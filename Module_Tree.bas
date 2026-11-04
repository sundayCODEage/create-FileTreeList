Attribute VB_Name = "Module_Tree"
Option Explicit
' ==============================================================================
' モジュール概要： 指定フォルダ配下のファイル構造をツリー形式でシートへ一括出力する。

' [VBAからの呼び出し窓口]
' └─ CreateFileTreeList               (メインプロシージャ：全体制御)
'      │
'      ├─ 1. [初期設定とルート取得]
'      │   ├─ MsgBox                  (一時ファイル除外設定の確認)
'      │   ├─ FileDialog              (対象ルートフォルダの選択)
'      │   ├─ InitializeSheet         (シート初期化・ヘッダー生成・枠固定)
'      │   └─ ReDim mBuffer           (動的メモリバッファの初期確保：2,048行)
'      │
'      ├─ 2. [再帰的フォルダ探索とメモリバッファ蓄積]
'      │   └─ ProcessFolderRecursively (親フォルダ, 現在インデント)  <───┐ [再帰呼び出し]
'      │        │                                                            │
'      │        ├─ SortAndFilterItems (Files)                               │
'      │        │   ├─ 除外判定       (Thumbs.db, ~$*, *.tmp のスキップ)   │
'      │        │   └─ QuickSortObjects (ファイル名昇順のクイックソート)   │
'      │        │                                                            │
'      │        ├─ SortAndFilterItems (SubFolders)                          │
'      │        │   ├─ 権限エラー検知 (エラー70発生時は安全にEmpty返却)    │
'      │        │   └─ QuickSortObjects (フォルダ名昇順のクイックソート)   │
'      │        │                                                            │
'      │        ├─ [A] ファイル群の走査ループ                               │
'      │        │   ├─ 枝記号の決定   (最終要素なら "└─" / 他は "├─")  │
'      │        │   └─ BufferFileInfo                                      │
'      │        │       └─ AddRowToBuffer (RowType=0:通常 / 2:長パス)      │
'      │        │                                                            │
'      │        └─ [B] サブフォルダ群の走査ループ                           │
'      │            ├─ 枝記号の決定   (最終要素なら "└─" / 他は "├─")   │
'      │            ├─ AddRowToBuffer (RowType=1:フォルダ行 [フォルダ名])   │
'      │            ├─ 次階層インデントの生成                               │
'      │            │   ├─ 最終フォルダの場合 ： indent + "       " (空白) │
'      │            │   └─ 後続がある場合     ： indent + "  │   " (縦線) │
'      │            └─ ProcessFolderRecursively (子フォルダ, 次インデント)─┘
'      │
'      ├─ 3. [シートへの一括出力と書式適用]
'      │   └─ FlushBufferToSheet
'      │        ├─ 2次元配列の構築    (A?D列の値配列 / E列のHYPERLINK数式配列)
'      │        ├─ 配列の一括書き込み (Resize.Value / Resize.Formula)
'      │        ├─ フォント一括上書き (数式適用後に BIZ UDゴシック 9pt を再適用)
'      │        └─ Unionバッチ着色    (50セル単位でフォルダ青字・長パス灰字を適用)
'      │
'      └─ 4. [レイアウト調整と終了処理]
'          ├─ Columns("A:A").AutoFit(列幅自動調整と最大幅70の制限適用)
'          ├─ Range.AutoFilter        (A3:E最終行へオートフィルタ設定)
'          └─ RestoreSettings         (バッファ解放・Excel環境設定の復元)

' ==============================================================================

' [モジュール定数・変数の宣言]
    ' ツリー描画用定数： 階層構造を表現するための罫線および空白文字列を定義する。
    Private Const TREE_BRANCH As String = "  ├─ "
    Private Const TREE_LAST   As String = "  └─ "
    Private Const TREE_VLINE  As String = "  │   "
    Private Const TREE_SPACE  As String = "       "
    
    ' フィルタ設定フラグ： 一時ファイルを除外するかどうかの選択結果を保持する。
    Private isFilterEnabled As Boolean
    
    ' 出力バッファ構造体： シートへの個別アクセスを回避するため1行分の出力情報を保持する。
    Private Type TreeRowBuffer
        ColA_Tree As String
        ColB_BaseName As String
        ColC_Modified As String
        ColD_Size As String
        ColE_Path As String
        rowType As Byte ' 行種別： 0は通常ファイル、1はフォルダ、2は250文字超の長パスを表す。
    End Type
    
    ' 動的配列バッファ： 探索中の全行データをメモリ上に蓄積する。
    Private mBuffer() As TreeRowBuffer
    Private mBufferCount As Long
    Private mBufferCapacity As Long
    
    ' 時間計測用変数： 処理開始時刻と総実行秒数を保持する。
    Private startTime As Double
    Private totalSeconds As Double
    
' ==============================================================================
' --- メイン処理 ---
' フォルダ選択からツリー生成およびシート出力までの全体を制御する。
Public Sub CreateFileTreeList()
    Dim fso As Object
    Dim targetSheet As Worksheet
    Dim rootPath As String
    Dim rootFolder As Object
    Dim writeRow As Long
    Dim res As VbMsgBoxResult
    
    ' [実行条件の確認と対象フォルダの取得]
    ' フィルタ設定の確認： 一時ファイルを除外するかどうかをダイアログで取得する。
    Dim strMsg As String
    strMsg = "除外設定を適用しますか？" & vbCrLf & vbCrLf & _
             " ・ 一時ファイル (~$*, *.tmp)" & vbCrLf & _
             " (・ システム管理ファイル (Thumbs.db) は常時除外)" & vbCrLf & vbCrLf & _
             "  ※メモリ一括出力方式により高速化されています。" & vbCrLf & _
             "  キャンセルは、「フォルダ選択」で選択" & vbCrLf & vbCrLf & _
             "  *** Esc *** 中断可（終了(E)を選択する）"
             
    res = MsgBox(strMsg, vbYesNo + vbQuestion, "フィルタ設定")
    isFilterEnabled = (res = vbYes)
    
    ' 親フォルダの選択： フォルダ選択ダイアログを表示し、キャンセル時は処理を終了する。
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ツリーを作成する親フォルダを選択してください"
        If .Show = -1 Then
            rootPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    ' 計測開始： 処理時間の算出用に現在時刻を記録する。
    startTime = Timer

    ' [Excel環境設定の最適化とシート初期化]
    ' 描画と計算の抑制： 処理速度向上のため画面更新、自動計算、イベントを一時停止する。
    On Error GoTo ErrorHandler
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayStatusBar = True
    End With

    ' シートとFSOの初期化： 出力先シートを初期化し、ルートフォルダオブジェクトを取得する。
    Set targetSheet = ThisWorkbook.Sheets(1)
    InitializeSheet targetSheet

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(rootPath)
    
    ' メモリバッファの初期化： 初期容量として2,048行分を確保する。
    mBufferCount = 0
    mBufferCapacity = 2048
    ReDim mBuffer(1 To mBufferCapacity)
    
    ' ルートフォルダ名の出力： 起点となる第4行にルートフォルダ名を太字で出力する。
    writeRow = 4
    With targetSheet.Cells(writeRow, 1)
        .Value = rootFolder.Name
        .Font.Bold = True
        .Font.Size = 10
    End With
    writeRow = writeRow + 1
    
    ' [再帰探索とシートへの一括出力]
    ' 再帰探索の実行： ルートフォルダを起点に全階層の情報をメモリバッファへ格納する。
    ProcessFolderRecursively fso, rootFolder, ""
    
    ' バッファの一括出力： 蓄積された全行データをシートへ一括で書き出す。
    If mBufferCount > 0 Then
        Application.StatusBar = "シートへ一括出力中... (" & Format(mBufferCount, "#,##0") & " 行)"
        DoEvents
        FlushBufferToSheet targetSheet, writeRow
        writeRow = writeRow + mBufferCount
    End If

    ' [レイアウト調整と終了処理]
    ' 列幅の自動調整： A列の幅を内容に合わせ、上限値70を超えた場合は制限する。
    targetSheet.Columns("A:A").AutoFit
    If targetSheet.Columns("A").ColumnWidth > 70 Then targetSheet.Columns("A").ColumnWidth = 70
    targetSheet.Range("C4").Select
    
    ' オートフィルタの適用： 既存のフィルタを解除し、データ範囲全体に再設定する。
    targetSheet.AutoFilterMode = False
    If writeRow > 4 Then
        targetSheet.Range("A3:E" & writeRow - 1).AutoFilter
    End If
    
    ' 環境設定の復元： 一時停止していたExcelの各種設定を規定値に戻す。
    GoSub RestoreSettings
    
    ' 経過時間の算出： 日付変更線をまたいだ場合のマイナス値を補正し、完了メッセージを表示する。
    totalSeconds = Timer - startTime
    If totalSeconds < 0 Then totalSeconds = totalSeconds + 86400
    
    MsgBox "* 処理完了" & vbCrLf & vbCrLf & _
           "    出力行数: " & Format(writeRow - 4, "#,##0") & " 行" & vbCrLf & _
           "    実行時間: " & Format(totalSeconds, "0.00") & "秒", vbInformation

    Exit Sub

    ' ==========================================================================
RestoreSettings:
    ' リソース解放と設定復元： 配列メモリとオブジェクトを解放し、Application設定を戻す。
    Erase mBuffer
    Application.StatusBar = False
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    Set rootFolder = Nothing
    Set fso = Nothing
    Return

    ' ==========================================================================
ErrorHandler:
    ' 例外処理： 環境設定を復元した上でエラー内容を通知する。
    GoSub RestoreSettings
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' --- バッファ追加処理 ---
' 1行分のデータをメモリ配列へ格納し、容量不足時は2倍に拡張する。
Private Sub AddRowToBuffer(ByVal colA As String, _
                           ByVal colB As String, _
                           ByVal colC As String, _
                           ByVal colD As String, _
                           ByVal colE As String, _
                           ByVal rowType As Byte)
    ' 配列サイズの動的拡張： 現在の要素数が上限を超えた場合、既存データを保持して2倍に拡張する。
    mBufferCount = mBufferCount + 1
    If mBufferCount > mBufferCapacity Then
        mBufferCapacity = mBufferCapacity * 2
        ReDim Preserve mBuffer(1 To mBufferCapacity)
    End If
    
    ' 行データの格納： 構造体の各メンバへ引数の値を代入する。
    With mBuffer(mBufferCount)
        .ColA_Tree = colA
        .ColB_BaseName = colB
        .ColC_Modified = colC
        .ColD_Size = colD
        .ColE_Path = colE
        .rowType = rowType
    End With
End Sub

' ==============================================================================
' --- シート一括出力処理 ---
' バッファを2次元配列へ変換し、シートへの一括書き込みと書式適用を行う。
Private Sub FlushBufferToSheet(ByVal ws As Worksheet, ByVal startRow As Long)
    Dim arrValues() As Variant
    Dim arrFormulasE() As Variant
    Dim i As Long
    Dim targetRow As Long
    Dim lastRow As Long
    
    ' 出力用2次元配列の確保： シートへ一括転送するためのVariant配列を初期化する。
    lastRow = startRow + mBufferCount - 1
    ReDim arrValues(1 To mBufferCount, 1 To 4)
    ReDim arrFormulasE(1 To mBufferCount, 1 To 1)
    
    ' 出力データの詰め替え： 構造体バッファからA～D列の値配列およびE列の数式配列を構築する。
    For i = 1 To mBufferCount
        With mBuffer(i)
            arrValues(i, 1) = .ColA_Tree
            arrValues(i, 2) = .ColB_BaseName
            arrValues(i, 3) = .ColC_Modified
            arrValues(i, 4) = .ColD_Size
            
            ' E列の出力形式分岐： フォルダは空欄、長パスは文字列、通常ファイルはHYPERLINK数式を設定する。
            If .ColE_Path = "" Then
                arrFormulasE(i, 1) = ""
            ElseIf .rowType = 2 Then
                arrFormulasE(i, 1) = .ColE_Path
            Else
                Dim escapedPath As String
                escapedPath = Replace(.ColE_Path, """", """""")
                arrFormulasE(i, 1) = "=HYPERLINK(""" & escapedPath & """,""" & escapedPath & """)"
            End If
        End With
    Next i
    
    ' 列書式の一括設定： C列の文字列書式とC～D列の右揃えを列単位で一括適用する。
    ws.Range("C" & startRow & ":C" & lastRow).NumberFormatLocal = "@"
    ws.Range("C" & startRow & ":D" & lastRow).HorizontalAlignment = xlRight
    
    ' 2次元配列の一括書き込み： セルへの個別アクセスを避け、対象領域へ一度に値を転写する。
    ws.Range("A" & startRow).Resize(mBufferCount, 4).Value = arrValues
    ws.Range("E" & startRow).Resize(mBufferCount, 1).Formula = arrFormulasE
    
    ' ハイパーリンク列のフォント統一： 数式入力で自動適用される標準スタイルをBIZ UDゴシックで上書きする。
    With ws.Range("E" & startRow & ":E" & lastRow).Font
        .Name = "BIZ UDゴシック"
        .Size = 9
        .Color = RGB(5, 99, 193)
        .Underline = xlUnderlineStyleSingle
    End With
    
    ' 個別行のバッチ着色： Unionを用いてフォルダ行（青字）と長パス行（灰字）を50セル単位で一括着色する。
    Dim folderRng As Range
    Dim grayRng As Range
    Dim folderBatchCount As Long
    Dim grayBatchCount As Long
    
    For i = 1 To mBufferCount
        targetRow = startRow + i - 1
        If mBuffer(i).rowType = 1 Then
            If folderRng Is Nothing Then
                Set folderRng = ws.Cells(targetRow, 1)
            Else
                Set folderRng = Union(folderRng, ws.Cells(targetRow, 1))
            End If
            folderBatchCount = folderBatchCount + 1
            ' Unionの肥大化による速度低下を防ぐため、50セルごとに書式を確定する。
            If folderBatchCount >= 50 Then
                folderRng.Font.Color = RGB(0, 102, 204)
                Set folderRng = Nothing
                folderBatchCount = 0
            End If
        ElseIf mBuffer(i).rowType = 2 Then
            If grayRng Is Nothing Then
                Set grayRng = ws.Cells(targetRow, 5)
            Else
                Set grayRng = Union(grayRng, ws.Cells(targetRow, 5))
            End If
            grayBatchCount = grayBatchCount + 1
            If grayBatchCount >= 50 Then
                grayRng.Font.Color = RGB(150, 150, 150)
                grayRng.Font.Underline = xlUnderlineStyleNone
                Set grayRng = Nothing
                grayBatchCount = 0
            End If
        End If
    Next i
    
    ' 残余セルの着色確定： 未適用のUnion範囲が残っている場合に書式を反映する。
    If Not folderRng Is Nothing Then folderRng.Font.Color = RGB(0, 102, 204)
    If Not grayRng Is Nothing Then
        grayRng.Font.Color = RGB(150, 150, 150)
        grayRng.Font.Underline = xlUnderlineStyleNone
    End If
End Sub

' ==============================================================================
' --- 再帰探索処理 ---
' フォルダ階層を深さ優先で走査し、ツリー構造データをメモリバッファへ蓄積する。
Private Sub ProcessFolderRecursively(ByVal fso As Object, _
                                     ByVal parentFolder As Object, _
                                     ByVal indent As String)
    
    Dim sortedFiles As Variant
    Dim sortedFolders As Variant
    Dim i As Long
    Dim currentIdx As Long
    Dim totalItems As Long
    
    ' 進捗表示の更新： 100件ごとにステータスバーを更新し、OSへ制御を戻してEsc中断を受け付ける。
    If mBufferCount Mod 100 = 0 Then
        Application.StatusBar = "探索中... " & Format(mBufferCount, "#,##0") & " 件目を取得中... (Escで中断可)"
        DoEvents
    End If

    ' 要素の取得と整列： 現在のフォルダ内のファイルとサブフォルダをフィルタリングおよび名前順ソートする。
    sortedFiles = SortAndFilterItems(parentFolder, False)
    sortedFolders = SortAndFilterItems(parentFolder, True)
    
    ' 要素総数の算出： 枝記号の終端（「└─」）を判定するため、現在階層の有効要素数を合算する。
    totalItems = 0
    If IsArray(sortedFiles) Then totalItems = totalItems + (UBound(sortedFiles) + 1)
    If IsArray(sortedFolders) Then totalItems = totalItems + (UBound(sortedFolders) + 1)
    
    currentIdx = 0
    
    ' ファイル要素の走査： ソート済みファイル配列を順次バッファへ登録する。
    If IsArray(sortedFiles) Then
        For i = LBound(sortedFiles) To UBound(sortedFiles)
            currentIdx = currentIdx + 1
            
            ' 枝記号の決定： 現在階層の最終要素であれば「└─」、後続があれば「├─」を選択する。
            Dim isLastItem As Boolean
            isLastItem = (currentIdx = totalItems)
            
            Dim prefix As String
            prefix = IIf(isLastItem, TREE_LAST, TREE_BRANCH)
            
            ' ファイル情報の登録： インデントと枝記号を連結してバッファへ追加する。
            BufferFileInfo fso, indent & prefix, sortedFiles(i)
            
            Set sortedFiles(i) = Nothing
        Next i
        Erase sortedFiles
    End If

    ' サブフォルダ要素の走査と再帰： フォルダ行を登録した後、インデントを更新して自身を再帰呼び出しする。
    If IsArray(sortedFolders) Then
        For i = LBound(sortedFolders) To UBound(sortedFolders)
            currentIdx = currentIdx + 1
            
            ' 枝記号の決定： 現在階層の最終フォルダかどうかを判定する。
            Dim isLastFolder As Boolean
            isLastFolder = (currentIdx = totalItems)
            
            Dim folderPrefix As String
            folderPrefix = IIf(isLastFolder, TREE_LAST, TREE_BRANCH)
            
            ' フォルダ行の登録： フォルダ名を角括弧で囲み、青字対象の行種別（RowType=1）として追加する。
            AddRowToBuffer indent & folderPrefix & "[" & sortedFolders(i).Name & "]", _
                           "", "", "", "", 1
            
            ' 次階層インデントの生成： 最終フォルダなら空白、後続があるなら縦線（「│」）を連結する。
            Dim nextIndent As String
            nextIndent = indent & IIf(isLastFolder, TREE_SPACE, TREE_VLINE)
            
            ' サブフォルダの再帰探索： 生成したインデントを引き継いで下位階層を走査する。
            ProcessFolderRecursively fso, sortedFolders(i), nextIndent
            
            Set sortedFolders(i) = Nothing
        Next i
        Erase sortedFolders
    End If
End Sub

' ==============================================================================
' --- ファイル情報抽出 ---
' ファイルオブジェクトから各属性を取得し、メモリバッファへ登録する。
Private Sub BufferFileInfo(ByVal fso As Object, _
                           ByVal treePrefix As String, _
                           ByVal fileObj As Object)
    Dim fName As String
    Dim fPath As String
    Dim baseName As String
    Dim modDate As String
    Dim sizeKB As String
    Dim rType As Byte
    
    ' 属性値の取得： 特殊なシステムファイル等での参照エラーを回避しつつ各プロパティを取得する。
    On Error Resume Next
    fName = fileObj.Name
    fPath = fileObj.Path
    baseName = fso.GetBaseName(fPath)
    modDate = Format(fileObj.DateLastModified, "yyyy/mm/dd hh:mm:ss")
    sizeKB = Format(fileObj.Size / 1024, "#,##0.0") & " KB"
    On Error GoTo 0
    
    ' パス長の判定： HYPERLINK関数の引数上限（255文字）を超える場合は長パス行（RowType=2）とする。
    If Len(fPath) > 250 Then
        rType = 2
    Else
        rType = 0
    End If
    
    AddRowToBuffer treePrefix & fName, baseName, modDate, sizeKB, fPath, rType
End Sub

' ==============================================================================
' --- フィルタリングと整列 ---
' コレクションから除外対象を省き、名前順にソートした配列を返す。
Private Function SortAndFilterItems(ByVal parentFolder As Object, ByVal isFolder As Boolean) As Variant
    Dim itemsCollection As Object
    Dim itemCount As Long
    
    ' コレクションの取得： アクセス権限のないフォルダ（エラー70等）ではEmptyを返して安全にスキップする。
    On Error Resume Next
    If isFolder Then
        Set itemsCollection = parentFolder.subFolders
    Else
        Set itemsCollection = parentFolder.Files
    End If
    itemCount = itemsCollection.Count
    If Err.Number <> 0 Or itemCount = 0 Then
        Err.Clear
        On Error GoTo 0
        SortAndFilterItems = Empty
        Exit Function
    End If
    On Error GoTo 0
    
    ' 作業用配列の確保： オブジェクト本体と比較用の名前文字列を格納する配列を用意する。
    Dim arrItems() As Object
    Dim arrNames() As String
    ReDim arrItems(0 To itemCount - 1)
    ReDim arrNames(0 To itemCount - 1)
    
    Dim item As Object
    Dim idx As Long
    Dim sName As String
    idx = 0
    
    ' 除外判定と配列格納： システム管理ファイルおよび一時ファイルを除外しながら有効要素を格納する。
    For Each item In itemsCollection
        sName = item.Name
        Dim skip As Boolean
        skip = False
        
        If Not isFolder Then
            If StrComp(sName, "Thumbs.db", vbTextCompare) = 0 Then skip = True
        End If
        
        If isFilterEnabled And Not skip Then
            If sName Like "~$*" Or sName Like "*.tmp" Then skip = True
        End If
        
        If Not skip Then
            Set arrItems(idx) = item
            arrNames(idx) = sName
            idx = idx + 1
        End If
    Next item
    
    ' 配列サイズの確定： 有効要素が0件ならEmptyを返し、1件以上なら実データ数に合わせてリサイズする。
    If idx = 0 Then
        SortAndFilterItems = Empty
        Exit Function
    Else
        ReDim Preserve arrItems(0 To idx - 1)
        ReDim Preserve arrNames(0 To idx - 1)
    End If
    
    ' 名前順ソートの実行： 要素が2件以上存在する場合にクイックソートで昇順に並べ替える。
    If UBound(arrItems) > 0 Then
        QuickSortObjects arrItems, arrNames, LBound(arrItems), UBound(arrItems)
    End If
    
    SortAndFilterItems = arrItems
End Function

' ==============================================================================
' --- クイックソート処理 ---
' 名前配列をキーとしてオブジェクト配列を再帰的に昇順ソートする。
Private Sub QuickSortObjects(ByRef arrObj() As Object, ByRef arrKey() As String, _
                             ByVal low As Long, ByVal high As Long)
    Dim pivot As String
    Dim i As Long, j As Long
    Dim tempObj As Object
    Dim tempKey As String

    ' 基準値の選出： 範囲の中央要素をピボットとして設定する。
    i = low
    j = high
    pivot = arrKey((low + high) \ 2)

    ' 要素の分割と交換： 大文字小文字を区別せず文字列比較を行い、ピボットを境に要素を入れ替える。
    Do While i <= j
        Do While StrComp(arrKey(i), pivot, vbTextCompare) = -1: i = i + 1: Loop
        Do While StrComp(arrKey(j), pivot, vbTextCompare) = 1: j = j - 1: Loop
        
        If i <= j Then
            tempKey = arrKey(i): arrKey(i) = arrKey(j): arrKey(j) = tempKey
            Set tempObj = arrObj(i): Set arrObj(i) = arrObj(j): Set arrObj(j) = tempObj
'            Set tempObj = arrObj(i): Set tempObj(i) = arrObj(j): Set arrObj(j) = tempObj
            
            i = i + 1
            j = j - 1
        End If
    Loop

    ' 部分配列の再帰ソート： 分割された左右の領域に対してソートを継続する。
    If low < j Then QuickSortObjects arrObj, arrKey, low, j
    If i < high Then QuickSortObjects arrObj, arrKey, i, high
End Sub

' ==============================================================================
' --- シート初期化処理 ---
' 既存データを消去し、基本フォント・ヘッダー・ウィンドウ枠固定を設定する。
Private Sub InitializeSheet(ByRef ws As Worksheet)
    ' データ領域のクリア： 第3行以降の既存データおよび書式をすべて消去する。
    ws.Rows("3:" & ws.Rows.Count).Clear
    
    ' 全体フォントの設定： シート全体の基準フォントをBIZ UDゴシック（9pt）に統一する。
    With ws.Cells.Font
        .Name = "BIZ UDゴシック"
        .Size = 9
    End With
    
    ' テーブルヘッダーの作成： 第3行に項目名を配置し、背景色と下罫線を設定する。
    Dim headers As Variant
    headers = Array("構成図", "ファイル名", "更新日時", "サイズ", "パス（ハイパーリンク）")
    
    With ws.Range("A3").Resize(1, UBound(headers) + 1)
        .Value = headers
        .Interior.Color = RGB(240, 240, 240)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    ' タイトルと作成日時の出力： A1～A2セルに帳票名と実行日時を記録する。
    With ws.Range("A1")
        .Value = "ファイル階層一覧"
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("A2")
        .Value = "作成日時: " & Format(Now, "yyyy/mm/dd HH:mm:ss")
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("A3")
        .Font.Color = RGB(255, 0, 0)
    End With
    
    ' ウィンドウ枠の固定： ヘッダー行（第3行まで）が常時表示されるよう第4行で枠を固定する。
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Rows(4).Select
    ActiveWindow.FreezePanes = True
End Sub

