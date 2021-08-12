Attribute VB_Name = "mdl_DiffLiner"
Option Explicit

Const ROW_START             As Integer = 2

Const COL_A                       As Integer = 1
Const COL_A_CODE            As Integer = 2
Const COL_A_PROC            As Integer = 3
Const COL_A_PROCNAME  As Integer = 4
Const COL_B                       As Integer = 5
Const COL_B_CODE            As Integer = 6
Const COL_B_proc            As Integer = 7
Const COL_B_PROCNAME  As Integer = 8
Const COL_LS                      As Integer = 9

Sub AutoOpen()
'ツール起動時処理

Dim ans As String
Dim rng_R As Range

Application.ScreenUpdating = False

'__init__
    ans = MsgBox("初期化して起動しますか。", vbOKCancel)
    If ans = vbOK Then
        GoTo init
    Else
        Exit Sub
    End If
init:
    Call InitProc
        ThisWorkbook.Save
End Sub

Function InitProc()
'初期化処理

Dim cnt As Integer
Dim i As Integer
Dim sheetName As String

Application.ScreenUpdating = False

    cnt = ThisWorkbook.Sheets.Count

    For i = 1 To cnt
        sheetName = ThisWorkbook.Sheets(i).Name
        Sheets(sheetName).Select
        Call sheetSetting(sheetName)
    Next

    ThisWorkbook.Sheets("menu").Select
Application.ScreenUpdating = True
End Function

Function sheetSetting(sheetName As String)
'初期化処理サブルーチン

    Select Case sheetName
    Case "menu"
    Case "Detail"
        Sheets(sheetName).Cells.ClearContents
        Call ClearColor(Sheets(sheetName).Cells)
        Call CellColor(Rows(1), 191, 191, 191, 0)
        Cells(1, 1).Select
        
    Case "Summary"
        Sheets(sheetName).Cells.ClearContents
        Call ClearColor(Sheets(sheetName).Cells)
        Call CellColor(Rows(1), 191, 191, 191, 0)
        Cells(1, 1).Select
    End Select
End Function

Function IsText(path As String) As Boolean
'textファイルか判別

Dim ret As String
    ret = GetExtension(path)
    If ret = "text" Or ret = "log" Or "bas" Then
        IsText = True
    Else
        IsText = False
    End If
End Function

Sub main()
'テキストファイルを取込み、比較する

Dim pathA As String, pathB As String
Dim ret As String

    MsgBox "分析対象のテキストファイルを選択してください。"
    pathA = GetFilePath
    If pathA = "False" Then
        MsgBox "処理を中止します。"
        Exit Sub
    Else
    End If
    
    ret = MsgBox("比較対象のテキストファイルを選択してください。", vbOKCancel)
    If ret <> vbCancel Then
        pathB = GetFilePath
    Else
        pathB = ""
    End If
    
    Call ImportSource(pathA, pathB)

End Sub

Private Sub ImportSource(pathA As String, pathB As String)
'テキストファイルからデータを取込む

Dim SummaryFlg As Boolean
Dim col As Integer
Dim i As Long, j As Long
Dim detail As String, summary As String
Dim rngA As Range, rngB As Range
Dim rowEnd As Long

'__init__
    Sheets("Detail").Cells.Clear
    Sheets("Summary").Cells.Clear
    SummaryFlg = Sheets("menu").Range("S4")
    
'__main__
    Sheets("Detail").Select
    Application.ScreenUpdating = False
    
    'テキストファイルを読み込む
    Call GetProc(pathA, col:=COL_A)
    
    If pathB <> "" Then
        Call GetProc(pathB, col:=COL_B)
    Else
    End If
    
    'Summaryの一致度を返す
    If SummaryFlg = True Then
        Debug.Print "AnalizeSummary"
    Else
    End If
    
    'DetailのDiff分析/行追加
    rowEnd = GetRowEnd
    
    Do While rowProcNext(i, COL_A_PROC) <= rowEnd + 1
        j = rowProcNext(i, COL_A_PROC)
        
        If j > rowProcNext(i, COL_B_proc) Then
            i = rowProcNext(i, COL_B_proc)
        Else
            i = j
        End If
        
        Debug.Print i
        
        Call Analyze_Detail(i)
        
        'セルの最終行を探索
        rowEnd = GetRowEnd
    Loop
    
    'セルの最終行を探索
        rowEnd = GetRowEnd
    
    'セルの背景色設定
        Call CellColor(Range(Cells(1, 1), Cells(1, colLast("Detail", 1))), 100, 100, 100, 0.5)
        Call CellColor(Range(Cells(1, 1), Cells(rowEnd, 1)), 100, 100, 100, 0.5)
        Call CellColor(Range(Cells(1, 5), Cells(rowEnd, 5)), 100, 100, 100, 0.5)
        Call CellColor(Range(Cells(2, 9), Cells(rowEnd, 9)), 100, 100, 100, 0.8)
        
        Application.ScreenUpdating = True
        
        ThisWorkbook.Save
        MsgBox "処理が完了しました。"
    
End Sub

Function GetRowEnd()
'セルの最終行を探索

Dim rowEnd As Long
        rowEnd = rowLast("Detail", COL_A)
        If rowEnd < rowLast("Detail", COL_B) Then rowEnd = rowLast("Detail", COL_B)
        
        GetRowEnd = rowEnd
End Function

Sub AnalyzeDetail_onSave()
'ファイル上書き保存時に再算出

Dim rowEnd  As Long
Dim row As Long
Dim col As Integer
Dim CurrentRng As Range
Dim i As Long, j As Long

    Application.ScreenUpdating = False
    Set CurrentRng = ActiveCell
    
    'セルの最終行を探索
    rowEnd = GetRowEnd
    
    'セルの背景色をクリアする
    Call ClearColor(Cells)
    
    'dummy行を削除する
    For row = ROW_START To rowEnd
        For col = COL_A To COL_B Step COL_B - COL_A
            If Sheets("Detail").Cells(row, col) = "dummy" Then
                Call DeleteRow(row:=row, col:=col, rowRng:=1, colRng:=3)
                row = row - 1
            Else
            End If
        Next
    Next
    
    '一致度を削除
    Sheets("Detail").Columns(COL_LS).ClearContents
    
    '一致度を再算出
    Do While rowProcNext(i, COL_A_PROC) <= rowEnd + 1
        j = rowProcNext(i, COL_A_PROC)
        If j > rowProcNext(i, COL_B_proc) Then
            i = rowProcNext(i, COL_B_proc)
        Else
            i = j
        End If
        
        Debug.Print i
        
        Call Analyze_Detail(i)
    Loop
    
    'セルの最終行を探索
    rowEnd = GetRowEnd
    
    'セルの背景色設定
    Sheets("Detail").Select
    Call CellColor(Range(Cells(1, 1), Cells(1, colLast("Detail", 1))), 100, 100, 100, 0.5)
    Call CellColor(Range(Cells(1, 1), Cells(rowEnd, 1)), 100, 100, 100, 0.5)
    Call CellColor(Range(Cells(1, 5), Cells(rowEnd, 5)), 100, 100, 100, 0.5)
    Call CellColor(Range(Cells(1, 9), Cells(rowEnd, 9)), 100, 100, 100, 0.8)
    
    'ヘッダ再設定
    For col = COL_A To COL_B Step COL_B - COL_A
        Sheets("Detail").Cells(1, col) = "行"
        Sheets("Detail").Cells(1, col + 1) = "ソース"
        Sheets("Detail").Cells(1, col + 2) = "proc種類"
        Sheets("Detail").Cells(1, col + 3) = "proc名"
        Sheets("Detail").Cells(1, col + 4) = "一致度"
    Next
    
    Application.ScreenUpdating = True
    
On Error Resume Next
    CurrentRng.Select
    
End Sub

Function Analyze_Detail(row As Long) As Double
'Detailシートのソースの各行を探索しながら一致度の高い行を並べて表示する

Dim row_base As Long, col_base As Long, row_try As Long, col_try As Long
Dim base As Range, try As Range
Dim score As Double
Dim threshold As Double
Dim lst As Variant
Dim matchingFlg As Boolean
Dim i As Long, j As Long
Dim keyA As String, keyB As String
Dim END_ROW As Long

'__init__
    row_base = row
    col_base = COL_A_CODE
    row_try = row
    col_try = COL_B_CODE
    
    '閾値
    threshold = Sheets("menu").Range("S2")
    
    'キーワードリスト
    lst = Range(Sheets("menu").Cells(2, 16), Sheets("menu").Cells(rowLast("menu", 16), 16))
    
'__main__
    Analyze_Detail = 0
    
    For i = row To rowProcEnd(row, COL_A_PROC)
        Set base = Sheets("Detail").Cells(i, COL_A_CODE)
        
        If Sheets("Detail").Cells(i, COL_A) = "Dummy" Or _
            Sheets("Detail").Cells(i, COL_A) = "END" Then GoTo Continue
            
        'COL_Aにダミー行でない空白行があったらLsDist=0で処理する
        If Sheets("Detail").Cells(i, COL_A_CODE) = "" Then
            Sheets("Detail").Cells(i, COL_LS) = 0
            Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME)), 255, 255, 0, 0.5)    'セルをハイライト
            GoTo Continue
        Else
        End If
        
        For j = i To rowProcEnd(row, COL_B_proc)
            Set try = Sheets("Detail").Cells(i, COL_B_CODE)
            If Sheets("Detail").Cells(i, COL_B) = "dummy" Or _
                    Sheets("Detail").Cells(i, COL_B) = "END" Then GoTo Continue
                    
            'COL_Bにダミー行でない空白行があったらLsDist=0で処理する
            If Sheets("Detail").Cells(i, COL_B_CODE) = "" Then
                Sheets("Detail").Cells(i, COL_LS) = 0
                Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME)), 255, 255, 0, 0.5)    'セルをハイライト
                GoTo Continue
            Else
            End If
            
            '__init__
                score = 0
                matchingFlg = False
                score = LsDist(base.Value, try.Value)
                        
            '      '=== デバッグ用 ====================
            '      '対象セルをハイライト
            '      Call ClearColor(base.Offset(-1, 0))
            '      Call ClearColor(try.Offset(-1, 0))
            '      Call CellColor(base, 255, 255, 255, 0, 0)
            '      Call CellColor(try, 255, 255, 0, 0)
            '      Stop
            '      '=============================
                        
            '__main__
                '一致度が閾値を超えている場合、一致フラグをたてる
                If score >= threshold Then
                    Sheets("Detail").Cells(j, COL_LS) = score
                    matchingFlg = True
                    Exit For
                Else
                End If
        Next j
        
        'ダミー業を追加する
        If matchingFlg = False Then
        
        '行内容が不一致
        If j >= rowProcEnd(row, COL_B_proc) Then j = i  '挿入行数の最大値設定
            Range(Sheets("Detail").Cells(i, COL_B), Sheets("Detail").Cells(j, COL_B_PROCNAME)).Insert shift:=xlDown '挿入
            Sheets("Detail").Cells(i, COL_B) = "dummy"  'dummy表示
            Sheets("Detail").Cells(j, COL_LS) = 0 '一致度入力
            Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(j, COL_B_PROCNAME)), 255, 255, 0, 0.5)    'セルをハイライト
            row = i
            GoTo Restart
            
        '行内容の一致度が閾値を超えているか、途中行をスキップしている場合
        ElseIf matchingFlg = True And j - i > 0 Then
            'iとjの差分行数をbase列に追加
            If j >= rowProcEnd(row, COL_B_proc) Then j = i '挿入行数の最大値設定
                Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(j - 1, COL_A_PROCNAME)).Insert shift:=xlDown '挿入
                Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(j - 1, COL_A)) = "dummy" 'dummy表示
                Range(Sheets("Detail").Cells(i, COL_LS), Sheets("Detail").Cells(j - 1, COL_LS)) = 0 '一致度入力
                Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(j - 1, COL_B_PROCNAME)), 255, 255, 0, 0.5)  'セルをハイライト
                row = i
                GoTo Restart
            Else
                If score < 1 Then   '行内容の一致度が閾値を超えているが不一致
                    Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME)), 255, 255, 0, 0.5) 'セルをハイライト
                Else    '完全に一致
                    Call ClearColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME))) 'セルの色付けをなくす
                End If
            End If

Continue:
        'プロシージャ名をハイライト
        If IsProcLine(Cells(i, COL_A_PROC)) = True Or IsProcLine(Cells(i, COL_B_proc)) = True Then
            Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME)), 255, 150, 200, 0)
        Else
        End If
        
        keyA = IsWords(Cells(i, COL_A_CODE), lst)
        keyB = IsWords(Cells(i, COL_B_CODE), lst)
        
        'キーワードハイライト
        If keyA <> "" Then
            Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME)), 100, 220, 100, 0.2)
            Cells(i, COL_B_PROCNAME) = keyA
        Else
        End If
        If keyB <> "" Then
            Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME)), 100, 220, 100, 0.2)
            Cells(i, COL_B_PROCNAME) = keyB
        Else
        End If
             END_ROW = rowLast("Detail", COL_A_PROC)
             If END_ROW < rowLast("Detail", COL_B_proc) Then END_ROW = rowLast("Detail", COL_B_proc)
             If i + 1 >= END_ROW Then Exit For
    Next i
Exit Function

Restart:
    Call Analyze_Detail(row)
End Function

Function rowProcEnd(CurrentRow As Long, col As Long) As Long
'現在のprocの最終行を取得する

Dim row As Long
Dim rowEnd As Long

On Error GoTo errExit
    rowEnd = rowLast("Detail", 3)
    If rowEnd < rowLast("Detail", 7) Then rowEnd = rowLast("Detail", 7)
    
    row = CurrentRow
    Do Until Cells(row + 1, col) <> ""
        row = row + 1
        If row > rowEnd Then Exit Do
    Loop

'__return__
    rowProcEnd = row
Exit Function

errExit:
    rowProcEnd = rowEnd
End Function

Function rowProcNext(CurrentRow As Long, col As Long) As Long
'次のprocの開始行を取得する

Dim row As Long
Dim rowEnd As Long

On Error GoTo errExit
    rowEnd = rowLast("Detail", 3)
    If rowEnd < rowLast("Detail", 7) Then rowEnd = rowLast("Detail", 7)
    
    row = CurrentRow
    Do Until Cells(row + 1, col) <> ""
        row = row + 1
        If row > rowEnd Then Exit Do
    Loop
    
'__return__
    rowProcNext = row + 1
Exit Function

errExit:
    rowProcNext = rowEnd + 1
End Function

Function Check_min(InsertScore, DeleteScore, StableScore) As String
'最小値の項目を返す

    If InsertScore <= DeleteScore Then
        If InsertScore <> DeleteScore Then
            If InsertScore < StableScore Then
                Check_min = "Insert"
            ElseIf InsertScore > StableScore Then
                Check_min = "Stable"
            Else
                Check_min = "InsertStable"
            End If
        Else
            If InsertScore < StableScore Then
                Check_min = "InsertDelete"
            ElseIf InsertScore > StableScore Then
                Check_min = "Stable"
            Else
                Check_min = "InsertDeleteStable"
            End If
        End If
    ElseIf DeleteScore < StableScore Then
        Check_min = "Delete"
    ElseIf DeleteScore >= StableScore Then
        Check_min = "DeleteStable"
    End If
End Function

Function AnalyzeSummary() As Double
'Summary分析

Dim baseRng As Range, tryRng As Range

'__main__
    '基準range
    Set baseRng = Range(Sheets(summary).Cells(ROW_START, COL_PROC_A), _
                                    Sheets("Summary").Cells(rowLast("Summary", COL_PROC_A), COL_PROC_A))
    '比較range
    Set tryRng = Range(Sheets("Summary").Cells(ROW_START, COL_PROC_B), _
                                    Sheets("Summary").Cells(rowLast("Summary", COL_PROC_B), COL_PROC_B))

    'Summaryの行数差、一致率を算出する
    Call CheclProcOrder
    
    'T_LS/T_match　生成/並べ替え
    Call LsRng(baseRng, tryRng, "Summary")
    
'__Return__
    '一致度を再計算し返す
    AnalyzeSummary = LsRng(baseRng, tryRng, "Summary")
End Function

Function CheckProcOrder()
'Summaryの行数さ、一致率を算出する
'proc名の並び順を基準Rangeにあわせ並び替える

Const colA As Long = 1
Const colB As Long = 5
Const colGap As Long = 9
Const colLS As Long = 10
Dim rowMax As Long, row As Long, col As Long

'__main__
    '最終行取得
    rowMax = rowLast("Summary", colStart)
    
    For row = rowStart To rowMax
        For col = colA To colB - 1
            Select Case col
            'モジュール名相違
            Case 1
                'skip
            'プロシージャ名相違
            Case 2
                '一致率算出しセルに入力
                Sheets("Summary").Cells(row, colLS) = LsDist(Sheets("Summart").Cells(row, col), Sheets("Summary").Cells(row, col + (colB - colA)))
            End Select
        Next
    Next
End Function

Function LsRng(basRng As Range, tryRng As Range, sheetName As String) As Double
'文字列の比較
'Levenshtein距離で類似度を測定し、一致度を返す
'retuen Max:1/min:0

Const SheetNameA As String = "T_LS"
Const SheetNameB As String = "T_match"

Dim LS_matrix() As Variant, matching_matrix() As Variant
Dim i As Integer, j As Integer, cost As Integer
Dim missCnt As Integer

Dim Base_R1C1 As String, Try_R1C1 As String
Dim row_base As Long, col_base As Long, row_try As Long, col_try As Long

'__init__
    '各rangeの開始行、開始列を取得する
    Base_R1C1 = basrRng.Address(ReferenceStyle:=xlR1C1)
    Base_R1C1 = Left(Base_R1C1, InStr(Vase_R1C1, ":") - 1)
    row_base = CLng(Replace(Replace(Base_R1C1, (Right(Base_R1C1, InStr(Base_R1C1, "C") - 1)), ""), "R", "")) - 1
    col_base = CLng(Replace(Base_R1C1, (Left(Base_R1C1, InStr(Base_R1C1, "C"))), ""))
    
    Try_R1C1 = tryRng.Address(ReferenceStyle:=xlR1C1)
    Try_R1C1 = Left(Try_R1C1, InStr(Try_R1C1, ":") - 1)
    
    row_try = CLng(Replace(Replace(Try_R1C1, (Right(Try_R1C1, InStr(Try_R1C1, "C") - 1)), ""), "R", "")) - 1
    coltry = CLng(Replace(Try_R1C1, (Left(Try_R1C1, InStr(Try_R1C1, "C"))), ""))
    
    '配列の再定義
        ReDim LS_matrix(baseRng.Count, tryRng.Count)
        ReDim matching_matrix(baseRng.Count, tryRng.Count)
    
    '__main__
        LsRng = 0
        
        '対象Rangeがなかった場合「０」で返す
        If baseRng.Count = 0 Then
            LsRng = 0
            Exit Function
        ElseIf tryRng.Count = 0 Then
            LsRng = 0
            Exit Function
        Else
        End If
        
        '各配列に初期値を与える
        For i = 0 To baseRng.Count
            LS_matrix(i, 0) = i
            matching_matrix(i, 0) = i
        Next i
        
        For j = 0 To tryRng.Count
            LS_matrix(0, j) = j
            matching.matrix(0, j) = j
        Next j
        
        '配列の可視化
        Range(Sheets(SheetNameA).Cells(1, 1), Sheets(SheetNameA).Cells(baseRng.Count + 1, tryRng.Count + 1)) = LS_matrix
        Range(Sheets(SheetNameA).Cells(1, 1), Sheets(SheetNameA).Cells(baseRng.Count + 1, tryRng.Count + 1)) = mathing_matrix
        
        '各セルを組合せ一致度を探索する
        For i = 1 To baseRng.Count
            For j = 1 To tryRng.Count
                
                '値が一致する場合は「０」、一致しない場合は「１」
                cost = IIf(Sheets(sheetName).Cells(row_base + i, col_base) = Sheets(sheetName).Cells(row_try + j, col_try), 0, 1)
                
                '一致組合せをmatrixに格納
                '値が一致する場合は「１」、一致しない場合は「０」に変換して格納
                matching_matrix(i, j) = Abs(cost - 1)
                
                '最小コストをmatrixに格納
                LS_matrix(i, j) = WorksheetFunction.Min(LS_matrix(i - 1, j) + 1, LS_matrix(i, j - 1) + 1, LS_matrix(i - 1, j - 1) + cost)
                  
                'LS_matrix(i - 1, j) + 1               '要素の削除
                'LS_matrix(i, j -1) + 1                '要素の挿入
                'LS_matrix(i - 1, j - 1) + cost     '要素の置換
                
''                '配列の可視化
''                '１行   ：i  / baseRngの要素番号
''                'A列    ：j / tryRngの要素番号
''                'セル値：一致/不一致の累積地
''
''                '正解  最小値が右に進む
''                '挿入  最小値が下に進む
''                '削除  最小値が右に進む
''                '置換  最小値が右下に進む
''                Sheets(SheetNameA).Cells(i + 1, j + 1) = LS_matrix(i, j)
            Next j
        Next i
        
        missCnt = LS_matrix(baseRng.Count, tryRng.Count)
        
        '配列の可視化
        Range(Sheets(SheetNameA).Cells(1, 1), Sheets(SheetNameA).Cells(baseRng.Count + 1, tryRng.Count + 1)) = LS_matrix
        Range(Sheets(SheetNameB).Cells(1, 1), Sheets(SheetNameB).Cells(baseRng.Count + 1, tryRng.Count + 1)) = matching_matrix
        
        '並べ替え
        Call LsOrder_Summary(LS_matrix, matching_matrix, baseRng, tryRng)
        
'__return__
    '一致度を返す
    LsRng = Format(Abs(1 - (missCnt / baseRng.Count)), "0.00")
    If LsRng < 0 Then LsRng = 0
        
End Function

Function LsOrder_Summary(LS_matrix As Variant, matching_matrix As Variant, baseRng As Range, tryRng As Range) As Boolean

Dim row_tg As Long, col_tg  As Long
Dim flg_dup As Integer
Dim rowOrg As Long, rowDest As Long, rowRng As Long

Const SheetNameA As String = "T_LS"
Const SheetNameB As String = "T_match"

Const COL_START = 2
Const COL_CntAst = 3
Const COL_CntArng = 4
Const COL_B_proc = 6
Const COL_CntBst = 7
Const COL_Brng = 8
Const COL_DUP = 11

Const COL_Detail = 5

Const COL_PROC_A = 2
Const COL_PROC_B = 6

'__main__
    '一致/置換しているセルを探索する
    '各行を一つずつ確認する
    For row_tg = COL_START To rowLast(SheetNameA, 1)
    
        '重複の確認
        flg_dup = WorksheetFunction.Sum(Sheets(SheetNameB).Rows(row_tg)) - (row_tg - 1)
        
        '重複フラグをたてる
        If flg_dup > 1 Then
            Sheets("Summary").Cells(row_tg, COL_DUP) = 1
        Else
            Sheets("Summary").Cells(row_tg, COL_DUP) = ""
        End If
        
        '各列をひとつずつ確認する
        For col_tg = COL_START To rowLast(SheetNameA, 1)
        
            '一致/置換しているセルか確認する
            If Sheets(SheetNameB).Cells(row_tg, col_tg) = 1 Then
            
                '一致したセルの行と列が同じか確認する
                If row_tg = col_tg Then
                
                    '同じ行で一致している
                    Debug.Print
                    Debug.Print row_tg; Sheets("Summary").Cells(row_tg, COL_B_proc)
                    Debug.Print col_tg; Sheets("Summary").Cells(col_tg, COL_B_proc)
                    Exit For
                Else
                    '重複ある場合は同じ行で一致/置換フラグがないか確かめる
                    If flg_dup > 1 Then
                        If Sheets(SheetNameB).Cells(row_tg, row_tg) = 1 Then
                            Debug.Print
                            Debug.Print row_tg; Sheets("Summary").Cells(row_tg, COL_B_proc)
                            Debug.Print col_tg; Sheets("Summary").Cells(col_tg, COL_B_proc)
                            Debug.Print "continue"
                            GoTo Continue
                        Else
                        End If
                    Else
                    End If
    
                    '重複がない/basRngと同じ行に一致/置換フラグがない場合は行を移動する
                    Debug.Print
                    Debug.Print row_tg; Sheets("Summary").Cells(row_tg, COL_B_proc)
                    Debug.Print col_tg; Sheets("Summary").Cells(col_tg, COL_B_proc)
        
                    'Summary挿入
                    Call procLineInsert(rowOrg:=col_tg, row_Dest:=row_tg, rowRng:=1, _
                                                    col:=COL_B_proc, colRng:=3, sheetName:="Summary")
                    'Detail挿入
                    rowOrg = Sheets("Summary").Cells(row_tg, col_CmtBst)
                    rowRng = Sheets("Summary").Cells(row_tg, col_CntBrng)
                    rowDest = Sheets("Summary").Cells(ROW_START, COL_CntBst) + _
                                    WorksheetFunction.Sum(Range(Sheets("Summary").Cells(ROW_START, col_CntBrng), _
                                    Sheets("Summary").Cells(row_tg - 1, col_CntBrng))) + 1
                                    
                    Debug.Print "rowOrg :   " & rowOrg
                    Debug.Print "rowRng :   " & rowRng
                    Debug.Print "rowDest :  " & rowDest 'COL_CntAst列は行の入れ替えをしていないので、D_RowNoで行番号を検索する必要はない
                
                    Call procLineInsert(rowOrg:=D_RowNo(rowOrg), rowDest:=rowDest, rowRng:=rowRng, col:=COL_Detail)
                    
                    '再計算
                    Set baseRng = Range(Sheets("Summary").Cells(ROW_START, COL_PROC_A), Sheets("Summary").Cells(rowLast("Summary", COL_PROC_A), COL_PROC_A))
                    Set tryRng = Range(Sheets("Summary").Cells(ROW_START, COL_PROC_B), Sheets("Summary").Cells(rowLast("Summary", COL_PROC_B), COL_PROC_B))
                    
                    Call CheckProcOrder
                    Call LsRng(baseRng, tryRng, summary)
                End If
            Else
                Debug.Print Sheets(SheetNameB).Cells(row_tg, col_tg)
            End If
Continue:
        Next
    Next

'__return__
    LsOrder_Summary = True

End Function

Function D_RowNo(rowNo As Long, Optional col As Long = 5) As Long
'Detailシートのソース行番号に合致するExcelシート業を返す
'return    Excelの行番号

Dim i As Long
    For i = rowStart To rowLast("Detail", col)
        If Sheets("Detail").Cells(i, col) = rowNo Then
            D_RowNo = i
            Exit Function
        Else
        End If
    Next

'__DefaultReturn__
    D_RowNo = 0
End Function

Function procLineInsert(rowOrg As Long, rowDest As Long, col As Integer, _
                Optional rowRng As Long = 1, Optional colRng As Integer = 4, _
                Optional sheetName As String = "Detail")
'行範囲を指定の行の前に挿入する
'rowOrigin      :   行番号
'rowRng         :   行数
'colRng           :    列数
'sheetName    :    シート名

Dim LineOrg As Range
Dim RngOrg As Range, rngOrgINSBellow As Range, RngOrgINSAbove As Range, RngDest As Range

'__init__
    Sheets(sheetName).Select

'__check__
    If rowOrg = rowDest Then
        Debug.Print "rowOrg :" & rowOrg & " =rowDest" & rowDest & " =>  Exec False"
        Exit Function
    End If
    
'__main__
    '値を取得
    'コメントアウト行の" '　"を残すため、値を変数に格納するのではなく、コピペで対応する
    
    'rowDest行を挿入
    Range(Sheets(sheetName).Cells(rowDest, col), Sheets(sheetName).Cells(rowDest + rowRng - 1, col + colRng - 1)).Select
    Range(Sheets(sheetName).Cells(rowDest, col), Sheets(sheetName).Cells(rowDest + rowRng - 1 + colRng - 1)).Insert shift:=xlDown
    
    'rowOrg行をコピー
    If rowOrg < rowDest Then
        Range(Sheets(sheetName).Cells(rowOrg, col), Sheets(sheetName).Cells(rowOrg + rowRng - 1, col + colRng - 1)).Select
        Range(Sheets(sheetName).Cells(rowOrg, col), Sheets(sheetName).Cells(rowOrg + rowRng - 1, col + colRng - 1)).Copy
    Else
        Range(Sheets(sheetName).Cells(rowOrg + rowRng, col), Sheets(sheetName).Cells(rowDest + rowRng - 1, col + colRng - 1)).Select
        Range(Sheets(sheetName).Cells(rowOrg + rowRng, col), Sheets(sheetName).Cells(rowDest + rowRng - 1, col + colRng - 1)).PasteSpecial
        
        'rowOrg行を削除
        If rowOrg < rowDest Then
            Call DeleteRow(rowOrg, col, rowRng:=rowRng, sheetName:=sheetName)
        Else
            Call DeleteRow(rowOrg + rowRng, col, rowRng:=rowRng, sheetName:=sheetName)
        End If

End Function

Function procLineUpDown(rowA As Long, rowB As Long, col As Integer, _
                Optional rowRng As Long = 1, Optional rowRngB As Long = 1, _
                Optional colRng As Integer = 4, Optional sheetName As String = "Detail")
'行範囲Aと行範囲Bを入れ替える

Dim LineA As Variant, LineB As Variant
Dim i As Long
Dim tmp

'__init__
    Sheets(sheetName).Select

    'rowBはrowAの後の行でなければいけない
    If rowA > rowB Then
        tmp = Empty
        tmp = rowA
        rowA = rowB
        rowB = tmp
        tmp = rowRngA
        rowRngA = rowRngB
        rowRngB = tmp
    Else
    End If
    
    '行範囲Aは行範囲Bの一部を含んではいけない
    If rowA + rowRngA - 1 > rowB Then
        Debug.Print (" rowB must be bigger than rowA +rowRngA.")
        MsgBox (" rowB must be bigger than rowA +rowRngA.")
        End
    Else
    End If
    
'__main__
    '値を取得
    LineA = Range(Sheets(sheetName).Cells(rowA, col), Sheets(sheetName).Cells(rowA + rowRngA - 1, col + colRng - 1))
    LineB = Range(Sheets(sheetName).Cells(rowB, col), Sheets(sheetName).Cells(rowB + rowrngG - 1, col + colRng - 1))
    
    '行を入れ替ええる
    'rowAの値→rowBの値に挿入する
    With Range(Sheets(sheetName).Cells(rowB, col), Sheets(sheetName).Cells(rowB + rowRngA - 1, col + colRng - 1))
        .Select
        .Insert shift:=xlDown
        .NumberFormatLocal = "@"
        .Value = LineB
    End With
    
    Call DeleteRow(rowB + rowRngB, col, rowRng:=rowRngB, sheetName:=sheetName)
                
End Function

Function Compare(valA, valB) As Boolean
'値を比較する
'return     True:一致/False:不一致

    If valA = valB Then
        Compare = True
    Else
        Compare = False
    End If
End Function

Function InsertRow(row As Long, col As Integer, _
                                Optional rngCnt As Long = 3, Optional sheetName As String = "Detail")
'指定の範囲に行を挿入する
    
    Range(Sheets(sheetName).Cells(row, col), Sheets(sheetName).Cells(row, col + rngCnt)).Insert shift:=xlDown
End Function

Function DeleteRow(row As Long, col As Integer, _
                                Optional rowRng As Long = 1, Optional colRng As Integer = 3, _
                                Optional sheetName As String = "Detail")
'指定の範囲の行を削除する

    Range(Sheets(sheetName).Cells(row, col), Sheets(sheetName).Cells(row + rowRng - 1, col + colRng)).Select
    Range(Sheets(sheetName).Cells(row, col), Sheets(sheetName).Cells(row + rowRng - 1, col + colRng)).Delete
End Function

Function GetModuleName(ByVal arr As Variant) As String
'モジュール名を取得する
Dim i As Long
Dim strTg As String
Dim strModuleName As String
    strTg = "Attribute VB_Name ="
    
    For i = LBound(arr) To UBound(arr)
        If InStr(arr(i), strTg) > 0 Then
            strModuleName = arr(i)
            strModuleName = Replace(strModuleName, strTg, "")
            strModuleName = Replace(strModuleName, """", "")
            GetModuleName = strModuleName
            Exit Function
        Else
        End If
    Next
    
'__return__
    GetModuleName = "NULL"
End Function

Function IsProcName(ByVal strLine As String) As String
'プロシージャ・プロパティ定義行かの判定

    strLine = Replace(strLine, " ", "")
    strLine = Replace(strLine, "Private", "")
    strLine = Replace(strLine, "Public", "")
    
    Select Case True
        Case Left(strLine, 1) = "'"
            IsProcName = ""
        Case strLine Like "Sub*"
            IsProcName = "Sub"
        Case strLine Like "Function*"
            IsProcName = "Function"
        Case strLine Like "Property*"
            IsProcName = "Property"
        Case Else
            IsProcName = ""
        End Select
End Function

Sub GetProc(path As String, Optional col As Integer = 1)
'テキストファイルからモジュールを読み込む

Dim i As Long, j As Long, k As Long
Dim arrProc  As Variant
Dim arrProcInfo() As Variant, arrProcSummary() As Variant
Dim ModuleName As String, ProcName As String, ProcKind As String
Dim ProcLine As String

On Error Resume Next

'__init__
    j = 0
    Erase arrProcInfo
    Erase arrProcSummary
    
    arrProc = Read_txt(path, OutputFlg:=False)

'__main__
    ModuleName = GetModuleName(arrProc)
    
    For i = 0 To UBound(arrProc)
        
        'コード内容
        ProcLine = ProcLine & Trim(arrProc(i))
        
        Debug.Print i & " " & ProcLine
        
        If ProcLine <> "" Then  '空白行はスキップする
            j = j + 1
            ReDim Preserve arrProcInfo(3, j)
            
            If Right(ProcLine, 1) <> "_" Then   '続き行がない場合
                ProcKind = IsProcName(arrProc(i)) 'コードの種類を取得
            
                If ProcKind <> "" Then  'プロシージャの先頭行の場合
                    ProcName = Replace(ProcLine, "Public", "")
                    ProcName = Replace(ProcLine, "Private", "")
                    ProcName = Replace(ProcName, ProcKind, "")
                    ProcName = Replace(ProcName, " ", "")
                    ProcName = Left(ProcName, InStrRev(ProcName, "(") - 1)
                    
                    'Summary
                    k = k + 1
                    ReDim Preserve arrProcSummary(3, k)
                    arrProcSummary(0, k) = ModuleName                                'モジュール名
                    arrProcSummary(1, k) = ProcName                                     'プロシージャ名
                    arrProcSummary(2, k) = j                                                    '開始行
                    arrProcSummary(3, k - 1) = j - arrProcSummary(2, k - 1)   '行数
        
                    'Detail
                    arrProcInfo(0, j) = j       '行数
                    arrProcInfo(1, j) = "'" & ProcLine    'ソース
                    arrProcInfo(2, j) = "'" & ProcKind    'proc種類
                    arrProcInfo(3, j) = "'" & ProcName  'proc名
                    
                Else    'プロシージャの先頭行以外
                    'Detail
                    arrProcInfo(0, j) = j       '行数
                    arrProcInfo(1, j) = "'" & ProcLine 'ソース
                End If
                ProcLine = ""
            Else
            End If
        Else
        End If
    Next
        '最終行から行数を求め格納する
        arrProcSummary(3, k) = j - arrProcSummary(2, k - 1) '行数
        
'__FormatSetting__
    'Detailの設定
    Range(Sheets("Detail").Cells(1, col), Sheets("Detail").Cells(UBound(arrProcInfo, 2) + 1, col + 3)) = WorksheetFunction.Transpose(arrProcInfo)
    With Sheets("Detail")
        .Cells(1, col) = "行"
        .Cells(1, col + 1) = "ソース"
        .Cells(1, col + 2) = "proc種類"
        .Cells(1, col + 3) = "proc名"
        .Cells(1, col + 4) = "一致度"
    End With
        
    'Summaryの設定
    Range(Sheets("Summary").Cells(1, col), Sheets("Summary").Cells(UBound(arrProcSummary, 2) + 1, col + 3)) = WorksheetFunction.Transpose(arrProcSummary)
    With Sheets("Summary")
        .Cells(1, col) = "モジュール名"
        .Cells(1, col + 1) = "プロシージャ名"
        .Cells(1, col + 2) = "開始行"
        .Cells(1, col + 3) = "行数"
        .Cells(1, col + 4) = "行数差"
        .Cells(1, col + 5) = "一致度"
        .Cells(1, col + 6) = "重複"
    End With
    
    'End の設定
    Sheets("Detail").Cells(j + 2, col) = "END"
    Sheets("Detail").Cells(j + 2, col + 2) = "END"
    
End Sub

Function IsProcLine(Optional strLine As String) As Boolean
'プロシージャ・プロパティ定義行の判

    Select Case True
        Case Left(strLine, 1) = "'"
            IsProcLine = False
        Case strLine Like "Sub*"
            IsProcLine = True
        Case strLine Like "Function*"
            IsProcLine = True
        Case strLine Like "Property*"
            IsProcLine = True
        Case Else
            IsProcLine = False
        End Select
End Function

Function IsWords(word As String, lst_words As Variant) As String
'設定キーワードか否か判定

Dim i As Integer
'__init__
    IsWords = Empty
    word = Replace(word, " ", "")

'__main__
    For i = LBound(lst_words) To UBound(lst_words)
        If lst_words(i, 1) = Left(word, Len(lst_words(i, 1))) Then
            IsWords = lst_words(i, 1)
            Exit Function
        Else
        End If
    Next

'__return__
    IsWords = Empty
End Function

