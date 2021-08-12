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
'�c�[���N��������

Dim ans As String
Dim rng_R As Range

Application.ScreenUpdating = False

'__init__
    ans = MsgBox("���������ċN�����܂����B", vbOKCancel)
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
'����������

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
'�����������T�u���[�`��

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
'text�t�@�C��������

Dim ret As String
    ret = GetExtension(path)
    If ret = "text" Or ret = "log" Or "bas" Then
        IsText = True
    Else
        IsText = False
    End If
End Function

Sub main()
'�e�L�X�g�t�@�C�����捞�݁A��r����

Dim pathA As String, pathB As String
Dim ret As String

    MsgBox "���͑Ώۂ̃e�L�X�g�t�@�C����I�����Ă��������B"
    pathA = GetFilePath
    If pathA = "False" Then
        MsgBox "�����𒆎~���܂��B"
        Exit Sub
    Else
    End If
    
    ret = MsgBox("��r�Ώۂ̃e�L�X�g�t�@�C����I�����Ă��������B", vbOKCancel)
    If ret <> vbCancel Then
        pathB = GetFilePath
    Else
        pathB = ""
    End If
    
    Call ImportSource(pathA, pathB)

End Sub

Private Sub ImportSource(pathA As String, pathB As String)
'�e�L�X�g�t�@�C������f�[�^���捞��

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
    
    '�e�L�X�g�t�@�C����ǂݍ���
    Call GetProc(pathA, col:=COL_A)
    
    If pathB <> "" Then
        Call GetProc(pathB, col:=COL_B)
    Else
    End If
    
    'Summary�̈�v�x��Ԃ�
    If SummaryFlg = True Then
        Debug.Print "AnalizeSummary"
    Else
    End If
    
    'Detail��Diff����/�s�ǉ�
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
        
        '�Z���̍ŏI�s��T��
        rowEnd = GetRowEnd
    Loop
    
    '�Z���̍ŏI�s��T��
        rowEnd = GetRowEnd
    
    '�Z���̔w�i�F�ݒ�
        Call CellColor(Range(Cells(1, 1), Cells(1, colLast("Detail", 1))), 100, 100, 100, 0.5)
        Call CellColor(Range(Cells(1, 1), Cells(rowEnd, 1)), 100, 100, 100, 0.5)
        Call CellColor(Range(Cells(1, 5), Cells(rowEnd, 5)), 100, 100, 100, 0.5)
        Call CellColor(Range(Cells(2, 9), Cells(rowEnd, 9)), 100, 100, 100, 0.8)
        
        Application.ScreenUpdating = True
        
        ThisWorkbook.Save
        MsgBox "�������������܂����B"
    
End Sub

Function GetRowEnd()
'�Z���̍ŏI�s��T��

Dim rowEnd As Long
        rowEnd = rowLast("Detail", COL_A)
        If rowEnd < rowLast("Detail", COL_B) Then rowEnd = rowLast("Detail", COL_B)
        
        GetRowEnd = rowEnd
End Function

Sub AnalyzeDetail_onSave()
'�t�@�C���㏑���ۑ����ɍĎZ�o

Dim rowEnd  As Long
Dim row As Long
Dim col As Integer
Dim CurrentRng As Range
Dim i As Long, j As Long

    Application.ScreenUpdating = False
    Set CurrentRng = ActiveCell
    
    '�Z���̍ŏI�s��T��
    rowEnd = GetRowEnd
    
    '�Z���̔w�i�F���N���A����
    Call ClearColor(Cells)
    
    'dummy�s���폜����
    For row = ROW_START To rowEnd
        For col = COL_A To COL_B Step COL_B - COL_A
            If Sheets("Detail").Cells(row, col) = "dummy" Then
                Call DeleteRow(row:=row, col:=col, rowRng:=1, colRng:=3)
                row = row - 1
            Else
            End If
        Next
    Next
    
    '��v�x���폜
    Sheets("Detail").Columns(COL_LS).ClearContents
    
    '��v�x���ĎZ�o
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
    
    '�Z���̍ŏI�s��T��
    rowEnd = GetRowEnd
    
    '�Z���̔w�i�F�ݒ�
    Sheets("Detail").Select
    Call CellColor(Range(Cells(1, 1), Cells(1, colLast("Detail", 1))), 100, 100, 100, 0.5)
    Call CellColor(Range(Cells(1, 1), Cells(rowEnd, 1)), 100, 100, 100, 0.5)
    Call CellColor(Range(Cells(1, 5), Cells(rowEnd, 5)), 100, 100, 100, 0.5)
    Call CellColor(Range(Cells(1, 9), Cells(rowEnd, 9)), 100, 100, 100, 0.8)
    
    '�w�b�_�Đݒ�
    For col = COL_A To COL_B Step COL_B - COL_A
        Sheets("Detail").Cells(1, col) = "�s"
        Sheets("Detail").Cells(1, col + 1) = "�\�[�X"
        Sheets("Detail").Cells(1, col + 2) = "proc���"
        Sheets("Detail").Cells(1, col + 3) = "proc��"
        Sheets("Detail").Cells(1, col + 4) = "��v�x"
    Next
    
    Application.ScreenUpdating = True
    
On Error Resume Next
    CurrentRng.Select
    
End Sub

Function Analyze_Detail(row As Long) As Double
'Detail�V�[�g�̃\�[�X�̊e�s��T�����Ȃ����v�x�̍����s����ׂĕ\������

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
    
    '臒l
    threshold = Sheets("menu").Range("S2")
    
    '�L�[���[�h���X�g
    lst = Range(Sheets("menu").Cells(2, 16), Sheets("menu").Cells(rowLast("menu", 16), 16))
    
'__main__
    Analyze_Detail = 0
    
    For i = row To rowProcEnd(row, COL_A_PROC)
        Set base = Sheets("Detail").Cells(i, COL_A_CODE)
        
        If Sheets("Detail").Cells(i, COL_A) = "Dummy" Or _
            Sheets("Detail").Cells(i, COL_A) = "END" Then GoTo Continue
            
        'COL_A�Ƀ_�~�[�s�łȂ��󔒍s����������LsDist=0�ŏ�������
        If Sheets("Detail").Cells(i, COL_A_CODE) = "" Then
            Sheets("Detail").Cells(i, COL_LS) = 0
            Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME)), 255, 255, 0, 0.5)    '�Z�����n�C���C�g
            GoTo Continue
        Else
        End If
        
        For j = i To rowProcEnd(row, COL_B_proc)
            Set try = Sheets("Detail").Cells(i, COL_B_CODE)
            If Sheets("Detail").Cells(i, COL_B) = "dummy" Or _
                    Sheets("Detail").Cells(i, COL_B) = "END" Then GoTo Continue
                    
            'COL_B�Ƀ_�~�[�s�łȂ��󔒍s����������LsDist=0�ŏ�������
            If Sheets("Detail").Cells(i, COL_B_CODE) = "" Then
                Sheets("Detail").Cells(i, COL_LS) = 0
                Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME)), 255, 255, 0, 0.5)    '�Z�����n�C���C�g
                GoTo Continue
            Else
            End If
            
            '__init__
                score = 0
                matchingFlg = False
                score = LsDist(base.Value, try.Value)
                        
            '      '=== �f�o�b�O�p ====================
            '      '�ΏۃZ�����n�C���C�g
            '      Call ClearColor(base.Offset(-1, 0))
            '      Call ClearColor(try.Offset(-1, 0))
            '      Call CellColor(base, 255, 255, 255, 0, 0)
            '      Call CellColor(try, 255, 255, 0, 0)
            '      Stop
            '      '=============================
                        
            '__main__
                '��v�x��臒l�𒴂��Ă���ꍇ�A��v�t���O�����Ă�
                If score >= threshold Then
                    Sheets("Detail").Cells(j, COL_LS) = score
                    matchingFlg = True
                    Exit For
                Else
                End If
        Next j
        
        '�_�~�[�Ƃ�ǉ�����
        If matchingFlg = False Then
        
        '�s���e���s��v
        If j >= rowProcEnd(row, COL_B_proc) Then j = i  '�}���s���̍ő�l�ݒ�
            Range(Sheets("Detail").Cells(i, COL_B), Sheets("Detail").Cells(j, COL_B_PROCNAME)).Insert shift:=xlDown '�}��
            Sheets("Detail").Cells(i, COL_B) = "dummy"  'dummy�\��
            Sheets("Detail").Cells(j, COL_LS) = 0 '��v�x����
            Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(j, COL_B_PROCNAME)), 255, 255, 0, 0.5)    '�Z�����n�C���C�g
            row = i
            GoTo Restart
            
        '�s���e�̈�v�x��臒l�𒴂��Ă��邩�A�r���s���X�L�b�v���Ă���ꍇ
        ElseIf matchingFlg = True And j - i > 0 Then
            'i��j�̍����s����base��ɒǉ�
            If j >= rowProcEnd(row, COL_B_proc) Then j = i '�}���s���̍ő�l�ݒ�
                Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(j - 1, COL_A_PROCNAME)).Insert shift:=xlDown '�}��
                Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(j - 1, COL_A)) = "dummy" 'dummy�\��
                Range(Sheets("Detail").Cells(i, COL_LS), Sheets("Detail").Cells(j - 1, COL_LS)) = 0 '��v�x����
                Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(j - 1, COL_B_PROCNAME)), 255, 255, 0, 0.5)  '�Z�����n�C���C�g
                row = i
                GoTo Restart
            Else
                If score < 1 Then   '�s���e�̈�v�x��臒l�𒴂��Ă��邪�s��v
                    Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME)), 255, 255, 0, 0.5) '�Z�����n�C���C�g
                Else    '���S�Ɉ�v
                    Call ClearColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME))) '�Z���̐F�t�����Ȃ���
                End If
            End If

Continue:
        '�v���V�[�W�������n�C���C�g
        If IsProcLine(Cells(i, COL_A_PROC)) = True Or IsProcLine(Cells(i, COL_B_proc)) = True Then
            Call CellColor(Range(Sheets("Detail").Cells(i, COL_A), Sheets("Detail").Cells(i, COL_B_PROCNAME)), 255, 150, 200, 0)
        Else
        End If
        
        keyA = IsWords(Cells(i, COL_A_CODE), lst)
        keyB = IsWords(Cells(i, COL_B_CODE), lst)
        
        '�L�[���[�h�n�C���C�g
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
'���݂�proc�̍ŏI�s���擾����

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
'����proc�̊J�n�s���擾����

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
'�ŏ��l�̍��ڂ�Ԃ�

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
'Summary����

Dim baseRng As Range, tryRng As Range

'__main__
    '�range
    Set baseRng = Range(Sheets(summary).Cells(ROW_START, COL_PROC_A), _
                                    Sheets("Summary").Cells(rowLast("Summary", COL_PROC_A), COL_PROC_A))
    '��rrange
    Set tryRng = Range(Sheets("Summary").Cells(ROW_START, COL_PROC_B), _
                                    Sheets("Summary").Cells(rowLast("Summary", COL_PROC_B), COL_PROC_B))

    'Summary�̍s�����A��v�����Z�o����
    Call CheclProcOrder
    
    'T_LS/T_match�@����/���בւ�
    Call LsRng(baseRng, tryRng, "Summary")
    
'__Return__
    '��v�x���Čv�Z���Ԃ�
    AnalyzeSummary = LsRng(baseRng, tryRng, "Summary")
End Function

Function CheckProcOrder()
'Summary�̍s�����A��v�����Z�o����
'proc���̕��я����Range�ɂ��킹���ёւ���

Const colA As Long = 1
Const colB As Long = 5
Const colGap As Long = 9
Const colLS As Long = 10
Dim rowMax As Long, row As Long, col As Long

'__main__
    '�ŏI�s�擾
    rowMax = rowLast("Summary", colStart)
    
    For row = rowStart To rowMax
        For col = colA To colB - 1
            Select Case col
            '���W���[��������
            Case 1
                'skip
            '�v���V�[�W��������
            Case 2
                '��v���Z�o���Z���ɓ���
                Sheets("Summary").Cells(row, colLS) = LsDist(Sheets("Summart").Cells(row, col), Sheets("Summary").Cells(row, col + (colB - colA)))
            End Select
        Next
    Next
End Function

Function LsRng(basRng As Range, tryRng As Range, sheetName As String) As Double
'������̔�r
'Levenshtein�����ŗގ��x�𑪒肵�A��v�x��Ԃ�
'retuen Max:1/min:0

Const SheetNameA As String = "T_LS"
Const SheetNameB As String = "T_match"

Dim LS_matrix() As Variant, matching_matrix() As Variant
Dim i As Integer, j As Integer, cost As Integer
Dim missCnt As Integer

Dim Base_R1C1 As String, Try_R1C1 As String
Dim row_base As Long, col_base As Long, row_try As Long, col_try As Long

'__init__
    '�erange�̊J�n�s�A�J�n����擾����
    Base_R1C1 = basrRng.Address(ReferenceStyle:=xlR1C1)
    Base_R1C1 = Left(Base_R1C1, InStr(Vase_R1C1, ":") - 1)
    row_base = CLng(Replace(Replace(Base_R1C1, (Right(Base_R1C1, InStr(Base_R1C1, "C") - 1)), ""), "R", "")) - 1
    col_base = CLng(Replace(Base_R1C1, (Left(Base_R1C1, InStr(Base_R1C1, "C"))), ""))
    
    Try_R1C1 = tryRng.Address(ReferenceStyle:=xlR1C1)
    Try_R1C1 = Left(Try_R1C1, InStr(Try_R1C1, ":") - 1)
    
    row_try = CLng(Replace(Replace(Try_R1C1, (Right(Try_R1C1, InStr(Try_R1C1, "C") - 1)), ""), "R", "")) - 1
    coltry = CLng(Replace(Try_R1C1, (Left(Try_R1C1, InStr(Try_R1C1, "C"))), ""))
    
    '�z��̍Ē�`
        ReDim LS_matrix(baseRng.Count, tryRng.Count)
        ReDim matching_matrix(baseRng.Count, tryRng.Count)
    
    '__main__
        LsRng = 0
        
        '�Ώ�Range���Ȃ������ꍇ�u�O�v�ŕԂ�
        If baseRng.Count = 0 Then
            LsRng = 0
            Exit Function
        ElseIf tryRng.Count = 0 Then
            LsRng = 0
            Exit Function
        Else
        End If
        
        '�e�z��ɏ����l��^����
        For i = 0 To baseRng.Count
            LS_matrix(i, 0) = i
            matching_matrix(i, 0) = i
        Next i
        
        For j = 0 To tryRng.Count
            LS_matrix(0, j) = j
            matching.matrix(0, j) = j
        Next j
        
        '�z��̉���
        Range(Sheets(SheetNameA).Cells(1, 1), Sheets(SheetNameA).Cells(baseRng.Count + 1, tryRng.Count + 1)) = LS_matrix
        Range(Sheets(SheetNameA).Cells(1, 1), Sheets(SheetNameA).Cells(baseRng.Count + 1, tryRng.Count + 1)) = mathing_matrix
        
        '�e�Z����g������v�x��T������
        For i = 1 To baseRng.Count
            For j = 1 To tryRng.Count
                
                '�l����v����ꍇ�́u�O�v�A��v���Ȃ��ꍇ�́u�P�v
                cost = IIf(Sheets(sheetName).Cells(row_base + i, col_base) = Sheets(sheetName).Cells(row_try + j, col_try), 0, 1)
                
                '��v�g������matrix�Ɋi�[
                '�l����v����ꍇ�́u�P�v�A��v���Ȃ��ꍇ�́u�O�v�ɕϊ����Ċi�[
                matching_matrix(i, j) = Abs(cost - 1)
                
                '�ŏ��R�X�g��matrix�Ɋi�[
                LS_matrix(i, j) = WorksheetFunction.Min(LS_matrix(i - 1, j) + 1, LS_matrix(i, j - 1) + 1, LS_matrix(i - 1, j - 1) + cost)
                  
                'LS_matrix(i - 1, j) + 1               '�v�f�̍폜
                'LS_matrix(i, j -1) + 1                '�v�f�̑}��
                'LS_matrix(i - 1, j - 1) + cost     '�v�f�̒u��
                
''                '�z��̉���
''                '�P�s   �Fi  / baseRng�̗v�f�ԍ�
''                'A��    �Fj / tryRng�̗v�f�ԍ�
''                '�Z���l�F��v/�s��v�̗ݐϒn
''
''                '����  �ŏ��l���E�ɐi��
''                '�}��  �ŏ��l�����ɐi��
''                '�폜  �ŏ��l���E�ɐi��
''                '�u��  �ŏ��l���E���ɐi��
''                Sheets(SheetNameA).Cells(i + 1, j + 1) = LS_matrix(i, j)
            Next j
        Next i
        
        missCnt = LS_matrix(baseRng.Count, tryRng.Count)
        
        '�z��̉���
        Range(Sheets(SheetNameA).Cells(1, 1), Sheets(SheetNameA).Cells(baseRng.Count + 1, tryRng.Count + 1)) = LS_matrix
        Range(Sheets(SheetNameB).Cells(1, 1), Sheets(SheetNameB).Cells(baseRng.Count + 1, tryRng.Count + 1)) = matching_matrix
        
        '���בւ�
        Call LsOrder_Summary(LS_matrix, matching_matrix, baseRng, tryRng)
        
'__return__
    '��v�x��Ԃ�
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
    '��v/�u�����Ă���Z����T������
    '�e�s������m�F����
    For row_tg = COL_START To rowLast(SheetNameA, 1)
    
        '�d���̊m�F
        flg_dup = WorksheetFunction.Sum(Sheets(SheetNameB).Rows(row_tg)) - (row_tg - 1)
        
        '�d���t���O�����Ă�
        If flg_dup > 1 Then
            Sheets("Summary").Cells(row_tg, COL_DUP) = 1
        Else
            Sheets("Summary").Cells(row_tg, COL_DUP) = ""
        End If
        
        '�e����ЂƂ��m�F����
        For col_tg = COL_START To rowLast(SheetNameA, 1)
        
            '��v/�u�����Ă���Z�����m�F����
            If Sheets(SheetNameB).Cells(row_tg, col_tg) = 1 Then
            
                '��v�����Z���̍s�Ɨ񂪓������m�F����
                If row_tg = col_tg Then
                
                    '�����s�ň�v���Ă���
                    Debug.Print
                    Debug.Print row_tg; Sheets("Summary").Cells(row_tg, COL_B_proc)
                    Debug.Print col_tg; Sheets("Summary").Cells(col_tg, COL_B_proc)
                    Exit For
                Else
                    '�d������ꍇ�͓����s�ň�v/�u���t���O���Ȃ����m���߂�
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
    
                    '�d�����Ȃ�/basRng�Ɠ����s�Ɉ�v/�u���t���O���Ȃ��ꍇ�͍s���ړ�����
                    Debug.Print
                    Debug.Print row_tg; Sheets("Summary").Cells(row_tg, COL_B_proc)
                    Debug.Print col_tg; Sheets("Summary").Cells(col_tg, COL_B_proc)
        
                    'Summary�}��
                    Call procLineInsert(rowOrg:=col_tg, row_Dest:=row_tg, rowRng:=1, _
                                                    col:=COL_B_proc, colRng:=3, sheetName:="Summary")
                    'Detail�}��
                    rowOrg = Sheets("Summary").Cells(row_tg, col_CmtBst)
                    rowRng = Sheets("Summary").Cells(row_tg, col_CntBrng)
                    rowDest = Sheets("Summary").Cells(ROW_START, COL_CntBst) + _
                                    WorksheetFunction.Sum(Range(Sheets("Summary").Cells(ROW_START, col_CntBrng), _
                                    Sheets("Summary").Cells(row_tg - 1, col_CntBrng))) + 1
                                    
                    Debug.Print "rowOrg :   " & rowOrg
                    Debug.Print "rowRng :   " & rowRng
                    Debug.Print "rowDest :  " & rowDest 'COL_CntAst��͍s�̓���ւ������Ă��Ȃ��̂ŁAD_RowNo�ōs�ԍ�����������K�v�͂Ȃ�
                
                    Call procLineInsert(rowOrg:=D_RowNo(rowOrg), rowDest:=rowDest, rowRng:=rowRng, col:=COL_Detail)
                    
                    '�Čv�Z
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
'Detail�V�[�g�̃\�[�X�s�ԍ��ɍ��v����Excel�V�[�g�Ƃ�Ԃ�
'return    Excel�̍s�ԍ�

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
'�s�͈͂��w��̍s�̑O�ɑ}������
'rowOrigin      :   �s�ԍ�
'rowRng         :   �s��
'colRng           :    ��
'sheetName    :    �V�[�g��

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
    '�l���擾
    '�R�����g�A�E�g�s��" '�@"���c�����߁A�l��ϐ��Ɋi�[����̂ł͂Ȃ��A�R�s�y�őΉ�����
    
    'rowDest�s��}��
    Range(Sheets(sheetName).Cells(rowDest, col), Sheets(sheetName).Cells(rowDest + rowRng - 1, col + colRng - 1)).Select
    Range(Sheets(sheetName).Cells(rowDest, col), Sheets(sheetName).Cells(rowDest + rowRng - 1 + colRng - 1)).Insert shift:=xlDown
    
    'rowOrg�s���R�s�[
    If rowOrg < rowDest Then
        Range(Sheets(sheetName).Cells(rowOrg, col), Sheets(sheetName).Cells(rowOrg + rowRng - 1, col + colRng - 1)).Select
        Range(Sheets(sheetName).Cells(rowOrg, col), Sheets(sheetName).Cells(rowOrg + rowRng - 1, col + colRng - 1)).Copy
    Else
        Range(Sheets(sheetName).Cells(rowOrg + rowRng, col), Sheets(sheetName).Cells(rowDest + rowRng - 1, col + colRng - 1)).Select
        Range(Sheets(sheetName).Cells(rowOrg + rowRng, col), Sheets(sheetName).Cells(rowDest + rowRng - 1, col + colRng - 1)).PasteSpecial
        
        'rowOrg�s���폜
        If rowOrg < rowDest Then
            Call DeleteRow(rowOrg, col, rowRng:=rowRng, sheetName:=sheetName)
        Else
            Call DeleteRow(rowOrg + rowRng, col, rowRng:=rowRng, sheetName:=sheetName)
        End If

End Function

Function procLineUpDown(rowA As Long, rowB As Long, col As Integer, _
                Optional rowRng As Long = 1, Optional rowRngB As Long = 1, _
                Optional colRng As Integer = 4, Optional sheetName As String = "Detail")
'�s�͈�A�ƍs�͈�B�����ւ���

Dim LineA As Variant, LineB As Variant
Dim i As Long
Dim tmp

'__init__
    Sheets(sheetName).Select

    'rowB��rowA�̌�̍s�łȂ���΂����Ȃ�
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
    
    '�s�͈�A�͍s�͈�B�̈ꕔ���܂�ł͂����Ȃ�
    If rowA + rowRngA - 1 > rowB Then
        Debug.Print (" rowB must be bigger than rowA +rowRngA.")
        MsgBox (" rowB must be bigger than rowA +rowRngA.")
        End
    Else
    End If
    
'__main__
    '�l���擾
    LineA = Range(Sheets(sheetName).Cells(rowA, col), Sheets(sheetName).Cells(rowA + rowRngA - 1, col + colRng - 1))
    LineB = Range(Sheets(sheetName).Cells(rowB, col), Sheets(sheetName).Cells(rowB + rowrngG - 1, col + colRng - 1))
    
    '�s�����ւ�����
    'rowA�̒l��rowB�̒l�ɑ}������
    With Range(Sheets(sheetName).Cells(rowB, col), Sheets(sheetName).Cells(rowB + rowRngA - 1, col + colRng - 1))
        .Select
        .Insert shift:=xlDown
        .NumberFormatLocal = "@"
        .Value = LineB
    End With
    
    Call DeleteRow(rowB + rowRngB, col, rowRng:=rowRngB, sheetName:=sheetName)
                
End Function

Function Compare(valA, valB) As Boolean
'�l���r����
'return     True:��v/False:�s��v

    If valA = valB Then
        Compare = True
    Else
        Compare = False
    End If
End Function

Function InsertRow(row As Long, col As Integer, _
                                Optional rngCnt As Long = 3, Optional sheetName As String = "Detail")
'�w��͈̔͂ɍs��}������
    
    Range(Sheets(sheetName).Cells(row, col), Sheets(sheetName).Cells(row, col + rngCnt)).Insert shift:=xlDown
End Function

Function DeleteRow(row As Long, col As Integer, _
                                Optional rowRng As Long = 1, Optional colRng As Integer = 3, _
                                Optional sheetName As String = "Detail")
'�w��͈̔͂̍s���폜����

    Range(Sheets(sheetName).Cells(row, col), Sheets(sheetName).Cells(row + rowRng - 1, col + colRng)).Select
    Range(Sheets(sheetName).Cells(row, col), Sheets(sheetName).Cells(row + rowRng - 1, col + colRng)).Delete
End Function

Function GetModuleName(ByVal arr As Variant) As String
'���W���[�������擾����
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
'�v���V�[�W���E�v���p�e�B��`�s���̔���

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
'�e�L�X�g�t�@�C�����烂�W���[����ǂݍ���

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
        
        '�R�[�h���e
        ProcLine = ProcLine & Trim(arrProc(i))
        
        Debug.Print i & " " & ProcLine
        
        If ProcLine <> "" Then  '�󔒍s�̓X�L�b�v����
            j = j + 1
            ReDim Preserve arrProcInfo(3, j)
            
            If Right(ProcLine, 1) <> "_" Then   '�����s���Ȃ��ꍇ
                ProcKind = IsProcName(arrProc(i)) '�R�[�h�̎�ނ��擾
            
                If ProcKind <> "" Then  '�v���V�[�W���̐擪�s�̏ꍇ
                    ProcName = Replace(ProcLine, "Public", "")
                    ProcName = Replace(ProcLine, "Private", "")
                    ProcName = Replace(ProcName, ProcKind, "")
                    ProcName = Replace(ProcName, " ", "")
                    ProcName = Left(ProcName, InStrRev(ProcName, "(") - 1)
                    
                    'Summary
                    k = k + 1
                    ReDim Preserve arrProcSummary(3, k)
                    arrProcSummary(0, k) = ModuleName                                '���W���[����
                    arrProcSummary(1, k) = ProcName                                     '�v���V�[�W����
                    arrProcSummary(2, k) = j                                                    '�J�n�s
                    arrProcSummary(3, k - 1) = j - arrProcSummary(2, k - 1)   '�s��
        
                    'Detail
                    arrProcInfo(0, j) = j       '�s��
                    arrProcInfo(1, j) = "'" & ProcLine    '�\�[�X
                    arrProcInfo(2, j) = "'" & ProcKind    'proc���
                    arrProcInfo(3, j) = "'" & ProcName  'proc��
                    
                Else    '�v���V�[�W���̐擪�s�ȊO
                    'Detail
                    arrProcInfo(0, j) = j       '�s��
                    arrProcInfo(1, j) = "'" & ProcLine '�\�[�X
                End If
                ProcLine = ""
            Else
            End If
        Else
        End If
    Next
        '�ŏI�s����s�������ߊi�[����
        arrProcSummary(3, k) = j - arrProcSummary(2, k - 1) '�s��
        
'__FormatSetting__
    'Detail�̐ݒ�
    Range(Sheets("Detail").Cells(1, col), Sheets("Detail").Cells(UBound(arrProcInfo, 2) + 1, col + 3)) = WorksheetFunction.Transpose(arrProcInfo)
    With Sheets("Detail")
        .Cells(1, col) = "�s"
        .Cells(1, col + 1) = "�\�[�X"
        .Cells(1, col + 2) = "proc���"
        .Cells(1, col + 3) = "proc��"
        .Cells(1, col + 4) = "��v�x"
    End With
        
    'Summary�̐ݒ�
    Range(Sheets("Summary").Cells(1, col), Sheets("Summary").Cells(UBound(arrProcSummary, 2) + 1, col + 3)) = WorksheetFunction.Transpose(arrProcSummary)
    With Sheets("Summary")
        .Cells(1, col) = "���W���[����"
        .Cells(1, col + 1) = "�v���V�[�W����"
        .Cells(1, col + 2) = "�J�n�s"
        .Cells(1, col + 3) = "�s��"
        .Cells(1, col + 4) = "�s����"
        .Cells(1, col + 5) = "��v�x"
        .Cells(1, col + 6) = "�d��"
    End With
    
    'End �̐ݒ�
    Sheets("Detail").Cells(j + 2, col) = "END"
    Sheets("Detail").Cells(j + 2, col + 2) = "END"
    
End Sub

Function IsProcLine(Optional strLine As String) As Boolean
'�v���V�[�W���E�v���p�e�B��`�s�̔�

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
'�ݒ�L�[���[�h���ۂ�����

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

