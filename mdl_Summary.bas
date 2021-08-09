Attribute VB_Name = "mdl_Summary"
Option Explicit

Const sheetName As String = "Summary"

Sub �R�[�h����()
    '�ϐ��錾
    Dim intWriteRow As Long         '�������ލs�ԍ�
    Dim intMdlTotalNum As Long    '���W���[������
    Dim intProTotalNum As Long    '�v���V�[�W������
    Dim intCodeTotalNum As Long  '�R�[�h����
    Dim strMdlName As String        '���W���[����
    Dim strProName As String        '�v���V�[�W����
    Dim intCodeNum As Long         '�R�[�h��
    Dim intRowCount As Long         '�f�[�^�̍s���J�E���g�p
    
    '�@���W���[�������擾
    Dim intMdlNum As Integer '���W���[����
    intMdlTotalNum = ActiveWorkbook.VBProject.VBComponents.Count
    
    '�A���W���[���������������[�v
    Dim i As Long, j As Long
    intProTotalNum = 0
    intCodeTotalNum = 0
    intRowCount = 1
    intWriteRow = 9
    For i = 1 To intMdlTotalNum
       With ActiveWorkbook.VBProject.VBComponents(i)
            '�B���W���[������ݒ�
            strMdlName = .Name
            With .CodeModule
                '�v���V�[�W���������������[�v
                For j = 1 To .CountOfLines
                    If strProName <> .ProcOfLine(j, 0) Then
                        '�C�v���V�[�W�����E�R�[�h�����Z�b�g
                        strProName = .ProcOfLine(j, 0)
                        intCodeNum = .ProcCountLines(strProName, 0)
                        
                        '�DNo�A���W���[�����A�v���V�[�W�����A�R�[�h�����Z���ɏ�������
                        With Worksheets(sheetName)
                            .Cells(intWriteRow, 2).Value = intRowCount  'No
                            .Cells(intWriteRow, 3).Value = strMdlName   '���W���[����
                            .Cells(intWriteRow, 4).Value = strProName   '�v���V�[�W����
                            .Cells(intWriteRow, 5).Value = intCodeNum  '�R�[�h��
                        End With
                                                
                        '�T�}���[�f�[�^�p�̃v���V�[�W�������A�R�[�h�������X�V
                        intProTotalNum = intProTotalNum + 1
                        intCodeTotalNum = intCodeTotalNum + intCodeNum
                        
                        '���ɏ������ލs�����X�V
                        intWriteRow = intWriteRow + 1
                        
                        'No���X�V
                        intRowCount = intRowCount + 1
                        
                    End If
                Next j
                
                '�v���V�[�W��������ɖ߂�
                strProName = ""
            End With
        End With
    Next i
    
    '�E�T�}���[�f�[�^��o�^
    Worksheets(sheetName).Cells(3, 3).Value = intMdlTotalNum
    Worksheets(sheetName).Cells(4, 3).Value = intProTotalNum
    Worksheets(sheetName).Cells(5, 3).Value = intCodeTotalNum
    
End Sub
