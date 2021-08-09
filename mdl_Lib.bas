Attribute VB_Name = "mdl_Lib"
Option Explicit

Function rowLast(sheetName As String, column As Long) As Long
'�ŏI�s�����߂�

'Arg
'sheetName     ��������V�[�g��
'column            ���������ԍ�

'On Error GoTo errExit
    rowLast = Sheets(sheetName).Columns(column).Find(What:="*", _
                                                                                            LookIn:=xlFormulas, _
                                                                                            SearchOrder:=xlByRows, _
                                                                                            SearchDirection:=xlPrevious).row
                                                                                
    Exit Function
errExit:
    rowLast = 0
End Function

Function colLast(sheetName As String, row As Long) As Integer
'�ŏI������߂�

'Arg
'sheetName     ��������V�[�g��
'column            ���������ԍ�

On Error GoTo errExit
    colLast = Sheets(sheetName).Rows(row).Find(What:="*", _
                                                                                LookIn:=xlFormulas, _
                                                                                SearchOrder:=xlByColumns, _
                                                                                SearchDirection:=xlPrevious).column
                                                                                
    Exit Function
errExit:
    colLast = 0
End Function

Function SetWidth(Width As Integer, StartCol As Integer, LastCol As Integer)
'�Z���̕���ݒ肷��

    Range(Columns(StartCol), Columns(LastCol)).ColumnWidth = Width
End Function

Function SetHight(Height As Integer, StartRow As Integer, LastRow As Integer)
'�Z���̍�����ݒ肷��

    Range(Rows(StartRow), Rows(LastRow)).RowHeight = Height
End Function

Function SetFileReadOnly()
'�t�@�C����ǎ���p�ɂ���

On Error Resume Next
    ActiveWorkbook.Saved = True
    ActiveWorkbook.ChangeFileAccess (xlReadOnly)
End Function

Function SetFileReadWrite()
'�t�@�C����ǎ���p����������

On Error Resume Next
    ActiveWorkbook.Saved = True
    ActiveWorkbook.ChangeFileAccess (xlReadWrite)
End Function

Function IsReadOnly()
'�t�@�C�����ǂݎ���p���m�F����

    IsReadOnly = ActiveWorkbook.ReadOnly
End Function

Function KillOwn()
'�v���O�����t�@�C�����g���폜����
'�ǂݎ���p�ŊJ���A�ǂݎ�茳�t�@�C�����폜����

    Call SetFileReadOnly
    Kill ThisWorkbook.FullName
End Function

Function CellColor(rngR As Range, _
                                intColorR As Long, intColorG As Long, intColorB As Long, _
                                Optional dblTintAndShade As Double)
'RGB�X�P�[���ŃZ���̐F��ς���

'RGB�p�����[�^
'   https://ironodata.info/

    With rngR.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = RGB(intColorR, intColorG, intColorB)
                    .TintAndShade = dblTintAndShade
                    .PatternTintAndShade = 0
    End With
End Function
                                
Function ClearColor(rngR As Range)
'�Z���̐F�ݒ���N���A����
    With rngR.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
    End With
End Function

Function GetFilePath() As String
'�_�C�A���O����t�@�C����I�����A�t�@�C���p�X���擾����

Dim FilePath As String

    FilePath = Application.GetOpenFilename
    
    If FilePath = "False" Then
        GetFilePath = False
        Exit Function
    Else
    End If
    
    GetFilePath = FilePath
End Function

Function GetDirPath() As String
'�_�C�A���O����t�H���_��I�����A�p�X���擾����

Dim FilePath As String

    FilePath = Application.FileDialog(msoFileDialogFolderPicker).Show
    
    If FilePath = "False" Then
        GetDirPath = False
        Exit Function
    Else
    End If
    
    GetDirPath = FilePath
End Function

Function GetFileName(FilePath As String, Optional ExtensionFlg As Boolean = True) As String
'�����Ŏw�肳�ꂽ�t�@�C�������擾����

'Arg     ExtensionFlg
'    True�FReturn�Ɋg���q����
'     False:Return�Ɋg���q�Ȃ�

    If ExtensionFlg = True Then
        GetFileName = Mid(FilePath, InStrRev(FilePath, "\") + 1)
    Else
        GetFileName = Replace(FilePath, Left(FilePath, InStrRev(FilePath, "\")), "")
        GetFileName = Replace(GetFileName, GetExtension(FilePath), "")
        GetFileName = Left(GetFileName, Len(GetFileName) - 1)
    End If
End Function

Function GetExtension(FilePath As String) As String
'�����Ŏw�肳�ꂽ�t�@�C���̊g���q��Ԃ�

Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetExtension = FSO.GetExtensionName(FilePath)
End Function

Function GetPCName() As String
'PC�̖��O���擾����

Dim WshNetworkObject As Object

    Set WshNetworkObject = CreateObject("Wscript.Network")
    GetPCName = WshNetworkObject.ComputerName
End Function

'Function GetUserID() As String
''���[�UID���擾����
'
'Dim objSysInfo As Object
'Dim objUser As Object
'
'    Set objSysInfo = CreateObject("ADSysteminfo")
'    Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
'    GetUserID = objUser.Name
'End Function

Function IsExist(FilePath As String) As Boolean
'�t�@�C���A�f�B���N�g���̑��݊m�F������

    If InStrRev(FilePath, ",") > 0 Then
        If Dir(FilePath) <> "" Then
            IsExist = True
        Else
            IsExist = False
        End If
    Else
        If Dir(FilePath, vbDirectory) <> "" Then
            IsExist = True
        Else
            IsExist = False
        End If
End Function
    End If

Function IsAppActivate(Title As String, Optional WaitTime As Single = 3) As Boolean
'�w��̃E�B���h�E���A�N�e�B�u���m�F����

Dim StartTime As Single
Dim ElapesedTime As Single

On Error Resume Next
    
    '�J�n���ԃZ�b�g
    StartTime = Timer
    
    '��莞�Ԃ̊ԁA���Ԋu���Ƃɏ��������݂�
    Do While ElapesedTime < WaitTime     '�o�ߎ��� <= �Ԋu(�b)
        
        '�Ώۉ�ʂ��N�����Ă��邩�m�F����
        AppActivate (Title)
        
        If Err = 0 Then
            IsAppActivate = True
            Exit Function
        Else
'            Debug.Print Err
        End If
        Err.Clear
        WaitTimeFor (0.1)                           '�����Ԋu
        ElapesedTime = Timer - StartTime     '�o�ߎ��ԎZ�o
        DoEvents
    Loop
    
    '��ʂ�������Ȃ��Ƃ���Flase�ŕԂ�
    On Error GoTo 0
        IsAppActivate = False
End Function

Function OpenDir(DirPath As String, Optional WaitTime As Single = 0.7)
'�t�H���_�p�X���w�肵�ăf�B���N�g�����J��

Dim StartTime As Single

    If IsExist(fokderpath) = False Then GoTo errExist
    
    Shell "C:\Windows\Explore.exe" & FolderPath, vbNormalFocus
    WaitTimeFor (WaitTime)
    StartTime = Timer
    
    '�t�H���_���\�������܂ő҂�
    '�T�b�҂��ĕ\������Ȃ�������G���[���o��
    Do Until IsAppActivate(GetFileName(FolderPath)) = True
        DoEvents
        
        If Timer - StartTime > 5 Then
            OpenDir = False
            Exit Do
        Else
        End If
        
    Loop
    
    OpenDir = True
    Exit Function

errExit:
    OpenDir = False
End Function

Function strTime(Time As Date) As String
'�����𕶎����"hhnn"�ŕԂ�

    strTime = Format(Time, "hhnn")
End Function

Function WaitTimeFor(WaitSecounds As Single)
'�w��̕b��������ҋ@������

Dim StartTime As String
    StartTime = Timer
    
    Do While Timer < StartTime + WaitSecounds
        DoEvents
    Loop
End Function

Function Read_txt(path As String, _
                                Optional row_n As Long = 1, Optional col_n As Long = 1, _
                                Optional sheetName As String, Optional OutputFlg As Boolean = False) _
                                As Variant
'txt/log/bas�t�@�C����Ǎ��ݕԂ�
'�Ή����镶���R�[�h ->ANSI

Dim buf As String
Dim buf_above As String
Dim array_buf() As Variant
Dim i As Long

'__init__
    Open path For Input As #1
    Erase array_buf
    i = 0 'I:�z��ԍ�

'__check__
    If OutputFlg = True Then
        If sheetName = Empty Then
            Debug.Print "Syntax Error: Input sheetName"
            End
        Else
        End If
    Else
        If sheetName <> Empty Then
            Debug.Print "Syntax Error: Not input sheetName"
            End
        Else
        End If
    End If

'__main__
    Do Until EOF(1)
        Line Input #1, buf
        
        '������'_'������ꍇ�͎��̍s�ƃ}�[�W����
        If buf_above <> "" Then buf = Trim(buf)
        
        If Right(buf, 2) = "_" Then
            buf = Left(buf, Len(buf) - 1)
            buf_above = buf_above & buf
            GoTo Continue
        Else
        End If
        
        'Excel�V�[�g�ɏo�͂���
        If OutputFlg = True Then
            If sheetName = Empty Then
                Cells(row_n, col_n) = "'" & buf_above & buf
                buf_above = ""
            End If
        
        '�z��Ɋi�[����
        Else
            ReDim Preserve array_buf(i)
            array_buf(i) = buf_above & buf
            buf_above = ""
            i = i + 1
        End If
Continue:
            row_n = row_n + 1
        Loop
    Close #1

'__return__
    If OutputFlg = False Then Read_txt = array_buf
End Function

Sub TEST_AppRun()
    Call AppRun("�{�C�X ���R�[�_�[")
    If IsAppActivate("�{�C�X ���R�[�_�[") = True Then
    Debug.Print
    WaitTimeFor (0.5)
        AppActivate ("�{�C�X ���R�[�_�[")

        SendKeys "^r"
        AppActivate (ThisWorkbook.Name)
        Application.WindowState = xlMaximized
    Else
    End If

End Sub


Function AppRun(AppName As String)
'���̃A�v���P�[�V�������N������

Dim AppUserModelID As String
    Select Case AppName
        Case "GetStarted"
            AppUserModelID = "Microsoft.Getstarted_8wekyb3d8bbwe!App"
        Case "Groove�~���[�W�b�N"
            AppUserModelID = "Microsoft.ZuneMusic_8wekyb3d8bbwe!Microsoft.ZuneMusic"
        Case "InternetExplorer"
            AppUserModelID = "Microsoft.InternetExplorer.Default"
        Case "MicrosoftEdge"
            AppUserModelID = "Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge"
        Case "ODBC�f�[�^ �\�[�X"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\odbcad32.exe"
        Case "OneDrive"
            AppUserModelID = "Microsoft.SkyDrive.Desktop"
        Case "PC"
            AppUserModelID = "Microsoft.Windows.Computer"
        Case "Snipping Tool"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\SnippingTool.exe"
        Case "PowerShell"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\WindowsPowerShell\v1.0\powershell.exe"
        Case "PowerShell ISE"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\WindowsPowerShell\v1.0\PowerShell_ISE.exe"
        Case "�����F��"
            AppUserModelID = "Microsoft.AutoGenerated.{866E6D2B-1C11-B7B7-2EBB-50D2B949F0B4}"
        Case "�A���[��&�N���b�N"
            AppUserModelID = "Microsoft.WindowsAlarms_8wekyb3d8bbwe!App"
        Case "�C�x���g �r���[�A�["
            AppUserModelID = "Microsoft.AutoGenerated.{A5294213-6473-6AEC-9FE8-C4DC1DFDD1B2}"
        Case "�G�N�X�v���[���["
            AppUserModelID = "Microsoft.Windows.Explorer"
        Case "�J����"
            AppUserModelID = "Microsoft.WindowsCamera_8wekyb3d8bbwe!App"
        Case "�J�����_�["
            AppUserModelID = "Microsoft.windowscommunicationsapps_8wekyb3d8bbwe!Microsoft.windowslive.Calendar"
        Case "�R�}���h �v�����v�g"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\cmd.exe"
        Case "�R���g���[���p�l��"
            AppUserModelID = "Microsoft.Windows.ControlPanel"
        Case "�R���s���[�^�[�̊Ǘ�"
            AppUserModelID = "Microsoft.AutoGenerated.{9BC0C182-2EB1-D242-F4F1-EB60E3978346}"
        Case "�R���|�[�l���g �T�[�r�X"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\comexp.msc"
        Case "�T�[�r�X"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\services.msc"
        Case "�V�X�e�����"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\msinfo32.exe"
        Case "�V�X�e���\��"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\msconfig.exe"
        Case "�X�N���[���L�[�{�[�h"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\osk.exe"
        Case "�X�e�b�v�L�^�c�[��"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\psr.exe"
        Case "�^�X�N�X�P�W���[��"
            AppUserModelID = "Microsoft.AutoGenerated.{FC715E43-DDE5-9F19-D0C0-A7336C7414D7}"
        Case "�^�X�N�}�l�[�W���["
            AppUserModelID = "Microsoft.AutoGenerated.{216F52FF-1A5B-FFC0-E638-5861AAE5CCCE}"
        Case "�f�B�X�N�N���[���A�b�v"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\cleanmgr.exe"
        Case "�f�o�C�X"
            AppUserModelID = "Microsoft.Windows.PCSettings.Devices"
        Case "�h���C�u�̃f�t���O�ƍœK��"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\dfrgui.exe"
        Case "�i���[�^�["
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\narrator.exe"
        Case "�p�t�H�[�}���X ���j�^�["
            AppUserModelID = "Microsoft.AutoGenerated.{CEDD44D2-A5B2-1CE7-2EC1-FA113DB2B1CF}"
        Case "�t�@�C�������w�肵�Ď��s"
            AppUserModelID = "Microsoft.Windows.Shell.RunDialog"
        Case "�t�H�g"
            AppUserModelID = "Microsoft.Windows.Photos_8wekyb3d8bbwe!App"
        Case "�y�C���g"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\mspaint.exe"
        Case "�{�C�X ���R�[�_�["
            AppUserModelID = "Microsoft.WindowsSoundRecorder_8wekyb3d8bbwe!App"
        Case "�}�b�v"
            AppUserModelID = "Microsoft.WindowsMaps_8wekyb3d8bbwe!App"
        Case "������"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\notepad.exe"
        Case "���\�[�X ���j�^�["
            AppUserModelID = "Microsoft.AutoGenerated.{FC9AE604-08AC-B162-5532-88CBC4322FD7}"
        Case "�����[�g �f�X�N�g�b�v�ڑ�"
            AppUserModelID = "Microsoft.Windows.RemoteDesktop"
        Case "���[�J�� �Z�L�����e�B �|���V�["
            AppUserModelID = "Microsoft.AutoGenerated.{C85B2B53-EA75-7151-6A6A-9728A5752150}"
        Case "���[�h�p�b�h"
            AppUserModelID = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}\Windows NT\Accessories\wordpad.exe"
        Case "�t�"
            AppUserModelID = "Microsoft.Windows.StickyNotes"
        Case "����̊Ǘ�"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\printmanagement.msc"
        Case "�������̓p�l��"
            AppUserModelID = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}\Common Files\Microsoft Shared\Ink\mip.exe"
        Case "�����R�[�h�\"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\charmap.exe"
        Case "����"
            AppUserModelID = "Microsoft.Windows.Cortana_cw5n1h2txyewy!CortanaUI"
        Case "�ݒ�"
            AppUserModelID = "Windows.immersivecontrolpanel_cw5n1h2txyewy!Microsoft.Windows.immersivecontrolpanel"
        Case "�d��"
            AppUserModelID = "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App"
        End Select
        
        Shell "explorer.exe shell:AppsFolder\" & AppUserModelID
End Function
