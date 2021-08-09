Attribute VB_Name = "mdl_Lib"
Option Explicit

Function rowLast(sheetName As String, column As Long) As Long
'最終行を求める

'Arg
'sheetName     検索するシート名
'column            検索する列番号

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
'最終列を求める

'Arg
'sheetName     検索するシート名
'column            検索する列番号

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
'セルの幅を設定する

    Range(Columns(StartCol), Columns(LastCol)).ColumnWidth = Width
End Function

Function SetHight(Height As Integer, StartRow As Integer, LastRow As Integer)
'セルの高さを設定する

    Range(Rows(StartRow), Rows(LastRow)).RowHeight = Height
End Function

Function SetFileReadOnly()
'ファイルを読取り専用にする

On Error Resume Next
    ActiveWorkbook.Saved = True
    ActiveWorkbook.ChangeFileAccess (xlReadOnly)
End Function

Function SetFileReadWrite()
'ファイルを読取り専用を解除する

On Error Resume Next
    ActiveWorkbook.Saved = True
    ActiveWorkbook.ChangeFileAccess (xlReadWrite)
End Function

Function IsReadOnly()
'ファイルが読み取り専用か確認する

    IsReadOnly = ActiveWorkbook.ReadOnly
End Function

Function KillOwn()
'プログラムファイル自身を削除する
'読み取り専用で開き、読み取り元ファイルを削除する

    Call SetFileReadOnly
    Kill ThisWorkbook.FullName
End Function

Function CellColor(rngR As Range, _
                                intColorR As Long, intColorG As Long, intColorB As Long, _
                                Optional dblTintAndShade As Double)
'RGBスケールでセルの色を変える

'RGBパラメータ
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
'セルの色設定をクリアする
    With rngR.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
    End With
End Function

Function GetFilePath() As String
'ダイアログからファイルを選択し、ファイルパスを取得する

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
'ダイアログからフォルダを選択し、パスを取得する

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
'引数で指定されたファイル名を取得する

'Arg     ExtensionFlg
'    True：Returnに拡張子あり
'     False:Returnに拡張子なし

    If ExtensionFlg = True Then
        GetFileName = Mid(FilePath, InStrRev(FilePath, "\") + 1)
    Else
        GetFileName = Replace(FilePath, Left(FilePath, InStrRev(FilePath, "\")), "")
        GetFileName = Replace(GetFileName, GetExtension(FilePath), "")
        GetFileName = Left(GetFileName, Len(GetFileName) - 1)
    End If
End Function

Function GetExtension(FilePath As String) As String
'引数で指定されたファイルの拡張子を返す

Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetExtension = FSO.GetExtensionName(FilePath)
End Function

Function GetPCName() As String
'PCの名前を取得する

Dim WshNetworkObject As Object

    Set WshNetworkObject = CreateObject("Wscript.Network")
    GetPCName = WshNetworkObject.ComputerName
End Function

'Function GetUserID() As String
''ユーザIDを取得する
'
'Dim objSysInfo As Object
'Dim objUser As Object
'
'    Set objSysInfo = CreateObject("ADSysteminfo")
'    Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
'    GetUserID = objUser.Name
'End Function

Function IsExist(FilePath As String) As Boolean
'ファイル、ディレクトリの存在確認をする

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
'指定のウィンドウがアクティブか確認する

Dim StartTime As Single
Dim ElapesedTime As Single

On Error Resume Next
    
    '開始時間セット
    StartTime = Timer
    
    '一定時間の間、一定間隔ごとに処理を試みる
    Do While ElapesedTime < WaitTime     '経過時間 <= 間隔(秒)
        
        '対象画面が起動しているか確認する
        AppActivate (Title)
        
        If Err = 0 Then
            IsAppActivate = True
            Exit Function
        Else
'            Debug.Print Err
        End If
        Err.Clear
        WaitTimeFor (0.1)                           '処理間隔
        ElapesedTime = Timer - StartTime     '経過時間算出
        DoEvents
    Loop
    
    '画面が見つからないときはFlaseで返す
    On Error GoTo 0
        IsAppActivate = False
End Function

Function OpenDir(DirPath As String, Optional WaitTime As Single = 0.7)
'フォルダパスを指定してディレクトリを開く

Dim StartTime As Single

    If IsExist(fokderpath) = False Then GoTo errExist
    
    Shell "C:\Windows\Explore.exe" & FolderPath, vbNormalFocus
    WaitTimeFor (WaitTime)
    StartTime = Timer
    
    'フォルダが表示されるまで待つ
    '５秒待って表示されなかったらエラーを出す
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
'時刻を文字列で"hhnn"で返す

    strTime = Format(Time, "hhnn")
End Function

Function WaitTimeFor(WaitSecounds As Single)
'指定の秒数処理を待機させる

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
'txt/log/basファイルを読込み返す
'対応する文字コード ->ANSI

Dim buf As String
Dim buf_above As String
Dim array_buf() As Variant
Dim i As Long

'__init__
    Open path For Input As #1
    Erase array_buf
    i = 0 'I:配列番号

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
        
        '文末に'_'がある場合は次の行とマージする
        If buf_above <> "" Then buf = Trim(buf)
        
        If Right(buf, 2) = "_" Then
            buf = Left(buf, Len(buf) - 1)
            buf_above = buf_above & buf
            GoTo Continue
        Else
        End If
        
        'Excelシートに出力する
        If OutputFlg = True Then
            If sheetName = Empty Then
                Cells(row_n, col_n) = "'" & buf_above & buf
                buf_above = ""
            End If
        
        '配列に格納する
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
    Call AppRun("ボイス レコーダー")
    If IsAppActivate("ボイス レコーダー") = True Then
    Debug.Print
    WaitTimeFor (0.5)
        AppActivate ("ボイス レコーダー")

        SendKeys "^r"
        AppActivate (ThisWorkbook.Name)
        Application.WindowState = xlMaximized
    Else
    End If

End Sub


Function AppRun(AppName As String)
'他のアプリケーションを起動する

Dim AppUserModelID As String
    Select Case AppName
        Case "GetStarted"
            AppUserModelID = "Microsoft.Getstarted_8wekyb3d8bbwe!App"
        Case "Grooveミュージック"
            AppUserModelID = "Microsoft.ZuneMusic_8wekyb3d8bbwe!Microsoft.ZuneMusic"
        Case "InternetExplorer"
            AppUserModelID = "Microsoft.InternetExplorer.Default"
        Case "MicrosoftEdge"
            AppUserModelID = "Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge"
        Case "ODBCデータ ソース"
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
        Case "音声認識"
            AppUserModelID = "Microsoft.AutoGenerated.{866E6D2B-1C11-B7B7-2EBB-50D2B949F0B4}"
        Case "アラーム&クロック"
            AppUserModelID = "Microsoft.WindowsAlarms_8wekyb3d8bbwe!App"
        Case "イベント ビューアー"
            AppUserModelID = "Microsoft.AutoGenerated.{A5294213-6473-6AEC-9FE8-C4DC1DFDD1B2}"
        Case "エクスプローラー"
            AppUserModelID = "Microsoft.Windows.Explorer"
        Case "カメラ"
            AppUserModelID = "Microsoft.WindowsCamera_8wekyb3d8bbwe!App"
        Case "カレンダー"
            AppUserModelID = "Microsoft.windowscommunicationsapps_8wekyb3d8bbwe!Microsoft.windowslive.Calendar"
        Case "コマンド プロンプト"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\cmd.exe"
        Case "コントロールパネル"
            AppUserModelID = "Microsoft.Windows.ControlPanel"
        Case "コンピューターの管理"
            AppUserModelID = "Microsoft.AutoGenerated.{9BC0C182-2EB1-D242-F4F1-EB60E3978346}"
        Case "コンポーネント サービス"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\comexp.msc"
        Case "サービス"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\services.msc"
        Case "システム情報"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\msinfo32.exe"
        Case "システム構成"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\msconfig.exe"
        Case "スクリーンキーボード"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\osk.exe"
        Case "ステップ記録ツール"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\psr.exe"
        Case "タスクスケジューラ"
            AppUserModelID = "Microsoft.AutoGenerated.{FC715E43-DDE5-9F19-D0C0-A7336C7414D7}"
        Case "タスクマネージャー"
            AppUserModelID = "Microsoft.AutoGenerated.{216F52FF-1A5B-FFC0-E638-5861AAE5CCCE}"
        Case "ディスククリーンアップ"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\cleanmgr.exe"
        Case "デバイス"
            AppUserModelID = "Microsoft.Windows.PCSettings.Devices"
        Case "ドライブのデフラグと最適化"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\dfrgui.exe"
        Case "ナレーター"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\narrator.exe"
        Case "パフォーマンス モニター"
            AppUserModelID = "Microsoft.AutoGenerated.{CEDD44D2-A5B2-1CE7-2EC1-FA113DB2B1CF}"
        Case "ファイル名を指定して実行"
            AppUserModelID = "Microsoft.Windows.Shell.RunDialog"
        Case "フォト"
            AppUserModelID = "Microsoft.Windows.Photos_8wekyb3d8bbwe!App"
        Case "ペイント"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\mspaint.exe"
        Case "ボイス レコーダー"
            AppUserModelID = "Microsoft.WindowsSoundRecorder_8wekyb3d8bbwe!App"
        Case "マップ"
            AppUserModelID = "Microsoft.WindowsMaps_8wekyb3d8bbwe!App"
        Case "メモ帳"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\notepad.exe"
        Case "リソース モニター"
            AppUserModelID = "Microsoft.AutoGenerated.{FC9AE604-08AC-B162-5532-88CBC4322FD7}"
        Case "リモート デスクトップ接続"
            AppUserModelID = "Microsoft.Windows.RemoteDesktop"
        Case "ローカル セキュリティ ポリシー"
            AppUserModelID = "Microsoft.AutoGenerated.{C85B2B53-EA75-7151-6A6A-9728A5752150}"
        Case "ワードパッド"
            AppUserModelID = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}\Windows NT\Accessories\wordpad.exe"
        Case "付箋"
            AppUserModelID = "Microsoft.Windows.StickyNotes"
        Case "印刷の管理"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\printmanagement.msc"
        Case "数式入力パネル"
            AppUserModelID = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}\Common Files\Microsoft Shared\Ink\mip.exe"
        Case "文字コード表"
            AppUserModelID = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}\charmap.exe"
        Case "検索"
            AppUserModelID = "Microsoft.Windows.Cortana_cw5n1h2txyewy!CortanaUI"
        Case "設定"
            AppUserModelID = "Windows.immersivecontrolpanel_cw5n1h2txyewy!Microsoft.Windows.immersivecontrolpanel"
        Case "電卓"
            AppUserModelID = "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App"
        End Select
        
        Shell "explorer.exe shell:AppsFolder\" & AppUserModelID
End Function
