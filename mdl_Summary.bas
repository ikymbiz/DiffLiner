Attribute VB_Name = "mdl_Summary"
Option Explicit

Const sheetName As String = "Summary"

Sub コード分析()
    '変数宣言
    Dim intWriteRow As Long         '書き込む行番号
    Dim intMdlTotalNum As Long    'モジュール総数
    Dim intProTotalNum As Long    'プロシージャ総数
    Dim intCodeTotalNum As Long  'コード総数
    Dim strMdlName As String        'モジュール名
    Dim strProName As String        'プロシージャ名
    Dim intCodeNum As Long         'コード数
    Dim intRowCount As Long         'データの行数カウント用
    
    '①モジュール数を取得
    Dim intMdlNum As Integer 'モジュール数
    intMdlTotalNum = ActiveWorkbook.VBProject.VBComponents.Count
    
    '②モジュール数分処理をループ
    Dim i As Long, j As Long
    intProTotalNum = 0
    intCodeTotalNum = 0
    intRowCount = 1
    intWriteRow = 9
    For i = 1 To intMdlTotalNum
       With ActiveWorkbook.VBProject.VBComponents(i)
            '③モジュール名を設定
            strMdlName = .Name
            With .CodeModule
                'プロシージャ数分処理をループ
                For j = 1 To .CountOfLines
                    If strProName <> .ProcOfLine(j, 0) Then
                        '④プロシージャ名・コード数をセット
                        strProName = .ProcOfLine(j, 0)
                        intCodeNum = .ProcCountLines(strProName, 0)
                        
                        '⑤No、モジュール名、プロシージャ名、コード数をセルに書き込む
                        With Worksheets(sheetName)
                            .Cells(intWriteRow, 2).Value = intRowCount  'No
                            .Cells(intWriteRow, 3).Value = strMdlName   'モジュール名
                            .Cells(intWriteRow, 4).Value = strProName   'プロシージャ名
                            .Cells(intWriteRow, 5).Value = intCodeNum  'コード数
                        End With
                                                
                        'サマリーデータ用のプロシージャ総数、コード総数を更新
                        intProTotalNum = intProTotalNum + 1
                        intCodeTotalNum = intCodeTotalNum + intCodeNum
                        
                        '次に書き込む行数を更新
                        intWriteRow = intWriteRow + 1
                        
                        'Noを更新
                        intRowCount = intRowCount + 1
                        
                    End If
                Next j
                
                'プロシージャ名を空に戻す
                strProName = ""
            End With
        End With
    Next i
    
    '⑥サマリーデータを登録
    Worksheets(sheetName).Cells(3, 3).Value = intMdlTotalNum
    Worksheets(sheetName).Cells(4, 3).Value = intProTotalNum
    Worksheets(sheetName).Cells(5, 3).Value = intCodeTotalNum
    
End Sub
