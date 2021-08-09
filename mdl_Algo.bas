Attribute VB_Name = "mdl_Algo"
Option Explicit

Function LsDist(baseText As String, tryText As String) As Double
'# •¶š—ñ‚Ì”äŠr
'Levenshtein‹——£‚Å—Ş—“x‚ğ‘ª’è‚µ¤ˆê’v“x‚ğ•Ô‚·
'Levenshtein‹——£
'   Œ³‚Ì•¶š—ñ‚ğ‰½•¶š•ÏX‚·‚ê‚ÎAvl•¶š—ñ‚É‚È‚é‚©‰ñ”‚Å‘ª‚é

'Arg
'Param1(String):    baseText    ”äŠrŒ³‚Ì•¶š—ñ
'Param2(String):    tryText       ”äŠr‘ÎÛ‚Ì•¶š—ñ

'Return(Double):    •¶š—ñ‚Ìˆê’v“x  min:0/Max:1

Dim matrix As Variant
Dim i As Integer, j As Integer, cost As Integer
Dim missCnt As Integer

    LsDist = 0
    
    If (baseText = tryText) Then
        LsDist = Format(1, "0.00")
        Exit Function
    End If
    If (Len(baseText) = 0) Then
        LsDist = Format(0, "0.00")
        Exit Function
    End If
    
    ReDim matrix(Len(baseText), Len(tryText))

    For i = 0 To Len(baseText)
        matrix(i, 0) = i
    Next i
    
    For j = 0 To Len(tryText)
        matrix(0, j) = j
    Next j
    
    For i = 1 To Len(baseText)
        For j = 1 To Len(tryText)
            cost = IIf(Mid$(baseText, i, 1) = Mid$(tryText, j, 1), 0, 1)
            matrix(i, j) = WorksheetFunction.Min(matrix(i - 1, j) + 1, matrix(i, j - 1) + 1, matrix(i - 1, j - 1) + cost)
            
                 'matrix(i - 1, j) + 1              '—v‘f‚Ìíœ
                 'matrix(i, j - 1) + 1              '—v‘f‚Ì‘}“ü
                 'matrix(i - 1, j - 1) + cost    '—v‘f‚Ì’uŠ·
        Next j
    Next i
    
    missCnt = matrix(Len(baseText), Len(tryText))
    
    'ˆê’v“x‚ğ•Ô‚·
'    LsDist = 1-(missCnt / Len(baseText))
    LsDist = (missCnt / Len(baseText))
    LsDist = 1 - LsDist / Len(baseText)
    LsDist = Format(LsDist, "0.00")
    If LsDist < 0 Then LsDist = Format(0, "0.00")
End Function

