Attribute VB_Name = "Module4"
Option Explicit

Public Sub Formulas(wname As String)
    On Error GoTo ErrorProc
    WriteForm.TextBox2.Value = "Creating Formulas"
    WriteForm.Repaint
    Call Form1Formulas(wname)
    Call Form4Formulas(wname)                    '2p2
    Call Form5Formulas(wname)                    '2p2
    Call Form7Formulas(wname)                    '2p2
    Call Form8AFormulas(wname)
    Call Form8DFormulas(wname)
    Call Form8EFormulas(wname)
    Call Form8FFormulas(wname)
    Call Form9Formulas(wname)                    '2p2
    Call Form10Formulas(wname)                   '2p2
    Call Form11AFormulas(wname)
    Call Form12Formulas(wname)
    Call Form13Formulas(wname)                   '2p2
    Call Form14Formulas(wname)                   '2p2
    WriteForm.Hide
    Call Speedon(False)
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Formulas " & Err.Description
    Err.Clear
End Sub

Public Sub ClearForms(wname As String)
    On Error GoTo ErrorProc
    WriteForm.TextBox2.Value = "Clearing Data"
    WriteForm.Repaint
    Dim ws As Worksheet
    Dim MaxRowCount As Long                      ' Do not use Integer, may be too small and cause overflow
    Dim i As Long
    MaxRowCount = 0
    'For Each ws In ActiveWorkbook.Worksheets 'Try replacing ActiveWorkbook with Workbooks(wname)
    'For Each ws In Workbooks(wname).Worksheets
    '  If ws.UsedRange.Rows.Count > MaxRowCount Then
    '    MaxRowCount = ws.UsedRange.Rows.Count
    '  End If
    'Next
    With Workbooks(wname)
        .Worksheets("F01").Range("D3:D69").ClearContents
        .Worksheets("F01").Range("D3:D22").NumberFormat = "@" 'Force format to be text so dates aren't autmoatically changed
        .Worksheets("F01").Range("G48").ClearContents 'Erase restart file name
        MaxRowCount = .Worksheets("F02A").UsedRange.Rows.Count
        .Worksheets("F02A").Range("B5:F" & MaxRowCount).ClearContents
        .Worksheets("F02A").Range("B5:F" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F02B").UsedRange.Rows.Count
        .Worksheets("F02B").Range("B5:E" & MaxRowCount).ClearContents
        .Worksheets("F02B").Range("B5:E" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F03").UsedRange.Rows.Count
        .Worksheets("F03").Range("B5:AX" & MaxRowCount).ClearContents
        .Worksheets("F03").Range("B5:AX" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F04").UsedRange.Rows.Count
        .Worksheets("F04").Range("B5:J" & MaxRowCount).ClearContents
        .Worksheets("F04").Range("B5:J" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F05").UsedRange.Rows.Count
        .Worksheets("F05").Range("B5:AA" & MaxRowCount).ClearContents
        .Worksheets("F05").Range("B5:AA" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F06").UsedRange.Rows.Count
        .Worksheets("F06").Range("B5:AY" & MaxRowCount).ClearContents
        .Worksheets("F06").Range("B5:AY" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F07").UsedRange.Rows.Count
        .Worksheets("F07").Range("B5:V" & MaxRowCount).ClearContents
        .Worksheets("F07").Range("B5:V" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F07C").UsedRange.Rows.Count
        .Worksheets("F07C").Range("B5:H" & MaxRowCount).ClearContents
        .Worksheets("F07C").Range("B5:H" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F07D").UsedRange.Rows.Count
        .Worksheets("F07D").Range("B5:I" & MaxRowCount).ClearContents
        .Worksheets("F07D").Range("B5:I" & MaxRowCount).NumberFormat = "General"
        'Headers of Form 8 are cleared with loop listed at end of procedure
        MaxRowCount = .Worksheets("F08A").UsedRange.Rows.Count
        .Worksheets("F08A").Range("B5:I" & MaxRowCount).ClearContents
        .Worksheets("F08A").Range("B5:I" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F08B").UsedRange.Rows.Count
        .Worksheets("F08B").Range("B4:BI" & MaxRowCount).ClearContents
        .Worksheets("F08B").Range("B4:BI" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F08C").UsedRange.Rows.Count
        .Worksheets("F08C").Range("B5:FE" & MaxRowCount).ClearContents
        .Worksheets("F08C").Range("B5:FE" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F08D").UsedRange.Rows.Count
        .Worksheets("F08D").Range("B7:BI" & MaxRowCount).ClearContents
        .Worksheets("F08D").Range("B7:BI" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F08E").UsedRange.Rows.Count
        .Worksheets("F08E").Range("B6:CW" & MaxRowCount).ClearContents
        .Worksheets("F08E").Range("B6:CW" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F08F").UsedRange.Rows.Count
        .Worksheets("F08F").Range("D5:BI" & MaxRowCount).ClearContents
        .Worksheets("F08F").Range("D5:BI" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F09").UsedRange.Rows.Count
        .Worksheets("F09").Range("F3:Y" & MaxRowCount).ClearContents
        .Worksheets("F09").Range("F3:Y" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F10").UsedRange.Rows.Count
        .Worksheets("F10").Range("B5:I" & MaxRowCount).ClearContents
        .Worksheets("F10").Range("B5:I" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F11A").UsedRange.Rows.Count
        .Worksheets("F11A").Range("B5:G" & MaxRowCount).ClearContents
        .Worksheets("F11A").Range("B5:G" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F11B").UsedRange.Rows.Count
        If MaxRowCount < 3 Then MaxRowCount = 3
        .Worksheets("F11B").Range("B3:AO" & MaxRowCount).ClearContents
        .Worksheets("F11B").Range("B3:AO" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F12").UsedRange.Rows.Count
        .Worksheets("F12").Range("D2:D3").ClearContents 'Clear Temp Tab
        .Worksheets("F12").Range("B5:H" & MaxRowCount).ClearContents
        .Worksheets("F12").Range("B5:H" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F13").UsedRange.Rows.Count
        .Worksheets("F13").Range("B4:E" & MaxRowCount).ClearContents
        .Worksheets("F13").Range("B4:E" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F14AB").UsedRange.Rows.Count
        .Worksheets("F14AB").Range("B4:L" & MaxRowCount).ClearContents
        .Worksheets("F14AB").Range("B4:L" & MaxRowCount).NumberFormat = "General"
        MaxRowCount = .Worksheets("F14C").UsedRange.Rows.Count
        .Worksheets("F14C").Range("B4:K" & MaxRowCount).ClearContents
        .Worksheets("F14C").Range("B4:K" & MaxRowCount).NumberFormat = "General"
        For i = 1 To 20                          'Clear header rows on Form 8
            With Workbooks(wname).Worksheets("F08D")
                .Cells(3, 1 + i * 3).ClearContents
                '.Range(.Cells(2, 1 + i * 3), .Cells(3, 1 + i * 3)).NumberFormat = "General"
                .Cells(3, 1 + i * 3).NumberFormat = "General"
            End With
            With Workbooks(wname).Worksheets("F08F")
                .Cells(3, 1 + i * 3).ClearContents
                .Cells(3, 1 + i * 3).NumberFormat = "General"
            End With
        Next i
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in Clearing and Formatting of Forms: " & Err.Description
    Err.Clear
End Sub

Private Sub Form1Formulas(wname As String)
    On Error GoTo ErrorProc
    With Workbooks(wname).Worksheets("F01")
        .Range("D23").Value2 = "17."
        .Range("D24").Value2 = "7."
        .Range("D25").Value2 = "2022."
        .Range("D26").Value2 = "0."
        .Range("D27").Value2 = "2."
        .Range("D28").Value2 = "1."
        .Range("D29:D33").Value2 = "0."
        .Range("D34").Formula = "=MAX(COUNT('F03'!B:B),1)" 'Form 1D
        .Range("D35").Formula = "=MAX(COUNT('F02A'!B:B,'F02B'!B:B),1)"
        .Range("D36").Formula = "=COUNT(F02B!B:B)"
        .Range("D37").Formula = "=MAX(COUNT('F06'!B:B),2)"
        .Range("D38").Value2 = "1.0" 'Added to prevent errors to branched junctions
        .Range("D39").Formula = "=COUNTA('F04'!B:B)-2"
        .Range("D40").Formula = "=COUNTA('F07'!B:B)-2"
        .Range("D41").Formula = "=COUNTA(F08A!B:B)-2"
        .Range("D42").Formula = "=COUNTA('F09'!3:3)-2"
        .Range("D43").Formula = "=COUNT(F11A!B:B)"
        .Range("D44").Value2 = "1."
        .Range("D45").Formula = "=COUNT('F10'!B:B)"
        .Range("D46").Formula = "=COUNT(F07C!B:B)"
        .Range("D47:D48").Value2 = "0."
        .Range("D57").Value2 = "68."
        .Range("D58:D63").Value2 = "0."
        .Range("D64").Value2 = "0.2"
        .Range("D65").Formula = "=COUNT(F07D!B:B)"
        .Range("D66").Formula = "30."
        .Range("D67").Formula = "0.5"
        .Range("D68").Formula = "0."
        .Range("D69").Formula = "=COUNTA(F14AB!B:B)-2"
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form1Formulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form4Formulas(wname As String)       '2p2
    On Error GoTo ErrorProc
    With Workbooks(wname).Worksheets("F04")
        .Range("J1").Formula = "=NUHS"
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form5Formulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form5Formulas(wname As String)       '2p2
    On Error GoTo ErrorProc
    With Workbooks(wname).Worksheets("F05")
        .Range("P1").Formula = "=NUHS"
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form4Formulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form7Formulas(wname As String)       '2p2
    On Error GoTo ErrorProc
    With Workbooks(wname).Worksheets("F07C")
        .Range("H1").Formula = "=NIFT"
    End With
    With Workbooks(wname).Worksheets("F07D")
        .Range("I1").Formula = "=NACFT"
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form7Formulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form8AFormulas(wname As String)      '(wname As String)
    On Error GoTo ErrorProc
    Dim i, rF08a, cF08a, rF08b, cF08b, cF08c As Integer
    Dim numGroup, numSect As String
    With Workbooks(wname).Worksheets("F08A")
        rF08a = 5                                'Starting row on F08A
        cF08a = 4                                'Starting column on F08A
        cF08b = 2                                'Column for F08b
        cF08c = 2                                'Column for F08c
        For i = 0 To 19 Step 1
            'Create formula for Number of Group Train Enter
            numGroup = "=if(R[0]C[-2]<>"""",COUNT(F08B!C" & cF08b + i * 3 & " )+1,"""")"
            .Range("D" & rF08a + i).FormulaR1C1 = numGroup
            'Create formula for Number of Track Sections in Route
            numSect = "=if(R[0]C[-3]<>"""",COUNT(F08C!C" & cF08c + i * 8 & " ),"""")"
            .Range("E" & rF08a + i).FormulaR1C1 = numSect
        Next
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form8AFormulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form8DFormulas(wname As String)
    On Error GoTo ErrorProc
    Dim i, rF08d, cF08d, cCount, cInterval, linknum As Integer
    Dim numStops, routelink As String
    With Workbooks(wname).Worksheets("F08D")
        rF08d = 2                                'Starting row for formula F08D
        cF08d = 4                                'Starting column for formula on F08D
        cCount = 2                               'Column to start counting
        cInterval = 3                            'Number of columns over for next count
        For i = 0 To 19 Step 1
            'Create formula for Number of Group Train Enter
            linknum = i + 5
            routelink = "=F08A!$B" & linknum
            numStops = "=COUNT(F08D!C" & cCount + i * cInterval & " )"
            '.Range("D" & cF08a + i).FormulaR1C1 = numStops
            .Cells(1, cF08d).Formula = routelink
            .Range(.Cells(rF08d, cF08d), .Cells(rF08d, cF08d)).FormulaR1C1 = numStops
            'Create formula for Number of Track Sections in Route
            cF08d = cF08d + cInterval
        Next
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form8DFormulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form8FFormulas(wname As String)
    On Error GoTo ErrorProc
    Dim i, rF08f, cF08f, cCount, cInterval, linknum As Integer
    Dim routelink As String
    With Workbooks(wname).Worksheets("F08F")
        rF08f = 2                                'Starting row for formula F08F
        cF08f = 4                                'Starting column for formula on F08F
        cCount = 2                               'Column to start counting
        cInterval = 3                            'number of columns to move over
        For i = 0 To 19 Step 1
            linknum = i + 5                      'Updated from 4 to 5
            routelink = "=F08A!$B" & linknum
            .Cells(1, cF08f - 1).Formula = routelink
            .Range(.Cells(rF08f, cF08f), .Cells(rF08f, cF08f)).FormulaR1C1 = "=COUNT(R[3]C[0]:R[1002]C[0])"
            'Create formula for Number of Track Sections in Route
            cF08f = cF08f + cInterval
        Next
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form8DFormulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form8EFormulas(wname As String)
    On Error GoTo ErrorProc
    Dim i, rF08e, cF08e, cCount, cInterval, linkcol, linknum As Integer
    Dim numSec, routelink As String
    With Workbooks(wname).Worksheets("F08E")
        rF08e = 2                                'Starting row for formula F08e
        cF08e = 6                                'Starting column for formula on F08e
        cCount = 2                               'Column to start counting
        cInterval = 5                            'number of columns to move over
        For i = 0 To 19 Step 1
            'Create formula for Number of Group Train Enter
            linkcol = cF08e - 2
            linknum = i + 5
            routelink = "=F08A!$B" & linknum
            .Cells(1, linkcol).Formula = routelink
            numSec = "=COUNT(R[4]C[-4]:R[103]C[-4])"
            .Range(.Cells(rF08e, cF08e), .Cells(rF08e, cF08e)).FormulaR1C1 = numSec
            'Create formula for Number of Track Sections in Route
            cF08e = cF08e + cInterval
        Next
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form8EFormulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form9Formulas(wname As String)       '2p2
    On Error GoTo ErrorProc
    With Workbooks(wname).Worksheets("F09")
        .Range("D1").Formula = "=TPO"
        .Range("E1").Formula = "=TPO"
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form10Formulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form10Formulas(wname As String)      '2p2
    On Error GoTo ErrorProc
    With Workbooks(wname).Worksheets("F10")
        .Range("G1").Formula = "=TPO"
        .Range("I1").Formula = "=NTOI"
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form10Formulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form11AFormulas(wname As String)
    On Error GoTo ErrorProc
    With Workbooks(wname).Worksheets("F11a")
        .Range("G1").Formula = "=ecleo"          'Reference to Form 1 C, Environmental Control Load Evaluation Option
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form11Formulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form12Formulas(wname As String)
    On Error GoTo ErrorProc
    With Workbooks(wname).Worksheets("F12")
        .Range("D3").Formula = "=COUNT('F12'!$B:$B )" 'Form 12 Equation for Print Groups
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form12Formulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form13Formulas(wname As String)      '2p2
    On Error GoTo ErrorProc
    With Workbooks(wname).Worksheets("F13")
        .Range("E1").Formula = "=NUHS"           'Form 12 Equation for Print Groups
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form12Formulas: " & Err.Description
    Err.Clear
End Sub

Private Sub Form14Formulas(wname As String)      '2p2
    On Error GoTo ErrorProc
    With Workbooks(wname).Worksheets("F14AB")
        .Range("I1").Formula = "=NCP"            'Form 14 Equation for Print Groups
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Form12Formulas: " & Err.Description
    Err.Clear
End Sub

