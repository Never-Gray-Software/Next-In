Attribute VB_Name = "Write_Input_Files"
'Copyright 2024, Never Gray, Justin Edenbaum P.Eng
'Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
'1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
'2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
'3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS “AS IS” AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'Write out input files

Option Explicit

Public wname As String
Public FormIn, Output, F14C As Worksheet
Public TPO As Integer
Dim TCO As Integer
Dim OFSO As Integer
Public NextRow As Long
Private Active As Boolean

Dim Row As Long
Dim col As Long
Dim NumCol As Long
Dim NextCol As Long

Dim LRtext As Long
Dim NumData As Long
Dim ECL As Long
'Write Procedures

Dim NumUnsteady As Long
Dim NumECZ As Long
Dim InitialReadOption As Long
Dim InitialFileName As String
Dim CPipe As Long
Dim FormInLastRow As Long
Dim NumofSSHS As Long
Dim R_off As Long
Dim NAeroType As Long
Dim NThermoType As Long
Dim numGroup As Long
Dim NumTrack As Long
Dim i As Integer
Dim numStops As Long
Dim numSecs As Long
Dim RowF11B As Long
Dim NumPrint As Long
Dim LastRow As Long
Dim ipversion As Boolean

Dim segtype As Integer

'Locations of information on the control sheet
Public Write_Options As Range
Public SES_Exe As Range
Public NextOut_Exe As Range
Public Visio_File As Range

Sub Get_Control_Values(wname)
    Set SES_Exe = Workbooks(wname).Worksheets("Control").Range("H14")
    Set NextOut_Exe = Workbooks(wname).Worksheets("Control").Range("H15")
    Set Write_Options = Workbooks(wname).Worksheets("Control").Range("C14")
    Set Visio_File = Workbooks(wname).Worksheets("Control").Range("H17")
End Sub

Public Sub WriteFile(Optional unit_name As String) 'Copy data from Form Worksheets to Output Worksheet
    'unit_name is used for unit_tests. Otherwise, the value should be empty
    On Error GoTo ErrorProc
    wname = ActiveWorkbook.Name
    Get_Control_Values (wname) 'Get settings from Control worksheet
    Dim num_sections, num_vents, num_line_sec As Integer 'Variables for Form 2
    Dim StartTime, Ftime, EndTime As Double
    StartTime = Timer
    WriteForm.Show vbModeless
    WriteForm.TextBox2.value = "Formulas are Calculating"
    WriteForm.Repaint
    'Start program
    Call Speedon(True)                           'Makes the program work faster.
    Dim TC, LC, RC, LastRow, numlines, rF2 As Integer
    Dim FormRange As Range
    TC = 5                                       'Top Cell with Data on most sheets
    LC = 2                                       'Most left cell with data
    Calculate                                    'make sure to recalculate any formulas
    ipversion = is_version_ip(wname)
    Workbooks(wname).Worksheets("Control").Range("H21").Value2 = Workbooks(wname).BuiltinDocumentProperties("Last Author")
    With Workbooks(wname)
        Set Output = Workbooks(wname).Worksheets("Output")
        WriteForm.TextBox2.value = "Producing Form 1"
        WriteForm.Repaint
        If .Worksheets("F01").Range("D3") = "" Then
            WriteForm.Hide
            Speedon (False)
            MsgBox ("Add System Identification to Form 1A to write the file")
            Exit Sub
        End If
        'Clear contents of Output Sheet
        Output.Cells.Delete Shift:=xlUp
        NextRow = 1
        ' Form 1 Workshop starting point
        Set FormIn = .Worksheets("F01")
        Row = 3
        col = 4
        ' Form 1A Title
        NumCol = 1
        'LRtext = FormIn.Range("D3:D21").End(xlDown).Row 'Last Row of titles with text'
        LRtext = FormIn.Range("D22").End(xlUp).Row 'Last Row of titles with text'
        If LRtext < 23 Then                      'Calculate number of lines to move down
            numlines = LRtext - Row + 1
        Else
            numlines = 1
            LRtext = 3
        End If
        'Force format of Form 1A cells to be text. This prevents converting Dates to numerical values
        'Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow + numlines - 1, 1)).NumberFormat = "@"
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow + numlines - 1, 1)).Value2 = _
                                                                                               FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(LRtext, col)).Value2
        NextRow = NextRow + numlines
        'Form 1B
        Row = 23
        NumData = 3
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumData)).Value2 = _
                                                                                      Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        NextRow = NextRow + 1
        Row = Row + NumData
        'Form 1C
        NumData = 8
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumData)).Value2 = _
                                                                                      Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        TPO = FormIn.Cells(Row, col).Value2      'Train performance option
        ECL = FormIn.Cells(Row + 3, col).Value2  'Environemtnal Control Load
        NextRow = NextRow + 1
        Row = Row + NumData
        'Form 1D
        NumData = 7
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumData)).Value2 = _
                                                                                      Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        NumUnsteady = FormIn.Cells(Row + 5, col).Value2
        If ipversion Then                        'adjust output sheet if using ipVersion
            Output.Cells(NextRow, 8).Value2 = Output.Cells(NextRow, 7).Value2
            Output.Cells(NextRow, 7).Value2 = Output.Cells(NextRow, 6).Value2
            Output.Cells(NextRow, 6).Value2 = 0
        End If
        NextRow = NextRow + 1                    'Move to next line for Form 1E
        Row = Row + NumData                      'Move to next row on F01
        'Form 1E
        NumData = 8
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumData)).Value2 = _
                                                                                      Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        NumECZ = FormIn.Cells(Row + 2, col).Value2
        InitialReadOption = FormIn.Cells(Row + 7, col).Value2
        If InitialReadOption <> 0 Then
            InitialFileName = FormIn.Cells(Row + 7, col + 3).Value2
        End If
        NextRow = NextRow + 1
        Row = Row + NumData
        'Form 1F
        NumData = 8
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumData)).Value2 = _
                                                                                      Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        NextRow = NextRow + 1
        Row = Row + NumData
        'Form 1G
        NumData = 8
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumData)).Value2 = _
                                                                                      Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        NextRow = NextRow + 1
        Row = Row + NumData
        'Form 1H
        If Not ipversion Then
            NumData = 5
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumData)).Value2 = _
                                                                                          Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
            NextRow = NextRow + 1
            CPipe = FormIn.Cells(Row + 4, col).Value2
            Row = Row + NumData
        End If
        ' Form 2A
        WriteForm.TextBox2.value = "Writing Form 2A"
        WriteForm.Repaint
        num_sections = .Worksheets("F01").Range("D35").Value2
        num_vents = .Worksheets("F01").Range("D36").Value2
        num_line_sec = num_sections - num_vents  'Number of Line Sections for tunnels
        rF2 = 5                                  'Starting row for Form 2A and 2B
        RC = 5                                   'Number of columns of data
        Set FormIn = .Worksheets("F02A")
        With FormIn
            Set FormRange = .Range(.Cells(rF2, LC), .Cells(rF2 + num_line_sec - 1, RC + LC - 1))
        End With
        With Output
            .Range(.Cells(NextRow, 1), .Cells(NextRow + num_line_sec - 1, RC)).Value2 = FormRange.Value2
        End With
        NextRow = NextRow + num_line_sec
        ' Form 2B
        If num_vents > 0 Then
            WriteForm.TextBox2.value = "Writing Form 2B"
            WriteForm.Repaint
            Set FormIn = .Worksheets("F02B")
            RC = 5                               'Number of columns of data
            With FormIn
                Set FormRange = .Range(.Cells(rF2, LC), .Cells(rF2 + num_vents - 1, RC + LC - 1))
            End With
            With Output
                Output.Range(.Cells(NextRow, 1), .Cells(NextRow + num_vents - 1, RC)).Value2 = FormRange.Value2
            End With
            NextRow = NextRow + num_vents
        End If
        ' Form 3
        WriteForm.TextBox2.value = "Producing Form 3"
        WriteForm.Repaint
        Set FormIn = .Worksheets("F03")
        'Set FormRange = FormIn.Range("A1:AX1500")
        Row = 5                                  'Starting Row
        col = 2                                  'Starting Column
        Do While FormIn.Cells(Row, col) <> ""
            'start of loop
            'Form 3A line
            NumCol = 3
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            'Form 3A line 2
            col = col + NumCol                   'Starting point for column
            NumCol = 5                           'Number of columns in this line of the form
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            'Form 3B line 1
            col = col + NumCol
            NumCol = 8
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            'Form 3B line 2
            col = col + NumCol
            NumCol = 8
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            'Form 3C
            col = col + NumCol
            NumCol = 7
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            NumofSub = FormIn.Cells(Row, col + 5).Value2 'Number of subsegments
            NumofSSHS = FormIn.Cells(Row, col + 6).Value2 'Number of Steady-State heat sources
            'Form 3D
            col = col + NumCol
            counter = 1
            NumCol = 6
            Do While ((counter < NumofSSHS) Or (counter = NumofSSHS)) ' Inner loop.
                Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                             FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
                NextRow = NextRow + 1            'Move one cell down for Output Sheet
                If counter > 99 Then
                    MsgBox "Error in Form 3D, Ending Sub-Segment Number"
                    End
                End If
                Row = Row + 1
                counter = counter + 1
            Loop
            Row = Row - counter + 1              'Resets Row back to start
            'Form 3E
            col = col + NumCol
            NumCol = 5
            EndSub = 0
            counter = 0
            Do While (EndSub < NumofSub)         ' Inner loop.
                EndSub = FormIn.Cells(Row, col + 1).Value2
                Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                             FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
                If EndSub < NumofSub Then Row = Row + 1
                NextRow = NextRow + 1
                counter = counter + 1
                If counter > 99 Then
                    MsgBox "Error in Form 3E, Ending Sub-Segment Number"
                    End
                End If
            Loop
            Row = Row - counter + 1
            'Form 3F
            col = col + NumCol
            NumCol = 7
            If ECL = 1 Or ECL = 2 Then
                Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                             FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
                NextRow = NextRow + 1
            End If
            'Restart at Form 3A, First figure out line to start
            R_off = 1                            'Need to figure out the starting row
            If R_off < counter Then R_off = counter
            If R_off < NumofSSHS Then R_off = NumofSSHS
            col = 2
            Row = Row + R_off
        Loop
        ' Form 4
        WriteForm.TextBox2.value = "Producing Form 4"
        WriteForm.Repaint
        Set FormIn = .Worksheets("F04")
        'Set FormRange = FormIn.Range("A1:AX1500")
        Row = 5                                  'Starting Row
        If NumUnsteady > 0 Then
            Do While FormIn.Cells(Row, col) <> ""
                NextCol = 1
                col = 2                          'Starting Column
                'Form 4 Line 1, Write Source Name
                Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, 1)).Value2 = _
                                                                                        FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col)).Value2
                'Form 4 Line 1, Write Location and Subsegement
                col = col + 1
                NextCol = 5
                NumCol = 2
                Output.Range(Output.Cells(NextRow, NextCol), Output.Cells(NextRow, NextCol + NumCol - 1)).Value2 = _
                                                                                                                 FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
                'Form 4 Line 2
                NextRow = NextRow + 1
                col = col + NumCol
                NumCol = 6
                Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                             FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
                NextRow = NextRow + 1
                Row = Row + 1
            Loop
        End If
        ' Form 5
        WriteForm.TextBox2.value = "Producing Form 5"
        WriteForm.Repaint
        Set FormIn = .Worksheets("F05")
        Row = 5                                  'Starting Row
        Do While FormIn.Cells(Row, col) <> ""
            'start of loop
            'Form 5A
            col = 2
            NumCol = 3
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            segtype = FormIn.Cells(Row, 3)
            'If FormIn.Cells(Row, 3) > 2 Then MsgBox "This spreadsheet does not support Vent shaft types other than 1 and 2. Updates coming."
            'Form 5B
            col = col + NumCol                   'Starting point for column
            NumCol = 8                           'Number of columns in this line of the form
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NumofSub = FormIn.Cells(Row, 5).Value2
            NextRow = NextRow + 1
            'Form 5C
            If segtype <> 3 Then                 'If segement type is not type 3
                col = col + NumCol
            Else: col = col + NumCol + 4
            End If
            NumCol = 4
            If .Worksheets("F01").Range("D40").Value2 > 0 Then
                Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                             FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
                NextRow = NextRow + 1
            End If
            'Form 5D
            If segtype <> 3 Then
                col = col + NumCol + 4
            Else: col = col + NumCol
            End If
            NumCol = 7
            counter = 1
            Do While ((counter < NumofSub) Or (counter = NumofSub)) ' Inner loop.
                Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                             FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
                NextRow = NextRow + 1            'Move one cell down for Output Sheet
                If counter > 99 Then
                    MsgBox "Error in Form 3D, Ending Sub-Segment Number"
                    End
                End If
                Row = Row + 1
                counter = counter + 1
            Loop
            'Restart at Form 5A, First figure out line to start
            If counter = 1 Then Row = Row + 1
        Loop
        'Form 6
        WriteForm.TextBox2.value = "Producing Form 6"
        WriteForm.Repaint
        Set FormIn = .Worksheets("F06")
        col = 2
        Row = 5                                  'Starting Row
        Do While FormIn.Cells(Row, col) <> ""
            'start of loop
            'Form 6A
            NumCol = 3
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NAeroType = FormIn.Cells(Row, col + 1).Value2
            NThermoType = FormIn.Cells(Row, col + 2).Value2
            NextRow = NextRow + 1
            'Form 6B Boundary Conditions (if Applicable)
            If NThermoType = 3 Then
                col = 5
                NumCol = 6
                Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                             FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
                NextRow = NextRow + 1
            End If
            Select Case NAeroType
            Case 0, 7
                NumCol = -1                      'This will skip writing the line
            Case 1                               'tunnel to tunnel crossover juction
                col = 11                         'Row K value corrected in 0p09
                NumCol = 5                       'Five data points
            Case 2                               'dividing wall termination junction
                col = 16                         'Row p
                NumCol = 3                       'Five data points
            Case 3                               '"T" Junction
                col = 19                         'S
                NumCol = 4
            Case 4                               'angled junction
                col = 23                         'w
                NumCol = 5
            Case 5                               '"Y"Junction
                col = 28
                NumCol = 5
            Case 6                               'Saccardo Nozzle
                col = 40                         'Column AN
                NumCol = 6
            Case 7                               'zero total pressure change jucntion
                NumCol = -1
            Case 8                               'PSD Junction
                col = 46                         'AT
                NumCol = 6
            Case Else                            ' Other values.
                MsgBox "Error in Form 6, Node Type does not exist"
            End Select
            If NumCol > 0 Then
                Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                             FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
                NextRow = NextRow + 1
            End If
            col = 2
            Row = Row + 1
        Loop                                     'Loop for Nodes on Form 6
        'Form 7A and B
        WriteForm.TextBox2.value = "Producing Form 7A and 7B"
        WriteForm.Repaint
        Set FormIn = .Worksheets("F07")
        col = 2
        Row = 5                                  'Starting Row
        Do While FormIn.Cells(Row, col) <> ""
            'Form 7A Line 1, Write Source Name
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, 1)).Value2 = _
                                                                                    FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col)).Value2
            'Form 7A Line 1, Write Location and Subsegement
            col = col + 1
            NextCol = 5
            NumCol = 4
            Output.Range(Output.Cells(NextRow, NextCol), Output.Cells(NextRow, NextCol + NumCol - 1)).Value2 = _
                                                                                                             FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            'Form 7B Line 1
            col = col + NumCol
            NumCol = 8
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            'Form 7B Line 2
            col = col + NumCol
            NumCol = 8
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            Row = Row + 1                        'Move one row down
            col = 2                              'Reset Column
        Loop
        'Form 7C
        WriteForm.TextBox2.value = "Producing Form 7C"
        WriteForm.Repaint
        Set FormIn = .Worksheets("F07C")
        col = 2
        Row = 5                                  'Starting Row
        Do While FormIn.Cells(Row, col) <> ""
            'Form 7C Line 1
            NumCol = 7
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            Row = Row + 1                        'Move one row down
            col = 2                              'Reset Column
        Loop
        'Form 7D needs to be inputted here
        WriteForm.TextBox2.value = "Producing Form 7D"
        WriteForm.Repaint
        Set FormIn = .Worksheets("F07D")
        col = 2
        Row = 5                                  'Starting Row
        Do While FormIn.Cells(Row, col) <> ""
            'Form 7D Line 1
            NumCol = 8
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            Row = Row + 1                        'Move one row down
        Loop
        'Form 8
        If TPO > 0 Then
            WriteForm.TextBox2.value = "Producing Form 8"
            WriteForm.Repaint
            Call WriteForm8
        End If                                   'If statement to skip Form 9 if train performance option is 0
        'Form 9
        If TPO > 0 Then
            WriteForm.TextBox2.value = "Producing Form 9"
            WriteForm.Repaint
            Call WriteForm9(NextRow)
        End If                                   'If statement to skip Form 9 if train performance option is 0
        'Form 10
        WriteForm.TextBox2.value = "Producing Form 10"
        WriteForm.Repaint
        Set FormIn = .Worksheets("F10")
        col = 2
        Row = 5                                  'Starting Row
        NumCol = 8
        Do While FormIn.Cells(Row, col) <> ""
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            Row = Row + 1
        Loop
        'Form 11
        WriteForm.TextBox2.value = "Producing Form 11"
        WriteForm.Repaint
        Call WriteForm11
        'Form 12
        Set FormIn = .Worksheets("F12")
        WriteForm.TextBox2.value = "Producing Form 12"
        WriteForm.Repaint
        NumData = 2
        Row = 2
        col = 4
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumData)).Value2 = _
                                                                                      Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData - 1, col)).Value2)
        'Form 12 Second and beyond entries
        NumPrint = FormIn.Cells(3, 4).Value2
        If NumPrint = 0 Then NumPrint = 1 'This forces blank values if Form 12 is empty.
        NextRow = NextRow + 1
        col = 2
        Row = 5
        NumCol = 7
        With FormIn
            Set FormRange = .Range(.Cells(Row, col), .Cells(Row + NumPrint - 1, col + NumCol - 1))
        End With
        With Output
            .Range(.Cells(NextRow, 1), .Cells(NextRow + NumPrint - 1, NumCol)).Value2 = FormRange.Value2
        End With
        NextRow = NextRow + NumPrint
        'Form 13
        WriteForm.TextBox2.value = "Producing Form 13"
        WriteForm.Repaint
        Set FormIn = .Worksheets("F13")
        col = 2
        Row = 4                                  'Starting Row of Form
        NumCol = 4
        Do While FormIn.Cells(Row, col) <> ""
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                         FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            NextRow = NextRow + 1
            Row = Row + 1                        'move a row down on Form 11A
        Loop
        'Form 14
        If CPipe > 0 Then
            WriteForm.TextBox2.value = "Producing Form 14"
            WriteForm.Repaint
            Call WriteForm14
        End If
        'Restart File
        If InitialReadOption <> 0 Then
            Output.Cells(NextRow, 1).Value2 = InitialFileName
        End If
    End With
    Call FormatINP                               'Adjust width and alignment of text
    WriteForm.TextBox2.value = "Formatting Numbers"
    WriteForm.Repaint
    Call FormatNumbersArray                      'Adjust numerical format
    EndTime = Timer
    Debug.Print unit_name & " Time after formating before save: " & (EndTime - StartTime)
    WriteForm.TextBox2.value = "Exporting File"
    Call Speedon(False)                          'Enable items that previously slowed down processing.
    If unit_name = "" Then
        Call WriteINP                            'Write out file to text file
    Else
        Call WriteINP(unit_name)
    End If
    WriteForm.Hide
    Exit Sub
ErrorProc:
    Speedon (False)
    MsgBox "Error in procedure Write File : " & Err.Description
    Err.Clear
End Sub

Sub InitializeRanges()
    Set ControlSheetRange = ThisWorkbook.Worksheets("Control").Range("I10")
    Set AnotherRange = ThisWorkbook.Worksheets("Control").Range("J20")
End Sub

Private Sub WriteINP(Optional unit_name As String)
    On Error GoTo ErrorProc
    Dim file_selected As Variant
    Dim savename, write_info As String
    Dim write_date, write_time As Variant
    Dim open_save_as_dialog, save_file As Boolean
    Dim overwrite_exiting_file As VbMsgBoxResult
    savename = ""
    If unit_name = "" Then
        open_save_as_dialog = True 'Open the save as dialog box
    Else
        savename = unit_name
        open_save_as_dialog = False 'skip the save as dialog box
        save_file = True 'save the file!
    End If
    While open_save_as_dialog
        file_selected = Application.GetSaveAsFilename(fileFilter:="SES Input File (*.inp), *.inp", Title:="Save SES Input File")
        If file_selected = False Then
            open_save_as_dialog = False
            save_file = False
        Else:
            savename = CStr(file_selected)
            If Dir(savename) <> "" Then 'Test if file already exists
                overwrite_exiting_file = MsgBox("The file already exists. Do you want to overwrite it?", vbYesNoCancel + vbExclamation, "File Exists")
                    Select Case overwrite_exiting_file
                        Case vbYes ' Overwrite the file
                            save_file = True
                            open_save_as_dialog = False
                        Case vbNo ' Ask user for a new file name
                            save_file = False
                            open_save_as_dialog = True
                        Case vbCancel
                            save_file = False
                            open_save_as_dialog = False
                    End Select
            End If
        End If
    Wend
    If save_file Then
        Application.DisplayAlerts = False 'False to suppress alert message
        ThisWorkbook.Sheets("Output").Copy
        ActiveWorkbook.SaveAs FileName:=savename, FileFormat:=xlTextPrinter, CreateBackup:=False
        ActiveWorkbook.Close
        Application.DisplayAlerts = True 'Renable alert messages
        'Add information about the file that was written to the control sheet
        write_date = Date
        write_time = Time
        write_info = "Last Wrote on " & write_date & " at " & write_time & ":"
        Workbooks(wname).Worksheets("Control").Range("B20").Value2 = write_info
        Workbooks(wname).Worksheets("Control").Range("H20").Value2 = savename 'Change sheet to say last saved
        Workbooks(wname).Worksheets("Control").Range("H21").Value2 = Workbooks(wname).BuiltinDocumentProperties("Last Author")
        If ipversion Then
            Workbooks(wname).Worksheets("Control").Range("G20").Value2 = "(SES 4.1)"
        Else
            Workbooks(wname).Worksheets("Control").Range("G20").Value2 = "(SES 6.0)"
        End If
        If Workbooks(wname).Worksheets("Control").Range(Write_Options.Address).Value2 = 2 Then
            WriteForm.TextBox2.value = "Running SES Simulation"
            WriteForm.Repaint
            Call_SES_Exe wname, savename
        ElseIf Workbooks(wname).Worksheets("Control").Range(Write_Options.Address).Value2 = 3 Then
            WriteForm.TextBox2.value = "Running SES and Next-Out"
            WriteForm.Repaint
            Call_NextOut wname, savename
        End If
    End If
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure WriteINP: " & Err.Description
    Err.Clear
End Sub

Public Sub choose_ses_exe(wname)
    On Error GoTo ErrorProc
    'Declare a variable as a FileDialog object.
    Dim FD As FileDialog
    'Create a FileDialog object as a File Picker dialog box.
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    'Declare a variable to contain the path
    'of each selected item. Even though the path is a String,
    'the variable must be a Variant because For Each...Next
    'routines only work with Variants and Objects.
    Dim vrtSelectedItem As Variant
    'Use a With...End With block to reference the FileDialog object.
    Get_Control_Values (wname)
    With FD
        'Use the Show method to display the File Picker dialog box and return the user's action.
        'The user pressed the action button.
        If .InitialFileName = "" Then .InitialFileName = ActiveWorkbook.Path
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Executable Files", "*.EXE", 1
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                'vrtSelectedItem is a String that contains the path of each selected item.
                'You can use any file I/O functions that you want to work with this path.
                'This example simply displays the path in a message box.
                'MsgBox "The path is: " & vrtSelectedItem
                Infile = vrtSelectedItem
            Next vrtSelectedItem
        Else: Infile = ""
        End If
    End With
    Workbooks(wname).Worksheets("Control").Range(SES_Exe.Address).Value2 = Infile
    'Set the object variable to Nothing.
    Set FD = Nothing
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure choose_ses_exe : " & Err.Description
    Err.Clear
End Sub

Public Sub choose_exe(wname, program_name As String)
    On Error GoTo ErrorProc
    'Declare a variable as a FileDialog object.
    Dim FD As FileDialog
    'Create a FileDialog object as a File Picker dialog box.
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    'Declare a variable to contain the path
    'of each selected item. Even though the path is a String,
    'the variable must be a Variant because For Each...Next
    'routines only work with Variants and Objects.
    Dim vrtSelectedItem As Variant
    'Use a With...End With block to reference the FileDialog object.
    Get_Control_Values (wname)
    With FD
        'Use the Show method to display the File Picker dialog box and return the user's action.
        'The user pressed the action button.
        If .InitialFileName = "" Then .InitialFileName = ActiveWorkbook.Path
        .AllowMultiSelect = False
        .Filters.Clear
        If program_name = "SES" Or program_name = "NextOut" Then
            .Filters.Add "Executable Files", "*.EXE", 1
        ElseIf program_name = "Visio" Then
            .Filters.Add "Visio Files", "*.vsdx", 1
        Else
            .Filters.Add "All Files", "*.E", 1
        End If
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                'vrtSelectedItem is a String that contains the path of each selected item.
                'You can use any file I/O functions that you want to work with this path.
                'This example simply displays the path in a message box.
                'MsgBox "The path is: " & vrtSelectedItem
                Infile = vrtSelectedItem
            Next vrtSelectedItem
        Else: Infile = ""
        End If
    End With
    If program_name = "SES" Then
        Workbooks(wname).Worksheets("Control").Range(SES_Exe.Address).Value2 = Infile
    ElseIf program_name = "NextOut" Then
        Workbooks(wname).Worksheets("Control").Range(NextOut_Exe.Address).Value2 = Infile
    ElseIf program_name = "Visio" Then
        Workbooks(wname).Worksheets("Control").Range(Visio_File.Address).Value2 = Infile
    End If
    'Set the object variable to Nothing.
    Set FD = Nothing
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure choose_ses_exe : " & Err.Description
    Err.Clear
End Sub

Private Sub FormatINP()
    'Left justify and create the write width
    On Error GoTo ErrorProc
    Workbooks(wname).Worksheets("Output").Columns("A:H").HorizontalAlignment = xlLeft
    Workbooks(wname).Worksheets("Output").Columns("A:H").ColumnWidth = 10.29
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure FormatINP: " & Err.Description
    Err.Clear
End Sub

Private Sub FormatNumbersArray()
    'Format numbers
    On Error GoTo ErrorProc
    Dim OutputArray() As Variant
    Dim OutputRange As Object
    Dim J As Integer
    With Workbooks(wname).Worksheets("Output")
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        Set OutputRange = .Range(.Cells(1, 1), .Cells(LastRow, 8))
        OutputArray = OutputRange.Value2
        OutputRange.NumberFormat = "0.#######"
        For i = LBound(OutputArray, 1) To UBound(OutputArray, 1)
            For J = LBound(OutputArray, 2) To UBound(OutputArray, 2)
                If Not IsEmpty(OutputArray(i, J)) And IsNumeric(OutputArray(i, J)) Then
                    If (Abs(OutputArray(i, J)) > 99999999) Or ((Abs(OutputArray(i, J)) < 0.00001) And (OutputArray(i, J) <> 0)) Then
                        'If value is greater than 8 places or smaller then 5 decimal places
                        OutputRange(i, J).NumberFormat = "0.000E+00"
                    ElseIf Len(OutputRange(i, J)) > 9 Then 'If the value is longer than 9 places
                        Select Case OutputArray(i, J)
                        Case Is > 10000000
                            OutputRange(i, J).NumberFormat = "0."
                        Case Is > 1000000
                            OutputRange(i, J).NumberFormat = "0.#"
                        Case Is > 100000
                            OutputRange(i, J).NumberFormat = "0.##"
                        Case Is > 10000
                            OutputRange(i, J).NumberFormat = "0.###"
                        Case Is > 1000
                            OutputRange(i, J).NumberFormat = "0.####"
                        Case Is > 100
                            OutputRange(i, J).NumberFormat = "0.#####"
                        Case Is > 10
                            OutputRange(i, J).NumberFormat = "0.######"
                        Case Is > 1
                            OutputRange(i, J).NumberFormat = "0.#######"
                        End Select
                    End If
                End If
            Next J
        Next i
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure FormatNumberArray : " & Err.Description
    Err.Clear
End Sub

Public Sub Speedon(ByVal SetOn As Boolean)
    'Speeds up processing by turning off some functionality
    'Sets the application to use Decimal Seperartor as a period.
    On Error GoTo ErrorProc
    Dim temp As Integer
    With Application
        If SetOn Then
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayAlerts = False
            .Cursor = xlWait
            .DisplayStatusBar = True
            .StatusBar = "Next-In Working..."
        Else
            .Calculation = xlCalculationAutomatic
            .ScreenUpdating = True
            .EnableEvents = True
            .DisplayAlerts = True
            .Cursor = xlDefault
            .StatusBar = False
        End If
    End With
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Speedon, hit Reset : " & Err.Description
    Err.Clear
End Sub

Private Sub WriteForm9(OutRow)
    On Error GoTo ErrorProc
    col = 6
    Row = 3                                      'Starting Row
    Set FormIn = Workbooks(wname).Worksheets("F09")
    Do While FormIn.Cells(Row, col) <> ""
        'Test code
        WriteForm.TextBox2.value = "Producing Form 9"
        WriteForm.Repaint
        'Form 9A Line 1, Write Source Name
        Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, 1)).Value2 = _
                                                                              FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col)).Value2
        'Form 9A Line 1, Write Location and Subsegement
        Row = Row + 1
        NextCol = 5
        NumData = 4
        Output.Range(Output.Cells(OutRow, NextCol), Output.Cells(OutRow, NextCol + NumData - 1)).Value2 = _
                                                                                                        Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData - 1, col)).Value2)
        OutRow = OutRow + 1
        'Form 9B
        Row = Row + NumData
        NumData = 5
        Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                    Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        OutRow = OutRow + 1
        'Form 9C
        Row = Row + NumData
        NumData = 6
        Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                    Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        OutRow = OutRow + 1
        'Form 9D Line1
        Row = Row + NumData
        NumData = 8
        Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                    Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        OutRow = OutRow + 1
        'Form 9D Line2
        Row = Row + NumData
        NumData = 8
        Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                    Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        OutRow = OutRow + 1
        'Form 9E
        Row = Row + NumData
        NumData = 6
        Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                    Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
        OutRow = OutRow + 1
        If TPO <> 3 Then
            'Form 9F Line 1, Write Source Name
            Row = Row + NumData
            NumData = 1
            Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, 1)).Value2 = _
                                                                                  Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col)).Value2)
            'Form 9F Line 1, Write Location and Subsegement
            Row = Row + 1
            NextCol = 5
            NumData = 2
            Output.Range(Output.Cells(OutRow, NextCol), Output.Cells(OutRow, NextCol + NumData - 1)).Value2 = _
                                                                                                            Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
            OutRow = OutRow + 1
            'Form 9F Line 2, Write Location and Subsegement
            Row = Row + NumData
            NumData = 5
            Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                        Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
            OutRow = OutRow + 1
            'Form 9G Line 1
            Row = Row + NumData
            NumData = 4
            Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                        Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
            OutRow = OutRow + 1
            'Form 9G Line 2
            Row = Row + NumData
            NumData = 4
            Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                        Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
            OutRow = OutRow + 1
            'Form 9G Line 3
            Row = Row + NumData
            NumData = 4
            Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                        Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
            OutRow = OutRow + 1
            'Form 9G Line 3
            Row = Row + NumData
            NumData = 1
            Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                        Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
            OutRow = OutRow + 1
            TCO = Int_International(FormIn.Cells(Row, col).Value2) 'Train Control Option
            If TCO = 2 Then
                'Form 9H Line 1 Train Controller
                Row = Row + NumData
                NumData = 5
                Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                            Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
                OutRow = OutRow + 1
                'Form 9H Line 2
                Row = Row + NumData
                NumData = 5
                Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                            Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
                OutRow = OutRow + 1
                OFSO = FormIn.Cells(Row + 4, col) 'onboard Flyewheel simulation option
            Else
                Row = Row + 5
                OFSO = 1                         'set onboard flywhell simulation option to bypass
            End If
            If TCO = 3 Then
                'Form 9H-A Line 1
                Row = 72
                NumData = 5
                Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                            Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
                OutRow = OutRow + 1
                'Form 9H-A Line 2
                Row = Row + NumData
                NumData = 5
                Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                            Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
                OutRow = OutRow + 1
                'Form 9H-A Line 3
                Row = Row + NumData
                NumData = 5
                Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                            Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
                OutRow = OutRow + 1
            End If
            'Form 9I Line 1
            Row = 87
            NumData = 5
            Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                        Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
            OutRow = OutRow + 1
            If TPO = 1 Then
                'Form 9J Line 1
                Row = Row + NumData
                NumData = 5
                Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                            Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
                OutRow = OutRow + 1
            Else: Row = Row + 5
            End If
            'Form 9K Line 1
            If OFSO = 2 Then
                Row = Row + NumData
                NumData = 5
                Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                            Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
                OutRow = OutRow + 1
                'Form 9L Line 1
                Row = Row + NumData
                NumData = 7
                Output.Range(Output.Cells(OutRow, 1), Output.Cells(OutRow, NumData)).Value2 = _
                                                                                            Application.Transpose(FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row + NumData, col)).Value2)
                OutRow = OutRow + 1
            End If
        End If
        'Reset to read next column
        col = col + 1                            'Move one row over
        Row = 3                                  'Reset Column
    Loop
    NextRow = OutRow
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure WriteForm9 : " & Err.Description
    Err.Clear
End Sub

Private Sub WriteForm14()
    On Error GoTo ErrorProc
    Dim i As Integer
    Dim RowF14C As Integer
    Dim ColF14C As Integer
    Dim NumInlet As Integer
    Dim numSec As Integer
    Dim SNum As Integer
  
    'NumPipes is the number of cool pies to read in
    Set FormIn = Workbooks(wname).Worksheets("F14AB")
    Set F14C = Workbooks(wname).Worksheets("F14C")
    Set Output = Workbooks(wname).Worksheets("Output")
    Row = 4                                      'Starting Row 14AB
    col = 2                                      'Starting Row 14AB
    RowF14C = 4                                  'Starting Row 14C
    ColF14C = 2                                  'Starting Row 14C
    Do While FormIn.Cells(Row, col) <> ""
        col = 2                                  'Starting Column
        'Form 14A Line 1, Write Cooling Pipe Description
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, 1)).Value2 = _
                                                                                FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col)).Value2
        NextRow = NextRow + 1
        'Form 14A Line 2, Write Location and Subsegement
        col = col + 1
        NextCol = 1
        NumCol = 7
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NextCol + NumCol - 1)).Value2 = _
                                                                                                   FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
        NumInlet = FormIn.Cells(Row, 3).Value2
        numSec = FormIn.Cells(Row, 4).Value2
        NextRow = NextRow + 1
        'Form 14B Line3 and all inlet IDs
        For i = 1 To NumInlet
            NextCol = 1
            NumCol = 3
            col = 10
            Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NextCol + NumCol - 1)).Value2 = _
                                                                                                       FormIn.Range(FormIn.Cells(Row, col), FormIn.Cells(Row, col + NumCol - 1)).Value2
            Row = Row + 1                        'Advance down a row on Form 14 AB
            NextRow = NextRow + 1
        Next i
        'Form 14C which needs to include + or - sign
        Dim NumArray() As Variant
        numSec = 1
        If numSec > 1 Then
            NumArray = F14C.Range(F14C.Cells(RowF14C, ColF14C), F14C.Cells(RowF14C + numSec - 1, ColF14C)).Value2
            'ActiveCell.Value2 = F14C.Range(F14C.Cells(RowF14C, ColF14C), F14C.Cells(RowF14C + NumSec - 1, ColF14C)).Value2
            'NumArray() = Range("B4:B5").Value2
            'Create String Array with + or -. Values are entered as strings faster than entering the value for each cell
            For i = LBound(NumArray) To UBound(NumArray)
                If NumArray(i, 1) > 0 Then
                    Output.Cells(NextRow, 1).Value2 = "'+" & CStr(NumArray(i, 1)) & "."
                Else
                    Output.Cells(NextRow, 1).Value2 = "'" & CStr(NumArray(i, 1)) & "."
                End If
                NextRow = NextRow + 1
            Next i
            Erase NumArray
        ElseIf numSec = 1 Then
            SNum = Int_International(F14C.Cells(RowF14C, ColF14C).Value2)
            If SNum > 0 Then
                Output.Cells(NextRow, 1).Value2 = "'+" & SNum & "."
            Else
                Output.Cells(NextRow, 1).Value2 = "'" & SNum & "."
            End If
            NextRow = NextRow + 1
        End If
        ColF14C = ColF14C + 1
        'remember to delete arrays
    Loop
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure WriteForm14 : " & Err.Description
    Err.Clear
End Sub

Private Sub WriteForm8()

    Dim Line8A As Long
    Dim Line8B As Long
    Dim Line8C As Long
    Dim Line8D As Long
    Dim Line8E As Long
    Dim Line8F As Long
    Dim Col8A As Long
    Dim Col8B As Long
    Dim Col8C As Long
    Dim Col8D As Long
    Dim Col8E As Long
    Dim Col8F As Long
    Dim numlines As Long
    Line8A = 5                                   'Starting line for Line 8A
    Line8B = 4                                   'Starting line for Line 8B
    Line8C = 5                                   'Starting line for Form 8C
    Line8D = 7                                   'Starting line for Form 8D
    Line8E = 6                                   'starting line for Form 8E
    Line8F = 5                                   'Starting line for Form 8F
    Col8A = 2                                    'starting column for Form 8A
    Col8B = 2                                    'starting column for Form 8B
    Col8C = 2                                    'starting column for Form 8C
    Col8D = 2                                    'starting column for Form 8D
    Col8E = 2                                    'starting column for Form 8E
    Col8F = 4                                    'starting column for form 8F
    Do While (Workbooks(wname).Worksheets("F08A").Cells(Line8A, 2).Value2 <> "")
        'Form 8A
        Set FormIn = Workbooks(wname).Worksheets("F08A")
        col = Col8A
        NumData = 1
        Call Range2DSame(NextRow, 1, Line8A, col, 1, NumData, Output, FormIn)
        col = col + NumData
        NumData = 7
        Call Range2DSame(NextRow, 1, Line8A, col, 1, NumData, Output, FormIn)
        Range2D(NextRow, 1, NumData, 1, Output) = Range2D(Line8A, col, NumData, 1, FormIn)
        numGroup = FormIn.Cells(Line8A, 4).Value2 'Number of train routes
        NumTrack = FormIn.Cells(Line8A, 5).Value2 'Number of train routes
        Line8A = Line8A + 1
        'Form 8B
        Set FormIn = Workbooks(wname).Worksheets("F08B")
        NumData = 3
        If numGroup > 1 Then
            Call Range2DSame(NextRow, 1, Line8B, Col8B, numGroup - 1, 3, Output, FormIn)
        End If
        Col8B = Col8B + 3
        'Form 8C
        If TPO = 1 Or TPO = 2 Then
            Set FormIn = Workbooks(wname).Worksheets("F08C")
            NumData = 8
            Call Range2DSame(NextRow, 1, Line8C, Col8C, NumTrack, NumData, Output, FormIn)
            Col8C = Col8C + 8
        End If
        'Form 8D
        If TPO = 1 Then
            Set FormIn = Workbooks(wname).Worksheets("F08D")
            'Form 8D first entry
            NumData = 2
            Call Vertical2Horizontal(NextRow, 1, Line8D - 5, Col8D + 2, 2, Output, FormIn)
            numStops = FormIn.Cells(Line8D - 5, Col8D + 2).Value2
            'Form 8D Second and beyond entries
            NumData = 3
            If numStops > 0 Then
                Call Range2DSame(NextRow, 1, Line8D, Col8D, numStops, NumData, Output, FormIn)
            End If
            Col8D = Col8D + 3
        End If
        'Form 8E
        If TPO = 2 Or TPO = 3 Then
            Set FormIn = Workbooks(wname).Worksheets("F08E")
            numlines = FormIn.Cells(Line8E - 4, Col8E + 4).Value2 'number of spee-time points
            Output.Cells(NextRow, 1).Value2 = numlines
            NextRow = NextRow + 1
            NumData = 5
            Call Range2DSame(NextRow, 1, Line8E, Col8E, numlines, NumData, Output, FormIn)
            Col8E = Col8E + 5
        End If
        'Form 8F
        Set FormIn = Workbooks(wname).Worksheets("F08F")
        NumData = 2
        Call Vertical2Horizontal(NextRow, 1, Line8F - 3, Col8F, 2, Output, FormIn)
        numSecs = FormIn.Cells(Line8F - 3, Col8F).Value2
        NumData = 1
        Call Range2DSame(NextRow, 1, Line8F, Col8F, numSecs, NumData, Output, FormIn)
        Col8F = Col8F + 3
    Loop
    Exit Sub

ErrorProc:
    MsgBox "Error in procedure WriteForm8 : " & Err.Description
    Err.Clear
End Sub

Private Sub WriteForm11()
    Dim RowF11A As Long
    Dim RowF11B As Long
    Dim ColF11A As Long
    Dim ColF11B As Long
    Dim Numzones As Long
    Dim F11B As Variant
    Dim numlines As Integer
    
    Set FormIn = Workbooks(wname).Worksheets("F11A")
    Set F11B = Workbooks(wname).Worksheets("F11B")
    col = 2
    RowF11A = 5                                  'Starting Row of Form F11a
    ColF11A = 2                                  'Starting Column of Form F11a
    RowF11B = 3                                  'starting Row of Form FllB
    ColF11B = 2                                  'Starting Column of Form F11b
    NumCol = 6
    Do While FormIn.Cells(RowF11A, ColF11A) <> ""
        Output.Range(Output.Cells(NextRow, 1), Output.Cells(NextRow, NumCol)).Value2 = _
                                                                                     FormIn.Range(FormIn.Cells(RowF11A, ColF11A), FormIn.Cells(RowF11A, ColF11A + NumCol - 1)).Value2
        NextRow = NextRow + 1
        If NumECZ > 1 Then
            Numzones = FormIn.Cells(RowF11A, ColF11A + 1)
            numlines = Int_International(Numzones / 8) + 1
            If (Numzones Mod 8 = 0) Then numlines = numlines - 1
            'NumCol = 8
            Row = RowF11B
            For i = 1 To numlines
                Call Vertical2Horizontal(NextRow, 1, Row, ColF11B, 8, Output, F11B)
                Row = Row + 8
            Next i
        End If
        RowF11A = RowF11A + 1
        ColF11B = ColF11B + 1
    Loop
            
End Sub

Sub Range2DSame(routput As Long, coutput As Long, rinput As Long, cinput As Long, rdata As Long, cdata As Long, outputname As Variant, inputname As Variant)
    Range2D(routput, coutput, cdata, rdata, outputname).Value2 = Range2D(rinput, cinput, cdata, rdata, inputname).Value2
    NextRow = NextRow + rdata
End Sub

Function Range2D(rstart As Long, cstart As Long, cdata As Long, rdata As Long, sname As Variant) As Range 'Starting Line, Number of Data Points. Consider adding additional data
    With sname
        Set Range2D = .Range(.Cells(rstart, cstart), .Cells(rstart + rdata - 1, cstart + cdata - 1))
    End With
End Function

Sub Vertical2Horizontal(routput As Long, coutput As Long, rinput As Long, cinput As Long, VerticalInput As Long, outputname As Variant, inputname As Variant)
    Range2D(routput, coutput, VerticalInput, 1, outputname).Value2 = Application.Transpose(Range2D(rinput, cinput, 1, VerticalInput, inputname).Value2)
    NextRow = NextRow + 1
End Sub


