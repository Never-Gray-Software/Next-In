Attribute VB_Name = "Read_File"
' Project Name: Next-In
' Description: Reads SES input files into Next-In Excel format
' Copyright (c) 2025 Justin Edenbaum, Never Gray
' This file is licensed under the MIT License.
' You may obtain a copy of the license at https://opensource.org/licenses/MIT

Option Explicit

Public File As String, mydate As String, mytime As String
Public ECO As Integer, counter As Integer, EndSub As Integer, NumofSub As Integer, RowAdd As Integer
Public THO As Integer
Public NSECT As Integer
Public NVENTS As Integer
Public NNODES As Integer
Public NUHS As Integer
Public NFANS As Integer
Public NTR As Integer
Public NTT As Integer
Public NECZ As Integer
Public NTIOAI As Integer
Public NIFT As Integer
Public Infile
'JME Variables
Private x, Index, Lstart As Integer
Private L As Integer
Private Test As Boolean
Public f1c01, f1c31, f1d01, f1d11, f1d21, f1d31, f1d41, f1d51, f1d61, f1e01, f1e11, f1e21, f1e51, f1e41, f1e71, f1h01, f1h41, f8f01 As Variant
Private numsegment As Integer                    'used for form 5 segment
Private segtype As Integer                       'used for form 5 segment
Public wname As String
Public DataArray() As String
Dim r As Long
Dim c As Long
Dim inline As String
Dim a As String
Dim b As String
Dim d As String
Dim ipversion As Boolean
Dim last_line_with_data As String

Sub ReadFile(Optional unit_test As String)
    On Error GoTo ErrorProc
    Dim StartTime As Variant
    Dim cell_value, read_date, read_time As Variant
    Dim read_info As String
    Dim FormIn, Output As Worksheet
    Dim FormRange As Range
    Dim directory_path As String
    wname = ActiveWorkbook.Name
    'Select Input file with selection screen
    If unit_test = "" Then 'Call dialog box
        directory_path = Extract_Directory_Path(last_read_file.Value2)
        Call choosefile(Infile, directory_path)
        If Infile = "" Then                      'Quit if there is no input file.
            Call Speedon(False)
            Exit Sub
        End If
    Else                                         ' unit test is being conducted so infile is equal to another string
        Infile = unit_test
    End If
    Call Speedon(True)
    StartTime = Timer
    ipversion = is_version_ip(wname)
    cell_value = last_read_version.Value2 'Value of last read in
    WriteForm.Show vbModeless
    WriteForm.TextBox2.value = "Adjusting Version"
    Call ip_switch(wname, ipversion, cell_value)
    If Not ipversion Then
        last_read_version.Value2 = "(SES 6.0)"
    Else
        last_read_version.Value2 = "(SES 4.1)"
    End If
    read_date = Date
    read_time = Time
    read_info = "Last Read on " & read_date & " at " & read_time & ":"
    last_read_time.Value2 = read_info
    last_read_file.Value2 = Infile
    Workbooks(wname).Worksheets("Control").Range("G21").Value2 = Workbooks(wname).BuiltinDocumentProperties("Last Author")
    WriteForm.TextBox2.value = "Reading input into memory"
    WriteForm.Repaint
    Call TextFileToArray(Infile)                 'Create an DataArray from the text file for faster processing
    WriteForm.TextBox2.value = "Clearing Forms"
    WriteForm.Repaint
    Call ClearForms(wname)
    WriteForm.TextBox2.value = "Writing Formulas" '2p3 Moved up so default files are overwritten
    WriteForm.Repaint
    Call Formulas(wname)
    WriteForm.TextBox2.value = "Reading Form 1"
    WriteForm.Repaint
    Call ReadForm1v2
    WriteForm.TextBox2.value = "Reading Form 2"
    WriteForm.Repaint
    Call ReadForm2v2
    WriteForm.TextBox2.value = "Reading Form 3"
    WriteForm.Repaint
    Call ReadForm3v2
    WriteForm.TextBox2.value = "Reading Form 4"
    WriteForm.Repaint
    Call ReadForm4v2
    Call ReadForm5v2
    WriteForm.TextBox2.value = "Reading Form 6"
    WriteForm.Repaint
    Call ReadForm6v2
    WriteForm.TextBox2.value = "Reading Form 7 A and B"
    WriteForm.Repaint
    Call ReadForm7ABv2
    WriteForm.TextBox2.value = "Reading Form 7 C and D"
    WriteForm.Repaint
    Call ReadForm7Cv2
    Call ReadForm7Dv2
    WriteForm.TextBox2.value = "Reading Form 8"
    WriteForm.Repaint
    Call ReadForm8v3
    WriteForm.TextBox2.value = "Reading Form 9"
    WriteForm.Repaint
    Call ReadForm9v3
    WriteForm.TextBox2.value = "Reading Form 10"
    WriteForm.Repaint
    Call ReadForm10v2
    WriteForm.TextBox2.value = "Reading Form 11"
    WriteForm.Repaint
    Call ReadForm11v2
    WriteForm.TextBox2.value = "Reading Form 12"
    WriteForm.Repaint
    Call ReadForm12v2
    WriteForm.TextBox2.value = "Reading Form 13"
    WriteForm.Repaint
    Call ReadForm13v2
    WriteForm.TextBox2.value = "Reading Form 14"
    WriteForm.Repaint
    Call ReadForm14v2
    WriteForm.TextBox2.value = "Restart File"
    WriteForm.Repaint
    Call ReadInitializationFile
    WriteForm.Hide
    Debug.Print unit_test & " Time to read Input is: " & (Timer - StartTime)
    Call Speedon(False)
    If unit_test = "" Then MsgBox "Finished Reading in File"
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Read : " & Err.Description
    Call Speedon(False)
    Err.Clear
End Sub

Public Sub choosefile(Infile, Optional directory_path As String)
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
    Dim open_file_dialog As Boolean
    open_file_dialog = True
    While open_file_dialog
        With FD
            'Use the Show method to display the File Picker dialog box and return the user's action.
            'The user pressed the action button.
            .InitialFileName = directory_path
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "SES Input Files", "*.SES; *.INP; *.SVS", 1
            If .Show = -1 Then
                For Each vrtSelectedItem In .SelectedItems
                    'vrtSelectedItem is a String that contains the path of each selected item.
                    Infile = vrtSelectedItem
                Next vrtSelectedItem
            Else: Infile = ""
            End If
            'FilePath must be local. I cannot be an http link, which can happen if Excel file is on Sharepoint.
            If InStr(1, Infile, "http://", vbTextCompare) > 0 Or _
                InStr(1, Infile, "https://", vbTextCompare) > 0 Then
                MsgBox "Please select an local input file on a lettered drive (c:\). The current file path references a website, which happens with files on sharepoint."
                open_file_dialog = True
            Else: open_file_dialog = False
            End If
        End With
        'Set the object variable to Nothing.
    Wend
    Set FD = Nothing
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure choosefile : " & Err.Description
    Err.Clear
End Sub

Function is_version_ip(wname)
    If si_ip_option = 2 Then
        is_version_ip = True
    Else
        is_version_ip = False
    End If
End Function

Private Sub ReadForm1v2()
    On Error GoTo ErrorProc
    L = 0                                        'Count Lines starts at Zero by setting this value to -1. This mataches array values
    Set FormIn = Workbooks(wname).Worksheets("F01")
    With Workbooks(wname).Worksheets("F01")
        r = 3
        c = 4
        .Cells(r, c).Value2 = JoinElements(L, 1, 8) 'system idenfication
        L = L + 1
        r = r + 1
        'Test text for ses data or additional title lines
        Test = False
        Do Until Test = True
            a = Trim(DataArray(L, 0))
            b = Trim(DataArray(L, 1))
            d = Trim(DataArray(L, 2))
            a = Application.WorksheetFunction.Text(a, 0#)
            b = Application.WorksheetFunction.Text(b, 0#)
            d = Application.WorksheetFunction.Text(d, 0#)
            If IsNumeric(a) And IsNumeric(b) And IsNumeric(d) Then Test = True
            If Test = False Then
                .Cells(r, c).Value2 = JoinElements(L, 1, 8)
                L = L + 1
                r = r + 1                        'column does not change, row does
                If L > 20 Then
                    MsgBox "Error in Form 1B"
                    Test = True
                End If
            End If
        Loop
        r = 23                                   'If integers then input as dates start of Row for Hour
        x = 1
        .Cells(r, c).Value2 = a                  'design hour
        r = r + 1
        .Cells(r, c).Value2 = b                  'design month
        r = r + 1
        .Cells(r, c).Value2 = d                  'design year
        r = r + 1
        L = L + 1
        Call AL2Vertical(L, 8, r, c, FormIn)     ' Form 1C
        f1c01 = Int_International(DataArray(L - 1, 0))                'Form 1C, Train Performance Option
        f1c31 = Int_International(DataArray(L - 1, 3))                'Form 1C, Environmental Control Load Option
        f1d01 = Int_International(DataArray(L, 0))                  'form 1D, 1. Number of Line segments
        r = r + 1
        x = x + 10
        f1d11 = Int_International(DataArray(L, 1))                  'form 1D, 2. Number of sections
        r = r + 1
        f1d21 = Int_International(DataArray(L, 2))                  'form 1D, 3. Vent shaft sections
        r = r + 1
        f1d31 = Int_International(DataArray(L, 3))                  'form 1D 4. Number of Nodes)
        r = r + 1
        .Cells(r, c).Value2 = Int_International(DataArray(L, 4))    'Form 1D 5. Number of Branched Junctions
        r = r + 1
        If Not ipversion Then                    'Form 1D is different for SI and IP versions.
            f1d51 = Int_International(DataArray(L, 5))              'Form 1D Unsteady heat sources
            r = r + 1
            f1d61 = Int_International(DataArray(L, 6))              'form 1D IP Number of fan types
            r = r + 1
        Else
            'Skip entry 6 (or array spot 5) for "Number of Portals" in IP Version
            f1d51 = Int_International(DataArray(L, 5 + 1))          'Form 1D Unsteady heat sources
            r = r + 1
            f1d61 = Int_International(DataArray(L, 6 + 1))          'form 1D IP Number of fan types
            r = r + 1
        End If
        L = L + 1
        f1e01 = Int_International(DataArray(L, 0))                  'form 1E Number of Train routes
        r = r + 1
        f1e11 = Int_International(DataArray(L, 1))                  'form 1E Number of train types
        r = r + 1
        f1e21 = Int_International(DataArray(L, 2))                  'form 1E Number of Envrio. zones
        r = r + 1
        .Cells(r, c).Value2 = DataArray(L, 3)    'form 1E Fan Stopping/Windmilling Option
        r = r + 1
        f1e41 = Int_International(DataArray(L, 4))                  'form 1E number of trains in Op at init.
        r = r + 1
        f1e51 = Int_International(DataArray(L, 5))                  'form 1E impulse fan types
        r = r + 1
        .Cells(r, c).Value2 = DataArray(L, 6)    'Initization fire writing option
        r = r + 1
        .Cells(r, c) = DataArray(L, 7)           'Initization file reading option
        f1e71 = Int_International(Trim(DataArray(L, 7)))
        r = r + 1
        L = L + 1                                'Next Line for form 1F
        Call AL2Vertical(L, 8, r, c, FormIn)     'Form 1F
        Call AL2Vertical(L, 8, r, c, FormIn)     'Form 1G
        'Form 1h
        If Not ipversion Then
            f1h01 = Int_International(DataArray(L, 0))              '1h Number of Air Curtain Types
            r = r + 1
            .Cells(r, c).Value2 = Int_International(DataArray(L, 1)) '1h Number of years in the heat sink
            r = r + 1
            .Cells(r, c).Value2 = DataArray(L, 2) 'Heat Sink Attenuation Factor
            r = r + 1
            .Cells(r, c).Value2 = DataArray(L, 3) 'Night time Cooling Option
            r = r + 1
            f1h41 = Int_International(DataArray(L, 4))              'Number of Cool Pipes
            L = L + 1
        End If
    End With
  
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm1 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm2v2()
    On Error GoTo ErrorProc
    Dim f2a As Integer
    Dim f2B As Integer
    If Int(f1d11 - f1d21) > 0 Then
        Set FormIn = Workbooks(wname).Worksheets("F02A")
        'form 2A
        r = 5
        c = 2
        Call AL2E2D(L, 5, r, c, Int(f1d11 - f1d21), FormIn)
    End If
    If Int(f1d21) > 0 Then                       'Conditional statement for empty Form 2
        Set FormIn = Workbooks(wname).Worksheets("F02B")
        r = 5
        c = 2
        Call AL2E2D(L, 5, r, c, Int(f1d21), FormIn)
    End If
    Exit Sub

ErrorProc:
    MsgBox "Error in procedure ReadForm2 v2 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm3v2()                        'reads information from data array into excel sheet
    On Error GoTo ErrorProc
    Dim nsegrange As Variant
    Dim f3 As Integer
    Dim f3c51 As String
    Dim f3c61 As String
    Dim counter2 As Integer
    Dim ezcontrol As Integer
    Dim Lstart, lend As Integer
  
    With Workbooks(wname).Worksheets("F03")
        Set FormIn = Workbooks(wname).Worksheets("F03")
        r = 5                                    'Starting Row
        nsegrange = f1d01
        For f3 = 1 To Int(f1d01)                 'number of segments
            c = 2                                'Starting Coloumn
            Call AL2E(L, 2, r, c, FormIn)        'Form 3A - First line, Indentification Number and Line Segment Type
            FormIn.Cells(r, c).Value2 = JoinElements(L - 1, 3, 8) 'Form 3A Identification title
            c = c + 1
            Call AL2E(L, 5, r, c, FormIn)        'Form 3A - Second Line
            Call AL2E(L, 8, r, c, FormIn)        'Form 3B Perimeters
            Call AL2E(L, 8, r, c, FormIn)        'Form 3B Roughness
            Call AL2E(L, 7, r, c, FormIn)        'Form 3C Line segment data
            f3c51 = Int_International(DataArray(L - 1, 5))
            f3c61 = Int_International(DataArray(L - 1, 6))
            counter = 0
            Do While (counter < Val(f3c61) - 1 Or counter = Val(f3c61) - 1)
                Call AL2E(L, 5, r, c, FormIn)    'Form 3D Line segment data
                FormIn.Cells(r, c).Value2 = JoinElements(L - 1, 6, 8) 'Form 3D Identification
                c = c + 1
                counter = counter + 1
                If counter < Val(f3c61) Then
                    r = r + 1
                    c = c - 6
                End If
            Loop
            If Val(f3c61) = 0 Then
                c = c + 6
                counter = 1
            End If
            r = r - (counter - 1)                'moves the row back up
            EndSub = 0                           'start with an ending subsgement of zero
            counter2 = 0
            Do While (EndSub < Val(f3c51))
                Call AL2E(L, 5, r, c, FormIn)    'Form 3D Line segment data
                EndSub = Int_International(DataArray(L - 1, 1))
                If EndSub < Val(f3c51) Then
                    r = r + 1
                    c = c - 5
                End If
                counter2 = counter2 + 1
                If counter > 99 Then
                    MsgBox "Error in Form 3E, Ending Sub-Segment Number"
                    End
                End If
            Loop
            r = r - (counter2 - 1)
            ezcontrol = Int(f1c31)
            If ezcontrol = 1 Or ezcontrol = 2 Then 'if ezcontrol is 1 or 2 then read line segment data
                Call AL2E(L, 7, r, c, FormIn)    'Form 3F Line segment data
            End If
            RowAdd = 1
            If counter > RowAdd Then RowAdd = counter
            If counter2 > RowAdd Then RowAdd = counter2
            r = r + RowAdd
        Next f3
    End With
  
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm3 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm4v2()                        'reads information form text file into memory
    On Error GoTo ErrorProc
    Dim f4 As Integer
    Set FormIn = Workbooks(wname).Worksheets("F04")
    r = 5
    If Val(f1d51) = 0 Then Exit Sub              'exit the subroutine if no form 4 data
    If Val(f1d51) > 0 Then                       'write form 4 data if there is unsteady sources
        For f4 = 1 To Val(f1d51)
            c = 2
            FormIn.Cells(r, c).Value2 = JoinElements(L, 1, 4) 'read in description
            c = c + 1
            FormIn.Cells(r, c) = DataArray(L, 4)
            c = c + 1
            FormIn.Cells(r, c) = DataArray(L, 5)
            L = L + 1                            'Move to Next Line
            c = c + 1
            Call AL2E(L, 6, r, c, FormIn)        ' Read in Form 4, Line 2
            r = r + 1
        Next f4
    End If
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm4v2 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm5v2()                        'reads information form text file into memory
    On Error GoTo ErrorProc
    Dim f3 As Integer
    Dim msg As String
    With Workbooks(wname).Worksheets("F05")
        Set FormIn = Workbooks(wname).Worksheets("F05")
        r = 5
        If Val(f1d21) > 0 Then
            For f3 = 1 To Val(f1d21)
                c = 2
                segtype = DataArray(L, 1)
                Call AL2E(L, 2, r, c, FormIn)    'Form 5A
                FormIn.Cells(r, c).Value2 = JoinElements(L - 1, 3, 8) 'Form 5A Identification title
                c = c + 1
                Call AL2E(L, 8, r, c, FormIn)    'Read in form 5B
                numsegment = Int_International(DataArray(L - 1, 0)) 'Number of subsegments
                If Val(f1d61) > 0 And segtype <> 3 Then 'if there is more then one fan read in form 5C
                    Call AL2E(L, 4, r, c, FormIn) 'Read in Form 5C
                Else: c = c + 4
                End If
                'Enter code for Form 5C-A
                If segtype = 3 Then
                    Call AL2E(L, 4, r, c, FormIn) 'Read in code for Form 5C - A
                Else: c = c + 4
                End If
                ' Form 5D
                counter = 0
                Do While counter < numsegment
                    Call AL2E(L, 7, r, c, FormIn) 'Read in Form 5D
                    counter = counter + 1
                    If counter < numsegment Then
                        c = c - 7
                        r = r + 1
                    End If
                    If counter > 99 Then
                        MsgBox "Error in Form 5B, Ending Segment Number"
                        End
                    End If
                Loop
                r = r + 1
            Next f3
        End If
    End With
  
    Exit Sub
ErrorProc:
    msg = "Error in procedure ReadForm5 : " & Err.Description
    If Err.Description = "Type mismatch" Then
        msg = msg & vbCrLf & "This can occur when the wrong Unit is selected: SI or IP."
        MsgBox msg
    Else
        MsgBox "Error in procedure ReadForm5 : " & Err.Description
        Err.Clear
    End If
End Sub

Private Sub ReadForm6v2()                        'reads information form text file into memory
    On Error GoTo ErrorProc
    Dim a(5, 1) As Integer
    Dim ntype As Integer
    Dim f6 As Integer
    Dim NTherm As Integer
    Dim MyArray As Variant
  
    Set FormIn = Workbooks(wname).Worksheets("F06")
    r = 5
    For f6 = 1 To Val(f1d31)
        c = 2
        Call AL2E(L, 3, r, c, FormIn)            ' Read in Form 6A. ntype and Ntherm called afterwards. Otherwise, the dataArray value gets a comma.
        ntype = Int_International(DataArray(L - 1, 1))           'Node Aero Type. L is subtracted by 1 because AL2E advances the line number
        NTherm = Int_International(DataArray(L - 1, 2))          'Node Thermo Type
        If NTherm = 3 Then
            Call AL2E(L, 6, r, c, FormIn)        'If type 3, read in form 6 boundary condtions
        End If
        'An array is used to keep track of the column number needed based on node
        'Node#          0   1   2   3   4   5   6   7   8
        MyArray = Array(0, 11, 16, 19, 23, 28, 40, 0, 46)
        If ntype <> 0 And ntype <> 7 Then        'Node type zero does not have any other entries
            c = MyArray(ntype)
            If ntype = 6 Or ntype = 8 Then       'If node type 6 or 8, read in six data points
                Call AL2E(L, 6, r, c, FormIn)
            Else
                Call AL2E(L, 5, r, c, FormIn)    'All other node types read in five dat points
            End If
        End If
        If NTherm = 2 Then                       'For thermodynamic node type 2
            Call AL2E(L, 7, r, c, FormIn)
        End If
        r = r + 1
    Next f6
  
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm6v2 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm7ABv2()                      'reads information form text file into memory
    On Error GoTo ErrorProc
    If f1d61 = 0 Then Exit Sub                   'there should be more then one fan
    Dim f7 As Integer
    Dim i As Integer
    Set FormIn = Workbooks(wname).Worksheets("F07")
    r = 5
    For f7 = 1 To Val(f1d61)
        c = 2
        FormIn.Cells(r, c).Value2 = JoinElements(L, 1, 4)
        c = c + 1
        For i = 4 To 7
            FormIn.Cells(r, c).Value2 = DataArray(L, i)
            c = c + 1
        Next i
        L = L + 1
        Call AL2E(L, 8, r, c, FormIn)            'Read in Form 7B
        Call AL2E(L, 8, r, c, FormIn)            'Read in Form 7B
        r = r + 1
    Next f7
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm7AB : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm7Cv2()
    On Error GoTo ErrorProc
    Dim f7 As Integer
    Dim i As Integer
   
    Set FormIn = Workbooks(wname).Worksheets("F07C")
    If f1e51 = 0 Then Exit Sub                   'there should be more then one fan
    r = 5
    c = 2
    Call AL2E2D(L, 7, r, c, Val(f1e51), FormIn)
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm7C : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm7Dv2()
    On Error GoTo ErrorProc
    Dim f7 As Integer
    Dim i As Integer
    
    Set FormIn = Workbooks(wname).Worksheets("F07D")
    If f1h01 = 0 Then Exit Sub                   'there should be more then one fan
    r = 5
    c = 2
    Call AL2E2D(L, 8, r, c, Val(f1h01), FormIn)
    
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm7D : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm8v3()
    On Error GoTo ErrorProc
    Dim f8 As Long
    Dim r8a As Long
    Dim r8b As Long
    Dim r8c As Long
    Dim r8d As Long
    Dim r8e As Long
    Dim r8f As Long
    Dim c8a As Long
    Dim c8b As Long
    Dim c8c As Long
    Dim c8d As Long
    Dim c8e As Long
    Dim c8f As Long
    Dim rk As Long
    Dim i As Integer
    Dim f8b As Long
    Dim f8c As Integer
    Dim ii As Integer
    Dim nstops As Integer
    Dim iii As Integer
    Dim numpp As Integer
    Dim f8f01 As Integer
    If Val(f1c01) = 0 Then Exit Sub              'skip if trainperformance equal zero
    r = 4
    r8a = 5                                      'initial row number for form 8a
    r8b = 4                                      'initial row number for form 8b
    r8c = 5                                      'initial row number for form 8c
    r8d = 7                                      'initial row number for form 8d
    r8e = 6                                      'initial row number for form 8e
    r8f = 5                                      'initial row number for form 8f
    c8a = 2                                      'initial column number for form 8b
    c8b = 2                                      'initial column number for form 8b
    c8c = 2                                      'initial column number for form 8c
    c8d = 2                                      'initial column number for form 8d
    c8e = 2                                      'initial column number for form 8e
    c8f = 4                                      'initial column number of form 8e
    rk = 1                                       'initial number for route key
    For f8 = 1 To Val(f1e01)
        Set FormIn = Workbooks(wname).Worksheets("F08A")
        c = c8a
        FormIn.Cells(r8a, c).Value2 = JoinElements(L, 1, 7) 'Form 8A Train Route Description, Line 1
        c = c + 1
        L = L + 1
        'Form 8A, Line 2 - Do not overwrite formulas that calculate number of groups of trains and tracks
        FormIn.Cells(r8a, c).Value2 = Val(DataArray(L, 0)) 'Train Schedule Origin
        f8b = Int_International(DataArray(L, 1))               ' number of groups of trains that could enter route
        f8c = Int_International(DataArray(L, 2))               ' number of groups of track sections
        FormIn.Cells(r8a, c + 3).Value2 = Val(DataArray(L, 3)) 'Time Delay
        FormIn.Cells(r8a, c + 4).Value2 = Int_International(DataArray(L, 4)) 'First Train Type
        FormIn.Cells(r8a, c + 5).Value2 = Val(DataArray(L, 5)) 'Minimum Coast Velocity
        FormIn.Cells(r8a, c + 6).Value2 = Int_International(DataArray(L, 6)) 'Coast option
        L = L + 1
        r8a = r8a + 1
        If f8b > 1 Then                          'Form 8B
            Set FormIn = Workbooks(wname).Worksheets("F08B")
            Call AL2E2D(L, 3, r8b, c8b, f8b - 1, FormIn)
        End If
        c8b = c8b + 3                            'Move column over to next route
        'Form 8C and 8D read in for Train Performance Option 1
        If Val(f1c01) = 1 Then                   'Form 1C, Train Performance Option
            Set FormIn = Workbooks(wname).Worksheets("F08C")
            Call AL2E2D(L, 8, r8c, c8c, f8c, FormIn)
            c8c = c8c + 8                        'keep track of rows for form 8c
            Set FormIn = Workbooks(wname).Worksheets("F08D")
            nstops = Int_International(DataArray(L, 0)) 'Formula calculates number of stops, so don't over this information
            FormIn.Cells(3, c8d + 2).Value2 = (DataArray(L, 1)) 'passenger at origin
            L = L + 1
            'Call AL2Vertical(L, 2, r8d - 5, c8d + 2, FormIn)
            'If there is one or stops read write them to the worksheet
            If nstops > 0 Then Call AL2E2D(L, 3, r8d, c8d, nstops, FormIn)
            c8d = c8d + 3                        'Move over columns for the next route
        End If
        If Val(f1c01) = 2 Or Val(f1c01) = 3 Then 'form 8 E
            Set FormIn = Workbooks(wname).Worksheets("F08E")
            numpp = Int_International(DataArray(L, 0))
            L = L + 1
            Call AL2E2D(L, 5, r8e, c8e, numpp, FormIn)
            c8e = c8e + 5                        'Move over columns for the next route
        End If
        Set FormIn = Workbooks(wname).Worksheets("F08F") 'read in form 8f data
        f8f01 = Int_International(DataArray(L, 0))             'number of sections
        'Call AL2Vertical(L, 2, r8f - 3, c8f, FormIn) 'Read in two
        FormIn.Cells(3, c8f).Value2 = (DataArray(L, 1)) 'Distance to Portal
        L = L + 1
        Call FirstArrayLines2Vertical(L, f8f01, r8f, c8f, FormIn)
        c8f = c8f + 3
    Next f8
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm8 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm9v3()
    On Error GoTo ErrorProc
    Dim f9 As Integer
    Dim i As Integer
    Dim ii As Integer
    Dim tcontrol As Integer
    Dim Flysim As Integer
  
    If Val(f1c01) = 0 Then Exit Sub              'skip if trainperformance equal zero
    Set FormIn = Workbooks(wname).Worksheets("F09")
    c = 6                                        'Starting column
    For f9 = 1 To Val(f1e11)
        r = 3                                    ' Starting row
        FormIn.Cells(r, c).Value2 = JoinElements(L, 1, 4) 'Read in Descrioption of Form 9A
        r = r + 1
        For i = 4 To 7
            FormIn.Cells(r, c).Value2 = DataArray(L, i)
            r = r + 1
        Next i
        L = L + 1
        Call AL2Vertical(L, 5, r, c, FormIn)     'Read in Form 9B
        Call AL2Vertical(L, 6, r, c, FormIn)     'Read in Form 9c
        Call AL2Vertical(L, 8, r, c, FormIn)     'Read in Form 9D, Part 1
        Call AL2Vertical(L, 8, r, c, FormIn)     'Read in Form 9D, Part 2
        Call AL2Vertical(L, 6, r, c, FormIn)     'Read in Form 9E
        If Val(f1c01) = 1 Then                   'if train performance option is 1
            FormIn.Cells(r, c).Value2 = JoinElements(L, 1, 4) 'Form 9F Motor ID Name
            r = r + 1
            For i = 4 To 5
                FormIn.Cells(r, c).Value2 = DataArray(L, i)
                r = r + 1
            Next i
            L = L + 1
            Call AL2Vertical(L, 5, r, c, FormIn) 'Form 9F - Continue
            Call AL2Vertical(L, 4, r, c, FormIn) 'Form 9G 1 of 4
            Call AL2Vertical(L, 4, r, c, FormIn) 'Form 9G 2 of 4
            Call AL2Vertical(L, 4, r, c, FormIn) 'Form 9G 3 of 4
            FormIn.Cells(r, c).Value2 = DataArray(L, 0)
            tcontrol = Int_International(DataArray(L, 0))      ' Form 9G 4 of 4
            L = L + 1
            r = r + 1
            If tcontrol = 2 Then                 'read in form 9h data if train control is 2
                Call AL2Vertical(L, 5, r, c, FormIn)
                Call AL2Vertical(L, 5, r, c, FormIn)
                Flysim = Int_International(DataArray(L - 1, 4))
            ElseIf tcontrol = 3 Then
                r = 72
                Call AL2Vertical(L, 5, r, c, FormIn) 'read in form 9H-A, Part 1
                Call AL2Vertical(L, 5, r, c, FormIn) 'read in form 9H-A, Part 2
                Call AL2Vertical(L, 5, r, c, FormIn) 'read in form 9H-A, Part 3
            End If
            r = 87                               'Form 9I starting row
            Call AL2Vertical(L, 5, r, c, FormIn) 'read in form 9I
            Call AL2Vertical(L, 5, r, c, FormIn) 'read in form 9J
            If Flysim = 2 Then                   'read in Flywheel data if fly wheels are simulated
                Call AL2Vertical(L, 5, r, c, FormIn) 'Form 9K
                Call AL2Vertical(L, 7, r, c, FormIn)
            End If
        End If
        c = c + 1                                'Go to next column
    Next f9
  
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm9V2 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm10v2()                       'Needs to be Updated with Array
    On Error GoTo ErrorProc
    Dim f10 As Integer
    Dim i As Integer
  
    If Val(f1c01) = 0 Then Exit Sub              'skip if the train performance is zero
    If Val(f1e41) < 1 Then Exit Sub
    Set FormIn = Workbooks(wname).Worksheets("F10")
    r = 5
    c = 2
    Call AL2E2D(L, 8, r, c, Val(f1e41), FormIn)
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm10 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm11v2() '2p4 Update to not overwrite formulas to count number of line and vent shafts
    On Error GoTo ErrorProc
    Dim R11A As Long
    Dim R11B As Long
    Dim C11A As Long
    Dim C11B As Long
    Dim f11 As Integer
    Dim i As Integer
    Dim NLS As Integer
    Dim Numline As Integer
    Dim ii As Integer
    Dim F11A, F11B As Variant
    Dim stop_loop As Boolean
    Set F11A = Workbooks(wname).Worksheets("F11A")
    Set F11B = Workbooks(wname).Worksheets("F11B")
    R11A = 5
    R11B = 3
    C11A = 2
    C11B = 2
    For f11 = 1 To Val(f1e21)
        F11A.Cells(R11A, C11A).Value2 = Int_International(DataArray(L, 0)) 'Zone Type
        NLS = Int_International(DataArray(L, 1))                      'Number of Line and Vent Segements to variable
        If Int_International(DataArray(L, 0)) = 1 Then
            For i = 2 To 5                                  'Skip Column for number of vent zones. Input remaining four values
                F11A.Cells(R11A, 2 + i).Value2 = DataArray(L, i)
            Next
        End If
        L = L + 1
        If Val(f1e21) > 1 Then                   'Environmental Control Zone
            Numline = Int_International(NLS / 8) + 1
            If (NLS Mod 8 = 0) Then Numline = Numline - 1
            Call EightArrayLines2Vertical(L, Numline, R11B, C11B, F11B)
        End If
        C11B = C11B + 1
        R11A = R11A + 1
    Next f11
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm11 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm12v2()
    On Error GoTo ErrorProc
    Dim npg As Integer
    Dim ii As Integer
    Dim i As Integer
  
    Set FormIn = Workbooks(wname).Worksheets("F12")
    'Line Input #1, inline 'Read in Form F12
    FormIn.Range("D2").value = DataArray(L, 0)
    npg = Int_International(DataArray(L, 1))                   'Number of groups
    L = L + 1                                    'Advance line count to the next cell
    r = 5
    c = 2
    Call AL2E2D(L, 7, r, c, npg, FormIn)
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm12 : " & Err.Description
    Err.Clear
    
End Sub

Private Sub ReadForm13v2()
    On Error GoTo ErrorProc
    Dim i As Integer
    Set FormIn = Workbooks(wname).Worksheets("F13")
    r = 4
    c = 2
    Call AL2E(L, 4, r, c, FormIn)
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm13 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadForm14v2()                       'Form 14c
    On Error GoTo ErrorProc
    Dim F14A, F14B, F14C, i, x, N, p As Integer
    Dim C14C As Long
    Dim NumIn As Integer
    Dim numSec As Integer
    Dim R14C As Long
    Dim FormF14AB As Variant
    Dim FormF14C As Variant
  
    If f1h41 >= 1 Then
        r = 4
        C14C = 2                                 'starting column on Form 14C
        Set FormF14AB = Workbooks(wname).Worksheets("F14AB")
        Set FormF14C = Workbooks(wname).Worksheets("F14C")
        For F14A = 1 To Val(f1h41)
            c = 2                                'start at left hand column
            FormF14AB.Cells(r, c).Value2 = JoinElements(L, 1, 8) ''Read form 14A
            L = L + 1
            c = c + 1
            Call AL2E(L, 7, r, c, FormF14AB)     'Form 14 A, Part 2
            NumIn = Int_International(DataArray(L - 1, 0))
            numSec = Int_International(DataArray(L - 1, 1))
            Call AL2E2D(L, 3, r, c, NumIn, FormF14AB) 'Form14B
            r = r + NumIn                        'Update Row Number
            R14C = 4
            Call AL2E2D(L, 1, R14C, C14C, numSec, FormF14C)
            C14C = C14C + 1                      ' Move to the next column
        Next F14A
    End If
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadForm14 : " & Err.Description
    Err.Clear
End Sub

Private Sub ReadInitializationFile()
    On Error GoTo ErrorProc
    If Val(f1e71) = 0 Then Exit Sub
    Set FormIn = Workbooks(wname).Worksheets("F01")
    FormIn.Cells(48, 7) = last_line_with_data
    If Len(last_line_with_data) > 255 Then
        MsgBox "Careful! Restart files longer than 255 characters may not work."
    End If
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure ReadInitializationFile: " & Err.Description
    Err.Clear

End Sub

Sub TextFileToArray(ByVal FilePath As String)
    'PURPOSE: Load an Array variable with data from a delimited text file
    'SOURCE: www.TheSpreadsheetGuru.com

    Dim TextFile As Integer
    Dim FileContent As String
    Dim LineArray() As String
    Dim rw As Long, col As Long
    Dim x, y, last_line_number As Integer
    Dim last_line As String
    On Error GoTo ErrorProc
  
    'Open the text file in a Read State
    TextFile = FreeFile
    Open FilePath For Input As TextFile
  
    'Store file content inside a variable
    FileContent = Input(LOF(TextFile), TextFile)

    'Close Text File
    Close TextFile
  
    'Separate Out lines of data
    LineArray() = Split(FileContent, vbCrLf)
    If UBound(LineArray) < 1 Then                'if line divider is not vbCrlF
        LineArray() = Split(FileContent, Chr(10))
    End If
    'Size DataArray for input file
    Erase DataArray
    'Debug.Print LBound(LineArray)
    'Debug.Print UBound(LineArray)
    ReDim Preserve DataArray(UBound(LineArray), 7)
    For x = LBound(LineArray) To UBound(LineArray)
        For y = 0 To 7
            DataArray(x, y) = Mid(LineArray(x), y * 10 + 1, 10)
        Next y
    Next x
    
    'Find last line with valid data because might be the restart file name.
    'Restart file is stored as a string because the array is limited to 80 characters
    last_line_number = UBound(LineArray)
    last_line_with_data = LineArray(last_line_number)
    Do While (Len(last_line_with_data) = 0)
        last_line_number = last_line_number - 1
        last_line_with_data = LineArray(last_line_number)
    Loop
    Exit Sub

ErrorProc:
    MsgBox "Error in procedure TextFileToArray: " & Err.Description
    Err.Clear
End Sub

'Determine if spreadsheet needs to change
Sub ip_switch(wname, ipversion, cell_value)
    On Error GoTo ErrorProc
    Dim Row3and4, w As Variant
    Dim ip_hide, si_hide, switch As Boolean
    Row3and4 = Array("F02A", "F02B", "F03", "F04", "F05", "F06", "F07", "F08A", "F08C", "F10", "F11A") 'SI and IP on 3 and 4
    'Last read in is NOT SI, but IP is selected
    switch = False
    If (cell_value <> "(SES 4.1)" And ipversion) Then
        si_hide = True
        ip_hide = False
        switch = True
    ElseIf (cell_value <> "(SES 6.0)" And Not ipversion) Then 'Switch to SI Only
        si_hide = False
        ip_hide = True
        switch = True
    End If
    If switch Then
        With Workbooks(wname)
            'Sheets with SI and IP on rows 3 and 4, respectively
            For Each w In Row3and4
                .Worksheets(w).Rows("3").Hidden = si_hide
                .Worksheets(w).Rows("4").Hidden = ip_hide
            Next
            'Other rows and columns to Hide
            .Worksheets("F01").Columns("E").Hidden = si_hide
            .Worksheets("F01").Columns("F").Hidden = ip_hide
            .Worksheets("F07C").Rows("3:3").Hidden = si_hide
            .Worksheets("F07C").Rows("4:4").Hidden = ip_hide
            .Worksheets("F08D").Rows("5").Hidden = si_hide
            .Worksheets("F08D").Rows("6").Hidden = ip_hide
            .Worksheets("F08E").Rows("4").Hidden = si_hide
            .Worksheets("F08E").Rows("5").Hidden = ip_hide
            .Worksheets("F09").Columns("D").Hidden = si_hide
            .Worksheets("F09").Columns("E").Hidden = ip_hide
        End With
    End If
    Exit Sub
ErrorProc:
    Workbooks(wname).Worksheets("Control").Select
    MsgBox "Error in Switching to IP or SI: " & Err.Description
    Err.Clear
End Sub

'Array Line to Excel Line (AL2E)
'Copy the line of the an array to an excel
'Input:
        'arrayline is the line number in the arrary where the data is located
        'ndata is the number of data points to take from the arrayline
        'Erow and Ecol are the row and column number on the excel sheet for the output
        'sname is excel worksheet for output
Sub AL2E(arrayline As Integer, ndata As Integer, ERow As Long, ECol As Long, sname As Variant)
    On Error GoTo ErrorProc
    OutRange(ERow, ECol, ndata, sname).Value2 = DataFromArray(arrayline, ndata)
    L = L + 1       'Advance the line number to read the next line of the array
    c = c + ndata   'Advance the number of columns for next input in Excel
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure AL2E: " & Err.Description
    Err.Clear
End Sub

'Starting Line, Number of Data Points. Consider adding additional data
Function DataFromArray(ALine As Integer, ndata As Integer) As Variant
    Dim i As Integer
    Dim Aindex As Integer
    Dim TempArray() As Variant
    On Error GoTo ErrorProc
    
    'Number of data points using array notation (starting at zero)
    Aindex = ndata - 1
    ReDim TempArray(Aindex)
    For i = 0 To Aindex
        TempArray(i) = DataArray(ALine, i)
    Next i
    DataFromArray = TempArray
    Exit Function

ErrorProc:
    MsgBox "Error in procedure DataFromArray: " & Err.Description
    Err.Clear
End Function

Function OutRange(rindex As Long, Cindex As Long, ndata As Integer, sname As Variant) As Range 'Starting Line, Number of Data Points. Consider adding additional data
    With sname
        Set OutRange = .Range(.Cells(rindex, Cindex), .Cells(rindex, Cindex + ndata - 1))
    End With
End Function

Sub AL2E2D(arrayline As Integer, ndata As Integer, ERow As Long, ECol As Long, nlines As Integer, sname As Variant) 'Array Line to Excel Line
    OutRange2D(ERow, ECol, ndata, nlines, sname).Value2 = DataFromArray2D(arrayline, ndata, nlines)
    L = L + nlines
End Sub

Function DataFromArray2D(ALine As Integer, ndata As Integer, nlines As Integer) As Variant 'Starting Line, Number of Data Points. Consider adding additional data
    Dim i, J As Integer
    Dim Aindex, Alines As Integer
    Dim TempArray() As Variant
    On Error GoTo ErrorProc
    
    'Number of data points using array notation (starting at zero)
    Aindex = ndata - 1
    Alines = nlines - 1
    ReDim TempArray(Alines, Aindex)
    For i = 0 To Alines
        For J = 0 To Aindex
            TempArray(i, J) = DataArray(ALine + i, J)
        Next J
    Next i
    DataFromArray2D = TempArray
    Exit Function

ErrorProc:
    MsgBox "Error in procedure DataFromArray2D: " & Err.Description
    Err.Clear
End Function

Function OutRange2D(rindex As Long, Cindex As Long, ndata As Integer, nlines As Integer, sname As Variant) As Range 'Starting Line, Number of Data Points. Consider adding additional data
    With sname
        Set OutRange2D = .Range(.Cells(rindex, Cindex), .Cells(rindex + nlines - 1, Cindex + ndata - 1))
    End With
End Function

Sub SkipLines(numlines As Integer)
    If numlines > 0 Then
        For x = 1 To numlines
            Line Input #1, inline
        Next x
    End If
End Sub

Function JoinElements(ALine As Integer, startcol As Integer, endcol As Integer) As String 'Combine array entries into one comprehesive string
    Dim TempString As String
    On Error GoTo ErrorProc
    JoinElements = ""                            'Initialize value to nothing
    'Number of data points using array notation (starting at zero)
    For x = startcol - 1 To endcol - 1
        JoinElements = JoinElements + DataArray(ALine, x)
    Next x

    Exit Function

ErrorProc:
    MsgBox "Error in procedure DataFromArray: " & Err.Description
    Err.Clear
End Function

Sub AL2Vertical(arrayline As Integer, ndata As Integer, ERow As Long, ECol As Long, sname As Variant) 'Array Line to Excel Line
    OutRange2D(ERow, ECol, 1, ndata, sname).Value2 = _
                                                   WorksheetFunction.Transpose(DataFromArray(arrayline, ndata))
    L = L + 1
    r = r + ndata
End Sub

Sub FirstArrayLines2Vertical(arrayline As Integer, nlines As Integer, ERow As Long, ECol As Long, sname As Variant) 'Muliple Lines of Array to Vertical Cells
    Dim i, J As Integer
    Dim Alines As Integer
    Dim TempArray() As Variant
    Alines = nlines - 1                          'array lines from nlines
    ReDim TempArray(Alines)
    For i = 0 To Alines
        TempArray(i) = DataArray(arrayline + i, 0)
    Next i
    OutRange2D(ERow, ECol, 1, nlines, sname).Value2 = _
                                                    WorksheetFunction.Transpose(TempArray)
    L = L + nlines                               'Move Lines and rows before assigning array incase error occurs
End Sub

Sub EightArrayLines2Vertical(arrayline As Integer, nlines As Integer, ERow As Long, ECol As Long, sname As Variant) 'Muliple Lines of Array to Vertical Cells
    Dim i, J, k As Integer
    Dim Alines As Integer
    Dim TempArray() As Variant
    Alines = nlines - 1                          'array lines from nlines
    ReDim TempArray(nlines * 8 - 1)
    k = 0
    For i = 0 To Alines
        For J = 0 To 7
            TempArray(k) = DataArray(arrayline + i, J)
            k = k + 1
        Next J
    Next i
    OutRange2D(ERow, ECol, 1, nlines * 8, sname).Value2 = _
                                                        WorksheetFunction.Transpose(TempArray)
    L = L + nlines                               'Move Lines and rows before assigning array incase error occurs
End Sub


