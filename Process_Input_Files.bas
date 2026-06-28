Attribute VB_Name = "Process_Input_Files"
' Project Name: Next-In
' Description: Connects buttons and information on Control Sheet to Macros.
' Copyright (c) 2025 Justin Edenbaum, Never Gray
' This file is licensed under the MIT License.
' You may obtain a copy of the license at https://opensource.org/licenses/MIT

Option Explicit

Public Sub Call_SES_Exe(workbook_name As String, input_file_path)
    Dim path_exe As String, shell_command As String
    On Error GoTo ErrorProc
    WriteForm.TextBox2.value = "Attempting to run SES"
    WriteForm.Repaint
    path_exe = SES_Exe
    If Dir(SES_Exe) <> "" Then
        shell_command = """" & path_exe & """ """ & input_file_path & """"
        Debug.Print shell_command
        Shell shell_command, vbNormalNoFocus  'Previously vbNormalFocus
    Else
        MsgBox "SES executable not found at: " & path_exe
    End If
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Call_SES_Exe: " & Err.Description
    Err.Clear
End Sub

Public Sub Call_NextOut(workbook_name As String, savename As Variant, _
    Optional next_in_path As String = "", Optional iteration_path As String = "", _
    Optional ses_version As String = "SI", Optional file_type As String = "next_in")

    On Error GoTo ErrorProc

    ' Status message
    If Len(iteration_path) = 0 Then
        WriteForm.TextBox2.value = "Attempting to run Next-Out, then SES"
    Else
        WriteForm.TextBox2.value = "Writing Iterations with Next-Out"
    End If
    WriteForm.Repaint

    ' Convert NextOut_Exe (Range ? String)
    Dim nextout_path As String
    nextout_path = CStr(NextOut_Exe.Value2)

    If Dir(nextout_path) = "" Then
        MsgBox "Next Out executable not found at: " & nextout_path
        Exit Sub
    End If

    Dim settings_dict As Object
    Dim shell_command As String
    Dim argument As String

    Set settings_dict = CreateObject("Scripting.Dictionary")

    ' --- BASIC SETTINGS ---
    settings_dict("output_conversion") = output_conversion_string
    settings_dict("file_type") = file_type
    settings_dict("output") = Get_Output_Setting(workbook_name)

    ' SES_Exe (Range ? String)
    settings_dict("path_exe") = Settings_File_Path(CStr(SES_Exe.Value2))

    ' ses_output_str MUST be a list
    If Len(CStr(savename)) = 0 Then
        settings_dict("ses_output_str") = Array("")
    Else
        settings_dict("ses_output_str") = Array(Settings_File_Path(CStr(savename)))
    End If

    settings_dict("simtime") = -1

    ' Visio template (Range ? String)
    settings_dict("visio_template") = Settings_File_Path(CStr(Visio_File.Value2))

    ' --- ITERATION PATH LOGIC (CRITICAL FIX) ---
    Dim iter_path As String
    iter_path = CStr(iteration_path)

    If Len(iter_path) = 0 Then
        ' Normal mode ? Python expects empty string
        settings_dict("iteration_path") = ""
    Else
        ' Iteration mode ? real folder path
        settings_dict("iteration_path") = Settings_File_Path(iter_path)
    End If

    ' --- SES VERSION ---
    settings_dict("ses_version") = ses_version

    ' --- NEXT-IN PATH ---
    settings_dict("next_in_path") = Settings_File_Path(CStr(next_in_path))

    ' --- SUMMARY NUMBERS (only if Summary is selected) ---
    If UBound(Filter(settings_dict("output"), "Summary")) >= 0 Then
        Dim seg As String
        seg = CStr(summary_numbers.Value2)
    
        If Len(seg) = 0 Then
            settings_dict("segments_2_lookup") = Array()
        ElseIf InStr(seg, ",") > 0 Then
            settings_dict("segments_2_lookup") = Split(seg, ",")
        Else
            settings_dict("segments_2_lookup") = Array(seg)
        End If
    End If

    ' --- SERIALIZE PYTHON DICT ---
    Dim settings_literal As String
    settings_literal = PyDict(settings_dict)
    
    ' Escape internal quotes for Windows Shell
    settings_literal = Replace(settings_literal, """", """""")
    
    argument = " --settings """ & settings_literal & """"

    Debug.Print "FINAL PYTHON DICT:"
    Debug.Print argument

    ' --- BUILD COMMAND ---
    shell_command = """" & nextout_path & """" & argument
    Debug.Print shell_command

    ' --- RUN NEXT-OUT ---
    Shell shell_command, vbNormalNoFocus

    WriteForm.TextBox2.value = "Running SES and Next-Out"
    WriteForm.Repaint
    Exit Sub

ErrorProc:
    MsgBox "Error in procedure Call_NextOut: " & Err.Description
    Err.Clear
End Sub








Function Get_Output_Setting(workbook_name As String) As Variant
    On Error GoTo ErrorProc

    Dim output_options As Collection
    Set output_options = New Collection

    Dim ws As Worksheet
    Set ws = Workbooks(workbook_name).Worksheets("Control")

    If ws.Shapes("NO_Excel").ControlFormat.value = xlOn Then
        output_options.Add "Excel"
    End If
    If ws.Shapes("NO_Route_Data").ControlFormat.value = xlOn Then
        output_options.Add "Route"
    End If
    If ws.Shapes("NO_H5_File").ControlFormat.value = xlOn Then
        output_options.Add "H5_file"
    End If
    If ws.Shapes("NO_Summary").ControlFormat.value = xlOn Then
        output_options.Add "Summary"
    End If
    If ws.Shapes("NO_Visio").ControlFormat.value = xlOn Then
        output_options.Add "Visio"

        If ws.Shapes("NO_PDF").ControlFormat.value = xlOn Then
            output_options.Add "visio_2_pdf"
        End If
        If ws.Shapes("NO_PNG").ControlFormat.value = xlOn Then
            output_options.Add "visio_2_png"
        End If
        If ws.Shapes("NO_SVG").ControlFormat.value = xlOn Then
            output_options.Add "visio_2_svg"
        End If
        If ws.Shapes("NO_Open_Visio").ControlFormat.value = xlOn Then
            output_options.Add "visio_open"
        End If
    End If

    ' Convert collection ? array
    Dim arr() As String
    ReDim arr(0 To output_options.Count - 1)

    Dim i As Long
    For i = 1 To output_options.Count
        arr(i - 1) = output_options(i)
    Next i

    Get_Output_Setting = arr
    Exit Function

ErrorProc:
    MsgBox "Error in procedure Get_Output_Setting: " & Err.Description
    Err.Clear
End Function

' abandoned
Function get_ses_version()
    If si_ip_option = 1 Then
        get_ses_version = "SI"
    Else
        get_ses_version = "IP"
    End If
End Function

Private Function PyValue(v As Variant) As String
    Dim s As String
    Dim i As Long

    If IsArray(v) Then
        s = "["
        For i = LBound(v) To UBound(v)
            s = s & PyValue(v(i)) & ", "
        Next i
        If Right(s, 2) = ", " Then s = Left(s, Len(s) - 2)
        PyValue = s & "]"

    ElseIf IsNumeric(v) Then
        PyValue = CStr(v)

    Else
        PyValue = "'" & Replace(CStr(v), "'", "''") & "'"
    End If
End Function


Private Function PyDict(dict As Object) As String
    Dim key As Variant
    Dim s As String

    s = "{"
    For Each key In dict.Keys
        s = s & "'" & key & "': " & PyValue(dict(key)) & ", "
    Next key

    If Right(s, 2) = ", " Then s = Left(s, Len(s) - 2)
    PyDict = s & "}"
End Function



