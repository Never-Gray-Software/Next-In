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
    'Get_Control_Values (workbook_name)
    path_exe = Range(SES_Exe.Address).Value2
    If path_exe <> "" Then
        shell_command = """" & path_exe & """ """ & input_file_path & """"
        Debug.Print shell_command
        Shell shell_command, vbNormalNoFocus  'Previously vbNormalFocus
    End If
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Call_SES_Exe: " & Err.Description
    Err.Clear
End Sub

Public Sub Call_NextOut(workbook_name As String, savename)
    On Error GoTo ErrorProc
    WriteForm.TextBox2.value = "Attempting to run SES and Next-Out"
    WriteForm.Repaint
    Dim shell_command As String, output_setting As String
    Dim Path_of_Next_Out As String
    Dim argument As String, NextOut_Path As String, msg As String
    Dim settings As Object
    Dim key As Variant
    Dim Proper_Path As String
    'Get_Control_Values (workbook_name)
    ' Path to your compiled PyInstaller .exe file
    ' Optional: Any command-line arguments you want to pass to the program
    ' <VARIABLES> in the argument statement are replaced below
    argument = " --settings ""{'conversion': '', 'file_type': 'input_file', 'output': [<OUTPUT_SETTING>], 'path_exe': '<SES_EXE>', 'results_folder_str': None, 'ses_output_str': ['<INPUT_FILE>'], 'simtime': -1, 'visio_template': '<VISIO_FILE>'}"""
    ' Construct the command to open cmd and run the program
    output_setting = Get_Output_Setting(workbook_name)
    Debug.Print output_setting
    Set settings = CreateObject("Scripting.Dictionary")
    settings.Add "<OUTPUT_SETTING>", CStr(output_setting)
    settings.Add "<INPUT_FILE>", Settings_File_Path(savename) 'Get proper format for argument
    settings.Add "<SES_EXE>", Settings_File_Path(Range(SES_Exe.Address).Value2)
    settings.Add "<VISIO_FILE>", Settings_File_Path(Range(Visio_File.Address).Value2)
    For Each key In settings.Keys
        'Debug.Print "Replacing " & key & " with " & settings(key)
        argument = Replace(argument, key, settings(key))
    Next key
    NextOut_Path = CStr(Range(NextOut_Exe.Address).Value2)
    shell_command = """" & NextOut_Path & """" & argument
    Debug.Print shell_command
    Shell shell_command, vbNormalNoFocus
    WriteForm.TextBox2.value = "Running SES and Next-Out"
    WriteForm.Repaint
    Exit Sub
ErrorProc:
    MsgBox "Error in procedure Call_NextOut: " & Err.Description
    Err.Clear
End Sub

Function Get_Output_Setting(workbook_name As String) As String
    On Error GoTo ErrorProc
    Dim str As String
    Dim output_options As Collection
    Set output_options = New Collection
    Dim Item As Variant
    If Workbooks(workbook_name).Worksheets("Control").Shapes("NO_Excel").ControlFormat.value = xlOn Then
        output_options.Add "Excel"
    End If
    If Workbooks(workbook_name).Worksheets("Control").Shapes("NO_Visio").ControlFormat.value = xlOn Then
        output_options.Add "Visio"
    End If
    If Workbooks(workbook_name).Worksheets("Control").Shapes("NO_Route_Data").ControlFormat.value = xlOn Then
        output_options.Add "Route"
    End If
    If Workbooks(workbook_name).Worksheets("Control").Shapes("NO_PDF").ControlFormat.value = xlOn Then
        output_options.Add "visio_2_pdf"
    End If
    If Workbooks(workbook_name).Worksheets("Control").Shapes("NO_PNG").ControlFormat.value = xlOn Then
        output_options.Add "visio_2_png"
    End If
    If Workbooks(workbook_name).Worksheets("Control").Shapes("NO_SVG").ControlFormat.value = xlOn Then
        output_options.Add "visio_2_svg"
    End If
    If Workbooks(workbook_name).Worksheets("Control").Shapes("NO_Open_Visio").ControlFormat.value = xlOn Then
        output_options.Add "visio_open"
    End If
    str = "' '"
    For Each Item In output_options
        str = str & ", '" & Item & "'"
    Next Item
    Get_Output_Setting = str
    Exit Function
ErrorProc:
    MsgBox "Error in procedure Get_Output_Settings: " & Err.Description
    Err.Clear
End Function

Function Settings_File_Path(ByVal Original_Path As String) As String
    ' Replace all backslashes with forward slashes
     Settings_File_Path = Replace(Original_Path, "\", "/")
End Function
Sub Pulsante22_Click()

End Sub
Sub Pulsante23_Click()

End Sub
