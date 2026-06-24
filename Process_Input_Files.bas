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

Public Sub Call_NextOut(workbook_name As String, savename As Variant, Optional iteration_path As String = "", Optional ses_version As String = "SI")
    On Error GoTo ErrorProc
    If iteration_path = "" Then
        WriteForm.TextBox2.value = "Attempting to run Next-Out, then SES"
    Else
        WriteForm.TextBox2.value = "Writing Iterations with Next-Out"
    End If
    WriteForm.Repaint
    If Dir(NextOut_Exe) = "" Then
        MsgBox "Next Out executable not found at: " & NextOut_Exe
        Exit Sub
    End If
    Dim shell_command As String, output_setting As String
    Dim Path_of_Next_Out As String
    Dim argument As String, NextOut_Path As String, msg As String
    Dim settings As Object
    Dim key As Variant
    Dim Proper_Path As String
    ' Path to your compiled PyInstaller .exe file
    ' Optional: Any command-line arguments you want to pass to the program
    ' <VARIABLES> in the argument statement are replaced below
    ' ERASE ME  --settings "{'file_type': 'input_file', 'output': [' ', 'H5_file', 'Visio'], 'path_exe': '-30', 'ses_output_str': ['<SES_OUTPUT_STR>'], 'simtime': -1, 'visio_template': '0', 'save_path': 'C:\temp', 'ses_version': 'SI'}"
    argument = " --settings ""{" & _
           "'conversion': '', " & _
           "'file_type': '<FILE_TYPE>', " & _
           "'output': [<OUTPUT_SETTING>], " & _
           "'path_exe': '<SES_EXE>', " & _
           "'ses_output_str': ['<SES_OUTPUT_STR>'], " & _
           "'simtime': -1, " & _
           "'visio_template': '<VISIO_FILE>', " & _
           "'iteration_path': '<ITERATION_PATH>', " & _
           "'ses_version': '<SES_VERSION>'" & _
           "}"""
    ' Construct the command to open cmd and run the program
    output_setting = Get_Output_Setting(workbook_name)
    Debug.Print output_setting
    Set settings = CreateObject("Scripting.Dictionary")
    settings.Add "<OUTPUT_SETTING>", CStr(output_setting)
    settings.Add "<SES_EXE>", Settings_File_Path(SES_Exe)
    settings.Add "<SES_OUTPUT_STR>", Settings_File_Path(savename)
    settings.Add "<VISIO_FILE>", Settings_File_Path(Visio_File)
    If iteration_path = "" Then 'Creating one input file and post-processing with Next-Out
        settings.Add "<FILE_TYPE>", "input_file"
    Else 'Next-Out will create iterations files in the save_path
        settings.Add "<FILE_TYPE>", "iteration"
        settings.Add "<ITERATION_PATH>", Settings_File_Path(iteration_path)
        settings.Add "<SES_VERSION>", get_ses_version()
    End If
    For Each key In settings.Keys
        'Debug.Print "Replacing " & key & " with " & settings(key)
        argument = Replace(argument, key, settings(key))
    Next key
    Debug.Print argument
    NextOut_Path = CStr(NextOut_Exe)
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
    If Workbooks(workbook_name).Worksheets("Control").Shapes("NO_Route_Data").ControlFormat.value = xlOn Then
        output_options.Add "Route"
    End If
    If Workbooks(workbook_name).Worksheets("Control").Shapes("NO_H5_File").ControlFormat.value = xlOn Then
        output_options.Add "H5_file"
    End If
    If Workbooks(workbook_name).Worksheets("Control").Shapes("NO_Visio").ControlFormat.value = xlOn Then
        output_options.Add "Visio"
        ' Add additional visio options if NO_visio is selected.
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

Function get_ses_version()
    If si_ip_option = 1 Then
        get_ses_version = "SI"
    Else
        get_ses_version = "IP"
    End If
End Function
