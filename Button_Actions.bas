Attribute VB_Name = "Button_Actions"
' Project Name: Next-In
' Description: Connects buttons and information on Control Sheet to Macros.
' Copyright (c) 2025 Justin Edenbaum, Never Gray
' This file is licensed under the MIT License.
' You may obtain a copy of the license at https://opensource.org/licenses/MIT

Option Explicit

Public si_ip_option As Range
Public input_conversion As Range
Public output_conversion_option As Range
Public output_conversion_string As String
Public Write_Option As Integer
Public summary_numbers As Range
Public last_read_file_name  As Range
Public last_read_version    As Range
Public last_read_date       As Range
Public last_read_time       As Range
Public last_read_file_path  As Range
Public last_write_file_name As Range
Public last_write_version   As Range
Public last_write_date      As Range
Public last_write_time      As Range
Public last_write_file_path As Range
Public last_used_by         As Range
Public ses_version As String

' IMPORTANT: these must be Range, not String
Public SES_Exe As Range
Public NextOut_Exe As Range
Public Visio_File As Range

Private Function CV(ws As Worksheet, addr As String) As Variant
    CV = ws.Range(addr).Value2
End Function

Sub Get_Control_Values(wname As String)

    Dim ctl As Worksheet
    Set ctl = Workbooks(wname).Worksheets("Control")

    ' --- Ranges that you modify later ---
    Set si_ip_option = ctl.Range("B2")
    Set input_conversion = ctl.Range("B6")
    Set output_conversion_option = ctl.Range("C6")

    ' --- Ranges for executable / template paths ---
    Set SES_Exe = ctl.Range("F8")
    Set NextOut_Exe = ctl.Range("F9")
    Set Visio_File = ctl.Range("F10")

    ' --- Ranges ---
    Set summary_numbers = ctl.Range("F12")

    Set last_read_file_name = ctl.Range("C17")
    Set last_read_version = ctl.Range("F17")
    Set last_read_date = ctl.Range("G17")
    Set last_read_time = ctl.Range("H17")
    Set last_read_file_path = ctl.Range("I17")

    Set last_write_file_name = ctl.Range("C18")
    Set last_write_version = ctl.Range("F18")
    Set last_write_date = ctl.Range("G18")
    Set last_write_time = ctl.Range("H18")
    Set last_write_file_path = ctl.Range("I18")

    Set last_used_by = ctl.Range("C20")

    ' --- Output conversion string ---
    Select Case output_conversion_option
        Case 1: output_conversion_string = ""
        Case 2: output_conversion_string = "IP_TO_SI"
        Case 3: output_conversion_string = "SI_TO_IP"
    End Select

End Sub

'Create a new, empty Next-In
Sub new_button()
    Dim wname As String
    wname = ActiveWorkbook.Name
    Get_Control_Values wname
    NewInput.SES4p1_Check1.value = is_version_ip()
    NewInput.Show
End Sub

Sub rest_button()
    Call Speedon(False)
End Sub

Sub read_button()
    Dim wname As String
    wname = ActiveWorkbook.Name
    Get_Control_Values wname
    Call ReadFile
End Sub

Sub write_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Write_Option = 1
    Get_Control_Values wname
    Call select_creation_method(wname)
End Sub

Sub convert_values_in_excel_button()
    Dim wname As String
    wname = ActiveWorkbook.Name
    Get_Control_Values wname
    cell_convert.Show
End Sub


Sub run_SES_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Write_Option = 2
    Get_Control_Values wname
    Call select_creation_method(wname)
End Sub

Sub run_next_out_button()
    On Error GoTo NextOutError

    Dim wname As String
    wname = ActiveWorkbook.Name

    Write_Option = 3
    Get_Control_Values wname

    ' Attempt the conversion
    Call select_creation_method(wname)

    Exit Sub

NextOutError:
    MsgBox "Something went wrong calling Next-Out." & vbCrLf & vbCrLf & _
           "Error details:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "If the problem persists, write the input file and run Next-Out manually.", _
           vbCritical, "Next-Out Error"
End Sub


Sub Select_Exe_button()
    Dim wname As String
    Dim program_name As String
    Dim directory_path As String
    wname = ActiveWorkbook.Name
    Get_Control_Values wname
    directory_path = Extract_Directory_Path(CStr(SES_Exe.Value2))
    program_name = "SES"
    choose_exe wname, program_name, directory_path
End Sub

Sub Select_NextOut_button()
    Dim wname As String
    Dim program_name As String
    Dim directory_path As String
    wname = ActiveWorkbook.Name
    Get_Control_Values wname
    directory_path = Extract_Directory_Path(CStr(NextOut_Exe.Value2))
    program_name = "NextOut"
    choose_exe wname, program_name, directory_path
End Sub

Sub Select_visio_button()
    Dim wname As String
    Dim program_name As String
    Dim directory_path As String
    wname = ActiveWorkbook.Name
    Get_Control_Values wname
    directory_path = Extract_Directory_Path(CStr(Visio_File.Value2))
    program_name = "Visio"
    choose_exe wname, program_name, directory_path
End Sub

Sub Write_Iteration_Files_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Get_Control_Values wname
    Call Write_Iteration_Files(wname)
End Sub

Sub Write_Iteration_and_run_next_out()
    MsgBox "Use Write Iteration Files and call Next-Out seperately. Future versions will enable running simulations directly from Next-In."
End Sub

Sub select_creation_method(wname As String, Optional unit_name As String)
    Dim ses_version As String
    Dim next_in_path As String
    If Len(unit_name) = 0 Then unit_name = ""
    If input_conversion.Value2 = 1 Then
        Call WriteFile(unit_name)
    Else
        Call Next_out_conversion(wname, unit_name)
    End If
End Sub
