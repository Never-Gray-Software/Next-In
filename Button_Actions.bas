Attribute VB_Name = "Button_Actions"
' Project Name: Next-In
' Description: Connects buttons and information on Control Sheet to Macros.
' Copyright (c) 2025 Justin Edenbaum, Never Gray
' This file is licensed under the MIT License.
' You may obtain a copy of the license at https://opensource.org/licenses/MIT

Option Explicit

Public si_ip_option As Range
Public conversion As Range
Public output_conversion_option As Variant
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
    Set conversion = ctl.Range("B6")

    ' --- Values ---
    output_conversion_option = CV(ctl, "C6")

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
    NewInput.SES4p1_Check1.value = is_version_ip(wname)
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

Sub run_SES_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Write_Option = 2
    Get_Control_Values wname
    Call select_creation_method(wname)
End Sub

Sub run_next_out_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Write_Option = 3
    Get_Control_Values wname
    Call select_creation_method(wname)
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

Sub select_creation_method(wname As String)
    Dim ses_version As String
    Dim next_in_path As String
    If conversion.Value2 = 1 Then
        Call WriteFile
    Else
        Call Next_out_conversion(wname)
    End If
End Sub

Sub Next_out_conversion(wname As String)
    Dim ses_version As String
    Dim next_in_path As String
    
    If conversion.Value2 = 2 And si_ip_option.Value2 = 1 Then
        ses_version = "SI_TO_IP"
    ElseIf conversion.Value2 = 3 And si_ip_option.Value2 = 2 Then
        ses_version = "IP_TO_SI"
    Else
        ses_version = "Conflict"
        MsgBox "Error! Input and conversion settings don't match."
        Exit Sub
    End If
    
    Dim savename As String
    Dim save_file As Boolean
    Call get_savename(savename, save_file)
    
    If Not save_file Then Exit Sub
    
    next_in_path = GetLocalCopyPath()
    
    Call_NextOut _
        workbook_name:=wname, _
        savename:=savename, _
        next_in_path:=next_in_path, _
        ses_version:=ses_version, _
        file_type:="next_in"

    If Write_Option = 2 Then
        WriteForm.TextBox2.value = "Running SES Simulation"
        WriteForm.Repaint
        Call_SES_Exe wname, savename

    ElseIf Write_Option = 3 Then
        WriteForm.TextBox2.value = "Running Next-Out, then SES"
        WriteForm.Repaint

        If ses_version = "IP_TO_SI" Then
            ses_version = "SI"
        Else
            ses_version = "IP"
        End If

        Call_NextOut wname, savename, ses_version:=ses_version, file_type:="input_file"

    Else
        MsgBox "Converted file created:" & vbCrLf & vbCrLf & savename, vbInformation, "Next-Out Complete"
    End If
End Sub


