Attribute VB_Name = "Button_Actions"
' Project Name: Next-In
' Description: Connects buttons and information on Control Sheet to Macros.
' Copyright (c) 2025 Justin Edenbaum, Never Gray
' This file is licensed under the MIT License.
' You may obtain a copy of the license at https://opensource.org/licenses/MIT

Option Explicit

Public si_ip_option As Variant
Public conversion As Variant
Public si_ip_cell As Range
Public output_conversion_string As String
Public output_conversion_option As Variant
Public Write_Option As Integer
Public SES_Exe As Range
Public NextOut_Exe As Range
Public Visio_File As Range
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

' Set location of information to read and write on Control Sheet
Sub Get_Control_Values(wname)
    Set si_ip_cell = Workbooks(wname).Worksheets("Control").Range("B2")
    si_ip_option = si_ip_cell.Value2
    Set conversion = Workbooks(wname).Worksheets("Control").Range("B6")
    conversion = conversion.Value2
    Set output_conversion_option = Workbooks(wname).Worksheets("Control").Range("C6").Value2
    Set SES_Exe = Workbooks(wname).Worksheets("Control").Range("F8")
    Set NextOut_Exe = Workbooks(wname).Worksheets("Control").Range("F9")
    Set Visio_File = Workbooks(wname).Worksheets("Control").Range("F10")
    Set summary_numbers = Workbooks(wname).Worksheets("Control").Range("F12")
    Set last_read_file_name = Workbooks(wname).Worksheets("Control").Range("C17")
    Set last_read_version = Workbooks(wname).Worksheets("Control").Range("F17")
    Set last_read_date = Workbooks(wname).Worksheets("Control").Range("G17")
    Set last_read_time = Workbooks(wname).Worksheets("Control").Range("H17")
    Set last_read_file_path = Workbooks(wname).Worksheets("Control").Range("I17")
    Set last_write_file_name = Workbooks(wname).Worksheets("Control").Range("C18")
    Set last_write_version = Workbooks(wname).Worksheets("Control").Range("F18")
    Set last_write_date = Workbooks(wname).Worksheets("Control").Range("G18")
    Set last_write_time = Workbooks(wname).Worksheets("Control").Range("H18")
    Set last_write_file_path = Workbooks(wname).Worksheets("Control").Range("I18")
    Set last_used_by = Workbooks(wname).Worksheets("Control").Range("C20")
    Select Case output_conversion_option
        Case 1
            output_conversion_string = ""
        Case 2
            output_conversion_string = "IP to SI"
        Case 3
            output_conversion_string = "SI to IP"
    End Select
End Sub

'Extract just the directory from a path that includes a file
Function Extract_Directory_Path(file_path As String) As String
    ' Check if the input is valid
    If file_path = "" Then
        Extract_Directory_Path = ""
    Else
        ' Extract the directory path
        Extract_Directory_Path = Left(file_path, InStrRev(file_path, "\"))
    End If
End Function

'Create a new, empty Next-In
Sub new_button()
    Dim wname As String
    wname = ActiveWorkbook.Name
    Get_Control_Values (wname)
    NewInput.SES4p1_Check1.value = is_version_ip(wname)
    NewInput.Show
End Sub

Sub rest_button()
    Call Speedon(False) ' Speed on is false
End Sub

Sub read_button()
    Dim wname As String
    wname = ActiveWorkbook.Name
    Get_Control_Values (wname)
    Call ReadFile
End Sub

Sub write_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Write_Option = 1
    Get_Control_Values (wname)
    Call select_creation_method(wname)
End Sub

Sub run_SES_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Write_Option = 2
    Get_Control_Values (wname)
    Call select_creation_method(wname)
End Sub
Sub run_next_out_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Write_Option = 3
    Get_Control_Values (wname)
    Call select_creation_method(wname)
End Sub

Sub Select_Exe_button()
    Dim wname As String
    Dim program_name As String
    Dim file_path As String
    Dim directory_path As String
    wname = ActiveWorkbook.Name
    Get_Control_Values (wname)
    directory_path = Extract_Directory_Path(SES_Exe.Value2)
    program_name = "SES"
    choose_exe wname, program_name, directory_path
End Sub

Sub Select_NextOut_button()
    Dim wname As String
    Dim program_name As String
    Dim directory_path As String
    wname = ActiveWorkbook.Name
    Get_Control_Values (wname)
    directory_path = Extract_Directory_Path(NextOut_Exe.Value2)
    program_name = "NextOut"
    choose_exe wname, program_name, directory_path
End Sub

Sub Select_visio_button()
    Dim wname As String
    Dim program_name As String
    Dim directory_path As String
    wname = ActiveWorkbook.Name
    Get_Control_Values (wname)
    directory_path = Extract_Directory_Path(Visio_File.Value2)
    program_name = "Visio"
    choose_exe wname, program_name, directory_path
End Sub

Sub Write_Iteration_Files_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Get_Control_Values (wname)
    Call Write_Iteration_Files(wname)
End Sub

Sub select_creation_method(wname As String)
    Dim ses_version As String
    Dim next_in_path As String
    If conversion = 1 Then ' Write using VBA script
        Call WriteFile
    Else
        Call Next_out_conversion(wname)
    End If
End Sub

Sub Next_out_conversion(wname As String)
    Dim ses_version As String
    Dim next_in_path As String
    
    ' Determine conversion direction
    If conversion = 2 And si_ip_option = 1 Then
        ses_version = "SI_2_IP"
    ElseIf conversion = 3 And si_ip_option = 2 Then
        ses_version = "IP_2_SI"
    Else
        ses_version = "Conflict"
        MsgBox "Error! Input and conversion settings don't match."
        Exit Sub
    End If
    
    ' Ask user for output file name
    Dim savename As String
    Dim save_file As Boolean
    Call get_savename(savename, save_file)
    
    If Not save_file Then Exit Sub
    
    ' Always create a guaranteed local copy of Next-In
    next_in_path = GetLocalCopyPath()
    
    ' Run Next-Out using the local copy
    Call_NextOut _
        workbook_name:=wname, _
        savename:=savename, _
        next_in_path:=next_in_path, _
        ses_version:=ses_version, _
        file_type:="next_in"

    ' If this was triggered by run_SES_button, also run SES directly
    If Write_Option = 2 Then
        WriteForm.TextBox2.value = "Running SES Simulation"
        WriteForm.Repaint
        Call_SES_Exe wname, savename

    ' If this was triggered by run_next_out_button, let Next-Out handle SES + post-processing
    ElseIf Write_Option = 3 Then
        WriteForm.TextBox2.value = "Running Next-Out, then SES"
        WriteForm.Repaint
        If ses_version = "IP_2_SI" Then
            ses_version = "SI"
        Else
            ses_version = "IP"
        End If
        Call_NextOut wname, savename, ses_version:=ses_version, file_type:="input_file"
        ' TODO need to fix "ses_version" to work properly
    End If
End Sub


