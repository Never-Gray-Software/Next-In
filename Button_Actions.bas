Attribute VB_Name = "Button_Actions"
' Project Name: Next-In
' Description: Connects buttons and information on Control Sheet to Macros.
' Copyright (c) 2025 Justin Edenbaum, Never Gray
' This file is licensed under the MIT License.
' You may obtain a copy of the license at https://opensource.org/licenses/MIT

Option Explicit

Public si_ip_option As Variant
Public si_ip_cell As Range
Public Write_Option As Integer
Public SES_Exe As Range
Public NextOut_Exe As Range
Public Visio_File As Range
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
    Set SES_Exe = Workbooks(wname).Worksheets("Control").Range("F8")
    Set NextOut_Exe = Workbooks(wname).Worksheets("Control").Range("F9")
    Set Visio_File = Workbooks(wname).Worksheets("Control").Range("F10")
    Set last_read_file_name = Workbooks(wname).Worksheets("Control").Range("C15")
    Set last_read_version = Workbooks(wname).Worksheets("Control").Range("F15")
    Set last_read_date = Workbooks(wname).Worksheets("Control").Range("G15")
    Set last_read_time = Workbooks(wname).Worksheets("Control").Range("H15")
    Set last_read_file_path = Workbooks(wname).Worksheets("Control").Range("I15")
    Set last_write_file_name = Workbooks(wname).Worksheets("Control").Range("C16")
    Set last_write_version = Workbooks(wname).Worksheets("Control").Range("F16")
    Set last_write_date = Workbooks(wname).Worksheets("Control").Range("G16")
    Set last_write_time = Workbooks(wname).Worksheets("Control").Range("H16")
    Set last_write_file_path = Workbooks(wname).Worksheets("Control").Range("I16")
    Set last_used_by = Workbooks(wname).Worksheets("Control").Range("C18")
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
    Call WriteFile
End Sub

Sub run_SES_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Write_Option = 2
    Get_Control_Values (wname)
    Call WriteFile
End Sub
Sub run_next_out_button()
    Dim wname As String
    Dim Write_Options As Integer
    wname = ActiveWorkbook.Name
    Write_Option = 3
    Get_Control_Values (wname)
    Call WriteFile
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

