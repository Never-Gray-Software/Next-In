Attribute VB_Name = "input_file_conversion"
Option Explicit

Public Sub Next_out_conversion(wname As String, Optional unit_name As String)
    Dim ses_version As String
    Dim next_in_path As String
    Dim unit_test_in_progress As Boolean
    If Len(unit_name) = 0 Then unit_name = ""
    ' Determine conversion direction
    If input_conversion.Value2 = 2 And si_ip_option.Value2 = 1 Then
        ses_version = "SI_TO_IP"
    ElseIf input_conversion.Value2 = 3 And si_ip_option.Value2 = 2 Then
        ses_version = "IP_TO_SI"
    Else
        MsgBox "Error! Read In Units and Input Conversion don't match."
        Exit Sub
    End If
    
    ' Ask user for save name
    If unit_name = "" Then
        Dim savename As String
        Dim save_file As Boolean
        get_savename savename, save_file
        If Not save_file Then Exit Sub
        unit_test_in_progress = False
    Else:
        savename = unit_name
        unit_test_in_progress = True
    End If
    ' Create guaranteed local copy
    next_in_path = GetLocalCopyPath(wname)
    
    ' Run Next-Out conversion
    Call_NextOut _
        workbook_name:=wname, _
        savename:=savename, _
        next_in_path:=next_in_path, _
        ses_version:=ses_version, _
        file_type:="next_in", _
        unit_test_in_progress:=unit_test_in_progress

    ' Handle SES or Next-Out+SES options
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
        If unit_name = "" Then
            MsgBox "Converted file created:" & vbCrLf & vbCrLf & savename, vbInformation, "Next-Out Complete"
        End If
    End If
End Sub


Public Sub convert_in_excel()
    On Error GoTo ConvertError

    Dim wname As String
    Dim from_units As Long
    Dim to_units As Long
    Dim ses_version As String
    Dim next_in_path As String
    Dim temp_input_2_read_in As String
    Dim temp_folder As String
    Dim before_conversion_read_file_name As String
    Dim before_conversion_read_file_path As String

    wname = ActiveWorkbook.Name

    ' Refresh global ranges FIRST
    Get_Control_Values wname

    ' Current units of workbook
    from_units = si_ip_option.Value2

    ' Decide conversion direction
    Select Case input_conversion.Value2
        Case 2      ' SI ? IP
            ses_version = "SI_TO_IP"
            to_units = 2

        Case 3      ' IP ? SI
            ses_version = "IP_TO_SI"
            to_units = 1

        Case Else
            MsgBox "Error! Invalid Input Conversion option.", vbCritical
            Exit Sub
    End Select

    ' Create guaranteed local copy
    next_in_path = GetLocalCopyPath(wname)

    ' Build temporary file path
    temp_folder = Extract_Directory_Path(next_in_path)
    temp_input_2_read_in = temp_folder & "temporary_input_file_for_convert_in_excel.inp"

    ' Run Next-Out conversion
    Call_NextOut _
        workbook_name:=wname, _
        savename:=temp_input_2_read_in, _
        next_in_path:=next_in_path, _
        ses_version:=ses_version, _
        file_type:="next_in", _
        wait_for_finish:=True

    cell_convert.Hide
    
    
    ' Save old read-file info
    before_conversion_read_file_name = CStr(last_read_file_name.Value2)
    before_conversion_read_file_path = CStr(last_read_file_path.Value2)
    
    ' IMPORTANT:
    ' Read the converted file after changing unit system
    si_ip_option.Value2 = to_units
    Call ReadFile(temp_input_2_read_in)

    ' Reset conversion dropdown
    input_conversion.Value2 = 1

    ' Update read-file metadata
    last_read_file_name.Value2 = before_conversion_read_file_name & " " & ses_version
    last_read_file_path.Value2 = before_conversion_read_file_path

    Exit Sub

ConvertError:
    MsgBox "Error converting values in Excel: " & Err.Description, vbCritical
    cell_convert.Hide
End Sub





