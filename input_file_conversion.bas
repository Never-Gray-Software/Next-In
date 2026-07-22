Attribute VB_Name = "input_file_conversion"
Option Explicit

Public Sub Next_out_conversion(wname As String, Optional unit_name As String)
    Dim ses_version As String
    Dim next_in_path As String
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
    Else: savename = unit_name
    End If
    ' Create guaranteed local copy
    next_in_path = GetLocalCopyPath(wname)
    
    ' Run Next-Out conversion
    Call_NextOut _
        workbook_name:=wname, _
        savename:=savename, _
        next_in_path:=next_in_path, _
        ses_version:=ses_version, _
        file_type:="next_in"

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

    wname = ActiveWorkbook.Name

    ' Current units of workbook
    from_units = si_ip_option.Value2

    ' Decide conversion direction
    Select Case input_conversion.Value2
        Case 2      ' SI ? IP
            If from_units <> 1 Then
                MsgBox "Error! Input Conversion must be 'IP to SI' for Read in Units of 'IP'.", vbCritical
                Exit Sub
            End If
            ses_version = "SI_TO_IP"
            to_units = 2

        Case 3      ' IP ? SI
            If from_units <> 2 Then
                MsgBox "Error! Input Conversion must be 'SI to IP' for Read In Units 'SI'.", vbCritical
                Exit Sub
            End If
            ses_version = "IP_TO_SI"
            to_units = 1

        Case Else
            MsgBox "Error! Invalid Input Conversion option.", vbCritical
            Exit Sub
    End Select

    ' Create guaranteed local copy
    next_in_path = GetLocalCopyPath(wname)
    If Dir(next_in_path) = "" Then
        MsgBox "Local Next-In file not found: " & next_in_path, vbCritical
        Exit Sub
    End If

    ' Build temporary file path
    temp_folder = Extract_Directory_Path(next_in_path)
    temp_input_2_read_in = temp_folder & "\temporary_input_file_for_convert_in_excel.inp"

    ' Run Next-Out conversion
    Call_NextOut _
        workbook_name:=wname, _
        savename:=temp_input_2_read_in, _
        next_in_path:=next_in_path, _
        ses_version:=ses_version, _
        file_type:="next_in"

    If Dir(temp_input_2_read_in) = "" Then
        MsgBox "Next-Out did not produce a converted file.", vbCritical
        Exit Sub
    End If

    cell_convert.Hide

    ' Update units BEFORE reading converted file
    si_ip_option.Value2 = to_units

    ' Copy previous read file name and path
    Dim before_conversion_read_file_name As String
    Dim before_conversion_read_file_path As String
    before_conversion_read_file_name = CStr(last_read_file_name.Value2)
    before_conversion_read_file_path = CStr(last_read_file_path.Value2)
    ' Read converted file
    Call ReadFile(temp_input_2_read_in)
    
    ' Update the name of read file and path because it was changed by ReadFile
    last_read_file_name.Value2 = before_conversion_read_file_name & " " & ses_version
    last_read_file_path.Value2 = before_conversion_read_file_path
    Exit Sub

ConvertError:
    MsgBox "Error converting values in Excel: " & Err.Description, vbCritical
    cell_convert.Hide
End Sub


