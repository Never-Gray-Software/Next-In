Attribute VB_Name = "file_helper"
Sub get_savename(ByRef savename As String, ByRef save_file As Boolean, Optional Unit_name As String)
    Dim directory_path As String
    Dim file_selected As Variant
    Dim overwrite_exiting_file As VbMsgBoxResult
    Dim open_save_as_dialog As Boolean

    savename = ""
    save_file = False
    open_save_as_dialog = True
    While open_save_as_dialog
        open_save_as_dialog = False

        directory_path = Extract_Directory_Path(last_write_file_path.Value2)
        If directory_path <> "" And Dir(directory_path, vbDirectory) <> "" Then
            ChDir directory_path
        End If

        file_selected = Application.GetSaveAsFilename( _
            fileFilter:="SES Input File (*.inp), *.inp", _
            Title:="Save SES Input File")

        If file_selected = False Then
            save_file = False
        Else
            savename = CStr(file_selected)
            save_file = True

            If Dir(savename) <> "" Then
                overwrite_exiting_file = MsgBox("The file already exists. Overwrite?", _
                                                vbYesNoCancel + vbExclamation)

                Select Case overwrite_exiting_file
                    Case vbYes
                        save_file = True
                    Case vbNo
                        open_save_as_dialog = True
                        save_file = False
                    Case vbCancel
                        save_file = False
                End Select
            End If
        End If
    Wend
End Sub

' Creates a guaranteed local copy of the current workbook
Public Function GetLocalCopyPath(wname As String) As String
    Dim tempPath As String
    tempPath = Environ$("TEMP") & "\NextIn_LocalCopy.xlsx"

    ' Clean up any previous temp file
    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    ' Save a fresh local copy of the workbook being converted
    Workbooks(wname).SaveCopyAs tempPath
    
    GetLocalCopyPath = tempPath
End Function


Function Settings_File_Path(ByVal Original_Path As String) As String
    ' Replace all backslashes with forward slashes
    Settings_File_Path = Replace(Original_Path, "\", "/")
End Function

'Extract just the directory from a path that includes a file
Function Extract_Directory_Path(file_path As String) As String
    If file_path = "" Then
        Extract_Directory_Path = ""
    Else
        Extract_Directory_Path = Left(file_path, InStrRev(file_path, "\"))
    End If
End Function
' Clean up temporary files
Public Sub delete_temp_files(wname As String)
    Dim localCopy As String
    Dim tempFile As String
    Dim folder As String

    localCopy = GetLocalCopyPath(wname)
    folder = Extract_Directory_Path(localCopy)

    tempFile = folder & "\temporary_input_file_for_convert_in_excel.inp"
    If Dir(tempFile) <> "" Then Kill tempFile

    If Dir(localCopy) <> "" Then Kill localCopy
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call delete_temp_files(ThisWorkbook.Name)
End Sub

Public Function is_version_ip() As Boolean
    If si_ip_option.Value2 = 2 Then
        is_version_ip = True
    Else
        is_version_ip = False
    End If
End Function
