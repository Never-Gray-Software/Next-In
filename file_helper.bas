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
Public Function GetLocalCopyPath() As String
    Dim tempPath As String
    tempPath = Environ$("TEMP") & "\NextIn_LocalCopy.xlsx"

    ' Clean up any previous temp file
    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    ' Save a fresh local copy
    ThisWorkbook.SaveCopyAs tempPath

    GetLocalCopyPath = tempPath
End Function

' Deletes the temporary local copy of Next-In, if it exists
Public Sub CleanupLocalCopies()
    Dim tempPath As String
    tempPath = Environ$("TEMP") & "\NextIn_LocalCopy.xlsx"
    
    On Error Resume Next
    Kill tempPath
    On Error GoTo 0
End Sub
