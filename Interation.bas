Attribute VB_Name = "Interation"
Public Function choose_directory(Optional start_path As String = "") As String
    On Error GoTo ErrorProc
    
    Dim FD As FileDialog
    Dim selectedPath As String
    Dim open_dialog As Boolean
    
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    FD.Title = "Select Directory for Iterations"   ' <-- Add this line
    
    open_dialog = True
    selectedPath = ""
    
    While open_dialog
        With FD
            .AllowMultiSelect = False
            
            'Set initial directory if provided
            If start_path <> "" Then
                .InitialFileName = start_path
            End If
            
            If .Show = -1 Then
                selectedPath = .SelectedItems(1)
            Else
                selectedPath = ""
            End If
            
            'Reject SharePoint/URL paths
            If InStr(1, selectedPath, "http://", vbTextCompare) > 0 Or _
               InStr(1, selectedPath, "https://", vbTextCompare) > 0 Then
               
                MsgBox "Please select a local directory on a lettered drive (C:\). The selected path references a website, which happens with SharePoint locations."
                open_dialog = True
            Else
                open_dialog = False
            End If
        End With
    Wend
    
    choose_directory = selectedPath
    Exit Function

ErrorProc:
    MsgBox "Error in function choose_directory: " & Err.Description
    Err.Clear
    choose_directory = ""
End Function


Sub Write_Iteration_Files(wname As String)
    Dim save_path As String
    Dim start_path As String
    next_in_path = ThisWorkbook.FullName
    If InStr(1, next_in_path, "http://", vbTextCompare) > 0 Or _
        InStr(1, next_in_path, "https://", vbTextCompare) > 0 Then
        MsgBox "Please save Next-In on a lettered drive (c:\). The current file path references a website, which happens with files on sharepoint."
        Exit Sub
    End If
    start_path = ThisWorkbook.Path
    save_path = choose_directory(start_path)  'optional initial directory
    If save_path <> "" Then
        Call_NextOut wname, next_in_path, save_path
    Else
        MsgBox "No folder selected."
    End If
End Sub



