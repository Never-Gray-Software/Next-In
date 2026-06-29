Attribute VB_Name = "Interation"
Public Function choose_directory(Optional start_path As String = "") As String
    On Error GoTo ErrorProc
    
    Dim FD As FileDialog
    Dim selectedPath As String
    Dim open_dialog As Boolean
    
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    FD.Title = "Select Directory for Iterations"
    FD.AllowMultiSelect = False
    
    open_dialog = True
    selectedPath = ""
    
    While open_dialog
        ' Set initial directory if provided
        If Len(start_path) > 0 Then
            FD.InitialFileName = start_path
        End If
        
        If FD.Show = -1 Then
            selectedPath = FD.SelectedItems(1)
        Else
            selectedPath = ""
        End If
        
        ' User cancelled
        If Len(selectedPath) = 0 Then
            open_dialog = False
        ' Reject SharePoint/URL paths
        ElseIf InStr(1, selectedPath, "http://", vbTextCompare) > 0 Or _
               InStr(1, selectedPath, "https://", vbTextCompare) > 0 Then
               
            MsgBox "Please select a local directory on a lettered drive (e.g., C:\). " & _
                   "The selected path is a web location (SharePoint/OneDrive URL)."
            open_dialog = True
        Else
            open_dialog = False
        End If
    Wend
    
    choose_directory = selectedPath
    Exit Function

ErrorProc:
    MsgBox "Error in function choose_directory: " & Err.Description
    choose_directory = ""
End Function


Sub Write_Iteration_Files(wname As String)
    Dim iterationPath As String
    Dim localCopyPath As String
    Dim initialFolder As String

    ' Create a guaranteed local copy of Next-In
    localCopyPath = GetLocalCopyPath(wname)

    ' Choose the folder where iteration files will be written
    initialFolder = ThisWorkbook.Path
    iterationPath = choose_directory(initialFolder)

    If Len(iterationPath) = 0 Then
        MsgBox "No folder selected."
        Exit Sub
    End If

    ' Execute Next-Out using the local copy
    Call_NextOut _
        workbook_name:=wname, _
        savename:="", _
        next_in_path:=localCopyPath, _
        iteration_path:=iterationPath
End Sub



