Attribute VB_Name = "Module6"
Dim blNotFirstIteration As Boolean
Dim Fil As File
Dim hFolder As Folder, SubFolder As Folder
Dim FileExt As String
Dim FSO As Scripting.FileSystemObject

Sub unit_test()
    Dim strFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
      If .Show = -1 Then strFolder = .SelectedItems(1)
    End With
    If strFolder <> "" Then
        Call FindFilesInFolders(strFolder, "inp")
    End If
End Sub

' From https://wellsr.com/vba/2018/excel/list-files-in-folder-and-subfolders-with-vba-filesystemobject/
' Variable declarations
' Recursive procedure for iterating through all files in all subfolders
' of a folder and locating specific file types by file extension.
Sub FindFilesInFolders(ByVal HostFolder As String, FileTypes As Variant)
'(1) This routine uses Early Binding so you must add reference to Microsoft Scripting Runtime:
' Tools > References > Microsoft Scripting Runtime
'(2) Call procedure using a command like:
' Call FindFilesInFolders("C:\Users\MHS\Documents", Array("xlsm", "xlsb"))
    i = 0
    If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
    Set hFolder = FSO.GetFolder(HostFolder)
    ' iterate through all files in the root of the main folder
    If Not blNotFirstIteration Then
      For Each Fil In hFolder.Files
        'cutomized code
        ReadFile (Fil.path)
        num = Len(Fil.path)
        unit_output = Left(Fil.path, num - 4) + ".nxi"
        WriteFile (unit_output)
      Next Fil
      ' make recursive call, if main folder contains subfolder
      If Not hFolder.SubFolders Is Nothing Then
          blNotFirstIteration = True
          Call FindFilesInFolders(HostFolder, FileTypes)
      End If
    
    ' iterate through all files in all the subfolders of the main folder
    Else
      For Each SubFolder In hFolder.SubFolders
          For Each Fil In SubFolder.Files
            'cutomized code
            ReadFile (Fil.path)
            num = Len(Fil.path)
            unit_output = Left(Fil.path, num - 4) + ".nxi"
            WriteFile (unit_output)
          Next Fil
          ' make recursive call, if subfolder contains subfolders
          If Not SubFolder.SubFolders Is Nothing Then _
              Call FindFilesInFolders(HostFolder & "\" & SubFolder.Name, FileTypes)
    
      Next SubFolder
    End If
    blNotFirstIteration = False
End Sub

