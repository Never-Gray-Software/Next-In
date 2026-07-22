Attribute VB_Name = "Unit_Testing"
Dim blNotFirstIteration As Boolean
Dim Fil As File
Dim hFolder As Folder, SubFolder As Folder
Dim FileExt As String
Dim FSO As Scripting.FileSystemObject
Dim unit_output As String

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
    Dim first_suffix As String
    Dim second_suffix As String
    first_suffix = "nxi"
    second_suffix = "UIN"
    i = 0
    If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
    Set hFolder = FSO.GetFolder(HostFolder)
    ' iterate through all files in the root of the main folder
    If Not blNotFirstIteration Then
      For Each Fil In hFolder.Files
        'cutomized code
        If Right(Fil.Path, 3) = first_suffix Then
            Call ReadFilesv2(2, Fil.Path)
            unit_name = Left(Fil.Path, InStrRev(Fil.Path, ".")) & second_suffix
            Call ReadFilesv2(12, unit_name)
            Call CompareArray3
            Call SavePDF(Fil.Path)
        End If
      Next Fil
      ' make recursive call, if main folder contains subfolder
      If Not hFolder.SubFolders Is Nothing Then
          blNotFirstIteration = True
          Call FindFilesInFolders(HostFolder, FileTypes)
      End If
    End If
    ' iterate through all files in all the subfolders of the main folder
    blNotFirstIteration = False
End Sub

