Attribute VB_Name = "unit_tests"
'Copyright 2024, Never Gray, Justin Edenbaum P.Eng
'Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
'1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
'2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
'3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS “AS IS” AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'Unit test to make sure software works correctly
'Create a folder with input files (IP or SI, not both).
'The unit test reads in all files and writes them out with a suffix of *.nxi.
'You can then compare the original file with the one created with Next-In

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
    MsgBox "Finished Unit Test", vbInformation, "Information"
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
            ReadFile (Fil.Path)
            num = Len(Fil.Path)
            unit_output = Left(Fil.Path, num - 4) + ".nxi"
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
                ReadFile (Fil.Path)
                num = Len(Fil.Path)
                unit_output = Left(Fil.Path, num - 4) + ".nxi"
                WriteFile (unit_output)
            Next Fil
            ' make recursive call, if subfolder contains subfolders
            If Not SubFolder.SubFolders Is Nothing Then _
               Call FindFilesInFolders(HostFolder & "\" & SubFolder.Name, FileTypes)
    
        Next SubFolder
    End If
    blNotFirstIteration = False
End Sub


