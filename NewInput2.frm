VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewInput2 
   Caption         =   "Create New Input"
   ClientHeight    =   5640
   ClientLeft      =   90
   ClientTop       =   315
   ClientWidth     =   7155
   OleObjectBlob   =   "NewInput2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewInput2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub CancelButton2_Click()
    NewInput.Hide
    MsgBox "PHEW! That was a close one. Be careful and be sure to save your files"
End Sub

Private Sub ContinueButton2_Click()
    Dim workbookname As String
    Dim ipversion As Boolean
    Dim wname_local As String
    ipversion = NewInput2.SES4p1_Check2.value
    NewInput2.Hide
    workbookname = ActiveWorkbook.Name
    WriteForm.Show vbModeless
    Call Speedon(True)
    Call ClearForms(workbookname)
    Call Formulas(workbookname)
    Call ip_switch(workbookname, ipversion, "")
    last_read_time.Value2 = "New File"
    last_read_version.Value2 = ""
    last_read_file.Value2 = ""
    If ipversion Then
        si_ip_cell.Value2 = "2 is the value"
    Else
        si_ip_cell.Value2 = "1 is the value"
    End If
    WriteForm.Hide
    Call Speedon(False)
End Sub

Private Sub UserForm_Initialize()
    ' Set the desired position (e.g., top-left corner of the primary monitor)
    Me.Left = 100 ' X position
    Me.Top = 100 ' Y position
End Sub
