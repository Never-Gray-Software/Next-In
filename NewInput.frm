VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewInput 
   Caption         =   "Create New Input"
   ClientHeight    =   5640
   ClientLeft      =   192
   ClientTop       =   732
   ClientWidth     =   7176
   OleObjectBlob   =   "NewInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton1_Click()
    NewInput.Hide
    MsgBox "PHEW! That was a close one. Be careful and be sure to save your files"
End Sub

Private Sub ContinueButton1_Click()
    NewInput2.SES4p1_Check2.value = NewInput.SES4p1_Check1.value
    NewInput.Hide
    NewInput2.Show
End Sub

Private Sub UserForm_Initialize()
    ' Set the desired position (e.g., top-left corner of the primary monitor)
    Me.Left = 100 ' X position
    Me.Top = 100 ' Y position
End Sub
