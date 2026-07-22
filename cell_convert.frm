VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cell_convert 
   Caption         =   "Create New Input"
   ClientHeight    =   7050
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8955.001
   OleObjectBlob   =   "cell_convert.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cell_convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CancelButton1_Click()
    cell_convert.Hide
    MsgBox "PHEW! That was a close one. Be careful and back-up your files"
End Sub

Private Sub ContinueButton_4_excel_conversion_Click()
    Call convert_in_excel
End Sub

Private Sub UserForm_Initialize()
    ' Set the desired position (e.g., top-left corner of the primary monitor)
    Me.Left = 100 ' X position
    Me.Top = 100 ' Y position
End Sub
