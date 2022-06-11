VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewInput2 
   Caption         =   "Create New Input"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4605
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
    ipversion = NewInput2.SES4p1_Check2.Value
    NewInput2.Hide
    workbookname = ActiveWorkbook.Name
    WriteForm.Show vbModeless
    Call Speedon(True)
    Call ClearForms(workbookname)
    Call Formulas(workbookname)
    Call ip_switch(workbookname, ipversion, "")
    Workbooks(workbookname).Worksheets("Control").ipCheckBox = ipversion
    Workbooks(workbookname).Worksheets("Control").Range("G19").Value2 = ""
    Workbooks(workbookname).Worksheets("Control").Range("H19").Value2 = "New File"
    WriteForm.Hide
    Call Speedon(False)
End Sub

