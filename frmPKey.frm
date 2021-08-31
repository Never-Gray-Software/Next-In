VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPKey 
   Caption         =   "Activate now"
   ClientHeight    =   2085
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   5370
   OleObjectBlob   =   "frmPKey.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ta As TurboActivate
Private OKClicked As Boolean

Private Sub btnActivate_Click()
  ' tell your app to handle errors in the section
  ' at the end of the sub
  On Error GoTo TAError
  
  ' save the new key
  If Not GetTA().CheckAndSavePKey(txtPkey.Text, TA_SYSTEM) Then
    Err.Raise _
      Number:=1, _
      Description:="The product key is not valid.", _
      Source:="TurboActivate.CheckAndSavePKey"
  End If
  ' try to activate and close the form
  Call GetTA().Activate
  OKClicked = True
  Unload Me
SubExit:
  Exit Sub
TAError:
  MsgBox Err.Description
  Resume SubExit
End Sub

Private Sub btnCancel_Click()
  Unload Me
End Sub

Public Function ShowDialog(ByRef turboAct As TurboActivate) As VbMsgBoxResult
  Set ta = turboAct
  ' get the existing product key
  If GetTA().IsProductKeyValid Then
    txtPkey.Text = GetTA().GetPKey()
  End If
  Me.Show vbModal
  ShowDialog = IIf(OKClicked, vbOK, vbCancel)
  Unload Me
End Function

