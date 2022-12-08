Attribute VB_Name = "Module2"
'Licensing procedures

Public Sub ActivationFromControl()
    On Error GoTo ErrorProc
        If Worksheets("Control").OptionButton4 Then
            If Module5.Local_Lock Then
                Sheet12.cmdRead.Enabled = True
                Sheet12.cmdReset.Enabled = True
                Sheet12.cmdWrite.Enabled = True
            Else
                Sheet12.cmdRead.Enabled = False
                Sheet12.cmdReset.Enabled = False
                Sheet12.cmdWrite.Enabled = False
            End If
        Exit Sub
        End If
        
        If Not IsGenuine Then
          Dim pkeyBox As frmPKey
          Set pkeyBox = New frmPKey
          Dim diagResult As VbMsgBoxResult
          diagResult = pkeyBox.ShowDialog(ta)
        Else
          Call AccountDeactivate
        End If
        Call ModValidKey.CheckActivation
        If Not CheckForValidKey Then Exit Sub
        Exit Sub
ErrorProc:
        MsgBox "Error in Procedure cmdActivation : " & Err.Description
        Err.Clear
End Sub

