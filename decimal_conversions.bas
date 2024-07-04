Attribute VB_Name = "decimal_conversions"

Function Int_International(value As Variant) As Integer
    If VarType(value) = vbString Then
        If Int("3.00") <> 3 Then 'if the period is not the decimal separator
            value = Replace(value, ".", ",")
        End If
    End If
    If value = "" Then
        Int_International = Int(0)
    Else
        Int_International = Int(value)
    End If
End Function

Sub DetectDecimalSeparator()
    Dim currentSeparator As String
    currentSeparator = Application.DecimalSeparator
    
    MsgBox "The current decimal separator is '" & currentSeparator & "'."
End Sub
