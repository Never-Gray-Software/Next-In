Attribute VB_Name = "decimal_conversions"
' Project Name: Next-In
' Description: Formats numbers depending on decimal separator being a period or comma (. or ,)
' Copyright (c) 2025 Justin Edenbaum, Never Gray
' This file is licensed under the MIT License.
' You may obtain a copy of the license at https://opensource.org/licenses/MIT

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
