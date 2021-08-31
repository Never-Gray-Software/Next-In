Attribute VB_Name = "Module5"
Option Compare Binary
Public Const MIN_ASC As Integer = 32
Public Const MAX_ASC As Integer = 126
Public Const NO_OF_CHARS As Integer = MAX_ASC - MIN_ASC + 1
Private Const keymasterkey As String = "JustinNeverGray2021KennedydkjafjdfaPLU380205VNKLS2"
Private Const gatekey As String = "Heather#Hidde$Square#l1923fRad9380gFlnew#%^%F3vs[]"

Sub Keymaster_maker()
    Dim str As String
    str = Pull
    b = encrypt(str, keymasterkey)
    t_s = "Send the Keymaster below to Justin@NeverGray.biz with subject line: Keymaster"
    t_e = "End of Message"
    tt = t_s + Chr(13) + Chr(13) + b + Chr(13) + Chr(13) + t_e
    Clipboard (tt)
    'c = decrypt(b, key)
    MsgBox ("Local Key Copied." + Chr(13) + "Open email editor and paste text (Ctrl + V) in the body of email." + Chr(13) + "Follow instructions.")
    Worksheets("Control").Range("N2").Value2 = "Gatekeeper:"
    Worksheets("Control").Range("O2").Interior.Pattern = xlNone
End Sub

'Compares Gatekeeper to internal pull
Function Local_Lock() As Boolean
    On Error GoTo ErrorProc
    Local_Lock = False
    gatekeeper = Worksheets("Control").Range("O2").Value2
    If Len(gatekeeper) = 0 Then
        Keymaster_maker
        Exit Function
    End If
    gate_key = decrypt(gatekeeper, gatekey)
    ll = Pull
    If gate_key = ll Then
        Local_Lock = True
    Else
        Local_Lock = False
        MsgBox ("Gatekeeper is invalid")
        Keymaster_maker
    End If
    Exit Function
ErrorProc:
    MsgBox "Error in function CheckForValidKey : " & Err.Description
    Err.Clear
End Function

Function MoveAsc(ByVal a, ByVal mLvl)
    'Move the Asc value so it stays inside interval MIN_ASC and MAX_ASC
    mLvl = mLvl Mod NO_OF_CHARS
    a = a + mLvl
    If a < MIN_ASC Then
        a = a + NO_OF_CHARS
    ElseIf a > MAX_ASC Then
        a = a - NO_OF_CHARS
    End If
    MoveAsc = a
End Function
Function encrypt(ByVal s As String, ByVal key As String)
    Dim p, keyPos, c, e, k, chkSum
    If key = "" Then
        encrypt = s
        Exit Function
    End If
    For p = 1 To Len(s) 'check if character is valid in ASCII
        If Asc(Mid(s, p, 1)) < MIN_ASC Or Asc(Mid(s, p, 1)) > MAX_ASC Then
            MsgBox "Char at position " & p & " is invalid!"
            Exit Function
        End If
    Next p
    For keyPos = 1 To Len(key) 'Calculate a checksum
        chkSum = chkSum + Asc(Mid(key, keyPos, 1)) * keyPos
    Next keyPos
    keyPos = 0
    For p = 1 To Len(s) 'front beginning to end of string
        c = Asc(Mid(s, p, 1)) 'go letter by letter
        keyPos = keyPos + 1
        If keyPos > Len(key) Then keyPos = 1
        k = Asc(Mid(key, keyPos, 1))
        c = MoveAsc(c, k)
        c = MoveAsc(c, k * Len(key))
        c = MoveAsc(c, chkSum * k)
        c = MoveAsc(c, p * k)
        c = MoveAsc(c, Len(s) * p) 'This is only for getting new chars for different word lengths
        e = e & Chr(c)
    Next p
    encrypt = e
End Function
Function decrypt(ByVal s As String, ByVal key As String)
    Dim p, keyPos, c, d, k, chkSum
    If key = "" Then
        decrypt = s
        Exit Function
    End If
    For keyPos = 1 To Len(key)
        chkSum = chkSum + Asc(Mid(key, keyPos, 1)) * keyPos
    Next keyPos
    keyPos = 0
    For p = 1 To Len(s)
        c = Asc(Mid(s, p, 1))
        keyPos = keyPos + 1
        If keyPos > Len(key) Then keyPos = 1
        k = Asc(Mid(key, keyPos, 1))
        'Do MoveAsc in reverse order from encrypt, and with a minus sign this time(to unmove)
        c = MoveAsc(c, -(Len(s) * p))
        c = MoveAsc(c, -(p * k))
        c = MoveAsc(c, -(chkSum * k))
        c = MoveAsc(c, -(k * Len(key)))
        c = MoveAsc(c, -k)
        d = d & Chr(c)
    Next p
    decrypt = d
End Function
Function Pull() As String
    Dim key_pull(0 To 8) As String 'Key array
    Dim part1, part2, pk As String
    Dim str As String
    key_pull(0) = Application.UserName
    key_pull(1) = Application.OperatingSystem
    key_pull(2) = Application.OrganizationName
    key_pull(3) = Application.ProductCode
    key_pull(4) = Environ("computername")
    key_pull(5) = Environ("username")
    key_pull(6) = Environ("PROCESSOR_IDENTIFIER")
    key_pull(7) = Environ("PROCESSOR_LEVEL")
    key_pull(8) = Environ("PROCESSOR_REVISION")
    For i = 0 To 8
        part1 = part1 + Format(Len(key_pull(i)), "#000")
        part2 = part2 + key_pull(i)
    Next i
    pk = part1 + part2
    Pull = pk
End Function

Function Clipboard(Optional StoreText As String) As String
'From https://www.thespreadsheetguru.com/blog/2015/1/13/how-to-use-vba-code-to-copy-text-to-the-clipboard
'PURPOSE: Read/Write to Clipboard
'Source: ExcelHero.com (Daniel Ferry)

Dim x As Variant

'Store as variant for 64-bit VBA support
  x = StoreText

'Create HTMLFile Object
  With CreateObject("htmlfile")
    With .parentWindow.clipboardData
      Select Case True
        Case Len(StoreText)
          'Write to the clipboard
            .setData "text", x
        Case Else
          'Read from the clipboard (no variable passed through)
            Clipboard = .GetData("text")
      End Select
    End With
  End With

End Function
