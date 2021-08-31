Attribute VB_Name = "TurboActivateHelper"
Option Explicit

Public Const TA_SYSTEM = 1
Public Const TA_USER = 2

' Use the TA_DISALLOW_VM in UseTrial() to disallow trials in virtual machines.
' If you use this flag in UseTrial() and the customer's machine is a Virtual
' Machine, then UseTrial() will return TA_E_IN_VM.
Public Const TA_DISALLOW_VM = 4

' Use this flag in TA_UseTrial() to tell TurboActivate to use client-side
' unverified trials. For more information about verified vs. unverified trials,
' see here: https://wyday.com/limelm/help/trials/
' Note: unverified trials are unsecured and can be reset by malicious customers.
' </summary>
Public Const TA_UNVERIFIED_TRIAL = 16

' Use the TA_VERIFIED_TRIAL flag to use verified trials instead
' of unverified trials. This means the trial is locked to a particular computer.
' The customer can't reset the trial.
' </summary>
Public Const TA_VERIFIED_TRIAL = 32

' Flags for IsDateValid
Public Const TA_HAS_NOT_EXPIRED = 1

Public Const TA_OK = &H0
Public Const TA_FAIL = &H1
Public Const TA_E_PKEY = &H2
Public Const TA_E_ACTIVATE = &H3
Public Const TA_E_INET = &H4
Public Const TA_E_INUSE = &H5
Public Const TA_E_REVOKED = &H6
Public Const TA_E_PDETS = &H8
Public Const TA_E_TRIAL = &H9
Public Const TA_E_COM = &HB
Public Const TA_E_TRIAL_EUSED = &HC
Public Const TA_E_TRIAL_EEXP = &HD
Public Const TA_E_EXPIRED = &HD
Public Const TA_E_INSUFFICIENT_BUFFER = &HE
Public Const TA_E_PERMISSION = &HF
Public Const TA_E_INVALID_FLAGS = &H10
Public Const TA_E_IN_VM = &H11
Public Const TA_E_EDATA_LONG = &H12
Public Const TA_E_INVALID_ARGS = &H13
Public Const TA_E_KEY_FOR_TURBOFLOAT = &H14
Public Const TA_E_INET_DELAYED = &H15
Public Const TA_E_FEATURES_CHANGED = &H16
Public Const TA_E_NO_MORE_DEACTIVATIONS = &H18
Public Const TA_E_ACCOUNT_CANCELED = &H19
Public Const TA_E_ALREADY_ACTIVATED = &H1A
Public Const TA_E_INVALID_HANDLE = &H1B
Public Const TA_E_ENABLE_NETWORK_ADAPTERS = &H1C
Public Const TA_E_ALREADY_VERIFIED_TRIAL = &H1D
Public Const TA_E_TRIAL_EXPIRED = &H1E
Public Const TA_E_MUST_SPECIFY_TRIAL_TYPE = &H1F
Public Const TA_E_MUST_USE_TRIAL = &H20
Public Const TA_E_NO_MORE_TRIALS_ALLOWED = &H21

' You can either hardcode the paths or retrieve them dynamically.
Public Function GetTADirectory() As String
    ' Get the directory for TurboFloat (TurboFloat.dll and/or TurboFloat.x64.dll).
    ' On Windows you can get the directory using the registry
    ' See: http://support.microsoft.com/kb/145679/en-us
    ' For example, if you have an installer for your extension just set the path
    ' in the regirstry in your intaller. Then, in this function, read that value from
    ' the registry.
    ' Here we're just getting the path of the current Excel workbook
    #If Mac Then
        MsgBox "This worksheet only works on Windows"
        ' You can specify the same folder for both versions of Office, or choose separate
        ' folders. It's completely up to you.

        ' Special case for Office 2016 and above
        #If MAC_OFFICE_VERSION >= 15 Then
            GetTADirectory = "/Library/Application Support/Microsoft/YourApp"
        #Else                                    ' Office 2011
            GetTADirectory = "/Library/Application Support/Microsoft/YourApp"
        #End If
    #Else                                        ' Windows
        If Worksheets("Control").OptionButton1 Then
            GetTADirectory = Left(Environ("windir"), 3) & "Program Files\Next-In"
        ElseIf Worksheets("Control").OptionButton2 Then
            GetTADirectory = Environ("ProgramFiles") & "\Next-In"
        Else
            GetTADirectory = "c:\Next-In"
        End If
        'Debug.Print GetTADirectory
    #End If
End Function

