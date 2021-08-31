VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TurboActivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private versGUID As String
Private handle As Long

' Flags for GENUINE_OPTIONS

' If the user activated using offline activation
' (ActivateRequestToFile(), ActivateFromFile() ), then with this
' flag IsGenuineEx() will still try to validate with the LimeLM
' servers, however instead of returning TA_E_INET (when within the
' grace period) or TA_FAIL (when past the grace period) it will
' instead only return TA_OK (if IsActivated()).
'
' If you still want to get the TA_E_INET error code, without
' deactivating after the grace period has expired, then use
' this flag in tandem with TA_OFFLINE_SHOW_INET_ERR.
'
' If the user activated using online activation then this flag
' is ignored.
Const TA_SKIP_OFFLINE = 1

' If the user activated using offline activation, and you're
' using this flag in tandem with TA_SKIP_OFFLINE, then IsGenuineEx()
' will return TA_E_INET on internet failure instead of TA_OK.

' If the user activated using online activation then this flag
' is ignored.
Const TA_OFFLINE_SHOW_INET_ERR = 2



Private Type ACTIVATE_OPTIONS
    nLength As Long
    #If Win64 Then
        sExtraData As LongPtr
    #Else ' Windows / Mac 32-bit
        sExtraData As Long
    #End If
End Type

Private Type GENUINE_OPTIONS
    nLength As Long
    flags As Long
    nDaysBetweenChecks As Long
    nGraceDaysOnInetErr As Long
End Type


Public Enum IsGenuineResult
    ' Is activated and genuine.
    Genuine = 0

    ' Is activated and genuine and the features changed.
    GenuineFeaturesChanged = 1

    ' Not genuine (note: use this in tandem with NotGenuineInVM).
    NotGenuine = 2

    ' Not genuine because you're in a Virtual Machine.
    NotGenuineInVM = 3

    ' Treat this error as a warning. That is, tell the user that the activation couldn't be validated with the servers and that they can manually recheck with the servers immediately.
    InternetError = 4
End Enum


#If Mac Then
    Private Declare Function TA_GetHandle Lib "/path/to/your/libTurboActivate.dylib" (ByVal versionGUID As String) As Long
    Private Declare Function TA_Activate Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal options As Long) As Long
    Private Declare Function TA_ActivateEx Lib "/path/to/your/libTurboActivate.dylib" Alias "TA_Activate" (ByVal handle As Long, ByRef options As ACTIVATE_OPTIONS) As Long
    Private Declare Function TA_ActivationRequestToFile Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal FileName As String, ByVal options As Long) As Long
    Private Declare Function TA_ActivationRequestToFileEx Lib "/path/to/your/libTurboActivate.dylib" Alias "TA_ActivationRequestToFile" (ByVal handle As Long, ByVal FileName As String, ByRef options As ACTIVATE_OPTIONS) As Long
    Private Declare Function TA_ActivateFromFile Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal FileName As String) As Long
    Private Declare Function TA_CheckAndSavePKey Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal productKey As String, ByVal flags As Long) As Long
    Private Declare Function TA_Deactivate Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal erasePkey As Byte) As Long
    Private Declare Function TA_DeactivationRequestToFile Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal FileName As String, ByVal erasePkey As Byte) As Long
    Private Declare Function TA_GetExtraData Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal lpValueStr As String, ByVal cchValue As Long) As Long
    Private Declare Function TA_GetFeatureValue Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal featureName As String, ByVal lpValueStr As String, ByVal cchValue As Long) As Long
    Private Declare Function TA_GetPKey Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal lpPKeyStr As String, ByVal cchPKey As Long) As Long
    Private Declare Function TA_IsActivated Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long) As Long
    Private Declare Function TA_IsDateValid Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal date_time As String, ByVal flags As Long) As Long
    Private Declare Function TA_IsGenuine Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long) As Long
    Private Declare Function TA_IsGenuineEx Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByRef options As GENUINE_OPTIONS) As Long
    Private Declare Function TA_GenuineDays Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal nDaysBetweenChecks As Long, ByVal nGraceDaysOnInetErr As Long, ByRef DaysRemaining As Long, ByRef inGracePeriod As Byte) As Long
    Private Declare Function TA_IsProductKeyValid Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long) As Long
    Private Declare Function TA_TrialDaysRemaining Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal useTrialFlags As Long, ByRef DaysRemaining As Long) As Long
    Private Declare Function TA_UseTrial Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal flags As Long, ByVal extra_data As String) As Long
    Private Declare Function TA_UseTrialVerifiedRequest Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal FileName As String, ByVal extra_data As String) As Long
    Private Declare Function TA_UseTrialVerifiedFromFile Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal FileName As String, ByVal flags As Long) As Long
    Private Declare Function TA_ExtendTrial Lib "/path/to/your/libTurboActivate.dylib" (ByVal handle As Long, ByVal flags As Long, ByVal TrialExtension As String) As Long
    Private Declare Function TA_PDetsFromPath Lib "/path/to/your/libTurboActivate.dylib" (ByVal FileName As String) As Long
#ElseIf Win64 Then
    Private Declare PtrSafe Function TA_GetHandle Lib "TurboActivate.x64.dll" (ByVal versionGUID As LongPtr) As Long
    Private Declare PtrSafe Function TA_Activate Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal options As Long) As Long
    Private Declare PtrSafe Function TA_ActivateEx Lib "TurboActivate.x64.dll" Alias "TA_Activate" (ByVal handle As Long, ByRef options As ACTIVATE_OPTIONS) As Long
    Private Declare PtrSafe Function TA_ActivationRequestToFile Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal FileName As LongPtr, ByVal options As Long) As Long
    Private Declare PtrSafe Function TA_ActivationRequestToFileEx Lib "TurboActivate.x64.dll" Alias "TA_ActivationRequestToFile" (ByVal handle As Long, ByVal FileName As LongPtr, ByRef options As ACTIVATE_OPTIONS) As Long
    Private Declare PtrSafe Function TA_ActivateFromFile Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal FileName As LongPtr) As Long
    Private Declare PtrSafe Function TA_CheckAndSavePKey Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal productKey As LongPtr, ByVal flags As Long) As Long
    Private Declare PtrSafe Function TA_Deactivate Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal erasePkey As Byte) As Long
    Private Declare PtrSafe Function TA_DeactivationRequestToFile Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal FileName As LongPtr, ByVal erasePkey As Byte) As Long
    Private Declare PtrSafe Function TA_GetExtraData Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal lpValueStr As LongPtr, ByVal cchValue As Long) As Long
    Private Declare PtrSafe Function TA_GetFeatureValue Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal featureName As LongPtr, ByVal lpValueStr As LongPtr, ByVal cchValue As Long) As Long
    Private Declare PtrSafe Function TA_GetPKey Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal lpPKeyStr As LongPtr, ByVal cchPKey As Long) As Long
    Private Declare PtrSafe Function TA_IsActivated Lib "TurboActivate.x64.dll" (ByVal handle As Long) As Long
    Private Declare PtrSafe Function TA_IsDateValid Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal date_time As LongPtr, ByVal flags As Long) As Long
    Private Declare PtrSafe Function TA_IsGenuine Lib "TurboActivate.x64.dll" (ByVal handle As Long) As Long
    Private Declare PtrSafe Function TA_IsGenuineEx Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByRef options As GENUINE_OPTIONS) As Long
    Private Declare PtrSafe Function TA_GenuineDays Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal nDaysBetweenChecks As Long, ByVal nGraceDaysOnInetErr As Long, ByRef DaysRemaining As Long, ByRef inGracePeriod As Byte) As Long
    Private Declare PtrSafe Function TA_IsProductKeyValid Lib "TurboActivate.x64.dll" (ByVal handle As Long) As Long
    Private Declare PtrSafe Function TA_TrialDaysRemaining Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal useTrialFlags As Long, ByRef DaysRemaining As Long) As Long
    Private Declare PtrSafe Function TA_UseTrial Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal flags As Long, ByVal extra_data As LongPtr) As Long
    Private Declare PtrSafe Function TA_UseTrialVerifiedRequest Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal FileName As LongPtr, ByVal extra_data As LongPtr) As Long
    Private Declare PtrSafe Function TA_UseTrialVerifiedFromFile Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal FileName As LongPtr, ByVal flags As Long) As Long
    Private Declare PtrSafe Function TA_ExtendTrial Lib "TurboActivate.x64.dll" (ByVal handle As Long, ByVal flags As Long, ByVal TrialExtension As LongPtr) As Long
    Private Declare PtrSafe Function TA_PDetsFromPath Lib "TurboActivate.x64.dll" (ByVal FileName As LongPtr) As Long
#Else ' Win32
    Private Declare Function TA_GetHandle Lib "TurboActivate.dll" (ByVal versionGUID As Long) As Long
    Private Declare Function TA_Activate Lib "TurboActivate.dll" (ByVal handle As Long, ByVal options As Long) As Long
    Private Declare Function TA_ActivateEx Lib "TurboActivate.dll" Alias "TA_Activate" (ByVal handle As Long, ByRef options As ACTIVATE_OPTIONS) As Long
    Private Declare Function TA_ActivationRequestToFile Lib "TurboActivate.dll" (ByVal handle As Long, ByVal FileName As Long, ByVal options As Long) As Long
    Private Declare Function TA_ActivationRequestToFileEx Lib "TurboActivate.dll" Alias "TA_ActivationRequestToFile" (ByVal handle As Long, ByVal FileName As Long, ByRef options As ACTIVATE_OPTIONS) As Long
    Private Declare Function TA_ActivateFromFile Lib "TurboActivate.dll" (ByVal handle As Long, ByVal FileName As Long) As Long
    Private Declare Function TA_CheckAndSavePKey Lib "TurboActivate.dll" (ByVal handle As Long, ByVal productKey As Long, ByVal flags As Long) As Long
    Private Declare Function TA_Deactivate Lib "TurboActivate.dll" (ByVal handle As Long, ByVal erasePkey As Byte) As Long
    Private Declare Function TA_DeactivationRequestToFile Lib "TurboActivate.dll" (ByVal handle As Long, ByVal FileName As Long, ByVal erasePkey As Byte) As Long
    Private Declare Function TA_GetExtraData Lib "TurboActivate.dll" (ByVal handle As Long, ByVal lpValueStr As Long, ByVal cchValue As Long) As Long
    Private Declare Function TA_GetFeatureValue Lib "TurboActivate.dll" (ByVal handle As Long, ByVal featureName As Long, ByVal lpValueStr As Long, ByVal cchValue As Long) As Long
    Private Declare Function TA_GetPKey Lib "TurboActivate.dll" (ByVal handle As Long, ByVal lpPKeyStr As Long, ByVal cchPKey As Long) As Long
    Private Declare Function TA_IsActivated Lib "TurboActivate.dll" (ByVal handle As Long) As Long
    Private Declare Function TA_IsDateValid Lib "TurboActivate.dll" (ByVal handle As Long, ByVal date_time As Long, ByVal flags As Long) As Long
    Private Declare Function TA_IsGenuine Lib "TurboActivate.dll" (ByVal handle As Long) As Long
    Private Declare Function TA_IsGenuineEx Lib "TurboActivate.dll" (ByVal handle As Long, ByRef options As GENUINE_OPTIONS) As Long
    Private Declare Function TA_GenuineDays Lib "TurboActivate.dll" (ByVal handle As Long, ByVal nDaysBetweenChecks As Long, ByVal nGraceDaysOnInetErr As Long, ByRef DaysRemaining As Long, ByRef inGracePeriod As Byte) As Long
    Private Declare Function TA_IsProductKeyValid Lib "TurboActivate.dll" (ByVal handle As Long) As Long
    Private Declare Function TA_TrialDaysRemaining Lib "TurboActivate.dll" (ByVal handle As Long, ByVal useTrialFlags As Long, ByRef DaysRemaining As Long) As Long
    Private Declare Function TA_UseTrial Lib "TurboActivate.dll" (ByVal handle As Long, ByVal flags As Long, ByVal extra_data As Long) As Long
    Private Declare Function TA_UseTrialVerifiedRequest Lib "TurboActivate.dll" (ByVal handle As Long, ByVal FileName As Long, ByVal extra_data As Long) As Long
    Private Declare Function TA_UseTrialVerifiedFromFile Lib "TurboActivate.dll" (ByVal handle As Long, ByVal FileName As Long, ByVal flags As Long) As Long
    Private Declare Function TA_ExtendTrial Lib "TurboActivate.dll" (ByVal handle As Long, ByVal flags As Long, ByVal TrialExtension As Long) As Long
    Private Declare Function TA_PDetsFromPath Lib "TurboActivate.dll" (ByVal FileName As Long) As Long
#End If

Private Const ERR_COM = "CoInitializeEx failed. Re-enable Windows Management Instrumentation (WMI) service. Contact your system admin for more information."
Private Const ERR_ACCT_CANCELED = "Can't activate because the LimeLM account is cancelled."
Private Const ERR_REVOKED = "The product key has been revoked."
Private Const ERR_INUSE = "The product key has already been activated with the maximum number of computers."
Private Const ERR_INTERNET = "Connection to the servers failed."
Private Const ERR_INVALID_PKEY = "The product key is invalid or there's no product key."
Private Const ERR_ACTIVATE = "The product needs to be activated."
Private Const ERR_PDETS = "The product details file ""TurboActivate.dat"" failed to load. It's either missing or corrupt."
Private Const ERR_HANDLE = "The handle is not valid. You must set the VersionGUID property."
Private Const ERR_TRIALUSED = "The trial extension has already been used."
Private Const ERR_TRIALEXP = "The trial extension has expired."
Private Const ERR_EXPIRED = "Failed because the system date and time settings are incorrect. Fix your date and time settings, restart your computer, and try to activate again."
Private Const ERR_PERMISSION = "Insufficient system permission. Either start your process as an admin / elevated user or call the function again with the TA_USER flag."
Private Const ERR_FLAGS = "The flags you passed to the function were invalid (or missing). Flags like ""TA_SYSTEM"" and ""TA_USER"" are mutually exclusive -- you can only use one or the other."
Private Const ERR_VM = "The function failed because this instance of your program is running inside a virtual machine / hypervisor and you've prevented the function from running inside a VM."
Private Const ERR_EXTRADATALONG = "The ""extra data"" was too long. You're limited to 255 UTF-8 characters. Or, on Windows, a Unicode string that will convert into 255 UTF-8 characters or less."
Private Const ERR_INVALIDARGS = "The arguments passed to the function are invalid. Double check your logic."
Private Const ERR_TFKEY = "The product key used is for TurboFloat, not TurboActivate."
Private Const ERR_NO_DEACTS = "No more deactivations are allowed for the product key. This product is still activated on this computer."
Private Const ERR_ENABLE_NET_ADAPT = "There are network adapters on the system that are disabled and TurboActivate couldn't read their hardware properties (even after trying and failing to enable the adapters automatically). Enable the network adapters, re-run the function, and TurboActivate will be able to ""remember"" the adapters even if the adapters are disabled in the future."
Private Const ERR_ALREADY_ACT = "You can't use a product key because your app is already activated with a product key. To use a new product key, then first deactivate using either the Deactivate() or DeactivationRequestToFile()."
Private Const ERR_ALREADY_VT = "The trial is already a verified trial. You need to use the ""TA_VERIFIED_TRIAL"" flag. Can't ""downgrade"" a verified trial to an unverified trial."
Private Const ERR_TRIAL_EXPIRED = "The verified trial has expired. You must request a trial extension from the company."
Private Const ERR_NO_MORE_TRIALS = "In the LimeLM account either the trial days is set to 0, OR the account is set to not auto-upgrade and thus no more verified trials can be made."
Private Const ERR_MUST_SPECIFY_TRIAL = "You must specify the trial type (TA_UNVERIFIED_TRIAL or TA_VERIFIED_TRIAL). And you can't use both flags. Choose one or the other. We recommend TA_VERIFIED_TRIAL."
Private Const ERR_MUST_USETRIAL = "You must call TA_UseTrial() before you can get the number of trial days remaining."


Public Sub Init(ByVal vGUID As String, Optional ByVal pdetsFilename As String = vbNullString)
    If pdetsFilename = "" Then
        pdetsFilename = GetTADirectory() & Application.PathSeparator & "TurboActivate.dat"
    End If

#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long

#If Mac Then
    ret = TA_PDetsFromPath(pdetsFilename)
#Else ' Windows
    ret = TA_PDetsFromPath(StrPtr(pdetsFilename))

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_OK ' successful
        Case TA_FAIL ' the TurboActivate.dat already loaded.

        Case TA_E_PDETS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_PDETS, _
                Source:="TurboActivate.Init"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="The TurboActivate.dat file failed to load.", _
                Source:="TurboActivate.Init"
    End Select

    versGUID = vGUID

#If Win32 Then
    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

#If Mac Then
    handle = TA_GetHandle(versGUID)
#Else ' Windows
    handle = TA_GetHandle(StrPtr(versGUID))

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    ' If there's no handle throw an exception immediately
    If handle = 0 Then
        Err.Raise _
            Number:=ret, _
            Description:=ERR_PDETS, _
            Source:="TurboActivate.Init"
    End If
End Sub


''' <summary>Activates the product on this computer. You must call <see cref="CheckAndSavePKey"/> with a valid product key or have used the TurboActivate wizard sometime before calling this function.</summary>
''' <param name="extraData">Extra data to pass to the LimeLM servers that will be visible for you to see and use. Maximum size is 255 UTF-8 characters.</param>
Public Sub Activate(Optional ByVal extraData As String = vbNullString)
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long

    If extraData <> "" Then
        Dim options As ACTIVATE_OPTIONS

        ' Len(options) screws up the struct size
        ' Thus we have to hardcode them.
#If Win64 Then
        options.nLength = 16
#Else ' Windows / Mac 32-bit
        options.nLength = 8
#End If

#If Mac Then
        Dim sABSTR As String

        ' Convert to ANSI
        sABSTR = StrConv(extraData, vbFromUnicode)

        options.sExtraData = StrPtr(sABSTR)
#Else ' Windows
        options.sExtraData = StrPtr(extraData)
#End If

        ret = TA_ActivateEx(handle, options)
    Else
        ret = TA_Activate(handle, 0)
    End If


#If Win32 Then
    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If


    Select Case ret
        Case TA_OK
            Exit Sub

        Case TA_E_PKEY
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALID_PKEY, _
                Source:="TurboActivate.Activate"

        Case TA_E_INET
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INTERNET, _
                Source:="TurboActivate.Activate"

        Case TA_E_INUSE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INUSE, _
                Source:="TurboActivate.Activate"

        Case TA_E_REVOKED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_REVOKED, _
                Source:="TurboActivate.Activate"

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.Activate"

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.Activate"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.Activate"

        Case TA_E_EXPIRED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_EXPIRED, _
                Source:="TurboActivate.Activate"

        Case TA_E_IN_VM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_VM, _
                Source:="TurboActivate.Activate"

        Case TA_E_EDATA_LONG
            Err.Raise _
                Number:=ret, _
                Description:=ERR_EXTRADATALONG, _
                Source:="TurboActivate.Activate"

        Case TA_E_INVALID_ARGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALIDARGS, _
                Source:="TurboActivate.Activate"

        Case TA_E_KEY_FOR_TURBOFLOAT
            Err.Raise _
                Number:=ret, _
                Description:=ERR_TFKEY, _
                Source:="TurboActivate.Activate"

        Case TA_E_ACCOUNT_CANCELED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ACCT_CANCELED, _
                Source:="TurboActivate.Activate"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to activate.", _
                Source:="TurboActivate.Activate"
    End Select
End Sub


''' <summary>Get the "activation request" file for offline activation.  You must call <see cref="CheckAndSavePKey"/> with a valid product key or have used the TurboActivate wizard sometime before calling this function.</summary>
''' <param name="filename">The location where you want to save the activation request file.</param>
''' <param name="extraData">Extra data to pass to the LimeLM servers that will be visible for you to see and use. Maximum size is 255 UTF-8 characters.</param>
Public Sub ActivationRequestToFile(ByVal FileName As String, Optional ByVal extraData As String = vbNullString)
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long

    If extraData <> "" Then
        Dim options As ACTIVATE_OPTIONS

        ' Len(options) screws up the struct size
        ' Thus we have to hardcode them.
#If Win64 Then
        options.nLength = 16
#Else ' Windows / Mac 32-bit
        options.nLength = 8
#End If

#If Mac Then
        Dim sABSTR As String

        ' Convert to ANSI
        sABSTR = StrConv(extraData, vbFromUnicode)

        options.sExtraData = StrPtr(sABSTR)
        ret = TA_ActivationRequestToFileEx(handle, FileName, options)
#Else ' Windows
        options.sExtraData = StrPtr(extraData)
        ret = TA_ActivationRequestToFileEx(handle, StrPtr(FileName), options)
#End If
    Else
#If Mac Then
        ret = TA_ActivationRequestToFile(handle, FileName, 0)
#Else ' Windows
        ret = TA_ActivationRequestToFile(handle, StrPtr(FileName), 0)
#End If
    End If


#If Win32 Then
    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_OK
            Exit Sub

        Case TA_E_PKEY
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALID_PKEY, _
                Source:="TurboActivate.ActivationRequestToFile"

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.ActivationRequestToFile"

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.ActivationRequestToFile"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.ActivationRequestToFile"

        Case TA_E_EDATA_LONG
            Err.Raise _
                Number:=ret, _
                Description:=ERR_EXTRADATALONG, _
                Source:="TurboActivate.ActivationRequestToFile"

        Case TA_E_INVALID_ARGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALIDARGS, _
                Source:="TurboActivate.ActivationRequestToFile"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to save the activation request to file.", _
                Source:="TurboActivate.ActivationRequestToFile"
    End Select
End Sub


''' <summary>Activate from the "activation response" file for offline activation.</summary>
''' <param name="filename">The location of the activation response file.</param>
Public Sub ActivateFromFile(ByVal FileName As String)
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long

#If Mac Then
    ret = TA_ActivateFromFile(handle, FileName)
#Else ' Windows
    ret = TA_ActivateFromFile(handle, StrPtr(FileName))

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_OK
            Exit Sub

        Case TA_E_PKEY
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALID_PKEY, _
                Source:="TurboActivate.ActivateFromFile"

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.ActivateFromFile"

        Case TA_E_INVALID_ARGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALIDARGS, _
                Source:="TurboActivate.ActivateFromFile"

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.ActivateFromFile"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.ActivateFromFile"

        Case TA_E_EXPIRED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_EXPIRED, _
                Source:="TurboActivate.ActivateFromFile"

        Case TA_E_IN_VM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_VM, _
                Source:="TurboActivate.ActivateFromFile"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to activate.", _
                Source:="TurboActivate.ActivateFromFile"
    End Select
End Sub

Public Function CheckAndSavePKey(ByVal productKey As String, Optional ByVal flags As Long = TA_USER) As Boolean
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long

#If Mac Then
    ret = TA_CheckAndSavePKey(handle, productKey, flags)
#Else ' Windows
    ret = TA_CheckAndSavePKey(handle, StrPtr(productKey), flags)

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_OK
            CheckAndSavePKey = True
            Exit Function

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.CheckAndSavePKey"

        Case TA_E_PERMISSION
            Err.Raise _
                Number:=ret, _
                Description:=ERR_PERMISSION, _
                Source:="TurboActivate.CheckAndSavePKey"

        Case TA_E_INVALID_FLAGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_FLAGS, _
                Source:="TurboActivate.CheckAndSavePKey"

        Case TA_E_ALREADY_ACTIVATED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ALREADY_ACT, _
                Source:="TurboActivate.CheckAndSavePKey"

        Case Else
            CheckAndSavePKey = False
    End Select
End Function


''' <summary>Deactivates the product on this computer.</summary>
''' <param name="eraseProductKey">Erase the product key so the user will have to enter a new product key if they wish to reactivate.</param>
Public Sub Deactivate(Optional ByVal eraseProductKey As Boolean = False)
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim eraseKey As Byte
    
    If eraseProductKey Then
        eraseKey = 1
    Else
        eraseKey = 0
    End If

    Dim ret As Long
    ret = TA_Deactivate(handle, eraseKey)

#If Win32 Then
    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_OK
            Exit Sub

        Case TA_E_PKEY
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALID_PKEY, _
                Source:="TurboActivate.Deactivate"

        Case TA_E_ACTIVATE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ACTIVATE, _
                Source:="TurboActivate.Deactivate"

        Case TA_E_INET
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INTERNET, _
                Source:="TurboActivate.Deactivate"

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.Deactivate"

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.Deactivate"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.Deactivate"

        Case TA_E_NO_MORE_DEACTIVATIONS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_NO_DEACTS, _
                Source:="TurboActivate.Deactivate"

        Case TA_E_INVALID_ARGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALIDARGS, _
                Source:="TurboActivate.Deactivate"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to deactivate.", _
                Source:="TurboActivate.Deactivate"
    End Select
End Sub


''' <summary>Get the "deactivation request" file for offline deactivation.</summary>
''' <param name="filename">The location where you want to save the deactivation request file.</param>
''' <param name="eraseProductKey">Erase the product key so the user will have to enter a new product key if they wish to reactivate.</param>
Public Sub DeactivationRequestToFile(ByVal FileName As String, Optional ByVal eraseProductKey As Boolean = False)
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim eraseKey As Byte
    
    If eraseProductKey Then
        eraseKey = 1
    Else
        eraseKey = 0
    End If

    Dim ret As Long

#If Mac Then
    ret = TA_DeactivationRequestToFile(handle, FileName, eraseKey)
#Else ' Windows
    ret = TA_DeactivationRequestToFile(handle, StrPtr(FileName), eraseKey)

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_OK
            Exit Sub

        Case TA_E_PKEY
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALID_PKEY, _
                Source:="TurboActivate.DeactivationRequestToFile"

        Case TA_E_ACTIVATE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ACTIVATE, _
                Source:="TurboActivate.DeactivationRequestToFile"

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.DeactivationRequestToFile"

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.DeactivationRequestToFile"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.DeactivationRequestToFile"

        Case TA_E_INVALID_ARGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALIDARGS, _
                Source:="TurboActivate.DeactivationRequestToFile"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to deactivate.", _
                Source:="TurboActivate.DeactivationRequestToFile"
    End Select
End Sub


''' <summary>Gets the extra data value you passed in when activating.</summary>
''' <returns>Returns the extra data if it exists, otherwise it returns null.</returns>
Public Function GetExtraData() As String
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim strValue As String
    Dim ret As Long

    ' get the length of the field value
    ret = TA_GetExtraData(handle, 0, 0)

#If Mac Then
    ' allocate the buffer to store the string
    strValue = String(ret, 0)
    ret = TA_GetExtraData(handle, strValue, ret)
#Else ' Windows
    ' allocate the buffer to store the string
    strValue = String(ret * 2, 0)
    ret = TA_GetExtraData(handle, StrPtr(strValue), ret)

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.GetExtraData"

        Case TA_OK
            GetExtraData = strValue
            Exit Function

        Case Else
            GetExtraData = vbNullString
            Exit Function
    End Select
End Function


''' <summary>Gets the value of a feature (custom license field).</summary>
''' <param name="fieldName">The name of the custom license field to retrieve the value for.</param>
''' <param name="defaultValue">The default value to return if the license field doesn't exist.</param>
''' <returns>Returns the custom license field value if it exists, otherwise it returns the default value.</returns>
Public Function GetFeatureValue(ByVal fieldName As String, ByVal defaultValue As String) As String
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim strValue As String
    Dim ret As Long

#If Mac Then
    ' get the length of the field value
    ret = TA_GetFeatureValue(handle, fieldName, 0, 0)

    ' allocate the buffer to store the string
    strValue = String(ret, 0)
    ret = TA_GetFeatureValue(handle, fieldName, strValue, ret)
#Else ' Windows
    ' get the length of the field value
    ret = TA_GetFeatureValue(handle, StrPtr(fieldName), 0, 0)

    ' allocate the buffer to store the string
    strValue = String(ret * 2, 0)
    ret = TA_GetFeatureValue(handle, StrPtr(fieldName), StrPtr(strValue), ret)

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.GetFeatureValue"

        Case TA_OK
            GetFeatureValue = strValue
            Exit Function

        Case Else
            GetFeatureValue = defaultValue
            Exit Function
    End Select
End Function


''' <summary>Gets the stored product key. NOTE: if you want to check if a product key is valid simply call <see cref="IsProductKeyValid"/>.</summary>
Public Function GetPKey() As String
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim PKey As String
    Dim ret As Long

#If Mac Then
    PKey = String(35, 0)
    ret = TA_GetPKey(handle, PKey, 35)
#Else ' Windows
    PKey = String(70, 0)
    ret = TA_GetPKey(handle, StrPtr(PKey), 35)

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_PKEY
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALID_PKEY, _
                Source:="TurboActivate.GetPKey"

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.GetPKey"

        Case TA_OK
            GetPKey = PKey
            Exit Function

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to get the product key.", _
                Source:="TurboActivate.GetPKey"
    End Select
End Function


''' <summary>Checks whether the computer has been activated.</summary>
''' <returns>True if the computer is activated. False otherwise.</returns>
Public Function IsActivated() As Boolean
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long
    ret = TA_IsActivated(handle)

#If Win32 Then
    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.IsActivated"

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.IsActivated"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.IsActivated"

        Case TA_E_IN_VM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_VM, _
                Source:="TurboActivate.IsActivated"

        Case TA_OK
            IsActivated = True
            Exit Function
    End Select

    IsActivated = False
End Function


' <summary>Checks if the string in the form "YYYY-MM-DD HH:mm:ss" is a valid date/time. The date must be in UTC time and "24-hour" format. If your date is in some other time format first convert it to UTC time before passing it into this function.</summary>
' <param name="date_time">The date time string to check.</param>
' <param name="flags">The type of date time check. Valid flags are <see cref="TF_DateCheckFlags.TF_HAS_NOT_EXPIRED"/>.</param>
' <returns>True if the date is valid, false if it's not</returns>
Public Function IsDateValid(ByVal date_time As String, Optional ByVal flags As Long = TA_HAS_NOT_EXPIRED) As Boolean
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long

#If Mac Then
    ret = TA_IsDateValid(handle, date_time, flags)
#Else ' Windows
    ret = TA_IsDateValid(handle, StrPtr(date_time), flags)

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_OK
            IsDateValid = True
            Exit Function

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.IsDateValid"

        Case TA_E_INVALID_FLAGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_FLAGS, _
                Source:="TurboActivate.IsDateValid"

        Case Else
            IsDateValid = False
            Exit Function
    End Select
End Function


''' <summary>Checks whether the computer is genuinely activated by verifying with the LimeLM servers.</summary>
''' <returns>IsGenuineResult</returns>
Public Function IsGenuine() As IsGenuineResult
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long
    ret = TA_IsGenuine(handle)

#If Win32 Then
    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INET
            IsGenuine = IsGenuineResult.InternetError
            Exit Function

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.IsGenuine"

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.IsGenuine"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.IsGenuine"

        Case TA_E_EXPIRED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_EXPIRED, _
                Source:="TurboActivate.IsGenuine"

        Case TA_E_IN_VM
            IsGenuine = IsGenuineResult.NotGenuineInVM
            Exit Function

        Case TA_E_FEATURES_CHANGED
            IsGenuine = IsGenuineResult.GenuineFeaturesChanged
            Exit Function

        Case TA_OK
            IsGenuine = IsGenuineResult.Genuine
            Exit Function
    End Select

    ' not genuine (TA_FAIL, TA_E_REVOKED, TA_E_ACTIVATE)
    IsGenuine = IsGenuineResult.NotGenuine
End Function


''' <summary>Checks whether the computer is activated, and every "daysBetweenChecks" days it check if the customer is genuinely activated by verifying with the LimeLM servers.</summary>
''' <param name="daysBetweenChecks">How often to contact the LimeLM servers for validation. 90 days recommended.</param>
''' <param name="graceDaysOnInetErr">If the call fails because of an internet error, how long, in days, should the grace period last (before returning deactivating and returning TA_FAIL).
'''
''' 14 days is recommended.</param>
''' <param name="skipOffline">If the user activated using offline activation
''' (ActivateRequestToFile(), ActivateFromFile() ), then with this
''' option IsGenuineEx() will still try to validate with the LimeLM
''' servers, however instead of returning <see cref="IsGenuineResult.InternetError"/> (when within the
''' grace period) or <see cref="IsGenuineResult.NotGenuine"/> (when past the grace period) it will
''' instead only return <see cref="IsGenuineResult.Genuine"/> (if IsActivated()).
'''
''' If the user activated using online activation then this option
''' is ignored.</param>
''' <param name="offlineShowInetErr">If the user activated using offline activation, and you're
''' using this option in tandem with skipOffline, then IsGenuineEx()
''' will return <see cref="IsGenuineResult.InternetError"/> on internet failure instead of <see cref="IsGenuineResult.Genuine"/>.
'''
''' If the user activated using online activation then this flag
''' is ignored.</param>
''' <returns>IsGenuineResult</returns>
Public Function IsGenuineEx(ByVal DaysBetweenChecks As Long, ByVal graceDaysOnInetErr As Long, Optional ByVal skipOffline As Boolean = False, Optional ByVal offlineShowInetErr As Boolean = False) As IsGenuineResult
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim options As GENUINE_OPTIONS
    options.nLength = Len(options)
    options.flags = 0
    options.nDaysBetweenChecks = DaysBetweenChecks
    options.nGraceDaysOnInetErr = graceDaysOnInetErr

    If skipOffline Then
        options.flags = TA_SKIP_OFFLINE

        If offlineShowInetErr Then
            options.flags = options.flags Or TA_OFFLINE_SHOW_INET_ERR
        End If
    End If

    Dim ret As Long
    ret = TA_IsGenuineEx(handle, options)

#If Win32 Then
    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INET, TA_E_INET_DELAYED
            IsGenuineEx = IsGenuineResult.InternetError
            Exit Function

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.IsGenuineEx"

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.IsGenuineEx"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.IsGenuineEx"

        Case TA_E_EXPIRED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_EXPIRED, _
                Source:="TurboActivate.IsGenuineEx"

        Case TA_E_IN_VM
            IsGenuineEx = IsGenuineResult.NotGenuineInVM
            Exit Function

        Case TA_E_INVALID_ARGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALIDARGS, _
                Source:="TurboActivate.IsGenuineEx"

        Case TA_E_FEATURES_CHANGED
            IsGenuineEx = IsGenuineResult.GenuineFeaturesChanged
            Exit Function

        Case TA_OK
            IsGenuineEx = IsGenuineResult.Genuine
            Exit Function
    End Select

    ' not genuine (TA_FAIL, TA_E_REVOKED, TA_E_ACTIVATE)
    IsGenuineEx = IsGenuineResult.NotGenuine
End Function


''' <summary>Get the number of days until the next time that the <see cref="IsGenuine"/> function contacts the LimeLM activation servers to reverify the activation.</summary>
''' <param name="daysBetweenChecks">How often to contact the LimeLM servers for validation. Use the exact same value as used in <see cref="IsGenuine"/>.</param>
''' <param name="graceDaysOnInetErr">If the call fails because of an internet error, how long, in days, should the grace period last (before returning deactivating and returning TA_FAIL). Again, use the exact same value as used in <see cref="IsGenuine"/>.</param>
''' <param name="inGracePeriod">Get whether the user is in the grace period.</param>
''' <returns>The number of days remaining. 0 days if both the days between checks and the grace period have expired. (E.g. 1 day means *at most* 1 day. That is, it could be 30 seconds.)</returns>
Public Function GenuineDays(ByVal DaysBetweenChecks As Long, ByVal graceDaysOnInetErr As Long, ByRef inGracePeriod As Boolean) As Long
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim daysRemain As Long
    Dim inGrace As Byte

    daysRemain = 0
    inGrace = 0

    Dim ret As Long
    ret = TA_GenuineDays(handle, DaysBetweenChecks, graceDaysOnInetErr, daysRemain, inGrace)

#If Win32 Then
    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.GenuineDays"

        Case TA_E_ACTIVATE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ACTIVATE, _
                Source:="TurboActivate.GenuineDays"

        Case TA_OK
            ' successful, just fall through to the calculation

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to get the genuine days.", _
                Source:="TurboActivate.GenuineDays"
    End Select

    inGracePeriod = (inGrace = 1)
    GenuineDays = daysRemain
End Function


''' <summary>Checks if the product key installed for this product is valid. This does NOT check if the product key is activated or genuine. Use <see cref="IsActivated"/> and <see cref="IsGenuine"/> instead.</summary>
''' <returns>True if the product key is valid.</returns>
Public Function IsProductKeyValid() As Boolean
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long
    ret = TA_IsProductKeyValid(handle)

#If Win32 Then
    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.IsProductKeyValid"

        Case TA_OK
            IsProductKeyValid = True
            Exit Function
    End Select

    IsProductKeyValid = False
End Function
Public Function TrialDaysRemaining(Optional ByVal useTrialFlags As Long = TA_SYSTEM Or TA_VERIFIED_TRIAL) As Long
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim daysRemain As Long
    Dim ret As Long
    ret = TA_TrialDaysRemaining(handle, useTrialFlags, daysRemain)

#If Win32 Then
    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.TrialDaysRemaining"

        Case TA_OK
            TrialDaysRemaining = daysRemain
            Exit Function

        Case TA_E_ALREADY_VERIFIED_TRIAL
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ALREADY_VT, _
                Source:="TurboActivate.TrialDaysRemaining"

        Case TA_E_MUST_SPECIFY_TRIAL_TYPE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_MUST_SPECIFY_TRIAL, _
                Source:="TurboActivate.TrialDaysRemaining"

        Case TA_E_MUST_USE_TRIAL
            Err.Raise _
                Number:=ret, _
                Description:=ERR_MUST_USETRIAL, _
                Source:="TurboActivate.TrialDaysRemaining"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to get the trial data.", _
                Source:="TurboActivate.TrialDaysRemaining"
    End Select

    TrialDaysRemaining = 0
End Function


''' <summary>Begins the trial the first time it's called. Calling it again will validate the trial data hasn't been tampered with.</summary>
''' <param name="flags">Whether to create the trial (verified or unverified) either user-wide or system-wide and whether to allow trials in virtual machines. Valid flags are <see cref="TA_Flags.TA_SYSTEM"/>, <see cref="TA_Flags.TA_USER"/>, <see cref="TA_Flags.TA_DISALLOW_VM"/>, <see cref="TA_Flags.TA_VERIFIED_TRIAL"/>, and <see cref="TA_Flags.TA_UNVERIFIED_TRIAL"/>.</param>
''' <param name="extraData">Extra data to pass to the LimeLM servers that will be visible for you to see and use. Maximum size is 255 UTF-8 characters.</param>
Public Sub UseTrial(Optional ByVal flags As Long = TA_SYSTEM Or TA_VERIFIED_TRIAL, Optional ByVal extraData As String = vbNullString)
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long

#If Mac Then
    ret = TA_UseTrial(handle, flags, extraData)

#Else ' Windows
  #If Win64 Then
    Dim exDataPtr As LongPtr
  #Else ' Win32
    Dim exDataPtr As Long
  #End If
    exDataPtr = 0

    If extraData <> "" Then
        exDataPtr = StrPtr(extraData)
    End If

    ret = TA_UseTrial(handle, flags, exDataPtr)

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_INET
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INTERNET, _
                Source:="TurboActivate.UseTrial"

        Case TA_OK ' successful
            Exit Sub

        Case TA_E_PERMISSION
            Err.Raise _
                Number:=ret, _
                Description:=ERR_PERMISSION, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_INVALID_FLAGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_FLAGS, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_IN_VM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_VM, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_ACCOUNT_CANCELED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ACCT_CANCELED, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_ALREADY_VERIFIED_TRIAL
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ALREADY_VT, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_EXPIRED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_EXPIRED, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_TRIAL_EXPIRED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_TRIAL_EXPIRED, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_MUST_SPECIFY_TRIAL_TYPE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_MUST_SPECIFY_TRIAL, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_EDATA_LONG
            Err.Raise _
                Number:=ret, _
                Description:=ERR_EXTRADATALONG, _
                Source:="TurboActivate.UseTrial"

        Case TA_E_NO_MORE_TRIALS_ALLOWED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_NO_MORE_TRIALS, _
                Source:="TurboActivate.UseTrial"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to save the trial data.", _
                Source:="TurboActivate.UseTrial"
    End Select
End Sub


''' <summary>Generate a "verified trial" offline request file. This file will then need to be submitted to LimeLM. You will then need to use the TA_UseTrialVerifiedFromFile() function with the response file from LimeLM to actually start the trial.</summary>
''' <param name="filename">The location where you want to save the trial request file.</param>
''' <param name="extraData">Extra data to pass to the LimeLM servers that will be visible for you to see and use. Maximum size is 255 UTF-8 characters.</param>
Public Sub UseTrialVerifiedRequest(ByVal FileName As String, Optional ByVal extraData As String = vbNullString)
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long

#If Mac Then
    ret = TA_UseTrialVerifiedRequest(handle, FileName, extraData)

#Else ' Windows
  #If Win64 Then
    Dim exDataPtr As LongPtr
  #Else ' Win32
    Dim exDataPtr As Long
  #End If
    exDataPtr = 0

    If extraData <> "" Then
        exDataPtr = StrPtr(extraData)
    End If

    ret = TA_UseTrialVerifiedRequest(handle, StrPtr(FileName), exDataPtr)

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.UseTrialVerifiedRequest"

        Case TA_OK ' successful
            Exit Sub

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.UseTrialVerifiedRequest"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.UseTrialVerifiedRequest"

        Case TA_E_EDATA_LONG
            Err.Raise _
                Number:=ret, _
                Description:=ERR_EXTRADATALONG, _
                Source:="TurboActivate.UseTrialVerifiedRequest"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to save the verified trial request file.", _
                Source:="TurboActivate.UseTrialVerifiedRequest"
    End Select
End Sub


''' <summary>Use the "verified trial response" from LimeLM to start the verified trial.</summary>
''' <param name="filename">The location of the trial response file.</param>
''' <param name="flags">Whether to create the trial (verified or unverified) either user-wide or system-wide and whether to allow trials in virtual machines. Valid flags are <see cref="TA_Flags.TA_SYSTEM"/>, <see cref="TA_Flags.TA_USER"/>, <see cref="TA_Flags.TA_DISALLOW_VM"/>, <see cref="TA_Flags.TA_VERIFIED_TRIAL"/>, and <see cref="TA_Flags.TA_UNVERIFIED_TRIAL"/>.</param>
Public Sub UseTrialVerifiedFromFile(ByVal FileName As String, Optional ByVal flags As Long = TA_SYSTEM Or TA_VERIFIED_TRIAL)
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long
#If Mac Then
    ret = TA_UseTrialVerifiedFromFile(handle, FileName, flags)
#Else ' Windows
    ret = TA_UseTrialVerifiedFromFile(handle, StrPtr(FileName), flags)

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.UseTrialVerifiedFromFile"

        Case TA_OK ' successful
            Exit Sub

        Case TA_E_PERMISSION
            Err.Raise _
                Number:=ret, _
                Description:=ERR_PERMISSION, _
                Source:="TurboActivate.UseTrialVerifiedFromFile"

        Case TA_E_INVALID_FLAGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_FLAGS, _
                Source:="TurboActivate.UseTrialVerifiedFromFile"

        Case TA_E_INVALID_ARGS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INVALIDARGS, _
                Source:="TurboActivate.UseTrialVerifiedFromFile"

        Case TA_E_COM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_COM, _
                Source:="TurboActivate.UseTrialVerifiedFromFile"

        Case TA_E_ENABLE_NETWORK_ADAPTERS
            Err.Raise _
                Number:=ret, _
                Description:=ERR_ENABLE_NET_ADAPT, _
                Source:="TurboActivate.UseTrialVerifiedFromFile"

        Case TA_E_MUST_SPECIFY_TRIAL_TYPE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_MUST_SPECIFY_TRIAL, _
                Source:="TurboActivate.UseTrialVerifiedFromFile"

        Case TA_E_IN_VM
            Err.Raise _
                Number:=ret, _
                Description:=ERR_VM, _
                Source:="TurboActivate.UseTrialVerifiedFromFile"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to save the trial data.", _
                Source:="TurboActivate.UseTrialVerifiedFromFile"
    End Select
End Sub

Public Sub ExtendTrial(ByVal TrialExtension As String, Optional ByVal useTrialFlags As Long = TA_SYSTEM Or TA_VERIFIED_TRIAL)
#If Win32 Then
    Dim currDir As String
    Dim taDir As String

    currDir = CurDir()
    taDir = GetTADirectory()

    ' set the TurboActivate path as the current directory
    ChDrive taDir
    ChDir taDir
#End If

    Dim ret As Long
#If Mac Then
    ret = TA_ExtendTrial(handle, useTrialFlags, TrialExtension)
#Else ' Windows
    ret = TA_ExtendTrial(handle, useTrialFlags, StrPtr(TrialExtension))

    ' restore the original current directory
    ChDrive currDir
    ChDir currDir
#End If

    Select Case ret
        Case TA_OK
            Exit Sub

        Case TA_E_INET
            Err.Raise _
                Number:=ret, _
                Description:=ERR_INTERNET, _
                Source:="TurboActivate.ExtendTrial"

        Case TA_E_INVALID_HANDLE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_HANDLE, _
                Source:="TurboActivate.ExtendTrial"

        Case TA_E_TRIAL_EUSED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_TRIALUSED, _
                Source:="TurboActivate.ExtendTrial"

        Case TA_E_EXPIRED
            Err.Raise _
                Number:=ret, _
                Description:=ERR_TRIALEXP, _
                Source:="TurboActivate.ExtendTrial"

        Case TA_E_MUST_SPECIFY_TRIAL_TYPE
            Err.Raise _
                Number:=ret, _
                Description:=ERR_MUST_SPECIFY_TRIAL, _
                Source:="TurboActivate.ExtendTrial"

        Case TA_E_MUST_USE_TRIAL
            Err.Raise _
                Number:=ret, _
                Description:=ERR_MUST_USETRIAL, _
                Source:="TurboActivate.ExtendTrial"

        Case Else
            Err.Raise _
                Number:=ret, _
                Description:="Failed to extend the trial.", _
                Source:="TurboActivate.ExtendTrial"
    End Select
End Sub
