Attribute VB_Name = "ModValidKey"
Option Explicit

Public IsGenuine As Boolean                      ' used for checking if the key is genuine
Public nExpiryState As Integer                   ' three state : 0 is valid
'               1 is expired as of date but within the grace days
'               2 is expired completly. cannot use the application
Public nExpiryBalanceDays As Integer             ' used for checking expiry
Public dtExpiryDate As Date                      ' used for checking expiry, expiry date as per key
Public noLongerActivated As Boolean              ' used for checking expiry

Dim trialFlags As Long

Dim DaysBetweenChecks As Long
Dim GracePeriodLength As Long                    ' number of days after the expiry date. This is used while checking for validity of the key
Public ta As TurboActivate

'Private Sub UserForm_Initialize()
Public Sub CheckActivation()
    If Worksheets("Control").OptionButton4 Then Exit Sub
    Dim dtExpiryDateWithGraceDays As Date
    ' Set the trial flags you want to use. Here we've selected that the
    ' trial data should be stored system-wide (TA_SYSTEM) and that we should
    ' use un-resetable verified trials (TA_VERIFIED_TRIAL).
    trialFlags = TA_SYSTEM Or TA_VERIFIED_TRIAL
  
    ' Don't use 0 for either of these values.
    ' We recommend 90, 14. But if you want to lower the values
    ' we don't recommend going below 7 days for each value.
    ' Anything lower and you're just punishing legit users.
    DaysBetweenChecks = 30
    GracePeriodLength = 7
    noLongerActivated = False
    ' create the new TurboFloat instance
    
    ' tell your app to handle errors in the section
    ' at the end of the sub
    On Error GoTo TAProcError
  
    'TODO: goto the version page at LimeLM and paste this GUID here
    Set ta = GetTA                               '.Init("e2w5q6vbvtkm4tgdesasnyyob76wbgy")
  
    Dim gr As IsGenuineResult
  
    ' Check if we're activated, and every 90 days verify it with the activation servers
    ' In this example we won't show an error if the activation was done offline
    ' (see the 3rd parameter of the IsGenuine() function)
    ' https://wyday.com/limelm/help/offline-activation/
    gr = GetTA().IsGenuineEx(DaysBetweenChecks, GracePeriodLength, True)
  
    IsGenuine = gr = IsGenuineResult.Genuine Or _
                gr = IsGenuineResult.GenuineFeaturesChanged Or _
                gr = IsGenuineResult.InternetError
  
  
    ' If IsGenuineEx() is telling us we're not activated
    ' but the IsActivated() function is telling us that the activation
    ' data on the computer is valid (i.e. the crypto-signed-fingerprint matches the computer)
    ' then that means that the customer has passed the grace period and they must re-verify
    ' with the servers to continue to use your app.
  
  
    ' the following code checks for expiry date
    If IsGenuine And GetTA().IsActivated Then
        Select Case GetTA().IsGenuine()
        Case IsGenuineResult.Genuine, IsGenuineResult.GenuineFeaturesChanged
            IsGenuine = True
        Case IsGenuineResult.NotGenuine, IsGenuineResult.NotGenuineInVM
            noLongerActivated = True
        End Select
        dtExpiryDate = GetTA().GetFeatureValue("Expiry", Format(Now(), "dd-MMM-yyyy"))
        dtExpiryDateWithGraceDays = DateAdd("d", GracePeriodLength, dtExpiryDate)
        nExpiryBalanceDays = DateDiff("d", Date, dtExpiryDateWithGraceDays) ' including grace period
        If Now() > dtExpiryDateWithGraceDays Then
            nExpiryState = 2
        ElseIf Date > dtExpiryDate And Date < dtExpiryDateWithGraceDays Then
            nExpiryState = 1
        Else
            nExpiryState = 0
        End If
    End If
ProcExit:
    Exit Sub
TAProcError:
    MsgBox "Failed to check if activated: " & Err.Description
    ' End your application immediately
    End
End Sub

' procedure to deactivate the key by the user.
Public Sub AccountDeactivate()
    If IsGenuine Then
        ' tell your app to handle errors in the section
        ' at the end of the sub
        On Error GoTo TAProcError
        'deactivate product without deleting the product key
        'allows the user to easily reactivate
        GetTA().Deactivate
        IsGenuine = False
    Else
        ' tell your app to handle errors in the section
        ' at the end of the sub
        On Error GoTo TAProcIsActError
        ' launch the product key form
        Dim pkeyBox As frmPKey
        Set pkeyBox = New frmPKey

        Dim diagResult As VbMsgBoxResult
        diagResult = pkeyBox.ShowDialog(ta)

        If diagResult = vbOK And GetTA().IsActivated Then
            IsGenuine = True
        End If
    End If
ProcExit:
    Exit Sub
TAProcError:
    MsgBox "Failed to deactivate: " & Err.Description
    Resume ProcExit
TAProcIsActError:
    MsgBox "Failed to check if activated: " & Err.Description
    Resume ProcExit
End Sub

' function to check validity of the key, including expiry, and enables / disables feature in the application.
Public Function CheckForValidKey()
    On Error GoTo ErrorProc
    If Not Worksheets("Control").OptionButton4 Then 'If not a local lock, perform the following
        CheckForValidKey = True
        Call ModValidKey.CheckActivation
    
        Sheet12.cmdRead.Enabled = IsGenuine
        Sheet12.cmdReset.Enabled = IsGenuine
        Sheet12.cmdWrite.Enabled = IsGenuine
      
        If noLongerActivated Then
            Sheet12.cmdRead.Enabled = False
            Sheet12.cmdReset.Enabled = False
            Sheet12.cmdWrite.Enabled = False
        End If
      
        If Not IsGenuine Or noLongerActivated Then
            CheckForValidKey = False
            Sheet12.cmdActivation.Caption = "Activate"
        Else
            Sheet12.cmdActivation.Caption = "Deactivate"
        End If
        Select Case nExpiryState
        Case 0
        Case 1
            MsgBox "Your license has expired." _
                 & vbCrLf & "The software will stop working in " & nExpiryBalanceDays & " days." _
                 & vbCrLf & "Please contact info@nevergray.biz to extend your license.", vbInformation + vbOKOnly, "Renew License"
        Case 2
            Sheet12.cmdRead.Enabled = False
            Sheet12.cmdReset.Enabled = False
            Sheet12.cmdWrite.Enabled = False
            Sheet12.cmdActivation.Caption = "Activate"
            CheckForValidKey = False
            MsgBox "Your Product Key has expired on " & dtExpiryDate, vbInformation + vbOKOnly, "Renew License"
        End Select
        Exit Function
    Else
        If local_lock_subroutines.Local_Lock Then
            CheckForValidKey = True
            Sheet12.cmdRead.Enabled = True
            Sheet12.cmdReset.Enabled = True
            Sheet12.cmdWrite.Enabled = True
        Else
            CheckForValidKey = False
            Sheet12.cmdRead.Enabled = False
            Sheet12.cmdReset.Enabled = False
            Sheet12.cmdWrite.Enabled = False
        End If
        Exit Function
    End If
    
ErrorProc:
    MsgBox "Error in function CheckForValidKey : " & Err.Description
    Err.Clear
End Function

Public Function GetTA() As TurboActivate
    If ta Is Nothing Then
        Set ta = New TurboActivate
        Call ta.Init("e2w5q6vbvtkm4tgdesasnyyob76wbgy")
    End If
    Set GetTA = ta
End Function


