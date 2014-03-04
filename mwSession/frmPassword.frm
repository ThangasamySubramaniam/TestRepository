VERSION 5.00
Begin VB.Form frmPassword 
   Caption         =   "Change Password"
   ClientHeight    =   5808
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9048
   ControlBox      =   0   'False
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5808
   ScaleWidth      =   9048
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameStrongPassword 
      Caption         =   "NOTE: Strong Password Validation Options must be used for Shore roles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   600
      TabIndex        =   10
      Top             =   3600
      Width           =   7815
      Begin VB.Label lblMixedCaseCharacters 
         Caption         =   "- Mixed upper && lower case characters"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lblPasswordHistory 
         Caption         =   "- Cannot use previous passwords or reversed"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lblCaseSensitive 
         Caption         =   "- Case sensitive"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lblIncludeNumbers 
         Caption         =   "- Must include numbers"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lblUseSpecialCharacters 
         Caption         =   "- Must contain special characters ~!@#$%^&*()_+-={}|[]\:"";'<>?,./"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label lblMinimumChar 
         Caption         =   "- Minimum number of characters allowed:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      Picture         =   "frmPassword.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Picture         =   "frmPassword.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox pvmReEnter 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3720
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox pvmNewPassword 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3720
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1275
      Width           =   3135
   End
   Begin VB.TextBox pvmOldPassword 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3720
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   735
      Width           =   3135
   End
   Begin VB.Label lblUserID 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "User ID:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   660
      TabIndex        =   3
      Top             =   300
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   2
      Top             =   1350
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Re-Type New Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   1875
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Old Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   810
      Width           =   3135
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCanceled As Boolean
Private Const ENCRYPT_PSWD = "Gray" & "bar" & "327"

'By N.Angelakis On 22 April 2009
'DEV-1174 Advance Password Settings
Dim mIsPswdHistory As Boolean
Dim mIsPswdDisallowedList As Boolean
Dim mIsPswdIncludeNumbers As Boolean
Dim mIsPswdUseSpecialChar As Boolean
Dim mIsPswdMixedCaseChar As Boolean
Dim mIntPswdMinLength As Integer
Dim mIntPswdFailedAttemts As Integer

Dim mIsForcePasswordChange As Boolean
Dim WithEvents moRsUser As Recordset
Attribute moRsUser.VB_VarHelpID = -1
#If LATE_BIND Then
   Dim moKeys As Object
#Else
   Dim moKeys As ConfigKeys
#End If

Public Function ForcePasswordChange(ByVal mData As Boolean)
    mIsForcePasswordChange = mData
End Function

Public Function IsCanceled() As Boolean
   IsCanceled = mCanceled
End Function

Private Sub cmdCancel_Click()
   mCanceled = True
   Me.Hide
End Sub

Private Sub cmdLogin_Click()
   Dim loEncrypt As New mwEncrypt
   Dim OldPassword As String
   Dim MwPassword As String
   Dim Validated As Boolean

   'By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   Dim loDataWork As mwDataWork
   Dim strOracleContinuationSQL As String
   Dim strOracleContinuationSQL2 As String
   Dim strSQL As String
   
   On Error GoTo SubError
   '
   ' Call mwSession Function to change password...
   '
   mCanceled = False
   Validated = False
   
   loEncrypt.EnableEncryption ENCRYPT_PSWD
   
   If pvmNewPassword.Text = pvmReEnter.Text Then
      OldPassword = goSession.User.GetExtendedProperty("MwPassword")
      
      If goSession.User.GetExtendedProperty("PasswordEncrypted") = "1" Or _
         goSession.User.GetExtendedProperty("PasswordEncrypted") = "-1" Or UCase(goSession.User.GetExtendedProperty("PasswordEncrypted")) = "TRUE" Then
         
         OldPassword = loEncrypt.DecryptString(OldPassword)
         If Trim(pvmOldPassword.Text) <> Trim(OldPassword) Then
            goSession.GUI.ImprovedMsgBox "Old password does not match current password.", vbExclamation, "Validate Current Password"
            Validated = False
         Else
            Validated = True
            MwPassword = Trim(pvmNewPassword.Text)
               

            'By N.Angelakis On 22 April 2009
            'DEV-1174 Advance Password Settings
            If Not PasswordStrongCheck(goSession.User.UserKey, MwPassword) Then
               'show message in PasswordStrongCheck depending on problem
               'goSession.GUI.ImprovedMsgBox "New password........ Please try again.", vbOKOnly, "Password Validation"
               Validated = False
            Else 'If Not PasswordStrongCheck Then
               'update new password into mwcUsers
               
               MwPassword = loEncrypt.EncryptString(MwPassword)
               
               UpdateUserPassword goSession.User.UserKey, MwPassword
            
               If goSession.ThisSite.IsPasswordStrong = True Then
                  
                  'save in history table
                  Call SavePasswordHistory(goSession.User.UserKey, MwPassword)
               End If
            End If 'If Not PasswordStrongCheck Then
         End If
      Else
         If Trim(pvmOldPassword.Text) <> Trim(OldPassword) Then
            goSession.GUI.ImprovedMsgBox "Old password does not match current password.", vbExclamation, "Validate Current Password"
            Validated = False
         Else
            Validated = True
            MwPassword = Trim(pvmNewPassword.Text)
            
            'By N.Angelakis On 22 April 2009
            'DEV-1174 Advance Password Settings
            If Not PasswordStrongCheck(goSession.User.UserKey, MwPassword) Then
               'show message in PasswordStrongCheck depending on problem
               'goSession.GUI.ImprovedMsgBox "New password...... Please try again.", vbOKOnly, "Password Validation"
               Validated = False
            Else 'If Not PasswordStrongCheck Then
               'update new password into mwcUsers
               
               UpdateUserPassword goSession.User.UserKey, MwPassword
               
               If goSession.ThisSite.IsPasswordStrong = True Then
                  'save in history table, always encrypted
                  MwPassword = loEncrypt.EncryptString(MwPassword)
                  
                  Call SavePasswordHistory(goSession.User.UserKey, Trim(MwPassword))
               End If
            End If 'If Not PasswordStrongCheck Then
         End If
      End If
   Else
      goSession.GUI.ImprovedMsgBox "New passwords don't match. Please try again.", vbOKOnly, "Password Validation"
      'pvmNewPassword.Text = ""
      'pvmReEnter.Text = ""
      Validated = False
   End If
   
   If Validated = True Then
      goSession.GUI.ImprovedMsgBox "Password has been changed.", vbInformation, "Change Password Complete"
      Me.Hide
   End If
   Exit Sub
SubError:
   goSession.RaiseError "General Error in frmPassword.", Err.Number, Err.Description
End Sub
Private Function UpdateUserPassword(ByVal mwcUsersKey As Long, ByVal NewPassword As String) As Boolean
   Dim strSQL As String
   
   On Error GoTo ErrorHandler
   CloseRecordset moRsUser
   strSQL = "SELECT * FROM mwcUsers WHERE ID=" & mwcUsersKey
   Set moRsUser = New Recordset
   moRsUser.CursorLocation = adUseClient
   moRsUser.Open strSQL, goCon, adOpenDynamic, adLockOptimistic

   moRsUser!MwPassword = NewPassword
   moRsUser!LoginFailedAttempts = 0
   moRsUser!PasswordChangedDate = Now()
   moRsUser.Update
   CloseRecordset moRsUser
   
   UpdateUserPassword = True
   
   Exit Function
ErrorHandler:
   UpdateUserPassword = False
   goSession.RaisePublicError "General Error in msWorkstation.frmPassword.UpdateUserPassword.", Err.Number, Err.Description
End Function

Private Function SavePasswordHistory(ByVal mwcUsersKey As Long, ByVal NewPassword As String) As Boolean
   'Modified By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   
   Dim strSQL As String
   Dim loRs As Recordset
   
   On Error GoTo ErrorHandler

   strSQL = "SELECT * FROM mwcPasswordHistory WHERE mwcUsersKey=" & mwcUsersKey & " ORDER BY ChangeDate"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open strSQL, goCon, adOpenDynamic, adLockOptimistic

   While loRs.RecordCount >= 10
      loRs.Delete
      loRs.MoveNext
   Wend
   
   loRs.AddNew
   loRs!ID = goSession.MakePK("mwcPasswordHistory")
   loRs!mwcUsersKey = mwcUsersKey
   loRs!UsedPassword = NewPassword
   loRs!ChangeDate = Now()
   
   CloseRecordset loRs
   
   SavePasswordHistory = True
   
   Exit Function
ErrorHandler:
   SavePasswordHistory = False
   goSession.RaisePublicError "General Error in msWorkstation.frmPassword.SavePasswordHistory.", Err.Number, Err.Description
End Function

Private Function PasswordStrongCheck(ByVal mwcUsersKey As Long, ByVal sNewPassword As String) As Boolean
   'Modified By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   
   Dim loEncrypt As New mwEncrypt
   Dim EncryptedPassword As String
   Dim RevEncryptedPassword As String
   
   Dim strSQL As String
   Dim loRs As Recordset

   On Error GoTo ErrorHandler

   PasswordStrongCheck = True

   If goSession.User.IsShoreUser = False Or goSession.ThisSite.IsPasswordStrong = False Then
      'password strong validation not required so need to check further
      PasswordStrongCheck = True
      Exit Function
   End If


   If mIsPswdHistory = True Then
      loEncrypt.EnableEncryption ENCRYPT_PSWD
      
      EncryptedPassword = loEncrypt.EncryptString(sNewPassword)
      RevEncryptedPassword = loEncrypt.EncryptString(StrReverse(sNewPassword))
      
      strSQL = "SELECT * FROM mwcPasswordHistory " & _
               " WHERE mwcUsersKey=" & mwcUsersKey & " and (UsedPassword='" & EncryptedPassword & "'" & _
               " or UsedPassword='" & RevEncryptedPassword & "')" & _
               " ORDER BY ID Desc"

           
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
      If Not (loRs.BOF And loRs.EOF) Then
         goSession.GUI.ImprovedMsgBox "New password has already been used within the last 10 passwords. Please try again.", vbOKOnly, "Password Validation"
         PasswordStrongCheck = False
         CloseRecordset loRs
         Exit Function
      End If
      CloseRecordset loRs
   End If
   
   ' TJM0502: 0123456789
   If mIsPswdIncludeNumbers Then
      If Not IsStringContainsChars(sNewPassword, "0123456789") Then
         goSession.GUI.ImprovedMsgBox "New passwords must also contain numbers 1-9. Please try again.", vbOKOnly, "Password Validation"
         PasswordStrongCheck = False
         CloseRecordset loRs
         Exit Function
      End If
   End If

   ' TJM0502: ~!@#$%^&*()_+-={}|[]\:";'<>?,./
   If mIsPswdUseSpecialChar Then
      If Not IsStringContainsChars(sNewPassword, "~!@#$%^&*()_+-={}|[]\:"";'<>?,./") Then
         goSession.GUI.ImprovedMsgBox "New passwords must have at least one special character. Please try again.", vbOKOnly, "Password Validation"
         PasswordStrongCheck = False
         CloseRecordset loRs
         Exit Function
      End If
   End If
   

   If mIsPswdMixedCaseChar Then
      If Not IsMixedCased(sNewPassword) Then
         goSession.GUI.ImprovedMsgBox "New passwords must contain upper and lower case characters. Please try again.", vbOKOnly, "Password Validation"
         PasswordStrongCheck = False
         CloseRecordset loRs
         Exit Function
      End If
   End If

   If mIntPswdMinLength > 0 Then
      If Not Len(sNewPassword) >= mIntPswdMinLength Then
         goSession.GUI.ImprovedMsgBox "New passwords must be at least " & mIntPswdMinLength & " characters/digits long. Please try again.", vbOKOnly, "Password Validation"
         PasswordStrongCheck = False
         CloseRecordset loRs
         Exit Function
      End If
   End If


   If mIsPswdDisallowedList Then
      strSQL = "SELECT * FROM mwcPasswordDisallowed WHERE DisallowedPassword='" & sNewPassword & "'"
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
      
      If Not (loRs.BOF And loRs.EOF) Then
         goSession.GUI.ImprovedMsgBox "New password in disallowed passwords list. Please try again.", vbOKOnly, "Password Validation"
         PasswordStrongCheck = False
         CloseRecordset loRs
         Exit Function
      End If
      CloseRecordset loRs
   End If


   PasswordStrongCheck = True
   
   Exit Function
ErrorHandler:
   PasswordStrongCheck = False
   goSession.RaisePublicError "General Error in msWorkstation.frmPassword.PasswordStrongCheck.", Err.Number, Err.Description
End Function
Private Function IsStringContainsChars(ByVal StringToCheck As String, ByVal SpecialCharacters As String) As Boolean

   Dim ThisChar As String
   Dim xx As Long
   
   For xx = Len(SpecialCharacters) To 1 Step -1
      ThisChar = mID(SpecialCharacters, xx, 1)
      If InStr(StringToCheck, ThisChar) > 0 Then
         IsStringContainsChars = True
         Exit Function
      End If
   Next xx
   
   IsStringContainsChars = False
   Exit Function
ErrorHandler:
   IsStringContainsChars = False
   goSession.RaisePublicError "General Error in msWorkstation.frmPassword.IsStringContainsChars.", Err.Number, Err.Description
   
End Function
Private Function IsMixedCased(ByVal StringToCheck As String) As Boolean
   'Modified By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   Dim sCharString As String
   Dim sCharStringUpper As String
   Dim LowerCaseFound As Boolean
   Dim UpperCaseFound As Boolean
   
   On Error GoTo ErrorHandler

   sCharString = "abcdefghijklmnopqrstuvwxyz"
   sCharStringUpper = UCase(sCharString)
   
   If IsStringContainsChars(StringToCheck, sCharString) And _
      IsStringContainsChars(StringToCheck, sCharStringUpper) Then
      
      IsMixedCased = True
   Else
      IsMixedCased = False
   End If

   Exit Function
ErrorHandler:
   IsMixedCased = False
   goSession.RaisePublicError "General Error in msWorkstation.frmPassword.IsMixedCased.", Err.Number, Err.Description

End Function

Private Sub Form_Load()
   On Error GoTo SubError
   'By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   If mIsForcePasswordChange = True Then
      'password expiry date has passed , show change password screen
      'dont allow user to cancel screen and proceed must change password
      cmdCancel.Enabled = False
   End If
   lblUserID.Caption = goSession.User.UserID
   
   If goSession.User.IsShoreUser = True And goSession.ThisSite.IsPasswordStrong = True Then
      frameStrongPassword.Visible = True
      'password strong validation required load defaults
      LoadPasswordStrongSettings
   Else
      frameStrongPassword.Visible = False
   End If
   Set moKeys = goSession.GetEventSecurityKeys(SW_EV_USERS)
   
   goSession.SetDotNetTheme Me
   
   Exit Sub
SubError:
   MsgBox "Error in mwSession.frmPassword.Form_Load ", Err.Number, Err.Description
End Sub
Private Function LoadPasswordStrongSettings() As Boolean
   'By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   Dim loRs As Recordset
   Dim strSQL As String
   
   strSQL = "SELECT IsPasswordHistory, IsPasswordIncludeNumbers, " & _
            " IsPasswordUseSpecialCharacters,PasswordMinLength, PasswordExpireNoDays, " & _
            " PasswordFailedAttempts, IsPasswordDisallowedList, IsPasswordMixedCase FROM mwcThisSite"
            
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   lblCaseSensitive.Visible = True
   
   If Not (loRs.BOF And loRs.EOF) Then

      mIsPswdHistory = loRs!IsPasswordHistory
      If mIsPswdHistory = True Then lblPasswordHistory.Visible = True
      
      mIsPswdIncludeNumbers = loRs!IsPasswordIncludeNumbers
      If mIsPswdIncludeNumbers = True Then lblIncludeNumbers.Visible = True
      
      mIsPswdUseSpecialChar = loRs!IsPasswordUseSpecialCharacters
      If mIsPswdUseSpecialChar = True Then lblUseSpecialCharacters.Visible = True
      
      mIntPswdMinLength = loRs!PasswordMinLength
      If mIntPswdMinLength >= 4 Then
         lblMinimumChar.Caption = lblMinimumChar.Caption & " " & mIntPswdMinLength
         lblMinimumChar.Visible = True
      End If
      
      mIsPswdMixedCaseChar = loRs!IsPasswordMixedCase
      If mIsPswdMixedCaseChar = True Then lblMixedCaseCharacters.Visible = True
      
      mIntPswdFailedAttemts = loRs!PasswordFailedAttempts
      mIsPswdDisallowedList = loRs!IsPasswordDisallowedList
      
   End If
   CloseRecordset loRs
   
   LoadPasswordStrongSettings = True
   
   Exit Function
ErrorHandler:
   LoadPasswordStrongSettings = False
   goSession.RaisePublicError "General Error in msWorkstation.frmPassword.LoadPasswordStrongSettings.", Err.Number, Err.Description
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo SubError
   
   KillObject moKeys
   
   Exit Sub
SubError:
   MsgBox "Error in mwSession.frmPassword.Form_Unload ", Err.Number, Err.Description
End Sub

Private Sub moRsUser_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
      
      If Not loWork.Initialize("mwcUsers") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   If ((goSession.Site.SiteType = SITE_TYPE_SHIP And moKeys.GetKeyValue("IsAllowvesselPwdChangestoFleet") = "1" And BoolNull(moRsUser!IsShoreUser) = False And BlankNull(moRsUser!WorkflowCfgOverride) = "SHIP") Or goSession.Site.SiteType = SITE_TYPE_SHORE) Then
      loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.frmPassword.moRsUser_WillChangeRecord. ", Err.Number, Err.Description
End Sub
