VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwUserSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'mwUserSecurity
' 8/3/2001


Option Explicit

Private mAllowTransportOverride As Boolean
Private mAllowEmailCarrierOverride As Boolean
Private mAllowEmailAddressOverride As Boolean
Private mAllowTransmitContainer As Boolean

Private mUserRegistryAccess As msUserRegistryAccess
Private mUserConfigAccess As String
Private mUserConfigPreferencesAccess As String
Private mUserConfigWorkflowAccess As String
Private mUserConfigTransportAccess As String
Private mUserConfigAdvancedAccess As String
Private mUserConfigLicensingAccess As String
Private mUserConfigLicenseDistribution As Boolean

Private moParent As mwUser

Public Property Set Parent(oParent As mwUser)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As mwUser
   Set Parent = moParent
End Property


Public Property Let AllowTransportOverride(ByVal vData As Boolean)
   mAllowTransportOverride = vData
   moParent.UpdateUserData "AllowTransportOverride", vData
End Property

Public Property Get AllowTransportOverride() As Boolean
    AllowTransportOverride = mAllowTransportOverride
End Property

Public Property Let AllowTransmitContainer(ByVal vData As Boolean)
   mAllowTransmitContainer = vData
   moParent.UpdateUserData "AllowTransmitContainer", vData
End Property

Public Property Get AllowTransmitContainer() As Boolean
    AllowTransmitContainer = mAllowTransmitContainer
End Property

Public Property Get UserConfigAccess() As String
    UserConfigAccess = mUserConfigAccess
End Property

Public Property Let UserConfigAccess(ByVal vData As String)
    mUserConfigAccess = vData
   moParent.UpdateUserData "UserConfigAccess", vData
End Property


Public Property Get UserConfigPreferencesAccess() As String
    UserConfigPreferencesAccess = mUserConfigPreferencesAccess
End Property

Public Property Let UserConfigPreferencesAccess(ByVal vData As String)
    mUserConfigPreferencesAccess = vData
   moParent.UpdateUserData "UserConfigPreferencesAccess", vData
End Property


Public Property Get UserConfigWorkflowAccess() As String
    UserConfigWorkflowAccess = mUserConfigWorkflowAccess
End Property

Public Property Let UserConfigWorkflowAccess(ByVal vData As String)
    mUserConfigWorkflowAccess = vData
   moParent.UpdateUserData "UserConfigWorkflowAccess", vData
End Property


Public Property Get UserConfigTransportAccess() As String
    UserConfigTransportAccess = mUserConfigTransportAccess
End Property

Public Property Let UserConfigTransportAccess(ByVal vData As String)
    mUserConfigTransportAccess = vData
   moParent.UpdateUserData "UserConfigTransportAccess", vData
End Property


Public Property Get UserConfigAdvancedAccess() As String
    UserConfigAdvancedAccess = mUserConfigAdvancedAccess
End Property

Public Property Let UserConfigAdvancedAccess(ByVal vData As String)
    mUserConfigAdvancedAccess = vData
   moParent.UpdateUserData "UserConfigAdvancedAccess", vData
End Property


Public Property Get UserConfigLicensingAccess() As String
    UserConfigLicensingAccess = mUserConfigLicensingAccess
End Property

Public Property Let UserConfigLicensingAccess(ByVal vData As String)
    mUserConfigLicensingAccess = vData
   moParent.UpdateUserData "UserConfigLicensingAccess", vData
End Property


Public Property Get UserConfigLicenseDistribution() As Boolean
    UserConfigLicenseDistribution = mUserConfigLicenseDistribution
End Property

Public Property Let UserConfigLicenseDistribution(ByVal vData As Boolean)
    mUserConfigLicenseDistribution = vData
End Property


Public Property Get AllowEmailCarrierOverride() As Boolean
    AllowEmailCarrierOverride = mAllowEmailCarrierOverride
End Property

Public Property Let AllowEmailCarrierOverride(ByVal vData As Boolean)
    mAllowEmailCarrierOverride = vData
   moParent.UpdateUserData "AllowEmailCarrierOverride", vData
End Property
'
Public Property Get UserRegistryAccess() As msUserRegistryAccess
    UserRegistryAccess = mUserRegistryAccess
End Property

Public Property Get AllowEmailAddressOverride() As Boolean
    AllowEmailAddressOverride = mAllowEmailAddressOverride
End Property

Public Property Let AllowEmailAddressOverride(ByVal vData As Boolean)
    mAllowEmailAddressOverride = vData
   moParent.UpdateUserData "AllowEmailAddressOverride", vData
End Property

