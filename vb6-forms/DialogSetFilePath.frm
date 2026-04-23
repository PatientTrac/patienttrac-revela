VERSION 5.00
Begin VB.Form frmDialogSetFilePath 
   Caption         =   "Application Settings"
   ClientHeight    =   4965
   ClientLeft      =   7710
   ClientTop       =   3720
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7065
   Begin VB.CheckBox chkShowStimSheetIcon 
      Caption         =   "Show StimSheet Icon in Toolbar"
      Height          =   285
      Left            =   3690
      TabIndex        =   19
      Top             =   3690
      Width           =   3045
   End
   Begin VB.Frame fraEncounterOpt 
      Caption         =   " Encounter Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   3690
      TabIndex        =   13
      Top             =   1080
      Width           =   3105
      Begin VB.CheckBox chkReqPayType 
         Caption         =   "Require Pay Type"
         Height          =   345
         Left            =   270
         TabIndex        =   18
         Top             =   990
         Width           =   1815
      End
      Begin VB.CheckBox chkUnlimitedEnctrs 
         Caption         =   "Unlimited New Encounters      (For Maintenance Only)"
         Height          =   405
         Left            =   270
         TabIndex        =   17
         Top             =   1410
         Width           =   2325
      End
      Begin VB.CheckBox chkServicePlace 
         Caption         =   "Require Service Place"
         Height          =   225
         Left            =   270
         TabIndex        =   16
         Top             =   360
         Width           =   2025
      End
      Begin VB.CheckBox chkReqEnctrType 
         Caption         =   "Require Encounter Type"
         Height          =   315
         Left            =   270
         TabIndex        =   15
         Top             =   660
         Width           =   2145
      End
      Begin VB.CheckBox chkAllowClosingAnyEnctr 
         Caption         =   "Allow Closing Old Encounters      (For Maintenance Only)"
         Height          =   495
         Left            =   270
         TabIndex        =   14
         Top             =   1860
         Width           =   2415
      End
   End
   Begin VB.Frame fraIntervention 
      Caption         =   " Intervention Display Options "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   330
      TabIndex        =   8
      Top             =   1080
      Width           =   3105
      Begin VB.CheckBox chkBilling 
         Caption         =   "Billing Approval"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   1530
         Width           =   2145
      End
      Begin VB.CheckBox chkShowTransCode 
         Caption         =   "Show Transaction Code"
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   1140
         Width           =   2145
      End
      Begin VB.CheckBox chkShowCPTMod 
         Caption         =   "Show CPT Modifier"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   780
         Width           =   1815
      End
      Begin VB.CheckBox chkShowFee 
         Caption         =   "Show Fee"
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   390
         Width           =   1065
      End
   End
   Begin VB.Frame fraLogon 
      Caption         =   " Logon Options "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   330
      TabIndex        =   5
      Top             =   3150
      Width           =   3105
      Begin VB.CheckBox chkReqPWD 
         Caption         =   "Require User Logon and Password"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   2835
      End
      Begin VB.CheckBox chkHidePasswordChange 
         Caption         =   "Hide Password Change"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2205
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3450
      TabIndex        =   4
      Top             =   4350
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   345
      Index           =   0
      Left            =   6360
      TabIndex        =   3
      Top             =   570
      Width           =   405
   End
   Begin VB.TextBox txtPath 
      Height          =   330
      Left            =   345
      TabIndex        =   1
      Text            =   "\\"
      Top             =   585
      Width           =   5910
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2130
      TabIndex        =   0
      Top             =   4350
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Set File Path for Consent Forms (full network path if known):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   345
      TabIndex        =   2
      Top             =   315
      Width           =   5145
   End
End
Attribute VB_Name = "frmDialogSetFilePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fsFilePath As String
Private fsCurrentDrive As String

Private Sub CancelButton_Click()
   Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub chkDisablePasswordChange_Click()

End Sub

Private Sub chkReqPatientType_Click()

End Sub

Private Sub cmdSave_Click()

   If Trim(txtPath.Text) = vbNullString Then
      MsgBox "Scanned Image Path Required", vbExclamation, App.Title
      Exit Sub
   End If

   If Not FolderExists(Trim(txtPath.Text)) Then
      MsgBox "Scanned Image Path Does Not Exists", vbExclamation, App.Title
      Exit Sub
   End If

'   SaveSetting appname:="PatientProfiler", Section:="General", Key:="ConsentFilePath", setting:=txtPath
   SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ImagePath", txtPath, REG_SZ
   psImageFilePath = txtPath
   
'Logon Options.......................................................................
   If chkReqPWD.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "PasswordRequired", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "PasswordRequired", "no", REG_SZ
   End If
   
   
   If chkHidePasswordChange.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "HideChangePassword", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "HideChangePassword", "no", REG_SZ
   End If
   
   
'Billing Aproval Options.......................................................................
   If chkShowFee.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ShowFee", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ShowFee", "no", REG_SZ
   End If
   
   If chkShowCPTMod.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ShowModifier", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ShowModifier", "no", REG_SZ
   End If
   
   If chkShowTransCode.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ShowTransCode", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ShowTransCode", "no", REG_SZ
   End If
   
   If chkBilling.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ShowBillingApproval", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ShowBillingApproval", "no", REG_SZ
   End If
    
    
'Encounter Options............................................................................

    'Place of Service............................................................................
   If chkServicePlace.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ServicePlaceReq", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ServicePlaceReq", "no", REG_SZ
   End If
   
   'Encounter Type............................................................................
   If chkReqEnctrType.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "EncounterTypeRequired", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "EncounterTypeRequired", "no", REG_SZ
   End If
   
  'Pay Type............................................................................
   If chkReqPayType.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "PayTypeRequired", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "PayTypeRequired", "no", REG_SZ
   End If
   
   
'Unlimited New Encounters............................................................................
   If chkUnlimitedEnctrs.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "UnlimitedEncounters", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "UnlimitedEncounters", "no", REG_SZ
   End If
   
'Allow old Encounter Closing............................................................................
   If chkAllowClosingAnyEnctr.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "CloseAnyEncounter", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "CloseAnyEncounter", "no", REG_SZ
   End If
   
'Show StimSheeet...........................................................................
   If chkShowStimSheetIcon.Value Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ShowStimSheetIcon", "yes", REG_SZ
   Else
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ShowStimSheetIcon", "no", REG_SZ
   End If
   
    If chkShowStimSheetIcon.Value Then
        frmMDI.Toolbar.buttons(5).Visible = True
    Else
        frmMDI.Toolbar.buttons(5).Visible = False
    End If
   
   Unload Me
End Sub
Private Sub Command1_Click(Index As Integer)
   Dim lstrPath As String
   'ChooseFolder = modSpecific.ChooseFolder
   lstrPath = ChooseFolder("Set File Path for Scanned Imagess")
   If Trim(lstrPath) <> vbNullString Then txtPath.Text = lstrPath
      
End Sub

Private Sub Form_Load()
Dim strOption As String


'Logon Options.......................................................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "PasswordRequired")
    If strOption = "yes" Then chkReqPWD.Value = 1
    
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "HideChangePassword")
    If strOption = "yes" Then chkHidePasswordChange.Value = 1



'Billing Aproval Options.......................................................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ShowFee")
    If strOption = "yes" Then chkShowFee.Value = 1
    
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ShowModifier")
    If strOption = "yes" Then chkShowCPTMod.Value = 1
    
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ShowTransCode")
    If strOption = "yes" Then chkShowTransCode.Value = 1
    
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ShowBillingApproval")
    If strOption = "yes" Then chkBilling.Value = 1
    
    
    
'Place of Service............................................................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ServicePlaceReq")
    If strOption = "yes" Then chkServicePlace.Value = 1
    
'Encounter Type............................................................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "EncounterTypeRequired")
    If strOption = "yes" Then chkReqEnctrType.Value = 1
    
'Pay Type............................................................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "PayTypeRequired")
    If strOption = "yes" Then chkReqPayType.Value = 1
        
'Unlimited New Encounters............................................................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "UnlimitedEncounters")
    If strOption = "yes" Then chkUnlimitedEnctrs.Value = 1
    
'Allow old Encounter Closing............................................................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "CloseAnyEncounter")
    If strOption = "yes" Then chkAllowClosingAnyEnctr.Value = 1
        
'Allow old Encounter Closing............................................................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ShowStimSheetIcon")
    If strOption = "yes" Then chkShowStimSheetIcon.Value = 1
    
    centerForm Me
    
    txtPath.Text = psImageFilePath
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
End Sub


