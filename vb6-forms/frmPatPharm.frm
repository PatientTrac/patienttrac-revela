VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmPatPharm 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Patient Pharmacy Information"
   ClientHeight    =   3210
   ClientLeft      =   11040
   ClientTop       =   6480
   ClientWidth     =   6660
   Icon            =   "frmPatPharm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   6660
   Begin VB.Frame fraCmd 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   2070
      TabIndex        =   0
      Top             =   2550
      Width           =   2415
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   495
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPatPharm.frx":094A
         AppearanceThemes=   3
         BorderDrawEvent =   1
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatPharm.frx":1344
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   495
         Left            =   1230
         TabIndex        =   2
         ToolTipText     =   "Close and do not Save"
         Top             =   30
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPatPharm.frx":1CC6
         AppearanceThemes=   3
         BorderDrawEvent =   1
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatPharm.frx":26C0
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmPatPharm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public piPatientID As Long
Public piPatientExtID As Long
Private lintComponentCounter As Integer
Private lintTabCounter As Integer
Private lbooMissingClass As Boolean
Private lstrClassName As String
Private lstrLibrary As String
Private mlngCurrentPatientID As Long
Private mlngCurrentEncounterID As Long
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private mlngFormID As Long
Public mintindex As Integer
Private mFirst As Boolean
Private mstrRemarks As String
Public mbooPatient As Boolean

Private Function Get_EncounterID(vlngPatientID As Long)
Dim strSQL As String
Dim iReturn As Integer
Dim strRS

    strSQL = "SELECT max(ENCOUNTER_ID) from ENCOUNTER WHERE PATIENT_ID = " & vlngPatientID
    iReturn = GetRecordSet(strRS, strSQL, psConnect(0))
    If iReturn > 0 Then
        mlngCurrentEncounterID = Val(strRS(0, 0))
    End If
    Exit Function
End Function

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub


Private Sub cmdSave_Click()
    fobjLib.WriteClinicalData
    Unload Me
    If mbooPatient = False Then
        frmPatients.Show
        frmPatients.piPatientID = piPatientID
    End If
End Sub

Private Sub Form_Activate()
    If mFirst Then
        Screen.MousePointer = vbDefault
        mFirst = False
        centerForm Me
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
End Sub

Private Sub Form_Load()
    mFirst = True
    Load_PatPharm
End Sub


Private Sub Load_PatPharm()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_PatPharm
    
    If piPatientID = 0 Then
        piPatientID = frmPatients.piPatientID
    End If

    strOCX = "FN_PP_PS"
    strControl = "PatPharmacy"
    
    intACTXID = Get_ComponentID(strOCX, strControl)
    
    Set fobjLib = Nothing
    Set fobjOCX = Nothing
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlPatPharmacy", Me)
        fobjOCX.Visible = True
        Me.Width = fobjOCX.Width + 300
        Me.Height = fobjOCX.Height + fraCmd.Height + 1000
        fobjOCX.Left = 100
        Set fobjLib = fobjOCX.Object
        fraCmd.Top = fobjOCX.Height + 250
        fobjOCX.Top = 100
        fraCmd.Left = (Me.Width / 2) - (fraCmd.Width / 2)
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        centerForm Me
        If piPatientID <> 0 Then
            If piProviderID = 0 Then
                Get_DefaultProvider
                piProviderID = mlngDefProviderID
            End If
            fobjLib.ProviderID = piProviderID
            fobjLib.EncounterID = mlngCurrentEncounterID
            Get_EncounterID piPatientID
            fobjLib.PatientID = piPatientID
            modSpecific.piPatientID = piPatientID
            fobjLib.ReadClinicalData
        End If
    Exit Sub
    
Error_Load_PatPharm:

    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If

    ErrHandler _
     "frmPatPharm.Load_PatPharm", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    If Licenses.Count >= 0 Then
        If Licenses.Item(0) <> vbNullString Then
            If Licenses.Item(0) <> vbNullString Then
                Licenses.Remove (0)
            End If
        End If
        Licenses.Clear
    End If
    Set fobjLib = Nothing
    Set fobjOCX = Nothing
    Unload Me
    Exit Sub
End Sub


