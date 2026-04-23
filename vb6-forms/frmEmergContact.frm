VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmEmergContact 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Patient Emergency Contact "
   ClientHeight    =   2775
   ClientLeft      =   8325
   ClientTop       =   3900
   ClientWidth     =   6375
   Icon            =   "frmEmergContact.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   6375
   Begin VB.Frame fraSaveClose 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2220
      TabIndex        =   0
      Top             =   2100
      Width           =   1965
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   345
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
         Picture         =   "frmEmergContact.frx":08CA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEmergContact.frx":0C64
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   345
         Left            =   1050
         TabIndex        =   2
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
         Picture         =   "frmEmergContact.frx":0FB6
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEmergContact.frx":19C8
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmEmergContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public piPatientID As Long
Public piProviderID As Long
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
Private Sub cmdClose_Click()
    Unload Me
    'PatientInfo.Show
    PatientInfo.Get_Partner
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    fobjLib.WriteClinicalData
    Unload Me
    PatientInfo.Get_Partner
End Sub

Private Sub Form_Activate()
If mFirst Then
    mFirst = False
    centerForm Me
End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
End Sub

Private Sub Form_Load()

    Me.Top = PatientInfo.Top - 630
    Me.Left = PatientInfo.Left + 1380

    Load_EmergencyContact
    
    mFirst = True
End Sub


Private Sub Load_EmergencyContact()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_EmergencyContact
    
    If piPatientID = 0 Then
        piPatientID = piPatientID
    End If
    strOCX = "FN_PP_PS"
    strControl = "PatEmergencyInfN"
    
    Set fobjLib = Nothing
    Set fobjOCX = Nothing
    
    intACTXID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlPatEmergencyInfo", Me)
        fobjOCX.Visible = True
        Me.Width = fobjOCX.Width + 10
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjOCX.Visible = True
        Me.Height = fobjOCX.Height + 1000
        fraSaveClose.Top = fobjOCX.Height + 30
        fobjLib.PatientID = piPatientID
        fobjLib.EncounterID = piEncounterID
        If piPatientID <> Empty Then
            fobjLib.ProviderID = piProviderID
            fobjLib.ReadClinicalData
            fobjOCX.Visible = True
        End If
    Exit Sub
Error_Load_EmergencyContact:
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If
    MsgBox Err.Description, vbInformation
    Exit Sub
    Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next
 
    
    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    
    For Each obj In frm.objects
      Set obj = Nothing
    Next obj
    Unload Me
End Sub


Private Sub fraPartner_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


