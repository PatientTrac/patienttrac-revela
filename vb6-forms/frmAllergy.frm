VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmAllergy 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Patient Allergies"
   ClientHeight    =   10590
   ClientLeft      =   3615
   ClientTop       =   1800
   ClientWidth     =   13575
   ControlBox      =   0   'False
   Icon            =   "frmAllergy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   345
      Left            =   1110
      TabIndex        =   1
      Top             =   30
      Width           =   975
      _ExtentX        =   1720
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
      Picture         =   "frmAllergy.frx":058A
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmAllergy.frx":0F9C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   30
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
      Picture         =   "frmAllergy.frx":12EE
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmAllergy.frx":1688
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmAllergy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mPatientID As Long
Public mPatientName As String
Public mFirst As Boolean
Public mstrConnectStringClinical As String
Public mstrConnectStringTerms As String
Public piPatientID As Long
Public piEncounterID As Long
Public piSelEncntrID As Long
Public piProviderID As Long
Public piPatientExtID As Long
Public piFacilityID As Long
Public mintProviderID As Integer
Private msEncounterDate As String
Public mstrProviderName As String
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
'dnn.................
Private mintTabs As Long
Private mintTabsPerRow As Long
Private fstrTermsConnect As String
Private fstrClinicalConnect As String
Private fstrReportDate As Date
Private fstrReportDate2 As Date
Private mbooFormLoaded As Boolean
Private mstrDefProvider As String
Private mstrDefFacility As String
Private mstrIniFile As String
Private mbooGet As Boolean
Private mstrReportFileName As String
'*** Note!   This collection below is a rudimentary first attempt to handle this.
'            We need to create a specific collection that allows to manage the
'            loaded controls on a page by page basis, instead of keeping all
'            controls for all pages in the same bag.
Private mbooformload As Boolean
'Dimension collection to keep track of loaded clinical components
Private mcolActxComp As New Collection
Private mbooSelectPat As Boolean
Private mbooRecurssive As Boolean
Private mlngLastPatientId As Long
Private mlngLastEncounterId As Long
Private mObjSelectPatient As PatientTracTypeLib.iClinical
'Private WithEvents PatientList As SelectPatient
Public mintindex As Integer



Private Sub Load_Form()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_Form

    piPatientID = modSpecific.piPatientID
    piEncounterID = modSpecific.piEncounterID
    piFacilityID = modSpecific.piFacilityID
    piProviderID = modSpecific.piProviderID
    piEncounterID = modSpecific.piEncounterID

    Validate_PiValues piPatientID, piEncounterID, piProviderID, piFacilityID


   ' vstrOcx
   strOCX = "PT_MEDICATIONS"
   strControl = "NALLERGY"
    
    intACTXID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctl" & strControl, Me)
        fobjOCX.Visible = True
        'fraControlPanel.Left = fobjOCX.Width + 125
        Me.Width = fobjOCX.Width + 1000
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
       ' fobjOCX.Left = 150
        fobjOCX.Top = 30
        fobjLib.PatientID = piPatientID
        fobjLib.EncounterID = piEncounterID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        
        Me.Height = fobjOCX.Height + 1000
    Exit Sub
    
Error_Load_Form:

    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If

    ErrHandler _
     "frmTeleNotes.Load_Form", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub cmdClose_Click()
    Dim iAnswer As Integer
    
    On Error GoTo Error_Handler
    

    
    Unload Me
    
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmAllergy.cmdClose_Click", _
     Err.Number, _
     Err.Description
     Resume
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Error_Handler
    fobjLib.WriteClinicalData
    Unload Me
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmAllergy.cmdSave_Click", _
     Err.Number, _
     Err.Description
     Resume
    Exit Sub
End Sub

Private Sub Form_Activate()
    On Error GoTo Error_Handler
    
    If mFirst Then
        mFirst = False
        piPatientID = mPatientID
        If mstrConnectStringClinical = vbNullString Then mstrConnectStringClinical = psConnect(0)
        If mstrConnectStringTerms = vbNullString Then mstrConnectStringTerms = psConnect(1)
         Load_Form
    End If
    centerForm Me
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmAllergy.Form_Activate", _
     Err.Number, _
     Err.Description
     Resume
    Exit Sub
End Sub
Private Sub Form_Load()
    mFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl

End Sub


