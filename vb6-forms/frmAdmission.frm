VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmAdmission 
   BackColor       =   &H00E0E0E0&
   Caption         =   "HOSPITAL ADMISSION / DISCHARGE"
   ClientHeight    =   9060
   ClientLeft      =   4380
   ClientTop       =   2820
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   15420
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   405
      Left            =   5820
      TabIndex        =   0
      Top             =   8670
      Width           =   3675
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   345
         Left            =   840
         TabIndex        =   1
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
         Picture         =   "frmAdmission.frx":0000
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAdmission.frx":039A
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   345
         Left            =   1890
         TabIndex        =   2
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
         Picture         =   "frmAdmission.frx":06EC
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAdmission.frx":10FE
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public piPatientID As Long
Public piEncounterID As Long
Public piSelEncntrID As Long
Public piProviderID As Long
Public piFacilityID As Long
Public piPatientExtID As Long
Public mbooNext As Boolean
Public mbooPrev As Boolean
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
Private mFirst As Boolean
Private mstrRemarks As String
Private mbooCancelAdmission As Boolean

Public Sub ReleaseObject(frm As Form)
   Set frm = Nothing
End Sub

Private Sub cmdClose_Click()
    Dim iAnswer As Integer
    
    On Error GoTo Error_Handler
    
    ' Check for Superbill Open
    If IsFormLoaded("frmSuperbill") Then
        Unload frmSuperBill
    End If
    If IsFormLoaded("frmPsychiatry") Then
       Unload frmPsychiatry
    End If
    ReleaseObject Me
    Unload Me
    frmPatients.Form_Load
    frmPatients.Show
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmAdmission.cmdClose_Click", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub cmdSave_Click()
Dim strSQL As String
Dim booret As Boolean
Dim varRS
Dim lngrecs As Long
Dim intResp As Byte

    fobjLib.WriteClinicalData
    piEncounterID = fobjLib.EncounterID
    modSpecific.piEncounterID = piEncounterID
    Get_Encounter piEncounterID
    Get_FacilityID
    piFacilityID = modSpecific.piFacilityID
    frmPatients.piFacilityID = piFacilityID

    
    ' Check to see encounter update
    strSQL = "select ENCNTR_TYPE, SERVICE_PLACE  " _
            & "From encounter " _
            & "Where patient_id = " & piPatientID & " and ENCOUNTER_ID = " & piEncounterID
    booret = SelectQuery(strSQL, psconnect(0), varRS, lngrecs)
    
    If booret Then
        If lngrecs > 0 Then
            clsEncounter.ServicePlace = varRS(1, 0)
            clsEncounter.EncounterType = varRS(0, 0)
        End If
    End If
    If clsEncounter.EncounterType = "Initial InPatient" Or clsEncounter.EncounterType = "Discharge" Then
        MsgBox "Do you want to Print " & clsEncounter.EncounterType & " Report? ", vbYesNo
        If intResp = vbYes Then
            If clsEncounter.EncounterType = "Initial InPatient" Then
                frmMDIReport.fstrLastFilePath = "C:\Program Files\FirstNotebook\PatientTrac EMR\Reports"
                frmMDIReport.fmstrfile = frmMDIReport.fstrLastFilePath & "\" & "Admission_1100.Txt"
            ElseIf clsEncounter.EncounterType = "Discharge" Then
                frmMDIReport.fstrLastFilePath = "C:\Program Files\FirstNotebook\PatientTrac EMR\Reports"
                frmMDIReport.fmstrfile = frmMDIReport.fstrLastFilePath & "\" & "Discharge_1100.Txt"
            End If
            frmMDIReport.Show
            frmMDIReport.cmdRun_Click
            frmMDIReport.cmdPrint_Click
       ElseIf intResp = vbNo Then
            MsgBox "Please enter Active Reports to Print " & clsEncounter.EncounterType & " Reports ", vbInformation
        End If
    End If
    frmEnctr.Init_Summary
    Unload Me
    Exit Sub
End Sub
Private Sub Form_Activate()
   If mFirst = True Then
        mFirst = False
   End If
   Screen.MousePointer = vbDefault
   Exit Sub
End Sub

Private Sub Form_Load()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    mlngLastPatientId = 0
    mlngLastEncounterId = 0
    centerForm Me
    mFirst = True
    Me.Top = frmMDI.Top + 100
    
   mbooPrev = False
   mbooNext = False
   
    LogFile "Start Log  Admission --------------------", True
        
        Load_Admission
        If mbooCancelAdmission Then Unload Me
        centerForm Me
        Screen.MousePointer = vbDefault
Exit Sub
End Sub
Private Sub Load_Admission()
Dim strOCX As String
Dim strControl As String
Dim intActxID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_Admission

    
    piSelEncntrID = piEncounterID
    
    Validate_PiValues piPatientID, piSelEncntrID, piProviderID, piFacilityID
    If clsEncounter.VisitType = vbNullString Then
        MsgBox "Application requires visit type, Please update encounter with a Visit Type!", vbInformation
        mbooCancelAdmission = True
        Exit Sub
    End If
    strOCX = "FN_PsyMse"
    strControl = "Admission"
    
    intActxID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlAdmission", Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 150
        fobjOCX.Top = Me.Top - 60
        'fraControlPanel.Left = fobjOCX.Width + 125
        frmAdmission.Width = fobjOCX.Width + 500
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intActxID
        fobjLib.ConnectStringTerms = psconnect(1)
        fobjLib.ConnectStringClinical = psconnect(0)
       ' fobjLib = psUserID
        Me.Height = fobjOCX.Height + 1000
        Frame1.Top = fobjOCX.Height + 30
        centerForm Me
        fobjLib.PatientID = piPatientID
        fobjLib.EncounterID = piEncounterID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        Screen.MousePointer = vbDefault
     
    Exit Sub

Error_Load_Admission:

    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If

    ErrHandler _
     "frmAdmission.Load_Admission", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set frmAdmission = Nothing
End Sub





Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.Objects
      Set obj = Nothing
    Next obj
End Sub


