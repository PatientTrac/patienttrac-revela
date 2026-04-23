VERSION 5.00
Begin VB.Form frmPracticeReport 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   11310
   ClientLeft      =   8715
   ClientTop       =   2670
   ClientWidth     =   13815
   FillColor       =   &H00800000&
   ForeColor       =   &H00800000&
   Icon            =   "frmPracticeReports.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11310
   ScaleWidth      =   13815
   Begin VB.PictureBox picPatientTrac 
      Height          =   6045
      Left            =   1740
      Picture         =   "frmPracticeReports.frx":09EA
      ScaleHeight     =   5985
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   3510
      Width           =   7695
   End
   Begin VB.Label lblPatientTrac 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "PATIENTTRAC REPORTS"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3105
      Left            =   210
      TabIndex        =   1
      Top             =   60
      Width           =   13455
   End
   Begin VB.Menu mnu_file 
      Caption         =   "File"
      Begin VB.Menu mnu_Save 
         Caption         =   "Save"
      End
      Begin VB.Menu mnu_Print 
         Caption         =   "Print"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnu_Patient 
      Caption         =   "Patients"
   End
   Begin VB.Menu mnu_Run 
      Caption         =   "Run Report"
      Begin VB.Menu mnu_CurRept 
         Caption         =   "Current Report"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_DateSelect 
         Caption         =   "Date Selection"
      End
      Begin VB.Menu mnu_MaxReports 
         Caption         =   "All Reports"
      End
   End
End
Attribute VB_Name = "frmPracticeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public piSelPatientID As Long
Public piSelEnctrID As Long
Public piPatientExtID As Long
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
Private mbooFormLoaded As Boolean
Private mstrDefProvider As String
Private mstrDefFacility As String
Private mstrIniFile As String
Private mbooGet As Boolean
Private mstrReportFileName As String
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
Private mbooNext As Boolean
Private mbooPrev As Boolean
Private Sub Load_Reports()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String
Dim Arr
Dim i%
Dim mbControlPres As Boolean

    On Error GoTo Error_Load_Reports
    piProviderID = clsEncounter.ProviderID

    strOCX = "PatientTrac_Report"
    strControl = "PracticeReports"

    intACTXID = Get_ComponentID(strOCX, strControl)

        Licenses.Add strOCX & "." & strControl
        Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctl" & strControl, Me)
        Set fobjLib = fobjOCX.Object
        fobjOCX.Visible = True
        fobjOCX.Left = 150
        fobjOCX.Top = 10
        Me.Width = fobjOCX.Width + 300
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = piSelPatientID
        fobjLib.EncounterID = piSelEnctrID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        picPatientTrac.ZOrder 1
        Me.Height = fobjOCX.Height + 1000
       ' Me.Height = 10000
    Exit Sub
    
Error_Load_Reports:
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If
    ErrHandler _
     "frmPracticeReports.Load_Reports", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
        centerForm Me
        Screen.MousePointer = vbDefault
    End If
     piSelPatientID = frmPatients.piPatientID
     piSelEnctrID = piEncounterID
    If piSelPatientID = 0 Or piSelEnctrID = 0 Then
         MsgBox "Unable to load Reports, Please select a Patient and Encounter!", vbInformation
         Exit Sub
    End If
    
    If piSelEnctrID > 0 Then Get_Encounter piSelEnctrID
    If piSelPatientID > 0 And piSelEnctrID > 0 Then Load_Reports
    Exit Sub
End Sub


Private Sub Form_Load()
    mFirst = True
    Screen.MousePointer = vbDefault

Exit Sub
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

Private Sub mnu_Close_Click()
    Unload Me
End Sub


Private Sub mnu_Run_Click()
    If piSelEnctrID = 0 Or piSelPatientID = 0 Then
        MsgBox "Cannot run report without current patient or encounter,!", vbCritical
        Exit Sub
    Else
        fobjLib.ReadClinicalData
    End If
    Exit Sub
End Sub


Private Sub mnuPatient_Click()

End Sub


