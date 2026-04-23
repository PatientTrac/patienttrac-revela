VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmScan 
   Caption         =   "Patient Scanned Files"
   ClientHeight    =   11130
   ClientLeft      =   7455
   ClientTop       =   2610
   ClientWidth     =   15210
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11130
   ScaleWidth      =   15210
   Begin VB.Frame fraButton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7500
      TabIndex        =   0
      Top             =   10650
      Width           =   945
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   345
         Left            =   0
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
         Picture         =   "frmScan.frx":0ECA
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmScan.frx":18DC
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private piSelPatientID As Long
Private piSelEnctrID As Long
Private piProviderID As Long
Private piPatientExtID As Long
Private mintProviderID As Integer
Private msEncounterDate As String
Private mstrProviderName As String
Private lintComponentCounter As Integer
Private lintTabCounter As Integer
Private lbooMissingClass As Boolean
Private lstrClassName As String
Private lstrLibrary As String
Private mlngCurrentPatientID As Long
Private mlngCurrentEncounterID As Long
Private mstrGridName As String
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private mlngFormID As Long
Private mstrIniFile As String
Private mboobypass As Boolean
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
Public EnctrType As String
Private mbooNoData As Boolean
Private mlngSelFacID As Long
Private mintGridRow As Integer
Private mintGridCol As Integer
Public mstrLoadType As String


Private Sub Load_ScanImage()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_Form
   ' vstrOcx
   strOCX = "FN_RevSys"
   strControl = "ScanImages"
    intACTXID = Get_ComponentID(strOCX, strControl)
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctl" & strControl, Me)
        fobjOCX.Visible = True
        Me.Width = fobjOCX.Width + 1000
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjOCX.Top = 30
        Me.Height = fobjOCX.Height + 300
        fraButton.Top = fobjOCX.hieght + 100
        fobjLib.PatientID = piSelPatientID
        fobjLib.ReadClinicalData
        
        Me.Height = fobjOCX.Height + 1000
        Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Form:
Screen.MousePointer = vbDefault
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub Form_Activate()
    On Error GoTo Error
    If mFirst Then
        mFirst = False
    End If
    If frmPatients.piPatientID <> 0 Then
        piSelPatientID = frmPatients.piPatientID
    Else
         piSelPatientID = frmPatientAppointment.mlngPatientID
    End If
    Load_ScanImage
    centerForm Me
    Screen.MousePointer = vbDefault
    Exit Sub
Error:
    Exit Sub
    Resume
End Sub

Private Sub Form_Load()
    mFirst = True
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
  Exit Sub
End Sub


