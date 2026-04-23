VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmEmployer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Employer"
   ClientHeight    =   5700
   ClientLeft      =   7365
   ClientTop       =   6300
   ClientWidth     =   9975
   Icon            =   "frmEmployer.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5700
   ScaleWidth      =   9975
   Begin VB.Frame fraPrimary 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9915
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   405
      Left            =   4890
      TabIndex        =   1
      ToolTipText     =   "Save, Restore, Close"
      Top             =   5220
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
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
      Picture         =   "frmEmployer.frx":058A
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmEmployer.frx":0F9C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   405
      Left            =   3690
      TabIndex        =   2
      ToolTipText     =   "Save, Restore, Close"
      Top             =   5220
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   714
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
      Picture         =   "frmEmployer.frx":12EE
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmEmployer.frx":1688
      ShowFocus       =   -1  'True
   End
   Begin VB.Menu Patient 
      Caption         =   "Patient Data"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu Select_Patient 
         Caption         =   "Select Patient"
         Shortcut        =   ^P
      End
      Begin VB.Menu Insurance 
         Caption         =   "Insurance"
         Checked         =   -1  'True
      End
      Begin VB.Menu Empl_Info 
         Caption         =   "Employer Information"
      End
   End
   Begin VB.Menu Encounter 
      Caption         =   "Encounter"
      Visible         =   0   'False
      Begin VB.Menu Prev_Enctr 
         Caption         =   "Previous Encounter"
      End
      Begin VB.Menu Next_Enctr 
         Caption         =   "Next Encounter"
      End
      Begin VB.Menu Current_Enctr 
         Caption         =   "Current Encounter"
      End
   End
   Begin VB.Menu Facility 
      Caption         =   "Facility Maintanence"
      Visible         =   0   'False
      Begin VB.Menu Manage_Facilities 
         Caption         =   "Manage Facilities"
      End
      Begin VB.Menu Manage_Providers 
         Caption         =   "Manage Providers"
      End
   End
   Begin VB.Menu Clinical_Data 
      Caption         =   "Clinical Data"
      Visible         =   0   'False
      Begin VB.Menu Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmEmployer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public mOCX As String
Public mControl As String
Public mlngPatientID As Long
Public mlngEncounterID As Long
Public mlngProviderID As Long
Public mstrCPTCode As String
Public mstrConnectStringClinical As String
Public mstrConnectStringTerms As String
Public mbControlLoaded As Boolean
Public mbSaved As Boolean
Public mstrParentRef As String
Public mstrSummary As String
Public mlngComponentID As Long
Private mFirst As Boolean
Private mobj As Object
Private mtypelib As PatientTracTypeLib.iClinical
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private SecobjLib As PatientTracTypeLib.iClinical ' Secondary Insurance
Private SecobjOCX As VBControlExtender ' Secondary Insurance
Private mlngFormID As Long
Private mbooPrev As Boolean

Private Sub Close_Click()

    Unload Me
End Sub


Private Sub cmdClose_Click()
        Unload Me
End Sub

Private Sub cmdSave_Click()
    fobjLib.WriteClinicalData
    Unload Me
End Sub


Private Sub cmdSave_DropDownClick()

    PopupMenu Clinical_Data, , cmdSave.Left, cmdSave.Height
    
End Sub

Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
    End If

End Sub

Private Sub Form_Load()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

    mFirst = True
    
    Load_Employer
    centerForm Me
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmSuperbill.Form_Load", _
     Err.Number, _
     Err.Description
     Resume
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set fobjLib = Nothing
    Set fobjOCX = Nothing
    Set SecobjLib = Nothing
    Set SecobjOCX = Nothing
End Sub



Private Sub Load_Employer()
Dim strOCX As String
Dim strControl As String
Dim intActxID As Integer
Dim strRemarks As String
Dim strUser As String
Dim strSQL As String
Dim booret As Boolean
Dim lngrecs As Long
Dim varRS


    On Error GoTo Error_Load_Employer

    Set fobjLib = Nothing
    Set fobjOCX = Nothing
    Set mtypelib = Nothing
    Set mobj = Nothing
    
    strOCX = "FN_PP_PS"
    strControl = "Employer"
    
    intActxID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlEmployer", fraPrimary)
        fobjOCX.Visible = True
        fobjOCX.Left = 0
        fobjOCX.Top = 500
        frmEmployer.Width = fobjOCX.Width + 300
    Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intActxID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = piPatientID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData

        centerForm Me
  Exit Sub
Error_Load_Employer:

    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If

    ErrHandler _
     "frmEmployer.Load_Employer", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume

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
  Exit Sub
End Sub


Private Sub Manage_Facilities_Click()

    frmFacilities.Show vbModal
    
End Sub

Private Sub Manage_Providers_Click()
    frmProviders.Show vbModal
End Sub


Private Sub Save_Click()
    fobjLib.WriteClinicalData
    SecobjLib.WriteClinicalData
    Unload Me
End Sub

