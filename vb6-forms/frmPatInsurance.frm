VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmPatInsurance 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Patient Insurance"
   ClientHeight    =   10755
   ClientLeft      =   5625
   ClientTop       =   1950
   ClientWidth     =   13380
   Icon            =   "frmPatInsurance.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10755
   ScaleWidth      =   13380
   Begin VB.Frame fraButton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   11280
      TabIndex        =   4
      Top             =   60
      Width           =   2055
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   405
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   "Save, Restore, Close"
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
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
         Picture         =   "frmPatInsurance.frx":08CA
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatInsurance.frx":12DC
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   405
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "Save, Restore, Close"
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
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
         Picture         =   "frmPatInsurance.frx":162E
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatInsurance.frx":19C8
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.Frame fraPrimary 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PRIMARY INSURANCE"
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
      Height          =   5055
      Left            =   30
      TabIndex        =   1
      Top             =   600
      Width           =   13305
   End
   Begin VB.Frame fraSecIns 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SECONDARY INSURANCE"
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
      Height          =   5055
      Left            =   30
      TabIndex        =   0
      Top             =   5670
      Width           =   13305
   End
   Begin PT_XP_Button.PT_XP cmdRevEnctr 
      Height          =   525
      Left            =   2940
      TabIndex        =   2
      ToolTipText     =   "Previous, Next, or Current Encounter"
      Top             =   30
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   926
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
      Picture         =   "frmPatInsurance.frx":1D1A
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Employer"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatInsurance.frx":25F4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdMaint 
      Height          =   525
      Left            =   30
      TabIndex        =   3
      ToolTipText     =   "Save, Restore, Close"
      Top             =   30
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   926
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
      Picture         =   "frmPatInsurance.frx":3246
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Insurance Company Maintanence"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   12582912
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatInsurance.frx":3B20
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmPatInsurance"
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
Private mbooNext As Boolean

 
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMaint_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "Insurers"
    frmLoadControl.Show vbModal
    Exit Sub
End Sub




Private Sub cmdRevEnctr_Click()
    mstrPatOcx = "FN_PP_PS"
    mstrPatCtl = "Employer"
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    fobjLib.WriteClinicalData
    SecobjLib.WriteClinicalData
    Unload Me
End Sub


Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
    End If
    
    Load_PatInsurance
    centerForm Me
    Exit Sub
End Sub

Private Sub Form_Load()

    mFirst = True
    

    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmPatInsurance.Form_Load", _
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



Private Sub Load_PatInsurance()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String
Dim strUser As String
Dim strSQL As String
Dim booRet As Boolean
Dim lngrecs As Long
Dim varRS


    On Error GoTo Error_Load_PatInsurance
    Screen.MousePointer = vbHourglass
    
    Set fobjLib = Nothing
    Set fobjOCX = Nothing
    Set mtypelib = Nothing
    Set mobj = Nothing
    
    strOCX = "FN_PP_PS"
    strControl = "PrimIns"
    
    intACTXID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlPrimIns", fraPrimary)
        fobjOCX.Visible = True
        fobjOCX.Left = 0
        fobjOCX.Top = 50
    Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = mlngPatientID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        Set SecobjLib = Nothing
        Set SecobjOCX = Nothing
    strOCX = "FN_PP_PS"
    strControl = "SecondaryIns"
    
    intACTXID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set SecobjOCX = Controls.Add(strOCX & "." & strControl, "ctlSecondaryIns", fraSecIns)
        SecobjOCX.Visible = True
        SecobjOCX.Left = 0
        SecobjOCX.Top = 50
    Set SecobjLib = SecobjOCX.Object
        SecobjLib.ComponentID = intACTXID
        SecobjLib.ConnectStringTerms = psConnect(1)
        SecobjLib.ConnectStringClinical = psConnect(0)
        SecobjLib.PatientID = piPatientID
        SecobjLib.ProviderID = piProviderID
        SecobjLib.ReadClinicalData
        centerForm Me
     Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_PatInsurance:
    Screen.MousePointer = vbDefault
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then
        Resume Next
    ElseIf Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    Else
        Write_Error Err.Description, Err.Number, "", "Load_PatIns", "frmPatIns"
        Screen.MousePointer = vbDefault
        Exit Sub
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
  Exit Sub
End Sub


 
Private Sub Save_Click()
    fobjLib.WriteClinicalData
    SecobjLib.WriteClinicalData
    Unload Me
End Sub

