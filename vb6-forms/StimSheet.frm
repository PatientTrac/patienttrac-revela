VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmStimSheet 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Stimulation Sheet"
   ClientHeight    =   9510
   ClientLeft      =   7320
   ClientTop       =   2805
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   15240
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   675
      Left            =   5520
      TabIndex        =   0
      ToolTipText     =   "Review Scanned Chart Data"
      Top             =   8760
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1191
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "StimSheet.frx":0000
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   12582912
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "StimSheet.frx":08DA
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   675
      Left            =   7560
      TabIndex        =   1
      ToolTipText     =   "Save, Restore, Close"
      Top             =   8760
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1191
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
      Picture         =   "StimSheet.frx":152C
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   12582912
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "StimSheet.frx":1E06
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmStimSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private mlngFormID As Long
Private Sub cmdClose_Click()
    Controls.Remove fobjOCX
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdSave_Click()
    fobjLib.WriteClinicalData
End Sub


Private Sub Form_Activate()
    Load_StimSheet
End Sub

Private Sub Form_Load()
    centerForm Me
    Screen.MousePointer = vbDefault
End Sub



Private Sub Load_StimSheet()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_StimSheet
    
    strOCX = "FN_Reproductive"
    strControl = "StimSheet"
    
    intACTXID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlSS", Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 30
        fobjOCX.Top = 50
        
    Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = piPatientID
        fobjLib.EncounterID = piEncounterID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        Me.Height = fobjOCX.Height * 1
        Me.Width = fobjOCX.Width + 300
        cmdSave.Top = fobjOCX.Height + 30
        cmdClose.Top = fobjOCX.Height + 30
        centerForm Me
    Exit Sub
    
Error_Load_StimSheet:

    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If

     ErrHandler _
     Me.Name & " Load_StimSheet", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Public Property Get FormID() As Long
    FormID = mlngFormID
End Property

Public Property Let FormID(ByVal vNewValue As Long)
    mlngFormID = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next
    Unload Me
    Exit Sub
End Sub


