VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmProvInsurance 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Provider Insurance Accounts"
   ClientHeight    =   12105
   ClientLeft      =   3195
   ClientTop       =   1770
   ClientWidth     =   11820
   Icon            =   "frmProvInsurance.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   12105
   ScaleWidth      =   11820
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   525
      Left            =   5970
      TabIndex        =   0
      ToolTipText     =   "Save, Restore, Close"
      Top             =   11310
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   926
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
      Picture         =   "frmProvInsurance.frx":058A
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmProvInsurance.frx":0E64
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   525
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Save, Restore, Close"
      Top             =   11310
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   926
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
      Picture         =   "frmProvInsurance.frx":1AB6
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmProvInsurance.frx":2390
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmProvInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public mOCX As String
Public mControl As String
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
Private mlngFormID As Long

Private Function Get_EncounterID()
Dim strSQL As String
Dim sDate As String
Dim varRS
Dim lngrecs As Long
Dim booret As Boolean


    If piEncounterID = 0 Then
        strSQL = "select max(encounter_id) from encounter where  " _
                & "patient_id = " & piPatientID
        booret = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
        
        If booret Then
            If lngrecs > 0 Then
                piEncounterID = varRS(0, 0)
            End If
        End If
    End If
    Exit Function
End Function

Private Sub Load_InsProv()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_InsProv

    strOCX = "FN_RevSys"
    strControl = "InsuranceProvider"
    
    intACTXID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlInsuranceProvider", Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 150
        fobjOCX.Top = 100
        'fraControlPanel.Left = fobjOCX.Width + 125
        Me.Width = fobjOCX.Width + 300
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = piPatientID
        If piEncounterID = 0 Then
            Get_EncounterID
        End If
        fobjLib.EncounterID = piEncounterID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        
        Me.Height = fobjOCX.Height * 0.95
    Exit Sub
    
Error_Load_InsProv:

    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If

    ErrHandler _
     "frmProvInsurance.Load_InsProv", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub cmdClose_Click()

    
    Unload Me
End Sub

Private Sub cmdSave_Click()

    Unload Me
End Sub


Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
    End If

End Sub

Private Sub Form_Load()

    mFirst = True
    centerForm Me
    Load_InsProv
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmProvInsurance.Form_Load", _
     Err.Number, _
     Err.Description
     Resume
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


