VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmSuperBill 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "SuperBill"
   ClientHeight    =   9525
   ClientLeft      =   8865
   ClientTop       =   3840
   ClientWidth     =   11820
   Icon            =   "frmSuperBill.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9525
   ScaleWidth      =   11820
   Begin VB.Frame fraButton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4568
      TabIndex        =   0
      Top             =   8940
      Width           =   2685
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   345
         Left            =   1410
         TabIndex        =   1
         ToolTipText     =   "Save, Restore, Close"
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
         Picture         =   "frmSuperBill.frx":058A
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   12632256
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSuperBill.frx":0F9C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   345
         Left            =   330
         TabIndex        =   2
         ToolTipText     =   "Save, Restore, Close"
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
         Picture         =   "frmSuperBill.frx":12EE
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   12632256
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSuperBill.frx":1688
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSuperBill"
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


    strSQL = "select max(encounter_id) from encounter where  " _
            & "patient_id = " & piPatientID
    booret = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
    
    If booret Then
        If lngrecs > 0 Then
            piEncounterID = varRS(0, 0)
        End If
    End If
    Exit Function
End Function

Private Sub Load_SuperBill()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_SuperBill

    strOCX = "FN_PsyMse"
    strControl = "SuperBill"
    If piPatientID = 0 Then piPatientID = frmPatients.piPatientID
    If piEncounterID = 0 Then Get_EncounterID
    intACTXID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlSuperBill", Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 150
        fobjOCX.Top = 100
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        If piPatientID = 0 Then
            MsgBox "No patient selected, please select a patient!" ' vbinformation
            Exit Sub
        Else
            fobjLib.PatientID = piPatientID
        End If
        fobjLib.EncounterID = piEncounterID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        fraButton.Top = Val(fobjOCX.Height) + 150
        Me.Height = (fobjOCX.Height + fraButton.Height) + 800
        Me.Width = fobjOCX.Width + 500
        fraButton.Left = (Me.Width / 2) - (fraButton.Width / 2)
        Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_SuperBill:
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If
    MsgBox Err.Description, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume Next
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    fobjLib.WriteClinicalData
    Unload Me
End Sub


Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
        centerForm Me
    End If
    Exit Sub
End Sub

Private Sub Form_Load()
Dim iAnswer As String

    mFirst = True
    If piPatientID = 0 Then
        iAnswer = MsgBox("Please select a Patient First", vbOKCancel + vbExclamation + vbApplicationModal + vbDefaultButton1, "Open Superbill")
           If iAnswer = vbOK Then
                Me.Hide
                If IsFormLoaded("frmpatients") Then
                    frmPatients.Show vbModal
                Else
                    frmPatients.Show vbModal
                     If piPatientID > 0 Then
                        Load_SuperBill
                    Else
                        Screen.MousePointer = vbDefault
                        Unload Me
                    End If
                End If
                
            ElseIf vbCancel Then
                 Screen.MousePointer = vbDefault
                Unload Me
            End If
    Else
        Load_SuperBill
         Screen.MousePointer = vbDefault
    End If
         Screen.MousePointer = vbDefault
    Exit Sub
Error_Handler:
    MsgBox Err.Description, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume Next
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


