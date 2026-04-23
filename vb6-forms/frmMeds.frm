VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmMeds 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Medications"
   ClientHeight    =   7800
   ClientLeft      =   3810
   ClientTop       =   1440
   ClientWidth     =   13575
   Icon            =   "frmMeds.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   13575
   Begin VB.Frame fraClose 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   5625
      TabIndex        =   0
      Top             =   6570
      Width           =   2325
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   345
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Save & Close"
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
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
         Picture         =   "frmMeds.frx":08CA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "OK"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMeds.frx":0C64
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCancel 
         Height          =   345
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Close without Save"
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
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
         Picture         =   "frmMeds.frx":0FB6
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMeds.frx":19C8
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMeds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lintComponentCounter As Integer
Private lintTabCounter As Integer
Private lbooMissingClass As Boolean
Private lstrClassName As String
Private lstrLibrary As String
Public piCurPatient As Long
Public mlngCurrentEncounterID As Long
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private mlngFormID As Long
Public mintindex As Integer
Private mFirst As Boolean
Private mstrRemarks As String
Private mbooSearch As Boolean
Public mstrSource As String

Private Function Load_Prescription()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_Prescription
    If mstrSource = "MDI" Then
        If frmPatients.piPatientID > 0 Then
            piCurPatient = frmPatients.piPatientID
            Get_Patient piCurPatient
            MsgBox "Current Patient is" & " " & clsPatient.Name & " " & "if not correct please change patient!", vbInformation
        Else
            MsgBox "Please select a Patient to write prescriptions!", vbInformation
        End If
    ElseIf mstrSource = "Telephone" Then
        If frmPatients.piPatientID > 0 Then
            piCurPatient = frmPatients.piPatientID
            Get_Patient piCurPatient
            MsgBox "Current Patient selected is" & " " & " " & clsPatient.Name & " " & "if not correct please change patient!", vbInformation
        Else
            MsgBox "Please select a Patient to write prescriptions!", vbInformation
        End If
    Else
        If piCurPatient <> 0 Then
             MsgBox "Current Patient selected is" & " " & " " & clsPatient.Name & " " & "if not correct please change patient!", vbInformation
        Else
            MsgBox "Please select a Patient to write prescriptions!", vbInformation
        End If
    End If
    If piProviderID = 0 Then Get_DefaultProvider
    strOCX = "PT_Medications"
    strControl = "Prescription"
    intACTXID = Get_ComponentID(strOCX, strControl)
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlPrescription", Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 150
        fobjOCX.Top = 0
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = piCurPatient
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        fraClose.Top = fobjOCX.Height + 100
        frmMeds.Height = fobjOCX.Height + fraClose.Height + 700
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Prescription:
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If
   Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_Prescription", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cmdCancel_Click()
Dim iAnswer As Integer
    Unload Me
   If Not IsFormLoaded("frmNursingHome") Then
        frmMDI.Get_AppointmentFormat
    End If
    Exit Sub
End Sub


Private Sub cmdSave_Click()

    fobjLib.WriteClinicalData
    Unload Me
    
    Exit Sub
End Sub


Private Sub Form_Activate()
   
    On Error GoTo Error_Handler

    If mFirst Then
       mFirst = False
        centerForm Me
    End If
    Me.ZOrder 0
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmMeds.Form_Activate", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub Form_Load()
    mFirst = True
    Load_Prescription
    centerForm Me
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    On Error Resume Next
    piCurPatient = 0
    If IsFormLoaded("frmInPatient") Then
        frmInPatient.ZOrder 0
    End If
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


