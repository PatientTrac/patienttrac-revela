VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmTelephone 
   Caption         =   "Patient Messaging"
   ClientHeight    =   6735
   ClientLeft      =   7020
   ClientTop       =   3015
   ClientWidth     =   11115
   Icon            =   "frmTelephone.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   11115
   Begin VB.Frame fraButton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   4035
      TabIndex        =   0
      Top             =   6150
      Width           =   3045
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   345
         Left            =   1410
         TabIndex        =   1
         ToolTipText     =   "Save, Restore, Close"
         Top             =   90
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
         Picture         =   "frmTelephone.frx":058A
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   13160660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTelephone.frx":0F9C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   345
         Left            =   330
         TabIndex        =   2
         ToolTipText     =   "Save, Restore, Close"
         Top             =   90
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
         Picture         =   "frmTelephone.frx":12EE
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   13160660
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTelephone.frx":1688
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmTelephone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private piSelPatientID As Long
Public piSelEnctrID As Long
Private piPatientExtID As Long
Private mintActXId As Long
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
Private mbooByPass As Boolean
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



Private Sub Load_Control_Phone()
Dim strOCX As String
Dim strControl As String
Dim lstrControl As String
Dim lstrSQL As String
Dim strSQL As String
Dim lbooReadOnly As Boolean
Dim cnGetTerms As New ADODB.Connection
Dim rsGetTerms As New ADODB.Recordset
Dim strPatCtl As String
Dim strPatOcx As String

    On Error GoTo Error_Load_Control

    If cnGetTerms.State = adStateClosed Then cnGetTerms.open psConnect(1)
    
    strPatOcx = "PatientTrac_Report"
    strPatCtl = "PhoneNotes"
    lstrSQL = "Select ACTX_ID FROM ACTIVE_X " _
        & "Where ACTX_NAME = '" & strPatOcx & "' And " _
        & "ACTX_CLASS = '" & strPatCtl & "'"
    Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
    With rsGetTerms
        If Not .EOF Then
            If Not IsNull(!actx_id) Then
                mintActXId = (!actx_id)
            End If
        End If
    End With
    If mintActXId = 0 Then
            strSQL = "INSERT INTO ACTIVE_X (ACTX_NAME, ACTX_CLASS,ACTX_REL_DIR, ACTX_FILE_NAME, ACTX_VERSION) " _
                & "values('" & strPatOcx & "', '" & strPatCtl & "', 'Components', '" & strPatOcx & "', '2')"
            Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
            
            strSQL = "Select ACTX_ID FROM ACTIVE_X " _
            & "Where ACTX_NAME = '" & strPatOcx & "'  And " _
            & "ACTX_CLASS = '" & strPatCtl & "'"
            Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
            With rsGetTerms
                If Not .EOF Then
                    If Not IsNull(!actx_id) Then
                        mintActXId = (!actx_id)
                    End If
                End If
            End With
    End If
    If cnGetTerms.State = adStateOpen Then
        cnGetTerms.Close
        Set cnGetTerms = Nothing
        If rsGetTerms.State = adStateOpen Then
            rsGetTerms.Close
           Set rsGetTerms = Nothing
        End If
    End If
    Set fobjOCX = Nothing
    Licenses.Add strPatOcx & "." & strPatCtl
    Set fobjOCX = Controls.Add(strPatOcx & "." & strPatCtl, "ctl" & strPatCtl, Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 172
        fobjOCX.Top = 24
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = mintActXId
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = mlngMessageID
        fobjLib.EncounterID = modSpecific.piEncounterID
        fobjLib.ProviderID = modSpecific.piProviderID
        fobjLib.ReadClinicalData
        fraButton.Top = Val(fobjOCX.Height) + 150
        Me.Height = (fobjOCX.Height + fraButton.Height) + 800
        Me.Width = fobjOCX.Width + 500
        fraButton.Left = (Me.Width / 2) - (fraButton.Width / 2)
     Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Control:
    Screen.MousePointer = vbDefault
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then
        Resume Next
    ElseIf Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    Else
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
    Resume
End Sub
Private Sub Load_Control_Contact()
Dim strOCX As String
Dim strControl As String
Dim lstrControl As String
Dim lstrSQL As String
Dim strSQL As String
Dim lbooReadOnly As Boolean
Dim cnGetTerms As New ADODB.Connection
Dim rsGetTerms As New ADODB.Recordset
Dim strPatCtl As String
Dim strPatOcx As String

    On Error GoTo Error_Load_Control

    If cnGetTerms.State = adStateClosed Then cnGetTerms.open psConnect(1)
    
    strPatOcx = "FN_Orders"
    strPatCtl = "PatientContact"
    lstrSQL = "Select ACTX_ID FROM ACTIVE_X " _
        & "Where ACTX_NAME = '" & strPatOcx & "' And " _
        & "ACTX_CLASS = '" & strPatCtl & "'"
    Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
    With rsGetTerms
        If Not .EOF Then
            If Not IsNull(!actx_id) Then
                mintActXId = (!actx_id)
            End If
        End If
    End With
    If mintActXId = 0 Then
            strSQL = "INSERT INTO ACTIVE_X (ACTX_NAME, ACTX_CLASS,ACTX_REL_DIR, ACTX_FILE_NAME, ACTX_VERSION) " _
                & "values('" & strPatOcx & "', '" & strPatCtl & "', 'Components', '" & strPatOcx & "', '2')"
            Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
            
            strSQL = "Select ACTX_ID FROM ACTIVE_X " _
            & "Where ACTX_NAME = '" & strPatOcx & "'  And " _
            & "ACTX_CLASS = '" & strPatCtl & "'"
            Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
            With rsGetTerms
                If Not .EOF Then
                    If Not IsNull(!actx_id) Then
                        mintActXId = (!actx_id)
                    End If
                End If
            End With
    End If
    If cnGetTerms.State = adStateOpen Then
        cnGetTerms.Close
        Set cnGetTerms = Nothing
        If rsGetTerms.State = adStateOpen Then
            rsGetTerms.Close
           Set rsGetTerms = Nothing
        End If
    End If
    Set fobjOCX = Nothing
    'If Licenses.Count > 0 Then Licenses.Clear
    Licenses.Add strPatOcx & "." & strPatCtl
    Set fobjOCX = Controls.Add(strPatOcx & "." & strPatCtl, "ctl" & strPatCtl, Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 172
        fobjOCX.Top = 24
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = mintActXId
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = modSpecific.piPatientID
        fobjLib.EncounterID = modSpecific.piEncounterID
        fobjLib.ProviderID = modSpecific.piProviderID
        fobjLib.ReadClinicalData
        fraButton.Top = Val(fobjOCX.Height) + 150
        Me.Height = (fobjOCX.Height + fraButton.Height) + 800
        Me.Width = fobjOCX.Width + 500
        fraButton.Left = (Me.Width / 2) - (fraButton.Width / 2)
     Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Control:
    Screen.MousePointer = vbDefault
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then
        Resume Next
    ElseIf Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    Else
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
    Resume
End Sub










Private Sub cmdClose_Click()
 
    Unload Me
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    fobjLib.WriteClinicalData
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
        If mstrPatContact = "Phone" Then
            Load_Control_Phone
        ElseIf mstrPatContact = "Contact" Then
            Load_Control_Contact
        End If
    End If
    centerForm Me
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub Form_Load()
    mFirst = True
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


