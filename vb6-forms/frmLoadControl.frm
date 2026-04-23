VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmLoadControl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8460
   ClientLeft      =   5955
   ClientTop       =   2535
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   Icon            =   "frmLoadControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11385
   Begin VB.Frame fraButton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   4710
      TabIndex        =   0
      Top             =   7800
      Width           =   1965
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   345
         Left            =   990
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
         Picture         =   "frmLoadControl.frx":0ECA
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmLoadControl.frx":18DC
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   345
         Left            =   0
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
         Picture         =   "frmLoadControl.frx":1C2E
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmLoadControl.frx":1FC8
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmLoadControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public piSelPatientID As Long
Public piSelEnctrID As Long
Private piPatientExtID As Long
Private mintActXId As Long
Private msEncounterDate As String
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
Private mcolActxComp As New Collection
Private mbooSelectPat As Boolean
Private mbooRecurssive As Boolean
Private mlngLastPatientId As Long
Private mlngLastEncounterId As Long
Private mObjSelectPatient As PatientTracTypeLib.iClinical
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






Function Load_Control()
Dim strOCX As String
Dim strControl As String
Dim lstrControl As String
Dim lstrSQL As String
Dim strSQL As String
Dim lbooReadOnly As Boolean
Dim cnGetTerms As New ADODB.Connection
Dim rsGetTerms As New ADODB.Recordset
 

    On Error GoTo Error_Load_Control
    Screen.MousePointer = vbHourglass
    If cnGetTerms.State = adStateClosed Then cnGetTerms.open psConnect(1)
    lstrSQL = "Get_ActiveX_ID '" & mstrPatOcx & "', '" & mstrPatCtl & "'"
    Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
    With rsGetTerms
        If Not .EOF Then
            If Not IsNull(!actx_id) Then
                mintActXId = (!actx_id)
            End If
        End If
    End With
    If mintActXId = 0 Then
        strSQL = "Insert_New_ActiveX '" & mstrPatOcx & "', '" & mstrPatCtl & "', 'Components', '" & mstrPatOcx & "', '2' "
        Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
        strSQL = "Get_ActiveX_ID '" & mstrPatOcx & "', '" & mstrPatCtl & "'"
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
    End If
    Set fobjOCX = Nothing
    If piSelPatientID = 0 Then piSelPatientID = piPatientID
    Licenses.Add mstrPatOcx & "." & mstrPatCtl
    Set fobjOCX = Controls.Add(mstrPatOcx & "." & mstrPatCtl, "ctl" & mstrPatCtl, Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 150
        fobjOCX.Top = 100
        Me.Width = fobjOCX.Width + 700
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = mintActXId
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = piSelPatientID
        fobjLib.EncounterID = piEncounterID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        Screen.MousePointer = vbDefault
        fraButton.Left = (Me.Width / 2) - (fraButton.Width / 2)
        fraButton.Top = Val(fobjOCX.Height) + 150
        Me.Height = (fobjOCX.Height + fraButton.Height) + 800
     Screen.MousePointer = vbDefault
     booAppTime = False
    Exit Function
Error_Load_Control:
    Screen.MousePointer = vbDefault
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then
        Resume Next
    ElseIf Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
        Resume Next
    Else
        Screen.MousePointer = vbDefault
        Write_Error Err.Description, Err.Number, strSQL, "Load_Control", App.EXEName & "." & Me.Name
        Exit Function
        Resume Next
    End If
    Exit Function
End Function



Private Sub cmdClose_Click()
   Unload Me
  ' If Not IsFormLoaded("frmNursingHome") Then
    '    frmMDI.Get_AppointmentFormat
   ' End If
End Sub

Private Sub cmdSave_Click()
    fobjLib.WriteClinicalData
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
   If mFirst = True Then
        mFirst = False
        '--Load_Control
        Me.Refresh
        centerForm Me
   End If
    Screen.MousePointer = vbDefault
   Exit Sub
End Sub

Private Sub Form_Load()
    mFirst = True
    Load_Control
    booAppTime = False
    Exit Sub
End Sub

Public Property Get FormID() As Long
    FormID = mlngFormID
End Property

Public Property Let FormID(ByVal vlngNewValue As Long)
    mlngFormID = vlngNewValue
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
End Sub

Public Sub Form_Unload(Cancel As Integer)
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
    booAppTime = True
    Unload Me
    Exit Sub
End Sub












