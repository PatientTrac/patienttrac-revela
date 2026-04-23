VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmSurgComment 
   Caption         =   "SURGEON RISK NOTES"
   ClientHeight    =   3885
   ClientLeft      =   12045
   ClientTop       =   3495
   ClientWidth     =   6570
   Icon            =   "frmSurgComment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   6570
   Begin VB.Frame fraSurNote 
      BackColor       =   &H00E0E0E0&
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      Begin VB.ComboBox cboAttitude 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3210
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1740
         Width           =   2715
      End
      Begin VB.ComboBox cboFee 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3210
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1020
         Width           =   2715
      End
      Begin VB.TextBox cctSurgNote 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2550
         Width           =   4965
      End
      Begin VB.ListBox lstRisk 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1860
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   300
         Width           =   2685
      End
      Begin VB.Frame fraButton 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   2250
         TabIndex        =   1
         Top             =   3300
         Width           =   2085
         Begin PT_XP_Button.PT_XP cmdClose 
            Height          =   405
            Left            =   1080
            TabIndex        =   2
            ToolTipText     =   "Save, Restore, Close"
            Top             =   30
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
            Picture         =   "frmSurgComment.frx":0442
            AppearanceThemes=   3
            BorderWidth     =   0
            TransparentColor=   14215660
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSurgComment.frx":0E54
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdSave 
            Height          =   405
            Left            =   0
            TabIndex        =   3
            ToolTipText     =   "Save, Restore, Close"
            Top             =   30
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
            Picture         =   "frmSurgComment.frx":11A6
            AppearanceThemes=   3
            BorderWidth     =   0
            TransparentColor=   14215660
            Caption         =   "Save"
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSurgComment.frx":1540
            ShowFocus       =   -1  'True
         End
      End
      Begin PT_XP_Button.PT_XP cmdRisk 
         Height          =   345
         Left            =   2910
         TabIndex        =   5
         ToolTipText     =   "Terms Maintanance for Surg Time"
         Top             =   270
         Width           =   405
         _ExtentX        =   714
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
         Picture         =   "frmSurgComment.frx":1892
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSurgComment.frx":1E2C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdFee 
         Height          =   345
         Left            =   6000
         TabIndex        =   10
         ToolTipText     =   "Terms Maintanance for Surg Time"
         Top             =   930
         Width           =   405
         _ExtentX        =   714
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
         Picture         =   "frmSurgComment.frx":217E
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSurgComment.frx":2718
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAtt 
         Height          =   345
         Left            =   6000
         TabIndex        =   13
         ToolTipText     =   "Terms Maintanance for Surg Time"
         Top             =   1680
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmSurgComment.frx":2A6A
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   ""
         CaptionPosition =   4
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmSurgComment.frx":3004
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblPatAtt 
         AutoSize        =   -1  'True
         Caption         =   "PATIENT ATTITUDE"
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
         Height          =   210
         Left            =   3240
         TabIndex        =   11
         Top             =   1500
         Width           =   1950
      End
      Begin VB.Label lblFee 
         AutoSize        =   -1  'True
         Caption         =   "FEE CHANGES"
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
         Height          =   210
         Left            =   3180
         TabIndex        =   8
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblComments 
         AutoSize        =   -1  'True
         Caption         =   "SURGEON NOTE FOR STAFF"
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
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   2610
      End
   End
End
Attribute VB_Name = "frmSurgComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrConnectStringClinical As String
Public mstrConnectStringTerms As String
Public mlngtop As Long
Public mlngleft As Long
Public mstrTerms As String
Public mvarSelectedTerms As String
Public mstrCurrentData As String
Private mbooByPass As Boolean
Public mstrListName As String
Public mFirst As Boolean
Public mbooCaption As Boolean
Private mstrItemClick As String

Private Function Read_Data()
Dim cnPlan As New ADODB.Connection
Dim rsPlan As New ADODB.Recordset
Dim strSQL As String
Dim lngPlanID As Long
Dim strDate As String
Dim strBullet As String
Dim strNote As String
Dim Arr
Dim IDX As Integer
Dim ArrWord
Dim i%

    If cnPlan.State = adStateClosed Then cnPlan.Open modSpecific.mstrConnectStringClinical

    strSQL = "SELECT STAFF_NOTE,COMPLEXITY FROM PATIENT_CONSULTATION_PLAN WHERE PATIENT_ID = " & modSpecific.mlngPatientID & " AND ENCOUNTER_ID = " & modSpecific.mlngEncounterID
    Set rsPlan = cnPlan.Execute(strSQL)
    If Not rsPlan.EOF Then
        If rsPlan.Fields("STAFF_NOTE") <> vbNullString Then strNote = rsPlan.Fields("STAFF_NOTE")
        If rsPlan.Fields("COMPLEXITY") <> vbNullString Then mstrCurrentData = rsPlan.Fields("COMPLEXITY")
    End If
    If strNote <> vbNullString Then
        Arr = Split(strNote, "|")
        If UBound(Arr) > 0 Then
            For IDX = LBound(Arr) To UBound(Arr)
                If TextFound("Surgeon Note", Arr(IDX), Len("Surgeon Note"), True) Then
                    ArrWord = Split(Arr(IDX), ":")
                    If UBound(ArrWord) > 0 Then
                        cctSurgNote.Text = ArrWord(1)
                    End If
                End If
                If TextFound("Patient Demeanor", Arr(IDX), Len("Patient Demeanor"), True) Then
                    ArrWord = Split(Arr(IDX), ":")
                    If UBound(ArrWord) > 0 Then
                        For i = 0 To cboAttitude.ListCount < 1
                            If cboAttitude.List(i) = ArrWord(1) Then
                                cboAttitude.Text = cboAttitude.List(i)
                                Exit For
                            End If
                        Next i
                    End If
                End If
                If TextFound("Fee Changes", Arr(IDX), Len("Fee Changes"), True) Then
                    ArrWord = Split(Arr(IDX), ":")
                    If UBound(ArrWord) > 0 Then
                        For i = 0 To cboFee.ListCount < 1
                            If cboFee.List(i) = ArrWord(1) Then
                                cboFee.Text = cboAttitude.List(i)
                                Exit For
                            End If
                        Next i
                    End If
                End If
            Next IDX
        End If
    End If
    If mstrCurrentData <> vbNullString Then
        For i = 0 To lstRisk.ListCount < 1
            If lstRisk.List(i) = mstrCurrentData Then
                lstRisk.Selected(i) = True
                Exit For
            End If
        Next i
    End If
     If cnPlan.State = adStateOpen Then
        cnPlan.Close
        Set cnPlan = Nothing
        Set rsPlan = Nothing
    End If
    Exit Function
End Function


Private Function Save_Data()
Dim cnPlan As New ADODB.Connection
Dim rsPlan As New ADODB.Recordset
Dim strSQL As String
Dim lngPlanID As Long
Dim strDate As String
Dim strBullet As String
Dim strNote As String


    If cboFee.Text <> vbNullString Then
        strNote = "Fee Changes suggested are:" & " " & Trim(cboFee.Text)
    End If
    If Trim(cctSurgNote.Text) <> vbNullString Then
        strNote = strNote & "|" & "Surgeon Note:" & " " & Trim(cctSurgNote.Text)
    End If
    If Trim(cboAttitude.Text) <> vbNullString Then
        strNote = strNote & "|" & "Patient Demeanor:" & " " & Trim(cboAttitude.Text)
    End If
    strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If cnPlan.State = adStateClosed Then cnPlan.Open modSpecific.mstrConnectStringClinical

    strSQL = "SELECT PLAN_ID FROM PATIENT_CONSULTATION_PLAN WHERE PATIENT_ID = " & modSpecific.mlngPatientID & " AND ENCOUNTER_ID = " & modSpecific.mlngEncounterID
    Set rsPlan = cnPlan.Execute(strSQL)
    If Not rsPlan.EOF Then
        lngPlanID = rsPlan.Fields("PLAN_ID")
    End If
    If lngPlanID = 0 Then
        strSQL = "INSERT INTO PATIENT_CONSULTATION_PLAN (PATIENT_ID,  ENCOUNTER_ID, STAFF_NOTE, COMPLEXITY, INSERT_DATE  " _
            & "Values (" & modSpecific.mlngPatientID & ", " & modSpecific.mlngEncounterID & ", '" & strNote & "','" & mstrCurrentData & "','" & strDate & "')"
        Set rsPlan = cnPlan.Execute(strSQL)
    Else
           strSQL = "Update  PATIENT_CONSULTATION_PLAN SET STAFF_NOTE  = '" & strNote & "',COMPLEXITY = '" & mstrCurrentData & "', UPDATE_DATE = '" & strDate & "' WHERE PLAN_ID = " & lngPlanID
             Set rsPlan = cnPlan.Execute(strSQL)
    End If
     If cnPlan.State = adStateOpen Then
        cnPlan.Close
        Set cnPlan = Nothing
        Set rsPlan = Nothing
    End If
    Exit Function
End Function

Private Sub cmdAtt_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Patient Attitude"
    frmNurseTemplates.fType = "Patient Attitude"
    frmNurseTemplates.fSubDomain = "sd_Risk_PatDemeanor"
    frmNurseTemplates.fConnectCR = mstrConnectStringClinical
    frmNurseTemplates.fConnectTM = mstrConnectStringTerms
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    GetTerms cboAttitude, "sd_Risk_PatDemeanor"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFee_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Fee Changes by Surgeon"
    frmNurseTemplates.fType = "Fee Changes by Surgeon"
    frmNurseTemplates.fSubDomain = "sd_Risk_FeeChange"
    frmNurseTemplates.fConnectCR = mstrConnectStringClinical
    frmNurseTemplates.fConnectTM = mstrConnectStringTerms
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    GetList cboFee, "sd_Risk_FeeChange"
End Sub

Private Sub cmdRisk_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Surgery Risk/Complexity"
    frmNurseTemplates.fType = "Surgery Risk/Complexity"
    frmNurseTemplates.fSubDomain = "sd_Proc_Risk"
    frmNurseTemplates.fConnectCR = mstrConnectStringClinical
    frmNurseTemplates.fConnectTM = mstrConnectStringTerms
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    GetList lstRisk, "sd_Proc_Risk"
End Sub


Private Sub cmdSave_Click()
    Save_Data
    Unload Me
End Sub

Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
       If cboAttitude.ListCount <= 0 Then GetTerms cboAttitude, "sd_Risk_PatDemeanor"
       If lstRisk.ListCount <= 0 Then GetList lstRisk, "sd_Proc_Risk"
       If cboFee.ListCount <= 0 Then GetTerms cboFee, "sd_Risk_FeeChange"
       Read_Data
       CenterForm Me
    End If
End Sub

Private Sub Form_Load()
    mFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next
 
    
    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    
    For Each obj In frm.objects
      Set obj = Nothing
    Next obj
End Sub


Private Sub lstRisk_ItemCheck(Item As Integer)
    mstrCurrentData = vbNullString
    If lstRisk.Selected(Item) Then
        mstrCurrentData = lstRisk.List(Item)
    End If
End Sub


