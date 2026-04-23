VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmMDSign 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PatientTrac Electronic Signature"
   ClientHeight    =   3795
   ClientLeft      =   8295
   ClientTop       =   8985
   ClientWidth     =   10845
   Icon            =   "frmMDSign.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   10845
   Begin VB.Frame fraHomRisk 
      BackColor       =   &H00E0E0E0&
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
      Height          =   3225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10785
      Begin PatientTrac.ccApprove ccNurseSig 
         Height          =   675
         Left            =   480
         TabIndex        =   16
         Top             =   2310
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1191
      End
      Begin PatientTrac.ccApprove ccMDSig 
         Height          =   825
         Left            =   480
         TabIndex        =   15
         Top             =   660
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   1455
      End
      Begin VB.Frame fraDefault 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Default Physician Review"
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
         Height          =   1575
         Left            =   4920
         TabIndex        =   4
         Top             =   540
         Visible         =   0   'False
         Width           =   5685
         Begin VB.TextBox txtDefaultReview 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   795
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   360
            Width           =   5355
         End
         Begin PT_XP_Button.PT_XP cmdDone 
            Height          =   285
            Left            =   2070
            TabIndex        =   6
            Top             =   1230
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
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
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Done"
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            ShowFocus       =   -1  'True
         End
      End
      Begin VB.ComboBox cboAttendPhy 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8220
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1110
         Width           =   2505
      End
      Begin VB.ComboBox cboAssistant 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2370
         Width           =   3135
      End
      Begin VB.TextBox cctAttendPhys 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1035
         Left            =   5040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Above read, confirmed, and/or corrected"
         Top             =   630
         Width           =   3075
      End
      Begin PT_XP_Button.PT_XP cmdDefaultReview 
         Height          =   465
         Left            =   9450
         TabIndex        =   7
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
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
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Set Default Review"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Clinician/Resident/Physician Assistant"
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
         Left            =   420
         TabIndex        =   12
         Top             =   1830
         Width           =   3840
      End
      Begin VB.Label lblMedDtr 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Medical Doctor"
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
         Height          =   255
         Left            =   420
         TabIndex        =   11
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label lblAttdPhys 
         BackStyle       =   0  'Transparent
         Caption         =   "Attending Physician"
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
         Height          =   225
         Left            =   8250
         TabIndex        =   10
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblNurse 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Physician Assistant"
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
         Height          =   285
         Left            =   5070
         TabIndex        =   9
         Top             =   2100
         Width           =   2595
      End
      Begin VB.Label lblMD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Physician Review"
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
         Height          =   315
         Left            =   5100
         TabIndex        =   8
         Top             =   330
         Width           =   2595
      End
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   345
      Left            =   5475
      TabIndex        =   13
      Top             =   3360
      Width           =   1065
      _ExtentX        =   1879
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
      Picture         =   "frmMDSign.frx":0442
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMDSign.frx":0E54
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   345
      Left            =   4305
      TabIndex        =   14
      Top             =   3360
      Width           =   1065
      _ExtentX        =   1879
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
      Picture         =   "frmMDSign.frx":11A6
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMDSign.frx":1540
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmMDSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Public mstrRef As String
Public piSignPatID As Long
Public piSignEnctr As Long
Public piSignProv As Long
Public vstrSource As String
Public vstrSummary As String
Private udtEnctr As udtEnctr
Private Type udtEnctr
    EnctrID As Long
    StartDate As String
    EndDate As String
    VisitReason As String
    FacID As Long
    Provider As String
    Admission As String
    Discharge As String
    ServicePlace As String
    ServicePlaceCode As String
    EncounterType As String
    Appointment As String
    ProviderID As Long
    VisitType As String
    Intrvn_CD As String
    VisitTime As String
End Type
Private Sub Clear_EnctrString()
    With udtEnctr
        .EnctrID = 0
        .StartDate = vbNullString
        .EndDate = vbNullString
        .VisitReason = vbNullString
        .FacID = 0
        .Provider = vbNullString
        .Admission = vbNullString
        .Discharge = vbNullString
        .ServicePlace = vbNullString
        .ServicePlaceCode = vbNullString
        .EncounterType = vbNullString
        .Appointment = vbNullString
        .ProviderID = 0
        .VisitType = vbNullString
        .Intrvn_CD = vbNullString
        .VisitTime = vbNullString
    End With
End Sub

Private Function Read_Default(vlngProviderID As Long)
Dim strSQL As String
Dim cnDefault As New ADODB.Connection
Dim rsDefault As New ADODB.Recordset
Dim strClinicalDesc As String

    On Error GoTo Error_Read_Default
    Screen.MousePointer = vbHourglass

    vstrSummary = vbNullString
    If cnDefault.State = adStateClosed Then cnDefault.open psConnect(1)
    strSQL = "Select CLINICAL_VALUE FROM REPORT_DEFAULTS " _
        & "WHERE  CLINICAL_DESC = '" & App.EXEName & "' And CLINICAL_VALUE_REF = '" & vstrSource & "' And PROVIDER_ID = " & vlngProviderID & ""
    Set rsDefault = cnDefault.Execute(strSQL)
    With rsDefault
        If Not .EOF Then
            If Not IsNull(!CLINICAL_VALUE) Then
                vstrSummary = Trim((!CLINICAL_VALUE))
            End If
        End If
    End With
    If cnDefault.State = adStateOpen Then
        cnDefault.Close
        Set cnDefault = Nothing
        Set rsDefault = Nothing
    End If
    Exit Function
    Screen.MousePointer = vbDefault
Error_Read_Default:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Default", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Save_Default(vlngProviderID As Long)
Dim strSQL As String
Dim strProvider As String
Dim strDefaultText As String
Dim booinsert As Boolean
Dim strClinicalDesc As String
Dim strObject As String
Dim intComponentID As Long
Dim lngDefaultID As Long
Dim sDate As String
Dim cnSave As New ADODB.Connection
Dim rsSave As New ADODB.Recordset

    strDefaultText = vstrSummary
    booinsert = False
    On Error GoTo Error_Save_Default
    Screen.MousePointer = vbHourglass
    If cnSave.State = adStateClosed Then cnSave.open psConnect(1)
    strSQL = "Select DEFAULT_ID FROM REPORT_DEFAULTS " _
            & "Where provider_id = '" & vlngProviderID & "' AND CLINICAL_VALUE_REF =  '" & vstrSource & "' AND CLINICAL_DESC = '" & App.EXEName & "'"
    Set rsSave = cnSave.Execute(strSQL)
    If Not rsSave.EOF Then
        lngDefaultID = rsSave.Fields("DEFAULT_ID")
    End If
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If lngDefaultID = 0 Then
        strSQL = "Insert into REPORT_DEFAULTS  (CLINICAL_VALUE, CLINICAL_VALUE_REF, CLINICAL_DESC, ACTX_ID, " _
                & "PROVIDER_ID, INSERT_DATE) " _
                & "Values('" & vstrSummary & "','" & vstrSource & "','" & App.EXEName & "',1, " & vlngProviderID & ",'" & sDate & "')"
        Set rsSave = cnSave.Execute(strSQL)
    Else  'Update
        strSQL = "Update REPORT_DEFAULTS " _
            & "Set CLINICAL_VALUE = '" & vstrSummary & "'," _
            & "CLINICAL_VALUE_REF = '" & vstrSource & "'," _
            & "CLINICAL_DESC = '" & App.EXEName & "'," _
            & "ACTX_ID = 1," _
            & "UPDATE_DATE = '" & sDate & "'" _
            & "WHERE DEFAULT_ID = " & lngDefaultID
        Set rsSave = cnSave.Execute(strSQL)
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Save_Default:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Save_Default", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function SaveSignature(vlngProviderID As Long, vstrValue As String, vstrRef As String, vstrRefData As String, vintSequence As Integer, vbooPrimary As Boolean)
Dim cnNotes As New ADODB.Connection
Dim rsNotes As New ADODB.Recordset
Dim strSQL As String
Dim strCntr As String
Dim booinsert As Boolean
Dim lngSigId As Long
Dim intComID As Long

    On Error GoTo Error_SaveSignature
    Screen.MousePointer = vbHourglass
    If piSignPatID = 0 Then
        MsgBox "Cannot Save Data - No Patient selected. Please select a Patient first!", vbInformation
        Exit Function
    End If
    If PsUserID = vbNullString Then PsUserID = Get_PsUserID(psConnect(0))
    intComID = 10
    If cnNotes.State = adStateClosed Then cnNotes.open psConnect(0)
    booinsert = True
    strCntr = App.EXEName & "." & vstrSource
    strSQL = "SELECT SIGNATURE_ID FROM SIGNATURES WHERE PATIENT_ID = " & piSignPatID & " AND ENCOUNTER_ID = " & piSignEnctr & " AND " _
        & "SIGNATURE_REF = '" & vstrRef & "' and PROVIDER_ID =  " & vlngProviderID & " AND REF_DATA = '" & vstrRefData & "' AND CLINICAL_DESC = '" & strCntr & "'"
        Set rsNotes = cnNotes.Execute(strSQL)
    If Not rsNotes.EOF Then
        lngSigId = rsNotes.Fields("SIGNATURE_ID")
    End If
    If Len(Trim(vstrValue)) <= 4 Then
        vstrValue = Replace(vstrValue, vbNewLine, vbNullString)
    End If
    If Trim(vstrValue) <> vbNullString Then
        If lngSigId = 0 Then
            strSQL = "INSERT INTO SIGNATURES ( PATIENT_ID,ENCOUNTER_ID, ACTX_ID, PROVIDER_ID, SIGNATURE_REF," _
                & "SEQUENCE_NO, SIGNATURE, CLINICAL_DESC, REF_DATA, INSERT_DATE, USER_ID) " _
                & "values(" & piSignPatID & ", " & piSignEnctr & "," & intComID & "," & vlngProviderID & ", '" & vstrRef & "'," & vintSequence & ",'" _
                & vstrValue & "','" & strCntr & "', '" & vstrRefData & "','" & Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "','" & PsUserID & "')"
            Set rsNotes = cnNotes.Execute(strSQL)
        Else
            strSQL = "UPDATE SIGNATURES SET SIGNATURE = '" & vstrValue & "', REF_DATA = '" & vstrRefData & "', " _
                & "UPDATE_DATE = '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "', SEQUENCE_NO = " & vintSequence & "," _
                & "PROVIDER_ID = " & vlngProviderID & ", INSERT_DATE = '" & Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "'  WHERE " _
                & "SIGNATURE_ID = " & lngSigId
            Set rsNotes = cnNotes.Execute(strSQL)
        End If
    End If
    If cnNotes.State = adStateOpen Then
        cnNotes.Close
        Set cnNotes = Nothing
        Set rsNotes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_SaveSignature:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "SaveSignature", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Sub SaveNotes(vlngProviderID As Long, ByVal vstrValue As String, ByVal vstrRef As String, ByVal vintSequence As Integer, ByVal vbooPrimary As Boolean)
Dim cnNotes As New ADODB.Connection
Dim rsNotes As New ADODB.Recordset
Dim strSQL As String
Dim strCntr As String
Dim booinsert As Boolean
Dim intComID As Integer

    On Error GoTo Error_SaveNotes
    Screen.MousePointer = vbHourglass
    If piSignPatID = 0 Then
        MsgBox "Cannot Save Data - No Patient selected. Please select a Patient first!", vbInformation
        Exit Sub
    End If
    intComID = 10
    If PsUserID = vbNullString Then PsUserID = Get_PsUserID(psConnect(0))
    If cnNotes.State = adStateClosed Then cnNotes.open psConnect(0)
    booinsert = True
    strCntr = App.EXEName & "." & vstrSource
    strSQL = "SELECT * FROM PATIENT_NOTES WHERE PATIENT_ID = " & piSignPatID & " AND ENCOUNTER_ID = " & piSignEnctr & " AND " _
        & "CLINICAL_DESC = '" & strCntr & "' and CLINICAL_VALUE_REF = '" & vstrRef & "'"
        Set rsNotes = cnNotes.Execute(strSQL)
    If Not rsNotes.EOF Then
        booinsert = False
    End If
    If Len(Trim(vstrValue)) <= 4 Then
        vstrValue = Replace(vstrValue, vbNewLine, vbNullString)
    End If
    If Trim(vstrValue) <> vbNullString Then
        If booinsert Then
            strSQL = "INSERT INTO PATIENT_NOTES ( PATIENT_ID, PROVIDER_ID, ENCOUNTER_ID, ACTX_ID, CLINICAL_VALUE_REF," _
                & "SEQUENCE_NO, CLINICAL_NOTES, CLINICAL_DESC, PRI_VALUE_IND, INSERT_DATE, USER_ID) " _
                & "values(" & piSignPatID & ", " & vlngProviderID & ", " & piSignEnctr & ", " & intComID & ", '" & vstrRef & "', 1,'" _
                & vstrValue & "','" & strCntr & "', 1,'" & Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "','" & PsUserID & "')"
            Set rsNotes = cnNotes.Execute(strSQL)
        Else
            strSQL = "UPDATE PATIENT_NOTES SET CLINICAL_NOTES = '" & vstrValue & "', UPDATE_DATE = '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," _
                 & "INSERT_DATE = '" & Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "' ,PROVIDER_ID =  " & vlngProviderID & "  WHERE " _
                & "PATIENT_ID = " & piSignPatID & " AND ENCOUNTER_ID = " & piSignEnctr & " AND CLINICAL_VALUE_REF = '" & vstrRef & "' AND  " _
                & "CLINICAL_DESC = '" & strCntr & "'"
                 Set rsNotes = cnNotes.Execute(strSQL)
        End If
    End If
    If cnNotes.State = adStateOpen Then
        cnNotes.Close
        Set cnNotes = Nothing
        Set rsNotes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_SaveNotes:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, vstrValue, "SaveNotes", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub SaveValue(vlngProviderID As Long, ByVal vstrValue As String, ByVal vstrRef As String, ByVal vintSequence As Integer, ByVal vbooPrimary As Boolean)
Dim cnNotes As New ADODB.Connection
Dim rsNotes As New ADODB.Recordset
Dim strSQL As String
Dim strCntr As String
Dim booinsert As Boolean
Dim intComID As Integer

    On Error GoTo Error_SaveValue
    Screen.MousePointer = vbHourglass
    If piSignPatID = 0 Then
        MsgBox "Cannot Save Data - No Patient selected. Please select a Patient first!", vbInformation
        Exit Sub
    End If
    intComID = 10
    If cnNotes.State = adStateClosed Then cnNotes.open psConnect(0)
    booinsert = True
    strCntr = App.EXEName & "." & vstrSource
    strSQL = "SELECT * FROM PATIENT_CLINICAL WHERE PATIENT_ID = " & piSignPatID & " AND ENCOUNTER_ID = " & piSignEnctr & " AND " _
        & "CLINICAL_DESC = '" & strCntr & "' and CLINICAL_VALUE_REF = '" & vstrRef & "'"
        Set rsNotes = cnNotes.Execute(strSQL)
    If Not rsNotes.EOF Then
        booinsert = False
    End If
    If Len(Trim(vstrValue)) <= 4 Then
        vstrValue = Replace(vstrValue, vbNewLine, vbNullString)
    End If
    If Trim(vstrValue) <> vbNullString Then
        If booinsert Then
            strSQL = "INSERT INTO PATIENT_CLINICAL ( PATIENT_ID, PROVIDER_ID, ENCOUNTER_ID, ACTX_ID, CLINICAL_VALUE_REF," _
                & "SEQUENCE_NO, CLINICAL_VALUE, CLINICAL_DESC, PRI_VALUE_IND, INSERT_DATE, USER_ID) " _
                & "values(" & piSignPatID & "," & vlngProviderID & "," & piSignEnctr & ", " & intComID & ", '" & vstrRef & "', 1,'" _
                & vstrValue & "','" & strCntr & "', 1,'" & Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "','" & PsUserID & "')"
            Set rsNotes = cnNotes.Execute(strSQL)
        Else
            strSQL = "UPDATE PATIENT_CLINICAL SET CLINICAL_VALUE = '" & vstrValue & "', UPDATE_DATE = '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," _
                 & "INSERT_DATE = '" & Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "',PROVIDER_ID =  " & vlngProviderID & " WHERE " _
                 & "PATIENT_ID = " & piSignPatID & " AND ENCOUNTER_ID = " & piSignEnctr & " AND CLINICAL_VALUE_REF = '" & vstrRef & "' AND  " _
                 & "CLINICAL_DESC = '" & strCntr & "'"
                 Set rsNotes = cnNotes.Execute(strSQL)
        End If
    End If
    If cnNotes.State = adStateOpen Then
        cnNotes.Close
        Set cnNotes = Nothing
        Set rsNotes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_SaveValue:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, vstrValue, "SaveValue", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Public Function Get_MedProvider()
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strMedProvider As String
Dim cnMedProv As New ADODB.Connection
Dim rsMedProv As New ADODB.Recordset
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim strSuffix As String
Dim strCity As String
Dim strState As String
Dim strZip As String
Dim strZipExt As String
Dim booSkip As Boolean
Dim lngProvID As Long
Dim i%

    On Error GoTo Error_Get_MedProvider
    Screen.MousePointer = vbHourglass
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    strSQL = "Load_Active_Providers"
    Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        Do Until .EOF
            Clear_Provider
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            lngProvID = 0
            booSkip = True
            If Not IsNull(!Title) Then
                strTitle = (!Title)
                If TextFound("MD", strTitle, Len("MD"), True) Or _
                    TextFound("M.D.", strTitle, Len("M.D."), True) Or _
                    TextFound("D.O.", strTitle, Len("D.O."), True) Or _
                    TextFound("DO", strTitle, Len("DO"), True) Then
                        booSkip = False
                    End If
            End If
            If Not booSkip Then
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Fld.Value <> vbNullString Then
                            If UCase(Fld.Name) = "FIRST_NAME" Then
                                strFirst = Fld.Value
                            ElseIf UCase(Fld.Name) = "MIDDLE_NAME" Then
                                strMiddle = Fld.Value
                            ElseIf UCase(Fld.Name) = "LAST_NAME" Then
                                strLast = Fld.Value
                            ElseIf UCase(Fld.Name) = "SUFFIX" Then
                                strSuffix = Fld.Value
                            ElseIf UCase(Fld.Name) = "TITLE" Then
                                strTitle = Fld.Value
                            ElseIf UCase(Fld.Name) = "PROVIDER_ID" Then
                                lngProvID = Fld.Value
                            End If
                        End If
                    End If
                Next
            End If
            .MoveNext
            If Trim(strFirst) <> vbNullString Then
                 clsProvider.Name = Trim(strFirst)
             End If
            If Trim(strMiddle) <> vbNullString Then
                strMiddle = IsProperCase(strMiddle)
                 clsProvider.Name = clsProvider.Name & " " & Left(strMiddle, 1)
             End If
            If Trim(strLast) <> vbNullString Then
                 clsProvider.Name = clsProvider.Name & " " & Trim(strLast)
             End If
            If Trim(strSuffix) <> vbNullString And Trim(clsProvider.Name) <> vbNullString Then
                 clsProvider.Name = clsProvider.Name & ", " & Trim(strSuffix)
             End If
            If Trim(strTitle) <> vbNullString And Trim(clsProvider.Name) <> vbNullString Then
                 clsProvider.Name = clsProvider.Name & ", " & Trim(strTitle)
                 clsProvider.Title = Trim(strTitle)
             End If
            If clsProvider.Name <> vbNullString Then
                cboAttendPhy.AddItem clsProvider.Name
                If IsNumeric(lngProvID) Then cboAttendPhy.ItemData(cboAttendPhy.NewIndex) = lngProvID
            End If
        Loop
    End With
    For i = 0 To cboAttendPhy.listcount - 1
        If cboAttendPhy.ItemData(i) = piSignProv Then
            cboAttendPhy.Text = cboAttendPhy.List(i)
        End If
    Next i
    If cnMedProv.State = adStateOpen Then
        cnMedProv.Close
        Set cnMedProv = Nothing
        Set rsMedProv = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_MedProvider:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_MedProvider", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub ccMDSig_AfterUpdate()
    If Not ccMDSig.EmptySignature Then
        If cboAttendPhy.Text = vbNullString Then
            MsgBox "Please select Provider Name from the Physician dropdown.", vbInformation
        End If
    End If
    Exit Sub
End Sub

Private Sub ccNurseSig_AfterUpdate()
    If Not ccNurseSig.EmptySignature Then
        If cboAssistant.Text = vbNullString Then
            MsgBox "Please select Provider Name from the Physician assistant dropdown.", vbInformation
        End If
    End If
    Exit Sub
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    Exit Sub
End Sub

Private Sub cmdDefaultReview_Click()
    fraDefault.Visible = True
End Sub

Private Sub cmdDone_Click()
Dim intActX_ID As Integer

    fraDefault.Visible = False
    If txtDefaultReview <> vbNullString Then
        If Trim(txtDefaultReview.Text) <> vbNullString Then vstrSummary = Trim(txtDefaultReview.Text)
        intActX_ID = Get_ComponentID("FN_PsyMse", "Signature")
        If txtDefaultReview.Text <> vbNullString Then Save_Default piSignProv
    End If
    Exit Sub
End Sub

Private Function ReadClinicalData()
Dim sStr As String
Dim lstrSignatureToken As String
Dim lbooInvalidData As Boolean
Dim i%
    On Error GoTo Error_ReadClinicalData
    If piSignEnctr > 0 Then Get_CurrentEnctr
    If PsUserID = vbNullString Then PsUserID = Get_PsUserID(psConnect(0))
    Clear_Signature
    Get_NurseProvider
    Get_MedProvider
    Get_Providers
    vstrSummary = vbNullString
    Read_Default piSignProv
    If Not TextFound(vstrSummary, txtDefaultReview.Text, Len(vstrSummary), True) And Trim(txtDefaultReview.Text) = vbNullString Then
        txtDefaultReview.Text = vstrSummary
    ElseIf Not TextFound(Trim(vstrSummary), Trim(txtDefaultReview.Text), Len(Trim(vstrSummary)), True) And Trim(txtDefaultReview.Text) <> vbNullString Then
        txtDefaultReview.Text = Trim(txtDefaultReview.Text) & " " & vstrSummary
    End If
    clsSignature.SignReason = mstrRef
    lstrSignatureToken = Get_Signature(App.EXEName, vstrSource, "ccMDSig", 1, piSignPatID, piSignEnctr, clsSignature.SignReason)
    If lstrSignatureToken <> vbNullString Then
        lbooInvalidData = ccMDSig.DecodeSignatureValues(lstrSignatureToken, clsSignature.SignReason)
        If ccMDSig.EmptySignature Then
            lstrSignatureToken = vbNullString
            lstrSignatureToken = Get_Signature(App.EXEName, vstrSource, "ccMDSig", 1, piSignPatID, piSignEnctr, vstrSummary)
            If lstrSignatureToken <> vbNullString Then
                lbooInvalidData = ccMDSig.DecodeSignatureValues(lstrSignatureToken, vstrSummary)
            End If
        End If
    End If
    lstrSignatureToken = vbNullString
    lstrSignatureToken = Get_Signature(App.EXEName, vstrSource, "ccNurseSig", 1, piSignPatID, piSignEnctr, clsSignature.SignReason)
    If lstrSignatureToken <> vbNullString Then
        lbooInvalidData = ccNurseSig.DecodeSignatureValues(lstrSignatureToken, clsSignature.SignReason)
    End If
    If ccNurseSig.EmptySignature Then
        txtDefaultReview.Text = vbNullString
        For i = 0 To cboAttendPhy.listcount - 1
            If cboAttendPhy.ItemData(i) = piSignPatID Then
                cboAttendPhy.Text = cboAttendPhy.List(i)
                Exit For
            End If
        Next i
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_ReadClinicalData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "ReadClinicalData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function WriteClinicalData()
Dim lstrSignatureToken As String
Dim strRefData As String
Dim lngProvID As Long
Dim i
Dim lngAsstID As Long

    On Error GoTo Error_WriteClinicalData
    Screen.MousePointer = vbHourglass
    Clear_Signature
    For i = 0 To cboAttendPhy.listcount - 1
        If cboAttendPhy.List(i) = cboAttendPhy.Text Then
            lngProvID = cboAttendPhy.ItemData(i)
            Exit For
        End If
    Next i
    For i = 0 To cboAssistant.listcount - 1
        If cboAssistant.List(i) = cboAssistant.Text Then
            lngAsstID = cboAssistant.ItemData(i)
            Exit For
        End If
    Next i
    clsSignature.SignReason = mstrRef
    vstrSummary = vbNullString
    If Trim(cboAssistant.Text) <> vbNullString Then SaveValue lngAsstID, cboAssistant.Text, cboAssistant, 1, True
    If Trim(cboAttendPhy.Text) <> vbNullString Then SaveValue lngProvID, cboAttendPhy.Text, cboAttendPhy, 1, True
    If Trim(cctAttendPhys.Text) <> vbNullString Then SaveNotes lngProvID, cctAttendPhys.Text, cctAttendPhys, 1, True
    If Trim(cctAttendPhys.Text) <> vbNullString Then vstrSummary = clsSignature.SignReason & " " & Trim(cctAttendPhys.Text)
    If Not ccMDSig.EmptySignature Then
        If vstrSummary <> vbNullString Then
            lstrSignatureToken = ccMDSig.EncodeSignatureValues(vstrSummary)
            SaveSignature lngProvID, lstrSignatureToken, "ccMDSig", vstrSummary, 2, True
        Else
            lstrSignatureToken = ccMDSig.EncodeSignatureValues(clsSignature.SignReason)
            SaveSignature lngProvID, lstrSignatureToken, "ccMDSig", clsSignature.SignReason, 1, True
        End If
    End If
    If Not ccNurseSig.EmptySignature Then
        For i = 0 To cboAttendPhy.listcount - 1
            If cboAssistant.List(i) = cboAssistant.Text Then
                lngProvID = cboAssistant.ItemData(i)
                Exit For
            End If
        Next i
        lstrSignatureToken = ccNurseSig.EncodeSignatureValues(clsSignature.SignReason)
        SaveSignature lngProvID, lstrSignatureToken, "ccNurseSig", clsSignature.SignReason, 1, True
    End If
    Clear_Signature
    Screen.MousePointer = vbDefault
    Exit Function
Error_WriteClinicalData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, clsSignature.SignReason, "WriteClinicalData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Public Function Get_NurseProvider()
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strMedProvider As String
Dim cnMedProv As New ADODB.Connection
Dim rsMedProv As New ADODB.Recordset
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim strSuffix As String
Dim strCity As String
Dim strState As String
Dim strZip As String
Dim strZipExt As String
Dim booSkip As Boolean
Dim i%
Dim lngProvID As Long

    On Error GoTo Error_Get_NurseProvider
    Screen.MousePointer = vbHourglass
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    strSQL = "Load_Active_Providers"
    Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        Do Until .EOF
            Clear_Provider
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            lngProvID = 0
            booSkip = True
            If Not IsNull(!Title) Then
                strTitle = (!Title)
                If TextFound("R.N.", strTitle, Len("R.N."), True) Or _
                    TextFound("RN", strTitle, Len("RN"), True) Or _
                    TextFound("L.P.N.", strTitle, Len("L.P.N."), True) Or _
                    TextFound("LPN", strTitle, Len("LPN"), True) Or _
                    TextFound("P.A.", strTitle, Len("P.A."), True) Or _
                    TextFound("PA", strTitle, Len("PA"), True) Then
                        booSkip = False
                    End If
            End If
            If Not booSkip Then
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Fld.Value <> vbNullString Then
                            If UCase(Fld.Name) = "FIRST_NAME" Then
                                strFirst = Fld.Value
                            ElseIf UCase(Fld.Name) = "MIDDLE_NAME" Then
                                strMiddle = Fld.Value
                            ElseIf UCase(Fld.Name) = "LAST_NAME" Then
                                strLast = Fld.Value
                            ElseIf UCase(Fld.Name) = "SUFFIX" Then
                                strSuffix = Fld.Value
                            ElseIf UCase(Fld.Name) = "TITLE" Then
                                strTitle = Fld.Value
                            ElseIf UCase(Fld.Name) = "PROVIDER_ID" Then
                                lngProvID = Fld.Value
                            End If
                        End If
                    End If
                Next
            End If
            .MoveNext
            If Trim(strFirst) <> vbNullString Then
                 clsProvider.Name = Trim(strFirst)
             End If
            If Trim(strMiddle) <> vbNullString Then
                strMiddle = IsProperCase(strMiddle)
                 clsProvider.Name = clsProvider.Name & " " & Left(strMiddle, 1)
             End If
            If Trim(strLast) <> vbNullString Then
                 clsProvider.Name = clsProvider.Name & " " & Trim(strLast)
             End If
            If Trim(strSuffix) <> vbNullString And Trim(clsProvider.Name) <> vbNullString Then
                 clsProvider.Name = clsProvider.Name & ", " & Trim(strSuffix)
             End If
            If Trim(strTitle) <> vbNullString And Trim(clsProvider.Name) <> vbNullString Then
                 clsProvider.Name = clsProvider.Name & ", " & Trim(strTitle)
                 clsProvider.Title = Trim(strTitle)
             End If
            If clsProvider.Name <> vbNullString Then
                cboAssistant.AddItem clsProvider.Name
                If IsNumeric(lngProvID) Then cboAssistant.ItemData(cboAssistant.NewIndex) = lngProvID
            End If
        Loop
    End With

    If cnMedProv.State = adStateOpen Then
        cnMedProv.Close
        Set cnMedProv = Nothing
        Set rsMedProv = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_NurseProvider:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_NurseProvider", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cmdSave_Click()
    WriteClinicalData
    Unload Me
    Exit Sub
End Sub


Private Sub Form_Activate()

    If mFirst Then
        mFirst = False
        Clear_EnctrString
        ReadClinicalData
    End If
    Exit Sub
End Sub
Private Function Get_CurrentEnctr()
Dim cnEncnt As New ADODB.Connection
Dim rsEncnt As New ADODB.Recordset
Dim Fld As ADODB.Field
Dim Flds As ADODB.Fields
Dim strSQL As String

    On Error GoTo Error_Get_CurrentEnctr
    Screen.MousePointer = vbHourglass
    Clear_EnctrString
    If cnEncnt.State = adStateClosed Then cnEncnt.open psConnect(0)
    strSQL = "Select * FROM ENCOUNTER WHERE ENCOUNTER_ID = " & piSignEnctr & " AND PURGE = 0 "
    Set rsEncnt = cnEncnt.Execute(strSQL)
    If Not rsEncnt.EOF Then
        Set Flds = rsEncnt.Fields
        For Each Fld In Flds
            If Not IsNull(Fld.Value) Then
                If Trim(Fld.Value) <> vbNullString Then
                    If UCase(Fld.Name) = "FACILITY_ID" Then
                        udtEnctr.FacID = Fld.Value
                    ElseIf UCase(Fld.Name) = "ENCNTR_START_DATE" Then
                         udtEnctr.StartDate = Fld.Value
                    ElseIf UCase(Fld.Name) = "PROVIDER_ID" Then
                         udtEnctr.ProviderID = Fld.Value
                    End If
                End If
            End If
        Next
    End If
    If cnEncnt.State = adStateOpen Then
        cnEncnt.Close
        Set cnEncnt = Nothing
        Set rsEncnt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_CurrentEnctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_CurrentEnctr", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub Form_Load()
    mFirst = True
    Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Clear_EnctrString
    ccMDSig.ClearSignature
    ccNurseSig.ClearSignature
    Unload Me
    Exit Sub
End Sub
