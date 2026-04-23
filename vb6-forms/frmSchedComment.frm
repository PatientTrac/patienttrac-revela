VERSION 5.00
Object = "{13ADDE2B-B303-4618-975F-648031250148}#5.0#0"; "PolarSpellChecker.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmSchedComment 
   Caption         =   "InPatient Roster Notes"
   ClientHeight    =   5790
   ClientLeft      =   10260
   ClientTop       =   3765
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchedComment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   10200
   Begin VB.TextBox cctComments 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   660
      Width           =   9990
   End
   Begin SPELLCHECKERLibCtl.SpellChecker SpellChecker1 
      Left            =   390
      OleObjectBlob   =   "frmSchedComment.frx":08CA
      Top             =   5190
   End
   Begin PT_XP_Button.PT_XP cmdErase 
      Height          =   345
      Left            =   4440
      TabIndex        =   1
      Top             =   5340
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmSchedComment.frx":08EE
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Erase"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedComment.frx":0E88
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdAdd 
      Height          =   345
      Left            =   3450
      TabIndex        =   2
      Top             =   5340
      Width           =   885
      _ExtentX        =   1561
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
      Picture         =   "frmSchedComment.frx":11DA
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedComment.frx":1774
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRemove 
      Height          =   345
      Left            =   5580
      TabIndex        =   3
      Top             =   5340
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
      Picture         =   "frmSchedComment.frx":1AC6
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Close"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedComment.frx":24D8
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSpell 
      Height          =   345
      Left            =   8550
      TabIndex        =   4
      ToolTipText     =   "Spell Check"
      Top             =   5340
      Width           =   1485
      _ExtentX        =   2619
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
      Picture         =   "frmSchedComment.frx":282A
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Spell Check"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedComment.frx":2DC4
      ShowFocus       =   -1  'True
   End
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   345
      Left            =   8130
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   16761024
      CalendarTitleForeColor=   8388608
      CalendarTrailingForeColor=   8421504
      CheckBox        =   -1  'True
      CustomFormat    =   "MM/dd/yyy"
      DateIsNull      =   -1  'True
      Format          =   86573059
      CurrentDate     =   38792
   End
   Begin VB.Label lblADate 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Status Comment Date:"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5850
      TabIndex        =   6
      Top             =   210
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmSchedComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst  As Boolean
Public mbooLoad As Boolean
Public mstrNotes As String
Private mlngStatus As Long
Public mstrType As String




Private Function Write_Note()
Dim strSQL As String
Dim cnWrite As New ADODB.Connection
Dim rsWrite As New ADODB.Recordset
Dim lngGridPrim As Long
Dim strData As String
Dim Arr
Dim Idx%


    On Error GoTo Error_Write_Note
    Screen.MousePointer = vbHourglass
    mstrNotes = vbNullString
    If TextFound(Trim(clsInsurance.Prim_Insurer), Trim(cctComments.Text), Len(Trim(clsInsurance.Prim_Insurer)), True) Then
         cctComments.Text = Replace(cctComments.Text, Trim(UCase(clsInsurance.Prim_Insurer)), "")
         cctComments.Text = Replace(cctComments.Text, "Primary Insurance:", "")
    ElseIf TextFound("No record for primary insurance", udtInPatSched.ApptComment, Len("No record for primary insurance"), True) Then
         cctComments.Text = Replace(cctComments.Text, "No record for primary insurance", "")
    End If
    Arr = Split(cctComments, vbNewLine)
    If UBound(Arr) > 0 Then
        cctComments = vbNullString
        For Idx = LBound(Arr) To UBound(Arr)
            If Arr(Idx) <> vbNullString And Len(Arr(Idx)) > 3 Then
                If cctComments = vbNullString Then
                    cctComments = Trim(Arr(Idx))
                Else
                    cctComments = cctComments & vbNewLine & Trim(Arr(Idx))
                End If
            End If
        Next Idx
    Else
        cctComments = Trim(cctComments)
    End If
    mstrNotes = cctComments
    If TextFound(clsInsurance.Prim_Insurer, mstrNotes, Len(clsInsurance.Prim_Insurer), True) Then
        mstrNotes = Replace(mstrNotes, clsInsurance.Prim_Insurer, "")
    End If
    strData = cctComments
    Arr = Split(strData, vbNewLine)
    If UBound(Arr) > 0 Then
        strData = vbNullString
        For Idx = LBound(Arr) To UBound(Arr)
            If Arr(Idx) <> vbNullString And Not TextFound(Arr(Idx), strData, Len(Arr(Idx)), True) Then
                If strData = vbNullString Then
                    strData = Trim(Arr(Idx))
                Else
                    strData = strData & vbNewLine & Trim(Arr(Idx))
                End If
            End If
        Next Idx
    End If
    strData = Replace(strData, "'", "''")
    If cnWrite.State = adStateClosed Then cnWrite.open psConnect(0)
    If mlngStatus > 0 Then
        strSQL = "Update_Patient_Status " & udtInPatSched.Facility_ID & ", '" _
            & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "', '" & Trim(strData) & "'," & mlngStatus
        Set rsWrite = cnWrite.Execute(strSQL)
    Else
        strSQL = "Insert_Patient_Status " & piPatientID & "," _
            & udtInPatSched.Facility_ID & ", '" & Trim(strData) & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'"
        Set rsWrite = cnWrite.Execute(strSQL)
    End If
    If cnWrite.State = adStateOpen Then
        cnWrite.Close
        Set cnWrite = Nothing
        Set rsWrite = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Note:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Note", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function
Private Function Get_Status_Notes()
Dim strSQL As String
Dim cnRead As New ADODB.Connection
Dim rsRead As New ADODB.Recordset
Dim strDate As String
Dim strUpdate As String
Dim strNote As String
Dim strUpNote As String
Dim Arr
Dim Idx%

    On Error GoTo Error_Get_Status_Notes
    Screen.MousePointer = vbHourglass
    mlngStatus = 0
    If cnRead.State = adStateClosed Then cnRead.open psConnect(0)
    strSQL = "Get_Patient_Status_PatientID  " & piPatientID
    Set rsRead = cnRead.Execute(strSQL)
    With rsRead
        If Not .EOF Then
            strDate = (!INSERT_DATE)
            If Not IsNull((!STATUS_NOTES)) Then strNote = Replace(Trim((!STATUS_NOTES)), "''", "'")
            mlngStatus = (!PATIENT_STATUS_ID)
        End If
    End With
    Get_Insurance Val(udtInPatSched.IntId)
    If Trim(clsInsurance.Prim_Insurer) <> vbNullString Then
        If Not TextFound(clsInsurance.Prim_Insurer, cctComments.Text, Len(clsInsurance.Prim_Insurer), True) Then
            If cctComments.Text = vbNullString Then
                cctComments.Text = "Primary Insurance:" & " " & Trim(clsInsurance.Prim_Insurer)
            Else
                cctComments.Text = clsInsurance.Prim_Insurer & vbNewLine & Trim(cctComments.Text)
            End If
        End If
    Else
        If cctComments.Text = vbNullString Then
            cctComments.Text = "No record for primary insurance"
        Else
            cctComments.Text = "No record for primary insurance" & vbNewLine & Trim(cctComments.Text)
        End If
    End If
    If Trim(strNote) <> vbNullString Then
        If cctComments.Text <> vbNullString Then
            cctComments.Text = cctComments.Text & vbNewLine & Replace(strNote, "''", "'")
        Else
            cctComments.Text = Replace(strNote, "''", "'")
        End If
    End If
'------
    Arr = Split(cctComments, vbNewLine)
    If UBound(Arr) > 0 Then
        cctComments = vbNullString
        For Idx = LBound(Arr) To UBound(Arr)
            If Arr(Idx) <> vbNullString And Not TextFound(Arr(Idx), cctComments, Len(Arr(Idx)), True) Then
                If cctComments = vbNullString Then
                    cctComments = Trim(Arr(Idx))
                Else
                    cctComments = cctComments & vbNewLine & Trim(Arr(Idx))
                End If
            End If
        Next Idx
    End If
'----------
    txtDate.Value = Format(strDate, "mm/dd/yyyy")
    If mstrType = "Appointment" Then
        udtInPatSched.Appointment_ID = 0
        strSQL = "SELECT * FROM INPATIENT_APPOINTMENTS  " _
            & "WHERE PATIENT_ID =  " & Val(udtInPatSched.IntId) & " AND " _
            & "FACILITY_ID = " & udtInPatSched.Facility_ID & " ORDER BY INSERT_DATE DESC"
        Set rsRead = cnRead.Execute(strSQL)
        With rsRead
            If Not .EOF Then
                'If Not IsNull((!INSERT_DATE)) Then
                    'If Trim((!INSERT_DATE)) <> vbNullString Then
                        'strDate = Format((!INSERT_DATE), "mm/dd/yyyy")
                    'End If
                   udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
                'End If
            End If
        End With
        If udtInPatSched.Appointment_ID > 0 Then
            strSQL = "SELECT REASON FROM INPATIENT_APPOINTMENTS WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
            With rsRead
                If Not .EOF Then
                    If Not IsNull((!Reason)) Then
                        If Trim((!Reason)) <> vbNullString Then
                            If cctComments = vbNullString Then
                                cctComments = Trim((!Reason))
                            Else
                                cctComments = cctComments & vbNewLine & Trim((!Reason))
                            End If
                        End If
                    End If
                End If
            End With
        End If
    End If
    Screen.MousePointer = vbDefault
    If cnRead.State = adStateOpen Then
        cnRead.Close
        Set cnRead = Nothing
        Set rsRead = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Status_Notes:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Status_Notes", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cctComments_Click()
    mbooHasChanged = True
End Sub

Private Sub cmdAdd_Click()
    Write_Note
    Unload Me
    Exit Sub
End Sub


Private Sub cmdErase_Click()
    cctComments.Text = vbNullString
    Exit Sub
End Sub

Private Sub cmdRemove_Click()
    Unload Me
End Sub

Private Sub cmdSpell_Click()
Dim strResultText As String
Dim vExitCode As Variant
    SpellChecker1.OpenDictionary mstrProgram & "\FirstNotebook\PatientTrac EMR\Dictionary\SpellUS.lex"
    cctComments.SetFocus
    strResultText = SpellChecker1.CheckText(cctComments.Text, vExitCode)
    If Trim(cctComments.Text) = Trim(strResultText) Then
        MsgBox "No spelling errors found!", vbInformation, "PatientTrac Spell Checker"
    End If
    cctComments.Text = Trim(strResultText)
    SpellChecker1.CloseDictionary mstrProgram & "\FirstNotebook\PatientTrac EMR\Dictionary\SpellUS.lex"
    Exit Sub
End Sub

Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
        Get_Status_Notes
        centerForm Me
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub Form_Load()
    mFirst = True
    Screen.MousePointer = vbDefault
End Sub
 


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
 cctComments.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next
    Exit Sub
End Sub


