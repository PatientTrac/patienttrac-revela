VERSION 5.00
Object = "{13ADDE2B-B303-4618-975F-648031250148}#5.0#0"; "PolarSpellChecker.dll"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmComment_All 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5880
   ClientLeft      =   9555
   ClientTop       =   4125
   ClientWidth     =   10230
   Icon            =   "frmComment_All.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   10230
   Begin VB.TextBox cctDate 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   4500
      TabIndex        =   11
      Top             =   60
      Width           =   2385
   End
   Begin VB.Frame fraInstr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3645
      Left            =   1425
      TabIndex        =   5
      ToolTipText     =   "Click to move around screen"
      Top             =   705
      Visible         =   0   'False
      Width           =   7305
      Begin PT_XP_Button.PT_XP cmdSaveInstruct 
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   3270
         Width           =   885
         _ExtentX        =   1561
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
         Picture         =   "frmComment_All.frx":058A
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Select"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmComment_All.frx":0924
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdInstruct 
         Height          =   3075
         Left            =   0
         TabIndex        =   7
         Top             =   150
         Width           =   7215
         _Version        =   458752
         _ExtentX        =   12726
         _ExtentY        =   5424
         _StockProps     =   64
         ColHeaderDisplay=   0
         DAutoCellTypes  =   0   'False
         DAutoFill       =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridColor       =   8388608
         MaxCols         =   2
         MaxRows         =   0
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         ShadowText      =   8388608
         SpreadDesigner  =   "frmComment_All.frx":0C76
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
         CellNoteIndicator=   3
      End
      Begin PT_XP_Button.PT_XP cmdTermMaint 
         Height          =   255
         Left            =   6780
         TabIndex        =   8
         ToolTipText     =   "Terms Maintanence Units"
         Top             =   3270
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   450
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
         MaskColor       =   128
         Picture         =   "frmComment_All.frx":0FD3
         AppearanceThemes=   3
         BorderWidth     =   2
         TransparentColor=   14737632
         Caption         =   ".."
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmComment_All.frx":156D
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCloseInstr 
         Height          =   285
         Left            =   3750
         TabIndex        =   9
         Top             =   3270
         Width           =   915
         _ExtentX        =   1614
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
         Picture         =   "frmComment_All.frx":18BF
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmComment_All.frx":22D1
         ShowFocus       =   -1  'True
      End
   End
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
      Height          =   4680
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   9990
   End
   Begin SPELLCHECKERLibCtl.SpellChecker SpellChecker1 
      Left            =   510
      OleObjectBlob   =   "frmComment_All.frx":2623
      Top             =   5220
   End
   Begin PT_XP_Button.PT_XP cmdErase 
      Height          =   345
      Left            =   4560
      TabIndex        =   0
      Top             =   5370
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
      Picture         =   "frmComment_All.frx":2647
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Erase"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmComment_All.frx":2BE1
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdAdd 
      Height          =   345
      Left            =   3570
      TabIndex        =   1
      Top             =   5370
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
      Picture         =   "frmComment_All.frx":2F33
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmComment_All.frx":34CD
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRemove 
      Height          =   345
      Left            =   5700
      TabIndex        =   2
      Top             =   5370
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
      Picture         =   "frmComment_All.frx":381F
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmComment_All.frx":4231
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSpell 
      Height          =   345
      Left            =   8670
      TabIndex        =   3
      ToolTipText     =   "Spell Check"
      Top             =   5370
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
      Picture         =   "frmComment_All.frx":4583
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Spell Check"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmComment_All.frx":4B1D
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPrevEncounter 
      Height          =   465
      Left            =   7830
      TabIndex        =   10
      Top             =   60
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Show Previous Encounter"
      CaptionPosition =   4
      DepthMode       =   1
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   465
      Left            =   180
      TabIndex        =   13
      Top             =   30
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
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
      Picture         =   "frmComment_All.frx":4E6F
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "INDENT"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmComment_All.frx":5209
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblVisit 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Visit Date"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   12
      Top             =   120
      Width           =   1425
   End
End
Attribute VB_Name = "frmComment_All"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrCommentItem As String
Private mFirst  As Boolean
Public mbooLoad As Boolean
Public CommentType As String


Private Function Get_Recommendations()
Dim strSQL As String
Dim cnRead As New ADODB.Connection
Dim rsRead As New ADODB.Recordset
Dim strStart As String

    On Error GoTo Error_Get_Recommendations
    Screen.MousePointer = vbHourglass
    cctDate.Text = vbNullString
    If clsEncounter.StartDate <> vbNullString Then
        strStart = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    Else
         strStart = Format(Now, "mm/dd/yyyy")
    End If
    If cmdPrevEncounter.Caption = "Show Current Encounter" Then
        strStart = DateAdd("d", -1, strStart)
    End If
    If cnRead.State = adStateClosed Then cnRead.open psConnect(0)
    strSQL = "SELECT NOTES_DESC, ENCOUNTER.ENCNTR_START_DATE  FROM INPATIENT_NOTES " _
        & "INNER JOIN ENCOUNTER ON ENCOUNTER.ENCOUNTER_ID = INPATIENT_NOTES.ENCOUNTER_ID " _
        & "WHERE INPATIENT_NOTES.PATIENT_ID = " & piPatientID & " AND " _
        & "CONVERT(DATETIME, ENCOUNTER.ENCNTR_START_DATE, 101)  <  '" & Format(strStart, "mm/dd/yyyy") & "' AND " _
        & "NOTE_TYPE like 'RECOMMENDATION%' AND NOTES_DESC IS NOT NULL ORDER BY INPATIENT_NOTES.INSERT_DATE DESC"
    Set rsRead = cnRead.Execute(strSQL)
    With rsRead
        Do Until .EOF
            If Trim(!NOTES_DESC) <> vbNullString Then
                cctComments.Text = Trim((!NOTES_DESC))
                If IsDate((!ENCNTR_START_DATE)) Then
                    cctDate.Text = Format((!ENCNTR_START_DATE), "MMMM DD, YYYY")
                End If
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    cctComments.Text = Replace(cctComments.Text, "''", "'")
    Screen.MousePointer = vbDefault
    If cnRead.State = adStateOpen Then
        cnRead.Close
        Set cnRead = Nothing
        Set rsRead = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Recommendations:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Recommendations", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_Add_Notes()
Dim strSQL As String
Dim cnRead As New ADODB.Connection
Dim rsRead As New ADODB.Recordset
Dim strStart As String

    On Error GoTo Error_Get_Add_Notes
    Screen.MousePointer = vbHourglass
    cctDate.Text = vbNullString
    If clsEncounter.StartDate <> vbNullString Then
        strStart = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    Else
         strStart = Format(Now, "mm/dd/yyyy")
    End If
    If cmdPrevEncounter.Caption = "Show Current Encounter" Then
        strStart = DateAdd("d", -1, strStart)
    End If
    If cnRead.State = adStateClosed Then cnRead.open psConnect(0)
    strSQL = "SELECT NOTES_DESC, ENCOUNTER.ENCNTR_START_DATE  FROM INPATIENT_NOTES " _
        & "INNER JOIN ENCOUNTER ON ENCOUNTER.ENCOUNTER_ID = INPATIENT_NOTES.ENCOUNTER_ID " _
        & "WHERE INPATIENT_NOTES.PATIENT_ID = " & piPatientID & " AND " _
        & "CONVERT(DATETIME,  ENCOUNTER.ENCNTR_START_DATE, 101)  <  '" & Format(strStart, "mm/dd/yyyy") & "' AND " _
        & "NOTE_TYPE like 'ADDITIONAL NOTE%' AND NOTES_DESC IS NOT NULL ORDER BY INPATIENT_NOTES.INSERT_DATE DESC"
    Set rsRead = cnRead.Execute(strSQL)
    With rsRead
        Do Until .EOF
            If Trim(!NOTES_DESC) <> vbNullString Then
                cctComments.Text = Trim((!NOTES_DESC))
                If IsDate((!ENCNTR_START_DATE)) Then
                    cctDate.Text = Format((!ENCNTR_START_DATE), "MMMM DD, YYYY")
                End If
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    cctComments.Text = Replace(cctComments.Text, "''", "'")
    Screen.MousePointer = vbDefault
    If cnRead.State = adStateOpen Then
        cnRead.Close
        Set cnRead = Nothing
        Set rsRead = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Add_Notes:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Add_Notes", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function



Private Function Format_Terms()
'Dim varColName
'Dim lngRow As Long
'Dim strList As String
'Dim lngX As Long
'Dim lngY As Long
'Dim lngWidth As Long
'Dim lngHeight As Long
'
'    On Error GoTo Error_Format_Goals
'    Screen.MousePointer = vbHourglass
'
'    With grdInstruct
'        .Redraw = False
'        .ClearRange 1, SpreadHeader, grdInstruct.MaxCols, grdInstruct.MaxRows, False
'        .ProcessTab = True
'        .FontBold = True
'        .ColHeaderRows = 2
'        .ColHeadersShow = True
'        .RowHeadersShow = False
'        .MaxCols = 1
'        .MaxRows = 0
'        Select Case clsEducation.PlanType
'            Case "Education"
'                clsEducation.Title = "PATIENT INFORM AND EDUCATION"
'            Case "Instructions"
'                clsEducation.Title = "PATIENT INSTRUCTIONS"
'        End Select
'        Format_CellTypeEdit grdInstruct, SpreadHeader, 1, "1", clsEducation.Title, "TypeHAlignCenter", "", "", HeaderColor
'        grdInstruct.SetCellBorder 1, lngRow, 1, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'        grdInstruct.SetCellBorder 2, lngRow, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'        Add_Spacer grdInstruct, SpreadHeader + 1
'        Get_Edu_Terms
'        .Redraw = True
'    End With
'    Screen.MousePointer = vbDefault
'    Exit Function
'Error_Format_Goals:
'    Exit Function
'    Resume Next
End Function
Private Function Get_Edu_Terms()
'Dim strSQL As String
'Dim rsTerms As New ADODB.Recordset
'Dim cnTerms As New ADODB.Connection
'Dim Flds As ADODB.Fields
'Dim Fld As ADODB.Field
'Dim lngRow As Long
'Dim varCellValue
'Dim Arr
'Dim Idx As Integer
'Dim strList As String
'Dim lngInterVent As Long
'
'    On Error GoTo Error_Get_Edu_Terms
'    Screen.MousePointer = vbHourglass
'
'    Select Case clsEducation.PlanType
'        Case "Education"
'            strList = "sd_PsyPlan_Education"
'        Case "Instructions"
'            strList = "sd_PsyPlan_Instruct"
'    End Select
'    If cnTerms.State = adStateClosed Then cnTerms.Open mstrConnectStringTerms
'    strSQL = "Select DISTINCT ATTRBTE_VALID_VALUE, valid_value_seq_nbr " _
'            & "From ATTR_VALID_VALUE " _
'            & "Where VALID_VALUE_SUBDOM = '" & strList & "' ORDER BY valid_value_seq_nbr"
'        Set rsTerms = cnTerms.Execute(strSQL)
'        With rsTerms
'            Do Until .EOF
'                If (!ATTRBTE_VALID_VALUE) <> vbNullString Then
'                    lngRow = lngRow + 1
'                    grdInstruct.MaxRows = lngRow
'                    Format_CellTypeStatic grdInstruct, lngRow, 1, "1", Trim(!ATTRBTE_VALID_VALUE), "TypeHAlignLeft", "TypeVAlignCenter"
'                    grdInstruct.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                End If
'                .MoveNext
'            Loop
'        End With
'    If cnTerms.State = adStateOpen Then
'        cnTerms.Close
'        Set rsTerms = Nothing
'        Set cnTerms = Nothing
'        Set Flds = Nothing
'    End If
'    Screen.MousePointer = vbDefault
'    Exit Function
'Error_Get_Edu_Terms:
'    Screen.MousePointer = vbDefault
'    Write_Error Err.Description, Err.Number, "", "Get_Edu_Terms", App.EXEName & "." & TypeName(Me)
'    Exit Function
'    Resume Next
End Function



Private Sub cctComments_Click()
    mbooHasChanged = True
End Sub

Private Sub cmdAdd_Click()
Dim iAnswer As Integer

    mstrCommentItem = vbNullString
    If cmdPrevEncounter.Caption = "Show Current Encounter" Then
        iAnswer = MsgBox("Do you want to save the previous encounter data into this current encounter?  If not please select Current Encounter First!", vbYesNo, "PatientTrac Electronic Medical Records")
        If iAnswer = vbYes Then
            mstrCommentItem = cctComments.Text
            Unload Me
        Else
            Exit Sub
        End If
    Else
        mstrCommentItem = cctComments.Text
        Unload Me
    End If
    Exit Sub
End Sub

Private Sub cmdCloseInstr_Click()
    fraInstr.Visible = False
End Sub

Private Sub cmdErase_Click()
    cctComments.Text = vbNullString
    Exit Sub
End Sub

Private Sub cmdPrevEncounter_Click()
    If cmdPrevEncounter.Caption = "Show Previous Encounter" Then
        cctComments.Text = vbNullString
        cmdPrevEncounter.Caption = "Show Current Encounter"
        Select Case CommentType
            Case "Additional Notes"
                Get_Add_Notes
            Case "Recommendation"
                Get_Recommendations
        End Select
    ElseIf cmdPrevEncounter.Caption = "Show Current Encounter" Then
        cctComments.Text = vbNullString
        cctComments.Text = mstrCommentItem
        If clsEncounter.StartDate <> vbNullString Then
            cctDate.Text = Format(clsEncounter.StartDate, "MMMM DD, YYYY")
        End If
        cmdPrevEncounter.Caption = "Show Previous Encounter"
    End If
    Exit Sub
End Sub

Private Sub cmdRemove_Click()
    mstrCommentItem = vbNullString
    cctComments.Text = vbNullString
    Unload Me
End Sub

Private Sub cmdSaveInstruct_Click()
'Dim lngRow As Long
'Dim varCellValue
'Dim lngCol As Long
'Dim strCollection As String
'Dim lngPrimKey As Long
'Dim strDate As String
'Dim cnGoal As New ADODB.Connection
'Dim rsGoal As New ADODB.Recordset
'Dim strSQL As String
'
'    Screen.MousePointer = vbDefault
'    On Error GoTo Error_cmdSaveInstruct_Click
'    clsEducation.Education = vbNullString
'
'    strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
'    With grdInstruct
'        For lngRow = 1 To .MaxRows
'            .Col = 1
'            .Row = lngRow
'            If .BackColor = grdCurItem Then
'                .GetText 1, lngRow, varCellValue
'                If Trim(varCellValue) <> vbNullString Then
'                    If Trim(clsEducation.Education) = vbNullString Then
'                        clsEducation.Education = Trim(varCellValue)
'                    Else
'                        clsEducation.Education = clsEducation.Education & vbNewLine & Trim(varCellValue)
'                    End If
'                End If
'            End If
'        Next lngRow
'    End With
'    With grdInstruct
'        .Redraw = False
'        For lngRow = 1 To .MaxRows
'            .BlockMode = True
'                .Col = 1
'                .Col2 = .MaxCols
'                .Row = lngRow
'                .Row2 = lngRow
'                .BackColor = NormalColor
'             .BlockMode = False
'         Next lngRow
'        .Redraw = True
'        .ClearRange 1, SpreadHeader, .MaxCols, .MaxRows, False
'    End With
'    Select Case clsEducation.PlanType
'        Case "Education"
'            cctEducation.Text = Trim(clsEducation.Education)
'        Case "Instructions"
'            cctNotes.Text = Trim(clsEducation.Education)
'    End Select
'    fraInstr.Visible = False
'    clsEducation.PlanType = vbNullString
'    clsEducation.Education = vbNullString
'    Screen.MousePointer = vbDefault
'    Exit Sub
'Error_cmdSaveInstruct_Click:
'    Screen.MousePointer = vbDefault
'    Write_Error Err.Description, Err.Number, "", "cmdSaveInstruct_Click", App.EXEName & "." & TypeName(Me)
'    Exit Sub
'    Resume Next

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
    cctComments.Text = strResultText
    SpellChecker1.CloseDictionary mstrProgram & "\FirstNotebook\PatientTrac EMR\Dictionary\SpellUS.lex"
    Exit Sub
End Sub

Private Sub cmdTermMaint_Click()
'Dim strSQL As String
'Dim cnCareIntrvn As New ADODB.Connection
'Dim rsCareIntrvn As New ADODB.Recordset
'Dim lngCareID As Long
'Dim strList As String
'Dim strCaption As String
'
'    On Error GoTo Error_Goal_Terms
'    Select Case clsEducation.PlanType
'        Case "Education"
'            strList = "sd_PsyPlan_Education"
'            strCaption = "PATIENT EDUCATION"
'        Case "Instructions"
'            strList = "sd_PsyPlan_Instruct"
'            strCaption = "PATIENT INSTRUCTIONS"
'    End Select
'    Load frmNurseTemplates
'    frmNurseTemplates.fCaption = strCaption
'    frmNurseTemplates.fType = strCaption
'    frmNurseTemplates.fSubDomain = strList
'    frmNurseTemplates.fConnectCR = mstrConnectStringClinical
'    frmNurseTemplates.fConnectTM = mstrConnectStringTerms
'    frmNurseTemplates.ReadData
'    frmNurseTemplates.Show vbModal
'    Format_Terms
'    Get_Edu_Terms
'    Read_InPatientNote clsEducation.PlanType
'    Screen.MousePointer = vbDefault
'    Exit Sub
'Error_Goal_Terms:
'    Screen.MousePointer = vbDefault
'    Write_Error Err.Description, Err.Number, strSQL, "Goal_Terms", App.EXEName & "." & TypeName(Me)
'    Exit Sub
'    Resume Next
End Sub

Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional vlngBackColor As Long)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        mbooByPass = True
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeEdit
            .TypeMaxEditLen = 32000
            .TypeEditMultiLine = True   'Make multi-lines
            .FontName = "Times New Roman"
            .FontSize = 8
            .FontBold = True
            .TypeVAlign = TypeVAlignCenter
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If vstrCellNote <> vbNullString Then
                .CellNote = vstrCellNote
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            End If
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Col = vlngCol
            .Text = vstrText
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote + 7
            If vlngBackColor > 0 Then .BackColor = vlngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
        mbooByPass = False
End Function

Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, vstrVAlign As String, Optional lngBackColor As Long, Optional vstrCellNote As String, Optional vstrCellTag As String, Optional vlngFontSize As Long)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeStaticText
            .TypeMaxEditLen = 32000
            .FontName = "Times New Roman"
            If vlngFontSize > 0 Then
                .FontSize = vlngFontSize
            Else
                .FontSize = 8
            End If
            .FontBold = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            If vstrVAlign = "TypeVAlignCenter" Then
                .TypeVAlign = TypeVAlignCenter
            ElseIf vstrVAlign = "TypeVAlignBottom" Then
                .TypeVAlign = TypeVAlignBottom
            Else
               .TypeVAlign = TypeVAlignTop
            End If
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            ElseIf vstrHAligh = "TypeHAlignRight" Then
                .TypeHAlign = TypeHAlignRight
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If vstrCellNote <> vbNullString Then
                .CellNote = vstrCellNote
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
            If lngBackColor > 0 Then .BackColor = lngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        .BlockMode = False
    End With
End Function

 


Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
        cctComments.Text = mstrCommentItem
        centerForm Me
        Screen.MousePointer = vbDefault
        If clsEncounter.StartDate <> vbNullString Then
            cctDate.Text = Format(clsEncounter.StartDate, "MMMM DD, YYYY")
        End If
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
  On Error Resume Next
    Unload Me
    Exit Sub
End Sub


Private Sub grdInstruct_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    With grdInstruct
        If Row >= 0 Then
            .BlockMode = True
                .Col = Col
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                 .BackColor = grdCurItem
             .BlockMode = False
        End If
    End With
    Exit Sub
End Sub


Private Sub grdInstruct_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row <= 0 Then Exit Sub
    With grdInstruct
        .BlockMode = True
            .Col = Col
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
             .BackColor = NormalColor
         .BlockMode = False
    End With
    Exit Sub
End Sub


