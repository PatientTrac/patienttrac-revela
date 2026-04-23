VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmCoding 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SuperBill Items"
   ClientHeight    =   9150
   ClientLeft      =   5295
   ClientTop       =   2490
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frmCoding.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9150
   ScaleWidth      =   9285
   Begin VB.ListBox lstSubGroup 
      ForeColor       =   &H00800000&
      Height          =   1020
      Left            =   2790
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   0
      Width           =   4605
   End
   Begin FPSpreadADO.fpSpread grdSuperBill 
      Height          =   6525
      Left            =   -30
      TabIndex        =   0
      Top             =   1530
      Width           =   9285
      _Version        =   458752
      _ExtentX        =   16431
      _ExtentY        =   11509
      _StockProps     =   64
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16761024
      MaxCols         =   25
      MaxRows         =   25
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      SpreadDesigner  =   "frmCoding.frx":09EA
      Appearance      =   1
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   405
      Left            =   3240
      TabIndex        =   4
      Top             =   8400
      Width           =   915
      _ExtentX        =   1614
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
      Picture         =   "frmCoding.frx":0C13
      AppearanceThemes=   3
      BorderDrawEvent =   1
      TransparentColor=   14737632
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmCoding.frx":11AD
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   405
      Left            =   4830
      TabIndex        =   5
      ToolTipText     =   "Close and do not Save"
      Top             =   8400
      Width           =   915
      _ExtentX        =   1614
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
      Picture         =   "frmCoding.frx":14FF
      AppearanceThemes=   3
      BorderDrawEvent =   1
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmCoding.frx":1F11
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblGroupItem 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Group Items"
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3210
      TabIndex        =   3
      Top             =   1140
      Width           =   2685
   End
   Begin VB.Label lblBillGroup 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Billing Group"
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   705
      TabIndex        =   2
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "frmCoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngRow As Long
Public mlngcol As Long
Public mFirst As Boolean
Public mintMSE_EM As Integer
Public mlngEncounterID As Long

Private Sub Format_Grid()
Dim CellNote As Long

'With grdSuperBill
'
'    .ProcessTab = True
'    .Width = 16000
'    .FontBold = True
'    .FontName = "Georgia"
'    .GridSolid = False
'    .GridShowHoriz = True
'    .GridShowVert = True
'    .RowHeadersShow = True
'    .ColHeaderRows = 1
'    .ColHeadersShow = True
'    .MaxCols = 2
'    .MaxRows = 1
'
'    .BlockMode = True
'    .Col = 1
'    .Col2 = .MaxCols
'    .Row = SpreadHeader
'    .Row2 = SpreadHeader
'    .TypeHAlign = TypeHAlignCenter
'    .TypeVAlign = TypeVAlignCenter
'    .EditMode = False
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = .MaxCols
'        .Row = 1
'        .Row2 = .MaxRows
'        .FontBold = True
'        .FontName = "Georgia"
'        .TypeHAlign = TypeHAlignLeft
'        .TypeVAlign = TypeVAlignCenter
'        .FontBold = True
'    .BlockMode = False
''------------------------------------------------------------
'    ' Row 1   First Header Rows
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 0
'        .Row2 = 0
'        .ColWidth(.Col) = 25
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .TypeVAlign = TypeVAlignCenter
'        .RowHeight(1) = 15
'        .SetText 1, .Row, "SuperBillGropu"
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 2
'        .Col2 = 2
'        .Row = 0
'        .Row2 = 0
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .ColWidth(2) = 21
'        .SetText 2, .Row, "Completed Data"
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col = 1
'        .Row = SpreadHeader + 1
'        .Row2 = SpreadHeader + 1
'        .AddCellSpan 1, SpreadHeader + 1, 3, 1
'        .RowHeight(1) = 15
'        .SetText 1, SpreadHeader + 1, "Mental Status Examination"
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 3
'        .Col2 = 3
'        .Row = 0
'        .Row2 = 0
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .ColWidth(3) = 25
'        .SetText 3, .Row, "Scoring Value"
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 1
'        .Row2 = 1
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .Text = "Appearance" & vbNewLine & "     " & "Remarkable Features" & vbNewLine & "     " _
'            & "Habitus" & vbNewLine & "     " & "Patient Appears" & vbNewLine & "     " & "Nuritional Status"
'        cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'        .SetText 1, 1, "Appearance" & vbNewLine & "     " & "Remarkable Features" & vbNewLine & "     " _
'            & "Habitus" & vbNewLine & "     " & "Patient Appears" & vbNewLine & "     " & "Nuritional Status"
'        .SetText 2, 1, mvarAppearance
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 2
'        .Row2 = 2
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .Text = "Muscle Strength" & vbNewLine & "     " & "Muscle Tone"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'        .SetText 1, 2, "Muscle Strength" & vbNewLine & "     " & "Muscle Tone"
'        .SetText 2, 2, mvarMuscleStrgth
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 3
'        .Row2 = 3
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 3, "Gait" & vbNewLine & "     " & "Station"
'        .SetText 2, 3, mvarGait
'        .Text = "Gait" & vbNewLine & "     " & "Station"
'        cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 4
'        .Row2 = 4
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 4, "Orientation"
'        .SetText 2, 4, mvarOrient
'        .Text = "Orientation"
'        cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 5
'        .Row2 = 5
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 5, "Speech" & vbNewLine & "     " & "Voice"
'        .SetText 2, 5, mvarSpeech
'        .Text = "Speech" & vbNewLine & "     " & "Voice"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 6
'        .Row2 = 6
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 6, "Thought Process" & vbNewLine & "     " & "ProvInter" & vbNewLine & "     " & "Similarities"
'        .SetText 2, 6, mvarThoughtProc
'        .Text = "Thought Process" & vbNewLine & "     " & "ProvInter" & vbNewLine & "     " & "Similarities"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 7
'        .Row2 = 7
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 7, "Associations"
'        .SetText 2, 7, mvarAssociations
'        .Text = "Associations"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 8
'        .Row2 = 8
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 8, "Insight" & vbNewLine & "     " & "Judgment"
'        .SetText 2, 8, mvarInsight
'        .Text = "Insight" & vbNewLine & "     " & "Judgment"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 9
'        .Row2 = 9
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 9, "Naming" & vbNewLine & "     " & "Repitition"
'        .SetText 2, 9, mvarNaming
'        .Text = "Naming" & vbNewLine & "     " & "Repitition"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 10
'        .Row2 = 10
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 10, "Mood" & vbNewLine & "     " & "Affect"
'        .SetText 2, 10, mvarMood
'        .Text = "Mood" & vbNewLine & "     " & "Affect"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 11
'        .Row2 = 11
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 11, "Registration" & vbNewLine & "     " & "Serial 7" & vbNewLine & "     " & "Spell"
'        .SetText 2, 11, mvarRegistration
'        .Text = "Registration" & vbNewLine & "     " & "Serial 7" & vbNewLine & "     " & "Spell"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 12
'        .Row2 = 12
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 12, "Delusions" & vbNewLine & "     " & "Hallucinations" & vbNewLine _
'            & "     " & "Suicide" & vbNewLine & "     " & "Homicide"
'        .SetText 2, 12, mvarDelusions
'        .Text = "Delusions" & vbNewLine & "     " & "Hallucinations" & vbNewLine _
'            & "     " & "Suicide" & vbNewLine & "     " & "Homicide"
'        .SetText 2, 12, mvarDelusions
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 13
'        .Row2 = 13
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 13, "Fund of Knowledge"
'        .SetText 2, 13, mvarFundKnow
'        .Text = "Fund of Knowledge"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'
'    .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 14
'        .Row2 = 14
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 14, "Remote Memory" & vbNewLine & "     " & "Gross Assessment"
'        .SetText 2, 14, mvarRemMem
'        .Text = "Remote Memory" & vbNewLine & "     " & "Gross Assessment"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'     .BlockMode = True
'        .Col = 1
'        .Col2 = 1
'        .Row = 15
'        .Row2 = 15
'        .TypeEditMultiLine = True   'Make multi-line
'        .TypeTextWordWrap = True  'Make multi-line
'        .TypeMaxEditLen = 2000
'        .SetText 1, 15, "Vital Sign"
'        .SetText 2, 15, mvarVitals
'        .Text = "Vital Sign"
'         cellnote = .MaxTextCellHeight
'        .RowHeight(.Row) = cellnote
'    .BlockMode = False
'
'    .SetCellBorder 2, 1, 3, 15, 1, 0, CellBorderStyleSolid
'    .SetCellBorder 1, 1, .MaxCols, .MaxRows, 4, 0, CellBorderStyleSolid
'
'
'' Other Documented
'
'    .SetText 1, 18, "Verbal Auditory Content"
'    .SetText 2, 18, mvaVerbalAudi
'    .SetText 1, 19, "Attitude"
'    .SetText 2, 19, mvarAttitude
'    .SetText 1, 20, "Conation"
'    .SetText 2, 20, mvarConation
'    .SetText 1, 21, "Abnormal Movement"
'    .SetText 2, 21, mvarAbnorMove
'    .SetText 1, 22, "Intelligence Estimate"
'    .SetText 2, 22, mvarIntell
'    .SetText 1, 23, "Eye Contact"
'    .SetText 2, 23, mvarEyeContact
'
'
'
'    .Row = 16
'    .Col = 3
'    .Formula = "SUM(A1:A15)"
'    .Width = Me.Width
'
'End With
End Sub


Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
    Format_Grid
    'MSEEM_Value
    End If
End Sub

Private Sub Form_Load()
    mFirst = True
    PatientTracCodeLib.Set_ClinicalConnectString psConnect(0)
    PatientTracCodeLib.Set_TermsConnectString psConnect(1)
    Get_SubGroup
    Exit Sub
End Sub

Public Sub Add_GridItem(vcol As Long, vrow As Long, vvar As Variant)

     
End Sub

Public Sub MSEEM_Value()
'Dim intEM As Integer
'Dim strVitals As String
'Dim varEMCalc
'
'    mintMSE_EM = 0
'
'    With grdSuperBill
'        .Redraw = False
'        ' Appearance Value
'            If mvarAppearance = "1" Then .SetText 2, 1, "Appearance" & vbNewLine
'            If mvarRemarkFeatures = "1" Then .SetText 2, 1, "Remarkable Features" & vbNewLine
'            If mvarHabitus = "1" Then .SetText 2, 1, "Habitus" & vbNewLine
'            If mvarPatientApp = "1" Then .SetText 2, 1, "Patient Appears" & vbNewLine
'            If mvarNuritionStat = "1" Then .SetText 2, 1, "Nurition Status" & vbNewLine
'            .GetText 2, 1, varEMCalc
'            If Trim(varEMCalc)<> vbnullstring Then .SetText 3, 1, "1"
'
'
'       ' Muscle Tone
'       If mvarMuscleStrgth = "1" Then .SetText 2, 2, "Muscle Strength" & vbNewLine
'       If mvarMuscleTone = "1" Then .SetText 2, 2, "Muscle Tone" & vbNewLine
'        If mvarMuscleStrgth = "1" And _
'           mvarMuscleTone = "1" Then
'             .SetText 3, 2, mvarMuscleStrgth
'        End If
'
'         'Gait / Station
'       If mvarGait = "1" Then .SetText 2, 3, "Gait" & vbNewLine
'       If mvarStation = "1" Then .SetText 2, 3, "Station" & vbNewLine
'        If mvarGait = "1" And mvarStation = "1" Then
'            .SetText 3, 3, mvarGait
'        End If
'
'        ' Orientation
'       If mvarOrient = "1" Then .SetText 2, 4, "Orientation" & vbNewLine
'        If mvarOrient = "1" Then
'            .SetText 3, 4, mvarOrient
'        End If
'
'        ' Speech
'       If mvarSpeech = "1" Then .SetText 2, 5, "Speech" & vbNewLine
'       If mvarVoice = "1" Then .SetText 2, 5, "Voice" & vbNewLine
'        If mvarSpeech = "1" And mvarVoice = "1" Then
'            .SetText 3, 5, mvarSpeech
'        End If
'
'       ' Need to Verify
'       If mvarThoughtProc = "1" Then .SetText 2, 6, "Thought Process" & vbNewLine
'       If mvarSimilarities = "1" Then .SetText 2, 6, "Similarities" & vbNewLine
'       If mvarProvInter = "1" Then .SetText 2, 6, "Proverb" & vbNewLine
'       If mvarSerial7 = "1" Then .SetText 2, 6, "Serial 7" & vbNewLine
'        If mvarThoughtProc = "1" And _
'            (mvarProvInter = "1" Or mvarSimilarities = "1") And _
'            mvarSerial7 = "1" Then
'            .SetText 3, 6, mvarThoughtProc
'        End If
'
'       If mvarAssociations = "1" Then .SetText 2, 7, "Associations" & vbNewLine
'       If mvarAssociations = "1" Then
'            .SetText 3, 7, mvarAssociations
'        End If
'
'       If mvarInsight = "1" Then .SetText 2, 8, "Insight" & vbNewLine
'       If mvarJudgement = "1" Then .SetText 2, 8, "Judgement" & vbNewLine
'        If mvarInsight = "1" And _
'            mvarJudgement = "1" Then
'            .SetText 3, 8, mvarInsight
'        End If
'
'       If mvarNaming = "1" Then .SetText 2, 9, "Naming" & vbNewLine
'       If mvarRepitition = "1" Then .SetText 2, 9, "Repitition" & vbNewLine
'       If mvarNaming = 1 Or _
'            mvarRepitition = "1" Then
'            .SetText 3, 9, mvarNaming
'        End If
'
'       If mvarMood = "1" Then .SetText 2, 10, "Mood" & vbNewLine
'       If mvarAffect = "1" Then .SetText 2, 10, "Affect" & vbNewLine
'        If mvarMood And mvarAffect Then
'            .SetText 3, 10, mvarMood
'        End If
'
'       If mvarRegistration = "1" Then .SetText 2, 11, "Registration" & vbNewLine
'       If mvarSerial7 = "1" Then .SetText 2, 11, "Serial 7" & vbNewLine
'       If mvarSpell = "1" Then .SetText 2, 11, "Spell" & vbNewLine
'        If mvarRegistration = "1" And _
'            (mvarSerial7 = "1" Or mvarSpell = "1") Then
'            .SetText 3, 11, mvarRegistration
'        End If
'
'       If mvarHallucination = "1" Then .SetText 2, 12, "Hallucination" & vbNewLine
'       If mvarDelusions = "1" Then .SetText 2, 12, "Delusions" & vbNewLine
'       If mvarSuicide = "1" Then .SetText 2, 12, "Suicide" & vbNewLine
'       If mvarHomicide = "1" Then .SetText 2, 12, "Homicide" & vbNewLine
'        If mvarHallucination = "1" And _
'            mvarDelusions = "1" And _
'            (mvarSuicide = "1") And _
'            (mvarHomicide = "1") Then
'            .SetText 3, 12, mvarSuicide
'        End If
'
'       If mvarFundKnow = "1" Then .SetText 2, 13, "Fund of Knowledge" & vbNewLine
'        'If cboFunKno.ListIndex > -1 Then
'        If mvarFundKnow = "1" Then
'            .SetText 3, 13, mvarFundKnow
'        End If
'
'       If mvarRemMem = "1" Then .SetText 2, 14, "Remote Memory" & vbNewLine
'       If mvarGrossAssess = "1" Then .SetText 2, 14, "Gross Assessment" & vbNewLine
'        If mvarRemMem = "1" And _
'            mvarGrossAssess = "1" Then
'            .SetText 3, 14, mvarRemMem
'        End If
'
'        mvarVitals = Get_ClinicalNotes("FN_PsyMse", "Vitals", mvarVitals, 1, mlngEncounterID)
'
'       If mvarVitals<> vbnullstring Then .SetText 2, 15, "Vital Signs" & vbNewLine
'        If mvarVitals<> vbnullstring Then
'            mvarVitals = "1"
'            .SetText 3, 15, mvarVitals
'        Else
'            mvarVitals = vbnullstring
'        End If
'     End With
'
'    mintMSE_EM = CInt(mvarAppearance) + CInt(mvarMuscleStrgth) + CInt(mvarGait) + CInt(mvarOrient) _
'        + CInt(mvarSpeech) + CInt(mvarThoughtProc) + CInt(mvarAssociations) + CInt(mvarMood) + CInt(mvarInsight) _
'        + CInt(mvarNaming) + CInt(mvarRegistration) + CInt(mvarVitals) + CInt(mvarSuicide) + CInt(mvarFundKnow)
'    With grdSuperBill
'        .Row = 16
'        .Col = 3
'        .Text = mintMSE_EM
'        .SetText 1, 16, "TOTAL MSE HCFA BULLETS"
'
'
'    If mvaVerbalAudi = "1" Then .SetText 2, 18, "Verbal Auditory Content"
'    If mvarAttitude = "1" Then .SetText 2, 19, "Attitude"
'    If mvarConation = "1" Then .SetText 2, 20, "Conation"
'    If mvarAbnorMove = "1" Then .SetText 2, 21, "Abnormal Movement"
'    If mvarIntell = "1" Then .SetText 2, 22, "Intelligence Estimate"
'    If mvarEyeContact = "1" Then .SetText 2, 23, "Eye Contact"
'
'
'    .BlockMode = True
'        .Col = 3
'        .Col2 = 3
'        .Row = .MaxRows
'        .Row2 = .MaxRows
'        .TypeHAlign = TypeHAlignCenter
'        .TypeVAlign = TypeVAlignCenter
'        .FontSize = 12
'        .FontBold = True
'    .BlockMode = False
'    .Redraw = False
'
'    End With
    
    Exit Sub
End Sub


Private Sub Get_SubGroup()
Dim strSQL As String
Dim booRet As Boolean
Dim varRS
Dim lngrecs As Long

    strSQL = "Select  SUBGROUP_ID " _
            & "From INTRVN_SUBGROUPS " _
            & "Where group_id = 2 "

    Load_ListBox lstSubGroup, strSQL, psConnect(1)
    
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
End Sub

Private Sub lstSubGroup_Click()

    Format_Grid
    
    
    Exit Sub
End Sub


