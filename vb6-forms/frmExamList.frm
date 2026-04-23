VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmExamList 
   BackColor       =   &H80000004&
   Caption         =   "Physical Exam Details"
   ClientHeight    =   5760
   ClientLeft      =   9135
   ClientTop       =   4170
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExamList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   12105
   Begin VB.Frame fraComment 
      Height          =   3135
      Left            =   4860
      TabIndex        =   14
      Top             =   510
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox cctComments 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   150
         Width           =   6360
      End
      Begin PT_XP_Button.PT_XP cmdCommentClose 
         Height          =   345
         Left            =   3240
         TabIndex        =   16
         Top             =   2670
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
         Picture         =   "frmExamList.frx":08CA
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmExamList.frx":12DC
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCommentAdd 
         Height          =   345
         Left            =   2190
         TabIndex        =   17
         Top             =   2670
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
         Picture         =   "frmExamList.frx":162E
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Done"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmExamList.frx":1BC8
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.Frame fraSpecLocation 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3375
      Left            =   2400
      TabIndex        =   8
      Top             =   1020
      Visible         =   0   'False
      Width           =   4695
      Begin FPSpreadADO.fpSpread grdSpecLoc 
         Height          =   2775
         Left            =   60
         TabIndex        =   9
         Top             =   150
         Width           =   4545
         _Version        =   458752
         _ExtentX        =   8017
         _ExtentY        =   4895
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
         SpreadDesigner  =   "frmExamList.frx":1F1A
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin PT_XP_Button.PT_XP cmdLocate 
         Height          =   285
         Left            =   1620
         TabIndex        =   10
         Top             =   3000
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
         Picture         =   "frmExamList.frx":21DA
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Select"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmExamList.frx":2BEC
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSpecLocTerm 
         Height          =   255
         Left            =   4140
         TabIndex        =   11
         ToolTipText     =   "Terms Maintanence Units"
         Top             =   3000
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
         Picture         =   "frmExamList.frx":2F3E
         AppearanceThemes=   3
         BorderWidth     =   2
         TransparentColor=   15790320
         Caption         =   ".."
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmExamList.frx":34D8
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.Frame fraGridTerm 
      BackColor       =   &H80000004&
      Height          =   4545
      Left            =   150
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   11880
      Begin PT_XP_Button.PT_XP cmdTermSave 
         Height          =   315
         Left            =   5633
         TabIndex        =   5
         Top             =   4080
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
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
         Picture         =   "frmExamList.frx":382A
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Done"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmExamList.frx":3BC4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRemoveRow 
         Height          =   315
         Left            =   1020
         TabIndex        =   6
         Top             =   -450
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
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
         Picture         =   "frmExamList.frx":3F16
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Remove"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmExamList.frx":44B0
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdChildTerm 
         Height          =   375
         Left            =   10500
         TabIndex        =   7
         ToolTipText     =   "Add Terms to Displayed List Box"
         Top             =   3990
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
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
         Picture         =   "frmExamList.frx":4802
         AppearanceThemes=   3
         BorderDrawEvent =   1
         TransparentColor=   15790320
         Caption         =   "Terms"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmExamList.frx":4F7C
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdTerms 
         Height          =   3945
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   11790
         _Version        =   458752
         _ExtentX        =   20796
         _ExtentY        =   6959
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
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   6
         MaxRows         =   0
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   4194304
         ShadowText      =   8388608
         SpreadDesigner  =   "frmExamList.frx":568E
         UserResize      =   0
         Appearance      =   2
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
         CellNoteIndicator=   2
      End
   End
   Begin FPSpreadADO.fpSpread grdGrid 
      Height          =   4485
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   11775
      _Version        =   458752
      _ExtentX        =   20770
      _ExtentY        =   7911
      _StockProps     =   64
      AutoCalc        =   0   'False
      DAutoSizeCols   =   0
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
      GrayAreaBackColor=   16777215
      GridColor       =   14737632
      MaxCols         =   6
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarMaxAlign=   0   'False
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmExamList.frx":594E
      UserResize      =   0
      Appearance      =   2
      ScrollBarTrack  =   1
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdTerms 
      Height          =   375
      Left            =   10830
      TabIndex        =   1
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      Picture         =   "frmExamList.frx":5C06
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Terms"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmExamList.frx":61A0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdAddList 
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   -390
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
      Picture         =   "frmExamList.frx":64F2
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmExamList.frx":6A8C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRemove 
      Height          =   345
      Left            =   1110
      TabIndex        =   3
      Top             =   -390
      Width           =   1185
      _ExtentX        =   2090
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
      Picture         =   "frmExamList.frx":6DDE
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Remove"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmExamList.frx":7378
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   405
      Left            =   5505
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
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
      Picture         =   "frmExamList.frx":76CA
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Close"
      DepthEvent      =   1
      PictureDisabled =   "frmExamList.frx":80DC
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmExamList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngtop As Long
Public mlngleft As Long
Public mstrTerms As String
Private mbooByPass As Boolean
Public mFirst As Boolean
Public mbooCaption As Boolean
Public mlngPatientID As Long
Public mlngEncounterID As Long
Private mlngGridRow As Long
Private mlngGridCol As Long
Private GrdLock As Boolean
Private mlngTermRow As Long
Private mlngTermCol As Long
Private flngDeltaX  As Long
Private flngDeltaY  As Long
Private mGrid As String
Private udtPhyExamDetails As udtPhyExamDetails
Private Type udtPhyExamDetails
    Child_Area As String
    Normal As String
    Abnormal As String
    GridRow As Long
    GridCol As Long
End Type
Private Function Clear_List_Details()
    With udtPhyExamDetails
        .Child_Area = vbNullString
        .Normal = vbNullString
        .Abnormal = vbNullString
        .GridRow = 0
        .GridCol = 0
    End With
End Function

Private Function Format_GridTerms()
Dim varColName
Dim lngRow As Long
Dim strList As String
Dim lngX As Long
Dim lngY As Long
Dim lngWidth As Long
Dim lngHeight As Long

    Screen.MousePointer = vbHourglass
    Clear_GridFormat
    With grdTerms
        .ReDraw = False
        .ClearRange 0, SpreadHeader, grdTerms.MaxCols, grdTerms.MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .MaxRows = 0
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .RowHeadersShow = False
        .MaxCols = 3
        grdTerms.ColWidth(1) = 30
        grdTerms.ColWidth(2) = 32
        grdTerms.ColWidth(3) = 32
        Format_CellTypeEdit grdTerms, SpreadHeader, 1, "3", UCase(udtPhyExam.ExamArea), "TypeHAlignCenter", "EXAM SUB_GROUP", "", HeaderColor
        grdTerms.SetCellBorder 1, SpreadHeader, 3, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeEdit grdTerms, SpreadHeader + 1, 1, "1", "AREA", "TypeHAlignCenter", "AREA", "", HeaderColor
        grdTerms.SetCellBorder 1, SpreadHeader + 1, 1, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeEdit grdTerms, SpreadHeader + 1, 2, "1", "FINDINGS", "TypeHAlignCenter", "DETAILS", ", HeaderColor"
        grdTerms.SetCellBorder 2, SpreadHeader + 1, 2, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeEdit grdTerms, SpreadHeader + 1, 3, "1", "COMMENTS", "TypeHAlignCenter", "COMMENTS", "Enter Comments", HeaderColor
        grdTerms.SetCellBorder 3, SpreadHeader + 1, 3, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer grdTerms, SpreadHeader + 2
        cmdClose.Visible = False
        fraGridTerm.ZOrder 0
        fraGridTerm.Visible = True
        grdGrid.Visible = False
    End With
    Screen.MousePointer = vbDefault
    Exit Function
End Function
Private Function Get_GridTerms(vstrSubDom As String)
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim lngRow As Long

    On Error GoTo Error_Get_Terms
    mstrTerms = vstrSubDom
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Get_Attr_Valid_Value_SubDom_Sequence  '" & vstrSubDom & "'"
    Set rsTerms = cnTerms.Execute(strSQL)
    With rsTerms
        Do Until .EOF
            If (!ATTRBTE_VALID_VALUE) <> vbNullString Then
                lngRow = lngRow + 1
                grdTerms.MaxRows = lngRow
                Format_CellTypeEdit grdTerms, lngRow, 1, "1", Trim(!ATTRBTE_VALID_VALUE), "EXAM SUB_GROUP", "TypeHAlignLeft", udtPhyExam.ExamArea
                Format_CellTypeStatic grdTerms, lngRow, 2, "1", "", "TypeHAlignLeft", "DETAILS", "Click to Select Terms"
                Format_CellTypeEdit grdTerms, lngRow, 3, "1", " ", "TypeHAlignLeft", "COMMENTS", "Enter Comments"
                grdTerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdTerms.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
            .MoveNext
        Loop
    End With
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set rsTerms = Nothing
        Set cnTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Terms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_GridTerms", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function

Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional vlngBackColor As Long, Optional vbooUnderLine As Boolean)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeEdit
            .TypeMaxEditLen = 32000
            .TypeEditMultiLine = True   'Make multi-lines
            .FontBold = True
            If vbooUnderLine Then .FontUnderline = True
            If vlngBackColor Then .BackColor = vlngBackColor
            If Trim(vstrCellNote) <> vbNullString Then .CellNote = vstrCellNote
            If vstrSpan <> vbNullString Then
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                    .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                    lngSpan = Val(vstrSpan)
                    If lngSpan > 0 Then
                        .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                    End If
                End If
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote + 7
            .TypeVAlign = TypeVAlignCenter
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            .EditMode = False
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional lngBackColor As Long)
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
            .FontSize = 8
            .FontBold = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
'            If vstrVAlign = "TypeVAlignCenter" Then
'                .TypeVAlign = TypeVAlignCenter
'            ElseIf vstrVAlign = "TypeVAlignBottom" Then
'                .TypeVAlign = TypeVAlignBottom
'            Else
'               .TypeVAlign = TypeVAlignTop
'            End If
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
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
            If lngBackColor > 0 Then .BackColor = lngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        .BlockMode = False
    End With
End Function
Private Function Format_CellTypeCombo(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, _
vstrString As String, vstrHAligh As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
              .Col = vlngCol
              .Col2 = vlngCol
              .Row = vlngRow
              .Row2 = vlngRow
              .CellType = CellTypeComboBox
              .TypeMaxEditLen = 32000
              .TypeComboBoxList = vstrString
              .TypeComboBoxWidth = 1
              If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
              If vstrCellTag <> vbNullString Then
                .CellTag = vstrCellTag
              End If
              If vstrSpan <> vbNullString Then
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                    .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                Else
                    lngSpan = Val(vstrSpan)
                    If lngSpan > 0 Then
                        .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                    End If
                End If
              End If
              If vlngBackColor > 0 Then
                 .BackColor = vlngBackColor
              End If
              .TypeHAlign = TypeHAlignCenter
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
          .BlockMode = False
    End With
End Function
Private Sub Add_Spacer(vObjGrid As Object, Optional vlngRow As Long)
Dim iRow As Long
    With vObjGrid
        If vlngRow >= 0 Then
             vlngRow = vlngRow + 1
            .MaxRows = vlngRow
            iRow = vlngRow
        Else
            iRow = vlngRow
        End If
        .BlockMode = True
            .Col = 1
            .Col2 = 1
            .Row = vlngRow
            .Row2 = vlngRow
            .RowHeight(vlngRow) = 3
            .BackColor = GridLine
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 1, vlngRow, .MaxCols, vlngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub

Private Function Format_SpecificGrid(vstrSubDom As String, vstrTag As String)
Dim varColName
Dim strList As String
Dim lngX As Long
Dim lngY As Long
Dim lngWidth As Long
Dim lngHeight As Long
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim lngRow As Long

    On Error GoTo Error_Format_SpecificGrid
    Screen.MousePointer = vbHourglass
    clsGridFormat.CellTag = vstrTag
    mstrTerms = vstrSubDom
    With grdSpecLoc
        .ReDraw = False
        .ClearRange 0, SpreadHeader, grdSpecLoc.MaxCols, grdSpecLoc.MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .ColHeaderRows = 2
        .ColHeadersShow = True
        .RowHeadersShow = False
        .MaxCols = 1
        grdSpecLoc.MaxRows = 0
        .ColWidth(1) = 35
        Format_CellTypeEdit grdSpecLoc, SpreadHeader, 1, "1", "SELECT ALL THAT APPLIES", "TypeHAlignCenter", vstrTag, "", HeaderColor
        grdSpecLoc.SetCellBorder 1, SpreadHeader, 1, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer grdSpecLoc, SpreadHeader + 1
        .ReDraw = True
    End With
    If vstrSubDom <> vbNullString Then
        Open_DB psConnect(1)
        strSQL = "Get_Attr_Valid_Value_Like_SubDom '" & vstrSubDom & "%'"
        Set rs = cn.Execute(strSQL)
        With rs
            Do Until .EOF
                lngRow = lngRow + 1
                If grdSpecLoc.MaxRows < lngRow Then grdSpecLoc.MaxRows = lngRow
                Format_CellTypeEdit grdSpecLoc, lngRow, 1, "1", (!ATTRBTE_VALID_VALUE), "TypeHAlignLeft", vstrTag
                grdSpecLoc.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
           .MoveNext
           Loop
        End With
    End If
    If mGrid = "grdGrid" Then
        With grdGrid
'            .Col = mlngGridCol
'            .Row = mlngGridRow
'            .GetCellPos .Col, .Row, lngX, lngY, lngWidth, lngHeight
'            fraSpecLocation.Top = lngY + lngHeight
'            fraSpecLocation.Left = grdGrid.Left + lngX
            fraSpecLocation.ZOrder 0
            fraSpecLocation.Visible = True
        End With
    Else
        With grdTerms
'            .Col = mlngTermCol
'            .Row = mlngTermRow
'            .GetCellPos .Col, .Row, lngX, lngY, lngWidth, lngHeight
'            fraSpecLocation.Top = lngY + lngHeight
'            fraSpecLocation.Left = grdGrid.Left + lngX
            fraSpecLocation.ZOrder 0
            fraSpecLocation.Visible = True
        End With
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_SpecificGrid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_SpecificGrid", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function

Private Sub Format_Headers()
Dim colID As Variant
Dim iCnt As Integer
Dim lngCol As Integer
Dim lngRow As Integer
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim varHeader
Dim strCount As String
Dim strSQL As String
Dim varSysName
Dim varField
    
    Screen.MousePointer = vbHourglass
    grdGrid.ClearRange 0, 1, grdGrid.MaxCols, grdGrid.MaxRows, False
    With grdGrid
        .ReDraw = False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 10
        .FontName = "Arial"
        .GridSolid = False
        .GridShowHoriz = False
        .GridShowVert = False
        .RowHeadersShow = False
        .ColHeadersShow = True
        .ColHeaderRows = 3
        .ColHeadersAutoText = DispBlank
        .MaxCols = 3
        .MaxRows = 0
    End With
    grdGrid.ColWidth(1) = 35
    grdGrid.ColWidth(2) = 27
    grdGrid.ColWidth(3) = 32
    Format_CellTypeEdit grdGrid, SpreadHeader, 1, "3", UCase(udtPhyExam.ExamSystem), "TypeHAlignCenter", "HEADER", "", HeaderColor
    grdGrid.SetCellBorder 1, SpreadHeader, 3, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Format_CellTypeEdit grdGrid, SpreadHeader + 1, 1, "1", "EXAMINATION AREA", "TypeHAlignCenter", "HEADER", "", HeaderColor
    grdGrid.SetCellBorder 1, SpreadHeader + 1, 2, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Format_CellTypeEdit grdGrid, SpreadHeader + 1, 2, "1", "FINDINGS DETAILS", "TypeHAlignCenter", "HEADER", "", HeaderColor
    grdGrid.SetCellBorder 3, SpreadHeader + 1, 4, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Format_CellTypeEdit grdGrid, SpreadHeader + 1, 3, "1", "COMMENTS", "TypeHAlignCenter", "HEADER", "", HeaderColor
    grdGrid.SetCellBorder 5, SpreadHeader + 1, 6, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer grdGrid, SpreadHeader + 2
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Function Read_Data()
Dim cnExam As New ADODB.Connection
Dim rsExam As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim varCellValue
Dim lngExamID As Long
Dim varGroup
Dim strGroup As String
    On Error GoTo Error_Read_Data
    Screen.MousePointer = vbHourglass
    If cnExam.State = adStateClosed Then cnExam.open psConnect(0)
    With grdGrid
        For lngRow = 1 To .MaxRows
            udtPhyExam.ExamArea = vbNullString
            .GetText 1, lngRow, varCellValue
            If varCellValue <> vbNullString Then
                udtPhyExam.ExamArea = Trim(varCellValue)
            End If
            udtPhyExam.Comment = vbNullString
            udtPhyExam.Finding = vbNullString
            If cnExam.State = adStateClosed Then cnExam.open psConnect(0)
            strSQL = "SELECT FINDING, COMMENTS, SEQ_NO FROM PATIENT_PHYSICAL_DETAILS WHERE PATIENT_ID = " & piPatientID & " AND " _
               & "ENCOUNTER_ID = " & piEncounterID & " AND PARENT_AREA = '" & udtPhyExam.ExamSystem & "' AND " _
               & "PARENT_SPECIFIC_AREA like '" & udtPhyExam.ExamArea & "'"
            Set rsExam = cnExam.Execute(strSQL)
            Do Until rsExam.EOF
                If Trim(rsExam.Fields("FINDING")) <> vbNullString Then
                    If udtPhyExam.Finding = vbNullString Then
                        udtPhyExam.Finding = Trim(rsExam.Fields("FINDING"))
                    Else
                        udtPhyExam.Finding = udtPhyExam.Finding & vbNewLine & Trim(rsExam.Fields("FINDING"))
                    End If
                End If
                If Trim(rsExam.Fields("COMMENTS")) <> vbNullString Then
                    If udtPhyExam.Comment = vbNullString Then
                        udtPhyExam.Comment = Trim(rsExam.Fields("COMMENTS"))
                    Else
                        udtPhyExam.Comment = udtPhyExam.Comment & vbNewLine & Trim(rsExam.Fields("COMMENTS"))
                    End If
                End If
                rsExam.MoveNext
            Loop
            If udtPhyExam.Finding = udtPhyExam.Comment Then udtPhyExam.Finding = vbNullString
            Format_CellTypeEdit grdGrid, lngRow, 2, "1", udtPhyExam.Finding, "TypeHAlignLeft", "DETAILS"
            Format_CellTypeEdit grdGrid, lngRow, 3, "1", udtPhyExam.Comment, "TypeHAlignLeft", "COMMENTS"
            grdGrid.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdGrid.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            udtPhyExam.Comment = vbNullString
            udtPhyExam.Finding = vbNullString
        Next lngRow
    End With
    udtPhyExam.ExamArea = vbNullString
     If cnExam.State = adStateOpen Then
        cnExam.Close
        Set cnExam = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Data", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Write_ChildGrid(vstrData As String, vstrExamArea As String)
Dim lngPrimKey As Long
Dim cnExam As New ADODB.Connection
Dim rsExam As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim booinsert As Boolean
Dim sDate As String
Dim strEnctrDate As String

    On Error GoTo Error_Write_ChildGrid
    Screen.MousePointer = vbHourglass
    If strEnctrDate = vbNullString Then
       strEnctrDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    Else
        strEnctrDate = Format(strEnctrDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    End If
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If cnExam.State = adStateClosed Then cnExam.open psConnect(0)
    strSQL = "Get_Patient_Physical_Exam_PT_Enctr " & piPatientID & ", " _
        & piEncounterID & ",'" & udtPhyExam.ExamSystem & "','" & udtPhyExam.ExamArea & "','" & udtPhyExam.ExamChild & "'"
        Set rsExam = cnExam.Execute(strSQL)
        With rsExam
            If Not .EOF Then
            lngPrimKey = (!EXAM_DET_ID)
            End If
        End With
    If lngPrimKey = 0 Then
       strSQL = "Write_Physical_Exam_Details " & piPatientID & "," & piEncounterID & ",'" _
           & udtPhyExam.ExamChild & "','" & udtPhyExam.ExamSystem & "','" & udtPhyExam.ExamArea & "','" _
           & vstrData & "','" & strEnctrDate & "')"
           Set rsExam = cnExam.Execute(strSQL)
    Else
        If vstrData <> vbNullString Then
            strSQL = "Update_Physical_Exam_Details_Finding  '" & Trim(vstrData) & "','" & sDate & "'," & lngPrimKey
            Set rsExam = cnExam.Execute(strSQL)
        End If
        If udtPhyExam.Comment <> vbNullString Then
            strSQL = "Update_Physical_Exam_Details_Comments '" & udtPhyExam.Comment & "', '" & sDate & "'," & lngPrimKey
            Set rsExam = cnExam.Execute(strSQL)
        End If
    End If
     If cnExam.State = adStateOpen Then
        cnExam.Close
        Set cnExam = Nothing
        Set rsExam = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_ChildGrid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_ChildGrid", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Write_Data(vstrData As String, vstrExamArea As String)
Dim lngPrimKey As Long
Dim cnExam As New ADODB.Connection
Dim rsExam As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim booinsert As Boolean
Dim sDate As String
Dim strEnctrDate As String

    On Error GoTo Error_Write_Data
    Screen.MousePointer = vbHourglass
    strEnctrDate = Get_Enctr_Date(mlngEncounterID)
    If strEnctrDate = vbNullString Then
       strEnctrDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    Else
        strEnctrDate = Format(strEnctrDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    End If
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If cnExam.State = adStateClosed Then cnExam.open psConnect(0)
    strSQL = "Get_Patient_Physical_Exam_PT_Enctr_NoChild " & piPatientID & "," _
        & piEncounterID & ",'" & udtPhyExam.ExamSystem & "','" & udtPhyExam.ExamArea & "'"
        Set rsExam = cnExam.Execute(strSQL)
            With rsExam
                If Not .EOF Then
                lngPrimKey = (!EXAM_DET_ID)
                End If
            End With
            If lngPrimKey = 0 Then
               strSQL = "Write_Physical_Exam_Details_No_Child " & piPatientID & "," & piEncounterID & ",'" _
                   & udtPhyExam.ExamSystem & "','" & vstrExamArea & "','" _
                   & udtPhyExam.Finding & "','" & udtPhyExam.Comment & "','" & strEnctrDate & "'"
                   Set rsExam = cnExam.Execute(strSQL)
            Else
                strSQL = "Update_Physical_Exam_Details_Finding '" & udtPhyExam.Finding & "','" & sDate & "'," & lngPrimKey
                Set rsExam = cnExam.Execute(strSQL)
                strSQL = "Update_Physical_Exam_Details_Comments  '" & udtPhyExam.Comment & "', '" & sDate & "'," & lngPrimKey
                Set rsExam = cnExam.Execute(strSQL)
            End If
     If cnExam.State = adStateOpen Then
        cnExam.Close
        Set cnExam = Nothing
        Set rsExam = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Data", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Set_GridCurSel(vObjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
Dim Idx As Integer
Dim booData As Boolean
Dim strSpan As String
Dim strTag As String
Dim strCellNote As String
Dim lngCol As Long
Dim lngSpan As Long
    With vObjGrid
        .ReDraw = False
        .Col = vlngCol
        .Row = vlngRow
        If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        If .CellType = CellTypeComboBox Then
            For Idx = 0 To .TypeComboBoxCount
            .TypeComboBoxIndex = Idx
            If TextFound(Trim(vstrData), Trim(.TypeComboBoxString), Len(Trim(vstrData)), True) Then
                .TypeComboBoxCurSel = Idx
                Exit For
            End If
            Next Idx
        End If
        If vstrSpan <> vbNullString Then
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                lngSpan = Val(vstrSpan)
                If lngSpan > 0 Then
                    .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                End If
            End If
        End If
        If vlngBackColor > 0 Then .BackColor = vlngBackColor
        If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
    .ReDraw = True
    End With
End Function

Private Sub cmdAddList_Click()
Dim i%
Dim lngRow As Long

    With grdGrid
        lngRow = .MaxRows + 1
        .MaxRows = lngRow
        Format_CellTypeEdit grdGrid, lngRow, 1, "1", "", "TypeHAlignCenter", "EXAM_SYSTEM"
        grdGrid.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeEdit grdGrid, lngRow, 2, "1", "", "TypeHAlignCenter", "FINDINGS"
        grdGrid.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeEdit grdGrid, lngRow, 3, "1", "", "TypeHAlignCenter", "COMMENTS"
        grdGrid.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    End With
    Exit Sub
End Sub
Private Sub cmdChildTerm_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = udtPhyExam.ExamSystem
    frmNurseTemplates.fType = udtPhyExam.ExamSystem
    frmNurseTemplates.fSubDomain = mstrTerms
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    Format_GridTerms
    Get_GridTerms mstrTerms
    Exit Sub
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdCommentAdd_Click()
Dim lngRow As Long
Dim lngCol As Long
Dim strNote As String
Dim varCellValue

    udtPhyExam.Comment = Trim(cctComments)
    With grdGrid
        lngCol = 3
        lngRow = .ActiveRow
        If grdGrid.CellNote <> vbNullString Then strNote = grdGrid.CellNote
        .GetText 1, lngRow, varCellValue
        udtPhyExam.ExamArea = Trim(varCellValue)
    End With
    
    lngCol = 3
    Format_CellTypeEdit grdGrid, lngRow, lngCol, "1", Trim(udtPhyExam.Comment), "TypeHAlignCenter", "COMMENTS", , grdCurItem
    grdGrid.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Write_Data Trim(udtPhyExam.Comment), Trim(udtPhyExam.ExamArea)
    fraComment.Visible = False
    Exit Sub
End Sub

Private Sub cmdCommentClose_Click()

    fraComment.Visible = False
End Sub


Private Sub cmdLocate_Click()
Dim strData As String
Dim lngRow As Long
Dim varCellValue
Dim strTag As String
Dim Arr
Dim lngCol As Long
Dim strCollection As String
Dim cnExam As New ADODB.Connection
Dim rsExam As New ADODB.Recordset
Dim strSQL As String
Dim lngPrimKey As Long
Dim strDate As String

    Screen.MousePointer = vbDefault
    On Error GoTo Error_cmdLocate
    With grdSpecLoc
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                .GetText 1, lngRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    If Trim(strCollection) = vbNullString Then
                        strCollection = Trim(varCellValue)
                    ElseIf Not TextFound(CStr(varCellValue), strCollection, Len(varCellValue), True) Then
                        strCollection = strCollection & "," & Trim(varCellValue)
                    End If
                End If
            End If
        Next lngRow
        
        If Trim(strCollection) <> vbNullString Then
             If mGrid = "grdGrid" Then
                grdGrid.Col = mlngGridCol
                grdGrid.Row = mlngGridRow
                If .CellTag <> vbNullString Then strTag = .CellTag
                grdGrid.GetText 3, mlngGridRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then udtPhyExam.Comment = Trim(varCellValue)
                Format_CellTypeEdit grdGrid, mlngGridRow, 3, "1", Trim(strCollection), "TypeHAlignLeft", Trim(udtPhyExam.ExamArea), "", DataColor
                grdGrid.SetCellBorder 3, mlngGridRow, 3, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Write_Data Trim(strCollection), Trim(udtPhyExam.ExamArea)
             Else
                grdTerms.Col = mlngTermCol
                grdTerms.Row = mlngTermRow
                If .CellTag <> vbNullString Then strTag = .CellTag
                grdTerms.GetText 1, mlngTermRow, varCellValue
                varCellValue = Replace(varCellValue, "_", "")
                If Trim(varCellValue) <> vbNullString Then strCollection = Trim(varCellValue) & ": " & strCollection
                grdTerms.GetText 2, mlngTermRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then udtPhyExam.Finding = Trim(varCellValue)
                grdTerms.GetText 3, mlngTermRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then udtPhyExam.Comment = Trim(varCellValue)
                Format_CellTypeEdit grdTerms, mlngTermRow, 3, "1", Trim(strCollection), "TypeHAlignLeft", Trim(udtPhyExam.ExamArea), "", DataColor
                grdTerms.SetCellBorder 3, mlngTermRow, 3, mlngTermRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Write_ChildGrid Trim(strCollection), Trim(udtPhyExam.ExamArea)
            End If
        End If
    End With
    fraSpecLocation.Visible = False
    GrdLock = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdLocate:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdLocate", App.EXEName & "." & TypeName(Me)
    Exit Sub
    Resume
End Sub

Private Sub cmdRemoveRow_Click()
'    If mlngGridRow > 0 Then
'        grdTerms.DeleteRows (mlngGridRow), 1
'        grdTerms.MaxRows = grdTerms.MaxRows - 1
'    End If
End Sub


Private Sub cmdSpecLocTerm_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = UCase(udtPhyExam.ExamSystem) & " " & UCase(udtPhyExam.ExamArea)
    frmNurseTemplates.fType = "TERMS MAINTANENCE"
    frmNurseTemplates.fSubDomain = mstrTerms
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    mGrid = "grdTerms"
    Format_SpecificGrid mstrTerms, udtPhyExam.ExamArea
    Exit Sub
End Sub

Private Sub cmdTerms_Click()
        mstrTerms = "sd_PhyExam_" & Trim(udtPhyExam.ExamSystem)
        frmNurseTemplates.fSubDomain = mstrTerms
        frmNurseTemplates.fConnectTM = psConnect(1)
        Load frmNurseTemplates
        frmNurseTemplates.Show vbModal
        AddList_Terms mstrTerms
    Exit Sub
End Sub

Private Sub cmdTermSave_Click()
Dim lngPrimKey As Long
Dim cnExam As New ADODB.Connection
Dim rsExam As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim strParent As String
Dim varCellValue
Dim sDate As String
Dim lngSeqNo As Long
Dim strChild  As String
Dim strParArea As String
Dim strCollection As String

    On Error GoTo Error_TermSave
    Screen.MousePointer = vbHourglass
    If udtPhyExam.ExamSystem <> vbNullString Then strParent = udtPhyExam.ExamSystem
    Clear_PhysExam
    udtPhyExam.ExamSystem = strParent
    udtPhyExam.ExamArea = strParArea
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    Get_Encounter mlngEncounterID
    If cnExam.State = adStateClosed Then cnExam.open psConnect(0)
    With grdTerms
        For lngRow = 1 To .MaxRows
            lngPrimKey = 0
            strCollection = vbNullString
            udtPhyExam.Finding = vbNullString
            udtPhyExam.Comment = vbNullString
            lngPrimKey = .GetRowItemData(lngRow)
            .Col = 1
            .Row = lngRow
            If .CellTag <> vbNullString Then udtPhyExam.ExamChild = .CellTag
            .GetText 1, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                udtPhyExam.ExamArea = Trim(varCellValue)
                .GetText 3, lngRow, varCellValue
                udtPhyExam.Finding = Trim(varCellValue)
                .GetText 5, lngRow, varCellValue
                udtPhyExam.Comment = Trim(varCellValue)
                If Trim(udtPhyExam.Comment) <> vbNullString Then
                    If lngPrimKey <= 0 Then
                         strSQL = "SELECT EXAM_DET_ID FROM PATIENT_PHYSICAL_DETAILS WHERE PATIENT_ID = " & piPatientID & " AND " _
                            & "ENCOUNTER_ID = " & piEncounterID & " AND PARENT_AREA = '" & udtPhyExam.ExamSystem & "' AND " _
                            & "PARENT_SPECIFIC_AREA = '" & udtPhyExam.ExamArea & "' AND CHILD_AREA ='" & udtPhyExam.ExamChild & "'"
                        Set rsExam = cnExam.Execute(strSQL)
                        If Not rsExam.EOF Then
                            If rsExam.Fields("EXAM_DET_ID") > 0 Then
                                lngPrimKey = rsExam.Fields("EXAM_DET_ID")
                            End If
                        End If
                    End If
                    If lngPrimKey = 0 Then
                         strSQL = "INSERT INTO PATIENT_PHYSICAL_DETAILS (PATIENT_ID, ENCOUNTER_ID, CHILD_AREA, " _
                            & "PARENT_AREA,PARENT_SPECIFIC_AREA, FINDING, COMMENTS,INSERT_DATE,SEQ_NO) " _
                            & "values(" & piPatientID & "," & piEncounterID & ",'" _
                            & udtPhyExam.ExamChild & "','" & udtPhyExam.ExamSystem & "','" & udtPhyExam.ExamArea & "','" _
                            & udtPhyExam.Finding & "','" & udtPhyExam.Comment & "','" & clsEncounter.StartDate & "'," & lngSeqNo & ")"
                            Set rsExam = cnExam.Execute(strSQL)
                     ElseIf udtPhyExam.Comment <> vbNullString Then
                        strSQL = "UPDATE  PATIENT_PHYSICAL_DETAILS SET  COMMENTS = '" & udtPhyExam.Comment & "'," _
                            & "UPDATE_DATE = '" & sDate & "' WHERE EXAM_DET_ID = " & lngPrimKey
                            Set rsExam = cnExam.Execute(strSQL)
                     End If
                End If
            End If
        Next lngRow
    End With
     If cnExam.State = adStateOpen Then
        cnExam.Close
        Set cnExam = Nothing
        Set rsExam = Nothing
    End If
    cmdClose.Visible = True
    grdGrid.Visible = True
    fraGridTerm.Visible = False
    Read_Data
    Screen.MousePointer = vbDefault
    Exit Sub
Error_TermSave:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "TermSave", App.EXEName & "." & TypeName(Me)
    Exit Sub
    Resume
End Sub

Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
        cmdTerms.Visible = True
        Format_Headers
        AddList_Terms mstrTerms
        Read_Data
        centerForm Me
    End If
    Exit Sub
End Sub
Public Sub AddList_Terms(vstrTerms As String)
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim lngRow As Long

    On Error GoTo Error_AddList_Terms
    Screen.MousePointer = vbHourglass
    If vstrTerms = vbNullString Then Exit Sub
    If mstrTerms <> vbNullString Then
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Get_Attr_Valid_Value_SubDom_Sequence   '" & vstrTerms & "'"
        Set rsTerms = cnTerms.Execute(strSQL)
        With rsTerms
            Do Until .EOF
                clsGridFormat.CellNote = vbNullString
                If Trim((!ATTRBTE_VALID_VALUE)) <> vbNullString Then
                    clsGridFormat.CellNote = Trim((!ATTRBTE_VALID_VALUE))
                End If
           .MoveNext
           If clsGridFormat.CellNote <> vbNullString Then
                lngRow = lngRow + 1
                If grdGrid.MaxRows < lngRow Then grdGrid.MaxRows = lngRow
                Format_CellTypeEdit grdGrid, lngRow, 1, "1", clsGridFormat.CellNote, "TypeHAlignLeft", udtPhyExam.ExamSystem, clsGridFormat.CellNote
                Format_CellTypeStatic grdGrid, lngRow, 2, "1", " ", "TypeHAlignLeft", "DETAILS", clsGridFormat.CellNote
                Format_CellTypeEdit grdGrid, lngRow, 3, "1", " ", "TypeHAlignCenter", "COMMENTS", clsGridFormat.CellNote
                grdGrid.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdGrid.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdGrid.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
           Loop
        End With
    End If
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set rsTerms = Nothing
        Set cnTerms = Nothing
    End If
    clsGridFormat.CellNote = vbNullString
    Read_Data
    Screen.MousePointer = vbDefault
    Exit Sub
Error_AddList_Terms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "AddList_Terms", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub Form_Load()
    mFirst = True
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
    grdGrid.ClearRange 0, 1, grdGrid.MaxCols, grdGrid.MaxRows, False
    grdSpecLoc.ClearRange 0, SpreadHeader, grdSpecLoc.MaxCols, grdSpecLoc.MaxRows, False
    grdTerms.ClearRange 0, SpreadHeader, grdTerms.MaxCols, grdTerms.MaxRows, False
    Unload Me
    Exit Sub
End Sub

Private Sub fraSpecLocation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    flngDeltaX = X
    flngDeltaY = Y
End Sub

Private Sub fraSpecLocation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        fraSpecLocation.Left = fraSpecLocation.Left + X - flngDeltaX
        fraSpecLocation.Top = fraSpecLocation.Top + Y - flngDeltaY
    End If
End Sub


Private Sub grdGrid_Change(ByVal Col As Long, ByVal Row As Long)
Dim strTag As String
Dim strNote As String
Dim varCellValue

    grdGrid.Col = Col
    grdGrid.Row = Row
    If grdGrid.CellTag <> vbNullString Then strTag = grdGrid.CellTag
    If grdGrid.CellNote <> vbNullString Then strNote = grdGrid.CellNote
    If Col = 2 Then
        grdGrid.GetText Col, Row, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtPhyExam.Comment = varCellValue
            Format_CellTypeEdit grdGrid, Row, Col, "1", Trim(udtPhyExam.Finding), "TypeHAlignCenter", strTag, strNote, grdCurItem
            grdGrid.SetCellBorder Col, Row, Col, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
    ElseIf Col = 3 Then
        grdGrid.GetText Col, Row, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtPhyExam.Comment = Trim(varCellValue)
            Format_CellTypeEdit grdGrid, Row, Col, "1", Trim(udtPhyExam.Comment), "TypeHAlignCenter", strTag, strNote, grdCurItem
            grdGrid.SetCellBorder Col, Row, Col, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Write_Data Trim(varCellValue), Trim(udtPhyExam.ExamArea)
        End If
    End If
    Exit Sub
End Sub

Private Sub grdGrid_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim varHeader
Dim CellNote As Long
Dim strTerms As String
Dim lngRow As Long
Dim strData As String
Dim varCellValue

    mlngGridRow = Row
    mlngGridCol = Col
    If Row <= 0 Then Exit Sub
    mGrid = "grdTerms"
    With grdGrid
'        For lngRow = 1 To .MaxRows
'            .BlockMode = True
'                .Col = 1
'                .Col2 = .MaxCols
'                .Row = lngRow
'                .Row2 = lngRow
'                .BackColor = NormalColor
'            .BlockMode = False
'        Next lngRow
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = grdCurItem
        .BlockMode = False
        .Col = mlngGridCol
        .Row = mlngGridRow
        If .CellTag = "DETAILS" Then
            .GetText 1, mlngGridRow, varHeader
             If Trim(varHeader) <> vbNullString Then udtPhyExam.ExamArea = varHeader
             If TextFound("Gait", udtPhyExam.ExamArea, Len("Gait"), True) Then
                 frmGait.mstrConnectStringTerms = psConnect(1)
                 frmGait.mstrConnectStringClinical = psConnect(0)
                 frmGait.mstrTerms = "sd_PysMse_Gait"
                 clsMSE.Gait = vbNullString
                 clsMSE.GaitSev = vbNullString
                 frmGait.Show vbModal
                 If clsMSE.Gait <> vbNullString Then
                     clsMSE.GaitSev = Replace(clsMSE.GaitSev, vbNewLine, "")
                     If Trim(clsMSE.GaitSev) <> vbNullString Then
                         strData = Trim(clsMSE.Gait) & " " & "of" & " " & LCase(clsMSE.GaitSev) & " " & "severity"
                         Format_CellTypeEdit grdGrid, mlngGridRow, 2, "1", strData, "TypeHAlignLeft", "DETAILS", "Click to Select Terms", grdCurItem
                         grdGrid.SetCellBorder 3, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                     ElseIf Trim(clsMSE.Gait) <> vbNullString Then
                         Format_CellTypeEdit grdGrid, mlngGridRow, 2, "1", Trim(clsMSE.Gait), "TypeHAlignLeft", "DETAILS", "Click to Select Terms", grdCurItem
                         grdGrid.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                     End If
                 End If
             ElseIf TextFound("Cranial", udtPhyExam.ExamArea, 5, True) Then
                 strTerms = "Cranial"
                 Format_GridTerms
                 mstrTerms = "sd_PhyExam_" & strTerms
                 Get_GridTerms "sd_PhyExam_" & strTerms
             Else
                 strTerms = Replace(udtPhyExam.ExamArea, " ", "_")
                 strTerms = Trim(strTerms)
                 If strTerms = "CEREBELLAR" Then
                     udtPhyExam.ExamChild = "Cerebellar"
                 ElseIf strTerms = "SENSORY" Then
                     udtPhyExam.ExamChild = "Sensory"
                 End If
                 mGrid = "grdGrid"
                 mstrTerms = "sd_PhyExam_" & udtPhyExam.ExamSystem & "_" & Trim(udtPhyExam.ExamArea)
                 Format_SpecificGrid "sd_PhyExam_" & udtPhyExam.ExamSystem & "_" & Trim(udtPhyExam.ExamArea), udtPhyExam.ExamChild
             End If
        ElseIf .CellTag = "COMMENTS" Then
            cctComments.Text = vbNullString
            .GetText Col, Row, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                cctComments.Text = Trim(varCellValue)
            End If
            fraComment.ZOrder 0
            fraComment.Visible = True
       End If
    End With
    Exit Sub
End Sub

Private Sub grdGrid_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim varCellValue
'Dim booInsert As Boolean
'Dim lngRow As Long
'Dim strSQL As String
'
'    booInsert = True
'    If mlngGridRow >= 1 Then
'        grdGrid.GetText 5, mlngGridRow, varCellValue
'        If Trim(varCellValue) <> vbNullString Then
'            Open_DB psConnect(1)
'            strSQL = "Select DISTINCT ATTRBTE_VALID_VALUE FROM ATTR_VALID_VALUE " _
'                & "Where VALID_VALUE_SUBDOM  = '" & mstrTerms & "' AND  " _
'                & "ATTRBTE_VALID_VALUE = '" & varCellValue & "'"
'            Set rs = cn.Execute(strSQL)
'            If Not rs.EOF Then
'                booInsert = False
'            Else
'                strSQL = "INSERT INTO ATTRBTE_VALID_VALUE (ATTRBTE_VALID_VALUE,VALID_VALUE_SUBDOM," _
'                    & "VALID_VALUE_SEQ_NBR) values ('" & varCellValue & "','" & mstrTerms & "'," _
'                    & mlngGridRow & ")"
'                Set rs = cn.Execute(strSQL)
'                MsgBox varCellValue & " " & "succesfully added to" & " " & mstrTerms, vbInformation
'            End If
'        End If
'        Close_DB psConnect(1)
'    End If
'
'    Exit Sub
End Sub


Private Sub grdGrid_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    With grdGrid
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = NormalColor
        .BlockMode = False
    End With
    Exit Sub
End Sub


Private Sub grdSpecLoc_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
    With grdSpecLoc
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


Private Sub grdSpecLoc_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row <= 0 Then Exit Sub
    With grdSpecLoc
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


Private Sub grdTerms_Change(ByVal Col As Long, ByVal Row As Long)
Dim strTag As String
Dim strNote As String
Dim varCellValue

    grdTerms.Col = Col
    grdTerms.Row = Row
    If grdTerms.CellTag <> vbNullString Then strTag = grdTerms.CellTag
    If grdTerms.CellNote <> vbNullString Then strNote = grdTerms.CellTag
    grdTerms.GetText Col, Row, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        udtPhyExam.Comment = varCellValue
        Format_CellTypeEdit grdTerms, Row, Col, "1", CStr(varCellValue), "TypeHAlignCenter", strTag, strNote, grdCurItem
        grdTerms.SetCellBorder Col, Row, Col, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdTerms.GetText 3, Row, varCellValue
        If Trim(varCellValue) <> vbNullString Then udtPhyExam.Finding = Trim(varCellValue)
         Write_ChildGrid Trim(varCellValue), Trim(udtPhyExam.ExamArea)
    End If
    Exit Sub
End Sub

Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue
Dim lngRow As Long

    mlngTermRow = Row
    mlngTermCol = Col
    If mlngTermRow < 0 Then Exit Sub
    With grdTerms
        .BlockMode = True
            .Col = Col
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
             .BackColor = grdCurItem
         .BlockMode = False
    End With
    grdTerms.GetText 1, mlngTermRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then udtPhyExam.ExamChild = Trim(varCellValue)
    Select Case Trim(udtPhyExam.ExamChild)
        Case "Olfactory CN I"
            udtPhyExam.ExamChild = "Olfatory"
        Case "Optical CN II"
            udtPhyExam.ExamChild = "Optical"
        Case "Oculomotor CN III"
            udtPhyExam.ExamChild = "Oculomotor"
        Case "Trochlear CN IV"
            udtPhyExam.ExamChild = "Trochlear"
        Case "Trigeminal CN V"
            udtPhyExam.ExamChild = "Trigeminal"
        Case "Abducens CN VI"
            udtPhyExam.ExamChild = "Ataducens"
        Case "Facial CN VII"
            udtPhyExam.ExamChild = "Facial"
        Case "Auditory CN VIII"
            udtPhyExam.ExamChild = "Auditory"
        Case "Glossopharyngeal CN IX"
            udtPhyExam.ExamChild = "Glossoph"
        Case "Vagus CN X"
            udtPhyExam.ExamChild = "Vagus"
        Case "Accessory CN XI"
            udtPhyExam.ExamChild = "Accessory"
        Case "Hypoglossal CN XII"
            udtPhyExam.ExamChild = "Hypoglossal"
    End Select
    If Trim(udtPhyExam.ExamChild) <> vbNullString Then
         Format_SpecificGrid "sd_PhyExam_" & Trim(udtPhyExam.ExamChild), udtPhyExam.ExamChild
    End If
    Exit Sub
End Sub


Private Sub grdTerms_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row <= 0 Then Exit Sub
    With grdTerms
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


