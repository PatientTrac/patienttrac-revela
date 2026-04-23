VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmDefaultOrder 
   BackColor       =   &H00E0E0E0&
   Caption         =   "  DEFAULT PATIENT ORDERS"
   ClientHeight    =   8850
   ClientLeft      =   3735
   ClientTop       =   2520
   ClientWidth     =   14310
   Icon            =   "frmDefaultOrder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   14310
   Begin VB.Frame fraGridTerm 
      BackColor       =   &H00E0E0E0&
      Height          =   3705
      Left            =   1200
      TabIndex        =   14
      ToolTipText     =   "Frame is movable with mouse!"
      Top             =   1740
      Visible         =   0   'False
      Width           =   5145
      Begin VB.TextBox cctSearch 
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
         Height          =   375
         Left            =   90
         TabIndex        =   15
         ToolTipText     =   "Enter either Name or Code"
         Top             =   390
         Visible         =   0   'False
         Width           =   3945
      End
      Begin PT_XP_Button.PT_XP cmdGridTerms 
         Height          =   315
         Left            =   4440
         TabIndex        =   16
         ToolTipText     =   "Terms Maintanence"
         Top             =   3330
         Width           =   465
         _ExtentX        =   820
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
         MaskColor       =   128
         Picture         =   "frmDefaultOrder.frx":08CA
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
         PictureDisabled =   "frmDefaultOrder.frx":0E64
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAddItem 
         Height          =   315
         Left            =   2100
         TabIndex        =   17
         Top             =   3270
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
         Picture         =   "frmDefaultOrder.frx":11B6
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Done"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDefaultOrder.frx":1550
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdTerms 
         Height          =   2355
         Left            =   90
         TabIndex        =   18
         Top             =   870
         Width           =   4935
         _Version        =   458752
         _ExtentX        =   8705
         _ExtentY        =   4154
         _StockProps     =   64
         ColHeaderDisplay=   0
         DAutoCellTypes  =   0   'False
         DAutoFill       =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         DisplayColHeaders=   0   'False
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
         MaxCols         =   2
         MaxRows         =   10
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   16761024
         ShadowText      =   8388608
         SpreadDesigner  =   "frmDefaultOrder.frx":18A2
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin VB.Label lblSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Type in desired diagnosis"
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
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   2115
      End
   End
   Begin VB.Frame fraTerms 
      BackColor       =   &H00FFC0C0&
      Height          =   1815
      Left            =   6060
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   4005
      Begin PT_XP_Button.PT_XP cmdMedHist 
         Height          =   495
         Left            =   210
         TabIndex        =   6
         Tag             =   "MEDICAL_ HISTORY"
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   873
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
         Caption         =   "Conditions"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdConsents 
         Height          =   495
         Left            =   2010
         TabIndex        =   7
         Tag             =   "MEDICAL_ HISTORY"
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   873
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
         Caption         =   "Consents"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdDone 
         Height          =   405
         Left            =   1410
         TabIndex        =   8
         Top             =   1080
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "frmDefaultOrder.frx":1BA4
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Done"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDefaultOrder.frx":1F3E
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.TextBox cctSummary 
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
      Height          =   1245
      Left            =   3690
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmDefaultOrder.frx":2290
      Top             =   9810
      Visible         =   0   'False
      Width           =   7845
   End
   Begin VB.Frame fraConditions 
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
      Height          =   1095
      Left            =   -30
      TabIndex        =   0
      Top             =   30
      Width           =   14265
      Begin VB.ListBox lstProviders 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   780
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   150
         Width           =   5115
      End
      Begin PT_XP_Button.PT_XP cmdTerms 
         Height          =   315
         Left            =   13680
         TabIndex        =   4
         Top             =   180
         Width           =   465
         _ExtentX        =   820
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
         MaskColor       =   128
         Picture         =   "frmDefaultOrder.frx":2296
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
         PictureDisabled =   "frmDefaultOrder.frx":2830
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblProvider 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Providers"
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
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   975
      End
   End
   Begin FPSpreadADO.fpSpread grdOrder 
      CausesValidation=   0   'False
      DragIcon        =   "frmDefaultOrder.frx":2B82
      Height          =   7125
      Left            =   0
      TabIndex        =   9
      Top             =   1710
      Width           =   14325
      _Version        =   458752
      _ExtentX        =   25268
      _ExtentY        =   12568
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      ColHeaderDisplay=   0
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   8388608
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   8
      MaxRows         =   0
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmDefaultOrder.frx":344C
      UnitType        =   0
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   10
      ScrollBarTrack  =   1
      ClipboardOptions=   0
      ShowScrollTips  =   1
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   495
      Left            =   13020
      TabIndex        =   10
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDefaultOrder.frx":3788
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      CaptionPosition =   4
      DepthMode       =   1
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   13003064
      ForeColorFocus  =   13003064
      ForeColorDown   =   13003064
      PictureDisabled =   "frmDefaultOrder.frx":419A
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdAddRow 
      Height          =   495
      Left            =   30
      TabIndex        =   11
      ToolTipText     =   "Enter New Item in Text Field First "
      Top             =   1170
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   873
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
      Picture         =   "frmDefaultOrder.frx":44EC
      AppearanceThemes=   3
      BorderDrawEvent =   1
      TransparentColor=   14737632
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmDefaultOrder.frx":4A86
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRemove 
      Height          =   495
      Left            =   1350
      TabIndex        =   12
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
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
      Picture         =   "frmDefaultOrder.frx":4DD8
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Remove"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmDefaultOrder.frx":5372
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   495
      Left            =   11430
      TabIndex        =   13
      Top             =   1170
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
      Picture         =   "frmDefaultOrder.frx":56C4
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmDefaultOrder.frx":5C5E
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmDefaultOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mControl As String
Public mlngPatientID As Long
Public mlngEncounterID As Long
Public mlngFacilityID As Long
Public mbControlLoaded As Boolean
Public mbSaved As Boolean
Private mbooByPass As Boolean
Private mFirstTime As Boolean
Public mFirst As Boolean
Private mlngGridCol As Long
Private mlngRow  As Long
Private flngDeltaX As Long
Private flngDeltaY As Long
Private fpDeltaLeft As Long
Private fpDeltaTop As Long
Private mstrSubGroup As String
Private Function Add_OrderRow()
Dim i%
Dim strList As String
Dim lngRow As Long
Dim lngCol As Long
Dim lngColWidth As Long
 
     With grdOrder
        .ReDraw = False
        .MaxRows = .MaxRows + 1
        lngRow = .MaxRows
        lngCol = 0
        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", CStr(lngRow), "TypeHAlignCenter", "ROW_COUNT"
        .SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 1
        If Trim(mstrSubGroup) = vbNullString Then Get_CPTGroups
        Format_CellTypeCombo grdOrder, lngRow, lngCol, "1", mstrSubGroup, "TypeHAlignCenter", "ORDER_TYPE"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeStatic grdOrder, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CPT_CODE"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeStatic grdOrder, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CONDITION"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'        lngCol = 4
'        strList = "Male" & Chr(9) & "Female" & Chr(9) & "Both"
'        Format_CellTypeCombo grdOrder, lngRow, lngCol, "1", strList, "TypeHAlignCenter", "GENDER"
'        .SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'        lngCol = 5
'        strList = "=>" & Chr(9) & "<="
'        Format_CellTypeCombo grdOrder, lngRow, lngCol, "1", strList, "TypeHAlignCenter", "TIME_RANGE"
'        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'        lngCol = 6
'        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "TIME_VALUE"
'        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeStatic grdOrder, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "ICD_CODE"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "COMMENT"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        strList = "Routine" & Chr(9) & "STAT"
        Format_CellTypeCombo grdOrder, lngRow, lngCol, "1", strList, "TypeHAlignCenter", "PRIORITY"
        
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 6, lngRow, 6, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    .ReDraw = True
    End With
    Exit Function

End Function

Private Function Get_CPTGroups()
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset

    On Error GoTo ERROR_Get_CPTGroups
    mstrSubGroup = vbNullString
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        cnCPT.open psConnect(1)
    Else
        cnCPT.open psConnect(1)
    End If
    strSQL = "Select * from INTRVN_SUBGROUPS "
    Set rsCPT = cnCPT.Execute(strSQL)
    With rsCPT
        Do Until .EOF
            If mstrSubGroup = vbNullString Then
                mstrSubGroup = Trim(!SUBGROUP_DESC)
            Else
                mstrSubGroup = mstrSubGroup & Chr(9) & Trim(!SUBGROUP_DESC)
            End If
        .MoveNext
        Loop
    End With
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set cnCPT = Nothing
     End If
    Exit Function
ERROR_Get_CPTGroups:
    Write_Error Err.Description, Err.Number, strSQL, "Get_CPTGroups", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function





Private Function Write_Default_Order()
Dim strSQL As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset
Dim strDate As String


    On Error GoTo Error_Write_DefaultOrder
    Screen.MousePointer = vbHourglass
    udtLaboratory.Lab_Id = 0
    strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    strSQL = "SELECT PROC_ID  FROM DEFAULT_ORDERS WHERE LAB_CODE = '" & Trim(udtLaboratory.CPT_Code) & "' AND PROVIDER_ID = " & piProviderID
    Set rsLab = cnLab.Execute(strSQL)
        If Not rsLab.EOF Then
            udtLaboratory.Lab_Id = rsLab.Fields("PROC_ID")
        End If
        If udtLaboratory.Lab_Id = 0 Then
            strSQL = "INSERT INTO DEFAULT_ORDERS (PROVIDER_ID, AGE, AGE_VALUE, GENDER, CONDITION, " _
                    & "PROC_GROUP, LAB_CODE, LAB_DESC, ICD_DESC, ICD_CODE, INTRVN_PRIORITY, COMMENT, INSERT_DATE) " _
                    & "Values(" & piProviderID & ",'" & udtLaboratory.Age & "','" & udtLaboratory.AgeRange & "','" _
                    & udtLaboratory.Gender & "','" & udtLaboratory.Med_Condition & "','" & udtLaboratory.OrderType & "','" _
                    & udtLaboratory.CPT_Code & "','" & udtLaboratory.CPT_Desc & "','" & udtLaboratory.Diag_Desc & "','" _
                    & udtLaboratory.Diag_Code & "', '" & udtLaboratory.Lab_Priority & "','" _
                    & udtLaboratory.Comment & "','" & strDate & "')"
                    Set rsLab = cnLab.Execute(strSQL)
        Else
            strSQL = "UPDATE DEFAULT_ORDERS SET AGE = '" & udtLaboratory.Age & "', AGE_VALUE = '" & udtLaboratory.AgeRange & "'," _
                    & "CONDITION = '" & udtLaboratory.Med_Condition & "', PROC_GROUP = '" & udtLaboratory.OrderType & "'," _
                    & "LAB_CODE = '" & udtLaboratory.CPT_Code & "', LAB_DESC = '" & udtLaboratory.CPT_Desc & "'," _
                    & "INTRVN_PRIORITY = '" & udtLaboratory.Lab_Priority & "', COMMENT = '" & udtLaboratory.Comment & "'," _
                    & "ICD_DESC = '" & udtLaboratory.Diag_Desc & "', ICD_CODE = '" & udtLaboratory.Diag_Code & "'," _
                    & "UPDATE_DATE = '" & strDate & "' WHERE PROC_ID =  " & udtLaboratory.Lab_Id
                Set rsLab = cnLab.Execute(strSQL)
        End If
        grdOrder.ReDraw = True
    If cnLab.State = adStateOpen Then
        cnLab.Close
        Set cnLab = Nothing
        Set rsLab = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_DefaultOrder:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_DefaultOrder", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cctSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(cctSearch.Text) > 2 Then Load_GridBox "ICD", "ICD"
End Sub


Private Sub cmdAddItem_Click()
Dim strData As String
Dim lngRow As Long
Dim varCellValue
Dim varCode
Dim Arr
Dim ArrCode
Dim lngCnt As Long
Dim strTag As String
Dim strList As String
Dim lngLabRow As Long
Dim lngCol As Long

    lngCnt = -1
    With grdTerms
        .ReDraw = True
        fraGridTerm.Visible = False
        .Col = 1
        .Row = lngRow
        If .CellTag <> vbNullString Then strTag = .CellTag
        Select Case strTag
            Case "CPT DESC", "CPT_CODE"
                For lngRow = 1 To .MaxRows
                    strData = vbNullString
                    .Col = 1
                    .Row = lngRow
                    If .BackColor = &HFFC0FF Then
                        .GetText 1, lngRow, varCellValue
                        .GetText 2, lngRow, varCode
                        If Trim(varCellValue) <> vbNullString Then
                            lngCnt = lngCnt + 1
                            strData = Trim(varCellValue) & " - " & Trim(varCode)
                            If lngCnt = 0 Then
                                Arr = Split(LookUpCellTag("CPT_CODE"), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                End If
                                Format_CellTypeEdit grdOrder, mlngRow, lngCol, "1", strData, "TypeHAlignLeft", "CPT_CODE"
                                grdOrder.SetCellBorder lngCol, mlngRow, lngCol, mlngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleFineSolid
                            Else
                                Add_OrderRow
                                lngLabRow = grdOrder.MaxRows
                                Arr = Split(LookUpCellTag("ORDER_TYPE"), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                End If
                                If Trim(mstrSubGroup) = vbNullString Then Get_CPTGroups
                                Format_CellTypeCombo grdOrder, lngLabRow, lngCol, "1", mstrSubGroup, "TypeHAlignCenter", "ORDER_TYPE"
                                Set_GridCurSel grdOrder, clsGridFormat.FormatTermGridType, lngCol, lngLabRow, "1", "ORDER_TYPE"
                                grdOrder.SetCellBorder lngCol, lngLabRow, lngCol, lngLabRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                Arr = Split(LookUpCellTag("CPT_CODE"), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                End If
                                Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", strData, "TypeHAlignLeft", "CPT_CODE"
                                grdOrder.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        End If
                    End If
                Next lngRow
            Case "ICD_DESC", "ICD_CODE"
                 For lngRow = 1 To grdTerms.MaxRows
                    .Col = 1
                    .Row = lngRow
                    If .BackColor = &HFFC0FF Then
                        .GetText 1, lngRow, varCellValue
                        .GetText 2, lngRow, varCode
                        If Trim(varCellValue) <> vbNullString Then
                            If strData = vbNullString Then
                                strData = Trim(varCellValue) & " - " & Trim(varCode)
                            Else
                                strData = strData & vbNewLine & Trim(varCellValue) & " - " & Trim(varCode)
                            End If
                        End If
                     End If
                Next lngRow
                lngRow = 0
                For lngRow = mlngRow To grdOrder.MaxRows
                    Arr = Split(LookUpCellTag("ICD_CODE"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                    End If
                    grdOrder.GetText lngCol, lngRow, varCellValue
                    If varCellValue = " - " Then varCellValue = vbNullString
                    If Trim(varCellValue) = vbNullString And Trim(strData) <> vbNullString Then
                        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", strData, "TypeHAlignLeft", "ICD_CODE"
                        grdOrder.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleFineSolid
                    End If
                Next lngRow
            Case "CONDITION"
                 For lngRow = 1 To grdTerms.MaxRows
                    .Col = 1
                    .Row = lngRow
                    If .BackColor = &HFFC0FF Then
                        .GetText 1, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            If strData = vbNullString Then
                                strData = Trim(varCellValue)
                            End If
                        End If
                     End If
                Next lngRow
                lngRow = 0
                Arr = Split(LookUpCellTag("CONDITION"), "|")
                If UBound(Arr) > 0 Then
                    lngCol = Arr(0)
                End If
                For lngRow = mlngRow To grdOrder.MaxRows
                    grdOrder.GetText lngCol, lngRow, varCellValue
                    If Trim(varCellValue) = vbNullString And Trim(strData) <> vbNullString Then
                        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", strData, "TypeHAlignLeft", "CONDITION"
                        grdOrder.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleFineSolid
                    End If
                Next lngRow
            Case "MODIFIER"
                 For lngRow = 1 To .MaxRows
                    .Col = 1
                    .Row = lngRow
                    If .BackColor = &HFFC0FF Then
                        .GetText 1, lngRow, varCellValue
                        .GetText 2, lngRow, varCode
                        If Trim(varCellValue) <> vbNullString Then
                            If strData = vbNullString Then
                                strData = Trim(varCellValue) & " - " & Trim(varCode)
                            Else
                                strData = strData & vbNewLine & Trim(varCellValue) & " - " & Trim(varCode)
                            End If
                        End If
                     End If
                Next lngRow
                Arr = Split(LookUpCellTag("MODIFIER"), "|")
                If UBound(Arr) > 0 Then
                    lngCol = Arr(0)
                End If
                If Trim(strData) <> vbNullString Then
                    Format_CellTypeEdit grdOrder, mlngRow, lngCol, "1", strData, "TypeHAlignLeft", "MODIFIER"
                    grdOrder.SetCellBorder lngCol, mlngRow, lngCol, mlngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleFineSolid
                End If
            End Select
        .ReDraw = True
    End With
End Sub

Private Sub cmdAddRow_Click()
    Add_OrderRow
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGridTerms_Click()
    Select Case clsGridFormat.CellTag
        Case "CPT_CODE"
            frmCPTMaint.Show vbModal
        Case "CONDITION"
            Load frmNurseTemplates
            frmNurseTemplates.fCaption = "Order Conditions"
            frmNurseTemplates.fType = "Order Conditions"
            frmNurseTemplates.fSubDomain = "sd_LabOrder_Condition"
            frmNurseTemplates.fConnectTM = psConnect(1)
            frmNurseTemplates.ReadData
            frmNurseTemplates.Show vbModal
            Load_GridBox "CONDITION", "CONDITION"
    End Select
    
    Exit Sub
End Sub

Private Sub grdOrder_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String

    clsGridFormat.FormatTermGridType = vbNullString
      With grdOrder
         mlngRow = Row
        .ReDraw = False
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then strTag = .CellTag
        If Trim(strTag) <> vbNullString Then
            Select Case UCase(strTag)
                Case "ORDER_TYPE"
                    If .TypeComboBoxCurSel >= 0 Then .TypeComboBoxIndex = .TypeComboBoxCurSel
                    clsGridFormat.FormatTermGridType = .TypeComboBoxString
            End Select
        End If
    End With
    Exit Sub
End Sub
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
        .TypeVAlign = TypeVAlignCenter
        If vlngBackColor > 0 Then .BackColor = vlngBackColor
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
        If vstrCellNote <> vbNullString Then .CellNote = Trim(vstrCellNote)
    .ReDraw = True
    End With
End Function

Private Sub grdOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim strSQL As String
Dim lngX As Long
Dim lngY As Long
Dim lngWidth As Long
Dim lngHeight As Long
Dim CellNote As Long
Dim varCellValue
Dim Arr
Dim lngCol As Long
Dim strCurrentDiag As String
Dim strTag As String
 
    clsGridFormat.CellTag = vbNullString
    With grdOrder
         mlngRow = Row
        .ReDraw = False
        .GetText 1, Row, varCellValue
        If Trim(varCellValue) <> vbNullString Then strTag = Trim(varCellValue)
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then clsGridFormat.CellTag = .CellTag
        Select Case clsGridFormat.CellTag
            Case "CPT_CODE"
                .Col = 2
                .Row = Row
                .GetCellPos .Col, .Row, lngX, lngY, lngWidth, lngHeight
                fpDeltaTop = lngY + 51.5
                fpDeltaLeft = lngX + 50
                fraGridTerm.ZOrder 0
                fraGridTerm.Top = .Top + fpDeltaTop
                fraGridTerm.Left = .Left + fpDeltaLeft
                Load_GridBox clsGridFormat.FormatTermGridType, "CPT_CODE"
                lblSearch.Visible = False
                cctSearch.Visible = False
                fraGridTerm.Visible = True
            Case "CONDITION"
                .Col = 3
                .Row = Row
                .GetCellPos .Col, .Row, lngX, lngY, lngWidth, lngHeight
                fpDeltaTop = lngY + 51.5
                fpDeltaLeft = lngX + 50
                fraGridTerm.ZOrder 0
                fraGridTerm.Top = .Top + fpDeltaTop
                fraGridTerm.Left = .Left + fpDeltaLeft
                If Trim(strTag) <> vbNullString Then Load_GridBox "CONDITION", "CONDITION"
                lblSearch.Visible = False
                cctSearch.Visible = False
                fraGridTerm.Visible = True
            Case "ICD_CODE"
                .Col = 4
                .Row = Row
                .GetCellPos .Col, .Row, lngX, lngY, lngWidth, lngHeight
                fpDeltaTop = lngY + 51.5
                fpDeltaLeft = lngX + 50
                fraGridTerm.ZOrder 0
                fraGridTerm.Top = .Top + fpDeltaTop
                fraGridTerm.Left = .Left + fpDeltaLeft
                If Trim(strTag) <> vbNullString Then Load_GridBox "ICD", "ICD"
                lblSearch.Visible = True
                cctSearch.Visible = True
                fraGridTerm.Visible = True
        End Select
        .ReDraw = True
    End With
    Exit Sub
End Sub
Private Function Load_CPTIntrvn(vstrSubGroup As String)
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim iCnt As Integer
Dim strData As String
Dim lngRow As Long
Dim mintGroupID As Integer
Dim mintLabGroupID As Integer

    On Error GoTo Error_Load_CptIntrvn
   Screen.MousePointer = vbHourglass
    If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
    strSQL = "Select  a.Group_id, b.SUBGROUP_ID " _
            & "From INTRVN_GROUPS a, INTRVN_SUBGROUPS b " _
            & "Where a.GROUP_DESC Like 'Superbill' And b.SUBGROUP_DESC Like '" & vstrSubGroup & "'"
        Set rsCPT = cnCPT.Execute(strSQL)
        With rsCPT
            If Not .EOF Then
                mintGroupID = (!Group_ID)
                mintLabGroupID = (!SUBGROUP_ID)
            End If
        End With
    strSQL = "Select CPT_INTRVN_DESC,  CPT_INTRVN_CODE " _
            & "From INTRVN_TEMPLATE " _
            & "Where group_id = " & mintGroupID & " And SUBGROUP_ID = " & mintLabGroupID _
            & " Order By SUBGROUP_ID "
       Set rsCPT = cnCPT.Execute(strSQL)
       With rsCPT
            Do Until .EOF
                If Not IsNull(!CPT_Intrvn_Desc) Then strData = (!CPT_Intrvn_Desc)
                If Trim(strData) <> vbNullString Then
                    lngRow = lngRow + 1
                    grdTerms.MaxRows = lngRow
                    Format_CellTypeEdit grdTerms, lngRow, 1, "1", Trim(strData), "TypeHAlignLeft", vstrSubGroup
                    If IsNumeric(!CPT_Intrvn_Code) Then grdTerms.SetRowItemData (lngRow), (!CPT_Intrvn_Code)
                    If Not IsNull(!CPT_Intrvn_Code) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!CPT_Intrvn_Code), "TypeHAlignCenter", vstrSubGroup
                    grdTerms.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
                .MoveNext
            Loop
        End With
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_CptIntrvn:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Lab", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Load_GridBox(vstrCol As String, vstrType As String)
Dim strSQL As String
Dim iRow As Long
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim iCnt As Integer
Dim strData As String
Dim lngRow As Long
Dim Arr
Dim booData As Boolean


    On Error GoTo Error_Load_GridBox
    Screen.MousePointer = vbHourglass
    If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
    Select Case Trim(vstrType)
        Case "CPT_CODE"
            Format_GridTerms "CPT_CODE"
            Load_CPTIntrvn clsGridFormat.FormatTermGridType
        Case "ICD"
            Format_GridTerms "ICD"
            If Trim(cctSearch.Text) <> vbNullString Then
                If Trim(cctSearch.Text) Like "[A-Za-z]###" Or Trim(cctSearch.Text) Like "###" Then
                    grdTerms.ReDraw = False
                    strSQL = "SELECT LONG_DESC, CODE FROM ICD_CODE WHERE CODE like '" & Trim(cctSearch.Text) & "%' ORDER BY CODE"
                    Set rsCPT = cnCPT.Execute(strSQL)
                    With rsCPT
                        Do Until .EOF
                            If Not IsNull(!Long_Desc) Then strData = (!Long_Desc)
                            If Trim(strData) <> vbNullString Then
                                booData = True
                                lngRow = lngRow + 1
                                grdTerms.MaxRows = lngRow
                                Format_CellTypeEdit grdOrder, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "ICD_DESC"
                                If Not IsNull(!code) Then
                                    If IsNumeric(!code) Then grdTerms.SetRowItemData (lngRow), (!code)
                                    If Not IsNull((!code)) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!code), "TypeHAlignCenter", "ICD_CODE"
                                    grdTerms.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                End If
                            End If
                            .MoveNext
                        Loop
                    End With
                Else
                    strSQL = "select LONG_DESC, CODE from icd_code Where LONG_DESC like  '%" & Trim(cctSearch.Text) & "%' ORDER BY LONG_DESC "
                    Set rsCPT = cnCPT.Execute(strSQL)
                    With rsCPT
                        Do Until .EOF
                            If (!Long_Desc) <> vbNullString Then strData = (!Long_Desc)
                            If Trim(strData) <> vbNullString Then
                                booData = True
                                lngRow = lngRow + 1
                                grdTerms.MaxRows = lngRow
                                Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "ICD_DESC"
                                If Not IsNull(!code) Then
                                    If IsNumeric(!code) Then grdTerms.SetRowItemData (lngRow), (!code)
                                    If Not IsNull((!code)) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!code), "TypeHAlignCenter", "ICD_CODE"
                                    grdTerms.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                End If
                            End If
                         .MoveNext
                        Loop
                    End With
                End If
            End If
        Case "MODIFIER"
            grdTerms.ReDraw = False
            Format_GridTerms "MODIFIER"
            strSQL = "select DISTINCT CPT_MOD_DESC, CPT_MOD_CODE " _
                & "From CPT_MOD_TEMPLATE  "
            Set rsCPT = cnCPT.Execute(strSQL)
            With rsCPT
                Do Until .EOF
                    If Not IsNull(!CPT_MOD_DESC) Then strData = (!CPT_MOD_DESC)
                    If Trim(strData) <> vbNullString Then
                        lngRow = lngRow + 1
                        grdTerms.MaxRows = lngRow
                        Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "MODIFIER"
                        If IsNumeric(!CPT_MOD_CODE) Then grdTerms.SetRowItemData (lngRow), (!CPT_MOD_CODE)
                        If Not IsNull(!CPT_MOD_CODE) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!CPT_MOD_CODE), "TypeHAlignLeft", "MODIFIER"
                        grdTerms.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                    .MoveNext
                Loop
            End With
        Case "CONDITION"
            grdTerms.ReDraw = False
            Format_GridTerms "CONDITION"
            strSQL = "SELECT ATTRBTE_VALID_VALUE,VALID_VALUE_SEQ_NBR FROM ATTR_VALID_VALUE WHERE " _
                & "VALID_VALUE_SUBDOM =  'sd_LabOrder_Condition' ORDER BY VALID_VALUE_SEQ_NBR"
            Set rsCPT = cnCPT.Execute(strSQL)
            With rsCPT
                Do Until .EOF
                    If Not IsNull(!ATTRBTE_VALID_VALUE) Then strData = (!ATTRBTE_VALID_VALUE)
                    If Trim(strData) <> vbNullString Then
                        lngRow = lngRow + 1
                        grdTerms.MaxRows = lngRow
                        Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "CONDITION"
                        grdTerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                    .MoveNext
                Loop
            End With
    End Select
    grdTerms.ReDraw = True
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_GridBox:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_GridBox", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function




Private Sub grdOrder_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim varCondition
    mlngRow = Row
    mlngGridCol = Col
     If mlngGridCol = 6 Then
        With grdOrder
            .GetText mlngGridCol, mlngRow, varCondition
            If varCondition <> Empty Then
                .BlockMode = True
                    .Col = mlngGridCol
                    .Col2 = mlngGridCol
                    .Row = mlngRow
                    .Row2 = mlngRow
                    .BackColor = &HC0C0FF
                    .CellNote = "NOTE: Pink Background denotes Negative." & vbNewLine & " Right click mouse to change to Patient reports Negative for" & " " & varCondition
                .BlockMode = False
            End If
        End With
  End If

    Exit Sub
End Sub

Private Sub grdOrder_TextTipFetch(ByVal Col As Long, ByVal Row As Long, Multiline As FPSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
    If grdOrder.IsFetchCellNote Then
        ShowTip = True
    Else
        ShowTip = False
    End If
End Sub



Private Function Select_DefaultOrder()
Dim Arr
Dim lngRow As Long
Dim varCellValue
Dim lngCol As Long

    On Error GoTo Error_Select_DefaultOrder
    Screen.MousePointer = vbHourglass

    With grdOrder
        For lngRow = 1 To grdOrder.MaxRows
            Clear_LabString
            udtLaboratory.Lab_Id = .GetRowItemData(lngRow)
            For lngCol = 1 To .MaxCols
                grdOrder.Col = lngCol
                grdOrder.Row = lngRow
                If grdOrder.CellTag <> vbNullString Then
                    Select Case UCase(grdOrder.CellTag)
                        Case "ORDER_TYPE"
                            .GetText lngCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                udtLaboratory.OrderType = Trim(varCellValue)
                            End If
                        Case "TIME_RANGE"
                            .GetText lngCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                udtLaboratory.AgeRange = Trim(varCellValue)
                            End If
                        Case "TIME_VALUE"
                            .GetText lngCol, lngRow, varCellValue
                            If Trim(Trim(varCellValue)) <> vbNullString Then
                                udtLaboratory.Age = Trim(varCellValue)
                            End If
                        Case "GENDER"
                            .GetText lngCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                udtLaboratory.Gender = Trim(varCellValue)
                            End If
                        Case "CONDITION"
                            .GetText lngCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                udtLaboratory.Med_Condition = Trim(varCellValue)
                            End If
                        Case "COMMENT"
                            .GetText lngCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                udtLaboratory.Spec_Instruct = Trim(varCellValue)
                            End If
                        Case "CPT_CODE"
                            .GetText lngCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                Arr = Split(varCellValue, "-")
                                If UBound(Arr) > 0 Then
                                    udtLaboratory.CPT_Desc = Trim(Arr(0))
                                    udtLaboratory.CPT_Code = Trim(Arr(1))
                                End If
                            End If
                        Case "PRIORITY"
                            .GetText lngCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                udtLaboratory.Lab_Priority = Trim(varCellValue)
                            End If
                        Case "ICD_CODE"
                            .GetText lngCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                Arr = Split(varCellValue, "-")
                                If UBound(Arr) > 0 Then
                                    udtLaboratory.Diag_Desc = Trim(Arr(0))
                                    udtLaboratory.Diag_Code = Trim(Arr(1))
                                End If
                            End If
                    End Select
                End If
            Next lngCol
            If Trim(udtLaboratory.OrderType) <> vbNullString Then
                Write_Default_Order
            End If
        Next lngRow
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_Select_DefaultOrder:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Select_DefaultOrder", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Public Function Read_DefaultOrder()
Dim strSQL As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim Arr
Dim lngRow As Long
Dim lngCol As Long
Dim strProvider As String
Dim i%

    On Error GoTo Error_Read_DefaultOrder
    Screen.MousePointer = vbHourglass
    List_Providers
    Get_Providers piProviderID
    strProvider = clsProvider.Name
    mbooByPass = True
    For i = 0 To lstProviders.listcount - 1
        If lstProviders.ItemData(i) = piProviderID Then
            lstProviders.Selected(i) = True
            Exit For
        End If
    Next i
    mbooByPass = False
    Format_Headers
    If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    strSQL = "SELECT * FROM DEFAULT_ORDERS WHERE PROVIDER_ID = " & piProviderID
    Set rsLab = cnLab.Execute(strSQL)
    With rsLab
        Do Until .EOF
            Clear_LabString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PROC_ID"
                                 udtLaboratory.Lab_Id = Fld.Value
                            Case "PROC_GROUP"
                                udtLaboratory.OrderType = Fld.Value
                            Case "AGE_VALUE"
                                udtLaboratory.AgeRange = Fld.Value
                            Case "AGE"
                                udtLaboratory.Age = Fld.Value
                            Case "GENDER"
                                udtLaboratory.Gender = Fld.Value
                            Case "CONDITION"
                                 udtLaboratory.Med_Condition = Fld.Value
                            Case "LAB_CODE"
                                udtLaboratory.CPT_Code = Fld.Value
                            Case "LAB_DESC"
                                udtLaboratory.CPT_Desc = Fld.Value
                            Case "ICD_CODE"
                                udtLaboratory.Diag_Code = Fld.Value
                            Case "ICD_DESC"
                                udtLaboratory.Diag_Desc = Fld.Value
                            Case "PROVIDER_ID"
                                udtLaboratory.ProviderID = Fld.Value
                            Case "CONSULT"
                                udtLaboratory.PatientNotice = Fld.Value
                            Case "INTRVN_PRIORITY"
                                udtLaboratory.Lab_Priority = Fld.Value
                            Case "COMMENTS"
                                udtLaboratory.Spec_Instruct = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtLaboratory.Lab_Id > 0 Then
                Add_OrderRow
                lngRow = grdOrder.MaxRows
                grdOrder.SetRowItemData lngRow, udtLaboratory.Lab_Id
                For lngCol = 1 To grdOrder.MaxCols
                    grdOrder.Col = lngCol
                    grdOrder.Row = lngRow
                    If grdOrder.CellTag <> vbNullString Then
                        Select Case UCase(grdOrder.CellTag)
                            Case "ROW_COUNT"
                                Format_CellTypeStatic grdOrder, lngRow, lngCol, "1", CStr(lngRow), "TypeHAlignCenter", "ROW_COUNT"
                            Case "ORDER_TYPE"
                                Set_GridCurSel grdOrder, udtLaboratory.OrderType, lngCol, lngRow, "1", "ORDER_TYPE"
                            Case "TIME_RANGE"
                                Format_CellTypeStatic grdOrder, lngRow, lngCol, "1", udtLaboratory.AgeRange, "TypeHAlignCenter", "TIME_RANGE"
                            Case "TIME_VALUE"
                                Format_CellTypeStatic grdOrder, lngRow, lngCol, "1", udtLaboratory.Age, "TypeHAlignCenter", "TIME_RANGE"
                            Case "GENDER"
                                Format_CellTypeStatic grdOrder, lngRow, lngCol, "1", udtLaboratory.Gender, "TypeHAlignCenter", "GENDER"
                            Case "CONDITION"
                                Format_CellTypeStatic grdOrder, lngRow, lngCol, "1", udtLaboratory.Med_Condition, "TypeHAlignCenter", "CONDITION"
                            Case "COMMENT"
                                Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", udtLaboratory.Spec_Instruct, "TypeHAlignCenter", "COMMENT"
                            Case "CPT_CODE"
                                Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", udtLaboratory.CPT_Desc & " - " & udtLaboratory.CPT_Code, "TypeHAlignCenter", "CPT_CODE"
                            Case "ICD_CODE"
                                Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", udtLaboratory.Diag_Desc & " - " & udtLaboratory.Diag_Code, "TypeHAlignCenter", "ICD_CODE"
                            Case "PRIORITY"
                                Set_GridCurSel grdOrder, udtLaboratory.Lab_Priority, lngCol, lngRow, "1", "PRIORITY"
                        End Select
                    End If
                Next lngCol
            End If
           Loop
        End With
        grdOrder.Col = 1
        grdOrder.Row = 1
    If cnLab.State = adStateOpen Then
        Set cnLab = Nothing
        Set rsLab = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_DefaultOrder:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, " Read_DefaultOrder", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function



Private Sub cmdConsents_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Legal Consents / Instructions"
    frmNurseTemplates.fType = "Legal Consents / Instructions"
    frmNurseTemplates.fSubDomain = "sd_Legal_Consents"
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
End Sub

Private Sub cmdDone_Click()
    fraTerms.Visible = False
    fraTerms.ZOrder 1
End Sub

Private Sub cmdMedHist_Click()
'    Load frmMedHistTerms
'    frmMedHistTerms.fCaption = "Medical History Terms"
'    frmMedHistTerms.fType = "Medical History Terms"
'    frmMedHistTerms.fSubDomain = "sd_LabOrder_Condition"
'    frmMedHistTerms.fConnectTM = psConnect(1)
'    frmMedHistTerms.ReadData
'    frmMedHistTerms.Show vbModal
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Order Conditions"
    frmNurseTemplates.fType = "Order Conditions"
    frmNurseTemplates.fSubDomain = "sd_LabOrder_Condition"
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
End Sub

Private Sub cmdRemove_Click()
Dim strSQL As String
Dim iRows As Long
Dim iAnswer As Integer
Dim varCode
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim lngCPTCode As Long
Dim lngPrimKey As Long

    grdOrder.Col = mlngGridCol
    grdOrder.Row = mlngRow
    On Error GoTo Error_Remove
    Screen.MousePointer = vbHourglass
    If mlngRow > 0 And grdOrder.BackColor = &HFFC0FF Then
        With grdOrder
            .ReDraw = False
            lngPrimKey = .GetRowItemData(mlngRow)
            iAnswer = MsgBox("Are you sure you want to delete " & varCode, vbYesNo)
            If iAnswer = vbYes Then
                    If lngPrimKey > 0 Then
                        If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
                        strSQL = "DELETE  DEFAULT_ORDERS WHERE PROC_ID = " & lngPrimKey
                        Set rsCPT = cnCPT.Execute(strSQL)
                    End If
                    grdOrder.DeleteRows mlngRow, 1
                    .MaxRows = .MaxRows - 1
                    MsgBox "succesfully removed!", vbInformation
            Else
                Exit Sub
            End If
            .ReDraw = False
        End With
    Else
        MsgBox "Please select row first!", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Remove:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Remove", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdSave_Click()
    Select_DefaultOrder
    Exit Sub
End Sub

Private Sub cmdTerms_Click()
    fraTerms = True
    fraTerms.Visible = True
    fraTerms.ZOrder 0
End Sub



Private Sub Form_Activate()
    If mFirst Then
        Read_DefaultOrder
        mFirst = False
    End If
End Sub

Private Sub Form_Load()
    mFirst = True
End Sub

Private Function Get_Procedures()
'Dim strSQL As String
'Dim vRS
'Dim bRet As Boolean
'Dim iRow As Long
'Dim iCnt As Integer
'
'        mstrProcList = Empty
'        mlngProcList = 0
'
'         Open_DB psConnect(1)
'        cn.S_ATTR_VALID_VALUE "sd_Order_ProcedureGroups", rs
'
'        If Not rs.EOF Then
'           While Not rs.EOF
'            If mstrProcList = Empty Then
'                If Not IsNull(rs.Fields("ATTRBTE_VALID_VALUE")) Then mstrProcList = rs.Fields("ATTRBTE_VALID_VALUE")
'            Else
'                If Not IsNull(rs.Fields("ATTRBTE_VALID_VALUE")) Then mstrProcList = mstrProcList & Chr(9) & rs.Fields("ATTRBTE_VALID_VALUE")
'            End If
'            rs.MoveNext
'    Wend
'    End If
'
'    Close_DB psConnect(1)
'
'    Exit Function
End Function
Private Function LookUpCPTCode(vstrCPTName As String) As String
Dim strSQL As String
Dim cnCode As New ADODB.Connection
Dim rsCode As New ADODB.Recordset
Dim strProvider As String

    On Error GoTo Error_LookUpCPTCode
    Screen.MousePointer = vbHourglass
    
    LookUpCPTCode = vbNullString
    If cnCode.State = adStateClosed Then cnCode.open psConnect(1)
    strSQL = "Select CPT_CODE, CPT_SUBGROUP " _
        & "From CPT_TEMPLATE " _
        & "where  CPT_NAME like  '" & vstrCPTName & "%' "
    Set rsCode = cnCode.Execute(strSQL)
    With rsCode
        If Not .EOF Then
            If Not IsNull((!CPT_Code)) Then
                If (!CPT_Code) <> vbNullString Then LookUpCPTCode = Trim(!CPT_Code)
            End If
        End If
        rsCode.Close
    End With
    If cnCode.State = adStateOpen Then
        cnCode.Close
        Set cnCode = Nothing
        Set rsCode = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUpCPTCode:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUpCPTCode", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function
Private Function LookUpCPTName(vstrCPTCode As String) As String
Dim i%
Dim strSQL As String
Dim cnCode As New ADODB.Connection
Dim rsCode As New ADODB.Recordset
Dim Arr

    On Error GoTo Error_LookUpCPTName
    Screen.MousePointer = vbHourglass
    LookUpCPTName = vbNullString
    Arr = Split(vstrCPTCode, ":")
    If UBound(Arr) > 0 Then
        vstrCPTCode = Arr(0)
    End If
     clsGridFormat.Text = vbNullString
    If vstrCPTCode <> vbNullString Then
        If cnCode.State = adStateClosed Then cnCode.open psConnect(1)
        strSQL = "Select CPT_NAME FROM CPT_TEMPLATE WHERE CPT_CODE like  '" & vstrCPTCode & "%'"
        Set rsCode = cnCode.Execute(strSQL)
        With rsCode
            If Not .EOF Then
                If Not IsNull((!CPT_Name)) Then
                    If (!CPT_Name) <> vbNullString Then LookUpCPTName = Trim(!CPT_Name)
                End If
            End If
        End With
        If LookUpCPTName = vbNullString Then
            strSQL = "Select CPT_INTRVN_NAME, CPT_FEE FROM INTRVN_TEMPLATE WHERE CPT_INTRVN_CODE like  '" & vstrCPTCode & "%'"
                Set rsCode = cnCode.Execute(strSQL)
                With rsCode
                    If Not .EOF Then
                        If Not IsNull((!CPT_Intrvn_Name)) Then
                            If (!CPT_Intrvn_Name) <> vbNullString Then LookUpCPTName = Trim(!CPT_Intrvn_Name)
                        End If
                        If Not IsNull((!CPT_Fee)) Then
                            If (!CPT_Fee) <> vbNullString Then clsGridFormat.Text = Trim(!CPT_Fee)
                        End If
                    End If
                End With
        End If
    End If
    If cnCode.State = adStateOpen Then
        cnCode.Close
        Set cnCode = Nothing
        Set rsCode = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUpCPTName:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUpCPTName", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function


Private Function LookUpCellTag(vstrname As String) As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim booExit As Boolean

    On Error GoTo Error_LookUpCellTag
    Screen.MousePointer = vbHourglass
    With grdOrder
        LookUpCellTag = vbNullString
        For lngCol = 1 To .MaxCols
            For lngRow = 1 To .MaxRows
                .Col = lngCol
                .Row = lngRow
                If .CellTag <> vbNullString Then
                    If TextFound(vstrname, .CellTag, Len(vstrname), True) Then
                        LookUpCellTag = lngCol & "|" & lngRow
                        booExit = True
                        Exit For
                    End If
                End If
            Next lngRow
            If booExit Then Exit For
        Next lngCol
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUpCellTag:
    Screen.MousePointer = vbDefault
   Exit Function
    Resume
End Function
Private Sub Format_Headers()
Dim lngRow As Long
Dim lngCol As Long

    Screen.MousePointer = vbHourglass
    With grdOrder
        .ReDraw = False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .FontName = "Times New Roman"
        .GridSolid = False
        .GridShowHoriz = False
        .GridShowVert = False
        .RowHeadersShow = True
        .RowHeadersAutoText = DispNumbers
        .ColHeadersShow = True
        .MaxCols = 6
        .ColHeadersAutoText = DispBlank
        Add_Header
        lngRow = SpreadHeader + 2
        lngCol = 0
        .ColWidth(lngCol) = 5
        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", "ROW", "TypeHAlignCenter", "ROW_COUNT", HeaderColor
        lngCol = 1
        .ColWidth(lngCol) = 20
        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", "ORDER TYPE", "TypeHAlignCenter", "ORDER_TYPE", HeaderColor
        lngCol = 2
        .ColWidth(lngCol) = 20
        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", "ORDER NAME", "TypeHAlignCenter", "CPT_CODE", HeaderColor
        lngCol = 3
        .ColWidth(lngCol) = 20
        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", "ORDER CONDITION", "TypeHAlignCenter", "CONDITION", HeaderColor
       ' lngCol = 4
       ' .ColWidth(lngCol) = 7
      '  Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", "GENDER", "TypeHAlignCenter", "GENDER", HeaderColor
     '   lngCol = 5
      '  .ColWidth(lngCol) = 7
      '  Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", "RANGE", "TypeHAlignCenter", "TIME_RANGE", HeaderColor
    '    lngCol = 6
      '  .ColWidth(lngCol) = 10
     '   Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", "AGE" & vbNewLine & "TIME", "TypeHAlignCenter", "TIME_VALUE", HeaderColor
        lngCol = 4
        .ColWidth(lngCol) = 20
        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", "DX REASON", "TypeHAlignCenter", "ICD_CODE", HeaderColor
        lngCol = 5
        .ColWidth(lngCol) = 15
        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", "INSTRUCTIONS", "TypeHAlignCenter", "COMMENT", HeaderColor
        lngCol = 6
        .ColWidth(lngCol) = 15
        Format_CellTypeEdit grdOrder, lngRow, lngCol, "1", "PRIORITY", "TypeHAlignCenter", "PRIORITY", HeaderColor
        grdOrder.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 6, lngRow, 6, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 7, lngRow, 7, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 8, lngRow, 8, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.SetCellBorder 9, lngRow, 9, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdOrder.ReDraw = True
         Add_Spacer SpreadHeader + 3, grdOrder
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Function Add_Header()
Dim CellNote As Long
Dim strPhoto As String
Dim Idx As Integer
Dim fso As New FileSystemObject

    On Error GoTo Error_Add_Header
    Screen.MousePointer = vbHourglass
    Get_DefaultFacility
    grdOrder.ColHeaderRows = 4
    grdOrder.ColHeadersShow = True
    If udtFacility.Logo <> vbNullString Then
        If fso.FileExists(udtFacility.Logo) Then
            Format_CellTypePicture grdOrder, SpreadHeader, 1, "1", udtFacility.Logo, "TypeHAlignCenter", "LOGO", "75"
            If fso.FileExists(psImageFilePath & mlngPatientID & "\Photo.jpg") And Trim(udtPatientName.IntId) <> vbNullString Then
                strPhoto = psImageFilePath & mlngPatientID & "\Photo.jpg "
                Format_CellTypeEdit grdOrder, SpreadHeader, 2, "7", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
                Format_CellTypePicture grdOrder, SpreadHeader, 9, "2", strPhoto, "TypeHAlignCenter"
                grdOrder.SetCellBorder 9, SpreadHeader, 10, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Else
                Format_CellTypeEdit grdOrder, SpreadHeader, 2, "maxcols", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
            End If
        End If
    End If
    Add_Spacer SpreadHeader + 1, grdOrder
    Screen.MousePointer = vbDefault
    Exit Function
Error_Add_Header:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Add_Header", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function

Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)
    With vObjGrid
        vlngRow = vlngRow
       If .MaxRows < vlngRow Then .MaxRows = vlngRow
        .BlockMode = True
            .Col = 1
            .Col2 = 1
            .Row = vlngRow
            .Row2 = vlngRow
            .RowHeight(vlngRow) = 3
            .BackColor = 12632256
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 1, vlngRow, .MaxCols, vlngRow, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub
Private Function Format_GridTerms(vstrType As String)
    Screen.MousePointer = vbHourglass
    With grdTerms
        .ReDraw = False
        .ClearRange 0, SpreadHeader, grdTerms.MaxCols, grdTerms.MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .ColHeaderRows = 2
        .ColHeadersShow = True
        .RowHeadersShow = False
        .ColWidth(1) = 28
        .ColWidth(2) = 10
        If vstrType = "ICD" Then
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "DIAGNOSIS", "TypeHAlignCenter", "ICD_DESC", &HC0FFFF
            Format_CellTypeEdit grdTerms, SpreadHeader, 2, "maxcols", "DX CODE", "TypeHAlignCenter", "ICD_CODE", &HC0FFFF
            grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        ElseIf vstrType = "CONDITION" Then
            .ColWidth(1) = 38
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "CONDITION", "TypeHAlignCenter", "CONDITION", &HC0FFFF
            grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        ElseIf vstrType = "CPT_CODE" Then
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "CPT DESC", "TypeHAlignCenter", "CPT DESC", &HC0FFFF
            Format_CellTypeEdit grdTerms, SpreadHeader, 2, "maxcols", "CPT CODE", "TypeHAlignCenter", "CPT_CODE", &HC0FFFF
            grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        Add_Spacer SpreadHeader + 1, grdTerms
        grdTerms.MaxRows = 0
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Function
End Function



Private Function List_Providers()
Dim strSQL As String
Dim cnProv As New ADODB.Connection
Dim rsProv As New ADODB.Recordset
Dim lngProv As Long
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
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

    On Error GoTo Error_List_Providers
    Screen.MousePointer = vbHourglass

    lstProviders.Clear
    If cnProv.State = adStateClosed Then cnProv.open psConnect(0)
    strSQL = "SELECT * FROM  PROVIDERS  WHERE ACTIVE = 1"
    Set rsProv = cnProv.Execute(strSQL)
    With rsProv
        Do Until .EOF
            lngProv = 0
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            If Not IsNull(!Title) Then
                strTitle = Trim((!Title))
            End If
            booSkip = True
            
            Select Case UCase(Replace(strTitle, ".", ""))
                Case "MD", "RN", "PHD", "PH.D", "LSW", "LCSW", "APN", "LPN", "LGSW", "MSW", "LGPC", "LCPC", "LCS", _
                    "PA", "APRN", "NP", "MA", "MS", "DO", "D.O.", "PA-C", "LPC", "LCSW-C", "LCSWC", "PSY-D", "CRNP", "LCADC", "LCAD-C", "MPH", "ICMS"
                    booSkip = False
            End Select
            If Not booSkip Then
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Fld.Value <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "FIRST_NAME"
                                    strFirst = Fix_Name(Fld.Value)
                                Case "MIDDLE_NAME"
                                    strMiddle = Fix_Name(Fld.Value)
                                Case "LAST_NAME"
                                    strLast = Fix_Name(Fld.Value)
                                Case "SUFFIX"
                                    strSuffix = Fix_Suffix(Fld.Value)
                                Case "TITLE"
                                    strTitle = Fix_Title(Fld.Value)
                                Case "PROVIDER_ID"
                                    lngProv = Fld.Value
                            End Select
                        End If
                    End If
                Next
            End If
            .MoveNext
            If lngProv > 0 Then
                strLast = Fix_Name(strLast)
                If strSuffix <> vbNullString Then
                    strSuffix = Fix_Suffix(strSuffix)
                    strLast = strLast & ", " & strSuffix
                End If
                If Trim(strFirst) <> vbNullString Then strFirst = Fix_Name(strFirst)
                If Trim(strLast) <> vbNullString Then
                    lstProviders.AddItem strLast & " " & strFirst & " " & strMiddle & " " & strTitle
                    lstProviders.ItemData(lstProviders.NewIndex) = lngProv
                End If
            End If
        Loop
    End With
    If cnProv.State = adStateOpen Then
        cnProv.Close
        Set cnProv = Nothing
        Set Flds = Nothing
        Set rsProv = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_List_Providers:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "List_Providers", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Format_CellTypeCombo(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, _
vstrString As String, vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional vlngBackColor As Long)
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
              .TypeComboBoxCurSel = -1
              If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
              If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
              .FontName = "Times New Roman"
              .FontSize = 10
              .FontBold = True
              If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
              ElseIf Trim(vstrSpan) <> vbNullString Then
                lngSpan = vstrSpan
                If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
              End If
              If vlngBackColor > 0 Then .BackColor = vlngBackColor
            .TypeVAlign = TypeVAlignCenter
          .BlockMode = False
    End With
End Function

Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
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
              .FontName = "Times New Roman"
              .FontSize = 10
              .FontBold = True
              'If vbooUnderLine Then .FontUnderline = True
              If vlngBackColor > 0 Then .BackColor = vlngBackColor
              If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
              If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
              .TypeVAlign = TypeVAlignCenter
                If vstrHAligh = "TypeHAlignCenter" Then
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .TypeHAlign = TypeHAlignLeft
                End If
              If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
              ElseIf Trim(vstrSpan) <> vbNullString Then
                lngSpan = vstrSpan
                If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
              End If
              .SetText vlngCol, vlngRow, Trim(vstrText)
              .Text = Trim(vstrText)
              CellNote = .MaxTextCellHeight
              If .RowHeight(vlngRow) < CellNote Then
                .RowHeight(vlngRow) = CellNote
              End If
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypeEdit_RowSpan(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vbooUnderLine As Boolean, Optional vlngBackColor As Long, Optional vstrCellNote As String)
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
              .FontName = "Times New Roman"
              .FontSize = 10
              .FontBold = True
              If vbooUnderLine Then .FontUnderline = True
              If vlngBackColor > 0 Then .BackColor = vlngBackColor
              If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
              If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
              .TypeVAlign = TypeVAlignTop
                If vstrHAligh = "TypeHAlignCenter" Then
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .TypeHAlign = TypeHAlignLeft
                End If
              If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, vlngCol, .MaxRows
              ElseIf Trim(vstrSpan) <> vbNullString Then
                lngSpan = vstrSpan
                If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, vlngCol, lngSpan
              End If
              .SetText vlngCol, vlngRow, Trim(vstrText)
              .Text = Trim(vstrText)
              CellNote = .MaxTextCellHeight
              If .RowHeight(vlngRow) < CellNote Then
                .RowHeight(vlngRow) = CellNote
              End If
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypePicture(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long)
Dim CellNote As Long
Dim lngSpan As Long
Dim fso As New FileSystemObject

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            If vlngRowHeight > 0 Then .RowHeight(vlngRow) = vlngRowHeight
            .CellType = CellTypePicture
            .TypePictMaintainScale = True
            .TypePictStretch = True
            .TypePictCenter = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            Else
               If lngSpan > 0 And vstrSpan <> vbNullString Then
                    lngSpan = Val(vstrSpan)
                    .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                End If
             End If
            If fso.FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, vstrCellTag As String, Optional lngBackColor As Long)
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
            .TypeVAlign = TypeVAlignCenter
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            ElseIf vstrHAligh = "TypeHAlignRight" Then
                .TypeHAlign = TypeHAlignRight
            Else
                .TypeHAlign = TypeHAlignLeft
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




Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
    With grdTerms
        cctSearch.Text = vbNullString
        If Row >= 0 Then
            .BlockMode = True
                .Col = Col
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                 .BackColor = &HFFC0FF
             .BlockMode = False
        End If
    End With
    Exit Sub
End Sub


Private Sub grdTerms_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    With grdTerms
        cctSearch.Text = vbNullString
        .BlockMode = True
            .Col = Col
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
             .BackColor = &H80000005
         .BlockMode = False
    End With
    Exit Sub
End Sub


Private Sub lstProviders_ItemCheck(Item As Integer)
    If mbooByPass Then mbooByPass = False: Exit Sub
    If lstProviders.Selected(Item) Then
        piProviderID = lstProviders.ItemData(Item)
    End If
    Read_DefaultOrder
    Exit Sub
End Sub


