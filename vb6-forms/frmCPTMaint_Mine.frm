VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmCPTMaint 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CPT Maintanence"
   ClientHeight    =   8580
   ClientLeft      =   5670
   ClientTop       =   3285
   ClientWidth     =   13005
   Icon            =   "frmCPTMaint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   13005
   Begin VB.PictureBox picLogo 
      BackColor       =   &H00FFFFFF&
      Height          =   7395
      Left            =   30
      ScaleHeight     =   7335
      ScaleWidth      =   12855
      TabIndex        =   2
      Top             =   510
      Width           =   12915
      Begin VB.PictureBox picPatientTrac 
         Height          =   6075
         Left            =   2655
         Picture         =   "frmCPTMaint.frx":058A
         ScaleHeight     =   6015
         ScaleWidth      =   7635
         TabIndex        =   3
         Top             =   840
         Width           =   7695
      End
      Begin VB.Label lblPatientTrac 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "PATIENTTRAC CPT CODE MAINTANENCE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   150
         TabIndex        =   4
         Top             =   210
         Width           =   12705
      End
   End
   Begin VB.Frame fraGridTerm 
      BackColor       =   &H00E0E0E0&
      Height          =   3705
      Left            =   1620
      TabIndex        =   6
      ToolTipText     =   "Frame is movable with mouse!"
      Top             =   1410
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
         TabIndex        =   7
         ToolTipText     =   "Enter either Name or Code"
         Top             =   390
         Visible         =   0   'False
         Width           =   3945
      End
      Begin PT_XP_Button.PT_XP cmdGridTerms 
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         ToolTipText     =   "Terms Maintanence"
         Top             =   3300
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
         Picture         =   "frmCPTMaint.frx":E63C
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
         PictureDisabled =   "frmCPTMaint.frx":EBD6
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdDone 
         Height          =   315
         Left            =   2100
         TabIndex        =   9
         Top             =   3300
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
         Picture         =   "frmCPTMaint.frx":EF28
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Done"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmCPTMaint.frx":F2C2
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdTerms 
         Height          =   2355
         Left            =   90
         TabIndex        =   10
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
         SpreadDesigner  =   "frmCPTMaint.frx":F614
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
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   2115
      End
   End
   Begin PT_XP_Button.PT_XP cmdAddPreferred 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
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
      Picture         =   "frmCPTMaint.frx":F8EA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmCPTMaint.frx":FE84
      PictureAlignment=   0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRemove 
      Height          =   375
      Left            =   1380
      TabIndex        =   1
      Top             =   60
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
      Picture         =   "frmCPTMaint.frx":101D6
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Remove"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmCPTMaint.frx":10770
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdPrefCPT 
      Height          =   7455
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   12975
      _Version        =   458752
      _ExtentX        =   22886
      _ExtentY        =   13150
      _StockProps     =   64
      AutoCalc        =   0   'False
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
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
      GridColor       =   8388608
      MaxCols         =   10
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarMaxAlign=   0   'False
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmCPTMaint.frx":10AC2
      UserResize      =   0
      Appearance      =   2
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   405
      Left            =   5355
      TabIndex        =   12
      Top             =   8040
      Width           =   1065
      _ExtentX        =   1879
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
      Picture         =   "frmCPTMaint.frx":10D89
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmCPTMaint.frx":11123
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   405
      Left            =   6585
      TabIndex        =   13
      Top             =   8040
      Width           =   1065
      _ExtentX        =   1879
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
      Picture         =   "frmCPTMaint.frx":11475
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmCPTMaint.frx":11E87
      ShowFocus       =   -1  'True
   End
   Begin VB.Menu mnu_CPT_Maint 
      Caption         =   "CPT Type"
      Begin VB.Menu mnu_Labs 
         Caption         =   "Laboratory - Radiology"
      End
      Begin VB.Menu mnu_EM 
         Caption         =   "Evaluataion & Management"
      End
      Begin VB.Menu mnu_ProcTest 
         Caption         =   "Procedures & Tests"
      End
   End
   Begin VB.Menu mnu_Print 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "frmCPTMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrConnectStringClinical As String
Public mstrConnectStringTerms As String
Public mlngPatientID As Long
Public mlngEncounterID As Long
Public mlngProviderID As Long
Private mbooHasChanged As Boolean
Private mstrRemarks As String
Private mbooByPass As Boolean
Private mParentRef As String
Private gbooBypass As Boolean
Private mstrRefData As String
Private mlngGridRow As Long
Private mlngGridCol As Long
Private mstrList As String
Private mstrColType As String
Private mstrSignatureToken As String
Private mlngInsertRows As Long
Private mFirst As Boolean
Public mstrSubGroup As String
Private mbooPrintPreview As Boolean
Private mlngStartRow As Long
Private mstrLabCode As String
Private flngDeltaX As Long
Private flngDeltaY As Long
Private mbooNoSave As Boolean
Private fpDeltaLeft As Long
Private fpDeltaTop As Long
Private mbooCancelPrint As Boolean
Private Function Add_Template_Row()
Dim lngRow As Long
Dim lngCol As Long

    
    lngRow = grdPrefCPT.MaxRows + 1
    grdPrefCPT.MaxRows = lngRow
    With grdPrefCPT
        lngCol = 1
        Format_CellTypeStatic grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CPT_SUBGROUP"
        .SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeStatic grdPrefCPT, lngRow, lngCol, "2", " ", "TypeHAlignCenter", "CPT_NAME"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeStatic grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CPT_CODE"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeStatic grdPrefCPT, lngRow, lngCol, "2", " ", "TypeHAlignCenter", "ENCNTR_TYPE"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeStatic grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "PLACE_SERVICE"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        Format_CellTypeStatic grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "PROVIDER"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CPT_FEE"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 8
        Format_CellTypeStatic grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "REPORT_ID"
        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    End With
End Function

Private Function Format_CellTypeBackColor(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, lngBackColor As Long, _
Optional vstrCellTag As String, Optional vstrCellNote As String)
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
              .SetText vlngCol, vlngRow, vstrText
              .TypeEditMultiLine = True   'Make multi-lines
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              .Text = vstrText
              If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
              If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
              CellNote = .MaxTextCellHeight
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
              If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
              .TypeVAlign = TypeVAlignCenter
                If vstrHAligh = "TypeHAlignCenter" Then
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .TypeHAlign = TypeHAlignLeft
                End If
                .BackColor = lngBackColor
          .BlockMode = False

    End With
End Function
Private Function Format_CellTypeCombo(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, _
vstrString As String, vstrHAligh As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .ReDraw = False
        .BlockMode = True
              .Col = vlngCol
              .Col2 = vlngCol
              .Row = vlngRow
              .Row2 = vlngRow
              .CellType = CellTypeComboBox
              .TypeMaxEditLen = 32000
              .TypeComboBoxList = vstrString
              .TypeComboBoxWidth = 1
              If vlngBackColor > 0 Then
                 .BackColor = vlngBackColor
              End If
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
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
          .BlockMode = False
        .ReDraw = True
    End With
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional vlngBackColor As Long, Optional vlngFontSize As Long)
Dim CellNote As Long
Dim lngSpan As Long

        
    With vGrid
        .ReDraw = False
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeEdit
            .TypeMaxEditLen = 32000
            .FontBold = True
            If vlngFontSize > 0 Then .FontSize = vlngFontSize
            .TypeEditMultiLine = True   'Make multi-lines
            'If vbooUnderLine Then .FontUnderline = Truer
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
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote + 5
            .TypeVAlign = TypeVAlignCenter
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If vlngBackColor Then .BackColor = vlngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            .EditMode = False
          .BlockMode = False
        .ReDraw = True
    End With
End Function
Private Function Format_CellTypePicture(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long
    With vGrid
        .ReDraw = False
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            If vlngRowHeight > 0 Then .RowHeight(vlngRow) = vlngRowHeight
            .CellType = CellTypePicture
            lngSpan = Val(vstrSpan)
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf lngSpan > 1 Then
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
             End If
            If FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
            End If
             If Trim(vstrCellNote) <> vbNullString Then .CellNote = vstrCellNote
            .TypePictMaintainScale = True
            .TypePictStretch = True
            .TypePictCenter = True
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
        .ReDraw = True
    End With
End Function
Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional lngBackColor As Long, Optional vstrVAlign As String)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .ReDraw = False
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
        .ReDraw = True
    End With
End Function
Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)

    With vObjGrid
        vlngRow = vlngRow
       If .MaxRows < vlngRow Then .MaxRows = vlngRow
        .BlockMode = True
            .Col = 1
            .Col2 = 1
            If vlngRow = -998 Then
                .Row = SpreadHeader + 2
                .Row2 = SpreadHeader + 2
            Else
                .Row = vlngRow
                .Row2 = vlngRow
            End If
            .RowHeight(vlngRow) = 3
            .BackColor = GridBurg
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 1, vlngRow, .MaxCols, vlngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub
Private Function Format_GridTerms(vstrType As String)

       Screen.MousePointer = vbHourglass
    With grdTerms
        .ReDraw = False
       .ClearRange 0, SpreadHeader, grdTerms.MaxRows, grdTerms.MaxCols, False
        .ProcessTab = True
        .FontBold = True
        .ColHeaderRows = 2
        .ColHeadersShow = True
        .RowHeadersShow = False
        .ColWidth(1) = 28
        .ColWidth(2) = 10
        Select Case vstrType
            Case "ICD"
                cctSearch.Visible = True
                Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "DIAGNOSIS", "TypeHAlignCenter", "", "", HeaderColor
                Format_CellTypeEdit grdTerms, SpreadHeader, 2, "maxcols", "DX CODE", "TypeHAlignCenter", "", "", HeaderColor
                grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Case "MODIFIER"
                Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "MODIFIER", "TypeHAlignCenter", "", "", HeaderColor
                Format_CellTypeEdit grdTerms, SpreadHeader, 2, "maxcols", "CPT CODE", "TypeHAlignCenter", "", "", HeaderColor
                grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Case "POS"
                Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "PLACE SERVICE", "TypeHAlignCenter", "", "", HeaderColor
                Format_CellTypeEdit grdTerms, SpreadHeader, 2, "maxcols", "POS CODE", "TypeHAlignCenter", "", "", HeaderColor
                grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Case "CPT"
                cctSearch.Visible = True
                Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "CPT DESC", "TypeHAlignCenter", "", "", HeaderColor
                Format_CellTypeEdit grdTerms, SpreadHeader, 2, "maxcols", "CPT CODE", "TypeHAlignCenter", "", "", HeaderColor
                grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Case "CPT_INTRVN"
                cctSearch.Visible = True
                Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "CPT DESC", "TypeHAlignCenter", "", "", HeaderColor
                Format_CellTypeEdit grdTerms, SpreadHeader, 2, "maxcols", "CPT CODE", "TypeHAlignCenter", "", "", HeaderColor
                grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Case "REPORT_ID"
                cctSearch.Visible = True
                Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "REPORT NAME", "TypeHAlignCenter", "", "", HeaderColor
                Format_CellTypeEdit grdTerms, SpreadHeader, 2, "maxcols", "REPORT ID", "TypeHAlignCenter", "", "", HeaderColor
                grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End Select
        Add_Spacer SpreadHeader + 1, grdTerms
        grdTerms.MaxRows = 0
        .ReDraw = True
        mstrRefData = vstrType
    End With
End Function
Private Function Get_Provider_Title()
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection

    On Error GoTo Error_Get_GridTerms
    Screen.MousePointer = vbHourglass
    
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(0)
    mstrList = vbNullString
    strSQL = "SELECT DISTINCT TITLE FROM PROVIDERS WHERE NPI_ID <> 11"
    Set rsTerms = cnTerms.Execute(strSQL, adAsyncExecute)
    With rsTerms
        Do Until .EOF
            If (!Title) <> vbNullString Then
                If Trim(mstrList) = vbNullString Then
                    mstrList = Trim(!Title)
                Else
                    mstrList = mstrList & Chr(9) & Trim(!Title)
                End If
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
Error_Get_GridTerms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", " Get_GridTerms", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Get_VisitType() As String
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset

    On Error GoTo Error_Get_InvtrvTemp
    Screen.MousePointer = vbHourglass
    Get_VisitType = vbNullString
     If cnCPT.State = adStateClosed Then cnCPT.open psConnect(0)
     strSQL = "Select DISTINCT VISIT_TYPE FROM PLACE_SERVICE "
       Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
       With rsCPT
            Do Until .EOF
                    If (!VISIT_TYPE) <> vbNullString Then
                        If Get_VisitType = vbNullString Then
                            Get_VisitType = (!VISIT_TYPE)
                        Else
                            Get_VisitType = Trim(Get_VisitType) & Chr(9) & (!VISIT_TYPE)
                        End If
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
Error_Get_InvtrvTemp:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_InvtrvTemp", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Get_InvtrvTemp() As String
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim strIntrvn As String

    On Error GoTo Error_Get_InvtrvTemp
    Screen.MousePointer = vbHourglass
    Get_InvtrvTemp = vbNullString
    If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
    strSQL = "Select * FROM INTRVN_SUBGROUPS ORDER BY  SUBGROUP_ID"
       Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
       With rsCPT
            Do Until .EOF
                    If (!SUBGROUP_DESC) <> vbNullString Then
                       strIntrvn = (!SUBGROUP_DESC)
                       If strIntrvn = "Lab" Then strIntrvn = "Laboratory"
                       If strIntrvn = "Rad" Then strIntrvn = "Radiology"
                       If strIntrvn = "Procedure" Then strIntrvn = "Procedure"
                       If strIntrvn = "Eval/Mgmt" Then strIntrvn = "Eval/Mgmt"
                       If strIntrvn = "Psych Test" Then strIntrvn = "Psych Test"
                       If strIntrvn = "Injectables" Then strIntrvn = "Injectables"
                       If strIntrvn = "Miscellaneous" Then strIntrvn = "Miscellaneous"
                        If Get_InvtrvTemp = vbNullString Then
                            Get_InvtrvTemp = (strIntrvn)
                        Else
                            Get_InvtrvTemp = Trim(Get_InvtrvTemp) & Chr(9) & (strIntrvn)
                        End If
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
Error_Get_InvtrvTemp:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_InvtrvTemp", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Get_EnctrType() As String
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset

    On Error GoTo Error_Get_InvtrvTemp
    Screen.MousePointer = vbHourglass
    Get_EnctrType = vbNullString
     If cnCPT.State = adStateClosed Then cnCPT.open psConnect(0)
     strSQL = "Select DISTINCT ENCNTR_TYPE FROM PLACE_SERVICE "
       Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
       With rsCPT
            Do Until .EOF
                    If (!Encntr_Type) <> vbNullString Then
                        If Get_EnctrType = vbNullString Then
                            Get_EnctrType = (!Encntr_Type)
                        Else
                            Get_EnctrType = Get_EnctrType & Chr(9) & (!Encntr_Type)
                        End If
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
Error_Get_InvtrvTemp:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_InvtrvTemp", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Get_InvtrvGrpSubGp(vlngGrp_ID As Long, vlngSubGrp_Id As Long)
Dim strSQL As String
Dim cnInvtr As New ADODB.Connection
Dim rsInvtr As New ADODB.Recordset

    On Error GoTo Error_Get_InvtrvTemp
    Screen.MousePointer = vbHourglass
    If cnInvtr.State = adStateClosed Then cnInvtr.open psConnect(1)
    strSQL = "Select GRP_SUBGROUP_ID FROM INTRVN_GRP_SUBGROUP WHERE GROUP_ID = " & vlngGrp_ID & " AND SUBGROUP_ID = " & vlngSubGrp_Id
    Set rsInvtr = cnInvtr.Execute(strSQL, adAsyncExecute)
    With rsInvtr
         If Not .EOF Then
            If (!Grp_SubGroup_ID) > 0 Then clsCptIntrvn.Grp_SubGroup_ID = (!Grp_SubGroup_ID)
          End If
     End With
    If cnInvtr.State = adStateOpen Then
        cnInvtr.Close
        Set cnInvtr = Nothing
        Set rsInvtr = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_InvtrvTemp:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_InvtrvTemp", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Get_SubGroupID(vstrSubGrp As String)
Dim strSQL As String
Dim cnInvtr As New ADODB.Connection
Dim rsInvtr As New ADODB.Recordset

    On Error GoTo Error_Get_SubGroupID
    Screen.MousePointer = vbHourglass
    If cnInvtr.State = adStateClosed Then cnInvtr.open psConnect(1)
    strSQL = "Select SUBGROUP_ID FROM INTRVN_SUBGROUPS WHERE SUBGROUP_DESC = '" & vstrSubGrp & "'"
    Set rsInvtr = cnInvtr.Execute(strSQL, adAsyncExecute)
    With rsInvtr
         If Not .EOF Then
            clsCptIntrvn.Grp_SubGroup_ID = (!SUBGROUP_ID)
          End If
     End With
    If cnInvtr.State = adStateOpen Then
        cnInvtr.Close
        Set cnInvtr = Nothing
        Set rsInvtr = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_SubGroupID:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_SubGroupID", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function GetTerms_Visible(vObjGrid As Object, vlngCol As Long, vlngRow As Long, Optional vstrTag As String, Optional vlngResize As Long)
Dim lngx As Long
Dim lngY As Long
Dim lngWidth As Long
Dim lngHeight As Long

    With vObjGrid
        .Col = vlngCol
        .Row = vlngRow
        .GetCellPos .Col, .Row, lngx, lngY, lngWidth, lngHeight
        fpDeltaTop = lngY + 51.5
        fpDeltaLeft = lngx + 50
        fraGridTerm.ZOrder 0
        fraGridTerm.Top = .Top + fpDeltaTop
        fraGridTerm.Left = .Left + fpDeltaLeft
      '  lblSearch.Visible = False
       ' cctSearch.Visible = False
        fraGridTerm.Visible = True
    End With
End Function
Private Function Load_GridBox(vstrType As String)
Dim strSQL As String
Dim iRow As Long
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim iCnt As Integer
Dim strData As String
Dim lngRow As Long
Dim Arr
Dim booData As Boolean
Dim strExistPOS As String

    On Error GoTo Error_Load_GridBox
    Screen.MousePointer = vbHourglass
    If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
    Select Case Trim(vstrType)
        Case "CPT"
            Format_GridTerms vstrType
            If Trim(cctSearch.Text) <> vbNullString Then
                If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
                    If Trim(cctSearch.Text) Like "[A-Za-z]###" Or IsNumeric(cctSearch.Text) Then
                        grdTerms.ReDraw = False
                        strSQL = "SELECT CPT_DESC, CPT_CODE FROM CPT_CODE WHERE CPT_CODE like '%" & Trim(cctSearch.Text) & "%'  ORDER BY CPT_CODE"
                        Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
                        With rsCPT
                            Do Until .EOF
                                If Not IsNull(!CPT_Desc) Then strData = (!CPT_Desc)
                                If Trim(strData) <> vbNullString Then
                                    booData = True
                                    lngRow = lngRow + 1
                                    grdTerms.MaxRows = lngRow
                                    Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "CPT NAME"
                                    If Not IsNull(!CPT_Code) Then
                                        If IsNumeric(!CPT_Code) Then grdTerms.SetRowItemData (lngRow), (!CPT_Code)
                                        If Not IsNull((!CPT_Code)) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!CPT_Code), "TypeHAlignCenter", "CPT CODE"
                                        grdTerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                        grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    End If
                                End If
                                .MoveNext
                            Loop
                        End With
                    Else
                        strSQL = "select CPT_DESC, CPT_CODE from CPT_CODE Where CPT_DESC like  '%" & Trim(cctSearch.Text) & "%' ORDER BY CPT_DESC "
                        Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
                        With rsCPT
                            Do Until .EOF
                                If (!CPT_Desc) <> vbNullString Then strData = (!CPT_Desc)
                                If Trim(strData) <> vbNullString Then
                                    booData = True
                                    lngRow = lngRow + 1
                                    grdTerms.MaxRows = lngRow
                                    Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "CPT NAME"
                                    If Not IsNull(!CPT_Code) Then
                                        If IsNumeric(!CPT_Code) Then grdTerms.SetRowItemData (lngRow), Trim(!CPT_Code)
                                        If Not IsNull((!CPT_Code)) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", Trim(!CPT_Code), "TypeHAlignCenter", "CPT CODE"
                                        grdTerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                        grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    End If
                                End If
                             .MoveNext
                            Loop
                        End With
                    End If
            End If
            If Trim(strData) = vbNullString Then
                Format_GridTerms vstrType
                LookUp_CPT
            End If
        Case "ICD", "SUG_ICD_CODE", "ICD CODE"
            Format_GridTerms "ICD"
            If Trim(cctSearch.Text) <> vbNullString Then
                If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
                    If Trim(cctSearch.Text) Like "[A-Za-z]###" Or IsNumeric(cctSearch.Text) Then
                        grdTerms.ReDraw = False
                        strSQL = "SELECT LONG_DESC, CODE FROM ICD_CODE WHERE CODE like '" & Trim(cctSearch.Text) & "%' ORDER BY CODE"
                        Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
                        With rsCPT
                            Do Until .EOF
                                If Not IsNull(!Long_Desc) Then strData = (!Long_Desc)
                                If Trim(strData) <> vbNullString Then
                                    booData = True
                                    lngRow = lngRow + 1
                                    grdTerms.MaxRows = lngRow
                                    Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "ICD"
                                    If Not IsNull(!code) Then
                                        If IsNumeric(!code) Then grdTerms.SetRowItemData (lngRow), (!code)
                                        If Not IsNull((!code)) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!code), "TypeHAlignCenter", "ICD"
                                        grdTerms.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                        grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    End If
                                End If
                                .MoveNext
                            Loop
                        End With
                    Else
                        strSQL = "select LONG_DESC, CODE from icd_code Where LONG_DESC like  '%" & Trim(cctSearch.Text) & "%' ORDER BY LONG_DESC "
                        Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
                        With rsCPT
                            Do Until .EOF
                                If (!Long_Desc) <> vbNullString Then strData = (!Long_Desc)
                                If Trim(strData) <> vbNullString Then
                                    booData = True
                                    lngRow = lngRow + 1
                                    grdTerms.MaxRows = lngRow
                                    Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "ICD"
                                    If Not IsNull(!code) Then
                                        If IsNumeric(!code) Then grdTerms.SetRowItemData (lngRow), (!code)
                                        If Not IsNull((!code)) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!code), "TypeHAlignCenter", "ICD"
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
            If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
            strSQL = "select DISTINCT CPT_MOD_DESC, CPT_MOD_CODE " _
                & "From CPT_MOD_TEMPLATE  "
            Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
            With rsCPT
                Do Until .EOF
                    If Not IsNull(!CPT_MOD_DESC) Then strData = (!CPT_MOD_DESC)
                    If Trim(strData) <> vbNullString Then
                        lngRow = lngRow + 1
                        grdTerms.MaxRows = lngRow
                        Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "MODIFIER"
                        If IsNumeric(!CPT_MOD_CODE) Then grdTerms.SetRowItemData (lngRow), (!CPT_MOD_CODE)
                        If Not IsNull(!CPT_MOD_CODE) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!CPT_MOD_CODE), "TypeHAlignLeft", "MODIFIER"
                    End If
                    .MoveNext
                Loop
            End With
        Case "CPT_INTRVN", "CPT_INTRVN_CODE", "CPT_INTRVN_DESC"
            grdTerms.ReDraw = False
            Format_GridTerms "CPT_INTRVN"
            If Trim(cctSearch.Text) <> vbNullString Then
                If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
                strSQL = "SELECT   CPT_INTRVN_NAME, CPT_INTRVN_CODE  FROM INTRVN_TEMPLATE  WHERE CPT_INTRVN_NAME like '" & cctSearch.Text & "%'"
                Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
                With rsCPT
                    Do Until .EOF
                        If Not IsNull(!CPT_Intrvn_Name) Then strData = (!CPT_Intrvn_Name)
                        If Trim(strData) <> vbNullString Then
                            lngRow = lngRow + 1
                            grdTerms.MaxRows = lngRow
                            Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "CPT_INTRVN"
                            If IsNumeric(!CPT_Intrvn_Code) Then grdTerms.SetRowItemData lngRow, (!CPT_Intrvn_Code)
                            If Not IsNull(!CPT_Intrvn_Code) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!CPT_Intrvn_Code), "TypeHAlignLeft", "CPT_INTRVN"
                            grdTerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                        .MoveNext
                    Loop
                End With
                If grdTerms.MaxRows < 1 Then
                    strSQL = "SELECT  CPT_INTRVN_NAME, CPT_INTRVN_CODE  FROM INTRVN_TEMPLATE  WHERE CPT_INTRVN_CODE like '" & cctSearch.Text & "%'"
                    Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
                    With rsCPT
                        Do Until .EOF
                            If Not IsNull(!CPT_Intrvn_Name) Then strData = (!CPT_Intrvn_Name)
                            If Trim(strData) <> vbNullString Then
                                lngRow = lngRow + 1
                                grdTerms.MaxRows = lngRow
                                Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "CPT_INTRVN"
                                If IsNumeric(!CPT_Intrvn_Code) Then grdTerms.SetRowItemData (lngRow), (!CPT_Intrvn_Code)
                                If Not IsNull(!CPT_Intrvn_Code) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!CPT_Intrvn_Code), "TypeHAlignLeft", "CPT_INTRVN"
                                grdTerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                            .MoveNext
                        Loop
                    End With
                End If
            End If
        Case "POS"
            grdTerms.ReDraw = False
            Format_GridTerms "POS"
            cnCPT.Close
            If cnCPT.State = adStateClosed Then cnCPT.open psConnect(0)
            strSQL = "SELECT * FROM PLACE_SERVICE ORDER BY POS_NAME"
            Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
            With rsCPT
                Do Until .EOF
                    If Not IsNull(!POS_NAME) Then strData = (!POS_NAME)
                    If Trim(strData) <> vbNullString And strData <> Trim(strExistPOS) Then
                        lngRow = lngRow + 1
                        grdTerms.MaxRows = lngRow
                        Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "POS"
                        If IsNumeric(!POS_CODE) Then grdTerms.SetRowItemData (lngRow), (!POS_CODE)
                        If Not IsNull(!POS_CODE) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!POS_CODE), "TypeHAlignLeft", "POS"
                        grdTerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        strExistPOS = strData
                    End If
                    .MoveNext
                Loop
            End With
        Case "REPORT_ID"
            grdTerms.ReDraw = False
            Format_GridTerms "REPORT_ID"
            cnCPT.Close
            If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
            strSQL = "SELECT * FROM PRACTICE_REPORTS ORDER BY REPORT_NAME"
            Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
            With rsCPT
                Do Until .EOF
                    If Not IsNull(!REPORT_NAME) Then strData = (!REPORT_NAME)
                    If Trim(strData) <> vbNullString And strData <> Trim(strExistPOS) Then
                        lngRow = lngRow + 1
                        grdTerms.MaxRows = lngRow
                        Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "REPORT_NAME"
                        If IsNumeric(!REPORT_ID) Then grdTerms.SetRowItemData lngRow, (!REPORT_ID)
                        If Not IsNull(!REPORT_ID) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!REPORT_ID), "TypeHAlignLeft", "REPORT_ID"
                        grdTerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        strExistPOS = strData
                    End If
                    .MoveNext
                Loop
            End With
    End Select
    grdTerms.ReDraw = True
    If grdTerms.MaxRows <= 0 And vstrType <> "ICD" Then
         LookUp_CPT
    End If
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_GridBox:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_CPT", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function LookUp_CPT()
Dim strSQL As String
Dim iRow As Long
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim iCnt As Integer
Dim strData As String
Dim lngRow As Long
Dim Arr
Dim booData As Boolean
Dim strExistPOS As String

    On Error GoTo Error_LookUp_CPT
    
    If Trim(cctSearch.Text) = vbNullString Then Exit Function
    Screen.MousePointer = vbHourglass
    If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
    If Trim(cctSearch.Text) <> vbNullString Then
        If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
        grdTerms.ReDraw = False
        strSQL = "SELECT CPT_DESC, CPT_CODE FROM CPT_CODE WHERE CPT_CODE like '%" & Trim(cctSearch.Text) & "%'  ORDER BY CPT_CODE"
        Set rsCPT = cnCPT.Execute(strSQL)
        With rsCPT
            Do Until .EOF
                If Not IsNull(!CPT_Desc) Then strData = (!CPT_Desc)
                If Trim(strData) <> vbNullString Then
                    booData = True
                    lngRow = lngRow + 1
                    grdTerms.MaxRows = lngRow
                    Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "CPT NAME"
                    If Not IsNull(!CPT_Code) Then
                        If IsNumeric(!CPT_Code) Then grdTerms.SetRowItemData (lngRow), (!CPT_Code)
                        If Not IsNull((!CPT_Code)) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!CPT_Code), "TypeHAlignCenter", "CPT CODE"
                        grdTerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                End If
                .MoveNext
            Loop
        End With
    Else
        strSQL = "select CPT_DESC, CPT_CODE from CPT_CODE Where CPT_DESC like  '%" & Trim(cctSearch.Text) & "%' ORDER BY CPT_DESC "
        Set rsCPT = cnCPT.Execute(strSQL)
        With rsCPT
            Do Until .EOF
                If (!CPT_Desc) <> vbNullString Then strData = (!CPT_Desc)
                If Trim(strData) <> vbNullString Then
                    booData = True
                    lngRow = lngRow + 1
                    grdTerms.MaxRows = lngRow
                    Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "CPT NAME"
                    If Not IsNull(!CPT_Code) Then
                        If IsNumeric(!CPT_Code) Then grdTerms.SetRowItemData (lngRow), Trim(!CPT_Code)
                        If Not IsNull((!CPT_Code)) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", Trim(!CPT_Code), "TypeHAlignCenter", "CPT CODE"
                        grdTerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                End If
             .MoveNext
            Loop
        End With
    End If
    grdTerms.ReDraw = True
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUp_CPT:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUp_CPT", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
 
Private Function Write_Cpt_Template()
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim strTag As String
Dim sDate As String
Dim lngPrimKey As Long

        On Error GoTo Error_Write_Cpt_Template
        Screen.MousePointer = vbHourglass
        sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
        If cnCPT.State = adStateClosed Then cnCPT.open psConnect(0)
        If clsCptTemp.CPT_Subgroup = vbNullString Then clsCptTemp.CPT_Subgroup = "Office"
        If clsCptTemp.Place_Service_CD = vbNullString Then
            strSQL = "SELECT POS_CODE FROM PLACE_SERVICE WHERE POS_NAME = '" & clsCptTemp.CPT_Subgroup & "'"
            Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
            If Not rsCPT.EOF Then
                clsCptTemp.Place_Service_CD = rsCPT.Fields("POS_CODE")
            End If
        End If
        If cnCPT.State = adStateOpen Then cnCPT.Close
        If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
        strSQL = "SELECT CPT_ID FROM CPT_TEMPLATE WHERE CPT_CODE = '" & clsCptTemp.CPT_Code & "' AND " _
            & "CPT_NAME = '" & clsCptTemp.CPT_Name & "' AND  PLACE_SERVICE_CD = '" & clsCptTemp.Place_Service_CD & "'"
        Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
        If Not rsCPT.EOF Then
            lngPrimKey = rsCPT.Fields("CPT_ID")
        End If
        If lngPrimKey <> clsCptTemp.CPT_ID Then
            clsCptTemp.CPT_ID = lngPrimKey
        End If
        If clsCptTemp.CPT_ID > 0 Then
            strSQL = "UPDATE CPT_TEMPLATE " _
                    & "Set CPT_GROUP = 'Superbill', CPT_SUBGROUP = '" & clsCptTemp.CPT_Subgroup & "', " _
                    & "CPT_NAME = '" & clsCptTemp.CPT_Name & "',CPT_FEE = '" & clsCptTemp.CPT_Fee & "',TAXABLE = '" & clsCptTemp.Taxable & "'," _
                    & "SERVICE_PLACE = '" & clsCptTemp.Service_Place & "', PLACE_SERVICE_CD = '" & clsCptTemp.Place_Service_CD & "'," _
                    & "ENCNTR_TYPE = '" & clsCptTemp.Encntr_Type & "', PROVIDER = '" & clsCptTemp.Provider & "', " _
                    & "UPDATE_DATE = '" & sDate & "', CPT_CODE = '" & clsCptTemp.CPT_Code & "' WHERE  CPT_ID = " & clsCptTemp.CPT_ID
            Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
            If clsGridFormat.GridReportID > 0 Then
                strSQL = "UPDATE CPT_TEMPLATE  Set REPORT_ID =  '" & clsGridFormat.GridReportID & "' WHERE  CPT_ID = " & clsCptTemp.CPT_ID
                Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
            End If
        Else
            strSQL = "INSERT INTO CPT_TEMPLATE (CPT_GROUP, CPT_SUBGROUP, CPT_CODE, CPT_NAME, CPT_FEE, " _
                & "TAXABLE, SERVICE_PLACE,PLACE_SERVICE_CD, ENCNTR_TYPE, PROVIDER, REPORT_ID, INSERT_DATE) " _
                & "values ('Superbill','" _
                    & clsCptTemp.CPT_Subgroup & "','" & clsCptTemp.CPT_Code & "','" _
                    & clsCptTemp.CPT_Name & "','" _
                    & clsCptTemp.CPT_Fee & "','" _
                    & clsCptTemp.Taxable & "','" _
                    & clsCptTemp.Service_Place & "','" _
                    & clsCptTemp.Place_Service_CD & "','" _
                    & clsCptTemp.Encntr_Type & "','" _
                    & clsCptTemp.Provider & "','" _
                    & clsGridFormat.GridReportID & "', '" _
                    & sDate & "')"
                Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
        End If
        Clear_CPT_Temp
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Cpt_Template:
    Screen.MousePointer = vbDefault
  Write_Error Err.Description, Err.Number, strSQL, "Write_Cpt_Template", App.EXEName & "." & Me.Name
  Exit Function
  Resume
End Function
Private Function Write_Intrvn_Template()
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim strSQL As String
Dim booValid As Boolean

    On Error GoTo Error_Write_Intrvn_Template
    Screen.MousePointer = vbHourglass
    If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
    If clsCptIntrvn.Group_ID > 0 And clsCptIntrvn.SubGroupID > 0 Then
        Get_InvtrvGrpSubGp clsCptIntrvn.Group_ID, clsCptIntrvn.SubGroupID
    End If
    If clsCptIntrvn.Grp_SubGroup_ID = 0 Then
        If clsCptIntrvn.Group_ID > 0 And clsCptIntrvn.SubGroupID > 0 Then
            strSQL = "INSERT INTO INTRVN_GRP_SUBGROUP (GROUP_ID, SUBGROUP_ID) VALUES (" & clsCptIntrvn.Group_ID & "," & clsCptIntrvn.SubGroupID & ")"
            Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
            Get_InvtrvGrpSubGp clsCptIntrvn.Group_ID, clsCptIntrvn.SubGroupID
        End If
    End If
    strSQL = "SELECT INTRVN_ID FROM INTRVN_TEMPLATE WHERE CPT_INTRVN_NAME = '" & clsCptIntrvn.CPT_Intrvn_Desc & "'"
    Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
    If Not rsCPT.EOF Then
        clsCptIntrvn.Intrvn_ID = rsCPT.Fields("INTRVN_ID")
    End If
    If clsCptIntrvn.Intrvn_ID > 0 Then
          strSQL = "UPDATE INTRVN_TEMPLATE SET GROUP_ID = " & clsCptIntrvn.Group_ID & ", SUBGROUP_ID = " & clsCptIntrvn.SubGroupID & "," _
            & "GRP_SUBGROUP_ID = " & clsCptIntrvn.Grp_SubGroup_ID & ", CPT_INTRVN_CODE = '" & clsCptIntrvn.CPT_Intrvn_Code & "'," _
            & "CPT_INTRVN_DESC = '" & clsCptIntrvn.CPT_Intrvn_Desc & "', CPT_INTRVN_NAME = '" & clsCptIntrvn.CPT_Intrvn_Desc & "', CPT_FEE ='" & clsCptIntrvn.CPT_Fee & "'," _
            & "CPT_MODIFIER ='" & clsCptIntrvn.CPT_Modifier & "', PROVIDER = '" & clsCptIntrvn.Provider & "'," _
            & "SUG_ICD_CODE ='" & clsCptIntrvn.SUG_ICD_Code & "', SEQ_NO = " & Val(clsGridFormat.Row) & " WHERE  INTRVN_ID = " & clsCptIntrvn.Intrvn_ID
            Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
    Else
        strSQL = "INSERT INTO INTRVN_TEMPLATE (GROUP_ID, SUBGROUP_ID, GRP_SUBGROUP_ID, CPT_INTRVN_CODE, " _
            & "CPT_INTRVN_DESC, CPT_INTRVN_NAME, CPT_FEE, CPT_MODIFIER, PROVIDER, SUG_ICD_CODE, SEQ_NO) " _
            & "values(" & clsCptIntrvn.Group_ID & "," & clsCptIntrvn.SubGroupID & "," & clsCptIntrvn.Grp_SubGroup_ID & ",'" _
            & clsCptIntrvn.CPT_Intrvn_Code & "','" & clsCptIntrvn.CPT_Intrvn_Desc & "','" & clsCptIntrvn.CPT_Intrvn_Desc & "','" _
            & clsCptIntrvn.CPT_Fee & "','" & clsCptIntrvn.CPT_Modifier & "','" & clsCptIntrvn.Provider & "','" _
            & clsCptIntrvn.SUG_ICD_Code & "'," & Val(clsGridFormat.Row) & ")"
            Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
        strSQL = "SELECT INTRVN_ID FROM INTRVN_TEMPLATE WHERE CPT_INTRVN_NAME = '" & clsCptIntrvn.CPT_Intrvn_Desc & "'"
        Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
        If Not rsCPT.EOF Then
            clsCptIntrvn.Intrvn_ID = rsCPT.Fields("INTRVN_ID")
        End If
    End If
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Intrvn_Template:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Write_Intrvn_Template", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Load_CPTIntrvn(Optional vstrType As String)
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim iCnt As Integer
Dim strData As String
Dim lngRow As Long
Dim lngSubGrpID As Long
Dim lngCol As Long
Dim lngColWidth As Long
Dim strTag As String

    On Error GoTo Error_Load_CptIntrvn
    Screen.MousePointer = vbHourglass
    grdPrefCPT.ReDraw = False
     If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
     If vstrType = vbNullString Then
        strSQL = "SELECT DISTINCT INTRVN_TEMPLATE.*, INTRVN_SUBGROUPS.SUBGROUP_DESC " _
            & "FROM INTRVN_TEMPLATE   INNER JOIN INTRVN_SUBGROUPS   " _
            & "ON INTRVN_SUBGROUPS.SUBGROUP_ID = INTRVN_TEMPLATE.SUBGROUP_ID WHERE " _
            & "INTRVN_TEMPLATE.PURGE = 0 AND (INTRVN_TEMPLATE.SUBGROUP_ID = 9 Or INTRVN_TEMPLATE.SUBGROUP_ID = 12) ORDER BY INTRVN_TEMPLATE.CPT_INTRVN_DESC"
    ElseIf vstrType = "Procedure" Then
        strSQL = "SELECT DISTINCT INTRVN_TEMPLATE.*, INTRVN_SUBGROUPS.SUBGROUP_DESC  " _
            & "FROM INTRVN_TEMPLATE INNER JOIN INTRVN_SUBGROUPS " _
            & "ON INTRVN_TEMPLATE.SUBGROUP_ID = INTRVN_SUBGROUPS.SUBGROUP_ID WHERE " _
            & "INTRVN_TEMPLATE.PURGE = 0 AND (INTRVN_TEMPLATE.SUBGROUP_ID   <> 9 AND INTRVN_TEMPLATE.SUBGROUP_ID <> 12) ORDER " _
            & "BY INTRVN_TEMPLATE.CPT_INTRVN_DESC"
    End If
       Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
       With rsCPT
            Do Until .EOF
                lngRow = lngRow + 1
                grdPrefCPT.MaxRows = lngRow
                 lngCol = 1
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "TYPE"
                grdPrefCPT.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "2", " ", "TypeHAlignCenter", "CPT_INTRVN_DESC"
                grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 4
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CPT_INTRVN_CODE"
                grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 5
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CPT_MODIFIER"
                grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 6
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "SUG_ICD_CODE"
                grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 7
                Get_Provider_Title
                Format_CellTypeCombo grdPrefCPT, lngRow, lngCol, "1", mstrList, "TypeHAlignLeft", "PROVIDER"
                grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 8
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CPT_FEE"
                grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                '--------
                If (!SUBGROUP_ID) <> vbNullString Then
                   lngSubGrpID = (!SUBGROUP_ID)
                   strData = (!SUBGROUP_DESC)
                    If vstrType = vbNullString Then
                        If strData = "Lab" Then
                            strData = "Laboratory"
                        ElseIf strData = "Rad" Then
                            strData = "Radiology"
                        End If
                    Else
                        If strData = "Procedure" Then
                            strData = "Procedure"
                        ElseIf strData = "Psych Test" Then
                            strData = "Psych Test"
                        ElseIf strData = "Injectables" Then
                            strData = "Injectables"
                        ElseIf strData = "Miscellaneous" Then
                            strData = "Miscellaneous"
                        End If
                    End If
                    If (!Group_ID) = 2 Then
                        Format_CellTypeEdit grdPrefCPT, lngRow, 0, "1", "SUPERBILL", "TypeHAlignCenter"
                    End If
                End If
                Set Flds = rsCPT.Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "INTRVN_ID"
                                    grdPrefCPT.SetRowItemData (lngRow), Fld.Value
                                Case "SUBGROUP_ID"
                                    lngCol = 1
                                    Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", strData, "TypeHAlignCenter", CStr(lngSubGrpID)
                                Case "CPT_INTRVN_DESC"
                                    lngCol = 2
                                    Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "2", Fld.Value, "TypeHAlignLeft", strTag
                                Case "CPT_INTRVN_CODE"
                                    lngCol = 4
                                    Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", strTag
                                Case "CPT_MODIFIER"
                                    lngCol = 5
                                    Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", strTag
                                Case "SUG_ICD_CODE"
                                    lngCol = 6
                                    Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", strTag
                                Case "PROVIDER"
                                    lngCol = 7
                                    Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", strTag
                                Case "CPT_FEE"
                                    lngCol = 8
                                    Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", strTag
                            End Select
                        End If
                    End If
                Next
                .MoveNext
            Loop
        End With
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_CptIntrvn:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_CptIntrvn", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Write_Data()
Dim lngRow As Long
Dim lngCol As Long
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim strTag As String
Dim sDate As String
Dim lngPrimKey As Long
Dim lngSubGrp As Long
Dim varCellValue
Dim strType As String

    On Error GoTo Error_Write_Data
    If clsCptIntrvn.CPT_Type = vbNullString Then Exit Function
    
    strType = clsCptIntrvn.CPT_Type
    Screen.MousePointer = vbHourglass
    
    Clear_CPT_Temp
    Clear_Intrvn_CPT
     
    With grdPrefCPT
        For lngRow = 1 To .MaxRows
            lngPrimKey = 0
            lngPrimKey = .GetRowItemData(lngRow)
            Clear_CPT_Temp
            Clear_Intrvn_CPT
            clsGridFormat.GridReportID = 0
            clsCptIntrvn.Intrvn_ID = 0
            For lngCol = 1 To .MaxCols
                .GetText 1, lngRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    clsCptIntrvn.CPT_Type = Trim(varCellValue)
                End If
                strTag = vbNullString
                .Col = lngCol
                .Row = lngRow
                If clsCptIntrvn.CPT_Type = "Evaluation" Then
                    .GetText lngCol, lngRow, varCellValue
                    .Col = lngCol
                    .Row = lngRow
                    If Trim(varCellValue) <> vbNullString Then
                        If Trim(.CellTag) <> vbNullString Then
                            If strTag = "TYPE" Then strTag = "CPT_SUBGROUP"
                            strTag = Replace(.CellTag, " ", "_")
                        End If
                        Select Case strTag
                           Case "CPT_SUBGROUP"
                                clsCptTemp.CPT_Subgroup = Trim(varCellValue)
                           Case "CPT_INTRVN_CODE", "CPT_CODE"
                                clsCptTemp.CPT_Code = Trim(varCellValue)
                           Case "CPT_INTRVN_DESC", "CPT_NAME"
                                clsCptTemp.CPT_Name = Trim(varCellValue)
                           Case "CPT_FEE"
                                clsCptTemp.CPT_Fee = Trim(varCellValue)
                           Case "PLACE_SERVICE"
                                clsCptTemp.Service_Place = Trim(varCellValue)
                                If .CellNote <> vbNullString Then
                                    clsCptTemp.Place_Service_CD = .CellNote
                                End If
                           Case "ENCNTR_TYPE"
                                clsCptTemp.Encntr_Type = Trim(varCellValue)
                           Case "CPT_DESC"
                                clsCptTemp.CPT_Desc = Trim(varCellValue)
                           Case "PROVIDER"
                                clsCptTemp.Provider = Trim(varCellValue)
                           Case "REPORT_ID"
                                clsGridFormat.GridReportID = Trim(varCellValue)
                        End Select
                        If lngPrimKey > 0 Then
                            clsCptTemp.CPT_ID = lngPrimKey
                        End If
                    End If
                ElseIf clsCptIntrvn.CPT_Type <> vbNullString Then
                    .GetText lngCol, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        clsCptIntrvn.Group_ID = 2
                        If Trim(.CellTag) <> vbNullString Then
                            strTag = Replace(.CellTag, " ", "_")
                        End If
                        If lngCol = 1 Then
                            If IsNumeric(.CellTag) Then
                                clsCptIntrvn.SubGroupID = .CellTag
                            Else
                                Get_SubGroupID CStr(varCellValue)
                            End If
                        End If
                        Select Case strTag
                            Case "CPT_INTRVN_DESC", "CPT_NAME"
                                 clsCptIntrvn.CPT_Intrvn_Desc = Trim(varCellValue)
                            Case "CPT_INTRVN_CODE", "CPT_CODE"
                                 clsCptIntrvn.CPT_Intrvn_Code = Trim(varCellValue)
                            Case "CPT_FEE"
                                 clsCptIntrvn.CPT_Fee = Trim(varCellValue)
                            Case "CPT_MODIFIER"
                                 clsCptIntrvn.CPT_Modifier = Trim(varCellValue)
                            Case "PROVIDER"
                                 clsCptIntrvn.Provider = Trim(varCellValue)
                            Case "SUG_ICD_CODE"
                                 clsCptIntrvn.SUG_ICD_Code = Trim(varCellValue)
                        End Select
                        If lngPrimKey > 0 Then
                            clsCptIntrvn.Intrvn_ID = lngPrimKey
                        End If
                    End If
                End If
            Next lngCol
            clsGridFormat.Row = lngRow
            If clsCptIntrvn.CPT_Type = "Evaluation" Then
                Write_Cpt_Template
            Else
                Write_Intrvn_Template
                If clsCptIntrvn.Intrvn_ID > 0 Then grdPrefCPT.SetRowItemData lngRow, Val(clsCptIntrvn.Intrvn_ID)
            End If
            lngPrimKey = 0
        Next lngRow
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Data", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Sub cctSearch_KeyUp(KeyCode As Integer, Shift As Integer)

    If Len(cctSearch.Text) >= 2 Then
        Load_GridBox clsGridFormat.CellTag
'        Select Case clsGridFormat.CellTag
'            Case "ICD", "SUG_ICD_CODE", "ICD CODE"
'                Load_GridBox clsGridFormat.CellTag
'            Case "CPT", "CPT_INTRVN_CODE"
'                Load_GridBox clsGridFormat.CellTag
'            Case "CPT_INTRVN", "CPT_INTRVN_DESC"
'                Load_GridBox mstrRefData
'        End Select
    End If

    Exit Sub
End Sub
Private Sub cmdAddPreferred_Click()
Dim lngRow As Long
Dim lngCol As Long
Dim lngColWidth As Long


    If clsCptIntrvn.CPT_Type = vbNullString Then
        MsgBox "Please select CPT Type first!", vbInformation, "PatientTrac CPT Maintanence"
        Exit Sub
    End If
     With grdPrefCPT
        .ReDraw = False
        .MaxRows = .MaxRows + 1
        lngRow = .MaxRows
        Select Case clsCptIntrvn.CPT_Type
            Case "Evaluation"
                Add_Template_Row
            Case clsCptIntrvn.CPT_Type
                Format_CellTypeEdit grdPrefCPT, lngRow, 0, "1", "GROUP", "TypeHAlignCenter"
                 grdPrefCPT.ColHidden = True
                lngCol = 1
                lngColWidth = 12
                grdPrefCPT.ColWidth(lngCol) = lngColWidth
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "TYPE"
                .SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                lngColWidth = 18
                grdPrefCPT.ColWidth(lngCol) = lngColWidth
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "2", " ", "TypeHAlignCenter", "CPT_INTRVN_DESC"
                .SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 4
                lngColWidth = 10
                grdPrefCPT.ColWidth(lngCol) = lngColWidth
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CPT_INTRVN_CODE"
                .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 5
                lngColWidth = 10
                grdPrefCPT.ColWidth(lngCol) = lngColWidth
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CPT_MODIFIER"
                .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 6
                lngColWidth = 17.5
                grdPrefCPT.ColWidth(lngCol) = lngColWidth
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "SUG_ICD_CODE", "DEFAULT SELECTION"
                .SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 7
                lngColWidth = 10
                grdPrefCPT.ColWidth(lngCol) = lngColWidth
                Get_Provider_Title
                Format_CellTypeCombo grdPrefCPT, lngRow, lngCol, "1", mstrList, "TypeHAlignLeft", "PROVIDER"
                .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 8
                lngColWidth = 10
                grdPrefCPT.ColWidth(lngCol) = lngColWidth
                Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "CPT_FEE"
                .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End Select
        .ReDraw = True
    End With
    Exit Sub
End Sub
Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub
Private Sub cmdDone_Click()
Dim strData As String
Dim lngRow As Long
Dim varCellValue
Dim varCode
Dim strCode As String
Dim Arr
Dim lngCol As Long
Dim booCPT As Boolean

    With grdTerms
        .ReDraw = True
        fraGridTerm.Visible = False
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                Select Case mstrRefData
                    Case "POS"
                        .GetText 1, lngRow, varCellValue
                        .GetText 2, lngRow, varCode
                        If Trim(varCellValue) <> vbNullString Then strData = Trim(varCellValue)
                        If Trim(varCode) <> vbNullString Then strCode = Trim(varCode)
                    Case "REPORT_ID"
                        .GetText 1, lngRow, varCellValue
                        .GetText 2, lngRow, varCode
                        If Trim(varCellValue) <> vbNullString Then strData = Trim(varCellValue)
                        If Trim(varCode) <> vbNullString Then strCode = Trim(varCode)
                    Case "CPT"
                        .GetText 1, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            If strData = vbNullString Then
                                strData = Trim(varCellValue)
                            Else
                                strData = strData & vbNewLine & Trim(varCellValue)
                            End If
                        End If
                        .GetText 2, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            If strCode = vbNullString Then
                                strCode = Trim(varCellValue)
                            Else
                                strCode = strCode & vbNewLine & Trim(varCellValue)
                            End If
                        End If
                    Case "CPT_INTRVN"
                        .GetText 1, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            If strData = vbNullString Then
                                strData = Trim(varCellValue)
                            Else
                                strData = strData & vbNewLine & Trim(varCellValue)
                            End If
                        End If
                        .GetText 2, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            If strCode = vbNullString Then
                                strCode = Trim(varCellValue)
                            Else
                                strCode = strCode & vbNewLine & Trim(varCellValue)
                            End If
                        End If
                    Case "ICD"
                        .GetText 1, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            If strData = vbNullString Then
                                strData = Trim(varCellValue)
                            Else
                                strData = strData & vbNewLine & Trim(varCellValue)
                            End If
                        End If
                        .GetText 2, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            If strCode = vbNullString Then
                                strCode = Trim(varCellValue)
                            Else
                                strCode = strCode & vbNewLine & Trim(varCellValue)
                            End If
                        End If
                    Case "MODIFIER"
                        .GetText 1, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            If strData = vbNullString Then
                                strData = Trim(varCellValue)
                            Else
                                strData = strData & vbNewLine & Trim(varCellValue)
                            End If
                        End If
                        .GetText 2, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            If strCode = vbNullString Then
                                strCode = Trim(varCellValue)
                            Else
                                strCode = strCode & vbNewLine & Trim(varCellValue)
                            End If
                        End If
                End Select
            End If
        Next lngRow
        grdTerms.ClearRange 1, 1, grdTerms.MaxCols, grdTerms.MaxRows, False
        Select Case mstrRefData
            Case "CPT"
                If strData <> vbNullString Then
                    Arr = Split(LookUpRow("CPT NAME"), "|")
                    If UBound(Arr) > 0 Then
                        booCPT = True
                        lngCol = Arr(0)
                         Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "1", strData, "TypeHAlignLeft", "CPT NAME"
                    End If
                    If Not booCPT Then
                        Arr = Split(LookUpRow("ITEM DESCRIPTION"), "|")
                        If UBound(Arr) > 0 Then
                            booCPT = True
                            lngCol = Arr(0)
                             Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "2", strData, "TypeHAlignLeft", "CPT_INTRVN_DESC"
                        End If
                    End If
                End If
                If strCode <> vbNullString Then
                    Arr = Split(LookUpRow("CPT CODE"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                         Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "1", strCode, "TypeHAlignLeft", "CPT CODE"
                    End If
                End If
            Case "CPT_INTRVN_DESC", "CPT_INTRVN"
                If strData <> vbNullString Then
                    Arr = Split(LookUpRow("CPT_INTRVN_DESC"), "|")
                    If UBound(Arr) > 0 Then
                        booCPT = True
                        lngCol = Arr(0)
                         Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "2", strData, "TypeHAlignLeft", "CPT NAME"
                    End If
                    If Not booCPT Then
                        Arr = Split(LookUpRow("ITEM DESCRIPTION"), "|")
                        If UBound(Arr) > 0 Then
                            booCPT = True
                            lngCol = Arr(0)
                             Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "2", strData, "TypeHAlignLeft", "CPT_INTRVN_DESC"
                        End If
                    End If
                End If
                If strCode <> vbNullString Then
                    Arr = Split(LookUpRow("CPT CODE"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                         Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "1", strCode, "TypeHAlignLeft", "CPT CODE"
                    End If
                End If
            Case "ICD"
                If strData <> vbNullString Then
                    Arr = Split(LookUpRow("SUG_ICD_CODE"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                         Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "1", strCode, "TypeHAlignLeft", "SUG_ICD_CODE"
                    End If
                End If
'                If strCode <> vbNullString Then
'                    Arr = Split(LookUpRow("ICD CODE"), "|")
'                    If UBound(Arr) > 0 Then
'                        lngCol = Arr(0)
'                         Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "1", strCode, "TypeHAlignLeft", "SUG_ICD_CODE"
'                    End If
'                End If
            Case "MODIFIER"
                If strData <> vbNullString Then
                    Arr = Split(LookUpRow("MODIFIER"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                         Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "1", strData, "TypeHAlignLeft", "CPT_MODIFIER"
                    End If
                End If
            Case "POS"
                If strData <> vbNullString Then
                    Arr = Split(LookUpRow("PLACE SERVICE"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                         Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "1", strData, "TypeHAlignLeft", "PLACE SERVICE", strCode
                    End If
                End If
            Case "REPORT_ID"
                If strCode <> vbNullString Then
                    Arr = Split(LookUpRow("REPORT TYPE"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                         Format_CellTypeEdit grdPrefCPT, mlngGridRow, lngCol, "1", strCode, "TypeHAlignLeft", "REPORT_ID", strCode
                    End If
                End If
        End Select
        .ReDraw = True
    End With
    Exit Sub
End Sub
Private Sub cmdGridTerms_Click()
'    Load frmNurseTemplates
'    frmNurseTemplates.fCaption = "TERMS MAINTANENCE"
'    frmNurseTemplates.fType = "TERMS MAINTANENCE"
' '   frmNurseTemplates.fSubDomain = 'mstrGridTerms
'  '  frmNurseTemplates.fConnectTM = psConnect(1)
'    frmNurseTemplates.ReadData
'    frmNurseTemplates.Show vbModal
End Sub
Private Sub Format_Evaluation()
Dim lngColWidth As Long
Dim lngCol As Long

    On Error GoTo Error_Format_Evaluation
    Screen.MousePointer = vbHourglass
    grdPrefCPT.ClearRange 0, SpreadHeader, grdPrefCPT.MaxCols, grdPrefCPT.MaxRows, False
    With grdPrefCPT
        .ReDraw = False
        .MaxRows = 0
        .MaxCols = 8
        Format_CellTypeEdit grdPrefCPT, SpreadHeader, 1, "maxcols", "Review/Manage Preferred Evaluation Management Codes", "TypeHAlignCenter", "EVALUATION", "", HeaderColor
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, 0, "1", "GROUP", "TypeHAlignCenter", "", "", HeaderColor
        .SetCellBorder 1, SpreadHeader, 9, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdPrefCPT.ColHidden = True
        lngCol = 1
        grdPrefCPT.ColWidth(lngCol) = 12
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "CODE TYPE", "TypeHAlignCenter", "CPT_SUBGROUP", "", HeaderColor
        .SetCellBorder 1, SpreadHeader + 1, 1, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        grdPrefCPT.ColWidth(lngCol) = 25
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "CPT NAME", "TypeHAlignCenter", "CPT_NAME", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "CPT CODE", "TypeHAlignCenter", "CPT_CODE", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        grdPrefCPT.ColWidth(lngCol) = 15
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "ENCOUNTER TYPE", "TypeHAlignCenter", "ENCNTR_TYPE", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "PLACE SERVICE", "TypeHAlignCenter", "PLACE_SERVICE", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "PROVIDER", "TypeHAlignCenter", "PROVIDER", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        grdPrefCPT.ColWidth(lngCol) = 8
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "FEE", "TypeHAlignCenter", "CPT_FEE", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 8
        grdPrefCPT.ColWidth(lngCol) = 12
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "REPORT TYPE", "TypeHAlignCenter", "REPORT_ID", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 2, grdPrefCPT
        .TopRow = 1
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Evaluation:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Evaluation", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Function LookUpRow(vstrName As String) As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim booData As Boolean
Dim Arr
Dim strFound As String
Dim booExit As Boolean

    On Error GoTo Error_LookUpRow
    Screen.MousePointer = vbHourglass
    
    With grdPrefCPT
        LookUpRow = vbNullString
        For lngCol = 1 To .MaxCols
                .Col = lngCol
                .Row = SpreadHeader + 1
                .GetText lngCol, SpreadHeader + 1, varCellValue
                If Trim(varCellValue) <> vbNullString And TextFound(vstrName, varCellValue, Len(vstrName), True) Then
                    strFound = Trim(varCellValue)
                     LookUpRow = lngCol & "|" & SpreadHeader + 1
                     booExit = True
                     Exit For
                End If
            If booExit Then Exit For
        Next lngCol
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUpRow:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "LookUpRow", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Sub cmdRemove_Click()
Dim strSQL As String
Dim iRows As Long
Dim IAnswer As Integer
Dim varCode
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim lngCPTCode As Long
Dim lngPrimKey As Long

    grdPrefCPT.Col = mlngGridCol
    grdPrefCPT.Row = mlngGridRow
    On Error GoTo Error_Remove
    Screen.MousePointer = vbHourglass
    
    If mlngGridRow > 0 And grdPrefCPT.BackColor = grdCurItem Then
        With grdPrefCPT
            .ReDraw = False
            lngPrimKey = .GetRowItemData(mlngGridRow)
            IAnswer = MsgBox("Are you sure you want to delete " & varCode, vbYesNo)
            If IAnswer = vbYes Then
                If clsCptIntrvn.CPT_Type = "Evaluation" Then
                    If lngPrimKey > 0 Then
                        If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
                        strSQL = "update  CPT_TEMPLATE set Purge = 1 where CPT_ID = " & lngPrimKey
                        Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
                    End If
                    grdPrefCPT.DeleteRows mlngGridRow, 1
                    .MaxRows = .MaxRows - 1
                    MsgBox "succesfully removed!", vbInformation
                Else
                    If lngPrimKey > 0 Then
                        If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
                        strSQL = "update  INTRVN_TEMPLATE set Purge = 1 where INTRVN_ID = " & lngPrimKey
                        Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
                    End If
                    .DeleteRows mlngGridRow, 1
                    .MaxRows = .MaxRows - 1
                    MsgBox "succesfully removed!", vbInformation
                End If
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
    Write_Data
    Exit Sub
End Sub
Private Sub Form_Activate()
    clsCptIntrvn.CPT_Type = vbNullString
End Sub
Private Sub fraGridTerm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    flngDeltaX = X
    flngDeltaY = Y
End Sub
Private Sub fraGridTerm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        fraGridTerm.Left = fraGridTerm.Left + X - flngDeltaX
        fraGridTerm.Top = fraGridTerm.Top + Y - flngDeltaY
    End If
End Sub
Private Sub grdPrefCPT_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
Dim strData As String
    grdPrefCPT.Col = Col
    grdPrefCPT.Row = Row
    If grdPrefCPT.CellType = CellTypeComboBox Then
        grdPrefCPT.TypeComboBoxIndex = grdPrefCPT.TypeComboBoxCurSel
        If grdPrefCPT.TypeComboBoxString <> vbNullString Then Get_SubGroupID grdPrefCPT.TypeComboBoxString
        grdPrefCPT.CellTag = clsCptIntrvn.Grp_SubGroup_ID
    End If
    Exit Sub
End Sub
Private Sub grdPrefCPT_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim strSQL As String
Dim strTag As String
Dim strList As String
Dim strRPN As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue

    If Row <= 0 Then Exit Sub
    mlngGridRow = Row: lngRow = Row
    mlngGridCol = Col
    With grdPrefCPT
        If mlngGridRow <= 0 Then Exit Sub
        If clsCptIntrvn.CPT_Type = "Orders" Then
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                .BackColor = grdCurItem
            .BlockMode = False
            .SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 2
            .SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 4
            .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 5
            .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 6
            .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 8
            .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 9
            .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Else
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                .BackColor = grdCurItem
             .BlockMode = False
            .SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 2
            .SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 4
            .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 5
            .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 6
            .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 8
            .SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 9
            .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        .Col = mlngGridCol
        .Row = mlngGridRow
        If .CellTag <> vbNullString Then
            clsGridFormat.CellTag = .CellTag
            Select Case UCase(clsGridFormat.CellTag)
                Case "SUG_ICD_CODE", "ICD CODE"
                    Load_GridBox "ICD"
                    GetTerms_Visible grdPrefCPT, Col, mlngGridRow, "ICD"
                    cctSearch.Visible = True
                Case "CPT_INTRVN_CODE"
                    .GetText Col, Row, varCellValue
                    cctSearch.Text = varCellValue
                    Load_GridBox "CPT_INTRVN"
                    GetTerms_Visible grdPrefCPT, Col, mlngGridRow, "CPT_INTRVN_CODE"
                    cctSearch.Visible = True
                Case "CPT_INTRVN_DESC"
                    .GetText Col, Row, varCellValue
                    cctSearch.Text = varCellValue
                    Load_GridBox "CPT_INTRVN"
                    GetTerms_Visible grdPrefCPT, Col, mlngGridRow, "CPT_INTRVN_DESC"
                    cctSearch.Visible = True
                Case "CPT_MODIFIER"
'                Load_GridBox "MODIFIER"
                    GetTerms_Visible grdPrefCPT, Col, mlngGridRow
                Case "ITEM DESCRIPTION", "CPT_CODE", "CPT_NAME"
                    Load_GridBox "CPT"
                    GetTerms_Visible grdPrefCPT, Col, mlngGridRow, clsGridFormat.CellTag
                    cctSearch.Visible = True
                Case "SUG_ICD_CODE", "ICD CODE"
                    GetTerms_Visible grdPrefCPT, Col, mlngGridRow
                Case "TYPE"
                    If clsCptIntrvn.CPT_Type = "Orders" Then
                        strList = Get_InvtrvTemp
                        Format_CellTypeCombo grdPrefCPT, mlngGridRow, Col, "1", strList, "TypeHAlignLeft", "TYPE"
                        .SetCellBorder Col, mlngGridRow, Col, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                Case "CPT_SUBGROUP", "CPT SUBGROUP"
                    If TextFound("Evaluation", clsCptIntrvn.CPT_Type, Len("Evaluation"), True) Then
                        strList = Get_VisitType
                        Format_CellTypeCombo grdPrefCPT, mlngGridRow, Col, "1", strList, "TypeHAlignLeft", "CPT_SUBGROUP"
                        .SetCellBorder Col, mlngGridRow, Col, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                Case "ENCNTR TYPE", "ENCNTR_TYPE"
                    If TextFound("Evaluation", clsCptIntrvn.CPT_Type, Len("Evaluation"), True) Then
                        strList = Get_EnctrType
                        Format_CellTypeCombo grdPrefCPT, mlngGridRow, Col, "1", strList, "TypeHAlignLeft", "ENCNTR_TYPE"
                        .SetCellBorder Col, mlngGridRow, Col, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                Case "PROVIDER"
                    Get_Provider_Title
                    Format_CellTypeCombo grdPrefCPT, mlngGridRow, Col, "1", mstrList, "TypeHAlignLeft", "PROVIDER"
                    .SetCellBorder Col, mlngGridRow, Col, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "PLACE_SERVICE"
                    Load_GridBox "POS"
                    GetTerms_Visible grdPrefCPT, Col, mlngGridRow
                Case "REPORT_ID"
                    Load_GridBox "REPORT_ID"
                    GetTerms_Visible grdPrefCPT, Col, mlngGridRow
            End Select
        End If
    End With
    Exit Sub
End Sub
Private Function Get_GridTerms(vstrSubDom As String)
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection

    On Error GoTo Error_Get_GridTerms
    Screen.MousePointer = vbHourglass
    
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    mstrList = vbNullString
    strSQL = "Select DISTINCT ATTRBTE_VALID_VALUE, valid_value_seq_nbr " _
        & "From ATTR_VALID_VALUE " _
        & "Where VALID_VALUE_SUBDOM = '" & vstrSubDom & "' ORDER BY valid_value_seq_nbr"
    Set rsTerms = cnTerms.Execute(strSQL, adAsyncExecute)
    With rsTerms
        Do Until .EOF
            If (!ATTRBTE_VALID_VALUE) <> vbNullString Then
                If Trim(mstrList) = vbNullString Then
                    mstrList = Trim(!ATTRBTE_VALID_VALUE)
                Else
                    mstrList = mstrList & Chr(9) & Trim(!ATTRBTE_VALID_VALUE)
                End If
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
Error_Get_GridTerms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", " Get_GridTerms", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Sub Load_CPT_Template(Optional vstrSerPlace As String)
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim lngCPT_ID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim strPos As String
Dim lngColWidth As Long
Dim strTag As String

    On Error GoTo Error_Load_CPT_Template
    Screen.MousePointer = vbHourglass
    grdPrefCPT.ReDraw = False
    If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
    If vstrSerPlace = vbNullString Then
        strSQL = "Select * FROM CPT_TEMPLATE Where Purge =  0  ORDER BY CPT_CODE "
    End If
    Set rsCPT = cnCPT.Execute(strSQL, adAsyncExecute)
        Do Until rsCPT.EOF
            Clear_CPT_Temp
            clsGridFormat.GridReportID = 0
            Set Flds = rsCPT.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "CPT_ID"
                                clsCptTemp.CPT_ID = Fld.Value
                            Case "CPT_SUBGROUP"
                                clsCptTemp.CPT_Subgroup = Fld.Value
                            Case "CPT_NAME"
                                clsCptTemp.CPT_Name = Fld.Value
                            Case "CPT_CODE"
                                clsCptTemp.CPT_Code = Fld.Value
                            Case "ENCNTR_TYPE"
                                clsCptTemp.Encntr_Type = Fld.Value
                            Case "SERVICE_PLACE"
                                clsCptTemp.Service_Place = Fld.Value
                            Case "PLACE_SERVICE_CD"
                                clsCptTemp.Place_Service_CD = Fld.Value
                            Case "PROVIDER"
                                clsCptTemp.Provider = Fld.Value
                            Case "CPT_FEE"
                                clsCptTemp.CPT_Fee = Fld.Value
                            Case "REPORT_ID"
                                clsGridFormat.GridReportID = Fld.Value
                        End Select
                    End If
                End If
            Next
            rsCPT.MoveNext
            Add_Template_Row
            lngRow = grdPrefCPT.MaxRows
            grdPrefCPT.SetRowItemData lngRow, Val(clsCptTemp.CPT_ID)
            lngCol = 1
            Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", clsCptTemp.CPT_Subgroup, "TypeHAlignCenter", "CPT_SUBGROUP"
            grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 2
            Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", clsCptTemp.CPT_Name, "TypeHAlignLeft", "CPT_NAME"
            grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 3
            Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", clsCptTemp.CPT_Code, "TypeHAlignCenter", "CPT_CODE"
            grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 4
            Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", clsCptTemp.Encntr_Type, "TypeHAlignCenter", "ENCNTR_TYPE"
            grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 5
            Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", clsCptTemp.Service_Place, "TypeHAlignLeft", "PLACE_SERVICE", clsCptTemp.Place_Service_CD
            grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 6
            Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", clsCptTemp.Provider, "TypeHAlignCenter", "PROVIDER"
            grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 7
            Format_CellTypeEdit grdPrefCPT, lngRow, lngCol, "1", clsCptTemp.CPT_Fee, "TypeHAlignCenter", "CPT_FEE"
            grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 8
            Format_CellTypeStatic grdPrefCPT, lngRow, lngCol, "1", CStr(clsGridFormat.GridReportID), "TypeHAlignCenter", "REPORT_ID"
            grdPrefCPT.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Loop
        grdPrefCPT.ReDraw = True
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
   Exit Sub
Error_Load_CPT_Template:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_CPT_Template", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub Format_Orders()
Dim lngColWidth As Long
Dim lngCol As Long

    
   On Error GoTo Error_Format_Orders
    Screen.MousePointer = vbHourglass
    grdPrefCPT.ClearRange 0, SpreadHeader, grdPrefCPT.MaxCols, grdPrefCPT.MaxRows, False
    With grdPrefCPT
        .UserColAction = UserColActionSort
        .ColHeadersUserSortIndex = 0
        .ReDraw = False
        .MaxRows = 0
        .MaxCols = 8
        Format_CellTypeEdit grdPrefCPT, SpreadHeader, 1, "maxcols", "Review/Manage Preferred CPT Codes", "TypeHAlignCenter", "ORDERS", "", HeaderColor
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, 0, "1", "GROUP", "TypeHAlignCenter", "", "", HeaderColor
        .SetCellBorder 1, SpreadHeader, 9, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdPrefCPT.ColHidden = True
        lngCol = 1
        grdPrefCPT.ColWidth(lngCol) = 12
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "TYPE", "TypeHAlignCenter", "TYPE", "", HeaderColor
        .SetCellBorder 1, SpreadHeader + 1, 1, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        grdPrefCPT.ColWidth(lngCol) = 25
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "2", "ITEM DESCRIPTION", "TypeHAlignCenter", "CPT_INTRVN_DESC", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol + 1, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "CPT CODE", "TypeHAlignCenter", "CPT_INTRVN_CODE", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "CPT SUBSET", "TypeHAlignCenter", "CPT_MODIFIER", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        grdPrefCPT.ColWidth(lngCol) = 15
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "DIAGNOSIS", "TypeHAlignCenter", "SUG_ICD_CODE", "DEFAULT SELECTION", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "PROVIDER", "TypeHAlignCenter", "PROVIDER", "DEFAULT SELECTION", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 8
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "FEE", "TypeHAlignCenter", "CPT_FEE", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 2, grdPrefCPT
        .TopRow = 1
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Orders:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Orders", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub Format_Procedure()
Dim lngColWidth As Long
Dim lngCol As Long

   On Error GoTo Error_Format_Procedure
    Screen.MousePointer = vbHourglass
    grdPrefCPT.ClearRange 0, SpreadHeader, grdPrefCPT.MaxCols, grdPrefCPT.MaxRows, False
    With grdPrefCPT
        .UserColAction = UserColActionSort
        .ColHeadersUserSortIndex = 0
        .ReDraw = False
        .MaxRows = 0
        .MaxCols = 9
        Format_CellTypeEdit grdPrefCPT, SpreadHeader, 1, "maxcols", "Review/Manage Preferred CPT Codes", "TypeHAlignCenter", "ORDERS", "", HeaderColor
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, 0, "1", "GROUP", "TypeHAlignCenter", "", "", HeaderColor
        .SetCellBorder 1, SpreadHeader, 9, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdPrefCPT.ColHidden = True
        lngCol = 1
        grdPrefCPT.ColWidth(lngCol) = 12
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "TYPE", "TypeHAlignCenter", "TYPE", "", HeaderColor
        .SetCellBorder 1, SpreadHeader + 1, 1, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        grdPrefCPT.ColWidth(lngCol) = 25
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "2", "ITEM DESCRIPTION", "TypeHAlignCenter", "CPT_INTRVN_DESC", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol + 1, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "CPT CODE", "TypeHAlignCenter", "CPT_INTRVN_CODE", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "CPT SUBSET", "TypeHAlignCenter", "CPT_MODIFIER", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        grdPrefCPT.ColWidth(lngCol) = 15
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "DIAGNOSIS", "TypeHAlignCenter", "SUG_ICD_CODE", "DEFAULT SELECTION", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol + 1, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "PROVIDER", "TypeHAlignCenter", "PROVIDER", "DEFAULT SELECTION", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 8
        grdPrefCPT.ColWidth(lngCol) = 10
        Format_CellTypeEdit grdPrefCPT, SpreadHeader + 1, lngCol, "1", "FEE", "TypeHAlignCenter", "CPT_FEE", "", HeaderColor
        .SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 1, SpreadHeader, .MaxCols, SpreadHeader, 8, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 2, grdPrefCPT
        .TopRow = 1
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Procedure:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Procedure", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub grdPrefCPT_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    With grdPrefCPT
        If Row >= 0 Then
            .BlockMode = True
                .Col = Col
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                .BackColor = &HFFFFFF
                .SetCellBorder 1, Row, .MaxCols, Row, CellBorderIndexOutline, 0, CellBorderStyleSolid
             .BlockMode = False
        End If
    End With
    Exit Sub
End Sub


Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
    With grdTerms
        cctSearch.Text = vbNullString
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


Private Sub grdTerms_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    With grdTerms
        If Row >= 0 Then
            .BlockMode = True
                .Col = Col
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                 .BackColor = &HFFFFFF
             .BlockMode = False
        End If
    End With
End Sub
Private Sub InitializeTerms()
    On Error GoTo Error_InitializeTerms

    Exit Sub

Error_InitializeTerms:

End Sub
Private Sub UserControl_Terminate()
    clsCptIntrvn.CPT_Type = vbNullString
End Sub
Private Sub mnu_EM_Click()
    Format_Evaluation
    Load_CPT_Template
    clsCptIntrvn.CPT_Type = "Evaluation"
    picLogo.Visible = False
    Exit Sub
End Sub
Private Sub mnu_Labs_Click()
    Format_Orders
    Load_CPTIntrvn
    clsCptIntrvn.CPT_Type = "Orders"
    picLogo.Visible = False
    Exit Sub
End Sub
Private Sub mnu_Print_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer
Dim CellNote As Long
Dim intPrint As Integer
Dim strRef As String

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    mbooCancelPrint = False
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation
        Exit Sub
    End If
    With grdPrefCPT
        For intPrint = 1 To Printer.Copies
                .PrintType = PrintTypeAll
                .hDCPrinter = Printer.hdc
                .PrintColHeaders = True
                .PrintRowHeaders = False
                .PrintScalingMethod = PrintScalingMethodSmartPrint
                .PrintCenterOnPageH = True
                .PrintMarginBottom = 770
                .PrintBorder = True
                .PrintColor = True
                .PrintOrientation = PrintOrientationPortrait
                .PrintShadows = False
                .PrintSheet
        Next intPrint
    mbooPrintPreview = False
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdPrint_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdPrint_Click", App.EXEName & "." & Me.Name
    Exit Sub
End Sub
Private Sub mnu_ProcTest_Click()
    Format_Procedure
    Load_CPTIntrvn "Procedure"
    clsCptIntrvn.CPT_Type = "Orders"
    picLogo.Visible = False
    Exit Sub
End Sub

