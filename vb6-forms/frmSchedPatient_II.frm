VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{6514F5A0-641C-11D2-9FD0-0020AF131A57}#3.0#0"; "fpFlp30.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmSchedPatient_II 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Schedule Patient Appointments"
   ClientHeight    =   9600
   ClientLeft      =   4020
   ClientTop       =   3705
   ClientWidth     =   21360
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchedPatient_II.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleMode       =   0  'User
   ScaleWidth      =   21360
   Visible         =   0   'False
   Begin LpADOLib.fpComboAdo fpComboAdo1 
      Height          =   315
      Left            =   0
      TabIndex        =   106
      Top             =   0
      Width           =   2085
      _Version        =   196608
      _ExtentX        =   3678
      _ExtentY        =   556
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   4194304
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   3
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   4194304
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   1
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   0
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      DataMemberList  =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   2
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      AllowAnimate    =   -1  'True
      ColDesigner     =   "frmSchedPatient_II.frx":058A
   End
   Begin VB.Frame fraGetAppts 
      Height          =   7305
      Left            =   15120
      TabIndex        =   100
      Top             =   4800
      Visible         =   0   'False
      Width           =   10515
      Begin MSComCtl2.MonthView ccSchedDate 
         Height          =   2670
         Left            =   90
         TabIndex        =   101
         Top             =   180
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   4710
         _Version        =   393216
         ForeColor       =   8388608
         BackColor       =   16744448
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   77201409
         TitleBackColor  =   16761024
         TitleForeColor  =   8388608
         TrailingForeColor=   8421504
         CurrentDate     =   41640
      End
      Begin FPSpreadADO.fpSpread grdAvailAppt 
         Height          =   7065
         Left            =   3480
         TabIndex        =   102
         Top             =   150
         Width           =   6930
         _Version        =   524288
         _ExtentX        =   12224
         _ExtentY        =   12462
         _StockProps     =   64
         AutoCalc        =   0   'False
         ColHeaderDisplay=   0
         DAutoCellTypes  =   0   'False
         DAutoFill       =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridColor       =   4194304
         GridSolid       =   0   'False
         MaxCols         =   2
         MaxRows         =   0
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         ShadowText      =   8388608
         SpreadDesigner  =   "frmSchedPatient_II.frx":084D
         UserResize      =   0
         Appearance      =   2
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
         CellNoteIndicator=   3
         ScrollBarStyle  =   3
      End
      Begin PT_XP_Button.PT_XP cmdSelectAvail 
         Height          =   405
         Left            =   750
         TabIndex        =   103
         Top             =   6750
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmSchedPatient_II.frx":0D68
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Select"
         CaptionAlignment=   5
         CaptionPosition =   4
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":1102
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCloseAvail 
         Height          =   405
         Left            =   1860
         TabIndex        =   104
         Top             =   6750
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmSchedPatient_II.frx":1454
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":1E66
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9060
      Top             =   2010
   End
   Begin VB.Frame fraPlan 
      BackColor       =   &H00E0E0E0&
      Height          =   2985
      Left            =   720
      TabIndex        =   90
      Top             =   5040
      Visible         =   0   'False
      Width           =   5865
      Begin FPSpreadADO.fpSpread grdPlan 
         Height          =   2175
         Left            =   120
         TabIndex        =   91
         ToolTipText     =   "Select a row and right click to remove an item."
         Top             =   240
         Width           =   5625
         _Version        =   524288
         _ExtentX        =   9922
         _ExtentY        =   3836
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         ColHeaderDisplay=   0
         DAutoCellTypes  =   0   'False
         DAutoFill       =   0   'False
         DAutoHeadings   =   0   'False
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
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridColor       =   14737632
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   3
         MaxRows         =   0
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         ShadowText      =   4194304
         SpreadDesigner  =   "frmSchedPatient_II.frx":21B8
         UnitType        =   0
         UserResize      =   0
         Appearance      =   2
         TextTipDelay    =   50
         ClipboardOptions=   0
         CellNoteIndicator=   3
      End
      Begin PT_XP_Button.PT_XP cmdPlan 
         Height          =   345
         Left            =   2550
         TabIndex        =   92
         Top             =   2520
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "frmSchedPatient_II.frx":26FC
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":310E
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.TextBox cctFacility 
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
      Height          =   705
      Left            =   6165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   78
      Top             =   105
      Width           =   3765
   End
   Begin VB.ComboBox cboSchedResource 
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
      Left            =   6165
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   1185
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ComboBox cboApptType 
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
      ItemData        =   "frmSchedPatient_II.frx":3460
      Left            =   2040
      List            =   "frmSchedPatient_II.frx":3476
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   105
      Width           =   2970
   End
   Begin PT_XP_Button.PT_XP cmdEditPatient 
      Height          =   345
      Left            =   16020
      TabIndex        =   44
      Top             =   315
      Visible         =   0   'False
      Width           =   1830
      _ExtentX        =   3228
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
      Picture         =   "frmSchedPatient_II.frx":34B0
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Update Patient"
      CaptionAlignment=   5
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedPatient_II.frx":3A4A
      ShowFocus       =   -1  'True
   End
   Begin VB.ComboBox cboProvider 
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
      Left            =   6165
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1575
      Width           =   3870
   End
   Begin PT_XP_Button.PT_XP cmdTakePic 
      Height          =   435
      Left            =   14190
      TabIndex        =   37
      Top             =   45
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSchedPatient_II.frx":3D9C
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmSchedPatient_II.frx":4676
      ShowFocus       =   -1  'True
   End
   Begin VB.Frame fraPic 
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
      Height          =   2145
      Left            =   10110
      TabIndex        =   35
      Top             =   60
      Width           =   3225
      Begin VB.Image imgPatient 
         Height          =   1935
         Left            =   60
         Stretch         =   -1  'True
         Top             =   150
         Width           =   3075
      End
      Begin VB.Label lblPhoto 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NO PHOTO AVAILABLE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   870
         TabIndex        =   36
         Top             =   840
         Visible         =   0   'False
         Width           =   1665
      End
   End
   Begin VB.TextBox txtList 
      BackColor       =   &H00FFFFFF&
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
      Height          =   735
      Left            =   4170
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   8070
      Visible         =   0   'False
      Width           =   4845
   End
   Begin VB.TextBox cctSearch 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   690
      Width           =   4875
   End
   Begin VB.ListBox lstListBox 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      ItemData        =   "frmSchedPatient_II.frx":52C8
      Left            =   120
      List            =   "frmSchedPatient_II.frx":52CA
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1155
      Width           =   4875
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   345
      Left            =   6780
      TabIndex        =   11
      Top             =   9255
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
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
      Picture         =   "frmSchedPatient_II.frx":52CC
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedPatient_II.frx":5CDE
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdApply 
      Height          =   345
      Left            =   5625
      TabIndex        =   10
      Top             =   9255
      Visible         =   0   'False
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
      Picture         =   "frmSchedPatient_II.frx":6030
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "&Apply"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedPatient_II.frx":63CA
      ShowFocus       =   -1  'True
   End
   Begin VB.Frame fraInfo 
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
      Height          =   7395
      Left            =   90
      TabIndex        =   14
      Top             =   2250
      Width           =   13275
      Begin VB.Frame fraPsycho 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3435
         Left            =   -1290
         TabIndex        =   95
         Top             =   5520
         Visible         =   0   'False
         Width           =   9075
         Begin FPSpreadADO.fpSpread grdPsychotherapy 
            Height          =   2385
            Left            =   30
            TabIndex        =   96
            Top             =   540
            Width           =   8985
            _Version        =   524288
            _ExtentX        =   15849
            _ExtentY        =   4207
            _StockProps     =   64
            AutoCalc        =   0   'False
            ColHeaderDisplay=   0
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
            GridColor       =   8388608
            MaxCols         =   9
            MaxRows         =   0
            MoveActiveOnFocus=   0   'False
            ProcessTab      =   -1  'True
            RetainSelBlock  =   0   'False
            ShadowColor     =   16761024
            ShadowDark      =   8388608
            ShadowText      =   8388608
            SpreadDesigner  =   "frmSchedPatient_II.frx":671C
            UserResize      =   0
            Appearance      =   2
            ClipboardOptions=   0
         End
         Begin PT_XP_Button.PT_XP cmdPsyAdd 
            Height          =   345
            Left            =   3510
            TabIndex        =   97
            Top             =   3000
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
            Picture         =   "frmSchedPatient_II.frx":6C0B
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Select"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSchedPatient_II.frx":6FA5
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdPsyClose 
            Height          =   345
            Left            =   4650
            TabIndex        =   98
            Top             =   3000
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
            Picture         =   "frmSchedPatient_II.frx":72F7
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSchedPatient_II.frx":7D09
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdRemAddOn 
            Height          =   315
            Left            =   7680
            TabIndex        =   99
            Top             =   180
            Width           =   1290
            _ExtentX        =   2275
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
            Picture         =   "frmSchedPatient_II.frx":805B
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Remove"
            DefaultThemes   =   2
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSchedPatient_II.frx":85F5
            ShowFocus       =   -1  'True
         End
      End
      Begin VB.TextBox cctRepeat 
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
         Height          =   615
         Left            =   9420
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   69
         Top             =   1170
         Width           =   3315
      End
      Begin VB.Frame fraRefer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Referring Physician"
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
         Height          =   3015
         Left            =   6870
         TabIndex        =   48
         Top             =   3510
         Visible         =   0   'False
         Width           =   6345
         Begin VB.ListBox lstReferral 
            Height          =   960
            Left            =   1230
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   77
            Top             =   1200
            Visible         =   0   'False
            Width           =   4965
         End
         Begin VB.TextBox cctReferral 
            Height          =   405
            Left            =   1230
            TabIndex        =   76
            Text            =   "Type Last Name"
            Top             =   780
            Width           =   4965
         End
         Begin VB.TextBox cctReasReferral 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            ToolTipText     =   "Double Click for Reason Terms List"
            Top             =   1575
            Width           =   6105
         End
         Begin PT_XP_Button.PT_XP cmdClear 
            Height          =   255
            Left            =   5490
            TabIndex        =   49
            Top             =   1260
            Width           =   645
            _ExtentX        =   1138
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
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Clear"
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdReferralPhys 
            Height          =   375
            Left            =   1290
            TabIndex        =   51
            Top             =   300
            Width           =   3255
            _ExtentX        =   5741
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
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Add New Referring Physician"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdRefClose 
            Height          =   345
            Left            =   2460
            TabIndex        =   52
            Top             =   2550
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
            Picture         =   "frmSchedPatient_II.frx":8947
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close "
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSchedPatient_II.frx":9359
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdRefTerms 
            Height          =   315
            Left            =   5490
            TabIndex        =   53
            ToolTipText     =   "Add Referral Reason Terms"
            Top             =   4440
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
            Picture         =   "frmSchedPatient_II.frx":96AB
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
            PictureDisabled =   "frmSchedPatient_II.frx":9C45
            ShowFocus       =   -1  'True
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Reason for Referral"
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
            Left            =   150
            TabIndex        =   55
            Top             =   1230
            Width           =   2025
         End
         Begin VB.Label lblReferralPhy 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Physician"
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
            Left            =   150
            TabIndex        =   54
            Top             =   870
            Width           =   960
         End
      End
      Begin VB.TextBox cctReferalReas 
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
         Height          =   345
         Left            =   8925
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   63
         Top             =   3990
         Width           =   4185
      End
      Begin VB.ComboBox cboReferBy 
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
         Height          =   330
         ItemData        =   "frmSchedPatient_II.frx":9F97
         Left            =   8925
         List            =   "frmSchedPatient_II.frx":9F99
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   3510
         Width           =   2865
      End
      Begin VB.ComboBox cboDuration 
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
         Height          =   330
         Left            =   1770
         TabIndex        =   57
         Top             =   6840
         Width           =   2145
      End
      Begin VB.ComboBox cboSerPlace 
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
         Height          =   330
         ItemData        =   "frmSchedPatient_II.frx":9F9B
         Left            =   2595
         List            =   "frmSchedPatient_II.frx":9F9D
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Tag             =   "Place of Service"
         Top             =   3240
         Width           =   3750
      End
      Begin MSComCtl2.DTPicker txtApptConfirm 
         Height          =   375
         Left            =   9420
         TabIndex        =   7
         Top             =   2850
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Format          =   77201411
         CurrentDate     =   38792
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   345
         Left            =   1080
         TabIndex        =   2
         Top             =   2310
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
         CalendarTrailingForeColor=   14737632
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyy"
         DateIsNull      =   -1  'True
         Format          =   77201411
         CurrentDate     =   38792
         MinDate         =   36526
      End
      Begin PT_XP_Button.PT_XP cmdExit 
         Height          =   345
         Left            =   6240
         TabIndex        =   12
         Top             =   6870
         Width           =   1215
         _ExtentX        =   2143
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
         Picture         =   "frmSchedPatient_II.frx":9F9F
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close Me"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":A9B1
         PictureAlignment=   6
         ShowFocus       =   -1  'True
      End
      Begin VB.ComboBox cboPatientType 
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
         Height          =   330
         Left            =   3615
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3750
         Width           =   2730
      End
      Begin VB.ComboBox cboFacility 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         ItemData        =   "frmSchedPatient_II.frx":AD03
         Left            =   2010
         List            =   "frmSchedPatient_II.frx":AD05
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2820
         Width           =   4335
      End
      Begin VB.TextBox txtComments 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   8925
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   4410
         Width           =   4185
      End
      Begin VB.ComboBox cboReason 
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
         Height          =   330
         Left            =   9420
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2370
         Width           =   3285
      End
      Begin VB.ComboBox cboStatus 
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
         Height          =   330
         ItemData        =   "frmSchedPatient_II.frx":AD07
         Left            =   9420
         List            =   "frmSchedPatient_II.frx":AD09
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1950
         Width           =   3285
      End
      Begin PT_XP_Button.PT_XP cmdStatus 
         Height          =   315
         Left            =   12720
         TabIndex        =   40
         Top             =   1950
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmSchedPatient_II.frx":AD0B
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmSchedPatient_II.frx":B2A5
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdReason 
         Height          =   315
         Left            =   12720
         TabIndex        =   41
         Top             =   2370
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmSchedPatient_II.frx":B5F7
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmSchedPatient_II.frx":BB91
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRefReason 
         Height          =   315
         Left            =   12720
         TabIndex        =   43
         Top             =   3600
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmSchedPatient_II.frx":BEE3
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmSchedPatient_II.frx":C47D
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdChange 
         Height          =   315
         Left            =   11820
         TabIndex        =   60
         ToolTipText     =   "Reset Referrals to Drop Down Selection!"
         Top             =   3510
         Width           =   375
         _ExtentX        =   661
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
         Picture         =   "frmSchedPatient_II.frx":C7CF
         AppearanceThemes=   3
         BorderWidth     =   2
         TransparentColor=   14737632
         Caption         =   ""
         CaptionAlignment=   5
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":CD69
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdChangeUser 
         Height          =   465
         Left            =   11460
         TabIndex        =   68
         Top             =   180
         Width           =   1770
         _ExtentX        =   3122
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
         Picture         =   "frmSchedPatient_II.frx":D0BB
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Manage Appointment"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":D655
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRepeat 
         Height          =   345
         Left            =   12750
         TabIndex        =   70
         ToolTipText     =   "Change Repeat Appts"
         Top             =   1170
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
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
         Picture         =   "frmSchedPatient_II.frx":D9A7
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Change Appt"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":DF41
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdShowAll 
         Height          =   345
         Left            =   2130
         TabIndex        =   73
         ToolTipText     =   "Show All"
         Top             =   4290
         Width           =   315
         _ExtentX        =   556
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
         Picture         =   "frmSchedPatient_II.frx":E293
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Change Appt"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":E82D
         ShowFocus       =   -1  'True
      End
      Begin MSComCtl2.DTPicker ccdTime 
         Height          =   345
         Left            =   4935
         TabIndex        =   75
         Top             =   2325
         Width           =   1455
         _ExtentX        =   2566
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
         CustomFormat    =   "hh:mm tt"
         Format          =   77201411
         UpDown          =   -1  'True
         CurrentDate     =   39812.3652777778
      End
      Begin PT_XP_Button.PT_XP cmdAcct 
         Height          =   480
         Left            =   9945
         TabIndex        =   80
         Top             =   180
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   847
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
         Picture         =   "frmSchedPatient_II.frx":EB7F
         AppearanceThemes=   3
         BorderDrawEvent =   1
         TransparentColor=   14737632
         Caption         =   "Account Balance"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":F119
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPhyPlan 
         Height          =   345
         Left            =   2550
         TabIndex        =   93
         ToolTipText     =   "Physician Follow-Up Instructions"
         Top             =   4290
         Width           =   345
         _ExtentX        =   609
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
         Picture         =   "frmSchedPatient_II.frx":F46B
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmSchedPatient_II.frx":FA05
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdProcedure 
         Height          =   2055
         Left            =   120
         TabIndex        =   94
         Top             =   4710
         Width           =   8625
         _Version        =   524288
         _ExtentX        =   15214
         _ExtentY        =   3625
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
         MaxCols         =   3
         MaxRows         =   10
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         ShadowText      =   8388608
         SpreadDesigner  =   "frmSchedPatient_II.frx":FD57
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin PT_XP_Button.PT_XP cmdSeeAvail 
         Height          =   345
         Left            =   3000
         TabIndex        =   105
         ToolTipText     =   "See Available Appts"
         Top             =   2310
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
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
         Picture         =   "frmSchedPatient_II.frx":10259
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Date"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":107F3
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4170
         TabIndex        =   64
         Top             =   4260
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblCPT 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "CPT Code"
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
         Left            =   3030
         TabIndex        =   65
         Top             =   4350
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblReasRefer 
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for Referral"
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
         Left            =   6870
         TabIndex        =   62
         Top             =   3930
         Width           =   2145
      End
      Begin VB.Label lblReferral 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Referred By"
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
         Left            =   6900
         TabIndex        =   61
         Top             =   3570
         Width           =   1410
      End
      Begin VB.Label lblDuration 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Duration"
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
         Left            =   75
         TabIndex        =   58
         Top             =   6900
         Width           =   900
      End
      Begin VB.Label lblPurposeVisit 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment Type"
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
         Left            =   135
         TabIndex        =   56
         Top             =   4350
         Width           =   1845
      End
      Begin VB.Label lblPatType 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Place of Service"
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
         Left            =   135
         TabIndex        =   47
         Top             =   3315
         Width           =   1755
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Height          =   675
         Left            =   3810
         TabIndex        =   28
         Top             =   1080
         Width           =   2730
      End
      Begin VB.Label lblPatientType 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Type (New or Established)"
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
         Height          =   480
         Left            =   135
         TabIndex        =   45
         Top             =   3840
         Width           =   3330
      End
      Begin VB.Label lblWorkPhone 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
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
         Left            =   1350
         TabIndex        =   39
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Work "
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
         Left            =   135
         TabIndex        =   38
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Repeating Appointment"
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
         Left            =   6900
         TabIndex        =   34
         Top             =   1110
         Width           =   2625
      End
      Begin VB.Label lblConfirm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Last Confirmed"
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
         Left            =   6900
         TabIndex        =   33
         Top             =   3000
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
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
         Left            =   6870
         TabIndex        =   31
         Top             =   4410
         Width           =   1035
      End
      Begin VB.Label lblReason 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason:"
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
         Left            =   6900
         TabIndex        =   30
         Top             =   2370
         Width           =   870
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   6900
         TabIndex        =   29
         Top             =   1950
         Width           =   1215
      End
      Begin VB.Label lblCell2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
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
         Left            =   1350
         TabIndex        =   27
         Top             =   1530
         Width           =   45
      End
      Begin VB.Label lblPhone2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Left            =   1350
         TabIndex        =   26
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label lblCell1 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Phone:"
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
         Left            =   135
         TabIndex        =   25
         Top             =   1500
         Width           =   1005
      End
      Begin VB.Label lblPhone1 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Home Phone:"
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
         Left            =   135
         TabIndex        =   24
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblName2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Left            =   2850
         TabIndex        =   23
         Top             =   270
         Width           =   45
      End
      Begin VB.Label lblName1 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name:"
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
         Left            =   135
         TabIndex        =   22
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblID2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID"
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
         Left            =   7020
         TabIndex        =   21
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lblID1 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID:"
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
         Left            =   5670
         TabIndex        =   20
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblAppt1 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Current  Appointment:"
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
         TabIndex        =   19
         Top             =   660
         Width           =   2265
      End
      Begin VB.Label lblAppt2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Left            =   2850
         TabIndex        =   18
         Top             =   660
         Width           =   45
      End
      Begin VB.Label lblADate 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   135
         TabIndex        =   17
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
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
         Left            =   4290
         TabIndex        =   16
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblFacility 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Facility/Location"
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
         Left            =   135
         TabIndex        =   15
         Top             =   2850
         Width           =   1680
      End
   End
   Begin PT_XP_Button.PT_XP cmdNewPatient 
      Height          =   345
      Left            =   14400
      TabIndex        =   74
      Top             =   315
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
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
      Picture         =   "frmSchedPatient_II.frx":10B45
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "New Patient"
      CaptionAlignment=   5
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedPatient_II.frx":110DF
      ShowFocus       =   -1  'True
   End
   Begin VB.Frame fraChange 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Update/Correct Appointment"
      ClipControls    =   0   'False
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
      Height          =   7200
      Left            =   9540
      TabIndex        =   81
      Top             =   4950
      Visible         =   0   'False
      Width           =   13605
      Begin LpADOLib.fpComboAdo fpTerm 
         Height          =   315
         Left            =   4350
         TabIndex        =   82
         Top             =   240
         Width           =   2085
         _Version        =   196608
         _ExtentX        =   3678
         _ExtentY        =   556
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         BackColor       =   -2147483643
         ForeColor       =   4194304
         Text            =   ""
         Columns         =   0
         Sorted          =   0
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   -1
         ColumnWidthScale=   3
         RowHeight       =   -1
         WrapList        =   0   'False
         WrapWidth       =   0
         AutoSearch      =   1
         SearchMethod    =   0
         VirtualMode     =   0   'False
         VRowCount       =   0
         DataSync        =   3
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   4194304
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ScrollHScale    =   2
         ScrollHInc      =   0
         ColsFrozen      =   0
         ScrollBarV      =   1
         NoIntegralHeight=   0   'False
         HighestPrecedence=   0
         AllowColResize  =   0
         AllowColDragDrop=   0
         ReadOnly        =   0   'False
         VScrollSpecial  =   0   'False
         VScrollSpecialType=   0
         EnableKeyEvents =   -1  'True
         EnableTopChangeEvent=   -1  'True
         DataAutoHeadings=   -1  'True
         DataAutoSizeCols=   1
         SearchIgnoreCase=   -1  'True
         ScrollBarH      =   1
         DataFieldList   =   ""
         ColumnEdit      =   -1
         ColumnBound     =   -1
         Style           =   0
         MaxDrop         =   8
         ListWidth       =   -1
         EditHeight      =   -1
         GrayAreaColor   =   -2147483633
         ListLeftOffset  =   0
         ComboGap        =   -2
         MaxEditLen      =   150
         VirtualPageSize =   0
         VirtualPagesAhead=   0
         ExtendCol       =   0
         ColumnLevels    =   1
         ListGrayAreaColor=   -2147483637
         GroupHeaderHeight=   -1
         GroupHeaderShow =   0   'False
         AllowGrpResize  =   0
         AllowGrpDragDrop=   0
         MergeAdjustView =   0   'False
         ColumnHeaderShow=   0   'False
         ColumnHeaderHeight=   -1
         GrpsFrozen      =   0
         BorderGrayAreaColor=   -2147483637
         ExtendRow       =   0
         ListPosition    =   0
         ButtonThreeDAppearance=   0
         DataMemberList  =   ""
         OLEDragMode     =   0
         OLEDropMode     =   0
         Redraw          =   -1  'True
         AutoSearchFill  =   0   'False
         AutoSearchFillDelay=   500
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   2
         AutoMenu        =   -1  'True
         EditAlignH      =   0
         EditAlignV      =   0
         AllowAnimate    =   -1  'True
         ColDesigner     =   "frmSchedPatient_II.frx":11431
      End
      Begin PT_XP_Button.PT_XP cmdRemoveRow 
         Height          =   465
         Left            =   1845
         TabIndex        =   83
         Top             =   480
         Width           =   1740
         _ExtentX        =   3069
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
         Picture         =   "frmSchedPatient_II.frx":116F4
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Remove Appointment"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":11C8E
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdAppointChange 
         Height          =   5415
         Left            =   135
         TabIndex        =   84
         Top             =   1050
         Width           =   12900
         _Version        =   524288
         _ExtentX        =   22754
         _ExtentY        =   9551
         _StockProps     =   64
         AutoCalc        =   0   'False
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   8388608
         MaxCols         =   4
         MaxRows         =   0
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         ShadowText      =   0
         SpreadDesigner  =   "frmSchedPatient_II.frx":11FE0
         UserResize      =   0
         Appearance      =   2
         ClipboardOptions=   0
         CellNoteIndicator=   3
      End
      Begin PT_XP_Button.PT_XP cmdUpdate 
         Height          =   465
         Left            =   120
         TabIndex        =   85
         Top             =   480
         Width           =   1635
         _ExtentX        =   2884
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
         Picture         =   "frmSchedPatient_II.frx":124E9
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Update Appt Row"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":12A83
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdFrameClose 
         Height          =   375
         Left            =   6210
         TabIndex        =   86
         Top             =   6630
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "frmSchedPatient_II.frx":12DD5
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":137E7
         ShowFocus       =   -1  'True
      End
      Begin MSComCtl2.DTPicker ccdRevisedDate 
         Height          =   345
         Left            =   9240
         TabIndex        =   87
         Top             =   570
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
         Format          =   77201411
         CurrentDate     =   38792
      End
      Begin PT_XP_Button.PT_XP cmdPrint 
         Height          =   465
         Left            =   11730
         TabIndex        =   88
         Top             =   450
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   820
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmSchedPatient_II.frx":13B39
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Print"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient_II.frx":140D3
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblPrevApptDate 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   8310
         TabIndex        =   89
         Top             =   645
         Width           =   450
      End
   End
   Begin VB.Label lblFac 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Facility"
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
      Height          =   420
      Left            =   5145
      TabIndex        =   79
      Top             =   150
      Width           =   885
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Resource"
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
      Left            =   5100
      TabIndex        =   72
      Top             =   1290
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblApptType 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment Type"
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
      TabIndex        =   67
      Top             =   150
      Width           =   2085
   End
   Begin VB.Label lblProvider 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Provider"
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
      Left            =   5100
      TabIndex        =   42
      Top             =   1650
      Width           =   915
   End
   Begin VB.Label lblPatInfo 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Last, First"
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
      Height          =   165
      Left            =   120
      TabIndex        =   32
      Top             =   450
      Width           =   1275
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_Save 
         Caption         =   "Save"
      End
      Begin VB.Menu mnu_Close 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmSchedPatient_II"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mStartHour() As String
Private mEndHour() As String
Private mStartIdx As Integer
Private mEndIdx As Integer
Public mlngtop As Long
Public mlngleft As Long
Public mstrTerms As String
Public mlngPatientID As Long
Private mbooByPass As Boolean
Public mstrListName As String
Public mFirst As Boolean
Public mbooCaption As Boolean
Private mbooSearch As Boolean
Public mbooAlreadyBook As Boolean
Private mstrCalendarDate As String
Private mUpdate As Boolean
Private mBeenHere As Boolean
Private bRepeat As Boolean
Private mstrOrigApptTime As String
Private mstrValidAppoint As String
Private mbooListValid As Boolean
Private mlngGridRow As Long
Private mlngGridCol As Long
Public mbLookupPatient As Boolean
Private mbExistSelection As Boolean
Private mbooCancel As Boolean
Private mbooGotYou As Boolean
Private mbooUpdate As Boolean
Public mbooNoMessage As Boolean
Private mstrAuthString As String
Private fpDeltaLeft As Long
Private fpDeltaTop As Long
Private flngDeltaX  As Long
Private flngDeltaY  As Long
Private mbooAuthSequence As Boolean
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal Hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long
Private clsWriteAppt As clsWriteAppt
Private Type clsWriteAppt
    CPT_Code As String
    Repeat_Appt As String
    Duration As String
    Appt_ID As Long
    Comments As String
    Confirm As String
    Reason As String
    Referral_Reas As String
    SchedTime As String
    SchedDate As String
    VisitType As String
    SchedType As String
    Add_Proc_CPT As String
    Add_Proc_DESC As String
    Status As String
    Referral As Long
    ProviderID As Long
    FacilityID As Long
    PatientType As String
    PurposeVisit As String
End Type
Private Function Set_Permanent_Blocks()
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim i%
Dim lngDateDiff As Long
Dim strRepDate As String
Dim booSkip As Boolean
Dim strStartRange As String
Dim strEndRange As String
Dim strTestTime As String
Dim strTestDay As String
Dim strWeek As String
Dim lngStartRow As Long
Dim lngEndRow As Long
Dim varCellValue

    On Error GoTo Error_Set_Permanent_Blocks
    Screen.MousePointer = vbHourglass
    If clsWriteAppt.SchedDate = vbNullString Then Exit Function
    strTestTime = DateAdd("d", 1, clsWriteAppt.SchedDate)
    strEndRange = DateAdd("d", -1, clsWriteAppt.SchedDate)
    strWeek = DateAdd("d", -30, clsWriteAppt.SchedDate)
    If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
    strSQL = "Get_Provider_Permanent_Blocks_Prov_Fac_Duration_EndDate " & clsWriteAppt.ProviderID & "," _
        & clsWriteAppt.FacilityID & " ,'Permanent' ,'" & Format(strEndRange, "mm/dd/yyyy") & "'"
    Set rsBlock = cnBlock.Execute(strSQL)
    Do Until rsBlock.EOF
        Clear_Block
        booSkip = False
        strRepDate = vbNullString
        Set Flds = rsBlock.Fields
        For Each Fld In Flds
            If Not IsNull(Fld.Value) Then
                If Trim(Fld.Value) <> vbNullString Then
                    Select Case UCase(Fld.Name)
                        Case "BLOCK_ID"
                            lngPrimKey = Trim(Fld.Value)
                        Case "PROVIDER_ID"
                            udtBlock.ProviderID = Trim(Fld.Value)
                        Case "START_DATE"
                             udtBlock.StartDate = Trim(Fld.Value)
                        Case "END_DATE"
                             udtBlock.EndDate = Trim(Fld.Value)
                        Case "START_TIME"
                             udtBlock.StartTime = Trim(Fld.Value)
                        Case "END_TIME"
                             udtBlock.EndTime = Trim(Fld.Value)
                        Case "REASON"
                             udtBlock.Reason = Trim(Fld.Value)
                        Case "DURATION"
                             udtBlock.Duration = Trim(Fld.Value)
                        Case "FACILITY_ID"
                             udtBlock.FacilityID = Trim(Fld.Value)
                        Case "DAYWEEK"
                             udtBlock.DayWeek = Trim(Fld.Value)
                    End Select
                End If
            End If
        Next
        rsBlock.MoveNext
        lngDateDiff = 0
        If Trim(udtBlock.DayWeek) <> vbNullString Then strRepDate = DayWeek(udtBlock.DayWeek)
        If IsDate(udtBlock.StartDate) And (Trim(udtBlock.DayWeek) = vbNullString Or UCase(udtBlock.DayWeek) = GetDay(clsWriteAppt.SchedDate)) Then
           If Not IsDate(strRepDate) Then strRepDate = udtBlock.EndDate
           If udtBlock.Duration = "Permanent" Then
              lngDateDiff = DateDiff("d", Format(clsWriteAppt.SchedDate, "mm/dd/yyyy"), udtBlock.StartDate)
              If lngDateDiff <= 0 Then
                 lngDateDiff = DateDiff("d", Format(clsWriteAppt.SchedDate, "mm/dd/yyyy"), strRepDate)
                 If lngDateDiff >= 0 Then
                        With grdAvailAppt
                            For lngRow = 1 To .MaxRows
                                .GetText 0, lngRow, varCellValue
                                If UCase(varCellValue) = UCase(udtBlock.StartTime) Then
                                    lngStartRow = lngRow
                                End If
                                .GetText 0, lngRow, varCellValue
                                If UCase(varCellValue) = UCase(udtBlock.EndTime) Then
                                    lngEndRow = lngRow
                                    Exit For
                                End If
                            Next lngRow
                            If lngStartRow > 1 And lngEndRow = 0 Then lngEndRow = lngRow
                            Format_CellTypeEdit_RowSpan grdAvailAppt, lngStartRow, 1, CStr(lngEndRow), "BLOCKED", "TypeHAlignCenter", "BOOKED", GridBlue
                            .SetCellBorder 1, lngRow, 1, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            Format_CellTypeEdit_RowSpan grdAvailAppt, lngStartRow, 2, CStr(lngEndRow), "1", "TypeHAlignCenter", "BOOKED", GridBlue
                            .SetCellBorder 2, lngRow, 2, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End With
                 End If
              End If
           Else
              lngDateDiff = DateDiff("d", udtBlock.StartDate, clsWriteAppt.SchedDate)
              If lngDateDiff >= 0 Then
                 lngDateDiff = DateDiff("d", clsWriteAppt.SchedDate, udtBlock.EndDate)
                 If lngDateDiff >= 0 Then
                        With grdAvailAppt
                            For lngRow = 1 To .MaxRows
                                .GetText 0, lngRow, varCellValue
                                If UCase(varCellValue) = UCase(udtBlock.StartTime) Then
                                    lngStartRow = lngRow
                                End If
                                .GetText 0, lngRow, varCellValue
                                If UCase(varCellValue) = UCase(udtBlock.EndTime) Then
                                    lngEndRow = lngRow
                                    Exit For
                                End If
                            Next lngRow
                            Format_CellTypeEdit_RowSpan grdAvailAppt, lngStartRow, 1, CStr(lngEndRow), "BLOCKED", "TypeHAlignCenter", "BOOKED"
                            .SetCellBorder 1, lngRow, 1, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            Format_CellTypeEdit_RowSpan grdAvailAppt, lngStartRow, 2, CStr(lngEndRow), "1", "TypeHAlignCenter", "BOOKED"
                            .SetCellBorder 2, lngRow, 2, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End With
                 End If
              End If
           End If
        End If
        If udtBlock.Reason = "Holiday" Then
            Exit Do
        End If
        If strTestDay <> udtBlock.DayWeek Then
            strStartRange = vbNullString
            strEndRange = vbNullString
        End If
        lngDateDiff = DateDiff("h", udtBlock.StartTime, udtBlock.EndTime)
        If lngDateDiff > 7 Then
            strStartRange = udtBlock.StartTime
            strEndRange = udtBlock.EndTime
        End If
        strTestDay = udtBlock.DayWeek
    Loop
    If cnBlock.State = adStateOpen Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_Permanent_Blocks:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Set_Permanent_Blocks", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Set_BlockTime()
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim i%
Dim lngDateDiff As Long
Dim strRepDate As String
Dim booSkip As Boolean
Dim strStartRange As String
Dim strEndRange As String
Dim strTestTime As String
Dim strTestDay As String
Dim strWeek As String
Dim lngStartRow As Long
Dim lngEndRow As Long
Dim varCellValue


    On Error GoTo Error_Set_BlockTime
    Screen.MousePointer = vbHourglass
    Set_Permanent_Blocks
    If udtAppointment.SchedDate = vbNullString Then Exit Function
    strTestTime = DateAdd("d", 1, udtAppointment.SchedDate)
    strEndRange = DateAdd("d", -1, udtAppointment.SchedDate)
    strWeek = DateAdd("d", -30, udtAppointment.SchedDate)
    If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
    strSQL = "Get_Provider_Appt_Block_Range " & udtAppointment.ProviderID & ", " & udtAppointment.FacilityID & ", '" _
        & Format(strTestTime, "mm/dd/yyyy") & "', '" & Format(strEndRange, "mm/dd/yyyy") & "', '" & Format(strWeek, "mm/dd/yyyy") & "'"
    Set rsBlock = cnBlock.Execute(strSQL)
    Do Until rsBlock.EOF
        Clear_Block
        booSkip = False
        strRepDate = vbNullString
        Set Flds = rsBlock.Fields
        For Each Fld In Flds
            If Not IsNull(Fld.Value) Then
                If Trim(Fld.Value) <> vbNullString Then
                    Select Case UCase(Fld.Name)
                        Case "BLOCK_ID"
                            lngPrimKey = Trim(Fld.Value)
                        Case "PROVIDER_ID"
                            udtBlock.ProviderID = Trim(Fld.Value)
                        Case "START_DATE"
                             udtBlock.StartDate = Trim(Fld.Value)
                        Case "END_DATE"
                             udtBlock.EndDate = Trim(Fld.Value)
                        Case "START_TIME"
                             udtBlock.StartTime = Trim(Fld.Value)
                        Case "END_TIME"
                             udtBlock.EndTime = Trim(Fld.Value)
                        Case "REASON"
                             udtBlock.Reason = Trim(Fld.Value)
                        Case "DURATION"
                             udtBlock.Duration = Trim(Fld.Value)
                        Case "FACILITY_ID"
                             udtBlock.FacilityID = Trim(Fld.Value)
                        Case "DAYWEEK"
                             udtBlock.DayWeek = Trim(Fld.Value)
                    End Select
                End If
            End If
        Next
        rsBlock.MoveNext
        If strStartRange <> vbNullString Then
            If DateDiff("n", Format(strStartRange, "hh:mm"), Format(udtBlock.StartTime, "hh:mm")) > 0 Then
                booSkip = True
            ElseIf DateDiff("n", Format(strEndRange, "hh:mm"), Format(udtBlock.StartTime, "hh:mm")) > 0 Then
                booSkip = True
            End If
        End If
        If Not booSkip Then
            lngDateDiff = 0
            If Trim(udtBlock.DayWeek) <> vbNullString Then strRepDate = DayWeek(udtBlock.DayWeek)
            If IsDate(udtBlock.StartDate) And (Trim(udtBlock.DayWeek) = vbNullString Or UCase(udtBlock.DayWeek) = GetDay(udtAppointment.SchedDate)) Then
               If Not IsDate(strRepDate) Then strRepDate = udtBlock.EndDate
               If udtBlock.Duration = "Permanent" Then
                  lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), udtBlock.StartDate)
                  If lngDateDiff <= 0 Then
                     lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), strRepDate)
                     If lngDateDiff >= 0 Then
                        With grdAvailAppt
                            For lngRow = 1 To .MaxRows
                                .GetText 0, lngRow, varCellValue
                                If UCase(varCellValue) = UCase(udtBlock.StartTime) Then
                                    lngStartRow = lngRow
                                End If
                                .GetText 0, lngRow, varCellValue
                                If UCase(varCellValue) = UCase(udtBlock.EndTime) Then
                                    lngEndRow = lngRow
                                    Exit For
                                End If
                            Next lngRow
                            Format_CellTypeEdit_RowSpan grdAvailAppt, lngStartRow, 1, CStr(lngEndRow), "BLOCKED", "TypeHAlignCenter", "BOOKED", GridBlue
                            .SetCellBorder 1, lngRow, 1, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            Format_CellTypeEdit_RowSpan grdAvailAppt, lngStartRow, 2, CStr(lngEndRow), "1", "TypeHAlignCenter", "BOOKED", GridBlue
                            .SetCellBorder 2, lngRow, 2, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End With
                     End If
                  End If
               Else
                  lngDateDiff = DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate)
                  If lngDateDiff >= 0 Then
                     lngDateDiff = DateDiff("d", udtAppointment.SchedDate, udtBlock.EndDate)
                     If lngDateDiff >= 0 Then
                        With grdAvailAppt
                            For lngRow = 1 To .MaxRows
                                .GetText 0, lngRow, varCellValue
                                If UCase(varCellValue) = UCase(udtBlock.StartTime) Then
                                    lngStartRow = lngRow
                                End If
                                .GetText 0, lngRow, varCellValue
                                If UCase(varCellValue) = UCase(udtBlock.EndTime) Then
                                    lngEndRow = lngRow
                                    Exit For
                                End If
                            Next lngRow
                            If lngStartRow > 1 And lngEndRow = 0 Then lngEndRow = lngRow
                            Format_CellTypeEdit_RowSpan grdAvailAppt, lngStartRow, 1, CStr(lngEndRow), "BLOCKED", "TypeHAlignCenter", "BOOKED"
                            .SetCellBorder 1, lngRow, 1, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            Format_CellTypeEdit_RowSpan grdAvailAppt, lngStartRow, 2, CStr(lngEndRow), "1", "TypeHAlignCenter", "BOOKED"
                            .SetCellBorder 2, lngRow, 2, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End With
                        If DateDiff("h", Format(udtBlock.StartTime, "hh:mm"), Format(udtBlock.EndTime, "hh:mm")) > 7 And _
                            DateDiff("d", Format(udtBlock.StartDate, "mm/dd/yyyy"), Format(udtBlock.EndDate, "mm/dd/yyyy")) >= 0 Then
                            Exit Do
                        End If
                     End If
                  End If
               End If
            End If
            If udtBlock.Reason = "Holiday" Then
                Exit Do
            End If
            If strTestDay <> udtBlock.DayWeek Then
                strStartRange = vbNullString
                strEndRange = vbNullString
            End If
            lngDateDiff = DateDiff("h", udtBlock.StartTime, udtBlock.EndTime)
            If lngDateDiff > 7 Then
                strStartRange = udtBlock.StartTime
                strEndRange = udtBlock.EndTime
            End If
            strTestDay = udtBlock.DayWeek
        End If
    Loop
    If cnBlock.State = adStateOpen Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_BlockTime:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Set_BlockTime", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

 
Private Function Check_OverBook_Change(Optional vstrTime As String) As Boolean
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim lngProviderID As Long
Dim sValue As String
Dim intCount As Integer
Dim iAnswer As Integer
Dim strTime As String
Dim strApptDate As String
Dim lngFacility As Long
Dim i%

    On Error GoTo Error_Check_OverBook_Change
    Screen.MousePointer = vbHourglass
    If clsWriteAppt.Duration = vbNullString Then clsWriteAppt.Duration = Trim(cboDuration)
    If vstrTime = vbNullString Then
        strTime = ccdTime.Value
        strTime = Format(strTime, "nn")
        strTime = DateAdd("n", Format(strTime, "nn") + 1, ccdTime.Value)
        sValue = Replace((clsWriteAppt.Duration), "minutes", "")
        sValue = DateAdd("n", Val(sValue) - 1, ccdTime.Value)
        strApptDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    Else
        strTime = clsWriteAppt.SchedTime
        strTime = Format(strTime, "nn")
        strTime = DateAdd("n", Format(strTime, "nn") + 1, clsWriteAppt.SchedTime)
        sValue = Replace((clsWriteAppt.Duration), "minutes", "")
        sValue = DateAdd("n", Val(sValue) - 1, clsWriteAppt.SchedTime)
        strApptDate = Format(clsWriteAppt.SchedDate, "mm/dd/yyyy")
    End If
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    strSQL = "SELECT * FROM APPOINTMENTS WHERE PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
        & "FACILITY_ID = " & udtAppointment.FacilityID & " AND  " _
        & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & Format(strApptDate, "mm/dd/yyyy") & "' AND " _
        & "CONVERT(DATETIME, APPT_TIME, 108)  BETWEEN '" & Format(strTime, "hh:mm AM/PM") & "' AND '" & Format(sValue, "hh:mm AM/PM") & "' AND ACTIVE = 1 "
    Set rsAppt = cnAppt.Execute(strSQL)
    intCount = 0
        With rsAppt
            Do Until .EOF
                If Not IsNull(!Patient_ID) Then intCount = intCount + 1
                .MoveNext
            Loop
        End With
    If intCount > 0 Then
        Read_OverBook
        If intCount >= Val(clsApptPref.OverbookNumber) Then
            MsgBox "This appointment time exceeds practice policy for overbooking and cannot be scheduled; please select another time and/or another duration.", vbInformation, "PatientTrac Scheduling"
            Check_OverBook_Change = True
        End If
    End If
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
     Exit Function
Error_Check_OverBook_Change:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Check_OverBook_Change", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Check_Repeat_Cancel(vlngRepeat_ID As Long) As String
Dim lngRepId As Long
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset

    On Error GoTo Error_Check_Repeat_Cancel
    Screen.MousePointer = vbHourglass
    Check_Repeat_Cancel = vbNullString
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    strSQL = "Get_Repeat_Cancel " & vlngRepeat_ID
        Set rsAppt = cnAppt.Execute(strSQL)
        With rsAppt
            If Not .EOF Then
                Check_Repeat_Cancel = "Cancelled"
            End If
        End With
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Check_Repeat_Cancel:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Check_Repeat_Cancel", App.EXEName & "." & "ModConst"
    Exit Function
    Resume Next
End Function

Private Function Clear_clsWriteAppt()
    With clsWriteAppt
        .CPT_Code = vbNullString
        .Repeat_Appt = vbNullString
        .Reason = vbNullString
        .Referral = 0
        .Appt_ID = 0
        .Confirm = vbNullString
        .Comments = vbNullString
        .Referral_Reas = vbNullString
        .SchedTime = vbNullString
        .SchedDate = vbNullString
        .VisitType = vbNullString
        .SchedType = vbNullString
        .Add_Proc_CPT = vbNullString
        .Add_Proc_DESC = vbNullString
        .Status = vbNullString
        .ProviderID = 0
        .FacilityID = 0
        .PatientType = vbNullString
    End With
End Function

Function Format_Appt_Lookup()
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim strEnd As String
Dim lngDateDiff As Long
Dim lngStartRow  As Long
Dim varCellValue


    On Error GoTo Error_Format_Appt_Lookup
    Screen.MousePointer = vbHourglass
    Get_TimeSlot
    With grdAvailAppt
        .ColHeaderRows = 4
        .Redraw = False
        .ClearRange 1, 1, .MaxCols, .MaxRows, True
        .MaxRows = 0
        .MaxCols = 2
        .ColWidth(0) = 7
        .ColWidth(1) = 31.4
        .ColWidth(2) = 16
        For Idx = mStartIdx To UBound(mStartHour)
            lngRow = .MaxRows + 1
            .MaxRows = lngRow
            Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", mStartHour(Idx) & " AM", "TypeHAlignLeft", mStartHour(Idx) & " AM", HeaderColor
            grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Select Case UCase(clsApptPref.SchedInterval)
                Case "EVERY HALF HOUR", "EVERY 30 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, lngCol, "0", mStartHour(Idx) & ":30 AM", "TypeHAlignLeft", mStartHour(Idx) & " AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 15 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":15 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":15 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 10 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":10 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":15 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":20 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":40 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":50 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 5 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":05 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":15 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":10 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":15 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":20 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":25 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":35 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":40 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":50 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":55 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 20 MIN"
                    lngRow = .MaxRows + 1
                   .MaxRows = lngRow
                   Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":20 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":20 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                   .MaxRows = lngRow
                   Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":40 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":40 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End Select
            grdAvailAppt.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdAvailAppt.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Next Idx
        lngStartRow = lngRow
        For Idx = LBound(mEndHour) To Left(mEndHour(mEndIdx), 2)
            lngRow = .MaxRows + 1
            .MaxRows = lngRow
            Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", mEndHour(Idx) & " PM", "TypeHAlignLeft", mEndHour(Idx) & " PM", HeaderColor
            grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Select Case UCase(clsApptPref.SchedInterval)
                Case "EVERY HALF HOUR", "EVERY 30 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 15 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":15 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":15 PM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":45 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 PM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 10 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":10 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":15 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":20 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":30 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":40 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":50 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 5 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":05 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":15 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":10 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":30 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":15 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":20 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":25 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":35 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":40 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":45 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":50 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":55 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 20 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":20 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":20 PM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdAvailAppt, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":40 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":40 PM", HeaderColor
                    grdAvailAppt.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End Select
            grdAvailAppt.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdAvailAppt.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Next Idx
        .Row = 1
        .LeftCol = 1
        For Idx = lngStartRow To grdAvailAppt.MaxRows
        lngRow = Idx
        .GetText 0, lngRow, varCellValue
            strEnd = mEndHour(mEndIdx)
            If Trim(varCellValue) <> vbNullString Then
                lngDateDiff = DateTime.DateDiff("n", Format(varCellValue, "hh:mm"), Format(mEndHour(mEndIdx) & "PM", "hh:mm"))
                If lngDateDiff < 0 Then
                     grdAvailAppt.DeleteRows lngRow, 1
                    .MaxRows = .MaxRows - 1
                    Idx = Idx - 1
                End If
            End If
        Next Idx
        .Redraw = True
    End With
    Load_ApptTime
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_Appt_Lookup:
    grdAvailAppt.Visible = True
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Appt_Lookup", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)

    With vObjGrid
        vlngRow = vlngRow
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = vlngRow
            .Row2 = vlngRow
            .BackColor = &H800000
            .RowHeight(vlngRow) = 3
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, 16, GridBlue, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub
Private Function Format_Plan()
Dim lngCol As Long
Dim lngRow As Long

    Screen.MousePointer = vbHourglass
    grdPlan.ProcessTab = True
    grdPlan.FontBold = True
    grdPlan.ForeColor = 0
    Get_PatientName piPatientID
    With grdPlan
        .Redraw = False
        .MaxRows = 0
        .MaxCols = 3
        .ColHeaderRows = 4
        .ClearRange 1, 1, grdPlan.MaxCols, grdPlan.MaxRows, False
        .ColWidth(1) = 15
        .ColWidth(2) = 13.5
        .ColWidth(3) = 15
        Format_CellTypeEdit grdPlan, SpreadHeader, 1, "9", "FOLLOW-UP PLAN FOR" & vbNewLine & UCase(udtPatientName.Name), "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdPlan.SetCellBorder 1, SpreadHeader, 9, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdPlan
        lngCol = 1
        Format_CellTypeStatic grdPlan, SpreadHeader + 2, lngCol, "1", "ENCOUNTER DATE", "TypeHAlignCenter", "ENCOUNTER_DATE", HeaderColor
        grdPlan.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeStatic grdPlan, SpreadHeader + 2, lngCol, "1", "FOLLOW-UP", "TypeHAlignCenter", "FOLLOW_UP", HeaderColor
        grdPlan.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeStatic grdPlan, SpreadHeader + 2, lngCol, "1", "APPT TYPE", "TypeHAlignCenter", "PATIENT_TYPE", HeaderColor
        grdPlan.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdPlan
        .Redraw = True
    End With
    Read_Plan
    Screen.MousePointer = vbDefault
    Exit Function
End Function
Private Function Format_PurposeGrid()
Dim lngCol As Long
Dim lngRow As Long

    Screen.MousePointer = vbHourglass
    grdProcedure.ProcessTab = True
    grdProcedure.FontBold = True
    grdProcedure.ForeColor = 0
    With grdProcedure
        .Redraw = False
        .MaxRows = 0
        .MaxCols = 5
        .ColHeaderRows = 4
        .ClearRange 1, 1, grdProcedure.MaxCols, grdProcedure.MaxRows, False
        .ColWidth(1) = 23.7
        .ColWidth(2) = 10
        .ColWidth(3) = 23
        .ColWidth(4) = 11.9
        .ColWidth(5) = 0
        Format_CellTypeEdit grdProcedure, SpreadHeader, 1, "4", "SCHEDULED VISIT CODES", "TypeHAlignCenter", "HEADER", HeaderColor
        grdProcedure.SetCellBorder 1, SpreadHeader, 4, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdProcedure

        If IsProperCase(udtAppointment.SchedType) = "Patient" Or udtAppointment.SchedType = vbNullString Then
            lngCol = 1
            Format_CellTypeStatic grdProcedure, SpreadHeader + 2, lngCol, "1", "VISIT TYPE", "TypeHAlignCenter", "VISIT TYPE", HeaderColor
            grdProcedure.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 2
            Format_CellTypeStatic grdProcedure, SpreadHeader + 2, lngCol, "1", "CPT CODE", "TypeHAlignCenter", "CPT_CODE", HeaderColor
            grdProcedure.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 3
            Format_CellTypeStatic grdProcedure, SpreadHeader + 2, lngCol, "1", "ADD-ON DESCRIPTION", "TypeHAlignCenter", "ADD_DESC", HeaderColor
            grdProcedure.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 4
            Format_CellTypeStatic grdProcedure, SpreadHeader + 2, lngCol, "1", "ADD-ON CODE", "TypeHAlignCenter", "ADD_CPT_CODE", HeaderColor
            grdProcedure.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Else
            lngCol = 1
            Format_CellTypeStatic grdProcedure, SpreadHeader + 2, lngCol, "2", "MEETING TYPE", "TypeHAlignCenter", "MEETING TYPE", HeaderColor
            grdProcedure.SetCellBorder lngCol, SpreadHeader + 2, 2, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 3
            Format_CellTypeStatic grdProcedure, SpreadHeader + 2, lngCol, "2", "DURATION", "TypeHAlignCenter", "DURATION", HeaderColor
            grdProcedure.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        lngCol = 5
        Format_CellTypeStatic grdProcedure, SpreadHeader + 2, lngCol, "1", "FACILITY TYPE", "TypeHAlignCenter", "FACILTY_TYPE", HeaderColor
        grdProcedure.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdProcedure
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Function
End Function

Private Function Get_Appt_Providers()
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
Dim booSkip As Boolean
Dim lngProvID As Long
Dim i%
Dim lngFacID As Long

    On Error GoTo Error_Get_Appt_Providers
    Screen.MousePointer = vbHourglass
    cboProvider.Clear
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    If cboFacility.listcount <= 0 Then
        Load_Facilities
    End If
    For i = 0 To cboFacility.listcount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            udtFacility.FacID = cboFacility.ItemData(i)
            Exit For
        End If
    Next i
    strSQL = "Get_Appt_Providers " & udtFacility.FacID
'    "SELECT PROVIDERS.*, APPOINTMENT_PROVIDERS.PROVIDER_ID FROM APPOINTMENT_PROVIDERS " _
'        & "INNER JOIN PROVIDERS ON APPOINTMENT_PROVIDERS.PROVIDER_ID = PROVIDERS.PROVIDER_ID " _
'        & "WHERE APPOINTMENT_PROVIDERS.FACILITY_ID = " & udtFacility.FacID & " AND " _
'        & "PROVIDERS.ACTIVE = 1 ORDER BY PROVIDERS.LAST_NAME, PROVIDERS.FIRST_NAME"
     Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        Do Until .EOF
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            lngProvID = 0
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "FIRST_NAME"
                                strFirst = Fld.Value
                            Case "MIDDLE_NAME"
                                strMiddle = Fld.Value
                            Case "LAST_NAME"
                                strLast = Fld.Value
                            Case "SUFFIX"
                                strSuffix = Fld.Value
                            Case "TITLE"
                                strTitle = Fld.Value
                            Case "PROVIDER_ID"
                                lngProvID = Fld.Value
                        End Select
                    End If
                End If
            Next
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
                 cboProvider.AddItem clsProvider.Name
                If IsNumeric(lngProvID) Then cboProvider.ItemData(cboProvider.NewIndex) = lngProvID
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
Error_Get_Appt_Providers:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Appt_Providers", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAlign As String, _
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
                .FontSize = 8
                .FontBold = True
                .ForeColor = 0
                If vlngBackColor > 0 Then .BackColor = vlngBackColor
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                  .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                ElseIf Trim(vstrSpan) <> vbNullString Then
                  lngSpan = vstrSpan
                  If lngSpan > 1 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                End If
                .SetText vlngCol, vlngRow, Trim(vstrText)
                .Text = Trim(vstrText)
                CellNote = .MaxTextCellHeight
                If .RowHeight(vlngRow) < CellNote Then
                  .RowHeight(vlngRow) = CellNote + 2
                End If
                If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
                .TypeVAlign = TypeVAlignCenter
                If vstrHAlign = "TypeHAlignCenter" Then
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .TypeHAlign = TypeHAlignLeft
                End If
                If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAlign As String, Optional vstrCellTag As String, Optional lngBackColor As Long, Optional vstrCellNote As String)
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
            .ForeColor = 0
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
'            If vstrVAlign = "TypeVAlignTop" Then
'                .TypeVAlign = TypeVAlignTop
'            ElseIf vstrVAlign = "TypeVAlignBottom" Then
'                .TypeVAlign = TypeVAlignBottom
'            Else
'
'            End If
            .TypeVAlign = TypeVAlignCenter
            If vstrHAlign = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            ElseIf vstrHAlign = "TypeHAlignRight" Then
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
 
Private Function Get_Count() As String
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field

    On Error GoTo Error_Get_Count
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    strSQL = "Select COUNT(APPT_ID) AS NUMBER FROM APPOINTMENTS WHERE " _
        & "APPT_DATE = '" & Format(ccSchedDate.Value, "mm/dd/yyyy") & "' AND " _
        & "PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
        & "FACILITY_ID = " & udtAppointment.FacilityID & " AND ACTIVE = 1 AND " _
        & "CONVERT(DATETIME, APPT_TIME, 108) BETWEEN '" & udtTime.StartTime & "' AND '" & udtTime.EndTime & "'"
    Set rsAppt = cnAppt.Execute(strSQL)
    With rsAppt
        If Not .EOF Then
            If Not IsNull(!Number) Then Get_Count = (!Number)
        End If
    End With
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
     Exit Function
Error_Get_Count:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Count", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Get_Grid_Status() As String
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim strTerms As String

    On Error GoTo Error_Get_Grid_Status
    Screen.MousePointer = vbHourglass
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Get_Attr_Valid_Value_SubDom_Alphabetic 'sd_Appointment_Status'"
    Set rsTerms = cnTerms.Execute(strSQL)
    Do Until rsTerms.EOF
        strTerms = vbNullString
        strTerms = IsProperCase(rsTerms.Fields("ATTRBTE_VALID_VALUE"))
        rsTerms.MoveNext
        If Get_Grid_Status = vbNullString Then
            Get_Grid_Status = strTerms
        Else
            Get_Grid_Status = Get_Grid_Status & Chr(9) & strTerms
        End If
    Loop
    Get_Grid_Status = Get_Grid_Status & Chr(9) & "Active"
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Grid_Status:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Grid_Status", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Get_GridDuration() As String
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset

    On Error GoTo Error_Get_GridDuration
    Screen.MousePointer = vbHourglass
    Get_GridDuration = vbNullString
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(1)
    strSQL = "Get_Appt_Types_Time"
    Set rsAppt = cnAppt.Execute(strSQL)
    Do Until rsAppt.EOF
        If rsAppt.Fields("APPT_TIME") <> vbNullString Then
            If Not TextFound("minutes", rsAppt.Fields("APPT_TIME"), Len("minutes"), True) Then
                If Get_GridDuration = vbNullString Then
                    Get_GridDuration = rsAppt.Fields("APPT_TIME") & " " & "minutes"
                Else
                    Get_GridDuration = Get_GridDuration & Chr(9) & rsAppt.Fields("APPT_TIME") & " " & "minutes"
                End If
            Else
                If Get_GridDuration = vbNullString Then
                    Get_GridDuration = rsAppt.Fields("APPT_TIME")
                Else
                    Get_GridDuration = Get_GridDuration & Chr(9) & rsAppt.Fields("APPT_TIME")
                End If
            End If
         End If
        rsAppt.MoveNext
    Loop
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_GridDuration:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_GridDuration", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Get_VisitGrid(vstrType As String) As String
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSerPlace As String
Dim strPatType As String
Dim strTitle As String
Dim i%
Dim Arr
Dim strProvider As String
Dim strCompare As String
Dim lngFacID As Long

    On Error GoTo Error_Get_VisitGrid
    For i = 0 To cboFacility.listcount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            lngFacID = cboFacility.ItemData(i)
            Exit For
        End If
    Next i
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    strCompare = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
    strCompare = Replace(strCompare, ".", "")
    strCompare = Trim(strCompare)
    If udtAppointment.ProviderID > 0 Then Get_Providers udtAppointment.ProviderID
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    If Trim(vstrType) <> vbNullString Then
        strSQL = "Get_Appt_Types_Visit_Type '" & vstrType & "%'"
    Else
        strSQL = "Get_Appt_Types_Facility_ID " & lngFacID
    End If
    Set rsCodes = cnCodes.Execute(strSQL)
    Do Until rsCodes.EOF
        strProvider = vbNullString
        If Not IsNull(rsCodes.Fields("PROVIDER_TITLE")) Then
            strProvider = UCase(rsCodes.Fields("PROVIDER_TITLE"))
        End If
        strProvider = Replace(strProvider, ".", "")
        strProvider = Trim(strProvider)
        If Trim(strProvider) <> vbNullString Then
            If strProvider = strCompare Or IsProperCase(strProvider) = "All" Then
                Set Flds = rsCodes.Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "APPT_CODE"
                                    If Get_VisitGrid = vbNullString Then
                                        Get_VisitGrid = Trim(Fld.Value)
                                    Else
                                        Get_VisitGrid = Get_VisitGrid & Chr(9) & Trim(Fld.Value)
                                    End If
                            End Select
                        End If
                    End If
                Next
            End If
        End If
        rsCodes.MoveNext
    Loop
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
     Exit Function
Error_Get_VisitGrid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_VisitGrid", App.EXEName & " " & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Is_NewAppt() As Boolean
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim strSQL As String
Dim Idx%

    On Error GoTo Error_Is_NewAppt
    Screen.MousePointer = vbHourglass
    If udtAppointment.FacilityID = 0 Then
        For Idx = 0 To cboFacility.listcount - 1
            udtAppointment.FacilityID = cboFacility.ItemData(Idx)
            Exit For
        Next Idx
    End If
    If udtAppointment.ProviderID = 0 Then
        For Idx = 0 To cboProvider.listcount - 1
            udtAppointment.ProviderID = cboProvider.ItemData(Idx)
            Exit For
        Next Idx
    End If
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
        strSQL = "Validatae_Is_NewAppt " & Val(udtSchedPatient.IntId) & "," & udtAppointment.ProviderID & "," _
            & udtAppointment.FacilityID & ",'" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "','" & Format(udtAppointment.SchedTime, "hh:mm AM/PM") & "'"
        Set rsAppt = cnAppt.Execute(strSQL)
        With rsAppt
            If Not .EOF Then
                Is_NewAppt = True
            End If
        End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_Is_NewAppt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Is_NewAppt", App.EXEName & "." & Me.Name
    Exit Function
End Function

Private Function Load_TimeGrid() As String

    Select Case UCase(clsApptPref.SchedInterval)
        Case "EVERY 15 MIN"
           Load_TimeGrid = "07:00 AM" & Chr(9) & "07:15 AM" & Chr(9) & "07:30 AM" & Chr(9) & "07:45 AM" & Chr(9) & "08:00 AM" & Chr(9) & "08:15 AM" & Chr(9) & "08:30 AM" & Chr(9) & "08:45 AM" _
           & Chr(9) & "09:00 AM" & Chr(9) & "09:15 AM" & Chr(9) & "09:30 AM" & Chr(9) & "09:45 AM" & Chr(9) & "10:00 AM" & Chr(9) & "10:15 AM" & Chr(9) & "10:30 AM" & Chr(9) & "10:45 AM" _
             & Chr(9) & "11:00 AM" & Chr(9) & "11:15 AM" & Chr(9) & "11:30 AM" & Chr(9) & "11:45 AM" & Chr(9) & "12:00 PM" & Chr(9) & "12:15 PM" & Chr(9) & "12:30 PM" & Chr(9) & "12:45 PM" _
             & Chr(9) & "01:00 PM" & Chr(9) & "01:15 PM" & Chr(9) & "01:30 PM" & Chr(9) & "01:45 PM" & Chr(9) & "02:00 PM" & Chr(9) & "02:15 PM" & Chr(9) & "02:30 PM" & Chr(9) & "02:45 PM" _
             & Chr(9) & "03:00 PM" & Chr(9) & "03:15 PM" & Chr(9) & "03:30 PM" & Chr(9) & "03:45 PM" & Chr(9) & "04:00 PM" & Chr(9) & "04:15 PM" & Chr(9) & "04:30 PM" & Chr(9) & "04:45 PM" _
             & Chr(9) & "05:00 PM" & Chr(9) & "05:15 PM" & Chr(9) & "05:30 PM" & Chr(9) & "05:45 PM" & Chr(9) & "06:00 PM" & Chr(9) & "06:15 PM" & Chr(9) & "06:30 PM" & Chr(9) & "06:45 PM" _
             & Chr(9) & "07:00 PM" & Chr(9) & "07:15 PM" & Chr(9) & "07:30 PM" & Chr(9) & "07:45 PM" & Chr(9) & "08:00 PM" & Chr(9) & "08:15 PM" & Chr(9) & "08:30 PM" & Chr(9) & "08:45 PM"
       Case "EVERY 5 MIN"
           Load_TimeGrid = "07:00 AM" & Chr(9) & "07:05 AM" & Chr(9) & "07:10 AM" & Chr(9) & "07:15 AM" & Chr(9) & "07:20 AM" & Chr(9) & "07:25 AM" & Chr(9) & "07:30 AM" & Chr(9) & "07:35 AM" _
             & Chr(9) & "07:40 AM" & Chr(9) & "07:45 AM" & Chr(9) & "07:50 AM" & Chr(9) & "07:55 AM" & Chr(9) & "08:00 AM" & Chr(9) & "08:05 AM" & Chr(9) & "08:10 AM" & Chr(9) & "08:15 AM" _
             & Chr(9) & "08:20 AM" & Chr(9) & "08:25 AM" & Chr(9) & "08:30 AM" & Chr(9) & "08:35 AM" & Chr(9) & "08:40 AM" & Chr(9) & "08:45 AM" & Chr(9) & "08:50 AM" & Chr(9) & "08:55 AM" _
             & Chr(9) & "09:00 AM" & Chr(9) & "09:05 AM" & Chr(9) & "09:10 AM" & Chr(9) & "09:15 AM" & Chr(9) & "09:20 AM" & Chr(9) & "09:25 AM" & Chr(9) & "09:30 AM" & Chr(9) & "09:35 AM" _
             & Chr(9) & "09:40 AM" & Chr(9) & "09:45 AM" & Chr(9) & "09:50 AM" & Chr(9) & "09:55 AM" & Chr(9) & "10:00 AM" & Chr(9) & "10:05 AM" & Chr(9) & "10:10 AM" & Chr(9) & "10:15 AM" _
             & Chr(9) & "10:20 AM" & Chr(9) & "10:25 AM" & Chr(9) & "10:30 AM" & Chr(9) & "10:35 AM" & Chr(9) & "10:40 AM" & Chr(9) & "10:45 AM" & Chr(9) & "10:50 AM" & Chr(9) & "10:55 AM" _
             & Chr(9) & "11:00 AM" & Chr(9) & "11:05 AM" & Chr(9) & "11:10 AM" & Chr(9) & "11:15 AM" & Chr(9) & "11:20 AM" & Chr(9) & "11:25 AM" & Chr(9) & "11:30 AM" & Chr(9) & "11:35 AM" _
             & Chr(9) & "11:40 AM" & Chr(9) & "11:45 AM" & Chr(9) & "11:50 AM" & Chr(9) & "11:55 AM" & Chr(9) & "12:00 PM" & Chr(9) & "12:05 PM" & Chr(9) & "12:10 PM" & Chr(9) & "12:15 PM" _
             & Chr(9) & "12:20 PM" & Chr(9) & "12:25 PM" & Chr(9) & "12:30 PM" & Chr(9) & "12:35 PM" & Chr(9) & "12:40 PM" & Chr(9) & "12:45 PM" & Chr(9) & "12:50 PM" & Chr(9) & "12:55 PM" _
             & Chr(9) & "01:00 PM" & Chr(9) & "01:05 PM" & Chr(9) & "01:10 PM" & Chr(9) & "01:15 PM" _
             & Chr(9) & "01:20 PM" & Chr(9) & "01:25 PM" & Chr(9) & "01:30 PM" & Chr(9) & "01:35 PM" & Chr(9) & "01:40 PM" & Chr(9) & "01:45 PM" & Chr(9) & "01:50 PM" & Chr(9) & "01:55 PM" _
             & Chr(9) & "02:00 PM" & Chr(9) & "02:05 PM" & Chr(9) & "02:10 PM" & Chr(9) & "02:15 PM" & Chr(9) & "02:20 PM" & Chr(9) & "02:25 PM" & Chr(9) & "02:30 PM" & Chr(9) & "02:35 PM" _
             & Chr(9) & "02:40 PM" & Chr(9) & "02:45 PM" & Chr(9) & "02:50 PM" & Chr(9) & "02:55 PM" & Chr(9) & "03:00 PM" & Chr(9) & "03:05 PM" & Chr(9) & "03:10 PM" & Chr(9) & "03:15 PM" _
             & Chr(9) & "03:20 PM" & Chr(9) & "03:25 PM" & Chr(9) & "03:30 PM" & Chr(9) & "03:35 PM" & Chr(9) & "03:40 PM" & Chr(9) & "03:45 PM" & Chr(9) & "03:50 PM" & Chr(9) & "03:55 PM" _
             & Chr(9) & "04:00 PM" & Chr(9) & "04:05 PM" & Chr(9) & "04:10 PM" & Chr(9) & "04:15 PM" & Chr(9) & "04:20 PM" & Chr(9) & "04:25 PM" & Chr(9) & "04:30 PM" & Chr(9) & "04:35 PM" _
             & Chr(9) & "04:40 PM" & Chr(9) & "04:45 PM" & Chr(9) & "04:50 PM" & Chr(9) & "04:55 PM" & Chr(9) & "05:00 PM" & Chr(9) & "05:05 PM" & Chr(9) & "05:10 PM" & Chr(9) & "05:15 PM" _
             & Chr(9) & "05:20 PM" & Chr(9) & "05:25 PM" & Chr(9) & "05:30 PM" & Chr(9) & "05:35 PM" & Chr(9) & "05:40 PM" & Chr(9) & "05:45 PM" & Chr(9) & "05:50 PM" & Chr(9) & "05:55 PM" _
             & Chr(9) & "06:00 PM" & Chr(9) & "06:05 PM" & Chr(9) & "06:10 PM" & Chr(9) & "06:15 PM" & Chr(9) & "06:20 PM" & Chr(9) & "06:25 PM" & Chr(9) & "06:30 PM" & Chr(9) & "06:35 PM" _
             & Chr(9) & "06:40 PM" & Chr(9) & "06:45 PM" & Chr(9) & "06:50 PM" & Chr(9) & "06:55 PM" & Chr(9) & "07:00 PM" & Chr(9) & "07:05 PM" & Chr(9) & "07:10 PM" & Chr(9) & "07:15 PM" _
             & Chr(9) & "07:20 PM" & Chr(9) & "07:25 PM" & Chr(9) & "07:30 PM" & Chr(9) & "07:35 PM" & Chr(9) & "07:40 PM" & Chr(9) & "07:45 PM" & Chr(9) & "07:50 PM" & Chr(9) & "07:55 PM" _
             & Chr(9) & "08:00 PM" & Chr(9) & "08:05 PM" & Chr(9) & "08:10 PM" & Chr(9) & "08:15 PM" & Chr(9) & "08:20 PM" & Chr(9) & "08:25 PM" & Chr(9) & "08:30 PM" & Chr(9) & "08:35 PM" _
             & Chr(9) & "08:40 PM" & Chr(9) & "08:45 PM" & Chr(9) & "08:50 PM" & Chr(9) & "08:55 PM"
        Case "EVERY 20 MIN"
           Load_TimeGrid = "07:00 AM" & Chr(9) & "07:20 AM" & Chr(9) & "07:40 AM" & Chr(9) & "08:00 AM" & Chr(9) & "08:20 AM" & Chr(9) & "08:40 AM" _
             & Chr(9) & "09:00 AM" & Chr(9) & "09:20 AM" & Chr(9) & "09:40 AM" & Chr(9) & "10:00 AM" & Chr(9) & "10:20 AM" & Chr(9) & "10:40 AM" _
             & Chr(9) & "11:00 AM" & Chr(9) & "11:20 AM" & Chr(9) & "11:40 AM" & Chr(9) & "12:00 PM" & Chr(9) & "12:20 PM" & Chr(9) & "12:40 PM" _
             & Chr(9) & "01:00 PM" & Chr(9) & "01:20 PM" & Chr(9) & "01:40 PM" & Chr(9) & "02:00 PM" & Chr(9) & "02:20 PM" & Chr(9) & "02:40 PM" _
             & Chr(9) & "03:00 PM" & Chr(9) & "03:20 PM" & Chr(9) & "03:40 PM" & Chr(9) & "04:00 PM" & Chr(9) & "04:20 PM" & Chr(9) & "04:40 PM" _
             & Chr(9) & "05:00 PM" & Chr(9) & "05:20 PM" & Chr(9) & "05:40 PM" & Chr(9) & "06:00 PM" & Chr(9) & "06:20 PM" & Chr(9) & "06:40 PM" _
             & Chr(9) & "07:00 PM" & Chr(9) & "07:20 PM" & Chr(9) & "07:40 PM" & Chr(9) & "08:00 PM" & Chr(9) & "08:20 PM" & Chr(9) & "08:40 PM"
       Case "EVERY 10 MIN"
        Load_TimeGrid = "07:00 AM" & Chr(9) & "07:10 AM" & Chr(9) & "07:20 AM" & Chr(9) & "07:30 AM" & Chr(9) & "07:40 AM" & Chr(9) & "07:50 AM" _
          & Chr(9) & "08:00 AM" & Chr(9) & "08:10 AM" & Chr(9) & "08:20 AM" & Chr(9) & "08:30 AM" & Chr(9) & "08:40 AM" & Chr(9) & "08:50 AM" _
          & Chr(9) & "09:00 AM" & Chr(9) & "09:10 AM" & Chr(9) & "09:20 AM" & Chr(9) & "09:30 AM" & Chr(9) & "09:40 AM" & Chr(9) & "09:50 AM" _
          & Chr(9) & "10:00 AM" & Chr(9) & "10:10 AM" & Chr(9) & "10:20 AM" & Chr(9) & "10:30 AM" & Chr(9) & "10:40 AM" & Chr(9) & "10:50 AM" _
          & Chr(9) & "11:00 AM" & Chr(9) & "11:10 AM" & Chr(9) & "11:20 AM" & Chr(9) & "11:30 AM" & Chr(9) & "11:40 AM" & Chr(9) & "11:50 AM" _
          & Chr(9) & "12:00 PM" & Chr(9) & "12:10 PM" & Chr(9) & "12:20 PM" & Chr(9) & "12:30 PM" & Chr(9) & "12:40 PM" & Chr(9) & "12:50 PM" _
          & Chr(9) & "01:00 PM" & Chr(9) & "01:10 PM" & Chr(9) & "01:20 PM" & Chr(9) & "01:30 PM" & Chr(9) & "01:40 PM" & Chr(9) & "01:50 PM" _
          & Chr(9) & "02:00 PM" & Chr(9) & "02:10 PM" & Chr(9) & "02:20 PM" & Chr(9) & "02:30 PM" & Chr(9) & "02:40 AM" & Chr(9) & "02:50 PM" _
          & Chr(9) & "03:00 PM" & Chr(9) & "03:10 PM" & Chr(9) & "03:20 PM" & Chr(9) & "03:30 PM" & Chr(9) & "03:40 PM" & Chr(9) & "03:50 PM" _
          & Chr(9) & "04:00 PM" & Chr(9) & "04:10 PM" & Chr(9) & "04:20 PM" & Chr(9) & "04:30 PM" & Chr(9) & "04:40 PM" & Chr(9) & "04:50 PM" _
          & Chr(9) & "05:00 PM" & Chr(9) & "05:10 PM" & Chr(9) & "05:20 PM" & Chr(9) & "05:30 PM" & Chr(9) & "05:40 PM" & Chr(9) & "05:50 PM" _
          & Chr(9) & "06:00 PM" & Chr(9) & "06:10 PM" & Chr(9) & "06:20 PM" & Chr(9) & "06:30 PM" & Chr(9) & "06:40 PM" & Chr(9) & "06:50 PM" _
          & Chr(9) & "07:00 PM" & Chr(9) & "07:10 PM" & Chr(9) & "07:20 PM" & Chr(9) & "07:30 PM" & Chr(9) & "07:40 PM" & Chr(9) & "07:50 PM" _
          & Chr(9) & "08:00 PM" & Chr(9) & "08:10 PM" & Chr(9) & "08:20 PM" & Chr(9) & "08:30 PM" & Chr(9) & "08:40 PM" & Chr(9) & "08:50 PM"
    End Select
End Function


Private Function Procedure_Grid_Action()
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim Idx%
Dim Arr
Dim strDuration As String

    If cboDuration.listcount <= 0 Then Load_Duration
    lblCode.Caption = vbNullString
    With grdProcedure
        If IsProperCase(udtAppointment.SchedType) = "Patient" Or udtAppointment.SchedType = vbNullString Then
            udtEncounter.Appt_Type = vbNullString
            udtEncounter.Intrvn_CD = vbNullString
            For lngRow = 1 To .MaxRows
                .BlockMode = True
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = lngRow
                    .Row2 = lngRow
                    .BackColor = NormalColor
                .BlockMode = False
            Next lngRow
            lngRow = .ActiveRow
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                .BackColor = grdCurItem
            .BlockMode = False
            .GetText 1, lngRow, varCellValue
            udtEncounter.Appt_Type = varCellValue
            LookUpCPTCode udtEncounter.Appt_Type
            udtAppointment.VisitType = Left(udtEncounter.Intrvn_CD, 5)
            lblCode.Caption = udtEncounter.Intrvn_CD
            Calc_Duration udtEncounter.Appt_Type
            clsWriteAppt.Duration = udtEncounter.VisitTime
            If cboDuration <> udtEncounter.VisitTime Then
                clsWriteAppt.Duration = Trim(cboDuration)
            End If
            clsWriteAppt.SchedDate = Format(txtDate.Value, "mm/dd/yyyy")
            clsWriteAppt.SchedTime = Format(ccdTime.Value, "hh:mm AM/PM")
             If Check_OverBook_Change Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            If Trim(udtEncounter.EncounterType) <> vbNullString Then
                For Idx = 0 To cboPatientType.listcount - 1
                    If cboPatientType.List(Idx) = udtEncounter.EncounterType Then
                        mbooByPass = True
                        cboPatientType.Text = cboPatientType.List(Idx)
                        Exit For
                    End If
                Next Idx
            End If
            If udtEncounter.VisitTime <> vbNullString Then
                With grdProcedure
                    .GetText 3, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        udtSchedPatient.Duration = varCellValue
                         Arr = Split(udtSchedPatient.Duration, " ")
                        If UBound(Arr) > 0 Then
                            For Idx = LBound(Arr) To UBound(Arr)
                                If IsNumeric(Arr(Idx)) Then
                                    udtEncounter.VisitTime = Trim(Arr(Idx))
                                    Exit For
                                End If
                            Next Idx
                            udtEncounter.VisitTime = udtEncounter.VisitTime & " " & "minutes"
                        End If
                    End If
                End With
            End If
            udtEncounter.Intrvn_CD = Left(udtEncounter.Intrvn_CD, 5)
            If udtEncounter.Intrvn_CD <> vbNullString Then
                lblCode.Caption = udtEncounter.Intrvn_CD
            Else
                LookUpCPTCode udtEncounter.Appt_Type
                lblCode.Caption = udtEncounter.Intrvn_CD
            End If
            udtAppointment.VisitType = udtEncounter.Appt_Type
        Else
            For lngRow = 1 To .MaxRows
                .BlockMode = True
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = lngRow
                    .Row2 = lngRow
                    If .BackColor = grdCurItem Then
                        .GetText 3, lngRow, varCellValue
                        strDuration = varCellValue
                        clsWriteAppt.Duration = vbNullString
                        clsWriteAppt.Duration = Trim(strDuration)
                        clsWriteAppt.SchedDate = Format(txtDate.Value, "mm/dd/yyyy")
                        clsWriteAppt.SchedTime = Format(ccdTime.Value, "hh:mm AM/PM")
                        If Check_OverBook_Change Then
                            Exit For
                            Screen.MousePointer = vbDefault
                            Exit Function
                        End If
                    End If
                    .BackColor = NormalColor
                .BlockMode = False
            Next lngRow
            lngRow = .ActiveRow
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                .BackColor = grdCurItem
            .BlockMode = False
            .GetText 1, lngRow, varCellValue
            udtEncounter.Appt_Type = varCellValue
            udtAppointment.VisitType = udtEncounter.Appt_Type
            .GetText 3, lngRow, varCellValue
            udtEncounter.VisitTime = varCellValue
            If strDuration <> udtEncounter.VisitTime And strDuration <> vbNullString Then
                clsWriteAppt.Duration = Replace(LCase(strDuration), "minutes", "")
            End If
        End If
    End With
    For Idx = 0 To cboDuration.listcount - 1
        If IsProperCase(Trim(cboDuration.List(Idx))) = IsProperCase(Trim(udtEncounter.VisitTime)) Then
            cboDuration.Text = cboDuration.List(Idx)
            Exit For
        End If
    Next Idx
    lblCPT.Visible = True
    lblCode.Visible = True

    Update_Facility_Class udtAppointment.FacilityID, grdProcedure.ActiveRow
    Screen.MousePointer = vbDefault
    Exit Function
End Function
Private Sub Get_TimeSlot()
Dim sValue As String
Dim Idx As Long
Dim Arr
Dim strError As String

    On Error GoTo Error_Get_TimeSlot
    strError = "registry End"
    strError = "registry Interval"
    clsApptPref.SchedInterval = UCase(clsApptPref.SchedInterval)
    Arr = Split(clsApptPref.SchedInterval, " ")
    If Arr(1) = 15 Then
        clsApptPref.SchedInterval = "EVERY 15 MIN"
    ElseIf Arr(1) = 5 Then
        clsApptPref.SchedInterval = "EVERY 5 MIN"
    ElseIf Arr(1) = 10 Then
        clsApptPref.SchedInterval = "EVERY 10 MIN"
    ElseIf Arr(1) = 20 Then
        clsApptPref.SchedInterval = "EVERY 20 MIN"
    End If
    ReDim mStartHour(11)
    mStartHour(0) = "12:00"
    mStartHour(1) = "01:00"
    mStartHour(2) = "02:00"
    mStartHour(3) = "03:00"
    mStartHour(4) = "04:00"
    mStartHour(5) = "05:00"
    mStartHour(6) = "06:00"
    mStartHour(7) = "07:00"
    mStartHour(8) = "08:00"
    mStartHour(9) = "09:00"
    mStartHour(10) = "10:00"
    mStartHour(11) = "11:00"
    mStartIdx = 8
    
    ReDim mEndHour(11)
    mEndHour(0) = "12:00"
    mEndHour(1) = "01:00"
    mEndHour(2) = "02:00"
    mEndHour(3) = "03:00"
    mEndHour(4) = "04:00"
    mEndHour(5) = "05:00"
    mEndHour(6) = "06:00"
    mEndHour(7) = "07:00"
    mEndHour(8) = "08:00"
    mEndHour(9) = "09:00"
    mEndHour(10) = "10:00"
    mEndHour(11) = "11:00"
    mEndIdx = 8
'    sValue = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScheduleStartHour")
    For Idx = 0 To UBound(mStartHour)
       If mStartHour(Idx) = sValue Then
          mStartIdx = Idx
          Exit For
       End If
    Next Idx
'    strError = "registry Start"
'    sValue = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScheduleEndHour")
    For Idx = 0 To UBound(mEndHour)
       If mEndHour(Idx) = sValue Then
          mEndIdx = Idx
          Exit For
       End If
    Next Idx
    Exit Sub
Error_Get_TimeSlot:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_TimeSlot", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Function Format_CellTypeFixed(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeStaticText
            .FontName = "Times New Roman"
            .FontSize = 8
            .FontBold = True
            .ForeColor = GridBlue
            If vlngBackColor > 0 Then .BackColor = vlngBackColor
            .SetText vlngCol, vlngRow, Trim(vstrText)
            .Text = Trim(vstrText)
            .TypeVAlign = TypeVAlignCenter
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
End Function

Private Function Read_Appt_Look()

End Function
Private Function Get_OpenAppointment()
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim Arr
Dim i%
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim Minutes As Double
Dim strPurpose As String
Dim intCount As Double

    On Error GoTo Error_Get_OpenAppointment
    Select Case UCase(clsApptPref.SchedInterval)
        Case "EVERY 20 MIN"
            Minutes = 20
        Case "EVERY 15 MIN"
            Minutes = 15
        Case "EVERY 10 MIN"
            Minutes = 10
        Case "EVERY 5 MIN"
            Minutes = 5
    End Select
    With grdAvailAppt
        .ClearRange 1, 1, .MaxCols, grdAvailAppt.MaxRows, False
        Format_CellTypeEdit grdAvailAppt, SpreadHeader, 1, "4", "APPOINTMENT AVAILABILITY FOR" & " " & UCase(Format(ccSchedDate.Value, "MMMM DD, YYYY")), "TypeHAlignCenter", "HEADER", HeaderColor
        .SetCellBorder 1, SpreadHeader, 4, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdAvailAppt
        Format_CellTypeEdit grdAvailAppt, SpreadHeader + 2, 0, "1", "TIME", "TypeHAlignCenter", "HEADER", HeaderColor
        .SetCellBorder 0, SpreadHeader + 2, 0, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeEdit grdAvailAppt, SpreadHeader + 2, 1, "1", "STATUS", "TypeHAlignCenter", "HEADER", HeaderColor
        .SetCellBorder 1, SpreadHeader + 2, 1, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeEdit grdAvailAppt, SpreadHeader + 2, 2, "1", "COUNT", "TypeHAlignCenter", "HEADER", HeaderColor
        .SetCellBorder 2, SpreadHeader + 2, 2, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdAvailAppt
    End With
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    strSQL = "Select APPT_TIME, DURATION, PURPOSE_VISIT, APPT_TYPE  FROM APPOINTMENTS where APPT_DATE = '" & Format(ccSchedDate.Value, "mm/dd/yyyy") & "' AND " _
            & "PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
            & "FACILITY_ID = " & udtAppointment.FacilityID & " AND ACTIVE = 1"
    Set rsAppt = cnAppt.Execute(strSQL)
    With rsAppt
        Do Until .EOF
            Clear_Time
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "APPT_TIME"
                                udtTime.StartTime = Fld.Value
                            Case "DURATION"
                                udtTime.Duration = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtTime.Purpose = Fld.Value
                            Case "APPT_TYPE"
                                If udtTime.ApptType = vbNullString Then
                                    udtTime.ApptType = Fld.Value
                                End If
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtTime.Purpose = vbNullString And UCase(udtTime.ApptType) = "VENDOR" Then udtTime.Purpose = "Meeting"
            Arr = Split(udtTime.Duration, " ")
            If UBound(Arr) > 0 Then
               If IsNumeric(Arr(0)) Then udtTime.TimeSpan = Val(Val(Arr(0)) / Val(Minutes))
            Else
               If IsNumeric(Minutes) Then udtTime.TimeSpan = Val(Val(udtTime.Duration) / Val(Minutes))
            End If
            lngCol = 0
            With grdAvailAppt
                For lngRow = 1 To .MaxRows
                    .GetText lngCol, lngRow, varCellValue
                    If varCellValue = udtTime.StartTime Then
                        udtTime.EndTime = DateAdd("n", Val(Minutes), udtTime.StartTime)
                        intCount = Get_Count
                        If Trim(Get_MultPurpose) <> vbNullString Then
                            udtTime.Purpose = Trim(Get_MultPurpose)
                        End If
                        Format_CellTypeEdit_RowSpan grdAvailAppt, lngRow, 1, CStr(udtTime.TimeSpan), udtTime.Purpose, "TypeHAlignCenter", "BOOKED"
                        .SetCellBorder 1, lngRow, 1, lngRow + (udtTime.TimeSpan - 1), CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Format_CellTypeEdit_RowSpan grdAvailAppt, lngRow, 2, CStr(udtTime.TimeSpan), CStr(intCount), "TypeHAlignCenter", "BOOKED"
                        .SetCellBorder 2, lngRow, 2, lngRow + (udtTime.TimeSpan - 1), CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Exit For
                    End If
                Next lngRow
            End With
        Loop
    End With
    Set_BlockTime
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
     Exit Function
Error_Get_OpenAppointment:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_OpenAppointment", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_MultPurpose() As String
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strApptType As String

    On Error GoTo Error_Get_Count
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    Get_MultPurpose = vbNullString
    strSQL = "SELECT PURPOSE_VISIT, APPT_TYPE FROM APPOINTMENTS WHERE " _
        & "APPT_DATE = '" & Format(ccSchedDate.Value, "mm/dd/yyyy") & "' AND " _
        & "PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
        & "FACILITY_ID = " & udtAppointment.FacilityID & " AND ACTIVE = 1 AND " _
        & "CONVERT(DATETIME, APPT_TIME, 108) BETWEEN '" & udtTime.StartTime & "' AND '" & udtTime.EndTime & "'"
    Set rsAppt = cnAppt.Execute(strSQL)
    With rsAppt
        Do Until .EOF
            If Not IsNull((!Appt_Type)) Then
                strApptType = (!Appt_Type)
            End If
            If Not IsNull(!PURPOSE_VISIT) Then
                If Get_MultPurpose = vbNullString Then
                    Get_MultPurpose = (!PURPOSE_VISIT)
                Else
                    Get_MultPurpose = Get_MultPurpose & vbNewLine & (!PURPOSE_VISIT)
                End If
            ElseIf IsProperCase(strApptType) = "Vendor" Then
                If Get_MultPurpose = vbNullString Then
                    Get_MultPurpose = "Meeting"
                Else
                    Get_MultPurpose = Get_MultPurpose & vbNewLine & "Meeting"
                End If
            End If
            .MoveNext
        Loop
    End With
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
     Exit Function
Error_Get_Count:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Count", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Function DayWeek(vstrDayWeek As String) As String
Dim strSQL As String
Dim nDays As Long
Dim Arr
Dim booinsert As Boolean
Dim RepeatIndex As Integer
Dim strRepDate As String
Dim Idx As Integer
Dim lngDateDiff As Long
Dim lngPatientID As Long
Dim FirstDay As VbDayOfWeek
Dim lngWeekDays As Long
Dim strOption As String
Dim strMonth As String
Dim strTestMonth As String
Dim lngEndDate As Long


    On Error GoTo Error_DayWeek
    Screen.MousePointer = vbHourglass
    
    If IsDate(udtAppointment.SchedDate) And IsDate(udtBlock.EndDate) Then
        lngDateDiff = DateDiff("ww", udtAppointment.SchedDate, udtBlock.EndDate)
        If udtBlock.Duration = "Permanent" Then lngDateDiff = 1
        If lngDateDiff >= 0 Then
            strRepDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
            If lngDateDiff = 0 Then lngDateDiff = 1
            For Idx = 1 To lngDateDiff
                FirstDay = Weekday(strRepDate, vbMonday)
                nDays = Val(FirstDay)
                If strRepDate <> vbNullString Then
                    Select Case udtBlock.DayWeek
                        Case "Monday"
                            FirstDay = Weekday(strRepDate, vbMonday)
                            DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                            Exit For
                        Case "Tuesday"
                            FirstDay = Weekday(strRepDate, vbTuesday)
                            DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                            Exit For
                        Case "Wednesday"
                            FirstDay = Weekday(strRepDate, vbWednesday)
                            DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                            Exit For
                        Case "Thursday"
                            FirstDay = Weekday(strRepDate, vbThursday)
                            DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                            Exit For
                        Case "Friday"
                            FirstDay = Weekday(strRepDate, vbFriday)
                            DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                            Exit For
                        Case "Saturday"
                            FirstDay = Weekday(strRepDate, vbSaturday)
                            DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                            Exit For
                        Case "Sunday"
                            FirstDay = Weekday(strRepDate, vbSunday)
                            DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                            Exit For
                    End Select
                End If
            Next
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_DayWeek:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "DayWeek", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Read_LegalData(vstrData As String, Optional vlngRef_ID As Long)
Dim strSQL As String
Dim lngRows As Long
Dim strCompany As String
Dim strFirstName As String
Dim strLastName As String
Dim cnReferral As New ADODB.Connection
Dim rsReferral As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strName As String
Dim lngVendorID As Long
Dim strType As String
Dim strSearch As String


    On Error GoTo Error_Read_LegalData
    Screen.MousePointer = vbHourglass
    lstListBox.Clear
    If cnReferral.State = adStateClosed Then cnReferral.open psConnect(0)
    If vlngRef_ID > 0 Then
        strType = "LEGAL ID"
    Else
        strType = "LAST NAME"
    End If
    Select Case Trim(strType)
        Case "LAST NAME"
            strSearch = cctSearch.Text
             strSQL = "SELECT * FROM LEGAL_CONTACT WHERE LAST_NAME like '" & strSearch & "%'  ORDER BY LAST_NAME, FIRST_NAME"
        Case "LEGAL ID"
            strSQL = "SELECT * FROM LEGAL_CONTACT WHERE CONVERT(VARCHAR, LEGAL_CONTACT.LEGAL_CONTACT_ID, 50) = " & vlngRef_ID
            'strSQL = "Search_Legal_Legal_ID "
    End Select
    Set rsReferral = cnReferral.Execute(strSQL)
        With rsReferral
            Do Until .EOF
                Clear_Legal
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                             Select Case UCase(Fld.Name)
                                Case "LEGAL_CONTACT_ID"
                                    clsLegal.LegalID = Fld.Value
                                Case "FACILITY_NAME"
                                    clsLegal.Practice = Fld.Value
                                Case "FIRST_NAME"
                                     clsLegal.FirstName = Fld.Value
                                Case "MIDDLE_NAME"
                                     clsLegal.MiddleName = Fld.Value
                                Case "LAST_NAME"
                                    clsLegal.LastName = Fld.Value
                                Case "SUFFIX"
                                     clsLegal.Suffix = Fld.Value
                                Case "TITLE"
                                     clsLegal.Title = Fld.Value
                                Case "LEGAL_ROLE"
                                    clsLegal.Role = Fld.Value
                                Case "CONTACT_EXT"
                                     clsLegal.PhoneExt = Fld.Value
                                Case "ADDRESS1"
                                     clsLegal.Address = Fld.Value
                                Case "ADDRESS2"
                                     clsLegal.Address2 = Fld.Value
                                Case "CITY"
                                    clsLegal.City = Fld.Value
                                Case "EMAIL"
                                    clsLegal.Email = Fld.Value
                                Case "STATE"
                                    clsLegal.State = Fld.Value
                                Case "ZIP_CODE"
                                     clsLegal.Zip = Fld.Value
                                Case "TELEPHONE"
                                     clsLegal.Phone = Fld.Value
                                Case "Facsimile"
                                     clsLegal.Facsimile = Fld.Value
                            End Select
                        End If
                    End If
                Next
                .MoveNext
                If Trim(clsLegal.FirstName) <> vbNullString Then
                   clsLegal.FirstName = Fix_Name(clsLegal.FirstName)
                   clsLegal.Name = clsLegal.FirstName
                End If
                If Trim(clsLegal.MiddleName) <> vbNullString Then
                   clsLegal.Name = clsLegal.Name & " " & clsLegal.MiddleName
                End If
                If Trim(clsLegal.LastName) <> vbNullString Then
                   clsLegal.LastName = Fix_Name(clsLegal.LastName)
                   clsLegal.Name = clsLegal.Name & " " & clsLegal.LastName
                End If
                If Trim(clsLegal.Suffix) <> vbNullString Then
                    clsLegal.Suffix = Fix_Suffix(clsLegal.Suffix)
                    clsLegal.Name = clsLegal.Name & ", " & Trim(clsLegal.Suffix)
                End If
                If Trim(clsLegal.Title) <> vbNullString Then
                    clsLegal.Name = clsLegal.Name & ", " & Trim(clsLegal.Title)
                End If
                If Trim(clsLegal.Practice) <> vbNullString Then
                    clsLegal.Name = clsLegal.Name & " - " & "FIRM NAME: " & clsLegal.Practice
                End If
                If Trim(clsLegal.Name) <> vbNullString Then
                    lstListBox.AddItem Trim(clsLegal.Name)
                    lstListBox.ItemData(lstListBox.NewIndex) = clsLegal.LegalID
                End If
            Loop
        End With
    If lstListBox.listcount < 0 Then
        MsgBox "No Legal Contact found!", vbInformation, "PatientTrac Legal Search"
        Clear_Legal
    End If
    Clear_Legal
    mbooByPass = False
    If cnReferral.State = adStateOpen Then
        cnReferral.Close
        Set cnReferral = Nothing
        Set rsReferral = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_LegalData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_LegalData", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function

Private Function Read_Plan()
Dim strSQL As String
Dim strInform As String
Dim cnPlanNote As New ADODB.Connection
Dim rsPlanNote As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strFollowUp As String
Dim strTerm As String
Dim strPurpose As String
Dim strDate As String
Dim Arr
Dim lngRow As Long
Dim lngCol As Long

    On Error GoTo Error_Get_PlanCollection
    Screen.MousePointer = vbHourglass
    If cnPlanNote.State = adStateClosed Then cnPlanNote.open psConnect(0)
    strSQL = "Get_Patient_Plan_All  " & piPatientID
     Set rsPlanNote = cnPlanNote.Execute(strSQL)
     With rsPlanNote
        Do Until .EOF
            strFollowUp = vbNullString
            strTerm = vbNullString
            strPurpose = vbNullString
            strDate = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "ENCNTR_START_DATE"
                                strDate = Fld.Value
                            Case "FOLLOW_UP"
                                strFollowUp = Fld.Value
                            Case "TERM"
                                strTerm = Fld.Value
                            Case "PURPOSE"
                                strPurpose = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If Trim(strTerm) <> vbNullString Then
                strFollowUp = Trim(strFollowUp) & " " & Trim(strTerm)
            End If
            If Trim(strPurpose) <> vbNullString Then
                Arr = Split(strPurpose, "-")
                If UBound(Arr) > 0 Then
                    strPurpose = Trim(Arr(0))
                Else
                    strPurpose = Trim(strPurpose)
                End If
                Select Case strPurpose
                    Case "Office Visit Level 4", "Office Visit Level 5", "Psychotherapy 40-50 min", "Psychotherapy 40-50 min with E&M"
                        strPurpose = "Extended Visit"
                    Case "Medication Management", "Office Visit Level 3", "Office Visit Level 2"
                        strPurpose = "Brief Visit"
                    Case "Psychotherapy 20-30 min with E&M", "Psychotherapy 20-30 min"
                        strPurpose = "20 Minute Visit"
                End Select
                lngRow = grdPlan.MaxRows + 1
                grdPlan.MaxRows = lngRow
                lngCol = 1
                Format_CellTypeStatic grdPlan, lngRow, lngCol, "1", Format(strDate, "mm/dd/yyyy"), "TypeHAlignLeft"
                grdPlan.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeStatic grdPlan, lngRow, lngCol, "1", strFollowUp, "TypeHAlignCenter", "FOLLOW_UP"
                grdPlan.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 3
                Format_CellTypeStatic grdPlan, lngRow, lngCol, "1", strPurpose, "TypeHAlignCenter", "FOLLOW_UP"
                grdPlan.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        Loop
    End With
   If cnPlanNote.State = adStateClosed Then
        cnPlanNote.Close
        Set cnPlanNote = Nothing
        Set rsPlanNote = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_PlanCollection:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_PlanCollection", App.EXEName & "." & "ModReports"
    Exit Function
    Resume Next
End Function

Private Function Set_GridCurSel(vObjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String, Optional vstrHAlign As String)
Dim Idx As Integer
Dim booData As Boolean
Dim strSpan As String
Dim strTag As String
Dim strCellNote As String
Dim lngCol As Long
Dim lngSpan As Long

    With vObjGrid
        .Redraw = False
        .Col = vlngCol
        .Row = vlngRow
        If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        If .CellType = CellTypeComboBox Then
            For Idx = 0 To .TypeComboBoxCount
                .TypeComboBoxIndex = Idx
                If IsProperCase(Trim(vstrData)) = IsProperCase(Trim(.TypeComboBoxString)) Then
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
        If Trim(vstrHAlign) = vbNullString Then
          .TypeHAlign = TypeHAlignCenter
        ElseIf vstrHAlign = "TypeHAlignLeft" Then
          .TypeHAlign = TypeHAlignLeft
        End If
        If vlngBackColor > 0 Then .BackColor = vlngBackColor
        If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
    .Redraw = True
    End With
End Function

Private Function Get_InsAuth()
Dim Idx%
Dim Arr

    mstrAuthString = vbNullString
    If IsNull(txtDate.Value) Then Exit Function
    mstrAuthString = txtComments.Text
    If udtSchedPatient.IntId <> vbNullString Then
        If Val(clsInsBillData.TreatNo) > 0 Then
            If Not TextFound(clsInsBillData.PreAuthNo, Trim(mstrAuthString), Len(clsInsBillData.PreAuthNo), True) Then
                If Len(clsInsBillData.PreAuthNo) > 2 Then
                    mstrAuthString = "authorization number: " & " " & Trim(clsInsBillData.PreAuthNo)
                End If
                If Trim(clsInsBillData.TreatRemain) <> vbNullString Then
                    mstrAuthString = mstrAuthString & "; " & "remaining number: " & " " & Trim(clsInsBillData.TreatRemain)
                End If
                If Trim(clsInsBillData.TreatApproved) <> vbNullString And IsProperCase(clsInsBillData.TreatApproved) <> "All" Then
                    mstrAuthString = mstrAuthString & "; " & "visit approved: " & " " & Trim(clsInsBillData.TreatApproved)
                End If
                If Trim(clsInsBillData.AuthContact) <> vbNullString > 5 Then
                     mstrAuthString = mstrAuthString & "; " & "contact: " & " " & Trim(clsInsBillData.AuthContact)
                End If
            End If
        End If
    End If
    If mstrAuthString <> vbNullString Then
        If txtComments.Text = vbNullString Then
            txtComments.Text = mstrAuthString
        ElseIf Not TextFound(mstrAuthString, Trim(txtComments.Text), Len(mstrAuthString), True) Then
            txtComments.Text = txtComments.Text & vbNewLine & mstrAuthString
        End If
    End If
    Arr = Split(txtComments.Text, "; ")
    If UBound(Arr) > 0 Then
        txtComments.Text = vbNullString
        For Idx = LBound(Arr) To UBound(Arr)
            If Trim(Arr(Idx)) <> vbNullString And Len(Trim(Arr(Idx))) > 2 Then
                If txtComments.Text = vbNullString Then
                    txtComments.Text = Trim(Arr(Idx))
                ElseIf Not TextFound(Trim(Arr(Idx)), Trim(txtComments.Text), Len(Trim(Arr(Idx))), True) Then
                    txtComments.Text = txtComments.Text & "; " & Trim(Arr(Idx))
                End If
            End If
        Next Idx
    End If
    Exit Function
End Function
Private Function Get_ApptRange(vlngPatientID As Long, vlngProviderID As Long)
Dim Arr
Dim vTime
Dim vName
Dim lngFacilityID As Long
Dim strDate As String
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim strAffirmed As String
Dim strRepeat As String
Dim i%
Dim lngPatientID As Long
Dim strTermDate As String
Dim strApptDate As String
Dim lngDateDiff As Long
Dim lngProviderID As Long
Dim strTime As String
Dim strNewApptDate As String
Dim strTestTime As String
Dim strTestDur As String
Dim strTestEnd As String
Dim lngApptID As Long
Dim lngRow As Long
Dim varCellValue

    On Error GoTo Error_Get_ApptRange
    Screen.MousePointer = vbHourglass
    clsReferral.ReferralID = 0
    If cboDuration.listcount <= 0 Then Load_Duration
    strDate = udtAppointment.SchedDate
    If cboFacility.listcount <= 0 Then Load_Facilities
    For i = 0 To cboFacility.listcount - 1
       If cboFacility.ItemData(i) = udtFacility.FacID Then
            cboFacility.Text = cboFacility.List(i)
            lngFacilityID = cboFacility.ItemData(i)
           Exit For
       End If
    Next i
    lngProviderID = vlngProviderID
    lngPatientID = vlngPatientID
    Clear_SchedPatient
    Get_SchedPatient lngPatientID
    If cnSched.State = adStateClosed Then cnSched.open psConnect(0)
    udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    strSQL = "Get_Patient_Appointment_Date " & lngPatientID & ", " & lngFacilityID & ", " & udtAppointment.ProviderID & ", '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "'"
        udtAppointment.VisitType = vbNullString
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            Do Until .EOF
                    strTestTime = vbNullString
                    strTestDur = vbNullString
                    lngApptID = 0
                    If Not IsNull((!Appt_Time)) Then
                        strTestTime = Format((!Appt_Time), "hh:mm AM/PM")
                    End If
                    If Not IsNull((!Duration)) Then
                        strTestDur = (!Duration)
                    End If
                    lngApptID = (!Appt_ID)
                .MoveNext
                If strTestDur <> vbNullString Then
                    strTestEnd = DateAdd("n", Val(Replace(strTestDur, "minutes", "")), strTestTime)
                    If Format(udtAppointment.SchedTime, "hh:mm AM/PM") >= Format(strTestTime, "hh:mm AM/PM") And Format(udtAppointment.SchedTime, "hh:mm AM/PM") <= Format(strTestEnd, "hh:mm AM/PM") Then
                        udtAppointment.SchedTime = Format(strTestTime, "hh:mm AM/PM")
                        txtDate.Value = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
                        ccdTime.Value = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                        mstrOrigApptTime = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                        Exit Do
                    End If
                End If
                 lngApptID = 0
            Loop
        End With
    If lngApptID > 0 Then
        strSQL = "Get_Appointment_Appt_ID " & lngApptID
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            If Not .EOF Then
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                           Select Case UCase(Fld.Name)
                               Case "STATUS"
                                   udtSchedPatient.Status = Trim(Fld.Value)
                               Case "DURATION"
                                   udtSchedPatient.Duration = Trim(Fld.Value)
                               Case "APPT_TIME"
                                   udtAppointment.SchedTime = Format(Fld.Value, "hh:mm AM/PM")
                               Case "PURPOSE_VISIT"
                                  udtAppointment.VisitType = Trim(Fld.Value)
                               Case "NOTES"
                                   udtSchedPatient.ApptComment = Trim(Fld.Value)
                               Case "PATIENT_ID"
                                   udtSchedPatient.IntId = Trim(Fld.Value)
                               Case "APPT_CONFIRMED"
                                   strAffirmed = Trim(Fld.Value)
                               Case "REASON"
                                   udtSchedPatient.Reason = Trim(Fld.Value)
                               Case "REPEAT_APPOINT"
                                   strRepeat = Trim(Fld.Value)
                               Case "TERMINATION_DATE"
                                   strTermDate = Trim(Fld.Value)
                               Case "PATIENT_TYPE"
                                   udtSchedPatient.PatientType = Trim(Fld.Value)
                               Case "APPT_TYPE"
                                   udtAppointment.SchedType = Trim(Fld.Value)
                           End Select
                       End If
                   End If
                Next
                For i = 0 To cboStatus.listcount - 1
                   If cboStatus.List(i) = Trim(udtSchedPatient.Status) Then
                       mbooByPass = True
                       cboStatus.Text = cboStatus.List(i)
                       Exit For
                   End If
                Next i
                For i = 0 To cboDuration.listcount - 1
                   If IsProperCase(cboDuration.List(i)) = IsProperCase(udtSchedPatient.Duration) Then
                       cboDuration.Text = cboDuration.List(i)
                       Exit For
                   End If
                Next i
                mbooByPass = False
                If IsProperCase(udtAppointment.SchedType) <> vbNullString And IsProperCase(udtAppointment.SchedType) <> "Patient" Then
                   If cboPatientType.listcount < 0 Then Load_Visit
                Else
                   If cboPatientType.listcount <= 0 Then Load_Visit
                   For i = 0 To cboPatientType.listcount - 1
                       If cboPatientType.List(i) = IsProperCase(udtSchedPatient.PatientType) Then
                           cboPatientType.Text = cboPatientType.List(i)
                           Load_Visit cboPatientType.Text
                           Exit For
                       End If
                   Next i
                End If
                For i = 0 To cboReason.listcount - 1
                   If cboReason.List(i) = udtSchedPatient.Reason Then
                       cboReason.Text = cboReason.List(i)
                       Exit For
                   End If
                Next i
                If strAffirmed <> vbNullString Then txtApptConfirm.Value = Format(strAffirmed, "mm/dd/yyyy")
                For lngRow = 1 To grdProcedure.MaxRows
                    With grdProcedure
                        .Col = 1
                        .Row = lngRow
                        .GetText 1, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            If UCase(varCellValue) = UCase(udtAppointment.VisitType) Then
                                Format_CellTypeStatic grdProcedure, lngRow, 1, "1", udtAppointment.VisitType, "TypeHAlignCenter", "VISIT TYPE"
                                grdProcedure.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                .BlockMode = True
                                    .Col = 1
                                    .Col2 = .MaxCols
                                    .Row = lngRow
                                    .Row2 = lngRow
                                    .BackColor = grdCurItem
                                .BlockMode = False
                                Exit For
                            End If
                        End If
                    End With
                Next lngRow
                If udtAppointment.SchedTime <> vbNullString And udtAppointment.SchedDate <> vbNullString Then
                    lblAppt2.Caption = Format(udtAppointment.SchedDate, "MMMM DD, YYYY") & " " & Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                End If
                If Trim(udtSchedPatient.ApptComment) <> vbNullString Then txtComments.Text = Trim(udtSchedPatient.ApptComment)
                mUpdate = True
                Select Case IsProperCase(udtAppointment.SchedType)
                   Case "Patient"
                       Load_Patient udtSchedPatient.IntId, "SPECIFIC PATIENT"
                   Case "Physician"
                       Read_ReferralData "REFERRAL ID", Val(udtSchedPatient.IntId)
                   Case "Vendor"
                       Read_VendorData "VENDOR ID", Val(udtSchedPatient.IntId)
                   Case "Legal"
                       Read_LegalData "LEGAL ID", Val(udtSchedPatient.IntId)
                End Select
                If udtSchedPatient.IntId <> vbNullString Then
                    For i = 0 To lstListBox.listcount - 1
                        If lstListBox.ItemData(i) = Val(udtSchedPatient.IntId) Then
                            mbooByPass = True
                            lstListBox.Selected(i) = True
                            Exit For
                        End If
                    Next i
                End If
                fraInfo.Visible = True
                cmdEditPatient.Visible = True
                If IsProperCase(udtAppointment.SchedType) = "Patient" Then Get_Referrals
                cctReferalReas.Text = Trim(clsReferral.Consult_Reason)
            End If
        End With
    End If
   If cnSched.State = adStateOpen Then
        cnSched.Close
        Set cnSched = Nothing
        Set rsSched = Nothing
    End If
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_ApptRange:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_ApptRange", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Get_RepeatData(vlngFacility As Long, vlngProvider As Long)
Dim Arr
Dim vTime
Dim vName
Dim lngFacilityID As Long
Dim strDate As String
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim strAffirmed As String
Dim strRepeat As String
Dim i%
Dim lngPatientID As Long
Dim strTermDate As String
Dim strApptDate As String
Dim lngDateDiff As Long


    On Error GoTo Error_Get_RepeatData
    Screen.MousePointer = vbHourglass
    If cnSched.State = adStateClosed Then cnSched.open psConnect(0)
    strSQL = "Get_Repeat_Appt_PatientID " & vlngFacility & "," & vlngProvider & "," & Val(udtSchedPatient.IntId)
        Set rsSched = cnSched.Execute(strSQL)
         With rsSched
            Do Until .EOF
                lngDateDiff = 0
                If Not IsNull(!START_DATE) Then
                    If (!START_DATE) <> vbNullString And Format((!START_DATE), "mm/dd/yyyy") <> "01/01/1900" Then strApptDate = (!START_DATE)
                    If Trim(strApptDate) <> vbNullString Then
                        strApptDate = Format(strApptDate, "mm/dd/yyyy")
                        lngDateDiff = DateDiff("d", udtAppointment.SchedDate, strApptDate)
                    End If
                Else
                    strApptDate = vbNullString
                End If
                If lngDateDiff > 0 Then
                    If Not IsNull(!REPEAT_DAY_WEEK) Then
                        If (!REPEAT_DAY_WEEK) <> vbNullString Then strRepeat = Trim(!REPEAT_DAY_WEEK)
                    End If
                    If Not IsNull(!TERMINATION_DATE) Then
                        strTermDate = Format((!TERMINATION_DATE), "MMMM DD, YYYY")
                        If (!TERMINATION_DATE) <> vbNullString And Format((!TERMINATION_DATE), "mm/dd/yyyy") <> "01/01/1900" Then strTermDate = Format((!TERMINATION_DATE), "MMMM DD, YYYY")
                    End If
                    Exit Do
                End If
                .MoveNext
            Loop
        End With
    If Trim(strRepeat) <> vbNullString Then
        cctRepeat.Text = "Subsequent Appointments every" & " " & strRepeat & " " & "From" & " " & Format(strApptDate, "MMMM DD, YYYY") & " " & "through" & " " & Format(strTermDate, "MMMM DD, YYYY")
    End If
   If cnSched.State = adStateOpen Then
        cnSched.Close
        Set cnSched = Nothing
        Set rsSched = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_RepeatData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_RepeatData", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Read_ApptGrid()
Dim cnApptGrid As New ADODB.Connection
Dim rsApptGrid As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr
Dim lngApptID As Long
Dim lngProviderID As Long
Dim strList As String
Dim i%
Dim strPatType As String
Dim strVisitType As String
Dim varCellValue
Dim lngRepeatID As Long
Dim strStatus As String


    On Error GoTo Error_Read_ApptGrid
    Screen.MousePointer = vbHourglass
    ccdRevisedDate.Value = Format(txtDate.Value, "mm/dd/yyyy")
    ccdRevisedDate.Enabled = False
    For i = 0 To cboProvider.listcount - 1
        If cboProvider.List(i) = cboProvider.Text Then
            lngProviderID = cboProvider.ItemData(i)
            Exit For
        End If
    Next i
    cboProvider.Locked = True
    grdAppointChange.ClearRange 1, 1, grdAppointChange.MaxCols, grdAppointChange.MaxRows, False
    grdAppointChange.MaxRows = 0
    If cnApptGrid.State = adStateClosed Then cnApptGrid.open psConnect(0)
    strSQL = "SELECT * FROM APPOINTMENTS WHERE PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
        & "PROVIDER_ID = " & lngProviderID & " AND  " _
        & "CONVERT(VARCHAR(10), APPT_DATE,101) = '" & Format(ccdRevisedDate.Value, "mm/dd/yyyy") & "' AND " _
        & "APPT_TYPE = '" & udtAppointment.SchedType & "' ORDER " _
        & "BY CONVERT(DATETIME, APPOINTMENTS.APPT_TIME, 108)  DESC"
    Set rsApptGrid = cnApptGrid.Execute(strSQL)
    With rsApptGrid
        Do Until .EOF
            lngApptID = 0
            strPatType = vbNullString
            lngApptID = (!Appt_ID)
            If lngApptID > 0 Then
                lngRow = grdAppointChange.MaxRows + 1
                grdAppointChange.MaxRows = lngRow
                Arr = Split(LookUpCellTag("APPT_TIME"), "|")
                If UBound(Arr) > 0 Then
                    lngCol = Arr(0)
                    If lngCol > 0 Then
                        Format_CellTypeCombo grdAppointChange, lngRow, lngCol, "1", Load_TimeGrid, "TypeHAlignCenter", "APPT_TIME"
                    End If
                End If
                Arr = Split(LookUpCellTag("PATIENT_TYPE"), "|")
                If UBound(Arr) > 0 Then
                    lngCol = Arr(0)
                    If lngCol > 0 Then
                        strList = "Business" & Chr(9) & "Established Patient" & Chr(9) & "New Patient"
                        Format_CellTypeCombo grdAppointChange, lngRow, lngCol, "1", strList, "TypeHAlignCenter", "PATIENT_TYPE"
                    End If
                End If
                Arr = Split(LookUpCellTag("DURATION"), "|")
                If UBound(Arr) > 0 Then
                    lngCol = Arr(0)
                    If lngCol > 0 Then
                        Format_CellTypeCombo grdAppointChange, lngRow, lngCol, "1", Get_GridDuration, "TypeHAlignCenter", "DURATION"
                    End If
                End If
                Arr = Split(LookUpCellTag("STATUS"), "|")
                If UBound(Arr) > 0 Then
                    lngCol = Arr(0)
                    If lngCol > 0 Then
                        Format_CellTypeCombo grdAppointChange, lngRow, lngCol, "1", Get_Grid_Status, "TypeHAlignCenter", "STATUS"
                    End If
                End If
            End If
            lngRepeatID = 0
            strStatus = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case Fld.Name
                            Case "REPEAT_ID"
                                lngRepeatID = Fld.Value
                            Case "APPT_DATE"
                                Arr = Split(LookUpCellTag(Fld.Name), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                    If lngCol > 0 Then
                                        Format_CellTypeStatic grdAppointChange, lngRow, lngCol, "1", Format(Fld.Value, "mm/dd/yyyy"), "TypeHAlignCenter", UCase(Fld.Name)
                                    End If
                                End If
                            Case "APPT_TIME"
                                Arr = Split(LookUpCellTag(Fld.Name), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                    If lngCol > 0 Then
                                        Set_GridCurSel grdAppointChange, Trim(Fld.Value), lngCol, lngRow, "1", "APPT_TIME"
                                    End If
                                End If
                            Case "PROVIDER_ID"
                                Arr = Split(LookUpCellTag(Fld.Name), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                    If lngCol > 0 Then
                                        Get_Providers Val(Fld.Value)
                                        Format_CellTypeStatic grdAppointChange, lngRow, lngCol, "1", clsProvider.Name, "TypeHAlignLeft", UCase(Fld.Name), 0, CStr(Fld.Value)
                                    End If
                                End If
                            Case "PATIENT_TYPE"
                                Arr = Split(LookUpCellTag(Fld.Name), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                    If lngCol > 0 Then
                                        Set_GridCurSel grdAppointChange, Trim(Fld.Value), lngCol, lngRow, "1", "PATIENT_TYPE"
                                        strPatType = Trim(Fld.Value)
                                    End If
                                End If
                            Case "PURPOSE_VISIT"
                                strVisitType = Trim(Fld.Value)
                            Case "DURATION"
                                Arr = Split(LookUpCellTag(Fld.Name), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                    If lngCol > 0 Then
                                       ' Format_CellTypeCombo grdAppointChange, lngRow, lngCol, "1", Get_GridDuration, "TypeHAlignCenter", "DURATION"
                                        Set_GridCurSel grdAppointChange, Trim(Fld.Value), lngCol, lngRow, "1", "DURATION"
                                    End If
                                End If
                            Case "STATUS"
                                Arr = Split(LookUpCellTag(Fld.Name), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                    If lngCol > 0 Then
                                        strStatus = Trim(Fld.Value)
                                        Set_GridCurSel grdAppointChange, Trim(strStatus), lngCol, lngRow, "1", "STATUS"
                                    End If
                                End If
                            Case "NOTES"
                                udtSchedPatient.ApptComment = vbNullString
                                Arr = Split(LookUpCellTag(Fld.Name), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                    If lngCol > 0 Then
                                        udtSchedPatient.ApptComment = Replace(Fld.Value, "'", "''")
                                        Format_CellTypeEdit grdAppointChange, lngRow, lngCol, "1", Trim(udtSchedPatient.ApptComment), "TypeHAlignLeft", UCase(Fld.Name)
                                    End If
                                End If
                        End Select
                    End If
                End If
            Next
            .MoveNext
            For lngCol = 1 To grdAppointChange.MaxCols
                grdAppointChange.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Next lngCol
            If Trim(strPatType) <> vbNullString Then
            End If
            Arr = Split(LookUpCellTag("PURPOSE_VISIT"), "|")
            If UBound(Arr) > 0 Then
                lngCol = Arr(0)
                If lngCol > 0 Then
                    Format_CellTypeCombo grdAppointChange, lngRow, lngCol, "1", Get_VisitGrid(strPatType), "TypeHAlignCenter", "PURPOSE_VISIT"
                    Set_GridCurSel grdAppointChange, Trim(strVisitType), lngCol, lngRow, "1", "PURPOSE_VISIT"
                End If
            End If
            If lngRepeatID > 0 And Trim(strStatus) = vbNullString Then
                If Check_Repeat_Cancel(lngRepeatID) = "Cancelled" Then
                    Arr = Split(LookUpCellTag("STATUS"), "|")
                    If UBound(Arr) > 0 Then
                       lngCol = Arr(0)
                       If lngCol > 0 Then
                           Set_GridCurSel grdAppointChange, "Cancelled", lngCol, lngRow, "1", "STATUS"
                       End If
                    End If
                End If
            End If
            If lngRow > 0 Then
                grdAppointChange.SetRowItemData lngRow, lngApptID
            End If
        Loop
    End With
    udtSchedPatient.ApptComment = vbNullString
    If cnApptGrid.State = adStateOpen Then
        cnApptGrid.Close
        Set cnApptGrid = Nothing
        Set rsApptGrid = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_ApptGrid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_ApptGrid", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub Clear_Fields()
    Clear_Enctr
    cboStatus.ListIndex = -1
    cboReason.ListIndex = -1
    txtComments.Text = vbNullString
    cboFacility.ListIndex = -1
    txtApptConfirm.Value = vbNullString
    cctSearch.Text = vbNullString
    txtDate.Value = vbNullString
    lstReferral.Clear
    cctReferral.Text = vbNullString
    cboPatientType.ListIndex = -1
    lblID2.Caption = vbNullString
    lblName2.Caption = vbNullString
    cboReferBy.Clear
    cctReasReferral.Text = vbNullString
    cboDuration.ListIndex = -1
    Exit Sub
End Sub






Private Function Load_Patient(vstrData As String, Optional vstrType As String)
Dim strSQL As String
Dim strSearch As String
Dim Cnt As Long
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strName As String
Dim strTitle As String
Dim strSuffix As String
Dim lngPatID As Long
Dim Arr

    On Error GoTo Error_Load_Patient
    Screen.MousePointer = vbHourglass
    lstListBox.Clear
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    If vstrType <> vbNullString Then
        strSQL = "Search_Patient_EMR_ID '" & vstrData & "%'"
    Else
        frmSchedPatient.Caption = "Search Patients"
        If cctSearch.Text = vbNullString Then
            MsgBox "Please enter a search item first.", vbInformation, "PatientTrac Patient Search"
            Exit Function
        End If
        strSearch = cctSearch.Text
        If TextFound("'", strSearch, Len("'"), True) Then
            strSearch = Replace(strSearch, "'", "%")
        End If
        If TextFound("'", strSearch, Len("'"), True) Then
            strSearch = Replace(strSearch, "'", "%")
        End If
        If TextFound(",", strSearch, Len(","), True) Then
            Arr = Split(strSearch, ",")
            If UBound(Arr) > 0 Then
                strFirst = Trim(Arr(1))
                strSearch = Arr(0)
                strFirst = Replace(strFirst, " ", "")
                strSearch = Replace(strSearch, " ", "")
                strSearch = Trim(strSearch)
            End If
        End If
        If Trim(strFirst) <> vbNullString Then
             strSQL = "Search_Patient_FirstName_Last_Name '" & Trim(strSearch) & "%', '" & Trim(strFirst) & "%'"
        Else
            strSQL = "Search_Patient_Row_Count_Last_Name 50, '" & Trim(strSearch) & "%' "
        End If
    End If
    If strSQL <> vbNullString Then
        Set rsPat = cnPat.Execute(strSQL)
        Do Until rsPat.EOF
            lngPatID = 0
            lngPatID = rsPat.Fields("PATIENT_ID")
            rsPat.MoveNext
            Get_PatientName lngPatID
            If Trim(udtPatientName.Suffix) <> vbNullString Then
                udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
            Else
                udtPatientName.Name = udtPatientName.LastName & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
            End If
            If Trim(udtPatientName.Name) <> vbNullString Then
                mbooByPass = True
                lstListBox.AddItem Trim(udtPatientName.Name)
                lstListBox.ItemData(lstListBox.NewIndex) = lngPatID
            End If
        Loop
        Clear_UdtPatient
        mbooByPass = True
    End If
    Clear_UdtPatient
    If cnPat.State = adStateOpen Then
        cnPat.Close
        Set cnPat = Nothing
        Set rsPat = Nothing
    End If
    LockWindowUpdate 0
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Patient", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Public Function Load_SpecificPatient()
Dim i%
Dim lngPatientID As Long

    If lngPatientID > 0 Then
        Load_Patient CStr(lngPatientID), "SPECIFIC PATIENT"
        For i = 0 To lstListBox.listcount - 1
            If lstListBox.ItemData(i) = lngPatientID Then
                lstListBox.Selected(i) = True
                Exit For
            End If
        Next i
    End If
    fraInfo.Visible = True
    Exit Function
End Function
Public Function Read_Data()
Dim strSQL As String
Dim R%
Dim i%
Dim Arr
Dim strName As String

    On Error GoTo Error_ReadData
    Screen.MousePointer = vbHourglass
    GetTerms cboReason, "sd_Appointment_Reason"
    GetTerms cboStatus, "sd_Appointment_Status"
    cboStatus.AddItem "Modify Appointment"
    If udtSchedPatient.IntId = vbNullString Then
        cmdApply.Enabled = False
    End If
    udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    mstrOrigApptTime = vbNullString
    If udtAppointment.SchedTime <> vbNullString Then mstrOrigApptTime = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
    Get_Appt_Providers
    If udtAppointment.ProviderID <> 0 Then
        For i = 0 To cboProvider.listcount - 1
            If cboProvider.ItemData(i) = udtAppointment.ProviderID Then
                cboProvider.Text = cboProvider.List(i)
                Exit For
            End If
        Next i
    End If
    mbooByPass = False
    mstrValidAppoint = "Authorized"
    fraInfo.Visible = False
    frmSchedPatient.Caption = "Search Patients"
    If lstListBox.listcount >= 0 Then
        lstListBox.Clear
    End If
    Load_Facilities
    For i = 0 To cboFacility.listcount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            udtFacility.FacID = cboFacility.ItemData(i)
            Exit For
        End If
    Next i
    If cboFacility.Text = vbNullString Then
     For i = 0 To cboFacility.listcount - 1
        If cboFacility.ItemData(i) = udtFacility.FacID Then
            cboFacility.Text = cboFacility.List(i)
            Exit For
        End If
    Next i
    End If
    If cboSerPlace.listcount <= 0 Then Load_Service Val(udtFacility.FacID)
    lstListBox.Visible = True
    cctSearch.Visible = True
    lblPatInfo.Visible = True
    cmdNewPatient.Visible = True
   Exit Function
Error_ReadData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "ReadData", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Public Function GetDay(ipDate As String)
    Dim lDay As Long

    lDay = Weekday(ipDate)

    Select Case lDay
       Case 1
          GetDay = "SUNDAY"
       Case 2
          GetDay = "MONDAY"
       Case 3
          GetDay = "TUESDAY"
       Case 4
          GetDay = "WEDNESDAY"
       Case 5
          GetDay = "THURSDAY"
       Case 6
          GetDay = "FRIDAY"
       Case 7
          GetDay = "SATURDAY"
    End Select
End Function
Private Function Load_ApptTime()
Dim Idx As Integer

    ccdTime.Value = Format(Now, "hh:mm AM/PM")
    Exit Function
End Function

Function Read_Referral()
Dim i%
Dim strSQL As String
Dim booinsert As Boolean
Dim lngAcctID As Long
Dim strReferral As String
Dim cnRef As New ADODB.Connection
Dim rsRef As New ADODB.Recordset

    On Error GoTo Error_Read_Referral
    Clear_ReferralStrings
    If Val(udtSchedPatient.IntId) = 0 Then Exit Function
    If cnRef.State = adStateClosed Then cnRef.open psConnect(0)
    If clsReferral.ReferralID = 0 Then
            strSQL = "Get_Appt_Referral " & Val(udtSchedPatient.IntId)
            Set rsRef = cnRef.Execute(strSQL)
            With rsRef
                Do Until .EOF
                    If Not IsNull(!Referral_ID) Then
                        clsReferral.ReferralID = (!Referral_ID)
                        If Not IsNull(!REFERRAL_REASON) Then cctReferalReas.Text = (!REFERRAL_REASON)
                        Exit Do
                    End If
                .MoveNext
                Loop
            End With
    End If
    If clsReferral.ReferralID > 0 Then
        Get_Referral clsReferral.ReferralID
        For i = 0 To lstReferral.listcount - 1
            If lstReferral.Selected(i) = clsReferral.ReferralID Then
                cctReferral.Text = clsReferral.Physician
            End If
        Next i
    End If
    If clsReferral.Consult_Reason <> vbNullString Then cctReferalReas.Text = clsReferral.Consult_Reason
    cctReasReferral.Text = clsReferral.Consult_Reason
    If cnRef.State = adStateOpen Then
        cnRef.Close
        Set cnRef = Nothing
        Set rsRef = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Referral:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Referral", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Function Rescheduled_Patient()
Dim vTime
Dim lngPatientID As Long
Dim lngProviderID As Long
Dim lngFacilityID As Long
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim booinsert As Boolean
Dim strPurpose As String
Dim strTime As String
Dim strDate As String
Dim strNotes As String
Dim strConfirm As String
Dim strStatus As String
Dim strReason As String
Dim strRepeatAppt As String
Dim nDays As Long
Dim i%
Dim lngApptID As Long
Dim sToday As String
Dim Idx%
Dim Arr
Dim lngRow As Long
Dim varCellValue

    On Error GoTo Error_Rescheduled_Patient
    Screen.MousePointer = vbHourglass
    mbooByPass = True
    
    If Trim(cboStatus.Text) = vbNullString Then
         MsgBox "Please Select Reschedule in status and provide a reason for reschedule.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
'    If Trim(mstrOrigApptTime) = vbNullString Then
'        Screen.MousePointer = vbDefault
'         MsgBox "No Information Available on Current Appointment, Please Reselect from Appointment Screen!", vbInformation, "PatientTrac Scheduling"
'        frmPatientAppointment.mBeenHere = True
'        Screen.MousePointer = vbDefault
'        mstrValidAppoint = vbNullString: mbooAlreadyBook = False: mbooCancel = False
'        fraInfo.Visible = False
'        Clear_PatientStrings
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
    strDate = Format(txtDate.Value, "mm/dd/yyyy")
    If udtAppointment.SchedTime = vbNullString Then
         udtAppointment.SchedTime = Format(ccdTime.Value, "hh:mm AM/PM")
    End If
    If udtSchedPatient.Duration = vbNullString Then udtSchedPatient.Duration = Trim(cboDuration.Text)
    strNotes = Trim(txtComments.Text)
    strNotes = Replace(txtComments, "'", "''")
    strConfirm = Format(txtApptConfirm.Value, "mm/dd/yyyy")
'    For lngRow = 1 To grdProcedure.MaxRows
'        With grdProcedure
'            .Col = 1
'            .Row = lngRow
'            If .BackColor = grdCurItem Then
'                .GetText 1, lngRow, varCellValue
'                If Trim(varCellValue) <> vbNullString Then
'                    strPurpose = Trim(varCellValue)
'                End If
'                Exit For
'            End If
'        End With
'    Next lngRow
    mbooByPass = True
    strReason = Trim(cboReason.Text)
    strNotes = Trim(txtComments.Text)
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
'    strSQL = "Get__Patient_Appt_ID " & udtAppointment.ProviderID & "," & udtAppointment.FacilityID & ", " _
'        & Val(udtSchedPatient.IntId) & ",'" & Format(mstrOrigApptTime, "hh:mm AM/PM") & "','" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "'"
'    Set rsAppt = cnAppt.Execute(strSQL)
'    If Not rsAppt.EOF Then
'        lngApptID = rsAppt.Fields("APPT_ID")
'    End If
    sToday = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    strNotes = Literals(strNotes, "'")
    strReason = Literals(strReason, "'")
    If udtAppointment.Appt_ID > 0 Then
        strSQL = "Update_Appointments_Rescheduled 'Rescheduled','" & strReason & "','" & strNotes & "','" & sToday & "'," & udtAppointment.Appt_ID
        Set rsAppt = cnAppt.Execute(strSQL)
    End If
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    cmdApply.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Rescheduled_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Rescheduled_Patient", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Cancel_Patient()
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim strTime As String
Dim strDate As String
Dim strNotes As String
Dim strConfirm As String
Dim strStatus As String
Dim strReason As String
Dim i%
Dim strNow As String
Dim lngApptID As Long

    On Error GoTo Error_Cancel_Patient
    Screen.MousePointer = vbHourglass
    mbooByPass = True
    If Trim(cboProvider.Text) <> vbNullString Then
        For i = 0 To cboProvider.listcount - 1
            If cboProvider.List(i) = cboProvider.Text Then
                udtAppointment.ProviderID = cboProvider.ItemData(i)
                Exit For
            End If
        Next i
    End If
    strNow = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    Select Case Trim(cboStatus.Text)
        Case "Cancelled", "Late Cancellation", "No Show"
        If mstrOrigApptTime = vbNullString Then mstrOrigApptTime = udtAppointment.SchedTime
        strReason = Trim(cboReason.Text)
        strNotes = Trim(txtComments.Text)
        strReason = Replace(strReason, "'", "''")
        strNotes = Replace(strNotes, "'", "''")
        If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
        If udtAppointment.Appt_ID = 0 Then
            strSQL = "Get__Patient_Appt_ID_Type  " & udtAppointment.ProviderID & ", " & udtAppointment.FacilityID & ", " _
                & Val(udtSchedPatient.IntId) & ", '" & udtAppointment.SchedTime & "', '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "','" _
                & udtAppointment.SchedType & "' "
            Set rsAppt = cnAppt.Execute(strSQL)
            With rsAppt
                If Not .EOF Then
                    If Not IsNull(!Appt_ID) Then
                        udtAppointment.Appt_ID = (!Appt_ID)
                    End If
                End If
            End With
            If udtAppointment.Appt_ID > 0 Then
                strSQL = "Update_Appointments_Rescheduled  '" & cboStatus & "','" & strReason & "','" & strNotes & "','" & strNow & "', " & udtAppointment.Appt_ID
                Set rsAppt = cnAppt.Execute(strSQL)
                udtAppointment.Status = Trim(udtAppointment.Status)
            End If
        Else
            strSQL = "Update_Appointments_Rescheduled  '" & cboStatus & "','" & strReason & "','" & strNotes & "','" & strNow & "', " & udtAppointment.Appt_ID
                Set rsAppt = cnAppt.Execute(strSQL)
                udtAppointment.Status = Trim(udtAppointment.Status)
        End If
        If cnAppt.State = adStateOpen Then
            cnAppt.Close
            Set cnAppt = Nothing
            Set rsAppt = Nothing
        End If
    End Select
    Screen.MousePointer = vbDefault
    Exit Function
Error_Cancel_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Cancel_Patient", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function
Public Function Set_PatientFrame(vlngPatientID As Long)
Dim i%
Dim strSQL
Dim lngPatientID As Long
Dim intActID As Integer
Dim strBestNumber  As String
Dim sFirstName As String
Dim sLastName As String
Dim strSuffix As String
Dim strTitle As String
Dim mlngLastPat As Long
Dim strCellPhone As String
Dim strPhone As String
Dim strPatient As String
Dim strReason As String
Dim strData As String
Dim strAddress1 As String
Dim strAddress2 As String
Dim strCity As String
Dim strState As String
Dim strZip As String
Dim fso As New FileSystemObject
Dim strCurrentTime As String
Dim lngRow As Long

    On Error GoTo Error_Set_PatientFrame
    Screen.MousePointer = vbHourglass
    If cboDuration.listcount <= 0 Then Load_Duration
    lngPatientID = vlngPatientID
    lblID2.Caption = vbNullString
    lblName2.Caption = vbNullString
    cctSearch.Text = vbNullString
    txtComments.Text = vbNullString
    cmdApply.Enabled = True
    If udtAppointment.SchedType = vbNullString Then
        udtAppointment.SchedType = cboApptType.Text
    End If
    If cboApptType.Text = vbNullString Then
        For i = 0 To cboApptType.listcount - 1
            If LCase(cboApptType.List(i)) = LCase(udtAppointment.SchedType) Then
                mbooByPass = True
                cboApptType.Text = cboApptType.List(i)
                Exit For
            End If
        Next i
    End If
    strCurrentTime = udtAppointment.SchedTime
    If Trim(udtSchedPatient.IntId) <> vbNullString Then
        Select Case IsProperCase(udtAppointment.SchedType)
            Case "Patient"
                Get_SchedPatient Val(udtSchedPatient.IntId)
            Case "Vendor", "Business"
                Get_SchedVendor Val(udtSchedPatient.IntId), udtAppointment.SchedTime
                Load_Visit "Vendor"
            Case "Legal", "Attorney"
                Get_Schedule_Attorney Val(udtSchedPatient.IntId)
            Case "Physician"
                Get_SchedPhysician Val(udtSchedPatient.IntId)
        End Select
    End If
    udtSchedPatient.ApptComment = Replace(udtSchedPatient.ApptComment, "''", "'")
    txtComments.Text = Trim(udtSchedPatient.ApptComment)
    udtAppointment.SchedTime = strCurrentTime
    cmdApply.Enabled = True
    mbooGotYou = True
    cmdEditPatient.Visible = True
    If udtSchedPatient.IntId = vbNullString Then
       udtSchedPatient.IntId = lngPatientID
    End If
    fraInfo.Visible = True
    lblID2.Caption = udtSchedPatient.IntId
    lblName2.Caption = Trim(udtSchedPatient.FirstName) & " " & Trim(udtSchedPatient.LastName)
    lblPhone2.Caption = Trim(udtSchedPatient.Phone)
    lblCell2.Caption = Trim(udtSchedPatient.CellPhone)
    If Len(udtSchedPatient.WorkPhone) > 3 Then lblWorkPhone.Caption = Trim(udtSchedPatient.WorkPhone)
    udtSchedPatient.Address = Replace(Trim(udtSchedPatient.Address), "Address:", "")
    lblAddress.Caption = "Address:" & " " & Trim(udtSchedPatient.Address)
    strBestNumber = udtSchedPatient.PrefContact
    If strBestNumber <> vbNullString Then
        If TextFound("Cell", strBestNumber, Len("Cell"), True) Then
            lblCell2.Caption = lblCell2.Caption & " - " & "Preferred"
            lblCell2.ToolTipText = "Cell Phone is preferred contact"
        ElseIf TextFound("Work", strBestNumber, Len("Work"), True) Then
            lblWorkPhone.Caption = lblWorkPhone.Caption & " - " & "Preferred"
            lblWorkPhone.ToolTipText = "Work Phone is preferred contact"
        ElseIf TextFound("Home", strBestNumber, Len("Home"), True) Then
            lblPhone2.Caption = lblPhone2.Caption & " - " & "Preferred"
            lblPhone2.ToolTipText = "Home Phone is preferred contact"
        End If
    End If
    mbooByPass = True
    If cboFacility.listcount <= 0 Then Load_Facilities
    For i = 0 To cboFacility.listcount - 1
       If cboFacility.ItemData(i) = udtFacility.FacID Then
            cboFacility.Text = cboFacility.List(i)
           Exit For
       End If
    Next i
    cboFacility.Locked = True
    cboFacility.Enabled = False
    If cboProvider.listcount <= 0 Then Get_Appt_Providers
    For i = 0 To cboProvider.listcount - 1
         If cboProvider.ItemData(i) = udtAppointment.ProviderID Then
             cboProvider.Text = cboProvider.List(i)
             Exit For
         End If
    Next i
    cboProvider.Enabled = False
    cboProvider.Locked = True
    mbooByPass = False
    If mbooGotYou And IsProperCase(udtAppointment.SchedType) = "Patient" Then
       psPhoto = psImageFilePath & lngPatientID & "\Photo.jpg"
       If fso.FileExists(psPhoto) Then
          imgPatient.Picture = LoadPicture(psPhoto)
          lblPhoto.Visible = False
       Else
          lblPhoto.Visible = True
       End If
       lstListBox.Visible = False: fraInfo.Visible = True
       cctSearch.Visible = False
       lblPatInfo.Visible = False
       cmdNewPatient.Visible = False
       Get_Info udtAppointment.ProviderID, lngPatientID
       cboApptType.Enabled = False
       lblApptType.Enabled = False
    ElseIf IsProperCase(udtAppointment.SchedType) <> "Patient" And lngPatientID > 0 Then
       Get_Info udtAppointment.ProviderID, lngPatientID
       cboPatientType.Enabled = False
       cboPatientType.AddItem "Business"
       For i = 0 To cboPatientType.listcount - 1
        If IsProperCase(cboPatientType.List(i)) = "Business" Then
            mbooByPass = True
             cboPatientType.Text = cboPatientType.List(i)
             mbooByPass = False
        End If
       Next i
       lblPatientType.Enabled = False
       lstListBox.Visible = False: fraInfo.Visible = True
       cctSearch.Visible = False
       lblPatInfo.Visible = False
       cmdNewPatient.Visible = False
       cboApptType.Enabled = False
       lblApptType.Enabled = False
    End If
    mbooGotYou = False
    mbooByPass = True
    mbooByPass = False
    If IsNull(txtDate.Value) Then
      txtDate.Value = udtAppointment.SchedDate
    End If
    cmdApply.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_PatientFrame:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Set_PatientFrame", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Update_Facility_Class(vlngFacility_ID As Long, Optional vlngRow As Long)
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim strSQL As String
Dim strClass As String
Dim lngRow As Long

    On Error GoTo Error_Update_Facility_Class
    Screen.MousePointer = vbHourglass
    Correlate_ApptVisit udtAppointment.VisitType
    lngRow = vlngRow
    If cnSched.State = adStateClosed Then cnSched.open psConnect(0)
    strSQL = "Get_Facility_Class_Facility_ID " & vlngFacility_ID
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            If Not .EOF Then
                strClass = (!FACILITY_CLASS)
            End If
        End With
    If clsEDI_Client.Site_ID = "9702" Then
        If udtSchedPatient.Current_Location <> vbNullString Then
            If udtSchedPatient.Current_Location <> vbNullString Then
                With grdProcedure
                    .Col = 5
                    .Row = lngRow
                    Set_GridCurSel grdProcedure, Trim(udtSchedPatient.Current_Location), 5, lngRow, "1", "FACILTY_TYPE"
                    .SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End With
            End If
        Else
            strClass = vbNullString
            Select Case clsCptTemp.CPT_Code
                Case "H0001", "H0004", "H0005", "H0015", "H0033"
                    strClass = "SUBSTANCE"
            End Select
            If strClass = vbNullString Then
                If UCase(clsProvider.Group_MemberShip) = "SUBSTANCE" Then
                    strClass = "SUBSTANCE"
                ElseIf UCase(clsProvider.Group_MemberShip) = "MENTAL HEALTH" Then
                    strClass = "MENTAL HEALTH"
                End If
            End If
            If strClass <> vbNullString Then
                With grdProcedure
                    .Col = 5
                    .Row = lngRow
                    Set_GridCurSel grdProcedure, Trim(strClass), 5, lngRow, "1", "FACILTY_TYPE"
                    .SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End With
            End If
        End If
    Else
        If strClass = vbNullString Then
            strClass = "MENTAL HEALTH"
        End If
        With grdProcedure
            For lngRow = 1 To .MaxRows
                .Col = 5
                .Row = lngRow
                Set_GridCurSel grdProcedure, Trim(strClass), 5, lngRow, "1", "FACILTY_TYPE"
                .SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Next lngRow
        End With
    End If
   If cnSched.State = adStateOpen Then
        cnSched.Close
        Set cnSched = Nothing
        Set rsSched = Nothing
    End If
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_Facility_Class:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_Facility_Class", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Validate_ProvAppt(vstrAppt_Time As String, vstrDuration As String) As Boolean
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim varCellValue
Dim strStart As String
Dim strEnd As String
Dim strData As String
Dim strTime As String
Dim strRange As String
Dim strEndTest As String
Dim nDays As Double
Dim booExit As Boolean
Dim lngDateDiff As Long
Dim lngDur As Long
Dim strWeekDay As String
Dim strFoundStart As String
Dim strFoundEnd As String
Dim lngApptCnt As Long
Dim strStartDate As String
Dim strEndDate As String

    On Error GoTo Error_Validate_ProvAppt
    Screen.MousePointer = vbHourglass
    Validate_ProvAppt = False
    If mbLookupPatient Then
        Validate_ProvAppt = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If udtSchedPatient.Status <> vbNullString Then
        Select Case udtSchedPatient.Status
            Case "Rescheduled", "Cancelled", "No Show", "Confirmed", "Present"
                Validate_ProvAppt = False
                Screen.MousePointer = vbDefault
                Exit Function
        End Select
    End If
    If Not IsNull(ccdRevisedDate.Value) Then
        strData = Format(ccdRevisedDate.Value, "mm/dd/yyyy")
    Else
        strData = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    End If
    If Not IsNull(ccdTime.Value) Then strTime = ccdTime.Value
    strTime = Format(strTime, "hh:mm AM/PM")
    If udtAppointment.SchedDate <> vbNullString Then
        strRange = DateAdd("d", -7, udtAppointment.SchedDate)
    End If
    If Trim(strTime) = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        'strStart = Format(strStart, "hh:mm AM/PM")
        strStart = Format(vstrAppt_Time)
        strStart = Format(strStart, "hh:mm AM/PM")
        vstrDuration = Replace(vstrDuration, "minutes", "")
        vstrDuration = Trim(vstrDuration)
        lngDur = Val(vstrDuration) - 1
        strEnd = DateAdd("n", lngDur, vstrAppt_Time)
        strEnd = Format(strEnd, "hh:mm AM/PM")
        If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
        strStartDate = DateAdd("d", -1, Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
        strEndDate = DateAdd("d", 1, Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
        strSQL = "Validate_ProvBlock_Phase1 " & udtAppointment.ProviderID & ",'" & Format(strStart, "hh:mm AM/PM") & "','" _
            & Format(strEnd, "hh:mm AM/PM") & "','" & Format(strStartDate, "mm/dd/yyyy") & "','" & Format(strEndDate, "mm/dd/yyyy") & "'"
            Set rsTerms = cnTerms.Execute(strSQL)
            With rsTerms
                Clear_Block
                Do Until .EOF
                    Set Flds = rsTerms.Fields
                        For Each Fld In Flds
                            If Not IsNull(Fld.Value) Then
                                If Trim(Fld.Value) <> vbNullString Then
                                    Select Case UCase(Fld.Name)
                                        Case "START_TIME"
                                            strFoundStart = Trim(Fld.Value)
                                        Case "END_TIME"
                                            strFoundEnd = Trim(Fld.Value)
'                                        Case "START_DATE"
'                                            strStartDate = Trim(Fld.Value)
'                                        Case "END_DATE"
'                                            strEndDate = Trim(Fld.Value)
                                    End Select
                                End If
                            End If
                        Next
                     .MoveNext
                    If DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(strStart, "hh:mm AM/PM")) > 0 And DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(strEnd, "hh:mm AM/PM")) < 0 Then
                        Validate_ProvAppt = True
                        MsgBox "Provider has this time period permanently blocked, time is blocked from" & " " & strFoundStart & " " & "until" & " " & strFoundEnd
                        Exit Do
                    End If
                    Loop
            End With
        '2nd Case Permanenent
        strStartDate = DateAdd("d", -1, Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
        strEndDate = DateAdd("d", 1, Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
        If Validate_ProvAppt = False Then
           strWeekDay = Weekday_Return(Weekday(udtAppointment.SchedDate))
            strSQL = "Validate_ProvBlock_Phase2 " & udtAppointment.ProviderID & ", " & udtFacility.FacID & ",'" & Format(strStart, "hh:mm AM/PM") & "','" _
                & Format(strEnd, "hh:mm AM/PM") & "', '" & Format(strStartDate, "mm/dd/yyyy") & "', '" & Format(strEndDate, "mm/dd/yyyy") & "','" _
                & strWeekDay & "'"
            Set rsTerms = cnTerms.Execute(strSQL)
            With rsTerms
                Clear_Block
                Do Until .EOF
                    Set Flds = rsTerms.Fields
                    For Each Fld In Flds
                        If Not IsNull(Fld.Value) Then
                            If Trim(Fld.Value) <> vbNullString Then
                                Select Case UCase(Fld.Name)
                                    Case "START_TIME"
                                        strFoundStart = Trim(Fld.Value)
                                    Case "END_TIME"
                                        strFoundEnd = Trim(Fld.Value)
                                    Case "START_DATE"
                                        strStartDate = Trim(Fld.Value)
                                    Case "END_DATE"
                                        strEndDate = Trim(Fld.Value)
                                End Select
                            End If
                        End If
                    Next
                    .MoveNext
                    If DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(vstrAppt_Time, "hh:mm AM/PM")) >= 0 And _
                        DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(strEnd, "hh:mm AM/PM")) < 0 Then
                        Validate_ProvAppt = True
                        MsgBox "Provider has this time period permanently blocked, time is blocked from" & " " & strFoundStart & " " & "until" & " " & strFoundEnd
                        Exit Do
                    End If
                Loop
            End With
        End If
        strStartDate = DateAdd("d", -1, Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
        strEndDate = DateAdd("d", 1, Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
        strEndTest = DateAdd("d", 1, Format(strEndDate, "mm/dd/yyyy 00:00:00 AM/PM"))
        strEndTest = Format(strEndTest, "mm/dd/yyyy 00:00:00 AM/PM")
        If Validate_ProvAppt = False Then
            strSQL = "Validate_ProvBlock_Phase3 " & udtAppointment.ProviderID & "," & udtFacility.FacID & ", '" _
                & Format(strStartDate, "mm/dd/yyyy") & "', '" & Format(strEndDate, "mm/dd/yyyy") & "'"
                 Set rsTerms = cnTerms.Execute(strSQL)
                With rsTerms
                    Clear_Block
                    Do Until .EOF
                        Set Flds = rsTerms.Fields
                        For Each Fld In Flds
                            If Not IsNull(Fld.Value) Then
                                If Trim(Fld.Value) <> vbNullString Then
                                    Select Case UCase(Fld.Name)
                                        Case "START_TIME"
                                            strFoundStart = Trim(Fld.Value)
                                        Case "END_TIME"
                                            strFoundEnd = Trim(Fld.Value)
                                    End Select
                                End If
                            End If
                        Next
                        .MoveNext
                        If DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(vstrAppt_Time, "hh:mm AM/PM")) >= 0 And _
                            DateDiff("n", Format(strEnd, "hh:mm AM/PM"), Format(strFoundEnd, "hh:mm AM/PM")) > 0 Then
                            Validate_ProvAppt = True
                            MsgBox "Provider has this time period permanently blocked, time is blocked from" & " " & strFoundStart & " " & "until" & " " & strFoundEnd
                            Exit Do
                        End If
                    Loop
                End With
        End If
        strStartDate = DateAdd("d", -1, Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
        strEndDate = DateAdd("d", 1, Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
        If Validate_ProvAppt = False Then
            strSQL = "Validate_ProvBlock_Phase4 " & udtAppointment.ProviderID & ", " & udtFacility.FacID & ",'" _
                & Format(strStartDate, "mm/dd/yyyy") & "', '" & Format(strEndDate, "mm/dd/yyyy") & "'"
                 Set rsTerms = cnTerms.Execute(strSQL)
                With rsTerms
                    Clear_Block
                    Do Until .EOF
                        Set Flds = rsTerms.Fields
                        For Each Fld In Flds
                            If Not IsNull(Fld.Value) Then
                                If Trim(Fld.Value) <> vbNullString Then
                                    Select Case UCase(Fld.Name)
                                        Case "START_TIME"
                                            strFoundStart = Trim(Fld.Value)
                                        Case "END_TIME"
                                            strFoundEnd = Trim(Fld.Value)
                                    End Select
                                End If
                            End If
                        Next
                        .MoveNext
                        If DateDiff("n", Format(vstrAppt_Time, "hh:mm AM/PM"), Format(strFoundStart, "hh:mm AM/PM")) < 0 Then
                            If DateDiff("n", Format(vstrAppt_Time, "hh:mm AM/PM"), Format(strFoundEnd, "hh:mm AM/PM")) > 0 Then
                                Validate_ProvAppt = True
                                MsgBox "Provider has this time period permanently blocked, time is blocked from" & " " & strFoundStart & " " & "until" & " " & strFoundEnd
                                Exit Do
                            End If
                        End If
                    Loop
                End With
        End If
        If Validate_ProvAppt = False Then
            strSQL = "Validate_ProvBlock_Phase5 " & udtAppointment.ProviderID & ", " & udtFacility.FacID & ",'" _
                & strRange & "'"
                    Set rsTerms = cnTerms.Execute(strSQL)
                    Do Until rsTerms.EOF
                        Clear_Block
                        Set Flds = rsTerms.Fields
                            For Each Fld In Flds
                                If Not IsNull(Fld.Value) Then
                                    If Trim(Fld.Value) <> vbNullString Then
                                        Select Case UCase(Fld.Name)
                                            Case "PROVIDER_ID"
                                                udtAppointment.ProviderID = Trim(Fld.Value)
                                            Case "START_TIME"
                                                strFoundStart = Trim(Fld.Value)
                                            Case "END_TIME"
                                                strFoundEnd = Trim(Fld.Value)
                                            Case "START_DATE"
                                                udtBlock.StartDate = Trim(Fld.Value)
                                            Case "END_DATE"
                                                udtBlock.EndDate = Trim(Fld.Value)
                                            Case "REASON"
                                                udtBlock.Reason = Trim(Fld.Value)
                                            Case "DURATION"
                                                udtBlock.Duration = Trim(Fld.Value)
                                            Case "FACILITY_ID"
                                                udtFacility.FacID = Trim(Fld.Value)
                                            Case "DAYWEEK"
                                                udtBlock.DayWeek = Trim(Fld.Value)
                                        End Select
                                    End If
                                End If
                            Next
                            rsTerms.MoveNext
                            Select Case IsProperCase(Trim(udtBlock.Duration))
                                Case "1 Week"
                                    nDays = 7
                                Case "2 Weeks", "2 Week"
                                    nDays = 14
                                Case "3 Weeks"
                                    nDays = 21
                                Case "1 Month", "1 Months"
                                    nDays = 30
                                Case "2 Months"
                                    nDays = 60
                                Case "3 Months"
                                    nDays = 90
                               Case "6 Months"
                                    nDays = 180
                               Case "1 Year"
                                    nDays = 365
                               Case "1 Day"
                                    nDays = 1
                            End Select
                            If udtBlock.StartDate <> vbNullString And udtBlock.EndDate <> vbNullString Then
                                If nDays > 1 Then
                                    If DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate) > 0 And _
                                        DateDiff("d", udtBlock.EndDate, udtAppointment.SchedDate) < 0 Then
                                        'Validate_ProvAppt = True
                                        If udtBlock.EndDate <> vbNullString Then
                                            lngDateDiff = DateDiff("d", Format(udtBlock.StartDate, "mm/dd/yyyy"), Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
                                            If lngDateDiff >= 0 Then
                                                lngDateDiff = DateDiff("n", strFoundStart, strTime)
                                                If lngDateDiff >= 0 Then
                                                    lngDateDiff = DateDiff("n", strTime, strFoundEnd)
                                                    If lngDateDiff >= 0 Then
                                                        lngDateDiff = DateDiff("d", Format(udtBlock.EndDate, "mm/dd/yyyy"), Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
                                                        If lngDateDiff <= 0 Then
                                                            If udtBlock.DayWeek <> vbNullString Then
                                                                 If UCase(udtBlock.DayWeek) = GetDay(udtAppointment.SchedDate) Then
                                                                    MsgBox "Provider has blocked this time period!", vbCritical, "PatientTrac Scheduling"
                                                                    Validate_ProvAppt = True
                                                                    Exit Do
                                                                End If
                                                            Else
                                                                MsgBox "Provider has blocked this time period!", vbCritical, "PatientTrac Scheduling"
                                                                Validate_ProvAppt = True
                                                                Exit Do
                                                            End If
                                                        Else
                                                            Validate_ProvAppt = False
                                                            Exit Do
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    ElseIf DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate) = 0 And _
                                        DateDiff("d", udtBlock.EndDate, udtAppointment.SchedDate) = 0 Then
                                        If lngDateDiff = DateDiff("n", udtAppointment.SchedTime, strFoundEnd) < 0 Then
                                            Validate_ProvAppt = False
                                            booExit = True
                                            Exit Do
                                        Else
                                            Validate_ProvAppt = True
                                            Exit Do
                                        End If
                                    End If
                                End If
                            End If
                        Loop
                    End If
        End If
    If Not Validate_ProvAppt Then
        strStart = DateAdd("n", -1, vstrAppt_Time)
        strStart = Format(strStart, "hh:mm AM/PM")
        lngDur = Val(vstrDuration) - 1
        strEnd = DateAdd("n", lngDur, vstrAppt_Time)
        strEnd = Format(strEnd, "hh:mm AM/PM")
        If Not Validate_ProvAppt Then
            If cnTerms.State = adStateOpen Then cnTerms.Close
            If cnTerms.State = adStateClosed Then cnTerms.open psConnect(0)
            strSQL = "Validate_ProvBlock_Phase6 " & udtAppointment.ProviderID & ", '" & Format(strData, "mm/dd/yyyy") & "','" _
                & Format(strStart, "hh:mm AM/PM") & "', '" & Format(strEnd, "hh:mm AM/PM") & "'"
                Set rsTerms = cnTerms.Execute(strSQL)
                With rsTerms
                   Do Until .EOF
                       lngApptCnt = lngApptCnt + (!BOOKED_APPT)
                       .MoveNext
                       If lngApptCnt >= 2 Then
                           Validate_ProvAppt = True
                           MsgBox "Provider already has two appointments scheduled during this time period!", vbInformation, "PatientTrac Scheduling"
                           Exit Do
                       End If
                   Loop
                End With
        End If
    End If
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_ProvAppt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_ProvAppt", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function LookUpCellTag(vstrName As String, Optional vlngCol As Long) As String
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String
    
    lngCol = 0
    With grdAppointChange
        LookUpCellTag = vbNullString
            lngRow = SpreadHeader + 2
            For lngCol = 1 To .MaxCols
                .Col = lngCol
                .Row = lngRow
                If .CellTag <> vbNullString Then strTag = .CellTag
                If Trim(strTag) <> vbNullString And TextFound(vstrName, Trim(strTag), Len(vstrName), True) Then
                    LookUpCellTag = lngCol & "|" & lngRow
                    Exit For
                End If
            Next lngCol
    End With
    Exit Function
End Function



Private Function Write_Appointment()
Dim cnApptGrid As New ADODB.Connection
Dim rsApptGrid As New ADODB.Recordset
Dim strSQL As String
Dim strNow As String


    On Error GoTo Error_Write_Appointment
    Screen.MousePointer = vbHourglass
    If cnApptGrid.State = adStateClosed Then cnApptGrid.open psConnect(0)
    clsWriteAppt.Status = Trim(cboStatus.Text)
    clsWriteAppt.Reason = Trim(cboReason.Text)
'    If udtAppointment.Appt_ID = 0 Then
'        strSQL = "SELECT APPT_ID, APPT_TIME FROM APPOINTMENTS WHERE PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
'            & "PROVIDER_ID = " & clsWriteAppt.ProviderID & " AND FACILITY_ID = " & clsWriteAppt.FacilityID & " AND " _
'            & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & Format(clsWriteAppt.SchedDate, "mm/dd/yyyy") & "' AND APPT_TYPE = '" & udtAppointment.SchedType & "'"
'        Set rsApptGrid = cnApptGrid.Execute(strSQL)
'        With rsApptGrid
'            Do Until .EOF
'                If Not IsNull(!Appt_ID) Then
'                    udtAppointment.Appt_ID = (!Appt_ID)
'                End If
'                If Not IsNull(!Appt_Time) Then
'                    clsWriteAppt.SchedTime = (!Appt_Time)
'                End If
'                .MoveNext
'                If clsWriteAppt.SchedTime <> vbNullString And Trim(udtAppointment.SchedTime) <> vbNullString Then
'                    If Format(clsWriteAppt.SchedTime, "hh:mm AM/PM") = Format(udtAppointment.SchedTime, "hh:mm AM/PM") Then
'                        Exit Do
'                    Else
'                        udtAppointment.Appt_ID = 0
'                    End If
'                End If
'            Loop
'        End With
'    End If
    If udtAppointment.Appt_ID = 0 Then
         clsWriteAppt.SchedTime = Format(ccdTime.Value, "hh:mm AM/PM")
        strSQL = "Write_Appointment " & Val(udtSchedPatient.IntId) & ", " & clsWriteAppt.FacilityID & "," _
            & clsWriteAppt.ProviderID & ",'" & clsWriteAppt.PatientType & "','" & clsWriteAppt.PurposeVisit & "','" _
            & clsWriteAppt.SchedDate & "','" _
            & clsWriteAppt.SchedTime & "','" & clsWriteAppt.Comments & "','" & clsWriteAppt.Confirm & "','" _
            & clsWriteAppt.Status & "','" & clsWriteAppt.Reason & "','" & clsWriteAppt.Repeat_Appt & "','" & clsWriteAppt.Duration & "',1," _
            & clsWriteAppt.Referral & ",'" & clsWriteAppt.Referral_Reas & "', '" _
            & udtAppointment.SchedType & "','" & udtSchedPatient.Current_Location & "','" & PsUserID & "' "
        Set rsApptGrid = cnApptGrid.Execute(strSQL)
         strSQL = "SELECT APPT_ID FROM APPOINTMENTS WHERE PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
            & "PROVIDER_ID = " & clsWriteAppt.ProviderID & " AND FACILITY_ID = " & clsWriteAppt.FacilityID & " AND " _
            & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & Format(clsWriteAppt.SchedDate, "mm/dd/yyyy") & "' AND " _
            & "CONVERT(VARCHAR(08), APPT_TIME, 108)  = '" & Format(clsWriteAppt.SchedTime, "hh:mm AM/PM") & "' AND " _
            & "APPT_TYPE = '" & udtAppointment.SchedType & "'"
        Set rsApptGrid = cnApptGrid.Execute(strSQL)
        With rsApptGrid
            If Not .EOF Then
                If Not IsNull(!Appt_ID) Then
                    udtAppointment.Appt_ID = (!Appt_ID)
                End If
            End If
        End With
    ElseIf udtSchedPatient.IntId <> vbNullString Then
        strNow = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
        strSQL = "Update_Appointments_Schedule '" & clsWriteAppt.PurposeVisit & "','" & clsWriteAppt.Duration & "', '" _
            & clsWriteAppt.Comments & "','" & clsWriteAppt.SchedTime & "','" & clsWriteAppt.PatientType & "','" & strNow & "','" _
            & clsWriteAppt.Confirm & "','" & clsWriteAppt.Status & "','" & clsWriteAppt.Reason & "','" & clsWriteAppt.Repeat_Appt & "'," _
            & clsWriteAppt.Referral & ",'" & clsWriteAppt.Referral_Reas & "','" _
            & udtAppointment.SchedType & "','" & PsUserID & "','" & udtSchedPatient.Current_Location & "'," & udtAppointment.Appt_ID
        Set rsApptGrid = cnApptGrid.Execute(strSQL)
    End If
    If udtAppointment.Appt_ID > 0 Then
         strSQL = "UPDATE APPOINTMENTS SET VISIT_CODE = '" & clsWriteAppt.CPT_Code & "'," _
            & "ADD_PROC_CODE = '" & clsWriteAppt.Add_Proc_DESC & "', ADD_PROC_DESC = '" _
            & clsWriteAppt.Add_Proc_DESC & "', ACTIVE = 1  WHERE APPT_ID = " & udtAppointment.Appt_ID
        Set rsApptGrid = cnApptGrid.Execute(strSQL)
    End If
    If Trim(cboStatus.Text) = vbNullString Then udtAppointment.Status = "New Appointment"
    If cnApptGrid.State = adStateOpen Then
        cnApptGrid.Close
        Set cnApptGrid = Nothing
        Set rsApptGrid = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Appointment:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Appointment", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cboApptType_Click()
    
    If Trim(cboApptType.Text) <> vbNullString Then
        udtAppointment.SchedType = cboApptType.Text
    End If
    Select Case IsProperCase(udtAppointment.SchedType)
        Case "Physician"
            lblID1.Caption = "Physician ID"
        Case "Patient"
            lblID1.Caption = "Patient ID"
        Case "Vendor", "Business"
            lblID1.Caption = "Vendor ID"
        Case "Legal", "Attorney"
            lblID1.Caption = "Legal ID"
    End Select
    Format_PurposeGrid
    lstListBox.Clear
    cctSearch.Text = vbNullString
    Exit Sub
End Sub


Private Sub cboDuration_Click()
Dim i%

    clsWriteAppt.Duration = cboDuration.Text
    clsWriteAppt.SchedDate = Format(txtDate.Value, "mm/dd/yyyy")
    clsWriteAppt.SchedTime = Format(ccdTime.Value, "hh:mm AM/PM")
    If Not mbooByPass Then
        If Check_OverBook_Change Then
            For i = 0 To cboDuration.listcount - 1
                If IsProperCase(cboDuration.List(i)) = IsProperCase(udtSchedPatient.Duration) Then
                    cboDuration.Text = cboDuration.List(i)
                    Exit For
                End If
            Next i
            Screen.MousePointer = vbDefault
        End If
    End If
    Exit Sub
End Sub

Private Sub cboFacility_Click()
Dim lngFacilityID As Long
Dim i%
Dim strText As String
Dim Arr

    On Error GoTo Error_cboFacility_Click
    If mbooByPass Then mbooByPass = False: Exit Sub
    For i = 0 To cboFacility.listcount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            lngFacilityID = cboFacility.ItemData(i)
            Exit For
        End If
    Next i
    mstrAuthString = vbNullString
    If cboFacility.listcount <= 0 Then Get_Appt_Providers
    Load_Service lngFacilityID
    If clsEDI_Client.Site_ID = "9702" Then
        Clear_Surg_Facility
        Read_Patient_Status Val(udtSchedPatient.IntId)
        If Val(clsFacStatus.FacilityID) > 0 Then
            cctFacility.BackColor = NormalColor
            Get_Facilities Val(clsFacStatus.FacilityID)
'            cctFacility.Text = clsFacility.Name
            If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
                cctFacility.Text = Trim(clsFacStatus.StatusDesc)
               ' cctFacility.Text = cctFacility.Text & vbNewLine & Trim(clsFacStatus.StatusDesc)
            End If
            cctFacility.Visible = True
            lblFac.Visible = True
        Else
            cctFacility.Text = cboFacility.Text
            cctFacility.BackColor = NormalColor
        End If
        strText = Trim(cctFacility.Text)
        If Trim(strText) <> vbNullString Then
            Arr = Split(strText, vbNewLine)
            If UBound(Arr) > 0 Then
                strText = Arr(0)
                If Trim(strText) <> Trim(cboFacility.Text) Then
                    cctFacility.BackColor = grdCurItem
                Else
                    cctFacility.BackColor = NormalColor
                End If
            Else
                If Trim(strText) <> Trim(cboFacility.Text) Then
                    cctFacility.BackColor = grdCurItem
                Else
                    cctFacility.BackColor = NormalColor
                End If
            End If
        End If
    Else
        Clear_Surg_Facility
        Read_Patient_Status Val(udtSchedPatient.IntId)
        If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
            Get_Facilities Val(clsFacStatus.FacilityID)
           ' cctFacility.Text = clsFacility.Name
            If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
                cctFacility.Text = Trim(clsFacStatus.StatusDesc)
            End If
            cctFacility.Visible = True
            lblFac.Visible = True
     '   Else
      '      cctFacility.Text = cboFacility.Text
        End If
        cctFacility.BackColor = grdCurItem
    End If
    Exit Sub
Error_cboFacility_Click:
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume Next
End Sub

Private Sub cboFacility_KeyUp(KeyCode As Integer, Shift As Integer)
    cboFacility.ToolTipText = cboFacility.Text
End Sub


Private Sub cboPatientType_Click()
    On Error GoTo Error_PatientType
    If mbooByPass Then mbooByPass = False: Exit Sub
    Load_Visit cboPatientType.Text
    Screen.MousePointer = vbDefault
    Exit Sub
Error_PatientType:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "PatientType", App.EXEName & " " & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Function Load_Visit(Optional vstrType As String)
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSerPlace As String
Dim strPatType As String
Dim strTitle As String
Dim i%
Dim Arr
Dim strProvider As String
Dim strCompare As String
Dim lngFacID As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strApptCode As String
Dim varCellValue
Dim booFound As Boolean
Dim strList As String

    On Error GoTo Error_Load_Visit
    For i = 0 To cboFacility.listcount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            lngFacID = cboFacility.ItemData(i)
            Exit For
        End If
    Next i
    Screen.MousePointer = vbHourglass
    grdProcedure.ClearRange 1, 1, grdProcedure.MaxCols, grdProcedure.MaxRows, False
    grdProcedure.MaxRows = 0
    If LCase(vstrType) <> "vendor" Then
        strCompare = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
        strCompare = Replace(strCompare, ".", "")
        strCompare = Trim(strCompare)
        clsPatAcct.CPT_Code = vbNullString
    End If
    'udtSchedPatient
    If udtAppointment.ProviderID > 0 Then Get_Providers udtAppointment.ProviderID
    strProvider = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    If LCase(udtAppointment.SchedType) = "vendor" Then
        strPatType = "business"
    ElseIf LCase(vstrType) <> vbNullString Then
        strPatType = vstrType
    End If
    If Trim(vstrType) <> vbNullString Then
        strSQL = "SELECT DISTINCT VISIT_TYPE, APPT_CODE, PROVIDER_TITLE, CPT_CODE, APPT_TIME " _
            & "FROM APPOINTMENT_TYPE WHERE VISIT_TYPE like '" & strPatType & "%' AND (PROVIDER_TITLE = 'ALL' OR PROVIDER_TITLE = '" & strProvider & "') ORDER BY  APPT_CODE, VISIT_TYPE"
    Else
        strSQL = "SELECT DISTINCT VISIT_TYPE, APPT_CODE, PROVIDER_TITLE, APPT_TIME " _
            & "FROM APPOINTMENT_TYPE WHERE " _
            & "FACILITY_ID = " & udtAppointment.FacilityID & "  AND (PROVIDER_TITLE = 'ALL' OR PROVIDER_TITLE = '" & strProvider & "') ORDER BY  APPT_CODE, VISIT_TYPE"
    End If
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        Do Until .EOF
            strProvider = vbNullString
            If Not IsNull((!PROVIDER_TITLE)) Then
                strProvider = UCase((!PROVIDER_TITLE))
            End If
            clsPatAcct.Duration = vbNullString
            clsPatAcct.CPT_Code = vbNullString
            strApptCode = vbNullString
            strPatType = vbNullString
            Set Flds = rsCodes.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "CPT_CODE"
                                clsPatAcct.CPT_Code = Trim(Fld.Value)
                            Case "APPT_TIME"
                                clsPatAcct.Duration = Trim(Fld.Value)
                            Case "APPT_CODE"
                                strApptCode = Trim(Fld.Value)
                            Case "VISIT_TYPE"
                                If strPatType = vbNullString Then
                                    strPatType = Trim(Fld.Value)
                                ElseIf Not TextFound(Trim(Fld.Value), strPatType, Len(Trim(Fld.Value)), True) Then
                                    strPatType = strPatType & vbNewLine & Trim(Fld.Value)
                                End If
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If IsProperCase(strApptCode) = "No Show" Then
                strApptCode = vbNullString
            End If
            If DateDiff("d", "01/01/2013", Format(txtDate.Value, "mm/dd/yyyy")) > 1 And IsProperCase(strApptCode) = "Medication Management" Then
                strApptCode = vbNullString
            End If
            booFound = False
            If Trim(strApptCode) <> vbNullString Then
                For lngRow = 1 To grdProcedure.MaxRows
                    grdProcedure.GetText 1, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        If Trim(varCellValue) = strApptCode Then
                            booFound = True
                        End If
                    End If
                Next lngRow
                If Not booFound Then
                    If IsProperCase(udtAppointment.SchedType) = "Patient" Or udtAppointment.SchedType = vbNullString Then
                        lngRow = grdProcedure.MaxRows + 1
                        grdProcedure.MaxRows = lngRow
                        lngCol = 1
                        Format_CellTypeEdit grdProcedure, lngRow, lngCol, "1", Trim(strApptCode), "TypeHAlignLeft", "VISIT TYPE"
                        grdProcedure.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 2
                        If strApptCode <> vbNullString And Trim(clsPatAcct.CPT_Code) = vbNullString Then
                            LookUpCPTCode strApptCode
                        End If
                        Format_CellTypeStatic grdProcedure, lngRow, lngCol, "1", Trim(clsPatAcct.CPT_Code), "TypeHAlignCenter", "CPT_CODE"
                        grdProcedure.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 3
                        Format_CellTypeEdit grdProcedure, lngRow, lngCol, "1", "", "TypeHAlignLeft", "ADD_DESC"
                        grdProcedure.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 4
                        Format_CellTypeEdit grdProcedure, lngRow, lngCol, "1", "", "TypeHAlignLeft", "ADD_CPT_CODE"
                        grdProcedure.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 5
                        strList = Get_Facility_Class
                        Format_CellTypeCombo grdProcedure, lngRow, lngCol, "1", strList, "TypeHAlignLeft", "FACILTY_TYPE"
                        grdProcedure.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Else
                        lngRow = grdProcedure.MaxRows + 1
                        grdProcedure.MaxRows = lngRow
                        lngCol = 1
                        Format_CellTypeEdit grdProcedure, lngRow, lngCol, "2", Trim(strApptCode), "TypeHAlignLeft", "VISIT TYPE"
                        grdProcedure.SetCellBorder lngCol, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 3
                        Format_CellTypeStatic grdProcedure, lngRow, lngCol, "2", UCase(clsPatAcct.Duration), "TypeHAlignCenter", "CPT_CODE"
                        grdProcedure.SetCellBorder lngCol, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 5
                        strList = Get_Facility_Class
                        Format_CellTypeCombo grdProcedure, lngRow, lngCol, "1", strList, "TypeHAlignLeft", "FACILTY_TYPE"
                        grdProcedure.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        
                    End If
                End If
            End If
        Loop
    End With
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    If Trim(strPatType) <> vbNullString Then
        Arr = Split(strPatType, vbNewLine)
        If UBound(Arr) > 0 Then
            For i = LBound(Arr) To UBound(Arr)
                If Trim(Arr(i)) <> vbNullString Then
                    cboPatientType.AddItem CStr(Arr(i))
                End If
            Next i
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Visit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Visit", App.EXEName & " " & Me.Name
    Exit Function
    Resume Next
End Function
Private Sub cboProvider_Click()
Dim i%

    If mbooByPass Then mbooByPass = False: Exit Sub
    For i = 0 To cboProvider.listcount - 1
        If cboProvider.List(i) = cboProvider.Text Then
            udtAppointment.ProviderID = cboProvider.ItemData(i)
            Exit For
        End If
    Next i
    
    mstrAuthString = vbNullString
    Get_Providers udtAppointment.ProviderID
    Exit Sub
End Sub



Private Function Read_ReferralData(vstrType As String, Optional vlngRef_ID As Long)
Dim i%
Dim strSQL As String
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim cnPhys As New ADODB.Connection
Dim rsPhys As New ADODB.Recordset
Dim strSuffix As String
Dim lngRefID As Long
Dim strSearch As String
Dim Arr
Dim strType As String

    On Error GoTo Error_Read_ReferralData
    Screen.MousePointer = vbHourglass
    lstListBox.Clear
    If cnPhys.State = adStateClosed Then cnPhys.open psConnect(0)
    If vlngRef_ID > 0 Then
        strType = "REFERRAL ID"
    Else
        strType = "LAST NAME"
    End If
    Select Case Trim(strType)
        Case "LAST NAME"
            strSearch = cctSearch.Text
            If TextFound("'", strSearch, Len("'"), True) Then
                strSearch = Replace(strSearch, "'", "%")
            End If
            If TextFound("'", strSearch, Len("'"), True) Then
                strSearch = Replace(strSearch, "'", "%")
            End If
            If TextFound(",", strSearch, Len(","), True) Then
                Arr = Split(strSearch, ",")
                If UBound(Arr) > 0 Then
                    strFirst = Trim(Arr(1))
                    strSearch = Arr(0)
                    strFirst = Replace(strFirst, " ", "")
                    strSearch = Replace(strSearch, " ", "")
                    strSearch = Trim(strSearch)
                End If
            End If
            If Trim(strFirst) <> vbNullString Then
                 strSQL = "Search_Referral_Last_Name_First_Name  '" & strSearch & "%', '%" & strFirst & "%'"
            Else
                strSQL = "Search_Referral_Last_Name '%" & strSearch & "%'"
            End If
        Case "REFERRAL ID"
            strSQL = "Get_Referral_By_ID " & vlngRef_ID
    End Select
    Set rsPhys = cnPhys.Execute(strSQL)
    With rsPhys
        Do Until .EOF
            lngRefID = 0
            strFirst = vbNullString
            strLast = vbNullString
            strMiddle = vbNullString
            strTitle = vbNullString
            strSuffix = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "REFERRER_ID"
                               lngRefID = Fld.Value
                            Case "LAST_NAME"
                               strLast = Fix_Name(Fld.Value)
                            Case "FIRST_NAME"
                                strFirst = Fix_Name(Fld.Value)
                            Case "MIDDLE_NAME"
                                strMiddle = Fix_Name(Fld.Value)
                            Case "SUFFIX"
                                strSuffix = Fix_Suffix(Fld.Value)
                            Case "TITLE"
                                strTitle = Fix_Title(Fld.Value)
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If Trim(strSuffix) = vbNullString Then
                mbooByPass = True
                lstListBox.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strTitle)
                If IsNumeric(lngRefID) Then lstListBox.ItemData(lstListBox.NewIndex) = lngRefID
            Else
                mbooByPass = True
                lstListBox.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strSuffix) & " " & Trim(strTitle)
                If IsNumeric(lngRefID) Then lstListBox.ItemData(lstListBox.NewIndex) = lngRefID
            End If
        Loop
    End With
    mbooByPass = False
    If cnPhys.State = adStateOpen Then
        cnPhys.Close
        Set cnPhys = Nothing
        Set rsPhys = Nothing
        Set Flds = Nothing
    End If
    mbooHasChanged = True
    Screen.MousePointer = vbDefault
   Exit Function
Error_Read_ReferralData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_ReferralData", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function


Private Sub cboReason_Click()
    If mbooByPass Then mbooByPass = False: Exit Sub
'    If cboStatus.Text = "Rescheduled" Then
'        Rescheduled_Patient
'        Exit Sub
'    End If
End Sub

Private Sub cboReferBy_Click()
  If mbooByPass Then mbooByPass = False: Exit Sub
    If cboReferBy.Text = "Doctor" Then
        Call MsgBox("Please select the referring physician from the choices provided. If you do not find the referral please add the referring physician to your database.  ", vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1, "History of Present Illness")
        fraRefer.Visible = True
        fraRefer.ZOrder 0
        cctReasReferral.Text = vbNullString
    ElseIf cboReferBy.Text <> vbNullString Then
        fraRefer.Visible = False
        fraRefer.ZOrder 1
    Else
        fraRefer.Visible = False
        fraRefer.ZOrder 1
    End If
    mbooHasChanged = True
    Exit Sub
End Sub



Private Function Load_Specific_Referral(vlngRefID As Long)
Dim strSQL As String
Dim cnReferral As New ADODB.Connection
Dim rsReferral As New ADODB.Recordset
Dim strName As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim strSuffix As String
Dim lngRefID As Long
Dim strRef As String

    On Error GoTo Error_Load_Specific_Referral
    Screen.MousePointer = vbHourglass
    lngRefID = vlngRefID
    If cnReferral.State = adStateClosed Then cnReferral.open psConnect(0)
    lstReferral.Clear
    strSQL = "Get_Referral_By_ID " & lngRefID
    Set rsReferral = cnReferral.Execute(strSQL)
    With rsReferral
        If Not .EOF Then
            lngRefID = 0
            strFirst = vbNullString
            strLast = vbNullString
            strMiddle = vbNullString
            strTitle = vbNullString
            strSuffix = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "REFERRER_ID"
                               lngRefID = Fld.Value
                            Case "LAST_NAME"
                               strLast = Fld.Value
                            Case "FIRST_NAME"
                                strFirst = Fld.Value
                            Case "MIDDLE_NAME"
                                strMiddle = Fld.Value
                            Case "SUFFIX"
                                strSuffix = Fld.Value
                            Case "TITLE"
                                strTitle = Fld.Value
                        End Select
                    End If
                End If
            Next
            If lngRefID > 0 Then
                strLast = Fix_Name(strLast)
                strFirst = Fix_Name(strFirst)
                strMiddle = Fix_Name(strMiddle)
                If Trim(strTitle) <> vbNullString Then
                    strTitle = Fix_Title(strTitle)
                End If
                If Trim(strSuffix) = vbNullString Then
                    lstReferral.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strTitle)
                Else
                    strSuffix = Fix_Suffix(strSuffix)
                    lstReferral.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strSuffix) & " " & Trim(strTitle)
                End If
                lstReferral.ItemData(lstReferral.NewIndex) = lngRefID
                mbooByPass = True
                lstReferral.Selected(lstReferral.NewIndex) = True
            End If
        End If
    End With
    If cnReferral.State = adStateOpen Then
        cnReferral.Close
        Set cnReferral = Nothing
        Set rsReferral = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
   Exit Function
Error_Load_Specific_Referral:
    Screen.MousePointer = vbDefault
     Write_Error Err.Description, Err.Number, strSQL, "Load_Specific_Referral", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cboSerPlace_Click()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim strSerPlace As String
Dim strPatType As String
Dim strTitle As String
Dim i%

    cboPatientType.Clear
    lblCode.Caption = vbNullString
    If cboPatientType.listcount <= 0 Then
        If TextFound("Office", cboSerPlace.Text, Len("Office"), True) Then
            cboPatientType.AddItem "New Patient"
            cboPatientType.AddItem "Established Patient"
            If mstrPractice <> "Plastic Surgery" Then cboPatientType.AddItem "Consultation"
            cboPatientType.AddItem "Second Opinion"
         ElseIf TextFound("Intensive", cboSerPlace.Text, Len("Intensive"), True) Then
            cboPatientType.AddItem "New Patient"
            cboPatientType.AddItem "Established Patient"
         ElseIf TextFound("Care", cboSerPlace.Text, Len("Care"), True) Then
            cboPatientType.AddItem "New Patient"
            cboPatientType.AddItem "Established Patient"
         ElseIf TextFound("InPatient Hospital", cboSerPlace.Text, Len("InPatient Hospital"), True) Then
            cboPatientType.AddItem "Initial InPatient"
            cboPatientType.AddItem "Subsequent Care"
        ElseIf TextFound("Nursing Facility", cboSerPlace.Text, Len("Nursing Facility"), True) Then
            cboPatientType.AddItem "Initial InPatient"
            cboPatientType.AddItem "Subsequent Care"
        ElseIf TextFound("Assisted Living", cboSerPlace.Text, Len("Assisted Living"), True) Then
            cboPatientType.AddItem "Initial InPatient"
            cboPatientType.AddItem "Subsequent Care"
        ElseIf cboSerPlace.Text = "Inpatient Psychiatric Facility" Then
            cboPatientType.AddItem "Initial InPatient"
            cboPatientType.AddItem "Subsequent Care"
        ElseIf TextFound("Nursing Home", cboSerPlace.Text, Len("Nursing Home"), True) Then
            cboPatientType.AddItem "Initial InPatient"
            cboPatientType.AddItem "Subsequent Care"
            cboPatientType.AddItem "Second Opinion"
         ElseIf TextFound("IOP", cboSerPlace.Text, Len("IOP"), True) Then
            cboPatientType.AddItem "Initial InPatient"
            cboPatientType.AddItem "Subsequent Care"
         ElseIf TextFound("PHP", cboSerPlace.Text, Len("PHP"), True) Then
            cboPatientType.AddItem "Initial InPatient"
            cboPatientType.AddItem "Subsequent Care"
         ElseIf TextFound("Ambulatory Surgical Center", cboSerPlace.Text, Len("Ambulatory Surgical Center"), True) Then
            cboPatientType.AddItem "Procedure"
            cboPatientType.AddItem "Admission"
            cboPatientType.AddItem "Recovery Room"
            cboPatientType.AddItem "Ambulatory Procedure Visit"
         End If
     End If
    If cboSerPlace.Text <> vbNullString Then
        For i = 0 To cboSerPlace.listcount - 1
            If cboSerPlace.List(i) = cboSerPlace.Text Then
                udtEncounter.ServicePlaceCode = cboSerPlace.ItemData(i)
            End If
        Next i
    End If
    Exit Sub
End Sub

Private Sub cboSerPlace_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim strSerPlace As String
Dim strPatType As String
Dim strTitle As String
Dim i%

    cboPatientType.Clear
    lblCode.Caption = vbNullString
    If cboStatus.listcount <= 0 Then
        If TextFound("Office", cboSerPlace.Text, Len("Office"), True) Then
            cboStatus.AddItem "Established Patient"
            cboStatus.AddItem "New Patient"
            If mstrPractice <> "Plastic Surgery" Then cboStatus.AddItem "Consultation"
            cboStatus.AddItem "Second Opinion"
         ElseIf TextFound("InPatient Hospital", cboSerPlace.Text, Len("InPatient Hospital"), True) Then
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
        ElseIf TextFound("Nursing Facility", cboSerPlace.Text, Len("Nursing Facility"), True) Then
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
        ElseIf TextFound("Assisted Living", cboSerPlace.Text, Len("Assisted Living"), True) Then
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
        ElseIf cboSerPlace.Text = "Inpatient Psychiatric Facility" Then
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
        ElseIf TextFound("Nursing Home", cboSerPlace.Text, Len("Nursing Home"), True) Then
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
            cboStatus.AddItem "Second Opinion"
         ElseIf TextFound("IOP", cboSerPlace.Text, Len("IOP"), True) Then
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
         ElseIf TextFound("PHP", cboSerPlace.Text, Len("PHP"), True) Then
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
         ElseIf TextFound("Ambulatory Surgical Center", cboSerPlace.Text, Len("Ambulatory Surgical Center"), True) Then
            cboStatus.AddItem "Procedure"
            cboStatus.AddItem "Admission"
            cboStatus.AddItem "Recovery Room"
            cboStatus.AddItem "Ambulatory Procedure Visit"
         End If
     End If
    If cboSerPlace.Text <> vbNullString Then
        For i = 0 To cboSerPlace.listcount - 1
            If cboSerPlace.List(i) = cboSerPlace.Text Then
                udtEncounter.ServicePlaceCode = cboSerPlace.ItemData(i)
            End If
        Next i
    End If
    Exit Sub
End Sub

Private Sub cboStatus_Click()
Dim i%

    If mbooByPass Then mbooByPass = False: Exit Sub
    If cboStatus.Text <> vbNullString Then
        Select Case IsProperCase(cboStatus)
            Case "Cancelled", "Late Cancellation"
                If cboReason.Text = vbNullString Then
                   MsgBox "Please enter a reason for the" & " " & cboStatus.Text, vbInformation, "PatientTrac Appointment Scheduling"
                   Screen.MousePointer = vbDefault
                End If
            Case "Rescheduled"
                MsgBox "Please enter a reason for the" & " " & cboStatus.Text, vbInformation, "PatientTrac Appointment Scheduling"
                Screen.MousePointer = vbDefault
            Case "Modify Appointment"
                If Val(udtSchedPatient.IntId) > 0 Then
                    Format_Grid
                    Read_ApptGrid
                    fraChange.Visible = True
                    Screen.MousePointer = vbDefault
                End If
        End Select
    End If
    udtAppointment.Status = vbNullString
    If udtSchedPatient.IntId <> vbNullString Then
        udtAppointment.Status = Val(udtSchedPatient.IntId) & "~" & Trim(cboStatus.Text)
        udtSchedPatient.Status = cboStatus.Text
    Else
        mbooCancel = True
    End If
    Exit Sub
End Sub

Private Sub cboStatus_KeyUp(KeyCode As Integer, Shift As Integer)
    If mbooByPass Then mbooByPass = False: Exit Sub
    If cboStatus.Text <> vbNullString Then
        If cboStatus.Text = "Cancelled" Then
            If cboReason.Text = vbNullString Then
                MsgBox "Please enter a reason for the" & " " & cboStatus.Text, vbInformation, "PatientTrac Appointment Scheduling"
                mbooCancel = True
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
    udtAppointment.Status = vbNullString
    udtAppointment.Status = udtSchedPatient.IntId & "~" & Trim(cboStatus.Text)
    Exit Sub
End Sub

Private Sub ccdTime_Click()



    Exit Sub
End Sub


Private Sub ccSchedDate_DateClick(ByVal DateClicked As Date)
    Format_Appt_Lookup
    Get_OpenAppointment
End Sub
Private Function Format_CellTypeEdit_RowSpan(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrNote As String)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .Redraw = False
        .BlockMode = True
            If IsNumeric(vstrSpan) Then lngSpan = vstrSpan
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
              If vlngBackColor > 0 Then .BackColor = vlngBackColor
              If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
              ElseIf lngSpan > 0 Then
                .AddCellSpan vlngCol, vlngRow, 1, lngSpan
              End If
              .SetText vlngCol, vlngRow, Trim(vstrText)
              .Text = Trim(vstrText)
              CellNote = .MaxTextCellHeight
              'If lngSpan <= 1 Then
                    If .RowHeight(vlngRow) < CellNote Then
                      .RowHeight(vlngRow) = Trim(CellNote)
                    End If
               ' End If
              .TypeVAlign = TypeVAlignCenter
                If vstrHAligh = "TypeHAlignCenter" Then
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .TypeHAlign = TypeHAlignLeft
                End If
              If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
              If vstrNote <> vbNullString Then .CellNote = vstrNote
          .BlockMode = False
    End With
End Function



Private Sub cctReferral_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strSQL As String
Dim cnReferral As New ADODB.Connection
Dim rsReferral As New ADODB.Recordset
Dim strName As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim strSuffix As String
Dim lngRefID As Long
Dim strRef As String

    On Error GoTo Error_Load_Referrals
    If Len(cctReferral.Text) < 2 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If cnReferral.State = adStateClosed Then cnReferral.open psConnect(0)
    lstReferral.Clear
    strRef = cctReferral.Text
    strSQL = "Select * FROM PHYSICIAN_REFERRAL WHERE LAST_NAME like '" & strRef & "%' ORDER BY LAST_NAME"
    Set rsReferral = cnReferral.Execute(strSQL)
    With rsReferral
        Do Until .EOF
            lngRefID = 0
            strFirst = vbNullString
            strLast = vbNullString
            strMiddle = vbNullString
            strTitle = vbNullString
            strSuffix = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "REFERRER_ID"
                               lngRefID = Fld.Value
                            Case "LAST_NAME"
                               strLast = Fld.Value
                            Case "FIRST_NAME"
                                strFirst = Fld.Value
                            Case "MIDDLE_NAME"
                                strMiddle = Fld.Value
                            Case "SUFFIX"
                                strSuffix = Fld.Value
                            Case "TITLE"
                                strTitle = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If lngRefID > 0 Then
                strLast = Fix_Name(strLast)
                strFirst = Fix_Name(strFirst)
                strMiddle = Fix_Name(strMiddle)
                If Trim(strTitle) <> vbNullString Then
                    strTitle = Fix_Title(strTitle)
                End If
                If Trim(strSuffix) = vbNullString Then
                    lstReferral.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strTitle)
                Else
                    strSuffix = Fix_Suffix(strSuffix)
                    lstReferral.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strSuffix) & " " & Trim(strTitle)
                End If
                lstReferral.ItemData(lstReferral.NewIndex) = lngRefID
            End If
        Loop
    End With
    If lstReferral.listcount > 0 Then lstReferral.Visible = True
    If cnReferral.State = adStateOpen Then
        cnReferral.Close
        Set cnReferral = Nothing
        Set rsReferral = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
   Exit Sub
Error_Load_Referrals:
    Screen.MousePointer = vbDefault
     Write_Error Err.Description, Err.Number, strSQL, "Load_Referrals", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Sub cctRepeat_Click()
    
    Exit Sub
End Sub


Private Sub cctSearch_Change()

    If Trim(cctSearch) <> vbNullString Then
        Timer1.Enabled = False
        Timer1.Enabled = True
    End If
    Exit Sub
End Sub

Private Sub Read_VendorData(vstrData As String, Optional vlngRef_ID As Long)
Dim strSQL As String
Dim lngRows As Long
Dim strCompany As String
Dim strFirstName As String
Dim strLastName As String
Dim cnReferral As New ADODB.Connection
Dim rsReferral As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strName As String
Dim lngVendorID As Long
Dim strType As String
Dim strSearch As String

    On Error GoTo ERROR_Read_VendorData
    Screen.MousePointer = vbHourglass
    lstListBox.Clear
    If cnReferral.State = adStateClosed Then cnReferral.open psConnect(0)
    If vlngRef_ID > 0 Then
        strType = "VENDOR ID"
    Else
        strType = "LAST NAME"
    End If
    Select Case Trim(strType)
        Case "LAST NAME"
            strSearch = cctSearch.Text
            strSQL = "Search_Vendor_LastName '%" & strSearch & "%'"
        Case "VENDOR ID"
            strSQL = "Get_Vendor_ID " & vlngRef_ID
    End Select
    Set rsReferral = cnReferral.Execute(strSQL)
    With rsReferral
        Do Until .EOF
            Clear_Vendor
            strName = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "VENDOR_NAME"
                                clsVendor.Name = Trim(Fld.Value)
                            Case "VENDOR_CONTACT"
                                clsVendor.Contact = Trim(Fld.Value)
                            Case "CONTACT_TITLE"
                                clsVendor.Contact_Title = Trim(Fld.Value)
                            Case "VENDOR_ID"
                                clsVendor.VendorID = Trim(Fld.Value)
                        End Select
                    End If
                End If
            Next
        .MoveNext
            If Trim(clsVendor.Name) <> vbNullString Then
                strName = Trim(clsVendor.Name)
            End If
            If Trim(clsVendor.Contact) <> vbNullString Then
                strName = strName & ", " & Trim(clsVendor.Contact)
            End If
            If Trim(clsVendor.Contact_Title) <> vbNullString Then
                strName = strName & ", " & Trim(clsVendor.Contact_Title)
            End If
            If Trim(strName) <> vbNullString Then
                lstListBox.AddItem strName
                If IsNumeric(clsVendor.VendorID) Then lstListBox.ItemData(lstListBox.NewIndex) = clsVendor.VendorID
            End If
        Loop
    End With
    If lstListBox.listcount < 0 Then
         strSQL = "Search_Vendor_Contact '" & vstrData & "%'"
        Set rsReferral = cnReferral.Execute(strSQL)
        With rsReferral
            Do Until .EOF
            Clear_Vendor
            strName = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "VENDOR_NAME"
                                clsVendor.Name = Trim(Fld.Value)
                            Case "VENDOR_CONTACT"
                                clsVendor.Contact = Trim(Fld.Value)
                            Case "CONTACT_TITLE"
                                clsVendor.Contact_Title = Trim(Fld.Value)
                            Case "VENDOR_ID"
                                clsVendor.VendorID = Trim(Fld.Value)
                        End Select
                    End If
                End If
            Next
            .MoveNext
                If Trim(clsVendor.Name) <> vbNullString Then
                    strName = Trim(clsVendor.Name)
                End If
                If Trim(clsVendor.Contact) <> vbNullString Then
                    strName = strName & ", " & Trim(clsVendor.Contact)
                End If
                If Trim(clsVendor.Contact_Title) <> vbNullString Then
                    strName = strName & ", " & Trim(clsVendor.Contact_Title)
                End If
                If Trim(strName) <> vbNullString Then
                    lstListBox.AddItem strName
                    If IsNumeric(clsVendor.VendorID) Then lstListBox.ItemData(lstListBox.NewIndex) = clsVendor.VendorID
                End If
            Loop
        End With
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_Read_VendorData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, " Read_VendorData", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Sub cmdAcct_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "FinancialSnapShot"
    frmLoadControl.piSelPatientID = udtSchedPatient.IntId
    frmLoadControl.Show vbModal, Me
    Exit Sub
End Sub

Private Sub cmdChange_Click()
    cctReferalReas.Text = vbNullString
    cboReferBy.Clear
    GetTerms cboReferBy, "sd_Patient_Referral"
    Exit Sub
End Sub

Private Sub cmdChangeUser_Click()
Dim lngUser As Long
    Format_Grid
    cboProvider.Enabled = False
    Read_ApptGrid
    fraChange.Visible = True
    Exit Sub
End Sub

Private Sub cmdClear_Click()
    cctReasReferral.Text = vbNullString
    Exit Sub
End Sub

Private Sub cmdCloseAvail_Click()
    fraGetAppts.Visible = False
End Sub

Private Sub cmdEditPatient_Click()
 Dim strSQL As String
Dim R%
Dim i%
Dim Arr
Dim strName As String
Dim lngPatID As Long

    Clear_PatientStrings
    PatientInfo.mNewPatient = False
    If udtSchedPatient.IntId = vbNullString Then
        udtSchedPatient.IntId = Val(lblID2.Caption)
    End If
    piPatientID = udtSchedPatient.IntId
    lngPatID = udtSchedPatient.IntId
    PatientInfo.Show vbModal
    udtSchedPatient.IntId = lngPatID
    For i = 0 To lstListBox.listcount - 1
        If lstListBox.ItemData(i) = piPatientID Then
            lstListBox.Selected(i) = True
            Exit For
        End If
    Next i
    If cboProvider.listcount <= 0 Then Get_Appt_Providers
    For i = 0 To cboProvider.listcount - 1
         If cboProvider.ItemData(i) = udtAppointment.ProviderID Then
             cboProvider.Text = cboProvider.List(i)
             Exit For
         End If
    Next i
    Get_Info udtAppointment.ProviderID
    For i = 0 To lstListBox.listcount - 1
        If lstListBox.ItemData(i) = Val(udtSchedPatient.IntId) Then
            lstListBox.Selected(i) = True
        End If
    Next i
    lstListBox.Clear
    Exit Sub
End Sub

Private Sub cmdExit_Click()
    fraInfo.Visible = False
    Clear_Fields
    cctFacility.Text = vbNullString
    cctFacility.Visible = False
    lblFac.Visible = False
    cctFacility.BackColor = NormalColor
    piPatientID = 0
    lstListBox.Visible = True
    cctSearch.Visible = True
    lblPatInfo.Visible = True
    lstListBox.Clear
    cmdNewPatient.Visible = True
    Clear_PatientStrings
    mbooByPass = False
    Exit Sub
End Sub

Private Sub cmdFrameClose_Click()
Dim Idx As Integer
Dim lngProv As Long

    For Idx = 0 To cboProvider.listcount - 1
        If cboProvider.List(Idx) = cboProvider.Text Then
            lngProv = cboProvider.ItemData(Idx)
            Exit For
        End If
    Next Idx
    Get_Info lngProv, Val(udtSchedPatient.IntId)
    fraChange.Visible = False
End Sub
Private Sub Format_Grid()
Dim lngCol As Long
Dim lngRow As Long
Dim lngPat As Long

    On Error GoTo Error_Format_Grid
    Screen.MousePointer = vbHourglass
    grdAppointChange.ProcessTab = True
    grdAppointChange.FontBold = True
    grdAppointChange.ForeColor = 0
    If Not IsNull(txtDate.Value) Then ccdRevisedDate.Value = Format(txtDate.Value, "mm/dd/yyyy")
    With grdAppointChange
        .Redraw = False
        .MaxRows = 0
        .MaxCols = 7
        .ColHeaderRows = 4
        .ClearRange 1, 1, grdAppointChange.MaxCols, grdAppointChange.MaxRows, False
        .ColWidth(1) = 10
        .ColWidth(2) = 10
        .ColWidth(3) = 15
        .ColWidth(4) = 20
        .ColWidth(5) = 19
        .ColWidth(6) = 10
        .ColWidth(7) = 20
        lngPat = udtSchedPatient.IntId
        If Trim(udtSchedPatient.IntId) <> vbNullString Then
            Select Case IsProperCase(udtAppointment.SchedType)
                Case "Patient"
                    Get_SchedPatient Val(udtSchedPatient.IntId)
                Case "Vendor", "Business"
                    Get_SchedVendor Val(udtSchedPatient.IntId), udtAppointment.SchedTime
                Case "Legal", "Attorney"
                    Get_Schedule_Attorney Val(udtSchedPatient.IntId)
                Case "Physician"
                    Get_SchedPhysician Val(udtSchedPatient.IntId)
            End Select
        End If
        Format_CellTypeEdit grdAppointChange, SpreadHeader, 1, "9", "APPOINTMENTS FOR" & " " & UCase(udtSchedPatient.FirstName & " " & udtSchedPatient.LastName) & vbNewLine & Format(ccdRevisedDate.Value, "MMMM DD, YYYY"), "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdAppointChange.SetCellBorder 1, SpreadHeader, 9, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdAppointChange
        lngCol = 1
        Format_CellTypeEdit grdAppointChange, SpreadHeader + 2, lngCol, "1", "APPT DATE", "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdAppointChange, SpreadHeader + 2, lngCol, "1", "APPT TIME", "TypeHAlignCenter", "APPT_TIME", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeEdit grdAppointChange, SpreadHeader + 2, lngCol, "1", "APPT TYPE", "TypeHAlignCenter", "PATIENT_TYPE", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeEdit grdAppointChange, SpreadHeader + 2, lngCol, "1", "VISIT TYPE", "TypeHAlignCenter", "PURPOSE_VISIT", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdAppointChange, SpreadHeader + 2, lngCol, "1", "DURATION", "TypeHAlignCenter", "DURATION", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        Format_CellTypeEdit grdAppointChange, SpreadHeader + 2, lngCol, "1", "APPT STATUS", "TypeHAlignCenter", "STATUS", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        Format_CellTypeEdit grdAppointChange, SpreadHeader + 2, lngCol, "1", "COMMENTS", "TypeHAlignCenter", "NOTES", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdAppointChange
        .Redraw = True
    End With
    udtSchedPatient.IntId = lngPat
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Grid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Grid", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
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
             .TypeComboBoxhWnd = fpTerm.Hwnd
             .TypeMaxEditLen = 32000
             .TypeComboBoxList = vstrString
            ' .TypeComboBoxWidth = 1
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
            .TypeVAlign = TypeVAlignCenter
              If vlngBackColor > 0 Then
                 .BackColor = vlngBackColor
              End If
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
          .BlockMode = False
          fpTerm.ListIndex = -1
    End With
End Function


Private Sub cmdNewPatient_Click()
Dim strSQL As String
Dim R%
Dim i%
Dim Arr
Dim strName As String

 
    Clear_PatientStrings
    udtSchedPatient.IntId = vbNullString
    PatientInfo.mNewPatient = True
    PatientInfo.Show vbModal
    udtSchedPatient.IntId = piPatientID
    lstListBox.Clear
    Open_DB psConnect(0)
    strSQL = "Select first_name,last_name, patient_ID from patient Order by Last_Name"
    Set rs = cn.Execute(strSQL)
    With rs
        Do Until .EOF
            If (!Last_Name) <> vbNullString Then
                 If TextFound("''", (!Last_Name), 2, True) Then
                    strName = Replace((!Last_Name), "''", "'")
                Else
                    strName = (!Last_Name)
                End If
            End If
            If (!First_Name) <> vbNullString Then
                 If TextFound("''", (!First_Name), 2, True) Then
                    strName = strName & " " & Replace((!First_Name), "''", "'")
                Else
                    strName = strName & " " & (!First_Name)
                End If
            End If
            lstListBox.AddItem strName
            If IsNumeric((!Patient_ID)) Then lstListBox.ItemData(lstListBox.NewIndex) = (!Patient_ID)
        .MoveNext
        Loop
    End With
    
    Close_DB psConnect(0)
    If cboProvider.listcount <= 0 Then Get_Appt_Providers
    For i = 0 To cboProvider.listcount - 1
         If cboProvider.ItemData(i) = udtAppointment.ProviderID Then
             cboProvider.Text = cboProvider.List(i)
             Exit For
         End If
    Next i
    Get_Info udtAppointment.ProviderID
    If piPatientID <> 0 Then
        For i = 0 To lstListBox.listcount - 1
            If lstListBox.ItemData(i) = Val(udtSchedPatient.IntId) Then
                lstListBox.Selected(i) = True
            End If
        Next i
    End If
    Exit Sub
End Sub




Private Sub cmdPhyPlan_Click()
    Format_Plan
    fraPlan.Visible = True
    Exit Sub
End Sub

Private Sub cmdPlan_Click()
    fraPlan.Visible = False
    Exit Sub
End Sub

Private Sub cmdPrint_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(1) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled.", vbInformation, "PatientTrac Appointment Scheduling"
        mbooCancelPrint = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    With grdAppointChange
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = udtSchedPatient.Name & "_" & "Appointment_" & udtAppointment.SchedDate
        .PrintColor = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintSmartPrint = True
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintBorder = True
        .PrintColHeaders = True
        .PrintRowHeaders = True
        .PrintScalingMethod = PrintScalingMethodSmartPrint
        .PrintMarginLeft = 650
        .PrintMarginRight = 700
        .PrintMarginBottom = 900
        .PrintMarginTop = 900
        .PrintUseDataMax = True
        .PrintColor = True
        .PrintOrientation = PrintOrientationPortrait
        .PrintShadows = False
        .PrintSheet
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdPrint_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Print", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdPsyAdd_Click()
Dim lngRow As Long
Dim strCode As String
Dim bFound As Boolean
Dim bGood As Boolean
Dim strUnit As String
Dim lngTest As Long
Dim bDup As Boolean
Dim strCodeType As String
Dim varCellValue
Dim lngTestRow As Long
Dim strDesc As String
Dim Arr
Dim strTime As String
Dim Idx%
Dim strVisit As String
Dim strVisitCode As String

    'what is code
    With grdProcedure
        .GetText 2, .ActiveRow, varCellValue
        strVisit = varCellValue
        .GetText 1, .ActiveRow, varCellValue
        strVisitCode = varCellValue
    End With
    With grdPsychotherapy
        For lngTestRow = 1 To .MaxRows
            strCode = vbNullString
            strUnit = vbNullString
            strCodeType = vbNullString
            bGood = False
            bDup = False
            .Col = 1
            .Row = lngTestRow
            If .BackColor = grdCurItem Then
                .GetText 1, lngTestRow, varCellValue
                strCode = varCellValue
                .GetText 2, lngTestRow, varCellValue
                strDesc = varCellValue
                .GetText 3, lngTestRow, varCellValue
                strUnit = varCellValue
                If strUnit = vbNullString Then strUnit = "1"
                Exit For
            End If
        Next lngTestRow
    End With
    If strVisit = "90792" Or strVisit = "90791" Then
        Select Case strCode
            Case "90833", "90836", "90838"
            strCode = vbNullString
            MsgBox "Cannot use this Add CPT Code on Initial Visit", vbInformation, "PatientTrac Scheduling"
        End Select
    End If
    If strCode <> vbNullString Then
        Format_CellTypeEdit grdProcedure, grdProcedure.ActiveRow, 3, "1", strDesc, "TypeHAlignCenter", "ADD_DESC"
        Format_CellTypeEdit grdProcedure, grdProcedure.ActiveRow, 4, "1", strCode, "TypeHAlignCenter", "ADD_CPT_CODE"
        Arr = Split(strDesc, " ")
        If UBound(Arr) > 0 Then
            For Idx = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(Idx)) Then
                    strTime = Trim(Arr(Idx))
                    Exit For
                End If
            Next Idx
            If strCode = "90875" Then
                Calc_Duration strVisitCode
                For Idx = 0 To cboDuration.listcount - 1
                    If cboDuration.List(Idx) = Trim(strTime) Then
                        cboDuration.Text = cboDuration.List(Idx)
                        Exit For
                    End If
                Next Idx
            Else
                strTime = strTime & " " & "minutes"
                For Idx = 0 To cboDuration.listcount - 1
                    If cboDuration.List(Idx) = Trim(strTime) Then
                        cboDuration.Text = cboDuration.List(Idx)
                        Exit For
                    End If
                Next Idx
            End If
        End If
    End If
    fraPsycho.Visible = False
    Exit Sub
End Sub

Private Sub cmdPsyClose_Click()
    fraPsycho.Visible = False
    Exit Sub
End Sub


Private Sub cmdRefClose_Click()
Dim Idx As Integer
Dim cnRef As New ADODB.Connection
Dim rsRef As New ADODB.Recordset
Dim strSQL As String
Dim lngReferral As Long
Dim i%

    
    Screen.MousePointer = vbHourglass
    fraRefer.Visible = False
    cctReferalReas.Text = Trim(cctReasReferral.Text)
    cctReferalReas.Visible = True
    cboReferBy.Clear
    If cctReferral.Text <> vbNullString Then
        cboReferBy.AddItem cctReferral.Text
        If clsReferral.ReferralID > 0 Then
            For i = 0 To cboReferBy.listcount - 1
                If cboReferBy.ItemData(i) = clsReferral.ReferralID Then
                    cboReferBy.Text = cboReferBy.List(i)
                    Exit For
                End If
            Next i
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cmdReferralPhys_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "ReferPhys"
    frmLoadControl.Show vbModal, Me
    Exit Sub
End Sub

Private Sub cmdRefReason_Click()

    Exit Sub
End Sub

Private Function Load_Service(vlngFacilityID As Long, Optional vintServCode As Integer)
Dim strSQL As String
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset
Dim lngIndex As Long
Dim Idx As Integer

    On Error GoTo Error_Load_Service
    Screen.MousePointer = vbHourglass
    cboSerPlace.Clear
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "Load_Service_Place"
    Set rsFac = cnFac.Execute(strSQL)
    With rsFac
       Do Until .EOF
            If (!POS_NAME) <> vbNullString Then cboSerPlace.AddItem rsFac.Fields("POS_NAME")
            If Not IsNull(!POS_CODE) Then cboSerPlace.ItemData(cboSerPlace.NewIndex) = (!POS_CODE)
            .MoveNext
        Loop
    End With
    strSQL = "Get_Facility_Place_Service " & vlngFacilityID
    Set rsFac = cnFac.Execute(strSQL)
    With rsFac
        If Not rsFac.EOF Then
            If Not IsNull(!pos) Then
                lngIndex = (!pos)
            End If
        End If
    End With
    If lngIndex > 0 Then
        For Idx = 0 To cboSerPlace.listcount - 1
           If cboSerPlace.ItemData(Idx) = lngIndex Then
                cboSerPlace.Text = cboSerPlace.List(Idx)
                Exit For
           End If
        Next Idx
    End If
    udtEncounter.ServicePlace = cboSerPlace.Text
    udtEncounter.ServicePlaceCode = lngIndex
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set cnFac = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Service:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Service", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cmdRemAddOn_Click()
Dim Arr
Dim cnApptGrid As New ADODB.Connection
Dim rsApptGrid As New ADODB.Recordset
Dim lngProviderID As Long
Dim lngFacilityID As Long
Dim strSQL As String
Dim strTime As String
Dim strDate As String
Dim lngApptID As Long
Dim lngRow As Long
Dim varCellValue
Dim strAdd_Proc As String
Dim strAdd_Proc_Code As String
Dim strPurpose As String
Dim lngDateDiff As Long
Dim strPatientType As String
Dim Idx%

    On Error GoTo Error_cmdRemAddOn
    Screen.MousePointer = vbHourglass
    udtSchedPatient.IntId = Trim(lblID2.Caption)
    If Trim(udtSchedPatient.IntId) <> vbNullString Then
        If Trim(cboFacility.Text) = vbNullString Then
            MsgBox "Please reselect a facility!", vbInformation, "PatientTrac Appointment Scheduling"
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf cboFacility.ListIndex >= 0 Then
            lngFacilityID = cboFacility.ItemData(cboFacility.ListIndex)
        End If
        If udtAppointment.ProviderID > 0 Then
            lngProviderID = udtAppointment.ProviderID
        ElseIf Trim(cboProvider.Text) <> vbNullString Then
            lngProviderID = cboProvider.ItemData(cboProvider.ListIndex)
        End If
        lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), txtDate.Value)
        If lngDateDiff >= 0 And Not IsNull(txtDate.Value) Then
            strDate = Format(txtDate.Value, "mm/dd/yyyy")
        Else
            strDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
        End If
        strTime = Trim(ccdTime.Value)
        strTime = Format(strTime, "hh:mm AM/PM")
        If Trim(cboPatientType.Text) <> vbNullString Then strPatientType = Trim(cboPatientType.Text)
        With grdPsychotherapy
            For lngRow = 1 To .MaxRows
                .BlockMode = True
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = lngRow
                    .Row2 = lngRow
                    .BackColor = NormalColor
                .BlockMode = False
            Next lngRow
        End With
        lngRow = 0
        For lngRow = 1 To grdProcedure.MaxRows
            With grdProcedure
                .Col = 1
                .Row = lngRow
                If .BackColor = grdCurItem Then
                    .GetText 1, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        strPurpose = Trim(varCellValue)
                    End If
                    .GetText 3, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        strAdd_Proc = Trim(varCellValue)
                    End If
                    .GetText 4, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        strAdd_Proc_Code = Trim(varCellValue)
                    End If
                    Exit For
                End If
            End With
        Next lngRow
        If lngRow > 0 Then
            Format_CellTypeEdit grdProcedure, lngRow, 3, "1", "", "TypeHAlignCenter", "ADD_DESC"
            Format_CellTypeEdit grdProcedure, lngRow, 4, "1", "", "TypeHAlignCenter", "ADD_CPT_CODE"
        End If
        If cnApptGrid.State = adStateClosed Then cnApptGrid.open psConnect(0)
        strSQL = "SELECT APPT_ID FROM APPOINTMENTS WHERE PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
            & "PROVIDER_ID = " & lngProviderID & " AND FACILITY_ID = " & lngFacilityID & " AND " _
            & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "' AND " _
            & "CONVERT(VARCHAR(08), APPT_TIME, 108)  = '" & Format(udtAppointment.SchedTime, "hh:mm AM/PM") & "' AND APPT_TYPE = '" & udtAppointment.SchedType & "'"
        Set rsApptGrid = cnApptGrid.Execute(strSQL)
        With rsApptGrid
            If Not .EOF Then
                If Not IsNull(!Appt_ID) Then
                    lngApptID = (!Appt_ID)
                End If
            End If
        End With
        If cnApptGrid.State = adStateClosed Then cnApptGrid.open psConnect(0)
        If lngApptID > 0 Then
             strSQL = "UPDATE APPOINTMENTS SET ADD_PROC_CODE = '', ADD_PROC_DESC = '' WHERE APPT_ID = " & lngApptID
            Set rsApptGrid = cnApptGrid.Execute(strSQL)
        End If
        Calc_Duration strPurpose
        If cboDuration.listcount <= 0 Then Load_Duration
        If Trim(udtEncounter.VisitTime) <> vbNullString Then
            For Idx = 0 To cboDuration.listcount - 1
                If cboDuration.List(Idx) = udtEncounter.VisitTime Then
                    cboDuration.Text = cboDuration.List(Idx)
                    Exit For
                End If
            Next Idx
        End If
    End If
    If cnApptGrid.State = adStateOpen Then
        cnApptGrid.Close
        Set cnApptGrid = Nothing
        Set rsApptGrid = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdRemAddOn:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "cmdRemAddOn", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdRemoveRow_Click()
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim lngPrimKey As Long
Dim lngRow As Long
Dim booFound As Boolean

    On Error GoTo Error_DeleteRow
    Screen.MousePointer = vbHourglass
    With grdAppointChange
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                booFound = True
                lngPrimKey = .GetRowItemData(lngRow)
                Exit For
            End If
        Next lngRow
    End With
    If lngPrimKey = 0 And booFound = False Then
        MsgBox "Please select the desired removal row first!", vbInformation, "PatientTrac Appointment Management"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If lngPrimKey > 0 And booFound Then
        If cnBlock.State = adStateClosed Then cnBlock.open psConnect(0)
        strSQL = "Delete_Appointment_Appt_ID " & lngPrimKey
        Set rsBlock = cnBlock.Execute(strSQL)
        If cnBlock.State = adStateClosed Then
            cnBlock.Close
            Set cnBlock = Nothing
            Set rsBlock = Nothing
        End If
    End If
    If lngRow > 0 And booFound Then
        With grdAppointChange
            .DeleteRows lngRow, 1
            .MaxRows = .MaxRows - 1
        End With
    End If
    cmdApply.Enabled = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_DeleteRow:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "DeleteRow", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdRepeat_Click()
Dim i As Integer
Dim lngRow As Long
Dim varCellValue

    udtRepeatAppt.Comment = Trim(txtComments.Text)
    udtRepeatAppt.Comment = Replace(udtRepeatAppt.Comment, "'", "''")
    udtRepeatAppt.Status = Trim(cboStatus.Text)
    For lngRow = 1 To grdProcedure.MaxRows
        With grdProcedure
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                .GetText 1, lngRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    udtRepeatAppt.PurposeVisit = Trim(varCellValue)
                End If
                Exit For
            End If
        End With
    Next lngRow
    If Not IsNull(ccdTime.Value) <> vbNullString Then udtRepeatAppt.StartTime = Format(ccdTime.Value, "hh:mm AM/PM")
    If Not IsNull(txtDate.Value) Then udtRepeatAppt.StartDate = Format(txtDate.Value, "mm/dd/yyyy")
    If Trim(cboDuration.Text) <> vbNullString Then udtRepeatAppt.Duration = Trim(cboDuration.Text)
    If udtAppointment.ProviderID > 0 Then udtRepeatAppt.ProviderID = udtAppointment.ProviderID
    If udtFacility.FacID > 0 Then udtRepeatAppt.FacilityID = udtFacility.FacID
    frmMultiAppt.Show vbModal
    Get_Info udtAppointment.ProviderID, Val(udtSchedPatient.IntId)
    Exit Sub
End Sub

Private Sub cmdSeeAvail_Click()

    Format_Appt_Lookup
    ccSchedDate.Value = txtDate.Value
    Get_OpenAppointment
    fraGetAppts.Visible = True
End Sub

Private Sub cmdSelectAvail_Click()
Dim i%

    If Not Check_OverBook_Change(clsWriteAppt.SchedTime) Then
        txtDate.Value = clsWriteAppt.SchedDate
        ccdTime.Value = clsWriteAppt.SchedTime
        fraGetAppts.Visible = False
    End If
    Exit Sub
End Sub

Private Sub cmdShowAll_Click()
Dim strType As String

    Select Case cboApptType.Text
        Case "Patient", "Partner", "Guardian"
            udtAppointment.SchedType = udtSchedPatient.PatientType
        Case "Vendor", "Physician"
            udtAppointment.SchedType = "Business"
        Case "Legal"
            udtAppointment.SchedType = "Legal"
    End Select
    Load_Visit
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim lngPrimKey As Long
Dim lngRow As Long
Dim booFound As Boolean
Dim varCellValue
Dim strDuration As String
Dim strNow As String
Dim Arr
Dim lngCol As Long
Dim lngActive As Long


    On Error GoTo Error_cmdUpdate
    Screen.MousePointer = vbHourglass
    strNow = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    With grdAppointChange
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                booFound = True
                lngPrimKey = .GetRowItemData(lngRow)
                Exit For
            End If
        Next lngRow
    End With
    If lngPrimKey = 0 And booFound = False Then
        MsgBox "Please select the desired row first!", vbInformation, "PatientTrac Appointment Management"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If lngPrimKey > 0 And booFound Then
        Arr = Split(LookUpCellTag("APPT_DATE"), "|")
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                grdAppointChange.GetText lngCol, lngRow, varCellValue
                udtAppointment.SchedDate = varCellValue
            End If
        End If
        Arr = Split(LookUpCellTag("APPT_TIME"), "|")
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                grdAppointChange.GetText lngCol, lngRow, varCellValue
                udtAppointment.SchedTime = varCellValue
            End If
        End If
        Arr = Split(LookUpCellTag("PROVIDER_ID"), "|")
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                If grdAppointChange.CellNote <> vbNullString Then udtAppointment.ProviderID = grdAppointChange.CellNote
            End If
        End If
        Arr = Split(LookUpCellTag("PATIENT_TYPE"), "|")
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                grdAppointChange.GetText lngCol, lngRow, varCellValue
                udtSchedPatient.PatientType = varCellValue
            End If
        End If
        Arr = Split(LookUpCellTag("PURPOSE_VISIT"), "|")
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                grdAppointChange.GetText lngCol, lngRow, varCellValue
                udtAppointment.VisitType = varCellValue
            End If
        End If
        Arr = Split(LookUpCellTag("DURATION"), "|")
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                grdAppointChange.GetText lngCol, lngRow, varCellValue
                udtSchedPatient.Duration = varCellValue
            End If
        End If
        Arr = Split(LookUpCellTag("STATUS"), "|")
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                grdAppointChange.GetText lngCol, lngRow, varCellValue
                udtAppointment.Status = varCellValue
            End If
        End If
        Arr = Split(LookUpCellTag("NOTES"), "|")
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                grdAppointChange.GetText lngCol, lngRow, varCellValue
                udtSchedPatient.ApptComment = varCellValue
            End If
        End If
        lngActive = 1
        If udtSchedPatient.Status <> vbNullString Then
            Select Case udtSchedPatient.Status
                Case "Cancelled", "No Show"
                    lngActive = 0
                Case "Active"
                    lngActive = 1
                    udtSchedPatient.Status = vbNullString
            End Select
        End If
        If lngActive = 0 And udtSchedPatient.Status = vbNullString Then
            MsgBox "Cannot Update an Inactive Appointment, if you want to modify this appointment please change the Status before saving!", vbCritical, "PatientTrac Scheduling"
            cmdApply.Enabled = False
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        PsUserID = Get_PsUserID(psConnect(0))
        udtSchedPatient.ApptComment = Replace(udtSchedPatient.ApptComment, "'", "''")
        clsWriteAppt.Duration = udtSchedPatient.Duration
        If Check_OverBook_Change(udtAppointment.SchedTime) Then
            cmdApply.Enabled = False
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If cnBlock.State = adStateClosed Then cnBlock.open psConnect(0)
            strSQL = "Update_Appointment_Schedule_Routine  " & lngActive & ",'" & udtAppointment.VisitType & "','" _
                & udtSchedPatient.Duration & "','" & udtSchedPatient.ApptComment & "','" _
                & udtAppointment.SchedTime & "','" & udtSchedPatient.PatientType & "','" & strNow & "','" _
                & udtSchedPatient.Status & "', '" & udtAppointment.SchedDate & "', '" _
                & PsUserID & "','" & udtAppointment.SchedType & "'," & lngPrimKey
        Set rsBlock = cnBlock.Execute(strSQL)
            'update reason
        If cnBlock.State = adStateClosed Then cnBlock.open psConnect(0)
        If udtAppointment.Status <> "Cancelled" Then
            strSQL = "Update_Appointments_Repeat_Boolean " & lngPrimKey & ", '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'"
            Set rsBlock = cnBlock.Execute(strSQL)
            strSQL = "UPDATE APPOINTMENTS SET REASON = '' WHERE APPT_ID = " & lngPrimKey
            Set rsBlock = cnBlock.Execute(strSQL)
        End If
        MsgBox "Appointment has been succesfully updated to" & " " & udtAppointment.SchedTime & ".", vbInformation, "PatientTrac Scheduling"
        If cnBlock.State = adStateClosed Then
            cnBlock.Close
            Set cnBlock = Nothing
            Set rsBlock = Nothing
        End If
    End If
    cmdApply.Enabled = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdUpdate:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "cmdUpdate", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub grdAppointChange_Click(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long

   If Row < 1 Then Exit Sub
    With grdAppointChange
        For lngRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                .BackColor = NormalColor
            .BlockMode = False
        Next lngRow
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = grdCurItem
        .BlockMode = False
    End With
    Exit Sub
End Sub


Private Sub grdAppointChange_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim lngRow As Long
Dim lngCol As Long
Dim Idx%
Dim varCellValue
Dim Arr
Dim strList As String
Dim bRet As Boolean
Dim strVisit As String
Dim strDuration As String
Dim strPatType As String

    On Error GoTo Error_grdAppointChange_ComboCloseUp
    If mbooByPass Then mbooByPass = False: Exit Sub
    If Row = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    clsGridFormat.CellTag = vbNullString
    With grdAppointChange
        .Col = Col
        .Row = Row
        If .BackColor = NormalColor Then
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                .BackColor = grdCurItem
            .BlockMode = False
        End If
        If .CellTag <> vbNullString Then clsGridFormat.CellTag = .CellTag
        Select Case clsGridFormat.CellTag
            Case "PATIENT_TYPE"
                .GetText Col, Row, varCellValue
                strPatType = varCellValue
                Arr = Split(LookUpCellTag("PURPOSE_VISIT"), "|")
                If UBound(Arr) > 0 Then
                    lngCol = Arr(0)
                    If lngCol > 0 Then
                        Format_CellTypeCombo grdAppointChange, Row, lngCol, "1", Get_VisitGrid(strPatType), "TypeHAlignCenter", "PURPOSE_VISIT"
                    End If
                End If
            Case "STATUS"
                .GetText Col, Row, varCellValue
                udtSchedPatient.Status = varCellValue
            Case "PURPOSE_VISIT"
            .GetText Col, Row, varCellValue
            strVisit = varCellValue
            If Trim(strVisit) <> vbNullString Then
                If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
                strSQL = "SELECT APPT_TIME FROM APPOINTMENT_TYPE WHERE APPT_CODE " _
                    & "like '" & strVisit & "%' AND PROVIDER_TITLE = '" & Replace(clsProvider.Title, ".", "") & "'"
                Set rsCodes = cnCodes.Execute(strSQL)
                If Not rsCodes.EOF Then
                    strDuration = rsCodes.Fields("APPT_TIME")
                End If
                If Trim(strDuration) <> vbNullString Then
                    Arr = Split(LookUpCellTag("DURATION"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            mbooByPass = True
                            Set_GridCurSel grdAppointChange, strDuration, lngCol, Row, "1", "DURATION"
                            mbooByPass = False
                        End If
                    End If
                End If
            End If
            Case "APPT_TIME"
                .GetText Col, Row, varCellValue
                strList = Format(varCellValue, "hh:mm AM/PM")
                varCellValue = vbNullString
                Arr = Split(LookUpCellTag("DURATION"), "|")
                If UBound(Arr) > 0 Then
                    lngCol = Arr(0)
                    If lngCol > 0 Then
                        .GetText lngCol, Row, varCellValue
                    End If
                End If
                If varCellValue <> vbNullString Then
                    bRet = Validate_ProvAppt(strList, CStr(varCellValue))
                    If bRet Then
                        MsgBox "Appointment has been reset to current appointment!", vbInformation, "PatientTrac Scheduling"
                        Format_Grid
                        Read_ApptGrid
                    Else
                        lngCol = 0
                        Arr = Split(LookUpCellTag("APPT_DATE"), "|")
                        If UBound(Arr) > 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                mbooByPass = True
                                Format_CellTypeEdit grdAppointChange, Row, lngCol, "1", Format(ccdRevisedDate.Value, "mm/dd/yyyy"), "TypeHAlignCenter", "APPT_DATE"
                                grdAppointChange.SetCellBorder lngCol, Row, lngCol, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                mbooByPass = True
                            End If
                        End If
                    End If
                Else
                    MsgBox "Please select duration first!", vbInformation, "PatientTrac Appointments"
                    Exit Sub
                End If
        End Select
    End With
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_grdAppointChange_ComboCloseUp:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "grdAppointChange_ComboCloseUp", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub grdAvailAppt_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue

    With grdAvailAppt
        .Col = 1
        .Row = Row
        If .BackColor = GridBlue Then
            MsgBox "Cannot schedule, time has been blocked for this provider!", vbInformation, "PatientTrac Reschedule"
            Exit Sub
        End If
        For lngRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                .BackColor = NormalColor
            .BlockMode = False
            .SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Next lngRow
        lngRow = .ActiveRow
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = grdCurItem
        .BlockMode = False
        .GetText 0, Row, varCellValue
        clsWriteAppt.SchedTime = varCellValue
        clsWriteAppt.SchedDate = Format(ccSchedDate.Value, "mm/dd/yyyy")
    End With
End Sub


Private Sub grdProcedure_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
Dim lngRow As Long
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim strSQL As String
Dim strFacility As String
Dim varCellValue
Dim Idx As Integer
Dim lngFac As Long
Dim strDesc As String
Dim Arr


    On Error GoTo Error_grdProcedure_ComboCloseUp
    Screen.MousePointer = vbHourglass
    With grdProcedure
        If Col = 5 Then
            .GetText Col, Row, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                strFacility = Trim(varCellValue)
            End If
        End If
        For lngRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                .BackColor = NormalColor
            .BlockMode = False
        Next lngRow
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = grdCurItem
        .BlockMode = False
        .GetText 1, Row, varCellValue
        udtEncounter.Appt_Type = varCellValue
        udtAppointment.VisitType = Left(udtEncounter.Intrvn_CD, 5)
        Calc_Duration udtEncounter.Appt_Type
        LookUpCPTCode udtEncounter.Appt_Type
        If Trim(udtEncounter.EncounterType) <> vbNullString Then
            For Idx = 0 To cboPatientType.listcount - 1
                If cboPatientType.List(Idx) = udtEncounter.EncounterType Then
                    mbooByPass = True
                    cboPatientType.Text = cboPatientType.List(Idx)
                    Exit For
                End If
            Next Idx
        End If
        If udtEncounter.VisitTime <> vbNullString Then
            With grdProcedure
                .GetText 3, Row, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    strDesc = varCellValue
                     Arr = Split(strDesc, " ")
                    If UBound(Arr) > 0 Then
                        For Idx = LBound(Arr) To UBound(Arr)
                            If IsNumeric(Arr(Idx)) Then
                                udtEncounter.VisitTime = Trim(Arr(Idx))
                                Exit For
                            End If
                        Next Idx
                        udtEncounter.VisitTime = udtEncounter.VisitTime & " " & "minutes"
                    End If
                End If
            End With
            If cboDuration.listcount <= 0 Then Load_Duration
            For Idx = 0 To cboDuration.listcount - 1
                If cboDuration.List(Idx) = udtEncounter.VisitTime Then
                    cboDuration.Text = cboDuration.List(Idx)
                    Exit For
                End If
            Next Idx
        End If
        lblCPT.Visible = True
        lblCode.Visible = True
        lblCode.Caption = vbNullString
        udtEncounter.Intrvn_CD = Left(udtEncounter.Intrvn_CD, 5)
        If udtEncounter.Intrvn_CD <> vbNullString Then
            lblCode.Caption = udtEncounter.Intrvn_CD
        Else
            LookUpCPTCode udtEncounter.Appt_Type
            lblCode.Caption = udtEncounter.Intrvn_CD
        End If
        udtAppointment.VisitType = udtEncounter.Appt_Type
    End With
   If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_grdProcedure_ComboCloseUp:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "grdProcedure_ComboCloseUp", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub grdProcedure_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim Idx As Integer
Dim varCellValue
Dim lngRow As Long
Dim lngFac As Long
Dim strDesc As String
Dim Arr
Dim strCode As String


    On Error GoTo Error_grdProcedure_DblClick
    Screen.MousePointer = vbHourglass
    mbooListValid = False
    With grdProcedure
        For lngRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                .BackColor = NormalColor
            .BlockMode = False
        Next lngRow
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = grdCurItem
        .BlockMode = False
        clsWriteAppt.CPT_Code = vbNullString
        If IsProperCase(udtAppointment.SchedType) = "Patient" Or udtAppointment.SchedType = vbNullString Then
            .GetText 2, Row, varCellValue
            clsWriteAppt.CPT_Code = varCellValue
            lblCode.Caption = clsWriteAppt.CPT_Code
            If Col = 3 Then
                .Col = 1
                .Row = Row
                If .BackColor = grdCurItem Then
                    Format_PsychoTherapy
                    Load_Codes_Psychotherapy
                    .GetText 4, Row, varCellValue
                    strCode = Trim(varCellValue)
                    With grdPsychotherapy
                        For lngRow = 1 To .MaxRows
                            .GetText 1, lngRow, varCellValue
                            If Trim(varCellValue) = Trim(strCode) Then
                                .BlockMode = True
                                    .Col = 1
                                    .Col2 = 3
                                    .Row = lngRow
                                    .Row2 = lngRow
                                    .BackColor = grdCurItem
                                .BlockMode = False
                                Exit For
                            End If
                        Next lngRow
                    End With
                    fraPsycho.Visible = True
                    Exit Sub
                End If
            End If
        End If
    End With
    Procedure_Grid_Action
    mbExistSelection = True
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_grdProcedure_DblClick:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "grdProcedure_DblClick", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Function Format_PsychoTherapy()
    Screen.MousePointer = vbHourglass
    With grdPsychotherapy
        .Redraw = False
       .ClearRange 0, SpreadHeader, grdPsychotherapy.MaxCols, grdPsychotherapy.MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .RowHeadersShow = False
        .MaxCols = 3
        .ColWidth(1) = 19
        .ColWidth(2) = 40
        .ColWidth(3) = 12
        Format_CellTypeStatic grdPsychotherapy, SpreadHeader, 1, "3", "MENTAL HEALTH CODES", "TypeHAlignCenter", "", HeaderColor
        Format_CellTypeStatic grdPsychotherapy, SpreadHeader + 1, 1, "1", "CPT CODE", "TypeHAlignCenter", "CPT CODE", HeaderColor
        Format_CellTypeStatic grdPsychotherapy, SpreadHeader + 1, 2, "1", "CPT DESC", "TypeHAlignCenter", "CPT DESC", HeaderColor
        Format_CellTypeStatic grdPsychotherapy, SpreadHeader + 1, 3, "1", "CPT UNITS", "TypeHAlignCenter", "CPT UNITS", HeaderColor
        grdPsychotherapy.SetCellBorder 1, SpreadHeader, 6, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdPsychotherapy.SetCellBorder 1, SpreadHeader + 1, 1, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdPsychotherapy.SetCellBorder 2, SpreadHeader + 1, 2, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdPsychotherapy.SetCellBorder 3, SpreadHeader + 1, 3, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 2, grdPsychotherapy
        grdPsychotherapy.SetCellBorder 1, SpreadHeader + 2, grdPsychotherapy.MaxCols, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdPsychotherapy.MaxRows = 0
        Load_Codes_Psychotherapy
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Function
End Function
Private Function Load_Codes_Psychotherapy()
Dim bFound As Boolean
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim lngRow As Long
Dim strData As String
Dim strCode As String
Dim bExit As Boolean
Dim strProvider As String

    On Error GoTo Error_Load_Codes_Psychotherapy
    Screen.MousePointer = vbHourglass
    strProvider = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
    If strProvider <> vbNullString Then
       Select Case UCase(Replace(strProvider, ".", ""))
            Case "MD", "DO", "APRN", "CRNP", "PA", "ARNP"
                bFound = True
                If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
                strSQL = "Load_Psychotherapy_AddOn  '" & Replace(clsProvider.Title, ".", "") & "'"
                Set rsCodes = cnCodes.Execute(strSQL)
                With rsCodes
                    Do Until .EOF
                        bExit = False
                        strData = vbNullString
                        strCode = vbNullString
                        If Not IsNull((!CPT_Code)) Then strCode = (!CPT_Code)
                        If Not IsNull((!CPT_Name)) Then
                            If Trim(!CPT_Name) <> vbNullString Then strData = (!CPT_Name)
                        End If
                        .MoveNext
                        Select Case Trim(strCode)
                            Case "90791", "90792", "99211", "99212", "99213", "99214", "99215", "90832", _
                                "90834", "90837", "90839", "90846", "90847", "90849", "90853", "99255", "99213", "99212", _
                                "99222", "99221", "90819", "90818", "99233", "99232", "99231", "90862", "99253", "99252", "99251", _
                                "99263", "99262", "99261", "90871", "90882", "90887", "99367", "99239", "99254", "99264", "99265", _
                                "99223"
                                bExit = True
                        End Select
                        If Not bExit Then
                            lngRow = lngRow + 1
                            grdPsychotherapy.MaxRows = lngRow
                            Format_CellTypeStatic grdPsychotherapy, lngRow, 1, "1", strCode, "TypeHAlignCenter", "CPT_CODE"
                            Format_CellTypeStatic grdPsychotherapy, lngRow, 2, "1", strData, "TypeHAlignLeft", "CPT_NAME"
                            Format_CellTypeEdit grdPsychotherapy, lngRow, 3, "1", "1", "TypeHAlignCenter", "CPT_UNITS"
                            grdPsychotherapy.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            grdPsychotherapy.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            grdPsychotherapy.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
'                        End If
                    Loop
                End With
        End Select
        If Not bFound Then
            If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
            strSQL = "Load_Psychotherapy_AddOn  '" & Replace(clsProvider.Title, ".", "") & "'"
            Set rsCodes = cnCodes.Execute(strSQL)
            With rsCodes
                Do Until .EOF
                    strData = vbNullString
                    strCode = vbNullString
                    If Not IsNull((!CPT_Code)) Then strCode = (!CPT_Code)
                    If Not IsNull((!CPT_Name)) Then
                        If Trim(!CPT_Name) <> vbNullString Then strData = (!CPT_Name)
                    End If
                    .MoveNext
                    If Trim(strCode) = "90836" Then
                        If Authorized_Prescriber(udtAppointment.ProviderID) Then
                            lngRow = lngRow + 1
                            grdPsychotherapy.MaxRows = lngRow
                            Format_CellTypeStatic grdPsychotherapy, lngRow, 1, "1", strCode, "TypeHAlignCenter", "CPT_CODE"
                            Format_CellTypeStatic grdPsychotherapy, lngRow, 2, "1", strData, "TypeHAlignLeft", "CPT_NAME"
                            Select Case strCode
                                Case "90832", "90834", "90837", "90845", "90846", "90847", "90849", "90853"
                                   Format_CellTypeStatic grdPsychotherapy, lngRow, 3, "1", "1", "TypeHAlignLeft", "CPT_UNITS"
                                   grdPsychotherapy.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                Case "90840", "90785"
                                    Format_CellTypeEdit grdPsychotherapy, lngRow, 3, "1", " ", "TypeHAlignLeft", "CPT_UNITS"
                                    grdPsychotherapy.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End Select
                            grdPsychotherapy.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            grdPsychotherapy.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                    Else
                         Select Case strCode
                            Case "90840", "90785"
                                lngRow = lngRow + 1
                                grdPsychotherapy.MaxRows = lngRow
                                Format_CellTypeStatic grdPsychotherapy, lngRow, 1, "1", strCode, "TypeHAlignCenter", "CPT_CODE"
                                Format_CellTypeStatic grdPsychotherapy, lngRow, 2, "1", strData, "TypeHAlignLeft", "CPT_NAME"
                                grdPsychotherapy.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                grdPsychotherapy.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                Format_CellTypeEdit grdPsychotherapy, lngRow, 3, "1", "1", "TypeHAlignLeft", "CPT_UNITS"
                                grdPsychotherapy.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End Select
                    End If
                Loop
            End With
        End If
    End If
    If cnCodes.State = adStateClosed Then
        cnCodes.open
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
     Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Codes_Psychotherapy:
     Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Codes_Psychotherapy", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function

Private Sub grdPsychotherapy_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue
Dim lngRow As Long
Dim lngCol As Long
Dim strEnterCode As String
Dim bFound As Boolean
Dim strVisitCode As String
Dim bCancel As Boolean
Dim bGood As Boolean
Dim lngTest As Long
Dim strCode As String
Dim strProvider As String

    If Row = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    strProvider = Get_Provider_Type(UCase(Replace(clsProvider.Title, ".", "")))
    Select Case strProvider
         Case "MD", "DO", "ARNP", "CRNP", "PA", "APRN"
             bFound = True
     End Select
    strVisitCode = lblCode.Caption
    With grdPsychotherapy
        For lngRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                .BackColor = NormalColor
            .BlockMode = False
        Next lngRow
        .GetText 1, Row, varCellValue
        strEnterCode = varCellValue
        For lngRow = 1 To .MaxRows
            .GetText 1, lngRow, varCellValue
            If Trim(varCellValue) = strEnterCode Then
              Exit For
            ElseIf Trim(varCellValue) = "90833" Or _
                Trim(varCellValue) = "90836" Or _
                Trim(varCellValue) = "90838" Or _
                Trim(varCellValue) = "90863" Then
                If strEnterCode <> Trim(varCellValue) Then
                    If Trim(lblCPT.Caption) = "90833" Or _
                        Trim(lblCPT.Caption) = "90836" Or _
                        Trim(lblCPT.Caption) = "90838" Or _
                        Trim(lblCPT.Caption) = "90863" Then
                    End If
                End If
            End If
            If bFound Then
                  Select Case strEnterCode
                    Case "90840"
                        If strVisitCode <> "90839" Then
                            MsgBox "90840 Add-On can only be added to a 98839 Psychotherapy for Crisis visit code!", vbInformation, "PatientTrac EM Coding"
                            bCancel = True
                            Exit For
                        End If
                      Case "90863" '
                        If strVisitCode = "99211" Or _
                            strVisitCode = "99212" Or _
                            strVisitCode = "99213" Or _
                            strVisitCode = "99214" Or _
                            strVisitCode = "99215" Then
                            MsgBox "Cannot bill the medication management add-on (90863) code with an" & " " & "EM Visit Code!", vbInformation, "PatientTrac EM Coding"
                            bCancel = True
                            Exit For
                        End If
                        If Not bCancel Then
                            If Not TextFound("90832", strVisitCode, Len("90832"), True) Or _
                                Not TextFound("90834", strVisitCode, Len("90834"), True) Or _
                                Not TextFound("90837", strVisitCode, Len("90837"), True) Then
                                    MsgBox "Cannot bill the medication management add-on (90863) code with an" & " " & strVisitCode & " " & "visit code!", vbInformation, "PatientTrac EM Coding"
                                    bCancel = True
                                    Exit For
                            End If
                        End If
                      Case "90785"
                              Select Case strVisitCode
                                  Case "90839", "90840", "90846", "990847", "90849"
                                      MsgBox "Cannot bill the 90785 code with" & " " & strVisitCode & "!", vbInformation, "PatientTrac EM Coding"
                                      bCancel = True
                                      Exit For
                                  Case "99212", "99213", "99214", "99215"
                                      With grdPsychotherapy
                                          For lngTest = 1 To .MaxRows
                                              .Col = 1
                                              .Row = lngTest
                                              If .BackColor = grdCurItem Then
                                                  .GetText 1, lngTest, varCellValue
                                                  If Trim(varCellValue) = "90833" Or _
                                                      Trim(varCellValue) = "90836" Or _
                                                      Trim(varCellValue) = "90838" Or _
                                                      Trim(varCellValue) = "90863" Then
                                                          bGood = True
                                                  End If
                                              End If
                                              If bGood Then
                                                  Exit For
                                              End If
                                          Next lngTest
                                          If Not bGood Then
                                              MsgBox "Cannot bill the 90785 code with" & " " & strVisitCode & " " & "EM Code must include psychotherapy add-on!", vbInformation, "PatientTrac EM Coding"
                                              bCancel = True
                                              Exit For
                                          End If
                                      End With
                                  End Select
                      End Select
            Else
                'Non Physician Rules
                  Select Case strEnterCode
                      Case "90863"
                        If Not Authorized_Prescriber(clsEncounter.ProviderID) Then
                          MsgBox "Currently not authorized to write prescriptions, if error contact Adminitrator", vbInformation, "PatientTrac EM Coding"
                          bCancel = True
                        ElseIf Not TextFound("90832", strVisitCode, Len("90832"), True) Or _
                            Not TextFound("90834", strVisitCode, Len("90834"), True) Or _
                            Not TextFound("90837", strVisitCode, Len("90837"), True) Then
                                MsgBox "Cannot bill the medication management add-on (90863) code with an" & " " & strVisitCode & " " & "visit code!", vbInformation, "PatientTrac EM Coding"
                                bCancel = True
                                Exit For
                        End If
                    Case "90840"
                        If strVisitCode <> "90839" Then
                            MsgBox "90840 Add-On can only be added to a 98839 Psychotherapy for Crisis visit code!", vbInformation, "PatientTrac EM Coding"
                            bCancel = True
                            Exit For
                        End If
                      Case "90785"
                            Select Case strVisitCode
                                Case "90839", "90840", "90846", "990847", "90849"
                                    MsgBox "Cannot bill the 90785 add-code with" & " " & strVisitCode & " " & "visit code!", vbInformation, "PatientTrac EM Coding"
                                    bCancel = True
                                    Exit For
                            End Select
                End Select
            End If
        Next lngRow
        If Not bCancel Then
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                .BackColor = grdCurItem
            .BlockMode = False
        End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub grdPsychotherapy_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    With grdPsychotherapy
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = NormalColor
        .BlockMode = False
    End With
End Sub


Private Sub lstListBox_ItemCheck(Item As Integer)
Dim i%
Dim strSQL
Dim lngPatientID As Long
Dim iAnswer As Integer
Dim strDate As String

    On Error GoTo Error_lstListBox_ItemCheck
    If mbooByPass Then mbooByPass = False: Exit Sub
    If lstListBox.listcount < 0 Then
        MsgBox "Please search for Patient first!", vbCritical, "PatientTrac Schedule Patient"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If lstListBox.listcount = 0 Then
        If lstListBox.List(Item) = vbNullString Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    cctFacility.Text = vbNullString
    cctFacility.Visible = False
    lblFac.Visible = False
    cctFacility.BackColor = NormalColor
    piPatientID = 0
    If mbooByPass Then mbooByPass = False: Exit Sub
    mbExistSelection = False
    cctSearch.Text = vbNullString
    cmdApply.Enabled = True
    If lstListBox.Selected(Item) Then
        If lstListBox.ItemData(Item) > 0 Then
            lngPatientID = lstListBox.ItemData(Item)
            mbooByPass = True
        End If
    End If
    If lngPatientID > 0 Then
        mbLookupPatient = True
        udtSchedPatient.IntId = lngPatientID
        piPatientID = lngPatientID
        Set_PatientFrame lngPatientID
        If clsEDI_Client.Site_ID = vbNullString Then
            Get_Edi_Static
        End If
        Read_Patient_Status lngPatientID
        If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
            Get_Facilities Val(clsFacStatus.FacilityID)
            cctFacility.Visible = True
            lblFac.Visible = True
        End If
        If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
            cctFacility.Text = Trim(clsFacStatus.StatusDesc)
        End If
        cctFacility.BackColor = grdCurItem
    End If
    cmdApply.Enabled = True
    mbooByPass = False
    mbLookupPatient = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_lstListBox_ItemCheck:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "lstListBox_ItemCheck", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub



Private Sub lstReferral_ItemCheck(Item As Integer)
Dim Idx%
    If mbooByPass Then mbooByPass = False: Exit Sub
    cctReferral.Text = vbNullString
    clsReferral.ReferralID = 0
    clsReferral.ReferralID = lstReferral.ItemData(Item)
    lstReferral.Visible = False
    If clsReferral.ReferralID > 0 Then
        Get_Referral clsReferral.ReferralID
        cctReferral.Text = clsReferral.Physician
    End If
    Exit Sub
End Sub



Private Sub mnu_Close_Click()
    frmPatientAppointment.mBeenHere = True
    If cboStatus.Text = vbNullString Then udtAppointment.Status = "New Appointment"
    Clear_PatientStrings
    Unload Me
    centerForm frmPatientAppointment
    If frmPatientAppointment.WindowState = 1 Then
        frmPatientAppointment.WindowState = 0
    End If
    Exit Sub
End Sub

Private Sub mnu_Save_Click()
Dim i%
Dim mbooReason As Boolean
Dim strPurpose As String
Dim strDuration As String
Dim mbooSave As Boolean
Dim lngRow As Long
Dim varCellValue
Dim Idx%
Dim Arr

    On Error GoTo Error_Apply
    If cmdApply.Enabled = False Then Exit Sub
    If fraPsycho.Visible Then
        MsgBox "Please close the Add On Grid prior to saving appointment!", vbInformation, "PatientTrac Scheduling"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Arr = Split(txtComments.Text, vbNewLine)
    If UBound(Arr) > 0 Then
        txtComments.Text = vbNullString
        For Idx = LBound(Arr) To UBound(Arr)
            If Trim(Arr(Idx)) <> vbNullString And Len(Trim(Arr(Idx))) > 2 Then
                If txtComments.Text = vbNullString Then
                    txtComments.Text = Trim(Arr(Idx))
                Else
                    txtComments.Text = txtComments.Text & vbNewLine & Trim(Arr(Idx))
                End If
            End If
            
        Next Idx
    End If
    mbooAuthSequence = True
    udtSchedPatient.Current_Location = vbNullString
    If IsNull(ccdTime.Value) Then
        MsgBox "Please select appointment time!", vbCritical, "PatientTrac Appointment Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If cboFacility.Text = vbNullString Then
        MsgBox "Please select facility!", vbInformation, "PatientTrac Appointment Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strPurpose = "Purpose"
    For lngRow = 1 To grdProcedure.MaxRows
        With grdProcedure
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                mbooReason = True
                strPurpose = vbNullString
                .GetText 5, lngRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    udtSchedPatient.Current_Location = Trim(varCellValue)
                End If
                Exit For
            End If
        End With
    Next lngRow
    If udtSchedPatient.Current_Location = vbNullString Then
        udtSchedPatient.Current_Location = "MENTAL HEALTH"
    End If
    If cboProvider.Text = vbNullString Then
        MsgBox "Please select a provider!", vbInformation, "PatientTrac Appointment Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If cboDuration.Text = vbNullString Then strDuration = "Duration"
    If cboStatus.Text <> vbNullString Then
        If udtAppointment.Status = vbNullString Then
            udtAppointment.Status = Trim(cboStatus.Text)
        End If
    End If
    If SaveData Then
        frmPatientAppointment.mBeenHere = True
        Screen.MousePointer = vbDefault
        mstrValidAppoint = vbNullString: mbooAlreadyBook = False: mbooCancel = False
        fraInfo.Visible = False
        Clear_PatientStrings
        Clear_clsWriteAppt
        Screen.MousePointer = vbDefault
        Unload Me
        frmPatientAppointment.Show
        centerForm frmPatientAppointment
        If frmPatientAppointment.WindowState = 1 Then
            frmPatientAppointment.WindowState = 0
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Apply:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Apply", App.EXEName & "." & Me.Name
    Resume Next
    Exit Sub
End Sub

Private Sub Timer1_Timer()
Dim strList As String
Dim Arr

        If mbooByPass Then mbooByPass = False: Exit Sub
        mbooSearch = True
        mstrAuthString = Replace(txtComments, mstrAuthString, vbNullString)
        txtComments = Replace(txtComments, vbNewLine, vbNullString)
        txtComments = Trim(txtComments)
        If Len(cctSearch.Text) > 0 Then
            strList = Trim(cctSearch.Text)
        End If
    Read_Patient_Status Val(udtSchedPatient.IntId)
    If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
        If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
            cctFacility.Text = Trim(clsFacStatus.StatusDesc)
        End If
        cctFacility.Visible = True
        lblFac.Visible = True
         cctFacility.BackColor = grdCurItem
     Else
         cctFacility.Text = vbNullString
         cctFacility.Visible = False
         lblFac.Visible = False
    End If
    lstListBox.Clear
    mbooByPass = True
    If Trim(strList) <> vbNullString And Len(strList) >= 2 Then
        Select Case IsProperCase(udtAppointment.SchedType)
            Case "Patient"
                Load_Patient strList
            Case "Physician"
                Read_ReferralData "LAST NAME"
            Case "Vendor"
                Read_VendorData "LAST NAME"
            Case "Legal"
                Read_LegalData "LAST NAME"
        End Select
    End If
    mbooSearch = False
    mbooByPass = False
    Timer1.Enabled = False
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub txtApptConfirm_Click()
    Exit Sub
End Sub


Private Sub txtApptConfirm_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i%

    If Not IsNull(txtApptConfirm.Value) Then
        For i = 0 To cboStatus.listcount - 1
            If cboStatus.List(i) = "Confirmed" Then
                cboStatus.Text = cboStatus.List(i)
            End If
        Next i
    Else
        cboStatus.ListIndex = -1
    End If
End Sub


Private Function SaveData() As Boolean
Dim cnApptGrid As New ADODB.Connection
Dim rsApptGrid As New ADODB.Recordset
Dim strSQL As String
Dim i%
Dim lngDateDiff As Long
Dim lngRow As Long
Dim varCellValue
Dim vTime
Dim Arr
Dim iAnswer%

    On Error GoTo Error_SaveData
    Screen.MousePointer = vbHourglass
    Clear_clsWriteAppt
    Clear_Time
    If PsUserID = vbNullString Then PsUserID = Get_PsUserID(psConnect(0))
    clsWriteAppt.Duration = Trim(cboDuration.Text)
    udtSchedPatient.IntId = Trim(lblID2.Caption)
    If Trim(udtSchedPatient.IntId) <> vbNullString Then
        Select Case IsProperCase(udtAppointment.SchedType)
            Case "Patient"
                If Trim(udtSchedPatient.IntId) <> vbNullString Then
                    Get_SchedPatient Val(udtSchedPatient.IntId)
                End If
            Case "Vendor", "Business"
                Get_SchedVendor Val(udtSchedPatient.IntId), udtAppointment.SchedTime
            Case "Legal", "Attorney"
                Get_Schedule_Attorney Val(udtSchedPatient.IntId)
            Case "Physician"
                Get_SchedPhysician Val(udtSchedPatient.IntId)
        End Select
        If Trim(cboFacility.Text) = vbNullString Then
            MsgBox "Please reselect a facility!", vbInformation, "PatientTrac Appointment Scheduling"
            SaveData = False
            Screen.MousePointer = vbDefault
            Exit Function
        ElseIf cboFacility.ListIndex >= 0 Then
            clsWriteAppt.FacilityID = cboFacility.ItemData(cboFacility.ListIndex)
        End If
        If udtAppointment.ProviderID > 0 Then
            clsWriteAppt.ProviderID = udtAppointment.ProviderID
        ElseIf Trim(cboProvider.Text) <> vbNullString Then
            clsWriteAppt.ProviderID = cboProvider.ItemData(cboProvider.ListIndex)
        End If
        For i = 0 To lstReferral.listcount - 1
            If lstReferral.Selected(i) Then
                clsWriteAppt.Referral = lstReferral.ItemData(i)
                Exit For
            End If
        Next i
        clsWriteAppt.Referral_Reas = Replace(cctReferalReas.Text, "'", "''")
        Select Case IsProperCase(cboStatus)
            Case "Cancelled", "Late Cancellation", "No Show"
                Cancel_Patient
                SaveData = True
                Exit Function
            Case "Rescheduled"
                Rescheduled_Patient
                udtAppointment.Appt_ID = 0
                iAnswer = MsgBox("Would you like to set another time now?", vbYesNo, "PatientTrac Scheduling")
                If iAnswer = vbYes Then
                    Format_Appt_Lookup
                    ccSchedDate.Value = txtDate.Value
                    Get_OpenAppointment
                    fraGetAppts.Visible = True
                    mbooByPass = True
                    cboReason.ListIndex = -1
                    mbooByPass = True
                    cboStatus.ListIndex = -1
                    mbooByPass = False
                    SaveData = False
                    Exit Function
                Else
                    MsgBox "This appointment has been updated as inactive!", vbInformation, "PatientTrac Appointment Scheduling"
                    SaveData = True
                    Exit Function
                End If
        End Select
        clsWriteAppt.SchedDate = Format(txtDate.Value, "mm/dd/yyyy")
        clsWriteAppt.SchedTime = Format(ccdTime.Value, "hh:mm AM/PM")
        If Trim(clsWriteAppt.SchedTime) = vbNullString Then
            MsgBox "Unable to save an appointment without a time value!", vbInformation, "PatientTrac Patient Scheduling"
            SaveData = False
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        clsWriteAppt.Comments = Trim(txtComments.Text)
        clsWriteAppt.Comments = Replace(clsWriteAppt.Comments, "'", "''")
        If Trim(cboPatientType.Text) <> vbNullString Then clsWriteAppt.PatientType = Trim(cboPatientType.Text)
        clsWriteAppt.Confirm = Format(txtApptConfirm.Value, "mm/dd/yyyy")
        With grdProcedure
            For lngRow = 1 To .MaxRows
                .Col = 1
                .Row = lngRow
                If .BackColor = grdCurItem Then
                    .GetText 1, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        clsWriteAppt.PurposeVisit = Trim(varCellValue)
                    End If
                    If IsProperCase(udtAppointment.SchedType) = "Patient" Or udtAppointment.SchedType = vbNullString Then
                        .GetText 2, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            clsWriteAppt.CPT_Code = Trim(varCellValue)
                        End If
                        .GetText 3, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            clsWriteAppt.Add_Proc_DESC = Trim(varCellValue)
                        End If
                        .GetText 4, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            clsWriteAppt.Add_Proc_CPT = Trim(varCellValue)
                        End If
                        Exit For
                    Else
                        .GetText 3, lngRow, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                            clsWriteAppt.Duration = Trim(varCellValue)
                        End If
                    End If
                    .GetText 5, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        udtSchedPatient.Current_Location = Trim(varCellValue)
                    End If
                    Exit For
                End If
            Next lngRow
        End With
        If cboDuration <> clsWriteAppt.Duration Then clsWriteAppt.Duration = Trim(cboDuration)
        If clsWriteAppt.PurposeVisit <> vbNullString And clsWriteAppt.Duration = vbNullString Then
            MsgBox "Please select a duration for the visit!", vbInformation, "PatientTrac Appointment Scheduling"
            SaveData = False
            Screen.MousePointer = vbDefault
            Exit Function
        ElseIf clsWriteAppt.PurposeVisit = vbNullString And clsWriteAppt.Duration = vbNullString Then
            MsgBox "Please select a duration and a purpose for the visit of the appointment first!", vbInformation, "PatientTrac Appointment Scheduling"
            Screen.MousePointer = vbDefault
            Exit Function
        ElseIf clsWriteAppt.PurposeVisit = vbNullString And clsWriteAppt.Duration <> vbNullString Then
            MsgBox "Please first select a Purpose for the visit for the appointment.", vbInformation, "PatientTrac Appointment Scheduling"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        If Check_OverBook_Change Then
            SaveData = False
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        Write_Appointment
        SaveData = True
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_SaveData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "SaveData", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Public Sub Get_Info(vlngProviderID As Long, Optional vlngPatientID As Long)
Dim Arr
Dim vTime
Dim vName
Dim lngFacilityID As Long
Dim strDate As String
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim strAffirmed As String
Dim strRepeat As String
Dim i%
Dim lngPatientID As Long
Dim strTermDate As String
Dim strApptDate As String
Dim lngDateDiff As Long
Dim lngProviderID As Long
Dim strTime As String
Dim strNewApptDate As String
Dim booMisMatch As Boolean
Dim lngRow As Long
Dim varCellValue

    On Error GoTo Error_Get_Info
    Screen.MousePointer = vbHourglass
    lngProviderID = vlngProviderID
    lngPatientID = vlngPatientID
    clsReferral.ReferralID = 0
    If cboDuration.listcount <= 0 Then Load_Duration
    strDate = udtAppointment.SchedDate
    ccdTime.Value = Format(Now, "hh:mm AM/PM")
    If mbLookupPatient = True Then
        If udtAppointment.SchedDate <> vbNullString And udtAppointment.SchedTime <> vbNullString Then
            txtDate.Value = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
            ccdTime.Value = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
            mstrOrigApptTime = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
        End If
    Else
        vTime = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
        txtDate.Value = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
        If vTime = vbNullString And mbooNoMessage = False Then
            MsgBox "Please select time!", vbInformation
            Screen.MousePointer = vbDefault
            Unload Me
        ElseIf vTime <> vbNullString Then
            vTime = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
            txtDate.Value = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
        End If
    End If
    mUpdate = False
    udtAppointment.Add_Proc_CPT = vbNullString
    udtAppointment.Add_Proc_DESC = vbNullString
    strNewApptDate = udtAppointment.SchedDate
    strTime = udtAppointment.SchedTime
    Clear_SchedPatient
    Get_SchedPatient lngPatientID
    If udtAppointment.SchedTime <> strTime Then
        Clear_SchedPatient
        udtAppointment.SchedTime = strTime
        udtAppointment.SchedDate = strNewApptDate
    End If
    If cnSched.State = adStateClosed Then cnSched.open psConnect(0)
    udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    strSQL = "Get_PT_Appointment " & lngPatientID & ", " & lngProviderID & ", " & udtAppointment.FacilityID & ", '" & udtAppointment.SchedDate & "','" & Format(udtAppointment.SchedTime, "hh:mm AM/PM") & "'"
        udtAppointment.VisitType = vbNullString
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            If Not .EOF Then
                udtAppointment.Add_Proc_CPT = vbNullString
                udtAppointment.Add_Proc_DESC = vbNullString
                If Not IsNull((!Appt_Time)) Then
                    If (!Appt_Time) <> vbNullString Then
                        If Format((!Appt_Time), "hh:mm AM/PM") = udtAppointment.SchedTime Then
                             Set Flds = .Fields
                             For Each Fld In Flds
                                 If Not IsNull(Fld.Value) Then
                                     If Trim(Fld.Value) <> vbNullString Then
                                        Select Case UCase(Fld.Name)
                                            Case "STATUS"
                                                udtSchedPatient.Status = Trim(Fld.Value)
                                            Case "DURATION"
                                                udtSchedPatient.Duration = Trim(Fld.Value)
                                            Case "APPT_TIME"
                                                udtAppointment.SchedTime = Format(Fld.Value, "hh:mm AM/PM")
                                            Case "PURPOSE_VISIT"
                                               udtAppointment.VisitType = Trim(Fld.Value)
                                            Case "NOTES"
                                                udtSchedPatient.ApptComment = Trim(Fld.Value)
                                            Case "PATIENT_ID"
                                                udtSchedPatient.IntId = Trim(Fld.Value)
                                            Case "APPT_CONFIRMED"
                                                strAffirmed = Trim(Fld.Value)
                                            Case "REASON"
                                                udtSchedPatient.Reason = Trim(Fld.Value)
                                            Case "REPEAT_APPOINT"
                                                strRepeat = Trim(Fld.Value)
                                            Case "TERMINATION_DATE"
                                                strTermDate = Trim(Fld.Value)
                                            Case "PATIENT_TYPE"
                                                udtSchedPatient.PatientType = Trim(Fld.Value)
                                            Case "APPT_TYPE"
                                                udtAppointment.SchedType = Trim(Fld.Value)
                                            Case "CURRENT_LOCATION"
                                              udtSchedPatient.Current_Location = Trim(Fld.Value)
                                            Case "ADD_PROC_CODE"
                                              udtAppointment.Add_Proc_CPT = Trim(Fld.Value)
                                            Case "ADD_PROC_DESC"
                                              udtAppointment.Add_Proc_DESC = Trim(Fld.Value)
                                        End Select
                                    End If
                                End If
                            Next
                            udtSchedPatient.IntId = lngPatientID
                             For i = 0 To cboStatus.listcount - 1
                                If cboStatus.List(i) = Trim(udtSchedPatient.Status) Then
                                    mbooByPass = True
                                    cboStatus.Text = cboStatus.List(i)
                                    Exit For
                                End If
                            Next i
                            For i = 0 To cboDuration.listcount - 1
                                If IsProperCase(cboDuration.List(i)) = IsProperCase(udtSchedPatient.Duration) Then
                                    cboDuration.Text = cboDuration.List(i)
                                    Exit For
                                End If
                            Next i
                             mbooByPass = False
                            If udtAppointment.SchedType <> vbNullString And IsProperCase(udtAppointment.SchedType) <> "Patient" Then
                                If cboPatientType.listcount < 0 Then Load_Visit
                            Else
                                If cboPatientType.listcount <= 0 Then Load_Visit
                                For i = 0 To cboPatientType.listcount - 1
                                    If cboPatientType.List(i) = IsProperCase(udtSchedPatient.PatientType) Then
                                        mbooByPass = True
                                        cboPatientType.Text = cboPatientType.List(i)
                                        Load_Visit cboPatientType.Text
                                        mbooByPass = False
                                        Exit For
                                    End If
                                Next i
                            End If
                            For i = 0 To cboReason.listcount - 1
                                If cboReason.List(i) = udtSchedPatient.Reason Then
                                    cboReason.Text = cboReason.List(i)
                                    Exit For
                                End If
                            Next i
                            If strAffirmed <> vbNullString Then txtApptConfirm.Value = Format(strAffirmed, "mm/dd/yyyy")
                           If grdProcedure.MaxRows = 0 Then
                                Load_Visit udtSchedPatient.PatientType
                           End If
                            For lngRow = 1 To grdProcedure.MaxRows
                                With grdProcedure
                                    .Col = 1
                                    .Row = lngRow
                                    .GetText 1, lngRow, varCellValue
                                    If Trim(varCellValue) <> vbNullString Then
                                        If UCase(varCellValue) = UCase(udtAppointment.VisitType) Then
                                            If udtAppointment.SchedType = "Patient" Or udtAppointment.SchedType = vbNullString Then
                                                Format_CellTypeStatic grdProcedure, lngRow, 1, "1", udtAppointment.VisitType, "TypeHAlignCenter", "VISIT TYPE"
                                                grdProcedure.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                                Format_CellTypeStatic grdProcedure, lngRow, 3, "1", udtAppointment.Add_Proc_DESC, "TypeHAlignCenter", "ADD_DESC"
                                                grdProcedure.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                                Format_CellTypeStatic grdProcedure, lngRow, 4, "1", udtAppointment.Add_Proc_CPT, "TypeHAlignCenter", "ADD_CPT_CODE"
                                                grdProcedure.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                            Else
                                                Format_CellTypeStatic grdProcedure, lngRow, 1, "2", udtAppointment.VisitType, "TypeHAlignCenter", "VISIT TYPE"
                                                grdProcedure.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                                Format_CellTypeStatic grdProcedure, lngRow, 3, "2", udtSchedPatient.Duration, "TypeHAlignCenter", "DURATION"
                                                grdProcedure.SetCellBorder 3, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                                                Format_CellTypeStatic grdProcedure, lngRow, 4, "1", udtAppointment.Add_Proc_CPT, "TypeHAlignCenter", "ADD_CPT_CODE"
'                                                grdProcedure.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                            End If
                                            .BlockMode = True
                                                .Col = 1
                                                .Col2 = .MaxCols
                                                .Row = lngRow
                                                .Row2 = lngRow
                                                .BackColor = grdCurItem
                                            .BlockMode = False
                                            If udtAppointment.VisitType <> vbNullString Then
                                                Update_Facility_Class udtAppointment.FacilityID, lngRow
                                            End If
                                            grdProcedure.TopRow = lngRow
                                            lblCPT.Visible = True
                                            lblCode.Visible = True
                                            grdProcedure.GetText 2, lngRow, varCellValue
                                            lblCode.Caption = Trim(varCellValue)
                                            .TopRow = lngRow
                                            Exit For
                                        End If
                                    End If
                                End With
                            Next lngRow
                             If udtAppointment.SchedTime <> vbNullString And udtAppointment.SchedDate <> vbNullString Then
                                 lblAppt2.Caption = Format(udtAppointment.SchedDate, "MMMM DD, YYYY") & " " & Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                             End If
                             If Trim(udtSchedPatient.ApptComment) <> vbNullString Then txtComments.Text = Trim(udtSchedPatient.ApptComment)
                             mUpdate = True
                             Select Case IsProperCase(udtAppointment.SchedType)
                                Case "Patient"
                                    Load_Patient udtSchedPatient.IntId, "SPECIFIC PATIENT"
                                Case "Physician"
                                    Read_ReferralData "REFERRAL ID", Val(udtSchedPatient.IntId)
                                Case "Vendor"
                                    Read_VendorData "VENDOR ID", Val(udtSchedPatient.IntId)
                                Case "Legal"
                                    Read_LegalData "LEGAL ID", Val(udtSchedPatient.IntId)
                             End Select
                             If udtSchedPatient.IntId <> vbNullString Then
                                 For i = 0 To lstListBox.listcount - 1
                                     If lstListBox.ItemData(i) = Val(udtSchedPatient.IntId) Then
                                         mbooByPass = True
                                         lstListBox.Selected(i) = True
                                         Exit For
                                     End If
                                 Next i
                             End If
                             fraInfo.Visible = True
                             cmdEditPatient.Visible = True
                             If IsProperCase(udtAppointment.SchedType) = "Patient" Then Get_Referrals
                             cctReferalReas.Text = Trim(clsReferral.Consult_Reason)
                        End If
                    End If
                End If
            ElseIf udtAppointment.Appt_ID > 0 Then
                booMisMatch = True
            End If
    End With
    If booMisMatch Then
        Get_ApptRange lngPatientID, lngProviderID
    End If
    Get_RepeatData udtAppointment.FacilityID, lngProviderID
   If cnSched.State = adStateOpen Then
        cnSched.Close
        Set cnSched = Nothing
        Set rsSched = Nothing
    End If
    mBeenHere = True
    If cboDuration.Text = vbNullString Then
        For i = 0 To cboDuration.listcount - 1
            If clsApptPref.SchedInterval = "EVERY HALF HOUR" Then
                If cboDuration.List(i) = "30 minutes" Then cboDuration.Text = cboDuration.List(i)
            ElseIf clsApptPref.SchedInterval = "EVERY 15 MIN" Then
                If cboDuration.List(i) = "15 minutes" Then cboDuration.Text = cboDuration.List(i)
            End If
        Next i
    End If
     mbooByPass = False
   Clear_Surg_Facility
   Read_Patient_Status Val(udtSchedPatient.IntId)
   If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
       If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
           cctFacility.Text = Trim(clsFacStatus.StatusDesc)
       End If
       cctFacility.Visible = True
       lblFac.Visible = True
        cctFacility.BackColor = grdCurItem
   End If
    udtSchedPatient.IntId = lngPatientID
    mbExistSelection = False
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Get_Info:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Info", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Public Function Load_Duration()
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset

    On Error GoTo Error_Load_Duration
    Screen.MousePointer = vbHourglass
    cboDuration.Clear
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(1)
    strSQL = "Get_Appt_Types_Time"
    Set rsAppt = cnAppt.Execute(strSQL)
    Do Until rsAppt.EOF
        If rsAppt.Fields("APPT_TIME") <> vbNullString Then
            If Not TextFound("minutes", rsAppt.Fields("APPT_TIME"), Len("minutes"), True) Then
                cboDuration.AddItem rsAppt.Fields("APPT_TIME") & " " & "minutes"
            Else
                cboDuration.AddItem rsAppt.Fields("APPT_TIME")
            End If
         End If
        rsAppt.MoveNext
    Loop
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Duration:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_Duration", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Sub cmdReason_Click()

    On Error GoTo Error_CmdReason_Click
    Screen.MousePointer = vbHourglass
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Reasons"
    frmNurseTemplates.fType = "Reasons"
    frmNurseTemplates.fSubDomain = "sd_Appointment_Reason"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    cboReason.Clear
    GetTerms cboReason, "sd_Appointment_Reason"
    Screen.MousePointer = vbDefault
    Exit Sub
Error_CmdReason_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdReason_Click", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub



Private Sub cmdStatus_Click()


    On Error GoTo Error_cmdStatus_Click
    Screen.MousePointer = vbHourglass
    
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Status"
    frmNurseTemplates.fType = "Status"
    frmNurseTemplates.fSubDomain = "sd_Appointment_Status"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    cboStatus.Clear
    GetTerms cboStatus, "sd_Appointment_Status"
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdStatus_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdStatus_Click", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Sub cmdTakePic_Click()

   Exit Sub
End Sub

Private Sub Load_Facilities()
Dim strSQL As String
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngFacID As Long
Dim Idx%
Dim booDefault As Boolean
Dim strName As String
Dim booSkip As Boolean
Dim lngDefID As Long


    On Error GoTo Error_Load_Facilities
    Screen.MousePointer = vbHourglass
    cboFacility.Clear
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
'    strSQL = "SELECT APPOINTMENT_PROVIDERS.DEFAULT_FACILITY, " _
'        & "FACILITIES.FACILITY_NAME, FACILITIES.FACILITY_ID " _
'        & "FROM APPOINTMENT_PROVIDERS INNER JOIN FACILITIES ON  " _
'        & "APPOINTMENT_PROVIDERS.FACILITY_ID = FACILITIES.FACILITY_ID " _
'        & "WHERE FACILITIES.ACTIVE  = 1 ORDER BY FACILITIES.FACILITY_NAME"
    strSQL = "Get_Appt_Facilities"
    Set rsFac = cnFac.Execute(strSQL)
    With rsFac
        Do Until .EOF
            lngFacID = 0
            booSkip = False
            booDefault = False
            strName = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "FACILITY_NAME"
                                strName = Fld.Value
                            Case "FACILITY_ID"
                                lngFacID = Fld.Value
                            Case "DEFAULT_FACILITY"
                                booDefault = True
                        End Select
                    End If
                End If
            Next
            .MoveNext
            strName = Fix_Name(strName)
            strName = Fix_Facility_Name(strName)
            If Trim(strName) <> vbNullString Then
                For Idx = 0 To cboFacility.listcount - 1
                    If TextFound(strName, cboFacility.List(Idx), Len(strName), True) Then
                        booSkip = True
                        Exit For
                    End If
                Next Idx
                If Not booSkip Then
                    cboFacility.AddItem strName
                    If lngFacID > 0 Then cboFacility.ItemData(cboFacility.NewIndex) = lngFacID
                End If
            End If
        Loop
    End With
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set cnFac = Nothing
        Set rsFac = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Facilities:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_Facilities", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Sub Form_Activate()
    mbooByPass = True
    On Error GoTo Error_Activate
    Clear_FacStatus
    mbooAuthSequence = False
    cmdApply.Enabled = True
    If clsFacStatus.StatusDesc = vbNullString Then
        cctFacility.Visible = False
        lblFac.Visible = False
    End If
    mbooByPass = False
    centerForm Me
    Exit Sub
Error_Activate:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Activate", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Sub Form_Load()
Dim Idx%

    On Error GoTo Error_Load
    mbooByPass = True
    mbooAuthSequence = False
    Clear_FacStatus
    mstrAuthString = vbNullString
    txtComments = vbNullString
    Screen.MousePointer = vbHourglass
    cctFacility.Text = vbNullString
    cctReferral.Text = vbNullString
    cctReferalReas.Text = vbNullString
    cctReferalReas.Text = vbNullString
    cctReasReferral.Text = vbNullString
    cboProvider.Locked = False
    cboProvider.Enabled = True
    cboReferBy.Clear
    lstReferral.Visible = False
    lstReferral.Clear
    cctFacility.Text = vbNullString
    Format_PurposeGrid
    Clear_RepeatAppt
    mFirst = False
    cctFacility.Visible = False
    lblFac.Visible = False
    Load_Facilities
    cboApptType.Enabled = True
    lblApptType.Enabled = True
    Read_Data
    If udtAppointment.SchedType = vbNullString Then
       cboApptType.ListIndex = 0
    ElseIf IsProperCase(udtAppointment.SchedType) = "Patient" Then
       cboApptType.ListIndex = 0
    ElseIf IsProperCase(udtAppointment.SchedType) = "Physician" Then
       cboApptType.ListIndex = 1
    ElseIf IsProperCase(udtAppointment.SchedType) = "Vendor" Then
       cboApptType.ListIndex = 2
    ElseIf IsProperCase(udtAppointment.SchedType) = "Legal" Then
       cboApptType.ListIndex = 3
    End If
    lblCode.Caption = vbNullString
    lblCode.Visible = False
    udtAppointment.Status = vbNullString
    If mbLookupPatient And Val(udtSchedPatient.IntId) <> 0 Then
       Set_PatientFrame Val(udtSchedPatient.IntId)
       If IsProperCase(udtAppointment.SchedType) = "Patient" Then
          lblID2.Caption = udtSchedPatient.IntId
          lblName2.Caption = udtSchedPatient.FirstName & " " & Left(udtSchedPatient.MiddleInitial, 1) _
                           & " " & udtSchedPatient.LastName
          If udtSchedPatient.Suffix <> vbNullString Then
             lblName2.Caption = lblName2.Caption & ", " & udtSchedPatient.Suffix
          End If
          If udtSchedPatient.Title <> vbNullString Then
             lblName2.Caption = lblName2.Caption & ", " & UCase(udtSchedPatient.Title)
          End If
       End If
       If Trim(cboReferBy.Text) = vbNullString Then GetTerms cboReferBy, "sd_Patient_Referral"
       If IsNull(txtDate.Value) Then
          txtDate.Value = udtAppointment.SchedDate
       End If
       If udtAppointment.SchedTime <> vbNullString Then
          ccdTime.Value = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
       End If
    Else
       cctFacility.Visible = False
       lblFac.Visible = False
       If IsNull(txtDate.Value) Then
          txtDate.Value = udtAppointment.SchedDate
       End If
       If udtAppointment.SchedTime <> vbNullString Then
          ccdTime.Value = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
       End If
    End If
    udtRepeatAppt.Weekday = Format(udtAppointment.SchedDate, "dddd")
    mbooNoMessage = False
    If udtSchedPatient.IntId = vbNullString Then cmdApply.Enabled = False
    If cboDuration.listcount <= 0 Then Load_Duration
    If cboPatientType.listcount <= 0 Then
       cboPatientType.AddItem ("New Patient")
       cboPatientType.AddItem ("Initial Patient")
       cboPatientType.AddItem ("Established Patient")
    End If
    If clsReferral.ReferralID > 0 Then
        For Idx = 0 To cboReferBy.listcount - 1
            If cboReferBy.ItemData(Idx) = clsReferral.ReferralID Then
                cboReferBy.Text = cboReferBy.List(Idx)
                Exit For
            End If
        Next Idx
    End If
    udtAppointment.Add_Proc_CPT = vbNullString
    udtAppointment.Add_Proc_DESC = vbNullString
    cmdApply.Enabled = True
    If clsFacStatus.StatusDesc <> vbNullString Then
        cctFacility.Visible = True
        cctFacility.ZOrder 0
    End If
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
    On Error Resume Next
    mstrTerms = vbNullString
    mbLookupPatient = False
    udtAppointment.Add_Proc_CPT = vbNullString
    udtAppointment.Add_Proc_DESC = vbNullString
    Unload Me
    Exit Sub
End Sub

Private Function Validate_Appointment()
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim vstrTime As String
Dim lngProviderID As Long
Dim sValue As String
Dim intCount As Integer
Dim iAnswer As Integer
Dim strTime As String
Dim strApptDate As String
Dim lngFacility As Long
Dim i%

    On Error GoTo Error_Validate_Appointment
    Screen.MousePointer = vbHourglass
    If udtSchedPatient.IntId = 0 Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    For i = 0 To cboProvider.listcount - 1
        If cboProvider.List(i) = cboProvider.Text Then
            lngProviderID = cboProvider.ItemData(i)
            Exit For
        End If
    Next i
    For i = 0 To cboFacility.listcount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            lngFacility = cboFacility.ItemData(i)
            Exit For
        End If
    Next i
    strTime = ccdTime.Value
    sValue = DateAdd("n", Val(Replace(Trim(cboDuration.Text), "minutes", "")), strTime)
    strApptDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    If mbooCancel Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    strSQL = "SELECT * FROM APPOINTMENTS WHERE PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
        & "FACILITY_ID = " & lngFacility & " AND  " _
        & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & Format(strApptDate, "mm/dd/yyyy") & "' AND " _
        & "CONVERT(DATETIME, APPT_TIME, 108)  BETWEEN '" & Format(strTime, "hh:mm AM/PM") & "' AND '" & Format(sValue, "hh:mm AM/PM") & "' AND ACTIVE = 1 "
    Set rsAppt = cnAppt.Execute(strSQL)
    intCount = 0
        With rsAppt
            Do Until .EOF
                If Not IsNull(!Patient_ID) Then intCount = intCount + 1
                .MoveNext
            Loop
        End With
    Read_OverBook
    If Not mbLookupPatient Then
        If intCount > Val(clsApptPref.OverbookNumber) Then
            MsgBox "Appointment exceeds practice policy for overbooking and cannot be scheduled. Current appointments during this time are " & " " & intCount & ".", vbInformation, "PatientTrac Scheduling"
            mstrValidAppoint = "Invalid Appointment"
            MsgBox "Please reschedule another time!", vbInformation
        End If
    End If
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
     Exit Function
Error_Validate_Appointment:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_Appointment", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Check_TimeBlocks()
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim varCellValue
Dim vDate
Dim iAnswer As Integer
Dim anchorcol  As Variant
Dim anchorrow As Variant
Dim colspan As Variant
Dim rowspan As Variant
Dim IsSpan As Boolean
Dim Span As Integer
Dim strTime As String
Dim strData As String
Dim nDays As Long
Dim Arr
Dim lngDateDiff As Long
Dim booExit As Boolean
Dim strRange As String


    On Error GoTo Error_Check_TimeBlocks
    Screen.MousePointer = vbHourglass
    mstrValidAppoint = vbNullString
    If Not IsNull(txtDate.Value) Then strData = Format(txtDate.Value, "mm/dd/yyyy")
    If Not IsNull(ccdTime.Value) Then strTime = ccdTime.Value
    strRange = DateAdd("d", -7, udtAppointment.SchedDate)
    If Trim(strTime) = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
        strSQL = "Get_SchedConflict " & udtAppointment.ProviderID & ", " & udtFacility.FacID & ", '" & Format(strRange, "mm/dd/yyyy") & "'"
        Set rsTerms = cnTerms.Execute(strSQL)
        Do Until rsTerms.EOF
            Clear_Block
            Set Flds = rsTerms.Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "PROVIDER_ID"
                                    udtBlock.ProviderID = Trim(Fld.Value)
                                Case "START_DATE"
                                    udtBlock.StartDate = Trim(Fld.Value)
                                Case "END_DATE"
                                    udtBlock.EndDate = Trim(Fld.Value)
                                Case "START_TIME"
                                    udtBlock.StartTime = Trim(Fld.Value)
                                Case "END_TIME"
                                    udtBlock.EndTime = Trim(Fld.Value)
                                Case "REASON"
                                    udtBlock.Reason = Trim(Fld.Value)
                                Case "DURATION"
                                    udtBlock.Duration = Trim(Fld.Value)
                                Case "FACILITY_ID"
                                    udtBlock.FacilityID = Trim(Fld.Value)
                                Case "DAYWEEK"
                                    udtBlock.DayWeek = Trim(Fld.Value)
                            End Select
                        End If
                    End If
                Next
                rsTerms.MoveNext
                Select Case IsProperCase(Trim(udtBlock.Duration))
                    Case "1 Week"
                        nDays = 7
                    Case "2 Weeks", "2 Week"
                        nDays = 14
                    Case "3 Weeks"
                        nDays = 21
                    Case "1 Month", "1 Months"
                        nDays = 30
                    Case "2 Months"
                        nDays = 60
                    Case "3 Months"
                        nDays = 90
                   Case "6 Months"
                        nDays = 180
                   Case "1 Year"
                        nDays = 365
                   Case "1 Day"
                        nDays = 1
                End Select
                If udtBlock.Duration = "Permanent" Then
                    If DateDiff("n", udtAppointment.SchedTime, udtBlock.StartTime) > 0 Then
                        If DateDiff("n", udtAppointment.SchedTime, udtBlock.EndTime) < 0 Then
                            mstrValidAppoint = False
                            booExit = True
                            Exit Do
                        Else
                            mstrValidAppoint = True
                            Exit Do
                        End If
                    Else
                        mstrValidAppoint = True
                        Exit Do
                    End If
                    mstrValidAppoint = False
                ElseIf nDays > 1 Then
                    If DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate) > 0 And _
                        DateDiff("d", udtBlock.EndDate, udtAppointment.SchedDate) < 0 Then
                        mstrValidAppoint = True
                        If udtBlock.EndDate <> vbNullString Then
                            lngDateDiff = DateDiff("d", Format(udtBlock.StartDate, "mm/dd/yyyy"), Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
                            If lngDateDiff >= 0 Then
                                lngDateDiff = DateDiff("n", udtBlock.StartTime, strTime)
                                If lngDateDiff >= 0 Then
                                    lngDateDiff = DateDiff("n", strTime, udtBlock.EndTime)
                                    If lngDateDiff >= 0 Then
                                        lngDateDiff = DateDiff("d", Format(udtBlock.EndDate, "mm/dd/yyyy"), Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
                                        If lngDateDiff <= 0 Then
                                            If udtBlock.DayWeek <> vbNullString Then
                                                 If UCase(udtBlock.DayWeek) = GetDay(udtAppointment.SchedDate) Then
                                                    mstrValidAppoint = "Invalid Appointment"
                                                    Exit Do
                                                 End If
                                            Else
                                                mstrValidAppoint = "Invalid Appointment"
                                                Exit Do
                                            End If
                                        Else
                                            mstrValidAppoint = "Authorized"
                                            Exit Do
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    ElseIf DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate) = 0 And _
                        DateDiff("d", udtBlock.EndDate, udtAppointment.SchedDate) = 0 Then
                        If lngDateDiff = DateDiff("n", udtAppointment.SchedTime, udtBlock.EndTime) < 0 Then
                            mstrValidAppoint = False
                            booExit = True
                            Exit Do
                        Else
                            mstrValidAppoint = True
                            Exit Do
                        End If
                    End If
                End If
            Loop
        End If
    If mstrValidAppoint = "Invalid Appointment" Then
        MsgBox "This Appointment time  for this provider has been blocked for" & " " & IsProperCase(udtBlock.Reason) & ". ", vbInformation, "PatientTrac Appointment Scheduling"
        Screen.MousePointer = vbDefault
    Else
        mstrValidAppoint = "Authorized"
    End If
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Check_TimeBlocks:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Check_TimeBlocks", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_Referrals()
Dim strSQL As String
Dim cnReferral As New ADODB.Connection
Dim rsReferral As New ADODB.Recordset
Dim lngReferralID As Long
Dim strRefReason As String
Dim Idx As Integer

    On Error GoTo Error_Get_Referrals
    If cnReferral.State = adStateClosed Then cnReferral.open psConnect(0)
    Screen.MousePointer = vbHourglass
    strSQL = "Get_Superbill_Referral " & Val(udtSchedPatient.IntId)
    Set rsReferral = cnReferral.Execute(strSQL)
    With rsReferral
        Do Until .EOF
            If Not IsNull(!Referral_ID) Then
                If (!Referral_ID) > 0 Then
                    clsReferral.ReferralID = (!Referral_ID): lngReferralID = clsReferral.ReferralID
                    If Not IsNull(!REFERRAL_REASON) Then clsReferral.Consult_Reason = Trim((!REFERRAL_REASON))
                    Exit Do
                End If
            End If
            .MoveNext
        Loop
    End With
    If lngReferralID > 0 Then
        Load_Specific_Referral clsReferral.ReferralID
        If clsReferral.Physician <> vbNullString Then
            cboReferBy.AddItem clsReferral.Physician
            mbooByPass = True
            cboReferBy.ItemData(cboReferBy.NewIndex) = clsReferral.ReferralID
            For Idx = 0 To cboReferBy.listcount - 1
                If cboReferBy.ItemData(Idx) = clsReferral.ReferralID Then
                    cboReferBy.Text = cboReferBy.List(Idx)
                    Exit For
                End If
            Next Idx
        End If
    End If
    If lngReferralID > 0 Then
        strSQL = "Get_Patient_Account_Referral " & Val(udtSchedPatient.IntId)
        Set rsReferral = cnReferral.Execute(strSQL)
        With rsReferral
            Do Until .EOF
                If Not IsNull(!Referral_ID) Then
                    clsReferral.ReferralID = (!Referral_ID): lngReferralID = clsReferral.ReferralID
                    If Not IsNull(!REFERRAL_REASON) Then clsReferral.Consult_Reason = Trim((!REFERRAL_REASON))
                End If
            .MoveNext
            Loop
            If lngReferralID > 0 Then
                Load_Specific_Referral clsReferral.ReferralID
                cboReferBy.AddItem clsReferral.Physician
                mbooByPass = True
                cboReferBy.ItemData(cboReferBy.NewIndex) = clsReferral.ReferralID
                For Idx = 0 To cboReferBy.listcount - 1
                    If cboReferBy.ItemData(Idx) = clsReferral.ReferralID Then
                        cboReferBy.Text = cboReferBy.List(Idx)
                        Exit For
                    End If
                Next Idx
            End If
        End With
    End If
    If cnReferral.State = adStateOpen Then
        cnReferral.Close
        Set cnReferral = Nothing
    End If
    If lngReferralID = 0 Then Check_ApptReferral
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Referrals:
    Screen.MousePointer = vbDefault
     Write_Error Err.Description, Err.Number, strSQL, "Get_Referrals", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Check_ApptReferral()
Dim cnRefer As New ADODB.Connection
Dim rsRefer As New ADODB.Recordset
Dim Idx%
Dim strSQL As String
Dim strDate As String
Dim lngDateDiff As Long
Dim strApptDate As String
Dim lngRefID As Long
Dim strRefReason As String

    On Error GoTo Error_Check_ApptReferral
    Clear_ReferralStrings
    If Val(udtSchedPatient.IntId) <= 0 Then Exit Function
    Screen.MousePointer = vbHourglass
    If Trim(clsEncounter.StartDate) <> vbNullString Then
        strDate = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    ElseIf piEncounterID > 0 Then
        Get_Encounter piEncounterID
        strDate = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    Else
        strDate = Format(Now, "mm/dd/yyyy")
    End If
    
    If cnRefer.State = adStateClosed Then cnRefer.open psConnect(0)
    If strDate = vbNullString Then strDate = Format(Now, "mm/dd/yyy")
    strSQL = "Get_Appt_Referral " & Val(udtSchedPatient.IntId)
    Set rsRefer = cnRefer.Execute(strSQL)
    With rsRefer
        Do Until .EOF
            If Not IsNull(!Referral_ID) Then lngRefID = (!Referral_ID)
            If (!Appt_Date) <> vbNullString Then
                strApptDate = (!Appt_Date)
                strApptDate = Format(strApptDate, "mm/dd/yyyy")
            End If
            If (!REFERRAL_REASON) <> vbNullString Then
                 clsReferral.Consult_Reason = (!REFERRAL_REASON)
            End If
            lngDateDiff = DateDiff("d", strApptDate, strDate)
            If lngDateDiff <= 0 And lngRefID > 0 Then
                If Not IsNull(!Referral_ID) Then
                    clsReferral.ReferralID = (!Referral_ID)
                    Exit Do
                End If
            ElseIf lngRefID > 0 Then
                clsReferral.ReferralID = lngRefID
            End If
        .MoveNext
        Loop
    End With
    If clsReferral.ReferralID > 0 Then
        strRefReason = clsReferral.Consult_Reason
        Load_Specific_Referral clsReferral.ReferralID
        Get_Referral clsReferral.ReferralID
        clsReferral.Consult_Reason = strRefReason
        mbooByPass = True
        cboReferBy.AddItem clsReferral.Physician
        cboReferBy.ItemData(cboReferBy.NewIndex) = clsReferral.ReferralID
        For Idx = 0 To cboReferBy.listcount - 1
            If cboReferBy.ItemData(Idx) = clsReferral.ReferralID Then
                If Trim(cboReferBy.List(Idx)) <> vbNullString Then
                    cboReferBy.Text = cboReferBy.List(Idx)
                End If
                Exit For
            End If
        Next Idx
    End If
    If cnRefer.State = adStateOpen Then
        cnRefer.Close
        Set cnRefer = Nothing
        Set rsRefer = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Check_ApptReferral:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Check_ApptReferral", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function

Private Function Calc_Duration(vstrReason As String)
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim strTitle As String

    On Error GoTo Error_Calc_Duration
    Screen.MousePointer = vbHourglass
    strTitle = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
    clsGridFormat.List = vbNullString
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(1)
    strSQL = "SELECT APPT_TIME FROM APPOINTMENT_TYPE WHERE " _
        & "APPT_CODE = '" & vstrReason & "' AND " _
        & "PROVIDER_TITLE = '" & strTitle & "' AND APPT_TIME IS NOT NULL"
    Set rsAppt = cnAppt.Execute(strSQL)
    With rsAppt
        If Not .EOF Then
            If (!Appt_Time) <> vbNullString Then udtEncounter.VisitTime = (!Appt_Time)
        End If
    End With
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Calc_Duration:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Calc_Duration", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

