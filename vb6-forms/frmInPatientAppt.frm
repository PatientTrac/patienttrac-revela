VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmInPatientAppt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Appointments"
   ClientHeight    =   11040
   ClientLeft      =   4470
   ClientTop       =   3000
   ClientWidth     =   18555
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInPatientAppt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   18555
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraFacilities 
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
      Height          =   5715
      Left            =   4140
      TabIndex        =   92
      Top             =   330
      Visible         =   0   'False
      Width           =   6165
      Begin VB.Frame fraPrevVisit 
         Height          =   3165
         Left            =   1110
         TabIndex        =   93
         Top             =   1350
         Visible         =   0   'False
         Width           =   4965
         Begin FPSpreadADO.fpSpread grdVisitDate 
            Height          =   2595
            Left            =   0
            TabIndex        =   94
            Top             =   60
            Width           =   4890
            _Version        =   524288
            _ExtentX        =   8625
            _ExtentY        =   4577
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FormulaSync     =   0   'False
            GrayAreaBackColor=   16777215
            GridColor       =   8388608
            MaxCols         =   1
            MaxRows         =   0
            RetainSelBlock  =   0   'False
            ScrollBarShowMax=   0   'False
            SelectBlockOptions=   0
            ShadowColor     =   16761024
            ShadowDark      =   8388608
            ShadowText      =   8388608
            SpreadDesigner  =   "frmInPatientAppt.frx":08CA
            UnitType        =   0
            UserResize      =   0
            Appearance      =   2
            TextTipDelay    =   30
            ClipboardOptions=   0
            AppearanceStyle =   0
         End
         Begin PT_XP_Button.PT_XP cmdSelectDate 
            Height          =   315
            Left            =   1620
            TabIndex        =   95
            Top             =   2730
            Width           =   945
            _ExtentX        =   1667
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
            Picture         =   "frmInPatientAppt.frx":0E26
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Select"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmInPatientAppt.frx":11C0
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdRemove 
            Height          =   315
            Left            =   2670
            TabIndex        =   96
            Top             =   2730
            Width           =   975
            _ExtentX        =   1720
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
            Picture         =   "frmInPatientAppt.frx":1512
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close"
            DefaultThemes   =   2
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmInPatientAppt.frx":1F24
            ShowFocus       =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread grdFacility 
         Height          =   4395
         Left            =   0
         TabIndex        =   97
         Top             =   150
         Width           =   6105
         _Version        =   524288
         _ExtentX        =   10769
         _ExtentY        =   7752
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
         MaxCols         =   1
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
         SpreadDesigner  =   "frmInPatientAppt.frx":2276
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
         CellNoteIndicator=   3
         AppearanceStyle =   0
      End
      Begin PT_XP_Button.PT_XP cmdSelectFac 
         Height          =   405
         Left            =   2610
         TabIndex        =   98
         Top             =   5220
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "frmInPatientAppt.frx":2794
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":31A6
         ShowFocus       =   -1  'True
      End
      Begin MSComCtl2.DTPicker ccdPrevDate 
         Height          =   375
         Left            =   1140
         TabIndex        =   99
         Top             =   4590
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         CustomFormat    =   "MM/dd/yyy hh:mm:ss"
         DateIsNull      =   -1  'True
         Format          =   85065731
         CurrentDate     =   38862.7362384259
         MinDate         =   32874
      End
      Begin PT_XP_Button.PT_XP cmdPrevVisitDate 
         Height          =   585
         Left            =   4140
         TabIndex        =   100
         Top             =   4590
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1032
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
         Picture         =   "frmInPatientAppt.frx":34F8
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Previous Visit Dates"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":3A92
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblAdmt 
         BackStyle       =   0  'Transparent
         Caption         =   "Visit Date"
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
         Left            =   30
         TabIndex        =   101
         Top             =   4620
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.TextBox cctFacility 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5220
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   91
      Top             =   60
      Width           =   4875
   End
   Begin VB.Frame fraPatient 
      Height          =   10965
      Left            =   4080
      TabIndex        =   48
      Top             =   60
      Visible         =   0   'False
      Width           =   14475
      Begin PT_XP_Button.PT_XP cmdUpdate 
         Height          =   345
         Left            =   2700
         TabIndex        =   63
         Top             =   1350
         Width           =   1695
         _ExtentX        =   2990
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
         Picture         =   "frmInPatientAppt.frx":3DE4
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Save Roster"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":437E
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrintRoster 
         Height          =   345
         Left            =   6300
         TabIndex        =   77
         Top             =   1350
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
         Picture         =   "frmInPatientAppt.frx":46D0
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Print"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":4C6A
         ShowFocus       =   -1  'True
      End
      Begin VB.ComboBox cboSearchFilter 
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
         ItemData        =   "frmInPatientAppt.frx":4FBC
         Left            =   1230
         List            =   "frmInPatientAppt.frx":4FBE
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   180
         Width           =   3765
      End
      Begin PT_XP_Button.PT_XP cmdRefresh 
         Height          =   345
         Left            =   4560
         TabIndex        =   79
         Top             =   1350
         Width           =   1635
         _ExtentX        =   2884
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
         Picture         =   "frmInPatientAppt.frx":4FC0
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Refresh"
         Depth           =   5
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":555A
         ShowFocus       =   -1  'True
      End
      Begin VB.Frame fraDeath 
         Height          =   1005
         Left            =   5730
         TabIndex        =   71
         Top             =   3390
         Visible         =   0   'False
         Width           =   2025
         Begin MSComCtl2.DTPicker ccdDeath 
            Height          =   375
            Left            =   240
            TabIndex        =   72
            Top             =   180
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   85065729
            CurrentDate     =   39288
         End
         Begin PT_XP_Button.PT_XP cmdCloseDeath 
            Height          =   285
            Left            =   540
            TabIndex        =   73
            Top             =   600
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
            Picture         =   "frmInPatientAppt.frx":58AC
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmInPatientAppt.frx":62BE
            ShowFocus       =   -1  'True
         End
      End
      Begin VB.Frame fraLive 
         Height          =   1005
         Left            =   3570
         TabIndex        =   58
         Top             =   3390
         Visible         =   0   'False
         Width           =   2025
         Begin MSComCtl2.DTPicker ccdLive 
            Height          =   375
            Left            =   270
            TabIndex        =   59
            Top             =   180
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   85065729
            CurrentDate     =   39288
         End
         Begin PT_XP_Button.PT_XP cmdLivAccept 
            Height          =   285
            Left            =   570
            TabIndex        =   60
            Top             =   600
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
            Picture         =   "frmInPatientAppt.frx":6610
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmInPatientAppt.frx":7022
            ShowFocus       =   -1  'True
         End
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
         ForeColor       =   &H00000000&
         Height          =   1260
         Left            =   8970
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   53
         Top             =   570
         Width           =   4875
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
         Left            =   8970
         TabIndex        =   49
         Top             =   180
         Width           =   4875
      End
      Begin FPSpreadADO.fpSpread grdPatient 
         Height          =   8535
         Left            =   30
         TabIndex        =   50
         Top             =   1860
         Width           =   14385
         _Version        =   524288
         _ExtentX        =   25374
         _ExtentY        =   15055
         _StockProps     =   64
         AutoCalc        =   0   'False
         ColHeaderDisplay=   0
         DAutoCellTypes  =   0   'False
         DAutoFill       =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSizeCols   =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
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
         MaxCols         =   2
         MaxRows         =   0
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         ShadowText      =   4194304
         SpreadDesigner  =   "frmInPatientAppt.frx":7374
         UnitType        =   0
         UserResize      =   0
         Appearance      =   1
         TextTipDelay    =   50
         ClipboardOptions=   0
         CellNoteIndicator=   3
         AppearanceStyle =   0
      End
      Begin PT_XP_Button.PT_XP cmdRemoveRow 
         Height          =   435
         Left            =   15270
         TabIndex        =   54
         Top             =   2130
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
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
         Picture         =   "frmInPatientAppt.frx":78B0
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Remove Patient"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":7E4A
         ShowFocus       =   -1  'True
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   345
         Left            =   3570
         TabIndex        =   55
         Top             =   1500
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
         Format          =   85065731
         CurrentDate     =   38792
      End
      Begin PT_XP_Button.PT_XP cmdCloseList 
         Height          =   345
         Left            =   5850
         TabIndex        =   57
         Top             =   10500
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
         Picture         =   "frmInPatientAppt.frx":819C
         AppearanceThemes=   3
         BorderDrawEvent =   1
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":8BAE
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdStatus 
         Height          =   315
         Left            =   13890
         TabIndex        =   61
         ToolTipText     =   "Manage Status"
         Top             =   180
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
         Picture         =   "frmInPatientAppt.frx":8F00
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmInPatientAppt.frx":949A
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdNewRepeat 
         Height          =   435
         Left            =   360
         TabIndex        =   74
         Top             =   -510
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   767
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
         Picture         =   "frmInPatientAppt.frx":97EC
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Add Patient"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":9D86
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdActive 
         Height          =   345
         Left            =   150
         TabIndex        =   75
         Top             =   1350
         Width           =   2415
         _ExtentX        =   4260
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
         Picture         =   "frmInPatientAppt.frx":A0D8
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Create Visit Roster"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":A672
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCreateVisit2 
         Height          =   345
         Left            =   11640
         TabIndex        =   76
         Top             =   10500
         Width           =   2415
         _ExtentX        =   4260
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
         Picture         =   "frmInPatientAppt.frx":A9C4
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Create Visit Roster"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":AF5E
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdLoad 
         Height          =   345
         Left            =   1230
         TabIndex        =   90
         Top             =   600
         Width           =   3765
         _ExtentX        =   6641
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
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Show Inactive"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":B2B0
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdExcel 
         Height          =   465
         Left            =   7620
         TabIndex        =   102
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
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
         Picture         =   "frmInPatientAppt.frx":B602
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Print - Excel"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":BB9C
         ShowFocus       =   -1  'True
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
         Left            =   2130
         TabIndex        =   56
         Top             =   1290
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search by:"
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
         TabIndex        =   52
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblEnterInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name, First Name"
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
         Left            =   2370
         TabIndex        =   51
         Top             =   780
         Visible         =   0   'False
         Width           =   3465
      End
   End
   Begin VB.Frame fraAddPatient 
      Height          =   4755
      Left            =   7710
      TabIndex        =   82
      Top             =   1170
      Visible         =   0   'False
      Width           =   10575
      Begin FPSpreadADO.fpSpread grdAppt 
         Height          =   3345
         Left            =   90
         TabIndex        =   87
         Top             =   810
         Width           =   10425
         _Version        =   524288
         _ExtentX        =   18389
         _ExtentY        =   5900
         _StockProps     =   64
         AutoClipboard   =   0   'False
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         DAutoSizeCols   =   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   16777215
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   5
         MaxRows         =   0
         ProcessTab      =   -1  'True
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         SpreadDesigner  =   "frmInPatientAppt.frx":BEEE
         UserResize      =   0
         Appearance      =   1
         ClipboardOptions=   0
         CellNoteIndicator=   3
         AppearanceStyle =   0
      End
      Begin VB.TextBox cctAddPatient 
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
         Height          =   375
         Left            =   90
         TabIndex        =   83
         Top             =   330
         Width           =   4875
      End
      Begin PT_XP_Button.PT_XP cmdAddRoster 
         Height          =   345
         Left            =   3360
         TabIndex        =   85
         Top             =   4290
         Width           =   2085
         _ExtentX        =   3678
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
         Picture         =   "frmInPatientAppt.frx":C419
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Add to Visits"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":C9B3
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCloseAdd 
         Height          =   345
         Left            =   5580
         TabIndex        =   86
         Top             =   4290
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
         Picture         =   "frmInPatientAppt.frx":CD05
         AppearanceThemes=   3
         BorderDrawEvent =   1
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatientAppt.frx":D717
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblInstr 
         BackStyle       =   0  'Transparent
         Caption         =   "Search (Last Name, First Name)"
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
         Left            =   120
         TabIndex        =   84
         Top             =   90
         Width           =   3855
      End
   End
   Begin VB.ComboBox cboExtender 
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
      Left            =   5220
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   69
      ToolTipText     =   "Case Manager"
      Top             =   780
      Width           =   4875
   End
   Begin VB.PictureBox picPsych 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   60
      ScaleHeight     =   825
      ScaleWidth      =   3945
      TabIndex        =   66
      Top             =   6030
      Width           =   3945
      Begin VB.Label lblCurScheAppt 
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment Count"
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
         Height          =   630
         Left            =   60
         TabIndex        =   67
         Top             =   30
         Visible         =   0   'False
         Width           =   3660
      End
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
      ItemData        =   "frmInPatientAppt.frx":DA69
      Left            =   11820
      List            =   "frmInPatientAppt.frx":DA6B
      Style           =   2  'Dropdown List
      TabIndex        =   64
      Top             =   90
      Width           =   3435
   End
   Begin VB.Timer schTimer 
      Interval        =   60000
      Left            =   9480
      Top             =   1320
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   3120
      Left            =   30
      TabIndex        =   24
      Top             =   30
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   5503
      _Version        =   393216
      ForeColor       =   8388608
      BackColor       =   16744448
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      StartOfWeek     =   85065730
      TitleBackColor  =   16761024
      TitleForeColor  =   8388608
      TrailingForeColor=   8421504
      CurrentDate     =   38806
   End
   Begin VB.Timer strTelTime 
      Interval        =   60000
      Left            =   4950
      Top             =   8970
   End
   Begin VB.Frame fraSchLegend 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scheduling Color Codes"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1425
      Left            =   30
      TabIndex        =   13
      Top             =   4530
      Width           =   4005
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   43
         Top             =   1050
         Width           =   285
      End
      Begin VB.PictureBox picPreOp 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   35
         ToolTipText     =   "Pre Operative Patient"
         Top             =   1470
         Width           =   285
      End
      Begin VB.PictureBox picPostOp 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   34
         Top             =   1770
         Width           =   285
      End
      Begin VB.PictureBox picExstNewPat 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   33
         ToolTipText     =   "Established Patient New Consultation"
         Top             =   1770
         Width           =   285
      End
      Begin VB.PictureBox picNewPatient 
         BackColor       =   &H00FFC0FF&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   32
         ToolTipText     =   "New Initial Patient"
         Top             =   1470
         Width           =   285
      End
      Begin VB.PictureBox picReturn 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   31
         Top             =   2040
         Width           =   285
      End
      Begin VB.PictureBox picSurgery 
         BackColor       =   &H00FFFF80&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   30
         Top             =   2040
         Width           =   285
      End
      Begin VB.PictureBox picHospital 
         BackColor       =   &H00FFFF80&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   28
         Top             =   780
         Width           =   285
      End
      Begin VB.PictureBox picBlock 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   26
         Top             =   780
         Width           =   285
      End
      Begin VB.PictureBox picPresent 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   17
         Top             =   210
         Width           =   285
      End
      Begin VB.PictureBox picCheckOut 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   16
         Top             =   510
         Width           =   285
      End
      Begin VB.PictureBox picOffice 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   15
         Top             =   510
         Width           =   285
      End
      Begin VB.PictureBox picCancel 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   14
         Top             =   210
         Width           =   285
      End
      Begin VB.Label lblBus 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Non Patient/Business"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   600
         TabIndex        =   44
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label lblNew 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "New Patient"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   540
         TabIndex        =   41
         Top             =   1500
         Width           =   930
      End
      Begin VB.Label lblEstablishNew 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Est New Patient"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   540
         TabIndex        =   40
         Top             =   1800
         Width           =   1230
      End
      Begin VB.Label lblPostOp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Post-Op"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   2595
         TabIndex        =   39
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label lblPreOp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Pre-Op"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   2610
         TabIndex        =   38
         Top             =   1500
         Width           =   555
      End
      Begin VB.Label lblFU 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "FollowUp Exam"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   525
         TabIndex        =   37
         Top             =   2070
         Width           =   1170
      End
      Begin VB.Label lblSurgery 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Surgery"
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
         Left            =   2625
         TabIndex        =   36
         Top             =   2070
         Width           =   780
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   30
         X2              =   3975
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2655
         TabIndex        =   29
         Top             =   810
         Width           =   615
      End
      Begin VB.Label lblBlockTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Block Time"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   585
         TabIndex        =   27
         Top             =   810
         Width           =   810
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel - No show"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2670
         TabIndex        =   21
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblOffice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Office"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2655
         TabIndex        =   20
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Checked out"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   570
         TabIndex        =   19
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Present"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   570
         TabIndex        =   18
         Top             =   240
         Width           =   555
      End
   End
   Begin PT_XP_Button.PT_XP cmdLookUpPatient 
      Height          =   345
      Left            =   2100
      TabIndex        =   12
      Top             =   3300
      Width           =   1935
      _ExtentX        =   3413
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
      Picture         =   "frmInPatientAppt.frx":DA6D
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Patient Roster"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":E007
      ShowFocus       =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16110
      Top             =   12660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin PT_XP_Button.PT_XP cmdCurrentWeek 
      Height          =   435
      Left            =   5370
      TabIndex        =   11
      Top             =   -480
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   767
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
      Picture         =   "frmInPatientAppt.frx":E359
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Current Week"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":E6F3
      ShowFocus       =   -1  'True
   End
   Begin RichTextLib.RichTextBox rtfChanges 
      Height          =   3735
      Left            =   0
      TabIndex        =   9
      Top             =   7290
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6588
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmInPatientAppt.frx":EA45
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PT_XP_Button.PT_XP cmdPrint 
      Height          =   375
      Left            =   13680
      TabIndex        =   7
      Top             =   -390
      Width           =   1485
      _ExtentX        =   2619
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
      Picture         =   "frmInPatientAppt.frx":EAC9
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Print"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":F063
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdNextWeek 
      Height          =   435
      Left            =   6660
      TabIndex        =   5
      Top             =   -480
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
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
      Picture         =   "frmInPatientAppt.frx":F3B5
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Next Week"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":F74F
      PictureAlignment=   5
      ShowFocus       =   -1  'True
   End
   Begin VB.ComboBox cboProviders 
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
      Left            =   5220
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   420
      Width           =   4875
   End
   Begin PT_XP_Button.PT_XP cmdViewProvider 
      Height          =   345
      Left            =   11370
      TabIndex        =   3
      Top             =   -480
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
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
      AppearanceMode  =   2
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "View By"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdData 
      DragIcon        =   "frmInPatientAppt.frx":FAA1
      Height          =   5715
      Left            =   4110
      TabIndex        =   0
      Top             =   1170
      Width           =   14385
      _Version        =   524288
      _ExtentX        =   25374
      _ExtentY        =   10081
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
         Name            =   "Georgia"
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
      MaxCols         =   7
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   0
      ShadowText      =   0
      SpreadDesigner  =   "frmInPatientAppt.frx":FEE3
      UnitType        =   0
      UserResize      =   0
      VisibleCols     =   7
      Appearance      =   2
      TextTip         =   3
      TextTipDelay    =   25
      ScrollBarTrack  =   1
      AppearanceStyle =   0
   End
   Begin PT_XP_Button.PT_XP cmdPrevWeek 
      Height          =   435
      Left            =   4080
      TabIndex        =   6
      Top             =   -480
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   767
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
      Picture         =   "frmInPatientAppt.frx":10472
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Previous Week"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":1080C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdBlockTimes 
      Height          =   495
      Left            =   -1800
      TabIndex        =   22
      Top             =   6450
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
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
      Picture         =   "frmInPatientAppt.frx":10B5E
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Block Time"
      DepthMode       =   1
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":110F8
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRepair 
      Height          =   405
      Left            =   2880
      TabIndex        =   23
      Top             =   7620
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Repair"
      DepthEvent      =   1
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPhone 
      Height          =   345
      Left            =   16650
      TabIndex        =   25
      Top             =   780
      Width           =   1695
      _ExtentX        =   2990
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
      Picture         =   "frmInPatientAppt.frx":1144A
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "New Message"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":119E4
      PictureAlignment=   0
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdMessages 
      Height          =   3765
      Left            =   4110
      TabIndex        =   42
      Top             =   7290
      Width           =   14415
      _Version        =   524288
      _ExtentX        =   25426
      _ExtentY        =   6641
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
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
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
      SpreadDesigner  =   "frmInPatientAppt.frx":11D36
      UnitType        =   0
      UserResize      =   0
      Appearance      =   2
      TextTipDelay    =   60000
      ClipboardOptions=   0
      AppearanceStyle =   0
   End
   Begin PT_XP_Button.PT_XP cmdReport 
      Height          =   345
      Left            =   2100
      TabIndex        =   45
      Top             =   3810
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Picture         =   "frmInPatientAppt.frx":12299
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Reset MD Roster"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":12833
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdStop 
      Height          =   495
      Left            =   -1470
      TabIndex        =   46
      Top             =   6450
      Width           =   1185
      _ExtentX        =   2090
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
      Picture         =   "frmInPatientAppt.frx":12B85
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Stop Refresh"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":1311F
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdNewPat 
      Height          =   345
      Left            =   60
      TabIndex        =   47
      Top             =   3300
      Width           =   1935
      _ExtentX        =   3413
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
      Picture         =   "frmInPatientAppt.frx":13471
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "New Patient"
      Depth           =   5
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":13A0B
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRosterReport 
      Height          =   345
      Left            =   60
      TabIndex        =   68
      Top             =   3810
      Width           =   1935
      _ExtentX        =   3413
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
      Picture         =   "frmInPatientAppt.frx":13D5D
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   " Roster Report"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":142F7
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPrintAppt 
      Height          =   345
      Left            =   15420
      TabIndex        =   78
      Top             =   780
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmInPatientAppt.frx":14649
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Print"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":14BE3
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   345
      Left            =   10200
      TabIndex        =   80
      Top             =   780
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
      Picture         =   "frmInPatientAppt.frx":14F35
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":154CF
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdResched 
      Height          =   345
      Left            =   11370
      TabIndex        =   81
      Top             =   780
      Width           =   2085
      _ExtentX        =   3678
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
      Picture         =   "frmInPatientAppt.frx":15821
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Add Patient Visit"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":15DBB
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPatientStatus 
      Height          =   345
      Left            =   13620
      TabIndex        =   88
      Top             =   780
      Width           =   1635
      _ExtentX        =   2884
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
      AppearanceMode  =   2
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Visit Status"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatientAppt.frx":1610D
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblResource 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Assistant"
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
      Left            =   4185
      TabIndex        =   70
      Top             =   870
      Width           =   930
   End
   Begin VB.Label lblApptType 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Roster Type"
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
      Left            =   10470
      TabIndex        =   65
      Top             =   150
      Width           =   1155
   End
   Begin VB.Label lblProviders 
      Alignment       =   1  'Right Justify
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
      Left            =   4185
      TabIndex        =   62
      Top             =   510
      Width           =   855
   End
   Begin VB.Label lblSchedChange 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scheduling Changes"
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
      Left            =   60
      TabIndex        =   10
      Top             =   7020
      Width           =   1935
   End
   Begin VB.Label lblProvider 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   10110
      TabIndex        =   8
      Top             =   13800
      Width           =   5295
   End
   Begin VB.Label lblSelectDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Select Date"
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
      Left            =   0
      TabIndex        =   2
      Top             =   -3510
      Width           =   4905
   End
   Begin VB.Label lblLocation 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Facility"
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
      Left            =   4185
      TabIndex        =   1
      Top             =   150
      Width           =   765
   End
   Begin VB.Menu View 
      Caption         =   "Views by Provider"
      Visible         =   0   'False
      Begin VB.Menu mnu_Day 
         Caption         =   "By Day"
      End
      Begin VB.Menu mnu_Week 
         Caption         =   "By Week"
      End
      Begin VB.Menu mnu_Month 
         Caption         =   "By Month"
      End
   End
   Begin VB.Menu mnu_Patient_Type 
      Caption         =   "Patient Type"
      Begin VB.Menu mnu_Active 
         Caption         =   "Active"
      End
      Begin VB.Menu mnu_Cancelled 
         Caption         =   "Cancelled"
      End
      Begin VB.Menu mnu_Postpone 
         Caption         =   "Postpone"
      End
      Begin VB.Menu mnu_Workup_Complete 
         Caption         =   "Workup Completed"
      End
      Begin VB.Menu mnu_All_PatientType 
         Caption         =   "All Patient Types"
      End
   End
End
Attribute VB_Name = "frmInPatientAppt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Private mStartHour() As String
Private mEndHour() As String
Private mstrShow As String
Private mbooLoad As Boolean
Private mStartIdx As Integer
Private mEndIdx As Integer
Public vTime As Variant
Public mlngCurProvID As Long
Public strSelDate As String
Private mLoadPatient As Boolean
Private mChange As Boolean
Public mBeenHere As Boolean
Private mlngGridRow As Long
Private mlngGridCol As Long
Private mbooUnknown As Boolean
Private mlngMessRow As Long
Private mlngMessCol As Long
Private mEnterRow As Boolean
Private mstrValidAppoint As String
Private mFirstLoad As Boolean
Public mstrApptMode As String
Public mbooTimeAvailable As Boolean
Private mstrBlockReason As String
Private mbooDoubleBook As Boolean
Private mbooReadWeek As Boolean
Private strPrevWeek As String
Private mmlngClickPrev As Long
Private mlngClickNext As Long
Private mGridClick As Boolean
Public mstrStatus As String
Public mlngPatientID As Long
Private mbooPatValid As Boolean
Private mlngActCol As Long
Private Color As OLE_COLOR
Private mstrPresent As String
Public mlngClickPrev As Long
Private mboobyppas As Boolean
Private mbooPrintPreview As Boolean
Private mstrRangeDate As String
Private mstrUpperRange As String
Dim mbooDup As Boolean
Private Const NO_SIDES = 0
Private Const LEFT_SIDE = 1
Private Const RIGHT_SIDE = 2
Private Const TOP_SIDE = 4
Private Const BOTTOM_SIDE = 8
Private Const ALL_SIDES = 16
Private Const SOLID = 1
Private mlngPatCnt As Long
Private mstrCheckTime As String
Private mbooSchChange As Boolean
Private mbooTeleChange As Boolean
Dim mbooSupport As Boolean
Private udtColor As udtColor
Private Type udtColor
    Office As Long
    Hospital As Long
    Vacation As Long
    Cancelled As Long
    CheckOut As Long
    NoShow As Long
    Present As Long
    Lunch As Long
    PostOp As Long
    PreOp As Long
    NewPatient As Long
    ExistPatient As Long
    Surgery As Long
    FollowUp As Long
End Type
Private udtEnctr As udtEnctr
Private Type udtEnctr
    EnctrID As Long
    StartDate As String
    EndDate As String
    VisitReason As String
    Facility As String
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
Private udtBlock As udtBlock
Private Type udtBlock
    StartDate As String
    EndDate As String
    StartTime As String
    EndTime As String
    Reason As String
    Provider  As String
    ProviderID As Long
    FacilityID As Long
    Duration As String
    DayWeek As String
End Type
Private clsShedGrid As clsShedGrid
Private Type clsShedGrid
    CellSpan As String
    CellTag As String
    CellNote As String
    BackColor As Long
    Text As String
    TextNote As String
    List As String
    Duration As Long
    InsertRow As Long
    ProvCol As Long
    ProvCol2 As Long
    SpanRow As Long
    ApptDate As String
    StartRow As Long
    EndRow As Long
    Col As Long
    Row As Long
    RowName As String
    ColHeader As String
    CellType As String
End Type
Private Function Add_NextVisit(vstrDate As String)
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngPurge As Double
Dim lngAdmit_ID As Long
Dim lngPrimKey As Long

    On Error GoTo Error_Add_NextVisit
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    If udtInPatSched.VisitType = vbNullString Then udtInPatSched.VisitType = "Subsequent Visit"
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "SELECT * FROM INPATIENT_APPOINTMENTS WHERE PATIENT_ID = " & Val(udtInPatSched.IntId) & " AND " _
        & "APPT_DATE = '" & Format(vstrDate, "mm/dd/yyyy") & "' AND PROVIDER_ID = " & udtInPatSched.Provider_ID
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        With rsReadAppt
            If Not .EOF Then
                lngPrimKey = (!INPATIENT_APPT_ID)
            End If
        End With
    If lngPrimKey > 0 Then
        If IsProperCase(udtInPatSched.Status) = "Cancelled" Or IsProperCase(udtInPatSched.Status) = "Rescheduled" Then
            lngPurge = 1
        Else
            lngPurge = 0
        End If
        strSQL = "UPDATE INPATIENT_APPOINTMENTS SET APPT_DATE = '" & vstrDate & "', " _
            & "ROOM_NUMBER = '" & udtInPatSched.RoomNumber & "', PURPOSE_VISIT = '" & udtInPatSched.VisitType & "'," _
            & "ACTIVE = " & lngPurge & "  WHERE INPATIENT_APPT_ID = " & lngPrimKey
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
    Else
        strSQL = "INSERT INTO INPATIENT_APPOINTMENTS (PATIENT_ID, PROVIDER_ID, FACILITY_ID, ROOM_NUMBER," _
            & "APPT_DATE, PURPOSE_VISIT, ACTIVE) values (" & Val(udtInPatSched.IntId) & ", " _
            & udtInPatSched.Provider_ID & "," & udtInPatSched.Facility_ID & ",'" & udtInPatSched.RoomNumber & "','" _
            & vstrDate & "', '" & udtInPatSched.VisitType & "',0)"
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
    End If
    strSQL = "SELECT PATIENT_ADMIT_ROOMID FROM PATIENT_ADMISSION_ROOMS WHERE " _
        & "PATIENT_ID = " & Val(udtInPatSched.IntId) & " AND FACILITY_ID = " & udtInPatSched.Facility_ID
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        With rsReadAppt
            If Not .EOF Then
                lngAdmit_ID = (!PATIENT_ADMIT_ROOMID)
            End If
        End With
        If lngAdmit_ID > 0 Then
            strSQL = "UPDATE PATIENT_ADMISSION_ROOMS SET ROOM = '" & udtInPatSched.RoomNumber & "' WHERE " _
                & "PATIENT_ADMIT_ROOMID = " & lngAdmit_ID
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
        Else
            strSQL = "INSERT INTO PATIENT_ADMISSION_ROOMS (PATIENT_ID,  FACILITY_ID, ROOM) values " _
                & "(" & Val(udtInPatSched.IntId) & ", " & udtInPatSched.Facility_ID & ",'" & udtInPatSched.RoomNumber & "')"
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
        End If
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Add_NextVisit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Add_NextVisit", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Calculate_NextVisit(vlngNumber As Double, strTime As String)
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
Dim strAptDay As String
Dim strWeekDay As String

    On Error GoTo Error_Calculate_NextVisit
    Screen.MousePointer = vbHourglass
    nDays = vlngNumber
    If strWeekDay = vbNullString Then
        strWeekDay = Format(udtInPatSched.Appt_Date, "dddd")
    End If
    If vlngNumber <= 0 Or Trim(strTime) = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If nDays >= 1 Then
        Select Case strTime
            Case "Day"
                Calculate_NextVisit = DateAdd("d", nDays, udtInPatSched.Appt_Date)
            Case "Weeks"
                strRepDate = DateAdd("ww", nDays, udtInPatSched.Appt_Date)
                If strRepDate <> vbNullString Then
                    Select Case strWeekDay
                        Case "Monday"
                            FirstDay = WeekDay(strRepDate, vbMonday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Tuesday"
                            FirstDay = WeekDay(strRepDate, vbTuesday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Wednesday"
                            FirstDay = WeekDay(strRepDate, vbWednesday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Thursday"
                            FirstDay = WeekDay(strRepDate, vbThursday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Friday"
                            FirstDay = WeekDay(strRepDate, vbFriday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Saturday"
                            FirstDay = WeekDay(strRepDate, vbSaturday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Sunday"
                            FirstDay = WeekDay(strRepDate, vbSunday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                    End Select
                End If
                Calculate_NextVisit = Format(strRepDate, "mm/dd/yyyy")
                
            Case "Month"
                strRepDate = DateAdd("m", nDays, udtInPatSched.Appt_Date)
                If Trim(udtInPatSched.Appt_Date) <> vbNullString Then
                    udtInPatSched.Appt_Date = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
                End If
                If strRepDate <> vbNullString Then
                    Select Case strWeekDay
                        Case "Monday"
                            FirstDay = WeekDay(strRepDate, vbMonday)
                            strRepDate = DateAdd("d", -FirstDay, strRepDate)
                            strTestMonth = Month(strRepDate)
                        Case "Tuesday"
                            FirstDay = WeekDay(strRepDate, vbTuesday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                            strTestMonth = Month(strRepDate)
                        Case "Wednesday"
                            FirstDay = WeekDay(strRepDate, vbWednesday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Thursday"
                            FirstDay = WeekDay(strRepDate, vbThursday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                            strTestMonth = Month(strRepDate)
                        Case "Friday"
                            FirstDay = WeekDay(strRepDate, vbFriday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                            strTestMonth = Month(strRepDate)
                        Case "Saturday"
                            FirstDay = WeekDay(strRepDate, vbSaturday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                            strTestMonth = Month(strRepDate)
                        Case "Sunday"
                            FirstDay = WeekDay(strRepDate, vbSunday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                            strTestMonth = Month(strRepDate)
                    End Select
                End If
                Calculate_NextVisit = Format(strRepDate, "mm/dd/yyyy")
            Case "Annual"
                strRepDate = DateAdd("y", nDays, udtInPatSched.Appt_Date)
                If Trim(udtInPatSched.Appt_Date) <> vbNullString Then
                    udtInPatSched.Appt_Date = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
                End If
                If strRepDate <> vbNullString Then
                    Select Case strWeekDay
                        Case "Monday"
                            FirstDay = WeekDay(strRepDate, vbMonday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Tuesday"
                            FirstDay = WeekDay(strRepDate, vbTuesday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Wednesday"
                            FirstDay = WeekDay(strRepDate, vbWednesday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Thursday"
                            FirstDay = WeekDay(strRepDate, vbThursday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Friday"
                            FirstDay = WeekDay(strRepDate, vbFriday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Saturday"
                            FirstDay = WeekDay(strRepDate, vbSaturday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Case "Sunday"
                            FirstDay = WeekDay(strRepDate, vbSunday)
                            strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                    End Select
                    Calculate_NextVisit = Format(strRepDate, "mm/dd/yyyy")
                End If
        End Select
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Calculate_NextVisit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Calculate_NextVisit", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Change_Facility()
Dim Idx As Integer
Dim lngProv As Long
Dim lngRow As Long
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim strSQL As String
Dim lngCnt As Long

    Screen.MousePointer = vbHourglass
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    mstrRangeDate = vbNullString
    strSQL = "SELECT * FROM PROVIDERS WHERE TITLE = 'MD' AND ACTIVE = 1"
    Set rsPat = cnPat.Execute(strSQL)
    With rsPat
        Do Until .EOF
            lngCnt = lngCnt + 1
            lngProv = (!Provider_ID)
            .MoveNext
        Loop
    End With
    If lngCnt = 1 Then
        For Idx = 0 To cboProviders.listcount - 1
            If cboProviders.ItemData(Idx) = lngProv Then
                mbooByPass = True
                cboProviders.Text = cboProviders.List(Idx)
                Exit For
            End If
        Next Idx
    End If
    Load_Resource_Providers
    mbooByPass = True
    mlngPatCnt = 0
    lblCurScheAppt.Caption = vbNullString
    grdPatient.ClearRange 1, 1, grdPatient.MaxCols, grdPatient.MaxRows, False
    grdPatient.MaxRows = 0
    Get_Appt_Providers
    ReadData
    mbooByPass = False
    If cnPat.State = adStateOpen Then
        cnPat.Close
        Set cnPat = Nothing
        Set rsPat = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
End Function

Private Function Determine_Visit_Type(vlngPatient As Long)
Dim cnSearch As New ADODB.Connection
Dim rsSearch As New ADODB.Recordset
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngPatientID As Long
Dim strAC As String
Dim Arr
Dim strDate As String
Dim lngDateDiff As Long
Dim strReAdmit As String

    On Error GoTo Error_Determine_Visit_Type
    Screen.MousePointer = vbHourglass
    lngPatientID = vlngPatient
    If rsSearch.State = adStateClosed Then cnSearch.open psConnect(0)
    strSQL = "SELECT * FROM INPATIENT_APPOINTMENTS  " _
        & "WHERE PATIENT_ID =  " & lngPatientID & " AND " _
        & "PROVIDER_ID = " & udtInPatSched.Provider_ID & " AND FACILITY_ID = " & udtInPatSched.Facility_ID & " ORDER BY INSERT_DATE DESC"
        Set rsSearch = cnSearch.Execute(strSQL)
    With rsSearch
       If Not .EOF Then
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PATIENT_TYPE"
                                udtInPatSched.PatientType = Trim(Fld.Value)
                            Case "CURRENT_LOCATION"
                                strReAdmit = Trim(Fld.Value)
                            Case "STATUS"
                                udtInPatSched.Status = Trim(Fld.Value)
                            Case "INSERT_DATE"
                                strDate = Trim(Fld.Value)
                        End Select
                    End If
                End If
            Next
        End If
    End With

    If strReAdmit = "Re-Admission" Then
        udtInPatSched.PatientType = "Re-Admission"
    End If
    If Trim(strDate) <> vbNullString Then
        lngDateDiff = DateDiff("d", Format(strDate, "mm/dd/yyyy"), Format(Now, "mm/dd/yyyy"))
    End If
    If udtInPatSched.Status = "Cancelled" Or udtInPatSched.Status = "InActive" Then
        If lngDateDiff > 0 And lngDateDiff < 10 Then
            Determine_Visit_Type = True
        End If
    End If
    If IsProperCase(udtInPatSched.Status) = "Workup Completed" Or IsProperCase(udtInPatSched.Status) = "Pospone" And lngDateDiff > 0 And lngDateDiff < 10 Then
        Determine_Visit_Type = True
    End If
    'End If
    If rsSearch.State = adStateClosed Then
        cnSearch.Close
        Set cnSearch = Nothing
        Set rsSearch = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Determine_Visit_Type:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Determine_Visit_Type", App.EXEName & "." & "ModConst"
    Exit Function
    Resume
End Function

Private Function Fix_Comments(vstrComment As String) As String
Dim Arr
Dim Idx%

    Arr = Split(vstrComment, vbNewLine)
    If UBound(Arr) > 0 Then
        For Idx = LBound(Arr) To UBound(Arr)
            If Fix_Comments = vbNullString Then
                Fix_Comments = Trim(Arr(Idx))
            Else
                Fix_Comments = Fix_Comments & vbNewLine & Trim(Arr(Idx))
            End If
        Next Idx
    Else
        Fix_Comments = Trim(vstrComment)
    End If
    Exit Function
End Function

Private Function Format_Add_Patient_Grid()
Dim lngCol As Long
Dim lngRow As Long
Dim i%

    Screen.MousePointer = vbHourglass
    grdAppt.ProcessTab = True
    grdAppt.FontBold = True
    grdAppt.ForeColor = 0
    With grdAppt
        .Redraw = False
        .MaxRows = 0
        .MaxCols = 5
        .ClearRange 1, 1, grdAppt.MaxCols, grdAppt.MaxRows, False
        .ColHeaderRows = 4
        .ColWidth(1) = 20
        .ColWidth(2) = 10
        .ColWidth(3) = 30
        .ColWidth(4) = 10
        .ColWidth(5) = 13.6
        Format_CellTypeStatic grdAppt, SpreadHeader, 1, "maxcols", "ADD PATIENT TO VISIT ROSTER", "TypeHAlignCenter", "APPT_DATE", "", HeaderColor
        grdAppt.SetCellBorder 1, SpreadHeader, 5, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdAppt
        lngCol = 1
        Format_CellTypeEdit grdAppt, SpreadHeader + 2, lngCol, "1", "PATIENT NAME", "TypeHAlignCenter", "PATIENT_ID", HeaderColor
        grdAppt.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdAppt, SpreadHeader + 2, lngCol, "1", "PATIENT STATUS", "TypeHAlignCenter", "PATIENT STATUS", HeaderColor
        grdAppt.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeEdit grdAppt, SpreadHeader + 2, lngCol, "1", "FACILITY", "TypeHAlignCenter", "FACILITY", HeaderColor
        grdAppt.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeEdit grdAppt, SpreadHeader + 2, lngCol, "1", "ROOM", "TypeHAlignCenter", "ROOM", HeaderColor
        grdAppt.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdAppt, SpreadHeader + 2, lngCol, "1", "VISIT TYPE", "TypeHAlignCenter", "VISIT TYPE", HeaderColor
        grdAppt.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdAppt
        .Redraw = True
    End With
    mbooByPass = True
    Screen.MousePointer = vbDefault
    Exit Function
End Function

Private Function Get_Appt_Range()
Dim lngRow As Long
Dim lngCol As Long
Dim strSQL As String
Dim cnEnctr As New ADODB.Connection
Dim rsEnctr As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngEnctr As Long
Dim strEntrDate As String
Dim strEnctrType As String
Dim strCurState As String
Dim lngCurEnctr As Long

    On Error GoTo Error_Get_Appt_Range
    If cnEnctr.State = adStateClosed Then cnEnctr.open psConnect(0)
    mstrUpperRange = vbNullString
    mstrRangeDate = vbNullString
'----------
    strSQL = "SELECT APPT_DATE FROM INPATIENT_APPOINTMENTS WHERE FACILITY_ID = " & udtInPatSched.Facility_ID & " AND " _
        & "ACTIVE = 0 AND (STATUS = 'PRESENT' OR STATUS = 'MD COMPLETED') ORDER BY APPT_DATE DESC"
    Set rsEnctr = cnEnctr.Execute(strSQL)
    With rsEnctr
        If Not .EOF Then
            mstrRangeDate = (!Appt_Date)
        End If
    End With
    mstrRangeDate = DateAdd("d", 1, mstrRangeDate)
'------------
    If mstrRangeDate = vbNullString Then
        strSQL = "SELECT APPT_DATE FROM INPATIENT_APPOINTMENTS WHERE FACILITY_ID = " & udtInPatSched.Facility_ID & " AND " _
            & "ACTIVE = 0 AND (STATUS <> 'PRESENT' AND STATUS <> 'MD COMPLETED' AND STATUS <> 'CANCELLED') ORDER BY APPT_DATE ASC"
        Set rsEnctr = cnEnctr.Execute(strSQL)
        With rsEnctr
            If Not .EOF Then
                mstrRangeDate = (!Appt_Date)
            End If
        End With
    End If
    strSQL = "SELECT APPT_DATE FROM INPATIENT_APPOINTMENTS WHERE FACILITY_ID = " & udtInPatSched.Facility_ID & " AND " _
        & "ACTIVE = 0 AND (STATUS <> 'PRESENT' AND STATUS <> 'MD COMPLETED' AND STATUS <> 'CANCELLED') ORDER BY APPT_DATE DESC"
    Set rsEnctr = cnEnctr.Execute(strSQL)
    With rsEnctr
        If Not .EOF Then
            mstrUpperRange = (!Appt_Date)
        End If
    End With
    strEntrDate = DateAdd("d", 1, udtInPatSched.Appt_Date)
    If mstrUpperRange = mstrRangeDate Then
        mstrUpperRange = Format(Now, "mm/dd/yyyy")
    ElseIf DateDiff("d", Format(strEntrDate, "mm/dd/yyyy"), Format(mstrUpperRange, "mm/dd/yyyy")) > 1 Then
        mstrUpperRange = Format(strEntrDate, "mm/dd/yyyy")
    End If
    If cnEnctr.State = adStateOpen Then
        cnEnctr.Close
        Set cnEnctr = Nothing
        Set rsEnctr = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Appt_Range:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Appt_Range", Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_New_Appt_Range()
Dim lngRow As Long
Dim lngCol As Long
Dim strSQL As String
Dim cnEnctr As New ADODB.Connection
Dim rsEnctr As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngEnctr As Long
Dim strEntrDate As String
Dim strEnctrType As String
Dim strCurState As String
Dim lngCurEnctr As Long

    On Error GoTo Error_Get_New_Appt_Range
    If cnEnctr.State = adStateClosed Then cnEnctr.open psConnect(0)
    mstrUpperRange = vbNullString
    mstrRangeDate = vbNullString
    strSQL = "SELECT APPT_DATE FROM INPATIENT_APPOINTMENTS WHERE FACILITY_ID = " & udtInPatSched.Facility_ID & " AND " _
        & "ACTIVE = 0 AND (STATUS = 'PRESENT' OR STATUS = 'MD COMPLETED') ORDER BY APPT_DATE DESC"
    Set rsEnctr = cnEnctr.Execute(strSQL)
    With rsEnctr
        If Not .EOF Then
            mstrRangeDate = (!Appt_Date)
        End If
    End With
    strSQL = "SELECT APPT_DATE FROM INPATIENT_APPOINTMENTS WHERE FACILITY_ID = " & udtInPatSched.Facility_ID & " AND " _
        & "ACTIVE = 0 AND (STATUS <> 'PRESENT' AND STATUS <> 'MD COMPLETED' AND STATUS <> 'CANCELLED') ORDER BY APPT_DATE DESC"
    Set rsEnctr = cnEnctr.Execute(strSQL)
    With rsEnctr
        If Not .EOF Then
            mstrUpperRange = (!Appt_Date)
        End If
    End With
    If mstrUpperRange = mstrRangeDate Or mstrUpperRange = vbNullString Then
        mstrUpperRange = Format(Now, "mm/dd/yyyy")
    End If
    If cnEnctr.State = adStateOpen Then
        cnEnctr.Close
        Set cnEnctr = Nothing
        Set rsEnctr = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_New_Appt_Range:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_New_Appt_Range", Me.Name
    Exit Function
    Resume Next
End Function

Private Function Get_Encounter_Not_Scheduled()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Arr
Dim lngPatID As Long
Dim Idx As Integer
Dim lngRow As Long
Dim varCellValue
Dim vTime
Dim i%
Dim lngCol As Long
Dim lngFoundID As Boolean
Dim lngProvID As Long
Dim strNotes As String
Dim bPatFound As Boolean
Dim lngApptID As Long
Dim lngDup As Long

    On Error GoTo Error_Get_Encounter_Not_Scheduled
    Screen.MousePointer = vbHourglass
    mlngPatCnt = 0
    rtfChanges.Text = vbNullString
    grdData.MaxRows = 0

    If mstrApptMode = vbNullString Then mstrApptMode = "Day"
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Get_Appointment_Data_Facility " & udtInPatSched.Facility_ID
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            If Not IsNull((!INSERT_DATE)) Then
                If Trim((!INSERT_DATE)) <> vbNullString Then
                    udtInPatSched.Appt_Date = Format((!INSERT_DATE), "mm/dd/yyyy")
                End If
            End If
            .MoveNext
        Loop
    End With
    If udtInPatSched.Provider_ID = 0 Then udtInPatSched.Provider_ID = 1000
    clsShedGrid.ApptDate = udtInPatSched.Appt_Date
    clsShedGrid.ApptDate = Format(clsShedGrid.ApptDate, "mm/dd/yyyy")
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Get_InPatient_Appointments '" & Format(udtInPatSched.Appt_Date, "mm/dd/yyyy") & "'," _
        & udtInPatSched.Facility_ID & "," & udtInPatSched.Provider_ID
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            Clear_InSched
            Clear_Appt
            lngPatID = 0
            udtInPatSched.Appointment_ID = 0
            lngApptID = 0
            udtInPatSched.SchedType = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "SUFFIX"
                                udtInPatSched.Suffix = Trim(Fld.Value)
                            Case "LAST_NAME"
                                udtInPatSched.LastName = IsProperCase(Fld.Value)
                            Case "MIDDLE_NAME"
                                udtInPatSched.MiddleInitial = IsProperCase(Fld.Value)
                            Case "FIRST_NAME"
                                udtInPatSched.FirstName = Fld.Value
                            Case "PREF_CONTACT"
                                udtInPatSched.RoomTelephone = Trim(Fld.Value)
                            Case "PATIENT_ID"
                                udtInPatSched.IntId = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtInPatSched.VisitType = Fld.Value
                            Case "APPT_TYPE"
                                udtInPatSched.SchedType = Trim(Fld.Value)
                            Case "STATUS"
                                udtInPatSched.Status = Trim(Fld.Value)
                            Case "STATUS_NOTES"
                                udtInPatSched.ApptComment = Fld.Value
                                udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "''", "'")
                            Case "PROVIDER_ID"
                                udtInPatSched.Provider_ID = Fld.Value
                            Case "APPT_DATE"
                                udtInPatSched.Appt_Date = Fld.Value
                            Case "FACILITY_ID"
                                udtInPatSched.Facility_ID = Fld.Value
                            Case "RESOURCE_TYPE"
                                udtInPatSched.ResourceType = Fld.Value
                            Case "ROOM_NUMBER"
                                udtInPatSched.RoomNumber = Fld.Value
                            Case "TELEPHONE"
                                udtInPatSched.RoomTelephone = Fld.Value
                            Case "DURATION"
                                udtInPatSched.Duration = Fld.Value
                            Case "PATIENT_TYPE"
                                udtInPatSched.PatientType = Fld.Value
                            Case "INPATIENT_APPT_ID"
                                lngApptID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If lngDup <> Val(udtInPatSched.IntId) Then
                If Not Validate_CurrentFacility(Val(udtInPatSched.IntId)) Then
                    udtInPatSched.Appointment_ID = lngApptID
                    lngFoundID = False
                    udtInPatSched.Title = Fix_Title(udtInPatSched.Title)
                    udtInPatSched.Suffix = Fix_Suffix(udtInPatSched.Suffix)
                    udtInPatSched.LastName = Fix_Name(udtInPatSched.LastName)
                    udtInPatSched.FirstName = Fix_Name(udtInPatSched.FirstName)
                    If udtInPatSched.FirstName <> vbNullString Then
                        udtInPatSched.Collection = Trim(udtInPatSched.LastName) & ", " & Trim(udtInPatSched.FirstName)
                    End If
                    If Trim(udtInPatSched.Suffix) <> vbNullString Then
                        udtInPatSched.Name = udtInPatSched.FirstName & " " & udtInPatSched.LastName & ", " & udtInPatSched.Suffix
                    Else
                        udtInPatSched.Name = udtInPatSched.FirstName & " " & udtInPatSched.LastName
                    End If
                    Get_Room_Number Val(udtInPatSched.IntId)
                    mlngPatCnt = mlngPatCnt + 1
                    udtInPatSched.Name = Replace(udtInPatSched.Name, "''", "'")
                    If grdData.MaxRows > 0 Then
                        For lngRow = 1 To grdData.MaxRows
                            grdData.Col = 1
                            grdData.Row = lngRow
                            lngPatID = grdData.GetRowItemData(lngRow)
                            If lngPatID <> Val(udtInPatSched.IntId) Then
                                grdData.MaxRows = grdData.MaxRows + 1
                                clsShedGrid.Row = grdData.MaxRows
                                Add_Row clsShedGrid.Row
                                lngFoundID = True
                                Exit For
                            Else
                               lngFoundID = True
                               Exit For
                            End If
                        Next lngRow
                    Else
                        grdData.MaxRows = grdData.MaxRows + 1
                        clsShedGrid.Row = grdData.MaxRows
                        Add_Row clsShedGrid.Row
                    End If
                End If
            End If
        Loop
    End With
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    'Get Encounter editions
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Encounter_Not_Scheduled:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Encounter_Not_Scheduled", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Get_Extender_Encounter()
Dim strSQL As String
Dim cnStatus As New ADODB.Connection
Dim rsStatus As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim strEnctTime As String
Dim lngEnctrID As Long
Dim iAnswer As Integer
Dim strDate As String
Dim lngTestProvider As Long
Dim i%
Dim Idx%
Dim strInsert As String
Dim varCellValue


    On Error GoTo Error_Get_Extender_Encounter
    Screen.MousePointer = vbHourglass
    Dump_EncounterStrData
    piEncounterID = 0
    For i = 0 To cboProviders.listcount - 1
        mbooByPass = True
        If cboProviders.List(i) = cboProviders.Text Then
            mbooByPass = True
            udtInPatSched.Provider_ID = cboProviders.ItemData(i)
            Exit For
        End If
    Next i
    Get_Patient piPatientID
    If IsProperCase(clsPatient.Status) = "Inactive" Or IsProperCase(clsPatient.Status) = "Deceased" Then
        MsgBox "Unable to create a visit for a" & " " & UCase(clsPatient.Status) & " " & "patient. Please correct record if patient is active!", vbInformation, "PatientTrac InPatient Roster"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    Get_Providers udtInPatSched.Provider_ID
    piEncounterID = 0
    If Trim(udtInPatSched.Appt_Date) <> vbNullString Then
        strDate = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
    Else
        strDate = Format(Now, "mm/dd/yyyy")
    End If
    udtInPatSched.Appt_Date = strDate
    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        .GetText 1, .ActiveRow, varCellValue
        udtInPatSched.Facility_ID = .GetRowItemData(.ActiveRow)
        udtInPatSched.Facility = varCellValue
    End With
    clsEncounter.EncounterID = 0
    Get_Appt_Range
    udtTime.StartDate = Format(strDate, "mm/dd/yyyy 00:00:00 AM/PM")
    udtTime.EndDate = DateAdd("h", 23, udtTime.StartDate)
    udtTime.EndDate = Format(udtTime.EndDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
    strSQL = "SELECT * FROM ENCOUNTER WHERE PATIENT_ID = " & piPatientID & " AND PROVIDER_ID = " & udtInPatSched.Provider_ID & " AND " _
        & "FACILITY_ID  = " & udtInPatSched.Facility_ID & " AND EXT_ENCOUNTER_ID IS NULL AND PURGE = 0 AND CURRENT_STATE <> 'MD COMPLETED' " _
        & "AND CONVERT(DATETIME, ENCNTR_START_DATE, 101) > '" & Format(DateAdd("d", -1, mstrRangeDate), "mm/dd/yyyy") & "' AND " _
        & "CURRENT_STATE IS NOT NULL ORDER BY ENCNTR_START_DATE DESC, CURRENT_STATE"
    Set rsStatus = cnStatus.Execute(strSQL)
    With rsStatus
        Do Until .EOF
            Dump_EncounterStrData
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case Fld.Name
                            Case "CURRENT_STATE"
                                clsEncounter.Encounter_State = Fld.Value
                            Case "ENCNTR_START_DATE"
                                clsEncounter.StartDate = Fld.Value
                            Case "RENDERING_PROVIDER"
                                clsEncounter.RendProv = Fld.Value
                            Case "FACILITY_ID"
                                clsEncounter.FacilityID = Fld.Value
                            Case "ENCOUNTER_ID"
                                clsEncounter.EncounterID = Fld.Value
                            Case "INSERT_DATE"
                                strInsert = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If clsEncounter.StartDate = vbNullString Then
                clsEncounter.EncounterID = 0
                Dump_EncounterStrData
            Else
                lngDateDiff = DateDiff("d", Format(clsEncounter.StartDate, "mm/dd/yyyy"), strDate)
                If lngDateDiff > 30 Then
                    clsEncounter.EncounterID = 0
                    Dump_EncounterStrData
                End If
            End If
            If clsEncounter.Encounter_State = "WORKUP" Then
                Exit Do
            End If
        Loop
    End With
    If clsEncounter.EncounterID > 0 Then
        strSQL = "Update_Current_Encounter_InPatient_Extender " & udtInPatSched.Provider_ID & "," _
            & mlngLogonUser & ",'" & udtTime.StartDate & "','WORKUP'," & clsEncounter.EncounterID
        Set rsStatus = cnStatus.Execute(strSQL)
        piEncounterID = clsEncounter.EncounterID
        Get_Encounter piEncounterID
'        If clsEncounter.EncounterType = vbNullString Then
'            If udtInPatSched.PatientType = "Follow-Up" Then
'
'            End If
'        End If
        If udtInPatSched.Facility_ID > 0 Then
            Get_Facilities udtInPatSched.Facility_ID
        End If
        If clsFacility.FacType <> vbNullString Then
            clsEncounter.ServicePlace = clsFacility.FacType
        End If
        If clsFacility.pos <> 0 Then
            clsEncounter.ServicePlaceCode = clsFacility.pos
        End If
        If udtInPatSched.PatientType = "Follow-Up" Then
            clsEncounter.EncounterType = "Subsequent Care"
            udtInPatSched.PatientType = "Subsequent Care"
        ElseIf udtInPatSched.PatientType = "Initial" Or udtInPatSched.PatientType = "Initial InPatient" Then
            clsEncounter.EncounterType = "Initial InPatient"
            udtInPatSched.PatientType = "Initial InPatient"
        ElseIf udtInPatSched.PatientType = "Re-Admission" Then
            clsEncounter.EncounterType = "Initial InPatient"
            udtInPatSched.PatientType = "Initial InPatient"
        End If
        If clsEncounter.ServicePlaceCode = "31" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99307"
                clsEncounter.VisitType = "Subsequent care, nursing facility E&M, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = "99304"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        ElseIf clsEncounter.ServicePlaceCode = "13" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99334"
                clsEncounter.VisitType = "Follow-Up visit, asst living facility, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = " 99325"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        End If
        If clsEncounter.ServicePlaceCode = "31" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99307"
                clsEncounter.VisitType = "Subsequent care, nursing facility E&M, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = "99304"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        ElseIf clsEncounter.ServicePlaceCode = "13" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99334"
                clsEncounter.VisitType = "Follow-Up visit, asst living facility, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = " 99325"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        End If
        clsEncounter.EncounterID = piEncounterID
        strSQL = "UPDATE ENCOUNTER SET ENCOUNT_CPT_INTRVN = '" & clsEncounter.Intrvn_CD & "', VISIT_TYPE = '" & clsEncounter.VisitType & "'," _
            & "ENCNTR_TYPE = '" & clsEncounter.EncounterType & "',SERVICE_PLACE = '" & clsEncounter.ServicePlace & "', " _
            & "SERVICE_PLACE_CD = '" & clsEncounter.ServicePlaceCode & "' WHERE ENCOUNTER_ID = " & piEncounterID
        Set rsStatus = cnStatus.Execute(strSQL)
    End If
    If clsEncounter.EncounterID = 0 Then
        strSQL = "Get_InPatient_Appt_ID " & piPatientID & ",'" & Format(udtInPatSched.Appt_Date, "mm/dd/yyyy") & "'," & udtInPatSched.Provider_ID
        Set rsStatus = cnStatus.Execute(strSQL)
        With rsStatus
            If Not .EOF Then
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "PATIENT_TYPE"
                                    udtInPatSched.PatientType = Trim(Fld.Value)
                                Case "PURPOSE_VISIT"
                                    udtAppointment.VisitType = Trim(Fld.Value)
                                Case "FACILITY_ID"
                                    clsEncounter.FacilityID = Fld.Value
                                    clsEncounter.FacilityID = Fld.Value
                            End Select
                        End If
                    End If
                Next
            End If
        End With
        Select Case udtAppointment.VisitType
            Case "Follow-Up", "Subsequent Care", "Subsequent Visit", "Subsequent care, nursing facility E&M, limited", "Follow-Up visit, asst living facility, limited"
            udtInPatSched.PatientType = "Subsequent Care"
            Case "Initial InPatient", "New Patient", "Initial Visit", "Initial E&M Care, Detailed Visit", "Initial E&M Care, Brief Visit", _
                "Initial E&M Care, Extended Visit", "Initial E&M Care, Complex Visit", "Diagnostic Interview Exam"
                udtInPatSched.PatientType = "Initial InPatient"
        End Select
        If PsUserID = vbNullString Then
            PsUserID = Get_PsUserID(psConnect(0))
        End If
        If udtInPatSched.PatientType = "Follow-Up" Then
            clsEncounter.EncounterType = "Subsequent Care"
            udtInPatSched.PatientType = "Subsequent Care"
        ElseIf udtInPatSched.PatientType = "Initial" Or udtInPatSched.PatientType = "Initial InPatient" Then
            clsEncounter.EncounterType = "Initial InPatient"
            udtInPatSched.PatientType = "Initial InPatient"
        ElseIf udtInPatSched.PatientType = "Re-Admission" Then
            clsEncounter.EncounterType = "Initial InPatient"
            udtInPatSched.PatientType = "Initial InPatient"
        End If
        If clsCptTemp.CPT_Code <> vbNullString Then
            clsEncounter.Intrvn_CD = clsCptTemp.CPT_Code
        End If
        If Trim(udtInPatSched.PatientType) <> vbNullString Then
            clsEncounter.EncounterType = udtInPatSched.PatientType
        End If
        If udtInPatSched.Facility_ID > 0 Then
            Get_Facilities udtInPatSched.Facility_ID
        End If
        If clsFacility.FacType <> vbNullString Then
            clsEncounter.ServicePlace = clsFacility.FacType
        End If
        If clsFacility.pos <> 0 Then
            clsEncounter.ServicePlaceCode = clsFacility.pos
        End If
        If clsEncounter.ServicePlaceCode = "31" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99307"
                clsEncounter.VisitType = "Subsequent care, nursing facility E&M, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = "99304"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        ElseIf clsEncounter.ServicePlaceCode = "13" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99334"
                clsEncounter.VisitType = "Follow-Up visit, asst living facility, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = " 99325"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        End If
        If clsEDI_Client.Site_ID = "1999" Then
            If udtInPatSched.Provider_ID = 0 Then udtInPatSched.Provider_ID = 1000
        End If
        If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
        strSQL = "Write_Encounter_Extender " & piPatientID & "," & udtInPatSched.Provider_ID & "," & udtInPatSched.Facility_ID & ",'" _
             & Format(udtTime.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "', '" & clsEncounter.Intrvn_CD & "','" & Trim(clsEncounter.VisitType) & "','" _
             & Trim(clsEncounter.EncounterType) & "','" & Trim(clsEncounter.ServicePlace) & "'," & clsEncounter.ServicePlaceCode & ",'" & PsUserID & "','" _
             & Format(udtTime.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," & mlngLogonUser & ",'" & Trim(udtAppointment.VisitType) & "','WORKUP'"
            Set rsStatus = cnStatus.Execute(strSQL)
            With rsStatus
                If Not .EOF Then
                    clsEncounter.EncounterID = (!Encounter_ID)
                End If
            End With
    End If
    piEncounterID = clsEncounter.EncounterID
    modSpecific.piEncounterID = clsEncounter.EncounterID
    piEncounterID = clsEncounter.EncounterID
    Get_Encounter piEncounterID
    If cnStatus.State = adStateOpen Then
        cnStatus.Close
        Set cnStatus = Nothing
        Set rsStatus = Nothing
        Set Flds = Nothing
    End If
    mbooPatValid = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Extender_Encounter:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Extender_Encounter", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_Cur_Sched_Enctr()
Dim strSQL As String
Dim cnStatus As New ADODB.Connection
Dim rsStatus As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim strEnctTime As String
Dim lngEnctrID As Long
Dim iAnswer As Integer
Dim strDate As String
Dim lngTestProvider As Long
Dim i%
Dim Idx%
Dim strInsert As String
Dim varCellValue


    On Error GoTo Error_Get_Cur_Sched_Enctr
    Screen.MousePointer = vbHourglass
    Dump_EncounterStrData
    piEncounterID = 0
    For i = 0 To cboProviders.listcount - 1
        mbooByPass = True
        If cboProviders.List(i) = cboProviders.Text Then
            mbooByPass = True
            udtInPatSched.Provider_ID = cboProviders.ItemData(i)
            Exit For
        End If
    Next i
'    Get_Patient piPatientID
'    If IsProperCase(clsPatient.Status) = "Inactive" Or IsProperCase(clsPatient.Status) = "Deceased" Then
'        MsgBox "Unable to create a visit for a" & " " & UCase(clsPatient.Status) & " " & "patient. Please correct record if patient is active!", vbInformation, "PatientTrac InPatient Roster"
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
    Get_Providers udtInPatSched.Provider_ID
    If Trim(udtInPatSched.Appt_Date) <> vbNullString Then
        strDate = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
    Else
        strDate = Format(Now, "mm/dd/yyyy")
    End If
    piEncounterID = 0
    udtInPatSched.Appt_Date = strDate
    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        .GetText 1, .ActiveRow, varCellValue
        udtInPatSched.Facility_ID = .GetRowItemData(.ActiveRow)
        udtInPatSched.Facility = varCellValue
    End With
    clsEncounter.EncounterID = 0
    udtTime.StartDate = Format(strDate, "mm/dd/yyyy 00:00:00 AM/PM")
    udtTime.EndDate = DateAdd("h", 23, udtTime.StartDate)
    udtTime.EndDate = Format(udtTime.EndDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
    strSQL = "SELECT * FROM ENCOUNTER WHERE PATIENT_ID = " & piPatientID & " AND PROVIDER_ID = " & udtInPatSched.Provider_ID & " AND " _
        & "FACILITY_ID  = " & udtInPatSched.Facility_ID & " AND EXT_ENCOUNTER_ID IS NULL AND PURGE = 0 AND CURRENT_STATE <> 'MD COMPLETED' " _
        & "AND CURRENT_STATE IS NOT NULL ORDER BY ENCNTR_START_DATE DESC, CURRENT_STATE"
    Set rsStatus = cnStatus.Execute(strSQL)
    With rsStatus
        Do Until .EOF
            Dump_EncounterStrData
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case Fld.Name
                            Case "CURRENT_STATE"
                                clsEncounter.Encounter_State = Fld.Value
                            Case "ENCNTR_START_DATE"
                                clsEncounter.StartDate = Fld.Value
                            Case "RENDERING_PROVIDER"
                                clsEncounter.RendProv = Fld.Value
                            Case "FACILITY_ID"
                                clsEncounter.FacilityID = Fld.Value
                            Case "ENCOUNTER_ID"
                                piEncounterID = Fld.Value
                            Case "INSERT_DATE"
                                strInsert = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If clsEncounter.StartDate = vbNullString Then
                piEncounterID = 0
                Dump_EncounterStrData
            Else
                lngDateDiff = DateDiff("d", Format(clsEncounter.StartDate, "mm/dd/yyyy"), strDate)
                If lngDateDiff > 30 Then
                    piEncounterID = 0
                    Dump_EncounterStrData
                End If
            End If
            If clsEncounter.Encounter_State = "WORKUP" Then
                Exit Do
            End If
        Loop
    End With
    If cnStatus.State = adStateOpen Then
        cnStatus.Close
        Set cnStatus = Nothing
        Set rsStatus = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Cur_Sched_Enctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Cur_Sched_Enctr", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Get_LastVisit()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim i%
Dim lngDateDiff As Long

    On Error GoTo Error_Get_LastVisit
    Screen.MousePointer = vbHourglass
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "SELECT APPT_DATE  FROM INPATIENT_APPOINTMENTS WHERE PATIENT_ID = " & Val(udtInPatSched.IntId) & "  ORDER BY CONVERT(DATETIME, APPT_DATE, 101) DESC"
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            Get_LastVisit = vbNullString
            Get_LastVisit = (!Appt_Date)
            .MoveNext
            If IsDate(Get_LastVisit) Then
                lngDateDiff = DateDiff("d", Format(Get_LastVisit, "mm/dd/yyyy"), Format(Now, "mm/dd/yyyy"))
                If lngDateDiff < 0 Then
                    Exit Do
                Else
                    Get_LastVisit = vbNullString
                End If
            End If
        Loop
    End With
    If Get_LastVisit = vbNullString Then
        strSQL = "SELECT ENCNTR_START_DATE, ENCOUNTER_ID FROM ENCOUNTER WHERE PATIENT_ID = " & piPatientID & " AND " _
            & "CONVERT(DATETIME, ENCNTR_START_DATE, 101) < '" & Format(Now, "mm/dd/yyyy") & "' AND " _
            & "PROVIDER_ID = " & udtInPatSched.Provider_ID & " AND PURGE = 0 ORDER BY ENCNTR_START_DATE DESC"
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
            With rsReadAppt
                Do Until .EOF
                    Get_LastVisit = vbNullString
                    Get_LastVisit = (!ENCNTR_START_DATE)
                    .MoveNext
                    If Get_LastVisit <> vbNullString Then
                        Exit Do
                    End If
                Loop
            End With
    End If
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_LastVisit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_LastVisit", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_Meds()
Dim strSQL As String
Dim cnCurMeds As New ADODB.Connection
Dim rsCurMeds As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim booSkip As Boolean
Dim lngDateDiff As Long
Dim lngCurMedID As Long
Dim lngMedID As Long
Dim strMed As String

    On Error GoTo Error_Get_Meds
    Screen.MousePointer = vbHourglass
    Clear_UdtMed
    If cnCurMeds.State = adStateClosed Then cnCurMeds.open psConnect(0)
    strSQL = "SELECT * FROM CURRENT_MEDICATIONS WHERE PATIENT_ID = " & piPatientID & " AND " _
        & "PURGE = 0 AND PATIENT_MED_ID > 0   ORDER BY MEDICATION,  MEDICATION_DATE DESC, CHANGE_DATE DESC"
    Set rsCurMeds = cnCurMeds.Execute(strSQL)
    With rsCurMeds
        Do Until .EOF
            udtMedData.IngredientName = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "CURMED_ID"
                                lngCurMedID = Fld.Value
                           Case "INGREDIENT_NAME"
                                udtMedData.IngredientName = Fld.Value
                           Case "PATIENT_MED_ID"
                                udtMedData.PatientMedId = Fld.Value
                           Case "PT_MEDICATION_ID"
                                 udtMedData.PT_MedID = Fld.Value
                           Case "REASON"
                                udtMedData.Reason = Fld.Value
                           Case "INSTRUCTIONS"
                                udtMedData.MedDirections = Fld.Value
                           Case "LISTING_SEQ_NO"
                                udtMedData.Listing = Fld.Value
                           Case "CURRENT_STATUS"
                                udtMedData.Status = Fld.Value
                           Case "ORDER_PHYSICIAN"
                                udtMedData.OrderPys = Fld.Value
                           Case "STRENGTH"
                                udtMedData.Strength = Fld.Value
                           Case "UNITS"
                                udtMedData.Units = Fld.Value
                           Case "MEDICATION_DATE"
                                udtMedData.StartDate = Format(Fld.Value, "mm/dd/yyyy")
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtMedData.IngredientName <> strMed Then
                If lngCurMedID > 0 Then
                    If Trim(udtMedData.Strength) <> vbNullString Then udtMedData.Strength = Trim(udtMedData.Strength)
                    If Trim(udtMedData.Units) <> vbNullString Then udtMedData.Units = Trim(udtMedData.Units)
                    If Trim(udtMedData.Refills) <> vbNullString Then udtMedData.Refills = Trim(udtMedData.Refills)
                    If udtMedData.IngredientName = vbNullString Then
                        If Trim(udtMedData.IngredientName) <> vbNullString Then udtMedData.IngredientName = Trim(udtMedData.IngredientName)
                    End If
                    If udtInPatSched.Medication = vbNullString Then
                        udtInPatSched.Medication = Trim(udtMedData.Strength) & " " & Trim(udtMedData.Units) & " " & Trim(udtMedData.IngredientName)
                    Else
                        udtInPatSched.Medication = udtInPatSched.Medication & vbNewLine & Trim(udtMedData.Strength) & " " & Trim(udtMedData.Units) & " " & Trim(udtMedData.IngredientName)
                    End If
                End If
            End If
            strMed = udtMedData.IngredientName
        Loop
    End With
    If cnCurMeds.State = adStateOpen Then
        cnCurMeds.Close
        Set cnCurMeds = Nothing
        Set rsCurMeds = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Meds:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Meds", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_MissingData(vlngPatMedID As Long, vstrDate As String, Optional vlngCurMedID As Long)
Dim strSQL As String
Dim cnMisData As New ADODB.Connection
Dim rsMisData As New ADODB.Recordset
Dim i%
Dim strMedDate As String
Dim lngDateDiff As Long

    On Error GoTo Error_Get_MissingData
    Screen.MousePointer = vbHourglass

    If cnMisData.State = adStateClosed Then cnMisData.open psConnect(0)
    If vlngCurMedID > 0 And Trim(vstrDate) <> vbNullString Then
        strSQL = "SELECT PATIENT_MED_ID FROM CURRENT_MEDICATIONS WHERE CURMED_ID = " & vlngCurMedID
        Set rsMisData = cnMisData.Execute(strSQL)
        If Not rsMisData.EOF Then
            vlngPatMedID = rsMisData.Fields("PATIENT_MED_ID")
        End If
    End If
    If cnMisData.State = adStateClosed Then cnMisData.open psConnect(0)
    strSQL = "SELECT * FROM PRESCRIPTION_HISTORY WHERE PATIENT_MED_ID = " & vlngPatMedID & " ORDER BY START_DATE DESC"
    Set rsMisData = cnMisData.Execute(strSQL)
    With rsMisData
        Do Until .EOF
            Dump_MedStrData
            If (!START_DATE) <> vbNullString Then strMedDate = Format((!START_DATE), "mm/dd/yyyy")
            If vstrDate <> vbNullString Then vstrDate = Format(vstrDate, "mm/dd/yyyy")
            If IsDate(vstrDate) Then
                lngDateDiff = DateDiff("d", strMedDate, vstrDate)
                If lngDateDiff = 0 Then
                    If Trim(udtCurrentMed.Strength) = vbNullString And (!Strength) <> vbNullString Then udtCurrentMed.Strength = (!Strength)
                    If Trim(udtCurrentMed.Units) = vbNullString And (!Units) <> vbNullString Then udtCurrentMed.Units = (!Units)
                    If Trim(udtCurrentMed.IngredientName) = vbNullString And (!medication_name) <> vbNullString Then udtCurrentMed.IngredientName = (!medication_name)
                    If Trim(udtCurrentMed.DosageForm) = vbNullString And (!Dosage_Form) <> vbNullString Then udtCurrentMed.DosageForm = (!Dosage_Form)
                    If Trim(udtCurrentMed.Route) = vbNullString And (!Route) <> vbNullString Then udtCurrentMed.Route = (!Route)
                    If Trim(udtCurrentMed.MedDirections) = vbNullString And (!Instructions) <> vbNullString Then udtCurrentMed.MedDirections = (!Instructions)
                    If Trim(udtCurrentMed.Dispense) = vbNullString And (!Dispense_No) <> vbNullString Then udtCurrentMed.Dispense = (!Dispense_No)
                    If Trim(udtCurrentMed.Refills) = vbNullString And (!Refills) <> vbNullString Then udtCurrentMed.Refills = (!Refills)
                    If Trim(udtCurrentMed.Listing) = vbNullString And (!Listing_Seq_No) <> vbNullString Then udtCurrentMed.Listing = (!Listing_Seq_No)
                    If Trim(udtCurrentMed.Reason) = vbNullString And (!Reason) <> vbNullString Then udtCurrentMed.Reason = (!Reason)
                    Exit Do
                End If
            End If
            .MoveNext
            Loop
        End With
    If udtCurrentMed.IngredientName = vbNullString Then
        strSQL = "SELECT * FROM PATIENT_MEDICATIONS WHERE PATIENT_MED_ID = " & vlngPatMedID
        Set rsMisData = cnMisData.Execute(strSQL)
        With rsMisData
            If Not .EOF Then
                If Trim(udtCurrentMed.Strength) = vbNullString And (!Strength) <> vbNullString Then udtCurrentMed.Strength = (!Strength)
                If Trim(udtCurrentMed.Units) = vbNullString And (!Unit) <> vbNullString Then udtCurrentMed.Units = (!Unit)
                If Trim(udtCurrentMed.IngredientName) = vbNullString And (!INGREDIENT_NAME) <> vbNullString Then udtCurrentMed.IngredientName = (!INGREDIENT_NAME)
                If Trim(udtCurrentMed.DosageForm) = vbNullString And (!Dosage_Form) <> vbNullString Then udtCurrentMed.DosageForm = (!Dosage_Form)
                If Trim(udtCurrentMed.Route) = vbNullString And (!ROUTE_NAME) <> vbNullString Then udtCurrentMed.Route = (!ROUTE_NAME)
                If Trim(udtCurrentMed.MedDirections) = vbNullString And (!MED_DIRECTIONS) <> vbNullString Then udtCurrentMed.MedDirections = (!MED_DIRECTIONS)
                If Trim(udtCurrentMed.Dispense) = vbNullString And (!Dispense_No) <> vbNullString Then udtCurrentMed.Dispense = (!Dispense_No)
                If Trim(udtCurrentMed.Refills) = vbNullString And (!Refill_number) <> vbNullString Then udtCurrentMed.Refills = (!Refill_number)
                If Trim(udtCurrentMed.Listing) = vbNullString And (!Listing_Seq_No) <> vbNullString Then udtCurrentMed.Listing = (!Listing_Seq_No)
                If Trim(udtCurrentMed.Reason) = vbNullString And (!Reason) <> vbNullString Then udtCurrentMed.Reason = (!Reason)
            End If
        End With
    End If
     If cnMisData.State = adStateOpen Then
        cnMisData.Close
        Set cnMisData = Nothing
        Set rsMisData = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_MissingData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_MissingData", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_NextVisit() As String
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim i%
Dim strTime As String
Dim lngNumber As Double
Dim strDate As String
Dim lngDateDiff As Long

    On Error GoTo Error_Get_NextVisit
    Screen.MousePointer = vbHourglass
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "SELECT EVALUATION_MANAGEMENT.FOLLOW_UP, EVALUATION_MANAGEMENT.TERM, ENCOUNTER.ENCNTR_START_DATE FROM " _
        & "EVALUATION_MANAGEMENT INNER JOIN ENCOUNTER ON ENCOUNTER.ENCOUNTER_ID = EVALUATION_MANAGEMENT.ENCOUNTER_ID WHERE " _
        & "ENCOUNTER.PATIENT_ID = " & Val(udtInPatSched.IntId) & " AND EVALUATION_MANAGEMENT.FOLLOW_UP IS NOT NULL AND EVALUATION_MANAGEMENT.TERM IS NOT NULL AND " _
        & "ENCOUNTER.PROVIDER_ID = " & udtInPatSched.Provider_ID & " ORDER BY ENCOUNTER.ENCNTR_START_DATE DESC"
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            strDate = Format((!ENCNTR_START_DATE), "mm/dd/yyyy")
            If Not IsNull((!FOLLOW_UP)) Then
                If Val((!FOLLOW_UP)) > 0 Then
                    lngNumber = (!FOLLOW_UP)
                End If
            End If
            If Not IsNull((!Term)) Then
                If Trim((!Term)) <> vbNullString Then
                    strTime = (!Term)
                End If
            End If
            .MoveNext
            If IsDate(strDate) And IsDate(udtInPatSched.Appt_Date) Then
                lngDateDiff = DateDiff("d", Format(strDate, "mm/dd/yyyy"), Format(udtInPatSched.Appt_Date, "mm/dd/yyyy"))
                If lngDateDiff = 0 Then
                    Exit Do
                Else
                    lngNumber = 0
                    strTime = vbNullString
                End If
            End If
            
        Loop
    End With
    If Trim(strTime) <> vbNullString And Val(lngNumber) > 0 Then Get_NextVisit = Calculate_NextVisit(lngNumber, strTime)
    If IsDate(Get_NextVisit) Then
        Add_NextVisit Get_NextVisit
    End If
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_NextVisit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_NextVisit", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_Roster_Status_Dates() As String
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String

    On Error GoTo Error_Get_Roster_Status_Dates
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "SELECT DISTINCT APPT_DATE FROM INPATIENT_APPOINTMENTS WHERE " _
        & "INPATIENT_APPOINTMENTS.STATUS = 'WORKUP COMPLETED' ORDER BY APPT_DATE ASC"
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        If Not .EOF Then
            If ((!Appt_Date)) <> vbNullString Then
                If Trim((!Appt_Date)) <> "01/01/1900" Then
                    Get_Roster_Status_Dates = (!Appt_Date)
                End If
            End If
        End If
    End With
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Roster_Status_Dates:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Roster_Status_Dates", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Load_Active_Patient()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Idx%
Dim lngCol As Long
Dim lngRow As Long
Dim strName As String
Dim lngPatID As Long
Dim strDate As String
Dim strList As String
Dim lngFoundID As Boolean
Dim lngProvID As Long
Dim strNotes As String
Dim bPatFound As Boolean
Dim lngApptID As Long
Dim varCellValue
Dim lngTestPatient As Long
Dim booSkip As Boolean
Dim lngTest As Long
Dim strReAdmit As String


    On Error GoTo Error_Load_Active_Patient
    Screen.MousePointer = vbHourglass
    If clsEDI_Client.Site_ID = vbNullString Then Get_Edi_Static
    mstrShow = "All Patients"
    mlngPatCnt = 0
    Format_Patient
    clsShedGrid.ApptDate = udtInPatSched.Appt_Date
    clsShedGrid.ApptDate = Format(clsShedGrid.ApptDate, "mm/dd/yyyy")
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Search_All_InPatient_Facility_Join " & udtInPatSched.Facility_ID
'    strSQL = "SELECT DISTINCT  PATIENT.*, STATUS.STATUS_NOTES FROM PATIENT " _
'        & "INNER JOIN PATIENT_STATUS as STATUS on STATUS.PATIENT_ID = PATIENT.PATIENT_ID " _
'        & "WHERE  STATUS.STATUS_ITEM = CONVERT(VARCHAR, @FACILITY_ID,(10)) " _
'        & "ORDER BY PATIENT.LAST_NAME, PATIENT.FIRST_NAME"
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            strReAdmit = vbNullString
            booSkip = False
            Clear_InSched
            clsPatient.Status = vbNullString
            Clear_Appt
            lngPatID = 0
            udtInPatSched.SchedType = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "SUFFIX"
                                udtInPatSched.Suffix = Trim(Fld.Value)
                            Case "LAST_NAME"
                                udtInPatSched.LastName = IsProperCase(Fld.Value)
                            Case "MIDDLE_NAME"
                                udtInPatSched.MiddleInitial = IsProperCase(Fld.Value)
                            Case "FIRST_NAME"
                                udtInPatSched.FirstName = Fld.Value
                            Case "PREF_CONTACT"
                                udtInPatSched.RoomTelephone = Trim(Fld.Value)
                            Case "PATIENT_ID"
                                udtInPatSched.IntId = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtInPatSched.VisitType = Fld.Value
                            Case "APPT_TYPE"
                                udtInPatSched.SchedType = Trim(Fld.Value)
                            Case "STATUS"
                                clsPatient.Status = Trim(IsProperCase(Fld.Value))
                            Case "PROVIDER_ID"
                                udtInPatSched.Provider_ID = Fld.Value
                            Case "APPT_DATE"
                                udtInPatSched.Appt_Date = Fld.Value
                            Case "FACILITY_ID"
                                udtInPatSched.Facility_ID = Fld.Value
                            Case "RESOURCE_TYPE"
                                udtInPatSched.ResourceType = Fld.Value
                            Case "ROOM_NUMBER"
                                udtInPatSched.RoomNumber = Fld.Value
                            Case "TELEPHONE"
                                udtInPatSched.RoomTelephone = Fld.Value
                            Case "DURATION"
                                udtInPatSched.Duration = Fld.Value
                            Case "PATIENT_TYPE"
                                udtInPatSched.PatientType = Fld.Value
                            Case "CURRENT_LOCATION"
                                strReAdmit = Fld.Value
                            Case "INPATIENT_APPT_ID"
                                lngApptID = Fld.Value
                            Case "STATUS_NOTES"
                                udtInPatSched.ApptComment = Fld.Value
                                udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "''", "'")
                        End Select
                    End If
                End If
            Next
            .MoveNext
            For lngTest = 1 To grdPatient.MaxRows
                 lngTestPatient = grdPatient.GetRowItemData(lngTest)
                 If lngTestPatient = Val(udtInPatSched.IntId) Then
                    booSkip = True
                 End If
            Next lngTest
            If clsEDI_Client.Site_ID = "1999" And Not booSkip Then
                If Determine_Visit_Type(Val(udtInPatSched.IntId)) Then
                    booSkip = True
                End If
            End If
            If Not booSkip Then
                If Not Validate_CurrentFacility(Val(udtInPatSched.IntId)) Then
                    mlngPatCnt = mlngPatCnt + 1
                    lngRow = grdPatient.MaxRows + 1
                    grdPatient.MaxRows = grdPatient.MaxRows + 1
                    grdPatient.SetRowItemData lngRow, Val(udtInPatSched.IntId)
                    Get_NewInPatient Val(udtInPatSched.IntId)
                    Get_Room_Number Val(udtInPatSched.IntId)
                    udtInPatSched.Schedule_Date = Format(txtDate.Value, "mm/dd/yyyy")
                    lngCol = 1
                    Format_CellTypeEdit grdPatient, lngRow, lngCol, "1", udtInPatSched.LastName & ", " & udtInPatSched.FirstName & vbNewLine & udtInPatSched.IntId, "TypeHAlignLeft", "PATIENT_ID", 0, CStr(lngPatID)
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 2
                    Format_CellTypeEdit grdPatient, lngRow, lngCol, "1", udtInPatSched.RoomNumber, "TypeHAlignCenter", "ROOM_NUMBER"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 3
                    Get_Edi_Static
                    If clsEDI_Client.Site_ID = "1999" Then
                        If Determine_Visit_Type(Val(udtInPatSched.IntId)) Then
                            booSkip = True
                        End If
                        strList = "Initial" & Chr(9) & "Follow-Up" & Chr(9) & "Re-Admission"
                        Format_CellTypeCombo grdPatient, lngRow, lngCol, "1", strList, "TypeHAlignLeft", "PURPOSE_VISIT"
                        If udtInPatSched.PatientType = vbNullString Then
                            Set_GridCurSel grdPatient, "Follow-Up", lngCol, lngRow, "1", "PURPOSE_VISIT"
                        Else
                            If udtInPatSched.PatientType = "Subsequent Care" Then
                                Set_GridCurSel grdPatient, "Follow-Up", lngCol, lngRow, "1", "PURPOSE_VISIT"
                            ElseIf udtInPatSched.PatientType = "Re-Admission" Then
                                Set_GridCurSel grdPatient, "Re-Admission", lngCol, lngRow, "1", "PURPOSE_VISIT"
                            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                                Set_GridCurSel grdPatient, "Initial", lngCol, lngRow, "1", "PURPOSE_VISIT"
                            End If
                        End If
                        grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Else
                        Format_CellTypeCombo grdPatient, lngRow, lngCol, "1", Load_Visit, "TypeHAlignLeft", "PURPOSE_VISIT"
                    End If
                    If clsEDI_Client.Site_ID = "1125" Then
                        Set_GridCurSel grdPatient, "Subsequent Visit", lngCol, lngRow, "1", "PURPOSE_VISIT"
                        grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 4
                    Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", clsPatient.Status, "TypeHAlignCenter", "STATUS"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 5
                    If TextFound("No record for primary insurance", udtInPatSched.ApptComment, Len("No record for primary insurance"), True) Then
                        udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "No record for primary insurance", vbNullString)
                        udtInPatSched.ApptComment = Trim(udtInPatSched.ApptComment)
                    End If
                    Get_Insurance Val(udtInPatSched.IntId)
                    If Trim(clsInsurance.Prim_Insurer) <> vbNullString Then
                        If Not TextFound(clsInsurance.Prim_Insurer, udtInPatSched.ApptComment, Len(clsInsurance.Prim_Insurer), True) Then
                            If udtInPatSched.ApptComment = vbNullString Then
                                udtInPatSched.ApptComment = Trim(clsInsurance.Prim_Insurer)
                            Else
                                udtInPatSched.ApptComment = clsInsurance.Prim_Insurer & vbNewLine & Trim(udtInPatSched.ApptComment)
                            End If
                        End If
                    Else
                        If udtInPatSched.ApptComment = vbNullString Then
                            udtInPatSched.ApptComment = "No record for primary insurance"
                        Else
                            udtInPatSched.ApptComment = "No record for primary insurance" & vbNewLine & Trim(udtInPatSched.ApptComment)
                        End If
                    End If
                    Format_CellTypeEdit grdPatient, lngRow, lngCol, "1", Fix_Comments(udtInPatSched.ApptComment), "TypeHAlignCenter", "NOTES"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
            End If
        Loop
    End With
    lblCurScheAppt.Caption = "Active Patient Roster"
    lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Active_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_Active_Patient", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Function Load_All_Patients()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Arr
Dim lngPatID As Long
Dim Idx As Integer
Dim lngRow As Long
Dim varCellValue
Dim vTime
Dim i%
Dim lngCol As Long
Dim lngFoundID As Boolean
Dim lngProvID As Long
Dim strNotes As String
Dim bPatFound As Boolean
Dim lngApptID As Long
Dim strReAdmit As String

    On Error GoTo Error_Load_All_Patients
    Screen.MousePointer = vbHourglass
    mlngPatCnt = 0
    rtfChanges.Text = vbNullString
    grdData.MaxRows = 0
    mbooByPass = True
    For i = 0 To cboProviders.listcount - 1
        mbooByPass = True
        If cboProviders.List(i) = cboProviders.Text Then
            udtInPatSched.Provider_ID = cboProviders.ItemData(i)
            mbooByPass = True
            udtInPatSched.Provider = cboProviders.List(i)
            Exit For
        End If
    Next i
    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        .GetText 1, .ActiveRow, varCellValue
        udtInPatSched.Facility_ID = .GetRowItemData(.ActiveRow)
        udtInPatSched.Facility = varCellValue
    End With
    If mstrApptMode = vbNullString Then mstrApptMode = "Day"
    clsShedGrid.ApptDate = udtInPatSched.Appt_Date
    clsShedGrid.ApptDate = Format(clsShedGrid.ApptDate, "mm/dd/yyyy")
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "SELECT DISTINCT  PATIENT.*, STATUS.STATUS_NOTES FROM PATIENT " _
        & "INNER JOIN PATIENT_STATUS as STATUS on STATUS.PATIENT_ID = PATIENT.PATIENT_ID " _
        & "WHERE  STATUS.STATUS_ITEM = " & udtInPatSched.Facility_ID & " " _
        & "ORDER BY PATIENT.LAST_NAME, PATIENT.FIRST_NAME"
    'strSQL = "Search_All_InPatient_Facility_Join " & udtInPatSched.Facility_ID
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            strReAdmit = vbNullString
            Clear_InSched
            clsPatient.Status = vbNullString
            Clear_Appt
            lngPatID = 0
            udtInPatSched.SchedType = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "SUFFIX"
                                udtInPatSched.Suffix = Trim(Fld.Value)
                            Case "LAST_NAME"
                                udtInPatSched.LastName = IsProperCase(Fld.Value)
                            Case "MIDDLE_NAME"
                                udtInPatSched.MiddleInitial = IsProperCase(Fld.Value)
                            Case "FIRST_NAME"
                                udtInPatSched.FirstName = Fld.Value
                            Case "PREF_CONTACT"
                                udtInPatSched.RoomTelephone = Trim(Fld.Value)
                            Case "PATIENT_ID"
                                udtInPatSched.IntId = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtInPatSched.VisitType = Fld.Value
                            Case "APPT_TYPE"
                                udtInPatSched.SchedType = Trim(Fld.Value)
                            Case "STATUS"
                                clsPatient.Status = Trim(IsProperCase(Fld.Value))
                            Case "STATUS_NOTES"
                                udtInPatSched.ApptComment = Fld.Value
                                udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "''", "'")
                            Case "PROVIDER_ID"
                                udtInPatSched.Provider_ID = Fld.Value
                            Case "APPT_DATE"
                                udtInPatSched.Appt_Date = Fld.Value
                            Case "FACILITY_ID"
                                udtInPatSched.Facility_ID = Fld.Value
                            Case "RESOURCE_TYPE"
                                udtInPatSched.ResourceType = Fld.Value
                            Case "ROOM_NUMBER"
                                udtInPatSched.RoomNumber = Fld.Value
                            Case "TELEPHONE"
                                udtInPatSched.RoomTelephone = Fld.Value
                            Case "DURATION"
                                udtInPatSched.Duration = Fld.Value
                            Case "PATIENT_TYPE"
                                udtInPatSched.PatientType = Fld.Value
                            Case "INPATIENT_APPT_ID"
                                lngApptID = Fld.Value
                            Case "CURRENT_LOCATION"
                                strReAdmit = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            lngFoundID = False
            udtInPatSched.Title = Fix_Title(udtInPatSched.Title)
            udtInPatSched.Suffix = Fix_Suffix(udtInPatSched.Suffix)
            udtInPatSched.LastName = Fix_Name(udtInPatSched.LastName)
            udtInPatSched.FirstName = Fix_Name(udtInPatSched.FirstName)
            If udtInPatSched.FirstName <> vbNullString Then
                udtInPatSched.Collection = Trim(udtInPatSched.LastName) & ", " & Trim(udtInPatSched.FirstName)
            End If
            If Trim(udtInPatSched.Suffix) <> vbNullString Then
                udtInPatSched.Name = udtInPatSched.FirstName & " " & udtInPatSched.LastName & ", " & udtInPatSched.Suffix
            Else
                udtInPatSched.Name = udtInPatSched.FirstName & " " & udtInPatSched.LastName
            End If
            Get_Room_Number Val(udtInPatSched.IntId)
            mlngPatCnt = mlngPatCnt + 1
            udtInPatSched.Name = Replace(udtInPatSched.Name, "''", "'")
            If grdData.MaxRows > 0 Then
                For lngRow = 1 To grdData.MaxRows
                    grdData.Col = 1
                    grdData.Row = lngRow
                    udtInPatSched.Status = clsPatient.Status
                    lngPatID = grdData.GetRowItemData(lngRow)
                    If lngPatID <> Val(udtInPatSched.IntId) Then
                        grdData.MaxRows = grdData.MaxRows + 1
                        clsShedGrid.Row = grdData.MaxRows
                        Add_Row clsShedGrid.Row
                        lngFoundID = True
                        Exit For
                    Else
                       lngFoundID = True
                       Exit For
                    End If
                Next lngRow
            Else
                grdData.MaxRows = grdData.MaxRows + 1
                clsShedGrid.Row = grdData.MaxRows
                Add_Row clsShedGrid.Row
            End If
        Loop
    End With
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    udtInPatSched.Appt_Date = clsShedGrid.ApptDate
    mbooByPass = False
    mFirstLoad = False
    mbooByPass = False
    udtInPatSched.SchedType = "Patient"
    Get_Changes clsShedGrid.ApptDate
    Set_Headers
    Clear_InSched
    udtInPatSched.IntId = vbNullString
    lblCurScheAppt.Visible = True
    lblCurScheAppt.Caption = "Current Patient Roster"
    lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
    mbooByPass = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_All_Patients:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_All_Patients", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Load_InActive()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Idx%
Dim lngCol As Long
Dim lngRow As Long
Dim strName As String
Dim lngPatID As Long
Dim strDate As String
Dim strList As String
Dim lngFoundID As Boolean
Dim lngProvID As Long
Dim strNotes As String
Dim bPatFound As Boolean
Dim lngApptID As Long
Dim varCellValue
Dim lngTest As Long
Dim lngTestPatient As Long
Dim booSkip As Boolean

    On Error GoTo Error_Load_InActive
    Screen.MousePointer = vbHourglass
    mlngPatCnt = 0
    mstrShow = "All Patients"
    Format_Patient
    clsShedGrid.ApptDate = udtInPatSched.Appt_Date
    clsShedGrid.ApptDate = Format(clsShedGrid.ApptDate, "mm/dd/yyyy")
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Search_All_InActive_InPatient_Facility_Join " & udtInPatSched.Facility_ID
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            booSkip = False
            Clear_InSched
            clsPatient.Status = vbNullString
            Clear_Appt
            lngPatID = 0
            udtInPatSched.SchedType = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "SUFFIX"
                                udtInPatSched.Suffix = Trim(Fld.Value)
                            Case "LAST_NAME"
                                udtInPatSched.LastName = IsProperCase(Fld.Value)
                            Case "MIDDLE_NAME"
                                udtInPatSched.MiddleInitial = IsProperCase(Fld.Value)
                            Case "FIRST_NAME"
                                udtInPatSched.FirstName = Fld.Value
                            Case "PREF_CONTACT"
                                udtInPatSched.RoomTelephone = Trim(Fld.Value)
                            Case "PATIENT_ID"
                                udtInPatSched.IntId = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtInPatSched.VisitType = Fld.Value
                            Case "APPT_TYPE"
                                udtInPatSched.SchedType = Trim(Fld.Value)
                            Case "STATUS"
                                clsPatient.Status = Trim(IsProperCase(Fld.Value))
                            Case "STATUS_NOTES"
                                udtInPatSched.ApptComment = Fld.Value
                                udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "''", "'")
                            Case "PROVIDER_ID"
                                udtInPatSched.Provider_ID = Fld.Value
                            Case "APPT_DATE"
                                udtInPatSched.Appt_Date = Fld.Value
                            Case "FACILITY_ID"
                                udtInPatSched.Facility_ID = Fld.Value
                            Case "RESOURCE_TYPE"
                                udtInPatSched.ResourceType = Fld.Value
                            Case "ROOM_NUMBER"
                                udtInPatSched.RoomNumber = Fld.Value
                            Case "TELEPHONE"
                                udtInPatSched.RoomTelephone = Fld.Value
                            Case "DURATION"
                                udtInPatSched.Duration = Fld.Value
                            Case "PATIENT_TYPE"
                                udtInPatSched.PatientType = Fld.Value
                            Case "INPATIENT_APPT_ID"
                                lngApptID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            For lngTest = 1 To grdPatient.MaxRows
                 lngTestPatient = grdPatient.GetRowItemData(lngTest)
                 If lngTestPatient = Val(udtInPatSched.IntId) Then
                    booSkip = True
                 End If
            Next lngTest
            If Not booSkip Then
                If Not Validate_CurrentFacility(Val(udtInPatSched.IntId)) Then
                    mlngPatCnt = mlngPatCnt + 1
                    lngRow = grdPatient.MaxRows + 1
                    grdPatient.MaxRows = grdPatient.MaxRows + 1
                    grdPatient.SetRowItemData lngRow, Val(udtInPatSched.IntId)
                    Get_NewInPatient Val(udtInPatSched.IntId)
                    Get_Room_Number Val(udtInPatSched.IntId)
                    udtInPatSched.Status = clsPatient.Status
                    udtInPatSched.Schedule_Date = Format(txtDate.Value, "mm/dd/yyyy")
                    lngCol = 1
                    Format_CellTypeEdit grdPatient, lngRow, lngCol, "1", udtInPatSched.LastName & ", " & udtInPatSched.FirstName & vbNewLine & udtInPatSched.IntId, "TypeHAlignLeft", "PATIENT_ID", 0, CStr(lngPatID)
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 2
                    Format_CellTypeEdit grdPatient, lngRow, lngCol, "1", udtInPatSched.RoomNumber, "TypeHAlignCenter", "ROOM_NUMBER"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 3
                    Get_Edi_Static
                    If clsEDI_Client.Site_ID = "1999" Then
                        strList = "Initial" & Chr(9) & "Follow-Up" & Chr(9) & "Re-Admission"
                        Format_CellTypeCombo grdPatient, lngRow, lngCol, "1", strList, "TypeHAlignLeft", "PURPOSE_VISIT"
                        Set_GridCurSel grdPatient, "Follow-Up", lngCol, lngRow, "1", "PURPOSE_VISIT"
                        grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Else
                        Format_CellTypeCombo grdPatient, lngRow, lngCol, "1", Load_Visit, "TypeHAlignLeft", "PURPOSE_VISIT"
                    End If
                    If clsEDI_Client.Site_ID = "1125" Then
                        Set_GridCurSel grdPatient, "Functionsequent Visit", lngCol, lngRow, "1", "PURPOSE_VISIT"
                        grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 4
                    Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", udtInPatSched.Status, "TypeHAlignCenter", "STATUS"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 5
                    If TextFound("No record for primary insurance", udtInPatSched.ApptComment, Len("No record for primary insurance"), True) Then
                        udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "No record for primary insurance", vbNullString)
                        udtInPatSched.ApptComment = Trim(udtInPatSched.ApptComment)
                    End If
                    Get_Insurance Val(udtInPatSched.IntId)
                    If Trim(clsInsurance.Prim_Insurer) <> vbNullString Then
                        If Not TextFound(clsInsurance.Prim_Insurer, udtInPatSched.ApptComment, Len(clsInsurance.Prim_Insurer), True) Then
                            If udtInPatSched.ApptComment = vbNullString Then
                                udtInPatSched.ApptComment = Trim(clsInsurance.Prim_Insurer)
                            Else
                                udtInPatSched.ApptComment = clsInsurance.Prim_Insurer & vbNewLine & Trim(udtInPatSched.ApptComment)
                            End If
                        End If
                    Else
                        If udtInPatSched.ApptComment = vbNullString Then
                            udtInPatSched.ApptComment = "No record for primary insurance"
                        Else
                            udtInPatSched.ApptComment = "No record for primary insurance" & vbNewLine & Trim(udtInPatSched.ApptComment)
                        End If
                    End If
                    Format_CellTypeEdit grdPatient, lngRow, lngCol, "1", Fix_Comments(udtInPatSched.ApptComment), "TypeHAlignCenter", "NOTES"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
            End If
        Loop
    End With
    lblCurScheAppt.Caption = "Inactive Patient Roster"
    lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_InActive:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_InActive", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Load_Facilities(Optional vlngFacID As Long)
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
Dim lngCurID As Long
Dim lngRow As Long
Dim varCellValue

    On Error GoTo Error_Load_Facilities
    Screen.MousePointer = vbHourglass
    Format_FacilityGrid
    lngCurID = vlngFacID
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "SELECT DISTINCT FACILITIES.FACILITY_ID, FACILITIES.FACILITY_NAME FROM APPOINTMENT_PROVIDERS " _
        & "INNER JOIN FACILITIES ON FACILITIES.FACILITY_ID = APPOINTMENT_PROVIDERS.FACILITY_ID " _
        & "WHERE APPOINTMENT_PROVIDERS.PROVIDER_ID = " & mlngLogonUser & " ORDER BY FACILITIES.FACILITY_NAME"
    Set rsFac = cnFac.Execute(strSQL)
    With rsFac
        Do Until .EOF
            lngFacID = 0
            booDefault = False
            strName = vbNullString
            booSkip = False
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
                lngRow = grdFacility.MaxRows + 1
                grdFacility.MaxRows = lngRow
                grdFacility.SetRowItemData lngRow, lngFacID
                Format_CellTypeEdit grdFacility, lngRow, 1, "1", strName, "TypeHAlignLeft", "FACILITY_NAME"
                grdFacility.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        Loop
    End With
    If lngCurID > 0 Then
        With grdFacility
            For lngRow = 1 To .MaxRows
                If .GetRowItemData(lngRow) = lngCurID Then
                    .ActiveRow = lngRow
                    .GetText 1, lngRow, varCellValue
                    .Col = 1
                    .Row = lngRow
                    .BackColor = grdCurItem
                    cctFacility.Text = varCellValue
                    Exit For
                End If
            Next lngRow
        End With
    End If
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set cnFac = Nothing
        Set rsFac = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Facilities:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Facilities", App.EXEName & "." & "modBilling"
    Exit Function
    Resume
End Function

Private Function Load_Patient_Status(vstrStatus As String)
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Idx%
Dim lngCol As Long
Dim lngRow As Long
Dim strName As String
Dim lngPatID As Long
Dim strDate As String
Dim strList As String
Dim lngFoundID As Boolean
Dim lngProvID As Long
Dim strNotes As String
Dim bPatFound As Boolean
Dim lngApptID As Long
Dim varCellValue
Dim lngTest As Long
Dim lngTestPatient As Long
Dim booSkip As Boolean
Dim strStart As String
Dim strEnd As String
Dim strReAdmit As String
Dim lngDup As Long
Dim strStatus As String
Dim strVisitType As String
Dim strSchedType As String
Dim strPatType As String
Dim strReason As String
Dim Arr


    On Error GoTo Error_Load_Patient_Status
    Screen.MousePointer = vbHourglass
    mlngPatCnt = 0
    Get_InPatientFormat
    clsShedGrid.ApptDate = udtInPatSched.Appt_Date
    strEnd = DateAdd("d", 1, Now)
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    If UCase(vstrStatus) = "WORKUP COMPLETED" Then
        strSQL = "SELECT INPATIENT_APPOINTMENTS.*, PATIENT.LAST_NAME, PATIENT.FIRST_NAME FROM INPATIENT_APPOINTMENTS  " _
            & "INNER JOIN PATIENT ON INPATIENT_APPOINTMENTS.PATIENT_ID = PATIENT.PATIENT_ID " _
            & "WHERE  INPATIENT_APPOINTMENTS.PROVIDER_ID = " & udtInPatSched.Provider_ID & " AND " _
            & "CONVERT(DATETIME, INPATIENT_APPOINTMENTS.APPT_DATE,101)  BETWEEN '" & Format(mstrRangeDate, "mm/dd/yyyy") & "' AND '" & Format(strEnd, "mm/dd/yyyy") & "' AND " _
            & "INPATIENT_APPOINTMENTS.STATUS = '" & vstrStatus & "' AND " _
            & "INPATIENT_APPOINTMENTS.FACILITY_ID = " & udtInPatSched.Facility_ID & " AND INPATIENT_APPOINTMENTS.ACTIVE = 0 " _
            & "ORDER BY PATIENT.LAST_NAME, PATIENT.FIRST_NAME, INPATIENT_APPOINTMENTS.INSERT_DATE DESC"
    Else
        strSQL = "SELECT INPATIENT_APPOINTMENTS.* FROM INPATIENT_APPOINTMENTS  " _
            & "INNER JOIN PATIENT ON INPATIENT_APPOINTMENTS.PATIENT_ID = PATIENT.PATIENT_ID " _
            & "WHERE  INPATIENT_APPOINTMENTS.PROVIDER_ID = " & udtInPatSched.Provider_ID & " AND " _
            & "INPATIENT_APPOINTMENTS.STATUS = '" & vstrStatus & "' AND " _
            & "INPATIENT_APPOINTMENTS.FACILITY_ID = " & udtInPatSched.Facility_ID & " AND " _
            & "CONVERT(DATETIME, INPATIENT_APPOINTMENTS.APPT_DATE,101)  BETWEEN '" & Format(mstrRangeDate, "mm/dd/yyyy") & "' AND  '" _
            & Format(strEnd, "mm/dd/yyyy") & "'  AND (INPATIENT_APPOINTMENTS.STATUS <> 'PRESENT' OR INPATIENT_APPOINTMENTS.STATUS <> 'MD COMPLETED') ORDER BY PATIENT.LAST_NAME, PATIENT.FIRST_NAME, INPATIENT_APPOINTMENTS.INSERT_DATE DESC"
    End If
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            booSkip = False
            strStatus = vbNullString
            strVisitType = vbNullString
            strPatType = vbNullString
            strSchedType = vbNullString
            strReason = vbNullString
            Clear_InSched
            Clear_Appt
            lngPatID = 0
            strReAdmit = vbNullString
            udtInPatSched.SchedType = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "SUFFIX"
                                udtInPatSched.Suffix = Trim(Fld.Value)
                            Case "LAST_NAME"
                                udtInPatSched.LastName = IsProperCase(Fld.Value)
                            Case "MIDDLE_NAME"
                                udtInPatSched.MiddleInitial = IsProperCase(Fld.Value)
                            Case "FIRST_NAME"
                                udtInPatSched.FirstName = Fld.Value
                            Case "PREF_CONTACT"
                                udtInPatSched.RoomTelephone = Trim(Fld.Value)
                            Case "PATIENT_ID"
                                udtInPatSched.IntId = Fld.Value
                            Case "PURPOSE_VISIT"
                                'udtInPatSched.VisitType = Fld.Value
                                strVisitType = Fld.Value
                            Case "APPT_TYPE"
                               ' udtInPatSched.SchedType = Trim(Fld.Value)
                                strSchedType = Trim(Fld.Value)
                            Case "STATUS"
                                strStatus = Trim(Fld.Value)
                            Case "NOTES"
                                udtInPatSched.ApptComment = Fld.Value
                                udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "''", "'")
                            Case "REASON"
                                strReason = Fld.Value
                                strReason = Replace(strReason, "''", "'")
                            Case "PROVIDER_ID"
                                udtInPatSched.Provider_ID = Fld.Value
                            Case "APPT_DATE"
                                udtInPatSched.Appt_Date = Fld.Value
                            Case "FACILITY_ID"
                                udtInPatSched.Facility_ID = Fld.Value
                            Case "RESOURCE_TYPE"
                                udtInPatSched.ResourceType = Fld.Value
                            Case "ROOM_NUMBER"
                                udtInPatSched.RoomNumber = Fld.Value
                            Case "TELEPHONE"
                                udtInPatSched.RoomTelephone = Fld.Value
                            Case "DURATION"
                                udtInPatSched.Duration = Fld.Value
                            Case "CURRENT_LOCATION"
                                strReAdmit = Fld.Value
                            Case "PATIENT_TYPE"
                                strPatType = Fld.Value
                               ' udtInPatSched.PatientType = Fld.Value
                            Case "INPATIENT_APPT_ID"
                                lngApptID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If lngDup <> Val(udtInPatSched.IntId) Then
                 If Not Validate_CurrentFacility(Val(udtInPatSched.IntId)) Then
                    udtInPatSched.Appointment_ID = lngApptID
                    Get_NewInPatient Val(udtInPatSched.IntId)
                    udtInPatSched.Status = strStatus
                    udtInPatSched.VisitType = strVisitType
                    udtInPatSched.SchedType = strSchedType
                    udtInPatSched.PatientType = strPatType
                    mlngPatCnt = mlngPatCnt + 1
                    udtInPatSched.Title = Fix_Title(udtInPatSched.Title)
                    udtInPatSched.Suffix = Fix_Suffix(udtInPatSched.Suffix)
                    udtInPatSched.LastName = Fix_Name(udtInPatSched.LastName)
                    udtInPatSched.FirstName = Fix_Name(udtInPatSched.FirstName)
                    If udtInPatSched.FirstName <> vbNullString Then
                        udtInPatSched.Collection = Trim(udtInPatSched.LastName) & ", " & Trim(udtInPatSched.FirstName)
                    End If
                    If Trim(udtInPatSched.Suffix) <> vbNullString Then
                        udtInPatSched.Name = udtInPatSched.FirstName & " " & udtInPatSched.LastName & ", " & udtInPatSched.Suffix
                    Else
                        udtInPatSched.Name = udtInPatSched.FirstName & " " & udtInPatSched.LastName
                    End If
                    If udtInPatSched.PatientType <> vbNullString Then
                        If strReAdmit = "Re-Admission" Then
                            udtInPatSched.PatientType = "Re-Admission"
                        ElseIf udtInPatSched.PatientType = "Subsequent Care" Then
                            udtInPatSched.PatientType = "Follow-Up"
                        ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                            udtInPatSched.PatientType = "Initial"
                        End If
                    End If
                    Get_Room_Number Val(udtInPatSched.IntId)
                    udtInPatSched.Name = Replace(udtInPatSched.Name, "''", "'")
                    If Trim(strReason) <> vbNullString Then
                        udtInPatSched.ApptComment = Trim(udtInPatSched.ApptComment) & vbNewLine & Trim(strReason)
                    End If
                    Arr = Split(udtInPatSched.ApptComment, vbNewLine)
                    If UBound(Arr) > 0 Then
                        udtInPatSched.ApptComment = vbNullString
                        For Idx = LBound(Arr) To UBound(Arr)
                            If Arr(Idx) <> vbNullString And Len(Arr(Idx)) > 2 Then
                                If Trim(udtInPatSched.ApptComment) = vbNullString Then
                                    udtInPatSched.ApptComment = Trim(Arr(Idx))
                                Else
                                    udtInPatSched.ApptComment = udtInPatSched.ApptComment & vbNewLine & Trim(Arr(Idx))
                                End If
                            End If
                        Next Idx
                    End If
                    If grdData.MaxRows > 0 Then
                        grdData.MaxRows = grdData.MaxRows + 1
                        clsShedGrid.Row = grdData.MaxRows
                        Add_Row clsShedGrid.Row
                        lngFoundID = True
                    Else
                        grdData.MaxRows = grdData.MaxRows + 1
                        clsShedGrid.Row = grdData.MaxRows
                        Add_Row clsShedGrid.Row
                    End If
                End If
            End If
            lngDup = Val(udtInPatSched.IntId)
        Loop
    End With
    lblCurScheAppt.Caption = vstrStatus & " " & "Appointment Roster"
    lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
    Format_CellTypeEdit grdData, SpreadHeader + 2, 1, ".maxcols", UCase(vstrStatus) & " " & "VISIT ROSTER" & vbNewLine & _
    "FOR PERIOD BETWEEN" & " " & Format(mstrRangeDate, "mm/dd/yyyy") & " " & "AND" & " " & Format(Now, "mm/dd/yyyy"), "TypeHAlignCenter", "ROOM_NUMBER", clsGridFormat.BackColor
    grdData.SetCellBorder 1, SpreadHeader + 2, grdData.MaxCols, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Patient_Status:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Patient_Status", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Load_Resource_Providers()
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

    On Error GoTo Error_Load_Resource_Providers
    Screen.MousePointer = vbHourglass
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    cboExtender.Clear
    strSQL = "Load_Extender_Providers_Facility " & udtInPatSched.Facility_ID
     Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        Do Until .EOF
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            lngProvID = 0
            booSkip = False
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
                For i = 0 To cboExtender.listcount - 1
                    If TextFound(clsProvider.Name, cboExtender.List(i), Len(clsProvider.Name), True) Then
                        booSkip = True
                        Exit For
                    End If
                Next i
                If Not booSkip Then
                    mbooByPass = True
                    cboExtender.AddItem clsProvider.Name
                    mbooByPass = True
                    If IsNumeric(lngProvID) Then cboExtender.ItemData(cboExtender.NewIndex) = lngProvID
                End If
            End If
        Loop
    End With
    mbooByPass = True
    For i = 0 To cboExtender.listcount - 1
        If cboExtender.ItemData(i) = mlngLogonUser Then
            mbooByPass = True
            cboExtender.Text = cboExtender.List(i)
            cboExtender.Locked = True
            cboApptType.Visible = False
            lblApptType.Visible = False
            Exit For
        End If
    Next i
    If cnMedProv.State = adStateOpen Then
        cnMedProv.Close
        Set cnMedProv = Nothing
        Set rsMedProv = Nothing
        Set Flds = Nothing
    End If
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Resource_Providers:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Resource_Providers", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Search_Add_Patient()
Dim strSQL As String
Dim strSearch As String
Dim Cnt As Long
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngPatID As Long
Dim strTitle As String
Dim lngInvID As Long
Dim strDupNo As String
Dim lngRow As Long
Dim lngCol As Long
Dim lngAcctID As Long
Dim strFirst As String
Dim Arr
Dim lngFac As Long
Dim strList As String
Dim lngCurFac As Long

    On Error GoTo Error_Search_Add_Patient
    Screen.MousePointer = vbHourglass
    lngCurFac = udtInPatSched.Facility_ID
    Clear_UdtPatient
    If cctAddPatient.Text = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    grdAppt.ClearRange 1, 1, grdAppt.MaxCols, grdAppt.MaxRows, False
    grdAppt.MaxRows = 0
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    If Trim(cctAddPatient.Text) = vbNullString Then
        MsgBox "Please enter a search term first!", vbInformation, "PatientTrac Patient Search"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    strList = "Follow-Up" & Chr(9) & "Initial" & Chr(9) & "Re-Admission"
    strSearch = cctAddPatient.Text
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
    strFirst = Replace(strFirst, "'", "''")
    If Trim(strFirst) <> vbNullString Then
        strSQL = "SELECT * FROM PATIENT WHERE LAST_NAME LIKE '" & strSearch & "%' AND FIRST_NAME like '%" & strFirst & "%' ORDER BY LAST_NAME, FIRST_NAME"
    Else
        strSQL = "SELECT * FROM PATIENT WHERE LAST_NAME LIKE '" & strSearch & "%' ORDER BY LAST_NAME, FIRST_NAME"
    End If
    If strSQL <> vbNullString Then
        Set rsPat = cnPat.Execute(strSQL)
        Do Until rsPat.EOF
            lngPatID = 0
            lngPatID = rsPat.Fields("PATIENT_ID")
            rsPat.MoveNext
            ' validate
            If lngAcctID <> lngPatID Then
                Get_PatientName lngPatID
                If Trim(udtPatientName.Suffix) <> vbNullString Then
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                Else
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                End If
                If Trim(udtPatientName.Name) <> vbNullString Then
                    mbooByPass = True
                    Clear_Surg_Facility
                    lngFac = clsFacStatus.FacilityID
                    Read_Patient_Status lngPatID
                    If clsFacStatus.FacilityID = 0 Then
                        clsSurgFac.Name = "No Facility"
                    Else
                        Get_SurgFac "", clsFacStatus.FacilityID
                    End If
                    lngRow = grdAppt.MaxRows + 1
                    grdAppt.MaxRows = lngRow
                    grdAppt.SetRowItemData lngRow, lngPatID
                    lngCol = 1
                    Format_CellTypeStatic grdAppt, lngRow, lngCol, "1", udtPatientName.Name, "TypeHAlignLeft", "PATIENT_ID"
                    grdAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 2
                    Format_CellTypeStatic grdAppt, lngRow, lngCol, "1", udtPatientName.CurrentStatus, "TypeHAlignCenter", "PATIENT STATUS"
                    grdAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 3
                    Format_CellTypeEdit grdAppt, lngRow, lngCol, "1", clsSurgFac.Name, "TypeHAlignCenter", "FACILITY"
                    grdAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    clsFacStatus.FacilityID = lngFac
                    lngCol = 4
                    udtInPatSched.RoomNumber = vbNullString
                    Get_Room_Number lngPatID
                    Format_CellTypeStatic grdAppt, lngRow, lngCol, "1", udtInPatSched.RoomNumber, "TypeHAlignCenter", "ROOM"
                    grdAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 5
                    Format_CellTypeCombo grdAppt, lngRow, lngCol, "1", strList, "TypeHAlignCenter", "VISIT STATUS"
                    grdAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
            End If
            lngAcctID = lngPatID
        Loop
        mbooByPass = True
    End If
    Clear_UdtPatient
    mbooByPass = False
    If cnPat.State = adStateOpen Then
        cnPat.Close
        Set cnPat = Nothing
        Set rsPat = Nothing
    End If
    udtInPatSched.Facility_ID = lngCurFac
    Screen.MousePointer = vbDefault
    Exit Function
Error_Search_Add_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Search_Add_Patient", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Validate_Provider_Permission(vlngPatientID As Long) As Boolean
Dim strSQL As String
Dim strEnctrStart As String
Dim cnPatient As New ADODB.Connection
Dim rsPatient As New ADODB.Recordset
Dim lngPatientID As Long
Dim strTime As String
Dim strStart As String
Dim strEnd As String
Dim lngFacility As Long
Dim lngCurrent As Long
Dim Idx%

    On Error GoTo Error_Validate_Provider_Permission
    Screen.MousePointer = vbHourglass
    lngPatientID = vlngPatientID
    If lngPatientID > 0 Then
        If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
        strSQL = "SELECT STATUS_ITEM FROM PATIENT_STATUS " _
            & "INNER JOIN APPOINTMENT_PROVIDERS ON PATIENT_STATUS.STATUS_ITEM = APPOINTMENT_PROVIDERS.FACILITY_ID " _
            & "WHERE PATIENT_ID = " & lngPatientID & " AND APPOINTMENT_PROVIDERS.PROVIDER_ID = " & mlngLogonUser
            Set rsPatient = cnPatient.Execute(strSQL)
            With rsPatient
                If Not .EOF Then
                    If Not IsNull(!STATUS_ITEM) Then
                        If IsNumeric((!STATUS_ITEM)) Then
                            Validate_Provider_Permission = True
                        End If
                    End If
                End If
            End With
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_Provider_Permission:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_Provider_Permission", App.EXEName & "." & "modBilling"
    Exit Function
    Resume
End Function

Private Function Load_SchedPatient(pCol As Long, pRow As Long)
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngApptID As Long
Dim i%


    On Error GoTo Error_Load_SchedPatient
    Screen.MousePointer = vbHourglass
    cboSearchFilter.Clear
    cboSearchFilter.AddItem "LAST NAME"
    cboSearchFilter.AddItem "EMR ID"
    cboSearchFilter.AddItem "HOSPITAL RECORD"
    cboSearchFilter.AddItem "SOCIAL SECURITY"
    For i = 0 To cboSearchFilter.listcount - 1
        If cboSearchFilter.List(i) = "LAST NAME" Then
            cboSearchFilter.Text = cboSearchFilter.List(i)
            Exit For
        End If
    Next i
    grdPatient.MaxRows = 0
    Format_Patient
    mbooByPass = True
    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        udtInPatSched.Facility_ID = .GetRowItemData(.ActiveRow)
    End With
    mbooByPass = True
    For i = 0 To cboProviders.listcount - 1
        mbooByPass = True
        If cboProviders.List(i) = udtInPatSched.Provider Then
            mbooByPass = True
            udtInPatSched.Provider_ID = cboProviders.ItemData(i)
            Exit For
        End If
    Next i
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Load_Schedule_Patient_PatientID '" & udtInPatSched.Appt_Date & "'," & udtInPatSched.Facility_ID
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        If Not .EOF Then
            Clear_InSched
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PATIENT_ID"
                                udtInPatSched.IntId = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtInPatSched.VisitType = Fld.Value
'                            Case "APPT_TYPE"
'                                udtInPatSched.SchedType = Trim(Fld.Value)
                            Case "STATUS"
                                udtInPatSched.Status = Trim(Fld.Value)
                            Case "PROVIDER_ID"
                                udtInPatSched.Provider_ID = Fld.Value
                            Case "FACILITY_ID"
                                udtInPatSched.Facility_ID = Fld.Value
                            Case "RESOURCE_TYPE"
                                udtInPatSched.ResourceType = Fld.Value
                            Case "ROOM_NUMBER"
                                udtInPatSched.RoomNumber = Fld.Value
                            Case "TELEPHONE"
                                udtInPatSched.RoomTelephone = Fld.Value
                            Case "DURATION"
                                udtInPatSched.Duration = Fld.Value
                            Case "PATIENT_TYPE"
                                udtInPatSched.PatientType = Fld.Value
                            Case "STATUS_NOTES"
                                udtInPatSched.ApptComment = Fld.Value
                                udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "''", "'")
                            Case "INPATIENT_APPT_ID"
                                udtInPatSched.Appointment_ID = Fld.Value
                        End Select
                    End If
                End If
            Next
            If Val(udtInPatSched.IntId) > 0 Then
                grdPatient.SetRowItemData lngRow, udtInPatSched.Appointment_ID
                lngRow = grdPatient.MaxRows + 1
                grdPatient.MaxRows = grdPatient.MaxRows + 1
                udtInPatSched.Appt_Date = Format(txtDate.Value, "mm/dd/yyyy")
                udtInPatSched.Provider = cboProviders.Text
                udtInPatSched.Facility = cctFacility.Text
                Get_InPatientSchedule Val(udtInPatSched.IntId)
                lngCol = 1
                Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", udtInPatSched.LastName & ", " & udtInPatSched.FirstName, "TypeHAlignLeft", "PATIENT_ID", "TypeVAlignCenter", 0, udtInPatSched.IntId
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", udtInPatSched.Appt_Date, "TypeHAlignCenter", "APPT_DATE"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 3
                Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", udtInPatSched.RoomNumber, "TypeHAlignLeft", "ROOM_NUMBER"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 4
                Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", Load_Visit, "TypeHAlignLeft", "PURPOSE_VISIT"
                Set_GridCurSel grdPatient, udtInPatSched.VisitType, lngCol, lngRow, "1", "PURPOSE_VISIT"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 5
                Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", Get_Grid_Status, "TypeHAlignCenter", "STATUS"
                Set_GridCurSel grdPatient, udtInPatSched.Status, lngCol, lngRow, "1", "STATUS"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 6
                Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", Fix_Comments(udtInPatSched.ApptComment), "TypeHAlignCenter", "NOTES"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        End If
    End With
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    fraPatient.Visible = True
    fraPatient.ZOrder 0
    Screen.MousePointer = vbDefault
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_SchedPatient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_SchedPatient", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function LookUp_RoomNumber()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset

    On Error GoTo Error_LookUp_RoomNumber
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    strSQL = "SELECT ROOM FROM PATIENT_ADMISSION_ROOMS WHERE " _
        & "PATIENT_ID = " & Val(udtInPatSched.IntId) & " AND FACILITY_ID = " & udtInPatSched.Facility_ID
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        If Not .EOF Then
            udtInPatSched.RoomNumber = (!room)
        End If
    End With
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
     Exit Function
Error_LookUp_RoomNumber:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUp_RoomNumber", App.EXEName & " " & Me.Name
    Exit Function
    Resume Next
End Function

Private Function PT_Support_Search()
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
Dim lngFac As Long

    On Error GoTo Error_PT_Support_Search
    Screen.MousePointer = vbHourglass
    'mbooSupport = True
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    If cctSearch.Text = vbNullString Then
        MsgBox "Please enter a search item first.", vbInformation, "PT_Support Patient Search"
        Exit Function
    End If
    lstListBox.Clear
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
        strSQL = "SELECT * FROM PATIENT WHERE LAST_NAME LIKE '" & strSearch & "%' AND FIRST_NAME like '%" & strFirst & "%'"
    Else
        strSQL = "SELECT * FROM PATIENT WHERE LAST_NAME LIKE '" & strSearch & "%'"
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
                lngFac = clsFacStatus.FacilityID
                Read_Patient_Status lngPatID
                If clsFacStatus.FacilityID = 0 Then
                    clsSurgFac.Name = "No Facility"
                Else
                    Get_SurgFac "", clsFacStatus.FacilityID
                End If
                clsFacStatus.FacilityID = lngFac
                lstListBox.AddItem Trim(udtPatientName.Name) & "-" & clsSurgFac.Name
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
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_PT_Support_Search:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "PT_Support_Search", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Read_Colors(vstrColorName As String, vlngColor As Long)
Dim cnColor As New ADODB.Connection
Dim rsColor As New ADODB.Recordset
Dim strSQL As String
Dim lngPrimKey As Long

    Screen.MousePointer = vbHourglass
    On Error GoTo Error_Read_Colors

    If cnColor.State = adStateClosed Then cnColor.open psConnect(1)
    strSQL = "SELECT * FROM APPT_COLOR WHERE APPT_TYPE = '" & vstrColorName & "' AND APPT_COLOR = " & vlngColor
    Set rsColor = cnColor.Execute(strSQL)
    Do Until rsColor.EOF
         If rsColor.Fields("APPT_TYPE") = "Office" Then udtColor.Office = rsColor.Fields("APPT_COLOR")
         If rsColor.Fields("APPT_TYPE") = "Hospital" Then udtColor.Hospital = rsColor.Fields("APPT_COLOR")
         If rsColor.Fields("APPT_TYPE") = "Vacation" Then udtColor.Vacation = rsColor.Fields("APPT_COLOR")
         If rsColor.Fields("APPT_TYPE") = "Cancelled" Then udtColor.Cancelled = rsColor.Fields("APPT_COLOR")
         If rsColor.Fields("APPT_TYPE") = "CheckOut" Then udtColor.CheckOut = rsColor.Fields("APPT_COLOR")
         If rsColor.Fields("APPT_TYPE") = "NoShow" Then udtColor.NoShow = rsColor.Fields("APPT_COLOR")
         If rsColor.Fields("APPT_TYPE") = "Present" Then udtColor.Present = rsColor.Fields("APPT_COLOR")
         If rsColor.Fields("APPT_TYPE") = "Lunch" Then udtColor.Lunch = rsColor.Fields("APPT_COLOR")
         If rsColor.Fields("APPT_TYPE") = "PostOp" Then udtColor.PostOp = rsColor.Fields("APPT_COLOR")
         If rsColor.Fields("APPT_TYPE") = "PreOp" Then udtColor.PreOp = rsColor.Fields("APPT_COLOR")
        If rsColor.Fields("APPT_TYPE") = "NewPatient" Then udtColor.NewPatient = rsColor.Fields("APPT_COLOR")
        If rsColor.Fields("APPT_TYPE") = "ExistPatient" Then udtColor.ExistPatient = rsColor.Fields("APPT_COLOR")
       If rsColor.Fields("APPT_TYPE") = "Surgery" Then udtColor.Surgery = rsColor.Fields("APPT_COLOR")
       If rsColor.Fields("APPT_TYPE") = "FollowUp" Then udtColor.FollowUp = rsColor.Fields("APPT_COLOR")
        rsColor.MoveNext
    Loop
    If cnColor.State = adStateOpen Then
        cnColor.Close
        Set cnColor = Nothing
        Set rsColor = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Colors:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Colors", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Read_Patient_Appt()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngApptID As Long
Dim i%
Dim strList As String
Dim varCellValue

    On Error GoTo Error_Read_Patient_Appt
    Screen.MousePointer = vbHourglass
    udtInPatSched.Appt_Date = Format(txtDate.Value, "mm/dd/yyyy")
    mbooByPass = True
    For i = 0 To cboProviders.listcount - 1
        mbooByPass = True
        If cboProviders.List(i) = cboProviders.Text Then
            mbooByPass = True
            udtInPatSched.Provider_ID = cboProviders.ItemData(i)
            udtInPatSched.Provider = cboProviders.List(i)
            Exit For
        End If
    Next i
    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        .GetText 1, .Row, varCellValue
        udtInPatSched.Facility_ID = .GetRowItemData(.ActiveRow)
        udtInPatSched.Facility = varCellValue
    End With
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Get_Appointment_Data_Facility " & udtInPatSched.Facility_ID
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            If Not IsNull((!INSERT_DATE)) Then
                If Trim((!INSERT_DATE)) <> vbNullString Then
                    udtInPatSched.Appt_Date = Format((!INSERT_DATE), "mm/dd/yyyy")
                End If
            End If
            .MoveNext
        Loop
    End With
    strSQL = "Get_InPatient_Appointment_Patient " & udtInPatSched.Facility_ID & "," _
        & udtInPatSched.Provider_ID & "," & piPatientID & ",'" & Format(udtInPatSched.Appt_Date, "mm/dd/yyyy") & "'"
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        If Not .EOF Then
            Clear_InSched
            Clear_Appt
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "SUFFIX"
                                udtInPatSched.Suffix = Trim(Fld.Value)
                            Case "LAST_NAME"
                                udtInPatSched.LastName = IsProperCase(Fld.Value)
                            Case "MIDDLE_NAME"
                                udtInPatSched.MiddleInitial = IsProperCase(Fld.Value)
                            Case "FIRST_NAME"
                                udtInPatSched.FirstName = Fld.Value
                            Case "PREF_CONTACT"
                                udtInPatSched.RoomTelephone = Trim(Fld.Value)
                            Case "PATIENT_ID"
                                udtInPatSched.IntId = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtInPatSched.VisitType = Fld.Value
                            Case "STATUS"
                                udtInPatSched.Status = Trim(Fld.Value)
                            Case "PROVIDER_ID"
                                udtInPatSched.Provider_ID = Fld.Value
                            Case "FACILITY_ID"
                                udtInPatSched.Facility_ID = Fld.Value
                            Case "RESOURCE_TYPE"
                                udtInPatSched.ResourceType = Fld.Value
                            Case "ROOM"
                                udtInPatSched.RoomNumber = Fld.Value
'                            Case "TELEPHONE"
'                                udtInPatSched.RoomTelephone = Fld.Value
                            Case "DURATION"
                                udtInPatSched.Duration = Fld.Value
                            Case "PATIENT_TYPE"
                                udtInPatSched.PatientType = Fld.Value
                            Case "STATUS_NOTES"
                                udtInPatSched.ApptComment = Fld.Value
                                udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "''", "'")
                            Case "INPATIENT_APPT_ID"
                                udtInPatSched.Appointment_ID = Fld.Value
                        End Select
                    End If
                End If
            Next
        
            If Val(udtInPatSched.IntId) > 0 Then
                Get_Room_Number Val(udtInPatSched.IntId)
                lngRow = grdPatient.MaxRows + 1
                grdPatient.MaxRows = grdPatient.MaxRows + 1
                grdPatient.SetRowItemData lngRow, udtInPatSched.Appointment_ID
                udtInPatSched.Appt_Date = Format(txtDate.Value, "mm/dd/yyyy")
                udtInPatSched.Provider = cboProviders.Text
                lngCol = 1
                Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", udtInPatSched.LastName & ", " & udtInPatSched.FirstName, "TypeHAlignLeft", "PATIENT_ID", "TypeVAlignCenter", 0, udtInPatSched.IntId
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", udtInPatSched.Appt_Date, "TypeHAlignCenter", "APPT_DATE"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 3
                Format_CellTypeEdit grdPatient, lngRow, lngCol, "1", udtInPatSched.RoomNumber, "TypeHAlignLeft", "ROOM_NUMBER"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 4
                If clsEDI_Client.Site_ID = "1999" Then
                    strList = "Psychiatric Diagnostic Exam" & Chr(9) & "Subsequent E&M Care, Brief"
                    Format_CellTypeCombo grdPatient, lngRow, lngCol, "1", strList, "TypeHAlignLeft", "PURPOSE_VISIT"
                    Set_GridCurSel grdPatient, "Subsequent Visit", lngCol, lngRow, "1", "PURPOSE_VISIT"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Else
                    Correlate_ApptVisit udtInPatSched.VisitType
                    Format_CellTypeCombo grdPatient, lngRow, lngCol, "1", Load_Visit, "TypeHAlignLeft", "PURPOSE_VISIT"
                    Set_GridCurSel grdPatient, Trim(clsCptTemp.CPT_Code), lngCol, lngRow, "1", "PURPOSE_VISIT"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
                lngCol = 5
                Format_CellTypeCombo grdPatient, lngRow, lngCol, "1", Get_Grid_Status, "TypeHAlignCenter", "STATUS"
                Set_GridCurSel grdPatient, udtInPatSched.Status, lngCol, lngRow, "1", "STATUS"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 6
                'add insurance data
                If TextFound("No record for primary insurance", udtInPatSched.ApptComment, Len("No record for primary insurance"), True) Then
                    udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "No record for primary insurance", vbNullString)
                    udtInPatSched.ApptComment = Trim(udtInPatSched.ApptComment)
                End If
                Get_Insurance Val(udtInPatSched.IntId)
                If Trim(clsInsurance.Prim_Insurer) <> vbNullString Then
                    If Not TextFound(clsInsurance.Prim_Insurer, udtInPatSched.ApptComment, Len(clsInsurance.Prim_Insurer), True) Then
                        If udtInPatSched.ApptComment = vbNullString Then
                            udtInPatSched.ApptComment = Trim(clsInsurance.Prim_Insurer)
                        Else
                            udtInPatSched.ApptComment = clsInsurance.Prim_Insurer & vbNewLine & Trim(udtInPatSched.ApptComment)
                        End If
                    End If
                Else
                    If udtInPatSched.ApptComment = vbNullString Then
                        udtInPatSched.ApptComment = "No record for primary insurance"
                    Else
                        udtInPatSched.ApptComment = "No record for primary insurance" & vbNewLine & Trim(udtInPatSched.ApptComment)
                    End If
                End If
                    Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", Fix_Comments(udtInPatSched.ApptComment), "TypeHAlignCenter", "NOTES"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
        End If
    End With
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    mbooByPass = False
    Clear_InSched
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Patient_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Read_Patient_Appt", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Read_Secondary()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Arr
Dim lngPatID As Long
Dim Idx As Integer
Dim lngRow As Long
Dim varCellValue
Dim vTime
Dim i%
Dim lngCol As Long
Dim lngFoundID As Boolean
Dim lngProvID As Long
Dim strNotes As String
Dim bPatFound As Boolean
Dim lngApptID As Long


    On Error GoTo Error_Read_Secondary
    Screen.MousePointer = vbHourglass
    mlngPatCnt = 0
    rtfChanges.Text = vbNullString
    grdData.MaxRows = 0
    mbooByPass = True
    For i = 0 To cboProviders.listcount - 1
        mbooByPass = True
        If cboProviders.List(i) = cboProviders.Text Then
            mbooByPass = True
            udtInPatSched.Provider_ID = cboProviders.ItemData(i)
            udtInPatSched.Provider = cboProviders.List(i)
            Exit For
        End If
    Next i
    mbooByPass = True

    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        .GetText 1, .Row, varCellValue
        udtInPatSched.Facility_ID = .GetRowItemData(.ActiveRow)
        udtInPatSched.Facility = varCellValue
    End With
    If mstrApptMode = vbNullString Then mstrApptMode = "Day"
'    If udtInPatSched.Appt_Date = vbNullString Then
'        udtInPatSched.Appt_Date = Format(Now, "mm/dd/yyyy")
'    Else
'        udtInPatSched.Appt_Date = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
'    End If
'    clsShedGrid.ApptDate = udtInPatSched.Appt_Date
'    clsShedGrid.ApptDate = Format(clsShedGrid.ApptDate, "mm/dd/yyyy")
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Get_Appointment_Data_Facility " & udtInPatSched.Facility_ID
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            If Not IsNull((!INSERT_DATE)) Then
                If Trim((!INSERT_DATE)) <> vbNullString Then
                    udtInPatSched.Appt_Date = Format((!INSERT_DATE), "mm/dd/yyyy")
                End If
            End If
            .MoveNext
        Loop
    End With
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Get_InPatient_Appointments_Status '" & Format(udtInPatSched.Appt_Date, "mm/dd/yyyy") & "'," _
        & udtInPatSched.Facility_ID & "," & udtInPatSched.Provider_ID & ",'Postpone'"
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            Clear_InSched
            Clear_Appt
            lngPatID = 0
            udtInPatSched.SchedType = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "SUFFIX"
                                udtInPatSched.Suffix = Trim(Fld.Value)
                            Case "LAST_NAME"
                                udtInPatSched.LastName = IsProperCase(Fld.Value)
                            Case "MIDDLE_NAME"
                                udtInPatSched.MiddleInitial = IsProperCase(Fld.Value)
                            Case "FIRST_NAME"
                                udtInPatSched.FirstName = Fld.Value
                            Case "PREF_CONTACT"
                                udtInPatSched.RoomTelephone = Trim(Fld.Value)
                            Case "PATIENT_ID"
                                udtInPatSched.IntId = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtInPatSched.VisitType = Fld.Value
                            Case "APPT_TYPE"
                                udtInPatSched.SchedType = Trim(Fld.Value)
                            Case "STATUS"
                                udtInPatSched.Status = Trim(Fld.Value)
                            Case "PROVIDER_ID"
                                udtInPatSched.Provider_ID = Fld.Value
                            Case "APPT_DATE"
                                udtInPatSched.Appt_Date = Fld.Value
                            Case "FACILITY_ID"
                                udtInPatSched.Facility_ID = Fld.Value
                            Case "RESOURCE_TYPE"
                                udtInPatSched.ResourceType = Fld.Value
                            Case "ROOM_NUMBER"
                                udtInPatSched.RoomNumber = Fld.Value
                            Case "TELEPHONE"
                                udtInPatSched.RoomTelephone = Fld.Value
                            Case "DURATION"
                                udtInPatSched.Duration = Fld.Value
                            Case "PATIENT_TYPE"
                                udtInPatSched.PatientType = Fld.Value
                            Case "STATUS_NOTES"
                                udtInPatSched.ApptComment = Fld.Value
                                udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "''", "'")
                            Case "INPATIENT_APPT_ID"
                                lngApptID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            lngFoundID = False
            udtInPatSched.Title = Fix_Title(udtInPatSched.Title)
            udtInPatSched.Suffix = Fix_Suffix(udtInPatSched.Suffix)
            udtInPatSched.LastName = Fix_Name(udtInPatSched.LastName)
            udtInPatSched.FirstName = Fix_Name(udtInPatSched.FirstName)
            If udtInPatSched.FirstName <> vbNullString Then
                udtInPatSched.Collection = Trim(udtInPatSched.LastName) & ", " & Trim(udtInPatSched.FirstName)
            End If
            If Trim(udtInPatSched.Suffix) <> vbNullString Then
                udtInPatSched.Name = udtInPatSched.FirstName & " " & udtInPatSched.LastName & ", " & udtInPatSched.Suffix
            Else
                udtInPatSched.Name = udtInPatSched.FirstName & " " & udtInPatSched.LastName
            End If
            mlngPatCnt = mlngPatCnt + 1
            udtInPatSched.Name = Replace(udtInPatSched.Name, "''", "'")
            If grdData.MaxRows > 0 Then
                For lngRow = 1 To grdData.MaxRows
                    grdData.Col = 1
                    grdData.Row = lngRow
                    lngPatID = grdData.GetRowItemData(lngRow)
                    If lngPatID <> Val(udtInPatSched.IntId) Then
                        grdData.MaxRows = grdData.MaxRows + 1
                        clsShedGrid.Row = grdData.MaxRows
                        Add_Row clsShedGrid.Row
                        lngFoundID = True
                        Exit For
                    Else
                       lngFoundID = True
                       Exit For
                    End If
                Next lngRow
            Else
                grdData.MaxRows = grdData.MaxRows + 1
                clsShedGrid.Row = grdData.MaxRows
                Add_Row clsShedGrid.Row
            End If
        Loop
    End With
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    udtInPatSched.Appt_Date = clsShedGrid.ApptDate
    mbooByPass = False
    mFirstLoad = False
    mbooByPass = False
    udtInPatSched.SchedType = "Patient"
    Get_Changes clsShedGrid.ApptDate
    Set_Headers
    Clear_InSched
    udtInPatSched.IntId = vbNullString
    lblCurScheAppt.Visible = True
    lblCurScheAppt.Caption = "Active Visit Roster"
    lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
    mbooByPass = True
    mbooByPass = True
    For i = 0 To cboProviders.listcount - 1
        mbooByPass = True
        If cboProviders.List(i) = udtInPatSched.Provider Then
            mbooByPass = True
            cboProviders.Text = cboProviders.List(i)
            Exit For
        End If
    Next i
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Secondary:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Read_Secondary", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Search_Unknown_Patient_Status()
Dim strSQL As String
Dim strSearch As String
Dim Cnt As Long
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngPatID As Long
Dim strTitle As String
Dim lngInvID As Long
Dim strDupNo As String
Dim lngRow As Long
Dim lngAcctID As Long
Dim strFirst As String
Dim Arr
Dim strFac As String

    On Error GoTo Error_Search_Unknown_Patient_Status
    Screen.MousePointer = vbHourglass
    Clear_UdtPatient
    If cctSearch.Text = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    lstListBox.Clear
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    If Trim(cctSearch.Text) = vbNullString Then
        MsgBox "Please enter a search term first!", vbInformation, "PatientTrac Patient Search"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    strSearch = cctSearch.Text
    If cboSearchFilter.Text = "FINANCIAL ACCT NO" Then
        strTitle = "EXTERNAL ID"
    Else
        strTitle = UCase(cboSearchFilter.Text)
    End If
    If clsEDI_Client.Site_ID = "1999" Then
        Select Case Trim(strTitle)
            Case "LAST NAME"
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
                strFirst = Replace(strFirst, "'", "''")
                If Trim(strFirst) <> vbNullString Then
                     strSQL = "Search_Patient_FirstName_Last_Name_Facility_Status   '" & strSearch & "%', '" & strFirst & "%'"
                Else
                    strSQL = "Search_Patient_Row_Count_Last_Name_Facility_Status '" & strSearch & "%'"
                End If
            Case "EMR ID"
                strSQL = "Search_Patient_EMR_ID_Facility_Status '" & strSearch & "%'"
            Case "SOCIAL SECURITY"
                strSQL = "Search_Patient_Social_Security_Facility_Status '" & strSearch & "%'"
            Case "EXTERNAL ID"
                strSQL = "Search_Patient_External_ID_Facility_Status 25,'" & strSearch & "%'"
            End Select
            mbooUnknown = True
    Else
        Select Case Trim(strTitle)
            Case "LAST NAME"
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
                strFirst = Replace(strFirst, "'", "''")
                If Trim(strFirst) <> vbNullString Then
                     strSQL = "Search_Patient_FirstName_Last_Name   '" & strSearch & "%', '" & strFirst & "%'"
                Else
    '                strSQL = "Search_InPatient_Row_Count_Last_Name 25, '" & strSearch & "%', " & udtInPatSched.Facility_ID
                    strSQL = "Search_Patient_Row_Count_Last_Name 25, '" & strSearch & "%'"
                End If
            Case "EMR ID"
                strSQL = "Search_Patient_EMR_ID '" & strSearch & "%'"
            Case "SOCIAL SECURITY"
                strSQL = "Search_Patient_Social_Security '" & strSearch & "%'"
            Case "EXTERNAL ID"
                strSQL = "Search_Patient_RowCount_External_ID 25,'" & strSearch & "%'"
            Case "INVOICE NO"
                If Len(strSearch) > 3 Then
                    lngInvID = Val(strSearch)
                    strSQL = "Search_Patient_Invoice_Invoice_No '" & strSearch & "%'"
                    Set rsPat = cnPat.Execute(strSQL)
                    If Not rsPat.EOF Then
                        lngPatID = rsPat.Fields("ACCOUNT_NO")
                    Else
                        MsgBox "This is an invalid Invoice Number! Please select valid Invoice.", vbInformation, "PatientTrac EDI Submission"
                        cctSearch.Text = vbNullString
                        Screen.MousePointer = vbDefault
                        If cnPat.State = adStateOpen Then
                            cnPat.Close
                            Set rsPat = Nothing
                            Set cnPat = Nothing
                        End If
                        Exit Function
                    End If
                End If
            End Select
        End If
        If strSQL <> vbNullString Then
            Set rsPat = cnPat.Execute(strSQL)
            Do Until rsPat.EOF
                Clear_UdtPatient
                lngInvID = 0
                Set Flds = rsPat.Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "ACCT_ID"
                                    lngAcctID = Fld.Value
                                Case "INVOICE_NO"
                                    lngInvID = Fld.Value
                                Case "FACILITY_NAME"
                                    strFac = Fld.Value
                                Case "FIRST_NAME"
                                    udtPatientName.FirstName = IsProperCase(Fld.Value)
                                Case "MIDDLE_NAME"
                                    udtPatientName.MiddleName = IsProperCase(Fld.Value)
                                Case "LAST_NAME"
                                    udtPatientName.LastName = IsProperCase(Fld.Value)
                                    If TextFound("''", udtPatientName.LastName, 2, True) Then
                                        udtPatientName.LastName = Replace(udtPatientName.LastName, "''", "'")
                                        udtPatientName.LastName = IsProperCase(udtPatientName.LastName)
                                    Else
                                        udtPatientName.LastName = IsProperCase(udtPatientName.LastName)
                                    End If
                                Case "PREF_CONTACT"
                                    udtPatientName.PrefContact = Fld.Value
                                Case "GENDER"
                                    udtPatientName.Gender = IsProperCase(Fld.Value)
                                    If Left(IsProperCase(udtPatientName.Gender), 1) = "F" Then
                                        udtPatientName.Gender = "Female"
                                        udtPatientName.Person = "She"
                                        udtPatientName.Possessive = "Her"
                                        udtPatientName.PerPosses = "Herself"
                                    ElseIf Left(IsProperCase(udtPatientName.Gender), 1) = "M" Then
                                        udtPatientName.Gender = "Male"
                                        udtPatientName.Person = "He"
                                        udtPatientName.Possessive = "His"
                                        udtPatientName.PerPosses = "Himself"
                                    End If
                                Case "TITLE"
                                    udtPatientName.Title = IsProperCase(Fld.Value)

                                Case "MARITAL_STATUS"
                                    udtPatientName.Marital_Status = IsProperCase(Fld.Value)
                                Case "SUFFIX"
                                    udtPatientName.Suffix = Fld.Value
                                    udtPatientName.Suffix = Replace(udtPatientName.Suffix, ",", "")
                                    If TextFound("Jr", udtPatientName.Suffix, 2, True) Or _
                                        TextFound("Sr", udtPatientName.Suffix, 2, True) Then
                                        udtPatientName.Suffix = IsProperCase(udtPatientName.Suffix)
                                    Else
                                        udtPatientName.Suffix = UCase(udtPatientName.Suffix)
                                    End If
                                    If TextFound("PHD", udtPatientName.Suffix, Len("phd"), True) Or _
                                        TextFound("PHD", udtPatientName.Suffix, 4, True) Then
                                        udtPatientName.Suffix = "Ph.D."
                                    End If
                                Case "PATIENT_EXT_REF"
                                    udtPatientName.ExtID = Fld.Value
                                Case "PATIENT_ID"
                                    udtPatientName.IntId = Fld.Value
                                    lngPatID = udtPatientName.IntId
                                Case "SOCIAL_SECURITY"
                                    udtPatientName.SS = Fld.Value
                                Case "ROOM_NUMBER"
                                    udtPatientName.RoomNumber = Fld.Value
                                Case "COUNTRY"
                                     udtPatientName.Country = Fld.Value
                                Case "CELL_PHONE"
                                     udtPatientName.CellPhone = Fld.Value
                                Case "EMAIL"
                                     udtPatientName.Email = LCase(Fld.Value)
                                Case "COUNTY"
                                     udtPatientName.County = LCase(Fld.Value)
                                Case "PATIENT_TYPE"
                                     udtPatientName.PatientType = LCase(Fld.Value)
                                Case "DRV_LICENSE"
                                     udtPatientName.DL = LCase(Fld.Value)
                                Case "WORK_PHONE"
                                     udtPatientName.WorkPhone = LCase(Fld.Value)
                                Case "LOCAL_INTERNATL_PHONE"
                                     udtPatientName.InterNalPhone = LCase(Fld.Value)
                                Case "PHONE"
                                     udtPatientName.Phone = LCase(Fld.Value)
                                Case "AREA_CODE"
                                    udtPatientName.AreaCode = LCase(Fld.Value)
                                Case "STATUS"
                                     udtPatientName.Status = LCase(Fld.Value)
                            End Select
                        End If
                    End If
                Next
                rsPat.MoveNext
                If Val(strDupNo) <> Val(udtPatientName.IntId) Then
                    If Left(IsProperCase(udtPatientName.Gender), 1) = "M" Then
                        udtPatientName.Title = "Mr."
                    ElseIf Left(IsProperCase(udtPatientName.Gender), 1) = "F" Then
                        If IsProperCase(udtPatientName.Marital_Status) = "Married" Then
                            udtPatientName.Title = "Mrs."
                        Else
                            udtPatientName.Title = "Ms."
                        End If
                    End If
                    If Trim(udtPatientName.Suffix) <> vbNullString Then
                        udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                    Else
                        udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                    End If
                    Select Case Trim(strTitle)
                        Case "EXTERNAL ID"
                            lblEnterInfo.Caption = "Last Name, First Name - External Reference"
                        Case "EMR ID"
                            lblEnterInfo.Caption = "Last Name, First Name - EMR - ID"
                    End Select
                    If Trim(udtPatientName.Name) <> vbNullString And udtPatientName.IntId <> strDupNo Then
                        lstListBox.AddItem udtPatientName.LastName & ", " & udtPatientName.FirstName & " -  " & strFac
                        lstListBox.ItemData(lstListBox.NewIndex) = Val(udtPatientName.IntId)
                    End If
                End If
                strDupNo = Val(udtPatientName.IntId)
            Loop
            udtPatientName.IntId = lngPatID
        End If
    Clear_UdtPatient
    mbooByPass = False
    If cnPat.State = adStateOpen Then
        cnPat.Close
        Set cnPat = Nothing
        Set rsPat = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Search_Unknown_Patient_Status:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Search_Unknown_Patient_Status", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Set_Status()

    udtInPatSched.Reason = Replace(udtInPatSched.Reason, "Reason:", "")
    udtInPatSched.Reason = Trim(udtInPatSched.Reason)
'    If udtInPatSched.Reason = "FollowUp: Cosmetic" Then udtInPatSched.Status = "Post Operative"
'    If udtInPatSched.Reason = "Appt: New Patient Cosmetic" Then udtInPatSched.Status = "Appointment"
'    If udtInPatSched.Reason = "PostOP" Then udtInPatSched.Status = "Consultation"
'    If udtInPatSched.Reason = 4 Then udtInPatSched.Status = "Surgery Scheduled"
'    If udtInPatSched.Reason = 5 Then udtInPatSched.Status = "Surgery Cancelled"
'    If udtInPatSched.Reason = 6 Then udtInPatSched.Status = "Surgery Completed"
'    If udtInPatSched.Reason = 10 Then udtInPatSched.Status = "OnGoing Services"
'    If udtInPatSched.Reason = 11 Then udtInPatSched.Status = "Patient Exit"
'    If udtInPatSched.Reason = 12 Then udtInPatSched.Status = "Revision Pending"
'    If udtInPatSched.Reason = 13 Then udtInPatSched.Status = "Revision Scheduled"
'    If udtInPatSched.Reason = 14 Then udtInPatSched.Status = "Revision Completed"
'    If udtInPatSched.Reason = 15 Then udtInPatSched.Status = "Plan Cancelled"
                
End Function

Private Function Check_Change(vstrType As String)
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim strNow As String
Dim lngDateDiff As Long
Dim strInsert As String

    mbooSchChange = False
    mbooTeleChange = False
    strNow = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If mstrCheckTime <> vbNullString Then
        If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
                If vstrType = "Appointment" Then
                    strSQL = "SELECT INSERT_DATE FROM INPATIENT_APPOINTMENTS ORDER BY INSERT_DATE DESC"
                    Set rsAppt = cnAppt.Execute(strSQL)
                    Do Until rsAppt.EOF
                        If rsAppt.Fields("INSERT_DATE") <> vbNullString Then
                            If Not IsNull(rsAppt.Fields("INSERT_DATE")) Then strInsert = rsAppt.Fields("INSERT_DATE")
                             lngDateDiff = DateDiff("d", mstrCheckTime, strInsert)
                             If lngDateDiff = 0 Then
                                lngDateDiff = DateDiff("h", mstrCheckTime, strInsert)
                                If lngDateDiff >= 0 Then
                                     mbooSchChange = True
                                     Exit Do
                                End If
                            End If
                        End If
                        rsAppt.MoveNext
                    Loop
                ElseIf vstrType = "Telephone" Then
                    strSQL = "SELECT INSERT_DATE FROM PHONE_MESSAGE  WHERE INSERT_DATE > '" & mstrCheckTime & "'"
                    Set rsAppt = cnAppt.Execute(strSQL)
                    Do Until rsAppt.EOF
                        If rsAppt.Fields("INSERT_DATE") <> vbNullString Then
                            lngDateDiff = DateDiff("n", mstrCheckTime, strNow)
                            If lngDateDiff > 20 Then
                                 mbooTeleChange = True
                                 Exit Do
                            End If
                        End If
                        rsAppt.MoveNext
                    Loop
                End If
            End If
    mstrCheckTime = strNow
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Exit Function
End Function
Private Function Check_Duplicates(vstrText As String)
Dim i%
Dim Arr
Dim strRTF As String

    mbooDup = False
    strRTF = Trim(rtfChanges.Text)
    If Trim(rtfChanges.Text) <> vbNullString Then
        Arr = Split(strRTF, vbNewLine)
        If UBound(Arr) > 0 Then
            For i = 0 To UBound(Arr)
                mbooDup = False
                If TextFound(vstrText, Arr(i), Len(Arr(i)), False) Then
                    mbooDup = True
                    Exit For
                End If
            Next i
        ElseIf TextFound(vstrText, strRTF, Len(vstrText), False) Then
            mbooDup = True
        End If
    End If
    Exit Function
End Function
Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)

    With vObjGrid
        vlngRow = vlngRow
        If .MaxRows < vlngRow Then .MaxRows = vlngRow
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = vlngRow
            .Row2 = vlngRow
            .BackColor = 12632256
            .RowHeight(vlngRow) = 3.5
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub
Private Function Load_Visit(Optional vstrType As String) As String
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

    On Error GoTo Error_Load_Visit
    If mbooByPass Then mbooByPass = False: Exit Function
    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        lngFacID = .GetRowItemData(.ActiveRow)
    End With
    Screen.MousePointer = vbHourglass
    If udtInPatSched.Provider_ID > 0 Then Get_Providers udtInPatSched.Provider_ID
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    If Trim(vstrType) <> vbNullString Then
        strSQL = "SELECT DISTINCT VISIT_TYPE, APPT_CODE , PROVIDER_TITLE FROM APPOINTMENT_TYPE WHERE VISIT_TYPE like '" & vstrType & "%'  AND PURGE = 'False'  ORDER BY PROVIDER_TITLE, APPT_CODE"
    Else
        strSQL = "SELECT DISTINCT VISIT_TYPE, APPT_CODE , PROVIDER_TITLE FROM APPOINTMENT_TYPE WHERE PURGE = 'False' ORDER BY PROVIDER_TITLE,APPT_CODE"
    End If
    Set rsCodes = cnCodes.Execute(strSQL)
    Do Until rsCodes.EOF
        strProvider = vbNullString
        If Not IsNull(rsCodes.Fields("PROVIDER_TITLE")) Then
            strProvider = UCase(rsCodes.Fields("PROVIDER_TITLE"))
        End If
'------------------
        If Trim(clsProvider.Title) <> vbNullString Then
            strCompare = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
            strCompare = Replace(strCompare, ".", "")
            strCompare = Trim(strCompare)
        Else
            If udtInPatSched.Provider_ID = 0 Then
                piProviderID = mlngLogonUser
                Get_Providers piProviderID
            Else
                Get_Providers udtInPatSched.Provider_ID
            End If
            strCompare = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
            strCompare = Replace(strCompare, ".", "")
            strCompare = Trim(strCompare)
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
                                    If Load_Visit = vbNullString Then
                                        Load_Visit = Trim(Fld.Value)
                                    Else
                                        Load_Visit = Load_Visit & Chr(9) & Trim(Fld.Value)
                                    End If
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
Error_Load_Visit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Visit", App.EXEName & " " & Me.Name
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
    If Validate_Extender_Provider Then
        Get_Grid_Status = "Active" & Chr(9) & "Cancelled" & Chr(9) & "Postpone" & Chr(9) & "Workup Completed"
    ElseIf Validate_MD_Provider Then
        Get_Grid_Status = "Cancelled" & Chr(9) & "Initial Round" & Chr(9) & "Present" & Chr(9) & "Second Round" & Chr(9) & "Present"
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Grid_Status:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Grid_Status", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Add_Row(Optional vlngRow As Long)
Dim lngCol As Long
Dim lngRow As Long
Dim strList As String

    clsGridFormat.BackColor = 0
    If Trim(udtInPatSched.Status) <> vbNullString Then
        Manage_Colors udtInPatSched.Status, Val(udtInPatSched.IntId)
    End If
    clsGridFormat.BackColor = 0
    Select Case IsProperCase(udtAppointment.VisitType)
        Case "Init Psych", "New Patient Office Visit"
            clsGridFormat.BackColor = grdYellow
    End Select
    If TextFound("Initial", udtAppointment.VisitType, Len("Initial"), True) Then
        clsGridFormat.BackColor = grdYellow
    End If
    Get_Cur_Sched_Enctr
    piPatientID = Val(udtInPatSched.IntId)
    clsPatAcct.ICD_Code = vbNullString
    Values_Psychiatry "DIAGNOSIS"
    If clsDiagnosis.ICD_Code <> vbNullString Then
        clsPatAcct.ICD_Code = clsDiagnosis.ICD_Code
    End If
    
    If clsPatAcct.ICD_Code = vbNullString Then
        Get_NurseDiagnosis
    End If
    clsGridFormat.BackColor = 0
'    Select Case IsProperCase(udtAppointment.SchedType)
'        Case "Physician"
'            clsGridFormat.BackColor = GridNonPatient
'        Case "Vendor"
'            clsGridFormat.BackColor = GridNonPatient
'    End Select
    Select Case IsProperCase(udtInPatSched.Status)
        Case "Present"
            clsGridFormat.BackColor = grdCurItem
'        Case "No Show", "Cancelled"
'            clsGridFormat.BackColor = grdGray
        Case "Check Out"
            clsGridFormat.BackColor = DataColor
        Case "Confirmed"
            clsGridFormat.BackColor = GridConfir
    End Select
    lngRow = vlngRow
    grdData.SetRowItemData lngRow, Val(udtInPatSched.IntId)
    lngCol = 1
    Format_CellTypeEdit grdData, lngRow, lngCol, CStr(1), udtInPatSched.RoomNumber, "TypeHAlignCenter", "ROOM_NUMBER", clsGridFormat.BackColor
    grdData.SetCellBorder lngCol, clsShedGrid.Row, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 2
    Format_CellTypeEdit grdData, lngRow, lngCol, CStr(1), udtInPatSched.LastName & ", " & udtInPatSched.FirstName, "TypeHAlignLeft", CStr(udtInPatSched.IntId), clsGridFormat.BackColor
    grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 3
    Format_CellTypeEdit grdData, lngRow, lngCol, "1", Trim(clsPatAcct.ICD_Code), "TypeHAlignCenter", "ICD_CODE", clsGridFormat.BackColor
    grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 4
    If Validate_Extender_Provider Then
        strList = "Active" & Chr(9) & "Cancelled" & Chr(9) & "Postpone" & Chr(9) & "Workup Completed"
        Format_CellTypeCombo grdData, lngRow, lngCol, "1", Trim(Get_Grid_Status), "TypeHAlignCenter", "STATUS", clsGridFormat.BackColor
        If Trim(udtInPatSched.Status) <> vbNullString Then
            Set_GridCurSel grdData, udtInPatSched.Status, lngCol, lngRow, "1", "STATUS"
        End If
    ElseIf Validate_MD_Provider Then
        strList = "Initial Round" & Chr(9) & "Second Round" & Chr(9) & "Cancelled" & Chr(9) & "Present"
        Format_CellTypeCombo grdData, lngRow, lngCol, "1", Trim(Get_Grid_Status), "TypeHAlignCenter", "STATUS", clsGridFormat.BackColor
        If udtInPatSched.Status = "Active" Or udtInPatSched.Status = "Workup Completed" Then udtInPatSched.Status = "Initial Round"
        If Trim(udtInPatSched.Status) <> vbNullString Then
            Set_GridCurSel grdData, udtInPatSched.Status, lngCol, lngRow, "1", "STATUS"
        End If
    End If
    If Trim(udtInPatSched.Status) <> vbNullString Then
        Set_GridCurSel grdData, udtInPatSched.Status, lngCol, lngRow, "1", "STATUS"
    End If
    grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 5
    strList = "Follow-Up" & Chr(9) & "Initial" & Chr(9) & "Re-Admission"
    Format_CellTypeCombo grdData, lngRow, lngCol, "1", strList, "TypeHAlignCenter", "STATUS", clsGridFormat.BackColor
    If Trim(udtInPatSched.PatientType) <> vbNullString Then
        Set_GridCurSel grdData, udtInPatSched.PatientType, lngCol, lngRow, "1", "STATUS"
    End If
    grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 6
    Get_Insurance Val(udtInPatSched.IntId)
    If Trim(udtInPatSched.ApptComment) <> vbNullString Then
        Format_CellTypeEdit grdData, lngRow, lngCol, "1", Trim(clsInsurance.Prim_Insurer) & vbNewLine & Trim(udtInPatSched.ApptComment), "TypeHAlignLeft", "NOTES", clsGridFormat.BackColor
    Else
        Format_CellTypeEdit grdData, lngRow, lngCol, "1", Trim(clsInsurance.Prim_Insurer), "TypeHAlignLeft", "NOTES", clsGridFormat.BackColor
    End If
    
    grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    grdData.Refresh
   Exit Function
End Function
Private Function Load_Enctr_Format()

    Select Case mstrPractice
        Case "Gyn", "IVF"
            If IsFormLoaded("frmSuperbill") Then
                Unload frmSuperBill
            ElseIf clsEncounter.Intrvn_CD = "99024" Then
                If clsEncounter.EncounterType <> vbNullString Then
                   If Not IsFormLoaded("frmInPatient") Then
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.mstrLoadType = "PostOpVisit"
                        frmInPatient.Read_Data piEncounterID
                        frmInPatient.ZOrder 0
                        frmInPatient.Show vbModal
                    Else
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.Read_Data piEncounterID
                        frmInPatient.ZOrder 0
                        frmInPatient.Show vbModal
                    End If
                End If
                Exit Function
           ElseIf Trim(clsOpNote.Procedure_Name) <> vbNullString And Trim(clsOpNote.Procedure_ID) <> vbNullString Then
                   If Not IsFormLoaded("frmInPatient") Then
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.mstrLoadType = "Surgery"
                        frmInPatient.Read_Data piEncounterID
                        frmInPatient.ZOrder 0
                        frmInPatient.Show vbModal
                      '  Me.Show
                    Else
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.ZOrder 0
                        frmInPatient.Show vbModal
                       ' Me.Show
                    End If
                Exit Function
            Else
                If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.Load_Reports
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show vbModal
                 Else
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show vbModal
                 End If
                Exit Function
            End If
        Case "Psychiatry"
            If IsFormLoaded("frmSuperbill") Then
                Unload frmSuperBill
            Else
                clsEncounter.EncounterType = Trim(clsEncounter.EncounterType)
                If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
                If clsEncounter.EncounterType = vbNullString Then clsEncounter.EncounterType = "Subsequent Patient"
                If clsEncounter.EncounterType <> vbNullString Then
                   If Not IsFormLoaded("frmInPatient") Then
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.Load_Reports
                        frmInPatient.mstrLoadType = "InPatient Appointment"
                        Load frmInPatient
                        frmInPatient.Show vbModal
                    Else
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.mstrLoadType = "InPatient Appointment"
                        Load frmInPatient
                        frmInPatient.Show vbModal
                    End If
                End If
                Exit Function
            End If
    Case "Plastic Surgery"
        If IsFormLoaded("frmSuperbill") Then
        ElseIf clsEncounter.Intrvn_CD = "99024" Then
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    Me.Hide
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "PostOpVisit"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show vbModal
                    Me.Show
                Else
                    Me.Hide
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show vbModal
                    Me.Show
                End If
            End If
           ' Unload Me
            Exit Function
        ElseIf clsEncounter.VisitType = "Ambulatory Procedure Visit" And TextFound("049X", clsEncounter.Intrvn_CD, Len("049X"), True) Then
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    Me.Hide
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "Ambulatory Procedure Visit"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show vbModal
                    Me.Show
                Else
                    Me.Hide
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show vbModal
                    Me.Show
                End If
            End If
           ' Unload Me
            Exit Function
        ElseIf clsEncounter.VisitType = "Pre-Operative Exam Anesthesia" And clsEncounter.Intrvn_CD = "99201" Then
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "PreOperative Anesthesia"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            End If
           ' Unload Me
            Exit Function
        ElseIf clsEncounter.VisitType = "Surgical Admission" And clsEncounter.Intrvn_CD = "90806" Then
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "Surgical Admission"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            End If
           ' Unload Me
            Exit Function
       ElseIf Trim(clsOpNote.Procedure_Name) <> vbNullString And Trim(clsOpNote.Procedure_ID) <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "Surgery"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            'Unload Me
            Exit Function
        Else
            'Create_Enctr_Default
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Load_Reports
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            End If
            Exit Function
        End If
    End Select
    Exit Function
End Function

Private Sub Clear_EnctrString()
    With udtEnctr
        .EnctrID = 0
        .StartDate = vbNullString
        .EndDate = vbNullString
        .VisitReason = vbNullString
        .Facility = vbNullString
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
Private Function Clear_Sched()
    With clsShedGrid
        .CellSpan = vbNullString
        .CellTag = vbNullString
        .CellNote = vbNullString
        .BackColor = 0
        .Text = vbNullString
        .TextNote = vbNullString
        .List = vbNullString
        .Duration = 0
        .InsertRow = 0
        .ProvCol = 0
        .ProvCol2 = 0
        .SpanRow = 0
        .ApptDate = vbNullString
        .StartRow = 0
        .EndRow = 0
        .Col = 0
        .Row = 0
        .RowName = vbNullString
        .ColHeader = vbNullString
        .CellType = vbNullString
    End With
End Function


Private Sub Clear_Block()
    With udtBlock
        .StartDate = vbNullString
        .EndDate = vbNullString
        .StartTime = vbNullString
        .EndTime = vbNullString
        .Reason = vbNullString
        .ProviderID = 0
        .Duration = vbNullString
        .FacilityID = 0
        .DayWeek = vbNullString
    End With
    Exit Sub
End Sub
Private Sub Format_Messages()
Dim strSQL As String
Dim cnMessage As New ADODB.Connection
Dim rsMessage As New ADODB.Recordset
Dim lngCol As Long
Dim strHeader As String
Dim lngCnt As String
Dim sDate As String

    On Error GoTo Error_Format_Messages
    Screen.MousePointer = vbHourglass
    mlngMessRow = 0
    Get_Providers
    grdMessages.MaxRows = mlngMessRow
    With grdMessages
        grdMessages.ClearRange 0, 0, .MaxCols, .MaxRows, False
        .Redraw = False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .FontName = "Georgia"
        .GridSolid = False
        .GridShowHoriz = False
        .GridShowVert = False
        .RowHeadersShow = False
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .RowHeight(0) = 20
    End With
    If cnMessage.State = adStateClosed Then cnMessage.open psConnect(1)
    strSQL = "Select HEADERS, FORM_ID, visible, SEQUENCE_NO from GRID_HEADERS where " _
        & "form_name = 'Telephone Messages' ORDER BY SEQUENCE_NO"
    Set rsMessage = cnMessage.Execute(strSQL)
        lngCnt = 0
        While Not rsMessage.EOF
            lngCnt = lngCnt + 1
            rsMessage.MoveNext
        Wend

    strSQL = "Select HEADERS, FORM_ID, visible, SEQUENCE_NO from GRID_HEADERS where form_name = 'Telephone Messages' ORDER BY SEQUENCE_NO "
    Set rsMessage = cnMessage.Execute(strSQL)
    With rsMessage
        Do Until .EOF
            If rsMessage.Fields("HEADERS") <> vbNullString Then
                lngCol = lngCol + 1
                grdMessages.MaxCols = lngCol
                strHeader = rsMessage.Fields("HEADERS")
                strHeader = Replace(strHeader, "_", "  ")
                grdMessages.Col = lngCol
                grdMessages.ColID = rsMessage.Fields("HEADERS")
                grdMessages.SetColItemData (lngCol), (!FORM_ID)
                Format_CellTypeStatic grdMessages, SpreadHeader + 2, lngCol, "1", strHeader, "TypeHAlignCenter", "", "", HeaderColor
                grdMessages.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                If lngCol > 1 Then
                    grdMessages.ColWidth(lngCol) = 20
                Else
                    grdMessages.ColWidth(lngCol) = 8
                End If
                If (!Visible) = 0 Then grdMessages.ColHidden = True
            End If
        .MoveNext
        Loop
    End With
    grdMessages.ColWidth(6) = 28
    Format_CellTypeStatic grdMessages, SpreadHeader, 1, "maxcols", "TELEPHONE MESSAGES FOR" & " " & clsProvider.Name, "TypeHAlignCenter", "", "", HeaderColor
    grdMessages.SetCellBorder 1, SpreadHeader, grdMessages.MaxCols, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid

    sDate = Format(Now, "dddd, MMMM DD, yyyy")
    Format_CellTypeStatic grdMessages, SpreadHeader + 1, 1, "maxcols", sDate, "TypeHAlignCenter", "", "", HeaderColor
    grdMessages.SetCellBorder 1, SpreadHeader + 1, grdMessages.MaxCols, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    grdMessages.Redraw = True
    If cnMessage.State = adStateOpen Then
        cnMessage.Close
        Set cnMessage = Nothing
        Set rsMessage = Nothing
    End If
    Exit Sub
    Screen.MousePointer = vbDefault
Error_Format_Messages:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_Messages", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub



Private Function Get_InPatientFormat()
Dim Arr
Dim lngCol As Long
Dim lngRow As Long
Dim Idx As Integer
Dim vTime
Dim vdata
Dim varData
Dim varDoubleBook
Dim booNoInsert As Boolean
Dim lngProviderID As Long
Dim lngProvCol As Long
Dim varCellValue
Dim CellNote As Long
Dim lngSpanRow As Long
Dim lngRowTime As Long
Dim intSpan As Integer
Dim strInt As String
Dim cnApptFormat As New ADODB.Connection
Dim rsApptFormat As New ADODB.Recordset
Dim strSQL As String
Dim strMin As String
Dim lngCnt As Long
Dim lngProvID As Long
Dim strSpan As String

    On Error GoTo Error_Get_InPatientFormat
    Screen.MousePointer = vbHourglass
    grdData.ColHeaderRows = 6
    grdData.MaxCols = 6
    grdData.MaxRows = 0
    lngCnt = 0
    grdData.ColWidth(1) = 7
    grdData.ColWidth(2) = 25
    grdData.ColWidth(3) = 15
    grdData.ColWidth(4) = 15
    grdData.ColWidth(5) = 15
    grdData.ColWidth(6) = 39
    lngCol = 1
    Format_CellTypeEdit grdData, SpreadHeader + 4, lngCol, "1", "ROOM", "TypeHAlignCenter", HeaderColor
    grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 2
    'grdData.ColWidth(lngCol) = 25
    Format_CellTypeEdit grdData, SpreadHeader + 4, lngCol, "1", "PATIENT", "TypeHAlignCenter", HeaderColor
    grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 3
    Format_CellTypeEdit grdData, SpreadHeader + 4, lngCol, "1", "DX CODE", "TypeHAlignCenter", HeaderColor
    grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 4
    Format_CellTypeEdit grdData, SpreadHeader + 4, lngCol, "1", "VISIT STATUS", "TypeHAlignCenter", HeaderColor
    grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 5
    Format_CellTypeEdit grdData, SpreadHeader + 4, lngCol, "1", "VISIT TYPE", "TypeHAlignCenter", HeaderColor
    grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 6
    Format_CellTypeEdit grdData, SpreadHeader + 4, lngCol, "1", "COMMENTS", "TypeHAlignCenter", HeaderColor
    grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 5, grdData
    Set_Headers
    If udtInPatSched.Facility_ID > 0 Then Get_Facilities udtInPatSched.Facility_ID
    mFirstLoad = False
    mbooByPass = True
    If cboApptType.Text = vbNullString Then
        For Idx = 0 To cboApptType.listcount - 1
            If cboApptType.List(Idx) = "Initial Round" Then
                cboApptType.Text = cboApptType.List(Idx)
                Exit For
            End If
        Next Idx
    End If
    mbooByPass = False
    grdData.AllowDragDrop = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_InPatientFormat:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_InPatientFormat", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function LookUpCellTag(vstrName As String, Optional vlngCol As Long) As String
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String
    
    lngCol = 0
    With grdData
        LookUpCellTag = vbNullString
        For lngRow = 1 To .MaxRows
            .Col = vlngCol
            If .CellTag <> vbNullString Then strTag = .CellTag
            If Trim(strTag) <> vbNullString And TextFound(vstrName, Trim(strTag), Len(vstrName), True) Then
                LookUpCellTag = lngCol & "|" & lngRow
                Exit For
            End If
        Next lngRow
    End With
    Exit Function
End Function

Public Sub Prepare_Grid()
Dim sValue As String
Dim Idx As Long
Dim Arr

    On Error GoTo Prepare_Grid
    Screen.MousePointer = vbHourglass
    mLoadPatient = False
    mGridClick = False
    mstrValidAppoint = vbNullString
    mbooByPass = False
    Clear_InSched
    Clear_Appt
    mstrApptMode = "Day"
    If mstrPractice = "Psychiatry" Then
        picPsych.Visible = True
    Else
        picPsych.Visible = False
    End If
    Get_InPatientFormat
    Format_Patient
    Format_Messages
    udtInPatSched.Appt_Date = Format(Now, "mm/dd/yyyy")
    mbooByPass = True
    MonthView1.Value = udtInPatSched.Appt_Date
    mbooByPass = False
    Clear_InSched
    Clear_Appt
    If PsUserID = vbNullString Then PsUserID = Get_PsUserID(psConnect(0))
    Create_Defaults
    mbooByPass = True
    If udtInPatSched.Facility_ID = 0 Then udtInPatSched.Facility_ID = piFacilityID
    mLoadPatient = False
    mlngClickPrev = 0
    mlngClickNext = 0
    mGridClick = False
    mstrFirstDay = vbNullString
    udtColor.Office = picOffice.BackColor
    udtColor.Hospital = picHospital.BackColor
    udtColor.Vacation = picBlock.BackColor
    udtColor.Cancelled = picCheckOut.BackColor
    udtColor.CheckOut = picCheckOut.BackColor
    udtColor.NoShow = picCancel.BackColor
    udtColor.Present = picPresent.BackColor
    udtColor.Lunch = picBlock.BackColor
    udtColor.PostOp = picPostOp.BackColor
    udtColor.PreOp = picPreOp.BackColor
    udtColor.NewPatient = picNewPatient.BackColor
    udtColor.ExistPatient = picExstNewPat.BackColor
    udtColor.Surgery = picSurgery.BackColor
    udtColor.FollowUp = picReturn.BackColor
    'grdData.ReDraw = False
    'grdData.ReDraw = True
    mFirstLoad = True
    mbooByPass = True
    mbooByPass = True
    If udtInPatSched.Appt_Date = vbNullString Then udtInPatSched.Appt_Date = Format(Now, "mm/dd/yyyy")
    MonthView1.Value = udtInPatSched.Appt_Date
    Load_Facilities
    mbooByPass = True
    Get_Appt_Providers
    For Idx = 0 To cboProviders.listcount - 1
        mbooByPass = True
        If cboProviders.ItemData(Idx) = udtInPatSched.Provider_ID Then
            cboProviders.Text = cboProviders.List(Idx)
            Exit For
        End If
    Next Idx
    mbooByPass = True
    For Idx = 0 To cboProviders.listcount - 1
        If cboProviders.ItemData(Idx) = mlngLogonUser Then
            cboProviders.Text = cboProviders.List(Idx)
            cboApptType.Clear
            cboApptType.AddItem "Cancellled"
            cboApptType.AddItem "Initial Round"
            cboApptType.AddItem "Second Round"
            cboApptType.AddItem "Present"
            cboApptType.ListIndex = -1
            Exit For
        End If
    Next Idx
    mbooByPass = True
    Load_Resource_Providers
    mbooByPass = True
    For Idx = 0 To cboExtender.listcount - 1
        If cboExtender.ItemData(Idx) = mlngLogonUser Then
            cboExtender.Text = cboExtender.List(Idx)
            cboExtender.Locked = True
            Exit For
        End If
    Next Idx
    mbooByPass = False
    mstrApptMode = "Day"
    mbooByPass = False
    If Not mbooLoad Then ReadData
    mbooByPass = True
    Read_Messages
    For Idx = 0 To cboProviders.listcount - 1
        If cboProviders.ItemData(Idx) = mlngLogonUser Then
            cboProviders.Text = cboProviders.List(Idx)
            cboApptType.Clear
            cboApptType.AddItem "Cancellled"
            cboApptType.AddItem "Initial Round"
            cboApptType.AddItem "Second Round"
            cboApptType.AddItem "Present"
            cboApptType.ListIndex = -1
            Exit For
        End If
    Next Idx
    Screen.MousePointer = vbDefault
    Exit Sub
Prepare_Grid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Prepare_Grid", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Function Create_Defaults()
Dim i%
Dim lngRow As Long
Dim varCellValue

    Set_DefaultFacility
    With grdFacility
        For lngRow = 1 To .MaxRows
            If .GetRowItemData(lngRow) = piFacilityID Then
                .ActiveRow = lngRow
                .GetText 1, lngRow, varCellValue
                .Col = 1
                .Row = lngRow
                .BackColor = grdCurItem
                cctFacility.Text = varCellValue
                Exit For
            End If
        Next lngRow
    End With
     Get_Appt_Providers
     For i = 0 To cboProviders.listcount - 1
        mbooByPass = True
         If cboProviders.ItemData(i) = udtInPatSched.Provider_ID Then
            mbooByPass = True
            cboProviders.Text = cboProviders.List(i)
            lblCurScheAppt.Visible = True
            lblCurScheAppt.Caption = "Active Roster"
            lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
            Exit For
         End If
     Next i
     mbooByPass = False
     Exit Function
End Function

Private Function Set_Headers()
Dim vDate
Dim lngFacID As Long
Dim Idx As Integer
Dim lngSpan As Long

    udtInPatSched.Appt_Date = Format(MonthView1.Value, "MMMM DD, YYYY")
    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        lngFacID = .GetRowItemData(.ActiveRow)
    End With
    mbooByPass = False
    Get_Providers mlngCurProvID
    lngSpan = grdData.MaxCols
    If lngFacID > 0 Then Get_Facilities lngFacID
    grdData.RowHeight(SpreadHeader) = 20
        Format_CellTypeEdit grdData, SpreadHeader, 1, ".maxcols", UCase(clsFacility.Name), "TypeHAlignCenter", "", HeaderColor, 14
        Add_Spacer SpreadHeader + 1, grdData
        Format_CellTypeEdit grdData, SpreadHeader + 2, 1, CStr(lngSpan), "VISIT ROSTER" & vbNewLine & "FOR PERIOD BETWEEN" & " " & Format(mstrRangeDate, "mm/dd/yyyy") & " " & "AND" & " " & Format(mstrUpperRange, "mm/dd/yyyy") & vbNewLine & UCase(clsProvider.Name), "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder 1, SpreadHeader + 2, lngSpan, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 3, grdData
    Exit Function
End Function
Private Function Check_Patient(vlngPatientID As Long, vstrApptDate As String)
Dim strSQL As String
Dim sDate As String
Dim cnValidate As New ADODB.Connection
Dim rsValidate As New ADODB.Recordset
Dim lngApptID As Long
Dim lngPatID As Long
Dim strNow As String
Dim strTime As String

    On Error GoTo Error_Check_Patient
    If cnValidate.State = adStateClosed Then cnValidate.open psConnect(0)
    strNow = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    Check_Patient = False
    lngPatID = vlngPatientID
    sDate = Format(vstrApptDate, "mm/dd/yyyy")
    strSQL = "Get_InPatient_Appointment_Patient  " & udtInPatSched.Facility_ID & "," & udtInPatSched.Provider_ID & ", " & lngPatID & ", '" & Format(udtInPatSched.Appt_Date, "mm/dd/yyyy") & "'"
    Set rsValidate = cnValidate.Execute(strSQL)
    With rsValidate
        If Not .EOF Then
            If (!Patient_ID) = lngPatID Then
                Check_Patient = True
                lngApptID = (!INPATIENT_APPT_ID)
            End If
        End If
    End With
    If Check_Patient Then
        'need to add extender
        If Validate_Extender_Provider Then
'            strSQL = "Update_InPatient_Status_Appt_ID " & lngApptID & ", 'Workup Completed'"
'            Set rsValidate = cnValidate.Execute(strSQL)
        ElseIf Validate_MD_Provider Then
            strSQL = "Update_InPatient_Status_Appt_ID " & lngApptID & ", 'Present'"
            Set rsValidate = cnValidate.Execute(strSQL)
        End If
    End If
    If cnValidate.State = adStateOpen Then
        cnValidate.Close
        Set cnValidate = Nothing
        Set rsValidate = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Check_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Check_Patient", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_NurseDiagnosis()
Dim cnPlan As New ADODB.Connection
Dim rsPlan As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim Arr
Dim strCode As String
Dim strDiag As String
Dim Arrdiag
Dim booSkip As Boolean
Dim lngProcRow As Long
Dim Idx As Integer
Dim strAddProcCode As String
Dim strAddProc As String
Dim i%
Dim strCstrRow As String
Dim lngCol As Long
Dim lngFee As Long
Dim strList As String
Dim varCellValue
Dim strProcCode As String
Dim strCodeDesc As String
Dim ArrCode
Dim lngChargeRow As Long

    On Error GoTo Error_Get_NurseDiagnosis
    Screen.MousePointer = vbHourglass
    If cnPlan.State = adStateClosed Then cnPlan.open psConnect(0)
    strSQL = "Get_Encounter_Service_Join_Patient_ID " & Val(udtInPatSched.IntId)
    Set rsPlan = cnPlan.Execute(strSQL)
    With rsPlan
        Do Until .EOF
            If Not IsNull((!DIAGNOSIS_CODE)) Then
                If Trim((!DIAGNOSIS_CODE)) <> vbNullString Then
                    clsPatAcct.ICD_Code = (!DIAGNOSIS_CODE)
                    Exit Do
                End If
            End If
            .MoveNext
        Loop
    End With
    clsPatAcct.ICD_Code = Replace(clsPatAcct.ICD_Code, "|", ";")
    If cnPlan.State = adStateOpen Then
        cnPlan.Close
        Set cnPlan = Nothing
        Set rsPlan = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_NurseDiagnosis:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_NurseDiagnosis", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Function Get_CurrentEnctr_MD(Optional vstrDate As String)
Dim strSQL As String
Dim cnStatus As New ADODB.Connection
Dim rsStatus As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim strEnctTime As String
Dim lngEnctrID As Long
Dim iAnswer As Integer
Dim strDate As String
Dim lngTestProvider As Long
Dim i%
Dim strInsert As String

    On Error GoTo Error_Get_CurrentEnctr_MD
    Screen.MousePointer = vbHourglass
    strDate = vstrDate
    Dump_EncounterStrData
    piEncounterID = 0
    If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
    Get_Providers mlngLogonUser
    piEncounterID = 0
    If Trim(udtInPatSched.Appt_Date) <> vbNullString Then
        strDate = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
    Else
        strDate = Format(Now, "mm/dd/yyyy")
    End If
    If clsEDI_Client.Site_ID = "1999" Then
        If udtInPatSched.Provider_ID = 0 Then udtInPatSched.Provider_ID = 1000
    End If
    udtInPatSched.Appt_Date = Format(strDate, "mm/dd/yyyy")
    udtTime.StartDate = Format(strDate, "mm/dd/yyyy 00:00:00 AM/PM")
    udtTime.EndDate = DateAdd("h", 23, udtTime.StartDate)
    udtTime.EndDate = Format(udtTime.EndDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
    strSQL = "SELECT * FROM ENCOUNTER WHERE PATIENT_ID = " & piPatientID & " AND PROVIDER_ID = " & mlngLogonUser & " AND " _
        & "FACILITY_ID  = " & udtInPatSched.Facility_ID & " AND EXT_ENCOUNTER_ID IS NULL AND PURGE = 0 AND " _
        & "CURRENT_STATE IS NOT NULL ORDER BY ENCNTR_START_DATE DESC, CURRENT_STATE "
    Set rsStatus = cnStatus.Execute(strSQL)
    With rsStatus
        Do Until .EOF
            Dump_EncounterStrData
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case Fld.Name
                            Case "CURRENT_STATE"
                                clsEncounter.Encounter_State = Fld.Value
                            Case "ENCNTR_START_DATE"
                                clsEncounter.StartDate = Fld.Value
                            Case "RENDERING_PROVIDER"
                                clsEncounter.RendProv = Fld.Value
                            Case "FACILITY_ID"
                                clsEncounter.FacilityID = Fld.Value
                            Case "ENCOUNTER_ID"
                                clsEncounter.EncounterID = Fld.Value
                            Case "INSERT_DATE"
                                strInsert = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If clsEncounter.StartDate = vbNullString Then
                clsEncounter.EncounterID = 0
                Dump_EncounterStrData
            Else
                lngDateDiff = DateDiff("d", Format(clsEncounter.StartDate, "mm/dd/yyyy"), strDate)
                If lngDateDiff > 30 Then
                    clsEncounter.EncounterID = 0
                    Dump_EncounterStrData
                End If
            End If
        Loop
    End With
    If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
    If clsEncounter.EncounterID > 0 Then
        strSQL = "Update_Encounter_InPatient_All " & mlngLogonUser & ", " _
            & clsEncounter.EncounterID & ", '" & Format(udtTime.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "'"
        Set rsStatus = cnStatus.Execute(strSQL)
        Get_Encounter piEncounterID
        If udtInPatSched.Facility_ID > 0 Then
            Get_Facilities udtInPatSched.Facility_ID
        End If
        If clsFacility.FacType <> vbNullString Then
            clsEncounter.ServicePlace = clsFacility.FacType
        End If
        If clsFacility.pos <> 0 Then
            clsEncounter.ServicePlaceCode = clsFacility.pos
        End If
        If udtInPatSched.PatientType = "Follow-Up" Then
            clsEncounter.EncounterType = "Subsequent Care"
            udtInPatSched.PatientType = "Subsequent Care"
        ElseIf udtInPatSched.PatientType = "Initial" Or udtInPatSched.PatientType = "Initial InPatient" Then
            clsEncounter.EncounterType = "Initial InPatient"
            udtInPatSched.PatientType = "Initial InPatient"
        ElseIf udtInPatSched.PatientType = "Re-Admission" Then
            clsEncounter.EncounterType = "Initial InPatient"
            udtInPatSched.PatientType = "Initial InPatient"
        End If
        If clsEncounter.ServicePlaceCode = "31" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99307"
                clsEncounter.VisitType = "Subsequent care, nursing facility E&M, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = "99304"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        ElseIf clsEncounter.ServicePlaceCode = "13" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99334"
                clsEncounter.VisitType = "Follow-Up visit, asst living facility, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = " 99325"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        End If
        If clsEncounter.ServicePlaceCode = "31" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99307"
                clsEncounter.VisitType = "Subsequent care, nursing facility E&M, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = "99304"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        ElseIf clsEncounter.ServicePlaceCode = "13" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99334"
                clsEncounter.VisitType = "Follow-Up visit, asst living facility, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = " 99325"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        End If
        strSQL = "UPDATE ENCOUNTER SET ENCOUNT_CPT_INTRVN = '" & clsEncounter.Intrvn_CD & "', " _
            & "ENCNTR_TYPE = '" & clsEncounter.EncounterType & "',SERVICE_PLACE = '" & clsEncounter.ServicePlace & "', " _
            & "SERVICE_PLACE_CD = '" & clsEncounter.ServicePlaceCode & "' WHERE ENCOUNTER_ID = " & clsEncounter.EncounterID
        Set rsStatus = cnStatus.Execute(strSQL)
    ElseIf clsEncounter.EncounterID = 0 Then
        strSQL = "Get_InPatient_Appt_ID " & piPatientID & ",'" & Format(udtInPatSched.Appt_Date, "mm/dd/yyyy") & "', " & mlngLogonUser
        Set rsStatus = cnStatus.Execute(strSQL)
        With rsStatus
            If Not .EOF Then
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "PATIENT_TYPE"
                                    udtInPatSched.PatientType = Trim(Fld.Value)
                                Case "PURPOSE_VISIT"
                                    udtAppointment.VisitType = Trim(Fld.Value)
                                Case "FACILITY_ID"
                                    clsEncounter.FacilityID = Fld.Value
                                    clsEncounter.FacilityID = Fld.Value
                            End Select
                        End If
                    End If
                Next
            End If
        End With
        If PsUserID = vbNullString Then
            PsUserID = Get_PsUserID(psConnect(0))
        End If
        Select Case udtAppointment.VisitType
            Case "Follow-Up", "Subsequent Care", "Subsequent Visit", "Subsequent care, nursing facility E&M, limited", "Follow-Up visit, asst living facility, limited"
                udtInPatSched.PatientType = "Subsequent Care"
            Case "Initial InPatient", "New Patient", "Initial Visit", "Initial E&M Care, Detailed Visit", "Initial E&M Care, Brief Visit", _
                "Initial E&M Care, Extended Visit", "Initial E&M Care, Complex Visit", "Diagnostic Interview Exam"
                udtInPatSched.PatientType = "Initial InPatient"
        End Select
        If Trim(udtInPatSched.PatientType) <> vbNullString Then
            clsEncounter.EncounterType = udtInPatSched.PatientType
        End If
        If clsEncounter.FacilityID > 0 Then
            Get_Facilities clsEncounter.FacilityID
        End If
        If clsFacility.FacType <> vbNullString Then
            clsEncounter.ServicePlace = clsFacility.FacType
        End If
        If clsFacility.pos <> 0 Then
            clsEncounter.ServicePlaceCode = clsFacility.pos
        End If
        If clsEncounter.ServicePlaceCode = "31" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99307"
                clsEncounter.VisitType = "Subsequent care, nursing facility E&M, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = "99304"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        ElseIf clsEncounter.ServicePlaceCode = "13" Then
            If udtInPatSched.PatientType = "Subsequent Care" Then
                clsEncounter.Intrvn_CD = "99334"
                clsEncounter.VisitType = "Follow-Up visit, asst living facility, limited"
            ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                clsEncounter.Intrvn_CD = " 99325"
                clsEncounter.VisitType = "Initial E&M Care, Brief Visit"
            End If
        End If
        strSQL = "Write_Encounter_Extender " & piPatientID & "," & mlngLogonUser & "," & udtInPatSched.Facility_ID & ",'" _
                & Format(udtTime.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "', '" & clsEncounter.Intrvn_CD & "','" & clsEncounter.VisitType & "','" _
                & clsEncounter.EncounterType & "','" & clsEncounter.ServicePlace & "'," & clsEncounter.ServicePlaceCode & ",'" & PsUserID & "','" _
                & Format(udtTime.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," & mlngLogonUser & ",'" & udtAppointment.VisitType & "','MD COMPLETED'"
            Set rsStatus = cnStatus.Execute(strSQL)
            With rsStatus
                If Not .EOF Then
                    clsEncounter.EncounterID = (!Encounter_ID)
                End If
            End With
    End If
    piEncounterID = clsEncounter.EncounterID
    Get_Encounter piEncounterID
    modSpecific.piEncounterID = piEncounterID
    If cnStatus.State = adStateOpen Then
        cnStatus.Close
        Set cnStatus = Nothing
        Set rsStatus = Nothing
        Set Flds = Nothing
    End If
    mbooPatValid = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_CurrentEnctr_MD:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_CurrentEnctr_MD", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function IsCellSpan(vlngRow As Long, Optional vlngCol As Long, Optional vlngTarRow As Long) As Boolean
Dim Col As Variant
Dim Row As Variant
Dim Col2 As Variant
Dim Row2 As Variant
Dim Span As Integer
Dim anchorcol As Variant
Dim anchorrow As Variant
Dim colspan As Variant
Dim rowspan As Variant
Dim i As Long
Dim j As Long
Dim IsSpan As Boolean
Dim strTag As String
Dim strCurTag As String
Dim varCellValue
Dim lngPatID As Long

    If Not mbooOverBook Then
        Exit Function
    End If
    grdData.getSelection 0, Col, Row, Col2, Row2
    grdData.Redraw = False
    If vlngRow > vlngTarRow And vlngTarRow > 0 Then
        grdData.Col = vlngCol
        grdData.Row = i
        If grdData.CellTag <> vbNullString Then
            IsCellSpan = True
        End If
    ElseIf vlngRow > 0 Then
        For i = vlngRow To vlngTarRow
            If vlngCol > 0 Then
                For j = vlngCol To grdData.MaxCols
                    Span = grdData.GetCellSpan(vlngCol, vlngRow, anchorcol, anchorrow, colspan, rowspan)
                    grdData.Col = vlngCol
                    grdData.Row = i
                    If grdData.CellTag <> vbNullString Then
                        IsCellSpan = True
                        Exit For
                    End If
                    If Span Then
                        If j = vlngCol Then anchorcol = "1"
                        If i = vlngRow Then
                            IsCellSpan = True
                        End If
                    End If
                Next j
            Else
                Span = grdData.GetCellSpan(1, vlngRow, anchorcol, anchorrow, colspan, rowspan)
                grdData.Redraw = False
                If Span Then
                     If i = vlngRow Then
                        IsCellSpan = True
                        Exit For
                    End If
                End If
            End If
        Next i
    End If
    grdData.Redraw = True
    Exit Function
End Function
Private Function DayMode_BlockTime(vstrBlockReason As String, vstrBlockTo As String, vstrBlockFrom As String, _
vstrBlockDuration As String, vstrStartDate As String, vlngProviderID As Long)
Dim i%
Dim lngProviderID As Long
Dim strSQL As String
Dim vTime
Dim vDate
Dim BlockColor As Long
Dim lngStartRow As Long
Dim lngEndRow As Long
Dim lngNoRows  As Long
Dim nDays As Double
Dim Arr
Dim lngDateDiff As Long
Dim lngBlockCol As Long
Dim lngBlockCol2 As Long
Dim lngCol As Long
Dim lngRow As Long
Dim varCellValue
Dim booCancel As Boolean
Dim booSet As Boolean
Dim strErrata As String

    On Error GoTo Error_DayMode_BlockTime
    If vstrBlockFrom = vbNullString Then Exit Function
    Screen.MousePointer = vbHourglass
    
    Select Case IsProperCase(Trim(vstrBlockDuration))
        Case "1 Week"
            nDays = 7
        Case "2 Weeks", "2 Week"
            nDays = 14
        Case "3 Week", "3 Weeks"
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
            nDays = 0
    End Select
    Set_Headers
    For lngCol = 1 To grdData.MaxCols
        grdData.Col = lngCol
        grdData.Row = SpreadHeader + 3
        If grdData.CellTag <> vbNullString Then
            If vlngProviderID = Val(grdData.CellTag) Then
                If lngBlockCol = 0 Then
                    lngBlockCol = lngCol
                Else
                    lngBlockCol2 = lngCol
                    Exit For
                End If
            End If
        End If
    Next lngCol
'    If ndays > 1 Then
'        vDate = DateAdd("d", ndays, vstrStartDate)
'    Else
    vDate = vstrStartDate
'    End If
    vDate = Format(vDate, "mm/dd/yyyy")
'    lngDateDiff = DateDiff("d", udtInPatSched.Appt_Date , vDate)
    lngDateDiff = DateDiff("d", udtInPatSched.Appt_Date, udtBlock.EndDate)
    If lngDateDiff >= 0 And (lngDateDiff <= nDays Or IsProperCase(vstrBlockDuration) = "Permanent") Then
'    If lngDateDiff <= 0 And lngDateDiff <= ndays Then
        For lngCol = 1 To grdData.MaxCols
            If IsProperCase(vstrBlockDuration) = "Permanent" Then
                grdData.Col = lngCol
                grdData.Row = SpreadHeader + 3
                If grdData.CellTag <> vbNullString Then
                    If vlngProviderID = grdData.CellTag Then
                        lngBlockCol = lngCol
                        Exit For
                    End If
                End If
            Else
                lngDateDiff = DateDiff("d", udtBlock.EndDate, udtInPatSched.Appt_Date)
                If lngDateDiff <= 0 Then
                    If vlngProviderID = grdData.GetColItemData(lngCol) Then
                        lngBlockCol = lngCol
                        Exit For
                    End If
                Else
                    booCancel = True
                    Exit For
                End If
            End If
        Next lngCol
        If grdData.MaxRows = 0 Then grdData.MaxRows = 1
        If Not booCancel Then
            For lngRow = 1 To grdData.MaxRows
                grdData.GetText 0, lngRow, vTime
                If UCase(vTime) = UCase(vstrBlockFrom) Then
                    lngStartRow = lngRow
                End If
                vTime = vbNullString
                grdData.GetText 0, lngRow, vTime
                If UCase(vTime) = UCase(vstrBlockTo) Then
                    lngEndRow = lngRow
                End If
            Next lngRow
            If lngEndRow = 0 Then lngEndRow = grdData.MaxRows
            lngNoRows = lngEndRow - lngStartRow
            Select Case IsProperCase(vstrBlockReason)
                Case "Lunch"
                    BlockColor = grdBlock
                Case "Hospital"
                    BlockColor = grdBlock
                Case "Vacation"
                    BlockColor = grdBlock
                Case "Out Of The Office"
                    BlockColor = grdBlock
            End Select
            lngNoRows = lngNoRows
            If lngBlockCol = 0 Then
                lngBlockCol = 1
            End If
            If lngStartRow >= 1 And lngEndRow > 1 Then
                For lngRow = lngStartRow To lngEndRow
                    clsGridFormat.CellNote = vbNullString
                    grdData.GetText lngBlockCol, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString And Not TextFound(vstrBlockReason, CStr(varCellValue), Len(vstrBlockReason), True) Then
                        If lngRow = lngStartRow Or lngRow > lngStartRow And lngRow < lngEndRow Then
                            grdData.Col = lngBlockCol
                            grdData.Row = lngRow
                            grdData.GetText lngBlockCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                clsGridFormat.CellNote = Trim(varCellValue)
                                Arr = Split(varCellValue, "INSURANCE AUTHORIZATION:")
                                If UBound(Arr) > 0 Then
                                   varCellValue = Trim(Arr(0))
                                End If
                                clsGridFormat.CellNote = Trim(varCellValue)
                                If UBound(Arr) > 0 Then
                                   clsGridFormat.CellNote = Arr(0)
                                End If
                                If Trim(clsGridFormat.CellNote) <> vbNullString Then
                                    grdData.GetText 0, lngRow, varCellValue
                                    If Trim(varCellValue) <> vbNullString Then
                                        strErrata = Trim(varCellValue)
                                    End If
                                    If Trim(clsGridFormat.CellNote) <> vbNullString Then strErrata = "Patient is scheduled in the block time at:" & " " & strErrata & vbNewLine & clsGridFormat.CellNote
                                    If Trim(rtfChanges.Text) <> vbNullString Then
                                        rtfChanges.Text = rtfChanges.Text & vbNewLine & Trim(strErrata)
                                    ElseIf Not TextFound(Trim(strErrata), Trim(rtfChanges.Text), Len(Trim(strErrata)), True) Then
                                        rtfChanges.Text = strErrata & vbNewLine & Trim(strErrata)
                                    End If
                                End If
                            End If
                            Exit For
                        End If
                    ElseIf lngRow > lngStartRow Then
                        lngNoRows = lngNoRows - 1
                    End If
                Next lngRow
            End If
        End If
        If Not booSet Then
            If lngNoRows = 0 Then lngNoRows = 1
            lngNoRows = lngEndRow - lngStartRow
            If lngNoRows = 0 Then lngNoRows = 1
            If lngBlockCol2 = 0 Then
                Format_CellTypeEdit_RowSpan grdData, lngStartRow, lngBlockCol, CStr(lngNoRows), vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
                grdData.SetCellBorder lngBlockCol, lngStartRow, lngBlockCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdData.SetCellBorder lngBlockCol, lngEndRow, lngBlockCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Else
                Format_CellTypeEdit_RowSpan grdData, lngStartRow, lngBlockCol, CStr(lngNoRows), vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
                grdData.SetCellBorder lngBlockCol, lngStartRow, lngBlockCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdData.SetCellBorder lngBlockCol, lngEndRow, lngBlockCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Format_CellTypeEdit_RowSpan grdData, lngStartRow, lngBlockCol2, CStr(lngNoRows), vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
                grdData.SetCellBorder lngBlockCol2, lngStartRow, lngBlockCol2, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdData.SetCellBorder lngBlockCol2, lngEndRow, lngBlockCol2, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_DayMode_BlockTime:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "DayMode_BlockTime", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Delete_Dup_Appt(vlngAppt_ID As Long)
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String

    On Error GoTo Error_Delete_Dup_Appt
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "DELETE INPATIENT_APPOINTMENTS WHERE INPATIENT_APPT_ID = " & vlngAppt_ID
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Delete_Dup_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Delete_Dup_Appt", App.EXEName & "." & Me.Name
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

    'If udtBlock.Duration = "Permanent" Then lngDateDiff = 1
    lngDateDiff = DateDiff("ww", udtInPatSched.Appt_Date, udtBlock.EndDate)
    If udtBlock.Duration = "Permanent" Then lngDateDiff = 1
    If lngDateDiff >= 0 Then
        strRepDate = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
        If lngDateDiff = 0 Then lngDateDiff = 1
        For Idx = 1 To lngDateDiff
            FirstDay = WeekDay(strRepDate, vbMonday)
            nDays = Val(FirstDay)
            If strRepDate <> vbNullString Then
                Select Case udtBlock.DayWeek
                    Case "Monday"
                        FirstDay = WeekDay(strRepDate, vbMonday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Tuesday"
                        FirstDay = WeekDay(strRepDate, vbTuesday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Wednesday"
                        FirstDay = WeekDay(strRepDate, vbWednesday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Thursday"
                        FirstDay = WeekDay(strRepDate, vbThursday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Friday"
                        FirstDay = WeekDay(strRepDate, vbFriday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Saturday"
                        FirstDay = WeekDay(strRepDate, vbSaturday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Sunday"
                        FirstDay = WeekDay(strRepDate, vbSunday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                End Select
            End If
        Next
    End If
    Exit Function

End Function
Public Function Get_Mode()
Dim i%

    Get_SurgFacility
    Select Case mstrApptMode
        Case "Day"
            If udtInPatSched.Provider_ID = 0 Then udtInPatSched.Provider_ID = mlngCurProvID
            If mstrGridType = vbNullString Then
                Get_GridType
            End If
            If mstrGridType = "Name Only" Then
            Else
                ReadData
            End If
        Case "Weekly"
            Format_Week
            Read_Week_InPatient_Appointments
    End Select
    mbooByPass = True
    For i = 0 To cboProviders.listcount - 1
        mbooByPass = True
        If cboProviders.List(i) = udtInPatSched.Provider Then
            mbooByPass = True
            cboProviders.Text = cboProviders.List(i)
            Exit For
        End If
    Next i
    Exit Function
End Function
Private Function Get_Changes(vstrDate As String)
Dim i%
Dim strSQL As String
Dim cnChange As New ADODB.Connection
Dim rsChange As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld  As ADODB.Field
Dim lngProviderID  As Long
Dim Arr
Dim sDate As String
Dim sProviderName As String
Dim strLastName As String
Dim strFirstName As String
Dim strStatus As String
Dim strReason As String
Dim strText As String
Dim strCollect As String
Dim strLine As String
Dim lngPatID As Long
Dim strDuration As String
Dim strType As String
Dim strTime As String
Dim lngRow As Long

    On Error GoTo Error_Get_Changes
    Screen.MousePointer = vbHourglass
    grdData.Redraw = False
    strLine = "_______________________"
    If mstrApptMode = "Day" Then
        sDate = vstrDate
    Else
        Arr = Split(vstrDate, "-")
        If UBound(Arr) > 0 Then
            sDate = Arr(1)
        Else
            sDate = vstrDate
        End If
    End If
    sDate = Format(sDate, "mm/dd/yyyy")
    If cnChange.State = adStateClosed Then cnChange.open psConnect(0)
    strSQL = "SELECT * FROM INPATIENT_APPOINTMENTS  " _
        & "WHERE PROVIDER_ID = " & udtInPatSched.Provider_ID & " AND  APPT_DATE = '" & Format(sDate, "mm/dd/yyyy") & "' AND FACILITY_ID = " & udtInPatSched.Facility_ID
    Set rsChange = cnChange.Execute(strSQL)
    With rsChange
        Do Until .EOF
            strCollect = vbNullString
            strLastName = vbNullString
            strFirstName = vbNullString
            strStatus = vbNullString
            strReason = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PROVIDER_ID"
                                For i = 0 To cboProviders.listcount - 1
                                    mbooByPass = True
                                    If cboProviders.ItemData(i) = Fld.Value Then
                                        mbooByPass = True
                                        sProviderName = cboProviders.List(i)
                                        Exit For
                                    End If
                                Next i
                            Case "REASON"
                                strReason = Fld.Value
                            Case "STATUS"
                                strStatus = Fld.Value
                            Case "PATIENT_ID"
                                lngPatID = Fld.Value
                            Case "DURATION"
                                strDuration = Fld.Value
                            Case "APPT_TYPE"
                                strType = Fld.Value
                            Case "APPT_TIME"
                                strTime = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If lngPatID > 0 Then
                Get_NewInPatient lngPatID
                If udtInPatSched.LastName <> vbNullString And Trim(strStatus) <> vbNullString Then
                    strLastName = udtInPatSched.LastName & ", " & udtInPatSched.FirstName & " " & Trim(strStatus) & " " & Trim(strReason)
                    Select Case IsProperCase(strStatus)
                        Case "Rescheduled"
                            If IsProperCase(strReason) = "Change" Then strReason = vbNullString
                            If Trim(strLastName) <> vbNullString Then
                                If Trim(rtfChanges.Text) = vbNullString Then
                                    rtfChanges.Text = Trim(strLastName)
                                ElseIf Not TextFound(Trim(strLastName), Trim(rtfChanges.Text), Len(Trim(strLastName)), True) Then
                                    rtfChanges.Text = Trim(rtfChanges.Text) & vbNewLine & strLine & vbNewLine & Trim(strLastName)
                                End If
                            End If
                        Case "Cancelled", "No Show"
                            If Trim(strLastName) <> vbNullString Then
                                If Trim(rtfChanges.Text) = vbNullString Then
                                    rtfChanges.Text = Trim(strLastName)
                                ElseIf Not TextFound(Trim(strLastName), Trim(rtfChanges.Text), Len(Trim(strLastName)), True) Then
                                    rtfChanges.Text = Trim(rtfChanges.Text) & vbNewLine & strLine & vbNewLine & Trim(strLastName)
                                End If
                            End If
                        Case "Present"
                            For lngRow = 1 To grdData.MaxRows
                                grdData.Col = 2
                                grdData.Row = lngRow
                                If grdData.CellTag = CStr(lngPatID) Then
                                    grdData.BlockMode = True
                                        grdData.Col = 1
                                        grdData.Col = grdData.MaxCols
                                        grdData.Row = lngRow
                                        grdData.Row2 = lngRow
                                        grdData.BackColor = grdCurItem
                                    grdData.BlockMode = False
                                    Exit For
                                End If
                            Next lngRow
                    End Select
                End If
            End If
        Loop
    End With
    If cnChange.State = adStateOpen Then
        cnChange.Close
        Set cnChange = Nothing
        Set rsChange = Nothing
        Set Flds = Nothing
    End If
    mbooByPass = False
    grdData.Redraw = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Changes:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Changes", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Public Sub Manage_Colors(vstrBlockReason As String, vlngPatientID As Long)
Dim lngBlockRow As Long
Dim lngCol As Long
Dim lngRow As Long
Dim mbLeaveCol As Boolean
Dim lngCurPatient As Long

    On Error GoTo Error_Manage_Colors
    Screen.MousePointer = vbHourglass
    lngCurPatient = vlngPatientID
    If Trim(vstrBlockReason) <> vbNullString Then
        If TextFound("Can", vstrBlockReason, Len("Can"), True) Then
                'clsGridFormat.BackColor = &HC0FFC0
        ElseIf vstrBlockReason = "Resched" Or vstrBlockReason = "Delete" Then
                clsGridFormat.BackColor = &HC0FFC0
        ElseIf TextFound("Present", vstrBlockReason, Len("Present"), True) Then
                clsGridFormat.BackColor = &HFFC0FF
        ElseIf TextFound("Present", vstrBlockReason, Len("Check Out"), True) Then
               clsGridFormat.BackColor = &HC0FFFF
        ElseIf TextFound("Confirmed", vstrBlockReason, Len("Confirmed"), True) Then
               clsGridFormat.BackColor = &HC0E0FF
        ElseIf TextFound("No Show", vstrBlockReason, Len("No Show"), True) Then
                clsGridFormat.BackColor = &HC0FFC0
        Else
                clsGridFormat.BackColor = &H80000005
        End If
        With grdData
            If clsGridFormat.BackColor <> &HFFFFFF Then
                For lngCol = 1 To .MaxCols
                    For lngRow = 1 To .MaxRows
                        grdData.Redraw = False
                        .Col = lngCol
                        .Row = lngRow
                        If .CellTag <> vbNullString Then lngCurPatient = Val(.CellTag)
                        If lngCurPatient = vlngPatientID Then
                            mbLeaveCol = True
                            Exit For
                        End If
                    Next lngRow
                    If mbLeaveCol Then Exit For
                Next lngCol
                If lngCurPatient <> 0 Then
                    .BlockMode = True
                        .Col = lngCol
                        .Col2 = lngCol
                        .Row = lngRow
                        .Row2 = lngRow
                        If clsGridFormat.BackColor = 0 Then clsGridFormat.BackColor = &H80000005
                        If clsGridFormat.BackColor <> 0 Then grdData.BackColor = clsGridFormat.BackColor
                    .BlockMode = False
                End If
            End If
            grdData.Redraw = True
        End With
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Manage_Colors:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Manage_Colors", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Function Remove_Data(vlngPatID As Long)
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim strDate As String
Dim lngProv As Long
Dim Idx%

    On Error GoTo Error_Remove_Data
    Screen.MousePointer = vbHourglass
    For Idx = 0 To cboProviders.listcount - 1
        If cboProviders.List(Idx) = cboProviders.Text Then
            lngProv = cboProviders.ItemData(Idx)
            Exit For
        End If
    Next Idx
    If Not IsNull(txtDate.Value) Then strDate = Format(txtDate.Value, "mm/dd/yyyy")
    If IsDate(strDate) Then
        If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
        strSQL = "Remove_InPatient_Appointment " & vlngPatID & ",'" & strDate & "'," & lngProv
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
        If cnReadAppt.State = adStateOpen Then
            cnReadAppt.Close
            Set cnReadAppt = Nothing
            Set rsReadAppt = Nothing
        End If
    End If
    Screen.MousePointer = vbDefault
    grdData.Redraw = False
    Exit Function
Error_Remove_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", " Remove_Data ", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Manage_Surgery(vstrType As String)
Dim vDate
Dim strDate As String
Dim Arr
Dim vPatient
Dim lngPatientID As Long
Dim vTime
Dim sFirstName As String
Dim sLastName As String
Dim strSQL As String
Dim cnData As New ADODB.Connection
Dim rsData As New ADODB.Recordset
Dim Idx As Integer
Dim lngSpanRow As Long
Dim strSpan As String
Dim lngCol As Long
Dim lngRow As Long
Dim strMoveTime As String
Dim varCellValue
Dim lngRowTime As Long

'    If Row < 1 Then Exit Sub
'    If Col < 1 Then Exit Sub
'    mlngGridCol = Col
'    mlngGridRow = Row
    Select Case vstrType
        Case "DblClick"
            frmMovePatient.Show vbModal
            If frmMovePatient.mbMovePatient Then
                'My notes :  Get The existing Column Name if the column name not Patient Surgery then update the Encounter set Encounter End Date = this date
                'Update_EnctrEndDate(strHeader)
                grdData.Col = mlngGridCol
                grdData.Row = mlngGridRow
                If grdData.CellTag <> vbNullString Then udtInPatSched.IntId = grdData.CellTag
                    If Val(udtInPatSched.IntId) > 0 Then
                    grdData.ClearRange mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, False
                    grdData.GetText mlngGridCol, mlngGridRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then udtInPatSched.Collection = Trim(varCellValue)
                    If udtInPatSched.Collection <> vbNullString And TextFound("Operating", frmMovePatient.mstrResource, Len("Operating"), True) Then 'And Resource = "Operating Room"
                        Arr = Split(udtInPatSched.Collection, vbNewLine)
                        If UBound(Arr) > 0 Then
                            For Idx = LBound(Arr) To UBound(Arr)
                                If TextFound("minutes", Arr(Idx), Len("minute"), True) Then
                                    udtInPatSched.Duration = Replace(Arr(Idx), "minutes", "")
                                     If IsNumeric(udtInPatSched.Duration) Then lngSpanRow = Val(udtInPatSched.Duration / lngRowTime)
                                     Exit For
                                End If
                            Next Idx
                        End If
                    End If
                    Arr = Split(LookUpColName("frmMovePatient.mstrResource"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                    End If
                    'Need Row
                    lngRow = 0
                    If udtAppointment.SchedTime <> vbNullString Then
                        
                    End If
                    For Idx = 1 To grdData.MaxRows
                        grdData.GetText 0, Idx, vTime
                        If vTime = vbNullString Then
                          grdData.GetText 0, Idx - 1, vTime
                        End If
                        If Format(strMoveTime, "hh:mm AM/PM") = CStr(vTime) Then
                            lngRow = Idx
                            Exit For
                        End If
                    Next Idx
                    
                    If lngSpanRow >= 0 Then strSpan = lngSpanRow
                    Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, strSpan, udtInPatSched.Collection, "TypeHAlignLeft", CStr(udtInPatSched.IntId), udtColor.Surgery
                    grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow + (Val(strSpan) - 1), CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Else
                    MsgBox "Please select a valid Patient Row to move!", vbInformation
                End If
            End If
    End Select
    Exit Function
End Function
Function Check_Schedule()
Dim lngRow As Long
Dim lngPatientID As Long
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim strSQL As String
Dim strApptDate As String
Dim Arr
Dim vApptDate
Dim strStatus As String
Dim lngCol As Long
Dim varCellValue
Dim strTag As String
Dim strNote As String
Dim lngColor As Long
Dim lngRowTime As Long


    On Error GoTo Error_Check_Schedule
    Screen.MousePointer = vbHourglass
    Get_Practice
    grdData.Redraw = False
    strApptDate = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
    With grdData
        If mstrApptMode = "Day" Then
                For lngRow = 1 To .MaxRows
                    varCellValue = vbNullString
                    strNote = vbNullString
                    lngPatientID = 0
                    .Col = 1
                    .Row = lngRow
                    If .CellTag <> vbNullString And IsNumeric(.CellTag) Then lngPatientID = Val(.CellTag)
                    If lngPatientID <> 0 Then
                        cnSched.open psConnect(0)
                            strSQL = "SELECT STATUS, DURATION FROM INPATIENT_APPOINTMENTS WHERE CONVERT(VARCHAR, APPT_DATE , 10) = '" & strApptDate & "' AND " _
                                & "PATIENT_ID = " & lngPatientID & " And Active = 1"
                            Set rsSched = cnSched.Execute(strSQL)
                                If Not rsSched.EOF Then
                                    If rsSched.Fields("DURATION") <> vbNullString Then udtInPatSched.Duration = rsSched.Fields("DURATION")
                                    Select Case UCase(clsApptPref.SchedInterval)
                                        Case "EVERY 20 MIN"
                                            lngRowTime = 20
                                        Case "EVERY 15 MIN"
                                            lngRowTime = 15
                                        Case "EVERY 10 MIN"
                                            lngRowTime = 10
                                        Case "EVERY 5 MIN"
                                            lngRowTime = 5
                                        Case "EVERY 30 MIN", "EVERY HALF HOUR"
                                            lngRowTime = 30
                                    End Select
                                    Arr = Split(udtInPatSched.Duration, " ")
                                    If UBound(Arr) > 0 Then
                                        If IsNumeric(Arr(0)) And lngRowTime > 0 Then strNote = Val(Arr(0) / lngRowTime)
                                    Else
                                        strNote = udtInPatSched.Duration
                                    End If
                                    If Not IsNull(rsSched.Fields("STATUS")) Then
                                        udtInPatSched.Status = rsSched.Fields("STATUS")
                                        'If strStatus <> vbNullString Then
                                            clsGridFormat.BackColor = 0
                                            Select Case IsProperCase(udtAppointment.SchedType)
                                                Case "Physician"
                                                    clsGridFormat.BackColor = GridNonPatient
                                                Case "Vendor"
                                                    clsGridFormat.BackColor = GridNonPatient
                                            End Select
                                            Select Case IsProperCase(udtInPatSched.Status)
                                                Case "Present"
                                                    clsGridFormat.BackColor = grdCurItem
                                                Case "No Show", "Cancelled"
                                                    clsGridFormat.BackColor = grdGray
                                                Case "Check Out"
                                                    clsGridFormat.BackColor = DataColor
                                                Case "Confirmed"
                                                    clsGridFormat.BackColor = GridConfir
                                            End Select
                                            Manage_Colors strStatus, lngPatientID
                                            lngCol = 1
                                            grdData.Col = 1
                                            grdData.Row = lngRow
                                            If grdData.CellTag <> vbNullString Then clsGridFormat.CellTag = grdData.CellTag
                                            If grdData.CellNote <> vbNullString Then clsGridFormat.CellNote = grdData.CellNote
                                            grdData.GetText grdData.Col, lngRow, varCellValue
                                            Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, strNote, CStr(varCellValue), "TypeHAlignCenter", clsGridFormat.CellTag, clsGridFormat.BackColor, clsGridFormat.CellNote
                                            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow + (Val(strNote) - 1), CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                        'End If
                                    End If
                                End If
                        If cnSched.State = adStateOpen Then
                            cnSched.Close
                            Set cnSched = Nothing
                        End If
                    End If
                Next lngRow
        Else
            For lngCol = 1 To .MaxCols
                .GetText lngCol, SpreadHeader + 2, vApptDate
                Arr = Split(vApptDate, "-")
                If UBound(Arr) > 0 Then
                    vApptDate = Arr(1)
                    vApptDate = Format(vApptDate, "mm/dd/yyyy")
                    If vApptDate <> vbNullString Then
                         For lngRow = 1 To .MaxRows
                            If .CellTag <> vbNullString Then lngPatientID = Val(.CellTag)
                                If lngPatientID <> 0 Then
                                    cnSched.open psConnect(0)
                                        strSQL = "SELECT STATUS FROM INPATIENT_APPOINTMENTS WHERE APPT_DATE = '" & strApptDate & "' AND " _
                                            & "PATIENT_ID = " & lngPatientID
                                        Set rsSched = cnSched.Execute(strSQL)
                                            If Not rsSched.EOF Then
                                                If Not IsNull(rsSched.Fields("STATUS")) Then Manage_Colors rsSched.Fields("STATUS"), lngPatientID
                                                If rsSched.Fields("DURATION") <> vbNullString Then udtInPatSched.Duration = rsSched.Fields("DURATION")
                                                Select Case UCase(clsApptPref.SchedInterval)
                                                    Case "EVERY 20 MIN"
                                                        lngRowTime = 20
                                                    Case "EVERY 15 MIN"
                                                        lngRowTime = 15
                                                    Case "EVERY 10 MIN"
                                                        lngRowTime = 10
                                                    Case "EVERY 5 MIN"
                                                        lngRowTime = 5
                                                    Case "EVERY 30 MIN", "EVERY HALF HOUR"
                                                        lngRowTime = 30
                                                End Select
                                                Arr = Split(udtInPatSched.Duration, " ")
                                                If UBound(Arr) > 0 Then
                                                    If IsNumeric(Arr(0)) And lngRowTime > 0 Then strNote = Val(Arr(0) / lngRowTime)
                                                Else
                                                    strNote = udtInPatSched.Duration
                                                End If
                                            End If
                                            grdData.Col = lngCol
                                            grdData.Row = lngRow
                                            If grdData.CellTag <> vbNullString Then clsGridFormat.CellTag = grdData.CellTag
                                            If grdData.CellNote <> vbNullString Then clsGridFormat.CellNote = grdData.CellNote
                                            grdData.GetText grdData.Col, lngRow, varCellValue
                                            Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, strNote, CStr(varCellValue), "TypeHAlignCenter", clsGridFormat.CellTag, Color, clsGridFormat.CellNote
                                            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow + (Val(strNote) - 1), CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    If cnSched.State = adStateOpen Then
                                        cnSched.Close
                                        Set cnSched = Nothing
                                    End If
                                End If
                        Next lngRow
                    End If
                End If
            Next lngCol
        End If
    End With
    grdData.Redraw = True
    Screen.MousePointer = vbDefault
    grdData.Redraw = False
    Exit Function
Error_Check_Schedule:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Check_Schedule", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
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
            .FontName = "Georgia"
            .FontSize = 8
            .FontBold = True
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
              .TypeVAlign = TypeVAlignCenter
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
              .FontName = "Georgia"
              .FontSize = 9
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vlngFontSize As Long)
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
                .FontName = "Georgia"
                If vlngFontSize > 0 Then
                    .FontSize = vlngFontSize
                Else
                    .FontSize = 9
                End If
                .FontBold = True
                If vlngBackColor > 0 Then .BackColor = vlngBackColor
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                    .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                ElseIf Trim(vstrSpan) <> vbNullString Then
                  lngSpan = vstrSpan
                  If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                End If
                .TypeVAlign = TypeVAlignCenter
                If vstrHAligh = "TypeHAlignRight" Then
                    .TypeHAlign = TypeHAlignRight
                ElseIf vstrHAligh = "TypeHAlignLeft" Then
                  .TypeHAlign = TypeHAlignLeft
                Else
                  .TypeHAlign = TypeHAlignCenter
                End If
                .SetText vlngCol, vlngRow, Trim(vstrText)
                .Text = Trim(vstrText)
                CellNote = .MaxTextCellHeight
                If .RowHeight(vlngRow) < CellNote Then
                  .RowHeight(vlngRow) = CellNote + 7
                End If
                If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypeEdit_RowSpan(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrNote As String)
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
              .SetText vlngCol, vlngRow, Trim(vstrText)
              .TypeEditMultiLine = True   'Make multi-lines
              .FontName = "Georgia"
              .FontSize = 9
              .FontBold = True
              If vlngBackColor > 0 Then .BackColor = vlngBackColor
              If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
              ElseIf Trim(vstrSpan) <> vbNullString Then
                lngSpan = vstrSpan
                If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, vlngCol, lngSpan
              End If
              .Text = Trim(vstrText)
              CellNote = .MaxTextCellHeight
              If .RowHeight(vlngRow) < CellNote Then
                .RowHeight(vlngRow) = CellNote
               End If
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

Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrVAlign As String, Optional lngBackColor As Long, Optional vstrCellNote As String)
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
            .FontName = "Georgia"
            .FontSize = 9
            .FontBold = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
            If vstrVAlign = "TypeVAlignTop" Then
                .TypeVAlign = TypeVAlignTop
            ElseIf vstrVAlign = "TypeVAlignBottom" Then
                .TypeVAlign = TypeVAlignBottom
            Else
               .TypeVAlign = TypeVAlignCenter
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
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
            If lngBackColor > 0 Then .BackColor = lngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        .BlockMode = False
    End With
End Function
Private Function Set_GridCurSel(vObjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, Optional vstrCellTag As String)
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
    .Redraw = True
    End With
End Function

Private Function LookUpColName(vstrName As String) As String
Dim lngRow As Long
Dim lngPrimKey As Long
Dim lngCol As Long
Dim varCellValue
Dim strRowName As String
Dim booData As Boolean
Dim strClinData As String
Dim varData
Dim CellNote As Long
Dim Arr
Dim strFldName As String
Dim strFound As String
Dim lngInsertRow As Long
Dim booExit As Boolean

    With grdData
        LookUpColName = vbNullString
        For lngCol = 1 To .MaxCols
            .Col = lngCol
            .Row = SpreadHeader + 3
            .GetText lngCol, SpreadHeader + 3, varCellValue
            If Trim(varCellValue) <> vbNullString And TextFound(vstrName, Trim(varCellValue), Len(vstrName), True) Then
                strFound = Trim(varCellValue)
                LookUpColName = lngCol & "|" & SpreadHeader + 3
                booExit = True
                Exit For
            End If
        Next lngCol
    End With
End Function

Private Sub Format_Week()
Dim Idx As Integer
Dim strSQL As String
Dim sDate As Date
Dim vDay As String
Dim sMonth As String
Dim Arr
Dim lngLoopCol As Long
Dim nDays As Long
Dim lngCol As Long
Dim lngRow As Long
Dim varCellValue
Dim FirstDay As VbDayOfWeek
Dim lngWeekDays As Long
Dim strOption As String

    On Error GoTo Error_Format_Week
    Screen.MousePointer = vbHourglass
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScheduleWeek")
    If Val(strOption) > 0 Then
        lngWeekDays = Val(strOption)
    Else
        lngWeekDays = 7
    End If
    strSelDate = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
    FirstDay = WeekDay(strSelDate, vbMonday)
    nDays = Val(FirstDay)
    Remove_InsertCells
    lngRow = SpreadHeader
    With grdData
        .Redraw = False
        .ScrollBarTrack = ScrollBarTrackVertical
        .AllowDragDrop = True
        .MaxCols = lngWeekDays
        For lngCol = 1 To .MaxCols
            .ColWidth(lngCol) = 20
        Next lngCol
        lngCol = 1
        Format_CellTypeFixed grdData, lngRow, lngCol, "1", "Monday", "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeFixed grdData, lngRow, lngCol, "1", "Tuesday", "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeFixed grdData, lngRow, lngCol, "1", "Wednesday", "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeFixed grdData, lngRow, lngCol, "1", "Thursday", "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeFixed grdData, lngRow, lngCol, "1", "Friday", "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        If lngWeekDays > 5 Then
            lngCol = 6
            Format_CellTypeFixed grdData, lngRow, lngCol, "1", "Saturday", "TypeHAlignCenter", "", HeaderColor
            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        If lngWeekDays > 6 Then
            lngCol = 7
            Format_CellTypeFixed grdData, lngRow, lngCol, "1", "Sunday", "TypeHAlignCenter", "", HeaderColor
            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        lngCol = nDays
        For Idx = 1 To grdData.MaxCols
            If Idx = lngCol Then
                clsGridFormat.Col = Idx
                sDate = strSelDate
                .GetText Idx, lngRow, varCellValue
                clsGridFormat.Text = Trim(varCellValue) & " - " & Format(sDate, "mm/dd/yyyy")
                Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
                grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            ElseIf Idx > lngCol Then
                clsGridFormat.Col = Idx
                nDays = Idx - lngCol
              '  ndays = Idx - 1
                sDate = DateAdd("d", nDays, strSelDate)
                .GetText Idx, lngRow, varCellValue
                clsGridFormat.Text = Trim(varCellValue) & " - " & Format(sDate, "mm/dd/yyyy")
                Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
                grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            ElseIf Idx < lngCol Then
                clsGridFormat.Col = Idx
                nDays = nDays - 1
                sDate = DateAdd("d", (-nDays), strSelDate)
                .GetText Idx, lngRow, varCellValue
                clsGridFormat.Text = Trim(varCellValue) & " - " & Format(sDate, "mm/dd/yyyy")
                Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
                grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        Next Idx
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Week:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Week", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Function Search_Patient()
Dim strSQL As String
Dim strSearch As String
Dim Cnt As Long
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngPatID As Long
Dim strTitle As String
Dim lngInvID As Long
Dim strDupNo As String
Dim lngRow As Long
Dim lngAcctID As Long
Dim strFirst As String
Dim Arr
Dim lngFac As Long

    On Error GoTo Error_Search_Patient
    Screen.MousePointer = vbHourglass
    Clear_UdtPatient
    If cctSearch.Text = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    lstListBox.Clear
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    If Trim(cctSearch.Text) = vbNullString Then
        MsgBox "Please enter a search term first!", vbInformation, "PatientTrac Patient Search"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    strSearch = cctSearch.Text
    If cboSearchFilter.Text = "FINANCIAL ACCT NO" Then
        strTitle = "EXTERNAL ID"
    Else
        strTitle = UCase(cboSearchFilter.Text)
    End If
'    If clsEDI_Client.Site_ID = "1999" Then
'        Select Case Trim(strTitle)
'            Case "LAST NAME"
'                If TextFound(",", strSearch, Len(","), True) Then
'                    Arr = Split(strSearch, ",")
'                    If UBound(Arr) > 0 Then
'                        strFirst = Trim(Arr(1))
'                        strSearch = Arr(0)
'                        strFirst = Replace(strFirst, " ", "")
'                        strSearch = Replace(strSearch, " ", "")
'                        strSearch = Trim(strSearch)
'                    End If
'                End If
'                strFirst = Replace(strFirst, "'", "''")
'                If Trim(strFirst) <> vbNullString Then
'                     strSQL = "Search_Patient_FirstName_Last_Name_Join   '" & strSearch & "%', '" & strFirst & "%'," & udtInPatSched.Facility_ID
'                Else
'                    strSQL = "Search_Patient_Row_Count_Last_Name_Join 25, '" & strSearch & "%'," & udtInPatSched.Facility_ID
'                End If
'            Case "EMR ID"
'                strSQL = "Search_Patient_EMR_ID_Join '" & strSearch & "%'," & udtInPatSched.Facility_ID
'            Case "SOCIAL SECURITY"
'                strSQL = "Search_Patient_Social_Security_Join '" & strSearch & "%'," & udtInPatSched.Facility_ID
'            Case "EXTERNAL ID"
'                strSQL = "Search_Patient_RowCount_External_ID_Join 25,'" & strSearch & "%'," & udtInPatSched.Facility_ID
'            Case "INVOICE NO"
'                If Len(strSearch) > 3 Then
'                    lngInvID = Val(strSearch)
'                    strSQL = "Search_Patient_Invoice_Invoice_No_Join '" & strSearch & "%'," & udtInPatSched.Facility_ID
'                    Set rsPat = cnPat.Execute(strSQL)
'                    If Not rsPat.EOF Then
'                        lngPatID = rsPat.Fields("ACCOUNT_NO")
'                    Else
'                        MsgBox "This is an invalid Invoice Number! Please select valid Invoice.", vbInformation, "PatientTrac EDI Submission"
'                        cctSearch.Text = vbNullString
'                        Screen.MousePointer = vbDefault
'                        If cnPat.State = adStateOpen Then
'                            cnPat.Close
'                            Set rsPat = Nothing
'                            Set cnPat = Nothing
'                        End If
'                        Exit Function
'                    End If
'                End If
'            End Select
'    Else
        Select Case Trim(strTitle)
            Case "LAST NAME"
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
                strFirst = Replace(strFirst, "'", "''")
                If Trim(strFirst) <> vbNullString Then
    '                 strSQL = "Search_InPatient_FirstName_Last_Name   '" & strSearch & "%', '" & strFirst & "%', " & udtInPatSched.Facility_ID
                     strSQL = "Search_Patient_FirstName_Last_Name   '" & strSearch & "%', '" & strFirst & "%'"
                Else
    '                strSQL = "Search_InPatient_Row_Count_Last_Name 25, '" & strSearch & "%', " & udtInPatSched.Facility_ID
                    strSQL = "Search_Patient_Row_Count_Last_Name 25, '" & strSearch & "%'"
                End If
            Case "EMR ID"
                strSQL = "Search_Patient_EMR_ID '" & strSearch & "%'"
            Case "SOCIAL SECURITY"
                strSQL = "Search_Patient_Social_Security '" & strSearch & "%'"
            Case "EXTERNAL ID"
                strSQL = "Search_Patient_RowCount_External_ID 25,'" & strSearch & "%'"
            Case "INVOICE NO"
                If Len(strSearch) > 3 Then
                    lngInvID = Val(strSearch)
                    strSQL = "Search_Patient_Invoice_Invoice_No '" & strSearch & "%'"
                    Set rsPat = cnPat.Execute(strSQL)
                    If Not rsPat.EOF Then
                        lngPatID = rsPat.Fields("ACCOUNT_NO")
                    Else
                        MsgBox "This is an invalid Invoice Number! Please select valid Invoice.", vbInformation, "PatientTrac EDI Submission"
                        cctSearch.Text = vbNullString
                        Screen.MousePointer = vbDefault
                        If cnPat.State = adStateOpen Then
                            cnPat.Close
                            Set rsPat = Nothing
                            Set cnPat = Nothing
                        End If
                        Exit Function
                    End If
                End If
            End Select
'        End If
        If strSQL <> vbNullString Then
            Set rsPat = cnPat.Execute(strSQL)
            Do Until rsPat.EOF
                Clear_UdtPatient
                lngInvID = 0
                Set Flds = rsPat.Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "ACCT_ID"
                                    lngAcctID = Fld.Value
                                Case "INVOICE_NO"
                                    lngInvID = Fld.Value
                                Case "FIRST_NAME"
                                    udtPatientName.FirstName = IsProperCase(Fld.Value)
                                Case "MIDDLE_NAME"
                                    udtPatientName.MiddleName = IsProperCase(Fld.Value)
                                Case "LAST_NAME"
                                    udtPatientName.LastName = IsProperCase(Fld.Value)
                                    If TextFound("''", udtPatientName.LastName, 2, True) Then
                                        udtPatientName.LastName = Replace(udtPatientName.LastName, "''", "'")
                                        udtPatientName.LastName = IsProperCase(udtPatientName.LastName)
                                    Else
                                        udtPatientName.LastName = IsProperCase(udtPatientName.LastName)
                                    End If
                                Case "PREF_CONTACT"
                                    udtPatientName.PrefContact = Fld.Value
                                Case "GENDER"
                                    udtPatientName.Gender = IsProperCase(Fld.Value)
                                    If Left(IsProperCase(udtPatientName.Gender), 1) = "F" Then
                                        udtPatientName.Gender = "Female"
                                        udtPatientName.Person = "She"
                                        udtPatientName.Possessive = "Her"
                                        udtPatientName.PerPosses = "Herself"
                                    ElseIf Left(IsProperCase(udtPatientName.Gender), 1) = "M" Then
                                        udtPatientName.Gender = "Male"
                                        udtPatientName.Person = "He"
                                        udtPatientName.Possessive = "His"
                                        udtPatientName.PerPosses = "Himself"
                                    End If
                                Case "TITLE"
                                    udtPatientName.Title = IsProperCase(Fld.Value)

                                Case "MARITAL_STATUS"
                                    udtPatientName.Marital_Status = IsProperCase(Fld.Value)
                                Case "SUFFIX"
                                    udtPatientName.Suffix = Fld.Value
                                    udtPatientName.Suffix = Replace(udtPatientName.Suffix, ",", "")
                                    If TextFound("Jr", udtPatientName.Suffix, 2, True) Or _
                                        TextFound("Sr", udtPatientName.Suffix, 2, True) Then
                                        udtPatientName.Suffix = IsProperCase(udtPatientName.Suffix)
                                    Else
                                        udtPatientName.Suffix = UCase(udtPatientName.Suffix)
                                    End If
                                    If TextFound("PHD", udtPatientName.Suffix, Len("phd"), True) Or _
                                        TextFound("PHD", udtPatientName.Suffix, 4, True) Then
                                        udtPatientName.Suffix = "Ph.D."
                                    End If
                                Case "PATIENT_EXT_REF"
                                    udtPatientName.ExtID = Fld.Value
                                Case "PATIENT_ID"
                                    udtPatientName.IntId = Fld.Value
                                    lngPatID = udtPatientName.IntId
                                Case "SOCIAL_SECURITY"
                                    udtPatientName.SS = Fld.Value
                                Case "ROOM_NUMBER"
                                    udtPatientName.RoomNumber = Fld.Value
                                Case "COUNTRY"
                                     udtPatientName.Country = Fld.Value
                                Case "CELL_PHONE"
                                     udtPatientName.CellPhone = Fld.Value
                                Case "EMAIL"
                                     udtPatientName.Email = LCase(Fld.Value)
                                Case "COUNTY"
                                     udtPatientName.County = LCase(Fld.Value)
                                Case "PATIENT_TYPE"
                                     udtPatientName.PatientType = LCase(Fld.Value)
                                Case "DRV_LICENSE"
                                     udtPatientName.DL = LCase(Fld.Value)
                                Case "WORK_PHONE"
                                     udtPatientName.WorkPhone = LCase(Fld.Value)
                                Case "LOCAL_INTERNATL_PHONE"
                                     udtPatientName.InterNalPhone = LCase(Fld.Value)
                                Case "PHONE"
                                     udtPatientName.Phone = LCase(Fld.Value)
                                Case "AREA_CODE"
                                    udtPatientName.AreaCode = LCase(Fld.Value)
                                Case "STATUS"
                                     udtPatientName.Status = LCase(Fld.Value)
                            End Select
                        End If
                    End If
                Next
                rsPat.MoveNext
                If Val(strDupNo) <> Val(udtPatientName.IntId) Then
                    If Left(IsProperCase(udtPatientName.Gender), 1) = "M" Then
                        udtPatientName.Title = "Mr."
                    ElseIf Left(IsProperCase(udtPatientName.Gender), 1) = "F" Then
                        If IsProperCase(udtPatientName.Marital_Status) = "Married" Then
                            udtPatientName.Title = "Mrs."
                        Else
                            udtPatientName.Title = "Ms."
                        End If
                    End If
                    If Trim(udtPatientName.Suffix) <> vbNullString Then
                        udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                    Else
                        udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                    End If
                    Select Case Trim(strTitle)
                        Case "EXTERNAL ID"
                            lblEnterInfo.Caption = "Last Name, First Name - External Reference"
                        Case "EMR ID"
                            lblEnterInfo.Caption = "Last Name, First Name - EMR - ID"
                        Case "INVOICE NO", "LAST_NAME"
                            lblEnterInfo.Caption = "Last Name, First Name"
                    End Select
                    If Trim(udtPatientName.Name) <> vbNullString Then
                         If clsEDI_Client.Site_ID = "1999" Then
                            mbooByPass = True
                            lngFac = clsFacStatus.FacilityID
                            Read_Patient_Status Val(Val(udtPatientName.IntId))
                            If clsFacStatus.FacilityID = 0 Then
                                clsSurgFac.Name = "No Facility"
                            Else
                                Get_SurgFac "", clsFacStatus.FacilityID
                            End If
                            clsFacStatus.FacilityID = lngFac
                            lstListBox.AddItem Trim(udtPatientName.Name) & "-" & clsSurgFac.Name
                            lstListBox.ItemData(lstListBox.NewIndex) = lngPatID
                        Else
                            lstListBox.AddItem udtPatientName.LastName & ", " & udtPatientName.FirstName
                            lstListBox.ItemData(lstListBox.NewIndex) = Val(udtPatientName.IntId)
                        End If
                    End If
                End If
                strDupNo = udtPatientName.IntId
            Loop
            udtPatientName.IntId = lngPatID
        End If
    Clear_UdtPatient
    mbooByPass = False
    If lstListBox.listcount <= 0 Then
        If PsUserID <> vbNullString Then
            If TextFound("Supp", PsUserID, Len("Supp"), True) Then
                mbooSupport = True
                PT_Support_Search
            ElseIf Len(strSearch) > 3 Then
                PT_Support_Search
            End If
        Else
            PsUserID = Get_PsUserID(psConnect(0))
            If TextFound("Supp", PsUserID, Len("Supp"), True) Then
                mbooSupport = True
                PT_Support_Search
            ElseIf Len(strSearch) > 3 Then
                PT_Support_Search
            End If
        End If
    End If
    If cnPat.State = adStateOpen Then
        cnPat.Close
        Set cnPat = Nothing
        Set rsPat = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Search_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Search_Patient", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Update_EnctrEndDate(vstrType As String)
Dim strSQL As String
Dim cnData As New ADODB.Connection
Dim rsData As New ADODB.Recordset
Dim lngEnctrID As Long
Dim sDate As String


    On Error GoTo Error_Update_EnctrEndDate
    Screen.MousePointer = vbHourglass
    If cnData.State = adStateClosed Then cnData.open psConnect(0)
    If udtAppointment.SchedTime <> vbNullString Then udtEnctr.EndDate = Format(udtAppointment.SchedTime, "mm/dd/yyyy hh:mm:ss AM/PM")
    
    strSQL = "Get_Current_Encounter_InPatient " & frmPatientAppointment.mlngPatientID
    Set rsData = cnData.Execute(strSQL)
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    With rsData
        Do Until .EOF
            If Not IsNull((!VISIT_TYPE)) Then
                If (!VISIT_TYPE) = vstrType Then
                    lngEnctrID = (!Encounter_ID)
                End If
            End If
            .MoveNext
        Loop
    End With
    strSQL = "UPDATE ENCOUNTER SET ENCNTR_END_DATE = '" & udtEnctr.EndDate & "', UPDATE_DATE = '" & udtEnctr.EndDate & "' WHERE ENCOUNTER_ID = " & lngEnctrID
    Set rsData = cnData.Execute(strSQL)
    If cnData.State = adStateOpen Then
        cnData.Close
        Set cnData = Nothing
        Set rsData = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_EnctrEndDate:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_EnctrEndDate", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Update_InPatient_Status()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngPurge As Double
Dim lngAdmit_ID As Long
Dim lngPatient As Long

    On Error GoTo Error_Update_InPatient_Status
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    udtInPatSched.Appointment_ID = 0
    lngRow = grdData.ActiveRow
    With grdData
        .Col = 1
        .Row = lngRow
        lngPatient = .GetRowItemData(lngRow)
    End With
    strSQL = "Get_InPatient_Appt_ID " & lngPatient & ",'" & Format(udtInPatSched.Appt_Date, "mm/dd/yyyy") & "'," & udtInPatSched.Provider_ID
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        If Not .EOF Then
            udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
        End If
    End With
    If udtInPatSched.Appointment_ID > 0 Then
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
        strSQL = "Update_InPatient_Status_Appt_ID  " & udtInPatSched.Appointment_ID & ", '" & udtInPatSched.Status & "'"
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
    End If
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_InPatient_Status:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_InPatient_Status", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Update_Patient_Status()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim booSkip As Boolean
Dim lngAdmit_ID As Long

    On Error GoTo Error_Update_Patient_Status
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    For lngRow = 1 To grdPatient.MaxRows
        udtInPatSched.IntId = vbNullString
        udtInPatSched.Status = vbNullString
        udtInPatSched.RoomNumber = vbNullString
        lngAdmit_ID = 0
        booSkip = False
        grdPatient.Col = 1
        grdPatient.Row = lngRow
        udtInPatSched.IntId = grdPatient.GetRowItemData(lngRow)
'        If grdPatient.CellTag <> vbNullString Then
'            udtInPatSched.IntId = grdPatient.CellTag
'        End If
        grdPatient.GetText 4, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtInPatSched.Status = Trim(varCellValue)
        End If
        grdPatient.GetText 2, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtInPatSched.RoomNumber = Trim(varCellValue)
        End If
        If udtInPatSched.Status = vbNullString Then udtInPatSched.Status = "Active"
        If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
        Select Case IsProperCase(udtInPatSched.Status)
            Case "Active", "Deceased"
                booSkip = True
            Case "Inactive"
                booSkip = False
        End Select
        If Val(udtInPatSched.IntId) > 0 And Not booSkip Then
            strSQL = "Update_Patient_Purge_Status " & Val(udtInPatSched.IntId) & ",'Active'"
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
        ElseIf Val(udtInPatSched.IntId) > 0 And Trim(udtInPatSched.RoomNumber) <> vbNullString And booSkip Then
            strSQL = "Get_InPatient_Room " & Val(udtInPatSched.IntId) & "," & udtInPatSched.Facility_ID
                Set rsReadAppt = cnReadAppt.Execute(strSQL)
                With rsReadAppt
                    If Not .EOF Then
                        lngAdmit_ID = (!PATIENT_ADMIT_ROOMID)
                    End If
                End With
                If lngAdmit_ID > 0 Then
                    strSQL = "Update_InPatient_Room '" & udtInPatSched.RoomNumber & "'," & lngAdmit_ID
                    Set rsReadAppt = cnReadAppt.Execute(strSQL)
                Else
                    strSQL = "Insert_InPatient_Room_Patient " & Val(udtInPatSched.IntId) & ", " & udtInPatSched.Facility_ID & ",'" & udtInPatSched.RoomNumber & "'"
                    Set rsReadAppt = cnReadAppt.Execute(strSQL)
                End If
        End If
    Next lngRow
    MsgBox "Patient Roster has been successfully updated", vbInformation, "PatientTrac InPatient Roster"
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_Patient_Status:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_Patient_Status", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Validate_CurrentFacility(vlngPatient As Long) As Boolean
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim lngPatient As Long


    On Error GoTo Error_Validate_CurrentFacility
    Screen.MousePointer = vbHourglass
    lngPatient = vlngPatient
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Get_Patient_Status_PatientID " & lngPatient
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        If Not .EOF Then
            If Not IsNull(!STATUS_ITEM) Then
                If IsNumeric((!STATUS_ITEM)) Then
                    If udtInPatSched.Facility_ID <> (!STATUS_ITEM) Then
                        Validate_CurrentFacility = True
                        Screen.MousePointer = vbDefault
                        Exit Function
                    End If
                End If
            End If
        End If
    End With
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_CurrentFacility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_CurrentFacility", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Validate_MD_Provider(Optional vlngProv As Long) As Boolean
Dim rsProvider As New ADODB.Recordset
Dim cnProvider As New ADODB.Connection
Dim strSQL As String
Dim lngProv As Long
Dim Idx%
Dim lngFac As Long

    On Error GoTo Error_Validate_MD_Provider
    Validate_MD_Provider = False
    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        lngFac = .GetRowItemData(.ActiveRow)
    End With
    If cnProvider.State = adStateClosed Then cnProvider.open psConnect(0)
    strSQL = "Validate_Extender_Providers_Facility " & mlngLogonUser & "," & lngFac
    Set rsProvider = cnProvider.Execute(strSQL)
    With rsProvider
        If Not rsProvider.EOF Then
            If Not IsNull((!Provider_Type)) Then
                If UCase((!Provider_Type)) = "PHYSICIAN" Then
                    Validate_MD_Provider = True
                End If
            End If
        End If
    End With
    If cnProvider.State = adStateOpen Then
        cnProvider.Close
        Set cnProvider = Nothing
        Set rsProvider = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_MD_Provider:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_MD_Provider", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Validate_Extender_Provider() As Boolean
Dim rsProvider As New ADODB.Recordset
Dim cnProvider As New ADODB.Connection
Dim strSQL As String
Dim lngProv As Long
Dim Idx%
Dim lngFac As Long

    On Error GoTo Error_Validate_Extender_Provider
    Validate_Extender_Provider = False
    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        lngFac = .GetRowItemData(.ActiveRow)
    End With
    If cnProvider.State = adStateClosed Then cnProvider.open psConnect(0)
    strSQL = "Validate_Extender_Providers_Facility " & mlngLogonUser & "," & lngFac
    Set rsProvider = cnProvider.Execute(strSQL)
    With rsProvider
        If Not rsProvider.EOF Then
            If Not IsNull((!Provider_Type)) Then
                If UCase((!Provider_Type)) = "CASE MANAGER" Then
                    Validate_Extender_Provider = True
                End If
            End If
        End If
    End With
    If Validate_Extender_Provider Then
        cboApptType.Visible = False
        lblApptType.Visible = False
        cboApptType.Clear
    End If
    If cnProvider.State = adStateOpen Then
        cnProvider.Close
        Set cnProvider = Nothing
        Set rsProvider = Nothing
    End If
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_Extender_Provider:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_Extender_Provider", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Function Validate_Patient_Status(vlngPatientID As Long)
Dim strSQL As String
Dim strEnctrStart As String
Dim cnPatient As New ADODB.Connection
Dim rsPatient As New ADODB.Recordset
Dim lngPatientID As Long
Dim strTime As String
Dim strStart As String
Dim strEnd As String

    On Error GoTo Error_Validate_Patient_Status
    Screen.MousePointer = vbHourglass
    lngPatientID = vlngPatientID
    If lngPatientID > 0 Then
        If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
        strSQL = "Get_Patient_Status_PatientID " & lngPatientID
            Set rsPatient = cnPatient.Execute(strSQL)
            With rsPatient
                If Not .EOF Then
                    If Not IsNull(!STATUS_ITEM) Then
                        If IsNumeric((!STATUS_ITEM)) Then
                            If (!STATUS_ITEM) = udtInPatSched.Facility_ID Then
                                Validate_Patient_Status = True
                            End If
                        End If
                    End If
                End If
            End With
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_Patient_Status:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_Patient_Status", App.EXEName & "." & "modBilling"
    Exit Function
    Resume
End Function

Private Function Write_Color(vstrColorName As String, vlngColor As Long)
Dim cnColor As New ADODB.Connection
Dim rsColor As New ADODB.Recordset
Dim strSQL As String
Dim lngPrimKey As Long

    On Error GoTo Error_Write_Color
    

    If cnColor.State = adStateClosed Then cnColor.open psConnect(1)
    strSQL = "SELECT * FROM APPT_COLOR WHERE APPT_TYPE = '" & vstrColorName & "' AND APPT_COLOR = " & vlngColor & ""
    Set rsColor = cnColor.Execute(strSQL)
    If Not rsColor.EOF Then
        lngPrimKey = rsColor.Fields("COLOR_ID")
    End If

    If lngPrimKey > 0 Then
        strSQL = "UPDATE APPT_COLOR SET APPT_TYPE = '" & vstrColorName & "'," _
            & "APPT_COLOR = " & vlngColor & " WHERE lngPrimkey = " & lngPrimKey
         Set rsColor = cnColor.Execute(strSQL)
    Else
        strSQL = "INSERT INTO APPT_COLOR (APPT_TYPE,APPT_COLOR)" _
            & "VALUES ('" & vstrColorName & "'," & vlngColor & ")"
        Set rsColor = cnColor.Execute(strSQL)
    End If
    If cnColor.State = adStateOpen Then
        cnColor.Close
        Set cnColor = Nothing
        Set rsColor = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Color:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Color", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Write_InPatient()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngPurge As Double
Dim lngAdmit_ID As Long
Dim strReAdmit As String
Dim lngCnt As Long
Dim lngWrite As Long
Dim strNewDate As String

    On Error GoTo Error_Write_InPatient
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    Get_Appt_Range
    If mstrRangeDate = vbNullString Then
        Get_New_Appt_Range
    End If
    strNewDate = DateAdd("d", 1, mstrRangeDate)
    udtInPatSched.Appt_Date = mstrUpperRange
    udtInPatSched.Appt_Date = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
    Get_Facilities udtInPatSched.Facility_ID
    For lngRow = 1 To grdPatient.MaxRows
        strReAdmit = vbNullString
        Clear_InSched
        grdPatient.Col = 1
        grdPatient.Row = lngRow
        udtInPatSched.IntId = grdPatient.GetRowItemData(lngRow)
        grdPatient.GetText 2, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtInPatSched.RoomNumber = Trim(varCellValue)
        End If
        grdPatient.GetText 3, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtInPatSched.PatientType = Trim(varCellValue)
        End If
        grdPatient.GetText 4, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtInPatSched.Status = Trim(varCellValue)
        End If
        grdPatient.GetText 5, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtInPatSched.ApptComment = Trim(varCellValue)
        End If
        Get_Insurance Val(udtInPatSched.IntId)
        If TextFound(Trim(clsInsurance.Prim_Insurer), Trim(udtInPatSched.ApptComment), Len(Trim(clsInsurance.Prim_Insurer)), True) Then
             udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, Trim(clsInsurance.Prim_Insurer), "")
        ElseIf TextFound("No record for primary insurance", udtInPatSched.ApptComment, Len("No record for primary insurance"), True) Then
             udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "No record for primary insurance", "")
        End If
        udtInPatSched.ApptComment = Trim(udtInPatSched.ApptComment)
        udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "'", "''")
        lngCnt = lngCnt + 1
        If IsProperCase(udtInPatSched.Status) = "Active" Then
            If udtInPatSched.PatientType = "Follow-Up" Then
                udtInPatSched.PatientType = "Subsequent Care"
            ElseIf udtInPatSched.PatientType = "Initial" Then
                 udtInPatSched.PatientType = "Initial InPatient"
            ElseIf udtInPatSched.PatientType = "Re-Admission" Then
                 udtInPatSched.PatientType = "Initial InPatient"
                 strReAdmit = "Re-Admission"
            ElseIf udtInPatSched.PatientType = vbNullString Then
                 udtInPatSched.PatientType = "Subsequent Care"
            End If
            If clsFacility.pos = "31" Then
                If udtInPatSched.PatientType = "Subsequent Care" Then
                     udtInPatSched.VisitType = "Subsequent care, nursing facility E&M, limited"
                ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                     udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
                End If
            ElseIf clsFacility.pos = "13" Then
                If udtInPatSched.PatientType = "Subsequent Care" Then
                     udtInPatSched.VisitType = "Follow-Up visit, asst living facility, limited"
                ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                     udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
                End If
            End If
            lngWrite = lngWrite + 1
            If clsEDI_Client.Site_ID = "1999" Then
                udtInPatSched.Provider_ID = 1000
            End If
           udtInPatSched.Appointment_ID = 0
            If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
            strSQL = "SELECT * FROM INPATIENT_APPOINTMENTS  " _
                & "WHERE PATIENT_ID =  " & Val(udtInPatSched.IntId) & " AND " _
                & "FACILITY_ID = " & udtInPatSched.Facility_ID & " AND " _
                & "(STATUS <> 'PRESENT' AND STATUS <> 'Cancelled') AND CONVERT( VARCHAR(10), INPATIENT_APPOINTMENTS.APPT_DATE, 101)  BETWEEN '" & Format(strNewDate, "mm/dd/yyyy") & "' AND  '" & Format(mstrUpperRange, "mm/dd/yyyy") & "' ORDER BY INSERT_DATE DESC"
                Set rsReadAppt = cnReadAppt.Execute(strSQL)
                With rsReadAppt
                    If Not .EOF Then
                        If Not IsNull((!Appt_Date)) Then
                            If Trim((!Appt_Date)) <> vbNullString Then
                                udtInPatSched.Appt_Date = Format((!Appt_Date), "mm/dd/yyyy")
                            End If
                           udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
                        End If
                    End If
                End With
                If udtInPatSched.Appt_Date = vbNullString Then
                    udtInPatSched.Appt_Date = Format(Now, "mm/dd/yyyy")
                Else
                    udtInPatSched.Appt_Date = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
                End If
            If udtInPatSched.Appointment_ID > 0 Then
                lngPurge = 0
                strSQL = "Update_InPatient_Appt '" & udtInPatSched.Appt_Date & "','" & udtInPatSched.RoomNumber & "','" _
                    & udtInPatSched.VisitType & "','" & udtInPatSched.PatientType & "','" _
                    & IsProperCase(udtInPatSched.Status) & "', " & lngPurge & ", '" & Fix_Comments(udtInPatSched.ApptComment) & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," _
                    & udtInPatSched.Appointment_ID
                Set rsReadAppt = cnReadAppt.Execute(strSQL)
                If strReAdmit <> vbNullString Then
                    strSQL = "UPDATE INPATIENT_APPOINTMENTS SET CURRENT_LOCATION = 'Re-Admission' WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
                    Set rsReadAppt = cnReadAppt.Execute(strSQL)
                End If
            Else
                strSQL = "Insert_InPatient_Appt " & Val(udtInPatSched.IntId) & ", " _
                    & udtInPatSched.Provider_ID & "," & udtInPatSched.Facility_ID & ",'" & udtInPatSched.RoomNumber & "','" _
                    & udtInPatSched.Appt_Date & "', '" & udtInPatSched.VisitType & "', '" & udtInPatSched.PatientType & "','" & IsProperCase(udtInPatSched.Status) & "','" _
                    & Fix_Comments(udtInPatSched.ApptComment) & "'," & mlngLogonUser
                Set rsReadAppt = cnReadAppt.Execute(strSQL)
                With rsReadAppt
                    If Not .EOF Then
                        udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
                    End If
                End With
                If strReAdmit <> vbNullString And udtInPatSched.Appointment_ID > 0 Then
                    strSQL = "UPDATE INPATIENT_APPOINTMENTS SET CURRENT_LOCATION = 'Re-Admission' WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
                    Set rsReadAppt = cnReadAppt.Execute(strSQL)
                End If
            End If
            lngAdmit_ID = 0
            strSQL = "Get_InPatient_Room_Appt " & Val(udtInPatSched.IntId) & "," & udtInPatSched.Facility_ID & "," & udtInPatSched.Appointment_ID
                Set rsReadAppt = cnReadAppt.Execute(strSQL)
                With rsReadAppt
                    If Not .EOF Then
                        lngAdmit_ID = (!PATIENT_ADMIT_ROOMID)
                    End If
                End With
                If lngAdmit_ID > 0 Then
                    strSQL = "Update_InPatient_Room '" & udtInPatSched.RoomNumber & "'," & lngAdmit_ID
                    Set rsReadAppt = cnReadAppt.Execute(strSQL)
                Else
                    strSQL = "Insert_InPatient_Room " & Val(udtInPatSched.IntId) & ", " & udtInPatSched.Facility_ID & "," & udtInPatSched.Appointment_ID & ",'" & udtInPatSched.RoomNumber & "'"
                    Set rsReadAppt = cnReadAppt.Execute(strSQL)
                End If
        End If
    Next lngRow
    MsgBox "Visit roster has been successfully updated.", vbInformation, "PatientTrac InPatient Scheduling"
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_InPatient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_InPatient", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Update_InPatient_Appt(vlngPatient As Long)
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngPurge As Double
Dim lngAdmit_ID As Long
Dim lngPatient As Long
Dim strDate As String
Dim lngDateDiff As Long
Dim strReAdmit As String
Dim strStatus As String
Dim lngDup As Long
Dim lngCnt As Long

    On Error GoTo Error_Update_InPatient_Appt
    lngPatient = vlngPatient
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    lngRow = grdData.ActiveRow
    Clear_InSched
    grdData.Col = 1
    grdData.Row = lngRow
    udtInPatSched.IntId = grdData.GetRowItemData(lngRow)
'    If grdData.CellNote <> vbNullString Then
'        udtInPatSched.IntId = Val(grdData.CellNote)
'    End If
    grdData.GetText 1, lngRow, varCellValue
    udtInPatSched.RoomNumber = varCellValue
    grdData.GetText 5, lngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        Select Case CStr(varCellValue)
            Case "Follow-Up"
                udtInPatSched.PatientType = "Subsequent Care"
            Case "Initial"
                 udtInPatSched.PatientType = "Initial InPatient"
            Case "Re-Admission"
                 udtInPatSched.PatientType = "Initial InPatient"
                 strReAdmit = "Re-Admission"
        End Select
    End If
    grdData.GetText 4, lngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        udtInPatSched.Status = Trim(varCellValue)
    End If
    grdData.GetText 6, lngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        udtInPatSched.ApptComment = Trim(varCellValue)
    End If
    If clsFacility.pos = "31" Then
        If udtInPatSched.PatientType = "Subsequent Care" Then
             udtInPatSched.VisitType = "Subsequent care, nursing facility E&M, limited"
        ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
             udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
        End If
    ElseIf clsFacility.pos = "13" Then
        If udtInPatSched.PatientType = "Subsequent Care" Then
             udtInPatSched.VisitType = "Follow-Up visit, asst living facility, limited"
        ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
             udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
        End If
    End If
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "SELECT * FROM INPATIENT_APPOINTMENTS  " _
        & "WHERE PATIENT_ID =  " & Val(udtInPatSched.IntId) & " AND " _
        & "FACILITY_ID = " & udtInPatSched.Facility_ID & " AND CONVERT( VARCHAR(10), APPT_DATE, 101)  BETWEEN '" & Format(mstrRangeDate, "mm/dd/yyyy") & "' AND  '" & Format(mstrUpperRange, "mm/dd/yyyy") & "' ORDER BY APPT_DATE DESC, STATUS"
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            strStatus = vbNullString
            If Not IsNull((!Appt_Date)) Then
                If Trim((!Appt_Date)) <> vbNullString Then
                    strDate = Format((!Appt_Date), "mm/dd/yyyy")
                End If
                If Not IsNull((!Status)) Then
                    strStatus = (!Status)
                End If
                If strStatus <> udtInPatSched.Status Then
                    udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
                Else
                    lngDup = (!INPATIENT_APPT_ID)
                End If
            End If
            .MoveNext
            lngCnt = lngCnt + 1
        Loop
    End With
    If lngDup > 0 And lngCnt > 1 Then
        Delete_Dup_Appt lngDup
    ElseIf lngDup > 0 And udtInPatSched.Appointment_ID = 0 Then
        udtInPatSched.Appointment_ID = lngDup
    End If
    
    lngDateDiff = DateDiff("d", Format(strDate, "mm/dd/yyyy"), Format(mstrUpperRange, "mm/dd/yyyy"))
    If lngDateDiff > 0 And lngDateDiff < 10 Then
        udtInPatSched.Appt_Date = Format(strDate, "mm/dd/yyyy")
    Else
        udtInPatSched.Appt_Date = Format(mstrUpperRange, "mm/dd/yyyy")
    End If
    If udtInPatSched.VisitType = vbNullString Then udtInPatSched.VisitType = "Subsequent Visit"
'    strSQL = "Get_InPatient_Appt_ID " & Val(lngPatient) & ",'" & Format(udtInPatSched.Appt_Date, "mm/dd/yyyy") & "'," & udtInPatSched.Provider_ID
'        Set rsReadAppt = cnReadAppt.Execute(strSQL)
'        With rsReadAppt
'            If Not .EOF Then
'                udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
'            End If
'        End With
    If udtInPatSched.ApptComment = "Review Staff Notes" Then udtInPatSched.ApptComment = vbNullString
    If clsFacility.pos = "31" Then
        If udtInPatSched.PatientType = "Subsequent Care" Then
             udtInPatSched.VisitType = "Subsequent care, nursing facility E&M, limited"
        ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
             udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
        End If
    ElseIf clsFacility.pos = "13" Then
        If udtInPatSched.PatientType = "Subsequent Care" Then
             udtInPatSched.VisitType = "Follow-Up visit, asst living facility, limited"
        ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
             udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
        End If
    End If
    If udtInPatSched.Appointment_ID > 0 Then
        If IsProperCase(udtInPatSched.Status) = "Cancelled" Or IsProperCase(udtInPatSched.Status) = "Rescheduled" Then
            lngPurge = 1
        Else
            lngPurge = 0
        End If
        strSQL = "Update_InPatient_Appt '" & udtInPatSched.Appt_Date & "','" & udtInPatSched.RoomNumber & "','" & udtInPatSched.VisitType & "','" & udtInPatSched.PatientType & "','" _
            & IsProperCase(udtInPatSched.Status) & "', " & lngPurge & ", '" & Fix_Comments(udtInPatSched.ApptComment) & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," _
            & udtInPatSched.Appointment_ID
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        If strReAdmit <> vbNullString Then
            strSQL = "UPDATE INPATIENT_APPOINTMENTS SET CURRENT_LOCATION = 'Re-Admission' WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
        End If
    Else
        strSQL = "Insert_InPatient_Appt " & Val(udtInPatSched.IntId) & ", " _
            & udtInPatSched.Provider_ID & "," & udtInPatSched.Facility_ID & ",'" & udtInPatSched.RoomNumber & "','" _
            & udtInPatSched.Appt_Date & "', '" & udtInPatSched.VisitType & "','" & udtInPatSched.PatientType & "','" & IsProperCase(udtInPatSched.Status) & "','" _
            & Fix_Comments(udtInPatSched.ApptComment) & "'," & mlngLogonUser
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
            With rsReadAppt
                If Not .EOF Then
                    udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
                End If
            End With
            If strReAdmit <> vbNullString Then
                strSQL = "UPDATE INPATIENT_APPOINTMENTS SET CURRENT_LOCATION = 'Re-Admission' WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
                Set rsReadAppt = cnReadAppt.Execute(strSQL)
            End If
    End If
    strSQL = "Get_InPatient_Room " & Val(udtInPatSched.IntId) & "," & udtInPatSched.Facility_ID
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        With rsReadAppt
            If Not .EOF Then
                lngAdmit_ID = (!PATIENT_ADMIT_ROOMID)
            End If
        End With
        If lngAdmit_ID > 0 Then
            strSQL = "Update_InPatient_Room '" & udtInPatSched.RoomNumber & "'," & lngAdmit_ID
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
        Else
            strSQL = "Insert_InPatient_Room " & Val(udtInPatSched.IntId) & ", " & udtInPatSched.Facility_ID & ",'" & udtInPatSched.RoomNumber & "'"
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
        End If
'    If cboApptType.Text = "Initial Round" Then
'        If udtInPatSched.Status = "Postpone" Then
'            MsgBox "Successfully updated to the Second Round!", vbInformation, "PatientTrac InPatient Roster"
'        Else
'            MsgBox "Successfully updated!", vbInformation, "PatientTrac InPatient Roster"
'        End If
'    ElseIf cboApptType.Text = "Second Round" Then
'        If udtInPatSched.Status = "Postpone" Then
'            MsgBox "Successfully updated to the Initial Round!", vbInformation, "PatientTrac InPatient Roster"
'        Else
'            MsgBox "Successfully updated!", vbInformation, "PatientTrac InPatient Roster"
'        End If
'    End If
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_InPatient_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_InPatient_Appt", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Update_InPatient_Reason_Notes()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngPurge As Double
Dim lngAdmit_ID As Long
Dim lngPatient As Long
Dim strDate As String
Dim lngDateDiff As Long
Dim strReAdmit As String

    On Error GoTo Error_Update_InPatient_Reason_Notes
    udtInPatSched.Appointment_ID = 0
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "SELECT * FROM INPATIENT_APPOINTMENTS  " _
        & "WHERE PATIENT_ID =  " & Val(udtInPatSched.IntId) & " AND " _
        & "FACILITY_ID = " & udtInPatSched.Facility_ID & " ORDER BY INSERT_DATE DESC"
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        If Not .EOF Then
            If Not IsNull((!INSERT_DATE)) Then
                If Trim((!INSERT_DATE)) <> vbNullString Then
                    strDate = Format((!INSERT_DATE), "mm/dd/yyyy")
                End If
               udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
            End If
        End If
    End With
    If udtInPatSched.Appointment_ID > 0 Then
        strSQL = "UPDATE INPATIENT_APPOINTMENTS SET REASON = '" & Trim(frmSchedComment.mstrNotes) & "' WHERE " _
            & "INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
    End If
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_InPatient_Reason_Notes:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_InPatient_Reason_Notes", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Write_New_Patient_Schedule(vlngPatient As Long) As Boolean
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngPurge As Double
Dim lngAdmit_ID As Long
Dim strReAdmit As String
Dim lngCnt As Long
Dim lngWrite As Long
Dim lngPatient As Long
Dim strAppt As String

    On Error GoTo Error_Write_New_Patient_Schedule
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    'udtInPatSched.Appt_Date = Format(Now, "mm/dd/yyyy")
    If IsProperCase(udtPatientName.CurrentStatus) <> "Active" Then
        MsgBox "Unable to schedule patient until status is changed to Active!", vbInformation, "PatientTrac InPatient Appointments"
    End If
    Get_Facilities udtInPatSched.Facility_ID
    udtInPatSched.ApptComment = vbNullString
    lngRow = grdAppt.ActiveRow
    strReAdmit = vbNullString
    Clear_InSched
    grdAppt.Col = 1
    grdAppt.Row = lngRow
    lngPatient = grdAppt.GetRowItemData(lngRow)
    grdAppt.GetText 1, lngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        udtPatientName.Name = Trim(varCellValue)
    End If
    grdAppt.GetText 4, lngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        udtInPatSched.RoomNumber = Trim(varCellValue)
    End If
    grdAppt.GetText 5, lngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        udtInPatSched.PatientType = Trim(varCellValue)
    End If
    grdAppt.GetText 2, lngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        udtPatientName.CurrentStatus = Trim(varCellValue)
    End If
    udtSchedPatient.IntId = lngPatient
    Get_Insurance Val(udtInPatSched.IntId)
    If TextFound(Trim(clsInsurance.Prim_Insurer), Trim(udtInPatSched.ApptComment), Len(Trim(clsInsurance.Prim_Insurer)), True) Then
         udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, Trim(clsInsurance.Prim_Insurer), "")
    ElseIf TextFound("No record for primary insurance", udtInPatSched.ApptComment, Len("No record for primary insurance"), True) Then
         udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "No record for primary insurance", "")
    End If
    udtInPatSched.ApptComment = Trim(udtInPatSched.ApptComment)
    udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "'", "''")
    lngCnt = lngCnt + 1
     If udtInPatSched.PatientType = "Follow-Up" Then
         udtInPatSched.PatientType = "Subsequent Care"
     ElseIf udtInPatSched.PatientType = "Initial" Then
          udtInPatSched.PatientType = "Initial InPatient"
     ElseIf udtInPatSched.PatientType = "Re-Admission" Then
          udtInPatSched.PatientType = "Initial InPatient"
          strReAdmit = "Re-Admission"
     ElseIf udtInPatSched.PatientType = vbNullString Then
          udtInPatSched.PatientType = "Subsequent Care"
     End If
     If clsFacility.pos = "31" Then
         If udtInPatSched.PatientType = "Subsequent Care" Then
              udtInPatSched.VisitType = "Subsequent care, nursing facility E&M, limited"
         ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
              udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
         End If
     ElseIf clsFacility.pos = "13" Then
         If udtInPatSched.PatientType = "Subsequent Care" Then
              udtInPatSched.VisitType = "Follow-Up visit, asst living facility, limited"
         ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
              udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
         End If
     End If
     lngWrite = lngWrite + 1
     If clsEDI_Client.Site_ID = "1999" Then
         udtInPatSched.Provider_ID = 1000
     End If
     udtInPatSched.Status = "Active"
    udtInPatSched.Appointment_ID = 0
    udtInPatSched.IntId = lngPatient
     If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
     strSQL = "SELECT * FROM INPATIENT_APPOINTMENTS  " _
         & "WHERE PATIENT_ID =  " & Val(udtInPatSched.IntId) & " AND " _
         & "FACILITY_ID = " & udtInPatSched.Facility_ID & " AND APPT_DATE BETWEEN '" & Format(mstrRangeDate, "mm/dd/yyyy") & "' AND '" & Format(mstrUpperRange, "mm/dd/yyyy") & "' ORDER BY INSERT_DATE DESC"
         Set rsReadAppt = cnReadAppt.Execute(strSQL)
         With rsReadAppt
             If Not .EOF Then
                 If Not IsNull((!INSERT_DATE)) Then
                     If Trim((!INSERT_DATE)) <> vbNullString Then
                         strAppt = Format((!INSERT_DATE), "mm/dd/yyyy")
                     End If
                    udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
                 End If
             End If
         End With
         If strAppt = vbNullString Then
             strAppt = Format(mstrUpperRange, "mm/dd/yyyy")
         Else
             strAppt = Format(strAppt, "mm/dd/yyyy")
         End If
     If udtInPatSched.Appointment_ID > 0 Then
         lngPurge = 0
         strSQL = "Update_InPatient_Appt '" & strAppt & "','" & udtInPatSched.RoomNumber & "','" _
             & udtInPatSched.VisitType & "','" & udtInPatSched.PatientType & "','" _
             & IsProperCase(udtPatientName.CurrentStatus) & "', " & lngPurge & ", '" & Fix_Comments(udtInPatSched.ApptComment) & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," _
             & udtInPatSched.Appointment_ID
         Set rsReadAppt = cnReadAppt.Execute(strSQL)
         If strReAdmit <> vbNullString Then
             strSQL = "UPDATE INPATIENT_APPOINTMENTS SET CURRENT_LOCATION = 'Re-Admission' WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
             Set rsReadAppt = cnReadAppt.Execute(strSQL)
         End If
     Else
         strSQL = "Insert_InPatient_Appt " & Val(udtInPatSched.IntId) & ", " _
             & udtInPatSched.Provider_ID & "," & udtInPatSched.Facility_ID & ",'" & udtInPatSched.RoomNumber & "','" _
             & strAppt & "', '" & udtInPatSched.VisitType & "', '" & udtInPatSched.PatientType & "','" & IsProperCase(udtInPatSched.Status) & "','" _
             & Fix_Comments(udtInPatSched.ApptComment) & "'," & mlngLogonUser
         Set rsReadAppt = cnReadAppt.Execute(strSQL)
         With rsReadAppt
             If Not .EOF Then
                 udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
             End If
         End With
         If strReAdmit <> vbNullString And udtInPatSched.Appointment_ID > 0 Then
             strSQL = "UPDATE INPATIENT_APPOINTMENTS SET CURRENT_LOCATION = 'Re-Admission' WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
             Set rsReadAppt = cnReadAppt.Execute(strSQL)
         End If
     End If
     lngAdmit_ID = 0
     strSQL = "Get_InPatient_Room_Appt " & Val(udtInPatSched.IntId) & "," & udtInPatSched.Facility_ID & "," & udtInPatSched.Appointment_ID
         Set rsReadAppt = cnReadAppt.Execute(strSQL)
         With rsReadAppt
             If Not .EOF Then
                 lngAdmit_ID = (!PATIENT_ADMIT_ROOMID)
             End If
         End With
         If lngAdmit_ID > 0 Then
             strSQL = "Update_InPatient_Room '" & udtInPatSched.RoomNumber & "'," & lngAdmit_ID
             Set rsReadAppt = cnReadAppt.Execute(strSQL)
         Else
             strSQL = "Insert_InPatient_Room " & Val(udtInPatSched.IntId) & ", " & udtInPatSched.Facility_ID & "," & udtInPatSched.Appointment_ID & ",'" & udtInPatSched.RoomNumber & "'"
             Set rsReadAppt = cnReadAppt.Execute(strSQL)
         End If
    MsgBox udtPatientName.Name & " " & "has been successfully added to current visit roster.", vbInformation, "PatientTrac InPatient Scheduling"
    Write_New_Patient_Schedule = True
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_New_Patient_Schedule:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_New_Patient_Schedule", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Write_Refresh()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngPurge As Double
Dim lngAdmit_ID As Long
Dim strReAdmit As String

    On Error GoTo Error_Write_Refresh
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    udtInPatSched.Appt_Date = Format(Now, "mm/dd/yyyy")
    Get_Facilities udtInPatSched.Facility_ID
    For lngRow = 1 To grdPatient.MaxRows
        strReAdmit = vbNullString
        Clear_InSched
        grdPatient.Col = 1
        grdPatient.Row = lngRow
        udtInPatSched.IntId = grdPatient.GetRowItemData(lngRow)
        grdPatient.GetText 2, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtInPatSched.RoomNumber = Trim(varCellValue)
        End If
        grdPatient.GetText 3, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtInPatSched.PatientType = Trim(varCellValue)
        End If
        grdPatient.GetText 4, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtInPatSched.Status = Trim(varCellValue)
        End If
        grdPatient.GetText 5, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            udtInPatSched.ApptComment = Trim(varCellValue)
        End If
        If udtInPatSched.PatientType <> "Follow-Up" And udtInPatSched.PatientType <> "Subsequent Care" And udtInPatSched.PatientType <> vbNullString Then
            Get_Insurance Val(udtInPatSched.IntId)
            If TextFound(Trim(clsInsurance.Prim_Insurer), Trim(udtInPatSched.ApptComment), Len(Trim(clsInsurance.Prim_Insurer)), True) Then
                 udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, Trim(clsInsurance.Prim_Insurer), "")
            ElseIf TextFound("No record for primary insurance", udtInPatSched.ApptComment, Len("No record for primary insurance"), True) Then
                 udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "No record for primary insurance", "")
            End If
            udtInPatSched.ApptComment = Trim(udtInPatSched.ApptComment)
            udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "'", "''")
            
            If IsProperCase(udtInPatSched.Status) = "Active" Then
                If udtInPatSched.PatientType = "Follow-Up" Then
                    udtInPatSched.PatientType = "Subsequent Care"
                ElseIf udtInPatSched.PatientType = "Initial" Then
                     udtInPatSched.PatientType = "Initial InPatient"
                ElseIf udtInPatSched.PatientType = "Re-Admission" Then
                     udtInPatSched.PatientType = "Initial InPatient"
                     strReAdmit = "Re-Admission"
                ElseIf udtInPatSched.PatientType = vbNullString Then
                     udtInPatSched.PatientType = "Subsequent Care"
                End If
                If clsFacility.pos = "31" Then
                    If udtInPatSched.PatientType = "Subsequent Care" Then
                         udtInPatSched.VisitType = "Subsequent care, nursing facility E&M, limited"
                    ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                         udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
                    End If
                ElseIf clsFacility.pos = "13" Then
                    If udtInPatSched.PatientType = "Subsequent Care" Then
                         udtInPatSched.VisitType = "Follow-Up visit, asst living facility, limited"
                    ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                         udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
                    End If
                End If
               udtInPatSched.Appointment_ID = 0
                If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
                strSQL = "Get_InPatient_Appt_ID " & Val(udtInPatSched.IntId) & ",'" & Format(udtInPatSched.Appt_Date, "mm/dd/yyyy") & "'," & udtInPatSched.Provider_ID
                    Set rsReadAppt = cnReadAppt.Execute(strSQL)
                    With rsReadAppt
                        If Not .EOF Then
                            udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
                        End If
                    End With
                If clsEDI_Client.Site_ID = "1999" Then
                    udtInPatSched.Provider_ID = 1000
                End If
                If udtInPatSched.Appointment_ID > 0 Then
                    lngPurge = 0
                    strSQL = "Update_InPatient_Appt '" & udtInPatSched.Appt_Date & "','" & udtInPatSched.RoomNumber & "','" _
                        & udtInPatSched.VisitType & "','" & udtInPatSched.PatientType & "','" _
                        & IsProperCase(udtInPatSched.Status) & "', " & lngPurge & ", '" & Fix_Comments(udtInPatSched.ApptComment) & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," _
                        & udtInPatSched.Appointment_ID
                    Set rsReadAppt = cnReadAppt.Execute(strSQL)
                    If strReAdmit <> vbNullString Then
                        strSQL = "UPDATE INPATIENT_APPOINTMENTS SET CURRENT_LOCATION = 'Re-Admission' WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
                        Set rsReadAppt = cnReadAppt.Execute(strSQL)
                    End If
                Else
                    strSQL = "Insert_InPatient_Appt " & Val(udtInPatSched.IntId) & ", " _
                        & udtInPatSched.Provider_ID & "," & udtInPatSched.Facility_ID & ",'" & udtInPatSched.RoomNumber & "','" _
                        & udtInPatSched.Appt_Date & "', '" & udtInPatSched.VisitType & "', '" & udtInPatSched.PatientType & "','" & IsProperCase(udtInPatSched.Status) & "','" _
                        & Fix_Comments(udtInPatSched.ApptComment) & "'," & mlngLogonUser
                    Set rsReadAppt = cnReadAppt.Execute(strSQL)
                    With rsReadAppt
                        If Not .EOF Then
                            udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
                        End If
                    End With
                    If strReAdmit <> vbNullString And udtInPatSched.Appointment_ID > 0 Then
                        strSQL = "UPDATE INPATIENT_APPOINTMENTS SET CURRENT_LOCATION = 'Re-Admission' WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
                        Set rsReadAppt = cnReadAppt.Execute(strSQL)
                    End If
                End If
                lngAdmit_ID = 0
                strSQL = "Get_InPatient_Room_Appt " & Val(udtInPatSched.IntId) & "," & udtInPatSched.Facility_ID & "," & udtInPatSched.Appointment_ID
                    Set rsReadAppt = cnReadAppt.Execute(strSQL)
                    With rsReadAppt
                        If Not .EOF Then
                            lngAdmit_ID = (!PATIENT_ADMIT_ROOMID)
                        End If
                    End With
                    If lngAdmit_ID > 0 Then
                        strSQL = "Update_InPatient_Room '" & udtInPatSched.RoomNumber & "'," & lngAdmit_ID
                        Set rsReadAppt = cnReadAppt.Execute(strSQL)
                    Else
                        strSQL = "Insert_InPatient_Room " & Val(udtInPatSched.IntId) & ", " & udtInPatSched.Facility_ID & "," & udtInPatSched.Appointment_ID & ",'" & udtInPatSched.RoomNumber & "'"
                        Set rsReadAppt = cnReadAppt.Execute(strSQL)
                    End If
            End If
        End If
    Next lngRow
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Refresh:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Refresh", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cboApptType_Click()
    If mbooByPass Then mbooByPass = False: Exit Sub
    If cboApptType.Text = "Initial Round" Then
        Get_InPatientFormat
        ReadData
    ElseIf cboApptType.Text = "Second Round" Then
        Get_InPatientFormat
        Read_Secondary
'    ElseIf cboApptType.Text = "All Patients" Then
'        Get_InPatientFormat
'        Load_All_Patients
    End If
    Exit Sub
End Sub


Private Sub cboExtender_Click()
Dim i%
    For i = 0 To cboExtender.listcount - 1
        If cboExtender.ItemData(i) = mlngLogonUser Then
            cboExtender.Locked = True
            cboApptType.Visible = False
            lblApptType.Visible = False
            Exit For
        End If
    Next i
    Exit Sub
End Sub

Private Function Format_FacilityGrid()
Dim lngRow As Long
    With grdFacility
        .FontSize = 11
        .ClearRange 1, SpreadHeader, 1, .MaxRows, False
        .ColHeaderRows = 2
        .ColWidth(1) = 47.8
        .MaxRows = 0
        lngRow = SpreadHeader
        Format_CellTypeEdit grdFacility, lngRow, 1, "1", "FACILITY NAME", "TypeHAlignCenter", "", HeaderColor, 12
        .SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdFacility.sheet = 1
        grdFacility.ActiveSheet = 1
        Add_Spacer SpreadHeader + 1, grdFacility
    End With
End Function
Private Sub cboProviders_Click()
Dim lngProvID As Long
Dim selProvider As String
Dim i%

    If mbooByPass Then mbooByPass = False:  Exit Sub
    If Trim(cboProviders.Text) <> vbNullString Then
        For i = 0 To cboProviders.listcount - 1
            mbooByPass = True
            If cboProviders.List(i) = cboProviders.Text Then
                mbooByPass = True
                udtInPatSched.Provider_ID = cboProviders.ItemData(i)
                udtInPatSched.Provider = cboProviders.List(i)
                Exit For
            End If
        Next i
    End If
    piProviderID = udtInPatSched.Provider_ID
    Get_InPatientFormat
    ReadData
    If Trim(cboProviders.Text) = vbNullString Then
    mbooByPass = True
    For i = 0 To cboProviders.listcount - 1
        mbooByPass = True
        If cboProviders.List(i) = udtInPatSched.Provider Then
            mbooByPass = True
            cboProviders.Text = cboProviders.List(i)
            Exit For
        End If
    Next i
    End If
    mbooByPass = False
    Exit Sub
End Sub

Private Sub cctAddPatient_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(cctAddPatient.Text) > 2 Then
        Search_Add_Patient
    End If
    Exit Sub
End Sub


Private Sub cctFacility_Click()
    fraFacilities.ZOrder 0
    fraFacilities.Visible = True
End Sub

Private Sub cctFacility_DblClick()
    fraFacilities.Visible = True
End Sub


Private Sub cctFacility_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lngRow As Long
Dim varCellValue
Dim lngLen As Long

    With grdFacility
        For lngRow = 1 To grdFacility.MaxRows
            .GetText 1, lngRow, varCellValue
            lngLen = Len(cctFacility.Text)
            If Left(IsProperCase(varCellValue), lngLen) = Left(IsProperCase(cctFacility.Text), lngLen) Then
                .TopRow = lngRow
                Exit For
            End If
        Next lngRow
    End With
End Sub


Private Sub cctSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strType As String
    If mbooByPass Then mbooByPass = False: Exit Sub
    cboSearchFilter = cboSearchFilter.Text
    If Len(cctSearch.Text) >= 2 Then
        Search_Patient
    End If
    Exit Sub
End Sub
Private Function Format_Patient()
Dim lngCol As Long
Dim lngRow As Long
Dim i%

    Screen.MousePointer = vbHourglass
    grdPatient.ProcessTab = True
    grdPatient.FontBold = True
    grdPatient.ForeColor = 0
    With grdPatient
        .Redraw = False
        .MaxRows = 0
        .MaxCols = 5
        .ColHeaderRows = 4
        .ClearRange 1, 1, grdPatient.MaxCols, grdPatient.MaxRows, False
        .ColWidth(1) = 30
        .ColWidth(2) = 10
        .ColWidth(3) = 15
        .ColWidth(4) = 15
        .ColWidth(5) = 46
        Format_CellTypeEdit grdPatient, SpreadHeader, 1, "5", "PATIENT REPORTS FOR" & vbNewLine & UCase(cctFacility.Text), "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdPatient.SetCellBorder 1, SpreadHeader, 65, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdPatient
        lngCol = 1
        Format_CellTypeStatic grdPatient, SpreadHeader + 2, lngCol, "1", "PATIENT NAME", "TypeHAlignCenter", "PATIENT_ID", "", 0, HeaderColor
        grdPatient.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeStatic grdPatient, SpreadHeader + 2, lngCol, "1", "ROOM", "TypeHAlignCenter", "ROOM_NUMBER", "", 0, HeaderColor
        grdPatient.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeStatic grdPatient, SpreadHeader + 2, lngCol, "1", "VISIT TYPE", "TypeHAlignCenter", "PURPOSE_VISIT", "", 0, HeaderColor
        grdPatient.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeStatic grdPatient, SpreadHeader + 2, lngCol, "1", "PATIENT STATUS", "TypeHAlignCenter", "STATUS", "", 0, HeaderColor
        grdPatient.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeStatic grdPatient, SpreadHeader + 2, lngCol, "1", "COMMENTS", "TypeHAlignCenter", "NOTES", "", 0, HeaderColor
        grdPatient.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdPatient
        .Redraw = True
    End With
    mbooByPass = True
    txtDate.Value = Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")
    Screen.MousePointer = vbDefault
    Exit Function
End Function
Private Sub cmdActive_Click()
    If grdPatient.MaxRows > 0 Then
        Write_InPatient
    End If
    lstListBox.Clear
    Get_InPatientFormat
    ReadData
    fraPatient.Visible = False
    Exit Sub
End Sub

Private Sub cmdAddRoster_Click()
Dim lngPatient As Long
Dim varCellValue

    With grdAppt
        .Col = 1
        .Row = .ActiveRow
        .GetText 5, .ActiveRow, varCellValue
        udtInPatSched.PatientType = varCellValue
        If IsProperCase(udtInPatSched.PatientType) = vbNullString Then
            MsgBox "Please document patient visit type first!", vbInformation, "PatientTrac InPatient Appointments"
            Exit Sub
        End If
        .GetText 3, .ActiveRow, varCellValue
        clsSurgFac.Name = Trim(varCellValue)
        If IsProperCase(clsSurgFac.Name) <> IsProperCase(cctFacility.Text) Then
            MsgBox "Unable to schedule patient from different facilities, please use that facility roster to add patients!", vbInformation, "PatientTrac InPatient Appointments"
            Exit Sub
        End If
        If .BackColor = grdCurItem Then
            lngPatient = .GetRowItemData(.ActiveRow)
            If Write_New_Patient_Schedule(lngPatient) Then
                cctAddPatient.Text = vbNullString
                ReadData
                fraAddPatient.Visible = False
            End If
        Else
            MsgBox "Invalid Patient Selection", vbInformation, "PatientTrac InPatient Appointments"
        End If
    End With
    Exit Sub
End Sub
Private Sub cmdBlockTimes_Click()
    schTimer.Enabled = False
    frmBlockTime.Show vbModal
    Get_Mode
    'schTimer.Enabled = True
    Exit Sub
End Sub

Private Sub cmdCloseAdd_Click()
    Format_Add_Patient_Grid
    fraAddPatient.Visible = False
    Exit Sub
End Sub

Private Sub cmdCloseDeath_Click()
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim strDate As String
Dim varCellValue
Dim lngRow As Long
Dim lngPatient As Long
Dim lngAppt As Long


    On Error GoTo Error_cmdCloseDeath
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    strDate = Format(ccdDeath.Value, "mm/dd/yyyy")
  
    If Trim(strDate) <> vbNullString Then
        With grdPatient
            .Col = 1
            .Row = .ActiveRow
            If .CellNote <> vbNullString Then
                lngPatient = .CellNote
            End If
        End With
    End If
    If lngPatient > 0 Then
        strSQL = "Update_Patient_Date_Death '" & strDate & "'," & lngPatient
        Set rsAppt = cnAppt.Execute(strSQL)
        strSQL = "Update_InPatient_Appt_Death '" & strDate & "'," & lngPatient
        Set rsAppt = cnAppt.Execute(strSQL)
    End If
    grdPatient.DeleteRows grdPatient.ActiveRow, 1
    grdPatient.MaxRows = grdPatient.MaxRows - 1
    MsgBox "Patient Appointments have been cancelland removed from active patients!", vbInformation, "PatientTrac InPatient Roster"
    fraDeath.Visible = False
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdCloseDeath:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "cmdCloseDeath", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdCloseList_Click()
    lstListBox.Clear
    Format_Patient
    Get_InPatientFormat
    ReadData
    fraPatient.Visible = False
    Exit Sub
End Sub

Private Sub cmdCreateVisit2_Click()
    If grdPatient.MaxRows > 0 Then
        Write_InPatient
    End If
    lstListBox.Clear
    Get_InPatientFormat
    ReadData
    fraPatient.Visible = False
    Exit Sub
End Sub

Private Sub cmdCurrentWeek_Click()
    mstrFirstDay = vbNullString
    mBeenHere = True
    grdData.ClearRange 0, SpreadHeader, grdData.MaxCols, grdData.MaxRows, False
    Remove_InsertCells
'    Format_Grid
'    Set_Headers
    If mlngClickPrev > 0 Or mlngClickNext > 0 Then
         udtInPatSched.Appt_Date = Format(Now, "mm/dd/yyyy")
    End If
    Get_Mode
    mlngClickPrev = 0
    mlngClickNext = 0
    Exit Sub
End Sub

Private Sub cmdExcel_Click()
Dim strName As String
Dim ret As Boolean
Dim strTitle As String
Dim Arr
Dim Idx%
Dim strTime As String

    On Error Resume Next
    strTime = Format(Now, "mm/dd/yyyy")
    strTitle = "Patient_Roster"
    If grdPatient.MaxRows > 0 Then
        strTime = Replace(strTime, "/", "_")
        strName = "C:\" & strTitle & "_" & strTime
        ret = grdPatient.ExportExcelBook(strName, "")
    End If
    Exit Sub
End Sub

Private Sub cmdLivAccept_Click()
Dim strDate As String
Dim lngDateDiff As Long
Dim varCellValue
Dim Arr
Dim lngRow As Long
Dim strList As String
Dim booFound As Boolean
Dim lngCol As Long

  strDate = Format(ccdLive.Value, "mm/dd/yyyy")
  fraLive.Visible = False
  If Trim(strDate) <> vbNullString Then
        For lngRow = 1 To grdPatient.MaxRows
            grdPatient.Col = 1
            grdPatient.Row = lngRow
            If grdPatient.BackColor = grdCurItem Then
                booFound = True
                Exit For
            End If
        Next lngRow
        If Not booFound Then
            lngRow = grdPatient.MaxRows
        End If
        udtInPatSched.Schedule_Date = Format(strDate, "mm/dd/yyyy")
        lngCol = 2
        Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", udtInPatSched.Schedule_Date, "TypeHAlignCenter", "APPT_DATE"
        grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
  End If
  Exit Sub
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo Error_InActive_Patients
    Screen.MousePointer = vbHourglass
    mlngPatCnt = 0
    If cmdLoad.Caption = "Show Inactive" Then
        cmdLoad.Caption = "Show Active"
        cmdActive.Visible = False
        cmdCreateVisit2.Visible = False
        Load_InActive
    ElseIf cmdLoad.Caption = "Show Active" Then
        cmdLoad.Caption = "Show Inactive"
        Load_Active_Patient
        cmdCreateVisit2.Visible = True
        cmdActive.Visible = True
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_InActive_Patients:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "InActive_Patients", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub cmdLookUpPatient_Click()
Dim i%

    mbooByPass = True
    lstListBox.Clear
    mbooByPass = True
    cctSearch.Text = vbNullString
    
    cboSearchFilter.Clear
    cboSearchFilter.AddItem "LAST NAME"
    cboSearchFilter.AddItem "EMR ID"
    cboSearchFilter.AddItem "HOSPITAL RECORD"
    cboSearchFilter.AddItem "SOCIAL SECURITY"
    
    For i = 0 To cboSearchFilter.listcount - 1
        If cboSearchFilter.List(i) = "LAST NAME" Then
            mbooByPass = True
            cboSearchFilter.Text = cboSearchFilter.List(i)
            Exit For
        End If
    Next i
    If Validate_Extender_Provider Or Validate_MD_Provider Then
        Format_Patient
        Load_Active_Patient
        fraPatient.Visible = True
        fraPatient.ZOrder 0
    End If
    mbooByPass = False
    Exit Sub
End Sub
Private Sub cmdNewPat_Click()
    On Error GoTo Error_New_Patient

    mbooByPass = True
    lstListBox.Clear
    mbooByPass = True
    cctSearch.Text = vbNullString
    mbooByPass = False
    Screen.MousePointer = vbHourglass
    modSpecific.piPatientID = 0
    modSpecific.piEncounterID = 0
    piPatientID = 0
    piEncounterID = 0
    modSpecific.piProviderID = udtInPatSched.Provider_ID
    Dump_PatientStrData
    PatientInfo.mNewPatient = True
    PatientInfo.Show vbModal
    Screen.MousePointer = vbDefault
    Exit Sub
Error_New_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "New_Patient", "PatientTracEMR.FrmPatient"
    Exit Sub
    Resume
End Sub

Private Sub cmdNewRepeat_Click()
Dim Idx%
Dim lngCol As Long
Dim lngRow As Long
Dim strName As String
Dim lngPatID As Long
Dim strDate As String
Dim strList As String
Dim booExit As Boolean

    strDate = Format(txtDate.Value, "mm/dd/yyyy")
    For Idx = 0 To lstListBox.listcount - 1
        If lstListBox.Selected(Idx) Then
            lngPatID = lstListBox.ItemData(Idx)
            Exit For
        End If
    Next Idx
    With grdPatient
        If lngPatID > 0 Then
            udtInPatSched.IntId = lngPatID
            For lngRow = 1 To grdPatient.MaxRows
                .Col = 1
                .Row = lngRow
                If .CellNote <> vbNullString Then
                    If .CellNote = lngPatID Then
                        booExit = True
                    End If
                End If
            Next lngRow
            If booExit Then
                MsgBox "Patient is already in the Current roster!", vbInformation, "PatientTrac InPatient Roster"
                lstListBox.Clear
                cctSearch.Text = vbNullString
                Exit Sub
            Else
                lngRow = grdPatient.MaxRows + 1
                grdPatient.MaxRows = grdPatient.MaxRows + 1
                udtInPatSched.Provider = cboProviders.Text
                udtInPatSched.Facility = cctFacility.Text
                Get_NewInPatient lngPatID
                Get_Room_Number lngPatID
                udtInPatSched.Schedule_Date = Format(txtDate.Value, "mm/dd/yyyy")
                lngCol = 1
                Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", udtInPatSched.LastName & ", " & udtInPatSched.FirstName, "TypeHAlignLeft", "PATIENT_ID", "TypeVAlignLeft", 0, CStr(lngPatID)
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeStatic grdPatient, lngRow, lngCol, "1", udtInPatSched.Schedule_Date, "TypeHAlignCenter", "APPT_DATE"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 3
                Format_CellTypeEdit grdPatient, lngRow, lngCol, "1", udtInPatSched.RoomNumber, "TypeHAlignCenter", "ROOM_NUMBER"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 4
                Get_Edi_Static
                If clsEDI_Client.Site_ID = "1999" Then
                    strList = "Psychiatric Diagnostic Exam" & Chr(9) & "Subsequent E&M Care, Brief"
                    Format_CellTypeCombo grdPatient, lngRow, lngCol, "1", strList, "TypeHAlignLeft", "PURPOSE_VISIT"
                    Set_GridCurSel grdPatient, "Subsequent Visit", lngCol, lngRow, "1", "PURPOSE_VISIT"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Else
                    Format_CellTypeCombo grdPatient, lngRow, lngCol, "1", Load_Visit, "TypeHAlignLeft", "PURPOSE_VISIT"
                End If
                If clsEDI_Client.Site_ID = "1125" Then
                    Set_GridCurSel grdPatient, "Subsequent Visit", lngCol, lngRow, "1", "PURPOSE_VISIT"
                    grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 5
                Format_CellTypeCombo grdPatient, lngRow, lngCol, "1", Get_Grid_Status, "TypeHAlignCenter", "STATUS"
                If udtInPatSched.Status = vbNullString Then
                    udtInPatSched.Status = "Active"
                End If
                Set_GridCurSel grdPatient, udtInPatSched.Status, lngCol, lngRow, "1", "STATUS"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 6
                Format_CellTypeEdit grdPatient, lngRow, lngCol, "1", "", "TypeHAlignCenter", "NOTES"
                grdPatient.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        End If
    End With
    lstListBox.Clear
    cctSearch.Text = vbNullString
    Exit Sub
End Sub

Private Sub cmdNextWeek_Click()
Dim Arr
Dim strDate As String
Dim nDays As Long
Dim lngCol As Long
Dim lngRow As Long
Dim vDay
Dim lngAddDay As Long
Dim vFirstDay As Variant
Dim lngAddWeek As Long
Dim strOption As String
Dim lngWeekDays As Long

    On Error GoTo Error_NextWeek
    Screen.MousePointer = vbHourglass
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScheduleWeek")
    If Val(strOption) > 0 Then
        lngWeekDays = Val(strOption)
    Else
        lngWeekDays = 7
    End If
    grdData.GetText 1, SpreadHeader + 2, vFirstDay
    If vFirstDay <> vbNullString Then
        Arr = Split(vFirstDay, "-")
        If UBound(Arr) > 0 Then
            vFirstDay = Arr(1)
        End If
    End If
    schTimer.Enabled = False
    lngAddWeek = 7
    mChange = True
    If vFirstDay = vbNullString Then
        strDate = udtInPatSched.Appt_Date
    Else
        strDate = vFirstDay
    End If
    nDays = 1 * lngAddWeek
    strDate = DateAdd("d", nDays, strDate)
    vDay = strDate
    With grdData
        .Redraw = False
        .ClearRange 0, SpreadHeader, .MaxCols, .MaxRows, True
        Remove_InsertCells
        Get_InPatientFormat
        'Set_Headers
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .MaxCols = lngWeekDays
        lngRow = SpreadHeader + 2
        lngCol = 1
        .ColWidth(lngCol) = 20
        clsGridFormat.Col = lngCol
        clsGridFormat.Text = "Monday" & " - " & Format(vDay, "mm/dd/yyyy")
        Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        nDays = 1
        lngCol = 2
        clsGridFormat.Col = lngCol
        vDay = DateAdd("d", nDays, vDay)
        .ColWidth(lngCol) = 20
        clsGridFormat.Text = "Tuesday" & " - " & Format(vDay, "mm/dd/yyyy")
        Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        clsGridFormat.Col = lngCol
        vDay = DateAdd("d", nDays, vDay)
        .ColWidth(lngCol) = 20
        clsGridFormat.Text = "Wednesday" & " - " & Format(vDay, "mm/dd/yyyy")
        Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        clsGridFormat.Col = lngCol
         vDay = DateAdd("d", nDays, vDay)
        .ColWidth(lngCol) = 20
        clsGridFormat.Text = "Thurdsday" & " - " & Format(vDay, "mm/dd/yyyy")
        Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        clsGridFormat.Col = lngCol
         vDay = DateAdd("d", nDays, vDay)
        .ColWidth(lngCol) = 20
        clsGridFormat.Text = "Friday" & " - " & Format(vDay, "mm/dd/yyyy")
        Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        If .MaxCols >= 6 Then
            vDay = DateAdd("d", nDays, vDay)
            lngCol = 6
            clsGridFormat.Col = lngCol
            .ColWidth(lngCol) = 20
            clsGridFormat.Text = "Saturday" & " - " & Format(vDay, "mm/dd/yyyy")
            Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
            grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        If .MaxCols >= 6 Then
            vDay = DateAdd("d", nDays, vDay)
             lngCol = 7
            clsGridFormat.Col = lngCol
            .ColWidth(lngCol) = 20
            clsGridFormat.Text = "Sunday" & " - " & Format(vDay, "mm/dd/yyyy")
            Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
            grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        .Redraw = True
    End With
    Read_Week_InPatient_Appointments
    rtfChanges.Text = vbNullString
    mChange = False
    'schTimer.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_NextWeek:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "NextWeek", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdPatientStatus_DropDownClick()
    PopupMenu mnu_Patient_Type, , cmdPatientStatus.Left + 40, cmdPatientStatus.Top + cmdPatientStatus.Height
    Exit Sub
End Sub


Private Sub cmdPhone_Click()
    mstrPatContact = "Phone"
    frmTelephone.Show vbModal
    Format_Messages
    Read_Messages
End Sub

Private Sub cmdPrevWeek_Click()
Dim Arr
Dim strDate As String
Dim Idx As Integer
Dim strSQL As String
Dim sDate As String
Dim vDay As String
Dim sMonth As String
Dim nDays As Long
Dim lngAddDay As Long
Dim vFirstDay As Variant
Dim lngAddWeek As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strOption As String
Dim lngWeekDays As Long

    On Error GoTo Error_PrevWeek
    Screen.MousePointer = vbHourglass
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScheduleWeek")
    If Val(strOption) > 0 Then
        lngWeekDays = Val(strOption)
    Else
        lngWeekDays = 7
    End If
    grdData.GetText 1, SpreadHeader + 2, vFirstDay
    If vFirstDay <> vbNullString Then
        Arr = Split(vFirstDay, "-")
        If UBound(Arr) > 0 Then
            vFirstDay = Arr(1)
        End If
    End If
    If IsDate(vFirstDay) Then
        schTimer.Enabled = False
        mlngClickNext = 0
        mlngClickPrev = mlngClickPrev + 1
        lngAddWeek = 7
        mChange = True
        If vFirstDay = vbNullString Then
            strDate = udtInPatSched.Appt_Date
        Else
            strDate = vFirstDay
        End If
        nDays = -1 * lngAddWeek
        strDate = DateAdd("d", nDays, strDate)
        vDay = strDate
    Else
        MsgBox "Please select week view first!", vbInformation, "PatientTrac Appointment Schedule"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'Format Columns
    With grdData
        .Redraw = False
        .ClearRange 0, SpreadHeader, .MaxCols, .MaxRows, False
        Remove_InsertCells
        Get_InPatientFormat
       ' Set_Headers
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .MaxCols = lngWeekDays
        lngRow = SpreadHeader + 2
        lngCol = 1
        .ColWidth(lngCol) = 20
        clsGridFormat.Col = lngCol
        clsGridFormat.Text = "Monday" & " - " & Format(vDay, "mm/dd/yyyy")
        Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        nDays = 1
        lngCol = 2
        vDay = DateAdd("d", nDays, vDay)
        .ColWidth(lngCol) = 20
        clsGridFormat.Col = lngCol
        clsGridFormat.Text = "Tuesday" & " - " & Format(vDay, "mm/dd/yyyy")
        Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        clsGridFormat.Col = lngCol
        vDay = DateAdd("d", nDays, vDay)
        .ColWidth(lngCol) = 20
        clsGridFormat.Text = "Wednesday" & " - " & Format(vDay, "mm/dd/yyyy")
        Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        clsGridFormat.Col = lngCol
         vDay = DateAdd("d", nDays, vDay)
        .ColWidth(lngCol) = 20
        clsGridFormat.Text = "Thurdsday" & " - " & Format(vDay, "mm/dd/yyyy")
        Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        clsGridFormat.Col = lngCol
         vDay = DateAdd("d", nDays, vDay)
        .ColWidth(lngCol) = 20
        clsGridFormat.Text = "Friday" & " - " & Format(vDay, "mm/dd/yyyy")
        Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        If .MaxCols >= 6 Then
            vDay = DateAdd("d", nDays, vDay)
            lngCol = 6
            clsGridFormat.Col = lngCol
            .ColWidth(lngCol) = 20
            clsGridFormat.Text = "Saturday" & " - " & Format(vDay, "mm/dd/yyyy")
            Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
            grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        If .MaxCols >= 6 Then
            vDay = DateAdd("d", nDays, vDay)
             lngCol = 7
            clsGridFormat.Col = lngCol
            .ColWidth(lngCol) = 20
            clsGridFormat.Text = "Sunday" & " - " & Format(vDay, "mm/dd/yyyy")
            Format_CellTypeFixed grdData, lngRow, clsGridFormat.Col, "1", clsGridFormat.Text, "TypeHAlignCenter", "", HeaderColor
            grdData.SetCellBorder clsGridFormat.Col, lngRow, clsGridFormat.Col, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        .Redraw = True
    End With
    Read_Week_InPatient_Appointments
    rtfChanges.Text = vbNullString
    mChange = False
    'schTimer.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_PrevWeek:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "PrevWeek", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdPrintAppt_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    mbooCancelPrint = False
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
      If mbooCancelPrint Then
            MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Appointment Scheduling"
            mbooCancelPrint = False
            Exit Sub
      End If
      With grdData
          .PrintSmartPrint = False
          .PrintAbortMsg = "Printing - Click Cancel to quit"
          .PrintJobName = "Scheudule"
          .PrintHeader = strHeaders
          .PrintFooter = "/cPage #/p/n"
          .PrintBorder = True
          .PrintColor = True
          .PrintGrid = True
          .PrintOrientation = PrintOrientationPortrait
          .PrintMarginTop = 200
          .PrintMarginBottom = 50
          .PrintMarginLeft = 100
          .PrintMarginRight = 100
          .hDCPrinter = Printer.hdc
          .PrintType = 0
          .PrintRowHeaders = True
          .PrintShadows = False
          If Printer.Copies > 1 Then
              For i = 1 To Printer.Copies
                  .PrintSheet
              Next i
          Else
              .PrintSheet
          End If
          mbooPrintPreview = False
      End With
      Exit Sub
Error_cmdPrint_Click:
    MsgBox Err.Description, vbInformation
    Exit Sub
End Sub

Private Sub cmdPrintRoster_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    mbooCancelPrint = False
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
      If mbooCancelPrint Then
            MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Appointment Scheduling"
            mbooCancelPrint = False
            Exit Sub
      End If
      With grdPatient
          .PrintSmartPrint = False
          .PrintAbortMsg = "Printing - Click Cancel to quit"
          .PrintJobName = "Scheudule"
          .PrintHeader = strHeaders
'          .PrintBestFitPagesWide
          .PrintFooter = "/cPage #/p/n"
          .PrintBorder = True
          .PrintColor = True
          .PrintGrid = True
          .PrintOrientation = PrintOrientationPortrait
          .PrintMarginTop = 200
          .PrintMarginBottom = 50
          .PrintMarginLeft = 100
          .PrintMarginRight = 100
          .hDCPrinter = Printer.hdc
          .PrintType = 0
          .PrintRowHeaders = True
          .PrintShadows = False
          If Printer.Copies > 1 Then
              For i = 1 To Printer.Copies
                  .PrintSheet
              Next i
          Else
              .PrintSheet
          End If
          mbooPrintPreview = False
      End With
      Exit Sub
Error_cmdPrint_Click:
    MsgBox Err.Description, vbInformation
    Exit Sub
End Sub

Private Sub cmdRefresh_Click()
Dim lngRow As Long
Dim lngCol As Long

    Format_Patient
    Load_Active_Patient
    Exit Sub
End Sub
Private Function Update_Refresh(vlngPatient As Long)
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngPurge As Double
Dim lngAdmit_ID As Long
Dim lngPatient As Long
Dim strDate As String
Dim lngDateDiff As Long
Dim strReAdmit As String

    On Error GoTo Error_Update_InPatient_Appt
    lngPatient = vlngPatient
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    lngRow = grdData.ActiveRow
    Clear_InSched
    grdData.Col = 1
    grdData.Row = lngRow
    udtInPatSched.IntId = grdData.GetRowItemData(lngRow)
'    If grdData.CellNote <> vbNullString Then
'        udtInPatSched.IntId = Val(grdData.CellNote)
'    End If
    grdData.GetText 1, lngRow, varCellValue
    udtInPatSched.RoomNumber = varCellValue
    grdData.GetText 5, lngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        Select Case CStr(varCellValue)
            Case "Follow-Up"
                udtInPatSched.PatientType = "Subsequent Care"
            Case "Initial"
                 udtInPatSched.PatientType = "Initial InPatient"
            Case "Re-Admission"
                 udtInPatSched.PatientType = "Initial InPatient"
                 strReAdmit = "Re-Admission"
        End Select
    End If
    grdData.GetText 4, lngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        udtInPatSched.Status = Trim(varCellValue)
    End If
    grdData.GetText 6, lngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then
        udtInPatSched.ApptComment = Trim(varCellValue)
    End If
    If clsFacility.pos = "31" Then
        If udtInPatSched.PatientType = "Subsequent Care" Then
             udtInPatSched.VisitType = "Subsequent care, nursing facility E&M, limited"
        ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
             udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
        End If
    ElseIf clsFacility.pos = "13" Then
        If udtInPatSched.PatientType = "Subsequent Care" Then
             udtInPatSched.VisitType = "Follow-Up visit, asst living facility, limited"
        ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
             udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
        End If
    End If
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "SELECT * FROM INPATIENT_APPOINTMENTS  " _
        & "WHERE PATIENT_ID =  " & Val(udtInPatSched.IntId) & " AND " _
        & "FACILITY_ID = " & udtInPatSched.Facility_ID & " ORDER BY INSERT_DATE DESC"
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        If Not .EOF Then
            If Not IsNull((!INSERT_DATE)) Then
                If Trim((!INSERT_DATE)) <> vbNullString Then
                    strDate = Format((!INSERT_DATE), "mm/dd/yyyy")
                End If
               udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
'                If Not IsNull((!Patient_Type)) Then udtInPatSched.PatientType = (!Patient_Type)
            End If
        End If
    End With
    lngDateDiff = DateDiff("d", Format(strDate, "mm/dd/yyyy"), Format(Now, "mm/dd/yyyy"))
    If lngDateDiff > 0 And lngDateDiff < 10 Then
        udtInPatSched.Appt_Date = Format(strDate, "mm/dd/yyyy")
    Else
        udtInPatSched.Appt_Date = Format(Now, "mm/dd/yyyy")
    End If
    If udtInPatSched.VisitType = vbNullString Then udtInPatSched.VisitType = "Subsequent Visit"
'    strSQL = "Get_InPatient_Appt_ID " & Val(lngPatient) & ",'" & Format(udtInPatSched.Appt_Date, "mm/dd/yyyy") & "'," & udtInPatSched.Provider_ID
'        Set rsReadAppt = cnReadAppt.Execute(strSQL)
'        With rsReadAppt
'            If Not .EOF Then
'                udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
'            End If
'        End With
    If udtInPatSched.ApptComment = "Review Staff Notes" Then udtInPatSched.ApptComment = vbNullString
    If clsFacility.pos = "31" Then
        If udtInPatSched.PatientType = "Subsequent Care" Then
             udtInPatSched.VisitType = "Subsequent care, nursing facility E&M, limited"
        ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
             udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
        End If
    ElseIf clsFacility.pos = "13" Then
        If udtInPatSched.PatientType = "Subsequent Care" Then
             udtInPatSched.VisitType = "Follow-Up visit, asst living facility, limited"
        ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
             udtInPatSched.VisitType = "Initial E&M Care, Brief Visit"
        End If
    End If
    If udtInPatSched.Appointment_ID > 0 Then
        If IsProperCase(udtInPatSched.Status) = "Cancelled" Or IsProperCase(udtInPatSched.Status) = "Rescheduled" Then
            lngPurge = 1
        Else
            lngPurge = 0
        End If
        strSQL = "Update_InPatient_Appt '" & udtInPatSched.Appt_Date & "','" & udtInPatSched.RoomNumber & "','" & udtInPatSched.VisitType & "','" & udtInPatSched.PatientType & "','" _
            & IsProperCase(udtInPatSched.Status) & "', " & lngPurge & ", '" & Fix_Comments(udtInPatSched.ApptComment) & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," _
            & udtInPatSched.Appointment_ID
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        If strReAdmit <> vbNullString Then
            strSQL = "UPDATE INPATIENT_APPOINTMENTS SET CURRENT_LOCATION = 'Re-Admission' WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
        End If
    Else
        strSQL = "Insert_InPatient_Appt " & Val(udtInPatSched.IntId) & ", " _
            & udtInPatSched.Provider_ID & "," & udtInPatSched.Facility_ID & ",'" & udtInPatSched.RoomNumber & "','" _
            & udtInPatSched.Appt_Date & "', '" & udtInPatSched.VisitType & "','" & udtInPatSched.PatientType & "','" & IsProperCase(udtInPatSched.Status) & "','" _
            & Fix_Comments(udtInPatSched.ApptComment) & "'," & mlngLogonUser
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
            With rsReadAppt
                If Not .EOF Then
                    udtInPatSched.Appointment_ID = (!INPATIENT_APPT_ID)
                End If
            End With
            If strReAdmit <> vbNullString Then
                strSQL = "UPDATE INPATIENT_APPOINTMENTS SET CURRENT_LOCATION = 'Re-Admission' WHERE INPATIENT_APPT_ID = " & udtInPatSched.Appointment_ID
                Set rsReadAppt = cnReadAppt.Execute(strSQL)
            End If
    End If
    strSQL = "Get_InPatient_Room " & Val(udtInPatSched.IntId) & "," & udtInPatSched.Facility_ID
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        With rsReadAppt
            If Not .EOF Then
                lngAdmit_ID = (!PATIENT_ADMIT_ROOMID)
            End If
        End With
        If lngAdmit_ID > 0 Then
            strSQL = "Update_InPatient_Room '" & udtInPatSched.RoomNumber & "'," & lngAdmit_ID
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
        Else
            strSQL = "Insert_InPatient_Room " & Val(udtInPatSched.IntId) & ", " & udtInPatSched.Facility_ID & ",'" & udtInPatSched.RoomNumber & "'"
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
        End If
'    If cboApptType.Text = "Initial Round" Then
'        If udtInPatSched.Status = "Postpone" Then
'            MsgBox "Successfully updated to the Second Round!", vbInformation, "PatientTrac InPatient Roster"
'        Else
'            MsgBox "Successfully updated!", vbInformation, "PatientTrac InPatient Roster"
'        End If
'    ElseIf cboApptType.Text = "Second Round" Then
'        If udtInPatSched.Status = "Postpone" Then
'            MsgBox "Successfully updated to the Initial Round!", vbInformation, "PatientTrac InPatient Roster"
'        Else
'            MsgBox "Successfully updated!", vbInformation, "PatientTrac InPatient Roster"
'        End If
'    End If
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_InPatient_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_InPatient_Appt", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cmdRemoveRow_Click()
Dim strDate As String
Dim lngDateDiff As Long
Dim varCellValue
Dim Arr
Dim lngRow As Long
Dim strList As String
Dim booFound As Boolean
Dim lngCol As Long
Dim lngPatID As Long
 
      For lngRow = 1 To grdPatient.MaxRows
        grdPatient.Col = 1
        grdPatient.Row = lngRow
        If grdPatient.BackColor = grdCurItem Then
            If grdPatient.CellTag <> vbNullString Then
                Remove_Data lngPatID
            End If
            grdPatient.DeleteRows lngRow, 1
            grdPatient.MaxRows = grdPatient.MaxRows - 1
            booFound = True
            Exit For
        End If
    Next lngRow
    If Not booFound Then
        lngRow = grdPatient.MaxRows
        grdPatient.DeleteRows lngRow, 1
        grdPatient.MaxRows = grdPatient.MaxRows - 1
    End If
    Exit Sub
End Sub

Private Sub cmdRepair_Click()
Dim i%
Dim strSQL As String
Dim cnApptCheck As New ADODB.Connection
Dim rsApptCheck As New ADODB.Recordset
Dim strApptTime As String
Dim strApptDate As String
Dim lngProviderID As String
Dim lngDuplicates As Long
Dim lngSelProv As Long

    ' For Code Mode Only Repairs
    cnApptCheck.open psConnect(0)
    strSQL = "select patient_id from INPATIENT_APPOINTMENTS where Active = 1 order by patient_id"
    Set rsApptCheck = cnApptCheck.Execute(strSQL)
    With rsApptCheck
        Do Until .EOF
           If Not IsNull(!Patient_ID) And (!Patient_ID) <> 0 Then
                lngSelProv = (!Patient_ID)
                strSQL = "Select Appt_Time, Appt_Date, Provider_Id from INPATIENT_APPOINTMENTS where " _
                    & "patient_id = " & lngSelProv & "  Order by appt_Date"
                cnApptCheck.open psConnect(0)
                Set rsApptCheck = cnApptCheck.Execute(strSQL)
                With rsApptCheck
                    Do Until rsApptCheck.EOF
                        If (!Appt_Time) <> vbNullString And Trim(!Appt_Time) = strApptTime Then
                            If (!Appt_Date) <> vbNullString And Trim(!Appt_Date) = strApptDate Then
                                If (!Provider_ID) <> 0 And (!Provider_ID) = lngProviderID Then
                                    ' We have a duplicate patient appointment
                                    lngDuplicates = lngDuplicates + 1
                                    MsgBox "Duplicates" & lngDuplicates, vbInformation
                                End If
                            End If
                        Else
                            If (!Appt_Time) <> vbNullString Then strApptTime = Trim(!Appt_Time)
                            If (!Appt_Date) <> vbNullString Then strApptDate = Trim(!Appt_Date)
                            If (!Provider_ID) <> 0 Then lngProviderID = (!Provider_ID)
                        End If
                    rsApptCheck.MoveNext
                    Loop
                End With
           End If
        .MoveNext
        Loop
    End With
    
    Close_DB psConnect(0)
    
End Sub

Private Sub cmdReport_Click()
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim cnMedProv As New ADODB.Connection
Dim rsMedProv As New ADODB.Recordset


    On Error GoTo Error_Update_Roster
    Screen.MousePointer = vbHourglass
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    strSQL = "Update_InPatient_Roster '" & Format(Now, "mm/dd/yyyy") & "', " _
        & udtInPatSched.Facility_ID & "," & udtInPatSched.Provider_ID
    If cnMedProv.State = adStateOpen Then
        cnMedProv.Close
        Set cnMedProv = Nothing
        Set rsMedProv = Nothing
        Set Flds = Nothing
    End If
    Get_InPatientFormat
    ReadData
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Update_Roster:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_Roster", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdResched_Click()
Dim i%
    cctAddPatient.Text = vbNullString
    Format_Add_Patient_Grid
    fraAddPatient.Visible = True
    fraAddPatient.ZOrder 0
    Exit Sub
End Sub

Private Sub cmdRosterReport_Click()

    mbooByPass = True
    lstListBox.Clear
    mbooByPass = True
    cctSearch.Text = vbNullString
    mbooByPass = False
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "InPatientRoster"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub cmdSave_Click()
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngPatient As Long

    With grdData
        If .ActiveRow > 0 Then
            lngRow = .ActiveRow
            .Col = 4
            .Row = lngRow
            lngPatient = .GetRowItemData(lngRow)
            .GetText 4, .ActiveRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                Update_InPatient_Appt lngPatient
                MsgBox "Patient visit has been successfully updated!", vbInformation, "PatientTrac InPatient Roster"
                If IsProperCase(Trim(varCellValue)) = "Workup Completed" Or IsProperCase(Trim(varCellValue)) = "Cancelled" Then
                    .DeleteRows lngRow, 1
                    .MaxRows = .MaxRows - 1
                End If
            End If
        End If
    End With
    Exit Sub
End Sub

Private Sub cmdSelectFac_Click()
Dim lngRow As Long
Dim lngCol As Long
Dim bData As Boolean

    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        If .BackColor = grdCurItem Then
            bData = True
        End If
    End With
    If Not bData Then
        MsgBox "Please select a facility first!", vbInformation, "PatientTrac InPatient Roster"
    Else
        Change_Facility
        fraFacilities.Visible = False
    End If
    Exit Sub
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
    frmNurseTemplates.Show vbModal
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdStatus_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdStatus_Click", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdStop_Click()

    If cmdStop.Caption = "Stop Refresh" Then
        schTimer.Enabled = False
        cmdStop.Caption = "Refresh Appt"
    Else
        'schTimer.Enabled = True
        cmdStop.Caption = "Stop Refresh"
    End If
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
    
    If grdPatient.MaxRows > 0 Then
        Update_Patient_Status
    End If
    Write_Refresh
    lstListBox.Clear
    MsgBox "Select Refresh to see changes!", vbInformation, "Patient Roster"
    Exit Sub
End Sub
Private Sub cmdViewProvider_DropDownClick()

    PopupMenu View, , cmdViewProvider.Left, cmdViewProvider.Top + cmdViewProvider.Height
 Exit Sub
End Sub


Private Sub Form_Activate()
    mFirst = False
    mstrRangeDate = vbNullString
    mbooByPass = False
    Exit Sub
End Sub

Private Sub Form_Load()
        mstrShow = vbNullString
        Me.Top = 0
        Me.Left = frmMDI.Left + 120
        Me.Width = 18700
        Me.Height = 11700
        cboApptType.Clear
        mlngPatCnt = 0
        mbooLoad = True
        Prepare_Grid
        mbooLoad = False
        mFirst = True
    Exit Sub
End Sub

Private Sub Form_Paint()
    
    grdPatient.ZOrder 0
    grdPatient.Redraw = True
    grdData.ZOrder 0
    grdData.Redraw = True
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    mFirstLoad = False
    If IsFormLoaded("frmInPatient") Then
        frmInPatient.ZOrder 0
    End If
    Unload Me
  Exit Sub
End Sub

Private Sub grdAppt_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngPatient As Long
Dim lngRow As Long
Dim varCellValue

    If Row < 1 Then Exit Sub
    If TextFound("Supp", PsUserID, Len("Supp"), True) Then
        mbooSupport = True
    End If
    With grdAppt
         lngPatient = .GetRowItemData(Row)
        If lngPatient > 0 And Validate_Provider_Permission(lngPatient) Then
             .BlockMode = True
                 .Col = 1
                 .Col2 = .MaxCols
                 .Row = Row
                 .Row2 = Row
                 .BackColor = grdCurItem
             .BlockMode = False
         ElseIf lngPatient > 0 And Not mbooSupport Then
             MsgBox "This patient's records are currently in the facility which has not been assigned to you. Please contact PatientTrac support to update facility status for this patient!", vbInformation, "PatientTrac Patient Roster"
             Exit Sub
         ElseIf lngPatient > 0 And mbooSupport Then
             .BlockMode = True
                 .Col = 1
                 .Col2 = .MaxCols
                 .Row = Row
                 .Row2 = Row
                 .BackColor = grdCurItem
             .BlockMode = False
         End If
        .GetText 2, Row, varCellValue
        udtPatientName.CurrentStatus = varCellValue
        .GetText 3, Row, varCellValue
        clsSurgFac.Name = varCellValue
        If clsSurgFac.Name = "No Facility" Then clsSurgFac.Name = vbNullString
        If IsProperCase(udtPatientName.CurrentStatus) <> "Active" Or clsSurgFac.Name = vbNullString Then
            lngPatient = .GetRowItemData(Row)
            MsgBox "Please update patient status first!", vbInformation, "PatientTrac InPatient Appointments"
            Dump_PatientStrData
            udtSchedPatient.IntId = lngPatient
            PatientInfo.mNewPatient = False
            PatientInfo.Show vbModal
            grdAppt.ZOrder 0
            grdAppt.Redraw = True
            Get_PatientName lngPatient
            Format_CellTypeStatic grdAppt, Row, 2, "1", udtPatientName.CurrentStatus, "TypeHAlignCenter", "PATIENT STATUS"
            grdAppt.SetCellBorder 2, Row, 2, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Get_Room_Number lngPatient
            Format_CellTypeStatic grdAppt, Row, 4, "1", udtInPatSched.RoomNumber, "TypeHAlignCenter", "ROOM"
            grdAppt.SetCellBorder 4, Row, 4, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Read_Patient_Status lngPatient
            Get_SurgFac "", clsFacStatus.FacilityID
            Format_CellTypeStatic grdAppt, Row, 3, "1", clsSurgFac.Name, "TypeHAlignCenter", "ROOM"
            grdAppt.SetCellBorder 3, Row, 3, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            If IsProperCase(udtPatientName.CurrentStatus) <> "Active" Then
                MsgBox "Unable to schedule patient until status is changed to Active!", vbInformation, "PatientTrac InPatient Appointments"
                Exit Sub
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
        lngPatient = .GetRowItemData(Row)
       If lngPatient > 0 And Validate_Provider_Permission(lngPatient) Then
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                .BackColor = grdCurItem
            .BlockMode = False
        ElseIf lngPatient > 0 And Not mbooSupport Then
            MsgBox "This patient's records are currently in the facility which has not been assigned to you. Please contact PatientTrac support to update facility status for this patient!", vbInformation, "PatientTrac Patient Roster"
            Exit Sub
        ElseIf lngPatient > 0 And mbooSupport Then
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                .BackColor = grdCurItem
            .BlockMode = False
        End If
    End With
    Exit Sub
End Sub


Private Sub grdData_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
Dim varCellValue
Dim lngPatient As Long

    With grdData
        .Col = 1
        .Row = Row
        If .CellNote <> vbNullString Then
            lngPatient = .CellNote
        End If
        lngPatient = .GetRowItemData(Row)
'        .GetText 5, Row, varCellValue
'        udtInPatSched.PatientType = varCellValue
        .GetText 4, Row, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            If Col = 4 Then
                udtInPatSched.Status = Trim(varCellValue)
                If udtInPatSched.Status = "Initial Round" Or udtInPatSched.Status = "Second Round" Then
                    Update_InPatient_Status
                    If cboApptType.Text = "Initial Round" Then
                        Get_InPatientFormat
                        ReadData
                    ElseIf cboApptType.Text = "Second Round" Then
                        Get_InPatientFormat
                        Read_Secondary
                    End If
                Else
                    Update_InPatient_Appt lngPatient
                End If
            ElseIf Col = 5 Then
                Update_InPatient_Appt lngPatient
            End If
        End If
    End With
End Sub

Private Sub grdData_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim vDate
Dim strDate As String
Dim Arr
Dim arr1
Dim vPatient
Dim lngPatientID As Long
Dim vTime
Dim sFirstName As String
Dim sLastName As String
Dim i%
Dim strSQL As String
Dim cnData As New ADODB.Connection
Dim rsData As New ADODB.Recordset
Dim varCellValue
Dim lngCol As Long
Dim strList As String

    On Error GoTo Error_GrdData_DblClick
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    If Col = 6 Then
        piPatientID = 0
        piPatientID = grdData.GetRowItemData(Row)
        udtInPatSched.IntId = piPatientID
        grdData.GetText Col, Row, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            If piPatientID > 0 Then
                frmSchedComment.mstrNotes = vbNullString
                frmSchedComment.mstrType = "Appointment"
                frmSchedComment.Show vbModal
                If frmSchedComment.mstrNotes <> vbNullString Then
                    Update_InPatient_Reason_Notes
                    Format_CellTypeEdit grdData, Row, Col, "1", Trim(frmSchedComment.mstrNotes), "TypeHAlignLeft", "NOTES"
                End If
                frmSchedComment.mstrNotes = vbNullString
            End If
        End If
        Exit Sub
    End If
    Clear_InSched
    Clear_Appt
'    Clear_Provider
    grdData.GetText 4, Row, varCellValue
    If Trim(varCellValue) = "Cancelled" Then Exit Sub
    grdData.GetText 5, Row, varCellValue
    udtInPatSched.PatientType = varCellValue
'    If Validate_MD_Provider Then
        Screen.MousePointer = vbHourglass
        If IsFormLoaded("frmInPatient") Then
            Unload frmInPatient
        End If
        lngPatientID = grdData.GetRowItemData(Row)
        If lngPatientID = 0 Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            With grdData
                .BlockMode = True
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = Row
                    .Row2 = Row
                    .BackColor = grdCurItem
                .BlockMode = False
            End With
        End If
        mlngPatientID = 0
        piPatientID = 0
        piEncounterID = 0
        modSpecific.piEncounterID = 0
        frmInPatient.piSelEnctrID = 0
        If cnData.State = adStateClosed Then cnData.open psConnect(0)
        strSQL = "Get_Appointment_Data_Facility " & udtInPatSched.Facility_ID
        Set rsData = cnData.Execute(strSQL)
        With rsData
            Do Until .EOF
                If Not IsNull((!INSERT_DATE)) Then
                    If Trim((!INSERT_DATE)) <> vbNullString Then
                        udtInPatSched.Appt_Date = Format((!INSERT_DATE), "mm/dd/yyyy")
                    End If
                End If
                .MoveNext
                If DateDiff("d", Format(Now, "mm/dd/yyyy"), Format(udtInPatSched.Appt_Date, "mm/dd/yyyy")) <= 0 Then
                    Exit Do
                End If
            Loop
        End With
        If udtInPatSched.SchedType = vbNullString Then udtInPatSched.SchedType = "Patient"
        If lngPatientID > 0 Then
            If Check_Patient(lngPatientID, udtInPatSched.Appt_Date) Then
                mbooPatValid = False
                mlngPatientID = lngPatientID
                piPatientID = lngPatientID
                modSpecific.piPatientID = piPatientID
                frmPatients.piPatientID = piPatientID
                If Validate_Extender_Provider Then
                    Get_Extender_Encounter
                ElseIf Validate_MD_Provider Then
                   Get_CurrentEnctr_MD udtInPatSched.Appt_Date
                End If
                If mbooPatValid Then
                    Get_Patient piPatientID
                    modSpecific.piEncounterID = piEncounterID
                End If
            Else
                mlngPatientID = lngPatientID
                piPatientID = lngPatientID
                modSpecific.piPatientID = piPatientID
                frmPatients.piPatientID = piPatientID
                If Validate_Extender_Provider Then
                    Get_Extender_Encounter
                ElseIf Validate_MD_Provider Then
                   Get_CurrentEnctr_MD udtInPatSched.Appt_Date
                End If
                If IsNumeric(mlngPatientID) Then
                    Get_Patient piPatientID
                    modSpecific.piEncounterID = piEncounterID
                    mbooPatValid = True
                Else
                    mbooPatValid = False
                End If
            End If
        Else
            mbooPatValid = False
        End If
        mGridClick = True
        If cnData.State = adStateOpen Then
            cnData.Close
            Set cnData = Nothing
            Set rsData = Nothing
        End If
        If mbooPatValid Then
            frmMDI.StatusBar.Panels(1).Text = "Patient: " & clsPatient.Name
            If mstrPractice = "Psychiatry" Then
                frmMDI.StatusBar.Panels(2).Text = "Hospital ID: " & clsPatient.ExtID
            Else
                frmMDI.StatusBar.Panels(2).Text = "Financial ID: " & clsPatient.ExtID
            End If
            Get_Encounter piEncounterID
            frmMDI.StatusBar.Panels(3).Text = "Patient EMR ID: " & piPatientID
            frmMDI.StatusBar.Panels(4).Text = "Encounter Number: " & piEncounterID
            frmMDI.StatusBar.Panels(5).Text = "Visit Date: " & Format(clsEncounter.StartDate, "MMMM DD,YYYY")
            frmMDI.StatusBar.Panels(6).Text = "Provider: " & clsEncounter.Provider
            If piEncounterID > 0 Then
                Load_Enctr_Format
                If Validate_MD_Provider Then
                    udtInPatSched.Status = "Present"
                    Update_InPatient_Status
                End If
'--------------------
                lngCol = 4
                If Validate_Extender_Provider Then
                    strList = "Active" & Chr(9) & "Cancelled" & Chr(9) & "Postpone" & Chr(9) & "Workup Completed"
                    Format_CellTypeCombo grdData, Row, lngCol, "1", strList, "TypeHAlignCenter", "STATUS"
                   ' Set_GridCurSel grdData, "Workup Completed", lngCol, Row, "1", "STATUS"
'                ElseIf Validate_MD_Provider Then
'                    grdData.DeleteRows Row, 1
'                    grdData.MaxRows = grdData.MaxRows - 1
                End If
            End If
        End If
        If cnData.State = adStateOpen Then
            cnData.Close
            Set cnData = Nothing
            Set rsData = Nothing
        End If
        Screen.MousePointer = vbDefault
    Exit Sub
Error_GrdData_DblClick:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "GrdData_DblClick", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub grdData_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
Dim vDate
Dim strDate As String
Dim Arr
Dim strTime As String
Dim varCellValue
Dim booValidAppt As Boolean
Dim lngPatID As Long

    On Error GoTo Error_New_Schedule
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    udtInPatSched.IntId = vbNullString
    With grdData
        lngPatID = .GetRowItemData(Row)
        If Col = 2 Then
            If lngPatID > 0 Then
                Clear_InSched
                grdData.Redraw = False
                If cboProviders.ListIndex >= 0 Then udtInPatSched.Provider_ID = cboProviders.ItemData(cboProviders.ListIndex)
                udtInPatSched.Appt_Date = Format(MonthView1.Value, "mm/dd/yyyy")
                With grdData
                    udtInPatSched.IntId = lngPatID
                    mLoadPatient = True
                    mGridClick = True
                    Load_SchedPatient Col, Row
                End With
            End If
        Else
            With grdData
                .Col = Col
                .Row = Row
                If .CellTag <> vbNullString Then
                    Select Case UCase(.CellTag)
                        Case "CURRENT_MEDS"
                            udtInPatSched.IntId = lngPatID
                            frmMeds.mstrSource = "InPatient Appointments"
                            frmMeds.piCurPatient = lngPatID
                            frmMeds.Show vbModal
                            ReadData
                    End Select
                End If
            End With
        End If
    End With
    'schTimer.Enabled = True
    grdData.Redraw = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_New_Schedule:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "New_Schedule", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Read_Messages()
Dim i%
Dim strSQL As String
Dim cnMessage As New ADODB.Connection
Dim rsMessage As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strDate As String
Dim lngCol As Long
Dim lngRow As Long
Dim booAddRow As Boolean
Dim strSymptoms As String
Dim lngMesID As Long
Dim strUrgent As String

    On Error GoTo Error_Read_Messages
    Screen.MousePointer = vbHourglass
    
    grdMessages.ClearRange 1, 0, grdMessages.MaxCols, grdMessages.MaxRows, True
    Format_Messages
    If grdMessages.MaxRows <= 1 Then
        Format_Messages
    End If
    If mlngLogonUser <= 0 Then PsUserID = Get_PsUserID(psConnect(0))
    If cnMessage.State = adStateClosed Then cnMessage.open psConnect(0)
    strSQL = "Read_Message_Provider " & mlngLogonUser
    Set rsMessage = cnMessage.Execute(strSQL)
    With rsMessage
        Do Until rsMessage.EOF
            booAddRow = False
            Clear_Phone
            lngMesID = 0
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "CALLER_ID"
                                udtTelephone.CallerID = Fld.Value
                            Case "MESSAGE_TIME"
                                udtTelephone.Time = Fld.Value
                            Case "SENDER_ID"
                                udtTelephone.SenderID = Fld.Value
                            Case "RECIEVOR_ID"
                                udtTelephone.ReceiverID = Fld.Value
                            Case "MEDICATION_STATUS"
                                udtTelephone.MedStatus = Fld.Value
                            Case "PATIENT_SYMPTOMS"
                                udtTelephone.Symptoms = Fld.Value
                            Case "MESSAGE_NOTE"
                                udtTelephone.MessNote = Fld.Value
                            Case "CONTACT_TYPE"
                                udtTelephone.ContactType = Fld.Value
                            Case "MESSAGE_ID"
                                lngMesID = Fld.Value
                            Case "CALLER_NAME"
                                udtTelephone.CallerName = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtTelephone.CallerID = 0 Then
                booAddRow = False
            End If
            If Not booAddRow Then
                If DateDiff("d", "01/01/1900", Format(udtTelephone.Time, "mm/dd/yyyy")) > 0 Then
                    If udtTelephone.CallerID > 0 And udtTelephone.SenderID = mlngLogonUser Or _
                        udtTelephone.ReceiverID = mlngLogonUser Then
                        booAddRow = True
                    Else
                        booAddRow = False
                    End If
                Else
                    booAddRow = False
                End If
            End If
            If booAddRow Then
                mlngMessRow = mlngMessRow + 1
                grdMessages.Row = mlngMessRow
                grdMessages.MaxRows = mlngMessRow
                Format_CellTypeStatic grdMessages, mlngMessRow, 1, "1", Format(udtTelephone.Time, "mm/dd/yyyy"), "TypeHAlignLeft", "MESSAGE_TIME"
                grdMessages.SetRowItemData mlngMessRow, lngMesID
                If udtTelephone.CallerName <> vbNullString Then
                    Format_CellTypeStatic grdMessages, mlngMessRow, 2, "1", udtTelephone.CallerName, "TypeHAlignLeft", "CALLER_NAME"
                End If
                If udtTelephone.MedStatus <> vbNullString Then
                    Format_CellTypeStatic grdMessages, mlngMessRow, 3, "1", udtTelephone.MedStatus, "TypeHAlignLeft", "MEDICATION_STATUS"
                End If
                If udtTelephone.Symptoms <> vbNullString Then
                   If TextFound("Urgent", udtTelephone.Symptoms, Len("Urgent"), True) Then
                        strUrgent = "Urgent"
                        udtTelephone.Symptoms = Replace(udtTelephone.Symptoms, "Urgent", "")
                        Format_CellTypeStatic grdMessages, mlngMessRow, 4, "1", udtTelephone.Symptoms, "TypeHAlignLeft", "PATIENT_SYMPTOMS"
                   Else
                        Format_CellTypeStatic grdMessages, mlngMessRow, 4, "1", udtTelephone.Symptoms, "TypeHAlignLeft", "PATIENT_SYMPTOMS"
                   End If
                End If
                If Trim(udtTelephone.MessNote) <> vbNullString Then
                    grdMessages.SetText 5, mlngMessRow, udtTelephone.MessNote
                    Format_CellTypeStatic grdMessages, mlngMessRow, 5, "1", udtTelephone.MessNote, "TypeHAlignLeft", "MESSAGE_NOTE"
                End If
                If Trim(udtTelephone.CallType) <> vbNullString And Trim(udtTelephone.CallType) = "Inquiry" Then
                    grdMessages.BlockMode = True
                        grdMessages.Col = 1
                        grdMessages.Col2 = grdMessages.MaxCols
                        grdMessages.Row = mlngMessRow
                        grdMessages.Row2 = mlngMessRow
                        grdMessages.BackColor = &HC0FFC0
                    grdMessages.BlockMode = False
                    grdMessages.Col = 1
                    grdMessages.Row = mlngMessRow
                    grdMessages.CellTag = "Inquiry"
                End If
                If strUrgent = "Urgent" Then
                    grdMessages.BlockMode = True
                        grdMessages.Col = 1
                        grdMessages.Col2 = grdMessages.MaxCols
                        grdMessages.Row = mlngMessRow
                        grdMessages.Row2 = mlngMessRow
                        grdMessages.BackColor = YellowColor
                    grdMessages.BlockMode = False
                End If
                grdMessages.SetCellBorder 1, mlngMessRow, 1, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdMessages.SetCellBorder 2, mlngMessRow, 2, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdMessages.SetCellBorder 3, mlngMessRow, 3, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdMessages.SetCellBorder 4, mlngMessRow, 4, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdMessages.SetCellBorder 5, mlngMessRow, 5, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
            If booAddRow Then
                grdMessages.SetCellBorder 1, mlngMessRow, grdMessages.MaxCols, mlngMessRow, 16, GridBlue, CellBorderStyleSolid
            End If
        Loop
    End With
    If cnMessage.State = adStateOpen Then
        cnMessage.Close
        Set cnMessage = Nothing
        Set rsMessage = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Read_Messages:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Messages", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub grdFacility_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim varCellValue
Dim lngCol As Long
    If Row <= 0 Then Exit Sub
    '------------------------
    cctFacility.Text = vbNullString
    With grdFacility
        For lngRow = 1 To grdFacility.MaxRows
            .Col = 1
            .Row = lngRow
            .BackColor = NormalColor
        Next lngRow
        lblApptType.Visible = False
        cboApptType.Visible = False
        If Not Validate_Extender_Provider Or Validate_MD_Provider Then
            MsgBox "You are not authorized to review this facility, please reselect valid facility", vbInformation, "PatientTrac InPatient Roster"
            Write_Error Err.Description, Err.Number, "Attempting to Access Unauthorized Facility", "Invalid Facility Access", App.EXEName & "." & Me.Name
            Exit Sub
        End If
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = grdCurItem
        .BlockMode = False
        .Col = Col
        .Row = Row
        udtInPatSched.Facility = .GetRowItemData(Row)
        piFacilityID = udtInPatSched.Facility
        .GetText Col, Row, varCellValue
        cctFacility.Text = varCellValue
    End With
    Exit Sub
End Sub


Private Sub grdMessages_Click(ByVal Col As Long, ByVal Row As Long)
Dim iRow%
    mlngMessRow = Row
    mlngMessCol = Col
     With grdMessages
        .Redraw = False
        If mlngMessRow >= 1 Then
            For iRow = 1 To .MaxRows
                .BlockMode = True
                    .Col = 0
                    .Col2 = .MaxCols
                    .Row = iRow
                    .Row2 = iRow
                    .FontBold = True
                    .BackColor = &HFFFFFF
                .BlockMode = False
            Next iRow
        End If
        .Redraw = True
    End With
    Exit Sub
End Sub


Private Sub grdMessages_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim iRow%
Dim booNewContact As Boolean
   mlngMessRow = Row
   mlngMessCol = Col
   booNewContact = False
   If mlngMessRow < 1 Then Exit Sub
   mlngMessageID = 0
   With grdMessages
        .Redraw = False
        mlngMessageID = 0
        mlngMessageID = .GetRowItemData(mlngMessRow)
            If mlngGridRow >= 1 Then
                For iRow = 1 To .MaxRows
                    .BlockMode = True
                        .Col = 0
                        .Col2 = .MaxCols
                        .Row = iRow
                        .Row2 = iRow
                        .FontBold = True
                        .BackColor = &HFFFFFF
                    .BlockMode = False
                Next iRow
            End If
            For iRow = 1 To .MaxRows
                .BlockMode = True
                    .Col = 0
                    .Col2 = .MaxCols
                    .Row = mlngMessRow
                    .Row2 = mlngMessRow
                    .FontBold = True
                    .BackColor = &HFFFF&
                .BlockMode = False
                Exit For
            Next iRow
        .Col = 1
        .Row = mlngMessRow
        If .CellTag = "Inquiry" Then
            booNewContact = True
            mstrPatContact = "Contact"
        Else
            booNewContact = True
            mstrPatContact = "Phone"
        End If
        If mlngMessageID <> 0 Then
            
            frmTelephone.Show vbModal
            Format_Messages
            Read_Messages
        End If
        .Redraw = True
   End With
    Exit Sub
End Sub


Private Sub grdPatient_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
Dim varCellValue
Dim lngPatient As Long


    mbooByPass = True
    lstListBox.Clear
    mbooByPass = True
    cctSearch.Text = vbNullString
    With grdPatient
        lngPatient = .GetRowItemData(Row)
        .GetText 4, Row, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            If Col = 4 Then
                Select Case IsProperCase(varCellValue)
                    Case "Deceased"
                        fraDeath.Visible = True
                        fraDeath.ZOrder 0
                        MsgBox "Please enter date of death!", vbInformation, "PatientTrac InPatient Scheduling"
                        Exit Sub
                End Select

            End If
        End If
    End With
    mbooByPass = True
    Exit Sub
End Sub

Private Sub grdPatient_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngCol As Long
Dim lngRow As Long
Dim varCellValue

    With grdPatient
        If Row > 0 Then
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
            .Col = Col
            .Row = Row
            If .CellTag <> vbNullString Then
                clsGridFormat.CellTag = .CellTag
            End If
            Select Case clsGridFormat.CellTag
                Case "PATIENT_ID"
                    udtInPatSched.IntId = .GetRowItemData(Row)
                    modSpecific.piPatientID = 0
                    modSpecific.piEncounterID = 0
                    piPatientID = 0
                    piEncounterID = 0
                    modSpecific.piProviderID = udtInPatSched.Provider_ID
                    Dump_PatientStrData
                    udtSchedPatient.IntId = Val(udtInPatSched.IntId)
                    PatientInfo.mNewPatient = False
                    PatientInfo.Show vbModal
                    grdPatient.ZOrder 0
                    grdPatient.Redraw = True
                    If cmdLoad.Caption = "Show Active" Then
                        If udtSchedPatient.IntId > 0 Then
                            Get_Patient Val(udtSchedPatient.IntId)
                            If IsProperCase(clsPatient.Status) = "Active" Then
                                .DeleteRows Row, 1
                                .MaxRows = .MaxRows - 1
                                mlngPatCnt = mlngPatCnt - 1
                                lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
                            End If
                        End If
                    End If
                Case "NOTES"
                    frmSchedComment.Show vbModal
                Case "STATUS"
                    .GetText Col, Row, varCellValue
                    If Trim(varCellValue) <> "Active" Then
                        udtInPatSched.IntId = .GetRowItemData(Row)
                        modSpecific.piPatientID = 0
                        modSpecific.piEncounterID = 0
                        piPatientID = 0
                        piEncounterID = 0
                        modSpecific.piProviderID = udtInPatSched.Provider_ID
                        Dump_PatientStrData
                        udtSchedPatient.IntId = Val(udtInPatSched.IntId)
                        PatientInfo.mNewPatient = False
                        PatientInfo.Show vbModal
    '                    If cmdLoad.Caption = "Show Inactive" Then
    '                        Load_Active_Patient
    '                    ElseIf cmdLoad.Caption = "Show Active" Then
    '                        Load_InActive
    '                    End If
                    End If
            End Select
        End If
    End With
    Exit Sub
End Sub


Private Sub lstListBox_ItemCheck(Item As Integer)
Dim i%
Dim strSQL
Dim lngPatientID As Long
Dim iAnswer As Integer
Dim strDate As String
Dim booSkip As Boolean
Dim lngRow As Long

    On Error GoTo Error_lstListBox_ItemCheck
    If mbooByPass Then mbooByPass = False: Exit Sub
    If lstListBox.listcount < 0 Then
        MsgBox "Please search for Patient first!", vbCritical, "PatientTrac Schedule Patient"
        Exit Sub
    End If
    For i = 0 To lstListBox.listcount - 1
        If lstListBox.ItemData(i) <> lstListBox.ItemData(Item) Then
            mbooByPass = True
            lstListBox.Selected(i) = False
        End If
    Next i
    mbooByPass = False
    Screen.MousePointer = vbHourglass
    cctSearch.Text = vbNullString
    If lstListBox.Selected(Item) Then
        If lstListBox.ItemData(Item) > 0 Then
            lngPatientID = lstListBox.ItemData(Item)
            mbooByPass = True
        End If
    End If
   If lngPatientID > 0 And Validate_Provider_Permission(lngPatientID) Then
        modSpecific.piPatientID = 0
        modSpecific.piEncounterID = 0
        piPatientID = 0
        piEncounterID = 0
        modSpecific.piProviderID = udtInPatSched.Provider_ID
        Dump_PatientStrData
        udtSchedPatient.IntId = lngPatientID
        PatientInfo.mNewPatient = False
        PatientInfo.Show vbModal
'        Format_Patient
'        Load_Active_Patient
    ElseIf lngPatientID > 0 And Not mbooSupport Then
        MsgBox "This patient's records are currently in the facility which has not been assigned to you. Please contact PatientTrac support to update facility status for this patient!", vbInformation, "PatientTrac Patient Roster"
        mbooByPass = True
        lstListBox.Selected(Item) = False
        mbooByPass = False
    ElseIf lngPatientID > 0 And mbooSupport Then
        modSpecific.piPatientID = 0
        modSpecific.piEncounterID = 0
        piPatientID = 0
        piEncounterID = 0
        modSpecific.piProviderID = udtInPatSched.Provider_ID
        Dump_PatientStrData
        udtSchedPatient.IntId = lngPatientID
        PatientInfo.mNewPatient = False
        PatientInfo.Show vbModal
'        Format_Patient
'        Load_Active_Patient
    End If
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_lstListBox_ItemCheck:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "lstListBox_ItemCheck", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub mnu_Active_Click()
    Load_Patient_Status "Active"
End Sub

Private Sub mnu_All_PatientType_Click()
    Get_InPatientFormat
    ReadData
    Exit Sub
End Sub

Private Sub mnu_Cancelled_Click()
    Load_Patient_Status "Cancelled"
End Sub

Private Sub mnu_Day_Click()
'    lblApptDate.Visible = True
    cmdPrevWeek.Visible = False
    cmdCurrentWeek.Visible = False
    cmdNextWeek.Visible = False
    mstrApptMode = "Day"
    Get_Mode
    Exit Sub
End Sub
Private Sub mnu_Month_Click()
    mstrApptMode = "Month"
    Get_Mode
    cmdPrevWeek.Visible = False
    cmdCurrentWeek.Visible = False
    cmdNextWeek.Visible = False
    lblProvider.Visible = False
    Exit Sub
End Sub


Private Sub mnu_Postpone_Click()
    Load_Patient_Status "Postpone"
End Sub

Private Sub mnu_Week_Click()
    mstrApptMode = "Weekly"
    Get_Mode
    cmdPrevWeek.Visible = True
    cmdCurrentWeek.Visible = True
    cmdNextWeek.Visible = True
'    lblApptDate.Visible = False
     rtfChanges.Text = vbNullString
End Sub

Private Function Remove_InsertCells()
Dim lngCol As Long
Dim lngRow As Long
Dim vDate
Dim vdate2
Dim anchorcol  As Variant
Dim anchorrow As Variant
Dim colspan As Variant
Dim rowspan As Variant
Dim IsSpan As Boolean
Dim Span As Integer

    With grdData
        .Redraw = False
        For lngCol = 1 To .MaxCols
            For lngRow = 1 To grdData.MaxRows
               .GetText 0, lngRow, vDate
              If lngRow > 0 Then .GetText 0, lngRow - 1, vdate2
                If vDate <> vbNullString And vDate <> vbNullString Then
                    If vDate = vdate2 Then
                        .DeleteRows (lngRow), 1
                        .MaxRows = .MaxRows - 1
                    End If
                End If
            Next lngRow
        Next lngCol
        .Redraw = True
    End With
    Exit Function
End Function
Private Function LookUpHeader(vstrName As String) As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
    
    lngCol = 0
    With grdData
        LookUpHeader = vbNullString
        For lngRow = 1 To .MaxRows
            .Col = lngCol
            .GetText lngCol, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString And TextFound(vstrName, Trim(varCellValue), Len(vstrName), True) Then
                LookUpHeader = lngCol & "|" & lngRow
                Exit For
            End If
        Next lngRow
    End With
    Exit Function
End Function
Private Function Remove_Span()
Dim Col As Variant
Dim Row As Variant
Dim Col2 As Variant
Dim Row2 As Variant
Dim Span As Integer
Dim anchorcol As Variant
Dim anchorrow As Variant
Dim colspan As Variant
Dim rowspan As Variant
Dim i As Long
Dim j As Long
Dim IsSpan As Boolean

    ' Get the selected cells
    grdData.getSelection 0, Col, Row, Col2, Row2
    If mbooOverBook = False Then Exit Function
    grdData.Redraw = False
    For i = 1 To 7
        For j = SpreadHeader To grdData.MaxRows
            Span = grdData.GetCellSpan(i, j, anchorcol, anchorrow, colspan, rowspan)
            
            grdData.Redraw = False
            If Span Then
                If grdData.Col = 1 Then anchorcol = "1"
                If j = SpreadHeader Then anchorrow = SpreadHeader
                If j = SpreadHeader + 1 Then anchorrow = SpreadHeader + 1
                If j = SpreadHeader + 2 Then anchorrow = SpreadHeader + 2
                grdData.RemoveCellSpan anchorcol, anchorrow
                IsSpan = True
            End If
        Next j
    Next i
    
    grdData.Redraw = True
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
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    cboProviders.Clear
    strSQL = "Load_MD_Providers_Facility " & udtInPatSched.Facility_ID
     Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        Do Until .EOF
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            lngProvID = 0
            booSkip = False
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
                For i = 0 To cboProviders.listcount - 1
                    mbooByPass = True
                    If TextFound(clsProvider.Name, cboProviders.List(i), Len(clsProvider.Name), True) Then
                        booSkip = True
                        Exit For
                    End If
                Next i
                If Not booSkip Then
                    mbooByPass = True
                    cboProviders.AddItem clsProvider.Name
                    mbooByPass = True
                    If IsNumeric(lngProvID) Then cboProviders.ItemData(cboProviders.NewIndex) = lngProvID
                End If
            End If
        Loop
    End With
    mbooByPass = True
    For i = 0 To cboProviders.listcount - 1
        If cboProviders.ItemData(i) = lngProvID Then
            mbooByPass = True
            udtInPatSched.Provider_ID = cboProviders.ItemData(i)
            cboProviders.Text = cboProviders.List(i)
            Exit For
        End If
    Next i
    If cnMedProv.State = adStateOpen Then
        cnMedProv.Close
        Set cnMedProv = Nothing
        Set rsMedProv = Nothing
        Set Flds = Nothing
    End If
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Appt_Providers:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Appt_Providers", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
 
Public Sub InfoFrame(pCol As Long, pRow As Long)
Dim vValue
Dim iAnswer As Integer
Dim lngPatientID As Long
Dim vDate
Dim Arr

    On Error GoTo Error_InfoFrame
    Screen.MousePointer = vbHourglass
'    mChange = False
'    grdData.GetText pCol, pRow, vValue
'    grdData.GetText 0, pRow, vTime
'    If vTime = vbNullString Then
'        grdData.GetText 0, pRow - 1, vTime
'    End If
'    udtAppointment.SchedTime = vTime
'    lblApptDate.AutoSize = True
'    lblApptDate.Caption = "Date:" & " " & Format(udtInPatSched.Appt_Date, "mmmm dd, yyyy")
'    If mstrApptMode <> "Day" Then
'        grdData.GetText pCol, SpreadHeader + 2, vDate
'        Arr = Split(vDate, "-")
'        If UBound(Arr) > 0 Then
'            vDate = Arr(1)
'            udtInPatSched.Appt_Date = vDate
'        Else
'            udtInPatSched.Appt_Date = vDate
'        End If
'    End If
    'schTimer.Enabled = True
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_InfoFrame:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "InfoFrame", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Public Function GetDay(ipDate As String)
    Dim lDay As Long

    lDay = WeekDay(ipDate)

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

Private Function ReadData()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Arr
Dim lngPatID As Long
Dim Idx As Integer
Dim lngRow As Long
Dim varCellValue
Dim vTime
Dim i%
Dim lngCol As Long
Dim lngFoundID As Boolean
Dim lngProvID As Long
Dim strNotes As String
Dim bPatFound As Boolean
Dim lngApptID As Long
Dim lngDup As Long
Dim strReAdmit As String
Dim strRange As String
Dim strAddDate As String

    On Error GoTo Error_ReadData
    Screen.MousePointer = vbHourglass
    mlngPatCnt = 0
    rtfChanges.Text = vbNullString
    grdData.MaxRows = 0
    mbooByPass = True
    For i = 0 To cboProviders.listcount - 1
        If cboProviders.List(i) = cboProviders.Text Then
            mbooByPass = True
            udtInPatSched.Provider_ID = cboProviders.ItemData(i)
            mbooByPass = True
            udtInPatSched.Provider = cboProviders.List(i)
            Exit For
        End If
    Next i
    mbooByPass = True
    With grdFacility
        .Col = 1
        .Row = .ActiveRow
        .GetText 1, .ActiveRow, varCellValue
        udtInPatSched.Facility_ID = .GetRowItemData(.ActiveRow)
        udtInPatSched.Facility = varCellValue
    End With
    
    If mstrApptMode = vbNullString Then mstrApptMode = "Day"
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Get_Appointment_Data_Facility " & udtInPatSched.Facility_ID
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            If Not IsNull((!Appt_Date)) Then
                If Trim((!Appt_Date)) <> vbNullString Then
                    udtInPatSched.Appt_Date = Format((!Appt_Date), "mm/dd/yyyy")
                End If
            End If
            .MoveNext
            If udtInPatSched.Appt_Date <> "01/01/1900" Then
                Exit Do
            End If
            
        Loop
    End With
    If DateDiff("d", udtInPatSched.Appt_Date, MonthView1.Value) < 0 Then
        udtInPatSched.Appt_Date = MonthView1.Value
    End If
    strRange = DateAdd("d", 1, udtInPatSched.Appt_Date)
    Get_Appt_Range
    strAddDate = DateAdd("d", 1, mstrUpperRange)
    If mstrRangeDate = vbNullString Then
        mstrRangeDate = udtInPatSched.Appt_Date
    End If
    If udtInPatSched.Provider_ID = 0 Then udtInPatSched.Provider_ID = 1000
    clsShedGrid.ApptDate = udtInPatSched.Appt_Date
    clsShedGrid.ApptDate = Format(clsShedGrid.ApptDate, "mm/dd/yyyy")
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    If Validate_MD_Provider Then
        strSQL = "SELECT  INPATIENT_APPOINTMENTS.*, PATIENT.TITLE, PATIENT.PREF_CONTACT, " _
            & "Patient.Suffix , Patient.Last_Name, Patient.First_Name, Patient.Middle_Name " _
            & "From INPATIENT_APPOINTMENTS " _
            & "INNER JOIN PATIENT ON PATIENT.PATIENT_ID = INPATIENT_APPOINTMENTS.PATIENT_ID " _
            & "WHERE CONVERT( VARCHAR(10), INPATIENT_APPOINTMENTS.APPT_DATE, 101)  BETWEEN '" & Format(mstrRangeDate, "mm/dd/yyyy") & "' AND  '" & Format(strAddDate, "mm/dd/yyyy") & "'  AND " _
            & "INPATIENT_APPOINTMENTS.FACILITY_ID = " & udtInPatSched.Facility_ID & " AND " _
            & "INPATIENT_APPOINTMENTS.PROVIDER_ID = " & udtInPatSched.Provider_ID & " AND INPATIENT_APPOINTMENTS.ACTIVE = 0  AND  " _
            & "(INPATIENT_APPOINTMENTS.STATUS <> 'Cancelled' AND  INPATIENT_APPOINTMENTS.STATUS <> 'Second Round') " _
            & "AND PATIENT.STATUS = 'Active' " _
            & "ORDER BY INPATIENT_APPOINTMENTS.ROOM_NUMBER ASC, PATIENT.LAST_NAME, PATIENT.FIRST_NAME"
    Else
        strSQL = "SELECT  INPATIENT_APPOINTMENTS.*, PATIENT.TITLE, PATIENT.PREF_CONTACT, " _
            & "Patient.Suffix , Patient.Last_Name, Patient.First_Name, Patient.Middle_Name " _
            & "From INPATIENT_APPOINTMENTS " _
            & "INNER JOIN PATIENT ON PATIENT.PATIENT_ID = INPATIENT_APPOINTMENTS.PATIENT_ID " _
            & "WHERE CONVERT( VARCHAR(10), INPATIENT_APPOINTMENTS.APPT_DATE, 101)  BETWEEN '" & Format(mstrRangeDate, "mm/dd/yyyy") & "' AND  '" & Format(strAddDate, "mm/dd/yyyy") & "'  AND " _
            & "INPATIENT_APPOINTMENTS.FACILITY_ID = " & udtInPatSched.Facility_ID & " AND " _
            & "INPATIENT_APPOINTMENTS.PROVIDER_ID = " & udtInPatSched.Provider_ID & " AND INPATIENT_APPOINTMENTS.ACTIVE = 0  AND  " _
            & "(INPATIENT_APPOINTMENTS.STATUS <> 'Cancelled' AND  INPATIENT_APPOINTMENTS.STATUS <> 'Second Round' AND  INPATIENT_APPOINTMENTS.STATUS <> 'Present' AND  INPATIENT_APPOINTMENTS.STATUS <> 'MD COMPLETED') " _
            & "AND PATIENT.STATUS = 'Active' " _
            & "ORDER BY INPATIENT_APPOINTMENTS.ROOM_NUMBER ASC, PATIENT.LAST_NAME, PATIENT.FIRST_NAME"
    End If
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            Clear_InSched
            Clear_Appt
            lngPatID = 0
            udtInPatSched.Appointment_ID = 0
            lngApptID = 0
            udtInPatSched.SchedType = vbNullString
            strReAdmit = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "SUFFIX"
                                udtInPatSched.Suffix = Trim(Fld.Value)
                            Case "LAST_NAME"
                                udtInPatSched.LastName = IsProperCase(Fld.Value)
                            Case "MIDDLE_NAME"
                                udtInPatSched.MiddleInitial = IsProperCase(Fld.Value)
                            Case "FIRST_NAME"
                                udtInPatSched.FirstName = Fld.Value
                            Case "PREF_CONTACT"
                                udtInPatSched.RoomTelephone = Trim(Fld.Value)
                            Case "PATIENT_ID"
                                udtInPatSched.IntId = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtInPatSched.VisitType = Fld.Value
                            Case "APPT_TYPE"
                                udtInPatSched.SchedType = Trim(Fld.Value)
                            Case "STATUS"
                                udtInPatSched.Status = Trim(Fld.Value)
                            Case "STATUS_NOTES"
                                udtInPatSched.ApptComment = Fld.Value
                                udtInPatSched.ApptComment = Replace(udtInPatSched.ApptComment, "''", "'")
                            Case "PROVIDER_ID"
                                udtInPatSched.Provider_ID = Fld.Value
                            Case "APPT_DATE"
                                udtInPatSched.Appt_Date = Fld.Value
                            Case "FACILITY_ID"
                                udtInPatSched.Facility_ID = Fld.Value
                            Case "RESOURCE_TYPE"
                                udtInPatSched.ResourceType = Fld.Value
                            Case "ROOM_NUMBER"
                                udtInPatSched.RoomNumber = Fld.Value
                            Case "TELEPHONE"
                                udtInPatSched.RoomTelephone = Fld.Value
                            Case "DURATION"
                                udtInPatSched.Duration = Fld.Value
                            Case "PATIENT_TYPE"
                                udtInPatSched.PatientType = Fld.Value
                            Case "CURRENT_LOCATION"
                                strReAdmit = Fld.Value
                            Case "INPATIENT_APPT_ID"
                                lngApptID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If lngDup <> Val(udtInPatSched.IntId) Then
                 If Not Validate_CurrentFacility(Val(udtInPatSched.IntId)) Then
                    udtInPatSched.Appointment_ID = lngApptID
                    mlngPatCnt = mlngPatCnt + 1
                    udtInPatSched.Title = Fix_Title(udtInPatSched.Title)
                    udtInPatSched.Suffix = Fix_Suffix(udtInPatSched.Suffix)
                    udtInPatSched.LastName = Fix_Name(udtInPatSched.LastName)
                    udtInPatSched.FirstName = Fix_Name(udtInPatSched.FirstName)
                    If udtInPatSched.FirstName <> vbNullString Then
                        udtInPatSched.Collection = Trim(udtInPatSched.LastName) & ", " & Trim(udtInPatSched.FirstName)
                    End If
                    If Trim(udtInPatSched.Suffix) <> vbNullString Then
                        udtInPatSched.Name = udtInPatSched.FirstName & " " & udtInPatSched.LastName & ", " & udtInPatSched.Suffix
                    Else
                        udtInPatSched.Name = udtInPatSched.FirstName & " " & udtInPatSched.LastName
                    End If
                    If udtInPatSched.PatientType <> vbNullString Then
                        If strReAdmit = "Re-Admission" Then
                            udtInPatSched.PatientType = "Re-Admission"
                        ElseIf udtInPatSched.PatientType = "Subsequent Care" Then
                            udtInPatSched.PatientType = "Follow-Up"
                        ElseIf udtInPatSched.PatientType = "Initial InPatient" Then
                            udtInPatSched.PatientType = "Initial"
                        End If
                    End If
                    Get_Room_Number Val(udtInPatSched.IntId)
                    udtInPatSched.Name = Replace(udtInPatSched.Name, "''", "'")
                    If grdData.MaxRows > 0 Then
                        grdData.MaxRows = grdData.MaxRows + 1
                        clsShedGrid.Row = grdData.MaxRows
                        Add_Row clsShedGrid.Row
                        lngFoundID = True
                    Else
                        grdData.MaxRows = grdData.MaxRows + 1
                        clsShedGrid.Row = grdData.MaxRows
                        Add_Row clsShedGrid.Row
                    End If
                End If
            Else
                'Debug.Print
                Delete_Dup_Appt lngApptID
            End If
            lngDup = Val(udtInPatSched.IntId)
        Loop
    End With
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    udtInPatSched.Appt_Date = clsShedGrid.ApptDate
    mbooByPass = False
    mFirstLoad = False
    mbooByPass = False
    udtInPatSched.SchedType = "Patient"
    Get_Changes clsShedGrid.ApptDate
    Set_Headers
    Clear_InSched
    udtInPatSched.IntId = vbNullString
    lblCurScheAppt.Visible = True
    lblCurScheAppt.Caption = "Active Visit Roster"
    lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
    mbooByPass = True
    For i = 0 To cboProviders.listcount - 1
        If cboProviders.List(i) = udtInPatSched.Provider Then
            mbooByPass = True
            cboProviders.Text = cboProviders.List(i)
            Exit For
        End If
    Next i
    'Get Encounter editions
    Screen.MousePointer = vbDefault
    Exit Function
Error_ReadData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "ReadData", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_Room_Number(vlngPatient_ID As Long)
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim strSQL As String
Dim lngPurge As Double
Dim lngPatient_ID As Long

    On Error GoTo Error_Get_Room_Number
    lngPatient_ID = vlngPatient_ID
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    If udtInPatSched.Appointment_ID > 0 Then
        strSQL = "Get_InPatient_Room_Appt " & lngPatient_ID & "," & udtInPatSched.Facility_ID & "," & udtInPatSched.Appointment_ID
    Else
        strSQL = "Get_InPatient_Room " & lngPatient_ID & "," & udtInPatSched.Facility_ID
    End If
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        If Not .EOF Then
            If Not IsNull((!room)) Then
                udtInPatSched.RoomNumber = Trim((!room))
            End If
        End If
    End With
    If udtInPatSched.RoomNumber = vbNullString Then
        strSQL = "Get_InPatient_Room " & lngPatient_ID & "," & udtInPatSched.Facility_ID
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        With rsReadAppt
            If Not .EOF Then
                If Not IsNull((!room)) Then
                    udtInPatSched.RoomNumber = Trim((!room))
                End If
            End If
        End With
    End If
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Room_Number:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Room_Number", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Sub Read_Week_InPatient_Appointments()

End Sub

Private Sub mnu_Workup_Complete_Click()
    Load_Patient_Status "Workup Completed"
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Dim strDay As String
Dim Arr
Dim i%
Dim strSelDate As String

    If mbooByPass Then mbooByPass = False: Exit Sub
    rtfChanges.Text = vbNullString
    strSelDate = MonthView1.Value
    If strSelDate <> vbNullString Then
        rtfChanges.Text = vbNullString
        udtInPatSched.Appt_Date = Format(strSelDate, "MMMM DD, YYYY")
'        lblApptDate.Caption = Format(strSelDate, "MMMM DD, YYYY")
        Get_Mode
    End If
    Set_Headers
   Exit Sub
End Sub

Private Sub picBlock_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picBlock.BackColor = SColor.oSelectedColor
        udtColor.Lunch = picBlock.BackColor
       Get_Mode
       Write_Color "Lunch", udtColor.Lunch
    End If
    Exit Sub
End Sub

Private Sub picCancel_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picCancel.BackColor = SColor.oSelectedColor
        udtColor.Cancelled = picCancel.BackColor
       Get_Mode
       Write_Color "Cancelled", udtColor.Cancelled
    End If
    Exit Sub
End Sub


Private Sub picCheckOut_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picCheckOut.BackColor = SColor.oSelectedColor
        udtColor.CheckOut = picCheckOut.BackColor
       Get_Mode
       Write_Color "CheckOut", udtColor.CheckOut
    End If
    Exit Sub
End Sub

Private Sub picExstNewPat_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picExstNewPat.BackColor = SColor.oSelectedColor
        udtColor.ExistPatient = picExstNewPat.BackColor
       Get_Mode
       Write_Color "ExistPatient", udtColor.ExistPatient
    End If
    Exit Sub
End Sub

Private Sub picHospital_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picHospital.BackColor = SColor.oSelectedColor
        udtColor.Hospital = picHospital.BackColor
       Get_Mode
       Write_Color "Hospital", udtColor.Hospital
    End If
    Exit Sub
End Sub

Private Sub picNewPatient_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picNewPatient.BackColor = SColor.oSelectedColor
        udtColor.NewPatient = picNewPatient.BackColor
       Get_Mode
       Write_Color "NewPatient", udtColor.NewPatient
    End If
    Exit Sub
End Sub

Private Sub picOffice_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picOffice.BackColor = SColor.oSelectedColor
        udtColor.Office = picOffice.BackColor
       Get_Mode
       Write_Color "Office", udtColor.Office
    End If
    Exit Sub
End Sub

Private Sub picPostOp_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picPostOp.BackColor = SColor.oSelectedColor
        udtColor.PostOp = picPostOp.BackColor
       Get_Mode
       Write_Color "PostOp", udtColor.PostOp
    End If
    Exit Sub
End Sub

Private Sub picPreOp_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picPreOp.BackColor = SColor.oSelectedColor
        udtColor.PreOp = picPreOp.BackColor
       Get_Mode
       Write_Color "PreOp", udtColor.PreOp
    End If
    Exit Sub
End Sub

Private Sub picPresent_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picPresent.BackColor = SColor.oSelectedColor
        udtColor.Present = picPresent.BackColor
       Get_Mode
       Write_Color "Present", udtColor.Present
    End If
    Exit Sub
End Sub

Private Sub picReturn_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picReturn.BackColor = SColor.oSelectedColor
        udtColor.FollowUp = picReturn.BackColor
       Get_Mode
       Write_Color "FollowUp", udtColor.FollowUp
    End If
    Exit Sub
End Sub

Private Sub picSurgery_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picSurgery.BackColor = SColor.oSelectedColor
        udtColor.Surgery = picSurgery.BackColor
       Get_Mode
       Write_Color "Surgery", udtColor.Surgery
    End If
    Exit Sub
End Sub
Private Sub schTimer_Timer()
    
'    If schTimer.Enabled Then
'        Check_Change "Appointment"
'        If mbooSchChange Then
'            Get_Mode
'            Get_Changes udtInPatSched.Appt_Date
'        End If
'    End If
    Exit Sub
End Sub

Private Sub strTelTime_Timer()
    Check_Change "Telephone"
    If mbooTeleChange Then
        Format_Messages
        Read_Messages
    End If
    Exit Sub
End Sub

Private Sub txtDate_Change()
    udtInPatSched.Appt_Date = Format(txtDate.Value, "mm/dd/yyyy")
    Exit Sub
End Sub


