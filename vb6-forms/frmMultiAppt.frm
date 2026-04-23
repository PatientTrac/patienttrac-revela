VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmMultiAppt 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Multiple Patient Appointments"
   ClientHeight    =   7815
   ClientLeft      =   8085
   ClientTop       =   3735
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMultiAppt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   13830
   Begin VB.Frame fraControl 
      BackColor       =   &H00E0E0E0&
      Height          =   6105
      Left            =   4080
      TabIndex        =   17
      Top             =   150
      Visible         =   0   'False
      Width           =   9585
      Begin VB.Frame fraPsycho 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   180
         TabIndex        =   48
         Top             =   2340
         Visible         =   0   'False
         Width           =   9075
         Begin FPSpreadADO.fpSpread grdPsychotherapy 
            Height          =   2385
            Left            =   30
            TabIndex        =   49
            Top             =   570
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
            SpreadDesigner  =   "frmMultiAppt.frx":08CA
            UserResize      =   0
            Appearance      =   2
            ClipboardOptions=   0
         End
         Begin PT_XP_Button.PT_XP cmdPsyAdd 
            Height          =   345
            Left            =   3510
            TabIndex        =   50
            Top             =   3210
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
            Picture         =   "frmMultiAppt.frx":0DB9
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Select"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmMultiAppt.frx":1153
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdPsyClose 
            Height          =   345
            Left            =   4650
            TabIndex        =   51
            Top             =   3210
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
            Picture         =   "frmMultiAppt.frx":14A5
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmMultiAppt.frx":1EB7
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdRemAddOn 
            Height          =   315
            Left            =   7680
            TabIndex        =   52
            Top             =   210
            Visible         =   0   'False
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
            Picture         =   "frmMultiAppt.frx":2209
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Remove"
            DefaultThemes   =   2
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmMultiAppt.frx":27A3
            ShowFocus       =   -1  'True
         End
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
         Height          =   330
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   4740
         Width           =   4155
      End
      Begin VB.ComboBox cboStartTime 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         ItemData        =   "frmMultiAppt.frx":2AF5
         Left            =   7740
         List            =   "frmMultiAppt.frx":2AF7
         TabIndex        =   39
         Top             =   1140
         Width           =   1740
      End
      Begin VB.ComboBox cboLocation 
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
         ItemData        =   "frmMultiAppt.frx":2AF9
         Left            =   210
         List            =   "frmMultiAppt.frx":2AFB
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   450
         Width           =   4860
      End
      Begin VB.ComboBox cboDayWeek 
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
         ItemData        =   "frmMultiAppt.frx":2AFD
         Left            =   5760
         List            =   "frmMultiAppt.frx":2B16
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1140
         Width           =   1845
      End
      Begin VB.ComboBox cboResProvider 
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
         Left            =   5295
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   450
         Width           =   4155
      End
      Begin VB.ComboBox cboRepAppt 
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
         ItemData        =   "frmMultiAppt.frx":2B5A
         Left            =   3900
         List            =   "frmMultiAppt.frx":2B6A
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1140
         Width           =   1695
      End
      Begin VB.ComboBox cboRepeatTerm 
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
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1140
         Width           =   1665
      End
      Begin VB.TextBox cctComments 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   4710
         Width           =   4155
      End
      Begin MSComCtl2.DTPicker ccdTermDate 
         Height          =   345
         Left            =   210
         TabIndex        =   23
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
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
         CustomFormat    =   "MM/dd/yyy"
         DateIsNull      =   -1  'True
         Format          =   86048771
         CurrentDate     =   38792
      End
      Begin MSComCtl2.DTPicker ccdStart 
         Height          =   345
         Left            =   210
         TabIndex        =   24
         Top             =   1140
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
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
         CustomFormat    =   "MM/dd/yyy"
         DateIsNull      =   -1  'True
         Format          =   86048771
         CurrentDate     =   38792
      End
      Begin PT_XP_Button.PT_XP cmdSched 
         Height          =   345
         Left            =   3540
         TabIndex        =   36
         Top             =   5640
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
         Picture         =   "frmMultiAppt.frx":2B88
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Apply"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMultiAppt.frx":2F22
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdExit 
         Height          =   345
         Left            =   4800
         TabIndex        =   37
         Top             =   5640
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
         Picture         =   "frmMultiAppt.frx":3274
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "&Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMultiAppt.frx":3C86
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdProcedure 
         Height          =   2055
         Left            =   180
         TabIndex        =   53
         Top             =   2370
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
         SpreadDesigner  =   "frmMultiAppt.frx":3FD8
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin VB.Label lblDuration 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Duration"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   240
         TabIndex        =   47
         Top             =   4500
         Width           =   900
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Facility"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   210
         TabIndex        =   34
         Top             =   180
         Width           =   765
      End
      Begin VB.Label lblDayWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "Day of Week"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   5760
         TabIndex        =   32
         ToolTipText     =   "Frequency Number"
         Top             =   900
         Width           =   1485
      End
      Begin VB.Label lblTermDate 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   210
         TabIndex        =   31
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Provider"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   5295
         TabIndex        =   30
         Top             =   180
         Width           =   855
      End
      Begin VB.Label lblTermValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Term Value"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   3900
         TabIndex        =   29
         ToolTipText     =   "Week, Month, Year"
         Top             =   900
         Width           =   1530
      End
      Begin VB.Label lblTerm 
         BackStyle       =   0  'Transparent
         Caption         =   "Every "
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   2100
         TabIndex        =   28
         ToolTipText     =   "Frequency Number"
         Top             =   900
         Width           =   945
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appt Time"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   7770
         TabIndex        =   27
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   210
         TabIndex        =   26
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label lblComment 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4680
         TabIndex        =   25
         Top             =   4470
         Width           =   1725
      End
   End
   Begin VB.Frame fraSchedConflicts 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   6945
      Left            =   3330
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   10320
      Begin FPSpreadADO.fpSpread grdTerms 
         Height          =   5805
         Left            =   90
         TabIndex        =   42
         Top             =   660
         Width           =   10155
         _Version        =   524288
         _ExtentX        =   17912
         _ExtentY        =   10239
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
         MaxCols         =   5
         MaxRows         =   0
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         ShadowText      =   8388608
         SpreadDesigner  =   "frmMultiAppt.frx":44DA
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin PT_XP_Button.PT_XP cmdExitConflict 
         Height          =   345
         Left            =   4680
         TabIndex        =   43
         ToolTipText     =   "Save, Restore, Close"
         Top             =   6510
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmMultiAppt.frx":49EE
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMultiAppt.frx":5400
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrintList 
         Height          =   375
         Left            =   9180
         TabIndex        =   44
         Top             =   210
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
         Picture         =   "frmMultiAppt.frx":5752
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Print"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMultiAppt.frx":5CEC
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.TextBox cctBlockConflict 
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   15840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2550
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ComboBox cboResource 
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   15720
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   990
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Frame fraNumber 
      BackColor       =   &H00E0E0E0&
      Height          =   1005
      Left            =   1740
      TabIndex        =   0
      Top             =   3660
      Visible         =   0   'False
      Width           =   2025
      Begin VB.ComboBox cboRepTime 
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
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   180
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker ccdDate 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   86048769
         CurrentDate     =   39288
      End
      Begin PT_XP_Button.PT_XP cmdAccept 
         Height          =   285
         Left            =   510
         TabIndex        =   2
         Top             =   630
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
         Picture         =   "frmMultiAppt.frx":603E
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMultiAppt.frx":6A50
         ShowFocus       =   -1  'True
      End
   End
   Begin PT_XP_Button.PT_XP cmdRemoveRow 
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   2910
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
      Picture         =   "frmMultiAppt.frx":6DA2
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Delete Appt"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":733C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdGridRow 
      Height          =   465
      Left            =   17640
      TabIndex        =   5
      ToolTipText     =   "Update Appointment Row"
      Top             =   3450
      Visible         =   0   'False
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
      Picture         =   "frmMultiAppt.frx":768E
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Change Appt"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":7C28
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPrint 
      Height          =   345
      Left            =   12495
      TabIndex        =   6
      Top             =   2910
      Width           =   1155
      _ExtentX        =   2037
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
      Picture         =   "frmMultiAppt.frx":7F7A
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Print"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":8514
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdResource 
      Height          =   285
      Left            =   18390
      TabIndex        =   7
      ToolTipText     =   "Manage Resources"
      Top             =   1020
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMultiAppt.frx":8866
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmMultiAppt.frx":8E00
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdProvTypes 
      Height          =   285
      Left            =   18390
      TabIndex        =   8
      ToolTipText     =   "Provider Types"
      Top             =   1350
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMultiAppt.frx":9152
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmMultiAppt.frx":96EC
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdMultiSave 
      Height          =   375
      Left            =   30
      TabIndex        =   10
      Top             =   1590
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
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
      Picture         =   "frmMultiAppt.frx":9A3E
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "&Apply"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":9DD8
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPatient 
      Height          =   465
      Left            =   90
      TabIndex        =   12
      Top             =   150
      Width           =   1785
      _ExtentX        =   3149
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
      Picture         =   "frmMultiAppt.frx":A12A
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Patient Select"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":A6C4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdConflict 
      Height          =   345
      Left            =   10350
      TabIndex        =   15
      Top             =   2910
      Width           =   2025
      _ExtentX        =   3572
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
      Picture         =   "frmMultiAppt.frx":AA16
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Repeat Conflicts"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":AFB0
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdRepeatHist 
      Height          =   2415
      Left            =   4110
      TabIndex        =   16
      Top             =   150
      Width           =   9600
      _Version        =   524288
      _ExtentX        =   16933
      _ExtentY        =   4260
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
      SpreadDesigner  =   "frmMultiAppt.frx":B302
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   100
      ScrollBarTrack  =   1
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdNewRepeat 
      Height          =   345
      Left            =   2310
      TabIndex        =   35
      Top             =   180
      Width           =   1665
      _ExtentX        =   2937
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
      Picture         =   "frmMultiAppt.frx":B804
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "New Repeat"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":BD9E
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDelete 
      Height          =   345
      Left            =   2310
      TabIndex        =   40
      Top             =   1020
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
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
      Picture         =   "frmMultiAppt.frx":C0F0
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Delete Set"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":C68A
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdMultiAppt 
      Height          =   3945
      Left            =   90
      TabIndex        =   45
      Top             =   3420
      Width           =   13620
      _Version        =   524288
      _ExtentX        =   24024
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
      SpreadDesigner  =   "frmMultiAppt.frx":C9DC
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   100
      ScrollBarTrack  =   1
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdUpdate 
      Height          =   345
      Left            =   2310
      TabIndex        =   54
      Top             =   600
      Width           =   1665
      _ExtentX        =   2937
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
      Picture         =   "frmMultiAppt.frx":CEDE
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Appt Details"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":D478
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblPatName 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   480
      Left            =   4110
      TabIndex        =   14
      Top             =   2850
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.Label lblBlock 
      BackStyle       =   0  'Transparent
      Caption         =   "Block Time Conflicts"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   15870
      TabIndex        =   11
      Top             =   2340
      Width           =   2760
   End
   Begin VB.Label lblRes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Resource"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   15720
      TabIndex        =   9
      Top             =   750
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnu_Report 
      Caption         =   "Report"
      Begin VB.Menu mnu_Print_Report 
         Caption         =   "Print Repeat Schedule"
      End
      Begin VB.Menu mnu_Unscheduled_Appt 
         Caption         =   "Unscheduled Appt"
      End
   End
End
Attribute VB_Name = "frmMultiAppt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Private mstrType As String
Private mbooByPass As Boolean
Private mlngGridRow As Long
Private mlngGridCol As Long
Private fpDeltaLeft As Long
Private fpDeltaTop As Long
Private flngDeltaX  As Long
Private flngDeltaY  As Long
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
Private clsWriteAppt As clsWriteAppt
Private Type clsWriteAppt
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

Private Function Clear_clsWriteAppt()
    With clsWriteAppt
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
 



Private Function Format_Conflict()
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr

    With grdTerms
        .Redraw = False
        .ClearRange 1, 1, .MaxCols, .MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 10
        .MaxCols = 4
        .FontName = "Times New Roman"
        .GridShowHoriz = True
        .GridShowVert = True
        .RowHeadersShow = False
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .MaxRows = 0
        lngRow = SpreadHeader + 1
        .RowHeight(lngRow) = 20
        lngCol = 1
        .ColWidth(lngCol) = 20
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "APPT DATE", "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        .ColWidth(lngCol) = 20
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "APPT TIME", "TypeHAlignCenter", "APPT_TIME", HeaderColor
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        .ColWidth(lngCol) = 20
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "VISIT TYPE", "TypeHAlignCenter", "VISIT_TYPE", HeaderColor
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        .ColWidth(lngCol) = 21.3
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "CONFLICT REASON", "TypeHAlignCenter", "CONFLICT REASON", HeaderColor
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeEdit grdTerms, SpreadHeader, 1, "6", "PATIENT SCHEDULING CONFLICT", "TypeHAlignCenter", "", HeaderColor
        grdTerms.SetCellBorder 1, SpreadHeader, 6, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 2, grdTerms
        .Redraw = False
    End With
    Exit Function
End Function
Private Function Get_Repeat_Data()
Dim lngRow As Long
Dim varCellValue
Dim strDesc As String
Dim Arr
Dim Idx%


    With grdProcedure
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                Exit For
            End If
        Next lngRow
        .GetText 1, lngRow, varCellValue
        udtRepeatAppt.PurposeVisit = varCellValue
        .BlockMode = False
        .GetText 2, lngRow, varCellValue
        udtRepeatAppt.VisitCode = Left(varCellValue, 5)
        Calc_Duration udtRepeatAppt.PurposeVisit
        LookUpCPTCode udtRepeatAppt.PurposeVisit
        If udtRepeatAppt.Duration = vbNullString Then
            With grdProcedure
                .GetText 3, lngRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    strDesc = varCellValue
                     Arr = Split(strDesc, " ")
                    If UBound(Arr) > 0 Then
                        For Idx = LBound(Arr) To UBound(Arr)
                            If IsNumeric(Arr(Idx)) Then
                                udtRepeatAppt.Duration = Trim(Arr(Idx))
                                Exit For
                            End If
                        Next Idx
                        udtRepeatAppt.Duration = udtRepeatAppt.Duration & " " & "minutes"
                    End If
                End If
            End With
        End If
        If udtRepeatAppt.Duration = vbNullString Then udtRepeatAppt.Duration = cboDuration.Text
        udtRepeatAppt.VisitCode = Left(udtRepeatAppt.VisitCode, 5)
    End With
    
End Function

Private Function Get_SchedConflict()
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngCol As Long
Dim lngRow As Long


    On Error GoTo Error_Get_SchedConflict
    Screen.MousePointer = vbHourglass
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(0)
    strSQL = "SELECT PRA.VISIT_TYPE, RC.*  " _
        & "FROM REPEAT_CONFLICT RC  " _
        & "INNER JOIN PATIENT_REPEAT_APPT PRA ON PRA.REPEAT_ID = RC.REPEAT_ID " _
        & "WHERE RC.REPEAT_ID = " & udtRepeatAppt.RepeatID
    Set rsTerms = cnTerms.Execute(strSQL)
    With rsTerms
        Do Until .EOF
            Clear_Conflict
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case Fld.Name
                             Case "APPT_DATE"
                                udtConflict.Appt_Date = Fld.Value
                             Case "APPT_TIME"
                                udtConflict.Appt_Time = Fld.Value
                             Case "PATIENT_ID"
                                udtConflict.Patient_ID = Fld.Value
                             Case "CONFLICT_REASON"
                                udtConflict.Conflict_Reason = Fld.Value
                             Case "APPT_ID"
                                udtConflict.Appt_ID = Fld.Value
                             Case "VISIT_TYPE"
                                udtConflict.Visit_Type = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            With grdTerms
                lngRow = grdTerms.MaxRows + 1
                grdTerms.MaxRows = lngRow
                lngCol = 1
                Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", Format(udtConflict.Appt_Date, "mm/dd/yyyy"), "TypeHAlignCenter", "APPT_DATE"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", Format(udtConflict.Appt_Time, "hh:mm AM/PM"), "TypeHAlignCenter", "APPT_TIME"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 3
                Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", Trim(udtConflict.Visit_Type), "TypeHAlignCenter", "VISIT_TYPE"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 4
                Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", Trim(udtConflict.Conflict_Reason), "TypeHAlignCenter", "CONFLICT REASON"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                .Redraw = False
            End With
        Loop
    End With
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_SchedConflict:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_SchedConflict", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Function Add_Header()
Dim CellNote As Long
Dim strPhoto As String
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr
Dim fso As New FileSystemObject

    On Error GoTo Error_Add_Header
    Screen.MousePointer = vbHourglass
    grdMultiAppt.ClearRange 1, SpreadHeader, grdMultiAppt.MaxCols, SpreadHeader + 2, False
    Remove_Span
    grdMultiAppt.ColHeaderRows = 5
    Get_DefaultFacility
    If udtFacility.Logo <> vbNullString Then
        If fso.FileExists(udtFacility.Logo) Then
            Format_CellTypePicture grdMultiAppt, SpreadHeader, 1, "1", udtFacility.Logo, "TypeHAlignCenter", "LOGO", "75"
            Format_CellTypeEdit grdMultiAppt, SpreadHeader, 2, "10", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
        Else
            Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "10", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
        End If
    End If
        grdMultiAppt.SetCellBorder 1, SpreadHeader + 2, 10, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdMultiAppt
        lngRow = SpreadHeader + 2
        lngCol = 1
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 2, lngCol, "10", "REPEAT APPOINTMENT SCHEDULE FOR" & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName), "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 2, 5, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngRow = SpreadHeader + 3
        lngCol = 1
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "FACILITY", "TypeHAlignCenter", "FACILITY_ID", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "APPT DATE", "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "APPT TIME", "TypeHAlignCenter", "APPT_TIME", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "APPT TYPE", "TypeHAlignCenter", "APPT_TYPE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "DURATION", "TypeHAlignCenter", "DURATION", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "START DATE", "TypeHAlignCenter", "START_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "END DATE", "TypeHAlignCenter", "TERMINATION_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 8
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "WEEK DAY", "TypeHAlignCenter", "REPEAT_DAY_WEEK", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 9
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "COMMENTS", "TypeHAlignCenter", "NOTES", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 10
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "APPT ID", "TypeHAlignCenter", "APPT_ID", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 4, grdMultiAppt
    Screen.MousePointer = vbDefault
    Exit Function
Error_Add_Header:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Add_Header", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function
Private Function Calc_Duration(vstrReason As String) As String
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim Arr

    On Error GoTo Error_Calc_Duration
    Screen.MousePointer = vbHourglass
    clsGridFormat.List = vbNullString
    Arr = Split(vstrReason, " - ")
    If UBound(Arr) > 0 Then
        udtRepeatAppt.PurposeVisit = Arr(0)
        udtRepeatAppt.VisitCode = Arr(1)
    Else
        udtRepeatAppt.PurposeVisit = vstrReason
    End If
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(1)
    strSQL = "Get_Appt_Time_APPT_CODE '" & udtRepeatAppt.PurposeVisit & "'"
    Set rsAppt = cnAppt.Execute(strSQL)
    With rsAppt
        If Not rsAppt.EOF Then
            If (!Appt_Time) <> vbNullString Then udtRepeatAppt.Duration = (!Appt_Time)
'            If Not IsNull((!VISIT_TYPE)) Then
'                If (!VISIT_TYPE) <> vbNullString Then udtTelephone.CallType = (!VISIT_TYPE)
'            End If
        End If
    End With
    Calc_Duration = udtRepeatAppt.Duration
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
Private Function Clear_Fields()
Dim i%

    cboRepeatTerm.ListIndex = -1
    cboRepAppt.ListIndex = -1
    cboDayWeek.ListIndex = -1
    cboStartTime.ListIndex = -1
    cboRepeatTerm.ListIndex = -1
    cboRepTime.ListIndex = -1
    cboDuration.ListIndex = -1
    cctBlockConflict = vbNullString
    cctComments = vbNullString
    Clear_Block
    Exit Function
End Function


Private Function Format_Repeat_History()
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr

    On Error GoTo Error_Format_Repeat_History
    Screen.MousePointer = vbHourglass

    With grdRepeatHist
        .Redraw = False
        .ClearRange 1, SpreadHeader, .MaxCols, .MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .MaxCols = 5
        .FontName = "Times New Roman"
        .GridShowHoriz = True
        .GridShowVert = True
        .RowHeadersShow = False
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .ScrollBars = ScrollBarsBoth
        grdRepeatHist.ColWidth(1) = 23
        grdRepeatHist.ColWidth(2) = 20
        grdRepeatHist.ColWidth(3) = 9
        grdRepeatHist.ColWidth(4) = 9
        grdRepeatHist.ColWidth(5) = 15
        lngRow = SpreadHeader
        lngCol = 1
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "8", "REPEAT APPOINTMENT HISTORY FOR" & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName), "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, 8, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .RowHeight(SpreadHeader) = 15
        lngRow = SpreadHeader + 1
        lngCol = 1
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", "FACILITY", "TypeHAlignCenter", "FACILITY_ID", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", "PROVIDER", "TypeHAlignCenter", "PROVIDER_ID", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", "START DATE", "TypeHAlignCenter", "START_DATE", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", "END DATE", "TypeHAlignCenter", "TERMINATION_DATE", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", "DETAILS", "TypeHAlignCenter", "ACTIVE", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 2, grdRepeatHist
        .Redraw = False
        .MaxRows = 0
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_Repeat_History:
    Screen.MousePointer = vbDefault
End Function

Public Function Get_Appt_Providers()
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
    cboResProvider.Clear
    If udtRepeatAppt.FacilityID = 0 Then udtRepeatAppt.FacilityID = clsFacility.BillFacility
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    strSQL = "Get_Appt_Providers " & udtRepeatAppt.FacilityID
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
                If clsGridFormat.List = vbNullString Then
                    clsGridFormat.List = clsProvider.Name
                Else
                    clsGridFormat.List = clsGridFormat.List & Chr(9) & clsProvider.Name
                End If
                 cboResProvider.AddItem clsProvider.Name
                If IsNumeric(lngProvID) Then cboResProvider.ItemData(cboResProvider.NewIndex) = lngProvID
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

Private Function Load_Form()
Dim Idx As Integer

    Get_Edi_Static
        If udtRepeatAppt.FacilityID = 0 Then udtRepeatAppt.FacilityID = clsFacility.BillFacility
        Load_Resources
        Get_Appt_Providers
        Load_ReportTime
        Load_Facilities
        ccdStart.Value = Now + 1
        If Trim(udtRepeatAppt.WeekDay) <> vbNullString Then
           For Idx = 0 To cboDayWeek.listcount - 1
               If cboDayWeek.List(Idx) = Trim(udtRepeatAppt.WeekDay) Then
                   cboDayWeek.Text = cboDayWeek.List(Idx)
                   Exit For
               End If
           Next Idx
        End If
        If Trim(udtRepeatAppt.ProviderID) > 0 Then
           For Idx = 0 To cboResProvider.listcount - 1
               If cboResProvider.ItemData(Idx) = udtRepeatAppt.ProviderID Then
                   cboResProvider.Text = cboResProvider.List(Idx)
                   Exit For
               End If
           Next Idx
        End If
        If cboRepAppt.listcount <= 0 Then Load_ReportTime
        Format_MultiAppt
        Format_Repeat_History
        If Trim(udtAppointment.SchedDate) <> vbNullString Then
           ccdStart.Value = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
           If DateDiff("d", Now, ccdStart.Value) <= 0 Then ccdStart.Value = Now + 1
        End If
        cctBlockConflict.Text = vbNullString
        cctBlockConflict.Visible = False
        lblBlock.Visible = False
        mbooByPass = True
        For Idx = 0 To cboLocation.listcount - 1
            If cboLocation.ItemData(Idx) = clsFacility.BillFacility Then
                 cboLocation.Text = cboLocation.List(Idx)
                 Exit For
            End If
        Next Idx
        grdMultiAppt.MaxRows = 0
        mbooByPass = False
    Exit Function
End Function

Private Sub Load_ReportTime()

    cboRepeatTerm.Clear
    cboRepeatTerm.AddItem "1"
    cboRepeatTerm.AddItem "2"
    cboRepeatTerm.AddItem "3"
    cboRepeatTerm.AddItem "4"
    cboRepeatTerm.AddItem "5"
    cboRepeatTerm.AddItem "6"
    cboRepeatTerm.AddItem "7"
    cboRepeatTerm.AddItem "8"
    cboRepeatTerm.AddItem "9"
    cboRepeatTerm.AddItem "10"
    cboRepeatTerm.AddItem "12"
    Exit Sub
End Sub

Private Function Load_Resources()
Dim cnApptFormat As New ADODB.Connection
Dim rsApptFormat As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim lngCnt As Long
Dim lngCol As Long
Dim lngRow As Long

    On Error GoTo Error_Load_Resources
    Screen.MousePointer = vbHourglass
    If cnApptFormat.State = adStateClosed Then cnApptFormat.open psConnect(0)
    strSQL = "Load_Resources"
    Set rsApptFormat = cnApptFormat.Execute(strSQL)
    With rsApptFormat
       Do Until .EOF
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        If UCase(Fld.Name) = "RESOURCE_NAME" Then
                            cboResource.AddItem Trim(Fld.Value)
                            cboResource.ItemData(cboResource.NewIndex) = (!RESOURCE_ID)
                        End If
                    End If
                End If
            Next
           .MoveNext
       Loop
    End With
    If cnApptFormat.State = adStateOpen Then
        cnApptFormat.Close
        Set cnApptFormat = Nothing
        Set rsApptFormat = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Resources:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Resources", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Load_ComboTime()
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim strTerms As String

    On Error GoTo Error_Load_ComboTime
    Screen.MousePointer = vbHourglass
    cboStartTime.Clear
    Select Case UCase(clsApptPref.SchedInterval)
        Case "EVERY 15 MIN"
           cboStartTime.AddItem "07:00 AM": cboStartTime.AddItem "07:15 AM": cboStartTime.AddItem "07:30 AM": cboStartTime.AddItem "07:45 AM"
           cboStartTime.AddItem "08:00 AM": cboStartTime.AddItem "08:15 AM": cboStartTime.AddItem "08:30 AM": cboStartTime.AddItem "08:45 AM"
           cboStartTime.AddItem "09:00 AM": cboStartTime.AddItem "09:15 AM": cboStartTime.AddItem "09:30 AM": cboStartTime.AddItem "09:45 AM"
           cboStartTime.AddItem "10:00 AM": cboStartTime.AddItem "10:15 AM": cboStartTime.AddItem "10:30 AM": cboStartTime.AddItem "10:45 AM"
           cboStartTime.AddItem "11:00 AM": cboStartTime.AddItem "11:15 AM": cboStartTime.AddItem "11:30 AM": cboStartTime.AddItem "11:45 AM"
           cboStartTime.AddItem "12:00 PM": cboStartTime.AddItem "12:15 PM": cboStartTime.AddItem "12:30 PM": cboStartTime.AddItem "12:45 PM"
           cboStartTime.AddItem "01:00 PM": cboStartTime.AddItem "01:15 PM": cboStartTime.AddItem "01:30 PM": cboStartTime.AddItem "01:45 PM"
           cboStartTime.AddItem "02:00 PM": cboStartTime.AddItem "02:15 PM": cboStartTime.AddItem "02:30 PM": cboStartTime.AddItem "02:45 PM"
           cboStartTime.AddItem "03:00 PM": cboStartTime.AddItem "03:15 PM": cboStartTime.AddItem "03:30 PM": cboStartTime.AddItem "03:45 PM"
           cboStartTime.AddItem "04:00 PM": cboStartTime.AddItem "04:15 PM": cboStartTime.AddItem "04:30 PM": cboStartTime.AddItem "04:45 PM"
           cboStartTime.AddItem "05:00 PM": cboStartTime.AddItem "05:15 PM": cboStartTime.AddItem "05:30 PM": cboStartTime.AddItem "05:45 PM"
           cboStartTime.AddItem "06:00 PM": cboStartTime.AddItem "06:15 PM": cboStartTime.AddItem "06:30 PM": cboStartTime.AddItem "06:45 PM"
           cboStartTime.AddItem "07:00 PM": cboStartTime.AddItem "07:15 PM": cboStartTime.AddItem "07:30 PM": cboStartTime.AddItem "07:45 PM"
           cboStartTime.AddItem "08:00 PM": cboStartTime.AddItem "08:15 PM": cboStartTime.AddItem "08:30 PM": cboStartTime.AddItem "08:45 PM"
       Case "EVERY 5 MIN"
           cboStartTime.AddItem "07:00 AM": cboStartTime.AddItem "07:05 AM": cboStartTime.AddItem "07:10 AM": cboStartTime.AddItem "07:15 AM"
           cboStartTime.AddItem "07:20 AM": cboStartTime.AddItem "07:25 AM": cboStartTime.AddItem "07:30 AM": cboStartTime.AddItem "07:35 AM"
           cboStartTime.AddItem "07:40 AM": cboStartTime.AddItem "07:45 AM": cboStartTime.AddItem "07:50 AM": cboStartTime.AddItem "07:55 AM"
           cboStartTime.AddItem "08:00 AM": cboStartTime.AddItem "08:05 AM": cboStartTime.AddItem "08:10 AM": cboStartTime.AddItem "08:15 AM"
           cboStartTime.AddItem "08:20 AM": cboStartTime.AddItem "08:25 AM": cboStartTime.AddItem "08:30 AM": cboStartTime.AddItem "08:35 AM"
           cboStartTime.AddItem "08:40 AM": cboStartTime.AddItem "08:45 AM": cboStartTime.AddItem "08:50 AM": cboStartTime.AddItem "08:55 AM"
           cboStartTime.AddItem "09:00 AM": cboStartTime.AddItem "09:05 AM": cboStartTime.AddItem "09:10 AM": cboStartTime.AddItem "09:15 AM"
           cboStartTime.AddItem "09:20 AM": cboStartTime.AddItem "09:25 AM": cboStartTime.AddItem "09:30 AM": cboStartTime.AddItem "09:35 AM"
           cboStartTime.AddItem "09:40 AM": cboStartTime.AddItem "09:45 AM": cboStartTime.AddItem "09:50 AM": cboStartTime.AddItem "09:55 AM"
           cboStartTime.AddItem "10:00 AM": cboStartTime.AddItem "10:05 AM": cboStartTime.AddItem "10:10 AM": cboStartTime.AddItem "10:15 AM"
           cboStartTime.AddItem "10:20 AM": cboStartTime.AddItem "10:25 AM": cboStartTime.AddItem "10:30 AM": cboStartTime.AddItem "10:35 AM"
           cboStartTime.AddItem "10:40 AM": cboStartTime.AddItem "10:45 AM": cboStartTime.AddItem "10:50 AM": cboStartTime.AddItem "10:55 AM"
           cboStartTime.AddItem "11:00 AM": cboStartTime.AddItem "11:05 AM": cboStartTime.AddItem "11:10 AM": cboStartTime.AddItem "11:15 AM"
           cboStartTime.AddItem "11:20 AM": cboStartTime.AddItem "11:25 AM": cboStartTime.AddItem "11:30 AM": cboStartTime.AddItem "11:35 AM"
           cboStartTime.AddItem "11:40 AM": cboStartTime.AddItem "11:45 AM": cboStartTime.AddItem "11:50 AM": cboStartTime.AddItem "11:55 AM"
           cboStartTime.AddItem "12:00 PM": cboStartTime.AddItem "12:05 PM": cboStartTime.AddItem "12:10 PM": cboStartTime.AddItem "12:15 PM"
           cboStartTime.AddItem "12:20 PM": cboStartTime.AddItem "12:25 PM": cboStartTime.AddItem "12:30 PM": cboStartTime.AddItem "12:35 PM"
           cboStartTime.AddItem "01:00 PM": cboStartTime.AddItem "01:05 PM": cboStartTime.AddItem "01:10 PM": cboStartTime.AddItem "01:15 PM"
           cboStartTime.AddItem "01:20 PM": cboStartTime.AddItem "01:25 PM": cboStartTime.AddItem "01:30 PM": cboStartTime.AddItem "01:35 PM"
           cboStartTime.AddItem "01:40 PM": cboStartTime.AddItem "01:45 PM": cboStartTime.AddItem "01:50 PM": cboStartTime.AddItem "01:55 PM"
           cboStartTime.AddItem "02:00 PM": cboStartTime.AddItem "02:05 PM": cboStartTime.AddItem "02:10 PM": cboStartTime.AddItem "02:15 PM"
           cboStartTime.AddItem "02:20 PM": cboStartTime.AddItem "02:25 PM": cboStartTime.AddItem "02:30 PM": cboStartTime.AddItem "02:35 PM"
           cboStartTime.AddItem "02:40 PM": cboStartTime.AddItem "02:45 PM": cboStartTime.AddItem "02:50 PM": cboStartTime.AddItem "02:55 PM"
           cboStartTime.AddItem "03:00 PM": cboStartTime.AddItem "03:05 PM": cboStartTime.AddItem "03:10 PM": cboStartTime.AddItem "03:15 PM"
           cboStartTime.AddItem "03:20 PM": cboStartTime.AddItem "03:25 PM": cboStartTime.AddItem "03:30 PM": cboStartTime.AddItem "03:35 PM"
           cboStartTime.AddItem "03:40 PM": cboStartTime.AddItem "03:45 PM": cboStartTime.AddItem "03:50 PM": cboStartTime.AddItem "03:55 PM"
           cboStartTime.AddItem "04:00 PM": cboStartTime.AddItem "04:05 PM": cboStartTime.AddItem "04:10 PM": cboStartTime.AddItem "04:15 PM"
           cboStartTime.AddItem "04:20 PM": cboStartTime.AddItem "04:25 PM": cboStartTime.AddItem "04:30 PM": cboStartTime.AddItem "04:35 PM"
           cboStartTime.AddItem "04:40 PM": cboStartTime.AddItem "04:45 PM": cboStartTime.AddItem "04:50 PM": cboStartTime.AddItem "04:55 PM"
           cboStartTime.AddItem "05:00 PM": cboStartTime.AddItem "05:05 PM": cboStartTime.AddItem "05:10 PM": cboStartTime.AddItem "05:15 PM"
           cboStartTime.AddItem "05:20 PM": cboStartTime.AddItem "05:25 PM": cboStartTime.AddItem "05:30 PM": cboStartTime.AddItem "05:35 PM"
           cboStartTime.AddItem "05:40 PM": cboStartTime.AddItem "05:45 PM": cboStartTime.AddItem "05:50 PM": cboStartTime.AddItem "05:55 PM"
           cboStartTime.AddItem "06:00 PM": cboStartTime.AddItem "06:05 PM": cboStartTime.AddItem "06:10 PM": cboStartTime.AddItem "06:15 PM"
           cboStartTime.AddItem "06:20 PM": cboStartTime.AddItem "06:25 PM": cboStartTime.AddItem "06:30 PM": cboStartTime.AddItem "06:35 PM"
           cboStartTime.AddItem "06:40 PM": cboStartTime.AddItem "06:45 PM": cboStartTime.AddItem "06:50 PM": cboStartTime.AddItem "06:55 PM"
           cboStartTime.AddItem "07:00 PM": cboStartTime.AddItem "07:05 PM": cboStartTime.AddItem "07:10 PM": cboStartTime.AddItem "07:15 PM"
           cboStartTime.AddItem "07:20 PM": cboStartTime.AddItem "07:25 PM": cboStartTime.AddItem "07:30 PM": cboStartTime.AddItem "07:35 PM"
           cboStartTime.AddItem "07:40 PM": cboStartTime.AddItem "07:45 PM": cboStartTime.AddItem "07:50 PM": cboStartTime.AddItem "07:55 PM"
           cboStartTime.AddItem "08:00 PM": cboStartTime.AddItem "08:05 PM": cboStartTime.AddItem "08:10 PM": cboStartTime.AddItem "08:15 PM"
           cboStartTime.AddItem "08:20 PM": cboStartTime.AddItem "08:25 PM": cboStartTime.AddItem "08:30 PM": cboStartTime.AddItem "08:35 PM"
           cboStartTime.AddItem "08:40 PM": cboStartTime.AddItem "08:45 PM": cboStartTime.AddItem "08:50 PM": cboStartTime.AddItem "08:55 PM"
       Case "EVERY 20 MIN"
           cboStartTime.AddItem "07:00 AM": cboStartTime.AddItem "07:20 AM": cboStartTime.AddItem "07:40 AM"
           cboStartTime.AddItem "08:00 AM": cboStartTime.AddItem "08:20 AM": cboStartTime.AddItem "08:40 AM"
           cboStartTime.AddItem "09:00 AM": cboStartTime.AddItem "09:20 AM": cboStartTime.AddItem "09:40 AM"
           cboStartTime.AddItem "10:00 AM": cboStartTime.AddItem "10:20 AM": cboStartTime.AddItem "10:40 AM"
           cboStartTime.AddItem "11:00 AM": cboStartTime.AddItem "11:20 AM": cboStartTime.AddItem "11:40 AM"
           cboStartTime.AddItem "12:00 PM": cboStartTime.AddItem "12:20 PM": cboStartTime.AddItem "12:40 PM"
           cboStartTime.AddItem "01:00 PM": cboStartTime.AddItem "01:20 PM": cboStartTime.AddItem "01:40 PM"
           cboStartTime.AddItem "02:00 PM": cboStartTime.AddItem "02:20 PM": cboStartTime.AddItem "02:40 PM"
           cboStartTime.AddItem "03:00 PM": cboStartTime.AddItem "03:20 PM": cboStartTime.AddItem "03:40 PM"
           cboStartTime.AddItem "04:00 PM": cboStartTime.AddItem "04:20 PM": cboStartTime.AddItem "04:40 PM"
           cboStartTime.AddItem "05:00 PM": cboStartTime.AddItem "05:20 PM": cboStartTime.AddItem "05:40 PM"
           cboStartTime.AddItem "06:00 PM": cboStartTime.AddItem "06:20 PM": cboStartTime.AddItem "06:40 PM"
           cboStartTime.AddItem "07:00 PM": cboStartTime.AddItem "07:20 PM": cboStartTime.AddItem "07:40 PM"
           cboStartTime.AddItem "08:00 PM": cboStartTime.AddItem "08:20 PM": cboStartTime.AddItem "08:40 PM"
       Case "EVERY 10 MIN"
        cboStartTime.AddItem "07:00 AM": cboStartTime.AddItem "07:10 AM": cboStartTime.AddItem "07:20 AM": cboStartTime.AddItem "07:30 AM": cboStartTime.AddItem "07:40 AM": cboStartTime.AddItem "07:50 AM"
        cboStartTime.AddItem "08:00 AM": cboStartTime.AddItem "08:10 AM": cboStartTime.AddItem "08:20 AM": cboStartTime.AddItem "08:30 AM": cboStartTime.AddItem "08:40 AM": cboStartTime.AddItem "08:50 AM"
        cboStartTime.AddItem "09:00 AM": cboStartTime.AddItem "09:10 AM": cboStartTime.AddItem "09:20 AM": cboStartTime.AddItem "09:30 AM": cboStartTime.AddItem "09:40 AM": cboStartTime.AddItem "09:50 AM"
        cboStartTime.AddItem "10:00 AM": cboStartTime.AddItem "10:10 AM": cboStartTime.AddItem "10:20 AM": cboStartTime.AddItem "10:30 AM": cboStartTime.AddItem "10:40 AM": cboStartTime.AddItem "10:50 AM"
        cboStartTime.AddItem "11:00 AM": cboStartTime.AddItem "11:10 AM": cboStartTime.AddItem "11:20 AM": cboStartTime.AddItem "11:30 AM": cboStartTime.AddItem "11:40 AM": cboStartTime.AddItem "11:50 AM"
        cboStartTime.AddItem "12:00 PM": cboStartTime.AddItem "12:10 PM": cboStartTime.AddItem "12:20 PM": cboStartTime.AddItem "12:30 PM": cboStartTime.AddItem "12:40 PM": cboStartTime.AddItem "12:50 PM"
        cboStartTime.AddItem "01:00 PM": cboStartTime.AddItem "01:10 PM": cboStartTime.AddItem "01:20 PM": cboStartTime.AddItem "01:30 PM": cboStartTime.AddItem "01:40 PM": cboStartTime.AddItem "01:50 PM"
        cboStartTime.AddItem "02:00 PM": cboStartTime.AddItem "02:10 PM": cboStartTime.AddItem "02:20 PM": cboStartTime.AddItem "02:30 PM": cboStartTime.AddItem "02:40 pm": cboStartTime.AddItem "02:50 PM"
        cboStartTime.AddItem "03:00 PM": cboStartTime.AddItem "03:10 PM": cboStartTime.AddItem "03:20 PM": cboStartTime.AddItem "03:30 PM": cboStartTime.AddItem "03:40 PM": cboStartTime.AddItem "03:50 PM"
        cboStartTime.AddItem "04:00 PM": cboStartTime.AddItem "04:10 PM": cboStartTime.AddItem "04:20 PM": cboStartTime.AddItem "04:30 PM": cboStartTime.AddItem "04:40 PM": cboStartTime.AddItem "04:50 PM"
        cboStartTime.AddItem "05:00 PM": cboStartTime.AddItem "05:10 PM": cboStartTime.AddItem "05:20 PM": cboStartTime.AddItem "05:30 PM": cboStartTime.AddItem "05:40 PM": cboStartTime.AddItem "05:50 PM"
        cboStartTime.AddItem "06:00 PM": cboStartTime.AddItem "06:10 PM": cboStartTime.AddItem "06:20 PM": cboStartTime.AddItem "06:30 PM": cboStartTime.AddItem "06:40 PM": cboStartTime.AddItem "06:50 PM"
        cboStartTime.AddItem "07:00 PM": cboStartTime.AddItem "07:10 PM": cboStartTime.AddItem "07:20 PM": cboStartTime.AddItem "07:30 PM": cboStartTime.AddItem "07:40 PM": cboStartTime.AddItem "07:50 PM"
        cboStartTime.AddItem "08:00 PM": cboStartTime.AddItem "08:10 PM": cboStartTime.AddItem "08:20 PM": cboStartTime.AddItem "08:30 PM": cboStartTime.AddItem "08:40 PM": cboStartTime.AddItem "08:50 PM"
     End Select
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_ComboTime:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_ComboTime", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
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
                strSQL = "SELECT * FROM CPT_TEMPLATE WHERE CPT_CODE like '%90%' AND " _
                    & "CONVERT(DATETIME, INSERT_DATE, 101) > '12/31/2012'  AND PROVIDER = 'MD' ORDER BY CPT_CODE"
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
            strSQL = "SELECT * FROM CPT_TEMPLATE WHERE PROVIDER = '" & Replace(clsProvider.Title, ".", "") & "' AND CONVERT(DATETIME, INSERT_DATE,101) > '12/31/2012'  ORDER BY CPT_CODE "
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
                        If Authorized_Prescriber(udtRepeatAppt.ProviderID) Then
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
'                                Case "90832", "90834", "90837", "90845", "90846", "90847", "90849", "90853"
'                                   Format_CellTypeStatic grdPsychotherapy, lngRow, 3, "1", "1", "TypeHAlignLeft", "CPT_UNITS"
'                                   grdPsychotherapy.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
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
'                         lngRow = lngRow + 1
'                        grdPsychotherapy.MaxRows = lngRow
'                        Format_CellTypeStatic grdPsychotherapy, lngRow, 1, "1", strCode, "TypeHAlignCenter", "CPT_CODE"
'                        Format_CellTypeStatic grdPsychotherapy, lngRow, 2, "1", strData, "TypeHAlignLeft", "CPT_NAME"
'                        Select Case strCode
'                            Case "90832", "90834", "90837", "90845", "90846", "90847", "90849", "90853"
'                               Format_CellTypeStatic grdPsychotherapy, lngRow, 3, "1", "1", "TypeHAlignLeft", "CPT_UNITS"
'                               grdPsychotherapy.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                            Case "90840", "90785"
'                                Format_CellTypeEdit grdPsychotherapy, lngRow, 3, "1", " ", "TypeHAlignLeft", "CPT_UNITS"
'                                grdPsychotherapy.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                        End Select
'                        grdPsychotherapy.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                        grdPsychotherapy.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
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
            .CellType = CellTypePicture
           .TypePictMaintainScale = True
           .TypePictStretch = True
            .TypePictCenter = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf vstrSpan <> vbNullString Then
               lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
             End If
            If FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            If vlngRowHeight > 0 Then .RowHeight(vlngRow) = vlngRowHeight
          .BlockMode = False
    End With
End Function


Private Function Format_CellTypeCombo(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, _
vstrString As String, vstrHAlign As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String, Optional vstrHndw As Long)
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
              .TypeVAlign = TypeVAlignCenter
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
          .BlockMode = False
    End With
End Function
Private Function Read_Parent_Visit()
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Idx%
Dim lngRow As Long
Dim varCellValue



    On Error GoTo Error_Read_Parent_Visit
    Screen.MousePointer = vbHourglass
    If cnSched.State = adStateClosed Then cnSched.open psConnect(0)
    udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    strSQL = "SELECT * FROM PATIENT_REPEAT_APPT WHERE REPEAT_ID = " & udtRepeatAppt.RepeatID
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            If Not .EOF Then
                Clear_RepeatAppt
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "PROVIDER_ID"
                                    udtRepeatAppt.ProviderID = Trim(Fld.Value)
                                Case "RESOURCE_ID"
                                    udtRepeatAppt.ResourseID = Trim(Fld.Value)
                                Case "DURATION"
                                    udtRepeatAppt.Duration = Trim(Fld.Value)
                                Case "FACILITY_ID"
                                    udtRepeatAppt.FacilityID = Trim(Fld.Value)
                                Case "START_TIME"
                                    udtRepeatAppt.StartTime = Trim(Fld.Value)
                                Case "END_TIME"
                                    udtRepeatAppt.EndTime = Trim(Fld.Value)
                                Case "TERMINATION_DATE"
                                    udtRepeatAppt.EndDate = Trim(Fld.Value)
                                Case "START_DATE"
                                    udtRepeatAppt.StartDate = Trim(Fld.Value)
                                Case "START_TIME"
                                    udtRepeatAppt.StartTime = Trim(Fld.Value)
                                Case "END_TIME"
                                    udtRepeatAppt.EndTime = Trim(Fld.Value)
                                Case "REPEAT_MODE"
                                    udtRepeatAppt.RepeatMode = Trim(Fld.Value)
                                Case "REPEAT_ID"
                                    udtRepeatAppt.RepeatID = Trim(Fld.Value)
                                Case "REPEAT_FREQ"
                                    udtRepeatAppt.RepeatNumber = Trim(Fld.Value)
                                Case "ACTIVE"
                                    udtRepeatAppt.Status = Trim(Fld.Value)
                                    If udtRepeatAppt.Status = "True" Then
                                        udtRepeatAppt.Status = "Active"
                                    Else
                                        udtRepeatAppt.Status = "InActive"
                                    End If
                                Case "ADD_PROC_CODE"
                                    udtRepeatAppt.Add_Proc = Trim(Fld.Value)
                                Case "REPEAT_DAY_WEEK"
                                    udtRepeatAppt.WeekDay = Trim(Fld.Value)
                                Case "ADD_PROC_DESC"
                                    udtRepeatAppt.Add_Proc_DESC = Trim(Fld.Value)
                                Case "VISIT_CODE"
                                    udtRepeatAppt.VisitCode = Trim(Fld.Value)
                                Case "COMMENT"
                                    udtRepeatAppt.Comment = Trim(Fld.Value)
                                Case "VISIT_TYPE"
                                    udtRepeatAppt.PurposeVisit = Trim(Fld.Value)
                            End Select
                        End If
                    End If
                Next
                If Not TextFound("minutes", udtRepeatAppt.Duration, Len("minutes"), True) Then
                    udtRepeatAppt.Duration = Trim(udtRepeatAppt.Duration) & " " & "minutes"
                End If
                cctBlockConflict.Text = udtRepeatAppt.Comment
                If Trim(udtRepeatAppt.WeekDay) <> vbNullString Then
                   For Idx = 0 To cboDayWeek.listcount - 1
                       If cboDayWeek.List(Idx) = Trim(udtRepeatAppt.WeekDay) Then
                           cboDayWeek.Text = cboDayWeek.List(Idx)
                           Exit For
                       End If
                   Next Idx
                End If
                ccdStart.Value = udtRepeatAppt.StartDate
                ccdTermDate.Value = udtRepeatAppt.EndDate
                mbooByPass = True
                For Idx = 0 To cboRepAppt.listcount - 1
                    If cboRepAppt.List(Idx) = udtRepeatAppt.RepeatMode Then
                         cboRepAppt.Text = cboRepAppt.List(Idx)
                         Exit For
                    End If
                Next Idx
                mbooByPass = True
                For Idx = 0 To cboRepeatTerm.listcount - 1
                    If cboRepeatTerm.List(Idx) = udtRepeatAppt.RepeatNumber Then
                         cboRepeatTerm.Text = cboRepeatTerm.List(Idx)
                         Exit For
                    End If
                Next Idx
                mbooByPass = True
                For Idx = 0 To cboStartTime.listcount - 1
                    If cboStartTime.List(Idx) = Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") Then
                         cboStartTime.Text = cboStartTime.List(Idx)
                         Exit For
                    End If
                Next Idx
                mbooByPass = True
                For Idx = 0 To cboLocation.listcount - 1
                    If cboLocation.ItemData(Idx) = udtRepeatAppt.FacilityID Then
                         cboLocation.Text = cboLocation.List(Idx)
                         Exit For
                    End If
                Next Idx
                mbooByPass = True
                For Idx = 0 To cboDuration.listcount - 1
                    If Trim(cboDuration.List(Idx)) = Replace(Trim(udtRepeatAppt.Duration), "  ", " ") Then
                         cboDuration.Text = cboDuration.List(Idx)
                         Exit For
                    End If
                Next Idx
                mbooByPass = True
                If Trim(udtRepeatAppt.ProviderID) > 0 Then
                   For Idx = 0 To cboResProvider.listcount - 1
                       If cboResProvider.ItemData(Idx) = udtRepeatAppt.ProviderID Then
                           cboResProvider.Text = cboResProvider.List(Idx)
                           Exit For
                       End If
                   Next Idx
                End If
                mbooByPass = True
                If grdProcedure.MaxRows = 0 Then
                     Load_Visit LCase(udtTelephone.CallType)
                End If
                With grdProcedure
                    For lngRow = 1 To grdProcedure.MaxRows
                         .Col = 1
                         .Row = lngRow
                         .GetText 1, lngRow, varCellValue
                         If Trim(varCellValue) <> vbNullString Then
                             If UCase(varCellValue) = UCase(udtRepeatAppt.PurposeVisit) Then
                                 If udtTelephone.CallType = "Patient" Or udtTelephone.CallType = vbNullString Then
                                     Format_CellTypeStatic grdProcedure, lngRow, 1, "1", udtRepeatAppt.PurposeVisit, "TypeHAlignCenter", "VISIT TYPE"
                                     grdProcedure.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                     Format_CellTypeStatic grdProcedure, lngRow, 3, "1", udtRepeatAppt.Add_Proc_DESC, "TypeHAlignCenter", "ADD_DESC"
                                     grdProcedure.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                     Format_CellTypeStatic grdProcedure, lngRow, 4, "1", udtRepeatAppt.Add_Proc, "TypeHAlignCenter", "ADD_CPT_CODE"
                                     grdProcedure.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                 Else
                                     Format_CellTypeStatic grdProcedure, lngRow, 1, "2", udtRepeatAppt.PurposeVisit, "TypeHAlignCenter", "VISIT TYPE"
                                     grdProcedure.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                     Format_CellTypeStatic grdProcedure, lngRow, 3, "2", udtRepeatAppt.Duration, "TypeHAlignCenter", "DURATION"
                                     grdProcedure.SetCellBorder 3, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                 End If
                                 .BlockMode = True
                                     .Col = 1
                                     .Col2 = .MaxCols
                                     .Row = lngRow
                                     .Row2 = lngRow
                                     .BackColor = grdCurItem
                                 .BlockMode = False
                                 .TopRow = lngRow
                                 Exit For
                             End If
                         End If
                    Next lngRow
                End With
            End If
        End With
    If cnSched.State = adStateOpen Then
        cnSched.Close
        Set cnSched = Nothing
        Set rsSched = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Parent_Visit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Parent_Visit", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
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
    For i = 0 To cboLocation.listcount - 1
        If cboLocation.List(i) = cboLocation.Text Then
            lngFacID = cboLocation.ItemData(i)
            Exit For
        End If
    Next i
    Screen.MousePointer = vbHourglass
    grdProcedure.ClearRange 1, 1, grdProcedure.MaxCols, grdProcedure.MaxRows, False
    grdProcedure.MaxRows = 0
    Get_Providers udtRepeatAppt.ProviderID
    If LCase(vstrType) <> "vendor" Then
        strCompare = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
        strCompare = Replace(strCompare, ".", "")
        strCompare = Trim(strCompare)
        clsPatAcct.CPT_Code = vbNullString
    End If
    If udtTelephone.CallType = vbNullString Then udtTelephone.CallType = "Patient"
    If udtRepeatAppt.ProviderID > 0 Then Get_Providers udtRepeatAppt.ProviderID
    strProvider = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    Select Case LCase(udtTelephone.CallType)
        Case "vendor", "vendors", "business"
            strPatType = "business"
        Case "patient"
            strPatType = "Established Patient"
    End Select
    If Trim(vstrType) <> vbNullString Then
        strSQL = "SELECT DISTINCT VISIT_TYPE, APPT_CODE, PROVIDER_TITLE, CPT_CODE, APPT_TIME " _
            & "FROM APPOINTMENT_TYPE WHERE VISIT_TYPE like '" & strPatType & "%' AND (PROVIDER_TITLE = 'ALL' OR PROVIDER_TITLE = '" & strProvider & "') ORDER BY  APPT_CODE, VISIT_TYPE"
    Else
        strSQL = "SELECT DISTINCT VISIT_TYPE, APPT_CODE, PROVIDER_TITLE, APPT_TIME " _
            & "FROM APPOINTMENT_TYPE WHERE  VISIT_TYPE like '" & strPatType & "%' AND " _
            & "FACILITY_ID = " & udtRepeatAppt.FacilityID & "  AND (PROVIDER_TITLE = 'ALL' OR PROVIDER_TITLE = '" & strProvider & "') ORDER BY  APPT_CODE, VISIT_TYPE"
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
'            If Trim(txtDate.Value) <> vbNullString Then
'                If DateDiff("d", "01/01/2013", Format(txtDate.Value, "mm/dd/yyyy")) > 1 And IsProperCase(strApptCode) = "Medication Management" Then
'                    strApptCode = vbNullString
'                End If
'            End If
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
                    If IsProperCase(udtTelephone.CallType) = "Patient" Or udtTelephone.CallType = vbNullString Then
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
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Visit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Visit", App.EXEName & " " & Me.Name
    Exit Function
    Resume Next
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
        .Redraw = False
        .Col = vlngCol
        .Row = vlngRow
        If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        If .CellType = CellTypeComboBox Then
            For Idx = 0 To .TypeComboBoxCount
                .TypeComboBoxIndex = Idx
                If Trim(vstrData) = Trim(.TypeComboBoxString) Then
                    .TypeComboBoxCurSel = Idx
                    Exit For
                End If
            Next Idx
        End If
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
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
    .Redraw = True
    End With
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAlign As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long)
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
vstrHAlign As String, Optional vstrCellTag As String, Optional vstrVAlign As String, Optional lngBackColor As Long, Optional vstrCellNote As String)
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
            If vstrVAlign = "TypeVAlignTop" Then
                .TypeVAlign = TypeVAlignTop
            ElseIf vstrVAlign = "TypeVAlignBottom" Then
                .TypeVAlign = TypeVAlignBottom
            Else
               .TypeVAlign = TypeVAlignCenter
            End If
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

Private Function Get_SchedConflict_Old(vstrDate As String, vstrTime As String) As Boolean
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim sValue As String
Dim intCount As Integer

    Get_SchedConflict_Old = False
    On Error GoTo Error_Get_SchedConflict_Old
    udtRepeatAppt.ConflictReason = vbNullString
    Screen.MousePointer = vbHourglass
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Get_SchedConflict " & udtRepeatAppt.ProviderID & ", " & udtRepeatAppt.FacilityID & ", '" & Format(vstrDate, "mm/dd/yyyy") & "'"
    Set rsBlock = cnTerms.Execute(strSQL)
    With rsBlock
        Do Until .EOF
            Clear_Block
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case Fld.Name
                            Case "DURATION"
                                udtBlock.Duration = Trim(Fld.Value)
                            Case "START_TIME"
                                udtBlock.StartTime = Trim(Fld.Value)
                            Case "END_TIME"
                                udtBlock.EndTime = Trim(Fld.Value)
                            Case "START_DATE"
                                udtBlock.StartDate = Fld.Value
                            Case "END_DATE"
                                udtBlock.EndDate = Fld.Value
                        End Select
                End If
            End If
            Next
            .MoveNext
            udtBlock.StartTime = Replace(udtBlock.StartTime, ".", "")
            If IsDate(vstrDate) And IsDate(udtBlock.StartDate) Then
                lngDateDiff = DateDiff("d", Format(vstrDate, "mm/dd/yyyy"), Format(udtBlock.StartDate, "mm/dd/yyyy"))
                If lngDateDiff = 0 Then
                    lngDateDiff = DateDiff("d", Format(vstrDate, "mm/dd/yyyy"), Format(udtBlock.EndDate, "mm/dd/yyyy"))
                    
                Else
                    If DateDiff("d", Format(udtBlock.EndDate, "mm/dd/yyyy"), Format(udtBlock.StartDate, "mm/dd/yyyy")) < 0 Then
                        If DateDiff("n", Format(udtBlock.StartTime, "hh:mm AM/PM"), Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM")) > 0 Then
                            If DateDiff("n", Format(udtBlock.EndTime, "hh:mm AM/PM"), Format(DateAdd("n", Val(udtRepeatAppt.Duration), udtRepeatAppt.ApptTime), "hh:mm AM/PM")) < 0 Then
                                Get_SchedConflict_Old = True
                                udtRepeatAppt.ConflictReason = "Block Times"
                                Exit Do
                            End If
                        ElseIf DateDiff("n", Format(udtBlock.StartTime, "hh:mm AM/PM"), Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM")) < 0 Then
                            If DateDiff("n", Format(udtBlock.StartTime, "hh:mm AM/PM"), Format(Format(DateAdd("n", Val(udtRepeatAppt.Duration), udtRepeatAppt.ApptTime), "hh:mm AM/PM"), "hh:mm AM/PM")) > 0 Then
                                Get_SchedConflict_Old = True
                                udtRepeatAppt.ConflictReason = "Block Times"
                                Exit Do
                            End If
                        ElseIf DateDiff("n", Format(udtBlock.StartTime, "hh:mm AM/PM"), Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM")) = 0 Then
                            Get_SchedConflict_Old = True
                            udtRepeatAppt.ConflictReason = "Block Times"
                            Exit Do
                        End If
                    End If
                    lngDateDiff = -1
                End If
                If lngDateDiff < 0 Then
                  If udtBlock.Duration <> "Permanent" Then
                        lngDateDiff = DateDiff("h", Format(vstrTime, "hh:mm AM/PM"), Format(udtBlock.StartTime, "hh:mm AM/PM"))
                        If lngDateDiff >= 0 Then
                            lngDateDiff = DateDiff("h", Format(vstrTime, "hh:mm AM/PM"), Format(udtBlock.StartTime, "hh:mm AM/PM"))
                        End If
                        If lngDateDiff < 0 Then
                            Get_SchedConflict_Old = True
                            udtRepeatAppt.ConflictReason = "Block Times"
                            Exit Do
                        End If
                    End If
                ElseIf lngDateDiff = 0 Then
                    Get_SchedConflict_Old = True
                    udtRepeatAppt.ConflictReason = "Block Times"
                    Exit Do
                End If
            End If
        Loop
    End With

    udtRepeatAppt.Duration = Replace(udtRepeatAppt.Duration, "minutes", "")
    sValue = DateAdd("n", Val(udtRepeatAppt.Duration), udtRepeatAppt.ApptTime)
    cnTerms.Close
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(0)
    strSQL = "SELECT * FROM APPOINTMENTS WHERE PROVIDER_ID = " & udtRepeatAppt.ProviderID & " AND " _
        & "FACILITY_ID = " & udtRepeatAppt.FacilityID & " AND  " _
        & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "' AND " _
        & "CONVERT(DATETIME, APPT_TIME, 108)  BETWEEN '" _
        & Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM") & "' AND '" _
        & Format(sValue, "hh:mm AM/PM") & "' AND ACTIVE = 1 "
    Set rsBlock = cnTerms.Execute(strSQL)
    intCount = 0
    With rsBlock
        Do Until .EOF
            If Not IsNull(!Patient_ID) Then intCount = intCount + 1
            .MoveNext
        Loop
    End With
    If intCount > 0 Then
        Read_OverBook
        If intCount >= Val(clsApptPref.OverbookNumber) Then
            Get_SchedConflict_Old = True
            udtRepeatAppt.ConflictReason = "Overbook Appt"
        End If
    End If
    If Validate_Between_Date(CLng(intCount)) Then
        Get_SchedConflict_Old = True
        udtRepeatAppt.ConflictReason = "Overbook Appt"
    End If
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsBlock = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_SchedConflict_Old:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_SchedConflict_Old", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Update_Repeat_Appt()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Arr
Dim strTerm As String
Dim lngDateDiff As Long

    On Error GoTo Error_Repeat_Appt
    Screen.MousePointer = vbHourglass
    If Not TextFound("minutes", udtRepeatAppt.Duration, Len("minutes"), True) Then
        udtRepeatAppt.Duration = Trim(udtRepeatAppt.Duration) & " " & "minutes"
    End If
    strSQL = "UPDATE APPOINTMENTS SET PURPOSE_VISIT = '" & udtRepeatAppt.PurposeVisit & "', " _
            & "APPT_TIME = '" & Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM") & "', NOTES = '" & udtRepeatAppt.Comment & "', " _
            & "APPT_TYPE = '" & udtTelephone.CallType & "', APPT_TYPE = '" & udtRepeatAppt.ApptType & "', " _
            & "TERMINATION_DATE = '" & udtRepeatAppt.EndDate & "',DURATION = '" & udtRepeatAppt.Duration & "', " _
            & "VISIT_CODE = '" & udtRepeatAppt.VisitCode & "', ACTIVE = 1 " _
            & "WHERE REPEAT_APPOINT = " & udtRepeatAppt.RepeatID
        Set rsCodes = cnCodes.Execute(strSQL)
    If Trim(clsWriteAppt.Add_Proc_DESC) <> vbNullString Then
         strSQL = "UPDATE APPOINTMENTS SET ADD_PROC_CODE = '" & clsWriteAppt.Add_Proc_DESC & "', ADD_PROC_DESC = '" _
            & clsWriteAppt.Add_Proc_DESC & "', ACTIVE = 1  WHERE REPEAT_APPOINT = " & udtRepeatAppt.RepeatID
         Set rsCodes = cnCodes.Execute(strSQL)
    End If
    If IsDate(udtRepeatAppt.EndDate) Then
        strSQL = "SELECT max(APPT_DATE) APPT_DATE FROM APPOINTMENTS WHERE REPEAT_ID = " & udtRepeatAppt.RepeatID
        Set rsCodes = cnCodes.Execute(strSQL)
        With rsCodes
            If Not .EOF Then
                strTerm = ((!Appt_Date))
            End If
        End With
        If DateDiff("d", Format(udtRepeatAppt.EndDate, "mm/dd/yyyy"), Format(strTerm, "mm/dd/yyyy")) > 0 Then
            strSQL = "DELETE APPOINTMENTS WHERE " _
                & "REPEAT_ID = " & udtRepeatAppt.RepeatID & " AND CONVERT(DATETIME, APPT_DATE, 101) > '" & Format(udtRepeatAppt.EndDate, "mm/dd/yyyy")
             Set rsCodes = cnCodes.Execute(strSQL)
        ElseIf DateDiff("d", Format(udtRepeatAppt.EndDate, "mm/dd/yyyy"), Format(strTerm, "mm/dd/yyyy")) < 0 Then
            strSQL = "DELETE APPOINTMENTS WHERE REPEAT_ID = " & udtRepeatAppt.RepeatID
            udtRepeatAppt.RepeatID = 0
            Repeat_Appt udtRepeatAppt.StartDate
        End If
    End If
    Read_MultiAppt Val(udtSchedPatient.IntId)
    Screen.MousePointer = vbDefault
    Exit Function
Error_Repeat_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Repeat_Appt", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Update_Row()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim Arr
Dim lngApptID As Long
Dim strApptDate As String
Dim strApptTime As String
Dim bRowFound As Boolean

    On Error GoTo Error_Update_Row
    Screen.MousePointer = vbHourglass
    bRowFound = False
    With grdMultiAppt
        For lngRow = 1 To .MaxRows
            udtRepeatAppt.Repeat_Appt_ID = 0
            udtRepeatAppt.RepeatID = 0
            .BlockMode = True
                .Col = 1
                .Row = lngRow
                If .BackColor = grdCurItem Then
                    bRowFound = True
                    udtRepeatAppt.Repeat_Appt_ID = .GetRowItemData(lngRow)
                    lngCol = 0
                    Arr = Split(LookUpCellTag("APPT_DATE"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            .GetText lngCol, lngRow, varCellValue
                            If IsDate(varCellValue) Then
                                strApptDate = Format(varCellValue, "mm/dd/yyyy")
                            End If
                        End If
                    End If
                    lngCol = 0
                    Arr = Split(LookUpCellTag("APPT_ID"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            .GetText lngCol, lngRow, varCellValue
                            udtRepeatAppt.Repeat_Appt_ID = varCellValue
                        End If
                    End If
                    lngCol = 0
                    Arr = Split(LookUpCellTag("APPT_TIME"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            .GetText lngCol, lngRow, varCellValue
                            strApptTime = Format(varCellValue, "hh:mm AM/PM")
                        End If
                    End If
                    Exit For
                End If
            .BlockMode = False
        Next lngRow
        If bRowFound = False Then
           MsgBox "Select Appointment Row First.", vbInformation, "PatientTrac Scheduling"
           Screen.MousePointer = vbDefault
           Exit Function
        End If
    End With
    If udtRepeatAppt.Repeat_Appt_ID > 0 Then
        If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
        strSQL = "Get_Appointment_Appt_ID " & udtRepeatAppt.Repeat_Appt_ID
        Set rsCodes = cnCodes.Execute(strSQL)
        With rsCodes
            If Not IsNull((!Appt_Time)) Then
                If (!Appt_Time) <> vbNullString Then udtRepeatAppt.StartTime = Format((!Appt_Time), "hh:mm AM/PM")
                If (!Appt_Date) <> vbNullString Then udtRepeatAppt.ApptDate = Format((!Appt_Date), "mm/dd/yyyy")
            End If
        End With
        strSQL = "Update_Appointment_Date_Time '" & Format(strApptDate, "mm/dd/yyyy") & "', '" _
            & Format(strApptTime, "hh:mm AM/PM") & "', '" _
            & Literals(udtRepeatAppt.Comment, "'") & "'," & udtRepeatAppt.Repeat_Appt_ID
        Set rsCodes = cnCodes.Execute(strSQL)
        MsgBox "Successfully updated Appointment from" & " " & udtRepeatAppt.ApptDate & " " & "at" & " " & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") _
            & " " & "to" & " " & Format(strApptDate, "hh:mm AM/PM") & " " & "at" & " " & strApptTime
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_Row:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_Row", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Write_Repeat_Appt()
Dim lngProviderID As Long
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim booinsert As Boolean
Dim strReason As String
Dim strConfirm As String
Dim strRepeatAppt As String
Dim nDays As Long
Dim i%
Dim lngApptID  As Long
Dim vTime
Dim Arr
Dim lngResID As Long
Dim booExit As Boolean
Dim strTerms As String
Dim lngRow As Long
Dim varCellValue
Dim strPatientType As String

    On Error GoTo Error_Write_Repeat_Appt
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    With grdProcedure
        clsWriteAppt.Add_Proc_CPT = vbNullString
        clsWriteAppt.Add_Proc_DESC = vbNullString
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                .GetText 1, lngRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    udtRepeatAppt.PurposeVisit = Trim(varCellValue)
                End If
                .GetText 2, lngRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    udtRepeatAppt.VisitCode = Trim(varCellValue)
                End If
                If IsProperCase(udtTelephone.CallType) = "Patient" Or udtTelephone.CallType = vbNullString Then
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
                        udtRepeatAppt.Duration = Trim(varCellValue)
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
    If cboDuration <> udtRepeatAppt.Duration Then udtRepeatAppt.Duration = Trim(cboDuration)
    If udtRepeatAppt.VisitCode <> vbNullString And udtRepeatAppt.Duration = vbNullString Then
        MsgBox "Please select a duration for the visit!", vbInformation, "PatientTrac Appointment Scheduling"
        Screen.MousePointer = vbDefault
        Exit Function
    ElseIf udtRepeatAppt.VisitCode = vbNullString And udtRepeatAppt.Duration = vbNullString Then
        MsgBox "Please select a duration and a purpose for the visit of the appointment first!", vbInformation, "PatientTrac Appointment Scheduling"
        Screen.MousePointer = vbDefault
        Exit Function
    ElseIf (udtRepeatAppt.VisitCode = vbNullString And udtTelephone.CallType = "vendors") And udtRepeatAppt.Duration <> vbNullString Then
        MsgBox "Please first select a Purpose for the visit for the appointment.", vbInformation, "PatientTrac Appointment Scheduling"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If Not TextFound("minutes", udtRepeatAppt.Duration, Len("minutes"), True) Then
        udtRepeatAppt.Duration = Trim(udtRepeatAppt.Duration) & " " & "minutes"
    End If
    udtRepeatAppt.Repeat_Appt_ID = 0
    strSQL = "SELECT * FROM APPOINTMENTS where PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
        & "FACILITY_ID = " & udtRepeatAppt.FacilityID & " AND " _
        & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "' " _
        & "AND CONVERT(VARCHAR(8), APPT_TIME, 108) = '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "' AND " _
        & "PROVIDER_ID = " & udtRepeatAppt.ProviderID & " AND PATIENT_TYPE = '" & udtTelephone.CallType & "'"
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        If Not .EOF Then
            udtRepeatAppt.Repeat_Appt_ID = (!Appt_ID)
        End If
    End With
    udtRepeatAppt.ApptType = udtTelephone.CallType
    If Trim(udtRepeatAppt.Duration) = vbNullString Then
        strTerms = Calc_Duration(udtRepeatAppt.PurposeVisit)
    End If
     Select Case IsProperCase(udtTelephone.CallType)
        Case "Patient"
            strPatientType = "Established Patient"
        Case "Vendor", "Vendors"
            strPatientType = "Business"
     End Select
    If udtRepeatAppt.Repeat_Appt_ID = 0 Then
         strSQL = "INSERT INTO APPOINTMENTS (PATIENT_ID, FACILITY_ID, PROVIDER_ID, " _
            & "RESOURCE_ID, PURPOSE_VISIT, TERMINATION_DATE, APPT_DATE, " _
            & "APPT_TIME, NOTES, APPT_TYPE, PATIENT_TYPE, DURATION, REPEAT_APPOINT, VISIT_CODE, REPEAT_ID) " _
            & "VALUES ( " & Val(udtSchedPatient.IntId) & ", " & udtRepeatAppt.FacilityID & "," & udtRepeatAppt.ProviderID & "," & udtRepeatAppt.ResourseID & ",'" _
            & udtRepeatAppt.PurposeVisit & "','" & Format(udtRepeatAppt.EndDate, "mm/dd/yyyy") & "','" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "','" _
            & Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM") & "','" & udtRepeatAppt.Comment & "','" & udtRepeatAppt.ApptType & "','" _
            & strPatientType & "','" & udtRepeatAppt.Duration & "','" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "','" & udtRepeatAppt.VisitCode & "', " & udtRepeatAppt.RepeatID & ")"
        Set rsCodes = cnCodes.Execute(strSQL)
        strSQL = "SELECT * FROM APPOINTMENTS where PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
            & "FACILITY_ID = " & udtRepeatAppt.FacilityID & " AND " _
            & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "' " _
            & "AND CONVERT(VARCHAR(8), APPT_TIME, 108) = '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "' AND " _
            & "PROVIDER_ID = " & udtRepeatAppt.ProviderID & " AND APPT_TYPE = '" & udtTelephone.CallType & "'"
        Set rsCodes = cnCodes.Execute(strSQL)
        With rsCodes
            If Not .EOF Then
                udtRepeatAppt.Repeat_Appt_ID = (!Appt_ID)
            End If
        End With
    Else
        strSQL = "UPDATE APPOINTMENTS SET PURPOSE_VISIT = '" & udtRepeatAppt.PurposeVisit & "', " _
            & "RESOURCE_ID = " & udtRepeatAppt.ResourseID & " , PROVIDER_ID = " & udtRepeatAppt.ProviderID & ", " _
            & "APPT_DATE = '" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "' , " _
            & "APPT_TIME = '" & Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM") & "', NOTES = '" & udtRepeatAppt.Comment & "', " _
            & "APPT_TYPE = '" & udtTelephone.CallType & "', APPT_TYPE = '" & udtRepeatAppt.ApptType & "', " _
            & "TERMINATION_DATE = '" & udtRepeatAppt.EndDate & "',  PATIENT_TYPE = '" & strPatientType & "', DURATION = '" & udtRepeatAppt.Duration & "', " _
            & "REPEAT_APPOINT = " & udtRepeatAppt.RepeatID & ",VISIT_CODE = '" & udtRepeatAppt.VisitCode & "', ACTIVE = 1 " _
            & "WHERE APPT_ID = " & udtRepeatAppt.Repeat_Appt_ID
        Set rsCodes = cnCodes.Execute(strSQL)
    End If
    If Trim(clsWriteAppt.Add_Proc_DESC) <> vbNullString Then
         strSQL = "UPDATE APPOINTMENTS SET ADD_PROC_CODE = '" & clsWriteAppt.Add_Proc_DESC & "', ADD_PROC_DESC = '" _
            & clsWriteAppt.Add_Proc_DESC & "', PATIENT_TYPE = '" & strPatientType & "', ACTIVE = 1  WHERE APPT_ID = " & udtRepeatAppt.Repeat_Appt_ID
         Set rsCodes = cnCodes.Execute(strSQL)
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Repeat_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Repeat_Appt", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Write_Repeat_Parent()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim strTerms As String

    On Error GoTo Error_Write_Repeat_Parent
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    If udtRepeatAppt.RepeatMode = vbNullString Then
        udtRepeatAppt.RepeatMode = cboRepAppt.Text
    End If
    If udtRepeatAppt.FacilityID = 0 Then udtRepeatAppt.FacilityID = clsFacility.BillFacility
    udtRepeatAppt.RepeatNumber = cboRepeatTerm
    udtRepeatAppt.StartTime = Format(udtRepeatAppt.StartTime, "hh:mm AM/PM")
    udtRepeatAppt.RepeatID = 0
    strTerms = Calc_Duration(udtRepeatAppt.PurposeVisit)
    udtRepeatAppt.Duration = Replace(Trim(udtRepeatAppt.Duration), "minutes", "")
    udtRepeatAppt.EndTime = DateAdd("n", Val(udtRepeatAppt.Duration), udtRepeatAppt.StartTime)
    udtRepeatAppt.Duration = Trim(udtRepeatAppt.Duration) & " " & "minutes"
    udtRepeatAppt.Duration = Replace(Trim(udtRepeatAppt.Duration), "  ", "")
    strSQL = "SELECT * FROM PATIENT_REPEAT_APPT where PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
        & "APPT_TYPE = '" & udtTelephone.CallType & "'  AND " _
        & "FACILITY_ID = " & udtRepeatAppt.FacilityID & " AND " _
        & "CONVERT(VARCHAR(10), START_DATE, 101) = '" & Format(udtRepeatAppt.StartDate, "mm/dd/yyyy") & "' " _
        & "AND CONVERT(VARCHAR(8), APPT_TIME, 108) = '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "' AND PROVIDER_ID = " & udtRepeatAppt.ProviderID
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        If Not .EOF Then
            udtRepeatAppt.RepeatID = (!REPEAT_ID)
        End If
    End With
    If udtRepeatAppt.RepeatID = 0 Then
        strSQL = "INSERT INTO PATIENT_REPEAT_APPT (PATIENT_ID, FACILITY_ID, PROVIDER_ID, " _
            & "RESOURCE_ID, APPT_TIME, APPT_TYPE, START_TIME, END_TIME, START_DATE, TERMINATION_DATE, " _
            & "REPEAT_FREQ, DURATION, REPEAT_MODE, REPEAT_DAY_WEEK, COMMENT, VISIT_CODE, VISIT_TYPE, ADD_PROC_CODE, ADD_PROC_DESC, ACTIVE) VALUES " _
            & "(" & Val(udtSchedPatient.IntId) & ", " & udtRepeatAppt.FacilityID & ", " & udtRepeatAppt.ProviderID & ", " _
            & udtRepeatAppt.ResourseID & ", '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "', '" _
            & udtTelephone.CallType & "','" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "','" _
            & Format(udtRepeatAppt.EndTime, "hh:mm AM/PM") & "', '" & Format(udtRepeatAppt.StartDate, "mm/dd/yyyy") & "', '" _
            & Format(udtRepeatAppt.EndDate, "mm/dd/yyyy") & "'," & Val(udtRepeatAppt.RepeatNumber) & ", '" _
            & udtRepeatAppt.Duration & "', '" & udtRepeatAppt.RepeatMode & "', '" & udtRepeatAppt.WeekDay & "', '" & udtRepeatAppt.Comment & "', '" _
            & udtRepeatAppt.VisitCode & "','" & udtRepeatAppt.PurposeVisit & "', '" & udtRepeatAppt.Add_Proc & "','" & udtRepeatAppt.Add_Proc_DESC & "', 1)"
        Set rsCodes = cnCodes.Execute(strSQL)
        strSQL = "SELECT * FROM PATIENT_REPEAT_APPT where PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
            & "APPT_TYPE = '" & udtTelephone.CallType & "'  AND " _
            & "FACILITY_ID = " & udtRepeatAppt.FacilityID & " AND " _
            & "CONVERT(VARCHAR(10), START_DATE, 101) = '" & Format(udtRepeatAppt.StartDate, "mm/dd/yyyy") & "' " _
            & "AND CONVERT(VARCHAR(8), APPT_TIME, 108) = '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "' AND PROVIDER_ID = " & udtRepeatAppt.ProviderID
        Set rsCodes = cnCodes.Execute(strSQL)
        With rsCodes
            If Not .EOF Then
                udtRepeatAppt.RepeatID = (!REPEAT_ID)
            End If
        End With
    Else
        strSQL = "UPDATE PATIENT_REPEAT_APPT  SET APPT_TIME = '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "'," _
            & "START_TIME = '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "',  " _
            & "END_TIME = '" & Format(udtRepeatAppt.EndTime, "hh:mm AM/PM") & "', " _
            & "START_DATE = '" & Format(udtRepeatAppt.StartDate, "mm/dd/yyyy") & "', " _
            & "TERMINATION_DATE = '" & Format(udtRepeatAppt.EndDate, "mm/dd/yyyy") & "', " _
            & "REPEAT_FREQ = '" & udtRepeatAppt.RepeatNumber & "', " _
            & "DURATION  =  '" & udtRepeatAppt.Duration & "', " _
            & "REPEAT_MODE =  '" & udtRepeatAppt.RepeatMode & "', " _
            & "REPEAT_DAY_WEEK = '" & udtRepeatAppt.WeekDay & "', " _
            & "COMMENT =  '" & udtRepeatAppt.Comment & "', " _
            & "VISIT_TYPE = '" & udtRepeatAppt.PurposeVisit & "', " _
            & "VISIT_CODE = '" & udtRepeatAppt.VisitCode & "'," _
            & "ADD_PROC_CODE =  '" & udtRepeatAppt.Add_Proc & "',  " _
            & "ADD_PROC_DESC  = '" & udtRepeatAppt.Add_Proc_DESC & "'  WHERE " _
            & "REPEAT_ID = " & udtRepeatAppt.RepeatID
        Set rsCodes = cnCodes.Execute(strSQL)
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Repeat_Parent:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Repeat_Parent", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Write_Repeat_Error()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngPrimKey As Long


    On Error GoTo Error_Write_Repeat_Error
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    strSQL = "SELECT CONFLICT_ID FROM REPEAT_CONFLICT WHERE REPEAT_ID = " & udtRepeatAppt.RepeatID & " AND " _
        & "PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
        & "APPT_DATE = '" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "'"
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        If Not .EOF Then
            lngPrimKey = (!CONFLICT_ID)
        End If
    End With
    If lngPrimKey = 0 Then
         strSQL = "INSERT INTO REPEAT_CONFLICT (REPEAT_ID, APPT_DATE, " _
            & "APPT_TIME, PATIENT_ID, CONFLICT_REASON, APPT_ID) " _
            & "VALUES ( " & udtRepeatAppt.RepeatID & ", '" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "','" _
            & Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM") & "'," _
            & Val(udtSchedPatient.IntId) & ", '" & udtRepeatAppt.ConflictReason & "'," _
            & udtRepeatAppt.Repeat_Appt_ID & ")"
        Set rsCodes = cnCodes.Execute(strSQL)
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Repeat_Error:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Repeat_Error", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cboDayWeek_Click()
    udtRepeatAppt.WeekDay = cboDayWeek.Text
End Sub
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

Private Sub cboLocation_Click()
Dim Idx%

    If mbooByPass Then mbooByPass = False: Exit Sub
    For Idx = 0 To cboLocation.listcount - 1
        If cboLocation.List(Idx) = cboLocation.Text Then
            udtRepeatAppt.FacilityID = cboLocation.ItemData(Idx)
            Get_Appt_Providers
            Exit For
        End If
    Next Idx

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


    On Error GoTo Error_Load_Facilities
    Screen.MousePointer = vbHourglass
    cboLocation.Clear
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "Get_Appt_Facilities"
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
                                If lngFacID = udtRepeatAppt.FacilityID Then
                                    booDefault = True
                                End If
                        End Select
                    End If
                End If
            Next
            .MoveNext
            strName = Fix_Name(strName)
            strName = Fix_Facility_Name(strName)
            If Trim(strName) <> vbNullString Then
                For Idx = 0 To cboLocation.listcount - 1
                    If TextFound(strName, cboLocation.List(Idx), Len(strName), True) Then
                        booSkip = True
                        Exit For
                    End If
                Next Idx
                If Not booSkip Then
                    cboLocation.AddItem strName
                    If lngFacID > 0 Then cboLocation.ItemData(cboLocation.NewIndex) = lngFacID
                End If
            End If
            If udtRepeatAppt.FacilityID <= 0 And lngFacID > 0 Then
                udtRepeatAppt.FacilityID = lngFacID
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


Private Sub cboRepAppt_Change()
    If cboRepAppt.Text = "Day" Then
       cboDayWeek.ListIndex = -1
       cboDayWeek.Enabled = False
       ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm) - 1, ccdStart.Value)
       ccdTermDate.Enabled = False
    Else
       If cboDayWeek.Enabled = False Then cboDayWeek.Enabled = True
       If ccdTermDate.Enabled = False Then ccdTermDate.Enabled = True
    End If
End Sub


Private Sub cboRepAppt_Click()
    If Trim(cboRepAppt.Text) = "Day" Then
       cboDayWeek.ListIndex = -1
       cboDayWeek.Enabled = False
       If DateDiff("d", ccdStart.Value, ccdTermDate.Value) <= 0 Then
        ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm), ccdStart.Value)
       End If
       'ccdTermDate.Enabled = False
    Else
       If cboDayWeek.Enabled = False Then cboDayWeek.Enabled = True
       If ccdTermDate.Enabled = False Then ccdTermDate.Enabled = True
    End If
End Sub


Private Sub cboRepeatTerm_Click()
    If cboRepAppt.Text = "Day" Then
       ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm) - 1, ccdStart.Value)
    End If
End Sub


Private Sub cboResProvider_Click()
    Dim Idx%
    If Trim(cboResProvider.Text) <> vbNullString Then
        For Idx = 0 To cboResProvider.listcount - 1
            If cboResProvider.List(Idx) = cboResProvider.Text Then
                udtRepeatAppt.ProviderID = cboResProvider.ItemData(Idx)
                Format_PurposeGrid
                Load_Visit
                Format_PsychoTherapy
                Load_Codes_Psychotherapy
                Exit For
            End If
        Next Idx
    End If
End Sub


Private Sub cboStartTime_Click()
    If cboStartTime.Text <> vbNullString Then
        udtRepeatAppt.StartTime = Format(cboStartTime.Text, "hh:mm AM/PM")
    End If
End Sub


Private Sub ccdStart_Change()
    udtRepeatAppt.StartDate = Format(ccdStart.Value, "mm/dd/yyyy")
    If cboRepAppt.Text = "Day" Then
       ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm) - 1, ccdStart.Value)
    End If
    Exit Sub
End Sub

Private Sub ccdStart_Click()
    udtRepeatAppt.StartDate = Format(ccdStart.Value, "mm/dd/yyyy")
    If cboRepAppt.Text = "Day" Then
       ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm) - 1, ccdStart.Value)
    End If
    Exit Sub
End Sub





Private Sub cctComments_KeyUp(KeyCode As Integer, Shift As Integer)
    udtRepeatAppt.Comment = cctComments.Text
End Sub


Private Sub cmdAccept_Click()
Dim strDate As String
Dim varCellValue
Dim lngRow As Long
Dim lngCol As Long

    With grdMultiAppt
        clsGridFormat.CellTag = vbNullString
        .Col = mlngGridCol
        .Row = mlngGridRow
        If .CellTag <> vbNullString Then clsGridFormat.CellTag = .CellTag
        If Trim(clsGridFormat.CellTag) <> vbNullString Then
            Select Case .CellTag
                Case "APPT_DATE"
                    strDate = Format(ccdDate.Value, "mm/dd/yyyy")
                    Format_CellTypeEdit grdMultiAppt, mlngGridRow, mlngGridCol, "1", strDate, "TypeHAlignCenter", clsGridFormat.CellTag
                    grdMultiAppt.SetCellBorder mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "APPT_TIME"
                    If cboRepTime.ListIndex <> -1 Then
                       strDate = cboRepTime.Text
                       Format_CellTypeEdit grdMultiAppt, mlngGridRow, mlngGridCol, "1", strDate, "TypeHAlignCenter", clsGridFormat.CellTag
                       grdMultiAppt.SetCellBorder mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
            End Select
        End If
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = mlngGridRow
            .Row2 = mlngGridRow
            .BackColor = grdCurItem
        .BlockMode = False
    End With
    fraNumber.Visible = False
    Exit Sub
End Sub

Private Function LookUpCellTag(vstrName As String, Optional vlngRow As Long) As String
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String
    
    lngCol = 0
    With grdMultiAppt
        LookUpCellTag = vbNullString
        For lngCol = 1 To .MaxCols
            .Col = lngCol
            .Row = SpreadHeader + 2
            If .CellTag <> vbNullString Then strTag = .CellTag
            If Trim(strTag) <> vbNullString And TextFound(vstrName, Trim(strTag), Len(vstrName), True) Then
                LookUpCellTag = lngCol & "|" & lngRow
                Exit For
            End If
        Next lngCol
    End With
    Exit Function
End Function
Private Function LookUpCellTag_RepeatHist(vstrName As String, Optional vlngRow As Long) As String
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String
    
    lngCol = 0
    With grdRepeatHist
        LookUpCellTag_RepeatHist = vbNullString
        For lngCol = 1 To .MaxCols
            .Col = lngCol
            .Row = SpreadHeader + 1
            If .CellTag <> vbNullString Then strTag = .CellTag
            If Trim(strTag) <> vbNullString And TextFound(vstrName, Trim(strTag), Len(vstrName), True) Then
                LookUpCellTag_RepeatHist = lngCol & "|" & lngRow
                Exit For
            End If
        Next lngCol
    End With
    Exit Function
End Function

Private Function LookUpCellTag_Header(vstrName As String, Optional vlngRow As Long) As String
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String

    lngCol = 0
    With grdMultiAppt
        LookUpCellTag_Header = vbNullString
            For lngCol = 1 To .MaxCols
                    .Col = lngCol
                    .Row = SpreadHeader + 1
                    If .CellTag <> vbNullString Then
                        If TextFound(vstrName, .CellTag, Len(vstrName), True) Then
                            LookUpCellTag_Header = lngCol & "|" & lngRow
                            Exit For
                        End If
                    End If
            Next lngCol
    End With
    Exit Function
End Function

Private Sub cmdConflict_Click()
Dim lngRow As Long

    udtRepeatAppt.RepeatID = 0
    With grdRepeatHist
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                udtRepeatAppt.RepeatID = .GetRowItemData(lngRow)
                Exit For
            End If
        Next lngRow
    End With
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please select patient first!", vbInformation, "PatientTrac Repeat Appointments"
        Exit Sub
    End If
    If udtRepeatAppt.RepeatID = 0 Then
        MsgBox "Please select repeat appointments row first!", vbInformation, "PatientTrac Repeat Appointments"
        Exit Sub
    End If
    lblPatName.Visible = True
    lblPatName.Caption = udtSchedPatient.Name
    Format_Conflict
    Get_SchedConflict
    fraSchedConflicts.Visible = True
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim lngRow As Long
Dim lngRepId As Long
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim lngDateDiff As Long
Dim strTermDate As String
Dim strApptDate As String
Dim varCellValue
Dim booFound As Boolean
Dim Arr
Dim lngCol As Long

    On Error GoTo Error_cmdRemoveRow
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please Select Patient First!", vbInformation, "PatientTrac Repeat Appointments"
        Exit Sub
    Else
        lblPatName.Visible = True
        lblPatName.Caption = udtSchedPatient.Name
    End If
    With grdRepeatHist
        For lngRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                If .BackColor = grdCurItem Then
                    booFound = True
                    lngRepId = .GetRowItemData(lngRow)
                    Exit For
                End If
            .BlockMode = False
        Next lngRow
    End With
    If lngRepId > 0 Then
        If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
        strSQL = "DELETE PATIENT_REPEAT_APPT  WHERE REPEAT_ID = " & lngRepId
          Set rsAppt = cnAppt.Execute(strSQL)
        strSQL = "DELETE APPOINTMENTS  WHERE REPEAT_ID = " & lngRepId
          Set rsAppt = cnAppt.Execute(strSQL)
        strSQL = "DELETE REPEAT_CONFLICT  WHERE REPEAT_ID = " & lngRepId
          Set rsAppt = cnAppt.Execute(strSQL)
    End If
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Format_MultiAppt
    Format_Repeat_History
    Load_Repeat_Appt
    Screen.MousePointer = vbDefault
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdRemoveRow:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdRemoveRow", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdExit_Click()
    fraControl.Visible = False
    Clear_Fields
    Format_Repeat_History
    Load_Repeat_Appt
    
    Exit Sub
End Sub

Private Sub cmdExitConflict_Click()
    fraSchedConflicts.Visible = False
End Sub
Private Sub cmdMultiSave_Click()
Dim Arr
Dim lngCol As Long
Dim varCellValue

    With grdMultiAppt
        Arr = Split(LookUpCellTag("APPT_DATE", mlngGridRow))
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
        End If
        If lngCol > 0 Then .GetText lngCol, mlngGridRow, varCellValue
        If Trim(varCellValue) <> vbNullString And lngCol > 0 Then
             udtRepeatAppt.StartDate = Format(varCellValue, "mm/dd/yyyy")
        End If
        varCellValue = vbNullString
        lngCol = 0
        Arr = Split(LookUpCellTag("APPT_TIME", mlngGridRow))
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
        End If
        If lngCol > 0 Then .GetText lngCol, mlngGridRow, varCellValue
        If Trim(varCellValue) <> vbNullString And lngCol > 0 Then
             udtRepeatAppt.StartTime = varCellValue
        End If
        lngCol = 0
        Arr = Split(LookUpCellTag("REPEAT_DAY_WEEK", mlngGridRow))
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
        End If
        If lngCol > 0 Then .GetText lngCol, mlngGridRow, varCellValue
        If Trim(varCellValue) <> vbNullString And lngCol > 0 Then
             udtRepeatAppt.WeekDay = varCellValue
        End If
        If Trim(udtRepeatAppt.StartDate) <> vbNullString And Trim(udtRepeatAppt.StartTime) <> vbNullString Then
            Save_RepeatData udtRepeatAppt.StartTime, udtRepeatAppt.StartDate
        End If
        .BlockMode = True
           .Col = 1
           .Col2 = 5
           .Row = mlngGridRow
           .Row2 = mlngGridRow
           .BackColor = NormalColor
        .BlockMode = False
    End With
    Exit Sub
End Sub

Private Sub cmdNewRepeat_Click()
    
    mstrType = "New"
    cmdSched.Visible = True
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please Select Patient First!", vbInformation, "PatientTrac Repeat Appointments"
       Exit Sub
    End If
    Clear_Fields
    Load_Form
    Load_ComboTime
    If cboDuration.listcount <= 0 Then Load_Duration
    grdProcedure.MaxRows = 0
    If grdProcedure.MaxRows = 0 Then
         Load_Visit LCase(udtTelephone.CallType)
    End If
    fraControl.Visible = True
    Exit Sub
End Sub
Function Load_Duration()
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
Private Sub cmdPatient_Click()
Dim i%

    On Error GoTo Error_Select_Patient
    Screen.MousePointer = vbHourglass
    udtSchedPatient.IntId = vbNullString
    Clear_SchedPatient
    grdMultiAppt.ClearRange 1, 1, grdMultiAppt.MaxRows, grdMultiAppt.ColsFrozen, False
    grdMultiAppt.MaxRows = 0
    frmListPatient.mlngListItemData = 0
    frmListPatient.Show vbModal
    If frmListPatient.mlngListItemData <> 0 Then
        If udtTelephone.CallType = vbNullString Then udtTelephone.CallType = "Patient"
        udtSchedPatient.IntId = Val(frmListPatient.mlngListItemData)
        Select Case IsProperCase(udtTelephone.CallType)
            Case "Patient"
                If Trim(udtSchedPatient.IntId) <> vbNullString Then
                    Get_SchedPatient Val(udtSchedPatient.IntId)
                End If
            Case "Vendor", "Business", "Vendors"
                Get_SchedVendor Val(udtSchedPatient.IntId), udtAppointment.SchedTime
            Case "Legal", "Attorney"
                Get_Schedule_Attorney Val(udtSchedPatient.IntId)
            Case "Physician"
                Get_SchedPhysician Val(udtSchedPatient.IntId)
        End Select
        Clear_RepeatAppt
        Format_MultiAppt
        Format_Repeat_History
        Load_Repeat_Appt
        If udtRepeatAppt.RepeatID > 0 Then
            Read_MultiAppt Val(udtSchedPatient.IntId)
        End If
        grdMultiAppt.MaxRows = 0
        If udtSchedPatient.Name = vbNullString Then udtSchedPatient.Name = udtSchedPatient.LastName
        lblPatName.Caption = udtSchedPatient.Name & vbNewLine & udtSchedPatient.IntId
        lblPatName.Visible = True
        Format_PurposeGrid
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Select_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, udtSchedPatient.Name, "Select_Patient", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Function Procedure_Grid_Action()
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim Idx%
Dim Arr
Dim strDuration As String
Dim lngTestRow As Long

    If cboDuration.listcount <= 0 Then Load_Duration
    With grdProcedure
        lngTestRow = .ActiveRow
        If IsProperCase(udtTelephone.CallType) = "Patient" Or udtTelephone.CallType = vbNullString Then
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
                .Row = lngTestRow
                .Row2 = lngTestRow
                .BackColor = grdCurItem
            .BlockMode = False
            .GetText 1, lngTestRow, varCellValue
            udtRepeatAppt.PurposeVisit = varCellValue
            LookUpCPTCode udtRepeatAppt.PurposeVisit
            udtRepeatAppt.VisitCode = Left(udtEncounter.Intrvn_CD, 5)
            udtRepeatAppt.Duration = vbNullString
            Calc_Duration udtRepeatAppt.PurposeVisit
'            If cboDuration <> udtEncounter.VisitTime Then
'                udtRepeatAppt.Duration = Trim(cboDuration)
'            End If
            If udtRepeatAppt.Duration = vbNullString Then
                With grdProcedure
                    .GetText 3, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        udtSchedPatient.Duration = varCellValue
                         Arr = Split(udtSchedPatient.Duration, " ")
                        If UBound(Arr) > 0 Then
                            For Idx = LBound(Arr) To UBound(Arr)
                                If IsNumeric(Arr(Idx)) Then
                                    udtRepeatAppt.Duration = Trim(Arr(Idx))
                                    Exit For
                                End If
                            Next Idx
                            udtRepeatAppt.Duration = udtRepeatAppt.Duration & " " & "minutes"
                        End If
                    End If
                End With
            End If
            'udtRepeatAppt.VisitCode = udtEncounter.Appt_Type
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
                        udtRepeatAppt.Duration = vbNullString
                        udtRepeatAppt.Duration = Trim(strDuration)
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
            udtRepeatAppt.PurposeVisit = udtEncounter.Appt_Type
            .GetText 2, lngRow, varCellValue
            udtRepeatAppt.VisitCode = varCellValue
            .GetText 3, lngRow, varCellValue
            udtRepeatAppt.Duration = varCellValue
            If strDuration <> udtRepeatAppt.Duration And strDuration <> vbNullString Then
                udtRepeatAppt.Duration = Replace(LCase(strDuration), "minutes", "")
            End If
        End If
    End With
    For Idx = 0 To cboDuration.listcount - 1
        If IsProperCase(Trim(cboDuration.List(Idx))) = IsProperCase(Trim(udtRepeatAppt.Duration)) Then
            cboDuration.Text = cboDuration.List(Idx)
            Exit For
        End If
    Next Idx
    Screen.MousePointer = vbDefault
    Exit Function
End Function
Private Sub cmdPrint_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(1) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    If grdMultiAppt.MaxRows = 0 Then
        MsgBox "No records to print!", vbInformation, "PatientTrac Multiple Appointments"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Multi Appointment Schedule"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    With grdMultiAppt
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Repeat Appointments"
        .PrintColor = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintSmartPrint = True
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Multiple Appointments"
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

Private Sub cmdPrintList_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(1) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    If grdTerms.MaxRows = 0 Then
        MsgBox "No records to print!", vbInformation, "PatientTrac Multiple Appointments"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Multi Appointment Schedule"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    With grdTerms
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Multiple Appointment"
        .PrintColor = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintSmartPrint = True
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Conflict Report"
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

Private Sub cmdProvTypes_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Provider Types"
    frmNurseTemplates.fType = "Provider Types"
    frmNurseTemplates.fSubDomain = "sd_Provider_ProviderType"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    Exit Sub
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
    udtRepeatAppt.Duration = strTime
    fraPsycho.Visible = False
    Exit Sub

End Sub

Private Sub cmdPsyClose_Click()
    fraPsycho.Visible = False
    Exit Sub
End Sub

Private Sub cmdRemAddOn_Click()
'Dim Arr
'Dim cnApptGrid As New ADODB.Connection
'Dim rsApptGrid As New ADODB.Recordset
'Dim lngProviderID As Long
'Dim lngFacilityID As Long
'Dim strSQL As String
'Dim strTime As String
'Dim strDate As String
'Dim lngApptID As Long
'Dim lngRow As Long
'Dim varCellValue
'Dim strAdd_Proc As String
'Dim strAdd_Proc_Code As String
'Dim strPurpose As String
'Dim lngDateDiff As Long
'Dim strPatientType As String
'Dim Idx%
'
'    On Error GoTo Error_cmdRemAddOn
'    Screen.MousePointer = vbHourglass
'    If Trim(udtSchedPatient.IntId) <> vbNullString Then
'        If Trim(cboLocation.Text) = vbNullString Then
'            MsgBox "Please reselect a facility!", vbInformation, "PatientTrac Appointment Scheduling"
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        ElseIf cboLocation.ListIndex >= 0 Then
'            lngFacilityID = cboLocation.ItemData(cboLocation.ListIndex)
'        End If
'        If udtRepeatAppt.ProviderID > 0 Then
'            lngProviderID = udtRepeatAppt.ProviderID
'        ElseIf Trim(cboResProvider.Text) <> vbNullString Then
'            lngProviderID = cboResProvider.ItemData(cboResProvider.ListIndex)
'        End If
'        strTime = Trim(ccdTime.Value)
'        strTime = Format(strTime, "hh:mm AM/PM")
'        If Trim(cboPatientType.Text) <> vbNullString Then strPatientType = Trim(cboPatientType.Text)
'        With grdPsychotherapy
'            For lngRow = 1 To .MaxRows
'                .BlockMode = True
'                    .Col = 1
'                    .Col2 = .MaxCols
'                    .Row = lngRow
'                    .Row2 = lngRow
'                    .BackColor = NormalColor
'                .BlockMode = False
'            Next lngRow
'        End With
'        lngRow = 0
'        For lngRow = 1 To grdProcedure.MaxRows
'            With grdProcedure
'                .Col = 1
'                .Row = lngRow
'                If .BackColor = grdCurItem Then
'                    .GetText 1, lngRow, varCellValue
'                    If Trim(varCellValue) <> vbNullString Then
'                        strPurpose = Trim(varCellValue)
'                    End If
'                    .GetText 3, lngRow, varCellValue
'                    If Trim(varCellValue) <> vbNullString Then
'                        strAdd_Proc = Trim(varCellValue)
'                    End If
'                    .GetText 4, lngRow, varCellValue
'                    If Trim(varCellValue) <> vbNullString Then
'                        strAdd_Proc_Code = Trim(varCellValue)
'                    End If
'                    Exit For
'                End If
'            End With
'        Next lngRow
'        If lngRow > 0 Then
'            Format_CellTypeEdit grdProcedure, lngRow, 3, "1", "", "TypeHAlignCenter", "ADD_DESC"
'            Format_CellTypeEdit grdProcedure, lngRow, 4, "1", "", "TypeHAlignCenter", "ADD_CPT_CODE"
'        End If
'        If cnApptGrid.State = adStateClosed Then cnApptGrid.Open psConnect(0)
'        strSQL = "SELECT APPT_ID FROM APPOINTMENTS WHERE PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
'            & "PROVIDER_ID = " & lngProviderID & " AND FACILITY_ID = " & lngFacilityID & " AND " _
'            & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "' AND " _
'            & "CONVERT(VARCHAR(08), APPT_TIME, 108)  = '" & Format(udtAppointment.SchedTime, "hh:mm AM/PM") & "' AND APPT_TYPE = '" & udtTelephone.CallType & "'"
'        Set rsApptGrid = cnApptGrid.Execute(strSQL)
'        With rsApptGrid
'            If Not .EOF Then
'                If Not IsNull(!Appt_ID) Then
'                    lngApptID = (!Appt_ID)
'                End If
'            End If
'        End With
'        If cnApptGrid.State = adStateClosed Then cnApptGrid.Open psConnect(0)
'        If lngApptID > 0 Then
'             strSQL = "UPDATE APPOINTMENTS SET ADD_PROC_CODE = '', ADD_PROC_DESC = '' WHERE APPT_ID = " & lngApptID
'            Set rsApptGrid = cnApptGrid.Execute(strSQL)
'        End If
'        Calc_Duration strPurpose
'        If cboDuration.ListCount <= 0 Then Load_Duration
'        If Trim(udtEncounter.VisitTime) <> vbNullString Then
'            For Idx = 0 To cboDuration.ListCount - 1
'                If cboDuration.List(Idx) = udtEncounter.VisitTime Then
'                    cboDuration.Text = cboDuration.List(Idx)
'                    Exit For
'                End If
'            Next Idx
'        End If
'    End If
'    If cnApptGrid.State = adStateOpen Then
'        cnApptGrid.Close
'        Set cnApptGrid = Nothing
'        Set rsApptGrid = Nothing
'    End If
'    Screen.MousePointer = vbDefault
'    Exit Sub
'Error_cmdRemAddOn:
'    Screen.MousePointer = vbDefault
'    Write_Error Err.Description, Err.Number, strSQL, "cmdRemAddOn", App.EXEName & "." & Me.Name
'    Exit Sub
'    Resume Next
End Sub


Private Sub cmdRemoveRow_Click()
Dim lngRow As Long
Dim lngRepId As Long
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim lngDateDiff As Long
Dim strTermDate As String
Dim strApptDate As String
Dim varCellValue
Dim booFound As Boolean
Dim Arr
Dim lngCol As Long

    On Error GoTo Error_cmdRemoveRow
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please Select Patient First!", vbInformation, "PatientTrac Repeat Appointments"
        Exit Sub
    Else
        lblPatName.Visible = True
        lblPatName.Caption = udtSchedPatient.Name
    End If
    With grdMultiAppt
        For lngRow = 1 To grdMultiAppt.MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                If .BackColor = grdCurItem Then
                    booFound = True
                    lngRepId = .GetRowItemData(lngRow)
                    .GetText 2, lngRow, varCellValue
                    If IsDate(varCellValue) Then
                        strTermDate = Format(varCellValue, "mm/dd/yyyy")
                    End If
                    Exit For
                End If
            .BlockMode = False
        Next lngRow
    End With
'    If booFound Then
'        lngCol = 0
'        Arr = Split(LookUpCellTag("APPT_DATE"), "|")
'        If UBound(Arr) >= 0 Then
'            lngCol = Arr(0)
'            If lngCol > 0 Then
'            End If
'        End If
'    Else
    If strTermDate = vbNullString Then
        MsgBox "Please Select Terminate Date First!", vbInformation, "PatientTrac Repeat Appointments"
        Exit Sub
    End If
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please Select Patient First!", vbInformation, "PatientTrac Repeat Appointments"
        Exit Sub
    Else
        lblPatName.Visible = True
        lblPatName.Caption = udtSchedPatient.Name
    End If
    Screen.MousePointer = vbHourglass
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    If Not IsDate(strTermDate) Then Exit Sub
    If lngRepId > 0 Then
        strSQL = "DELETE APPOINTMENTS WHERE APPT_ID = " & lngRepId
            Set rsAppt = cnAppt.Execute(strSQL)
            grdMultiAppt.DeleteRows lngRow, 1
            grdMultiAppt.MaxRows = grdMultiAppt.MaxRows - 1
            MsgBox "Selected Repeat Appt has been successfully removed!", vbInformation, "PatientTrac Scheduling"
    End If
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdRemoveRow:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdRemoveRow", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Sub cmdResource_Click()
    frmResource.Show vbModal
    Exit Sub
End Sub

Private Sub cmdSched_Click()
Dim Idx%
Dim lngRow As Long
Dim lngCurRow As Long



    On Error GoTo Error_cmdSched
    Screen.MousePointer = vbHourglass
    Clear_RepeatAppt
    If cboLocation.Text = vbNullString Then
        MsgBox "Appointment requires a facility! Please select a facility.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        For Idx = 0 To cboLocation.listcount - 1
            If cboLocation.List(Idx) = cboLocation.Text Then
                udtRepeatAppt.FacilityID = cboLocation.ItemData(Idx)
                Exit For
            End If
        Next Idx
    End If
    If cboResProvider.Text = vbNullString Then
        MsgBox "Appointment requires a provider! Please select a provider.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        For Idx = 0 To cboResProvider.listcount - 1
            If cboResProvider.List(Idx) = cboResProvider.Text Then
                udtRepeatAppt.ProviderID = cboResProvider.ItemData(Idx)
                Exit For
            End If
        Next Idx
    End If
    Get_Repeat_Data
    If udtRepeatAppt.PurposeVisit = vbNullString Then
        MsgBox "Please Select Appt Type.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If cboRepeatTerm.ListIndex = -1 Then
       MsgBox "Please Select Occurance Interval (Every).", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If cboRepAppt.ListIndex = -1 Then
       MsgBox "Please Select Term Value.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If cboDayWeek.ListIndex = -1 And (cboRepAppt.Text <> "Day" And cboRepAppt.Text <> "Month") Then
        MsgBox "Please Select Day Of Week.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If cboStartTime.Text = vbNullString Then
        MsgBox "Please Confirm Appt Time.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If IsNull(ccdTermDate.Value) Or ccdTermDate.Value = vbNullString Then
        MsgBox "Please Confirm End Date.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If IsNull(ccdStart.Value) Then
       MsgBox "Please enter Start Date.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If IsNull(ccdTermDate.Value) Then
        MsgBox "Please enter End Date.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If DateDiff("d", Now, ccdStart.Value) <= 0 Then
        MsgBox "Start Date Must Be Greater Than Current Date.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If DateDiff("d", ccdStart.Value, ccdTermDate.Value) < 0 Then
        MsgBox "End Date Must Be Greater Than Or Equal To Start Date.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Trim(cboResProvider.Text) <> vbNullString Then
        For Idx = 0 To cboResProvider.listcount - 1
            If cboResProvider.List(Idx) = cboResProvider.Text Then
                udtRepeatAppt.ProviderID = cboResProvider.ItemData(Idx)
                Exit For
            End If
        Next Idx
    End If
    grdMultiAppt.ClearRange 1, 1, grdMultiAppt.MaxCols, grdMultiAppt.MaxRows, False
    grdMultiAppt.MaxRows = 0
    If cboStartTime.Text <> vbNullString Then
        udtRepeatAppt.StartTime = Format(cboStartTime.Text, "hh:mm AM/PM")
    End If
    If udtRepeatAppt.StartTime <> vbNullString And mstrType = "New" Then
        Repeat_Appt Format(udtRepeatAppt.StartTime, "hh:mm AM/PM")
    ElseIf udtRepeatAppt.StartTime <> vbNullString And mstrType = "Update" Then
        Update_Repeat_Appt
    End If
    Load_Repeat_Appt
    With grdRepeatHist
        For lngRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                .BackColor = NormalColor
            .BlockMode = False
            If .GetRowItemData(lngRow) = udtRepeatAppt.RepeatID Then
                lngCurRow = lngRow
            End If
        Next lngRow
        .BlockMode = True
           .Col = 1
           .Col2 = .MaxCols
           .Row = lngCurRow
           .Row2 = lngCurRow
           .BackColor = grdCurItem
           cmdDelete.Visible = True
        .BlockMode = False
    End With
    fraControl.Visible = False
    Clear_Fields
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdSched:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdSched", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Function Repeat_Appt(vstrTime As String)
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
Dim intMonth As Double
Dim intDay As Double
Dim intYear As Double
Dim lngRepeat As Long
Dim bError As Boolean

    On Error GoTo Error_Repeat_Appt
    Screen.MousePointer = vbHourglass
    cctBlockConflict.Text = vbNullString
    If Val(udtSchedPatient.IntId) >= 0 Then
        lngPatientID = Val(udtSchedPatient.IntId)
    Else
        MsgBox "Please select a patient first!", vbInformation, "PatientTrac Repeat Appointments Schedule"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If Not IsNull(ccdTermDate.Value) Then
        udtRepeatAppt.EndDate = Format(ccdTermDate.Value, "mm/dd/yyyy")
    End If
    If Not IsNull(ccdStart.Value) Then
        udtRepeatAppt.StartDate = Format(ccdStart.Value, "mm/dd/yyyy")
    End If
    udtRepeatAppt.WeekDay = cboDayWeek
    Select Case cboRepAppt
        Case "Day"
            nDays = Val(cboRepeatTerm.Text)
        Case "Week"
            nDays = Val(cboRepeatTerm.Text) * 7
        Case "Month"
            nDays = Val(cboRepeatTerm.Text) * 30
        Case "Bi-Month"
            nDays = Val(cboRepeatTerm.Text) * 15
        Case "Annual"
            nDays = Val(cboRepeatTerm.Text) * 365
    End Select
    If Trim(cctComments.Text) <> vbNullString Then
        udtRepeatAppt.Comment = Trim(cctComments.Text)
    End If
'    If Trim(cboVisitType.Text) <> vbNullString Then
'        udtRepeatAppt.PurposeVisit = cboVisitType.Text
'    End If
    udtRepeatAppt.ApptTime = Format(vstrTime, "hh:mm AM/PM")
    Select Case IsProperCase(udtTelephone.CallType)
        Case "Patient"
            udtRepeatAppt.ApptType = "Patient"
        Case "Vendor", "Business"
            udtRepeatAppt.ApptType = "Business"
            udtTelephone.CallType = "Business"
        Case "Legal", "Attorney"
            udtRepeatAppt.ApptType = "Legal"
            udtTelephone.CallType = "Legal"
        Case "Physician"
            udtRepeatAppt.ApptType = "Physician"
            udtTelephone.CallType = "Physician"
    End Select
    udtRepeatAppt.Duration = Calc_Duration(udtRepeatAppt.PurposeVisit)
    Write_Repeat_Parent
    lngRepeat = udtRepeatAppt.RepeatID
    If nDays >= 1 Then
        Select Case cboRepAppt.Text
            Case "Day"
                udtRepeatAppt.ApptDate = udtRepeatAppt.StartDate
                lngDateDiff = (DateDiff("d", udtRepeatAppt.StartDate, udtRepeatAppt.EndDate)) + 1
                If lngDateDiff >= 0 Then
                    For Idx = 0 To lngDateDiff
                        If udtRepeatAppt.ApptDate = vbNullString Then udtRepeatAppt.ApptDate = udtRepeatAppt.StartDate
                        udtRepeatAppt.ApptDate = DateAdd("d", 1, udtRepeatAppt.ApptDate)
                        If Trim(udtRepeatAppt.ApptDate) <> vbNullString Then
                            lngEndDate = DateDiff("d", udtRepeatAppt.ApptDate, udtRepeatAppt.EndDate)
                            If lngEndDate >= 0 Then
                                udtRepeatAppt.RepeatID = lngRepeat
                                If Get_SchedConflict_Old(udtRepeatAppt.ApptDate, udtRepeatAppt.ApptTime) Then
                                    Write_Repeat_Error
                                    bError = True
                                Else
                                    Write_Repeat_Appt
                                End If
                            End If
                        End If
                        
                    Next Idx
                End If
            Case "Week"
                udtRepeatAppt.ApptDate = vbNullString
                lngDateDiff = DateDiff("ww", udtRepeatAppt.StartDate, udtRepeatAppt.EndDate)
                lngDateDiff = lngDateDiff / Val(cboRepeatTerm.Text)
                If lngDateDiff > 0 Then
                    For Idx = 0 To lngDateDiff
                        If udtRepeatAppt.StartDate <> vbNullString Then
                            If udtRepeatAppt.ApptDate = vbNullString Then udtRepeatAppt.ApptDate = udtRepeatAppt.StartDate
                            'If udtRepeatAppt.Weekday <> vbNullString Then
                                Select Case udtRepeatAppt.WeekDay
                                    Case "Monday"
                                        FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbMonday)
                                        udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                    Case "Tuesday"
                                        FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbTuesday)
                                        udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                    Case "Wednesday"
                                        FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbWednesday)
                                        udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                    Case "Thursday"
                                        FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbThursday)
                                        udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                    Case "Friday"
                                        FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbFriday)
                                        udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                    Case "Saturday"
                                        FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbSaturday)
                                        udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                    Case "Sunday"
                                        FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbSunday)
                                        udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                End Select
                            If DateDiff("d", udtRepeatAppt.StartDate, udtRepeatAppt.ApptDate) >= 0 Then
                                If Trim(udtRepeatAppt.ApptDate) <> vbNullString Then
                                    udtRepeatAppt.RepeatID = lngRepeat
                                    lngEndDate = DateDiff("d", udtRepeatAppt.ApptDate, udtRepeatAppt.EndDate)
                                    If lngEndDate >= 0 Then
                                        If Get_SchedConflict_Old(udtRepeatAppt.ApptDate, udtRepeatAppt.ApptTime) Then
                                            Write_Repeat_Error
                                            bError = True
                                        Else
                                            Write_Repeat_Appt
                                        End If
                                    End If
                                End If
                            End If
                            udtRepeatAppt.ApptDate = DateAdd("d", nDays, udtRepeatAppt.ApptDate)
                        End If
                    Next Idx
                End If
            Case "Month"
                udtRepeatAppt.ApptDate = vbNullString
                lngDateDiff = DateDiff("m", udtAppointment.SchedDate, udtRepeatAppt.EndDate)
                If lngDateDiff > 0 Then
                    For Idx = 0 To lngDateDiff
                        If udtRepeatAppt.ApptDate = vbNullString Then udtRepeatAppt.ApptDate = udtRepeatAppt.StartDate
                        If udtRepeatAppt.WeekDay <> vbNullString Then
                            Select Case udtRepeatAppt.WeekDay
                                Case "Monday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbMonday)
                                    udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Tuesday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbTuesday)
                                    udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Wednesday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbWednesday)
                                    udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Thursday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbThursday)
                                    udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Friday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbFriday)
                                    udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Saturday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbSaturday)
                                    udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Sunday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbSunday)
                                    udtRepeatAppt.ApptDate = DateAdd("d", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                            End Select
                        Else
                            FirstDay = WeekDay(udtRepeatAppt.ApptDate)
                            If FirstDay = 1 Or FirstDay = 7 Then
                                FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbMonday)
                                udtRepeatAppt.ApptDate = DateAdd("d", nDays, udtRepeatAppt.ApptDate)
                            End If
                        End If
                        If Trim(udtRepeatAppt.ApptDate) <> vbNullString Then
                            If DateDiff("d", udtRepeatAppt.StartDate, udtRepeatAppt.ApptDate) >= 0 Then
                                udtRepeatAppt.RepeatID = lngRepeat
                                lngEndDate = DateDiff("d", udtRepeatAppt.ApptDate, udtRepeatAppt.EndDate)
                                If lngEndDate >= 0 Then
                                    If Get_SchedConflict_Old(udtRepeatAppt.ApptDate, udtRepeatAppt.ApptTime) Then
                                        Write_Repeat_Error
                                        bError = True
                                    Else
                                        Write_Repeat_Appt
                                    End If
                                End If
                            End If
                        End If
                        intDay = Format(udtRepeatAppt.StartDate, "d")
                        intMonth = Format(udtRepeatAppt.ApptDate, "m")
                        intMonth = Format(intMonth, "##")
                        intYear = Format(udtRepeatAppt.ApptDate, "yyyy")
                        udtRepeatAppt.ApptDate = DateAdd("m", 1, udtRepeatAppt.ApptDate)
                        intMonth = Format(udtRepeatAppt.ApptDate, "m")
                        intYear = Format(udtRepeatAppt.ApptDate, "yyyy")
                        udtRepeatAppt.ApptDate = intMonth & "/" & intDay & "/" & intYear
                        udtRepeatAppt.ApptDate = Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy")
                    Next Idx
                End If
            Case "Annual"
                udtRepeatAppt.ApptDate = vbNullString
                lngDateDiff = DateDiff("yyyy", udtAppointment.SchedDate, udtRepeatAppt.EndDate)
                lngDateDiff = lngDateDiff / nDays
                If lngDateDiff > 0 Then
                    If udtRepeatAppt.ApptDate = vbNullString Then udtRepeatAppt.ApptDate = udtRepeatAppt.StartDate
                    For Idx = 1 To lngDateDiff
                        If udtRepeatAppt.ApptDate <> vbNullString Then
                            Select Case udtRepeatAppt.WeekDay
                                Case "Monday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbMonday)
                                    udtRepeatAppt.ApptDate = DateAdd("yyyy", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Tuesday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbTuesday)
                                    udtRepeatAppt.ApptDate = DateAdd("yyyy", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Wednesday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbWednesday)
                                    udtRepeatAppt.ApptDate = DateAdd("yyyy", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Thursday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbThursday)
                                    udtRepeatAppt.ApptDate = DateAdd("yyyy", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Friday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbFriday)
                                    udtRepeatAppt.ApptDate = DateAdd("yyyy", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Saturday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbSaturday)
                                    udtRepeatAppt.ApptDate = DateAdd("yyyy", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                                Case "Sunday"
                                    FirstDay = WeekDay(udtRepeatAppt.ApptDate, vbSunday)
                                    udtRepeatAppt.ApptDate = DateAdd("yyyy", -(FirstDay - 1), udtRepeatAppt.ApptDate)
                            End Select
                            If Trim(udtRepeatAppt.ApptDate) <> vbNullString Then
                                lngEndDate = DateDiff("d", udtRepeatAppt.ApptDate, udtRepeatAppt.EndDate)
                                udtRepeatAppt.RepeatID = lngRepeat
                                If lngEndDate >= 0 Then
                                    If Get_SchedConflict_Old(udtRepeatAppt.ApptDate, udtRepeatAppt.ApptTime) Then
                                        Write_Repeat_Error
                                        bError = True
                                    Else
                                        Write_Repeat_Appt
                                    End If
                                End If
                            End If
                             udtRepeatAppt.ApptDate = DateAdd("yyyy", 1, udtRepeatAppt.ApptDate)
                        End If
                    Next Idx
                End If
        End Select
    End If
    Read_MultiAppt lngPatientID
    If bError Then
        MsgBox "One or more appointments could not be set due to scheduling conflicts. Please review Repeat Conflicts list!", vbInformation, "PatientTrac Multiple Appointments"
    Else
        MsgBox "Repeat Appointments has been successfully saved!", vbInformation, "PatientTrac Scheduling"
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Repeat_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Repeat_Appt", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Save_RepeatData(vstrTime As String, vstrDate As String) As Boolean
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim strTerms As String

    On Error GoTo Error_Save_RepeatData
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    vstrDate = Format(vstrDate, "mm/dd/yyyy")
    udtRepeatAppt.StartTime = vstrTime
    If udtRepeatAppt.RepeatMode = vbNullString Then
        udtRepeatAppt.RepeatMode = cboRepAppt.Text
    End If
    udtRepeatAppt.RepeatNumber = cboRepeatTerm
    udtRepeatAppt.StartTime = Format(udtRepeatAppt.StartTime, "hh:mm AM/PM")
    udtRepeatAppt.RepeatID = 0
    strTerms = Calc_Duration(udtRepeatAppt.PurposeVisit)
    udtRepeatAppt.StartTime = Format(udtRepeatAppt.StartTime, "hh:mm AM/PM")
    udtRepeatAppt.EndTime = DateAdd("mm", Val(Replace(strTerms, "minutes", "")), udtRepeatAppt.StartTime)
    strSQL = "Get_Repeat_Appt_ID " & Val(udtSchedPatient.IntId) & ",  " & udtRepeatAppt.FacilityID & ",'" _
            & Format(udtRepeatAppt.StartDate, "mm/dd/yyyy") & "', '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "', " & udtRepeatAppt.ProviderID
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        If Not .EOF Then
            udtRepeatAppt.RepeatID = (!REPEAT_ID)
        End If
    End With
    If udtRepeatAppt.RepeatID = 0 Then
        strSQL = "Write_Repeat_Appt " & Val(udtSchedPatient.IntId) & ", " & udtRepeatAppt.FacilityID & "," & udtRepeatAppt.ProviderID & ", " _
            & udtRepeatAppt.ResourseID & ",'" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "','" & udtRepeatAppt.ApptType & "','" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "','" _
            & Format(udtRepeatAppt.EndTime, "hh:mm AM/PM") & "','" & Format(udtRepeatAppt.StartDate, "mm/dd/yyyy") & "','" & Format(udtRepeatAppt.EndDate, "mm/dd/yyyy") & "'," _
            & Val(udtRepeatAppt.RepeatNumber) & ",'" & udtRepeatAppt.Duration & "','" & udtRepeatAppt.RepeatMode & "','" _
            & udtRepeatAppt.WeekDay & "','" & udtRepeatAppt.Comment & "'"
        Set rsCodes = cnCodes.Execute(strSQL)
        strSQL = "SELECT * FROM PATIENT_REPEAT_APPT where PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
            & "FACILITY_ID = " & udtRepeatAppt.FacilityID & " AND " _
            & "CONVERT(VARCHAR(10), START_DATE, 101) = '" & Format(udtRepeatAppt.StartDate, "mm/dd/yyyy") & "' " _
            & "AND CONVERT(VARCHAR(8), APPT_TIME, 108) = '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "' AND PROVIDER_ID = " & udtRepeatAppt.ProviderID
        Set rsCodes = cnCodes.Execute(strSQL)
        With rsCodes
            If Not .EOF Then
                udtRepeatAppt.RepeatID = (!REPEAT_ID)
            End If
        End With
    Else
        strSQL = "Update_Repeat_Appt '" & udtRepeatAppt.StartTime & "'," & udtRepeatAppt.ResourseID & ", " _
            & udtRepeatAppt.ProviderID & ", '" & udtRepeatAppt.ApptType & "', '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "',' " _
            & Format(udtRepeatAppt.EndTime, "hh:mm AM/PM") & "','" & Format(udtRepeatAppt.StartDate, "mm/dd/yyyy") & "','" & Format(udtRepeatAppt.EndDate, "mm/dd/yyyy") & "', " _
            & Val(udtRepeatAppt.RepeatNumber) & ",'" & udtRepeatAppt.Duration & "', '" & udtRepeatAppt.RepeatMode & "',' " _
            & udtRepeatAppt.WeekDay & "','" & udtRepeatAppt.Comment & "'," & udtRepeatAppt.RepeatID
        Set rsCodes = cnCodes.Execute(strSQL)
    End If
    
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Save_RepeatData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Save_RepeatData", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Read_MultiAppt(Optional vlngPatientID As Long)
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim i%
Dim lngPatientID As Long
Dim lngFacilityID As Long
Dim lngResID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim strRes As String
Dim strProv As String
Dim lngDateDiff As Long
Dim strTerm As String
Dim strStart As String
Dim strApptDate As String
Dim Arr

    On Error GoTo Error_Read_MultiAppt
    Screen.MousePointer = vbHourglass
    grdMultiAppt.MaxRows = 0
    If cboResProvider.listcount < 0 Then Get_Appt_Providers
    If cnSched.State = adStateClosed Then cnSched.open psConnect(0)
    udtRepeatAppt.ProviderID = 0
    udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    strSQL = "Read_MultiAppt  '" & udtTelephone.CallType & "', " & udtRepeatAppt.RepeatID
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            Do Until .EOF
                Clear_RepeatAppt
                lngDateDiff = 0
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "APPT_ID"
                                    udtRepeatAppt.Repeat_Appt_ID = Trim(Fld.Value)
                                Case "REPEAT_ID"
                                    udtRepeatAppt.RepeatID = Trim(Fld.Value)
                                Case "APPT_DATE"
                                    udtRepeatAppt.ApptDate = Trim(Fld.Value)
                                Case "PROVIDER_ID"
                                    udtRepeatAppt.ProviderID = Trim(Fld.Value)
                                Case "RESOURCE_ID"
                                    udtRepeatAppt.ResourseID = Trim(Fld.Value)
                                Case "FACILITY_ID"
                                    udtRepeatAppt.FacilityID = Trim(Fld.Value)
                                Case "APPT_TIME"
                                    udtRepeatAppt.StartTime = Trim(Fld.Value)
                                Case "PURPOSE_VISIT"
                                    udtRepeatAppt.PurposeVisit = Trim(Fld.Value)
                                Case "DURATION"
                                    udtRepeatAppt.Duration = Trim(Fld.Value)
                                Case "REPEAT_FREQ"
                                    udtRepeatAppt.RepeatNumber = Trim(Fld.Value)
                                Case "TERMINATION_DATE"
                                    udtRepeatAppt.EndDate = Trim(Fld.Value)
                                Case "START_DATE"
                                    udtRepeatAppt.StartDate = Trim(Fld.Value)
                                Case "START_TIME"
                                    udtRepeatAppt.StartTime = Trim(Fld.Value)
                                Case "END_TIME"
                                    udtRepeatAppt.EndTime = Trim(Fld.Value)
                                Case "REPEAT_MODE"
                                    udtRepeatAppt.RepeatMode = Trim(Fld.Value)
                                Case "REPEAT_DAY_WEEK"
                                    udtRepeatAppt.WeekDay = Trim(Fld.Value)
                                Case "NOTES"
                                    udtRepeatAppt.Comment = Trim(Literals(Fld.Value, "'"))
                                Case "APPT_TYPE"
                                    udtTelephone.CallType = Trim(Fld.Value)
                                Case "VISIT_CODE"
                                    udtRepeatAppt.VisitCode = Trim(Fld.Value)
                            End Select
                        End If
                    End If
                Next
                .MoveNext
                If Not TextFound("minutes", udtRepeatAppt.Duration, Len("minutes"), True) Then
                    udtRepeatAppt.Duration = Trim(udtRepeatAppt.Duration) & " " & "minutes"
                End If
                If Trim(udtRepeatAppt.ApptDate) <> vbNullString Then
                    With grdMultiAppt
                        lngRow = .MaxRows + 1
                        .MaxRows = lngRow
                        .SetRowItemData lngRow, udtRepeatAppt.Repeat_Appt_ID
                        lngCol = 0
                        Arr = Split(LookUpCellTag("APPT_DATE"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.ApptDate, "TypeHAlignCenter", "APPT_DATE"
                            End If
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag("FACILITY_ID"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If udtRepeatAppt.FacilityID > 0 Then
                                Get_ReportFacility udtRepeatAppt.FacilityID
                            End If
                            Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtFacility.Name, "TypeHAlignCenter", "FACILITY_ID"
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag("APPT_TIME"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.StartTime, "TypeHAlignCenter", "APPT_TIME"
                            End If
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag("START_DATE"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.StartDate, "TypeHAlignCenter", "START_DATE"
                            End If
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag("DURATION"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.Duration, "TypeHAlignCenter", "DURATION"
                            End If
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag("REPEAT_DAY_WEEK"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.WeekDay, "TypeHAlignCenter", "REPEAT_DAY_WEEK"
                            End If
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag("APPT_TYPE"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.PurposeVisit, "TypeHAlignCenter", "APPT_TYPE"
                            End If
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag("TERMINATION_DATE"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.EndDate, "TypeHAlignCenter", "TERMINATION_DATE"
                            End If
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag("NOTES"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.Comment, "TypeHAlignCenter", "NOTES"
                            End If
                        End If
                        lngCol = 0
                        Arr = Split(LookUpCellTag("APPT_ID"), "|")
                        If UBound(Arr) >= 0 Then
                           lngCol = Arr(0)
                           If lngCol > 0 Then
                              Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", CStr(udtRepeatAppt.Repeat_Appt_ID), "TypeHAlignCenter", "APPT_ID"
                           End If
                        End If
                        For lngCol = 1 To grdMultiAppt.MaxCols
                            .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Next lngCol
                    End With
                End If
            Loop
        End With
        If udtRepeatAppt.ProviderID > 0 Then
            Select Case IsProperCase(udtTelephone.CallType)
                Case "Patient"
                    Get_SchedPatient Val(udtSchedPatient.IntId)
                    Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "10", "REPEAT APPOINTMENT SCHEDULE FOR" _
                        & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName) & vbNewLine _
                        & "Scheduled Provider:" & " " & UCase(Get_Appt_Prov(udtRepeatAppt.ProviderID)), "TypeHAlignCenter", "APPT_DATE", HeaderColor
                Case "Vendor", "Business", "Vendors"
                    Get_SchedVendor Val(udtSchedPatient.IntId), udtAppointment.SchedTime
                    Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "9", "REPEAT APPOINTMENT SCHEDULE FOR" _
                        & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName) & vbNewLine _
                        & "Scheduled Provider:" & " " & UCase(Get_Appt_Prov(udtRepeatAppt.ProviderID)), "TypeHAlignCenter", "APPT_DATE", HeaderColor
                Case "Legal", "Attorney"
                    Get_Schedule_Attorney Val(udtSchedPatient.IntId)
                    Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "9", "REPEAT APPOINTMENT SCHEDULE FOR" _
                        & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName) _
                        & vbNewLine & "Scheduled Provider:" & " " & UCase(Get_Appt_Prov(udtRepeatAppt.ProviderID)), "TypeHAlignCenter", "APPT_DATE", HeaderColor
                Case "Physician"
                    Get_SchedPhysician Val(udtSchedPatient.IntId)
                    Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "9", "REPEAT APPOINTMENT SCHEDULE FOR" _
                        & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName) _
                        & vbNewLine & "Scheduled Provider:" & " " & UCase(Get_Appt_Prov(udtRepeatAppt.ProviderID)), "TypeHAlignCenter", "APPT_DATE", HeaderColor
            End Select
        End If
    If cnSched.State = adStateOpen Then
        cnSched.Close
        Set cnSched = Nothing
        Set rsSched = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_MultiAppt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_MultiAppt", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cmdUpdate_Click()
Dim lngRow As Long
Dim lngRepeat As Long

    cmdSched.Visible = False
    With grdRepeatHist
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                lngRepeat = .GetRowItemData(lngRow)
                Exit For
            End If
        Next lngRow
    End With
    If lngRepeat = 0 Then
        MsgBox "Please select repeat appointment first!", vbInformation, "PatientTrac Repeat Appointments"
        Exit Sub
    Else
        udtRepeatAppt.RepeatID = lngRepeat
    End If
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please Select Patient First!", vbInformation, "PatientTrac Repeat Appointments"
       Exit Sub
    End If
    Clear_Fields
    Load_ComboTime
    If cboDuration.listcount <= 0 Then Load_Duration
    Read_Parent_Visit
    fraControl.Visible = True
    Exit Sub
End Sub
Private Sub Form_Activate()
Dim sValue As String
Dim Idx As Integer

    Exit Sub
End Sub

Private Sub Form_Load()
    
    lblPatName.Caption = vbNullString
    Clear_SchedPatient
    Clear_Block
    Clear_RepeatAppt
    ccdTermDate.Value = vbNullString
    ccdStart.Value = vbNullString
    Clear_RepeatAppt
    Load_Form
    Clear_clsWriteAppt
    centerForm Me
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Clear_RepeatAppt
    Unload Me
End Sub
Private Function Load_Repeat_Appt()
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim i%
Dim lngPatientID As Long
Dim lngFacilityID As Long
Dim lngResID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim strRes As String
Dim strProv As String
Dim lngDateDiff As Long
Dim strTerm As String
Dim strStart As String
Dim strApptDate As String
Dim Arr
Dim lngRepeat As Long
Dim lngTestRow As Long
Dim bExit As Boolean


    On Error GoTo Error_Load_Repeat_Appt
    Screen.MousePointer = vbHourglass
    grdRepeatHist.MaxRows = 0
    If cnSched.State = adStateClosed Then cnSched.open psConnect(0)
    udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    strSQL = "Get_Repeat_Appt_All " & Val(udtSchedPatient.IntId)
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            Do Until .EOF
                Clear_RepeatAppt
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "PROVIDER_ID"
                                    udtRepeatAppt.ProviderID = Trim(Fld.Value)
                                Case "RESOURCE_ID"
                                    udtRepeatAppt.ResourseID = Trim(Fld.Value)
                                Case "FACILITY_ID"
                                    udtRepeatAppt.FacilityID = Trim(Fld.Value)
                                Case "TERMINATION_DATE"
                                    udtRepeatAppt.EndDate = Trim(Fld.Value)
                                Case "START_DATE"
                                    udtRepeatAppt.StartDate = Trim(Fld.Value)
                                Case "REPEAT_DAY_WEEK"
                                    udtRepeatAppt.WeekDay = Trim(Fld.Value)
                                Case "REPEAT_ID"
                                    udtRepeatAppt.RepeatID = Trim(Fld.Value)
                                    lngRepeat = udtRepeatAppt.RepeatID
                                Case "ACTIVE"
                                    udtRepeatAppt.Status = Trim(Fld.Value)
                                    If udtRepeatAppt.Status = "True" Then
                                        udtRepeatAppt.Status = "Active"
                                    Else
                                        udtRepeatAppt.Status = "InActive"
                                    End If
                            End Select
                        End If
                    End If
                Next
                .MoveNext
                With grdRepeatHist
                    bExit = False
                    For lngTestRow = 1 To .MaxRows
                        If lngRepeat = .GetRowItemData(lngTestRow) Then
                            bExit = True
                            Exit For
                        End If
                    Next lngTestRow
                    If Not bExit Then
                        lngRow = .MaxRows + 1
                        .MaxRows = lngRow
                        .SetRowItemData lngRow, udtRepeatAppt.RepeatID
                        lngCol = 0
                         Arr = Split(LookUpCellTag_RepeatHist("FACILITY_ID"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If udtRepeatAppt.FacilityID > 0 Then
                                Get_ReportFacility udtRepeatAppt.FacilityID
                            End If
                            Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", udtFacility.Name, "TypeHAlignCenter", "FACILITY_ID"
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag_RepeatHist("PROVIDER_ID"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If udtRepeatAppt.ProviderID > 0 Then
                                Get_Providers udtRepeatAppt.ProviderID
                            End If
                            Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", clsProvider.Name, "TypeHAlignCenter", "PROVIDER_ID"
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag_RepeatHist("START_DATE"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", udtRepeatAppt.StartDate, "TypeHAlignCenter", "START_DATE"
                            End If
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag_RepeatHist("TERMINATION_DATE"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", udtRepeatAppt.EndDate, "TypeHAlignCenter", "TERMINATION_DATE"
                            End If
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag_RepeatHist("REPEAT_DAY_WEEK"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeStatic grdRepeatHist, lngRow, lngCol, "1", udtRepeatAppt.WeekDay, "TypeHAlignCenter", "REPEAT_DAY_WEEK"
                            End If
                        End If
                        lngCol = 0
                         Arr = Split(LookUpCellTag_RepeatHist("ACTIVE"), "|")
                        If UBound(Arr) >= 0 Then
                            lngCol = Arr(0)
                            If lngCol > 0 Then
                                Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", udtRepeatAppt.Status, "TypeHAlignCenter", "ACTIVE"
                            End If
                        End If
                    End If
                    For lngCol = 1 To .MaxCols
                        .SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Next lngCol
                End With
            Loop
        End With
    If cnSched.State = adStateOpen Then
        cnSched.Close
        Set cnSched = Nothing
        Set rsSched = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Repeat_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Repeat_Appt", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Format_MultiAppt()
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr

    On Error GoTo Error_Format_MultiAppt
    Screen.MousePointer = vbHourglass
    If udtTelephone.CallType = vbNullString Then udtTelephone.CallType = "Patient"
    With grdMultiAppt
        .Redraw = False
        .ClearRange 1, SpreadHeader, .MaxCols, .MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 10
        .MaxCols = 9
        .FontName = "Times New Roman"
        .GridShowHoriz = True
        .GridShowVert = True
        .RowHeadersShow = False
        .ColHeaderRows = 4
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .ColWidth(1) = 18
        .ColWidth(2) = 12
        .ColWidth(3) = 12
        .ColWidth(4) = 15
        .ColWidth(5) = 10
        .ColWidth(6) = 9
        .ColWidth(7) = 9
        .ColWidth(8) = 12
        .ColWidth(9) = 15.3
        .ColWidth(10) = 0
        .MaxRows = 0
        .Row = SpreadHeader
        .Col = 10
        .ColHidden = True
        lngRow = SpreadHeader
        lngCol = 1
        Select Case IsProperCase(udtTelephone.CallType)
            Case "Patient"
                Get_SchedPatient Val(udtSchedPatient.IntId)
                Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "10", "REPEAT APPOINTMENT SCHEDULE FOR" & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName), "TypeHAlignCenter", "APPT_DATE", HeaderColor
            Case "Vendor", "Business", "Vendors"
                Get_SchedVendor Val(udtSchedPatient.IntId), udtAppointment.SchedTime
                Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "9", "REPEAT APPOINTMENT SCHEDULE FOR" & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName), "TypeHAlignCenter", "APPT_DATE", HeaderColor
            Case "Legal", "Attorney"
                Get_Schedule_Attorney Val(udtSchedPatient.IntId)
                Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "9", "REPEAT APPOINTMENT SCHEDULE FOR" & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName), "TypeHAlignCenter", "APPT_DATE", HeaderColor
            Case "Physician"
                Get_SchedPhysician Val(udtSchedPatient.IntId)
                Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "9", "REPEAT APPOINTMENT SCHEDULE FOR" & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName), "TypeHAlignCenter", "APPT_DATE", HeaderColor
        End Select
        grdMultiAppt.SetCellBorder 1, SpreadHeader, 9, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .RowHeight(SpreadHeader) = 20.5
        Add_Spacer SpreadHeader + 1, grdMultiAppt
        lngRow = SpreadHeader + 2
        lngCol = 1
        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", "FACILITY", "TypeHAlignCenter", "FACILITY_ID", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", "APPT DATE", "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", "APPT TIME", "TypeHAlignCenter", "APPT_TIME", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", "APPT TYPE", "TypeHAlignCenter", "APPT_TYPE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", "DURATION", "TypeHAlignCenter", "DURATION", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", "START DATE", "TypeHAlignCenter", "START_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", "END DATE", "TypeHAlignCenter", "TERMINATION_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 8
        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", "WEEK DAY", "TypeHAlignCenter", "REPEAT_DAY_WEEK", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 9
        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", "COMMENTS", "TypeHAlignCenter", "NOTES", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 10
        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", "APPT ID", "TypeHAlignCenter", "APPT_ID", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdMultiAppt
        .Redraw = False
        .MaxRows = 0
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_MultiAppt:
    Screen.MousePointer = vbDefault
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

    grdMultiAppt.getSelection 0, Col, Row, Col2, Row2
    grdMultiAppt.Redraw = False
    For i = 1 To 7
        For j = SpreadHeader To SpreadHeader + 2
            Span = grdMultiAppt.GetCellSpan(i, j, anchorcol, anchorrow, colspan, rowspan)
            grdMultiAppt.Redraw = False
            If Span Then
                If grdMultiAppt.Col = 1 Then anchorcol = "1"
                If j = SpreadHeader Then anchorrow = SpreadHeader
                If j = SpreadHeader Then anchorrow = SpreadHeader
                grdMultiAppt.RemoveCellSpan anchorcol, anchorrow
                IsSpan = True
            End If
        Next j
    Next i
    grdMultiAppt.Redraw = True
    Exit Function
End Function

Private Sub fraNumber_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    flngDeltaX = X
    flngDeltaY = y
End Sub

Private Sub fraNumber_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        fraNumber.Left = fraNumber.Left + X - flngDeltaX
        fraNumber.Top = fraNumber.Top + y - flngDeltaY
    End If
End Sub


Private Sub grdMultiAppt_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim lngCol As Long
Dim lngx As Long
Dim lngY As Long
Dim lngWidth As Long
Dim lngHeight As Long
Dim varCellValue
Dim Idx As Integer
Dim Idx1 As Integer
Dim Arr

    
    If Row <= 0 Then Exit Sub
    With grdMultiAppt
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
Public Sub SetcboTime(ipCBO As Object, ipHour As String, ipAMPM As String, ipInt As Long, ipEnd As Long)
    Dim Idx As Long
    
    For Idx = 0 To ipEnd Step ipInt
        ipCBO.AddItem ipHour & ":" & Format(Idx, "00") & " " & ipAMPM
    Next Idx
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
        udtRepeatAppt.PurposeVisit = varCellValue
        .BlockMode = False
        .GetText 2, Row, varCellValue
        udtRepeatAppt.VisitCode = Left(varCellValue, 5)
        Calc_Duration udtEncounter.Appt_Type
        LookUpCPTCode udtEncounter.Appt_Type
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
        Else
            If cboDuration.listcount <= 0 Then Load_Duration
            For Idx = 0 To cboDuration.listcount - 1
                If cboDuration.List(Idx) = udtEncounter.VisitTime Then
                    cboDuration.Text = cboDuration.List(Idx)
                    Exit For
                End If
            Next Idx
        End If
        udtRepeatAppt.VisitCode = Left(udtRepeatAppt.VisitCode, 5)
        udtRepeatAppt.VisitCode = udtEncounter.Appt_Type
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
            .Col2 = 5
            .Row = Row
            .Row2 = Row
            .BackColor = grdCurItem
        .BlockMode = False
        If IsProperCase(udtTelephone.CallType) = "Patient" Or udtTelephone.CallType = vbNullString Then
            If Col = 3 Then
                .Col = 1
                .Row = Row
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
    End With
    Procedure_Grid_Action
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_grdProcedure_DblClick:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "grdProcedure_DblClick", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Function Validate_Between_Date(vlngCnt As Long) As Boolean
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim sValue As String
Dim intCount As Integer
Dim strDuration As String
Dim strAppt_Time As String


    On Error GoTo Error_Validate_Between_Date
    Screen.MousePointer = vbHourglass
    Validate_Between_Date = False
    sValue = DateAdd("n", Val(udtRepeatAppt.Duration), udtRepeatAppt.ApptTime)
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(0)
    strSQL = "SELECT * FROM APPOINTMENTS WHERE PROVIDER_ID = " & udtRepeatAppt.ProviderID & " AND " _
        & "FACILITY_ID = " & udtRepeatAppt.FacilityID & " AND  " _
        & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "' AND " _
        & "ACTIVE = 1 "
    Set rsBlock = cnTerms.Execute(strSQL)
    With rsBlock
        Do Until .EOF
            Clear_Block
            udtBlock.StartTime = (!Appt_Time)
            udtBlock.Duration = Replace((!Duration), "minutes", "")
            .MoveNext
            udtBlock.EndTime = DateAdd("n", Val(udtBlock.Duration), udtBlock.StartTime)
            If DateDiff("n", Format(udtBlock.StartTime, "hh:mm AM/PM"), Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM")) > 0 Then
                If DateDiff("n", Format(udtBlock.EndTime, "hh:mm AM/PM"), Format(sValue, "hh:mm AM/PM")) < 0 Then
                    intCount = intCount + 1
                End If
            ElseIf DateDiff("n", Format(udtBlock.StartTime, "hh:mm AM/PM"), Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM")) < 0 Then
                If DateDiff("n", Format(udtBlock.StartTime, "hh:mm AM/PM"), Format(sValue, "hh:mm AM/PM")) > 0 Then
                    intCount = intCount + 1
                End If
            ElseIf DateDiff("n", Format(udtBlock.StartTime, "hh:mm AM/PM"), Format(udtRepeatAppt.ApptTime, "hh:mm AM/PM")) = 0 Then
                intCount = intCount + 1
            End If
        Loop
    End With
    If intCount > 0 Then
        If intCount >= Val(vlngCnt) Then
            Validate_Between_Date = True
            udtRepeatAppt.ConflictReason = "Overbook Appt"
        End If
    End If
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsBlock = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_Between_Date:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_Between_Date", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


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

Private Sub grdPsychotherapy_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue
Dim lngRow As Long
Dim lngCol As Long
Dim strEnterCode As String
Dim bFound As Boolean
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
                    If Trim(udtRepeatAppt.VisitCode) = "90833" Or _
                        Trim(udtRepeatAppt.VisitCode) = "90836" Or _
                        Trim(udtRepeatAppt.VisitCode) = "90838" Or _
                        Trim(udtRepeatAppt.VisitCode) = "90863" Then
                    End If
                End If
            End If
            If bFound Then
                  Select Case strEnterCode
                    Case "90840"
                        If udtRepeatAppt.VisitCode <> "90839" Then
                            MsgBox "90840 Add-On can only be added to a 98839 Psychotherapy for Crisis visit code!", vbInformation, "PatientTrac EM Coding"
                            bCancel = True
                            Exit For
                        End If
                      Case "90863" '
                        If udtRepeatAppt.VisitCode = "99211" Or _
                            udtRepeatAppt.VisitCode = "99212" Or _
                            udtRepeatAppt.VisitCode = "99213" Or _
                            udtRepeatAppt.VisitCode = "99214" Or _
                            udtRepeatAppt.VisitCode = "99215" Then
                            MsgBox "Cannot bill the medication management add-on (90863) code with an" & " " & "EM Visit Code!", vbInformation, "PatientTrac EM Coding"
                            bCancel = True
                            Exit For
                        End If
                        If Not bCancel Then
                            If Not TextFound("90832", udtRepeatAppt.VisitCode, Len("90832"), True) Or _
                                Not TextFound("90834", udtRepeatAppt.VisitCode, Len("90834"), True) Or _
                                Not TextFound("90837", udtRepeatAppt.VisitCode, Len("90837"), True) Then
                                    MsgBox "Cannot bill the medication management add-on (90863) code with an" & " " & udtRepeatAppt.VisitCode & " " & "visit code!", vbInformation, "PatientTrac EM Coding"
                                    bCancel = True
                                    Exit For
                            End If
                        End If
                      Case "90785"
                              Select Case udtRepeatAppt.VisitCode
                                  Case "90839", "90840", "90846", "990847", "90849"
                                      MsgBox "Cannot bill the 90785 code with" & " " & udtRepeatAppt.VisitCode & "!", vbInformation, "PatientTrac EM Coding"
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
                                              MsgBox "Cannot bill the 90785 code with" & " " & udtRepeatAppt.VisitCode & " " & "EM Code must include psychotherapy add-on!", vbInformation, "PatientTrac EM Coding"
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
                        ElseIf Not TextFound("90832", udtRepeatAppt.VisitCode, Len("90832"), True) Or _
                            Not TextFound("90834", udtRepeatAppt.VisitCode, Len("90834"), True) Or _
                            Not TextFound("90837", udtRepeatAppt.VisitCode, Len("90837"), True) Then
                                MsgBox "Cannot bill the medication management add-on (90863) code with an" & " " & udtRepeatAppt.VisitCode & " " & "visit code!", vbInformation, "PatientTrac EM Coding"
                                bCancel = True
                                Exit For
                        End If
                    Case "90840"
                        If udtRepeatAppt.VisitCode <> "90839" Then
                            MsgBox "90840 Add-On can only be added to a 98839 Psychotherapy for Crisis visit code!", vbInformation, "PatientTrac EM Coding"
                            bCancel = True
                            Exit For
                        End If
                      Case "90785"
                            Select Case udtRepeatAppt.VisitCode
                                Case "90839", "90840", "90846", "990847", "90849"
                                    MsgBox "Cannot bill the 90785 add-code with" & " " & udtRepeatAppt.VisitCode & " " & "visit code!", vbInformation, "PatientTrac EM Coding"
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

        If IsProperCase(udtTelephone.CallType) = "Patient" Or udtTelephone.CallType = vbNullString Then
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


Private Sub grdRepeatHist_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long

    On Error GoTo Error_GridRepeat
    
    With grdRepeatHist
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
           cmdDelete.Visible = True
        .BlockMode = False
        udtRepeatAppt.RepeatID = .GetRowItemData(Row)
        Format_MultiAppt
        Read_MultiAppt
    End With
    Exit Sub
Error_GridRepeat:
    Exit Sub
    Resume Next
End Sub

Private Sub mnu_Close_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub mnu_Print_Report_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(1) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    Add_Header
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Multi Appointment Schedule"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    With grdMultiAppt
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Repeat Appointments"
        .PrintColor = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintSmartPrint = True
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Multiple Appointments"
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

Private Sub mnu_Unscheduled_Appt_Click()

    If Val(udtSchedPatient.IntId) = 0 Then
        MsgBox "Please select patient first!", vbInformation, "PatientTrac Repeat Appointments"
        Exit Sub
    End If
    If udtRepeatAppt.RepeatID = 0 Then
        MsgBox "Please select repeat appointments row first!", vbInformation, "PatientTrac Repeat Appointments"
        Exit Sub
    End If
    lblPatName.Visible = True
    lblPatName.Caption = udtSchedPatient.Name
    Format_Conflict
    Get_SchedConflict
    fraSchedConflicts.Visible = True
    
    Exit Sub
End Sub


