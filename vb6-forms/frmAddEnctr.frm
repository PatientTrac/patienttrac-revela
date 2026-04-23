VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmAddEnctr 
   Caption         =   "New Encounter Details"
   ClientHeight    =   9180
   ClientLeft      =   3330
   ClientTop       =   750
   ClientWidth     =   14310
   Icon            =   "frmAddEnctr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   14310
   Begin VB.Frame fraEnctr 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   60
      TabIndex        =   27
      Top             =   2130
      Width           =   14205
      Begin VB.Frame fraMaint 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Encounter ToolBar"
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
         Height          =   1515
         Left            =   90
         TabIndex        =   35
         Top             =   60
         Width           =   3315
         Begin VB.Frame fraAppoint 
            BackColor       =   &H00E0E0E0&
            Height          =   1125
            Left            =   60
            TabIndex        =   36
            Top             =   2010
            Visible         =   0   'False
            Width           =   5745
            Begin VB.ComboBox ccNextAppt 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   2730
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   900
               Visible         =   0   'False
               Width           =   2595
            End
            Begin VB.CheckBox chkNoAppt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "No Appointment"
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
               Left            =   2760
               TabIndex        =   37
               Top             =   600
               Width           =   2625
            End
            Begin MSComCtl2.DTPicker ccApptDate 
               Height          =   375
               Left            =   30
               TabIndex        =   39
               Top             =   120
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
               Format          =   85721091
               CurrentDate     =   38792
            End
            Begin PT_XP_Button.PT_XP cmdAddTerms 
               Height          =   285
               Left            =   5370
               TabIndex        =   40
               ToolTipText     =   "Add Terms"
               Top             =   960
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               ForeColor       =   8388608
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AppearanceThemes=   3
               TransparentColor=   14215660
               Caption         =   "..."
               DepthEvent      =   1
               ForeColorOver   =   8388608
               ForeColorFocus  =   8388608
               ForeColorDown   =   8388608
               PictureDisabled =   "frmAddEnctr.frx":0442
               PictureAlignment=   4
               ShowFocus       =   -1  'True
            End
            Begin VB.Label lblNextAppt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Next Appointment"
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
               Height          =   210
               Left            =   0
               TabIndex        =   41
               Top             =   1050
               Width           =   1605
            End
         End
         Begin PT_XP_Button.PT_XP cmdSave 
            Height          =   345
            Left            =   1980
            TabIndex        =   42
            Top             =   300
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
            Picture         =   "frmAddEnctr.frx":0794
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "Save"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAddEnctr.frx":0D2E
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdNew 
            Height          =   345
            Left            =   120
            TabIndex        =   43
            Top             =   270
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
            Picture         =   "frmAddEnctr.frx":1080
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "New"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAddEnctr.frx":161A
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdUpdate 
            Height          =   345
            Left            =   1980
            TabIndex        =   44
            Top             =   780
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
            Picture         =   "frmAddEnctr.frx":196C
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "Update"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAddEnctr.frx":1F06
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdEndEncounter 
            Height          =   345
            Left            =   120
            TabIndex        =   45
            Top             =   690
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
            Picture         =   "frmAddEnctr.frx":2258
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "End"
            Depth           =   5
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAddEnctr.frx":27F2
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdDeleteEncounter 
            Height          =   345
            Left            =   120
            TabIndex        =   46
            Top             =   1110
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
            Picture         =   "frmAddEnctr.frx":2B44
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "Delete"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAddEnctr.frx":30DE
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblEnctrTools 
            BackStyle       =   0  'Transparent
            Caption         =   "Encounter Toolbar"
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
            Left            =   690
            TabIndex        =   47
            Top             =   -30
            Width           =   2025
         End
      End
      Begin VB.Frame fraAdmission 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   3360
         TabIndex        =   28
         Top             =   60
         Visible         =   0   'False
         Width           =   6165
         Begin MSComCtl2.DTPicker ccdDisharge 
            Height          =   375
            Left            =   3600
            TabIndex        =   29
            Top             =   270
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   16761024
            CalendarTitleForeColor=   8388608
            CalendarTrailingForeColor=   8421504
            CheckBox        =   -1  'True
            CustomFormat    =   "MM/dd/yyy hh:mm:ss"
            DateIsNull      =   -1  'True
            Format          =   85721091
            CurrentDate     =   38862.7362384259
            MinDate         =   37257
         End
         Begin MSComCtl2.DTPicker ccdAdmit 
            Height          =   375
            Left            =   690
            TabIndex        =   30
            Top             =   270
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   16761024
            CalendarTitleForeColor=   8388608
            CalendarTrailingForeColor=   8421504
            CheckBox        =   -1  'True
            CustomFormat    =   "MM/dd/yyy hh:mm:ss"
            DateIsNull      =   -1  'True
            Format          =   85721091
            CurrentDate     =   38862.7362384259
            MinDate         =   37257
         End
         Begin PT_XP_Button.PT_XP cmdDischarge 
            Height          =   225
            Left            =   5400
            TabIndex        =   31
            ToolTipText     =   "Enter Discharge Date"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   397
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
            Picture         =   "frmAddEnctr.frx":3430
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "OK"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAddEnctr.frx":37CA
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdAdmit 
            Height          =   225
            Left            =   2550
            TabIndex        =   32
            ToolTipText     =   "Enter/Reset Admission Date"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   397
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
            Picture         =   "frmAddEnctr.frx":3B1C
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "OK"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAddEnctr.frx":3EB6
            ShowFocus       =   -1  'True
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Discharge Date"
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
            Left            =   3630
            TabIndex        =   34
            Top             =   30
            Width           =   1725
         End
         Begin VB.Label lblAdmt 
            BackStyle       =   0  'Transparent
            Caption         =   "Admission Date"
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
            Left            =   690
            TabIndex        =   33
            Top             =   60
            Width           =   1695
         End
      End
      Begin FPSpreadADO.fpSpread grdEnctrSummary 
         Height          =   4515
         Left            =   0
         TabIndex        =   48
         Top             =   1770
         Width           =   14115
         _Version        =   458752
         _ExtentX        =   24897
         _ExtentY        =   7964
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
         MaxCols         =   15
         MaxRows         =   50
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   16761024
         ShadowText      =   8388608
         SpreadDesigner  =   "frmAddEnctr.frx":4208
         UnitType        =   0
         UserResize      =   0
         Appearance      =   2
         ClipboardOptions=   0
      End
      Begin PT_XP_Button.PT_XP cmdNext 
         Height          =   405
         Left            =   6630
         TabIndex        =   49
         Top             =   1170
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
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
         Picture         =   "frmAddEnctr.frx":451F
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Next"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAddEnctr.frx":48B9
         PictureAlignment=   5
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrev 
         Height          =   405
         Left            =   5250
         TabIndex        =   50
         Top             =   1170
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
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
         Picture         =   "frmAddEnctr.frx":4C0B
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Previous "
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAddEnctr.frx":4FA5
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCancel 
         Height          =   345
         Left            =   7110
         TabIndex        =   51
         Top             =   6330
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
         Picture         =   "frmAddEnctr.frx":52F7
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAddEnctr.frx":5D09
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdEnctrOK 
         Height          =   345
         Left            =   6000
         TabIndex        =   52
         Top             =   6330
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
         Picture         =   "frmAddEnctr.frx":605B
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "OK"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAddEnctr.frx":63F5
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdEntrDefault 
         Height          =   345
         Left            =   12360
         TabIndex        =   53
         Top             =   6390
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
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
         Picture         =   "frmAddEnctr.frx":6747
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Set Defaults"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAddEnctr.frx":6AE1
         ShowFocus       =   -1  'True
      End
      Begin MSComCtl2.DTPicker ccdStartDate 
         Height          =   375
         Left            =   10920
         TabIndex        =   54
         Top             =   210
         Width           =   2745
         _ExtentX        =   4842
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
         CalendarForeColor=   0
         CalendarTitleBackColor=   16761024
         CalendarTitleForeColor=   8388608
         CalendarTrailingForeColor=   8421504
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyy hh:mm:ss"
         DateIsNull      =   -1  'True
         Format          =   85721091
         CurrentDate     =   38862.7355555556
         MinDate         =   37257
      End
      Begin MSComCtl2.DTPicker ccdEndDate 
         Height          =   375
         Left            =   10920
         TabIndex        =   55
         Top             =   720
         Width           =   2745
         _ExtentX        =   4842
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
         CalendarForeColor=   0
         CalendarTitleBackColor=   16761024
         CalendarTitleForeColor=   8388608
         CalendarTrailingForeColor=   8421504
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyy hh:mm:ss"
         DateIsNull      =   -1  'True
         Format          =   85721091
         CurrentDate     =   38862.7362384259
         MinDate         =   38353
      End
      Begin PT_XP_Button.PT_XP cmdLookUpPatient 
         Height          =   465
         Left            =   11550
         TabIndex        =   56
         Top             =   2070
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
         Picture         =   "frmAddEnctr.frx":6E33
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Schedule Patient"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAddEnctr.frx":782D
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblHistory 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encounter History"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   5550
         OLEDropMode     =   1  'Manual
         TabIndex        =   59
         Top             =   840
         Width           =   2130
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         Left            =   9780
         TabIndex        =   58
         Top             =   810
         Width           =   915
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
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
         Left            =   9750
         TabIndex        =   57
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.Frame fraTitleBar 
      BackColor       =   &H00E0E0E0&
      Height          =   2085
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   14235
      Begin VB.Frame fraVisitReason 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
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
         Height          =   525
         Left            =   7620
         TabIndex        =   16
         Top             =   1500
         Width           =   6495
         Begin VB.TextBox txtVisitReason 
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
            Height          =   465
            Left            =   1950
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   60
            Width           =   4455
         End
         Begin VB.ComboBox cboVisitReason 
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
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Please select Type of Post Operative Visit"
            Top             =   60
            Visible         =   0   'False
            Width           =   4455
         End
         Begin PT_XP_Button.PT_XP cmdVisitReas 
            Height          =   255
            Left            =   1530
            TabIndex        =   19
            ToolTipText     =   "Search Variables"
            Top             =   240
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
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
            TransparentColor=   14215660
            Caption         =   "..."
            DepthMode       =   1
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAddEnctr.frx":81AF
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblVisit 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Visit Reason "
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
            Left            =   60
            TabIndex        =   20
            Top             =   150
            Width           =   1365
         End
      End
      Begin VB.Frame fraEnctrDetail 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   2340
         TabIndex        =   2
         Top             =   120
         Width           =   11775
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
            Left            =   7230
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   0
            Width           =   4455
         End
         Begin VB.ComboBox cboFacility 
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
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Width           =   3345
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
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmAddEnctr.frx":8501
            Left            =   1890
            List            =   "frmAddEnctr.frx":8503
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Tag             =   "Place of Service"
            Top             =   405
            Width           =   3345
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
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmAddEnctr.frx":8505
            Left            =   7230
            List            =   "frmAddEnctr.frx":8507
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "New or Established Patient"
            Top             =   435
            Width           =   4455
         End
         Begin VB.ComboBox cboPurVisit 
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
            Left            =   7230
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   870
            Width           =   4455
         End
         Begin PT_XP_Button.PT_XP cmdUpdProc 
            Height          =   255
            Left            =   5400
            TabIndex        =   8
            ToolTipText     =   "Upate Surgical Procedure"
            Top             =   1110
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
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
            TransparentColor=   14215660
            Caption         =   "..."
            DepthMode       =   1
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAddEnctr.frx":8509
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblProvider 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   5400
            TabIndex        =   15
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lblFacility 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   30
            TabIndex        =   14
            Top             =   150
            Width           =   765
         End
         Begin VB.Label lblPatType 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Place of Service:"
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
            Left            =   30
            TabIndex        =   13
            Top             =   540
            Width           =   1755
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Status:"
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
            Left            =   5400
            TabIndex        =   12
            Top             =   510
            Width           =   1470
         End
         Begin VB.Label lblUser 
            BackStyle       =   0  'Transparent
            Caption         =   "Current Logon User: "
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
            Height          =   405
            Left            =   30
            TabIndex        =   11
            Top             =   870
            Width           =   1785
         End
         Begin VB.Label lblUserName 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   465
            Left            =   1890
            TabIndex        =   10
            Top             =   870
            Width           =   3345
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblPurVisit 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Encounter Type"
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
            Left            =   5400
            TabIndex        =   9
            Top             =   900
            Width           =   1725
         End
      End
      Begin VB.ComboBox cboRefPhys 
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
         Left            =   4230
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1500
         Width           =   3375
      End
      Begin PT_XP_Button.PT_XP cmdAdminHistory 
         Height          =   495
         Left            =   150
         TabIndex        =   21
         Top             =   630
         Width           =   1965
         _ExtentX        =   3466
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
         TransparentColor=   14215660
         Caption         =   "Hospital Admission History"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPatient 
         Height          =   405
         Left            =   150
         TabIndex        =   22
         Top             =   150
         Width           =   1965
         _ExtentX        =   3466
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
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Patient Select"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdBillingInfo 
         Height          =   495
         Left            =   150
         TabIndex        =   23
         Top             =   570
         Width           =   1965
         _ExtentX        =   3466
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
         TransparentColor=   14215660
         Caption         =   "Billing Information"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSurgLog 
         Height          =   465
         Left            =   150
         TabIndex        =   24
         Top             =   1170
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
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
         Picture         =   "frmAddEnctr.frx":885B
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Surgery Log"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAddEnctr.frx":9255
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClear 
         Height          =   255
         Left            =   3810
         TabIndex        =   25
         ToolTipText     =   "Clear Referral"
         Top             =   1560
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
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
         TransparentColor=   14215660
         Caption         =   "..."
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAddEnctr.frx":9BD7
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblRendProv 
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
         Height          =   540
         Left            =   2370
         TabIndex        =   26
         Top             =   1470
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmAddEnctr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean

Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
    End If
    
    centerForm Me
    Exit Sub
End Sub


Private Sub Form_Load()
    mFirst = True
    Screen.MousePointer = vbDefault
Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Exit Sub
End Sub


