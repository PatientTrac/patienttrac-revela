VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{76F358A5-0FAA-4EA4-8073-9E9D9A484F37}#1.0#0"; "CIMAGE~1.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmSysConfig 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Configuration"
   ClientHeight    =   6030
   ClientLeft      =   9120
   ClientTop       =   3330
   ClientWidth     =   9270
   ControlBox      =   0   'False
   Icon            =   "frmSysConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraProvider 
      Height          =   5505
      Left            =   2100
      TabIndex        =   62
      Top             =   30
      Visible         =   0   'False
      Width           =   7035
      Begin PT_XP_Button.PT_XP cmdClosTitle 
         Height          =   345
         Left            =   3825
         TabIndex        =   63
         Top             =   5010
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
         Picture         =   "frmSysConfig.frx":08CA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":12DC
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdProviders 
         Height          =   4335
         Left            =   150
         TabIndex        =   64
         Top             =   540
         Width           =   6705
         _Version        =   524288
         _ExtentX        =   11827
         _ExtentY        =   7646
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
         SpreadDesigner  =   "frmSysConfig.frx":162E
         UserResize      =   0
         Appearance      =   2
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin PT_XP_Button.PT_XP cmdRemoveTitle 
         Height          =   300
         Left            =   1050
         TabIndex        =   65
         Top             =   180
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
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
         Picture         =   "frmSysConfig.frx":1B04
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Remove"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":209E
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAddTitle 
         Height          =   300
         Left            =   150
         TabIndex        =   66
         Top             =   180
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
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
         Picture         =   "frmSysConfig.frx":23F0
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Add"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":298A
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSaveTitle 
         Height          =   345
         Left            =   2670
         TabIndex        =   67
         Top             =   5010
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
         Picture         =   "frmSysConfig.frx":2CDC
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":3076
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.Frame fraPrivStatement 
      Caption         =   "PHOTO PRIVACY DISCLAIMER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2112
      Left            =   1332
      TabIndex        =   9
      Top             =   1710
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox txtPrivStatement 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   4095
      End
      Begin PT_XP_Button.PT_XP cmdOk 
         Height          =   264
         Left            =   1176
         TabIndex        =   11
         Top             =   1740
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   476
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
         Picture         =   "frmSysConfig.frx":33C8
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":3762
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrivCancel 
         Height          =   264
         Left            =   2280
         TabIndex        =   12
         Top             =   1740
         Width           =   1068
         _ExtentX        =   1879
         _ExtentY        =   476
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
         Picture         =   "frmSysConfig.frx":3AB4
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Cancel"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":44C6
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.Frame fraImage 
      Height          =   1788
      Left            =   36
      TabIndex        =   0
      Top             =   36
      Visible         =   0   'False
      Width           =   5895
      Begin CIMAGEBUFFERXLib.CImageBufferX cimbx 
         Height          =   195
         Left            =   4620
         TabIndex        =   49
         Top             =   240
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   344
         _StockProps     =   0
      End
      Begin VB.Frame fraImageSettings 
         BorderStyle     =   0  'None
         Height          =   1284
         Left            =   60
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   5775
         Begin VB.TextBox txtImgPath 
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
            Height          =   330
            Left            =   60
            TabIndex        =   3
            Text            =   "\\"
            Top             =   828
            Width           =   5160
         End
         Begin PT_XP_Button.PT_XP cmdSelDevice 
            Height          =   384
            Left            =   60
            TabIndex        =   13
            ToolTipText     =   "Click to take picture"
            Top             =   60
            Width           =   1344
            _ExtentX        =   2381
            _ExtentY        =   688
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
            Picture         =   "frmSysConfig.frx":4818
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Select Device"
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSysConfig.frx":50F2
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdImgFolder 
            Height          =   228
            Left            =   5316
            TabIndex        =   14
            ToolTipText     =   "Save, Restore, Close"
            Top             =   828
            Width           =   408
            _ExtentX        =   714
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
            AppearanceThemes=   3
            BorderWidth     =   0
            TransparentColor=   14737632
            Caption         =   "..."
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSysConfig.frx":5D44
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdSetup 
            Height          =   348
            Left            =   1872
            TabIndex        =   15
            ToolTipText     =   "Save, Restore, Close"
            Top             =   60
            Width           =   1212
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
            AppearanceThemes=   3
            BorderWidth     =   0
            TransparentColor=   14737632
            Caption         =   "Setup"
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSysConfig.frx":6096
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdPrivStatement 
            Height          =   348
            Left            =   3396
            TabIndex        =   16
            ToolTipText     =   "Save, Restore, Close"
            Top             =   60
            Width           =   2328
            _ExtentX        =   4101
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
            BorderWidth     =   0
            TransparentColor=   14737632
            Caption         =   "Photo Privacy Disclaimer"
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSysConfig.frx":63E8
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Image Path:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   192
            Left            =   60
            TabIndex        =   4
            Top             =   552
            Width           =   1416
         End
      End
      Begin VB.CheckBox chkCapture 
         Caption         =   "ENABLE IMAGE CAPTURE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   2595
      End
   End
   Begin VB.Frame fraAlert 
      Height          =   2352
      Left            =   1188
      TabIndex        =   56
      Top             =   1440
      Visible         =   0   'False
      Width           =   4695
      Begin FPSpreadADO.fpSpread grdTerms 
         Height          =   1584
         Left            =   60
         TabIndex        =   57
         Top             =   432
         Width           =   4548
         _Version        =   524288
         _ExtentX        =   8022
         _ExtentY        =   2794
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
         SpreadDesigner  =   "frmSysConfig.frx":673A
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin PT_XP_Button.PT_XP cmdCloseList 
         Height          =   240
         Left            =   2400
         TabIndex        =   58
         Top             =   2064
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   423
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
         Picture         =   "frmSysConfig.frx":6C10
         AppearanceThemes=   3
         BorderDrawEvent =   1
         TransparentColor=   14737632
         Caption         =   "Close"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":7622
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdTermSave 
         Height          =   240
         Left            =   1416
         TabIndex        =   59
         Top             =   2064
         Width           =   948
         _ExtentX        =   1667
         _ExtentY        =   423
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
         Picture         =   "frmSysConfig.frx":7974
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":7D0E
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRemoveRow 
         Height          =   240
         Left            =   996
         TabIndex        =   60
         Top             =   156
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   423
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
         Picture         =   "frmSysConfig.frx":8060
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Remove"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":85FA
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdGridRow 
         Height          =   240
         Left            =   96
         TabIndex        =   61
         Top             =   156
         Width           =   876
         _ExtentX        =   1535
         _ExtentY        =   423
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
         Picture         =   "frmSysConfig.frx":894C
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Add"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":8EE6
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.TextBox cctCount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   54
      Top             =   5604
      Width           =   1215
   End
   Begin VB.Frame fraAppt 
      Height          =   5535
      Left            =   5940
      TabIndex        =   23
      Top             =   36
      Width           =   3285
      Begin VB.ComboBox cboSchedWeek 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmSysConfig.frx":9238
         Left            =   1380
         List            =   "frmSysConfig.frx":923A
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1290
         Width           =   1815
      End
      Begin VB.CheckBox chkRoomManagement 
         Caption         =   "ROOM MANAGEMENT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   150
         TabIndex        =   41
         Top             =   4290
         Width           =   2625
      End
      Begin VB.ComboBox cboCountry 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3720
         Width           =   2685
      End
      Begin VB.ComboBox cboState 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   3000
         Width           =   2685
      End
      Begin VB.CheckBox chkOverBook 
         Caption         =   "Multiple Patient Appt Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   90
         TabIndex        =   34
         Top             =   1800
         Width           =   2595
      End
      Begin VB.TextBox cctOverBook 
         Alignment       =   2  'Center
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
         Left            =   2130
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   2190
         Width           =   675
      End
      Begin VB.Frame fraEnableSched 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   60
         TabIndex        =   24
         Top             =   150
         Width           =   3225
         Begin VB.ComboBox cboEndHour 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmSysConfig.frx":923C
            Left            =   1320
            List            =   "frmSysConfig.frx":9246
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   360
            Width           =   1395
         End
         Begin VB.ComboBox cboStartHour 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmSysConfig.frx":9269
            Left            =   1320
            List            =   "frmSysConfig.frx":926B
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   0
            Width           =   1395
         End
         Begin VB.ComboBox cboApptInt 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmSysConfig.frx":926D
            Left            =   1320
            List            =   "frmSysConfig.frx":9280
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   740
            Width           =   1815
         End
         Begin VB.Label lblPM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PM"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2760
            TabIndex        =   32
            Top             =   435
            Width           =   255
         End
         Begin VB.Label lblAM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AM"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2760
            TabIndex        =   31
            Top             =   60
            Width           =   270
         End
         Begin VB.Label lblApptInterval 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Appt Interval"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   0
            TabIndex        =   30
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lblEHour 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ending Hour"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   0
            TabIndex        =   29
            Top             =   435
            Width           =   990
         End
         Begin VB.Label lblSHour 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Starting Hour"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   0
            TabIndex        =   28
            Top             =   60
            Width           =   1095
         End
      End
      Begin MSComCtl2.UpDown cnUpYears 
         Height          =   405
         Left            =   2850
         TabIndex        =   36
         Top             =   2190
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   714
         _Version        =   393216
         Alignment       =   0
         Enabled         =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdTerms 
         Height          =   495
         Left            =   1710
         TabIndex        =   50
         ToolTipText     =   "Add Terms"
         Top             =   4770
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmSysConfig.frx":92C9
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Patient Folders"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":9BA3
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAlert 
         Height          =   495
         Left            =   180
         TabIndex        =   55
         ToolTipText     =   "Add Terms"
         Top             =   4800
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmSysConfig.frx":A7F5
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Alert Folders"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSysConfig.frx":B0CF
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblScheduleWeek 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule Week"
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
         Left            =   90
         TabIndex        =   42
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label lblDefCountry 
         AutoSize        =   -1  'True
         Caption         =   "DEFAULT COUNTRY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   3450
         Width           =   1635
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         Caption         =   "DEFAULT STATE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   2730
         Width           =   1275
      End
      Begin VB.Label lblNoPatient 
         Caption         =   "Max Number of Patient Per time slot"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   120
         TabIndex        =   35
         Top             =   2190
         Width           =   1905
      End
   End
   Begin VB.Frame fraInfo 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   30
      TabIndex        =   6
      Top             =   72
      Width           =   5895
      Begin VB.ComboBox cboSchedType 
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
         ItemData        =   "frmSysConfig.frx":BD21
         Left            =   120
         List            =   "frmSysConfig.frx":BD2B
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   3690
         Width           =   2235
      End
      Begin VB.TextBox cctProviderScan 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1800
         TabIndex        =   48
         Text            =   "\\"
         Top             =   3030
         Width           =   3465
      End
      Begin VB.TextBox cctPatientScan 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1800
         TabIndex        =   46
         Text            =   "\\"
         Top             =   2610
         Width           =   3465
      End
      Begin VB.TextBox txtAccessDB 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1800
         TabIndex        =   45
         Text            =   "\\"
         Top             =   2190
         Width           =   3465
      End
      Begin VB.ComboBox cboSpecialty 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1740
         Width           =   3345
      End
      Begin VB.ComboBox cboDefaultProv 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1290
         Width           =   3345
      End
      Begin VB.TextBox txtDefaultLogo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   660
         Left            =   1590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   570
         Width           =   3690
      End
      Begin VB.ComboBox cboFacility 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   180
         Width           =   3375
      End
      Begin VB.Label lblScheduleType 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Schedule Type"
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
         TabIndex        =   69
         Top             =   3450
         Width           =   2175
      End
      Begin VB.Label lblProvScan 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Provider Scanned"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   3060
         Width           =   1365
      End
      Begin VB.Label lblScan 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Scanned"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label lblNetword 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Acess DB Location"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   2220
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "SPECIALTY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label lblDefaultProv 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "DEFAULT PROVIDER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "FACILITY LOGO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label lblFacility 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "DEFAULT FACILITY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1560
      End
   End
   Begin PT_XP_Button.PT_XP cmdDone 
      Height          =   324
      Left            =   4656
      TabIndex        =   51
      Top             =   5640
      Width           =   1092
      _ExtentX        =   1931
      _ExtentY        =   582
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
      Picture         =   "frmSysConfig.frx":BD4D
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSysConfig.frx":C75F
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   324
      Left            =   3492
      TabIndex        =   52
      Top             =   5640
      Width           =   1092
      _ExtentX        =   1931
      _ExtentY        =   582
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
      Picture         =   "frmSysConfig.frx":CAB1
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSysConfig.frx":CE4B
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdEmdeon 
      Height          =   384
      Left            =   7716
      TabIndex        =   53
      Top             =   5604
      Width           =   1392
      _ExtentX        =   2461
      _ExtentY        =   688
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSysConfig.frx":D19D
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "E-Prescribe Load"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSysConfig.frx":D917
      ShowFocus       =   -1  'True
   End
   Begin VB.Menu mnu_Enable_Photo 
      Caption         =   "Enable Photo"
   End
   Begin VB.Menu mnu_PhotoStatement 
      Caption         =   "Photo Privacy Statement"
   End
   Begin VB.Menu mnu_Provider_Title 
      Caption         =   "Provider Title"
   End
   Begin VB.Menu mnu_Magnifier 
      Caption         =   "Windows Magnifier"
      Begin VB.Menu mn_Sys_Normal 
         Caption         =   "Normal"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_Tablet 
         Caption         =   "125%"
      End
   End
End
Attribute VB_Name = "frmSysConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fPrivStatement As String
Private fFndPrivStatement As Boolean
Private Sub Add_Spacer(vObjGrid As Object, Optional vlngRow As Long)
Dim lngRow As Long

    With vObjGrid
        If vlngRow >= 0 Then
             vlngRow = vlngRow + 1
            .MaxRows = vlngRow
            lngRow = vlngRow
        Else
            lngRow = vlngRow
        End If
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
Private Function Format_Provider_Title()
Dim lngColWidth As Long
Dim varCellValue
Dim varHeader
Dim lngCnt  As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim strHeaders As String

    On Error GoTo Error_Format_Provider_Title
    Screen.MousePointer = vbHourglass
    With grdProviders
        .Redraw = False
        .ProcessTab = True
        .FontBold = True
        .FontName = "Arial"
        .RowHeadersShow = False
        .ColHeadersShow = True
        .MaxCols = 4
        .MaxRows = 0
        .ColHeaderRows = 4
        .ColHeadersAutoText = DispBlank
    End With
    grdProviders.ColWidth(1) = 12 * mlngSystemFileSize
    grdProviders.ColWidth(2) = 16.3 * mlngSystemFileSize
    grdProviders.ColWidth(3) = 12 * mlngSystemFileSize
    grdProviders.ColWidth(4) = 12 * mlngSystemFileSize
    lngCol = 1
    lngRow = SpreadHeader
    Format_CellTypeEdit grdProviders, lngRow, lngCol, "4", "PROVIDER TITLES", "TypeHAlignCenter", "", "", HeaderColor
    grdProviders.SetCellBorder lngCol, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer grdProviders, SpreadHeader + 1
    lngRow = SpreadHeader + 2
    lngCol = 1
    Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", "TITLE", "TypeHAlignCenter", "TITLE", "", HeaderColor
    grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 2
    Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", "DESCRIPTION", "TypeHAlignCenter", "DESCRIPTION", "", HeaderColor
    grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 3
    Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", "CPT GROUP", "TypeHAlignCenter", "CPT GROUP", "", HeaderColor
    grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 4
    Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", "REPORT GROUP", "TypeHAlignCenter", "REPORT GROUP", "", HeaderColor
    grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer grdProviders, SpreadHeader + 3
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Read_Provider_Title
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_Provider_Title:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_Provider_Title", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function
Private Function Get_Title_Group() As String
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim strTerm As String
Dim strValue As String

    On Error GoTo Error_Read_Folder
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    strSQL = "SELECT PROVIDER_GROUP FROM PROVIDER_TITLE ORDER BY PROVIDER_GROUP"
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        Do Until .EOF
            strTerm = vbNullString
            If Not IsNull((!PROVIDER_GROUP)) Then strTerm = (!PROVIDER_GROUP)
            .MoveNext
            If strTerm = vbNullString Then
                Get_Title_Group = strTerm
            Else
                Get_Title_Group = Get_Title_Group & Chr(9) & strTerm
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
Error_Read_Folder:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Folder", App.EXEName & " " & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Read_Provider_Title()
Dim rsProvider As New ADODB.Recordset
Dim cnProvider As New ADODB.Connection
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim lngProv As Long
Dim Idx%
Dim lngFac As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strTitle As String
Dim strDesc As String
Dim strGroup As String
Dim strReptGroup As String


    On Error GoTo Error_Read_Provider_Title
    grdProviders.MaxRows = 0
    If cnProvider.State = adStateClosed Then cnProvider.open psConnect(1)
    strSQL = "SELECT * FROM PROVIDER_TITLE ORDER BY PROVIDER_TITLE"
    Set rsProvider = cnProvider.Execute(strSQL)
    With rsProvider
        Do Until .EOF
           strTitle = vbNullString
           strDesc = vbNullString
           strGroup = vbNullString
           strReptGroup = vbNullString
           lngProv = 0
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PROVIDER_TITLE"
                                strTitle = Trim(Fld.Value)
                            Case "PROVIDER_DESC"
                                strDesc = Trim(Fld.Value)
                            Case "PROVIDER_GROUP"
                                strGroup = Trim(Fld.Value)
                            Case "REPORT_GROUP"
                                strReptGroup = Trim(Fld.Value)
                            Case "TITLE_ID"
                                lngProv = Trim(Fld.Value)
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If Trim(strTitle) <> vbNullString Then
                lngRow = lngRow + 1
                grdProviders.MaxRows = lngRow
                grdProviders.SetRowItemData lngRow, lngProv
                lngCol = 1
                Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", strTitle, "TypeHAlignCenter", "TITLE"
                grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", strDesc, "TypeHAlignCenter", "DESCRIPTION"
                grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 3
                Format_CellTypeCombo grdProviders, lngRow, lngCol, "1", Get_Title_Group, "TypeHAlignCenter", "CPT GROUP"
                Set_GridCurSel grdProviders, strGroup, lngCol, lngRow, "1", "CPT GROUP"
                'Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", "FEE GROUP", "TypeHAlignCenter", "CPT GROUP"
                grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 4
                Format_CellTypeCombo grdProviders, lngRow, lngCol, "1", Get_Title_Group, "TypeHAlignCenter", "REPORT GROUP"
                Set_GridCurSel grdProviders, strReptGroup, lngCol, lngRow, "1", "CPT GROUP"
                'Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", "REPORT GROUP", "TypeHAlignCenter", "REPORT GROUP"
                grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        Loop
    End With
    If cnProvider.State = adStateOpen Then
        cnProvider.Close
        Set cnProvider = Nothing
        Set rsProvider = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Provider_Title:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Provider_Title", App.EXEName & "." & "ModConst"
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
                If Trim(vstrData) = Trim(.TypeComboBoxString) Then
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
        .TypeVAlign = TypeVAlignCenter
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
vstrString As String, vstrHAligh As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, _
Optional vstrCellNote As String, Optional vstrHAlign As String)
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
            .TypeVAlign = TypeVAlignCenter
              If Trim(vstrHAlign) = vbNullString Then
                .TypeHAlign = TypeHAlignCenter
              Else
                .TypeHAlign = vstrHAlign
              End If
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
          .BlockMode = False
    End With
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
Private Function Format_CellTypePicture(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long, Optional vstrCellNote As String)
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
            If fso.FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
            End If
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(lngSpan) Then
                lngSpan = vstrSpan
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
             End If
             If Trim(vstrCellNote) <> vbNullString Then .CellNote = vstrCellNote
            .TypePictMaintainScale = True
           ' .TypePictStretch = True
            .TypePictCenter = True
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            'grdPic.SetCellBorder vlngCol, vlngRow, vlngCol, vlngRow, CellBorderIndexOutline, &H800000, CellBorderStyleSolid
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
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
            If lngBackColor > 0 Then .BackColor = lngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        .BlockMode = False
    End With
End Function

Private Sub Format_Grid()
Dim lngCol As Long
Dim lngRow As Long

    Screen.MousePointer = vbHourglass
    grdTerms.ProcessTab = True
    grdTerms.FontBold = True
    With grdTerms
        .Redraw = False
        grdTerms.ClearRange 1, 1, grdTerms.MaxCols, grdTerms.MaxRows, False
        .MaxCols = 5
        .MaxRows = 0
        .ColHeaderRows = 2
        .ColWidth(1) = 20
        .ColWidth(2) = 10
        lngCol = 1
        Format_CellTypeEdit grdTerms, SpreadHeader, lngCol, "1", "FOLDER NAME", "TypeHAlignCenter", "FOLDER_NAME", "", HeaderColor
        grdTerms.SetCellBorder lngCol, SpreadHeader, lngCol, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdTerms, SpreadHeader, lngCol, "1", "ALERT", "TypeHAlignCenter", "ALERT_DESC", "", HeaderColor
        grdTerms.SetCellBorder lngCol, SpreadHeader, lngCol, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Add_Spacer grdTerms, SpreadHeader + 1
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Function Load_Appt()
Dim Idx As Integer
Dim strOption As String

    On Error GoTo Error
    cboStartHour.Clear
    cboStartHour.AddItem "12:00"
    cboStartHour.AddItem "01:00"
    cboStartHour.AddItem "02:00"
    cboStartHour.AddItem "03:00"
    cboStartHour.AddItem "04:00"
    cboStartHour.AddItem "05:00"
    cboStartHour.AddItem "06:00"
    cboStartHour.AddItem "07:00"
    cboStartHour.AddItem "08:00"
    cboStartHour.AddItem "09:00"
    cboStartHour.AddItem "10:00"
    cboStartHour.AddItem "11:00"
    cboStartHour.ListIndex = 8
    cboEndHour.Clear
    cboEndHour.AddItem "12:00"
    cboEndHour.AddItem "01:00"
    cboEndHour.AddItem "02:00"
    cboEndHour.AddItem "03:00"
    cboEndHour.AddItem "04:00"
    cboEndHour.AddItem "05:00"
    cboEndHour.AddItem "06:00"
    cboEndHour.AddItem "07:00"
    cboEndHour.AddItem "08:00"
    cboEndHour.AddItem "09:00"
    cboEndHour.AddItem "10:00"
    cboEndHour.AddItem "11:00"
    cboEndHour.ListIndex = 8
    Read_Sched_Options
    If clsApptPref.Appt_Start = vbNullString Then
        clsApptPref.Appt_Start = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScheduleStartHour")
        Write_CustomForm "START_TIME", Trim(clsApptPref.Appt_Start)
    End If
    For Idx = 0 To cboStartHour.listcount - 1
        If UCase(cboStartHour.List(Idx)) = UCase(clsApptPref.Appt_Start) Then
           cboStartHour.ListIndex = Idx
           Exit For
        End If
    Next Idx
    If clsApptPref.Appt_End = vbNullString Then
        clsApptPref.Appt_End = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScheduleEndHour")
        Write_CustomForm "END_TIME", Trim(clsApptPref.Appt_End)
    End If
    For Idx = 0 To cboEndHour.listcount - 1
        If UCase(cboEndHour.List(Idx)) = UCase(clsApptPref.Appt_End) Then
           cboEndHour.ListIndex = Idx
           Exit For
        End If
    Next Idx
    If clsApptPref.SchedInterval = vbNullString Then
        clsApptPref.SchedInterval = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "SchedulingInterval")
        Write_CustomForm "APPT_INTERVAL", Trim(clsApptPref.SchedInterval)
    End If
    For Idx = 0 To cboApptInt.listcount - 1
        If TextFound(Trim(clsApptPref.SchedInterval), Trim(cboApptInt.List(Idx)), Len(Trim(clsApptPref.SchedInterval)), True) Then
           cboApptInt.ListIndex = Idx
           Exit For
        End If
    Next Idx
    cboSchedWeek.Clear
    cboSchedWeek.AddItem "Mon - Friday"
    cboSchedWeek.ItemData(cboSchedWeek.NewIndex) = 5
    cboSchedWeek.AddItem "Mon - Saturday"
    cboSchedWeek.ItemData(cboSchedWeek.NewIndex) = 6
    cboSchedWeek.AddItem "Mon - Sunday"
    cboSchedWeek.ItemData(cboSchedWeek.NewIndex) = 7
    If clsApptPref.SchedWeek = vbNullString Then
        clsApptPref.SchedWeek = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScheduleWeek")
        Write_CustomForm "APPT_WEEK", Trim(clsApptPref.SchedWeek)
    Else
        Write_CustomForm "APPT_WEEK", Trim(clsApptPref.SchedWeek)
    End If
    For Idx = 0 To cboSchedWeek.listcount - 1
        If Trim(cboSchedWeek.List(Idx)) = Trim(clsApptPref.SchedWeek) Then
           cboSchedWeek.Text = cboSchedWeek.List(Idx)
           Exit For
        End If
    Next Idx
    If clsApptPref.OverbookNumber = 0 Then
        clsApptPref.OverbookNumber = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "MaxBookPatient")
        Write_CustomForm "OVERBOOK", CStr(clsApptPref.OverbookNumber)
    Else
        cctOverBook.Text = CStr(clsApptPref.OverbookNumber)
    End If
    If clsApptPref.PatientType <> vbNullString Then
        For Idx = 0 To cboSchedType.listcount - 1
            If Trim(cboSchedType.List(Idx)) = Trim(clsApptPref.PatientType) Then
               cboSchedType.Text = cboSchedType.List(Idx)
               Exit For
            End If
        Next Idx
    End If
    If Val(clsApptPref.OverbookNumber) <= 1 Then
        clsApptPref.OverBook = False
    ElseIf Val(clsApptPref.OverbookNumber) > 1 Then
        clsApptPref.OverBook = True
    End If
    If clsApptPref.OverbookNumber > 1 Then
        chkOverBook.Value = 1
        lblNoPatient.Visible = True
        cnUpYears.Visible = True
        cctOverBook.Visible = True
        cctOverBook.Text = CStr(clsApptPref.OverbookNumber)
    End If
    Exit Function
Error:
    Resume Next
    Exit Function
End Function
Private Function Write_CustomForm(vstrControl As String, vstrType As String)
Dim strSQL As String
Dim strValue As String
Dim i%
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset

    On Error GoTo Error_Write_CustomForm
    Screen.MousePointer = vbHourglass
    clsGridFormat.FormName = "APPOINTMENTS"
    clsGridFormat.RowName = vstrControl
    clsGridFormat.Text = vstrType
    clsGridFormat.GridReportID = 0
    If cnFac.State = adStateClosed Then cnFac.open psConnect(1)
    strSQL = "SELECT FORM_ID FROM CUSTOM_PRACTICE_FORM WHERE FORM_NAME = '" & clsGridFormat.FormName & "' AND ROW_NAME = '" & clsGridFormat.RowName & "'"
    Set rsFac = cnFac.Execute(strSQL)
    With rsFac
        If Not .EOF Then
            clsGridFormat.GridReportID = (!FORM_ID)
        End If
    End With
    If clsGridFormat.GridReportID = 0 Then
        strSQL = "INSERT INTO CUSTOM_PRACTICE_FORM (FORM_NAME, ROW_NAME, RESPONSE_TYPE) " _
            & "values('" & clsGridFormat.FormName & "','" & clsGridFormat.RowName & "', '" _
            & Trim(clsGridFormat.Text) & "')"
        Set rsFac = cnFac.Execute(strSQL)
    ElseIf Trim(clsGridFormat.RowName) <> vbNullString Then
        strSQL = "UPDATE CUSTOM_PRACTICE_FORM SET  RESPONSE_TYPE = '" & Trim(clsGridFormat.Text) & "' WHERE  FORM_ID = " & clsGridFormat.GridReportID
        Set rsFac = cnFac.Execute(strSQL)
    End If
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set rsFac = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_CustomForm:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Write_CustomForm", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Read_Folder()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim strTerm As String
Dim strValue As String

    On Error GoTo Error_Read_Folder
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    strSQL = "SELECT DISTINCT ATTRBTE_VALID_VALUE, VALID_VALUE_DEF FROM ATTR_VALID_VALUE " _
        & "WHERE VALID_VALUE_SUBDOM = 'sd_Patient_ScannedFolders' ORDER BY APPT_CODE"
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        Do Until .EOF
            strTerm = vbNullString
            strValue = vbNullString
            If Not IsNull((!VALID_VALUE_DEF)) Then strTerm = (!ATTRBTE_VALID_VALUE)
            If Not IsNull((!VALID_VALUE_DEF)) Then strValue = (!VALID_VALUE_DEF)
            .MoveNext
            If Trim(strTerm) <> vbNullString Then
                lngRow = grdTerms.MaxRows + 1
                grdTerms.MaxRows = lngRow
                lngCol = 1
                Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", strTerm, "TypeHAlignCenter", "ATTRBTE_VALID_VALUE"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", "True" & Chr(9) & "False", "TypeHAlignCenter", "VALID_VALUE_DEF"
                Set_GridCurSel grdTerms, strValue, lngCol, lngRow, "1", "VALID_VALUE_DEF", 0, "", "TypeHAlignCenter"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
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
Error_Read_Folder:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Folder", App.EXEName & " " & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Load_Providers()
Dim strSQL As String
Dim iRows As Long
Dim i As Integer
Dim lstrRemarks As String
Dim sName As String
Dim cnProv As New ADODB.Connection
Dim rsProv As New ADODB.Recordset
Dim strValue As String
    
    If cnProv.State = adStateClosed Then cnProv.open psConnect(0)
    strSQL = "Select LAST_NAME, FIRST_NAME, MIDDLE_NAME, SUFFIX, TITLE, PROVIDER_ID " _
            & "From PROVIDERS  " _
            & "Where Title = 'M.D.' or Title = 'MD' or Title = 'RN' or " _
                & "Title = 'LPN' or Title = 'PA' Or Title = 'APRN' or Title = 'NP' or Title = 'MA' or Title = 'MS' "
    Set rsProv = cnProv.Execute(strSQL)
    With rsProv
        Do Until .EOF
            If Not IsNull(!First_Name) Then sName = (!First_Name)
            If Not IsNull(!Middle_Name) Then sName = sName & " " & (!Middle_Name)
            If Not IsNull(!Last_Name) Then sName = sName & " " & (!Last_Name)
            If Trim((!Suffix)) <> vbNullString Or Trim((!Suffix)) <> vbNullString Then sName = sName & ", " & (!Suffix)
            If Not IsNull(!Title) Then sName = sName & ", " & (!Title)
            cboDefaultProv.AddItem sName
            cboDefaultProv.ItemData(cboDefaultProv.NewIndex) = (!Provider_ID)
        .MoveNext
        Loop
    End With
   If cnProv.State = adStateOpen Then
        cnProv.Close
        Set cnProv = Nothing
        Set rsProv = Nothing
   End If
    strValue = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultProvider")
    If strValue <> vbNullString Then
        For i = 0 To cboDefaultProv.listcount - 1
            If cboDefaultProv.ItemData(i) = Val(strValue) Then
                 cboDefaultProv.Text = cboDefaultProv.List(i)
                 Exit For
            End If
        Next i
    End If
   Exit Function
End Function

Private Function Load_Specialty()
Dim strValue As String
Dim i%

    strValue = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "PracticeType")
    LoadAttrValidValue psConnect(1), cboSpecialty, "sd_Provider_Specialty"
    If Trim(strValue) <> vbNullString Then
       For i = 0 To cboSpecialty.listcount - 1
          If cboSpecialty.List(i) = strValue Then
                cboSpecialty.Text = cboSpecialty.List(i)
          End If
       Next i
    End If
End Function

Private Function Load_State()
Dim strSQL As String
Dim strState As String
Dim strCountry As String
Dim i%


    strSQL = "Select DISTINCT state_name from States where state_ind = 'y' order by state_name"
    LoadItems cboState, strSQL, psConnect(1)
    
    strSQL = "Select DISTINCT country_name from Countries order by country_name"
    LoadItems cboCountry, strSQL, psConnect(1)
    cboCountry.Text = "United States"
    
    strState = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultState")
    If strState = vbNullString Then
      SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultState", "", REG_SZ
      strState = ""
    End If
    strCountry = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultCountry")
    If strCountry = vbNullString Then
      SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultCountry", "United States", REG_SZ
      strCountry = "United States"
    End If
    
    For i = 0 To cboState.listcount - 1
        If cboState.List(i) = strState Then
             cboState.Text = cboState.List(i)
        End If
    Next i
    For i = 0 To cboCountry.listcount - 1
        If cboCountry.List(i) = strCountry Then
             cboCountry.Text = cboCountry.List(i)
        End If
    Next i
    
    Exit Function
End Function


Private Function Update_DefFacility(vlngFacilityID As Long)
Dim i%
Dim strSQL As String
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset

    If vlngFacilityID = 0 Then Exit Function
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "UPDATE FACILITIES SET DEFAULT_PRAC = 'TRUE' WHERE FACILITY_ID = " & vlngFacilityID
    Set rsFac = cnFac.Execute(strSQL)
    If Trim(txtDefaultLogo.Text) <> vbNullString Then
        strSQL = "UPDATE FACILITIES SET Logo = '" & Trim(txtDefaultLogo.Text) & "' WHERE FACILITY_ID = " & vlngFacilityID
        Set rsFac = cnFac.Execute(strSQL)
    End If
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set cnFac = Nothing
        Set rsFac = Nothing
    End If
    Exit Function
End Function

Private Function Write_Title()
Dim rsProvider As New ADODB.Recordset
Dim cnProvider As New ADODB.Connection
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim lngProv As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strTitle As String
Dim strDesc As String
Dim strGroup As String
Dim strReptGroup As String
Dim varCellValue

    On Error GoTo Error_Write_Title
    Screen.MousePointer = vbHourglass
    If cnProvider.State = adStateClosed Then cnProvider.open psConnect(1)
    With grdProviders
        For lngRow = 1 To grdTerms.MaxRows
            lngProv = 0
            .GetText 1, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                strTitle = Trim(varCellValue)
            End If
            .GetText 2, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                strDesc = Trim(varCellValue)
            End If
            .GetText 3, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                strGroup = Trim(varCellValue)
            End If
            .GetText 4, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                strReptGroup = Trim(varCellValue)
            End If
            lngProv = .GetRowItemData(lngRow)
            strTitle = vbNullString
            strDesc = vbNullString
            strGroup = vbNullString
            strReptGroup = vbNullString
            If Trim(strTitle) <> vbNullString Then
                If lngProv = 0 Then
                    strSQL = "SELECT TITLE_ID FROM PROVIDER_TITLE " _
                        & "WHERE PROVIDER_TITLE = '" & strTitle & "'  "
                    Set rsProvider = cnProvider.Execute(strSQL)
                    With rsProvider
                        If Not .EOF Then
                            lngProv = (!TITLE_ID)
                        End If
                    End With
                End If
                If lngProv > 0 Then
                    strSQL = "UPDATE PROVIDER_TITLE SET " _
                        & "PROVIDER_DESC = '" & strDesc & "'," _
                        & "PROVIDER_GROUP = '" & strGroup & "'," _
                        & "REPORT_GROUP = '" & strGroup & "' WHERE TITLE_ID = " & lngProv
                    Set rsProvider = cnProvider.Execute(strSQL)
                Else
                    strSQL = "INSERT INTO PROVIDER_TITLE (PROVIDER_TITLE," _
                        & "PROVIDER_DESC,PROVIDER_GROUP,REPORT_GROUP) values " _
                        & "('" & strTitle & "','" & strDesc & "','" _
                        & strGroup & "','" & strReptGroup & "', )  "
                    Set rsProvider = cnProvider.Execute(strSQL)
                End If
            End If
        Next lngRow
    End With
    MsgBox "Provider Title Groups have been successfully Updated!", vbInformation, "PatientTrac EMR Alerts"
    If cnProvider.State = adStateOpen Then
        cnProvider.Close
        Set cnProvider = Nothing
        Set rsProvider = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Title:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Title", App.EXEName & " " & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cboCountry_KeyUp(KeyCode As Integer, Shift As Integer)
    SearchList cboCountry, cboCountry.Text, True
End Sub


Private Sub cboDefaultProv_KeyUp(KeyCode As Integer, Shift As Integer)
    SearchList cboDefaultProv, cboDefaultProv.Text, True
End Sub


Private Sub cboFacility_Click()
Dim i%

    For i = 0 To cboFacility.listcount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            piFacilityID = cboFacility.ItemData(i)
        End If
    Next i
End Sub


Private Sub cboFacility_KeyUp(KeyCode As Integer, Shift As Integer)
    SearchList cboFacility, cboFacility.Text, True
End Sub


Private Sub cboSpecialty_KeyUp(KeyCode As Integer, Shift As Integer)
    SearchList cboSpecialty, cboSpecialty.Text, True
End Sub


Private Sub cboState_KeyUp(KeyCode As Integer, Shift As Integer)
    SearchList cboState, cboState.Text, True
End Sub


Private Sub chkCapture_Click()
    If chkCapture.Value Then fraImageSettings.Visible = True Else fraImageSettings.Visible = False
End Sub

Private Sub cmdAddTitle_Click()
Dim lngCol As Long
Dim lngRow As Long

    With grdProviders
        lngRow = lngRow + 1
        .MaxRows = lngRow
        lngCol = 1
        lngRow = SpreadHeader + 2
        Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "TITLE"
        grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "DESCRIPTION"
        grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeCombo grdProviders, lngRow, lngCol, "1", Get_Title_Group, "TypeHAlignCenter", "CPT GROUP"
        grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeCombo grdProviders, lngRow, lngCol, "1", Get_Title_Group, "TypeHAlignCenter", "REPORT GROUP"
        grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    End With
End Sub

Private Sub cmdClosTitle_Click()
    fraProvider.Visible = False
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub



Private Sub cmdEmdeon_Click()
Dim cnPatient As New ADODB.Connection
Dim rsPatient As New ADODB.Recordset
Dim strSQL As String
Dim lngPatientID As Long

    If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
    Screen.MousePointer = vbHourglass
    mlngCnt = 0
    cctCount.Visible = True
    strSQL = "SELECT EMDEON_USER_NAME, EMDEON_PASSWORD FROM PROVIDERS WHERE EMDEON_USER_NAME IS NOT NULL"
    Set rsPatient = cnPatient.Execute(strSQL)
    If Not rsPatient.EOF Then
        clsProvider.Emdeon_Password = rsPatient.Fields("EMDEON_PASSWORD")
        clsProvider.Emdeon_User = rsPatient.Fields("EMDEON_USER_NAME")
    End If
    If Trim(clsProvider.Emdeon_Password) <> vbNullString Then
        If PsUserID = vbNullString Then PsUserID = Get_PsUserID(psConnect(0))
        If Left(PsUserID, 7) = "Support" Or Left(PsUserID, 7) = "PTSuppo" Then
            If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
            strSQL = "SELECT PATIENT_ID FROM PATIENT WHERE PURGE = 0"
            Set rsPatient = cnPatient.Execute(strSQL)
            With rsPatient
                Do Until .EOF
                    lngPatientID = rsPatient.Fields("PATIENT_ID")
                    .MoveNext
                    If lngPatientID > 0 Then
                        Upload_PatientData_Routine lngPatientID
                        mlngCnt = mlngCnt + 1
                        cctCount.Text = CStr(mlngCnt)
                    End If
                Loop
            End With
        Else
            MsgBox "Access Denied!", vbInformation, "PatientTrac E-Prescribe"
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cmdGridRow_Click()
Dim lngCol As Long
Dim lngRow As Long

    lngCol = 1
    Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "", "TypeHAlignCenter", "ATTRBTE_VALID_VALUE NAME"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 2
    Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", "True" & Chr(9) & "False", "TypeHAlignCenter", "VALID_VALUE_DEF"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Exit Sub
End Sub

Private Sub cmdImgFolder_Click()
   Dim lstrPath As String

   lstrPath = ChooseFolder("Set File Path for Images")
   If Trim(lstrPath) <> vbNullString Then txtImgPath.Text = lstrPath
End Sub

Private Sub cmdOK_Click()
    ObjEnabled True
    fraPrivStatement.Visible = False
End Sub

Private Sub cmdPrivCancel_Click()
    txtPrivStatement.Text = fPrivStatement
    ObjEnabled True
    fraPrivStatement.Visible = False
End Sub

Private Sub cmdPrivStatement_Click()
    On Error GoTo Error_Handler
    
    fPrivStatement = txtPrivStatement.Text
    ObjEnabled False
    fraPrivStatement.Visible = True
    txtPrivStatement.SetFocus
   
   Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmSysConfig.cmdPrivStatement_Click", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume Next
End Sub

Private Sub cmdRemoveTitle_Click()
Dim lngRow As Long
Dim bFound As Boolean

    With grdProviders
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                bFound = True
                .DeleteRows lngRow, 1
                .MaxRows = .MaxRows - 1
                Exit For
            End If
        Next lngRow
    End With
    If Not bFound Then
        MsgBox "Please select a row first!", vbInformation, "PatientTrac Provider Titles"
    End If
    Exit Sub
End Sub

Private Sub cmdSave_Click()
Dim lngFacID As Long
Dim lngProvID As Long
Dim i%
Dim fso As New FileSystemObject
Dim strShed As String

    On Error GoTo Error_Handler
    If cboFacility.ListIndex = -1 Then
       MsgBox "Please select a facility", vbExclamation, App.Title
       Exit Sub
    Else
        For i = 0 To cboFacility.listcount - 1
            If cboFacility.List(i) = cboFacility.Text Then
                lngFacID = cboFacility.ItemData(i)
                Exit For
            End If
        Next i
    End If
    If cboDefaultProv.ListIndex = -1 Then
       MsgBox "Please select a provider", vbExclamation, App.Title
       Exit Sub
    Else
        For i = 0 To cboDefaultProv.listcount - 1
            If cboDefaultProv.List(i) = cboDefaultProv.Text Then
                lngProvID = cboDefaultProv.ItemData(i)
                Exit For
            End If
        Next i
    End If
'    If cboAgeDate.ListIndex = -1 Then
'    Else
'        For i = 0 To cboAgeDate.listcount - 1
'            If cboAgeDate.List(i) = cboAgeDate.Text Then
'                strShed = cboAgeDate.Text
'                Exit For
'            End If
'        Next i
'    End If
    If strShed <> vbNullString Then
        Write_CustomForm "PATIENTTRAC_AGINGDATE", strShed
    End If
    If chkCapture.Value = 1 Then
       If Trim(txtImgPath.Text) = vbNullString Then
          MsgBox "Image Path Required", vbExclamation, App.Title
          Exit Sub
       End If
       If Not fso.FolderExists(Trim(txtImgPath.Text)) Then
          MsgBox "Image Path Does Not Exists", vbExclamation, App.Title
          Exit Sub
       End If
       SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ImagePath", txtImgPath, REG_SZ
       psImageFilePath = Trim(txtImgPath.Text)
       If Mid$(psImageFilePath, Len(psImageFilePath), 1) <> "\" Then psImageFilePath = psImageFilePath & "\"
       pbCaptureSetting = True
       SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "CaptureSetting", "True", REG_SZ
    Else
       pbCaptureSetting = False
       SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "CaptureSetting", "False", REG_SZ
    End If
    Save_PrivacyStatement
    SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "FacilityID", lngFacID, REG_SZ
    SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultFacility", lngFacID, REG_SZ
    If Trim(txtDefaultLogo.Text) <> vbNullString Then
        Update_DefFacility Val(lngFacID)
    End If
    If cboSchedType = vbNullString Then
        clsApptPref.PatientType = "Full Patient Info"
    End If
    If cboSchedType <> vbNullString Then
        clsApptPref.PatientType = Trim(cboSchedType)
        Write_CustomForm "PATIENTTRAC_SCHED_TYPE", clsApptPref.PatientType
    End If
    If cboState <> vbNullString Then
        clsApptPref.Default_State = Trim(cboState)
        Write_CustomForm "DEFAULT_STATE", clsApptPref.Default_State
    End If
    If cboStartHour.Text <> vbNullString Then
        clsApptPref.Appt_Start = Trim(cboStartHour.Text)
        Write_CustomForm "START_TIME", Trim(clsApptPref.Appt_Start)
    End If
    If cboEndHour.Text <> vbNullString Then
        clsApptPref.Appt_End = Trim(cboEndHour.Text)
        Write_CustomForm "END_TIME", Trim(clsApptPref.Appt_End)
    End If
    If cboApptInt.Text <> vbNullString Then
        clsApptPref.SchedInterval = Trim(cboApptInt.Text)
        Write_CustomForm "APPT_INTERVAL", Trim(clsApptPref.SchedInterval)
    End If
    If cboSchedWeek.Text <> vbNullString Then
        clsApptPref.SchedWeek = Trim(cboSchedWeek.Text)
        Write_CustomForm "APPT_WEEK", Trim(clsApptPref.SchedWeek)
    End If
    If cctOverBook.Text <> vbNullString Then
        clsApptPref.OverbookNumber = Val(cctOverBook.Text)
        Write_CustomForm "OVERBOOK", Trim(clsApptPref.OverbookNumber)
    End If
    SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultCountry", cboCountry.Text, REG_SZ
    SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultProvider", lngProvID, REG_SZ
    SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "PracticeType", cboSpecialty.Text, REG_SZ
    Unload Me
   Exit Sub
Error_Handler:
    MsgBox Err.Description, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume Next
End Sub

Private Sub cmdSaveTitle_Click()

    Write_Title
    fraProvider.Visible = True
    Exit Sub
End Sub

Private Sub cmdSelDevice_Click()
    Select_Video cimbx
End Sub
Private Sub cmdSetup_Click()
    Setup_Video cimbx
End Sub

Private Sub cmdTerms_Click()
    Patient_Folders
    frmNurseTemplates.fCaption = "Patient Folders"
    frmNurseTemplates.fType = "Patient Folders"
    frmNurseTemplates.fSubDomain = "sd_Patient_ScannedFolders"
    frmNurseTemplates.fConnectCR = psConnect(1)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Get_Terms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Terms", App.EXEName & "." & TypeName(Me)
    Exit Sub
    Resume Next
End Sub
Private Function Patient_Folders()
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim lngRow As Long
Dim strValue As String

    On Error GoTo Error_Patient_Folders
    Screen.MousePointer = vbHourglass
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "SELECT DISTINCT ATTRBTE_VALID_VALUE, valid_value_seq_nbr " _
        & "From ATTR_VALID_VALUE " _
        & "Where VALID_VALUE_SUBDOM = 'sd_Patient_ScannedFolders' ORDER BY ATTRBTE_VALID_VALUE"
    Set rsTerms = cnTerms.Execute(strSQL)
    With rsTerms
        If Not .EOF Then
            If Trim((!ATTRBTE_VALID_VALUE)) <> vbNullString Then
                strValue = (!ATTRBTE_VALID_VALUE)
            End If
        End If
    End With
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    If Trim(strValue) = vbNullString Then
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Correspondence', 1, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('H & P', 2, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Insurance', 3, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Laboratory', 4, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Radiology', 5, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Patient Medical Records', 6, 'sd_Patient_ScannedFolders')"
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Referral Physician Correspondence', 7, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Patient Consents', 8, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Record Request', 9, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Third Party', 10, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Medication', 11, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Legal Correspondence', 12, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
        strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR, VALID_VALUE_SUBDOM) " _
            & "values('Progress Notes', 13, 'sd_Patient_ScannedFolders')"
            Set rsTerms = cnTerms.Execute(strSQL)
    End If
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set rsTerms = Nothing
        Set cnTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Patient_Folders:
   Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Patient_Folders", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Function Write_Folder()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim strTerm As String
Dim strValue As String
Dim varCellValue

    On Error GoTo Error_Write_Folder
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    For lngRow = 1 To grdTerms.MaxRows
        grdTerms.GetText 1, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            clsFolder.Folder_Type = Trim(varCellValue)
        End If
        grdTerms.GetText 2, lngRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            clsFolder.Folder_Show = Trim(varCellValue)
        End If
        If Trim(clsFolder.Folder_Type) <> vbNullString Then
            strSQL = "SELECT ATTR_ID FROM ATTR_VALID_VALUE " _
                & "WHERE VALID_VALUE_SUBDOM = 'sd_Patient_ScannedFolders' AND " _
                & "ATTRBTE_VALID_VALUE = '" & clsFolder.Folder_Type & "'"
            Set rsCodes = cnCodes.Execute(strSQL)
            With rsCodes
                If Not .EOF Then
                    clsFolder.Alert_ID = (!ATTR_ID)
                End If
            End With
            If clsFolder.Alert_ID > 0 Then
                strSQL = "UPDATE ATTR_VALID_VALUE SET " _
                    & "VALID_VALUE_DEF = '" & clsFolder.Folder_Show & "'  WHERE " _
                    & "ATTR_ID = " & clsFolder.Alert_ID
                Set rsCodes = cnCodes.Execute(strSQL)
            Else
                strSQL = "INSERT INTO ATTR_VALID_VALUE (ATTRBTE_VALID_VALUE, VALID_VALUE_DEF) values " _
                    & "('" & clsFolder.Folder_Type & "', '" & clsFolder.Folder_Show & "')  "
                Set rsCodes = cnCodes.Execute(strSQL)
            End If
        End If
    Next lngRow
    MsgBox "Folder Alerts have been successfully Updated!", vbInformation, "PatientTrac EMR Alerts"
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Folder:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Folder", App.EXEName & " " & Me.Name
    Exit Function
    Resume Next
End Function
Private Sub cnUpYears_DownClick()
    If Val(cctOverBook.Text) >= 1 Then cctOverBook.Text = Val(cctOverBook.Text) - 1
    Exit Sub
End Sub


Private Sub cnUpYears_UpClick()
    If cnUpYears.Increment = 1 Then cctOverBook.Text = Val(cctOverBook.Text) + 1
    Exit Sub
End Sub


Private Sub Form_Load()
    On Error GoTo Error_Handler
    chkCapture.Value = 0
    If pbCaptureSetting Then chkCapture.Value = 1
    chkRoomManagement.Value = 0
    If pbRoomMan Then chkRoomManagement.Value = 1
    txtImgPath.Text = psImageFilePath
    Get_PrivacyStatement
    Load_Appt
    Load_Facilities
    Load_Providers
    Load_Specialty
    Load_State
    centerForm Me
   Exit Sub
Error_Handler:
     MsgBox Err.Description, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume Next
End Sub

Public Sub ObjEnabled(pEnabled As Boolean)
    fraImage.Enabled = pEnabled
    fraInfo.Enabled = pEnabled
    cmdSave.Enabled = pEnabled
    cmdDone.Enabled = pEnabled
End Sub

Private Sub Get_PrivacyStatement()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    On Error GoTo Error_Get_PrivacyStatement
  
    cn.open psConnect(1)
   
    'Call stored procedure to retrieve data values for combo
    cn.S_ATTR_VALID_VALUE "photo_privacy_stmnt", rs
    
    fFndPrivStatement = False
    
    With rs
       If Not .EOF Then
          fFndPrivStatement = True
          txtPrivStatement.Text = (!ATTRBTE_VALID_VALUE) & ""
       End If
       .Close
    End With
    
    Set rs = Nothing
    
    cn.Close
    Set cn = Nothing


    Exit Sub
    
Error_Get_PrivacyStatement:
    MsgBox Err.Description, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume Next
End Sub
Private Sub Save_PrivacyStatement()
Dim cnPriv As New ADODB.Connection
    
    On Error GoTo Error_Save_PrivacyStatement
    If cnPriv.State = adStateClosed Then cnPriv.open psConnect(1)
    If Not fFndPrivStatement Then
       cnPriv.i_attr_valid_value "photo_privacy_stmnt", 1, txtPrivStatement.Text, ""
    Else
       cnPriv.u_attr_valid_value txtPrivStatement.Text, "", "photo_privacy_stmnt", 1
    End If
    If cnPriv.State = adStateOpen Then
        cnPriv.Close
        Set cnPriv = Nothing
    End If
    Exit Sub

Error_Save_PrivacyStatement:
    MsgBox Err.Description, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume Next
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
Dim strValidate As String
Dim i%
Dim strValue As String

    On Error GoTo Error_Load_Facilities
    cboFacility.Clear
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "Load__Active_Facilities"
    Set rsFac = cnFac.Execute(strSQL)
    With rsFac
        Do Until .EOF
                If Not IsNull(!Facility_Name) Then
                If Trim(!Facility_Name) <> vbNullString Then
                    strName = Trim(!Facility_Name)
                    strName = Fix_Name(strName)
                    strName = Fix_Facility_Name(strName)
                    cboFacility.AddItem strName
                    If (!Facility_ID) > 0 Then cboFacility.ItemData(cboFacility.NewIndex) = (!Facility_ID)
                End If
            End If
            .MoveNext
        Loop
    End With
    Get_DefaultFacility
    For i = 0 To cboFacility.listcount - 1
        If cboFacility.ItemData(i) = udtFacility.FacID Then
             cboFacility.Text = cboFacility.List(i)
        End If
    Next i
    If udtFacility.Logo <> vbNullString Then txtDefaultLogo.Text = Trim(udtFacility.Logo)
    If cboFacility.Text = vbNullString Then
        strValue = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultFacility")
        If strValue <> vbNullString Then
            For i = 0 To cboFacility.listcount - 1
                If cboFacility.ItemData(i) = Val(strValue) Then
                     cboFacility.Text = cboFacility.List(i)
                End If
            Next i
        End If
    End If
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set cnFac = Nothing
        Set rsFac = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Facilities:
    MsgBox Err.Description, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Unload Me
  Exit Sub
End Sub


Private Sub grdProviders_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long

    If Row < 1 Then Exit Sub
    With grdProviders
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


Private Sub grdProviders_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    With grdProviders
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = NormalColor
        .BlockMode = False
    End With
End Sub


Private Sub mn_Sys_Normal_Click()
    If mnu_Tablet.Checked Then mnu_Tablet.Checked = False
    SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "Windows_Magnified", "1", REG_SZ
    mlngSystemFileSize = 1
    Exit Sub
End Sub

Private Sub mnu_Enable_Photo_Click()
    fraImage.Visible = True
    Exit Sub
End Sub

Private Sub mnu_PhotoStatement_Click()
    fraPrivStatement.Visible = True
    Exit Sub
End Sub


Private Sub mnu_Provider_Title_Click()
    Format_Provider_Title
    fraProvider.Visible = True
    Exit Sub
End Sub
Private Sub mnu_Tablet_Click()
    If mn_Sys_Normal.Checked Then mn_Sys_Normal.Checked = False
    SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "Windows_Magnified", "1.122", REG_SZ
    mlngSystemFileSize = 1.122
    Exit Sub
End Sub
