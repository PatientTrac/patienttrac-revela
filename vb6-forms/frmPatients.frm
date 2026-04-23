VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPatients 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Patient Selection"
   ClientHeight    =   7875
   ClientLeft      =   8400
   ClientTop       =   4080
   ClientWidth     =   14310
   Icon            =   "frmPatients.frx":0000
   LockControls    =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   14310
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   13470
      Top             =   7470
   End
   Begin VB.Frame fraPatient 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6645
      Left            =   30
      TabIndex        =   0
      Top             =   510
      Width           =   14265
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
         ItemData        =   "frmPatients.frx":058A
         Left            =   1320
         List            =   "frmPatients.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   60
         Width           =   3765
      End
      Begin FPSpreadADO.fpSpread grdPatient 
         Height          =   3075
         Left            =   90
         TabIndex        =   15
         Top             =   3120
         Width           =   14115
         _Version        =   524288
         _ExtentX        =   24897
         _ExtentY        =   5424
         _StockProps     =   64
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
         MaxRows         =   0
         RowHeaderDisplay=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         ShadowText      =   8388608
         SpreadDesigner  =   "frmPatients.frx":058E
         Appearance      =   2
         AppearanceStyle =   0
      End
      Begin VB.Frame fraSearch 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   2955
         Left            =   120
         TabIndex        =   2
         Top             =   420
         Width           =   14115
         Begin VB.TextBox cctPatNote 
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
            Height          =   645
            Left            =   5730
            MultiLine       =   -1  'True
            TabIndex        =   33
            ToolTipText     =   "Provider Patient Notes"
            Top             =   1980
            Width           =   5445
         End
         Begin VB.TextBox txtSearch 
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
            Height          =   345
            Left            =   0
            TabIndex        =   31
            Top             =   300
            Width           =   5295
         End
         Begin VB.TextBox txtPharName 
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
            Left            =   5730
            TabIndex        =   24
            ToolTipText     =   "Use Phone Notes to Update"
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox txtPharPhone 
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
            Left            =   8790
            TabIndex        =   23
            ToolTipText     =   "Use Phone Notes to Update"
            Top             =   1560
            Width           =   2385
         End
         Begin VB.ComboBox cboNoteSel 
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
            Left            =   4230
            TabIndex        =   13
            Text            =   "Plan"
            Top             =   3240
            Width           =   2145
         End
         Begin VB.TextBox txtParner 
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
            Left            =   8790
            Locked          =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   9
            Top             =   270
            Width           =   2385
         End
         Begin VB.Frame fraPic 
            BackColor       =   &H00E0E0E0&
            Height          =   2625
            Left            =   11220
            TabIndex        =   8
            Top             =   -90
            Width           =   2865
            Begin VB.Label lblPhoto 
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "NO  PHOTO AVAILABLE"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   735
               Left            =   660
               TabIndex        =   16
               Top             =   900
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Image imgPatient 
               Height          =   2385
               Left            =   60
               Stretch         =   -1  'True
               Top             =   150
               Width           =   2745
            End
         End
         Begin VB.ListBox lstPatient 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            Left            =   0
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   690
            Width           =   5325
         End
         Begin PT_XP_Button.PT_XP cmdClear 
            Height          =   345
            Left            =   3270
            TabIndex        =   17
            Top             =   2010
            Width           =   1845
            _ExtentX        =   3254
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
            Picture         =   "frmPatients.frx":09F6
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Clear ListBox"
            CaptionAlignment=   7
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmPatients.frx":0D90
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdStatus 
            Height          =   315
            Left            =   5370
            TabIndex        =   34
            ToolTipText     =   "Intenal Patient Note"
            Top             =   2010
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
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
            Picture         =   "frmPatients.frx":10E2
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmPatients.frx":167C
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblSpeech 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Primary Language"
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
            Left            =   8790
            TabIndex        =   28
            Top             =   690
            Width           =   2055
         End
         Begin VB.Label lblPrimLang 
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
            Height          =   375
            Left            =   8790
            TabIndex        =   27
            Top             =   900
            Width           =   2385
         End
         Begin VB.Label lblPharPhTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pharmacy Phone"
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
            Left            =   8790
            TabIndex        =   26
            Top             =   1320
            Width           =   1650
         End
         Begin VB.Label lblPhNameTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pharmacy Name"
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
            Left            =   5730
            TabIndex        =   25
            Top             =   1320
            Width           =   1620
         End
         Begin VB.Label lblPatient 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Patient Search"
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
            TabIndex        =   21
            Top             =   30
            Width           =   1695
         End
         Begin VB.Label lblPatTitle 
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
            Height          =   375
            Left            =   5730
            TabIndex        =   20
            Top             =   270
            Width           =   2655
         End
         Begin VB.Label lblPatName 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Patient Name"
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
            Left            =   5730
            TabIndex        =   19
            Top             =   30
            Width           =   1365
         End
         Begin VB.Label lblPatChoice 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Note Selection"
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
            Left            =   4245
            TabIndex        =   14
            Top             =   3000
            Width           =   1905
         End
         Begin VB.Label lblRace 
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
            Height          =   375
            Left            =   5730
            TabIndex        =   12
            Top             =   900
            Width           =   2655
         End
         Begin VB.Label lblPartnerName 
            BackColor       =   &H00BFE8FD&
            BackStyle       =   0  'Transparent
            Caption         =   "Partner Name"
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
            Left            =   8790
            TabIndex        =   11
            Top             =   30
            Width           =   1725
         End
         Begin VB.Label lblRaceTitle 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Race"
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
            Left            =   5730
            TabIndex        =   10
            Top             =   690
            Width           =   975
         End
         Begin VB.Label lblTimeFilter 
            Height          =   105
            Left            =   570
            TabIndex        =   5
            Top             =   540
            Width           =   795
         End
         Begin VB.Label lblPatCount 
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Count"
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
            Left            =   90
            TabIndex        =   4
            Top             =   2160
            Width           =   1395
         End
         Begin VB.Label lblCount 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
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
            Height          =   345
            Left            =   1530
            TabIndex        =   3
            Top             =   2100
            Width           =   1575
         End
      End
      Begin PT_XP_Button.PT_XP cmdTakePic 
         Height          =   405
         Left            =   13620
         TabIndex        =   32
         ToolTipText     =   "Take Picture"
         Top             =   0
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPatients.frx":19CE
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmPatients.frx":22A8
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   90
         TabIndex        =   30
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label lblUserInstruct 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Please select the Active Patient in the Large Grid Box. "
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   4545
         TabIndex        =   7
         Top             =   6330
         Visible         =   0   'False
         Width           =   5235
      End
   End
   Begin PT_XP_Button.PT_XP cmdCancel 
      Height          =   345
      Left            =   7193
      TabIndex        =   18
      Top             =   7410
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
      Picture         =   "frmPatients.frx":2EFA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatients.frx":390C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdEnctrOK 
      Height          =   345
      Left            =   6083
      TabIndex        =   1
      Top             =   7410
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
      Picture         =   "frmPatients.frx":3C5E
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "OK"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatients.frx":3FF8
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPrint 
      Height          =   375
      Left            =   12630
      TabIndex        =   22
      ToolTipText     =   "Print Face Sheet"
      Top             =   60
      Width           =   1635
      _ExtentX        =   2884
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
      Picture         =   "frmPatients.frx":434A
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Face Sheet"
      DepthMode       =   1
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatients.frx":48E4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdEncntr 
      Height          =   380
      Left            =   3750
      TabIndex        =   35
      Top             =   60
      Width           =   2085
      _ExtentX        =   3678
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
      Picture         =   "frmPatients.frx":4C36
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Encounter/Visits"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatients.frx":51D0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdBillingInfo 
      Height          =   380
      Left            =   5910
      TabIndex        =   36
      Top             =   60
      Width           =   1635
      _ExtentX        =   2884
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
      Picture         =   "frmPatients.frx":5522
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Billing Info"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatients.frx":5ABC
      PictureAlignment=   0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdUpdateRecord 
      Height          =   380
      Left            =   1740
      TabIndex        =   37
      Top             =   60
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Picture         =   "frmPatients.frx":5E0E
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Update Patient"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatients.frx":63A8
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdNewPat 
      Height          =   375
      Left            =   30
      TabIndex        =   38
      Top             =   60
      Width           =   1635
      _ExtentX        =   2884
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
      Picture         =   "frmPatients.frx":66FA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "New Patient"
      Depth           =   5
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatients.frx":6C94
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdInsType 
      Height          =   375
      Left            =   7620
      TabIndex        =   39
      ToolTipText     =   "Patient Insurance"
      Top             =   60
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
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
      Picture         =   "frmPatients.frx":6FE6
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Insurance"
      CaptionAlignment=   5
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatients.frx":7580
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdAppt 
      Height          =   375
      Left            =   9090
      TabIndex        =   40
      Top             =   60
      Width           =   1785
      _ExtentX        =   3149
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
      Picture         =   "frmPatients.frx":78D2
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Appt History"
      DepthMode       =   1
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatients.frx":7E6C
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmPatients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fpimage As Picture
Private fstrDefProvider As String
Public piPatientID As Long
Public piProviderID As Long
Private fUpdateMode As String
Private mLocID As Integer
Private mbAdmit As Boolean
Private mbDischarge As Boolean
Private mRoomAssign As String
Private mNewPatient As Boolean
Private mstrSaveData As String
Private mRow As Integer
Public mstrSearchItem As String
Public gbooBypass As Boolean
Public mFirst As Boolean
Private mstrPlan As String
Private fstrNetSerPath As String
Private fbooByPassEvent As Boolean
Private fbooEncounterTypeReq As Boolean
Private fbooCloseAnyEnctr As Boolean
Private fbooUnlimitedEnctrs As Boolean
Private fbooServicePlaceReq As Boolean
Private piSelEncntrID As Long
Private fbooEndClicked As Boolean
Private fbooSaveClicked As Boolean
Public mbooNext As Boolean
Public mbooPrev As Boolean
Private mbooUpdate As Boolean
Private fbooEncounterClosed As Boolean
Private mstrPatNote As String
Private mfpNote As String
Private mlngGridRow As Long
Private mlngGridCol As Long
Private mbooSupport As Boolean
Private mbooNewEncntr As Boolean
Private mbooAddNewEnctr As Boolean
Public mbShowPatient As Boolean
Private mbooSearch As Boolean
Private mbooFoundPat As Boolean
Private mbooSkip As Boolean
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal Hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long

Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long
Sub AddTopBorder(Row As Long)

End Sub
Private Sub Clear_Fields()
    lblPatTitle.Caption = vbNullString
    lblRace.Caption = vbNullString
    txtParner.Text = vbNullString
    txtPharName.Text = vbNullString
    txtPharPhone.Text = vbNullString
    lblPrimLang.Caption = vbNullString
    mlngPrevEnctrID = 0
    piEncounterID = 0
    mLastEnctrDate = vbNullString
    Exit Sub
End Sub

Private Function Look_UpPatient_Facility(vstrData As String, Optional vstrType As String)
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim cnPhys As New ADODB.Connection
Dim rsPhys As New ADODB.Recordset
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim strSuffix As String
Dim lngRefID As Long
Dim strExtID As String
Dim strName As String
Dim strSearch As String
Dim lngInvID As Long
Dim lngPatID As Long
Dim Arr
Dim strFirst As String
Dim strTitle As String

    On Error GoTo Error_Look_UpPatient_Facility
    Screen.MousePointer = vbHourglass
    udtInPatSched.Facility_ID = piFacilityID
    LockWindowUpdate Hwnd
    lstPatient.Clear
    If Trim(vstrData) = vbNullString Then
        strSearch = Trim(txtSearch.Text)
    Else
        strSearch = Trim(vstrData)
    End If
    If TextFound("'", strSearch, Len("'"), True) Then
        strSearch = Replace(strSearch, "'", "%")
        strSearch = Replace(strSearch, " ", "%")
    End If
    Clear_UdtPatient
    clsListSearch.SearchType = Trim(cboSearchFilter.Text)
    If cnPhys.State = adStateClosed Then cnPhys.open psConnect(0)
    Select Case Trim(UCase(clsListSearch.SearchType))
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
                 strSQL = "Search_Patient_FirstName_Last_Name_Join   '" & strSearch & "%', '" & strFirst & "%'," & udtInPatSched.Facility_ID
            Else
                strSQL = "Search_Patient_Row_Count_Last_Name_Join 25, '" & strSearch & "%'," & udtInPatSched.Facility_ID
            End If
        Case "EMR ID"
            strSQL = "Search_Patient_EMR_ID_Join '" & strSearch & "%'," & udtInPatSched.Facility_ID
        Case "SOCIAL SECURITY"
            strSQL = "Search_Patient_Social_Security_Join '" & strSearch & "%'," & udtInPatSched.Facility_ID
        Case "EXTERNAL ID"
            strSQL = "Search_Patient_RowCount_External_ID_Join 25,'" & strSearch & "%'," & udtInPatSched.Facility_ID
        Case "INVOICE NO"
            strSQL = "Search_Patient_Invoice_Invoice_No_Join '" & strSearch & "%'," & udtInPatSched.Facility_ID
            Set rsPat = cnPat.Execute(strSQL)
            If Not rsPat.EOF Then
                lngPatID = rsPat.Fields("ACCOUNT_NO")
            End If
            If lngPatID = 0 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            strSQL = "Get_Patient_BY_Id " & lngPatID
    End Select
    If strSQL <> vbNullString Then
        If cnPhys.State = adStateClosed Then cnPhys.open psConnect(0)
        Set rsPhys = cnPhys.Execute(strSQL)
         With rsPhys
             Do Until .EOF
                lngRefID = 0
                Clear_UdtPatient
                lngRefID = 0
                lngRefID = (!Patient_ID)
                If Not IsNull(!PATIENT_EXT_REF) Then
                    strExtID = Trim((!PATIENT_EXT_REF))
                End If
                rsPhys.MoveNext
                Get_PatientName lngRefID
                If Trim(udtPatientName.Suffix) <> vbNullString Then
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                Else
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                End If
                Select Case UCase(clsListSearch.SearchType)
                  Case "LAST NAME"
                      If Trim(udtPatientName.Name) <> vbNullString And Len(udtPatientName.Name) > 1 Then
                          lstPatient.AddItem Trim(udtPatientName.Name)
                          If lngRefID > 0 Then lstPatient.ItemData(lstPatient.NewIndex) = lngRefID
                      End If
                  Case "EMR ID"
                      If Trim(strName) <> vbNullString And Len(strName) > 1 Then
                          lstPatient.AddItem Trim(strName)
                          If lngRefID > 0 Then lstPatient.ItemData(lstPatient.NewIndex) = lngRefID
                      ElseIf Len(Trim(udtPatientName.Name)) > 1 Then
                          lstPatient.AddItem Trim(udtPatientName.Name)
                          If lngRefID > 0 Then lstPatient.ItemData(lstPatient.NewIndex) = lngRefID
                      End If
                  Case "EXTERNAL ID"
                      If Trim(udtPatientName.Name) <> vbNullString And Trim(strExtID) <> vbNullString And Trim(strExtID) <> "0" Then
                          lstPatient.AddItem Trim(udtPatientName.Name) & " - " & strExtID
                          If lngRefID > 0 Then lstPatient.ItemData(lstPatient.NewIndex) = lngRefID
                      End If
                End Select
             Loop
          End With
    End If
    If cnPat.State = adStateOpen Then
        cnPat.Close
        Set cnPat = Nothing
        Set rsPat = Nothing
    End If
    If cnPhys.State = adStateOpen Then
        cnPhys.Close
        Set cnPhys = Nothing
        Set rsPhys = Nothing
        Set Flds = Nothing
    End If
    Clear_UdtPatient
    LockWindowUpdate 0
    Screen.MousePointer = vbDefault
    Exit Function
Error_Look_UpPatient_Facility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Look_UpPatient_Facility", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub PatientGrid()
Dim lngPatientID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim mfpNote As String
Dim lngPatID
Dim varPatRow
Dim varPatSecRow
Dim booNoPhoto As Boolean
Dim fso As New FileSystemObject


    On Error GoTo Error_PatientGrid

    Screen.MousePointer = vbHourglass
    If Right(psImageFilePath, 1) <> "\" Then
        psImageFilePath = psImageFilePath & "\"
    End If
    
    If fso.FileExists(psImageFilePath & piPatientID & "\Photo.jpg") Then
        psPhoto = psImageFilePath & piPatientID & "\Photo.jpg"
        imgPatient.Picture = LoadPicture(psPhoto)
        lblPhoto.Visible = False
'        If Not booNoPhoto Then
'            imgPatient.Picture = LoadPicture(psPhoto)
'            lblPhoto.Visible = False
'        ElseIf fso.FileExists(psImageFilePath & "\" & piPatientID & "\Photo.jpg") Then
'            psPhoto = psImageFilePath & "\" & piPatientID & "\Photo.jpg"
'
'        End If
    Else
       Set imgPatient.Picture = Nothing
       lblPhoto.Visible = True
    End If
    With grdPatient
        .Redraw = False
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = 1
            .Row2 = .MaxRows
            .FontBold = True
            .BackColor = grdGray
             txtSearch.Text = vbNullString
        .BlockMode = False

        .Col = mlngGridCol
        .Row = mlngGridRow
        If mlngGridCol >= 1 And mlngGridRow <> 0 Then
            lngPatID = grdPatient.GetRowItemData(mlngGridRow)
            If lngPatID = 0 Then
                lngPatID = grdPatient.GetRowItemData(mlngGridRow - 1)
            End If
            If lngPatID = 0 Then Exit Sub
            piPatientID = lngPatID
            If piPatientID > 0 Then modSpecific.piPatientID = piPatientID
            Get_Patient piPatientID
            .Col = mlngGridCol
            .Row = mlngGridRow
            .BlockMode = True
                .Col = 0
                .Col2 = .MaxCols
                .Row = mlngGridRow
                .Row2 = mlngGridRow
                .FontBold = True
                .BackColor = grdGray
            .BlockMode = False
             lblPatTitle.Caption = clsPatient.Name
             lblRace.Caption = clsPatient.Race
             lblPrimLang.Caption = clsPatient.Language
             Get_Partner piPatientID
             txtParner.Text = clsPartner.Name
             Get_PatientClinical
             .GetText 2, 1, varPatRow
             .GetText 2, 3, varPatSecRow
             If varPatRow = varPatSecRow Then .MaxRows = 2
       Else
            Get_Patient piPatientID
            lblPatTitle.Caption = clsPatient.Name
            lblRace.Caption = clsPatient.Race
            lblPrimLang.Caption = clsPatient.Language
            Get_Partner piPatientID
            txtParner.Text = clsPartner.Name
            If grdPatient.CellType <> CellTypePicture Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            If grdPatient.TypePictPicture <> fpimage Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            lngRow = mlngGridRow + 1
            grdPatient.MaxRows = lngRow
            If grdPatient.RowHidden = True Then
                grdPatient.RowHidden = False
            Else
                grdPatient.RowHidden = True
            End If
       End If
    .Redraw = True
   End With
  ' Read_Patient_Flag
    Screen.MousePointer = vbDefault
   Exit Sub
Error_PatientGrid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "grdPatient", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Function Set_Patient(vnlgPatientID As Long)
Dim lngRow As Long
Dim strPhoto As String
Dim fso As New FileSystemObject


    On Error GoTo Error_Set_Patient
    If pbCancel Then
        mbooByPass = True
        lstPatient.Clear
        grdPatient.ClearRange 0, 1, grdPatient.MaxRows, grdPatient.MaxRows, True
        grdPatient.MaxRows = 0
        mbooByPass = True
        Clear_Fields
        Dump_PatientStrData
        modSpecific.piPatientID = 0
        piPatientID = 0
        mbooByPass = False
        pbCancel = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    If piPatientID <> 0 Then
        Get_PatientName piPatientID
    Else
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    mLastEnctrDate = vbNullString
    lngRow = lngRow + 1
    grdPatient.MaxRows = lngRow
    If udtPatientName.Gender = vbNullString Then
         MsgBox "No Gender!", vbInformation
         mbShowPatient = True
         pbCancel = False
     End If
    If udtPatientName.DOB = vbNullString Then
         MsgBox "No Date of Birth!", vbInformation
         mbShowPatient = True
         pbCancel = False
     End If
    mfpNote = " Birthday: " & udtPatientName.DOB & vbNewLine
    mfpNote = mfpNote & " Address: " & udtPatientName.Address & vbNewLine
    mfpNote = mfpNote & " " & udtPatientName.CityStateZip & vbNewLine
    mfpNote = mfpNote & " Social_Security " & udtPatientName.SS
    With grdPatient
        strPhoto = mstrProgram & "\FirstNotebook\PatientTrac EMR\Images\CARD.BMP"
        If fso.FileExists(strPhoto) Then
            Format_CellTypePicture grdPatient, lngRow, 1, "1", strPhoto, "TypeHAlignCenter", "", 15, mfpNote
            grdPatient.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        Format_CellTypeEdit grdPatient, lngRow, 2, "1", udtPatientName.Name, "TypeHAlignCenter", "PATIENT_ID"
        grdPatient.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeEdit grdPatient, lngRow, 3, "1", CStr(udtPatientName.IntId), "TypeHAlignCenter", "PATIENT_ID"
        grdPatient.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeEdit grdPatient, lngRow, 4, "1", CStr(udtPatientName.ExtID), "TypeHAlignCenter", "EXT_PATIENT_ID"
        grdPatient.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdPatient.SetRowItemData lngRow, piPatientID
        modSpecific.piPatientID = piPatientID
        If piPatientID > 0 Then Get_LastVisit piPatientID
        Format_CellTypeEdit grdPatient, lngRow, 5, "1", Format(mLastEnctrDate, "MMMM DD YYYY"), "TypeHAlignCenter", "LAST_ENCTR"
        grdPatient.SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        mlngGridRow = lngRow
        Screen.MousePointer = vbDefault
        PatientGrid
    End With
    If mbShowPatient Then
        mbooSkip = True
        PatientInfo.piPatientID = piPatientID
        PatientInfo.mNewPatient = False
        Screen.MousePointer = vbDefault
        udtSchedPatient.IntId = piPatientID
        PatientInfo.Show
        mbShowPatient = False
    End If
    pbCancel = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Set_Patient", App.EXEName & "." & Me.Name
    Exit Function
    Resume
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

Private Sub cboNoteSel_Change()
Dim i%

    
    For i = 0 To cboNoteSel.listcount - 1
        If cboNoteSel.Text = "Insurance" Then
            mstrPatNote = cboNoteSel.Text
        ElseIf cboNoteSel.Text = "Plan" Then
            mstrPatNote = cboNoteSel.Text
        End If
    Next i
End Sub

Private Sub cboNoteSel_Click()
Dim i%

    
    For i = 0 To cboNoteSel.listcount - 1
        If cboNoteSel.Text = "Insurance" Then
            mstrPatNote = cboNoteSel.Text
        ElseIf cboNoteSel.Text = "Plan" Then
            mstrPatNote = cboNoteSel.Text
        End If
    Next i
'    If cboNoteSel.Text <> vbNullString Then Update_PatientGrid
End Sub
Private Function FixPatient(vlngPatientID As Long)
Dim i%
Dim strSQL As String
Dim cnFixPat As New ADODB.Connection
Dim rsFixPat As New ADODB.Recordset

    mbooFoundPat = False
    If vlngPatientID > 0 Then
        If cnFixPat.State = adStateClosed Then cnFixPat.open psConnect(0)
        strSQL = "Get_Patient_BY_Id " & vlngPatientID
        Set rsFixPat = cnFixPat.Execute(strSQL)
        If Not rsFixPat.EOF Then
            If rsFixPat.Fields("PURGE") = "True" Then mbooFoundPat = True
        Else
            mbooFoundPat = False
        End If
        If mbooFoundPat Then
            strSQL = "Update_Patient_Purge " & vlngPatientID
            Set rsFixPat = cnFixPat.Execute(strSQL)
        End If
    End If
    If cnFixPat.State = adStateOpen Then
        cnFixPat.Close
        Set cnFixPat = Nothing
        Set rsFixPat = Nothing
    End If
    Exit Function
End Function
Private Function Get_EncounterID()
Dim strSQL As String
Dim sDate As String
Dim cnEntr As New ADODB.Connection
Dim rsEntr As New ADODB.Recordset

    Screen.MousePointer = vbHourglass
    If cnEntr.State = adStateClosed Then cnEntr.open psConnect(0)
    strSQL = "Get_Encounter_Patient_ID " & piPatientID
    Set rsEntr = cnEntr.Execute(strSQL)
    If Not rsEntr.EOF Then
        piEncounterID = rsEntr.Fields("ENCOUNTER_ID")
    End If
    If cnEntr.State = adStateOpen Then
        cnEntr.Close
        Set cnEntr = Nothing
        Set rsEntr = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
End Function
Private Sub cboSearchFilter_Click()
    lstPatient.Clear
    Dump_PatientStrData
    modSpecific.piPatientID = 0
    piPatientID = 0
    txtSearch.Text = vbNullString
    txtSearch.SetFocus
    grdPatient.ClearRange 0, 1, grdPatient.MaxRows, grdPatient.MaxRows, True
    grdPatient.MaxRows = 0
    Clear_Fields
    Exit Sub
End Sub


Private Sub cmdAppt_Click()
    frmPatAppHist.Show vbModal
    Exit Sub
End Sub

Private Sub cmdBillingInfo_Click()
    Screen.MousePointer = vbHourglass
    If piEncounterID = 0 Then Get_EncounterID
    If piEncounterID = 0 Then
        MsgBox "Billing information is not available since this patient has no encounters!", vbInformation, "PatientTrac Patient Selection"
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        mstrPatOcx = "FN_PsyMse"
        mstrPatCtl = "SuperBill"
        frmLoadControl.piSelEnctrID = piEncounterID
        frmLoadControl.piSelPatientID = piPatientID
        Load frmLoadControl
        frmLoadControl.Show vbModal
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
   If Not IsFormLoaded("frmNursingHome") Then
        frmMDI.Get_AppointmentFormat
    End If
    Exit Sub
End Sub

Private Sub cmdClear_Click()
Dim i%


    For i = 0 To lstPatient.listcount - 1
        If lstPatient.Selected(i) Then
            mbooByPass = True
            lstPatient.Selected(i) = False
            mbooByPass = False
        End If
    Next i
    Exit Sub
End Sub


Public Sub cmdEncntr_Click()
Dim i%
Dim varEnctrID

    On Error GoTo Error_Select_Enctr
     Screen.MousePointer = vbHourglass
    piEncounterID = 0
    frmMDI.StatusBar.Panels(1).Text = "Patient: " & clsPatient.Name
    frmMDI.StatusBar.Panels(2).Text = "External ID: " & clsPatient.ExtID
    frmMDI.StatusBar.Panels(3).Text = "Internal ID: " & clsPatient.IntId
    frmMDI.StatusBar.Panels(4).Text = "Encounter: " & piEncounterID
    frmMDI.StatusBar.Panels(5).Text = "Visit Date: " & Format(Now, "MMMM DD, YYYY")
    frmMDI.StatusBar.Panels(6).Text = "Provider: " & clsProvider.Name
    If IsFormLoaded("frmEnctr") Then Unload frmEnctr
    If IsFormLoaded("frmPatientAppointment") Then Unload frmPatientAppointment
    If piPatientID > 0 Then
        frmInPatient.mstrLoadType = "Patient"
        frmEnctr.piPatientID = piPatientID
        frmEnctr.Show
        Unload Me
     Else
        MsgBox "Please select Patient first!", vbInformation
    End If
     Screen.MousePointer = vbDefault
    Exit Sub
Error_Select_Enctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Select_Enctr", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdEncntr_GotFocus()
    cmdEncntr_KeyPress 13
End Sub

Private Sub cmdEncntr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
        cmdEncntr_Click
    End If
End Sub


Private Sub cmdEncntr_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblUserInstruct.Visible = True
End Sub
Private Sub cmdEnctrOK_Click()
Dim i%
Dim varEnctrID
    
    On Error GoTo Error
 
    If IsFormLoaded("frmEnctr") Then Unload frmEnctr
    If IsFormLoaded("frmPatientAppointment") Then Unload frmPatientAppointment
    frmMDI.StatusBar.Panels(1).Text = "Patient: " & clsPatient.Name
    frmMDI.StatusBar.Panels(2).Text = "Hospital ID: " & clsPatient.ExtID
    frmMDI.StatusBar.Panels(3).Text = "Patient EMR ID: " & clsPatient.IntId
    piEncounterID = 0
    frmMDI.StatusBar.Panels(4).Text = "Encounter Number: " & vbNullString
    frmMDI.StatusBar.Panels(5).Text = "Visit Date:: " & Format(Now, "MMMM DD, YYYY")
    frmMDI.StatusBar.Panels(6).Text = "Provider: " & clsProvider.Name
    If piPatientID <> 0 Then
        frmInPatient.mstrLoadType = "Patient"
        frmEnctr.piPatientID = piPatientID
        frmEnctr.Show
        Unload Me
     Else
        MsgBox "Please select Patient first!", vbInformation
    End If

Error:
    Exit Sub
    Resume
End Sub


Private Sub cmdInsType_Click()
   If piPatientID <= 0 Then
        MsgBox "Please select a Patient first!", vbInformation, "PatientTrac Scheduling"
        Exit Sub
    Else
        mstrPatCtl = "PrimIns"
        mstrPatOcx = "FN_PP_PS"
        frmLoadControl.piSelPatientID = piPatientID
        Load frmLoadControl
        frmLoadControl.Show vbModal
    End If
    Exit Sub
End Sub

Private Sub cmdNewPat_Click()
    On Error GoTo Error_New_Patient
    Screen.MousePointer = vbHourglass
    lstPatient.Clear
    lblPatTitle.Caption = vbNullString
    lblCount.Caption = vbNullString
    grdPatient.MaxRows = 1
    grdPatient.ClearRange 0, 0, grdPatient.MaxCols, grdPatient.MaxRows, False
    modSpecific.piPatientID = 0
    modSpecific.piEncounterID = 0
    piPatientID = 0
    piEncounterID = 0
    Dump_PatientStrData
    PatientInfo.mNewPatient = True
    udtSchedPatient.IntId = vbNullString
    PatientInfo.Show vbModal
    mbShowPatient = False
    GetData piPatientID
    Screen.MousePointer = vbDefault
    Exit Sub
Error_New_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "New_Patient", Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdPrint_Click()
    If piPatientID = 0 Then
        MsgBox "Please select patient first!", vbInformation, "PatientTrac Patient Selection"
        mRow = 0
        InitSpread
        Set_Patient piPatientID
        Exit Sub
    End If
    mstrPatOcx = "PatientTrac_Report": mstrPatCtl = "PatientDemographic"
    frmLoadControl.Caption = "PATIENT DEMOGRAPHICS"
    frmLoadControl.piSelPatientID = piPatientID
    Load frmLoadControl
    frmLoadControl.Show vbModal
    mRow = 0
    InitSpread
    Set_Patient piPatientID
    Exit Sub
End Sub



Private Sub cmdStatus_Click()
    frmComment.Show vbModal
    Get_Add_Notes
    Exit Sub
End Sub

Private Sub cmdTakePic_Click()
Dim fso As New FileSystemObject

   On Error GoTo Error_Take_Photo
   CaptureImage frmImageCapture, piPatientID
   If Right(psImageFilePath, 1) = "\" Then
        psPhoto = psImageFilePath & piPatientID & "\Photo.jpg"
    Else
        psPhoto = psImageFilePath & "\" & piPatientID & "\Photo.jpg"
    End If
   If fso.FileExists(psPhoto) Then
      imgPatient.Picture = LoadPicture(psPhoto)
      lblPhoto.Visible = False
   Else
      imgPatient.Picture = LoadPicture()
      lblPhoto.Visible = True
   End If
   Exit Sub
Error_Take_Photo:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, psPhoto, "Take_Photo", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdUpdateRecord_Click()
Dim strSQL As String
Dim bRet As Boolean
Dim lngrecs As Long
Dim varRS
Dim strUser As String

    On Error GoTo Error
    If piPatientID > 0 Then
        Create_Folder piPatientID
    End If
    PatientInfo.piPatientID = piPatientID
    If piProviderID = 0 Then Get_DefaultProvider
    PatientInfo.mNewPatient = False
    Me.Hide
    udtSchedPatient.IntId = piPatientID
    PatientInfo.Show vbModal
    mbShowPatient = False
    Me.Show
    Exit Sub
Error:
    MsgBox Err.Description, vbInformation
    Exit Sub
    Resume
End Sub

Private Sub Form_Activate()
Dim i%

    On Error GoTo Error_Form_Activate
    Screen.MousePointer = vbHourglass
    mbooSkip = False
    mbooSupport = False
    mbShowPatient = False
    If mFirst Then
        mFirst = False
        If Not mNewPatient Then
            If IsFormLoaded("frmInPatient") Then Unload frmInPatient
        End If
        mEnctrCounter = 0
        mRow = 0
        If pbCaptureSetting Then cmdTakePic.Visible = True
        Set fpimage = LoadPicture(App.Path & "\images\card.bmp")
        psImageFilePath = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ImagePath")
        If psImageFilePath = vbNullString Then
           psImageFilePath = "C:\"
        End If
        mbShowPatient = False
        GetTerms cboNoteSel, "sd_Patient_Items"
        If pbCaptureSetting Then cmdTakePic.Visible = True
        mbShowPatient = False
        Screen.MousePointer = vbDefault
        If mstrPractice = "Psychiatry" Then
            If cboSearchFilter.listcount <= 0 Then
                cboSearchFilter.AddItem "LAST NAME"
                cboSearchFilter.AddItem "SOCIAL SECURITY"
                cboSearchFilter.AddItem "EMR ID"
                cboSearchFilter.AddItem "EXTERNAL ID"
            End If
        Else
            If cboSearchFilter.listcount <= 0 Then
                cboSearchFilter.AddItem "LAST NAME"
                cboSearchFilter.AddItem "SOCIAL SECURITY"
                cboSearchFilter.AddItem "EMR ID"
                cboSearchFilter.AddItem "FINANCIAL ACCT NO"
            End If
        End If
        If cboSearchFilter.listcount >= 0 Then
            For i = 0 To cboSearchFilter.listcount - 1
                If cboSearchFilter.List(i) = "LAST NAME" Then
                    mbooByPass = True
                    cboSearchFilter.Text = cboSearchFilter.List(i)
                    Exit For
                End If
            Next i
        End If
        If piPatientID > 0 Then
            InitSpread
            Set_Patient piPatientID
        Else
            InitSpread
        End If
        lstPatient.ListIndex = -1
        lstPatient.Clear
        centerForm Me
    Else
        lstPatient.ListIndex = -1
        lstPatient.Clear
        InitSpread
        Set_Patient piPatientID
        centerForm Me
    End If
    Get_Patient_Count
    Screen.MousePointer = vbDefault
  Exit Sub
Error_Form_Activate:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Form_Activate", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Public Sub Form_Load()
    mFirst = True

    Exit Sub
End Sub

Sub InitSpread()

    Screen.MousePointer = vbHourglass
    
    With grdPatient
        .Redraw = False
        .MaxCols = 5
        .MaxRows = mRow
        .ColHeaderRows = 2
        .ClearRange 0, 0, grdPatient.MaxCols, grdPatient.MaxRows, False
        .RowHeadersShow = False  'turn off row headers
        .EditModePermanent = True   'hide the focus rectangle
        .ScrollBars = ScrollBarsBoth    'Only show vertical scrollbars
        .GridShowHoriz = False  'turn off horizontal grid lines
        .GridShowVert = False   'turn off vertical grid lines
        .BlockMode = True
            .Col = 0
            .Row = SpreadHeader
            .Col2 = .MaxCols
            .Row2 = SpreadHeader
            .ColWidth(1) = 7       'set picture col width
            .ColWidth(2) = 60   'Set name, memo col width
            .ColWidth(3) = 12
            .ColWidth(4) = 14
            .ColWidth(5) = 20
            .shadowColor = &HFFC0FF
            .FontBold = True
        .BlockMode = False
        .RowHeight(SpreadHeader) = 20
        Format_CellTypeStatic grdPatient, SpreadHeader, 1, "1", " ", "TypeHAlignCenter", " ", HeaderColor
        grdPatient.SetCellBorder 1, SpreadHeader, 1, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeStatic grdPatient, SpreadHeader, 2, "1", "Patient Name", "TypeHAlignCenter", " ", HeaderColor
        grdPatient.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeStatic grdPatient, SpreadHeader, 3, "1", "PatientTrac ID", "TypeHAlignCenter", " ", HeaderColor
        grdPatient.SetCellBorder 3, SpreadHeader, 3, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        If mstrPractice = "Gyn" Then
            Format_CellTypeStatic grdPatient, SpreadHeader, 4, "1", "Financial ID", "TypeHAlignCenter", " ", HeaderColor
        grdPatient.SetCellBorder 4, SpreadHeader, 4, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Else
            Format_CellTypeStatic grdPatient, SpreadHeader, 4, "1", "Hospital / ACCT ID", "TypeHAlignCenter", " ", HeaderColor
            grdPatient.SetCellBorder 4, SpreadHeader, 4, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        Format_CellTypeStatic grdPatient, SpreadHeader, 5, "1", "Last Visit Date", "TypeHAlignCenter", " ", HeaderColor
        grdPatient.SetCellBorder 5, SpreadHeader, 5, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .TextTip = TextTipFixed
        .SetTextTipAppearance "Times New Roman", 10, True, False, &HFFFF&, &H800000
        Add_Spacer SpreadHeader + 1, grdPatient
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub


Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
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
            .SetText vlngCol, vlngRow, Trim(vstrText)
            .TypeEditMultiLine = True   'Make multi-lines
            .FontName = "Times New Roman"
            .FontSize = 8
            .FontBold = True
            'If vbooUnderline Then .FontUnderline = True
            If vlngBackColor > 0 Then .BackColor = vlngBackColor
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString Then
                lngSpan = vstrSpan
            If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            .Text = Trim(vstrText)
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then
                .RowHeight(vlngRow) = CellNote
            End If
            .TypeVAlign = TypeVAlignCenter
            If vstrHAligh = "TypeHAlignLeft" Then
                .TypeHAlign = TypeHAlignLeft
            Else
                .TypeHAlign = TypeHAlignCenter
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional lngBackColor As Long, Optional vstrCellNote As String, Optional vstrVAlign As String)
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
             .TypeVAlign = TypeVAlignCenter
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
Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)

    With vObjGrid
        vlngRow = vlngRow
       If .MaxRows < vlngRow Then .MaxRows = vlngRow
        .BlockMode = True
            .Col = 0
            .Col2 = 0
            .Row = vlngRow
            .Row2 = vlngRow
            .RowHeight(vlngRow) = 3.5
            .BackColor = 12632256
            .AddCellSpan 0, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
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
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
          .BlockMode = False
    End With
End Function

Private Function Format_CellTypePicture(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long
Dim fso As New FileSystemObject

    With vGrid
        .Redraw = False
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
            .TypePictStretch = True
            .TypePictCenter = True
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
        .Redraw = True
    End With
End Function
Sub GetData(vlngPatientID As Long)
Dim sFirstName As String
Dim sLastName As String
Dim sPatientName As String
Dim strSQL As String
Dim varRS
Dim lngrecs As Long
Dim booRet
Dim i%
Dim R%
Dim varBirth
Dim lngPatExtRef As Long
Dim lngPatID
Dim vartemp
Dim lngRow As Long
Dim lngCol As Long
Dim varFilterTime As Long
Dim varCurrentDate
Dim varTime
Dim lngCount As Long
Dim sName As String

    On Error GoTo Error_GetData
    Screen.MousePointer = vbHourglass
    lstPatient.ListIndex = -1
    lstPatient.Clear
    Get_Patient piPatientID
    
    If Val(clsPatient.IntId) > 0 Then
        mbooByPass = True
        lstPatient.AddItem Trim(clsPatient.Name)
        For i = 0 To lstPatient.listcount - 1
            If lstPatient.List(i) = clsPatient.Name Then
                lstPatient.Text = lstPatient.List(i)
                lstPatient_ItemCheck (i)
                Exit For
            End If
        Next i
    End If
    If piPatientID > 0 Then Create_Folder piPatientID
    Screen.MousePointer = vbDefault
  Exit Sub
Error_GetData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "GetData", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Form_Paint()
    ColorBurst Me, "BLUE", 128, 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next
    If IsFormLoaded("frmPatientAppointment") Then
        frmPatientAppointment.AutoRedraw = True
    ElseIf IsFormLoaded("frmInPatientAppt") Then
        frmInPatientAppt.AutoRedraw = True
    End If
    mfpNote = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Exit Sub
End Sub

Public Sub grdPatient_Click(ByVal Col As Long, ByVal Row As Long)

    If Row <= 0 Then Exit Sub
    mlngGridRow = Row
    mlngGridCol = Col
    Exit Sub
End Sub

Private Sub grdPatient_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim strSQL As String
Dim bRet As Boolean
Dim lngrecs As Long
Dim varRS
Dim strUser As String

    PatientInfo.piPatientID = piPatientID
    If piProviderID = 0 Then Get_DefaultProvider
    PatientInfo.mNewPatient = False
    Set_Patient piPatientID
    udtSchedPatient.IntId = piPatientID
    PatientInfo.Show vbModal
    mbShowPatient = False
    Set_Patient piPatientID
    Exit Sub
End Sub


Private Sub grdPatient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
        cmdEncntr.SetFocus
    End If
End Sub

Private Sub grdPatient_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Col As Long, Row As Long
    
    grdPatient.Row = Row
    grdPatient.Col = Col
    If grdPatient.CellType = CellTypePicture Then
        If grdPatient.TypePictPicture = fpimage Then
            If grdPatient.CursorStyle <> CursorStyleArrow Then grdPatient.CursorStyle = CursorStyleArrow
            grdPatient.ToolTipText = "Click Picture to review previous PLAN Note."
        End If
    Else
        If grdPatient.CursorStyle <> CursorStyleDefault Then grdPatient.CursorStyle = CursorStyleDefault
        If grdPatient.Col = 2 Then
            grdPatient.ToolTipText = "Double Click to review Complete Patient Information or to Update Record; Right Click to go to encounter!"
        End If
    End If
End Sub

Private Sub grdPatient_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If ClickType = 0 Then
        If piPatientID <> 0 Then
            frmEnctr.piPatientID = piPatientID
            frmEnctr.Show
            Unload Me
        Else
            MsgBox "Please select Patient First!", vbInformation
            Exit Sub
        End If
    End If
End Sub

Private Sub grdPatient_TextTipFetch(ByVal Col As Long, ByVal Row As Long, Multiline As FPSpreadADO.TextTipFetchMultilineConstants, TipWidth As Long, TipText As String, ShowTip As Boolean)
    
    'Only display text tip if a cell note
    If grdPatient.IsFetchCellNote Then
        ShowTip = True
    Else
        ShowTip = False
    End If
End Sub



Private Function Get_Patient_Count()
Dim cnPhys As New ADODB.Connection
Dim rsPhys As New ADODB.Recordset
Dim strSQL As String

    If cnPhys.State = adStateClosed Then cnPhys.open psConnect(0)
    strSQL = "Get_Active_Patient_Count"
    Set rsPhys = cnPhys.Execute(strSQL)
    If Not rsPhys.EOF Then
        lblCount.Caption = rsPhys.Fields(0)
    End If
     If cnPhys.State = adStateOpen Then
        cnPhys.Close
        Set cnPhys = Nothing
        Set rsPhys = Nothing
    End If
    Exit Function
End Function

Public Sub lstPatient_Click()


End Sub


Private Sub lstPatient_ItemCheck(Item As Integer)
Dim i%
Dim lngRow As Long
Dim lngPatID As Long
Dim strSQL As String
Dim bSkip As Boolean

    If mbooByPass Then mbooByPass = False: Exit Sub
    On Error GoTo Error_Select_Patient
    If Item < 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If Item < 0 Then Exit Sub
    If Not lstPatient.Selected(Item) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If lstPatient.Selected(Item) Then
        If PsUserID = vbNullString Then PsUserID = Get_PsUserID(psConnect(0))
        If clsEDI_Client.Site_ID = "1999" Then
            If TextFound("Supp", PsUserID, Len("Supp"), True) Or TextFound("kame", PsUserID, Len("kame"), True) Then
                mbooSupport = True
            End If
            lngPatID = lstPatient.ItemData(Item)
            If lngPatID > 0 And Validate_Provider_Permission(lngPatID) Then
                bSkip = False
            ElseIf lngPatID > 0 And Not mbooSupport Then
                MsgBox "This patient's records are currently in the facility which has not been assigned to you. Please contact PatientTrac support to update facility status for this patient!", vbInformation, "PatientTrac Patient Roster"
                bSkip = True
                mbooByPass = True
                lstPatient.Selected(Item) = False
                mbooByPass = False
            ElseIf lngPatID > 0 And mbooSupport Then
                bSkip = False
            End If
        End If
        If Not bSkip Then
             If lstPatient.listcount <= 0 Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            frmLoadControl.piSelPatientID = 0
            grdPatient.Redraw = False
            For i = 0 To lstPatient.listcount - 1
                If lstPatient.ItemData(i) <> lstPatient.ItemData(Item) Then
                    mbooByPass = True
                    lstPatient.Selected(i) = False
                End If
            Next i
            If Not mbooSkip Then
                Dump_PatientStrData
                modSpecific.piPatientID = 0
                piPatientID = 0
                lngRow = 0
                InitSpread
            End If
            grdPatient.MaxRows = lngRow
            lngPatID = lstPatient.ItemData(Item)
            If lngPatID <> 0 Then
                piPatientID = lngPatID
                txtSearch.Text = vbNullString
                Set_Patient lngPatID
                Get_Add_Notes
            End If
            cmdUpdateRecord.Visible = True
            mbooByPass = True
            lstPatient.Selected(Item) = True
            mbooByPass = False
            cmdInsType.Visible = True
        End If
    Else
        grdPatient.ClearRange 0, 1, grdPatient.MaxRows, grdPatient.MaxRows, True
        grdPatient.MaxRows = 0
        Clear_Fields
        Dump_PatientStrData
        modSpecific.piPatientID = 0
        piPatientID = 0
        lngRow = 0
    End If
    pbCancel = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Select_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Select_Patient", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
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
Private Function Get_Add_Notes()
Dim strSQL As String
Dim cnRead As New ADODB.Connection
Dim rsRead As New ADODB.Recordset

    On Error GoTo Error_Get_Add_Notes
    Screen.MousePointer = vbHourglass
    cctPatNote.Text = vbNullString
    If cnRead.State = adStateClosed Then cnRead.open psConnect(0)
    strSQL = "Collect_InPatient_Notes_PT_Type 'INTERNAL_PATIENT_NOTE'," & piPatientID
    Set rsRead = cnRead.Execute(strSQL)
    With rsRead
        Do Until .EOF
            If Trim(!NOTES_DESC) <> vbNullString Then
                cctPatNote.Text = Trim((!NOTES_DESC))
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    cctPatNote.Text = Replace(cctPatNote.Text, "''", "'")
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
    Resume
End Function

Private Sub Timer1_Timer()
Dim strList As String
    
    strList = Trim(txtSearch.Text)
    grdPatient.ClearRange 0, 1, grdPatient.MaxRows, grdPatient.MaxRows, True
    grdPatient.MaxRows = 0
    Clear_Fields
    If Trim(strList) <> vbNullString And Len(strList) > 1 Then
        Select Case cboSearchFilter.Text
            Case "LAST NAME"
                Look_UpPatient strList, cboSearchFilter.Text
            Case "SOCIAL SECURITY", "EXTERNAL ID"
                Look_UpPatient strList, cboSearchFilter.Text
            Case "EMR ID", "FINANCIAL ACCT NO"
                If IsNumeric(txtSearch.Text) And Len(txtSearch.Text) >= 2 Then
                    If cboSearchFilter.Text = "Financial Acct No" Then
                        Look_UpPatient strList, "FINANCIAL ACCT NO"
                    Else
                        Look_UpPatient strList, cboSearchFilter.Text
                    End If
                Else
                    MsgBox cboSearchFilter.Text & " " & "requires numeric values!", vbInformation, "PatientTrac Patient Search"
                    txtSearch.Text = vbNullString
                    imgPatient.Visible = False
                    lstPatient.Clear
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
        End Select
    End If
    mbooByPass = False
    Timer1.Enabled = False
  Exit Sub
End Sub

Private Sub txtPharName_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    frmPatPharm.piPatientID = piPatientID
    Get_PatientClinical
    frmPatPharm.Show vbModal
    Exit Sub
End Sub


Private Sub txtPharPhone_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    frmPatPharm.piPatientID = piPatientID
    frmPatPharm.Show
    Exit Sub
End Sub

Private Sub Get_PatientClinical()
Dim i%
Dim strSQL As String
Dim lngrecs As Long
Dim booRet As Boolean
Dim intActX_ID As Integer
Dim strPharmName As String
Dim strPharmPhone As String
Dim strLastPapSmear As String

    
    Screen.MousePointer = vbHourglass
    strPharmName = Get_ProfilerNote("FN_PP_PS", "PatPharmacy", "txtAddr", 1, piPatientID)
    
    If strPharmName <> vbNullString Then
        txtPharName.Text = strPharmName
    End If
    
    strPharmName = Get_ProfilerNote("FN_PP_PS", "PatPharmacy", "ccPhone", 1, piPatientID)
    If strPharmName <> vbNullString Then
        txtPharPhone.Text = strPharmName
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Function Look_UpPatient(vstrData As String, Optional vstrType As String)
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim cnPhys As New ADODB.Connection
Dim rsPhys As New ADODB.Recordset
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim strSuffix As String
Dim lngRefID As Long
Dim strExtID As String
Dim strName As String
Dim strSearch As String
Dim lngInvID As Long
Dim lngPatID As Long
Dim Arr
Dim strFirst As String
Dim strTitle As String

    On Error GoTo Error_Look_UpPatient
    Get_Edi_Static
    If clsEDI_Client.Site_ID = "1999" Then
        If TextFound("Supp", PsUserID, Len("Supp"), True) Then
            mbooSupport = True
            PT_Support_Search
        Else
            mbooSupport = False
            PT_Support_Search
        End If
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    LockWindowUpdate Hwnd
    lstPatient.Clear
    If Trim(vstrData) = vbNullString Then
        strSearch = Trim(txtSearch.Text)
    Else
        strSearch = Trim(vstrData)
    End If
    If TextFound("'", strSearch, Len("'"), True) Then
        strSearch = Replace(strSearch, "'", "%")
        strSearch = Replace(strSearch, " ", "%")
    End If
    Clear_UdtPatient
    clsListSearch.SearchType = Trim(cboSearchFilter.Text)
    If cnPhys.State = adStateClosed Then cnPhys.open psConnect(0)
    Select Case Trim(UCase(clsListSearch.SearchType))
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
                 strSQL = "Search_Patient_FirstName_Last_Name   '" & strSearch & "%', '%" & strFirst & "%'"
            Else
                strSQL = "Search_Patient_Row_Count_Last_Name 25, '" & strSearch & "%'"
            End If
        Case "EMR ID"
            strSQL = "Search_Patient_EMR_ID '" & strSearch & "%'"
        Case "SOCIAL SECURITY"
            strSQL = "Search_Patient_Social_Security '" & strSearch & "%'"
        Case "EXTERNAL ID"
            strSQL = "Search_Patient_External_ID '" & strSearch & "%'"
        Case "INVOICE NO"
            strSQL = "Search_Patient_Invoice_Invoice_No '" & strSearch & "%'"
            Set rsPat = cnPat.Execute(strSQL)
            If Not rsPat.EOF Then
                lngPatID = rsPat.Fields("ACCOUNT_NO")
            End If
            If lngPatID = 0 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            strSQL = "Get_Patient_BY_Id " & lngPatID
    End Select
    If strSQL <> vbNullString Then
        If cnPhys.State = adStateClosed Then cnPhys.open psConnect(0)
        Set rsPhys = cnPhys.Execute(strSQL)
         With rsPhys
             Do Until .EOF
                lngRefID = 0
                Clear_UdtPatient
                lngRefID = 0
                lngRefID = (!Patient_ID)
                If Not IsNull(!PATIENT_EXT_REF) Then
                    strExtID = Trim((!PATIENT_EXT_REF))
                End If
                rsPhys.MoveNext
                Get_PatientName lngRefID
                If Trim(udtPatientName.Suffix) <> vbNullString Then
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                Else
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                End If
                Select Case UCase(clsListSearch.SearchType)
                  Case "LAST NAME"
                      If Trim(udtPatientName.Name) <> vbNullString And Len(udtPatientName.Name) > 1 Then
                          lstPatient.AddItem Trim(udtPatientName.Name)
                          If lngRefID > 0 Then lstPatient.ItemData(lstPatient.NewIndex) = lngRefID
                      End If
                  Case "EMR ID"
                      If Trim(strName) <> vbNullString And Len(strName) > 1 Then
                          lstPatient.AddItem Trim(strName)
                          If lngRefID > 0 Then lstPatient.ItemData(lstPatient.NewIndex) = lngRefID
                      ElseIf Len(Trim(udtPatientName.Name)) > 1 Then
                          lstPatient.AddItem Trim(udtPatientName.Name)
                          If lngRefID > 0 Then lstPatient.ItemData(lstPatient.NewIndex) = lngRefID
                      End If
                  Case "EXTERNAL ID"
                      If Trim(udtPatientName.Name) <> vbNullString And Trim(strExtID) <> vbNullString And Trim(strExtID) <> "0" Then
                          lstPatient.AddItem Trim(udtPatientName.Name) & " - " & strExtID
                          If lngRefID > 0 Then lstPatient.ItemData(lstPatient.NewIndex) = lngRefID
                      End If
                End Select
             Loop
          End With
    End If
    If cnPat.State = adStateOpen Then
        cnPat.Close
        Set cnPat = Nothing
        Set rsPat = Nothing
    End If
    If cnPhys.State = adStateOpen Then
        cnPhys.Close
        Set cnPhys = Nothing
        Set rsPhys = Nothing
        Set Flds = Nothing
    End If
    Clear_UdtPatient
    LockWindowUpdate 0
    Screen.MousePointer = vbDefault
    Exit Function
Error_Look_UpPatient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Look_UpPatient", App.EXEName & "." & Me.Name
    Exit Function
    Resume
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
Dim lngDupPat As Long

    On Error GoTo Error_PT_Support_Search
    Screen.MousePointer = vbHourglass
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    If txtSearch.Text = vbNullString Then
        MsgBox "Please enter a search item first.", vbInformation, "PT_Support Patient Search"
        Exit Function
    End If
    lstPatient.Clear
    strSearch = txtSearch.Text
    If TextFound("'", strSearch, Len("'"), True) Then
        strSearch = Replace(strSearch, "'", "%")
        strSearch = Replace(strSearch, " ", "%")
    End If
    Clear_UdtPatient
    clsListSearch.SearchType = Trim(cboSearchFilter.Text)
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    Select Case Trim(UCase(clsListSearch.SearchType))
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
                 strSQL = "SELECT * FROM PATIENT WHERE LAST_NAME LIKE '" & strSearch & "%' AND FIRST_NAME like '%" & strFirst & "%'"
                ' Search_Patient_FirstName_Last_Name   '" & strSearch & "%', '%" & strFirst & "%'"
            Else
                strSQL = "Search_Patient_Row_Count_Last_Name 25, '" & strSearch & "%'"
            End If
        Case "EMR ID"
            strSQL = "Search_Patient_EMR_ID '" & strSearch & "%'"
        Case "SOCIAL SECURITY"
            strSQL = "Search_Patient_Social_Security '" & strSearch & "%'"
        Case "EXTERNAL ID"
            strSQL = "Search_Patient_External_ID '" & strSearch & "%'"
        Case "INVOICE NO"
            strSQL = "Search_Patient_Invoice_Invoice_No '" & strSearch & "%'"
            Set rsPat = cnPat.Execute(strSQL)
            If Not rsPat.EOF Then
                lngPatID = rsPat.Fields("ACCOUNT_NO")
            End If
            If lngPatID = 0 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            strSQL = "Get_Patient_BY_Id " & lngPatID
    End Select

    If strSQL <> vbNullString Then
        Set rsPat = cnPat.Execute(strSQL)
        Do Until rsPat.EOF
            lngPatID = 0
            lngPatID = rsPat.Fields("PATIENT_ID")
            rsPat.MoveNext
            If lngDupPat <> lngPatID Then
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
                    lstPatient.AddItem Trim(udtPatientName.Name) & " -  " & clsSurgFac.Name
                    lstPatient.ItemData(lstPatient.NewIndex) = lngPatID
                End If
            End If
            lngDupPat = lngPatID
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

Private Sub Get_ProfilerData()

End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch) <> vbNullString Then
        Timer1.Enabled = False
        Timer1.Enabled = True
    End If
    Exit Sub
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim strList As String
'
'    strList = Trim(txtSearch.Text)
'    grdPatient.ClearRange 0, 1, grdPatient.MaxRows, grdPatient.MaxRows, True
'    grdPatient.MaxRows = 0
'    Clear_Fields
'    If Trim(strList) <> vbNullString And Len(strList) > 1 Then
'        Select Case cboSearchFilter.Text
'            Case "LAST NAME"
'                Look_UpPatient strList, cboSearchFilter.Text
'            Case "SOCIAL SECURITY", "EXTERNAL ID"
'                Look_UpPatient strList, cboSearchFilter.Text
'            Case "EMR ID", "FINANCIAL ACCT NO"
'                If IsNumeric(txtSearch.Text) And Len(txtSearch.Text) >= 2 Then
'                    If cboSearchFilter.Text = "Financial Acct No" Then
'                        Look_UpPatient strList, "FINANCIAL ACCT NO"
'                    Else
'                        Look_UpPatient strList, cboSearchFilter.Text
'                    End If
'                Else
'                    MsgBox cboSearchFilter.Text & " " & "requires numeric values!", vbInformation, "PatientTrac Patient Search"
'                    txtSearch.Text = vbNullString
'                    imgPatient.Visible = False
'                    lstPatient.Clear
'                    Screen.MousePointer = vbDefault
'                    Exit Sub
'                End If
'        End Select
'    End If
'     mbooByPass = False
'  Exit Sub
End Sub


