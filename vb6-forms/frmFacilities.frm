VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmFacilities 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Maintenance"
   ClientHeight    =   9015
   ClientLeft      =   5355
   ClientTop       =   2490
   ClientWidth     =   13965
   ControlBox      =   0   'False
   Icon            =   "frmFacilities.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab sstTab 
      Height          =   8925
      Left            =   0
      TabIndex        =   19
      Top             =   30
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   15743
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Facilities Maintenance"
      TabPicture(0)   =   "frmFacilities.frx":094A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ccFmain"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Facilities Demographics"
      TabPicture(1)   =   "frmFacilities.frx":0966
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ccfFFac"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame ccFmain 
         BackColor       =   &H00E0E0E0&
         Height          =   8235
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   13665
         Begin FPSpreadADO.fpSpread grdData 
            Height          =   7215
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   13395
            _Version        =   458752
            _ExtentX        =   23627
            _ExtentY        =   12726
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            GridColor       =   -2147483645
            OperationMode   =   2
            RowHeaderDisplay=   0
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   16761024
            SpreadDesigner  =   "frmFacilities.frx":0982
            UserResize      =   0
         End
         Begin PT_XP_Button.PT_XP cmdExit 
            Height          =   525
            Left            =   12390
            TabIndex        =   51
            Top             =   7560
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   926
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
            Picture         =   "frmFacilities.frx":0C75
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmFacilities.frx":154F
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdActivate 
            Height          =   555
            Left            =   120
            TabIndex        =   52
            Top             =   7560
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   979
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
            Picture         =   "frmFacilities.frx":21A1
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Activate"
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmFacilities.frx":2A7B
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdAdd 
            Height          =   555
            Left            =   3015
            TabIndex        =   53
            Top             =   7530
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   979
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
            Picture         =   "frmFacilities.frx":36CD
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Add"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmFacilities.frx":3FA7
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdUpdate 
            Height          =   555
            Left            =   5895
            TabIndex        =   54
            Top             =   7530
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Picture         =   "frmFacilities.frx":4BF9
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Update"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmFacilities.frx":54D3
            ShowFocus       =   -1  'True
         End
         Begin VB.CommandButton cmdHelp1 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13080
            Picture         =   "frmFacilities.frx":6125
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   7560
            Visible         =   0   'False
            Width           =   465
         End
      End
      Begin VB.Frame ccfFFac 
         BackColor       =   &H00E0E0E0&
         Height          =   6975
         Left            =   180
         TabIndex        =   20
         Top             =   990
         Width           =   13305
         Begin MSMask.MaskEdBox cctFEIN 
            Height          =   315
            Left            =   6120
            TabIndex        =   1
            Top             =   780
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "##-#######"
            PromptChar      =   "_"
         End
         Begin VB.ListBox lstProviders 
            Height          =   2310
            Left            =   9660
            Style           =   1  'Checkbox
            TabIndex        =   48
            Top             =   1410
            Width           =   3045
         End
         Begin VB.ComboBox cboFacType 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   6120
            TabIndex        =   0
            Top             =   330
            Width           =   2565
         End
         Begin VB.TextBox txtFacilityID 
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2880
            TabIndex        =   46
            Top             =   360
            Width           =   1155
         End
         Begin VB.CommandButton cmdHelp2 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   12720
            Picture         =   "frmFacilities.frx":6227
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   6360
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.ComboBox cccState 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   2910
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3720
            Width           =   2385
         End
         Begin VB.TextBox txtZipcode 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2910
            MaxLength       =   9
            TabIndex        =   9
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox txtCity 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2910
            MaxLength       =   32
            TabIndex        =   7
            Top             =   3150
            Width           =   5925
         End
         Begin VB.TextBox txtAddress2 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2910
            MaxLength       =   48
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   2580
            Width           =   5955
         End
         Begin VB.TextBox txtAddress1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2910
            MaxLength       =   48
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   2040
            Width           =   5955
         End
         Begin VB.TextBox txtFacilityName 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2910
            MaxLength       =   48
            TabIndex        =   4
            Top             =   1440
            Width           =   5925
         End
         Begin VB.TextBox txtPhone23 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4050
            MaxLength       =   4
            TabIndex        =   18
            Top             =   5190
            Width           =   645
         End
         Begin VB.TextBox txtPhone21 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2910
            MaxLength       =   3
            TabIndex        =   11
            Top             =   5190
            Width           =   465
         End
         Begin VB.TextBox txtFacsimile13 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4050
            MaxLength       =   4
            TabIndex        =   16
            Top             =   5640
            Width           =   645
         End
         Begin VB.TextBox txtFacsimile12 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   3450
            MaxLength       =   3
            TabIndex        =   15
            Top             =   5640
            Width           =   465
         End
         Begin VB.TextBox txtFacsimile11 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2910
            MaxLength       =   3
            TabIndex        =   12
            Top             =   5640
            Width           =   465
         End
         Begin VB.TextBox txtPhone11 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2910
            MaxLength       =   3
            TabIndex        =   10
            Top             =   4740
            Width           =   465
         End
         Begin VB.TextBox txtPhone12 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   3450
            MaxLength       =   3
            TabIndex        =   13
            Top             =   4740
            Width           =   465
         End
         Begin VB.TextBox txtPhone13 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4050
            MaxLength       =   4
            TabIndex        =   14
            Top             =   4740
            Width           =   645
         End
         Begin VB.TextBox txtPhone22 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   3450
            MaxLength       =   3
            TabIndex        =   17
            Top             =   5190
            Width           =   465
         End
         Begin VB.Frame fraHospitalInd 
            BackColor       =   &H00E0E0E0&
            Height          =   525
            Left            =   11340
            TabIndex        =   33
            Top             =   5490
            Width           =   1695
            Begin VB.OptionButton optHospitalIndNo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "No"
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
               Height          =   195
               Left            =   900
               TabIndex        =   3
               Top             =   180
               Value           =   -1  'True
               Width           =   585
            End
            Begin VB.OptionButton optHospitalIndYes 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Yes"
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
               Height          =   195
               Left            =   90
               TabIndex        =   2
               Top             =   180
               Width           =   825
            End
         End
         Begin PT_XP_Button.PT_XP cmdSave 
            Height          =   435
            Left            =   5190
            TabIndex        =   55
            Top             =   6300
            Width           =   1065
            _ExtentX        =   1879
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
            Picture         =   "frmFacilities.frx":6329
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Save"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmFacilities.frx":66C3
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdCancel 
            Height          =   435
            Left            =   6390
            TabIndex        =   56
            Top             =   6300
            Width           =   1065
            _ExtentX        =   1879
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
            Picture         =   "frmFacilities.frx":6A15
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmFacilities.frx":7427
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblFEINNo 
            BackStyle       =   0  'Transparent
            Caption         =   "Federal Tax ID:"
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
            Left            =   4560
            TabIndex        =   50
            Top             =   780
            Width           =   1485
         End
         Begin VB.Label lblProviders 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Providers :"
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
            Left            =   9660
            TabIndex        =   49
            Top             =   1170
            Width           =   1050
         End
         Begin VB.Label lblFacType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Facility Type:"
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
            Left            =   4680
            TabIndex        =   47
            Top             =   390
            Width           =   1350
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "("
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
            Index           =   0
            Left            =   2850
            TabIndex        =   42
            Top             =   4800
            Width           =   75
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ")"
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
            Index           =   1
            Left            =   3360
            TabIndex        =   41
            Top             =   4800
            Width           =   75
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Index           =   2
            Left            =   3960
            TabIndex        =   40
            Top             =   4830
            Width           =   60
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "("
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
            Index           =   3
            Left            =   2850
            TabIndex        =   39
            Top             =   5700
            Width           =   75
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ")"
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
            Index           =   4
            Left            =   3360
            TabIndex        =   38
            Top             =   5700
            Width           =   75
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Index           =   5
            Left            =   3930
            TabIndex        =   37
            Top             =   5700
            Width           =   60
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "("
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
            Index           =   6
            Left            =   2850
            TabIndex        =   36
            Top             =   5280
            Width           =   75
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ")"
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
            Index           =   7
            Left            =   3360
            TabIndex        =   35
            Top             =   5280
            Width           =   75
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Index           =   8
            Left            =   3960
            TabIndex        =   34
            Top             =   5280
            Width           =   60
         End
         Begin VB.Label lblFacilityID1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Facility ID :"
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
            Left            =   1590
            TabIndex        =   31
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label lblFaciltyName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Facility Name :"
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
            Left            =   1305
            TabIndex        =   30
            Top             =   1500
            Width           =   1485
         End
         Begin VB.Label lblAdd1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address 1 :"
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
            Left            =   1740
            TabIndex        =   29
            Top             =   2145
            Width           =   1020
         End
         Begin VB.Label lblAdd2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address 2 :"
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
            Left            =   1740
            TabIndex        =   28
            Top             =   2670
            Width           =   1035
         End
         Begin VB.Label lblCity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City :"
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
            Left            =   2250
            TabIndex        =   27
            Top             =   3285
            Width           =   510
         End
         Begin VB.Label lblState 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "State :"
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
            Left            =   2115
            TabIndex        =   26
            Top             =   3810
            Width           =   615
         End
         Begin VB.Label lblZip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zipcode :"
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
            Left            =   1920
            TabIndex        =   25
            Top             =   4320
            Width           =   840
         End
         Begin VB.Label lblPhone2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone 2 :"
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
            Left            =   1455
            TabIndex        =   24
            Top             =   5250
            Width           =   1275
         End
         Begin VB.Label lblPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone 1 :"
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
            Left            =   1455
            TabIndex        =   23
            Top             =   4860
            Width           =   1260
         End
         Begin VB.Label lblFacs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Facsimile :"
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
            Left            =   1710
            TabIndex        =   22
            Top             =   5700
            Width           =   1050
         End
         Begin VB.Label lblHosIND 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hospital :"
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
            Left            =   10410
            TabIndex        =   21
            Top             =   5670
            Visible         =   0   'False
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "frmFacilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_lCurrRow As Long
Dim m_lMode As Integer
Dim fbAdd As Boolean

Private Sub cmdActivate_Click()
Dim sValue
Dim i As Integer

    On Error GoTo cmdActivate_Error:

    If grdData.MaxRows < 1 Then Exit Sub
    
    If cmdActivate.Caption = "Activate" Then
       grdData.SetText 6, grdData.ActiveRow, "Yes"
       cmdActivate.Caption = "DeActivate"
       i = 1
    Else
       grdData.SetText 6, grdData.ActiveRow, "No"
       cmdActivate.Caption = "Activate"
       i = 0
    End If
    
    grdData.GetText 0, grdData.ActiveRow, sValue
    
   Open_DB psConnect(0)
    
    cn.u_Facility_Active i, sValue

    Close_DB psConnect(0)
    Exit Sub
    
cmdActivate_Error:
    ErrHandler _
     "frmFacilities.cmdActivate_Click", _
     Err.Number, _
     Err.Description
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo cmdAdd_Error:
    
    ccfFFac.Enabled = True
    ClearTextValues
    Load_Providers
    sstTab.TabEnabled(0) = False
    sstTab.TabEnabled(1) = True
    sstTab.Tab = 1
    fbAdd = True
    
    Exit Sub
    
cmdAdd_Error:
    ErrHandler _
     "frmFacilities.cmdAdd_Click", _
     Err.Number, _
     Err.Description
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo Error_Handler
    
    sstTab.Tab = 0
    sstTab.TabEnabled(1) = False
    sstTab.TabEnabled(0) = True
    ccfFFac.Enabled = False
    
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmFacilities.cmdCancel_Click", _
     Err.Number, _
     Err.Description
     Exit Sub
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Not Validate_Entries Then Exit Sub
   
    If SaveData = True Then
       sstTab.Tab = 0
       ccfFFac.Enabled = False
    End If
End Sub

Private Sub cmdUpdate_Click()

    On Error GoTo cmdUpdate_Error:
    
    If grdData.MaxRows = 0 Then Exit Sub
    
    ccfFFac.Enabled = True
    ClearTextValues
    LoadTextValues
    Load_Providers
    sstTab.TabEnabled(0) = False
    sstTab.TabEnabled(1) = True
    sstTab.Tab = 1
    fbAdd = False
    
    Exit Sub
    
cmdUpdate_Error:
    ErrHandler _
     "frmFacilities.cmdUpdate_Click", _
     Err.Number, _
     Err.Description
End Sub

Private Sub Load_Service()
Dim strSQL As String
Dim lvarRS
Dim lbooRet As Boolean
Dim lintRows As Long
Dim lstrRemarks As String
Dim r%


On Error GoTo Error_Load_Service

    Open_DB psConnect(0)
    strSQL = "Select DISTINCT ENCOUNTER_TYPE, ENCTRTYPE_GROUP_ID " _
        & "From HCFA_CODES "
                
    Set rs = cn.Execute(strSQL)
    With rs
        Do Until .EOF
            If Not IsNull((!ENCOUNTER_TYPE)) Then cboFacType.AddItem (!ENCOUNTER_TYPE)
            If IsNumeric((!ENCTRTYPE_GROUP_ID)) Then cboFacType.ItemData(cboFacType.NewIndex) = (!ENCTRTYPE_GROUP_ID)
        .MoveNext
        Loop
    End With
    Close_DB psConnect(0)
    Exit Sub
    
Error_Load_Service:
    ErrHandler _
     Me.Name & " Sub Load_Service", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub



Private Sub Form_Load()
        
    Screen.MousePointer = vbHourglass
    grdData.ZOrder 0
    sstTab.Tab = 0
    sstTab.TabEnabled(1) = False
    Format_Grid
    LoadGrid
    Load_States
    Load_Service
    ccfFFac.Enabled = False
    centerForm Me
    Me.Top = Me.Top + 300
    
    Screen.MousePointer = vbDefault
    
End Sub
Private Function LoadGrid() As Boolean

    Dim sValue
    
    On Error GoTo LoadGrid_Error:

    Open_DB psConnect(0)
    
    cn.s_Facilities rs
    
    With rs
       Do Until .EOF
          grdData.MaxRows = grdData.MaxRows + 1
          grdData.Row = grdData.MaxRows
          grdData.SetText 0, grdData.Row, (!FACILITY_ID)
          grdData.Col = 1
          grdData.CellType = CellTypeStaticText
          If (!Hospital_Ind) Then grdData.SetText 1, grdData.Row, "Yes" Else grdData.SetText 6, grdData.Row, "No"
          grdData.Col = 2
          grdData.CellType = CellTypeStaticText
          grdData.SetText 2, grdData.Row, (!Facility_Name) & ""
          grdData.Col = 3
          grdData.CellType = CellTypeStaticText
          grdData.SetText 3, grdData.Row, (!City) & ""
          grdData.Col = 4
          grdData.CellType = CellTypeStaticText
          grdData.SetText 4, grdData.Row, (!State) & ""
          grdData.Col = 5
          grdData.CellType = CellTypeStaticText
          grdData.SetText 5, grdData.Row, (!Telephone_1) & ""
          grdData.Col = 6
          grdData.CellType = CellTypeStaticText
          If (!Active) Then grdData.SetText 6, grdData.Row, "Yes" Else grdData.SetText 6, grdData.Row, "No"
          
          .MoveNext
       Loop
    End With
    
    Close_DB psConnect(0)
    
    If grdData.MaxRows > 0 Then
       grdData.Row = 1
       grdData.GetText 6, 1, sValue
       If sValue = "Yes" Then cmdActivate.Caption = "DeActivate" Else cmdActivate.Caption = "Activate"
       cmdActivate.Visible = True
    End If

    Exit Function

LoadGrid_Error:
    ErrHandler _
     "frmFacilities.LoadGrid", _
     Err.Number, _
     Err.Description
End Function

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
  Exit Sub
End Sub

Private Sub grdData_GotFocus()
    Dim sValue
    
    On Error GoTo Error_Handler
    
    grdData.GetText 6, grdData.ActiveRow, sValue
    grdData.Row = grdData.ActiveRow
    
    If sValue = "Yes" Then cmdActivate.Caption = "DeActivate" Else cmdActivate.Caption = "Activate"
    If cmdActivate.Visible = False Then cmdActivate.Visible = True
    
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmFacilities.grdData_GotFocus", _
     Err.Number, _
     Err.Description
End Sub
Private Sub grdData_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sValue
    
    On Error GoTo Error_Handler
    
    grdData.GetText 6, grdData.ActiveRow, sValue
    
    If sValue = "Yes" Then cmdActivate.Caption = "DeActivate" Else cmdActivate.Caption = "Activate"
    If cmdActivate.Visible = False Then cmdActivate.Visible = True
    
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmFacilities.grdData_KeyUp", _
     Err.Number, _
     Err.Description
End Sub
Private Sub grdData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sValue
    
    On Error GoTo Error_Handler
    
    grdData.GetText 6, grdData.ActiveRow, sValue
    
    If sValue = "Yes" Then cmdActivate.Caption = "DeActivate" Else cmdActivate.Caption = "Activate"
    If cmdActivate.Visible = False Then cmdActivate.Visible = True
    
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmFacilities.grdData_MouseUp", _
     Err.Number, _
     Err.Description
End Sub
Private Sub LoadTextValues()
Dim sPhone1 As String
Dim varPhone
Dim vFacilityID
Dim Idx As Integer
Dim strSQL As String
Dim lvarRS
Dim lbooRet As Boolean
Dim lintRows As Long
Dim lstrRemarks As String
Dim r%
Dim i%
Dim lngFacilityID As Long

    On Error GoTo LoadTextValues_Error:
    
    grdData.GetText 0, grdData.ActiveRow, vFacilityID
    txtFacilityID.Text = Trim(vFacilityID)
    
     Open_DB psConnect(0)
    
    cn.s_FacilityInfo vFacilityID, rs
    
    With rs
       Do While Not .EOF
          If (!Hospital_Ind) Then optHospitalIndYes.Value = True Else optHospitalIndNo.Value = True
          cctFEIN.Text = Trim((!FEIN)) & ""
          txtFacilityName.Text = Trim((!Facility_Name)) & ""
          txtAddress1.Text = Trim((!ADDRESS_1)) & ""
          txtAddress2.Text = Trim((!ADDRESS_2)) & ""
          txtCity.Text = Trim((!City)) & ""

          For Idx = 0 To cccState.ListCount - 1
              If cccState.List(Idx) = Trim((!State)) Then
                 cccState.ListIndex = Idx
                 Exit For
              End If
          Next Idx
          
         If Not IsNull(!ZIP_CODE) Then txtZipcode.Text = Trim((!ZIP_CODE))
          
          If InStr((!Telephone_1), "-") Then
              varPhone = Split((!Telephone_1), "-")
              txtPhone11.Text = varPhone(0)
              txtPhone12.Text = varPhone(1)
              txtPhone13.Text = varPhone(2)
              Erase varPhone
          Else
             sPhone1 = Trim((!Telephone_1))
             If Len(sPhone1) > 0 Then
                txtPhone11.Text = Left(sPhone1, 3)
                If Len(sPhone1) > 3 Then
                   sPhone1 = Mid(sPhone1, 4)
                   txtPhone12.Text = Left(sPhone1, 3)
                   If Len(sPhone1) > 3 Then
                      txtPhone13.Text = Mid(sPhone1, 4)
                   End If
                End If
             End If
          End If
          
          If InStr((!Telephone_2), "-") Then
              varPhone = Split((!Telephone_2), "-")
              txtPhone21.Text = varPhone(0)
              txtPhone22.Text = varPhone(1)
              txtPhone23.Text = varPhone(2)
              Erase varPhone
          Else
             If Not IsNull(!Telephone_2) Then
                sPhone1 = Trim((!Telephone_2))
                If Len(sPhone1) > 0 Then
                    txtPhone12.Text = Left(sPhone1, 3)
                    If Len(sPhone1) > 3 Then
                       sPhone1 = Mid(sPhone1, 4)
                       txtPhone22.Text = Left(sPhone1, 3)
                       If Len(sPhone1) > 3 Then
                          txtPhone23.Text = Mid(sPhone1, 4)
                       End If
                    End If
                End If
            End If
          End If
          
          If InStr((!Facsimile), "-") Then
              varPhone = Split((!Facsimile), "-")
              txtFacsimile11.Text = varPhone(0)
              txtFacsimile12.Text = varPhone(1)
              txtFacsimile13.Text = varPhone(2)
              Erase varPhone
          Else
            If Not IsNull(!Facsimile) Then
             sPhone1 = Trim((!Facsimile))
             If Len(sPhone1) > 0 Then
                txtFacsimile12.Text = Left(sPhone1, 3)
                If Len(sPhone1) > 3 Then
                   sPhone1 = Mid(sPhone1, 4)
                   txtFacsimile12.Text = Left(sPhone1, 3)
                   If Len(sPhone1) > 3 Then
                      txtFacsimile13.Text = Mid(sPhone1, 4)
                   End If
                End If
             End If
            End If
          End If
          
          .MoveNext
       Loop
    End With

        lstProviders.Clear
        lngFacilityID = vFacilityID
        strSQL = " SELECT distinct provider_id FROM facility_providers WHERE facility_id  = " & lngFacilityID
        Set rs = cn.Execute(strSQL)
        
        With rs
            Do Until .EOF
                    For i = 0 To lstProviders.ListCount - 1
                        If lstProviders.ItemData(i) = (!PROVIDER_ID) Then
                            lstProviders.Selected(i) = True
                        End If
                    Next i
            .MoveNext
            Loop
        End With
   Close_DB psConnect(0)
    Exit Sub
     Screen.MousePointer = vbDefault
LoadTextValues_Error:
    ErrHandler _
     "frmFacilities.LoadTextValues", _
     Err.Number, _
     Err.Description
End Sub

Private Sub ClearTextValues()
    On Error GoTo ClearTextValues_Error:
    txtFacilityID.Text = vbNullString
    cccState.ListIndex = -1
    txtFacilityName.Text = vbNullString
    txtAddress1.Text = vbNullString
    txtAddress2.Text = vbNullString
    txtCity.Text = vbNullString
    txtZipcode.Text = vbNullString
    txtPhone11.Text = vbNullString
    txtPhone12.Text = vbNullString
    txtPhone13.Text = vbNullString
    txtPhone21.Text = vbNullString
    txtPhone22.Text = vbNullString
    txtPhone23.Text = vbNullString
    txtFacsimile11.Text = vbNullString
    txtFacsimile12.Text = vbNullString
    txtFacsimile13.Text = vbNullString
    optHospitalIndYes.Value = True
    optHospitalIndNo.Value = False
    cctFEIN.Text = vbNullString
    Exit Sub
    
ClearTextValues_Error:
    ErrHandler _
     "frmFacilities.ClearTextValues", _
     Err.Number, _
     Err.Description
End Sub

Private Function SaveData() As Boolean
Dim strSQL As String
Dim bHospitalIndYesNo
Dim strPractice As String
Dim strFacility As String
Dim sPhone1 As String
Dim sPhone2 As String
Dim sFacs As String
Dim Idx As Integer
Dim booRet As Boolean
Dim lngrecs As Long
Dim varRS
Dim lngFacilityID As Long
Dim r%
    
    On Error GoTo SaveData_Error:
    
    Screen.MousePointer = vbHourglass
    
    If optHospitalIndYes.Value = True Then
       bHospitalIndYesNo = 1
    ElseIf optHospitalIndNo.Value = True Then
       bHospitalIndYesNo = 0
    End If
        
    strFacility = Replace(txtFacilityName.Text, "'", "''")
        
    sPhone1 = Trim(txtPhone11.Text) & "-" & Trim(txtPhone12.Text) & "-" & Trim(txtPhone13.Text)
    sPhone2 = Trim(txtPhone21.Text) & "-" & Trim(txtPhone22.Text) & "-" & Trim(txtPhone23.Text)
    sFacs = Trim(txtFacsimile11.Text) & "-" & Trim(txtFacsimile12.Text) & "-" & Trim(txtFacsimile13.Text)

   Open_DB psConnect(0)
    
    If fbAdd Then
       cn.i_Facility "", 0, _
                     Trim(strFacility), _
                     Trim(Replace(cctFEIN.Text, "'", "''")), _
                     Trim(Replace(txtAddress1.Text, "'", "''")), _
                     Trim(Replace(txtAddress2.Text, "'", "''")), _
                     Trim(Replace(txtCity.Text, "'", "''")), _
                     Trim(Replace(cccState.Text, "'", "''")), _
                     Trim(Replace(txtZipcode.Text, "'", "''")), _
                     Trim(Replace(sPhone1, "'", "''")), _
                     Trim(Replace(sPhone2, "'", "''")), _
                     Trim(Replace(sFacs, "'", "''")), _
                     "", _
                     bHospitalIndYesNo, _
                     0, _
                     Trim(Replace(cboFacType.Text, "'", "' '")), _
                     ""

       grdData.MaxRows = grdData.MaxRows + 1
       grdData.Row = grdData.MaxRows
       
       If bHospitalIndYesNo Then
          grdData.SetText 1, grdData.Row, "Yes"
       Else
          grdData.SetText 1, grdData.Row, "No"
       End If
       
       ' Get New Facility ID
       strSQL = "Select facility_id from Facilities where facility_name = '" & strFacility & "'"
        Set rs = cn.Execute(strSQL)
        
       With rs
            If Not .EOF Then
                lngFacilityID = (!FACILITY_ID)
            End If
       End With
        
            For r = 0 To lstProviders.ListCount - 1
                If lstProviders.Selected(r) Then
                    'Is it already there
                        cn.S_FACILITIES_PROVIDERS lstProviders.ItemData(r)
                   With rs
                        If Not .EOF Then
                            cn.i_Facility_Providers lngFacilityID, lstProviders.ItemData(r)
                        End If
                    End With
                End If
            Next r


       grdData.SetText 2, grdData.Row, strFacility
       grdData.SetText 3, grdData.Row, Trim(Replace(txtCity.Text, "'", "''"))
       grdData.SetText 4, grdData.Row, Trim(Replace(cccState.Text, "'", "''"))
       grdData.SetText 5, grdData.Row, Trim(Replace(sPhone1, "'", "''"))
       grdData.SetText 6, grdData.Row, "Yes"
       grdData.SetText 7, grdData.Row, Trim(Replace(cboFacType.Text, "'", "' '"))
       
       cn.s_FacilityID strFacility, rs
       With rs
          If Not .EOF Then
             grdData.SetText 0, grdData.Row, (!FacilityID)
          Else
             grdData.SetText 0, grdData.Row, "???"
             MsgBox "Problem retrieving Facility ID", vbExclamation
             Exit Function
          End If
          .Close
       End With
       
       For Idx = 1 To 7
           grdData.Col = Idx
           grdData.CellType = CellTypeStaticText
       Next Idx
       
    Else
       cn.u_Facility Trim(strFacility), _
                     Trim(Replace(cctFEIN.Text, "'", "''")), _
                     Trim(Replace(txtAddress1.Text, "'", "''")), _
                     Trim(Replace(txtAddress2.Text, "'", "''")), _
                     Trim(Replace(txtCity.Text, "'", "''")), _
                     Trim(Replace(cccState.Text, "'", "''")), _
                     Trim(Replace(txtZipcode.Text, "'", "''")), _
                     Trim(Replace(sPhone1, "'", "''")), _
                     Trim(Replace(sPhone2, "'", "''")), _
                     Trim(Replace(sFacs, "'", "''")), _
                     bHospitalIndYesNo, _
                     txtFacilityID.Text, _
                     Trim(Replace(cboFacType.Text, "'", "' '"))
   
       grdData.SetText 2, grdData.ActiveRow, strFacility
       grdData.SetText 3, grdData.ActiveRow, Trim(Replace(txtCity.Text, "'", "''"))
       grdData.SetText 4, grdData.ActiveRow, Trim(Replace(cccState.Text, "'", "''"))
       grdData.SetText 5, grdData.ActiveRow, Trim(Replace(sPhone1, "'", "''"))
    End If
    
    ' Need to Delete a Provider from Facility
    'cn.d_Facility_Providers vProvID
    'sProviderID = vProvID
    'Find provider
    
    Close_DB psConnect(0)

    Screen.MousePointer = vbDefault
    fbAdd = False
    SaveData = True
    
    Exit Function
    
SaveData_Error:
    Screen.MousePointer = vbDefault
    SaveData = False
    
    ErrHandler _
     "frmFacilities.SaveData", _
     Err.Number, _
     Err.Description
End Function


Private Sub txtFacsimile11_KeyPress(KeyAscii As Integer)
   If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txtFacsimile11_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode >= 32 And KeyCode <= 127 Then
      If Len(txtFacsimile11.Text) = txtFacsimile11.MaxLength Then
         If txtFacsimile12.Enabled Then txtFacsimile12.SetFocus
      End If
   End If
End Sub

Private Sub txtFacsimile12_KeyPress(KeyAscii As Integer)
   If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
    
End Sub

Private Sub txtFacsimile12_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode >= 32 And KeyCode <= 127 Then
      If Len(txtFacsimile12.Text) = txtFacsimile12.MaxLength Then
         If txtFacsimile13.Enabled Then txtFacsimile13.SetFocus
      End If
   End If
End Sub

Private Sub txtFacsimile13_KeyPress(KeyAscii As Integer)
    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub
Private Sub txtPhone11_KeyPress(KeyAscii As Integer)
    
   If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txtPhone11_KeyUp(KeyCode As Integer, Shift As Integer)
    
   If KeyCode >= 32 And KeyCode <= 127 Then
      If Len(txtPhone11.Text) = txtPhone11.MaxLength Then
         If txtPhone12.Enabled Then txtPhone12.SetFocus
      End If
   End If
    
End Sub

Private Sub txtPhone12_KeyPress(KeyAscii As Integer)
    
   If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txtPhone12_KeyUp(KeyCode As Integer, Shift As Integer)
    
   If KeyCode >= 32 And KeyCode <= 127 Then
      If Len(txtPhone12.Text) = txtPhone12.MaxLength Then
         If txtPhone13.Enabled Then txtPhone13.SetFocus
      End If
   End If
    
End Sub

Private Sub txtPhone13_KeyPress(KeyAscii As Integer)
    
    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
    
End Sub
Private Sub txtPhone21_KeyPress(KeyAscii As Integer)
    
   If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
    
End Sub

Private Sub txtPhone21_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode >= 32 And KeyCode <= 127 Then
      If Len(txtPhone21.Text) = txtPhone21.MaxLength Then
         If txtPhone22.Enabled Then txtPhone22.SetFocus
      End If
   End If
End Sub

Private Sub txtPhone22_KeyPress(KeyAscii As Integer)
    
   If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txtPhone22_KeyUp(KeyCode As Integer, Shift As Integer)
    
   If KeyCode >= 32 And KeyCode <= 127 Then
      If Len(txtPhone22.Text) = txtPhone22.MaxLength Then
         If txtPhone23.Enabled Then txtPhone23.SetFocus
      End If
   End If
    
End Sub

Private Sub txtPhone23_KeyPress(KeyAscii As Integer)
    
    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
    
End Sub

Private Sub Format_Grid()
Dim Idx As Integer
Dim strSQL As String

    
    On Error GoTo Err_Format_Grid

    Screen.MousePointer = vbHourglass

    With grdData
        .ColsFrozen = 1
        .MaxCols = 7
        .MaxRows = 0
        .CursorStyle = CursorStyleArrow
        
        .ColWidth(0) = 4
        .SetText 0, 0, "ID"
        
        .ColWidth(1) = 8
        .SetText 1, 0, "Hospital"
       
        .ColWidth(2) = 30
        .SetText 2, 0, "Facility Name"
        
        .ColWidth(3) = 25
        .SetText 3, 0, "City"
        
        .ColWidth(4) = 20
        .SetText 4, 0, "State"
        
        .ColWidth(5) = 12
        .SetText 5, 0, "Phone"
        
        .ColWidth(6) = 7
        .SetText 6, 0, "Active"
        
        .ColWidth(7) = 30
        .SetText 7, 0, "Facility Type"
    End With
    
    Screen.MousePointer = vbDefault

    Exit Sub

Err_Format_Grid:
    Screen.MousePointer = vbDefault

    ErrHandler _
     "frmFacilities.Format_Grid", _
     Err.Number, _
     Err.Description
End Sub
Private Sub Load_States()
Dim strSQL As String
Dim i%

    On Error GoTo Error_Load_States
    
    Open_DB psConnect(0)
    strSQL = "Select state_name  From States where state_ind = 'y' Order by state_name"
    Set rs = cn.Execute(strSQL)
    
    With rs
        Do Until .EOF
            cccState.AddItem Trim((!state_name))
      .MoveNext
       Loop
   End With
   Close_DB psConnect(0)
   
    Exit Sub
    
Error_Load_States:
    ErrHandler _
     "frmFacilities.Load_States", _
     Err.Number, _
     Err.Description
End Sub

Private Function Validate_Entries() As Boolean
    Validate_Entries = True
    
    If Trim(txtFacilityName) = "" Then
       MsgBox "Facility Name is Required!", vbInformation + vbOKCancel
       Validate_Entries = False
       Exit Function
    End If
End Function


Private Sub Load_Providers()
Dim strSQL As String
Dim r%
Dim i%
Dim vArray
Dim lngFacilityID As Long
Dim strFirstName As String
Dim strMiddleName As String
Dim strLastName As String
Dim strTitle As String
Dim strName As String


    On Error GoTo Error_Load_Providers
    
    Open_DB psConnect(0)
    
    strSQL = "Select  first_name, middle_name, last_name, title , provider_id " _
            & "From PROVIDERS  " _
            & "Where Title = 'M.D.' or Title = 'MD' or Title = 'RN' or Title = 'LPN' or Title = 'PA' Or Title = 'APRN' or Title = 'NP'"
            
    Set rs = cn.Execute(strSQL)
    With rs
        Do Until .EOF
            If Not IsNull((!First_Name)) Then strFirstName = (!First_Name)
            If Not IsNull((!Middle_Name)) Then strMiddleName = (!Middle_Name)
            If Not IsNull((!Last_Name)) Then strLastName = (!Last_Name)
            If Not IsNull((!Title)) Then strTitle = (!Title)
                        If strFirstName <> vbNullString Then strName = strFirstName
                        If strMiddleName <> vbNullString Then strName = strName & " " & strMiddleName
                        If strLastName <> vbNullString Then strName = strName & " " & strLastName
                        If strTitle <> vbNullString Then strName = strName & ", " & strTitle
                        lstProviders.AddItem strName
                        lstProviders.ItemData(lstProviders.NewIndex) = (!PROVIDER_ID)
        .MoveNext
        Loop
    End With
        

    Close_DB psConnect(0)
    Exit Sub
    
Error_Load_Providers:
    ErrHandler _
     "frmFacilities > Load_Providers", _
     Err.Number, _
     Err.Description
     Exit Sub
     Resume
End Sub


