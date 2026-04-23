VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmProviders 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provider Data Maintenance"
   ClientHeight    =   10395
   ClientLeft      =   4500
   ClientTop       =   1605
   ClientWidth     =   14370
   ControlBox      =   0   'False
   Icon            =   "frmProviders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   14370
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab sstTab 
      Height          =   10395
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   18336
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Provider Maintenance"
      TabPicture(0)   =   "frmProviders.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ccFmain"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Provider Demographics"
      TabPicture(1)   =   "frmProviders.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ccfFFac"
      Tab(1).ControlCount=   1
      Begin VB.Frame ccfFFac 
         BackColor       =   &H00E0E0E0&
         Height          =   9885
         Left            =   -74850
         TabIndex        =   7
         Top             =   420
         Width           =   14085
         Begin VB.Frame fraNumber 
            BackColor       =   &H00C0FFFF&
            Height          =   1005
            Left            =   3060
            TabIndex        =   123
            Top             =   8010
            Visible         =   0   'False
            Width           =   2025
            Begin MSComCtl2.DTPicker ccdDate 
               Height          =   315
               Left            =   240
               TabIndex        =   124
               Top             =   180
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   393216
               Format          =   85262337
               CurrentDate     =   39288
            End
            Begin PT_XP_Button.PT_XP cmdAccept 
               Height          =   285
               Left            =   510
               TabIndex        =   125
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
               Picture         =   "frmProviders.frx":0902
               AppearanceThemes=   3
               TransparentColor=   14737632
               Caption         =   "Close"
               DepthEvent      =   1
               ForeColorOver   =   8388608
               ForeColorFocus  =   8388608
               ForeColorDown   =   8388608
               PictureDisabled =   "frmProviders.frx":1314
               ShowFocus       =   -1  'True
            End
         End
         Begin VB.ListBox lstTaxonomy 
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
            Height          =   1620
            Left            =   10650
            TabIndex        =   121
            Top             =   6210
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.TextBox cctTaxonomyID 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10650
            MaxLength       =   25
            TabIndex        =   119
            Top             =   5520
            Width           =   2145
         End
         Begin VB.TextBox cctTaxName 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10650
            MaxLength       =   25
            TabIndex        =   117
            Top             =   4950
            Width           =   2145
         End
         Begin VB.TextBox txtUserId 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            MaxLength       =   12
            TabIndex        =   110
            Top             =   3975
            Width           =   2475
         End
         Begin VB.TextBox txtPassword1 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   5790
            MaxLength       =   12
            PasswordChar    =   "*"
            TabIndex        =   109
            Top             =   3930
            Width           =   2475
         End
         Begin VB.TextBox txtPassword2 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   11220
            MaxLength       =   12
            PasswordChar    =   "*"
            TabIndex        =   108
            Top             =   3930
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.TextBox cctCellPhone2 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11940
            MaxLength       =   4
            TabIndex        =   103
            Top             =   3000
            Width           =   645
         End
         Begin VB.TextBox cctCellArea 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10770
            MaxLength       =   3
            TabIndex        =   102
            Top             =   3000
            Width           =   465
         End
         Begin VB.TextBox cctCellPhone1 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11310
            MaxLength       =   3
            TabIndex        =   101
            Top             =   3000
            Width           =   465
         End
         Begin VB.TextBox cctFaxArea 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6720
            MaxLength       =   3
            TabIndex        =   100
            Top             =   3000
            Width           =   465
         End
         Begin VB.TextBox cctPagePhone2 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7890
            MaxLength       =   4
            TabIndex        =   91
            Top             =   3420
            Width           =   645
         End
         Begin VB.TextBox cctFaxPhone2 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7890
            MaxLength       =   4
            TabIndex        =   90
            Top             =   3000
            Width           =   645
         End
         Begin VB.TextBox cctFaxPhone1 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7260
            MaxLength       =   3
            TabIndex        =   89
            Top             =   3000
            Width           =   465
         End
         Begin VB.TextBox cctPagePhone1 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7260
            MaxLength       =   3
            TabIndex        =   88
            Top             =   3420
            Width           =   465
         End
         Begin VB.TextBox cctPageArea 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6720
            MaxLength       =   3
            TabIndex        =   87
            Top             =   3420
            Width           =   465
         End
         Begin VB.TextBox cctOfficePhone2 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2700
            MaxLength       =   4
            TabIndex        =   78
            Top             =   3000
            Width           =   645
         End
         Begin VB.TextBox cctOfficeArea 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1530
            MaxLength       =   3
            TabIndex        =   77
            Top             =   3000
            Width           =   465
         End
         Begin VB.TextBox cctOfficePhone1 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2070
            MaxLength       =   3
            TabIndex        =   76
            Top             =   3000
            Width           =   465
         End
         Begin VB.TextBox cctHomePhone2 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2700
            MaxLength       =   4
            TabIndex        =   75
            Top             =   3360
            Width           =   645
         End
         Begin VB.TextBox cctHomeArea 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1530
            MaxLength       =   3
            TabIndex        =   74
            Top             =   3360
            Width           =   465
         End
         Begin VB.TextBox cctHomePhone1 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2070
            MaxLength       =   3
            TabIndex        =   73
            Top             =   3360
            Width           =   465
         End
         Begin VB.TextBox cctNPI 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   70
            Top             =   5520
            Width           =   2145
         End
         Begin VB.FileListBox lstCME 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   9300
            TabIndex        =   39
            Top             =   6960
            Width           =   3765
         End
         Begin VB.TextBox txtDeaNumber 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   38
            Top             =   6105
            Width           =   2145
         End
         Begin VB.TextBox txtPractice 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5640
            TabIndex        =   37
            Top             =   120
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.ComboBox cboSuffix 
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
            Left            =   12060
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   580
            Width           =   1065
         End
         Begin VB.ComboBox cboSpecialty 
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
            Left            =   5580
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   5505
            Width           =   2145
         End
         Begin VB.ComboBox cboTitle 
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
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1020
            Width           =   2600
         End
         Begin VB.CommandButton cmdSpecTerms 
            Caption         =   "..."
            Height          =   285
            Left            =   7830
            TabIndex        =   33
            Top             =   5505
            Width           =   285
         End
         Begin VB.Frame fraSex 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "Gender"
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
            Left            =   1440
            TabIndex        =   30
            Top             =   2400
            Width           =   1965
            Begin VB.OptionButton CCOFemale 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Female"
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
               Left            =   930
               TabIndex        =   32
               Top             =   165
               Width           =   975
            End
            Begin VB.OptionButton CCOMale 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Male"
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
               Height          =   225
               Left            =   120
               TabIndex        =   31
               Top             =   165
               Width           =   795
            End
         End
         Begin VB.TextBox txtDiscount 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   17130
            TabIndex        =   29
            Top             =   3840
            Width           =   735
         End
         Begin VB.CheckBox chkExternal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "External Type:"
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
            Left            =   15120
            TabIndex        =   28
            Top             =   3570
            Width           =   1755
         End
         Begin VB.ComboBox cccState 
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
            Left            =   6630
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1875
            Width           =   2235
         End
         Begin VB.TextBox txtZipcode 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10320
            MaxLength       =   9
            TabIndex        =   26
            Top             =   1785
            Width           =   1635
         End
         Begin VB.TextBox txtCity 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            MaxLength       =   32
            TabIndex        =   25
            Top             =   1905
            Width           =   2325
         End
         Begin VB.TextBox txtAddress2 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6600
            MaxLength       =   48
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   1395
            Width           =   2235
         End
         Begin VB.TextBox txtAddress1 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            MaxLength       =   48
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   1460
            Width           =   3675
         End
         Begin VB.TextBox txtFirstName 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            MaxLength       =   24
            TabIndex        =   22
            Top             =   580
            Width           =   2205
         End
         Begin VB.TextBox txtMiddleName 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5700
            MaxLength       =   24
            TabIndex        =   21
            Top             =   580
            Width           =   1965
         End
         Begin VB.TextBox txtLastName 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9030
            MaxLength       =   24
            TabIndex        =   20
            Top             =   580
            Width           =   2205
         End
         Begin VB.TextBox txtUPIN 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5580
            MaxLength       =   10
            TabIndex        =   19
            Top             =   4980
            Width           =   2145
         End
         Begin VB.ListBox lstFacility 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   10380
            Style           =   1  'Checkbox
            TabIndex        =   18
            Top             =   6120
            Width           =   3495
         End
         Begin VB.TextBox txtProviderID 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            TabIndex        =   17
            Top             =   150
            Width           =   1155
         End
         Begin VB.TextBox txtSS1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11280
            MaxLength       =   3
            TabIndex        =   15
            Top             =   2430
            Width           =   420
         End
         Begin VB.TextBox txtSS2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11880
            MaxLength       =   2
            TabIndex        =   14
            Top             =   2430
            Width           =   330
         End
         Begin VB.TextBox txtSS3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   12390
            MaxLength       =   4
            TabIndex        =   13
            Top             =   2430
            Width           =   660
         End
         Begin VB.TextBox txtEIN1 
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
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   12
            Top             =   4980
            Width           =   420
         End
         Begin VB.TextBox txtEIN2 
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
            Left            =   2460
            MaxLength       =   7
            TabIndex        =   11
            Top             =   4980
            Width           =   1260
         End
         Begin PT_XP_Button.PT_XP cmdProvInsur 
            Height          =   525
            Left            =   12300
            TabIndex        =   8
            Top             =   8610
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   926
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
            BackColorDown   =   15133676
            TransparentColor=   14737632
            Caption         =   "Insurance Provider ID's"
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdCancel 
            Height          =   345
            Left            =   7253
            TabIndex        =   9
            Top             =   9420
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
            Picture         =   "frmProviders.frx":1666
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Cancel"
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmProviders.frx":1A00
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdSave 
            Height          =   345
            Left            =   6023
            TabIndex        =   10
            Top             =   9420
            Width           =   915
            _ExtentX        =   1614
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
            Picture         =   "frmProviders.frx":1D52
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Save"
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmProviders.frx":20EC
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdDocuments 
            Height          =   285
            Left            =   13140
            TabIndex        =   16
            Top             =   7350
            Width           =   585
            _ExtentX        =   1032
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
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Path"
            CaptionPosition =   0
            DepthMode       =   1
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmProviders.frx":243E
            PictureAlignment=   5
            ShowFocus       =   -1  'True
         End
         Begin MSMask.MaskEdBox txtBirthDate 
            Height          =   315
            Left            =   6660
            TabIndex        =   40
            Top             =   2430
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "MM/DD/YYYY"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ccdDeaExpire 
            Height          =   315
            Left            =   6180
            TabIndex        =   41
            Top             =   6090
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "MM/DD/YYYY"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin PT_XP_Button.PT_XP cmdTitleTerms 
            Height          =   285
            Left            =   4620
            TabIndex        =   42
            Top             =   1050
            Width           =   315
            _ExtentX        =   556
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
            MaskColor       =   128
            Picture         =   "frmProviders.frx":2DC0
            AppearanceThemes=   3
            BorderWidth     =   2
            TransparentColor=   14737632
            Caption         =   ""
            CaptionPosition =   4
            DepthMode       =   1
            DepthEvent      =   1
            ForeColorDisabled=   -2147483629
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmProviders.frx":335A
            PictureAlignment=   4
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdSuffixMaint 
            Height          =   285
            Left            =   13170
            TabIndex        =   43
            Top             =   600
            Width           =   315
            _ExtentX        =   556
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
            MaskColor       =   128
            Picture         =   "frmProviders.frx":36AC
            AppearanceThemes=   3
            BorderWidth     =   2
            TransparentColor=   14737632
            Caption         =   ""
            CaptionPosition =   4
            DepthMode       =   1
            DepthEvent      =   1
            ForeColorDisabled=   -2147483629
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmProviders.frx":3C46
            PictureAlignment=   4
            ShowFocus       =   -1  'True
         End
         Begin FPSpreadADO.fpSpread grdMedLicense 
            Height          =   1305
            Left            =   90
            TabIndex        =   114
            ToolTipText     =   "Right Click on Cell to Remove Item"
            Top             =   7800
            Width           =   11715
            _Version        =   458752
            _ExtentX        =   20664
            _ExtentY        =   2302
            _StockProps     =   64
            AutoCalc        =   0   'False
            AutoClipboard   =   0   'False
            ColHeaderDisplay=   1
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
            MaxCols         =   12
            MaxRows         =   1
            ProcessTab      =   -1  'True
            RetainSelBlock  =   0   'False
            ScrollBars      =   1
            ScrollBarShowMax=   0   'False
            SelectBlockOptions=   0
            ShadowColor     =   16761024
            ShadowDark      =   8388608
            ShadowText      =   8388608
            SpreadDesigner  =   "frmProviders.frx":3F98
            UnitType        =   0
            UserResize      =   0
            Appearance      =   2
            TextTip         =   4
            TextTipDelay    =   10
            ScrollBarTrack  =   2
            ClipboardOptions=   0
         End
         Begin PT_XP_Button.PT_XP cmdAddLicence 
            Height          =   345
            Left            =   120
            TabIndex        =   115
            Top             =   7440
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
            Picture         =   "frmProviders.frx":42B0
            AppearanceThemes=   3
            BorderDrawEvent =   1
            TransparentColor=   14737632
            Caption         =   "Add"
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmProviders.frx":484A
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblTaxID 
            BackStyle       =   0  'Transparent
            Caption         =   "TAXOMONY ID"
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
            Left            =   9210
            TabIndex        =   120
            Top             =   5580
            Width           =   1740
         End
         Begin VB.Label cctTaxType 
            BackStyle       =   0  'Transparent
            Caption         =   "TAXOMONY NAME"
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
            Left            =   9210
            TabIndex        =   118
            Top             =   5010
            Width           =   1740
         End
         Begin VB.Label lblUserID 
            BackStyle       =   0  'Transparent
            Caption         =   "PatientTrac User ID"
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
            Left            =   150
            TabIndex        =   113
            Top             =   3975
            Width           =   1200
         End
         Begin VB.Label lblPassword 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Left            =   4500
            TabIndex        =   112
            Top             =   4005
            Width           =   1200
         End
         Begin VB.Label lblReEnter 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
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
            Left            =   9300
            TabIndex        =   111
            Top             =   4020
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "("
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
            Index           =   7
            Left            =   10665
            TabIndex        =   107
            Top             =   3015
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ")"
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
            Index           =   8
            Left            =   11220
            TabIndex        =   106
            Top             =   3015
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Index           =   9
            Left            =   11805
            TabIndex        =   105
            Top             =   3060
            Width           =   90
         End
         Begin VB.Label lblCellPhone 
            BackStyle       =   0  'Transparent
            Caption         =   "Cell Phone"
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
            Left            =   9390
            TabIndex        =   104
            Top             =   3135
            Width           =   1260
         End
         Begin VB.Label lblPager 
            BackStyle       =   0  'Transparent
            Caption         =   "Pager"
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
            Left            =   5340
            TabIndex        =   99
            Top             =   3510
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "("
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
            Index           =   4
            Left            =   6615
            TabIndex        =   98
            Top             =   3015
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ")"
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
            Index           =   5
            Left            =   7170
            TabIndex        =   97
            Top             =   3015
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Index           =   6
            Left            =   7755
            TabIndex        =   96
            Top             =   3015
            Width           =   90
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Facsimile"
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
            Left            =   5340
            TabIndex        =   95
            Top             =   3090
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Index           =   10
            Left            =   7755
            TabIndex        =   94
            Top             =   3435
            Width           =   90
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ")"
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
            Index           =   11
            Left            =   7170
            TabIndex        =   93
            Top             =   3435
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "("
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
            Index           =   12
            Left            =   6615
            TabIndex        =   92
            Top             =   3435
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "("
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
            Index           =   0
            Left            =   1425
            TabIndex        =   86
            Top             =   3015
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ")"
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
            Index           =   2
            Left            =   1980
            TabIndex        =   85
            Top             =   3015
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Index           =   3
            Left            =   2565
            TabIndex        =   84
            Top             =   3015
            Width           =   90
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Office"
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
            Left            =   150
            TabIndex        =   83
            Top             =   3090
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "("
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
            Index           =   1
            Left            =   1425
            TabIndex        =   82
            Top             =   3435
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ")"
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
            Index           =   13
            Left            =   1980
            TabIndex        =   81
            Top             =   3435
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Index           =   14
            Left            =   2565
            TabIndex        =   80
            Top             =   3375
            Width           =   90
         End
         Begin VB.Label lblHome 
            BackStyle       =   0  'Transparent
            Caption         =   "Home"
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
            Left            =   150
            TabIndex        =   79
            Top             =   3450
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CME, LICENSES, CERTIFICATES, LEGAL"
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
            Left            =   9300
            TabIndex        =   72
            Top             =   6690
            Width           =   3690
         End
         Begin VB.Label lblNPI 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPI NUMBER"
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
            TabIndex        =   71
            Top             =   5610
            Width           =   1215
         End
         Begin VB.Line Line1 
            X1              =   30
            X2              =   14055
            Y1              =   4830
            Y2              =   4830
         End
         Begin VB.Label lblGender 
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
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
            Left            =   150
            TabIndex        =   69
            Top             =   2550
            Width           =   1200
         End
         Begin VB.Label lblDeaNumber 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DEA NUMBER"
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
            TabIndex        =   68
            Top             =   6165
            Width           =   1275
         End
         Begin VB.Label lblSuffix 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suffix"
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
            Left            =   11400
            TabIndex        =   67
            Top             =   675
            Width           =   570
         End
         Begin VB.Label lblSpecialty 
            BackStyle       =   0  'Transparent
            Caption         =   "Medical Specialty"
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
            Height          =   495
            Left            =   4500
            TabIndex        =   66
            Top             =   5490
            Width           =   990
         End
         Begin VB.Label lblExternalInst 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmProviders.frx":4B9C
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
            Height          =   825
            Left            =   14970
            TabIndex        =   65
            Top             =   4320
            Width           =   3315
         End
         Begin VB.Label lblDiscount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Applied Discount (%):"
            Enabled         =   0   'False
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
            Left            =   15000
            TabIndex        =   64
            Top             =   3945
            Width           =   2085
         End
         Begin VB.Label lblZip 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Zipcode"
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
            Left            =   9390
            TabIndex        =   63
            Top             =   1905
            Width           =   735
         End
         Begin VB.Label lblState 
            BackStyle       =   0  'Transparent
            Caption         =   "State"
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
            Left            =   5400
            TabIndex        =   62
            Top             =   1995
            Width           =   1200
         End
         Begin VB.Label lblCity 
            BackStyle       =   0  'Transparent
            Caption         =   "City"
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
            Left            =   150
            TabIndex        =   61
            Top             =   2025
            Width           =   1200
         End
         Begin VB.Label lblAdd2 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 2"
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
            Left            =   5370
            TabIndex        =   60
            Top             =   1500
            Width           =   1200
         End
         Begin VB.Label lblAdd1 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 1"
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
            Left            =   150
            TabIndex        =   59
            Top             =   1575
            Width           =   1200
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
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
            Left            =   150
            TabIndex        =   58
            Top             =   1125
            Width           =   1200
         End
         Begin VB.Label lblMSFacility 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Facilities:"
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
            Left            =   9270
            TabIndex        =   57
            Top             =   6180
            Width           =   975
         End
         Begin VB.Label lblPrID 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Provider ID :"
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
            Left            =   150
            TabIndex        =   56
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblFirstName 
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
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
            Left            =   150
            TabIndex        =   55
            Top             =   675
            Width           =   1200
         End
         Begin VB.Label lblMiddleName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name"
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
            Left            =   4320
            TabIndex        =   54
            Top             =   675
            Width           =   1275
         End
         Begin VB.Label lblLastName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
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
            Left            =   7860
            TabIndex        =   53
            Top             =   675
            Width           =   1035
         End
         Begin VB.Label lblUPIN 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UPIN"
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
            Left            =   4590
            TabIndex        =   52
            Top             =   5070
            Width           =   480
         End
         Begin VB.Label lblLicenseNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Medical Licenses"
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
            TabIndex        =   51
            Top             =   7200
            Width           =   1605
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Social Security No:"
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
            Left            =   9390
            TabIndex        =   50
            Top             =   2490
            Width           =   1800
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FEIN NUMBER"
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
            TabIndex        =   49
            Top             =   5040
            Width           =   1335
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth"
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
            Left            =   3720
            TabIndex        =   48
            Top             =   2550
            Width           =   2880
         End
         Begin VB.Label lblDash 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   6
            Left            =   11730
            TabIndex        =   47
            Top             =   2430
            Width           =   150
         End
         Begin VB.Label lblDash 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   7
            Left            =   12240
            TabIndex        =   46
            Top             =   2430
            Width           =   150
         End
         Begin VB.Label lblDash 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   2280
            TabIndex        =   45
            Top             =   5010
            Width           =   150
         End
         Begin VB.Label lblDeaExp 
            BackStyle       =   0  'Transparent
            Caption         =   "Expiration Date"
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
            Left            =   4470
            TabIndex        =   44
            Top             =   6150
            Visible         =   0   'False
            Width           =   1545
         End
      End
      Begin VB.Frame ccFmain 
         BackColor       =   &H00E0E0E0&
         Height          =   8655
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   13545
         Begin PT_XP_Button.PT_XP PT_XP1 
            Height          =   345
            Left            =   10080
            TabIndex        =   122
            Top             =   8070
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorDown   =   15133676
            TransparentColor=   14737632
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdExit 
            Height          =   525
            Left            =   12270
            TabIndex        =   3
            Top             =   7980
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
            Picture         =   "frmProviders.frx":4C27
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmProviders.frx":5501
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
            Left            =   12960
            Picture         =   "frmProviders.frx":6153
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   8040
            Visible         =   0   'False
            Width           =   465
         End
         Begin PT_XP_Button.PT_XP cmdActivate 
            Height          =   555
            Left            =   150
            TabIndex        =   4
            Top             =   7980
            Width           =   1455
            _ExtentX        =   2566
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
            Picture         =   "frmProviders.frx":6255
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Activate"
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmProviders.frx":6B2F
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdAdd 
            Height          =   525
            Left            =   4725
            TabIndex        =   5
            Top             =   7980
            Width           =   1545
            _ExtentX        =   2725
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
            Picture         =   "frmProviders.frx":7781
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Add"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmProviders.frx":805B
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdUpdate 
            Height          =   555
            Left            =   6570
            TabIndex        =   6
            Top             =   7980
            Width           =   1635
            _ExtentX        =   2884
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
            Picture         =   "frmProviders.frx":8CAD
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Update"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmProviders.frx":9587
            ShowFocus       =   -1  'True
         End
         Begin FPSpreadADO.fpSpread grdData 
            Height          =   7785
            Left            =   30
            TabIndex        =   116
            Top             =   120
            Width           =   13485
            _Version        =   458752
            _ExtentX        =   23786
            _ExtentY        =   13732
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
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            GridColor       =   8388608
            MaxCols         =   10
            MaxRows         =   1
            NoBorder        =   -1  'True
            ProcessTab      =   -1  'True
            RetainSelBlock  =   0   'False
            RowHeaderDisplay=   0
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   16761024
            ShadowDark      =   8388608
            ShadowText      =   8388608
            SpreadDesigner  =   "frmProviders.frx":A1D9
            UserResize      =   0
            Appearance      =   1
            TextTip         =   2
            TextTipDelay    =   60
            ClipboardOptions=   0
         End
      End
   End
End
Attribute VB_Name = "frmProviders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fbAdd As Boolean
Private mbooAddRow As Boolean
Private mlngGridRow As Long
Private mlngGridCol As Long
Private mlngMedRow As Long
Private mlngMedCol As Long
Private mFirst As Boolean
Private grdGrdLock As Boolean
Private fpDeltaLeft As Long
Private fpDeltaTop As Long
Private flngDeltaX  As Long
Private flngDeltaY  As Long
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal Hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long

Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long
Private udtMedLicense As udtMedLicense
Private Type udtMedLicense
    State As String
    LicNo As String
    ExpDate As String
    IssueDate As String
    Status As String
    Primary As String
    Term As String
End Type

Private Sub Add_Provider()
    ccfFFac.Enabled = True
    ClearTextValues
    sstTab.TabEnabled(0) = False
    sstTab.TabEnabled(1) = True
    sstTab.Tab = 1
    fbAdd = True
    Exit Sub
End Sub

Private Sub Clear_MedLicense()
    With udtMedLicense
        .State = vbNullString
        .LicNo = vbNullString
        .ExpDate = vbNullString
        .IssueDate = vbNullString
        .Status = vbNullString
        .Primary = vbNullString
        .Term = vbNullString
    End With
End Sub


Private Function Validate_License()
Dim lngRow As Long
Dim varCellValue
Dim booAdd As Boolean
Dim strState As String

    On Error GoTo Error_Validate_License
    
    strState = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultState")
    With grdMedLicense
        .ReDraw = False
        If .MaxRows >= 1 Then
            For lngRow = 1 To .MaxRows
                .GetText 3, lngRow, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    If Trim(varCellValue) = udtSpecProvider.Prim_License_No Then
                        booAdd = False
                        Exit For
                    End If
                End If
            Next lngRow
        Else
            booAdd = True
        End If
        If booAdd Then
            Add_Row
            lngRow = .MaxRows
            Set_GridCurSel grdMedLicense, Trim(strState), 1, lngRow, "2", "STATE", "Type to Enter Selection"
            grdMedLicense.SetCellBorder 1, lngRow, 2, lngRow, 16, GridBlue, CellBorderStyleDot
            Format_CellTypeEdit grdMedLicense, lngRow, 3, "2", udtSpecProvider.Prim_License_No, "TypeHAlignLeft", "LICENSE_NO", " ", DataColor
            .SetCellBorder 3, lngRow, 4, lngRow, 16, GridBlue, CellBorderStyleDot
            Format_CellTypeEdit grdMedLicense, lngRow, 7, "2", udtSpecProvider.Prim_License_Exp, "TypeHAlignLeft", "EXPIRATION_DATE", "", DataColor
            .SetCellBorder 7, lngRow, 8, lngRow, 16, GridBlue, CellBorderStyleDot
        End If
        .ReDraw = True
    End With
     Exit Function
    Screen.MousePointer = vbDefault
Error_Validate_License:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Validate_License", App.EXEName & "." & Me.Name
    Exit Function
    Resume

End Function

Private Function Write_Med_License()
Dim cnProvider As New ADODB.Connection
Dim rsProvider As New ADODB.Recordset
Dim strSQL As String
Dim lngProvID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim strState As String
Dim varCellValue


    On Error GoTo Error_Write_Med_License
    Screen.MousePointer = vbHourglass
    lngProvID = udtSpecProvider.ProviderID
    strState = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultState")
    
    If cnProvider.State = adStateClosed Then cnProvider.open psConnect(0)
    With grdMedLicense
        .ReDraw = False
        For lngRow = 1 To .MaxRows
            Clear_MedLicense
            lngPrimKey = 0
             lngPrimKey = .GetRowItemData(lngRow)
            .GetText 1, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then udtMedLicense.State = Trim(varCellValue)
            If Trim(strState) = udtMedLicense.State Then udtMedLicense.Primary = "True"
            .GetText 3, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then udtMedLicense.LicNo = Trim(varCellValue)
            .GetText 5, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then udtMedLicense.IssueDate = Trim(varCellValue)
            .GetText 7, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then udtMedLicense.ExpDate = Trim(varCellValue)
            .GetText 9, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then udtMedLicense.Term = Trim(varCellValue)
            .GetText 11, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then udtMedLicense.Status = Trim(varCellValue)
            If lngPrimKey > 0 Then
                strSQL = "SELECT * FROM PROVIDER_MEDICAL_LICENSE WHERE MED_LIC_ID = " & lngPrimKey
                Set rsProvider = cnProvider.Execute(strSQL)
                If Not rsProvider.EOF Then
                    lngPrimKey = rsProvider.Fields("MED_LIC_ID")
                Else
                    lngPrimKey = 0
                End If
            End If
            If lngPrimKey = 0 Then
                strSQL = "INSERT INTO  PROVIDER_MEDICAL_LICENSE (PROVIDER_ID , STATE, LICENSE_NO, ISSUE_DATE, EXPIRATION_DATE, TERM, STATUS, PRIMARY) " _
                        & "values (" & udtSpecProvider.ProviderID & ",'" & udtMedLicense.State & "', '" & udtMedLicense.LicNo & "'," _
                        & udtMedLicense.IssueDate & "','" & udtMedLicense.ExpDate & "','" & udtMedLicense.ExpDate & "', '" & udtMedLicense.Term & "'," _
                        & udtMedLicense.Status & "','" & udtMedLicense.Primary & "')"
                Set rsProvider = cnProvider.Execute(strSQL)
            Else

                strSQL = "UPDATE PROVIDER_MEDICAL_LICENSE SET STATE = '" & udtMedLicense.State & "'," _
                        & "LICENSE_NO = '" & udtMedLicense.LicNo & "', ISSUE_DATE = '" & udtMedLicense.IssueDate & "'," _
                        & "EXPIRATION_DATE = '" & udtMedLicense.ExpDate & "', TERM = '" & udtMedLicense.Term & "'," _
                        & "STATUS = '" & udtMedLicense.Status & "', PRIMARY = '" & udtMedLicense.Primary & "' WHERE " _
                        & "MED_LIC_ID = " & lngPrimKey
                Set rsProvider = cnProvider.Execute(strSQL)
            End If
        Next lngRow
        .ReDraw = True
    End With
     Exit Function
    Screen.MousePointer = vbDefault
Error_Write_Med_License:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Med_License", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Read_Med_License()
Dim cnProvider As New ADODB.Connection
Dim rsProvider As New ADODB.Recordset
Dim strSQL As String
Dim lngProvID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim strState As String


    On Error GoTo Error_Read_Med_License
    Screen.MousePointer = vbHourglass
    lngProvID = udtSpecProvider.ProviderID
    If cnProvider.State = adStateClosed Then cnProvider.open psConnect(0)
    grdMedLicense.MaxRows = lngRow
    grdMedLicense.ReDraw = True
    strState = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultState")
    strSQL = "SELECT * FROM PROVIDER_MEDICAL_LICENSE WHERE PROVIDER_ID = " & udtSpecProvider.ProviderID
    Set rsProvider = cnProvider.Execute(strSQL)
    With rsProvider
        Do Until .EOF
            lngRow = lngRow + 1
            grdMedLicense.MaxRows = lngRow
            If (!MED_LIC_ID) <> vbNullString Then
                grdMedLicense.SetRowItemData (lngRow), Val(!MED_LIC_ID)
            End If
            If Trim(!State) <> vbNullString Then
                Format_CellTypeEdit grdMedLicense, lngRow, 1, "2", Trim(!State), "TypeHAlignLeft", "STATE", " ", DataColor
               '  Set_GridCurSel grdMedLicense, Trim(!State), 1, lngRow, "2", "STATE", "Type to Enter Selection"
                grdMedLicense.SetCellBorder 1, lngRow, 2, lngRow, 16, GridBlue, CellBorderStyleDot
            End If
            If Trim(!License_No) <> vbNullString Then
                 Format_CellTypeEdit grdMedLicense, lngRow, 3, "2", Trim(!License_No), "TypeHAlignLeft", "LICENSE_NO", " ", DataColor
                grdMedLicense.SetCellBorder 3, lngRow, 4, lngRow, 16, GridBlue, CellBorderStyleDot
            End If
            If Trim(!ISSUED_DATE) <> vbNullString Then
                Format_CellTypeEdit grdMedLicense, lngRow, 5, "1", Trim(!ISSUED_DATE), "TypeHAlignLeft", "ISSUED_DATE", " ", DataColor
                grdMedLicense.SetCellBorder 5, lngRow, 6, lngRow, 16, GridBlue, CellBorderStyleDot
            End If
            If Trim(!EXPIRATION_DATE) <> vbNullString Then
                Format_CellTypeEdit grdMedLicense, lngRow, 7, "2", Trim(!EXPIRATION_DATE), "TypeHAlignLeft", "EXPIRATION_DATE", "", DataColor
                grdMedLicense.SetCellBorder 7, lngRow, 8, lngRow, 16, GridBlue, CellBorderStyleDot
            End If
            If Trim(!Term) <> vbNullString Then
                Set_GridCurSel grdMedLicense, Trim((!Term)), 9, lngRow, "2", "TERM", "Type to Enter Selection"
                grdMedLicense.SetCellBorder 9, lngRow, 10, lngRow, 16, GridBlue, CellBorderStyleDot
            End If
            If Trim(!Status) <> vbNullString Then
                Format_CellTypeEdit grdMedLicense, lngRow, 11, "12", Trim(!Status), "TypeHAlignLeft", "STATUS", "", DataColor
                grdMedLicense.SetCellBorder 11, lngRow, 12, lngRow, 16, GridBlue, CellBorderStyleDot
            End If
            .MoveNext
        Loop
        grdMedLicense.ReDraw = True
    End With
     Exit Function
    Screen.MousePointer = vbDefault
Error_Read_Med_License:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Med_License", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Sub cctCellArea_KeyPress(KeyAscii As Integer)
    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub cctCellArea_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(cctCellArea.Text) = cctCellArea.MaxLength Then
         If cctCellPhone1.Enabled Then cctCellPhone1.SetFocus
      End If
End Sub


Private Sub cctCellPhone1_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub cctCellPhone1_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(cctCellPhone1.Text) = cctCellPhone1.MaxLength Then
         If cctCellPhone2.Enabled Then cctCellPhone2.SetFocus
      End If
End Sub


Private Sub cctCellPhone2_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub


Private Sub cctFaxArea_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub cctFaxArea_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(cctFaxArea.Text) = cctFaxArea.MaxLength Then
         If cctFaxPhone1.Enabled Then cctFaxPhone1.SetFocus
      End If
End Sub


Private Sub cctFaxPhone1_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub cctFaxPhone1_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(cctFaxPhone1.Text) = cctFaxPhone1.MaxLength Then
         If cctFaxPhone2.Enabled Then cctFaxPhone2.SetFocus
      End If
End Sub


Private Sub cctFaxPhone2_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub


Private Sub cctHomeArea_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub cctHomeArea_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(cctHomeArea.Text) = cctHomeArea.MaxLength Then
         If cctHomePhone1.Enabled Then cctHomePhone1.SetFocus
      End If
End Sub


Private Sub cctHomePhone1_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub cctHomePhone1_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(cctHomePhone1.Text) = cctHomePhone1.MaxLength Then
         If cctHomePhone2.Enabled Then cctHomePhone2.SetFocus
      End If
End Sub


Private Sub cctHomePhone2_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub cctOfficeArea_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub cctOfficeArea_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(cctOfficeArea.Text) = cctOfficeArea.MaxLength Then
         If cctOfficePhone1.Enabled Then cctOfficePhone1.SetFocus
      End If
End Sub


Private Sub cctOfficePhone1_KeyPress(KeyAscii As Integer)
    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub cctOfficePhone1_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(cctOfficePhone1.Text) = cctOfficePhone1.MaxLength Then
         If cctOfficePhone2.Enabled Then cctOfficePhone2.SetFocus
      End If
End Sub


Private Sub cctOfficePhone2_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub


Private Sub cctPageArea_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub cctPageArea_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(cctPageArea.Text) = cctPageArea.MaxLength Then
         If cctPagePhone1.Enabled Then cctPagePhone1.SetFocus
      End If
End Sub


Private Sub cctPagePhone1_KeyPress(KeyAscii As Integer)
   If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub cctPagePhone1_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(cctPagePhone1.Text) = cctPagePhone1.MaxLength Then
         If cctPagePhone2.Enabled Then cctPagePhone2.SetFocus
      End If
End Sub


Private Sub cctPagePhone2_KeyPress(KeyAscii As Integer)

    If CheckNumberOnly(KeyAscii) = False Then KeyAscii = 0
End Sub


Private Sub cctTaxName_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    If lstTaxonomy.listcount <= 0 Then Get_TaxonomyData
    strList = Trim(cctTaxName.Text)
    If strList <> vbNullString And Len(strList) > 2 Then
            lstTaxonomy.ListIndex = SendMessage(lstTaxonomy.Hwnd, LB_FINDSTRING, -1, _
            ByVal strList)
    End If
    lstTaxonomy.Visible = True
    Exit Sub
End Sub


Private Sub chkExternal_Click()
    On Error GoTo ErrHandler
    
    If chkExternal.Value Then
        lblDiscount.Enabled = True
        txtDiscount.Enabled = True
        txtDiscount.BackColor = vbWhite
    Else
        lblDiscount.Enabled = False
        txtDiscount.Enabled = False
        txtDiscount.Text = vbNullString
        txtDiscount.BackColor = &H80000000
    End If
    
    Exit Sub
    
ErrHandler:
    ErrHandler _
     "frmProviders.chkExternal_Click", _
     Err.Number, _
     Err.Description
End Sub

Private Sub cmdAccept_Click()
Dim strDate As String
Dim lngDateDiff As Long
Dim varCellValue
Dim strTag As String

  strDate = Format(ccdDate.Value, "MMMM DD, YYYY")
  fraNumber.Visible = False
  If Not IsNull(ccdDate.Value) Then
    With grdMedLicense
        .Col = mlngMedCol
        .Row = mlngMedRow
        If .CellTag <> vbNullString Then
            Format_CellTypeEdit grdMedLicense, mlngMedRow, mlngMedCol, "2", "", "TypeHAlignLeft", strTag, " ", DataColor
            .SetCellBorder mlngMedCol, mlngMedRow, mlngMedCol + 1, mlngMedRow, 16, GridBlue, CellBorderStyleDot
        End If
    End With
  End If
  grdGrdLock = False
  Exit Sub
End Sub

Private Sub cmdActivate_Click()
Dim strSQL As String
Dim cnProvider As New ADODB.Connection
Dim rsProvider As New ADODB.Recordset
Dim sValue
Dim i As Integer
Dim lngProvID As Long

    On Error GoTo cmdActivate_Error

    If mlngGridRow < 1 Then
        MsgBox "Please select appropriate row first!", vbInformation
        Exit Sub
    End If
    
    If cmdActivate.Caption = "Activate" Then
        Format_CellTypeStatic grdData, mlngGridRow, 8, "1", "Yes", "TypeHAlignLeft"
        grdData.SetCellBorder 8, mlngGridRow, 8, mlngGridRow, 16, GridBlue, CellBorderStyleDot
       cmdActivate.Caption = "DeActivate"
    Else
        Format_CellTypeStatic grdData, mlngGridRow, 8, "1", "No", "TypeHAlignLeft"
        grdData.SetCellBorder 8, mlngGridRow, 8, mlngGridRow, 16, GridBlue, CellBorderStyleDot
       cmdActivate.Caption = "Activate"
    End If

    lngProvID = grdData.GetRowItemData(mlngGridRow)
 
    If cnProvider.State = adStateClosed Then cnProvider.open psConnect(0)
    If lngProvID > 0 Then
        If cmdActivate.Caption = "DeActivate" Then
            strSQL = "UPDATE PROVIDERS SET ACTIVE = TRUE WHERE PROVIDER_ID = " & lngProvID
            Set rsProvider = cnProvider.Execute(strSQL)
        ElseIf cmdActivate.Caption = "DeActivate" Then
            strSQL = "UPDATE PROVIDERS SET ACTIVE = FALSE WHERE PROVIDER_ID = " & lngProvID
            Set rsProvider = cnProvider.Execute(strSQL)
        End If
    End If
    If cnProvider.State = adStateOpen Then
        cnProvider.Close
        Set cnProvider = Nothing
        Set rsProvider = Nothing
   End If
    Exit Sub
cmdActivate_Error:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "cmdActivate", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdAdd_Click()
Dim strSQL As String
Dim booret As Boolean
Dim varRS
Dim lngrecs As Long
Dim cnAuth As New ADODB.Connection
Dim rsAuth As New ADODB.Recordset
Dim booSkip As Boolean


    On Error GoTo cmdAdd_Error
    If cnAuth.State = adStateClosed Then cnAuth.open psConnect(0)
    strSQL = "SELECT ACCT_PROVIDER FROM FACILITIES"
    Set rsAuth = cnAuth.Execute(strSQL)
    With rsAuth
        If Not IsNull(!ACCT_PROVIDER) Then
            If Val((!ACCT_PROVIDER)) = 0 Then
                MsgBox "No Authorized Licenses for PatientTrac. Contact PatientTrac Customer Support.", vbInformation
                Unload Me
            Else
                If Val(!ACCT_PROVIDER) <= grdData.MaxRows Then
                    MsgBox "Please contact your PatientTrac Representative for Additional Client Licenses!", vbInformation
                    Exit Sub
                Else
                    Add_Provider
                End If
            End If
        End If
    End With
    If cnAuth.State = adStateOpen Then
        cnAuth.Close
        Set cnAuth = Nothing
        Set rsAuth = Nothing
    End If
    Exit Sub
cmdAdd_Error:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "cmdAdd_Error", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub cmdAddLicence_Click()
    Add_Row
End Sub

Private Sub cmdCancel_Click()
        
    On Error GoTo cmdCancel_Error
    
    fbAdd = False

    sstTab.Tab = 0
    sstTab.TabEnabled(1) = False
    sstTab.TabEnabled(0) = True
    ccfFFac.Enabled = False
    
    Exit Sub

cmdCancel_Error:
    ErrHandler _
     "frmProviders.cmdCancel_Click", _
     Err.Number, _
     Err.Description
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdProvInsur_Click()
  frmProvInsurance.Show vbModal
  Exit Sub
End Sub

Private Sub cmdSave_Click()
    If Not Validate_Entries Then Exit Sub
   
    If SaveData = True Then
       sstTab.Tab = 0
       ccfFFac.Enabled = False
       Write_Med_License
    End If
End Sub

Private Sub cmdSpecTerms_Click()
    Dim sStr As String
    
    On Error GoTo Error_cmdSpecTerms_Click
    
    sStr = cboSpecialty.Text
    
    frmNurseTemplates.fCaption = "Provider Specialties"
    frmNurseTemplates.fType = "Provider Specialties"
    frmNurseTemplates.fSubDomain = "sd_Provider_Specialty"
    frmNurseTemplates.fConnectCR = psConnect(1)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    GetList cboSpecialty, "sd_Provider_Specialty"
    If Trim(sStr) <> vbNullString Then
       cboSpecialty.ListIndex = SetListIndex(cboSpecialty, sStr)
       If cboSpecialty.ListIndex = -1 Then
          cboSpecialty.AddItem sStr
          cboSpecialty.ListIndex = cboSpecialty.NewIndex
       End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdSpecTerms_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdSpecTerms_Click", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdSuffixMaint_Click()
    Dim sStr As String
    
    On Error GoTo ErrHandler
    
    sStr = cboSuffix.Text
    
    frmNurseTemplates.fCaption = "Name Suffix"
    frmNurseTemplates.fType = "Name Suffix"
    frmNurseTemplates.fSubDomain = "sd_Suffix"
    frmNurseTemplates.fConnectCR = psConnect(1)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    
    LoadAttrValidValue psConnect(1), cboSuffix, "sd_Suffix"
    
    If Trim(sStr) <> vbNullString Then
       cboSuffix.ListIndex = SetListIndex(cboSuffix, sStr)
       If cboSuffix.ListIndex = -1 Then
          cboSuffix.AddItem sStr
          cboSuffix.ListIndex = cboSuffix.NewIndex
       End If
    End If
    
    Exit Sub

ErrHandler:
    ErrHandler _
     "frmProviders.cmdSuffixMaint_Click", _
     Err.Number, _
     Err.Description
End Sub

Private Sub cmdTitleTerms_Click()
    Dim sStr As String
    
    On Error GoTo ErrHandler
    
    sStr = cboTitle.Text
    
    frmNurseTemplates.fCaption = "Provider Titles"
    frmNurseTemplates.fType = "Provider Titles"
    frmNurseTemplates.fSubDomain = "sd_Provider_Title"
    frmNurseTemplates.fConnectCR = psConnect(1)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    
    LoadAttrValidValue psConnect(1), cboTitle, "sd_Provider_Title"
    
    If Trim(sStr) <> vbNullString Then
       cboTitle.ListIndex = SetListIndex(cboTitle, sStr)
       If cboTitle.ListIndex = -1 Then
          cboTitle.AddItem sStr
          cboTitle.ListIndex = cboTitle.NewIndex
       End If
    End If
    
    
    Exit Sub

ErrHandler:
    ErrHandler _
     "frmProviders.cmdTitleTerms_Click", _
     Err.Number, _
     Err.Description
End Sub
Private Sub cmdUpdate_Click()

    On Error GoTo cmdUpdate_Error
    
    If grdData.MaxRows = 0 Then Exit Sub
    
    ccfFFac.Enabled = True
    ClearTextValues
    LoadTextValues
    sstTab.TabEnabled(0) = False
    sstTab.TabEnabled(1) = True
    sstTab.Tab = 1
    fbAdd = False
    
    Exit Sub
    
cmdUpdate_Error:
    ErrHandler _
     "frmProviders.cmdUpdate_Click", _
     Err.Number, _
     Err.Description
End Sub

Private Sub Add_Row()
Dim lngCol As Long
Dim lngGridRow As Long
Dim strList As String

        With grdMedLicense
            .ReDraw = True
            lngGridRow = lngGridRow + 1
            .MaxRows = lngGridRow
            Format_CellTypeCombo grdMedLicense, lngGridRow, 1, "2", Load_States, "TypeHAlignLeft", "STATE", DataColor, "Type to Enter Selection"
            .SetCellBorder 1, lngGridRow, 2, lngGridRow, 16, GridBlue, CellBorderStyleDot
            Format_CellTypeEdit grdMedLicense, lngGridRow, 3, "2", "", "TypeHAlignLeft", "LICENSE_NO", " ", DataColor
            .SetCellBorder 3, lngGridRow, 4, lngGridRow, 16, GridBlue, CellBorderStyleDot
            Format_CellTypeEdit grdMedLicense, lngGridRow, 5, "2", " ", "TypeHAlignLeft", "ISSUED_DATE", " ", DataColor
            .SetCellBorder 5, lngGridRow, 6, lngGridRow, 16, GridBlue, CellBorderStyleDot
            Format_CellTypeEdit grdMedLicense, lngGridRow, 7, "2", " ", "TypeHAlignLeft", "EXPIRATION_DATE", "", DataColor
            .SetCellBorder 7, lngGridRow, 8, lngGridRow, 16, GridBlue, CellBorderStyleDot
            strList = "One Year" & Chr(9) & "Two Years" & Chr(9) & "Three Years" & Chr(9) & "Four Years" & Chr(9) & "Five Years" & Chr(9) & "Six Years"
            Format_CellTypeCombo grdMedLicense, lngGridRow, 9, "2", strList, "TypeHAlignLeft", "TERM", DataColor, "Type to Enter Selection"
            .SetCellBorder 9, lngGridRow, 10, lngGridRow, 16, GridBlue, CellBorderStyleDot
            Format_CellTypeEdit grdMedLicense, 11, "2", " ", "TypeHAlignLeft", "STATUS", "", DataColor
            .SetCellBorder 11, lngGridRow, 12, lngGridRow, 16, GridBlue, CellBorderStyleDot
            .ReDraw = True
        End With
    Exit Sub
End Sub

Private Sub Format_Med_License()
Dim lngCol As Long
Dim lngGridRow As Long

    Screen.MousePointer = vbHourglass
    With grdMedLicense
        .ReDraw = False
        .ClearRange 0, SpreadHeader, grdMedLicense.MaxRows, SpreadHeader, False
        .ProcessTab = True
        .FontBold = True
        .Font = "Times New Roman"
        .FontSize = 8
        .ColHeaderRows = 2
        .ColHeadersShow = True
        .RowHeadersShow = False
        .MaxCols = 12
        .ColWidth(1) = 15
        .ColWidth(2) = 1
        .ColWidth(3) = 14
        .ColWidth(4) = 2
        .ColWidth(5) = 14
        .ColWidth(6) = 1
        .ColWidth(7) = 15
        .ColWidth(8) = 1
        .ColWidth(9) = 15
        .ColWidth(10) = 1
        .ColWidth(11) = 15
        .ColWidth(12) = 1
        
        Format_CellTypeStatic grdMedLicense, SpreadHeader, 1, "2", "STATE", "TypeHAlignCenter"
        Format_CellTypeStatic grdMedLicense, SpreadHeader, 3, "2", "LICENSE NUMBER", "TypeHAlignCenter"
        Format_CellTypeStatic grdMedLicense, SpreadHeader, 5, "2", "ISSUED DATE", "TypeHAlignCenter"
        Format_CellTypeStatic grdMedLicense, SpreadHeader, 7, "2", "EXPIRATION DATE", "TypeHAlignCenter"
        Format_CellTypeStatic grdMedLicense, SpreadHeader, 9, "2", "TERM", "TypeHAlignCenter"
        Format_CellTypeStatic grdMedLicense, SpreadHeader, 11, "2", "STATUS", "TypeHAlignCenter"
        .ReDraw = True
    End With
    Add_Spacer SpreadHeader + 1, grdMedLicense
    
    If udtSpecProvider.Prim_License_No <> vbNullString Then Read_Med_License
    Screen.MousePointer = vbDefault
End Sub
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
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
            If vstrCellNote <> vbNullString Then
                .CellNote = vstrCellNote
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            End If
            If lngBackColor > 0 Then .BackColor = lngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        .BlockMode = False
    End With
End Function

Private Sub Form_Activate()
Dim strSQL As String
Dim booret As Boolean
Dim varRS
Dim i%
Dim strTitle
Dim lngrecs As Long

    On Error GoTo cmdUpdate_Error
    If mFirst Then mFirst = False
    Screen.MousePointer = vbHourglass
    grdData.ZOrder 0
    sstTab.Tab = 0
    sstTab.TabEnabled(1) = False
    mbooAddRow = False
    Format_Grid
    LoadGrid
    Load_States
    Load_Facilities
    Format_Med_License
    
    GetTerms cboTitle, "sd_Provider_Title"
    GetTerms cboSuffix, "sd_Suffix"
    GetTerms cboSpecialty, "sd_Provider_Specialty"
    ccfFFac.Enabled = False
    centerForm Me
    Me.Top = Me.Top + 300
    Screen.MousePointer = vbDefault
    Exit Sub
    
cmdUpdate_Error:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_TaxonomyData", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Form_Load()
    mFirst = True
End Sub
Private Function LoadGrid() As Boolean
Dim sValue
Dim lngRow As Long
Dim strSQL As String
Dim cnProvider As New ADODB.Connection
Dim rsProvider As New ADODB.Recordset

    On Error GoTo LoadGrid_Error
    If cnProvider.State = adStateClosed Then cnProvider.open psConnect(0)
    strSQL = "Load_Active_Providers"
    Set rsProvider = cnProvider.Execute(strSQL)
    With rsProvider
       Do Until .EOF
            If (!Provider_ID) <> vbNullString Then
                lngRow = grdData.MaxRows + 1
                grdData.MaxRows = lngRow
                If (!Provider_ID) <> vbNullString Then
                    Format_CellTypeStatic grdData, lngRow, 1, "1", (!Provider_ID), "TypeHAlignLeft"
                    grdData.SetCellBorder 1, lngRow, 1, lngRow, 16, GridBlue, CellBorderStyleDot
                    grdData.SetRowItemData (lngRow), (!Provider_ID)
                End If
                
                If (!First_Name) <> vbNullString Then
                    Format_CellTypeStatic grdData, lngRow, 2, "1", (!First_Name), "TypeHAlignLeft"
                    grdData.SetCellBorder 2, lngRow, 2, lngRow, 16, GridBlue, CellBorderStyleDot
                End If
                If (!Middle_Name) <> vbNullString Then
                    Format_CellTypeStatic grdData, lngRow, 3, "1", (!Middle_Name), "TypeHAlignLeft"
                    grdData.SetCellBorder 3, lngRow, 3, lngRow, 16, GridBlue, CellBorderStyleDot
                End If
                If (!Last_Name) <> vbNullString Then
                    Format_CellTypeStatic grdData, lngRow, 4, "1", (!Last_Name), "TypeHAlignLeft"
                    grdData.SetCellBorder 4, lngRow, 4, lngRow, 16, GridBlue, CellBorderStyleDot
                End If
                If (!Suffix) <> vbNullString Then
                    Format_CellTypeStatic grdData, lngRow, 5, "1", (!Suffix), "TypeHAlignLeft"
                    grdData.SetCellBorder 5, lngRow, 5, lngRow, 16, GridBlue, CellBorderStyleDot
                End If
                If (!Title) <> vbNullString Then
                    Format_CellTypeStatic grdData, lngRow, 6, "1", (!Title), "TypeHAlignLeft"
                    grdData.SetCellBorder 6, lngRow, 6, lngRow, 16, GridBlue, CellBorderStyleDot
                End If
                If (!Specialty_1) <> vbNullString Then
                    Format_CellTypeStatic grdData, lngRow, 7, "1", (!Specialty_1), "TypeHAlignLeft"
                    grdData.SetCellBorder 7, lngRow, 7, lngRow, 16, GridBlue, CellBorderStyleDot
                End If
                If Trim(!Active) <> vbNullString Then
                    If (!Active) = True Then
                        Format_CellTypeStatic grdData, lngRow, 8, "1", "Yes", "TypeHAlignLeft"
                        grdData.SetCellBorder 8, lngRow, 8, lngRow, 16, GridBlue, CellBorderStyleDot
                    Else
                        Format_CellTypeStatic grdData, lngRow, 8, "1", "No", "TypeHAlignLeft"
                        grdData.SetCellBorder 8, lngRow, 8, lngRow, 16, GridBlue, CellBorderStyleDot
                        grdData.BlockMode = True
                            grdData.Col = 1
                            grdData.Col2 = grdData.MaxCols
                            grdData.Row = lngRow
                            grdData.Row2 = lngRow
                            grdData.BackColor = grdGray
                        grdData.BlockMode = False
                    End If
                End If
            End If
          .MoveNext
       Loop
    End With
    If cnProvider.State = adStateOpen Then
        cnProvider.Close
        Set cnProvider = Nothing
        Set rsProvider = Nothing
    End If
    Exit Function
LoadGrid_Error:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LoadGrid", App.EXEName & "." & Me.Name
    Exit Function
    Resume
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
              .FontName = "Times New Roman"
              .TypeHAlign = TypeHAlignCenter
              .FontSize = 8
              .FontBold = True
          .BlockMode = False
    End With
End Function
Private Sub Form_Unload(Cancel As Integer)
 
  On Error Resume Next

    Unload Me
    Exit Sub
End Sub

Private Sub grdData_Click(ByVal Col As Long, ByVal Row As Long)
Dim iRow As Long



    If Row <= 0 Then Exit Sub
    mlngGridRow = Row
    mlngGridCol = Col
   
    
    udtSpecProvider.ProviderID = grdData.GetRowItemData(mlngGridRow)
    If udtSpecProvider.ProviderID > 0 Then cmdUpdate_Click
    With grdData
         For iRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col = .MaxCols
                .Row = iRow
                .Row2 = iRow
                .BackColor = grdWhite
            .BlockMode = False
         Next iRow
        If Row >= 0 Then
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                .BackColor = SelectColor
             .BlockMode = False
        End If
    End With
 
    
End Sub

Private Sub grdData_GotFocus()
    Dim sValue
    
    On Error GoTo ErrHandler
    
    grdData.GetText 6, grdData.ActiveRow, sValue
    grdData.Row = grdData.ActiveRow
    
    If sValue = "Yes" Then cmdActivate.Caption = "DeActivate" Else cmdActivate.Caption = "Activate"
    If cmdActivate.Visible = False Then cmdActivate.Visible = True

    Exit Sub

ErrHandler:
    ErrHandler _
     "frmProviders.grdData_GotFocus", _
     Err.Number, _
     Err.Description
End Sub
Private Sub LoadTextValues()
Dim sPhone1 As String
Dim strPhone As String
Dim vProvID As Long
Dim Idx As Integer
Dim strSocialSec As String
Dim strEIN As String
Dim Arr
Dim fso As New FileSystemObject
Dim strFolder As String
Dim bNewDialog As Boolean
Dim bNoNewFolder As Boolean
Dim lngScanFile As Long
Dim strLegalPath As String
Dim strScan As String

    On Error GoTo Error_LoadTextValues
    vProvID = grdData.GetRowItemData(mlngGridRow)
    If vProvID > 0 Then
            Set_SpecificProv vProvID
            txtProviderID.Text = udtSpecProvider.ProviderID
            txtFirstName.Text = Replace(udtSpecProvider.FirstName, "''", "'")
            txtMiddleName.Text = Replace(udtSpecProvider.MiddleName, "''", "'")
            txtLastName.Text = Replace(udtSpecProvider.LastName, "''", "'")
            
            txtAddress1.Text = udtSpecProvider.Address
            txtAddress2.Text = udtSpecProvider.Address2
            txtCity.Text = udtSpecProvider.City
            For Idx = 0 To cccState.listcount - 1
                If cccState.List(Idx) = udtSpecProvider.State Then
                   cccState.ListIndex = Idx
                   Exit For
                End If
            Next Idx
            For Idx = 0 To cboTitle.listcount - 1
                If cboTitle.List(Idx) = udtSpecProvider.Title Then
                   cboTitle.ListIndex = Idx
                   Exit For
                End If
            Next Idx
            For Idx = 0 To cboSpecialty.listcount - 1
                If cboSpecialty.List(Idx) = udtSpecProvider.Specialty_1 Then
                   cboSpecialty.ListIndex = Idx
                   Exit For
                End If
            Next Idx
                 
            For Idx = 0 To cboSuffix.listcount - 1
                If cboSuffix.List(Idx) = udtSpecProvider.Suffix Then
                   cboSuffix.ListIndex = Idx
                   Exit For
                End If
            Next Idx
            If udtSpecProvider.NPI_No <> vbNullString Then cctNPI.Text = udtSpecProvider.NPI_No
            
            If InStr(udtSpecProvider.Telephone, "-") Then
                Arr = Split(udtSpecProvider.Telephone, "-")
                If UBound(Arr) > 0 Then
                     cctOfficeArea.Text = Arr(0)
                    cctOfficePhone1.Text = Arr(1)
                    cctOfficePhone2.Text = Arr(2)
                End If
            End If
                 
            If InStr(udtSpecProvider.CellPhone, "-") Then
                Arr = Split(udtSpecProvider.CellPhone, "-")
                If UBound(Arr) > 1 Then
                    cctCellArea.Text = Arr(0)
                    cctCellPhone1.Text = Arr(1)
                    cctCellPhone2.Text = Arr(2)
                End If
            End If
                 
            If InStr(udtSpecProvider.Home_Phone, "-") Then
                Arr = Split(udtSpecProvider.Home_Phone, "-")
                If UBound(Arr) > 1 Then
                    cctHomeArea.Text = Arr(0)
                    cctHomePhone1.Text = Arr(1)
                    cctHomePhone2.Text = Arr(2)
                End If
            End If
                 
            If InStr(udtSpecProvider.Facsimile, "-") Then
                Arr = Split(udtSpecProvider.Facsimile, "-")
                If UBound(Arr) > 1 Then
                    cctFaxArea.Text = Arr(0)
                    cctFaxPhone1.Text = Arr(1)
                    cctFaxPhone2.Text = Arr(2)
                End If
            End If
                 
            If InStr(udtSpecProvider.Pager, "-") Then
                Arr = Split(udtSpecProvider.Pager, "-")
                If UBound(Arr) > 1 Then
                    cctPageArea.Text = Arr(0)
                    cctPagePhone1.Text = Arr(1)
                    cctPagePhone2.Text = Arr(2)
                End If
            End If
        
           If udtSpecProvider.Birth_Date <> vbNullString Then txtBirthDate.Text = Format(udtSpecProvider.Birth_Date, "mm/dd/yyyy")
            If Trim(udtSpecProvider.DEA_Exp) <> vbNullString Then ccdDeaExpire.Text = Format(udtSpecProvider.DEA_Exp, "mm/dd/yyyy")
            If Trim(udtSpecProvider.DEA_Number) <> vbNullString Then txtDeaNumber.Text = Trim(udtSpecProvider.DEA_Number)
           
           
            If udtSpecProvider.Gender = "M" Then CCOMale.Value = True
            If udtSpecProvider.Gender = "F" Then CCOFemale.Value = True
                   
            If Trim(udtSpecProvider.Social_Security_No) <> vbNullString Then
                If InStr(udtSpecProvider.Social_Security_No, "-") > 0 Then
                   Arr = Split(strSocialSec, "-")
                   txtSS1.Text = Arr(0)
                   txtSS2.Text = Arr(1)
                   txtSS3.Text = Arr(2)
                End If
             End If
    
            If udtSpecProvider.FEIN <> vbNullString Then
                 If InStr(udtSpecProvider.FEIN, "-") > 0 Then
                    Arr = Split(udtSpecProvider.FEIN, "-")
                    txtEIN1.Text = Arr(0)
                    txtEIN2.Text = Arr(1)
                 End If
            End If
            txtUPIN.Text = udtSpecProvider.UPIN
            txtPractice.Text = udtSpecProvider.Practice
            txtUserId.Text = udtSpecProvider.UserID
            If txtUserId.Text = vbNullString Then txtUserId.Text = Mid$(Replace(txtLastName.Text, " ", vbNullString), 1, 8) & txtProviderID.Text
            txtPassword1.Text = Trim(udtSpecProvider.Password)
            txtPassword2.Visible = False
            lblReEnter.Visible = False
       Get_Facilities txtProviderID.Text
    End If
    strFolder = udtSpecProvider.ProviderID
    strLegalPath = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ProviderLegal")
    strScan = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScannedFiles")
    If Trim(strLegalPath) = vbNullString Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "ProviderLegal", " ", REG_SZ
    End If
 
    If fso.FolderExists(strLegalPath) Then
        If Right(strLegalPath, 1) <> "\" Then strLegalPath = strLegalPath & "\"
        lstCME.Path = strLegalPath
    End If
    If fso.FolderExists(lstCME.Path & "\" & strFolder) Then
        If TextFound(":\", lstCME.Path, Len(":\"), True) Then
             lstCME.Path = lstCME.Path & strFolder
        Else
             lstCME.Path = lstCME.Path & "\" & strFolder
        End If
    Else
        fso.CreateFolder lstCME.Path & "\" & strFolder
        lstCME.Path = lstCME.Path & "\" & strFolder
    End If
    Format_Med_License
    If udtSpecProvider.Prim_License_No <> vbNullString Then
        Write_Med_License
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_LoadTextValues:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "LoadTextValues_", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub ClearTextValues()
Dim i As Integer
Dim lobjControl As Object
Dim lstrControlType As String
Dim frm As Form
Dim obj As Object
Dim Idx%
Dim Control As Control

    On Error GoTo ClearTextValues_Error

    For Idx = 0 To lstFacility.listcount - 1
        lstFacility.Selected(Idx) = False
    Next Idx
    lstFacility.ListIndex = -1
    cboTitle.ListIndex = -1
    cccState.ListIndex = -1
    cboSpecialty.ListIndex = -1
    cboSuffix.ListIndex = -1
    txtProviderID.Text = vbNullString
    txtFirstName.Text = vbNullString
    txtMiddleName.Text = vbNullString
    txtLastName.Text = vbNullString
    txtAddress1.Text = vbNullString
    txtAddress2.Text = vbNullString
    txtCity.Text = vbNullString
    txtZipcode.Text = vbNullString
    cctOfficeArea.Text = vbNullString
    cctOfficePhone1.Text = vbNullString
    cctOfficePhone2.Text = vbNullString
    cctFaxArea.Text = vbNullString
    cctFaxPhone1.Text = vbNullString
    cctFaxPhone2.Text = vbNullString
    cctCellArea.Text = vbNullString
    cctCellPhone1.Text = vbNullString
    cctCellPhone2.Text = vbNullString
    txtSS1.Text = vbNullString
    txtSS2.Text = vbNullString
    txtSS3.Text = vbNullString
    cctHomeArea.Text = vbNullString
    cctHomePhone1.Text = vbNullString
    cctHomePhone2.Text = vbNullString
    cctPageArea.Text = vbNullString
    cctPagePhone1.Text = vbNullString
    cctPagePhone2.Text = vbNullString
    txtUserId.Text = vbNullString
    txtPassword1.Text = vbNullString
    txtPassword2.Text = vbNullString
    txtEIN1.Text = vbNullString
    txtEIN2.Text = vbNullString
    txtUPIN.Text = vbNullString
    cctNPI.Text = vbNullString
    cctTaxName.Text = vbNullString
    cctTaxonomyID.Text = vbNullString
    txtDeaNumber.Text = vbNullString
    txtBirthDate.Text = "__/__/____"
    ccdDeaExpire.Text = "__/__/____"
    CCOMale.Value = False
    CCOFemale.Value = False
    chkExternal.Value = 0
    txtPassword2.Visible = False
    lblReEnter.Visible = False
    Exit Sub
ClearTextValues_Error:
     Exit Sub
     Resume
End Sub

Private Function SaveData() As Boolean
Dim cnProvider As New ADODB.Connection
Dim rsProvider As New ADODB.Recordset
Dim strSQL As String
Dim lngProvID As Long
Dim booinsert As Boolean
Dim lngActive As Boolean

    On Error GoTo SaveData_Error
    Screen.MousePointer = vbHourglass
    If Trim(cctTaxName.Text) <> vbNullString Then Get_TaxonomyID Trim(cctTaxName.Text)
    lngProvID = udtSpecProvider.ProviderID
    Clear_SpecProvider
    If cnProvider.State = adStateClosed Then cnProvider.open psConnect(0)
    If CCOMale.Value Then udtSpecProvider.Gender = "Male"
    If CCOFemale.Value Then udtSpecProvider.Gender = "Female"
       
    udtSpecProvider.Birth_Date = Format(txtBirthDate.Text, "mm/dd/yyyy 00:00:00 AM/PM")
    If ccdDeaExpire.Text <> "__/__/____" Then udtSpecProvider.DEA_Exp = Format(ccdDeaExpire.Text, "mm/dd/yyyy 00:00:00 AM/PM") Else udtSpecProvider.DEA_Exp = vbNullString
    
    udtSpecProvider.Social_Security_No = txtSS1 & "-" & txtSS2 & "-" & txtSS3
    udtSpecProvider.FEIN = txtEIN1.Text & "-" & txtEIN2.Text
    udtSpecProvider.CellPhone = cctCellArea.Text & "-" & cctCellPhone1.Text & "-" & cctCellPhone2.Text
    udtSpecProvider.Home_Phone = cctHomeArea.Text & "-" & cctHomePhone1.Text & "-" & cctHomePhone2.Text
    udtSpecProvider.Facsimile = cctFaxArea.Text & "-" & cctFaxPhone1.Text & "-" & cctFaxPhone2.Text
    udtSpecProvider.Telephone = cctOfficeArea.Text & "-" & cctOfficePhone1.Text & "-" & cctOfficePhone2.Text

    udtSpecProvider.FirstName = Replace(Trim(txtFirstName.Text), "'", "''")
    udtSpecProvider.MiddleName = Replace(Trim(txtMiddleName.Text), "'", "''")
    udtSpecProvider.LastName = Replace(Trim(txtLastName.Text), "'", "''")
    udtSpecProvider.City = Replace(Trim(txtCity.Text), "'", "''")
    udtSpecProvider.State = Trim(cccState.Text)
    udtSpecProvider.ZipCode = Replace(Trim(txtZipcode.Text), "'", "''")
    udtSpecProvider.Address = Replace(Trim(txtAddress1.Text), "'", "''")
    udtSpecProvider.Address2 = Replace(Trim(txtAddress2.Text), "'", "''")
    udtSpecProvider.Specialty_1 = Trim(cboSpecialty.Text)
    udtSpecProvider.TaxonomyString = Trim(cctTaxName.Text)
    udtSpecProvider.TaxonomyID = Trim(cctTaxonomyID.Text)
    udtSpecProvider.NPI_No = cctNPI.Text
    udtSpecProvider.UPIN = txtUPIN.Text
    udtSpecProvider.UserID = txtUserId.Text
    udtSpecProvider.Password = txtPassword1.Text
    
    strSQL = "SELECT PROVIDER_ID FROM PROVIDER WHERE  FIRST_NAME = '" & udtSpecProvider.FirstName & "' AND LAST_NAME = '" & udtSpecProvider.LastName & "'"
    Set rsProvider = cnProvider.Execute(strSQL)
    If Not rsProvider.EOF Then
        If rsProvider.Fields("PROVIDER_ID") = lngProvID Then
            booinsert = False
        Else
             booinsert = True
        End If
    End If
    If udtSpecProvider.Active = "Yes" Then
        lngActive = True
    Else
        lngActive = False
    End If
    If booinsert Then
            strSQL = "INSERT INTO PROVIDERS (PRACTICE ,FIRST_NAME ,MIDDLE_NAME ,LAST_NAME ,TITLE ,SUFFIX ,ADDRESS_1 ,ADDRESS_2 ,CITY " _
                    & ",STATE ,ZIP_CODE ,SPECIALTY_1 , SOCIAL_SECURITY_NO ,EMPLOYER_ID_NO ,BIRTH_DATE ,TELEPHONE_1 " _
                    & ",FACSIMILE,CELL_PHONE,HOME_PHONE, PAGER, UPIN, LICENSE_NO, DEA_NUMBER, SEX, USER_ID " _
                    & ",USER_PASSWORD, MED_LICENSE_EXPIRE, DEA_LICENSE_EXPIRE, NPI_ID, TAXONOMY_NO, TAXONOMY_NAME " _
                    & ",ACTIVE)  Values " _
                    & "('" & udtSpecProvider.Practice & "','" & udtSpecProvider.FirstName & "','" & udtSpecProvider.MiddleName & "','" _
                    & udtSpecProvider.LastName & "','" & udtSpecProvider.Title & "','" & udtSpecProvider.Suffix & "','" _
                    & udtSpecProvider.Address & "','" & udtSpecProvider.Address2 & "','" & udtSpecProvider.City & "','" & udtSpecProvider.State & "'," _
                    & udtSpecProvider.ZipCode & "','" & udtSpecProvider.Specialty_1 & "','" & udtSpecProvider.Social_Security_No & "','" _
                    & udtSpecProvider.FEIN & "','" & udtSpecProvider.Birth_Date & "','" & udtSpecProvider.Telephone & "','" _
                    & udtSpecProvider.Facsimile & "','" & udtSpecProvider.CellPhone & "','" & udtSpecProvider.Home_Phone & "','" _
                    & udtSpecProvider.Pager & "','" & udtSpecProvider.UPIN & "','" & udtSpecProvider.Prim_License_No & "','" _
                    & udtSpecProvider.DEA_Number & "','" & udtSpecProvider.Gender & "','" & udtSpecProvider.UserID & "','" _
                    & udtSpecProvider.Password & "','" & udtSpecProvider.Prim_License_Exp & "','" & udtSpecProvider.DEA_Exp & "','" _
                    & udtSpecProvider.NPI_No & "','" & udtSpecProvider.TaxonomyID & "','" & udtSpecProvider.TaxonomyString & "'," _
                    & lngActive & ")"
            Set rsProvider = cnProvider.Execute(strSQL)
    Else
            strSQL = "UPDATE PROVIDERS SET PRACTICE = '" & udtSpecProvider.Practice & "'," _
                & "FIRST_NAME = '" & udtSpecProvider.FirstName & "', MIDDLE_NAME = '" & udtSpecProvider.MiddleName & "'," _
                & "LAST_NAME = '" & udtSpecProvider.LastName & "', TITLE = '" & udtSpecProvider.Title & "'," _
                & "SUFFIX = '" & udtSpecProvider.Suffix & "', ADDRESS_1 = '" & udtSpecProvider.Address & "'," _
                & "ADDRESS_2 = '" & udtSpecProvider.Address2 & "', CITY  = '" & udtSpecProvider.City & "'," _
                & "STATE = '" & udtSpecProvider.State & "', ZIP_CODE  = '" & udtSpecProvider.ZipCode & "'," _
                & "SPECIALTY_1 = '" & udtSpecProvider.Specialty_1 & "', SOCIAL_SECURITY_NO   = '" & udtSpecProvider.Social_Security_No & "'," _
                & "EMPLOYER_ID_NO = '" & udtSpecProvider.FEIN & "', BIRTH_DATE = '" & udtSpecProvider.Birth_Date & "'," _
                & "TELEPHONE_1 = '" & udtSpecProvider.Telephone & "', FACSIMILE = '" & udtSpecProvider.Facsimile & "'," _
                & "CELL_PHONE = '" & udtSpecProvider.CellPhone & "', HOME_PHONE = '" & udtSpecProvider.Home_Phone & "'," _
                & "PAGER = '" & udtSpecProvider.Pager & "', UPIN = '" & udtSpecProvider.UPIN & "', LICENSE_NO ='" & udtSpecProvider.Prim_License_No & "'," _
                & "DEA_NUMBER = '" & udtSpecProvider.DEA_Number & "', SEX = '" & udtSpecProvider.Gender & "'," _
                & "USER_ID = '" & udtSpecProvider.UserID & "',USER_PASSWORD = '" & udtSpecProvider.Password & "'," _
                & "MED_LICENSE_EXPIRE = '" & udtSpecProvider.Prim_License_Exp & "', DEA_LICENSE_EXPIRE = '" & udtSpecProvider.DEA_Exp & "'," _
                & "NPI_ID = '" & udtSpecProvider.NPI_No & "', TAXONOMY_NO = '" & udtSpecProvider.TaxonomyID & "'," _
                & "TAXONOMY_NAME = '" & udtSpecProvider.TaxonomyString & "', ACTIVE = " & lngActive & " WHERE PROVIDER_ID = " & lngProvID
            Set rsProvider = cnProvider.Execute(strSQL)
    End If
    Screen.MousePointer = vbDefault
    fbAdd = False
    SaveData = True
    Exit Function
SaveData_Error:
    Screen.MousePointer = vbDefault
    SaveData = False
    Write_Error Err.Description, Err.Number, strSQL, "SaveData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Sub grdData_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sValue
    
    On Error GoTo ErrHandler
    
    grdData.GetText 6, grdData.ActiveRow, sValue
    
    If sValue = "Yes" Then cmdActivate.Caption = "DeActivate" Else cmdActivate.Caption = "Activate"
    If cmdActivate.Visible = False Then cmdActivate.Visible = True
    
    Exit Sub
    
ErrHandler:
 
End Sub
Private Sub grdData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sValue
    
    On Error GoTo ErrHandler
    
    grdData.GetText 6, grdData.ActiveRow, sValue
    
    If sValue = "Yes" Then cmdActivate.Caption = "DeActivate" Else cmdActivate.Caption = "Activate"
    If cmdActivate.Visible = False Then cmdActivate.Visible = True
   
    Exit Sub
    
ErrHandler:
    ErrHandler _
     "frmProviders.grdData_MouseUp", _
     Err.Number, _
     Err.Description
End Sub



Private Function Get_TaxonomyData(Optional vstrData As String)
Dim cnTaxon As New ADODB.Connection
Dim rsTaxon As New ADODB.Recordset
Dim strSQL As String

    On Error GoTo Error_Get_TaxonomyData
    Screen.MousePointer = vbHourglass
    lstTaxonomy.Clear
    If cnTaxon.State = adStateClosed Then cnTaxon.open psConnect(1)
    LockWindowUpdate Hwnd
    strSQL = "SELECT DISTINCT SPECIALTY_NAME, TAXONOMY_NUMBER  FROM TAXONOMY ORDER BY SPECIALTY_NAME"
    Set rsTaxon = cnTaxon.Execute(strSQL)
    With rsTaxon
        Do Until .EOF
            If (!SPECIALTY_NAME) <> vbNullString Then lstTaxonomy.AddItem (!SPECIALTY_NAME)
            .MoveNext
        Loop
    End With
    If cnTaxon.State = adStateClosed Then
        cnTaxon.Close
        Set cnTaxon = Nothing
        Set rsTaxon = Nothing
    End If
    LockWindowUpdate 0
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_TaxonomyData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_TaxonomyData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Get_TaxonomyID(Optional vstrData As String)
Dim cnTaxon As New ADODB.Connection
Dim rsTaxon As New ADODB.Recordset
Dim strSQL As String

    On Error GoTo Error_Get_TaxonomyID
    Screen.MousePointer = vbHourglass
    If cnTaxon.State = adStateClosed Then cnTaxon.open psConnect(1)
    strSQL = "SELECT TAXONOMY_NUMBER  FROM TAXONOMY WHERE SPECIALTY_NAME = '" & vstrData & "'"
    Set rsTaxon = cnTaxon.Execute(strSQL)
    With rsTaxon
        Do Until .EOF
            If (!TAXONOMY_NUMBER) <> vbNullString Then cctTaxonomyID.Text = (!TAXONOMY_NUMBER)
            .MoveNext
        Loop
    End With
    If cnTaxon.State = adStateClosed Then
        cnTaxon.Close
        Set cnTaxon = Nothing
        Set rsTaxon = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_TaxonomyID:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_TaxonomyID", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub grdData_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)


    If Row <= 0 Then Exit Sub
    With grdData
        .BlockMode = True
            .Col = 1
            .Col = 2
            .Row = Row
            .Row2 = Row
            .BackColor = grdWhite
        .BlockMode = False
    End With
End Sub

Private Sub grdMedLicense_Click(ByVal Col As Long, ByVal Row As Long)
Dim iRow As Long
Dim lngx As Long
Dim lngY As Long
Dim lngWidth As Long
Dim lngHeight As Long
Dim booEmbed As Boolean
Dim CellNote As Long
Dim varMeas
Dim strLst As String
Dim varCellValue
Dim lngCol As Long

    If Row <= 0 Then Exit Sub
    If grdGrdLock Then Exit Sub
    mlngMedRow = Row
    mlngMedCol = Col
    With grdMedLicense
             For iRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col = 2
                .Row = iRow
                .Row2 = iRow
                .BackColor = grdWhite
            .BlockMode = False
         Next iRow
        If Row >= 0 Then
            .BlockMode = True
                .Col = Col
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                .BackColor = SelectColor
             .BlockMode = False
        End If
            .Col = Col
            .Row = Row
            If .CellTag <> vbNullString Then
                grdGrdLock = True
                If TextFound("_DATE", .CellTag, Len("_DATE"), True) Then
                    .GetCellPos Col, Row, lngx, lngY, lngWidth, lngHeight
                    grdGrdLock = True
                    fraNumber.ZOrder 0
                    lngY = lngY + 51.5
                    lngx = lngx + 50
                    fraNumber.Top = .Top + lngY
                    fraNumber.Left = .Left + lngx
                    fraNumber.Visible = True
                End If
            End If
     End With
End Sub


Private Sub grdMedLicense_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row <= 0 Then Exit Sub
    If grdGrdLock Then Exit Sub
    With grdMedLicense
        .BlockMode = True
            .Col = 1
            .Col = 2
            .Row = Row
            .Row2 = Row
            .BackColor = grdWhite
        .BlockMode = False
    End With
End Sub


Private Sub lstTaxonomy_Click()
    
    cctTaxName.Text = lstTaxonomy.Text
    lstTaxonomy.Visible = False
    Exit Sub
End Sub

Private Sub PT_XP1_Click()
    Format_Grid
    LoadGrid
End Sub

Private Sub txtDeaNumber_Click()
    ccdDeaExpire.Visible = True
   lblDeaExp.Visible = True
End Sub

Private Sub txtEIN1_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(txtEIN1.Text) = (txtEIN1.MaxLength) Then
         If txtEIN2.Enabled Then txtEIN2.SetFocus
      End If
End Sub


Private Sub txtPassword1_Change()
    lblReEnter.Visible = True
    txtPassword2.Visible = True
End Sub

Private Sub Load_Facilities()
Dim strSQL As String

    On Error GoTo Error_Load_Facilities
    
    strSQL = "Select facility_name, facility_id From FACILITIES Order by facility_name"
    LoadItemsKey lstFacility, strSQL, psConnect(0)
    
    Exit Sub
    
Error_Load_Facilities:
    ErrHandler _
     "frmProviders.Load_Facilities", _
     Err.Number, _
     Err.Description
     Exit Sub
     Resume
End Sub
Private Function Load_States() As String
Dim strSQL As String
Dim i%

    On Error GoTo Error_Load_States
    
    Open_DB psConnect(1)
    Load_States = vbNullString
    strSQL = "SELECT STATE_NAME  FROM STATES WHERE STATE_IND = 'y' ORDER BY STATE_NAME"
    Set rs = cn.Execute(strSQL)
    With rs
        Do Until .EOF
            If Trim(!state_name) <> vbNullString Then
                cccState.AddItem Trim(!state_name)
                If Load_States = vbNullString Then
                    Load_States = Trim(!state_name)
                Else
                    Load_States = Load_States & Chr(9) & Trim(!state_name)
                End If
            End If
        .MoveNext
        Loop
    End With
    Close_DB psConnect(1)
    Exit Function
    
Error_Load_States:
 
End Function

Private Sub Get_Facilities(vintProviderID As Integer)
Dim R As Integer
Dim i As Integer

    On Error GoTo Get_Facilities
    
    Open_DB psConnect(0)
    
    cn.S_FACILITIES_PROVIDERS vintProviderID, rs
    
    With rs
       Do While Not .EOF
          For i = 0 To lstFacility.listcount - 1
              If lstFacility.ItemData(i) = (!Facility_ID) Then
                 lstFacility.Selected(i) = True
                 Exit For
              End If
           Next i
           .MoveNext
       Loop
    End With
    
    Close_DB psConnect(0)
    
    Exit Sub
    
Get_Facilities:
    ErrHandler _
     "frmProviders.Get_Facilities", _
     Err.Number, _
     Err.Description
     Exit Sub
     Resume
End Sub
Private Sub Format_Grid()
Dim Idx As Integer
Dim strSQL As String
    
    On Error GoTo Err_Format_Grid
    udtSpecProvider.ProviderID = 0
    Screen.MousePointer = vbHourglass
    With grdData
        .ColHeaderRows = 2
        .ColsFrozen = 1
        .ReDraw = False
        .ClearRange 0, SpreadHeader, grdData.MaxRows, SpreadHeader, False
        .MaxRows = 0
        .MaxCols = 8
        .CursorStyle = CursorStyleArrow
        .ColWidth(1) = 7
        .ColWidth(2) = 15
        .ColWidth(3) = 15
        .ColWidth(4) = 15
        .ColWidth(5) = 15
        .ColWidth(6) = 15
        .ColWidth(7) = 15
        .ColWidth(8) = 15
        Format_CellTypeStatic grdData, SpreadHeader, 1, "1", "PROVIDER ID", "TypeHAlignLeft"
        
        Format_CellTypeStatic grdData, SpreadHeader, 2, "1", "FIRST NAME", "TypeHAlignLeft"
        Format_CellTypeStatic grdData, SpreadHeader, 3, "1", "MIDDLE NAME", "TypeHAlignLeft"
        Format_CellTypeStatic grdData, SpreadHeader, 4, "1", "LAST NAME", "TypeHAlignLeft"
        Format_CellTypeStatic grdData, SpreadHeader, 5, "1", "SUFFIX", "TypeHAlignLeft"
        Format_CellTypeStatic grdData, SpreadHeader, 6, "1", "TITLE", "TypeHAlignLeft"
        Format_CellTypeStatic grdData, SpreadHeader, 7, "1", "SPECIALTY", "TypeHAlignLeft"
        Format_CellTypeStatic grdData, SpreadHeader, 8, "1", "ACTIVE", "TypeHAlignLeft"
    End With
    Add_Spacer SpreadHeader + 1, grdData
    Screen.MousePointer = vbDefault
    Exit Sub
Err_Format_Grid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Grid", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Function Set_GridCurSel(vObjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, _
Optional vstrCellTag As String, Optional vstrCellNote As String, Optional vlngBackColor As Long)
Dim Idx As Integer
Dim booData As Boolean
Dim strSpan As String
Dim strTag As String
Dim strCellNote As String
Dim lngCol As Long
Dim lngSpan As Long

    With vObjGrid
        .Col = vlngCol
        .Row = vlngRow
        If .CellType = CellTypeComboBox Then
            For Idx = 0 To .TypeComboBoxCount
            .TypeComboBoxIndex = Idx
            If TextFound(Trim(vstrData), Trim(.TypeComboBoxString), Len(Trim(vstrData)), True) Then
                .TypeComboBoxCurSel = Idx
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                    .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                    lngSpan = Val(vstrSpan)
                    .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                End If
                If vstrCellNote <> vbNullString Then
                    .CellNote = vstrCellNote
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                End If
                If vlngBackColor > 0 Then .BackColor = vlngBackColor
                If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
                .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
                Exit For
            End If
            Next Idx
        End If
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
            .BackColor = 12632256
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 1, vlngRow, .MaxCols, vlngRow, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub
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
            .Font = "Times New Roman"
            .FontSize = 8
            If vbooUnderLine Then .FontUnderline = True
            If vlngBackColor Then .BackColor = vlngBackColor
            If Trim(vstrCellNote) <> vbNullString Then .CellNote = vstrCellNote
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
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
            .EditMode = False
          .BlockMode = False
    End With
End Function

Private Function Validate_Entries() As Boolean
    Dim Idx As Integer
    Dim bFnd As Boolean
    Dim sMonth As String
    Dim sDay As String
    Dim sYear As String
    
    On Error GoTo ErrHandler
    
    Validate_Entries = True
    
    If Trim(txtFirstName) = vbNullString Then
       MsgBox "First Name is Required!", vbInformation + vbOKCancel
       Validate_Entries = False
       Exit Function
    End If
    
    If Trim(txtLastName) = vbNullString Then
       MsgBox "Last Name is Required!", vbInformation + vbOKCancel
       Validate_Entries = False
       Exit Function
    End If
    
    If Trim(txtPassword1.Text) = vbNullString Then
       MsgBox "Password is Required!", vbInformation + vbOKCancel
       Validate_Entries = False
       Exit Function
    End If
    
    If txtPassword2.Visible = True And Trim(txtPassword2.Text) <> Trim(txtPassword1.Text) Then
       MsgBox "ReEntered Password not same as entered password!", vbInformation + vbOKCancel
       Validate_Entries = False
       Exit Function
    End If
    
    If Not chkExternal.Value Then
       bFnd = False
       For Idx = 0 To lstFacility.listcount - 1
           If lstFacility.Selected(Idx) = True Then
              bFnd = True
              Exit For
           End If
       Next Idx
       
       If Not bFnd Then
          Validate_Entries = False
          MsgBox "Select at least one facility.", vbExclamation, "Save Provider"
          Exit Function
       End If
    End If
    
    sMonth = Mid$(txtBirthDate.Text, 1, 2)
    sDay = Mid$(txtBirthDate.Text, 4, 2)
    sYear = Mid$(txtBirthDate.Text, 7, 4)

    If txtBirthDate.Text <> "__/__/____" Then
       If Not IsNumeric(sMonth) Then
          MsgBox "Invalid Month for Birth Date!", vbExclamation
          Validate_Entries = False
          Exit Function
       End If
       If CInt(sMonth) < 1 Or CInt(sMonth) > 12 Then
          MsgBox "Invalid Month for Birth Date!", vbExclamation
          Validate_Entries = False
          Exit Function
       End If

       If Not IsNumeric(sYear) Or Len(sYear) <> 4 Then
          MsgBox "Invalid Year for Birth Date!", vbExclamation
          Validate_Entries = False
          Exit Function
       End If
       If CInt(sYear) > Year(Now) Then
          MsgBox "Invalid Year for Birth Date!", vbExclamation
          Validate_Entries = False
          Exit Function
       End If

       If Not IsNumeric(sDay) Then
          MsgBox "Invalid Day for Birth Date!", vbExclamation
          Validate_Entries = False
          Exit Function
       End If

       If Not IsDayCheck(CInt(sMonth), CInt(sDay), CInt(sYear)) Then
          MsgBox "Invalid Day for Birth Date!", vbExclamation
          Validate_Entries = False
          Exit Function
       End If
    Else
       MsgBox "Birth Date is required!", vbExclamation
       Validate_Entries = False
       Exit Function
    End If
    
    Exit Function

ErrHandler:
    ErrHandler _
     "frmProviders.Validate_Entries", _
     Err.Number, _
     Err.Description
     Exit Function
     Resume
End Function

Private Sub txtSS1_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(txtSS1.Text) = (txtSS1.MaxLength) Then
         If txtSS2.Enabled Then txtSS2.SetFocus
      End If
End Sub


Private Sub txtSS2_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(txtSS2.Text) = (txtSS2.MaxLength) Then
         If txtSS3.Enabled Then txtSS3.SetFocus
      End If
End Sub


Private Sub txtSS3_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(txtSS3.Text) = (txtSS3.MaxLength) Then
         If txtEIN1.Enabled Then txtEIN1.SetFocus
      End If
End Sub


