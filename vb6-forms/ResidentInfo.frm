VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form ResidentInfo 
   BackColor       =   &H00C00000&
   Caption         =   "RESIDENT INFORMATION"
   ClientHeight    =   6945
   ClientLeft      =   6630
   ClientTop       =   2775
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ResidentInfo.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "ResidentInfo.frx":08CA
   ScaleHeight     =   6945
   ScaleWidth      =   12375
   Begin VB.Frame fraPatient 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   3660
         TabIndex        =   113
         Top             =   2220
         Width           =   915
      End
      Begin VB.TextBox cctHouseUnit 
         Height          =   315
         Left            =   1770
         TabIndex        =   111
         Top             =   2220
         Width           =   915
      End
      Begin VB.ComboBox cboLanguage 
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "ResidentInfo.frx":0DBE
         Left            =   4275
         List            =   "ResidentInfo.frx":0DC0
         Style           =   2  'Dropdown List
         TabIndex        =   107
         ToolTipText     =   "Select Patient Primary Language"
         Top             =   4890
         Width           =   1995
      End
      Begin VB.Frame fraPartName 
         BackColor       =   &H00E0E0E0&
         Height          =   2925
         Left            =   12900
         TabIndex        =   1
         Top             =   3240
         Visible         =   0   'False
         Width           =   5595
         Begin VB.TextBox txtParFirstName 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            TabIndex        =   10
            Top             =   990
            Width           =   2355
         End
         Begin VB.TextBox txtParMidName 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            TabIndex        =   9
            Top             =   1350
            Width           =   2355
         End
         Begin VB.TextBox txtParLastName 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            TabIndex        =   8
            Top             =   1710
            Width           =   2355
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Caption         =   "ADDRESS"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   60
            TabIndex        =   5
            Top             =   150
            Width           =   5475
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00800000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               Picture         =   "ResidentInfo.frx":0DC2
               ScaleHeight     =   255
               ScaleWidth      =   375
               TabIndex        =   6
               Top             =   60
               Width           =   375
            End
            Begin VB.Label lblPartTitle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PARTNER FULL NAME"
               BeginProperty Font 
                  Name            =   "Georgia"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   570
               TabIndex        =   7
               Top             =   90
               Width           =   2205
            End
         End
         Begin VB.ComboBox cboParTitle 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   600
            Width           =   2355
         End
         Begin VB.ComboBox cboPartRel 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2490
            Width           =   2985
         End
         Begin VB.ComboBox cboSuffix 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   2070
            Width           =   1485
         End
         Begin PT_XP_Button.PT_XP cmdParOk 
            Height          =   375
            Left            =   3990
            TabIndex        =   11
            Top             =   690
            Width           =   1035
            _ExtentX        =   1826
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
            Picture         =   "ResidentInfo.frx":134C
            AppearanceThemes=   3
            TransparentColor=   12582912
            Caption         =   "OK"
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "ResidentInfo.frx":16E6
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdPartCancel 
            Height          =   375
            Left            =   3990
            TabIndex        =   12
            Top             =   1200
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
            Picture         =   "ResidentInfo.frx":1A38
            AppearanceThemes=   3
            TransparentColor=   12582912
            Caption         =   "Cancel"
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "ResidentInfo.frx":244A
            ShowFocus       =   -1  'True
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title:"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   150
            TabIndex        =   18
            Top             =   690
            Width           =   465
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First:"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   150
            TabIndex        =   17
            Top             =   1050
            Width           =   465
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle:"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   150
            TabIndex        =   16
            Top             =   1380
            Width           =   630
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last:"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   150
            TabIndex        =   15
            Top             =   1800
            Width           =   420
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Relationship"
            ForeColor       =   &H00800000&
            Height          =   510
            Left            =   150
            TabIndex        =   14
            Top             =   2490
            Width           =   1080
         End
         Begin VB.Label lblSuffix 
            BackColor       =   &H00BFE8FD&
            BackStyle       =   0  'Transparent
            Caption         =   "Suffix:"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   150
            TabIndex        =   13
            Top             =   2160
            Width           =   705
         End
      End
      Begin VB.Frame fraPartner 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   1935
         Left            =   7080
         TabIndex        =   52
         Top             =   3900
         Visible         =   0   'False
         Width           =   4515
         Begin VB.TextBox txtParner 
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   99
            Top             =   1470
            Width           =   3075
         End
         Begin VB.ComboBox cboRelation 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1410
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   660
            Width           =   3075
         End
         Begin VB.TextBox txtEmergPh2 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   3105
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   97
            Top             =   1050
            Width           =   795
         End
         Begin VB.TextBox txtEmergPh1 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   2280
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   96
            Top             =   1050
            Width           =   615
         End
         Begin VB.TextBox txtEmergAC 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   95
            Top             =   1050
            Width           =   615
         End
         Begin VB.TextBox txtEmergContact 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1410
            TabIndex        =   94
            Top             =   300
            Width           =   3075
         End
         Begin VB.Label lblPartnerName 
            BackColor       =   &H00BFE8FD&
            BackStyle       =   0  'Transparent
            Caption         =   "Partner Name"
            ForeColor       =   &H00800000&
            Height          =   435
            Left            =   30
            TabIndex        =   106
            Top             =   1530
            Width           =   1275
         End
         Begin VB.Label lblRel 
            BackStyle       =   0  'Transparent
            Caption         =   "Relationship"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   60
            TabIndex        =   105
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label lblEmergPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone:"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   60
            TabIndex        =   104
            Top             =   1080
            Width           =   585
         End
         Begin VB.Label lblLeftParen 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "("
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
            Index           =   2
            Left            =   1410
            TabIndex        =   103
            Top             =   1095
            Width           =   75
         End
         Begin VB.Label lblRightPaen 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ")"
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
            Index           =   2
            Left            =   2175
            TabIndex        =   102
            Top             =   1095
            Width           =   75
         End
         Begin VB.Label lblDash 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Index           =   1
            Left            =   2970
            TabIndex        =   101
            Top             =   1080
            Width           =   105
         End
         Begin VB.Label lblEmContactName 
            BackColor       =   &H00BFE8FD&
            BackStyle       =   0  'Transparent
            Caption         =   "Emergency Contact"
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   60
            TabIndex        =   100
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Frame fraPic 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   9750
         TabIndex        =   50
         Top             =   270
         Width           =   2205
         Begin VB.Image imgPatient 
            Height          =   1635
            Left            =   90
            Stretch         =   -1  'True
            Top             =   180
            Width           =   1995
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
            Left            =   270
            TabIndex        =   51
            Top             =   600
            Visible         =   0   'False
            Width           =   1665
         End
      End
      Begin VB.TextBox txtAreaCode 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6540
         MaxLength       =   3
         TabIndex        =   49
         Top             =   870
         Width           =   465
      End
      Begin VB.TextBox txtPhone1 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   7170
         MaxLength       =   3
         TabIndex        =   48
         Top             =   870
         Width           =   465
      End
      Begin VB.TextBox txtPhone2 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   7785
         MaxLength       =   4
         TabIndex        =   47
         Top             =   870
         Width           =   700
      End
      Begin VB.TextBox txtCellPhone2 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   7785
         MaxLength       =   4
         TabIndex        =   46
         Top             =   1380
         Width           =   700
      End
      Begin VB.TextBox txtCellPhone1 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   7170
         MaxLength       =   3
         TabIndex        =   45
         Top             =   1380
         Width           =   465
      End
      Begin VB.TextBox txtCellAreaCode 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6540
         MaxLength       =   3
         TabIndex        =   44
         Top             =   1380
         Width           =   465
      End
      Begin VB.TextBox txtWorkPh2 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   7785
         MaxLength       =   4
         TabIndex        =   43
         Top             =   1800
         Width           =   700
      End
      Begin VB.TextBox txtWorkPh1 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   7170
         MaxLength       =   3
         TabIndex        =   42
         Top             =   1800
         Width           =   465
      End
      Begin VB.TextBox txtWorkAC 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6540
         MaxLength       =   3
         TabIndex        =   41
         Top             =   1800
         Width           =   465
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         FillColor       =   &H00314993&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00314993&
         Height          =   525
         Left            =   30
         Picture         =   "ResidentInfo.frx":279C
         ScaleHeight     =   525
         ScaleWidth      =   435
         TabIndex        =   40
         Top             =   720
         Width           =   435
      End
      Begin VB.TextBox txtFullName 
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
         Left            =   1770
         TabIndex        =   39
         Top             =   780
         Width           =   3165
      End
      Begin VB.TextBox txtAddress 
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
         Height          =   795
         Left            =   1740
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   1290
         Width           =   3135
      End
      Begin VB.TextBox txtAge 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4830
         MaxLength       =   3
         TabIndex        =   37
         Top             =   2730
         Width           =   600
      End
      Begin VB.TextBox txtSS3 
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
         Height          =   330
         Left            =   3540
         MaxLength       =   4
         TabIndex        =   36
         Top             =   3270
         Width           =   720
      End
      Begin VB.TextBox txtSS2 
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
         Height          =   330
         Left            =   2970
         MaxLength       =   2
         TabIndex        =   35
         Top             =   3270
         Width           =   390
      End
      Begin VB.TextBox txtSS1 
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
         Height          =   330
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   34
         Top             =   3270
         Width           =   510
      End
      Begin VB.TextBox txtDriverLicense 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   33
         Top             =   3690
         Width           =   1980
      End
      Begin VB.Frame fraGender 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   1980
         TabIndex        =   30
         Top             =   3930
         Width           =   2955
         Begin VB.OptionButton optFemale 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Female"
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
            Height          =   240
            Left            =   270
            TabIndex        =   32
            Top             =   150
            Width           =   1305
         End
         Begin VB.OptionButton optMale 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Male"
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
            Height          =   240
            Left            =   1770
            TabIndex        =   31
            Top             =   150
            Width           =   825
         End
      End
      Begin VB.TextBox txtBloodType 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1965
         TabIndex        =   29
         Top             =   4470
         Width           =   1545
      End
      Begin VB.ComboBox cboPatientRace 
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "ResidentInfo.frx":3066
         Left            =   1965
         List            =   "ResidentInfo.frx":3068
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   4890
         Width           =   2145
      End
      Begin VB.TextBox txtEmailAddress 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1965
         TabIndex        =   27
         Top             =   7140
         Width           =   2940
      End
      Begin VB.TextBox txtExtRef 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9720
         TabIndex        =   26
         ToolTipText     =   "Inmate Registration or Resident Registration Number"
         Top             =   3180
         Width           =   2265
      End
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   30
         Picture         =   "ResidentInfo.frx":306A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   25
         Top             =   1230
         Width           =   480
      End
      Begin VB.TextBox txtYear 
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
         Height          =   330
         Left            =   2850
         MaxLength       =   4
         TabIndex        =   24
         Top             =   2820
         Width           =   600
      End
      Begin VB.TextBox txtDay 
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
         Height          =   330
         Left            =   2250
         MaxLength       =   2
         TabIndex        =   23
         Top             =   2820
         Width           =   435
      End
      Begin VB.TextBox txtMonth 
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
         Height          =   330
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   22
         Top             =   2820
         Width           =   435
      End
      Begin VB.TextBox txtInternationalPhone 
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6480
         TabIndex        =   21
         Top             =   2190
         Width           =   2010
      End
      Begin VB.ComboBox cboPatMarital 
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "ResidentInfo.frx":3934
         Left            =   1965
         List            =   "ResidentInfo.frx":3936
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   5400
         Width           =   2655
      End
      Begin VB.ComboBox cboFacility 
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
         Height          =   330
         Left            =   7140
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3630
         Width           =   4905
      End
      Begin PT_XP_Button.PT_XP cmdName 
         Height          =   345
         Left            =   630
         TabIndex        =   53
         Top             =   780
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
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   "Name"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAddress 
         Height          =   405
         Left            =   630
         TabIndex        =   54
         Top             =   1290
         Width           =   1035
         _ExtentX        =   1826
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
         TransparentColor=   12582912
         Caption         =   "Address"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdNew 
         Height          =   405
         Left            =   3720
         TabIndex        =   55
         Top             =   6390
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "ResidentInfo.frx":3938
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   "New"
         CaptionPosition =   4
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "ResidentInfo.frx":3ED2
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdUpdate 
         Height          =   405
         Left            =   180
         TabIndex        =   56
         Top             =   90
         Width           =   1845
         _ExtentX        =   3254
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
         Picture         =   "ResidentInfo.frx":4224
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   "Update Patient"
         CaptionAlignment=   5
         CaptionPosition =   4
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "ResidentInfo.frx":47BE
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCancel 
         Height          =   405
         Left            =   6120
         TabIndex        =   57
         Top             =   6390
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
         Picture         =   "ResidentInfo.frx":4B10
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   "Cancel"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "ResidentInfo.frx":4EAA
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   405
         Left            =   7290
         TabIndex        =   58
         Top             =   6390
         Visible         =   0   'False
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
         Picture         =   "ResidentInfo.frx":51FC
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "ResidentInfo.frx":5C0E
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   405
         Left            =   4980
         TabIndex        =   59
         Top             =   6390
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
         Picture         =   "ResidentInfo.frx":5F60
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "ResidentInfo.frx":62FA
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdTakePic 
         Height          =   465
         Left            =   9030
         TabIndex        =   60
         Top             =   390
         Width           =   615
         _ExtentX        =   1085
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
         Picture         =   "ResidentInfo.frx":664C
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "ResidentInfo.frx":6F26
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdLangTerm 
         Height          =   225
         Left            =   5820
         TabIndex        =   109
         ToolTipText     =   "Manage Language Terms"
         Top             =   4620
         Width           =   375
         _ExtentX        =   661
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
         TransparentColor=   12582912
         Caption         =   "...."
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRace 
         Height          =   225
         Left            =   1500
         TabIndex        =   110
         ToolTipText     =   "Manage Race Terms"
         Top             =   4950
         Width           =   375
         _ExtentX        =   661
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
         TransparentColor=   12582912
         Caption         =   "...."
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin MSComCtl2.DTPicker ccdAdmitDate 
         Height          =   375
         Left            =   1980
         TabIndex        =   116
         Top             =   5910
         Width           =   2115
         _ExtentX        =   3731
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
         CustomFormat    =   "MM/dd/yyy"
         DateIsNull      =   -1  'True
         Format          =   61079553
         CurrentDate     =   38862.7355555556
         MinDate         =   37257
      End
      Begin VB.Label lblDateAdmission 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Admission"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   510
         TabIndex        =   115
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label lblBed 
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "Bunk No"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2850
         TabIndex        =   114
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label lblHouseUnit 
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "Housing Unit"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   540
         TabIndex        =   112
         Top             =   2310
         Width           =   1200
      End
      Begin VB.Label lblLang 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Language"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4590
         TabIndex        =   108
         Top             =   4665
         Width           =   855
      End
      Begin VB.Label lblPatient_id 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9720
         TabIndex        =   93
         Top             =   2520
         Width           =   2235
      End
      Begin VB.Label lblEMRNo 
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "EMR NUMBER"
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
         Left            =   8100
         TabIndex        =   92
         Top             =   2610
         Width           =   1515
      End
      Begin VB.Label lblFacRefNo 
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "Resident ID"
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
         Left            =   8130
         TabIndex        =   91
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label lblDash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Index           =   0
         Left            =   7680
         TabIndex        =   90
         Top             =   900
         Width           =   75
      End
      Begin VB.Label lblRightPaen 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")"
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
         Index           =   0
         Left            =   6945
         TabIndex        =   89
         Top             =   900
         Width           =   165
      End
      Begin VB.Label lblLeftParen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "("
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
         Index           =   0
         Left            =   6450
         TabIndex        =   88
         Top             =   900
         Width           =   75
      End
      Begin VB.Label lblPhone 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   5070
         TabIndex        =   87
         Top             =   900
         Width           =   1260
      End
      Begin VB.Label lblCell 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   5220
         TabIndex        =   86
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label lblRightPaen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")"
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
         Index           =   1
         Left            =   7035
         TabIndex        =   85
         Top             =   1395
         Width           =   165
      End
      Begin VB.Label lblDash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Index           =   2
         Left            =   7680
         TabIndex        =   84
         Top             =   1410
         Width           =   75
      End
      Begin VB.Label lblLeftParen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "("
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
         Index           =   1
         Left            =   6450
         TabIndex        =   83
         Top             =   1395
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Work Phone:"
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
         Left            =   5070
         TabIndex        =   82
         Top             =   1860
         Width           =   1230
      End
      Begin VB.Label lblDash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Index           =   3
         Left            =   7680
         TabIndex        =   81
         Top             =   1830
         Width           =   75
      End
      Begin VB.Label lblRightPaen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")"
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
         Index           =   3
         Left            =   7035
         TabIndex        =   80
         Top             =   1830
         Width           =   75
      End
      Begin VB.Label lblLeftParen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "("
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
         Index           =   3
         Left            =   6450
         TabIndex        =   79
         Top             =   1830
         Width           =   75
      End
      Begin VB.Label Label7 
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   540
         TabIndex        =   78
         Top             =   2730
         Width           =   1200
      End
      Begin VB.Label lblAge 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Age:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3630
         TabIndex        =   77
         Top             =   2790
         Width           =   1140
      End
      Begin VB.Label lblSocialSecurity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Social Security #"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   540
         TabIndex        =   76
         Top             =   3330
         Width           =   1470
      End
      Begin VB.Label lblDash 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   7
         Left            =   3390
         TabIndex        =   75
         Top             =   3300
         Width           =   150
      End
      Begin VB.Label lblDash 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   6
         Left            =   2820
         TabIndex        =   74
         Top             =   3300
         Width           =   150
      End
      Begin VB.Label lblDriverLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver License #"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   540
         TabIndex        =   73
         Top             =   3750
         Width           =   1410
      End
      Begin VB.Label lblGender 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   540
         TabIndex        =   72
         Top             =   4140
         Width           =   690
      End
      Begin VB.Label lblBloodType 
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Type"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   540
         TabIndex        =   71
         Top             =   4530
         Width           =   975
      End
      Begin VB.Label lblRace 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Race:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   540
         TabIndex        =   70
         Top             =   4965
         Width           =   465
      End
      Begin VB.Label lblEmailAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   540
         TabIndex        =   69
         Top             =   7260
         Width           =   570
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "yyyy"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2970
         TabIndex        =   68
         Top             =   2580
         Width           =   480
      End
      Begin VB.Label lblDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dd"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2310
         TabIndex        =   67
         Top             =   2595
         Width           =   210
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1710
         TabIndex        =   66
         Top             =   2595
         Width           =   360
      End
      Begin VB.Label lblSlash2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   2700
         TabIndex        =   65
         Top             =   2835
         Width           =   165
      End
      Begin VB.Label lblSlash1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   2130
         TabIndex        =   64
         Top             =   2835
         Width           =   135
      End
      Begin VB.Label lblInternational 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "International:"
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
         Left            =   4950
         TabIndex        =   63
         Top             =   2220
         Width           =   1425
      End
      Begin VB.Label lblMarital 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marital Status:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   540
         TabIndex        =   62
         Top             =   5490
         Width           =   1305
      End
      Begin VB.Label lblFacility 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Initial Facility"
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
         Height          =   330
         Left            =   5460
         TabIndex        =   61
         Top             =   3720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ResidentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public piPatientID As Long
Public piEncounterID As Long
Public piProviderID As Long
Public mstrConnectStringClinical As String
Public mstrConnectStringTerms As String
Private fUpdateMode As String
Private mLocID As Integer
Private mbAdmit As Boolean
Private mbDischarge As Boolean
Private mRoomAssign As String
Public mNewPatient As Boolean
Private mFirst As Boolean
Private bEnable As Boolean
Private mintRelationPatID  As Integer
Public piFacilityID As Long
Private mbooClickOnce  As Boolean
Public mbNewPatInsert As Boolean
Public mbNotSave As Boolean
Public mbooData As Boolean
Public mbNewPatReg As Boolean
Private mbooValidSS As Boolean

Private Function Check_ValidPhone() As Boolean
Dim i%
Dim booBadPhone As Boolean

    If Len(txtAreaCode.Text) < txtAreaCode.MaxLength And Len(txtAreaCode.Text) > 1 Then txtAreaCode.Text = vbNullString: booBadPhone = True
    If Len(txtPhone1.Text) < txtPhone1.MaxLength And Len(txtPhone1.Text) > 1 Then txtPhone1.Text = vbNullString: booBadPhone = True
    If Len(txtPhone2.Text) < txtPhone2.MaxLength And Len(txtPhone2.Text) > 1 Then txtPhone2.Text = vbNullString: booBadPhone = True
    
    If Len(txtCellAreaCode.Text) < txtCellAreaCode.MaxLength And Len(txtCellAreaCode.Text) > 1 Then txtCellAreaCode.Text = vbNullString: booBadPhone = True
    If Len(txtCellPhone1.Text) < txtCellPhone1.MaxLength And Len(txtCellPhone1.Text) > 1 Then txtCellPhone1.Text = vbNullString: booBadPhone = True
    If Len(txtCellPhone2.Text) < txtCellPhone2.MaxLength And Len(txtCellPhone2.Text) > 1 Then txtCellPhone2.Text = vbNullString: booBadPhone = True
    
    If Len(txtWorkAC.Text) < txtWorkAC.MaxLength And Len(txtWorkAC.Text) > 1 Then txtWorkAC.Text = vbNullString: booBadPhone = True
    If Len(txtWorkPh1.Text) < txtWorkPh1.MaxLength And Len(txtWorkPh1.Text) > 1 Then txtWorkPh1.Text = vbNullString: booBadPhone = True
    If Len(txtWorkPh2.Text) < txtWorkPh2.MaxLength And Len(txtWorkPh2.Text) > 1 Then txtWorkPh2.Text = vbNullString: booBadPhone = True
    Check_ValidPhone = True
    If booBadPhone Then
        Check_ValidPhone = False
        MsgBox "Incomplete Telephone Number. Please enter a correct Telephone Number!", vbInformation
        Exit Function
    End If
    Exit Function
    
End Function

Private Sub Load_Facilities()
    Dim i As Long
    Dim strSQL As String

    On Error GoTo Error_Load_Facilities
    
    If psDriver(0) = "SQL SERVER" Then
       strSQL = "Select distinct b.facility_name, a.facility_id, a.default_facility From Appointment_Providers a, FACILITIES b where a.facility_id = b.facility_id Order by facility_name"
    Else
       strSQL = "Select distinct b.facility_name, a.facility_id, a.default_facility From Appointment_Providers a, FACILITIES b where a.facility_id = b.facility_id Order by facility_name"
    End If
    Screen.MousePointer = vbHourglass
     Open_DB psConnect(0)
    Set rs = cn.Execute(strSQL)
    cboFacility.Clear
    i = -1
    Do Until rs.EOF
       cboFacility.AddItem Trim(rs!Facility_Name)
       cboFacility.ItemData(cboFacility.NewIndex) = (rs!FACILITY_ID)
       If (rs!Default_Facility) = "Y" Then i = cboFacility.NewIndex
       rs.MoveNext
    Loop
    Close_DB psConnect(0)
    If cboFacility.ListCount > 0 Then
       If i <> -1 Then cboFacility.ListIndex = i Else cboFacility.ListIndex = 0
    End If
    Screen.MousePointer = vbDefault
    mboobypass = False
    Exit Sub
Error_Load_Facilities:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbInformation
    Exit Sub
    Resume
End Sub

Private Function Get_Partner()
Dim strParName As String
Dim intActX_ID As Integer
Dim strSQL As String
Dim varRS
Dim lngrecs As Long
Dim strTitle As String
Dim booRet As Boolean
Dim i As Integer
Dim strProfPhone As String
Dim Arr
Dim varEmergContact
Dim varRelation
Dim varBloodType
Dim sParFName As String
Dim sParLName As String
Dim sParMName As String
Dim strRelID As String
Dim sSuffix As String
Dim lngRelID As Long

    mlngRelationPatID = 0
    
    Open_DB psConnect(0)
    strSQL = "Select a.title, a.first_name, a.middle_name, a.last_name, a.suffix, b.related_patient_id, b.RELATION_ID  " _
            & " FROM PATIENT a, PATIENT_RELATIONS b " _
            & "Where b.patient_id = " & piPatientID & " and a.patient_id = b.related_patient_id "
    Set rs = cn.Execute(strSQL)
    With rs
        If Not .EOF Then
            If (!Title) <> vbNullString Then strTitle = Trim(!Title)
            If (!First_Name) <> vbNullString Then sParFName = Trim(!First_Name)
            If (!Last_Name) <> vbNullString Then sParLName = Trim(!Last_Name)
            If (!Middle_Name) <> vbNullString Then sParMName = Trim(!Middle_Name)
            If (!Suffix) <> vbNullString Then sSuffix = Trim(!Suffix)
            If (!Related_Patient_Id) <> 0 Then mlngRelationPatID = (!Related_Patient_Id)
            If (!RELATION_ID) <> 0 Then lngRelID = (!RELATION_ID)
            strParName = Trim(sParFName) & " " & Trim(sParMName) & " " & Trim(sParLName)
            If cboSuffix.ListCount <= 0 Then
                LoadAttrValidValue psConnect(1), cboSuffix, "sd_Suffix"
             End If
            If Trim(sSuffix) <> vbNullString Then
                strParName = Trim(strParName) & ", " & Trim(sSuffix)
            End If
            If cboParTitle.ListCount <= 0 Then
                LoadAttrValidValue psConnect(1), cboParTitle, "sd_Suffix"
             End If
            For i = 0 To cboParTitle.ListCount - 1
                If cboParTitle.List(i) = Trim(strTitle) Then
                    cboParTitle.Text = cboParTitle.List(i)
                    Exit For
                End If
            Next i
        End If
    End With
    If mlngRelationPatID <> 0 Then
        txtParner.Text = Trim(strParName)
        txtParFirstName.Text = Trim(sParFName)
        txtParMidName.Text = Trim(sParMName)
        txtParLastName.Text = Trim(sParLName)
        For i = 0 To cboPartRel.ListCount - 1
            If cboPartRel.ItemData(i) = lngRelID Then
                cboPartRel.Text = cboPartRel.List(i)
                Exit For
            End If
        Next i
        For i = 0 To cboRelation.ListCount - 1
            If cboRelation.ItemData(i) = lngRelID Then
                cboRelation.Text = cboRelation.List(i)
                Exit For
            End If
        Next i
        For i = 0 To cboSuffix.ListCount - 1
            If cboSuffix.List(i) = Trim(sSuffix) Then
                cboSuffix.Text = cboSuffix.List(i)
                Exit For
            End If
        Next i
    Else
        txtParner.Text = vbNullString
    End If
    varBloodType = Get_ProfilerValue("FN_PP_PS", "PersonalInfoN", "cboBloodType", 1, piPatientID)
    If varBloodType <> vbNullString Then txtBloodType.Text = varBloodType
    If varBloodType = vbNullString Then
        varBloodType = Get_ProfilerValue("FN_PP_PS", "NPersonalInfo", "cboBloodType", 1, piPatientID)
    End If
    If varBloodType <> vbNullString Then txtBloodType.Text = varBloodType
    Close_DB psConnect(0)
    Exit Function

End Function
Private Sub Clear_Fields()
    On Error GoTo Error_Handler

    txtFullName.Text = vbNullString
    cboPatientRace.ListIndex = -1
    cboPatientRace.Clear
    txtEmailAddress.Text = vbNullString
    txtAreaCode.Text = vbNullString
    txtPhone1.Text = vbNullString
    txtPhone2.Text = vbNullString
    txtCellAreaCode.Text = vbNullString
    txtCellPhone1.Text = vbNullString
    txtCellPhone2.Text = vbNullString
    txtInternationalPhone.Text = vbNullString
    txtDriverLicense.Text = vbNullString
    txtMonth.Text = vbNullString
    txtYear.Text = vbNullString
    txtDay.Text = vbNullString
    txtAge.Text = vbNullString
    txtSS1.Text = vbNullString
    txtSS2.Text = vbNullString
    txtSS3.Text = vbNullString
    optMale.Value = False
    optFemale.Value = False
  '  cboTitle.Clear
    txtParner.Text = vbNullString
    txtEmergContact.Text = vbNullString
    cboRelation.Clear
    txtEmergAC.Text = vbNullString
    txtEmergPh1.Text = vbNullString
    txtEmergPh2.Text = vbNullString
    txtBloodType.Text = vbNullString
    lblPatient_id.Caption = vbNullString
    txtExtRef.Text = vbNullString
    Exit Sub
    
Error_Handler:
    MsgBox Err.Description, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub


Public Sub GetComboTerms(vobjCombo As Object, ByVal vstrSubDomain As String)
Dim strSQL As String
Dim booRet As Boolean
Dim varRS
Dim lngrecs As Long
Dim r%
    
    strSQL = "select  distinct attrbte_valid_Value, Valid_Value_Seq_Nbr from ATTR_VALID_VALUE where VALID_VALUE_SUBDOM = '" & vstrSubDomain & "'"
    booRet = SelectQuery(strSQL, psConnect(1), varRS, lngrecs)
    
    If booRet Then
        If lngrecs > 0 Then
            For r = LBound(varRS, 2) To UBound(varRS, 2)
                vobjCombo.AddItem varRS(0, r)
            Next r
        End If
    End If
    Exit Sub
End Sub


Private Function Update_Archive(vstrExtRef As String, vstrTitle As String, vstrPatFName As String, vstrPatMName As String, _
vstrPatLName As String, vstrPatAddress As String, vstrPatAddress2 As String, vstrPatCity As String, vstrPatState As String, vstrPatZip As String, _
vstrPatZipExt As String, vstrPatCountry As String, vstrAreaCode As String, vstrPhone As String, vstrCellPhone As String, vstrDL As String, _
vstrRace As String, vstrStatus As String, vstrPatSuffix As String, vstrPatType As String, vstrRoom As String, vstrRoomPhone As String, _
vstrBirth As String, vstrSocSec As String, vstrGender As String, vstrEmailAddress As String, vstrIntnPhone As String, vstrPatCounty As String, _
vstrBloodType As String, vstrMaritalStat As String, vstrWorkPhone As String, vstrEmerContact As String, vstrEMERG_PHONE As String, vstrEmergRel As String)
Dim cnPatient As New ADODB.Connection
Dim rsPatient As New ADODB.Recordset
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim booSave As Boolean
Dim strDate As String
Dim strChgDate As String
Dim lngDateDiff As Long
Dim booinsert As Boolean
Dim lngPrimKey As Long
Dim sDate As String

    If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
    strSQL = "SELECT * FROM PATIENT WHERE PATIENT_ID = " & piPatientID
    Set rsPatient = cnPatient.Execute(strSQL)
    If Not rsPatient.EOF Then
        Set Flds = rsPatient.Fields
        For Each Fld In Flds
            If Trim(Fld.Value) <> vbNullString Then
                If Fld.Name = UCase("Patient_Ext_Ref") And Fld.Value <> vstrExtRef Then booSave = True
                If Fld.Name = UCase("title") And Fld.Value <> vstrTitle Then booSave = True
                If Fld.Name = UCase("first_name") And Fld.Value <> vstrPatFName Then booSave = True
                If Fld.Name = UCase("middle_name") And Not Fld.Value <> vstrPatMName Then booSave = True
                If Fld.Name = UCase("last_name") And Fld.Value <> vstrPatLName Then booSave = True
                If Fld.Name = UCase("address1") And Fld.Value <> vstrPatAddress Then booSave = True
                If Fld.Name = UCase("address2") And Fld.Value <> vstrPatAddress2 Then booSave = True
                If Fld.Name = UCase("city") And Fld.Value <> vstrPatCity Then booSave = True
                If Fld.Name = UCase("state") And Fld.Value <> vstrPatState Then booSave = True
                If Fld.Name = UCase("zipcode") And Fld.Value <> vstrPatZip Then booSave = True
                If Fld.Name = UCase("zipcode_ext") And Fld.Value <> vstrPatZipExt Then booSave = True
                If Fld.Name = UCase("country") And Fld.Value <> vstrPatCountry Then booSave = True
                If Fld.Name = UCase("area_code") And Fld.Value <> vstrAreaCode Then booSave = True
                If Fld.Name = UCase("phone") And Fld.Value <> vstrPhone Then booSave = True
                If Fld.Name = UCase("cell_phone") And Fld.Value <> vstrCellPhone Then booSave = True
                If Fld.Name = UCase("drv_license") And Fld.Value <> vstrDL Then booSave = True
                If Fld.Name = UCase("Race") And Fld.Value <> vstrRace Then booSave = True
                If Fld.Name = UCase("Status") And Fld.Value <> vstrStatus Then booSave = True
                If Fld.Name = UCase("Suffix") And Fld.Value <> vstrPatSuffix Then booSave = True
                If Fld.Name = UCase("Patient_Type") And Fld.Value <> vstrPatType Then booSave = True
                If Fld.Name = UCase("Room_Number") And Fld.Value <> vstrRoom Then booSave = True
                If Fld.Name = UCase("Room_Phone") And Fld.Value <> vstrRoomPhone Then booSave = True
                If Fld.Name = UCase("birth") And Fld.Value <> vstrBirth Then booSave = True
                If Fld.Name = UCase("social_security") And Fld.Value <> vstrSocSec Then booSave = True
                If Fld.Name = UCase("gender") And Fld.Value <> vstrGender Then booSave = True
                If Fld.Name = UCase("email") And Fld.Value <> vstrEmailAddress Then booSave = True
                If Fld.Name = UCase("internatl_phone") And Fld.Value <> vstrIntnPhone Then booSave = True
                If Fld.Name = UCase("county") And Fld.Value <> vstrPatCounty Then booSave = True
                If Fld.Name = UCase("blood_Type") And Fld.Value <> vstrBloodType Then booSave = True
                If Fld.Name = UCase("marital_status") And Fld.Value <> vstrMaritalStat Then booSave = True
                If Fld.Name = UCase("work_phone") And Fld.Value <> vstrWorkPhone Then booSave = True
                If Fld.Name = UCase("EMERGENCY_CONTACT") And Fld.Value <> vstrEmerContact Then booSave = True
                If Fld.Name = UCase("EMERG_PHONE") And Fld.Value <> vstrEMERG_PHONE Then booSave = True
                If Fld.Name = UCase("EMERG_RELATION") And Fld.Value <> vstrEmergRel Then booSave = True
            End If
            If booSave Then Exit For
        Next
    End If
    If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
    strSQL = "exec S_PREVPATINFO " & piPatientID
    Set rsPatient = cnPatient.Execute(strSQL)
    sDate = Format(Now, "mm/dd/yyyy")
    With rsPatient
        If Not .EOF Then
            If (!CHANGE_DATE) <> vbNullString Then
                strChgDate = Format((!CHANGE_DATE), "mm/dd/yyyy")
                lngPrimKey = (!PREV_DEMO_ID)
                lngDateDiff = DateDiff("d", strChgDate, sDate)
                If lngDateDiff < 0 Then
                    lngPrimKey = 0
                End If
            End If
        End If
    End With

    If lngPrimKey > 0 Then
        cnPatient.U_PATARCHIVE clsPatient.ExtID, clsPatient.Title, clsPatient.FirstName, clsPatient.MiddleName, _
            clsPatient.LastName, clsPatient.Address, clsPatient.Address2, clsPatient.City, clsPatient.State, clsPatient.Zip, _
            clsPatient.Zip_Ext, clsPatient.Country, clsPatient.AreaCode, clsPatient.Phone, clsPatient.CellPhone, clsPatient.DL, _
            clsPatient.Race, clsPatient.Status, clsPatient.Suffix, clsPatient.PatientType, clsPatient.RoomNumber, vstrRoomPhone, _
            clsPatient.DOB, clsPatient.SS, clsPatient.Gender, clsPatient.Email, clsPatient.InterNalPhone, clsPatient.County, _
            clsPatient.Blood_Type, clsPatient.Marital_Status, clsPatient.WorkPhone, clsPatient.EmergContact, clsPatient.EmergPhone, clsPatient.EmergRelation, lngPrimKey
    Else
         cnPatient.I_PATARCHIVE clsPatient.ExtID, clsPatient.Title, clsPatient.FirstName, clsPatient.MiddleName, _
            clsPatient.LastName, clsPatient.Address, clsPatient.Address2, clsPatient.City, clsPatient.State, clsPatient.Zip, _
            clsPatient.Zip_Ext, clsPatient.Country, clsPatient.AreaCode, clsPatient.Phone, clsPatient.CellPhone, clsPatient.DL, _
            clsPatient.Race, clsPatient.Status, clsPatient.Suffix, clsPatient.PatientType, clsPatient.RoomNumber, vstrRoomPhone, _
            clsPatient.DOB, clsPatient.SS, clsPatient.Gender, clsPatient.Email, clsPatient.InterNalPhone, clsPatient.County, _
            clsPatient.Blood_Type, clsPatient.Marital_Status, clsPatient.WorkPhone, clsPatient.EmergContact, clsPatient.EmergPhone, clsPatient.EmergRelation, piPatientID
    End If

    If cnPatient.State = adStateOpen Then
        cnPatient.Close
        Set cnPatient = Nothing
        Set rsPatient = Nothing
    End If
    Exit Function
End Function

Private Function Update_Inform()
Dim cnInterface As New ADODB.Connection
Dim rsInterface As New ADODB.Recordset
Dim strSQL As String

    If cnInterface.State = adStateOpen Then cnInterface.Close
    'If psInformConnect = vbNullString Then Get_Inform
    cnInterface.open psInformConnect

    strSQL = " UPDATE tblpatient a, tblsalutation b SET a.label =  , a.First, a.Last, a.Middle, " _
'                & "a.EMail, a.SSN, a.BirthDate, a.Address1, a.Address2, a.City, " _
'                & "a.State, a.Zip, a.NightPhone, a.DayPhone, a.CellPhone , a.Fax, a.[M/F], a.DateofContact " _
'                & "FROM  tblpatient a, tblsalutation b " _
'                & "WHERE a.SalutationID = b.id ORDER BY a.DateofContact desc, a.Last,a.First, a.BirthDate " _
'                 , ConnPC
    
    Set rsInterface = cnInterface.Execute(strSQL)
    If cnInterface.State = adStateOpen Then
        cnInterface.Close
        Set cnInterface = Nothing
        Set rsInterface = Nothing
    End If

End Function

Private Function Update_ThirdParty()
Dim cnInterface As New ADODB.Connection
Dim rsInterface As New ADODB.Recordset
Dim strSQL As String


    If cnInterface.State = adStateOpen Then cnInterface.Close
    cnInterface.open psConnect(0)
    strSQL = "SELECT * from INTERFACE_CONTROL "
    Set rsInterface = cnInterface.Execute(strSQL)
    With rsInterface
        Do Until .EOF
            If (!INFORM) = 1 Then 'True
                Update_Inform
                Exit Do
            ElseIf (!MEDICAL_MANGER) = True Then
            
            
            ElseIf (!PRIME_CLINICAL) = True Then
            
            End If
            .MoveNext
        Loop
    End With
    If cnInterface.State = adStateOpen Then cnInterface.Close
    Exit Function
End Function

Private Function Validate_Entries() As Boolean
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim iAnswer As Integer
    Dim Date1 As Date
    Dim Date2 As Date
    Dim Arr
    Dim i%
    Validate_Entries = True
    Arr = Split(txtFullName.Text, " ")
    If UBound(Arr) < 1 Then
      MsgBox "First name And last name required", vbExclamation
        Validate_Entries = False
    End If
 
    If Trim(txtMonth.Text) <> vbNullString Then
       If Not IsNumeric(txtMonth.Text) Then
          MsgBox "Invalid Month for Birth Date", vbExclamation
          Validate_Entries = False
          Exit Function
       End If
       If CInt(txtMonth.Text) < 1 Or CInt(txtMonth.Text) > 12 Then
          MsgBox "Invalid Month for Birth Date", vbExclamation
          Validate_Entries = False
          Exit Function
       End If
    End If
    
    If Trim(txtYear.Text) <> vbNullString Then
       If Not IsNumeric(txtYear.Text) Or Len(txtYear.Text) <> 4 Then
          MsgBox "Invalid Year for Birth Date", vbExclamation
          Validate_Entries = False
          Exit Function
       End If
       If CInt(txtYear.Text) > Year(Now) Then
          MsgBox "Invalid Year for Birth Date", vbExclamation
          Validate_Entries = False
          Exit Function
       End If
    End If
    
    If Trim(txtDay.Text) <> vbNullString Then
       If Not IsNumeric(txtDay.Text) Then
          MsgBox "Invalid Day for Birth Date", vbExclamation
          Validate_Entries = False
          Exit Function
       End If
       If Not IsDayCheck(CInt(txtMonth), CInt(txtDay), CInt(txtYear)) Then
          MsgBox "Invalid Day for Birth Date", vbExclamation
          Validate_Entries = False
          Exit Function
       End If
       If mstrPractice = "Psychiatry" Then
            If Val(txtAge.Text) <= 18 And Trim(txtParFirstName.Text) = vbNullString Then
                 MsgBox "Patient's Parent or Legal Guardian is a mandatory field, Please complete field.!", vbInformation
                 lblPartnerName.Caption = "Legal Guardian"
                 Validate_Entries = False
                 txtParner.SetFocus
                 Exit Function
             End If
        End If
    End If
        If Trim(txtEmergPh2.Text) <> vbNullString And Len(txtEmergPh2.Text) < 4 Then
             Call MsgBox("Invalid data entry ! Telephone fields require complete numeric data. Erroneous data will be removed. Please re-enter valid data.  ", vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1, "Patient Demographic Maintanence")
             txtEmergPh2.Text = vbNullString
            Validate_Entries = False
            txtEmergPh2.SetFocus
            Exit Function
         End If
        If Trim(txtEmergPh1.Text) <> vbNullString And Len(txtEmergPh1.Text) < 3 Then
            Call MsgBox("Invalid data entry ! Telephone fields require numeric data. Erroneous data will not be removed. Please re-enter valid data.  ", vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1, "Patient Demographic Maintanence")
            txtEmergPh1.Text = vbNullString
            txtEmergPh2.SetFocus
            Validate_Entries = False
            Exit Function
        End If
Exit Function
End Function

Private Sub cboFacility_Click()
Dim i%
Dim strSQL As String
Dim strFacilityName As String

    For i = 0 To cboFacility.ListCount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            strFacilityName = cboFacility.Text
                If piFacilityID = 0 Then
                    strSQL = "select FACILITY_ID from Facilities where facility_name = '" & strFacilityName & "'"
                    Open_DB psConnect(0)
                    Set rs = cn.Execute(strSQL)
                    With rs
                        If Not .EOF Then
                            piFacilityID = (!FACILITY_ID)
                        End If
                    End With
                    Close_DB psConnect(0)
                End If
        End If
    Next i
Exit Sub
End Sub

Private Sub cboParTitle_KeyUp(KeyCode As Integer, Shift As Integer)
Dim k As Integer
Dim lRow As Long
Dim sItem As String
        If cboParTitle.Text <> vbNullString Then
            SearchList cboParTitle, cboParTitle.Text, False
        End If
End Sub


Private Sub cboPatientRace_Click()
    If cboPatientRace.ListCount < 0 Then
        cboPatientRace.Clear
        GetTerms cboPatientRace, "sd_Patient_Race"
    End If
    Exit Sub
End Sub


Private Sub cboPatientRace_KeyUp(KeyCode As Integer, Shift As Integer)
Dim k As Integer
Dim lRow As Long
Dim sItem As String
        If cboPatientRace.Text <> vbNullString Then
            SearchList cboPatientRace, cboPatientRace.Text, False
        End If
    Exit Sub
End Sub

Private Sub cboPatMarital_KeyUp(KeyCode As Integer, Shift As Integer)
Dim k As Integer
Dim lRow As Long
Dim sItem As String
        If cboPatMarital.Text <> vbNullString Then
            SearchList cboPatMarital, cboPatMarital.Text, False
        End If
End Sub


Private Sub cboRelation_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 1 Then cboRelation.ListIndex = -1
End Sub


Private Sub cmdAddress_Click()
Dim strAddress As String
    If mboobypass Then mboobypass = False: Exit Sub
    txtAddress.Text = vbNullString
    frmPatientAdress.Show vbModal
    If frmPatientAdress.mstrPatAddress <> vbNullString Then
        If Trim(frmPatientAdress.mstrPatAddress) <> vbNullString Then strAddress = Trim(frmPatientAdress.mstrPatAddress)
        If Trim(frmPatientAdress.mstrPatAddress2) <> vbNullString Then strAddress = Trim(strAddress) & ", " & Trim(frmPatientAdress.mstrPatAddress2)
        If Trim(frmPatientAdress.mstrPatCity) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(frmPatientAdress.mstrPatCity)
        If Trim(frmPatientAdress.mstrPatState) <> vbNullString Then strAddress = Trim(strAddress) & ", " & Trim(frmPatientAdress.mstrPatState)
        If Trim(frmPatientAdress.mstrPatZip) <> vbNullString Then strAddress = Trim(strAddress) & " " & Trim(frmPatientAdress.mstrPatZip)
        If Trim(frmPatientAdress.mstrPatCountry) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(frmPatientAdress.mstrPatCountry)
        If Trim(frmPatientAdress.mstrPatCounty) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(frmPatientAdress.mstrPatCounty)
        txtAddress.Text = Trim(strAddress)
    End If
    
    Exit Sub
    
End Sub


Private Sub cmdCancel_Click()
    mNewPatient = False
    Clear_Fields
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    piPatientID = 0
    Unload Me
Exit Sub
End Sub

Private Sub cmdClose_Click()
    If mNewPatient = True Then piPatientID = 0
    mNewPatient = False
    cmdNew.Enabled = False
    mbooClickOnce = True
    Unload Me
End Sub


Private Sub cmdLangTerm_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Patient Language"
    frmNurseTemplates.fType = "Patient Languages"
    frmNurseTemplates.fSubDomain = "sd_Patient_Language"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    GetTerms cboLanguage, "sd_Patient_Language"
End Sub

Private Sub cmdName_Click()


    If mboobypass Then mboobypass = False: Exit Sub
    txtFullName.Text = vbNullString
    If mNewPatient Then
        frmPatientName.mstrPatTitle = vbNullString
        frmPatientName.mstrPatFName = vbNullString
        frmPatientName.mstrPatMName = vbNullString
        frmPatientName.mstrPatLName = vbNullString
        frmPatientName.mstrPatSuffix = vbNullString
    End If
   frmPatientName.Show vbModal
        If Trim(frmPatientName.mstrPatTitle) <> vbNullString Then txtFullName.Text = Trim(frmPatientName.mstrPatTitle)
        If Trim(frmPatientName.mstrPatFName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(frmPatientName.mstrPatFName)
        If Trim(frmPatientName.mstrPatMName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(frmPatientName.mstrPatMName)
        If Trim(frmPatientName.mstrPatLName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(frmPatientName.mstrPatLName)
        If Trim(frmPatientName.mstrPatSuffix) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & ", " & Trim(frmPatientName.mstrPatSuffix)
    Exit Sub
End Sub

Private Sub cmdNew_Click()
Dim sStr As String
Dim strSQL As String
Dim strState As String
    
    mNewPatient = True
    fUpdateMode = "New"
    Clear_Fields
    piPatientID = 0
    bEnable = True
    Screen_Action
    GetTerms cboPatientRace, "sd_Patient_Race"

    cmdUpdate.Visible = False
    cmdSave.Visible = True
    cmdSave.Left = 5400
    cmdCancel.Enabled = True
    fUpdateMode = "Update"
End Sub

Public Sub SaveNotes _
 ( _
 ByRef rconClinical As ADODB.Connection, _
 ByVal vstrValue As String, _
 ByVal vstrRef As String, _
 ByVal vintSequence As Integer, _
 ByVal vbooPrimary As Boolean, _
 ByVal strOCX As String, _
 ByVal strControl _
 )


    'Set error handler
    On Error GoTo Error_SaveNotes
    
    LogFile "Public Sub SaveNotes" & String(30, "-")

    LogFile "A_PATIENT_PROFILER_NOTES > " & vstrRef

    If strControl = "PersonalInfoN" Then
            rconClinical.A_PATIENT_PROFILER_NOTES _
            "X", piPatientID, strOCX & "." & strControl, vstrRef, vintSequence, vstrValue, vbooPrimary
    End If
    
    If strControl = "PersonalInfoN" Then
            rconClinical.A_PATIENT_PROFILER _
            "X", piPatientID, strOCX & "." & strControl, vstrRef, vintSequence, vstrValue, vbooPrimary
    End If
    Exit Sub
    
    

Error_SaveNotes:

    ErrHandler _
     "frmPatientInfo.SaveNotes", _
     Err.Number, _
     Err.Description
     Exit Sub
     Resume
End Sub
    


Private Sub cmdParOk_Click()
Dim strFName As String
Dim strLName As String

    strFName = Trim(txtParFirstName.Text)
    strLName = Trim(txtParLastName.Text)
    If txtParLastName.Text <> vbNullString Then
        txtParner.Text = Trim(cboParTitle.Text) & " " & Trim(txtParFirstName.Text) & " " & Trim(txtParMidName.Text) _
            & " " & Trim(txtParLastName.Text)
        If cboSuffix.Text <> vbNullString Then
            txtParner.Text = Trim(txtParner.Text) & ", " & Trim(cboSuffix.Text)
        End If
    End If
    
'    If cboPartRel.Text = vbNullString And Trim(txtParner.Text) <> vbNullString Then
'        MsgBox "Patient's" & " " & lblPartTitle.Caption & " " & "relationship is a mandatory field, Please complete field!", vbInformation
'        Exit Sub
'    End If
    If strFName = vbNullString Or strLName = vbNullString Then
        MsgBox "Please complete the full name and relationship for Partner, otherwise select CANCEL!", vbInformation
        Exit Sub
    End If
    fraPartName.Visible = False

    Exit Sub

End Sub

Private Sub cmdPartCancel_Click()
        fraPartName.Visible = False
End Sub


Private Sub cmdRace_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Patient Race"
    frmNurseTemplates.fType = "Patient Race"
    frmNurseTemplates.fSubDomain = "sd_Patient_Race"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    GetTerms cboPatientRace, "sd_Patient_Race"
End Sub

Private Sub cmdSave_Click()
    Dim intIndex As Integer
        On Error GoTo Error_Handler
        
    If cboFacility.Text = vbNullString Then
        MsgBox "Please select initial facililty. ", vbInformation
        Exit Sub
    End If
    

    If Validate_Entries() Then
       If mNewPatient Then
          piPatientID = 0
          piPatientID = Save_Data()
            If Not mNewPatient Then
                mNewPatient = False
                Unload Me
                If mbNotSave Then Exit Sub
            Else
                If mbNotSave Then Exit Sub
            End If
       Else
          UpdatePatientInfo
          mNewPatient = False
          If mbNotSave Then Exit Sub
       End If
    Else
       Exit Sub
    End If
    Write_PatientProfiler
    mNewPatient = False
    mbooClickOnce = True
   Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
    
Error_Handler:
    MsgBox Err.Description, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub



Private Sub cmdUpdate_Click()
Dim sStr As String
Dim strSQL As String
Dim strState As String
    
    mNewPatient = False
    fUpdateMode = "Update"
    bEnable = True
    Screen_Action
    txtAge.Enabled = False
    If Trim(cboPatientRace.Text) = vbNullString Then
        If cboPatientRace.ListCount <= 0 Then
            LoadAttrValidValue psConnect(1), cboPatientRace, "sd_Patient_Race"
        End If
    End If

    cmdSave.ZOrder 0
    cmdSave.Visible = True
    cmdSave.Enabled = True
    cmdSave.Left = cmdNew.Left
    cmdNew.Visible = False
    Exit Sub
End Sub

Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
    centerForm Me
    End If
    Exit Sub
End Sub

Private Function Save_Data() As Long
Dim strExtRef As String
Dim strFields As String
Dim strValues As String
Dim strPhone As String
Dim strCellPhone As String
Dim strBirth As String
Dim strSocSec As String
Dim bRet As Boolean
Dim strGender As String
Dim strRoom As String
Dim lPatID As Long
Dim strStatus As String
Dim strRoomPhone As String
Dim strPatType  As String
Dim strBloodType As String
Dim strMaritalStatus As String
Dim strWorkPhone As String
Dim strAreaCode As String
Dim strSQL As String
Dim strSSFound As String
Dim iAnswer As Integer
Dim sTestLast As String
Dim sFirstTest As String
Dim sEmerContact As String
Dim sEmerRel As String
Dim sEmrPhone As String
Dim booSS As Boolean


    On Error GoTo Error_Save_Data
    
'    If mNewPatient = False Then
'        If piPatientID = 0 Then
'            MsgBox "No Patient Entered. Exiting now!", vbInformation
'            Exit Function
'        End If
'    End If
'    frmPatientName.mstrPatLName = Literals(frmPatientName.mstrPatLName, "'")
'    frmPatientName.mstrPatFName = Literals(frmPatientName.mstrPatFName, "'")
'    ' 999-99-9099 = Unknown Social Security Number
'    Validate_Entries
'    mbooData = True
'
'    strBloodType = Literals(Trim(txtBloodType), "'")
'    strExtRef = Literals(Trim(txtExtRef), "'")
'
'    strMaritalStatus = cboPatMarital.Text
'
'    Open_DB psConnect(0)
'
'    If txtSS1.Text <> vbNullString Then
'        strSocSec = txtSS1.Text & "-" & txtSS2.Text & "-" & txtSS3.Text
'    Else
'         Call MsgBox("Social Security Number is required. Please enter Social Security Number!  ", vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1, "Patient Demographic Maintanence")
'           mbNewPatReg = False
'           mbNotSave = True
'          GoTo Exit_Function
'    End If
'
'    Open_DB psConnect(0)
'    strSQL = "SELECT FIRST_NAME, LAST_NAME, SOCIAL_SECURITY from PATIENT WHERE SOCIAL_SECURITY = '" & strSocSec & "'"
'    Set rs = cn.Execute(strSQL)
'    If Not rs.EOF Then
'        If rs.Fields("SOCIAL_SECURITY") <> vbNullString And rs.Fields("SOCIAL_SECURITY") <> "999-99-9999" Then
'            If rs.Fields("LAST_NAME") <> vbNullString Then
'                If TextFound("''", rs.Fields("LAST_NAME"), 2, True) Then
'                    sTestLast = Replace(rs.Fields("LAST_NAME"), "''", "'")
'                Else
'                    sTestLast = rs.Fields("LAST_NAME")
'                End If
'                If TextFound("''", rs.Fields("FIRST_NAME"), 2, True) Then
'                    sFirstTest = Replace(rs.Fields("FIRST_NAME"), "''", "'")
'                Else
'                    sFirstTest = rs.Fields("FIRST_NAME")
'                End If
'                If TextFound(sTestLast, frmPatientName.mstrPatLName, Len(frmPatientName.mstrPatLName), True) Then booSS = True
'                If TextFound(sFirstTest, frmPatientName.mstrPatFName, Len(frmPatientName.mstrPatFName), True) Then booSS = True
'                If booSS Then
'                        strSSFound = sFirstTest & " " & sTestLast & " " & "has this social security number in PatientTrac records!" _
'                            & " " & "Cannot save this patient."
'                    iAnswer = MsgBox(strSSFound & " " & "Do you want to re-enter a valid social security number?", vbYesNoCancel + vbQuestion + vbApplicationModal + vbDefaultButton1, "Patient Maintance")
'                    If iAnswer = vbYes Then
'                        mbNotSave = True
'                        mbNewPatReg = False
'                        Close_DB psConnect(0)
'                        txtSS1.Text = vbNullString
'                        txtSS2.Text = vbNullString
'                        txtSS3.Text = vbNullString
'                        Close_DB psConnect(0)
'                        Exit Function
'                    ElseIf iAnswer = vbNo Then
'                        Close_DB psConnect(0)
'                        mbNewPatReg = False
'                        txtSS1.Text = vbNullString
'                        txtSS2.Text = vbNullString
'                        txtSS3.Text = vbNullString
'                        GoTo Exit_Function
'                    ElseIf iAnswer = vbCancel Then
'                        Close_DB psConnect(0)
'                        txtSS1.Text = vbNullString
'                        txtSS2.Text = vbNullString
'                        txtSS3.Text = vbNullString
'                        mbNewPatReg = False
'                        mbNotSave = True
'                        Exit Function
'                    End If
'                End If
'            End If
'        End If
'    End If
'   If strSocSec <> vbNullString Then
'        Set rs = Nothing
'        cn.S_ISPATIENT UCase(frmPatientName.mstrPatFName), UCase(frmPatientName.mstrPatLName), strSocSec, rs
'    Else
'        Call MsgBox("Social Security Number is required. Please enter Social Security Number!", vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1, "Patient Demographic Maintanence")
'           mbNotSave = True
'            mbNewPatReg = False
'          GoTo Exit_Function
'    End If
'    With rs
'       If Not .EOF Then
'          Save_Data = (!Patient_ID)
'          mNewPatient = False
'          MsgBox "Patient already exists!", vbExclamation
'          mbNewPatReg = False
'          GoTo Exit_Function
'        End If
'        .Close
'    End With
'
'    'Title---------------
'
'    'cboTitle.Text = strTitle
'    If Check_ValidPhone Then
'        If txtAreaCode.Text <> vbNullString Then strAreaCode = txtAreaCode.Text
'        If txtPhone1.Text <> vbNullString Then strPhone = Trim(txtPhone1.Text)
'        If txtPhone2.Text <> vbNullString Then strPhone = Trim(strPhone) & "-" & Trim(txtPhone2.Text)
'        If txtCellAreaCode.Text <> vbNullString Then strCellPhone = Trim(txtCellAreaCode.Text)
'        If txtCellPhone1.Text <> vbNullString Then strCellPhone = Trim(strCellPhone) & "-" & txtCellPhone1 & "-" & txtCellPhone2
'        If txtWorkAC.Text <> vbNullString Then strWorkPhone = txtWorkAC
'        If txtWorkPh1.Text <> vbNullString Then strWorkPhone = Trim(strWorkPhone) & "-" & txtWorkPh1 & "-" & txtWorkPh2
'    End If
'    If txtMonth.Text <> vbNullString Then strBirth = txtMonth.Text
'    If txtDay.Text <> vbNullString Then strBirth = Trim(strBirth) & "/" & txtDay.Text
'    If txtYear.Text <> vbNullString Then strBirth = Trim(strBirth) & "/" & txtYear.Text
'    If Trim(strBirth) <> vbNullString Then strBirth = Format(strBirth, "mm/dd/yyyy hh:mm:ss AM/PM")
'
'    If txtSS1.Text <> vbNullString Then
'        txtSS1.Text = Format(txtSS1.Text, "000")
'        txtSS2.Text = Format(txtSS2.Text, "00")
'        txtSS3.Text = Format(txtSS3.Text, "0000")
'        strSocSec = txtSS1.Text & "-" & txtSS2.Text & "-" & txtSS3.Text
'    End If
'    If Trim(strBirth) = vbNullString Then
'        mbNotSave = True
'        MsgBox "Date of Birth is mandatory field!", vbInformation
'        Exit Function
'    ElseIf Trim(strPhone) = vbNullString Then
'        mbNotSave = True
'        MsgBox "Telephone is mandatory field!", vbInformation
'        Exit Function
'    End If
'    If optMale.Value Then strGender = "M"
'    If optFemale.Value Then strGender = "F"
'
'    strRoom = vbNullString
'    strStatus = vbNullString
'    strRoomPhone = vbNullString
'    strPatType = vbNullString
'    strRoom = vbNullString
'
'    frmPatientName.mstrPatTitle = StrConv(frmPatientName.mstrPatTitle, vbProperCase)
'    frmPatientName.mstrPatFName = StrConv(frmPatientName.mstrPatFName, vbProperCase)
'    frmPatientName.mstrPatMName = StrConv(frmPatientName.mstrPatMName, vbProperCase)
'    frmPatientName.mstrPatLName = StrConv(frmPatientName.mstrPatLName, vbProperCase)
'    frmPatientName.mstrPatSuffix = StrConv(frmPatientName.mstrPatSuffix, vbProperCase)
'
'    Open_DB psConnect(0)
'    cn.i_Patient strExtRef, frmPatientName.mstrPatTitle, frmPatientName.mstrPatFName, frmPatientName.mstrPatMName, _
'                frmPatientName.mstrPatLName, frmPatientAdress.mstrPatAddress, frmPatientAdress.mstrPatAddress2, _
'                frmPatientAdress.mstrPatCity, frmPatientAdress.mstrPatState, frmPatientAdress.mstrPatZip, _
'                frmPatientAdress.mstrPatZipExt, frmPatientAdress.mstrPatCountry, Literals(strAreaCode, "'"), _
'                Literals(strPhone, "'"), Literals(strCellPhone, "'"), _
'                Literals(txtDriverLicense, "'"), Literals(cboPatientRace.Text, "'"), strStatus, _
'                frmPatientName.mstrPatSuffix, strPatType, strRoom, strRoomPhone, _
'                strBirth, Literals(strSocSec, "'"), strGender, Literals(txtEmailAddress, "'"), _
'                Literals(txtInternationalPhone, "'"), frmPatientAdress.mstrPatCounty, strBloodType, strMaritalStatus, strWorkPhone
'
'    Open_DB psConnect(0)
'    Set rs = Nothing
'    cn.S_ISPATIENT UCase(frmPatientName.mstrPatFName), UCase(frmPatientName.mstrPatLName), strSocSec, rs
'    With rs
'       If Not .EOF Then
'            lPatID = (!Patient_ID)
'            mbNotSave = False
'            mbNewPatReg = True
'       Else
'            MsgBox "Problem retreiving patient information!", vbCritical
'            mbNewPatReg = False
'            GoTo Exit_Function
'       End If
'    End With
'
'    Save_Data = lPatID
'    If lPatID <> 0 Then
'        Open_DB psConnect(0)
'        mbNotSave = False
'        strSQL = "Update patient set purge = 0 where patient_id = " & lPatID
'        bRet = ActionQuery(strSQL, psConnect(0))
'    End If
'    If Trim(txtEmergContact.Text) <> vbNullString Then
'        sEmerContact = Trim(txtEmergContact.Text)
'        If sEmerContact <> vbNullString Then sEmerContact = StrConv(sEmerContact, vbProperCase)
'        sEmerRel = Trim(cboRelation.Text)
'        sEmrPhone = Trim(txtEmergAC.Text) & "-" & Trim(txtEmergPh1.Text) & "-" & Trim(txtEmergPh2.Text)
'        strSQL = "UPDATE PATIENT SET EMERGENCY_CONTACT = '" & sEmerContact & "',EMERG_PHONE = '" & sEmrPhone & "', " _
'             & "EMERG_RELATION = '" & sEmerRel & "' WHERE PATIENT_ID = " & lPatID
'        Set rs = cn.Execute(strSQL)
'    End If
'    If Trim(cboLanguage.Text) <> vbNullString Then
'        clsPatient.Language = Trim(cboLanguage.Text)
'        strSQL = "UPDATE PATIENT SET LANGUAGE = '" & clsPatient.Language & "' WHERE PATIENT_ID =   " & lPatID
'        Set rs = cn.Execute(strSQL)
'    End If
'    Close_DB psConnect(0)
'    frmPatientName.mstrPatTitle = vbNullString
'    frmPatientName.mstrPatFName = vbNullString
'    frmPatientName.mstrPatMName = vbNullString
'    frmPatientName.mstrPatLName = vbNullString
'    frmPatientName.mstrPatSuffix = vbNullString
'    Close_DB psConnect(0)
'
'    Exit Function
'Exit_Function:
'    Close_DB psConnect(0)
'    Exit Function
    
Error_Save_Data:
     Screen.MousePointer = vbDefault
    Call MsgBox(Err.Description & " " & "PatientTrac encountered a problem saving patient data. Please contact technical support!. Thank You.", vbOKOnly + vbExclamation + vbApplicationModal + vbDefaultButton1, "Patient Demographic Maintanence")
     Exit Function
     Resume
End Function

Public Sub UpdatePatientInfo()
Dim strExtRef As String
Dim strFields As String
Dim strValues As String
Dim strPhone As String
Dim strCellPhone As String
Dim strBirth As String
Dim strSocSec As String
Dim strGender As String
Dim lPatID As Long
Dim strRoom As String
Dim strStatus As String
Dim strRoomPhone As String
Dim strPatType  As String
Dim strBloodType As String
Dim strMaritalStatus As String
Dim strWorkPhone As String
Dim iAnswer As Integer
Dim strSQL As String
Dim strCount As String
Dim strAreaCode As String
Dim sEmerContact As String
Dim sEmerRel As String
Dim sEmrPhone As String
Dim sTestLast As String
Dim sFirstTest As String
Dim strSSFound As String
Dim Arr
Dim booSS As Boolean
Dim lngPatTestID As Long

    On Error GoTo Error_UpdatePatientInfo
    
    If piPatientID = 0 Then Exit Sub
    Get_Patient piPatientID
    frmPatientName.mstrPatLName = Literals(frmPatientName.mstrPatLName, "'")
    frmPatientName.mstrPatFName = Literals(frmPatientName.mstrPatFName, "'")
    Validate_Entries
    
    Open_DB psConnect(0)
    
    strSQL = "Select first_name, last_name, patient_id from Patient where last_name  = '" & frmPatientName.mstrPatLName & "' and patient_id = " & piPatientID
     
    Set rs = cn.Execute(strSQL)
    strCount = 0
    With rs
        If Not .EOF Then
            If (!Patient_ID) = piPatientID Then
               ' MsgBox "Patient validated!", vbInformation
                strCount = strCount + 1
            End If
        End If
    End With
    If strCount < 1 Then
        iAnswer = MsgBox("Do you want to change the Last Name for" & " " & frmPatientName.mstrPatFName & " " & frmPatientName.mstrPatLName & "?", vbYesNo)
        If iAnswer = vbNo Then
            Unload Me
            Exit Sub
        End If
    End If
    strExtRef = Replace(Trim(txtExtRef), "'", "''")

    If cboPatMarital.Text <> vbNullString Then strMaritalStatus = Trim(cboPatMarital.Text)
    If Check_ValidPhone Then
        If txtAreaCode.Text <> vbNullString Then strAreaCode = txtAreaCode.Text
        If txtPhone1.Text <> vbNullString Then strPhone = Trim(txtPhone1.Text)
        If txtPhone2.Text <> vbNullString Then strPhone = Trim(strPhone) & "-" & Trim(txtPhone2.Text)
        If txtCellAreaCode.Text <> vbNullString Then strCellPhone = Trim(txtCellAreaCode.Text)
        If txtCellPhone1.Text <> vbNullString Then strCellPhone = Trim(strCellPhone) & "-" & txtCellPhone1.Text & "-" & txtCellPhone2.Text
        If txtWorkAC.Text <> vbNullString Then strWorkPhone = txtWorkAC.Text
        If txtWorkPh1.Text <> vbNullString Then strWorkPhone = Trim(strWorkPhone) & "-" & txtWorkPh1.Text & "-" & txtWorkPh2.Text
    End If
    If txtMonth.Text <> vbNullString Then strBirth = txtMonth.Text
    If txtDay.Text <> vbNullString Then strBirth = Trim(strBirth) & "/" & txtDay.Text
    If txtYear.Text <> vbNullString Then strBirth = Trim(strBirth) & "/" & txtYear.Text
    If Trim(strBirth) <> vbNullString Then strBirth = Format(strBirth, "mm/dd/yyyy hh:mm:ss AM/PM")
    If txtSS1.Text <> vbNullString Then strSocSec = txtSS1.Text & "-" & txtSS2.Text & "-" & txtSS3.Text
    ' -------------------------------------------------------------------------------------
    Get_Patient piPatientID
    Open_DB psConnect(0)
    strSQL = "SELECT FIRST_NAME, LAST_NAME, SOCIAL_SECURITY, PATIENT_ID from PATIENT WHERE SOCIAL_SECURITY = '" & strSocSec & "'"
    Set rs = cn.Execute(strSQL)
    If Not rs.EOF Then
        If rs.Fields("SOCIAL_SECURITY") <> vbNullString And rs.Fields("SOCIAL_SECURITY") <> "999-99-9999" Then
            If rs.Fields("LAST_NAME") <> vbNullString Then
                If Not TextFound("''", rs.Fields("LAST_NAME"), 2, True) Then
                    sTestLast = Replace(rs.Fields("LAST_NAME"), "''", "'")
                End If
                If Not TextFound("''", rs.Fields("FIRST_NAME"), 2, True) Then
                    sFirstTest = Replace(rs.Fields("FIRST_NAME"), "''", "'")
                End If
                lngPatTestID = rs.Fields("PATIENT_ID")
                booSS = False
                If Not TextFound(sTestLast, clsPatient.LastName, Len(clsPatient.LastName), True) And lngPatTestID <> piPatientID Then booSS = True
                If Not TextFound(sFirstTest, clsPatient.FirstName, Len(clsPatient.FirstName), True) And lngPatTestID <> piPatientID Then booSS = True
               If booSS And Trim(strSocSec) <> "999-99-9999" Then
                    strSSFound = sFirstTest & " " & sTestLast & " " & "has this social security number in PatientTrac records!" _
                        & " " & "Cannot save this patient."
                    iAnswer = MsgBox(strSSFound & " " & "Do you want to re-enter a valid social security number?", vbYesNoCancel + vbQuestion + vbApplicationModal + vbDefaultButton1, "Patient Maintance")
                    If iAnswer = vbYes Then
                        mbNotSave = True
                        txtSS1.Text = vbNullString
                        txtSS2.Text = vbNullString
                        txtSS3.Text = vbNullString
                        Exit Sub
                    ElseIf iAnswer = vbNo Then
                        If clsPatient.SS <> vbNullString Then strSocSec = clsPatient.SS
                        Arr = Split(strSocSec, "-")
                        If UBound(Arr) > 1 Then
                            txtSS1.Text = Arr(0)
                            txtSS2.Text = Arr(1)
                            txtSS3.Text = Arr(2)
                        End If
                        Exit Sub
                    ElseIf iAnswer = vbCancel Then
                        If clsPatient.SS <> vbNullString Then strSocSec = clsPatient.SS
                        Arr = Split(strSocSec, "-")
                        If UBound(Arr) > 1 Then
                            txtSS1.Text = Arr(0)
                            txtSS2.Text = Arr(1)
                            txtSS3.Text = Arr(2)
                        End If
                        mbNotSave = True
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    
    '---------------------------------------------------------------------------------------
    
    If optMale.Value Then strGender = "M"
    If optFemale.Value Then strGender = "F"

    strRoom = vbNullString
    strStatus = vbNullString
    strRoomPhone = vbNullString
    strPatType = vbNullString
    strRoom = vbNullString
    strBloodType = Literals(txtBloodType.Text, "'")
    frmPatientName.mstrPatLName = Literals(frmPatientName.mstrPatLName, "'")
    frmPatientName.mstrPatFName = Literals(frmPatientName.mstrPatFName, "'")
    If frmPatientAdress.mstrPatAddress = vbNullString Then
        MsgBox "Address is mandatory field!", vbInformation
        mbNotSave = True
        Exit Sub
    ElseIf Trim(strBirth) = vbNullString Then
        MsgBox "Birth is mandatory field!", vbInformation
        mbNotSave = True
        Exit Sub
  '  ElseIf Trim(strPhone) = vbNullString Then
       ' MsgBox "Phone is mandatory field!", vbInformation
      '  mbNotSave = True
      '  Exit Sub
    End If
    frmPatientName.mstrPatTitle = StrConv(frmPatientName.mstrPatTitle, vbProperCase)
    frmPatientName.mstrPatFName = StrConv(frmPatientName.mstrPatFName, vbProperCase)
    frmPatientName.mstrPatMName = StrConv(frmPatientName.mstrPatMName, vbProperCase)
    frmPatientName.mstrPatLName = StrConv(frmPatientName.mstrPatLName, vbProperCase)
    frmPatientName.mstrPatSuffix = StrConv(frmPatientName.mstrPatSuffix, vbProperCase)
    
  '  If Check_ValidPhone = False Then Exit Sub
  'Test Changes
     If Trim(txtEmergContact) <> vbNullString Then
        sEmerContact = Trim(txtEmergContact.Text)
        sEmerRel = Trim(cboRelation.Text)
        sEmrPhone = Trim(txtEmergAC.Text) & "-" & Trim(txtEmergPh1.Text) & "-" & Trim(txtEmergPh2.Text)
        strSQL = "UPDATE PATIENT  SET EMERGENCY_CONTACT = '" & sEmerContact & "',EMERG_PHONE = '" & sEmrPhone & "', " _
             & "EMERG_RELATION = '" & sEmerRel & "' WHERE PATIENT_ID = " & piPatientID
        Set rs = cn.Execute(strSQL)
    End If
    Open_DB psConnect(0)

    Update_Archive strExtRef, frmPatientName.mstrPatTitle, frmPatientName.mstrPatFName, frmPatientName.mstrPatMName, _
                 frmPatientName.mstrPatLName, frmPatientAdress.mstrPatAddress, frmPatientAdress.mstrPatAddress2, _
                 frmPatientAdress.mstrPatCity, frmPatientAdress.mstrPatState, frmPatientAdress.mstrPatZip, _
                 frmPatientAdress.mstrPatZipExt, frmPatientAdress.mstrPatCountry, Literals(strAreaCode, "'"), _
                 Literals(strPhone, "'"), Literals(strCellPhone, "'"), Literals(txtDriverLicense, "'"), _
                 Literals(cboPatientRace.Text, "'"), Literals(strStatus, "'"), _
                 frmPatientName.mstrPatSuffix, Literals(strPatType, "'"), _
                 Literals(strRoom, "'"), strRoomPhone, strBirth, Literals(strSocSec, "'"), _
                 strGender, Literals(txtEmailAddress, "'"), Literals(txtInternationalPhone, "'"), _
                 frmPatientAdress.mstrPatCounty, strBloodType, strMaritalStatus, strWorkPhone, Literals(sEmerContact, "'"), sEmrPhone, sEmerRel
                 
    cn.u_Patient strExtRef, frmPatientName.mstrPatTitle, frmPatientName.mstrPatFName, frmPatientName.mstrPatMName, _
                 frmPatientName.mstrPatLName, frmPatientAdress.mstrPatAddress, frmPatientAdress.mstrPatAddress2, _
                 frmPatientAdress.mstrPatCity, frmPatientAdress.mstrPatState, frmPatientAdress.mstrPatZip, _
                 frmPatientAdress.mstrPatZipExt, frmPatientAdress.mstrPatCountry, Literals(strAreaCode, "'"), _
                 Literals(strPhone, "'"), Literals(strCellPhone, "'"), Literals(txtDriverLicense, "'"), _
                 Literals(cboPatientRace.Text, "'"), Literals(strStatus, "'"), _
                 frmPatientName.mstrPatSuffix, Literals(strPatType, "'"), _
                 Literals(strRoom, "'"), strRoomPhone, strBirth, Literals(strSocSec, "'"), _
                 strGender, Literals(txtEmailAddress, "'"), Literals(txtInternationalPhone, "'"), _
                 frmPatientAdress.mstrPatCounty, strBloodType, strMaritalStatus, strWorkPhone, _
                 piPatientID, rs
            
    mbNotSave = False
    If Trim(cboLanguage.Text) <> vbNullString Then
        clsPatient.Language = Trim(cboLanguage.Text)
        strSQL = "UPDATE PATIENT SET LANGUAGE = '" & clsPatient.Language & "' WHERE PATIENT_ID = " & piPatientID
        Set rs = cn.Execute(strSQL)
    End If
    Screen.MousePointer = vbDefault
    Close_DB psConnect(0)
    Write_PatientProfiler
 
    Exit Sub
Error_UpdatePatientInfo:
        Screen.MousePointer = vbDefault
        Close_DB psConnect(0)
        MsgBox Err.Description, vbInformation
        Exit Sub
        Resume
End Sub


Private Sub Form_Load()
Dim strBirth As String
Dim strSocialSec As String
Dim varSocSec
Dim varCellPhone
Dim sStr As String
Dim strSQL As String
Dim strState As String
Dim i%

    mFirst = True
    bEnable = True
    Clear_Fields
    Clear_Emerg
    Clear_PatientStrings
    mbNotSave = False
    Get_RelationshipTerms cboRelation
    Get_RelationshipTerms cboPartRel
    LoadAttrValidValue psConnect(1), cboPatientRace, "sd_Patient_Race"
    LoadAttrValidValue psConnect(1), cboPatMarital, "sd_Patient_MaritalStat"
    LoadAttrValidValue psConnect(1), cboParTitle, "sd_Patient_Title"
    If mNewPatient = False Then
        Set_Patient
    Else
        lblPhoto.Visible = True
        cmdNew.Enabled = False
        cmdSave.Enabled = True
        cmdSave.ZOrder 0
        cmdSave.Visible = True
        cmdUpdate.Visible = False
    End If
    cboFacility.Clear
    Load_Facilities
    
    If piFacilityID = 0 Then
        piFacilityID = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultFacility")
    End If
    Screen.MousePointer = vbDefault
    If mNewPatient Then
        fraPartner.Visible = True
    Else
        fraPartner.Visible = True
    End If
    GetTerms cboLanguage, "sd_Patient_Language"
    Exit Sub
Error_Handler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbInformation
    Exit Sub
    Resume
End Sub

Private Sub Screen_Action()
   
   
    txtEmailAddress.Enabled = bEnable
    txtAreaCode.Enabled = bEnable
    txtPhone1.Enabled = bEnable
    txtPhone2.Enabled = bEnable
    txtCellAreaCode.Enabled = bEnable
    txtCellPhone1.Enabled = bEnable
    txtCellPhone2.Enabled = bEnable
    txtInternationalPhone.Enabled = bEnable
    txtDriverLicense.Enabled = bEnable
    txtMonth.Enabled = bEnable
    txtYear.Enabled = bEnable
    txtDay.Enabled = bEnable
    txtAge.Enabled = bEnable
    txtSS1.Enabled = bEnable
    txtSS2.Enabled = bEnable
    txtSS3.Enabled = bEnable
    optMale.Enabled = bEnable
    optFemale.Enabled = bEnable
    cboPatientRace.Enabled = bEnable
    txtParner.Enabled = bEnable
    txtEmergContact.Enabled = bEnable
    cboRelation.Enabled = bEnable
    txtEmergAC.Enabled = bEnable
    txtEmergPh1.Enabled = bEnable
    txtEmergPh2.Enabled = bEnable
    txtBloodType.Enabled = bEnable
    lblPatient_id.Enabled = bEnable
    txtExtRef.Enabled = bEnable
    If bEnable Then
    ' font bold
       
       
        txtEmailAddress.FontBold = True
        txtAreaCode.FontBold = True
        txtPhone1.FontBold = True
        txtPhone2.FontBold = True
        txtCellAreaCode.FontBold = True
        txtCellPhone1.FontBold = True
        txtCellPhone2.FontBold = True
        txtInternationalPhone.FontBold = True
        txtDriverLicense.FontBold = True
        txtMonth.FontBold = True
        txtYear.FontBold = True
        txtDay.FontBold = True
        txtAge.FontBold = True
        txtSS1.FontBold = True
        txtSS2.FontBold = True
        txtSS3.FontBold = True
        optMale.FontBold = True
        optFemale.FontBold = True
        cboPatientRace.FontBold = True
        txtParner.FontBold = True
        txtEmergContact.FontBold = True
        cboRelation.FontBold = True
        txtEmergAC.FontBold = True
        txtEmergPh1.FontBold = True
        txtEmergPh2.FontBold = True
        txtBloodType.FontBold = True
        lblPatient_id.FontBold = True
        txtExtRef.FontBold = True
    End If
    
End Sub

Private Sub Form_Paint()
    ColorBurst Me, "BLUE", 128, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object

    On Error Resume Next
    Clear_PatientStrings
    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.objects
      Set obj = Nothing
    Next obj
End Sub



Private Sub Set_Patient()
Dim strBirth As String
Dim strSocialSec As String
Dim varSocSec
Dim strCellPhone
Dim sStr As String
Dim strSQL As String
Dim strState As String
Dim i%
Dim strBloodType  As String
Dim strMaritalStatus As String
Dim strWorkPhone As String
Dim varWorkPhone
Dim strMarried As String
Dim sRace As String
Dim Arr

    Clear_Fields
    Get_Partner
    
    If cboPatientRace.ListCount <= 0 Then LoadAttrValidValue psConnect(1), cboPatientRace, "sd_Patient_Race"
    
    Open_DB psConnect(0)
    
    If modSpecific.piPatientID = 0 Then modSpecific.piPatientID = piPatientID
    piPatientID = modSpecific.piPatientID
    cn.s_PatientInfo piPatientID, rs
        
    With rs
        If Not .EOF Then
            If Not IsNull(!Title) And Trim((!Title)) <> vbNullString Then frmPatientName.mstrPatTitle = Trim((!Title))
            If Not IsNull(!Patient_Ext_Ref) And Trim((!Patient_Ext_Ref)) <> vbNullString Then txtExtRef.Text = Trim((!Patient_Ext_Ref)) & ""
            If Not IsNull(!First_Name) Then
                If TextFound("''", (!First_Name), 2, True) Then
                    frmPatientName.mstrPatFName = Replace((!First_Name), "''", "'")
                Else
                    frmPatientName.mstrPatFName = (!First_Name)
                End If
            End If
            If Not IsNull(!Middle_Name) Then frmPatientName.mstrPatMName = (!Middle_Name)
            If Not IsNull(!Last_Name) Then
                If TextFound("''", (!Last_Name), 2, True) Then
                    frmPatientName.mstrPatLName = Replace((!Last_Name), "''", "'")
                Else
                    frmPatientName.mstrPatLName = Trim((!Last_Name))
                End If
            End If
            If Not IsNull(!Address1) And Trim((!Address1)) <> vbNullString Then frmPatientAdress.mstrPatAddress = Trim((!Address1))
            If Not IsNull(!Address2) And Trim((!Address2)) <> vbNullString Then frmPatientAdress.mstrPatAddress2 = Trim((!Address2))
            If Not IsNull(!City) And Trim((!City)) <> vbNullString Then frmPatientAdress.mstrPatCity = Trim((!City))
            If Not IsNull(!State) And Trim((!State)) <> vbNullString Then frmPatientAdress.mstrPatState = Trim((!State))
            If Not IsNull(!ZipCode) And Trim((!ZipCode)) <> vbNullString Then frmPatientAdress.mstrPatZip = Trim((!ZipCode))
            If Not IsNull(!ZipCode_Ext) And Trim((!ZipCode_Ext)) <> vbNullString Then frmPatientAdress.mstrPatZipExt = Trim((!ZipCode_Ext))
            If Not IsNull(!Country) And Trim((!Country)) <> vbNullString Then frmPatientAdress.mstrPatCountry = Trim((!Country))
            If Not IsNull(!County) And Trim((!County)) <> vbNullString Then frmPatientAdress.mstrPatCounty = Trim((!County))
            If Not IsNull(!Race) And Trim((!Race)) <> vbNullString Then
                cboPatientRace.AddItem "white"
                For i = 0 To cboPatientRace.ListCount - 1
                    If TextFound((!Race), cboPatientRace.List(i), Len((!Race)), True) Then
                        cboPatientRace.Text = cboPatientRace.List(i)
                        Exit For
                    End If
                Next
            End If
            If Not IsNull(!Suffix) Then frmPatientName.mstrPatSuffix = (!Suffix)
            If Not IsNull(!Blood_Type) Then txtBloodType.Text = (!Blood_Type) & ""
            If Trim(!Phone) <> vbNullString Then
                If (!Area_Code) <> vbNullString Then txtAreaCode.Text = Format((!Area_Code), "000")
                Arr = Split((!Phone), "-")
                If UBound(Arr) > 0 Then
                   txtPhone1.Text = Format(Arr(0), "000")
                   txtPhone2.Text = Format(Arr(1), "0000")
                End If
            Else
                txtAreaCode.Text = vbNullString
                txtPhone1.Text = vbNullString
                txtPhone2.Text = vbNullString
            End If
            Arr = vbNullString
            If Trim((!WORK_PHONE)) <> vbNullString Then
                varWorkPhone = Split((!WORK_PHONE), "-")
                If UBound(varWorkPhone) > 0 Then
                    If Trim(varWorkPhone(0)) <> vbNullString Then txtWorkAC.Text = Format(varWorkPhone(0), "000")
                    If Trim(varWorkPhone(1)) <> vbNullString Then txtWorkPh1.Text = Format(varWorkPhone(1), "000")
                    If Trim(varWorkPhone(2)) <> vbNullString Then txtWorkPh2.Text = Format(varWorkPhone(2), "0000")
                Else
                    txtWorkAC.Text = vbNullString
                    txtWorkPh1.Text = vbNullString
                    txtWorkPh2.Text = vbNullString
                End If
            End If
            If Trim(!Patient_Ext_Ref) <> vbNullString Then pxPatientID = (!Patient_Ext_Ref) & ""
            If (!Cell_Phone) <> vbNullString Then
                strCellPhone = Split((!Cell_Phone), "-")
                    If UBound(strCellPhone) > 0 Then
                            txtCellAreaCode.Text = Format(strCellPhone(0), "000")
                            txtCellPhone1.Text = Format(strCellPhone(1), "000")
                            txtCellPhone2.Text = Format(strCellPhone(2), "0000")
                    End If

            Else
                txtCellAreaCode.Text = vbNullString
                txtCellPhone1.Text = vbNullString
                txtCellPhone2.Text = vbNullString
            End If
            If Trim(!Drv_License) <> vbNullString Then txtDriverLicense.Text = Trim(!Drv_License) & ""
            strBirth = (!Birth) & ""
                If IsDate(strBirth) Then
                    strBirth = Format(strBirth, "mm/dd/yyyy")
                    txtMonth.Text = Left(strBirth, 2)
                    txtDay.Text = Mid(strBirth, 4, 2)
                    txtYear.Text = Mid(strBirth, 7, 4)
                    If clsPatient.DOB = vbNullString Then Get_Patient piPatientID
                    If Trim(clsPatient.DOB) <> vbNullString Then clsPatient.Age = Int(DateDiff("y", clsPatient.DOB, Now) / 365.25)
                    If Not IsNull(clsPatient.Age) Then txtAge.Text = Trim(clsPatient.Age)
                    If Val(txtAge.Text) <= 18 And mstrPractice = "Psychiatry" Then
                        lblPartnerName.Caption = "Parent or Legal Guardian"
                        lblPartTitle.Caption = "PARENT OR LEGAL GUARDIAN"
                    End If
                End If
                If Trim((!Social_Security)) <> vbNullString Then
                   Arr = Split((!Social_Security), "-")
                   If UBound(Arr) > 0 Then
                      If Trim(Arr(0)) <> vbNullString Then txtSS1.Text = Format(Arr(0), "000")
                      If Trim(Arr(1)) <> vbNullString Then txtSS2.Text = Format(Arr(1), "00")
                      If Trim(Arr(2)) <> vbNullString Then txtSS3.Text = Format(Arr(2), "0000")
                   End If
                End If
        
                If Not IsNull((!Gender)) Then
                   If (!Gender) = "M" Then
                      optMale.Value = True
                   ElseIf (!Gender) = "F" Then
                       optFemale.Value = True
                   End If
                End If
                If Trim(!Email) <> vbNullString Then txtEmailAddress.Text = (!Email) & ""
                If Trim(!internatl_phone) <> vbNullString Then txtInternationalPhone.Text = (!internatl_phone) & ""
    
                If Not IsNull(!Marital_Status) Then strMarried = Trim(!Marital_Status)
                If Trim(strMarried) <> vbNullString Then cboPatMarital.ListIndex = SetListIndex(cboPatMarital, strMarried & "")
                If (!EMERGENCY_CONTACT) <> vbNullString Then txtEmergContact.Text = Trim(!EMERGENCY_CONTACT)
                If (!EMERG_PHONE) <> vbNullString Then
                    Arr = Split((!EMERG_PHONE), "-")
                    If Arr(0) <> vbNullString Then txtEmergAC.Text = Trim(Arr(0))
                    If Arr(1) <> vbNullString Then txtEmergPh1.Text = Arr(1)
                    If Arr(2) <> vbNullString Then txtEmergPh2 = Arr(2)
                End If
                If (!EMERG_RELATION) <> vbNullString Then
                    If cboRelation.ListCount <= 0 Then Get_RelationshipTerms cboRelation
                    For i = 0 To cboRelation.ListCount - 1
                        If cboRelation.List(i) = Trim((!EMERG_RELATION)) Then
                            cboRelation.Text = cboRelation.List(i)
                            Exit For
                        End If
                    Next i
                End If
        End If
    End With
        
        Close_DB psConnect(0)
            
        If Trim(frmPatientName.mstrPatTitle) <> vbNullString Then txtFullName.Text = Trim(frmPatientName.mstrPatTitle)
        If Trim(frmPatientName.mstrPatFName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(frmPatientName.mstrPatFName)
        If Trim(frmPatientName.mstrPatMName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & frmPatientName.mstrPatMName
        If Trim(frmPatientName.mstrPatLName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(frmPatientName.mstrPatLName)
        If Trim(frmPatientName.mstrPatSuffix) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & ", " & Trim(frmPatientName.mstrPatSuffix)
        If Trim(frmPatientAdress.mstrPatAddress) <> vbNullString Then txtAddress.Text = Trim(frmPatientAdress.mstrPatAddress)
        If Trim(frmPatientAdress.mstrPatAddress2) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & ", " & Trim(frmPatientAdress.mstrPatAddress2)
        If Trim(frmPatientAdress.mstrPatCity) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & vbNewLine & Trim(frmPatientAdress.mstrPatCity)
        If Trim(frmPatientAdress.mstrPatState) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & " " & Trim(frmPatientAdress.mstrPatState)
        If Trim(frmPatientAdress.mstrPatZip) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & " " & Trim(frmPatientAdress.mstrPatZip)
        If Trim(frmPatientAdress.mstrPatCountry) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & vbNewLine & Trim(frmPatientAdress.mstrPatCountry)
        If Trim(frmPatientAdress.mstrPatCounty) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & vbNewLine & Trim(frmPatientAdress.mstrPatCounty)
    
        'Get_Patient
        If piPatientID <> 0 Then lblPatient_id.Caption = piPatientID
        mboobypass = True
        
        clsPatient.Language = vbNullString
        Get_Language piPatientID
        If clsPatient.Language <> vbNullString Then
            For i = 0 To cboLanguage.ListCount - 1
                If cboLanguage.List(i) = clsPatient.Language Then cboLanguage.Text = cboLanguage.List(i)
            Next i
        Else
             For i = 0 To cboLanguage.ListCount - 1
                If cboLanguage.List(i) = "English" Then cboLanguage.Text = cboLanguage.List(i)
            Next i
        End If
        If piPatientID > 0 Then
            fraPic.ZOrder 0
            If Right(psImageFilePath, 1) <> "\" Then
                psImageFilePath = psImageFilePath & "\"
            End If
            If FileExists(psImageFilePath & piPatientID & "\Photo.jpg") Then
               psPhoto = psImageFilePath & piPatientID & "\Photo.jpg"
               imgPatient.Picture = LoadPicture(psPhoto)
               lblPhoto.Visible = False
               imgPatient.Visible = True
            Else
              imgPatient.Visible = False
              lblPhoto.Visible = True
            End If
        End If
'
'        psPhoto = psImageFilePath & piPatientID & "\Photo.jpg"
'
'        If FileExists(psPhoto) Then
'            imgPatient.Picture = LoadPicture(psPhoto)
'            lblPhoto.Visible = False
'        Else
'            imgPatient.Picture = LoadPicture()
'            lblPhoto.Visible = True
'        End If
        
        cmdNew.Visible = False
        cmdUpdate.Visible = True
        cmdClose.Visible = False
        optMale.Visible = True
        optFemale.Visible = True
        bEnable = False
        Screen_Action
    
        Screen.MousePointer = vbDefault
    Exit Sub

End Sub



Private Sub txtAddress_Click()
Dim strAddress As String

    If mboobypass Then mboobypass = False: Exit Sub
    txtAddress.Text = vbNullString
    frmPatientAdress.Show vbModal
    If frmPatientAdress.mstrPatAddress <> vbNullString Then
        
        If Trim(frmPatientAdress.mstrPatAddress) <> vbNullString Then strAddress = Trim(frmPatientAdress.mstrPatAddress)
        If Trim(frmPatientAdress.mstrPatAddress2) <> vbNullString Then strAddress = Trim(strAddress) & ", " & Trim(frmPatientAdress.mstrPatAddress2)
        If Trim(frmPatientAdress.mstrPatCity) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(frmPatientAdress.mstrPatCity)
        If Trim(frmPatientAdress.mstrPatState) <> vbNullString Then strAddress = Trim(strAddress) & ", " & Trim(frmPatientAdress.mstrPatState)
        If Trim(frmPatientAdress.mstrPatZip) <> vbNullString Then strAddress = Trim(strAddress) & " " & Trim(frmPatientAdress.mstrPatZip)
        If Trim(frmPatientAdress.mstrPatCountry) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(frmPatientAdress.mstrPatCountry)
        If Trim(frmPatientAdress.mstrPatCounty) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(frmPatientAdress.mstrPatCounty)
        txtAddress.Text = Trim(strAddress)
    End If
    
    Exit Sub
End Sub


Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strAddress As String

    txtAddress.Text = vbNullString
    frmPatientAdress.Show vbModal
        If Trim(clsPatient.Address) <> vbNullString Then strAddress = Trim(clsPatient.Address)
        If Trim(clsPatient.Address2) <> vbNullString Then strAddress = Trim(strAddress) & ", " & Trim(clsPatient.Address2)
        If Trim(clsPatient.City) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(clsPatient.City)
        If Trim(clsPatient.State) <> vbNullString Then strAddress = Trim(strAddress) & ", " & Trim(clsPatient.State)
        If Trim(clsPatient.Zip) <> vbNullString Then strAddress = Trim(strAddress) & " " & Trim(clsPatient.Zip)
        If Trim(clsPatient.Country) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(clsPatient.Country)
        If Trim(clsPatient.County) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(clsPatient.County)
        txtAddress.Text = Trim(strAddress)
    Exit Sub
    
End Sub


Private Sub txtAddress_KeyUp(KeyCode As Integer, Shift As Integer)

    If Len(Trim(txtAreaCode.Text)) < 3 Then txtAreaCode.Text = vbNullString
    
End Sub


Private Sub txtAreaCode_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(txtAreaCode.Text) = (txtAreaCode.MaxLength) Then
         txtPhone1.Enabled = True
         txtPhone1.SetFocus
      End If
End Sub


Private Sub txtCellAreaCode_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(txtCellAreaCode.Text) = (txtCellAreaCode.MaxLength) Then
            txtCellPhone1.Enabled = True
            txtCellPhone1.SetFocus
      End If
End Sub


Private Sub txtCellPhone1_KeyUp(KeyCode As Integer, Shift As Integer)

    If Len(txtCellPhone1.Text) >= (txtCellPhone1.MaxLength) Then
         txtCellPhone2.Enabled = True
         txtCellPhone2.SetFocus
     End If
End Sub


Private Sub txtCellPhone2_Change()
mbooHasChanged = True
End Sub

Private Sub txtCellPhone2_LostFocus()
    If Len(txtCellPhone2.Text) > 1 And Len(txtCellPhone2.Text) < 4 Then
        Call MsgBox("Invalid data entry ! Telephone fields require numeric data. Erroneous data will not be removed. Please re-enter valid data.  ", vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1, "Patient Demographic Maintanence")
        txtCellPhone2.Text = vbNullString
    End If
    Exit Sub
End Sub


Private Sub txtDay_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lstrBirthDate As String
   
   lstrBirthDate = txtMonth.Text & "/" & txtDay.Text & "/" & txtYear.Text

   If Len(txtDay.Text) = 2 And IsDate(lstrBirthDate) Then
      lstrBirthDate = txtMonth.Text & "/" & txtDay.Text & "/" & txtYear.Text
      txtAge.Text = DateDiff("yyyy", lstrBirthDate, Now)
    
      If DateDiff("d", txtMonth.Text & "/" & txtDay.Text & "/" & Year(Now), Now) < 0 Then txtAge.Text = DateDiff("yyyy", lstrBirthDate, Now) - 1
   End If
 
   If KeyCode >= 32 And KeyCode <= 127 Then
      If Len(txtDay.Text) = txtDay.MaxLength Then
         If txtYear.Enabled Then txtYear.SetFocus
      End If
   End If
End Sub


Private Sub txtEmergAC_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(txtEmergAC.Text) = (txtEmergAC.MaxLength) Then
         txtEmergPh1.Enabled = True
         txtEmergPh1.SetFocus
    End If
      Exit Sub
End Sub


Private Sub txtEmergContact_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cnEmerg As New ADODB.Connection
Dim rsEmerg As New ADODB.Recordset
Dim strSQL As String
Dim Arr
Dim Idx As Integer


    If piPatientID = 0 Then
        MsgBox "Emergency Contact can only be updated on existing patients! Please save this Patient First. ", vbInformation
    Else
        If piProviderID = 0 Then piProviderID = modSpecific.piProviderID
        mstrPatOcx = "FN_PP_PS"
        mstrPatCtl = "EmergencyContact"
        frmLoadControl.Show vbModal
        If cnEmerg.State = adStateClosed Then cnEmerg.open psConnect(0)
        strSQL = "SELECT EMERGENCY_CONTACT, EMERG_PHONE, EMERG_RELATION FROM PATIENT WHERE PATIENT_ID = " & piPatientID
        Set rsEmerg = cnEmerg.Execute(strSQL)
        With rsEmerg
            If Not .EOF Then
                If (!EMERGENCY_CONTACT) <> vbNullString Then txtEmergContact.Text = Trim((!EMERGENCY_CONTACT))
                If (!EMERG_PHONE) <> vbNullString Then
                    Arr = Split((!EMERG_PHONE), "-")
                    If UBound(Arr) > 0 Then
                        For Idx = LBound(Arr) To UBound(Arr)
                            If UBound(Arr) > 1 Then
                                If Idx = 0 Then txtEmergAC.Text = Arr(Idx)
                                If Idx = 1 Then txtEmergPh1.Text = Arr(Idx)
                                If Idx = 2 Then txtEmergPh2.Text = Arr(Idx)
                            End If
                        Next Idx
                    End If
                End If
                If (!EMERG_RELATION) <> vbNullString Then
                    If cboRelation.ListCount <= 0 Then Get_RelationshipTerms cboRelation
                    For Idx = 0 To cboRelation.ListCount - 1
                        If cboRelation.List(Idx) = Trim(!EMERG_RELATION) Then
                            cboRelation.Text = cboRelation.List(Idx)
                            Exit For
                        End If
                    Next Idx
                End If
            End If
        End With
        'End If
    End If
    If cnEmerg.State = adStateOpen Then
        cnEmerg.Close
        Set cnEmerg = Nothing
        Set rsEmerg = Nothing
    End If
    Exit Sub
End Sub


Private Sub txtEmergPh1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(txtEmergPh1.Text) = (txtEmergPh1.MaxLength) Then
         txtEmergPh2.Enabled = True
         txtEmergPh2.SetFocus
    End If
End Sub


Private Sub txtFullName_Click()
Dim strName As String

    If mboobypass Then mboobypass = False: Exit Sub
    If mNewPatient Then
        frmPatientName.mstrPatTitle = vbNullString
        frmPatientName.mstrPatFName = vbNullString
        frmPatientName.mstrPatMName = vbNullString
        frmPatientName.mstrPatLName = vbNullString
        frmPatientName.mstrPatSuffix = vbNullString
    End If
         txtFullName.Text = Empty
        frmPatientName.Show vbModal
   If frmPatientName.mstrPatFName <> Empty Then
        If Trim(frmPatientName.mstrPatTitle) <> vbNullString Then txtFullName.Text = Trim(frmPatientName.mstrPatTitle)
        If Trim(frmPatientName.mstrPatFName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(frmPatientName.mstrPatFName)
        If Trim(frmPatientName.mstrPatMName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(frmPatientName.mstrPatMName)
        If Trim(frmPatientName.mstrPatLName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(frmPatientName.mstrPatLName)
        If Trim(frmPatientName.mstrPatSuffix) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & ", " & Trim(frmPatientName.mstrPatSuffix)
    End If
    
    
    'Strip unwanted commas
    If Right(txtFullName.Text, 1) = ", " And Trim(frmPatientName.mstrPatSuffix) = vbNullString Then txtFullName.Text = Replace(txtFullName.Text, ", ", "")
    txtFullName.Text = Trim((txtFullName.Text))
    txtFullName.Text = StrConv(txtFullName.Text, vbProperCase)
    
    Exit Sub
    
End Sub

Private Sub txtFullName_KeyDown(KeyCode As Integer, Shift As Integer)
         txtFullName.Text = vbNullString
        frmPatientName.Show vbModal
        If Trim(clsPatient.Title) <> vbNullString Then txtFullName.Text = Trim(clsPatient.Title)
        If Trim(clsPatient.FirstName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.FirstName)
        If Trim(clsPatient.MiddleName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.MiddleName)
        If Trim(clsPatient.LastName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.LastName)
        If Trim(clsPatient.Suffix) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & ", " & Trim(clsPatient.Suffix)
    Exit Sub
End Sub


Private Sub txtMonth_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(txtMonth.Text) = (txtMonth.MaxLength) Then
            txtDay.Enabled = True
            txtDay.SetFocus
      End If
      Exit Sub
End Sub


Private Sub txtParner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngPatID As Long
    lngPatID = Save_Data
    If lngPatID > 0 Then
        piPatientID = lngPatID
        frmLoadControl.Show vbModal
        Get_Partner lngPatID
        txtParner.Text = clsPartner.Name
    End If
    Exit Sub
End Sub


Private Sub txtPhone1_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(txtPhone1.Text) = (txtPhone1.MaxLength) Then
            txtPhone2.Enabled = True
            txtPhone2.SetFocus
     End If
      Exit Sub
End Sub


Private Sub txtPhone2_KeyUp(KeyCode As Integer, Shift As Integer)

    Exit Sub
End Sub


Private Sub txtPhone2_LostFocus()
   If Len(txtPhone2.Text) > 1 And Len(txtPhone2.Text) < 4 Then
        Call MsgBox("Invalid data entry ! Telephone fields require numeric data. Erroneous data will not be removed. Please re-enter valid data.  ", vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1, "Patient Demographic Maintanence")
        txtPhone2.Text = vbNullString
    End If
End Sub


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
    If Len(txtSS3.Text) = 4 Then
        txtSS3.Text = Format(txtSS3.Text, "0000")
    End If
    Exit Sub
End Sub


Private Sub txtSS3_LostFocus()
    txtSS3.Text = Format(txtSS3.Text, "0000")
    Exit Sub
End Sub


Private Sub txtWorkAC_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(txtWorkAC.Text) = (txtWorkAC.MaxLength) Then
         If txtWorkPh1.Enabled Then txtWorkPh1.SetFocus
      End If
End Sub


Private Sub txtWorkPh1_Change()
    mbooHasChanged = True
End Sub

Private Sub txtWorkPh1_GotFocus()
    mbooHasChanged = True
End Sub

Private Sub txtWorkPh1_KeyUp(KeyCode As Integer, Shift As Integer)
      If Len(txtWorkPh1.Text) = (txtWorkPh1.MaxLength) Then
         txtWorkPh2.Enabled = True
         txtWorkPh2.SetFocus
      End If
End Sub


Private Sub txtWorkPh2_LostFocus()
   If Len(txtWorkPh2.Text) > 1 And Len(txtWorkPh2.Text) < 4 Then
        Call MsgBox("Invalid data entry ! Telephone fields require numeric data. Erroneous data will not be removed. Please re-enter valid data.  ", vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1, "Patient Demographic Maintanence")
        txtWorkPh2.Text = vbNullString
    End If
    Exit Sub
End Sub


Private Sub txtYear_Change()
   Dim lstrBirthDate As String
   
   On Error GoTo ERROR
   lstrBirthDate = txtMonth.Text & "/" & txtDay.Text & "/" & txtYear.Text
    If Len(txtYear.Text) = 4 Then
        Age lstrBirthDate, Now
    End If
    
    If Len(txtYear.Text) = 4 Then
        If Not IsNull(clsPatient.Age) Then txtAge.Text = Trim(clsPatient.Age)
        If Val(txtAge.Text) <= 18 And mstrPractice = "Psychiatry" Then
            MsgBox "Please Complete Patient's Parent or Legal Guardian Information!", vbInformation
            lblPartnerName.Caption = "Parent/Legal Guardian"
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR:
    MsgBox Err.Description & " " & Err.Number, vbInformation
    Exit Sub
    Resume
End Sub



Private Sub Write_PatientProfiler()
Dim lstrTitle As String
Dim lstrFirstName As String
Dim lstrMiddleName As String
Dim lstrLastName As String
Dim lstrAction As String
Dim lintEmailRecpt As Integer
Dim strSQL As String
Dim llngReturn As Long
Dim strRS
Dim lReturn As Integer
Dim lconClinical As New ADODB.Connection
Dim strTitle As String
Dim strFirstName As String
Dim strLastName As String
Dim strMidName As String
Dim strParSum As String
Dim strEmerContact As String
Dim strEmerRelation As String
Dim strEmerPhoneAC As String
Dim strEmerPhone1 As String
Dim strEmerPhone2 As String
Dim strPartName As String
Dim lngRelKey As String
Dim i%
Dim cnPart As New ADODB.Connection
Dim rsPart As New ADODB.Recordset
Dim strSuffix As String
    
    
    On Error GoTo Error_Handler
    
    lconClinical.open psConnect(0)
    mbooHasChanged = True
    lstrFirstName = Literals(txtParFirstName.Text, "'")
    lstrMiddleName = Literals(txtParMidName.Text, "'")
    lstrLastName = Literals(txtParLastName.Text, "'")
    If cboParTitle.Text <> vbNullString Then
        strTitle = cboParTitle.Text
    End If
    strParSum = strTitle & txtParFirstName.Text & " " _
        & txtParMidName.Text & " " & txtParLastName.Text
    If Trim(strParSum) = vbNullString Then Exit Sub
   lstrAction = "I"
   If mlngRelationPatID <> 0 Then lstrAction = "U"
   If lstrAction = "U" Then
        strSQL = "exec MAINT_PATIENT_NAME '" & lstrAction & "','" & _
                    mlngRelationPatID & "','" & strTitle & "','" & _
                    lstrFirstName & "','" & lstrMiddleName & "','" & _
                    lstrLastName & "'"
        llngReturn = ActionQuery(strSQL, psConnect(0))
        If cboSuffix.Text <> vbNullString Then
            strSuffix = Trim(cboSuffix.Text)
            If cnPart.State = adStateClosed Then cnPart.open psConnect(0)
            strSQL = "update PATIENT set SUFFIX = '" & strSuffix & "' where PATIENT_ID = " & mlngRelationPatID
            Set rsPart = cnPart.Execute(strSQL)
            If cnPart.State = adStateClosed Then
                cnPart.Close
                Set cnPart = Nothing
                Set rsPart = Nothing
            End If
        End If
   Else
        strSQL = "SET NOCOUNT ON;" & _
            "exec MAINT_PATIENT_NAME '" & lstrAction & "','','" & _
            strTitle & "','" & _
            lstrFirstName & "','" & lstrMiddleName & "','" & _
            lstrLastName & "';" & _
            "Select @@IDENTITY;"
        
        lReturn = GetRecordSet(strRS, strSQL, psConnect(0))
        If lReturn < 0 Then
            Exit Sub
        Else
            mlngRelationPatID = CLng(strRS(0, 0))
        End If

        For i = 0 To cboPartRel.ListCount - 1
            If cboPartRel.List(i) = cboPartRel.Text Then
                lngRelKey = cboPartRel.ItemData(i)
                Exit For
            End If
        Next i
        strSQL = "exec MAINT_PATIENT_RELATIONS '" & lstrAction & "','" & piPatientID & "','" & lngRelKey & "','" & mlngRelationPatID & "'"
        llngReturn = ActionQuery(strSQL, psConnect(0))
        If cboSuffix.Text <> vbNullString Then
            strSuffix = Trim(cboSuffix.Text)
            If cnPart.State = adStateClosed Then cnPart.open psConnect(0)
            strSQL = "update PATIENT set SUFFIX = '" & strSuffix & "' where PATIENT_ID = " & mlngRelationPatID
            Set rsPart = cnPart.Execute(strSQL)
            If cnPart.State = adStateClosed Then
                cnPart.Close
                Set cnPart = Nothing
                Set rsPart = Nothing
            End If
        End If
   End If
   Screen.MousePointer = vbDefault
Exit Sub
Error_Handler:
    MsgBox Err.Description, vbInformation
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
End Sub
