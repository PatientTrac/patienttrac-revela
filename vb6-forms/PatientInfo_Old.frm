VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form PatientInfo_Old 
   BackColor       =   &H00C00000&
   Caption         =   "PATIENT CONTACT INFORMATION"
   ClientHeight    =   6945
   ClientLeft      =   5565
   ClientTop       =   2790
   ClientWidth     =   12405
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PatientInfo_Old.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "PatientInfo_Old.frx":08CA
   ScaleHeight     =   6945
   ScaleWidth      =   12405
   Begin VB.Frame fraPatient 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6945
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   12375
      Begin VB.ComboBox cboBloodType 
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "PatientInfo_Old.frx":0DBE
         Left            =   1980
         List            =   "PatientInfo_Old.frx":0DDA
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   4500
         Width           =   1305
      End
      Begin VB.ComboBox cboStatus 
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "PatientInfo_Old.frx":0E06
         Left            =   6570
         List            =   "PatientInfo_Old.frx":0E13
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3660
         Width           =   1995
      End
      Begin VB.Frame fraDOD 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Date of Death"
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   9150
         TabIndex        =   116
         Top             =   3630
         Visible         =   0   'False
         Width           =   2895
         Begin MSComCtl2.DTPicker ccdDateDeath 
            CausesValidation=   0   'False
            Height          =   375
            Left            =   1080
            TabIndex        =   53
            Top             =   0
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   83230723
            CurrentDate     =   39523
            MinDate         =   36526
         End
         Begin VB.Label lblDateDeath 
            BackColor       =   &H00BFE8FD&
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Death"
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
            TabIndex        =   117
            Top             =   0
            Width           =   825
         End
      End
      Begin VB.ComboBox cboPrefContact 
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "PatientInfo_Old.frx":0E33
         Left            =   6540
         List            =   "PatientInfo_Old.frx":0E40
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   450
         Width           =   1965
      End
      Begin VB.ComboBox cboLanguage 
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "PatientInfo_Old.frx":0E68
         Left            =   4275
         List            =   "PatientInfo_Old.frx":0E6A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Select Patient Primary Language"
         Top             =   4890
         Width           =   1995
      End
      Begin VB.Frame fraPartName 
         BackColor       =   &H00E0E0E0&
         Height          =   2925
         Left            =   13080
         TabIndex        =   35
         Top             =   3390
         Visible         =   0   'False
         Width           =   5595
         Begin VB.TextBox txtParFirstName 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            TabIndex        =   44
            Top             =   990
            Width           =   2355
         End
         Begin VB.TextBox txtParMidName 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            TabIndex        =   43
            Top             =   1350
            Width           =   2355
         End
         Begin VB.TextBox txtParLastName 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            TabIndex        =   42
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
            TabIndex        =   39
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
               Picture         =   "PatientInfo_Old.frx":0E6C
               ScaleHeight     =   255
               ScaleWidth      =   375
               TabIndex        =   40
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
               TabIndex        =   41
               Top             =   90
               Width           =   2205
            End
         End
         Begin VB.ComboBox cboParTitle 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   600
            Width           =   2355
         End
         Begin VB.ComboBox cboPartRel 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   2490
            Width           =   2985
         End
         Begin VB.ComboBox cboSuffix 
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   2070
            Width           =   1485
         End
         Begin PT_XP_Button.PT_XP cmdParOk 
            Height          =   375
            Left            =   3990
            TabIndex        =   45
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
            Picture         =   "PatientInfo_Old.frx":13F6
            AppearanceThemes=   3
            TransparentColor=   12582912
            Caption         =   "OK"
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "PatientInfo_Old.frx":1790
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdPartCancel 
            Height          =   375
            Left            =   3990
            TabIndex        =   46
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
            Picture         =   "PatientInfo_Old.frx":1AE2
            AppearanceThemes=   3
            TransparentColor=   12582912
            Caption         =   "Cancel"
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "PatientInfo_Old.frx":24F4
            ShowFocus       =   -1  'True
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title:"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   150
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   1800
            Width           =   420
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Relationship"
            ForeColor       =   &H00800000&
            Height          =   510
            Left            =   150
            TabIndex        =   48
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
            TabIndex        =   47
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
         Height          =   2235
         Left            =   7080
         TabIndex        =   61
         Top             =   3900
         Visible         =   0   'False
         Width           =   4515
         Begin VB.TextBox txtParner 
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   1470
            Width           =   3015
         End
         Begin VB.ComboBox cboRelation 
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1410
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   660
            Width           =   3075
         End
         Begin VB.TextBox txtEmergPh2 
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   3045
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   101
            Top             =   1050
            Width           =   795
         End
         Begin VB.TextBox txtEmergPh1 
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2220
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   100
            Top             =   1050
            Width           =   615
         End
         Begin VB.TextBox txtEmergAC 
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1470
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   99
            Top             =   1050
            Width           =   615
         End
         Begin VB.TextBox txtEmergContact 
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   300
            Width           =   3075
         End
         Begin VB.Label lblPartnerName 
            BackColor       =   &H00BFE8FD&
            BackStyle       =   0  'Transparent
            Caption         =   "Partner Name"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   30
            TabIndex        =   110
            Top             =   1530
            Width           =   1275
         End
         Begin VB.Label lblRel 
            BackStyle       =   0  'Transparent
            Caption         =   "Relationship"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   60
            TabIndex        =   109
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblEmergPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   60
            TabIndex        =   108
            Top             =   1080
            Width           =   525
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
            Left            =   1350
            TabIndex        =   107
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
            Left            =   2115
            TabIndex        =   106
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
            Left            =   2910
            TabIndex        =   105
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
            TabIndex        =   104
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
         TabIndex        =   59
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
            TabIndex        =   60
            Top             =   600
            Visible         =   0   'False
            Width           =   1665
         End
      End
      Begin VB.TextBox txtAreaCode 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6540
         MaxLength       =   3
         TabIndex        =   16
         Top             =   870
         Width           =   465
      End
      Begin VB.TextBox txtPhone1 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7170
         MaxLength       =   3
         TabIndex        =   17
         Top             =   870
         Width           =   465
      End
      Begin VB.TextBox txtPhone2 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7785
         MaxLength       =   4
         TabIndex        =   18
         Top             =   870
         Width           =   700
      End
      Begin VB.TextBox txtCellPhone2 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7785
         MaxLength       =   4
         TabIndex        =   21
         Top             =   1380
         Width           =   700
      End
      Begin VB.TextBox txtCellPhone1 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7170
         MaxLength       =   3
         TabIndex        =   20
         Top             =   1380
         Width           =   465
      End
      Begin VB.TextBox txtCellAreaCode 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6540
         MaxLength       =   3
         TabIndex        =   19
         Top             =   1380
         Width           =   465
      End
      Begin VB.TextBox txtWorkPh2 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7785
         MaxLength       =   4
         TabIndex        =   24
         Top             =   1800
         Width           =   700
      End
      Begin VB.TextBox txtWorkPh1 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7170
         MaxLength       =   3
         TabIndex        =   23
         Top             =   1800
         Width           =   465
      End
      Begin VB.TextBox txtWorkAC 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6540
         MaxLength       =   3
         TabIndex        =   22
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
         Picture         =   "PatientInfo_Old.frx":2846
         ScaleHeight     =   525
         ScaleWidth      =   405
         TabIndex        =   58
         Top             =   720
         Width           =   405
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1740
         TabIndex        =   0
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
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   1740
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1290
         Width           =   3135
      End
      Begin VB.TextBox txtAge 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4830
         MaxLength       =   3
         TabIndex        =   57
         Top             =   2790
         Width           =   600
      End
      Begin VB.TextBox txtSS3 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   7
         Top             =   3270
         Width           =   720
      End
      Begin VB.TextBox txtSS2 
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2670
         MaxLength       =   2
         TabIndex        =   6
         Top             =   3270
         Width           =   390
      End
      Begin VB.TextBox txtSS1 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1980
         MaxLength       =   3
         TabIndex        =   5
         Top             =   3270
         Width           =   510
      End
      Begin VB.TextBox txtDriverLicense 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1980
         MaxLength       =   25
         TabIndex        =   8
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
         TabIndex        =   9
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
            TabIndex        =   56
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
            TabIndex        =   55
            Top             =   150
            Width           =   825
         End
      End
      Begin VB.ComboBox cboPatientRace 
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "PatientInfo_Old.frx":3110
         Left            =   1965
         List            =   "PatientInfo_Old.frx":3112
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4890
         Width           =   2145
      End
      Begin VB.TextBox txtEmailAddress 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1965
         TabIndex        =   14
         Top             =   5820
         Width           =   2940
      End
      Begin VB.TextBox txtExtRef 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9720
         TabIndex        =   26
         ToolTipText     =   "For any other Identification Identification "
         Top             =   2880
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
         Picture         =   "PatientInfo_Old.frx":3114
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   54
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2850
         MaxLength       =   4
         TabIndex        =   4
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2250
         MaxLength       =   2
         TabIndex        =   3
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   2
         Top             =   2820
         Width           =   435
      End
      Begin VB.TextBox txtInternationalPhone 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6480
         TabIndex        =   25
         Top             =   2190
         Width           =   2010
      End
      Begin VB.ComboBox cboPatMarital 
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "PatientInfo_Old.frx":39DE
         Left            =   1965
         List            =   "PatientInfo_Old.frx":39E0
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   5400
         Width           =   2655
      End
      Begin VB.ComboBox cboFacility 
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6570
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3240
         Width           =   5445
      End
      Begin PT_XP_Button.PT_XP cmdName 
         Height          =   345
         Left            =   630
         TabIndex        =   62
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
         Height          =   375
         Left            =   630
         TabIndex        =   63
         Top             =   1350
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
         Left            =   3863
         TabIndex        =   33
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
         Picture         =   "PatientInfo_Old.frx":39E2
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
         PictureDisabled =   "PatientInfo_Old.frx":3F7C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdUpdate 
         Height          =   405
         Left            =   180
         TabIndex        =   64
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
         Picture         =   "PatientInfo_Old.frx":42CE
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
         PictureDisabled =   "PatientInfo_Old.frx":4868
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCancel 
         Height          =   405
         Left            =   6263
         TabIndex        =   31
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
         Picture         =   "PatientInfo_Old.frx":4BBA
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   "Cancel"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "PatientInfo_Old.frx":4F54
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   405
         Left            =   7433
         TabIndex        =   32
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
         Picture         =   "PatientInfo_Old.frx":52A6
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "PatientInfo_Old.frx":5CB8
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   405
         Left            =   5123
         TabIndex        =   30
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
         Picture         =   "PatientInfo_Old.frx":600A
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "PatientInfo_Old.frx":63A4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdTakePic 
         Height          =   465
         Left            =   9030
         TabIndex        =   65
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
         Picture         =   "PatientInfo_Old.frx":66F6
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "PatientInfo_Old.frx":6FD0
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdLangTerm 
         Height          =   225
         Left            =   5880
         TabIndex        =   112
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
         TabIndex        =   113
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
      Begin PT_XP_Button.PT_XP cmdPrint 
         Height          =   345
         Left            =   11160
         TabIndex        =   118
         Top             =   6360
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
         Picture         =   "PatientInfo_Old.frx":7C22
         AppearanceThemes=   3
         TransparentColor=   12582912
         Caption         =   "Print"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "PatientInfo_Old.frx":81BC
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblPatStatus 
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS"
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
         Left            =   5460
         TabIndex        =   115
         Top             =   3720
         Width           =   825
      End
      Begin VB.Label lblPref 
         BackStyle       =   0  'Transparent
         Caption         =   "Preferred Contact"
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
         Left            =   5100
         TabIndex        =   114
         Top             =   420
         Width           =   990
      End
      Begin VB.Label lblLang 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Language:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4290
         TabIndex        =   111
         Top             =   4665
         Width           =   915
      End
      Begin VB.Label lblPatient_id 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9720
         TabIndex        =   98
         Top             =   2520
         Width           =   2265
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
         TabIndex        =   97
         Top             =   2610
         Width           =   1515
      End
      Begin VB.Label lblFacRefNo 
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "AUX EMR ID"
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
         TabIndex        =   96
         ToolTipText     =   "Secondary EMR/Hospital ID Number"
         Top             =   2940
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
         Top             =   1830
         Width           =   75
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   300
         TabIndex        =   83
         Top             =   2880
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
         TabIndex        =   82
         Top             =   2850
         Width           =   1140
      End
      Begin VB.Label lblSocialSecurity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Social Security #"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   450
         TabIndex        =   81
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
         Left            =   3090
         TabIndex        =   80
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
         Left            =   2520
         TabIndex        =   79
         Top             =   3300
         Width           =   150
      End
      Begin VB.Label lblDriverLicense 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver License #"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   450
         TabIndex        =   78
         Top             =   3750
         Width           =   1410
      End
      Begin VB.Label lblGender 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   450
         TabIndex        =   77
         Top             =   4140
         Width           =   690
      End
      Begin VB.Label lblBloodType 
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Type"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   450
         TabIndex        =   76
         Top             =   4530
         Width           =   975
      End
      Begin VB.Label lblRace 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Race:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   450
         TabIndex        =   75
         Top             =   4965
         Width           =   465
      End
      Begin VB.Label lblEmailAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   450
         TabIndex        =   74
         Top             =   5955
         Width           =   570
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "yyyy"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2970
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
         Top             =   2220
         Width           =   1425
      End
      Begin VB.Label lblMarital 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marital Status:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   450
         TabIndex        =   67
         Top             =   5490
         Width           =   1305
      End
      Begin VB.Label lblFacility 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "FACILITY"
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
         Left            =   5460
         TabIndex        =   66
         Top             =   3330
         Width           =   930
      End
   End
End
Attribute VB_Name = "PatientInfo_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public piPatientID As Long
Public piEncounterID As Long
Public piProviderID As Long
Private fUpdateMode As String
Private mLocID As Integer
Private mbAdmit As Boolean
Private mbDischarge As Boolean
Private mRoomAssign As String
Public mNewPatient As Boolean
Private mFirst As Boolean
Private bEnable As Boolean
Private mintRelationPatID  As Integer
Private mbooClickOnce  As Boolean
Public mbNewPatInsert As Boolean
Public mbNotSave As Boolean
Public mbooByPass As Boolean
Public mbooData As Boolean
Public mbNewPatReg As Boolean
Private mbooValidSS As Boolean
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long

Private Function IsNewPatient() As Boolean
Dim rsPatient As New ADODB.Recordset
Dim cnPatient As New ADODB.Connection
Dim strCount As String
Dim strSQL As String
Dim lngDateDiff As Long
Dim strBirth As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field


    On Error GoTo Error_IsNewPatient
    Screen.MousePointer = vbHourglass
    If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
    IsNewPatient = True
    clsPatient.IntId = 0
    strSQL = "SELECT SOCIAL_SECURITY, BIRTH, PATIENT_ID FROM PATIENT WHERE LAST_NAME  = '" & clsPatient.LastName & "' AND FIRST_NAME =  '" & clsPatient.FirstName & "'"
    Set rsPatient = cnPatient.Execute(strSQL)
    Do Until rsPatient.EOF
        strBirth = vbNullString
        Set Flds = rsPatient.Fields
        clsPatient.IntId = rsPatient.Fields("PATIENT_ID")
        For Each Fld In Flds
            If Not IsNull(Fld.Value) Then
                If Trim(Fld.Value) <> vbNullString Then
                    Select Case UCase(Fld.Name)
                        Case "SOCIAL_SECURITY"
                            If clsPatient.SS = Fld.Value Then
                                IsNewPatient = False
                            End If
                        Case "BIRTH"
                            strBirth = Format(Fld.Value, "mm/dd/yyyy")
                            If IsDate(strBirth) And strBirth <> "01/01/1900" Then
                                clsPatient.DOB = Format(clsPatient.DOB, "mm/dd/yyyy")
                                lngDateDiff = DateDiff("d", strBirth, clsPatient.DOB)
                                If lngDateDiff = 0 Then
                                    IsNewPatient = False
                                    clsPatient.IntId = vbNullString
                                    Exit For
                                Else
                                    IsNewPatient = True
                                End If
                            Else
                                IsNewPatient = True
                            End If
                    End Select
                End If
            End If
        Next
        rsPatient.MoveNext
        clsPatient.IntId = vbNullString
    Loop
    If Val(clsPatient.IntId) > 0 Then
        mNewPatient = False
        mbNewPatReg = False
    End If
    If cnPatient.State = adStateOpen Then
        cnPatient.Close
        Set cnPatient = Nothing
        Set rsPatient = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_IsNewPatient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "IsNewPatient", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function
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

Private Function Load_Facilities()
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


    On Error GoTo Error_Load_Facilities
    Screen.MousePointer = vbHourglass
    cboFacility.Clear
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "SELECT a.DEFAULT_FACILITY,  b.FACILITY_NAME, b.FACILITY_ID  FROM APPOINTMENT_PROVIDERS a, " _
        & "FACILITIES b where a.FACILITY_ID = b.FACILITY_ID AND b.ACTIVE = 1 ORDER BY b.FACILITY_NAME"
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
            If Trim(strName) <> vbNullString Then
                strName = Fix_Name(strName)
                strName = Fix_Facility_Name(strName)
                For Idx = 0 To cboFacility.ListCount - 1
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
    lngFacID = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultFacility")
    If lngFacID > 0 Then
        For Idx = 0 To cboFacility.ListCount - 1
            If cboFacility.ItemData(Idx) = lngFacID Then
                cboFacility.Text = cboFacility.List(Idx)
                Exit For
            End If
        Next Idx
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
    Write_Error Err.Description, Err.Number, "", "Load_Facilities", App.EXEName & "." & Me.Name
    Exit Function
    Resume
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
    cboLanguage.ListIndex = -1
    cboPrefContact.ListIndex = -1
    txtParner.Text = vbNullString
    txtEmergContact.Text = vbNullString
    cboRelation.Clear
    txtEmergAC.Text = vbNullString
    txtEmergPh1.Text = vbNullString
    txtEmergPh2.Text = vbNullString
    cboBloodType.ListIndex = -1
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
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim sStr As String

    On Error GoTo Error_GetTerms
    Screen.MousePointer = vbHourglass
   If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    vobjCombo.Clear
    sStr = "SELECT DISTINCT ATTRBTE_VALID_VALUE FROM ATTR_VALID_VALUE WHERE VALID_VALUE_SUBDOM = '" & vstrSubDomain & "'"
    Set rsTerms = cnTerms.Execute(sStr)
    With rsTerms
        Do Until .EOF
            If Trim(!ATTRBTE_VALID_VALUE) <> vbNullString Then
                vobjCombo.AddItem Trim(!ATTRBTE_VALID_VALUE)
            End If
            .MoveNext
        Loop
    End With
    If cnTerms.State = adStateOpen Then
         cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_GetTerms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, sStr, "GetTerms", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Function Get_RelationshipTerms(vobjCombo As Object)
Dim strSQL As String
Dim cnRelation As New ADODB.Connection
Dim rsRelation As New ADODB.Recordset

    On Error GoTo Error_Get_RelationshipTerms
    Screen.MousePointer = vbHourglass
    vobjCombo.Clear
    If cnRelation.State = adStateClosed Then cnRelation.open psConnect(1)
    strSQL = "Select * from RELATION_CODES "
    Set rsRelation = cnRelation.Execute(strSQL)
    With rsRelation
        Do Until .EOF
            If (!Relationship) <> vbNullString Then vobjCombo.AddItem (!Relationship)
            If IsNumeric(!Relationship_Code) Then vobjCombo.ItemData(vobjCombo.NewIndex) = (!Relationship_Code)
            .MoveNext
        Loop
    End With
    If cnRelation.State = adStateOpen Then
        cnRelation.Close
        Set cnRelation = Nothing
        Set rsRelation = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_RelationshipTerms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_RelationshipTerms", App.EXEName & "." & "PatientInfo"
    Exit Function
    Resume
End Function


Private Function Update_Archive(vstrExtRef As String, vstrTitle As String, vstrPatFName As String, vstrPatMName As String, _
vstrPatLName As String, vstrPatAddress As String, vstrPatAddress2 As String, vstrPatCity As String, vstrPatState As String, vstrPatZip As String, _
vstrPatZipExt As String, vstrPatCountry As String, vstrAreaCode As String, vstrPhone As String, vstrCellPhone As String, vstrDL As String, _
vstrRace As String, vstrStatus As String, vstrPatSuffix As String, _
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
    On Error GoTo Error_Update_Archive
    Screen.MousePointer = vbHourglass
    strSQL = "SELECT * FROM PATIENT WHERE PATIENT_ID = " & piPatientID
    Set rsPatient = cnPatient.Execute(strSQL)
    If Not rsPatient.EOF Then
        Set Flds = rsPatient.Fields
        For Each Fld In Flds
            If Not IsNull(Fld.Value) Then
                If Trim(Fld.Value) <> vbNullString Then
                    Select Case UCase(Fld.Name)
                        Case UCase("Patient_Ext_Ref")
                            If Fld.Value <> vstrExtRef Then booSave = True
                        Case UCase("title")
                            If Fld.Value <> Fld.Value <> vstrTitle Then booSave = True
                        Case UCase("first_name")
                            If Fld.Value <> Fld.Value <> vstrPatFName Then booSave = True
                        Case UCase("middle_name")
                            If Fld.Value <> Fld.Value <> vstrPatMName Then booSave = True
                        Case UCase("last_name")
                            If Fld.Value <> Fld.Value <> vstrPatLName Then booSave = True
                        Case UCase("address1")
                            If Fld.Value <> Fld.Value <> vstrPatAddress Then booSave = True
                        Case UCase("address2")
                            If Fld.Value <> Fld.Value <> vstrPatAddress2 Then booSave = True
                        Case UCase("city")
                            If Fld.Value <> Fld.Value <> vstrPatCity Then booSave = True
                        Case UCase("state")
                            If Fld.Value <> Fld.Value <> vstrPatState Then booSave = True
                        Case UCase("zipcode")
                            If Fld.Value <> Fld.Value <> vstrPatZip Then booSave = True
                        Case UCase("zipcode_ext")
                            If Fld.Value <> Fld.Value <> vstrPatZipExt Then booSave = True
                        Case UCase("country")
                            If Fld.Value <> Fld.Value <> vstrPatCountry Then booSave = True
                        Case UCase("area_code")
                            If Fld.Value <> Fld.Value <> vstrAreaCode Then booSave = True
                        Case UCase("phone")
                            If Fld.Value <> Fld.Value <> vstrPhone Then booSave = True
                        Case UCase("cell_phone")
                            If Fld.Value <> Fld.Value <> vstrCellPhone Then booSave = True
                        Case UCase("drv_license")
                            If Fld.Value <> Fld.Value <> vstrDL Then booSave = True
                        Case UCase("Race")
                            If Fld.Value <> Fld.Value <> vstrRace Then booSave = True
                        Case UCase("Status")
                            If Fld.Value <> Fld.Value <> vstrStatus Then booSave = True
                        Case UCase("Suffix")
                            If Fld.Value <> Fld.Value <> vstrPatSuffix Then booSave = True
                        Case UCase("birth")
                            If Fld.Value <> Fld.Value <> vstrBirth Then booSave = True
                        Case UCase("social_security")
                            If Fld.Value <> Fld.Value <> vstrSocSec Then booSave = True
                        Case UCase("gender")
                            If Fld.Value <> Fld.Value <> vstrGender Then booSave = True
                        Case UCase("email")
                            If Fld.Value <> Fld.Value <> vstrEmailAddress Then booSave = True
                        Case UCase("internatl_phone")
                            If Fld.Value <> Fld.Value <> vstrIntnPhone Then booSave = True
                        Case UCase("county")
                            If Fld.Value <> Fld.Value <> vstrPatCounty Then booSave = True
                        Case UCase("blood_Type")
                            If Fld.Value <> Fld.Value <> vstrBloodType Then booSave = True
                        Case UCase("marital_status")
                            If Fld.Value <> Fld.Value <> vstrMaritalStat Then booSave = True
                        Case UCase("work_phone")
                            If Fld.Value <> Fld.Value <> vstrWorkPhone Then booSave = True
                        Case UCase("EMERGENCY_CONTACT")
                            If Fld.Value <> Fld.Value <> vstrEmerContact Then booSave = True
                        Case UCase("EMERG_PHONE")
                            If Fld.Value <> Fld.Value <> vstrEMERG_PHONE Then booSave = True
                        Case UCase("EMERG_RELATION")
                            If Fld.Value <> Fld.Value <> vstrEmergRel Then booSave = True
                    End Select
                End If
            End If
            If booSave Then Exit For
        Next
    End If
    Get_ArchiveData
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
    strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If udtPatientName.WorkPhone <> vbNullString Then
        udtPatientName.WorkPhone = Replace(udtPatientName.WorkPhone, "(", vbNullString)
        udtPatientName.WorkPhone = Replace(udtPatientName.WorkPhone, ")", "-")
        udtPatientName.WorkPhone = Replace(udtPatientName.WorkPhone, " ", "")
        udtPatientName.WorkPhone = Trim(udtPatientName.WorkPhone)
    End If
    udtPatientName.ExtID = Replace(udtPatientName.ExtID, "'", "''")
    If TextFound("p.o.", udtPatientName.Address, Len("p.o."), True) Then
        udtPatientName.Address = Replace(udtPatientName.Address, "p.o.", "Post Office")
    End If
    PsUserID = Get_PsUserID(psConnect(0))
    udtPatientName.Address2 = Fix_Name(udtPatientName.Address2)
    udtPatientName.Address = Fix_Name(udtPatientName.Address)
    udtPatientName.County = Replace(udtPatientName.County, "'", "''")
    udtPatientName.LastName = Replace(udtPatientName.LastName, "'", "''")
    udtPatientName.Address2 = Replace(udtPatientName.Address2, "'", "''")
    udtPatientName.Address = Replace(udtPatientName.Address, "'", "''")
    udtPatientName.FirstName = Replace(udtPatientName.FirstName, "'", "''")
    udtPatientName.Address = Replace(udtPatientName.Address, "'", "''")
    udtPatientName.Address2 = Replace(udtPatientName.Address2, "'", "''")
    udtPatientName.City = Replace(udtPatientName.City, "'", "''")
    
    If lngPrimKey > 0 Then
        strSQL = "UPDATE PATIENT_ARCHIVE_DEMOGRAPHIC SET PATIENT_EXT_REF = '" & udtPatientName.ExtID & "', TITLE = '" & udtPatientName.Title & "', FIRST_NAME  = '" & udtPatientName.FirstName & "'," _
            & "MIDDLE_NAME = '" & udtPatientName.MiddleName & "', LAST_NAME = '" & udtPatientName.LastName & "', RACE  = '" & udtPatientName.Race & "'," _
            & "PREF_CONTACT = '" & udtPatientName.PrefContact & "',  STATUS  = '" & udtPatientName.Status & "', COUNTY  = '" & udtPatientName.County & "'," _
            & "SUFFIX = '" & udtPatientName.Suffix & "', ADDRESS1  = '" & udtPatientName.Address & "',  ADDRESS2 =  '" & udtPatientName.Address2 & "'," _
            & "CITY = '" & udtPatientName.City & "', STATE = '" & udtPatientName.State & "',  PHONE = '" & udtPatientName.Phone & "'," _
            & "INTERNATL_PHONE = '" & udtPatientName.InterNalPhone & "', SOCIAL_SECURITY = '" & udtPatientName.SS & "',  MARITAL_STATUS = '" & udtPatientName.Marital_Status & "'," _
            & "AREA_CODE = '" & udtPatientName.AreaCode & "', BIRTH = '" & udtPatientName.DOB & "', LANGUAGE = '" & udtPatientName.Language & "'," _
            & "ZIPCODE = '" & Left(udtPatientName.Zip, 5) & "', ZIPCODE_EXT = '" & udtPatientName.Zip_Ext & "', GENDER = '" & Left(udtPatientName.Gender, 1) & "'," _
            & "CELL_PHONE =  '" & udtPatientName.CellPhone & "', WORK_PHONE = '" & udtPatientName.WorkPhone & "', EMAIL = '" & udtPatientName.Email & "', DRV_LICENSE = '" & udtPatientName.DL & " '," _
            & "USER_NAME = '" & PsUserID & "', UPDATE_DATE = '" & strDate & "', CHANGE_DATE = '" & strDate & "'  WHERE PREV_DEMO_ID = " & lngPrimKey
            Set rsPatient = cnPatient.Execute(strSQL)
    Else
        strSQL = "INSERT INTO PATIENT_ARCHIVE_DEMOGRAPHIC (PATIENT_ID, CHANGE_DATE, PATIENT_EXT_REF, TITLE, FIRST_NAME, MIDDLE_NAME, LAST_NAME,RACE, PREF_CONTACT," _
            & "STATUS, COUNTY,  SUFFIX, ADDRESS1, ADDRESS2, CITY,STATE, PHONE, INTERNATL_PHONE, SOCIAL_SECURITY, MARITAL_STATUS, AREA_CODE, BIRTH, LANGUAGE, " _
            & "ZIPCODE, ZIPCODE_EXT, GENDER, CELL_PHONE, WORK_PHONE, EMAIL, DRV_LICENSE, USER_NAME, INSERT_DATE, PURGE) values (" & piPatientID & ", '" _
            & strDate & "', '" & udtPatientName.ExtID & "','" & udtPatientName.Title & "','" _
            & udtPatientName.FirstName & "','" & udtPatientName.MiddleName & "','" & udtPatientName.LastName & "','" & udtPatientName.Race & "','" _
            & udtPatientName.PrefContact & "','" & udtPatientName.Status & "','" & udtPatientName.County & "','" _
            & udtPatientName.Suffix & "','" & udtPatientName.Address & "','" & udtPatientName.Address2 & "','" _
            & udtPatientName.City & "','" & udtPatientName.State & "','" & Trim(udtPatientName.Phone) & "','" & Trim(udtPatientName.InterNalPhone) & "','" & udtPatientName.SS & "','" _
            & udtPatientName.Marital_Status & "', '" & udtPatientName.AreaCode & "','" & udtPatientName.DOB & "','" & udtPatientName.Language & "','" & Left(udtPatientName.Zip, 5) & "','" _
            & udtPatientName.Zip_Ext & "','" & Left(udtPatientName.Gender, 1) & "','" & udtPatientName.CellPhone & "','" & udtPatientName.WorkPhone & "','" & udtPatientName.Email & "','" _
            & udtPatientName.DL & "', '" & PsUserID & "','" & strDate & "',0 )"
            Set rsPatient = cnPatient.Execute(strSQL)
    End If
    If cnPatient.State = adStateOpen Then
        cnPatient.Close
        Set cnPatient = Nothing
        Set rsPatient = Nothing
    End If
    Clear_UdtPatient
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_Archive:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Error_Update_Archive", "PatientTracSchedule.PatientInfo"
    Exit Function
End Function
Private Function Get_ArchiveData()
Dim strSQL As String
Dim strEnctrStart As String
Dim cnPatient As New ADODB.Connection
Dim rsPatient As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim strChDate As String
Dim strInqDate As String
Dim strFormatChg As String
Dim lngPrimKey As Long


    On Error GoTo Error_Get_ArchiveData
    Screen.MousePointer = vbHourglass
    Clear_UdtPatient
    If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
    If piPatientID > 0 Then
        strSQL = "SELECT * From PATIENT WHERE PATIENT_ID  = " & piPatientID
        Set rsPatient = cnPatient.Execute(strSQL)
        If Not rsPatient.EOF Then
            Set Flds = rsPatient.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PREF_CONTACT"
                                udtPatientName.PrefContact = Fld.Value
                            Case "OTHER_INFO3"
                                udtPatientName.DateDeath = Format(Fld.Value, "mm/dd/yyyy")
                            Case "LANGUAGE"
                                udtPatientName.Language = IsProperCase(Fld.Value)
                            Case "FIRST_NAME"
                                udtPatientName.FirstName = Fix_Name(Fld.Value)
                            Case "MIDDLE_NAME"
                                udtPatientName.MiddleName = IsProperCase(Fld.Value)
                            Case "LAST_NAME"
                                udtPatientName.LastName = IsProperCase(Fld.Value)
                                If TextFound("''", udtPatientName.LastName, 2, True) Then
                                    udtPatientName.LastName = Replace(udtPatientName.LastName, "''", "'")
                                    udtPatientName.LastName = udtPatientName.LastName
                                Else
                                    udtPatientName.LastName = udtPatientName.LastName
                                End If
                                udtPatientName.LastName = Fix_Name(Fld.Value)
                            Case "GENDER"
                                udtPatientName.Gender = IsProperCase(Fld.Value)
                                If udtPatientName.Gender = "F" Or udtPatientName.Gender = "Female" Then
                                    udtPatientName.Gender = "Female"
                                    udtPatientName.Person = "She"
                                    udtPatientName.Possessive = "Her"
                                    udtPatientName.PerPosses = "Herself"
                                Else
                                    udtPatientName.Gender = "Male"
                                    udtPatientName.Person = "He"
                                    udtPatientName.Possessive = "His"
                                    udtPatientName.PerPosses = "Himself"
                                End If
                            Case "TITLE"
                                udtPatientName.Title = IsProperCase(Fld.Value)
                                If udtPatientName.Title = vbNullString Then
                                    If udtPatientName.Gender = "M" Or udtPatientName.Gender = "Male" Then
                                        udtPatientName.Title = "Mr."
                                    Else
                                        If udtPatientName.Marital_Status = "married" Or udtPatientName.Marital_Status = UCase("Married") And _
                                            udtPatientName.Gender = "F" Then
                                            udtPatientName.Title = "Mrs."
                                        Else
                                            udtPatientName.Title = "Ms."
                                        End If
                                    End If
                                End If
                                udtPatientName.Title = Fix_Title(udtPatientName.Title)
                            Case "MARITAL_STATUS"
                                udtPatientName.Marital_Status = IsProperCase(Fld.Value)
                            Case "SUFFIX"
                                udtPatientName.Suffix = IsProperCase(Fld.Value)
                                udtPatientName.Suffix = Replace(udtPatientName.Suffix, ",", "")
                                udtPatientName.Suffix = Fix_Suffix(udtPatientName.Suffix)
                            Case "BIRTH"
                                udtPatientName.DOB = Format(Fld.Value, "mm/dd/yyyy")
                            Case "PATIENT_EXT_REF"
                                udtPatientName.ExtID = Fld.Value
                            Case "PATIENT_ID"
                                udtPatientName.IntId = Fld.Value
                            Case "SOCIAL_SECURITY"
                                udtPatientName.SS = Fld.Value
                            Case "ADDRESS1"
                                udtPatientName.Address = Fix_Name(Fld.Value)
                            Case "ADDRESS2"
                                udtPatientName.Address2 = Fix_Name(Fld.Value)
                            Case "CITY"
                                udtPatientName.City = Fix_Name(Fld.Value)
                            Case "STATE"
                                udtPatientName.State = IsProperCase(Fld.Value)
                            Case "ZIPCODE"
                                udtPatientName.Zip = Fld.Value
                            Case "ZIPCODE_EXT"
                                udtPatientName.Zip_Ext = Fld.Value
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
                            Case "INTERNATL_PHONE"
                                 udtPatientName.InterNalPhone = LCase(Fld.Value)
                            Case "PHONE"
                                 udtPatientName.Phone = LCase(Fld.Value)
                            Case "AREA_CODE"
                                udtPatientName.AreaCode = LCase(Fld.Value)
                               ' udtPatientName.Phone = udtPatientName.AreaCode & "-" & udtPatientName.Phone
                            Case "EMERG_PHONE"
                                udtPatientName.EmergPhone = LCase(Fld.Value)
                                clsEmerg.Phone = udtPatientName.EmergPhone
                            Case "EMERG_RELATION"
                                udtPatientName.EmergRelation = LCase(Fld.Value)
                                clsEmerg.Relation = udtPatientName.EmergRelation
                            Case "EMERGENCY_CONTACT"
                                udtPatientName.EmergContact = LCase(Fld.Value)
                                clsEmerg.Name = udtPatientName.EmergContact
                            Case "STATUS"
                                 udtPatientName.Status = LCase(Fld.Value)
                            Case "ROOM_PHONE"
                                 udtPatientName.RoomPhone = LCase(Fld.Value)
                            Case "RACE"
                                 udtPatientName.Race = LCase(Fld.Value)
                            Case "BLOOD_TYPE"
                                 udtPatientName.Blood_Type = LCase(Fld.Value)
                        End Select
                    End If
                End If
            Next
            If udtPatientName.DOB <> vbNullString Then
                If clsEncounter.StartDate <> vbNullString Then
                    strEnctrStart = Format(clsEncounter.StartDate, "mm/dd/yyyy")
                Else
                    strEnctrStart = Format(Now, "mm/dd/yyyy")
                End If
                udtPatientName.Age = Int(DateDiff("y", udtPatientName.DOB, strEnctrStart) / 365.25)
            End If
            udtPatientName.Suffix = Fix_Suffix(udtPatientName.Suffix)
            udtPatientName.Address = Fix_Name(udtPatientName.Address)
            udtPatientName.FirstName = Fix_Name(udtPatientName.FirstName)
            udtPatientName.LastName = Fix_Name(udtPatientName.LastName)
            udtPatientName.City = Fix_Name(udtPatientName.City)
            If Trim(udtPatientName.Suffix) <> vbNullString Then
                udtPatientName.Name = udtPatientName.FirstName & " " & udtPatientName.LastName & ", " & udtPatientName.Suffix
            Else
                udtPatientName.Name = udtPatientName.FirstName & " " & udtPatientName.LastName
            End If
            udtPatientName.City = Replace(udtPatientName.City, " - ", "-")
            udtPatientName.City = Trim(udtPatientName.City)
            If Trim(udtPatientName.City) <> vbNullString Then
                 udtPatientName.CityStateZip = Trim(udtPatientName.City)
             End If
            If Trim(udtPatientName.State) <> vbNullString Then
                 udtPatientName.CityStateZip = udtPatientName.CityStateZip & ", " & Trim(udtPatientName.State)
             End If
            If Trim(udtPatientName.Zip) <> vbNullString Then
                 udtPatientName.CityStateZip = udtPatientName.CityStateZip & " " & Trim(udtPatientName.Zip)
             End If
            If udtPatientName.Title = vbNullString Then
                If udtPatientName.Gender = "M" Or udtPatientName.Gender = "Male" Then
                    udtPatientName.Title = "Mr."
                Else
                    If udtPatientName.Marital_Status = "married" Or udtPatientName.Marital_Status = UCase("Married") And _
                        udtPatientName.Gender = "F" Then
                        udtPatientName.Title = "Mrs."
                    Else
                        udtPatientName.Title = "Ms."
                    End If
                End If
            End If
        End If
    End If
    If cnPatient.State = adStateOpen Then
         cnPatient.Close
         Set cnPatient = Nothing
         Set rsPatient = Nothing
         Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_ArchiveData:
   Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_ArchiveData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
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
Dim Date1 As Date
Dim Date2 As Date
Dim Arr
Dim i%

    Validate_Entries = True
    Arr = Split(txtFullName.Text, " ")
    If UBound(Arr) < 1 Then
      MsgBox "First name And last name required", vbExclamation, "PatientTrac Patient Information"
        Validate_Entries = False
    End If
    If Trim(txtMonth.Text) <> vbNullString Then
       If Not IsNumeric(txtMonth.Text) Then
          MsgBox "Invalid Month for Birth Date", vbExclamation, "PatientTrac Patient Information"
          Validate_Entries = False
          Exit Function
       End If
       If CInt(txtMonth.Text) < 1 Or CInt(txtMonth.Text) > 12 Then
          MsgBox "Invalid Month for Birth Date", vbExclamation, "PatientTrac Patient Information"
          Validate_Entries = False
          Exit Function
       End If
    End If
    If optFemale.Value = False And optMale.Value = False Then
          MsgBox "Patient requires a Gender!", vbExclamation, "PatientTrac Patient Information"
          Validate_Entries = False
          Exit Function
    End If
    If Trim(txtAddress.Text) = vbNullString Then
          MsgBox "Address is a mandatory field. Please enter an address!", vbExclamation, "PatientTrac Patient Information"
          Validate_Entries = False
          Exit Function
    End If
    If Trim(txtYear.Text) <> vbNullString Then
       If Not IsNumeric(txtYear.Text) Or Len(txtYear.Text) <> 4 Then
          MsgBox "Invalid Year for Birth Date", vbExclamation, "PatientTrac Patient Information"
          Validate_Entries = False
          Exit Function
       End If
       If CInt(txtYear.Text) > Year(Now) Then
          MsgBox "Invalid Year for Birth Date", vbExclamation, "PatientTrac Patient Information"
          Validate_Entries = False
          Exit Function
       End If
    End If
    If Trim(txtDay.Text) <> vbNullString Then
       If Not IsNumeric(txtDay.Text) Then
          MsgBox "Invalid Day for Birth Date", vbExclamation, "PatientTrac Patient Information"
          Validate_Entries = False
          Exit Function
       End If
       If Not IsDayCheck(CInt(txtMonth), CInt(txtDay), CInt(txtYear)) Then
          MsgBox "Invalid Day for Birth Date", vbExclamation, "PatientTrac Patient Information"
          Validate_Entries = False
          Exit Function
       End If
       If mstrPractice = "Psychiatry" Then
            If Val(txtAge.Text) < 18 And Trim(txtParFirstName.Text) = vbNullString Then
                 lblPartnerName.Caption = "Guardian"
             End If
        End If
    End If
        If Trim(txtEmergPh2.Text) <> vbNullString And Len(txtEmergPh2.Text) < 4 Then
             MsgBox "Invalid data entry! Telephone fields require numeric data. Erroneous data will not be removed. Please re-enter valid data. ", vbInformation, "Patient Demographic Maintanence"
             txtEmergPh2.Text = vbNullString
            Validate_Entries = False
            txtEmergPh2.SetFocus
            Exit Function
         End If
        If Trim(txtEmergPh1.Text) <> vbNullString And Len(txtEmergPh1.Text) < 3 Then
            MsgBox "Invalid data entry! Telephone fields require numeric data. Erroneous data will not be removed. Please re-enter valid data. ", vbInformation, "Patient Demographic Maintanence"
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

    If mbooByPass Then mbooByPass = False: Exit Sub
    If cboFacility.ListCount > 0 And Trim(cboFacility.Text) <> vbNullString Then
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
    End If
    Exit Sub
End Sub

Private Sub cboLanguage_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    If Len(cboLanguage.Text) >= 1 Then
        strList = cboLanguage.Text
        cboLanguage.ListIndex = SendMessage(cboLanguage.hwnd, LB_FINDSTRING, -1, _
        ByVal strList)
    End If
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


Private Sub cboPartRel_Change()
Dim strList As String
    If Len(cboPartRel.Text) >= 1 Then
        strList = cboPartRel.Text
        cboPartRel.ListIndex = SendMessage(cboPartRel.hwnd, LB_FINDSTRING, -1, _
        ByVal strList)
    End If
    Exit Sub
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
Dim strList As String
    If Len(cboPatMarital.Text) >= 1 Then
        strList = cboPatMarital.Text
        cboPatMarital.ListIndex = SendMessage(cboPatMarital.hwnd, LB_FINDSTRING, -1, _
        ByVal strList)
    End If
    Exit Sub
End Sub


Private Sub cboPrefContact_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String

    If Len(cboPrefContact.Text) >= 1 Then
        strList = cboPrefContact.Text
        cboPrefContact.ListIndex = SendMessage(cboPrefContact.hwnd, LB_FINDSTRING, -1, _
        ByVal strList)
    End If
    Exit Sub
End Sub


Private Sub cboRelation_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    If Len(cboRelation.Text) >= 1 Then
        strList = cboRelation.Text
        cboRelation.ListIndex = SendMessage(cboRelation.hwnd, LB_FINDSTRING, -1, _
        ByVal strList)
    End If
    Exit Sub
End Sub


Private Sub cboStatus_Click()
    If cboStatus.Text = "DECEASED" Then
        ccdDateDeath.Enabled = True
        ccdDateDeath.CheckBox = True
        fraDOD.Visible = True
    Else
        ccdDateDeath.Enabled = False
        ccdDateDeath.CheckBox = True
        fraDOD.Visible = False
    End If
    Exit Sub
End Sub


Private Sub cboSuffix_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    If Len(cboSuffix.Text) >= 1 Then
        strList = cboSuffix.Text
        cboSuffix.ListIndex = SendMessage(cboSuffix.hwnd, LB_FINDSTRING, -1, _
        ByVal strList)
    End If
    Exit Sub
End Sub


Private Sub cmdAddress_Click()
Dim strAddress As String
        If mbooByPass Then mbooByPass = False: Exit Sub
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


Private Sub cmdCancel_Click()
    mNewPatient = False
    Clear_Fields
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    piPatientID = 0
    frmPatients.mbShowPatient = False
    pbCancel = True
    Unload Me
    'frmPatients.Show
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
    If mbooByPass Then mbooByPass = False: Exit Sub
    txtFullName.Text = vbNullString
    If mNewPatient Then
        clsPatient.Title = vbNullString
        clsPatient.FirstName = vbNullString
        clsPatient.MiddleName = vbNullString
        clsPatient.LastName = vbNullString
        clsPatient.Suffix = vbNullString
    End If
   frmPatientName.Show vbModal
    If piPatientID > 0 And clsPatient.LastName = vbNullString Then
        Get_Patient piPatientID
    End If
    If Trim(clsPatient.Title) <> vbNullString Then txtFullName.Text = Trim(clsPatient.Title)
    If Trim(clsPatient.FirstName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.FirstName)
    If Trim(clsPatient.MiddleName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.MiddleName)
    If Trim(clsPatient.LastName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.LastName)
    If Trim(clsPatient.Suffix) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & ", " & Trim(clsPatient.Suffix)
    Exit Sub
End Sub

Private Sub cmdNew_Click()
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

    On Error GoTo Error_SaveNotes
    Screen.MousePointer = vbHourglass
    If strControl = "PersonalInfoN" Then
            rconClinical.A_PATIENT_PROFILER_NOTES _
            "X", piPatientID, strOCX & "." & strControl, vstrRef, vintSequence, vstrValue, vbooPrimary
    End If
    If strControl = "PersonalInfoN" Then
            rconClinical.A_PATIENT_PROFILER _
            "X", piPatientID, strOCX & "." & strControl, vstrRef, vintSequence, vstrValue, vbooPrimary
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_SaveNotes:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "SaveNotes", "PatientTracSchedule.PatientInfo"
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


Private Sub cmdPrint_Click()
    cmdSave_Click
    mstrPatOcx = "PatientTrac_Report": mstrPatCtl = "PatientDemographic"
    frmLoadControl.Caption = "PATIENT DEMOGRAPHICS"
    frmLoadControl.piSelPatientID = piPatientID
    frmLoadControl.Show vbModal
    Me.Show vbModal
    Exit Sub
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
'    If cboFacility.Text = vbNullString Then
'        MsgBox "Please select initial facililty. ", vbInformation
'        Exit Sub
'    End If
    If Validate_Entries() Then
       If mNewPatient Then
            piPatientID = 0
            piPatientID = Save_Data()
            If Not mNewPatient Then
                mNewPatient = False
                If mbNotSave Then
                    Get_Patient piPatientID
                    Exit Sub
                End If
                Unload Me
            Else
                If mbNotSave Then
                    Get_Patient piPatientID
                    Exit Sub
                End If
            End If
       Else
            UpdatePatientInfo
            mNewPatient = False
            If mbNotSave Then
                Get_Patient piPatientID
                Exit Sub
            End If
       End If
    Else
       Get_Patient piPatientID
       Exit Sub
    End If
    mNewPatient = False
    mbooClickOnce = True
    frmPatients.mbShowPatient = False
    Screen.MousePointer = vbDefault
    Unload Me
    frmPatients.Show
    Exit Sub
Error_Handler:
    MsgBox Err.Description, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub
Private Sub cmdUpdate_Click()
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
    cboRelation.Enabled = True
    cmdSave.ZOrder 0
    cmdSave.Visible = True
    cmdSave.Enabled = True
    cmdSave.Left = cmdNew.Left
    cmdNew.Visible = False
    Exit Sub
End Sub

Private Sub Form_Activate()
Dim strBirth As String
Dim strSocialSec As String
Dim varSocSec
Dim varCellPhone
Dim sStr As String
Dim strSQL As String
Dim strState As String
Dim i%

    If mFirst = True Then
        mFirst = False
        bEnable = True
        Clear_Fields
        Clear_Emerg
        Clear_PatientStrings
        Dump_PatientStrData
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
            piPatientID = 0
            Clear_PatientStrings
            Clear_SchedPatient
            Dump_PatientStrData
            udtSchedPatient.IntId = vbNullString
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
        If cboLanguage.ListCount <= 0 Then GetTerms cboLanguage, "sd_Patient_Language"
        centerForm Me
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Handler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbInformation
    Exit Sub
    Resume
    Exit Sub
End Sub

Private Function Save_Data() As Long
Dim rsPatient As New ADODB.Recordset
Dim cnPatient As New ADODB.Connection
Dim strSQL As String

    On Error GoTo Error_Save_Data
    Screen.MousePointer = vbHourglass
    If mNewPatient = False Then
        If piPatientID = 0 Then
            MsgBox "No Patient Entered. Exiting now!", vbInformation, "PatientTrac Patient Repository"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    clsPatient.LastName = Literals(clsPatient.LastName, "'")
    clsPatient.FirstName = Literals(clsPatient.FirstName, "'")
    Validate_Entries
    mbooData = True
    If Check_ValidPhone Then
        If txtAreaCode.Text <> vbNullString Then clsPatient.AreaCode = Trim(txtAreaCode.Text)
        If txtPhone1.Text <> vbNullString Then clsPatient.Phone = Trim(txtPhone1.Text)
        If txtPhone2.Text <> vbNullString Then clsPatient.Phone = Trim(clsPatient.Phone) & "-" & Trim(txtPhone2.Text)
        If txtCellAreaCode.Text <> vbNullString Then clsPatient.CellPhone = Trim(txtCellAreaCode.Text)
        If txtCellPhone1.Text <> vbNullString Then clsPatient.CellPhone = Trim(clsPatient.CellPhone) & "-" & txtCellPhone1 & "-" & txtCellPhone2
        If txtWorkAC.Text <> vbNullString Then clsPatient.WorkPhone = txtWorkAC
        If txtWorkPh1.Text <> vbNullString Then clsPatient.WorkPhone = Trim(clsPatient.WorkPhone) & "-" & txtWorkPh1 & "-" & txtWorkPh2
    End If
    If Trim(cboLanguage.Text) <> vbNullString Then clsPatient.Language = Trim(cboLanguage.Text)
    If Trim(cboPrefContact.Text) <> vbNullString Then clsPatient.PrefContact = Trim(cboPrefContact.Text)
    If Trim(cboBloodType.Text) <> vbNullString Then clsPatient.Blood_Type = Trim(cboBloodType.Text)
    If Trim(cboStatus.Text) <> vbNullString Then clsPatient.Status = Trim(cboStatus.Text)
    If Trim(cboPatientRace.Text) <> vbNullString Then clsPatient.Race = Trim(cboPatientRace.Text)
    If Trim(txtDriverLicense.Text) <> vbNullString Then clsPatient.DL = Trim(txtDriverLicense.Text)
    If Trim(txtInternationalPhone.Text) <> vbNullString Then clsPatient.InterNalPhone = Trim(txtInternationalPhone.Text)
    If Trim(txtEmailAddress.Text) <> vbNullString Then clsPatient.Email = Trim(txtEmailAddress.Text)
    If txtMonth.Text <> vbNullString Then clsPatient.DOB = txtMonth.Text
    If txtDay.Text <> vbNullString Then clsPatient.DOB = Trim(clsPatient.DOB) & "/" & txtDay.Text
    If txtYear.Text <> vbNullString Then clsPatient.DOB = Trim(clsPatient.DOB) & "/" & txtYear.Text
    If Trim(clsPatient.DOB) <> vbNullString Then clsPatient.DOB = Format(clsPatient.DOB, "mm/dd/yyyy hh:mm:ss AM/PM")
    If Not IsDate(clsPatient.DOB) Then
        mbNotSave = True
        MsgBox "Date of Birth is mandatory field!", vbInformation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If optMale.Value Then clsPatient.Gender = "M"
    If optFemale.Value Then clsPatient.Gender = "F"
    clsPatient.Title = Fix_Title(clsPatient.Title)
    clsPatient.FirstName = Fix_Name(clsPatient.FirstName)
    clsPatient.MiddleName = Fix_Name(clsPatient.MiddleName)
    clsPatient.LastName = Fix_Name(clsPatient.LastName)
    clsPatient.Suffix = Fix_Suffix(clsPatient.Suffix)
    clsPatient.Blood_Type = cboBloodType.Text
    clsPatient.ExtID = Literals(Trim(txtExtRef), "'")
    clsPatient.Marital_Status = cboPatMarital.Text
    If txtSS1.Text <> vbNullString Then
        clsPatient.SS = txtSS1.Text & "-" & txtSS2.Text & "-" & txtSS3.Text
    Else
        txtSS1.Text = "999"
        txtSS2.Text = "99"
        txtSS3.Text = "9999"
        clsPatient.SS = txtSS1.Text & "-" & txtSS2.Text & "-" & txtSS3.Text
    End If
    If Not IsNewPatient Then
         MsgBox "PatientTrac cannot save data since the Patient's Name, Date of Birth and Social Security Number" _
            & " " & "matches an existing Patient record. Please select Cancel or enter a new Date of Birth or Social Security Number", vbInformation, "Patient Maintance"
            mbNotSave = True
            mbNewPatReg = False
            txtSS1.Text = vbNullString
            txtSS2.Text = vbNullString
            txtSS3.Text = vbNullString
    Else
        piPatientID = 0
        clsPatient.IntId = vbNullString
        Insert_Patient
        If piPatientID > 0 Then
            Save_Data = piPatientID
            mbNotSave = False
            mbNewPatReg = True
            clsPatient.FirstName = Replace(clsPatient.FirstName, "''", "'")
            clsPatient.LastName = Replace(clsPatient.LastName, "''", "'")
            MsgBox clsPatient.FirstName & " " & clsPatient.LastName & " " & "is successfully saved!", vbInformation, "PatientTrac Patient Information"
        End If
    End If
    
    If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
    If Trim(txtEmergContact.Text) <> vbNullString Then
        clsPatient.EmergContact = Trim(txtEmergContact.Text)
        If clsPatient.EmergContact <> vbNullString Then clsPatient.EmergContact = Fix_Name(clsPatient.EmergContact)
        clsPatient.EmergRelation = Trim(cboRelation.Text)
        clsPatient.EmergPhone = Trim(txtEmergAC.Text) & "-" & Trim(txtEmergPh1.Text) & "-" & Trim(txtEmergPh2.Text)
        strSQL = "UPDATE PATIENT SET EMERGENCY_CONTACT = '" & clsPatient.EmergContact & "',EMERG_PHONE = '" & clsPatient.EmergPhone & "', " _
             & "EMERG_RELATION = '" & clsPatient.EmergRelation & "' WHERE PATIENT_ID = " & piPatientID
        Set rsPatient = cnPatient.Execute(strSQL)
    End If
    
    If Not IsNull(ccdDateDeath.Value) Then
        clsPatient.DateDeath = Format(ccdDateDeath.Value, "mm/dd/yyyy")
        strSQL = "UPDATE PATIENT SET OTHER_INFO3 = '" & clsPatient.DateDeath & "' WHERE PATIENT_ID =   " & piPatientID
        Set rsPatient = cnPatient.Execute(strSQL)
    End If
    If cnPatient.State = adStateOpen Then
        cnPatient.Close
        Set cnPatient = Nothing
        Set rsPatient = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Save_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Save_Data", "PatientTracSchedule.PatientInfo"
    Exit Function
    Resume
End Function

Private Function UpdatePatientInfo()
Dim strSQL As String
Dim rsPatient As New ADODB.Recordset
Dim cnPatient As New ADODB.Connection
Dim strDate As String
Dim Arr
Dim Idx%
Dim i%


    On Error GoTo Error_UpdatePatientInfo
    Screen.MousePointer = vbHourglass
    If piPatientID = 0 Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If Validate_Entries Then
        If Check_ValidPhone Then
            clsPatient.AreaCode = Trim(txtAreaCode.Text)
            clsPatient.Phone = Trim(txtPhone1.Text)
            clsPatient.Phone = Trim(clsPatient.Phone) & "-" & Trim(txtPhone2.Text)
            clsPatient.CellPhone = Trim(txtCellAreaCode.Text)
            clsPatient.CellPhone = Trim(clsPatient.CellPhone) & "-" & txtCellPhone1 & "-" & txtCellPhone2
            clsPatient.WorkPhone = txtWorkAC
            clsPatient.WorkPhone = Trim(clsPatient.WorkPhone) & "-" & txtWorkPh1 & "-" & txtWorkPh2
        End If
        clsPatient.Language = Trim(cboLanguage.Text)
        clsPatient.PrefContact = Trim(cboPrefContact.Text)
        clsPatient.Blood_Type = Trim(cboBloodType.Text)
        clsPatient.Status = Trim(cboStatus.Text)
        clsPatient.Race = Trim(cboPatientRace.Text)
        clsPatient.DL = Trim(txtDriverLicense.Text)
        clsPatient.InterNalPhone = Trim(txtInternationalPhone.Text)
        clsPatient.Email = Trim(txtEmailAddress.Text)
        clsPatient.DOB = txtMonth.Text
        clsPatient.DOB = Trim(clsPatient.DOB) & "/" & txtDay.Text
        clsPatient.DOB = Trim(clsPatient.DOB) & "/" & txtYear.Text
        If Trim(clsPatient.DOB) <> vbNullString Then clsPatient.DOB = Format(clsPatient.DOB, "mm/dd/yyyy hh:mm:ss AM/PM")
        If Not IsDate(clsPatient.DOB) Then
            mbNotSave = True
            MsgBox "Date of Birth is mandatory field!", vbInformation, "PatientTrac Patient Information"
            Exit Function
        End If
        If optMale.Value Then clsPatient.Gender = "M"
        If optFemale.Value Then clsPatient.Gender = "F"
        If clsPatient.Gender = vbNullString Then
            mbNotSave = True
            MsgBox "Gender is mandatory field!", vbInformation, "PatientTrac Patient Information"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        clsPatient.Title = Fix_Title(clsPatient.Title)
        clsPatient.FirstName = Fix_Name(clsPatient.FirstName)
        clsPatient.MiddleName = Fix_Name(clsPatient.MiddleName)
        clsPatient.LastName = Fix_Name(clsPatient.LastName)
        clsPatient.Suffix = Fix_Suffix(clsPatient.Suffix)
        clsPatient.Blood_Type = cboBloodType.Text
        clsPatient.ExtID = Literals(Trim(txtExtRef), "'")
        clsPatient.Marital_Status = cboPatMarital.Text
        If txtSS1.Text <> vbNullString Then
            clsPatient.SS = txtSS1.Text & "-" & txtSS2.Text & "-" & txtSS3.Text
        Else
            txtSS1.Text = "999"
            txtSS2.Text = "99"
            txtSS3.Text = "9999"
            clsPatient.SS = txtSS1.Text & "-" & txtSS2.Text & "-" & txtSS3.Text
        End If
    
        clsPatient.ExtID = Replace(Trim(txtExtRef.Text), "'", "''")
        If clsPatient.WorkPhone <> vbNullString Then
            clsPatient.WorkPhone = Replace(clsPatient.WorkPhone, "(", vbNullString)
            clsPatient.WorkPhone = Replace(clsPatient.WorkPhone, ")", "-")
            clsPatient.WorkPhone = Replace(clsPatient.WorkPhone, " ", "")
            clsPatient.WorkPhone = Trim(clsPatient.WorkPhone)
        End If
        strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
        clsPatient.ExtID = Replace(clsPatient.ExtID, "'", "''")
        If TextFound("p.o.", clsPatient.Address, Len("p.o."), True) Then
            clsPatient.Address = Replace(clsPatient.Address, "p.o.", "Post Office")
        End If
        PsUserID = Get_PsUserID(psConnect(0))
        clsPatient.Gender = Left(clsPatient.Gender, 1)
        clsPatient.Address2 = Fix_Name(clsPatient.Address2)
        clsPatient.Address = Fix_Name(clsPatient.Address)
        clsPatient.County = Replace(clsPatient.County, "'", "''")
        clsPatient.LastName = Replace(clsPatient.LastName, "'", "''")
        clsPatient.Address2 = Replace(clsPatient.Address2, "'", "''")
        clsPatient.Address = Replace(clsPatient.Address, "'", "''")
        clsPatient.FirstName = Replace(clsPatient.FirstName, "'", "''")
        clsPatient.Address = Replace(clsPatient.Address, "'", "''")
        clsPatient.Address2 = Replace(clsPatient.Address2, "'", "''")
        clsPatient.City = Replace(clsPatient.City, "'", "''")
        
        If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
         If Trim(txtEmergContact) <> vbNullString Then
            clsPatient.EmergContact = Trim(txtEmergContact.Text)
            clsPatient.EmergRelation = Trim(cboRelation.Text)
            clsPatient.EmergPhone = Trim(txtEmergAC.Text) & "-" & Trim(txtEmergPh1.Text) & "-" & Trim(txtEmergPh2.Text)
            strSQL = "UPDATE PATIENT  SET EMERGENCY_CONTACT = '" & clsPatient.EmergContact & "',EMERG_PHONE = '" & clsPatient.EmergContact & "', " _
                 & "EMERG_RELATION = '" & clsPatient.EmergRelation & "' WHERE PATIENT_ID = " & piPatientID
            Set rsPatient = cnPatient.Execute(strSQL)
        End If
    
        Update_Archive clsPatient.ExtID, clsPatient.Title, clsPatient.FirstName, clsPatient.MiddleName, _
                     clsPatient.LastName, clsPatient.Address, clsPatient.Address2, _
                     clsPatient.City, clsPatient.State, clsPatient.Zip, _
                     clsPatient.Zip_Ext, clsPatient.Country, Literals(clsPatient.AreaCode, "'"), _
                     Literals(clsPatient.Phone, "'"), Literals(clsPatient.CellPhone, "'"), Literals(clsPatient.DL, "'"), _
                     Literals(cboPatientRace.Text, "'"), Literals(clsPatient.Status, "'"), _
                     clsPatient.Suffix, clsPatient.DOB, Literals(clsPatient.SS, "'"), _
                     Left(clsPatient.Gender, 1), Literals(clsPatient.Email, "'"), Literals(clsPatient.InterNalPhone, "'"), _
                     clsPatient.County, clsPatient.Blood_Type, clsPatient.Marital_Status, clsPatient.WorkPhone, Literals(clsPatient.EmergPhone, "'"), clsPatient.EmergPhone, clsPatient.EmergRelation
                     
            If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
            strSQL = "UPDATE PATIENT SET PATIENT_EXT_REF = '" & clsPatient.ExtID & "', TITLE = '" & clsPatient.Title & "', FIRST_NAME  = '" & clsPatient.FirstName & "'," _
                & "MIDDLE_NAME = '" & clsPatient.MiddleName & "', LAST_NAME = '" & clsPatient.LastName & "', RACE  = '" & clsPatient.Race & "'," _
                & "PREF_CONTACT = '" & clsPatient.PrefContact & "',  STATUS  = '" & clsPatient.Status & "', COUNTY  = '" & clsPatient.County & "'," _
                & "SUFFIX = '" & clsPatient.Suffix & "', ADDRESS1  = '" & clsPatient.Address & "',  ADDRESS2 =  '" & clsPatient.Address2 & "'," _
                & "CITY = '" & clsPatient.City & "', STATE = '" & clsPatient.State & "',  PHONE = '" & clsPatient.Phone & "'," _
                & "INTERNATL_PHONE = '" & clsPatient.InterNalPhone & "', SOCIAL_SECURITY = '" & clsPatient.SS & "',  MARITAL_STATUS = '" & clsPatient.Marital_Status & "'," _
                & "AREA_CODE = '" & clsPatient.AreaCode & "', BIRTH = '" & clsPatient.DOB & "', LANGUAGE = '" & clsPatient.Language & "'," _
                & "ZIPCODE = '" & Left(clsPatient.Zip, 5) & "', ZIPCODE_EXT = '" & clsPatient.Zip_Ext & "', GENDER = '" & Left(clsPatient.Gender, 1) & "'," _
                & "CELL_PHONE =  '" & clsPatient.CellPhone & "', WORK_PHONE = '" & clsPatient.WorkPhone & "', EMAIL = '" & clsPatient.Email & "', DRV_LICENSE = '" & clsPatient.DL & " '," _
                & "USER_NAME = '" & PsUserID & "', UPDATE_DATE = '" & strDate & "', BLOOD_TYPE = '" & clsPatient.Blood_Type & "' WHERE PATIENT_ID = " & piPatientID
                Set rsPatient = cnPatient.Execute(strSQL)
        mbNotSave = False
        strSQL = "UPDATE PATIENT SET UPDATE_DATE = '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "' WHERE PATIENT_ID = " & piPatientID
        Set rsPatient = cnPatient.Execute(strSQL)
        If Trim(cboLanguage.Text) <> vbNullString Then
            clsPatient.Language = Trim(cboLanguage.Text)
            strSQL = "UPDATE PATIENT SET LANGUAGE = '" & clsPatient.Language & "' WHERE PATIENT_ID = " & piPatientID
            Set rsPatient = cnPatient.Execute(strSQL)
        End If
        If Trim(cboPrefContact.Text) <> vbNullString Then
            clsPatient.PrefContact = Trim(cboPrefContact.Text)
            strSQL = "UPDATE PATIENT SET PREF_CONTACT = '" & clsPatient.PrefContact & "' WHERE PATIENT_ID = " & piPatientID
            Set rsPatient = cnPatient.Execute(strSQL)
        End If
        If Trim(cboStatus.Text) <> vbNullString Then
            clsPatient.Status = Trim(cboStatus.Text)
            strSQL = "UPDATE PATIENT SET STATUS = '" & clsPatient.Status & "' WHERE PATIENT_ID =   " & piPatientID
            Set rsPatient = cnPatient.Execute(strSQL)
        End If
        If Not IsNull(ccdDateDeath.Value) Then
            clsPatient.DateDeath = Format(ccdDateDeath.Value, "mm/dd/yyyy")
            strSQL = "UPDATE PATIENT SET OTHER_INFO3 = '" & clsPatient.DateDeath & "' WHERE PATIENT_ID =   " & piPatientID
            Set rsPatient = cnPatient.Execute(strSQL)
        End If
        If cnPatient.State = adStateOpen Then
            cnPatient.Close
            Set cnPatient = Nothing
            Set rsPatient = Nothing
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_UpdatePatientInfo:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "UpdatePatientInfo", "PatientTracSchedule.PatientInfo"
    Exit Function
    Resume
End Function


Private Sub Form_Load()

     mFirst = True
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
    txtEmergAC.Enabled = bEnable
    txtEmergPh1.Enabled = bEnable
    txtEmergPh2.Enabled = bEnable
    cboBloodType.Enabled = bEnable
    lblPatient_id.Enabled = bEnable
    cboPatMarital.Enabled = bEnable
    txtExtRef.Enabled = bEnable
    txtWorkAC.Enabled = bEnable
    txtWorkPh1.Enabled = bEnable
    txtWorkPh2.Enabled = bEnable
    cboPrefContact.Enabled = bEnable
    cboLanguage.Enabled = bEnable
    txtFullName.Enabled = bEnable
    txtAddress.Enabled = bEnable
    cboFacility.Enabled = bEnable
    cboStatus.Enabled = bEnable
    txtEmergContact.Enabled = bEnable
    cboRelation.Enabled = bEnable
    If bEnable Then
        cboStatus.FontBold = True
        cboFacility.FontBold = True
        txtAddress.FontBold = True
        txtEmailAddress.FontBold = True
        txtFullName.FontBold = True
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
        cboBloodType.FontBold = True
        lblPatient_id.FontBold = True
        txtExtRef.FontBold = True
        cboPatMarital.FontBold = True
        txtWorkAC.FontBold = True
        txtWorkPh1.FontBold = True
        txtWorkPh2.FontBold = True
        cboPrefContact.FontBold = True
        cboLanguage.FontBold = True
    End If
    Exit Sub

    
End Sub

Private Function Insert_Patient()
Dim rsPatInsert As New ADODB.Recordset
Dim cnPatInsert As New ADODB.Connection
Dim strSQL As String
Dim Arr
Dim Idx As Integer
Dim i%
Dim strDate As String
Dim strInsert As String
Dim lngDateDiff As Long


    On Error GoTo Error_Insert_Patient
    Screen.MousePointer = vbHourglass
    If clsPatient.WorkPhone <> vbNullString Then
        clsPatient.WorkPhone = Replace(clsPatient.WorkPhone, "(", vbNullString)
        clsPatient.WorkPhone = Replace(clsPatient.WorkPhone, ")", "-")
        clsPatient.WorkPhone = Replace(clsPatient.WorkPhone, " ", "")
        clsPatient.WorkPhone = Trim(clsPatient.WorkPhone)
    End If
    strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    clsPatient.ExtID = Replace(clsPatient.ExtID, "'", "''")
    If TextFound("p.o.", clsPatient.Address, Len("p.o."), True) Then
        clsPatient.Address = Replace(clsPatient.Address, "p.o.", "Post Office")
    End If
    PsUserID = Get_PsUserID(psConnect(0))
    clsPatient.Address2 = Fix_Name(clsPatient.Address2)
    clsPatient.Address = Fix_Name(clsPatient.Address)
    clsPatient.County = Replace(clsPatient.County, "'", "''")
    clsPatient.LastName = Replace(clsPatient.LastName, "'", "''")
    clsPatient.Address2 = Replace(clsPatient.Address2, "'", "''")
    clsPatient.Address = Replace(clsPatient.Address, "'", "''")
    clsPatient.FirstName = Replace(clsPatient.FirstName, "'", "''")
    clsPatient.Address = Replace(clsPatient.Address, "'", "''")
    clsPatient.Address2 = Replace(clsPatient.Address2, "'", "''")
    clsPatient.City = Replace(clsPatient.City, "'", "''")
    If cnPatInsert.State = adStateClosed Then cnPatInsert.open psConnect(0)
    strSQL = "INSERT INTO PATIENT (PATIENT_EXT_REF, TITLE, FIRST_NAME, MIDDLE_NAME, LAST_NAME,RACE, PREF_CONTACT, STATUS, COUNTY, " _
        & "SUFFIX, ADDRESS1, ADDRESS2, CITY,STATE, PHONE, INTERNATL_PHONE, SOCIAL_SECURITY, MARITAL_STATUS, AREA_CODE, BIRTH, LANGUAGE, " _
        & "ZIPCODE, ZIPCODE_EXT, GENDER, CELL_PHONE, WORK_PHONE, EMAIL, DRV_LICENSE, USER_NAME, INSERT_DATE, PURGE, BLOOD_TYPE) values ('" & clsPatient.ExtID & "','" & clsPatient.Title & "','" & clsPatient.FirstName & "','" _
        & clsPatient.MiddleName & "','" & clsPatient.LastName & "','" & clsPatient.Race & "','" & clsPatient.PrefContact & "','" & clsPatient.Status & "','" & clsPatient.County & "','" _
        & clsPatient.Suffix & "','" & clsPatient.Address & "','" & clsPatient.Address2 & "','" _
        & clsPatient.City & "','" & clsPatient.State & "','" & Trim(clsPatient.Phone) & "','" & Trim(clsPatient.InterNalPhone) & "','" & clsPatient.SS & "','" _
        & clsPatient.Marital_Status & "', '" & clsPatient.AreaCode & "','" & clsPatient.DOB & "','" & clsPatient.Language & "','" & Left(clsPatient.Zip, 5) & "','" _
        & clsPatient.Zip_Ext & "','" & Left(clsPatient.Gender, 1) & "','" & clsPatient.CellPhone & "','" & clsPatient.WorkPhone & "','" & clsPatient.Email & "','" _
        & clsPatient.DL & "', '" & PsUserID & "','" & strDate & "',0,'" & clsPatient.Blood_Type & "')"
        Set rsPatInsert = cnPatInsert.Execute(strSQL)
    strSQL = "SELECT * FROM PATIENT WHERE LAST_NAME = '" & clsPatient.LastName & "' and INSERT_DATE = '" & strDate & "'"
    Set rsPatInsert = cnPatInsert.Execute(strSQL)
    If Not rsPatInsert.EOF Then
        piPatientID = rsPatInsert.Fields("PATIENT_ID")
        clsPatient.IntId = piPatientID
    End If
    If cnPatInsert.State = adStateOpen Then
        cnPatInsert.Close
        Set cnPatInsert = Nothing
        Set rsPatInsert = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Insert_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Insert_Patient", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function

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
    mFirst = False
    Exit Sub
End Sub



Private Sub Set_Patient()
Dim strSQL As String
Dim strEnctrStart As String
Dim cnPatient As New ADODB.Connection
Dim rsPatient As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim i%
Dim Arr
Dim fso As New FileSystemObject

    On Error GoTo Error_SetPatient
    Screen.MousePointer = vbHourglass
    Clear_Fields
    Dump_PatientStrData
    Get_Partner piPatientID
    If cboRelation.ListCount <= 0 Then Get_RelationshipTerms cboRelation
    If cboLanguage.ListCount <= 0 Then GetTerms cboLanguage, "sd_Patient_Language"
    If cboPatientRace.ListCount <= 0 Then LoadAttrValidValue psConnect(1), cboPatientRace, "sd_Patient_Race"
    If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
    strSQL = "SELECT * FROM PATIENT WHERE PATIENT_ID = " & piPatientID
    Set rsPatient = cnPatient.Execute(strSQL)
    If Not rsPatient.EOF Then
        Set Flds = rsPatient.Fields
        For Each Fld In Flds
            If Not IsNull(Fld.Value) Then
                If Trim(Fld.Value) <> vbNullString Then
                    Select Case UCase(Fld.Name)
                        Case "FIRST_NAME"
                            clsPatient.FirstName = Fld.Value
                        Case "MIDDLE_NAME"
                            clsPatient.MiddleName = IsProperCase(Fld.Value)
                        Case "LAST_NAME"
                            clsPatient.LastName = Fld.Value
                            If TextFound("''", clsPatient.LastName, 2, True) Then
                                clsPatient.LastName = Replace(clsPatient.LastName, "''", "'")
                                clsPatient.LastName = IsProperCase(clsPatient.LastName)
                            Else
                                clsPatient.LastName = IsProperCase(clsPatient.LastName)
                            End If
                        Case "GENDER"
                            clsPatient.Gender = IsProperCase(Fld.Value)
                            If clsPatient.Gender = "F" Or clsPatient.Gender = "Female" Then
                                clsPatient.Gender = "Female"
                                clsPatient.Person = "She"
                                optFemale.Value = True
                            Else
                                optMale.Value = True
                                clsPatient.Gender = "Male"
                                clsPatient.Person = "He"
                            End If
                        Case "TITLE"
                            clsPatient.Title = IsProperCase(Fld.Value)
                            If clsPatient.Title = vbNullString Then
                                If clsPatient.Gender = "M" Or clsPatient.Gender = "Male" Then
                                    clsPatient.Title = "Mr."
                                Else
                                    If clsPatient.Marital_Status = "married" Or clsPatient.Marital_Status = UCase("Married") And _
                                        clsPatient.Gender = "F" Then
                                        clsPatient.Title = "Mrs."
                                    Else
                                        clsPatient.Title = "Ms."
                                    End If
                                End If
                            End If
                        Case "MARITAL_STATUS"
                            clsPatient.Marital_Status = IsProperCase(Fld.Value)
                        Case "SUFFIX"
                            clsPatient.Suffix = IsProperCase(Fld.Value)
                            clsPatient.Suffix = Replace(clsPatient.Suffix, ",", "")
                        Case "BIRTH"
                                clsPatient.DOB = Fld.Value
                                If clsPatient.DOB <> vbNullString Then
                                    If clsEncounter.StartDate <> vbNullString Then
                                        strEnctrStart = Format(clsEncounter.StartDate, "mm/dd/yyyy")
                                    Else
                                        strEnctrStart = Format(Now, "mm/dd/yyyy")
                                    End If
                                    clsPatient.Age = Int(DateDiff("y", clsPatient.DOB, strEnctrStart) / 365.25)
                                End If
                        Case "PATIENT_EXT_REF"
                            clsPatient.ExtID = Fld.Value
                            txtExtRef.Text = clsPatient.ExtID
                        Case "PATIENT_ID"
                            clsPatient.IntId = Fld.Value
                        Case "SOCIAL_SECURITY"
                            clsPatient.SS = Fld.Value
                        Case "ADDRESS1"
                            clsPatient.Address = Fld.Value
                        Case "ADDRESS2"
                            clsPatient.Address2 = IsProperCase(Fld.Value)
                        Case "CITY"
                            clsPatient.City = IsProperCase(Fld.Value)
                            clsPatient.City = Replace(clsPatient.City, " - ", "-")
                            clsPatient.City = Trim(clsPatient.City)
                        Case "STATE"
                            clsPatient.State = IsProperCase(Fld.Value)
                        Case "ZIPCODE"
                            clsPatient.Zip = Fld.Value
                         Case "ZIPCODE_EXT"
                            clsPatient.Zip_Ext = Fld.Value
                         Case "ROOM_NUMBER"
                            clsPatient.RoomNumber = Fld.Value
                        Case "COUNTRY"
                            clsPatient.Country = Fld.Value
                        Case "CELL_PHONE"
                            clsPatient.CellPhone = Fld.Value
                        Case "EMAIL"
                            clsPatient.Email = LCase(Fld.Value)
                            txtEmailAddress.Text = clsPatient.Email
                        Case "COUNTY"
                            clsPatient.County = IsProperCase(Fld.Value)
                        Case "PATIENT_TYPE"
                            clsPatient.PatientType = LCase(Fld.Value)
                        Case "DRV_LICENSE"
                            clsPatient.DL = LCase(Fld.Value)
                            txtDriverLicense.Text = clsPatient.DL
                        Case "WORK_PHONE"
                            clsPatient.WorkPhone = LCase(Fld.Value)
                        Case "INTERNATL_PHONE"
                            clsPatient.InterNalPhone = LCase(Fld.Value)
                            txtInternationalPhone.Text = clsPatient.InterNalPhone
                        Case "PHONE"
                            clsPatient.Phone = LCase(Fld.Value)
                        Case "AREA_CODE"
                            clsPatient.AreaCode = LCase(Fld.Value)
                        Case "EMERG_PHONE"
                            clsPatient.EmergPhone = LCase(Fld.Value)
                            clsEmerg.Phone = clsPatient.EmergPhone
                            Arr = Split(clsPatient.EmergPhone, "-")
                            If UBound(Arr) > 0 Then
                                If Arr(0) <> vbNullString Then txtEmergAC.Text = Trim(Arr(0))
                                If Arr(1) <> vbNullString Then txtEmergPh1.Text = Arr(1)
                            ElseIf UBound(Arr) > 1 Then
                                If Arr(2) <> vbNullString Then txtEmergPh2 = Arr(2)
                            End If
                        Case "EMERG_RELATION"
                            clsPatient.EmergRelation = LCase(Fld.Value)
                            clsEmerg.Relation = clsPatient.EmergRelation
                            For i = 0 To cboRelation.ListCount - 1
                                If IsProperCase(cboRelation.List(i)) = IsProperCase(clsEmerg.Relation) Then
                                    cboRelation.Text = cboRelation.List(i)
                                    Exit For
                                End If
                            Next i
                        Case "EMERGENCY_CONTACT"
                            clsPatient.EmergContact = LCase(Fld.Value)
                            clsEmerg.Name = clsPatient.EmergContact
                            txtEmergContact.Text = clsPatient.EmergContact
                        Case "STATUS"
                            clsPatient.Status = Fld.Value
                            For i = 0 To cboStatus.ListCount - 1
                                If cboStatus.List(i) = clsPatient.Status Then
                                    cboStatus.Text = cboStatus.List(i)
                                    Exit For
                                End If
                            Next i
                            If UCase(clsPatient.Status) = "DECEASED" Then fraDOD.Visible = True
                        Case "PURGE"
                            If Fld.Value = True Or Fld.Value = "1" Then
                                For i = 0 To cboStatus.ListCount - 1
                                    If cboStatus.List(i) = "INACTIVE" Then
                                        cboStatus.Text = cboStatus.List(i)
                                        clsPatient.Status = "Inactive Acct"
                                        Exit For
                                    End If
                                Next i
                            End If
                        Case "ROOM_PHONE"
                            clsPatient.RoomPhone = LCase(Fld.Value)
                        Case "RACE"
                            clsPatient.Race = LCase(Fld.Value)
                            cboPatientRace.AddItem "white"
                            For i = 0 To cboPatientRace.ListCount - 1
                                If TextFound(clsPatient.Race, cboPatientRace.List(i), Len(clsPatient.Race), True) Then
                                    cboPatientRace.Text = cboPatientRace.List(i)
                                    Exit For
                                End If
                            Next i
                        Case "BLOOD_TYPE"
                            clsPatient.Blood_Type = LCase(Fld.Value)
                            For i = 0 To cboBloodType.ListCount - 1
                                If cboBloodType.List(i) = UCase(clsPatient.Blood_Type) Then
                                    cboBloodType.Text = cboBloodType.List(i)
                                    Exit For
                                End If
                            Next i
                        Case "PREF_CONTACT"
                            clsPatient.PrefContact = Fld.Value
                            For i = 0 To cboPrefContact.ListCount - 1
                                If cboPrefContact.List(i) = clsPatient.PrefContact Then
                                    cboPrefContact.Text = cboPrefContact.List(i)
                                    Exit For
                                End If
                            Next i
                        Case "OTHER_INFO3"
                            clsPatient.DateDeath = Format(Fld.Value, "mm/dd/yyyy")
                            ccdDateDeath.Enabled = True
                            ccdDateDeath.Value = clsPatient.DateDeath
                            fraDOD.Visible = True
                        Case "LANGUAGE"
                            clsPatient.Language = IsProperCase(Fld.Value)
                            For i = 0 To cboLanguage.ListCount - 1
                                If TextFound(clsPatient.Language, cboLanguage.List(i), Len(clsPatient.Language), True) Then
                                    cboLanguage.Text = cboLanguage.List(i)
                                    Exit For
                                End If
                            Next
                    End Select
                End If
            End If
        Next
    End If
    If Trim(clsPatient.DateDeath) = vbNullString Then
        ccdDateDeath.Enabled = False
        fraDOD.Visible = False
    End If
    If clsPatient.Phone <> vbNullString Then
        If clsPatient.AreaCode <> vbNullString Then txtAreaCode.Text = Format(clsPatient.AreaCode, "000")
        Arr = Split(clsPatient.Phone, "-")
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
    If Trim(clsPatient.WorkPhone) <> vbNullString Then
        Arr = Split(clsPatient.WorkPhone, "-")
        If UBound(Arr) > 0 Then
            If Trim(Arr(0)) <> vbNullString Then txtWorkAC.Text = Format(Arr(0), "000")
            If Trim(Arr(1)) <> vbNullString Then txtWorkPh1.Text = Format(Arr(1), "000")
            If UBound(Arr) > 1 Then
                If Trim(Arr(2)) <> vbNullString Then txtWorkPh2.Text = Format(Arr(2), "0000")
            End If
        Else
            txtWorkAC.Text = vbNullString
            txtWorkPh1.Text = vbNullString
            txtWorkPh2.Text = vbNullString
        End If
    End If
    If clsPatient.CellPhone <> vbNullString Then
        Arr = vbNullString
        Arr = Split(clsPatient.CellPhone, "-")
        If UBound(Arr) > 0 Then
            txtCellAreaCode.Text = Format(Arr(0), "000")
            txtCellPhone1.Text = Format(Arr(1), "000")
            If UBound(Arr) > 1 Then
                txtCellPhone2.Text = Format(Arr(2), "0000")
            End If
        End If
    Else
        txtCellAreaCode.Text = vbNullString
        txtCellPhone1.Text = vbNullString
        txtCellPhone2.Text = vbNullString
    End If

    If IsDate(clsPatient.DOB) Then
        clsPatient.DOB = Format(clsPatient.DOB, "mm/dd/yyyy")
        txtMonth.Text = Left(clsPatient.DOB, 2)
        txtDay.Text = Mid(clsPatient.DOB, 4, 2)
        txtYear.Text = Mid(clsPatient.DOB, 7, 4)
    End If
    txtAge.Text = Trim(clsPatient.Age)
    If Val(txtAge.Text) <= 18 And mstrPractice = "Psychiatry" Then
        lblPartnerName.Caption = "Parent or Legal Guardian"
        lblPartTitle.Caption = "PARENT OR LEGAL GUARDIAN"
    End If

    If Trim(clsPatient.SS) <> vbNullString Then
       Arr = Split(clsPatient.SS, "-")
       If UBound(Arr) > 0 Then
            If Trim(Arr(0)) <> vbNullString Then txtSS1.Text = Format(Arr(0), "000")
            If Trim(Arr(1)) <> vbNullString Then txtSS2.Text = Format(Arr(1), "00")
            If UBound(Arr) > 1 Then
                If Trim(Arr(2)) <> vbNullString Then txtSS3.Text = Format(Arr(2), "0000")
            End If
       End If
    End If
    If clsPatient.Marital_Status <> vbNullString Then
        For i = 0 To cboPatMarital.ListCount - 1
            If IsProperCase(cboPatMarital.List(i)) = clsPatient.Marital_Status Then
                cboPatMarital.Text = cboPatMarital.List(i)
                Exit For
            End If
        Next i
    End If
    If cnPatient.State = adStateOpen Then
        cnPatient.Close
        Set cnPatient = Nothing
        Set rsPatient = Nothing
        Set Flds = Nothing
    End If
    clsPatient.Suffix = Fix_Suffix(clsPatient.Suffix)
    clsPatient.Address = Fix_Name(clsPatient.Address)
    clsPatient.FirstName = Fix_Name(clsPatient.FirstName)
    clsPatient.LastName = Fix_Name(clsPatient.LastName)
    clsPatient.Title = Fix_Title(clsPatient.Title)
    clsPatient.City = Fix_Name(clsPatient.City)
    clsPatient.City = Replace(clsPatient.City, " - ", "-")
    clsPatient.City = Trim(clsPatient.City)
    If Trim(clsPatient.Suffix) <> vbNullString Then
        clsPatient.Name = clsPatient.FirstName & " " & clsPatient.LastName & ", " & clsPatient.Suffix
    Else
        clsPatient.Name = clsPatient.FirstName & " " & clsPatient.LastName
    End If
    If Trim(clsPatient.City) <> vbNullString Then
         clsPatient.CityStateZip = Trim(clsPatient.City)
     End If
    If Trim(clsPatient.State) <> vbNullString Then
         clsPatient.CityStateZip = clsPatient.CityStateZip & ", " & Trim(clsPatient.State)
     End If
    If Trim(clsPatient.Zip) <> vbNullString Then
         clsPatient.CityStateZip = clsPatient.CityStateZip & " " & Trim(clsPatient.Zip)
     End If
     If clsPatient.Zip_Ext <> vbNullString Then clsPatient.CityStateZip = clsPatient.CityStateZip & "-" & clsPatient.Zip_Ext
    If clsPatient.Title = vbNullString Then
        If clsPatient.Gender = "M" Or clsPatient.Gender = "Male" Then
            clsPatient.Title = "Mr."
        Else
            If clsPatient.Marital_Status = "married" Or clsPatient.Marital_Status = UCase("Married") And _
                clsPatient.Gender = "F" Then
                clsPatient.Title = "Mrs."
            Else
                clsPatient.Title = "Ms."
            End If
        End If
    End If
    If Trim(clsPatient.Title) <> vbNullString Then clsPatient.Name = Trim(clsPatient.Title) & " " & Trim(clsPatient.Name)
    If Trim(clsPatient.Name) <> vbNullString Then txtFullName.Text = Trim(clsPatient.Name)
    If Trim(clsPatient.Address) <> vbNullString Then txtAddress.Text = Trim(clsPatient.Address)
    If Trim(clsPatient.Address2) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & ", " & Trim(clsPatient.Address2)
    If Trim(clsPatient.City) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & vbNewLine & Trim(clsPatient.City)
    If Trim(clsPatient.State) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & " " & Trim(clsPatient.State)
    If Trim(clsPatient.Zip) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & " " & Trim(clsPatient.Zip)
    If Trim(clsPatient.Country) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & vbNewLine & Trim(clsPatient.Country)
    If Trim(clsPatient.County) <> vbNullString Then txtAddress.Text = Trim(txtAddress.Text) & vbNewLine & Trim(clsPatient.County)
    If piPatientID <> 0 Then lblPatient_id.Caption = piPatientID
    mbooByPass = True
    Get_Partner piPatientID
    txtParner.Text = clsPartner.Name
    Get_EmergContact piPatientID
    txtEmergContact.Text = clsEmerg.Name
    If clsEmerg.Phone <> vbNullString Then
        Arr = Split(clsEmerg.Phone, "-")
        If UBound(Arr) > 0 Then
            For i = LBound(Arr) To UBound(Arr)
                If UBound(Arr) > 1 Then
                    If i = 0 Then txtEmergAC.Text = Arr(i)
                    If i = 1 Then txtEmergPh1.Text = Arr(i)
                    If i = 2 Then txtEmergPh2.Text = Arr(i)
                End If
            Next i
        End If
    End If
    If clsEmerg.Relation <> vbNullString Then
        If cboRelation.ListCount <= 0 Then Get_RelationshipTerms cboRelation
        For i = 0 To cboRelation.ListCount - 1
            If cboRelation.List(i) = Trim(clsEmerg.Relation) Then
                cboRelation.Text = cboRelation.List(i)
                Exit For
            End If
        Next i
    End If
    If piPatientID > 0 Then
        fraPic.ZOrder 0
        If Right(psImageFilePath, 1) <> "\" Then
            psImageFilePath = psImageFilePath & "\"
        End If
        If fso.FileExists(psImageFilePath & piPatientID & "\Photo.jpg") Then
           psPhoto = psImageFilePath & piPatientID & "\Photo.jpg"
           imgPatient.Picture = LoadPicture(psPhoto)
           lblPhoto.Visible = False
           imgPatient.Visible = True
        Else
          imgPatient.Visible = False
          lblPhoto.Visible = True
        End If
    End If
    Get_Bill_Facility
    Load_Facilities
    If cboFacility.ListCount > 0 Then
        For i = 0 To cboFacility.ListCount - 1
            If cboFacility.ItemData(i) = clsFacility.BillFacility Then
                mbooByPass = True
                cboFacility.Text = cboFacility.List(i)
                mbooByPass = False
                Exit For
            End If
        Next i
    End If
    cmdNew.Visible = False
    cmdUpdate.Visible = True
    cmdClose.Visible = False
    optMale.Visible = True
    optFemale.Visible = True
    bEnable = False
    Screen_Action
    Screen.MousePointer = vbDefault
    Exit Sub
Error_SetPatient:
    Screen.MousePointer = vbHourglass
    Write_Error Err.Description, Err.Number, strSQL, "Set_Patient", "PatientTrac.PatientInfo"
    Exit Sub
    Resume
End Sub
Private Sub txtAddress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim strAddress As String
    If mbooByPass Then mbooByPass = False: Exit Sub
    txtAddress.Text = vbNullString
    frmPatientAdress.Show vbModal
    If clsPatient.Address <> vbNullString Then
        If Trim(clsPatient.Address) <> vbNullString Then strAddress = Trim(clsPatient.Address)
        If Trim(clsPatient.Address2) <> vbNullString Then strAddress = Trim(strAddress) & ", " & Trim(clsPatient.Address2)
        If Trim(clsPatient.City) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(clsPatient.City)
        If Trim(clsPatient.State) <> vbNullString Then strAddress = Trim(strAddress) & ", " & Trim(clsPatient.State)
        If Trim(clsPatient.Zip) <> vbNullString Then strAddress = Trim(strAddress) & " " & Trim(clsPatient.Zip)
        If Trim(clsPatient.Country) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(clsPatient.Country)
        If Trim(clsPatient.County) <> vbNullString Then strAddress = Trim(strAddress) & vbNewLine & Trim(clsPatient.County)
        txtAddress.Text = Trim(strAddress)
    End If
    Exit Sub
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


Private Sub txtEmergContact_Click()
Dim lngPatID As Long
Dim Arr
Dim i%
Dim Idx%

    On Error GoTo Error_Load_Emergency
    If mNewPatient = True Then
        If txtFullName.Text <> vbNullString And clsPatient.LastName <> vbNullString Then
            lngPatID = Save_Data
            Get_Patient lngPatID
            mNewPatient = False
        Else
            MsgBox "Please complete patient name first!, PatientTrac saves this new patient first before creating an emergency contact.", vbInformation, "PatientTrac Patient Data"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        lngPatID = piPatientID
    End If
    If lngPatID > 0 Then
        piPatientID = lngPatID
        mstrPatCtl = "EmergencyContact"
        mstrPatOcx = "FN_PP_PS"
        frmPatients.piPatientID = piPatientID
        modSpecific.piPatientID = piPatientID
        frmLoadControl.piSelPatientID = piPatientID
        frmLoadControl.Show vbModal
        Get_EmergContact lngPatID
        txtEmergContact.Text = clsEmerg.Name
        If clsEmerg.Phone <> vbNullString Then
            Arr = Split(clsEmerg.Phone, "-")
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
        If clsEmerg.Relation <> vbNullString Then
            If cboRelation.ListCount <= 0 Then Get_RelationshipTerms cboRelation
            For Idx = 0 To cboRelation.ListCount - 1
                If IsProperCase(cboRelation.List(Idx)) = IsProperCase(clsEmerg.Relation) Then
                    cboRelation.Text = cboRelation.List(Idx)
                    Exit For
                End If
            Next Idx
        End If
        mNewPatient = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Emergency:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_Emergency", App.EXEName & "." & "PatientInfo"
    Exit Sub
    Resume
End Sub

Private Sub txtEmergContact_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lngPatID As Long
Dim Arr
Dim i%
Dim Idx%

    If mNewPatient = True Then
        lngPatID = Save_Data
    Else
        lngPatID = piPatientID
    End If
    If lngPatID > 0 Then
        piPatientID = lngPatID
        mstrPatCtl = "EmergencyContact"
        mstrPatOcx = "FN_PP_PS"
        frmLoadControl.Show vbModal
        Get_EmergContact lngPatID
        txtEmergContact.Text = clsEmerg.Name
        If clsEmerg.Phone <> vbNullString Then
            Arr = Split(clsEmerg.Phone, "-")
            If UBound(Arr) > 0 Then
                For Idx = LBound(Arr) To UBound(Arr)
                    If UBound(Arr) > 1 Then
                        If Idx = 0 Then txtEmergAC.Text = Arr(i)
                        If Idx = 1 Then txtEmergPh1.Text = Arr(i)
                        If Idx = 2 Then txtEmergPh2.Text = Arr(i)
                    End If
                Next Idx
            End If
        End If
        If clsEmerg.Relation <> vbNullString Then
            If cboRelation.ListCount <= 0 Then Get_RelationshipTerms cboRelation
            For Idx = 0 To cboRelation.ListCount - 1
                If cboRelation.List(Idx) = Trim(clsEmerg.Relation) Then
                    cboRelation.Text = cboRelation.List(Idx)
                    Exit For
                End If
            Next Idx
        End If
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

    If mbooByPass Then mbooByPass = False: Exit Sub
    
    If mNewPatient Then
        Dump_PatientStrData
    End If
    txtFullName.Text = vbNullString
    frmPatientName.Show vbModal
    mbooByPass = True
    If Trim(clsPatient.Title) <> vbNullString Then txtFullName.Text = Trim(clsPatient.Title)
    If Trim(clsPatient.FirstName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.FirstName)
    If Trim(clsPatient.MiddleName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.MiddleName)
    If Trim(clsPatient.LastName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.LastName)
    If Trim(clsPatient.Suffix) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & ", " & Trim(clsPatient.Suffix)
    If Right(txtFullName.Text, 1) = ", " And Trim(clsPatient.Suffix) = vbNullString Then txtFullName.Text = Replace(txtFullName.Text, ", ", "")
    txtFullName.Text = Trim((txtFullName.Text))
    txtFullName.Text = StrConv(txtFullName.Text, vbProperCase)
    mbooByPass = False
    Exit Sub
End Sub

Private Sub txtFullName_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strName As String
    If mbooByPass Then mbooByPass = False: Exit Sub
    If mNewPatient Then
        Dump_PatientStrData
    End If
    txtFullName.Text = vbNullString
    frmPatientName.Show vbModal
    mbooByPass = True
    If Trim(clsPatient.Title) <> vbNullString Then txtFullName.Text = Trim(clsPatient.Title)
    If Trim(clsPatient.FirstName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.FirstName)
    If Trim(clsPatient.MiddleName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.MiddleName)
    If Trim(clsPatient.LastName) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & " " & Trim(clsPatient.LastName)
    If Trim(clsPatient.Suffix) <> vbNullString Then txtFullName.Text = Trim(txtFullName.Text) & ", " & Trim(clsPatient.Suffix)
    If Right(txtFullName.Text, 1) = ", " And Trim(clsPatient.Suffix) = vbNullString Then txtFullName.Text = Replace(txtFullName.Text, ", ", "")
    txtFullName.Text = Trim((txtFullName.Text))
    txtFullName.Text = StrConv(txtFullName.Text, vbProperCase)
    mbooByPass = False
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

    On Error GoTo Error_Load_Partner
    Screen.MousePointer = vbHourglass
    If mNewPatient = True Then
        If txtFullName.Text <> vbNullString And Trim(clsPatient.LastName) <> vbNullString Then
            lngPatID = Save_Data
            Get_Patient lngPatID
            mNewPatient = False
        Else
            MsgBox "Please complete patient name first!, PatientTrac saves this new patient first before creating Partner or Guardian.", vbInformation, "PatientTrac Patient Data"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        lngPatID = piPatientID
    End If
    If lngPatID > 0 Then
        piPatientID = lngPatID
        mstrPatOcx = "FN_PP_PS"
        mstrPatCtl = "Partner_Guardian"
        frmLoadControl.Show vbModal
        Get_Partner lngPatID
        txtParner.Text = clsPartner.Name
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Partner:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_Partner", App.EXEName & "." & "PatientInfo"
    Exit Sub
    Resume
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
        If Val(txtAge.Text) <= 18 And mstrPractice = "Psychiatry" And clsPartner.Name = vbNullString Then
            If txtParner.Text = vbNullString Then MsgBox "Please Complete Patient's Parent or Legal Guardian Information!", vbInformation
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
End Sub
