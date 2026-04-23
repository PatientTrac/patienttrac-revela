VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{6514F5A0-641C-11D2-9FD0-0020AF131A57}#3.0#0"; "fpFlp30.ocx"
Begin VB.Form frmSchedPatient 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Schedule Patient Appointments"
   ClientHeight    =   9315
   ClientLeft      =   4635
   ClientTop       =   2850
   ClientWidth     =   13365
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchedPatient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleMode       =   0  'User
   ScaleWidth      =   13365
   Visible         =   0   'False
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
      TabIndex        =   72
      Top             =   1185
      Visible         =   0   'False
      Width           =   3825
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
      ItemData        =   "frmSchedPatient.frx":058A
      Left            =   2040
      List            =   "frmSchedPatient.frx":059D
      Style           =   2  'Dropdown List
      TabIndex        =   67
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
      Picture         =   "frmSchedPatient.frx":05D0
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Update Patient"
      CaptionAlignment=   5
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedPatient.frx":0B6A
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
      Picture         =   "frmSchedPatient.frx":0EBC
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmSchedPatient.frx":1796
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
      Top             =   11070
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
      Left            =   90
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
      Left            =   90
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1155
      Width           =   4875
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   345
      Left            =   6773
      TabIndex        =   11
      Top             =   8865
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
      Picture         =   "frmSchedPatient.frx":23E8
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedPatient.frx":2DFA
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdApply 
      Height          =   345
      Left            =   5603
      TabIndex        =   10
      Top             =   8865
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
      Picture         =   "frmSchedPatient.frx":314C
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "&Apply"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedPatient.frx":34E6
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdNewPatient 
      Height          =   345
      Left            =   14400
      TabIndex        =   75
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
      Picture         =   "frmSchedPatient.frx":3838
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "New Patient"
      CaptionAlignment=   5
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSchedPatient.frx":3DD2
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
      Height          =   6525
      Left            =   60
      TabIndex        =   14
      Top             =   2310
      Width           =   13275
      Begin VB.ComboBox cboTime 
         Height          =   315
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   2340
         Width           =   1455
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
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   70
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
         Top             =   3465
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
            Picture         =   "frmSchedPatient.frx":4124
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close "
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmSchedPatient.frx":4B36
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
            Picture         =   "frmSchedPatient.frx":4E88
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
            PictureDisabled =   "frmSchedPatient.frx":5422
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
         Height          =   825
         Left            =   8925
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
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
         ItemData        =   "frmSchedPatient.frx":5774
         Left            =   8925
         List            =   "frmSchedPatient.frx":5776
         Style           =   2  'Dropdown List
         TabIndex        =   60
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
         Left            =   1830
         TabIndex        =   58
         Top             =   5850
         Width           =   2145
      End
      Begin VB.ListBox lstPurVisit 
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
         Height          =   1020
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   56
         Top             =   4740
         Width           =   6225
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
         ItemData        =   "frmSchedPatient.frx":5778
         Left            =   2595
         List            =   "frmSchedPatient.frx":577A
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
         Format          =   76873731
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
         CalendarTrailingForeColor=   8421504
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyy"
         DateIsNull      =   -1  'True
         Format          =   76873731
         CurrentDate     =   38792
      End
      Begin PT_XP_Button.PT_XP cmdExit 
         Height          =   345
         Left            =   6090
         TabIndex        =   12
         Top             =   6015
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
         Picture         =   "frmSchedPatient.frx":577C
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close Me"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient.frx":618E
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
         ItemData        =   "frmSchedPatient.frx":64E0
         Left            =   2010
         List            =   "frmSchedPatient.frx":64E2
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
         Height          =   1080
         Left            =   8925
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   4950
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
         ItemData        =   "frmSchedPatient.frx":64E4
         Left            =   9420
         List            =   "frmSchedPatient.frx":64E6
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
         Picture         =   "frmSchedPatient.frx":64E8
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmSchedPatient.frx":6A82
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
         Picture         =   "frmSchedPatient.frx":6DD4
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmSchedPatient.frx":736E
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
         Picture         =   "frmSchedPatient.frx":76C0
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmSchedPatient.frx":7C5A
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdChange 
         Height          =   315
         Left            =   11820
         TabIndex        =   61
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
         Picture         =   "frmSchedPatient.frx":7FAC
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
         PictureDisabled =   "frmSchedPatient.frx":8546
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdChangeUser 
         Height          =   465
         Left            =   11475
         TabIndex        =   69
         Top             =   210
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
         Picture         =   "frmSchedPatient.frx":8898
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Manage Appointment"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient.frx":8E32
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRepeat 
         Height          =   345
         Left            =   12750
         TabIndex        =   71
         ToolTipText     =   "Change Repeat Appts"
         Top             =   1170
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
         Picture         =   "frmSchedPatient.frx":9184
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Change Appt"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient.frx":971E
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdShowAll 
         Height          =   345
         Left            =   2130
         TabIndex        =   74
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
         Picture         =   "frmSchedPatient.frx":9A70
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Change Appt"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient.frx":A00A
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAcct 
         Height          =   450
         Left            =   9945
         TabIndex        =   80
         Top             =   210
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   794
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
         Picture         =   "frmSchedPatient.frx":A35C
         AppearanceThemes=   3
         BorderDrawEvent =   1
         TransparentColor=   14737632
         Caption         =   "Account Balance"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient.frx":A8F6
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
         TabIndex        =   65
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
         TabIndex        =   66
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         Left            =   135
         TabIndex        =   59
         Top             =   5910
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
         TabIndex        =   57
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
         Top             =   4950
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
         Left            =   135
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
         Left            =   4080
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
      Height          =   5880
      Left            =   60
      TabIndex        =   81
      Top             =   3000
      Visible         =   0   'False
      Width           =   13215
      Begin LpADOLib.fpComboAdo fpTerm 
         Height          =   315
         Left            =   4335
         TabIndex        =   82
         Top             =   6150
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
         ColDesigner     =   "frmSchedPatient.frx":AC48
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
         Picture         =   "frmSchedPatient.frx":AF0B
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Remove Appointment"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient.frx":B4A5
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdAppointChange 
         Height          =   3945
         Left            =   135
         TabIndex        =   84
         Top             =   1050
         Width           =   12900
         _Version        =   458752
         _ExtentX        =   22754
         _ExtentY        =   6959
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
         SpreadDesigner  =   "frmSchedPatient.frx":B7F7
         UserResize      =   0
         Appearance      =   1
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
         Picture         =   "frmSchedPatient.frx":BABE
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Update Appt Row"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient.frx":C058
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdFrameClose 
         Height          =   375
         Left            =   6165
         TabIndex        =   86
         Top             =   5250
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
         Picture         =   "frmSchedPatient.frx":C3AA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient.frx":CDBC
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
         Format          =   76873731
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
         Picture         =   "frmSchedPatient.frx":D10E
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Print"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSchedPatient.frx":D6A8
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
      TabIndex        =   73
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
      Left            =   90
      TabIndex        =   68
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
      Left            =   90
      TabIndex        =   32
      Top             =   450
      Width           =   1275
   End
End
Attribute VB_Name = "frmSchedPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngtop As Long
Public mlngleft As Long
Public mstrTerms As String
Public mlngPatientID As Long
Private mbooByPass As Boolean
Public mstrListName As String
Public mFirst As Boolean
Public mbooCaption As Boolean
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
Public mbooNoMessage As Boolean
Private fpDeltaLeft As Long
Private fpDeltaTop As Long
Private flngDeltaX  As Long
Private flngDeltaY  As Long
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long
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
    Resume
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
Private Function Get_Grid_Status() As String
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim strTerms As String

    On Error GoTo Error_Get_Grid_Status
    Screen.MousePointer = vbHourglass
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Get_Attr_Valid_Value_SubDom_Sequence 'sd_Appointment_Status' "
    Set rsTerms = cnTerms.Execute(strSQL)
    With rsTerms
        Do Until .EOF
            strTerms = vbNullString
            strTerms = IsProperCase(Trim(!ATTRBTE_VALID_VALUE))
            .MoveNext
            If Get_Grid_Status = vbNullString Then
                Get_Grid_Status = strTerms
            ElseIf Not TextFound(strTerms, Get_Grid_Status, Len(strTerms), True) Then
                Get_Grid_Status = Get_Grid_Status & Chr(9) & strTerms
            End If
        Loop
    End With
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
    Resume
End Function

Private Function Get_GridDuration() As String
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset

    On Error GoTo Error_Get_GridDuration
    Screen.MousePointer = vbHourglass
    Get_GridDuration = vbNullString
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(1)
    strSQL = "Get_Appt_Time"
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
    Resume
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
    If Trim(vstrType) <> vbNullString Then
        strSQL = "Get_Appt_Types '" & vstrType & "%' "
    Else
        strSQL = "Get_Appt_Types_Facility_ID " & lngFacID
    End If
    Set rsCodes = cnCodes.Execute(strSQL)
    Do Until rsCodes.EOF
        strProvider = vbNullString
        If Not IsNull(rsCodes.Fields("PROVIDER_TITLE")) Then
            strProvider = UCase(rsCodes.Fields("PROVIDER_TITLE"))
        End If
        If Trim(clsProvider.Title) <> vbNullString Then
            strCompare = UCase(clsProvider.Title)
            strCompare = Replace(strCompare, ".", "")
            strCompare = Trim(strCompare)
        Else
            If piProviderID = 0 Then
                Get_DefaultProvider
                Get_Providers piProviderID
            Else
                Get_Providers piProviderID
            End If
            strCompare = UCase(clsProvider.Title)
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
    Resume
End Function
Private Function Is_NewAppt() As Boolean
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim strSQL As String
Dim Idx%

    On Error GoTo Error_Is_NewAppt
    Screen.MousePointer = vbHourglass
    If udtSchedPatient.IntId = vbNullString Then udtSchedPatient.IntId = lblID2.Caption
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
        strSQL = "Get__Patient_Appt_ID " & udtAppointment.ProviderID & ", " & udtAppointment.FacilityID & "," _
            & Val(udtSchedPatient.IntId) & ",  '" & Format(udtAppointment.SchedTime, "hh:mm AM/PM") & "' '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "' "
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

    Select Case UCase(frmPatientAppointment.mInterval)
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
        .ReDraw = False
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
        If Trim(vstrHAlign) = vbNullString Then
          .TypeHAlign = TypeHAlignCenter
        ElseIf vstrHAlign = "TypeHAlignLeft" Then
          .TypeHAlign = TypeHAlignLeft
        End If
        If vlngBackColor > 0 Then .BackColor = vlngBackColor
        If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
    .ReDraw = True
    End With
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

    On Error GoTo Error_Get_ApptRange
    Screen.MousePointer = vbHourglass
    lngProviderID = vlngProviderID
    lngPatientID = vlngPatientID
    Clear_SchedPatient
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
    Get_SchedPatient lngPatientID
    If cnSched.State = adStateClosed Then cnSched.open psConnect(0)
    udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    strSQL = "Get_PT_Appt_ProviderID_Facility " & lngPatientID & ", " & udtAppointment.ProviderID & ", " _
        & lngFacilityID & ", '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "'"
        udtAppointment.VisitType = vbNullString
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            Do Until .EOF
                    strTestTime = vbNullString
                    strTestDur = vbNullString
                    lngApptID = 0
                    If Not IsNull((!APPT_TIME)) Then
                        strTestTime = Format((!APPT_TIME), "hh:mm AM/PM")
                    End If
                    If Not IsNull((!Duration)) Then
                        strTestDur = (!Duration)
                    End If
                    lngApptID = (!APPT_ID)
                .MoveNext
                If strTestDur <> vbNullString Then
                    strTestEnd = DateAdd("n", Val(Replace(strTestDur, "minutes", "")), strTestTime)
                    If Format(udtAppointment.SchedTime, "hh:mm AM/PM") >= Format(strTestTime, "hh:mm AM/PM") And Format(udtAppointment.SchedTime, "hh:mm AM/PM") <= Format(strTestEnd, "hh:mm AM/PM") Then
                        udtAppointment.SchedTime = Format(strTestTime, "hh:mm AM/PM")
                        txtDate.Value = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
                       ' ccdTime.Value = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                        For i = 0 To cboTime.listcount - 1
                            If cboTime.List(i) = udtAppointment.SchedTime Then
                                cboTime.Text = cboTime.List(i)
                                Exit For
                            End If
                        Next i
                        mstrOrigApptTime = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                        Exit Do
                    End If
                End If
                 lngApptID = 0
            Loop
        End With
    If lngApptID > 0 Then
        strSQL = "Get_Appointment_Appt_ID" & lngApptID
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
                If udtAppointment.SchedType <> vbNullString And udtAppointment.SchedType <> "Patient" Then
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
                'Load_Visit
                For i = 0 To lstPurVisit.listcount - 1
                   If lstPurVisit.List(i) = udtAppointment.VisitType Then
                       mbooByPass = True
                       lstPurVisit.Selected(i) = True
                       Exit For
                   End If
                Next i
                If udtAppointment.SchedTime <> vbNullString And udtAppointment.SchedDate <> vbNullString Then
                    lblAppt2.Caption = Format(udtAppointment.SchedDate, "MMMM DD, YYYY") & " " & Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                End If
                If Trim(udtSchedPatient.ApptComment) <> vbNullString Then txtComments.Text = Trim(udtSchedPatient.ApptComment)
                mUpdate = True
                udtSchedPatient.IntId = lngPatientID
                Select Case udtAppointment.SchedType
                   Case "Patient"
                       Load_Patient udtSchedPatient.IntId
                   Case "Physician"
                       Read_ReferralData udtSchedPatient.IntId
                   Case "Vendor"
                       Read_VendorData udtSchedPatient.IntId
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
                If udtAppointment.SchedType = "Patient" Then Get_Referrals
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
    Resume
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
    strSQL = "Get_Repeat_Appt_PatientID " & vlngFacility & ", " & vlngProvider & "," & Val(udtSchedPatient.IntId)
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
    Resume
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
    strSQL = "SELECT * FROM APPOINTMENTS WHERE PATIENT_ID = " & Val(udtSchedPatient.IntId) & "  AND " _
        & "PROVIDER_ID = " & lngProviderID & " AND " _
        & "CONVERT(VARCHAR, APPT_DATE , 10) = '" & Format(ccdRevisedDate.Value, "mm/dd/yyyy") & "' ORDER BY PROVIDER_ID, APPT_TIME DESC"
    Set rsApptGrid = cnApptGrid.Execute(strSQL)
    With rsApptGrid
        Do Until .EOF
            lngApptID = 0
            strPatType = vbNullString
            lngApptID = (!APPT_ID)
            If lngApptID > 0 Then
                lngRow = grdAppointChange.MaxRows + 1
                grdAppointChange.MaxRows = lngRow
                'do terms
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
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case Fld.Name
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
                                      '  Format_CellTypeCombo grdAppointChange, lngRow, lngCol, "1", Load_TimeGrid, "TypeHAlignCenter", "APPT_TIME"
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
                                        'Format_CellTypeCombo grdAppointChange, lngRow, lngCol, "1", Get_Grid_Status, "TypeHAlignCenter", "STATUS"
                                        Set_GridCurSel grdAppointChange, Trim(Fld.Value), lngCol, lngRow, "1", "STATUS"
                                        'Format_CellTypeStatic grdAppointChange, lngRow, lngCol, "1", Trim(Fld.Value), "TypeHAlignCenter", UCase(Fld.Name)
                                    End If
                                End If
                            Case "NOTES"
                                Arr = Split(LookUpCellTag(Fld.Name), "|")
                                If UBound(Arr) > 0 Then
                                    lngCol = Arr(0)
                                    If lngCol > 0 Then
                                        Format_CellTypeEdit grdAppointChange, lngRow, lngCol, "1", Trim(Fld.Value), "TypeHAlignLeft", UCase(Fld.Name)
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
            If lngRow > 0 Then
                grdAppointChange.SetRowItemData lngRow, lngApptID
            End If
        Loop
    End With
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
    Resume
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






Private Function Load_Patient(vstrData As String)
Dim strSQL As String
Dim strName As String
Dim strLastName As String
Dim cnSearch As New ADODB.Connection
Dim rsSearch As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngPatID As Long
Dim i%
Dim lngDbPatID As Long
Dim strList As String
Dim strFirst As String
Dim Arr

    On Error GoTo Error_Load_Patient
    Screen.MousePointer = vbHourglass
    frmSchedPatient.Caption = "Search Patients"
    If vstrData = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    lstListBox.Clear
    vstrData = Replace(vstrData, "'", "''")
    If TextFound(",", vstrData, Len(","), True) Then
        Arr = Split(vstrData, ",")
        If UBound(Arr) > 0 Then
            strFirst = Trim(Arr(1))
            strList = Arr(0)
            strFirst = Replace(strFirst, " ", "")
            strList = Replace(strList, " ", "")
            strList = Trim(strList)
        End If
    Else
        strList = vstrData
    End If
    If TextFound("'", strList, Len("'"), True) Then
        strList = Replace(strList, "'", "%")
    End If
    If cnSearch.State = adStateClosed Then cnSearch.open psConnect(0)
    If IsNumeric(vstrData) Then
        strSQL = "Search_Patient_EMR_ID  '" & vstrData & "%'"
        Set rsSearch = cnSearch.Execute(strSQL)
        If rsSearch.EOF Then
            strSQL = "Search_Patient_External_ID '%" & vstrData & "%'"
            Set rsSearch = cnSearch.Execute(strSQL)
            With rsSearch
                Set Flds = rsSearch.Fields
                lngDbPatID = 0
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            If UCase(Fld.Name) = "PATIENT_ID" Then
                                lngDbPatID = Fld.Value
                               Exit For
                            End If
                        End If
                    End If
                Next
                .MoveNext
                Get_PatientName lngDbPatID
                If Trim(udtPatientName.Suffix) <> vbNullString Then
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                Else
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                End If
                If Trim(udtPatientName.Name) <> vbNullString Then
                    lstListBox.AddItem Trim(udtPatientName.Name)
                    lstListBox.ItemData(lstListBox.NewIndex) = lngDbPatID
                End If
            End With
        Else
            With rsSearch
                Do Until .EOF
                Set Flds = rsSearch.Fields
                lngDbPatID = 0
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            If UCase(Fld.Name) = "PATIENT_ID" Then
                               lngDbPatID = Fld.Value
                               Exit For
                            End If
                        End If
                    End If
                Next
                .MoveNext
                Get_PatientName lngDbPatID
                If Trim(udtPatientName.Suffix) <> vbNullString Then
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                Else
                    udtPatientName.Name = udtPatientName.LastName & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                End If
                If Trim(udtPatientName.Name) <> vbNullString Then
                    lstListBox.AddItem Trim(udtPatientName.Name)
                    lstListBox.ItemData(lstListBox.NewIndex) = lngDbPatID
                End If
                Loop
            End With
        End If
    Else
        If Trim(strFirst) <> vbNullString Then
             strSQL = "Search_Patient_FirstName_Last_Name '" & strList & "%', '%" & strFirst & "%'"
        Else
            strSQL = "Search_Patient_Last_Name'" & strList & "%'"
        End If
        Set rsSearch = cnSearch.Execute(strSQL)
            With rsSearch
                Do Until .EOF
                Set Flds = rsSearch.Fields
                lngDbPatID = 0
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            If UCase(Fld.Name) = "PATIENT_ID" Then
                               lngDbPatID = Fld.Value
                               Exit For
                            End If
                        End If
                    End If
                Next
                .MoveNext
                Get_PatientName lngDbPatID
                If Trim(udtPatientName.Suffix) <> vbNullString Then
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                Else
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                End If
                If Trim(udtPatientName.Name) <> vbNullString Then
                    mbooByPass = True
                    lstListBox.AddItem Trim(udtPatientName.Name)
                    lstListBox.ItemData(lstListBox.NewIndex) = lngDbPatID
                End If
            Loop
        End With
    End If
    Clear_UdtPatient
    If cnSearch.State = adStateOpen Then
        cnSearch.Close
        Set cnSearch = Nothing
        Set rsSearch = Nothing
        Set Flds = Nothing
    End If
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Patient", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Public Function Load_SpecificPatient()
Dim i%
Dim lngPatientID As Long

    lngPatientID = piPatientID
    If lngPatientID > 0 Then
        Load_Patient CStr(lngPatientID)
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
    If cboTime.listcount <= 0 Then Load_ApptTime
    lstListBox.Visible = True
    cctSearch.Visible = True
    lblPatInfo.Visible = True
    cmdNewPatient.Visible = True
   Exit Function
Error_ReadData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "ReadData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Load_ApptTime()
Dim Idx As Integer
Dim strTime As String
Dim Arr

    cboTime.Clear
    Select Case UCase(frmPatientAppointment.mInterval)
        Case "EVERY 15 MIN"
           strTime = "07:00 AM" & "|" & "07:15 AM" & "|" & "07:30 AM" & "|" & "07:45 AM" & "|" & "08:00 AM" & "|" & "08:15 AM" & "|" & "08:30 AM" & "|" & "08:45 AM" _
           & "|" & "09:00 AM" & "|" & "09:15 AM" & "|" & "09:30 AM" & "|" & "09:45 AM" & "|" & "10:00 AM" & "|" & "10:15 AM" & "|" & "10:30 AM" & "|" & "10:45 AM" _
             & "|" & "11:00 AM" & "|" & "11:15 AM" & "|" & "11:30 AM" & "|" & "11:45 AM" & "|" & "12:00 PM" & "|" & "12:15 PM" & "|" & "12:30 PM" & "|" & "12:45 PM" _
             & "|" & "01:00 PM" & "|" & "01:15 PM" & "|" & "01:30 PM" & "|" & "01:45 PM" & "|" & "02:00 PM" & "|" & "02:15 PM" & "|" & "02:30 PM" & "|" & "02:45 PM" _
             & "|" & "03:00 PM" & "|" & "03:15 PM" & "|" & "03:30 PM" & "|" & "03:45 PM" & "|" & "04:00 PM" & "|" & "04:15 PM" & "|" & "04:30 PM" & "|" & "04:45 PM" _
             & "|" & "05:00 PM" & "|" & "05:15 PM" & "|" & "05:30 PM" & "|" & "05:45 PM" & "|" & "06:00 PM" & "|" & "06:15 PM" & "|" & "06:30 PM" & "|" & "06:45 PM" _
             & "|" & "07:00 PM" & "|" & "07:15 PM" & "|" & "07:30 PM" & "|" & "07:45 PM" & "|" & "08:00 PM" & "|" & "08:15 PM" & "|" & "08:30 PM" & "|" & "08:45 PM"
       Case "EVERY 5 MIN"
           strTime = "07:00 AM" & "|" & "07:05 AM" & "|" & "07:10 AM" & "|" & "07:15 AM" & "|" & "07:20 AM" & "|" & "07:25 AM" & "|" & "07:30 AM" & "|" & "07:35 AM" _
             & "|" & "07:40 AM" & "|" & "07:45 AM" & "|" & "07:50 AM" & "|" & "07:55 AM" & "|" & "08:00 AM" & "|" & "08:05 AM" & "|" & "08:10 AM" & "|" & "08:15 AM" _
             & "|" & "08:20 AM" & "|" & "08:25 AM" & "|" & "08:30 AM" & "|" & "08:35 AM" & "|" & "08:40 AM" & "|" & "08:45 AM" & "|" & "08:50 AM" & "|" & "08:55 AM" _
             & "|" & "09:00 AM" & "|" & "09:05 AM" & "|" & "09:10 AM" & "|" & "09:15 AM" & "|" & "09:20 AM" & "|" & "09:25 AM" & "|" & "09:30 AM" & "|" & "09:35 AM" _
             & "|" & "09:40 AM" & "|" & "09:45 AM" & "|" & "09:50 AM" & "|" & "09:55 AM" & "|" & "10:00 AM" & "|" & "10:05 AM" & "|" & "10:10 AM" & "|" & "10:15 AM" _
             & "|" & "10:20 AM" & "|" & "10:25 AM" & "|" & "10:30 AM" & "|" & "10:35 AM" & "|" & "10:40 AM" & "|" & "10:45 AM" & "|" & "10:50 AM" & "|" & "10:55 AM" _
             & "|" & "11:00 AM" & "|" & "11:05 AM" & "|" & "11:10 AM" & "|" & "11:15 AM" & "|" & "11:20 AM" & "|" & "11:25 AM" & "|" & "11:30 AM" & "|" & "11:35 AM" _
             & "|" & "11:40 AM" & "|" & "11:45 AM" & "|" & "11:50 AM" & "|" & "11:55 AM" & "|" & "12:00 PM" & "|" & "12:05 PM" & "|" & "12:10 PM" & "|" & "12:15 PM" _
             & "|" & "12:20 PM" & "|" & "12:25 PM" & "|" & "12:30 PM" & "|" & "12:35 PM" & "|" & "12:40 PM" & "|" & "12:45 PM" & "|" & "12:50 PM" & "|" & "12:55 PM" _
             & "01:00 PM" & "|" & "01:05 PM" & "|" & "01:10 PM" & "|" & "01:15 PM" _
             & "|" & "01:20 PM" & "|" & "01:25 PM" & "|" & "01:30 PM" & "|" & "01:35 PM" & "|" & "01:40 PM" & "|" & "01:45 PM" & "|" & "01:50 PM" & "|" & "01:55 PM" _
             & "|" & "02:00 PM" & "|" & "02:05 PM" & "|" & "02:10 PM" & "|" & "02:15 PM" & "|" & "02:20 PM" & "|" & "02:25 PM" & "|" & "02:30 PM" & "|" & "02:35 PM" _
             & "|" & "02:40 PM" & "|" & "02:45 PM" & "|" & "02:50 PM" & "|" & "02:55 PM" & "|" & "03:00 PM" & "|" & "03:05 PM" & "|" & "03:10 PM" & "|" & "03:15 PM" _
             & "|" & "03:20 PM" & "|" & "03:25 PM" & "|" & "03:30 PM" & "|" & "03:35 PM" & "|" & "03:40 PM" & "|" & "03:45 PM" & "|" & "03:50 PM" & "|" & "03:55 PM" _
             & "|" & "04:00 PM" & "|" & "04:05 PM" & "|" & "04:10 PM" & "|" & "04:15 PM" & "|" & "04:20 PM" & "|" & "04:25 PM" & "|" & "04:30 PM" & "|" & "04:35 PM" _
             & "|" & "04:40 PM" & "|" & "04:45 PM" & "|" & "04:50 PM" & "|" & "04:55 PM" & "|" & "05:00 PM" & "|" & "05:05 PM" & "|" & "05:10 PM" & "|" & "05:15 PM" _
             & "|" & "05:20 PM" & "|" & "05:25 PM" & "|" & "05:30 PM" & "|" & "05:35 PM" & "|" & "05:40 PM" & "|" & "05:45 PM" & "|" & "05:50 PM" & "|" & "05:55 PM" _
             & "|" & "06:00 PM" & "|" & "06:05 PM" & "|" & "06:10 PM" & "|" & "06:15 PM" & "|" & "06:20 PM" & "|" & "06:25 PM" & "|" & "06:30 PM" & "|" & "06:35 PM" _
             & "|" & "06:40 PM" & "|" & "06:45 PM" & "|" & "06:50 PM" & "|" & "06:55 PM" & "|" & "07:00 PM" & "|" & "07:05 PM" & "|" & "07:10 PM" & "|" & "07:15 PM" _
             & "|" & "07:20 PM" & "|" & "07:25 PM" & "|" & "07:30 PM" & "|" & "07:35 PM" & "|" & "07:40 PM" & "|" & "07:45 PM" & "|" & "07:50 PM" & "|" & "07:55 PM" _
             & "|" & "08:00 PM" & "|" & "08:05 PM" & "|" & "08:10 PM" & "|" & "08:15 PM" & "|" & "08:20 PM" & "|" & "08:25 PM" & "|" & "08:30 PM" & "|" & "08:35 PM" _
             & "|" & "08:40 PM" & "|" & "08:45 PM" & "|" & "08:50 PM" & "|" & "08:55 PM"
        Case "EVERY 20 MIN"
           strTime = "07:00 AM" & "|" & "07:20 AM" & "|" & "07:40 AM" & "|" & "08:00 AM" & "|" & "08:20 AM" & "|" & "08:40 AM" _
             & "|" & "09:00 AM" & "|" & "09:20 AM" & "|" & "09:40 AM" & "|" & "10:00 AM" & "|" & "10:20 AM" & "|" & "10:40 AM" _
             & "|" & "11:00 AM" & "|" & "11:20 AM" & "|" & "11:40 AM" & "|" & "12:00 PM" & "|" & "12:20 PM" & "|" & "12:40 PM" _
             & "|" & "01:00 PM" & "|" & "01:20 PM" & "|" & "01:40 PM" & "|" & "02:00 PM" & "|" & "02:20 PM" & "|" & "02:40 PM" _
             & "|" & "03:00 PM" & "|" & "03:20 PM" & "|" & "03:40 PM" & "|" & "04:00 PM" & "|" & "04:20 PM" & "|" & "04:40 PM" _
             & "|" & "05:00 PM" & "|" & "05:20 PM" & "|" & "05:40 PM" & "|" & "06:00 PM" & "|" & "06:20 PM" & "|" & "06:40 PM" _
             & "|" & "07:00 PM" & "|" & "07:20 PM" & "|" & "07:40 PM" & "|" & "08:00 PM" & "|" & "08:20 PM" & "|" & "08:40 PM"
       Case "EVERY 10 MIN"
        strTime = "07:00 AM" & "|" & "07:10 AM" & "|" & "07:20 AM" & "|" & "07:30 AM" & "|" & "07:40 AM" & "|" & "07:50 AM" _
          & "|" & "08:00 AM" & "|" & "08:10 AM" & "|" & "08:20 AM" & "|" & "08:30 AM" & "|" & "08:40 AM" & "|" & "08:50 AM" _
          & "|" & "09:00 AM" & "|" & "09:10 AM" & "|" & "09:20 AM" & "|" & "09:30 AM" & "|" & "09:40 AM" & "|" & "09:50 AM" _
          & "|" & "10:00 AM" & "|" & "10:10 AM" & "|" & "10:20 AM" & "|" & "10:30 AM" & "|" & "10:40 AM" & "|" & "10:50 AM" _
          & "|" & "11:00 AM" & "|" & "11:10 AM" & "|" & "11:20 AM" & "|" & "11:30 AM" & "|" & "11:40 AM" & "|" & "11:50 AM" _
          & "|" & "12:00 PM" & "|" & "12:10 PM" & "|" & "12:20 PM" & "|" & "12:30 PM" & "|" & "12:40 PM" & "|" & "12:50 PM" _
          & "|" & "01:00 PM" & "|" & "01:10 PM" & "|" & "01:20 PM" & "|" & "01:30 PM" & "|" & "01:40 PM" & "|" & "01:50 PM" _
          & "|" & "02:00 PM" & "|" & "02:10 PM" & "|" & "02:20 PM" & "|" & "02:30 PM" & "|" & "02:40 AM" & "|" & "02:50 PM" _
          & "|" & "03:00 PM" & "|" & "03:10 PM" & "|" & "03:20 PM" & "|" & "03:30 PM" & "|" & "03:40 PM" & "|" & "03:50 PM" _
          & "|" & "04:00 PM" & "|" & "04:10 PM" & "|" & "04:20 PM" & "|" & "04:30 PM" & "|" & "04:40 PM" & "|" & "04:50 PM" _
          & "|" & "05:00 PM" & "|" & "05:10 PM" & "|" & "05:20 PM" & "|" & "05:30 PM" & "|" & "05:40 PM" & "|" & "05:50 PM" _
          & "|" & "06:00 PM" & "|" & "06:10 PM" & "|" & "06:20 PM" & "|" & "06:30 PM" & "|" & "06:40 PM" & "|" & "06:50 PM" _
          & "|" & "07:00 PM" & "|" & "07:10 PM" & "|" & "07:20 PM" & "|" & "07:30 PM" & "|" & "07:40 PM" & "|" & "07:50 PM" _
          & "|" & "08:00 PM" & "|" & "08:10 PM" & "|" & "08:20 PM" & "|" & "08:30 PM" & "|" & "08:40 PM" & "|" & "08:50 PM"
    End Select
    Arr = Split(strTime, "|")
    If UBound(Arr) > 0 Then
        For Idx = LBound(Arr) To UBound(Arr)
            If Trim(Arr(Idx)) <> vbNullString Then
                cboTime.AddItem CStr(Arr(Idx))
            End If
        Next Idx
    End If
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
    Resume
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

    On Error GoTo Error_Rescheduled_Patient
    Screen.MousePointer = vbHourglass
    If Trim(cboStatus.Text) = vbNullString Then
         MsgBox "Please Select Reschedule in status and provide a reason for reschedule.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If Trim(mstrOrigApptTime) = vbNullString Then
        Screen.MousePointer = vbDefault
         MsgBox "No Information Available on Current Appointment, Please Reselect from Appointment Screen!", vbInformation, "PatientTrac Scheduling"
        frmPatientAppointment.mBeenHere = True
        Screen.MousePointer = vbDefault
        mstrValidAppoint = vbNullString: mbooAlreadyBook = False: mbooCancel = False
        fraInfo.Visible = False
        Clear_PatientStrings
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    strDate = Format(txtDate.Value, "mm/dd/yyyy")
    If udtAppointment.SchedTime = vbNullString Then
         udtAppointment.SchedTime = Format(cboTime.Text, "hh:mm AM/PM")
    End If
    If udtSchedPatient.Duration = vbNullString Then udtSchedPatient.Duration = Trim(cboDuration.Text)
    strNotes = Trim(txtComments.Text)
    strNotes = Replace(txtComments, "'", "''")
    strConfirm = Format(txtApptConfirm.Value, "mm/dd/yyyy")
    For i = 0 To lstPurVisit.listcount - 1
        If lstPurVisit.Selected(i) Then
            strPurpose = lstPurVisit.List(i)
            Exit For
        End If
    Next i
    strReason = Trim(cboReason.Text)
    strNotes = Trim(txtComments.Text)
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    strSQL = "Get__Patient_Appt_ID " & udtAppointment.ProviderID & " ," & udtAppointment.FacilityID & ", " & Val(udtSchedPatient.IntId) & ", '" _
        & Format(mstrOrigApptTime, "hh:mm AM/PM") & "', '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "'"
    Set rsAppt = cnAppt.Execute(strSQL)
    If Not rsAppt.EOF Then
        lngApptID = rsAppt.Fields("APPT_ID")
    End If
    sToday = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    strNotes = Literals(strNotes, "'")
    strReason = Literals(strReason, "'")
    If lngApptID > 0 Then
        strSQL = "Update_Appointments_Rescheduled 'Rescheduled', '" & strReason & "', '" & strNotes & "', '" & sToday & "'," & lngApptID
        Set rsAppt = cnAppt.Execute(strSQL)
        MsgBox "Please close and select new Appointment Date and Time, this appointment has been updated as inactive!", vbInformation, "PatientTrac Appointment Scheduling"
        cboStatus.ListIndex = -1
        cboReason.ListIndex = -1
        txtComments.Text = vbNullString
        Screen.MousePointer = vbDefault
    End If
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    frmPatientAppointment.mBeenHere = True
    Screen.MousePointer = vbDefault
    mstrValidAppoint = vbNullString: mbooAlreadyBook = False: mbooCancel = False
    fraInfo.Visible = False
    Clear_PatientStrings
    Screen.MousePointer = vbDefault
  '  Unload Me
    Exit Function
Error_Rescheduled_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Rescheduled_Patient", App.EXEName & "." & Me.Name
    Exit Function
    Resume
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
    If Trim(cboProvider.Text) <> vbNullString Then
        For i = 0 To cboProvider.listcount - 1
            If cboProvider.List(i) = cboProvider.Text Then
                udtAppointment.ProviderID = cboProvider.ItemData(i)
                Exit For
            End If
        Next i
    End If
    strNow = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If Trim(cboStatus.Text) = vbNullString Or Trim(cboStatus.Text) <> "Cancelled" Then Exit Function
    If mstrOrigApptTime = vbNullString Then mstrOrigApptTime = udtAppointment.SchedTime
    strReason = Trim(cboReason.Text)
    strNotes = Trim(txtComments.Text)
    strReason = Replace(strReason, "'", "''")
    strNotes = Replace(strNotes, "'", "''")
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    strSQL = "Get__Patient_Appt_ID_Type  " & udtAppointment.ProviderID & ", " & udtAppointment.FacilityID & ", " _
        & Val(udtSchedPatient.IntId) & ", '" & udtAppointment.SchedTime & "', '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "','" _
        & udtAppointment.SchedType & "' "
        Set rsAppt = cnAppt.Execute(strSQL)
        With rsAppt
            If Not .EOF Then
                If Not IsNull(!APPT_ID) Then
                    lngApptID = (!APPT_ID)
                End If
            End If
        End With
    If lngApptID > 0 Then
        If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
        strSQL = "Update_Appointments_Rescheduled  'Cancelled','" & strReason & "','" & strNotes & "','" & strNow & "', " & lngApptID
        Set rsAppt = cnAppt.Execute(strSQL)
    End If
    udtAppointment.Status = Trim(udtAppointment.Status)
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Cancel_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Cancel_Patient", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
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
    strCurrentTime = udtAppointment.SchedTime
    If udtAppointment.SchedType = "Patient" Then
       Get_SchedPatient lngPatientID
    ElseIf udtAppointment.SchedType = "Physician" Then
       Get_SchedPhysician lngPatientID
    ElseIf udtAppointment.SchedType = "Vendor" Then
       Get_SchedVendor lngPatientID
    End If
    udtSchedPatient.ApptComment = Replace(udtSchedPatient.ApptComment, "''", "'")
    txtComments.Text = Trim(udtSchedPatient.ApptComment)
    udtAppointment.SchedTime = strCurrentTime
    cmdApply.Enabled = True
    mbooGotYou = True
    cmdEditPatient.Visible = True
    fraInfo.Visible = True
    lblID2.Caption = lngPatientID
    udtSchedPatient.IntId = lngPatientID
    lblName2.Caption = Trim(udtSchedPatient.FirstName) & " " & Trim(udtSchedPatient.LastName)
    lblPhone2.Caption = Trim(udtSchedPatient.Phone)
    lblCell2.Caption = Trim(udtSchedPatient.CellPhone)
    If Len(udtSchedPatient.WorkPhone) > 3 Then lblWorkPhone.Caption = Trim(udtSchedPatient.WorkPhone)
    udtSchedPatient.Address = Replace(Trim(udtSchedPatient.Address), "Address:", "")
    lblAddress.Caption = "Address:" & " " & Trim(udtSchedPatient.Address)
    strBestNumber = udtSchedPatient.PrefContact
    If strBestNumber <> vbNullString Then
        If TextFound("Cell", strBestNumber, Len("Cell"), True) Then
            lblCell2.Caption = lblCell2.Caption & " " & "Preferred"
            lblCell2.ToolTipText = "Cell Phone is preferred contact"
        ElseIf TextFound("Work", strBestNumber, Len("Work"), True) Then
            lblWorkPhone.Caption = lblWorkPhone.Caption & " " & "Preferred"
            lblWorkPhone.ToolTipText = "Work Phone is preferred contact"
        ElseIf TextFound("Home", strBestNumber, Len("Home"), True) Then
            lblPhone2.Caption = lblPhone2.Caption & " " & "Preferred"
            lblPhone2.ToolTipText = "Home Phone is preferred contact"
        End If
    End If
    If mbooGotYou And udtAppointment.SchedType = "Patient" Then
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
    ElseIf udtAppointment.SchedType <> "Patient" And lngPatientID > 0 Then
       Get_Info udtAppointment.ProviderID, lngPatientID
       cboPatientType.Enabled = False
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
    Resume
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
    
    If udtSchedPatient.Status <> vbNullString Then
        Select Case udtSchedPatient.Status
            Case "Rescheduled", "Cancelled", "No Show", "Confirmed", "Present"
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
    strRange = DateAdd("d", -7, udtAppointment.SchedDate)
    If Trim(strTime) = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        strStart = Format(vstrAppt_Time)
        strStart = Format(strStart, "hh:mm AM/PM")
        vstrDuration = Replace(vstrDuration, "minutes", "")
        vstrDuration = Trim(vstrDuration)
        lngDur = Val(vstrDuration) + 1
        strEnd = DateAdd("n", lngDur, vstrAppt_Time)
        strEnd = Format(strEnd, "hh:mm AM/PM")
        If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
        strStartDate = DateAdd("d", -1, Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
        strEndDate = DateAdd("d", 1, Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
        '1st Case  Day of Appt
        strSQL = "Validate_ProvBlock_Phase1 " & udtAppointment.ProviderID & ", '" & Format(strStart, "hh:mm AM/PM") & "','" _
            & Format(strEnd, "hh:mm AM/PM") & "', '" & Format(strStartDate, "mm/dd/yyyy") & "', '" & Format(strEndDate, "mm/dd/yyyy") & "'"
    '        SELECT START_TIME, END_TIME, START_DATE, END_DATE FROM APPOINTMENTS_BLOCKS WHERE PROVIDER_ID = AND " _
    '           & "ACTIVE = 1 AND " _
    '           & "convert(DATETIME, START_TIME, 108) BETWEEN '" & Format(strStart, "hh:mm AM/PM") & "' AND '" & Format(strEnd, "hh:mm AM/PM") & "' AND " _
    '           & "DURATION = '1 Day' AND  CONVERT(VARCHAR, START_DATE, 10)  > '" & Format(strStartDate, "mm/dd/yyyy") & "' AND CONVERT(VARCHAR, START_DATE, 10) < '" & Format(strEndDate, "mm/dd/yyyy") & "' And " _
    '            & "CONVERT(VARCHAR,  END_DATE, 10) < '" & Format(strEndDate, "mm/dd/yyyy") & "'  ORDER BY START_TIME DESC"
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
           strWeekDay = Weekday_Return(WeekDay(udtAppointment.SchedDate))
            strSQL = "Validate_ProvBlock_Phase2 " & udtAppointment.ProviderID & ", " & udtFacility.FacID & ",'" & Format(strStart, "hh:mm AM/PM") & "','" _
                & Format(strEnd, "hh:mm AM/PM") & "', '" & Format(strStartDate, "mm/dd/yyyy") & "', '" & Format(strEndDate, "mm/dd/yyyy") & "','" _
                & strWeekDay & "'"
'            strSQL = "SELECT START_TIME, END_TIME, START_DATE, END_DATE FROM APPOINTMENTS_BLOCKS WHERE PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
'                & "FACILITY_ID = " & udtFacility.FacID & " AND DURATION = 'Permanent' AND convert(DATETIME, START_TIME, 108) " _
'                & "BETWEEN '" & Format(strStart, "hh:mm AM/PM") & "' AND '" & Format(strEnd, "hh:mm AM/PM") & "' AND  " _
'                & "ACTIVE = 1 AND DAYWEEK = '" & strWeekDay & "' AND  CONVERT(DATETIME, START_DATE, 101)  > '" & Format(strStartDate, "mm/dd/yyyy") & "' AND CONVERT(DATETIME, START_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "' And " _
'                & "CONVERT(DATETIME, END_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "'  ORDER BY START_TIME DESC"
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
            strSQL = "Validate_ProvBlock_Phase3 " & udtAppointment.ProviderID & ", " & udtFacility.FacID & ", '" _
                & Format(strStartDate, "mm/dd/yyyy") & "', '" & Format(strEndDate, "mm/dd/yyyy") & "', "
'            strSQL = "SELECT START_TIME, END_TIME FROM APPOINTMENTS_BLOCKS WHERE PROVIDER_ID = " & udtAppointment.ProviderID & " AND DAYWEEK = '' AND " _
'                & "FACILITY_ID = " & udtFacility.FacID & " AND DURATION = 'Permanent' AND ACTIVE = 1 AND " _
'                & "CONVERT(DATETIME, START_DATE, 101)  > '" & Format(strStartDate, "mm/dd/yyyy") & "' AND CONVERT(DATETIME, START_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "' And " _
'                & "CONVERT(DATETIME, END_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "'  ORDER BY START_TIME DESC"
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
'            strSQL = "SELECT START_TIME, END_TIME FROM APPOINTMENTS_BLOCKS WHERE PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
'                & "FACILITY_ID = " & udtFacility.FacID & " AND DURATION = 'Permanent' AND " _
'                & "CONVERT(DATETIME, START_DATE, 101)  > '" & Format(strStartDate, "mm/dd/yyyy") & "' AND CONVERT(DATETIME, START_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "' And " _
'                & "CONVERT(DATETIME, END_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "' AND ACTIVE = 1  " _
'                & "AND CONVERT(DATETIME, END_DATE, 101) > '" & Format(strStartDate, "mm/dd/yyyy") & "'  ORDER BY START_TIME DESC"
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
        '4th Duration Block out
        If Validate_ProvAppt = False Then
            strSQL = "Validate_ProvBlock_Phase5 " & udtAppointment.ProviderID & ", " & udtFacility.FacID & ",'" _
                & strRange & "'"
    '            strSQL = "SELECT * FROM APPOINTMENTS_BLOCKS WHERE PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
    '                & "FACILITY_ID = " & udtFacility.FacID & " And DateDiff(Day, '" & strRange & "', START_DATE) > 0 AND " _
    '                & "DURATION <> 'Permanent' AND DURATION <> '1 Day' AND ACTIVE = 1  ORDER BY START_TIME DESC"
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
                                            Case "REASON"
                                                udtBlock.Reason = Trim(Fld.Value)
                                            Case "DURATION"
                                                udtBlock.Duration = Trim(Fld.Value)
                                            Case "FACILITY_ID"
                                                udtFacility.FacID = Trim(Fld.Value)
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
                            If nDays > 1 Then
                                If DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate) > 0 And _
                                    DateDiff("d", udtBlock.EndDate, udtAppointment.SchedDate) < 0 Then
                                    Validate_ProvAppt = True
                                    If udtBlock.EndDate <> vbNullString Then
                                        lngDateDiff = DateDiff("d", Format(udtBlock.StartDate, "mm/dd/yyyy"), Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
                                        If lngDateDiff >= 0 Then
                                            lngDateDiff = DateDiff("n", strFoundStart, strTime)
                                            If lngDateDiff >= 0 Then
                                                lngDateDiff = DateDiff("n", strTime, strFoundEnd)
                                                If lngDateDiff >= 0 Then
                                                    lngDateDiff = DateDiff("d", Format(udtBlock.EndDate, "mm/dd/yyyy"), Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
                                                    If lngDateDiff <= 0 Then
                                                        MsgBox "Provider has blocked this time period!", vbCritical, "PatientTrac Scheduling"
                                                        Validate_ProvAppt = True
                                                        Exit Do
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
                        Loop
                    End If
        End If
    If Not Validate_ProvAppt Then
        strStart = DateAdd("n", -1, vstrAppt_Time)
        strStart = Format(strStart, "hh:mm AM/PM")
        lngDur = Val(vstrDuration) + 1
        strEnd = DateAdd("n", lngDur, vstrAppt_Time)
        strEnd = Format(strEnd, "hh:mm AM/PM")
        If Not Validate_ProvAppt Then
            If cnTerms.State = adStateOpen Then cnTerms.Close
            If cnTerms.State = adStateClosed Then cnTerms.open psConnect(0)
            strSQL = "Validate_ProvBlock_Phase6 " & udtAppointment.ProviderID & ", '" & Format(strData, "mm/dd/yyy") & "','" _
                & Format(strStart, "hh:mm AM/PM") & "', '" & Format(strEnd, "hh:mm AM/PM") & "'"
'            strSQL = "SELECT APPT_TIME, COUNT(APPT_TIME) As BOOKED_APPT, ACTIVE, PROVIDER_ID, APPT_DATE  " _
'                & "FROM APPOINTMENTS  " _
'                & "GROUP BY  APPT_TIME, PROVIDER_ID, PROVIDER_ID, ACTIVE, APPT_DATE " _
'                & "Having COUNT(APPT_TIME) > 0 AND CONVERT(VARCHAR, APPT_DATE, 10)  = '" & Format(strData, "mm/dd/yyy") & "' AND CONVERT(DATETIME, APPT_TIME, 108) " _
'                & "BETWEEN '" & Format(strStart, "hh:mm AM/PM") & "' AND '" & Format(strEnd, "hh:mm AM/PM") & "' AND ACTIVE = 1 AND PROVIDER_ID = " & udtAppointment.ProviderID
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
    Resume
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



Private Sub cboApptType_Click()
    
    If Trim(cboApptType.Text) <> vbNullString Then
        udtAppointment.SchedType = cboApptType.Text
    End If
    If IsProperCase(udtAppointment.SchedType) = "Physician" Then
        lblID1.Caption = "Physician ID"
    ElseIf IsProperCase(udtAppointment.SchedType) = "Patient" Then
        lblID1.Caption = "Patient ID"
    ElseIf IsProperCase(udtAppointment.SchedType) = "Vendor" Then
        lblID1.Caption = "Vendor ID"
    End If
    lstListBox.Clear
    cctSearch.Text = vbNullString
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
    If cboFacility.listcount <= 0 Then Get_Appt_Providers
    Load_Service lngFacilityID
    If clsEDI_Client.Site_ID = "9702" Then
        Clear_Surg_Facility
        Read_Patient_Status Val(udtSchedPatient.IntId)
        If Val(clsFacStatus.FacilityID) > 0 Then
            cctFacility.BackColor = NormalColor
            Get_Facilities Val(clsFacStatus.FacilityID)
            cctFacility.Text = clsFacility.Name
            If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
                cctFacility.Text = cctFacility.Text & vbNewLine & Trim(clsFacStatus.StatusDesc)
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
        cctFacility.Text = cboFacility.Text
    End If
    Exit Sub
Error_cboFacility_Click:
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
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
    Resume
End Sub

Private Sub Load_Visit(Optional vstrType As String)
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
    If mbooByPass Then mbooByPass = False: Exit Sub
    If vstrType = vbNullString Then
        cboPatientType.Clear
    End If
    For i = 0 To cboFacility.listcount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            lngFacID = cboFacility.ItemData(i)
            Exit For
        End If
    Next i
    Screen.MousePointer = vbHourglass
    lstPurVisit.Clear
    If udtAppointment.ProviderID > 0 Then Get_Providers udtAppointment.ProviderID
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    If Trim(vstrType) <> vbNullString Then
        strSQL = "Get_Appt_Types '" & vstrType & "%'"
    Else
        strSQL = "Get_Appt_Types_Facility_ID" & lngFacID
    End If
    Set rsCodes = cnCodes.Execute(strSQL)
    Do Until rsCodes.EOF
        strProvider = vbNullString
        If Not IsNull(rsCodes.Fields("PROVIDER_TITLE")) Then
            strProvider = UCase(rsCodes.Fields("PROVIDER_TITLE"))
        End If
'------------------
        If Trim(clsProvider.Title) <> vbNullString Then
            strCompare = UCase(clsProvider.Title)
            strCompare = Replace(strCompare, ".", "")
            strCompare = Trim(strCompare)
        Else
            If piProviderID = 0 Then
                'modGeneral.piProviderID = 0
               ' Get_DefaultProvider
                piProviderID = mlngLogonUser
                Get_Providers piProviderID
            Else
                Get_Providers piProviderID
            End If
            strCompare = UCase(clsProvider.Title)
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
                                    lstPurVisit.AddItem Trim(Fld.Value)
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
    If Trim(strPatType) <> vbNullString Then
        Arr = Split(strPatType, vbNewLine)
        If UBound(Arr) > 0 Then
            For i = LBound(Arr) To UBound(Arr)
                If Arr(i) <> vbNullString Then
                    cboPatientType.AddItem CStr(Arr(i))
                End If
            Next i
        End If
    End If
    Screen.MousePointer = vbDefault
     Exit Sub
Error_Load_Visit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Visit", App.EXEName & " " & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub cboProvider_Click()
Dim i%
    For i = 0 To cboProvider.listcount - 1
        If cboProvider.List(i) = cboProvider.Text Then
            udtAppointment.ProviderID = cboProvider.ItemData(i)
            Exit For
        End If
    Next i
    Get_Providers udtAppointment.ProviderID
    Exit Sub
End Sub



Private Function Read_ReferralData(vstrName As String)
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

    On Error GoTo Error_Read_ReferralData
    Screen.MousePointer = vbHourglass
    If lstListBox.listcount >= 0 Then
        lstListBox.Clear
    End If
    If cnPhys.State = adStateClosed Then cnPhys.open psConnect(0)
    strSQL = "Load_Referrals"
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
                lstListBox.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strTitle)
                If IsNumeric(lngRefID) Then lstListBox.ItemData(lstListBox.NewIndex) = lngRefID
            Else
                lstListBox.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strSuffix) & " " & Trim(strTitle)
                If IsNumeric(lngRefID) Then lstListBox.ItemData(lstListBox.NewIndex) = lngRefID
            End If
        Loop
    End With
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
    Resume
End Function


Private Sub cboReason_Click()
    If cboStatus.Text = "Rescheduled" Then
        Rescheduled_Patient
        Exit Sub
    End If
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
    Resume
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
        If cboStatus.Text = "Cancelled" Then
             If cboReason.Text = vbNullString Then
                MsgBox "Please enter a reason for the" & " " & cboStatus.Text, vbInformation, "PatientTrac Appointment Scheduling"
                Screen.MousePointer = vbDefault
            End If
        End If
    End If
    If cboStatus.Text = "Rescheduled" Then
        MsgBox "Please enter a reason for the" & " " & cboStatus.Text, vbInformation, "PatientTrac Appointment Scheduling"
        Screen.MousePointer = vbDefault
    End If
    If cboStatus.Text = "Modify Appointment" Then
        If Val(udtSchedPatient.IntId) > 0 Then
            Format_Grid
            Read_ApptGrid
            fraChange.Visible = True
            Screen.MousePointer = vbDefault
        End If
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
        If cboStatus.Text = "Cancelled" Or cboStatus.Text = "Rescheduled" Then
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
    Resume
End Sub


Private Sub cctRepeat_Click()
    
    Exit Sub
End Sub


Private Sub cctSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
Dim Arr

    'mbooSearch = True
    If Len(cctSearch.Text) > 0 Then
        strList = Trim(cctSearch.Text)
    End If
    lstListBox.Clear
    mbooByPass = True
    If strList <> vbNullString Then
        If udtAppointment.SchedType = "Patient" Then
            If Trim(strList) <> vbNullString And Len(strList) >= 2 Then
                Load_Patient strList
            End If
        ElseIf udtAppointment.SchedType = "Physician" Then
            Read_ReferralData strList
        ElseIf udtAppointment.SchedType = "Vendor" Then
            Read_VendorData strList
        End If
    End If
    mbooByPass = False
    'mbooSearch = False
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub Read_VendorData(vstrData As String)
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

    On Error GoTo ERROR_Read_VendorData
    Screen.MousePointer = vbHourglass
    lstListBox.Clear
    
    If cnReferral.State = adStateClosed Then cnReferral.open psConnect(0)
    strSQL = "Search_Vendor_LastName '" & vstrData & "%'"
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
         strSQL = "SELECT * FROM VENDORS WHERE VENDOR_CONTACT like '" & vstrData & "%'  ORDER BY VENDOR_NAME"
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
    Resume
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
    fraChange.ZOrder 0
    fraChange.Visible = True
    Exit Sub
End Sub

Private Sub cmdClear_Click()
    cctReasReferral.Text = vbNullString
    Exit Sub
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

    Screen.MousePointer = vbHourglass
    grdAppointChange.ProcessTab = True
    grdAppointChange.FontBold = True
    grdAppointChange.ForeColor = 0
    If Not IsNull(txtDate.Value) Then ccdRevisedDate.Value = Format(txtDate.Value, "mm/dd/yyyy")
    With grdAppointChange
        .ReDraw = False
        .MaxRows = 0
        .MaxCols = 7
        .ColHeaderRows = 4
        .ClearRange 1, 1, grdAppointChange.MaxCols, grdAppointChange.MaxRows, False
        .ColWidth(1) = 10
        .ColWidth(2) = 10
        .ColWidth(3) = 15
        .ColWidth(4) = 20
        .ColWidth(5) = 19
'        .ColWidth(6) = 9
        .ColWidth(6) = 10
        .ColWidth(7) = 20
        '.ColWidth(8) = 22
        Get_SchedPatient Val(udtSchedPatient.IntId)
        Format_CellTypeEdit grdAppointChange, SpreadHeader, 1, "9", "APPOINTMENTS FOR" & " " & UCase(udtSchedPatient.Name) & vbNewLine & Format(ccdRevisedDate.Value, "MMMM DD, YYYY"), "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdAppointChange.SetCellBorder 1, SpreadHeader, 9, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdAppointChange
        lngCol = 1
        Format_CellTypeStatic grdAppointChange, SpreadHeader + 2, lngCol, "1", "APPT DATE", "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeStatic grdAppointChange, SpreadHeader + 2, lngCol, "1", "APPT TIME", "TypeHAlignCenter", "APPT_TIME", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
'        lngCol = 3
'        Format_CellTypeStatic grdAppointChange, SpreadHeader + 2, lngCol, "1", "PROVIDER", "TypeHAlignCenter", "PROVIDER_ID", HeaderColor
'        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeStatic grdAppointChange, SpreadHeader + 2, lngCol, "1", "APPT TYPE", "TypeHAlignCenter", "PATIENT_TYPE", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeStatic grdAppointChange, SpreadHeader + 2, lngCol, "1", "VISIT TYPE", "TypeHAlignCenter", "PURPOSE_VISIT", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
'        lngCol = 6
'        Format_CellTypeEdit grdAppointChange, SpreadHeader + 2, lngCol, "1", "VISIT REASON", "TypeHAlignCenter", "CPT_CODE", HeaderColor
'        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeStatic grdAppointChange, SpreadHeader + 2, lngCol, "1", "DURATION", "TypeHAlignCenter", "DURATION", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 6
        Format_CellTypeStatic grdAppointChange, SpreadHeader + 2, lngCol, "1", "APPT STATUS", "TypeHAlignCenter", "STATUS", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 7
        Format_CellTypeStatic grdAppointChange, SpreadHeader + 2, lngCol, "1", "COMMENTS", "TypeHAlignCenter", "NOTES", HeaderColor
        grdAppointChange.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdAppointChange
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
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
             .TypeComboBoxhWnd = fpTerm.hwnd
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
    CD_ShowPrinter (Me.hwnd)
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
    Resume
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
    cboReferBy.AddItem cctReferral.Text
    cboReferBy.ItemData(cboReferBy.NewIndex) = clsReferral.ReferralID
    If clsReferral.ReferralID <> 0 Then
        For i = 0 To cboReferBy.listcount - 1
            If cboReferBy.ItemData(i) = clsReferral.ReferralID Then
                cboReferBy.Text = cboReferBy.List(i)
                Exit For
            End If
        Next i
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
    
'    Load frmNurseTemplates
'    frmNurseTemplates.fCaption = "Purpose of Visit"
'    frmNurseTemplates.fType = "Purpose of Visit"
'    frmNurseTemplates.fSubDomain = "sd_Appointment_lstPurVisit"
'    frmNurseTemplates.fConnectCR = psConnect(0)
'    frmNurseTemplates.fConnectTM = psConnect(1)
'    frmNurseTemplates.ReadData
'    frmNurseTemplates.Show vbModal
'    lstPurVisit.Clear
'    GetList lstPurVisit, "sd_Appointment_lstPurVisit"
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
    strSQL = "Get_Facility_ServicePlace " & vlngFacilityID
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
    Resume
End Function

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
        strSQL = "DELETE APPOINTMENTS  WHERE APPT_ID = " & lngPrimKey
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
    Resume
End Sub

Private Sub cmdRepeat_Click()
Dim i As Integer
    udtRepeatAppt.Comment = Trim(txtComments.Text)
    udtRepeatAppt.Comment = Replace(udtRepeatAppt.Comment, "'", "''")
    udtRepeatAppt.Status = Trim(cboStatus.Text)
    For i = 0 To lstPurVisit.listcount - 1
        If lstPurVisit.Selected(i) Then
            udtRepeatAppt.PurposeVisit = lstPurVisit.List(i)
            Exit For
        End If
    Next i
    If Trim(cboTime.Text) <> vbNullString Then udtRepeatAppt.StartTime = Format(cboTime.Text, "hh:mm AM/PM")
    If Not IsNull(txtDate.Value) Then udtRepeatAppt.StartDate = Format(txtDate.Value, "mm/dd/yyyy")
    If Trim(cboDuration.Text) <> vbNullString Then udtRepeatAppt.Duration = Trim(cboDuration.Text)
    If udtAppointment.ProviderID > 0 Then udtRepeatAppt.ProviderID = udtAppointment.ProviderID
    If udtFacility.FacID > 0 Then udtRepeatAppt.FacilityID = udtFacility.FacID
    frmMultiAppt.Show vbModal
    Get_Info udtAppointment.ProviderID, Val(udtSchedPatient.IntId)
    Exit Sub
End Sub

Private Sub cmdShowAll_Click()
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
                Case "Rescheduled", "Cancelled"
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
        If cnBlock.State = adStateClosed Then cnBlock.open psConnect(0)
            strSQL = "Update APPOINTMENTS set ACTIVE = " & lngActive & ", PURPOSE_VISIT = '" & udtAppointment.VisitType & "', Duration = '" & udtSchedPatient.Duration & "'," _
                & "NOTES = '" & udtSchedPatient.ApptComment & "', APPT_TIME  = '" & udtAppointment.SchedTime & "', " _
                & "PATIENT_TYPE = '" & udtSchedPatient.PatientType & "', UPDATE_DATE = '" & strNow & "'," _
                & "STATUS = '" & udtSchedPatient.Status & "',  " _
                & "USER_ID = '" & PsUserID & "', APPT_TYPE = '" & udtAppointment.SchedType & "' WHERE APPT_ID = " & lngPrimKey
        Set rsBlock = cnBlock.Execute(strSQL)
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
    Resume
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
                                grdAppointChange.SetCellBorder lngCol, Row, lngCol, Row, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
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
    Resume
End Sub

Private Sub lstListBox_ItemCheck(Item As Integer)
Dim i%
Dim strSQL
Dim lngPatientID As Long
Dim IAnswer As Integer
Dim strDate As String

    On Error GoTo Error_lstListBox_ItemCheck
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
    udtSchedPatient.IntId = vbNullString
    piPatientID = 0
    If lngPatientID > 0 Then
        mbLookupPatient = True
        udtSchedPatient.IntId = lngPatientID
        piPatientID = lngPatientID
        Set_PatientFrame lngPatientID
      '  Get_InsAuth
        If clsEDI_Client.Site_ID = vbNullString Then
            Get_EDI_Static
        End If
        If clsEDI_Client.Site_ID = "9702" Then
            Read_Patient_Status lngPatientID
            If Val(clsFacStatus.FacilityID) > 0 Then
                cctFacility.BackColor = NormalColor
                Get_Facilities Val(clsFacStatus.FacilityID)
                cctFacility.Text = clsFacility.Name
                cctFacility.Visible = True
                lblFac.Visible = True
                If Trim(cctFacility.Text) <> vbNullString Then
                    If Trim(cctFacility.Text) <> Trim(cboFacility.Text) And Trim(cboFacility.Text) <> vbNullString Then
                        cctFacility.BackColor = grdCurItem
                    End If
                End If
            End If
            If Trim(clsFacStatus.StatusDesc) <> vbNullString Then
                cctFacility.Text = cctFacility.Text & vbNewLine & Trim(clsFacStatus.StatusDesc)
            End If
            strDate = Format(txtDate.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
'            If IsDate(strDate) Then
'                Calculate_TreatRemain lngPatientID, strDate
'                If Val(clsInsBillData.TreatRemain) <= 2 And clsInsBillData.TreatRemain <> vbNullString Then
'                    MsgBox "This Patient has" & " " & clsInsBillData.TreatRemain & " " & "Insurance Authorizations remaining for this Appointment Date!", vbInformation, "PatientTrac Scheduling"
'                End If
'            End If
            cctFacility.Locked = True
            Get_PatientName lngPatientID
            If UCase(udtPatientName.Status) = "INACTIVE" Or UCase(udtPatientName.Status) = "DECEASED" Then
                txtList.Text = txtList.Text & vbNewLine & "ALERT:" & " " & UCase(udtPatientName.Status)
                If IsProperCase(udtPatientName.Status) = "Inactive" Then
                    IAnswer = MsgBox("You have selected an inactive patient. Are you sure you want to continue?", vbYesNo, "PatientTrac Encounter/Charges")
                ElseIf IsProperCase(udtPatientName.Status) = "Deceased" Then
                  IAnswer = MsgBox("You have selected a deceased patient. Are you sure you want to continue?", vbYesNo, "PatientTrac Encounter/Charges")
                End If
                If IAnswer = vbNo Then
                    Screen.MousePointer = vbDefault
                    udtPatientName.IntId = vbNullString
                    Clear_UdtPatient
                    Exit Sub
                End If
            End If
            Clear_UdtPatient
        Else
            cctFacility.Text = Trim(cboFacility.Text)
        End If
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
    Resume
End Sub

Private Sub lstPurVisit_ItemCheck(Item As Integer)
Dim Idx As Integer

    On Error GoTo Error_lstPurVisit_ItemCheck
    If mbooByPass Then mbooByPass = False: Exit Sub
    Screen.MousePointer = vbHourglass
    udtEncounter.Appt_Type = vbNullString
    udtEncounter.Intrvn_CD = vbNullString
    If mbooByPass Then mbooByPass = False: Exit Sub
    mbooListValid = False
    udtEncounter.Appt_Type = lstPurVisit.List(Item)
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
    If udtEncounter.Intrvn_CD <> vbNullString Then
        lblCode.Caption = udtEncounter.Intrvn_CD
    Else
        LookUpCPTCode udtEncounter.Appt_Type
        lblCode.Caption = udtEncounter.Intrvn_CD
    End If
    mbExistSelection = True
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_lstPurVisit_ItemCheck:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "lstPurVisit_ItemCheck", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub lstReferral_ItemCheck(Item As Integer)
Dim Idx%
    If mbooByPass Then mbooByPass = False: Exit Sub
    clsReferral.ReferralID = 0
    clsReferral.ReferralID = lstReferral.ItemData(Item)
    lstReferral.Visible = False
    If clsReferral.ReferralID > 0 Then Get_Referral clsReferral.ReferralID
    cctReferral.Text = clsReferral.Physician
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


Private Sub txtDate_Change()
Dim i%
Dim sDate As String
Dim strCurrentDate As String
    Clear_Auth
    If Trim(cboTime.Text) <> vbNullString Then
        udtAppointment.SchedTime = Format(cboTime.Text, "hh:mm AM/PM")
        If IsNull(txtDate.Value) Then
            MsgBox "Please select date!", vbInformation, "PatientTrac Patient Scheduling"
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            sDate = Format(txtDate.Value, "mm/dd/yyyy")
            strCurrentDate = Format(Now, "mm/dd/yyyy")
            sDate = DateDiff("d", strCurrentDate, sDate)
            If Val(sDate) < 0 Then
                MsgBox "Invalid Appointment Date, occurs prior to today's date!", vbInformation, "PatientTrac Patient Scheduling"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            If Validate_ProvAppt(sDate, cboDuration.Text) Then
                Exit Sub
            End If
        End If
    End If
    Get_InsAuth
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cctSearch_Change()
Dim sItem As String
    mbooByPass = False
    'mbooSearch = False
    Exit Sub
End Sub

Private Sub cmdApply_Click()
Dim i%
Dim mbooReason As Boolean
Dim strPurpose As String
Dim strDuration As String
Dim mbooSave As Boolean

    On Error GoTo Error_Apply
    Screen.MousePointer = vbHourglass
    udtSchedPatient.IntId = Val(lblID2.Caption)
    If Trim(cboTime.Text) = vbNullString Then
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
    For i = 0 To lstPurVisit.listcount - 1
        If lstPurVisit.Selected(i) Then
            mbooReason = True
            strPurpose = vbNullString
            Exit For
        End If
    Next i
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
    If udtAppointment.SchedType = "Patient" Then
        If mbooCancel = False Then
            If strPurpose <> vbNullString And strDuration = vbNullString Then
                MsgBox "Please select a Purpose for the visit!", vbInformation, "PatientTrac Appointment Scheduling"
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf strPurpose <> vbNullString And strDuration <> vbNullString Then
                MsgBox "Please select a Purpose for the visit and the Duration of the Appointment first!", vbInformation, "PatientTrac Appointment Scheduling"
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf strPurpose = vbNullString And strDuration <> vbNullString Then
                MsgBox "Please first select the Duration for the appointment.", vbInformation, "PatientTrac Appointment Scheduling"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
    mstrValidAppoint = vbNullString
    Check_TimeBlocks
    If Not Is_NewAppt Then
        If Validate_ProvAppt(cboTime.Text, cboDuration.Text) Then
            mstrValidAppoint = vbNullString: mbooAlreadyBook = False: mbooCancel = False
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    If mstrValidAppoint = "Authorized" Then
        mbooSave = True
    ElseIf mstrValidAppoint = "Invalid Appointment" Then
        MsgBox "Unable to schedule appointment!", vbInformation, "PatientTrac Appointment Scheduling"
        mstrValidAppoint = vbNullString: mbooAlreadyBook = False: mbooCancel = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If mbooAlreadyBook = False Then
        mbooSave = True
    Else
        mstrValidAppoint = vbNullString: mbooAlreadyBook = False: mbooCancel = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not mbooCancel Then
        mbooSave = True
    Else
        mstrValidAppoint = vbNullString: mbooAlreadyBook = False: mbooCancel = False
        Screen.MousePointer = vbDefault
        Exit Sub
   End If
   If mbooSave Then SaveData
    frmPatientAppointment.mBeenHere = True
    Screen.MousePointer = vbDefault
    mstrValidAppoint = vbNullString: mbooAlreadyBook = False: mbooCancel = False
    fraInfo.Visible = False
    Clear_PatientStrings
    Screen.MousePointer = vbDefault
    Unload Me
'    frmPatientAppointment.Show
'    centerForm frmPatientAppointment
'    If frmPatientAppointment.WindowState = 1 Then
'        frmPatientAppointment.WindowState = 0
'    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Apply:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Apply", App.EXEName & "." & Me.Name
    Exit Sub
End Sub

Private Function SaveData() As Boolean
Dim vTime
Dim Arr
Dim lngProviderID As Long
Dim lngFacilityID As Long
Dim strSQL As String
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
Dim strDuration As String
Dim strPatientType As String
Dim lngApptID As Long
Dim lngDateDiff As Long
Dim strReferral As String
Dim lngReferral As Long
Dim strNow As String

    On Error GoTo Error_SaveData
    Screen.MousePointer = vbHourglass
    PsUserID = Get_PsUserID(psConnect(0))
    strDuration = Trim(cboDuration.Text)
    If Trim(lblID2.Caption) <> vbNullString Then udtSchedPatient.IntId = Trim(lblID2.Caption)
    If Trim(udtSchedPatient.IntId) <> vbNullString Then
        If Trim(cboFacility.Text) = vbNullString Then
            MsgBox "Please reselect a facility!", vbInformation, "PatientTrac Appointment Scheduling"
            Screen.MousePointer = vbDefault
            Exit Function
        ElseIf cboFacility.ListIndex >= 0 Then
            lngFacilityID = cboFacility.ItemData(cboFacility.ListIndex)
        End If
        If udtAppointment.ProviderID > 0 Then
            lngProviderID = udtAppointment.ProviderID
        ElseIf Trim(cboProvider.Text) <> vbNullString Then
            lngProviderID = cboProvider.ItemData(cboProvider.ListIndex)
        End If
        For i = 0 To lstReferral.listcount - 1
            If lstReferral.Selected(i) Then
                lngReferral = lstReferral.ItemData(i)
                Exit For
            End If
        Next i
        udtAppointment.SchedTime = Format(cboTime.Text, "hh:mm AM/PM")
        strReferral = Replace(cctReasReferral.Text, "'", "''")
        If Trim(cboStatus.Text) = "Rescheduled" Then
            Rescheduled_Patient
            Exit Function
        ElseIf Trim(cboStatus.Text) = "Cancelled" Then
            Cancel_Patient
            Exit Function
        End If
        lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), txtDate.Value)
        If lngDateDiff >= 0 And Not IsNull(txtDate.Value) Then
            strDate = Format(txtDate.Value, "mm/dd/yyyy")
        Else
            strDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
        End If
        strTime = Trim(cboTime.Text)
        If Trim(strTime) = vbNullString Then
            MsgBox "Unable to save an appointment without a time value!", vbInformation, "PatientTrac Patient Scheduling"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        strTime = Format(strTime, "hh:mm AM/PM")
        strNotes = Trim(txtComments.Text)
        strNotes = Replace(strNotes, "'", "''")
        If Trim(cboPatientType.Text) <> vbNullString Then strPatientType = Trim(cboPatientType.Text)
        strConfirm = Format(txtApptConfirm.Value, "mm/dd/yyyy")
        For i = 0 To lstPurVisit.listcount - 1
            If lstPurVisit.Selected(i) Then
                strPurpose = lstPurVisit.List(i)
            End If
        Next i
        strStatus = Trim(cboStatus.Text)
        strReason = Trim(cboReason.Text)
        Open_DB psConnect(0)
        strSQL = "Get__Patient_Appt_ID " & lngProviderID & ", " & lngFacilityID & ", " & Val(udtSchedPatient.IntId) & ",'" _
            & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "', '" & Format(udtAppointment.SchedTime, "hh:mm AM/PM") & "'"
        Set rs = cn.Execute(strSQL)
        With rs
            If Not .EOF Then
                If Not IsNull(!APPT_ID) Then
                    lngApptID = (!APPT_ID)
                End If
            End If
        End With
        Open_DB psConnect(0)
        If lngApptID = 0 And udtSchedPatient.IntId <> vbNullString Then
            strSQL = "Write_Appointment " & Val(udtSchedPatient.IntId) & ", " & lngFacilityID & "," & lngProviderID & ",'" & strPatientType & "','" & strPurpose & "','" & Format(strDate, "mm/dd/yyyy") & "','" _
                & Format(strTime, "hh:mm AM/PM") & "','" & strNotes & "','" & strConfirm & "','" & strStatus & "','" & strReason & "','" & strRepeatAppt & "','" & strDuration & "',1," _
                & lngReferral & ",'" & strReferral & "', '" & udtAppointment.SchedType & "','" & PsUserID & "' "
            Set rs = cn.Execute(strSQL)
        ElseIf udtSchedPatient.IntId <> vbNullString Then
            strNow = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
            strSQL = "Update_Appointments  '" & Format(strDate, "mm/dd/yyyy") & "', '" & strPurpose & "', '" & strDuration & "','" _
                & strNotes & "', '" & Format(strTime, "hh:mm AM/PM") & "', '" & strPatientType & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "','" _
                & strConfirm & "', '" & strStatus & "', '" & strReason & "', '" & strRepeatAppt & "', " _
                & lngReferral & ", '" & strReferral & "', '" & udtAppointment.SchedType & "', '" & PsUserID & "', " & lngApptID
            Set rs = cn.Execute(strSQL)
        End If
        Close_DB psConnect(0)
    End If
    If Trim(cboStatus.Text) = vbNullString Then udtAppointment.Status = "New Appointment"
    Screen.MousePointer = vbDefault
    Exit Function
Error_SaveData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "SaveData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
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

    On Error GoTo Error_Get_Info
    Screen.MousePointer = vbHourglass
    lngProviderID = vlngProviderID
    If vlngPatientID > 0 Then lngPatientID = vlngPatientID
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
    If cboPatientType.listcount <= 0 Then Load_Visit
    'cboTime.Text = Format(Now, "hh:mm AM/PM")
    If mbLookupPatient = True Then
        If udtAppointment.SchedDate <> vbNullString And udtAppointment.SchedTime <> vbNullString Then
            txtDate.Value = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
           ' ccdTime.Value = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
            For i = 0 To cboTime.listcount - 1
                If cboTime.List(i) = udtAppointment.SchedTime Then
                    cboTime.Text = cboTime.List(i)
                    Exit For
                End If
            Next i
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
    If cboProvider.listcount <= 0 Then Get_Appt_Providers
    For i = 0 To cboProvider.listcount - 1
         If cboProvider.ItemData(i) = udtAppointment.ProviderID Then
             cboProvider.Text = cboProvider.List(i)
             Exit For
         End If
    Next i
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
    strSQL = "Get_PT_Appointment " & lngPatientID & ", " & lngProviderID & ", " & lngFacilityID & ", '" _
        & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "', '" & Format(udtAppointment.SchedTime, "hh:mm AM/PM") & "'"
        udtAppointment.VisitType = vbNullString
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            If Not .EOF Then
                If Not IsNull((!APPT_TIME)) Then
                    If (!APPT_TIME) <> vbNullString Then
                        If Format((!APPT_TIME), "hh:mm AM/PM") = udtAppointment.SchedTime Then
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
                            If udtAppointment.SchedType <> vbNullString And udtAppointment.SchedType <> "Patient" Then
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
                            'Load_Visit
                            For i = 0 To lstPurVisit.listcount - 1
                                If lstPurVisit.List(i) = udtAppointment.VisitType Then
                                    mbooByPass = True
                                    lstPurVisit.Selected(i) = True
                                    Exit For
                                End If
                            Next i
                             If udtAppointment.SchedTime <> vbNullString And udtAppointment.SchedDate <> vbNullString Then
                                 lblAppt2.Caption = Format(udtAppointment.SchedDate, "MMMM DD, YYYY") & " " & Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                             End If
                             If Trim(udtSchedPatient.ApptComment) <> vbNullString Then txtComments.Text = Trim(udtSchedPatient.ApptComment)
                             mUpdate = True
                             Select Case udtAppointment.SchedType
                                Case "Patient"
                                    Load_Patient udtSchedPatient.IntId
                                Case "Physician"
                                    Read_ReferralData udtSchedPatient.IntId
                                Case "Vendor"
                                    Read_VendorData udtSchedPatient.IntId
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
                             If udtAppointment.SchedType = "Patient" Then Get_Referrals
                             cctReferalReas.Text = Trim(clsReferral.Consult_Reason)
                        End If
                    End If
                End If
            Else
                booMisMatch = True
            End If
    End With
    If booMisMatch Then
        Get_ApptRange lngPatientID, lngProviderID
    End If
    Get_RepeatData lngFacilityID, lngProviderID
   If cnSched.State = adStateOpen Then
        cnSched.Close
        Set cnSched = Nothing
        Set rsSched = Nothing
    End If
    mBeenHere = True
    If cboDuration.Text = vbNullString Then
        For i = 0 To cboDuration.listcount - 1
            If frmPatientAppointment.mInterval = "EVERY HALF HOUR" Then
                If cboDuration.List(i) = "30 minutes" Then cboDuration.Text = cboDuration.List(i)
            ElseIf frmPatientAppointment.mInterval = "EVERY 15 MIN" Then
                If cboDuration.List(i) = "15 minutes" Then cboDuration.Text = cboDuration.List(i)
            End If
        Next i
    End If
     mbooByPass = False
    If Trim(udtAppointment.SchedType) <> vbNullString And IsProperCase(udtAppointment.SchedType) <> "Patient" Then
        Load_Visit udtAppointment.SchedType
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
    Resume
End Sub

Public Function Load_Duration()
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset

    On Error GoTo Error_Load_Duration
    Screen.MousePointer = vbHourglass
    cboDuration.Clear
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(1)
    strSQL = "Get_Appt_Time"
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
    Resume
End Function
Private Sub cmdClose_Click()
    frmPatientAppointment.mBeenHere = True
    If cboStatus.Text = vbNullString Then udtAppointment.Status = "New Appointment"
    Clear_PatientStrings
    Unload Me
'    centerForm frmPatientAppointment
'    If frmPatientAppointment.WindowState = 1 Then
'        frmPatientAppointment.WindowState = 0
'    End If
    Exit Sub
End Sub


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
    Resume
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
    Resume
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
    Resume
End Sub
Private Sub Form_Activate()
    Dim i%
    On Error GoTo Error_Activate
    centerForm Me
    'lblID2.Caption = vbNullString
    cboProvider.Enabled = True
    cboProvider.Locked = False
    cmdApply.Enabled = True
    Exit Sub
Error_Activate:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Activate", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub Form_Load()
Dim Idx%
Dim i%

'    mFirst = True
    On Error GoTo Error_Load
    Screen.MousePointer = vbHourglass
    cctReferral.Text = vbNullString
    cctReferalReas.Text = vbNullString
    cctReferalReas.Text = vbNullString
    cctReasReferral.Text = vbNullString
    lblID2.Caption = vbNullString
    cboProvider.Locked = False
    cboProvider.Enabled = True
    cboReferBy.Clear
    lstReferral.Visible = False
    lstReferral.Clear
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
    ElseIf udtAppointment.SchedType = "Patient" Then
       cboApptType.ListIndex = 0
    ElseIf udtAppointment.SchedType = "Physician" Then
       cboApptType.ListIndex = 1
    ElseIf udtAppointment.SchedType = "Vendor" Then
       cboApptType.ListIndex = 2
    End If
    lblCode.Caption = vbNullString
    lblCode.Visible = False
    udtAppointment.Status = vbNullString
    If mbLookupPatient And Val(udtSchedPatient.IntId) <> 0 Then
       Set_PatientFrame Val(udtSchedPatient.IntId)
       If Trim(cboSerPlace.Text) <> vbNullString And lstPurVisit.listcount <= 0 Then Load_Visit
       If udtAppointment.SchedType = "Patient" Then
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
            For i = 0 To cboTime.listcount - 1
                If cboTime.List(i) = udtAppointment.SchedTime Then
                    cboTime.Text = cboTime.List(i)
                    Exit For
                End If
            Next i
       End If
    Else
       If IsNull(txtDate.Value) Then
          txtDate.Value = udtAppointment.SchedDate
       End If
       If udtAppointment.SchedTime <> vbNullString Then
            For i = 0 To cboTime.listcount - 1
                If cboTime.List(i) = udtAppointment.SchedTime Then
                    cboTime.Text = cboTime.List(i)
                    Exit For
                End If
            Next i
       End If
    End If
    udtRepeatAppt.WeekDay = Format(udtAppointment.SchedDate, "dddd")
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
    cmdApply.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    mstrTerms = Empty
    mbLookupPatient = False
End Sub

Private Function Validate_Appointment()
Dim i%
Dim strSQL As String
Dim vstrTime As String
Dim lngProviderID As Long
Dim sValue As String
Dim intCount As Integer
Dim IAnswer As Integer
Dim strTime As String
Dim strApptDate As String

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
    strTime = cboTime.Text
    strApptDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    If mbooCancel Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    strSQL = "Validate_Appts_Provider " & udtAppointment.ProviderID & ", " & piFacilityID & ", '" & Format(strApptDate, "mm/dd/yyyy") & "', '" & Format(strTime, "hh:mm AM/PM") & "'"
    Open_DB psConnect(0)
    Set rs = cn.Execute(strSQL)
    intCount = 0
    With rs
        Do Until .EOF
            If Not IsNull(!Patient_ID) Then intCount = intCount + 1
        .MoveNext
        Loop
    End With
    sValue = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "MaxBookPatient")
    If sValue <> vbNullString And Not mbLookupPatient Then
        If intCount > Val(sValue) Then
            IAnswer = MsgBox("Is there an  authorization for this patient? Appointment exceeds practice policy for overbooking. Current Patients  at this time  are" & " " & intCount, vbYesNo)
            If IAnswer = vbYes Then
                mstrValidAppoint = "Authorized"
                Screen.MousePointer = vbDefault
            Else
                mstrValidAppoint = "Invalid Appointment"
                MsgBox "Please reschedule another time!", vbInformation
                mbooOverBook = True
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
    End If
    If cn.State = adStateClosed Then cn.open psConnect(0)
    strSQL = "Get_PT_Appt_Time " & Val(udtSchedPatient.IntId) & ", " & lngProviderID & ", '" & Format(strApptDate, "mm/dd/yyyy") & "'"
    Set rs = cn.Execute(strSQL)
    With rs
        If Not .EOF Then
            If Not IsNull(!APPT_TIME) Then
                If (!APPT_TIME) <> vbNullString Then
                    mstrOrigApptTime = Trim((!APPT_TIME))
                    If strTime <> Trim((!APPT_TIME)) Then
                        IAnswer = MsgBox("Patient has an appointment at" & " " & (!APPT_TIME) & " " & "on" & " " & strApptDate & _
                            "do you want to reschedule appointment to" & " " & strTime & "?", vbYesNoCancel)
                        If IAnswer = vbYes Then
                            Rescheduled_Patient
                            mbooAlreadyBook = True
                            Screen.MousePointer = vbDefault
                        Else
                            mbooAlreadyBook = True
                            Screen.MousePointer = vbDefault
                            Exit Function
                        End If
                   End If
                End If
            End If
        End If
    End With
    Close_DB psConnect(0)
    Screen.MousePointer = vbDefault
     Exit Function
Error_Validate_Appointment:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_Appointment", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Check_TimeBlocks()
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim varCellValue
Dim vDate
Dim IAnswer As Integer
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
    
    If Trim(cboTime.Text) <> vbNullString Then strTime = cboTime.Text

    strRange = DateAdd("d", -7, udtAppointment.SchedDate)
    If Trim(strTime) = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
        strSQL = "Get_SchedConflict " & udtAppointment.ProviderID & ", " & udtFacility.FacID & ", '" & Format(strRange, "mm/dd/yyyy") & "'"
'        strSQL = "SELECT * FROM APPOINTMENTS_BLOCKS WHERE ACTIVE = 1 AND PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
'            & "FACILITY_ID = " & udtFacility.FacID & " And DateDiff(Day, '" & strRange & "', START_DATE) > 0"
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
                                            mstrValidAppoint = "Invalid Appointment"
                                            Exit Do
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
    Resume
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
    strSQL = "Get_Superbill_Referral_PatientID" & Val(udtSchedPatient.IntId)
    Set rsReferral = cnReferral.Execute(strSQL)
    With rsReferral
        Do Until .EOF
            If Not IsNull(!Referral_ID) Then
                If clsReferral.ReferralID > 0 Then
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
    Else
        strSQL = "SELECT  REFERRER_ID, REFERRAL_REASON FROM PATIENT_ORDERS  WHERE  PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
            & "ENCOUNTER_ID = " & piEncounterID
        Set rsReferral = cnReferral.Execute(strSQL)
        
        With rsReferral
            Do Until .EOF
                If Not IsNull(!REFERRER_ID) Then
                    clsReferral.ReferralID = (!REFERRER_ID): lngReferralID = clsReferral.ReferralID
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
    If lngReferralID > 0 Then
        strSQL = "SELECT  REFERRAL_ID, REFERRAL_REASON FROM PATIENT_ACCOUNTS  WHERE  PATIENT_ID = " & Val(udtSchedPatient.IntId)
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
    Resume
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

    On Error GoTo Error_Check_ApptReferral
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
        Load_Specific_Referral clsReferral.ReferralID
        Get_Referral clsReferral.ReferralID
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
    Resume
End Function

Private Function Calc_Duration(vstrReason As String)
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset

    On Error GoTo Error_Calc_Duration
    Screen.MousePointer = vbHourglass
    
    clsGridFormat.List = vbNullString
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(1)
    strSQL = "Get_Appt_Time_APPT_CODE '" & vstrReason & "'"
    Set rsAppt = cnAppt.Execute(strSQL)
    With rsAppt
        If Not .EOF Then
            If (!APPT_TIME) <> vbNullString Then udtEncounter.VisitTime = (!APPT_TIME)
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
    Resume
End Function

