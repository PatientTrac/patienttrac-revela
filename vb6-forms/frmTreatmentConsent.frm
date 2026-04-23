VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmTreatmentConsent 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10500
   ClientLeft      =   6480
   ClientTop       =   1815
   ClientWidth     =   13140
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTreatmentConsent.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   13140
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame fraConsent 
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
      Height          =   10395
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   13095
      Begin PatientTrac.ccApprove ccApprove1 
         Height          =   795
         Left            =   8760
         TabIndex        =   49
         Top             =   7800
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1402
         ccApproveBackColor=   14737632
         Enabled         =   0   'False
      End
      Begin PatientTrac.ccApprove ccAppWitness 
         Height          =   555
         Left            =   8790
         TabIndex        =   48
         Top             =   8700
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   979
      End
      Begin VB.Frame fraReceiver 
         Caption         =   "RECIPIENT INFORMATION"
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
         Height          =   3915
         Left            =   4170
         TabIndex        =   34
         Top             =   2400
         Visible         =   0   'False
         Width           =   4785
         Begin VB.TextBox cctRecPhone2 
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
            Height          =   360
            Left            =   3090
            MaxLength       =   4
            TabIndex        =   45
            Top             =   2610
            Width           =   825
         End
         Begin VB.TextBox cctRecPh1 
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
            Height          =   360
            Left            =   2190
            MaxLength       =   3
            TabIndex        =   44
            Top             =   2610
            Width           =   825
         End
         Begin VB.TextBox cctRecAC 
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
            Height          =   360
            Left            =   1410
            MaxLength       =   3
            TabIndex        =   41
            Top             =   2610
            Width           =   645
         End
         Begin VB.TextBox cctRecCityState 
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
            Height          =   435
            Left            =   1410
            TabIndex        =   37
            Top             =   1980
            Width           =   3075
         End
         Begin VB.TextBox cctRecAddress 
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
            Height          =   375
            Left            =   1410
            TabIndex        =   36
            Top             =   1395
            Width           =   3075
         End
         Begin VB.TextBox cctRecName 
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
            Height          =   360
            Left            =   1410
            TabIndex        =   35
            Top             =   840
            Width           =   3075
         End
         Begin PT_XP_Button.PT_XP cmdRecClose 
            Height          =   315
            Left            =   1890
            TabIndex        =   43
            Top             =   3420
            Width           =   1095
            _ExtentX        =   1931
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
            Picture         =   "frmTreatmentConsent.frx":058A
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmTreatmentConsent.frx":0F9C
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdLookUp 
            Height          =   465
            Left            =   3060
            TabIndex        =   47
            Top             =   210
            Width           =   1455
            _ExtentX        =   2566
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
            Picture         =   "frmTreatmentConsent.frx":12EE
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "LookUp Provider"
            DefaultThemes   =   2
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmTreatmentConsent.frx":1CE8
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblRecTele 
            Caption         =   "Telephone"
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
            Left            =   150
            TabIndex        =   42
            Top             =   2670
            Width           =   915
         End
         Begin VB.Label lblRecCityState 
            Caption         =   "City, State"
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
            Left            =   150
            TabIndex        =   40
            Top             =   2040
            Width           =   915
         End
         Begin VB.Label lblRecAdd 
            Caption         =   "Address"
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
            Left            =   150
            TabIndex        =   39
            Top             =   1500
            Width           =   915
         End
         Begin VB.Label lblRecName 
            Caption         =   "Name"
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
            Left            =   150
            TabIndex        =   38
            Top             =   930
            Width           =   825
         End
      End
      Begin VB.ComboBox cboPatient 
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
         Height          =   345
         ItemData        =   "frmTreatmentConsent.frx":266A
         Left            =   5400
         List            =   "frmTreatmentConsent.frx":2674
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   8100
         Width           =   2085
      End
      Begin VB.PictureBox picImage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   5325
         Left            =   4500
         ScaleHeight     =   5265
         ScaleWidth      =   4725
         TabIndex        =   24
         Top             =   1230
         Visible         =   0   'False
         Width           =   4785
         Begin PT_XP_Button.PT_XP cmdDefaultImg 
            Height          =   465
            Left            =   2730
            TabIndex        =   25
            Top             =   120
            Visible         =   0   'False
            Width           =   1905
            _ExtentX        =   3360
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
            Picture         =   "frmTreatmentConsent.frx":268C
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Default Path"
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmTreatmentConsent.frx":3086
            ShowFocus       =   -1  'True
         End
         Begin VB.Frame fraImage 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Select Form"
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
            Height          =   4035
            Left            =   150
            TabIndex        =   26
            Top             =   360
            Width           =   4515
            Begin VB.FileListBox filNavigator 
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
               Height          =   1890
               Left            =   120
               TabIndex        =   29
               ToolTipText     =   "Select to Preview Form"
               Top             =   1920
               Width           =   4230
            End
            Begin VB.DirListBox dirNavigator 
               Height          =   1170
               Left            =   120
               TabIndex        =   28
               Top             =   660
               Width           =   4230
            End
            Begin VB.DriveListBox drvNavigator 
               Height          =   360
               Left            =   120
               TabIndex        =   27
               Top             =   270
               Width           =   2310
            End
         End
         Begin PT_XP_Button.PT_XP cmdGetForm 
            Height          =   465
            Left            =   930
            TabIndex        =   30
            Top             =   4710
            Width           =   1395
            _ExtentX        =   2461
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
            Picture         =   "frmTreatmentConsent.frx":3A08
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Accept"
            CaptionPosition =   4
            DefaultThemes   =   2
            DepthMode       =   1
            DepthEvent      =   1
            ForeColorDisabled=   -2147483629
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmTreatmentConsent.frx":3DA2
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdVisible 
            Height          =   465
            Left            =   2490
            TabIndex        =   31
            Top             =   4710
            Width           =   1395
            _ExtentX        =   2461
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
            Picture         =   "frmTreatmentConsent.frx":40F4
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmTreatmentConsent.frx":4B06
            ShowFocus       =   -1  'True
         End
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
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   8160
         Width           =   3735
      End
      Begin VB.ListBox lstConsents 
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
         Height          =   1110
         Left            =   120
         TabIndex        =   14
         Top             =   8640
         Width           =   4755
      End
      Begin VB.ComboBox cboStaff 
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
         Left            =   8790
         TabIndex        =   10
         Top             =   9330
         Width           =   3285
      End
      Begin VB.TextBox txtWitnessName 
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
         Height          =   345
         Left            =   5490
         TabIndex        =   8
         Top             =   12450
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.PictureBox picWitness 
         AutoRedraw      =   -1  'True
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
         Height          =   525
         Left            =   8790
         MousePointer    =   11  'Hourglass
         ScaleHeight     =   23.25
         ScaleMode       =   2  'Point
         ScaleWidth      =   159.75
         TabIndex        =   5
         Top             =   8700
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.PictureBox picSignature 
         AutoRedraw      =   -1  'True
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
         Height          =   465
         Left            =   8760
         MousePointer    =   11  'Hourglass
         ScaleHeight     =   20.25
         ScaleMode       =   2  'Point
         ScaleWidth      =   159.75
         TabIndex        =   3
         Top             =   7890
         Visible         =   0   'False
         Width           =   3255
      End
      Begin MSComDlg.CommonDialog ComDlg 
         Left            =   1785
         Top             =   1905
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox rtfBox 
         Height          =   6885
         Left            =   120
         TabIndex        =   0
         Top             =   645
         Width           =   12795
         _ExtentX        =   22569
         _ExtentY        =   12144
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         RightMargin     =   5
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmTreatmentConsent.frx":4E58
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PT_XP_Button.PT_XP cmdAccept 
         Height          =   345
         Left            =   5220
         TabIndex        =   11
         Top             =   9870
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
         Picture         =   "frmTreatmentConsent.frx":4EDE
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Accept"
         CaptionAlignment=   5
         CaptionPosition =   4
         DefaultThemes   =   2
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTreatmentConsent.frx":5278
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   345
         Left            =   6360
         TabIndex        =   12
         Top             =   9870
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
         Picture         =   "frmTreatmentConsent.frx":55CA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTreatmentConsent.frx":5FDC
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdTerms 
         Height          =   315
         Left            =   12150
         TabIndex        =   13
         ToolTipText     =   "Uses Same Terms as Patient Interviewed With"
         Top             =   9330
         Visible         =   0   'False
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
         Picture         =   "frmTreatmentConsent.frx":632E
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
         PictureDisabled =   "frmTreatmentConsent.frx":68C8
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPatient 
         Height          =   495
         Left            =   180
         TabIndex        =   18
         Top             =   90
         Width           =   1785
         _ExtentX        =   3149
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
         Picture         =   "frmTreatmentConsent.frx":6C1A
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "SELECT PATIENT"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTreatmentConsent.frx":71B4
         PictureAlignment=   0
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdMedConsent 
         Height          =   525
         Left            =   2040
         TabIndex        =   19
         Top             =   60
         Width           =   1785
         _ExtentX        =   3149
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
         Picture         =   "frmTreatmentConsent.frx":7506
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Medications"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTreatmentConsent.frx":7DE0
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdGeneral 
         Height          =   525
         Left            =   3900
         TabIndex        =   20
         Top             =   60
         Width           =   1785
         _ExtentX        =   3149
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
         Picture         =   "frmTreatmentConsent.frx":8A32
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Practice General"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTreatmentConsent.frx":930C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdEducation 
         Height          =   525
         Left            =   5760
         TabIndex        =   21
         Top             =   60
         Width           =   1785
         _ExtentX        =   3149
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
         Picture         =   "frmTreatmentConsent.frx":9F5E
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Patient Education"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTreatmentConsent.frx":A838
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrint 
         Height          =   525
         Left            =   11070
         TabIndex        =   22
         ToolTipText     =   "PRINT"
         Top             =   60
         Width           =   1785
         _ExtentX        =   3149
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
         Picture         =   "frmTreatmentConsent.frx":B48A
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Print"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTreatmentConsent.frx":BD64
         PictureAlignment=   5
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdInform 
         Height          =   525
         Left            =   7650
         TabIndex        =   23
         Top             =   30
         Width           =   1785
         _ExtentX        =   3149
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
         Picture         =   "frmTreatmentConsent.frx":C9B6
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Inform Consent"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTreatmentConsent.frx":D290
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSearchName 
         Height          =   435
         Left            =   3240
         TabIndex        =   32
         ToolTipText     =   "Select Search Type"
         Top             =   7650
         Width           =   1815
         _ExtentX        =   3201
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
         Picture         =   "frmTreatmentConsent.frx":DEE2
         AppearanceMode  =   1
         BackColorDown   =   15133676
         TransparentColor=   14737632
         Caption         =   "Search Forms"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmTreatmentConsent.frx":E65C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdReciever 
         Height          =   495
         Left            =   9540
         TabIndex        =   46
         ToolTipText     =   "Recipient Data"
         Top             =   60
         Width           =   1335
         _ExtentX        =   2355
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
         TransparentColor=   14737632
         Caption         =   "Receipient"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblSearch 
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
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
         Left            =   150
         TabIndex        =   17
         Top             =   7890
         Width           =   975
      End
      Begin VB.Label lblMeds 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Education/Consents"
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
         Left            =   -240
         TabIndex        =   15
         Top             =   7890
         Width           =   4215
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   8010
         TabIndex        =   9
         Top             =   9420
         Width           =   645
      End
      Begin VB.Label lblWitness 
         BackStyle       =   0  'Transparent
         Caption         =   "Witness"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5430
         TabIndex        =   7
         Top             =   8640
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sign here:"
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
         Left            =   7620
         TabIndex        =   6
         Top             =   8640
         Width           =   975
      End
      Begin VB.Label lblPatient 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient/Guardian"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5400
         TabIndex        =   4
         Top             =   7860
         Width           =   2385
      End
      Begin VB.Label lblSignHere 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sign here:"
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
         Left            =   7620
         TabIndex        =   2
         Top             =   7890
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTreatmentConsent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbolErased As Boolean
Private mbolform As Boolean
Private mbolSig As Boolean
Private mbolGetSig As Boolean
Private mbolActivate As Boolean
Private mobjCheckBox As Control
Public mlngPatientID As Long
Public lEncounterID As Long
Public lCompID As Long
Public lProviderID As Long
Public psConsentFilePath As String
Public mobjCmd As Control
Public lngCurrColIndex As Long
Public lngCurrIndex As String
Private mbooPrintPreview As Boolean
Private mSaveConsent As Boolean
Public mstrConsent As String
Private mFirst As Boolean
Private flngDeltaX As Long
Private flngDeltaY As Long
Private fintPreviewIndex As Integer
Const BLOCK_SIZE = 16384

Private udtReciever As udtReceiver
Private Type udtReceiver
    Name As String
    Address As String
    CityState As String
    Telephone As String
End Type

Private Const AnInch As Long = 1440   '1440 twips per inch
Private Const QuarterInch As Long = 360

Const Signature = "PSig"
Const Patient = "PName"
Const TodayDate = "PDate"
Const VisitDate = "pVisitDate"
Const WitnessSign = "WSig"
Const WitnessName = "WName"
Const Facility = "PFacility"
Const FacPhone = "pFacPhone"
Const Provider = "pProvider"
Const Receiver = "pReceiver"
Const MRN = "pEMR"
Const DOB = "pDOB"
Const Guardian = "pGuard"
Const GuardName = "pGuardName"
Const FacAdd = "pFacAdd"
Private Sub Clear_Reciepient()
    With udtReciever
        .Name = vbNullString
        .Address = vbNullString
        .CityState = vbNullString
        .Telephone = vbNullString
    End With
End Sub

Private Function Load_Consent()
Dim strSQL As String
Dim i%
Dim rsConsent As New ADODB.Recordset
Dim cnConsent As New ADODB.Connection

    On Error GoTo Error_Load_Consent
    Screen.MousePointer = vbHourglass
    
    If cnConsent.State = adStateClosed Then cnConsent.open psConnect(1)
    strSQL = "SELECT CONSENT_NAME, FORM_ID FROM PRACTICE_CONSENTS"
    Set rsConsent = cnConsent.Execute(strSQL)
    With rsConsent
        Do Until .EOF
            If Trim((!CONSENT_NAME)) <> vbNullString Then lstConsents.AddItem (!CONSENT_NAME)
            If IsNumeric((!FORM_ID)) Then lstConsents.ItemData(lstConsents.NewIndex) = (!FORM_ID)
        .MoveNext
        Loop
    End With
    If cnConsent.State = adStateClosed Then
        cnConsent.Close
        Set rsConsent = Nothing
        Set cnConsent = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Consent:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Consent", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Function Load_Rtf(vstrConsent As String)
Dim cnConsent As New ADODB.Connection
Dim rsConsent As New ADODB.Recordset
Dim strSQL As String
Dim strFile As String
Dim strPath As String
Dim LoadStream As ADODB.Stream
Dim strConsent As String
Dim strFullName As String
Dim Fld As ADODB.Field
Dim F As Long, Data() As Byte, FileSize As Long
Dim strFileText As String
Dim lngPrimKey As Long
Dim myStream As New ADODB.Stream

    On Error GoTo Error_Load_Rtf
    Screen.MousePointer = vbHourglass
    If cnConsent.State = adStateClosed Then cnConsent.open psConnect(1)
    vstrConsent = Replace(vstrConsent, ".rtf", "")
    rsConsent.open "SELECT * FROM PRACTICE_CONSENTS WHERE CONSENT_NAME = '" & vstrConsent & "'", cnConsent, adOpenKeyset, adLockOptimistic
    'strSQL = "SELECT * FROM PRACTICE_CONSENTS WHERE CONSENT_NAME = '" & Trim(vstrConsent) & "'"
    'Set rsConsent = cnConsent.Execute(strSQL)
    With rsConsent
        Do Until .EOF
            Set myStream = New ADODB.Stream
            myStream.Type = adTypeText
            myStream.open
            If (!CONSENT_NAME) <> vbNullString Then
                strFile = (!CONSENT_NAME)
            End If
            If (!FILE_PATH) <> vbNullString Then
                strPath = (!FILE_PATH)
            End If
            If (!FORM_ID) > 0 Then
                lngPrimKey = (!FORM_ID)
            End If
            If strFile <> vbNullString And strPath <> vbNullString Then
                strFullName = strPath & "\" & strFile & ".rtf"
                If FileExists(strFullName) Then
                     myStream.WriteText strFullName
                     If Dir("C:\TempConsent.txt") <> "" Then
                         Kill "C:\TempConsent.txt"
                     End If
                    myStream.SaveToFile ("C:\TempConsent.txt")
                     rtfBox.LoadFile strFullName
                     rsConsent.Update
                End If
            End If
        rsConsent.MoveNext
        Loop
    End With
        
'    'Populate Report
'    Const sKeyWord6 = "PFacility"
'    Const sKeyWord7 = "pProvider"
    Clipboard.Clear
    rtfBox.HideSelection = False

     'Provider Name
    rtfBox.Find Provider ' , lintPos, lintPos + Provider
    If rtfBox.SelStart > 0 Then rtfBox.SelText = clsProvider.Name
    rtfBox.SelItalic = True

    rtfBox.SelStart = 0
    rtfBox.Find Facility ', lintPos, lintPos + Facility
    rtfBox.SelText = clsFacility.Name
    rtfBox.SelItalic = True
    
    rtfBox.SelStart = 0
    rtfBox.Find FacPhone ' , lintPos, lintPos + FPhone
    If rtfBox.SelStart > 0 Then rtfBox.SelText = clsFacility.Telephone
    rtfBox.SelItalic = True
    
    Word_Wrap
    If cnConsent.State = adStateClosed Then
        cnConsent.Close
        Set rsConsent = Nothing
        Set cnConsent = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Rtf:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Rtf", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Function Read_Consent(vstrConsent As String)
Dim strSQL As String
Dim cnRead As New ADODB.Connection
Dim rsRead As New ADODB.Recordset
Dim LoadStream As ADODB.Stream
Dim myStream As ADODB.Stream
Dim Fld As ADODB.Field
Dim Flds As ADODB.Fields
Dim strConsent As String
Dim strFullName As String
Dim F As Long, Data() As Byte, FileSize As Long
Dim strFileText As String
Dim lngPrimKey As Long
Dim strFile As String
Dim strPath As String

    On Error GoTo Error_Read_Consent
    Screen.MousePointer = vbHourglass
    myStream.Type = adTypeText
    If cnRead.State = adStateClosed Then cnRead.open psConnect(0)
    'rs.open "SELECT * FROM CONSENT_FORM WHERE FORM =  '" & vstrConsent & "' AND PATIENT_ID = " & mlngPatientID & ", cn, adOpenForwardOnly, adLockReadOnly"
    strSQL = "SELECT * FROM CONSENT_FORM WHERE FORM =  '" & vstrConsent & "' AND PATIENT_ID = " & mlngPatientID
    Set rsRead = cnRead.Execute(strSQL)
    myStream.open
    Set Flds = rsRead.Fields
    For Each Fld In Flds
       If Fld.Name = "FORM_IMAGE" Then
          If Not IsNull(rsRead.Fields("CONSENT_FORM").Value) Then myStream.Write rsRead.Fields("CONSENT_FORM").Value
           myStream.SaveToFile ("C:\TempConsent.txt")
           myStream.Close
       End If
    Next
    If cnRead.State = adStateOpen Then
       cnRead.Close
       Set myStream = Nothing
       Set cnRead = Nothing
       Set rsRead = Nothing
    End If
    rtfBox.LoadFile "C:\TempConsent.txt"
    Add_Providers
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Consent:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Consent", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Private Sub PrintDoc()
Dim i%
Dim PrintableWidth As Long
Dim PrintableHeight As Long
Dim X As Single

    On Error GoTo Error_Consent_Print
    Screen.MousePointer = vbHourglass
    mbooCancelPrint = False
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation
        Exit Sub
    End If
        If Printer.Copies > 1 And mbooPrintPreview Then
            For i = 1 To Printer.Copies
                rtfBox.SelPrint Printer.hdc
            Next i
        Else
            rtfBox.SelPrint Printer.hdc
        End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Consent_Print:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Consent_Print", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Function Read_Data()
Dim i%
Dim sDate As String

    rtfBox.Text = vbNullString
    Load_Consent
    mSaveConsent = False
    ccApprove1.ClearSignature
    mbolActivate = True
    If piPatientID > 0 Then Get_Patient piPatientID
    sDate = Format(Now, "MMMM DD, YYYY")
    Me.Caption = clsPatient.Name & " " & "Date of Birth" & " " & clsPatient.DOB & " " & "Date:" & " " & sDate
    Add_Providers
    Get_Facilities
    Get_Providers
    If mstrConsent <> vbNullString Then
        For i = 0 To lstConsents.ListCount - 1
            If lstConsents.List(i) = Trim(mstrConsent) Then
                lstConsents.Text = lstConsents.List(i)
                Exit For
            End If
        Next i
    End If
    
    Exit Function
End Function

Private Sub cboStaff_Change()
    txtWitnessName.Text = cboStaff.Text
End Sub

Private Sub cctSearch_KeyUp(KeyCode As Integer, Shift As Integer)

    SearchList lstConsents, cctSearch.Text, True
    Exit Sub
End Sub


Private Sub cmdAccept_Click()
   If Not ValidateData Then Exit Sub
   Call WriteSignature
   mSaveConsent = True
   cmdPatient.Enabled = True
    Exit Sub
End Sub

Private Sub WriteSignature()
Dim lintPos As Long
Dim lstrFile As String
Dim lstrDate As String
Dim lstrName As String
Dim lstrLastName As String
Dim lstrFirstName As String
Dim lstrMiddleName As String
Dim lstrReturnMsg As String
Dim lintRet As Integer
Dim lstrSQL As String
Dim lstrMsg As String
Dim varRS()
Dim lReturn As Integer
Dim intPrevPos As Long
Dim intMargin As Integer
Dim strStartPosPat As String
Dim strStarPosWit As String
   
   On Error GoTo Error_WriteSignature
   Screen.MousePointer = vbHourglass

   Clipboard.Clear
    Get_Facilities
    Get_Providers
    lstrDate = Trim(Format(Now, "mmmm dd, yyyy"))
    rtfBox.HideSelection = False

    rtfBox.Find Date  ' , lintPos, lintPos + PDate
    If rtfBox.SelStart > 0 Then rtfBox.SelText = lstrDate
    rtfBox.SelItalic = True
    
    rtfBox.Find Facility, lintPos, lintPos + Facility
    If rtfBox.SelStart > 0 Then rtfBox.SelText = clsFacility.Name
    rtfBox.SelItalic = True
 
   lstrFile = "C:\Program Files\FirstNotebook\PatientTrac EMR" & "\sig" & Format(Now, "HHMMSS") & ".bmp"
 

   If FileExists(lstrFile) Then
        SetAttr lstrFile, vbNormal
   End If

   picSignature.Picture = ccApprove1.Picture
   SavePicture picSignature.Image, lstrFile
   
   Clipboard.SetData LoadPicture(lstrFile), vbCFBitmap

    rtfBox.SelUnderline = True
    lintPos = rtfBox.Find(Signature)

'Inject Signature......................................................
    rtfBox.SelUnderline = True
    If lintPos > 0 Then rtfBox.OLEObjects.Add , , lstrFile
    
    intPrevPos = lintPos
    rtfBox.SelUnderline = True

    'Patient Name
    lstrName = clsPatient.Name
   
   If Len(Patient) - Len(lstrName) > 1 Then
        intMargin = (Len(Patient) - Len(lstrName)) / 2
        lstrName = String(intMargin, "_") & lstrName & String(intMargin, "_")
   End If

    lintPos = rtfBox.Find(Patient, lintPos - 1)
    rtfBox.SelItalic = True
    rtfBox.SelUnderline = True
    
    If lintPos > 0 Then rtfBox.SelText = lstrName
    If lintPos > intPrevPos Then intPrevPos = lintPos Else lintPos = intPrevPos
  
    
'Get first date location................................
   lintPos = rtfBox.Find(Date, lintPos + Len(Date) + 1)
   rtfBox.SelItalic = True
   rtfBox.SelUnderline = True
   
   'Inject Date....................................
   If lintPos > 0 Then rtfBox.SelText = Space(1) & lstrDate & Space(1)
   If lintPos > intPrevPos Then intPrevPos = lintPos Else lintPos = intPrevPos

    
   picWitness.Picture = ccAppWitness.Picture
   SavePicture picWitness.Image, lstrFile
   
'Get 2nd Signature location
   lintPos = rtfBox.Find(WitnessSign, lintPos + 1)
   
    'Inject 2nd Signature ......................................................
   If lintPos > 0 Then rtfBox.OLEObjects.Add , , lstrFile            ' ??? was commented out
   
   If lintPos > intPrevPos Then intPrevPos = lintPos Else lintPos = intPrevPos
   
   rtfBox.SelStart = lintPos + 1
   rtfBox.SelText = "    "
   
'Get 2nd date location................................
   lintPos = rtfBox.Find("WitnessName", lintPos + 1)
   rtfBox.SelItalic = True
   rtfBox.SelUnderline = True
   txtWitnessName.Text = cboStaff.Text
     
   'Inject 2nd Date................................
   If lintPos > 0 Then rtfBox.SelText = txtWitnessName.Text
   
    rtfBox.SelItalic = True
    rtfBox.SelUnderline = True
   
   If lintPos > 0 Then
        rtfBox.Find ("Patient's Name")
   End If
   
   If lintPos > intPrevPos Then intPrevPos = lintPos Else lintPos = intPrevPos

   rtfBox.SelStart = lintPos + 500
   rtfBox.HideSelection = True
   Kill lstrFile
   Screen.MousePointer = vbDefault
   Call UpdateSignature

    ' Save RTF to PATIENT_CONSENT TABLE (Write_Consent)
    Screen.MousePointer = vbDefault
    Exit Sub
Error_WriteSignature:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "picWitness.Image", "WriteSignature", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Function Write_Consent()
Dim cnWrite As New ADODB.Connection
Dim rsWrite As New ADODB.Recordset
Dim strSQL As String
Dim strFile As String
Dim strPath As String
Dim LoadStream As ADODB.Stream
Dim strConsent As String
Dim strFullName As String
Dim Fld As ADODB.Field
Dim F As Long, Data() As Byte, FileSize As Long
Dim strFileText As String
Dim lngPrimKey As Long
Dim myStream As New ADODB.Stream
Dim booinsert As Boolean
Dim strDate As String

    On Error GoTo Error_Write_Consent
    Screen.MousePointer = vbHourglass
    strConsent = Trim(lstConsents.Text)
    booinsert = True
    strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If cnWrite.State = adStateClosed Then cnWrite.open psConnect(0)
    rsWrite.open "SELECT CONSENT_ID FROM CONSENT_FORMS WHERE FORM = '" & strConsent & "' AND PATIENT_ID = " & mlngPatientID & " ", cnWrite, adOpenKeyset, adLockOptimistic
    With rsWrite
        If Not .EOF Then
            If Not IsNull(!CONSENT_ID) Then
                booinsert = False
                lngPrimKey = (!CONSENT_ID)
            End If
        Else
            booinsert = True
        End If
    End With
            
    If booinsert Then
        strSQL = "INSERT INTO CONSENT_FORMS (PATIENT_ID, FORM, ACCEPTED, INSERT_DATE) " _
            & "values(" & mlngPatientID & ",'" & strConsent & "', 1,'" & strDate & "')"
    
        Set rsWrite = cnWrite.Execute(strSQL)
        strSQL = "SELECT CONSENT_ID FROM CONSENT_FORMS WHERE FORM = '" & strConsent & "' AND PATIENT_ID = " & mlngPatientID & ""
        Set rsWrite = cnWrite.Execute(strSQL)
        With rsWrite
            If Not .EOF Then
                If Not IsNull(!CONSENT_ID) Then
                    lngPrimKey = (!CONSENT_ID)
                End If
            End If
        End With
    Else
        strSQL = "UPDATE CONSENT_FORMS SET FORM = '" & strConsent & "', ACCEPTED = 1," _
            & "UPDATE_DATE = '" & strDate & "' WHERE CONSENT_ID = " & lngPrimKey
        Set rsWrite = cnWrite.Execute(strSQL)
    End If
    If rsWrite.State = adStateOpen Then rsWrite.Close
        'rsWrite.open "SELECT FORM_IMAGE FROM CONSENT_FORMS WHERE CONSENT_ID =  " & lngPrimKey & " ", cnWrite, adOpenKeyset, adLockOptimistic
        strSQL = "SELECT FORM_IMAGE FROM CONSENT_FORMS WHERE CONSENT_ID =  " & lngPrimKey
        Set rsWrite = cnWrite.Execute(strSQL)
                Set myStream = New ADODB.Stream
                myStream.Type = adTypeText
                myStream.open
                rtfBox.SaveFile ("C:\TempConsent.txt")
                If Not IsNull(rsWrite.Fields("FORM_IMAGE")) Then
                    rsWrite.Fields("FORM_IMAGE").Value = myStream.ReadText
                Else
                    myStream.LoadFromFile ("C:\TempConsent.txt")
                    rsWrite.Fields("FORM_IMAGE").Value = myStream.ReadText
                End If
            rsWrite.Update
    If cnWrite.State = adStateClosed Then
        cnWrite.Close
        Set rsWrite = Nothing
        Set cnWrite = Nothing
        Set myStream = Nothing
    End If
    cmdPatient.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Consent:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Consent", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cmdClose_Click()
    If mSaveConsent Then
        Write_Consent
    End If
    Unload Me
End Sub

Private Sub cmdEducation_Click()
Dim strSQL As String
Dim i%
Dim rsConsent As New ADODB.Recordset
Dim cnConsent As New ADODB.Connection
Dim lngCnt As Long

    On Error GoTo Error_Load_Consent
    Screen.MousePointer = vbHourglass
    
    If cnConsent.State = adStateClosed Then cnConsent.open psConnect(1)
    strSQL = "SELECT CONSENT_NAME, FORM_ID FROM PRACTICE_CONSENTS WHERE CONSENT_TYPE = 'Patient_Education'"
    Set rsConsent = cnConsent.Execute(strSQL)
    With rsConsent
        Do Until .EOF
            If lngCnt = 0 Then lstConsents.Clear
            lngCnt = lngCnt + 1
            If Trim((!CONSENT_NAME)) <> vbNullString Then lstConsents.AddItem (!CONSENT_NAME)
            If IsNumeric((!FORM_ID)) Then lstConsents.ItemData(lstConsents.NewIndex) = (!FORM_ID)
        .MoveNext
        Loop
    End With
    If cnConsent.State = adStateClosed Then
        cnConsent.Close
        Set rsConsent = Nothing
        Set cnConsent = Nothing
    End If
    cmdPatient.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Consent:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Consent", App.EXEName & "." & Me.Name
    Exit Sub
    Resume

End Sub

Private Sub cmdGeneral_Click()
Dim strSQL As String
Dim i%
Dim rsConsent As New ADODB.Recordset
Dim cnConsent As New ADODB.Connection
Dim lngCnt As Long

    On Error GoTo Error_Load_Consent
    Screen.MousePointer = vbHourglass
    
    If cnConsent.State = adStateClosed Then cnConsent.open psConnect(1)
    strSQL = "SELECT CONSENT_NAME, FORM_ID FROM PRACTICE_CONSENTS WHERE CONSENT_TYPE = 'Practice_General'"
    Set rsConsent = cnConsent.Execute(strSQL)
    With rsConsent
        Do Until .EOF
            If lngCnt = 0 Then lstConsents.Clear
            lngCnt = lngCnt + 1
            If Trim((!CONSENT_NAME)) <> vbNullString Then lstConsents.AddItem (!CONSENT_NAME)
            If IsNumeric((!FORM_ID)) Then lstConsents.ItemData(lstConsents.NewIndex) = (!FORM_ID)
        .MoveNext
        Loop
    End With
    If cnConsent.State = adStateClosed Then
        cnConsent.Close
        Set rsConsent = Nothing
        Set cnConsent = Nothing
    End If
    cmdPatient.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Consent:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Consent", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdGetForm_Click()
Dim strFilePath As String
    Screen.MousePointer = vbHourglass
    If Right$(filNavigator.Path, 1) <> "\" Then
        strFilePath = filNavigator.Path & "\" & filNavigator.FileName
    Else
        strFilePath = filNavigator.Path & filNavigator.FileName
    End If
    If FileExists(strFilePath) Then
        Load_Rtf filNavigator.FileName
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cmdInform_Click()
Dim strSQL As String
Dim i%
Dim rsConsent As New ADODB.Recordset
Dim cnConsent As New ADODB.Connection
Dim lngCnt As Long

    On Error GoTo Error_Load_Consent
    Screen.MousePointer = vbHourglass
    
    If cnConsent.State = adStateClosed Then cnConsent.open psConnect(1)
    strSQL = "SELECT CONSENT_NAME, FORM_ID FROM PRACTICE_CONSENTS WHERE CONSENT_TYPE = 'Patient_Inform'"
    Set rsConsent = cnConsent.Execute(strSQL)
    With rsConsent
        Do Until .EOF
            If lngCnt = 0 Then lstConsents.Clear
            lngCnt = lngCnt + 1
            If Trim((!CONSENT_NAME)) <> vbNullString Then lstConsents.AddItem (!CONSENT_NAME)
            If IsNumeric((!FORM_ID)) Then lstConsents.ItemData(lstConsents.NewIndex) = (!FORM_ID)
        .MoveNext
        Loop
    End With
    If cnConsent.State = adStateClosed Then
        cnConsent.Close
        Set rsConsent = Nothing
        Set cnConsent = Nothing
    End If
    cmdPatient.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Consent:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Consent", App.EXEName & "." & Me.Name
    Exit Sub
    Resume

End Sub

Private Sub cmdMedConsent_Click()
Dim strSQL As String
Dim i%
Dim rsConsent As New ADODB.Recordset
Dim cnConsent As New ADODB.Connection
Dim lngCnt As Long

    On Error GoTo Error_Load_Consent
    Screen.MousePointer = vbHourglass
    
    If cnConsent.State = adStateClosed Then cnConsent.open psConnect(1)
    strSQL = "SELECT CONSENT_NAME, FORM_ID FROM PRACTICE_CONSENTS WHERE CONSENT_TYPE = 'Medication'"
    Set rsConsent = cnConsent.Execute(strSQL)
    With rsConsent
        Do Until .EOF
            If lngCnt = 0 Then lstConsents.Clear
            lngCnt = lngCnt + 1
            If Trim((!CONSENT_NAME)) <> vbNullString Then lstConsents.AddItem (!CONSENT_NAME)
            If IsNumeric((!FORM_ID)) Then lstConsents.ItemData(lstConsents.NewIndex) = (!FORM_ID)
        .MoveNext
        Loop
    End With
    If cnConsent.State = adStateClosed Then
        cnConsent.Close
        Set rsConsent = Nothing
        Set cnConsent = Nothing
    End If
    Screen.MousePointer = vbDefault
   cmdPatient.Enabled = True
    Exit Sub
Error_Load_Consent:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Consent", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub cmdPatient_Click()
    Clear_Reciepient
    frmListPatient.Read_PatientData
    frmListPatient.Show vbModal
    If frmListPatient.mlngListItemData > 0 Then piPatientID = frmListPatient.mlngListItemData
    Read_Data
    Exit Sub
End Sub

Private Sub cmdPrint_Click()
   'If Not ValidateData Then Exit Sub
   PrintDoc
   cmdPatient.Enabled = True
End Sub

Private Sub cmdRecClose_Click()
    If cctRecName.Name <> vbNullString Then
        udtReciever.Name = Trim(cctRecName.Name)
        udtReciever.Name = vbNullString
    End If
    If cctRecAddress.Text <> vbNullString Then
        udtReciever.Address = Trim(cctRecAddress.Text)
    End If
    If cctRecAddress.Text <> vbNullString Then
        udtReciever.Address = Trim(cctRecAddress.Text)
    End If
    If cctRecCityState.Text <> vbNullString Then
        udtReciever.CityState = Trim(cctRecCityState.Text)
    End If
    If cctRecAC.Text <> vbNullString Then
        udtReciever.Telephone = Trim(cctRecAC.Text)
    End If
    If cctRecPh1.Text <> vbNullString Then
        udtReciever.Telephone = "(" & Trim(udtReciever.Telephone) & ")" & Trim(cctRecPh1.Text)
    End If
    If cctRecPhone2.Text <> vbNullString Then
        udtReciever.Telephone = udtReciever.Telephone & "-" & Trim(cctRecPhone2.Text)
    End If
    
    If udtReciever.Name <> vbNullString Then
        rtfBox.Find Receiver  ' , lintPos, lintPos + PDate
        If rtfBox.SelStart > 0 Then rtfBox.SelText = udtReciever.Name & vbNewLine & udtReciever.Address & vbNewLine & udtReciever.CityState & vbNewLine & udtReciever.Telephone
        rtfBox.SelItalic = True
    End If
    fraReceiver.Visible = False
    Exit Sub
End Sub

Private Sub cmdReciever_Click()
    fraReceiver.Visible = True
End Sub
Private Sub cmdSearchName_Click()
    picImage.Visible = True
   cmdPatient.Enabled = True
End Sub

Private Sub cmdTerms_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Patient Interview"
    frmNurseTemplates.fType = "Patient_Interview"
    frmNurseTemplates.fSubDomain = "sd_Psych_Interview"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    GetTerms cboStaff, "sd_Psych_Interview"
End Sub

Private Sub cmdVisible_Click()
        picImage.Visible = False
End Sub

Private Sub dirNavigator_Change()
    filNavigator.Path = dirNavigator.Path
End Sub

Private Sub dirNavigator_KeyUp(KeyCode As Integer, Shift As Integer)
        filNavigator.Path = dirNavigator.Path
End Sub


Private Sub filNavigator_Click()
Dim intTIFpages As Integer
Dim i As Integer
Dim strFilePath As String
    If mboobypass Then mboobypass = False: Exit Sub
    If Right$(filNavigator.Path, 1) <> "\" Then
        strFilePath = filNavigator.Path & "\" & filNavigator.FileName
    Else
        strFilePath = filNavigator.Path & filNavigator.FileName
    End If
    Exit Sub
End Sub


Private Sub Form_Activate()
Dim i%
Dim sDate As String

    On Error GoTo Error_Activate
    Clear_Reciepient
    If mFirst Then
        mFirst = False
    End If
    If mbolActivate Then Exit Sub
    If piPatientID > 0 Then
        Read_Data
        cmdPatient.Enabled = False
    End If
    centerForm Me
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Activate:
    Exit Sub
    Resume
End Sub
Private Sub Form_Load()
    mbolActivate = False
    mFirst = True
    mbolSig = False
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
  Exit Sub
End Sub

Private Function ValidateData() As Boolean
   ValidateData = (Not ccApprove1.EmptySignature)
   If Not ValidateData Then
     MsgBox "Enter signature and proceed..."
     mbolSig = False
   Else
    mbolSig = True
   End If
End Function

Public Sub SetTreatmentConsentFile()
    'dnn............
    If Not Read_Signature Then
        rtfBox.LoadFile pobjConsentColl(lngCurrColIndex).FileName
    End If
End Sub

Public Sub UpdateSignature()
   Dim lstrSignature As String
   lstrSignature = ccApprove1.EncodeSignatureValues(".")
End Sub

Private Function Read_Signature() As Boolean
Dim strFile As String
Dim strPatientFile As String
Dim strPath As String
Dim lintFiles As Integer
Dim lstrFile() As String
Dim i As Integer


    'Identify file .................
   strPath = pobjConsentColl(lngCurrColIndex).FileName
   strFile = Get_FileName(strPath)
   strPath = Get_FilePath(strPath) & "\SignedRtf\"
   lintFiles = Read_DirFiles(strPath & "*.*", lstrFile())
    
    'Look for...............
    strPatientFile = Replace(strFile, ".", "_" & mlngPatientID & "_" & lEncounterID & ".")
    strFile = strPath & strPatientFile
    'Check In Form-English_1029_2.rtf
    
    If lintFiles > 0 Then
        For i = LBound(lstrFile) To UBound(lstrFile)
            If strPatientFile = lstrFile(i) Then
                rtfBox.LoadFile strFile
                Read_Signature = True
                Exit For
            End If
        Next i
    End If
    Add_Providers
    Exit Function
End Function


Private Function Add_Providers()
Dim strSQL As String
Dim cnProv As New ADODB.Connection
Dim rsProv As New ADODB.Recordset
Dim strProvider As String

    On Error GoTo Error_Add_Providers
    Screen.MousePointer = vbHourglass
    If cnProv.State = adStateClosed Then cnProv.open psConnect(0)
    strSQL = "Select FIRST_NAME, MIDDLE_NAME, LAST_NAME, SUFFIX, TITLE, PROVIDER_ID " _
        & "FROM PROVIDERS "
    Set rsProv = cnProv.Execute(strSQL)
    With rsProv
            Do Until .EOF
            If Not IsNull(!First_Name) Then strProvider = (!First_Name)
            If Not IsNull(!Middle_Name) Then strProvider = Trim(strProvider) & " " & (!Middle_Name)
            If Not IsNull(!Last_Name) Then strProvider = Trim(strProvider) & " " & (!Last_Name)
            If Trim((!Suffix)) <> vbNullString Or Trim((!Suffix)) <> vbNullString Then strProvider = Trim(strProvider) & ", " & (!Suffix)
            If Not IsNull(!Title) Then strProvider = Trim(strProvider) & ", " & (!Title)
            If strProvider <> vbNullString Then cboStaff.AddItem strProvider
            If IsNumeric((!PROVIDER_ID)) Then cboStaff.ItemData(cboStaff.NewIndex) = (!PROVIDER_ID)
            .MoveNext
            Loop
        End With
    If cnProv.State = adStateOpen Then
        cnProv.Close
        Set cnProv = Nothing
        If rsProv.State = adStateOpen Then rsProv.Close
        Set rsProv = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Add_Providers:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Add_Providers", App.EXEName & "." & "ModPatientData"
    Exit Function
    Resume
End Function

Private Sub lstConsents_Click()
Dim strSQL As String
Dim iAnswer As Integer
Dim strConsentName As String
Dim cnList As New ADODB.Connection
Dim rsList As New ADODB.Recordset

    On Error GoTo Error_lstConsents_Click
    Screen.MousePointer = vbHourglass
    strConsentName = Trim(lstConsents.Text)
    If cnList.State = adStateClosed Then cnList.open psConnect(0)
    strSQL = "SELECT CONSENT_ID, INSERT_DATE FROM CONSENT_FORMS WHERE PATIENT_ID = " & piPatientID & " AND FORM = '" & strConsentName & " '"
    Set rsList = cnList.Execute(strSQL)
    With rsList
        If Not .EOF Then
            If piPatientID > 0 Then
                iAnswer = MsgBox("Patient has signed this consent on." & " " & (!INSERT_DATE) & " " & ". Do you want to execute a new consent.", vbYesNoCancel + vbQuestion + vbApplicationModal + vbDefaultButton1, "Patient Informed Consents")
                Select Case iAnswer
                    Case vbYes:
                      Load_Rtf strConsentName
                    Case vbNo:
                      Read_Consent strConsentName
                    Case vbCancel:
                      Exit Sub
                End Select
            Else
                Load_Rtf strConsentName
            End If
        Else
            Load_Rtf strConsentName
        End If
    End With
    If cnList.State = adStateClosed Then
        cnList.Close
        Set rsList = Nothing
        Set cnList = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_lstConsents_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "lstConsents_Click", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub Word_Wrap()
    Dim i As Byte
     SetViewMode ercWordWrap, rtfBox
End Sub

Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    flngDeltaX = X
    flngDeltaY = Y
End Sub


Private Sub picImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picImage.Left = picImage.Left + X - flngDeltaX
        picImage.Top = picImage.Top + Y - flngDeltaY
    End If
End Sub



