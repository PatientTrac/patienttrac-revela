VERSION 5.00
Object = "{2A5B1681-9AFA-4521-B035-B9AD0C75F381}#1.0#0"; "DedicatedControls.ocx"
Object = "{C08D2A54-E5CB-4AFE-8AAD-B3E08A43486C}#1.3#0"; "XButton.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmMDIReport 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Report"
   ClientHeight    =   9285
   ClientLeft      =   5895
   ClientTop       =   3180
   ClientWidth     =   12705
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMDIReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   12705
   Begin VB.FileListBox lstReports 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2400
      Pattern         =   "*.txt"
      TabIndex        =   30
      ToolTipText     =   "Please select desired report. "
      Top             =   90
      Width           =   3585
   End
   Begin VB.PictureBox picPatientSel 
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
      Height          =   8205
      Left            =   120
      ScaleHeight     =   8145
      ScaleWidth      =   8565
      TabIndex        =   7
      Top             =   1050
      Visible         =   0   'False
      Width           =   8625
      Begin VB.Frame fraMultiPage 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Multiple Report Print"
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
         Height          =   2595
         Left            =   90
         TabIndex        =   22
         Top             =   4800
         Visible         =   0   'False
         Width           =   8205
         Begin VB.Frame fraEnctr 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Encounter Dates"
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
            Height          =   2355
            Left            =   2460
            TabIndex        =   23
            Top             =   150
            Width           =   4755
            Begin DedicatedControls.ccDate ccdDateTo 
               Height          =   510
               Left            =   2490
               TabIndex        =   47
               Top             =   240
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   900
               Caption         =   "TO"
               CaptionPosition =   1
               Text            =   ""
               BackColor       =   14737632
               DateBackColor   =   -2147483643
               ForeColor       =   8388608
               BackStyle       =   0
               FontName        =   "Georgia"
               FontBold        =   -1  'True
            End
            Begin VB.Frame fraEnctrFrom 
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
               Height          =   1545
               Left            =   90
               TabIndex        =   26
               Top             =   780
               Visible         =   0   'False
               Width           =   2985
               Begin VB.ListBox lstEnctrFrom 
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   825
                  Left            =   150
                  Style           =   1  'Checkbox
                  TabIndex        =   27
                  Top             =   270
                  Width           =   2625
               End
               Begin PT_XP_Button.PT_XP cmdFEnctSel 
                  Height          =   255
                  Left            =   660
                  TabIndex        =   42
                  Top             =   1200
                  Width           =   795
                  _ExtentX        =   1402
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
                  Caption         =   "Select"
                  DepthEvent      =   1
                  ForeColorOver   =   8388608
                  ForeColorFocus  =   8388608
                  ForeColorDown   =   8388608
                  ShowFocus       =   -1  'True
               End
               Begin PT_XP_Button.PT_XP cmdFClose 
                  Height          =   255
                  Left            =   1500
                  TabIndex        =   44
                  Top             =   1200
                  Width           =   795
                  _ExtentX        =   1402
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
                  Caption         =   "Close"
                  DepthEvent      =   1
                  ForeColorOver   =   8388608
                  ForeColorFocus  =   8388608
                  ForeColorDown   =   8388608
                  ShowFocus       =   -1  'True
               End
            End
            Begin XBtn.XButton cmdSelTo 
               Height          =   255
               Left            =   4170
               TabIndex        =   32
               Top             =   510
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Georgia"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderColor     =   8388608
               Bevel           =   6
               Caption         =   "...."
               Picture         =   "frmMDIReport.frx":0442
               PictAlign       =   4
               CaptionAlign    =   4
               PushTimer       =   0   'False
               PushInterval    =   300
               MousePointer    =   0
               PicRaster       =   12583114
            End
            Begin XBtn.XButton cmdSelFrom 
               Height          =   255
               Left            =   1800
               TabIndex        =   31
               Top             =   480
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   450
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Georgia"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderColor     =   8388608
               Bevel           =   6
               Caption         =   "...."
               Picture         =   "frmMDIReport.frx":045E
               PictAlign       =   4
               CaptionAlign    =   4
               PushTimer       =   0   'False
               PushInterval    =   300
               MousePointer    =   0
               PicRaster       =   12583114
            End
            Begin DedicatedControls.ccCheck chkAllEnctr 
               Height          =   195
               Left            =   300
               TabIndex        =   25
               Top             =   1200
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   344
               Caption         =   "All Encounters"
               CaptionPosition =   4
               ForeColor       =   8388608
               FontName        =   "Georgia"
               FontBold        =   -1  'True
               BackColor       =   14737632
            End
            Begin VB.Frame fraEnctrTo 
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
               Height          =   1545
               Left            =   1710
               TabIndex        =   28
               Top             =   780
               Visible         =   0   'False
               Width           =   2985
               Begin VB.ListBox lstEnctrTo 
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   825
                  Left            =   150
                  Style           =   1  'Checkbox
                  TabIndex        =   29
                  Top             =   270
                  Width           =   2625
               End
               Begin PT_XP_Button.PT_XP cmdTEnctSel 
                  Height          =   255
                  Left            =   630
                  TabIndex        =   41
                  Top             =   1170
                  Width           =   795
                  _ExtentX        =   1402
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
                  Caption         =   "Select"
                  DepthEvent      =   1
                  ForeColorOver   =   8388608
                  ForeColorFocus  =   8388608
                  ForeColorDown   =   8388608
                  ShowFocus       =   -1  'True
               End
               Begin PT_XP_Button.PT_XP cmdTClose 
                  Height          =   255
                  Left            =   1470
                  TabIndex        =   43
                  Top             =   1170
                  Width           =   795
                  _ExtentX        =   1402
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
                  Caption         =   "Close"
                  DepthEvent      =   1
                  ForeColorOver   =   8388608
                  ForeColorFocus  =   8388608
                  ForeColorDown   =   8388608
                  ShowFocus       =   -1  'True
               End
            End
            Begin XBtn.XButton cmdPrintMulti 
               Height          =   435
               Left            =   1830
               TabIndex        =   33
               Top             =   1620
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   767
               BackColor       =   12640511
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Georgia"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderColor     =   4210752
               Bevel           =   6
               Caption         =   "PRINT"
               Picture         =   "frmMDIReport.frx":047A
               PictAlign       =   4
               CaptionAlign    =   4
               PushTimer       =   0   'False
               PushInterval    =   300
               MousePointer    =   0
               PicRaster       =   12583114
            End
            Begin DedicatedControls.ccDate ccdDateFrom 
               Height          =   510
               Left            =   270
               TabIndex        =   48
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   900
               Caption         =   "FROM"
               CaptionPosition =   1
               Text            =   ""
               BackColor       =   14737632
               DateBackColor   =   -2147483643
               ForeColor       =   8388608
               BackStyle       =   0
               FontName        =   "Georgia"
               FontBold        =   -1  'True
            End
         End
         Begin VB.Label lblPrintPat 
            BackColor       =   &H00FFFFFF&
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
            Height          =   405
            Left            =   120
            TabIndex        =   24
            Top             =   780
            Width           =   3195
         End
      End
      Begin VB.Frame fraFilter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search Filter"
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
         Height          =   1155
         Left            =   4620
         TabIndex        =   11
         Top             =   210
         Width           =   3735
         Begin VB.OptionButton optByPatientName 
            BackColor       =   &H00E0E0E0&
            Caption         =   "By Last Name"
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
            Left            =   150
            TabIndex        =   15
            Top             =   300
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton optByExternalId 
            BackColor       =   &H00E0E0E0&
            Caption         =   "By External ID"
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
            Left            =   150
            TabIndex        =   14
            Top             =   840
            Width           =   1755
         End
         Begin VB.OptionButton optByInternalId 
            BackColor       =   &H00E0E0E0&
            Caption         =   "By Internal ID"
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
            Left            =   150
            TabIndex        =   13
            Top             =   570
            Width           =   1725
         End
         Begin VB.CheckBox chkWithEnctrs 
            BackColor       =   &H00E0E0E0&
            Caption         =   "w\Encounters"
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
            Left            =   1950
            TabIndex        =   12
            Top             =   540
            Width           =   1695
         End
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   10
         Top             =   360
         Width           =   3195
      End
      Begin VB.ListBox lstEncounters 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   4620
         TabIndex        =   9
         Top             =   1680
         Width           =   3735
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
         Height          =   3210
         Left            =   120
         TabIndex        =   8
         Top             =   1050
         Width           =   4305
      End
      Begin PT_XP_Button.PT_XP cmdMultPrint 
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   4410
         Width           =   2865
         _ExtentX        =   5054
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
         Picture         =   "frmMDIReport.frx":0496
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Multiple Print Job Mode"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMDIReport.frx":0A30
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSelCancel 
         Height          =   345
         Left            =   4410
         TabIndex        =   45
         Top             =   7590
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
         Picture         =   "frmMDIReport.frx":0D82
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Cancel"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMDIReport.frx":1794
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSelOK 
         Height          =   345
         Left            =   3420
         TabIndex        =   46
         Top             =   7590
         Width           =   765
         _ExtentX        =   1349
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
         Picture         =   "frmMDIReport.frx":1AE6
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "OK"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMDIReport.frx":1E80
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblCount 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3420
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl_Count 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Count"
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
         Left            =   3420
         TabIndex        =   19
         Top             =   150
         Width           =   585
      End
      Begin VB.Label lblHeaders 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Last Name, First Name"
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
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   780
         Width           =   4275
      End
      Begin VB.Label lblSelEncounter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encounters:"
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
         Left            =   4650
         TabIndex        =   17
         Top             =   1410
         Width           =   1170
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Search:"
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
         TabIndex        =   16
         Top             =   120
         Width           =   1500
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1800
      Top             =   8580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSignature 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7740
      ScaleHeight     =   315
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   3615
   End
   Begin DedicatedControls.ccApprove ccaSignature 
      Height          =   495
      Left            =   2610
      TabIndex        =   2
      Top             =   8600
      Visible         =   0   'False
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   873
      Enabled         =   0   'False
   End
   Begin FPSpreadADO.fpSpread grdMDIReport 
      Height          =   8475
      Left            =   30
      TabIndex        =   0
      Top             =   840
      Width           =   12645
      _Version        =   458752
      _ExtentX        =   22304
      _ExtentY        =   14949
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16638955
      MaxCols         =   10
      MaxRows         =   0
      ScrollBars      =   2
      ShadowColor     =   14737632
      SpreadDesigner  =   "frmMDIReport.frx":21D2
      TextTipDelay    =   0
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Frame fraCMD 
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
      Height          =   855
      Left            =   1830
      TabIndex        =   5
      Top             =   30
      Width           =   10935
      Begin VB.ComboBox cboForms 
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
         Left            =   8430
         TabIndex        =   6
         Top             =   450
         Width           =   2325
      End
      Begin PT_XP_Button.PT_XP cmdReports_Select 
         Height          =   315
         Left            =   90
         TabIndex        =   35
         Top             =   90
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmMDIReport.frx":23F2
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         CaptionPosition =   4
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmMDIReport.frx":298C
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrint 
         Height          =   315
         Left            =   4260
         TabIndex        =   36
         Top             =   90
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmMDIReport.frx":2CDE
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         CaptionPosition =   4
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   13003064
         ForeColorFocus  =   13003064
         ForeColorDown   =   13003064
         PictureDisabled =   "frmMDIReport.frx":3278
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdMult2 
         Height          =   315
         Left            =   4740
         TabIndex        =   37
         Top             =   90
         Width           =   2865
         _ExtentX        =   5054
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
         Picture         =   "frmMDIReport.frx":35CA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Multiple Print Job Mode"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMDIReport.frx":3B64
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   345
         Left            =   9780
         TabIndex        =   38
         Top             =   60
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
         Picture         =   "frmMDIReport.frx":3EB6
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMDIReport.frx":48C8
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRefresh 
         Height          =   345
         Left            =   8460
         TabIndex        =   39
         Top             =   60
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
         Picture         =   "frmMDIReport.frx":4C1A
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Refresh"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMDIReport.frx":51B4
         ShowFocus       =   -1  'True
      End
   End
   Begin PT_XP_Button.PT_XP cmdPatient 
      Height          =   495
      Left            =   90
      TabIndex        =   34
      Top             =   60
      Width           =   1485
      _ExtentX        =   2619
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
      Picture         =   "frmMDIReport.frx":5506
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Select Patient"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMDIReport.frx":5DE0
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblPatName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1980
      TabIndex        =   21
      Top             =   90
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.Label lblHidden 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "hidden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   570
      Visible         =   0   'False
      Width           =   885
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMDIReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private fclsEncounter As EncounterInfo
Private mlngFromDateEnctrID As Long
Private mlngToDateEnctrID As Long
Private fudtAudit() As BooleanAudit
Private fudtGridHeader() As GridHeader
Private fintRowInserts As Integer
Private fbooRun As Boolean
Public fstrTitle As String
Public fmstrfile As String
Private fstrSavePath As String
Private fstrIncludedObject() As String
Private fstrExcludedObject() As String
Public fstrLastFilePath As String
Private fstrScriptLine() As String
Private fbooSkip As Boolean
Private fbooRemRowNoData As Boolean
Private flngRowHidden() As Long
Private fudtRowInsert() As RowInsert
Private fintInsertedRows  As Integer
Private piPatientID As Long
Private piEncounterID As Long
Private mFirst As Boolean
Public mbooPrint As Boolean
Private Type RowInsert
    StartRow As Integer
    InsertedRows As Integer
End Type


Private Type BooleanAudit
    OCX As String
    Control As String
    Object As String
    DataPresent As Boolean
    ShowWhen As Boolean
End Type

Private Type GridHeader
    StartCol As Long
    EndCol As Long
    Row As Long
End Type

Private Type EncounterInfo
    StartDate As String
    EndDate As String
    VisitReason As String
    Facility As String
    Provider As String
    ServicePlace As String
    ServicePlaceCode As String
    EncounterType As String
End Type

Private Sub cboForms_Change()
    If cboForms.Text = "Browser" Then
        frmMDIReport.Hide
        frmBrowser.Show
    End If
End Sub


Private Sub cboForms_Click()
    If cboForms.Text = "Browser" Then
        frmMDIReport.Hide
        frmBrowser.Show
    End If
End Sub

Private Sub cmdClose_Click()
    Reset_Grid
    Unload Me
    
End Sub

Private Sub cmdFClose_Click()
    fraEnctrFrom.Visible = False
    lstPatient.Enabled = True
    lstEncounters.Enabled = True
End Sub

Private Sub cmdFEnctSel_Click()
    lstEnctrFrom_Click
    fraEnctrFrom.Visible = False
    lstPatient.Enabled = True
    lstEncounters.Enabled = True

    Exit Sub
End Sub

Private Sub cmdMult2_Click()
    picPatientSel.ZOrder 0
    picPatientSel.Visible = True
    picPatientSel.Top = 960
    picPatientSel.Left = 480
    fraMultiPage.Visible = True
    fraMultiPage.ZOrder 0
End Sub

Private Sub cmdMultPrint_Click()
    fraMultiPage.Visible = True
End Sub

Private Sub cmdPatient_Click()
    Screen.MousePointer = vbHourglass
    txtSearch.Text = ""
    lstPatient.Clear
    lstEncounters.Clear
    Load_Patients
    picPatientSel.Visible = True
    picPatientSel.ZOrder 0
    picPatientSel.Top = 960
    picPatientSel.Left = 480
    txtSearch.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Load_Patients()
Dim strSQL As String
Dim lstrRemarks As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim r As Integer
Dim c As Integer

LogFile "Private Sub Load_Patients-----------------------------------"

    Set vRS = Nothing
    strSQL = "Select patient_id, rtrim(last_name) last_name, rtrim(first_name) first_name " _
        & " From Patient  " _
        & "Order by rtrim(last_name),rtrim(first_name)"
    
    'iRows = Run_SelectQuery(psDB, strSQL, vData)
    bRet = SelectQuery(strSQL, psConnect(0), vRS, iRows)

    'lblCount.Refresh
    
If iRows > 0 Then
    For r = LBound(vRS, 2) To UBound(vRS, 2)
        lstPatient.AddItem vRS(1, r) & ", " & vRS(2, r)
        If IsNumeric(vRS(0, r)) Then lstPatient.ItemData(r) = vRS(0, r)
        If vRS(0, r) = piPatientID Then lstPatient.ListIndex = lstPatient.NewIndex
    Next r
        lblCount.Caption = lstPatient.ListCount
End If

    
End Sub



Private Sub cmdPrintMulti_Click()
Dim strSQL As String
Dim i As Integer
Dim lstrRemarks As String
Dim booRet As Boolean
Dim lngrecs As Long
Dim varRS
    On Error GoTo Error_cmdPrintMulti_Click
    
    picPatientSel.Enabled = False
    If mlngToDateEnctrID < 0 Then mlngToDateEnctrID = mlngFromDateEnctrID
    
    Screen.MousePointer = vbHourglass
    fmstrfile = fstrLastFilePath & "\" & lstReports.FileName
    If lstReports.FileName = vbNullString Then
        MsgBox "Please Select Report From Menu!", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
     
    If chkAllEnctr.Value Then
    
    strSQL = "Select encounter_id, patient_id  " _
        & " From ENCOUNTER " _
        & "Where patient_id = " & piPatientID
        
    Else
    
    strSQL = "Select encounter_id, patient_id  " _
        & " From ENCOUNTER " _
        & "Where encounter_id between " & mlngFromDateEnctrID & " And " & mlngToDateEnctrID & "  And " _
        & "patient_id = " & piPatientID
        
    End If
       
        booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
        
        If booRet Then
            If lngrecs Then
                For i = LBound(varRS, 2) To UBound(varRS, 2)
                    If varRS(1, i) > 0 Then
                        piEncounterID = varRS(0, i)
                         'Get_Provider_Id
                        If piProviderID = 0 Then
                            MsgBox "Patient does not have a Provider Selected for encounter" & varRS(0, i) & _
                                "Print Job Aborted, Please reselect valid encounter(s).", vbCritical
                                Screen.MousePointer = vbDefault
                            Exit Sub
                        Else
                            cmdRun_Click
                            cmdPrint_Click
                            grdMDIReport.Refresh
                        End If
                    End If
                Next i
            End If
        End If
        Screen.MousePointer = vbDefault
        fraMultiPage.Visible = False
        chkAllEnctr.Value = False
        picPatientSel.Enabled = True
        lstPatient.Enabled = True
        fraEnctr.Visible = False
        lstEncounters.Enabled = True
        Screen.MousePointer = vbDefault
   Exit Sub
        
Error_cmdPrintMulti_Click:
    Screen.MousePointer = vbDefault
    picPatientSel.Enabled = True
    fraMultiPage.Visible = False
    lstPatient.Enabled = True
    fraEnctr.Visible = False
    lstEncounters.Enabled = True
   ErrHandler _
     "frmMDIReport.cmdPrintMulti_Click", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
End Sub


Private Sub cmdRefresh_Click()
    grdMDIReport.MaxCols = 5
    grdMDIReport.MaxRows = 0
    fbooRun = True
    Get_Collection piPatientID
    Open_Script
    fbooRun = False
End Sub

Private Sub cmdReports_Select_Click()
    Reset
    fmstrfile = ""
    
    If fstrLastFilePath = "" Then
        fstrLastFilePath = pstrMainPath & "\Reports"
        If Not FolderExists(fstrLastFilePath) Then
            MkDir fstrLastFilePath
       End If
    End If
    
    fmstrfile = CommonDialogAction(CommonDialog, "Open", "Report Script", fstrLastFilePath, "*.txt")
    
    
    If fmstrfile <> vbNullString Then
       ' fstrReportContent = File_Method(fmstrFile, "open")
        cmdRun_Click
    End If

End Sub

Public Sub cmdPrint_Click()
    On Error GoTo Error_Print
    
    ShowPrinter Me.hwnd
    
    Print_Grid grdMDIReport, "Date: " & Now
    mbooPrint = True
    
    Exit Sub
    
Error_Print:
    MsgBox Err.Description, vbInformation, "Print"
    Exit Sub
End Sub

Public Sub cmdRun_Click()

    fbooRun = True
    frmMDIReport.Visible = False
    modSpecific.piPatientID = piPatientID
    modSpecific.piEncounterID = piEncounterID
    Get_Collection piPatientID
    Open_Script
    frmMDIReport.Height = grdMDIReport.Height + 1500
    frmMDIReport.Visible = True
    fbooRun = False

End Sub

Private Sub Reset()
    fstrAction = ""

    'shpSelect.Visible = False

    fbooBackColor = False
    fbooBorderOutline = False
    fbooBorderVert = False
    fbooBorderHoriz = False
    fbooBorderCells = False
    fbooBorderTop = False
    fbooBorderBottom = False
    fbooBorderLeft = False
    fbooBorderRight = False
    'fbooByPass = False
    'fbooByPass2 = False
    fbooCellSpan = False
    'fbooCellSpanExists = False
    'fbooChanged = False
    'fbooColWidthsExists = False
    'fbooEncounterInfoExists = False
    fbooErase = False
    'fbooEraseMode = False
    'fbooFetchBooleanCompExists = False
    'fbooFetchBooleanDataExists = False
    'fbooFetchDataExists = False
    'fbooFetchListExists = False
    fbooFont = False
    fbooHorizAlignLeft = False
    fbooHorizAlignCenter = False
    fbooHorizAlignRight = False
    fbooMergeCols = False
    fbooMergeRows = False
    fbooMoveRange = False
    fbooSkip = False
    fbooVertAlignTop = False
    fbooVertAlignCenter = False
    fbooVertAlignBottom = False
    fbooWordWrap = False
    
End Sub


Private Sub Reset_Grid()
        
    grdMDIReport.MaxCols = 0
    grdMDIReport.MaxRows = 0
    grdMDIReport.MaxCols = 15
    grdMDIReport.MaxRows = 50
    grdMDIReport.ClearRange 1, 1, grdMDIReport.MaxCols, grdMDIReport.MaxRows, False
    grdMDIReport.Width = 13695
    grdMDIReport.Height = 10000
End Sub


Private Sub cmdSelCancel_Click()
    picPatientSel.Visible = False
End Sub

Private Sub cmdSelFrom_Click()
Dim lstrSQL As String
Dim i As Integer
Dim lstrRemarks As String
Dim booRet As Boolean
Dim lngrecs As Long
Dim varRS

LogFile "Private Sub Get_Encounters-----------------------------------"

If piPatientID = 0 Then Exit Sub

    lstrSQL = "Select encntr_start_date, encounter_id " _
        & "From ENCOUNTER " _
        & "Where patient_id = " & piPatientID & " Order By ENCNTR_START_DATE desc"
        
         booRet = SelectQuery(lstrSQL, psConnect(0), varRS, lngrecs)
         Load_ComboBox lstEnctrFrom, lstrSQL, psConnect(0)
         
    fraEnctrFrom.Visible = True
    fraEnctrFrom.ZOrder 0
    lstPatient.Enabled = False
    lstEncounters.Enabled = False
    
    Exit Sub
End Sub

Private Sub cmdSelOK_Click()
    picPatientSel.Visible = False
    If piProviderID = 0 Then
        Exit Sub
    Else
        picPatientSel.Visible = False
        Reset
        cmdRun_Click
    End If
Exit Sub
End Sub
Private Sub Get_Provider_Id()
Dim strSQL As String
Dim varRS
Dim lngrecs As Long
Dim booRet As Boolean
Dim varMessage(2)
Dim lstrRemarks As String

 On Error GoTo Error_Handler

     strSQL = "Select rtrim(a.first_name) + ' ' + rtrim(a.last_name), b.provider_id  " _
        & "From ENCOUNTER b, Providers a " _
        & "Where b.patient_id = '" & piPatientID & "' And b.encounter_id = '" & piEncounterID & "'" _
        & "And a.provider_id = b.provider_id "
        
        
    booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
    
    
    If booRet Then
        If lngrecs > 0 Then
            piProviderID = varRS(1, 0)
            mstrProviderName = varRS(0, 0)
        Else
            MsgBox "Patient does not have a provider. Please go to Add/Manage Patients and assign a Provider.", vbCritical
            Exit Sub
        End If
    End If
    
    modSpecific.piProviderID = piProviderID
    Exit Sub
Error_Handler:
    ErrHandler _
     "frmMDIReport.Get_Provider_Id", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub



Public Function Get_Provider(vintProviderID As Integer) As String
Dim lstrSQL As String
Dim lvarRS
Dim llngRows As Long
Dim lbooRet As Boolean
Dim lvarMessage(2)
Dim lstrRemarks As String

On Error GoTo Error_Get_Provider

    lstrSQL = "Select rtrim(first_name) + ' ' + rtrim(last_name) +  ' ' + title " _
        & "From PROVIDERS " _
        & "Where provider_id = " & vintProviderID
    
    lvarMessage(0) = lstrSQL
    
    lbooRet = adoQuery(lstrSQL, psConnect(0), lvarRS, llngRows)
    
    If llngRows > 0 Then
        Get_Provider = lvarRS(0, 0)
    End If
    
    Exit Function
    
Error_Get_Provider:
    ErrHandler _
     "modSpecific > Get_Provider", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function


Private Sub cmdSelTo_Click()
Dim lstrSQL As String
Dim i As Integer
Dim lstrRemarks As String
Dim booRet As Boolean
Dim lngrecs As Long
Dim varRS

LogFile "Private Sub Get_Encounters-----------------------------------"

If piPatientID = 0 Then Exit Sub
lstEnctrTo.Clear
    lstrSQL = "Select encntr_start_date, encounter_id " _
        & "From ENCOUNTER " _
        & "Where patient_id = " & piPatientID & " Order By ENCNTR_START_DATE desc"
        
         booRet = SelectQuery(lstrSQL, psConnect(0), varRS, lngrecs)
         
        Load_ComboBox lstEnctrTo, lstrSQL, psConnect(0)
        
    fraEnctrTo.Visible = True
    fraEnctrTo.ZOrder 0
    lstPatient.Enabled = False
    lstEncounters.Enabled = False
    
    Exit Sub
End Sub

Private Sub cmdTClose_Click()
    fraEnctrTo.Visible = False
    lstPatient.Enabled = True
    lstEncounters.Enabled = True
End Sub


Private Sub cmdTEnctSel_Click()
    lstEnctrTo_Click
    fraEnctrTo.Visible = False
    lstPatient.Enabled = False
    lstEncounters.Enabled = False
End Sub


Private Sub Form_Activate()
    mbooPrint = False
    If mFirst Then
        Screen.MousePointer = vbHourglass
        Me.Caption = "PatientTrac Medical Records Report Application" & fstrTitle
        piPatientID = modSpecific.piPatientID
        piEncounterID = modSpecific.piEncounterID
        If piPatientID > 0 And piEncounterID > 0 Then
            Set_Patient
        End If
        If piPatientID = 0 Then
            MsgBox "Please select Patient and Encounter", vbCritical
            frmMDIReport.Caption = "PatientTrac Medical Records Report Application"
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            Me.Caption = clsPatient.Name & " " & clsPatient.IntId & " " & clsEncounter.StartDate & " " & fstrTitle
            fstrLastFilePath = pstrMainPath & "\Reports"
            lstReports.Path = pstrMainPath & "\Reports"
          '  Get_Referral
            Form_Resize
        End If
        mFirst = False
        Screen.MousePointer = vbDefault
    End If
    If piPatientID < 0 Then piPatientID = modSpecific.piPatientID
    If piEncounterID < 0 Then piEncounterID = modSpecific.piEncounterID
    Form_Resize
    Exit Sub
End Sub

Private Sub Form_Resize()
    If mFirst Then
        frmMDIReport.Width = grdMDIReport.Width + 300    ' 1.5
        Me.Height = grdMDIReport.Height + 8000
        frmMDIReport.Top = frmMDI.Top + 100
        frmMDIReport.Left = 0
    End If
End Sub

Private Sub Form_Load()
    piPatientID = modSpecific.piPatientID
    piEncounterID = modSpecific.piEncounterID
    grdMDIReport.EditModePermanent = True
    grdMDIReport.ProcessTab = True
    grdMDIReport.BackColorStyle = BackColorStyleUnderGrid
    cboForms.AddItem "Browser"
    cboForms.AddItem "Active Report"
    mFirst = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub



Private Sub Open_Script()
Dim intNum As Integer
Dim strLine As String
Dim i As Integer
Dim strTitle As String

    intNum = FreeFile
    If lstReports.FileName <> vbNullString Then
        fmstrfile = fstrLastFilePath & "\" & lstReports.FileName
    End If
    
    If fmstrfile = vbNullString Then
        fmstrfile = CommonDialogAction(CommonDialog, "Open", "Report Script", fstrLastFilePath, "*.txt")
    End If
    
    If lstReports.FileName <> vbNullString Then
        strTitle = Left(lstReports.FileName, Len(lstReports.FileName) - 9)
    Else
        If fmstrfile <> vbNullString Then
            strTitle = Right(fmstrfile, Len(fmstrfile) - 55)
            strTitle = Left(strTitle, Len(strTitle) - 9)
        End If
    End If
    frmMDIReport.fstrTitle = strTitle
    Me.Caption = clsPatient.Name & " " & clsPatient.IntId & " " & clsEncounter.StartDate & " " & fstrTitle
    
    If fmstrfile <> vbNullString Then
    Open fmstrfile For Input As intNum
    
        Do While Not EOF(intNum)
            Line Input #intNum, strLine
            
            If strLine <> vbNullString And Left(strLine, 1) = "<" Then
                ReDim Preserve fstrScriptLine(i)
                fstrScriptLine(i) = strLine
                i = i + 1
                Execute strLine
            End If
       
        Loop
    
    Close intNum

    Else
        MsgBox "No Report Selected!", vbInformation
        Exit Sub
    End If
    
    Remove_ResidualTags
Exit Sub
    Remove_ResidualTags
End Sub

Private Sub Execute(vstrLine As String)
Dim strItem() As String
Dim lngCol As Long
Dim lngRow As Long
Dim lngCol2 As Long
Dim lngRow2 As Long
Dim strDate As String
Dim strData As String
Dim intSide As Integer
Dim intStyle As Integer
Dim strEncodedSignature As String
Dim booInvalid As Boolean
Dim strRefData As String
Dim varImage
Dim strPath As String
Dim booData As Boolean
Dim varRS
Dim c%, r%, i%, k%, g%, n1%, n2%, n3%, n4%
Dim intSeq As Integer
Dim strObjType As String
Dim booBooleanType As Boolean
Dim strRowHeader() As String
Dim strColHeader() As String
Dim strColHeaderWidth() As String
Dim intLastCol As Integer
Dim strExcludedObject() As String
'Dim strExpectedObject() As String
Dim booSkip As Boolean
Dim booMissing As Boolean
Dim strtype As String
Dim strFontType() As String
Dim strComponentVal As String
Dim strObject As String
Dim booCase As Boolean
Dim strOCX As String
Dim strControl As String
Dim intPos1 As Integer
Dim intPos2 As Integer
Dim booNoteData As Boolean
Dim varNotes
Dim intSpan As Integer
Dim intTotSpan As Integer
Dim booDataPresent As Boolean

'Dynamic Row height..............
Dim varSpan(3)
Dim X%, Y%, z%
Dim intSpanWidth As Integer
Dim booShowData As Boolean
Dim intRowInserts As Integer
Dim booRowAdded As Boolean
Dim intPrevSeq As Integer
Dim intColWidth As Integer
Dim intInsertedRows As Integer
Dim varCell
Dim boogrdSSData As Boolean


    On Error GoTo Error_Execute
    grdMDIReport.ReDraw = False
    strItem() = Split(vstrLine, "~")

    
    booNoteData = False
    
    If modSpecific.piEncounterID = 0 Then modSpecific.piEncounterID = piEncounterID
    
    Select Case strItem(0)
    
        Case "<Alignment>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            lngCol2 = strItem(3)
            lngRow2 = strItem(4)
            intInsertedRows = Get_InsertRowCount(lngRow)
            Set_Block lngCol, lngRow + intInsertedRows, lngCol2, lngRow2 + intInsertedRows
            
            If strItem(5) = "HorizAlignLeft" Then
                grdMDIReport.TypeHAlign = TypeHAlignLeft
            ElseIf strItem(5) = "HorizAlignCenter" Then
                grdMDIReport.TypeHAlign = TypeHAlignCenter
            ElseIf strItem(5) = "HorizAlignRight" Then
                grdMDIReport.TypeHAlign = TypeHAlignRight
            ElseIf strItem(5) = "VertAlignTop" Then
                grdMDIReport.TypeVAlign = TypeVAlignTop
            ElseIf strItem(5) = "VertAlignCenter" Then
                grdMDIReport.TypeVAlign = TypeVAlignCenter
            ElseIf strItem(5) = "VertAlignBottom" Then
                grdMDIReport.TypeVAlign = TypeVAlignBottom
            End If
            
            grdMDIReport.BlockMode = False
            
        Case "<BackColor>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            lngCol2 = strItem(3)
            lngRow2 = strItem(4)
            intInsertedRows = Get_InsertRowCount(lngRow)
            Set_Block lngCol, lngRow + intInsertedRows, lngCol2, lngRow2 + intInsertedRows
            grdMDIReport.BackColor = strItem(5)
            grdMDIReport.BlockMode = False
            
        Case "<Border>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            lngCol2 = strItem(3)
            lngRow2 = strItem(4)
            
            Select Case strItem(5)  'Sides
                Case "None"
                    intSide = 0
                Case "Left"
                    intSide = 1
                Case "Right"
                    intSide = 2
                Case "Top"
                    intSide = 4
                Case "Bottom"
                    intSide = 8
                Case "All"
                    intSide = 16
            End Select
            
            Select Case strItem(7)  'Style
                Case "None"
                    intStyle = 0
                Case "Solid"
                    intStyle = 1
                Case "Dash"
                    intStyle = 13
            End Select
            
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetCellBorder lngCol, lngRow + intInsertedRows, lngCol2, lngRow2 + intInsertedRows, intSide, strItem(6), intStyle
        
        Case "<BorderInterior>"
            
        Case "<CellSpan>"
            lngCol = strItem(1) 'Start Col
            lngRow = strItem(2) 'Start Row
            lngCol2 = strItem(3) 'Quantity
            lngRow2 = strItem(4) 'Quantity
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.AddCellSpan lngCol, lngRow + intInsertedRows, lngCol2, lngRow2 '+ intInsertedRows
            
        Case "<ColWidth>"
            lngCol = strItem(1)
            grdMDIReport.ColWidth(lngCol) = strItem(2)
            
        Case "<ColHeaders>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            strColHeader() = Split(strItem(3), ",")
            'intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.Row = lngRow '+ intInsertedRows
            
            For i = LBound(strColHeader) To UBound(strColHeader)
                grdMDIReport.Col = lngCol + i
                grdMDIReport.SetText lngCol + i, lngRow, CVar(strColHeader(i))
                grdMDIReport.TypeHAlign = TypeHAlignCenter
                grdMDIReport.FontBold = True
            Next i
            
            ReDim Preserve fudtGridHeader(UBound(fudtGridHeader) + 1)
            fudtGridHeader(UBound(fudtGridHeader)).StartCol = lngCol
            fudtGridHeader(UBound(fudtGridHeader)).EndCol = lngCol + i
            fudtGridHeader(UBound(fudtGridHeader)).Row = lngRow
            
        Case "<ColHeaderWidths>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            strColHeaderWidth() = Split(strItem(3), ",")
            
            For i = LBound(strColHeaderWidth) To UBound(strColHeaderWidth)
                grdMDIReport.ColWidth(lngCol + i) = CInt(strColHeaderWidth(i))
            Next i
        
        Case "<Date>", "<DateTime>"
            lngCol = strItem(1) 'Start Col
            lngRow = strItem(2) 'Start Row
            strDate = Format(Now, strItem(3))
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(strDate)
            
        Case "<IncludedObjects>"
            lngCol = strItem(1) 'Start Col
            lngRow = strItem(2) 'Start Row
            
            If InStr(strItem(3), ",") > 0 Then
                fstrIncludedObject() = Split(strItem(3), ",")
            Else
                ReDim fstrIncludedObject(0)
                fstrIncludedObject(0) = strItem(3)
            End If
            
        Case "<EncounterStartDate>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.StartDate)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.StartDate = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterEndDate>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.EndDate)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.EndDate = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterVisitReason>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.VisitReason)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.VisitReason = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterFacility>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.Facility)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.Facility = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterProvider>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.Provider)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.Provider = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterServicePlace>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.ServicePlace)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.ServicePlace = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterServicePlaceCode>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.ServicePlaceCode)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.ServicePlaceCode = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterType>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.EncounterType)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.EncounterType = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
        
        Case "<ExcludedObjects>"
            lngCol = strItem(1) 'Start Col
            lngRow = strItem(2) 'Start Row
            
            If InStr(strItem(3), ",") > 0 Then
                fstrExcludedObject() = Split(strItem(3), ",")
            Else
                If strItem(3) = "" Then strItem(3) = "|"
                ReDim fstrExcludedObject(0)
                fstrExcludedObject(0) = strItem(3)
            End If
             
        Case "<FetchData>"  'Returns one value for a specific object on a control
            lngCol = strItem(1)
            lngRow = strItem(2)
'            If lngRow > 40 And lngRow < 44 Then
'                Debug.Print
'            End If
            intInsertedRows = Get_InsertRowCount(lngRow)
            If strItem(4) = "PATIENT_CLINICAL" Then
                '                             OCX       Control      objRef        Sequence
                strData = Get_ClinicalValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
            
                'dnn   6/7/2004.....if vbnullstring try PATIENT_NOTES table............
                If strData = "" Then
                    strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                    If strData <> "" Then booNoteData = True
                End If
                
                'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                If strData = "" Then
                    strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                End If
                
            ElseIf strItem(4) = "PATIENT_NOTES" Then
                strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                If strData <> "" Then booNoteData = True
            ElseIf strItem(4) = "PATIENT_PROFILER" Then
                strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
            ElseIf strItem(4) = "PATIENT_PROFILER_NOTES" Then
                strData = Get_ProfilerNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                If strData <> "" Then booNoteData = True
            End If
            
            strData = Trim(strData)
            grdMDIReport.Col = lngCol
            grdMDIReport.Row = lngRow + intInsertedRows
            'Reproductive Note exception
            If strItem(7) <> "cclNote" Then
            'dnn 6/6/2004  strip carriage returns..(Accrued data types)..........
                If strData <> "" And strItem(7) <> "cctSummary" And strItem(7) <> "cctAllergies" Then
                    If InStr(strData, vbNewLine) > 0 Then
                        strData = Replace(strData, vbNewLine, " ")
                            If InStr(strData, ";") > 0 Then
                                strData = Replace(strData, ";", ", ")
                            End If
                    End If
                End If
            End If
            grdMDIReport.GetText lngCol, lngRow + intInsertedRows, varCell
            
            If varCell = vbNullString Then
                Select Case strItem(3)
                    Case "StaticText"
                        grdMDIReport.CellType = CellTypeStaticText
                        grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(strData)
                        
                    Case "CheckBox"
                        grdMDIReport.CellType = CellTypeCheckBox
                        If strData = "True" Then strData = "1" Else strData = "0"
                        grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(strData)
                        
                        
                    Case "OptionButton"
                        grdMDIReport.CellType = CellTypeStaticText
                        grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(strData)
                    
                End Select
            End If
            
            grdMDIReport.GetText lngCol, lngRow + intInsertedRows, varCell
            
            If varCell = vbNullString Then
                If HideRow(lngRow + intInsertedRows) And strData = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            End If
            
        Case "<FetchBooleanData>" 'Outputs text based on a checkbox or option button True or False case.................
            lngCol = strItem(1)
            lngRow = strItem(2)
            If strItem(4) = "PATIENT_CLINICAL" Then
                '                             OCX       Control      objRef        Sequence
                strData = Get_ClinicalValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
            
                'dnn   6/7/2004.....if vbnullstring try PATIENT_NOTES table............
                If strData = "" Then
                    strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                    If strData <> "" Then booNoteData = True
                End If
                
                'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                If strData = "" Then
                    strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                End If
            
            ElseIf strItem(4) = "PATIENT_NOTES" Then
                strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                If strData <> "" Then booNoteData = True
            ElseIf strItem(4) = "PATIENT_PROFILER" Then
                strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
            ElseIf strItem(4) = "PATIENT_PROFILER_NOTES" Then
                strData = Get_ProfilerNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                If strData <> "" Then booNoteData = True
            End If
            
            grdMDIReport.Col = lngCol
            grdMDIReport.Row = lngRow
            grdMDIReport.CellType = CellTypeStaticText
            
            Select Case Trim(strData)

                Case "True", "1", "-1", "yes", "Yes", "YES", "y", "Y", "A"
                    booCase = True
                    
                Case "False", "", "no", "N", "NO", "I", "n", "0", "No"
                    booCase = False
                    
                Case Else
                    booCase = False
                
            End Select
            
            If booCase = CBool(strItem(9)) And strData = "True" Then
                grdMDIReport.SetText lngCol, lngRow, CVar(strItem(3))
            End If
            
            'dnn  6/7/2004...........6/24/2004
            If strData <> "False" And strData <> "" And strItem(9) = "True" Then
                grdMDIReport.SetText lngCol, lngRow, CVar(strItem(3))
            ElseIf strData <> "" Then
            
                'dnn 6/6/2004  strip carriage returns..(Accrued data types)..........
                If strData <> "" Then
                    If InStr(strData, vbNewLine) > 0 Then
                        'get data between carriage returns..........
                        intPos1 = InStr(strData, vbNewLine) + 2
                        intPos2 = InStr(intPos1, strData, vbNewLine)
                        strData = Trim(Mid(strData, intPos1 + 1, intPos2 - intPos1 - 1))
                        
                        If InStr(strData, ";") > 0 Then
                            strData = Replace(strData, ";", ", ")
                        End If
                    End If
                End If
                
                If strData <> "False" And strData <> "True" Then
                    grdMDIReport.SetText lngCol, lngRow, CVar(strData)
                    
                End If
            End If
            
            
            '6/24/2004................
            If strData = "" Or strData = "False" And strItem(9) = "False" Then
                grdMDIReport.SetText lngCol, lngRow, ""
                
            End If
            
            'If HideRow(0) > 0 Then strData
                'If HideRow(lngRow) And strItem(3) = "" Then grdMDIReport.RowHeight(lngRow) = 0
                If HideRow(lngRow) And Not booCase Then
                    grdMDIReport.ROWHEIGHT(lngRow) = 0
                Else
                    grdMDIReport.ROWHEIGHT(lngRow) = 9
                    
                End If
            'End If
            
            
        Case "<FetchBooleanComp>" 'Outputs component value based on associated checkbox or option button True or False case.................
            lngCol = strItem(1)
            lngRow = strItem(2)
            
           ' If lngRow = 24 Then
                'Stop
            'End If
            'If booCase = False Then
            If strItem(4) = "PATIENT_CLINICAL" Then
                '                             OCX       Control      objRef        Sequence
                strData = Get_ClinicalValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
            
                'dnn   6/7/2004.....if vbnullstring try PATIENT_NOTES table............
                If strData = "" Then
                    strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                    If strData <> "" Then booNoteData = True
                End If
                
                'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                If strData = "" Then
                    strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                End If
                
            ElseIf strItem(4) = "PATIENT_NOTES" Then
                strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                If strData <> "" Then booNoteData = True
            ElseIf strItem(4) = "PATIENT_PROFILER" Then
                strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
            ElseIf strItem(4) = "PATIENT_PROFILER_NOTES" Then
                strData = Get_ProfilerNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                If strData <> "" Then booNoteData = True
            End If
            
            grdMDIReport.Col = lngCol
            grdMDIReport.Row = lngRow
            grdMDIReport.CellType = CellTypeStaticText
            
            strObject = Mid(strItem(3), InStr(strItem(3), ":") + 1)
           ' End If
            Select Case Trim(strData)

                Case "True", "1", "-1", "yes", "Yes", "YES", "y", "Y", "A"
                    booCase = True
                    
                Case "False", "", "no", "N", "NO", "I", "n", "0", "No"
                    booCase = False
                
                Case Else
                    booCase = False
                    
            End Select
            
            If booCase = CBool(strItem(9)) Then
                strObjType = strItem(10)
                strOCX = Left(strItem(3), InStr(strItem(3), ".") - 1)
                strControl = Mid(strItem(3), InStr(strItem(3), ".") + 1)
                strControl = Left(strControl, InStr(strControl, ":") - 1)
                
                If strObjType = "ccList" Or strObjType = "ccListGrid" Or strObjType = "ccMultiCheck" Or strObjType = "ListBox" Then
                    fbooSkip = False
                Else
                    If strItem(4) = "PATIENT_CLINICAL" Then
                        '                                     OCX       Control      objRef        Sequence
                        'strComponentVal = Get_ClinicalValue(strItem(5), strItem(6), strObject, CInt(strItem(8)), piPatientID, piEncounterID)
                        strComponentVal = Get_ClinicalValue(strOCX, strControl, strObject, CInt(strItem(8)), piPatientID, piEncounterID)
                        
                        'if vbnullstring try PATIENT_NOTES table............
                        If strComponentVal = "" Then
                            'dnn 6/7/2004............
                            'strComponentVal = Get_ClinicalNote(strItem(5), strItem(6), strObject, CInt(strItem(8)), piPatientID, piEncounterID)
                            strComponentVal = Get_ClinicalNote(strOCX, strControl, strObject, CInt(strItem(8)), piPatientID, piEncounterID)
                            If strComponentVal <> "" Then booNoteData = True
                        End If
                        
                        'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                        If strComponentVal = "" Then
                            strComponentVal = Get_ProfilerValue(strOCX, strControl, strObject, CInt(strItem(8)), piPatientID)
                        End If
                    
                    ElseIf strItem(4) = "PATIENT_NOTES" Then
                        strComponentVal = Get_ClinicalNote(strItem(5), strItem(6), strObject, CInt(strItem(8)), piPatientID, piEncounterID)
                        If strData <> "" Then booNoteData = True
                    ElseIf strItem(4) = "PATIENT_PROFILER" Then
                        strComponentVal = Get_ProfilerValue(strItem(5), strItem(6), strObject, CInt(strItem(8)), piPatientID)
                    ElseIf strItem(4) = "PATIENT_PROFILER_NOTES" Then
                        strComponentVal = Get_ProfilerNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                        If strData <> "" Then booNoteData = True
                    End If
                
                End If
                If strComponentVal <> "" Then
                    grdMDIReport.SetText lngCol, lngRow, CVar(strComponentVal)
                End If
            Else
                'grdMDIReport.SetText lngCol, lngRow, ""
            End If
            
            ReDim Preserve fudtAudit(UBound(fudtAudit) + 1)
            fudtAudit(UBound(fudtAudit)).OCX = Left(strItem(3), InStr(strItem(3), ".") - 1)
            fudtAudit(UBound(fudtAudit)).Control = Mid(strItem(3), InStr(strItem(3), ".") + 1, (InStr(strItem(3), ":") - InStr(strItem(3), ".") - 1))
            fudtAudit(UBound(fudtAudit)).Object = Mid(strItem(3), InStr(strItem(3), ":") + 1)
            fudtAudit(UBound(fudtAudit)).DataPresent = booCase
            fudtAudit(UBound(fudtAudit)).ShowWhen = strItem(9)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow) And strComponentVal = "" Then grdMDIReport.ROWHEIGHT(lngRow) = 0
            'End If
        
            
                    
        Case "<FetchAnnotation>"

            ccaImage.Image = strItem(7)
            'Get_Image ccaImage, strItem(3), strItem(4), strItem(5), CInt(strItem(6)), piPatientId, piEncounterID
            
            'strPath = Get_FilePath(strItem(7))
            strPath = App.Path & "\Images"
            If Not FolderExists(strPath) Then
                MkDir strPath
            End If
            
            ccaImage.Object.SaveFile App.Path & "\Images\image.bmp"
            
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            
            grdMDIReport.Col = lngCol
            grdMDIReport.Row = lngRow + intInsertedRows
            grdMDIReport.CellType = CellTypePicture
            grdMDIReport.TypePictCenter = True
                
            If FileExists(App.Path & "\Images\image.bmp") Then
                grdMDIReport.TypePictPicture = LoadPicture(App.Path & "\Images\image.bmp")
            End If
            
         Case "<FetchLabData>"
        
        
        Case "<FetchList>"
            Select Case strItem(3) 'Table
                Case "PATIENT_CLINICAL"
                    '                            OCX        Control     objRef
                    booData = Get_ClinicalList(strItem(4), strItem(5), strItem(6), varRS, piPatientID, piEncounterID)
                
                    'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                    If Not booData Then
                        '                                   OCX        Control     objRef
                        booData = Get_ProfilerClinicalList(strItem(4), strItem(5), strItem(6), varRS, piPatientID)
                    End If
                
            End Select
            
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            
            If booData Then
                i = 0
                
                'If frmFarpointReports.chkConcatList.Value = 1 Then
                    'For r = LBound(varRS, 2) To UBound(varRS, 2)
                        'If InStr(varRS(0, r), vbNewLine) = 0 Then
                            'strData = strData & varRS(0, r) & ", "
                        'End If
                   'Next r
                    
                    For r = LBound(varRS, 2) To UBound(varRS, 2)
                        If InStr(varRS(0, r), vbNewLine) = 0 Then
                            strData = strData & varRS(0, r) & ", "
                        Else
                            strData = varRS(0, r)
                        End If
                    Next r
                    
                    
                    If Right(strData, 2) = ", " Then
                        strData = Left(strData, Len(strData) - 2)
                    End If
                    
                    strData = Replace(strData, vbNewLine, " ,")
                    
                    grdMDIReport.SetText lngCol, lngRow + intInsertedRows, strData
                    'grdMDIReport.AddCellSpan lngCol, lngRow, strItem(7), 1
                            

            Else
                'grdMDIReport.SetText lngCol, lngRow, ""
            End If
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow) And strData = "" Then grdMDIReport.ROWHEIGHT(lngRow) = 0
            'End If
            
            
        Case "<FetchRecordSet>"
            booShowData = True
            'Check Audit trail....................
            For X = LBound(fudtAudit) To UBound(fudtAudit)
            
                If fudtAudit(X).OCX = strItem(4) _
                    And fudtAudit(X).Control = strItem(5) _
                    And fudtAudit(X).Object = strItem(6) Then
                    'Debug.Print
                 
                    If fudtAudit(X).ShowWhen Then
                       'Show grid when there is data
                        If fudtAudit(X).DataPresent Then
                            booShowData = True
                        Else
                            booShowData = False
                        End If
                    Else
                        'Show grid when there is no data
                        If fudtAudit(X).DataPresent Then
                            booShowData = False
                        Else
                            booShowData = True
                        End If
                    End If
                    Exit For
                End If
            Next X
            
            If Not booShowData Then
                'Destroy grid header..................
                lngCol = strItem(1)
                lngRow = strItem(2)
                'grdMDIReport.RowHeight(lngRow - 1) = 0
                For X = LBound(fudtGridHeader) To UBound(fudtGridHeader)
                    If fudtGridHeader(X).StartCol = lngCol And fudtGridHeader(X).Row = lngRow - 1 Then
                        For Y = fudtGridHeader(X).StartCol To fudtGridHeader(X).EndCol
                            grdMDIReport.SetText Y, lngRow - 1 + fintRowInserts, ""
                        Next Y
                    End If
                Next X
                
            End If
            
            If Not fbooSkip And booShowData Then
                Select Case strItem(3) 'Table
                    Case "PATIENT_CLINICAL"
                    
                        If strItem(6) = "grdSS" Then
                            If strItem(5) = "Medication" Then
                                booData = Get_MedicationData(varRS, piPatientID, piEncounterID, piProviderID)
                            ElseIf strItem(5) = "Lab_Results" Then
                                booData = Get_Lab_ResultsData(piPatientID, piEncounterID, piProviderID)
                                'Set varRS, pudtLabs returns multiple record sets ..............
                                lngCol = strItem(1): lngRow = strItem(2)
                                For n1 = LBound(pudtLabs, 1) To UBound(pudtLabs, 1) 'Encounters
                                    For n2 = LBound(pudtLabs, 2) To UBound(pudtLabs, 2) 'Orders
                                        For n4 = LBound(pudtLabs, 4) To UBound(pudtLabs, 4) 'Lab Rows
                                            ReDim varRS(UBound(pudtLabs, 3), UBound(pudtLabs, 4))
                                            For n3 = LBound(pudtLabs, 3) To UBound(pudtLabs, 3) 'Lab Cols
                                                varCell = CVar(pudtLabs(n1, n2, n3, n4).Value)
                                                varRS(n3, n4) = varCell
                                            Next n3
                                        Next n4
                                        GoSub ProcessRS
                                    Next n2
                                Next n1
                                i = i - 1
                            End If
                        Else
                            '                            OCX        Control
                            booData = Get_ClinicalData(strItem(4), strItem(5), varRS, piPatientID, piEncounterID)
                        End If
                        
                        'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                        If Not booData Then
                            '                                    OCX        Control
                            booData = Get_ProfilerClinicalData(strItem(4), strItem(5), varRS, piPatientID)
                        End If
                        
                End Select
                
                lngCol = strItem(1)
                lngRow = strItem(2)
                
                booRowAdded = False
                If booData Then
                    If strItem(5) <> "Lab_Results" Then
                    intLastCol = strItem(8)
                    'grdMDIReport.InsertRows lngRow + 1, 1
                    i = 0
                    
                    'If fintRowInserts > 0 Then
                    '    fintRowInserts = fintRowInserts - 1
                    'Else
                    '    intRowInserts = fintRowInserts
                    'End If
                    intRowInserts = fintRowInserts
                    boogrdSSData = False
                    
                    'If strItem(5) <> "Lab_Results" Then
                    For r = LBound(varRS, 2) To UBound(varRS, 2)
                    
                        intSpan = 0
                        booSkip = False
                        For k = LBound(fstrExcludedObject) To UBound(fstrExcludedObject)
                            If InStr(varRS(0, r), fstrExcludedObject(k)) > 0 Then
                                booSkip = True
                                Exit For
                            End If
                        Next k
                                                
                        'Avoid accrued values..................
                        If InStr(varRS(2, r), vbNewLine) = 0 And Not booSkip Then
                            
                            If strItem(6) = "grdSS" And booData Then
                                If strItem(5) = "Medication" Then
                                    LoadRow_Medication grdMDIReport, varRS, r, lngCol, lngRow + fintRowInserts, fstrIncludedObject(), fstrExcludedObject()
                                'ElseIf strItem(5) = "Lab_Results" Then
                                '    LoadRow_LabResults grdMDIReport, varRS, r, lngCol, lngRow + fintRowInserts, fstrIncludedObject(), fstrExcludedObject()
                                End If
                                    i = i + 1
                                    
                                    If i > 1 Then
                                        grdMDIReport.InsertRows lngRow + i + fintRowInserts, 1
                                    End If
                                    
                                    fintRowInserts = fintRowInserts + 1
                            
                                    If i = UBound(varRS, 2) + 1 Then
                                        i = i - 1
                                    End If
                                
                                
                            Else
                                intSeq = varRS(1, r)
                                strObjType = Left(varRS(0, r), 3)
                            
                                grdMDIReport.Col = IdentifyColumn(lngCol, varRS(0, r)) 'lngCol + c
                            End If
                            
                            'Any cell spans applied to this column?........
                            '                                             ColAnchor,  RowAnchor,  NumCols,    NumRows
                             grdMDIReport.GetCellSpan grdMDIReport.Col, lngRow, varSpan(0), varSpan(1), varSpan(2), varSpan(3)
                            
                            If varSpan(2) > 0 Then 'There is a cell span
                                'If grdMDIReport.Col > grdMDIReport.Col * varSpan(2) And grdMDIReport.Col <= varSpan(0) + varSpan(2) - 1 Then
                                    grdMDIReport.Col = grdMDIReport.Col * varSpan(2) - lngCol ' varSpan(0) + varSpan(2)
                                'End If
                                
                                'Does data exists in this cell?
                                grdMDIReport.GetText grdMDIReport.Col, lngRow, varCell
                                'If varCell <> vbnullstring Then
                                    'grdMDIReport.Col = varSpan(0) + varSpan(2)
                                'End If
                                
                                intSpan = varSpan(2) - 1
                                intTotSpan = intTotSpan + intSpan
                                
                            End If
                            'grdMDIReport.Col = lngCol + c
                            
                            If intPrevSeq <> intSeq And r > 0 And intSeq > 1 Then
                
                                i = i + 1
                                If fintRowInserts > 1 Then
                                    grdMDIReport.InsertRows lngRow + i + fintRowInserts, 1
                                Else
                                    grdMDIReport.InsertRows lngRow + i, 1
                                End If
                                'intTotSpan = 0
                                'ReDim Preserve fudtRowInsert(fintRowInserts)
                                fintRowInserts = fintRowInserts + 1
                                'fudtRowInsert(fintRowInserts).InsertedRows = fintRowInserts
                                'fudtRowInsert(fintRowInserts).StartRowlngRow
                                
                                'grdMDIReport.AddCellSpan grdMDIReport.Col, lngRow + i, varSpan(2), 1
                                'grdMDIReport.InsertRows lngRow + i, 1
                                'intRowInserts = intRowInserts + 1
                                booRowAdded = True
                            End If
                            
                            If booRowAdded Then
                                grdMDIReport.AddCellSpan grdMDIReport.Col, lngRow + i, varSpan(2), 1
                            End If
                            grdMDIReport.Row = lngRow + i + intRowInserts
                            
                            If grdMDIReport.Col > 0 Then
                                Select Case strObjType
                                    Case "chk", "ccc", "cco", "opt"
                                        grdMDIReport.CellType = CellTypeCheckBox
                                        booBooleanType = True
                                        
                                    Case Else
                                        booBooleanType = False
                                End Select
                            
                                If strItem(6) <> "grdSS" Then
                                    grdMDIReport.TypeHAlign = TypeHAlignCenter
                                    grdMDIReport.SetText grdMDIReport.Col, lngRow + i + intRowInserts, varRS(2, r)
                                
                                
                                    Select Case strItem(7) 'Interior Borders
                                    
                                        Case "Vert"
                                            grdMDIReport.SetCellBorder grdMDIReport.Col, lngRow + i + intRowInserts, grdMDIReport.Col + intSpan, lngRow + i, 1, 0, 1  'Left
                                            grdMDIReport.SetCellBorder grdMDIReport.Col, lngRow + i + intRowInserts, grdMDIReport.Col + intSpan, lngRow + i, 2, 0, 1  'Right
                                        Case "Horiz"
                                            grdMDIReport.SetCellBorder grdMDIReport.Col, lngRow + i + intRowInserts, grdMDIReport.Col + intSpan, lngRow + i + intRowInserts, 4, 0, 1 'Top
                                            grdMDIReport.SetCellBorder grdMDIReport.Col, lngRow + i + intRowInserts, grdMDIReport.Col + intSpan, lngRow + i + intRowInserts, 8, 0, 1 'Bottom
                                        Case "Both"
                                            grdMDIReport.SetCellBorder grdMDIReport.Col, lngRow + i + intRowInserts, grdMDIReport.Col + intSpan, lngRow + i, 16, 0, 1
                                    End Select
                                End If
                            
                             End If
                            
                            'If HideRow(0) > 0 Then
                                If HideRow(lngRow + fintRowInserts) And varRS(2, r) = "" Then grdMDIReport.ROWHEIGHT(lngRow + fintRowInserts) = 0
                            'End If
    
                            c = c + 1
                            
                            If r > 0 And r < UBound(varRS, 2) Then
                                If intSeq <> intPrevSeq Then 'varRS(1, r + 1)
                                    c = 0
                                    
                                    'grdMDIReport.InsertRows lngRow + i, 1
                                    'fintRowInserts = fintRowInserts + i
                                    'i = i + 1
                                End If
                            End If
    
                        End If
                       
                       intPrevSeq = intSeq
                    Next r
                    End If
                    'wayne 12/12
                    'Outline .......................
                   ' grdMDIReport.SetCellBorder lngCol, lngRow + intRowInserts, lngCol + intLastCol - 1 + intTotSpan / (i + 1), lngRow + i + intRowInserts, 16, 0, 1
                    grdMDIReport.SetCellBorder lngCol, lngRow + intRowInserts, lngCol + intLastCol - 1 + intTotSpan, lngRow + i + intRowInserts, 16, 0, 1
                    'Outline Header..................................
                   ' grdMDIReport.SetCellBorder lngCol, lngRow - 1 + intRowInserts, lngCol + intLastCol - 1 + intTotSpan / (i + 1), lngRow - 1 + intRowInserts, 16, 0, 1
                   ' grdMDIReport.SetCellBorder lngCol, lngRow - 1 + intRowInserts, lngCol + intLastCol - 1 + intTotSpan / (i + 1), lngRow - 1 + intRowInserts, 2, 0, 1
                    grdMDIReport.SetCellBorder lngCol, lngRow - 1 + intRowInserts, lngCol + intLastCol - 1 + intTotSpan, lngRow - 1 + intRowInserts, 16, 0, 1
                    grdMDIReport.SetCellBorder lngCol, lngRow - 1 + intRowInserts, lngCol + intLastCol - 1 + intTotSpan, lngRow - 1 + intRowInserts, 2, 0, 1
                    
                    
                    ReDim Preserve fudtRowInsert(UBound(fudtRowInsert) + 1)
                    fudtRowInsert(UBound(fudtRowInsert)).InsertedRows = fintRowInserts
                    fudtRowInsert(UBound(fudtRowInsert)).StartRow = lngRow
                    'SuppressUnusedCols grdMDIReport, lngCol, lngRow, lngCol + CLng(intLastCol) - 1, lngRow + CLng(i)
                    
                End If
                fbooSkip = False
            End If
            
        Case "<Font>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            'lngCol2 = strItem(3)
            'lngRow2 = strItem(4)
            'Set_Block lngCol, lngRow, lngCol2, lngRow2
            '    grdMDIReport.FontName = strItem(5)
            '    grdMDIReport.FontSize = CInt(strItem(6))
            '    grdMDIReport.ForeColor = CLng(strItem(7))
                
            'Set_Block lngCol, lngRow + intInsertedRows, lngCol, lngRow + intInsertedRows
             Set_Block lngCol, lngRow, lngCol, lngRow + intInsertedRows
                grdMDIReport.FontName = strItem(3)
                grdMDIReport.FontSize = CInt(strItem(4))
                grdMDIReport.ForeColor = CLng(strItem(5))
                
                'Separate Type into distinct words.................
'                If strItem(8) <> "" Then
'                    strType = Separate_Words(strItem(8))
                If strItem(6) <> "" Then
                    strtype = Separate_Words(strItem(6))
                    
                    If InStr(strtype, " ") > 0 Then
                        strFontType() = Split(strtype, " ")
                        For r = LBound(strFontType) To UBound(strFontType)
                            Select Case strFontType(r)
                                Case "Bold"
                                    grdMDIReport.FontBold = True
                                Case "Italic"
                                    grdMDIReport.FontItalic = True
                                Case "Underline"
                                    grdMDIReport.FontUnderline = True
                            End Select
                        Next r
                    Else
                        Select Case strtype
                            Case "Bold"
                                grdMDIReport.FontBold = True
                            Case "Italic"
                                grdMDIReport.FontItalic = True
                            Case "Underline"
                                grdMDIReport.FontUnderline = True
                        End Select
                    End If
                End If
            
            grdMDIReport.BlockMode = False
            
            
'        Case "<ForeColor>"
'            lngCol2 = strItem(3)
'            lngRow2 = strItem(4)
'            Set_Block lngCol, lngRow, lngCol2, lngRow2
'            grdMDIReport.ForeColor = strItem(5)
'            grdMDIReport.BlockMode = False
'
'        Case "<FontName>"
'            lngCol2 = strItem(3)
'            lngRow2 = strItem(4)
'            Set_Block lngCol, lngRow, lngCol2, lngRow2
'            grdMDIReport.FontName = strItem(5)
'            grdMDIReport.BlockMode = False
'
'        Case "<FontSize>"
'            lngCol2 = strItem(3)
'            lngRow2 = strItem(4)
'            Set_Block lngCol, lngRow, lngCol2, lngRow2
'            grdMDIReport.FontSize = strItem(5)
'            grdMDIReport.BlockMode = False
'
'        Case "<FontType>"
'            lngCol2 = strItem(3)
'            lngRow2 = strItem(4)
'            Set_Block lngCol, lngRow, lngCol2, lngRow2
'
'            Select Case strItem(5)
'                Case "Bold"
'                    grdMDIReport.FontBold = True
'                Case "Italic"
'                    grdMDIReport.FontItalic = True
'                Case "Underline"
'                    grdMDIReport.FontUnderline = True
'                Case "Strike"
'                    grdMDIReport.FontStrikethru = True
'            End Select
'
'            grdMDIReport.BlockMode = False
            
        Case "<GridDim>"

           ' centerForm Me
            grdMDIReport.Width = strItem(1)
            grdMDIReport.Height = strItem(2) + 3000
            Me.Width = strItem(1) + 400
            Me.Height = strItem(2) + 3000
            
        Case "<GridLines>"
            grdMDIReport.GridShowVert = strItem(1)
            grdMDIReport.GridShowHoriz = strItem(2)
              
        Case "<Headers>"
            grdMDIReport.ColHeadersShow = strItem(1)
            grdMDIReport.RowHeadersShow = strItem(2)
            
        Case "<Image>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            grdMDIReport.Col = lngCol
            grdMDIReport.Row = lngRow
            grdMDIReport.CellType = CellTypePicture
            grdMDIReport.TypePictCenter = True
            
            strPath = Get_FilePath(strItem(3))
            
            If Not FolderExists(strPath) Then
                MkDir strPath
            End If
            
            If FileExists(strItem(3)) Then
                grdMDIReport.TypePictPicture = LoadPicture(strItem(3))
            Else
                MsgBox "Image " & strItem(3) & " not found.", vbInformation
            End If
            
        Case "<Max>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            'intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.MaxCols = lngCol
            grdMDIReport.MaxRows = lngRow
            
            
        Case "<MaxChars>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.Col = lngCol
            grdMDIReport.Row = lngRow + intInsertedRows
            grdMDIReport.TypeMaxEditLen = strItem(3)
            
            
        Case "<MergeCols>" 'spans cells in that column having the same content
            lngCol = strItem(1)
            lngCol2 = strItem(2)
            For c = lngCol To lngCol2
                grdMDIReport.Col = c
                grdMDIReport.ColMerge = MergeAlways
            Next c
            
        Case "<MergeRows>" 'spans cells in that row having the same content
            lngRow = strItem(1)
            lngRow2 = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            
            For r = lngRow + intInsertedRows To lngRow2 + intInsertedRows
                grdMDIReport.Row = r
                grdMDIReport.RowMerge = MergeAlways
            Next r
        
        Case "<PatientName>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.Name)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.Name = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
        
        Case "<PatientDOB>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.DOB)
            
         'If HideRow(0) > 0 Then
            If HideRow(lngRow + intInsertedRows) And clsPatient.DOB = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
         'End If
        
        
        Case "<PatientSex>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.Gender)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.Gender = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientAddress>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.Address)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.Address = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientCityStateZip>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.CityStateZip)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.CityStateZip = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientPhone>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.Phone)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.Phone = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientEmail>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.Email)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.Email = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientSS>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.SS)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.SS = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientFCC>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.FCC)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.FCC = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientRace>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.Race)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.Race = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientAge>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.Age)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.Age = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientIntID>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(piPatientID)
        
        Case "<PatientExtID>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(clsPatient.ExtID)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And clsPatient.ExtID = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<ReferralPracticeName>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.Practice)
            If HideRow(lngRow + intInsertedRows) And udtReferral.Practice = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0

        Case "<ReferralPhysician>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.Physician)
            If HideRow(lngRow + intInsertedRows) And udtReferral.Physician = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
        
        Case "<ReferralAddress>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.Address)
            If HideRow(lngRow + intInsertedRows) And udtReferral.Address = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
           
        Case "<ReferralCityStateZipCode>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.CityStateZip)
            If HideRow(lngRow + intInsertedRows) And udtReferral.CityStateZip = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralPhone>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.Phone)
            If HideRow(lngRow + intInsertedRows) And udtReferral.Phone = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralCellPhone>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.CellPhone)
            If HideRow(lngRow + intInsertedRows) And udtReferral.CellPhone = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralFAX>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.FAX)
            If HideRow(lngRow + intInsertedRows) And udtReferral.FAX = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralEmail>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.Email)
            If HideRow(lngRow + intInsertedRows) And udtReferral.Email = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralUPIN>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.UPIN)
            If HideRow(lngRow + intInsertedRows) And udtReferral.UPIN = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralLicense>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.License)
            If HideRow(lngRow + intInsertedRows) And udtReferral.License = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralSpecialty>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(udtReferral.Specialty)
            If HideRow(lngRow + intInsertedRows) And udtReferral.Specialty = "" Then grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = 0

        Case "<RemoveRow>"
            If strItem(1) = "NoData" Then
                If UBound(strItem) > 1 Then
                    ReDim flngRowHidden(UBound(strItem) - 2) ' the trailing numbers in the <RemoveRow>~NoData tag
                    For i = LBound(flngRowHidden) To UBound(flngRowHidden)
                        flngRowHidden(i) = strItem(i + 2)
                    Next i
                End If
            End If
            
        Case "<RowHeaders>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            strRowHeader() = Split(strItem(3), ",")
            
            For i = LBound(strRowHeader) To UBound(strRowHeader)
                grdMDIReport.SetText lngCol + i, lngRow + intInsertedRows, CVar(strRowHeader(i))
                grdMDIReport.Col = lngCol + i
                grdMDIReport.Row = lngRow + intInsertedRows
                grdMDIReport.TypeHAlign = TypeHAlignCenter
                grdMDIReport.FontBold = True
                grdMDIReport.SetCellBorder lngCol + i, lngRow + intInsertedRows, lngCol + i, lngRow + intInsertedRows, 16, 0, 1
            Next i
        
        Case "<RowHeadersColWidth>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            strRowHeader() = Split(strItem(3), ",")
            
            For i = LBound(strRowHeader) To UBound(strRowHeader)
                grdMDIReport.ColWidth(lngCol + i) = CInt(strRowHeader(i))
            Next i
        
        Case "<RowHeight>"
            lngRow = strItem(1)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdMDIReport.ROWHEIGHT(lngRow + intInsertedRows) = strItem(2)
            
        Case "<ScrollBars>"
                If strItem(1) = "False" And strItem(2) = "False" Then
                    grdMDIReport.ScrollBars = ScrollBarsNone
                ElseIf strItem(1) = "True" And strItem(2) = "False" Then
                    grdMDIReport.ScrollBars = ScrollBarsVertical
                ElseIf strItem(1) = "False" And strItem(2) = "True" Then
                    grdMDIReport.ScrollBars = ScrollBarsHorizontal
                ElseIf strItem(1) = "True" And strItem(2) = "True" Then
                    grdMDIReport.ScrollBars = ScrollBarsBoth
                End If
        
        Case "<Signature>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            
            '                                    OCX       Control        Ref             Seq
            strEncodedSignature = Get_Signature(strItem(3), strItem(4), strItem(5), CInt(strItem(6)), piPatientID, piEncounterID, strRefData)
            If strEncodedSignature <> vbNullString Then
                'booInvalid = ccaSignature.DecodeSignatureValues(strEncodedSignature, "CurrentEstraceESTRACE2mgBUCCAL0TID04/28/200300621")
                booInvalid = ccaSignature.DecodeSignatureValues(strEncodedSignature, strRefData)
                picSignature.Picture = ccaSignature.Picture
                SavePicture picSignature.Image, App.Path & "\Images\sig.bmp"
                grdMDIReport.Col = lngCol
                grdMDIReport.Row = lngRow + intInsertedRows
                grdMDIReport.CellType = CellTypePicture
                grdMDIReport.TypePictCenter = True
                grdMDIReport.TypePictPicture = LoadPicture(App.Path & "\Images\sig.bmp")
            End If
            
        
        Case "<Text>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            'grdMDIReport.TopRow = lngRow - 8
            grdMDIReport.SetText lngCol, lngRow + intInsertedRows, CVar(strItem(3))
            
            
        Case "<WordWrap>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            lngCol2 = strItem(3)
            lngRow2 = strItem(4)
            intInsertedRows = Get_InsertRowCount(lngRow)
            Set_Block lngCol, lngRow + intRowInserts, lngCol2, lngRow2 + intRowInserts
            'grdMDIReport.Col = lngCol
            'grdMDIReport.Row = lngRow
            
            'If strItem(3) = "True" Then
                grdMDIReport.TypeEditMultiLine = True
                grdMDIReport.TypeTextWordWrap = True
            'Else
            '    grdMDIReport.TypeEditMultiLine = False
            '    grdMDIReport.TypeTextWordWrap = False
            'End If
            grdMDIReport.BlockMode = False
            
    End Select
    
    If booNoteData Then
        intSpanWidth = 0
        'Check for cell span...............................
       '                                      ColAnchor,   RowAnchor,  NumCols,    NumRows
        grdMDIReport.GetCellSpan lngCol, lngRow, varSpan(0), varSpan(1), varSpan(2), varSpan(3)
        
        For z = 0 To varSpan(2) - 1
            intSpanWidth = intSpanWidth + grdMDIReport.ColWidth(lngCol + z)
        Next z
        
        If intSpanWidth = 0 Then intSpanWidth = grdMDIReport.ColWidth(lngCol)
        
        'Set cell to Wordwrap................
        Set_Block lngCol, lngRow, lngCol, lngRow
        grdMDIReport.TypeEditMultiLine = True
        grdMDIReport.TypeTextWordWrap = True
        grdMDIReport.BlockMode = False
        
        'Expand row size.................
        lblHidden.Width = intSpanWidth * 121
        grdMDIReport.GetText lngCol, lngRow, varNotes
        lblHidden.Caption = CStr(varNotes)

        If varNotes <> vbNullString Then
        Dim cellheight As Long
        grdMDIReport.BlockMode = True
        grdMDIReport.Col = lngCol
        grdMDIReport.Row = lngRow
        ' Define cell type as edit
        grdMDIReport.CellType = CellTypeStaticText
        grdMDIReport.TypeTextWordWrap = True
        'grdMDIReport.TypeEditMultiLine = True
    ' Provide text for the cell
        grdMDIReport.Text = varNotes
    ' Return the value of the selected cell
        cellheight = grdMDIReport.MaxTextCellHeight
        If grdMDIReport.ROWHEIGHT(lngRow) < cellheight Then
            grdMDIReport.ROWHEIGHT(lngRow) = cellheight
        End If
        grdMDIReport.BlockMode = False

        End If
        'grdMDIReport.TopRow = lngRow - 8
        
        
    End If
    
'    For g = 1 To grdMDIReport.MaxCols
'        intColWidth = grdMDIReport.MaxTextColWidth(g)
'        grdMDIReport.ColWidth(g) = intColWidth
'    Next g
            
    grdMDIReport.SetActiveCell 0, 0
    
    Exit Sub
    
ProcessRS: 'Lab Result record sets......
    'booRowAdded = False
    intLastCol = strItem(8)
    'i = 0
    'intRowInserts = fintRowInserts
    'boogrdSSData = False
    
    For r = LBound(varRS, 2) To UBound(varRS, 2)
        intSpan = 0
        booSkip = False
        LoadRow_LabResults grdMDIReport, varRS, r, lngCol, lngRow + fintRowInserts, fstrIncludedObject(), fstrExcludedObject()
        i = i + 1
        
        If i > 1 Then
            grdMDIReport.InsertRows lngRow + i + fintRowInserts, 1
        End If
        
        fintRowInserts = fintRowInserts + 1

        If i = UBound(varRS, 2) + 1 Then
            'i = i - 1
        End If
        
         'Any cell spans applied to this column?........
         '                                             ColAnchor,  RowAnchor,  NumCols,    NumRows
        grdMDIReport.GetCellSpan grdMDIReport.Col, lngRow, varSpan(0), varSpan(1), varSpan(2), varSpan(3)
         
        If varSpan(2) > 0 Then 'There is a cell span
             
            grdMDIReport.Col = grdMDIReport.Col * varSpan(2) - lngCol
    
            'Does data exists in this cell?
            grdMDIReport.GetText grdMDIReport.Col, lngRow, varCell
    
            intSpan = varSpan(2) - 1
            intTotSpan = intTotSpan + intSpan
             
        End If
        
        If intPrevSeq <> intSeq And r > 0 And intSeq > 1 Then
            i = i + 1
            If fintRowInserts > 1 Then
                grdMDIReport.InsertRows lngRow + i + fintRowInserts, 1
            Else
                grdMDIReport.InsertRows lngRow + i, 1
            End If

            fintRowInserts = fintRowInserts + 1
            booRowAdded = True
        End If
        
        If booRowAdded Then
            grdMDIReport.AddCellSpan grdMDIReport.Col, lngRow + i, varSpan(2), 1
        End If
        
        grdMDIReport.Row = lngRow + i + intRowInserts
        booBooleanType = False
        
        If HideRow(lngRow + fintRowInserts) And varRS(2, r) = "" Then grdMDIReport.ROWHEIGHT(lngRow + fintRowInserts) = 0
                         
            c = c + 1
            
            If r > 0 And r < UBound(varRS, 2) Then
                If intSeq <> intPrevSeq Then
                    c = 0
                End If
            End If
    
        'End If
                       
        intPrevSeq = intSeq
                   
    Next r
    Return
    
    grdMDIReport.ReDraw = True
Error_Execute:
    'Debug.Print
    
    If Err.Number = 9 Then
        
        If strItem(0) = "<FetchBooleanComp>" Then
            ReDim fudtAudit(0): Resume Next
        ElseIf strItem(0) = "<ColHeaders>" Then
            ReDim fudtGridHeader(0): Resume Next
        ElseIf strItem(0) = "<FetchRecordSet>" Then
            ReDim fudtRowInsert(0): Resume Next
        End If
    End If
    Exit Sub

    
  Resume
  
End Sub

Private Sub SuppressUnusedCols(vObjGrid As Object, lngCol As Long, lngRow As Long, lngCol2 As Long, lngRow2 As Long)
Dim r As Integer
Dim c As Integer

    For c = lngCol To lngCol2
        
        vObjGrid.Col = c
        vObjGrid.ColHidden = True
        
        For r = lngRow To lngRow2
            vObjGrid.Row = r
            If vObjGrid.Text <> vbNullString And vObjGrid.Text <> "0" Then
                vObjGrid.ColHidden = False
                Exit For
            End If
            
        Next r
    
    Next c
    
End Sub


Private Sub Set_Block(vlngCol, vlngRow, vlngCol2, vlngRow2)

    With grdMDIReport
        .BlockMode = True
        .Col = vlngCol
        .Row = vlngRow
        .Col2 = vlngCol2
        .Row2 = vlngRow2
    End With
    
End Sub
Private Sub Form_Paint()
    ColorBurst Me, "BLUE", 128, 0

End Sub

Private Sub Print_Grid(vgrd As fpSpread, vstrHeader As String)
Dim strNote(1) As String
Dim strNotes As String
Dim i As Integer
    
    grdMDIReport.ReDraw = False

    With vgrd
        .PrintSmartPrint = False
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = vstrHeader
        .PrintHeader = vstrHeader
        .PrintFooter = "/cPage #/p/n"
        .PrintBorder = False
        .PrintColHeaders = True
        .PrintColor = True
        .PrintGrid = True
        .PrintMarginTop = 50
        .PrintMarginBottom = 50
        .PrintMarginLeft = 50
        .PrintMarginRight = 50
        '.PrintType=
        .PrintOrientation = PrintOrientationPortrait
        .PrintRowHeaders = True
        .PrintShadows = False
        .PrintUseDataMax = True
        
        .PrintSheet

    grdMDIReport.ReDraw = True
    End With
End Sub





Private Function IdentifyColumn(vlngCol As Long, vvarObject) As Integer
Dim i As Integer

    For i = LBound(fstrIncludedObject) To UBound(fstrIncludedObject)
    
        If InStr(vvarObject, fstrIncludedObject(i)) > 0 Then
            IdentifyColumn = vlngCol + i
            Exit For
        End If
        
    Next i
    
End Function

Private Sub Get_EncounterData()

    Get_Encounter piEncounterID

    fclsEncounter.StartDate = clsEncounter.StartDate
    fclsEncounter.EndDate = clsEncounter.EndDate
    fclsEncounter.VisitReason = clsEncounter.VisitReason
    fclsEncounter.Facility = clsEncounter.Facility
    fclsEncounter.Provider = clsEncounter.Provider
    fclsEncounter.ServicePlace = clsEncounter.ServicePlace
    fclsEncounter.ServicePlaceCode = clsEncounter.ServicePlaceCode
    fclsEncounter.EncounterType = clsEncounter.EncounterType
Exit Sub
End Sub

Private Sub Remove_ResidualTags()
Dim c!, r!
Dim varCell

    For r = 1 To grdMDIReport.MaxRows
    
        For c = 1 To grdMDIReport.MaxCols
            grdMDIReport.GetText c, r, varCell
            
            If varCell <> vbNullString Then
                If Left(varCell, 1) = "<" And Right(varCell, 2) = "/>" Then
                   grdMDIReport.SetText c, r, ""
                End If
            End If
            
        Next c
        
    Next r
    
End Sub

Private Function HideRow(vlngRow As Long) As Boolean
Dim i As Integer
Dim intRow As Integer
Dim intAbsRow As Integer
Dim intStartRow As Integer


    On Error GoTo Error_HideRow
    
    intAbsRow = vlngRow
    
    If fintRowInserts > 0 Then
        For i = LBound(fudtRowInsert) To UBound(fudtRowInsert)
            
            If vlngRow > fudtRowInsert(i).StartRow And fudtRowInsert(i).StartRow > 0 Then
                intAbsRow = vlngRow - fintRowInserts
                Exit For
            'Else
                'intAbsRow = vlngRow
            End If
        
        Next i
    End If
    
    For i = LBound(flngRowHidden) To UBound(flngRowHidden)
        'If intAbsRow > 33 And intAbsRow < 50 Then
        '    Debug.Print vlngRow 'fudtRowInsert(0).StartRow
        'End If
'
'        If vlngRow > fudtRowInsert(0).StartRow Then
'            intRow = flngRowHidden(i) + fintRowInserts
'        Else
'            intRow = flngRowHidden(i)
'        End If
        
        If flngRowHidden(i) = intAbsRow Then
            HideRow = True
            Exit For
        End If
    Next i
    
    Exit Function
Error_HideRow:
    If Err.Number = 9 Then
        Exit Function
    Else
        ReDim flngRowHidden(0)
    End If
    Exit Function
    
    Resume
End Function

Private Function Get_InsertRowCount(vlngRow As Long) As Integer
Dim g%
     
 '    On Error Resume Next
     On Error GoTo ErrHandler

    For g = LBound(fudtRowInsert) To UBound(fudtRowInsert)
        If vlngRow >= fudtRowInsert(g).StartRow Then
            Get_InsertRowCount = fudtRowInsert(g).InsertedRows
        End If
    Next g
    
    Exit Function
    
ErrHandler:
 Exit Function
    
End Function

Private Sub Get_Encounters(vlngPatientID As Long, Optional vlngEncounterID As Long)
Dim lstrSQL As String
Dim i As Integer
Dim lstrRemarks As String

LogFile "Private Sub Get_Encounters-----------------------------------"
'If lstEncounters.ListIndex = -1 Then Exit Sub
If vlngPatientID = 0 Then Exit Sub

    lstrSQL = "Select (cast(encntr_start_date as varchar)  + ' - ' + case when encntr_end_date is null then '' else cast(encntr_end_date as varchar) end) as dur, encounter_id " _
        & "From ENCOUNTER " _
        & "Where patient_id = " & vlngPatientID & " Order By ENCNTR_START_DATE desc"
        
              
        Load_ComboBox lstEncounters, lstrSQL, psConnect(0)
        
        If vlngEncounterID > 0 Then
            For i = 0 To lstEncounters.ListCount - 1
                If lstEncounters.ItemData(i) = vlngEncounterID Then
                    lstEncounters.ListIndex = i  'lstEncounters.NewIndex
                    Exit For
                End If
            Next i
            
            If lstEncounters.ListIndex = 0 Then cmdNext.Enabled = False
            If lstEncounters.ListIndex = lstEncounters.ListCount - 1 Then cmdPrev.Enabled = False
            mintindex = lstEncounters.ListIndex
        End If
        
    Exit Sub
End Sub

Private Sub mnuReports_Exit_Click()
    Unload Me
End Sub






Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.objects
      Set obj = Nothing
    Next obj
  Exit Sub
End Sub

Private Sub lstEncounters_Click()
Dim i%

    If mboobypass Then mboobypass = False: Exit Sub
    If lstEncounters.ListIndex = -1 Then Exit Sub

    piEncounterID = lstEncounters.ItemData(lstEncounters.ListIndex)

    modSpecific.piPatientID = piPatientID
    modSpecific.piEncounterID = piEncounterID
    Get_Provider_Id
    frmPatients.piProviderID = piProviderID
    Get_EncounterData
    ccdDateFrom.Text = fclsEncounter.StartDate
    Exit Sub
End Sub

Private Sub lstEnctrFrom_Click()
Dim i%
Dim strDate As String

    If mboobypass Then mboobypass = False: Exit Sub
    If lstEnctrFrom.ListIndex = -1 Then Exit Sub

    mlngFromDateEnctrID = lstEnctrFrom.ItemData(lstEnctrFrom.ListIndex)
    ccdDateFrom.Text = vbNullString
    strDate = Format(Now(), "MM/DD/YYYY hh:mm AM/PM")
    strDate = Format(lstEnctrFrom.Text(), "MM/DD/YYYY hh:mm AM/PM")
    ccdDateFrom.Text = Mid$(strDate, 1, 10)
    fraEnctrFrom.Visible = False
    lstPatient.Enabled = True
    lstEncounters.Enabled = True
    
    Exit Sub
End Sub


Private Sub lstEnctrTo_Click()
Dim i%
Dim strDate As String

    If mboobypass Then mboobypass = False: Exit Sub
    If lstEnctrTo.ListIndex = -1 Then Exit Sub
    
    mlngToDateEnctrID = lstEnctrTo.ItemData(lstEnctrTo.ListIndex)
    ccdDateTo.Text = vbNullString
    strDate = Format(Now(), "MM/DD/YYYY hh:mm AM/PM")
    strDate = Format(lstEnctrTo.Text(), "MM/DD/YYYY hh:mm AM/PM")
    ccdDateTo.Text = Mid$(strDate, 1, 10)
    fraEnctrTo.Visible = False
    lstPatient.Enabled = False
    lstEncounters.Enabled = False
    Exit Sub
End Sub


Private Sub lstPatient_Click()
Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim strPatientExtID As String

    'lstEncounters.Clear
    If lstPatient.ListIndex = -1 Then Exit Sub
    Get_Encounters lstPatient.ItemData(lstPatient.ListIndex)   ', piEncounterID
    piPatientID = lstPatient.ItemData(lstPatient.ListIndex)
    piPatientID = piPatientID
    modSpecific.piPatientID = piPatientID
    Get_Patient piPatientID
    'lblPatName.Caption = clsPatient.Name
    lblPrintPat.Caption = clsPatient.Name
    frmMDIReport.Caption = clsPatient.Name & " " & clsPatient.IntId & " " & clsEncounter.StartDate

Exit Sub

End Sub


Public Sub lstReports_Click()
    Reset
    Reset_Grid
    fmstrfile = ""
    
    If fstrLastFilePath = "" Then
        fstrLastFilePath = pstrMainPath & "\Reports"
        If Not FolderExists(fstrLastFilePath) Then
            MkDir fstrLastFilePath
       End If
    End If
    
    'fmstrfile = CommonDialogAction(CommonDialog, "Open", "Report Script", fstrLastFilePath, "*.txt")

    fmstrfile = fstrLastFilePath & "\" & lstReports.FileName
    
    If fmstrfile <> vbNullString Then
       ' fstrReportContent = File_Method(fmstrFile, "open")
        cmdRun_Click
    End If
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim k As Integer
Dim lRow As Long
Dim sItem As String

    'If optByInternalId.Value Or optByExternalId.Value Then
       ' SearchList lstPatient, "(" & txtSearch.Text & ")", False
   ' Else
        SearchList lstPatient, txtSearch.Text, True
   ' End If
End Sub




Private Sub Set_Patient()
Dim i

Load_Patients
    For i = 0 To lstPatient.ListCount - 1
        If lstPatient.ItemData(i) = piPatientID Then
            lstPatient.Text = lstPatient.List(i)
            Exit For
        End If
    Next i
    
    For i = 0 To lstEncounters.ListCount - 1
        If lstEncounters.ItemData(i) = piEncounterID Then
            lstEncounters.Text = lstEncounters.List(i)
            Exit For
        End If
    Next i
    
Exit Sub

End Sub

