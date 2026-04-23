VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmOrders 
   Caption         =   "Patient Orders"
   ClientHeight    =   9570
   ClientLeft      =   5205
   ClientTop       =   2580
   ClientWidth     =   13770
   Icon            =   "frmOrders.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   13770
   Begin VB.Frame fraGridTerm 
      BackColor       =   &H00E0E0E0&
      Height          =   3705
      Left            =   7560
      TabIndex        =   13
      ToolTipText     =   "Frame is movable with mouse!"
      Top             =   1980
      Visible         =   0   'False
      Width           =   5145
      Begin VB.TextBox cctSearch 
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
         Left            =   90
         TabIndex        =   0
         ToolTipText     =   "Enter either Name or Code"
         Top             =   390
         Visible         =   0   'False
         Width           =   3945
      End
      Begin PT_XP_Button.PT_XP cmdGridTerms 
         Height          =   315
         Left            =   4440
         TabIndex        =   14
         ToolTipText     =   "Terms Maintanence"
         Top             =   3300
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
         Picture         =   "frmOrders.frx":09EA
         AppearanceThemes=   3
         BorderWidth     =   2
         TransparentColor=   14215660
         Caption         =   ".."
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmOrders.frx":0F84
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdDone 
         Height          =   315
         Left            =   2070
         TabIndex        =   15
         Top             =   3270
         Width           =   885
         _ExtentX        =   1561
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
         Picture         =   "frmOrders.frx":12D6
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Done"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmOrders.frx":1670
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdTerms 
         Height          =   2355
         Left            =   90
         TabIndex        =   1
         Top             =   870
         Width           =   4935
         _Version        =   458752
         _ExtentX        =   8705
         _ExtentY        =   4154
         _StockProps     =   64
         ColHeaderDisplay=   0
         DAutoCellTypes  =   0   'False
         DAutoFill       =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
         DisplayColHeaders=   0   'False
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
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   2
         MaxRows         =   10
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   16761024
         ShadowText      =   8388608
         SpreadDesigner  =   "frmOrders.frx":19C2
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin VB.Label lblSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Type in desired diagnosis"
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
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   2115
      End
   End
   Begin VB.Frame fraSign 
      BackColor       =   &H00E0E0E0&
      Height          =   2835
      Left            =   2400
      TabIndex        =   25
      Top             =   2700
      Visible         =   0   'False
      Width           =   4605
      Begin PT_XP_Button.PT_XP cmdErase 
         Height          =   420
         Left            =   3420
         TabIndex        =   31
         Top             =   1500
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
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
         Picture         =   "frmOrders.frx":1CC4
         AppearanceMode  =   1
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Erase"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmOrders.frx":225E
         ShowFocus       =   -1  'True
      End
      Begin PatientTrac.ccApprove ccaSignature 
         Height          =   1155
         Left            =   120
         TabIndex        =   30
         Top             =   870
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   2037
      End
      Begin VB.ComboBox cboAttendPhy 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   360
         Width           =   2985
      End
      Begin PT_XP_Button.PT_XP cmdSigExit 
         Height          =   345
         Left            =   2550
         TabIndex        =   32
         ToolTipText     =   "Save, Restore, Close"
         Top             =   2310
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
         Picture         =   "frmOrders.frx":25B0
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   15790320
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmOrders.frx":2FC2
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSaveSign 
         Height          =   345
         Left            =   1470
         TabIndex        =   33
         ToolTipText     =   "Save, Restore, Close"
         Top             =   2310
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
         Picture         =   "frmOrders.frx":3314
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   15790320
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmOrders.frx":36AE
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblMedDtr 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Order Physician"
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
         Height          =   255
         Left            =   150
         TabIndex        =   28
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblPhys 
         BackStyle       =   0  'Transparent
         Caption         =   "Physician"
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
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   -210
         Width           =   1365
      End
   End
   Begin VB.PictureBox picLogo 
      BackColor       =   &H00FFFFFF&
      Height          =   7740
      Left            =   60
      ScaleHeight     =   7680
      ScaleWidth      =   13530
      TabIndex        =   17
      Top             =   1020
      Width           =   13590
      Begin VB.PictureBox picPatientTrac 
         Height          =   6075
         Left            =   2700
         Picture         =   "frmOrders.frx":3A00
         ScaleHeight     =   6015
         ScaleWidth      =   7635
         TabIndex        =   18
         Top             =   1065
         Width           =   7695
      End
      Begin VB.Label lblPatientTrac 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "PATIENTTRAC ORDERS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   585
         Left            =   2190
         TabIndex        =   19
         Top             =   60
         Width           =   8505
      End
   End
   Begin VB.Frame fraRecICD 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Suggested Reasons  Lab Tests"
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
      Height          =   2415
      Left            =   5070
      TabIndex        =   9
      Top             =   1530
      Visible         =   0   'False
      Width           =   6585
      Begin PT_XP_Button.PT_XP cmdSelect 
         Height          =   315
         Left            =   2010
         TabIndex        =   11
         Top             =   2010
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Select"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin VB.ListBox lstICD 
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
         Height          =   1500
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   450
         Width           =   6135
      End
      Begin PT_XP_Button.PT_XP cmdCloseList 
         Height          =   315
         Left            =   3180
         TabIndex        =   12
         ToolTipText     =   "Close List to select ICD code manaully"
         Top             =   2010
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
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
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Close List"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
   End
   Begin PT_XP_Button.PT_XP cmdOrderRemove 
      Height          =   435
      Left            =   1155
      TabIndex        =   2
      Top             =   540
      Width           =   1215
      _ExtentX        =   2143
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
      Picture         =   "frmOrders.frx":11AB2
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Remove"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":11E4C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdLAB 
      Height          =   465
      Left            =   7830
      TabIndex        =   3
      Top             =   540
      Width           =   1515
      _ExtentX        =   2672
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
      Picture         =   "frmOrders.frx":1219E
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Print Laboratory"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":12738
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdAddOrder 
      Height          =   435
      Left            =   30
      TabIndex        =   4
      Top             =   540
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmOrders.frx":12A8A
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":13024
      ShowFocus       =   -1  'True
   End
   Begin VB.PictureBox picSignature 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   8970
      ScaleHeight     =   0.63
      ScaleMode       =   0  'User
      ScaleWidth      =   5.16
      TabIndex        =   5
      Top             =   10050
      Visible         =   0   'False
      Width           =   2985
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   435
      Left            =   6870
      TabIndex        =   6
      Top             =   8895
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmOrders.frx":13376
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":13D88
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   435
      Left            =   5760
      TabIndex        =   7
      Top             =   8910
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmOrders.frx":140DA
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":14474
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdTerms 
      Height          =   315
      Left            =   13170
      TabIndex        =   8
      Top             =   9000
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
      Picture         =   "frmOrders.frx":147C6
      AppearanceThemes=   3
      BorderWidth     =   2
      TransparentColor=   14215660
      Caption         =   ".."
      DepthMode       =   1
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":14D60
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRAD 
      Height          =   465
      Left            =   9450
      TabIndex        =   20
      Top             =   540
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
      Picture         =   "frmOrders.frx":150B2
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Print Radiology"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":1564C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPROC 
      Height          =   465
      Left            =   10935
      TabIndex        =   21
      Top             =   540
      Width           =   1575
      _ExtentX        =   2778
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
      Picture         =   "frmOrders.frx":1599E
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Print Procedure"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":15F38
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdALL 
      Height          =   465
      Left            =   12600
      TabIndex        =   22
      Top             =   540
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmOrders.frx":1628A
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Print All"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":16824
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSign 
      Height          =   465
      Left            =   6180
      TabIndex        =   23
      Top             =   540
      Width           =   1605
      _ExtentX        =   2831
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
      Picture         =   "frmOrders.frx":16B76
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Signature Pad"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":17588
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdSS 
      CausesValidation=   0   'False
      DragIcon        =   "frmOrders.frx":178DA
      Height          =   7755
      Left            =   60
      TabIndex        =   24
      Top             =   1020
      Width           =   13605
      _Version        =   458752
      _ExtentX        =   23998
      _ExtentY        =   13679
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   8388608
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   8
      MaxRows         =   0
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmOrders.frx":181A4
      UnitType        =   0
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   10
      ScrollBarTrack  =   1
      ClipboardOptions=   0
      ShowScrollTips  =   1
   End
   Begin PT_XP_Button.PT_XP cmdPatient 
      Height          =   435
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   1515
      _ExtentX        =   2672
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
      Picture         =   "frmOrders.frx":184BC
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "SELECT PATIENT"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":18A56
      PictureAlignment=   0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdStandNew 
      Height          =   465
      Left            =   2790
      TabIndex        =   34
      Top             =   540
      Width           =   1485
      _ExtentX        =   2619
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
      Picture         =   "frmOrders.frx":18DA8
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "New Patient"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":19342
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdEstablish 
      Height          =   465
      Left            =   4350
      TabIndex        =   35
      Top             =   540
      Width           =   1485
      _ExtentX        =   2619
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
      Picture         =   "frmOrders.frx":19694
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Established Patient"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":19C2E
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDefault 
      Height          =   315
      Left            =   12420
      TabIndex        =   36
      ToolTipText     =   "Default Orders"
      Top             =   9000
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
      Picture         =   "frmOrders.frx":19F80
      AppearanceThemes=   3
      BorderWidth     =   2
      TransparentColor=   15790320
      Caption         =   ".."
      DepthMode       =   1
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmOrders.frx":1A51A
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mstrRemarks As String
Private mbooByPass As Boolean
Private mobjParentParent As Object
Private mParentRef As String
Private mReferId As String
Private mRefId() As Long
Private mPhone() As String
Private gbooBypass As Boolean
Private mstrRefData As String
Private mlngGridRow As Long
Private mlngGridCol As Long
Private mstrItemSelected As String
Private mRow As Long
Private mintRefRow As Integer
Private mstrSignatureToken As String
Private mlngInsertRows As Long
Public mstrCPTCode As String
Public mintComponentID As Integer
Public mintObjPoolID As Integer
Private mFirst As Boolean
Public mstrSubGroup As String
Private mbooPrintPreview As Boolean
Private mlngStartRow As Long
Private mstrLabCode As String
Private flngDeltaX As Long
Private flngDeltaY As Long
Private mbooNoSave As Boolean
Private fpDeltaLeft As Long
Private fpDeltaTop As Long
Private mlngHeadRow As Long
Private mstrComment As String
Private udtEnctr As udtEnctr
Private Type udtEnctr
    EnctrID As Long
    StartDate As String
    EndDate As String
    VisitReason As String
    Facility As String
    Provider As String
    Admission As String
    Discharge As String
    ServicePlace As String
    ServicePlaceCode As String
    EncounterType As String
    Appointment As String
    ProviderID As Long
    VisitType As String
    Intrvn_CD As String
    VisitTime As String
    FacilityID As Long
End Type
Private Sub Clear_EnctrString()
    With udtEnctr
        .EnctrID = 0
        .StartDate = vbNullString
        .EndDate = vbNullString
        .VisitReason = vbNullString
        .Facility = vbNullString
        .Provider = vbNullString
        .Admission = vbNullString
        .Discharge = vbNullString
        .ServicePlace = vbNullString
        .ServicePlaceCode = vbNullString
        .EncounterType = vbNullString
        .Appointment = vbNullString
        .ProviderID = 0
        .FacilityID = 0
        .VisitType = vbNullString
        .Intrvn_CD = vbNullString
        .VisitTime = vbNullString
    End With
End Sub
Private Function Get_CurrentEnctr(Optional vlngEnctr As Long)
Dim sDate As String
Dim strSQL As String
Dim cnStatus As New ADODB.Connection
Dim rsStatus As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim strEnctTime As String
Dim lngEnctrID As Long

    On Error GoTo Error_Get_CurrentEnctr
    Screen.MousePointer = vbHourglass

    Clear_Provider
    If piProviderID = 0 Then
        piProviderID = mlngLogonUser
    End If
    If vlngEnctr > 0 Then
        lngEnctrID = vlngEnctr
    End If
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    Get_Providers piProviderID
    If lngEnctrID = 0 Then
         If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
         strSQL = "SELECT * FROM ENCOUNTER WHERE PATIENT_ID = " & udtPatientName.IntId & " AND " _
                 & "PROVIDER_ID = " & piProviderID & " AND PURGE = 0 "
         Set rsStatus = cnStatus.Execute(strSQL)
         
         With rsStatus
             Do Until .EOF
                 If (!ENCNTR_START_DATE) <> vbNullString Then strEnctTime = (!ENCNTR_START_DATE)
                 If Trim(strEnctTime) <> vbNullString Then
                     strEnctTime = Format(strEnctTime, "mm/dd/yyyy")
                     lngDateDiff = DateDiff("d", strEnctTime, Format(sDate, "mm/dd/yyyy"))
                     If lngDateDiff = 0 Then
                         lngEnctrID = (!ENCOUNTER_ID)
                         Get_Enctr lngEnctrID
                         Exit Do
                     End If
                 End If
             .MoveNext
             Loop
         End With
    Else
        udtEnctr.EnctrID = lngEnctrID
        Get_Enctr udtEnctr.EnctrID
    End If
    If lngEnctrID = 0 Then
        udtEnctr.VisitType = "Activities Of Patient Care (patient Not Present)"
        udtEnctr.FacilityID = piFacilityID
        udtEnctr.EncounterType = "Established Patient"
        If PsUserID = vbNullString Then
            PsUserID = Get_PsUserID(psConnect(0))
        End If
        udtEnctr.Intrvn_CD = "99361"
        If udtEnctr.FacilityID > 0 Then
            Get_Facilities udtEnctr.FacilityID
        End If
        If clsFacility.FacType <> vbNullString Then
            udtEnctr.ServicePlace = clsFacility.FacType
        End If
        If clsFacility.pos <> 0 Then
            udtEnctr.ServicePlaceCode = clsFacility.pos
        End If
        If udtEnctr.ServicePlace = vbNullString Then udtEnctr.ServicePlace = "Office"
        If Val(udtEnctr.ServicePlaceCode) = 0 Then udtEnctr.ServicePlaceCode = "11"
        strSQL = "INSERT INTO ENCOUNTER (PATIENT_ID, PROVIDER_ID, FACILITY_ID, VISIT_TYPE, SERVICE_PLACE," _
             & "ENCNTR_TYPE, ENCOUNT_CPT_INTRVN, RENDERING_PROVIDER, ENCNTR_START_DATE, ENCOUNT_APPT_TYPE, INSERT_DATE, SERVICE_PLACE_CD, USER_ID, PURGE) " _
             & "values  (" & udtPatientName.IntId & "," & piProviderID & "," & udtEnctr.FacilityID & ",'" _
             & udtEnctr.VisitType & "','" _
             & udtEnctr.ServicePlace & "','" _
             & udtEnctr.EncounterType & "','" _
             & udtEnctr.Intrvn_CD & "'," & piProviderID & ", '" & sDate & "','" _
             & udtEnctr.VisitType & "','" & sDate & "'," & udtEnctr.ServicePlaceCode & ",'" & PsUserID & "',0)"
            Set rsStatus = cnStatus.Execute(strSQL)
        strSQL = "SELECT  ENCOUNTER_ID FROM ENCOUNTER   " _
              & "WHERE ENCNTR_START_DATE = '" & sDate & "' AND " _
              & "PATIENT_ID = " & udtPatientName.IntId & " AND PURGE = 0"
            Set rsStatus = cnStatus.Execute(strSQL)
            If Not rsStatus.EOF Then
                 udtEnctr.EnctrID = rsStatus.Fields("ENCOUNTER_ID")
            End If
            Get_Enctr udtEnctr.EnctrID
    End If
    If udtEnctr.ServicePlaceCode = "9" Then udtEnctr.ServicePlaceCode = "09"
    If cnStatus.State = adStateOpen Then
        cnStatus.Close
        Set cnStatus = Nothing
        Set rsStatus = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_CurrentEnctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_CurrentEnctr", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Add_Signature()
Dim lngRow As Long
Dim strPhoto As String
Dim Arr
Dim fso As New FileSystemObject

    On Error GoTo Error_Add_Signature
    Screen.MousePointer = vbHourglass
    If clsSignature.Signature = vbNullString Then
        InitSignature
    End If
    If Trim(clsProvider.Name) = vbNullString Then
        Get_Providers piProviderID
    End If
    Arr = Split(LookUpCellTag("Physician_Signature"), "|")
    If UBound(Arr) > 0 Then
        lngRow = Arr(1)
    End If
        
    With grdSS
        If lngRow = 0 Then
            lngRow = .MaxRows + 2
            .MaxRows = lngRow
            If Trim(clsSignature.Signature) <> vbNullString Then
                picSignature.Width = 4000
                picSignature.Height = 1000
                picSignature.BackColor = NormalColor
                picSignature.Picture = ccaSignature.Picture
                SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                If fso.FileExists(App.Path & "\Images\SignNote.bmp") Then
                    SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                    strPhoto = App.Path & "\Images\SignNote.bmp"
                End If
                Format_CellTypeSignature grdSS, lngRow, 1, "4", strPhoto, "TypeHAlignCenter", "Physician_Signature", 47
            Else
                Format_CellTypeStatic grdSS, lngRow, 1, "4", "PHYSICIAN SIGNATURE", "TypeHAlignCenter", "TypeVAlignCenter", 0, "", "Physician_Signature"
            End If
            lngRow = lngRow + 1
            .MaxRows = lngRow
            Format_CellTypeStatic grdSS, lngRow, 1, "4", clsProvider.Name, "TypeHAlignCenter"
            lngRow = .MaxRows + 1
            .MaxRows = lngRow
            Format_CellTypeStatic grdSS, lngRow, 1, "4", "Thank You - Please contact our office for any additional  information.", "TypeHAlignCenter"
        Else
            If Trim(clsSignature.Signature) <> vbNullString Then
                picSignature.Width = 4000
                picSignature.Height = 1000
                picSignature.BackColor = NormalColor
                picSignature.Picture = ccaSignature.Picture
                SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                If fso.FileExists(App.Path & "\Images\SignNote.bmp") Then
                    SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                    strPhoto = App.Path & "\Images\SignNote.bmp"
                End If
                Format_CellTypeSignature grdSS, lngRow, 1, "4", strPhoto, "TypeHAlignCenter", "Physician_Signature", 47
            Else
                Format_CellTypeStatic grdSS, lngRow, 1, "4", "PHYSICIAN SIGNATURE", "TypeHAlignCenter", "Physician_Signature"
            End If
        End If
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_Add_Signature:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Add_Signature", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Private Function Get_Enctr(vlngEncounterID As Long)
Dim sDate As String
Dim vDate As String
Dim lngDateDiff As Long
Dim cnEncnt As New ADODB.Connection
Dim rsEncnt As New ADODB.Recordset
Dim Fld As ADODB.Field
Dim Flds As ADODB.Fields
Dim strSQL As String
Dim lngEnctrID As Long
Dim lngPatID As Long

    On Error GoTo Error_Get_Enctr
    lngEnctrID = vlngEncounterID
    If lngEnctrID = 0 Then Exit Function
    Screen.MousePointer = vbHourglass
    Clear_EnctrString
    If cnEncnt.State = adStateClosed Then cnEncnt.open psConnect(0)
    strSQL = "Select * FROM ENCOUNTER WHERE ENCOUNTER_ID = " & lngEnctrID & " AND PURGE = 0 Order by  ENCNTR_START_DATE DESC "
    Set rsEncnt = cnEncnt.Execute(strSQL)
    If Not rsEncnt.EOF Then
        Set Flds = rsEncnt.Fields
        For Each Fld In Flds
            If Not IsNull(Fld.Value) Then
                If Trim(Fld.Value) <> vbNullString Then
                    Select Case UCase(Fld.Name)
                        Case "FACILITY_ID"
                            udtEnctr.FacilityID = Fld.Value
                        Case "ENCNTR_START_DATE"
                             udtEnctr.StartDate = Fld.Value
                        Case "ENCNTR_END_DATE"
                             udtEnctr.EndDate = Fld.Value
                        Case "VISIT_REASON"
                             udtEnctr.VisitReason = Fld.Value
                        Case "SERVICE_PLACE"
                             udtEnctr.ServicePlace = Fld.Value
                        Case "SERVICE_PLACE_CD"
                             udtEnctr.ServicePlaceCode = Fld.Value
                        Case "ENCNTR_TYPE"
                             udtEnctr.EncounterType = Fld.Value
                        Case "PATIENT_ID"
                            lngPatID = Fld.Value
                        Case "NEXTAPPT"
                             udtEnctr.Appointment = Fld.Value
                        Case "PROVIDER_ID"
                             udtEnctr.ProviderID = Fld.Value
                        Case "VISIT_TYPE"
                            udtEnctr.VisitType = Fld.Value
                        Case "ENCOUNTER_ID"
                            udtEnctr.EnctrID = Fld.Value
                    End Select
                    If udtEnctr.Intrvn_CD <> vbNullString Then
                        If TextFound("Initial Consultation", udtEnctr.EncounterType, Len("Initial Consultation"), True) Or Trim(udtEnctr.Intrvn_CD) = "90801" Then
                            mstrMedReason = "Pre Surgery"
                        ElseIf Trim(udtEnctr.Intrvn_CD) = "99242" Or TextFound("Follow-Up", udtEnctr.EncounterType, Len("Follow-Up"), True) Then
                            mstrMedReason = "Surgery Follow-Up"
                        End If
                    End If
                End If
            End If
        Next
    End If
    If udtEnctr.ProviderID > 0 Then
        Get_ProviderName (udtEnctr.ProviderID)
    Else
        Get_DefaultProvider
        Get_ProviderName piProviderID
    End If
    If udtEnctr.FacilityID > 0 Then Get_EnctFacName (udtEnctr.FacilityID)
    If udtEnctr.ServicePlaceCode = "9" Then udtEnctr.ServicePlaceCode = "09"
    If cnEncnt.State = adStateOpen Then
        cnEncnt.Close
        Set cnEncnt = Nothing
        Set rsEncnt = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Enctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Enctr", App.EXEName & "." & "ModPatientData"
    Exit Function
    Resume
End Function

Private Function Remove_Signature()
Dim lngRow As Long
Dim Arr

    On Error GoTo Error_Remove_Signature
    Screen.MousePointer = vbHourglass
    Arr = Split(LookUpCellTag("Physician_Signature"), "|")
    If UBound(Arr) > 0 Then
        lngRow = Arr(1)
        grdSS.ClearRange 1, lngRow, grdSS.MaxCols, grdSS.MaxRows, False
        grdSS.MaxRows = lngRow - 2
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Remove_Signature:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "Remove Error", "Remove_Signature", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Private Function Set_GridCurSel(vObjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
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
            If TextFound(Trim(vstrData), Trim(.TypeComboBoxString), Len(Trim(vstrData)), True) Then
                .TypeComboBoxCurSel = Idx
                Exit For
            End If
            Next Idx
        End If
        .TypeVAlign = TypeVAlignCenter
        If vlngBackColor > 0 Then .BackColor = vlngBackColor
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
        If vstrCellNote <> vbNullString Then .CellNote = Trim(vstrCellNote)
    .ReDraw = True
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
vstrString As String, vstrHAligh As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .ReDraw = False
        .BlockMode = True
              .Col = vlngCol
              .Col2 = vlngCol
              .Row = vlngRow
              .Row2 = vlngRow
              .CellType = CellTypeComboBox
              .TypeMaxEditLen = 32000
              .TypeComboBoxList = vstrString
              .TypeComboBoxWidth = 1
              If vlngBackColor > 0 Then
                 .BackColor = vlngBackColor
              End If
              If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
              If vstrCellTag <> vbNullString Then
                .CellTag = vstrCellTag
              End If
              .TypeHAlign = TypeHAlignCenter
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
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
          .BlockMode = False
        .ReDraw = True
    End With
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional vlngBackColor As Long, Optional vlngFontSize As Long)
Dim CellNote As Long
Dim lngSpan As Long

        
    With vGrid
        .ReDraw = False
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeEdit
            .TypeMaxEditLen = 32000
            .SetText vlngCol, vlngRow, vstrText
            .TypeEditMultiLine = True   'Make multi-lines
            .FontBold = True
            If vlngFontSize > 0 Then .FontSize = vlngFontSize
            'If vbooUnderLine Then .FontUnderline = True
            .Text = vstrText
            If vlngBackColor Then .BackColor = &HE0E0E0
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
        .ReDraw = True
    End With
End Function
Private Function Format_CellTypePicture(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long)
Dim CellNote As Long
Dim lngSpan As Long
Dim fso As New FileSystemObject

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypePicture
           .TypePictMaintainScale = True
           .TypePictStretch = True
            .TypePictCenter = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf vstrSpan <> vbNullString Then
               lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
             End If
            If fso.FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            If vlngRowHeight > 0 Then .RowHeight(vlngRow) = vlngRowHeight
          .BlockMode = False
    End With
End Function

Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrVAlign As String, Optional lngBackColor As Long, Optional vstrCellNote As String, Optional vstrCellTag As String)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .ReDraw = False
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
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
            If lngBackColor > 0 Then .BackColor = lngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        .BlockMode = False
        .ReDraw = True
    End With
End Function
Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)
    With vObjGrid
        vlngRow = vlngRow
       If .MaxRows < vlngRow Then .MaxRows = vlngRow
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
Private Function Format_CellTypeSignature(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long)
Dim CellNote As Long
Dim lngSpan As Long
Dim fso As New FileSystemObject

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypePicture
            .TypePictCenter = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf vstrSpan <> vbNullString Then
               lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
             End If
            If fso.FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
            End If
            .CellTag = "Physician_Signature"
            .RowHeight(vlngRow) = 49
          .BlockMode = False
    End With
End Function
Private Function LookUpCellTag(vstrname As String) As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim booExit As Boolean

    On Error GoTo Error_LookUpCellTag
    Screen.MousePointer = vbHourglass
    With grdSS
        LookUpCellTag = vbNullString
        For lngCol = 1 To .MaxCols
            For lngRow = 1 To .MaxRows
                .Col = lngCol
                .Row = lngRow
                If .CellTag <> vbNullString Then
                    If TextFound(vstrname, .CellTag, Len(vstrname), True) Then
                        LookUpCellTag = lngCol & "|" & lngRow
                        booExit = True
                        Exit For
                    End If
                End If
            Next lngRow
            If booExit Then Exit For
        Next lngCol
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUpCellTag:
    Screen.MousePointer = vbDefault
   Exit Function
    Resume
End Function



Private Function Format_GridTerms(vstrType As String)

    Screen.MousePointer = vbHourglass
    With grdTerms
        .ReDraw = False
        .ClearRange 0, SpreadHeader, grdTerms.MaxCols, grdTerms.MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .ColHeaderRows = 2
        .ColHeadersShow = True
        .RowHeadersShow = False
        .ColWidth(1) = 28
        .ColWidth(2) = 10
        If vstrType = "ICD" Then
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "DIAGNOSIS", "TypeHAlignCenter", "ICD_DESC", "", &HC0FFFF
            Format_CellTypeEdit grdTerms, SpreadHeader, 2, ".MaxCols", "DX CODE", "TypeHAlignCenter", "ICD_CODE", "", &HC0FFFF
            grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        ElseIf vstrType = "MODIFIER" Then
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "MODIFIER", "TypeHAlignCenter", "MODIFIER", "", &HC0FFFF
            Format_CellTypeEdit grdTerms, SpreadHeader, 2, ".MaxCols", "CPT CODE", "TypeHAlignCenter", "", "", &HC0FFFF
            grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        ElseIf vstrType = "CPT_CODE" Then
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "CPT DESC", "TypeHAlignCenter", "CPT DESC", "", &HC0FFFF
            Format_CellTypeEdit grdTerms, SpreadHeader, 2, ".MaxCols", "CPT CODE", "TypeHAlignCenter", "CPT_CODE", "", &HC0FFFF
            grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        Add_Spacer SpreadHeader + 1, grdTerms
        grdTerms.MaxRows = 0
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Function
End Function

Private Function Get_LabDetails(vstrCPT As String)
Dim i%
Dim strSQL As String
Dim strCode As String
Dim strDesc As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset

     On Error GoTo ERROR_Get_LabDetails
    If mstrLabCode <> vbNullString Then
        lstICD.Clear
    End If
    If vstrCPT = vbNullString Then Exit Function
    If cnLab.State = adStateClosed Then cnLab.open psConnect(1)
    strSQL = "SELECT SUG_ICD_CODE, SUG_ICD_DESC FROM INTRVN_TEMPLATE WHERE " _
        & "CPT_INTRVN_DESC like  '%" & vstrCPT & "%'"
    Set rsLab = cnLab.Execute(strSQL)
    With rsLab
        Do Until .EOF
            If Not IsNull(!SUG_ICD_Desc) Then strDesc = LCase((!SUG_ICD_Desc))
            If Not IsNull(!SUG_ICD_Code) Then strCode = LCase((!SUG_ICD_Code))
        .MoveNext
            If mstrLabCode = vbNullString Then
                If Trim(strDesc) <> vbNullString And Trim(strCode) <> vbNullString Then
                    mstrLabCode = Trim(strDesc) & "-" & Trim(strCode)
                End If
            Else
                If Trim(strDesc) <> vbNullString And Trim(strCode) <> vbNullString Then
                     mstrLabCode = mstrLabCode & Chr(9) & Trim(strDesc) & "-" & Trim(strCode)
                End If
            End If
        Loop
    End With
    If cnLab.State = adStateOpen Then
        cnLab.Close
        Set cnLab = Nothing
        Set rsLab = Nothing
    End If
    Exit Function
ERROR_Get_LabDetails:
    Write_Error Err.Description, Err.Number, strSQL, "Get_LabDetails", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Load_GridBox(vstrCol As String, vstrType As String)
Dim strSQL As String
Dim iRow As Long
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim iCnt As Integer
Dim strData As String
Dim lngRow As Long
Dim Arr
Dim booData As Boolean


    On Error GoTo Error_Load_GridBox
    Screen.MousePointer = vbHourglass
    If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
    Select Case Trim(vstrType)
        Case "CPT_CODE"
            Format_GridTerms "CPT_CODE"
            Load_CPTIntrvn vstrCol
        Case "ICD"
            Format_GridTerms "ICD"
            If Trim(cctSearch.Text) <> vbNullString Then
                If Trim(cctSearch.Text) Like "[A-Za-z]###" Or Trim(cctSearch.Text) Like "###" Then
                    grdTerms.ReDraw = False
                    strSQL = "SELECT LONG_DESC, CODE FROM ICD_CODE WHERE CODE like '" & Trim(cctSearch.Text) & "%' ORDER BY CODE"
                    Set rsCPT = cnCPT.Execute(strSQL)
                    With rsCPT
                        Do Until .EOF
                            If Not IsNull(!Long_Desc) Then strData = (!Long_Desc)
                            If Trim(strData) <> vbNullString Then
                                booData = True
                                lngRow = lngRow + 1
                                grdTerms.MaxRows = lngRow
                                Format_CellTypeEdit grdSS, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "ICD_DESC"
                                If Not IsNull(!code) Then
                                    If IsNumeric(!code) Then grdTerms.SetRowItemData (lngRow), (!code)
                                    If Not IsNull((!code)) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!code), "TypeHAlignCenter", "ICD_CODE"
                                    grdTerms.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                End If
                            End If
                            .MoveNext
                        Loop
                    End With
                Else
                    strSQL = "select LONG_DESC, CODE from icd_code Where LONG_DESC like  '%" & Trim(cctSearch.Text) & "%' ORDER BY LONG_DESC "
                    Set rsCPT = cnCPT.Execute(strSQL)
                    With rsCPT
                        Do Until .EOF
                            If (!Long_Desc) <> vbNullString Then strData = (!Long_Desc)
                            If Trim(strData) <> vbNullString Then
                                booData = True
                                lngRow = lngRow + 1
                                grdTerms.MaxRows = lngRow
                                Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "ICD_DESC"
                                If Not IsNull(!code) Then
                                    If IsNumeric(!code) Then grdTerms.SetRowItemData (lngRow), (!code)
                                    If Not IsNull((!code)) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!code), "TypeHAlignCenter", "ICD_CODE"
                                    grdTerms.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                End If
                            End If
                         .MoveNext
                        Loop
                    End With
                End If
            End If
        Case "MODIFIER"
            grdTerms.ReDraw = False
            Format_GridTerms "MODIFIER"
            strSQL = "select DISTINCT CPT_MOD_DESC, CPT_MOD_CODE " _
                & "From CPT_MOD_TEMPLATE  "
            Set rsCPT = cnCPT.Execute(strSQL)
            With rsCPT
                Do Until .EOF
                    If Not IsNull(!CPT_MOD_DESC) Then strData = (!CPT_MOD_DESC)
                    If Trim(strData) <> vbNullString Then
                        lngRow = lngRow + 1
                        grdTerms.MaxRows = lngRow
                        Format_CellTypeEdit grdTerms, lngRow, 1, "1", IsProperCase(Trim(strData)), "TypeHAlignLeft", "MODIFIER"
                        If IsNumeric(!CPT_MOD_CODE) Then grdTerms.SetRowItemData (lngRow), (!CPT_MOD_CODE)
                        If Not IsNull(!CPT_MOD_CODE) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!CPT_MOD_CODE), "TypeHAlignLeft", "MODIFIER"
                        grdTerms.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                    .MoveNext
                Loop
            End With
    End Select
    grdTerms.ReDraw = True
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_GridBox:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_GridBox", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Load_CPTIntrvn(vstrSubGroup As String)
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim iCnt As Integer
Dim strData As String
Dim lngRow As Long
Dim mintGroupID As Integer
Dim mintLabGroupID As Integer

    On Error GoTo Error_Load_CptIntrvn
   Screen.MousePointer = vbHourglass
    If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
    strSQL = "Select  a.Group_id, b.SUBGROUP_ID " _
            & "From INTRVN_GROUPS a, INTRVN_SUBGROUPS b " _
            & "Where a.GROUP_DESC Like 'Superbill' And b.SUBGROUP_DESC Like '" & vstrSubGroup & "'"
        Set rsCPT = cnCPT.Execute(strSQL)
        With rsCPT
            If Not .EOF Then
                mintGroupID = (!Group_ID)
                mintLabGroupID = (!SUBGROUP_ID)
            End If
        End With
    strSQL = "Select CPT_INTRVN_DESC,  CPT_INTRVN_CODE " _
            & "From INTRVN_TEMPLATE " _
            & "Where group_id = " & mintGroupID & " And SUBGROUP_ID = " & mintLabGroupID _
            & " Order By SUBGROUP_ID "
       Set rsCPT = cnCPT.Execute(strSQL)
       With rsCPT
            Do Until .EOF
                If Not IsNull(!CPT_Intrvn_Desc) Then strData = (!CPT_Intrvn_Desc)
                If Trim(strData) <> vbNullString Then
                    lngRow = lngRow + 1
                    grdTerms.MaxRows = lngRow
                    Format_CellTypeEdit grdTerms, lngRow, 1, "1", Trim(strData), "TypeHAlignLeft", vstrSubGroup
                    If IsNumeric(!CPT_Intrvn_Code) Then grdTerms.SetRowItemData (lngRow), (!CPT_Intrvn_Code)
                    If Not IsNull(!CPT_Intrvn_Code) Then Format_CellTypeEdit grdTerms, lngRow, 2, "1", (!CPT_Intrvn_Code), "TypeHAlignCenter", vstrSubGroup
                    grdTerms.SetCellBorder 1, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    grdTerms.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
                .MoveNext
            Loop
        End With
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_CptIntrvn:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Lab", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Write_Data()
Dim strSQL As String
Dim intSeq As Integer
Dim lngRow As Long
Dim varCellValue
Dim lngLabID As Long
Dim lngPrimKey As Long
Dim sDate As String
Dim strTable As String
Dim lngCol As Long
Dim strMultiDiag As String
Dim Arr
Dim ArrICD
Dim i%
Dim strBullet As String
Dim strMessage As String

    On Error GoTo Error_Write_Data
    Screen.MousePointer = vbHourglass
    strBullet = "* "
    With grdSS
        For lngRow = mlngHeadRow + 1 To grdSS.MaxRows
            Clear_Order
            grdSS.GetText 1, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then clsOrders.Type = Trim(varCellValue)
            grdSS.GetText 2, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                Arr = Split(varCellValue, " - ")
                If UBound(Arr) > 0 Then
                    clsOrders.CPT_Code = Trim(Arr(1))
                    clsOrders.CPT_Desc = Trim(Arr(0))
                Else
                    clsOrders.CPT_Desc = Trim(Arr(0))
                End If
            End If
            grdSS.GetText 3, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then clsOrders.Modifier = Trim(varCellValue)
            
            grdSS.GetText 4, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                Arr = Split(varCellValue, vbNewLine)
                If UBound(Arr) > 0 Then
                    For i = LBound(Arr) To UBound(Arr)
                        ArrICD = Split(Arr(i), " - ")
                         If UBound(ArrICD) > 0 Then
                            If clsOrders.ICD_Desc = vbNullString Then
                                clsOrders.ICD_Desc = Trim(ArrICD(0))
                            Else
                                clsOrders.ICD_Desc = Trim(clsOrders.ICD_Desc) & "~" & Trim(ArrICD(0))
                            End If
                            If clsOrders.ICD_Code = vbNullString Then
                               clsOrders.ICD_Code = Trim(ArrICD(1))
                            Else
                                clsOrders.ICD_Code = Trim(clsOrders.ICD_Code) & "~" & Trim(ArrICD(1))
                            End If
                        End If
                    Next i
                Else
                    ArrICD = Split(varCellValue, " - ")
                     If UBound(ArrICD) > 0 Then
                        If clsOrders.ICD_Desc = vbNullString Then
                            clsOrders.ICD_Desc = Trim(ArrICD(0))
                        Else
                            clsOrders.ICD_Desc = Trim(clsOrders.ICD_Desc) & "~" & Trim(ArrICD(0))
                        End If
                        If clsOrders.ICD_Code = vbNullString Then
                            clsOrders.ICD_Code = Trim(ArrICD(1))
                        Else
                            clsOrders.ICD_Code = Trim(clsOrders.ICD_Code) & "~" & Trim(ArrICD(0))
                        End If
                    End If
                End If
            End If
            grdSS.GetText 5, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then clsOrders.Status = Trim(varCellValue)
            grdSS.GetText 6, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then clsOrders.Comment = Trim(varCellValue)
            If TextFound("Lab", clsOrders.Type, Len("Lab"), True) Then
                If clsOrders.ICD_Desc = vbNullString Or clsOrders.ICD_Code = vbNullString Then
                    MsgBox "Labs require a diagnostic code - please enter diagnostic code for" & " " & clsOrders.CPT_Desc, vbInformation
                    mbooNoSave = True
                    Exit Function
                End If
                If clsOrders.CPT_Desc <> vbNullString Then
                    Write_LabOrder clsOrders.Type, clsOrders.CPT_Code, clsOrders.CPT_Desc, clsOrders.Modifier, clsOrders.Comment, clsOrders.Status, clsOrders.ICD_Desc, clsOrders.ICD_Code
                    Write_Order clsOrders.Type, clsOrders.CPT_Code, clsOrders.CPT_Desc, clsOrders.Modifier, clsOrders.Comment, clsOrders.Status, clsOrders.ICD_Desc, clsOrders.ICD_Code
                End If
            Else
                If clsOrders.CPT_Desc <> vbNullString Then
                Write_Order clsOrders.Type, clsOrders.CPT_Code, clsOrders.CPT_Desc, clsOrders.Modifier, clsOrders.Comment, clsOrders.Status, clsOrders.ICD_Desc, clsOrders.ICD_Code
                End If
           End If
        Next lngRow
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, clsOrders.CPT_Code & "~" & clsOrders.CPT_Desc, "Write_Data", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cboAttendPhy_Change()
    Clear_Signature
    Exit Sub
End Sub

Private Sub cctSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(cctSearch.Text) > 2 Then Load_GridBox "ICD", "ICD"
End Sub

Private Sub cmdAddOrder_Click()
Dim varCellValue
Dim strList As String
Dim lngRow As Long

    If mlngHeadRow > 5 Then
        Remove_Signature
    End If
    lstICD.Clear
     Add_Order
    Exit Sub
End Sub

Public Sub Read_LabOrder()
Dim strSQL As String
Dim i%
Dim R%
Dim vName
Dim vAddress
Dim vCityStateZip
Dim lngRow As Long
Dim Arr
Dim strList As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset

    On Error GoTo Error_Read_LabOrder
    Screen.MousePointer = vbHourglass
    lngRow = grdSS.MaxRows + 1
    If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    strSQL = "Select DISTINCT a.LAB_ID,a.cpt_desc, a.cpt_intrvn_cd, a.cpt_mod, a.comment," _
        & "a.lab_priority, b.ICD_CODE, b.ICD_DESC  From LAB_HEADER a, LAB_DIAG b " _
        & "Where a.PATIENT_ID = " & Val(udtPatientName.IntId) & " And " _
        & "a.ENCOUNTER_ID = " & udtEnctr.EnctrID & " AND a.LAB_ID = b.LAB_HEADER_ID AND a.ENCOUNTER_ID = b.ENCOUNTER_ID AND " _
        & "b.PATIENT_ID = a.PATIENT_ID"
     Set rsLab = cnLab.Execute(strSQL)
     With rsLab
        Do Until .EOF
            grdTerms.ReDraw = True
            lngRow = lngRow + 1
            grdSS.MaxRows = lngRow
            grdSS.SetRowItemData (lngRow), (!Lab_Id)
            If Trim(mstrSubGroup) = vbNullString Then Get_CPTGroups
            Format_CellTypeCombo grdSS, lngRow, 1, "1", mstrSubGroup, "TypeHAlignCenter"
            Set_GridCurSel grdSS, "Lab", 1, lngRow, "1", "ORDER_TYPE"
            If Trim((!CPT_Desc)) <> vbNullString And (!Cpt_Intrvn_Cd) <> vbNullString Then
                Format_CellTypeEdit grdSS, lngRow, 2, "1", Trim((!CPT_Desc)) & " - " & Trim((!Cpt_Intrvn_Cd)), "TypeHAlignCenter", "CPT_CODE"
            End If
            If Trim((!CPT_MOD)) <> vbNullString Then
                Format_CellTypeEdit grdSS, lngRow, 3, "1", Trim((!CPT_MOD)), "TypeHAlignCenter", "MODIFIER"
            End If
            If Trim((!ICD_Desc)) <> vbNullString And (!ICD_Code) <> vbNullString Then
                Format_CellTypeEdit grdSS, lngRow, 4, "1", (!ICD_Desc) & " - " & (!ICD_Code), "TypeHAlignCenter", "ICD_CODE"
            End If
            strList = "Routine" & Chr(9) & "Stat"
            Format_CellTypeCombo grdSS, lngRow, 5, "1", strList, "TypeHAlignCenter", "STATUS"
            If (!Lab_Priority) <> vbNullString Then Set_GridCurSel grdSS, (!Lab_Priority), 5, lngRow, "1", "STATUS"
            If (!Comment) <> vbNullString Then
                Format_CellTypeEdit grdSS, lngRow, 6, "2", Trim((!Comment)), "TypeHAlignCenter", "SPECIAL_INSTRUCTIONS"
            End If
            grdSS.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdSS.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdSS.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdSS.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdSS.SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdSS.SetCellBorder 6, lngRow, 7, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .MoveNext
        Loop
        grdTerms.ReDraw = True
    End With
    
    If cnLab.State = adStateOpen Then
        cnLab.Close
        Set cnLab = Nothing
        Set rsLab = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Read_LabOrder:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_LabOrder", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub Write_SuperBill(vstrType As String, vstrCPT As String, vstrCPTDesc As String, vstrMod As String, vstrComments As String, _
vstrPrior As String, vstrICDdesc As String, vstrICD As String)
Dim cnSuperBill As New ADODB.Connection
Dim rsSuperBill As New ADODB.Recordset
Dim iReturn As Integer
Dim strFee As String
Dim lngPrimKey As Long
Dim sDate As String
Dim strSQL As String
Dim ArrCurrent
Dim ArrDB
Dim strStatus As String
Dim lngBillID As Long
Dim Idx%
Dim ArrCPT
Dim strDesc As String
Dim booinsert As Boolean

    On Error GoTo ERROR_Write_SuperBill

    If cnSuperBill.State = adStateOpen Then cnSuperBill.Close
    If cnSuperBill.State = adStateClosed Then cnSuperBill.open psConnect(0)
    If piFacilityID = 0 Then
        strSQL = "Select FACILITY_ID FROM FACILITIES WHERE Facility_Name = '" & udtEnctr.Facility & "'"
        Set rsSuperBill = cnSuperBill.Execute(strSQL)
        If Not rsSuperBill.EOF Then
           If rsSuperBill.Fields("FACILITY_ID") > 0 Then piFacilityID = rsSuperBill.Fields("FACILITY_ID")
        End If
    End If
    booinsert = True
    clsPatBill.AcctType = "Patient"
    clsPatAcct.Acct_ID = 0
    strSQL = "select ACCT_ID from PATIENT_ACCOUNTS where PATIENT_ID = " & Val(udtPatientName.IntId) & " AND " _
        & "PATIENT_TYPE = 'Patient'"
    Set rsSuperBill = cnSuperBill.Execute(strSQL)
    booinsert = True
    With rsSuperBill
        If Not .EOF Then
            clsPatAcct.Acct_ID = (!Acct_ID)
        End If
    End With
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If clsPatAcct.Acct_ID = 0 Then
        strSQL = "INSERT INTO PATIENT_ACCOUNTS (PATIENT_ID, FACILITY_ID, PROVIDER_ID," _
            & "PATIENT_TYPE, PATIENT_STATUS,  POS, COMMENTS,  ACCOUNT_NO, INSERT_DATE) " _
            & "values (" & Val(udtPatientName.IntId) & "," _
            & piFacilityID & "," _
            & piProviderID & ",'Patient','" _
            & udtEnctr.VisitType & "','" _
            & udtEnctr.ServicePlaceCode & "','" _
            & clsPatBill.BillNote & "'," _
            & Val(udtPatientName.IntId) & ",'" _
            & sDate & "')"
        Set rsSuperBill = cnSuperBill.Execute(strSQL)
    End If
    If clsPatAcct.Acct_ID = 0 Then
        strSQL = "select ACCT_ID from PATIENT_ACCOUNTS where PATIENT_ID = " & Val(udtPatientName.IntId) & " AND " _
        & "PATIENT_TYPE = 'Patient'"
        Set rsSuperBill = cnSuperBill.Execute(strSQL)
        booinsert = True
        With rsSuperBill
            If Not .EOF Then
                If Not IsNull((!Acct_ID)) Then
                    clsPatAcct.Acct_ID = (!Acct_ID)
                End If
            End If
        End With
    End If
    If cnSuperBill.State = adStateOpen Then cnSuperBill.Close
    ArrCPT = Split(vstrCPTDesc, "|")
    If UBound(ArrCPT) > 0 Then
        For Idx = LBound(ArrCPT) To UBound(ArrCPT)
            strDesc = ArrCPT(Idx)
            If strDesc <> vbNullString Then
                If cnSuperBill.State = adStateClosed Then cnSuperBill.open psConnect(1)
                strSQL = "SELECT CPT_FEE From CPT_TEMPLATE WHERE CPT_NAME = '" & strDesc & "'"
                Set rsSuperBill = cnSuperBill.Execute(strSQL)
                With rsSuperBill
                    If Not .EOF Then
                        If strFee = vbNullString Then
                            If (!CPT_Fee) <> vbNullString Then strFee = (!CPT_Fee)
                        Else
                            If (!CPT_Fee) <> vbNullString Then strFee = strFee & "|" & (!CPT_Fee)
                        End If
                    End If
                End With
                If cnSuperBill.State = adStateOpen Then
                    cnSuperBill.Close
                    Set cnSuperBill = Nothing
                End If
            End If
        Next Idx
    Else
        strDesc = ArrCPT(0)
             If strDesc <> vbNullString Then
              If cnSuperBill.State = adStateClosed Then cnSuperBill.open psConnect(1)
                strSQL = "SELECT CPT_FEE From CPT_TEMPLATE WHERE CPT_NAME = '" & strDesc & "'"
                Set rsSuperBill = cnSuperBill.Execute(strSQL)
                With rsSuperBill
                    If Not .EOF Then
                        If (!CPT_Fee) <> vbNullString Then strFee = (!CPT_Fee)
                    End If
                End With
                If cnSuperBill.State = adStateOpen Then
                    cnSuperBill.Close
                    Set cnSuperBill = Nothing
                End If
            End If
    End If
    If cnSuperBill.State = adStateOpen Then cnSuperBill.Close
    If cnSuperBill.State = adStateClosed Then cnSuperBill.open psConnect(0)
    strSQL = "Select * From SUPERBILL " _
       & "Where patient_id  = " & Val(udtPatientName.IntId) & " And " _
       & "encounter_id  =  " & udtEnctr.EnctrID
        Set rsSuperBill = cnSuperBill.Execute(strSQL)
        If Not rsSuperBill.EOF Then
            If rsSuperBill.Fields("BILLING_ID") > 0 Then
                strStatus = "Update"
                lngPrimKey = rsSuperBill.Fields("BILLING_ID")
                'CPT_Desc
                If rsSuperBill.Fields("CPT_DESC") <> vbNullString Then
                   ArrDB = Split(rsSuperBill.Fields("CPT_DESC"), "|")
                   If UBound(ArrDB) > 0 Then
                       For Idx = LBound(ArrDB) To UBound(ArrDB)
                           If Not TextFound(vstrCPTDesc, ArrDB(Idx), Len(vstrCPTDesc), True) Then
                               vstrCPTDesc = ArrDB(Idx) & "|" & vstrCPTDesc
                           End If
                       Next Idx
                   Else
                       If Not TextFound(vstrCPTDesc, rsSuperBill.Fields("CPT_DESC"), Len(vstrCPTDesc), True) Then
                           vstrCPTDesc = rsSuperBill.Fields("CPT_DESC") & "|" & vstrCPTDesc
                       End If
                   End If
                End If
            End If
            'Code
            If rsSuperBill.Fields("CPT_CODE") <> vbNullString Then
               ArrDB = Split(rsSuperBill.Fields("CPT_CODE"), "|")
               If UBound(ArrDB) > 0 Then
                   For Idx = LBound(ArrDB) To UBound(ArrDB)
                       If Not TextFound(vstrCPT, ArrDB(Idx), Len(vstrCPT), True) Then
                           vstrCPT = ArrDB(Idx) & "|" & vstrCPT
                       End If
                   Next Idx
               Else
                   If Not TextFound(vstrCPT, rsSuperBill.Fields("CPT_CODE"), Len(vstrCPT), True) Then
                       vstrCPT = rsSuperBill.Fields("CPT_CODE") & "|" & vstrCPT
                   End If
               End If
            End If
            If rsSuperBill.Fields("FEE") <> vbNullString Then
               ArrDB = Split(rsSuperBill.Fields("FEE"), "|")
               If UBound(ArrDB) > 0 Then
                   For Idx = LBound(ArrDB) To UBound(ArrDB)
                       If Not TextFound(strFee, ArrDB(Idx), Len(strFee), True) Then
                           strFee = ArrDB(Idx) & "|" & strFee
                       End If
                   Next Idx
               Else
                   If Not TextFound(strFee, rsSuperBill.Fields("FEE"), Len(strFee), True) Then
                       strFee = rsSuperBill.Fields("FEE") & "|" & strFee
                   End If
               End If
            End If
        End If
    lngPrimKey = 0
    sDate = Format(Now, "mm/dd/yyyy hh:nn:ss AM/PM")
    If cnSuperBill.State = adStateClosed Then cnSuperBill.open psConnect(0)
    strSQL = "SELECT BILLING_ID From SUPERBILL " _
            & "WHERE PATIENT_ID  = " & Val(udtPatientName.IntId) & " And " _
            & "ENCOUNTER_ID  = " & udtEnctr.EnctrID
    Set rsSuperBill = cnSuperBill.Execute(strSQL)
    With rsSuperBill
        If Not .EOF Then
            If (!BILLING_ID) > 0 Then lngPrimKey = (!BILLING_ID)
        End If
    End With
 
    If lngPrimKey <= 0 Then
        strSQL = "Insert into SUPERBILL (Acct_ID, FACILITY_ID, patient_id, Encounter_id, POS, cpt_desc, cpt_code, insert_date, status, fee, CPT_GROUP) " _
            & "values (" & clsPatAcct.Acct_ID & "," & piFacilityID & "," & Val(udtPatientName.IntId) & "," _
                        & udtEnctr.EnctrID & "," _
                        & Val(udtEnctr.ServicePlaceCode) & ",'" _
                        & vstrCPTDesc & "','" _
                        & vstrCPT & "','" _
                        & sDate & "','" _
                        & vstrPrior & "','" _
                        & strFee & "','" _
                        & vstrType & "' )"
        Set rsSuperBill = cnSuperBill.Execute(strSQL)
    Else
        strSQL = "Update SUPERBILL " _
            & "Set cpt_desc = '" & vstrCPTDesc & "', cpt_code =  '" & vstrCPT & "'," _
             & "Status = '" & vstrPrior & "',POS = " & Val(udtEnctr.ServicePlaceCode) & ", CPT_GROUP = '" & vstrType & "'," _
            & "FEE = '" & strFee & "', UPDATE_DATE = '" & sDate & "', FACILITY_ID = " & piFacilityID & " WHERE " _
            & "BILLING_ID = " & lngPrimKey
        Set rsSuperBill = cnSuperBill.Execute(strSQL)
    End If
    
    If cnSuperBill.State = adStateOpen Then
        cnSuperBill.Close
        Set cnSuperBill = Nothing
        Set rsSuperBill = Nothing
    End If
    Exit Sub
ERROR_Write_SuperBill:
    Write_Error Err.Description, Err.Number, strSQL, "Write_SuperBill", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Public Function Read_Order()
Dim strSQL As String
Dim lngRow As Long
Dim Arr
Dim strList As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngOrderID As Long

    On Error GoTo Error_Read_Order
    Screen.MousePointer = vbHourglass
    If Val(udtPatientName.IntId) <= 0 Then
        Exit Function
    End If
    grdTerms.ReDraw = False
    lngRow = grdSS.MaxRows
    If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    strSQL = "Select * FROM PATIENT_ORDERS " _
        & "Where PATIENT_ID = " & Val(udtPatientName.IntId) & " And " _
        & "ENCOUNTER_ID = " & udtEnctr.EnctrID
    Set rsLab = cnLab.Execute(strSQL)
    With rsLab
        Do Until .EOF
            Clear_LabString
            lngOrderID = 0
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "ORDER_ID"
                               lngOrderID = Fld.Value
                            Case "CPT_PROCEDURE"
                               udtLaboratory.CPT_Proc = Fld.Value
                            Case "CPT_DESC"
                               udtLaboratory.CPT_Desc = Fld.Value
                            Case "CPT_CODE"
                               udtLaboratory.CPT_Code = Fld.Value
                            Case "CPT_MOD"
                               udtLaboratory.CPT_MOD = Fld.Value
                            Case "DIAG_DESC"
                               udtLaboratory.Diag_Desc = Fld.Value
                            Case "DIAG_CODE"
                               udtLaboratory.Diag_Code = Fld.Value
                            Case "SPECIAL_INSTRUCTIONS"
                               udtLaboratory.Spec_Instruct = Fld.Value
                            Case "PRIORITY"
                               udtLaboratory.Lab_Status = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
             If lngOrderID > 0 Then
                lngRow = lngRow + 1
                grdSS.MaxRows = lngRow
                grdSS.SetRowItemData lngRow, lngOrderID
                If Trim(mstrSubGroup) = vbNullString Then Get_CPTGroups
                Format_CellTypeCombo grdSS, lngRow, 1, "1", mstrSubGroup, "TypeHAlignCenter", "ORDER_TYPE"
                Set_GridCurSel grdSS, "Lab", 1, lngRow, "1", Trim(udtLaboratory.CPT_Proc)
                Format_CellTypeEdit grdSS, lngRow, 2, "1", udtLaboratory.CPT_Desc & " - " & udtLaboratory.CPT_Code, "TypeHAlignLeft", "CPT_CODE"
                Format_CellTypeEdit grdSS, lngRow, 3, "1", udtLaboratory.CPT_MOD, "TypeHAlignLeft", "MODIFIER"
                Format_CellTypeEdit grdSS, lngRow, 4, "1", udtLaboratory.Diag_Desc & " - " & udtLaboratory.Diag_Code, "TypeHAlignLeft", "ICD_CODE"
                strList = "Routine" & Chr(9) & "Stat"
                Format_CellTypeCombo grdSS, lngRow, 5, "1", strList, "TypeHAlignCenter", "STATUS"
                If Trim(udtLaboratory.Lab_Status) <> vbNullString Then Set_GridCurSel grdSS, udtLaboratory.Lab_Status, 5, lngRow, "1", "STATUS"
                Format_CellTypeEdit grdSS, lngRow, 6, "2", udtLaboratory.Spec_Instruct, "TypeHAlignLeft", "SPECIAL_INSTRUCTIONS"
                grdSS.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdSS.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdSS.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdSS.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdSS.SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdSS.SetCellBorder 6, lngRow, 7, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        Loop
    End With
    If lngOrderID > 0 Then
        InitSignature
        Add_Signature
    End If

    grdTerms.ReDraw = True
    If cnLab.State = adStateOpen Then
        cnLab.Close
        Set cnLab = Nothing
        Set rsLab = Nothing
        Set Flds = Nothing
    End If
    Clear_LabString
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Order:
    Write_Error Err.Description, Err.Number, strSQL, "Read_Order", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Write_Order(vstrType As String, vstrCPT As String, vstrCPTDesc As String, vstrMod As String, vstrComments As String, _
vstrPrior As String, vstrICDdesc As String, vstrICD As String)
Dim strSQL As String
Dim booFound As Boolean
Dim intSeq As Integer
Dim strGroup As String
Dim strSubGroup As String
Dim lngPrimKey As Long
Dim sDate As String
Dim lngRef As Long
Dim booinsert As Boolean
Dim strStatus As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset

    On Error GoTo Error_Write_Order
    Screen.MousePointer = vbHourglass
    
   If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    If udtPatientName.IntId <= 0 Then
        MsgBox "Please select patient prior to saving order!", vbInformation, "PatientTrac Orders"
        Exit Function
    End If
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    lngRef = 0
    'Insert or Update?......................
    strSQL = "Select ORDER_ID  From PATIENT_ORDERS " _
            & "Where patient_id = " & Val(udtPatientName.IntId) & " And " _
            & "encounter_id = " & udtEnctr.EnctrID & " And " _
            & "CPT_CODE = '" & vstrCPT & "'"
    Set rsLab = cnLab.Execute(strSQL)
     With rsLab
        If Not .EOF Then
            If (!ORDER_ID) > 0 Then lngPrimKey = (!ORDER_ID)
        End If
    End With
    If vstrPrior <> "Administered" Then
        strStatus = "Ordered"
    Else
        strStatus = Trim(vstrPrior)
    End If
    If lngPrimKey <= 0 Then
        strSQL = "Insert into PATIENT_ORDERS (patient_id, encounter_id, provider_id, CPT_PROCEDURE, CPT_CODE, " _
                & "CPT_DESC, CPT_MOD, DIAG_CODE,DIAG_DESC, SPECIAL_INSTRUCTIONS," _
                & "STATUS, PRIORITY, insert_date) " _
                & "Values(" & Val(udtPatientName.IntId) & "," _
                        & udtEnctr.EnctrID & "," _
                        & piProviderID & ",'" _
                        & vstrType & "','" _
                        & vstrCPT & "','" _
                        & vstrCPTDesc & "','" _
                        & vstrMod & "','" _
                        & vstrICD & "','" _
                        & vstrICDdesc & "','" _
                        & vstrComments & "','" & strStatus & "','" _
                        & vstrPrior & "','" _
                        & sDate & "')"
        Set rsLab = cnLab.Execute(strSQL)
    ElseIf lngPrimKey > 0 Then
        strSQL = "Update PATIENT_ORDERS " _
            & "Set cpt_desc = '" & vstrCPTDesc & "'," _
            & "SPECIAL_INSTRUCTIONS = '" & vstrComments & "'," _
            & "CPT_PROCEDURE =  '" & vstrType & "'," _
            & "CPT_CODE =  '" & vstrCPT & "'," _
            & "DIAG_CODE =  '" & vstrICD & "'," _
            & "DIAG_DESC =  '" & vstrICDdesc & "'," _
            & "STATUS = '" & strStatus & "'," _
            & "PRIORITY = '" & vstrPrior & "'," _
            & "CPT_MOD = '" & vstrMod & "'," _
            & "UPDATE_DATE = '" & sDate & "' " _
            & "WHERE patient_id = " & Val(udtPatientName.IntId) & " And " _
            & "encounter_id = " & udtEnctr.EnctrID & " And " _
            & "ORDER_ID = " & lngPrimKey
        Set rsLab = cnLab.Execute(strSQL)
    End If
    mbooNoSave = False
    If cnLab.State = adStateOpen Then
        cnLab.Close
        Set cnLab = Nothing
        Set rsLab = Nothing
    End If
    If vstrPrior = "Administered" Then Write_SuperBill vstrType, vstrCPT, vstrCPTDesc, vstrMod, vstrComments, vstrPrior, vstrICDdesc, vstrICD
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Order:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Order", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub Write_LabOrder(vstrType As String, vstrCPT As String, vstrCPTDesc As String, vstrMod As String, vstrComments As String, _
vstrPrior As String, vstrICDdesc As String, vstrICD As String)
Dim strSQL As String
Dim booFound As Boolean
Dim intSeq As Integer
Dim R%
Dim strGroup As String
Dim strSubGroup As String
Dim lngPrimKey As Long
Dim sDate As String
Dim booinsert As Boolean
Dim lngLabID As Long
Dim lngHeaderID As Long
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset

    On Error GoTo Error_Write_LabOrder
    Screen.MousePointer = vbHourglass
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If piProviderID = 0 Then piProviderID = modSpecific.piProviderID
    If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    If vstrICD <> vbNullString Then vstrICD = Format(vstrICD, "000.00")
    booinsert = True
    strSQL = "Delete from LAB_RESULTS Where patient_id = " & Val(udtPatientName.IntId) & " And encounter_id = " & udtEnctr.EnctrID
    Set rsLab = cnLab.Execute(strSQL)
    
    'Insert or Update?......................
    strSQL = "Select LAB_ID  From LAB_HEADER " _
            & "Where patient_id = " & Val(udtPatientName.IntId) & " And " _
            & "encounter_id = " & udtEnctr.EnctrID & " And " _
            & "cpt_intrvn_cd = '" & vstrCPT & "' "
    Set rsLab = cnLab.Execute(strSQL)
    With rsLab
        If Not .EOF Then
            booinsert = False
            If (!Lab_Id) > 0 Then lngPrimKey = (!Lab_Id)
        End If
    End With
    PsUserID = Get_PsUserID(psConnect(0))
    If lngPrimKey <= 0 Then
        strSQL = "Insert into LAB_HEADER (patient_id, encounter_id, provider_id, cpt_intrvn_cd, " _
                & "cpt_desc, cpt_mod, COMMENT," _
                & "lab_status, lab_priority, insert_user, insert_date) " _
                & "Values(" & Val(udtPatientName.IntId) & "," _
                        & udtEnctr.EnctrID & "," _
                        & piProviderID & ",'" _
                        & vstrCPT & "','" _
                        & vstrCPTDesc & "','" _
                        & vstrMod & "','" _
                        & vstrComments & "','Ordered','" _
                        & vstrPrior & "','" _
                        & PsUserID & "','" _
                        & Date & "')"
    Else  'Update
        strSQL = "Update LAB_HEADER " _
            & "Set cpt_desc = '" & vstrCPTDesc & "'," _
            & "COMMENT = '" & vstrComments & "'," _
            & "cpt_intrvn_cd =  '" & vstrCPT & "'," _
            & "lab_status = 'Ordered'," _
            & "lab_priority = '" & vstrPrior & "'," _
            & "insert_user = '" & PsUserID & "'," _
            & "insert_date = '" & sDate & "' " _
            & "Where patient_id = " & Val(udtPatientName.IntId) & " And " _
            & "encounter_id = " & udtEnctr.EnctrID & " And " _
            & "LAB_ID =  " & lngPrimKey
    End If
    Set rsLab = cnLab.Execute(strSQL)

    'Get primary key......................
    strSQL = "Select LAB_ID  From LAB_HEADER " _
            & "Where patient_id = " & Val(udtPatientName.IntId) & " And " _
            & "encounter_id = " & udtEnctr.EnctrID & " And " _
            & "cpt_intrvn_cd = '" & vstrCPT & "' "
    Set rsLab = cnLab.Execute(strSQL)
    With rsLab
        If Not .EOF Then
            If (!Lab_Id) > 0 Then lngHeaderID = (!Lab_Id)
        End If
    End With
    intSeq = 0
    lngLabID = 0
    intSeq = intSeq + 1
    If udtEnctr.EnctrID > 0 And Val(udtPatientName.IntId) > 0 Then
        'Need to validate
        strSQL = "SELECT LAB_ID FROM LAB_DIAG WHERE PATIENT_ID = " & Val(udtPatientName.IntId) & " AND ENCOUNTER_ID = " & udtEnctr.EnctrID & " AND " _
            & "CPT_CODE = '" & vstrCPT & "'"
        Set rsLab = cnLab.Execute(strSQL)
        If Not rsLab.EOF Then
            lngLabID = rsLab.Fields("LAB_ID")
        End If
        If lngLabID <= 0 Then
            strSQL = "INSERT INTO LAB_DIAG ( LAB_HEADER_ID,PATIENT_ID, ENCOUNTER_ID,CPT_CODE,ICD_CODE,ICD_DESC, SEQUENCE_NO) " _
                & "values (" & lngHeaderID & "," & Val(udtPatientName.IntId) & "," & udtEnctr.EnctrID & ",'" & vstrCPT & "','" _
                & vstrICD & "','" & vstrICDdesc & "'," & intSeq & ")"
             Set rsLab = cnLab.Execute(strSQL)
        Else
            strSQL = "UPDATE LAB_DIAG  SET LAB_HEADER_ID =  " & lngHeaderID & ", CPT_CODE = '" & vstrCPT & "',ICD_CODE = '" & vstrICD & "', ICD_DESC = '" & vstrICDdesc & "'," _
                & "SEQUENCE_NO = " & intSeq & " WHERE LAB_ID = " & lngLabID
        Set rsLab = cnLab.Execute(strSQL)
        End If
    End If
    mbooNoSave = False
    If cnLab.State = adStateOpen Then
        cnLab.Close
        Set cnLab = Nothing
        Set rsLab = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Write_LabOrder:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_LabOrder", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub cmdALL_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer
Dim CellNote As Long
Dim intPrint As Integer
Dim strRef As String
Dim lngRow As Long

    On Error GoTo Error_Print_All
    mbooCancelPrint = False
    Write_Data
    If clsProvider.Name = vbNullString Then
        modSpecific.piProviderID = piProviderID
        ModPatientData.Get_Providers
    End If
     Remove_Signature
     Add_Signature
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation
        Exit Sub
    End If
    With grdSS
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintJobName = "Laboratory Order" & "_" & Val(udtPatientName.IntId)
        .PrintColHeaders = True
        .PrintRowHeaders = False
        .PrintScalingMethod = PrintScalingMethodSmartPrint
        .PrintCenterOnPageH = True
        '.PrintCenterOnPageV = True
        .PrintMarginBottom = 770
        .PrintBorder = True
        .PrintColor = True
        .PrintOrientation = PrintOrientationPortrait
        .PrintShadows = False
        .PrintSheet
        mbooPrintPreview = False
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Print_All:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "Printer", "Print_All", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdClose_Click()
    ccaSignature.ClearSignature
    ccaSignature.Enabled = False
    ccaSignature.ShowSignedDate = False
    Unload Me
    Exit Sub
End Sub

Private Sub cmdCloseList_Click()
    fraRecICD.Visible = False
    lstICD.Clear
    grdSS_DblClick 5, mlngGridRow
     Exit Sub
End Sub

Private Sub cmdDefault_Click()
    frmDefaultOrder.Show vbModal
    Exit Sub
End Sub

Private Sub cmdDone_Click()
Dim strData As String
Dim lngRow As Long
Dim varCellValue
Dim varCode
Dim Arr
Dim ArrCode
Dim lngGridRow As Long
Dim lngCnt As Long
Dim strTag As String
Dim strList As String
Dim lngLabRow As Long

    lngCnt = -1
    With grdTerms
        .ReDraw = True
        fraGridTerm.Visible = False
        .Col = 1
        .Row = SpreadHeader
        If .CellTag <> vbNullString Then strTag = .CellTag
        Select Case strTag
            Case "CPT DESC"
                For lngRow = 1 To .MaxRows
                    strData = vbNullString
                    .Col = 1
                    .Row = lngRow
                    If .BackColor = &HFFC0FF Then
                        .GetText 1, lngRow, varCellValue
                        .GetText 2, lngRow, varCode
                        If Trim(varCellValue) <> vbNullString Then
                            lngCnt = lngCnt + 1
                            strData = Trim(varCellValue) & " - " & Trim(varCode)
                            If lngCnt = 0 Then
                                Format_CellTypeEdit grdSS, mlngGridRow, 2, "1", strData, "TypeHAlignLeft", "ORDER_TYPE"
                                grdSS.SetCellBorder 2, mlngGridRow, 2, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleFineSolid
                            Else
                                Add_Order
                                lngLabRow = grdSS.MaxRows
                                Set_GridCurSel grdSS, "Lab", 1, lngLabRow, "1", "ORDER_TYPE"
                                Format_CellTypeEdit grdSS, lngLabRow, 2, "1", strData, "TypeHAlignLeft", "CPT_CODE"
                                grdSS.SetCellBorder 1, lngLabRow, 1, lngLabRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                grdSS.SetCellBorder 2, lngLabRow, 2, lngLabRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        End If
                    End If
                Next lngRow
            Case "ICD_DESC", "ICD_CODE"
                 For lngRow = 1 To grdTerms.MaxRows
                    .Col = 1
                    .Row = lngRow
                    If .BackColor = &HFFC0FF Then
                        .GetText 1, lngRow, varCellValue
                        .GetText 2, lngRow, varCode
                        If Trim(varCellValue) <> vbNullString Then
                            If strData = vbNullString Then
                                strData = Trim(varCellValue) & " - " & Trim(varCode)
                            Else
                                strData = strData & vbNewLine & Trim(varCellValue) & " - " & Trim(varCode)
                            End If
                        End If
                     End If
                Next lngRow
                lngRow = 0
                For lngRow = mlngGridRow To grdSS.MaxRows
                    grdSS.GetText 1, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString And Trim(varCellValue) <> vbNullString Then
                        Format_CellTypeEdit grdSS, lngRow, 4, "1", strData, "TypeHAlignLeft", "ICD_CODE"
                        grdSS.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleFineSolid
                    End If
                Next lngRow
            Case "MODIFIER"
                 For lngRow = 1 To .MaxRows
                    .Col = 1
                    .Row = lngRow
                    If .BackColor = &HFFC0FF Then
                        .GetText 1, lngRow, varCellValue
                        .GetText 2, lngRow, varCode
                        If Trim(varCellValue) <> vbNullString Then
                            If strData = vbNullString Then
                                strData = Trim(varCellValue) & " - " & Trim(varCode)
                            Else
                                strData = strData & vbNewLine & Trim(varCellValue) & " - " & Trim(varCode)
                            End If
                        End If
                     End If
                Next lngRow
                If Trim(strData) <> vbNullString Then
                    Format_CellTypeEdit grdSS, mlngGridRow, 3, "1", strData, "TypeHAlignLeft"
                    grdSS.SetCellBorder 3, mlngGridRow, 3, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleFineSolid
                End If
            End Select
        grdTerms.ClearRange 1, 1, grdTerms.MaxCols, grdTerms.MaxRows, False
        .ReDraw = True
    End With
    Exit Sub
End Sub

Private Sub cmdErase_Click()
    Clear_Signature
    Exit Sub
End Sub
Private Function Add_Order()
Dim varCellValue
Dim strList As String
Dim lngRow As Long

    With grdSS
        .ReDraw = True
        lngRow = grdSS.MaxRows + 1
        grdSS.MaxRows = lngRow
        .Col = 0
        Format_CellTypeEdit grdSS, lngRow, 0, "1", "Superbill", "TypeHAlignCenter"
        grdSS.ColHidden = True
        If Trim(mstrSubGroup) = vbNullString Then Get_CPTGroups
        Format_CellTypeCombo grdSS, lngRow, 1, "1", mstrSubGroup, "TypeHAlignCenter", "ORDER_TYPE"
        Format_CellTypeEdit grdSS, lngRow, 2, "1", "", "TypeHAlignCenter", "CPT_CODE"
        Format_CellTypeEdit grdSS, lngRow, 3, "1", "", "TypeHAlignCenter", "MODIFIER"
        Format_CellTypeEdit grdSS, lngRow, 4, "1", "", "TypeHAlignCenter", "ICD_CODE"
        strList = "Routine" & Chr(9) & "Stat"
        Format_CellTypeCombo grdSS, lngRow, 5, "1", strList, "TypeHAlignCenter", "STATUS"
        Format_CellTypeEdit grdSS, lngRow, 6, "2", "", "TypeHAlignCenter", "", "SPECIAL_INSTRUCTIONS"
        .SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 6, lngRow, 7, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .ReDraw = False
    End With
    Exit Function
End Function

Private Sub cmdEstablish_Click()
Dim strSQL As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim Arr
Dim lngRow As Long
Dim lngCol As Long
Dim strProvider As String
Dim i%

    On Error GoTo Error_Read_New_Patient_Order
    Screen.MousePointer = vbHourglass
    Get_Providers piProviderID
    strProvider = clsProvider.Name
    If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    strSQL = "SELECT * FROM DEFAULT_ORDERS WHERE PROVIDER_ID = " & piProviderID & " AND CONDITION = 'Established Patient'"
    Set rsLab = cnLab.Execute(strSQL)
    With rsLab
        Do Until .EOF
            Clear_LabString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PROC_ID"
                                 udtLaboratory.Lab_Id = Fld.Value
                            Case "PROC_GROUP"
                                udtLaboratory.OrderType = Fld.Value
                            Case "AGE_VALUE"
                                udtLaboratory.AgeRange = Fld.Value
                            Case "AGE"
                                udtLaboratory.Age = Fld.Value
                            Case "GENDER"
                                udtLaboratory.Gender = Fld.Value
                            Case "CONDITION"
                                 udtLaboratory.Med_Condition = Fld.Value
                            Case "LAB_CODE"
                                udtLaboratory.CPT_Code = Fld.Value
                            Case "LAB_DESC"
                                udtLaboratory.CPT_Desc = Fld.Value
                            Case "ICD_CODE"
                                udtLaboratory.Diag_Code = Fld.Value
                            Case "ICD_DESC"
                                udtLaboratory.Diag_Desc = Fld.Value
                            Case "PROVIDER_ID"
                                udtLaboratory.ProviderID = Fld.Value
                            Case "CONSULT"
                                udtLaboratory.PatientNotice = Fld.Value
                            Case "INTRVN_PRIORITY"
                                udtLaboratory.Lab_Priority = Fld.Value
                            Case "COMMENTS"
                                udtLaboratory.Spec_Instruct = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtLaboratory.Lab_Id > 0 Then
                Add_Order
                lngRow = grdSS.MaxRows
                For lngCol = 1 To grdSS.MaxCols
                    grdSS.Col = lngCol
                    grdSS.Row = lngRow
                    If grdSS.CellTag <> vbNullString Then
                        Select Case UCase(grdSS.CellTag)
                            Case "ORDER_TYPE"
                                Set_GridCurSel grdSS, udtLaboratory.OrderType, lngCol, lngRow, "1", "ORDER_TYPE"
                            Case "SPECIAL_INSTRUCTIONS"
                                Format_CellTypeEdit grdSS, lngRow, lngCol, "1", udtLaboratory.Spec_Instruct, "TypeHAlignCenter", "SPECIAL_INSTRUCTIONS"
                            Case "CPT_CODE"
                                Format_CellTypeEdit grdSS, lngRow, lngCol, "1", udtLaboratory.CPT_Desc & " - " & udtLaboratory.CPT_Code, "TypeHAlignCenter", "CPT_CODE"
                            Case "ICD_CODE"
                                Format_CellTypeEdit grdSS, lngRow, lngCol, "1", udtLaboratory.Diag_Desc & " - " & udtLaboratory.Diag_Code, "TypeHAlignCenter", "ICD_CODE"
                            Case "STATUS"
                                Set_GridCurSel grdSS, udtLaboratory.Lab_Priority, lngCol, lngRow, "1", "STATUS"
                        End Select
                    End If
                Next lngCol
            End If
           Loop
        End With
        grdSS.Col = 1
        grdSS.Row = 1
    If cnLab.State = adStateOpen Then
        Set cnLab = Nothing
        Set rsLab = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Read_New_Patient_Order:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, " Read_New_Patient_Order", App.EXEName & "." & Me.Name
    Exit Sub
    Resume

End Sub

Private Sub cmdGridTerms_Click()
    frmCPTMaint.Show vbModal
    Exit Sub
End Sub

Private Sub cmdLAB_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer
Dim CellNote As Long
Dim intPrint As Integer
Dim strRef As String
Dim lngRow As Long
Dim varCellValue

    On Error GoTo Error_Print_Lab
    Screen.MousePointer = vbHourglass
    With grdSS
        .ReDraw = False
        Write_Data
        For lngRow = 6 To .MaxRows
            .Row = lngRow
            .GetText 1, lngRow, varCellValue
            If Not TextFound("Lab", Trim(varCellValue), Len("Lab"), True) Then .RowHidden = True
        Next lngRow
    End With
    If clsProvider.Name = vbNullString Then
        modSpecific.piProviderID = piProviderID
        ModPatientData.Get_Providers
    End If
     Remove_Signature
     Add_Signature
     Printer.TrackDefault = True
     CD_ShowPrinter (Me.hwnd)
     If mbooCancelPrint Then
        For lngRow = 6 To grdSS.MaxRows
            grdSS.Row = lngRow
            If grdSS.RowHidden Then grdSS.RowHidden = False
        Next lngRow
        MsgBox "Print Command Cancelled", vbInformation
        Exit Sub
     End If
    With grdSS
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintJobName = "Laboratory Order" & "_" & Val(udtPatientName.IntId)
        .PrintColHeaders = True
        .PrintRowHeaders = False
        .PrintScalingMethod = PrintScalingMethodSmartPrint
        .PrintCenterOnPageH = True
        '.PrintCenterOnPageV = True
        .PrintMarginBottom = 770
        .PrintBorder = True
        .PrintColor = True
        .PrintOrientation = PrintOrientationPortrait
        .PrintShadows = False
        .PrintSheet
        mbooPrintPreview = False
        For lngRow = 6 To .MaxRows
            .Row = lngRow
            If .RowHidden Then .RowHidden = False
        Next lngRow
        .ReDraw = True
    End With
    Exit Sub
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Print_Lab:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "Lab Print Error", "Print_Lab", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdOrderRemove_Click()

    If mlngGridRow <= 5 Then
        Call MsgBox("Please select a data row first!  ", vbOKOnly + vbInformation + vbApplicationModal + vbDefaultButton1, "Remove Order Item")
    Else
        If grdSS.MaxRows >= mlngStartRow Then
            grdSS.DeleteRows mlngGridRow, 1
            grdSS.MaxRows = grdSS.MaxRows - 1
            mRow = grdSS.MaxRows
        End If
    End If
End Sub

Private Sub InitSignature()
Dim cnNotes As New ADODB.Connection
Dim rsNotes As New ADODB.Recordset
Dim strSQL As String
Dim strCntr As String
Dim booRet As Boolean
Dim lngProvID As Long
Dim strDate As String
    
    On Error GoTo Error_InitSignature
    If clsPatient.IntId = vbNullString Then Exit Sub
    Screen.MousePointer = vbHourglass
    Get_MedProvider
    Clear_Signature
    If cnNotes.State = adStateClosed Then cnNotes.open psConnect(0)
    clsSignature.SignReason = clsPatient.IntId & "_" & udtEnctr.EnctrID & "_" & "Orders"
    strSQL = "SELECT * FROM PROVIDER_SIGNATURES WHERE PROVIDER_ID = " & piProviderID & " AND REF_DATA = '" & clsSignature.SignReason & "'"
    Set rsNotes = cnNotes.Execute(strSQL)
    With rsNotes
        If Not .EOF Then
            clsSignature.Signature = rsNotes.Fields("SIGNATURE")
        End If
    End With
    ccaSignature.ClearSignature
    booRet = ccaSignature.DecodeSignatureValues(clsSignature.Signature, clsSignature.SignReason)
    Screen.MousePointer = vbDefault
    Exit Sub
Error_InitSignature:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, clsSignature.SignReason, "InitSignature", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Public Function SaveSignature()
Dim cnNotes As New ADODB.Connection
Dim rsNotes As New ADODB.Recordset
Dim strSQL As String
Dim strCntr As String
Dim booinsert As Boolean
Dim lngProvID As Long
Dim strDate As String
Dim i%

    On Error GoTo Error_SaveSignature
    Screen.MousePointer = vbHourglass
    If udtEnctr.StartDate <> vbNullString Then
        strDate = Format(udtEnctr.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    Else
        strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    End If
    Clear_Signature
    PsUserID = Get_PsUserID(psConnect(0))
    For i = 0 To cboAttendPhy.ListCount - 1
        If cboAttendPhy.List(i) = cboAttendPhy.Text Then
            lngProvID = cboAttendPhy.ItemData(i)
            Exit For
        End If
    Next i
    clsSignature.Patient = clsPatient.IntId & "_" & udtEnctr.EnctrID
    clsSignature.SignReason = clsPatient.IntId & "_" & udtEnctr.EnctrID & "_" & "Orders"
    If clsSignature.SignReason = vbNullString Then clsSignature.SignReason = "OrderDefault"
    If Not ccaSignature.EmptySignature Then
        clsSignature.Signature = ccaSignature.EncodeSignatureValues(clsSignature.SignReason)
    End If
    If Trim(clsSignature.Signature) <> vbNullString Then
        If cnNotes.State = adStateClosed Then cnNotes.open psConnect(0)
        strSQL = "SELECT * FROM PROVIDER_SIGNATURES WHERE PROVIDER_ID = " & lngProvID & " AND REF_DATA = '" & clsSignature.SignReason & "'"
        Set rsNotes = cnNotes.Execute(strSQL)
        With rsNotes
            If Not .EOF Then
                If rsNotes.Fields("SIGNATURE") <> vbNullString Then
                    clsSignature.Signature = rsNotes.Fields("SIGNATURE")
                    booinsert = False
                    lngProvID = rsNotes.Fields("PROVIDER_ID")
                End If
            Else
                booinsert = True
            End If
        End With
        If lngProvID = 0 Then
            If cnNotes.State = adStateClosed Then cnNotes.open psConnect(0)
            strSQL = "SELECT * FROM PROVIDER_SIGNATURES WHERE REF_DATA = '" & clsSignature.SignReason & "'"
            Set rsNotes = cnNotes.Execute(strSQL)
            With rsNotes
                If Not .EOF Then
                    If rsNotes.Fields("SIGNATURE") <> vbNullString Then
                        clsSignature.Signature = rsNotes.Fields("SIGNATURE")
                        lngProvID = rsNotes.Fields("PROVIDER_ID")
                        booinsert = False
                    End If
                Else
                    booinsert = True
                End If
            End With
        End If
        If lngProvID = 0 And booinsert = False Then
            lngProvID = piProviderID
            booinsert = False
        ElseIf lngProvID <> piProviderID And piProviderID = 0 And lngProvID > 0 Then
            booinsert = False
        End If
        If booinsert Then
            strSQL = "INSERT INTO PROVIDER_SIGNATURES (PROVIDER_ID, SIGNATURE_REF, SIGNATURE, REF_DATA, INSERT_DATE, USER_ID) values " _
                & "(" & piProviderID & ",'" & clsSignature.Patient & "', '" & clsSignature.Signature & "',  '" & clsSignature.SignReason & "','" _
                & strDate & "', '" & PsUserID & "')"
            Set rsNotes = cnNotes.Execute(strSQL)
        Else
             strSQL = "UPDATE PROVIDER_SIGNATURES SET SIGNATURE_REF = '" & clsSignature.Patient & "'," _
                & "SIGNATURE = '" & clsSignature.Signature & "', UPDATE_DATE = '" & strDate & "'," _
                & "USER_ID = '" & PsUserID & "' WHERE REF_DATA = '" & clsSignature.SignReason & "' AND " _
                & "PROVIDER_ID  = " & lngProvID
            Set rsNotes = cnNotes.Execute(strSQL)
        End If
    End If
    Clear_Signature
    Screen.MousePointer = vbDefault
    Exit Function
Error_SaveSignature:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, clsSignature.SignReason, "SaveSignature", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cmdPatient_Click()
Dim lngEnctrID As Long

    Clear_UdtPatient
    Clear_Signature
    udtPatientName.IntId = vbNullString
    mstrComment = vbNullString
    frmListPatient.mbShowEnctr = True
    frmListPatient.Caption = "Search Patients"
    frmListPatient.Show vbModal
    If frmListPatient.mlngListItemData > 0 Then udtPatientName.IntId = frmListPatient.mlngListItemData
    If Trim(udtPatientName.IntId) <> vbNullString And Val(udtPatientName.IntId) > 0 Then
         Get_Patient Val(udtPatientName.IntId)
         lngEnctrID = frmListPatient.mlngListEnctr
         Get_CurrentEnctr lngEnctrID
         Format_Grid
         picLogo.Visible = False
    End If
    Read_Order
    frmListPatient.mbShowEnctr = False
    Exit Sub
End Sub

Private Sub cmdPROC_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer
Dim CellNote As Long
Dim intPrint As Integer
Dim strRef As String
Dim lngRow As Long
Dim varCellValue

    On Error GoTo Error_Print_Procedure
    Screen.MousePointer = vbHourglass
    With grdSS
        .ReDraw = False
        Write_Data
        For lngRow = 6 To .MaxRows
            .Row = lngRow
            .GetText 1, lngRow, varCellValue
            If IsProperCase(varCellValue) <> "Procedure" Then .RowHidden = True
        Next lngRow
    End With
    If clsProvider.Name = vbNullString Then
        modSpecific.piProviderID = piProviderID
        ModPatientData.Get_Providers
    End If
    Remove_Signature
    Add_Signature
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.hwnd)
    If mbooCancelPrint Then
        For lngRow = 6 To grdSS.MaxRows
            grdSS.Row = lngRow
            If grdSS.RowHidden Then grdSS.RowHidden = False
        Next lngRow
        MsgBox "Print Command Cancelled", vbInformation
        Exit Sub
    End If
    With grdSS
        .PrintType = PrintTypeAll
        .PrintJobName = "Procedure Order" & "_" & Val(udtPatientName.IntId)
        .hDCPrinter = Printer.hdc
        .PrintColHeaders = True
        .PrintRowHeaders = False
        .PrintScalingMethod = PrintScalingMethodSmartPrint
        .PrintCenterOnPageH = True
        '.PrintCenterOnPageV = True
        .PrintMarginBottom = 770
        .PrintBorder = True
        .PrintColor = True
        .PrintOrientation = PrintOrientationPortrait
        .PrintShadows = False
        .PrintSheet
        mbooPrintPreview = False
        For lngRow = 6 To .MaxRows
            .Row = lngRow
            If .RowHidden Then .RowHidden = False
        Next lngRow
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Print_Procedure:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "Print Procedure", "Print_Procedure", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdRAD_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer
Dim CellNote As Long
Dim intPrint As Integer
Dim strRef As String
Dim lngRow As Long
Dim varCellValue

    On Error GoTo Error_Print_Radiology
    Screen.MousePointer = vbHourglass
    With grdSS
        .ReDraw = False
        Write_Data
        For lngRow = 6 To .MaxRows
            .Row = lngRow
            .GetText 1, lngRow, varCellValue
            If IsProperCase(varCellValue) <> "Rad" Then .RowHidden = True
        Next lngRow
    End With
    If clsProvider.Name = vbNullString Then
        modSpecific.piProviderID = piProviderID
        ModPatientData.Get_Providers
    End If
    Remove_Signature
    Add_Signature
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation
        For lngRow = 6 To grdSS.MaxRows
            grdSS.Row = lngRow
            If grdSS.RowHidden Then grdSS.RowHidden = False
        Next lngRow
        Exit Sub
    End If
    With grdSS
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintJobName = "Radiology Order" & "_" & Val(udtPatientName.IntId)
        .PrintColHeaders = True
        .PrintRowHeaders = False
        .PrintScalingMethod = PrintScalingMethodSmartPrint
        .PrintCenterOnPageH = True
        '.PrintCenterOnPageV = True
        .PrintMarginBottom = 770
        .PrintBorder = True
        .PrintColor = True
        .PrintOrientation = PrintOrientationPortrait
        .PrintShadows = False
        .PrintSheet
        mbooPrintPreview = False
        For lngRow = 6 To .MaxRows
            .Row = lngRow
            If .RowHidden Then .RowHidden = False
        Next lngRow
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Print_Radiology:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "Print Radiology", "Print_Radiology", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdSave_Click()
Dim Arr
Dim lngRow As Long
    lngRow = grdSS.MaxRows
    If lngRow > 3 Then
        Arr = Split(LookUpCellTag("Physician_Signature"), "|")
        If UBound(Arr) <= 0 Then
            MsgBox "Please Sign Patient Order, Thank You", vbInformation, "Patient Orders"
            Exit Sub
        End If
        Write_Data
        Clear_Order
        If mbooNoSave Then
            Exit Sub
        Else
            ccaSignature.ClearSignature
            Unload Me
        End If
    Else
        MsgBox "Please select Close when not saving data!", vbInformation, "Patient Orders"
        Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cmdSaveSign_Click()
    SaveSignature
    Add_Signature
    fraSign.Visible = False
    Exit Sub
End Sub

Private Sub cmdSelect_Click()
Dim i%
Dim strICD As String
Dim CellNote As Long

    For i = 0 To lstICD.ListCount - 1
        If lstICD.Selected(i) Then
            If strICD <> vbNullString Then
                strICD = lstICD.List(i)
            Else
                strICD = Trim(strICD) & Chr(9) & lstICD.List(i)
            End If
        End If
    Next i
    
    With grdSS
        .ReDraw = False
        If strICD <> vbNullString Then
            .BlockMode = True
                .Col = 4
                .Col2 = 4
                .Row = mlngGridRow
                .Row2 = mlngGridRow
                .SetText 4, mlngGridRow, CVar(strICD)
                .TypeEditMultiLine = True   'Make multi-line
                .TypeTextWordWrap = True  'Make multi-line
                .TypeMaxEditLen = 20000
                .Text = CVar(strICD)
                CellNote = grdSS.MaxTextCellWidth
                If .RowHeight(mlngGridRow) < CellNote Then
                    .RowHeight(mlngGridRow) = CellNote
                End If
                .TypeHAlign = TypeHAlignLeft
                .TypeVAlign = TypeVAlignCenter
            .BlockMode = False
        End If
        .ReDraw = True
        grdSS.SetText 5, mlngGridRow, "Routine"
    End With
    fraRecICD.Visible = False
End Sub

Private Sub cmdSigExit_Click()
    fraSign.Visible = False
End Sub

Private Sub cmdSign_Click()
    InitSignature
    fraSign.Visible = True
    Exit Sub
End Sub

Private Sub cmdStandNew_Click()
Dim strSQL As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim Arr
Dim lngRow As Long
Dim lngCol As Long
Dim strProvider As String
Dim i%

    On Error GoTo Error_Read_New_Patient_Order
    Screen.MousePointer = vbHourglass
    If piProviderID = 0 Then piProviderID = modSpecific.piProviderID
    Get_Providers piProviderID
    strProvider = clsProvider.Name
    If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    strSQL = "SELECT * FROM DEFAULT_ORDERS WHERE PROVIDER_ID = " & piProviderID & " AND CONDITION = 'New Patient'"
    Set rsLab = cnLab.Execute(strSQL)
    With rsLab
        Do Until .EOF
            Clear_LabString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PROC_ID"
                                 udtLaboratory.Lab_Id = Fld.Value
                            Case "PROC_GROUP"
                                udtLaboratory.OrderType = Fld.Value
                            Case "AGE_VALUE"
                                udtLaboratory.AgeRange = Fld.Value
                            Case "AGE"
                                udtLaboratory.Age = Fld.Value
                            Case "GENDER"
                                udtLaboratory.Gender = Fld.Value
                            Case "CONDITION"
                                 udtLaboratory.Med_Condition = Fld.Value
                            Case "LAB_CODE"
                                udtLaboratory.CPT_Code = Fld.Value
                            Case "LAB_DESC"
                                udtLaboratory.CPT_Desc = Fld.Value
                            Case "ICD_CODE"
                                udtLaboratory.Diag_Code = Fld.Value
                            Case "ICD_DESC"
                                udtLaboratory.Diag_Desc = Fld.Value
                            Case "PROVIDER_ID"
                                udtLaboratory.ProviderID = Fld.Value
                            Case "CONSULT"
                                udtLaboratory.PatientNotice = Fld.Value
                            Case "INTRVN_PRIORITY"
                                udtLaboratory.Lab_Priority = Fld.Value
                            Case "COMMENTS"
                                udtLaboratory.Spec_Instruct = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtLaboratory.Lab_Id > 0 Then
                Add_Order
                lngRow = grdSS.MaxRows
                For lngCol = 1 To grdSS.MaxCols
                    grdSS.Col = lngCol
                    grdSS.Row = lngRow
                    If grdSS.CellTag <> vbNullString Then
                        Select Case UCase(grdSS.CellTag)
                            Case "ORDER_TYPE"
                                Set_GridCurSel grdSS, udtLaboratory.OrderType, lngCol, lngRow, "1", "ORDER_TYPE"
                            Case "SPECIAL_INSTRUCTIONS"
                                Format_CellTypeEdit grdSS, lngRow, lngCol, "1", udtLaboratory.Spec_Instruct, "TypeHAlignCenter", "SPECIAL_INSTRUCTIONS"
                            Case "CPT_CODE"
                                Format_CellTypeEdit grdSS, lngRow, lngCol, "1", udtLaboratory.CPT_Desc & " - " & udtLaboratory.CPT_Code, "TypeHAlignCenter", "CPT_CODE"
                            Case "ICD_CODE"
                                Format_CellTypeEdit grdSS, lngRow, lngCol, "1", udtLaboratory.Diag_Desc & " - " & udtLaboratory.Diag_Code, "TypeHAlignCenter", "ICD_CODE"
                            Case "STATUS"
                                Set_GridCurSel grdSS, udtLaboratory.Lab_Priority, lngCol, lngRow, "1", "STATUS"
                        End Select
                    End If
                Next lngCol
            End If
           Loop
        End With
        grdSS.Col = 1
        grdSS.Row = 1
    If cnLab.State = adStateOpen Then
        Set cnLab = Nothing
        Set rsLab = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Read_New_Patient_Order:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, " Read_New_Patient_Order", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdTerms_Click()
    frmCPTMaint.Show vbModal
    Exit Sub
End Sub

Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
        Clear_Order
        mlngHeadRow = 0
        PsUserID = Get_PsUserID(psConnect(0))
        mlngInsertRows = 0
        Get_DefaultFacility
        Clear_EnctrString
        Clear_BillStrings
        centerForm Me
        mlngGridRow = 0
    End If
    Exit Sub
End Sub

Private Sub Form_Load()
    mFirst = True
    Exit Sub
End Sub

Private Function Add_ReportHeaders()
Dim lngRow As Long
Dim strList As String
Dim lngColWidth As Long
Dim varCellValue
Dim lngGridRow As Long
Dim varDiagnosis1
Dim varDiagnosis2
Dim varMedDiag
Dim strHeader As String
Dim strDate As String
'
    With grdSS
        .ClearRange 1, SpreadHeader, .MaxCols, .MaxRows, False
        Screen.MousePointer = vbHourglass
        grdSS.ProcessTab = True
        grdSS.FontBold = True
        grdSS.ColHeaderRows = 4
        grdSS.RowHeadersShow = False
        grdSS.MaxRows = 0
       .ColWidth(1) = 15
       .ColWidth(2) = 20
       .ColWidth(3) = 15
       .ColWidth(4) = 20
       .ColWidth(5) = 14
       .ColWidth(6) = 17
        lngRow = 1
        Get_DefaultFacility
        If udtEnctr.ProviderID > 0 Then
            Get_Providers udtEnctr.ProviderID
        Else
            Get_Providers
        End If
        grdSS.ColHeaderRows = 4
        If udtEnctr.StartDate <> vbNullString Then
            strDate = Format(udtEnctr.StartDate, "MMMM DD YYYY")
        Else
            strDate = Format(Now, "MMMM DD YYYY")
        End If
        strHeader = udtFacility.Name & vbNewLine & " " & udtFacility.Address & " " & udtFacility.Address2 _
                & vbNewLine & udtFacility.CityStateZip & vbNewLine & "Telephone: " & udtFacility.Telephone & "   :   " & "Fax: " & udtFacility.Facsimile
        Format_CellTypePicture grdSS, SpreadHeader, 1, "1", udtFacility.Logo, "TypeHAlignCenter", "LOGO", "75"
        Format_CellTypeEdit grdSS, SpreadHeader, 2, "7", strHeader, "TypeHAlignCenter", "", "", HeaderColor, 14
        Add_Spacer SpreadHeader + 1, grdSS
        Format_CellTypeEdit grdSS, SpreadHeader + 2, 1, "7", "PATIENT ORDERS" & "   DATED:" & "   " & strDate, "TypeHAlignCenter", "", "", HeaderColor, 14
        .SetCellBorder 1, SpreadHeader + 2, 7, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdSS
        .MaxRows = lngRow
        Format_CellTypeEdit grdSS, lngRow, 1, "1", "PATIENT NAME", "TypeHAlignLeft", "HEADER"
        Format_CellTypeEdit grdSS, lngRow, 2, "1", clsPatient.Name, "TypeHAlignLeft", "PATIENT_NAME"
        Format_CellTypeEdit grdSS, lngRow, 5, "1", "PATIENT EMR ID", "TypeHAlignLeft", "HEADER"
        Format_CellTypeEdit grdSS, lngRow, 6, "2", CStr(clsPatient.IntId), "TypeHAlignLeft", "EMR_ID"
        grdSS.SetCellBorder 1, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdSS.SetCellBorder 5, lngRow, 7, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngRow = lngRow + 1
        grdSS.MaxRows = lngRow
        Format_CellTypeEdit grdSS, lngRow, 1, "1", "SOCIAL SECURITY", "TypeHAlignLeft", "HEADER"
        Format_CellTypeEdit grdSS, lngRow, 2, "2", clsPatient.SS, "TypeHAlignLeft", "SS"
        Format_CellTypeEdit grdSS, lngRow, 5, "1", "EXTERNAL ID", "TypeHAlignLeft", "HEADER"
        Format_CellTypeEdit grdSS, lngRow, 6, "2", CStr(clsPatient.ExtID), "TypeHAlignLeft", "EXT_ID"
        grdSS.SetCellBorder 1, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdSS.SetCellBorder 5, lngRow, 7, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngRow = lngRow + 1
        .MaxRows = lngRow
        Format_CellTypeEdit grdSS, lngRow, 1, "1", "DATE OF BIRTH", "TypeHAlignLeft", "HEADER"
        Format_CellTypeEdit grdSS, lngRow, 2, "2", clsPatient.DOB, "TypeHAlignLeft", "BIRTH"
        Format_CellTypeEdit grdSS, lngRow, 5, "1", "PROVIDER NAME", "TypeHAlignLeft", "HEADER"
        Format_CellTypeEdit grdSS, lngRow, 6, "2", clsProvider.Name, "TypeHAlignLeft", "PROVIDER_NAME"
        grdSS.SetCellBorder 1, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdSS.SetCellBorder 5, lngRow, 7, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    End With
    Screen.MousePointer = vbDefault
    Exit Function
End Function
Private Sub Format_Grid()
Dim lngCol As Long
Dim CellNote As Long
Dim lngRow As Long

    On Error GoTo Error_Format_Grid
    Screen.MousePointer = vbHourglass
    Add_ReportHeaders
    lngRow = grdSS.MaxRows + 1
    grdSS.MaxRows = lngRow
    Add_Spacer lngRow, grdSS
    With grdSS
        .ReDraw = False
        lngRow = grdSS.MaxRows + 1
        .MaxRows = lngRow
        .Col = 0
        Format_CellTypeEdit grdSS, lngRow, 0, "1", "GROUP", "TypeHAlignCenter", "GROUP", "", HeaderColor
        grdSS.ColHidden = True
        Format_CellTypeEdit grdSS, lngRow, 1, "1", "TYPE", "TypeHAlignCenter", "TYPE", "", HeaderColor
        Format_CellTypeEdit grdSS, lngRow, 2, "1", "ITEM DESCRIPTION" & vbNewLine & "CPT CODE", "TypeHAlignCenter", "ITEM DESCRIPTION", "", HeaderColor
        Format_CellTypeEdit grdSS, lngRow, 3, "1", "MODIFIER" & vbNewLine & "CPT CODE", "TypeHAlignCenter", "MODIFIER", "", HeaderColor
        Format_CellTypeEdit grdSS, lngRow, 4, "1", "REASON FOR TEST" & vbNewLine & "ICD CODE", "TypeHAlignCenter", "", "", HeaderColor
        Format_CellTypeEdit grdSS, lngRow, 5, "1", "PRIORITY/STATUS", "TypeHAlignCenter", "STATUS", "", HeaderColor
        Format_CellTypeEdit grdSS, lngRow, 6, "2", "SPECIAL INSTRUCTIONS", "TypeHAlignCenter", "SPECIAL INSTRUCTIONS", "", HeaderColor
        grdSS.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdSS.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdSS.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdSS.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdSS.SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdSS.SetCellBorder 6, lngRow, 7, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdSS.MaxCols = 7
        lngRow = grdSS.MaxRows + 1
        grdSS.MaxRows = lngRow
        mlngHeadRow = lngRow
        Add_Spacer lngRow, grdSS
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Grid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Grid", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Function Get_CPTGroups()
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset

    On Error GoTo ERROR_Get_CPTGroups
    mstrSubGroup = vbNullString
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        cnCPT.open psConnect(1)
    Else
        cnCPT.open psConnect(1)
    End If
    strSQL = "Select * from INTRVN_SUBGROUPS "
    Set rsCPT = cnCPT.Execute(strSQL)
    With rsCPT
        Do Until .EOF
            If mstrSubGroup = vbNullString Then
                mstrSubGroup = Trim(!SUBGROUP_DESC)
            Else
                mstrSubGroup = mstrSubGroup & Chr(9) & Trim(!SUBGROUP_DESC)
            End If
        .MoveNext
        Loop
    End With
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set cnCPT = Nothing
     End If
    Exit Function
ERROR_Get_CPTGroups:
    Write_Error Err.Description, Err.Number, strSQL, "Get_CPTGroups", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
    On Error Resume Next
    ccaSignature.ClearSignature
    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.objects
      Set obj = Nothing
    Next obj
    Unload Me
    Clear_Order
    Exit Sub
End Sub

Private Sub fraGridTerm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    flngDeltaX = X
    flngDeltaY = Y
End Sub


Private Sub fraGridTerm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        fraGridTerm.Left = fraGridTerm.Left + X - flngDeltaX
        fraGridTerm.Top = fraGridTerm.Top + Y - flngDeltaY
    End If
End Sub


Private Sub fraRecICD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    flngDeltaX = X
    flngDeltaY = Y
End Sub


Private Sub fraRecICD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
         fraRecICD.Left = fraRecICD.Left + X - flngDeltaX
         fraRecICD.Top = fraRecICD.Top + Y - flngDeltaY
    End If
End Sub


Private Sub grdSS_Change(ByVal Col As Long, ByVal Row As Long)
Dim CellNote As Long
Dim varCellValue

    
    With grdSS
        .Col = Col
        .Row = Row
        If .CellTag = "SPECIAL_INSTRUCTIONS" Then
            .GetText mlngGridCol, mlngGridRow, varCellValue
            If varCellValue <> vbNullString Then
                Format_CellTypeEdit grdSS, Row, Col, "2", CStr(varCellValue), "TypeHAlignLeft", "SPECIAL_INSTRUCTIONS"
                grdSS.SetCellBorder 6, Row, 7, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        End If
        .LeftCol = 1
    End With
    If mlngGridCol > 6 Then
        grdSS.LeftCol = 1
    End If
    Exit Sub
End Sub


Private Sub grdSS_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim strType As String
Dim strCurrentDiag As String

    If udtEnctr.EnctrID > 0 Then
        With grdSS
            If Col = 1 Then
                .Col = Col
                .Row = Row
                If .TypeComboBoxCurSel >= 0 Then .TypeComboBoxIndex = .TypeComboBoxCurSel
                strType = .TypeComboBoxString
                If strType <> vbNullString Then
                    If mstrPractice = "Psychiatry" Then
                        strCurrentDiag = Get_ClinicalNotes("FN_PsyMse", "Diagnosis", "cctAxis1", 1, udtEnctr.EnctrID)
                        If strCurrentDiag = vbNullString Then
                            strCurrentDiag = Get_ClinicalNotes("FN_PsyMse", "Diagnosis", "cctAxisI", 1, udtEnctr.EnctrID)
                        End If
                        If Trim(strCurrentDiag) <> vbNullString Then
                            Format_CellTypeEdit grdSS, Row, 4, "1", strCurrentDiag, "TypeHAlignLeft", "ICD_CODE"
                        End If
                    End If
                End If
            End If
        End With
    End If
    Exit Sub
End Sub
Private Sub grdSS_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim strSQL As String
Dim varRS
Dim lngrecs As Long
Dim booRet As Boolean
Dim i%
Dim lngX As Long
Dim lngY As Long
Dim lngWidth As Long
Dim lngHeight As Long
Dim varDiag
Dim CellNote As Long
Dim varCellValue
Dim Arr
Dim ArrCPT
Dim ArrDesc
Dim ArrCode
Dim lngCol As Long
Dim iCol%
Dim strCurrentDiag As String
Dim strTag As String

    varDiag = vbNullString
    clsGridFormat.CellTag = vbNullString
    With grdSS
         mlngGridRow = Row
        .ReDraw = False
        .GetText 1, Row, varCellValue
        If Trim(varCellValue) <> vbNullString Then strTag = Trim(varCellValue)
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then clsGridFormat.CellTag = .CellTag
        
        Select Case clsGridFormat.CellTag
            Case "CPT_CODE"
                .Col = 2
                .Row = Row
                .GetCellPos .Col, .Row, lngX, lngY, lngWidth, lngHeight
                fpDeltaTop = lngY + 51.5
                fpDeltaLeft = lngX + 50
                fraGridTerm.ZOrder 0
                fraGridTerm.Top = .Top + fpDeltaTop
                fraGridTerm.Left = .Left + fpDeltaLeft
                Load_GridBox strTag, "CPT_CODE"
                lblSearch.Visible = False
                cctSearch.Visible = False
                fraGridTerm.Visible = True
            Case "MODIFIER"
                .Col = 3
                .Row = Row
                .GetCellPos .Col, .Row, lngX, lngY, lngWidth, lngHeight
                fpDeltaTop = lngY + 51.5
                fpDeltaLeft = lngX + 50
                fraGridTerm.ZOrder 0
                fraGridTerm.Top = .Top + fpDeltaTop
                fraGridTerm.Left = .Left + fpDeltaLeft
                If Trim(strTag) <> vbNullString Then Load_GridBox "MODIFIER", "MODIFIER"
                lblSearch.Visible = False
                cctSearch.Visible = False
                fraGridTerm.Visible = True
            Case "ICD_CODE"
                .Col = 4
                .Row = Row
                .GetCellPos .Col, .Row, lngX, lngY, lngWidth, lngHeight
                fpDeltaTop = lngY + 51.5
                fpDeltaLeft = lngX + 50
                fraGridTerm.ZOrder 0
                fraGridTerm.Top = .Top + fpDeltaTop
                fraGridTerm.Left = .Left + fpDeltaLeft
                If Trim(strTag) <> vbNullString Then Load_GridBox "ICD", "ICD"
                lblSearch.Visible = True
                cctSearch.Visible = True
                fraGridTerm.Visible = True
        End Select
        .ReDraw = True
    End With
    Exit Sub
End Sub


Private Sub grdss_KeyUp(KeyCode As Integer, Shift As Integer)

    Debug.Print
    
End Sub


Private Sub grdSS_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    With grdSS
        If Row >= 0 Then
            .BlockMode = True
                .Col = Col
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                 .BackColor = &HFFFFFF
             .BlockMode = False
        End If
    End With
    Exit Sub
    
End Sub

Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
    With grdTerms
        cctSearch.Text = vbNullString
        If Row >= 0 Then
            .BlockMode = True
                .Col = Col
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                 .BackColor = &HFFC0FF
             .BlockMode = False
        End If
    End With
    Exit Sub
End Sub


Private Sub grdTerms_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    With grdTerms
        cctSearch.Text = vbNullString
        .BlockMode = True
            .Col = Col
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
             .BackColor = &H80000005
         .BlockMode = False
    End With
    Exit Sub
End Sub

Public Function Get_MedProvider()
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
Dim strName As String
Dim strCity As String
Dim strState As String
Dim strZip As String
Dim strZipExt As String
Dim booSkip As Boolean
Dim i%
Dim lngProvID As Long

    On Error GoTo Error_Get_MedProvider
    Screen.MousePointer = vbHourglass
    cboAttendPhy.Clear
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    strSQL = "Select * FROM PROVIDERS WHERE ACTIVE = 1 ORDER BY TITLE, LAST_NAME, FIRST_NAME"
    Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        Do Until .EOF
            Clear_Provider
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            lngProvID = 0
            booSkip = True
            If Not IsNull(!Title) Then
                strTitle = (!Title)
                If TextFound("MD", strTitle, Len("MD"), True) Or _
                    TextFound("M.D.", strTitle, Len("M.D."), True) Or _
                    TextFound("D.O.", strTitle, Len("D.O."), True) Or _
                    TextFound("DO", strTitle, Len("DO"), True) Then
                        booSkip = False
                    End If
            End If
            If Not booSkip Then
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Fld.Value <> vbNullString Then
                            If UCase(Fld.Name) = "FIRST_NAME" Then
                                strFirst = Fld.Value
                            ElseIf UCase(Fld.Name) = "MIDDLE_NAME" Then
                                strMiddle = Fld.Value
                            ElseIf UCase(Fld.Name) = "LAST_NAME" Then
                                strLast = Fld.Value
                            ElseIf UCase(Fld.Name) = "SUFFIX" Then
                                strSuffix = Fld.Value
                            ElseIf UCase(Fld.Name) = "TITLE" Then
                                strTitle = Fld.Value
                            ElseIf UCase(Fld.Name) = "PROVIDER_ID" Then
                                lngProvID = Fld.Value
                            End If
                        End If
                    End If
                Next
            End If
            .MoveNext
            If Trim(strFirst) <> vbNullString Then
                 strName = Fix_Name(strFirst)
             End If
            If Trim(strMiddle) <> vbNullString Then
                strMiddle = IsProperCase(strMiddle)
                strName = strName & " " & Left(strMiddle, 1)
             End If
            If Trim(strLast) <> vbNullString Then
                 strName = strName & " " & Fix_Name(strLast)
             End If
            If Trim(strSuffix) <> vbNullString And Trim(strName) <> vbNullString Then
                 strName = strName & ", " & Fix_Suffix(strSuffix)
             End If
            If Trim(strTitle) <> vbNullString And Trim(strName) <> vbNullString Then
                 strName = strName & ", " & Fix_Title(strTitle)
             End If
            If strName <> vbNullString Then
                cboAttendPhy.AddItem strName
                If IsNumeric(lngProvID) Then cboAttendPhy.ItemData(cboAttendPhy.NewIndex) = lngProvID
            End If
        Loop
    End With
    For i = 0 To cboAttendPhy.ListCount - 1
        If cboAttendPhy.ItemData(i) = piProviderID Then
            cboAttendPhy.Text = cboAttendPhy.List(i)
            Exit For
        End If
    Next i
    If cnMedProv.State = adStateOpen Then
        cnMedProv.Close
        Set cnMedProv = Nothing
        Set rsMedProv = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_MedProvider:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_MedProvider", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
