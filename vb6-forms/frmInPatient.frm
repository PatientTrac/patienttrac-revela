VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmInPatient 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Electronic Medical Record"
   ClientHeight    =   10530
   ClientLeft      =   8415
   ClientTop       =   2415
   ClientWidth     =   15765
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   702
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1051
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox cctPatNote 
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
      Height          =   915
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      ToolTipText     =   "Provider Patient Notes"
      Top             =   1140
      Width           =   3705
   End
   Begin FPSpreadADO.fpSpread grdSpec 
      Height          =   2835
      Left            =   0
      TabIndex        =   24
      Top             =   4950
      Width           =   4020
      _Version        =   524288
      _ExtentX        =   7091
      _ExtentY        =   5001
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
      RetainSelBlock  =   0   'False
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmInPatient.frx":058A
      UnitType        =   2
      UserResize      =   0
      Appearance      =   2
      TextTipDelay    =   30
      ClipboardOptions=   0
      AppearanceStyle =   0
   End
   Begin VB.Frame fraButton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   13740
      TabIndex        =   17
      Top             =   30
      Width           =   1965
      Begin PT_XP_Button.PT_XP cmdAppt 
         Height          =   285
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Width           =   1035
         _ExtentX        =   1826
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
         Picture         =   "frmInPatient.frx":0A9E
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Appt"
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatient.frx":1038
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCloseList 
         Height          =   285
         Left            =   990
         TabIndex        =   18
         Top             =   30
         Width           =   975
         _ExtentX        =   1720
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
         Picture         =   "frmInPatient.frx":138A
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatient.frx":1D9C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   285
         Left            =   0
         TabIndex        =   19
         ToolTipText     =   "Save, Restore, Close"
         Top             =   30
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmInPatient.frx":20EE
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatient.frx":2488
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRefresh 
         Height          =   285
         Left            =   990
         TabIndex        =   21
         Top             =   -750
         Width           =   1215
         _ExtentX        =   2143
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
         Picture         =   "frmInPatient.frx":27DA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Refresh"
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmInPatient.frx":2D74
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.TextBox txtRoomNumber 
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
      Height          =   330
      Left            =   2460
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1395
   End
   Begin FPSpreadADO.fpSpread grdOrders 
      Height          =   2715
      Left            =   0
      TabIndex        =   8
      Top             =   7800
      Width           =   4020
      _Version        =   524288
      _ExtentX        =   7091
      _ExtentY        =   4789
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
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmInPatient.frx":30C6
      UnitType        =   2
      UserResize      =   0
      Appearance      =   2
      TextTipDelay    =   30
      ClipboardOptions=   0
      AppearanceStyle =   0
   End
   Begin VB.ComboBox cboPurpose 
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
      Left            =   0
      TabIndex        =   5
      Top             =   780
      Width           =   3885
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
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin FPSpreadADO.fpSpread grdPatient 
      Height          =   2805
      Left            =   0
      TabIndex        =   2
      Top             =   2100
      Width           =   4020
      _Version        =   524288
      _ExtentX        =   7091
      _ExtentY        =   4948
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
      RetainSelBlock  =   0   'False
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmInPatient.frx":35F3
      UnitType        =   2
      UserResize      =   0
      Appearance      =   2
      TextTipDelay    =   30
      ClipboardOptions=   0
      AppearanceStyle =   0
   End
   Begin PT_XP_Button.PT_XP cmdAlert 
      Height          =   435
      Left            =   3720
      TabIndex        =   9
      ToolTipText     =   "Manage Alerts"
      Top             =   11790
      Width           =   495
      _ExtentX        =   873
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
      Picture         =   "frmInPatient.frx":3B2B
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthMode       =   1
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   13003064
      ForeColorFocus  =   13003064
      ForeColorDown   =   13003064
      PictureDisabled =   "frmInPatient.frx":42A5
      PictureAlignment=   0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdNewPat 
      Height          =   390
      Left            =   90
      TabIndex        =   10
      Top             =   2850
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
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
      Picture         =   "frmInPatient.frx":49B7
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "New Patient"
      Depth           =   5
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatient.frx":4F51
      PictureAlignment=   0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdGrdView 
      Height          =   450
      Left            =   120
      TabIndex        =   13
      Top             =   3750
      Visible         =   0   'False
      Width           =   2100
      _ExtentX        =   3704
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
      Picture         =   "frmInPatient.frx":52A3
      AppearanceMode  =   1
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "View Encounters"
      Depth           =   5
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatient.frx":583D
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdUpdate 
      Height          =   390
      Left            =   1320
      TabIndex        =   14
      ToolTipText     =   "Update Patient Encounter Information"
      Top             =   2850
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
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
      Picture         =   "frmInPatient.frx":5B8F
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Update Status"
      Depth           =   5
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatient.frx":6129
      ShowFocus       =   -1  'True
   End
   Begin MSComCtl2.DTPicker ccdAdmit 
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   2385
      _ExtentX        =   4207
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
      CustomFormat    =   "MM/dd/yyy hh:mm:ss"
      DateIsNull      =   -1  'True
      Format          =   119668739
      CurrentDate     =   38862.7355555556
      MinDate         =   38353
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
      Left            =   4500
      ScaleHeight     =   315
      ScaleWidth      =   2805
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   2865
   End
   Begin PT_XP_Button.PT_XP cmdReportBldr 
      Height          =   375
      Left            =   180
      TabIndex        =   22
      ToolTipText     =   "Report Builder"
      Top             =   11910
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
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
      Picture         =   "frmInPatient.frx":647B
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatient.frx":6A15
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdReset 
      Height          =   375
      Left            =   660
      TabIndex        =   23
      ToolTipText     =   "Report Builder"
      Top             =   11910
      Width           =   675
      _ExtentX        =   1191
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
      Caption         =   "Reset"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInPatient.frx":6D67
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdStatus 
      Height          =   315
      Left            =   -60
      TabIndex        =   26
      ToolTipText     =   "Intenal Patient Note"
      Top             =   1200
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
      Picture         =   "frmInPatient.frx":70B9
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmInPatient.frx":7653
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblApptDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Date"
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
      Height          =   180
      Left            =   990
      TabIndex        =   16
      Top             =   1350
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblRoomNumber 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Room Number"
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
      Left            =   2490
      TabIndex        =   12
      Top             =   1350
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CURRENT PATIENT:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4320
      TabIndex        =   7
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose of Assessment or Evaluation"
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
      Height          =   390
      Left            =   15
      TabIndex        =   6
      Top             =   570
      Width           =   3705
   End
   Begin VB.Label lblFacility 
      BackStyle       =   0  'Transparent
      Caption         =   "Facility"
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
      Left            =   15
      TabIndex        =   4
      Top             =   30
      Width           =   855
   End
   Begin VB.Label lblPatName 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4500
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Menu mnu_ProgressNote 
      Caption         =   "Submit CCD"
   End
End
Attribute VB_Name = "frmInPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public mstrMyControl As String
Public piSelPatientID As Long
Public piSelEnctrID As Long
Private piPatientExtID As Long
Private mintActXId As Long
Private lintComponentCounter As Integer
Private lintTabCounter As Integer
Private lbooMissingClass As Boolean
Private lstrClassName As String
Private lstrLibrary As String
Private mlngCurrentPatientID As Long
Private mlngCurrentEncounterID As Long
Private mstrGridName As String
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private mlngFormID As Long
Private mstrIniFile As String
Private mbooByPass As Boolean
Private mcolActxComp As New Collection
Private mbooSelectPat As Boolean
Private mbooRecurssive As Boolean
Private mlngLastPatientId As Long
Private mlngLastEncounterId As Long
Private mObjSelectPatient As PatientTracTypeLib.iClinical
Public mintindex As Integer
Private mFirst As Boolean
Private mstrRemarks As String
Private mbooNext As Boolean
Private mbooPrev As Boolean
Public EnctrType As String
Private mbooNoData As Boolean
Private mlngSelFacID As Long
Private mlngGridRow As Long
Private mlngGridCol As Long
Public mstrLoadType As String


Private Sub Format_PatientGrid()

    With grdPatient
        .Redraw = False
        .MaxCols = 3
        .MaxRows = 0
        .ClearRange 0, 0, grdPatient.MaxCols, grdPatient.MaxRows, False
        .ColWidth(1) = 2000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .SetText 1, 0, "PATIENT NAME"
        .SetText 2, 0, "REASON"
        .SetText 3, 0, "ROOM NO"
        'init text tips
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = 0
            .Row2 = 0
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        .Redraw = True
    End With
    Load_Patients
    Exit Sub
End Sub

Private Sub Format_EncntrGrid()

    With grdPatient
        .Redraw = False
        .MaxCols = 3
        .MaxRows = 0
        .RowHeadersShow = False
        .ClearRange 0, 0, grdPatient.MaxCols, grdPatient.MaxRows, False
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
         Format_CellTypeEdit grdPatient, SpreadHeader, 1, "3", "VISIT HISTORY", "TypeHAlignCenter", "TITLE", NormalColor
         grdPatient.SetCellBorder 1, SpreadHeader, 3, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
         Add_Spacer SpreadHeader + 1, grdPatient
         Format_CellTypeEdit grdPatient, SpreadHeader + 2, 1, "1", "VISIT DATE", "TypeHAlignCenter", "VISIT DATE", HeaderColor
         grdPatient.SetCellBorder 1, SpreadHeader + 2, 1, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
         Format_CellTypeEdit grdPatient, SpreadHeader + 2, 2, "1", "REASON", "TypeHAlignCenter", "REASON", HeaderColor
         grdPatient.SetCellBorder 2, SpreadHeader + 2, 2, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
         Format_CellTypeEdit grdPatient, SpreadHeader + 2, 3, "1", "TYPE", "TypeHAlignCenter", "TYPE", HeaderColor
         grdPatient.SetCellBorder 3, SpreadHeader + 2, 3, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
         Add_Spacer SpreadHeader + 3, grdPatient
        .Redraw = True
    End With
    Load_Enctrs
    Exit Sub
End Sub

Private Sub Format_SpecGrid()

    With grdSpec
        .Redraw = False
        .ColHeaderRows = 2
        .MaxCols = 1
        .MaxRows = 0
        .RowHeadersShow = False
        .ClearRange 0, 0, grdSpec.MaxCols, grdSpec.MaxRows, False
        .ColWidth(1) = 3650
         Format_CellTypeEdit grdSpec, SpreadHeader, 1, "1", "RATING SCALES", "TypeHAlignCenter", "TITLE", HeaderColor
         grdSpec.SetCellBorder 1, SpreadHeader, 1, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
         Add_Spacer SpreadHeader + 1, grdSpec
        .Redraw = True
    End With
    Load_Scales
    Exit Sub
End Sub
Private Sub Get_ActionItems()
Dim cnReport As New ADODB.Connection
Dim rsReport As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim strInsert As String
Dim lngDateDiff As Long
Dim strAlert As String
Dim Arr
Dim Idx%

    On Error GoTo Error_Read_Data
    Screen.MousePointer = vbHourglass
    Format_PatientAlerts piSelPatientID
    Clear_FormatAlert
    If cnReport.State = adStateClosed Then cnReport.open psConnect(0)
    strSQL = "Load_Active_Alert " & piSelPatientID
    Set rsReport = cnReport.Execute(strSQL)
    With rsReport
        Do Until .EOF
            Clear_FormatAlert
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "EFFECTIVE_DATE"
                                udtFormatAlert.EffectDate = Trim(Fld.Value)
                            Case "ALERT"
                                udtFormatAlert.Alert = Trim(Fld.Value)
                            Case "PRIORITY"
                                udtFormatAlert.AlertType = Trim(Fld.Value)
                            Case "RESPONSIBLE_PARTY"
                                udtFormatAlert.Assign = Trim(Fld.Value)
                            Case "DESCRIPTION"
                                udtFormatAlert.AlertDesc = Trim(Fld.Value)
                            Case "INTERVENTION"
                                udtFormatAlert.Intervention = Trim(Fld.Value)
                            Case "STATUS"
                                udtFormatAlert.Status = Trim(Fld.Value)
                            Case "REASON_FLAG"
                                udtFormatAlert.PatData = Trim(Fld.Value)
                            Case "LAST_ORDER_DATE"
                                udtFormatAlert.LastDate = Format(Fld.Value, "mm/dd/yyyy")
                            Case "ALERT_TIME"
                                udtFormatAlert.Time = Trim(Fld.Value)
                            Case "ALERT_TIME_VALUE"
                                udtFormatAlert.TimeValue = Trim(Fld.Value)
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If Val(udtFormatAlert.Time) > 0 Then
                If LCase(udtFormatAlert.TimeValue) = "months" Then
                    udtFormatAlert.Time = Val(udtFormatAlert.Time) * 30
                End If
            End If
            If strAlert <> udtFormatAlert.Alert Then
                If Format(udtFormatAlert.LastDate, "mm/dd/yyyy") <> "01/01/1900" Then
                    If clsEncounter.StartDate <> vbNullString Then
                        lngDateDiff = DateDiff("d", Format(udtFormatAlert.LastDate, "mm/dd/yyyy"), Format(clsEncounter.StartDate, "mm/dd/yyyy"))
                    End If
                End If
                If udtFormatAlert.Alert <> vbNullString And udtFormatAlert.PatData <> vbNullString Then
                    'clean this up
                    Arr = Split(udtFormatAlert.PatData, vbNewLine)
                    If UBound(Arr) > 0 Then
                        udtFormatAlert.PatData = vbNullString
                        For Idx = LBound(Arr) To UBound(Arr)
                            If Trim(Arr(Idx)) <> vbNullString And Len(Trim(Arr(Idx))) > 2 Then
                                If udtFormatAlert.PatData = vbNullString Then
                                    udtFormatAlert.PatData = Trim(Arr(Idx))
                                ElseIf Not TextFound(Trim(Arr(Idx)), udtFormatAlert.PatData, Len(Trim(Arr(Idx))), True) Then
                                    udtFormatAlert.PatData = Trim(udtFormatAlert.PatData) & vbNewLine & Trim(Arr(Idx))
                                End If
                            End If
                        Next Idx
                    End If
                    Add_Row
                    lngRow = grdOrders.MaxRows
                    Format_CellTypeEdit grdOrders, lngRow, 1, "1", udtFormatAlert.Alert, "TypeHAlignLeft"
                    grdOrders.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Format_CellTypeEdit grdOrders, lngRow, 2, "1", udtFormatAlert.PatData & vbNewLine & "ACTION:" & " " & udtFormatAlert.AlertDesc, "TypeHAlignLeft"
                    grdOrders.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    If TextFound("No Order", udtFormatAlert.PatData, Len("No Order"), True) Or _
                        TextFound("No BP", udtFormatAlert.PatData, Len("No BP"), True) Or _
                        TextFound("No BP", udtFormatAlert.PatData, Len("No BP"), True) Or _
                        lngDateDiff >= Val(udtFormatAlert.Time) Then
                        grdOrders.BlockMode = True
                            grdOrders.Col = 1
                            grdOrders.Col2 = 2
                            grdOrders.Row = lngRow
                            grdOrders.Row2 = lngRow
                            grdOrders.BackColor = grdCurItem
                        grdOrders.BlockMode = False
                    End If
                End If
            End If
            strAlert = udtFormatAlert.Alert
        Loop
    End With
    If cnReport.State = adStateOpen Then
        cnReport.Close
        Set cnReport = Nothing
        Set rsReport = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Read_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Data", App.EXEName & "." & TypeName(Me)
    Exit Sub
    Resume
End Sub

Private Function Get_InPatient_Data()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Arr
Dim Idx%
Dim strApptDate As String

    On Error GoTo Error_Get_InPatient_Data
    Screen.MousePointer = vbHourglass
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "Get_InPatient_Data_PatientID " & piPatientID & ", " & clsEncounter.ProviderID & ", " _
        & piFacilityID & ", '" & Format(clsEncounter.StartDate, "mm/dd/yyyy") & "'"
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        If Not .EOF Then
            Clear_InSched
            udtInPatSched.SchedType = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PATIENT_ID"
                                udtInPatSched.IntId = Trim(Fld.Value)
                            Case "PURPOSE_VISIT"
                                udtInPatSched.VisitType = Trim(Fld.Value)
                            Case "APPT_TYPE"
                                udtInPatSched.SchedType = Trim(Fld.Value)
                            Case "STATUS"
                                udtInPatSched.Status = Trim(Fld.Value)
                            Case "PROVIDER_ID"
                                udtInPatSched.Provider_ID = Fld.Value
                            Case "APPT_DATE"
                                udtInPatSched.Appt_Date = Fld.Value
                            Case "FACILITY_ID"
                                udtInPatSched.Facility_ID = Fld.Value
                            Case "RESOURCE_TYPE"
                                udtInPatSched.ResourceType = Trim(Fld.Value)
                            Case "ROOM_NUMBER"
                                udtInPatSched.RoomNumber = Trim(Fld.Value)
                            Case "TELEPHONE"
                                udtInPatSched.RoomTelephone = Trim(Fld.Value)
                            Case "DURATION"
                                udtInPatSched.Duration = Trim(Fld.Value)
                            Case "PATIENT_TYPE"
                                udtInPatSched.PatientType = Trim(Fld.Value)
                        End Select
                    End If
                End If
            Next
        End If
    End With
    If Trim(udtInPatSched.VisitType) = "Subsequent Visit" Then udtInPatSched.VisitType = "Subsequent Care"
    If cboPurpose.listcount <= 0 Then Load_Visits
    If Trim(udtInPatSched.VisitType) <> vbNullString Then
        For Idx = 0 To cboPurpose.listcount - 1
            If TextFound(IsProperCase(udtInPatSched.VisitType), cboPurpose.List(Idx), Len(udtInPatSched.VisitType), True) Then
                cboPurpose.Text = Trim(cboPurpose.List(Idx))
                Exit For
            End If
        Next Idx
    End If
    txtRoomNumber.Text = UCase(udtInPatSched.RoomNumber)
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_InPatient_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_InPatient_Data", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Public Sub Load_Reports()
Dim strOCX As String
Dim strControl As String
Dim strRemarks As String
Dim Arr
Dim i%
Dim mbControlPres As Boolean

    On Error GoTo Error_Load_Reports
    Screen.MousePointer = vbHourglass
    Set_ReportType
    piSelPatientID = 0
    If mstrLoadType = "Appointment" Then
        If AppointmentFormat = "InPatient" Then
        piSelPatientID = frmInPatientAppt.mlngPatientID
        Else
            piSelPatientID = frmPatientAppointment.mlngPatientID
        End If
    Else
         piSelPatientID = frmPatients.piPatientID
    End If
    If clsEncounter.FacilityID > 0 Then
        piFacilityID = clsEncounter.FacilityID
    End If
    Validate_Enctr
    piProviderID = clsEncounter.ProviderID
    If EnctrType = vbNullString Then
        mbooNoData = True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Reports:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_Reports", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Function Load_Scales()
Dim i%
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim booLoad As Boolean
Dim strActName As String
Dim lngRow As Long
Dim varCellValue
Dim Arr
Dim strValue As String

    On Error GoTo Error_Load_Scales
    Screen.MousePointer = vbHourglass
   ' Normalize_Scales
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Get_Attr_Valid_Value_SubDom_Sequence 'sd_Psych_Scales'"
    Set rsTerms = cnTerms.Execute(strSQL)
    With rsTerms
        Do Until .EOF
            strActName = vbNullString
            strValue = vbNullString
            If .Fields("ATTRBTE_VALID_VALUE") <> vbNullString Then
                strActName = .Fields("ATTRBTE_VALID_VALUE")
            End If
            .MoveNext
             If TextFound("Geriatric", Replace(strActName, "_", " "), Len("Geriatric"), True) Then
                 strValue = "GERIATRIC DEPRESSION"
                 strActName = "GERIATRIC DEPRESSION"
             ElseIf TextFound("BIPOLAR", Replace(strActName, "_", " "), Len("BIPOLAR"), True) Then
                 strActName = "BIPOLAR"
                 strValue = "BIPOLAR"
             ElseIf TextFound("Hamilton", strActName, Len("Hamilton"), True) Then
                 strValue = "HAMILTON RATING"
                 strActName = "HAMILTON RATING"
             ElseIf TextFound("Clinical Dementia", Replace(strActName, "_", " "), Len("Clinical Dementia"), True) Then
                 strValue = "CLINICAL DEMENTIA SCALE"
                 strActName = "CLINICAL DEMENTIA SCALE"
             ElseIf TextFound("KATZ", Replace(strActName, "_", " "), Len("KATZ"), True) Then
                 strValue = "KATZ INDEX OF INDEPENDENCE IN ADL"
                 strActName = "KATZ INDEX OF INDEPENDENCE IN ADL"
             ElseIf TextFound("GERIATRIC PAIN", Replace(strActName, "_", " "), Len("GERIATRIC PAIN"), True) Then
                 strValue = "GERIATRIC PAIN MEASURE (GPM)"
                 strActName = "GERIATRIC PAIN MEASURE (GPM)"
             ElseIf TextFound("Decision Capacity", Replace(strActName, "_", " "), Len("Decision Capacity"), True) Then
                 strActName = "DECISION CAPACITY"
                 strValue = "DECISION CAPACITY"
             ElseIf TextFound("YMRS", Replace(strActName, "_", " "), Len("YMRS"), True) Then
                 strActName = "YMRS"
                 strValue = "YOUNG MANIC RATING"
            ElseIf TextFound("MDQ", Replace(strActName, "_", " "), Len("MDQ"), True) Then
                 strActName = "MDQ"
                 strValue = "MDQ"
            ElseIf TextFound("PSEUDOBULBAR", Replace(strActName, "_", " "), Len("Pseudobulbar"), True) Then
                 strValue = "PSEUDOBULBAR AFFECT SCALE"
                 strActName = "PSEUDOBULBAR AFFECT SCALE"
             ElseIf TextFound("PHQ-9", Replace(strActName, "_", " "), Len("PHQ-9"), True) Then
                 strActName = "PHQ-9"
                 strValue = "PHQ-9"
             ElseIf TextFound("BPRS", Replace(strActName, "_", " "), Len("BPRS"), True) Then
                 strActName = "BPRS"
                 strValue = "BPRS"
             ElseIf TextFound("RISCI", Replace(strActName, "_", " "), Len("RISCI"), True) Then
                 strActName = "RISCI"
                 strValue = "BARIATRIC SURGERY CLEARANCE"
             ElseIf TextFound("NEUROSTIMULATION", Replace(strActName, "_", " "), Len("NEUROSTIMULATION"), True) Then
                 strActName = "NEUROSTIMULATION"
                 strValue = "NEUROSTIMULATION SURGERY CLEARANCE"
             ElseIf TextFound("YBOCS", Replace(strActName, "_", " "), Len("YBOCS"), True) Then
                 strActName = "YBOCS"
                 strValue = "YALE-BROWN OBSESSIVE COMPULSIVE RATING"
             ElseIf TextFound("PTSD MIL", Replace(strActName, "_", " "), Len("PTSD MIL"), True) Then
                 strActName = "PTSD_MIL"
                 strValue = "POST TRAUMATIC STRESS DISORDER - MILITARY"
             ElseIf TextFound("PTSD CIV", Replace(strActName, "_", " "), Len("PTSD CIV"), True) Then
                 strActName = "PTSD_CIV"
                 strValue = "POST TRAUMATIC STRESS DISORDER - C"
             ElseIf TextFound("ADHD", Replace(strActName, "_", " "), Len("ADHD"), True) Then
                 strActName = "ADHD"
                 strValue = "ATTENTION DEFICIT HYPERACTIVITY DISORDER"
             ElseIf TextFound("MMSE", Replace(strActName, "_", " "), Len("MMSE"), True) Then
                 strActName = "MMSE"
                 strValue = "MINI MENTAL STATUS"
             ElseIf TextFound("AIMS", Replace(strActName, "_", " "), Len("AIMS"), True) Then
                 strActName = "AIMS"
                 strValue = "AIMS"
             ElseIf TextFound("MECH INVENTORY", Replace(strActName, "_", " "), Len("MECH INVENTORY"), True) Then
                 strActName = "MECH INVENTORY"
                 strValue = "MECH INVENTORY"
             ElseIf TextFound("SUBOXONE", Replace(strActName, "_", " "), Len("SUBOXONE"), True) Then
                 strActName = "SUBOXONE"
                 strValue = "SUBOXONE"
             ElseIf TextFound("SUICIDE", Replace(strActName, "_", " "), Len("SUICIDE"), True) Then
                 strActName = "SUICIDE"
                 strValue = "SUICIDE"
             ElseIf TextFound("ADULT DEVELOPMENT", Replace(strActName, "_", " "), Len("ADULT DEVELOPMENT"), True) Then
                 strActName = "ADULT DEVELOPMENTAL HISTORY"
                 strValue = "ADULT DEVELOPMENT"
             ElseIf TextFound("CHILD DEVELOPMENT", Replace(strActName, "_", " "), Len("CHILD DEVELOPMENT"), True) Then
                 strActName = "CHILD DEVELOPMENTAL HISTORY"
                 strValue = "CHILD DEVELOPMENT"
             ElseIf TextFound("BARIATRIC SURGERY SCREENING EVALUTION", Replace(strActName, "_", " "), Len("BARIATRIC SURGERY SCREENING EVALUTION"), True) Then
                 strActName = "BARIATRIC SURGERY SCREENING EVALUTION"
                 strValue = "BARIATRIC SURGERY SCREENING EVALUTION"
             End If
            If Trim(strValue) <> vbNullString And Trim(strActName) <> vbNullString Then
                lngRow = grdSpec.MaxRows + 1
                grdSpec.MaxRows = lngRow
                Format_CellTypeEdit grdSpec, lngRow, 1, "1", Trim(strValue), "TypeHAlighLeft", UCase(strActName)
            End If
        Loop
    End With
    If cnTerms.State = adStateClosed Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Scales:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Scales", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function

Private Function Load_Visits()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSerPlace As String
Dim strPatType As String
Dim strName As String
Dim strCode As String
Dim strProvider As String
Dim strCompare As String
Dim strCPT As String
Dim i%
Dim strTitle As String
Dim strServCode As String

    If mbooByPass Then mbooByPass = False: Exit Function
    On Error GoTo Error_Get_Status_Visit
    Screen.MousePointer = vbHourglass
    cboPurpose.Clear
    If clsEDI_Client.Site_ID = "1999" Then
        If clsEncounter.EncounterType = "Initial InPatient" Then
                If clsEncounter.ServicePlaceCode = "31" Then
                    cboPurpose.AddItem "Initial E&M Care, Brief Visit - 99325"
'                    cboPurpose.ItemData(cboPurpose.NewIndex) = 1005
                    cboPurpose.AddItem "Initial E&M Care, Detailed Visit - 99326"
                    cboPurpose.AddItem "Initial E&M Care, Extended Visit - 99327"
                    cboPurpose.AddItem "Initial E&M Care, Complex Visit - 99328"
                ElseIf clsEncounter.ServicePlaceCode = "13" Then
                    cboPurpose.AddItem "Initial E&M Care, Brief Visit - 99304"
                    cboPurpose.AddItem "Initial E&M Care, Extended Visit - 99305"
                    cboPurpose.AddItem "Initial E&M Care, Complex Visit - 99306"
                End If
        ElseIf clsEncounter.EncounterType = "Subsequent Care" Then
            If clsEncounter.ServicePlaceCode = "31" Then
                cboPurpose.AddItem "Subsequent care, nursing facility E&M, limited - 99334"
                cboPurpose.AddItem "Subsequent E&M Care, Brief - 99307"
                cboPurpose.AddItem "Subsequent E&M Care, Detailed - 99308"
                cboPurpose.AddItem "Subsequent E&M Care, Extended - 99309"
                cboPurpose.AddItem "Subsequent E&M Care, Complex - 99310"
            ElseIf clsEncounter.ServicePlaceCode = "13" Then
                cboPurpose.AddItem "Subsequent care, nursing facility E&M, limited - 99307"
                cboPurpose.AddItem "Subsequent E&M Care, Extended - 99336"
                cboPurpose.AddItem "Subsequent E&M Care, Detailed - 99335"
                cboPurpose.AddItem "Subsequent E&M Care, Brief - 99334"
            End If
        End If
    Else
        If piProviderID = 0 Then piProviderID = mlngLogonUser
        If piProviderID = 0 And clsEncounter.ProviderID = 0 Then
            clsEncounter.ProviderID = mlngLogonUser
            Get_Providers clsEncounter.ProviderID
        ElseIf clsEncounter.ProviderID = 0 Then
            clsEncounter.ProviderID = piProviderID
            Get_Providers piProviderID
        Else
            Get_Providers clsEncounter.ProviderID
        End If
        strCompare = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
        strCompare = Replace(strCompare, ".", "")
        strCompare = Trim(strCompare)
        If Trim(strCompare) = vbNullString Then strCompare = "MD"
        If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
        strSQL = "Get_Appt_Types_Provider '" & strCompare & "'"
        Set rsCodes = cnCodes.Execute(strSQL)
         Do Until rsCodes.EOF
            strName = vbNullString
            strCode = vbNullString
            Set Flds = rsCodes.Fields
            For Each Fld In Flds
                Select Case Fld.Name
                    Case "TYPE_DESC"
                        strName = Fld.Value
                    Case "APPT_CODE"
                        strCode = Fld.Value
                End Select
            Next
            rsCodes.MoveNext
            If Trim(strName) <> vbNullString Then
                cboPurpose.AddItem strName & " - " & strCode
                If IsNumeric(strCode) Then cboPurpose.ItemData(cboPurpose.NewIndex) = strCode
            End If
        Loop
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Status_Visit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Status_Visit", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Normalize_Scales()
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim strSQL As String


    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Delete_Attr_Valid_Value_Valid_Value_Subdom 'sd_Psych_Scales'"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','BARIATRIC SURGERY SCREENING EVALUTION', 1"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales', 'BPRS', 2"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','BIPOLAR',3"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','PHQ-9', 4"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name  'sd_Psych_Scales','MDQ', 5 "
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name  'sd_Psych_Scales','ADHD',6 "
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name  'sd_Psych_Scales','AIMS', 7"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','MMSE', 8 "
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','GERIATRIC DEPRESSION', 9 "
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name  'sd_Psych_Scales','SUICIDE', 10"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name  'sd_Psych_Scales','YMRS', 11"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name  'sd_Psych_Scales','HAMILTON DEPRESSION',12"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name  'sd_Psych_Scales','PANNS', 13"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name  'sd_Psych_Scales','PTSD_CIV', 14"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','PTSD_MIL', 15"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','YBOCS', 16"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','ADULT DEVELOPMENTAL HISTORY', 17"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','CHILD DEVELOPMENTAL HISTORY',18"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','QUIDS-SR', 19"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','SUBOXONE', 20"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','MECH INVENTORY', 21"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name  'sd_Psych_Scales','RISCI', 22"
    Set rsTerms = cnTerms.Execute(strSQL)
    strSQL = "Insert_Attr_Valid_Value_Seq_Name 'sd_Psych_Scales','NEUROSTIMULATION', 23"
    Set rsTerms = cnTerms.Execute(strSQL)
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Scales:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Scales", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function

Private Sub Set_ReportType()
Dim i%
Dim strPatientType As String

    If clsEncounter.ServicePlace <> vbNullString Then
        If clsEncounter.ServicePlace = "Assisting Living Facility" Then clsEncounter.ServicePlace = "Assisted Living Facility"
        If clsEncounter.ServicePlace = "Office" Then
            ccdAdmit.Visible = False
            lblApptDate.Visible = False
            txtRoomNumber.Visible = False
            lblRoomNumber.Visible = False
            If clsEncounter.EncounterType = "New Patient" Then
                 strPatientType = "Office Initial"
            ElseIf clsEncounter.EncounterType = "Established Patient" Then
                 strPatientType = "Office Established"
            End If
            If cboPurpose.listcount <= 0 Then
                Get_Reports
            End If
            For i = 0 To cboPurpose.listcount - 1
                If TextFound(Trim(strPatientType), cboPurpose.List(i), Len(strPatientType), True) Then
                    cboPurpose.Text = cboPurpose.List(i)
                    mlngReportID = cboPurpose.ItemData(i)
                    modSpecific.mlngReportID = mlngReportID
                    Exit For
                End If
            Next i
        ElseIf clsEncounter.ServicePlace = "InPatient" Or clsEncounter.ServicePlace = "Skilled Nursing Facility" Or clsEncounter.ServicePlace = "Assisted Living Facility" Then
            ccdAdmit.Visible = True
            lblApptDate.Visible = True
            txtRoomNumber.Visible = True
            lblRoomNumber.Visible = True
            Get_InPatient_Data
            txtRoomNumber.Text = UCase(udtInPatSched.RoomNumber)
            ccdAdmit.Value = Format(clsEncounter.Admission, "mm/dd/yyyy hh:mm:ss AM/PM")
        ElseIf clsEncounter.EncounterType = "Initial InPatient" Then
             strPatientType = "InPatient Initial"
        ElseIf clsEncounter.EncounterType = "Established InPatient" _
            Or clsEncounter.EncounterType = "Established Patient" Then
             strPatientType = "InPatient Subsequent"
        ElseIf clsEncounter.EncounterType = "Initial InPatient" And _
             TextFound("Consult", clsEncounter.VisitType, Len("consult"), True) Then
             strPatientType = "InPatient Consult Initial"
        ElseIf TextFound("Established", clsEncounter.VisitType, Len("Established"), True) And _
             TextFound("Consult", clsEncounter.VisitType, Len("consult"), True) Then
             strPatientType = "InPatient Consult Subsequent"
        End If
        Load_Visits
       If cboPurpose.listcount > 0 Then
            For i = 0 To cboPurpose.listcount - 1
                If TextFound(Trim(udtInPatSched.VisitType), Trim(cboPurpose.List(i)), Len(udtInPatSched.VisitType), True) Then
                    cboPurpose.Text = cboPurpose.List(i)
                    mlngReportID = cboPurpose.ItemData(i)
                    modSpecific.mlngReportID = mlngReportID
                    Exit For
                End If
            Next i
            If mlngReportID = 0 Then
                If clsEDI_Client.Site_ID = "1999" Then
                    If Trim(clsEncounter.EncounterType) = "Subsequent Care" Then
                        mlngReportID = 1005
                    End If
                    If clsEncounter.EncounterType = "Initial InPatient" Or clsEncounter.EncounterType = "Re-Admission" Then
                        mlngReportID = 1002
                    End If
                End If
            End If
        Else
            Debug.Print
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Function Get_CurrentEnctr()
 
End Function
Private Sub Add_Row()
Dim lngRow As Long
Dim lngCol As Long
Dim Arr
Dim strList As String
    Screen.MousePointer = vbHourglass
    lngRow = grdOrders.MaxRows + 1
    grdOrders.MaxRows = lngRow
    lngCol = 1
    Format_CellTypeCombo grdOrders, lngRow, lngCol, "1", Load_Alerts, "TypeHAlignCenter", "ALERT"
    grdOrders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 2
    Format_CellTypeEdit grdOrders, lngRow, lngCol, "1", "", "TypeHAlignCenter", "REASON_FLAG"
    grdOrders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub Load_Control()
Dim strOCX As String
Dim strControl As String
Dim lstrControl As String
Dim lstrSQL As String
Dim strSQL As String
Dim lbooReadOnly As Boolean
Dim cnGetTerms As New ADODB.Connection
Dim rsGetTerms As New ADODB.Recordset
Dim Idx%

    On Error GoTo Error_Load_Control
    If frmInPatient.mstrLoadType = "PostOpVisit" Then
        If mstrPatOcx = vbNullString Then mstrPatOcx = "FN_Orders"
        If mstrPatCtl = vbNullString Then mstrPatCtl = "ProgNote"
    ElseIf frmInPatient.mstrLoadType = "Surgery" Then
        If mstrPatOcx = vbNullString Then mstrPatOcx = "PatientTrac_Report"
        If mstrPatCtl = vbNullString Then mstrPatCtl = "Procedures"
    ElseIf frmInPatient.mstrLoadType = "Ambulatory Procedure Visit" Then
        If mstrPatOcx = vbNullString Then mstrPatOcx = "FN_Orders"
        If mstrPatCtl = vbNullString Then mstrPatCtl = "OpRoomRecord"
    ElseIf frmInPatient.mstrLoadType = "Pre-Operative Preparation" Then
        If mstrPatOcx = vbNullString Then mstrPatOcx = "FN_Orders"
        If mstrPatCtl = vbNullString Then mstrPatCtl = "Surg_HoldArea"
    ElseIf frmInPatient.mstrLoadType = "Surgical Admission" Then
        If mstrPatOcx = vbNullString Then mstrPatOcx = "PatientTrac_Report"
        If mstrPatCtl = vbNullString Then mstrPatCtl = "SurgAdmit"
    ElseIf frmInPatient.mstrLoadType = "PreOperative Anesthesia" Then
        If mstrPatOcx = vbNullString Then mstrPatOcx = "PatientTrac_Report"
        If mstrPatCtl = vbNullString Then mstrPatCtl = "AnesthQuest"
    Else
        mstrPatOcx = "PatientTrac_Report"
        mstrPatCtl = "Admission"
    End If
    If piSelPatientID = 0 Then piSelPatientID = modSpecific.piPatientID
    If piEncounterID = 0 Then piEncounterID = modSpecific.piEncounterID
    If piSelPatientID = 0 Then piSelPatientID = piPatientID
    If piEncounterID = 0 Then piEncounterID = piEncounterID
    If cnGetTerms.State = adStateClosed Then cnGetTerms.open psConnect(1)
    lstrSQL = "Get_ActiveX_ID '" & mstrPatOcx & "', '" & mstrPatCtl & "'"
    Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
    With rsGetTerms
        If Not .EOF Then
            If Not IsNull(!actx_id) Then
                mintActXId = (!actx_id)
            End If
        End If
    End With
    If mintActXId = 0 Then
        strSQL = "Insert_New_ActiveX '" & mstrPatOcx & "', '" & mstrPatCtl & "', 'Components', '" & mstrPatOcx & "', '2' "
        Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
        strSQL = lstrSQL = "Get_ActiveX_ID '" & mstrPatOcx & "', '" & mstrPatCtl & "'"
        Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
        With rsGetTerms
            If Not .EOF Then
                If Not IsNull(!actx_id) Then
                    mintActXId = (!actx_id)
                End If
            End If
        End With
    End If
    If cnGetTerms.State = adStateOpen Then
        cnGetTerms.Close
        Set cnGetTerms = Nothing
        Set rsGetTerms = Nothing
    End If
    mstrMyControl = "ctl" & mstrPatCtl
    Set fobjOCX = Nothing
    If Licenses.Count > 0 Then
        For Idx = 0 To Licenses.Count
            If Licenses.Count > 0 Then
                Licenses.Remove (Idx)
            Else
                Exit For
            End If
        Next Idx
        Licenses.Clear
    End If
    Licenses.Add mstrPatOcx & "." & mstrPatCtl
    Me.Controls.Remove mstrMyControl
    Set fobjOCX = Controls.Add(mstrPatOcx & "." & mstrPatCtl, "ctl" & mstrPatCtl, Me)
        fobjOCX.Visible = True
        fobjOCX.Left = grdPatient.Width + 2
        fobjOCX.Top = cboFacility.Top
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = mintActXId
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = piSelPatientID
        fobjLib.EncounterID = piEncounterID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
    Exit Sub
Error_Load_Control:
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Or Err.Number = 730 Then Resume Next
    If Err.Number = 731 Or Err.Number = 747 Then
        If Licenses.Count > 0 Then Licenses.Clear
        Licenses.Add strOCX & "." & strControl
    Else
        Screen.MousePointer = vbDefault
        Write_Error Err.Description, Err.Number, strSQL, "Load_Control", App.EXEName & "." & Me.Name
        Exit Sub
        Resume
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
End Sub
Private Function Set_Visit()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSerPlace As String
Dim strPatType As String
Dim strName As String
Dim strCode As String
Dim strProvider As String
Dim strCompare As String
Dim strCPT As String
Dim i%
Dim strTitle As String
Dim strServCode As String
Dim Arr
Dim strSearchTitle As String

    If mbooByPass Then mbooByPass = False: Exit Function
    On Error GoTo Error_Set_Visit
    Screen.MousePointer = vbHourglass
    cboPurpose.Clear
    If Trim(clsProvider.Title) <> vbNullString Then
        strSearchTitle = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
        strSearchTitle = Replace(strSearchTitle, ".", "")
        strSearchTitle = Trim(strSearchTitle)
    Else
        If piProviderID = 0 Then
            Get_DefaultProvider
            clsEncounter.ProviderID = piProviderID
            Get_Providers clsEncounter.ProviderID
        Else
            Get_Providers piProviderID
            clsEncounter.ProviderID = piProviderID
            
        End If
        strSearchTitle = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
        strSearchTitle = Replace(strSearchTitle, ".", "")
        strSearchTitle = Trim(strSearchTitle)
    End If
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    strServCode = clsEncounter.ServicePlaceCode
    strSQL = "Get_CPT_Template_By_Type_ServPl_EnctrType " & clsEncounter.ServicePlaceCode & ",'" _
        & clsEncounter.EncounterType & "%', 'Miscellaneous'"
    Set rsCodes = cnCodes.Execute(strSQL)
     Do Until rsCodes.EOF
        strName = vbNullString
        strCode = vbNullString
        strProvider = vbNullString
        If rsCodes.Fields("PROVIDER") <> vbNullString Then
            strProvider = UCase(rsCodes.Fields("PROVIDER"))
            strProvider = Replace(strProvider, ".", "")
            strProvider = Trim(strProvider)
            Set Flds = rsCodes.Fields
            For Each Fld In Flds
                Select Case strProvider
                    Case strSearchTitle
                        If Not IsNull(Fld.Value) Then
                            If Fld.Value <> vbNullString Then
                                If Fld.Name = "CPT_NAME" Then
                                    strName = Fld.Value
                                ElseIf Fld.Name = "CPT_CODE" Then
                                    strCode = Fld.Value
                                End If
                            End If
                        End If
                End Select
            Next
        End If
        rsCodes.MoveNext
        If Trim(strName) <> vbNullString Then
            cboPurpose.AddItem strName & " - " & strCode
            If IsNumeric(strCode) Then cboPurpose.ItemData(cboPurpose.NewIndex) = strCode
        End If
    Loop
    If cboPurpose.listcount <= 0 Then
        strPatType = Replace(strPatType, "Patient", "")
        strPatType = Replace(strPatType, "Office", "")
        strSQL = "Get_CPT_Template_By_Type_ServPl_EnctrType " & clsEncounter.ServicePlaceCode & ", '" _
            & clsEncounter.EncounterType & "%', 'Miscellaneous'"
        Set rsCodes = cnCodes.Execute(strSQL)
         Do Until rsCodes.EOF
            strName = vbNullString
            strCode = vbNullString
            strProvider = vbNullString
            If rsCodes.Fields("PROVIDER") <> vbNullString Then
                strProvider = UCase(rsCodes.Fields("PROVIDER"))
                strProvider = Replace(strProvider, ".", "")
                strProvider = Trim(strProvider)
                Set Flds = rsCodes.Fields
                For Each Fld In Flds
                Select Case strProvider
                    Case strSearchTitle
                        If Not IsNull(Fld.Value) Then
                            If Fld.Value <> vbNullString Then
                                If Fld.Name = "CPT_NAME" Then
                                    strName = Fld.Value
                                ElseIf Fld.Name = "CPT_CODE" Then
                                    strCode = Fld.Value
                                End If
                            End If
                        End If
                    End Select
                Next
            End If
            rsCodes.MoveNext
            If Trim(strName) <> vbNullString Then
                cboPurpose.AddItem strName & " - " & strCode
                If IsNumeric(strCode) Then cboPurpose.ItemData(cboPurpose.NewIndex) = strCode
            End If
        Loop
    End If
    For i = 0 To cboPurpose.listcount - 1
        If cboPurpose.List(i) = clsEncounter.VisitType Then
            cboPurpose.Text = cboPurpose.List(i)
            Exit For
        End If
    Next i
    If clsEDI_Client.Site_ID = "1999" Then
        If cboPurpose.Text = vbNullString And clsEncounter.EncounterType = "Subsequent Care" Then
            For i = 0 To cboPurpose.listcount - 1
                Arr = Split(Trim(cboPurpose.List(i)), "-")
                If UBound(Arr) > 0 Then
                    strCode = Arr(1)
                    If strCode <> vbNullString Then
                        If Trim(strCode) = Left(clsEncounter.Intrvn_CD, 5) Then
                            cboPurpose.Text = cboPurpose.List(i)
                            Exit For
                        End If
                    End If
                End If
            Next i
            If cboPurpose.Text = vbNullString Then
                If clsEncounter.ServicePlaceCode = "31" Then
                    cboPurpose.AddItem "Subsequent care, nursing facility E&M, limited - 99334"
                    cboPurpose.AddItem "Subsequent E&M Care, Brief - 99307"
                    cboPurpose.AddItem "Subsequent E&M Care, Detailed - 99308"
                    cboPurpose.AddItem "Subsequent E&M Care, Extended - 99309"
                    cboPurpose.AddItem "Subsequent E&M Care, Complex - 99310"
                ElseIf clsEncounter.ServicePlaceCode = "13" Then
                    cboPurpose.AddItem "Subsequent care, nursing facility E&M, limited - 99307"
                    cboPurpose.AddItem "Subsequent E&M Care, Extended - 99336"
                    cboPurpose.AddItem "Subsequent E&M Care, Detailed - 99335"
                    cboPurpose.AddItem "Subsequent E&M Care, Brief - 99334"
                End If
            End If
        ElseIf cboPurpose.Text = vbNullString And clsEncounter.EncounterType = "Initial InPatient" Then
            For i = 0 To cboPurpose.listcount - 1
                Arr = Split(Trim(cboPurpose.List(i)), "-")
                If UBound(Arr) > 0 Then
                    strCode = Arr(1)
                    If strCode <> vbNullString Then
                        If Trim(strCode) = Left(clsEncounter.Intrvn_CD, 5) Then
                            cboPurpose.Text = cboPurpose.List(i)
                            Exit For
                        End If
                    End If
                End If
            Next i
            If cboPurpose.Text = vbNullString Then
                If clsEncounter.ServicePlaceCode = "31" Then
                    cboPurpose.AddItem "Initial E&M Care, Brief Visit - 99325"
                    cboPurpose.AddItem "Initial E&M Care, Detailed Visit - 99326"
                    cboPurpose.AddItem "Initial E&M Care, Extended Visit - 99327"
                    cboPurpose.AddItem "Initial E&M Care, Complex Visit - 99328"
                ElseIf clsEncounter.ServicePlaceCode = "13" Then
                    cboPurpose.AddItem "Initial E&M Care, Brief Visit - 99304"
                    cboPurpose.AddItem "Initial E&M Care, Extended Visit - 99305"
                    cboPurpose.AddItem "Initial E&M Care, Complex Visit - 99306"
                End If
                For i = 0 To cboPurpose.listcount - 1
                    Arr = Split(Trim(cboPurpose.List(i)), "-")
                    If UBound(Arr) > 0 Then
                        strCode = Arr(1)
                        If strCode <> vbNullString Then
                            If Trim(strCode) = Left(clsEncounter.Intrvn_CD, 5) Then
                                cboPurpose.Text = cboPurpose.List(i)
                                Exit For
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_Visit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Set_Visit", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub Update_RoomInfo()
Dim strSQL As String
Dim strRoom As String

    If Trim(txtRoomNumber.Text) <> vbNullString Then
        strRoom = Literals(txtRoomNumber.Text, "'")
    End If
    Open_DB psConnect(0)
    strSQL = "Update_Room_Number " & strRoom & "', " & mlngPatientID
    Set rs = cn.Execute(strSQL)
    Close_DB psConnect(0)
    Exit Sub
End Sub

Private Sub cboFacility_Click()
Dim lngFacilityID As Long
Dim i%

    On Error GoTo Error_cboFacility_Click

    For i = 0 To cboFacility.listcount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            mlngSelFacID = cboFacility.ItemData(i)
            Exit For
        End If
    Next i
    Exit Sub
Error_cboFacility_Click:
     Write_Error Err.Description, Err.Number, "", "cboFacility_Click", App.EXEName & "." & Me.Name
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub cboFacility_KeyUp(KeyCode As Integer, Shift As Integer)
    SearchList cboFacility, cboFacility.Text, True
End Sub


Public Sub Read_Data(vlngEncounterID As Long)
Dim i%
    
    cmdGrdView.Caption = "View Encounters"
    cmdGrdView_Click
    Load_Facilities
    Format_ActionGrid
    Get_ActionItems
    Load_Visits
    If vlngEncounterID = 0 Then vlngEncounterID = piEncounterID
        Dump_PatientStrData
        Dump_EncounterStrData
        piEncounterID = vlngEncounterID
        If mstrLoadType = "Appointment" Then
            If AppointmentFormat = "InPatient" And frmInPatientAppt.mlngPatientID <> 0 Then
                piSelPatientID = frmInPatientAppt.mlngPatientID
            Else
                piSelPatientID = frmPatientAppointment.mlngPatientID
            End If
        Else
             piSelPatientID = frmPatients.piPatientID
        End If
        If piEncounterID <> 0 Then modSpecific.piEncounterID = piEncounterID
        If piSelPatientID <> 0 Then modSpecific.piPatientID = piSelPatientID
        piPatientID = piSelPatientID
        Get_Patient piSelPatientID
        Me.Caption = "Current Patient" & " " & UCase(clsPatient.Name) & " " & clsPatient.IntId & " " & clsPatient.ExtID
        Get_Encounter piEncounterID
        If clsEncounter.ProviderID <> 0 Then
            piProviderID = clsEncounter.ProviderID
        Else
            clsEncounter.ProviderID = piProviderID
        End If
        Load_Control
        For i = 0 To cboFacility.listcount - 1
            If cboFacility.List(i) = clsEncounter.Facility Then
                cboFacility.Text = cboFacility.List(i)
                piFacilityID = cboFacility.ItemData(i)
                Exit For
            End If
        Next i
    Exit Sub
End Sub

Private Sub Get_Reports()
Dim i%
Dim strSQL As String
Dim strPatType As String

    If cboPurpose.Text <> vbNullString Then
        strPatType = cboPurpose.Text
    Else
        strPatType = clsEncounter.VisitType
    End If
    Open_DB psConnect(1)
    cboPurpose.Clear
     strSQL = "Get_Reports_Practice '" & mstrPractice & "'"
     Set rs = cn.Execute(strSQL)
     Do Until rs.EOF
        If Not IsNull(rs.Fields("PATIENT_TYPE")) Then
            cboPurpose.AddItem rs.Fields("PATIENT_TYPE")
            If IsNumeric(rs.Fields("REPORT_ID")) Then cboPurpose.ItemData(cboPurpose.NewIndex) = rs.Fields("REPORT_ID")
        End If
     rs.MoveNext
     Loop
    Close_DB psConnect(1)
    For i = 0 To cboPurpose.listcount - 1
        If cboPurpose.List(i) = strPatType Then
            cboPurpose.Text = cboPurpose.List(i)
            Exit For
        End If
    Next i
    Exit Sub
End Sub
Private Function Validate_Enctr()
Dim cnEnctr As New ADODB.Connection
Dim rsEnctr As New ADODB.Recordset
Dim strSQL As String

    On Error GoTo Error_Validate_Enctr
    Screen.MousePointer = vbHourglass
    If cnEnctr.State = adStateClosed Then cnEnctr.open psConnect(0)
    strSQL = "Get_Encntr_Date_PatientID_Enctr_ID " & piSelPatientID & ", " & piSelEnctrID
     Set rsEnctr = cnEnctr.Execute(strSQL)
     If Not rsEnctr.EOF Then
        If rsEnctr.Fields("ENCNTR_START_DATE") = vbNullString Then
           piSelEnctrID = piEncounterID
        End If
     End If
     If cnEnctr.State = adStateOpen Then
        cnEnctr.Close
        Set cnEnctr = Nothing
        Set rsEnctr = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_Enctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_Enctr", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub ccdAdmit_Click()
Dim cnAdmit As New ADODB.Connection
Dim rsAdmit As New ADODB.Recordset
Dim sDisDate As String
Dim sAdmitDate As String
Dim strEnctrDate As String
Dim lngDateDiff As String
Dim booAdmit As Boolean
Dim booDischarge As Boolean
Dim sDate As String
Dim lngAdmitID As Long
Dim strSQL As String

    If cnAdmit.State = adStateClosed Then cnAdmit.open psConnect(0)
    booAdmit = True
    booDischarge = True
    strSQL = "SELECT * FROM PATIENT_ADMISSION  WHERE PATIENT_ID = " & piSelPatientID & " AND PURGE = 0 ORDER BY ADMISSION_DATE DESC"
    Set rsAdmit = cnAdmit.Execute(strSQL)
    With rsAdmit
        If Not .EOF Then
            If Not IsNull(!PATIENT_ADMITID) Then lngAdmitID = (!PATIENT_ADMITID)
            sAdmitDate = Format(ccdAdmit.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
            Open_DB psConnect(0)
            strSQL = "UPDATE PATIENT_ADMISSION SET ADMISSION_DATE = '" & sAdmitDate & "'," _
            & "UPDATE_USER = '" & PsUserID & "', PURGE = 0, UPDATE = '" & sDate & "' WHERE PATIENT_ADMITID = " & lngAdmitID
            Set rs = cn.Execute(strSQL)
        End If
    End With
    clsEncounter.Admission = Format(ccdAdmit.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
    strSQL = "Update encounter set purge = 0, ADMISSION_DATE = '" & clsEncounter.Admission & "'  where ENCOUNTER_ID = " & piEncounterID
    Set rs = cn.Execute(strSQL)
    Close_DB psConnect(0)
    If cnAdmit.State = adStateOpen Then
        cnAdmit.Close
        Set cnAdmit = Nothing
        If rsAdmit.State = adStateOpen Then
            rsAdmit.Close
            Set rsAdmit = Nothing
        End If
    End If
    Exit Sub
End Sub





Private Sub Load_Patients()
Dim Arr
Dim iCol As Long
Dim iRow As Long
Dim Idx As Integer
Dim vTime
Dim vdata
Dim varData
Dim sFirstName As String
Dim sLastName As String
Dim strSuffix As String
Dim strTitle As String
Dim strPatient As String
Dim strReason As String
Dim booNoInsert As Boolean
Dim lngProviderID As Long
Dim lngProvCol As Long
Dim strSQL As String
Dim varCellValue
Dim strSelDate As String

    On Error GoTo Error_Load_Patients

    Open_DB psConnect(0)
    strSelDate = Format(Now, "mm/dd/yyyy")
    If mlngSelFacID = 0 Then mlngSelFacID = piFacilityID
        strSQL = "SELECT  b.first_name, b.last_name, b.Patient_ID, b.ROOM_NUMBER, a.Reason, b.Patient_Ext_Ref " _
                & "FROM APPOINTMENTS a, PATIENT b " _
                & "WHERE APPT_DATE = '" & strSelDate & "' AND FACILITY_ID = " & mlngSelFacID & " AND " _
                & "b.PATIENT_ID = a.PATIENT_ID AND a.ACTIVE = 1  " _
                & "Order by Appt_time"
        Set rs = cn.Execute(strSQL)
    With rs
       Do Until .EOF
        iRow = iRow + 1
        grdPatient.MaxRows = iRow
        If (!Last_Name) <> vbNullString Then
            If TextFound("''", (!Last_Name), 2, True) Then
                strPatient = Replace((!Last_Name), "''", "'")
            Else
                strPatient = (!Last_Name)
            End If
        End If
        If Trim((!First_Name)) <> vbNullString Then strPatient = Trim(strPatient) & ", " & Trim((!First_Name))
        If Trim((!Reason)) <> vbNullString Then strReason = Trim((!Reason))
        grdPatient.Redraw = False
             grdPatient.BlockMode = True
                grdPatient.Col = iCol
                grdPatient.Col2 = iCol
                grdPatient.Row = iRow
                grdPatient.Row2 = iRow
                grdPatient.TypeEditMultiLine = True   'Make multi-line
                grdPatient.TypeMaxEditLen = 20000
                If Trim(sLastName) <> vbNullString Then strPatient = Trim(sLastName)
                grdPatient.SetText 1, iRow, Trim(strPatient)
                If Not IsNull(!Patient_ID) Then grdPatient.SetRowItemData (iRow), Val(!Patient_ID)
                grdPatient.SetText 2, iRow, Trim(strReason)
                If Not IsNull(!ROOM_NUMBER) Then grdPatient.SetText 3, iRow, (!ROOM_NUMBER)
                grdPatient.TypeHAlign = TypeHAlignCenter
                grdPatient.TypeVAlign = TypeVAlignCenter
                grdPatient.SetCellBorder 1, iRow, grdPatient.MaxCols, iRow, 16, &H80&, CellBorderStyleSolid
                grdPatient.SetCellBorder iCol, iRow, iCol + 1, iRow, 1, &H80&, CellBorderStyleSolid
            grdPatient.BlockMode = False
                sFirstName = vbNullString
                sLastName = vbNullString
                strSuffix = vbNullString
                strTitle = vbNullString
                strPatient = vbNullString
                strReason = vbNullString
        .MoveNext
       Loop
    End With
    Close_DB psConnect(0)
   mbooByPass = False
    Exit Sub
    
Error_Load_Patients:
    MsgBox "Contact PatientTrac Technical Support" & " " & Err.Description
     Exit Sub
     Resume
End Sub
Private Sub Load_Enctrs()
Dim Arr
Dim lngCol As Long
Dim lngRow As Long
Dim Idx As Integer
Dim vTime
Dim vdata
Dim varData
Dim sFirstName As String
Dim sLastName As String
Dim strSuffix As String
Dim strTitle As String
Dim strPatient As String
Dim strReason As String
Dim booNoInsert As Boolean
Dim lngProviderID As Long
Dim lngProvCol As Long
Dim strSQL As String
Dim varCellValue
Dim strSelDate As String

    On Error GoTo Error_Load_Enctrs
    Screen.MousePointer = vbHourglass
    If piSelPatientID = 0 Then piSelPatientID = piPatientID
    Open_DB psConnect(0)
    strSelDate = Format(Now, "mm/dd/yyyy")
    grdPatient.ClearRange 1, 1, grdPatient.MaxCols, grdPatient.MaxRows, False
    lngRow = 0
    If mlngSelFacID = 0 Then mlngSelFacID = piFacilityID
        strSQL = "Get_Encounter_Patient_ID " & piSelPatientID
        Set rs = cn.Execute(strSQL)
    With rs
       Do Until .EOF
            lngRow = lngRow + 1
            grdPatient.MaxRows = lngRow
            grdPatient.Redraw = False
            If Not IsNull((!ENCNTR_START_DATE)) Then
                grdPatient.SetRowItemData lngRow, (!Encounter_ID)
                Format_CellTypeEdit grdPatient, lngRow, 1, "1", Format((!ENCNTR_START_DATE), "MMMM DD, YYYY"), "TypeHAlignLeft"
                grdPatient.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
            If Not IsNull((!VISIT_REASON)) Then
                Format_CellTypeEdit grdPatient, lngRow, 2, "1", (!VISIT_REASON), "TypeHAlignLeft"
                grdPatient.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
            If Not IsNull((!Visit_Type)) Then
                Format_CellTypeEdit grdPatient, lngRow, 3, "1", (!Visit_Type), "TypeHAlignLeft"
                grdPatient.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        .MoveNext
       Loop
    End With
    Close_DB psConnect(0)
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Enctrs:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_Enctrs", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Load_Facilities()
Dim strSQL As String
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset
Dim strName As String

    On Error GoTo Error_Load_Facilities
    Screen.MousePointer = vbHourglass
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
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set cnFac = Nothing
        Set rsFac = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Facilities:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_Facilities", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub cmdAlert_Click()
    mstrPatOcx = "Reporter"
    mstrPatCtl = "PatientAlert"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub cmdAppt_Click()
'    Me.Hide
'    frmPatientAppointment.Show
'    Exit Sub
End Sub

Private Sub cmdCloseList_Click()
    Unload Me
'    frmMDI.Get_AppointmentFormat
End Sub

Private Sub cmdGrdView_Click()

    If mbooByPass Then mbooByPass = False: Exit Sub
    If cmdGrdView.Caption = "View Encounters" Then
        cmdGrdView.ToggleValue = False
        cmdGrdView.Caption = "View Current Patients"
        mstrGridName = "Encntr"
        Format_EncntrGrid
    ElseIf cmdGrdView.Caption = "View Current Patients" Then
        cmdGrdView.ToggleValue = True
        cmdGrdView.Caption = "View Encounters"
        Format_PatientGrid
        mstrGridName = "Today's Patient List"
    End If
    Exit Sub
End Sub

Private Sub cmdNewPat_Click()
    Me.Caption = vbNullString
    modSpecific.piPatientID = 0
    modSpecific.piEncounterID = 0
    piSelPatientID = 0
    piEncounterID = 0
    modSpecific.piProviderID = piProviderID
    Dump_PatientStrData
    PatientInfo.mNewPatient = True
    PatientInfo.Show vbModal
    If PatientInfo.piPatientID <> 0 Then
        modSpecific.piPatientID = 0
        Get_Patient PatientInfo.piPatientID
        frmPatients.piPatientID = PatientInfo.piPatientID
        piSelPatientID = PatientInfo.piPatientID
    End If
    frmEnctr.piPatientID = piSelPatientID
    frmEnctr.Show
    Exit Sub
End Sub

Private Sub cmdRefresh_Click()
        fobjLib.ComponentID = mintActXId
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = piSelPatientID
        fobjLib.EncounterID = piEncounterID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        Me.ZOrder 0
    Exit Sub
End Sub

Private Sub cmdReportBldr_Click()
    fraButton.Visible = False
End Sub

Private Sub cmdReset_Click()
    fraButton.Visible = True
End Sub


Private Sub cmdSave_Click()
    fobjLib.ComponentID = mintActXId
    fobjLib.ConnectStringTerms = psConnect(1)
    fobjLib.ConnectStringClinical = psConnect(0)
    fobjLib.PatientID = piSelPatientID
    fobjLib.EncounterID = piEncounterID
    fobjLib.ProviderID = piProviderID
    fobjLib.WriteClinicalData
    MsgBox "Successfully saved data!", vbInformation
    Exit Sub
End Sub
Private Sub Format_ActionGrid()

    With grdOrders
        .Redraw = False
        .MaxCols = 2
        .MaxRows = 0
        .ClearRange 0, 0, grdPatient.MaxCols, grdPatient.MaxRows, False
        .ColWidth(1) = 1500
        .ColWidth(2) = 2500
         Format_CellTypeEdit grdOrders, SpreadHeader, 1, "2", "PATIENT ALERTS", "TypeHAlignCenter", "TITLE", NormalColor
         grdOrders.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
         Add_Spacer SpreadHeader + 1, grdOrders
         Format_CellTypeEdit grdOrders, SpreadHeader + 2, 1, "1", "ALERT", "TypeHAlignCenter", "ALERT", HeaderColor
         grdOrders.SetCellBorder 1, SpreadHeader + 2, 1, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
         Format_CellTypeEdit grdOrders, SpreadHeader + 2, 2, "1", "PATIENT DATA", "TypeHAlignCenter", "REASON_FLAG", HeaderColor
         grdOrders.SetCellBorder 2, SpreadHeader + 2, 2, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
         Add_Spacer SpreadHeader + 3, grdOrders
        .Redraw = True
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
              .TypeVAlign = TypeVAlignCenter
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
                If vstrHAligh = "TypeHAlignCenter" Then
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .TypeHAlign = TypeHAlignLeft
                End If
              If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
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
Private Function Set_GridCurSel(vObjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, Optional vstrCellTag As String)
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
            If TextFound(Trim(vstrData), Trim(.TypeComboBoxString), Len(Trim(vstrData)), True) Then
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
    .Redraw = True
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
            .RowHeight(vlngRow) = 10
            .BackColor = 12632256
            .AddCellSpan 0, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub
Private Function Get_Add_Notes()
Dim strSQL As String
Dim cnRead As New ADODB.Connection
Dim rsRead As New ADODB.Recordset

    On Error GoTo Error_Get_Add_Notes
    Screen.MousePointer = vbHourglass
    cctPatNote.Text = vbNullString
    If cnRead.State = adStateClosed Then cnRead.open psConnect(0)
    strSQL = "Get_Additional_Notes_PatientID " & piSelPatientID
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

Private Sub cmdStatus_Click()
    frmComment.Show vbModal
    Get_Add_Notes
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
Dim strEnctrType As String
Dim strSQL As String

    MsgBox "Please Correct Data in PatientTrac Encounter, Current Encounter is shown on task menu below. Thank You!", vbInformation
    frmEnctr.Show
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()

    On Error GoTo Error
    Me.ZOrder 0
    centerForm Me
    mFirst = False
    Exit Sub
Error:
    Write_Error Err.Description, Err.Number, "", "Form_Activate", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Form_Load()
    mFirst = True
    
'    If IsFormLoaded("frmPatientAppointment") Then
'        Unload frmPatientAppointment
'    End If
    cctPatNote.Text = vbNullString
    Read_Data piEncounterID
    mFirst = False
    Screen.MousePointer = vbDefault
    If piSelPatientID = 0 Then piSelPatientID = piPatientID
    If piSelEnctrID = 0 Then piSelEnctrID = piEncounterID
    If piSelPatientID = 0 Or piSelEnctrID = 0 Then
         MsgBox "Unable to load Reports, Please select a Patient and Encounter!", vbInformation
         Exit Sub
    End If
    If piSelEnctrID > 0 Then Get_Encounter piSelEnctrID
    If piSelPatientID > 0 And piSelEnctrID > 0 Then Load_Reports
    Get_Patient piSelPatientID
    lblPatient.Caption = lblPatient.Caption & " " & UCase(clsPatient.FirstName) & " " & UCase(clsPatient.LastName)
    Get_Add_Notes
    If piSelEnctrID > 0 Then Get_Encounter piSelEnctrID
    Set_Visit
    If mstrPractice = "Psychiatry" Then Format_SpecGrid
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i%

  On Error Resume Next
    mbooNoData = False
    If Licenses.Count >= 0 Then
        If Licenses.Item(0) <> vbNullString Then
            If Licenses.Item(0) <> vbNullString Then
                Licenses.Remove (0)
            End If
        End If
        Licenses.Clear
    End If
    Set fobjLib = Nothing
    Set fobjOCX = Nothing
End Sub

Private Sub Form_Resize()
    Me.ZOrder 0
    Exit Sub
End Sub

Private Sub Form_Terminate()
    Debug.Print
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Unload Me
'    frmMDI.Get_AppointmentFormat
End Sub

Private Sub grdOrders_DblClick(ByVal Col As Long, ByVal Row As Long)

    mstrPatOcx = "Reporter"
    mstrPatCtl = "PatientAlert"
    Me.Hide
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Me.Show
    Exit Sub
End Sub


Private Sub grdPatient_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngGrid As Long

    On Error GoTo Error_Load

    If Row < 0 Then
        MsgBox "No Patient History to show!", vbInformation, "PatientTrac Progress Note"
        Exit Sub
    End If
    lngGrid = grdPatient.GetRowItemData(Row)
    frmPreview.mlngGrdEncr = lngGrid
    frmPreview.piSelPatientID = 0
    frmPreview.piSelPatientID = piSelPatientID
    frmPreview.mstrLoad_Type = "Patient Record"
    Me.Hide
    Load frmPreview
    frmPreview.Show vbModal
    Me.Show
    Exit Sub
Error_Load:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load", App.EXEName & "." & Me.Name
    Resume Next
    Exit Sub
End Sub


Private Sub grdSpec_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim lngCol As Long

    On Error GoTo Error_Rating_Scale
    With grdSpec
        lngRow = Row
        lngCol = Col
        .Col = lngCol
        .Row = lngRow
        If .CellTag <> vbNullString Then
            Me.Hide
            frmLoadControl.piSelPatientID = piSelPatientID
            Read_Scales .CellTag
            fobjLib.ReadClinicalData
            Me.Show
            Me.Refresh
        End If
    End With
    Exit Sub
Error_Rating_Scale:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load", App.EXEName & "." & Me.Name
    Resume Next
    Exit Sub
End Sub


Private Sub mnu_ProgressNote_Click()
'    mstrPatOcx = "PatientTrac_Report"
'    mstrPatCtl = "Progress_Template"
'    Me.Hide
'    Load frmLoadControl
'    frmLoadControl.Show vbModal
'    Me.Show
'    Exit Sub
End Sub


