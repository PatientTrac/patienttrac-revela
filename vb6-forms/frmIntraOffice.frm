VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmIntraOffice 
   Caption         =   "Practice Alerts"
   ClientHeight    =   9060
   ClientLeft      =   6990
   ClientTop       =   3270
   ClientWidth     =   13680
   Icon            =   "frmIntraOffice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   13680
   Begin VB.Frame fraMODI 
      Height          =   8985
      Left            =   870
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   12195
      Begin VB.TextBox txtCurPage 
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
         Height          =   405
         Left            =   3630
         TabIndex        =   11
         Top             =   8430
         Width           =   555
      End
      Begin VB.TextBox txtTotalPage 
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
         Left            =   1290
         TabIndex        =   10
         Top             =   8430
         Width           =   645
      End
      Begin PT_XP_Button.PT_XP cmdRotate 
         Height          =   375
         Left            =   3330
         TabIndex        =   2
         Top             =   180
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
         Picture         =   "frmIntraOffice.frx":0ECA
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Rotate"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":1464
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrintRecords 
         Height          =   375
         Left            =   10890
         TabIndex        =   3
         ToolTipText     =   "Print Images"
         Top             =   180
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "frmIntraOffice.frx":17B6
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   15790320
         Caption         =   "Print"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":1D50
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrev 
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "frmIntraOffice.frx":20A2
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Previous "
         CaptionAlignment=   5
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":243C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSingleView 
         Height          =   375
         Left            =   1710
         TabIndex        =   5
         Top             =   180
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
         Picture         =   "frmIntraOffice.frx":278E
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Single Page"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":2D28
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdMultiView 
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   180
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
         Picture         =   "frmIntraOffice.frx":307A
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Multi Page"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":3614
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   375
         Left            =   5730
         TabIndex        =   7
         Top             =   8430
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmIntraOffice.frx":3966
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":4378
         ShowFocus       =   -1  'True
      End
      Begin RichTextLib.RichTextBox rtfScan 
         Height          =   7725
         Left            =   90
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   13626
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmIntraOffice.frx":46CA
      End
      Begin PT_XP_Button.PT_XP cmdNext 
         Height          =   375
         Left            =   5790
         TabIndex        =   9
         Top             =   180
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
         Picture         =   "frmIntraOffice.frx":474C
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Next"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":4AE6
         PictureAlignment=   5
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdFinancial 
         CausesValidation=   0   'False
         DragIcon        =   "frmIntraOffice.frx":4E38
         Height          =   7740
         Left            =   60
         TabIndex        =   14
         Top             =   600
         Width           =   12045
         _Version        =   458752
         _ExtentX        =   21246
         _ExtentY        =   13652
         _StockProps     =   64
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         DAutoSave       =   0   'False
         DAutoSizeCols   =   1
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
         GrayAreaBackColor=   16777215
         GridColor       =   8388608
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   5
         MaxRows         =   0
         MoveActiveOnFocus=   0   'False
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         ShadowText      =   8388608
         SpreadDesigner  =   "frmIntraOffice.frx":5702
         UnitType        =   0
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   10
         ScrollBarTrack  =   3
         ClipboardOptions=   0
         ShowScrollTips  =   1
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Current  Page"
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
         Left            =   2190
         TabIndex        =   13
         Top             =   8520
         Width           =   1425
      End
      Begin VB.Label lblTotalPage 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pages"
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
         Left            =   90
         TabIndex        =   12
         Top             =   8520
         Width           =   1155
      End
   End
   Begin VB.Frame fraSearch 
      Height          =   2565
      Left            =   1590
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   4875
      Begin VB.ComboBox cboFolders 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   510
         Width           =   2805
      End
      Begin VB.TextBox cctPatient 
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
         Height          =   405
         Left            =   150
         TabIndex        =   16
         Top             =   870
         Visible         =   0   'False
         Width           =   3945
      End
      Begin PT_XP_Button.PT_XP cmdNewPat 
         Height          =   465
         Left            =   150
         TabIndex        =   18
         Top             =   300
         Width           =   1545
         _ExtentX        =   2725
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
         Picture         =   "frmIntraOffice.frx":5A03
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Select Patient"
         Depth           =   5
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":5F9D
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdTo 
         Height          =   225
         Left            =   4350
         TabIndex        =   19
         ToolTipText     =   "Enter/Reset Date"
         Top             =   1290
         Width           =   315
         _ExtentX        =   556
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
         Picture         =   "frmIntraOffice.frx":62EF
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "OK"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":6689
         ShowFocus       =   -1  'True
      End
      Begin MSComCtl2.DTPicker ccdDisharge 
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   1590
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
         Format          =   136314883
         CurrentDate     =   38862.7362384259
         MinDate         =   32874
      End
      Begin MSComCtl2.DTPicker ccdAdmit 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1590
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
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
         Format          =   136314883
         CurrentDate     =   38862.7362384259
         MinDate         =   32874
      End
      Begin PT_XP_Button.PT_XP cmdFrom 
         Height          =   225
         Left            =   1980
         TabIndex        =   22
         ToolTipText     =   "Enter/Reset Date"
         Top             =   1290
         Width           =   315
         _ExtentX        =   556
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
         Picture         =   "frmIntraOffice.frx":69DB
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "OK"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":6D75
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRemove 
         Height          =   345
         Left            =   2460
         TabIndex        =   23
         Top             =   2010
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
         Picture         =   "frmIntraOffice.frx":70C7
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Close"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":7AD9
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdEnctrOK 
         Height          =   345
         Left            =   1170
         TabIndex        =   24
         Top             =   2010
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
         Picture         =   "frmIntraOffice.frx":7E2B
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Select"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmIntraOffice.frx":81C5
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblAdmt 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   2520
         TabIndex        =   26
         Top             =   1320
         Width           =   1725
      End
      Begin VB.Label lblFolders 
         BackStyle       =   0  'Transparent
         Caption         =   "Specific Folder"
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
         Left            =   1920
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
   End
   Begin FPSpreadADO.fpSpread grdTerms 
      Height          =   9045
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   13600
      _Version        =   458752
      _ExtentX        =   23989
      _ExtentY        =   15954
      _StockProps     =   64
      ColHeaderDisplay=   0
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
      MaxCols         =   2
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmIntraOffice.frx":8517
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   100
      ScrollBarTrack  =   1
      ClipboardOptions=   0
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnu_Alert 
      Caption         =   "Alert Type"
      Begin VB.Menu mnu_Message 
         Caption         =   "Message"
      End
      Begin VB.Menu mnu_Scanned 
         Caption         =   "New Scanned Document"
      End
   End
   Begin VB.Menu mnu_Search 
      Caption         =   "Search"
   End
   Begin VB.Menu mnu_Print 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "frmIntraOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecuteA Lib "shell32.dll" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private WithEvents clsPDF As CPDF 'Activate event support
Attribute clsPDF.VB_VarHelpID = -1
Private mFirst As Boolean
Private piPatSelect As Long

Private Function Get_File_Data()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field

    On Error GoTo Error_Get_File_Data
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    strSQL = "Get_Practice_Alert_ID " & clsFolder.Alert_ID
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        If Not .EOF Then
            Clear_clsFolder
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "ALERT_ID"
                                clsFolder.Alert_ID = Fld.Value
                            Case "FOLDER_ID"
                                clsFolder.Folder_ID = Fld.Value
                            Case "FOLDER_NAME"
                                clsFolder.Folder_Name = Fld.Value
                            Case "FILE_NAME"
                                clsFolder.File_Name = Fld.Value
                            Case "FILE_DATE"
                                clsFolder.Create_Date = Fld.Value
                            Case "FILE_PATH"
                                clsFolder.File_Path = Fld.Value
                        End Select
                    End If
                End If
            Next
        End If
    End With

    clsFolder.File_Name = Replace(clsFolder.File_Name, "''", "'")
    clsFolder.File_Path = Replace(clsFolder.File_Path, "''", "'")
    clsFolder.Folder_Name = Replace(clsFolder.Folder_Name, "''", "'")
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_File_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_File_Data", App.EXEName & " " & Me.Name
    Exit Function
    Resume
End Function

Private Function Load_Folder_Name()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field

    On Error GoTo Error_Load_Folder_Name
    Screen.MousePointer = vbHourglass
    cboFolders.Clear
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    strSQL = "Collect_Practice_Alert_Folders"
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        Do Until .EOF
            cboFolders.AddItem (!Folder_Name)
        .MoveNext
        Loop
    End With
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Folder_Name:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Folder_Name", App.EXEName & " " & Me.Name
    Exit Function
    Resume
End Function

Private Function Read_Folder_Search()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim strTerm As String
Dim strValue As String
Dim strStart As String
Dim strEnd As String
Dim strType As String
Dim strList As String
Dim strFolder As String

    On Error GoTo Error_Read_Folder_Search
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    If piPatSelect > 0 Then
        strType = "Patient"
    End If
    If cboFolders.Text <> vbNullString Then
        If Trim(strType) <> vbNullString Then
            strType = strType & "|" & "Folder"
        Else
            strType = "Folder"
        End If
    End If
    If ccdAdmit.CheckBox Then
        If IsDate(ccdAdmit.Value) Then
            If Trim(strType) <> vbNullString Then
                strType = strType & "|" & "Date"
                strStart = Format(ccdAdmit.Value, "mm/dd/yyyy")
            Else
                strType = "Date"
                strStart = Format(ccdAdmit.Value, "mm/dd/yyyy")
            End If
        End If
    End If
    If ccdDisharge.CheckBox Then
        If IsDate(ccdDisharge.Value) Then
            If Trim(strType) <> vbNullString Then
                strType = strType & "|" & "Range"
                strEnd = Format(ccdDisharge.Value, "mm/dd/yyyy")
            Else
                strType = "Range"
                strEnd = Format(ccdDisharge.Value, "mm/dd/yyyy")
            End If
        End If
    End If
    strFolder = Replace(Trim(cboFolders.Text), "'", "''")
    Select Case strType
        Case "Patient"
            strSQL = "Search_Practice_Alert_Patient " & piPatSelect
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" & " " & UCase(udtPatientName.Name), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Case "Patient|Folder"
            strSQL = "Search_Practice_Alert_Patient_Folder " & piPatSelect & ", '" & Trim(cboFolders.Text) & "'"
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" _
            & " " & UCase(udtPatientName.Name) & " " & "FOR" & " " & UCase(cboFolders.Text), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Case "Patient|Folder|Range"
            strEnd = DateAdd("d", 1, strEnd)
            strSQL = "Search_Practice_Alert_Patient_Folder_Range " & piPatSelect & ",'" _
                & Trim(cboFolders.Text) & "','" & Format(strStart, "mm/dd/yyyy") & "','" & Format(strEnd, "mm/dd/yyyy") & "'"
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" _
            & " " & UCase(udtPatientName.Name) & " " & "FOR" & " " & UCase(cboFolders.Text) & " " & "PERIOD FOR" _
            & " " & Format(strStart, "MMMM DD, YYYY") & " " & "-" & " " & Format(strEnd, "MMMM DD, YYYY"), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Case "Patient|Folder|Date"
            strSQL = "Search_Practice_Alert_Patient_Folder_Date " & piPatSelect & ",'" _
                & Trim(cboFolders.Text) & "','" & Format(strStart, "mm/dd/yyyy") & "'"
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" _
            & " " & UCase(udtPatientName.Name) & " " & "FOR" & " " & UCase(cboFolders.Text) & " " & Format(strStart, "MMMM DD, YYYY"), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Case "Patient|Range", "Patient|Date|Range"
            strEnd = DateAdd("d", 1, strEnd)
            strSQL = "Search_Practice_Alert_Patient_Range " & piPatSelect & ",'" & Format(strStart, "mm/dd/yyyy") & "','" & Format(strEnd, "mm/dd/yyyy") & "' "
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" _
            & " " & UCase(udtPatientName.Name) & " " & "FOR PERIOD FOR" _
            & " " & Format(strStart, "MMMM DD, YYYY") & " " & "-" & " " & Format(strEnd, "MMMM DD, YYYY"), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Case "Patient|Date"
            strSQL = "Search_Practice_Alert_Patient_Date " & piPatSelect & ",'" & Format(strStart, "mm/dd/yyyy") & "' "
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" _
            & " " & UCase(udtPatientName.Name) & " " & "FOR" & " " & Format(strStart, "MMMM DD, YYYY"), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Case "Folder"
            strSQL = "Search_Practice_Alert_Folder '" & Trim(strFolder) & "'"
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" _
            & " " & UCase(cboFolders.Text), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Case "Folder|Range", "Folder|Date|Range"
            strEnd = DateAdd("d", 1, strEnd)
            strSQL = "Search_Practice_Alert_Folder_Range '" & Trim(strFolder) & "','" _
                & Format(strStart, "mm/dd/yyyy") & "','" & Format(strEnd, "mm/dd/yyyy") & "' "
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" _
            & " " & UCase(cboFolders.Text) & " " & "PERIOD FOR" _
            & " " & Format(strStart, "MMMM DD, YYYY") & " " & "-" & " " & Format(strEnd, "MMMM DD, YYYY"), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Case "Folder|Date"
            strSQL = "Search_Practice_Alert_Folder_Date '" & Trim(strFolder) & "','" & Format(strStart, "mm/dd/yyyy") & "' "
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" _
            & " " & UCase(cboFolders.Text) & " " & "PERIOD FOR" _
            & " " & Format(strStart, "MMMM DD, YYYY"), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Case "Range", "Date|Range"
            strEnd = DateAdd("d", 1, strEnd)
            strSQL = "Search_Practice_Alert_Range '" & Format(strStart, "mm/dd/yyyy") & "','" & Format(strEnd, "mm/dd/yyyy") & "'"
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" _
            & " " & "PERIOD FOR" _
            & " " & Format(strStart, "MMMM DD, YYYY") & " " & "-" & " " & Format(strEnd, "MMMM DD, YYYY"), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Case "Date"
            strSQL = "Search_Practice_Alert_Date '" & Format(strStart, "mm/dd/yyyy") & "'"
            Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "SCANNED DOCUMENTS FOR" _
            & " " & "PERIOD FOR" _
            & " " & Format(strStart, "MMMM DD, YYYY"), _
                "TypeHAlignCenter", "HEADER", "", HeaderColor
            grdTerms.SetCellBorder 1, SpreadHeader, grdTerms.MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
    End Select
    If Trim(strSQL) <> vbNullString Then
        Set rsCodes = cnCodes.Execute(strSQL)
        With rsCodes
            Do Until .EOF
                Clear_clsFolder
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "ALERT_ID"
                                    clsFolder.Alert_ID = Fld.Value
                                Case "FOLDER_ID"
                                    clsFolder.Folder_ID = Fld.Value
                                Case "FOLDER_NAME"
                                    clsFolder.Folder_Name = Fld.Value
                                Case "FILE_NAME"
                                    clsFolder.File_Name = Fld.Value
                                Case "FILE_DATE"
                                    clsFolder.Create_Date = Fld.Value
                            Case "STATUS_ACTION"
                                clsFolder.Status_Action = Fld.Value
                            End Select
                        End If
                    End If
                Next
                .MoveNext
                clsFolder.File_Name = Replace(clsFolder.File_Name, "''", "'")
                clsFolder.File_Path = Replace(clsFolder.File_Path, "''", "'")
                clsFolder.Folder_Name = Replace(clsFolder.Folder_Name, "''", "'")
                If Trim(clsFolder.Folder_ID) <> vbNullString Then
                    Select Case clsFolder.Folder_Name
                        Case "Laboratory", "Radiology", "Correspondence", "Consultations", "Consultation", _
                            "Physician Referral Correspondence", "Record Requests", _
                            "Radiology Report", "Revoked Consents", "Physicains Referral Notes", "Referring Physician Notes", _
                            "Referral Physician Correspondence", "Physician Referral Records", _
                            "Legal Correspondence", "Referring Doctor Notes", "Referring_Physican", _
                            "Questions Regarding Prescription(s)", "Prescription & Refills", "Tests", "Progress Notes"
                            lngRow = grdTerms.MaxRows + 1
                            grdTerms.MaxRows = lngRow
                            grdTerms.SetRowItemData lngRow, clsFolder.Alert_ID
                            lngCol = 1
                            If IsNumeric(clsFolder.Folder_ID) Then Get_Patient Val(clsFolder.Folder_ID)
                            Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", clsPatient.Name, "TypeHAlignCenter", "FOLDER_ID", clsFolder.Folder_ID
                            grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            lngCol = 2
                            Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", clsFolder.Folder_Name, "TypeHAlignCenter", "FOLDER_NAME"
                            grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            lngCol = 3
                            Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", clsFolder.File_Name, "TypeHAlignCenter", "FILE_NAME"
                            grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            lngCol = 4
                            strList = "HOLD" & Chr(9) & "COMPLETED" & Chr(9) & "FOLLOW UP"
                            Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", strList, "TypeHAlignCenter", "STATUS_ACTION"
                            Set_GridCurSel grdTerms, clsFolder.Status_Action, lngCol, lngRow, "1", "STATUS_ACTION"
                            grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            lngCol = 5
                            If IsDate(clsFolder.Create_Date) Then clsFolder.Create_Date = Format(clsFolder.Create_Date, "mm/dd/yyyy hh:mm:ss AM/PM")
                            Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", clsFolder.Create_Date, "TypeHAlignCenter", "FILE_DATE"
                            grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End Select
                End If
            Loop
        End With
    End If
    Clear_clsFolder
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Folder_Search:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Folder_Search", App.EXEName & " " & Me.Name
    Exit Function
    Resume
End Function

Private Function Update_Alert()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset

    On Error GoTo Error_Update_Alert
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    If Trim(clsProvider.Name) = vbNullString Then
        Get_Providers mlngLogonUser
    End If
    strSQL = "UPDATE PRACTICE_ALERT SET  REVIEW_DATE = '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "', " _
        & "REVIEW_BY = '" & clsProvider.Name & "', STATUS_ALERT = '" & clsFolder.Status_Action & "'  WHERE ALERT_ID = " & clsFolder.Alert_ID
    Set rsCodes = cnCodes.Execute(strSQL)
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_Alert:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL & " " & clsFolder.File_Path, "Update_Alert", App.EXEName & " " & Me.Name
    Exit Function
    Resume
End Function

Private Function Write_File()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset

    On Error GoTo Error_Write_File
    Screen.MousePointer = vbHourglass
    clsFolder.File_Name = Replace(clsFolder.File_Name, "'", "''")
    clsFolder.File_Path = Replace(clsFolder.File_Path, "'", "''")
    clsFolder.Folder_Name = Replace(clsFolder.Folder_Name, "'", "''")
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    strSQL = "Get_Practice_Alert_Collection '" & clsFolder.Folder_ID & "','" & Trim(clsFolder.Folder_Name) & "','" _
        & Trim(clsFolder.File_Name) & "','" & Format(clsFolder.Create_Date, "mm/dd/yyyy") & "'"
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        If Not .EOF Then
            clsFolder.Alert_ID = (!Alert_ID)
        End If
    End With
    If clsFolder.Alert_ID = 0 Then
        strSQL = "Write_Practice_Alert '" & clsFolder.Folder_ID & "','" & clsFolder.Folder_Name & "', '" _
            & clsFolder.File_Name & "','" & clsFolder.Create_Date & "', '" & clsFolder.File_Path & "'"
'        strSQL = "INSERT INTO PRACTICE_ALERT (FOLDER_ID , FOLDER_NAME, FILE_NAME, FILE_DATE) " _
'            & "values ('" & clsFolder.Folder_ID & "','" & clsFolder.Folder_Name & "', '" _
'            & clsFolder.File_Name & "','" & clsFolder.Create_Date & "')"
        Set rsCodes = cnCodes.Execute(strSQL)
    Else
        strSQL = "UPDATE PRACTICE_ALERT SET  REVIEW_DATE = '" & clsFolder.Review_Date & "', " _
            & "REVIEW_BY = '" & clsFolder.Review_By & "', STATUS_ALERT = '" & clsFolder.Status_Action & "' WHERE ALERT_ID = " & clsFolder.Alert_ID
        Set rsCodes = cnCodes.Execute(strSQL)
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_File:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL & " " & clsFolder.Folder_Name & " " & clsFolder.File_Name & " " & clsFolder.Create_Date, "Write_File", App.EXEName & " " & Me.Name
    Exit Function
    Resume
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
vstrString As String, vstrHAligh As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, _
Optional vstrCellNote As String, Optional vstrHAlign As String)
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
              If Trim(vstrHAlign) = vbNullString Then
                .TypeHAlign = TypeHAlignCenter
              Else
                .TypeHAlign = vstrHAlign
              End If
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
          .BlockMode = False
    End With
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
                If Trim(vstrData) = Trim(.TypeComboBoxString) Then
                    .TypeComboBoxCurSel = Idx
                    Exit For
                End If
            Next Idx
        End If
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
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
        If vlngBackColor > 0 Then .BackColor = vlngBackColor
        If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
    .ReDraw = True
    End With
End Function

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
            If vbooUnderLine Then .FontUnderline = True
            If vlngBackColor Then .BackColor = vlngBackColor
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
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            .EditMode = False
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypePicture(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long
Dim FSO As New FileSystemObject

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            If vlngRowHeight > 0 Then .RowHeight(vlngRow) = vlngRowHeight
            .CellType = CellTypePicture
            If FSO.FileExists(vstrPhotoPath) Then
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
           ' .TypePictStretch = True
            .TypePictCenter = True
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            'grdPic.SetCellBorder vlngCol, vlngRow, vlngCol, vlngRow, CellBorderIndexOutline, &H800000, CellBorderStyleSolid
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
Private Sub Format_Grid()
Dim lngCol As Long
Dim lngRow As Long

    Screen.MousePointer = vbHourglass
    grdTerms.ProcessTab = True
    grdTerms.FontBold = True
    With grdTerms
        .ReDraw = False
        grdTerms.ClearRange 1, 1, grdTerms.MaxCols, grdTerms.MaxRows, False
        .MaxCols = 5
        .MaxRows = 0
        .ColHeaderRows = 4
        .ColWidth(1) = 20
        .ColWidth(2) = 25
        .ColWidth(3) = 25
        .ColWidth(4) = 20
        .ColWidth(5) = 20
        Format_CellTypeEdit grdTerms, SpreadHeader, 1, "maxcols", "RECENT SCANNED DOCUMENTS", "TypeHAlignCenter", "HEADER", "", HeaderColor
        grdTerms.SetCellBorder 1, SpreadHeader, .MaxCols, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Add_Spacer grdTerms, SpreadHeader + 1
        lngCol = 1
        Format_CellTypeEdit grdTerms, SpreadHeader + 2, lngCol, "1", "PATIENT NAME", "TypeHAlignCenter", "FOLDER_ID", "", HeaderColor
        grdTerms.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdTerms, SpreadHeader + 2, lngCol, "1", "FOLDER NAME", "TypeHAlignCenter", "FOLDER_NAME", "", HeaderColor
        grdTerms.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeEdit grdTerms, SpreadHeader + 2, lngCol, "1", "FILE NAME", "TypeHAlignCenter", "FOLDER_NAME", "", HeaderColor
        grdTerms.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeEdit grdTerms, SpreadHeader + 2, lngCol, "1", "STATUS ACTION", "TypeHAlignCenter", "STATUS_ACTION", "", HeaderColor
        grdTerms.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdTerms, SpreadHeader + 2, lngCol, "1", "RECEIVED DATE", "TypeHAlignCenter", "RECEIVED_DATE", "", HeaderColor
        grdTerms.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Add_Spacer grdTerms, SpreadHeader + 3
        .ReDraw = True
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
Private Sub Add_Spacer(vObjGrid As Object, Optional vlngRow As Long)
Dim lngRow As Long

    With vObjGrid
        If vlngRow >= 0 Then
             vlngRow = vlngRow + 1
            .MaxRows = vlngRow
            lngRow = vlngRow
        Else
            lngRow = vlngRow
        End If
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
Private Function Read_Folder()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim strTerm As String
Dim strValue As String
Dim strStart As String
Dim strList As String

    On Error GoTo Error_Read_Folder
    Screen.MousePointer = vbHourglass
    strStart = DateAdd("d", -30, Now)
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    strSQL = "Collect_Practice_Alert_Greater_Date '" & Format(strStart, "mm/dd/yyyy") & "'"
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        Do Until .EOF
            Clear_clsFolder
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "ALERT_ID"
                                clsFolder.Alert_ID = Fld.Value
                            Case "FOLDER_ID"
                                clsFolder.Folder_ID = Fld.Value
                            Case "FOLDER_NAME"
                                clsFolder.Folder_Name = Fld.Value
                            Case "FILE_NAME"
                                clsFolder.File_Name = Fld.Value
                            Case "FILE_DATE"
                                clsFolder.Create_Date = Fld.Value
                            Case "STATUS_ACTION"
                                clsFolder.Status_Action = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext

            clsFolder.File_Name = Replace(clsFolder.File_Name, "''", "'")
            clsFolder.File_Path = Replace(clsFolder.File_Path, "''", "'")
            clsFolder.Folder_Name = Replace(clsFolder.Folder_Name, "''", "'")
            If Trim(clsFolder.Folder_ID) <> vbNullString Then
                Select Case clsFolder.Folder_Name
                    Case "Laboratory", "Radiology", "Correspondence", "Consultations", "Consultation", _
                        "Physician Referral Correspondence", "Record Requests", _
                        "Radiology Report", "Revoked Consents", "Physicains Referral Notes", "Referring Physician Notes", _
                        "Referral Physician Correspondence", "Physician Referral Records", _
                        "Legal Correspondence", "Referring Doctor Notes", "Referring_Physican", _
                        "Questions Regarding Prescription(s)", "Prescription & Refills", "Tests"
                        lngRow = grdTerms.MaxRows + 1
                        grdTerms.MaxRows = lngRow
                        grdTerms.SetRowItemData lngRow, clsFolder.Alert_ID
                        lngCol = 1
                        If IsNumeric(clsFolder.Folder_ID) Then Get_Patient Val(clsFolder.Folder_ID)
                        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", clsPatient.Name, "TypeHAlignCenter", "FOLDER_ID", clsFolder.Folder_ID
                        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 2
                        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", clsFolder.Folder_Name, "TypeHAlignCenter", "FOLDER_NAME"
                        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 3
                        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", clsFolder.File_Name, "TypeHAlignCenter", "FILE_NAME"
                        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 4
                        strList = "HOLD" & Chr(9) & "COMPLETED" & Chr(9) & "FOLLOW UP"
                        Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", strList, "TypeHAlignCenter", "STATUS_ACTION"
                        Set_GridCurSel grdTerms, clsFolder.Status_Action, lngCol, lngRow, "1", "STATUS_ACTION"
                        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 5
                        If IsDate(clsFolder.Create_Date) Then clsFolder.Create_Date = Format(clsFolder.Create_Date, "mm/dd/yyyy hh:mm:ss AM/PM")
                        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", clsFolder.Create_Date, "TypeHAlignCenter", "FILE_DATE"
                        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End Select
            End If
        Loop
    End With
    Clear_clsFolder
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Folder:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Folder", App.EXEName & " " & Me.Name
    Exit Function
    Resume
End Function

Private Sub cmdClose_Click()
    fraMODI.Visible = False
    Exit Sub
End Sub

Private Sub cmdEnctrOK_Click()
    Format_Grid
    Read_Folder_Search
    Exit Sub
End Sub

Private Sub cmdNewPat_Click()

    cctPatient.Text = vbNullString
    piPatSelect = 0
    Clear_UdtPatient
    frmListPatient.mlngListItemData = 0
    frmListPatient.Show vbModal
    If frmListPatient.mlngListItemData <> 0 Then
        piPatSelect = Val(frmListPatient.mlngListItemData)
    End If
    Get_PatientName piPatSelect
    cctPatient.Text = udtPatientName.Name
    Screen.MousePointer = vbDefault
    cctPatient.Visible = True
    Exit Sub
End Sub

Private Sub cmdPrintRecords_Click()
    On Error GoTo Error_PrintRecord
    Screen.MousePointer = vbHourglass
    mbooCancelPrint = False
    If rtfScan.Visible Then
        CD_ShowPrinter (Me.hwnd)
        If mbooCancelPrint Then
            MsgBox "Print Command Cancelled", vbInformation
            mbooCancelPrint = False
            Exit Sub
        Else
            rtfScan.SelPrint Printer.hdc
        End If
'    Else
'        Printer.TrackDefault = True
'        CD_ShowPrinter (Me.hwnd)
'        If mbooCancelPrint Then
'            MsgBox "Print Command Cancelled", vbInformation
'            mbooCancelPrint = False
'            Exit Sub
'        Else
'            miDoc.PrintOut Printer.hdc
'        End If
    End If
    mbooCancelPrint = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_PrintRecord:
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
End Sub

Private Sub cmdRemove_Click()
    fraSearch.Visible = False
    Exit Sub
End Sub

Private Sub cmdRotate_Click()
'Dim i%
'    For i = 0 To miDoc.images.Count - 1
'        If i = MiDocMODI.PageNum Then
'            miDoc.images(i).rotate 90
'            Exit For
'        End If
'    Next i
'    Exit Sub
End Sub

Private Sub cmdSingleView_Click()
'    MiDocMODI.DocViewMode = miDOCVIEWMODE_SINGLEPAGEVIEW
'
'    cmdRotate.Visible = True
'    cmdNext.Visible = True
'    cmdPrev.Visible = True
'    Exit Sub
End Sub

Private Sub Form_Activate()

    On Error GoTo Error_Active
    If mFirst Then
        mFirst = False
        Format_Grid
        Read_Folder
    End If
    centerForm Me
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Active:
    Exit Sub
    Resume
End Sub

Private Sub Form_Load()
    mFirst = True
    Exit Sub
End Sub

Private Function Get_Excel(vstrFile As String)
Dim y As Boolean, z As Boolean
Dim Var As Variant
Dim X As Integer, listcount As Integer, Handle As Integer
Dim List(10) As String

    On Error Resume Next
    
    With grdFinancial
        ' Check if file is an Excel file and set result to x
        X = .IsExcelFile(vstrFile)
    
        ' If file is Excel file, tell user, import sheet
        ' list, and set result to y
        If X = 1 Then
            y = .GetExcelSheetList(vstrFile, List, listcount, "C:\ILOGFILE.TXT", Handle, True)
            ' If received sheet list, tell user, import file,
            ' and set result to z
            If y = True Then
                MsgBox "Got Excel File.", , "Status"
                z = .ImportExcelSheet(Handle, 0)
                ' Tell user result based on T/F value of z
                If z = True Then
                    MsgBox "Import complete.", , "Result"
                    .Visible = True
                Else
                    MsgBox "Import did not succeed.", , "Result"
                    .Visible = False
                End If
            Else
                ' Tell user cannot obtain sheet list
                MsgBox "Cannot return information for Excel file.", , "Result"
            End If
        Else
            ' Tell user file is not Excel file or is locked
            MsgBox "File is not an Excel file or is locked and cannot be imported.", , "Invalid File Type or Locked"
        End If
    End With
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Exit Sub
End Sub

Private Sub grdTerms_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset

    On Error GoTo Error_Update_Status
    If Row < 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    clsFolder.Status_Action = vbNullString
    With grdTerms
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then
            If .CellTag = "STATUS_ACTION" Then
                .GetText Col, Row, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    clsFolder.Status_Action = varCellValue
                End If
                clsFolder.Alert_ID = .GetRowItemData(Row)
            End If
        End If
    End With
    If Trim(clsFolder.Status_Action) <> vbNullString And Val(clsFolder.Alert_ID) > 0 Then
        If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
        If Trim(clsProvider.Name) = vbNullString Then
            Get_Providers mlngLogonUser
        End If
        strSQL = "Update_Practice_Alert_Status '" & clsFolder.Status_Action & "', '" _
            & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," & clsFolder.Alert_ID
        Set rsCodes = cnCodes.Execute(strSQL)
        If cnCodes.State = adStateOpen Then
            cnCodes.Close
            Set cnCodes = Nothing
            Set rsCodes = Nothing
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Update_Status:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_Status", App.EXEName & " " & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim strPath As String
Dim FileType As String
Dim sFile As String, sExt As String, lPos As Long, oApp As Object
Dim miImage As MODI.Image
Dim strCmd As String

    On Error GoTo Error_GridTerm_Select
    Clear_clsFolder
    txtCurPage.Text = vbNullString
    txtTotalPage.Text = vbNullString
    With grdTerms
        .BlockMode = True
            .Col = Col
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = grdCurItem
         .BlockMode = False
        .Col = 1
        .Row = Row
         clsFolder.Alert_ID = .GetRowItemData(Row)
         If clsFolder.Alert_ID > 0 Then Get_File_Data
        .Col = 4
        .Row = Row
        If .CellTag <> vbNullString Then
            If .CellTag = "STATUS_ACTION" Then
                .GetText Col, Row, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    clsFolder.Status_Action = varCellValue
                End If
                clsFolder.Alert_ID = .GetRowItemData(Row)
            End If
        End If
    End With
    If clsFolder.Alert_ID > 0 Then
        grdFinancial.Visible = False
        fraMODI.Visible = False
        rtfScan.Visible = False
        If FileExists(Trim(clsFolder.File_Path)) Then
            FileType = Right(clsFolder.File_Name, 4)
            If Left(FileType, 1) = "." Then
                FileType = Right(FileType, 3)
            End If
            Select Case UCase(FileType)
                Case "TIF", "TIFF", "MDI"
                    fraMODI.Visible = False
                    ShellExecuteA Me.hwnd, "open", Trim(clsFolder.File_Path), vbNullString, vbNullString, 1
                Case "PDF"
                    fraMODI.Visible = False
                    ShellExecuteA Me.hwnd, "open", Trim(clsFolder.File_Path), vbNullString, vbNullString, 1
                    Set clsPDF = Nothing
                Case "RFT", "DOC", "TXT"
                    cmdRotate.Visible = False
                    cmdPrintRecords.Visible = True
                    rtfScan.Visible = True
                    cmdSingleView.Visible = False
                    cmdMultiView.Visible = False
                    If FileExists(clsFolder.File_Path) Then rtfScan.LoadFile clsFolder.File_Path
                    txtCurPage.Text = "1"
                    txtTotalPage.Text = "1"
                    fraMODI.Visible = True
                Case "XLS"
                    rtfScan.Visible = False
                    Get_Excel clsFolder.File_Path
                    grdFinancial.Visible = True
                    fraMODI.Visible = True
            End Select
        End If
        If clsFolder.Status_Action = vbNullString Then
            clsFolder.Status_Action = "COMPLETED"
        End If
        If Trim(clsFolder.File_Name) <> vbNullString And Trim(clsFolder.Status_Action) <> vbNullString Then
            Update_Alert
        End If
    End If
    Exit Sub
Error_GridTerm_Select:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, clsFolder.File_Name, "GridTerm_Select", App.EXEName & " " & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub grdTerms_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row <= 0 Then Exit Sub
    With grdTerms
        .BlockMode = True
            .Col = Col
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
             .BackColor = NormalColor
         .BlockMode = False
    End With
    Exit Sub
End Sub


Private Sub mnu_Close_Click()
    If Check_Scanned_Docs Then
        frmMDI.cmdInterAlert.BackColor = &HD6D8FE
        frmMDI.cmdInterAlert.BorderColor = &H404080
        frmMDI.cmdInterAlert.ForeColor = &H0&
    Else
        frmMDI.cmdInterAlert.BackColor = &HFFFFFF
        frmMDI.cmdInterAlert.BorderColor = &H743C00
        frmMDI.cmdInterAlert.ForeColor = &H800000
    End If
    Unload Me
    Exit Sub
End Sub


Private Sub mnu_Message_Click()
    frmTask.Show vbModal
    Exit Sub
End Sub


Private Sub mnu_Print_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    mbooCancelPrint = False
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.hwnd)
      If mbooCancelPrint Then
            MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Practice Alerts"
            mbooCancelPrint = False
            Exit Sub
      End If
      With grdTerms
        .PrintSmartPrint = False
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Practice_Alert" & "_" & Format(Now, "MM_DD_YYYY")
        .PrintHeader = "Practice_Alert" & "_" & Format(Now, "MM_DD_YYYY")
        .PrintFooter = "/cPage #/p/n"
        .PrintBorder = True
        .PrintColor = True
        .PrintGrid = True
        .PrintOrientation = PrintOrientationLandscape
        .PrintMarginTop = 200
        .PrintMarginBottom = 50
        .PrintMarginLeft = 300
        .PrintMarginRight = 100
        .hDCPrinter = Printer.hdc
        .PrintType = 0
        .PrintRowHeaders = True
        .PrintShadows = False
        If Printer.Copies > 1 Then
            For i = 1 To Printer.Copies
                .PrintSheet
            Next i
        Else
            .PrintSheet
        End If
      End With
      Exit Sub
Error_cmdPrint_Click:
    MsgBox Err.Description, vbInformation
    Exit Sub
End Sub

Private Sub mnu_Scanned_Click()
    If IsFormLoaded("frmTask") Then
        Unload frmTask
    End If
End Sub


Private Sub mnu_Search_Click()
    piPatSelect = 0
    Clear_UdtPatient
    fraSearch.Visible = True
    Load_Folder_Name
    Exit Sub
End Sub


