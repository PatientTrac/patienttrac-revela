VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{6514F5A0-641C-11D2-9FD0-0020AF131A57}#3.0#0"; "fpFlp30.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmMultiAppt 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Multiple Patient Appointments"
   ClientHeight    =   7875
   ClientLeft      =   5115
   ClientTop       =   2910
   ClientWidth     =   13785
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMultiAppt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   13785
   Begin VB.Frame fraControl 
      BackColor       =   &H00E0E0E0&
      Height          =   3465
      Left            =   1470
      TabIndex        =   19
      Top             =   3390
      Visible         =   0   'False
      Width           =   9615
      Begin LpADOLib.fpComboAdo cboVisitType 
         Height          =   330
         Left            =   6450
         TabIndex        =   20
         Top             =   1290
         Width           =   2895
         _Version        =   196608
         _ExtentX        =   5106
         _ExtentY        =   582
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         BackColor       =   -2147483643
         ForeColor       =   0
         Text            =   ""
         Columns         =   0
         Sorted          =   2
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   -1
         ColumnWidthScale=   3
         RowHeight       =   -1
         WrapList        =   -1  'True
         WrapWidth       =   0
         AutoSearch      =   1
         SearchMethod    =   2
         VirtualMode     =   0   'False
         VRowCount       =   0
         DataSync        =   3
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483627
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   14737632
         BorderWidth     =   1
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   14737632
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
         Style           =   2
         MaxDrop         =   8
         ListWidth       =   -1
         EditHeight      =   -1
         GrayAreaColor   =   -2147483633
         ListLeftOffset  =   0
         ComboGap        =   -2
         MaxEditLen      =   1000
         VirtualPageSize =   0
         VirtualPagesAhead=   0
         ExtendCol       =   1
         ColumnLevels    =   1
         ListGrayAreaColor=   -2147483637
         GroupHeaderHeight=   -1
         GroupHeaderShow =   0   'False
         AllowGrpResize  =   0
         AllowGrpDragDrop=   0
         MergeAdjustView =   0   'False
         ColumnHeaderShow=   -1  'True
         ColumnHeaderHeight=   -1
         GrpsFrozen      =   0
         BorderGrayAreaColor=   14737632
         ExtendRow       =   0
         ListPosition    =   0
         ButtonThreeDAppearance=   0
         DataMemberList  =   ""
         OLEDragMode     =   0
         OLEDropMode     =   0
         Redraw          =   -1  'True
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   500
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   0   'False
         EditAlignH      =   0
         EditAlignV      =   0
         AllowAnimate    =   -1  'True
         ColDesigner     =   "frmMultiAppt.frx":08CA
      End
      Begin VB.ComboBox cboStartTime 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   4350
         TabIndex        =   44
         Top             =   1305
         Width           =   1800
      End
      Begin VB.ComboBox cboLocation 
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
         ItemData        =   "frmMultiAppt.frx":0BFD
         Left            =   210
         List            =   "frmMultiAppt.frx":0BFF
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   570
         Width           =   4860
      End
      Begin VB.ComboBox cboDayWeek 
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
         ItemData        =   "frmMultiAppt.frx":0C01
         Left            =   4320
         List            =   "frmMultiAppt.frx":0C1A
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2070
         Width           =   1845
      End
      Begin VB.ComboBox cboResProvider 
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
         Left            =   5415
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   570
         Width           =   3915
      End
      Begin VB.ComboBox cboRepAppt 
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
         ItemData        =   "frmMultiAppt.frx":0C5E
         Left            =   2280
         List            =   "frmMultiAppt.frx":0C60
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2070
         Width           =   1695
      End
      Begin VB.ComboBox cboRepeatTerm 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2070
         Width           =   1725
      End
      Begin VB.TextBox cctComments 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6450
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   2040
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker ccdTermDate 
         Height          =   345
         Left            =   2310
         TabIndex        =   26
         Top             =   1290
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
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
         Format          =   171900931
         CurrentDate     =   38792
      End
      Begin MSComCtl2.DTPicker ccdStart 
         Height          =   345
         Left            =   240
         TabIndex        =   27
         Top             =   1290
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
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
         Format          =   171900931
         CurrentDate     =   38792
      End
      Begin PT_XP_Button.PT_XP cmdSched 
         Height          =   345
         Left            =   3570
         TabIndex        =   40
         Top             =   2970
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
         Picture         =   "frmMultiAppt.frx":0C62
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Apply"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMultiAppt.frx":0FFC
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdExit 
         Height          =   345
         Left            =   4860
         TabIndex        =   41
         Top             =   2970
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
         Picture         =   "frmMultiAppt.frx":134E
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "&Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMultiAppt.frx":1D60
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Facility"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   240
         TabIndex        =   38
         Top             =   330
         Width           =   765
      End
      Begin VB.Label lblDayWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "Day of Week"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   4320
         TabIndex        =   36
         ToolTipText     =   "Frequency Number"
         Top             =   1830
         Width           =   1485
      End
      Begin VB.Label lblTermDate 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2310
         TabIndex        =   35
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Provider"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   5415
         TabIndex        =   34
         Top             =   330
         Width           =   855
      End
      Begin VB.Label lblTermValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Term Value"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2280
         TabIndex        =   33
         ToolTipText     =   "Week, Month, Year"
         Top             =   1830
         Width           =   1530
      End
      Begin VB.Label lblTerm 
         BackStyle       =   0  'Transparent
         Caption         =   "Every "
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   240
         TabIndex        =   32
         ToolTipText     =   "Frequency Number"
         Top             =   1830
         Width           =   945
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appt Time"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4380
         TabIndex        =   31
         Top             =   1050
         Width           =   1035
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label lblApptType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appt Type"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6450
         TabIndex        =   29
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label lblComment 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6480
         TabIndex        =   28
         Top             =   1800
         Width           =   1725
      End
   End
   Begin VB.TextBox cctBlockConflict 
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   17220
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2580
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ComboBox cboResource 
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   15720
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   990
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Frame fraNumber 
      BackColor       =   &H00C0FFFF&
      Height          =   1005
      Left            =   1740
      TabIndex        =   0
      Top             =   3660
      Visible         =   0   'False
      Width           =   2025
      Begin VB.ComboBox cboRepTime 
         Height          =   330
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   180
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker ccdDate 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         DateIsNull      =   -1  'True
         Format          =   171900929
         CurrentDate     =   39288
      End
      Begin PT_XP_Button.PT_XP cmdAccept 
         Height          =   285
         Left            =   510
         TabIndex        =   2
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
         Picture         =   "frmMultiAppt.frx":20B2
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMultiAppt.frx":2AC4
         ShowFocus       =   -1  'True
      End
   End
   Begin PT_XP_Button.PT_XP cmdRemoveRow 
      Height          =   465
      Left            =   3510
      TabIndex        =   4
      Top             =   2850
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
      Picture         =   "frmMultiAppt.frx":2E16
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Terminate Repeats"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":33B0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdGridRow 
      Height          =   465
      Left            =   15870
      TabIndex        =   5
      ToolTipText     =   "Update Appointment Row"
      Top             =   6450
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmMultiAppt.frx":3702
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Change Appt"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":3C9C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPrint 
      Height          =   465
      Left            =   12735
      TabIndex        =   6
      Top             =   2850
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
      Picture         =   "frmMultiAppt.frx":3FEE
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Print"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":4588
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdResource 
      Height          =   285
      Left            =   18390
      TabIndex        =   7
      ToolTipText     =   "Manage Resources"
      Top             =   1020
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMultiAppt.frx":48DA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmMultiAppt.frx":4E74
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdProvTypes 
      Height          =   285
      Left            =   18390
      TabIndex        =   8
      ToolTipText     =   "Provider Types"
      Top             =   1350
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMultiAppt.frx":51C6
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmMultiAppt.frx":5760
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdMultiClose 
      Height          =   375
      Left            =   6390
      TabIndex        =   10
      Top             =   7440
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
      Picture         =   "frmMultiAppt.frx":5AB2
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "&Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":64C4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdMultiSave 
      Height          =   375
      Left            =   5130
      TabIndex        =   11
      Top             =   8340
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
      Picture         =   "frmMultiAppt.frx":6816
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "&Apply"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":6BB0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPatient 
      Height          =   405
      Left            =   90
      TabIndex        =   13
      Top             =   120
      Width           =   1785
      _ExtentX        =   3149
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
      Picture         =   "frmMultiAppt.frx":6F02
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Patient Select"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":749C
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdMultiAppt 
      Height          =   4005
      Left            =   60
      TabIndex        =   15
      Top             =   3360
      Width           =   13710
      _Version        =   524288
      _ExtentX        =   24183
      _ExtentY        =   7064
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
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      ShadowText      =   8388608
      SpreadDesigner  =   "frmMultiAppt.frx":77EE
      UserResize      =   0
      Appearance      =   1
      ClipboardOptions=   0
      CellNoteIndicator=   3
      AppearanceStyle =   0
   End
   Begin PT_XP_Button.PT_XP cmdConflict 
      Height          =   465
      Left            =   9570
      TabIndex        =   17
      Top             =   2850
      Visible         =   0   'False
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
      Picture         =   "frmMultiAppt.frx":7CEF
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Scheduling Conflicts"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":8289
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdRepeatHist 
      Height          =   2805
      Left            =   4170
      TabIndex        =   18
      Top             =   0
      Width           =   9600
      _Version        =   524288
      _ExtentX        =   16933
      _ExtentY        =   4948
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   16777215
      GridColor       =   4194304
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
      SpreadDesigner  =   "frmMultiAppt.frx":85DB
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   100
      ScrollBarTrack  =   1
      ClipboardOptions=   0
      AppearanceStyle =   0
   End
   Begin PT_XP_Button.PT_XP cmdNewRepeat 
      Height          =   465
      Left            =   90
      TabIndex        =   39
      Top             =   2850
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
      Picture         =   "frmMultiAppt.frx":8AD5
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "New Repeat Appt"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":906F
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdUpdate 
      Height          =   465
      Left            =   1800
      TabIndex        =   42
      Top             =   2850
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
      Picture         =   "frmMultiAppt.frx":93C1
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Update Appt Row"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMultiAppt.frx":995B
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblPatName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblBlock 
      BackStyle       =   0  'Transparent
      Caption         =   "Block Time Conflicts"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   17250
      TabIndex        =   12
      Top             =   2370
      Width           =   2760
   End
   Begin VB.Label lblRes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Resource"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   15720
      TabIndex        =   9
      Top             =   750
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "frmMultiAppt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Private mbooByPass As Boolean
Private mlngGridRow As Long
Private mlngGridCol As Long
Private fpDeltaLeft As Long
Private fpDeltaTop As Long
Private flngDeltaX  As Long
Private flngDeltaY  As Long
Private udtBlock As udtBlock
Private Type udtBlock
    StartDate As String
    EndDate As String
    StartTime As String
    EndTime As String
    Reason As String
    Provider  As String
    ProviderID As Long
    FacilityID As Long
    Duration As String
    DayWeek As String
End Type
Private Function Add_Header()
Dim CellNote As Long
Dim strPhoto As String
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr
Dim fso As New FileSystemObject

    On Error GoTo Error_Add_Header
    Screen.MousePointer = vbHourglass
    grdMultiAppt.ClearRange 1, SpreadHeader, grdMultiAppt.MaxCols, SpreadHeader + 2, False
    Remove_Span
    grdMultiAppt.ColHeaderRows = 5
    Get_DefaultFacility
    If udtFacility.Logo <> vbNullString Then
        If fso.FileExists(udtFacility.Logo) Then
            Format_CellTypePicture grdMultiAppt, SpreadHeader, 1, "1", udtFacility.Logo, "TypeHAlignCenter", "LOGO", "75"
            Format_CellTypeEdit grdMultiAppt, SpreadHeader, 2, "8", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
        Else
            Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "10", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
        End If
    End If
        grdMultiAppt.SetCellBorder 1, SpreadHeader + 2, 10, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdMultiAppt
        lngRow = SpreadHeader + 2
        lngCol = 1
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 2, lngCol, "10", "REPEAT APPOINTMENT SCHEDULE FOR" & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName), "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 2, 5, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngRow = SpreadHeader + 3
        lngCol = 1
        grdMultiAppt.RemoveCellSpan 1, lngRow
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 3, lngCol, "1", "PROVIDER", "TypeHAlignCenter", "PROVIDER_ID", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 3, lngCol, SpreadHeader + 3, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 3, lngCol, "1", "APPT DATE", "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 3, lngCol, SpreadHeader + 3, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 3, lngCol, "1", "APPT TIME", "TypeHAlignCenter", "APPT_TIME", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 3, lngCol, SpreadHeader + 3, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 3, lngCol, "1", "APPT TYPE", "TypeHAlignCenter", "APPT_TYPE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 3, lngCol, SpreadHeader + 3, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 3, lngCol, "1", "DURATION", "TypeHAlignCenter", "DURATION", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 3, lngCol, SpreadHeader + 3, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 4, grdMultiAppt
    Screen.MousePointer = vbDefault
    Exit Function
Error_Add_Header:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Add_Header", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function
Private Function Calc_Duration(vstrReason As String) As String
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset

    On Error GoTo Error_Calc_Duration
    Screen.MousePointer = vbHourglass
    
    clsGridFormat.List = vbNullString
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(1)
    strSQL = "SELECT APPT_TIME, VISIT_TYPE FROM APPOINTMENT_TYPE WHERE APPT_CODE = '" & vstrReason & "'"
    Set rsAppt = cnAppt.Execute(strSQL)
    If Not rsAppt.EOF Then
        If rsAppt.Fields("APPT_TIME") <> vbNullString Then udtRepeatAppt.Duration = rsAppt.Fields("APPT_TIME")
        If Not IsNull(rsAppt.Fields("VISIT_TYPE")) Then
            If rsAppt.Fields("VISIT_TYPE") <> vbNullString Then udtAppointment.SchedType = rsAppt.Fields("VISIT_TYPE")
        End If
    End If
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

Private Function Format_Repeat_History()
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr

    On Error GoTo Error_Format_Repeat_History
    Screen.MousePointer = vbHourglass

    With grdRepeatHist
        .Redraw = False
        .ClearRange 1, SpreadHeader, .MaxCols, .MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .MaxCols = 5
        .FontName = "Times New Roman"
        .GridShowHoriz = True
        .GridShowVert = True
        .RowHeadersShow = False
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .ScrollBars = ScrollBarsBoth
        grdRepeatHist.ColWidth(1) = 23
        grdRepeatHist.ColWidth(2) = 20
        grdRepeatHist.ColWidth(3) = 9
        grdRepeatHist.ColWidth(4) = 9
        grdRepeatHist.ColWidth(5) = 15
        lngRow = SpreadHeader
        lngCol = 1
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "8", "REPEAT APPOINTMENT HISTORY FOR" & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName), "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, 8, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .RowHeight(SpreadHeader) = 15
        lngRow = SpreadHeader + 1
        lngCol = 1
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", "FACILITY", "TypeHAlignCenter", "FACILITY_ID", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", "PROVIDER", "TypeHAlignCenter", "PROVIDER_ID", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", "START DATE", "TypeHAlignCenter", "START_DATE", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", "END DATE", "TypeHAlignCenter", "TERMINATION_DATE", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", "STATUS", "TypeHAlignCenter", "ACTIVE", HeaderColor
        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 2, grdRepeatHist
        .Redraw = False
'        If Trim(udtSchedPatient.IntId) <> vbNullString Then Read_MultiAppt Val(udtSchedPatient.IntId)
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_Repeat_History:
    Screen.MousePointer = vbDefault
End Function

Public Function Get_Appt_Providers()
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
    cboResProvider.Clear
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    strSQL = "SELECT b.*, a.PROVIDER_ID from APPOINTMENT_PROVIDERS a, PROVIDERS b " _
        & "where a.FACILITY_ID = " & udtRepeatAppt.FacilityID & " and b.PROVIDER_ID = a.PROVIDER_ID and b.ACTIVE = 1 ORDER BY b.LAST_NAME, b.FIRST_NAME"
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
                If clsGridFormat.List = vbNullString Then
                    clsGridFormat.List = clsProvider.Name
                Else
                    clsGridFormat.List = clsGridFormat.List & Chr(9) & clsProvider.Name
                End If
                 cboResProvider.AddItem clsProvider.Name
                If IsNumeric(lngProvID) Then cboResProvider.ItemData(cboResProvider.NewIndex) = lngProvID
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

Private Function Load_Form()
Dim Idx As Integer

        Load_Resources
        Get_Appt_Providers
        Load_ReportTime
        Load_Facilities
        ccdStart.Value = Now + 1
        If Trim(udtRepeatAppt.WeekDay) <> vbNullString Then
           For Idx = 0 To cboDayWeek.listcount - 1
               If cboDayWeek.List(Idx) = Trim(udtRepeatAppt.WeekDay) Then
                   cboDayWeek.Text = cboDayWeek.List(Idx)
                   Exit For
               End If
           Next Idx
        End If
        If Trim(udtRepeatAppt.ProviderID) > 0 Then
           For Idx = 0 To cboResProvider.listcount - 1
               If cboResProvider.ItemData(Idx) = udtRepeatAppt.ProviderID Then
                   cboResProvider.Text = cboResProvider.List(Idx)
                   Exit For
               End If
           Next Idx
        End If
        If cboRepAppt.listcount <= 0 Then Load_ReportTime
        Format_MultiAppt
        Format_Repeat_History
        'Load_Repeat_Appt
        If Trim(udtAppointment.SchedDate) <> vbNullString Then
           ccdStart.Value = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
           If DateDiff("d", Now, ccdStart.Value) <= 0 Then ccdStart.Value = Now + 1
        End If
        cctBlockConflict.Text = vbNullString
        cctBlockConflict.Visible = False
        lblBlock.Visible = False
End Function

Private Sub Load_ReportTime()

    cboRepAppt.Clear
    cboRepAppt.AddItem "Day"
    cboRepAppt.AddItem "Week"
    cboRepAppt.AddItem "Bi-Month"
    cboRepAppt.AddItem "Month"
    cboRepAppt.AddItem "Annual"
    cboRepeatTerm.Clear
    cboRepeatTerm.AddItem "1"
    cboRepeatTerm.AddItem "2"
    cboRepeatTerm.AddItem "3"
    cboRepeatTerm.AddItem "4"
    cboRepeatTerm.AddItem "5"
    cboRepeatTerm.AddItem "6"
    cboRepeatTerm.AddItem "7"
    cboRepeatTerm.AddItem "8"
    cboRepeatTerm.AddItem "9"
    cboRepeatTerm.AddItem "10"
    cboRepeatTerm.AddItem "12"
    Exit Sub
End Sub

Private Function Load_Resources()
Dim cnApptFormat As New ADODB.Connection
Dim rsApptFormat As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim lngCnt As Long
Dim lngCol As Long
Dim lngRow As Long

    On Error GoTo Error_Load_Resources
    Screen.MousePointer = vbHourglass
    If cnApptFormat.State = adStateClosed Then cnApptFormat.open psConnect(0)
    strSQL = "SELECT RESOURCE_NAME, RESOURCE_ID, SEQUENCE_NO FROM APPOINTMENT_RESOURCE WHERE RESOURCE_HEADER = 'True' Order by SEQUENCE_NO"
    Set rsApptFormat = cnApptFormat.Execute(strSQL)
    With rsApptFormat
       Do Until .EOF
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        If UCase(Fld.Name) = "RESOURCE_NAME" Then
                            cboResource.AddItem Trim(Fld.Value)
                            cboResource.ItemData(cboResource.NewIndex) = (!RESOURCE_ID)
                        End If
                    End If
                End If
            Next
           .MoveNext
       Loop
    End With
    If cnApptFormat.State = adStateOpen Then
        cnApptFormat.Close
        Set cnApptFormat = Nothing
        Set rsApptFormat = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Resources:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Resources", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Load_ComboTime()
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim strTerms As String

    On Error GoTo Error_Load_ComboTime
    Screen.MousePointer = vbHourglass
    cboStartTime.Clear
    Select Case UCase(frmPatientAppointment.mInterval)
        Case "EVERY 15 MIN"
           cboStartTime.AddItem "07:00 AM": cboStartTime.AddItem "07:15 AM": cboStartTime.AddItem "07:30 AM": cboStartTime.AddItem "07:45 AM"
           cboStartTime.AddItem "08:00 AM": cboStartTime.AddItem "08:15 AM": cboStartTime.AddItem "08:30 AM": cboStartTime.AddItem "08:45 AM"
           cboStartTime.AddItem "09:00 AM": cboStartTime.AddItem "09:15 AM": cboStartTime.AddItem "09:30 AM": cboStartTime.AddItem "09:45 AM"
           cboStartTime.AddItem "10:00 AM": cboStartTime.AddItem "10:15 AM": cboStartTime.AddItem "10:30 AM": cboStartTime.AddItem "10:45 AM"
           cboStartTime.AddItem "11:00 AM": cboStartTime.AddItem "11:15 AM": cboStartTime.AddItem "11:30 AM": cboStartTime.AddItem "11:45 AM"
           cboStartTime.AddItem "12:00 PM": cboStartTime.AddItem "12:15 PM": cboStartTime.AddItem "12:30 PM": cboStartTime.AddItem "12:45 PM"
           cboStartTime.AddItem "01:00 PM": cboStartTime.AddItem "01:15 PM": cboStartTime.AddItem "01:30 PM": cboStartTime.AddItem "01:45 PM"
           cboStartTime.AddItem "02:00 PM": cboStartTime.AddItem "02:15 PM": cboStartTime.AddItem "02:30 PM": cboStartTime.AddItem "02:45 PM"
           cboStartTime.AddItem "03:00 PM": cboStartTime.AddItem "03:15 PM": cboStartTime.AddItem "03:30 PM": cboStartTime.AddItem "03:45 PM"
           cboStartTime.AddItem "04:00 PM": cboStartTime.AddItem "04:15 PM": cboStartTime.AddItem "04:30 PM": cboStartTime.AddItem "04:45 PM"
           cboStartTime.AddItem "05:00 PM": cboStartTime.AddItem "05:15 PM": cboStartTime.AddItem "05:30 PM": cboStartTime.AddItem "05:45 PM"
           cboStartTime.AddItem "06:00 PM": cboStartTime.AddItem "06:15 PM": cboStartTime.AddItem "06:30 PM": cboStartTime.AddItem "06:45 PM"
           cboStartTime.AddItem "07:00 PM": cboStartTime.AddItem "07:15 PM": cboStartTime.AddItem "07:30 PM": cboStartTime.AddItem "07:45 PM"
           cboStartTime.AddItem "08:00 PM": cboStartTime.AddItem "08:15 PM": cboStartTime.AddItem "08:30 PM": cboStartTime.AddItem "08:45 PM"
       Case "EVERY 5 MIN"
           cboStartTime.AddItem "07:00 AM": cboStartTime.AddItem "07:05 AM": cboStartTime.AddItem "07:10 AM": cboStartTime.AddItem "07:15 AM"
           cboStartTime.AddItem "07:20 AM": cboStartTime.AddItem "07:25 AM": cboStartTime.AddItem "07:30 AM": cboStartTime.AddItem "07:35 AM"
           cboStartTime.AddItem "07:40 AM": cboStartTime.AddItem "07:45 AM": cboStartTime.AddItem "07:50 AM": cboStartTime.AddItem "07:55 AM"
           cboStartTime.AddItem "08:00 AM": cboStartTime.AddItem "08:05 AM": cboStartTime.AddItem "08:10 AM": cboStartTime.AddItem "08:15 AM"
           cboStartTime.AddItem "08:20 AM": cboStartTime.AddItem "08:25 AM": cboStartTime.AddItem "08:30 AM": cboStartTime.AddItem "08:35 AM"
           cboStartTime.AddItem "08:40 AM": cboStartTime.AddItem "08:45 AM": cboStartTime.AddItem "08:50 AM": cboStartTime.AddItem "08:55 AM"
           cboStartTime.AddItem "09:00 AM": cboStartTime.AddItem "09:05 AM": cboStartTime.AddItem "09:10 AM": cboStartTime.AddItem "09:15 AM"
           cboStartTime.AddItem "09:20 AM": cboStartTime.AddItem "09:25 AM": cboStartTime.AddItem "09:30 AM": cboStartTime.AddItem "09:35 AM"
           cboStartTime.AddItem "09:40 AM": cboStartTime.AddItem "09:45 AM": cboStartTime.AddItem "09:50 AM": cboStartTime.AddItem "09:55 AM"
           cboStartTime.AddItem "10:00 AM": cboStartTime.AddItem "10:05 AM": cboStartTime.AddItem "10:10 AM": cboStartTime.AddItem "10:15 AM"
           cboStartTime.AddItem "10:20 AM": cboStartTime.AddItem "10:25 AM": cboStartTime.AddItem "10:30 AM": cboStartTime.AddItem "10:35 AM"
           cboStartTime.AddItem "10:40 AM": cboStartTime.AddItem "10:45 AM": cboStartTime.AddItem "10:50 AM": cboStartTime.AddItem "10:55 AM"
           cboStartTime.AddItem "11:00 AM": cboStartTime.AddItem "11:05 AM": cboStartTime.AddItem "11:10 AM": cboStartTime.AddItem "11:15 AM"
           cboStartTime.AddItem "11:20 AM": cboStartTime.AddItem "11:25 AM": cboStartTime.AddItem "11:30 AM": cboStartTime.AddItem "11:35 AM"
           cboStartTime.AddItem "11:40 AM": cboStartTime.AddItem "11:45 AM": cboStartTime.AddItem "11:50 AM": cboStartTime.AddItem "11:55 AM"
           cboStartTime.AddItem "12:00 PM": cboStartTime.AddItem "12:05 PM": cboStartTime.AddItem "12:10 PM": cboStartTime.AddItem "12:15 PM"
           cboStartTime.AddItem "12:20 PM": cboStartTime.AddItem "12:25 PM": cboStartTime.AddItem "12:30 PM": cboStartTime.AddItem "12:35 PM"
           cboStartTime.AddItem "01:00 PM": cboStartTime.AddItem "01:05 PM": cboStartTime.AddItem "01:10 PM": cboStartTime.AddItem "01:15 PM"
           cboStartTime.AddItem "01:20 PM": cboStartTime.AddItem "01:25 PM": cboStartTime.AddItem "01:30 PM": cboStartTime.AddItem "01:35 PM"
           cboStartTime.AddItem "01:40 PM": cboStartTime.AddItem "01:45 PM": cboStartTime.AddItem "01:50 PM": cboStartTime.AddItem "01:55 PM"
           cboStartTime.AddItem "02:00 PM": cboStartTime.AddItem "02:05 PM": cboStartTime.AddItem "02:10 PM": cboStartTime.AddItem "02:15 PM"
           cboStartTime.AddItem "02:20 PM": cboStartTime.AddItem "02:25 PM": cboStartTime.AddItem "02:30 PM": cboStartTime.AddItem "02:35 PM"
           cboStartTime.AddItem "02:40 PM": cboStartTime.AddItem "02:45 PM": cboStartTime.AddItem "02:50 PM": cboStartTime.AddItem "02:55 PM"
           cboStartTime.AddItem "03:00 PM": cboStartTime.AddItem "03:05 PM": cboStartTime.AddItem "03:10 PM": cboStartTime.AddItem "03:15 PM"
           cboStartTime.AddItem "03:20 PM": cboStartTime.AddItem "03:25 PM": cboStartTime.AddItem "03:30 PM": cboStartTime.AddItem "03:35 PM"
           cboStartTime.AddItem "03:40 PM": cboStartTime.AddItem "03:45 PM": cboStartTime.AddItem "03:50 PM": cboStartTime.AddItem "03:55 PM"
           cboStartTime.AddItem "04:00 PM": cboStartTime.AddItem "04:05 PM": cboStartTime.AddItem "04:10 PM": cboStartTime.AddItem "04:15 PM"
           cboStartTime.AddItem "04:20 PM": cboStartTime.AddItem "04:25 PM": cboStartTime.AddItem "04:30 PM": cboStartTime.AddItem "04:35 PM"
           cboStartTime.AddItem "04:40 PM": cboStartTime.AddItem "04:45 PM": cboStartTime.AddItem "04:50 PM": cboStartTime.AddItem "04:55 PM"
           cboStartTime.AddItem "05:00 PM": cboStartTime.AddItem "05:05 PM": cboStartTime.AddItem "05:10 PM": cboStartTime.AddItem "05:15 PM"
           cboStartTime.AddItem "05:20 PM": cboStartTime.AddItem "05:25 PM": cboStartTime.AddItem "05:30 PM": cboStartTime.AddItem "05:35 PM"
           cboStartTime.AddItem "05:40 PM": cboStartTime.AddItem "05:45 PM": cboStartTime.AddItem "05:50 PM": cboStartTime.AddItem "05:55 PM"
           cboStartTime.AddItem "06:00 PM": cboStartTime.AddItem "06:05 PM": cboStartTime.AddItem "06:10 PM": cboStartTime.AddItem "06:15 PM"
           cboStartTime.AddItem "06:20 PM": cboStartTime.AddItem "06:25 PM": cboStartTime.AddItem "06:30 PM": cboStartTime.AddItem "06:35 PM"
           cboStartTime.AddItem "06:40 PM": cboStartTime.AddItem "06:45 PM": cboStartTime.AddItem "06:50 PM": cboStartTime.AddItem "06:55 PM"
           cboStartTime.AddItem "07:00 PM": cboStartTime.AddItem "07:05 PM": cboStartTime.AddItem "07:10 PM": cboStartTime.AddItem "07:15 PM"
           cboStartTime.AddItem "07:20 PM": cboStartTime.AddItem "07:25 PM": cboStartTime.AddItem "07:30 PM": cboStartTime.AddItem "07:35 PM"
           cboStartTime.AddItem "07:40 PM": cboStartTime.AddItem "07:45 PM": cboStartTime.AddItem "07:50 PM": cboStartTime.AddItem "07:55 PM"
           cboStartTime.AddItem "08:00 PM": cboStartTime.AddItem "08:05 PM": cboStartTime.AddItem "08:10 PM": cboStartTime.AddItem "08:15 PM"
           cboStartTime.AddItem "08:20 PM": cboStartTime.AddItem "08:25 PM": cboStartTime.AddItem "08:30 PM": cboStartTime.AddItem "08:35 PM"
           cboStartTime.AddItem "08:40 PM": cboStartTime.AddItem "08:45 PM": cboStartTime.AddItem "08:50 PM": cboStartTime.AddItem "08:55 PM"
       Case "EVERY 20 MIN"
           cboStartTime.AddItem "07:00 AM": cboStartTime.AddItem "07:20 AM": cboStartTime.AddItem "07:40 AM"
           cboStartTime.AddItem "08:00 AM": cboStartTime.AddItem "08:20 AM": cboStartTime.AddItem "08:40 AM"
           cboStartTime.AddItem "09:00 AM": cboStartTime.AddItem "09:20 AM": cboStartTime.AddItem "09:40 AM"
           cboStartTime.AddItem "10:00 AM": cboStartTime.AddItem "10:20 AM": cboStartTime.AddItem "10:40 AM"
           cboStartTime.AddItem "11:00 AM": cboStartTime.AddItem "11:20 AM": cboStartTime.AddItem "11:40 AM"
           cboStartTime.AddItem "12:00 PM": cboStartTime.AddItem "12:20 PM": cboStartTime.AddItem "12:40 PM"
           cboStartTime.AddItem "01:00 PM": cboStartTime.AddItem "01:20 PM": cboStartTime.AddItem "01:40 PM"
           cboStartTime.AddItem "02:00 PM": cboStartTime.AddItem "02:20 PM": cboStartTime.AddItem "02:40 PM"
           cboStartTime.AddItem "03:00 PM": cboStartTime.AddItem "03:20 PM": cboStartTime.AddItem "03:40 PM"
           cboStartTime.AddItem "04:00 PM": cboStartTime.AddItem "04:20 PM": cboStartTime.AddItem "04:40 PM"
           cboStartTime.AddItem "05:00 PM": cboStartTime.AddItem "05:20 PM": cboStartTime.AddItem "05:40 PM"
           cboStartTime.AddItem "06:00 PM": cboStartTime.AddItem "06:20 PM": cboStartTime.AddItem "06:40 PM"
           cboStartTime.AddItem "07:00 PM": cboStartTime.AddItem "07:20 PM": cboStartTime.AddItem "07:40 PM"
           cboStartTime.AddItem "08:00 PM": cboStartTime.AddItem "08:20 PM": cboStartTime.AddItem "08:40 PM"
       Case "EVERY 10 MIN"
        cboStartTime.AddItem "07:00 AM": cboStartTime.AddItem "07:10 AM": cboStartTime.AddItem "07:20 AM": cboStartTime.AddItem "07:30 AM": cboStartTime.AddItem "07:40 AM": cboStartTime.AddItem "07:50 AM"
        cboStartTime.AddItem "08:00 AM": cboStartTime.AddItem "08:10 AM": cboStartTime.AddItem "08:20 AM": cboStartTime.AddItem "08:30 AM": cboStartTime.AddItem "08:40 AM": cboStartTime.AddItem "08:50 AM"
        cboStartTime.AddItem "09:00 AM": cboStartTime.AddItem "09:10 AM": cboStartTime.AddItem "09:20 AM": cboStartTime.AddItem "09:30 AM": cboStartTime.AddItem "09:40 AM": cboStartTime.AddItem "09:50 AM"
        cboStartTime.AddItem "10:00 AM": cboStartTime.AddItem "10:10 AM": cboStartTime.AddItem "10:20 AM": cboStartTime.AddItem "10:30 AM": cboStartTime.AddItem "10:40 AM": cboStartTime.AddItem "10:50 AM"
        cboStartTime.AddItem "11:00 AM": cboStartTime.AddItem "11:10 AM": cboStartTime.AddItem "11:20 AM": cboStartTime.AddItem "11:30 AM": cboStartTime.AddItem "11:40 AM": cboStartTime.AddItem "11:50 AM"
        cboStartTime.AddItem "12:00 PM": cboStartTime.AddItem "12:10 PM": cboStartTime.AddItem "12:20 PM": cboStartTime.AddItem "12:30 PM": cboStartTime.AddItem "12:40 PM": cboStartTime.AddItem "12:50 PM"
        cboStartTime.AddItem "01:00 PM": cboStartTime.AddItem "01:10 PM": cboStartTime.AddItem "01:20 PM": cboStartTime.AddItem "01:30 PM": cboStartTime.AddItem "01:40 PM": cboStartTime.AddItem "01:50 PM"
        cboStartTime.AddItem "02:00 PM": cboStartTime.AddItem "02:10 PM": cboStartTime.AddItem "02:20 PM": cboStartTime.AddItem "02:30 PM": cboStartTime.AddItem "02:40 pm": cboStartTime.AddItem "02:50 PM"
        cboStartTime.AddItem "03:00 PM": cboStartTime.AddItem "03:10 PM": cboStartTime.AddItem "03:20 PM": cboStartTime.AddItem "03:30 PM": cboStartTime.AddItem "03:40 PM": cboStartTime.AddItem "03:50 PM"
        cboStartTime.AddItem "04:00 PM": cboStartTime.AddItem "04:10 PM": cboStartTime.AddItem "04:20 PM": cboStartTime.AddItem "04:30 PM": cboStartTime.AddItem "04:40 PM": cboStartTime.AddItem "04:50 PM"
        cboStartTime.AddItem "05:00 PM": cboStartTime.AddItem "05:10 PM": cboStartTime.AddItem "05:20 PM": cboStartTime.AddItem "05:30 PM": cboStartTime.AddItem "05:40 PM": cboStartTime.AddItem "05:50 PM"
        cboStartTime.AddItem "06:00 PM": cboStartTime.AddItem "06:10 PM": cboStartTime.AddItem "06:20 PM": cboStartTime.AddItem "06:30 PM": cboStartTime.AddItem "06:40 PM": cboStartTime.AddItem "06:50 PM"
        cboStartTime.AddItem "07:00 PM": cboStartTime.AddItem "07:10 PM": cboStartTime.AddItem "07:20 PM": cboStartTime.AddItem "07:30 PM": cboStartTime.AddItem "07:40 PM": cboStartTime.AddItem "07:50 PM"
        cboStartTime.AddItem "08:00 PM": cboStartTime.AddItem "08:10 PM": cboStartTime.AddItem "08:20 PM": cboStartTime.AddItem "08:30 PM": cboStartTime.AddItem "08:40 PM": cboStartTime.AddItem "08:50 PM"
     End Select
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_ComboTime:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_ComboTime", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Load_Visit()
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
    Screen.MousePointer = vbHourglass
    cboVisitType.Clear
    If Trim(cboResProvider.Text) <> vbNullString Then
        For i = 0 To cboResProvider.listcount - 1
            If cboResProvider.List(i) = cboResProvider.Text Then
                udtRepeatAppt.ProviderID = cboResProvider.ItemData(i)
                Exit For
            End If
        Next i
    End If
    Get_Providers udtRepeatAppt.ProviderID
    strCompare = Get_Provider_Type(Replace(clsProvider.Title, ".", ""))
    strCompare = Replace(strCompare, ".", "")
    strCompare = Trim(strCompare)
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    strSQL = "SELECT DISTINCT APPT_CODE, PROVIDER_TITLE FROM APPOINTMENT_TYPE  ORDER BY PROVIDER_TITLE, APPT_CODE"
    Set rsCodes = cnCodes.Execute(strSQL)
    Do Until rsCodes.EOF
        strProvider = vbNullString
        If Not IsNull(rsCodes.Fields("PROVIDER_TITLE")) Then
            strProvider = UCase(rsCodes.Fields("PROVIDER_TITLE"))
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
                                    cboVisitType.AddItem Trim(Fld.Value)
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
Error_Load_Visit:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Visit", App.EXEName & " " & Me.Name
    Exit Function
    Resume
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


Private Function Format_CellTypeCombo(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, _
vstrString As String, vstrHAlign As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String, Optional vstrHndw As Long)
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
              .TypeHAlign = TypeHAlignCenter
              .TypeVAlign = TypeVAlignCenter
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
        .Redraw = False
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
    .Redraw = True
    End With
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAlign As String, _
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
                .TypeEditMultiLine = True   'Make multi-lines
                .FontName = "Times New Roman"
                .FontSize = 8
                .FontBold = True
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
vstrHAlign As String, Optional vstrCellTag As String, Optional vstrVAlign As String, Optional lngBackColor As Long, Optional vstrCellNote As String)
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

Private Function Get_SchedConflict(vstrDate As String, vstrTime As String) As Boolean
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long

    Get_SchedConflict = False
    On Error GoTo Error_Get_SchedConflict
    Screen.MousePointer = vbHourglass
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "SELECT * FROM APPOINTMENTS_BLOCKS WHERE  PROVIDER_ID = " & udtRepeatAppt.ProviderID & " AND FACILITY_ID = " & udtRepeatAppt.FacilityID & " AND " _
        & "ACTIVE = 1 AND convert(DATETIME, START_DATE, 101) > '" & Format(vstrDate, "mm/dd/yyyy") & "'"
    Set rsBlock = cnTerms.Execute(strSQL)
    With rsBlock
        Do Until rsBlock.EOF
            Set Flds = rsBlock.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case Fld.Name
                            Case "DURATION"
                                udtBlock.Duration = Fld.Value
                            Case "START_TIME"
                                udtBlock.StartTime = Trim(Fld.Value)
                            Case "END_TIME"
                                udtBlock.EndTime = Trim(Fld.Value)
                            Case "START_DATE"
                                udtBlock.StartDate = Fld.Value
                            Case "END_DATE"
                                udtBlock.EndDate = Fld.Value
                        End Select
                End If
            End If
            Next
            rsBlock.MoveNext
            udtBlock.StartTime = Replace(udtBlock.StartTime, ".", "")
            If IsDate(vstrDate) And IsDate(udtBlock.StartDate) Then
                lngDateDiff = DateDiff("d", Format(vstrDate, "mm/dd/yyyy"), Format(udtBlock.StartDate, "mm/dd/yyyy"))
                If lngDateDiff = 0 And udtBlock.Duration <> "Permanent" Then
                    lngDateDiff = DateDiff("d", Format(vstrDate, "mm/dd/yyyy"), Format(udtBlock.EndDate, "mm/dd/yyyy"))
                Else
                    lngDateDiff = -1
                End If
                If lngDateDiff < 0 Then
                  If udtBlock.Duration <> "Permanent" Then
                        lngDateDiff = DateDiff("h", Format(vstrTime, "hh:mm AM/PM"), Format(udtBlock.StartTime, "hh:mm AM/PM"))
                        If lngDateDiff >= 0 Then
                            lngDateDiff = DateDiff("h", Format(vstrTime, "hh:mm AM/PM"), Format(udtBlock.StartTime, "hh:mm AM/PM"))
                        End If
                        If lngDateDiff < 0 Then
                            Get_SchedConflict = True
                            Exit Do
                        End If
                    End If
                ElseIf lngDateDiff = 0 Then
                    Get_SchedConflict = True
                    Exit Do
                End If
            End If
        Loop
    End With
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsBlock = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_SchedConflict:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_SchedConflict", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Update_Row()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim Arr
Dim sApptID As String
Dim strApptDate As String
Dim strApptTime As String
Dim bRowFound As Boolean

    On Error GoTo Error_Update_Row
    Screen.MousePointer = vbHourglass
    udtRepeatAppt.RepeatID = 0
    bRowFound = False
    With grdMultiAppt
        For lngRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Row = lngRow
                If .BackColor = grdCurItem Then
                    bRowFound = True
                    udtRepeatAppt.RepeatID = .GetRowItemData(lngRow)
                    lngCol = 0
                    Arr = Split(LookUpCellTag("APPT_ID"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                           .GetText lngCol, lngRow, varCellValue
                           sApptID = varCellValue
                        End If
                    End If
                    lngCol = 0
                    Arr = Split(LookUpCellTag("APPT_DATE"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            .GetText lngCol, lngRow, varCellValue
                            If IsDate(varCellValue) Then
                                strApptDate = Format(varCellValue, "mm/dd/yyyy")
                            End If
                        End If
                    End If
                    lngCol = 0
                    Arr = Split(LookUpCellTag("APPT_TIME"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            .GetText lngCol, lngRow, varCellValue
                            If IsDate(varCellValue) Then
                                strApptTime = Format(varCellValue, "hh:mm AM/PM")
                            End If
                        End If
                    End If
                    Exit For
                End If
            .BlockMode = False
        Next lngRow
        If bRowFound = False Then
           MsgBox "Select Appointment Row First.", vbInformation, "PatientTrac Scheduling"
           Screen.MousePointer = vbDefault
           Exit Function
        End If
    End With
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    strSQL = "SELECT APPT_DATE, APPT_TIME FROM APPOINTMENTS WHERE APPT_ID = " & sApptID
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
       udtRepeatAppt.StartTime = Format((!Appt_Time), "hh:mm AM/PM")
       udtRepeatAppt.ApptDate = Format((!Appt_Date), "mm/dd/yyyy")
    End With
    If sApptID <> vbNullString Then
        strApptDate = Format(strApptDate, "mm/dd/yyyy")
        strSQL = "UPDATE APPOINTMENTS SET APPT_DATE =  '" & strApptDate & "'," _
            & "APPT_TIME =  '" & strApptTime & "'," _
            & "NOTES = '" & Literals(udtRepeatAppt.Comment, "'") & "'" _
            & " WHERE APPT_ID = " & sApptID
        Set rsCodes = cnCodes.Execute(strSQL)
        MsgBox "Successfully updated Appointment from" & " " & udtRepeatAppt.ApptDate & " " & "at" & " " & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") _
            & " " & "to" & " " & Format(strApptDate, "MMMM DD, YYYY") & " " & "at" & " " & strApptTime
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_Row:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_Row", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Write_Repeat_Appt()
Dim lngProviderID As Long
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim booinsert As Boolean
Dim strReason As String
Dim strConfirm As String
Dim strRepeatAppt As String
Dim nDays As Long
Dim i%
Dim lngApptID  As Long
Dim vTime
Dim Arr
Dim lngResID As Long
Dim booExit As Boolean
Dim strTerms As String

    On Error GoTo Error_Write_Repeat_Appt
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    strSQL = "SELECT * FROM APPOINTMENTS WHERE REPEAT_ID = " & udtRepeatAppt.RepeatID & " AND " _
            & "CONVERT(VARCHAR, APPT_DATE , 10) = '" & udtRepeatAppt.ApptDate & "' "
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        If Not .EOF Then
            lngApptID = (!APPT_ID)
        End If
    End With
    udtRepeatAppt.ApptType = "Patient"
    If Trim(udtRepeatAppt.Duration) = vbNullString Then
        strTerms = Calc_Duration(udtRepeatAppt.PurposeVisit)
    End If
    If lngApptID = 0 Then
        strSQL = "INSERT INTO APPOINTMENTS (PATIENT_ID, FACILITY_ID, PROVIDER_ID, RESOURCE_ID, PURPOSE_VISIT, TERMINATION_DATE, APPT_DATE, " _
            & "APPT_TIME, NOTES, APPT_TYPE, PATIENT_TYPE, DURATION, REPEAT_APPOINT, REPEAT_ID)" _
            & "Values(" & Val(udtSchedPatient.IntId) & ", " & udtRepeatAppt.FacilityID & "," & udtRepeatAppt.ProviderID & "," & udtRepeatAppt.ResourseID & ",'" _
            & udtRepeatAppt.PurposeVisit & "','" & Format(udtRepeatAppt.EndDate, "mm/dd/yyyy") & "','" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "','" _
            & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "','" & udtRepeatAppt.Comment & "','" & udtRepeatAppt.ApptType & "','" _
            & udtAppointment.SchedType & "','" & udtRepeatAppt.Duration & "','" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "'," & udtRepeatAppt.RepeatID & ")"
        Set rsCodes = cnCodes.Execute(strSQL)
    Else
        strSQL = "UPDATE APPOINTMENTS SET PURPOSE_VISIT = '" & udtRepeatAppt.PurposeVisit & "'," _
            & "RESOURCE_ID = " & udtRepeatAppt.ResourseID & ", PROVIDER_ID = " & udtRepeatAppt.ProviderID & "," _
            & "APPT_DATE = '" & Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy") & "',APPT_TIME = '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "'," _
            & "NOTES = '" & udtRepeatAppt.Comment & "', PATIENT_TYPE = '" & udtAppointment.SchedType & "'," _
            & "APPT_TYPE = '" & udtRepeatAppt.ApptType & "', TERMINATION_DATE = '" & Format(udtRepeatAppt.EndDate, "mm/dd/yyyy") & "'," _
            & "DURATION = '" & udtRepeatAppt.Duration & "', REPEAT_APPOINT =  '" & Format(udtRepeatAppt.StartDate, "mm/dd/yyyy") & "', ACTIVE = 1 " _
            & "WHERE APPT_ID = " & lngApptID
        Set rsCodes = cnCodes.Execute(strSQL)
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Repeat_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Repeat_Appt", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cboDayWeek_Click()
    udtRepeatAppt.WeekDay = cboDayWeek.Text
End Sub
Private Sub Clear_Block()
    With udtBlock
        .StartDate = vbNullString
        .EndDate = vbNullString
        .StartTime = vbNullString
        .EndTime = vbNullString
        .Reason = vbNullString
        .ProviderID = 0
        .Duration = vbNullString
        .FacilityID = 0
        .DayWeek = vbNullString
    End With
    Exit Sub
End Sub

Private Sub cboLocation_Click()
Dim Idx%

    For Idx = 0 To cboLocation.listcount - 1
        If cboLocation.List(Idx) = cboLocation.Text Then
            udtRepeatAppt.FacilityID = cboLocation.ItemData(Idx)
            Get_Appt_Providers
            Exit For
        End If
    Next Idx
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


    On Error GoTo Error_Load_Facilities
    Screen.MousePointer = vbHourglass
    cboLocation.Clear
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "SELECT a.DEFAULT_FACILITY, b.FACILITY_NAME, b.FACILITY_ID  FROM APPOINTMENT_PROVIDERS a, " _
        & "FACILITIES b where a.FACILITY_ID = b.FACILITY_ID AND b.ACTIVE = 1 ORDER BY b.FACILITY_NAME"
    Set rsFac = cnFac.Execute(strSQL)
    With rsFac
        Do Until .EOF
            lngFacID = 0
            booDefault = False
            strName = vbNullString
            booSkip = False
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
                                If lngFacID = udtRepeatAppt.FacilityID Then
                                    booDefault = True
                                End If
                        End Select
                    End If
                End If
            Next
            .MoveNext
            strName = Fix_Name(strName)
            strName = Fix_Facility_Name(strName)
            If Trim(strName) <> vbNullString Then
                For Idx = 0 To cboLocation.listcount - 1
                    If TextFound(strName, cboLocation.List(Idx), Len(strName), True) Then
                        booSkip = True
                        Exit For
                    End If
                Next Idx
                If Not booSkip Then
                    cboLocation.AddItem strName
                    If lngFacID > 0 Then cboLocation.ItemData(cboLocation.NewIndex) = lngFacID
                End If
            End If
            If udtRepeatAppt.FacilityID <= 0 And lngFacID > 0 Then
                udtRepeatAppt.FacilityID = lngFacID
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


Private Sub cboRepAppt_Change()
    If cboRepAppt.Text = "Day" Then
       cboDayWeek.ListIndex = -1
       cboDayWeek.Enabled = False
       ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm) - 1, ccdStart.Value)
       ccdTermDate.Enabled = False
    Else
       If cboDayWeek.Enabled = False Then cboDayWeek.Enabled = True
       If ccdTermDate.Enabled = False Then ccdTermDate.Enabled = True
    End If
End Sub


Private Sub cboRepAppt_Click()
    If cboRepAppt.Text = "Day" Then
       cboDayWeek.ListIndex = -1
       cboDayWeek.Enabled = False
       ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm) - 1, ccdStart.Value)
       ccdTermDate.Enabled = False
    Else
       If cboDayWeek.Enabled = False Then cboDayWeek.Enabled = True
       If ccdTermDate.Enabled = False Then ccdTermDate.Enabled = True
    End If
End Sub


Private Sub cboRepeatTerm_Change()
    If cboRepAppt.Text = "Day" Then
       ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm) - 1, ccdStart.Value)
    End If
End Sub


Private Sub cboRepeatTerm_Click()
    If cboRepAppt.Text = "Day" Then
       ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm) - 1, ccdStart.Value)
    End If
End Sub


Private Sub cboResProvider_Click()
    Dim Idx%
    If Trim(cboResProvider.Text) <> vbNullString Then
        For Idx = 0 To cboResProvider.listcount - 1
            If cboResProvider.List(Idx) = cboResProvider.Text Then
                udtRepeatAppt.ProviderID = cboResProvider.ItemData(Idx)
                Load_Visit
                Exit For
            End If
        Next Idx
    End If
End Sub


Private Sub cboStartTime_Click()
    If cboStartTime.Text <> vbNullString Then
        udtRepeatAppt.StartTime = Format(cboStartTime.Text, "hh:mm AM/PM")
    End If
End Sub


Private Sub ccdStart_Change()
    udtRepeatAppt.StartDate = Format(ccdStart.Value, "mm/dd/yyyy")
    If cboRepAppt.Text = "Day" Then
       ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm) - 1, ccdStart.Value)
    End If
    Exit Sub
End Sub

Private Sub ccdStart_Click()
    udtRepeatAppt.StartDate = Format(ccdStart.Value, "mm/dd/yyyy")
    If cboRepAppt.Text = "Day" Then
       ccdTermDate.Value = DateAdd("d", Val(cboRepeatTerm) - 1, ccdStart.Value)
    End If
    Exit Sub
End Sub





Private Sub cctComments_KeyUp(KeyCode As Integer, Shift As Integer)
    udtRepeatAppt.Comment = cctComments.Text
End Sub


Private Sub cmdAccept_Click()
Dim strDate As String
Dim varCellValue
Dim lngRow As Long
Dim lngCol As Long

    With grdMultiAppt
        clsGridFormat.CellTag = vbNullString
        .Col = mlngGridCol
        .Row = mlngGridRow
        If .CellTag <> vbNullString Then clsGridFormat.CellTag = .CellTag
        If Trim(clsGridFormat.CellTag) <> vbNullString Then
            Select Case .CellTag
                Case "APPT_DATE"
                    strDate = Format(ccdDate.Value, "mm/dd/yyyy")
                    Format_CellTypeEdit grdMultiAppt, mlngGridRow, mlngGridCol, "1", strDate, "TypeHAlignCenter", clsGridFormat.CellTag
                    grdMultiAppt.SetCellBorder mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "APPT_TIME"
                    If cboRepTime.ListIndex <> -1 Then
                       strDate = cboRepTime.Text
                       Format_CellTypeEdit grdMultiAppt, mlngGridRow, mlngGridCol, "1", strDate, "TypeHAlignCenter", clsGridFormat.CellTag
                       grdMultiAppt.SetCellBorder mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
            End Select
        End If
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = mlngGridRow
            .Row2 = mlngGridRow
            .BackColor = grdCurItem
        .BlockMode = False
    End With
    fraNumber.Visible = False
    Exit Sub
End Sub

Private Function LookUpCellTag(vstrName As String, Optional vlngRow As Long) As String
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String
    
    lngCol = 0
    With grdMultiAppt
        LookUpCellTag = vbNullString
        For lngCol = 1 To .MaxCols
            .Col = lngCol
            .Row = SpreadHeader + 1
            If .CellTag <> vbNullString Then strTag = .CellTag
            If Trim(strTag) <> vbNullString And TextFound(vstrName, Trim(strTag), Len(vstrName), True) Then
                LookUpCellTag = lngCol & "|" & lngRow
                Exit For
            End If
        Next lngCol
    End With
    Exit Function
End Function
Private Function LookUpCellTag_RepeatHist(vstrName As String, Optional vlngRow As Long) As String
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String
    
    lngCol = 0
    With grdRepeatHist
        LookUpCellTag_RepeatHist = vbNullString
        For lngCol = 1 To .MaxCols
            .Col = lngCol
            .Row = SpreadHeader + 1
            If .CellTag <> vbNullString Then strTag = .CellTag
            If Trim(strTag) <> vbNullString And TextFound(vstrName, Trim(strTag), Len(vstrName), True) Then
                LookUpCellTag_RepeatHist = lngCol & "|" & lngRow
                Exit For
            End If
        Next lngCol
    End With
    Exit Function
End Function

Private Function LookUpCellTag_Header(vstrName As String, Optional vlngRow As Long) As String
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String

    lngCol = 0
    With grdMultiAppt
        LookUpCellTag_Header = vbNullString
            For lngCol = 1 To .MaxCols
                    .Col = lngCol
                    .Row = SpreadHeader + 1
                    If .CellTag <> vbNullString Then
                        If TextFound(vstrName, .CellTag, Len(vstrName), True) Then
                            LookUpCellTag_Header = lngCol & "|" & lngRow
                            Exit For
                        End If
                    End If
            Next lngCol
    End With
    Exit Function
End Function

Private Sub cmdConflict_Click()
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please Select Patient First!", vbInformation, "PatientTrac Repeat Appointments"
    Else
        lblPatName.Visible = True
        lblPatName.Caption = udtSchedPatient.Name
    End If
End Sub

Private Sub cmdExit_Click()
    fraControl.Visible = False
    Format_Repeat_History
    Load_Repeat_Appt
End Sub

Private Sub cmdMultiClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdMultiSave_Click()
Dim Arr
Dim lngCol As Long
Dim varCellValue

    With grdMultiAppt
        Arr = Split(LookUpCellTag("APPT_DATE", mlngGridRow))
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
        End If
        If lngCol > 0 Then .GetText lngCol, mlngGridRow, varCellValue
        If Trim(varCellValue) <> vbNullString And lngCol > 0 Then
             udtRepeatAppt.StartDate = Format(varCellValue, "mm/dd/yyyy")
        End If
        varCellValue = vbNullString
        lngCol = 0
        Arr = Split(LookUpCellTag("APPT_TIME", mlngGridRow))
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
        End If
        If lngCol > 0 Then .GetText lngCol, mlngGridRow, varCellValue
        If Trim(varCellValue) <> vbNullString And lngCol > 0 Then
             udtRepeatAppt.StartTime = varCellValue
        End If
        lngCol = 0
        Arr = Split(LookUpCellTag("REPEAT_DAY_WEEK", mlngGridRow))
        If UBound(Arr) > 0 Then
            lngCol = Arr(0)
        End If
        If lngCol > 0 Then .GetText lngCol, mlngGridRow, varCellValue
        If Trim(varCellValue) <> vbNullString And lngCol > 0 Then
             udtRepeatAppt.WeekDay = varCellValue
        End If
        If Trim(udtRepeatAppt.StartDate) <> vbNullString And Trim(udtRepeatAppt.StartTime) <> vbNullString Then
            Save_RepeatData udtRepeatAppt.StartTime, udtRepeatAppt.StartDate
        End If
        .BlockMode = True
           .Col = 1
           .Col2 = 5
           .Row = mlngGridRow
           .Row2 = mlngGridRow
           .BackColor = NormalColor
        .BlockMode = False
    End With
    Exit Sub
End Sub

Private Sub cmdNewRepeat_Click()
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please Select Patient First!", vbInformation, "PatientTrac Repeat Appointments"
       Exit Sub
    End If
    Load_Form
    Load_ComboTime
    fraControl.Visible = True
End Sub
Private Sub cmdPatient_Click()
Dim i%

    On Error GoTo Error_Select_Patient
    Screen.MousePointer = vbHourglass
    udtSchedPatient.IntId = vbNullString
    Clear_SchedPatient
    grdMultiAppt.ClearRange 1, 1, grdMultiAppt.MaxRows, grdMultiAppt.ColsFrozen, False
    grdMultiAppt.MaxRows = 0
    frmListPatient.mlngListItemData = 0
    frmListPatient.Show vbModal
    If frmListPatient.mlngListItemData <> 0 Then
        udtSchedPatient.IntId = Val(frmListPatient.mlngListItemData)
        Get_SchedPatient Val(udtSchedPatient.IntId)
        Format_MultiAppt
        Format_Repeat_History
        Load_Repeat_Appt
        Read_MultiAppt Val(udtSchedPatient.IntId)
        lblPatName.Caption = udtSchedPatient.Name
        lblPatName.Visible = True
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Select_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, udtSchedPatient.Name, "Select_Patient", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
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
    Add_Header
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Multi Appointment Schedule"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    With grdMultiAppt
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Multiple Appointment"
        .PrintColor = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintSmartPrint = True
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Multiple Appointments"
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

Private Sub cmdProvTypes_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Provider Types"
    frmNurseTemplates.fType = "Provider Types"
    frmNurseTemplates.fSubDomain = "sd_Provider_ProviderType"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    Exit Sub
End Sub


Private Sub cmdRemoveRow_Click()
Dim lngRow As Long
Dim lngRepId As Long
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim lngDateDiff As Long
Dim strTermDate As String
Dim strApptDate As String
Dim varCellValue
Dim booFound As Boolean
Dim Arr
Dim lngCol As Long

    On Error GoTo Error_cmdRemoveRow
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please Select Patient First!", vbInformation, "PatientTrac Repeat Appointments"
    Else
        lblPatName.Visible = True
        lblPatName.Caption = udtSchedPatient.Name
    End If
    With grdMultiAppt
        For lngRow = 1 To grdMultiAppt.MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                If .BackColor = grdCurItem Then
                    booFound = True
                    Exit For
                End If
            .BlockMode = False
        Next lngRow
    End With
    If booFound Then
        lngCol = 0
        Arr = Split(LookUpCellTag("APPT_DATE"), "|")
        If UBound(Arr) >= 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                grdMultiAppt.GetText 2, lngRow, varCellValue
            End If
        End If
        If IsDate(varCellValue) Then
            strTermDate = Format(varCellValue, "mm/dd/yyyy")
        End If
    Else
       MsgBox "Please Select Terminate Date First!", vbInformation, "PatientTrac Repeat Appointments"
        Exit Sub
    End If
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please Select Patient First!", vbInformation, "PatientTrac Repeat Appointments"
    Else
        lblPatName.Visible = True
        lblPatName.Caption = udtSchedPatient.Name
    End If
    Screen.MousePointer = vbHourglass
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    If Not IsDate(strTermDate) Then Exit Sub
    For lngRow = lngRow To grdMultiAppt.MaxRows
        lngRepId = grdMultiAppt.GetRowItemData(lngRow)
        If lngRepId > 0 Then
          strSQL = "UPDATE APPOINTMENTS SET TERMINATION_DATE = '" & Format(strTermDate, "mm/dd/yyyy") & "'," _
            & "ACTIVE = 0, STATUS = 'Cancelled' WHERE REPEAT_ID = " & lngRepId & " AND convert(varchar, APPT_DATE,10) > '" & Format(strTermDate, "mm/dd/yyyy") & "'"
            Set rsAppt = cnAppt.Execute(strSQL)
          strSQL = "UPDATE PATIENT_REPEAT_APPT SET TERMINATION_DATE = '" & Format(strTermDate, "mm/dd/yyyy") & "'," _
            & "ACTIVE = 0 WHERE REPEAT_ID = " & lngRepId
            Set rsAppt = cnAppt.Execute(strSQL)
            Exit For
        End If
    Next lngRow
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Format_MultiAppt
    Format_Repeat_History
    Load_Repeat_Appt
    Screen.MousePointer = vbDefault
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Load_Repeat_Appt
    MsgBox "Selected Repeat Appt has been successfully removed!", vbInformation, "PatientTrac Scheduling"
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdRemoveRow:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdRemoveRow", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub cmdResource_Click()
    frmResource.Show vbModal
    Exit Sub
End Sub

Private Sub cmdSched_Click()



    On Error GoTo Error_cmdSched
    Screen.MousePointer = vbHourglass
    If cboLocation.Text = vbNullString Then
        MsgBox "Appointment requires a facility! Please select a facility.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If cboVisitType.ListIndex = -1 Then
       MsgBox "Please Select Appt Type.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If cboRepeatTerm.ListIndex = -1 Then
       MsgBox "Please Select Occurance Interval (Every).", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If cboRepAppt.ListIndex = -1 Then
       MsgBox "Please Select Term Value.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If cboDayWeek.ListIndex = -1 And cboRepAppt.Text <> "Day" Then
        MsgBox "Please Select Day Of Week.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If cboStartTime.Text = vbNullString Then
        MsgBox "Please Confirm Appt Time.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If IsNull(ccdTermDate.Value) Or ccdTermDate.Value = vbNullString Then
        MsgBox "Please Confirm End Date.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If IsNull(ccdStart.Value) Then
       MsgBox "Please enter Start Date.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If IsNull(ccdTermDate.Value) Then
        MsgBox "Please enter End Date.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If DateDiff("d", Now, ccdStart.Value) <= 0 Then
        MsgBox "Start Date Must Be Greater Than Current Date.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If DateDiff("d", ccdStart.Value, ccdTermDate.Value) < 0 Then
        MsgBox "End Date Must Be Greater Than Or Equal To Start Date.", vbInformation, "PatientTrac Scheduling"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    grdMultiAppt.ClearRange 1, 1, grdMultiAppt.MaxCols, grdMultiAppt.MaxRows, False
    grdMultiAppt.MaxRows = 0
    If cboStartTime.Text <> vbNullString Then
        udtRepeatAppt.StartTime = Format(cboStartTime.Text, "hh:mm AM/PM")
    End If
    If udtRepeatAppt.StartTime <> vbNullString Then
        Repeat_Appt Format(udtRepeatAppt.StartTime, "hh:mm AM/PM")
    End If
    MsgBox "Repeat Appointments has been successfully saved!", vbInformation, "PatientTrac Scheduling"
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdSched:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdSched", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Function Repeat_Appt(vstrTime As String)
Dim strSQL As String
Dim nDays As Long
Dim Arr
Dim booinsert As Boolean
Dim RepeatIndex As Integer
Dim strRepDate As String
Dim Idx As Integer
Dim lngDateDiff As Long
Dim lngPatientID As Long
Dim FirstDay As VbDayOfWeek
Dim lngWeekDays As Long
Dim strOption As String
Dim strMonth As String
Dim strTestMonth As String
Dim lngEndDate As Long

    On Error GoTo Error_Repeat_Appt
    Screen.MousePointer = vbHourglass
    cctBlockConflict.Text = vbNullString
    If Val(udtSchedPatient.IntId) >= 0 Then
        lngPatientID = Val(udtSchedPatient.IntId)
    Else
        MsgBox "Please select a patient first!", vbInformation, "PatientTrac Repeat Appointments Schedule"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If Not IsNull(ccdTermDate.Value) Then
        udtRepeatAppt.EndDate = Format(ccdTermDate.Value, "mm/dd/yyyy")
    End If
    If Not IsNull(ccdStart.Value) Then
        udtRepeatAppt.StartDate = Format(ccdStart.Value, "mm/dd/yyyy")
    End If
    nDays = cboRepeatTerm.Text
    If udtRepeatAppt.WeekDay = vbNullString Then
        udtRepeatAppt.WeekDay = Format(udtAppointment.SchedDate, "dddd")
    End If
    If Trim(cctComments.Text) <> vbNullString Then
        udtRepeatAppt.Comment = Trim(cctComments.Text)
    End If
    If Trim(cboVisitType.Text) <> vbNullString Then
        udtRepeatAppt.PurposeVisit = cboVisitType.Text
    End If
    udtRepeatAppt.ApptType = "Patient"
    ' Need Duration
    udtRepeatAppt.Duration = Calc_Duration(udtRepeatAppt.PurposeVisit)
    If cboRepAppt.Text = "Months" Then
        
    End If
    If nDays >= 1 Then
        Select Case cboRepAppt.Text
            Case "Day"
                udtRepeatAppt.ApptDate = vbNullString
'                lngDateDiff = DateDiff("d", udtAppointment.SchedDate, udtRepeatAppt.EndDate)
                lngDateDiff = (DateDiff("d", udtRepeatAppt.StartDate, udtRepeatAppt.EndDate)) + 1
'                lngDateDiff = lngDateDiff / ndays
                If lngDateDiff >= 0 Then
                    strRepDate = Format(udtRepeatAppt.StartDate, "mm/dd/yyyy")
                    For Idx = 0 To lngDateDiff
                        If Trim(strRepDate) <> vbNullString Then
                            udtRepeatAppt.ApptDate = strRepDate
                            lngEndDate = DateDiff("d", strRepDate, udtRepeatAppt.EndDate)
                            If lngEndDate >= 0 Then
                                udtRepeatAppt.ApptDate = strRepDate
                                Save_RepeatData vstrTime, udtRepeatAppt.ApptDate
                                If Get_SchedConflict(strRepDate, vstrTime) Then
                                    If cctBlockConflict.Text = vbNullString Then
                                        cctBlockConflict.Text = "These Appointments conflict with Block Times" & vbNewLine & strRepDate & " " & vstrTime
                                    Else
                                        cctBlockConflict.Text = cctBlockConflict.Text & vbNewLine & cctBlockConflict.Text
                                    End If
                                    lblBlock.Visible = True
                                    cctBlockConflict.Visible = True
                                End If
                            End If
                        End If
                        strRepDate = DateAdd("d", Idx, strRepDate)
                    Next Idx
                End If
            Case "Week"
                udtRepeatAppt.ApptDate = vbNullString
                lngDateDiff = DateDiff("ww", udtAppointment.SchedDate, udtRepeatAppt.EndDate)
                lngDateDiff = lngDateDiff / nDays
                If lngDateDiff > 0 Then
                    If Trim(udtRepeatAppt.ApptDate) <> vbNullString Then
                        strRepDate = Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy")
                    Else
                        strRepDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
                    End If
                    For Idx = 0 To lngDateDiff
                        If strRepDate <> vbNullString Then
                            Select Case udtRepeatAppt.WeekDay
                                Case "Monday"
                                    FirstDay = WeekDay(strRepDate, vbMonday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Tuesday"
                                    FirstDay = WeekDay(strRepDate, vbTuesday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Wednesday"
                                    FirstDay = WeekDay(strRepDate, vbWednesday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Thursday"
                                    FirstDay = WeekDay(strRepDate, vbThursday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Friday"
                                    FirstDay = WeekDay(strRepDate, vbFriday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Saturday"
                                    FirstDay = WeekDay(strRepDate, vbSaturday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Sunday"
                                    FirstDay = WeekDay(strRepDate, vbSunday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                            End Select
                            If Trim(strRepDate) <> vbNullString Then
                                udtRepeatAppt.ApptDate = strRepDate
                                lngEndDate = DateDiff("d", strRepDate, udtRepeatAppt.EndDate)
                                If lngEndDate >= 0 Then
                                    udtRepeatAppt.ApptDate = strRepDate
                                    Save_RepeatData vstrTime, udtRepeatAppt.ApptDate
                                    If Get_SchedConflict(strRepDate, vstrTime) Then
                                        If cctBlockConflict.Text = vbNullString Then
                                            cctBlockConflict.Text = "These Appointments conflict with Block Times" & vbNewLine & strRepDate & " " & vstrTime
                                        Else
                                            cctBlockConflict.Text = cctBlockConflict.Text & vbNewLine & cctBlockConflict.Text
                                        End If
                                        lblBlock.Visible = True
                                        cctBlockConflict.Visible = True
                                    End If
                                End If
                            End If
                        End If
                        strRepDate = DateAdd("ww", nDays, strRepDate)
                    Next Idx
                End If
            Case "Month"
                udtRepeatAppt.ApptDate = vbNullString
                lngDateDiff = DateDiff("m", udtAppointment.SchedDate, udtRepeatAppt.EndDate)
                lngDateDiff = lngDateDiff / nDays
                udtRepeatAppt.RepeatNumber = lngDateDiff
                If lngDateDiff > 0 Then
                    If Trim(udtRepeatAppt.ApptDate) <> vbNullString Then
                        strRepDate = Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy")
                    Else
                        strRepDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
                    End If
                    For Idx = 0 To lngDateDiff
                        If strRepDate <> vbNullString Then
                            Select Case udtRepeatAppt.WeekDay
                                Case "Monday"
                                    FirstDay = WeekDay(strRepDate, vbMonday)
                                    strRepDate = DateAdd("d", -FirstDay, strRepDate)
                                    strTestMonth = Month(strRepDate)
                                Case "Tuesday"
                                    FirstDay = WeekDay(strRepDate, vbTuesday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                    strTestMonth = Month(strRepDate)
                                Case "Wednesday"
                                    FirstDay = WeekDay(strRepDate, vbWednesday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Thursday"
                                    FirstDay = WeekDay(strRepDate, vbThursday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                    strTestMonth = Month(strRepDate)
                                Case "Friday"
                                    FirstDay = WeekDay(strRepDate, vbFriday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                    strTestMonth = Month(strRepDate)
                                Case "Saturday"
                                    FirstDay = WeekDay(strRepDate, vbSaturday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                    strTestMonth = Month(strRepDate)
                                Case "Sunday"
                                    FirstDay = WeekDay(strRepDate, vbSunday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                    strTestMonth = Month(strRepDate)
                            End Select
                            If Trim(strRepDate) <> vbNullString Then
                                lngEndDate = DateDiff("d", strRepDate, udtRepeatAppt.EndDate)
                                If lngEndDate >= 0 Then
                                    udtRepeatAppt.ApptDate = strRepDate
                                    Save_RepeatData vstrTime, udtRepeatAppt.ApptDate
                                    If Get_SchedConflict(strRepDate, vstrTime) Then
                                        If cctBlockConflict.Text = vbNullString Then
                                            cctBlockConflict.Text = "These Appointments conflict with Block Times" & vbNewLine & strRepDate & " " & vstrTime
                                        Else
                                            cctBlockConflict.Text = cctBlockConflict.Text & vbNewLine & cctBlockConflict.Text
                                        End If
                                        lblBlock.Visible = True
                                        cctBlockConflict.Visible = True
                                    End If
                                End If
                            End If
                        End If
                        strRepDate = DateAdd("m", 1, strRepDate)
                        strMonth = Month(strRepDate)
                    Next Idx
                End If
            Case "Annual"
                udtRepeatAppt.ApptDate = vbNullString
                lngDateDiff = DateDiff("yyyy", udtAppointment.SchedDate, udtRepeatAppt.EndDate)
                lngDateDiff = lngDateDiff / nDays
                If lngDateDiff > 0 Then
                    If Trim(udtRepeatAppt.ApptDate) <> vbNullString Then
                        strRepDate = Format(udtRepeatAppt.ApptDate, "mm/dd/yyyy")
                    Else
                        strRepDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
                    End If
                    For Idx = 1 To lngDateDiff
                        strRepDate = DateAdd("yyyy", Idx, strRepDate)
                        If strRepDate <> vbNullString Then
                            Select Case udtRepeatAppt.WeekDay
                                Case "Monday"
                                    FirstDay = WeekDay(strRepDate, vbMonday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Tuesday"
                                    FirstDay = WeekDay(strRepDate, vbTuesday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Wednesday"
                                    FirstDay = WeekDay(strRepDate, vbWednesday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Thursday"
                                    FirstDay = WeekDay(strRepDate, vbThursday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Friday"
                                    FirstDay = WeekDay(strRepDate, vbFriday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Saturday"
                                    FirstDay = WeekDay(strRepDate, vbSaturday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                                Case "Sunday"
                                    FirstDay = WeekDay(strRepDate, vbSunday)
                                    strRepDate = DateAdd("d", -(FirstDay - 1), strRepDate)
                            End Select
                            If Trim(strRepDate) <> vbNullString Then
                                lngEndDate = DateDiff("d", strRepDate, udtRepeatAppt.EndDate)
                                If lngEndDate <= 0 Then
                                    If Get_SchedConflict(strRepDate, vstrTime) Then
                                        If cctBlockConflict.Text = vbNullString Then
                                            cctBlockConflict.Text = "These Appointments conflict with Block Times" & vbNewLine & strRepDate & " " & vstrTime
                                        Else
                                            cctBlockConflict.Text = cctBlockConflict.Text & vbNewLine & cctBlockConflict.Text
                                        End If
                                        lblBlock.Visible = True
                                        cctBlockConflict.Visible = True
                                    End If
                                    udtRepeatAppt.ApptDate = strRepDate
                                    Save_RepeatData vstrTime, udtRepeatAppt.ApptDate
                                End If
                            End If
                        End If
                    Next Idx
                End If
        End Select
    End If
    Read_MultiAppt lngPatientID
    Screen.MousePointer = vbDefault
    Exit Function
Error_Repeat_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Repeat_Appt", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Save_RepeatData(vstrTime As String, vstrDate As String) As Boolean
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim strTerms As String

    On Error GoTo Error_Save_RepeatData
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(0)
    vstrDate = Format(vstrDate, "mm/dd/yyyy")
    udtRepeatAppt.StartTime = vstrTime
    If udtRepeatAppt.RepeatMode = vbNullString Then
        udtRepeatAppt.RepeatMode = cboRepAppt.Text
    End If
    udtRepeatAppt.StartTime = Format(udtRepeatAppt.StartTime, "hh:mm AM/PM")
    udtRepeatAppt.RepeatID = 0
    strTerms = Calc_Duration(udtRepeatAppt.PurposeVisit)
    strSQL = "SELECT * FROM PATIENT_REPEAT_APPT where PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND FACILITY_ID = " & udtRepeatAppt.FacilityID & " AND " _
            & "convert(VARCHAR, START_DATE, 10) = '" & Format(udtRepeatAppt.StartDate, "mm/dd/yyyy") & "' AND START_TIME = '" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "' AND PROVIDER_ID = " & udtRepeatAppt.ProviderID
    Set rsCodes = cnCodes.Execute(strSQL)
    With rsCodes
        If Not .EOF Then
            udtRepeatAppt.RepeatID = (!REPEAT_ID)
        End If
    End With
    If udtRepeatAppt.RepeatID = 0 Then
        strSQL = "INSERT INTO PATIENT_REPEAT_APPT (PATIENT_ID, FACILITY_ID, PROVIDER_ID, " _
            & "RESOURCE_ID, APPT_TIME, APPT_TYPE, START_TIME, END_TIME, START_DATE, TERMINATION_DATE, " _
            & "REPEAT_FREQ, DURATION, REPEAT_MODE, REPEAT_DAY_WEEK, COMMENT, ACTIVE ) " _
            & "Values(" & Val(udtSchedPatient.IntId) & ", " & udtRepeatAppt.FacilityID & "," & udtRepeatAppt.ProviderID & ", " _
            & udtRepeatAppt.ResourseID & ",'" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "','" & udtRepeatAppt.ApptType & "','" & Format(udtRepeatAppt.StartTime, "hh:mm AM/PM") & "','" _
            & Format(udtRepeatAppt.EndTime, "hh:mm AM/PM") & "','" & udtRepeatAppt.StartDate & "','" & udtRepeatAppt.EndDate & "','" _
            & udtRepeatAppt.RepeatNumber & "','" & udtRepeatAppt.Duration & "','" & udtRepeatAppt.RepeatMode & "','" _
            & udtRepeatAppt.WeekDay & "','" & udtRepeatAppt.Comment & "', 1)"
        Set rsCodes = cnCodes.Execute(strSQL)
        strSQL = "SELECT * FROM PATIENT_REPEAT_APPT where PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND FACILITY_ID = " & udtRepeatAppt.FacilityID & " AND " _
                & "START_DATE = '" & udtRepeatAppt.StartDate & "' AND PROVIDER_ID = " & udtRepeatAppt.ProviderID
        Set rsCodes = cnCodes.Execute(strSQL)
        With rsCodes
            If Not .EOF Then
                udtRepeatAppt.RepeatID = (!REPEAT_ID)
            End If
        End With
    Else
        strSQL = "UPDATE PATIENT_REPEAT_APPT SET APPT_TIME = '" & udtRepeatAppt.StartTime & "', " _
            & "RESOURCE_ID = " & udtRepeatAppt.ResourseID & ", PROVIDER_ID = " & udtRepeatAppt.ProviderID & "," _
            & "APPT_TYPE = '" & udtRepeatAppt.ApptType & "',START_TIME = '" & udtRepeatAppt.StartTime & "', " _
            & "END_TIME = '" & udtRepeatAppt.EndTime & "', START_DATE = '" & udtRepeatAppt.StartDate & "'," _
            & "TERMINATION_DATE = '" & udtRepeatAppt.EndDate & "', REPEAT_FREQ = '" & udtRepeatAppt.RepeatNumber & "', " _
            & "DURATION = '" & udtRepeatAppt.Duration & "', REPEAT_MODE = '" & udtRepeatAppt.RepeatMode & "', " _
            & "REPEAT_DAY_WEEK = '" & udtRepeatAppt.WeekDay & "', COMMENT = '" & udtRepeatAppt.Comment & "' WHERE " _
            & "REPEAT_ID = " & udtRepeatAppt.RepeatID
        Set rsCodes = cnCodes.Execute(strSQL)
    End If
    Write_Repeat_Appt
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Save_RepeatData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Save_RepeatData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Read_MultiAppt(Optional vlngPatientID As Long)
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim i%
Dim lngPatientID As Long
Dim lngFacilityID As Long
Dim lngResID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim strRes As String
Dim strProv As String
Dim lngDateDiff As Long
Dim strTerm As String
Dim strStart As String
Dim strApptDate As String
Dim Arr

    On Error GoTo Error_Read_MultiAppt
    Screen.MousePointer = vbHourglass
    grdMultiAppt.MaxRows = 0
    If cboResProvider.listcount < 0 Then Get_Appt_Providers
    If cnSched.State = adStateClosed Then cnSched.open psConnect(0)
    udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    strSQL = "SELECT  PATIENT_REPEAT_APPT.*, APPOINTMENTS.* FROM PATIENT_REPEAT_APPT " _
        & "INNER JOIN APPOINTMENTS ON APPOINTMENTS.REPEAT_ID = PATIENT_REPEAT_APPT.REPEAT_ID " _
        & "WHERE PATIENT_REPEAT_APPT.ACTIVE = 1 AND PATIENT_REPEAT_APPT.REPEAT_ID = " & udtRepeatAppt.RepeatID
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            Do Until .EOF
                Clear_RepeatAppt
                lngDateDiff = 0
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "APPT_ID"
                                    udtRepeatAppt.ApptID = Trim(Fld.Value)
                                Case "REPEAT_ID"
                                    udtRepeatAppt.RepeatID = Trim(Fld.Value)
                                Case "APPT_DATE"
                                    udtRepeatAppt.ApptDate = Trim(Fld.Value)
                                Case "PROVIDER_ID"
                                    udtRepeatAppt.ProviderID = Trim(Fld.Value)
                                Case "RESOURCE_ID"
                                    udtRepeatAppt.ResourseID = Trim(Fld.Value)
                                Case "FACILITY_ID"
                                    udtRepeatAppt.FacilityID = Trim(Fld.Value)
                                Case "APPT_TIME"
                                    udtRepeatAppt.StartTime = Trim(Fld.Value)
                                Case "PURPOSE_VISIT"
                                    udtRepeatAppt.PurposeVisit = Trim(Fld.Value)
                                Case "DURATION"
                                    udtRepeatAppt.Duration = Trim(Fld.Value)
                                Case "REPEAT_FREQ"
                                    udtRepeatAppt.RepeatNumber = Trim(Fld.Value)
                                Case "TERMINATION_DATE"
                                    udtRepeatAppt.EndDate = Trim(Fld.Value)
                                Case "START_DATE"
                                    udtRepeatAppt.StartDate = Trim(Fld.Value)
                                Case "START_TIME"
                                    udtRepeatAppt.StartTime = Trim(Fld.Value)
                                Case "END_TIME"
                                    udtRepeatAppt.EndTime = Trim(Fld.Value)
                                Case "REPEAT_MODE"
                                    udtRepeatAppt.RepeatMode = Trim(Fld.Value)
                                Case "REPEAT_DAY_WEEK"
                                    udtRepeatAppt.WeekDay = Trim(Fld.Value)
                                Case "NOTES"
                                    udtRepeatAppt.Comment = Trim(Literals(Fld.Value, "'"))
                            End Select
                        End If
                    End If
                Next
                .MoveNext
                If Trim(udtRepeatAppt.ApptDate) <> vbNullString Then
                    lngRow = grdMultiAppt.MaxRows + 1
                    grdMultiAppt.MaxRows = lngRow
                    grdMultiAppt.SetRowItemData lngRow, udtRepeatAppt.RepeatID
                    lngCol = 0
                    Arr = Split(LookUpCellTag("APPT_DATE"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.ApptDate, "TypeHAlignCenter", "APPT_DATE"
                            grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                    End If
                    grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 0
                     Arr = Split(LookUpCellTag("FACILITY_ID"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If udtRepeatAppt.FacilityID > 0 Then
                            Get_ReportFacility udtRepeatAppt.FacilityID
                        End If
                        Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtFacility.Name, "TypeHAlignCenter", "FACILITY_ID"
                        grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                    grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 0
                     Arr = Split(LookUpCellTag("APPT_TIME"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.StartTime, "TypeHAlignCenter", "APPT_TIME"
                            grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                    End If
                    grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 0
                     Arr = Split(LookUpCellTag("START_DATE"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.StartDate, "TypeHAlignCenter", "START_DATE"
                            grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                    End If
                    grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 0
                     Arr = Split(LookUpCellTag("DURATION"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.Duration, "TypeHAlignCenter", "DURATION"
                            grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                    End If
                    grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 0
                     Arr = Split(LookUpCellTag("REPEAT_DAY_WEEK"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.WeekDay, "TypeHAlignCenter", "REPEAT_DAY_WEEK"
                            grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                    End If
                    grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 0
                     Arr = Split(LookUpCellTag("APPT_TYPE"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.PurposeVisit, "TypeHAlignCenter", "APPT_TYPE"
                            grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                    End If
                    grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 0
                     Arr = Split(LookUpCellTag("TERMINATION_DATE"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.EndDate, "TypeHAlignCenter", "TERMINATION_DATE"
                            grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                    End If
                    grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 0
                     Arr = Split(LookUpCellTag("NOTES"), "|")
                    If UBound(Arr) >= 0 Then
                        lngCol = Arr(0)
                        If lngCol > 0 Then
                            Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", udtRepeatAppt.Comment, "TypeHAlignCenter", "NOTES"
                            grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                    End If
                    grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngCol = 0
                    Arr = Split(LookUpCellTag("APPT_ID"), "|")
                    If UBound(Arr) >= 0 Then
                       lngCol = Arr(0)
                       If lngCol > 0 Then
                          Format_CellTypeEdit grdMultiAppt, lngRow, lngCol, "1", CStr(udtRepeatAppt.ApptID), "TypeHAlignCenter", "APPT_ID"
                          grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                       End If
                    End If
                    grdMultiAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
            Loop
        End With
    If cnSched.State = adStateOpen Then
        cnSched.Close
        Set cnSched = Nothing
        Set rsSched = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_MultiAppt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_MultiAppt", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cmdUpdate_Click()
    If Val(udtSchedPatient.IntId) = 0 Then
       MsgBox "Please Select Patient First!", vbInformation, "PatientTrac Repeat Appointments"
    Else
        lblPatName.Visible = True
        lblPatName.Caption = udtSchedPatient.Name
    End If
    Update_Row
    Exit Sub
End Sub
Private Sub Form_Activate()
Dim sValue As String
Dim Idx As Integer

    Exit Sub
End Sub

Private Sub Form_Load()
    
    lblPatName.Caption = vbNullString
    Clear_SchedPatient
    Clear_Block
    Clear_RepeatAppt
    ccdTermDate.Value = vbNullString
    ccdStart.Value = vbNullString
    Load_Form
    centerForm Me
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

  On Error Resume Next

  Unload Me
  Exit Sub
End Sub
Private Function Load_Repeat_Appt()
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim i%
Dim lngPatientID As Long
Dim lngFacilityID As Long
Dim lngResID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim strRes As String
Dim strProv As String
Dim lngDateDiff As Long
Dim strTerm As String
Dim strStart As String
Dim strApptDate As String
Dim Arr

    On Error GoTo Error_Load_Repeat_Appt
    Screen.MousePointer = vbHourglass
    grdRepeatHist.MaxRows = 0
    If cnSched.State = adStateClosed Then cnSched.open psConnect(0)
    udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    strSQL = "SELECT * FROM PATIENT_REPEAT_APPT WHERE PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND ACTIVE = 1"
        Set rsSched = cnSched.Execute(strSQL)
        With rsSched
            Do Until .EOF
                Clear_RepeatAppt
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "PROVIDER_ID"
                                    udtRepeatAppt.ProviderID = Trim(Fld.Value)
                                Case "RESOURCE_ID"
                                    udtRepeatAppt.ResourseID = Trim(Fld.Value)
                                Case "FACILITY_ID"
                                    udtRepeatAppt.FacilityID = Trim(Fld.Value)
                                Case "TERMINATION_DATE"
                                    udtRepeatAppt.EndDate = Trim(Fld.Value)
                                Case "START_DATE"
                                    udtRepeatAppt.StartDate = Trim(Fld.Value)
                                Case "REPEAT_DAY_WEEK"
                                    udtRepeatAppt.WeekDay = Trim(Fld.Value)
                                Case "REPEAT_ID"
                                    udtRepeatAppt.RepeatID = Trim(Fld.Value)
                                Case "ACTIVE"
                                    udtRepeatAppt.Status = Trim(Fld.Value)
                                    If udtRepeatAppt.Status = "True" Then
                                        udtRepeatAppt.Status = "Active"
                                    Else
                                        udtRepeatAppt.Status = "InActive"
                                    End If
                            End Select
                        End If
                    End If
                Next
                .MoveNext
                lngRow = grdRepeatHist.MaxRows + 1
                grdRepeatHist.MaxRows = lngRow
                grdRepeatHist.SetRowItemData lngRow, udtRepeatAppt.RepeatID
                lngCol = 0
                 Arr = Split(LookUpCellTag_RepeatHist("FACILITY_ID"), "|")
                If UBound(Arr) >= 0 Then
                    lngCol = Arr(0)
                    If udtRepeatAppt.FacilityID > 0 Then
                        Get_ReportFacility udtRepeatAppt.FacilityID
                    End If
                    Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", udtFacility.Name, "TypeHAlignCenter", "FACILITY_ID"
                    grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
                lngCol = 0
                 Arr = Split(LookUpCellTag_RepeatHist("PROVIDER_ID"), "|")
                If UBound(Arr) >= 0 Then
                    lngCol = Arr(0)
                    If udtRepeatAppt.ProviderID > 0 Then
                        Get_Providers udtRepeatAppt.ProviderID
                    End If
                    Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", clsProvider.Name, "TypeHAlignCenter", "PROVIDER_ID"
                    grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
                lngCol = 0
                 Arr = Split(LookUpCellTag_RepeatHist("START_DATE"), "|")
                If UBound(Arr) >= 0 Then
                    lngCol = Arr(0)
                    If lngCol > 0 Then
                        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", udtRepeatAppt.StartDate, "TypeHAlignCenter", "START_DATE"
                        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                End If
                lngCol = 0
                 Arr = Split(LookUpCellTag_RepeatHist("TERMINATION_DATE"), "|")
                If UBound(Arr) >= 0 Then
                    lngCol = Arr(0)
                    If lngCol > 0 Then
                        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", udtRepeatAppt.EndDate, "TypeHAlignCenter", "TERMINATION_DATE"
                        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                End If
                lngCol = 0
                 Arr = Split(LookUpCellTag_RepeatHist("REPEAT_DAY_WEEK"), "|")
                If UBound(Arr) >= 0 Then
                    lngCol = Arr(0)
                    If lngCol > 0 Then
                        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", udtRepeatAppt.WeekDay, "TypeHAlignCenter", "REPEAT_DAY_WEEK"
                        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                End If
                lngCol = 0
                 Arr = Split(LookUpCellTag_RepeatHist("ACTIVE"), "|")
                If UBound(Arr) >= 0 Then
                    lngCol = Arr(0)
                    If lngCol > 0 Then
                        Format_CellTypeEdit grdRepeatHist, lngRow, lngCol, "1", udtRepeatAppt.Status, "TypeHAlignCenter", "ACTIVE"
                        grdRepeatHist.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                End If
            Loop
        End With
    If cnSched.State = adStateOpen Then
        cnSched.Close
        Set cnSched = Nothing
        Set rsSched = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Repeat_Appt:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Repeat_Appt", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Format_MultiAppt()
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr

    On Error GoTo Error_Format_MultiAppt
    Screen.MousePointer = vbHourglass
    With grdMultiAppt
        .Redraw = False
        .ClearRange 1, SpreadHeader, .MaxCols, .MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .MaxCols = 10
        .FontName = "Times New Roman"
        .GridShowHoriz = True
        .GridShowVert = True
        .RowHeadersShow = False
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .ScrollBars = ScrollBarsBoth
        grdMultiAppt.ColWidth(1) = 20
        grdMultiAppt.ColWidth(2) = 9
        grdMultiAppt.ColWidth(3) = 9
        grdMultiAppt.ColWidth(4) = 9
        grdMultiAppt.ColWidth(5) = 12
        grdMultiAppt.ColWidth(6) = 9
        grdMultiAppt.ColWidth(7) = 9
        grdMultiAppt.ColWidth(8) = 18
        grdMultiAppt.ColWidth(9) = 15
        .Row = SpreadHeader
        .Col = 10
        .ColHidden = True
        lngRow = SpreadHeader
        lngCol = 1
        Format_CellTypeEdit grdMultiAppt, SpreadHeader, 1, "9", "REPEAT APPOINTMENT SCHEDULE FOR" & " " & UCase(udtSchedPatient.FirstName) & " " & UCase(udtSchedPatient.LastName), "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdMultiAppt.SetCellBorder 1, SpreadHeader, 9, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .RowHeight(SpreadHeader) = 20
        lngRow = SpreadHeader + 1
        lngCol = 1
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "FACILITY", "TypeHAlignCenter", "FACILITY_ID", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "APPT DATE", "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "APPT TIME", "TypeHAlignCenter", "APPT_TIME", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "APPT TYPE", "TypeHAlignCenter", "APPT_TYPE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "DURATION", "TypeHAlignCenter", "DURATION", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "START DATE", "TypeHAlignCenter", "START_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "END DATE", "TypeHAlignCenter", "TERMINATION_DATE", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 8
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "COMMENTS", "TypeHAlignCenter", "NOTES", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 9
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "WEEK DAY", "TypeHAlignCenter", "REPEAT_DAY_WEEK", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 10
        Format_CellTypeEdit grdMultiAppt, SpreadHeader + 1, lngCol, "1", "APPT ID", "TypeHAlignCenter", "APPT_ID", HeaderColor
        grdMultiAppt.SetCellBorder lngCol, SpreadHeader + 1, lngCol, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 2, grdMultiAppt
        .Redraw = False
        .MaxRows = 0
    End With
    Screen.MousePointer = vbDefault
'    Format_Repeat_History
'    Load_Repeat_Appt
    Exit Function
Error_Format_MultiAppt:
    Screen.MousePointer = vbDefault
End Function
Private Function Remove_Span()
Dim Col As Variant
Dim Row As Variant
Dim Col2 As Variant
Dim Row2 As Variant
Dim Span As Integer
Dim anchorcol As Variant
Dim anchorrow As Variant
Dim colspan As Variant
Dim rowspan As Variant
Dim i As Long
Dim j As Long
Dim IsSpan As Boolean

    grdMultiAppt.getSelection 0, Col, Row, Col2, Row2
    grdMultiAppt.Redraw = False
    For i = 1 To 7
        For j = SpreadHeader To SpreadHeader + 2
            Span = grdMultiAppt.GetCellSpan(i, j, anchorcol, anchorrow, colspan, rowspan)
            grdMultiAppt.Redraw = False
            If Span Then
                If grdMultiAppt.Col = 1 Then anchorcol = "1"
                If j = SpreadHeader Then anchorrow = SpreadHeader
                If j = SpreadHeader Then anchorrow = SpreadHeader
                grdMultiAppt.RemoveCellSpan anchorcol, anchorrow
                IsSpan = True
            End If
        Next j
    Next i
    
    grdMultiAppt.Redraw = True
    Exit Function
End Function
Private Sub fraNumber_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    flngDeltaX = X
    flngDeltaY = y
End Sub

Private Sub fraNumber_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        fraNumber.Left = fraNumber.Left + X - flngDeltaX
        fraNumber.Top = fraNumber.Top + y - flngDeltaY
    End If
End Sub


Private Sub grdMultiAppt_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim lngCol As Long
Dim lngx As Long
Dim lngY As Long
Dim lngWidth As Long
Dim lngHeight As Long
Dim varCellValue
Dim Idx As Integer
Dim Idx1 As Integer
Dim Arr

    
    If Row <= 0 Then Exit Sub
    mlngGridRow = Row
    mlngGridCol = Col
    With grdMultiAppt
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
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then
            .GetText Col, Row, varCellValue
            Select Case .CellTag
                Case "APPT_DATE"
                    .GetCellPos Col, Row, lngx, lngY, lngWidth, lngHeight
                    fraNumber.ZOrder 0
                    lngY = lngY + 51.5
                    lngx = lngx + 50
                    fraNumber.Top = .Top + lngY
                    fraNumber.Left = .Left + lngx
                    cboRepTime.Visible = False
                    ccdDate.Visible = True
                    ccdDate.Format = dtpShortDate
                    If Trim(varCellValue) <> vbNullString Then
                        ccdDate.Value = Format(varCellValue, "mm/dd/yyyy")
                    Else
                        ccdDate.Value = Format(Now, "mm/dd/yyyy")
                    End If
                    fraNumber.Visible = True
                Case "PROVIDER_ID"
                Case "APPT_TIME"
                    If cboRepTime.listcount <= 0 Then
                         If UCase(frmPatientAppointment.mInterval) = "EVERY 10 MIN" Then
                            SetcboTime cboRepTime, "07", "AM", 10, 50
                            SetcboTime cboRepTime, "08", "AM", 10, 50
                            SetcboTime cboRepTime, "09", "AM", 10, 50
                            SetcboTime cboRepTime, "10", "AM", 10, 50
                            SetcboTime cboRepTime, "11", "AM", 10, 50
                            SetcboTime cboRepTime, "12", "PM", 10, 50
                            SetcboTime cboRepTime, "01", "PM", 10, 50
                            SetcboTime cboRepTime, "02", "PM", 10, 50
                            SetcboTime cboRepTime, "03", "PM", 10, 50
                            SetcboTime cboRepTime, "04", "PM", 10, 50
                            SetcboTime cboRepTime, "05", "PM", 10, 50
                            SetcboTime cboRepTime, "06", "PM", 10, 50
                            SetcboTime cboRepTime, "07", "PM", 10, 50
                            SetcboTime cboRepTime, "08", "PM", 10, 50
                        ElseIf UCase(frmPatientAppointment.mInterval) = "EVERY 15 MIN" Then
                               cboRepTime.AddItem "08:00 AM": cboRepTime.AddItem "08:15 AM": cboRepTime.AddItem "08:30 AM": cboRepTime.AddItem "08:45 AM"
                               cboRepTime.AddItem "09:00 AM": cboRepTime.AddItem "09:15 AM": cboRepTime.AddItem "09:30 AM": cboRepTime.AddItem "09:45 AM"
                               cboRepTime.AddItem "10:00 AM": cboRepTime.AddItem "10:15 AM": cboRepTime.AddItem "10:30 AM": cboRepTime.AddItem "10:45 AM"
                               cboRepTime.AddItem "11:00 AM": cboRepTime.AddItem "11:15 AM": cboRepTime.AddItem "11:30 AM": cboRepTime.AddItem "11:45 AM"
                               cboRepTime.AddItem "12:00 PM": cboRepTime.AddItem "12:15 PM": cboRepTime.AddItem "12:30 PM": cboRepTime.AddItem "12:45 PM"
                               cboRepTime.AddItem "01:00 PM": cboRepTime.AddItem "01:15 PM": cboRepTime.AddItem "01:30 PM": cboRepTime.AddItem "01:45 PM"
                               cboRepTime.AddItem "02:00 PM": cboRepTime.AddItem "02:15 PM": cboRepTime.AddItem "02:30 PM": cboRepTime.AddItem "02:45 PM"
                               cboRepTime.AddItem "03:00 PM": cboRepTime.AddItem "03:15 PM": cboRepTime.AddItem "03:30 PM": cboRepTime.AddItem "03:45 PM"
                               cboRepTime.AddItem "04:00 PM": cboRepTime.AddItem "04:15 PM": cboRepTime.AddItem "04:30 PM": cboRepTime.AddItem "04:45 PM"
                               cboRepTime.AddItem "05:00 PM": cboRepTime.AddItem "05:15 PM": cboRepTime.AddItem "05:30 PM": cboRepTime.AddItem "05:45 PM"
                               cboRepTime.AddItem "06:00 PM": cboRepTime.AddItem "06:15 PM": cboRepTime.AddItem "06:30 PM": cboRepTime.AddItem "06:45 PM"
                               cboRepTime.AddItem "07:00 PM": cboRepTime.AddItem "07:15 PM": cboRepTime.AddItem "07:30 PM": cboRepTime.AddItem "07:45 PM"
                               cboRepTime.AddItem "08:00 PM": cboRepTime.AddItem "08:15 PM": cboRepTime.AddItem "08:30 PM": cboRepTime.AddItem "08:45 PM"
                        ElseIf UCase(frmPatientAppointment.mInterval) = "EVERY 20 MIN" Then
                             cboRepTime.AddItem "08:00 AM": cboRepTime.AddItem "08:20 AM": cboRepTime.AddItem "08:40 AM"
                             cboRepTime.AddItem "09:00 AM": cboRepTime.AddItem "09:20 AM": cboRepTime.AddItem "09:40 AM"
                             cboRepTime.AddItem "10:00 AM": cboRepTime.AddItem "10:20 AM": cboRepTime.AddItem "10:40 AM"
                             cboRepTime.AddItem "11:00 AM": cboRepTime.AddItem "11:20 AM": cboRepTime.AddItem "11:40 AM"
                             cboRepTime.AddItem "12:00 PM": cboRepTime.AddItem "12:20 PM": cboRepTime.AddItem "12:40 PM"
                             cboRepTime.AddItem "01:00 PM": cboRepTime.AddItem "01:20 PM": cboRepTime.AddItem "01:40 PM"
                             cboRepTime.AddItem "02:00 PM": cboRepTime.AddItem "02:20 PM": cboRepTime.AddItem "02:40 PM"
                             cboRepTime.AddItem "03:00 PM": cboRepTime.AddItem "03:20 PM": cboRepTime.AddItem "03:40 PM"
                             cboRepTime.AddItem "04:00 PM": cboRepTime.AddItem "04:20 PM": cboRepTime.AddItem "04:40 PM"
                             cboRepTime.AddItem "05:00 PM": cboRepTime.AddItem "05:20 PM": cboRepTime.AddItem "05:40 PM"
                             cboRepTime.AddItem "06:00 PM": cboRepTime.AddItem "06:20 PM": cboRepTime.AddItem "06:40 PM"
                             cboRepTime.AddItem "07:00 PM": cboRepTime.AddItem "07:20 PM": cboRepTime.AddItem "07:40 PM"
                             cboRepTime.AddItem "08:00 PM": cboRepTime.AddItem "08:20 PM": cboRepTime.AddItem "08:40 PM"
                          End If
                    End If
                    .GetCellPos Col, Row, lngx, lngY, lngWidth, lngHeight
                    fraNumber.ZOrder 0
                    lngY = lngY + 51.5
                    lngx = lngx + 50
                    fraNumber.Top = .Top + lngY
                    fraNumber.Left = .Left + lngx
                    ccdDate.Visible = False
                    If Trim(varCellValue) <> vbNullString Then
                        For Idx = 0 To cboRepTime.listcount - 1
                            If cboRepTime.List(Idx) = varCellValue Then
                                cboRepTime.Text = cboRepTime.List(Idx)
                                Exit For
                            End If
                        Next Idx
                    End If
                    cboRepTime.Visible = True
                    fraNumber.Visible = True
            End Select
        End If
        lngCol = 0
        Arr = Split(LookUpCellTag("APPT_DATE"), "|")
        If UBound(Arr) >= 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                .GetText lngCol, lngRow, varCellValue
                If IsDate(varCellValue) Then
                    udtRepeatAppt.ApptDate = Format(varCellValue, "mm/dd/yyyy")
                End If
            End If
        End If
        lngCol = 0
         Arr = Split(LookUpCellTag("APPT_TIME"), "|")
        If UBound(Arr) >= 0 Then
            lngCol = Arr(0)
            If lngCol > 0 Then
                .GetText lngCol, lngRow, varCellValue
                If IsDate(varCellValue) Then
                    udtRepeatAppt.StartTime = Format(varCellValue, "mm/dd/yyyy")
                End If
            End If
        End If
    '-----------------------------------------
    End With
    Exit Sub
End Sub
Private Sub grdRepeatHist_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long

    With grdRepeatHist
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
        udtRepeatAppt.RepeatID = .GetRowItemData(Row)
        Format_MultiAppt
        Read_MultiAppt
    End With
End Sub


Public Sub SetcboTime(ipCBO As Object, ipHour As String, ipAMPM As String, ipInt As Long, ipEnd As Long)
    Dim Idx As Long
    
    For Idx = 0 To ipEnd Step ipInt
        ipCBO.AddItem ipHour & ":" & Format(Idx, "00") & " " & ipAMPM
    Next Idx
End Sub
