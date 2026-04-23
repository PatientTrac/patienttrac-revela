VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{6514F5A0-641C-11D2-9FD0-0020AF131A57}#3.0#0"; "fpFlp30.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmBlockTime 
   Caption         =   "Set Schedule Blocks"
   ClientHeight    =   7770
   ClientLeft      =   5070
   ClientTop       =   3345
   ClientWidth     =   14715
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBlockTime.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   14715
   Begin LpADOLib.fpComboAdo fpTerm 
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   -600
      Width           =   2925
      _Version        =   196608
      _ExtentX        =   5159
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
      ForeColor       =   8388608
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
      ThreeDInsideShadowColor=   8388608
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   8388608
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   8388608
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8388608
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
      ColDesigner     =   "frmBlockTime.frx":08CA
   End
   Begin VB.Frame fraRemove 
      BackColor       =   &H00E0E0E0&
      Height          =   1785
      Left            =   1890
      TabIndex        =   40
      Top             =   1470
      Visible         =   0   'False
      Width           =   3795
      Begin VB.TextBox cboStaff 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   900
         Width           =   3435
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
         Height          =   330
         ItemData        =   "frmBlockTime.frx":0BF9
         Left            =   930
         List            =   "frmBlockTime.frx":0C03
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   300
         Width           =   2625
      End
      Begin PT_XP_Button.PT_XP cmdRemReason 
         Height          =   315
         Left            =   2040
         TabIndex        =   44
         Top             =   1410
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "frmBlockTime.frx":0C29
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBlockTime.frx":163B
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSelectRemove 
         Height          =   315
         Left            =   870
         TabIndex        =   45
         Top             =   1410
         Width           =   1035
         _ExtentX        =   1826
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
         Picture         =   "frmBlockTime.frx":198D
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBlockTime.frx":1D27
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblReason 
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   120
         TabIndex        =   43
         Top             =   375
         Width           =   855
      End
      Begin VB.Label lblStaff 
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   120
         TabIndex        =   41
         Top             =   660
         Width           =   1185
      End
   End
   Begin VB.Frame fraSubmitDate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Date Range Selection"
      ForeColor       =   &H00800000&
      Height          =   1395
      Left            =   990
      TabIndex        =   32
      Top             =   30
      Visible         =   0   'False
      Width           =   3825
      Begin MSComCtl2.DTPicker ccdSubTo 
         Height          =   285
         Left            =   2040
         TabIndex        =   33
         Top             =   540
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
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
         CustomFormat    =   "MM/dd/yyy 11:59:00 PM"
         Format          =   77660163
         CurrentDate     =   40179.0006944444
      End
      Begin MSComCtl2.DTPicker ccdSubFrom 
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
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
         CustomFormat    =   "MM/dd/yyy "
         Format          =   77660163
         CurrentDate     =   40179.0006944444
      End
      Begin PT_XP_Button.PT_XP cmdSelect 
         Height          =   315
         Left            =   210
         TabIndex        =   35
         Top             =   1020
         Width           =   1065
         _ExtentX        =   1879
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
         Picture         =   "frmBlockTime.frx":2079
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Select"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBlockTime.frx":2413
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCloseSubmit 
         Height          =   315
         Left            =   1380
         TabIndex        =   36
         Top             =   1020
         Width           =   1035
         _ExtentX        =   1826
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
         Picture         =   "frmBlockTime.frx":2765
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBlockTime.frx":3177
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdShowAll_Submit 
         Height          =   315
         Left            =   2490
         TabIndex        =   37
         Top             =   1020
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "frmBlockTime.frx":34C9
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Show All"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBlockTime.frx":3A63
         ShowFocus       =   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "From"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "To"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2040
         TabIndex        =   39
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame fraNumber 
      BackColor       =   &H00FFC0C0&
      Height          =   1515
      Left            =   4320
      TabIndex        =   7
      Top             =   1530
      Visible         =   0   'False
      Width           =   4515
      Begin VB.ComboBox cboRepTimeTo 
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
         Left            =   2370
         TabIndex        =   29
         Top             =   630
         Width           =   1875
      End
      Begin VB.ComboBox cboRepTimeFrom 
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
         Left            =   180
         TabIndex        =   13
         Top             =   630
         Width           =   1875
      End
      Begin PT_XP_Button.PT_XP cmdAccept 
         Height          =   285
         Left            =   1890
         TabIndex        =   8
         Top             =   1050
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
         Picture         =   "frmBlockTime.frx":3DB5
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBlockTime.frx":47C7
         ShowFocus       =   -1  'True
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   210
         TabIndex        =   28
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2370
         TabIndex        =   27
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   225
         Left            =   60
         TabIndex        =   17
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Date Range Selection"
      ForeColor       =   &H00800000&
      Height          =   2475
      Left            =   8430
      TabIndex        =   18
      Top             =   1410
      Visible         =   0   'False
      Width           =   4125
      Begin VB.TextBox cctDuration 
         Height          =   330
         Left            =   1680
         TabIndex        =   47
         Top             =   1500
         Width           =   735
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
         Height          =   330
         ItemData        =   "frmBlockTime.frx":4B19
         Left            =   2460
         List            =   "frmBlockTime.frx":4B2F
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1500
         Width           =   1365
      End
      Begin VB.ComboBox cboWeek 
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
         ItemData        =   "frmBlockTime.frx":4B6B
         Left            =   1620
         List            =   "frmBlockTime.frx":4B84
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1020
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker ccdEndDate 
         Height          =   285
         Left            =   2340
         TabIndex        =   19
         Top             =   510
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
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
         Format          =   77660163
         CurrentDate     =   38795.9993055556
      End
      Begin MSComCtl2.DTPicker ccdStartDate 
         Height          =   315
         Left            =   105
         TabIndex        =   20
         Top             =   510
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
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
         CustomFormat    =   "MM/dd/yyy "
         DateIsNull      =   -1  'True
         Format          =   77660163
         CurrentDate     =   38795.0006944444
      End
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   315
         Left            =   1020
         TabIndex        =   21
         Top             =   2010
         Width           =   1035
         _ExtentX        =   1826
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
         Picture         =   "frmBlockTime.frx":4BC2
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Select"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBlockTime.frx":4F5C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCancel 
         Height          =   315
         Left            =   2190
         TabIndex        =   22
         Top             =   2010
         Width           =   1035
         _ExtentX        =   1826
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
         Picture         =   "frmBlockTime.frx":52AE
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBlockTime.frx":5CC0
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblDuration 
         BackStyle       =   0  'Transparent
         Caption         =   "Duration"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   120
         TabIndex        =   30
         Top             =   1575
         Width           =   1605
      End
      Begin VB.Label lblDayWeek 
         BackStyle       =   0  'Transparent
         Caption         =   "Day of Week"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   1110
         Width           =   1605
      End
      Begin VB.Label lblFrom 
         BackColor       =   &H00E0E0E0&
         Caption         =   "From"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label lblTo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "To"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2340
         TabIndex        =   23
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame fraSchedConflicts 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   6675
      Left            =   1140
      TabIndex        =   9
      Top             =   450
      Visible         =   0   'False
      Width           =   12150
      Begin FPSpreadADO.fpSpread grdTerms 
         Height          =   6195
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   12075
         _Version        =   524288
         _ExtentX        =   21299
         _ExtentY        =   10927
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         MaxCols         =   5
         MaxRows         =   0
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   0
         ShadowText      =   0
         SpreadDesigner  =   "frmBlockTime.frx":6012
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin PT_XP_Button.PT_XP cmdExit 
         Height          =   345
         Left            =   5550
         TabIndex        =   11
         ToolTipText     =   "Save, Restore, Close"
         Top             =   6270
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmBlockTime.frx":6512
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   14215660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBlockTime.frx":6F24
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrintList 
         Height          =   375
         Left            =   10530
         TabIndex        =   12
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
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
         Picture         =   "frmBlockTime.frx":7276
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Print"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBlockTime.frx":7810
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.ComboBox cboProviders 
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
      Left            =   5250
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   60
      Width           =   3555
   End
   Begin PT_XP_Button.PT_XP cmdDurationBlock 
      Height          =   345
      Left            =   14190
      TabIndex        =   0
      ToolTipText     =   "Duration Terms"
      Top             =   7320
      Width           =   315
      _ExtentX        =   556
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
      Picture         =   "frmBlockTime.frx":7B62
      BackColorDown   =   15133676
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmBlockTime.frx":80FC
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cboReasonTerms 
      Height          =   345
      Left            =   13770
      TabIndex        =   1
      ToolTipText     =   "Reason Terms"
      Top             =   7320
      Width           =   315
      _ExtentX        =   556
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
      Picture         =   "frmBlockTime.frx":844E
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmBlockTime.frx":89E8
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   495
      Left            =   7365
      TabIndex        =   2
      Top             =   7230
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
      Picture         =   "frmBlockTime.frx":8D3A
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmBlockTime.frx":974C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSaveBlock 
      Height          =   495
      Left            =   5715
      TabIndex        =   3
      Top             =   7230
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
      Picture         =   "frmBlockTime.frx":9A9E
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Set Block Time"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmBlockTime.frx":A038
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdBlock 
      Height          =   6705
      Left            =   30
      TabIndex        =   4
      Top             =   420
      Width           =   14625
      _Version        =   524288
      _ExtentX        =   25797
      _ExtentY        =   11827
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
      MaxCols         =   7
      MaxRows         =   0
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmBlockTime.frx":A38A
      UserResize      =   0
      Appearance      =   2
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
   Begin PT_XP_Button.PT_XP cmdRemoveRow 
      Height          =   375
      Left            =   1350
      TabIndex        =   5
      Top             =   30
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmBlockTime.frx":A88B
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Remove"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmBlockTime.frx":AE25
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdGridRow 
      Height          =   375
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmBlockTime.frx":B177
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmBlockTime.frx":B711
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblMessages 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Provider"
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
      Height          =   225
      Left            =   4050
      TabIndex        =   16
      Top             =   120
      Width           =   1185
   End
   Begin VB.Menu mnu_Show 
      Caption         =   "Show All"
   End
   Begin VB.Menu mnu_Date_Range 
      Caption         =   "Select Date Range"
   End
   Begin VB.Menu mnu_PrintBlock 
      Caption         =   "Print Block Schedule"
   End
   Begin VB.Menu mnu_Conflict 
      Caption         =   "Conflcting Appointments"
   End
End
Attribute VB_Name = "frmBlockTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbooData As Boolean
Private mbooByPass As Boolean
Private mlngGridRow As Long
Private mlngGridCol As Long
Private mStartHour() As String
Private mEndHour() As String
Private mInterval As String
Private mStartIdx As Integer
Private mEndIdx As Integer
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
Private udtBlockTime As udtBlockTime
Private Type udtBlockTime
    EndDate As String
    StartDate As String
    StartTime As String
    EndTime As String
    Range As String
End Type

Private Function Clear_BlockTerm()
   With udtBlock
        .StartDate = vbNullString
        .EndDate = vbNullString
        .StartTime = vbNullString
        .EndTime = vbNullString
        .Reason = vbNullString
        .Provider = vbNullString
        .ProviderID = 0
        .FacilityID = 0
        .Duration = vbNullString
        .DayWeek = vbNullString
    End With
End Function

Private Function Clear_BlockTime()
    With udtBlockTime
        .EndDate = vbNullString
        .StartDate = vbNullString
        .StartTime = vbNullString
        .EndTime = vbNullString
        .Range = vbNullString
    End With
End Function

Private Function DayWeek()

End Function

Private Function Get_LogonProvider()
Dim lngFacilityID As Long
Dim strSQL As String
Dim cnProv As New ADODB.Connection
Dim rsProv As New ADODB.Recordset
Dim lngProv As Long
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim strSuffix As String

    On Error GoTo Error_Get_Providers
    Screen.MousePointer = vbHourglass
    If cnProv.State = adStateClosed Then cnProv.open psConnect(0)
    strSQL = "Get_Providers " & mlngLogonUser
    Set rsProv = cnProv.Execute(strSQL)
    With rsProv
        If Not .EOF Then
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
                        End Select
                    End If
                End If
            Next
        End If
    End With
    clsProvider.Address = Fix_Name(clsProvider.Address)
    If Trim(strFirst) <> vbNullString Then
        strFirst = Trim(strFirst)
        cboStaff.Text = Fix_Name(strFirst)
    End If
    If Trim(strMiddle) <> vbNullString Then
       strMiddle = IsProperCase(strMiddle)
        cboStaff.Text = cboStaff.Text & " " & Left(strMiddle, 1)
    End If
    If Trim(strLast) <> vbNullString Then
        strLast = Trim(strLast)
        strLast = Fix_Name(strLast)
        cboStaff.Text = cboStaff.Text & " " & strLast
    End If
    If Trim(strSuffix) <> vbNullString Then
        strSuffix = Fix_Suffix(strSuffix)
        cboStaff.Text = cboStaff.Text & ", " & Trim(strSuffix)
    End If
    If Trim(strTitle) <> vbNullString Then
        strTitle = Fix_Title(strTitle)
        cboStaff.Text = cboStaff.Text & ", " & Trim(strTitle)
    End If
    If cnProv.State = adStateOpen Then
       cnProv.Close
       Set cnProv = Nothing
       If rsProv.State = adStateOpen Then rsProv.Close
       Set rsProv = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Providers:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Providers", App.EXEName & "." & "ModPatientData"
    Exit Function
    Resume Next
End Function

Private Function Load_ComboTime()
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim strTerms As String

    On Error GoTo Error_Load_ComboTime
    Screen.MousePointer = vbHourglass
    cboRepTimeFrom.Clear
    cboRepTimeTo.Clear
    cboDuration.Clear
'    cboDuration.AddItem "Date Range"
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Get_Attr_Valid_Value_SubDom_Alphabetic 'sd_Appointment_BlockDuration'"
    Set rsTerms = cnTerms.Execute(strSQL)
    Do Until rsTerms.EOF
        strTerms = vbNullString
        strTerms = IsProperCase(rsTerms.Fields("ATTRBTE_VALID_VALUE"))
        rsTerms.MoveNext
        If strTerms <> vbNullString Then
            cboDuration.AddItem strTerms
        End If
    Loop
    Select Case UCase(frmPatientAppointment.mInterval)
        Case "EVERY 15 MIN"
           cboRepTimeFrom.AddItem "07:00 AM": cboRepTimeFrom.AddItem "07:15 AM": cboRepTimeFrom.AddItem "07:30 AM": cboRepTimeFrom.AddItem "07:45 AM"
           cboRepTimeFrom.AddItem "08:00 AM": cboRepTimeFrom.AddItem "08:15 AM": cboRepTimeFrom.AddItem "08:30 AM": cboRepTimeFrom.AddItem "08:45 AM"
           cboRepTimeFrom.AddItem "09:00 AM": cboRepTimeFrom.AddItem "09:15 AM": cboRepTimeFrom.AddItem "09:30 AM": cboRepTimeFrom.AddItem "09:45 AM"
           cboRepTimeFrom.AddItem "10:00 AM": cboRepTimeFrom.AddItem "10:15 AM": cboRepTimeFrom.AddItem "10:30 AM": cboRepTimeFrom.AddItem "10:45 AM"
           cboRepTimeFrom.AddItem "11:00 AM": cboRepTimeFrom.AddItem "11:15 AM": cboRepTimeFrom.AddItem "11:30 AM": cboRepTimeFrom.AddItem "11:45 AM"
           cboRepTimeFrom.AddItem "12:00 PM": cboRepTimeFrom.AddItem "12:15 PM": cboRepTimeFrom.AddItem "12:30 PM": cboRepTimeFrom.AddItem "12:45 PM"
           cboRepTimeFrom.AddItem "01:00 PM": cboRepTimeFrom.AddItem "01:15 PM": cboRepTimeFrom.AddItem "01:30 PM": cboRepTimeFrom.AddItem "01:45 PM"
           cboRepTimeFrom.AddItem "02:00 PM": cboRepTimeFrom.AddItem "02:15 PM": cboRepTimeFrom.AddItem "02:30 PM": cboRepTimeFrom.AddItem "02:45 PM"
           cboRepTimeFrom.AddItem "03:00 PM": cboRepTimeFrom.AddItem "03:15 PM": cboRepTimeFrom.AddItem "03:30 PM": cboRepTimeFrom.AddItem "03:45 PM"
           cboRepTimeFrom.AddItem "04:00 PM": cboRepTimeFrom.AddItem "04:15 PM": cboRepTimeFrom.AddItem "04:30 PM": cboRepTimeFrom.AddItem "04:45 PM"
           cboRepTimeFrom.AddItem "05:00 PM": cboRepTimeFrom.AddItem "05:15 PM": cboRepTimeFrom.AddItem "05:30 PM": cboRepTimeFrom.AddItem "05:45 PM"
           cboRepTimeFrom.AddItem "06:00 PM": cboRepTimeFrom.AddItem "06:15 PM": cboRepTimeFrom.AddItem "06:30 PM": cboRepTimeFrom.AddItem "06:45 PM"
           cboRepTimeFrom.AddItem "07:00 PM": cboRepTimeFrom.AddItem "07:15 PM": cboRepTimeFrom.AddItem "07:30 PM": cboRepTimeFrom.AddItem "07:45 PM"
           cboRepTimeFrom.AddItem "08:00 PM": cboRepTimeFrom.AddItem "08:15 PM": cboRepTimeFrom.AddItem "08:30 PM": cboRepTimeFrom.AddItem "08:45 PM"
           cboRepTimeTo.AddItem "07:00 AM": cboRepTimeTo.AddItem "07:15 AM": cboRepTimeTo.AddItem "07:30 AM": cboRepTimeTo.AddItem "07:45 AM"
           cboRepTimeTo.AddItem "08:00 AM": cboRepTimeTo.AddItem "08:15 AM": cboRepTimeTo.AddItem "08:30 AM": cboRepTimeTo.AddItem "08:45 AM"
           cboRepTimeTo.AddItem "09:00 AM": cboRepTimeTo.AddItem "09:15 AM": cboRepTimeTo.AddItem "09:30 AM": cboRepTimeTo.AddItem "09:45 AM"
           cboRepTimeTo.AddItem "10:00 AM": cboRepTimeTo.AddItem "10:15 AM": cboRepTimeTo.AddItem "10:30 AM": cboRepTimeTo.AddItem "10:45 AM"
           cboRepTimeTo.AddItem "11:00 AM": cboRepTimeTo.AddItem "11:15 AM": cboRepTimeTo.AddItem "11:30 AM": cboRepTimeTo.AddItem "11:45 AM"
           cboRepTimeTo.AddItem "12:00 PM": cboRepTimeTo.AddItem "12:15 PM": cboRepTimeTo.AddItem "12:30 PM": cboRepTimeTo.AddItem "12:45 PM"
           cboRepTimeTo.AddItem "01:00 PM": cboRepTimeTo.AddItem "01:15 PM": cboRepTimeTo.AddItem "01:30 PM": cboRepTimeTo.AddItem "01:45 PM"
           cboRepTimeTo.AddItem "02:00 PM": cboRepTimeTo.AddItem "02:15 PM": cboRepTimeTo.AddItem "02:30 PM": cboRepTimeTo.AddItem "02:45 PM"
           cboRepTimeTo.AddItem "03:00 PM": cboRepTimeTo.AddItem "03:15 PM": cboRepTimeTo.AddItem "03:30 PM": cboRepTimeTo.AddItem "03:45 PM"
           cboRepTimeTo.AddItem "04:00 PM": cboRepTimeTo.AddItem "04:15 PM": cboRepTimeTo.AddItem "04:30 PM": cboRepTimeTo.AddItem "04:45 PM"
           cboRepTimeTo.AddItem "05:00 PM": cboRepTimeTo.AddItem "05:15 PM": cboRepTimeTo.AddItem "05:30 PM": cboRepTimeTo.AddItem "05:45 PM"
           cboRepTimeTo.AddItem "06:00 PM": cboRepTimeTo.AddItem "06:15 PM": cboRepTimeTo.AddItem "06:30 PM": cboRepTimeTo.AddItem "06:45 PM"
           cboRepTimeTo.AddItem "07:00 PM": cboRepTimeTo.AddItem "07:15 PM": cboRepTimeTo.AddItem "07:30 PM": cboRepTimeTo.AddItem "07:45 PM"
           cboRepTimeTo.AddItem "08:00 PM": cboRepTimeTo.AddItem "08:15 PM": cboRepTimeTo.AddItem "08:30 PM": cboRepTimeTo.AddItem "08:45 PM"
       Case "EVERY 5 MIN"
           cboRepTimeFrom.AddItem "07:00 AM": cboRepTimeFrom.AddItem "07:05 AM": cboRepTimeFrom.AddItem "07:10 AM": cboRepTimeFrom.AddItem "07:15 AM"
           cboRepTimeFrom.AddItem "07:20 AM": cboRepTimeFrom.AddItem "07:25 AM": cboRepTimeFrom.AddItem "07:30 AM": cboRepTimeFrom.AddItem "07:35 AM"
           cboRepTimeFrom.AddItem "07:40 AM": cboRepTimeFrom.AddItem "07:45 AM": cboRepTimeFrom.AddItem "07:50 AM": cboRepTimeFrom.AddItem "07:55 AM"
           cboRepTimeFrom.AddItem "08:00 AM": cboRepTimeFrom.AddItem "08:05 AM": cboRepTimeFrom.AddItem "08:10 AM": cboRepTimeFrom.AddItem "08:15 AM"
           cboRepTimeFrom.AddItem "08:20 AM": cboRepTimeFrom.AddItem "08:25 AM": cboRepTimeFrom.AddItem "08:30 AM": cboRepTimeFrom.AddItem "08:35 AM"
           cboRepTimeFrom.AddItem "08:40 AM": cboRepTimeFrom.AddItem "08:45 AM": cboRepTimeFrom.AddItem "08:50 AM": cboRepTimeFrom.AddItem "08:55 AM"
           cboRepTimeFrom.AddItem "09:00 AM": cboRepTimeFrom.AddItem "09:05 AM": cboRepTimeFrom.AddItem "09:10 AM": cboRepTimeFrom.AddItem "09:15 AM"
           cboRepTimeFrom.AddItem "09:20 AM": cboRepTimeFrom.AddItem "09:25 AM": cboRepTimeFrom.AddItem "09:30 AM": cboRepTimeFrom.AddItem "09:35 AM"
           cboRepTimeFrom.AddItem "09:40 AM": cboRepTimeFrom.AddItem "09:45 AM": cboRepTimeFrom.AddItem "09:50 AM": cboRepTimeFrom.AddItem "09:55 AM"
           cboRepTimeFrom.AddItem "10:00 AM": cboRepTimeFrom.AddItem "10:05 AM": cboRepTimeFrom.AddItem "10:10 AM": cboRepTimeFrom.AddItem "10:15 AM"
           cboRepTimeFrom.AddItem "10:20 AM": cboRepTimeFrom.AddItem "10:25 AM": cboRepTimeFrom.AddItem "10:30 AM": cboRepTimeFrom.AddItem "10:35 AM"
           cboRepTimeFrom.AddItem "10:40 AM": cboRepTimeFrom.AddItem "10:45 AM": cboRepTimeFrom.AddItem "10:50 AM": cboRepTimeFrom.AddItem "10:55 AM"
           cboRepTimeFrom.AddItem "11:00 AM": cboRepTimeFrom.AddItem "11:05 AM": cboRepTimeFrom.AddItem "11:10 AM": cboRepTimeFrom.AddItem "11:15 AM"
           cboRepTimeFrom.AddItem "11:20 AM": cboRepTimeFrom.AddItem "11:25 AM": cboRepTimeFrom.AddItem "11:30 AM": cboRepTimeFrom.AddItem "11:35 AM"
           cboRepTimeFrom.AddItem "11:40 AM": cboRepTimeFrom.AddItem "11:45 AM": cboRepTimeFrom.AddItem "11:50 AM": cboRepTimeFrom.AddItem "11:55 AM"
           cboRepTimeFrom.AddItem "12:00 PM": cboRepTimeFrom.AddItem "12:05 PM": cboRepTimeFrom.AddItem "12:10 PM": cboRepTimeFrom.AddItem "12:15 PM"
           cboRepTimeFrom.AddItem "12:20 PM": cboRepTimeFrom.AddItem "12:25 PM": cboRepTimeFrom.AddItem "12:30 PM": cboRepTimeFrom.AddItem "12:35 PM"
           cboRepTimeFrom.AddItem "01:00 PM": cboRepTimeFrom.AddItem "01:05 PM": cboRepTimeFrom.AddItem "01:10 PM": cboRepTimeFrom.AddItem "01:15 PM"
           cboRepTimeFrom.AddItem "01:20 PM": cboRepTimeFrom.AddItem "01:25 PM": cboRepTimeFrom.AddItem "01:30 PM": cboRepTimeFrom.AddItem "01:35 PM"
           cboRepTimeFrom.AddItem "01:40 PM": cboRepTimeFrom.AddItem "01:45 PM": cboRepTimeFrom.AddItem "01:50 PM": cboRepTimeFrom.AddItem "01:55 PM"
           cboRepTimeFrom.AddItem "02:00 PM": cboRepTimeFrom.AddItem "02:05 PM": cboRepTimeFrom.AddItem "02:10 PM": cboRepTimeFrom.AddItem "02:15 PM"
           cboRepTimeFrom.AddItem "02:20 PM": cboRepTimeFrom.AddItem "02:25 PM": cboRepTimeFrom.AddItem "02:30 PM": cboRepTimeFrom.AddItem "02:35 PM"
           cboRepTimeFrom.AddItem "02:40 PM": cboRepTimeFrom.AddItem "02:45 PM": cboRepTimeFrom.AddItem "02:50 PM": cboRepTimeFrom.AddItem "02:55 PM"
           cboRepTimeFrom.AddItem "03:00 PM": cboRepTimeFrom.AddItem "03:05 PM": cboRepTimeFrom.AddItem "03:10 PM": cboRepTimeFrom.AddItem "03:15 PM"
           cboRepTimeFrom.AddItem "03:20 PM": cboRepTimeFrom.AddItem "03:25 PM": cboRepTimeFrom.AddItem "03:30 PM": cboRepTimeFrom.AddItem "03:35 PM"
           cboRepTimeFrom.AddItem "03:40 PM": cboRepTimeFrom.AddItem "03:45 PM": cboRepTimeFrom.AddItem "03:50 PM": cboRepTimeFrom.AddItem "03:55 PM"
           cboRepTimeFrom.AddItem "04:00 PM": cboRepTimeFrom.AddItem "04:05 PM": cboRepTimeFrom.AddItem "04:10 PM": cboRepTimeFrom.AddItem "04:15 PM"
           cboRepTimeFrom.AddItem "04:20 PM": cboRepTimeFrom.AddItem "04:25 PM": cboRepTimeFrom.AddItem "04:30 PM": cboRepTimeFrom.AddItem "04:35 PM"
           cboRepTimeFrom.AddItem "04:40 PM": cboRepTimeFrom.AddItem "04:45 PM": cboRepTimeFrom.AddItem "04:50 PM": cboRepTimeFrom.AddItem "04:55 PM"
           cboRepTimeFrom.AddItem "05:00 PM": cboRepTimeFrom.AddItem "05:05 PM": cboRepTimeFrom.AddItem "05:10 PM": cboRepTimeFrom.AddItem "05:15 PM"
           cboRepTimeFrom.AddItem "05:20 PM": cboRepTimeFrom.AddItem "05:25 PM": cboRepTimeFrom.AddItem "05:30 PM": cboRepTimeFrom.AddItem "05:35 PM"
           cboRepTimeFrom.AddItem "05:40 PM": cboRepTimeFrom.AddItem "05:45 PM": cboRepTimeFrom.AddItem "05:50 PM": cboRepTimeFrom.AddItem "05:55 PM"
           cboRepTimeFrom.AddItem "06:00 PM": cboRepTimeFrom.AddItem "06:05 PM": cboRepTimeFrom.AddItem "06:10 PM": cboRepTimeFrom.AddItem "06:15 PM"
           cboRepTimeFrom.AddItem "06:20 PM": cboRepTimeFrom.AddItem "06:25 PM": cboRepTimeFrom.AddItem "06:30 PM": cboRepTimeFrom.AddItem "06:35 PM"
           cboRepTimeFrom.AddItem "06:40 PM": cboRepTimeFrom.AddItem "06:45 PM": cboRepTimeFrom.AddItem "06:50 PM": cboRepTimeFrom.AddItem "06:55 PM"
           cboRepTimeFrom.AddItem "07:00 PM": cboRepTimeFrom.AddItem "07:05 PM": cboRepTimeFrom.AddItem "07:10 PM": cboRepTimeFrom.AddItem "07:15 PM"
           cboRepTimeFrom.AddItem "07:20 PM": cboRepTimeFrom.AddItem "07:25 PM": cboRepTimeFrom.AddItem "07:30 PM": cboRepTimeFrom.AddItem "07:35 PM"
           cboRepTimeFrom.AddItem "07:40 PM": cboRepTimeFrom.AddItem "07:45 PM": cboRepTimeFrom.AddItem "07:50 PM": cboRepTimeFrom.AddItem "07:55 PM"
           cboRepTimeFrom.AddItem "08:00 PM": cboRepTimeFrom.AddItem "08:05 PM": cboRepTimeFrom.AddItem "08:10 PM": cboRepTimeFrom.AddItem "08:15 PM"
           cboRepTimeFrom.AddItem "08:20 PM": cboRepTimeFrom.AddItem "08:25 PM": cboRepTimeFrom.AddItem "08:30 PM": cboRepTimeFrom.AddItem "08:35 PM"
           cboRepTimeFrom.AddItem "08:40 PM": cboRepTimeFrom.AddItem "08:45 PM": cboRepTimeFrom.AddItem "08:50 PM": cboRepTimeFrom.AddItem "08:55 PM"
           cboRepTimeTo.AddItem "07:00 AM": cboRepTimeTo.AddItem "07:05 AM": cboRepTimeTo.AddItem "07:10 AM": cboRepTimeTo.AddItem "07:15 AM"
           cboRepTimeTo.AddItem "07:20 AM": cboRepTimeTo.AddItem "07:25 AM": cboRepTimeTo.AddItem "07:30 AM": cboRepTimeTo.AddItem "07:35 AM"
           cboRepTimeTo.AddItem "07:40 AM": cboRepTimeTo.AddItem "07:45 AM": cboRepTimeTo.AddItem "07:50 AM": cboRepTimeTo.AddItem "07:55 AM"
           cboRepTimeTo.AddItem "08:00 AM": cboRepTimeTo.AddItem "08:05 AM": cboRepTimeTo.AddItem "08:10 AM": cboRepTimeTo.AddItem "08:15 AM"
           cboRepTimeTo.AddItem "08:20 AM": cboRepTimeTo.AddItem "08:25 AM": cboRepTimeTo.AddItem "08:30 AM": cboRepTimeTo.AddItem "08:35 AM"
           cboRepTimeTo.AddItem "08:40 AM": cboRepTimeTo.AddItem "08:45 AM": cboRepTimeTo.AddItem "08:50 AM": cboRepTimeTo.AddItem "08:55 AM"
           cboRepTimeTo.AddItem "09:00 AM": cboRepTimeTo.AddItem "09:05 AM": cboRepTimeTo.AddItem "09:10 AM": cboRepTimeTo.AddItem "09:15 AM"
           cboRepTimeTo.AddItem "09:20 AM": cboRepTimeTo.AddItem "09:25 AM": cboRepTimeTo.AddItem "09:30 AM": cboRepTimeTo.AddItem "09:35 AM"
           cboRepTimeTo.AddItem "09:40 AM": cboRepTimeTo.AddItem "09:45 AM": cboRepTimeTo.AddItem "09:50 AM": cboRepTimeTo.AddItem "09:55 AM"
           cboRepTimeTo.AddItem "10:00 AM": cboRepTimeTo.AddItem "10:05 AM": cboRepTimeTo.AddItem "10:10 AM": cboRepTimeTo.AddItem "10:15 AM"
           cboRepTimeTo.AddItem "10:20 AM": cboRepTimeTo.AddItem "10:25 AM": cboRepTimeTo.AddItem "10:30 AM": cboRepTimeTo.AddItem "10:35 AM"
           cboRepTimeTo.AddItem "10:40 AM": cboRepTimeTo.AddItem "10:45 AM": cboRepTimeTo.AddItem "10:50 AM": cboRepTimeTo.AddItem "10:55 AM"
           cboRepTimeTo.AddItem "11:00 AM": cboRepTimeTo.AddItem "11:05 AM": cboRepTimeTo.AddItem "11:10 AM": cboRepTimeTo.AddItem "11:15 AM"
           cboRepTimeTo.AddItem "11:20 AM": cboRepTimeTo.AddItem "11:25 AM": cboRepTimeTo.AddItem "11:30 AM": cboRepTimeTo.AddItem "11:35 AM"
           cboRepTimeTo.AddItem "11:40 AM": cboRepTimeTo.AddItem "11:45 AM": cboRepTimeTo.AddItem "11:50 AM": cboRepTimeTo.AddItem "11:55 AM"
           cboRepTimeTo.AddItem "12:00 PM": cboRepTimeTo.AddItem "12:05 PM": cboRepTimeTo.AddItem "12:10 PM": cboRepTimeTo.AddItem "12:15 PM"
           cboRepTimeTo.AddItem "12:20 PM": cboRepTimeTo.AddItem "12:25 PM": cboRepTimeTo.AddItem "12:30 PM": cboRepTimeTo.AddItem "12:35 PM"
           cboRepTimeTo.AddItem "01:00 PM": cboRepTimeTo.AddItem "01:05 PM": cboRepTimeTo.AddItem "01:10 PM": cboRepTimeTo.AddItem "01:15 PM"
           cboRepTimeTo.AddItem "01:20 PM": cboRepTimeTo.AddItem "01:25 PM": cboRepTimeTo.AddItem "01:30 PM": cboRepTimeTo.AddItem "01:35 PM"
           cboRepTimeTo.AddItem "01:40 PM": cboRepTimeTo.AddItem "01:45 PM": cboRepTimeTo.AddItem "01:50 PM": cboRepTimeTo.AddItem "01:55 PM"
           cboRepTimeTo.AddItem "02:00 PM": cboRepTimeTo.AddItem "02:05 PM": cboRepTimeTo.AddItem "02:10 PM": cboRepTimeTo.AddItem "02:15 PM"
           cboRepTimeTo.AddItem "02:20 PM": cboRepTimeTo.AddItem "02:25 PM": cboRepTimeTo.AddItem "02:30 PM": cboRepTimeTo.AddItem "02:35 PM"
           cboRepTimeTo.AddItem "02:40 PM": cboRepTimeTo.AddItem "02:45 PM": cboRepTimeTo.AddItem "02:50 PM": cboRepTimeTo.AddItem "02:55 PM"
           cboRepTimeTo.AddItem "03:00 PM": cboRepTimeTo.AddItem "03:05 PM": cboRepTimeTo.AddItem "03:10 PM": cboRepTimeTo.AddItem "03:15 PM"
           cboRepTimeTo.AddItem "03:20 PM": cboRepTimeTo.AddItem "03:25 PM": cboRepTimeTo.AddItem "03:30 PM": cboRepTimeTo.AddItem "03:35 PM"
           cboRepTimeTo.AddItem "03:40 PM": cboRepTimeTo.AddItem "03:45 PM": cboRepTimeTo.AddItem "03:50 PM": cboRepTimeTo.AddItem "03:55 PM"
           cboRepTimeTo.AddItem "04:00 PM": cboRepTimeTo.AddItem "04:05 PM": cboRepTimeTo.AddItem "04:10 PM": cboRepTimeTo.AddItem "04:15 PM"
           cboRepTimeTo.AddItem "04:20 PM": cboRepTimeTo.AddItem "04:25 PM": cboRepTimeTo.AddItem "04:30 PM": cboRepTimeTo.AddItem "04:35 PM"
           cboRepTimeTo.AddItem "04:40 PM": cboRepTimeTo.AddItem "04:45 PM": cboRepTimeTo.AddItem "04:50 PM": cboRepTimeTo.AddItem "04:55 PM"
           cboRepTimeTo.AddItem "05:00 PM": cboRepTimeTo.AddItem "05:05 PM": cboRepTimeTo.AddItem "05:10 PM": cboRepTimeTo.AddItem "05:15 PM"
           cboRepTimeTo.AddItem "05:20 PM": cboRepTimeTo.AddItem "05:25 PM": cboRepTimeTo.AddItem "05:30 PM": cboRepTimeTo.AddItem "05:35 PM"
           cboRepTimeTo.AddItem "05:40 PM": cboRepTimeTo.AddItem "05:45 PM": cboRepTimeTo.AddItem "05:50 PM": cboRepTimeTo.AddItem "05:55 PM"
           cboRepTimeTo.AddItem "06:00 PM": cboRepTimeTo.AddItem "06:05 PM": cboRepTimeTo.AddItem "06:10 PM": cboRepTimeTo.AddItem "06:15 PM"
           cboRepTimeTo.AddItem "06:20 PM": cboRepTimeTo.AddItem "06:25 PM": cboRepTimeTo.AddItem "06:30 PM": cboRepTimeTo.AddItem "06:35 PM"
           cboRepTimeTo.AddItem "06:40 PM": cboRepTimeTo.AddItem "06:45 PM": cboRepTimeTo.AddItem "06:50 PM": cboRepTimeTo.AddItem "06:55 PM"
           cboRepTimeTo.AddItem "07:00 PM": cboRepTimeTo.AddItem "07:05 PM": cboRepTimeTo.AddItem "07:10 PM": cboRepTimeTo.AddItem "07:15 PM"
           cboRepTimeTo.AddItem "07:20 PM": cboRepTimeTo.AddItem "07:25 PM": cboRepTimeTo.AddItem "07:30 PM": cboRepTimeTo.AddItem "07:35 PM"
           cboRepTimeTo.AddItem "07:40 PM": cboRepTimeTo.AddItem "07:45 PM": cboRepTimeTo.AddItem "07:50 PM": cboRepTimeTo.AddItem "07:55 PM"
           cboRepTimeTo.AddItem "08:00 PM": cboRepTimeTo.AddItem "08:05 PM": cboRepTimeTo.AddItem "08:10 PM": cboRepTimeTo.AddItem "08:15 PM"
           cboRepTimeTo.AddItem "08:20 PM": cboRepTimeTo.AddItem "08:25 PM": cboRepTimeTo.AddItem "08:30 PM": cboRepTimeTo.AddItem "08:35 PM"
           cboRepTimeTo.AddItem "08:40 PM": cboRepTimeTo.AddItem "08:45 PM": cboRepTimeTo.AddItem "08:50 PM": cboRepTimeTo.AddItem "08:55 PM"
       Case "EVERY 20 MIN"
           cboRepTimeFrom.AddItem "07:00 AM": cboRepTimeFrom.AddItem "07:20 AM": cboRepTimeFrom.AddItem "07:40 AM"
           cboRepTimeFrom.AddItem "08:00 AM": cboRepTimeFrom.AddItem "08:20 AM": cboRepTimeFrom.AddItem "08:40 AM"
           cboRepTimeFrom.AddItem "09:00 AM": cboRepTimeFrom.AddItem "09:20 AM": cboRepTimeFrom.AddItem "09:40 AM"
           cboRepTimeFrom.AddItem "10:00 AM": cboRepTimeFrom.AddItem "10:20 AM": cboRepTimeFrom.AddItem "10:40 AM"
           cboRepTimeFrom.AddItem "11:00 AM": cboRepTimeFrom.AddItem "11:20 AM": cboRepTimeFrom.AddItem "11:40 AM"
           cboRepTimeFrom.AddItem "12:00 PM": cboRepTimeFrom.AddItem "12:20 PM": cboRepTimeFrom.AddItem "12:40 PM"
           cboRepTimeFrom.AddItem "01:00 PM": cboRepTimeFrom.AddItem "01:20 PM": cboRepTimeFrom.AddItem "01:40 PM"
           cboRepTimeFrom.AddItem "02:00 PM": cboRepTimeFrom.AddItem "02:20 PM": cboRepTimeFrom.AddItem "02:40 PM"
           cboRepTimeFrom.AddItem "03:00 PM": cboRepTimeFrom.AddItem "03:20 PM": cboRepTimeFrom.AddItem "03:40 PM"
           cboRepTimeFrom.AddItem "04:00 PM": cboRepTimeFrom.AddItem "04:20 PM": cboRepTimeFrom.AddItem "04:40 PM"
           cboRepTimeFrom.AddItem "05:00 PM": cboRepTimeFrom.AddItem "05:20 PM": cboRepTimeFrom.AddItem "05:40 PM"
           cboRepTimeFrom.AddItem "06:00 PM": cboRepTimeFrom.AddItem "06:20 PM": cboRepTimeFrom.AddItem "06:40 PM"
           cboRepTimeFrom.AddItem "07:00 PM": cboRepTimeFrom.AddItem "07:20 PM": cboRepTimeFrom.AddItem "07:40 PM"
           cboRepTimeFrom.AddItem "08:00 PM": cboRepTimeFrom.AddItem "08:20 PM": cboRepTimeFrom.AddItem "08:40 PM"
           
           cboRepTimeTo.AddItem "07:00 AM": cboRepTimeTo.AddItem "07:20 AM": cboRepTimeTo.AddItem "07:40 AM"
           cboRepTimeTo.AddItem "08:00 AM": cboRepTimeTo.AddItem "08:20 AM": cboRepTimeTo.AddItem "08:40 AM"
           cboRepTimeTo.AddItem "09:00 AM": cboRepTimeTo.AddItem "09:20 AM": cboRepTimeTo.AddItem "09:40 AM"
           cboRepTimeTo.AddItem "10:00 AM": cboRepTimeTo.AddItem "10:20 AM": cboRepTimeTo.AddItem "10:40 AM"
           cboRepTimeTo.AddItem "11:00 AM": cboRepTimeTo.AddItem "11:20 AM": cboRepTimeTo.AddItem "11:40 AM"
           cboRepTimeTo.AddItem "12:00 PM": cboRepTimeTo.AddItem "12:20 PM": cboRepTimeTo.AddItem "12:40 PM"
           cboRepTimeTo.AddItem "01:00 PM": cboRepTimeTo.AddItem "01:20 PM": cboRepTimeTo.AddItem "01:40 PM"
           cboRepTimeTo.AddItem "02:00 PM": cboRepTimeTo.AddItem "02:20 PM": cboRepTimeTo.AddItem "02:40 PM"
           cboRepTimeTo.AddItem "03:00 PM": cboRepTimeTo.AddItem "03:20 PM": cboRepTimeTo.AddItem "03:40 PM"
           cboRepTimeTo.AddItem "04:00 PM": cboRepTimeTo.AddItem "04:20 PM": cboRepTimeTo.AddItem "04:40 PM"
           cboRepTimeTo.AddItem "05:00 PM": cboRepTimeTo.AddItem "05:20 PM": cboRepTimeTo.AddItem "05:40 PM"
           cboRepTimeTo.AddItem "06:00 PM": cboRepTimeTo.AddItem "06:20 PM": cboRepTimeTo.AddItem "06:40 PM"
           cboRepTimeTo.AddItem "07:00 PM": cboRepTimeTo.AddItem "07:20 PM": cboRepTimeTo.AddItem "07:40 PM"
           cboRepTimeTo.AddItem "08:00 PM": cboRepTimeTo.AddItem "08:20 PM": cboRepTimeTo.AddItem "08:40 PM"
       Case "EVERY 10 MIN"
        cboRepTimeFrom.AddItem "07:00 AM": cboRepTimeFrom.AddItem "07:10 AM": cboRepTimeFrom.AddItem "07:20 AM": cboRepTimeFrom.AddItem "07:30 AM": cboRepTimeFrom.AddItem "07:40 AM": cboRepTimeFrom.AddItem "07:50 AM"
        cboRepTimeFrom.AddItem "08:00 AM": cboRepTimeFrom.AddItem "08:10 AM": cboRepTimeFrom.AddItem "08:20 AM": cboRepTimeFrom.AddItem "08:30 AM": cboRepTimeFrom.AddItem "08:40 AM": cboRepTimeFrom.AddItem "08:50 AM"
        cboRepTimeFrom.AddItem "09:00 AM": cboRepTimeFrom.AddItem "09:10 AM": cboRepTimeFrom.AddItem "09:20 AM": cboRepTimeFrom.AddItem "09:30 AM": cboRepTimeFrom.AddItem "09:40 AM": cboRepTimeFrom.AddItem "09:50 AM"
        cboRepTimeFrom.AddItem "10:00 AM": cboRepTimeFrom.AddItem "10:10 AM": cboRepTimeFrom.AddItem "10:20 AM": cboRepTimeFrom.AddItem "10:30 AM": cboRepTimeFrom.AddItem "10:40 AM": cboRepTimeFrom.AddItem "10:50 AM"
        cboRepTimeFrom.AddItem "11:00 AM": cboRepTimeFrom.AddItem "11:10 AM": cboRepTimeFrom.AddItem "11:20 AM": cboRepTimeFrom.AddItem "11:30 AM": cboRepTimeFrom.AddItem "11:40 AM": cboRepTimeFrom.AddItem "11:50 AM"
        cboRepTimeFrom.AddItem "12:00 PM": cboRepTimeFrom.AddItem "12:10 PM": cboRepTimeFrom.AddItem "12:20 PM": cboRepTimeFrom.AddItem "12:30 PM": cboRepTimeFrom.AddItem "12:40 PM": cboRepTimeFrom.AddItem "12:50 PM"
        cboRepTimeFrom.AddItem "01:00 PM": cboRepTimeFrom.AddItem "01:10 PM": cboRepTimeFrom.AddItem "01:20 PM": cboRepTimeFrom.AddItem "01:30 PM": cboRepTimeFrom.AddItem "01:40 PM": cboRepTimeFrom.AddItem "01:50 PM"
        cboRepTimeFrom.AddItem "02:00 PM": cboRepTimeFrom.AddItem "02:10 PM": cboRepTimeFrom.AddItem "02:20 PM": cboRepTimeFrom.AddItem "02:30 PM": cboRepTimeFrom.AddItem "02:40 PM": cboRepTimeFrom.AddItem "02:50 PM"
        cboRepTimeFrom.AddItem "03:00 PM": cboRepTimeFrom.AddItem "03:10 PM": cboRepTimeFrom.AddItem "03:20 PM": cboRepTimeFrom.AddItem "03:30 PM": cboRepTimeFrom.AddItem "03:40 PM": cboRepTimeFrom.AddItem "03:50 PM"
        cboRepTimeFrom.AddItem "04:00 PM": cboRepTimeFrom.AddItem "04:10 PM": cboRepTimeFrom.AddItem "04:20 PM": cboRepTimeFrom.AddItem "04:30 PM": cboRepTimeFrom.AddItem "04:40 PM": cboRepTimeFrom.AddItem "04:50 PM"
        cboRepTimeFrom.AddItem "05:00 PM": cboRepTimeFrom.AddItem "05:10 PM": cboRepTimeFrom.AddItem "05:20 PM": cboRepTimeFrom.AddItem "05:30 PM": cboRepTimeFrom.AddItem "05:40 PM": cboRepTimeFrom.AddItem "05:50 PM"
        cboRepTimeFrom.AddItem "06:00 PM": cboRepTimeFrom.AddItem "06:10 PM": cboRepTimeFrom.AddItem "06:20 PM": cboRepTimeFrom.AddItem "06:30 PM": cboRepTimeFrom.AddItem "06:40 PM": cboRepTimeFrom.AddItem "06:50 PM"
        cboRepTimeFrom.AddItem "07:00 PM": cboRepTimeFrom.AddItem "07:10 PM": cboRepTimeFrom.AddItem "07:20 PM": cboRepTimeFrom.AddItem "07:30 PM": cboRepTimeFrom.AddItem "07:40 PM": cboRepTimeFrom.AddItem "07:50 PM"
        cboRepTimeFrom.AddItem "08:00 PM": cboRepTimeFrom.AddItem "08:10 PM": cboRepTimeFrom.AddItem "08:20 PM": cboRepTimeFrom.AddItem "08:30 PM": cboRepTimeFrom.AddItem "08:40 PM": cboRepTimeFrom.AddItem "08:50 PM"
            
        cboRepTimeTo.AddItem "07:00 AM": cboRepTimeTo.AddItem "07:10 AM": cboRepTimeTo.AddItem "07:20 AM": cboRepTimeTo.AddItem "07:30 AM": cboRepTimeTo.AddItem "07:40 AM": cboRepTimeTo.AddItem "07:50 AM"
        cboRepTimeTo.AddItem "08:00 AM": cboRepTimeTo.AddItem "08:10 AM": cboRepTimeTo.AddItem "08:20 AM": cboRepTimeTo.AddItem "08:30 AM": cboRepTimeTo.AddItem "08:40 AM": cboRepTimeTo.AddItem "08:50 AM"
        cboRepTimeTo.AddItem "09:00 AM": cboRepTimeTo.AddItem "09:10 AM": cboRepTimeTo.AddItem "09:20 AM": cboRepTimeTo.AddItem "09:30 AM": cboRepTimeTo.AddItem "09:40 AM": cboRepTimeTo.AddItem "09:50 AM"
        cboRepTimeTo.AddItem "10:00 AM": cboRepTimeTo.AddItem "10:10 AM": cboRepTimeTo.AddItem "10:20 AM": cboRepTimeTo.AddItem "10:30 AM": cboRepTimeTo.AddItem "10:40 AM": cboRepTimeTo.AddItem "10:50 AM"
        cboRepTimeTo.AddItem "11:00 AM": cboRepTimeTo.AddItem "11:10 AM": cboRepTimeTo.AddItem "11:20 AM": cboRepTimeTo.AddItem "11:30 AM": cboRepTimeTo.AddItem "11:40 AM": cboRepTimeTo.AddItem "11:50 AM"
        cboRepTimeTo.AddItem "12:00 PM": cboRepTimeTo.AddItem "12:10 PM": cboRepTimeTo.AddItem "12:20 PM": cboRepTimeTo.AddItem "12:30 PM": cboRepTimeTo.AddItem "12:40 PM": cboRepTimeTo.AddItem "12:50 PM"
        cboRepTimeTo.AddItem "01:00 PM": cboRepTimeTo.AddItem "01:10 PM": cboRepTimeTo.AddItem "01:20 PM": cboRepTimeTo.AddItem "01:30 PM": cboRepTimeTo.AddItem "01:40 PM": cboRepTimeTo.AddItem "01:50 PM"
        cboRepTimeTo.AddItem "02:00 PM": cboRepTimeTo.AddItem "02:10 PM": cboRepTimeTo.AddItem "02:20 PM": cboRepTimeTo.AddItem "02:30 PM": cboRepTimeTo.AddItem "02:40 AM": cboRepTimeTo.AddItem "02:50 PM"
        cboRepTimeTo.AddItem "03:00 PM": cboRepTimeTo.AddItem "03:10 PM": cboRepTimeTo.AddItem "03:20 PM": cboRepTimeTo.AddItem "03:30 PM": cboRepTimeTo.AddItem "03:40 PM": cboRepTimeTo.AddItem "03:50 PM"
        cboRepTimeTo.AddItem "04:00 PM": cboRepTimeTo.AddItem "04:10 PM": cboRepTimeTo.AddItem "04:20 PM": cboRepTimeTo.AddItem "04:30 PM": cboRepTimeTo.AddItem "04:40 PM": cboRepTimeTo.AddItem "04:50 PM"
        cboRepTimeTo.AddItem "05:00 PM": cboRepTimeTo.AddItem "05:10 PM": cboRepTimeTo.AddItem "05:20 PM": cboRepTimeTo.AddItem "05:30 PM": cboRepTimeTo.AddItem "05:40 PM": cboRepTimeTo.AddItem "05:50 PM"
        cboRepTimeTo.AddItem "06:00 PM": cboRepTimeTo.AddItem "06:10 PM": cboRepTimeTo.AddItem "06:20 PM": cboRepTimeTo.AddItem "06:30 PM": cboRepTimeTo.AddItem "06:40 PM": cboRepTimeTo.AddItem "06:50 PM"
        cboRepTimeTo.AddItem "07:00 PM": cboRepTimeTo.AddItem "07:10 PM": cboRepTimeTo.AddItem "07:20 PM": cboRepTimeTo.AddItem "07:30 PM": cboRepTimeTo.AddItem "07:40 PM": cboRepTimeTo.AddItem "07:50 PM"
        cboRepTimeTo.AddItem "08:00 PM": cboRepTimeTo.AddItem "08:10 PM": cboRepTimeTo.AddItem "08:20 PM": cboRepTimeTo.AddItem "08:30 PM": cboRepTimeTo.AddItem "08:40 PM": cboRepTimeTo.AddItem "08:50 PM"
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
    Resume Next
End Function

Private Function Load_Providers()
Dim strSQL As String
Dim sName As String
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim strSuffix As String
Dim lngProvID As Long
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim cnProv As New ADODB.Connection
Dim rsProv As New ADODB.Recordset

    On Error GoTo Error_Load_Providers
    cboProviders.Clear
    If cnProv.State = adStateClosed Then cnProv.open psConnect(0)
    strSQL = "Load_Active_Providers"
    Set rsProv = cnProv.Execute(strSQL)
    With rsProv
        Do Until .EOF
            strFirst = vbNullString
            strLast = vbNullString
            strMiddle = vbNullString
            strTitle = vbNullString
            strSuffix = vbNullString
            sName = vbNullString
            lngProvID = 0
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "FIRST_NAME"
                                strFirst = Fix_Name(Fld.Value)
                            Case "MIDDLE_NAME"
                                strMiddle = Fix_Name(Fld.Value)
                            Case "LAST_NAME"
                                strLast = Fix_Name(Fld.Value)
                            Case "SUFFIX"
                                strSuffix = Fix_Suffix(Fld.Value)
                            Case "TITLE"
                                strTitle = Fix_Title(Fld.Value)
                            Case "PROVIDER_ID"
                                lngProvID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If Trim(strFirst) <> vbNullString Then
                strFirst = Fix_Name(strFirst)
                sName = strFirst
             End If
            If Trim(strMiddle) <> vbNullString Then
                 sName = sName & " " & Left(strMiddle, 1)
             End If
            If Trim(strLast) <> vbNullString Then
                strLast = Fix_Name(strLast)
                sName = sName & " " & strLast
             End If
            If Trim(strSuffix) <> vbNullString Then
                strSuffix = Fix_Suffix(strSuffix)
                sName = sName & ", " & Trim(strSuffix)
             End If
            If Trim(strTitle) <> vbNullString Then
                 sName = sName & ", " & Trim(strTitle)
             End If
             If Trim(sName) <> vbNullString Then
                sName = Replace(sName, ",,", ",")
                cboProviders.AddItem sName
                cboProviders.ItemData(cboProviders.NewIndex) = lngProvID
            End If
        Loop
    End With
    If cnProv.State = adStateOpen Then
        cnProv.Close
        Set cnProv = Nothing
        Set rsProv = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
   Exit Function
Error_Load_Providers:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Providers", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Function Add_ConflictRow()
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr

    lngRow = grdTerms.MaxRows + 1
    grdTerms.MaxRows = lngRow
    With grdTerms
        lngCol = 1
        '.ColWidth(lngCol) = 20
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", udtSchedPatient.Collection, "TypeHAlignCenter", "PATIENT"
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdTerms.SetRowItemData lngRow, Val(udtSchedPatient.IntId)
        lngCol = 2
        '.ColWidth(lngCol) = 19
        Get_Appt_Providers
        Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", clsGridFormat.List, "TypeHAlignCenter", "PROVIDER_ID"
        Set_Appt_Providers udtBlock.ProviderID, lngRow, lngCol, "GrdTerms"
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        '.ColWidth(lngCol) = 12
        clsGridFormat.List = vbNullString
        Load_Facilities
        Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", clsGridFormat.List, "TypeHAlignCenter", "FACILITY_ID"
        Set_Appt_Facility udtBlock.FacilityID, lngRow, lngCol, "GrdTerms"
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
       ' .ColWidth(lngCol) = 17
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", udtAppointment.SchedDate, "TypeHAlignCenter", "APPT_DATE"
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        '.ColWidth(lngCol) = 12
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", udtAppointment.SchedTime, "TypeHAlignCenter", "APPT_TIME"
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        '.ColWidth(lngCol) = 12
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", udtSchedPatient.Duration, "TypeHAlignCenter", "DURATION"
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .Redraw = False
    End With
    Exit Function
End Function

Private Function Load_Facilities()
Dim strSQL As String
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset
Dim strName As String
Dim Arr
Dim Idx%
Dim lngRow As Long

    On Error GoTo Error_Load_Facilities
    Screen.MousePointer = vbHourglass
    clsGridFormat.List = vbNullString
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "Load__Active_Facilities"
    Set rsFac = cnFac.Execute(strSQL)
    Do Until rsFac.EOF
        If Not IsNull(rsFac.Fields("FACILITY_NAME")) Then
            If rsFac.Fields("FACILITY_NAME") <> vbNullString Then
                strName = vbNullString
                strName = rsFac.Fields("FACILITY_NAME")
                strName = Fix_Name(strName)
                strName = Fix_Facility_Name(strName)
                If TextFound("-", strName, Len("-"), True) Then
                    Arr = Split(strName, "-")
                    If UBound(Arr) > 0 Then
                        strName = vbNullString
                        For Idx = LBound(Arr) To UBound(Arr)
                            If Arr(0) <> vbNullString Then
                                If Arr(0) = "Ucs" Then
                                    strName = UCase(Arr(0))
                                Else
                                    strName = Arr(0)
                                End If
                            End If
                            If Idx >= 1 Then
                                strName = strName & " - " & Arr(Idx)
                            End If
                        Next Idx
                    End If
                End If
                If Trim(strName) <> vbNullString And Len(strName) > 2 Then
                    If clsGridFormat.List = vbNullString Then
                        clsGridFormat.List = strName
                    ElseIf Not TextFound(strName, clsGridFormat.List, Len(strName), True) Then
                        clsGridFormat.List = clsGridFormat.List & Chr(9) & strName
                    End If
                End If
            End If
        End If
        rsFac.MoveNext
    Loop
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set cnFac = Nothing
        Set rsFac = Nothing
    End If
    Screen.MousePointer = vbDefault
Exit Function
Error_Load_Facilities:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Facilities", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Add_Row()
Dim lngRow As Long
Dim lngCol As Long

    lngRow = grdBlock.MaxRows + 1
    grdBlock.MaxRows = lngRow
    mlngGridRow = lngRow
        clsGridFormat.List = vbNullString
        lngCol = 1
        Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", " ", "TypeHAlignLeft", "FACILITY_ID"
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", " ", "TypeHAlignLeft", "REASON"
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        clsGridFormat.List = Get_BlockTerms
        Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", " ", "TypeHAlignLeft", "START_TIME"
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        clsGridFormat.List = Get_BlockTerms
        Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", " ", "TypeHAlignLeft", "END_TIME"
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", " ", "TypeHAlignLeft", "START_DATE"
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", " ", "TypeHAlignLeft", "END_DATE"
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        clsGridFormat.List = Get_GridTerms("sd_Appointment_BlockDuration")
        Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "DURATION"
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdBlock.Col = 8
        lngCol = 8
        Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", " ", "TypeHAlignCenter", "DAYWEEK"
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        
        Exit Function
End Function

Private Function Get_BlockTerms() As String
Dim Idx As Integer
Dim Arr

    Screen.MousePointer = vbHourglass
    clsGridFormat.List = vbNullString
    mInterval = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "SchedulingInterval")
    mInterval = UCase(mInterval)
    Arr = Split(mInterval, " ")
    If Arr(1) = 15 Then mInterval = "EVERY 15 MIN" Else mInterval = "EVERY 20 MIN"
    For Idx = mStartIdx To UBound(mStartHour)
        clsGridFormat.List = mStartHour(Idx) & " AM"
        If UCase(mInterval) = "EVERY HALF HOUR" Then
            clsGridFormat.List = clsGridFormat.List & Chr(9) & mStartHour(Idx) & ":30 AM"
        ElseIf UCase(mInterval) = "EVERY 15 MIN" Then
            clsGridFormat.List = clsGridFormat.List & Chr(9) & Mid$(mStartHour(Idx), 1, 2) & ":15 AM"
            clsGridFormat.List = clsGridFormat.List & Chr(9) & Mid$(mStartHour(Idx), 1, 2) & ":30 AM"
            clsGridFormat.List = clsGridFormat.List & Chr(9) & Mid$(mStartHour(Idx), 1, 2) & ":45 AM"
        Else
            clsGridFormat.List = clsGridFormat.List & Chr(9) & Mid$(mStartHour(Idx), 1, 2) & ":20 AM"
            clsGridFormat.List = clsGridFormat.List & Chr(9) & Mid$(mStartHour(Idx), 1, 2) & ":40 AM"
        End If
    Next Idx
    For Idx = 0 To mEndIdx
           clsGridFormat.List = clsGridFormat.List & Chr(9) & mEndHour(Idx) & " PM"
        If UCase(mInterval) = "EVERY HALF HOUR" Then
            clsGridFormat.List = clsGridFormat.List & Chr(9) & mEndHour(Idx) & ":30 PM"
        ElseIf UCase(mInterval) = "EVERY 15 MIN" Then
            clsGridFormat.List = clsGridFormat.List & Chr(9) & mEndHour(Idx) & ":15 PM"
            clsGridFormat.List = clsGridFormat.List & Chr(9) & mEndHour(Idx) & ":30 PM"
            clsGridFormat.List = clsGridFormat.List & Chr(9) & mEndHour(Idx) & ":45 PM"
        Else
            clsGridFormat.List = clsGridFormat.List & Chr(9) & mEndHour(Idx) & ":20 PM"
            clsGridFormat.List = clsGridFormat.List & Chr(9) & mEndHour(Idx) & ":40 PM"
        End If
    Next Idx
    Screen.MousePointer = vbDefault
    Exit Function
End Function
Private Function Format_Grid()
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr

    With grdBlock
        .Redraw = False
        .ClearRange 1, 1, .MaxCols, .MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 10
        .MaxCols = 8
        .FontName = "Georgia"
        .GridShowHoriz = True
        .GridShowVert = True
        .RowHeadersShow = False
        .ColHeaderRows = 2
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        lngRow = SpreadHeader
        .RowHeight(lngRow) = 20
        
        lngCol = 1
        .ColWidth(lngCol) = 30.2
        Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", "FACILITY", "TypeHAlignCenter", "FACILITY_ID", HeaderColor
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        .ColWidth(lngCol) = 16.9
        Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", "REASON", "TypeHAlignCenter", "REASON", HeaderColor
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        .ColWidth(lngCol) = 12
        Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", "START TIME", "TypeHAlignCenter", "START_TIME", HeaderColor
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        .ColWidth(lngCol) = 12
        Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", "END TIME", "TypeHAlignCenter", "END_TIME", HeaderColor
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        .ColWidth(lngCol) = 12
        Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", "START DATE", "TypeHAlignCenter", "START_DATE", HeaderColor
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        .ColWidth(lngCol) = 12
        Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", "END DATE", "TypeHAlignCenter", "END_DATE", HeaderColor
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        .ColWidth(lngCol) = 12
        Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", "DURATION", "TypeHAlignCenter", "DURATION", HeaderColor
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 8
        .ColWidth(lngCol) = 11
        Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", "DAY OF WEEK", "TypeHAlignCenter", "DAYWEEK", HeaderColor
        grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdBlock
        .Redraw = False
    End With
    Exit Function
End Function

Private Function Format_Conflict()
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr

    With grdTerms
        .Redraw = False
        .ClearRange 1, 1, .MaxCols, .MaxRows, False
        .ProcessTab = True
        '.FontBold = True
        .FontSize = 10
        .MaxCols = 6
        .FontName = "Georgia"
        .GridShowHoriz = True
        .GridShowVert = True
        .RowHeadersShow = False
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .MaxRows = 0
        lngRow = SpreadHeader + 1
        .RowHeight(lngRow) = 20
        lngCol = 1
        .ColWidth(lngCol) = 25
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "PATIENT", "TypeHAlignCenter", "PATIENT", HeaderColor
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        .ColWidth(lngCol) = 19
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "PROVIDER", "TypeHAlignCenter", "PROVIDER", HeaderColor
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        .ColWidth(lngCol) = 12
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "FACILITY", "TypeHAlignCenter", "FACILITY", HeaderColor
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        .ColWidth(lngCol) = 17
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "APPT DATE", "TypeHAlignCenter", "APPT_DATE", HeaderColor
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        .ColWidth(lngCol) = 12
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "APPT TIME", "TypeHAlignCenter", "APPT_TIME", HeaderColor
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        .ColWidth(lngCol) = 12
        Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "DURATION", "TypeHAlignCenter", "DURATION", HeaderColor
        grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        
        Format_CellTypeEdit grdTerms, SpreadHeader, 1, "6", "PATIENT SCHEDULING CONFLICT", "TypeHAlignCenter", "", HeaderColor
        grdTerms.SetCellBorder 1, SpreadHeader, 6, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 2, grdTerms
        .Redraw = False
    End With
    Exit Function
End Function
Private Function Format_CellTypeCombo(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, _
vstrString As String, vstrHAlign As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
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
              .TypeComboBoxhWnd = fpTerm.Hwnd
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
              .FontName = "Georgia"
              .FontSize = 8
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
          .BlockMode = False
    End With
End Function
Private Function LookUp_Facility(vstrFacility As String) As Long
Dim strSQL As String
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset
Dim lngFacID As Long
Dim strFacility As String
Dim strCompareFac As String
Dim strFacName As String
Dim Arr
Dim Idx%

    On Error GoTo Error_LookUp_Facility
    Screen.MousePointer = vbHourglass
    strFacility = Replace(vstrFacility, ",", ", ")
    strFacility = Replace(vstrFacility, ".", "")
    strFacility = Replace(vstrFacility, "  -  ", " - ")
    strFacility = Replace(vstrFacility, ".", "")
    'strFacility = Replace(vstrFacility, "-", " - ")
    strFacility = Trim(strFacility)
    Arr = Split(strFacility, "-")
    If UBound(Arr) > 0 Then
        strFacility = vbNullString
        For Idx = LBound(Arr) To UBound(Arr)
            If Trim(Arr(Idx)) <> vbNullString Then
                If Trim(strFacility) = vbNullString Then
                   strFacility = Trim(Arr(Idx))
                Else
                   strFacility = strFacility & " - " & Trim(Arr(Idx))
                End If
            End If
        Next Idx
    Else
        strFacility = Trim(strFacility)
    End If
    clsGridFormat.List = vbNullString
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "Load__Active_Facilities"
    Set rsFac = cnFac.Execute(strSQL)
    Do Until rsFac.EOF
        If Not IsNull(rsFac.Fields("FACILITY_NAME")) Then
            lngFacID = rsFac.Fields("FACILITY_ID")
            strCompareFac = rsFac.Fields("FACILITY_NAME")
            strCompareFac = Replace(Trim(strCompareFac), ".", "")
            strCompareFac = Replace(Trim(strCompareFac), "  ", " ")
            strCompareFac = Trim(strCompareFac)
            If strCompareFac <> vbNullString Then
                Arr = Split(strCompareFac, "-")
                If UBound(Arr) > 0 Then
                    strCompareFac = vbNullString
                    For Idx = LBound(Arr) To UBound(Arr)
                        If Trim(Arr(Idx)) <> vbNullString Then
                            If Trim(strCompareFac) = vbNullString Then
                               strCompareFac = Trim(Arr(Idx))
                            Else
                               strCompareFac = strCompareFac & " - " & Trim(Arr(Idx))
                            End If
                        End If
                    Next Idx
                Else
                    strCompareFac = Trim(strCompareFac)
                End If
                strCompareFac = Trim(strCompareFac)
                If IsProperCase(Trim(strFacility)) = IsProperCase(Trim(strCompareFac)) Then
                    LookUp_Facility = lngFacID
                    Exit Do
                End If
            End If
        End If
        rsFac.MoveNext
    Loop
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set cnFac = Nothing
        Set rsFac = Nothing
    End If
    Screen.MousePointer = vbDefault
Exit Function
Error_LookUp_Facility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUp_Facility", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function LookUp_Provider(vstrProvider As String) As Long
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

    On Error GoTo Error_LookUp_Provider
    Screen.MousePointer = vbHourglass
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    strSQL = "Get_Appt_Providers " & udtFacility.FacID
     Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        Do Until .EOF
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            Clear_Provider
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
                If TextFound(Replace(clsProvider.Name, ".", ""), Replace(vstrProvider, ".", ""), Len(Replace(clsProvider.Name, ".", "")), True) Then
                    LookUp_Provider = lngProvID
                    Exit Do
                End If
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
Error_LookUp_Provider:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUp_Provider", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function LookUpCellTag(vstrName As String) As String
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String

    With grdBlock
        LookUpCellTag = vbNullString
        lngRow = SpreadHeader
        For lngCol = 1 To grdBlock.MaxCols
            .Col = lngCol
            If .CellTag <> vbNullString Then strTag = .CellTag
            If Trim(strTag) <> vbNullString And TextFound(vstrName, Trim(strTag), Len(vstrName), True) Then
                LookUpCellTag = lngCol & "|" & lngRow
                Exit For
            End If
        Next lngCol
    End With
    Exit Function
End Function

Private Function Read_Blocks_Provider_All()
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim strDate As String


    On Error GoTo Error_Read_Blocks_ByProvider
    Screen.MousePointer = vbHourglass
    grdBlock.ClearRange 1, 1, grdBlock.MaxCols, grdBlock.MaxRows, False
    grdBlock.MaxRows = 0
    If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
    strSQL = "Read_BlockTime_ProviderID " & udtBlock.ProviderID
    Set rsBlock = cnBlock.Execute(strSQL)
    Do Until rsBlock.EOF
        Clear_BlockTerm
        Set Flds = rsBlock.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "BLOCK_ID"
                                lngPrimKey = Trim(Fld.Value)
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
                            Case "DAYWEEK"
                                 udtBlock.DayWeek = Trim(Fld.Value)
                        End Select
                    End If
                End If
            Next
            rsBlock.MoveNext
            udtBlock.Reason = Replace(udtBlock.Reason, "''", "'")
            lngRow = grdBlock.MaxRows + 1
            grdBlock.MaxRows = lngRow
            clsGridFormat.List = vbNullString
            grdBlock.SetRowItemData lngRow, lngPrimKey
            lngCol = 1
            clsGridFormat.List = vbNullString
            Load_Facilities
            Format_CellTypeCombo grdBlock, lngRow, lngCol, "1", clsGridFormat.List, "TypeHAlignLeft", "FACILITY_ID"
            Set_Appt_Facility udtBlock.FacilityID, lngRow, lngCol
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 2
            clsGridFormat.List = vbNullString
            clsGridFormat.List = Get_GridTerms("sd_Appointment_BlockReason")
            Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", IsProperCase(udtBlock.Reason), "TypeHAlignLeft", "REASON"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 3
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", udtBlock.StartTime, "TypeHAlignCenter", "START_TIME"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 4
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", udtBlock.EndTime, "TypeHAlignCenter", "END_TIME"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 5
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", Format(udtBlock.StartDate, "mm/dd/yyyy"), "TypeHAlignCenter", "START_DATE"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 6
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", Format(udtBlock.EndDate, "mm/dd/yyyy"), "TypeHAlignCenter", "END_DATE"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 7
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", IsProperCase(udtBlock.Duration), "TypeHAlignCenter", "DURATION", "TypeHAlignCenter", 0, ""
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 8
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", IsProperCase(udtBlock.DayWeek), "TypeHAlignCenter", "DAYWEEK", "TypeHAlignCenter", 0, ""
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Loop
    If cnBlock.State = adStateClosed Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Blocks_ByProvider:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Read_Blocks_ByProvider", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Set_Appt_Facility(vlngFacID As Long, Optional vlngRow As Long, Optional vlngCol As Long, Optional vstrGrid As String)
Dim strSQL As String
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset
Dim lngFacID As Long
Dim strFac As String
Dim Arr
Dim Idx%

    On Error GoTo Error_Load_Facilities
    Screen.MousePointer = vbHourglass
    lngFacID = vlngFacID
    Clear_Facility
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "Get_Facility_By_ID " & lngFacID
    Set rsFac = cnFac.Execute(strSQL)
    If Not rsFac.EOF Then
        If Not IsNull(rsFac.Fields("FACILITY_NAME")) Then
            If rsFac.Fields("FACILITY_NAME") <> vbNullString Then
                strFac = rsFac.Fields("FACILITY_NAME")
                strFac = Fix_Name(strFac)
                strFac = Fix_Facility_Name(strFac)
            End If
        End If
    End If
    If TextFound("-", strFac, Len("-"), True) Then
        Arr = Split(strFac, "-")
        If UBound(Arr) > 0 Then
            strFac = vbNullString
            For Idx = LBound(Arr) To UBound(Arr)
                If Arr(0) <> vbNullString Then
                    If Arr(0) = "Ucs" Then
                        strFac = UCase(Arr(0))
                    Else
                        strFac = Arr(0)
                    End If
                End If
                If Idx >= 1 Then
                    strFac = strFac & " - " & Arr(Idx)
                End If
            Next Idx
        End If
    End If
    If strFac <> vbNullString Then
        If Trim(vstrGrid) <> vbNullString Then
            Format_CellTypeStatic grdTerms, vlngRow, vlngCol, "1", strFac, "TypeHAlignLeft", "FACILITY_ID", "TypeHAlignLeft", 0, CStr(lngFacID)
        Else
            Format_CellTypeStatic grdBlock, vlngRow, vlngCol, "1", strFac, "TypeHAlignLeft", "FACILITY_ID", "TypeHAlignLeft", 0, CStr(lngFacID)
        End If
    End If
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set cnFac = Nothing
        Set rsFac = Nothing
    End If
    Screen.MousePointer = vbDefault
    vstrGrid = vbNullString
Exit Function
Error_Load_Facilities:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Facilities", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
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
Dim strSelect As String
Dim Arr
Dim strStr  As String
Dim strTest As String
Dim ITest As Integer

    strTest = Trim(vstrData)
    Arr = Split(strTest, "-")
    If UBound(Arr) > 0 Then
        strTest = vbNullString
        For Idx = LBound(Arr) To UBound(Arr)
            If strTest = vbNullString Then
                strTest = Arr(Idx)
            Else
                strTest = strTest & " " & Arr(Idx)
            End If
        Next Idx
    End If
    With vObjGrid
        .Redraw = False
        .Col = vlngCol
        .Row = vlngRow
        If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        
        If .CellType = CellTypeComboBox Then
            For Idx = 0 To .TypeComboBoxCount
                strStr = vbNullString
                .TypeComboBoxIndex = Idx
                strStr = Trim(.TypeComboBoxString)
                Arr = Split(strStr, "-")
                If UBound(Arr) > 0 Then
                    strStr = vbNullString
                    For ITest = LBound(Arr) To UBound(Arr)
                        If strStr = vbNullString Then
                            strStr = Trim(Arr(ITest))
                        Else
                            strStr = Trim(strStr) & " " & Trim(Arr(ITest))
                        End If
                    Next ITest
                End If
                Arr = Split(strTest, " ")
                If UBound(Arr) > 0 Then
                    strTest = vbNullString
                    For ITest = LBound(Arr) To UBound(Arr)
                        If Trim(Arr(ITest)) <> vbNullString Then
                            If strTest = vbNullString Then
                                strTest = Trim(Arr(ITest))
                            Else
                                strTest = Trim(strTest) & " " & Trim(Arr(ITest))
                            End If
                        End If
                    Next ITest
                End If
                If Trim(strTest) = Trim(strStr) Then
                    .TypeComboBoxCurSel = Idx
                    Exit For
                End If
            Next Idx
        End If
        .TypeHAlign = TypeHAlignLeft
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

Private Function Set_GridCurSel_ItemData(vObjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
Dim Idx As Integer
Dim booData As Boolean
Dim strSpan As String
Dim strTag As String
Dim strCellNote As String
Dim lngCol As Long
Dim lngSpan As Long
Dim strSelect As String

    With vObjGrid
        .Redraw = False
        .TypeHAlign = TypeHAlignLeft
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
                .FontName = "Georgia"
                .FontSize = 10
'                .FontBold = True
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
            .FontName = "Georgia"
            .FontSize = 10
'            .FontBold = True
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

Private Function Validate_Save() As Boolean
Dim Arr
Dim varCellValue
Dim lngCol As Long

    Screen.MousePointer = vbHourglass
    
'    If Validate_BlockTime Then
'        Validate_Save = True
'        Exit Function
'    End If
    
    If cboProviders.Text = vbNullString Then
        MsgBox "Please select a Provider First!", vbInformation, "PatientTrac Block Mode"
        Validate_Save = True
        Exit Function
    End If
    If Val(udtBlock.FacilityID) = 0 Then
        Arr = Split(LookUpCellTag("FACILITY_ID"), "|")
        If UBound(Arr) > 0 Then
          lngCol = Arr(0)
          grdBlock.GetText lngCol, mlngGridRow, varCellValue
          If Trim(varCellValue) <> vbNullString Then
                udtBlock.FacilityID = LookUp_Facility(CStr(varCellValue))
            End If
        End If
        If udtBlock.FacilityID = 0 Then
            MsgBox "Please select a Facility first!", vbInformation, "PatientTrac Block Mode"
            Validate_Save = True
            Exit Function
        End If
    End If
    If Not IsDate(udtBlock.StartDate) Then
         MsgBox "Please select a Start Date first!", vbInformation, "PatientTrac Block Mode"
        Validate_Save = True
        Exit Function
    End If
    If Not IsDate(udtBlock.StartTime) Then
         MsgBox "Please select a Start Time first!", vbInformation, "PatientTrac Block Mode"
        Validate_Save = True
        Exit Function
    End If
    If Not IsDate(udtBlock.EndTime) Then
         MsgBox "Please select a End Time first!", vbInformation, "PatientTrac Block Mode"
        Validate_Save = True
        Exit Function
    End If
    If Not IsDate(udtBlock.EndDate) Then
         MsgBox "Please select a End Date first!", vbInformation, "PatientTrac Block Mode"
        Validate_Save = True
        Exit Function
    End If
    If DateDiff("d", udtBlock.StartDate, udtBlock.EndDate) < 0 Then
         MsgBox "Invalid date range End Date is before Start Date!", vbInformation, "PatientTrac Block Mode"
        Validate_Save = True
        Exit Function
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_BlockTime:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Validate_BlockTime", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Read_Blocks_ByProvider(vlngProvider As Long)
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim strDate As String
Dim lngProv As Long


    On Error GoTo Error_Read_Blocks_ByProvider
    Screen.MousePointer = vbHourglass
    grdBlock.ClearRange 1, 1, grdBlock.MaxCols, grdBlock.MaxRows, False
    grdBlock.MaxRows = 0
    lngProv = vlngProvider
    If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
    strDate = DateAdd("d", -30, Now)
'    strSQL = "Read_BlockTime_ProviderID " & vlngProvider
    strSQL = "SELECT CONVERT(DATETIME, START_DATE, 101) AS START_DATE, " _
        & "CONVERT(DATETIME, END_DATE, 101) AS END_DATE, PROVIDER_ID, BLOCK_ID, " _
        & "CONVERT(VARCHAR(8), START_TIME, 108) AS START_TIME, CONVERT(VARCHAR(8), END_TIME, 108) AS  END_TIME, " _
        & "REASON, DURATION, FACILITY_ID, DAYWEEK FROM APPOINTMENTS_BLOCKS  WHERE " _
        & "PROVIDER_ID = " & lngProv & " AND ACTIVE = 1 AND " _
        & "CONVERT(DATETIME, START_DATE, 101) >  '" & Format(strDate, "mm/dd/yyyy") & "' ORDER " _
        & "BY END_DATE DESC, START_TIME DESC"
    Set rsBlock = cnBlock.Execute(strSQL)
    Do Until rsBlock.EOF
        Clear_BlockTerm
        Set Flds = rsBlock.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "BLOCK_ID"
                                lngPrimKey = Trim(Fld.Value)
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
                            Case "DAYWEEK"
                                 udtBlock.DayWeek = Trim(Fld.Value)
                        End Select
                    End If
                End If
            Next
            rsBlock.MoveNext
            udtBlock.Reason = Replace(udtBlock.Reason, "''", "'")
            lngRow = grdBlock.MaxRows + 1
            grdBlock.MaxRows = lngRow
            clsGridFormat.List = vbNullString
            grdBlock.SetRowItemData lngRow, lngPrimKey
            lngCol = 1
            clsGridFormat.List = vbNullString
            Load_Facilities
            Format_CellTypeCombo grdBlock, lngRow, lngCol, "1", clsGridFormat.List, "TypeHAlignLeft", "FACILITY_ID"
            Set_Appt_Facility udtBlock.FacilityID, lngRow, lngCol
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 2
            clsGridFormat.List = vbNullString
            clsGridFormat.List = Get_GridTerms("sd_Appointment_BlockReason")
            Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", IsProperCase(udtBlock.Reason), "TypeHAlignLeft", "REASON"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 3
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", udtBlock.StartTime, "TypeHAlignCenter", "START_TIME"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 4
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", udtBlock.EndTime, "TypeHAlignCenter", "END_TIME"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 5
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", Format(udtBlock.StartDate, "mm/dd/yyyy"), "TypeHAlignCenter", "START_DATE"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 6
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", Format(udtBlock.EndDate, "mm/dd/yyyy"), "TypeHAlignCenter", "END_DATE"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 7
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", IsProperCase(udtBlock.Duration), "TypeHAlignCenter", "DURATION", "TypeHAlignCenter", 0, ""
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 8
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", IsProperCase(udtBlock.DayWeek), "TypeHAlignCenter", "DAYWEEK", "TypeHAlignCenter", 0, ""
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Loop
    If cnBlock.State = adStateClosed Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Blocks_ByProvider:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Read_Blocks_ByProvider", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Read_Blocks_ByProvider_Date(vlngProvider As Long)
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim lngProv As Long

    On Error GoTo Error_Read_Blocks_ByProvider_Date
    Screen.MousePointer = vbHourglass
    lngProv = vlngProvider
    grdBlock.ClearRange 1, 1, grdBlock.MaxCols, grdBlock.MaxRows, False
    grdBlock.MaxRows = 0
    If udtTime.EndDate = vbNullString Then
        udtTime.EndDate = udtTime.StartDate
    End If
    udtTime.StartDate = Format(udtTime.StartDate, "mm/dd/yyyy")
    udtTime.EndDate = Format(udtTime.EndDate, "mm/dd/yyyy")
    If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
    strSQL = "Read_BlockTime_ProviderID_Date " & lngProv & ",'" _
        & Format(udtTime.StartDate, "mm/dd/yyyy") & "', '" & Format(udtTime.EndDate, "mm/dd/yyyy") & "'"
        Set rsBlock = cnBlock.Execute(strSQL)
    Do Until rsBlock.EOF
        Clear_BlockTerm
        Set Flds = rsBlock.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "BLOCK_ID"
                                lngPrimKey = Trim(Fld.Value)
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
                            Case "DAYWEEK"
                                 udtBlock.DayWeek = Trim(Fld.Value)
                        End Select
                    End If
                End If
            Next
            rsBlock.MoveNext
            udtBlock.Reason = Replace(udtBlock.Reason, "''", "'")
            lngRow = grdBlock.MaxRows + 1
            grdBlock.MaxRows = lngRow
            clsGridFormat.List = vbNullString
            grdBlock.SetRowItemData lngRow, lngPrimKey
            lngCol = 1
            clsGridFormat.List = vbNullString
            Load_Facilities
            Format_CellTypeCombo grdBlock, lngRow, lngCol, "1", clsGridFormat.List, "TypeHAlignCenter", "FACILITY_ID"
            Set_Appt_Facility udtBlock.FacilityID, lngRow, lngCol
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 2
            clsGridFormat.List = vbNullString
            clsGridFormat.List = Get_GridTerms("sd_Appointment_BlockReason")
            Format_CellTypeEdit grdBlock, lngRow, lngCol, "1", IsProperCase(udtBlock.Reason), "TypeHAlignLeft", "REASON"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 3
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", udtBlock.StartTime, "TypeHAlignCenter", "START_TIME"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 4
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", udtBlock.EndTime, "TypeHAlignCenter", "END_TIME"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 5
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", Format(udtBlock.StartDate, "mm/dd/yyyy"), "TypeHAlignCenter", "START_DATE"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 6
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", Format(udtBlock.EndDate, "mm/dd/yyyy"), "TypeHAlignCenter", "END_DATE"
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 7
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", IsProperCase(udtBlock.Duration), "TypeHAlignCenter", "DURATION", "TypeHAlignCenter", 0, ""
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            lngCol = 8
            Format_CellTypeStatic grdBlock, lngRow, lngCol, "1", IsProperCase(udtBlock.DayWeek), "TypeHAlignCenter", "DAYWEEK", "TypeHAlignCenter", 0, ""
            grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Loop
    If cnBlock.State = adStateClosed Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Blocks_ByProvider_Date:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Read_Blocks_ByProvider_Date", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Save_Block_Mode()
Dim sDate As String
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngBlockID As Long
Dim strProvider As String
Dim strFacility As String
Dim strCellTag As String
Dim i%
Dim booExit As Boolean

    On Error GoTo Error_Save_Block_Mode
    Screen.MousePointer = vbHourglass
    grdBlock.Redraw = False
   ' For lngRow = 1 To grdBlock.MaxRows
        lngRow = mlngGridRow
        Clear_BlockTerm
        lngBlockID = 0
        For lngCol = 1 To grdBlock.MaxCols
            grdBlock.Col = lngCol
            grdBlock.Row = lngRow
            If grdBlock.CellTag <> vbNullString Then
                strCellTag = Trim(grdBlock.CellTag)
                grdBlock.GetText lngCol, lngRow, varCellValue
                Select Case UCase(strCellTag)
                    Case "FACILITY_ID"
                        strFacility = Trim(varCellValue)
                        If grdBlock.CellNote <> vbNullString Then udtBlock.FacilityID = grdBlock.CellNote
                    Case "START_DATE"
                         udtBlock.StartDate = Trim(varCellValue)
                     Case "END_DATE"
                         udtBlock.EndDate = Trim(varCellValue)
                     Case "START_TIME"
                         udtBlock.StartTime = Trim(varCellValue)
                     Case "END_TIME"
                        udtBlock.EndTime = Trim(varCellValue)
                    Case "REASON"
                        udtBlock.Reason = Trim(varCellValue)
                    Case "DURATION"
                        udtBlock.Duration = Trim(varCellValue)
                    Case "DAYWEEK"
                        udtBlock.DayWeek = Trim(varCellValue)
                End Select
            End If
        Next lngCol
        udtBlock.Duration = IsProperCase(udtBlock.Duration)
        If udtBlock.Duration <> vbNullString And udtBlock.Reason <> vbNullString Then
            lngBlockID = grdBlock.GetRowItemData(lngRow)
        End If
        For i = 0 To cboProviders.listcount - 1
            If cboProviders.List(i) = cboProviders.Text Then
                udtBlock.ProviderID = cboProviders.ItemData(i)
                Exit For
            End If
        Next i
        If udtBlock.FacilityID <= 0 Then udtBlock.FacilityID = LookUp_Facility(strFacility)
        If Not Validate_Save Then
            If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
            If lngBlockID = 0 Then
                strSQL = "Get_Appt_Block_ID " & udtBlock.ProviderID & ", " & udtBlock.FacilityID & ", '" & Trim(udtBlock.Duration) & "','" _
                    & Format(udtBlock.StartDate, "mm/dd/yyyy") & "', '" & Format(udtBlock.StartTime, "hh:mm AM/PM") & "'"
'                strSQL = "SELECT BLOCK_ID from APPOINTMENTS_BLOCKS where PROVIDER_ID = " & udtBlock.ProviderID & " AND " _
'                    & "FACILITY_ID = " & udtBlock.FacilityID & " AND " _
'                    & "CONVERT(DATETIME,START_DATE,101)  = '" & Format(udtBlock.StartDate, "mm/dd/yyyy") & "' AND " _
'                    & "DURATION = '" & Trim(udtBlock.Duration) & "' AND CONVERT(DATETIME,START_TIME,108)  = '" & Format(udtBlock.StartTime, "hh:mm AM/PM") & "'"
                Set rsBlock = cnBlock.Execute(strSQL)
                With rsBlock
                    If Not .EOF Then
                        If Not IsNull(!BLOCK_ID) Then
                            lngBlockID = (!BLOCK_ID)
                        End If
                    End If
                End With
            End If
            udtBlock.Reason = Replace(udtBlock.Reason, "'", "''")
            udtBlock.StartTime = Format(udtBlock.StartTime, "hh:mm AM/PM")
            sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
            If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
            If lngBlockID = 0 Then
                    strSQL = "Insert_Appt_Block '" & udtBlock.DayWeek & "','" & Trim(udtBlock.StartTime) & "','" & Trim(udtBlock.EndTime) & "','" _
                        & Trim(udtBlock.Duration) & "','" & Trim(udtBlock.Reason) & "'," & udtBlock.ProviderID & "," _
                        & udtBlock.FacilityID & ", '" & Trim(udtBlock.StartDate) & "','" & Trim(udtBlock.EndDate) & "','" & sDate & "' "
'                    strSQL = "INSERT INTO APPOINTMENTS_BLOCKS (DAYWEEK, START_TIME, END_TIME, DURATION, REASON, " _
'                        & "PROVIDER_ID,FACILITY_ID, START_DATE, END_DATE, INSERT_DATE) " _
'                        & "Values ('" & udtBlock.DayWeek & "','" & Trim(udtBlock.StartTime) & "','" & Trim(udtBlock.EndTime) & "','" _
'                        & Trim(udtBlock.Duration) & "','" & Trim(udtBlock.Reason) & "'," & udtBlock.ProviderID & "," _
'                        & udtBlock.FacilityID & ", '" & Trim(udtBlock.StartDate) & "','" & Trim(udtBlock.EndDate) & "','" & sDate & "')"
                        Set rsBlock = cnBlock.Execute(strSQL)
                        If udtBlock.StartDate <> vbNullString Then
                            Get_SchedConflict
                        End If
                    strSQL = "Get_Appt_Block_ID " & udtBlock.ProviderID & ", " & udtBlock.FacilityID & ", '" & Trim(udtBlock.Duration) & "','" _
                        & Format(udtBlock.StartDate, "mm/dd/yyyy") & "', '" & Format(udtBlock.StartTime, "hh:mm AM/PM") & "'"
'                    strSQL = "SELECT BLOCK_ID from APPOINTMENTS_BLOCKS where PROVIDER_ID = " & udtBlock.ProviderID & " AND " _
'                        & "FACILITY_ID = " & udtBlock.FacilityID & " AND " _
'                        & "CONVERT(DATETIME,START_DATE,101)  = '" & Format(udtBlock.StartDate, "mm/dd/yyyy") & "' AND " _
'                        & "DURATION = '" & Trim(udtBlock.Duration) & "' AND CONVERT(DATETIME,START_TIME,108)  = '" & Format(udtBlock.StartTime, "hh:mm AM/PM") & "'"
                    Set rsBlock = cnBlock.Execute(strSQL)
                    With rsBlock
                        If Not .EOF Then
                            If Not IsNull(!BLOCK_ID) Then
                                lngBlockID = (!BLOCK_ID)
                            End If
                        End If
                    End With
                    grdBlock.SetRowItemData lngRow, lngBlockID
            Else
                strSQL = "Update_Appt_Block '" & udtBlock.DayWeek & "','" & Trim(udtBlock.StartTime) & "','" & Trim(udtBlock.EndTime) & "','" _
                        & Trim(udtBlock.Duration) & "','" & Trim(udtBlock.Reason) & "'," & udtBlock.ProviderID & "," _
                        & udtBlock.FacilityID & ", '" & Trim(udtBlock.StartDate) & "','" & Trim(udtBlock.EndDate) & "','" & sDate & "', " & lngBlockID
'                strSQL = "UPDATE  APPOINTMENTS_BLOCKS SET START_TIME = '" & Trim(udtBlock.StartTime) & "', DAYWEEK = '" & Trim(udtBlock.DayWeek) & "'," _
'                    & "END_TIME = '" & Trim(udtBlock.EndTime) & "', ACTIVE = 1, END_DATE = '" & Trim(udtBlock.EndDate) & "'," _
'                    & "DURATION = '" & Trim(udtBlock.Duration) & " ',UPDATE_DATE = '" & sDate & "',PROVIDER_ID = " & udtBlock.ProviderID & "," _
'                    & "FACILITY_ID = " & udtBlock.FacilityID & ", START_DATE = '" & Trim(udtBlock.StartDate) & "', REASON = '" & Trim(udtBlock.Reason) & "' WHERE " _
'                    & "BLOCK_ID = " & lngBlockID
                    Set rsBlock = cnBlock.Execute(strSQL)
                    If Trim(udtBlock.DayWeek) <> vbNullString Then
                        strSQL = "Update_Appt_BlockDay_Week '" & Trim(udtBlock.DayWeek) & "', " & lngBlockID
'                        strSQL = "UPDATE  APPOINTMENTS_BLOCKS SET DAYWEEK = '" & Trim(udtBlock.DayWeek) & "' WHERE " _
'                        & "BLOCK_ID = " & lngBlockID
                        Set rsBlock = cnBlock.Execute(strSQL)
                    End If
                booExit = False
            End If
            booExit = True
        Else
            Screen.MousePointer = vbDefault
            Exit Function
        End If
  '  Next lngRow
    If Not Validate_BlockTime Then
        If Not Validate_Save Then
            If grdTerms.MaxRows >= 1 Then
                MsgBox "The New Block Time Conflicts with the following Patient Appointments," _
                    & " " & "If necessary please reschedule patients or cancel block!", vbInformation, "PatientTrac Scheduling Conflicts"
                fraSchedConflicts.Visible = True
            End If
            MsgBox "Block Times have been successfully updated!", vbInformation, "PatientTrac Block Mode"
        End If
    End If
    If cnBlock.State = adStateOpen Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    mbooData = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Save_Block_Mode:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Save_Block_Mode", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Validate_BlockTime() As Boolean
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
Dim i%
Dim Arr
Dim strEndTest As String
Dim lngBlock As Long
Dim lngPrimKey As Long

    On Error GoTo Error_Validate_BlockTime
    Screen.MousePointer = vbHourglass
   If Trim(udtBlock.StartDate) = vbNullString Then
        If Not IsNull(ccdStartDate.Value) Then
            udtBlock.StartDate = Format(ccdStartDate.Value, "mm/dd/yyyy")
        Else
            MsgBox "Unable to validate block time, please update the Start Date!", vbInformation, "PatientTrac Block Time"
            Validate_BlockTime = True
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    If udtBlock.EndDate = vbNullString Then
        If Not IsNull(ccdEndDate.Value) Then
            udtBlock.EndDate = Format(ccdEndDate.Value, "mm/dd/yyyy")
        Else
            MsgBox "Unable to validate block time, please update the End Date!", vbInformation, "PatientTrac Block Time"
            Validate_BlockTime = True
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    strRange = DateAdd("d", -7, udtBlock.StartDate)
    For i = 0 To cboProviders.listcount - 1
        If cboProviders.List(i) = cboProviders.Text Then
            udtBlock.ProviderID = cboProviders.ItemData(i)
            Exit For
        End If
    Next i
    If Trim(udtBlock.StartDate) = vbNullString Then
        Validate_BlockTime = True
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If cboRepTimeFrom.Text <> vbNullString Then udtBlock.StartTime = cboRepTimeFrom.Text
    If cboRepTimeTo.Text <> vbNullString Then udtBlock.EndTime = cboRepTimeTo.Text
    If IsDate(udtBlock.StartTime) Then
       ' strStart = DateAdd("n", -1, udtBlock.StartTime)
        strStart = udtBlock.StartTime
    Else
        MsgBox "Unable to validate block time, please update the Start Time!", vbInformation, "PatientTrac Block Time"
        Validate_BlockTime = True
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    strStart = Format(strStart, "hh:mm AM/PM")
    lngDur = Val(udtBlock.Duration) + 1
    If IsDate(udtBlock.EndTime) Then
        strEnd = Format(udtBlock.EndTime, "hh:mm AM/PM")
        strEnd = Format(strEnd, "hh:mm AM/PM")
    Else
        MsgBox "Unable to validate block time, please update the End Time!", vbInformation, "PatientTrac Block Time"
        Validate_BlockTime = True
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    lngBlock = grdBlock.GetRowItemData(mlngGridRow)
    strStartDate = Format(udtBlock.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    udtBlock.EndDate = Format(udtBlock.EndDate, "mm/dd/yyyy 00:00:00 AM/PM")
    strEndDate = DateAdd("h", 23, udtBlock.EndDate)
    strEndDate = Format(strEndDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Validate_BlockTime_1 " & udtBlock.ProviderID & ", '" & Format(strStartDate, "mm/dd/yyyy") & "','" _
        & Format(strEndDate, "mm/dd/yyyy") & "', '" & Format(strStart, "hh:mm AM/PM") & "', '" & Format(strEnd, "hh:mm AM/PM") & "'"
'     strSQL = "SELECT BLOCK_ID, START_TIME, END_TIME, START_DATE, END_DATE FROM APPOINTMENTS_BLOCKS WHERE PROVIDER_ID = " & udtBlock.ProviderID & " AND " _
'        & "ACTIVE = 1 AND " _
'        & "convert(DATETIME, START_TIME, 108) BETWEEN '" & Format(strStart, "hh:mm AM/PM") & "' AND '" & Format(strEnd, "hh:mm AM/PM") & "' AND " _
'        & "CONVERT(VARCHAR, START_DATE, 10)  > '" & Format(strStartDate, "mm/dd/yyyy") & "' AND CONVERT(VARCHAR, START_DATE, 10) < '" & Format(strEndDate, "mm/dd/yyyy") & "' And " _
'        & "CONVERT(VARCHAR, END_DATE, 10) < '" & Format(strEndDate, "mm/dd/yyyy") & "'  ORDER BY START_TIME DESC"
        'DURATION = '1 Day' AND
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
                                    Case "BLOCK_ID"
                                        lngPrimKey = Fld.Value
                                End Select
                            End If
                        End If
                    Next
                    .MoveNext
                    If DateDiff("d", udtBlock.StartDate, strStartDate) > 0 And _
                        DateDiff("d", udtBlock.EndDate, strEndDate) < 0 Then
                        If DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(strStart, "hh:mm AM/PM")) > 0 And DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(strEnd, "hh:mm AM/PM")) < 0 Then
                            If lngPrimKey <> lngBlock Then
                                Validate_BlockTime = True
                                MsgBox "Provider has this time period blocked, time is blocked from" & " " & strFoundStart & " " & "until" & " " & strFoundEnd
                                Exit Do
                            End If
                       End If
                    End If
            Loop
        End With
    If Validate_BlockTime = False Then
           strWeekDay = Weekday_Return(WeekDay(udtBlock.StartDate))
            strSQL = "Validate_BlockTime_2 " & udtBlock.ProviderID & ", " & udtBlock.FacilityID & ", '" & Format(strStartDate, "mm/dd/yyyy") & "','" _
                & Format(strEndDate, "mm/dd/yyyy") & "', '" & Format(strStart, "hh:mm AM/PM") & "', '" & Format(strEnd, "hh:mm AM/PM") & "','" _
                & strWeekDay & "'"
'            strSQL = "SELECT BLOCK_ID, START_TIME, END_TIME, START_DATE, END_DATE FROM APPOINTMENTS_BLOCKS WHERE PROVIDER_ID = " & udtBlock.ProviderID & " AND " _
'                & "FACILITY_ID = " & udtBlock.FacilityID & " AND DURATION = 'Permanent' AND convert(DATETIME, START_TIME, 108) " _
'                & "BETWEEN '" & Format(strStart, "hh:mm AM/PM") & "' AND '" & Format(strEnd, "hh:mm AM/PM") & "' AND  " _
'                & "ACTIVE = 1 AND DAYWEEK = '" & strWeekDay & "' AND CONVERT(DATETIME, START_DATE, 101)  > '" & Format(strStartDate, "mm/dd/yyyy") & "' AND CONVERT(DATETIME, START_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "' And " _
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
                                        Case "BLOCK_ID"
                                            lngPrimKey = Fld.Value
                                        Case "START_DATE"
                                            strStartDate = Trim(Fld.Value)
                                        Case "END_DATE"
                                            strEndDate = Trim(Fld.Value)
                                    End Select
                                End If
                            End If
                        Next
                        .MoveNext
                        If DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(strStart, "hh:mm AM/PM")) >= 0 And _
                            DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(strEnd, "hh:mm AM/PM")) < 0 Then
    '                        If Format(strFoundStart, "hh:mm AM/PM") <= Format(strStart, "hh:mm AM/PM") And Format(strFoundStart, "hh:mm AM/PM") <= Format(strEnd, "hh:mm AM/PM") Then
                            If lngPrimKey <> lngBlock Then
                                Validate_BlockTime = True
                                MsgBox "Provider has this time period permanently blocked, time is blocked from" & " " & strFoundStart & " " & "until" & " " & strFoundEnd
                                Exit Do
                            End If
                        End If
                    Loop
                End With
    End If
    
    strStartDate = Format(udtBlock.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    udtBlock.EndDate = Format(udtBlock.EndDate, "mm/dd/yyyy 00:00:00 AM/PM")
    strEndDate = DateAdd("h", 23, udtBlock.EndDate)
    strEndDate = Format(strEndDate, "mm/dd/yyyy hh:mm:ss AM/PM")
   ' strStart = DateAdd("n", -1, udtBlock.StartTime)
    strEndTest = DateAdd("d", 1, Format(udtBlock.EndDate, "mm/dd/yyyy 00:00:00 AM/PM"))
    strEndTest = Format(strEndTest, "mm/dd/yyyy 00:00:00 AM/PM")
  '  strEnd = DateAdd("n", 1, udtBlock.EndTime)
    If Validate_BlockTime = False And strWeekDay = vbNullString Then
        strSQL = "Validate_BlockTime_3 " & udtBlock.ProviderID & ", " & udtBlock.FacilityID & ", '" & Format(strStartDate, "mm/dd/yyyy") & "','" & Format(strEndDate, "mm/dd/yyyy") & "'"
'        strSQL = "SELECT BLOCK_ID, START_TIME, END_TIME, START_DATE, END_DATE  FROM APPOINTMENTS_BLOCKS WHERE PROVIDER_ID = AND DAYWEEK = '' AND " _
'            & "FACILITY_ID =  AND DURATION = 'Permanent' AND ACTIVE = 1 AND " _
'            & "CONVERT(DATETIME, START_DATE, 101)  > '" & Format(strStartDate, "mm/dd/yyyy") & "' AND CONVERT(DATETIME, START_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "' And " _
'            & "CONVERT(DATETIME, END_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "' AND " _
'            & "CONVERT(DATETIME, END_DATE, 101) < '" & Format(strEndTest, "mm/dd/yyyy") & "' ORDER BY START_TIME DESC"
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
                                    Case "BLOCK_ID"
                                        lngPrimKey = Fld.Value
                                    Case "START_DATE"
                                        strStartDate = Trim(Fld.Value)
                                    Case "END_DATE"
                                        strEndDate = Trim(Fld.Value)
                                End Select
                            End If
                        End If
                    Next
                    .MoveNext
                    If DateDiff("d", Format(udtBlock.StartDate, "mm/dd/yyyy"), Format(strStartDate, "mm/dd/yyyy")) >= 0 And _
                        DateDiff("d", Format(udtBlock.EndDate, "mm/dd/yyyy"), Format(strEndDate, "mm/dd/yyyy")) <= 0 Then
                          '  If DateDiff("h", strFoundStart, strFoundEnd) >= 8 Then
'                                If lngPrimKey <> lngBlock Then
'                                    Validate_BlockTime = True
'                                    MsgBox "Provider has this time period permanently blocked, time is blocked from" & " " & strFoundStart & " " & "until" & " " & strFoundEnd
'                                    Exit Do
'                                End If
                           ' Else
                                If DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(strStart, "hh:mm AM/PM")) <= 0 And _
                                    DateDiff("n", Format(strEnd, "hh:mm AM/PM"), Format(strFoundEnd, "hh:mm AM/PM")) <= 0 Then
                                    If lngPrimKey <> lngBlock Then
                                        Validate_BlockTime = True
                                        MsgBox "Provider has this time period permanently blocked, time is blocked from" & " " & strFoundStart & " " & "until" & " " & strFoundEnd
                                        Exit Do
                                    End If
                                End If
                           ' End If
                    End If
                Loop
            End With
    End If
    If Validate_BlockTime = False And strWeekDay = vbNullString Then
            strSQL = "Validate_BlockTime_4  " & udtBlock.ProviderID & ", " & udtBlock.FacilityID & ", '" _
                & Format(strStartDate, "mm/dd/yyyy") & "','" & Format(strEndDate, "mm/dd/yyyy") & "'"
'            strSQL = "SELECT BLOCK_ID, START_TIME, END_TIME FROM APPOINTMENTS_BLOCKS WHERE PROVIDER_ID = " & udtBlock.ProviderID & " AND " _
'                & "FACILITY_ID = " & udtBlock.FacilityID & " AND DURATION = 'Permanent' AND " _
'                & "CONVERT(DATETIME, START_DATE, 101)  > '" & Format(strStartDate, "mm/dd/yyyy") & "' AND CONVERT(DATETIME, START_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "' And " _
'                & "CONVERT(DATETIME, END_DATE, 101) < '" & Format(strEndDate, "mm/dd/yyyy") & "' AND ACTIVE = 1"
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
                                        Case "BLOCK_ID"
                                            lngPrimKey = Fld.Value
                                    End Select
                                End If
                            End If
                        Next
                        .MoveNext
                        If DateDiff("d", Format(udtBlock.StartDate, "mm/dd/yyyy"), Format(strStartDate, "mm/dd/yyyy")) >= 0 And _
                            DateDiff("d", Format(udtBlock.EndDate, "mm/dd/yyyy"), Format(strEndDate, "mm/dd/yyyy")) <= 0 Then
                            If DateDiff("n", Format(strFoundStart, "hh:mm AM/PM"), Format(udtBlock.StartTime, "hh:mm AM/PM")) <= 0 And _
                                    DateDiff("n", Format(strEnd, "hh:mm AM/PM"), Format(strFoundEnd, "hh:mm AM/PM")) <= 0 Then
    '                        If Format(strFoundStart, "hh:mm AM/PM") <= Format(udtBlock.StartTime, "hh:mm AM/PM") And Format(strFoundStart, "hh:mm AM/PM") <= Format(strEnd, "hh:mm AM/PM") Then
    '                            If Format(udtBlock.StartTime, "hh:mm AM/PM") > Format(strFoundStart, "hh:mm AM/PM") Then
    '                                If Format(udtBlock.StartTime, "hh:mm AM/PM") <= Format(strEnd, "hh:mm AM/PM") Then
                                        If lngPrimKey <> lngBlock Then
                                            Validate_BlockTime = True
                                            MsgBox "Provider has this time period permanently blocked, time is blocked from" & " " & strFoundStart & " " & "until" & " " & strFoundEnd
                                            Exit Do
                                        End If
                                   ' End If
                               ' End If
                            End If
                        End If
                    Loop
                End With
        End If
        If Validate_BlockTime = False And strWeekDay = vbNullString Then
            strSQL = "Validate_BlockTime_5 " & udtBlock.ProviderID & "," & udtBlock.FacilityID & ",'" & Format(strRange, "mm/dd/yyyy") & "'"
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
                                        strFoundStart = Trim(Fld.Value)
                                    Case "END_TIME"
                                        strFoundEnd = Trim(Fld.Value)
                                    Case "REASON"
                                        udtBlock.Reason = Trim(Fld.Value)
                                    Case "DURATION"
                                        udtBlock.Duration = Trim(Fld.Value)
                                    Case "FACILITY_ID"
                                        udtBlock.FacilityID = Trim(Fld.Value)
                                    Case "BLOCK_ID"
                                        lngPrimKey = Fld.Value
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
                            Validate_BlockTime = True
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
                                                Validate_BlockTime = True
                                                Exit Do
                                            Else
                                                Validate_BlockTime = False
                                                Exit Do
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate) = 0 And _
                            DateDiff("d", udtBlock.EndDate, udtAppointment.SchedDate) = 0 Then
                            If lngDateDiff = DateDiff("n", udtAppointment.SchedTime, strFoundEnd) < 0 Then
                                Validate_BlockTime = False
                                booExit = True
                                Exit Do
                            Else
                                If lngPrimKey <> lngBlock Then
                                    Validate_BlockTime = True
                                    Exit Do
                                End If
                            End If
                        End If
                    End If
                Loop
            End If
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_BlockTime:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_BlockTime", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cboDuration_Click()
    If cboDuration.Text = "Permanent" Then
        cctDuration.Text = vbNullString
        cctDuration.Enabled = False
    Else
        cctDuration.Enabled = True
    End If
    SetEndDate cboDuration.Text
    If cboDuration.Text <> "1 Day" Then
       cboWeek.Enabled = True
    End If
End Sub
Private Sub cboProviders_Click()
Dim i%
Dim lngProvID As Long

    Format_Grid
    For i = 0 To cboProviders.listcount - 1
        If cboProviders.List(i) = cboProviders.Text Then
            udtBlock.ProviderID = cboProviders.ItemData(i)
            Exit For
        End If
    Next i
    If udtBlock.ProviderID > 0 Then Read_Blocks_ByProvider udtBlock.ProviderID
    Exit Sub
End Sub


Private Sub cboReasonTerms_Click()

    On Error GoTo Error_ReasonTerms
    frmNurseTemplates.fCaption = "REASON"
    frmNurseTemplates.fType = "REASON"
    frmNurseTemplates.fSubDomain = "sd_Appointment_BlockReason"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    Screen.MousePointer = vbDefault
    Exit Sub
Error_ReasonTerms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "ReasonTerms", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cboWeek_Click()
Dim strWeek As String
Dim strDate As String
Dim strRepDate As String
Dim FirstDay As VbDayOfWeek
Dim FirstDate As Date


   If cboWeek.Text = vbNullString Then
        Exit Sub
   End If
    strWeek = cboWeek.Text
    If Not IsNull(ccdStartDate.Value) Then
        strRepDate = Format(ccdStartDate.Value, "mm/dd/yyyy")
    Else
        MsgBox "Please select Start Date for this block time. First" & " " & strWeek & " " & "of desired block start time.", vbInformation, "PatientTrac Financial Management"
        cboWeek.ListIndex = -1
        Exit Sub
    End If
    Select Case strWeek
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
    If strRepDate <> vbNullString Then
        ccdStartDate.Value = Format(strRepDate, "mm/dd/yyyy")
    End If
    Exit Sub
End Sub
Private Function Day_Week(vstrDate As String, vstrDayWeek As String) As String
Dim strDayWeek As String
Dim FirstDay As VbDayOfWeek

    
    On Error GoTo Err_Day_Week
    strDayWeek = vstrDate
    Select Case vstrDayWeek
        Case "Monday"
            FirstDay = WeekDay(vstrDate, vbMonday)
            strDayWeek = DateAdd("d", -(FirstDay - 1), strDayWeek)
        Case "Tuesday"
            FirstDay = WeekDay(vstrDate, vbTuesday)
            strDayWeek = DateAdd("d", -(FirstDay - 1), strDayWeek)
        Case "Wednesday"
            FirstDay = WeekDay(vstrDate, vbWednesday)
            strDayWeek = DateAdd("d", -(FirstDay - 1), strDayWeek)
        Case "Thursday"
            FirstDay = WeekDay(vstrDate, vbThursday)
            strDayWeek = DateAdd("d", -(FirstDay - 1), strDayWeek)
        Case "Friday"
            FirstDay = WeekDay(vstrDate, vbFriday)
            strDayWeek = DateAdd("d", -(FirstDay - 1), strDayWeek)
        Case "Saturday"
            FirstDay = WeekDay(vstrDate, vbSaturday)
            strDayWeek = DateAdd("d", -(FirstDay - 1), strDayWeek)
        Case "Sunday"
            FirstDay = WeekDay(vstrDate, vbSunday)
            strDayWeek = DateAdd("d", -(FirstDay - 1), strDayWeek)
    End Select
    Day_Week = strDayWeek
    Exit Function
Err_Day_Week:
    Exit Function
    Resume Next
End Function
Private Function GetDay(ipDate As String) As String
    Dim lDay As Long

    lDay = WeekDay(ipDate)
    Select Case lDay
       Case 1
          GetDay = "SUNDAY"
       Case 2
          GetDay = "MONDAY"
       Case 3
          GetDay = "TUESDAY"
       Case 4
          GetDay = "WEDNESDAY"
       Case 5
          GetDay = "THURSDAY"
       Case 6
          GetDay = "FRIDAY"
       Case 7
          GetDay = "SATURDAY"
    End Select

End Function

Private Sub ccdStartDate_Change()
    SetEndDate cboDuration.Text
End Sub
Private Sub ccdStartDate_Click()
    SetEndDate cboDuration.Text
End Sub

Private Sub cctDuration_Click()
    If cboDuration.Text = "Permanent" Then
        cctDuration.Text = vbNullString
        cctDuration.Enabled = False
    Else
        cctDuration.Enabled = True
    End If
    SetEndDate cboDuration.Text
    If cboDuration.Text <> "1 Day" Then
       cboWeek.Enabled = True
    End If
End Sub


Private Sub cmdAccept_Click()
Dim strDate As String
Dim varCellValue
Dim lngRow As Long
Dim lngCol As Long
Dim Arr

    fraNumber.Visible = False
    mbooData = True
    clsGridFormat.CellTag = vbNullString
    With grdBlock
        .Col = mlngGridCol
        .Row = mlngGridRow
        If .CellTag <> vbNullString Then clsGridFormat.CellTag = .CellTag
        If clsGridFormat.CellTag <> vbNullString Then
            Select Case clsGridFormat.CellTag
                  Case "START_TIME", "END_TIME"
                        Arr = Split(LookUpCellTag("START_TIME"), "|")
                        If UBound(Arr) > 0 Then
                          lngCol = Arr(0)
                        End If
                        Format_CellTypeEdit grdBlock, mlngGridRow, lngCol, "1", cboRepTimeFrom.Text, "TypeHAlignCenter", "START_TIME"
                        grdBlock.SetCellBorder lngCol, mlngGridRow, lngCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Arr = Split(LookUpCellTag("END_TIME"), "|")
                        If UBound(Arr) > 0 Then
                          lngCol = Arr(0)
                        End If
                        Format_CellTypeEdit grdBlock, mlngGridRow, lngCol, "1", cboRepTimeTo.Text, "TypeHAlignCenter", "END_TIME"
                        grdBlock.SetCellBorder lngCol, mlngGridRow, lngCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
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
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    fraDate.Visible = False
End Sub

Private Sub cmdClose_Click()
    mlngGridRow = 0
    Unload Me
End Sub

Private Sub cmdCloseSubmit_Click()
    mbooData = False
    mlngGridRow = 0
    fraSubmitDate.Visible = False
End Sub
Private Sub cmdDurationBlock_Click()

    On Error GoTo Error_DurationTerms
    frmNurseTemplates.fCaption = "DURATION"
    frmNurseTemplates.fType = "DURATION"
    frmNurseTemplates.fSubDomain = "sd_Appointment_BlockDuration"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    Screen.MousePointer = vbDefault
    Exit Sub
Error_DurationTerms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "DurationTerms", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Sub cmdExit_Click()
    fraSchedConflicts.Visible = False
    Exit Sub
End Sub

Private Sub cmdGridRow_Click()
Dim iAnswer As Integer

    If mlngGridRow > 0 And mbooData Then
        iAnswer = MsgBox("Do you want to save your previous changes? Please Save Changes before updating or adding a New Block Time", vbYesNo, "PatientTrac Block Times")
        If iAnswer = vbYes Then
            Save_Block_Mode
            Exit Sub
        ElseIf iAnswer = vbNo Then
            With grdBlock
                .BlockMode = True
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = mlngGridRow
                    .Row2 = mlngGridRow
                    .BackColor = NormalColor
                .BlockMode = False
            End With
            mlngGridRow = 0
            mbooData = False
            Read_Blocks_ByProvider udtBlock.ProviderID
        End If
    End If
    fraNumber.Visible = False
    fraDate.Visible = False
    If cboProviders.Text = vbNullString Then
        MsgBox "Please select a provider first!", vbInformation, "PatientTrac Block Time Management"
    Else
        Add_Row
        If grdBlock.MaxRows > 0 Then mlngGridRow = grdBlock.MaxRows
    End If
    Exit Sub
End Sub
Private Sub cmdPrintList_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(1) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Scheduling Conflicts"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    With grdTerms
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Block Time Appointments"
        .PrintFooter = "/cPage #/p/n"
        .PrintColor = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintSmartPrint = True
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Phone Messages"
        .PrintHeader = strHeader(1) & vbNewLine & strHeader(0)
        .PrintFooter = "/cPage #/p/n" & strHeader(0)
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
        .PrintOrientation = PrintOrientationLandscape
        .PrintShadows = False
         .PrintSheet
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdPrint_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Print", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdRemoveRow_Click()
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim varCellValue
Dim lngPrimKey As Long

    On Error GoTo Error_DeleteRow
    Screen.MousePointer = vbHourglass
    If mlngGridRow = 0 Then
        MsgBox "Please select the desired removal row first!", vbInformation, "PatientTrac Appointment Block Time"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    With grdBlock
        .GetText 1, mlngGridRow, varCellValue
        If Trim(varCellValue) = vbNullString Then
            .DeleteRows mlngGridRow, 1
            .MaxRows = .MaxRows - 1
        Else
            If mlngLogonUser > 0 Then
                Get_LogonProvider
            End If
           fraRemove.Visible = True
        End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_DeleteRow:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "DeleteRow", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdRemReason_Click()
    fraRemove.Visible = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdRemReason:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdRemReason", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Sub cmdSave_Click()
Dim varCellValue
Dim lngRow As Long
Dim lngCol As Long
Dim strEndState As String
Dim strDate As String
Dim strDuration As String
Dim Arr
Dim strStart As String
Dim iAnswer As Integer
Dim lngDateDiff As Long
Dim Idx%

'    If Trim(cboWeek.Text) = vbNullString And cboDuration.Text <> "1 Day" And cboDuration.Text <> "Permanent" Then
'       MsgBox "Please Select Day Of Week", vbInformation, "PatientTrac Block Mode"
'       Exit Sub
'    End If
    If cboDuration.Text <> "Permanent" Then SetEndDate cboDuration.Text
    If cboDuration.Text = vbNullString And IsDate(ccdStartDate.Value) Then
        If Not IsDate(ccdEndDate.Value) Then
            MsgBox "Please select a valid end date or select duration!", vbInformation, "PatientTrac Block Time"
            Exit Sub
        Else
            lngDateDiff = DateDiff("d", Format(ccdStartDate.Value, "mm/dd/yyyy"), Format(ccdEndDate.Value, "mm/dd/yyyy"))
            If lngDateDiff > 1 Then
                If lngDateDiff < 7 Then
                    cctDuration.Text = CStr(lngDateDiff)
                    For Idx = 0 To cboDuration.listcount - 1
                        If IsProperCase(cboDuration.List(Idx)) = "Days" Then
                            cboDuration.Text = cboDuration.List(Idx)
                        End If
                    Next Idx
                ElseIf lngDateDiff = 28 Then
                    cctDuration.Text = "4"
                    For Idx = 0 To cboDuration.listcount - 1
                        If IsProperCase(cboDuration.List(Idx)) = "Weeks" Then
                            cboDuration.Text = cboDuration.List(Idx)
                        End If
                    Next Idx
                ElseIf lngDateDiff = 14 Then
                    cctDuration.Text = "2"
                    For Idx = 0 To cboDuration.listcount - 1
                        If IsProperCase(cboDuration.List(Idx)) = "Weeks" Then
                            cboDuration.Text = cboDuration.List(Idx)
                        End If
                    Next Idx
                ElseIf lngDateDiff = 21 Then
                    cctDuration.Text = "3"
                    For Idx = 0 To cboDuration.listcount - 1
                        If IsProperCase(cboDuration.List(Idx)) = "Weeks" Then
                            cboDuration.Text = cboDuration.List(Idx)
                        End If
                    Next Idx
                ElseIf lngDateDiff > 7 And lngDateDiff < 30 Then
                    cctDuration.Text = lngDateDiff
                    For Idx = 0 To cboDuration.listcount - 1
                        If IsProperCase(cboDuration.List(Idx)) = "Days" Then
                            cboDuration.Text = cboDuration.List(Idx)
                        End If
                    Next Idx
                ElseIf lngDateDiff = 7 Then
                    cctDuration.Text = "1"
                    For Idx = 0 To cboDuration.listcount - 1
                        If IsProperCase(cboDuration.List(Idx)) = "Week" Then
                            cboDuration.Text = cboDuration.List(Idx)
                        End If
                    Next Idx
                ElseIf lngDateDiff > 7 And lngDateDiff < 31 Then
                    cctDuration.Text = "1"
                    For Idx = 0 To cboDuration.listcount - 1
                        If IsProperCase(cboDuration.List(Idx)) = "Month" Then
                            cboDuration.Text = cboDuration.List(Idx)
                        End If
                    Next Idx
                ElseIf lngDateDiff > 30 And lngDateDiff > 61 Then
                    lngDateDiff = DateDiff("m", Format(ccdStartDate.Value, "mm/dd/yyyy"), Format(ccdEndDate.Value, "mm/dd/yyyy"))
                    cctDuration.Text = CStr(lngDateDiff)
                    For Idx = 0 To cboDuration.listcount - 1
                        If IsProperCase(cboDuration.List(Idx)) = "Months" Then
                            cboDuration.Text = cboDuration.List(Idx)
                        End If
                    Next Idx
                End If
            Else
                cctDuration.Text = "1"
                For Idx = 0 To cboDuration.listcount - 1
                    If IsProperCase(cboDuration.List(Idx)) = "Day" Then
                        cboDuration.Text = cboDuration.List(Idx)
                    End If
                Next Idx
            End If
        End If
    ElseIf Not IsDate(ccdStartDate.Value) Then
        MsgBox "Please select a valid start date!", vbInformation, "PatientTrac Block Time"
        Exit Sub
    End If
'    'fix end data
'    Select Case cboDuration.Text
'        If cctDuration = vbNullString Then cctDuration = "1"
'        Case "Week", "Weeks"
'            udtBlock.EndDate = DateAdd("ww", Val(cctDuration), udtBlock.StartDate)
'        Case "Month", "Months"
'            udtBlock.EndDate = DateAdd("m", Val(cctDuration), udtBlock.StartDate)
'        Case "Year", "Years"
'        udtBlock.EndDate = DateAdd("y", Val(cctDuration), udtBlock.StartDate)
'    End Select
    mbooData = True
    udtBlock.StartDate = Format(ccdStartDate.Value, "mm/dd/yyyy")
    udtBlock.EndDate = Format(ccdEndDate.Value, "mm/dd/yyyy")
    Arr = Split(LookUpCellTag("START_DATE"), "|")
    If UBound(Arr) > 0 Then
      lngCol = Arr(0)
    End If
    Format_CellTypeEdit grdBlock, mlngGridRow, lngCol, "1", udtBlock.StartDate, "TypeHAlignCenter", "START_DATE"
    grdBlock.SetCellBorder mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Arr = Split(LookUpCellTag("END_DATE"), "|")
    If UBound(Arr) > 0 Then
      lngCol = Arr(0)
    End If
    Format_CellTypeEdit grdBlock, mlngGridRow, lngCol, "1", udtBlock.EndDate, "TypeHAlignCenter", "END_DATE"
    grdBlock.SetCellBorder mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid

    udtBlock.Duration = cctDuration.Text & " " & cboDuration.Text
    udtBlock.DayWeek = cboWeek.Text
    If cboWeek.Text <> vbNullString Then
        If ccdStartDate.CheckBox Then
            strStart = ccdStartDate.Value
            If cboWeek.Text <> vbNullString Then strStart = Day_Week(strStart, cboWeek.Text)
            If DateDiff("d", Format(ccdStartDate.Value, "mm/dd/yyyy"), Format(strStart, "mm/dd/yyyy")) <> 0 Then
                MsgBox "Invalid Start Date Selection, Please select" & " " & "first" & " " & cboWeek.Text & " " & "of the desired start week!", vbInformation, "PatientTrac Block Times"
                Exit Sub
            End If
        Else
            MsgBox "Please select a valid start date first!", vbInformation, "PatientTrac Block Times"
            Exit Sub
        End If
    End If
    If udtBlock.StartDate <> vbNullString And udtBlock.EndDate <> vbNullString Then
       If Validate_BlockTime Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    fraDate.Visible = False
    clsGridFormat.CellTag = vbNullString
    With grdBlock
        .Col = mlngGridCol
        .Row = mlngGridRow
        If .CellTag <> vbNullString Then clsGridFormat.CellTag = .CellTag
        If clsGridFormat.CellTag <> vbNullString Then
            Select Case clsGridFormat.CellTag
                Case "START_DATE", "END_DATE", "DURATION", "DAYWEEK"
                    If Not IsNull(ccdStartDate.Value) Then
                        strDate = Format(ccdStartDate.Value, "mm/dd/yyyy")
                        Arr = Split(LookUpCellTag("START_DATE"), "|")
                        If UBound(Arr) > 0 Then
                          lngCol = Arr(0)
                        End If
                        Format_CellTypeEdit grdBlock, mlngGridRow, lngCol, "1", strDate, "TypeHAlignCenter", "START_DATE"
                        grdBlock.SetCellBorder mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                    If Not IsNull(ccdEndDate.Value) Then
                        strDate = Format(ccdEndDate.Value, "mm/dd/yyyy")
                        Arr = Split(LookUpCellTag("END_DATE"), "|")
                        If UBound(Arr) > 0 Then
                          lngCol = Arr(0)
                        End If
                        Format_CellTypeEdit grdBlock, mlngGridRow, lngCol, "1", strDate, "TypeHAlignCenter", "END_DATE"
                        grdBlock.SetCellBorder mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                    If Trim(cboDuration.Text) <> vbNullString Then
                        Arr = Split(LookUpCellTag("DURATION"), "|")
                        If UBound(Arr) > 0 Then
                          lngCol = Arr(0)
                        End If
                        Format_CellTypeStatic grdBlock, mlngGridRow, lngCol, "1", Trim(udtBlock.Duration), "TypeHAlignCenter", "DURATION", "TypeHAlignCenter", 0, ""
                        grdBlock.SetCellBorder mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                    Arr = Split(LookUpCellTag("DAYWEEK"), "|")
                    If UBound(Arr) > 0 Then
                      lngCol = Arr(0)
                    End If
                    Format_CellTypeStatic grdBlock, mlngGridRow, lngCol, "1", Trim(cboWeek.Text), "TypeHAlignCenter", "DAYWEEK", "TypeHAlignCenter", 0, ""
                    grdBlock.SetCellBorder mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
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
    cctDuration.Text = vbNullString
    cboWeek.ListIndex = -1
    Exit Sub
End Sub

Private Sub cmdSaveBlock_Click()
 
    On Error GoTo Error_Set_BlockTime
    Screen.MousePointer = vbHourglass
    If mbooData Then Save_Block_Mode
    mlngGridRow = 0
    mbooData = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Set_BlockTime:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Set_BlockTime", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Function Get_SchedConflict()
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim nDays As String
Dim strDuration As String
Dim strEndTime As String
Dim strStart As String
Dim strStartDate As String
Dim strEndDate As String
Dim lngPatID As Long
Dim lngApptID As Long


    On Error GoTo Error_Get_SchedConflict
    Screen.MousePointer = vbHourglass
    strStart = DateAdd("n", -1, udtBlock.StartTime)
    strStartDate = DateAdd("n", 1, udtBlock.EndTime)
    strEndTime = DateAdd("n", 1, udtBlock.EndTime)
    If Trim(udtBlock.StartDate) <> vbNullString Then
        strStartDate = Format(udtBlock.StartDate, "mm/dd/yyyy")
        strStartDate = Format(udtBlock.StartDate, "mm/dd/yyyy 00:00:00 AM/PM")
    End If
    If Trim(udtBlock.EndDate) <> vbNullString Then
            strEndDate = Format(udtBlock.EndDate, "mm/dd/yyyy")
            strEndDate = DateAdd("h", 23, strEndDate)
            strEndDate = Format(strEndDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    Else
            strEndDate = Format(strStartDate, "mm/dd/yyyy")
            strEndDate = DateAdd("h", 23, strEndDate)
            strEndDate = Format(strEndDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    End If
    Format_Conflict
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(0)
    If DateDiff("d", udtBlock.StartDate, udtBlock.EndDate) > 0 Then
        strEndDate = DateAdd("d", 1, Format(strEndDate, "mm/dd/yyyy"))
'        strSQL = "Block_Check_Provider_Conflict_Range  " & udtBlock.ProviderID & ",  " & udtBlock.FacilityID & ",'" _
'            & Format(strStartDate, "mm/dd/yyyy") & "','" & Format(strEndDate, "mm/dd/yyyy") & "', '" & Format(strStart, "hh:mm AM/PM") & "', '" _
'            & Format(strEndTime, "hh:mm AM/PM") & "'"
        strSQL = "SELECT APPOINTMENTS.*, PATIENT.LAST_NAME FROM APPOINTMENTS INNER JOIN PATIENT ON " _
            & "APPOINTMENTS.PATIENT_ID = PATIENT.PATIENT_ID " _
            & "WHERE CONVERT(DATETIME, APPOINTMENTS.APPT_DATE, 101) BETWEEN '" & Format(strStartDate, "mm/dd/yyyy") & "' " _
            & "AND  '" & Format(strEndDate, "mm/dd/yyyy") & "'  AND " _
            & "APPOINTMENTS.PROVIDER_ID = " & udtBlock.ProviderID & " AND APPOINTMENTS.FACILITY_ID = " & udtBlock.FacilityID & " AND " _
            & "APPOINTMENTS.ACTIVE = 1 AND  CONVERT(DATETIME, APPOINTMENTS.APPT_TIME, 108) BETWEEN '" & Format(strStart, "hh:mm AM/PM") & "' " _
            & "AND '" & Format(strEndTime, "hh:mm AM/PM") & "' ORDER BY CONVERT(DATETIME, APPOINTMENTS.APPT_DATE, 101) ASC, PATIENT.LAST_NAME"
        Set rsTerms = cnTerms.Execute(strSQL)
    Else
'        strSQL = "Block_Check_Provider_Conflict_Date  " & udtBlock.ProviderID & ",  " & udtBlock.FacilityID & ",'" _
'            & Format(strStartDate, "mm/dd/yyyy") & "', '" & Format(strStart, "hh:mm AM/PM") & "', '" _
'            & Format(strEndTime, "hh:mm AM/PM") & "'"
        strSQL = "SELECT APPOINTMENTS.*, PATIENT.LAST_NAME  FROM APPOINTMENTS INNER JOIN PATIENT ON " _
            & "APPOINTMENTS.PATIENT_ID = Patient.PATIENT_ID " _
            & "WHERE CONVERT(DATETIME, APPOINTMENTS.APPT_DATE, 101) = '" & Format(strStartDate, "mm/dd/yyyy") & "'  " _
            & "AND APPOINTMENTS.PROVIDER_ID = " & udtBlock.ProviderID & " AND APPOINTMENTS.FACILITY_ID = " & udtBlock.FacilityID & " AND " _
            & "APPOINTMENTS.ACTIVE = 1 AND  CONVERT(DATETIME,APPOINTMENTS.APPT_TIME, 108) " _
            & "BETWEEN '" & Format(strStart, "hh:mm AM/PM") & "' AND '" & Format(strEndTime, "hh:mm AM/PM") & "' ORDER " _
            & "BY CONVERT(DATETIME, APPOINTMENTS.APPT_DATE, 101) ASC, PATIENT.LAST_NAME"
        Set rsTerms = cnTerms.Execute(strSQL)
    End If
    Set rsTerms = cnTerms.Execute(strSQL)
    Do Until rsTerms.EOF
        lngApptID = 0
        udtSchedPatient.IntId = vbNullString
        Set Flds = rsTerms.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PATIENT_ID"
                                udtSchedPatient.IntId = Fld.Value
                            Case "APPT_TYPE"
                                udtAppointment.SchedType = Trim(Fld.Value)
                            Case "STATUS"
                                udtSchedPatient.Status = Trim(Fld.Value)
                            Case "PROVIDER_ID"
                                udtAppointment.ProviderID = Fld.Value
                            Case "FACILITY_ID"
                                udtAppointment.FacilityID = Fld.Value
                            Case "APPT_TIME"
                                udtAppointment.SchedTime = Fld.Value
                            Case "DURATION"
                                udtSchedPatient.Duration = Fld.Value
                            Case "APPT_DATE"
                                udtAppointment.SchedDate = Fld.Value
                            Case "APPT_ID"
                                lngApptID = Fld.Value
                        End Select
                    End If
                End If
            Next
            rsTerms.MoveNext
            If Not Valid_RepeatAppt(lngApptID) Then
                If IsProperCase(GetDay(udtAppointment.SchedDate)) = IsProperCase(udtBlock.DayWeek) Then
                    If lngPatID <> Val(udtSchedPatient.IntId) Then
                        Select Case IsProperCase(udtAppointment.SchedType)
                            Case "Patient"
                                If Trim(udtSchedPatient.IntId) <> vbNullString Then
                                    Get_SchedPatient Val(udtSchedPatient.IntId)
                                End If
                            Case "Vendor", "Business"
                                Get_SchedVendor Val(udtSchedPatient.IntId)
                            Case "Legal", "Attorney"
                                Get_Schedule_Attorney Val(udtSchedPatient.IntId)
                            Case "Physician"
                                Get_SchedPhysician Val(udtSchedPatient.IntId)
                        End Select
                        If Val(Replace(udtSchedPatient.Duration, "minutes", "")) > 0 Then
                          strEndTime = DateAdd("n", Val(Replace(udtSchedPatient.Duration, "minutes", "")), udtAppointment.SchedTime)
                          strEndTime = Format(strEndTime, "hh:mm AM/PM")
                        End If
                        
                        If Trim(udtSchedPatient.Collection) <> vbNullString Then
                            If IsDate(udtAppointment.SchedTime) And IsDate(udtBlock.StartTime) Then
                                lngDateDiff = DateDiff("n", Format(udtBlock.StartTime, "hh:mm AM/PM"), Format(udtAppointment.SchedTime, "hh:mm AM/PM"))
                                If lngDateDiff > 0 Then
                                    lngDateDiff = DateDiff("n", Format(udtBlock.StartTime, "hh:mm AM/PM"), Format(udtBlock.EndTime, "hh:mm AM/PM"))
                                    If lngDateDiff > 0 Then
                                        If DateDiff("n", Format(udtBlock.EndTime, "hh:mm AM/PM"), Format(strEndTime, "hh:mm AM/PM")) < 0 Then
                                            Add_ConflictRow
                                        Else
                                            lngDateDiff = DateDiff("n", Format(udtAppointment.SchedTime, "hh:mm AM/PM"), Format(udtBlock.EndTime, "hh:mm AM/PM"))
                                            If lngDateDiff > 0 Then
                                                Add_ConflictRow
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    lngPatID = Val(udtSchedPatient.IntId)
                End If
            End If
        Loop
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_SchedConflict:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_SchedConflict", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cmdSelect_Click()

    
    fraSubmitDate.Visible = False
    fraSubmitDate.ZOrder 1
    mbooData = False
    mlngGridRow = 0
    Clear_Time
    If Not IsNull(ccdSubFrom.Value) Then
        udtTime.StartDate = Format(ccdSubFrom.Value, "mm/dd/yyyy 00:00:00 AM/PM")
    End If
    If Not IsNull(ccdSubTo.Value) Then
        udtTime.EndDate = Format(ccdSubTo.Value, "mm/dd/yyyy 00:00:00 AM/PM")
        udtTime.EndDate = DateAdd("h", 23, udtTime.EndDate)
    Else
        udtTime.EndDate = DateAdd("h", 23, udtTime.StartDate)
    End If
    If DateDiff("d", udtTime.EndDate, udtTime.StartDate) > 0 Then
        MsgBox "End Date is before the Start Date, please select an appropriate Date Range!", vbCritical, "PatientTrac Block Time"
        Exit Sub
    End If
    If udtTime.StartDate <> vbNullString And udtTime.EndDate <> vbNullString Then
        Read_Blocks_ByProvider_Date udtBlock.ProviderID
    End If
    Exit Sub
End Sub
Private Sub cmdSelectRemove_Click()
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim varCellValue
Dim lngPrimKey As Long
Dim strUser As String
Dim strReason As String
Dim lngProv As Long
Dim Idx%


    On Error GoTo Error_cmdRemReason
    Screen.MousePointer = vbHourglass
    fraNumber.Visible = False
    fraDate.Visible = False
    For Idx = 0 To cboProviders.listcount - 1
        If cboProviders.List(Idx) = cboProviders.Text Then
            lngProv = cboProviders.ItemData(Idx)
            Exit For
        End If
    Next Idx
    If mlngLogonUser > 0 Then
        Get_LogonProvider
    End If
    strUser = cboStaff.Text
    strReason = cboReason.Text
    If Trim(strUser) = vbNullString Or Trim(strReason) = vbNullString Then
        MsgBox "User Name and Reason is Mandatory to Remove Block Times!", vbCritical, "PatientTrac Block Times"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lngPrimKey = grdBlock.GetRowItemData(mlngGridRow)
    If lngPrimKey > 0 Then
        If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
        strSQL = "UPDATE APPOINTMENTS_BLOCKS SET ACTIVE = 0, CANCEL_REASON = '" & strReason & "'," _
            & "CANCEL_USER = '" & strUser & "', UPDATE_DATE = '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "' WHERE BLOCK_ID = " & lngPrimKey
        Set rsBlock = cnBlock.Execute(strSQL)
        If cnBlock.State = adStateOpen Then
            cnBlock.Close
            Set cnBlock = Nothing
            Set rsBlock = Nothing
            Format_Grid
            Read_Blocks_ByProvider lngProv
        End If
    End If
    MsgBox "Block Time has been Successfully Removed!", vbInformation, "PatientTrac Block Times"
'    If mlngGridRow > 0 Then
'        With grdBlock
'            .cmdRemReasons mlngGridRow, 1
'            .MaxRows = .MaxRows - 1
'        End With
'    End If
    fraRemove.Visible = False
    mlngGridRow = 0
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdRemReason:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "cmdRemReason", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdShowAll_Submit_Click()
    mbooData = False
    fraSubmitDate.Visible = False
    mlngGridRow = 0
    If udtBlock.ProviderID > 0 Then Read_Blocks_ByProvider udtBlock.ProviderID
End Sub

Private Sub Form_Activate()
    mbooData = False
    Exit Sub
End Sub

Private Sub Form_Load()
Dim sValue As String
Dim Idx As Integer
'    mFirst = True
     
     centerForm Me
     mbooData = False
     Clear_Block
     Get_Appt_Providers
     Format_Grid
     ReDim mStartHour(11)
     mStartHour(0) = "12:00"
     mStartHour(1) = "01:00"
     mStartHour(2) = "02:00"
     mStartHour(3) = "03:00"
     mStartHour(4) = "04:00"
     mStartHour(5) = "05:00"
     mStartHour(6) = "06:00"
     mStartHour(7) = "07:00"
     mStartHour(8) = "08:00"
     mStartHour(9) = "09:00"
     mStartHour(10) = "10:00"
     mStartHour(11) = "11:00"
     mStartIdx = 8
        
     ReDim mEndHour(11)
     mEndHour(0) = "12:00"
     mEndHour(1) = "01:00"
     mEndHour(2) = "02:00"
     mEndHour(3) = "03:00"
     mEndHour(4) = "04:00"
     mEndHour(5) = "05:00"
     mEndHour(6) = "06:00"
     mEndHour(7) = "07:00"
     mEndHour(8) = "08:00"
     mEndHour(9) = "09:00"
     mEndHour(10) = "10:00"
     mEndHour(11) = "11:00"
     mEndIdx = 8
     Load_Providers
     Load_ComboTime
     sValue = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScheduleStartHour")
     If sValue <> vbNullString Then
        For Idx = 0 To UBound(mStartHour)
            If mStartHour(Idx) = sValue Then
               mStartIdx = Idx
               Exit For
            End If
        Next Idx
        sValue = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ScheduleEndHour")
        For Idx = 0 To UBound(mEndHour)
            If mEndHour(Idx) = sValue Then
               mEndIdx = Idx
               Exit For
            End If
        Next Idx
    End If
    
'    SetWindowOnTop Me, True

    Exit Sub
End Sub

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
Dim Idx%

    On Error GoTo Error_Get_Appt_Providers
    Screen.MousePointer = vbHourglass
     clsGridFormat.List = vbNullString
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    strSQL = "Get_Appt_Providers " & udtFacility.FacID
     Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        Do Until .EOF
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            Clear_Provider
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
                clsProvider.FirstName = Trim(strFirst)
                clsProvider.FirstName = Fix_Name(clsProvider.FirstName)
                clsProvider.Name = clsProvider.FirstName
            End If
            If Trim(strMiddle) <> vbNullString Then
               strMiddle = IsProperCase(strMiddle)
                clsProvider.Name = clsProvider.Name & " " & Left(strMiddle, 1)
            End If
            If Trim(strLast) <> vbNullString Then
                clsProvider.LastName = Trim(strLast)
                clsProvider.LastName = Fix_Name(clsProvider.LastName)
                clsProvider.Name = clsProvider.Name & " " & clsProvider.LastName
            End If
            If Trim(strSuffix) <> vbNullString Then
                strSuffix = Fix_Suffix(strSuffix)
                clsProvider.Name = clsProvider.Name & ", " & Trim(strSuffix)
            Else
                clsProvider.Name = clsProvider.Name
            End If
            If Trim(strTitle) <> vbNullString Then
                strTitle = Fix_Title(strTitle)
                clsProvider.Name = clsProvider.Name & ", " & Trim(strTitle)
                clsProvider.Title = Trim(strTitle)
            End If
            If clsProvider.Name <> vbNullString Then
                cboProviders.AddItem "clsProvider.Name"
                cboProviders.ItemData(cboProviders.NewIndex) = lngProvID
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
    Resume Next
End Function
Private Function Set_Appt_Providers(vlngProvId As Long, Optional vlngRow As Long, Optional vlngCol As Long, Optional vstrGrid As String)
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

    On Error GoTo Error_Set_Appt_Providers
    Screen.MousePointer = vbHourglass
    lngProvID = vlngProvId
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    strSQL = "Get_Appt_Provider_Provider_ID " & lngProvID
'    strSQL = "SELECT b.*, a.PROVIDER_ID from APPOINTMENT_PROVIDERS a, PROVIDERS b " _
'        & "where a.PROVIDER_ID = " & lngProvID & " and b.PROVIDER_ID = a.PROVIDER_ID AND b.ACTIVE = 1 ORDER BY b.LAST_NAME, b.FIRST_NAME"
     Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        If Not .EOF Then
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            Clear_Provider
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
            If Trim(strFirst) <> vbNullString Then
                clsProvider.FirstName = Trim(strFirst)
                clsProvider.FirstName = Fix_Name(clsProvider.FirstName)
                clsProvider.Name = clsProvider.FirstName
            End If
            If Trim(strMiddle) <> vbNullString Then
               strMiddle = IsProperCase(strMiddle)
                clsProvider.Name = clsProvider.Name & " " & Left(strMiddle, 1)
            End If
            If Trim(strLast) <> vbNullString Then
                clsProvider.LastName = Trim(strLast)
                clsProvider.LastName = Fix_Name(clsProvider.LastName)
                clsProvider.Name = clsProvider.Name & " " & clsProvider.LastName
            End If
            If Trim(strSuffix) <> vbNullString Then
                strSuffix = Fix_Suffix(strSuffix)
                clsProvider.Name = clsProvider.Name & ", " & Trim(strSuffix)
            Else
                clsProvider.Name = clsProvider.Name
            End If
            If Trim(strTitle) <> vbNullString Then
                strTitle = Fix_Title(strTitle)
                clsProvider.Name = clsProvider.Name & ", " & Trim(strTitle)
                clsProvider.Title = Trim(strTitle)
            End If
            If clsProvider.Name <> vbNullString Then
                If Trim(vstrGrid) <> vbNullString Then
                   ' Set_GridCurSel grdTerms, clsProvider.Name, vlngCol, vlngRow, "1", "PROVIDER_ID", 0, CStr(lngProvID)
                    Format_CellTypeStatic grdTerms, vlngRow, vlngCol, "1", clsProvider.Name, "TypeHAlignLeft", "PROVIDER_ID", "TypeHAlignLeft", 0, CStr(lngProvID)
                Else
                    Format_CellTypeStatic grdBlock, vlngRow, vlngCol, "1", clsProvider.Name, "TypeHAlignLeft", "PROVIDER_ID", "TypeHAlignLeft", 0, CStr(lngProvID)
                End If
            End If
        End If
    End With
    If cnMedProv.State = adStateOpen Then
        cnMedProv.Close
        Set cnMedProv = Nothing
        Set rsMedProv = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_Appt_Providers:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Set_Appt_Providers", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub Form_Unload(Cancel As Integer)
     On Error Resume Next
    grdTerms.ClearRange 1, 1, grdTerms.MaxCols, grdTerms.MaxRows, False
    grdTerms.MaxRows = 0
    grdBlock.ClearRange 1, 1, grdBlock.MaxCols, grdBlock.MaxRows, False
    grdBlock.MaxRows = 0
    Unload Me
    Exit Sub
End Sub
Public Function Get_GridTerms(vstrSubDomain As String) As String
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset


    On Error GoTo Error_Get_GridTerms
    Screen.MousePointer = vbHourglass
    clsGridFormat.List = vbNullString
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Get_Attr_Valid_Value_SubDom_Alphabetic '" & vstrSubDomain & "'"
    Set rsTerms = cnTerms.Execute(strSQL)
    Do Until rsTerms.EOF
        If Get_GridTerms = vbNullString Then
            Get_GridTerms = IsProperCase(rsTerms.Fields("ATTRBTE_VALID_VALUE"))
        Else
            Get_GridTerms = Get_GridTerms & Chr(9) & IsProperCase(rsTerms.Fields("ATTRBTE_VALID_VALUE"))
        End If
        rsTerms.MoveNext
    Loop
    clsGridFormat.List = Get_GridTerms
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_GridTerms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_GridTerms", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub grdBlock_Change(ByVal Col As Long, ByVal Row As Long)
    mbooData = True
End Sub

Private Sub grdBlock_Click(ByVal Col As Long, ByVal Row As Long)
Dim Idx As Integer
Dim strProv As String
Dim strType As String
Dim lngPrimKey As Long
Dim varCellValue
Dim lngCol As Long
Dim lngRow As Long
Dim iAnswer As Integer

    On Error GoTo Error_Block_Click
    If Row <= 0 Then Exit Sub
    If mlngGridRow > 0 And mlngGridRow <> Row And mbooData Then
        iAnswer = MsgBox("Do you want to save your previous changes? Please Save Changes before updating or adding a New Block Time", vbYesNo, "PatientTrac Block Times")
        If iAnswer = vbYes Then
            Save_Block_Mode
            Exit Sub
        ElseIf iAnswer = vbNo Then
            With grdBlock
                .BlockMode = True
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = mlngGridRow
                    .Row2 = mlngGridRow
                    .BackColor = NormalColor
                .BlockMode = False
            End With
            mlngGridRow = 0
            mbooData = False
            Read_Blocks_ByProvider udtBlock.ProviderID
        End If
    End If
    mlngGridRow = Row
    If cboProviders.Text = vbNullString Then
        MsgBox "Please select a provider first!", vbInformation, "PatientTrac Block Time"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    With grdBlock
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then
            If .CellTag = "PROVIDER_ID" Then
                lngCol = Col
                lngRow = Row
                Get_Appt_Providers
                Format_CellTypeCombo grdBlock, lngRow, lngCol, "1", clsGridFormat.List, "TypeHAlignLeft", "PROVIDER_ID"
                grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            
            ElseIf .CellTag = "FACILITY_ID" Then
                clsGridFormat.Text = vbNullString
                clsGridFormat.CellNote = vbNullString
                lngCol = Col
                lngRow = Row
                If .CellNote <> vbNullString Then
                    clsGridFormat.CellNote = .CellNote
                    .GetText Col, Row, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        clsGridFormat.Text = Trim(varCellValue)
                    End If
                End If
                Load_Facilities
                Format_CellTypeCombo grdBlock, lngRow, lngCol, "1", clsGridFormat.List, "TypeHAlignLeft", "FACILITY_ID"
                Set_GridCurSel grdBlock, clsGridFormat.Text, lngCol, lngRow, "1", "FACILITY_ID", 0, clsGridFormat.CellNote
                grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            ElseIf .CellTag = "REASON" Then
                lngCol = Col
                lngRow = Row
                clsGridFormat.List = Get_GridTerms("sd_Appointment_BlockReason")
                Format_CellTypeCombo grdBlock, lngRow, lngCol, "1", clsGridFormat.List, "TypeHAlignLeft", "REASON"
                grdBlock.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Block_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Block_Click", App.EXEName & "." & TypeName(Me)
    Exit Sub
End Sub

Private Sub grdBlock_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
Dim Idx As Integer
Dim strProv As String
Dim strType As String
Dim lngPrimKey As Long
Dim varCellValue
Dim iAnswer As Integer


    On Error GoTo Error_Combo_Close
    Screen.MousePointer = vbHourglass
    If mlngGridRow > 0 And mlngGridRow <> Row And mbooData Then
        iAnswer = MsgBox("Do you want to save your previous changes? Please Save Changes before updating or adding a New Block Time", vbYesNo, "PatientTrac Block Times")
        If iAnswer = vbYes Then
            Save_Block_Mode
            Exit Sub
        ElseIf iAnswer = vbNo Then
            With grdBlock
                .BlockMode = True
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = mlngGridRow
                    .Row2 = mlngGridRow
                    .BackColor = NormalColor
                .BlockMode = False
            End With
            mlngGridRow = 0
            mbooData = False
            Read_Blocks_ByProvider udtBlock.ProviderID
        End If
    End If
    mbooData = True
    mlngGridRow = Row
    With grdBlock
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then
            If .CellTag = "PROVIDER_ID" Then
                strType = "Provider"
                .GetText Col, Row, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    strProv = varCellValue
                    lngPrimKey = LookUp_Provider(strProv)
                End If
                If lngPrimKey > 0 And strProv <> vbNullString Then
                    Set_GridCurSel grdBlock, strProv, Col, Row, "1", "PROVIDER_ID", 0, CStr(lngPrimKey)
                    grdBlock.SetCellBorder Col, Row, Col, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
            ElseIf .CellTag = "FACILITY_ID" Then
                strType = "Facility"
                .GetText Col, Row, varCellValue
                If Trim(varCellValue) <> vbNullString Then
                    strProv = varCellValue
                    lngPrimKey = LookUp_Facility(strProv)
                    udtBlock.FacilityID = lngPrimKey
                End If
                If lngPrimKey > 0 And strProv <> vbNullString Then
                    Set_GridCurSel grdBlock, strProv, Col, Row, "1", "FACILITY_ID", 0, CStr(lngPrimKey)
                    grdBlock.SetCellBorder Col, Row, Col, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
            End If
        End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Combo_Close:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Combo_Close", App.EXEName & "." & TypeName(Me)
    Exit Sub
End Sub

Private Sub grdBlock_ComboDropDown(ByVal Col As Long, ByVal Row As Long)
Dim iAnswer As Integer

    If Row <= 0 Then Exit Sub
    mbooData = True
    If mlngGridRow > 0 And mlngGridRow <> Row And mbooData Then
        iAnswer = MsgBox("Do you want to save your previous changes? Please Save Changes before updating or adding a New Block Time", vbYesNo, "PatientTrac Block Times")
        If iAnswer = vbYes Then
            Save_Block_Mode
            Exit Sub
        ElseIf iAnswer = vbNo Then
            With grdBlock
                .BlockMode = True
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = mlngGridRow
                    .Row2 = mlngGridRow
                    .BackColor = NormalColor
                .BlockMode = False
            End With
            mlngGridRow = 0
            mbooData = False
            Read_Blocks_ByProvider udtBlock.ProviderID
        End If
    End If
End Sub

Private Sub grdBlock_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim lngCol As Long
Dim lngx As Long
Dim lngY As Long
Dim lngWidth As Long
Dim lngHeight As Long
Dim varCellValue
Dim Idx As Integer
Dim Arr
Dim iAnswer As Integer

    If Row <= 0 Then Exit Sub
    If mlngGridRow > 0 And mlngGridRow <> Row And mbooData Then
        iAnswer = MsgBox("Do you want to save your previous changes? Please Save Changes before updating or adding a New Block Time", vbYesNo, "PatientTrac Block Times")
        If iAnswer = vbYes Then
            Save_Block_Mode
            Exit Sub
        ElseIf iAnswer = vbNo Then
            With grdBlock
                .BlockMode = True
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = mlngGridRow
                    .Row2 = mlngGridRow
                    .BackColor = NormalColor
                .BlockMode = False
            End With
            mlngGridRow = 0
            mbooData = False
            Read_Blocks_ByProvider udtBlock.ProviderID
        End If
    End If
    With grdBlock
        mlngGridRow = Row
        For lngRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                .BackColor = NormalColor
            .BlockMode = False
        Next lngRow
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then
            mlngGridCol = Col
            mlngGridRow = Row
            Select Case .CellTag
                Case "START_DATE", "END_DATE", "DAYWEEK", "DURATION"
                    .GetCellPos Col, Row, lngx, lngY, lngWidth, lngHeight
                    fraDate.ZOrder 0
                    Arr = Split(LookUpCellTag("START_DATE"), "|")
                    If UBound(Arr) > 0 Then
                      lngCol = Arr(0)
                    End If
                    grdBlock.GetText lngCol, Row, varCellValue
                    If IsDate(varCellValue) Then
                        ccdStartDate.Format = dtpShortDate
                        ccdStartDate.Value = Format(varCellValue, "mm/dd/yyyy")
                    Else
                        ccdStartDate.Format = dtpShortDate
                        ccdStartDate.Value = Format(Now, "mm/dd/yyyy")
                    End If
                    Arr = Split(LookUpCellTag("END_DATE"), "|")
                    If UBound(Arr) > 0 Then
                      lngCol = Arr(0)
                    End If
                    grdBlock.GetText lngCol, Row, varCellValue
                    If IsDate(varCellValue) Then
                        ccdEndDate.Format = dtpShortDate
                        ccdEndDate.Value = Format(varCellValue, "mm/dd/yyyy")
                    Else
                        ccdEndDate.Format = dtpShortDate
                        ccdEndDate.Value = Format(Now, "mm/dd/yyyy")
                    End If
                    Arr = Split(LookUpCellTag("DURATION"), "|")
                    If UBound(Arr) > 0 Then
                      lngCol = Arr(0)
                    End If
                    grdBlock.GetText lngCol, Row, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        For Idx = 0 To cboDuration.listcount - 1
                            If cboDuration.List(Idx) = Trim(varCellValue) Then
                                cboDuration.Text = cboDuration.List(Idx)
                                Exit For
                            End If
                        Next Idx
                    Else
                        cboDuration.ListIndex = -1
                    End If
                    If cboDuration.Text <> varCellValue Then
                        Arr = Split(varCellValue, " ")
                        If UBound(Arr) > 0 Then
                            cctDuration.Text = Arr(0)
                             For Idx = 0 To cboDuration.listcount - 1
                                If cboDuration.List(Idx) = Trim(Arr(1)) Then
                                    cboDuration.Text = cboDuration.List(Idx)
                                    Exit For
                                End If
                            Next Idx
                        End If
                    End If
                    Arr = Split(LookUpCellTag("DAYWEEK"), "|")
                    If UBound(Arr) > 0 Then
                      lngCol = Arr(0)
                    End If
                    grdBlock.GetText lngCol, Row, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        For Idx = 0 To cboWeek.listcount - 1
                            If cboWeek.List(Idx) = Trim(varCellValue) Then
                                cboWeek.Text = cboWeek.List(Idx)
                                Exit For
                            End If
                        Next Idx
                    Else
                        cboWeek.ListIndex = -1
                    End If
'----------------------
                    udtBlock.StartTime = vbNullString
                    udtBlock.EndTime = vbNullString
                    Arr = Split(LookUpCellTag("START_TIME"), "|")
                    If UBound(Arr) > 0 Then
                      lngCol = Arr(0)
                    End If
                    grdBlock.GetText lngCol, Row, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        udtBlock.StartTime = varCellValue
                    End If
                    lngCol = 0
                    Arr = Split(LookUpCellTag("END_TIME"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                    End If
                    If lngCol > 0 Then
                        grdBlock.GetText lngCol, Row, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                           udtBlock.EndTime = varCellValue
                        End If
                    End If
                    fraDate.Visible = True
                Case "START_TIME", "END_TIME"
                    .GetCellPos Col, Row, lngx, lngY, lngWidth, lngHeight
                    fraNumber.ZOrder 0
                    Arr = Split(LookUpCellTag("START_TIME"), "|")
                    If UBound(Arr) > 0 Then
                      lngCol = Arr(0)
                    End If
                    grdBlock.GetText lngCol, Row, varCellValue
                    If Trim(varCellValue) <> vbNullString Then
                        For Idx = 0 To cboRepTimeFrom.listcount - 1
                            If cboRepTimeFrom.List(Idx) = varCellValue Then
                                cboRepTimeFrom.Text = cboRepTimeFrom.List(Idx)
                                Exit For
                            End If
                        Next Idx
                    End If
                    lngCol = 0
                    Arr = Split(LookUpCellTag("END_TIME"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                    End If
                    If lngCol > 0 Then
                        grdBlock.GetText lngCol, Row, varCellValue
                        If Trim(varCellValue) <> vbNullString Then
                           For Idx = 0 To cboRepTimeTo.listcount - 1
                               If cboRepTimeTo.List(Idx) = varCellValue Then
                                   cboRepTimeTo.Text = cboRepTimeTo.List(Idx)
                                   Exit For
                               End If
                           Next Idx
                        End If
                    End If
                    fraNumber.Visible = True
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
    Exit Sub
End Sub
Public Sub SetEndDate(ipDuration As String)
Dim strDay As String
Dim Arr
Dim Idx%
Dim lngNo As Long
Dim lngDateAdd As Double
    strDay = ipDuration
    Arr = Split(strDay, " ")
    If UBound(Arr) > 0 Then
        strDay = Arr(1)
        lngNo = Arr(0)
    Else
        strDay = Trim(ipDuration)
        If IsNumeric(cctDuration.Text) Then
             lngNo = cctDuration.Text
        End If
    End If
    If lngNo = 0 Then lngNo = 1
    ccdEndDate.Enabled = True
    If IsNull(ccdStartDate.Value) Then
        MsgBox "Please select start date first!", vbInformation, "PatientTrac Block Time"
        Exit Sub
    End If
    ccdEndDate.Enabled = True
    Select Case IsProperCase(strDay)
       Case "Week", "Weeks"
            ccdEndDate.Value = DateAdd("ww", lngNo, ccdStartDate.Value)
            ccdEndDate.Value = DateAdd("d", -1, ccdEndDate.Value)
            ccdEndDate.Enabled = False
       Case "Month", "Months"
            ccdEndDate.Value = DateAdd("m", lngNo, ccdStartDate.Value)
            ccdEndDate.Value = DateAdd("d", -1, ccdEndDate.Value)
            ccdEndDate.Enabled = False
       Case "Day", "Days"
            ccdEndDate.Value = DateAdd("d", lngNo, ccdStartDate.Value)
            ccdEndDate.Value = DateAdd("d", -1, ccdEndDate.Value)
            ccdEndDate.Enabled = False
       Case "Year", "Years"
            ccdEndDate.Value = DateAdd("y", lngNo, ccdStartDate.Value)
            ccdEndDate.Value = DateAdd("d", -1, ccdEndDate.Value)
            ccdEndDate.Enabled = False
       Case Else
            ccdEndDate.Enabled = True
    End Select
    Exit Sub
End Sub

Private Sub grdBlock_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    With grdBlock
        .Col = Col
        .Row = Row
        If .CellType = CellTypeComboBox Then
            .TypeComboBoxCurSel = -1
        Else
            If .CellTag <> vbNullString Then
                clsGridFormat.CellTag = .CellTag
            End If
            Format_CellTypeEdit grdBlock, Row, Col, "1", "", "TypeHAlignCenter", clsGridFormat.CellTag
            grdBlock.SetCellBorder Col, Row, Col, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
    End With
    mlngGridRow = 0
    Exit Sub
End Sub


Private Sub grdTerms_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
End Sub


Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
End Sub


Private Sub mnu_Conflict_Click()
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim bFound As Boolean
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field

    On Error GoTo Error_Conflict_Click
    Screen.MousePointer = vbHourglass
    With grdBlock
        If .ActiveRow > 0 Then
            lngPrimKey = .GetRowItemData(.ActiveRow)
            bFound = True
        End If
    End With
    If Not bFound Then
        MsgBox "Please Select the appropriate block time row first!", vbInformation, "PatientTrac Block Time"
    Else
        If lngPrimKey > 0 Then
            If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
            strSQL = "SELECT * FROM APPOINTMENTS_BLOCKS WHERE BLOCK_ID = " & lngPrimKey
            Set rsTerms = cnTerms.Execute(strSQL)
            With rsTerms
                If Not .EOF Then
                    Clear_BlockTerm
                    Set Flds = .Fields
                    For Each Fld In Flds
                        If Not IsNull(Fld.Value) Then
                            If Trim(Fld.Value) <> vbNullString Then
                                Select Case UCase(Fld.Name)
                                    Case "BLOCK_ID"
                                        lngPrimKey = Trim(Fld.Value)
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
                                    Case "DAYWEEK"
                                         udtBlock.DayWeek = Trim(Fld.Value)
                                End Select
                            End If
                        End If
                    Next
                    Get_SchedConflict
                End If
            End With
        End If
    End If
    If grdTerms.MaxRows > 0 Then
        fraSchedConflicts.Visible = True
    Else
        MsgBox "No Conflicting appointments found!", vbInformation, "PatientTrac Block Times"
    End If
    If cnTerms.State = adStateClosed Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Conflict_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Conflict_Click", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Sub mnu_Date_Range_Click()
    fraSubmitDate.Visible = True
    fraSubmitDate.ZOrder 0
End Sub

Private Sub mnu_PrintBlock_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(1) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Scheduling Block Time"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    With grdBlock
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Block Time Appointments"
        .PrintFooter = "/cPage #/p/n"
        .PrintColor = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintSmartPrint = True
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Phone Messages"
        .PrintHeader = strHeader(1) & vbNewLine & strHeader(0)
        .PrintFooter = "/cPage #/p/n" & strHeader(0)
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
        .PrintOrientation = PrintOrientationLandscape
        .PrintShadows = False
         .PrintSheet
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdPrint_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Print", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub mnu_Show_Click()
    Read_Blocks_Provider_All
    Exit Sub
End Sub


