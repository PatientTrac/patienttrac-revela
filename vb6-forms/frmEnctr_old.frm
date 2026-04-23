VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{6514F5A0-641C-11D2-9FD0-0020AF131A57}#3.0#0"; "fpFlp30.ocx"
Begin VB.Form frmEnctr_old 
   Caption         =   "Patient Visits"
   ClientHeight    =   9120
   ClientLeft      =   1665
   ClientTop       =   2850
   ClientWidth     =   17505
   Icon            =   "frmEnctr_old.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   17505
   Begin VB.Frame fraTitleBar 
      BackColor       =   &H00E0E0E0&
      Height          =   2925
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   14505
      Begin VB.TextBox cctRend 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9600
         TabIndex        =   85
         Text            =   "Type Last Name"
         Top             =   2460
         Width           =   4035
      End
      Begin VB.ListBox lstReferral 
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
         Left            =   4230
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   84
         Top             =   1920
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox cctReferral 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4230
         TabIndex        =   83
         Text            =   "Type Last Name"
         Top             =   1500
         Width           =   3375
      End
      Begin VB.Frame fraVisitReason 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
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
         Height          =   555
         Left            =   7620
         TabIndex        =   66
         Top             =   1440
         Width           =   6495
         Begin VB.TextBox txtVisitReason 
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
            Height          =   495
            Left            =   1980
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   68
            Top             =   0
            Width           =   4425
         End
         Begin VB.ComboBox cboVisitReason 
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
            Height          =   315
            Left            =   1995
            Style           =   2  'Dropdown List
            TabIndex        =   67
            ToolTipText     =   "Please select Type of Post Operative Visit"
            Top             =   60
            Visible         =   0   'False
            Width           =   4395
         End
         Begin PT_XP_Button.PT_XP cmdVisitReas 
            Height          =   255
            Left            =   5430
            TabIndex        =   69
            ToolTipText     =   "Search Variables"
            Top             =   210
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
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
            TransparentColor=   14215660
            Caption         =   "..."
            DepthMode       =   1
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmEnctr_old.frx":09EA
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblVisit 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Visit Reason "
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
            Height          =   315
            Left            =   90
            TabIndex        =   70
            Top             =   120
            Width           =   1365
         End
      End
      Begin VB.Frame fraEnctrDetail 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   2340
         TabIndex        =   52
         Top             =   120
         Width           =   11775
         Begin LpADOLib.fpComboAdo cboFacility 
            Height          =   360
            Left            =   900
            TabIndex        =   87
            Top             =   30
            Width           =   4365
            _Version        =   196608
            _ExtentX        =   7699
            _ExtentY        =   635
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   9.75
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
            BorderDropShadowWidth=   0
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
            ColDesigner     =   "frmEnctr_old.frx":0D3C
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
            Left            =   7260
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   0
            Width           =   4455
         End
         Begin VB.ComboBox cboFacility_Old 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   60
            Width           =   4305
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
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmEnctr_old.frx":1037
            Left            =   1860
            List            =   "frmEnctr_old.frx":1039
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Tag             =   "Place of Service"
            Top             =   465
            Width           =   3375
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
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmEnctr_old.frx":103B
            Left            =   7260
            List            =   "frmEnctr_old.frx":103D
            Style           =   2  'Dropdown List
            TabIndex        =   54
            ToolTipText     =   "New or Established Patient"
            Top             =   435
            Width           =   4455
         End
         Begin VB.ComboBox cboPurVisit 
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
            Left            =   7260
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   870
            Width           =   4455
         End
         Begin PT_XP_Button.PT_XP cmdUpdProc 
            Height          =   195
            Left            =   12210
            TabIndex        =   58
            ToolTipText     =   "Upate Surgical Procedure"
            Top             =   1110
            Visible         =   0   'False
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   344
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
            TransparentColor=   14215660
            Caption         =   "..."
            DepthMode       =   1
            DepthEvent      =   1
            ForeColorDisabled=   8388608
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmEnctr_old.frx":103F
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblProvider 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   5400
            TabIndex        =   65
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lblFacility 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Height          =   210
            Left            =   30
            TabIndex        =   64
            Top             =   120
            Width           =   765
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
            Left            =   30
            TabIndex        =   63
            Top             =   570
            Width           =   1755
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Status"
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
            Left            =   5400
            TabIndex        =   62
            Top             =   510
            Width           =   1410
         End
         Begin VB.Label lblUser 
            BackStyle       =   0  'Transparent
            Caption         =   "Current Logon User: "
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
            Left            =   30
            TabIndex        =   61
            Top             =   870
            Width           =   1785
         End
         Begin VB.Label lblUserName 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   465
            Left            =   1890
            TabIndex        =   60
            Top             =   870
            Width           =   3345
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblPurVisit 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Encounter Type"
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
            Left            =   5400
            TabIndex        =   59
            Top             =   900
            Width           =   1725
         End
      End
      Begin VB.TextBox cctRefReason 
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
         Height          =   495
         Left            =   4230
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   1950
         Width           =   3375
      End
      Begin VB.TextBox cctDiagnosis 
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
         Height          =   435
         Left            =   9600
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Top             =   2040
         Width           =   4455
      End
      Begin PT_XP_Button.PT_XP cmdAdminHistory 
         Height          =   495
         Left            =   150
         TabIndex        =   71
         Top             =   630
         Width           =   1965
         _ExtentX        =   3466
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
         TransparentColor=   14215660
         Caption         =   "Hospital Admission History"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPatient 
         Height          =   405
         Left            =   150
         TabIndex        =   72
         Top             =   150
         Width           =   1965
         _ExtentX        =   3466
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
         TransparentColor=   14215660
         Caption         =   "Patient Select"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdBillingInfo 
         Height          =   495
         Left            =   150
         TabIndex        =   73
         Top             =   1890
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
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
         TransparentColor=   14215660
         Caption         =   "Billing Information"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSurgLog 
         Height          =   465
         Left            =   150
         TabIndex        =   74
         Top             =   1170
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
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
         Picture         =   "frmEnctr_old.frx":1391
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Surgery Log"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":1D8B
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClear 
         Height          =   255
         Left            =   3690
         TabIndex        =   75
         ToolTipText     =   "Clear Referral"
         Top             =   1560
         Width           =   345
         _ExtentX        =   609
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
         Picture         =   "frmEnctr_old.frx":270D
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "..."
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":2CA7
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdDiag 
         Height          =   345
         Left            =   9150
         TabIndex        =   76
         ToolTipText     =   "Select Diagnosi"
         Top             =   2130
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         Caption         =   "Dx"
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":2FF9
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClearRend 
         Height          =   345
         Left            =   9150
         TabIndex        =   77
         ToolTipText     =   "Clear Referral Physician List"
         Top             =   2520
         Width           =   375
         _ExtentX        =   661
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
         Picture         =   "frmEnctr_old.frx":334B
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   ""
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":36E5
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRendProvDefault 
         Height          =   315
         Left            =   13710
         TabIndex        =   82
         ToolTipText     =   "Add Current Facility Provider"
         Top             =   2520
         Width           =   345
         _ExtentX        =   609
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
         Picture         =   "frmEnctr_old.frx":3A37
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":3FD1
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdStatus 
         Height          =   315
         Left            =   14115
         TabIndex        =   88
         ToolTipText     =   "Load Visit Type"
         Top             =   1005
         Width           =   345
         _ExtentX        =   609
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
         Picture         =   "frmEnctr_old.frx":4323
         AppearanceThemes=   3
         TransparentColor=   15790320
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":48BD
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblRendProv 
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
         Height          =   540
         Left            =   2340
         TabIndex        =   81
         Top             =   1470
         Width           =   1680
      End
      Begin VB.Label lblRefReason 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Referral Reason "
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
         Height          =   315
         Left            =   2340
         TabIndex        =   80
         Top             =   1980
         Width           =   1695
      End
      Begin VB.Label lblRendering 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rendering Provider"
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
         Left            =   7770
         TabIndex        =   79
         Top             =   2460
         Width           =   1170
      End
      Begin VB.Label lblDiag 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Diagnosis"
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
         Left            =   7770
         TabIndex        =   78
         Top             =   2070
         Width           =   945
      End
   End
   Begin VB.Frame fraEnctr 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6165
      Left            =   30
      TabIndex        =   17
      Top             =   2910
      Width           =   14475
      Begin VB.ListBox lstRend 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   9600
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   86
         Top             =   0
         Width           =   4035
      End
      Begin VB.Frame fraMaint 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Encounter ToolBar"
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
         Height          =   1515
         Left            =   90
         TabIndex        =   25
         Top             =   150
         Width           =   3315
         Begin VB.Frame fraAppoint 
            BackColor       =   &H00E0E0E0&
            Height          =   1125
            Left            =   60
            TabIndex        =   26
            Top             =   2010
            Visible         =   0   'False
            Width           =   5745
            Begin VB.ComboBox ccNextAppt 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   2730
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   900
               Visible         =   0   'False
               Width           =   2595
            End
            Begin VB.CheckBox chkNoAppt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "No Appointment"
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
               Left            =   2760
               TabIndex        =   27
               Top             =   600
               Width           =   2625
            End
            Begin MSComCtl2.DTPicker ccApptDate 
               Height          =   375
               Left            =   30
               TabIndex        =   29
               Top             =   120
               Width           =   2625
               _ExtentX        =   4630
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
               Format          =   114098179
               CurrentDate     =   38792
            End
            Begin PT_XP_Button.PT_XP cmdAddTerms 
               Height          =   285
               Left            =   5370
               TabIndex        =   30
               ToolTipText     =   "Add Terms"
               Top             =   960
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               ForeColor       =   8388608
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AppearanceThemes=   3
               TransparentColor=   14215660
               Caption         =   "..."
               DepthEvent      =   1
               ForeColorOver   =   8388608
               ForeColorFocus  =   8388608
               ForeColorDown   =   8388608
               PictureDisabled =   "frmEnctr_old.frx":4C0F
               PictureAlignment=   4
               ShowFocus       =   -1  'True
            End
            Begin VB.Label lblNextAppt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Next Appointment"
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
               Height          =   210
               Left            =   0
               TabIndex        =   31
               Top             =   1050
               Width           =   1605
            End
         End
         Begin PT_XP_Button.PT_XP cmdSave 
            Height          =   345
            Left            =   1980
            TabIndex        =   32
            Top             =   330
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
            Picture         =   "frmEnctr_old.frx":4F61
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "Save"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmEnctr_old.frx":54FB
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdNew 
            Height          =   345
            Left            =   120
            TabIndex        =   33
            Top             =   300
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
            Picture         =   "frmEnctr_old.frx":584D
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "New"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmEnctr_old.frx":5DE7
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdUpdate 
            Height          =   345
            Left            =   1980
            TabIndex        =   34
            Top             =   810
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
            Picture         =   "frmEnctr_old.frx":6139
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "Update"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmEnctr_old.frx":66D3
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdEndEncounter 
            Height          =   345
            Left            =   120
            TabIndex        =   35
            Top             =   720
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
            Picture         =   "frmEnctr_old.frx":6A25
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "End"
            Depth           =   5
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmEnctr_old.frx":6FBF
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdDeleteEncounter 
            Height          =   345
            Left            =   120
            TabIndex        =   36
            Top             =   1140
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
            Picture         =   "frmEnctr_old.frx":7311
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "Delete"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmEnctr_old.frx":78AB
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblEnctrTools 
            BackStyle       =   0  'Transparent
            Caption         =   "Encounter Toolbar"
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
            Left            =   690
            TabIndex        =   37
            Top             =   0
            Width           =   2025
         End
      End
      Begin VB.Frame fraAdmission 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   3360
         TabIndex        =   18
         Top             =   90
         Visible         =   0   'False
         Width           =   6165
         Begin MSComCtl2.DTPicker ccdDisharge 
            Height          =   375
            Left            =   3600
            TabIndex        =   19
            Top             =   270
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
            CustomFormat    =   "MM/dd/yyy hh:mm:ss"
            DateIsNull      =   -1  'True
            Format          =   114098179
            CurrentDate     =   38862.7362384259
            MinDate         =   37257
         End
         Begin MSComCtl2.DTPicker ccdAdmit 
            Height          =   375
            Left            =   690
            TabIndex        =   20
            Top             =   270
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
            CustomFormat    =   "MM/dd/yyy hh:mm:ss"
            DateIsNull      =   -1  'True
            Format          =   114098179
            CurrentDate     =   38862.7362384259
            MinDate         =   37257
         End
         Begin PT_XP_Button.PT_XP cmdDischarge 
            Height          =   225
            Left            =   5400
            TabIndex        =   21
            ToolTipText     =   "Enter Discharge Date"
            Top             =   0
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
            Picture         =   "frmEnctr_old.frx":7BFD
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "OK"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmEnctr_old.frx":7F97
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdAdmit 
            Height          =   225
            Left            =   2550
            TabIndex        =   22
            ToolTipText     =   "Enter/Reset Admission Date"
            Top             =   0
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
            Picture         =   "frmEnctr_old.frx":82E9
            AppearanceThemes=   3
            TransparentColor=   14215660
            Caption         =   "OK"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmEnctr_old.frx":8683
            ShowFocus       =   -1  'True
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Discharge Date"
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
            Left            =   3630
            TabIndex        =   24
            Top             =   30
            Width           =   1725
         End
         Begin VB.Label lblAdmt 
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
            Height          =   165
            Left            =   690
            TabIndex        =   23
            Top             =   60
            Width           =   1695
         End
      End
      Begin FPSpreadADO.fpSpread grdEnctrSummary 
         Height          =   3795
         Left            =   0
         TabIndex        =   38
         Top             =   1740
         Width           =   14235
         _Version        =   458752
         _ExtentX        =   25109
         _ExtentY        =   6694
         _StockProps     =   64
         AutoCalc        =   0   'False
         DAutoCellTypes  =   0   'False
         DAutoFill       =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   15
         MaxRows         =   50
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   16761024
         ShadowText      =   8388608
         SpreadDesigner  =   "frmEnctr_old.frx":89D5
         UnitType        =   0
         UserResize      =   0
         Appearance      =   2
         ClipboardOptions=   0
      End
      Begin PT_XP_Button.PT_XP cmdNext 
         Height          =   405
         Left            =   6630
         TabIndex        =   39
         Top             =   1170
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
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
         Picture         =   "frmEnctr_old.frx":8CC8
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Next"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":9062
         PictureAlignment=   5
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrev 
         Height          =   405
         Left            =   5250
         TabIndex        =   40
         Top             =   1170
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
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
         Picture         =   "frmEnctr_old.frx":93B4
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Previous "
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":974E
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCancel 
         Height          =   345
         Left            =   7320
         TabIndex        =   41
         Top             =   5640
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
         Picture         =   "frmEnctr_old.frx":9AA0
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":A4B2
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdEnctrOK 
         Height          =   345
         Left            =   6090
         TabIndex        =   42
         Top             =   5640
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
         Picture         =   "frmEnctr_old.frx":A804
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "OK"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":AB9E
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdEntrDefault 
         Height          =   345
         Left            =   12360
         TabIndex        =   43
         Top             =   5700
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
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
         Picture         =   "frmEnctr_old.frx":AEF0
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Set Defaults"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":B28A
         ShowFocus       =   -1  'True
      End
      Begin MSComCtl2.DTPicker ccdStartDate 
         Height          =   375
         Left            =   10920
         TabIndex        =   44
         Top             =   210
         Width           =   2745
         _ExtentX        =   4842
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
         Format          =   114098179
         CurrentDate     =   38862.7355555556
         MinDate         =   37257
      End
      Begin MSComCtl2.DTPicker ccdEndDate 
         Height          =   375
         Left            =   10920
         TabIndex        =   45
         Top             =   720
         Width           =   2745
         _ExtentX        =   4842
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
         Format          =   114098179
         CurrentDate     =   38862.7362384259
         MinDate         =   38353
      End
      Begin VB.Label lblHistory 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encounter History"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   5550
         OLEDropMode     =   1  'Manual
         TabIndex        =   48
         Top             =   840
         Width           =   2130
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         Left            =   9780
         TabIndex        =   47
         Top             =   810
         Width           =   915
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
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
         Left            =   9750
         TabIndex        =   46
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.Frame fraEnctrProc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Encounter Type - Procedure:     Select All Relevant Procedures"
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
      Height          =   4785
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   13875
      Begin VB.ComboBox cboRedo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4050
         Width           =   2445
      End
      Begin VB.ComboBox cboFilter 
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
         Height          =   315
         Left            =   7260
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   420
         Width           =   2205
      End
      Begin VB.TextBox cctProcString 
         Height          =   405
         Left            =   10560
         TabIndex        =   6
         Text            =   " "
         Top             =   4020
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00E0E0E0&
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
         Height          =   1335
         Left            =   15150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   5505
      End
      Begin VB.TextBox cctSearch 
         Height          =   375
         Left            =   750
         TabIndex        =   1
         Top             =   450
         Width           =   4755
      End
      Begin PT_XP_Button.PT_XP cmdCloseList 
         Height          =   405
         Left            =   6210
         TabIndex        =   3
         Top             =   4230
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
         Picture         =   "frmEnctr_old.frx":B5DC
         AppearanceThemes=   3
         BorderDrawEvent =   1
         TransparentColor=   13160660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":BFEE
         ShowFocus       =   -1  'True
      End
      Begin VB.ListBox lstListBox 
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
         Height          =   2100
         Left            =   14160
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   930
         Width           =   6615
      End
      Begin FPSpreadADO.fpSpread grdColumn 
         Height          =   2925
         Left            =   5970
         TabIndex        =   7
         Top             =   1020
         Width           =   7785
         _Version        =   458752
         _ExtentX        =   13732
         _ExtentY        =   5159
         _StockProps     =   64
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
         MaxCols         =   6
         MaxRows         =   0
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   16761024
         ShadowText      =   8388608
         SpreadDesigner  =   "frmEnctr_old.frx":C340
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
      Begin FPSpreadADO.fpSpread grdProc 
         Height          =   2925
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   5745
         _Version        =   458752
         _ExtentX        =   10134
         _ExtentY        =   5159
         _StockProps     =   64
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
         MaxCols         =   2
         MaxRows         =   0
         MoveActiveOnFocus=   0   'False
         ProcessTab      =   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   16761024
         ShadowText      =   8388608
         SpreadDesigner  =   "frmEnctr_old.frx":C5EE
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin PT_XP_Button.PT_XP cmdDone 
         Height          =   315
         Left            =   9360
         TabIndex        =   9
         Top             =   3990
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
         Picture         =   "frmEnctr_old.frx":C887
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Done"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":CC21
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdReportToolBox 
         Height          =   435
         Left            =   9690
         TabIndex        =   11
         Top             =   360
         Width           =   1635
         _ExtentX        =   2884
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
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Show All CPT"
         Depth           =   5
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":CF73
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdTermRedo 
         Height          =   345
         Left            =   3930
         TabIndex        =   14
         ToolTipText     =   "Revision = Practice Redo Procedure, Redo = Other Surgeon Procedure Rework"
         Top             =   4050
         Width           =   405
         _ExtentX        =   714
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
         Picture         =   "frmEnctr_old.frx":DBC5
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmEnctr_old.frx":E15F
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdTest 
         Height          =   585
         Left            =   15780
         TabIndex        =   16
         Top             =   1230
         Width           =   2475
         _Version        =   458752
         _ExtentX        =   4366
         _ExtentY        =   1032
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmEnctr_old.frx":E4B1
      End
      Begin VB.Label lblRedo 
         BackStyle       =   0  'Transparent
         Caption         =   "Revision Type"
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
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblProcFilter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Show CPT by:"
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
         Height          =   225
         Left            =   5970
         TabIndex        =   12
         Top             =   450
         Width           =   1425
      End
      Begin VB.Label lblSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search"
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
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   330
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmEnctr_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fpimage As Picture
Private fstrDefProvider As String
Public piPatientID As Long
Public piEncounterID As Long
Public piProviderID As Long
Private mlngSelFacility As Long
Private fUpdateMode As String
Private mLocID As Integer
Private mbAdmit As Boolean
Private mbDischarge As Boolean
Private mRoomAssign As String
Private mNewPatient As Boolean
Private mstrSaveData As String
Private mRow As Integer
Public mstrSearchItem As String
Public gbooBypass As Boolean
Public mFirst As Boolean
Private mstrPlan As String
'Encounter Options.........
Private fbooByPassEvent As Boolean
Private fbooEncounterTypeReq As Boolean
Private fbooCloseAnyEnctr As Boolean
Private fbooUnlimitedEnctrs As Boolean
Private fbooServicePlaceReq As Boolean
Private piSelEncntrID As Long
Private fbooEndClicked As Boolean
Private fbooSaveClicked As Boolean
Public mbooNext As Boolean
Public mbooPrev As Boolean
Private mbooUpdate As Boolean
Private fbooEncounterClosed As Boolean
Private mstrPatNote As String
Private mfpNote As Variant
Private mlngGridRow As Long
Private mlngGridCol As Long
Private mstrRendType As String
Private mbooNewEncntr As Boolean
Private mbooAddNewEnctr As Boolean
Private mstrDenyEncntrReas As String
Private strCurProvName As String
Private mbDateCheck As Boolean
Private mstrNextAppt As String
Private mstrAppDate As String
Private mbooGrid As Boolean
Private mNoClick As Boolean
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long

Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long
Private mlngProcRow As Long
Private mlngProcCol As Long
Private mstrProc As String
Private mlngProcID As Long
Private mstrColName As String
Private mstrColCPT As String
Private mlngProSubID As Long
Private mbooGridClick As Boolean
Private mlngColumnRow As Long
Private mlngColumnCol As Long
Private GrdLock As Boolean
Private Function Get_AdmissionDate()
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
Dim Arr
Dim strEnctr As String
Dim strSQL As String
Dim lngEnctrID As Long

    On Error GoTo Error_Get_AdmissionDate
    Screen.MousePointer = vbHourglass
    If udtPatientName.IntId = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If clsEncounter.Admission = "1/1/1900" Then clsEncounter.Admission = vbNullString
    If clsEncounter.Admission <> vbNullString Then
        ccdAdmit.Enabled = True
        ccdAdmit.Value = Format(clsEncounter.Admission, "mm/dd/yyyy hh:mm:ss AM/PM")
    End If
    If clsEncounter.Discharge = "1/1/1900" Then clsEncounter.Discharge = vbNullString
    If clsEncounter.Discharge <> vbNullString Then
        ccdDisharge.Enabled = True
        ccdDisharge.Value = Format(clsEncounter.Discharge, "mm/dd/yyyy hh:mm:ss AM/PM")
    End If
'    lngDateDiff = -1
'
'    booAdmit = True
'    booDischarge = True
'    lngAdmitID = 0
'    strEnctrDate = Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM")
'    If clsEncounter.EncounterType = vbNullString Then
'        clsEncounter.EncounterType = cboStatus.Text
'    End If
'    strSQL = "SELECT * FROM PATIENT_ADMISSION  WHERE PATIENT_ID = " & val(udtPatientName.IntId) & " AND " _
'        & "PURGE = 0 ORDER BY ADMISSION_DATE DESC"
'    Set rsAdmit = cnAdmit.Execute(strSQL)
'    With rsAdmit
'        Do Until .EOF
'            booDischarge = False
'            If Not IsNull(!DISCHARGE_DATE) Then
'                If (!DISCHARGE_DATE) <> vbNullString And (!DISCHARGE_DATE) <> "1/1/1900" Then
'                    sDisDate = Format((!DISCHARGE_DATE), "mm/dd/yyyy")
'                    If Trim(sDisDate) <> vbNullString Then
'                        lngDateDiff = DateDiff("d", sDisDate, strEnctrDate)
'                        If lngDateDiff <= 0 Then
'                            booDischarge = True
'                        End If
'                    End If
'                End If
'            End If
'            If Not booDischarge Then
'                ccdAdmit.Enabled = True
'            End If
'            If Not IsNull(!ADMISSION_DATE) Then
'                If (!ADMISSION_DATE) <> vbNullString And (!ADMISSION_DATE) <> "1/1/1900" Then
'                    sAdmitDate = Format((!ADMISSION_DATE), "mm/dd/yyyy")
'                End If
'            End If
'            If Trim(sAdmitDate) <> vbNullString Then
'                If Trim(sDisDate) <> vbNullString And sDisDate <> "01/01/1900" Then
'                    lngDateDiff = DateDiff("d", sDisDate, sAdmitDate)
'                Else
'                     lngDateDiff = DateDiff("d", sDisDate, sAdmitDate)
'                End If
'                If lngDateDiff <= 0 Then
'                    lngEnctrID = (!ENCOUNTER_ID)
'                    lngAdmitID = (!PATIENT_ADMITID)
'                    ccdAdmit.Enabled = True
'                    If sAdmitDate <> vbNullString Then
'                        ccdAdmit.Enabled = True
'                        ccdAdmit.Value = Format((!ADMISSION_DATE), "mm/dd/yyyy hh:mm:ss AM/PM")
'                    End If
'                    If Trim(sDisDate) <> vbNullString And sDisDate <> "01/01/1900" Then
'                        ccdDisharge.Enabled = True
'                        ccdDisharge.Value = Format(sDisDate, "mm/dd/yyyy hh:mm:ss AM/PM")
'                    Else
'                        ccdDisharge.Enabled = True
'                    End If
'                    Exit Do
'                End If
'            End If
'        .MoveNext
'        Loop
'    End With
    ccdAdmit.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_AdmissionDate:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_AdmissionDate", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


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
    strSQL = "Select * FROM PHYSICIAN_REFERRAL WHERE REFERRER_ID  = " & lngRefID
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
                    lstReferral.ItemData(lstReferral.NewIndex) = lngRefID
                Else
                    strSuffix = Fix_Suffix(strSuffix)
                    lstReferral.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strSuffix) & " " & Trim(strTitle)
                    lstReferral.ItemData(lstReferral.NewIndex) = lngRefID
                End If
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

Private Function Load_Specific_Rendering(vlngRefID As Long)
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
    lstRend.Clear
    lstRend.Visible = False
    strSQL = "Select * FROM PHYSICIAN_REFERRAL WHERE REFERRER_ID  = " & lngRefID
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
                    lstRend.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strTitle)
                    cctRend.Text = Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strTitle)
                Else
                    strSuffix = Fix_Suffix(strSuffix)
                    lstRend.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strSuffix) & " " & Trim(strTitle)
                    cctRend.Text = Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strTitle)
                End If
                mbooByPass = True
                lstRend.ItemData(lstRend.NewIndex) = lngRefID
                mbooByPass = True
                lstRend.Selected(lstRend.NewIndex) = True
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


Private Function LookUp_ChildCPT(vstrCPT As String)
Dim strSQL As String
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim lngProcID As Long
Dim iCount As Integer
Dim strChildName As String
Dim strSentence As String
Dim Arr
Dim strChildCPT As String


    On Error GoTo Error_LookUp_ChildCPT
    Screen.MousePointer = vbHourglass
    Clear_ProcDetail
    If cnCPT.State = adStateOpen Then cnCPT.Close
    If cnCPT.State = adStateClosed Then cnCPT.open psConnect(1)
    'Get ProcID
    strSQL = "SELECT  PROCEDURE_ID, PROCEDUR_NAME FROM SURG_PROCEDURE_CODES WHERE  CPT_CODE like  '" & vstrCPT & "'"
    Set rsCPT = cnCPT.Execute(strSQL)
    If Not rsCPT.EOF Then
        If rsCPT.Fields("PROCEDURE_ID") <> vbNullString Then
            lngProcID = rsCPT.Fields("PROCEDURE_ID")
        End If
    End If
    If lngProcID > 0 Then
        Arr = Split(vstrCPT, ".")
        If UBound(Arr) > 0 Then
            strChildCPT = Trim(Arr(1))
        End If
        If strChildCPT <> vbNullString Then
            For iCount = 1 To Len(strChildCPT)
                If iCount = 1 Then
                    If InStr(1, strChildCPT, Mid$(strChildCPT, iCount, 1)) <> vbNullString Then
                        strSQL = "SELECT FIRST_COLUMN FROM SURGERY_CHILD_CPT WHERE PROCEDURE_ID = " & mlngProcID
                        Set rsCPT = cnCPT.Execute(strSQL)
                        If Not rsCPT.EOF Then
                              If rsCPT.Fields("FIRST_COLUMN") <> vbNullString Then strChildName = rsCPT.Fields("FIRST_COLUMN")
                        End If
                        If Trim(strChildName) = "AREA" Or Trim(strChildName) = "BODY PART" Then
                            clsProcDetail.LipoDetails = vbNullString
                            clsProcDetail.LipoDetails = Read_Child_Selection(piPatientID, "CPT_COL_1", lngProcID)
                        End If
                    End If
                End If
'                If iCount = 2 Then clsProcDetail.ProcIQues2 = iCount
'                If iCount = 3 Then clsProcDetail.ProcIQues3 = iCount
'                If iCount = 4 Then clsProcDetail.ProcIQues4 = iCount
'                If iCount = 5 Then clsProcDetail.ProcIQues5 = iCount
'                If iCount = 6 Then clsProcDetail.ProcIQues6 = iCount
            Next
        End If
    End If
    If cnCPT.State = adStateOpen Then
        cnCPT.Close
        Set cnCPT = Nothing
        Set rsCPT = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUp_ChildCPT:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUp_ChildCPT", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function LookUp_CPT(vstrVisitType As String)
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Arr
Dim strFee As String
Dim strProvider As String
Dim strCompare As String
Dim strCPT As String

    On Error GoTo Error_LookUp_CPT
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    If Trim(clsProvider.Title) <> vbNullString Then
        strCompare = UCase(clsProvider.Title)
        strCompare = Replace(strCompare, ".", "")
        strCompare = Trim(strCompare)
    Else
        If piProviderID = 0 Then
            modSpecific.piProviderID = 0
            Get_DefaultProvider
            piProviderID = modSpecific.piProviderID
            Get_Providers piProviderID
        Else
            Get_Providers piProviderID
        End If
        strCompare = UCase(clsProvider.Title)
        strCompare = Replace(strCompare, ".", "")
        strCompare = Trim(strCompare)
    End If
    clsEncounter.Intrvn_CD = vbNullString
    strSQL = "SELECT  CPT_NAME, PROVIDER, CPT_CODE FROM CPT_TEMPLATE WHERE CPT_NAME = '" & vstrVisitType & "'"
    Set rsCodes = cnCodes.Execute(strSQL)
    Do Until rsCodes.EOF
        strProvider = vbNullString
        If rsCodes.Fields("PROVIDER") <> vbNullString Then
            strProvider = UCase(rsCodes.Fields("PROVIDER"))
            strProvider = Replace(strProvider, ".", "")
            strProvider = Trim(strProvider)
            If strProvider = strCompare Then
              Set Flds = rsCodes.Fields
              For Each Fld In Flds
                  If Not IsNull(Fld.Value) Then
                      If Fld.Value <> vbNullString Then
                         strProvider = Replace(strProvider, ".", "")
                          Select Case strProvider
                              Case "MD", "DO", "PA", "PA-C", "APN"
                                  If UCase(Fld.Name) = "CPT_CODE" Then
                                      clsEncounter.Intrvn_CD = Fld.Value
                                      Exit Do
                                  End If
                              Case "MS", "MA", "LSW", "LPN", "CLINICIAN", "LCSW", "LPC", "PHD", "LSW", "LCSW", "LPN", "LGSW", "MSW", "LGPC", "LCPC", "LCS", "APRN", _
                                   "LPC", "LCSW-C", "LCSWC", "PSY-D", "CRNP", "APN", "CSC-AD", "LCADC", "LCAD-C", "CSCAD"
                                  If UCase(Fld.Name) = "CPT_CODE" Then
                                      clsEncounter.Intrvn_CD = Fld.Value
                                      Exit Do
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
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUp_CPT:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUp_CPT", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function LookUpColName(vstrname As String) As String
Dim lngRow As Long
Dim lngPrimKey As Long
Dim lngCol As Long
Dim varCellValue
Dim strRowName As String
Dim booData As Boolean
Dim strClinData As String
Dim varData
Dim CellNote As Long
Dim Arr
Dim strFldName As String
Dim strFound As String
Dim lngInsertRow As Long
Dim booExit As Boolean

    With grdEnctrSummary
        LookUpColName = vbNullString
        For lngCol = 1 To .MaxCols
            .Col = lngCol
            .Row = SpreadHeader
            .GetText lngCol, SpreadHeader, varCellValue
            If Trim(varCellValue) <> vbNullString And TextFound(vstrname, Trim(varCellValue), Len(vstrname), True) Then
                strFound = Trim(varCellValue)
                LookUpColName = lngCol & "|" & SpreadHeader
                booExit = True
                Exit For
            End If
        Next lngCol
    End With
End Function
Private Function Collect_ProcData()
Dim varCellValue
Dim strProc As String
Dim iRow As Long
Dim booDone As Boolean
Dim varCPTCode
Dim iCol As Long
Dim varColName
Dim lngPrimKey As Long
Dim Idx As Integer
Dim strRisk As String
Dim lngRow As Long
Dim lngCol As Long
Dim CellNote As Long
Dim booData As Boolean
Dim lngColWidth As Long
Dim i%
Dim strColCPT As String
Dim strCPT1 As String
Dim strCPT2 As String
Dim strCPT3 As String
Dim strCPT4 As String
Dim strCPT5 As String
Dim strCPT6 As String


    mbooByPass = True
    booData = False
    Clear_ProcDetail
    With grdColumn
        For iCol = 1 To .MaxCols
            For iRow = 1 To .MaxRows
                .Col = iCol
                .Row = iRow
                If .BackColor = grdCurItem Then
                    booData = True
                    varColName = vbNullString
                    .GetText iCol, iRow, varCellValue
                    .GetText iCol, SpreadHeader, varColName
                    .Col = iCol
                    .Row = iRow
                    If iCol = 1 Then
                        If Trim(varCellValue) <> vbNullString Then
                            If .CellTag <> vbNullString Then
                                strCPT1 = .CellTag
                            End If
                            If Trim(clsProcDetail.ProcIQues1) = vbNullString Then
                                clsProcDetail.ProcIQues1 = varColName & "   *" & Trim(varCellValue)
                            Else
                                clsProcDetail.ProcIQues1 = clsProcDetail.ProcIQues1 & vbNewLine & "   *" & Trim(varCellValue)
                            End If
                        End If
                    ElseIf iCol = 2 Then
                        If Trim(varCellValue) <> vbNullString Then
                            If .CellTag <> vbNullString Then
                                strCPT2 = .CellTag
                            End If
                            If Trim(clsProcDetail.ProcIQues2) = vbNullString Then
                                clsProcDetail.ProcIQues2 = varColName & "   *" & Trim(varCellValue)
                            Else
                                clsProcDetail.ProcIQues2 = clsProcDetail.ProcIQues2 & vbNewLine & "   *" & Trim(varCellValue)
                            End If
                        End If
                    ElseIf iCol = 3 Then
                        If Trim(varCellValue) <> vbNullString Then
                            If .CellTag <> vbNullString Then
                                strCPT3 = .CellTag
                            End If
                            If Trim(clsProcDetail.ProcIQues3) = vbNullString Then
                                clsProcDetail.ProcIQues3 = varColName & "   *" & Trim(varCellValue)
                            Else
                                clsProcDetail.ProcIQues3 = clsProcDetail.ProcIQues3 & vbNewLine & "   *" & Trim(varCellValue)
                            End If
                        End If
                    ElseIf iCol = 4 Then
                        If Trim(varCellValue) <> vbNullString Then
                            If .CellTag <> vbNullString Then
                                strCPT4 = .CellTag
                            End If
                            If Trim(clsProcDetail.ProcIQues4) = vbNullString Then
                                clsProcDetail.ProcIQues4 = varColName & "   *" & Trim(varCellValue)
                            Else
                                clsProcDetail.ProcIQues4 = clsProcDetail.ProcIQues4 & vbNewLine & "   *" & Trim(varCellValue)
                            End If
                        End If
                    ElseIf iCol = 5 Then
                        If Trim(varCellValue) <> vbNullString Then
                            If .CellTag <> vbNullString Then
                                strCPT5 = .CellTag
                            End If
                            If Trim(clsProcDetail.ProcIQues5) = vbNullString Then
                                clsProcDetail.ProcIQues5 = varColName & "   *" & Trim(varCellValue)
                            Else
                                clsProcDetail.ProcIQues5 = clsProcDetail.ProcIQues5 & vbNewLine & "   *" & Trim(varCellValue)
                            End If
                        End If
                    ElseIf iCol = 6 Then
                        If Trim(varCellValue) <> vbNullString Then
                            If .CellTag <> vbNullString Then
                                strCPT6 = .CellTag
                            End If
                            If Trim(clsProcDetail.ProcIQues6) = vbNullString Then
                                clsProcDetail.ProcIQues6 = varColName & "   *" & Trim(varCellValue)
                            Else
                                clsProcDetail.ProcIQues6 = clsProcDetail.ProcIQues6 & vbNewLine & "   *" & Trim(varCellValue)
                            End If
                        End If
                    End If
                End If
            Next iRow
        Next iCol
    End With
    If mstrColCPT <> vbNullString Then
        Write_Child_Selection piPatientID, mstrColCPT, mlngProcID, mstrProc
    End If
    If strCPT1 = vbNullString Then strCPT1 = 0
    If strCPT2 = vbNullString Then strCPT2 = 0
    If strCPT3 = vbNullString Then strCPT3 = 0
    If strCPT4 = vbNullString Then strCPT4 = 0
    If strCPT5 = vbNullString Then strCPT5 = 0
    If strCPT6 = vbNullString Then strCPT6 = 0
    mstrColCPT = mstrColCPT & "." & strCPT1 & strCPT2 & strCPT3 & strCPT4 & strCPT5 & strCPT6
    
    grdColumn.ClearRange 1, SpreadHeader, grdColumn.MaxCols, grdColumn.MaxRows, False
   Exit Function
End Function

Private Function Format_ColumnGrid()
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim strSQL As String
Dim iRow As Integer
Dim strQuestion As String
Dim CellNote As Long
Dim iCol%
Dim varColName
Dim strSeperator As String
Dim booData As Boolean
Dim strColItem As String  'CPT Subsets

    On Error GoTo Error_Format_ColumnGrid
    With grdColumn
            .ReDraw = False
            .ClearRange 1, SpreadHeader, .MaxCols, .MaxRows, False
            .ProcessTab = True
            .FontBold = True
            .FontName = "Times New Roman"
            .FontSize = 8
            .RowHeadersShow = False
            .ColsFrozen = 2
            .ColHeaderRows = 2
            .ColHeadersShow = True
            .ColHeadersAutoText = DispBlank
            .PrintBorder = False
            .ColHeaderRows = 2
            .TextTip = TextTipFloatingFocusOnly
            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            .FontBold = True
            .ScrollBars = ScrollBarsVertical
            .ScrollBarTrack = ScrollBarTrackVertical
        Add_Spacer SpreadHeader + 1, grdProc
        If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
        strSQL = "SELECT * FROM SURGERY_CHILD_CPT WHERE PROCEDURE_ID = " & mlngProcID
        Set rsTerms = cnTerms.Execute(strSQL)
        If Not rsTerms.EOF Then
                booData = True
                If rsTerms.Fields("PROC_SUBGROUP_ID") > 0 Then mlngProSubID = rsTerms.Fields("PROC_SUBGROUP_ID")
                If rsTerms.Fields("FIRST_COLUMN") <> vbNullString Then
                        strQuestion = vbNullString
                        strQuestion = Trim(rsTerms.Fields("FIRST_COLUMN"))
                        Format_CellTypeEdit grdColumn, SpreadHeader, 1, "1", strQuestion, "TypeHAlignCenter", strColItem
                        grdColumn.SetCellBorder 1, SpreadHeader, 1, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        .ColWidth(1) = 12
                End If
                If rsTerms.Fields("SECOND_COLUMN") <> vbNullString Then
                        strQuestion = vbNullString
                        If Trim(rsTerms.Fields("SECOND_COLUMN")) <> vbNullString Then
                            strQuestion = Trim(rsTerms.Fields("SECOND_COLUMN"))
                            Format_CellTypeEdit grdColumn, SpreadHeader, 2, "1", strQuestion, "TypeHAlignCenter", strColItem
                            grdColumn.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                        .ColWidth(2) = 12
                End If
                If rsTerms.Fields("THIRD_COLUMN") <> vbNullString Then
                        strQuestion = vbNullString
                        If Trim(rsTerms.Fields("THIRD_COLUMN")) <> vbNullString Then strQuestion = Trim(rsTerms.Fields("THIRD_COLUMN"))
                        Format_CellTypeEdit grdColumn, SpreadHeader, 3, "1", strQuestion, "TypeHAlignCenter", strColItem
                        grdColumn.SetCellBorder 3, SpreadHeader, 3, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        .ColWidth(3) = 12
                End If
                If rsTerms.Fields("FOURTH_COLUMN") <> vbNullString Then
                        strQuestion = vbNullString
                        If Trim(rsTerms.Fields("FOURTH_COLUMN")) <> vbNullString Then strQuestion = Trim(rsTerms.Fields("FOURTH_COLUMN"))
                        Format_CellTypeEdit grdColumn, SpreadHeader, 4, "1", strQuestion, "TypeHAlignCenter", strColItem
                        grdColumn.SetCellBorder 4, SpreadHeader, 4, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        .ColWidth(4) = 15
                End If
                If rsTerms.Fields("FIFTH_COLUMN") <> vbNullString Then
                        strQuestion = vbNullString
                        If Trim(rsTerms.Fields("FIFTH_COLUMN")) <> vbNullString Then strQuestion = Trim(rsTerms.Fields("FIFTH_COLUMN"))
                        Format_CellTypeEdit grdColumn, SpreadHeader, 5, "1", strQuestion, "TypeHAlignCenter", strColItem
                        grdColumn.SetCellBorder 5, SpreadHeader, 5, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        .ColWidth(5) = 12
                End If
                If rsTerms.Fields("SIXTH_COLUMN") <> vbNullString Then
                        strQuestion = vbNullString
                        If Trim(rsTerms.Fields("SIXTH_COLUMN")) <> vbNullString Then strQuestion = Trim(rsTerms.Fields("SIXTH_COLUMN"))
                        Format_CellTypeEdit grdColumn, SpreadHeader, 6, "1", strQuestion, "TypeHAlignCenter", strColItem
                        grdColumn.SetCellBorder 6, SpreadHeader, 6, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        .ColWidth(6) = 12
                End If
        End If
        If booData Then
             For iCol = 1 To .MaxCols
                .GetText iCol, SpreadHeader, varColName
                If Trim(varColName) = vbNullString Then
                    .Col = iCol
                    .ColHidden = True
                End If
            Next iCol
        End If
        .ReDraw = True
    End With
    
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Get_ChildTerms
    Exit Function
Error_Format_ColumnGrid:
    Write_Error Err.Description, Err.Number, strSQL, "Format_ColumnGrid", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Format_Procedures()
Dim R%
Dim i%
Dim Arr
Dim strProc As String
Dim ArrProc
Dim Idx As Integer

    With grdProc
        .ReDraw = False
        .ClearRange 1, SpreadHeader, .MaxCols, .MaxRows, False
        .ProcessTab = True
        .FontBold = True
        .FontName = "Times New Roman"
        .FontSize = 8
        .ColHeaderRows = 2
        .RowHeadersShow = False
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .PrintBorder = False
        .ColHeaderRows = 2
        .MaxCols = 2
        .TextTip = TextTipFloatingFocusOnly
        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        .FontBold = True
        .ColWidth(1) = 35
        .ColWidth(2) = 7
        .ScrollBars = ScrollBarsVertical
        .ScrollBarTrack = ScrollBarTrackVertical
        .MaxRows = 0
         Format_CellTypeEdit grdProc, SpreadHeader, 1, "1", "PROCEDURES", "TypeHAlignCenter", "PROCEDURES"
         Format_CellTypeEdit grdProc, SpreadHeader, 2, "1", "CPT", "TypeHAlignCenter", "CPT"
        Add_Spacer SpreadHeader + 1, grdProc
    End With
    
    If cboFilter.ListCount <= 0 Then
        cboFilter.AddItem "BODY"
        cboFilter.AddItem "BREAST"
        cboFilter.AddItem "FACE"
        cboFilter.AddItem "NOSE"
        cboFilter.AddItem "Other"
    End If
'    If mstrBodyFilter <> vbNullString Then
'        For i = 0 To cboFilter.ListCount - 1
'            If cboFilter.List(i) = Trim(mstrBodyFilter) Then
'                cboFilter.Text = cboFilter.List(i)
'                Exit For
'            End If
'        Next i
'    End If
    'GetTerms cboRedo, "sd_Proc_Redo"
    Exit Function
End Function

Private Function Get_Referrals(Optional vlngEncrID As Long)
Dim strSQL As String
Dim cnReferral As New ADODB.Connection
Dim rsReferral As New ADODB.Recordset
Dim lngReferralID As Long
Dim strRefReason As String
Dim Idx As Integer
Dim lngEnctrID As Long

    On Error GoTo Error_Get_Referrals
    If cnReferral.State = adStateClosed Then cnReferral.open psConnect(0)
    Screen.MousePointer = vbHourglass
    Clear_ReferralStrings
    cctRefReason.Text = vbNullString
    lstReferral.Clear
    If vlngEncrID > 0 Then lngEnctrID = vlngEncrID
    strSQL = "SELECT ENCOUNTER_ID, REFERRAL_ID, REFERRAL_REASON FROM SUPERBILL  WHERE  PATIENT_ID = " & piPatientID & "  ORDER BY INSERT_DATE DESC"
    Set rsReferral = cnReferral.Execute(strSQL)
    With rsReferral
        Do Until .EOF
            If lngEnctrID > 0 And (!ENCOUNTER_ID) = lngEnctrID Then
                If Not IsNull(!Referral_ID) Then
                    clsReferral.ReferralID = (!Referral_ID)
                    If Not IsNull(!REFERRAL_REASON) Then clsReferral.Consult_Reason = Trim((!REFERRAL_REASON))
                    cctRefReason.Text = Trim(clsReferral.Consult_Reason)
                    cctRefReason.Visible = True
                End If
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    If clsReferral.ReferralID > 0 Then
        Get_Referral clsReferral.ReferralID
        clsReferral.Consult_Reason = Trim(cctRefReason.Text)
        If Trim(clsReferral.Physician) <> vbNullString Then
            Load_Specific_Referral lngReferralID
            cctReferral.Text = clsReferral.Physician
        End If
    Else
        strSQL = "SELECT  REFERRER_ID, REFERRAL_REASON FROM PATIENT_ORDERS  WHERE  PATIENT_ID = " & piPatientID & " AND " _
            & "ENCOUNTER_ID = " & lngEnctrID
        Set rsReferral = cnReferral.Execute(strSQL)
        With rsReferral
            Do Until .EOF
                If Not IsNull(!REFERRER_ID) Then
                    clsReferral.ReferralID = (!REFERRER_ID)
                    If Not IsNull(!REFERRAL_REASON) Then clsReferral.Consult_Reason = Trim((!REFERRAL_REASON))
                    cctRefReason.Text = Trim(clsReferral.Consult_Reason)
                    Load_Specific_Referral clsReferral.ReferralID
                    cctRefReason.Text = Trim(clsReferral.Consult_Reason)
                    cctRefReason.Visible = True
                    Exit Do
                End If
            .MoveNext
            Loop
        End With
    End If
    If cnReferral.State = adStateOpen Then
        cnReferral.Close
        Set cnReferral = Nothing
        Set rsReferral = Nothing
    End If
    cctRefReason.Text = vbNullString
    If clsReferral.ReferralID = 0 Then
        Check_ApptReferral
        Get_Referral clsReferral.ReferralID
    End If
    If clsReferral.ReferralID > 0 Then
        lstReferral.Clear
        lstReferral.AddItem clsReferral.Physician
        lstReferral.ItemData(lstReferral.NewIndex) = clsReferral.ReferralID
        For Idx = 0 To lstReferral.ListCount - 1
            If lstReferral.ItemData(Idx) = clsReferral.ReferralID Then
                mbooByPass = True
                lstReferral.Selected(Idx) = True
                Exit For
            End If
        Next Idx
        cctReferral.Text = Trim(clsReferral.Physician)
        cctRefReason.Text = clsReferral.Consult_Reason
        cctRefReason.Visible = True
        Else
    End If
    
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
    
    Screen.MousePointer = vbHourglass
    If Trim(clsEncounter.StartDate) <> vbNullString Then
        strDate = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    Else
        Get_Encounter piEncounterID
        strDate = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    End If
    If cnRefer.State = adStateClosed Then cnRefer.open psConnect(0)
    If strDate = vbNullString Then strDate = Format(Now, "mm/dd/yyy")
    strSQL = "SELECT REFERRAL_ID,REFERRAL_REASON,APPT_DATE  FROM APPOINTMENTS WHERE " _
        & "PATIENT_ID = " & piPatientID & " ORDER BY APPT_DATE DESC"
    Set rsRefer = cnRefer.Execute(strSQL)
    With rsRefer
        Do Until .EOF
            If Not IsNull(!Referral_ID) Then lngRefID = (!Referral_ID)
            If (!Appt_Date) <> vbNullString Then
                strApptDate = (!Appt_Date)
                strApptDate = Format(strApptDate, "mm/dd/yyyy")
            End If
            If lngRefID > 0 Then
                lngDateDiff = DateDiff("d", strApptDate, strDate)
                If lngDateDiff = 0 Then
                    If Not IsNull(!Referral_ID) Then
                        If Not IsNull(!REFERRAL_REASON) Then
                            cctRefReason.Text = Trim(!REFERRAL_REASON)
                            clsReferral.Consult_Reason = (!REFERRAL_REASON)
                            cctRefReason.Visible = True
                        End If
                        clsReferral.ReferralID = lngRefID
                        Exit Do
                    End If
'                ElseIf lngRefID > 0 Then
'                    clsReferral.ReferralID = lngRefID
'                    If Not IsNull(!REFERRAL_REASON) Then
'                        cctRefReason.Text = Trim(!REFERRAL_REASON)
'                        clsReferral.Consult_Reason = (!REFERRAL_REASON)
'                        cctRefReason.Visible = True
'                    End If
'                    Exit Do
                End If
            End If
        .MoveNext
        Loop
    End With
    If clsReferral.ReferralID > 0 Then
        Load_Specific_Referral clsReferral.ReferralID
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

Private Function Create_Enctr_Default()
Dim i%
Dim lngFacID As Long
Dim lngProvID As Long


    For i = 0 To cboFacility.ListCount - 1
        If cboFacility.Text <> vbNullString Then
            If cboFacility.List(i) = cboFacility.Text Then
                lngFacID = cboFacility.ItemData(i)
                Exit For
            End If
        End If
    Next i
    For i = 0 To cboProvider.ListCount - 1
        If cboProvider.Text <> vbNullString Then
            If cboProvider.List(i) = cboProvider.Text Then
                lngProvID = cboProvider.ItemData(i)
                Exit For
            End If
        End If
    Next i
    If lngFacID > 1 Then SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultFacility", lngFacID, REG_SZ
    If lngFacID > 1 Then SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "FacilityID", lngFacID, REG_SZ
    If lngProvID > 1 Then SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultProvider", lngProvID, REG_SZ
    Exit Function
End Function



Private Function Get_EnctrInfo()
Dim strSQL As String
Dim lstrRemarks As String
Dim strAdmission As String
Dim strDischarge As String
Dim lngrecs As Long
Dim strStartDate As String
Dim varAdmisStartDate
Dim varDisStartDate
Dim booInsertAdmission As Boolean
Dim booInsertDisc As Boolean
Dim varTime
Dim strCurrentDate As String
Dim lngdelta As Long
Dim i%

    Get_Providers
    Get_ApptData
    Get_Facilities
    If clsFacility.FacType <> vbNullString Then clsEncounter.ServicePlace = clsFacility.FacType
    If cboSerPlace.Text = vbNullString Then
        For i = 0 To cboSerPlace.ListCount - 1
           If cboSerPlace.List(i) = clsEncounter.ServicePlace Then
                clsEncounter.ServicePlaceCode = cboSerPlace.ItemData(i)
                cboSerPlace.Text = cboSerPlace.List(i)
                Exit For
           End If
        Next i
    Else
        clsEncounter.ServicePlace = cboSerPlace.Text
    End If

    If cboStatus.Text <> vbNullString Then
        clsEncounter.EncounterType = Trim(cboStatus.Text)
    ElseIf cboStatus.Text = vbNullString And clsEncounter.EncounterType <> vbNullString Then
        For i = 0 To cboStatus.ListCount - 1
            If cboStatus.List(i) = clsEncounter.EncounterType Then
                cboStatus.Text = cboStatus.List(i)
                Exit For
            End If
        Next i
    ElseIf cboStatus.Text = vbNullString And clsEncounter.EncounterType = vbNullString Then
        MsgBox "Please select and update Patient Status", vbInformation
    End If
    If clsEncounter.VisitType <> vbNullString Then
        For i = 0 To cboPurVisit.ListCount - 1
            If cboPurVisit.List(i) = clsEncounter.VisitType Then
                cboPurVisit.Text = cboPurVisit.List(i)
                Exit For
            End If
        Next i
    End If
    Exit Function
End Function
Public Function Create_NewEncounter(bInsert As Boolean)
Dim strSQL As String
Dim strRS
Dim lstrRemarks As String
Dim strAdmission As String
Dim strDischarge As String
Dim intProviderID As Integer
Dim intFacilityID As Integer
Dim varVisitReason
Dim lngrecs As Long
Dim varStartDate
Dim strEnctrType
Dim varAdmisStartDate
Dim lReturn As Integer
Dim varDisStartDate
Dim varServicePlace
Dim i%
Dim booInsertAdmission As Boolean
Dim booInsertDisc As Boolean
Dim varTime
Dim varCurrentDate
Dim vartemp

    On Error GoTo Error_Create_NewEncounter
    If bInsert <> True Then Exit Function
    Screen.MousePointer = vbHourglass
    Validate_Enctr
    If mbooAddNewEnctr = False Then
        MsgBox "No new encounter created for" & " " & Trim(clsPatient.Name) & " " & mstrDenyEncntrReas, vbCritical
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        WriteData ("I")
        modSpecific.piEncounterID = piEncounterID
        Init_Summary
        Set_CurrentEncounter
    End If
    Screen.MousePointer = vbDefault
    mbooAddNewEnctr = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Create_NewEncounter:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Create_NewEncounter", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function



Private Function Manage_Admission(vlngEncounterID As Long)
Dim cnAdmit As New ADODB.Connection
Dim rsAdmit As New ADODB.Recordset
Dim sSQL As String
Dim cnUpdate As New ADODB.Connection
Dim rsUpdate As New ADODB.Recordset
Dim sDisDate As String
Dim sAdmitDate As String
Dim strEnctrDate As String
Dim lngDateDiff As String
Dim sDate As String
Dim lngAdmitID As Long
Dim Arr
Dim strEnctr As String
Dim strSQL As String
Dim lngEnctrID As Long
Dim lngUpdateID As Long

    On Error GoTo Error_Manage_Admission
    Screen.MousePointer = vbHourglass
    lngEnctrID = vlngEncounterID
    Get_EDI_Static
    If clsEDI_Client.Site_ID = "1125" Then
        If cnAdmit.State = adStateClosed Then cnAdmit.open psConnect(0)
        strSQL = "SELECT * FROM PATIENT_ADMISSION WHERE PATIENT_ID = " & piPatientID
        Set rsAdmit = cnAdmit.Execute(strSQL)
        With rsAdmit
            If Not .EOF Then
                If Not IsNull(!ADMISSION_DATE) Then
                    If IsDate(!ADMISSION_DATE) Then
                        clsInPatRecord.AdmissionDate = (!ADMISSION_DATE)
                    End If
                End If
            End If
        End With
        If clsInPatRecord.AdmissionDate = vbNullString Then
            strSQL = "SELECT * FROM ENCOUNTER WHERE PATIENT_ID = " & piPatientID & "  AND PURGE = 0  ORDER BY ENCNTR_START_DATE ASC"
            Set rsAdmit = cnAdmit.Execute(strSQL)
            With rsAdmit
                If Not .EOF Then
                    If Not IsNull(!ENCNTR_START_DATE) Then
                        If IsDate(!ENCNTR_START_DATE) Then
                            clsInPatRecord.AdmissionDate = (!ENCNTR_START_DATE)
                        End If
                    End If
                End If
            End With
        End If
        If Trim(clsInPatRecord.AdmissionDate) <> vbNullString Then
            ccdAdmit.Enabled = True
            ccdAdmit.CheckBox = True
            If IsNull(ccdAdmit.Value) Then
                ccdAdmit.Value = Format(clsInPatRecord.AdmissionDate, "mm/dd/yyyy")
            End If
            If Not IsNull(ccdAdmit.Value) Then sAdmitDate = Format(ccdAdmit.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
            If Not IsNull(ccdDisharge.Value) Then sDisDate = Format(ccdDisharge.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
            strSQL = "SELECT * FROM PATIENT_ADMISSION  WHERE PATIENT_ID = " & piPatientID & " AND ENCOUNTER_ID = " & lngEnctrID
            Set rsAdmit = cnAdmit.Execute(strSQL)
            With rsAdmit
                If Not .EOF Then
                    lngAdmitID = (!PATIENT_ADMITID)
                End If
            End With
            If cnAdmit.State = adStateClosed Then cnAdmit.open psConnect(0)
            If lngAdmitID = 0 Then
                strSQL = "INSERT INTO PATIENT_ADMISSION (PATIENT_ID, ENCOUNTER_ID, ADMISSION_DATE, DISCHARGE_DATE, INSERT_DATE, PURGE) " _
                    & "values(" & piPatientID & "," & lngEnctrID & ",'" & sAdmitDate & "','" & sDisDate & "','" _
                    & sDate & "',0)"
                Set rsAdmit = cnAdmit.Execute(strSQL)
            Else
                strSQL = "UPDATE PATIENT_ADMISSION SET ADMISSION_DATE = '" & sAdmitDate & "',DISCHARGE_DATE = '" & sDisDate & "'," _
                    & "UPDATE_DATE = '" & sDate & "' where PATIENT_ADMITID = " & lngAdmitID
                Set rsAdmit = cnAdmit.Execute(strSQL)
            End If
            strSQL = "UPDATE ENCOUNTER SET ADMISSION_DATE = '" & sAdmitDate & "' WHERE ENCOUNTER_ID = " & lngEnctrID
            Set rsAdmit = cnAdmit.Execute(strSQL)
            If sDisDate <> vbNullString Then
                strSQL = "UPDATE ENCOUNTER SET DISCHARGE_DATE = '" & sDisDate & "' WHERE ENCOUNTER_ID = " & lngEnctrID
                Set rsAdmit = cnAdmit.Execute(strSQL)
            End If
        End If
    Else
        If IsNull(ccdAdmit.Value) Then
            ccdAdmit.Enabled = True
            ccdAdmit.CheckBox = True
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        If clsEncounter.ServicePlaceCode <> "11" Then
            If cnAdmit.State = adStateClosed Then cnAdmit.open psConnect(0)
            strEnctrDate = Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM")
            sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
            If Not IsNull(ccdAdmit.Value) Then sAdmitDate = Format(ccdAdmit.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
            If Not IsNull(ccdDisharge.Value) Then sDisDate = Format(ccdDisharge.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
            strSQL = "SELECT * FROM PATIENT_ADMISSION  WHERE PATIENT_ID = " & piPatientID & " AND ENCOUNTER_ID = " & lngEnctrID
            Set rsAdmit = cnAdmit.Execute(strSQL)
            With rsAdmit
                If Not .EOF Then
                    lngAdmitID = (!PATIENT_ADMITID)
                    clsInPatRecord.AdmissionDate = (!ADMISSION_DATE)
                End If
            End With
            If cnAdmit.State = adStateClosed Then cnAdmit.open psConnect(0)
            If lngAdmitID = 0 Then
                strSQL = "INSERT INTO PATIENT_ADMISSION (PATIENT_ID, ENCOUNTER_ID, ADMISSION_DATE, DISCHARGE_DATE, INSERT_DATE, PURGE) " _
                    & "values(" & piPatientID & "," & lngEnctrID & ",'" & sAdmitDate & "','" & sDisDate & "','" _
                    & sDate & "',0)"
                Set rsAdmit = cnAdmit.Execute(strSQL)
            Else
                strSQL = "UPDATE PATIENT_ADMISSION SET ADMISSION_DATE = '" & sAdmitDate & "',DISCHARGE_DATE = '" & sDisDate & "'," _
                    & "UPDATE_DATE = '" & sDate & "' where PATIENT_ADMITID = " & lngAdmitID
                Set rsAdmit = cnAdmit.Execute(strSQL)
            End If
            strSQL = "UPDATE ENCOUNTER SET ADMISSION_DATE = '" & sAdmitDate & "' WHERE ENCOUNTER_ID = " & lngEnctrID
            Set rsAdmit = cnAdmit.Execute(strSQL)
            If sDisDate <> vbNullString Then
                strSQL = "UPDATE ENCOUNTER SET DISCHARGE_DATE = '" & sDisDate & "' WHERE ENCOUNTER_ID = " & lngEnctrID
                Set rsAdmit = cnAdmit.Execute(strSQL)
            End If
        End If
    End If
    ccdAdmit.Enabled = True
    ccdAdmit.CheckBox = True
    If cnAdmit.State = adStateOpen Then
        cnAdmit.Close
        Set cnAdmit = Nothing
        Set rsAdmit = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Manage_Admission:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Manage_Admission", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Provider_HasOpenEncounter() As Boolean
Dim strSQL As String
Dim lintProviderID As Integer
Dim lstrRemarks As String
Dim strStartDate As String
Dim strEndDate As String
Dim cnEnctr As New ADODB.Connection
Dim rsEnctr As New ADODB.Recordset
Dim strNow As String

    On Error GoTo Error_Provider_HasOpenEncounter
    Screen.MousePointer = vbHourglass
    strNow = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    strStartDate = Format(Now, "mm/dd/yyyy 00:00:000")
    strEndDate = Format(Now, "mm/dd/yyyy 23:59:001")
    If cnEnctr.State = adStateClosed Then cnEnctr.open psConnect(0)
    strSQL = "Select ENCOUNTER_ID, ENCNTR_END_DATE, VISIT_REASON FROM ENCOUNTER WHERE PATIENT_ID = " & piPatientID & " AND " _
        & "PROVIDER_ID = " & piProviderID & " AND ENCNTR_START_DATE " _
        & "BETWEEN '" & strStartDate & "' AND '" & strEndDate & "'"
    Set rsEnctr = cnEnctr.Execute(strSQL)
        Do Until rsEnctr.EOF
        If Not IsNull(rsEnctr.Fields("VISIT_REASON")) Then
            If Trim(rsEnctr.Fields("VISIT_REASON")) <> vbNullString Then
                If Not TextFound("Telephone Visit", rsEnctr.Fields("VISIT_REASON"), Len("Telephone Visit"), True) Then
                    If IsNull(rsEnctr.Fields("ENCNTR_END_DATE")) Then
                        Provider_HasOpenEncounter = True
                        Exit Do
                    Else
                        Provider_HasOpenEncounter = False
                    End If
                End If
            End If
        Else
            Provider_HasOpenEncounter = False
        End If
        rsEnctr.MoveNext
    Loop
    If cnEnctr.State = adStateOpen Then
        cnEnctr.Close
        Set cnEnctr = Nothing
        Set rsEnctr = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Provider_HasOpenEncounter:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Provider_HasOpenEncounter", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub CheckForChanges()
Dim iResp As Byte
Dim iAnswer As Integer

    If mbooNewEncntr Then
        iAnswer = MsgBox("Do you want to save the new encounter?", vbYesNo + vbQuestion)
        If iAnswer = vbYes Then
            cmdSave_Click
        Else
            Exit Sub
        End If
    End If
    mbooNewEncntr = False
    
    If fbooEndClicked Then
        iAnswer = MsgBox("Do you still want to close this Encounter?", vbYesNo + vbQuestion)
        If iAnswer = vbYes Then
            cmdEndEncounter_Click
        Else
            Exit Sub
        End If
    End If
    Exit Sub
End Sub


Private Function Save_SurgLog()
Dim cnWrite As New ADODB.Connection
Dim rsWrite As New ADODB.Recordset
Dim strSQL As String
Dim booinsert As Boolean
Dim strSurDate As String
Dim strEnctrDate As String
Dim lngDateDiff As Long
Dim lngSurgID As Long
Dim strSurgDate As String
Dim lngMaxID As Long
Dim lngPrimKey As Long
Dim lngFacID As Long


    If mstrPractice <> "Plastic Surgery" Then Exit Function
    If cnWrite.State = adStateClosed Then cnWrite.open psConnect(0)
    strEnctrDate = Format(ccdStartDate.Value, "mm/dd/yyyy")
    clsInPatRecord.SurgicalID = 0
    If mstrPractice = "Plastic Surgery" Then
            strSQL = "SELECT SURGERY_ID FROM PATIENT_SURGERY_ID WHERE PATIENT_ID = " & piPatientID
            Set rsWrite = cnWrite.Execute(strSQL)
            If Not rsWrite.EOF Then
                clsInPatRecord.SurgicalID = rsWrite.Fields("SURGERY_ID")
            End If
    End If
    'Change Surgery Id
    udtFacility.FacID = 0
    Get_SurgFacility
    If udtFacility.FacID > 0 Then
        piFacilityID = udtFacility.FacID
    End If
    If clsInPatRecord.SurgicalID = 0 Then
        strSQL = "INSERT INTO PATIENT_SURGERY_ID (PATIENT_ID, FACILITY_ID,INSERT_DATE) values" _
            & "(" & piPatientID & ", " & piFacilityID & ",'" & strEnctrDate & "')"
        Set rsWrite = cnWrite.Execute(strSQL)
    Else
        strSQL = "UPDATE PATIENT_SURGERY_ID SET  FACILITY_ID = " & piFacilityID & ", UPDATE_DATE = '" & strEnctrDate & "' WHERE " _
            & "SURGERY_ID = " & Val(clsInPatRecord.SurgicalID)
        Set rsWrite = cnWrite.Execute(strSQL)
    End If
    strSQL = "SELECT SURGLOG_ID, SURG_DATE FROM SURGERY_LOG WHERE  PATIENT_ID = " & piPatientID & " AND " _
        & "ENCOUNTER_ID = " & piEncounterID & " ORDER BY SURG_DATE DESC"
        Set rsWrite = cnWrite.Execute(strSQL)
        Do Until rsWrite.EOF
            If rsWrite.Fields("SURG_DATE") <> vbNullString Then strSurgDate = rsWrite.Fields("SURG_DATE")
            If strSurDate <> vbNullString Then
                strSurgDate = Format(strSurgDate, "mm/dd/yyyy")
                lngDateDiff = DateDiff("d", strEnctrDate, strSurgDate)
                If lngDateDiff = 0 Then
                    booinsert = False
                    lngSurgID = rsWrite.Fields("SURGLOG_ID")
                    Exit Do
                Else
                    booinsert = True
                End If
            End If
            rsWrite.MoveNext
        Loop
    strEnctrDate = Format(ccdStartDate.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
    If clsEncounter.Intrvn_CD <> vbNullString Then
        clsOpNote.Preop_Code = Replace(clsEncounter.Intrvn_CD, "049X" & "-", "")
    Else
        clsOpNote.Preop_Code = "049X"
    End If
    If clsOpNote.Procedure_Name = vbNullString Then clsOpNote.Procedure_Name = clsEncounter.VisitReason
    If booinsert Then
        strSQL = "SELECT MAX(SURG_ID) FROM SURGERY_LOG"
        Set rsWrite = cn.Write(strSQL)
        If Not rsWrite.EOF Then
            lngMaxID = rsWrite.Fields(0)
        End If
        If lngMaxID <> 0 And lngMaxID > 5000 Then
            lngMaxID = lngMaxID + 1
        Else
            lngMaxID = 5000
        End If
        
        strSQL = "INSERT INTO SURGERY_LOG (SURGLOG_ID, ENCOUNTER_ID, PATIENT_ID, SURG_DATE, SURGERY_ID, SURG_PROCEDURE, CPT_CODE, PROVIDER_ID, INSERT_DATE) " _
            & "values (" & lngMaxID & "," & piEncounterID & ", " & piPatientID & ",'" & strEnctrDate & "'," & Val(clsInPatRecord.SurgicalID) & ",'" & clsOpNote.Procedure_Name & "','" & clsOpNote.Preop_Code & "'," _
            & piProviderID & ",'" & strEnctrDate & "')"
         Set rsWrite = cnWrite.Execute(strSQL)
    Else
        strSQL = "UPDATE SURGERY_LOG Set SURG_DATE = '" & strEnctrDate & "', ENCOUNTER_ID = " & piEncounterID & ", SURG_PROCEDURE = '" & clsOpNote.Procedure_Name & "'," _
            & "PROVIDER_ID = " & piProviderID & ", CPT_CODE = '" & clsOpNote.Preop_Code & "', UPDATE_DATE = '" & strEnctrDate & "' WHERE " _
            & "SURGLOG_ID = " & lngSurgID
         Set rsWrite = cnWrite.Execute(strSQL)
    End If
    If clsEncounter.StartDate <> vbNullString Then clsEncounter.StartDate = Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM")

    If cnWrite.State = adStateClosed Then cnWrite.open psConnect(0)
    strSQL = "SELECT OPNOTE_ID FROM PATIENT_OPERATIVE_NOTE WHERE PATIENT_ID = " & piPatientID & " AND " _
            & "ENCOUNTER_ID = " & piEncounterID
    Set rsWrite = cnWrite.Execute(strSQL)
    With rsWrite
        If Not rsWrite.EOF Then
            booinsert = False
            lngPrimKey = (!Opnote_Id)
        End If
    End With
    If lngPrimKey = 0 Then
        strSQL = "INSERT INTO PATIENT_OPERATIVE_NOTE (PATIENT_ID, ENCOUNTER_ID, SURGERY_DATE) " _
            & "values(" & piPatientID & ",'" & piEncounterID & "','" & clsEncounter.StartDate & "')"
            Set rsWrite = cnWrite.Execute(strSQL)
    End If
    If cnWrite.State = adStateOpen Then
        cnWrite.Close
        Set cnWrite = Nothing
        Set rsWrite = Nothing
    End If
    Exit Function
ERROR:
    MsgBox "Problem encountered saving Surgery Log!", vbInformation
    Exit Function
    Resume
End Function

Function Set_Default_Enctr()
Dim i%
Dim sReg As String
Dim strValue As String

    If clsEDI_Client.Site_ID = "1125" Then
        If piFacilityID > 0 Then
            For i = 0 To cboFacility.ListCount - 1
                If cboFacility.ItemData(i) = piFacilityID Then
                    cboFacility.Text = cboFacility.List(i)
                    Exit For
                End If
            Next i
        End If
        Get_NurseReferral
        Manage_Admission piEncounterID
        Get_NurseDiagnosis
    Else
        If cboFacility.Text = vbNullString Then
            psRegGeneral = "SOFTWARE\FirstNoteBook\PatientTrac\General"
            sReg = "DefaultFacility"
            strValue = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, sReg)
            If strValue <> vbNullString Then
                For i = 0 To cboFacility.ListCount - 1
                    If cboFacility.ItemData(i) = Val(strValue) Then
                        cboFacility.Text = cboFacility.List(i)
                        Exit For
                    End If
                Next i
            End If
        End If
    End If

    If cboProvider.Text = vbNullString Then
            psRegGeneral = "SOFTWARE\FirstNoteBook\PatientTrac\General"
            sReg = "DefaultProvider"
            mlngDefProviderID = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, sReg)
            If mlngDefProviderID > 1 Then
                For i = 0 To cboProvider.ListCount - 1
                    If cboProvider.ItemData(i) = mlngDefProviderID Then
                        cboProvider.Text = cboProvider.List(i)
                        Exit For
                    End If
                Next i
            End If
    End If
    Exit Function
End Function
Private Function Get_NurseDiagnosis()
Dim cnPlan As New ADODB.Connection
Dim rsPlan As New ADODB.Recordset
Dim strSQL As String
Dim lngRow As Long
Dim Arr
Dim strCode As String
Dim strDiag As String
Dim ArrDiag
Dim booSkip As Boolean
Dim lngProcRow As Long
Dim Idx As Integer
Dim strAddProcCode As String
Dim strAddProc As String
Dim i%
Dim strCstrRow As String
Dim lngCol As Long
Dim lngFee As Long
Dim strList As String
Dim varCellValue
Dim strProcCode As String
Dim strCodeDesc As String
Dim ArrCode
Dim lngChargeRow As Long

    On Error GoTo Error_Get_NurseDiagnosis
    Screen.MousePointer = vbHourglass
    If cnPlan.State = adStateClosed Then cnPlan.open psConnect(0)
    strSQL = "SELECT a.DIAGNOSIS_CODE FROM ENCOUNTER_SERVICE a, ENCOUNTER b WHERE a.EDI_BILL_ID = b.ENCOUNTER_ID " _
        & "AND ASSIGNED_NUMBER = '1' AND b.PATIENT_ID = " & piPatientID & " ORDER BY b.INSERT_DATE DESC"
    Set rsPlan = cnPlan.Execute(strSQL)
    With rsPlan
        Do Until .EOF
            If Not IsNull((!DIAGNOSIS_CODE)) Then
                If Trim((!DIAGNOSIS_CODE)) <> vbNullString Then
                    cctDiagnosis.Text = (!DIAGNOSIS_CODE)
                    Exit Do
                End If
            End If
            .MoveNext
        Loop
    End With
    If Trim(cctDiagnosis.Text) <> vbNullString Then
        cctDiagnosis.Text = Look_UpDiagName(Trim(cctDiagnosis.Text))
    End If
    If cnPlan.State = adStateOpen Then
        cnPlan.Close
        Set cnPlan = Nothing
        Set rsPlan = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_NurseDiagnosis:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_NurseDiagnosis", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Look_UpDiagName(vstrType As String) As String
Dim strSQL As String
Dim strName As String
Dim strType As String
Dim rsDiag As New ADODB.Recordset
Dim cnDiag As New ADODB.Connection
Dim strCode As String
Dim Arr
Dim Idx%

    On Error GoTo Error_Look_UpDiagName
    Screen.MousePointer = vbHourglass
    Look_UpDiagName = vbNullString
    If cnDiag.State = adStateClosed Then cnDiag.open psConnect(1)
    If Trim(vstrType) <> vbNullString Then
        Arr = Split(vstrType, "|")
        If UBound(Arr) > 0 Then
            For Idx = LBound(Arr) To UBound(Arr)
                strCode = vbNullString
                strName = vbNullString
                If Len(Arr(Idx)) > 2 Then strCode = Trim(Arr(Idx))
                If Trim(strCode) <> vbNullString Then
                    strSQL = "select Name, Number, Chapter from  [DSM-IV] Where Number like '" & strCode & "%'"
                    Set rsDiag = cnDiag.Execute(strSQL)
                    With rsDiag
                       If Not .EOF Then
                            If Not IsNull((!Number)) Then
                                If (!Number) <> vbNullString Then
                                    If Not IsNull((!Name)) Then
                                        strName = (!Name)
                                        strName = Replace(strName, "''", "'")
                                        If Trim(strCode) Like "[A-Za-z]###" Or Trim(strCode) Like "###" Then
                                            If (!Number) <> vbNullString Then
                                                If Trim(Look_UpDiagName) = vbNullString Then
                                                    Look_UpDiagName = "AXIS I" & " " & strName & "-" & Format((!Number), "000.00")
                                                ElseIf Not TextFound((!Number), Look_UpDiagName, Len((!Number)), True) Then
                                                    Look_UpDiagName = Look_UpDiagName & vbNewLine & strName & "-" & Format((!Number), "000.00")
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End With
                End If
            Next Idx
        Else
            strCode = Trim(vstrType)
            strSQL = "select Name, Number, Chapter from  [DSM-IV] Where Number like '" & strCode & "%'"
            Set rsDiag = cnDiag.Execute(strSQL)
            With rsDiag
                If Not .EOF Then
                    If Not IsNull((!Number)) Then
                        If (!Number) <> vbNullString Then
                            If (!Number) <> vbNullString Then
                                strName = (!Name)
                                Look_UpDiagName = "AXIS I" & " " & strName & "-" & Format((!Number), "000.00")
                            End If
                        End If
                    End If
                End If
            End With
        End If
    End If
    If Look_UpDiagName = vbNullString Then
    strSQL = "select long_desc, code from icd_code Where code like '" & strCode & "%' ORDER BY code"
        Set rsDiag = cnDiag.Execute(strSQL)
        With rsDiag
        Do Until .EOF
                If Not IsNull(rsDiag.Fields("long_desc")) Then
                    strName = rsDiag.Fields("long_desc")
                End If
                rsDiag.MoveNext
                If Trim(strName) <> vbNullString Then
                    Look_UpDiagName = "AXIS I" & " " & strName & "-" & Format(strCode, "000.00")
                    Exit Do
                End If
            Loop
        End With
    End If
    If cnDiag.State = adStateOpen Then
        cnDiag.Close
        Set cnDiag = Nothing
        Set rsDiag = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Look_UpDiagName:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Look_UpDiagName", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Private Function Get_NurseReferral()
Dim strSQL As String
Dim cnReferral As New ADODB.Connection
Dim rsReferral As New ADODB.Recordset
Dim lngReferralID As Long
Dim strRefReason As String
Dim Idx As Integer
Dim lngEnctrID As Long

    On Error GoTo Error_Get_NurseReferral
    If cnReferral.State = adStateClosed Then cnReferral.open psConnect(0)
    Screen.MousePointer = vbHourglass
    Clear_ReferralStrings
    cctRefReason.Text = vbNullString
    lstReferral.Clear
    strSQL = "SELECT ENCOUNTER_ID, REFERRAL_ID, REFERRAL_REASON FROM SUPERBILL  WHERE  " _
        & "PATIENT_ID = " & piPatientID & "  ORDER BY INSERT_DATE DESC"
    Set rsReferral = cnReferral.Execute(strSQL)
    With rsReferral
        Do Until .EOF
            If Not IsNull(!Referral_ID) Then
                clsReferral.ReferralID = (!Referral_ID)
                If Not IsNull(!REFERRAL_REASON) Then clsReferral.Consult_Reason = Trim((!REFERRAL_REASON))
                cctRefReason.Text = Trim(clsReferral.Consult_Reason)
                cctRefReason.Visible = True
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    If clsReferral.ReferralID > 0 Then
        Get_Referral clsReferral.ReferralID
        clsReferral.Consult_Reason = Trim(cctRefReason.Text)
        If Trim(clsReferral.Physician) <> vbNullString Then
            Load_Specific_Referral clsReferral.ReferralID
            cctReferral.Text = clsReferral.Physician
        End If
    Else
        strSQL = "SELECT  REFERRER_ID, REFERRAL_REASON FROM PATIENT_ORDERS  WHERE " _
            & "PATIENT_ID = " & mlngPatientID
        Set rsReferral = cnReferral.Execute(strSQL)
        With rsReferral
            Do Until .EOF
                If Not IsNull(!REFERRER_ID) Then
                    clsReferral.ReferralID = (!REFERRER_ID)
                    If Not IsNull(!REFERRAL_REASON) Then clsReferral.Consult_Reason = Trim((!REFERRAL_REASON))
                    cctRefReason.Text = Trim(clsReferral.Consult_Reason)
                    Load_Specific_Referral clsReferral.ReferralID
                    cctRefReason.Text = Trim(clsReferral.Consult_Reason)
                    cctRefReason.Visible = True
                    Exit Do
                End If
            .MoveNext
            Loop
        End With
    End If
    If cnReferral.State = adStateOpen Then
        cnReferral.Close
        Set cnReferral = Nothing
        Set rsReferral = Nothing
    End If
    cctRefReason.Text = vbNullString
    If clsReferral.ReferralID = 0 Then
        Check_ApptReferral
        Get_Referral clsReferral.ReferralID
    End If
    If clsReferral.ReferralID > 0 Then
        lstReferral.Clear
        lstReferral.AddItem clsReferral.Physician
        lstReferral.ItemData(lstReferral.NewIndex) = clsReferral.ReferralID
        For Idx = 0 To lstReferral.ListCount - 1
            If lstReferral.ItemData(Idx) = clsReferral.ReferralID Then
                mbooByPass = True
                lstReferral.Selected(Idx) = True
                Exit For
            End If
        Next Idx
        cctReferral.Text = Trim(clsReferral.Physician)
        cctRefReason.Text = clsReferral.Consult_Reason
        cctRefReason.Visible = True
        Else
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_NurseReferral:
    Screen.MousePointer = vbDefault
     Write_Error Err.Description, Err.Number, strSQL, "Get_NurseReferral", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Private Function Set_CurrentEncounter()
Dim strSQL As String
Dim lngRow As Long
Dim varEnctrID
Dim lngEnctrID As Long

    If piEncounterID = 0 Then Exit Function
    With grdEnctrSummary
        For lngRow = 1 To grdEnctrSummary.MaxRows
        If lngRow > 0 Then
          .GetText 0, lngRow, varEnctrID
            If varEnctrID <> vbNullString Then
                If Val(varEnctrID) = piEncounterID Then
                grdEnctrSummary_DblClick 1, lngRow
                Exit For
                End If
            End If
        End If
        Next lngRow
    End With
    If clsEDI_Client.Site_ID = "1125" Then
        If Trim(cctDiagnosis.Text) = vbNullString Or Trim(cctReferral.Text) = vbNullString Or IsNull(ccdAdmit.Value) Then
            Get_NurseReferral
            Manage_Admission piEncounterID
            Get_NurseDiagnosis
        End If
    End If
    Exit Function
End Function


Private Function Update_ReferPhys(vlngRefID As Long)
Dim strSQL As String
Dim cnReferral As New ADODB.Connection
Dim rsReferral As New ADODB.Recordset
Dim Idx%
Dim strDate As String
Dim lngPrimKey As Long
Dim lngRefID As Long


    On Error GoTo Error_Update_ReferPhys
    If cnReferral.State = adStateClosed Then cnReferral.open psConnect(0)
    clsReferral.ReferralID = 0
    lngRefID = vlngRefID
    Clear_ReferralStrings
    If clsEncounter.StartDate = vbNullString Then Get_Encounter piEncounterID
    clsReferral.Consult_Reason = Trim(cctRefReason.Text)
    strDate = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    Screen.MousePointer = vbHourglass
    If lngRefID = 0 Then
        For Idx = 0 To lstReferral.ListCount - 1
            If lstReferral.Selected(Idx) Then
                lngRefID = lstReferral.ItemData(Idx)
                Exit For
            End If
        Next Idx
    End If
    If Trim(cctRefReason.Text) <> vbNullString Then
        clsReferral.Consult_Reason = Trim(cctRefReason.Text)
        clsReferral.Consult_Reason = Replace(clsReferral.Consult_Reason, "'", "''")
    End If
    If lngRefID > 0 Then clsReferral.ReferralID = lngRefID
    If clsReferral.ReferralID > 0 And Trim(clsReferral.Consult_Reason) = vbNullString Then
        strSQL = "SELECT REFERRAL_REASON FROM APPOINTMENTS WHERE PATIENT_ID = " & piPatientID & " AND APPT_DATE = '" & strDate & "'"
        Set rsReferral = cnReferral.Execute(strSQL)
        If Not rsReferral.EOF Then
            If Not IsNull(rsReferral.Fields("REFERRAL_REASON")) Then
                clsReferral.Consult_Reason = rsReferral.Fields("REFERRAL_REASON")
            End If
        End If
    ElseIf clsReferral.ReferralID > 0 Then
        strSQL = "UPDATE APPOINTMENTS SET REFERRAL_ID = " & clsReferral.ReferralID & "," _
            & "REFERRAL_REASON = '" & clsReferral.Consult_Reason & "' WHERE PATIENT_ID = " & piPatientID & " and " _
            & "APPT_DATE = '" & strDate & "'"
            Set rsReferral = cnReferral.Execute(strSQL)
        strSQL = "SELECT ACCT_ID, REFERRAL_REASON FROM PATIENT_ACCOUNTS WHERE " _
            & "PATIENT_TYPE = 'Patient' AND PATIENT_ID = " & piPatientID
            Set rsReferral = cnReferral.Execute(strSQL)
            If Not rsReferral.EOF Then
                lngPrimKey = rsReferral.Fields("ACCT_ID")
                If Not IsNull(rsReferral.Fields("REFERRAL_REASON")) Then
                    If Trim(rsReferral.Fields("REFERRAL_REASON")) <> vbNullString And clsReferral.Consult_Reason = vbNullString Then clsReferral.Consult_Reason = rsReferral.Fields("REFERRAL_REASON")
                End If
            End If
        If lngPrimKey = 0 Then
             strSQL = "INSERT INTO PATIENT_ACCOUNTS (PATIENT_TYPE, PATIENT_ID, REFERRAL_ID, REFERRAL_REASON, ACCOUNT_NO) values " _
                & "('Patient', " & piPatientID & "," & clsReferral.ReferralID & ",'" & clsReferral.Consult_Reason & "'," & piPatientID & ")"
                Set rsReferral = cnReferral.Execute(strSQL)
        Else
            strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
            strSQL = "UPDATE PATIENT_ACCOUNTS SET REFERRAL_ID = " & clsReferral.ReferralID & "," _
                & "REFERRAL_REASON = '" & clsReferral.Consult_Reason & "', UPDATE_DATE = '" & strDate & "' WHERE ACCT_ID = " & lngPrimKey
                Set rsReferral = cnReferral.Execute(strSQL)
        End If
        
        lngPrimKey = 0
         strSQL = "SELECT ORDER_ID, REFERRAL_REASON FROM PATIENT_ORDERS WHERE PATIENT_ID = " & piPatientID & " AND ENCOUNTER_ID = " & piEncounterID
            Set rsReferral = cnReferral.Execute(strSQL)
            If Not rsReferral.EOF Then
                lngPrimKey = rsReferral.Fields("ORDER_ID")
                If Not IsNull(rsReferral.Fields("REFERRAL_REASON")) Then
                    If Trim(rsReferral.Fields("REFERRAL_REASON")) <> vbNullString And Trim(clsReferral.Consult_Reason) = vbNullString Then
                        clsReferral.Consult_Reason = rsReferral.Fields("REFERRAL_REASON")
                    End If
                End If
            End If
        If lngPrimKey > 0 Then
              strSQL = "INSERT INTO PATIENT_ORDERS (PATIENT_ID, ENCOUNTER_ID, REFERRER_ID, REFERRAL_REASON) values " _
                & "(" & piPatientID & "," & piEncounterID & "," & clsReferral.ReferralID & ",'" & clsReferral.Consult_Reason & "')"
                Set rsReferral = cnReferral.Execute(strSQL)
        Else
            strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
            strSQL = "UPDATE PATIENT_ORDERS SET REFERRER_ID = " & clsReferral.ReferralID & "," _
                & "REFERRAL_REASON = '" & clsReferral.Consult_Reason & "', UPDATE_DATE = '" & strDate & "'  WHERE ORDER_ID = " & lngPrimKey
                Set rsReferral = cnReferral.Execute(strSQL)
        End If
        lngPrimKey = 0
         strSQL = "SELECT BILLING_ID, REFERRAL_REASON  FROM SUPERBILL WHERE PATIENT_ID = " & piPatientID & " AND ENCOUNTER_ID = " & piEncounterID
            Set rsReferral = cnReferral.Execute(strSQL)
            If Not rsReferral.EOF Then
                lngPrimKey = rsReferral.Fields("BILLING_ID")
                If Not IsNull(rsReferral.Fields("REFERRAL_REASON")) Then
                    If Trim(rsReferral.Fields("REFERRAL_REASON")) <> vbNullString And clsReferral.Consult_Reason = vbNullString Then clsReferral.Consult_Reason = rsReferral.Fields("REFERRAL_REASON")
                End If
            End If
        If lngPrimKey = 0 Then
              strSQL = "INSERT INTO SUPERBILL (PATIENT_ID, ENCOUNTER_ID, REFERRAL_ID, REFERRAL_REASON) values " _
                & "(" & piPatientID & "," & piEncounterID & "," & clsReferral.ReferralID & ",'" & clsReferral.Consult_Reason & "')"
                Set rsReferral = cnReferral.Execute(strSQL)
        Else
            strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
            strSQL = "UPDATE SUPERBILL SET REFERRAL_ID = " & clsReferral.ReferralID & "," _
                & "REFERRAL_REASON = '" & clsReferral.Consult_Reason & "', UPDATE_DATE = '" & strDate & "'  WHERE BILLING_ID = " & lngPrimKey
                Set rsReferral = cnReferral.Execute(strSQL)
        End If
    End If
    If cnReferral.State = adStateOpen Then
        cnReferral.Close
        Set cnReferral = Nothing
        Set rsReferral = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_ReferPhys:
    Screen.MousePointer = vbDefault
     Write_Error Err.Description, Err.Number, strSQL, "Update_ReferPhys", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cboFacility_Change()
Dim IntFacID As Integer
Dim i%

    On Error GoTo Error_cboFacility
    Screen.MousePointer = vbHourglass
    cboFacility.ColFromName = "FAC"
    For i = 0 To cboFacility.ListCount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            IntFacID = cboFacility.ItemData(i)
            piFacilityID = IntFacID
            Exit For
        End If
    Next i
    If IntFacID > 0 Then
        Load_Service Val(IntFacID)
        Load_Providers
    End If
    cmdUpdate.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cboFacility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cboFacility", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cboFacility_CloseUp()
Dim IntFacID As Integer
Dim i%

    On Error GoTo Error_cboFacility
    Screen.MousePointer = vbHourglass
    cboFacility.ColFromName = "FAC"
    For i = 0 To cboFacility.ListCount - 1
        If cboFacility.List(i) = cboFacility.Text Then
            IntFacID = cboFacility.ItemData(i)
            Exit For
        End If
    Next i
    If IntFacID > 0 Then Load_Service Val(IntFacID)
    cmdUpdate.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cboFacility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cboFacility", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cboFilter_Click()
    If cboFilter.Text <> vbNullString Then
        Get_BodySpecProcedure Trim(cboFilter.Text)
        mstrBodyFilter = Trim(cboFilter.Text)
        Get_BodySpecProcedure (mstrBodyFilter)
    End If
End Sub


Private Sub cboProvider_Click()
Dim Idx As Integer
Dim lngProv As Long

    If mbooByPass Then mbooByPass = False: Exit Sub
    lngProv = piProviderID
    
    If cboProvider.Text <> vbNullString Then
        For Idx = 0 To cboProvider.ListCount - 1
            If cboProvider.List(Idx) = cboProvider.Text Then
                piProviderID = cboProvider.ItemData(Idx)
                Exit For
            End If
        Next Idx
        Get_Providers piProviderID
'        If piProviderID <> mlngLogonUser Then
'            Clear_PracticeRole
'            Get_SecurityRole mlngLogonUser
'            If UCase(clsPracticeRoles.Permission) <> "ALL" Then
'                MsgBox "This is not your Patient Encounter! Please contact" _
'                    & " " & clsEncounter.Provider & " " & "for access permission or use Patient" _
'                    & "History to review this visit data. Thank you!", vbInformation, "PatientTrac Security Management"
'                    Screen.MousePointer = vbDefault
'                    Exit Sub
'                piProviderID = lngProv
'                mbooByPass = True
'                For Idx = 0 To cboProvider.ListCount - 1
'                    If cboProvider.ItemData(Idx) = lngProv Then
'                        cboProvider.Text = cboProvider.List(Idx)
'                        Exit For
'                    End If
'                Next Idx
'                mbooByPass = False
'            End If
'        End If
    End If
    mbooByPass = False
    Exit Sub
End Sub


Private Sub cboPurVisit_Click()
Dim i%
    If mbooByPass Then mbooByPass = False: Exit Sub
    If cboStatus.Text = "Procedure" Then Exit Sub
    For i = 0 To cboPurVisit.ListCount - 1
        If cboPurVisit.List(i) = cboPurVisit.Text Then
           clsEncounter.Intrvn_CD = LookUpCPTCode(Trim(cboPurVisit.Text))
            Exit For
        End If
    Next i
    cmdUpdate.Enabled = True
    
    cmdUpdate.Enabled = True
    Select Case clsEncounter.ServicePlaceCode
        Case "21", "51", "52"
            fraAdmission.Visible = True
            ccdAdmit.Enabled = True
            ccdDisharge.Enabled = True
            Get_AdmissionDate
            If TextFound("Initial", cboPurVisit.Text, Len("Initial"), True) Then
                If Trim(txtVisitReason.Text) <> vbNullString And Not TextFound("Admission", txtVisitReason.Text, Len("Admission"), True) Then
                    txtVisitReason.Text = "Admission" & vbNewLine & txtVisitReason.Text
                ElseIf Trim(txtVisitReason.Text) = vbNullString Then
                    txtVisitReason.Text = "Admission"
                End If
            End If
            fraAdmission.Visible = True
    End Select
'
'    If TextFound("Initial", cboPurVisit.Text, Len("Initial"), True) And fraAdmission.Visible Then
'        fraAdmission.Visible = True
'        ccdAdmit.Enabled = True
'        Get_AdmissionDate
'        If Trim(txtVisitReason.Text) <> vbNullString And Not TextFound("Admission", txtVisitReason.Text, Len("Admission"), True) Then
'            txtVisitReason.Text = "Admission" & vbNewLine & txtVisitReason.Text
'        ElseIf Trim(txtVisitReason.Text) = vbNullString Then
'            txtVisitReason.Text = "Admission"
'        End If
'    End If
    If TextFound("Subsequent", cboPurVisit.Text, Len("Subsequent"), True) Then fraAdmission.Visible = True: ccdDisharge.Enabled = True
    If TextFound("Discharge", cboPurVisit.Text, Len("Discharge"), True) Then
            fraAdmission.Visible = True
            ccdDisharge.Enabled = True
            Get_AdmissionDate
    End If
    If TextFound("Discharge", cboPurVisit.Text, Len("Discharge"), True) Then
        txtVisitReason.Text = "Discharge"
    End If
    If TextFound("Procedure", cboPurVisit.Text, Len("Procedure"), True) And cboStatus.Text <> "Established Patient" Then
        fraAdmission.Visible = True
        ccdAdmit.Enabled = True
        Get_AdmissionDate
        txtVisitReason.Text = "Admission"
        ccdDisharge.Enabled = False
        If mstrPractice <> "Psychiatry" Then cmdSurgLog.Visible = True
    End If
    If TextFound("Recovery Room", cboPurVisit.Text, Len("Recovery Room"), True) Then
        fraAdmission.Visible = True
        ccdAdmit.Enabled = True
        Get_AdmissionDate
        txtVisitReason.Text = "Recovery Room"
    End If
    If TextFound("Post Operative Visit", cboPurVisit.Text, Len("Post Operative Visit"), True) Or _
        TextFound("Post-Operative Visit", cboPurVisit.Text, Len("Post-Operative Visit"), True) Then
        If cboVisitReason.ListCount <= 0 Then GetList cboVisitReason, "sd_PostSurg_VisitType"
        cboVisitReason.Visible = True: cboVisitReason.ZOrder 0
        txtVisitReason.Visible = False
        txtVisitReason.ZOrder 1
    End If
    Exit Sub
End Sub
Private Function LookUpCPTCode(vstrCPTName As String) As String
Dim strSQL As String
Dim cnCode As New ADODB.Connection
Dim rsCode As New ADODB.Recordset
Dim strProvider As String

    On Error GoTo Error_LookUpCPTCode
    Screen.MousePointer = vbHourglass
    
    LookUpCPTCode = vbNullString
    If cnCode.State = adStateClosed Then cnCode.open psConnect(1)
    strSQL = "Select CPT_CODE, CPT_SUBGROUP " _
        & "From CPT_TEMPLATE " _
        & "where  CPT_NAME like  '" & vstrCPTName & "%' "
    Set rsCode = cnCode.Execute(strSQL)
    With rsCode
        If Not .EOF Then
            If Not IsNull((!CPT_Code)) Then
                If (!CPT_Code) <> vbNullString Then LookUpCPTCode = Trim(!CPT_Code)
            End If
        End If
        rsCode.Close
    End With
    If cnCode.State = adStateOpen Then
        cnCode.Close
        Set cnCode = Nothing
        Set rsCode = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUpCPTCode:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUpCPTCode", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function
Private Function LookUpCPTName(vstrCPTCode As String) As String
Dim i%
Dim strSQL As String
Dim cnCode As New ADODB.Connection
Dim rsCode As New ADODB.Recordset
Dim Arr

    On Error GoTo Error_LookUpCPTName
    Screen.MousePointer = vbHourglass
    LookUpCPTName = vbNullString
    Arr = Split(vstrCPTCode, ":")
    If UBound(Arr) > 0 Then
        vstrCPTCode = Arr(0)
    End If
     clsGridFormat.Text = vbNullString
    If vstrCPTCode <> vbNullString Then
        If cnCode.State = adStateClosed Then cnCode.open psConnect(1)
        strSQL = "Select CPT_NAME FROM CPT_TEMPLATE WHERE CPT_CODE like  '" & vstrCPTCode & "%'"
        Set rsCode = cnCode.Execute(strSQL)
        With rsCode
            If Not .EOF Then
                If Not IsNull((!CPT_Name)) Then
                    If (!CPT_Name) <> vbNullString Then LookUpCPTName = Trim(!CPT_Name)
                End If
            End If
        End With
        If LookUpCPTName = vbNullString Then
            strSQL = "Select CPT_INTRVN_NAME, CPT_FEE FROM INTRVN_TEMPLATE WHERE CPT_INTRVN_CODE like  '" & vstrCPTCode & "%'"
                Set rsCode = cnCode.Execute(strSQL)
                With rsCode
                    If Not .EOF Then
                        If Not IsNull((!CPT_Intrvn_Name)) Then
                            If (!CPT_Intrvn_Name) <> vbNullString Then LookUpCPTName = Trim(!CPT_Intrvn_Name)
                        End If
                        If Not IsNull((!CPT_Fee)) Then
                            If (!CPT_Fee) <> vbNullString Then clsGridFormat.Text = Trim(!CPT_Fee)
                        End If
                    End If
                End With
        End If
    End If
    If cnCode.State = adStateOpen Then
        cnCode.Close
        Set cnCode = Nothing
        Set rsCode = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUpCPTName:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUpCPTName", App.EXEName & "." & TypeName(Me)
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

    cboStatus.Clear
    clsEncounter.ServicePlaceCode = vbNullString

    If cboStatus.ListCount <= 0 Then
        If TextFound("Office", cboSerPlace.Text, Len("Office"), True) Then
            cboStatus.AddItem "Established Patient"
            cboStatus.AddItem "New Patient"
            If mstrPractice <> "Plastic Surgery" Then cboStatus.AddItem "Consultation"
            cboStatus.AddItem "Second Opinion"
         ElseIf TextFound("InPatient Hospital", cboSerPlace.Text, Len("InPatient Hospital"), True) Then
            fraAdmission.Visible = True
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
        ElseIf TextFound("Nursing Facility", cboSerPlace.Text, Len("Nursing Facility"), True) Then
            fraAdmission.Visible = True
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
        ElseIf TextFound("Assisted Living", cboSerPlace.Text, Len("Assisted Living"), True) Then
            fraAdmission.Visible = True
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
        ElseIf cboSerPlace.Text = "Inpatient Psychiatric Facility" Then
            fraAdmission.Visible = True
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
        ElseIf TextFound("Nursing Home", cboSerPlace.Text, Len("Nursing Home"), True) Then
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
            cboStatus.AddItem "Second Opinion"
         ElseIf TextFound("IOP", cboSerPlace.Text, Len("IOP"), True) Then
            fraAdmission.Visible = True
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
         ElseIf TextFound("PHP", cboSerPlace.Text, Len("PHP"), True) Then
            fraAdmission.Visible = True
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
         ElseIf TextFound("Hosp", cboSerPlace.Text, Len("Hosp"), True) Then
            fraAdmission.Visible = True
            cboStatus.AddItem "Initial InPatient"
            cboStatus.AddItem "Subsequent Care"
         ElseIf TextFound("OutPatient", cboSerPlace.Text, Len("OutPatient"), True) Then
            fraAdmission.Visible = True
            cboStatus.AddItem "Intake"
            cboStatus.AddItem "Established Patient"
         ElseIf TextFound("Ambulatory Surgical Center", cboSerPlace.Text, Len("Ambulatory Surgical Center"), True) Then
            fraAdmission.Visible = True
            cboStatus.AddItem "Procedure"
            cboStatus.AddItem "Admission"
            cboStatus.AddItem "Recovery Room"
            cboStatus.AddItem "Ambulatory Procedure Visit"
         End If
     End If
    If cboSerPlace.Text <> vbNullString Then
        For i = 0 To cboSerPlace.ListCount - 1
            If cboSerPlace.List(i) = cboSerPlace.Text Then
                clsEncounter.ServicePlaceCode = cboSerPlace.ItemData(i)
                Exit For
            End If
        Next i
    End If
    If TextFound("Office", cboSerPlace.Text, Len("Office"), True) Then
        ccdAdmit.Value = vbNullString
        fraAdmission.Visible = False
    End If
    Exit Sub
End Sub

Private Sub cboSerPlace_KeyUp(KeyCode As Integer, Shift As Integer)

        Exit Sub
End Sub


Private Sub cboStatus_Click()
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

    If mbooByPass Then mbooByPass = False: Exit Sub
    Screen.MousePointer = vbHourglass
    cboPurVisit.Clear
    lstListBox.Clear
    If Trim(clsProvider.Title) <> vbNullString Then
        strCompare = UCase(clsProvider.Title)
        strCompare = Replace(strCompare, ".", "")
        strCompare = Trim(strCompare)
    Else
        If piProviderID = 0 Then
            modSpecific.piProviderID = 0
            Get_DefaultProvider
            piProviderID = modSpecific.piProviderID
            Get_Providers piProviderID
        Else
            Get_Providers piProviderID
        End If
        strCompare = UCase(clsProvider.Title)
        strCompare = Replace(strCompare, ".", "")
        strCompare = Trim(strCompare)
    End If
    
    If Not TextFound("Admission", cboStatus.Text, Len("Admission"), True) And _
        Not TextFound("Recovery Room", cboStatus.Text, Len("Recovery Room"), True) Then
            If cboStatus.Text <> vbNullString Then
                If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
                strSerPlace = Trim(cboSerPlace.Text)
                For i = 0 To cboSerPlace.ListCount - 1
                    If cboSerPlace.List(i) = cboSerPlace.Text Then
                        clsEncounter.ServicePlaceCode = cboSerPlace.ItemData(i)
                        Exit For
                    End If
                Next i
                strPatType = Trim(cboStatus.Text)
                strTitle = clsProvider.Title
                cboPurVisit.Clear
                strSQL = "SELECT * FROM CPT_TEMPLATE " _
                   & "WHERE PLACE_SERVICE_CD = " & clsEncounter.ServicePlaceCode & " AND ENCNTR_TYPE like '" & strPatType & "%' ORDER BY CPT_NAME"
                Set rsCodes = cnCodes.Execute(strSQL)
 '-----------------
                 Do Until rsCodes.EOF
                    strName = vbNullString
                    strCode = vbNullString
                    strProvider = vbNullString
                    If rsCodes.Fields("PROVIDER") <> vbNullString Then
                        strProvider = UCase(rsCodes.Fields("PROVIDER"))
                        strProvider = Replace(strProvider, ".", "")
                        strProvider = Trim(strProvider)
                        If strProvider = strCompare Then
                            Set Flds = rsCodes.Fields
                            For Each Fld In Flds
                                Select Case strProvider
                                    Case "MD", "DO", "PA", "APN", "PA-C"
                                            If Not IsNull(Fld.Value) Then
                                                If Fld.Value <> vbNullString Then
                                                    If Fld.Name = "CPT_NAME" Then
                                                        strName = Fld.Value
                                                    ElseIf Fld.Name = "CPT_CODE" Then
                                                        strCode = Fld.Value
                                                    End If
                                                End If
                                            End If
                                    Case "MS", "MA", "LSW", "LPN", "CLINICIAN", "LCSW", "LPC", "PHD", "LSW", "LCSW", "LPN", "LGSW", "MSW", "LGPC", "LCPC", "LCS", "APRN", _
                                             "LPC", "LCSW-C", "LCSWC", "PSY-D", "CRNP", "APN", "CSC-AD", "LCADC", "LCAD-C", "MPH", "ICMS", "MA"
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
                    End If
                    rsCodes.MoveNext
                    If Trim(strName) <> vbNullString Then
                        cboPurVisit.AddItem strName
                        If IsNumeric(strCode) Then cboPurVisit.ItemData(cboPurVisit.NewIndex) = strCode
                    End If
                Loop
                If cboPurVisit.ListCount <= 0 Then
                    strPatType = Replace(strPatType, "Patient", "")
                    strPatType = Replace(strPatType, "Office", "")
                    strSQL = "SELECT * FROM CPT_TEMPLATE " _
                       & "WHERE PLACE_SERVICE_CD like '" & clsEncounter.ServicePlaceCode & "%' AND ENCNTR_TYPE like '%" & Trim(strPatType) & "%' ORDER BY CPT_NAME"
                    Set rsCodes = cnCodes.Execute(strSQL)
                     Do Until rsCodes.EOF
                        strName = vbNullString
                        strCode = vbNullString
                        strProvider = vbNullString
                        If rsCodes.Fields("PROVIDER") <> vbNullString Then
                            strProvider = UCase(rsCodes.Fields("PROVIDER"))
                            strProvider = Replace(strProvider, ".", "")
                            strProvider = Trim(strProvider)
                            If strProvider = strCompare Then
                                Set Flds = rsCodes.Fields
                                For Each Fld In Flds
                                Select Case strProvider
                                        Case "MD", "DO", "PA", "PA-C"
                                                If Not IsNull(Fld.Value) Then
                                                    If Fld.Value <> vbNullString Then
                                                        If Fld.Name = "CPT_NAME" Then
                                                            strName = Fld.Value
                                                        ElseIf Fld.Name = "CPT_CODE" Then
                                                            strCode = Fld.Value
                                                        End If
                                                    End If
                                                End If
                                        Case "MS", "MA", "LSW", "LPN", "CLINICIAN", "LCSW", "LPC", "PHD", "LSW", "LCSW", "LPN", "LGSW", "MSW", "LGPC", "LCPC", "LCS", "APRN", _
                                             "LPC", "LCSW-C", "LCSWC", "PSY-D", "CRNP", "APN", "CSC-AD", "LCADC", "LCAD-C"
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
                        End If
                        rsCodes.MoveNext
                        If Trim(strName) <> vbNullString Then
                            cboPurVisit.AddItem strName
                            If IsNumeric(strCode) Then cboPurVisit.ItemData(cboPurVisit.NewIndex) = strCode
                        End If
                    Loop
                End If
            End If
    End If
    If Trim(txtVisitReason.Text) = vbNullString Then
        If TextFound("Procedure", cboStatus.Text, Len("Procedure"), True) Or _
        TextFound("Admission", cboStatus.Text, Len("Admission"), True) Or _
        TextFound("Recovery Room", cboStatus.Text, Len("Recovery Room"), True) And mstrPractice <> "Psychiatry" Then
            If Not mbooGridClick Then
                Format_Procedures
                Load_Procedures
                fraEnctrProc.Visible = True
                cctProcString.Text = vbNullString
                cctProcString.Tag = vbNullString
                fraEnctrProc.ZOrder 0
                fraEnctr.Enabled = False
            End If
        End If
    End If
    If cboPurVisit.ListCount <= 0 And cboStatus.Text = "Procedure" Then
        cboPurVisit.AddItem "Procedure"
        cboPurVisit.ListIndex = 0
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
        Set Flds = Nothing
    End If
    cmdUpdate.Enabled = True
    If mstrPractice <> "Psychiatry" Then cmdSurgLog.Visible = True
    Screen.MousePointer = vbDefault
     Exit Sub
End Sub

Private Sub cboStatus_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim strSerPlace As String
Dim strPatType As String
Dim strTitle As String
Dim i%

    If mbooByPass Then mbooByPass = False: Exit Sub

    cboPurVisit.Clear
    lstListBox.Clear
    If Not TextFound("Admission", cboStatus.Text, Len("Admission"), True) And _
        Not TextFound("Recovery Room", cboStatus.Text, Len("Recovery Room"), True) Then
            If cboStatus.Text <> vbNullString Then
                If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
                strSerPlace = Trim(cboSerPlace.Text)
                strPatType = Trim(cboStatus.Text)
                strTitle = clsProvider.Title
                cboPurVisit.Clear
                strSQL = "Select CPT_CODE, CPT_NAME " _
                   & "From CPT_TEMPLATE " _
                   & "where PLACE_SERVICE_CD = " & clsEncounter.ServicePlaceCode & " AND ENCNTR_TYPE = '" & strPatType & "'"   'AND " _
                  ' & "PROVIDER = '" & strTitle & "'"
                Set rsCodes = cnCodes.Execute(strSQL)
                With rsCodes
                    Do Until .EOF
                        cboPurVisit.AddItem (!CPT_Name)
                        If IsNumeric(!CPT_Code) Then cboPurVisit.ItemData(cboPurVisit.NewIndex) = (!CPT_Code)
                    .MoveNext
                    Loop
                End With
                If cnCodes.State = adStateOpen Then
                    cnCodes.Close
                    Set cnCodes = Nothing
                End If
            End If
    End If
    If TextFound("Procedure", cboStatus.Text, Len("Procedure"), True) Or _
                 TextFound("Admission", cboStatus.Text, Len("Admission"), True) Or _
                 TextFound("Recovery Room", cboStatus.Text, Len("Recovery Room"), True) And mstrPractice <> "Psychiatry" Then
                 Load_Procedures
                fraEnctrProc.Visible = True
                fraEnctrProc.ZOrder 0
                fraEnctr.Enabled = False
    End If
    cmdUpdate.Enabled = True
    If mstrPractice <> "Psychiatry" Then cmdSurgLog.Visible = True
     Exit Sub
End Sub

Private Function Load_Procedures(Optional vstrProcedure)
Dim strServiceType As String
Dim strEncntrType As String
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim strSerPlace As String
Dim strPatType As String
Dim strTitle As String
Dim i%
Dim strVisitReas As String
Dim strCPT As String
Dim strNCpt As String
Dim lngRow As Long

    On Error GoTo Error_Load_Procedures
    Screen.MousePointer = vbHourglass
    cboRedo.Clear
    GetTerms cboRedo, "sd_Proc_Redo"
    If cboFilter.Text <> vbNullString Then mstrBodyFilter = cboFilter.Text
    modSpecific.piPatientID = piPatientID
    cboRedo.ListIndex = -1
    strTitle = Trim(clsProvider.Title)
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    If mstrPractice <> "Plastic Surgery" Then
        strSQL = "Select CPT_ID, CPT_CODE, CPT_NAME " _
            & "From CPT_TEMPLATE " _
            & "where ENCNTR_TYPE = 'PROCEDURE'"
        Set rsCodes = cnCodes.Execute(strSQL)
        With rsCodes
            Do Until .EOF
                If (!CPT_Name) <> vbNullString Then
                    lngRow = grdProc.MaxRows + 1
                    grdProc.MaxRows = lngRow
                    Format_CellTypeEdit grdProc, lngRow, 1, "1", (!CPT_Name), "TypeHAlignLeft", "PROCEDURES"
                    grdProc.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Format_CellTypeEdit grdProc, lngRow, 2, "1", (!CPT_Code), "TypeHAlignCenter", "CPT"
                    grdProc.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
            .MoveNext
            Loop
        End With
    ElseIf mstrPractice = "Plastic Surgery" Then
        If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
            strSQL = "Select PROCEDURE_ID, PROCEDURE_NAME, CPT_CODE " _
               & "FROM SURG_PROCEDURE_CODES ORDER BY BODY_AREA"
            Set rsCodes = cnCodes.Execute(strSQL)
            With rsCodes
                Do Until .EOF
                    If (!Procedure_Name) <> vbNullString Then
                        strCPT = Trim(!Procedure_Name)
                        If strCPT <> vbNullString Then
                            strCPT = IsProperCase(strCPT)
                        End If
                        lngRow = grdProc.MaxRows + 1
                        grdProc.MaxRows = lngRow
                        Format_CellTypeEdit grdProc, lngRow, 1, "1", strCPT, "TypeHAlignCenter", "PROCEDURES"
                        grdProc.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        If (!CPT_Code) <> vbNullString Then
                            Format_CellTypeEdit grdProc, lngRow, 2, "1", (!CPT_Code), "TypeHAlignCenter", "CPT CODE"
                            grdProc.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End If
                        If IsNumeric(!Procedure_ID) Then grdProc.SetRowItemData (lngRow), (!Procedure_ID)
                    End If
                .MoveNext
                Loop
            End With
            If cnCodes.State = adStateOpen Then
                cnCodes.Close
                Set cnCodes = Nothing
            End If
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Procedures:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Procedures", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub ccdDisharge_Click()
    If piEncounterID > 0 Then Manage_Admission piEncounterID
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
    If lstReferral.ListCount > 0 Then lstReferral.Visible = True
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

Private Sub cctRend_KeyUp(KeyCode As Integer, Shift As Integer)
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
Dim Idx As Integer

    On Error GoTo Error_Load_Referrals
    If Len(cctRend.Text) < 2 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If cnReferral.State = adStateClosed Then cnReferral.open psConnect(0)
    lstRend.Clear
    strRef = cctRend.Text
    If mstrRendType = "Referring Physician" Then
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
                        lstRend.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strTitle)
                    Else
                        strSuffix = Fix_Suffix(strSuffix)
                        lstRend.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strSuffix) & " " & Trim(strTitle)
                    End If
                    lstRend.ItemData(lstRend.NewIndex) = lngRefID
                End If
            Loop
        End With
        If cnReferral.State = adStateOpen Then
            cnReferral.Close
            Set cnReferral = Nothing
            Set rsReferral = Nothing
            Set Flds = Nothing
        End If
    Else
        Load_RenProv
        For Idx = 0 To lstRend.ListCount - 1
            If lstRend.List(Idx) = strRef Then
                lstRend.Selected(Idx) = True
                Exit For
            End If
        Next Idx
    End If
    If lstRend.ListCount > 0 Then lstRend.Visible = True
    Screen.MousePointer = vbDefault
   Exit Sub
Error_Load_Referrals:
    Screen.MousePointer = vbDefault
     Write_Error Err.Description, Err.Number, strSQL, "Load_Referrals", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub cctSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    strList = Trim(cctSearch.Text)
    If lstListBox.ListCount > 0 Then
        lstListBox.ListIndex = SendMessage(lstListBox.hwnd, LB_FINDSTRING, -1, _
        ByVal strList)
    End If
End Sub


Private Sub cmdAddTerms_Click()
'    pstrMaintType = "NextAppointment"
'    frmTermsMaint.PassedControl = ccNextAppt
'    frmTermsMaint.Show vbModal
End Sub


Private Sub cmdAdminHistory_Click()
    If piPatientID = 0 Then Exit Sub
    frmAdmitHistory.Show vbModal
    Exit Sub
End Sub



Private Sub cmdAdmit_Click()
Dim sDate As String
    mbAdmit = True
    ccdAdmit.Value = Format(Now(), "MM/DD/YYYY hh:mm:ss AM/PM")
    Exit Sub
End Sub


Private Sub cmdBillingInfo_Click()
    If piEncounterID = 0 Then
        MsgBox "Please select Visit/Encounter first!", vbInformation
        Exit Sub
    Else
        frmLoadControl.piSelPatientID = piPatientID
        frmLoadControl.piSelEnctrID = piEncounterID
        mstrPatOcx = "FN_PsyMse"
        mstrPatCtl = "Superbill"
        frmLoadControl.Caption = UCase("SUPERBILL")
        frmLoadControl.Show vbModal
    End If
    Exit Sub
End Sub


Private Sub cmdCancel_Click()
    CheckForChanges
    frmPatients.Show
    Unload Me
    Exit Sub
End Sub

Private Sub cmdClear_Click()
Dim i%
Dim iAnswer As Integer
Dim lngRefID As Long
Dim cnRefer As New ADODB.Connection
Dim rsRefer As New ADODB.Recordset
Dim strSQL As String
Dim strRef As String
Dim lngEnctrID As Long

    On Error GoTo Error_Clear_Referral
    Screen.MousePointer = vbHourglass
    
    If cnRefer.State = adStateClosed Then cnRefer.open psConnect(0)
    If Trim(cctReferral.Text) <> vbNullString Then
        iAnswer = MsgBox("Do you want to remove" & " " & cctReferral.Text & " " & "as referring physician?", vbYesNo, "PatientTrac Encounter")
        If iAnswer = vbYes Then
            lngRefID = clsReferral.ReferralID
            If lngRefID > 0 Then
                strSQL = "UPDATE APPOINTMENTS SET REFERRAL_ID = 0, REFERRAL_REASON = '' WHERE PATIENT_ID = " & Val(udtPatientName.IntId) & " AND REFERRAL_ID = " & lngRefID & " "
                Set rsRefer = cnRefer.Execute(strSQL)
                strSQL = "UPDATE SUPERBILL SET REFERRAL_ID = 0, REFERRAL_REASON = '' WHERE PATIENT_ID = " & Val(udtPatientName.IntId) & " AND REFERRAL_ID = " & lngRefID & " "
                Set rsRefer = cnRefer.Execute(strSQL)
                strSQL = "UPDATE  PATIENT_ACCOUNTS SET REFERRAL_ID = 0, REFERRAL_REASON = '' WHERE PATIENT_ID = " & Val(udtPatientName.IntId) & " AND REFERRAL_ID = " & lngRefID & " "
                Set rsRefer = cnRefer.Execute(strSQL)
                Clear_ReferralStrings
            End If
            cctReferral.Text = vbNullString
            lstReferral.Clear
        End If
    End If
    If cnRefer.State = adStateOpen Then
        cnRefer.Close
        Set cnRefer = Nothing
        Set rsRefer = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Clear_Referral:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Clear_Referral", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdClearRend_Click()
    lstRend.Clear
    cctRend.Text = vbNullString
    mstrRendType = "Referral Physician"
    Exit Sub
End Sub

Private Sub cmdCloseList_Click()
Dim i%
Dim strProcedures As String
Dim Arr
Dim Idx%
Dim iRow As Long

    Collect_ProcData

    clsOpNote.Preop_Code = vbNullString
    txtDesc.Text = vbNullString
    txtDesc.Visible = False
    clsOpNote.Procedure_Name = vbNullString
    clsOpNote.Procedure_ID = vbNullString
    If TextFound("Procedure", cboStatus.Text, Len("Procedure"), True) Then
        If cctProcString.Text <> vbNullString Then
            clsOpNote.Procedure_Name = Trim(cctProcString.Text)
            Arr = Split(cctProcString.Tag, vbNewLine)
            If UBound(Arr) > 0 Then
                For Idx = LBound(Arr) To UBound(Arr)
                    If Arr(Idx) <> vbNullString Then
                        If clsOpNote.Preop_Code = vbNullString Then
                            clsOpNote.Preop_Code = "049X" & " " & Arr(Idx)
                        ElseIf Not TextFound(Trim(Arr(Idx)), Trim(clsOpNote.Preop_Code), Len(Trim(Arr(Idx))), True) Then
                            clsOpNote.Preop_Code = clsOpNote.Preop_Code & "|" & "049X" & " " & Arr(Idx)
                        End If
                    End If
                Next Idx
                If clsOpNote.Preop_Code = vbNullString Then clsOpNote.Preop_Code = "049X"
            Else
                clsOpNote.Procedure_Name = Trim(mstrProc)
                clsOpNote.Preop_Code = "049X" & " " & mstrColCPT
            End If
        End If
    End If
    If clsOpNote.Procedure_Name = vbNullString Then clsOpNote.Procedure_Name = mstrProc
    If clsOpNote.Preop_Code = vbNullString Then clsOpNote.Preop_Code = "049X" & " " & Format(mstrColCPT, "00000")
    If clsOpNote.Procedure_ID = vbNullString Then clsOpNote.Procedure_ID = mlngProcID
    If clsOpNote.Procedure_Name <> vbNullString Then
        If Trim(txtVisitReason.Text) = vbNullString Then
            txtVisitReason.Text = clsOpNote.Procedure_Name
        ElseIf Not TextFound(Trim(clsOpNote.Procedure_Name), Trim(txtVisitReason.Text), Len(Trim(clsOpNote.Procedure_Name)), True) Then
             txtVisitReason.Text = Trim(clsOpNote.Procedure_Name) & vbNewLine & Trim(txtVisitReason.Text)
        End If
        clsEncounter.Intrvn_CD = clsOpNote.Preop_Code
        If TextFound("|", clsEncounter.Intrvn_CD, Len("|"), True) Then clsEncounter.Intrvn_CD = Replace(clsEncounter.Intrvn_CD, "|", vbNewLine)
        
    End If
    Get_Fat_Lipo_Details
    If Not TextFound(clsProcDetail.LipoDetails, txtVisitReason.Text, Len(clsProcDetail.LipoDetails), True) Then
        clsProcDetail.LipoDetails = Replace(clsProcDetail.LipoDetails, "Liposuction", "Liposuction of")
        clsProcDetail.LipoDetails = Replace(clsProcDetail.LipoDetails, vbNewLine & "   *", ",")
        Arr = Split(txtVisitReason.Text, "|")
        If UBound(Arr) > 0 Then
            For Idx = LBound(Arr) To UBound(Arr)
                If TextFound("Liposuction of", Arr(Idx), Len("Liposuction of"), True) Then
                    txtVisitReason.Text = Replace(txtVisitReason.Text, "Liposuction of", clsProcDetail.LipoDetails)
                ElseIf TextFound("Liposuction", Arr(Idx), Len("Liposuction"), True) Then
                    txtVisitReason.Text = Replace(txtVisitReason.Text, "Liposuction", clsProcDetail.LipoDetails)
                End If
            Next Idx
        Else
            If TextFound("Liposuction of", txtVisitReason.Text, Len("Liposuction of"), True) Then
                txtVisitReason.Text = Replace(txtVisitReason.Text, "Liposuction of", clsProcDetail.LipoDetails)
            ElseIf TextFound("Liposuction", txtVisitReason.Text, Len("Liposuction"), True) Then
                txtVisitReason.Text = Replace(txtVisitReason.Text, "Liposuction", clsProcDetail.LipoDetails)
            End If
        End If
    End If
    For i = 0 To cboPurVisit.ListCount - 1
        If cboPurVisit.List(i) = "Ambulatory Procedure Visit" Then
            cboPurVisit.Text = cboPurVisit.List(i)
            Exit For
        End If
    Next i
    grdProc.BlockMode = True
        grdProc.Col = 1
        grdProc.Col2 = grdProc.MaxCols
        grdProc.Row = mlngProcRow
        grdProc.Row2 = mlngProcRow
        grdProc.BackColor = NormalColor
    grdProc.BlockMode = False
    If clsOpNote.Preop_Code <> vbNullString Then clsEncounter.Intrvn_CD = clsOpNote.Preop_Code

    If TextFound("|", txtVisitReason.Text, Len("|"), True) Then txtVisitReason.Text = Replace(txtVisitReason.Text, "|", vbNewLine)
    'Code for Surgical Procedure 049X - for Ambulatory Surgery Center
    txtVisitReason.Text = Trim(txtVisitReason.Text)
    fraEnctrProc.Visible = False
    fraEnctrProc.ZOrder 1
    fraEnctr.Enabled = True
    cmdUpdProc.Visible = False
    Update_GridRow
    cboRedo.ListIndex = -1
    With grdColumn
        mlngProcRow = 0
        .ReDraw = False
        For iRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = iRow
                .Row2 = iRow
                .BackColor = NormalColor
                .CellType = CellTypeEdit
            .BlockMode = False
        Next iRow
        .ReDraw = True
    End With
    Exit Sub
End Sub

Private Function Update_GridRow()
Dim Arr
Dim Idx As Integer
Dim lngCol As Long

    Arr = Split(LookUpColName("REASON VISIT"), "|")
    If UBound(Arr) > 0 Then
        lngCol = Arr(0)
        Format_CellTypeEdit grdEnctrSummary, mlngGridRow, lngCol, "1", clsOpNote.Procedure_Name, "TypeHAlignCenter", "VISIT_REASON", &HC0FFFF
    End If
    Arr = Split(LookUpColName("CPT CODE"), "|")
    If UBound(Arr) > 0 Then
        lngCol = Arr(0)
        Format_CellTypeEdit grdEnctrSummary, mlngGridRow, lngCol, "1", clsEncounter.Intrvn_CD, "TypeHAlignCenter", "ENCOUNT_CPT_INTRVN", &HC0FFFF
    End If
    Exit Function
End Function

Private Sub cmdDeleteEncounter_Click()
Dim intResp As Byte
Dim strTables(6) As String
Dim strSQL As String
Dim cnPurge As New ADODB.Connection
Dim rsPurge As New ADODB.Recordset
Dim i%
Dim booUsed As Boolean

    On Error GoTo Error_Delete
    Screen.MousePointer = vbHourglass
    If piEncounterID = 0 Then
        MsgBox "Select an encounter first.", vbInformation, "PatientTrac Encounter"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    intResp = MsgBox("Are you sure you want to delete selected encounter?", vbYesNo, "PatientTrac Encounter")
    If intResp = vbYes Then
        Screen.MousePointer = vbHourglass
        Values_Psychiatry "REVIEW OF SYSTEMS"
        Values_Psychiatry "PHYSICAL EXAMINATION"
        Values_Psychiatry "MENTAL STATUS"
        Values_Psychiatry "FOLLOW-UP"
        If mstrRowData <> vbNullString Then
            MsgBox "Encounter is in use, cannot delete. Please contact PatientTrac Support to remove erroneous Encounter: (786) 350-3900", vbExclamation, "PatientTrac Encounter"
        Else
            If cnPurge.State = adStateClosed Then cnPurge.open psConnect(0)
            strSQL = "UPDATE ENCOUNTER SET PURGE = 1 WHERE ENCOUNTER_ID = " & piEncounterID
            Set rsPurge = cnPurge.Execute(strSQL)
            strSQL = "DELETE SUPERBILL WHERE ENCOUNTER_ID  = " & piEncounterID
            Set rsPurge = cnPurge.Execute(strSQL)
            strSQL = "DELETE PATIENT_PAYMENTS FROM PATIENT_PAYMENTS a, PATIENT_INVOICE b " _
                & "WHERE b.ENCOUNTER_ID  = " & piEncounterID & " And a.INVOICE_NO = b.INVOICE_NO"
            Set rsPurge = cnPurge.Execute(strSQL)
            strSQL = "DELETE PATIENT_INVOICE WHERE ENCOUNTER_ID  = " & piEncounterID
            Set rsPurge = cnPurge.Execute(strSQL)
        End If
    End If
    Init_Summary
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Delete:
    Screen.MousePointer = vbDefault
     Write_Error Err.Description, Err.Number, strSQL, "Delete Encounter", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdDiag_Click()

    If piEncounterID > 0 Then
        Select Case mstrPractice
            Case "Gyn", "IVF"
                    frmDiagnosis.mbooAdmin = False
                    frmDiagnosis.Caption = "DIAGNOSIS"
                    frmDiagnosis.Show vbModal
                    Values_ObGyn "DIAGNOSIS"
                    If Trim(clsReport.Diagnosis) <> vbNullString Then
                        cctDiagnosis.Text = Trim(clsReport.Diagnosis)
                    End If
            Case "Psychiatry"
                    mstrPatOcx = "FN_PsyMse"
                    mstrPatCtl = "Diagnosis"
                    frmLoadControl.Caption = "DIAGNOSIS"
                    frmLoadControl.piSelEnctrID = piEncounterID
                    frmLoadControl.piSelPatientID = piPatientID
                    frmLoadControl.Show vbModal
                    Values_Psychiatry "DIAGNOSIS"
                    If Trim(clsReport.AXIS_I) <> vbNullString Then
                        cctDiagnosis.Text = Trim(clsReport.AXIS_I)
                    End If
        End Select
    Else
        MsgBox "Save Encounter first, then Add Diagnosis and Select Update!, Must have an encounter for a diagnosis", vbInformation, "PatientTrac Encounter"
    End If
    Exit Sub
End Sub


Private Sub cmdDischarge_Click()
Dim sDate As String
    ccdDisharge.Enabled = True
    ccdDisharge.Value = Format(Now(), "MM/DD/YYYY hh:mm:ss AM/PM")
    Exit Sub
End Sub

Private Sub cmdDone_Click()
Dim iRow As Long

    Collect_ProcData
    cboRedo.ListIndex = -1
    With grdColumn
        mlngProcRow = 0
        .ReDraw = False
        For iRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = iRow
                .Row2 = iRow
                .BackColor = NormalColor
                .CellType = CellTypeEdit
            .BlockMode = False
        Next iRow
        .ReDraw = True
    End With
    Exit Sub
End Sub

Private Sub cmdEnctrOK_Click()
Dim i As Integer

    fraAdmission.Visible = False
    If IsFormLoaded("frmPatients") Then Unload frmPatients
    If piEncounterID = 0 Then
        MsgBox "Please select an encounter from the Encounter Summary First!", vbInformation
        Exit Sub
    End If
    If cboPurVisit.Text = vbNullString Then
        If clsEncounter.VisitType <> vbNullString Then
            For i = 0 To cboPurVisit.ListCount - 1
                If cboPurVisit.List(i) = clsEncounter.VisitType Then
                    mbooByPass = True
                    cboPurVisit.Text = cboPurVisit.List(i)
                    Exit For
                    mbooByPass = False
                End If
            Next i
        Else
            MsgBox "PatientTrac requires Visit Type, Please Update Encounter with a Visit Type!", vbCritical
        End If
        Exit Sub
    End If
    CheckForChanges
    Get_Patient piPatientID
    frmMDI.StatusBar.Panels(1).Text = "Patient: " & clsPatient.Name
    frmMDI.StatusBar.Panels(2).Text = "External ID: " & clsPatient.ExtID
    frmMDI.StatusBar.Panels(3).Text = "Internal ID: " & clsPatient.IntId
    frmMDI.StatusBar.Panels(4).Text = "Encounter: " & piEncounterID
    frmMDI.StatusBar.Panels(5).Text = "Visit Date: " & Format(clsEncounter.StartDate, "MMMM DD, YYYY")
    frmMDI.StatusBar.Panels(6).Text = "Provider: " & clsEncounter.Provider

    If mstrPractice = "Gyn" Or mstrPractice = "IVF" Then
        If IsFormLoaded("frmSuperbill") Then
            Create_Enctr_Default
            Unload Me
            Exit Sub
        ElseIf clsEncounter.Intrvn_CD = "99024" Then
            Create_Enctr_Default
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "PostOpVisit"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            End If
            Unload Me
            Exit Sub

       ElseIf Trim(clsOpNote.Procedure_Name) <> vbNullString And Trim(clsOpNote.Procedure_ID) <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "Surgery"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            Unload Me
            Exit Sub
        Else
            Create_Enctr_Default
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.Load_Reports
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.Show
                    frmInPatient.ZOrder 0
                End If
            Unload Me
            Exit Sub
        End If
    ElseIf mstrPractice = "Psychiatry" Then
        If IsFormLoaded("frmSuperbill") Then
            Create_Enctr_Default
            Unload Me
            Exit Sub
        Else
            Create_Enctr_Default
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Load_Reports
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            End If
            Unload Me
            Exit Sub
        End If
    ElseIf mstrPractice = "Plastic Surgery" Then
        Get_Encounter piEncounterID
        If IsFormLoaded("frmSuperbill") Then
            Create_Enctr_Default
            Unload Me
            Exit Sub
        ElseIf clsEncounter.Intrvn_CD = "99024" Then
            Create_Enctr_Default
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "PostOpVisit"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            End If
            Unload Me
            Exit Sub
        ElseIf clsEncounter.VisitType = "Ambulatory Procedure Visit" And TextFound("049X", clsEncounter.Intrvn_CD, Len("049X"), True) Then
            Create_Enctr_Default
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "Ambulatory Procedure Visit"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            End If
            Unload Me
            Exit Sub
        ElseIf clsEncounter.VisitType = "Pre-Operative Exam Anesthesia" And clsEncounter.Intrvn_CD = "99201" Then
            Create_Enctr_Default
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "PreOperative Anesthesia"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            End If
            Unload Me
            Exit Sub
            'Waiting for new "Pre-Operative Preparation"
'        ElseIf clsEncounter.VisitType = "Pre-Operative Preparation" Then
'            Create_Enctr_Default
'            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
'            If clsEncounter.EncounterType <> vbNullString Then
'               If Not IsFormLoaded("frmInPatient") Then
'                    frmInPatient.piSelEnctrID = 0
'                    frmInPatient.piSelEnctrID = piEncounterID
'                    frmInPatient.mstrLoadType = "Pre-Operative Preparation"
'                    frmInPatient.Read_Data piEncounterID
'                    frmInPatient.ZOrder 0
'                    frmInPatient.Show
'                Else
'                    frmInPatient.piSelEnctrID = piEncounterID
'                    frmInPatient.Read_Data piEncounterID
'                    frmInPatient.ZOrder 0
'                    frmInPatient.Show
'                End If
'            End If
'            Unload Me
'            Exit Sub
        ElseIf clsEncounter.VisitType = "Surgical Admission" And clsEncounter.Intrvn_CD = "90806" Then
            Create_Enctr_Default
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "Surgical Admission"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            End If
            Unload Me
            Exit Sub
       ElseIf Trim(clsOpNote.Procedure_Name) <> vbNullString And Trim(clsOpNote.Procedure_ID) <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "Surgery"
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            Unload Me
            Exit Sub
        Else
            Create_Enctr_Default
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Load_Reports
                    frmInPatient.Read_Data piEncounterID
                    
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                Else
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Read_Data piEncounterID
                    frmInPatient.ZOrder 0
                    frmInPatient.Show
                End If
            End If
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
     Exit Sub
End Sub


Private Sub cmdEnctrOK_GotFocus()
    cmdEnctrOK_KeyPress 13
End Sub


Private Sub cmdEnctrOK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
        cmdEnctrOK_Click
    End If
End Sub


Private Sub cmdEndEncounter_Click()
Dim strSQL As String
Dim bRet As Boolean
Dim lngencounterid As Long
Dim i As Integer
Dim lPatID As Long
Dim strAppt As String
Dim strStart As String
Dim strEnd As String
Dim lngdelta As Long
Dim strUser As String
Dim strNextAppt As String
Dim bRet1 As Boolean
Dim strRange As String
Dim strRoom As String
Dim strRoomPhone As String
Dim booRet As Boolean
Dim mstrSaveData As String

    On Error GoTo Error_EndEnctr
        fbooEndClicked = True
        If piEncounterID = 0 Then
            MsgBox "Select an encounter first.", vbInformation
            Exit Sub
        End If
        If clsEncounter.EndDate <> vbNullString Then
            MsgBox "Encounter is already closed", vbInformation
            SetControls True, False, False, True, False
            Exit Sub
        End If
        If cboStatus.Text = vbNullString Then
            MsgBox "Select an encounter type.", vbInformation
            Exit Sub
        End If
        If cboSerPlace.Text = vbNullString Then
            MsgBox "Select a place of service.", vbInformation
            Exit Sub
        End If
        If Not Date_Validation Then fbooEndClicked = False: fraEnctr.Visible = True: Exit Sub
        clsEncounter.Appointment = Format(Get_NextAppointment, "mm/dd/yyyy hh:mm:ss AM/PM")
        cmdSave.Enabled = False
        bRet = WriteData("E")
        If bRet Then
             strSQL = "exec a_enc_maint_checkout " _
              & "@ACTION_CODE = 'UA'," _
              & "@EncounterId = " & piEncounterID & "," _
              & "@PatientId = " & piPatientID & "," _
              & "@Comment = ''," _
              & "@NextAppt = '" & strNextAppt & "'"
            bRet1 = ActionQuery(strSQL, psConnect(0))
            mstrSaveData = lPatID
            strRange = Format(strStart, "mmm dd yyyy hh:nn AM/PM") & " - " & Format(strEnd, "mmm dd yyyy hh:nn AM/PM")
            MsgBox "The Encounter has been closed.", vbInformation
        End If
        Init_Summary
        SetControls True, False, False, True, False
        fbooEndClicked = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_EndEnctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "EndEnctr", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub cmdEntrDefault_Click()
    frmEnctrDefault.Show vbModal
    piProviderID = frmEnctrDefault.piProviderID
End Sub

Private Function Date_Validation() As Boolean
Dim i As Integer
Dim strAppt As String
Dim strStart As String
Dim strEnd As String
Dim lngdelta As Long
Dim strEndDate As String
Dim intPos As Integer
Dim bytResp As Byte

    Date_Validation = True
    If ccdStartDate.Value <> vbNullString Then strStart = Format(ccdStartDate.Value, "mm/dd/yyyy")
    If ccdEndDate.Value <> vbNullString Then strEnd = Format(ccdEndDate.Value, "mm/dd/yyyy")
    If fbooEndClicked Then
        If IsDate(strEnd) Then
            lngdelta = DateDiff("s", strStart, strEnd)
            If lngdelta < 0 Then
                MsgBox "Encounter start date can't exceed End date.", vbExclamation
                Date_Validation = False
                Exit Function
            End If
        Else
            bytResp = MsgBox("End date has not been entered.  Do you want to use today's date?", vbYesNo + vbInformation, "Close Encouter")
            If bytResp = vbYes Then
                strEnd = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
                ccdEndDate.Value = strEnd
                Date_Validation = True
            Else
                Date_Validation = False
            End If
            Exit Function
        End If
    End If
    Exit Function
End Function
Public Sub Enctr_Buttons()
Dim strDate As String
Dim strPrev As String
Dim strNext As String
Dim strEnd As String

    With grdEnctrSummary
        .ReDraw = False
        If mlngGridRow > 1 Then
            cmdNext.Visible = True
            cmdNext.Enabled = True
            mbooNext = True
        Else
            cmdNext.Enabled = False
            mbooNext = False
        End If
        If mlngGridRow < .MaxRows Then
            cmdPrev.Visible = True
            cmdPrev.Enabled = True
            mbooPrev = True
        Else
            cmdPrev.Enabled = False
            mbooPrev = False
        End If
        If clsEncounter.EndDate <> vbNullString Then
            SetControls True, False, False, True, False
            cmdDeleteEncounter.Enabled = False
        Else
            SetControls True, False, True, True, True
            cmdDeleteEncounter.Enabled = True
        End If
        .ReDraw = True
    End With
Exit Sub
End Sub
Public Sub Reset_Encounter()
    gbooBypass = True
    txtVisitReason.Text = vbNullString
    cboFacility.ListIndex = -1
    cboSerPlace.ListIndex = -1
    cboStatus.ListIndex = -1
    cboFacility.ListIndex = -1
    cboProvider.ListIndex = -1
    cboSerPlace.ListIndex = -1
    cboStatus.ListIndex = -1
    cctReferral.Text = vbNullString
    lstReferral.Clear
    lstRend.Clear
    cctRend.Text = vbNullString
    cctRefReason.Text = vbNullString
    cctRefReason.Visible = False
    cctDiagnosis.Text = vbNullString
    ccdStartDate.Value = vbNullString
    ccdEndDate.Value = vbNullString
    ccdAdmit.Value = vbNullString
    ccdAdmit.Enabled = False
    ccdDisharge.Value = vbNullString
    cctDiagnosis.Text = vbNullString
    mbAdmit = False
    mbDischarge = False
    cmdSave.Enabled = True
    cmdNew.Enabled = True
    cmdSave.Enabled = True
    cmdUpdate.Enabled = True
    cmdCancel.Enabled = True
    cmdEndEncounter.Enabled = True
    gbooBypass = False
    modSpecific.piEncounterID = 0
    Exit Sub
End Sub
Private Sub Set_Options()
Dim strOption As String

    On Error GoTo ERROR
     'Unlimited New Encounters......................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ServicePlaceReq")
    
    If strOption = vbNullString Then
        strOption = "no"
    End If
    
    If strOption = "yes" Then
        fbooServicePlaceReq = True
    Else
        fbooServicePlaceReq = False
    End If
    
    'Unlimited New Encounters......................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "UnlimitedEncounters")
    
    If strOption = vbNullString Then
        strOption = "no"
    End If
    
    If strOption = "yes" Then
        fbooUnlimitedEnctrs = True
    Else
        fbooUnlimitedEnctrs = False
    End If
    
    
    'Allow Closing Any Encounter......................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "CloseAnyEncounter")
    
    If strOption = vbNullString Then
        strOption = "no"
    End If
    
    If strOption = "yes" Then
        fbooCloseAnyEnctr = True
    Else
        fbooCloseAnyEnctr = False
    End If
    
    'Require Encounter Type......................................
    strOption = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "EncounterTypeRequired")
    
    If strOption = vbNullString Then
        strOption = "no"
    End If
    
    If strOption = "yes" Then
        fbooEncounterTypeReq = True
    Else
        fbooEncounterTypeReq = False
    End If

   Exit Sub
ERROR:
    Screen.MousePointer = vbDefault
     MsgBox Err.Description & " " & Err.Number, vbInformation
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub
Private Sub SetControls(bNew As Boolean, bSave As Boolean, bUpdate As Boolean, bCancel As Boolean, bEnctr As Boolean)
Dim bSet As Boolean
Dim i As Byte
Dim lstrRemarks As String


On Error GoTo Error_SetControls

    cmdNew.Enabled = bNew
    cmdSave.Enabled = bSave
    cmdUpdate.Enabled = bUpdate
    cmdCancel.Enabled = bCancel
    cmdEndEncounter.Enabled = bEnctr
    If fbooCloseAnyEnctr Then cmdEndEncounter.Enabled = True
    If bNew Then
        bSet = False
    Else
        bSet = True
    End If
    If bSave Then
        cmdCancel.Enabled = False
    Else
        cmdCancel.Enabled = True
    End If
    If pbooLockClosedEncounterData Then
        If fbooEncounterClosed Then
            cmdSave.Enabled = False
            cmdUpdate.Enabled = False
        End If
    End If
     Screen.MousePointer = vbDefault
    Exit Sub
Error_SetControls:
     Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "SetControls", App.EXEName & "." & Me.Name
     Exit Sub
     Resume
End Sub

Public Function Get_RoomPhone(pPatId As Long) As String

    On Error GoTo Error_Get_RoomPhone
    Get_RoomPhone = vbNullString
    Open_DB psConnect(0)
        cn.s_RoomInfoByPatID pPatId, rs
        With rs
            If Not .EOF Then
                If (!Telephone) <> vbNullString Then Get_RoomPhone = (!Telephone)
            End If
        End With
    Close_DB psConnect(0)
    Exit Function
Error_Get_RoomPhone:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_RoomPhone", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Private Function Get_Surgery_CPT(vstrname As String)
Dim Idx%
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim strSQL As String
Dim Arr
Dim strProc As String

    If vstrname = vbNullString Then Exit Function
    clsOpNote.Preop_Code = vbNullString
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    Arr = Split(vstrname, vbNewLine)
    If UBound(Arr) Then
        For Idx = LBound(Arr) To UBound(Arr)
            If Arr(Idx) <> vbNullString Then
                strProc = Arr(Idx)
                strSQL = "SELECT CPT_CODE, PROCEDURE_ID FROM SURG_PROCEDURE_CODES WHERE PROCEDURE_NAME = '" & strProc & "'"
                Set rsTerms = cnTerms.Execute(strSQL)
                If Not rsTerms.EOF Then
                   If rsTerms.Fields("CPT_CODE") <> vbNullString Then
                       If clsOpNote.Preop_Code = vbNullString Then
                           clsOpNote.Preop_Code = "049X" & "-" & rsTerms.Fields("CPT_CODE")
                       ElseIf Not TextFound(rsTerms.Fields("CPT_CODE"), clsOpNote.Preop_Code, Len(rsTerms.Fields("CPT_CODE")), True) Then
                           clsOpNote.Preop_Code = clsOpNote.Preop_Code & "|" & "049X" & "-" & rsTerms.Fields("CPT_CODE")
                       End If
                   End If
''                   If rsTerms.Fields("PROCEDURE_ID") <> vbNullString Then
''                       If clsOpNote.Procedure_ID = vbNullString Then
''                           clsOpNote.Procedure_ID = rsTerms.Fields("PROCEDURE_ID")
''                       ElseIf Not TextFound(rsTerms.Fields("CPT_CODE"), clsOpNote.Procedure_ID, Len(rsTerms.Fields("PROCEDURE_ID")), True) Then
''                           clsOpNote.Procedure_ID = Trim(clsOpNote.Procedure_ID) & "|" & rsTerms.Fields("PROCEDURE_ID")
''                       End If
''                   End If
                End If
            End If
        Next Idx
    Else
        strSQL = "SELECT CPT_CODE, PROCEDURE_ID FROM SURG_PROCEDURE_CODES WHERE PROCEDURE_NAME = '" & vstrname & "'"
        Set rsTerms = cnTerms.Execute(strSQL)
        If Not rsTerms.EOF Then
            If rsTerms.Fields("CPT_CODE") <> vbNullString Then
                If clsOpNote.Preop_Code = vbNullString Then
                    clsOpNote.Preop_Code = "049X" & "-" & rsTerms.Fields("CPT_CODE")
                ElseIf Not TextFound(rsTerms.Fields("CPT_CODE"), clsOpNote.Preop_Code, Len(rsTerms.Fields("CPT_CODE")), True) Then
                    clsOpNote.Preop_Code = clsOpNote.Preop_Code & "|" & "049X" & "-" & rsTerms.Fields("CPT_CODE")
                End If
'                If rsTerms.Fields("PROCEDURE_ID") <> vbNullString Then
'                    If clsOpNote.Procedure_ID = vbNullString Then
'                        clsOpNote.Procedure_ID = rsTerms.Fields("PROCEDURE_ID")
'                    ElseIf Not TextFound(rsTerms.Fields("CPT_CODE"), clsOpNote.Procedure_ID, Len(rsTerms.Fields("PROCEDURE_ID")), True) Then
'                        clsOpNote.Procedure_ID = Trim(clsOpNote.Procedure_ID) & "|" & rsTerms.Fields("PROCEDURE_ID")
'                    End If
'                End If
            End If
        End If
    End If
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Exit Function
End Function


Private Function Get_NextAppointment() As String
Dim strSQL As String
Dim sDate As String
Dim lngDateDiff As Long
 
    Get_NextAppointment = vbNullString
    If piEncounterID = 0 Then
        MsgBox "Please select an encounter first!, vbinformation"
        Exit Function
    End If
    If clsEncounter.StartDate = vbNullString Then Get_Encounter piEncounterID
    sDate = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    Open_DB psConnect(0)
    strSQL = "Select appt_date, appt_Time from APPOINTMENTS where patient_id = " & piPatientID & " AND " _
            & "ACTIVE = 1 ORDER BY APPT_DATE DESC"
    Set rs = cn.Execute(strSQL)
    With rs
        Do Until .EOF
           lngDateDiff = DateDiff("d", sDate, (!Appt_Date))
           If lngDateDiff > 1 Then
                If (!Appt_Date) <> vbNullString And (!APPT_TIME) <> vbNullString Then
                     mstrNextAppt = "Next Appointment" & " " & (!Appt_Date) & " " & (!APPT_TIME)
                     Get_NextAppointment = (!Appt_Date) & " " & (!APPT_TIME)
                     ccApptDate.Value = Format(Get_NextAppointment, "mm/dd/yyyy hh:mm:ss AM/PM")
                 End If
                Exit Do
            End If
        .MoveNext
        Loop
    End With
    Close_DB psConnect(0)
    Exit Function
End Function
Private Sub cmdNew_Click()
Dim i As Integer
Dim lbytResp As Byte
Dim strSQL As String
Dim varPrevPlan
Dim lngEnctr As Long
Dim intActX_ID As Integer
Dim varNumber
Dim varTerm
Dim varReason

    On Error GoTo Error_Select_NewEncounter
    Screen.MousePointer = vbHourglass
    mstrRendType = "Referral Physician"
    cctReferral.Text = vbNullString
    cctRend.Text = vbNullString
    cctRefReason.Text = vbNullString
    lstReferral.Visible = False
    lstReferral.Clear
    lstRend.Clear
    lstRend.Visible = False
    mlngPrevEnctrID = 0
    If piProviderID > 0 Then Get_PrevEncounter piEncounterID, piProviderID
    cmdSave.Enabled = True
    mbooNewEncntr = True
    txtVisitReason.Text = vbNullString
    mbDateCheck = True
    Validate_Enctr
    If mlngPrevEnctrID = 0 Then
        If piEncounterID > 0 Then
            If piProviderID > 0 Then Get_PrevEncounter piEncounterID, piProviderID
        End If
    End If
    lngEnctr = mlngPrevEnctrID
    txtVisitReason.Text = vbNullString
    If mstrPractice = "Psychiatry" Then
            varReason = Get_ClinicalValue("FN_PsyMse", "Orders", "cboPurpose_" & intActX_ID, 1, piPatientID, lngEnctr)
            If varReason <> vbNullString Then clsEncounter.VisitReason = varReason: ' txtVisitReason.Text = clsencounter.VisitReason
    ElseIf mstrPractice = "IVF" Or mstrPractice = "Gyn" Then
        If Trim(txtVisitReason.Text) = vbNullString Then
            varReason = Get_ClinicalValue("FN_ObGyn", "Orders", "cboPurpose_" & intActX_ID, 1, piPatientID, lngEnctr)
            varReason = Get_ClinicalValue("FN_ObGyn", "Orders", "cboPurpose", 1, piPatientID, lngEnctr)
        End If
    End If
    If Not fbooUnlimitedEnctrs Then
        If Provider_HasOpenEncounter Then
            lbytResp = MsgBox("Provider selected already has an open encounter." & vbNewLine _
                    & "Either close current encounter or select another provider." & vbNewLine & _
                    "Do you want to select another provider?", vbInformation + vbYesNo)
            If lbytResp = vbYes Then
                Reset_Encounter
                ccdStartDate.Value = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
                SetControls False, True, False, True, True
                Dump_EncounterStrData
                cmdCancel.Enabled = True
                Load_Facilities
                Load_Providers
                If cboSerPlace.Text <> vbNullString Then cboSerPlace_Click
                If fbooCloseAnyEnctr Then cmdEndEncounter.Enabled = True
                Get_EnctrInfo
                cmdCancel.Enabled = True
                cmdUpdate.Enabled = False
                cmdDeleteEncounter.Enabled = True
                Set_Default_Enctr
                ccdStartDate.Enabled = True
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                SetControls False, True, True, True, True
                cmdCancel.Enabled = True
                cmdDeleteEncounter.Enabled = True
                ccdStartDate.Enabled = True
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
         Else
            Reset_Encounter
            ccdStartDate.Value = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
            SetControls False, True, False, True, False
            cmdCancel.Enabled = True
            Load_Facilities
            Load_Providers
            If cboSerPlace.Text <> vbNullString Then cboSerPlace_Click
            If fbooCloseAnyEnctr Then cmdEndEncounter.Enabled = True
            Set_Default_Enctr
            cmdDeleteEncounter.Enabled = True
            cmdCancel.Enabled = True
            cmdUpdate.Enabled = False
            cmdDeleteEncounter.Enabled = True
            Screen.MousePointer = vbDefault
            ccdStartDate.Enabled = True
            Exit Sub
         End If
    Else
            fbooByPassEvent = True
            Reset_Encounter
            ccdStartDate.Value = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
            SetControls False, True, False, True, True
            cmdCancel.Enabled = True
            Load_Facilities
            Load_Providers
            If cboSerPlace.Text <> vbNullString Then cboSerPlace_Click
            cmdEndEncounter.Enabled = False
            If fbooCloseAnyEnctr Then cmdEndEncounter.Enabled = True
            fbooEndClicked = False
            fbooByPassEvent = False
            Get_EnctrInfo
            Set_Default_Enctr
            cmdDeleteEncounter.Enabled = True
            cmdCancel.Enabled = True
            cmdUpdate.Enabled = False
            ccdStartDate.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
    cmdDeleteEncounter.Enabled = True
    cmdCancel.Enabled = True
    cmdUpdate.Enabled = False
    ccdStartDate.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Select_NewEncounter:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Select_NewEncounter", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub cmdNext_Click()
Dim lintIndex As Integer
Dim lintResp As Integer
Dim booSkipSave As Boolean
Dim strSQL As String
Dim varRS
Dim lngrecs As Long
Dim bRet As Boolean
Dim lngNextID As Long
Dim i%
Dim lngEnctr As Long
Dim iRow%
Dim intNewRow As Long
    mbooNext = True
    If cmdPrev.AppearanceThemes = 1 Then
        cmdPrev.AppearanceThemes = 3
        mEnctrCounter = 0
    End If
    If cmdNext.AppearanceThemes = 1 Then
        cmdNext.AppearanceThemes = 1
        mEnctrCounter = 1
    Else
        mEnctrCounter = mEnctrCounter + 1
    End If

    Screen.MousePointer = vbHourglass
    With grdEnctrSummary
        .ReDraw = False
            If Val(mlngGridRow - 1) > 0 Then
                intNewRow = Val(mlngGridRow - 1)
                grdEnctrSummary.Col = 0
                grdEnctrSummary.Row = intNewRow
                lngNextID = .GetRowItemData(intNewRow)
            Else
                MsgBox "Invalid selection", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        If lngNextID > 0 Then
            grdEnctrSummary_DblClick 1, intNewRow
        End If
    End With
    frmMDI.StatusBar.Panels(4).Text = "Encounter: " & piEncounterID
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub


Private Sub cmdPatient_Click()
Dim i%
Dim varPatientID

    On Error GoTo Error_Patient
        Reset_Encounter
        Dump_EncounterStrData
        frmPatients.Show
        Unload Me
    Exit Sub
Error_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Patient_Select", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub cmdPrev_Click()
Dim lintIndex As Integer
Dim lintResp As Integer
Dim booSkipSave As Boolean
Dim strSQL As String
Dim varRS
Dim lngrecs As Long
Dim bRet As Boolean
Dim lngNextID As Long
Dim i%
Dim iRow
Dim intNewRow As Long

    mbooPrev = True
    If cmdNext.AppearanceThemes = 1 Then
        cmdNext.AppearanceThemes = 3
        mEnctrCounter = 0
    End If
    If cmdPrev.AppearanceThemes = 3 Then
        cmdPrev.AppearanceThemes = 1
        mEnctrCounter = 1
    Else
        mEnctrCounter = mEnctrCounter + 1
     End If
    Screen.MousePointer = vbHourglass

    With grdEnctrSummary
        .ReDraw = False
            grdEnctrSummary.Row = mlngGridRow
            lngNextID = .GetRowItemData(mlngGridRow)
            If Val(mlngGridRow + 1) <= grdEnctrSummary.MaxRows Then
                intNewRow = Val(mlngGridRow + 1)
                grdEnctrSummary.Col = 0
                grdEnctrSummary.Row = intNewRow
                lngNextID = .GetRowItemData(intNewRow)
            Else
                MsgBox "No Previous Encounters", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        If intNewRow <> 0 Then
            grdEnctrSummary_DblClick 1, intNewRow
        End If
        .ReDraw = True
    End With
    frmMDI.StatusBar.Panels(4).Text = "Encounter: " & piEncounterID
    Screen.MousePointer = vbDefault
Exit Sub

End Sub


Private Sub cmdRendProvDefault_Click()
Dim Idx As Integer

    mstrRendType = "Internal Provider"
    If piProviderID = 0 Then
        Get_DefaultProvider
    End If
    If piProviderID > 0 Then
        lstRend.Clear
        cctRend.Text = vbNullString
        Load_RenProv
        For Idx = 0 To lstRend.ListCount - 1
            If lstRend.ItemData(Idx) = piProviderID Then
                mbooByPass = True
                lstRend.Selected(Idx) = True
                cctRend.Text = lstRend.List(Idx)
                Exit For
            End If
        Next Idx
    End If
    Exit Sub
End Sub

Private Sub cmdReportToolBox_Click()
    Get_SurgProc
End Sub

Private Function Get_SurgProc()
Dim strSQL As String
Dim cnSurg As New ADODB.Connection
Dim rsSurg As New ADODB.Recordset
Dim lngRow As Long
Dim CellNote As Long
Dim strCPT As String

    On Error GoTo ERROR_Get_SurgProc

    If cnSurg.State = adStateClosed Then cnSurg.open psConnect(1)
    Format_Procedures
    strSQL = "SELECT PROCEDURE_NAME, PROCEDURE_ID, CPT_CODE FROM SURG_PROCEDURE_CODES WHERE SPECIALTY = '" & mstrPractice & "' ORDER BY PROCEDURE_NAME"
    Set rsSurg = cnSurg.Execute(strSQL)
    With rsSurg
        Do Until .EOF
            If (!Procedure_Name) <> vbNullString Then
                strCPT = Trim(!Procedure_Name)
                If strCPT <> vbNullString Then
                    strCPT = IsProperCase(strCPT)
                End If
            End If
            If strCPT <> vbNullString Then
                lngRow = grdProc.MaxRows + 1
                grdProc.MaxRows = lngRow
                Format_CellTypeEdit grdProc, lngRow, 1, "1", (!Procedure_Name), "TypeHAlignCenter", "PROCEDURES"
                grdProc.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                If (!CPT_Code) <> vbNullString Then
                    Format_CellTypeEdit grdProc, lngRow, 2, "1", (!CPT_Code), "TypeHAlignCenter", "CPT CODE"
                    grdProc.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
                If IsNumeric(!Procedure_ID) Then grdProc.SetRowItemData (lngRow), (!Procedure_ID)
            End If
            .MoveNext
        Loop
    End With
    If cnSurg.State = adStateOpen Then
        cnSurg.Close
        Set cnSurg = Nothing
        Set rsSurg = Nothing
    End If
    Exit Function
ERROR_Get_SurgProc:
    Write_Error Err.Description, Err.Number, strSQL, "Get_SurgProc", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cmdSave_Click()
Dim strSQL As String
Dim bRet As Boolean
Dim lngTxnmy As Long
Dim intFacilityID As Integer
Dim intProviderID As Integer
Dim strUser As String
Dim i As Integer
Dim strAppt As String
Dim strStart As String
Dim strEnd As String
Dim lngdelta As Long
Dim strRoom As String
Dim lPatID As Long
Dim strRoomPhone As String
Dim strVisitType As String
Dim sDate As String
Dim lngAdmitID As Long
Dim Arr

    On Error GoTo Error_Save_Encounter
    Screen.MousePointer = vbHourglass
    mbooNewEncntr = False
'    If TextFound("Consult", cboPurVisit.Text, Len("Consult"), True) And Trim(cctReferral.Text) = vbNullString Then
'        MsgBox "Please enter the Referring Physician Name and Reason for Consultation!", vbInformation, "PatientTrac Encounters"
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    If cboStatus.Text = vbNullString Then
        MsgBox "Select an encounter type.", vbInformation, "PatientTrac Encounter"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If cboSerPlace.Text = vbNullString Then
        MsgBox "Select a place of service.", vbInformation, "PatientTrac Encounter"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If cboPurVisit.Text = vbNullString Then
        MsgBox "PatientTrac requires Visit Type, Please Update Encounter with a Visit Type!", vbCritical, "PatientTrac Encounter"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    cmdCancel.Enabled = False
    fbooSaveClicked = True
    If Not Date_Validation Then
        fbooSaveClicked = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    strUser = PsUserID
    If Not IsNull(ccdStartDate.Value) Then strStart = Trim(ccdStartDate.Value)
    If Not IsNull(ccdEndDate) Then strEnd = Trim(ccdEndDate.Value)
    fbooSaveClicked = False
    If cboFacility.Text <> vbNullString Then
        For i = 0 To cboFacility.ListCount - 1
            If cboFacility.List(i) = cboFacility.Text Then
                piFacilityID = cboFacility.ItemData(i)
                Exit For
            End If
        Next i
    Else
        MsgBox "Select a facility.", vbExclamation, "PatientTrac Encounter"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If cboProvider.Text <> vbNullString Then
        For i = 0 To cboProvider.ListCount - 1
            If cboProvider.List(i) = Trim(cboProvider.Text) Then
                piProviderID = cboProvider.ItemData(i)
                Exit For
            End If
        Next i
    Else
        MsgBox "Select a provider.", vbExclamation, "PatientTrac Encounter"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    bRet = WriteData("I")
    If bRet Then
        MsgBox "New Encounter saved.", vbInformation, "PatientTrac Encounter"
    End If
    
    cmdEndEncounter.Enabled = True
    strVisitType = Trim(cboPurVisit.Text)
    Open_DB psConnect(0)
    If strVisitType <> vbNullString Then
         sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
        strSQL = "UPDATE ENCOUNTER SET PURGE = 0,  VISIT_TYPE = '" & strVisitType & "', UPDATE_DATE = '" & sDate & "' WHERE ENCOUNTER_ID = " & piEncounterID
        Set rs = cn.Execute(strSQL)
    End If
    Close_DB psConnect(0)
    If Trim(clsEncounter.Intrvn_CD) <> vbNullString Then
        Arr = Split(clsEncounter.Intrvn_CD, "-")
        If UBound(Arr) > 0 Then
            If Arr(0) = "049X" Then ' We have a procedure
                Save_SurgLog
             End If
        End If
    End If
    Init_Summary
    Enctr_Buttons
    cmdUpdate.Enabled = True
    mstrSaveData = lPatID
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Save_Encounter:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Save_Encounter", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub cmdStatus_Click()
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

    
    If mbooByPass Then mbooByPass = False: Exit Sub
    If cboPurVisit.ListCount <= 1 Then
       Screen.MousePointer = vbHourglass
       lstListBox.Clear
       If Trim(clsProvider.Title) <> vbNullString Then
           strCompare = UCase(clsProvider.Title)
           strCompare = Replace(strCompare, ".", "")
           strCompare = Trim(strCompare)
       Else
           If piProviderID = 0 Then
               modSpecific.piProviderID = 0
               Get_DefaultProvider
               piProviderID = modSpecific.piProviderID
               Get_Providers piProviderID
           Else
               Get_Providers piProviderID
           End If
           strCompare = UCase(clsProvider.Title)
           strCompare = Replace(strCompare, ".", "")
           strCompare = Trim(strCompare)
       End If
       
       If Not TextFound("Admission", cboStatus.Text, Len("Admission"), True) And _
           Not TextFound("Recovery Room", cboStatus.Text, Len("Recovery Room"), True) Then
               If cboStatus.Text <> vbNullString Then
                   If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
                   strSerPlace = Trim(cboSerPlace.Text)
                   For i = 0 To cboSerPlace.ListCount - 1
                       If cboSerPlace.List(i) = cboSerPlace.Text Then
                           clsEncounter.ServicePlaceCode = cboSerPlace.ItemData(i)
                           Exit For
                       End If
                   Next i
                   strPatType = Trim(cboStatus.Text)
                   strTitle = clsProvider.Title
                   cboPurVisit.Clear
                   strSQL = "SELECT * FROM CPT_TEMPLATE " _
                      & "WHERE PLACE_SERVICE_CD = " & clsEncounter.ServicePlaceCode & " AND ENCNTR_TYPE like '" & strPatType & "%' ORDER BY CPT_NAME"
                   Set rsCodes = cnCodes.Execute(strSQL)
    '-----------------
                    Do Until rsCodes.EOF
                       strName = vbNullString
                       strCode = vbNullString
                       strProvider = vbNullString
                       If rsCodes.Fields("PROVIDER") <> vbNullString Then
                           strProvider = UCase(rsCodes.Fields("PROVIDER"))
                           strProvider = Replace(strProvider, ".", "")
                           strProvider = Trim(strProvider)
                           If strProvider = strCompare Then
                               Set Flds = rsCodes.Fields
                               For Each Fld In Flds
                                   Select Case strProvider
                                       Case "MD", "DO", "PA", "APN", "PA-C"
                                               If Not IsNull(Fld.Value) Then
                                                   If Fld.Value <> vbNullString Then
                                                       If Fld.Name = "CPT_NAME" Then
                                                           strName = Fld.Value
                                                       ElseIf Fld.Name = "CPT_CODE" Then
                                                           strCode = Fld.Value
                                                       End If
                                                   End If
                                               End If
                                       Case "MS", "MA", "LSW", "LPN", "CLINICIAN", "LCSW", "LPC", "PHD", "LSW", "LCSW", "LPN", "LGSW", "MSW", "LGPC", "LCPC", "LCS", "APRN", _
                                                "LPC", "LCSW-C", "LCSWC", "PSY-D", "CRNP", "APN", "CSC-AD", "LCADC", "LCAD-C", "MPH", "ICMS", "MA"
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
                       End If
                       rsCodes.MoveNext
                       If Trim(strName) <> vbNullString Then
                           cboPurVisit.AddItem strName
                           If IsNumeric(strCode) Then cboPurVisit.ItemData(cboPurVisit.NewIndex) = strCode
                       End If
                   Loop
                   If cboPurVisit.ListCount <= 0 Then
                       strPatType = Replace(strPatType, "Patient", "")
                       strPatType = Replace(strPatType, "Office", "")
                       strSQL = "SELECT * FROM CPT_TEMPLATE " _
                          & "WHERE PLACE_SERVICE_CD like '" & clsEncounter.ServicePlaceCode & "%' AND ENCNTR_TYPE like '%" & Trim(strPatType) & "%' ORDER BY CPT_NAME"
                       Set rsCodes = cnCodes.Execute(strSQL)
                        Do Until rsCodes.EOF
                           strName = vbNullString
                           strCode = vbNullString
                           strProvider = vbNullString
                           If rsCodes.Fields("PROVIDER") <> vbNullString Then
                               strProvider = UCase(rsCodes.Fields("PROVIDER"))
                               strProvider = Replace(strProvider, ".", "")
                               strProvider = Trim(strProvider)
                               If strProvider = strCompare Then
                                   Set Flds = rsCodes.Fields
                                   For Each Fld In Flds
                                   Select Case strProvider
                                           Case "MD", "DO", "PA", "PA-C"
                                                   If Not IsNull(Fld.Value) Then
                                                       If Fld.Value <> vbNullString Then
                                                           If Fld.Name = "CPT_NAME" Then
                                                               strName = Fld.Value
                                                           ElseIf Fld.Name = "CPT_CODE" Then
                                                               strCode = Fld.Value
                                                           End If
                                                       End If
                                                   End If
                                           Case "MS", "MA", "LSW", "LPN", "CLINICIAN", "LCSW", "LPC", "PHD", "LSW", "LCSW", "LPN", "LGSW", "MSW", "LGPC", "LCPC", "LCS", "APRN", _
                                                "LPC", "LCSW-C", "LCSWC", "PSY-D", "CRNP", "APN", "CSC-AD", "LCADC", "LCAD-C", "MPH", "ICMS", "MA"
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
                           End If
                           rsCodes.MoveNext
                           If Trim(strName) <> vbNullString Then
                               cboPurVisit.AddItem strName
                               If IsNumeric(strCode) Then cboPurVisit.ItemData(cboPurVisit.NewIndex) = strCode
                           End If
                       Loop
                   End If
               End If
       End If
       If Trim(txtVisitReason.Text) = vbNullString Then
           If TextFound("Procedure", cboStatus.Text, Len("Procedure"), True) Or _
           TextFound("Admission", cboStatus.Text, Len("Admission"), True) Or _
           TextFound("Recovery Room", cboStatus.Text, Len("Recovery Room"), True) And mstrPractice <> "Psychiatry" Then
               If Not mbooGridClick Then
                   Format_Procedures
                   Load_Procedures
                   fraEnctrProc.Visible = True
                   cctProcString.Text = vbNullString
                   cctProcString.Tag = vbNullString
                   fraEnctrProc.ZOrder 0
                   fraEnctr.Enabled = False
               End If
           End If
       End If
       If cboPurVisit.ListCount <= 0 And cboStatus.Text = "Procedure" Then
           cboPurVisit.AddItem "Procedure"
           cboPurVisit.ListIndex = 0
       End If
    End If
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
        Set Flds = Nothing
    End If
    cmdUpdate.Enabled = True
    If mstrPractice <> "Psychiatry" Then cmdSurgLog.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub

End Sub

Private Sub cmdSurgLog_Click()
    Me.Hide
    mstrPatOcx = "PatientTrac_Report"
    mstrPatCtl = "SurgLog"
    frmLoadControl.Show vbModal
    Me.Show
End Sub

Private Sub cmdTermRedo_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Surgery Revision Type"
    frmNurseTemplates.fType = "Surgery Revision Type"
    frmNurseTemplates.fSubDomain = "sd_Proc_Redo"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    GetTerms cboRedo, "sd_Proc_Redo"
End Sub

Private Sub cmdUpdate_Click()
Dim bRet As Boolean
Dim i%
Dim strRoom As String
Dim lPatID As Long
Dim cnUpdate As New ADODB.Connection
Dim rsUpdate As New ADODB.Recordset
Dim strFields As String
Dim strValues As String
Dim strRoomPhone As String
Dim strSQL As String
Dim booRet As Boolean
Dim lngrecs As Long
Dim varRS
Dim Arr

    On Error GoTo Error_Update_Encounter
    Screen.MousePointer = vbHourglass
    If piEncounterID = 0 Then
        MsgBox "Select an encounter first.", vbInformation, "PatientTrac Encounters"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
'    If TextFound("Consult", cboPurVisit.Text, Len("Consult"), True) And Trim(cctReferral.Text) = vbNullString Then
'        MsgBox "Please enter the Referring Physician Name and Reason for Consultation!", vbInformation, "PatientTrac Encounters"
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    If cboPurVisit.Text = vbNullString Then
        If clsEncounter.VisitType <> vbNullString Then
            For i = 0 To cboPurVisit.ListCount - 1
                If cboPurVisit.List(i) = clsEncounter.VisitType Then
                    mbooByPass = True
                    cboPurVisit.Text = cboPurVisit.List(i)
                    Exit For
                    mbooByPass = False
                End If
            Next i
        End If
        MsgBox "PatientTrac requires Visit Type, Please Update Encounter with a Visit Type!", vbCritical, "PatientTrac Encounter"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Trim(cboStatus.Text) = vbNullString Then
        MsgBox "Select an encounter type.", vbInformation, "PatientTrac Encounter"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Trim(cboSerPlace.Text) = vbNullString Then
        MsgBox "Select a place of service.", vbInformation, "PatientTrac Encounter"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not Date_Validation Then
        MsgBox "Encounter not updated!", vbInformation, "PatientTrac Encounter"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If cboProvider.Text = vbNullString Or cboFacility.Text = vbNullString Then Get_EnctrInfo
    bRet = WriteData("U")
    If mstrPractice <> "Psychiatry" Then
         If Trim(clsEncounter.Intrvn_CD) <> vbNullString Then
            Arr = Split(clsEncounter.Intrvn_CD, "-")
            If UBound(Arr) > 0 Then
                If Arr(0) = "049X" Then ' We have a procedure
                    Save_SurgLog
                 End If
            End If
        End If
    End If
    mbooByPass = False
    lPatID = piPatientID
    Init_Summary
    cboVisitReason.Visible = False
    txtVisitReason.Visible = True
    txtVisitReason.ZOrder 0
    Enctr_Buttons
    MsgBox "Encounter Successfully Updated.", vbInformation, "PatientTrac Encounter"
    mbooNewEncntr = False
    Get_AdmissionDate
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Update_Encounter:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_Encounter", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Public Sub Init_Summary()
Dim i As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol  As Long
Dim varCellValue
Dim cnEnctr As New ADODB.Connection
Dim rsEnctr As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strCPT As String
Dim Arr
Dim strEnctrVisit As String
Dim booSkip As Boolean

    On Error GoTo Error_Init_Summary
    ccdEndDate.Value = vbNullString
    If piPatientID = 0 Then Exit Sub
    Get_PatientName piPatientID
    If cboFacility.ListCount <= 0 Then Get_Facilities
    With grdEnctrSummary
        .ReDraw = False
        .ClearRange 0, 0, grdEnctrSummary.MaxCols, grdEnctrSummary.MaxRows, False
        fraEnctr.Visible = True
        grdEnctrSummary.MaxCols = 10
        .ColWidth(0) = 7
        .ColWidth(1) = 9
        .ColWidth(2) = 9
        .ColWidth(3) = 14
        .ColWidth(4) = 18
        .ColWidth(5) = 10
        .ColWidth(6) = 18
        .ColWidth(7) = 12
        .ColWidth(8) = 15
        .ColWidth(9) = 15
        .ColWidth(10) = 15
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 0, "1", "Enctr_ID", "TypeHAlignCenter", "ENCOUNTER_ID", HeaderColor
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 1, "1", "START DATE", "TypeHAlignCenter", "ENCNTR_START_DATE", HeaderColor
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 2, "1", "PATIENT TYPE", "TypeHAlignCenter", "ENCNTR_TYPE", HeaderColor
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 3, "1", "REASON VISIT", "TypeHAlignCenter", "VISIT_REASON", HeaderColor
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 4, "1", "VISIT TYPE", "TypeHAlignCenter", "VISIT_TYPE", HeaderColor
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 5, "1", "APPT TYPE", "TypeHAlignCenter", "ENCOUNT_APPT_TYPE", HeaderColor
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 6, "1", "CPT CODE", "TypeHAlignCenter", "ENCOUNT_CPT_INTRVN", HeaderColor
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 7, "1", "NEXT APPT", "TypeHAlignCenter", "NEXTAPPT", HeaderColor
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 8, "1", "PROVIDER", "TypeHAlignCenter", "PROVIDER_ID", HeaderColor
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 9, "1", "PLACE SERVICE", "TypeHAlignCenter", "SERVICE_PLACE", HeaderColor
        Format_CellTypeEdit grdEnctrSummary, SpreadHeader, 10, "1", "FACILITY", "TypeHAlignCenter", "FACILITY_ID", HeaderColor
        
        If .RowHeight(SpreadHeader) < 15 Then .RowHeight(SpreadHeader) = 15
        Add_Spacer SpreadHeader + 1, grdEnctrSummary
 'CellBorders
        .SetCellBorder 0, SpreadHeader, 0, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 1, SpreadHeader, 1, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 3, SpreadHeader, 3, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 4, SpreadHeader, 4, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 5, SpreadHeader, 5, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 6, SpreadHeader, 6, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 7, SpreadHeader, 7, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 8, SpreadHeader, 8, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 9, SpreadHeader, 9, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .SetCellBorder 10, SpreadHeader, 10, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngRow = 0
        If cnEnctr.State = adStateClosed Then cnEnctr.open psConnect(0)
        strSQL = "SELECT * FROM ENCOUNTER WHERE PATIENT_ID = " & piPatientID & " AND PURGE = 0 Order by  ENCNTR_START_DATE desc "
        Set rsEnctr = cnEnctr.Execute(strSQL)
        Do Until rsEnctr.EOF
            booSkip = False
            strEnctrVisit = vbNullString
            If Not IsNull(rsEnctr.Fields("ENCOUNT_CPT_INTRVN")) Then
                If Trim(rsEnctr.Fields("ENCOUNT_CPT_INTRVN")) <> vbNullString Then
                    Select Case rsEnctr.Fields("ENCOUNT_CPT_INTRVN")
                        Case "99361", "99373", "99442", "99443", "90000", "99455", "99456", "G0180", _
                            "99361", "99058", "99071", "99075", "99082", "99080", "99089", "99085", "99441"
                            booSkip = True
                    End Select
                End If
            End If
            If Not booSkip Then
                Set Flds = rsEnctr.Fields
                lngRow = lngRow + 1
                .MaxRows = lngRow
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString And Fld.Value <> "1/1/1900" Then
                            Select Case UCase(Fld.Name)
                                Case "ENCOUNTER_ID"
                                    Arr = Split(LookUp_ColHeader("ENCOUNTER_ID"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        .SetRowItemData (lngRow), Fld.Value
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", Fld.Name
                                    End If
                                Case "ENCNTR_START_DATE"
                                    Arr = Split(LookUp_ColHeader("ENCNTR_START_DATE"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        .ColWidth(lngCol) = 14
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", Format(Fld.Value, "MMMM DD, YYYY"), "TypeHAlignCenter", Fld.Name
                                    End If
                                Case "ENCNTR_END_DATE"
                                    Arr = Split(LookUp_ColHeader("ENCNTR_END_DATE"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        .ColWidth(lngCol) = 14
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", Format(Fld.Value, "MMMM DD YYYY"), "TypeHAlignCenter", Fld.Name
                                    End If
                                Case "ENCOUNT_APPT_TYPE"
                                    Arr = Split(LookUp_ColHeader("ENCOUNT_APPT_TYPE"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        .ColWidth(lngCol) = 10
                                        Get_Appointment_Types
                                        Format_CellTypeCombo grdEnctrSummary, lngRow, lngCol, "1", clsGridFormat.List, "TypeHAlignCenter", Fld.Name
                                        Set_GridCurSel grdEnctrSummary, Fld.Value, lngCol, lngRow, "1", Fld.Name
                                    End If
                                Case "NEXTAPPT"
                                    Arr = Split(LookUp_ColHeader("NEXTAPPT"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", Fld.Name
                                    End If
                                Case "VISIT_REASON"
                                    Arr = Split(LookUp_ColHeader("VISIT_REASON"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", Fld.Name
                                    End If
                                Case "VISIT_TYPE"
                                    Arr = Split(LookUp_ColHeader("VISIT_TYPE"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        strEnctrVisit = Fld.Value
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", strEnctrVisit, "TypeHAlignCenter", Fld.Name
                                    End If
                                Case "FACILITY_ID"
                                    Arr = Split(LookUp_ColHeader("FACILITY_ID"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        Get_EnctFacName Fld.Value
                                        Get_Facilities
                                        clsEncounter.Facility = clsFacility.Name
                                        clsEncounter.FacilityID = Fld.Value
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", clsEncounter.Facility, "TypeHAlignCenter", Fld.Name
                                    End If
                                Case "PROVIDER_ID"
                                    Arr = Split(LookUp_ColHeader("PROVIDER_ID"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        Get_EnctrProviderName Fld.Value
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", clsEncounter.Provider, "TypeHAlignCenter", Fld.Name
                                    End If
                                Case "SERVICE_PLACE"
                                    Arr = Split(LookUp_ColHeader("SERVICE_PLACE"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", Fld.Name
                                    End If
                                Case "ENCOUNT_CPT_INTRVN"
                                    Arr = Split(LookUp_ColHeader("ENCOUNT_CPT_INTRVN"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                    End If
                                    .ColWidth(lngCol) = 14
                                    If Fld.Value = "049X-" Then
                                        If rsEnctr.Fields("VISIT_REASON") <> vbNullString Then
                                            clsOpNote.Preop_Code = vbNullString
                                            Get_Surgery_CPT rsEnctr.Fields("VISIT_REASON")
                                            If Trim(clsOpNote.Preop_Code) <> vbNullString Then
                                                strCPT = clsOpNote.Preop_Code
                                            End If
                                        End If
                                    Else
                                        strCPT = Fld.Value
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", strCPT, "TypeHAlignCenter", Fld.Name
                                    End If
                                Case "ENCNTR_TYPE"
                                    Arr = Split(LookUp_ColHeader("ENCNTR_TYPE"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngCol = Arr(0)
                                        Format_CellTypeEdit grdEnctrSummary, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", Fld.Name
                                    End If
                            End Select
                        End If
                    End If
                Next
            End If
            rsEnctr.MoveNext
            Arr = Split(LookUp_ColHeader("ENCOUNT_APPT_TYPE"), "|")
            If UBound(Arr) > 0 Then
                lngCol = Arr(0)
            End If
            If lngCol > 0 Then
                grdEnctrSummary.GetText lngCol, lngRow, varCellValue
                If Trim(varCellValue) = vbNullString And Trim(strEnctrVisit) <> vbNullString Then
                    Set_GridCurSel grdEnctrSummary, Get_ApptCode(strEnctrVisit), lngCol, lngRow, "1", "ENCOUNT_APPT_TYPE"
                End If
            End If
            
            Clear_Provider
            .SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 6, lngRow, 6, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 7, lngRow, 7, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 8, lngRow, 8, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 9, lngRow, 9, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .SetCellBorder 10, lngRow, 10, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Loop
        .MaxRows = lngRow
    End With
    If cnEnctr.State = adStateOpen Then
        cnEnctr.Close
        Set cnEnctr = Nothing
        Set rsEnctr = Nothing
    End If
    If piEncounterID <> 0 Then Set_CurrentEncounter
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Init_Summary:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, clsPatient.Name, "Init_Summary", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Function LookUp_ColHeader(vstrname As String) As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim booExit As Boolean

    On Error GoTo Error_LookUp_ColHeader
    Screen.MousePointer = vbHourglass
    With grdEnctrSummary
        LookUp_ColHeader = vbNullString
        For lngCol = 0 To .MaxCols
                .Col = lngCol
                .Row = SpreadHeader
                If .CellTag <> vbNullString Then
                    varCellValue = .CellTag
                    If Trim(varCellValue) <> vbNullString And TextFound(vstrname, varCellValue, Len(vstrname), True) Then
                         LookUp_ColHeader = lngCol & "|" & lngRow
                         booExit = True
                         Exit For
                    End If
                End If
        Next lngCol
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUp_ColHeader:
    Screen.MousePointer = vbDefault
    Exit Function
    Resume
End Function


Public Sub Reset_SummaryGrid()
    grdEnctrSummary.ReDraw = False
    grdEnctrSummary.MaxRows = 0
    grdEnctrSummary.ClearRange 0, 0, grdEnctrSummary.MaxCols, grdEnctrSummary.MaxRows, False
    grdEnctrSummary.ReDraw = True
End Sub


Private Sub cmdUpdProc_Click()
    txtVisitReason.Text = vbNullString
    cctProcString.Text = vbNullString
    fraEnctrProc.Visible = True
    fraEnctrProc.ZOrder 0
    fraEnctr.Enabled = False
    Exit Sub
End Sub

Private Sub cmdVisitReas_Click()
    frmNurseTemplates.fCaption = "Visit Reason"
    frmNurseTemplates.fType = "Visit Reason"
    frmNurseTemplates.fSubDomain = "sd_PostSurg_VisitType"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
   GetList cboVisitReason, "sd_PostSurg_VisitType"
End Sub

Private Sub Form_Activate()
Dim i%

    On Error GoTo Error_Activate
    Screen.MousePointer = vbHourglass
    mstrRendType = "Referral Physician"
    cctReferral.Text = vbNullString
    cctRend.Text = vbNullString
    cctRefReason.Text = vbNullString
    lstReferral.Visible = False
    lstReferral.Clear
    lstRend.Clear
    lstRend.Visible = False
    If mFirst Then
       mFirst = False
        Clear_Enctr
        modSpecific.piEncounterID = 0
        Clear_ReferralStrings
        mstrPatOcx = vbNullString
        mstrPatCtl = vbNullString
        Screen.MousePointer = vbHourglass
        cboVisitReason.Visible = False
        txtVisitReason.Visible = True
        txtVisitReason.Text = vbNullString
        mstrColCPT = vbNullString
        mstrColName = vbNullString
        mstrProc = vbNullString
        mlngProcCol = 0
        mlngProSubID = 0
        mlngProcID = 0
        mlngProcRow = 0
        mlngProcCol = 0
        cctProcString.Text = vbNullString
        cctProcString.Tag = vbNullString
        txtVisitReason.ZOrder 0
        mbDischarge = False
        If mstrPractice = "Psychiatry" Then
            cmdUpdProc.Visible = False
            cmdVisitReas.Visible = False
            cmdSurgLog.Visible = False
        End If
        mbAdmit = False
        Dump_EncounterStrData
        Get_Patient piPatientID
        Clear_Provider
        If PsUserID = vbNullString Then PsUserID = Get_PsUserID(psConnect(0))
        Get_Providers mlngLogonUser
        lblUserName.Caption = clsProvider.Name
        If clsPatient.Name <> vbNullString Then
                Me.Caption = "for" & " " & clsPatient.Name & " " & "Date of Birth" & " " & Format(clsPatient.DOB, "mmmm dd, yyyy")
            Else
                Get_Patient piPatientID
                 Me.Caption = "for" & " " & clsPatient.Name & " " & "Date of Birth" & " " & Format(clsPatient.DOB, "mmmm dd, yyyy")
            End If
        mEnctrCounter = 0
        piPatientID = frmPatients.piPatientID
        modSpecific.piPatientID = 0
        modSpecific.piPatientID = piPatientID
        Get_Patient piPatientID
        psImageFilePath = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ImagePath")
        If psImageFilePath = vbNullString Then
           psImageFilePath = "C:\"
        End If
        If piPatientID <> 0 Then
            Me.Caption = "NAME:" & " " & clsPatient.Name & " " & "EMR ID" & " " & clsPatient.IntId & " " & "BIRTH" & " " & Format(clsPatient.DOB, "mmmm dd, yyyy")
        End If
        Set_Options
        Screen.MousePointer = vbDefault
        Load_Providers
        Load_Facilities
        mbDateCheck = False
        mlngPrevEnctrID = 0
        Validate_Enctr
        Init_Summary
        Enctr_Buttons
        clsOpNote.Procedure_Name = vbNullString
        clsOpNote.Procedure_ID = vbNullString
        centerForm Me
    End If
  Screen.MousePointer = vbDefault
  Exit Sub
Error_Activate:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Activate", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Public Sub Form_Load()
    On Error GoTo Error_Handler
    mFirst = True
    Exit Sub
Error_Handler:
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
End Sub


Private Sub Form_Paint()

    On Error GoTo ERROR
    
    ColorBurst Me, "BLUE", 128, 0
    centerForm Me
    
ERROR:
    Exit Sub
    Resume

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
    On Error Resume Next
    If IsFormLoaded("frmEnctr") Then Unload frmEnctr
    Set frmEnctr = Nothing
    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
  Exit Sub
End Sub


Private Sub grdColumn_Click(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue
    mlngColumnCol = Col
    mlngColumnRow = Row
    grdProc.Enabled = False
   With grdColumn
        .BlockMode = True
            .Col = mlngColumnCol
            .Col2 = mlngColumnCol
            .Row = mlngColumnRow
            .Row2 = mlngColumnRow
            .BackColor = grdCurItem
            .EditMode = False
            .SetCellBorder .Col, .Row, .Col, .Row, CellBorderIndexOutline, 0, CellBorderStyleSolid
             .GetText mlngColumnCol, SpreadHeader, varCellValue
        .BlockMode = False
    End With
    grdProc.Enabled = True
    Exit Sub
End Sub


Private Sub grdColumn_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
Dim iRow%
    mlngColumnRow = 0
    mlngColumnRow = Row
    mlngColumnCol = Col
    If mlngColumnRow = 0 Then Exit Sub
    With grdColumn
        .ReDraw = False
            .BlockMode = True
                .Col = mlngColumnCol
                .Col2 = mlngColumnCol
                .Row = mlngColumnRow
                .Row2 = mlngColumnRow
                .BackColor = NormalColor
            .BlockMode = False
        .ReDraw = True
    End With
    Exit Sub
End Sub


Private Sub grdEnctrSummary_Click(ByVal Col As Long, ByVal Row As Long)
Dim lngCol As Long
Dim strTag As String
    mlngGridRow = Row
    With grdEnctrSummary
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then strTag = .CellTag
        If strTag = "ENCOUNT_APPT_TYPE" Then
            Get_Appointment_Types
            Format_CellTypeCombo grdEnctrSummary, Row, lngCol, "1", clsGridFormat.List, "TypeHAlignCenter", "ENCOUNT_APPT_TYPE"
            .SetCellBorder 0, Row, 0, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
    End With
    Exit Sub
End Sub

Private Sub grdEnctrSummary_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
Dim i%
    With grdEnctrSummary
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString And .CellTag = "ENCOUNT_APPT_TYPE" Then
            If clsEncounter.VisitType <> vbNullString Then
                For i = 0 To cboPurVisit.ListCount - 1
                    If cboPurVisit.List(i) = clsEncounter.VisitType Then
                        mbooByPass = True
                        cboPurVisit.Text = cboPurVisit.List(i)
                        mbooByPass = False
                        Exit For
                    End If
                Next i
            End If
        End If
    End With
    Exit Sub
End Sub

Private Sub grdEnctrSummary_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim i%
Dim lngEnctrID As Long
Dim varStartDate
Dim strSQL As String
Dim lngrecs As Long
Dim booRet As Boolean
Dim strUser As String
Dim varRS
Dim lngProviderID As Long
Dim strStart As String
Dim strEnd As String
Dim varFacility
Dim varProvider
Dim varService
Dim VarType
Dim varEndDate
Dim varReason
Dim varVisit
Dim lngCol As Long
Dim varCellValue
Dim Idx As Long
Dim booSkip As Boolean
Dim Arr
Dim lngRow As Long

    If mbooByPass Then mbooByPass = False: Exit Sub
    On Error GoTo Error_grdEnctrSummary_DblClick
    Screen.MousePointer = vbHourglass
    mbooByPass = True
    Clear_Super_Provider
    cctReferral.Text = vbNullString
    cctRend.Text = vbNullString
    cctRefReason.Text = vbNullString
    cctRefReason.Visible = False
    mbooByPass = True
    lstReferral.Visible = False
    mbooByPass = True
    lstReferral.Clear
    mbooByPass = True
    lstRend.Clear
    mbooByPass = True
    cboPurVisit.ListIndex = -1
    mbooByPass = True
    lstRend.Visible = False
    clsEncounter.StartDate = vbNullString
    modSpecific.piEncounterID = 0
    cctDiagnosis.Text = vbNullString
    cboVisitReason.Visible = False
    txtVisitReason.Visible = True
    txtVisitReason.ZOrder 0
    Me.Caption = "for" & " " & clsPatient.Name & " " & "Date of Birth" & " " & Format(clsPatient.DOB, "MMMM dd, yyyy")
    mbooGrid = True
    mlngGridRow = Row
    mlngGridCol = Col
    With grdEnctrSummary
        .ReDraw = False
        If cboProvider.ListCount < 0 Then Load_Providers
        If cboFacility.ListCount < 0 Then Load_Facilities
        lngEnctrID = 0
        If gbooBypass = False Then
            Dump_EncounterStrData
            Reset_Encounter
        End If
        If mlngGridRow >= 1 Then
            For lngRow = 1 To .MaxRows
                .BlockMode = True
                    .Col = 0
                    .Col2 = .MaxCols
                    .Row = lngRow
                    .Row2 = lngRow
                    .FontBold = True
                    .BackColor = &HFFFFFF
                .BlockMode = False
            Next lngRow
        End If
        If mlngGridRow > 0 Then
            .BlockMode = True
                .Col = 0
                .Col2 = .MaxCols
                .Row = mlngGridRow
                .Row2 = mlngGridRow
                .FontBold = True
                .BackColor = &HC0E0FF
            .BlockMode = False
            lngEnctrID = .GetRowItemData(mlngGridRow)
            If lngEnctrID > 0 Then
                piEncounterID = lngEnctrID
                Get_Encounter lngEnctrID
'                If clsEncounter.ProviderID <> mlngLogonUser Then
'                    If clsEncounter.RendProv <> mlngLogonUser Then
'                        Clear_PracticeRole
'                        Get_SecurityRole mlngLogonUser
'                        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
'                            MsgBox "This is not your Patient Encounter! Please contact" _
'                                & " " & clsEncounter.Provider & " " & "for access permission or use Patient" _
'                                & "History to review this visit data. Thank you!", vbInformation, "PatientTrac Security Management"
'                                Screen.MousePointer = vbDefault
'                                Exit Sub
'                        End If
'                    End If
'                End If
                mlngPrevEnctrID = 0
            End If
            For i = 0 To cboFacility.ListCount - 1
                If cboFacility.ItemData(i) = clsEncounter.FacilityID Then
                    mbooByPass = False
                    cboFacility.Text = cboFacility.List(i)
                    Exit For
                End If
            Next i
            For i = 0 To cboProvider.ListCount - 1
                If clsEncounter.ProviderID > 0 Then
                    If cboProvider.ItemData(i) = clsEncounter.ProviderID Then
                        cboProvider.Text = cboProvider.List(i)
                        Exit For
                    End If
                Else
                    If cboProvider.ItemData(i) = clsEncounter.ProviderID Then
                        cboProvider.Text = cboProvider.List(i)
                        Exit For
                    End If
                End If
            Next i
            Load_Service clsEncounter.FacilityID
            For i = 0 To cboSerPlace.ListCount - 1
                If clsEncounter.ServicePlaceCode > 0 Then
                    If cboSerPlace.ItemData(i) = clsEncounter.ServicePlaceCode Then
                        cboSerPlace.Text = cboSerPlace.List(i)
                        Exit For
                    End If
                End If
            Next i
            For i = 0 To cboStatus.ListCount - 1
                If cboStatus.List(i) = clsEncounter.EncounterType Then
                    If mbooGridClick Then mbooByPass = True
                    cboStatus.Text = cboStatus.List(i)
                    Exit For
                End If
            Next i
            mbooByPass = False
            txtVisitReason.Text = vbNullString
            If Trim(clsEncounter.VisitReason) <> vbNullString Then txtVisitReason.Text = clsEncounter.VisitReason
            If cboStatus.Text <> "Procedure" Then
                For i = 0 To cboPurVisit.ListCount - 1
                    If cboPurVisit.List(i) = clsEncounter.VisitType Then
                        cboPurVisit.Text = cboPurVisit.List(i)
                        Exit For
                    End If
                Next i
            End If
        End If
    End With
    cctRend.Text = vbNullString
    If Val(clsEncounter.RendProv) > 0 Then
       Load_Specific_Rendering Val(clsEncounter.RendProv)
    End If
    If clsEncounter.RendProv <> vbNullString And Trim(cctRend.Text) = vbNullString Then
        cctRend.Text = vbNullString
        For i = 0 To lstRend.ListCount - 1
            If lstRend.ItemData(i) = clsEncounter.RendProv Then
                lstRend.Selected(i) = True
                Exit For
            End If
        Next i
        cctRend.Text = Trim(clsEncounter.Provider)
    End If
    Get_EnctrInfo
    If IsNumeric(lngEnctrID) And lngEnctrID > 0 Then
        Enctr_Buttons
        If lngEnctrID > 0 Then piEncounterID = lngEnctrID
        ccdStartDate.Enabled = True
        If clsEncounter.StartDate <> "1/1/1900" Then
            ccdStartDate.Value = Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM")
        End If
        If clsEncounter.EndDate <> vbNullString And clsEncounter.EndDate <> "1/1/1900" Then
            ccdEndDate.Value = Format(clsEncounter.EndDate, "mm/dd/yyyy hh:mm:ss AM/PM")
        Else
            ccdEndDate.Value = vbNullString
        End If
        If Not IsNull(ccdEndDate.Value) Then
            ccdStartDate.Enabled = False
            ccdEndDate.Enabled = False
            cmdUpdate.Enabled = False
        End If
    End If
    If clsEncounter.ServicePlace <> "Office" And clsEncounter.ServicePlaceCode <> "11" Then
        Get_AdmissionDate
    End If
    gbooBypass = False
    mbooGrid = False
    If cboPurVisit.Text = vbNullString And Not booSkip Then
            If lngEnctrID > 0 Then Get_Encounter lngEnctrID
            If clsEncounter.VisitType <> vbNullString Then
                For i = 0 To cboPurVisit.ListCount - 1
                    If cboPurVisit.List(i) = clsEncounter.VisitType Then
                        mbooByPass = True
                        cboPurVisit.Text = cboPurVisit.List(i)
                        mbooByPass = False
                        Exit For
                    End If
                Next i
            Else
                MsgBox "PatientTrac requires Visit Type, Please Update Encounter with a Visit Type!", vbCritical
            End If
    Else
        clsEncounter.VisitType = Trim(cboPurVisit.Text)
    End If
    cmdUpdate.Enabled = True
    Get_ApptData clsEncounter.StartDate
    If clsEncounter.VisitType <> cboPurVisit.Text Then
        clsEncounter.VisitType = Trim(cboPurVisit.Text)
        LookUp_CPT clsEncounter.VisitType
        grdEnctrSummary.GetText 6, Row, varCellValue
        If Trim(varCellValue) <> CStr(clsEncounter.Intrvn_CD) Then
            Format_CellTypeEdit grdEnctrSummary, Row, 6, "1", CStr(clsEncounter.Intrvn_CD), "TypeHAlignCenter", "ENCOUNT_CPT_INTRVN", &HC0E0FF
        End If
    End If
    Arr = Split(LookUpColName("APPT TYPE"), "|")
    If UBound(Arr) > 0 Then
        lngCol = Arr(0)
        If mlngGridRow > 0 Then
            Get_Appointment_Types
            Format_CellTypeCombo grdEnctrSummary, mlngGridRow, lngCol, "1", clsGridFormat.List, "TypeHAlignCenter", "ENCOUNT_APPT_TYPE"
            If clsEncounter.Appt_Type <> vbNullString Then Set_GridCurSel grdEnctrSummary, clsEncounter.Appt_Type, lngCol, mlngGridRow, "1", "ENCOUNT_APPT_TYPE"
        End If
    End If
    If Trim(txtVisitReason.Text) <> vbNullString Then
        fraEnctrProc.Visible = False
        fraEnctr.Enabled = True
        cmdUpdProc.Visible = True
    ElseIf Trim(txtVisitReason.Text) = vbNullString Then
        If TextFound("Procedure", cboStatus.Text, Len("Procedure"), True) Or _
            TextFound("Admission", cboStatus.Text, Len("Admission"), True) Or _
            TextFound("Recovery Room", cboStatus.Text, Len("Recovery Room"), True) And mstrPractice <> "Psychiatry" Then
            fraEnctrProc.Visible = True
            fraEnctrProc.ZOrder 0
            fraEnctr.Enabled = False
        End If
    End If
    piEncounterID = lngEnctrID
    modSpecific.piEncounterID = piEncounterID
    Get_Referrals lngEnctrID
    Select Case mstrPractice
        Case "Gyn", "IVF"
            Values_ObGyn "DIAGNOSIS"
            If Trim(clsReport.Diagnosis) <> vbNullString Then
                cctDiagnosis.Text = Trim(clsReport.Diagnosis)
            End If
        Case "Psychiatry"
            Values_Psychiatry "DIAGNOSIS"
            If Trim(clsReport.AXIS_I) <> vbNullString Then
                cctDiagnosis.Text = Trim(clsReport.AXIS_I)
            End If
    End Select
    mbooByPass = False
    cmdDeleteEncounter.Enabled = True
    frmMDI.StatusBar.Panels(4).Text = "Encounter: " & piEncounterID
    frmMDI.StatusBar.Panels(5).Text = "Visit Date: " & Format(clsEncounter.StartDate, "MMMM DD, YYYY")
    frmMDI.StatusBar.Panels(6).Text = "Provider: " & clsEncounter.Provider
    ccdEndDate.Enabled = True
    ccdStartDate.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_grdEnctrSummary_DblClick:
    MsgBox "One or more fields contain incorrect encounter information. Please correct and update record!", vbCritical
    Screen.MousePointer = vbDefault
     Write_Error Err.Description, Err.Number, CStr(lngRow), "grdEnctrSummary", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub grdEnctrSummary_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
       cmdEnctrOK_Click
    End If
End Sub


Private Sub Load_Facilities()
Dim strSQL As String
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset
Dim strFacName As String
Dim lngFacID As Long

    On Error GoTo Error_Load_Facilities
    Screen.MousePointer = vbHourglass
'-------------

    cboFacility.Clear
    With cboFacility
        .Columns = 1
        .ColumnHeaderShow = False
        .ColumnHeaderHeight = 315
        .EditHeight = 350
        .MaxDrop = 10
        .Row = -1
        .Col = 0
        .ColName = "FAC"
        .ColHeaderText = "FACILITY NAME"
        .WrapList = True
    End With
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "SELECT FACILITY_NAME, FACILITY_ID FROM FACILITIES WHERE ACTIVE = 1 ORDER BY FACILITY_NAME"
    Set rsFac = cnFac.Execute(strSQL)
    With rsFac
        Do Until .EOF
            '-------------
            If IsNumeric(!Facility_ID) Then lngFacID = (!Facility_ID)
            If Not IsNull((!Facility_Name)) Then
                strFacName = (!Facility_Name)
            End If
        .MoveNext
        strFacName = Fix_Name(strFacName)
        strFacName = Fix_Facility_Name(strFacName)
            If cboFacility.Row = -1 Then
                cboFacility.InsertRow = strFacName
                cboFacility.Row = cboFacility.NewIndex
                cboFacility.ItemData = lngFacID
                cboFacility.BorderStyle = BorderStyleSingleLine
            Else
                cboFacility.Row = cboFacility.Row + 1
                cboFacility.InsertRow = strFacName
                cboFacility.Row = cboFacility.NewIndex
                cboFacility.ItemData = lngFacID
            End If
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
    Write_Error Err.Description, Err.Number, strSQL, "Load_Facilities", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
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
Private Function Get_MaxColWidth(grdObjGrid As Object, vlngCol As Long, vstrText As String) As Long
Dim CellNote As Long
Dim Arr
Dim Idx As Integer
Dim strTest As String
Dim lngColWidth As Long

    Get_MaxColWidth = 0
    With grdObjGrid
        Arr = Split(vstrText, vbNewLine)
        If UBound(Arr) > 0 Then
            For Idx = LBound(Arr) To UBound(Arr)
                strTest = Arr(Idx)
                If strTest <> vbNullString Then
                  .BlockMode = True
                        .Col = 1
                        .Col2 = 1
                        .Row = 1
                        .Row2 = 1
                        .SetText 1, 1, strTest
                        .TypeEditMultiLine = False   'Make multi-lines
                        .Text = Trim(strTest)
                        CellNote = .MaxTextCellWidth
                        If CellNote > lngColWidth Then lngColWidth = CellNote + 2
                    .BlockMode = False
                End If
            Next Idx
        Else
            If vstrText <> vbNullString Then
              .BlockMode = True
                    .Col = 1
                    .Col2 = 1
                    .Row = 1
                    .Row2 = 1
                    .SetText 1, 1, vstrText
                    .TypeEditMultiLine = False   'Make multi-lines
                    .Text = Trim(vstrText)
                    CellNote = .MaxTextCellWidth
                    If CellNote > lngColWidth Then lngColWidth = CellNote + 2
                .BlockMode = False
            End If
        End If
        If lngColWidth > 0 Then Get_MaxColWidth = lngColWidth
    End With
End Function
Private Function Get_BodySpecProcedure(vstrBodyArea As String)
Dim strSQL As String
Dim cnSurg As New ADODB.Connection
Dim rsSurg As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim CellNote As Long

    On Error GoTo ERROR_Get_BodySpecProcedure
    If cnSurg.State = adStateClosed Then cnSurg.open psConnect(1)
    With grdProc
        .ReDraw = False
        .ClearRange 1, 1, .MaxCols, .MaxRows, False
        .MaxRows = 0
        
    strSQL = "SELECT PROCEDURE_NAME, PROCEDURE_ID, CPT_CODE FROM SURG_PROCEDURE_CODES WHERE SPECIALTY = '" & mstrPractice & "' AND BODY_AREA = '" & vstrBodyArea & "' ORDER BY PROCEDURE_NAME"
    Set rsSurg = cnSurg.Execute(strSQL)
        Do Until rsSurg.EOF
            Set Flds = rsSurg.Fields
            For Each Fld In Flds
                If Fld.Name = "PROCEDURE_NAME" Then
                    lngRow = lngRow + 1
                    .MaxRows = lngRow
                      .BlockMode = True
                        .Col = 1
                        .Col2 = 1
                        .Row = lngRow
                        .Row2 = lngRow
                        .CellType = CellTypeStaticText
                         .SetText 1, lngRow, IsProperCase(Fld.Value)
                        .TypeMaxEditLen = 32000
                        .TypeEditMultiLine = True   'Make multi-lines
                        .Text = IsProperCase(Fld.Value)
                        CellNote = .MaxTextCellHeight
                        .RowHeight(lngRow) = CellNote
                        .TypeVAlign = TypeVAlignCenter
                        .TypeHAlign = TypeHAlignLeft
                    .BlockMode = False
                   
                ElseIf Fld.Name = "CPT_CODE" Then
                      .BlockMode = True
                        .Col = 2
                        .Col2 = 2
                        .Row = lngRow
                        .Row2 = lngRow
                        .CellType = CellTypeStaticText
                        .SetText 2, lngRow, Format(Fld.Value, "00000")
                        .Text = Format(Fld.Value, "00000")
                        .TypeVAlign = TypeVAlignCenter
                        .TypeHAlign = TypeHAlignCenter
                    .BlockMode = False
                    
                ElseIf Fld.Name = "PROCEDURE_ID" Then
                    .SetRowItemData (lngRow), Fld.Value
                End If
            Next
            rsSurg.MoveNext
        Loop

        .BlockMode = True
            .Col = 1
            .Col2 = 1
            .Row = SpreadHeader + 1
            .Row2 = SpreadHeader + 1
            .RowHeight(SpreadHeader + 1) = 3
            .BackColor = 12632256
            .AddCellSpan 1, SpreadHeader + 1, .MaxCols, 1
            .SetCellBorder 1, SpreadHeader + 1, .MaxCols, SpreadHeader + 1, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
        .ReDraw = True
    End With
    If cnSurg.State = adStateOpen Then
        cnSurg.Close
        Set cnSurg = Nothing
        Set rsSurg = Nothing
    End If
    Exit Function
ERROR_Get_BodySpecProcedure:
    Write_Error Err.Description, Err.Number, strSQL, "Get_BodySpecProcedure", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Get_ChildTerms()
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim strSQL As String
Dim iRow As Long
Dim varCellValue
Dim iCol As Long
Dim varColName
Dim CellNote As Long
Dim strItem As String
Dim strSeperator As String
Dim strItemData As String
Dim strColValue As String

    On Error GoTo Error_Get_ChildTerms
    With grdColumn
        For iCol = 1 To .MaxCols
            mstrColName = vbNullString
            .GetText iCol, SpreadHeader, varColName
            If Trim(varColName) <> vbNullString Then mstrColName = Trim(varColName)
                If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
                strSQL = "SELECT * FROM SURG_CHILD_TERMS WHERE COLUMN_NAME = '" & mstrColName & "'  AND PROC_SUBGROUP_ID = " & mlngProSubID & " ORDER BY COLUMN_ITEMDATA"
                Set rsTerms = cnTerms.Execute(strSQL)
                Do Until rsTerms.EOF
                    iRow = iRow + 1
                    If .MaxRows < iRow Then .MaxRows = iRow
                    .SetRowItemData iRow, rsTerms.Fields("SURG_TERM_ID")
                    strItem = vbNullString
                    strItemData = vbNullString
                    If rsTerms.Fields("COLUMN_ITEMDATA") <> vbNullString Then strItemData = rsTerms.Fields("COLUMN_ITEMDATA")
                    If rsTerms.Fields("COLUMN_ITEM") <> vbNullString Then strItem = rsTerms.Fields("COLUMN_ITEM")
                    If Trim(rsTerms.Fields("COLUMN_SEPERATOR")) <> vbNullString Then strSeperator = rsTerms.Fields("COLUMN_SEPERATOR")
                    If Trim(strItemData) <> vbNullString Then Format_CellTypeEdit grdColumn, iRow, iCol, "1", strItem, "TypeVAlignCenter", strItemData
                    rsTerms.MoveNext
                Loop
            iRow = 0
            .GetText iCol, SpreadHeader, varCellValue
            If Trim(varCellValue) <> vbNullString Then strColValue = Trim(varCellValue)
            If Trim(strSeperator) <> vbNullString Then Format_CellTypeEdit grdColumn, SpreadHeader, iCol, "1", strColValue, "TypeVAlignCenter", strSeperator
            strSeperator = vbNullString
        Next iCol
        If cnTerms.State = adStateOpen Then
            cnTerms.Close
            Set cnTerms = Nothing
            Set rsTerms = Nothing
        End If
        .ReDraw = True
    End With
Exit Function
Error_Get_ChildTerms:
    Write_Error Err.Description, Err.Number, strSQL, "Get_ChildTerms", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Load_Questions()
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim strSQL As String
Dim iRow As Long
Dim iCol As Long
Dim varCellValue
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strName As String
Dim lngRow As Long
Dim varColName

    On Error GoTo ERROR_Load_Questions

    With grdColumn
        .ClearRange 1, 1, .MaxCols, .MaxRows, False
        .MaxRows = 0
        mlngProSubID = 0
        If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
        strSQL = "SELECT * FROM SURGERY_CHILD_CPT WHERE PROCEDURE_ID = " & mlngProcID
        Set rsTerms = cnTerms.Execute(strSQL)
        If Not rsTerms.EOF Then
            mlngProSubID = rsTerms.Fields("PROC_SUBGROUP_ID")
            .SetRowItemData (1), mlngProSubID
            Set Flds = rsTerms.Fields
            For Each Fld In Flds
                If Fld.Name <> vbNullString Then
                    strName = Fld.Name
                    If strName = "FIRST_COLUMN" Then
                        lngRow = 1
                        .Col = 1
                        .Row = lngRow
                        .SetText 0, lngRow, Replace(Fld.Name, "_", " ")
                        .ColWidth(0) = 15
                        If Trim(Fld.Value) <> vbNullString Then .SetText 1, lngRow, Fld.Value
                        .ColWidth(1) = 20
                        .CellTag = Fld.Name
                    ElseIf strName = "SECOND_COLUMN" Then
                        lngRow = 2
                        .Col = 1
                        .Row = lngRow
                        .SetText 0, lngRow, Replace(Fld.Name, "_", " ")
                        .ColWidth(0) = 15
                        If Trim(Fld.Value) <> vbNullString Then .SetText 1, lngRow, Fld.Value
                        .ColWidth(1) = 20
                        .CellTag = Fld.Name
                    ElseIf strName = "THIRD_COLUMN" Then
                        lngRow = 3
                        .Col = 1
                        .Row = lngRow
                        .SetText 0, lngRow, Replace(Fld.Name, "_", " ")
                        .ColWidth(0) = 15
                        If Trim(Fld.Value) <> vbNullString Then .SetText 1, lngRow, Fld.Value
                        .ColWidth(1) = 20
                        .CellTag = Fld.Name
                    ElseIf strName = "FOURTH_COLUMN" Then
                        lngRow = 4
                        .Col = 1
                        .Row = lngRow
                        .SetText 0, lngRow, Replace(Fld.Name, "_", " ")
                        .ColWidth(0) = 15
                        If Trim(Fld.Value) <> vbNullString Then .SetText 1, lngRow, Fld.Value
                        .ColWidth(1) = 20
                        .CellTag = Fld.Name
                    ElseIf strName = "FIFTH_COLUMN" Then
                        lngRow = 5
                        .Col = 1
                        .Row = lngRow
                        .SetText 0, lngRow, Replace(Fld.Name, "_", " ")
                        .ColWidth(0) = 15
                        If Trim(Fld.Value) <> vbNullString Then .SetText 1, lngRow, Fld.Value
                        .ColWidth(1) = 20
                        .CellTag = Fld.Name
                    ElseIf strName = "SIXTH_COLUMN" Then
                        lngRow = 5
                        .Col = 1
                        .Row = lngRow
                        .SetText 0, lngRow, Replace(Fld.Name, "_", " ")
                        .ColWidth(0) = 15
                        If Trim(Fld.Value) <> vbNullString Then .SetText 1, lngRow, Fld.Value
                        .ColWidth(1) = 20
                        .CellTag = Fld.Name
                    End If
                End If
            Next
        End If


        For iRow = 1 To .MaxRows
            For iCol = 0 To 1
            .GetText iCol, iRow, varCellValue
                  .BlockMode = True
                    .Col = iCol
                    .Col2 = iCol
                    .Row = iRow
                    .Row2 = iRow
                    .CellType = CellTypeStaticText
                    .TypeMaxEditLen = 32000
                    .TypeEditMultiLine = True   'Make multi-lines
                    .Text = varCellValue
                    .TypeVAlign = TypeVAlignCenter
                    .TypeHAlign = TypeHAlignLeft
                .BlockMode = False
            Next iCol
        Next iRow
        If cnTerms.State = adStateOpen Then
            cnTerms.Close
            Set cnTerms = Nothing
            Set rsTerms = Nothing
        End If

        For iCol = 1 To .MaxCols
            .GetText iCol, SpreadHeader, varColName
            If Trim(varColName) = vbNullString Then
                .Col = iCol
                .ColHidden = True
            End If
        Next iCol
        .ReDraw = True
    End With
    Exit Function
ERROR_Load_Questions:
    Write_Error Err.Description, Err.Number, strSQL, "Load_Questions", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Get_Fat_Lipo_Details()
Dim lngPrimKey As Long
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim cnDefault As New ADODB.Connection
Dim rsDefault As New ADODB.Recordset
Dim strSQL As String
Dim iCol As Long
Dim iRow As Long
Dim strLipoArea As String
Dim varCellValue
Dim varColName
Dim lngCnt As Long
Dim Arr
Dim strCPT As String
Dim Idx As Integer

    mbooByPass = True
    mbooHasChanged = True
    On Error GoTo Error_Get_LipoDetails
    'clsEncounter.Intrvn_CD
    If clsOpNote.Procedure_Name <> vbNullString Then
        If TextFound("LIPO", clsOpNote.Procedure_Name, Len("LIPO"), True) Or _
            TextFound("FAT", clsOpNote.Procedure_Name, Len("FAT"), True) Then
                Arr = Split(clsEncounter.Intrvn_CD, vbNewLine)
                If UBound(Arr) > 0 Then
                    For Idx = LBound(Arr) To UBound(Arr)
                        strCPT = Trim(Arr(Idx))
                        If Len(strCPT) > 7 Then
                            LookUp_ChildCPT ("strCPT")
                        End If
                    Next Idx
                End If
        End If
    End If
    mbooByPass = False
    mbooHasChanged = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_LipoDetails:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_LipoDetails", App.EXEName & "." & Me.Name
    Exit Function
    Resume
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
            Else
                .TypeComboBoxCurSel = -1
                .TypeComboBoxIndex = -1
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
    .ReDraw = True
    End With
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
    If cboProvider.ListCount > 0 Then Exit Function
    cboProvider.Clear
    If cnProv.State = adStateClosed Then cnProv.open psConnect(0)
    strSQL = "SELECT * FROM  PROVIDERS a, FACILITY_PROVIDERS b  WHERE a.ACTIVE = 1 " _
        & "AND a.PROVIDER_ID = b.PROVIDER_ID AND b.FACILITY_ID = " & piFacilityID & " ORDER BY a.LAST_NAME, a.FIRST_NAME"
    Set rsProv = cnProv.Execute(strSQL)
    With rsProv
        Do Until .EOF
            strFirst = vbNullString
            strLast = vbNullString
            strMiddle = vbNullString
            strTitle = vbNullString
            strSuffix = vbNullString
            sName = vbNullString
            Select Case UCase((!Title))
                Case "MD", "RN", "PHD", "PH.D", "LSW", "Clinician", "LCSW", "APN", "LPN", "LGSW", "MSW", "LGPC", "LCPC", "LCS", _
                        "PA", "APRN", "NP", "MA", "MS", "DO", "D.O.", "PA-C", "LPC", "LCSW-C", "LCSWC", "PSY-D", "CRNP", "LCADC", "LCAD-C", "MPH", "ICMS"
                        
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
            End Select
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
                cboProvider.AddItem sName
                cboProvider.ItemData(cboProvider.NewIndex) = lngProvID
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
    Resume
End Function
Private Sub Load_RenProv()
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

    On Error GoTo Error_Load_RenProv
    lstRend.Clear
    Open_DB psConnect(0)
    strSQL = "SELECT * FROM PROVIDERS WHERE ACTIVE = 1 ORDER BY LAST_NAME, FIRST_NAME"
    Set rs = cn.Execute(strSQL)
    With rs
        Do Until .EOF
            strFirst = vbNullString
            strLast = vbNullString
            strMiddle = vbNullString
            strTitle = vbNullString
            strSuffix = vbNullString
            lngProvID = 0
            sName = vbNullString
            If Not IsNull((!Title)) Then
                strTitle = Replace((!Title), ".", "")
            End If
            If Trim(strTitle) <> vbNullString Then
                Select Case UCase(strTitle)
                    Case "MD", "RN", "PHD", "PH.D", "LSW", "Clinician", "LCSW", "APN", "LPN", "LGSW", "MSW", "LGPC", "LCPC", "LCS", _
                        "PA", "APRN", "NP", "MA", "MS", "DO", "D.O.", "PA-C", "LPC", "LCSW-C", "LCSWC", "PSY-D", "CRNP", "LCADC", "LCAD-C", "MPH", "ICMS"
                        lngProvID = 0
                        Set Flds = .Fields
                        For Each Fld In Flds
                            If Not IsNull(Fld.Value) Then
                                If Fld.Value <> vbNullString Then
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
                End Select
            End If
            .MoveNext
            strLast = Fix_Name(strLast)
            strFirst = Fix_Name(strFirst)
            strMiddle = Fix_Name(strMiddle)
            If Trim(strFirst) <> vbNullString Then
                 sName = strFirst
             End If
            If Trim(strMiddle) <> vbNullString Then
                 sName = sName & " " & Left(strMiddle, 1)
             End If
            If Trim(strLast) <> vbNullString Then
                 sName = sName & " " & strLast
             End If
            If Trim(strSuffix) <> vbNullString Then
                 sName = sName & ", " & Trim(strSuffix)
             End If
            If Trim(strTitle) <> vbNullString Then
                 sName = sName & ", " & Trim(strTitle)
             End If
            If sName <> vbNullString Then
                lstRend.AddItem sName
                lstRend.ItemData(lstRend.NewIndex) = lngProvID
            End If
        Loop
     End With
    Close_DB psConnect(0)
   Exit Sub
Error_Load_RenProv:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_RenProv", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub Load_Service(vlngFacilityID As Long, Optional vlngServCode As Integer)
Dim strSQL As String
Dim iRows As Long
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset
Dim cnCPT As New ADODB.Connection
Dim rsCPT As New ADODB.Recordset
Dim strServPlaceCd As String
Dim strPos As String
Dim i%
Dim lngFacilityID As Long
Dim Arr
Dim ArrCode
Dim varVisit
Dim strCode As String

        On Error GoTo Error_Load_Service
        Screen.MousePointer = vbHourglass
        cboSerPlace.Clear
        If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
        strPos = vbNullString
        strSQL = "Select  POS FROM  FACILITIES  WHERE FACILITY_ID = " & vlngFacilityID
        Set rsFac = cnFac.Execute(strSQL)
        With rsFac
            If Not .EOF Then
                If Not IsNull((!pos)) Then
                    If (!pos) <> vbNullString Then strPos = (!pos)
                End If
'                If Not IsNull((!SEC_POS)) Then
'                   If (!SEC_POS) <> vbNullString And Not TextFound((!SEC_POS), strPos, Len((!SEC_POS)), True) Then strPos = Trim(strPos) & "~" & (!SEC_POS)
'                End If
'                If Not IsNull((!THIRD_POS)) Then
'                   If (!THIRD_POS) <> vbNullString And Not TextFound((!THIRD_POS), strPos, Len((!THIRD_POS)), True) Then strPos = Trim(strPos) & "~" & (!THIRD_POS)
'                End If
           End If
        End With
        If strPos <> vbNullString Then
'            Arr = Split(strPos, "~")
'            If UBound(Arr) > 0 Then
'                For i = LBound(Arr) To UBound(Arr)
'                    If strCode <> Arr(i) Then
'                        strCode = Arr(i)
'                        strSQL = "SELECT DISTINCT POS_NAME FROM PLACE_SERVICE WHERE  POS_CODE  = " & Val(Arr(i))
'                        Set rsFac = cnFac.Execute(strSQL)
'                        If Not rsFac.EOF Then
'                            If Not IsNull(rsFac.Fields("POS_NAME")) Then
'                                If rsFac.Fields("POS_NAME") <> vbNullString And Not TextFound((rsFac.Fields("POS_NAME")), strServPlaceCd, Len((rsFac.Fields("POS_NAME"))), True) Then
'                                    If strServPlaceCd = vbNullString Then
'                                        strServPlaceCd = rsFac.Fields("POS_NAME") & "-" & Val(Arr(i))
'                                    Else
'                                        strServPlaceCd = strServPlaceCd & "~" & rsFac.Fields("POS_NAME") & "-" & Val(Arr(i))
'                                    End If
'                                End If
'                            End If
'                        Else
'                           If Not IsNull(rsFac.Fields("POS_NAME")) Then
'                                If rsFac.Fields("POS_NAME") <> vbNullString And _
'                                    Not TextFound((rsFac.Fields("POS_NAME")), strServPlaceCd, Len((rsFac.Fields("POS_NAME"))), True) Then
'                                    strServPlaceCd = Trim(strServPlaceCd) & "~" & rsFac.Fields("POS_NAME") & "-" & Val(Arr(i))
'                                End If
'                            End If
'                        End If
'                    End If
'
'                Next i
         '   Else
            If strPos <> vbNullString Then
                strSQL = "SELECT DISTINCT POS_NAME FROM PLACE_SERVICE WHERE  POS_CODE  = " & Val(strPos)
                Set rsFac = cnFac.Execute(strSQL)
                If Not rsFac.EOF Then
                    If Not IsNull(rsFac.Fields("POS_NAME")) Then cboSerPlace.AddItem rsFac.Fields("POS_NAME")
                    cboSerPlace.ItemData(cboSerPlace.NewIndex) = Val(strPos)
                End If
            End If

            For i = 0 To cboSerPlace.ListCount - 1
'                  If clsEncounter.EncounterType = "Subsequent Care" Then strPos = "Inpatient Psychiatric Facility"
'                  If clsEncounter.EncounterType = "Subsequent Consultation" Then strPos = "Inpatient Hospital"
                  If cboSerPlace.ItemData(i) = Val(strPos) Then
                      cboSerPlace.Text = cboSerPlace.List(i)
                      Exit For
                  End If
              Next i
''                End If
'            End If
        End If
        If cnFac.State = adStateOpen Then
            cnFac.Close
            Set cnFac = Nothing
        End If
    Exit Sub
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Service:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Service", App.EXEName & "." & TypeName(Me)
    Exit Sub
    Resume
End Sub


Public Sub Validate_Enctr()
Dim strSQL As String
Dim strStartTimeValue As String
Dim intAdddays As Integer
Dim strCurrentDate As String
Dim strCalDate As String
Dim iAnswer As Integer
Dim strValue As String
Dim sReg As String
Dim iReturn As Integer
Dim strRS
Dim lngDateDiff As Long

    On Error GoTo Error_Validate_Enctr
    If piPatientID = 0 Then Exit Sub
    If piProviderID = 0 Then
        Get_DefaultProvider
        piProviderID = modSpecific.piProviderID
        If piProviderID <> 0 Then
            mbooAddNewEnctr = True
            mstrDenyEncntrReas = "Default Provider Used"
        Else
            mbooAddNewEnctr = False
            MsgBox "Contact Network Administrator or Add Provider to this Patient Record- No Default Provider configured!", vbInformation
        End If
    Else
        mbooAddNewEnctr = True
    End If
    If piFacilityID = 0 Then
        psRegGeneral = "SOFTWARE\FirstNoteBook\PatientTrac\General"
        sReg = "DefaultFacility"
        strValue = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, sReg)
        If strValue <> vbNullString Then piFacilityID = Val(strValue)
        mbooAddNewEnctr = True: mstrDenyEncntrReas = "Default Provider Used"
    Else
        mbooAddNewEnctr = True
    End If
    If PatientInfo.mNewPatient = True Then
        PatientInfo.mNewPatient = False
        mbooAddNewEnctr = True
        Exit Sub
    End If
    If mbDateCheck Then
        strCurrentDate = Format(Now, "mm/dd/YYYY")
        Open_DB psConnect(0)
        strSQL = "select max(ENCNTR_START_DATE) from ENCOUNTER WHERE PATIENT_ID = " & piPatientID & " AND PURGE = 0"
        iReturn = GetRecordSet(strRS, strSQL, psConnect(0))
        If iReturn > 0 Then
            If strRS(0, 0) <> vbNullString Then
                strCalDate = Format(strRS(0, 0), "mm/dd/yyyy")
                lngDateDiff = DateDiff("d", strCalDate, strCurrentDate)
                If lngDateDiff = 0 Then
                   iAnswer = MsgBox("There is already an encounter for today's date. Do you still want to add another encounter?", vbYesNo)
                    If iAnswer = vbYes Then
                         mbooAddNewEnctr = True
                    Else
                         mbooAddNewEnctr = False
                    End If
                End If
            End If
        End If
        Close_DB psConnect(0)
    End If
    Screen.MousePointer = vbDefault
Exit Sub
Error_Validate_Enctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_Enctr", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Public Function WriteData(vstrAction As String) As Boolean
Dim strSQL As String
Dim lngTxnmy As Long
Dim strUser As String
Dim i As Integer
Dim strAppt As String
Dim strStart As String
Dim strEnd As String
Dim strVisitReason As String
Dim lstrRemarks As String
Dim lngEMID As Long
Dim varRS
Dim booRet As Boolean
Dim lngProviderID As Long
Dim lngFacilityID As Long
Dim sDate As String
Dim strFacilityName As String
Dim strNull As String
Dim iReturn As Integer
Dim Idx As Integer
Dim Arr
Dim lngCol As Long
Dim lngRendProv As Long
Dim cnEnctr As New ADODB.Connection
Dim rsEnctr As New ADODB.Recordset
Dim lngDateDiff As Long


    On Error GoTo Error_WriteData
    Screen.MousePointer = vbHourglass
    If clsEncounter.StartDate = vbNullString Then
         clsEncounter.StartDate = Format(ccdStartDate.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
    End If

    If clsEncounter.StartDate <> vbNullString Then
        sDate = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    Else
        Get_Encounter piEncounterID
        sDate = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    End If
    strUser = PsUserID
    If clsEncounter.VisitReason = vbNullString And Trim(txtVisitReason.Text) <> vbNullString Then
        clsEncounter.VisitReason = txtVisitReason.Text
    Else
        If Not TextFound(Trim(txtVisitReason.Text), clsEncounter.VisitReason, Len(txtVisitReason.Text), True) Then
            clsEncounter.VisitReason = clsEncounter.VisitReason & vbNewLine & txtVisitReason.Text
        End If
    End If
    If Not IsNull(ccdStartDate.Value) Then
        clsEncounter.StartDate = Format(ccdStartDate.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
    Else
        clsEncounter.StartDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    End If
    If Not IsNull(ccdEndDate.Value) Then
        clsEncounter.EndDate = Format(ccdEndDate.Value, "mm/dd/yyyy hh:mm:ss AM/PM")
    Else
        clsEncounter.EndDate = vbNullString
    End If
    If cboSerPlace.Text <> vbNullString Then
        clsEncounter.ServicePlace = cboSerPlace.Text
    End If
    If clsEncounter.ServicePlace = vbNullString Then
        clsEncounter.ServicePlace = clsFacility.FacType
    End If
    If clsEncounter.ServicePlaceCode = vbNullString Then
        For i = 0 To cboSerPlace.ListCount - 1
            If cboSerPlace.List(i) = clsEncounter.ServicePlace Then
                clsEncounter.ServicePlaceCode = cboSerPlace.ItemData(i)
                Exit For
            End If
        Next i
    End If
    If TextFound("Office", clsEncounter.ServicePlace, Len("Office"), True) Then
        clsEncounter.ServicePlaceCode = "11"
    End If
    If cboStatus.Text <> vbNullString Then
        clsEncounter.EncounterType = cboStatus.Text
    End If
    If clsEncounter.EncounterType = vbNullString Then
        clsEncounter.EncounterType = clsEncounter.EncounterType
    End If

    If cboProvider.Text <> vbNullString Then
        For i = 0 To cboProvider.ListCount - 1
            If cboProvider.List(i) = Trim(cboProvider.Text) Then
                clsEncounter.ProviderID = cboProvider.ItemData(i)
                Exit For
            End If
        Next i
    End If
    If Trim(cctRend.Text) <> vbNullString Then
        For i = 0 To lstRend.ListCount - 1
            If lstRend.Selected(i) Then
                clsEncounter.RendProv = lstRend.ItemData(i)
                Exit For
            End If
        Next i
    Else
        clsEncounter.RendProv = 0
    End If
    If cboVisitReason.Visible And Trim(cboVisitReason.Text) <> vbNullString Then
        If cboVisitReason.ListCount <= 0 Then GetTerms cboVisitReason, "sd_PostSurg_VisitType"
        For i = 0 To cboVisitReason.ListCount - 1
            If cboVisitReason.List(i) = cboVisitReason.Text Then
                clsEncounter.VisitReason = cboVisitReason.List(i)
                Exit For
            End If
        Next i
    ElseIf Trim(txtVisitReason.Text) <> vbNullString Then
        clsEncounter.VisitReason = Trim(txtVisitReason.Text)
    ElseIf Trim(txtVisitReason.Text) = vbNullString Then
        clsEncounter.VisitReason = vbNullString
    End If

    If clsEncounter.Appointment = vbNullString Then
        Get_ApptData clsEncounter.StartDate
    End If
    If cboFacility.Text <> vbNullString Then
        For i = 0 To cboFacility.ListCount - 1
            If cboFacility.List(i) = cboFacility.Text Then
                lngFacilityID = cboFacility.ItemData(i)
                clsEncounter.FacilityID = lngFacilityID
                Exit For
            End If
        Next i
    End If
    If cboStatus.Text <> vbNullString Then clsEncounter.EncounterType = cboStatus.Text
    If clsEncounter.EncounterType = "Procedure" Then
        If clsEncounter.Intrvn_CD = vbNullString Then
            If clsEncounter.VisitReason <> vbNullString Then
                clsOpNote.Preop_Code = vbNullString
                Get_Surgery_CPT clsEncounter.VisitReason
                If clsOpNote.Preop_Code <> vbNullString Then
                    clsEncounter.Intrvn_CD = clsOpNote.Preop_Code
                End If
            Else
                 clsEncounter.Intrvn_CD = "049X"
            End If
        End If
        udtFacility.FacID = 0
        Get_SurgFacility
        If udtFacility.FacID > 0 Then
            clsEncounter.FacilityID = udtFacility.FacID
        End If
    Else
        If clsEncounter.VisitType <> vbNullString Then
            LookUp_CPT Trim(clsEncounter.VisitType)
        Else
            If cboPurVisit.Text <> vbNullString Then
                clsEncounter.VisitType = cboPurVisit.Text
                LookUp_CPT Trim(clsEncounter.VisitType)
            End If
        End If
    End If
    If TextFound("Office", clsEncounter.ServicePlace, Len("Office"), True) Then
        clsEncounter.ServicePlaceCode = "11"
    End If
    If clsEncounter.VisitType <> vbNullString Then
        clsEncounter.Appt_Type = Get_ApptCode(clsEncounter.VisitType)
    End If
    If Val(clsEncounter.RendProv) <= 0 Then clsEncounter.RendProv = clsEncounter.ProviderID
    If cnEnctr.State = adStateClosed Then cnEnctr.open psConnect(0)
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If vstrAction = "I" Then
        strSQL = "INSERT INTO ENCOUNTER (PATIENT_ID, PROVIDER_ID, FACILITY_ID, ENCNTR_START_DATE," _
            & "ENCOUNT_CPT_INTRVN, VISIT_REASON, VISIT_TYPE, ENCNTR_TYPE, SERVICE_PLACE, SERVICE_PLACE_CD," _
            & "USER_ID, APPOINTMENT_DATE, ENCNTR_END_DATE,RENDERING_PROVIDER,ENCOUNT_APPT_TYPE, INSERT_DATE) " _
            & "values (" & piPatientID & "," & clsEncounter.ProviderID & "," & clsEncounter.FacilityID & ",'" _
            & clsEncounter.StartDate & "','" & clsEncounter.Intrvn_CD & "','" & clsEncounter.VisitReason & "','" _
            & clsEncounter.VisitType & "','" & clsEncounter.EncounterType & "','" & clsEncounter.ServicePlace & "','" _
            & clsEncounter.ServicePlaceCode & "','" & PsUserID & "','" & clsEncounter.Appointment & "','" & clsEncounter.EndDate & "'," _
            & clsEncounter.RendProv & ",'" & clsEncounter.Appt_Type & "','" & sDate & "')"
            Set rsEnctr = cnEnctr.Execute(strSQL)
    Else
        sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
        strSQL = "UPDATE ENCOUNTER SET PROVIDER_ID = " & clsEncounter.ProviderID & ", FACILITY_ID = " & clsEncounter.FacilityID & "," _
            & "ENCNTR_START_DATE = '" & clsEncounter.StartDate & "', ENCNTR_END_DATE = '" & clsEncounter.EndDate & "'," _
            & "ENCOUNT_CPT_INTRVN = '" & clsEncounter.Intrvn_CD & "',VISIT_REASON = '" & clsEncounter.VisitReason & "'," _
            & "VISIT_TYPE = '" & clsEncounter.VisitType & "',ENCNTR_TYPE = '" & clsEncounter.EncounterType & "'," _
            & "SERVICE_PLACE = '" & clsEncounter.ServicePlace & "',SERVICE_PLACE_CD = '" & clsEncounter.ServicePlaceCode & "'," _
            & "USER_ID = '" & PsUserID & "',APPOINTMENT_DATE = '" & clsEncounter.Appointment & "', RENDERING_PROVIDER = " & clsEncounter.RendProv & "," _
            & "ENCOUNT_APPT_TYPE = '" & clsEncounter.Appt_Type & "', UPDATE_DATE = '" & sDate & "' WHERE ENCOUNTER_ID = " & piEncounterID
            Set rsEnctr = cnEnctr.Execute(strSQL)
        strSQL = "UPDATE APPOINTMENTS SET PROVIDER_ID = " & clsEncounter.ProviderID & " FROM APPOINTMENTS a, ENCOUNTER b WHERE " _
            & "a.PATIENT_ID = " & piPatientID & " AND a.PATIENT_ID = b.PATIENT_ID AND " _
            & "a.APPT_DATE = '" & Format(clsEncounter.StartDate, "mm/dd/yyyy") & "' AND a.PURPOSE_VISIT = b.ENCOUNT_APPT_TYPE AND b.ENCOUNTER_ID  = " & piEncounterID
            Set rsEnctr = cnEnctr.Execute(strSQL)
        strSQL = "UPDATE ENCOUNTER SET PROVIDER_ID = " & clsEncounter.ProviderID & ", RENDERING_PROVIDER = " & piProviderID & "  WHERE ENCOUNTER_ID = " & piEncounterID
            Set rsEnctr = cnEnctr.Execute(strSQL)
        strSQL = "UPDATE EVALUATION_MANAGEMENT SET RENDERING_PROVIDER = " & clsEncounter.ProviderID & " WHERE ENCOUNTER_ID = " & piEncounterID
            Set rsEnctr = cnEnctr.Execute(strSQL)
    End If
    WriteData = True
    clsEncounter.VisitReason = vbNullString
    If vstrAction = "I" Then
        strSQL = "SELECT  ENCOUNTER_ID, ENCNTR_START_DATE FROM ENCOUNTER WHERE PATIENT_ID = " & piPatientID & " AND " _
            & "PROVIDER_ID = " & clsEncounter.ProviderID & " ORDER BY  INSERT_DATE DESC"
        Set rsEnctr = cnEnctr.Execute(strSQL)
        Do Until rsEnctr.EOF
                If Not IsNull(rsEnctr.Fields("ENCNTR_START_DATE")) Then
                    lngDateDiff = DateDiff("d", Format(rsEnctr.Fields("ENCNTR_START_DATE"), "mm/dd/yyyy"), Format(clsEncounter.StartDate, "mm/dd/yyyy"))
                End If
                If lngDateDiff = 0 Then
                    If rsEnctr.Fields("ENCOUNTER_ID") > 0 Then
                        piEncounterID = rsEnctr.Fields("ENCOUNTER_ID")
                        Exit Do
                    End If
                End If
            rsEnctr.MoveNext
        Loop
        If piEncounterID > 0 Then Get_Encounter piEncounterID
    End If
    If cboPurVisit.Text <> vbNullString Then
        clsEncounter.VisitType = Trim(cboPurVisit.Text)
        strSQL = "Update encounter set visit_Type = '" & clsEncounter.VisitType & "' " _
            & "where patient_id = " & piPatientID & " AND " & "Encounter_Id = " & piEncounterID
        Set rsEnctr = cnEnctr.Execute(strSQL)
        If mlngGridRow > 0 Then
            Arr = Split(LookUpColName("APPT TYPE"), "|")
            If UBound(Arr) > 0 Then
                lngCol = Arr(0)
                If lngCol > 0 Then
                    grdEnctrSummary.Row = mlngGridRow
                    grdEnctrSummary.Col = lngCol
                    If grdEnctrSummary.CellType = CellTypeComboBox Then
                        If grdEnctrSummary.TypeComboBoxCurSel > -1 Then
                            grdEnctrSummary.TypeComboBoxIndex = grdEnctrSummary.TypeComboBoxCurSel
                            
                            If grdEnctrSummary.TypeComboBoxString <> vbNullString Then clsEncounter.Appt_Type = grdEnctrSummary.TypeComboBoxString
                        End If
                    End If
                End If
            End If
        End If
        If clsEncounter.Appt_Type <> vbNullString Then
            strSQL = "UPDATE ENCOUNTER SET ENCOUNT_APPT_TYPE = '" & clsEncounter.Appt_Type & "' " _
                & "where PATIENT_ID = " & piPatientID & " AND  ENCOUNTER_ID = " & piEncounterID
            Set rsEnctr = cnEnctr.Execute(strSQL)
        End If
    End If
    If Trim(cctRend.Text) <> vbNullString Then
        For Idx = 0 To lstRend.ListCount - 1
            If lstRend.Selected(Idx) Then
                lngRendProv = lstRend.ItemData(Idx)
                Exit For
            End If
        Next Idx
        If clsEncounter.RendProv = 0 Then clsEncounter.RendProv = clsEncounter.ProviderID
        strSQL = "UPDATE ENCOUNTER SET RENDERING_PROVIDER = " & lngRendProv & " WHERE " _
            & "PATIENT_ID = " & piPatientID & " AND  ENCOUNTER_ID = " & piEncounterID
            Set rsEnctr = cnEnctr.Execute(strSQL)
            
        strSQL = "SELECT * FROM EVALUATION_MANAGEMENT WHERE ENCOUNTER_ID = " & piEncounterID
            Set rsEnctr = cnEnctr.Execute(strSQL)
        With rsEnctr
            If Not .EOF Then
                lngEMID = (!EM_ID)
            End If
        End With
        If lngEMID > 0 Then
            strSQL = "UPDATE EVALUATION_MANAGEMENT SET RENDERING_PROVIDER = " & clsEncounter.RendProv & ", UPDATE_DATE = '" & sDate & "' WHERE " _
                        & "PATIENT_ID = " & piPatientID & " AND  ENCOUNTER_ID = " & piEncounterID
            Set rsEnctr = cnEnctr.Execute(strSQL)
        Else
            strSQL = "INSERT INTO EVALUATION_MANAGEMENT (PATIENT_ID, ENCOUNTER_ID, RENDERING_PROVIDER, INSERT_DATE) " _
                & "values (" & piPatientID & ", " & piEncounterID & ", " & clsEncounter.RendProv & ", '" & sDate & "')"
            Set rsEnctr = cnEnctr.Execute(strSQL)
        End If
    End If
'    Select Case clsEncounter.ServicePlaceCode
'        case "21",
'
'    End Select
'    If clsEncounter.ServicePlaceCode <> "11 " And clsEncounter.ServicePlaceCode <> "32 " And clsEncounter.ServicePlaceCode <> "33 " Then
'         Manage_Admission piEncounterID
'        If TextFound("Initial", clsEncounter.EncounterType, Len("Initial"), True) Then
'            txtVisitReason.Text = "Admission"
'        End If
'    End If
    For Idx = 0 To lstReferral.ListCount - 1
        If lstReferral.Selected(Idx) Then
            clsReferral.ReferralID = lstReferral.ItemData(Idx)
            clsReferral.Consult_Reason = Trim(cctRefReason.Text)
            Update_ReferPhys clsReferral.ReferralID
            Exit For
        End If
    Next Idx
'    If clsEncounter.ServicePlaceCode <> "11" And Not IsNull(ccdAdmit.Value) Then
'        Manage_Admission piEncounterID
'    End If
    If cnEnctr.State = adStateClosed Then
        cnEnctr.Close
        Set cnEnctr = Nothing
        Set rsEnctr = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_WriteData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "WriteData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
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
            .RowHeight(vlngRow) = 3.5
            .BackColor = 12632256
            .AddCellSpan 0, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub





Private Sub grdProc_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue
Dim iRow%
Dim varCPTCode
        
        
    If mstrProc <> vbNullString Then
        If Trim(cctProcString.Text) = vbNullString Then
            If Trim(cboRedo.Text) <> vbNullString Then
                cctProcString.Text = cboRedo.Text & " " & Trim(mstrProc)
            Else
                cctProcString.Text = Trim(mstrProc)
            End If
        ElseIf Not TextFound(mstrProc, cctProcString.Text, Len(mstrProc), True) Then
            cctProcString.Text = Trim(cctProcString.Text) & vbNewLine & Trim(mstrProc)
        End If
        If mstrColCPT <> vbNullString Then
            If Trim(cctProcString.Tag) = vbNullString Then
                cctProcString.Tag = mstrColCPT
            ElseIf Not TextFound(mstrColCPT, cctProcString.Tag, Len(mstrColCPT), True) Then
                cctProcString.Tag = Trim(cctProcString.Tag) & vbNewLine & mstrColCPT
            End If
        End If
    End If
    mlngProcID = 0
    mstrColCPT = vbNullString
    mstrColName = vbNullString
    mstrProc = vbNullString
    mlngProcRow = Row
    mlngProcCol = Col
    cboRedo.ListIndex = -1
    With grdProc
        For iRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = mlngProcRow
                .Row2 = mlngProcRow
                .BackColor = NormalColor
            .BlockMode = False
        Next iRow
        mlngProcID = .GetRowItemData(mlngProcRow)
        .GetText 1, mlngProcRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then mstrProc = Trim(varCellValue)
        If varCellValue <> vbNullString Then
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = mlngProcRow
                .Row2 = mlngProcRow
                .BackColor = SelectColor
                .SetCellBorder 1, mlngProcRow, .MaxCols, mlngProcRow, CellBorderIndexOutline, 0, CellBorderStyleSolid
            .BlockMode = False
            .GetText 2, mlngProcRow, varCPTCode
            mstrColCPT = Trim(varCPTCode)
            Format_ColumnGrid
        End If
    End With
    Exit Sub
End Sub

Private Sub grdProc_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
Dim iRow%
Dim varCellValue
Dim strCPT As String
Dim strProc As String
    mlngProcRow = 0
    mlngProcRow = Row
    mlngProcCol = Col
    If mlngProcRow = 0 Then Exit Sub
    With grdProc
        .ReDraw = False
        .GetText 1, mlngProcRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            strProc = Trim(varCellValue)
        End If
        .GetText 2, mlngProcRow, varCellValue
        If Trim(varCellValue) <> vbNullString Then
            strCPT = Trim(varCellValue)
        End If
        cctProcString.Tag = Replace(cctProcString.Tag, strCPT, "")
        cctProcString.Text = Replace(cctProcString.Text, strProc, "")
        For iRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = mlngProcRow
                .Row2 = mlngProcRow
                .BackColor = NormalColor
                .CellType = CellTypeEdit
            .BlockMode = False
        Next iRow
        .ReDraw = True
    End With
    Exit Sub
End Sub


Private Sub lstListBox_ItemCheck(Item As Integer)
Dim i%
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim strSQL As String
Dim strProc As String
Dim strCPT As String
Dim lngProvID As Long

    If mbooByPass Then mbooByPass = False: Exit Sub
    txtDesc.Text = vbNullString
    txtDesc.Visible = False
    strProc = lstListBox.List(Item)
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "SELECT CPT_DESC, CPT_CODE FROM SURGERY_CODE_DICTIONARY WHERE PROCEDURE_NAME = '" & strProc & "'"
    Set rsTerms = cnTerms.Execute(strSQL)
    If Not rsTerms.EOF Then
        If rsTerms.Fields("CPT_CODE") <> vbNullString Then strCPT = rsTerms.Fields("CPT_CODE")
        If rsTerms.Fields("CPT_DESC") <> vbNullString Then
            txtDesc.Text = rsTerms.Fields("CPT_DESC")
            txtDesc.Visible = True
        End If
    End If
    strSQL = "SELECT  CPT_CODE, PROCEDURE_ID FROM  SURG_PROCEDURE_CODES WHERE PROCEDURE_NAME = '" & strProc & "'"
    Set rsTerms = cnTerms.Execute(strSQL)
    If Not rsTerms.EOF Then
        If rsTerms.Fields("CPT_CODE") <> vbNullString Then strCPT = rsTerms.Fields("CPT_CODE")
        If rsTerms.Fields("PROCEDURE_ID") <> vbNullString Then lngProvID = rsTerms.Fields("PROCEDURE_ID")
    End If

    If strCPT <> vbNullString And Trim(txtDesc.Text) <> vbNullString Then
        txtDesc.Text = Trim(txtDesc.Text) & vbNewLine & "CPT CODE" & " " & strCPT
    End If
    If strProc <> vbNullString Then clsOpNote.Procedure_Name = strProc
    If lngProvID > 0 Then
        If clsOpNote.Procedure_ID = vbNullString Then
            clsOpNote.Procedure_ID = lngProvID
        Else
            If Not TextFound(lngProvID, clsOpNote.Procedure_ID, Len("lngProvID"), True) Then
                clsOpNote.Procedure_ID = clsOpNote.Procedure_ID & "|" & lngProvID
            End If
        End If
    End If
    
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Exit Sub
End Sub

Private Sub lstReferral_ItemCheck(Item As Integer)
Dim Idx%

    If mbooByPass Then mbooByPass = False: Exit Sub
    clsReferral.ReferralID = 0
    clsReferral.ReferralID = lstReferral.ItemData(Item)
    lstReferral.Visible = False
    If clsReferral.ReferralID > 0 Then Get_Referral clsReferral.ReferralID
    cctReferral.Text = clsReferral.Physician
    cctRefReason.Visible = True
    Exit Sub
End Sub


Private Sub lstRend_ItemCheck(Item As Integer)

    If mbooByPass Then mbooByPass = False: Exit Sub
    If lstRend.Selected(Item) Then
        cctRend.Text = lstRend.List(Item)
        lstRend.Visible = False
    End If
    Exit Sub
End Sub





