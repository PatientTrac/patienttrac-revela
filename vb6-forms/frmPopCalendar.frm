VERSION 5.00
Object = "{C08D2A54-E5CB-4AFE-8AAD-B3E08A43486C}#1.3#0"; "XButton.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "pt_xp.ocx"
Object = "{2A5B1681-9AFA-4521-B035-B9AD0C75F381}#1.0#0"; "DedicatedControls.ocx"
Begin VB.Form frmPopCalendar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PopForm"
   ClientHeight    =   2820
   ClientLeft      =   7050
   ClientTop       =   3825
   ClientWidth     =   3405
   Icon            =   "frmPopCalendar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picControlBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   0
      ScaleHeight     =   2760
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   30
      Width           =   3405
      Begin VB.PictureBox picJumpBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   30
         ScaleHeight     =   1440
         ScaleWidth      =   2505
         TabIndex        =   61
         Top             =   345
         Visible         =   0   'False
         Width           =   2535
         Begin DedicatedControls.ccNumber ccnYear 
            Height          =   255
            Left            =   315
            TabIndex        =   63
            Top             =   645
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "Year"
            BackColor       =   16761024
            ForeColor       =   8388608
            BackStyle       =   0
            MaxLength       =   4
            FontName        =   "Georgia"
            FontBold        =   -1  'True
         End
         Begin DedicatedControls.ccDrop cccMonth 
            Height          =   315
            Left            =   180
            TabIndex        =   62
            Top             =   195
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            Caption         =   "Month"
            Text            =   ""
            BackColor       =   16761024
            ForeColor       =   8388608
            FontName        =   "Georgia"
            FontBold        =   -1  'True
         End
         Begin XBtn.XButton cmdNextYear 
            Height          =   225
            Left            =   2130
            TabIndex        =   81
            Top             =   690
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   397
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   6
            Caption         =   "»"
            Picture         =   "frmPopCalendar.frx":08CA
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdPrevYear 
            Height          =   225
            Left            =   1860
            TabIndex        =   82
            Top             =   690
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   397
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   6
            Caption         =   "«"
            Picture         =   "frmPopCalendar.frx":08E6
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin PT_XP_Button.PT_XP cmdSelect 
            Height          =   315
            Left            =   690
            TabIndex        =   86
            Top             =   1080
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
            AppearanceThemes=   3
            TransparentColor=   14737632
            Caption         =   "Set"
            DepthEvent      =   1
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            ShowFocus       =   -1  'True
         End
      End
      Begin VB.PictureBox picTimeBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   30
         ScaleHeight     =   705
         ScaleWidth      =   2535
         TabIndex        =   52
         Top             =   2040
         Width           =   2535
         Begin XBtn.XButton cmdKey 
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   64
            Top             =   390
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   5
            Caption         =   "0"
            Picture         =   "frmPopCalendar.frx":0902
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdKey 
            Height          =   285
            Index           =   1
            Left            =   330
            TabIndex        =   65
            Top             =   390
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   5
            Caption         =   "1"
            Picture         =   "frmPopCalendar.frx":091E
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdKey 
            Height          =   285
            Index           =   2
            Left            =   570
            TabIndex        =   66
            Top             =   390
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   5
            Caption         =   "2"
            Picture         =   "frmPopCalendar.frx":093A
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdKey 
            Height          =   285
            Index           =   3
            Left            =   810
            TabIndex        =   67
            Top             =   390
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   5
            Caption         =   "3"
            Picture         =   "frmPopCalendar.frx":0956
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdKey 
            Height          =   285
            Index           =   4
            Left            =   1050
            TabIndex        =   68
            Top             =   390
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   5
            Caption         =   "4"
            Picture         =   "frmPopCalendar.frx":0972
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdKey 
            Height          =   285
            Index           =   5
            Left            =   1290
            TabIndex        =   69
            Top             =   390
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   5
            Caption         =   "5"
            Picture         =   "frmPopCalendar.frx":098E
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdKey 
            Height          =   285
            Index           =   6
            Left            =   1530
            TabIndex        =   70
            Top             =   390
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   5
            Caption         =   "6"
            Picture         =   "frmPopCalendar.frx":09AA
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdKey 
            Height          =   285
            Index           =   7
            Left            =   1770
            TabIndex        =   71
            Top             =   390
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   5
            Caption         =   "7"
            Picture         =   "frmPopCalendar.frx":09C6
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdKey 
            Height          =   285
            Index           =   8
            Left            =   2010
            TabIndex        =   72
            Top             =   390
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   5
            Caption         =   "8"
            Picture         =   "frmPopCalendar.frx":09E2
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdKey 
            Height          =   285
            Index           =   9
            Left            =   2250
            TabIndex        =   73
            Top             =   390
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   5
            Caption         =   "9"
            Picture         =   "frmPopCalendar.frx":09FE
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdNextDigit 
            Height          =   225
            Left            =   2220
            TabIndex        =   83
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   397
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   6
            Caption         =   "»"
            Picture         =   "frmPopCalendar.frx":0A1A
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdPrevDigit 
            Height          =   225
            Left            =   1950
            TabIndex        =   84
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   397
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   6
            Caption         =   "«"
            Picture         =   "frmPopCalendar.frx":0A36
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin VB.Label lblTimeDigit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "0"
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
            Height          =   330
            Index           =   0
            Left            =   405
            TabIndex        =   60
            Top             =   30
            Width           =   165
         End
         Begin VB.Label lblTimeDigit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "0"
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
            Height          =   330
            Index           =   1
            Left            =   600
            TabIndex        =   59
            Top             =   30
            Width           =   165
         End
         Begin VB.Label lblTimeDigit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   ":"
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
            Height          =   330
            Index           =   2
            Left            =   750
            TabIndex        =   58
            Top             =   30
            Width           =   105
         End
         Begin VB.Label lblTimeDigit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "0"
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
            Height          =   330
            Index           =   3
            Left            =   855
            TabIndex        =   57
            Top             =   30
            Width           =   195
         End
         Begin VB.Label lblTimeDigit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "0"
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
            Height          =   330
            Index           =   4
            Left            =   1035
            TabIndex        =   56
            Top             =   30
            Width           =   195
         End
         Begin VB.Label lblTimeDigit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   ":"
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
            Height          =   330
            Index           =   5
            Left            =   1215
            TabIndex        =   55
            Top             =   30
            Width           =   135
         End
         Begin VB.Label lblTimeDigit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "0"
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
            Height          =   330
            Index           =   6
            Left            =   1335
            TabIndex        =   54
            Top             =   30
            Width           =   165
         End
         Begin VB.Label lblTimeDigit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "0"
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
            Height          =   330
            Index           =   7
            Left            =   1485
            TabIndex        =   53
            Top             =   30
            Width           =   225
         End
      End
      Begin VB.PictureBox picDateBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2025
         Left            =   0
         ScaleHeight     =   2025
         ScaleWidth      =   2625
         TabIndex        =   1
         Top             =   30
         Width           =   2625
         Begin XBtn.XButton cmdPrevious 
            Height          =   285
            Left            =   30
            TabIndex        =   79
            Top             =   0
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   6
            Caption         =   "«"
            Picture         =   "frmPopCalendar.frx":0A52
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin XBtn.XButton cmdNext 
            Height          =   285
            Left            =   2220
            TabIndex        =   80
            Top             =   0
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   8388608
            Bevel           =   6
            Caption         =   "»"
            Picture         =   "frmPopCalendar.frx":0A6E
            PictAlign       =   4
            CaptionAlign    =   4
            PushTimer       =   0   'False
            PushInterval    =   300
            MousePointer    =   0
            PicRaster       =   12583114
         End
         Begin VB.Shape shpSelectedDateMark 
            BorderColor     =   &H8000000E&
            BorderWidth     =   2
            Height          =   300
            Left            =   2130
            Top             =   750
            Width           =   420
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "September 2002"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   390
            TabIndex        =   51
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lblDow 
            Alignment       =   2  'Center
            BackColor       =   &H80000015&
            Caption         =   "Sat"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   195
            Index           =   6
            Left            =   2190
            TabIndex        =   50
            Top             =   315
            Width           =   375
         End
         Begin VB.Label lblDow 
            Alignment       =   2  'Center
            BackColor       =   &H80000015&
            Caption         =   "Fri"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   195
            Index           =   5
            Left            =   1830
            TabIndex        =   49
            Top             =   315
            Width           =   375
         End
         Begin VB.Label lblDow 
            Alignment       =   2  'Center
            BackColor       =   &H80000015&
            Caption         =   "Thu"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   195
            Index           =   4
            Left            =   1470
            TabIndex        =   48
            Top             =   315
            Width           =   375
         End
         Begin VB.Label lblDow 
            Alignment       =   2  'Center
            BackColor       =   &H80000015&
            Caption         =   "Wed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   195
            Index           =   3
            Left            =   1110
            TabIndex        =   47
            Top             =   315
            Width           =   375
         End
         Begin VB.Label lblDow 
            Alignment       =   2  'Center
            BackColor       =   &H80000015&
            Caption         =   "Tue"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   195
            Index           =   2
            Left            =   750
            TabIndex        =   46
            Top             =   315
            Width           =   375
         End
         Begin VB.Label lblDow 
            Alignment       =   2  'Center
            BackColor       =   &H80000015&
            Caption         =   "Mon"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   195
            Index           =   1
            Left            =   390
            TabIndex        =   45
            Top             =   315
            Width           =   375
         End
         Begin VB.Label lblDow 
            Alignment       =   2  'Center
            BackColor       =   &H80000015&
            Caption         =   "Sun"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000014&
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   44
            Top             =   315
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   41
            Left            =   2190
            TabIndex        =   43
            Top             =   1710
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   40
            Left            =   1830
            TabIndex        =   42
            Top             =   1710
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   39
            Left            =   1470
            TabIndex        =   41
            Top             =   1710
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   38
            Left            =   1110
            TabIndex        =   40
            Top             =   1710
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   37
            Left            =   750
            TabIndex        =   39
            Top             =   1710
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   36
            Left            =   390
            TabIndex        =   38
            Top             =   1710
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   35
            Left            =   30
            TabIndex        =   37
            Top             =   1710
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   34
            Left            =   2190
            TabIndex        =   36
            Top             =   1470
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   33
            Left            =   1830
            TabIndex        =   35
            Top             =   1470
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   32
            Left            =   1470
            TabIndex        =   34
            Top             =   1470
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   31
            Left            =   1110
            TabIndex        =   33
            Top             =   1470
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   30
            Left            =   750
            TabIndex        =   32
            Top             =   1470
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   29
            Left            =   390
            TabIndex        =   31
            Top             =   1470
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   28
            Left            =   30
            TabIndex        =   30
            Top             =   1470
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   27
            Left            =   2190
            TabIndex        =   29
            Top             =   1230
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   26
            Left            =   1830
            TabIndex        =   28
            Top             =   1230
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   25
            Left            =   1470
            TabIndex        =   27
            Top             =   1230
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   24
            Left            =   1110
            TabIndex        =   26
            Top             =   1230
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   23
            Left            =   750
            TabIndex        =   25
            Top             =   1230
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   22
            Left            =   390
            TabIndex        =   24
            Top             =   1230
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   21
            Left            =   30
            TabIndex        =   23
            Top             =   1230
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   2190
            TabIndex        =   22
            Top             =   990
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   19
            Left            =   1830
            TabIndex        =   21
            Top             =   990
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   18
            Left            =   1470
            TabIndex        =   20
            Top             =   990
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   1110
            TabIndex        =   19
            Top             =   990
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   750
            TabIndex        =   18
            Top             =   990
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   390
            TabIndex        =   17
            Top             =   990
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   30
            TabIndex        =   16
            Top             =   990
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   2190
            TabIndex        =   15
            Top             =   750
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   1830
            TabIndex        =   14
            Top             =   750
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   1470
            TabIndex        =   13
            Top             =   750
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   1110
            TabIndex        =   12
            Top             =   750
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   750
            TabIndex        =   11
            Top             =   750
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   390
            TabIndex        =   10
            Top             =   750
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   30
            TabIndex        =   9
            Top             =   750
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   2190
            TabIndex        =   8
            Top             =   510
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   1830
            TabIndex        =   7
            Top             =   510
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1470
            TabIndex        =   6
            Top             =   510
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   1110
            TabIndex        =   5
            Top             =   510
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   750
            TabIndex        =   4
            Top             =   510
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   390
            TabIndex        =   3
            Top             =   510
            Width           =   375
         End
         Begin VB.Label lblDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   30
            TabIndex        =   2
            Top             =   510
            Width           =   375
         End
      End
      Begin XBtn.XButton cmdNow 
         Height          =   285
         Left            =   2610
         TabIndex        =   74
         Top             =   2250
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
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
         Caption         =   "Now"
         Picture         =   "frmPopCalendar.frx":0A8A
         PictAlign       =   4
         CaptionAlign    =   4
         PushTimer       =   0   'False
         PushInterval    =   300
         MousePointer    =   0
         PicRaster       =   12583114
      End
      Begin XBtn.XButton cmdOK 
         Height          =   285
         Left            =   2640
         TabIndex        =   75
         Top             =   30
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
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
         Caption         =   "OK"
         Picture         =   "frmPopCalendar.frx":0AA6
         PictAlign       =   4
         CaptionAlign    =   4
         PushTimer       =   0   'False
         PushInterval    =   300
         MousePointer    =   0
         PicRaster       =   12583114
      End
      Begin XBtn.XButton cmdCancel 
         Height          =   285
         Left            =   2640
         TabIndex        =   76
         Top             =   390
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
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
         Caption         =   "Cancel"
         Picture         =   "frmPopCalendar.frx":0AC2
         PictAlign       =   4
         CaptionAlign    =   4
         PushTimer       =   0   'False
         PushInterval    =   300
         MousePointer    =   0
         PicRaster       =   12583114
      End
      Begin XBtn.XButton cmdJump 
         Height          =   285
         Left            =   2640
         TabIndex        =   77
         Top             =   750
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8388608
         Bevel           =   6
         Caption         =   "Go to ..."
         Picture         =   "frmPopCalendar.frx":0ADE
         PictAlign       =   4
         CaptionAlign    =   4
         PushTimer       =   0   'False
         PushInterval    =   300
         MousePointer    =   0
         PicRaster       =   12583114
      End
      Begin XBtn.XButton cmdToday 
         Height          =   285
         Left            =   2640
         TabIndex        =   78
         Top             =   1110
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
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
         Caption         =   "Today"
         Picture         =   "frmPopCalendar.frx":0AFA
         PictAlign       =   4
         CaptionAlign    =   4
         PushTimer       =   0   'False
         PushInterval    =   300
         MousePointer    =   0
         PicRaster       =   12583114
      End
      Begin VB.CheckBox chkJump 
         Caption         =   "Check1"
         Height          =   225
         Left            =   2760
         TabIndex        =   85
         Top             =   780
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmPopCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdtsReference As Date
Private mstrPreviousValue As String
Private mctlTarget As Control
Private mctlTarget2 As Object
Private mstrTime As String
Private mbytTimeDigitIndex As Byte
Private menuMode As enmBoxType


Public Sub PopulateCalendar()

    Dim lintOffset As Integer
    Dim lintCounter As Integer
    Dim lintLastVisibleCell As Integer
    Dim ldtsDate As Date
    
    'Set the title label for the month
    lblTitle.Caption = Format(mdtsReference, "mmmm yyyy")
    
    'Make sure selected date marker always visible
    shpSelectedDateMark.Visible = True
    
    'Find the offset for the first day of the month
    lintOffset = Format(DateSerial(Year(mdtsReference), Month(mdtsReference), 1), "w")
    
    'Find the date corresponding to the first label
    ldtsDate = DateAdd("d", DateSerial(Year(mdtsReference), Month(mdtsReference), 1), 1 - lintOffset)

    'Cycle through all day labels in the month's matrix
    For lintCounter = 0 To 41
        
        'Test if the month corresponding to the label in question does not match the reference date
        If Month(ldtsDate) <> Month(mdtsReference) Then
        
            'Characteristics for labels for dates outside the month in question
            lblDay(lintCounter).Caption = ""
            If lblDay(lintCounter).Top > lblDay(lintLastVisibleCell).Top Then
            
                lblDay(lintCounter).Visible = False
            Else
                lblDay(lintCounter).Visible = True
                
            End If
        Else
        
            'Otherwise, show the label and set the caption for the day
            lblDay(lintCounter).Visible = True
            lintLastVisibleCell = lintCounter
            lblDay(lintCounter).Caption = Day(ldtsDate)
            
            'Test if this is the selected date and mark accordingly
            If DateValue(ldtsDate) = DateValue(mdtsReference) Then
                HighlightCell lintCounter
            End If
        
        End If
        
        'Call routine to define cell colors
        SetCellColors lintCounter, ((DateValue(ldtsDate) = DateValue(Now)) And (Month(ldtsDate) = Month(mdtsReference)))
            
        'increase the date counter and continue with loop
        ldtsDate = DateAdd("d", ldtsDate, 1)
        
    Next
    
    'If in date only mode, trim the bottom of the form to avoid vbNullString space
    If menuMode = enuDate Then
        picControlBack.Height = lblDay(lintLastVisibleCell).Top + lblDay(lintLastVisibleCell).Height + 6 * Screen.TwipsPerPixelY
        Me.Height = 3 * Screen.TwipsPerPixelY + picControlBack.Height
    End If
    
End Sub

Private Sub cccMonth_AfterUpdate()
    
    ReferenceDate = (DateSerial(Year(mdtsReference), cccMonth.ListIndex + 1, Day(mdtsReference))) + _
     TimeValue(mstrTime)
     
End Sub

Private Sub ccnYear_AfterUpdate()
    
    ReferenceDate = DateSerial(Val(ccnYear.Text), Month(mdtsReference), Day(mdtsReference)) + _
     TimeValue(mstrTime)
     
End Sub

Private Sub chkJump_Click()

    If chkJump.Value = 1 Then
        cccMonth.Text = Format(mdtsReference, "mmmm")
        ccnYear.Text = Format(mdtsReference, "yyyy")
        picDateBack.Visible = False
        picJumpBack.Visible = True
        cmdOK.Enabled = False
        cmdCancel.Enabled = False
    Else
        PopulateCalendar
        picDateBack.Visible = True
        picJumpBack.Visible = False
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
    End If
    
End Sub

Private Sub cmdJump_Click()
    chkJump.Value = 1
    chkJump_Click
End Sub


Private Sub cmdNext_Click()

    ReferenceDate = DateAdd("m", 1, mdtsReference)
    PopulateCalendar

End Sub

Private Sub cmdNextYear_Click()

    ccnYear.Text = Val(ccnYear.Text) + 1
    ccnYear_AfterUpdate
    
End Sub

Private Sub cmdNow_Click()

    mstrTime = Format$(Now, "HH:NN:SS")
    PopulateTimeEntry
    TimeDigitIndex = 0
    
    'Find the date corresponding to the selected label
    ReferenceDate = DateValue(mdtsReference) + TimeValue((mstrTime))
    
End Sub

Private Sub cmdOK_Click()
    mbooPopOK = True
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    mbooPopCancel = True
    mctlTarget.Text = mstrPreviousValue
    mctlTarget2.Trigger_AfterUpdate
    mbooPopCancel = False
    Unload Me
    
End Sub

Public Property Set Target(ByRef rctlNewReference As Control)

    Set mctlTarget = rctlNewReference
    
    InitializeView
    
End Property

Public Property Set Target2(ByRef rctlNewReference As Object)
    Set mctlTarget2 = rctlNewReference
End Property
Private Sub cmdPrevious_Click()

    ReferenceDate = DateAdd("m", -1, mdtsReference)
    PopulateCalendar

End Sub

Private Sub cmdPrevYear_Click()

    ccnYear.Text = Val(ccnYear.Text) - 1
    ccnYear_AfterUpdate
    
End Sub

Private Sub cmdSelect_Click()
    chkJump.Value = 0
    chkJump_Click
End Sub

Private Sub cmdToday_Click()

    ReferenceDate = DateValue(Now) + TimeValue((mstrTime))
    cccMonth.Text = Format(mdtsReference, "mmmm")
    ccnYear.Text = Format(mdtsReference, "yyyy")
    PopulateCalendar
    
End Sub

Private Sub Form_Load()

    Dim lintCounter As Integer
    
    For lintCounter = 1 To 12
        cccMonth.AddItem Format$(DateSerial(1980, lintCounter, 1), "mmmm")
    Next
    mbooPopOK = False
    mbooPopCancel = False
End Sub

Private Sub lblDay_Click(Index As Integer)

    If lblDay(Index).Caption <> "" Then
        HighlightCell Index
        
        'Find the date corresponding to the selected label
        ReferenceDate = DateSerial(Year(mdtsReference), Month(mdtsReference), lblDay(Index).Caption) + TimeValue((mstrTime))
    End If

End Sub

Private Sub HighlightCell(ByVal vintIndex As Integer)
    shpSelectedDateMark.Visible = True
    shpSelectedDateMark.Move lblDay(vintIndex).Left - Screen.TwipsPerPixelX, lblDay(vintIndex).Top - Screen.TwipsPerPixelY
End Sub

Private Sub SetCellColors(ByVal vintIndex As Integer, ByVal vbooToday As Boolean)
    
    'Highlight the cell if it corresponds to the current date
    If vbooToday Then
        lblDay(vintIndex).BackColor = &H80000014
        lblDay(vintIndex).ForeColor = &H80000010
    Else
        'If cell is not for the current date, define color based on day of the week
        If (vintIndex Mod 7 = 0) Or (vintIndex Mod 7 = 6) Then
            'Saturday and Sundays
            lblDay(vintIndex).BackColor = &H80000010
            lblDay(vintIndex).ForeColor = &H8000000E
        Else
            'Weekdays
            lblDay(vintIndex).BackColor = &H8000000F
            lblDay(vintIndex).ForeColor = &H80000012
        End If
    End If
    
End Sub

Private Property Let ReferenceDate(ByVal vdtsNewValue As Date)

    mdtsReference = vdtsNewValue
    
    mctlTarget.Text = Format$(mdtsReference, mctlTarget.Format)
    
End Property

Private Sub PopulateTime()

    Dim lbytDigitCounter As Byte
    
    For lbytDigitCounter = 1 To Len(mstrTime)
        lblTimeDigit(lbytDigitCounter - 1).Caption = Mid$(mstrTime, lbytDigitCounter, 1)
    Next
    
End Sub


Private Sub EnableValidKeys()

    Dim lbytKeyCounter As Byte

    Select Case mbytTimeDigitIndex
        Case 0
            For lbytKeyCounter = 0 To 9
                cmdKey(lbytKeyCounter).Enabled = (lbytKeyCounter < 3)
            Next
        Case 1
            For lbytKeyCounter = 0 To 9
                cmdKey(lbytKeyCounter).Enabled = _
                 Choose(lblTimeDigit(0).Caption + 1, True, True, (lbytKeyCounter < 4))
            Next
        Case 3, 6
            For lbytKeyCounter = 0 To 9
                cmdKey(lbytKeyCounter).Enabled = (lbytKeyCounter < 6)
            Next
        Case 4, 7
            For lbytKeyCounter = 0 To 9
                cmdKey(lbytKeyCounter).Enabled = True
            Next
    End Select
    
End Sub

Private Sub cmdKey_Click(Index As Integer)

    lblTimeDigit(mbytTimeDigitIndex).Caption = cmdKey(Index).Caption
    Mid(mstrTime, mbytTimeDigitIndex + 1, 1) = cmdKey(Index).Caption
        
    cmdNextDigit_Click
    'Find the date corresponding to the selected label
    ReferenceDate = DateValue(mdtsReference) + TimeValue(CVDate(mstrTime))
    
    
End Sub


Private Sub lblTimeDigit_Click(Index As Integer)

    If Index <> 2 And Index <> 5 Then
    
        TimeDigitIndex = Index
        
    End If

End Sub

Private Property Get TimeDigitIndex() As Byte

    TimeDigitIndex = mbytTimeDigitIndex
    
End Property

Private Property Let TimeDigitIndex(ByVal vbytNewValue As Byte)

    lblTimeDigit(mbytTimeDigitIndex).BackColor = &H80000010
    lblTimeDigit(mbytTimeDigitIndex).ForeColor = &H8000000E
    mbytTimeDigitIndex = vbytNewValue
    lblTimeDigit(mbytTimeDigitIndex).BackColor = &H8000000F
    lblTimeDigit(mbytTimeDigitIndex).ForeColor = &H8000000D
    
    EnableValidKeys
    
    If cmdKey(lblTimeDigit(mbytTimeDigitIndex).Caption).Enabled = False Then
        lblTimeDigit(mbytTimeDigitIndex).Caption = "0"
        Mid(mstrTime, mbytTimeDigitIndex + 1, 1) = "0"
    End If
    
End Property


Private Sub cmdPrevDigit_Click()

    Select Case mbytTimeDigitIndex
        Case 7, 4, 1
            TimeDigitIndex = mbytTimeDigitIndex - 1
        Case 3, 6
            TimeDigitIndex = mbytTimeDigitIndex - 2
    End Select

End Sub

Private Sub cmdNextDigit_Click()

    Select Case mbytTimeDigitIndex
        Case 0, 3, 6
            TimeDigitIndex = mbytTimeDigitIndex + 1
        Case 1, 4
            TimeDigitIndex = mbytTimeDigitIndex + 2
    End Select
    
End Sub


Private Sub InitializeView()

    mstrPreviousValue = mctlTarget.Text
    
    If Not IsDate(mstrPreviousValue) Then
        ReferenceDate = Now
    Else
        ReferenceDate = CVDate(mstrPreviousValue)
    End If
    
    mstrTime = Format$(mdtsReference, "HH:NN:SS")
    
    If InStr(mstrPreviousValue, ":") <> 0 Then
        If InStr(mstrPreviousValue, "/") <> 0 Then
            menuMode = enuDateTime
        Else
            menuMode = enuTime
        End If
    Else
        menuMode = enuDate
    End If
    
    If menuMode = enuDate Or menuMode = enuDateTime Then
        PopulateCalendar
        
        If menuMode = enuDate Then
            picTimeBack.Visible = False
            cmdNow.Visible = False
        End If
    End If
    
    If menuMode = enuTime Or menuMode = enuDateTime Then
        
        TimeDigitIndex = 0
        PopulateTimeEntry
        
        If menuMode = enuTime Then
            picDateBack.Visible = False
            cmdToday.Visible = False
            chkJump.Visible = False
            picTimeBack.Top = picDateBack.Top
            cmdNow.Top = cmdToday.Top
        
            picControlBack.Height = cmdNow.Top + cmdNow.Height + 6 * Screen.TwipsPerPixelY
            Me.Height = 3 * Screen.TwipsPerPixelY + picControlBack.Height
        
        End If
    
    End If
    
    
End Sub

Private Sub PopulateTimeEntry()

    Dim lintCounter As Integer
    
    For lintCounter = 0 To 7
        lblTimeDigit(lintCounter).Caption = Mid$(mstrTime, lintCounter + 1, 1)
    Next

End Sub
