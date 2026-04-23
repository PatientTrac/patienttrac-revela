VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmPlasticSurgery 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9930
   ClientLeft      =   1545
   ClientTop       =   1785
   ClientWidth     =   16335
   ForeColor       =   &H00800000&
   Icon            =   "frmPlasticSurgery.frx":0000
   ScaleHeight     =   9930
   ScaleWidth      =   16335
   Begin PT_XP_Button.PT_XP cmdLogs 
      Height          =   585
      Left            =   13830
      TabIndex        =   72
      Top             =   3720
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1032
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
      Picture         =   "frmPlasticSurgery.frx":10E3
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Logs"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPlasticSurgery.frx":185D
      ShowFocus       =   -1  'True
   End
   Begin VB.Frame fraLogs 
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   5205
      Left            =   14280
      TabIndex        =   64
      Top             =   4410
      Visible         =   0   'False
      Width           =   1965
      Begin PT_XP_Button.PT_XP cmdAddMed 
         Height          =   585
         Left            =   90
         TabIndex        =   65
         Top             =   4410
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1032
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
         Caption         =   "Manage OR Medications "
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP PT_XP5 
         Height          =   615
         Left            =   120
         TabIndex        =   66
         Top             =   3000
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP PT_XP6 
         Height          =   615
         Left            =   120
         TabIndex        =   67
         ToolTipText     =   "PACU"
         Top             =   2280
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP fmdSurgLog 
         Height          =   615
         Left            =   120
         TabIndex        =   68
         Top             =   180
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Surgery Log"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdSpecimenLog 
         Height          =   615
         Left            =   120
         TabIndex        =   69
         Top             =   855
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Specimen Log"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdEquiptLog 
         Height          =   615
         Left            =   120
         TabIndex        =   70
         Top             =   1575
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Equipt Log"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP PT_XP15 
         Height          =   615
         Left            =   120
         TabIndex        =   71
         Top             =   3690
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Operative Logs"
         DepthEvent      =   1
      End
   End
   Begin VB.ComboBox cboSchedule 
      Height          =   315
      Left            =   240
      TabIndex        =   33
      Top             =   8970
      Visible         =   0   'False
      Width           =   3285
   End
   Begin PT_XP_Button.PT_XP cmdPatient 
      Height          =   1305
      Left            =   13860
      TabIndex        =   20
      ToolTipText     =   "Select Patient"
      Top             =   990
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   2302
      BackColor       =   14536141
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
      Picture         =   "frmPlasticSurgery.frx":1F6F
      AppearanceThemes=   3
      BackColorDown   =   14536141
      BackColorOver   =   13432063
      BackColorFocus  =   14536141
      BackColorDisabled=   14536141
      TransparentColor=   14737632
      Caption         =   ""
      CaptionEffect   =   1
      CaptionPosition =   4
      Depth           =   5
      DepthEvent      =   1
      PictureDisabled =   "frmPlasticSurgery.frx":3C51
      PictureOver     =   "frmPlasticSurgery.frx":16447
      PictureAlignment=   4
      PictureOverEffect=   3
   End
   Begin VB.PictureBox picOR 
      Height          =   2385
      Left            =   10590
      Picture         =   "frmPlasticSurgery.frx":28C3D
      ScaleHeight     =   2325
      ScaleWidth      =   3105
      TabIndex        =   11
      Top             =   1950
      Width           =   3165
   End
   Begin VB.PictureBox picAdmin 
      Height          =   2835
      Left            =   150
      Picture         =   "frmPlasticSurgery.frx":2B60C
      ScaleHeight     =   2775
      ScaleWidth      =   3075
      TabIndex        =   10
      Top             =   2040
      Width           =   3135
   End
   Begin VB.PictureBox picOrders 
      Height          =   2265
      Left            =   6900
      Picture         =   "frmPlasticSurgery.frx":30732
      ScaleHeight     =   2205
      ScaleWidth      =   3285
      TabIndex        =   9
      Top             =   1980
      Width           =   3345
   End
   Begin VB.Frame fraConsult 
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   3735
      Left            =   3330
      TabIndex        =   8
      Top             =   4410
      Visible         =   0   'False
      Width           =   3375
      Begin PT_XP_Button.PT_XP cmdPhyExam 
         Height          =   615
         Left            =   60
         TabIndex        =   47
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Physical Examination"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdBreastExam 
         Height          =   615
         Left            =   60
         TabIndex        =   48
         Top             =   840
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Breast Exam"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdGlutealPlasty 
         Height          =   615
         Left            =   1680
         TabIndex        =   49
         Top             =   870
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Glutealplasty"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdBodyExam 
         Height          =   615
         Left            =   60
         TabIndex        =   50
         Top             =   1530
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Body Exam"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdUpperEye 
         Height          =   615
         Left            =   1680
         TabIndex        =   51
         Top             =   1530
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Upper Eye"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdFace 
         Height          =   615
         Left            =   60
         TabIndex        =   52
         Top             =   2220
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Face"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdLowEye 
         Height          =   615
         Left            =   1680
         TabIndex        =   53
         Top             =   2910
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Lower Eyes"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdForeHair 
         Height          =   615
         Left            =   60
         TabIndex        =   54
         Top             =   2880
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "ForeHead / Hair"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdNose 
         Height          =   615
         Left            =   1680
         TabIndex        =   55
         Top             =   2220
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Nose"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdProgress 
         Height          =   645
         Left            =   1680
         TabIndex        =   56
         Top             =   150
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1138
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Progress Note"
         DepthEvent      =   1
      End
   End
   Begin VB.Frame fraOR 
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   5115
      Left            =   10500
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   3735
      Begin PT_XP_Button.PT_XP cmdMedInv 
         Height          =   585
         Left            =   60
         TabIndex        =   63
         Top             =   4410
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1032
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
         Caption         =   "Medication Inventory - Dispense"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdOpMeds 
         Height          =   645
         Left            =   1890
         TabIndex        =   31
         Top             =   3000
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1138
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Post Operative Orders"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdOperativeRep 
         Height          =   615
         Left            =   60
         TabIndex        =   30
         Top             =   3690
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Operative Reports"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdPostOpExam 
         Height          =   645
         Left            =   1890
         TabIndex        =   29
         Top             =   2280
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1138
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Post Operative Exam"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdDis 
         Height          =   615
         Left            =   60
         TabIndex        =   28
         Top             =   3000
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Discharge Sheet"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdRecoveryRoom 
         Height          =   615
         Left            =   60
         TabIndex        =   27
         ToolTipText     =   "PACU"
         Top             =   2280
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Recovery Room "
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdAnesEvaluation 
         Height          =   645
         Left            =   1890
         TabIndex        =   22
         Top             =   150
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1138
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Anesthesia Evaluation - Questionairre"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdPatCheck 
         Height          =   615
         Left            =   60
         TabIndex        =   21
         Top             =   180
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Patient Check In"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdPreHP 
         Height          =   615
         Left            =   60
         TabIndex        =   23
         Top             =   855
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "History and Physical"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdPreCheck 
         Height          =   645
         Left            =   1890
         TabIndex        =   24
         Top             =   855
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1138
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Pre-Operative Checklist"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdAnest 
         Height          =   615
         Left            =   60
         TabIndex        =   25
         Top             =   1575
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Anesthesia"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdIntrOp 
         Height          =   645
         Left            =   1890
         TabIndex        =   26
         Top             =   1575
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1138
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Operating Room Record"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdOpLogs 
         Height          =   615
         Left            =   1890
         TabIndex        =   32
         Top             =   3690
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Operative Logs"
         DepthEvent      =   1
      End
   End
   Begin VB.Frame fraOrders 
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   4395
      Left            =   6750
      TabIndex        =   6
      Top             =   4410
      Visible         =   0   'False
      Width           =   3675
      Begin PT_XP_Button.PT_XP cmdOpPlan 
         Height          =   615
         Left            =   60
         TabIndex        =   35
         Top             =   150
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Operative Plan"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdFollPlan 
         Height          =   615
         Left            =   1890
         TabIndex        =   36
         Top             =   150
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Follow-Up Plan"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdLabOrders 
         Height          =   615
         Left            =   60
         TabIndex        =   37
         Top             =   825
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Surgery Orders"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdMeds 
         Height          =   645
         Left            =   1890
         TabIndex        =   38
         Top             =   825
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1138
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Medication Orders"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdUR 
         Height          =   615
         Left            =   60
         TabIndex        =   39
         Top             =   2250
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Utilization Review"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdResults 
         Height          =   645
         Left            =   1890
         TabIndex        =   40
         Top             =   1545
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1138
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Results"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP PT_XP9 
         Height          =   615
         Left            =   1890
         TabIndex        =   41
         ToolTipText     =   "PACU"
         Top             =   3660
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BackColor       =   14536141
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdSpecimens 
         Height          =   615
         Left            =   60
         TabIndex        =   42
         Top             =   2970
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Specimen Logs"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdOtherRep 
         Height          =   645
         Left            =   1890
         TabIndex        =   43
         Top             =   2250
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1138
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Other Reports"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP PT_XP12 
         Height          =   615
         Left            =   60
         TabIndex        =   44
         Top             =   3660
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdProgressNote 
         Height          =   645
         Left            =   1890
         TabIndex        =   45
         Top             =   2970
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1138
         BackColor       =   14536141
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Progress Notes"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdReferrals 
         Height          =   615
         Left            =   60
         TabIndex        =   46
         Top             =   1560
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Referrals"
         DepthEvent      =   1
      End
   End
   Begin VB.Frame frPatientInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   30
      TabIndex        =   3
      Top             =   60
      Width           =   15945
      Begin PT_XP_Button.PT_XP cmdSchedule 
         Height          =   795
         Left            =   12420
         TabIndex        =   19
         Top             =   0
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1402
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Schedule"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdMedication 
         Height          =   795
         Left            =   10320
         TabIndex        =   18
         Top             =   0
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1402
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Prescriptions"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdConsentTreat 
         Height          =   795
         Left            =   7560
         TabIndex        =   17
         Top             =   0
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   1402
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPlasticSurgery.frx":34FAE
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Consent and Treatment Forms"
         DepthEvent      =   1
         PictureDisabled =   "frmPlasticSurgery.frx":35888
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdHospPhone 
         Height          =   795
         Left            =   5220
         TabIndex        =   16
         Top             =   0
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1402
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Hospital Room Phone No"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAdministration 
         Height          =   795
         Left            =   3210
         TabIndex        =   15
         Top             =   0
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1402
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Administration"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPatientCommand 
         Height          =   795
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   1402
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmPlasticSurgery.frx":364DA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Patient Contact"
         DepthEvent      =   1
         PictureDisabled =   "frmPlasticSurgery.frx":373B4
         ShowFocus       =   -1  'True
      End
      Begin VB.Label LblPatient 
         AutoSize        =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4350
         TabIndex        =   5
         Top             =   180
         Width           =   765
      End
      Begin VB.Label LblEncounterDate 
         AutoSize        =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9060
         TabIndex        =   4
         Top             =   180
         Width           =   45
      End
   End
   Begin VB.Frame fraAdmin 
      BackColor       =   &H00800000&
      ForeColor       =   &H00800000&
      Height          =   2445
      Left            =   120
      TabIndex        =   2
      Top             =   4980
      Visible         =   0   'False
      Width           =   3165
      Begin PT_XP_Button.PT_XP cmdProposals 
         Height          =   615
         Left            =   30
         TabIndex        =   57
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Proposals"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdScheduling 
         Height          =   615
         Left            =   30
         TabIndex        =   58
         Top             =   840
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Scheduling"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdPracMgmt 
         Height          =   615
         Left            =   1590
         TabIndex        =   59
         Top             =   150
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Practice Management"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdPracRept 
         Height          =   615
         Left            =   1590
         TabIndex        =   60
         Top             =   840
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         BackColor       =   14536141
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
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         TransparentColor=   14737632
         Caption         =   "Reports"
         DepthEvent      =   1
      End
   End
   Begin VB.PictureBox picConsult 
      Height          =   2235
      Left            =   3600
      Picture         =   "frmPlasticSurgery.frx":38F06
      ScaleHeight     =   2175
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   2010
      Width           =   2895
   End
   Begin VB.Label lblPatName 
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
      Height          =   495
      Left            =   60
      TabIndex        =   62
      Top             =   900
      Width           =   3135
   End
   Begin VB.Label lblShowBy 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Patient Data"
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
      Left            =   300
      TabIndex        =   34
      Top             =   8640
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label lblOR 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Operating Room"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   795
      Left            =   10140
      TabIndex        =   14
      Top             =   1530
      Width           =   3555
   End
   Begin VB.Label lblPatientAdmin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Administration"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   60
      TabIndex        =   13
      Top             =   1530
      Width           =   3195
   End
   Begin VB.Label lblOrders 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Orders/Plan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   825
      Left            =   6930
      TabIndex        =   12
      Top             =   1530
      Width           =   2745
   End
   Begin VB.Label lblMedical 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Consultation/ Progress Notes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   795
      Left            =   3210
      TabIndex        =   1
      Top             =   1200
      Width           =   3585
   End
End
Attribute VB_Name = "frmPlasticSurgery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public piPatientID As Long
Public piEncounterID As Long
Public piSelEncntrID As Long
Public piProviderID As Long
Public piPatientExtID As Long
Public piFacilityID As Long
Public mintProviderID As Integer
Private msEncounterDate As String
Public mstrProviderName As String
Private lintComponentCounter As Integer
Private lintTabCounter As Integer
Private lbooMissingClass As Boolean
Private lstrClassName As String
Private lstrLibrary As String
Private mlngCurrentPatientID As Long
Private mlngCurrentEncounterID As Long
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private mlngFormID As Long
'dnn.................
Private mintTabs As Long
Private mintTabsPerRow As Long
Private fstrTermsConnect As String
Private fstrClinicalConnect As String
Private fstrReportDate As Date
Private fstrReportDate2 As Date
Private mbooFormLoaded As Boolean
Private mstrDefProvider As String
Private mstrDefFacility As String
Private mstrIniFile As String
Private mbooGet As Boolean
Private mstrReportFileName As String
'*** Note!   This collection below is a rudimentary first attempt to handle this.
'            We need to create a specific collection that allows to manage the
'            loaded controls on a page by page basis, instead of keeping all
'            controls for all pages in the same bag.
Private mbooformload As Boolean
'Dimension collection to keep track of loaded clinical components
Private mcolActxComp As New Collection
Private mbooSelectPat As Boolean
Private mbooRecurssive As Boolean
Private mlngLastPatientId As Long
Private mlngLastEncounterId As Long
Private mObjSelectPatient As PatientTracTypeLib.iClinical
'Private WithEvents PatientList As SelectPatient
Public mintindex As Integer
Private mFirst As Boolean
Private mstrRemarks As String
Private mbooNext As Boolean
Private mbooPrev As Boolean









Private Sub cmdAddMed_Click()
    frmMedMaint.Show vbModal
    Exit Sub
End Sub


Private Sub cmdAdministration_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Load_ChildForm "FN_Orders", "Administration"
    frmPlastiSurg.Show
    Exit Sub
End Sub

Private Sub cmdAnesEvaluation_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Load_ChildForm "FN_Orders", "AnesEvaluation"
    frmPlastiSurg.Show
End Sub

Private Sub cmdAnest_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Load_ChildForm "FN_Orders", "Anesthesia"
    frmPlastiSurg.Show
End Sub

Private Sub cmdBodyExam_Click()
    frmBreastExam.Show vbModal
End Sub

Private Sub cmdBreastExam_Click()
    frmBreastExam.piPatientID = piPatientID
    frmBreastExam.piEncounterID = piEncounterID
    frmBreastExam.piProviderID = piProviderID
    frmBreastExam.Show vbModal
End Sub

Private Sub cmdDiag_Click()
    'If piEncounterID < 1 Then MsgBox "Select an encounter first.", vbExclamation: Exit Sub
   ' 'Load_Control "FN_InstClinical", "Diagnosis", True
End Sub


Private Sub cmdConsentTreat_Click()
    frmConsent.piPatientID = modSpecific.piPatientID
    frmConsent.Show vbModal
    Exit Sub
End Sub

Private Sub cmdIntrOp_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Load_ChildForm "FN_Orders", "OpRoomRecord"
    frmPlastiSurg.Show
End Sub

Private Sub cmdLabOrders_Click()
    If piPatientID = 0 Then
        MsgBox "Please select patient first!", vbInformation
        frmAdmin.Show vbModal
        If piPatientID <> 0 Then
            frmPlastiSurg.piPatientID = piPatientID
            frmPlastiSurg.piEncounterID = piEncounterID
            frmPlastiSurg.piProviderID = piProviderID
            frmPlastiSurg.Load_ChildForm "FN_Orders", "CosmeticOrders"
            frmPlastiSurg.Show
        Else
            Exit Sub
        End If
    Else
        frmPlastiSurg.piPatientID = piPatientID
        frmPlastiSurg.piEncounterID = piEncounterID
        frmPlastiSurg.piProviderID = piProviderID
        frmPlastiSurg.Load_ChildForm "FN_Orders", "CosmeticOrders"
        frmPlastiSurg.Show
    End If
    
    Exit Sub
End Sub

Private Sub cmdLogs_Click()
    fraLogs.Visible = True
End Sub

Private Sub cmdMedication_Click()
 
    If piPatientID = 0 Then piPatientID = modSpecific.piPatientID
        If piPatientID = 0 Then
            MsgBox "Please select patient first!", vbInformation
            Exit Sub
        Else
            frmPlastiSurg.piPatientID = modSpecific.piPatientID
            frmPlastiSurg.piEncounterID = modSpecific.piEncounterID
            frmPlastiSurg.piProviderID = modSpecific.piProviderID
            frmPlastiSurg.Load_ChildForm "PT_Medication", "Prescription"
            frmPlastiSurg.Show
        End If
    Exit Sub
End Sub

Private Sub cmdMedInv_Click()
   'frmMedInventory.Show
End Sub

Private Sub cmdOperativeRep_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Load_ChildForm "FN_Orders", "OpReport"
    frmPlastiSurg.Show
End Sub

Private Sub cmdPatCheck_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Load_ChildForm "FN_Orders", "SurgSet"
    frmPlastiSurg.Show
End Sub

Private Sub cmdPatient_Click()

    Unload Me
    modSpecific.piPatientID = piPatientID
    modSpecific.piEncounterID = piEncounterID
    modSpecific.piFacilityID = piFacilityID
    modSpecific.piProviderID = piProviderID
    frmPatients.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPatientCommand_Click()
    frmAdmin.Show vbModal
    lblPatName.Caption = clsPatient.Name
End Sub

Private Sub cmdPhyExam_Click()
    fraConsult.Visible = True
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Show
End Sub

Private Sub cmdPostOpPhy_Click()
    fraConsult.Visible = True
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Show
End Sub

Private Sub cmdPreOpAneth_Click()
    fraConsult.Visible = True
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Show
End Sub

Private Sub cmdPreOpPhys_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Show vbModal
End Sub

Private Sub cmdPreCheck_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
  ' frmPlastiSurg.Load_ChildForm "FN_Orders", "AnesEvaluation"
    frmPlastiSurg.Show
End Sub

Private Sub cmdPreHP_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Load_ChildForm "FN_Orders", "HistPhy"
    If frmPlastiSurg.mbNotLoad = False Then frmPlastiSurg.Show
    Exit Sub
End Sub

Private Sub cmdProgress_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Load_ChildForm "FN_Orders", "ProgNote"
    frmPlastiSurg.Show
End Sub

Private Sub cmdTelephone_Click()
    frmAdmin.Show
    Exit Sub
End Sub

Private Sub cmdProgressNote_Click()
    frmPlastiSurg.piPatientID = piPatientID
    frmPlastiSurg.piEncounterID = piEncounterID
    frmPlastiSurg.piProviderID = piProviderID
    frmPlastiSurg.Load_ChildForm "FN_Orders", "ProgNote"
    frmPlastiSurg.Show
    Exit Sub
End Sub

Private Sub cmdSchedule_Click()
 frmPatientAppointment.Show
End Sub

Private Sub cmdScheduling_Click()
     frmPatientAppointment.Show
End Sub

Private Sub Form_Activate()

   If mFirst = True Then
        mFirst = False
   End If
   Screen.MousePointer = vbDefault

   Exit Sub
   
    If piPatientID = 0 Then piPatientID = frmPatients.piPatientID
    If piEncounterID = 0 Then piPatientID = frmPatients.piEncounterID

    If clsPatient.Name <> vbNullString Then
        frmPlasticSurgery.Caption = clsPatient.Name
    End If
    
    If clsEncounter.Provider <> vbNullString Then
        frmPlasticSurgery.Caption = frmPlasticSurgery.Caption & "      " & clsEncounter.Provider
    End If
    
    If clsEncounter.StartDate <> vbNullString Then
        frmPlasticSurgery.Caption = frmPlasticSurgery.Caption & "      " & clsEncounter.StartDate
    End If
    Exit Sub
End Sub

Public Property Get FormID() As Long
    FormID = mlngFormID
End Property
Public Property Let FormID(ByVal vlngNewValue As Long)
    mlngFormID = vlngNewValue
End Property


Private Sub Form_Load()
    Screen.MousePointer = vbHourglass


    If piPatientID = 0 Then piPatientID = frmPatients.piPatientID
    If piEncounterID = 0 Then piPatientID = frmPatients.piEncounterID
    If clsPatient.Name <> vbNullString Then
        frmPlasticSurgery.Caption = UCase(clsPatient.Name)
    End If
    
    If clsEncounter.Provider <> vbNullString Then
         frmPlasticSurgery.Caption = UCase(clsEncounter.Provider)
    End If
    
    If clsEncounter.StartDate <> vbNullString Then
         frmPlasticSurgery.Caption = UCase(clsEncounter.StartDate)
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Function Validate_Entries() As Boolean

    Validate_Entries = True

   ' If Trim(txtFirstName) = vbnullstring Then Validate_Entries = False
    'If Trim(txtLastName) = vbnullstring Then Validate_Entries = False
'    If Trim(txtCity) = vbnullstring Then Validate_Entries = False
'    If cboState.ListIndex = -1 Then Validate_Entries = False
'
End Function

'Private Function Save_Data() As Long
'Dim strSQL As String
'Dim booret As Boolean
'Dim lngRows As Long
'Dim varRS
'Dim lstrRemarks As String
'Dim strFisrtName As String
'Dim strMiddleName As String
'Dim strLastName As String
'Dim strCity As String
'Dim strFields As String
'Dim strValues As String
'Dim strTitle As String
'Dim strPhone As String
'Dim strCellPhone As String
'Dim strBirth As String
'Dim strSocSec As String
'Dim strGender As String
'Dim strAddress1 As String
'Dim strAddress2 As String
'Dim strSuffix As String
'
'    On Error GoTo Error_Save_Data
'
'    strFisrtName = UCase(Literals(Trim(txtFirstName), "'"))
'    strMiddleName = UCase(Literals(Trim(txtMiddleName), "'"))
'    strLastName = UCase(Literals(Trim(txtLastName), "'"))
'    wh add suffix
'    strSuffix = UCase(Literals(Trim(txtSuffix), "'"))
'    strCity = UCase(Literals(Trim(txtCity), "'"))
'    strAddress1 = Literals(Trim(txtAddress1), "'")
'    strAddress2 = Literals(Trim(txtAddress2), "'")
'
'    Does Patient already exists?.............................
'    strSQL = "Select count(*) From PATIENT " _
'        & "Where upper(first_name) = '" & strFisrtName & "' And " _
'        & "upper(last_name) = '" & strLastName & "' And " _
'        & "upper(city) = '" & strCity & "' And " _
'        & "state = '" & cboState.Text & "'"
'
'        lstrRemarks = strSQL
'
'        booret = adoSelectQuery(strSQL, psConnect(0), varRS, , lngRows)
'
'        If booret Then
'            If lngRows > 0 Then
'                If varRS(0, 0) > 0 Then
'                    MsgBox "Patient already exists,", vbExclamation, "Save Data"
'                    Exit Function
'                End If
'            End If
'        End If
'
'    Insert Record.......................
'    strFields = "title, first_name, middle_name, last_name, address1, address2 ," _
'        & "city, state, zipcode, zipcode_ext, country, area_code, phone, " _
'        & "cell_phone, drv_license, birth, social_security, gender, email, patient_type"
'
'    strFields = "title, first_name, middle_name, last_name, suffix,address1, address2 ," _
'        & "city, state, zipcode, zipcode_ext, country, county " _
'        & "drv_license, birth, social_security, gender, email, race, room_number,SVP, DOC, patient_ext_ref, status, Other_Info1"
'
'    Title---------------
'    If optMr.Value Then strTitle = "Mr."
'    If optMrs.Value Then strTitle = "Mrs."
'    If optMs.Value Then strTitle = "Ms."
'    If optDr.Value Then strTitle = "Dr."
'
'    strPhone = txtPhone1 & "-" & txtPhone2
'    strCellPhone = txtCellAreaCode & "-" & txtCellPhone1 & "-" & txtCellPhone2
'
'    strBirth = "'" & txtMonth & "/" & txtDay & "/" & txtYear & "'"
'    If Not IsDate(txtMonth & "/" & txtDay & "/" & txtYear) Then strBirth = "NULL"
'
'    strSocSec = txtSS1 & "-" & txtSS2 & "-" & txtSS3
'
'    If optMale.Value Then strGender = "M"
'    If optFemale.Value Then strGender = "F"
'
'
'      strValues = "'" & strTitle & "','" & strFisrtName & "','" & strMiddleName & "','" & strLastName & "','" & strSuffix & "'" _
'        & strAddress1 & "','" & strAddress2 & "','" & txtCity & "','" & cboState.Text & "','" & txtZipCode & "','" _
'        & txtZipCodeExt & "','" & cboCountry.Text & "','" & cboCounty.Text & "','" _
'        & txtDriverLicense & "'," & strBirth & ",'" & strSocSec & "','" & strGender & "','" & txtEmailAddress & "','" & cboPatientRace.Text & "','" _
'        & cboHousingUnit.Text & "', '" & txtSVP.Text & "','" & txtDOC.Text & "','" & txtSVP.Text & "','" & cboStatus.Text & "'"
'
'    strSQL = "Insert into PATIENT (" & strFields & ") Values(" & strValues & ")"
'    lstrRemarks = strSQL
'    txtCurrRoom.Text = cboHousingUnit.Text
'    adoActionQuery strSQL, psConnect(0)
'
'    Get PatientID.............................
'    strSQL = "Select max(patient_id) From Patient Where last_name = '" & strLastName & "'"
'    lstrRemarks = strSQL
'    booret = adoSelectQuery(strSQL, psConnect(0), varRS, , lngRows)
'
'    If lngRows > 0 Then Save_Data = varRS(0, 0)
'    Load_HousingUnit
'
'    Exit Function
'
'Error_Save_Data:
'    LogErrorHandler _
'     psConnect(0), _
'     Me.Name & " Save_Data", _
'     Err.Number, _
'     Err.Description, _
'     App.EXEName, _
'     App.Path, _
'     App.Major & "." & App.Minor & "." & App.Revision, _
'     piPatientID, _
'     piEncounterID, _
'     Get_UserName(), _
'     Get_ComputerName(), _
'     lstrRemarks
Private Function Get_FormID(vstrFormName As String) As Integer
Dim strSQL As String
Dim varRS
Dim lngRows As Long
Dim booret As Boolean
Dim lstrRemarks As String


    strSQL = "Select form_id From Forms Where form_caption = '" & vstrFormName & "'"
    lstrRemarks = strSQL
    
    booret = SelectQuery(strSQL, psConnect(1), varRS, lngRows)
    
    
    If booret Then
        If lngRows > 0 Then
            Get_FormID = varRS(0, 0)
        Else
            MsgBox "Form " & vstrFormName & " not found.", vbExclamation
        End If
    Else
        LogErrorHandler _
         psConnect(0), _
         Me.Name & " Get_FormID", _
         Err.Number, _
         Err.Description, _
         App.EXEName, _
         App.Path, _
         App.Major & "." & App.Minor & "." & App.Revision, _
         piPatientID, _
         piEncounterID, _
         Get_UserName(), _
         Get_ComputerName(), _
         lstrRemarks
    End If
    
End Function

Private Sub Write_Control()
End Sub

Private Sub grdPatientStatus_Advance(ByVal AdvanceNext As Boolean)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.Objects
      Set obj = Nothing
    Next obj
  Exit Sub
End Sub

Private Sub picAdmin_Click()
    fraAdmin.Visible = True
End Sub

Private Sub picConsult_Click()

    fraConsult.Visible = True
End Sub

Private Sub picOR_Click()
    fraOR.Visible = True
End Sub

Private Sub picOrders_Click()
    fraOrders.Visible = True
End Sub


Private Sub picResident_Click()
    Unload Me
    modSpecific.piPatientID = piPatientID
    modSpecific.piEncounterID = piEncounterID
    modSpecific.piFacilityID = piFacilityID
    modSpecific.piProviderID = piProviderID
    frmPatients.Show
    Screen.MousePointer = vbDefault
    
End Sub





Private Sub PT_XP3_Click()

End Sub


