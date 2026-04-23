VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPatientAppointment_Old 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Appointments"
   ClientHeight    =   9585
   ClientLeft      =   4410
   ClientTop       =   2640
   ClientWidth     =   15015
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatientAppointment_Old.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraGetAppts 
      Height          =   7305
      Left            =   4050
      TabIndex        =   54
      Top             =   2190
      Visible         =   0   'False
      Width           =   10515
      Begin MSComCtl2.MonthView ccSchedDate 
         Height          =   2670
         Left            =   90
         TabIndex        =   55
         Top             =   180
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   4710
         _Version        =   393216
         ForeColor       =   8388608
         BackColor       =   16744448
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   197394433
         TitleBackColor  =   16761024
         TitleForeColor  =   8388608
         TrailingForeColor=   8421504
         CurrentDate     =   41640
      End
      Begin FPSpreadADO.fpSpread grdAvailAppt 
         Height          =   7065
         Left            =   3480
         TabIndex        =   56
         Top             =   150
         Width           =   6930
         _Version        =   524288
         _ExtentX        =   12224
         _ExtentY        =   12462
         _StockProps     =   64
         ColHeaderDisplay=   0
         DAutoCellTypes  =   0   'False
         DAutoFill       =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
         DInformActiveRowChange=   0   'False
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
         SpreadDesigner  =   "frmPatientAppointment_Old.frx":08CA
         UserResize      =   0
         Appearance      =   2
         TextTip         =   4
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin PT_XP_Button.PT_XP cmdSelectAvail 
         Height          =   405
         Left            =   750
         TabIndex        =   57
         Top             =   6750
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmPatientAppointment_Old.frx":0DED
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Select"
         CaptionAlignment=   5
         CaptionPosition =   4
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatientAppointment_Old.frx":1187
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCloseAvail 
         Height          =   405
         Left            =   1860
         TabIndex        =   58
         Top             =   6750
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmPatientAppointment_Old.frx":14D9
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatientAppointment_Old.frx":1EEB
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.Frame fraSchLegend 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scheduling Color Codes"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   90
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   4005
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   37
         Top             =   1050
         Width           =   285
      End
      Begin VB.PictureBox picPreOp 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   29
         ToolTipText     =   "Pre Operative Patient"
         Top             =   3360
         Width           =   285
      End
      Begin VB.PictureBox picPostOp 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   28
         Top             =   3270
         Width           =   285
      End
      Begin VB.PictureBox picExstNewPat 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   27
         ToolTipText     =   "Established Patient New Consultation"
         Top             =   3000
         Width           =   285
      End
      Begin VB.PictureBox picNewPatient 
         BackColor       =   &H00FFC0FF&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   26
         ToolTipText     =   "New Initial Patient"
         Top             =   2700
         Width           =   285
      End
      Begin VB.PictureBox picReturn 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   25
         Top             =   2040
         Width           =   285
      End
      Begin VB.PictureBox picSurgery 
         BackColor       =   &H00FFFF80&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   24
         Top             =   2040
         Width           =   285
      End
      Begin VB.PictureBox picHospital 
         BackColor       =   &H00FFFF80&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   22
         Top             =   780
         Width           =   285
      End
      Begin VB.PictureBox picBlock 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   20
         Top             =   780
         Width           =   285
      End
      Begin VB.PictureBox picPresent 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   13
         Top             =   210
         Width           =   285
      End
      Begin VB.PictureBox picCheckOut 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   210
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   12
         Top             =   510
         Width           =   285
      End
      Begin VB.PictureBox picOffice 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   11
         Top             =   510
         Width           =   285
      End
      Begin VB.PictureBox picCancel 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   2250
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   10
         Top             =   210
         Width           =   285
      End
      Begin PT_XP_Button.PT_XP cmdCancel 
         Height          =   345
         Left            =   1230
         TabIndex        =   53
         Top             =   1410
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
         Picture         =   "frmPatientAppointment_Old.frx":223D
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatientAppointment_Old.frx":2C4F
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblBus 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Non Patient/Business"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   600
         TabIndex        =   38
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label lblNew 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "New Patient"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   540
         TabIndex        =   35
         Top             =   3000
         Width           =   930
      End
      Begin VB.Label lblEstablishNew 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Est New Patient"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   540
         TabIndex        =   34
         Top             =   3300
         Width           =   1230
      End
      Begin VB.Label lblPostOp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Post-Op"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   2595
         TabIndex        =   33
         Top             =   3690
         Width           =   660
      End
      Begin VB.Label lblPreOp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Pre-Op"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   2610
         TabIndex        =   32
         Top             =   3390
         Width           =   555
      End
      Begin VB.Label lblFU 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "FollowUp Exam"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   525
         TabIndex        =   31
         Top             =   2070
         Width           =   1170
      End
      Begin VB.Label lblSurgery 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Surgery"
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
         Left            =   2625
         TabIndex        =   30
         Top             =   2070
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2655
         TabIndex        =   23
         Top             =   810
         Width           =   615
      End
      Begin VB.Label lblBlockTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Block Time"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   585
         TabIndex        =   21
         Top             =   810
         Width           =   810
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Encounter Closed"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2610
         TabIndex        =   17
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label lblOffice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Office"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   2655
         TabIndex        =   16
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Checked out"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   570
         TabIndex        =   15
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Present"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   570
         TabIndex        =   14
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame fraFacility 
      Height          =   3045
      Left            =   7110
      TabIndex        =   48
      Top             =   360
      Visible         =   0   'False
      Width           =   5895
      Begin PT_XP_Button.PT_XP cmdCloseFac 
         Height          =   345
         Left            =   2550
         TabIndex        =   49
         Top             =   2610
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
         Picture         =   "frmPatientAppointment_Old.frx":2FA1
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatientAppointment_Old.frx":39B3
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdFacility 
         Height          =   2355
         Left            =   150
         TabIndex        =   50
         Top             =   210
         Width           =   5565
         _Version        =   524288
         _ExtentX        =   9816
         _ExtentY        =   4154
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
         SpreadDesigner  =   "frmPatientAppointment_Old.frx":3D05
         UserResize      =   0
         Appearance      =   2
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
         CellNoteIndicator=   2
      End
   End
   Begin VB.TextBox cctFacilities 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7110
      TabIndex        =   47
      Top             =   30
      Width           =   4845
   End
   Begin VB.TextBox cctProvider 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7110
      TabIndex        =   44
      Top             =   420
      Width           =   4845
   End
   Begin VB.Frame fraProvider 
      Height          =   3045
      Left            =   7110
      TabIndex        =   41
      Top             =   360
      Visible         =   0   'False
      Width           =   5895
      Begin PT_XP_Button.PT_XP cmdCloseList 
         Height          =   345
         Left            =   2550
         TabIndex        =   42
         Top             =   2610
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
         Picture         =   "frmPatientAppointment_Old.frx":4233
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatientAppointment_Old.frx":4C45
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpread grdProviders 
         Height          =   2355
         Left            =   150
         TabIndex        =   43
         Top             =   210
         Width           =   5565
         _Version        =   524288
         _ExtentX        =   9816
         _ExtentY        =   4154
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
         SpreadDesigner  =   "frmPatientAppointment_Old.frx":4F97
         UserResize      =   0
         Appearance      =   2
         TextTipDelay    =   100
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
   End
   Begin VB.Timer schTimer 
      Interval        =   60000
      Left            =   15990
      Top             =   8880
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   3120
      Left            =   30
      TabIndex        =   19
      Top             =   30
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   5503
      _Version        =   393216
      ForeColor       =   8388608
      BackColor       =   16744448
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      StartOfWeek     =   197394433
      TitleBackColor  =   16761024
      TitleForeColor  =   8388608
      TrailingForeColor=   8421504
      CurrentDate     =   38806
   End
   Begin VB.Timer strTelTime 
      Interval        =   60000
      Left            =   4950
      Top             =   8970
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16110
      Top             =   12660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin PT_XP_Button.PT_XP cmdCurrentWeek 
      Height          =   435
      Left            =   8400
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
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
      Picture         =   "frmPatientAppointment_Old.frx":54C5
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Current Week"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatientAppointment_Old.frx":585F
      ShowFocus       =   -1  'True
   End
   Begin RichTextLib.RichTextBox rtfChanges 
      Height          =   3705
      Left            =   19290
      TabIndex        =   6
      Top             =   2670
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6535
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmPatientAppointment_Old.frx":5BB1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PT_XP_Button.PT_XP cmdNextWeek 
      Height          =   465
      Left            =   9690
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Picture         =   "frmPatientAppointment_Old.frx":5C35
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Next Week"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatientAppointment_Old.frx":5FCF
      PictureAlignment=   5
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdData 
      DragIcon        =   "frmPatientAppointment_Old.frx":6321
      Height          =   5265
      Left            =   4050
      TabIndex        =   0
      Top             =   1350
      Width           =   10995
      _Version        =   524288
      _ExtentX        =   19394
      _ExtentY        =   9287
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      ColHeaderDisplay=   0
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
      MaxCols         =   1
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmPatientAppointment_Old.frx":6763
      UnitType        =   0
      UserResize      =   0
      Appearance      =   2
      TextTip         =   3
      TextTipDelay    =   25
      ScrollBarStyle  =   3
   End
   Begin PT_XP_Button.PT_XP cmdPrevWeek 
      Height          =   435
      Left            =   7110
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
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
      Picture         =   "frmPatientAppointment_Old.frx":6CE8
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Previous Week"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatientAppointment_Old.frx":7082
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRepair 
      Height          =   405
      Left            =   19260
      TabIndex        =   18
      Top             =   4980
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Repair"
      DepthEvent      =   1
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdMessages 
      Height          =   2925
      Left            =   4050
      TabIndex        =   36
      Top             =   6630
      Width           =   10995
      _Version        =   524288
      _ExtentX        =   19394
      _ExtentY        =   5159
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
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   10
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarMaxAlign=   0   'False
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmPatientAppointment_Old.frx":73D4
      UnitType        =   0
      UserResize      =   0
      Appearance      =   2
      TextTipDelay    =   60000
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdStop 
      Height          =   495
      Left            =   30
      TabIndex        =   39
      Top             =   3240
      Width           =   1185
      _ExtentX        =   2090
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
      Picture         =   "frmPatientAppointment_Old.frx":7947
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Stop Refresh"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatientAppointment_Old.frx":7EE1
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdProvider 
      Height          =   345
      Left            =   5940
      TabIndex        =   45
      Top             =   420
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPatientAppointment_Old.frx":8233
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Providers"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatientAppointment_Old.frx":87CD
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdFacility 
      Height          =   345
      Left            =   5910
      TabIndex        =   46
      Top             =   30
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPatientAppointment_Old.frx":8B1F
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Facilities"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatientAppointment_Old.frx":90B9
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdChanges 
      Height          =   4755
      Left            =   60
      TabIndex        =   51
      ToolTipText     =   "Select a row and right click to remove an item."
      Top             =   4770
      Width           =   3945
      _Version        =   524288
      _ExtentX        =   6959
      _ExtentY        =   8387
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      ColHeaderDisplay=   0
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
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
      FormulaSync     =   0   'False
      GrayAreaBackColor=   16777215
      GridColor       =   14737632
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   4
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   0
      ShadowText      =   0
      SpreadDesigner  =   "frmPatientAppointment_Old.frx":940B
      UnitType        =   0
      UserResize      =   0
      Appearance      =   1
      TextTipDelay    =   50
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
   Begin VB.Label lblCurScheAppt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment Count"
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
      Height          =   240
      Left            =   150
      TabIndex        =   52
      Top             =   4080
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Label lblApptDate 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Height          =   300
      Left            =   4110
      TabIndex        =   40
      Top             =   -720
      Width           =   3345
   End
   Begin VB.Label lblSchedChange 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scheduling Changes"
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
      Left            =   17610
      TabIndex        =   7
      Top             =   2010
      Width           =   1935
   End
   Begin VB.Label lblProvider 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   10110
      TabIndex        =   5
      Top             =   13800
      Width           =   5295
   End
   Begin VB.Label lblSelectDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Select Date"
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
      Left            =   0
      TabIndex        =   2
      Top             =   -3510
      Width           =   4905
   End
   Begin VB.Label lblLocation 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Facility:"
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
      Left            =   18150
      TabIndex        =   1
      Top             =   3600
      Width           =   825
   End
   Begin VB.Menu mnu_Color 
      Caption         =   "Color Codes"
   End
   Begin VB.Menu mnu_New_Patient 
      Caption         =   "New Patient"
   End
   Begin VB.Menu mnu_Block_Time 
      Caption         =   "Manage Block Time"
   End
   Begin VB.Menu mnu_Message 
      Caption         =   "Messages"
   End
   Begin VB.Menu mnu_Print 
      Caption         =   "Print"
   End
   Begin VB.Menu mnu_Schedule_Patient 
      Caption         =   "Schedule Patient"
   End
   Begin VB.Menu mnu_AvailAppt 
      Caption         =   "Available Appointments"
      Visible         =   0   'False
   End
   Begin VB.Menu mnu_View 
      Caption         =   "Views by Provider"
      Begin VB.Menu mnu_Day 
         Caption         =   "By Day"
      End
      Begin VB.Menu mnu_Week 
         Caption         =   "By Week"
      End
   End
End
Attribute VB_Name = "frmPatientAppointment_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Private mStartHour() As String
Private mEndHour() As String
Private mStartIdx As Integer
Private mEndIdx As Integer
Private mProviders() As String
Public vTime As Variant
Public mlngCurProvID As Long
Public strSelDate As String
Private mLoadPatient As Boolean
Private mChange As Boolean
Public mBeenHere As Boolean
Private mlngGridRow As Long
Private mlngGridCol As Long
Private mlngMessRow As Long
Private mlngMessCol As Long
Private mEnterRow As Boolean
Private mstrValidAppoint As String
Private mFirstLoad As Boolean
Public mstrApptMode As String
Public mbooTimeAvailable As Boolean
Private mstrBlockReason As String
Private mbooDoubleBook As Boolean
Private mbooReadWeek As Boolean
Private strPrevWeek As String
Private mmlngClickPrev As Long
Private mlngClickNext As Long
Private mGridClick As Boolean
Public mstrStatus As String
Public mlngPatientID As Long
Private mbooPatValid As Boolean
Private mlngActCol As Long
Private Color As OLE_COLOR
Private mstrPresent As String
Public mlngClickPrev As Long
Private mboobyppas As Boolean
Private mbooPrintPreview As Boolean
Dim mbooDup As Boolean
Private Const NO_SIDES = 0
Private Const LEFT_SIDE = 1
Private Const RIGHT_SIDE = 2
Private Const TOP_SIDE = 4
Private Const BOTTOM_SIDE = 8
Private Const ALL_SIDES = 16
Private Const SOLID = 1
Private mlngPatCnt As Long
Private mstrCheckTime As String
Private mbooSchChange As Boolean
Private mbooTeleChange As Boolean
Private udtColor As udtColor
Private Type udtColor
    Office As Long
    Hospital As Long
    Vacation As Long
    Cancelled As Long
    CheckOut As Long
    NoShow As Long
    Present As Long
    Lunch As Long
    PostOp As Long
    PreOp As Long
    NewPatient As Long
    ExistPatient As Long
    Surgery As Long
    FollowUp As Long
End Type
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
End Type
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
Private clsShedGrid As clsShedGrid
Private Type clsShedGrid
    CellSpan As String
    CellTag As String
    CellNote As String
    BackColor As Long
    Text As String
    TextNote As String
    List As String
    Duration As Long
    InsertRow As Long
    ProvCol As Long
    ProvCol2 As Long
    SpanRow As Long
    ApptDate As String
    StartRow As Long
    EndRow As Long
    Col As Long
    Row As Long
    RowName As String
    ColHeader As String
    CellType As String
End Type

Private Function Change_Providers()
Dim lngProvID As Long
Dim selProvider As String
Dim i%
Dim lngProv As Long
Dim booFound As Boolean
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim bNext As Boolean

    On Error Resume Next
    If Trim(cctProvider.Text) <> vbNullString Then
        lngProv = grdProviders.GetRowItemData(grdProviders.ActiveRow)
        If lngProv <> mlngLogonUser And clsEDI_Client.Site_ID <> 1000 Then
            If lngProv > 0 Then
                If clsEDI_Client.Site_ID = "2777" And Get_SuperProv Then
                    If piProviderID <> lngProv Then
                        MsgBox "Facility rules do not allow access to other providers' schedule. Please contact your administrator for information regarding other providers' schedule!", vbInformation, "PatientTrac Schedule"
                        udtAppointment.ProviderID = piProviderID
                        Get_Providers udtAppointment.ProviderID
                        cctProvider.Text = clsProvider.Name
                        With grdProviders
                            For lngRow = 1 To .MaxRows
                                .BlockMode = True
                                    .Col = 1
                                    .Col2 = 2
                                    .Row = lngRow
                                    .Row2 = lngRow
                                    .BackColor = NormalColor
                                .BlockMode = True
                            Next lngRow
                            For lngRow = 1 To .MaxRows
                                .Col = 1
                                .Row = lngRow
                                If .GetRowItemData(lngRow) = udtAppointment.ProviderID Then
                                    .BlockMode = True
                                        .Col = 1
                                        .Col2 = 2
                                        .Row = lngRow
                                        .Row2 = lngRow
                                        .BackColor = grdCurItem
                                    .BlockMode = True
                                    Exit For
                                End If
                            Next lngRow
                        End With
                    Else
                        udtAppointment.ProviderID = lngProv
                        piProviderID = udtAppointment.ProviderID
                        Set_Headers
                        Format_Messages
                        Read_Messages
                        Get_Mode
                    End If
                Else
                    Get_SecurityRole mlngLogonUser
                    If UCase(clsPracticeRoles.Permission) <> "ALL" Then
                          clsPracticeRoles.Permission = vbNullString
                          Get_Indiv_Priv "SCHEDULE", mlngLogonUser
                          If UCase(clsPracticeRoles.Permission) <> "ALL" Then
                            If Not TextFound("Schedule", clsPracticeRoles.Permission, Len("Schedule"), True) Then
                                MsgBox "Facility rules do not allow access to other providers' schedule. Please contact your administrator for information regarding other providers' schedule!", vbInformation, "PatientTrac EMR"
                                With grdProviders
                                    For lngRow = 1 To .MaxRows
                                        If mlngLogonUser = .GetRowItemData(lngRow) Then
                                            .GetText 1, lngRow, varCellValue
                                            .BlockMode = True
                                                .Col = 1
                                                .Col2 = 2
                                                .Row = lngRow
                                                .Row2 = lngRow
                                                .BackColor = grdCurItem
                                            .BlockMode = False
                                            cctProvider.Text = Trim(varCellValue)
                                            Exit For
                                        End If
                                    Next lngRow
                                End With
                            End If
                        End If
                    Else
                        udtAppointment.ProviderID = lngProv
                        piProviderID = udtAppointment.ProviderID
                        Set_Headers
                        Format_Messages
                        Read_Messages
                        Get_Mode
                    End If
                End If
            End If
        Else
            udtAppointment.ProviderID = lngProv
            piProviderID = udtAppointment.ProviderID
            Set_Headers
            Format_Messages
            Read_Messages
            Get_Mode
        End If
    End If
    mbooByPass = False
    Exit Function
End Function

Private Function Check_Conflict_Facility()
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim i%
Dim lngDateDiff As Long
Dim strRepDate As String
Dim strStart As String
Dim strEnd As String
Dim lngCnt As Long
Dim lngTime As Double
Dim lngBlockFac As Long


    On Error GoTo Error_Check_Conflict_Facility
    Screen.MousePointer = vbHourglass
    If udtAppointment.SchedDate = vbNullString Then
        udtAppointment.SchedDate = Format(Now, "mm/dd/yyyy")
    Else
        udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    End If
    If cnBlock.State = adStateClosed Then cnBlock.open psConnect(0)
    strSQL = "Appt_Check_Provider_Conflict " & udtAppointment.ProviderID & ", '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "'"
    Set rsBlock = cnBlock.Execute(strSQL)
    With rsBlock
    Do Until .EOF
        Set Flds = rsBlock.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "STATUS"
                                udtSchedPatient.Status = Trim(Fld.Value)
                            Case "FACILITY_ID"
                                udtAppointment.FacilityID = Fld.Value
                            Case "APPT_TIME"
                                udtAppointment.SchedTime = Fld.Value
                                udtAppointment.SchedTime = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                            Case "DURATION"
                                udtSchedPatient.Duration = Fld.Value
                            Case "PATIENT_TYPE"
                                udtAppointment.PatientType = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtAppointment.FacilityID > 0 Then
                If (udtAppointment.FacilityID <> udtFacility.FacID) And udtSchedPatient.Status <> "No Show" Then
                    lngBlockFac = udtAppointment.FacilityID
                    lngCnt = lngCnt + 1
                    If lngCnt = 1 Then
                        strStart = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                        strEnd = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                        strRepDate = udtSchedPatient.Duration
                    End If
                    strEnd = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                    strRepDate = udtSchedPatient.Duration
                End If
            End If
        Loop
    End With
    If lngCnt > 0 Then
        If strStart <> vbNullString And strEnd <> vbNullString Then
            udtAppointment.SchedTime = udtAppointment.SchedDate & " " & strEnd
            If strEnd <> vbNullString Then
                lngTime = Val(Left(strRepDate, 2))
                strEnd = DateAdd("n", lngTime, Format(udtAppointment.SchedTime, "mm/dd/yyyy hh:mm AM/PM"))
                strEnd = Format(strEnd, "hh:mm AM/PM")
                Block_Facility strStart, strEnd, lngBlockFac
            End If
        End If
    End If
    If cnBlock.State = adStateOpen Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Check_Conflict_Facility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Check_Conflict_Facility", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Block_Facility(Optional vstrStart As String, Optional vstrEnd As String, Optional Facility As Long)
Dim i%
Dim lngProviderID As Long
Dim strSQL As String
Dim vTime
Dim vDate
Dim BlockColor As Long
Dim lngStartRow As Long
Dim lngEndRow As Long
Dim lngNoRows  As Long
Dim nDays As Double
Dim Arr
Dim lngDateDiff As Long
Dim lngBlockCol As Long
Dim lngCol As Long
Dim lngOverCol As Long
Dim lngRow As Long
Dim varCellValue
Dim booCancel As Boolean
Dim strReason As String
Dim strTime As String
Dim strStart As String
Dim strEnd As String
Dim bExit As Boolean
Dim lngTest As Long
Dim strConflict As String
Dim lngChgCol As Long

    On Error GoTo Error_Block_Facility
    Screen.MousePointer = vbHourglass
    For lngCol = 1 To grdData.MaxCols
        grdData.Col = lngCol
        If mstrApptMode = "Day" Then
            grdData.Row = SpreadHeader + 3
        ElseIf mstrApptMode = "Weekly" Then
             grdData.Row = SpreadHeader + 4
        End If
        If grdData.CellTag <> vbNullString Then
            If udtAppointment.ProviderID = grdData.CellTag Then
                If lngBlockCol = 0 Then
                    lngBlockCol = lngCol
                Else
                    lngOverCol = lngCol
                    Exit For
                End If
            End If
        End If
    Next lngCol
    If lngOverCol = 0 Then lngOverCol = 1
    Get_Facilities udtAppointment.FacilityID
    If vstrStart <> vbNullString Then
        strStart = Format(vstrStart, "hh:mm AM/PM")
        strEnd = Format(vstrEnd, "hh:mm AM/PM")
    End If
     For lngRow = 1 To grdData.MaxRows
        grdData.GetText 0, lngRow, vTime
        If UCase(vTime) = UCase(strStart) Then
            lngStartRow = lngRow
        End If
        grdData.GetText 0, lngRow, vTime
        If UCase(vTime) = UCase(strEnd) Then
            lngEndRow = lngRow
        End If
    Next lngRow
    If Facility > 0 Then
        Get_Facilities Facility
    End If
    strReason = "Provider scheduled in" & " " & UCase(clsFacility.Name)
    Get_Facilities udtFacility.FacID
    If lngEndRow = 0 Then lngEndRow = grdData.MaxRows
    If lngStartRow = 0 Then lngStartRow = 1
    lngNoRows = lngEndRow - lngStartRow
    BlockColor = &HFFFFC0
    If lngStartRow >= 1 And lngEndRow > 1 Then
        For lngRow = lngStartRow To lngEndRow
            clsGridFormat.CellNote = vbNullString
            grdData.GetText lngCol, lngRow, varCellValue
            If lngRow = lngStartRow Then
                    grdData.Col = lngBlockCol
                    grdData.Row = lngRow
                    grdData.GetText lngBlockCol, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString And Not TextFound(strReason, CStr(varCellValue), Len(strReason), True) Then
                        grdData.Col = lngCol
                        grdData.Row = lngRow
                        If grdData.CellTag <> vbNullString Then
                            udtSchedPatient.IntId = grdData.CellTag
                            If Trim(udtSchedPatient.IntId) <> vbNullString Then
                                If grdData.CellNote <> vbNullString Then
                                    udtAppointment.SchedType = grdData.CellNote
                                End If
                                If udtAppointment.SchedType = vbNullString Then udtAppointment.SchedType = "Patient"
                                strConflict = vbNullString
                                grdData.GetText 0, lngRow, varCellValue
                                strConflict = varCellValue
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
                                If Trim(udtSchedPatient.Name) <> vbNullString Then
                                    bExit = False
                                    With grdChanges
                                        For lngTest = 1 To .MaxRows
                                            .Col = 1
                                            .Row = lngTest
                                            If .GetRowItemData(lngTest) = Val(udtSchedPatient.IntId) Then
                                                bExit = True
                                                Exit For
                                            End If
                                        Next lngTest
                                        If Not bExit Then
                                            clsGridFormat.CellNote = "Appointment at" & " " & strConflict & " " & "Provider scheduled in" & " " & UCase(clsFacility.Name)
                                            lngTest = .MaxRows + 1
                                            .MaxRows = lngTest
                                            .SetRowItemData lngTest, Val(udtSchedPatient.IntId)
                                            lngChgCol = 1
                                            Format_CellTypeStatic grdChanges, lngTest, lngChgCol, "1", Trim(udtSchedPatient.Name), "TypeHAlignLeft", "", "", 0, udtAppointment.SchedType
                                            grdChanges.SetCellBorder lngChgCol, lngTest, lngChgCol, lngTest, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                            lngChgCol = 2
                                            Format_CellTypeEdit grdChanges, lngTest, lngChgCol, "1", Trim(clsGridFormat.CellNote), "TypeHAlignCenter"
                                            grdChanges.SetCellBorder lngChgCol, lngTest, lngChgCol, lngTest, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                            lngCol = 3
                                            If udtSchedPatient.CellPhone <> vbNullString Then
                                                Format_CellTypeStatic grdChanges, lngTest, lngCol, "1", udtSchedPatient.CellPhone, "TypeHAlignCenter"
                                            Else
                                                Format_CellTypeStatic grdChanges, lngTest, lngCol, "1", udtSchedPatient.Phone, "TypeHAlignCenter"
                                                grdChanges.SetCellBorder lngCol, lngTest, lngCol, lngTest, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                            End If
                                        End If
                                    End With
                                End If
                            End If
                        End If
                    End If
'                    If lngBlockCol = 0 Then lngBlockCol = 1
'                    Format_CellTypeEdit_RowSpan grdData, lngRow, lngBlockCol, CStr(lngNoRows), strReason, "TypeHAlignCenter", "BlockMode", grdBlock, "", CStr(lngOverCol)
'                    grdData.SetCellBorder lngBlockCol, lngStartRow, lngBlockCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                    grdData.SetCellBorder lngBlockCol, lngEndRow, lngBlockCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Exit For
                End If
        Next lngRow
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Block_Facility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Block_Facility", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Function Format_Facilities()
Dim lngColWidth As Long
Dim varCellValue
Dim varHeader
Dim lngCnt  As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim strHeaders As String

    On Error GoTo Error_Format_Problem
    Screen.MousePointer = vbHourglass
    With grdFacility
        .RemoveCellSpan 1, SpreadHeader
        .RemoveCellSpan 1, SpreadHeader + 2
        .ClearRange 1, SpreadHeader, 2, .MaxRows, False
        .ClearRange 1, SpreadHeader + 2, 2, .MaxRows, False
        .Redraw = False
        .ProcessTab = True
        .FontBold = True
        .FontName = "Arial"
        .RowHeadersShow = False
        .ColHeadersShow = True
        .MaxCols = 2
        .MaxRows = 0
        .ColHeaderRows = 4
        .ColHeadersAutoText = DispBlank
    End With
    grdFacility.ColWidth(1) = 25 * mlngSystemFileSize
    grdFacility.ColWidth(2) = 18 * mlngSystemFileSize
    lngCol = 1
    lngRow = SpreadHeader
    Format_CellTypeEdit grdFacility, lngRow, lngCol, "2", "APPOINTMENT FACILITIES", "TypeHAlignCenter", " ", HeaderColor
    grdFacility.SetCellBorder lngCol, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 1, grdFacility
    lngCol = 1
    lngRow = SpreadHeader + 2
    Format_CellTypeEdit grdFacility, lngRow, lngCol, "1", "FACILITY  NAME", "TypeHAlignCenter", "FACILITY_NAME", HeaderColor
    grdFacility.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 2
    Format_CellTypeEdit grdFacility, lngRow, lngCol, "1", "FACILITY ID", "TypeHAlignCenter", "FACILITY_ID", HeaderColor
    grdFacility.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 3, grdFacility
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_Problem:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_Problem", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function

Private Function Format_Providers()
Dim lngColWidth As Long
Dim varCellValue
Dim varHeader
Dim lngCnt  As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim strHeaders As String

    On Error GoTo Error_Format_Providers
    Screen.MousePointer = vbHourglass
    With grdProviders
        .Redraw = False
        .ProcessTab = True
        .FontBold = True
        .FontName = "Arial"
        .RowHeadersShow = False
        .ColHeadersShow = True
        .MaxCols = 2
        .MaxRows = 0
        .ColHeaderRows = 4
        .ColHeadersAutoText = DispBlank
    End With
    grdProviders.ColWidth(1) = 25.2 * mlngSystemFileSize
    grdProviders.ColWidth(2) = 18 * mlngSystemFileSize
    lngCol = 1
    lngRow = SpreadHeader
    Format_CellTypeEdit grdProviders, lngRow, lngCol, "2", "APPOINTMENT PROVIDERS", "TypeHAlignCenter", "", HeaderColor
    grdProviders.SetCellBorder lngCol, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 1, grdProviders
    lngCol = 1
    lngRow = SpreadHeader + 2
    Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", "PROVIDER  NAME", "TypeHAlignCenter", "PROVIDER_NAME", HeaderColor
    grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 2
    Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", "PROVIDER ID", "TypeHAlignCenter", "PROVIDER_ID", HeaderColor
    grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 3, grdProviders
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_Providers:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_Providers", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function
Private Function Is_Closed_Encounter() As Boolean
Dim sDate As String
Dim strSQL As String
Dim cnStatus As New ADODB.Connection
Dim rsStatus As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim strEnctTime As String
Dim lngEnctrID As Long
Dim iAnswer As Integer
Dim strDate As String
Dim Idx As Integer
Dim lngProvID As Long

    On Error GoTo Error_Is_Closed_Encounter
    Screen.MousePointer = vbHourglass
    Is_Closed_Encounter = False
    If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
    strSQL = "SELECT ENCNTR_END_DATE FROM ENCOUNTER WHERE PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
        & "CONVERT(VARCHAR(10), ENCNTR_START_DATE, 101) = '" & Format(clsShedGrid.ApptDate, "mm/dd/yyyy") & "' AND " _
        & "PROVIDER_ID = " & udtAppointment.ProviderID & " AND PURGE = 0  AND FACILITY_ID = " & udtFacility.FacID
        Set rsStatus = cnStatus.Execute(strSQL)
        With rsStatus
            If Not .EOF Then
                If Not IsNull((!ENCNTR_END_DATE)) Then
                    strEnctTime = (!ENCNTR_END_DATE)
                End If
            End If
        End With
        If Format(strEnctTime, "mm/dd/yyyy") = Format(clsShedGrid.ApptDate, "mm/dd/yyyy") Then
            Is_Closed_Encounter = True
        End If
    If cnStatus.State = adStateOpen Then
        cnStatus.Close
        Set cnStatus = Nothing
        Set rsStatus = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Is_Closed_Encounter:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Is_Closed_Encounter", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Load_SchedPatient(pCol As Long, pRow As Long)
Dim vValue
Dim iAnswer As Integer
Dim lngPatientID As Long
Dim vDate
Dim Arr

    On Error GoTo Error_Load_SchedPatient
    Screen.MousePointer = vbHourglass
    mChange = False
    grdData.GetText pCol, pRow, vValue
    grdData.GetText 0, pRow, vTime
    If vTime = vbNullString Then
        grdData.GetText 0, pRow - 1, vTime
    End If
    If vTime <> vbNullString Then
        udtAppointment.SchedTime = vTime
    End If
    lblApptDate.Caption = "Date:" & " " & Format(udtAppointment.SchedDate, "mmmm dd, yyyy")
    lblApptDate.AutoSize = True
    If mstrApptMode <> "Day" Then
        grdData.GetText pCol, SpreadHeader + 2, vDate
        Arr = Split(vDate, "-")
        If UBound(Arr) > 0 Then
            vDate = Arr(1)
            udtAppointment.SchedDate = vDate
        Else
            udtAppointment.SchedDate = vDate
        End If
    End If
    If mLoadPatient And mGridClick Then
       ' If vValue <> vbNullString And mbooDoubleBook = False Then
          '  If udtAppointment.ProviderID <> 0 Then piProviderID = udtAppointment.ProviderID
       ' Else
            If udtAppointment.ProviderID <> 0 Then piProviderID = udtAppointment.ProviderID
            frmSchedPatient.mbLookupPatient = True
            Screen.MousePointer = vbDefault
           ' Me.Hide
            frmSchedPatient.Show vbModal
            udtSchedPatient.IntId = vbNullString
            Screen.MousePointer = vbHourglass
            Get_Mode
            mBeenHere = True
       ' End If
    Else
        If udtAppointment.ProviderID <> 0 Then piProviderID = udtAppointment.ProviderID
        frmSchedPatient.mbLookupPatient = False
        udtSchedPatient.IntId = vbNullString
         frmSchedPatient.Show vbModal
        Screen.MousePointer = vbHourglass
        udtSchedPatient.IntId = vbNullString
        Get_Mode
        mBeenHere = True
    End If
    Screen.MousePointer = vbDefault
    schTimer.Enabled = True
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_SchedPatient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_SchedPatient", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Validate_Appt_Time(vlngPatient_ID As Long)
Dim strSQL As String
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim strStartTime As String
Dim strStart As String
Dim strEnd As String
Dim strDuration As String

    On Error GoTo Error_Validate_Appt_Time
    Screen.MousePointer = vbHourglass
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    strSQL = "Validate_Appt_Time " & vlngPatient_ID & ", " & udtAppointment.ProviderID & ", " & udtAppointment.FacilityID & ", '" & udtAppointment.SchedDate & "'"
    Set rsPat = cnPat.Execute(strSQL)
    With rsPat
    If Not .EOF Then
        strStartTime = (!Appt_Time)
        strDuration = (!Duration)
    End If
    End With
    strDuration = Replace(strDuration, "minutes", "")
    strDuration = Trim(strDuration)
    If strDuration <> vbNullString Then
        strEnd = DateAdd("n", Val(strDuration), strStartTime)
        strEnd = Format(strEnd, "hh:mm AM/PM")
        strStart = Format(strStartTime, "hh:mm AM/PM")
    End If
    If Format(udtAppointment.SchedTime, "hh:mm AM/PM") >= Format(strStart, "hh:mm AM/PM") And _
        Format(udtAppointment.SchedTime, "hh:mm AM/PM") < Format(strEnd, "hh:mm AM/PM") Then
        udtAppointment.SchedTime = Format(strStart, "hh:mm AM/PM")
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_Appt_Time:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_Appt_Time", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Read_Colors(vstrColorName As String, vlngColor As Long)
Dim cnColor As New ADODB.Connection
Dim rsColor As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim lngPrimKey As Long

    Screen.MousePointer = vbHourglass
    On Error GoTo Error_Read_Colors
    If cnColor.State = adStateClosed Then cnColor.open psConnect(1)
    strSQL = "Get_Appt_Color '" & vstrColorName & "', " & vlngColor
    Set rsColor = cnColor.Execute(strSQL)
    With rsColor
        Do Until .EOF
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) Then
                        Select Case (!Appt_Type)
                            Case "Office"
                                udtColor.Office = (!APPT_COLOR)
                            Case "Hospital"
                                udtColor.Hospital = (!APPT_COLOR)
                            Case "Vacation"
                                udtColor.Vacation = (!APPT_COLOR)
                            Case "Cancelled"
                                udtColor.Cancelled = (!APPT_COLOR)
                            Case "CheckOut"
                                udtColor.CheckOut = (!APPT_COLOR)
                            Case "NoShow"
                                udtColor.NoShow = (!APPT_COLOR)
                            Case "Present"
                                udtColor.Present = (!APPT_COLOR)
                            Case "Lunch"
                                udtColor.Lunch = (!APPT_COLOR)
                            Case "PostOp"
                                udtColor.PostOp = (!APPT_COLOR)
                            Case "PreOp"
                                udtColor.PreOp = (!APPT_COLOR)
                            Case "NewPatient"
                                udtColor.NewPatient = (!APPT_COLOR)
                            Case "ExistPatient"
                                udtColor.ExistPatient = (!APPT_COLOR)
                            Case "Surgery"
                                udtColor.Surgery = (!APPT_COLOR)
                            Case "FollowUp"
                                udtColor.FollowUp = (!APPT_COLOR)
                        End Select
                    End If
                End If
            Next
            .MoveNext
        Loop
    End With
    If cnColor.State = adStateOpen Then
        cnColor.Close
        Set cnColor = Nothing
        Set rsColor = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Colors:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Colors", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Set_Status()

    udtSchedPatient.Reason = Replace(udtSchedPatient.Reason, "Reason:", "")
    udtSchedPatient.Reason = Trim(udtSchedPatient.Reason)
'    If udtSchedPatient.Reason = "FollowUp: Cosmetic" Then udtSchedPatient.Status = "Post Operative"
'    If udtSchedPatient.Reason = "Appt: New Patient Cosmetic" Then udtSchedPatient.Status = "Appointment"
'    If udtSchedPatient.Reason = "PostOP" Then udtSchedPatient.Status = "Consultation"
'    If udtSchedPatient.Reason = 4 Then udtSchedPatient.Status = "Surgery Scheduled"
'    If udtSchedPatient.Reason = 5 Then udtSchedPatient.Status = "Surgery Cancelled"
'    If udtSchedPatient.Reason = 6 Then udtSchedPatient.Status = "Surgery Completed"
'    If udtSchedPatient.Reason = 10 Then udtSchedPatient.Status = "OnGoing Services"
'    If udtSchedPatient.Reason = 11 Then udtSchedPatient.Status = "Patient Exit"
'    If udtSchedPatient.Reason = 12 Then udtSchedPatient.Status = "Revision Pending"
'    If udtSchedPatient.Reason = 13 Then udtSchedPatient.Status = "Revision Scheduled"
'    If udtSchedPatient.Reason = 14 Then udtSchedPatient.Status = "Revision Completed"
'    If udtSchedPatient.Reason = 15 Then udtSchedPatient.Status = "Plan Cancelled"
                
End Function

Private Function Check_Change(vstrType As String)
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim strNow As String
Dim lngDateDiff As Long
Dim strInsert As String

    mbooSchChange = False
    mbooTeleChange = False
    strNow = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If mstrCheckTime <> vbNullString Then
        If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
                If vstrType = "Appointment" Then
                    strSQL = "Check_Appt_Change '" & Format(mstrCheckTime, "mm/dd/yyyy") & "'"
                    Set rsAppt = cnAppt.Execute(strSQL)
                    Do Until rsAppt.EOF
                        If rsAppt.Fields("INSERT_DATE") <> vbNullString Then
                            If Not IsNull(rsAppt.Fields("INSERT_DATE")) Then strInsert = rsAppt.Fields("INSERT_DATE")
                             lngDateDiff = DateDiff("d", mstrCheckTime, strInsert)
                             If lngDateDiff = 0 Then
                                lngDateDiff = DateDiff("h", Format(mstrCheckTime, "hh:mm AM/PM"), Format(strInsert, "hh:mm AM/PM"))
                                If lngDateDiff >= 0 Then
                                     mbooSchChange = True
                                     Exit Do
                                End If
                            End If
                        End If
                        rsAppt.MoveNext
                    Loop
                ElseIf vstrType = "Telephone" Then
                    strSQL = "Check_Message_Change '" & mstrCheckTime & "'"
                    Set rsAppt = cnAppt.Execute(strSQL)
                    Do Until rsAppt.EOF
                        If rsAppt.Fields("INSERT_DATE") <> vbNullString Then
                            lngDateDiff = DateDiff("n", mstrCheckTime, strNow)
                            If lngDateDiff > 20 Then
                                 mbooTeleChange = True
                                 Exit Do
                            End If
                        End If
                        rsAppt.MoveNext
                    Loop
                End If
            End If
    mstrCheckTime = strNow
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Exit Function
End Function

Private Function Check_Duplicates(vstrText As String)
Dim i%
Dim Arr
Dim strRTF As String

    mbooDup = False
    strRTF = Trim(rtfChanges.Text)
    If Trim(rtfChanges.Text) <> vbNullString Then
        Arr = Split(strRTF, vbNewLine)
        If UBound(Arr) > 0 Then
            For i = 0 To UBound(Arr)
                mbooDup = False
                If TextFound(vstrText, Arr(i), Len(Arr(i)), False) Then
                    mbooDup = True
                    Exit For
                End If
            Next i
        ElseIf TextFound(vstrText, strRTF, Len(vstrText), False) Then
            mbooDup = True
        End If
    End If
    Exit Function
End Function
Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)

    With vObjGrid
        vlngRow = vlngRow
        If .MaxRows < vlngRow Then .MaxRows = vlngRow
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = vlngRow
            .Row2 = vlngRow
            .BackColor = 12632256
            .ClearRange 1, vlngRow, .MaxCols, vlngRow, False
            .RowHeight(vlngRow) = 3.5
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub
Private Function Add_SchedRow(vlngCol As Long)
Dim lngCol As Long
Dim lngRow As Long


    On Error GoTo Error_Add_Row
    
    lngCol = vlngCol
    If TextFound("MTG", udtSchedPatient.Collection, Len("MTG"), True) Then
        udtAppointment.SchedType = "Vendor"
    End If
    If Trim(udtAppointment.Add_Proc_DESC) <> vbNullString Then
        udtSchedPatient.Collection = udtSchedPatient.Collection & " " & vbNewLine & udtAppointment.Add_Proc_DESC
    End If
    Select Case IsProperCase(udtAppointment.VisitType)
        Case "Init Psych", "New Patient Office Visit", "Initial Clinical Office Visit", "Initial Office Visit"
            clsGridFormat.BackColor = grdYellow
    End Select
    If TextFound("Psychiatric Diagnostic", udtAppointment.VisitType, Len("Psychiatric Diagnostic"), True) Then
        clsGridFormat.BackColor = grdYellow
    End If
    Select Case IsProperCase(udtAppointment.SchedType)
        Case "Physician"
            clsGridFormat.BackColor = GridNonPatient
        Case "Vendor"
            clsGridFormat.BackColor = GridNonPatient
    End Select
    Select Case IsProperCase(udtSchedPatient.Status)
        Case "Present"
            clsGridFormat.BackColor = grdCurItem
        Case "No Show", "Cancelled"
            clsGridFormat.BackColor = grdGray
        Case "Check Out"
            clsGridFormat.BackColor = DataColor
        Case "Confirmed"
            clsGridFormat.BackColor = GridConfir
    End Select
    If clsShedGrid.ApptDate <> vbNullString Then
        If Is_Closed_Encounter Then
            clsGridFormat.BackColor = grdGreen
        End If
    End If
    If lngCol = clsShedGrid.ProvCol2 Then
        For lngRow = clsShedGrid.Row To (clsShedGrid.Row + (clsShedGrid.SpanRow - 1))
            If mstrGridType = "Name Only" Then
                If lngRow = clsShedGrid.Row Then
                    Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, CStr(clsShedGrid.SpanRow), udtSchedPatient.Name & " - " & CStr(udtSchedPatient.IntId), "TypeHAlignCenter", CStr(udtSchedPatient.IntId), clsGridFormat.BackColor, udtAppointment.SchedType & "|" & CStr(udtAppointment.Appt_ID)
                    grdData.SetCellBorder lngCol, clsShedGrid.Row, lngCol, clsShedGrid.Row + clsShedGrid.SpanRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Else
                    Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, "", "", "TypeHAlignCenter", CStr(udtSchedPatient.IntId), clsGridFormat.BackColor, udtAppointment.SchedType & "|" & CStr(udtAppointment.Appt_ID)
                End If
            Else
                If lngRow = clsShedGrid.Row Then
                    Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, CStr(clsShedGrid.SpanRow), udtSchedPatient.Collection, "TypeHAlignCenter", CStr(udtSchedPatient.IntId), clsGridFormat.BackColor, udtAppointment.SchedType & "|" & CStr(udtAppointment.Appt_ID)
                    grdData.SetCellBorder lngCol, clsShedGrid.Row, lngCol, clsShedGrid.Row + clsShedGrid.SpanRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Else
                    Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, "", "", "TypeHAlignCenter", CStr(udtSchedPatient.IntId), clsGridFormat.BackColor, udtAppointment.SchedType & "|" & CStr(udtAppointment.Appt_ID)
                End If
            End If
        Next lngRow
    Else
        
        If clsShedGrid.SpanRow > 0 Then
            For lngRow = clsShedGrid.Row To (clsShedGrid.Row + (clsShedGrid.SpanRow - 1))
                If lngRow = clsShedGrid.Row Then
                    Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, CStr(clsShedGrid.SpanRow), udtSchedPatient.Collection, "TypeHAlignCenter", CStr(udtSchedPatient.IntId), clsGridFormat.BackColor, udtAppointment.SchedType & "|" & CStr(udtAppointment.Appt_ID)
                    grdData.SetCellBorder lngCol, clsShedGrid.Row, clsShedGrid.ProvCol2, clsShedGrid.Row + clsShedGrid.SpanRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Else
                    Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, "", "", "TypeHAlignCenter", CStr(udtSchedPatient.IntId), clsGridFormat.BackColor, udtAppointment.SchedType & "|" & CStr(udtAppointment.Appt_ID)
                    grdData.SetCellBorder lngCol, clsShedGrid.Row, clsShedGrid.ProvCol2, clsShedGrid.Row + clsShedGrid.SpanRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
            Next lngRow
        Else
            clsShedGrid.SpanRow = 1
            For lngRow = clsShedGrid.Row To (clsShedGrid.Row + (clsShedGrid.SpanRow - 1))
                Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, CStr(clsShedGrid.SpanRow), udtSchedPatient.Collection, "TypeHAlignCenter", CStr(udtSchedPatient.IntId), clsGridFormat.BackColor, udtAppointment.SchedType & "|" & CStr(udtAppointment.Appt_ID)
                grdData.SetCellBorder lngCol, clsShedGrid.Row, clsShedGrid.ProvCol2, clsShedGrid.Row + clsShedGrid.SpanRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Next lngRow
        End If
    End If
    If Trim(udtSchedPatient.Status) <> vbNullString Then
        Manage_Colors udtSchedPatient.Status, Val(udtSchedPatient.IntId)
    End If
    clsGridFormat.BackColor = 0
    grdData.Refresh
   Exit Function
Error_Add_Row:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Add_Row", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Add_WeekColumn(vlngCol As Long)
Dim lngCol As Long
Dim lngRow As Long
Dim strDate As String
Dim varCellValue
Dim lngInsertCol As Long
Dim Arr

    On Error GoTo Error_Add_WeekColumn
    lngRow = SpreadHeader + 2
    lngCol = vlngCol
    With grdData
        .GetText lngCol, SpreadHeader + 2, varCellValue
        Arr = Split(varCellValue, "-")
        If UBound(Arr) > 0 Then
            strDate = Arr(1)
        End If
        .MaxCols = .MaxCols + 1
        lngInsertCol = lngCol + 1
        .InsertCols lngInsertCol, 1
        lngCol = lngInsertCol
        .ColWidth(lngCol) = 20 * mlngSystemFileSize
        Format_CellTypeFixed grdData, lngRow, lngCol, "1", GetDay(strDate) & " - " & strDate, "TypeHAlignCenter", "OVERBOOKED", HeaderColor
        grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdData
        grdData.RemoveCellSpan 1, SpreadHeader + 4
        Format_CellTypeFixed grdData, SpreadHeader + 4, lngCol, "1", strDate & " " & "Overbook", "TypeHAlignCenter", CStr(udtAppointment.ProviderID), HeaderColor
        grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        For lngRow = 1 To .MaxRows
            Format_CellTypeEdit grdData, lngRow, lngCol, "1", " ", "TypeHAlignCenter"
            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Next lngRow
    End With
    Exit Function
Error_Add_WeekColumn:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", " Add_WeekColumn ", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub Resize_Grid()

    Me.Top = frmMDI.Top + 1450
    Me.Left = frmMDI.Left + 120
    Me.Height = 10455
    Me.Width = 15255
    If mstrApptMode = "Day" Then
        grdData.ColWidth(1) = 78.7 * mlngSystemFileSize
        grdChanges.Visible = True
    Else
        grdData.ZOrder 0
    End If
    Exit Sub
End Sub
Private Function Load_Enctr_Format()
    schTimer.Enabled = False
    Select Case mstrPractice
        Case "Gyn", "IVF"
            If IsFormLoaded("frmSuperbill") Then
                Unload frmSuperBill
                Exit Function
            ElseIf clsEncounter.Intrvn_CD = "99024" Then
                If clsEncounter.EncounterType <> vbNullString Then
                   If Not IsFormLoaded("frmInPatient") Then
                       ' Me.Hide
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.mstrLoadType = "PostOpVisit"
                        Load frmInPatient
                        frmInPatient.Show vbModal
                      '  Me.Show
                    Else
                      '  Me.Hide
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = piEncounterID
                        Load frmInPatient
                        frmInPatient.Show vbModal
                     '   Me.Show
                    End If
                End If
                Exit Function
           ElseIf Trim(clsOpNote.Procedure_Name) <> vbNullString And Trim(clsOpNote.Procedure_ID) <> vbNullString Then
                   If Not IsFormLoaded("frmInPatient") Then
                        Me.Hide
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.mstrLoadType = "Surgery"
                        Load frmInPatient
                        frmInPatient.Show vbModal
                     '   Me.Show
                    Else
                        Me.Hide
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        Load frmInPatient
                        frmInPatient.Show vbModal
                     '   Me.Show
                    End If
                Exit Function
            Else
                If Not IsFormLoaded("frmInPatient") Then
                        Me.Hide
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.Load_Reports
                        Load frmInPatient
                        frmInPatient.Show vbModal
                     '   Me.Show
                 Else
                        Me.Hide
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        Load frmInPatient
                        frmInPatient.Show vbModal
                     '   Me.Show
                 End If
                Exit Function
            End If
        Case "Psychiatry"
            If IsFormLoaded("frmSuperbill") Then
                Unload frmSuperBill
            Else
                If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
                If clsEncounter.EncounterType <> vbNullString Then
                   If Not IsFormLoaded("frmInPatient") Then
                        Me.Hide
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.mstrLoadType = "Appointment"
                        Load frmInPatient
                        frmInPatient.Show vbModal
                        Load Me
                        Me.Refresh
                        Me.Show
                    Else
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                        frmInPatient.mstrLoadType = "Appointment"
                        Load frmInPatient
                        frmInPatient.Show vbModal
                        Load Me
                        Me.Refresh
                        Me.Show
                    End If
                End If
                Exit Function
            End If
    Case "Plastic Surgery"
        If IsFormLoaded("frmSuperbill") Then
            Unload Me
            Exit Function
        ElseIf clsEncounter.Intrvn_CD = "99024" Then
 
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                        Me.Hide
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "PostOpVisit"
                    Load frmInPatient
                    frmInPatient.Show vbModal
                   ' Me.Show
                Else
                   ' Me.Hide
                    frmInPatient.mstrLoadType = "PostOpVisit"
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                    Load frmInPatient
                    frmInPatient.Show vbModal
                   ' Me.Show
                End If
            End If
            Unload Me
            Exit Function
        ElseIf clsEncounter.VisitType = "Ambulatory Procedure Visit" And TextFound("049X", clsEncounter.Intrvn_CD, Len("049X"), True) Then
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                   ' Me.Hide
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "Ambulatory Procedure Visit"
                    Load frmInPatient
                    frmInPatient.Show vbModal
                  '  Me.Show
                Else
                  '  Me.Hide
                        frmInPatient.piSelPatientID = 0
                        frmInPatient.piSelPatientID = piPatientID
                        frmInPatient.piSelEnctrID = 0
                        frmInPatient.piSelEnctrID = piEncounterID
                    Load frmInPatient
                    frmInPatient.Show vbModal
                 '   Me.Show
                End If
            End If
           ' Unload Me
            Exit Function
        ElseIf clsEncounter.VisitType = "Pre-Operative Exam Anesthesia" And clsEncounter.Intrvn_CD = "99201" Then
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "PreOperative Anesthesia"
                    Load frmInPatient
                    frmInPatient.Show vbModal
                Else
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    Load frmInPatient
                    frmInPatient.Show vbModal
                End If
            End If
          '  Unload Me
            Exit Function
        ElseIf clsEncounter.VisitType = "Surgical Admission" And clsEncounter.Intrvn_CD = "90806" Then
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "Surgical Admission"
                    Load frmInPatient
                    frmInPatient.Show vbModal
                Else
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    Load frmInPatient
                    frmInPatient.Show vbModal
                End If
            End If
            Exit Function
       ElseIf Trim(clsOpNote.Procedure_Name) <> vbNullString And Trim(clsOpNote.Procedure_ID) <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.mstrLoadType = "Surgery"
                    Load frmInPatient
                    frmInPatient.Show vbModal
                Else
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    Load frmInPatient
                    frmInPatient.Show vbModal
                End If
            Unload Me
            Exit Function
        Else
            'Create_Enctr_Default
            If clsEncounter.EncounterType = vbNullString Then Get_Encounter piEncounterID
            If clsEncounter.EncounterType <> vbNullString Then
               If Not IsFormLoaded("frmInPatient") Then
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    frmInPatient.Load_Reports
                    Load frmInPatient
                    frmInPatient.Show vbModal
                    frmInPatient.ZOrder 0
'                    frmInPatient.Show
                Else
                    frmInPatient.piSelPatientID = 0
                    frmInPatient.piSelPatientID = piPatientID
                    frmInPatient.piSelEnctrID = 0
                    frmInPatient.piSelEnctrID = piEncounterID
                    Load frmInPatient
                    frmInPatient.Show vbModal
                End If
            End If
            Exit Function
        End If
    End Select
    Exit Function
End Function
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
        .VisitType = vbNullString
        .Intrvn_CD = vbNullString
        .VisitTime = vbNullString
    End With
End Sub
Private Function Clear_Sched()
    With clsShedGrid
        .CellSpan = vbNullString
        .CellTag = vbNullString
        .CellNote = vbNullString
        .BackColor = 0
        .Text = vbNullString
        .TextNote = vbNullString
        .List = vbNullString
        .Duration = 0
        .InsertRow = 0
        .ProvCol = 0
        .ProvCol2 = 0
        .SpanRow = 0
        .ApptDate = vbNullString
        .StartRow = 0
        .EndRow = 0
        .Col = 0
        .Row = 0
        .RowName = vbNullString
        .ColHeader = vbNullString
        .CellType = vbNullString
    End With
End Function


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
Private Sub Format_Messages()
Dim strSQL As String
Dim cnMessage As New ADODB.Connection
Dim rsMessage As New ADODB.Recordset
Dim lngCol As Long
Dim strHeader As String
Dim lngCnt As String
Dim sDate As String

    On Error GoTo Error_Format_Messages
    Screen.MousePointer = vbHourglass
    mlngMessRow = 0
    Get_Providers
    grdMessages.MaxRows = mlngMessRow
    With grdMessages
        grdMessages.ClearRange 0, 0, .MaxCols, .MaxRows, False
        .Redraw = False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 10
        .FontName = "Times New Roman"
        .GridSolid = False
        .GridShowHoriz = False
        .GridShowVert = False
        .RowHeadersShow = False
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .RowHeight(0) = 20
    End With
    If cnMessage.State = adStateClosed Then cnMessage.open psConnect(1)
    strSQL = "Get_Grid_Headers 'Telephone Messages'"
    Set rsMessage = cnMessage.Execute(strSQL)
        lngCnt = 0
        With rsMessage
            Do Until .EOF
                lngCnt = lngCnt + 1
                .MoveNext
            Loop
        End With
    strSQL = "Get_Grid_Headers 'Telephone Messages'"
    Set rsMessage = cnMessage.Execute(strSQL)
    With rsMessage
        Do Until .EOF
            If rsMessage.Fields("HEADERS") <> vbNullString Then
                lngCol = lngCol + 1
                grdMessages.MaxCols = lngCol
                strHeader = rsMessage.Fields("HEADERS")
                strHeader = Replace(strHeader, "_", "  ")
                grdMessages.Col = lngCol
                grdMessages.ColID = rsMessage.Fields("HEADERS")
                grdMessages.SetColItemData (lngCol), (!FORM_ID)
                Format_CellTypeStatic grdMessages, SpreadHeader + 2, lngCol, "1", strHeader, "TypeHAlignCenter", "", "", HeaderColor
                grdMessages.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                If lngCol > 1 Then
                    grdMessages.ColWidth(lngCol) = 20 * mlngSystemFileSize
                Else
                    grdMessages.ColWidth(lngCol) = 8 * mlngSystemFileSize
                End If
                If (!Visible) = 0 Then grdMessages.ColHidden = True
            End If
        .MoveNext
        Loop
    End With
    Format_CellTypeStatic grdMessages, SpreadHeader, 1, "maxcols", "TELEPHONE MESSAGES FOR" & " " & clsProvider.Name, "TypeHAlignCenter", "", "", HeaderColor
    grdMessages.SetCellBorder 1, SpreadHeader, grdMessages.MaxCols, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    sDate = Format(Now, "dddd, MMMM DD, yyyy")
    Format_CellTypeStatic grdMessages, SpreadHeader + 1, 1, "maxcols", sDate, "TypeHAlignCenter", "", "", HeaderColor
    grdMessages.SetCellBorder 1, SpreadHeader + 1, grdMessages.MaxCols, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    grdMessages.Redraw = True
    If cnMessage.State = adStateOpen Then
        cnMessage.Close
        Set cnMessage = Nothing
        Set rsMessage = Nothing
    End If
    Exit Sub
    Screen.MousePointer = vbDefault
Error_Format_Messages:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_Messages", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub



Private Function Get_SurgeryFormat()
Dim Arr
Dim lngCol As Long
Dim lngRow As Long
Dim Idx As Integer
Dim vTime
Dim vdata
Dim varData
Dim varDoubleBook
Dim booNoInsert As Boolean
Dim lngProviderID As Long
Dim lngProvCol As Long
Dim varCellValue
Dim CellNote As Long
Dim lngSpanRow As Long
Dim lngRowTime As Long
Dim intSpan As Integer
Dim strInt As String
Dim cnApptFormat As New ADODB.Connection
Dim rsApptFormat As New ADODB.Recordset
Dim strSQL As String
Dim strMin As String
Dim lngCnt As Long
Dim lngProvID As Long
Dim strSpan As String

    On Error GoTo Error_Get_SurgeryFormat
    Screen.MousePointer = vbHourglass
    Exit Function
Error_Get_SurgeryFormat:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_SurgeryFormat", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

 
Private Function LookUpCellTag(vstrName As String, Optional vlngCol As Long) As String
Dim lngRow As Long
Dim lngCol As Long
Dim strTag As String
    
    lngCol = 0
    With grdData
        LookUpCellTag = vbNullString
        For lngRow = 1 To .MaxRows
            .Col = vlngCol
            If .CellTag <> vbNullString Then strTag = .CellTag
            If Trim(strTag) <> vbNullString And TextFound(vstrName, Trim(strTag), Len(vstrName), True) Then
                LookUpCellTag = lngCol & "|" & lngRow
                Exit For
            End If
        Next lngRow
    End With
    Exit Function
End Function

Public Sub Prepare_Grid()
Dim sValue As String
Dim Idx As Long
Dim Arr

    On Error GoTo Prepare_Grid
    Screen.MousePointer = vbHourglass
    mLoadPatient = False
    mGridClick = False
    mstrValidAppoint = vbNullString
    mbooByPass = False
    Clear_SchedPatient
    Clear_Appt
    mstrApptMode = "Day"
    Clear_SchedPatient
    Clear_Appt
    udtAppointment.SchedDate = Format(Now, "mm/dd/yyyy")
    mbooByPass = True
    MonthView1.Value = udtAppointment.SchedDate
    mbooByPass = False
    Create_Defaults
    If Get_SuperProv Then
        udtAppointment.ProviderID = piProviderID
    Else
        udtAppointment.ProviderID = piProviderID
    End If
    mLoadPatient = False
    mlngClickPrev = 0
    mlngClickNext = 0
    mGridClick = False
    mstrFirstDay = vbNullString
    udtColor.Office = picOffice.BackColor
    udtColor.Hospital = picHospital.BackColor
    udtColor.Vacation = picBlock.BackColor
    udtColor.Cancelled = picCheckOut.BackColor
    udtColor.CheckOut = picCheckOut.BackColor
    udtColor.NoShow = picCancel.BackColor
    udtColor.Present = picPresent.BackColor
    udtColor.Lunch = picBlock.BackColor
    udtColor.PostOp = picPostOp.BackColor
    udtColor.PreOp = picPreOp.BackColor
    udtColor.NewPatient = picNewPatient.BackColor
    udtColor.ExistPatient = picExstNewPat.BackColor
    udtColor.Surgery = picSurgery.BackColor
    udtColor.FollowUp = picReturn.BackColor
    Get_TimeSlot
    grdData.Redraw = True
    mFirstLoad = True
    mbooByPass = True
    mbooByPass = True
    If udtAppointment.SchedDate = vbNullString Then udtAppointment.SchedDate = Format(Now, "mm/dd/yyyy")
    MonthView1.Value = udtAppointment.SchedDate
    lblApptDate.Caption = "Appointment Date:  " & Format(udtAppointment.SchedDate, "MMMM DD, YYYY")
    Format_Grid
    Load_Facilities
    Get_Appt_Providers
    If mstrPractice = "Surgery" Then
        Get_SurgFacility
    Else
        Get_ReportFacility piFacilityID
    End If
    If piFacilityID = udtFacility.FacID And mstrPractice = "Surgery" Then
        mstrApptMode = "Surgery"
        Get_SurgeryFormat
    Else
        mstrApptMode = "Day"
        Get_Mode
    End If
    mbooByPass = False
    Format_Messages
    Read_Messages
    Screen.MousePointer = vbDefault
    Exit Sub
Prepare_Grid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Prepare_Grid", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Function Create_Defaults()
'Dim i%
'    Set_DefaultFacility
'     For i = 0 To cboLocation.listcount - 1
'         If cboLocation.ItemData(i) = piFacilityID Then
'             mbooByPass = True
'             cboLocation.Text = cboLocation.List(i)
'             Exit For
'         End If
'     Next i
'     Get_Appt_Providers
''     For i = 0 To cboProviders.listcount - 1
''        mbooByPass = True
''         If cboProviders.ItemData(i) = piProviderID Then
''            cboProviders.Text = cboProviders.List(i)
''            lblCurScheAppt.Visible = True
''            lblCurScheAppt.Caption = "Current Patient Appointments"
''            lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
''            Exit For
''         End If
''     Next i
'    mbooByPass = False
     Exit Function
End Function

Private Function Set_Headers()
Dim vDate
Dim bSpan As Boolean
Dim Span As Integer
Dim pvanchorrow As Variant
Dim pvanchorcol As Variant
Dim varCols As Variant
Dim varRows As Variant
Dim lngFacID As Long
Dim Idx As Integer
Dim lngSpan As Long

    udtAppointment.SchedDate = Format(MonthView1.Value, "MMMM DD, YYYY")
    Get_Providers udtAppointment.ProviderID
    grdData.ClearRange 1, SpreadHeader + 3, 2, SpreadHeader + 3, False
    grdData.ColHeaderRows = 4
    lngSpan = grdData.MaxCols
    If lngFacID > 0 Then Get_Facilities lngFacID
    grdData.RowHeight(SpreadHeader) = 20
    If grdData.MaxCols = 1 Then
        grdData.ColWidth(1) = 78.7 * mlngSystemFileSize
        Format_CellTypeEdit grdData, SpreadHeader, 1, CStr(lngSpan), cctFacilities.Text, "TypeHAlignCenter", "", HeaderColor, 14
        grdData.SetCellBorder 1, SpreadHeader, 1, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeStatic grdData, SpreadHeader + 1, 1, CStr(lngSpan), "Current Date" & " " & Format(udtAppointment.SchedDate, "dddd mmmm dd, yyyy"), "TypeHAlignCenter", "", "", HeaderColor
        Add_Spacer SpreadHeader + 2, grdData
        Format_CellTypeStatic grdData, SpreadHeader + 3, 1, CStr(1), UCase(clsProvider.Name), "TypeHAlignCenter", CStr(udtAppointment.ProviderID), "", HeaderColor
        grdData.SetCellBorder 1, SpreadHeader + 3, 1, SpreadHeader + 3, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    ElseIf grdData.MaxCols = 2 Then
        grdData.ColWidth(1) = 39.35 * mlngSystemFileSize
        grdData.ColWidth(2) = 39.35 * mlngSystemFileSize
        Format_CellTypeEdit grdData, SpreadHeader, 1, CStr(lngSpan), cctFacilities.Text, "TypeHAlignCenter", "", HeaderColor, 14
        grdData.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Format_CellTypeStatic grdData, SpreadHeader + 1, 1, CStr(lngSpan), "Current Date" & " " & Format(udtAppointment.SchedDate, "dddd mmmm dd, yyyy"), "TypeHAlignCenter", "", "", HeaderColor
        grdData.SetCellBorder 1, SpreadHeader + 1, 2, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 2, grdData
        Format_CellTypeStatic grdData, SpreadHeader + 3, 1, CStr(lngSpan), UCase(clsProvider.Name), "TypeHAlignCenter", CStr(udtAppointment.ProviderID), "", HeaderColor
        grdData.SetCellBorder 1, SpreadHeader + 3, 2, SpreadHeader + 3, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    End If
    
    Exit Function
End Function
Private Function Check_Patient(vlngPatientID As Long, vstrApptTime As String, vstrApptDate As String)
Dim strSQL As String
Dim sDate As String
Dim cnValidate As New ADODB.Connection
Dim rsValidate As New ADODB.Recordset
Dim lngApptID As Long
Dim lngPatID As Long
Dim strNow As String
Dim strTime As String

    On Error GoTo Error_Check_Patient
    If cnValidate.State = adStateClosed Then cnValidate.open psConnect(0)
    strNow = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    Check_Patient = False
    lngPatID = vlngPatientID
    sDate = Format(vstrApptDate, "mm/dd/yyyy")
    strTime = Format(vstrApptTime, "hh:mm AM/PM")
    Validate_Appt_Time lngPatID
    strSQL = "Appt_Validate_PT_Appt " & lngPatID & ", '" & Format(strTime, "hh:mm AM/PM") & "', '" & Format(sDate, "mm/dd/yyyy") & "'"
    Set rsValidate = cnValidate.Execute(strSQL)
    With rsValidate
        If Not .EOF Then
            If (!Patient_ID) = lngPatID Then
                Check_Patient = True
                lngApptID = (!Appt_ID)
            End If
        End If
    End With
    If Check_Patient Then
        strSQL = "Appt_Update_Status " & lngApptID & ", 'Present',  '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'"
        Set rsValidate = cnValidate.Execute(strSQL)
    End If
    If cnValidate.State = adStateOpen Then
        cnValidate.Close
        Set cnValidate = Nothing
        Set rsValidate = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Check_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Check_Patient", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Get_CurrentEnctr(Optional vstrDate As String)
Dim sDate As String
Dim strSQL As String
Dim cnStatus As New ADODB.Connection
Dim rsStatus As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngDateDiff As Long
Dim strEnctTime As String
Dim lngEnctrID As Long
Dim iAnswer As Integer
Dim strDate As String
Dim Idx As Integer
Dim lngRow As Long
Dim varCellValue
Dim lngPrimKey As Long

    On Error GoTo Error_Get_CurrentEnctr
    Screen.MousePointer = vbHourglass
    strDate = vstrDate
    Dump_EncounterStrData
    piEncounterID = 0
    Clear_SchedPatient
    Clear_Appt
    Clear_Provider
    piEncounterID = 0
    udtAppointment.SchedDate = strDate
    If Trim(udtAppointment.SchedDate) <> vbNullString Then
        sDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    Else
        sDate = Format(Now, "mm/dd/yyyy")
    End If
    udtTime.StartDate = Format(sDate, "mm/dd/yyyy 00:00:00 AM/PM")
    udtTime.EndDate = DateAdd("h", 23, udtTime.StartDate)
    udtTime.EndDate = Format(udtTime.EndDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    If cctProvider.Text <> vbNullString Then
        With grdProviders
            For lngRow = 1 To .MaxRows
                .GetText 1, lngRow, varCellValue
                If varCellValue = cctProvider Then
                    .BlockMode = True
                        .Col = 1
                        .Col2 = 2
                        .Row = lngRow
                        .Row2 = lngRow
                        .BackColor = grdCurItem
                    .BlockMode = False
                    udtAppointment.ProviderID = .GetRowItemData(lngRow)
                    Exit For
                End If
            Next lngRow
        End With
    End If
    If cctFacilities.Text <> vbNullString Then
        With grdFacility
            For lngRow = 1 To .MaxRows
                .GetText 1, lngRow, varCellValue
                If varCellValue = cctFacilities Then
                    .BlockMode = True
                        .Col = 1
                        .Col2 = 2
                        .Row = lngRow
                        .Row2 = lngRow
                        .BackColor = grdCurItem
                    .BlockMode = False
                    udtAppointment.FacilityID = .GetRowItemData(lngRow)
                    udtFacility.FacID = udtAppointment.FacilityID
                    Exit For
                End If
            Next lngRow
        End With
    End If
    If clsEDI_Client.Site_ID = "2777" Then
        If Get_SuperProv Then
            If udtAppointment.ProviderID = 0 Then
                udtAppointment.ProviderID = piProviderID
            End If
        Else
            If udtAppointment.ProviderID = 0 Then
                udtAppointment.ProviderID = mlngLogonUser
            End If
        End If
    Else
        If udtAppointment.ProviderID = 0 Then
            udtAppointment.ProviderID = mlngLogonUser
        End If
    End If
    piFacilityID = udtAppointment.FacilityID
    Get_Providers udtAppointment.ProviderID
    If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
    strSQL = "Appt_Get_Current_Enctr " & piPatientID & ", '" & Format(udtTime.StartDate, "mm/dd/yyyy") & "', '" _
        & Format(udtTime.EndDate, "mm/dd/yyyy") & "'," & udtAppointment.ProviderID & "," & udtAppointment.FacilityID
    Set rsStatus = cnStatus.Execute(strSQL)
    With rsStatus
        Do Until .EOF
            If Not IsNull((!Visit_Type)) Then
                If Not TextFound("Telephone Visit", (!Visit_Type), Len("Telephone Visit"), True) Then
                    lngEnctrID = (!Encounter_ID)
                    Exit Do
                End If
            Else
                lngEnctrID = (!Encounter_ID)
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    If lngEnctrID = 0 Then
        strSQL = "SELECT * FROM ENCOUNTER WHERE PATIENT_ID = " & piPatientID & " AND " _
            & "CONVERT(VARCHAR(10), ENCNTR_START_DATE, 101) = '" & Format(udtTime.StartDate, "mm/dd/yyyy") & "' AND " _
            & "USER_ID like '%Profiler' And PURGE = 0 And Provider_ID = " & udtAppointment.ProviderID
        Set rsStatus = cnStatus.Execute(strSQL)
        With rsStatus
            If Not .EOF Then
                lngEnctrID = (!Encounter_ID)
            End If
        End With
    End If
    If lngEnctrID > 0 Then
        Get_Encounter lngEnctrID
        If clsEncounter.ProviderID <> piProviderID Then
             MsgBox "There is an existing encounter for this Patient today with another Provider! If you want to add a new encounter" _
                & " " & "please use the encounter component!", vbInformation, "PatientTrac Encounter Management"
            mbooPatValid = False
            If clsEncounter.ServicePlaceCode = "9" Then clsEncounter.ServicePlaceCode = "09"
            modSpecific.piEncounterID = piEncounterID
            Exit Function
        End If
    End If
    clsCptTemp.CPT_Code = vbNullString
    strSQL = "Get_PT_Appt_ProviderID  " & piPatientID & ", " & udtAppointment.ProviderID & "," & udtAppointment.FacilityID & ",'" _
        & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "'"
    Set rsStatus = cnStatus.Execute(strSQL)
    With rsStatus
        If Not .EOF Then
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PATIENT_TYPE"
                                udtSchedPatient.PatientType = Trim(Fld.Value)
                            Case "PURPOSE_VISIT"
                                udtAppointment.VisitType = Trim(Fld.Value)
                            Case "VISIT_CODE"
                                clsCptTemp.CPT_Code = Trim(Fld.Value)
                            Case "FACILITY_ID"
                                clsEncounter.FacilityID = Fld.Value
                                clsEncounter.FacID = Fld.Value
                        End Select
                    End If
                End If
            Next
        End If
    End With
    If Trim(udtSchedPatient.PatientType) <> vbNullString Then
        clsEncounter.EncounterType = udtSchedPatient.PatientType
    End If
    If PsUserID = vbNullString Then
        PsUserID = Get_PsUserID(psConnect(0))
    End If
    If clsCptTemp.CPT_Code = vbNullString Then
     Correlate_ApptVisit udtAppointment.VisitType
    End If
    If clsCptTemp.CPT_Code <> vbNullString Then
        clsEncounter.Intrvn_CD = clsCptTemp.CPT_Code
    End If
    Get_Facilities udtAppointment.FacilityID
    If clsFacility.FacType <> vbNullString Then
        clsEncounter.ServicePlace = clsFacility.FacType
    End If
    If clsFacility.pos <> 0 Then
        clsEncounter.ServicePlaceCode = clsFacility.pos
    End If
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If clsEncounter.ServicePlace = vbNullString Then clsEncounter.ServicePlace = "Office"
    If Val(clsEncounter.ServicePlaceCode) = 0 Then clsEncounter.ServicePlaceCode = "11"
    If clsEncounter.VisitType = vbNullString Then
        clsEncounter.VisitType = LookUpCPTName(clsEncounter.Intrvn_CD)
    End If
    If clsEncounter.StartDate = vbNullString Then clsEncounter.StartDate = udtAppointment.SchedDate
    If DateDiff("d", "01/01/2013", clsEncounter.StartDate) >= 1 Then
        If Trim(clsEncounter.Intrvn_CD) = "90801" Then
            Get_Providers clsEncounter.ProviderID
            If UCase(Replace(clsProvider.Title, ".", "")) = "MD" Or _
                UCase(Replace(clsProvider.Title, ".", "")) = "DO" Or _
                UCase(Replace(clsProvider.Title, ".", "")) = "ARNP" Or _
                UCase(Replace(clsProvider.Title, ".", "")) = "CRNP" Or _
                UCase(Replace(clsProvider.Title, ".", "")) = "PA" Then
                    clsEncounter.Intrvn_CD = "90792"
            Else
                 clsEncounter.Intrvn_CD = "90791"
            End If
            clsEncounter.VisitType = LookUpCPTName(clsEncounter.Intrvn_CD)
            If lngEnctrID > 0 Then
                clsEncounter.EncounterID = lngEnctrID
                strSQL = "UPDATE ENCOUNTER SET ENCOUNT_CPT_INTRVN = '" & clsEncounter.Intrvn_CD & "'," _
                    & "VISIT_TYPE = '" & clsEncounter.VisitType & "' WHERE ENCOUNTER_ID = " & clsEncounter.EncounterID
                Set rsStatus = cnStatus.Execute(strSQL)
            End If
        End If
    End If
    If clsEncounter.Intrvn_CD <> vbNullString And lngEnctrID > 0 Then
        strSQL = "Update_Encounter_CPT_Only '" & clsEncounter.Intrvn_CD & "'," & lngEnctrID
        Set rsStatus = cnStatus.Execute(strSQL)
    End If
    If lngEnctrID = 0 Then
        strSQL = "Create_Encounter " & piPatientID & "," & udtAppointment.ProviderID & "," & udtAppointment.FacilityID & ",'" _
             & clsEncounter.VisitType & "','" _
             & clsEncounter.ServicePlace & "','" _
             & clsEncounter.EncounterType & "','" _
             & clsEncounter.Intrvn_CD & "'," & udtAppointment.ProviderID & ", '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "','" _
             & udtAppointment.VisitType & "','" & Format(sDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "'," & clsEncounter.ServicePlaceCode & ",'" & PsUserID & "',0 "
            Set rsStatus = cnStatus.Execute(strSQL)
            If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
        strSQL = "Get_Encntr_ID_Date_PatientID " & piPatientID & " ,   '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "', " & udtAppointment.ProviderID
            Set rsStatus = cnStatus.Execute(strSQL)
            If Not rsStatus.EOF Then
                 piEncounterID = rsStatus.Fields("ENCOUNTER_ID")
                 modSpecific.piEncounterID = piEncounterID
            End If
        strSQL = "Get_Encounter_Service_ID_Enctr_ID_01  " & piEncounterID
        Set rsStatus = cnStatus.Execute(strSQL)
        With rsStatus
            If Not .EOF Then
                lngPrimKey = (!SERVICE_LINE_ID)
            End If
        End With
        Get_CPTFee Left(clsEncounter.Intrvn_CD, 5)
        If lngPrimKey = 0 Then
            strSQL = "Write_Encounter_Service_EMR_Enctr " & piEncounterID & ",'" _
               & clsEncounter.Intrvn_CD & "','" & clsCptTemp.CPT_Fee & "', '1','01','1','" & clsEncounter.VisitType & "'"
            Set rsStatus = cnStatus.Execute(strSQL)
        ElseIf clsEncounter.Intrvn_CD <> vbNullString Then
            strSQL = "Update_Encounter_Service_Proc_Code_Routine '" & clsEncounter.Intrvn_CD & "','" _
                    & clsEncounter.VisitType & "','" & clsCptTemp.CPT_Fee & "'," & lngPrimKey
            Set rsStatus = cnStatus.Execute(strSQL)
        End If
'        If clsEncounter.ServicePlaceCode <> "11" And Not IsNull(ccdAdmit.Value) And piEncounterID > 0 Then
'            Manage_Admission piEncounterID
'            If clsInPatRecord.AdmissionDate = vbNullString And IsNull(ccdAdmit.Value) Then
'                MsgBox "Please enter Admission Date", vbInformation, "PatientTrac Encounter"
'                Screen.MousePointer = vbDefault
'                Exit Function
'            End If
'        End If
        Collect_Add_Proc_Appointments
    Else
        piEncounterID = lngEnctrID
    End If
    Get_Encounter piEncounterID
    If clsEncounter.ServicePlaceCode = "9" Then clsEncounter.ServicePlaceCode = "09"
    modSpecific.piEncounterID = piEncounterID
    If cnStatus.State = adStateOpen Then
        cnStatus.Close
        Set cnStatus = Nothing
        Set rsStatus = Nothing
        Set Flds = Nothing
    End If
    mbooByPass = False
    Create_Folder piPatientID
    mbooPatValid = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_CurrentEnctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_CurrentEnctr", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Public Function Check_TimeBlocks(vlngCol As Long, vlngRow As Long, Optional vstrBook As String)
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim varCellValue
Dim vDate
Dim iAnswer As Integer
Dim anchorcol  As Variant
Dim anchorrow As Variant
Dim colspan As Variant
Dim rowspan As Variant
Dim IsSpan As Boolean
Dim Span As Integer
Dim strTime As String
Dim strData As String
Dim nDays As Long
Dim Arr
Dim lngDateDiff As Long
Dim strRange As String
Dim booExit As Boolean

    On Error GoTo Error_Check_TimeBlocks
    Screen.MousePointer = vbHourglass
    mbooTimeAvailable = True
    
    grdData.Redraw = False
    grdData.GetText vlngCol, vlngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then strData = varCellValue
    grdData.GetText 0, vlngRow, varCellValue
    If Trim(varCellValue) <> vbNullString Then strTime = varCellValue
    strRange = DateAdd("d", -7, udtAppointment.SchedDate)
    If Trim(strTime) = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
        strSQL = "Check_TimeBlocks_Provider_Facility_Range " & udtAppointment.ProviderID & ", " & udtFacility.FacID & ", '" & strRange & "' "
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
                                udtBlock.StartTime = Trim(Fld.Value)
                            Case "END_TIME"
                                udtBlock.EndTime = Trim(Fld.Value)
                            Case "REASON"
                                udtBlock.Reason = Trim(Fld.Value)
                            Case "DURATION"
                                udtBlock.Duration = Trim(Fld.Value)
                            Case "FACILITY_ID"
                                udtBlock.FacilityID = Trim(Fld.Value)
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
            If udtBlock.Duration = "Permanent" Then
                If DateDiff("n", udtAppointment.SchedTime, udtBlock.StartTime) > 0 Then
                    If DateDiff("n", udtAppointment.SchedTime, udtBlock.EndTime) < 0 Then
                        mbooTimeAvailable = False
                        booExit = True
                        Exit Do
                    Else
                        mbooTimeAvailable = True
                        Exit Do
                    End If
                Else
                    mbooTimeAvailable = True
                    Exit Do
                End If
                mbooTimeAvailable = False
            ElseIf nDays > 1 Then
                If DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate) > 0 And _
                    DateDiff("d", udtBlock.EndDate, udtAppointment.SchedDate) < 0 Then
                    mbooTimeAvailable = True
                    If udtBlock.EndDate <> vbNullString Then
                        lngDateDiff = DateDiff("d", Format(udtBlock.StartDate, "mm/dd/yyyy"), Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
                        If lngDateDiff >= 0 Then
                            lngDateDiff = DateDiff("n", udtBlock.StartTime, strTime)
                            If lngDateDiff >= 0 Then
                                lngDateDiff = DateDiff("n", strTime, udtBlock.EndTime)
                                If lngDateDiff >= 0 Then
                                    lngDateDiff = DateDiff("d", Format(udtBlock.EndDate, "mm/dd/yyyy"), Format(udtAppointment.SchedDate, "mm/dd/yyyy"))
                                    If lngDateDiff <= 0 Then
                                        mbooTimeAvailable = False
                                        Exit Do
                                    Else
                                        mbooTimeAvailable = True
                                        Exit Do
                                    End If
                                End If
                            End If
                        End If
                    End If
                ElseIf DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate) = 0 And _
                    DateDiff("d", udtBlock.EndDate, udtAppointment.SchedDate) = 0 Then
                    If lngDateDiff = DateDiff("n", udtAppointment.SchedTime, udtBlock.EndTime) < 0 Then
                        mbooTimeAvailable = False
                        booExit = True
                        Exit Do
                    Else
                        mbooTimeAvailable = True
                        Exit Do
                    End If
                End If
            End If
        Loop
    End If
    If clsApptPref.OverBook Then
        mbooTimeAvailable = True
    Else
        If Not clsApptPref.OverBook Then
            Span = grdData.GetCellSpan(vlngCol, vlngRow, anchorcol, anchorrow, colspan, rowspan)
            If Span Then
                mbooTimeAvailable = False
                MsgBox "Cannot schedule an appointment for this time, violation of practice overbooking policy!", vbInformation, "PatientTrac Appointment Scheduling"
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
    End If
    If Not clsApptPref.OverBook And mbooTimeAvailable = False Then
        MsgBox "Unable to schedule an appointment time. The Provider's appointment time you selected has been blocked for" & " " & IsProperCase(udtBlock.Reason) & ". ", vbInformation, "PatientTrac Appointment Scheduling"
        mstrValidAppoint = "Invalid Appointment"
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        mstrValidAppoint = "Authorized"
        mbooTimeAvailable = True
    End If
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
        Set Flds = Nothing
    End If
    grdData.Redraw = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Check_TimeBlocks:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Check_TimeBlocks", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Collect_Add_Proc_Appointments()
Dim strSQL As String
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim sDate As String
Dim strAdd_Proc_Code As String
Dim strAdd_Proc_Desc As String


    On Error GoTo Error_Collect_Add_Proc_Appointments
    Screen.MousePointer = vbHourglass
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    strSQL = "SELECT ADD_PROC_CODE, ADD_PROC_DESC FROM APPOINTMENTS " _
        & "INNER JOIN ENCOUNTER ON APPOINTMENTS.PATIENT_ID = ENCOUNTER.PATIENT_ID  " _
        & "WHERE CONVERT(VARCHAR(10), APPOINTMENTS.APPT_DATE, 101)  = CONVERT(VARCHAR(10), ENCOUNTER.ENCNTR_START_DATE, 101) AND " _
        & "ADD_PROC_CODE IS NOT NULL AND ADD_PROC_DESC IS NOT NULL AND " _
        & "APPOINTMENTS.PROVIDER_ID = ENCOUNTER.PROVIDER_ID AND APPOINTMENTS.ACTIVE = 1 AND ENCOUNTER.ENCOUNTER_ID = " & piEncounterID
    Set rsAppt = cnAppt.Execute(strSQL)
    With rsAppt
        If Not .EOF Then
            strAdd_Proc_Code = vbNullString
            If (!ADD_PROC_CODE) <> vbNullString Then
                strAdd_Proc_Code = Trim((!ADD_PROC_CODE))
            End If
            If (!Add_Proc_DESC) <> vbNullString Then
                strAdd_Proc_Desc = Trim((!Add_Proc_DESC))
            End If
        End If
    End With
    If Trim(strAdd_Proc_Code) <> vbNullString Then
        Write_EM_Appointment strAdd_Proc_Code, strAdd_Proc_Desc
    End If
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Collect_Add_Proc_Appointments:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Collect_Add_Proc_Appointments", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function
Private Function Write_EM_Appointment(vstrCode As String, vstrDesc As String)
Dim cnEnctr As New ADODB.Connection
Dim rsEnctr As New ADODB.Recordset
Dim strSQL As String
Dim strDate As String
Dim lngPrimKey As Long
Dim lngCnt As Long
Dim bSkip As Boolean
Dim strModifier As String
Dim strUnit As String
Dim strCnt As String

    On Error GoTo Error_Write_EM_Appointment
    Screen.MousePointer = vbHourglass
    If cnEnctr.State = adStateClosed Then cnEnctr.open psConnect(0)
    strSQL = "SELECT * FROM ENCOUNTER_SERVICE WHERE EDI_BILL_ID = " & piEncounterID
    Set rsEnctr = cnEnctr.Execute(strSQL)
    With rsEnctr
        Do Until .EOF
            .MoveNext
             lngCnt = lngCnt + 1
        Loop
    End With
    If lngCnt > 1 Then
        strSQL = "SELECT *  FROM ENCOUNTER_SERVICE WHERE EDI_BILL_ID = " & piEncounterID & " AND  PROCEDURE_CODE = '" & vstrCode & "'"
        Set rsEnctr = cnEnctr.Execute(strSQL)
         With rsEnctr
            If Not .EOF Then
                bSkip = True
            End If
        End With
    End If
    If vstrCode = "96101" Or _
        vstrCode = "96116" Or _
        vstrCode = "96102" Or _
        vstrCode = "96103" Or _
        vstrCode = "96111" Then
        strModifier = "25"
    Else
        strModifier = "22"
    End If
    strCnt = Val(lngCnt) + 1
    If Not bSkip Then
        strUnit = 1
        Get_RowFees vstrCode, vstrDesc
        strSQL = "Write_Encounter_Service " & piEncounterID & ", '" & vstrCode & "','" _
            & strModifier & "','','','','" & Format(clsCptTemp.CPT_Fee, "####00.00") & "','" _
            & CStr(strUnit) & "', '" & strCnt & "','" & strCnt & "','" & vstrDesc & "','',''"
        Set rsEnctr = cnEnctr.Execute(strSQL)
    End If
    clsCptTemp.CPT_Code = vbNullString
    clsCptTemp.CPT_Fee = vbNullString
    clsCptTemp.CPT_Subgroup = vbNullString
    If cnEnctr.State = adStateOpen Then
        cnEnctr.Close
        Set cnEnctr = Nothing
        Set rsEnctr = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_EM_Appointment:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_EM_Appointment", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function
Private Function Get_RowFees(vstrCode As String, Optional vstrDesc As String)
Dim cnFee As New ADODB.Connection
Dim rsFee As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr
Dim strFee As String
Dim varCellValue
Dim lngDateDiff As Long
Dim strUpdate As String
Dim strEnctr As String
Dim strCPTCode As String
Dim strProvider As String
Dim strCompare As String
Dim i%
Dim strTitle As String
Dim strSearchTitle As String


    On Error GoTo Error_Get_RowFees
    clsCptTemp.CPT_Code = vbNullString
    clsCptTemp.CPT_Fee = vbNullString
    clsCptTemp.CPT_Subgroup = vbNullString
    If clsEncounter.StartDate <> vbNullString Then
        strEnctr = Format(clsEncounter.StartDate, "mm/dd/yyyy")
    Else
        strEnctr = Format(Now, "mm/dd/yyyy")
    End If
    Get_Providers clsEncounter.ProviderID
    If Trim(clsProvider.Title) <> vbNullString Then
        strSearchTitle = UCase(clsProvider.Title)
        strSearchTitle = Replace(strCompare, ".", "")
        strSearchTitle = Trim(strSearchTitle)
    Else
        If clsEncounter.ProviderID = 0 Then
            Get_DefaultProvider
            clsEncounter.ProviderID = piProviderID
            Get_Providers clsEncounter.ProviderID
        Else
            Get_Providers clsEncounter.ProviderID
        End If
        strSearchTitle = UCase(clsProvider.Title)
        strSearchTitle = Replace(strSearchTitle, ".", "")
        strSearchTitle = Trim(strSearchTitle)
    End If
    strSearchTitle = Get_Provider_Type(strSearchTitle)
    strCPTCode = vstrCode
    Arr = Split(strCPTCode, ":")
    If UBound(Arr) > 0 Then
        strCPTCode = Arr(0)
    End If
    If cnFee.State = adStateClosed Then cnFee.open psConnect(1)
    If strCPTCode <> vbNullString Then strCPTCode = Left(strCPTCode, 5)
    strSQL = "SELECT * FROM CPT_TEMPLATE WHERE CPT_CODE like '" & strCPTCode & "%' AND " _
        & "PLACE_SERVICE_CD = '" & clsEncounter.ServicePlaceCode & "'"
    Set rsFee = cnFee.Execute(strSQL)
        Do Until rsFee.EOF
            strProvider = vbNullString
            strUpdate = vbNullString
            If rsFee.Fields("PROVIDER") <> vbNullString Then
                strProvider = UCase(rsFee.Fields("PROVIDER"))
                strProvider = Replace(strProvider, ".", "")
                strProvider = Trim(strProvider)
                Set Flds = rsFee.Fields
                For Each Fld In Flds
                    If Not IsNull(Trim(Fld.Value)) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case strProvider
                                Case strSearchTitle
                                    If Fld.Name = "CPT_FEE" Then
                                        clsCptTemp.CPT_Fee = Trim(Fld.Value)
                                    ElseIf Fld.Name = "CPT_CODE" Then
                                        clsCptTemp.CPT_Code = Trim(Fld.Value)
                                    ElseIf Fld.Name = "UPDATE_DATE" Then
                                        If Trim(Fld.Value) <> "01/01/1900" Or Trim(Fld.Value) <> "1/1/1900" Then
                                            strUpdate = Format(Trim(Fld.Value), "mm/dd/yyyy")
                                        End If
                                    End If
                            End Select
                        End If
                    End If
                Next
            End If
            rsFee.MoveNext
            If clsCptTemp.CPT_Fee <> vbNullString And clsCptTemp.CPT_Code <> vbNullString Then
                Exit Do
            End If
        Loop
    If strUpdate <> vbNullString Then
        lngDateDiff = DateDiff("d", strUpdate, strEnctr)
        If lngDateDiff < 0 Then
            strSQL = "SELECT * FROM CPT_TEMPLATE WHERE CPT_CODE like '" & strCPTCode & "%' AND PLACE_SERVICE_CD = '" & clsEncounter.ServicePlaceCode & "'"
            Set rsFee = cnFee.Execute(strSQL)
            Do Until rsFee.EOF
                strProvider = vbNullString
                strUpdate = vbNullString
                If rsFee.Fields("PROVIDER") <> vbNullString Then
                    strProvider = UCase(rsFee.Fields("PROVIDER"))
                    strProvider = Replace(strProvider, ".", "")
                    strProvider = Trim(strProvider)
                    Set Flds = rsFee.Fields
                    For Each Fld In Flds
                        If Not IsNull(Trim(Fld.Value)) Then
                            If Trim(Fld.Value) <> vbNullString Then
                                Select Case strProvider
                                    Case strSearchTitle
                                        If Fld.Name = "PREV_CPT_FEE" Then
                                            clsCptTemp.CPT_Fee = Trim(Fld.Value)
                                        ElseIf Fld.Name = "CPT_CODE" Then
                                            clsCptTemp.CPT_Code = Trim(Fld.Value)
                                        ElseIf Fld.Name = "UPDATE_DATE" Then
                                            If Trim(Fld.Value) <> "01/01/1900" Or Trim(Fld.Value) <> "1/1/1900" Then
                                                strUpdate = Format(Trim(Fld.Value), "mm/dd/yyyy")
                                            End If
                                        End If
                                End Select
                            End If
                        End If
                    Next
                End If
                rsFee.MoveNext
                If clsCptTemp.CPT_Fee <> vbNullString And clsCptTemp.CPT_Code <> vbNullString Then
                    Exit Do
                End If
            Loop
        End If
    End If
    strUpdate = vbNullString
    If clsCptTemp.CPT_Fee = vbNullString Then
        strSQL = "SELECT * FROM INTRVN_TEMPLATE WHERE CPT_INTRVN_CODE like '" & strCPTCode & "%'"
        Set rsFee = cnFee.Execute(strSQL)
        Do Until rsFee.EOF
            strProvider = vbNullString
            strUpdate = vbNullString
            If rsFee.Fields("PROVIDER") <> vbNullString Then
                strProvider = UCase(rsFee.Fields("PROVIDER"))
                strProvider = Replace(strProvider, ".", "")
                strProvider = Trim(strProvider)
                Set Flds = rsFee.Fields
                For Each Fld In Flds
                    If Not IsNull(Trim(Fld.Value)) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case strProvider
                                Case strSearchTitle
                                    If Fld.Name = "CPT_FEE" Then
                                        clsCptTemp.CPT_Fee = Trim(Fld.Value)
                                    ElseIf Fld.Name = "CPT_INTRVN_CODE" Then
                                        clsCptTemp.CPT_Code = Trim(Fld.Value)
                                    ElseIf Fld.Name = "UPDATE_DATE" Then
                                        If Trim(Fld.Value) <> "01/01/1900" Or Trim(Fld.Value) <> "1/1/1900" Then
                                            strUpdate = Format(Trim(Fld.Value), "mm/dd/yyyy")
                                        End If
                                    End If
                            End Select
                        End If
                    End If
                Next
            End If
            rsFee.MoveNext
            If clsCptTemp.CPT_Fee <> vbNullString And clsCptTemp.CPT_Code <> vbNullString Then
                Exit Do
            End If
        Loop
    End If
    If strUpdate <> vbNullString Then
        lngDateDiff = DateDiff("d", strUpdate, strEnctr)
        If lngDateDiff < 0 Then
            strSQL = "SELECT * FROM INTRVN_TEMPLATE WHERE CPT_INTRVN_CODE like '" & strCPTCode & "%'"
            Set rsFee = cnFee.Execute(strSQL)
            Do Until rsFee.EOF
                strProvider = vbNullString
                strUpdate = vbNullString
                If rsFee.Fields("PROVIDER") <> vbNullString Then
                    strProvider = UCase(rsFee.Fields("PROVIDER"))
                    strProvider = Replace(strProvider, ".", "")
                    strProvider = Trim(strProvider)
                    Set Flds = rsFee.Fields
                    For Each Fld In Flds
                        If Not IsNull(Trim(Fld.Value)) Then
                            If Trim(Fld.Value) <> vbNullString Then
                                Select Case strProvider
                                    Case strSearchTitle
                                        If Fld.Name = "CPT_FEE" Then
                                            clsCptTemp.CPT_Fee = Trim(Fld.Value)
                                        ElseIf Fld.Name = "CPT_INTRVN_CODE" Then
                                            clsCptTemp.CPT_Code = Trim(Fld.Value)
                                        ElseIf Fld.Name = "UPDATE_DATE" Then
                                            If Trim(Fld.Value) <> "01/01/1900" Or Trim(Fld.Value) <> "1/1/1900" Then
                                                strUpdate = Format(Trim(Fld.Value), "mm/dd/yyyy")
                                            End If
                                        End If
                                End Select
                            End If
                        End If
                    Next
                End If
                rsFee.MoveNext
                If clsCptTemp.CPT_Fee <> vbNullString And clsCptTemp.CPT_Code <> vbNullString Then
                    Exit Do
                End If
            Loop
        End If
    End If
    If cnFee.State = adStateOpen Then
        Set rsFee = Nothing
        cnFee.Close
        Set cnFee = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_RowFees:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_RowFees", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function

Private Function IsCellSpan(vlngRow As Long, Optional vlngCol As Long, Optional vlngTarRow As Long) As Boolean
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
Dim strTag As String
Dim strCurTag As String
Dim varCellValue
Dim lngPatID As Long

    If Not clsApptPref.OverBook Then
        Exit Function
    End If
    grdData.getSelection 0, Col, Row, Col2, Row2
    grdData.Redraw = False
    If vlngRow > vlngTarRow And vlngTarRow > 0 Then
        grdData.Col = vlngCol
        grdData.Row = i
        If grdData.CellTag <> vbNullString Then
            IsCellSpan = True
        End If
    ElseIf vlngRow > 0 Then
        For i = vlngRow To vlngTarRow
            If vlngCol > 0 Then
                For j = vlngCol To grdData.MaxCols
                    Span = grdData.GetCellSpan(vlngCol, vlngRow, anchorcol, anchorrow, colspan, rowspan)
                    grdData.Col = vlngCol
                    grdData.Row = i
                    If grdData.CellTag <> vbNullString Then
                        IsCellSpan = True
                        Exit For
                    End If
                    If Span Then
                        If j = vlngCol Then anchorcol = "1"
                        If i = vlngRow Then
                            IsCellSpan = True
                        End If
                    End If
                Next j
            Else
                Span = grdData.GetCellSpan(1, vlngRow, anchorcol, anchorrow, colspan, rowspan)
                grdData.Redraw = False
                If Span Then
                     If i = vlngRow Then
                        IsCellSpan = True
                        Exit For
                    End If
                End If
            End If
        Next i
    End If
    grdData.Redraw = True
    Exit Function
End Function
Private Function DayMode_BlockTime(vstrBlockReason As String, vstrBlockTo As String, vstrBlockFrom As String, _
vstrBlockDuration As String, vstrStartDate As String, vlngProviderID As Long)
Dim i%
Dim lngProviderID As Long
Dim strSQL As String
Dim vTime
Dim vDate
Dim BlockColor As Long
Dim lngStartRow As Long
Dim lngEndRow As Long
Dim lngNoRows  As Long
Dim nDays As Double
Dim Arr
Dim lngDateDiff As Long
Dim lngBlockCol As Long
Dim lngBlockCol2 As Long
Dim lngCol As Long
Dim lngRow As Long
Dim varCellValue
Dim booCancel As Boolean
Dim booSet As Boolean
Dim strErrata As String
Dim lngNo As Double
Dim strDuration As String
Dim strNote As String
Dim strLine As String


    On Error GoTo Error_DayMode_BlockTime
    If vstrBlockFrom = vbNullString Then Exit Function
    Screen.MousePointer = vbHourglass
        Arr = Split(vstrBlockDuration, " ")
    If UBound(Arr) > 0 Then
        strDuration = Arr(1)
        lngNo = Arr(0)
    Else
        strDuration = vstrBlockDuration
    End If
    Select Case IsProperCase(strDuration)
        Case "Days", "Day"
            nDays = lngNo
        Case "Weeks", "Week"
            nDays = lngNo * 6
        Case "Month", "Months"
            nDays = lngNo * 29
        Case "Year", "Years"
            nDays = lngNo * 364
    End Select
    For lngCol = 1 To grdData.MaxCols
        grdData.Col = lngCol
        grdData.Row = SpreadHeader + 3
        If grdData.CellTag <> vbNullString Then
            If vlngProviderID = Val(grdData.CellTag) Then
                If lngBlockCol = 0 Then
                    lngBlockCol = lngCol
                Else
                    lngBlockCol2 = lngCol
                    Exit For
                End If
            End If
        End If
    Next lngCol
    If grdData.MaxCols = 2 And lngBlockCol2 < 2 Then
        lngBlockCol2 = 2
    End If
    vDate = vstrStartDate
    vDate = Format(vDate, "mm/dd/yyyy")
    lngDateDiff = DateDiff("d", udtAppointment.SchedDate, udtBlock.EndDate)
    If lngDateDiff >= 0 And (lngDateDiff <= nDays Or IsProperCase(vstrBlockDuration) = "Permanent") Then
        For lngCol = 1 To grdData.MaxCols
            If IsProperCase(vstrBlockDuration) = "Permanent" Then
                grdData.Col = lngCol
                grdData.Row = SpreadHeader + 3
                If grdData.CellTag <> vbNullString Then
                    If vlngProviderID = grdData.CellTag Then
                        lngBlockCol = lngCol
                        Exit For
                    End If
                End If
            Else
                lngDateDiff = DateDiff("d", udtBlock.EndDate, udtAppointment.SchedDate)
                If lngDateDiff <= 0 Then
                    If vlngProviderID = grdData.GetColItemData(lngCol) Then
                        lngBlockCol = lngCol
                        Exit For
                    End If
                Else
                    booCancel = True
                    Exit For
                End If
            End If
        Next lngCol
        If grdData.MaxRows = 0 Then grdData.MaxRows = 1
        If Not booCancel Then
            For lngRow = 1 To grdData.MaxRows
                grdData.GetText 0, lngRow, vTime
                If UCase(vTime) = UCase(vstrBlockFrom) Then
                    lngStartRow = lngRow
                End If
                vTime = vbNullString
                grdData.GetText 0, lngRow, vTime
                If UCase(vTime) = UCase(vstrBlockTo) Then
                    lngEndRow = lngRow
                End If
            Next lngRow
            If lngStartRow = 0 Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            If lngEndRow = 0 Then lngEndRow = grdData.MaxRows
            lngNoRows = lngEndRow - lngStartRow
            Select Case IsProperCase(vstrBlockReason)
                Case "Lunch"
                    BlockColor = grdBlock
                Case "Hospital"
                    BlockColor = grdBlock
                Case "Vacation"
                    BlockColor = grdBlock
                Case "Out Of The Office"
                    BlockColor = grdBlock
            End Select
            lngNoRows = lngNoRows
            If lngBlockCol = 0 Then
                lngBlockCol = 1
            End If
            If lngStartRow >= 1 And lngEndRow > 1 Then
                For lngRow = lngStartRow To lngEndRow
                    clsGridFormat.CellNote = vbNullString
                    grdData.GetText lngBlockCol, lngRow, varCellValue
                    If Trim(varCellValue) <> vbNullString And Not TextFound(vstrBlockReason, CStr(varCellValue), Len(vstrBlockReason), True) Then
                        If lngRow = lngStartRow Or lngRow > lngStartRow And lngRow < lngEndRow Then
                            grdData.Col = lngBlockCol
                            grdData.Row = lngRow
                            grdData.GetText lngBlockCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                clsGridFormat.CellNote = Trim(varCellValue)
                                Arr = Split(varCellValue, "INSURANCE AUTHORIZATION:")
                                If UBound(Arr) > 0 Then
                                   varCellValue = Trim(Arr(0))
                                End If
                                clsGridFormat.CellNote = Trim(varCellValue)
                                If UBound(Arr) > 0 Then
                                   clsGridFormat.CellNote = Arr(0)
                                End If
                                If udtAppointment.SchedType = vbNullString Then udtAppointment.SchedType = "Patient"
                                If Trim(clsGridFormat.CellNote) <> vbNullString Then
                                    If IsProperCase(Trim(clsGridFormat.CellNote)) = "Patient" Or _
                                        IsProperCase(Trim(clsGridFormat.CellNote)) = "Vendor" Or _
                                        IsProperCase(Trim(clsGridFormat.CellNote)) = "Physician" Or _
                                        IsProperCase(Trim(clsGridFormat.CellNote)) = "Legal" Then
                                        grdData.GetText 0, lngRow, varCellValue
                                        If Trim(varCellValue) <> vbNullString Then
                                            strErrata = Trim(varCellValue)
                                        End If
                                        If Trim(clsGridFormat.CellNote) <> vbNullString Then strNote = udtAppointment.SchedType & " " & "is scheduled in the block time at:" & " " & strErrata & vbNewLine & clsGridFormat.CellNote
                                        If Trim(rtfChanges.Text) = vbNullString Then
                                            rtfChanges.Text = Trim(rtfChanges.Text) & vbNewLine & Trim(strNote)
                                        ElseIf Not TextFound(Trim(strNote), Trim(rtfChanges.Text), Len(Trim(strNote)), True) Then
                                            rtfChanges.Text = Trim(rtfChanges.Text) & vbNewLine & strLine & vbNewLine & Trim(strNote)
                                        End If
                                    End If
                                End If
                            End If
                            Exit For
                        End If
                    ElseIf lngRow > lngStartRow Then
                        lngNoRows = lngNoRows - 1
                    End If
                Next lngRow
            End If
        End If
        If Not booSet Then
            If lngNoRows = 0 Then lngNoRows = 1
            lngNoRows = lngEndRow - lngStartRow
            If lngNoRows = 0 Then lngNoRows = 1
            If lngBlockCol2 = 0 Then
                Format_CellTypeEdit_RowSpan grdData, lngStartRow, lngBlockCol, CStr(lngNoRows), vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
                grdData.SetCellBorder lngBlockCol, lngStartRow, lngBlockCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdData.SetCellBorder lngBlockCol, lngEndRow, lngBlockCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Else
                Format_CellTypeEdit_RowSpan grdData, lngStartRow, lngBlockCol, CStr(lngNoRows), vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
                grdData.SetCellBorder lngBlockCol, lngStartRow, lngBlockCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdData.SetCellBorder lngBlockCol, lngEndRow, lngBlockCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Format_CellTypeEdit_RowSpan grdData, lngStartRow, lngBlockCol2, CStr(lngNoRows), vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
                grdData.SetCellBorder lngBlockCol2, lngStartRow, lngBlockCol2, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdData.SetCellBorder lngBlockCol2, lngEndRow, lngBlockCol2, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_DayMode_BlockTime:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "DayMode_BlockTime", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function DayWeek(vstrDayWeek As String) As String
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

    'If udtBlock.Duration = "Permanent" Then lngDateDiff = 1
    lngDateDiff = DateDiff("ww", udtAppointment.SchedDate, udtBlock.EndDate)
    If udtBlock.Duration = "Permanent" Then lngDateDiff = 1
    If lngDateDiff >= 0 Then
        strRepDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
        If lngDateDiff = 0 Then lngDateDiff = 1
        For Idx = 1 To lngDateDiff
            FirstDay = Weekday(strRepDate, vbMonday)
            nDays = Val(FirstDay)
            If strRepDate <> vbNullString Then
                Select Case udtBlock.DayWeek
                    Case "Monday"
                        FirstDay = Weekday(strRepDate, vbMonday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Tuesday"
                        FirstDay = Weekday(strRepDate, vbTuesday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Wednesday"
                        FirstDay = Weekday(strRepDate, vbWednesday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Thursday"
                        FirstDay = Weekday(strRepDate, vbThursday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Friday"
                        FirstDay = Weekday(strRepDate, vbFriday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Saturday"
                        FirstDay = Weekday(strRepDate, vbSaturday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                    Case "Sunday"
                        FirstDay = Weekday(strRepDate, vbSunday)
                        DayWeek = DateAdd("d", -(FirstDay - 1), strRepDate)
                        Exit For
                End Select
            End If
        Next
    End If
    Exit Function

End Function
Public Function Get_Mode()

    clsShedGrid.ProvCol2 = 1
    Get_SurgFacility
    If piFacilityID = udtFacility.FacID And mstrPractice = "Plastic Surgery" Then
        mstrApptMode = "Surgery"
    End If
    Select Case mstrApptMode
        Case "Day"
            grdData.MaxCols = 1
            If piProviderID = 0 Then piProviderID = udtAppointment.ProviderID
            Resize_Grid
            Format_Grid
            If mstrGridType = vbNullString Then
                Get_GridType
            End If
            If mstrGridType = "Name Only" Then
                ReadData_NameOnly
            Else
                ReadData
            End If
        Case "Weekly"
            Format_Week Format(MonthView1.Value, "mm/dd/yyyy")
            Read_Week_Appointments
        Case "Surgery"
            If udtFacility.FacID > 0 Then piFacilityID = udtFacility.FacID
            Get_SurgeryFormat
    End Select
    mbooByPass = True
    grdData.Redraw = True
    'mbooTimer = True
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Mode:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Mode", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Function Set_BlockTime(Optional vlngProviderID As Long)
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim i%
Dim lngDateDiff As Long
Dim strRepDate As String
Dim booSkip As Boolean
Dim strStartRange As String
Dim strEndRange As String
Dim strTestTime As String
Dim strTestDay As String
Dim strWeek As String

    On Error GoTo Error_Set_BlockTime
    Screen.MousePointer = vbHourglass
    
    If vlngProviderID > 0 Then
        udtAppointment.ProviderID = vlngProviderID
    ElseIf udtAppointment.ProviderID = 0 Then
        udtAppointment.ProviderID = piProviderID
    End If
    Set_Permanent_Blocks
    If udtAppointment.SchedDate = vbNullString Then Exit Function
    strTestTime = DateAdd("d", 1, udtAppointment.SchedDate)
    strEndRange = DateAdd("d", -1, udtAppointment.SchedDate)
    strWeek = DateAdd("d", -30, udtAppointment.SchedDate)
    If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
    strSQL = "SELECT * FROM APPOINTMENTS_BLOCKS WHERE " _
        & "PROVIDER_ID = " & udtAppointment.ProviderID & " AND FACILITY_ID = " & udtFacility.FacID & " AND " _
        & "(CONVERT(DATETIME, START_DATE, 101) BETWEEN '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "' AND '" & Format(DateAdd("d", 1, udtAppointment.SchedDate), "mm/dd/yyyy") & "' OR " _
        & "CONVERT(DATETIME, END_DATE, 101) BETWEEN '" & Format(udtAppointment.SchedDate, "mm/dd/yyyy") & "' AND '" & Format(DateAdd("d", 1, udtAppointment.SchedDate), "mm/dd/yyyy") & "') AND " _
        & "ACTIVE = 1  ORDER BY CONVERT(DATETIME,START_DATE, 101) DESC, REASON, CONVERT(DATETIME,START_TIME, 108) DESC, FACILITY_ID"
'    strSQL = "Get_Provider_Appt_Block_Range " & udtAppointment.ProviderID & ", " & udtFacility.FacID & ", '" _
'        & Format(strTestTime, "mm/dd/yyyy") & "', '" & Format(strEndRange, "mm/dd/yyyy") & "', '" & Format(strWeek, "mm/dd/yyyy") & "'"
    Set rsBlock = cnBlock.Execute(strSQL)
    Do Until rsBlock.EOF
        Clear_Block
        booSkip = False
        strRepDate = vbNullString
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
        If strStartRange <> vbNullString Then
            If DateDiff("n", Format(strStartRange, "hh:mm"), Format(udtBlock.StartTime, "hh:mm")) > 0 Then
                booSkip = True
'            ElseIf DateDiff("n", Format(udtBlock.StartTime, "hh:mm"), Format(strStartRange, "hh:mm")) = 0 Then
'                booSkip = False
            ElseIf DateDiff("n", Format(strEndRange, "hh:mm"), Format(udtBlock.StartTime, "hh:mm")) > 0 Then
                booSkip = True
            End If
'           strEndRange = udtBlock.EndTime
        End If
        If Not booSkip Then
            lngDateDiff = 0
            If Trim(udtBlock.DayWeek) <> vbNullString Then strRepDate = DayWeek(udtBlock.DayWeek)
            If IsDate(udtBlock.StartDate) And (Trim(udtBlock.DayWeek) = vbNullString Or UCase(udtBlock.DayWeek) = GetDay(udtAppointment.SchedDate)) Then
               If Not IsDate(strRepDate) Then strRepDate = udtBlock.EndDate
               If udtBlock.Duration = "Permanent" Then
                  lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), udtBlock.StartDate)
                  If lngDateDiff <= 0 Then
                     lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), strRepDate)
                     If lngDateDiff >= 0 Then
                        If mstrApptMode = "Day" Then
                           DayMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, strRepDate, udtBlock.ProviderID
                        ElseIf mstrApptMode = "Weekly" Then
                           WeekMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, udtBlock.StartDate, udtBlock.ProviderID
                        End If
                     End If
                  End If
               Else
                  lngDateDiff = DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate)
                  If lngDateDiff >= 0 Then
                     lngDateDiff = DateDiff("d", udtAppointment.SchedDate, udtBlock.EndDate)
                     If lngDateDiff >= 0 Then
                        If mstrApptMode = "Day" Then
                           DayMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, udtAppointment.SchedDate, udtBlock.ProviderID
                        ElseIf mstrApptMode = "Weekly" Then
                           WeekMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, udtAppointment.SchedDate, udtBlock.ProviderID
                        End If
                        If DateDiff("h", Format(udtBlock.StartTime, "hh:mm"), Format(udtBlock.EndTime, "hh:mm")) > 7 And _
                            DateDiff("d", Format(udtBlock.StartDate, "mm/dd/yyyy"), Format(udtBlock.EndDate, "mm/dd/yyyy")) >= 0 Then
                            Exit Do
                        End If
                     End If
                  End If
               End If
            End If
            If udtBlock.Reason = "Holiday" Then
                Exit Do
            End If
            If strTestDay <> udtBlock.DayWeek Then
                strStartRange = vbNullString
                strEndRange = vbNullString
            End If
            lngDateDiff = DateDiff("h", udtBlock.StartTime, udtBlock.EndTime)
            If lngDateDiff > 7 Then
                strStartRange = udtBlock.StartTime
                strEndRange = udtBlock.EndTime
            End If
            strTestDay = udtBlock.DayWeek
        End If
    Loop
    If cnBlock.State = adStateOpen Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    Check_Conflict_Facility
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_BlockTime:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Set_BlockTime", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Set_Permanent_Blocks()
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim i%
Dim lngDateDiff As Long
Dim strRepDate As String
Dim booSkip As Boolean
Dim strStartRange As String
Dim strEndRange As String
Dim strTestTime As String
Dim strTestDay As String
Dim strWeek As String

    On Error GoTo Error_Set_Permanent_Blocks
    Screen.MousePointer = vbHourglass
    If udtAppointment.SchedDate = vbNullString Then Exit Function
    strTestTime = DateAdd("d", 1, udtAppointment.SchedDate)
    strEndRange = DateAdd("d", -1, udtAppointment.SchedDate)
    strWeek = DateAdd("d", -30, udtAppointment.SchedDate)
    If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
'    strSQL = "Get_Provider_Permanent_Blocks_Prov_Fac_Duration_EndDate " & udtAppointment.ProviderID & "," _
'        & udtFacility.FacID & " ,'Permanent' ,'" & Format(strEndRange, "mm/dd/yyyy") & "'"
    strSQL = "SELECT * FROM APPOINTMENTS_BLOCKS WHERE ACTIVE = 1 " _
        & "AND PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
        & "DURATION = 'Permanent' AND FACILITY_ID =  " & udtFacility.FacID & " AND " _
        & "CONVERT(DATETIME, END_DATE, 101) > '" & Format(strEndRange, "mm/dd/yyyy") & "' ORDER " _
        & "BY CONVERT(DATETIME, START_DATE, 101) DESC, REASON, CONVERT(DATETIME, START_TIME, 108) DESC"
    Set rsBlock = cnBlock.Execute(strSQL)
    Do Until rsBlock.EOF
        Clear_Block
        booSkip = False
        strRepDate = vbNullString
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
        lngDateDiff = 0
        If Trim(udtBlock.DayWeek) <> vbNullString Then strRepDate = DayWeek(udtBlock.DayWeek)
        If IsDate(udtBlock.StartDate) And (Trim(udtBlock.DayWeek) = vbNullString Or UCase(udtBlock.DayWeek) = GetDay(udtAppointment.SchedDate)) Then
           If Not IsDate(strRepDate) Then strRepDate = udtBlock.EndDate
           If udtBlock.Duration = "Permanent" Then
              lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), udtBlock.StartDate)
              If lngDateDiff <= 0 Then
                 lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), strRepDate)
                 If lngDateDiff >= 0 Then
                    If mstrApptMode = "Day" Then
                       DayMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, strRepDate, udtBlock.ProviderID
                    ElseIf mstrApptMode = "Weekly" Then
                       WeekMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, udtBlock.StartDate, udtBlock.ProviderID
                    End If
                 End If
              End If
           Else
              lngDateDiff = DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate)
              If lngDateDiff >= 0 Then
                 lngDateDiff = DateDiff("d", udtAppointment.SchedDate, udtBlock.EndDate)
                 If lngDateDiff >= 0 Then
                    If mstrApptMode = "Day" Then
                       DayMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, udtAppointment.SchedDate, udtBlock.ProviderID
                    ElseIf mstrApptMode = "Weekly" Then
                       WeekMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, udtAppointment.SchedDate, udtBlock.ProviderID
                    End If
                 End If
              End If
           End If
        End If
        If udtBlock.Reason = "Holiday" Then
            Exit Do
        End If
        If strTestDay <> udtBlock.DayWeek Then
            strStartRange = vbNullString
            strEndRange = vbNullString
        End If
        lngDateDiff = DateDiff("h", udtBlock.StartTime, udtBlock.EndTime)
        If lngDateDiff > 7 Then
            strStartRange = udtBlock.StartTime
            strEndRange = udtBlock.EndTime
        End If
        strTestDay = udtBlock.DayWeek
    Loop
    If cnBlock.State = adStateOpen Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_Permanent_Blocks:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Set_Permanent_Blocks", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Get_Changes(vstrDate As String)
Dim i%
Dim strSQL As String
Dim cnChange As New ADODB.Connection
Dim rsChange As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld  As ADODB.Field
Dim lngProviderID  As Long
Dim Arr
Dim sDate As String
Dim sProviderName As String
Dim strLastName As String
Dim strFirstName As String
Dim strStatus As String
Dim strReason As String
Dim strText As String
Dim strCollect As String
Dim strLine As String
Dim lngPatID As Long
Dim strDuration As String
Dim strType As String
Dim strTime As String
Dim lngRow As Long
Dim lngCol As Long
Dim bExit As Boolean
Dim lngTest As Long
Dim strName As String


    On Error GoTo Error_Get_Changes
    Screen.MousePointer = vbHourglass
    grdData.Redraw = False
    sDate = Format(vstrDate, "mm/dd/yyyy")
    If cnChange.State = adStateClosed Then cnChange.open psConnect(0)
    strSQL = "Get_Appt_Change_Provider " & udtAppointment.ProviderID & ",'" _
        & udtAppointment.SchedType & "', '" & Format(sDate, "mm/dd/yyyy") & "'," _
        & udtFacility.FacID
    Set rsChange = cnChange.Execute(strSQL)
    With rsChange
        Do Until .EOF
            strName = vbNullString
            strCollect = vbNullString
            strLastName = vbNullString
            strFirstName = vbNullString
            strStatus = vbNullString
            strReason = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PROVIDER_ID"
                                sProviderName = Fld.Value
                            Case "REASON"
                                strReason = Fld.Value
                            Case "STATUS"
                                strStatus = Fld.Value
                            Case "PATIENT_ID"
                                lngPatID = Fld.Value
                            Case "DURATION"
                                strDuration = Fld.Value
                            Case "APPT_TYPE"
                                strType = Fld.Value
                            Case "APPT_TIME"
                                strTime = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If strType = vbNullString Then strType = "Patient"
             Select Case IsProperCase(strType)
                Case "Patient"
                    If lngPatID > 0 Then
                        Get_PatientName lngPatID
                        strName = udtPatientName.LastName & ", " & udtPatientName.FirstName
                    End If
                Case "Vendor", "Business"
                    Get_VendorInfo lngPatID
                    strName = clsVendor.Name
                Case "Legal", "Attorney"
                    Get_Legal_Contact lngPatID
                    If clsLegal.Practice <> vbNullString Then
                        strName = clsLegal.Name & ", " & clsLegal.Practice
                    Else
                        strName = clsLegal.Name
                    End If
                Case "Physician"
                    Get_Referral lngPatID
                    strName = clsReferral.Physician
            End Select
            Select Case IsProperCase(strStatus)
                Case "Rescheduled"
                    If IsProperCase(strReason) = "Change" Then strReason = vbNullString
                    If Trim(strName) <> vbNullString Then
                        bExit = False
                        With grdChanges
                            For lngTest = 1 To .MaxRows
                                .Col = 1
                                .Row = lngTest
                                If .GetRowItemData(lngTest) = lngPatID Then
                                    bExit = True
                                    Exit For
                                End If
                            Next lngTest
                            If Not bExit Then
                                lngRow = .MaxRows + 1
                                .MaxRows = lngRow
                                .SetRowItemData lngRow, lngPatID
                                lngCol = 1
                                Format_CellTypeStatic grdChanges, lngRow, lngCol, "1", strName, "TypeHAlignLeft", "", "", 0, strType
                                grdChanges.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                lngCol = 2
                                Format_CellTypeStatic grdChanges, lngRow, lngCol, "1", Trim(strStatus), "TypeHAlignCenter"
                                grdChanges.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        End With
                    End If
                Case "Cancelled", "No Show", "Late Cancellation"
                    Remove_Data lngPatID, strDuration, strTime
                    With grdChanges
                        For lngTest = 1 To .MaxRows
                            If .GetRowItemData(lngRow) = lngPatID Then
                                .DeleteRows lngRow, 1
                                .MaxRows = .MaxRows - 1
                                Exit For
                            End If
                        Next lngTest
                    End With
                    If Trim(strName) <> vbNullString Then
                        bExit = False
                        With grdChanges
                            For lngTest = 1 To .MaxRows
                                .Col = 1
                                .Row = lngTest
                                If .GetRowItemData(lngTest) = lngPatID Then
                                    bExit = True
                                    Exit For
                                End If
                            Next lngTest
                            If Not bExit Then
                                lngRow = .MaxRows + 1
                                .MaxRows = lngRow
                                .SetRowItemData lngRow, lngPatID
                                lngCol = 1
                                Format_CellTypeStatic grdChanges, lngRow, lngCol, "1", strName, "TypeHAlignLeft", "", "", 0, strType
                                grdChanges.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                lngCol = 2
                                Format_CellTypeStatic grdChanges, lngRow, lngCol, "1", Trim(strStatus), "TypeHAlignCenter"
                                grdChanges.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        End With
                    End If
            End Select
        Loop
    End With
    mbooByPass = False
    If cnChange.State = adStateOpen Then
        cnChange.Close
        Set cnChange = Nothing
        Set rsChange = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Changes:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Changes", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Public Sub Manage_Colors(vstrBlockReason As String, vlngPatientID As Long)
Dim lngBlockRow As Long
Dim lngCol As Long
Dim lngRow As Long
Dim mbLeaveCol As Boolean
Dim lngCurPatient As Long

    On Error GoTo Error_Manage_Colors
    Screen.MousePointer = vbHourglass
    
    If Trim(vstrBlockReason) <> vbNullString Then
        If TextFound("Can", vstrBlockReason, Len("Can"), True) Then
                clsGridFormat.BackColor = &HC0FFC0
        ElseIf vstrBlockReason = "Resched" Or vstrBlockReason = "Delete" Then
                clsGridFormat.BackColor = &HC0FFC0
        ElseIf TextFound("Present", vstrBlockReason, Len("Present"), True) Then
                clsGridFormat.BackColor = &HFFC0FF
        ElseIf TextFound("Present", vstrBlockReason, Len("Check Out"), True) Then
               clsGridFormat.BackColor = &HC0FFFF
        ElseIf TextFound("Confirmed", vstrBlockReason, Len("Confirmed"), True) Then
               clsGridFormat.BackColor = &HC0E0FF
'        ElseIf TextFound("No Show", vstrBlockReason, Len("No Show"), True) Then
'                clsGridFormat.BackColor = &HC0FFC0
        ElseIf TextFound("EAP", vstrBlockReason, Len("EAP"), True) Then
            clsGridFormat.BackColor = &HF4D9EA
        Else
                clsGridFormat.BackColor = &H80000005
        End If
        With grdData
            If clsGridFormat.BackColor <> &HFFFFFF Then
                For lngCol = 1 To 2
                    For lngRow = 1 To .MaxRows
                        grdData.Redraw = False
                        .Col = lngCol
                        .Row = lngRow
                        If Trim(.CellTag) <> vbNullString Then
                            If IsNumeric(.CellTag) And Len(.CellTag) > 1 Then
                                If .CellTag = lngCurPatient Then
                                     If lngCurPatient <> 0 Then
                                        .BlockMode = True
                                            .Col = lngCol
                                            .Col2 = lngCol
                                            .Row = lngRow
                                            .Row2 = lngRow
                                            If clsGridFormat.BackColor = 0 Then clsGridFormat.BackColor = &H80000005
                                            If clsGridFormat.BackColor <> 0 Then grdData.BackColor = clsGridFormat.BackColor
                                        .BlockMode = False
                                        mbLeaveCol = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next lngRow
                    If mbLeaveCol Then Exit For
                Next lngCol
            End If
            grdData.Redraw = True
        End With
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Manage_Colors:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Manage_Colors", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Function Remove_Data(vlngPatID As Long, vstrDuration As String, Optional vstrTime As String)
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Arr
Dim lngRow As Long
Dim varCellValue
Dim lngProvID As Long
Dim lngCol As Long
Dim lngFoundID As Long
Dim lngSpan As Long
Dim lngRowTime As Long
Dim lngSpanRow As Long
Dim booRet As Boolean
Dim booExit As Boolean
Dim strFoundStart As String
Dim strFoundEnd As String
Dim lngApptCnt As Long
Dim strStart As String
Dim strEnd As String
Dim lngDur As Long

    On Error GoTo Error_Remove_Data
    Screen.MousePointer = vbHourglass
    Get_NewSchedPatient vlngPatID
    lngCol = 0
    strStart = Format(vstrTime)
    strStart = Format(strStart, "hh:mm AM/PM")
    vstrDuration = Replace(vstrDuration, "minutes", "")
    vstrDuration = Trim(vstrDuration)
    lngDur = Val(vstrDuration) + 1
    strEnd = DateAdd("n", lngDur, vstrTime)
    strEnd = Format(strEnd, "hh:mm AM/PM")
    Arr = Split(vstrDuration, " ")
'    If UBound(Arr) > 0 Then
'       If IsNumeric(Arr(0)) Then lngSpanRow = Val(Arr(0) / lngRowTime)
'    Else
'       If IsNumeric(vstrDuration) Then lngSpanRow = Val(vstrDuration / lngRowTime)
'    End If
    With grdData
        For lngRow = 0 To grdData.MaxRows
            .GetText 0, lngRow, varCellValue
            If Trim(varCellValue) = vstrTime Then
                Exit For
            End If
        Next lngRow
        For lngCol = 1 To grdData.MaxCols
            If grdData.GetColItemData(lngCol) = udtAppointment.ProviderID Then
                .Col = lngCol
                .Row = lngRow
                If .CellTag <> vbNullString Then
                    If .CellTag = udtSchedPatient.IntId Then
                        IsCellSpan lngRow, lngCol
                        Format_CellTypeEdit grdData, lngRow, lngCol, "1", "", "TypeHAlignCenter", "", NormalColor
                        grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Exit For
                    End If
                End If
            End If
        Next lngCol
    End With
    lblProvider.Caption = Trim(cctProvider.Text)
    Screen.MousePointer = vbDefault
    grdData.Redraw = False
    Exit Function
Error_Remove_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", " Remove_Data ", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Manage_Surgery(vstrType As String)
Dim vDate
Dim strDate As String
Dim Arr
Dim vPatient
Dim lngPatientID As Long
Dim vTime
Dim sFirstName As String
Dim sLastName As String
Dim strSQL As String
Dim cnData As New ADODB.Connection
Dim rsData As New ADODB.Recordset
Dim Idx As Integer
Dim lngSpanRow As Long
Dim strSpan As String
Dim lngCol As Long
Dim lngRow As Long
Dim strMoveTime As String
Dim varCellValue
Dim lngRowTime As Long

    Select Case vstrType
        Case "DblClick"
            frmMovePatient.Show vbModal
            If frmMovePatient.mbMovePatient Then
                grdData.Col = mlngGridCol
                grdData.Row = mlngGridRow
                If grdData.CellTag <> vbNullString Then udtSchedPatient.IntId = grdData.CellTag
                    If Val(udtSchedPatient.IntId) > 0 Then
                    grdData.ClearRange mlngGridCol, mlngGridRow, mlngGridCol, mlngGridRow, False
                    grdData.GetText mlngGridCol, mlngGridRow, varCellValue
                    If Trim(varCellValue) <> vbNullString Then udtSchedPatient.Collection = Trim(varCellValue)
                    If udtSchedPatient.Collection <> vbNullString And TextFound("Operating", frmMovePatient.mstrResource, Len("Operating"), True) Then 'And Resource = "Operating Room"
                        Arr = Split(udtSchedPatient.Collection, vbNewLine)
                        If UBound(Arr) > 0 Then
                            For Idx = LBound(Arr) To UBound(Arr)
                                If TextFound("minutes", Arr(Idx), Len("minute"), True) Then
                                    udtSchedPatient.Duration = Replace(Arr(Idx), "minutes", "")
                                     If IsNumeric(udtSchedPatient.Duration) Then lngSpanRow = Val(udtSchedPatient.Duration / lngRowTime)
                                     Exit For
                                End If
                            Next Idx
                        End If
                    End If
                    Arr = Split(LookUpColName("frmMovePatient.mstrResource"), "|")
                    If UBound(Arr) > 0 Then
                        lngCol = Arr(0)
                    End If
                    'Need Row
                    lngRow = 0
                    If udtAppointment.SchedTime <> vbNullString Then
                        
                    End If
                    For Idx = 1 To grdData.MaxRows
                        grdData.GetText 0, Idx, vTime
                        If vTime = vbNullString Then
                          grdData.GetText 0, Idx - 1, vTime
                        End If
                        If Format(strMoveTime, "hh:mm AM/PM") = CStr(vTime) Then
                            lngRow = Idx
                            Exit For
                        End If
                    Next Idx
                    
                    If lngSpanRow >= 0 Then strSpan = lngSpanRow
                    Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, strSpan, udtSchedPatient.Collection, "TypeHAlignLeft", CStr(udtSchedPatient.IntId), udtColor.Surgery
                    grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow + (Val(strSpan) - 1), CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Else
                    MsgBox "Please select a valid Patient Row to move!", vbInformation
                End If
            End If
    End Select
    Exit Function
End Function
Function Check_Schedule()
Dim lngRow As Long
Dim lngPatientID As Long
Dim cnSched As New ADODB.Connection
Dim rsSched As New ADODB.Recordset
Dim strSQL As String
Dim strApptDate As String
Dim Arr
Dim vApptDate
Dim strStatus As String
Dim lngCol As Long
Dim varCellValue
Dim strTag As String
Dim strNote As String
Dim lngColor As Long
Dim lngRowTime As Long


    On Error GoTo Error_Check_Schedule
    Screen.MousePointer = vbHourglass
    Get_Practice
    grdData.Redraw = False
    strApptDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    With grdData
        If mstrApptMode = "Day" Then
                For lngRow = 1 To .MaxRows
                    varCellValue = vbNullString
                    strNote = vbNullString
                    lngPatientID = 0
                    .Col = 1
                    .Row = lngRow
                    If .CellTag <> vbNullString And IsNumeric(.CellTag) Then lngPatientID = Val(.CellTag)
                    If lngPatientID <> 0 Then
                        cnSched.open psConnect(0)
                            strSQL = "Get_Status_Duration_PatientAppt " & lngPatientID & ", '" & Format(strApptDate, "mm/dd/yyyy") & "'"
                            Set rsSched = cnSched.Execute(strSQL)
                                If Not rsSched.EOF Then
                                    If rsSched.Fields("DURATION") <> vbNullString Then udtSchedPatient.Duration = rsSched.Fields("DURATION")
                                    Select Case UCase(clsApptPref.SchedInterval)
                                        Case "EVERY 30 MIN"
                                            lngRowTime = 30
                                        Case "EVERY 20 MIN"
                                            lngRowTime = 20
                                        Case "EVERY 15 MIN"
                                            lngRowTime = 15
                                        Case "EVERY 10 MIN"
                                            lngRowTime = 10
                                        Case "EVERY 5 MIN"
                                            lngRowTime = 5
                                    End Select
                                    Arr = Split(udtSchedPatient.Duration, " ")
                                    If UBound(Arr) > 0 Then
                                        If IsNumeric(Arr(0)) And lngRowTime > 0 Then strNote = Val(Arr(0) / lngRowTime)
                                    Else
                                        strNote = udtSchedPatient.Duration
                                    End If
                                    If Not IsNull(rsSched.Fields("STATUS")) Then
                                        udtSchedPatient.Status = rsSched.Fields("STATUS")
                                            clsGridFormat.BackColor = 0
                                            Select Case IsProperCase(udtAppointment.SchedType)
                                                Case "Physician"
                                                    clsGridFormat.BackColor = GridNonPatient
                                                Case "Vendor"
                                                    clsGridFormat.BackColor = GridNonPatient
                                            End Select
                                            Select Case IsProperCase(udtSchedPatient.Status)
                                                Case "Present"
                                                    clsGridFormat.BackColor = grdCurItem
                                                Case "No Show", "Cancelled"
                                                    clsGridFormat.BackColor = grdGray
                                                Case "Check Out"
                                                    clsGridFormat.BackColor = DataColor
                                                Case "Confirmed"
                                                    clsGridFormat.BackColor = GridConfir
                                            End Select
                                            Manage_Colors strStatus, lngPatientID
                                            lngCol = 1
                                            grdData.Col = 1
                                            grdData.Row = lngRow
                                            If grdData.CellTag <> vbNullString Then clsGridFormat.CellTag = grdData.CellTag
                                            If grdData.CellNote <> vbNullString Then clsGridFormat.CellNote = grdData.CellNote
                                            grdData.GetText grdData.Col, lngRow, varCellValue
                                            Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, strNote, CStr(varCellValue), "TypeHAlignCenter", clsGridFormat.CellTag, clsGridFormat.BackColor, clsGridFormat.CellNote
                                            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow + (Val(strNote) - 1), CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                        'End If
                                    End If
                                End If
                        If cnSched.State = adStateOpen Then
                            cnSched.Close
                            Set cnSched = Nothing
                        End If
                    End If
                Next lngRow
        Else
            For lngCol = 1 To .MaxCols
                .GetText lngCol, SpreadHeader + 2, vApptDate
                Arr = Split(vApptDate, "-")
                If UBound(Arr) > 0 Then
                    vApptDate = Arr(1)
                    vApptDate = Format(vApptDate, "mm/dd/yyyy")
                    If vApptDate <> vbNullString Then
                         For lngRow = 1 To .MaxRows
                            If .CellTag <> vbNullString Then lngPatientID = Val(.CellTag)
                                If lngPatientID <> 0 Then
                                    cnSched.open psConnect(0)
                                        strSQL = "Get_Status_Duration_PatientAppt " & lngPatientID & ", '" & Format(strApptDate, "mm/dd/yyyy") & "'"
                                        Set rsSched = cnSched.Execute(strSQL)
                                            If Not rsSched.EOF Then
                                                If Not IsNull(rsSched.Fields("STATUS")) Then Manage_Colors rsSched.Fields("STATUS"), lngPatientID
                                                If rsSched.Fields("DURATION") <> vbNullString Then udtSchedPatient.Duration = rsSched.Fields("DURATION")
                                                Select Case UCase(clsApptPref.SchedInterval)
                                                    Case "EVERY 30 MIN"
                                                        lngRowTime = 30
                                                    Case "EVERY 20 MIN"
                                                        lngRowTime = 20
                                                    Case "EVERY 15 MIN"
                                                        lngRowTime = 15
                                                    Case "EVERY 10 MIN"
                                                        lngRowTime = 10
                                                    Case "EVERY 5 MIN"
                                                        lngRowTime = 5
                                                End Select
                                                Arr = Split(udtSchedPatient.Duration, " ")
                                                If UBound(Arr) > 0 Then
                                                    If IsNumeric(Arr(0)) And lngRowTime > 0 Then strNote = Val(Arr(0) / lngRowTime)
                                                Else
                                                    strNote = udtSchedPatient.Duration
                                                End If
                                            End If
                                            grdData.Col = lngCol
                                            grdData.Row = lngRow
                                            If grdData.CellTag <> vbNullString Then clsGridFormat.CellTag = grdData.CellTag
                                            If grdData.CellNote <> vbNullString Then clsGridFormat.CellNote = grdData.CellNote
                                            grdData.GetText grdData.Col, lngRow, varCellValue
                                            Format_CellTypeEdit_RowSpan grdData, lngRow, lngCol, strNote, CStr(varCellValue), "TypeHAlignCenter", clsGridFormat.CellTag, Color, clsGridFormat.CellNote
                                            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow + (Val(strNote) - 1), CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    If cnSched.State = adStateOpen Then
                                        cnSched.Close
                                        Set cnSched = Nothing
                                    End If
                                End If
                        Next lngRow
                    End If
                End If
            Next lngCol
        End If
    End With
    grdData.Redraw = True
    Screen.MousePointer = vbDefault
    grdData.Redraw = False
    Exit Function
Error_Check_Schedule:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Check_Schedule", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Sub Format_Grid()
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim strEnd As String
Dim lngDateDiff As Long
Dim lngStartRow  As Long
Dim varCellValue


    On Error GoTo Error_Format_Grid
    Screen.MousePointer = vbHourglass
    Get_TimeSlot
    With grdData
        .ColHeaderRows = 4
        .Redraw = False
        .MaxRows = 0
        .ColWidth(0) = 10 * mlngSystemFileSize
        For Idx = mStartIdx To UBound(mStartHour)
            lngRow = .MaxRows + 1
            .MaxRows = lngRow
            Format_CellTypeFixed grdData, lngRow, 0, "0", mStartHour(Idx) & " AM", "TypeHAlignLeft", mStartHour(Idx) & " AM", HeaderColor
            grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Select Case UCase(clsApptPref.SchedInterval)
                Case "EVERY HALF HOUR", "EVERY 30 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, lngCol, "0", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", "TypeHAlignLeft", mStartHour(Idx) & " AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 15 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":15 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":15 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 10 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":10 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":15 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":20 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":40 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":50 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 5 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":05 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":15 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":10 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":15 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":20 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":25 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":30 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":35 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":40 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":50 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":55 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 20 MIN"
                    lngRow = .MaxRows + 1
                   .MaxRows = lngRow
                   Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":20 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":20 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                
                    lngRow = .MaxRows + 1
                   .MaxRows = lngRow
                   Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mStartHour(Idx), 1, 2) & ":40 AM", "TypeHAlignLeft", Mid$(mStartHour(Idx), 1, 2) & ":40 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End Select
        Next Idx
        lngStartRow = lngRow
        For Idx = LBound(mEndHour) To Left(mEndHour(mEndIdx), 2)
            lngRow = .MaxRows + 1
            .MaxRows = lngRow
            Format_CellTypeFixed grdData, lngRow, 0, "0", mEndHour(Idx) & " PM", "TypeHAlignLeft", mEndHour(Idx) & " PM", HeaderColor
            grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Select Case UCase(clsApptPref.SchedInterval)
                Case "EVERY HALF HOUR", "EVERY 30 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 15 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":15 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":15 PM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":45 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 PM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 10 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":10 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":15 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":20 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":30 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":40 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":50 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 5 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":05 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":15 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":10 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":30 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":15 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":20 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":25 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":30 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":35 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":40 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":45 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":50 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":55 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":45 AM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Case "EVERY 20 MIN"
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":20 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":20 PM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    lngRow = .MaxRows + 1
                    .MaxRows = lngRow
                    Format_CellTypeFixed grdData, lngRow, 0, "0", Mid$(mEndHour(Idx), 1, 2) & ":40 PM", "TypeHAlignLeft", Mid$(mEndHour(Idx), 1, 2) & ":40 PM", HeaderColor
                    grdData.SetCellBorder 0, lngRow, 0, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End Select
        Next Idx
        .Row = 1
        .LeftCol = 1
        For Idx = lngStartRow To grdData.MaxRows
        lngRow = Idx
        .GetText 0, lngRow, varCellValue
            strEnd = mEndHour(mEndIdx)
            If Trim(varCellValue) <> vbNullString Then
                lngDateDiff = DateTime.DateDiff("n", Format(varCellValue, "hh:mm"), Format(mEndHour(mEndIdx) & "PM", "hh:mm"))
                If lngDateDiff < 0 Then
                     grdData.DeleteRows lngRow, 1
                    .MaxRows = .MaxRows - 1
                    Idx = Idx - 1
                End If
            End If
'            If mlngSystemFileSize > 0 Then
'                .RowHeight(lngRow) = .RowHeight(lngRow) * mlngSystemFileSize
'            End If
        Next Idx
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Grid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Grid", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Function Format_CellTypeFixed(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vlngCol_ID As Long)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeStaticText
            .FontName = "Times New Roman"
            .FontSize = 10
            .FontBold = True
            .ForeColor = GridBlue
            If vlngBackColor > 0 Then .BackColor = vlngBackColor
            .SetText vlngCol, vlngRow, Trim(vstrText)
            .Text = Trim(vstrText)
            .TypeVAlign = TypeVAlignCenter
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
              .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString Then
              lngSpan = vstrSpan
              If lngSpan > 1 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
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
              .FontSize = 10
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional FontSize As Long)
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
                .TypeEditMultiLine = True
                .FontName = "Times New Roman"
                If FontSize > 0 Then
                    .FontSize = FontSize
                Else
                    .FontSize = 10
                End If
                .FontBold = True
                .ForeColor = GridBlue
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
                  .RowHeight(vlngRow) = CellNote + 7
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
Private Function Format_CellTypeWeekCol(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vlngCol_ID As Long)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeEdit
            .FontName = "Times New Roman"
            .FontSize = 10
            .FontBold = True
            .ForeColor = GridBlue
            If vlngBackColor > 0 Then .BackColor = vlngBackColor
            .SetText vlngCol, vlngRow, Trim(vstrText)
            .Text = Trim(vstrText)
            .TypeVAlign = TypeVAlignCenter
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
              .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString Then
              lngSpan = vstrSpan
              If lngSpan > 1 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            If .RowHeight(vlngRow) < CellNote Then
              .RowHeight(vlngRow) = CellNote + 7
            End If
        .BlockMode = False
    End With
End Function

Private Function Format_CellTypeEdit_RowSpan(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrNote As String, Optional vstrColSpan As String)
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
              .FontSize = 10
              .FontBold = True
              'If vbooUnderline Then .FontUnderline = True
              If vlngBackColor > 0 Then .BackColor = vlngBackColor
              If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
              ElseIf Trim(vstrSpan) <> vbNullString Then
                If vstrColSpan = vbNullString Then
                    lngSpan = vstrSpan
                    If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, 1, lngSpan
                Else
                    lngSpan = vstrSpan
                    If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, Val(vstrColSpan), lngSpan
                End If
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
              If vstrNote <> vbNullString Then .CellNote = vstrNote
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
            .FontSize = 10
            .FontBold = True
            .ForeColor = GridBlue
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
Private Function LookUpColName(vstrName As String) As String
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

    With grdData
        LookUpColName = vbNullString
        For lngCol = 1 To .MaxCols
            .Col = lngCol
            .Row = SpreadHeader + 3
            .GetText lngCol, SpreadHeader + 3, varCellValue
            If Trim(varCellValue) <> vbNullString And TextFound(vstrName, Trim(varCellValue), Len(vstrName), True) Then
                strFound = Trim(varCellValue)
                LookUpColName = lngCol & "|" & SpreadHeader + 3
                booExit = True
                Exit For
            End If
        Next lngCol
    End With
End Function

Private Sub Format_Week(ipDate As String)
Dim sDate As Date
Dim lngCol As Long
Dim FirstDay As VbDayOfWeek
Dim lngWeekDays As Long
Dim strOption As String
Dim lngRow As Long

    On Error GoTo Error_Format_Week
    Screen.MousePointer = vbHourglass
    Format_Grid
    Get_TimeSlot
    If clsApptPref.SchedWeek = vbNullString Then clsApptPref.SchedWeek = "Mon - Friday"
    Select Case clsApptPref.SchedWeek
        Case "Mon - Friday"
            lngWeekDays = 5
        Case "Mon - Saturday"
            lngWeekDays = 6
        Case "Mon - Sunday"
            lngWeekDays = 7
    End Select
    strSelDate = Format(ipDate, "mm/dd/yyyy")
    FirstDay = Weekday(strSelDate, vbMonday) - 1
    sDate = DateAdd("d", (-FirstDay), strSelDate)
    With grdData
        .MaxCols = 1
        .ClearRange 1, SpreadHeader + 1, .MaxCols, SpreadHeader + 1, False
        .ClearRange 1, SpreadHeader + 2, .MaxCols, SpreadHeader + 2, False
        .ClearRange 1, SpreadHeader + 3, .MaxCols, SpreadHeader + 3, False
        .ClearRange 1, SpreadHeader + 4, .MaxCols, SpreadHeader + 3, False
         .Redraw = False
         .ScrollBarTrack = ScrollBarTrackVertical
         .AllowDragDrop = True
         .MaxCols = lngWeekDays
         For lngCol = 1 To .MaxCols
            .RemoveCellSpan lngCol, SpreadHeader
            .ColWidth(lngCol) = 20 * mlngSystemFileSize
         Next lngCol
         Format_Week_Col 0, 1, "Monday", Format(sDate, "mm/dd/yyyy")
         sDate = DateAdd("d", 1, sDate)
         Format_Week_Col 0, 2, "Tuesday", Format(sDate, "mm/dd/yyyy")
         sDate = DateAdd("d", 1, sDate)
         Format_Week_Col 0, 3, "Wednesday", Format(sDate, "mm/dd/yyyy")
         sDate = DateAdd("d", 1, sDate)
         Format_Week_Col 0, 4, "Thursday", Format(sDate, "mm/dd/yyyy")
         sDate = DateAdd("d", 1, sDate)
         Format_Week_Col 0, 5, "Friday", Format(sDate, "mm/dd/yyyy")
         If lngWeekDays > 5 Then
            sDate = DateAdd("d", 1, sDate)
            Format_Week_Col 0, 6, "Saturday", Format(sDate, "mm/dd/yyyy")
         End If
         If lngWeekDays > 6 Then
            sDate = DateAdd("d", 1, sDate)
            Format_Week_Col 0, 7, "Sunday", Format(sDate, "mm/dd/yyyy")
         End If
         .Redraw = True
         For lngCol = 1 To .MaxCols
            For lngRow = 1 To .MaxRows
                grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdData.RowHeight(lngRow) = grdData.RowHeight(lngRow) * mlngSystemFileSize
            Next lngRow
         Next lngCol
    End With
    If grdData.MaxCols > 1 Then
        Format_CellTypeEdit grdData, SpreadHeader, 1, CStr(grdData.MaxCols), cctFacilities.Text, "TypeHAlignCenter", "", HeaderColor, 14
        grdData.SetCellBorder 1, SpreadHeader, 1, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdData
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Week:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Week", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Sub Format_Week_Col(ipRow As Long, ipCol As Long, ipDay As String, ipDate As String)
Dim lngRow As Long
Dim lngCol As Long
Dim lngTest As Long

    On Error GoTo Error_Format_Week_Col
    If ipRow = 0 Then ipRow = SpreadHeader
    lngCol = ipCol
'    If udtFacility.FacID = 0 Then
'        With grdFacility
'            For lngTest = 1 To .MaxRows
'                .Col = 1
'                .Row = lngTest
'                If .BackColor = grdCurItem Then
'                     udtFacility.FacID = grdFacility.GetRowItemData(lngTest)
'                     Get_Facilities udtFacility.FacID
'                End If
'            Next lngTest
'        End With
'    End If
    With grdData
        .ColHeaderRows = 6
        grdData.RemoveCellSpan 1, SpreadHeader + 2
        grdData.RemoveCellSpan 1, SpreadHeader + 3
        .RowHeight(SpreadHeader + 2) = 15
        Format_CellTypeFixed grdData, SpreadHeader + 2, lngCol, "1", ipDay & " - " & ipDate, "TypeHAlignCenter", "", HeaderColor
        grdData.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdData
        Format_CellTypeEdit grdData, SpreadHeader + 4, 1, "1", clsFacility.Name, "TypeHAlignCenter", CStr(udtAppointment.ProviderID), HeaderColor
        grdData.SetCellBorder 1, SpreadHeader + 4, grdData.MaxCols, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        For lngRow = 1 To .MaxRows
            Format_CellTypeEdit grdData, lngRow, lngCol, "1", " ", "TypeHAlignCenter"
        Next lngRow
        Add_Spacer SpreadHeader + 5, grdData
    End With
    Exit Sub
Error_Format_Week_Col:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_Week_Col", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Sub Get_TimeSlot()
Dim sValue As String
Dim Idx As Long
Dim Arr
Dim sDate As String
Dim strError As String


    grdData.Redraw = False
    clsApptPref.SchedInterval = UCase(clsApptPref.SchedInterval)
    If clsApptPref.SchedInterval = "EVERY HALF HOUR" Then clsApptPref.SchedInterval = "EVERY 30 MIN"
    Arr = Split(clsApptPref.SchedInterval, " ")
    If Arr(1) = 15 Then
        clsApptPref.SchedInterval = "EVERY 15 MIN"
    ElseIf Arr(1) = 5 Then
        clsApptPref.SchedInterval = "EVERY 5 MIN"
    ElseIf Arr(1) = 10 Then
        clsApptPref.SchedInterval = "EVERY 10 MIN"
    ElseIf Arr(1) = 20 Then
        clsApptPref.SchedInterval = "EVERY 20 MIN"
    ElseIf Arr(1) = 30 Then
        clsApptPref.SchedInterval = "EVERY 30 MIN"
    End If
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
    For Idx = 0 To UBound(mStartHour)
       If mStartHour(Idx) = clsApptPref.Appt_Start Then
          mStartIdx = Idx
          Exit For
       End If
    Next Idx
    For Idx = 0 To UBound(mEndHour)
       If mEndHour(Idx) = clsApptPref.Appt_End Then
          mEndIdx = Idx
          Exit For
       End If
    Next Idx
    grdData.Redraw = True
    Exit Sub
End Sub


Private Function Update_Data()
Dim iAnswer As Integer

       If mChange = True Then
          iAnswer = MsgBox("Do you want to save your changes?", vbYesNo, "Appointment")
          If iAnswer = vbYes Then
            ' UpddateDeata
'             If Not UpdateData Then
'                MsgBox "Problem Updating Data.", vbInformation
'             End If
          End If
       End If
       mChange = False
End Function


Private Function Update_EnctrEndDate(vstrType As String)
'Dim strSQL As String
'Dim cnData As New ADODB.Connection
'Dim rsData As New ADODB.Recordset
'Dim lngEnctrID As Long
'Dim sDate As String
'
'
'    On Error GoTo Error_Update_EnctrEndDate
'    Screen.MousePointer = vbHourglass
'    If cnData.State = adStateClosed Then cnData.open psConnect(0)
'    If udtAppointment.SchedTime <> vbNullString Then udtEnctr.EndDate = Format(udtAppointment.SchedTime, "mm/dd/yyyy hh:mm:ss AM/PM")
'    strSQL = "Select * from ENCOUNTER WHERE patient_id = " & frmPatientAppointment.mlngPatientID & " AND PURGE = 0  ORDER BY VISIT_TYPE"
'    Set rsData = cnData.Execute(strSQL)
'    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
'    With rsData
'        Do Until .EOF
'            If (!VISIT_TYPE) = vstrType Then
'                lngEnctrID = (!ENCOUNTER_ID)
'            End If
'            .MoveNext
'        Loop
'    End With
'    strSQL = "UPDATE ENCOUNTER SET ENCNTR_END_DATE = '" & udtEnctr.EndDate & "', UPDATE_DATE = '" & udtEnctr.EndDate & "' WHERE ENCOUNTER_ID = " & lngEnctrID
'    Set rsData = cnData.Execute(strSQL)
'    If cnData.State = adStateOpen Then
'        cnData.Close
'        Set cnData = Nothing
'        Set rsData = Nothing
'    End If
'    Screen.MousePointer = vbDefault
'    Exit Function
'Error_Update_EnctrEndDate:
'    Screen.MousePointer = vbDefault
'    Write_Error Err.Description, Err.Number, strSQL, "Update_EnctrEndDate", App.EXEName & "." & Me.Name
'    Exit Function
'    Resume Next
End Function
Private Function Update_SurgeryStatus(vstrVisitType As String, Optional vstrType As String)
Dim sDate As String
Dim strSQL As String
Dim strEncntrStart As String
Dim strEnctTime As String
Dim strVisitType As String
Dim strPatType As String
Dim strEnctrType As String
Dim strSerPlace As String
Dim intServCode As String
Dim lngFacID As Long
Dim cnStatus As New ADODB.Connection
Dim rsStatus As New ADODB.Recordset
Dim lngDateDiff As Long
Dim booinsert As Boolean
Dim lngEnctrID As Long

    On Error GoTo Error_Update_SurgeryStatus
    Screen.MousePointer = vbHourglass
    Clear_EnctrString
    If cnStatus.State = adStateClosed Then cnStatus.open psConnect(1)
   ' strSQL = "SELECT  CPT_CODE,  ENCNTR_TYPE, CPT_NAME, SERVICE_PLACE, PLACE_SERVICE_CD, ENCNTR_TYPE WHERE RESOURCE_NAME like 'vstrType%'"
    Set rsStatus = cnStatus.Execute(strSQL)
    With rsStatus
       If Not rsStatus.EOF Then
          If (!CPT_Code) <> vbNullString Then udtEnctr.Intrvn_CD = (!CPT_Code)
          If (!CPT_Name) <> vbNullString Then udtEnctr.VisitType = (!CPT_Name)
          If (!Service_Place) <> vbNullString Then udtEnctr.ServicePlace = (!Service_Place)
          If (!Place_Service_CD) <> vbNullString Then udtEnctr.ServicePlaceCode = (!Place_Service_CD)
          If (!Encntr_Type) <> vbNullString Then udtEnctr.EncounterType = (!Place_Service_CD)
          If (!CPT_Desc) <> vbNullString Then udtEnctr.VisitReason = (!CPT_Desc)
       End If
    End With
    If udtEnctr.ServicePlaceCode = "9" Then udtEnctr.ServicePlaceCode = "09"
    cnStatus.Close
    If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
    Get_SurgFacility
    strSQL = "Select * from ENCOUNTER WHERE patient_id = " & Val(udtSchedPatient.IntId) & " AND PURGE = 0  ORDER BY VISIT_TYPE"
    Set rsStatus = cnStatus.Execute(strSQL)
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    With rsStatus
        Do Until .EOF
            If (!Visit_Type) = vstrVisitType Then
                If (!ENCNTR_START_DATE) <> vbNullString Then strEnctTime = (!ENCNTR_START_DATE)
                    If strEnctTime <> vbNullString Then
                    strEnctTime = Format(strEnctTime, "mm/dd/yyyy")
                    lngDateDiff = DateDiff("d", strEnctTime, Format(sDate, "mm/dd/yyyy"))
                    If lngDateDiff = 0 Then
                        udtEnctr.EnctrID = (!Encounter_ID)
                        Exit Do
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    If udtEnctr.EnctrID = 0 Then
        If piProviderID > 0 Then udtEnctr.ProviderID = piProviderID
        udtEnctr.StartDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
        sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
        strSQL = "Insert into ENCOUNTER (PATIENT_ID, PROVIDER_ID, FACILITY_ID, VISIT_TYPE, SERVICE_PLACE," _
             & "ENCNTR_TYPE, ENCNTR_START_DATE, RENDERING_PROVIDER,  INSERT_DATE, SERVICE_PLACE_CD, USER_ID, PURGE, ENCOUNT_CPT_INTRVN) " _
             & "values  (" & Val(udtSchedPatient.IntId) & "," & udtEnctr.ProviderID & "," & udtFacility.FacID & ",'" _
             & udtEnctr.VisitType & "','" _
             & udtEnctr.ServicePlace & "','" _
             & udtEnctr.EncounterType & "','" _
             & udtEnctr.StartDate & "'," & udtEnctr.ProviderID & ",'" _
             & sDate & "'," & Val(udtEnctr.ServicePlaceCode) & ",'" & PsUserID & "',0, '" & clsEncounter.Intrvn_CD & "')"
        Set rsStatus = cnStatus.Execute(strSQL)
        strSQL = "SELECT ENCOUNTER_ID FROM ENCOUNTER WHERE PATIENT_ID = " & Val(udtSchedPatient.IntId) & " AND " _
            & "VISIT_TYPE = '" & udtEnctr.VisitType & "' AND ENCNTR_START_DATE = '" & udtEnctr.StartDate & "'"
        Set rsStatus = cnStatus.Execute(strSQL)
        If Not rsStatus.EOF Then
             udtEnctr.EnctrID = rsStatus.Fields("ENCOUNTER_ID")
             modSpecific.piEncounterID = udtEnctr.EnctrID
             Get_Encounter udtEnctr.EnctrID
        End If
    Else
        strSQL = "UPDATE ENCOUNTER SET RENDERING_PROVIDER = " & udtEnctr.ProviderID & ", ENCNTR_START_DATE = '" & udtEnctr.StartDate & "', UPDATE_DATE = '" & sDate & "' WHERE ENCOUNTER_ID = " & udtEnctr.EnctrID
        Set rsStatus = cnStatus.Execute(strSQL)
    End If

    If cnStatus.State = adStateOpen Then
        cnStatus.Close
        Set cnStatus = Nothing
        Set rsStatus = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_SurgeryStatus:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_SurgeryStatus", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Write_Color(vstrColorName As String, vlngColor As Long)
Dim cnColor As New ADODB.Connection
Dim rsColor As New ADODB.Recordset
Dim strSQL As String
Dim lngPrimKey As Long

    On Error GoTo Error_Write_Color
    
    If cnColor.State = adStateClosed Then cnColor.open psConnect(1)
    strSQL = "Get_Appt_Color_Type_Color '" & vstrColorName & "', " & vlngColor & "," & udtAppointment.ProviderID
    Set rsColor = cnColor.Execute(strSQL)
    If Not rsColor.EOF Then
        lngPrimKey = rsColor.Fields("COLOR_ID")
    End If
    If lngPrimKey > 0 Then
        strSQL = "Update_Appt_Color_Type_Color '" & vstrColorName & "', " & vlngColor & "," & lngPrimKey
         Set rsColor = cnColor.Execute(strSQL)
    Else
        strSQL = "Write_Appt_Color_Type_Color '" & vstrColorName & "'," & vlngColor & ", " & udtAppointment.ProviderID
        Set rsColor = cnColor.Execute(strSQL)
    End If
    If cnColor.State = adStateOpen Then
        cnColor.Close
        Set cnColor = Nothing
        Set rsColor = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Color:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Color", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cmdCancel_Click()
    fraSchLegend.Visible = False
End Sub

Private Sub cmdCloseFac_Click()
    fraFacility.Visible = False
    Exit Sub
End Sub
Private Function Format_Changes()
Dim lngCol As Long
Dim lngRow As Long


    On Error GoTo Error_Format_Changes
    With grdChanges
        .Redraw = False
        .ClearRange 0, SpreadHeader, grdChanges.MaxCols, grdChanges.MaxRows, False
        .ProcessTab = True
        .ColHeaderRows = 4
        .MaxCols = 2
        .MaxRows = 0
        .ColHeadersShow = True
        .RowHeadersShow = False
        .ColWidth(1) = 15 * mlngSystemFileSize
        .ColWidth(2) = 15 * mlngSystemFileSize
    End With
    Format_CellTypeEdit grdChanges, SpreadHeader, 1, "4", "SCHEDULING CHANGES FOR" & vbNewLine & clsProvider.Name, "TypeHAlignCenter", "", HeaderColor
    grdChanges.SetCellBorder 1, SpreadHeader, 3, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 1, grdChanges
    lngCol = 1
    Format_CellTypeEdit grdChanges, SpreadHeader + 2, lngCol, "1", "PATIENT", "TypeHAlignCenter", "", HeaderColor
    grdChanges.SetCellBorder lngCol, SpreadHeader, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 2
    Format_CellTypeEdit grdChanges, SpreadHeader + 2, lngCol, "1", "STATUS", "TypeHAlignCenter", "", HeaderColor
    grdChanges.SetCellBorder lngCol, SpreadHeader, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 3, grdChanges
    Exit Function
Error_Format_Changes:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", " Remove_Data ", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Sub cmdCloseList_Click()
    fraProvider.Visible = False
    Exit Sub
End Sub

Private Sub cmdCurrentWeek_Click()
    mstrFirstDay = vbNullString
    mBeenHere = True
    If mlngClickPrev > 0 Or mlngClickNext > 0 Then
         udtAppointment.SchedDate = Format(Now, "mm/dd/yyyy")
    End If
    Format_Week Format(MonthView1.Value, "mm/dd/yyyy")
    Read_Week_Appointments
    grdData.Redraw = True
    mlngClickPrev = 0
    mlngClickNext = 0
    Exit Sub
End Sub

Private Sub cmdFacility_Click()
    fraFacility.ZOrder 0
    fraFacility.Visible = True
    Exit Sub
End Sub


Private Sub cmdNextWeek_Click()
Dim Arr
Dim strDate As String
Dim nDays As Long
Dim lngCol As Long
Dim lngRow As Long
Dim vDay
Dim lngAddDay As Long
Dim vFirstDay As Variant
Dim lngAddWeek As Long
Dim strOption As String
Dim lngWeekDays As Long

    On Error GoTo Error_NextWeek
    Screen.MousePointer = vbHourglass
    mlngClickNext = mlngClickNext + 1
    grdData.GetText 1, SpreadHeader + 2, vFirstDay
    If vFirstDay <> vbNullString Then
        Arr = Split(vFirstDay, "-")
        If UBound(Arr) > 0 Then
            vFirstDay = Arr(1)
        End If
    End If
    lngAddWeek = 7
    mChange = True
    If vFirstDay = vbNullString Then
        strDate = udtAppointment.SchedDate
    Else
        strDate = vFirstDay
    End If
    nDays = 1 * lngAddWeek
    If strDate <> vbNullString Then
        strDate = DateAdd("d", nDays, strDate)
        vDay = strDate
        With grdData
            Format_Week Format(vDay, "mm/dd/yyyy")
            Read_Week_Appointments
        End With
    End If
    grdData.Redraw = True
    mChange = False
    'mbooTimer = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_NextWeek:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "NextWeek", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Sub cmdPrevWeek_Click()
Dim Arr
Dim strDate As String
Dim Idx As Integer
Dim strSQL As String
Dim sDate As String
Dim vDay As String
Dim sMonth As String
Dim nDays As Long
Dim lngAddDay As Long
Dim vFirstDay As Variant
Dim lngAddWeek As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strOption As String
Dim lngWeekDays As Long

    On Error GoTo Error_PrevWeek
    Screen.MousePointer = vbHourglass
    If IsDate(vFirstDay) Then
        mlngClickNext = 0
        mlngClickPrev = mlngClickPrev + 1
        lngAddWeek = 7
        mChange = True
        If vFirstDay = vbNullString Then
            strDate = udtAppointment.SchedDate
        Else
            strDate = vFirstDay
        End If
        nDays = -1 * lngAddWeek
        strDate = DateAdd("d", nDays, strDate)
        vDay = strDate
    Else
        MsgBox "Please select week view first!", vbInformation, "PatientTrac Appointment Schedule"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    rtfChanges.Text = vbNullString
    grdData.Redraw = False
    Format_Week Format(vDay, "mm/dd/yyyy")
    Read_Week_Appointments
    grdData.Redraw = True
    mChange = False
    Screen.MousePointer = vbDefault
    'mbooTimer = True
    Exit Sub
Error_PrevWeek:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "PrevWeek", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub



Private Sub cmdProvider_Click()
    fraProvider.ZOrder 0
    fraProvider.Visible = True
    Exit Sub
End Sub

Private Sub cmdRepair_Click()
Dim i%
Dim strSQL As String
Dim cnApptCheck As New ADODB.Connection
Dim rsApptCheck As New ADODB.Recordset
Dim strApptTime As String
Dim strApptDate As String
Dim lngProviderID As String
Dim lngDuplicates As Long
Dim lngSelProv As Long

    ' For Code Mode Only Repairs
    cnApptCheck.open psConnect(0)
    strSQL = "select patient_id from appointments where Active = 1 order by patient_id"
    Set rsApptCheck = cnApptCheck.Execute(strSQL)
    With rsApptCheck
        Do Until .EOF
           If Not IsNull(!Patient_ID) And (!Patient_ID) <> 0 Then
                lngSelProv = (!Patient_ID)
                strSQL = "Select Appt_Time, Appt_Date, Provider_Id from appointments where " _
                    & "patient_id = " & lngSelProv & "  Order by appt_Date"
                cnApptCheck.open psConnect(0)
                Set rsApptCheck = cnApptCheck.Execute(strSQL)
                With rsApptCheck
                    Do Until rsApptCheck.EOF
                        If (!Appt_Time) <> vbNullString And Trim(!Appt_Time) = strApptTime Then
                            If (!Appt_Date) <> vbNullString And Trim(!Appt_Date) = strApptDate Then
                                If (!Provider_ID) <> 0 And (!Provider_ID) = lngProviderID Then
                                    ' We have a duplicate patient appointment
                                    lngDuplicates = lngDuplicates + 1
                                    MsgBox "Duplicates" & lngDuplicates, vbInformation
                                End If
                            End If
                        Else
                            If (!Appt_Time) <> vbNullString Then strApptTime = Trim(!Appt_Time)
                            If (!Appt_Date) <> vbNullString Then strApptDate = Trim(!Appt_Date)
                            If (!Provider_ID) <> 0 Then lngProviderID = (!Provider_ID)
                        End If
                    rsApptCheck.MoveNext
                    Loop
                End With
           End If
        .MoveNext
        Loop
    End With
    Close_DB psConnect(0)
End Sub


Private Sub cmdStop_Click()

    If cmdStop.Caption = "Stop Refresh" Then
        schTimer.Enabled = False
        cmdStop.Caption = "Refresh Appt"
    Else
        schTimer.Enabled = True
        cmdStop.Caption = "Stop Refresh"
    End If
    Exit Sub
End Sub



Private Sub Form_Activate()
    mFirst = True
End Sub

Private Sub Form_Load()
    mFirst = False
    mbooByPass = False
    Prepare_Grid
    Set_Headers
    Me.ZOrder 0
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    mFirstLoad = False
    schTimer.Enabled = False
    If IsFormLoaded("frmInPatient") Then
        frmInPatient.ZOrder 0
    End If
End Sub

Private Sub Form_Resize()
    
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)


    Unload Me
    Exit Sub
End Sub

Private Function Reset_SelProvider()
Dim lngProvID As Long
Dim strReason As String
Dim i%

'    For i = 0 To cboProviders.listcount - 1
'        mboobyppas = True
'        If cboProviders.ItemData(i) = udtAppointment.ProviderID Then
'            mbooByPass = True
'            cboProviders.Text = cboProviders.List(i)
'             mboobyppas = True
'            udtAppointment.ProviderID = cboProviders.ItemData(i)
'            mlngCurProvID = udtAppointment.ProviderID
'            Exit For
'        End If
'    Next i
    Exit Function
End Function

Private Sub grdChanges_Click(ByVal Col As Long, ByVal Row As Long)
    If Row <= 0 Then Exit Sub
End Sub


Private Sub grdChanges_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim bFound As Boolean
Dim Arr
Dim i%

    On Error GoTo Error_grdChanges_DblClick
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    'mbooTimer = False
    udtAppointment.SchedTime = vbNullString
    grdData.Redraw = False
    mFirst = False
    udtAppointment.SchedType = vbNullString
    Clear_SchedPatient
    udtSchedPatient.IntId = vbNullString
    'mbooTimer = False
    mlngPatientID = 0
    grdData.Col = 1
    grdData.Row = SpreadHeader + 2
    If grdData.CellTag <> vbNullString Then
       udtAppointment.ProviderID = grdData.CellTag
    End If
    With grdChanges
        .Col = 1
        .Row = Row
         udtSchedPatient.IntId = grdChanges.GetRowItemData(Row)
        If .CellNote <> vbNullString Then
            Arr = Split(.CellNote, vbNewLine)
            If UBound(Arr) > 0 Then
                udtAppointment.SchedType = Arr(0)
            Else
                udtAppointment.SchedType = .CellNote
            End If
        End If
    End With
    If udtAppointment.SchedType = vbNullString Then udtAppointment.SchedType = "Patient"
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
    With grdData
        For lngRow = 1 To grdData.MaxRows
            For lngCol = 1 To .MaxCols
                .Col = lngCol
                .Row = lngRow
                If .CellTag = udtSchedPatient.IntId Then
                    .GetText 0, lngRow, varCellValue
                    udtAppointment.SchedTime = varCellValue
                    bFound = True
                    Exit For
                End If
            Next lngCol
        Next lngRow
    End With
    mlngPatientID = Val(udtSchedPatient.IntId)
    piPatientID = Val(udtSchedPatient.IntId)
    If bFound Then
        With grdData
            .Col = lngCol
            .Row = lngRow
            If .CellTag <> vbNullString Then
                mbooTimeAvailable = True
            Else
                Check_TimeBlocks lngCol, lngRow
            End If
        End With
    End If
    If mbooTimeAvailable Then
        mLoadPatient = True
        mGridClick = True
        Load_SchedPatient Col, Row
'        Me.Hide
'        grdData.Refresh
'        mLoadPatient = True
'        frmSchedPatient.mbLookupPatient = True
'        mGridClick = True
'        InfoFrame Col, Row
'        mGridClick = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mGridClick = True
'    Me.Show
'    Me.ZOrder 0
    schTimer.Enabled = True
    mbooDoubleBook = False
    grdData.Redraw = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_grdChanges_DblClick:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "grdChanges_DblClick", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Sub grdData_Click(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue
Dim strDate As String
Dim Arr

    If mstrGridType = "Weekly" Then
        With grdData
            Format_Changes
            .GetText Col, SpreadHeader + 2, varCellValue
            strDate = varCellValue
            Arr = Split(varCellValue, "-")
            If UBound(Arr) > 0 Then
                Get_Changes Format(Arr(1), "mm/dd/yyyy")
            End If
        End With
    End If
End Sub

Private Sub grdData_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim vDate
Dim strDate As String
Dim Arr
Dim arr1
Dim vPatient
Dim lngPatientID As Long
Dim vTime
Dim sFirstName As String
Dim sLastName As String
Dim i%
Dim strSQL As String
Dim cnData As New ADODB.Connection
Dim rsData As New ADODB.Recordset
Dim lngCurFac As Long
Dim lngTestProv As Long
Dim lngTest As Long

    On Error GoTo Error_GrdData_DblClick
    booAppTime = False
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    schTimer.Enabled = False
    If IsFormLoaded("frmInPatient") Then
        Unload frmInPatient
    End If
    grdData.Col = Col
    grdData.Row = Row
    clsGridFormat.CellTag = vbNullString
    If Trim(grdData.CellTag) = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        clsGridFormat.CellTag = grdData.CellTag
    End If
    If grdData.CellNote <> vbNullString Then
        Arr = Split(grdData.CellNote, vbNewLine)
        If UBound(Arr) > 0 Then
            udtAppointment.SchedType = Arr(0)
        Else
            udtAppointment.SchedType = grdData.CellNote
        End If
        Arr = Split(udtAppointment.SchedType, "|")
        If UBound(Arr) > 0 Then
            udtAppointment.SchedType = Arr(0)
            udtAppointment.Appt_ID = Arr(1)
        End If
    Else
        udtAppointment.SchedType = "Patient"
    End If
    schTimer.Enabled = False
    udtAppointment.SchedTime = vbNullString
    mlngGridCol = Col
    mlngGridRow = Row
    If mstrApptMode = "Surgery" Then
        Manage_Surgery "DblClick"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    mlngPatientID = 0
    piPatientID = 0
    piEncounterID = 0
    modSpecific.piEncounterID = 0
    frmInPatient.piSelEnctrID = 0
    If mstrGridType = vbNullString Then mstrGridType = "Normal"
    With grdProviders
       For lngTest = 1 To .MaxRows
           .Col = 1
           .Row = lngTest
           If .BackColor = grdCurItem Then
                lngTestProv = .GetRowItemData(lngTest)
                Exit For
           End If
        Next lngTest
    End With
    Select Case mstrGridType
        Case "Normal", "Name Only"
             grdData.Col = Col
            grdData.Row = SpreadHeader + 2
            If grdData.CellTag <> vbNullString Then
                udtAppointment.ProviderID = grdData.CellTag
            End If
            grdData.Col = Col
            grdData.Row = Row
            If grdData.CellTag <> vbNullString Then
                If IsNumeric(grdData.CellTag) Then udtSchedPatient.IntId = grdData.CellTag
            End If
            If grdData.CellNote <> vbNullString Then
                Arr = Split(grdData.CellNote, vbNewLine)
                If UBound(Arr) > 0 Then
                    udtAppointment.SchedType = Arr(0)
                Else
                    udtAppointment.SchedType = grdData.CellNote
                End If
                Arr = Split(udtAppointment.SchedType, "|")
                If UBound(Arr) > 0 Then
                    udtAppointment.SchedType = Arr(0)
                    udtAppointment.Appt_ID = Arr(1)
                End If
            End If
            udtAppointment.SchedDate = Format(MonthView1.Value, "mm/dd/yyyy")
    Case "Weekly"
        With grdData
            .GetText Col, SpreadHeader + 2, vDate
            strDate = vDate
            Arr = Split(vDate, "-")
            If UBound(Arr) > 0 Then
                udtAppointment.SchedDate = Format(Arr(1), "mm/dd/yyyy")
            End If
            lngCurFac = udtAppointment.FacilityID
            udtAppointment.FacilityID = .GetColItemData(Col)
            udtFacility.FacID = udtAppointment.FacilityID
            grdData.Col = Col
            grdData.Row = Row
            If grdData.CellTag <> vbNullString Then
                If IsNumeric(grdData.CellTag) Then udtSchedPatient.IntId = grdData.CellTag
            End If
            If grdData.CellNote <> vbNullString Then
                Arr = Split(grdData.CellNote, vbNewLine)
                If UBound(Arr) > 0 Then
                    udtAppointment.SchedType = Arr(0)
                Else
                    udtAppointment.SchedType = grdData.CellNote
                End If
                Arr = Split(udtAppointment.SchedType, "|")
                If UBound(Arr) > 0 Then
                    udtAppointment.SchedType = Arr(0)
                    udtAppointment.Appt_ID = Arr(1)
                End If
            End If
        End With
    End Select
    If Val(clsGridFormat.CellTag) > 0 Then
        If udtAppointment.ProviderID <> mlngLogonUser Then
                Clear_PracticeRole
                Get_SecurityRole mlngLogonUser
                If Get_SuperProv Then
                    If lngTestProv <> piProviderID Then
                        MsgBox "Practice rules preclude reviewing other provider's schedule, Please contact Administration for other providers schedule.!", vbInformation, "PatientTrac Schedule"
                        udtAppointment.ProviderID = piProviderID
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                ElseIf UCase(clsPracticeRoles.Permission) <> "ALL" Then
                    MsgBox "This is not your Patient Encounter use Patient" _
                        & "History to review this visit data. Thank you!", vbInformation, "PatientTrac Security Management"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
        End If
    End If
    grdData.GetText 0, Row, vTime
    udtAppointment.SchedTime = vbNullString
    udtAppointment.SchedTime = CStr(vTime)
    If udtAppointment.SchedType = vbNullString Then udtAppointment.SchedType = "Patient"
    If IsProperCase(udtAppointment.SchedType) = "Patient" And Val(clsGridFormat.CellTag) > 0 Then
        udtAppointment.SchedTime = CStr(vTime)
        If Check_Patient(Val(clsGridFormat.CellTag), udtAppointment.SchedTime, udtAppointment.SchedDate) Then
            mbooPatValid = False
            mlngPatientID = Val(clsGridFormat.CellTag)
            piPatientID = Val(clsGridFormat.CellTag)
            modSpecific.piPatientID = piPatientID
            frmPatients.piPatientID = piPatientID
            Get_CurrentEnctr udtAppointment.SchedDate
            If mbooPatValid Then
                Get_Patient piPatientID
                modSpecific.piEncounterID = piEncounterID
            End If
        Else
            mlngPatientID = Val(clsGridFormat.CellTag)
            piPatientID = Val(clsGridFormat.CellTag)
            modSpecific.piPatientID = piPatientID
            frmPatients.piPatientID = piPatientID
            Get_CurrentEnctr udtAppointment.SchedDate
            If IsNumeric(mlngPatientID) Then
                Get_Patient piPatientID
                modSpecific.piEncounterID = piEncounterID
                mbooPatValid = True
            Else
                mbooPatValid = False
            End If
        End If
    Else
        mbooPatValid = False
    End If
    mGridClick = True
    If cnData.State = adStateOpen Then
        cnData.Close
        Set cnData = Nothing
        Set rsData = Nothing
    End If
    If mbooPatValid Then
        frmMDI.StatusBar.Panels(1).Text = "Patient: " & clsPatient.Name
        If mstrPractice = "Psychiatry" Then
            frmMDI.StatusBar.Panels(2).Text = "Hospital ID: " & clsPatient.ExtID
        Else
            frmMDI.StatusBar.Panels(2).Text = "Financial ID: " & clsPatient.ExtID
        End If
        Get_Encounter piEncounterID
        frmMDI.StatusBar.Panels(3).Text = "Patient EMR ID: " & piPatientID
        frmMDI.StatusBar.Panels(4).Text = "Encounter Number: " & piEncounterID
        frmMDI.StatusBar.Panels(5).Text = "Visit Date: " & Format(clsEncounter.StartDate, "MMMM DD,YYYY")
        frmMDI.StatusBar.Panels(6).Text = "Provider: " & clsEncounter.Provider
        If piEncounterID > 0 Then
            Load_Enctr_Format
        End If
    End If
    Screen.MousePointer = vbDefault
    Get_Mode
    If Me.WindowState = 1 Then
        Me.WindowState = 0
    End If
    booAppTime = True
    Exit Sub
Error_GrdData_DblClick:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "GrdData_DblClick", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub


Private Sub grdData_DragDropBlock(ByVal Col As Long, ByVal Row As Long, ByVal Col2 As Long, ByVal Row2 As Long, ByVal NewCol As Long, ByVal NewRow As Long, ByVal NewCol2 As Long, ByVal NewRow2 As Long, ByVal Overwrite As Boolean, Action As Integer, DataOnly As Boolean, Cancel As Boolean)
'Dim X As Long
'Dim Y As Long
'Dim iAnswer As Integer
'Dim strSQL As String
'Dim lngPatientID As Long
'Dim vTime
'Dim vPrevTm
'Dim vOldD
'Dim vDate
'Dim lngProviderId As Long
'Dim lngFacilityID As Long
'Dim i%
'Dim vExistingData
'Dim lngExistPat As Long
'Dim lngMovePatRow As Long
'Dim varHeader
'Dim varPatData
'Dim booinsert  As Boolean
'
'    With grdData
'        If mstrApptMode = "Surgery" Then
'            'My notes :  Get The existing Column Name if the column name not Patient Surgery then update the Encounter set Encounter End Date = this date
'            .GetText Col, SpreadHeader + 3, varHeader
'            If TextFound("OPERATING ROOM", varHeader, Len("OPERATING ROOM"), True) Then
'                booinsert = False
'            Else
'                booinsert = True
'            End If
'            .GetText NewCol, NewRow, vExistingData
'            .GetText Col, Row, varPatData
'            If vExistingData <> vbNullString And booinsert Then
'                If vExistingData <> vbNullString Then
'                    .InsertRows NewRow + 1, 1
'                    lngMovePatRow = NewRow + 1
'                    .MaxRows = .MaxRows + 1
'                    lngExistPat = .GetRowItemData(Row)
'                    .SetText Col, NewRow, vbNullString
'                    .SetText Col, lngMovePatRow, vExistingData
'                    .SetRowItemData lngMovePatRow, lngExistPat
'                    Format_CellTypeStatic grdData, lngMovePatRow, NewCol, "1", CStr(varPatData), "TypeHAlignLeft", CStr(varHeader), "", grdCurItem
'                Else
'                    Format_CellTypeStatic grdData, NewRow, NewCol, "1", CStr(varPatData), "TypeHAlignLeft", CStr(varHeader), "", grdCurItem
'                End If
'                Update_SurgeryStatus CStr(varHeader)
'            Else
'                MsgBox "This resource is in use, please select another Resource!", vbInformation
'                Exit Sub
'            End If
'        Else
'             Validate_Appointment
'             If mstrValidAppoint = "Invalid Appointment" Then
'                Cancel = True
'                Exit Sub
'             End If
'            With grdData
'                .GetText NewCol, NewRow, vExistingData
'                If vExistingData <> vbNullString Then
'                    .InsertRows NewRow + 1, 1
'                    lngMovePatRow = NewRow + 1
'                    .MaxRows = .MaxRows + 1
'                    lngExistPat = .GetRowItemData(NewRow)
'                    .SetText Col, NewRow, vbNullString
'                    .SetText Col, lngMovePatRow, vExistingData
'                    .SetRowItemData lngMovePatRow, lngExistPat
'                End If
'            End With
'
'            Open_DB psConnect(0)
'            For i = 0 To cboLocation.ListCount - 1
'                If cboLocation.List(i) = cboLocation.Text Then
'                    lngFacilityID = cboLocation.ItemData(i)
'                End If
'            Next i
'
'            With grdData
'                If NewRow <> Row Then
'                    iAnswer = MsgBox("Do you want to save changes?", vbYesNoCancel)
'                    If iAnswer = vbYes Then
'                        If mstrApptMode = "Day" Then
'                            .GetText 0, NewRow, vTime
'                            If vTime = vbNullString Then .GetText 0, NewRow - 1, vTime
'                            .GetText 0, Row, vPrevTm
'                            If vTime = vbNullString Then .GetText 0, Row - 1, vPrevTm
'                            lngPatientID = .GetRowItemData(Row)
'                            .SetRowItemData NewRow, lngPatientID
'                            lngProviderId = .GetColItemData(Col)
'                            strSQL = "Update Appointments set appt_Time = '" & vTime & "', appt_date = '" & udtAppointment.SchedDate & "', provider_id = " & lngProviderId & "  " _
'                                & "WHERE patient_id = " & lngPatientID & " And facility_id = " & lngFacilityID & " and appt_time = '" & vPrevTm & "'"
'                            Set rs = cn.Execute(strSQL)
'                        Else
'                            .GetText 0, NewRow, vTime
'                            .GetText 0, NewRow, vPrevTm
'                            lngPatientID = .GetRowItemData(Row)
'                            For i = 0 To cboProviders.ListCount - 1
'                                If cboProviders.List(i) = cboProviders.Text Then
'                                    lngProviderId = cboProviders.ItemData(i)
'                                End If
'                            Next i
'                            .GetText NewCol, 0, vDate
'                            .GetText Col, 0, vOldD
'                            lngPatientID = .GetRowItemData(Row)
'                            strSQL = "Update Appointments set appt_Time = '" & vTime & "', appt_date = '" & vDate & "', provider_id = " & lngProviderId & "" _
'                                & "WHERE patient_id = " & lngPatientID & " And facility_id = " & lngFacilityID & " and appt_time = '" & vPrevTm & "' And " _
'                                & "Appt_Date = '" & vOldD & "'"
'
'                            Set rs = cn.Execute(strSQL)
'                        End If
'                    Else
'                        Cancel = True
'                    End If
'                End If
'            End With
'            Close_DB psConnect(0)
'        End If
'    End With
'    Exit Sub
End Sub

Private Sub grdData_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
Dim vDate
Dim strDate As String
Dim Arr
Dim strTime As String
Dim varCellValue
Dim booValidAppt As Boolean
Dim Idx%
Dim lngRow As Long
Dim lngCol As Long

    On Error GoTo Error_New_Schedule
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    schTimer.Enabled = False
    Clear_SchedPatient
    udtSchedPatient.IntId = vbNullString
    If udtAppointment.ProviderID <> mlngLogonUser Then
        Clear_PracticeRole
        Get_SecurityRole mlngLogonUser
        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
            MsgBox "Unable to Schedule another Provider Appointment, please schedule through Administration!", vbInformation, "PatientTrac Security Management"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    With grdData
        .Redraw = False
        .GetText 0, Row, varCellValue
        If Trim(varCellValue) <> vbNullString Then strTime = Trim(varCellValue)
        If mstrApptMode <> "Day" Then
            grdData.Col = Col
            grdData.Row = SpreadHeader + 3
            If grdData.CellTag <> vbNullString Then
                udtAppointment.ProviderID = grdData.CellTag
            End If
        Else
            With grdProviders
                For lngRow = 1 To .MaxRows
                    .GetText 1, lngRow, varCellValue
                    If varCellValue = cctProvider Then
                        .BlockMode = True
                            .Col = 1
                            .Col2 = 2
                            .Row = lngRow
                            .Row2 = lngRow
                            .BackColor = grdCurItem
                        .BlockMode = False
                        udtAppointment.ProviderID = .GetRowItemData(lngRow)
                        Exit For
                    End If
                Next lngRow
            End With
        End If
        mbooByPass = False
        udtAppointment.SchedTime = vbNullString
        If mstrApptMode = "Day" Then
            udtAppointment.SchedDate = Format(MonthView1.Value, "mm/dd/yyyy")
        ElseIf mstrApptMode = "Weekly" Then
            .GetText 0, Row, vDate
            strDate = vDate
            Arr = Split(strDate, "-")
            If UBound(Arr) > 0 Then
                udtAppointment.SchedDate = Format(Arr(1), "mm/dd/yyyy")
            Else
                udtAppointment.SchedDate = strDate
            End If
        End If
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then
            If IsNumeric(.CellTag) Then
                udtSchedPatient.IntId = .CellTag
            End If
            If .CellNote <> vbNullString Then
                Arr = Split(.CellNote, vbNewLine)
                If UBound(Arr) > 0 Then
                    udtAppointment.SchedType = Arr(0)
                Else
                    udtAppointment.SchedType = .CellNote
                End If
                Arr = Split(udtAppointment.SchedType, "|")
                If UBound(Arr) > 0 Then
                    udtAppointment.SchedType = Arr(0)
                    udtAppointment.Appt_ID = Arr(1)
                End If
            End If
        End If
    End With
    mLoadPatient = True
    mGridClick = True
    Load_SchedPatient Col, Row
    If Me.WindowState = 1 Then
        Me.WindowState = 0
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_New_Schedule:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "New_Schedule", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub Read_Messages()
Dim i%
Dim strSQL As String
Dim cnMessage As New ADODB.Connection
Dim rsMessage As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strDate As String
Dim lngCol As Long
Dim lngRow As Long
Dim booAddRow As Boolean
Dim strSymptoms As String
Dim lngMesID As Long
Dim strUrgent As String

    On Error GoTo Error_Read_Messages
    Screen.MousePointer = vbHourglass
    grdMessages.ClearRange 1, 0, grdMessages.MaxCols, grdMessages.MaxRows, True
    Format_Messages
    If grdMessages.MaxRows <= 1 Then
        Format_Messages
    End If
    If mlngLogonUser <= 0 Then PsUserID = Get_PsUserID(psConnect(0))
    If cnMessage.State = adStateClosed Then cnMessage.open psConnect(0)
    strSQL = "Read_Message_Provider " & mlngLogonUser
    Set rsMessage = cnMessage.Execute(strSQL)
    With rsMessage
        Do Until rsMessage.EOF
            booAddRow = False
            Clear_Phone
            lngMesID = 0
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "CALLER_ID"
                                udtTelephone.CallerID = Fld.Value
                            Case "MESSAGE_TIME"
                                udtTelephone.Time = Fld.Value
                            Case "SENDER_ID"
                                udtTelephone.SenderID = Fld.Value
                            Case "RECIEVOR_ID"
                                udtTelephone.ReceiverID = Fld.Value
                            Case "MEDICATION_STATUS"
                                udtTelephone.MedStatus = Fld.Value
                            Case "PATIENT_SYMPTOMS"
                                udtTelephone.Symptoms = Fld.Value
                            Case "MESSAGE_NOTE"
                                udtTelephone.MessNote = Fld.Value
                            Case "CONTACT_TYPE"
                                udtTelephone.ContactType = Fld.Value
                            Case "MESSAGE_ID"
                                lngMesID = Fld.Value
                            Case "CALLER_NAME"
                                udtTelephone.CallerName = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtTelephone.CallerID = 0 Then
                booAddRow = False
            End If
            If Not booAddRow Then
                If DateDiff("d", "01/01/1900", Format(udtTelephone.Time, "mm/dd/yyyy")) > 0 Then
                    If udtTelephone.CallerID > 0 And udtTelephone.SenderID = mlngLogonUser Or _
                        udtTelephone.ReceiverID = mlngLogonUser Then
                        booAddRow = True
                    Else
                        booAddRow = False
                    End If
                Else
                    booAddRow = False
                End If
            End If
            If booAddRow Then
                mlngMessRow = mlngMessRow + 1
                grdMessages.Row = mlngMessRow
                grdMessages.MaxRows = mlngMessRow
                Format_CellTypeStatic grdMessages, mlngMessRow, 1, "1", Format(udtTelephone.Time, "mm/dd/yyyy"), "TypeHAlignLeft", "MESSAGE_TIME"
                grdMessages.SetRowItemData mlngMessRow, lngMesID
                If udtTelephone.CallerName <> vbNullString Then
                    Format_CellTypeStatic grdMessages, mlngMessRow, 2, "1", udtTelephone.CallerName, "TypeHAlignLeft", "CALLER_NAME"
                End If
                If udtTelephone.MedStatus <> vbNullString Then
                    Format_CellTypeStatic grdMessages, mlngMessRow, 3, "1", udtTelephone.MedStatus, "TypeHAlignLeft", "MEDICATION_STATUS"
                End If
                If udtTelephone.Symptoms <> vbNullString Then
                   If TextFound("Urgent", udtTelephone.Symptoms, Len("Urgent"), True) Then
                        strUrgent = "Urgent"
                        udtTelephone.Symptoms = Replace(udtTelephone.Symptoms, "Urgent", "")
                        Format_CellTypeStatic grdMessages, mlngMessRow, 4, "1", udtTelephone.Symptoms, "TypeHAlignLeft", "PATIENT_SYMPTOMS"
                   Else
                        Format_CellTypeStatic grdMessages, mlngMessRow, 4, "1", udtTelephone.Symptoms, "TypeHAlignLeft", "PATIENT_SYMPTOMS"
                   End If
                End If
                If Trim(udtTelephone.MessNote) <> vbNullString Then
                    grdMessages.SetText 5, mlngMessRow, udtTelephone.MessNote
                    Format_CellTypeStatic grdMessages, mlngMessRow, 5, "1", udtTelephone.MessNote, "TypeHAlignLeft", "MESSAGE_NOTE"
                End If
                If Trim(udtTelephone.CallType) <> vbNullString And Trim(udtTelephone.CallType) = "Inquiry" Then
                    grdMessages.BlockMode = True
                        grdMessages.Col = 1
                        grdMessages.Col2 = grdMessages.MaxCols
                        grdMessages.Row = mlngMessRow
                        grdMessages.Row2 = mlngMessRow
                        grdMessages.BackColor = &HC0FFC0
                    grdMessages.BlockMode = False
                    grdMessages.Col = 1
                    grdMessages.Row = mlngMessRow
                    grdMessages.CellTag = "Inquiry"
                End If
                If strUrgent = "Urgent" Then
                    grdMessages.BlockMode = True
                        grdMessages.Col = 1
                        grdMessages.Col2 = grdMessages.MaxCols
                        grdMessages.Row = mlngMessRow
                        grdMessages.Row2 = mlngMessRow
                        grdMessages.BackColor = YellowColor
                    grdMessages.BlockMode = False
                End If
                grdMessages.SetCellBorder 1, mlngMessRow, 1, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdMessages.SetCellBorder 2, mlngMessRow, 2, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdMessages.SetCellBorder 3, mlngMessRow, 3, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdMessages.SetCellBorder 4, mlngMessRow, 4, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                grdMessages.SetCellBorder 5, mlngMessRow, 5, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
            If booAddRow Then
                grdMessages.SetCellBorder 1, mlngMessRow, grdMessages.MaxCols, mlngMessRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        Loop
    End With
    If cnMessage.State = adStateOpen Then
        cnMessage.Close
        Set cnMessage = Nothing
        Set rsMessage = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Read_Messages:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Messages", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub grdFacility_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim varCellValue

    If Row > 0 Then
        With grdFacility
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
        End With
        grdData.SetText 1, 0, " "
        udtFacility.FacID = grdFacility.GetRowItemData(Row)
        udtAppointment.FacilityID = udtFacility.FacID
        piFacilityID = udtFacility.FacID
        Get_Appt_Providers
        grdFacility.GetText 1, Row, varCellValue
        cctFacilities.Text = varCellValue
        Set_Headers
        Get_Mode
        Screen.MousePointer = vbDefault
        fraFacility.Visible = False
    End If
End Sub


Private Sub grdMessages_Click(ByVal Col As Long, ByVal Row As Long)
Dim iRow%
    mlngMessRow = Row
    mlngMessCol = Col
     With grdMessages
        .Redraw = False
        If mlngMessRow >= 1 Then
            For iRow = 1 To .MaxRows
                .BlockMode = True
                    .Col = 0
                    .Col2 = .MaxCols
                    .Row = iRow
                    .Row2 = iRow
                    .FontBold = True
                    .BackColor = &HFFFFFF
                .BlockMode = False
            Next iRow
        End If
        .Redraw = True
    End With
    Exit Sub
End Sub


Private Sub grdMessages_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim iRow%
Dim booNewContact As Boolean
   mlngMessRow = Row
   mlngMessCol = Col
   booNewContact = False
   If mlngMessRow < 1 Then Exit Sub
   mlngMessageID = 0
   With grdMessages
        .Redraw = False
        mlngMessageID = 0
        mlngMessageID = .GetRowItemData(mlngMessRow)
            If mlngGridRow >= 1 Then
                For iRow = 1 To .MaxRows
                    .BlockMode = True
                        .Col = 0
                        .Col2 = .MaxCols
                        .Row = iRow
                        .Row2 = iRow
                        .FontBold = True
                        .BackColor = &HFFFFFF
                    .BlockMode = False
                Next iRow
            End If
            For iRow = 1 To .MaxRows
                .BlockMode = True
                    .Col = 0
                    .Col2 = .MaxCols
                    .Row = mlngMessRow
                    .Row2 = mlngMessRow
                    .FontBold = True
                    .BackColor = &HFFFF&
                .BlockMode = False
                Exit For
            Next iRow
        .Col = 1
        .Row = mlngMessRow
        If .CellTag = "Inquiry" Then
            booNewContact = True
            mstrPatContact = "Contact"
        Else
            booNewContact = True
            mstrPatContact = "Phone"
        End If
        If mlngMessageID <> 0 Then
            frmTelephone.Show vbModal
            Format_Messages
            Read_Messages
        End If
        .Redraw = True
   End With
    Exit Sub
End Sub

Private Sub grdProviders_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim varCellValue
Dim lngProv As Long

    If Row > 0 Then
        With grdProviders
            lngProv = .GetRowItemData(Row)
            If lngProv <> mlngLogonUser And clsEDI_Client.Site_ID <> "1000" Then
                Get_Indiv_Priv "Administrator", mlngLogonUser
                If IsProperCase(clsPracticeRoles.RoleTitle) = "Administrator" Then
                    
                ElseIf clsEDI_Client.Site_ID = "2777" And Get_SuperProv Then
                    If piProviderID <> lngProv Then
                        MsgBox "Facility rules do not allow access to other providers'schedule." _
                            & " " & "Please contact your administrator for information regarding other providers' schedule!", vbInformation, "PatientTrac Schedule"
                        udtAppointment.ProviderID = piProviderID
                        Exit Sub
                    End If
                Else
                    MsgBox "Facility rules do not allow access to other providers'schedule." _
                            & " " & "Please contact your administrator for information regarding other providers' schedule!", vbInformation, "PatientTrac Schedule"
                    udtAppointment.ProviderID = mlngLogonUser
                    Exit Sub
                End If
            End If
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
            udtAppointment.ProviderID = .GetRowItemData(Row)
            .GetText 1, Row, varCellValue
            If cctProvider.Text = vbNullString Then cctProvider.Text = Trim(varCellValue)
            Change_Providers
            
            fraProvider.Visible = False
        End With
    End If
    Exit Sub
End Sub


Private Sub mnu_Block_Time_Click()
    schTimer.Enabled = False
    frmBlockTime.Show vbModal
    Get_Mode
    schTimer.Enabled = True
    Exit Sub
End Sub

Private Sub mnu_Day_Click()
    lblApptDate.Visible = True
    cmdPrevWeek.Visible = False
    cmdCurrentWeek.Visible = False
    cmdNextWeek.Visible = False
    mstrApptMode = "Day"
    grdData.MaxCols = 1
    Set_Headers
    Get_Mode
    Exit Sub
End Sub


Private Sub mnu_Message_Click()
    mstrPatContact = "Phone"
    frmTelephone.Show vbModal
    Format_Messages
    Read_Messages
End Sub

Private Sub mnu_Month_Click()
    mstrApptMode = "Month"
    Get_Mode
    cmdPrevWeek.Visible = False
    cmdCurrentWeek.Visible = False
    cmdNextWeek.Visible = False
    lblProvider.Visible = False
    Exit Sub
End Sub


Private Sub mnu_New_Patient_Click()
    On Error GoTo Error_New_Patient
    Screen.MousePointer = vbHourglass
    modSpecific.piPatientID = 0
    modSpecific.piEncounterID = 0
    piPatientID = 0
    piEncounterID = 0
    modSpecific.piProviderID = piProviderID
    Dump_PatientStrData
    PatientInfo.mNewPatient = True
    PatientInfo.Show vbModal
    Screen.MousePointer = vbDefault
    Exit Sub
Error_New_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "New_Patient", "PatientTracEMR.FrmPatient"
    Exit Sub
    Resume Next
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
    CD_ShowPrinter (Me.Hwnd)
      If mbooCancelPrint Then
            MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Appointment Scheduling"
            mbooCancelPrint = False
            Exit Sub
      End If
      With grdData
          .PrintSmartPrint = False
          .PrintAbortMsg = "Printing - Click Cancel to quit"
          .PrintJobName = "Scheudule"
          .PrintHeader = strHeaders
          .PrintFooter = "/cPage #/p/n"
          .PrintBorder = True
          .PrintColor = True
          .PrintGrid = True
          .PrintOrientation = PrintOrientationPortrait
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
          mbooPrintPreview = False
      End With
      Exit Sub
Error_cmdPrint_Click:
    MsgBox Err.Description, vbInformation
    Exit Sub
End Sub

Private Sub mnu_Schedule_Patient_Click()
    If cctProvider.Text = vbNullString Then
        MsgBox "Please select a Provider first!", vbInformation, "PatientTrac Scheduling"
        Exit Sub
    Else
        frmSchedPatient.mbLookupPatient = False
        mLoadPatient = False
        udtSchedPatient.IntId = vbNullString
        Me.Hide
        udtAppointment.SchedDate = Format(MonthView1.Value, "mm/dd/yyyy")
        frmSchedPatient.Show vbModal
        Get_Mode
        Me.Show
        If cmdStop.Caption = "Stop Refresh" Then
             Get_Mode
             schTimer.Enabled = True
         End If
    End If
    If Me.WindowState = 1 Then
        Me.WindowState = 0
    End If
    Exit Sub
End Sub

Private Sub mnu_Week_Click()
    mstrApptMode = "Weekly"
    Get_Mode
    cmdPrevWeek.Visible = True
    cmdCurrentWeek.Visible = True
    cmdNextWeek.Visible = True
    lblApptDate.Visible = False
End Sub

Private Function Remove_InsertCells()
Dim lngCol As Long
Dim lngRow As Long
Dim vDate
Dim vdate2
Dim anchorcol  As Variant
Dim anchorrow As Variant
Dim colspan As Variant
Dim rowspan As Variant
Dim IsSpan As Boolean
Dim Span As Integer

    With grdData
        .Redraw = False
        For lngCol = 1 To .MaxCols
            For lngRow = 1 To grdData.MaxRows
               .GetText 0, lngRow, vDate
              If lngRow > 0 Then .GetText 0, lngRow - 1, vdate2
                If vDate <> vbNullString And vDate <> vbNullString Then
                    If vDate = vdate2 Then
                        .DeleteRows (lngRow), 1
                        .MaxRows = .MaxRows - 1
                    End If
                End If
            Next lngRow
        Next lngCol
        .Redraw = True
    End With
    Exit Function
End Function
Private Function LookUpHeader(vstrName As String) As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
    
    lngCol = 0
    With grdData
        LookUpHeader = vbNullString
        For lngRow = 1 To .MaxRows
            .Col = lngCol
            .GetText lngCol, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString And TextFound(vstrName, Trim(varCellValue), Len(vstrName), True) Then
                LookUpHeader = lngCol & "|" & lngRow
                Exit For
            End If
        Next lngRow
    End With
    Exit Function
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

    ' Get the selected cells
    grdData.getSelection 0, Col, Row, Col2, Row2
    If clsApptPref.OverBook = False Then Exit Function
    grdData.Redraw = False
    For i = 1 To 7
        For j = SpreadHeader To grdData.MaxRows
            Span = grdData.GetCellSpan(i, j, anchorcol, anchorrow, colspan, rowspan)
            
            grdData.Redraw = False
            If Span Then
                If grdData.Col = 1 Then anchorcol = "1"
                If j = SpreadHeader Then anchorrow = SpreadHeader
                If j = SpreadHeader + 1 Then anchorrow = SpreadHeader + 1
                If j = SpreadHeader + 2 Then anchorrow = SpreadHeader + 2
                grdData.RemoveCellSpan anchorcol, anchorrow
                IsSpan = True
            End If
        Next j
    Next i
    
    grdData.Redraw = True
    Exit Function
End Function

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
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim bFound As Boolean
Dim lngProvTest As Long

    On Error GoTo Error_Get_Appt_Providers
    Screen.MousePointer = vbHourglass
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    Format_Providers
    udtFacility.FacID = piFacilityID
    If udtFacility.FacID = 0 Then
        Set_DefaultFacility
        udtFacility.FacID = piFacilityID
    End If
    strSQL = "Get_Appt_Providers " & udtFacility.FacID
     Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        Do Until .EOF
            strFirst = vbNullString
            strMiddle = vbNullString
            strLast = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            lngProvID = 0
            booSkip = False
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
                booSkip = False
                With grdProviders
                    For lngRow = 1 To .MaxRows
                        .GetText 1, lngRow, varCellValue
                        If Trim(varCellValue) = clsProvider.Name Then
                            booSkip = True
                        End If
                    Next lngRow
                    If Not booSkip Then
                        lngRow = grdProviders.MaxRows + 1
                        grdProviders.MaxRows = lngRow
                        grdProviders.SetRowItemData lngRow, lngProvID
                        lngCol = 1
                        Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", clsProvider.Name, "TypeHAlignLeft"
                        grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 2
                        Format_CellTypeEdit grdProviders, lngRow, lngCol, "1", CStr(lngProvID), "TypeHAlignCenter"
                        grdProviders.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        
                    End If
                End With
            End If
        Loop
    End With
    If Get_SuperProv Then
        lngProvTest = piProviderID
    End If
    If lngProvTest > 0 Then
        lngProvTest = piProviderID
    Else
        lngProvTest = mlngLogonUser
    End If
    If lngProvTest > 0 Then
        With grdProviders
            For lngRow = 1 To .MaxRows
                If lngProvTest = .GetRowItemData(lngRow) Then
                    .GetText 1, lngRow, varCellValue
                    .BlockMode = True
                        .Col = 1
                        .Col2 = 2
                        .Row = lngRow
                        .Row2 = lngRow
                        .BackColor = grdCurItem
                    .BlockMode = False
                    bFound = True
                    cctProvider.Text = Trim(varCellValue)
                    Exit For
                End If
            Next lngRow
        End With
        If bFound And Get_SuperProv Then
            udtAppointment.ProviderID = piProviderID
        Else
            udtAppointment.ProviderID = mlngLogonUser
        End If
    End If
    udtAppointment.ProviderID = mlngLogonUser
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
Dim lngCurID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue

    On Error GoTo Error_Load_Facilities
    Screen.MousePointer = vbHourglass
    Format_Facilities
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "Get_Appt_Facilities"
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
                                strName = Trim(Fld.Value)
                            Case "FACILITY_ID"
                                lngFacID = Fld.Value
                            Case "DEFAULT_FACILITY"
                                booDefault = True
                        End Select
                    End If
                End If
            Next
            .MoveNext
            strName = Fix_Name(strName)
            strName = Fix_Facility_Name(strName)
            If Trim(strName) <> vbNullString Then
                booSkip = False
                With grdFacility
                    For lngRow = 1 To .MaxRows
                        .GetText 1, lngRow, varCellValue
                        If Trim(varCellValue) = strName Then
                            booSkip = True
                        End If
                    Next lngRow
                    If Not booSkip Then
                        lngRow = grdFacility.MaxRows + 1
                        grdFacility.MaxRows = lngRow
                        grdFacility.SetRowItemData lngRow, lngFacID
                        lngCol = 1
                        Format_CellTypeEdit grdFacility, lngRow, 1, "1", strName, "TypeHAlignLeft"
                        grdFacility.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 2
                        Format_CellTypeEdit grdFacility, lngRow, lngCol, "1", CStr(lngFacID), "TypeHAlignCenter"
                        grdFacility.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                End With
            End If
        Loop
    End With
     With grdFacility
        For lngRow = 1 To .MaxRows
            .GetText 1, lngRow, varCellValue
            If varCellValue = cctFacilities Then
                .BlockMode = True
                    .Col = 1
                    .Col2 = 2
                    .Row = lngRow
                    .Row2 = lngRow
                    .BackColor = grdCurItem
                .BlockMode = False
                udtFacility.FacID = .GetRowItemData(lngRow)
                piFacilityID = udtFacility.FacID
                Exit For
            End If
        Next lngRow
    End With
    If udtFacility.FacID = 0 Then
        Get_Current_Facility
        piFacilityID = udtFacility.FacID
    End If
    If udtFacility.FacID = 0 Then
        Set_DefaultFacility
        udtFacility.FacID = piFacilityID
    End If
    If piFacilityID > 0 Then
        With grdFacility
            For lngRow = 1 To .MaxRows
                If piFacilityID = .GetRowItemData(lngRow) Then
                    .GetText 1, lngRow, varCellValue
                    .BlockMode = True
                        .Col = 1
                        .Col2 = 2
                        .Row = lngRow
                        .Row2 = lngRow
                        .BackColor = grdCurItem
                    .BlockMode = False
                    cctFacilities.Text = Trim(varCellValue)
                    Exit For
                End If
            Next lngRow
        End With
        udtFacility.FacID = piFacilityID
    End If
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
    Write_Error Err.Description, Err.Number, strSQL, "Load_Facilities", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Public Sub InfoFrame(pCol As Long, pRow As Long)
Dim vValue
Dim iAnswer As Integer
Dim lngPatientID As Long
Dim vDate
Dim Arr

    On Error GoTo Error_InfoFrame
    Screen.MousePointer = vbHourglass
    mChange = False
    grdData.GetText pCol, pRow, vValue
    grdData.GetText 0, pRow, vTime
    If vTime = vbNullString Then
        grdData.GetText 0, pRow - 1, vTime
    End If
    If vTime <> vbNullString Then
        udtAppointment.SchedTime = vTime
        Validate_Appt_Time Val(udtSchedPatient.IntId)
    End If
    lblApptDate.AutoSize = True
    lblApptDate.Caption = "Appt Date:" & " " & Format(udtAppointment.SchedDate, "mmmm dd, yyyy")
    If mstrApptMode <> "Day" Then
        grdData.GetText pCol, SpreadHeader + 2, vDate
        Arr = Split(vDate, "-")
        If UBound(Arr) > 0 Then
            vDate = Arr(1)
            udtAppointment.SchedDate = vDate
        Else
            udtAppointment.SchedDate = vDate
        End If
    End If
    schTimer.Enabled = True
    mbooByPass = False
    Screen.MousePointer = vbDefault
    Exit Sub
Error_InfoFrame:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "InfoFrame", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Public Function GetDay(ipDate As String)
    Dim lDay As Long

    lDay = Weekday(ipDate)

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

Private Function WeekMode_BlockTime(vstrBlockReason As String, vstrBlockTo As String, _
vstrBlockFrom As String, vstrBlockDuration As String, vstrStartDate As String, vlngProviderID As Long)
'Dim i%
'Dim lngDateValue As Long
'Dim lngProviderID As Long
'Dim strSQL As String
'Dim vTime
'Dim vDate
'Dim BlockColor As Long
'Dim lngStartRow As Long
'Dim lngEndRow As Long
'Dim lngNoRows  As Long
'Dim nDays As Long
'Dim lngRow As Long
'Dim Arr
'Dim booData As Boolean
'Dim lngDateDiff As Long
'Dim lngCol As Long
'Dim booBlock As Boolean
'Dim strApptDate As String
'Dim varCellValue
'Dim booCancel As Boolean
'Dim strSched As String
'Dim lngTestRow As Long
'Dim vApptDate
'
'    On Error GoTo Error_WeekMode_BlockTime
'    Screen.MousePointer = vbHourglass
''    For lngCol = 1 To grdData.MaxCols
''        grdData.GetText lngCol, SpreadHeader, varCellValue
''        Arr = Split(varCellValue, " - ")
''        If UBound(Arr) > 0 Then
''            strApptDate = Trim(Arr(1))
''            strApptDate = Format(strApptDate, "mm/dd/yyyy")
''        End If
''
'    For lngCol = 1 To grdData.MaxCols
'        booCancel = False
'        grdData.GetText lngCol, SpreadHeader, vApptDate
'        Arr = Split(vApptDate, " - ")
'        If UBound(Arr) > 0 Then
'            vApptDate = Trim(Arr(1))
'            vApptDate = Format(vApptDate, "mm/dd/yyyy")
'        End If
'        If IsProperCase(vstrBlockDuration) = "Permanent" Then
'            lngDateDiff = DateDiff("d", Format(vApptDate, "mm/dd/yyyy"), Format(udtBlock.StartDate, "mm/dd/yyyy"))
'            If lngDateDiff <= 0 Then
'                booCancel = False
'            Else
'                booCancel = True
'            End If
'        Else
'            lngDateDiff = DateDiff("d", Format(vApptDate, "mm/dd/yyyy"), Format(udtBlock.StartDate, "mm/dd/yyyy"))
'            If lngDateDiff <= 0 Then
'                lngDateDiff = DateDiff("d", Format(vApptDate, "mm/dd/yyyy"), Format(udtBlock.EndDate, "mm/dd/yyyy"))
'                If lngDateDiff > 0 Then
'                    booCancel = False
'                Else
'                    booCancel = True
'                End If
'            Else
'                booCancel = True
'            End If
'        End If
'        If Not booCancel Then
'            For lngRow = 1 To grdData.MaxRows
'                grdData.GetText 0, lngRow, vTime
'                If vTime = Trim(vstrBlockFrom) Then
'                    lngStartRow = lngRow
'                End If
'                grdData.GetText 0, lngRow, vTime
'                If vTime = Trim(vstrBlockTo) Then
'                    lngEndRow = lngRow
'                End If
'                If lngStartRow > 0 And lngEndRow > 0 Then
'                    Exit For
'                End If
'            Next lngRow
'            If lngStartRow > 0 And lngEndRow = 0 Then lngEndRow = grdData.MaxRows
'            lngNoRows = lngEndRow - lngStartRow
'            If IsProperCase(vstrBlockReason) = "Lunch" Then
'                BlockColor = grdBlock
'            ElseIf IsProperCase(vstrBlockReason) = "Hospital" Then
'                BlockColor = grdBlock
'            ElseIf IsProperCase(vstrBlockReason) = "Vacation" Then
'                BlockColor = grdBlock
'            End If
'            lngNoRows = lngNoRows
'            If lngStartRow >= 1 And lngEndRow > 1 Then
'                For lngRow = lngStartRow To lngEndRow
'                    grdData.GetText lngCol, lngRow, varCellValue
'                    If Trim(varCellValue) = vbNullString Or TextFound(vstrBlockReason, CStr(varCellValue), Len(vstrBlockReason), True) Then
'                        If lngRow = lngStartRow Then
'                            For lngTestRow = lngRow To lngEndRow - 1
'                                grdData.Col = lngCol
'                                grdData.Row = lngTestRow
'                                grdData.GetText lngCol, lngTestRow, varCellValue
'                                If Trim(varCellValue) = vbNullString Then
'                                    If grdData.CellTag <> vbNullString And grdData.CellTag <> "BlockMode" Then
'                                        Format_CellTypeEdit_RowSpan grdData, lngTestRow, lngCol, "1", vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
'                                        grdData.SetCellBorder lngCol, lngTestRow, lngCol, lngTestRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                                    ElseIf grdData.CellTag <> vbNullString And Not IsNumeric(grdData.CellTag) Then
'                                        Format_CellTypeEdit_RowSpan grdData, lngTestRow, lngCol, "1", vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
'                                        grdData.SetCellBorder lngCol, lngTestRow, lngCol, lngTestRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                                    Else
'                                        Format_CellTypeEdit_RowSpan grdData, lngTestRow, lngCol, "1", vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
'                                        grdData.SetCellBorder lngCol, lngTestRow, lngCol, lngTestRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                                    End If
'                                Else
'                                    grdData.Col = lngCol
'                                    grdData.Row = lngTestRow
'                                    If IsNumeric(grdData.CellTag) Then
'                                        booData = True
'                                    End If
'                                End If
'                            Next lngTestRow
'                            If Not booData Then
'                                Format_CellTypeEdit_RowSpan grdData, lngStartRow, lngCol, CStr(lngNoRows), vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
'                                grdData.SetCellBorder lngCol, lngStartRow, lngCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                            End If
'                        Else
'                            booData = False
'                            lngNoRows = lngNoRows - 1
'                            For lngTestRow = lngRow To lngEndRow - 1
'                                grdData.Col = lngCol
'                                grdData.Row = lngTestRow
'                                grdData.GetText lngCol, lngTestRow, varCellValue
'                                If Trim(varCellValue) = vbNullString Then
'                                    If grdData.CellTag <> vbNullString And grdData.CellTag <> "BlockMode" Then
'                                        Format_CellTypeEdit_RowSpan grdData, lngTestRow, lngCol, "1", vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
'                                        grdData.SetCellBorder lngCol, lngTestRow, lngCol, lngTestRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                                    ElseIf grdData.CellTag <> vbNullString And Not IsNumeric(grdData.CellTag) Then
'                                        Format_CellTypeEdit_RowSpan grdData, lngTestRow, lngCol, "1", vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
'                                        grdData.SetCellBorder lngCol, lngTestRow, lngCol, lngTestRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                                    Else
'                                        Format_CellTypeEdit_RowSpan grdData, lngTestRow, lngCol, "1", vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
'                                        grdData.SetCellBorder lngCol, lngTestRow, lngCol, lngTestRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                                    End If
'                               Else
'                                    grdData.Col = lngCol
'                                    grdData.Row = lngTestRow
'                                    If IsNumeric(grdData.CellTag) Then
'                                        booData = True
'                                    End If
'                                End If
'                            Next lngTestRow
'                            If Not booData Then
'                                Format_CellTypeEdit_RowSpan grdData, lngStartRow, lngCol, CStr(lngNoRows + 1), vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
'                                grdData.SetCellBorder lngCol, lngStartRow, lngCol, lngStartRow + lngNoRows, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                                grdData.SetCellBorder lngCol, lngEndRow + 1, lngCol, lngEndRow + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                            End If
'                            Exit For
'                        End If
'                    ElseIf lngRow > lngStartRow Then
'                        lngNoRows = lngNoRows - 1
'                    End If
'                Next lngRow
'            End If
'        End If
'        lngEndRow = 0
'        lngStartRow = 0
'    Next lngCol
'    Screen.MousePointer = vbDefault
'    Exit Function
'Error_WeekMode_BlockTime:
'    Screen.MousePointer = vbDefault
'    Write_Error Err.Description, Err.Number, "", "WeekMode_BlockTime", App.EXEName & "." & Me.Name
'    Exit Function
'    Resume Next
Dim i%
Dim lngDateValue As Long
Dim lngProviderID As Long
Dim strSQL As String
Dim vTime
Dim vDate
Dim BlockColor As Long
Dim lngStartRow As Long
Dim lngEndRow As Long
Dim lngNoRows  As Long
Dim nDays As Long
Dim lngRow As Long
Dim Arr
Dim booData As Boolean
Dim lngDateDiff As Long
Dim lngCol As Long
Dim booBlock As Boolean
Dim strApptDate As String
Dim varCellValue
Dim booCancel As Boolean
Dim strSched As String
Dim lngTestRow As Long
Dim vApptDate
Dim lngNo As Double
Dim strDuration As String
Dim strTestDate As String
Dim lngSpan As Long
Dim lngStartCol As Long
Dim lngEndCol As Long
Dim strColTime As String
Dim strMuliFac As String
Dim strErrata As String
Dim strNote As String
Dim strLine As String


    On Error GoTo Error_WeekMode_BlockTime
    Screen.MousePointer = vbHourglass
    strLine = "_______________________"
    Arr = Split(vstrBlockDuration, " ")
    If UBound(Arr) > 0 Then
        strDuration = Arr(1)
        lngNo = Arr(0)
    Else
        strDuration = vstrBlockDuration
    End If
    Select Case IsProperCase(strDuration)
        Case "Days", "Day"
            nDays = lngNo
        Case "Weeks", "Week"
            nDays = lngNo * 6
        Case "Month", "Months"
            nDays = lngNo * 29
        Case "Year", "Years"
            nDays = lngNo * 364
    End Select
    If IsProperCase(vstrBlockReason) = "Lunch" Then
        BlockColor = grdBlock
    ElseIf IsProperCase(vstrBlockReason) = "Hospital" Then
        BlockColor = grdBlock
    ElseIf IsProperCase(vstrBlockReason) = "Vacation" Then
        BlockColor = grdBlock
    End If
    For lngCol = 1 To grdData.MaxCols
        booCancel = False
        If lngCol = lngEndCol Then
            booCancel = True
        End If
        grdData.GetText lngCol, SpreadHeader + 2, varCellValue
        strColTime = varCellValue
        Arr = Split(strColTime, " - ")
        If UBound(Arr) > 0 Then
            strColTime = Trim(Arr(1))
            strColTime = Format(strColTime, "mm/dd/yyyy")
        Else
            booCancel = True
        End If
        If Not booCancel Then
            If grdData.GetColItemData(lngCol) = udtBlock.FacilityID Then
                If DateDiff("d", Format(udtBlock.StartDate, "mm/dd/yyyy"), Format(strColTime, "mm/dd/yyyy")) >= 0 Or _
                    DateDiff("d", Format(udtBlock.EndDate, "mm/dd/yyyy"), Format(strColTime, "mm/dd/yyyy")) <= 0 Then
                    booCancel = False
                Else
                    booCancel = True
                End If
                If Not booCancel Then
                    For lngRow = 1 To grdData.MaxRows
                        grdData.GetText 0, lngRow, vTime
                        If vTime = Trim(vstrBlockFrom) Then
                            lngStartRow = lngRow
                        End If
                        grdData.GetText 0, lngRow, vTime
                        If vTime = Trim(vstrBlockTo) Then
                            lngEndRow = lngRow
                        End If
                        If lngStartRow > 0 And lngEndRow > 0 Then
                            Exit For
                        End If
                    Next lngRow
                    If lngStartRow > 0 And lngEndRow = 0 Then lngEndRow = grdData.MaxRows
                    lngNoRows = lngEndRow - lngStartRow
                    lngNoRows = lngNoRows
                    If lngStartRow >= 1 And lngEndRow > 1 Then
                        For lngRow = lngStartRow To lngEndRow
                            grdData.GetText lngCol, lngRow, varCellValue
                            If Trim(varCellValue) <> vbNullString Then
                                If Not TextFound(vstrBlockReason, CStr(varCellValue), Len(vstrBlockReason), True) Then
                                    grdData.Col = lngCol
                                    grdData.Row = lngRow
                                    If grdData.CellTag <> vbNullString Then
                                        If grdData.CellTag <> "BlockMode" Then
                                            grdData.Col = lngCol
                                            grdData.Row = lngRow
                                            grdData.GetText lngCol, lngRow, varCellValue
                                            If Trim(varCellValue) <> vbNullString Then
                                                clsGridFormat.CellNote = Trim(varCellValue)
                                                Arr = Split(varCellValue, "INSURANCE AUTHORIZATION:")
                                                If UBound(Arr) > 0 Then
                                                   varCellValue = Trim(Arr(0))
                                                End If
                                                clsGridFormat.CellNote = Trim(varCellValue)
                                                If UBound(Arr) > 0 Then
                                                   clsGridFormat.CellNote = Arr(0)
                                                End If
                                                If udtAppointment.SchedType = vbNullString Then udtAppointment.SchedType = "Patient"
                                                If Trim(clsGridFormat.CellNote) <> vbNullString Then
                                                    If IsProperCase(Trim(clsGridFormat.CellNote)) = "Patient" Or _
                                                        IsProperCase(Trim(clsGridFormat.CellNote)) = "Vendor" Or _
                                                        IsProperCase(Trim(clsGridFormat.CellNote)) = "Physician" Or _
                                                        IsProperCase(Trim(clsGridFormat.CellNote)) = "Legal" Then
                                                        grdData.GetText 0, lngRow, varCellValue
                                                        If Trim(varCellValue) <> vbNullString Then
                                                            strErrata = Trim(varCellValue)
                                                        End If
                                                        If Trim(clsGridFormat.CellNote) <> vbNullString Then strNote = udtAppointment.SchedType & " " & "is scheduled in the block time at:" & " " & strErrata & vbNewLine & clsGridFormat.CellNote
                                                        If Trim(rtfChanges.Text) = vbNullString Then
                                                            rtfChanges.Text = Trim(rtfChanges.Text) & vbNewLine & Trim(strNote)
                                                        ElseIf Not TextFound(Trim(strNote), Trim(rtfChanges.Text), Len(Trim(strNote)), True) Then
                                                            rtfChanges.Text = Trim(rtfChanges.Text) & vbNewLine & strLine & vbNewLine & Trim(strNote)
                                                        End If
                                                        lblSchedChange.Visible = True
                                                        rtfChanges.Visible = True
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next lngRow
                    End If
                    If grdData.MaxCols >= lngCol + 1 Then
                        lngEndCol = 0
                        If grdData.GetColItemData(lngCol + 1) = udtBlock.FacilityID Then
                            grdData.GetText lngCol + 1, SpreadHeader + 2, varCellValue
                            If Format(varCellValue, "mm/dd/yyyy") = Format(strColTime, "mm/dd/yyyy") Or Trim(varCellValue) = vbNullString Then
                                lngEndCol = lngCol + 1
                            End If
                        ElseIf lngCol > 1 And grdData.GetColItemData(lngCol - 1) = udtBlock.FacilityID Then
                            grdData.GetText lngCol - 1, SpreadHeader + 2, varCellValue
                            If Format(varCellValue, "mm/dd/yyyy") = Format(strColTime, "mm/dd/yyyy") Or Trim(varCellValue) = vbNullString Then
                                lngCol = lngEndCol
                            End If
                        End If
                    ElseIf grdData.GetColItemData(lngCol) = udtBlock.FacilityID Then
                            grdData.GetText lngCol, SpreadHeader + 2, varCellValue
                            If Format(varCellValue, "mm/dd/yyyy") = Format(strColTime, "mm/dd/yyyy") Then
                                lngCol = lngEndCol
                            End If
                        End If
                    End If
                    If lngEndCol > 0 And lngCol > 0 Then
                        Format_CellTypeEdit_RowSpan grdData, lngStartRow, lngCol, CStr(lngNoRows), vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock, "", 2
                        grdData.SetCellBorder lngCol, lngStartRow, lngCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Else
                        Format_CellTypeEdit_RowSpan grdData, lngStartRow, lngCol, CStr(lngNoRows), vstrBlockReason, "TypeHAlignCenter", "BlockMode", grdBlock
                        grdData.SetCellBorder lngCol, lngStartRow, lngCol, lngEndRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                End If
            End If
            lngEndRow = 0
            lngStartRow = 0
        Next lngCol
    Screen.MousePointer = vbDefault
    Exit Function
Error_WeekMode_BlockTime:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "WeekMode_BlockTime", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Sub ReadData()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Arr
Dim lngPatID As Long
Dim Idx As Integer
Dim lngRow As Long
Dim varCellValue
Dim vTime
Dim i%
Dim lngCol As Long
Dim lngFoundID As Boolean
Dim lngProvID As Long
Dim strNotes As String
Dim bPatFound As Boolean
Dim lngApptID As Long
Dim strConflict As String
Dim bExit As Boolean
Dim lngTest As Long
Dim lngChgCol As Long


    On Error GoTo Error_ReadData
    Screen.MousePointer = vbHourglass
    Get_Providers udtAppointment.ProviderID
    Format_Changes
    clsShedGrid.ProvCol = 1
    grdData.MaxCols = 1
    lngProvID = udtAppointment.ProviderID
    grdData.ClearRange 1, 1, grdData.MaxCols, grdData.MaxRows, False
    Format_CellTypeStatic grdData, SpreadHeader + 1, 1, CStr(1), "Current Date" & " " & Format(udtAppointment.SchedDate, "dddd mmmm dd, yyyy"), "TypeHAlignCenter", "", "", HeaderColor
    Add_Spacer SpreadHeader + 2, grdData
    For lngCol = 1 To grdData.MaxCols
        For lngRow = 1 To grdData.MaxRows
            grdData.RowHeight(lngRow) = 15 * mlngSystemFileSize
            grdData.RemoveCellSpan lngCol, lngRow
            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Next lngRow
    Next lngCol
    If clsShedGrid.ProvCol2 = 0 Then clsShedGrid.ProvCol2 = 1
    If mstrApptMode = vbNullString Then mstrApptMode = "Day"
    If udtAppointment.SchedDate = vbNullString Then
        udtAppointment.SchedDate = Format(Now, "mm/dd/yyyy")
    Else
        udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    End If
    If cctFacilities.Text <> vbNullString Then
        With grdFacility
            For lngRow = 1 To .MaxRows
                .GetText 1, lngRow, varCellValue
                If varCellValue = cctFacilities Then
                    .BlockMode = True
                        .Col = 1
                        .Col2 = 2
                        .Row = lngRow
                        .Row2 = lngRow
                        .BackColor = grdCurItem
                    .BlockMode = False
                    udtFacility.FacID = .GetRowItemData(lngRow)
                    Exit For
                End If
            Next lngRow
        End With
    End If
    If udtFacility.FacID = 0 Then udtFacility.FacID = piFacilityID
    If udtAppointment.ProviderID = 0 Then
        If cctProvider.Text <> vbNullString Then
            With grdProviders
                For lngRow = 1 To .MaxRows
                    .GetText 1, lngRow, varCellValue
                    If varCellValue = cctProvider Then
                        .BlockMode = True
                            .Col = 1
                            .Col2 = 2
                            .Row = lngRow
                            .Row2 = lngRow
                            .BackColor = grdCurItem
                        .BlockMode = False
                        udtAppointment.ProviderID = .GetRowItemData(lngRow)
                        Exit For
                    End If
                Next lngRow
            End With
        End If
    End If
    clsShedGrid.ApptDate = udtAppointment.SchedDate
    clsShedGrid.ApptDate = Format(clsShedGrid.ApptDate, "mm/dd/yyyy")
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    If lngProvID > 0 Then
        strSQL = "Get_Provider_Appt_Facility  " & lngProvID & ", " & udtFacility.FacID & ", '" & clsShedGrid.ApptDate & "'"
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        With rsReadAppt
            Do Until .EOF
                clsGridFormat.CellNote = vbNullString
                Clear_SchedPatient
                Clear_Appt
                udtAppointment.SchedType = vbNullString
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) > vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "PATIENT_ID"
                                    udtSchedPatient.IntId = Fld.Value
                                Case "PURPOSE_VISIT"
                                    udtAppointment.VisitType = Fld.Value
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
                                    udtAppointment.SchedTime = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                                Case "DURATION"
                                    udtSchedPatient.Duration = Fld.Value
                                Case "PATIENT_TYPE"
                                    udtAppointment.PatientType = Fld.Value
                                Case "APPT_ID"
                                    udtAppointment.Appt_ID = Fld.Value
                                Case "ADD_PROC_CODE"
                                  udtAppointment.Add_Proc_CPT = Trim(Fld.Value)
                                Case "ADD_PROC_DESC"
                                  udtAppointment.Add_Proc_DESC = Trim(Fld.Value)
                            End Select
                        End If
                    End If
                Next
            .MoveNext
            If Not Valid_RepeatAppt(udtAppointment.Appt_ID) Then
                lngFoundID = False
                If Not udtSchedPatient.Status = "No Show" Or udtSchedPatient.Status = "Cancelled" Then
                    If Trim(udtAppointment.SchedType) <> vbNullString Then
                        Arr = Split(LookUpHeader(udtAppointment.SchedTime), "|")
                        If UBound(Arr) > 0 Then
                            clsShedGrid.Row = Arr(1)
                        End If
                        udtAppointment.SchedDate = clsShedGrid.ApptDate
                        Select Case IsProperCase(udtAppointment.SchedType)
                            Case "Patient"
                                If Trim(udtSchedPatient.IntId) <> vbNullString Then
                                    Get_SchedPatient Val(udtSchedPatient.IntId)
                                End If
                            Case "Vendor", "Business"
                                If udtAppointment.SchedTime <> vbNullString Then
                                    Get_SchedVendor Val(udtSchedPatient.IntId), udtAppointment.SchedTime
                                Else
                                    Get_SchedVendor Val(udtSchedPatient.IntId)
                                End If
                            Case "Legal", "Attorney"
                                Get_Schedule_Attorney Val(udtSchedPatient.IntId)
                            Case "Physician"
                                Get_SchedPhysician Val(udtSchedPatient.IntId)
                        End Select
                        Select Case UCase(clsApptPref.SchedInterval)
                            Case "EVERY 20 MIN"
                                clsShedGrid.Duration = 20
                            Case "EVERY 30 MIN"
                                clsShedGrid.Duration = 30
                            Case "EVERY 15 MIN"
                                clsShedGrid.Duration = 15
                            Case "EVERY 10 MIN"
                                clsShedGrid.Duration = 10
                            Case "EVERY 5 MIN"
                                clsShedGrid.Duration = 5
                        End Select
                        If udtSchedPatient.Duration = "1 Hour" Then
                            udtSchedPatient.Duration = "60 minutes"
                        End If
                        Arr = Split(udtSchedPatient.Duration, " ")
                        If UBound(Arr) > 0 Then
                            If Val(clsShedGrid.Duration) > 0 Then
                                If IsNumeric(Arr(0)) Then clsShedGrid.SpanRow = Val(Arr(0) / Val(clsShedGrid.Duration))
                            End If
                        ElseIf Val(clsShedGrid.Duration) > 0 Then
                           If IsNumeric(udtSchedPatient.Duration) Then clsShedGrid.SpanRow = Val(udtSchedPatient.Duration / Val(clsShedGrid.Duration))
                        End If
                        udtSchedPatient.Name = Replace(udtSchedPatient.Name, "''", "'")
                        If clsShedGrid.Row > 0 Then
                            For lngCol = clsShedGrid.ProvCol To clsShedGrid.ProvCol2
                                If Not lngFoundID Then
                                   For lngRow = clsShedGrid.Row To Val(clsShedGrid.Row + (clsShedGrid.SpanRow - 1))
                                            grdData.Col = lngCol
                                            grdData.Row = lngRow
                                            If grdData.CellTag <> vbNullString Then
                                            If grdData.CellTag <> udtSchedPatient.IntId Then
                                                strConflict = vbNullString
                                                grdData.GetText 0, lngRow, varCellValue
                                                strConflict = varCellValue
                                                clsGridFormat.CellNote = "OverBook" & vbNewLine & strConflict
                                                'End If
                                                If Trim(clsGridFormat.CellNote) <> vbNullString And Not IsDate(grdData.CellTag) Then
                                                    If Trim(clsGridFormat.CellNote) <> vbNullString Then
                                                        bExit = False
                                                        With grdChanges
                                                            For lngTest = 1 To .MaxRows
                                                                .Col = 1
                                                                .Row = lngTest
                                                                If .GetRowItemData(lngTest) = Val(udtSchedPatient.IntId) Then
                                                                    bExit = True
                                                                    Exit For
                                                                End If
                                                            Next lngTest
                                                            If Not bExit Then
                                                                lngTest = .MaxRows + 1
                                                                .MaxRows = lngTest
                                                                .SetRowItemData lngTest, Val(udtSchedPatient.IntId)
                                                                lngChgCol = 1
                                                                Format_CellTypeStatic grdChanges, lngTest, lngChgCol, "1", Trim(udtSchedPatient.Name), "TypeHAlignLeft", "", "", 0, udtAppointment.SchedType
                                                                grdChanges.SetCellBorder lngChgCol, lngTest, lngChgCol, lngTest, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                                                lngChgCol = 2
                                                                Format_CellTypeEdit grdChanges, lngTest, lngChgCol, "1", Trim(clsGridFormat.CellNote), "TypeHAlignCenter"
                                                                grdChanges.SetCellBorder lngChgCol, lngTest, lngChgCol, lngTest, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                                            End If
                                                        End With
                                                    End If
                                                End If
                                                If clsShedGrid.ProvCol2 <= clsShedGrid.ProvCol Then
                                                    clsShedGrid.ProvCol2 = clsShedGrid.ProvCol + 1
                                                    Format_AdditionalCol
                                                End If
                                                Add_SchedRow clsShedGrid.ProvCol2
                                                lngFoundID = True
                                                Exit For
                                            Else
                                               lngFoundID = True
                                               Exit For
                                            End If
                                        Else
                                            lngFoundID = False
                                            Exit For
                                        End If
                                    Next lngRow
                                End If
                                If Not lngFoundID Then
                                    Add_SchedRow clsShedGrid.ProvCol
                                    Exit For
                                End If
                            Next lngCol
                            End If
                        End If
                    End If
                End If
            Loop
        End With
    End If
    'get count
    mlngPatCnt = 0
    strSQL = "SELECT DISTINCT PATIENT_ID FROM APPOINTMENTS WHERE " _
        & "PROVIDER_ID = " & lngProvID & " AND FACILITY_ID = " _
        & udtFacility.FacID & " AND ACTIVE = 1 AND " _
        & "CONVERT(VARCHAR(10), APPT_DATE, 101) = '" & clsShedGrid.ApptDate & "'"
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        With rsReadAppt
            Do Until .EOF
                .MoveNext
                mlngPatCnt = mlngPatCnt + 1
            Loop
        End With
    If cnReadAppt.State = adStateClosed Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    udtAppointment.SchedDate = clsShedGrid.ApptDate
    mbooByPass = False
    Set_BlockTime udtAppointment.ProviderID
    If mlngGridRow > 0 Then
       grdData.TopRow = mlngGridRow
    End If
    udtAppointment.Appt_ID = 0
    If grdData.Row >= 1 Then InfoFrame grdData.Col, grdData.Row
    mFirstLoad = False
    mbooByPass = False
    udtAppointment.Add_Proc_CPT = vbNullString
    udtAppointment.Add_Proc_DESC = vbNullString
    udtAppointment.SchedType = "Patient"
    Get_Changes clsShedGrid.ApptDate
    Clear_SchedPatient
    udtSchedPatient.IntId = vbNullString
    lblCurScheAppt.Visible = True
    lblCurScheAppt.Caption = "Current Patient Appointments:" & "    " & CStr(mlngPatCnt)
    Screen.MousePointer = vbDefault
    Exit Sub
Error_ReadData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "ReadData", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Function Look_Up_Column_Date_Week(vstrDate As String, Optional bOverBook As Boolean)
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim lngFirst As Long
Dim lngSecond As Long
Dim lngInsertCol As Long
Dim Arr
Dim strDate As String

    lngCol = 0
    grdData.Redraw = False
    With grdData
        For lngCol = 1 To .MaxCols
            .Col = lngCol
            .Row = SpreadHeader + 2
            .GetText lngCol, SpreadHeader + 2, varCellValue
            Arr = Split(varCellValue, "-")
            If UBound(Arr) > 0 Then
                strDate = Arr(1)
            End If
            If Format(strDate, "mm/dd/yyyy") = Format(vstrDate, "mm/dd/yyyy") Then
                If .GetColItemData(lngCol) = udtAppointment.FacilityID Then
                    If lngFirst = 0 Then
                        lngFirst = lngCol
                    ElseIf lngSecond = 0 Then
                        lngSecond = lngCol
                        Exit For
                    End If
                ElseIf .GetColItemData(lngCol) = 0 Then
                    .SetColItemData lngCol, udtAppointment.FacilityID
                    clsShedGrid.ProvCol = lngCol
                    Exit For
                ElseIf .GetColItemData(lngCol + 1) = 0 Then
                    .GetText lngCol, SpreadHeader + 2, varCellValue
                    Arr = Split(varCellValue, "-")
                    If UBound(Arr) > 0 Then
                        strDate = Arr(1)
                    End If
                    If Format(strDate, "mm/dd/yyyy") = Format(vstrDate, "mm/dd/yyyy") Then
                        .SetColItemData lngCol, udtAppointment.FacilityID
                        clsShedGrid.ProvCol = lngCol + 1
                        Exit For
                    End If
                ElseIf .GetColItemData(lngCol + 1) = udtAppointment.FacilityID Then
                    .GetText lngCol, SpreadHeader + 2, varCellValue
                    Arr = Split(varCellValue, "-")
                    If UBound(Arr) > 0 Then
                        strDate = Arr(1)
                    End If
                    If Format(strDate, "mm/dd/yyyy") = Format(vstrDate, "mm/dd/yyyy") Then
                        clsShedGrid.ProvCol = lngCol + 1
                        Exit For
                    End If
'                Else
'                    clsShedGrid.ProvCol = lngCol
'                    clsShedGrid.ProvCol2 = lngCol + 1
'                    .MaxCols = .MaxCols + 1
'                    lngInsertCol = clsShedGrid.ProvCol2
'                    .InsertCols lngInsertCol, 1
'                    .ColWidth(clsShedGrid.ProvCol2) = 20
'                    grdData.RemoveCellSpan clsShedGrid.ProvCol, SpreadHeader + 4
'                    Format_CellTypeFixed grdData, SpreadHeader + 2, clsShedGrid.ProvCol, "2", GetDay(vstrDate) & " - " & vstrDate, "TypeHAlignCenter", "", HeaderColor
'                    grdData.SetCellBorder clsShedGrid.ProvCol, SpreadHeader + 2, clsShedGrid.ProvCol2, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                    Add_Spacer SpreadHeader + 3, grdData
'                    grdData.RemoveCellSpan clsShedGrid.ProvCol, SpreadHeader + 4
'                    Format_CellTypeFixed grdData, SpreadHeader + 4, clsShedGrid.ProvCol, "2", clsFacility.Name, "TypeHAlignCenter", CStr(udtAppointment.ProviderID), HeaderColor
'                    grdData.SetCellBorder clsShedGrid.ProvCol2, SpreadHeader + 4, clsShedGrid.ProvCol2, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                    .SetColItemData clsShedGrid.ProvCol, udtAppointment.FacilityID
'                    .SetColItemData clsShedGrid.ProvCol2, udtAppointment.FacilityID
'                    For lngRow = 1 To .MaxRows
'                        Format_CellTypeEdit grdData, lngRow, clsShedGrid.ProvCol2, "1", " ", "TypeHAlignCenter"
'                        grdData.SetCellBorder clsShedGrid.ProvCol2, lngRow, clsShedGrid.ProvCol2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
'                    Next lngRow
'                    Exit For
                End If
            Else
                Debug.Print
            End If
        Next lngCol
        If lngFirst > 0 Then clsShedGrid.ProvCol = lngFirst
        clsShedGrid.ProvCol2 = lngSecond
        If clsShedGrid.ProvCol = 0 Then
            Debug.Print
        End If
        If bOverBook Then
            If clsShedGrid.ProvCol2 = 0 Then
                clsShedGrid.ProvCol2 = clsShedGrid.ProvCol + 1
                .MaxCols = .MaxCols + 1
                lngInsertCol = clsShedGrid.ProvCol2
                .InsertCols lngInsertCol, 1
                .ColWidth(clsShedGrid.ProvCol2) = 20 * mlngSystemFileSize
                grdData.RemoveCellSpan clsShedGrid.ProvCol, SpreadHeader + 4
                Format_CellTypeFixed grdData, SpreadHeader + 2, clsShedGrid.ProvCol, "2", GetDay(vstrDate) & " - " & vstrDate, "TypeHAlignCenter", "", HeaderColor
                grdData.SetCellBorder clsShedGrid.ProvCol, SpreadHeader + 2, clsShedGrid.ProvCol2, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Add_Spacer SpreadHeader + 3, grdData
                grdData.RemoveCellSpan clsShedGrid.ProvCol, SpreadHeader + 4
                Format_CellTypeFixed grdData, SpreadHeader + 4, clsShedGrid.ProvCol, "2", clsFacility.Name, "TypeHAlignCenter", CStr(udtAppointment.ProviderID), HeaderColor
                grdData.SetCellBorder clsShedGrid.ProvCol2, SpreadHeader + 4, clsShedGrid.ProvCol2, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                .SetColItemData clsShedGrid.ProvCol, udtAppointment.FacilityID
                .SetColItemData clsShedGrid.ProvCol2, udtAppointment.FacilityID
                For lngRow = 1 To .MaxRows
                    Format_CellTypeEdit grdData, lngRow, clsShedGrid.ProvCol2, "1", " ", "TypeHAlignCenter"
                    grdData.SetCellBorder clsShedGrid.ProvCol2, lngRow, clsShedGrid.ProvCol2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Next lngRow
            End If
        End If
    End With
    grdData.Redraw = True
    Exit Function
End Function

Private Function Get_Current_Facility()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim strTime As String
Dim strEnd As String

    On Error GoTo Error_Get_Current_Facility
    Screen.MousePointer = vbHourglass
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strTime = DateAdd("n", -30, Now)
    strTime = Format(strTime, "hh:mm AM/PM")
    strEnd = DateAdd("n", 10, Now)
    strEnd = Format(strEnd, "hh:mm AM/PM")
    strSQL = "SELECT FACILITY_ID FROM APPOINTMENTS WHERE PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
        & "APPT_DATE = '" & Format(Now, "mm/dd/yyyy") & "' AND APPT_TIME  BETWEEN  '" & strTime & "' AND '" & strEnd & "'"
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        With rsReadAppt
            If Not .EOF Then
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) > vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "FACILITY_ID"
                                    udtAppointment.FacilityID = Fld.Value
                            End Select
                        End If
                    End If
                Next
            End If
        End With
        If udtAppointment.FacilityID = 0 Then
            strSQL = "SELECT FACILITY_ID FROM APPOINTMENTS WHERE " _
                & "PROVIDER_ID = " & udtAppointment.ProviderID & " AND ACTIVE = 1 AND APPT_DATE = '" & Format(Now, "mm/dd/yyyy") & "'"
            Set rsReadAppt = cnReadAppt.Execute(strSQL)
            With rsReadAppt
                If Not .EOF Then
                    Set Flds = .Fields
                    For Each Fld In Flds
                        If Not IsNull(Fld.Value) Then
                            If Trim(Fld.Value) > vbNullString Then
                                Select Case UCase(Fld.Name)
                                    Case "FACILITY_ID"
                                        udtAppointment.FacilityID = Fld.Value
                                End Select
                            End If
                        End If
                    Next
                End If
            End With
        End If
    udtFacility.FacID = udtAppointment.FacilityID
    If cnReadAppt.State = adStateClosed Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    Exit Function
Error_Get_Current_Facility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Current_Facility", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Function Format_AdditionalCol()
Dim lngRow As Long
    With grdData
        .ColHeaderRows = 4
        If .MaxCols < clsShedGrid.ProvCol2 Then
            .ColWidth(clsShedGrid.ProvCol) = 39.35 * mlngSystemFileSize
            .MaxCols = grdData.MaxCols + 1
            .ColWidth(clsShedGrid.ProvCol2) = 39.35 * mlngSystemFileSize
            .ClearRange 1, SpreadHeader, grdData.MaxCols, SpreadHeader, True
            Format_CellTypeEdit grdData, SpreadHeader, 1, CStr(grdData.MaxCols), cctFacilities.Text, "TypeHAlignCenter", CStr(udtAppointment.ProviderID), HeaderColor
            .SetCellBorder 1, SpreadHeader, grdData.MaxCols, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Format_CellTypeStatic grdData, SpreadHeader + 1, 1, CStr(grdData.MaxCols), "Current Date" & " " & Format(udtAppointment.SchedDate, "dddd mmmm dd, yyyy"), "TypeHAlignCenter", "", "", HeaderColor
            .SetCellBorder 1, SpreadHeader + 1, 2, SpreadHeader + 1, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Add_Spacer SpreadHeader + 2, grdData
            Format_CellTypeEdit grdData, SpreadHeader + 3, clsShedGrid.ProvCol, "2", clsProvider.Name, "TypeHAlignCenter", CStr(udtAppointment.ProviderID), HeaderColor
            .SetCellBorder clsShedGrid.ProvCol, SpreadHeader + 3, clsShedGrid.ProvCol2, SpreadHeader + 3, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        For lngRow = 1 To .MaxRows
            .SetCellBorder clsShedGrid.ProvCol2, lngRow, clsShedGrid.ProvCol2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Next lngRow
    End With
End Function
Private Sub ReadData_NameOnly()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Arr
Dim lngPatID As Long
Dim Idx As Integer
Dim lngRow As Long
Dim varCellValue
Dim vTime
Dim i%
Dim lngCol As Long
Dim lngFoundID As Boolean
Dim lngProvID As Long
Dim strNotes As String
Dim bPatFound As Boolean

    On Error GoTo Error_ReadData_NameOnly
    Screen.MousePointer = vbHourglass
    lngProvID = udtAppointment.ProviderID
    mlngPatCnt = 0
    grdData.ClearRange 1, 1, grdData.MaxCols, grdData.MaxRows, False
    For lngCol = 1 To grdData.MaxCols
        For lngRow = 1 To grdData.MaxRows
            grdData.RowHeight(lngRow) = 10
            grdData.RemoveCellSpan lngCol, lngRow
            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Next lngRow
    Next lngCol
    If clsShedGrid.ProvCol2 = 0 Then clsShedGrid.ProvCol2 = 1
    rtfChanges.Text = vbNullString
    If mstrApptMode = vbNullString Then mstrApptMode = "Day"
    If udtAppointment.SchedDate = vbNullString Then
        udtAppointment.SchedDate = Format(Now, "mm/dd/yyyy")
    Else
        udtAppointment.SchedDate = Format(udtAppointment.SchedDate, "mm/dd/yyyy")
    End If
    udtFacility.FacID = grdFacility.GetRowItemData(grdFacility.ActiveRow)
    clsShedGrid.ApptDate = udtAppointment.SchedDate
    clsShedGrid.ApptDate = Format(clsShedGrid.ApptDate, "mm/dd/yyyy")
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    If lngProvID > 0 Then
        strSQL = "Get_Provider_Appt_Facility  " & lngProvID & ", " & udtFacility.FacID & ", '" & clsShedGrid.ApptDate & "'"
    Else
        strSQL = "Get_Appointments_Facility " & udtFacility.FacID & ", '" & clsShedGrid.ApptDate & "'"
    End If
    Set rsReadAppt = cnReadAppt.Execute(strSQL)
    With rsReadAppt
        Do Until .EOF
            Clear_SchedPatient
            Clear_Appt
            udtAppointment.SchedType = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) > vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PATIENT_ID"
                                udtSchedPatient.IntId = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtAppointment.VisitType = Fld.Value
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
                                udtAppointment.SchedTime = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                            Case "DURATION"
                                udtSchedPatient.Duration = Fld.Value
                            Case "PATIENT_TYPE"
                                udtAppointment.PatientType = Fld.Value
                            Case "APPT_ID"
                                udtAppointment.Appt_ID = Fld.Value
                        End Select
                    End If
                End If
            Next
        .MoveNext
        lngFoundID = False
        If Not udtSchedPatient.Status = "No Show" Or udtSchedPatient.Status = "Cancelled" Then
            If Trim(udtAppointment.SchedType) <> vbNullString Then
                Arr = Split(LookUpHeader(udtAppointment.SchedTime), "|")
                If UBound(Arr) > 0 Then
                    clsShedGrid.Row = Arr(1)
                End If
                udtAppointment.SchedDate = clsShedGrid.ApptDate
                Select Case IsProperCase(udtAppointment.SchedType)
                    Case "Patient"
                        If Trim(udtSchedPatient.IntId) <> vbNullString Then
                            Get_SchedPatient Val(udtSchedPatient.IntId)
                            mlngPatCnt = mlngPatCnt + 1
                        End If
                    Case "Vendor", "Business"
                        Get_SchedVendor Val(udtSchedPatient.IntId)
                    Case "Legal", "Attorney"
                        Get_Schedule_Attorney Val(udtSchedPatient.IntId)
                    Case "Physician"
                        Get_SchedPhysician Val(udtSchedPatient.IntId)
                End Select
                Select Case UCase(clsApptPref.SchedInterval)
                    Case "EVERY 30 MIN"
                        clsShedGrid.Duration = 30
                    Case "EVERY 20 MIN"
                        clsShedGrid.Duration = 20
                    Case "EVERY 15 MIN"
                        clsShedGrid.Duration = 15
                    Case "EVERY 10 MIN"
                        clsShedGrid.Duration = 10
                    Case "EVERY 5 MIN"
                        clsShedGrid.Duration = 5
                End Select
                If udtSchedPatient.Duration = "1 Hour" Then
                    udtSchedPatient.Duration = "60 minutes"
                End If
                Arr = Split(udtSchedPatient.Duration, " ")
                If UBound(Arr) > 0 Then
                   If IsNumeric(Arr(0)) Then clsShedGrid.SpanRow = Val(Arr(0) / Val(clsShedGrid.Duration))
                Else
                   If IsNumeric(udtSchedPatient.Duration) Then clsShedGrid.SpanRow = Val(udtSchedPatient.Duration / Val(clsShedGrid.Duration))
                End If
                udtSchedPatient.Name = Replace(udtSchedPatient.Name, "''", "'")
                If clsShedGrid.Row > 0 Then
                    For lngCol = 1 To grdData.MaxCols
                        If grdData.GetColItemData(lngCol) = udtAppointment.ProviderID Then
                            clsShedGrid.ProvCol = lngCol
                            Exit For
                        End If
                     Next lngCol
                     lngCol = 0
                     For lngCol = clsShedGrid.ProvCol To grdData.MaxCols
                        If grdData.GetColItemData(lngCol) = udtAppointment.ProviderID Then
                            If lngCol > clsShedGrid.ProvCol Then
                                clsShedGrid.ProvCol2 = lngCol
                                Exit For
                            End If
                        End If
                     Next lngCol
                    For lngCol = clsShedGrid.ProvCol To clsShedGrid.ProvCol2
                        If Not lngFoundID Then
                     '       For lngRow = clsShedGrid.Row - 1 To Val(clsShedGrid.Row + clsShedGrid.SpanRow)
                           For lngRow = clsShedGrid.Row To Val(clsShedGrid.Row + (clsShedGrid.SpanRow - 1))
                                    grdData.Col = lngCol
                                    grdData.Row = lngRow
                                    If grdData.CellTag <> vbNullString Then
                                        If grdData.CellTag <> udtSchedPatient.IntId Then
                                            grdData.GetText lngCol, grdData.Row, varCellValue
                                            clsGridFormat.CellNote = vbNullString
                                            Arr = Split(varCellValue, "INSURANCE AUTHORIZATION:")
                                            If UBound(Arr) > 0 Then
                                               varCellValue = Trim(Arr(0))
                                            End If
                                            clsGridFormat.CellNote = Trim(varCellValue)
                                            Arr = Split(udtSchedPatient.Collection, "INSURANCE AUTHORIZATION:")
                                            If UBound(Arr) > 0 Then
                                               clsGridFormat.CellNote = Trim(clsGridFormat.CellNote) & vbNewLine & Arr(0)
                                            End If
                                            If Trim(clsGridFormat.CellNote) <> vbNullString Then
                                               clsGridFormat.CellNote = "Scheduling Conflict" & vbNewLine & Trim(clsGridFormat.CellNote)
                                            End If
                                            If Trim(clsGridFormat.CellNote) <> vbNullString And _
                                               Not TextFound(Trim(clsGridFormat.CellNote), Trim(rtfChanges.Text), Len(Trim(rtfChanges.Text)), True) Then
                                               If Trim(rtfChanges.Text) <> vbNullString Then
                                                   rtfChanges.Text = rtfChanges.Text & vbNewLine & clsGridFormat.CellNote
                                               Else
                                                   rtfChanges.Text = Trim(clsGridFormat.CellNote)
                                               End If
                                            End If
                                            clsShedGrid.ProvCol2 = clsShedGrid.ProvCol + 1
                                            Add_SchedRow clsShedGrid.ProvCol2
                                            lngFoundID = True
                                            Exit For
                                        Else
                                           lngFoundID = True
                                           Exit For
                                        End If
                                    Else
                                        lngFoundID = False
                                        Exit For
                                    End If
                            Next lngRow
                        End If
                        If Not lngFoundID Then
                            Add_SchedRow clsShedGrid.ProvCol
                            Exit For
                        End If
                    Next lngCol
                    End If
                End If
            End If
        'End If
        Loop
    End With
    If cnReadAppt.State = adStateClosed Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    udtAppointment.SchedDate = clsShedGrid.ApptDate
    mbooByPass = False
    If grdData.MaxCols > 1 Then
        With grdData
            grdData.Row = SpreadHeader
            For lngCol = 1 To .MaxCols
                lngProvID = .GetColItemData(lngCol)
                Set_BlockTime lngProvID
            Next lngCol
        End With
    Else
        Set_BlockTime udtAppointment.ProviderID
    End If
    If mlngGridRow > 0 Then
       grdData.TopRow = mlngGridRow
    End If
    If mBeenHere Then mBeenHere = False: GoTo Exit_Function
    If grdData.Row >= 1 Then InfoFrame grdData.Col, grdData.Row
    mFirstLoad = False
    mbooByPass = False
    Reset_SelProvider
    schTimer.Enabled = True
    udtAppointment.SchedType = "Patient"
    Get_Changes clsShedGrid.ApptDate
    lblCurScheAppt.Visible = True
    lblCurScheAppt.Caption = "Current Patient Appointments"
    lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
    GoTo Exit_Function
Error_ReadData_NameOnly:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "ReadData_NameOnly", App.EXEName & "." & Me.Name
Exit_Function:
    Screen.MousePointer = vbDefault
    grdData.Visible = True
    Exit Sub
    Resume Next
End Sub

Private Sub Read_Week_Appointments()
Dim cnReadAppt As New ADODB.Connection
Dim rsReadAppt  As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSQL As String
Dim Arr
Dim lngPatID As Long
Dim Idx As Integer
Dim lngRow As Long
Dim varCellValue
Dim vTime
Dim i%
Dim lngCol As Long
Dim lngFoundID As Boolean
Dim lngProvID As Long
Dim strNotes As String
Dim bPatFound As Boolean
Dim lngApptID As Long
Dim lngTopRow As Long
Dim lngTest As Long
Dim lngCnt As Long
Dim bExit As Boolean
Dim strDate As String
Dim strEndDate As String
Dim lngChgCol As Long
Dim lngCurFac As Long


    On Error GoTo Error_Read_Week_Appointments
    Screen.MousePointer = vbHourglass
    mlngPatCnt = 0
    Get_Practice
    grdData.ClearRange 1, 1, grdData.MaxCols, grdData.MaxRows, False
    For lngCol = 1 To grdData.MaxCols
        For lngRow = 1 To grdData.MaxRows
            grdData.RowHeight(lngRow) = 15 * mlngSystemFileSize
            grdData.RemoveCellSpan lngCol, lngRow
            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Next lngRow
    Next lngCol
    lngRow = 0
    If mstrApptMode = vbNullString Then mstrApptMode = "Weekly"
    If udtAppointment.ProviderID = 0 Then
        With grdProviders
           For lngTest = 1 To .MaxRows
               .Col = 1
               .Row = lngTest
               If .BackColor = grdCurItem Then
                   udtAppointment.ProviderID = .GetRowItemData(lngTest)
                    Get_Providers udtAppointment.ProviderID
                    cctProvider.Text = clsProvider.Name
                    Exit For
               End If
           Next lngTest
        End With
    End If
    If udtFacility.FacID = 0 Then
        With grdFacility
           For lngTest = 1 To .MaxRows
               .Col = 1
               .Row = lngTest
               If .BackColor = grdCurItem Then
                   udtFacility.FacID = .GetRowItemData(lngTest)
                    Exit For
               End If
           Next lngTest
        End With
    End If
    lngCnt = grdData.MaxCols
    grdData.GetText 1, SpreadHeader + 2, varCellValue
    Arr = Split(varCellValue, "-")
    If UBound(Arr) > 0 Then
        If IsDate(Trim(Arr(1))) Then strDate = Trim(Arr(1))
        strDate = Format(strDate, "mm/dd/yyyy")
    Else
        strDate = Format(varCellValue, "mm/dd/yyyy")
    End If
    lngCurFac = udtFacility.FacID
    strEndDate = DateAdd("d", 6, strDate)
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "SELECT  * FROM APPOINTMENTS WHERE  CONVERT(DATETIME, APPT_DATE,101) " _
        & "BETWEEN '" & Format(strDate, "mm/dd/yyyy") & "' AND '" & Format(strEndDate, "mm/dd/yyyy") & "' AND " _
        & "PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
        & "ACTIVE = 1 ORDER BY CONVERT(DATETIME, APPT_DATE, 101), CONVERT(DATETIME, APPT_TIME, 108)"
       Set rsReadAppt = cnReadAppt.Execute(strSQL)
         With rsReadAppt
             Do Until .EOF
                Clear_SchedPatient
                Clear_Appt
                lngApptID = 0
                udtAppointment.SchedType = vbNullString
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                       If Trim(Fld.Value) > vbNullString Then
                          Select Case UCase(Fld.Name)
                             Case "PATIENT_ID"
                                 udtSchedPatient.IntId = Fld.Value
                             Case "PURPOSE_VISIT"
                                 udtAppointment.VisitType = Fld.Value
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
                                 udtAppointment.SchedTime = Format(udtAppointment.SchedTime, "hh:mm AM/PM")
                             Case "DURATION"
                                 udtSchedPatient.Duration = Fld.Value
                             Case "APPT_ID"
                                 lngApptID = Fld.Value
                             Case "PATIENT_TYPE"
                                 udtAppointment.PatientType = Fld.Value
                             Case "APPT_DATE"
                                 strDate = Format(Fld.Value, "mm/dd/yyyy")
                          End Select
                       End If
                    End If
                 Next
                 .MoveNext
                 If Not Valid_RepeatAppt(lngApptID) Then
                     lngFoundID = False
                     If udtSchedPatient.Status <> "No Show" And udtSchedPatient.Status <> "Cancelled" Then
                         If Trim(udtAppointment.SchedType) <> vbNullString Then
                             Arr = Split(LookUpHeader(udtAppointment.SchedTime), "|")
                             If UBound(Arr) > 0 Then
                                 clsShedGrid.Row = Arr(1)
                             End If
                             If lngTopRow = 0 Then lngTopRow = clsShedGrid.Row
                             udtAppointment.SchedDate = clsShedGrid.ApptDate
                             Select Case IsProperCase(udtAppointment.SchedType)
                                 Case "Patient"
                                     If Trim(udtSchedPatient.IntId) <> vbNullString Then
                                         Get_SchedPatient Val(udtSchedPatient.IntId), strDate
                                     End If
                                 Case "Vendor", "Business"
                                     Get_SchedVendor Val(udtSchedPatient.IntId)
                                 Case "Legal", "Attorney"
                                     Get_Schedule_Attorney Val(udtSchedPatient.IntId)
                                 Case "Physician"
                                     Get_SchedPhysician Val(udtSchedPatient.IntId)
                             End Select
                             Select Case UCase(clsApptPref.SchedInterval)
                                 Case "EVERY 30 MIN"
                                     clsShedGrid.Duration = 30
                                 Case "EVERY 20 MIN"
                                     clsShedGrid.Duration = 20
                                 Case "EVERY 15 MIN"
                                     clsShedGrid.Duration = 15
                                 Case "EVERY 10 MIN"
                                     clsShedGrid.Duration = 10
                                 Case "EVERY 5 MIN"
                                     clsShedGrid.Duration = 5
                             End Select
                             If udtSchedPatient.Duration = "1 Hour" Then
                                 udtSchedPatient.Duration = "60 minutes"
                             End If
                             Arr = Split(udtSchedPatient.Duration, " ")
                             If UBound(Arr) > 0 Then
                                If IsNumeric(Arr(0)) Then clsShedGrid.SpanRow = Val(Arr(0) / Val(clsShedGrid.Duration))
                             Else
                                If IsNumeric(udtSchedPatient.Duration) Then clsShedGrid.SpanRow = Val(udtSchedPatient.Duration / Val(clsShedGrid.Duration))
                             End If
                             Get_Facilities udtAppointment.FacilityID
                             udtSchedPatient.Name = Replace(udtSchedPatient.Name, "''", "'")
                             If clsShedGrid.Row > 0 Then
                                clsShedGrid.ProvCol = 0
                                Look_Up_Column_Date_Week (strDate)
                                grdData.GetText clsShedGrid.ProvCol, SpreadHeader + 4, varCellValue
                                If Trim(varCellValue) = vbNullString Then
                                    If clsShedGrid.ProvCol2 = 0 Then
                                        Format_CellTypeWeekCol grdData, SpreadHeader + 4, clsShedGrid.ProvCol, "1", clsFacility.Name, "TypeHAlignCenter", CStr(udtAppointment.ProviderID), HeaderColor, udtAppointment.FacilityID
                                        grdData.SetCellBorder clsShedGrid.ProvCol, SpreadHeader + 4, clsShedGrid.ProvCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    Else
                                        Format_CellTypeWeekCol grdData, SpreadHeader + 4, clsShedGrid.ProvCol, "2", clsFacility.Name, "TypeHAlignCenter", CStr(udtAppointment.ProviderID), HeaderColor, udtAppointment.FacilityID
                                        grdData.SetCellBorder clsShedGrid.ProvCol, SpreadHeader + 4, clsShedGrid.ProvCol2, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                    End If
                                Else
                                    Debug.Print
                                End If
                                 If clsShedGrid.ProvCol2 > 0 Then
                                     For lngCol = clsShedGrid.ProvCol To clsShedGrid.ProvCol2
                                         If Not lngFoundID Then
                                            For lngRow = clsShedGrid.Row To Val(clsShedGrid.Row + (clsShedGrid.SpanRow - 1))
                                                     grdData.Col = lngCol
                                                     grdData.Row = lngRow
                                                     If grdData.CellTag <> vbNullString Then
                                                         If grdData.CellTag <> udtSchedPatient.IntId Then
                                                             grdData.GetText lngCol, grdData.Row, varCellValue
                                                             clsGridFormat.CellNote = vbNullString
                                                             Arr = Split(varCellValue, "INSURANCE AUTHORIZATION:")
                                                             If UBound(Arr) > 0 Then
                                                                varCellValue = Trim(Arr(0))
                                                             End If
                                                             clsGridFormat.CellNote = Trim(varCellValue)
                                                             Arr = Split(udtSchedPatient.Collection, "INSURANCE AUTHORIZATION:")
                                                             If UBound(Arr) > 0 Then
                                                                clsGridFormat.CellNote = Trim(clsGridFormat.CellNote) & vbNewLine & Arr(0)
                                                             End If
                                                             If Trim(clsGridFormat.CellNote) <> vbNullString Then
                                                                clsGridFormat.CellNote = "Scheduling Conflict" & vbNewLine & Trim(clsGridFormat.CellNote)
                                                             End If
                                                             If Trim(clsGridFormat.CellNote) <> vbNullString And Not IsDate(grdData.CellTag) Then
                                                                 If Trim(clsGridFormat.CellNote) <> vbNullString Then
                                                                     bExit = False
                                                                     With grdChanges
                                                                         For lngTest = 1 To .MaxRows
                                                                             .Col = 1
                                                                             .Row = lngTest
                                                                             If .GetRowItemData(lngTest) = Val(udtSchedPatient.IntId) Then
                                                                                 bExit = True
                                                                                 Exit For
                                                                             End If
                                                                         Next lngTest
                                                                     End With
                                                                 End If
                                                             End If
                                                            Look_Up_Column_Date_Week (strDate), True
                                                             Add_SchedRow clsShedGrid.ProvCol2
                                                             lngFoundID = True
                                                             Exit For
                                                         Else
                                                            lngFoundID = True
                                                            Exit For
                                                         End If
                                                     Else
                                                         lngFoundID = False
                                                         Exit For
                                                     End If
                                             Next lngRow
                                         End If
                                         If Not lngFoundID Then
                                             Add_SchedRow clsShedGrid.ProvCol
                                             Exit For
                                         End If
                                     Next lngCol
                                 Else
                                    If Not lngFoundID Then
                                       For lngRow = clsShedGrid.Row To Val(clsShedGrid.Row + (clsShedGrid.SpanRow - 1))
                                            grdData.Col = clsShedGrid.ProvCol
                                            grdData.Row = lngRow
                                            If grdData.CellTag <> vbNullString Then
                                                If grdData.CellTag <> udtSchedPatient.IntId Then
                                                    grdData.GetText clsShedGrid.ProvCol, grdData.Row, varCellValue
                                                    clsGridFormat.CellNote = vbNullString
                                                    Arr = Split(varCellValue, "INSURANCE AUTHORIZATION:")
                                                    If UBound(Arr) > 0 Then
                                                       varCellValue = Trim(Arr(0))
                                                    End If
                                                    clsGridFormat.CellNote = Trim(varCellValue)
                                                    Arr = Split(udtSchedPatient.Collection, "INSURANCE AUTHORIZATION:")
                                                    If UBound(Arr) > 0 Then
                                                       clsGridFormat.CellNote = Trim(clsGridFormat.CellNote) & vbNewLine & Arr(0)
                                                    End If
                                                    If Trim(clsGridFormat.CellNote) <> vbNullString Then
                                                       clsGridFormat.CellNote = "Scheduling Conflict" & vbNewLine & Trim(clsGridFormat.CellNote)
                                                    End If
                                                    If Trim(clsGridFormat.CellNote) <> vbNullString And Not IsDate(grdData.CellTag) Then
                                                        If Trim(clsGridFormat.CellNote) <> vbNullString Then
                                                            bExit = False
                                                            With grdChanges
                                                                For lngTest = 1 To .MaxRows
                                                                    .Col = 1
                                                                    .Row = lngTest
                                                                    If .GetRowItemData(lngTest) = Val(udtSchedPatient.IntId) Then
                                                                        bExit = True
                                                                        Exit For
                                                                    End If
                                                                Next lngTest
                                                                If Not bExit Then
                                                                    clsGridFormat.CellNote = Replace(clsGridFormat.CellNote, Trim(udtSchedPatient.Name), "")
                                                                    lngTest = .MaxRows + 1
                                                                    .MaxRows = lngTest
                                                                    .SetRowItemData lngTest, Val(udtSchedPatient.IntId)
                                                                    lngChgCol = 1
                                                                    Format_CellTypeStatic grdChanges, lngTest, lngChgCol, "1", Trim(udtSchedPatient.Name), "TypeHAlignCenter"
                                                                    grdData.SetCellBorder lngChgCol, lngTest, lngChgCol, lngTest, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                                                    lngChgCol = 2
                                                                    Format_CellTypeStatic grdChanges, lngTest, lngChgCol, "1", Trim(clsGridFormat.CellNote), "TypeHAlignCenter"
                                                                    grdChanges.SetCellBorder lngChgCol, lngTest, lngChgCol, lngTest, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                                                                End If
                                                            End With
                                                        End If
                                                    End If
                                                    Look_Up_Column_Date_Week (strDate), True
                                                    Add_SchedRow clsShedGrid.ProvCol2
                                                    lngFoundID = True
                                                    Exit For
                                                Else
                                                   lngFoundID = True
                                                   Exit For
                                                End If
                                            Else
                                                lngFoundID = False
                                                Exit For
                                            End If
                                        Next lngRow
                                    End If
                                    If Not lngFoundID Then
                                        Add_SchedRow clsShedGrid.ProvCol
                                    End If
                                 End If
                             End If
                         End If
                     End If
                 End If
             Loop
         End With
    udtAppointment.SchedDate = strDate
    lngCurFac = udtFacility.FacID
    Set_Week_Blocks strDate, strEndDate
    If cnReadAppt.State = adStateOpen Then
        cnReadAppt.Close
        Set cnReadAppt = Nothing
        Set rsReadAppt = Nothing
        Set Flds = Nothing
    End If
    Add_Spacer SpreadHeader + 1, grdData
    Add_Spacer SpreadHeader + 5, grdData
    mbooByPass = False
    If mBeenHere Then
       mBeenHere = False
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    mlngPatCnt = 0
    grdData.GetText 1, SpreadHeader + 2, varCellValue
    Arr = Split(varCellValue, "-")
    If UBound(Arr) > 0 Then
        If IsDate(Trim(Arr(1))) Then strDate = Trim(Arr(1))
        strDate = Format(strDate, "mm/dd/yyyy")
    Else
        strDate = Format(varCellValue, "mm/dd/yyyy")
    End If
    strEndDate = DateAdd("d", 6, strDate)
    If cnReadAppt.State = adStateClosed Then cnReadAppt.open psConnect(0)
    strSQL = "SELECT  DISTINCT PATIENT_ID FROM APPOINTMENTS WHERE  CONVERT(DATETIME, APPT_DATE,101) " _
        & "BETWEEN '" & Format(strDate, "mm/dd/yyyy") & "' AND '" & Format(strEndDate, "mm/dd/yyyy") & "' AND " _
        & "PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
        & "ACTIVE = 1 AND STATUS <> 'No Show'"
        Set rsReadAppt = cnReadAppt.Execute(strSQL)
        With rsReadAppt
            Do Until .EOF
                .MoveNext
                mlngPatCnt = mlngPatCnt + 1
            Loop
        End With
    mBeenHere = True
    lblSchedChange.Visible = False
    udtAppointment.SchedType = "Patient"
    lblCurScheAppt.Visible = True
    lblCurScheAppt.Caption = "Current Patient Appointments"
    lblCurScheAppt.Caption = lblCurScheAppt.Caption & "     " & mlngPatCnt
    udtAppointment.FacilityID = lngCurFac
    udtFacility.FacID = lngCurFac
    Get_Facilities lngCurFac
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Read_Week_Appointments:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL & CStr(lngRow), "Read_Week_Appointments", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub
Private Function Set_Week_Blocks(vstrStart As String, vstrEnd As String)
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim i%
Dim lngDateDiff As Long
Dim strRepDate As String
Dim booSkip As Boolean
Dim strStartRange As String
Dim strEndRange As String
Dim strTestTime As String
Dim strTestDay As String
Dim strWeek As String

    On Error GoTo Error_Set_Week_Blocks
    Screen.MousePointer = vbHourglass
    Set_Prem_Block_Week
    If udtAppointment.SchedDate = vbNullString Then Exit Function
    If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
    strSQL = "SELECT * FROM APPOINTMENTS_BLOCKS WHERE " _
        & "PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
        & "(CONVERT(DATETIME, START_DATE, 101) BETWEEN '" & Format(vstrStart, "mm/dd/yyyy") & "' AND '" & Format(vstrStart, "mm/dd/yyyy") & "' OR " _
        & "CONVERT(DATETIME, END_DATE, 101) BETWEEN '" & Format(vstrEnd, "mm/dd/yyyy") & "'  AND  '" & Format(vstrEnd, "mm/dd/yyyy") & "') AND " _
        & "ACTIVE = 1  ORDER BY CONVERT(DATETIME,START_DATE, 101) DESC, REASON, CONVERT(DATETIME,START_TIME, 108) DESC, FACILITY_ID"
    Set rsBlock = cnBlock.Execute(strSQL)
    Do Until rsBlock.EOF
        Clear_Block
        booSkip = False
        strRepDate = vbNullString
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
        If strStartRange <> vbNullString Then
            If DateDiff("n", Format(strStartRange, "hh:mm"), Format(udtBlock.StartTime, "hh:mm")) > 0 Then
                booSkip = True
            ElseIf DateDiff("n", Format(strEndRange, "hh:mm"), Format(udtBlock.StartTime, "hh:mm")) > 0 Then
                booSkip = True
            End If
        End If
        If Not booSkip Then
            lngDateDiff = 0
            If Trim(udtBlock.DayWeek) <> vbNullString Then strRepDate = DayWeek(udtBlock.DayWeek)
            If IsDate(udtBlock.StartDate) And (Trim(udtBlock.DayWeek) = vbNullString Or UCase(udtBlock.DayWeek) = GetDay(udtAppointment.SchedDate)) Then
               If Not IsDate(strRepDate) Then strRepDate = udtBlock.EndDate
               If udtBlock.Duration = "Permanent" Then
                  lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), udtBlock.StartDate)
                  If lngDateDiff <= 0 Then
                     lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), strRepDate)
                     If lngDateDiff >= 0 Then
                           WeekMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, udtBlock.StartDate, udtBlock.ProviderID
                     End If
                  End If
               Else
                  lngDateDiff = DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate)
                  If lngDateDiff >= 0 Then
                     lngDateDiff = DateDiff("d", udtAppointment.SchedDate, udtBlock.EndDate)
                     If lngDateDiff >= 0 Then
                        WeekMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, udtAppointment.SchedDate, udtBlock.ProviderID
                        If DateDiff("h", Format(udtBlock.StartTime, "hh:mm"), Format(udtBlock.EndTime, "hh:mm")) > 7 And _
                            DateDiff("d", Format(udtBlock.StartDate, "mm/dd/yyyy"), Format(udtBlock.EndDate, "mm/dd/yyyy")) >= 0 Then
                            Exit Do
                        End If
                     End If
                  End If
               End If
            End If
            If udtBlock.Reason = "Holiday" Then
                Exit Do
            End If
            If strTestDay <> udtBlock.DayWeek Then
                strStartRange = vbNullString
                strEndRange = vbNullString
            End If
            lngDateDiff = DateDiff("h", udtBlock.StartTime, udtBlock.EndTime)
            If lngDateDiff > 7 Then
                strStartRange = udtBlock.StartTime
                strEndRange = udtBlock.EndTime
            End If
            strTestDay = udtBlock.DayWeek
        End If
    Loop
    If cnBlock.State = adStateOpen Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_Week_Blocks:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Set_Week_Blocks", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Set_Prem_Block_Week()
Dim strSQL As String
Dim cnBlock As New ADODB.Connection
Dim rsBlock As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim lngPrimKey As Long
Dim i%
Dim lngDateDiff As Long
Dim strRepDate As String
Dim booSkip As Boolean
Dim strStartRange As String
Dim strEndRange As String
Dim strTestTime As String
Dim strTestDay As String
Dim strWeek As String

    On Error GoTo Error_Set_Prem_Block_Week
    Screen.MousePointer = vbHourglass
    If udtAppointment.SchedDate = vbNullString Then Exit Function
    strTestTime = DateAdd("d", 1, udtAppointment.SchedDate)
    strEndRange = DateAdd("d", -1, udtAppointment.SchedDate)
    strWeek = DateAdd("d", -30, udtAppointment.SchedDate)
    If cnBlock.State = adStateClosed Then cnBlock.open psConnect(1)
    strSQL = "SELECT * FROM APPOINTMENTS_BLOCKS WHERE ACTIVE = 1 " _
        & "AND PROVIDER_ID = " & udtAppointment.ProviderID & " AND " _
        & "DURATION = 'Permanent'  AND " _
        & "CONVERT(DATETIME, END_DATE, 101) > '" & Format(strEndRange, "mm/dd/yyyy") & "' ORDER " _
        & "BY CONVERT(DATETIME, START_DATE, 101) DESC, REASON, CONVERT(DATETIME, START_TIME, 108) DESC"
    Set rsBlock = cnBlock.Execute(strSQL)
    Do Until rsBlock.EOF
        Clear_Block
        booSkip = False
        strRepDate = vbNullString
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
        lngDateDiff = 0
        If Trim(udtBlock.DayWeek) <> vbNullString Then strRepDate = DayWeek(udtBlock.DayWeek)
        If IsDate(udtBlock.StartDate) And (Trim(udtBlock.DayWeek) = vbNullString Or UCase(udtBlock.DayWeek) = GetDay(udtAppointment.SchedDate)) Then
           If Not IsDate(strRepDate) Then strRepDate = udtBlock.EndDate
           If udtBlock.Duration = "Permanent" Then
              lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), udtBlock.StartDate)
              If lngDateDiff <= 0 Then
                 lngDateDiff = DateDiff("d", Format(udtAppointment.SchedDate, "mm/dd/yyyy"), strRepDate)
                 If lngDateDiff >= 0 Then
                    WeekMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, udtBlock.StartDate, udtBlock.ProviderID
                 End If
              End If
           Else
              lngDateDiff = DateDiff("d", udtBlock.StartDate, udtAppointment.SchedDate)
              If lngDateDiff >= 0 Then
                 lngDateDiff = DateDiff("d", udtAppointment.SchedDate, udtBlock.EndDate)
                 If lngDateDiff >= 0 Then
                    WeekMode_BlockTime udtBlock.Reason, udtBlock.EndTime, udtBlock.StartTime, udtBlock.Duration, udtAppointment.SchedDate, udtBlock.ProviderID
                 End If
              End If
           End If
        End If
        If udtBlock.Reason = "Holiday" Then
            Exit Do
        End If
        If strTestDay <> udtBlock.DayWeek Then
            strStartRange = vbNullString
            strEndRange = vbNullString
        End If
        lngDateDiff = DateDiff("h", udtBlock.StartTime, udtBlock.EndTime)
        If lngDateDiff > 7 Then
            strStartRange = udtBlock.StartTime
            strEndRange = udtBlock.EndTime
        End If
        strTestDay = udtBlock.DayWeek
    Loop
    If cnBlock.State = adStateOpen Then
        cnBlock.Close
        Set cnBlock = Nothing
        Set rsBlock = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_Prem_Block_Week:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Set_Prem_Block_Week", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function


Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Dim strDay As String
Dim Arr
Dim i%

    If mbooByPass Then mbooByPass = False: Exit Sub
    strSelDate = MonthView1.Value
    If strSelDate <> vbNullString Then
        udtAppointment.SchedDate = Format(strSelDate, "MMMM DD, YYYY")
        lblApptDate.Caption = Format(strSelDate, "MMMM DD, YYYY")
        Get_Mode
    End If
   Exit Sub
End Sub

Private Sub picBlock_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picBlock.BackColor = SColor.oSelectedColor
        udtColor.Lunch = picBlock.BackColor
       Get_Mode
       Write_Color "Lunch", udtColor.Lunch
    End If
    Exit Sub
End Sub

Private Sub picCancel_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picCancel.BackColor = SColor.oSelectedColor
        udtColor.Cancelled = picCancel.BackColor
       Get_Mode
       Write_Color "Cancelled", udtColor.Cancelled
    End If
    Exit Sub
End Sub


Private Sub picCheckOut_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picCheckOut.BackColor = SColor.oSelectedColor
        udtColor.CheckOut = picCheckOut.BackColor
       Get_Mode
       Write_Color "CheckOut", udtColor.CheckOut
    End If
    Exit Sub
End Sub

Private Sub picExstNewPat_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picExstNewPat.BackColor = SColor.oSelectedColor
        udtColor.ExistPatient = picExstNewPat.BackColor
       Get_Mode
       Write_Color "ExistPatient", udtColor.ExistPatient
    End If
    Exit Sub
End Sub

Private Sub picHospital_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picHospital.BackColor = SColor.oSelectedColor
        udtColor.Hospital = picHospital.BackColor
       Get_Mode
       Write_Color "Hospital", udtColor.Hospital
    End If
    Exit Sub
End Sub

Private Sub picNewPatient_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picNewPatient.BackColor = SColor.oSelectedColor
        udtColor.NewPatient = picNewPatient.BackColor
       Get_Mode
       Write_Color "NewPatient", udtColor.NewPatient
    End If
    Exit Sub
End Sub

Private Sub picOffice_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picOffice.BackColor = SColor.oSelectedColor
        udtColor.Office = picOffice.BackColor
       Get_Mode
       Write_Color "Office", udtColor.Office
    End If
    Exit Sub
End Sub

Private Sub picPostOp_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picPostOp.BackColor = SColor.oSelectedColor
        udtColor.PostOp = picPostOp.BackColor
       Get_Mode
       Write_Color "PostOp", udtColor.PostOp
    End If
    Exit Sub
End Sub

Private Sub picPreOp_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picPreOp.BackColor = SColor.oSelectedColor
        udtColor.PreOp = picPreOp.BackColor
       Get_Mode
       Write_Color "PreOp", udtColor.PreOp
    End If
    Exit Sub
End Sub

Private Sub picPresent_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picPresent.BackColor = SColor.oSelectedColor
        udtColor.Present = picPresent.BackColor
       Get_Mode
       Write_Color "Present", udtColor.Present
    End If
    Exit Sub
End Sub

Private Sub picReturn_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picReturn.BackColor = SColor.oSelectedColor
        udtColor.FollowUp = picReturn.BackColor
       Get_Mode
       Write_Color "FollowUp", udtColor.FollowUp
    End If
    Exit Sub
End Sub

Private Sub picSurgery_Click()
Dim SColor As SelectedColor
    SColor = ShowColor(Me.Hwnd)
    If Not SColor.bCanceled Then
        picSurgery.BackColor = SColor.oSelectedColor
        udtColor.Surgery = picSurgery.BackColor
       Get_Mode
       Write_Color "Surgery", udtColor.Surgery
    End If
    Exit Sub
End Sub

Private Sub schTimer_Timer()
    
    If schTimer.Enabled Then
        Check_Change "Appointment"
        If mbooSchChange Then
            Get_Mode
            Reset_SelProvider
            Get_Changes udtAppointment.SchedDate
        End If
    End If
    Exit Sub
End Sub

Private Sub strTelTime_Timer()
    Check_Change "Telephone"
    If mbooTeleChange Then
        Format_Messages
        Read_Messages
    End If
    Exit Sub
End Sub

