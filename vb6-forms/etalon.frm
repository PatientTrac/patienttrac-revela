VERSION 5.00
Object = "{76F358A5-0FAA-4EA4-8073-9E9D9A484F37}#1.0#0"; "CIMAGE~1.OCX"
Begin VB.Form etalon 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   Caption         =   "Etalon for 3 projects"
   ClientHeight    =   9225
   ClientLeft      =   6525
   ClientTop       =   2715
   ClientWidth     =   16455
   FillColor       =   &H00C0E0FF&
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   16455
   Begin CIMAGEBUFFERXLib.CImageBufferX CImageBufferX1 
      Height          =   945
      Left            =   9180
      TabIndex        =   30
      Top             =   1950
      Width           =   1635
      _Version        =   65536
      _ExtentX        =   2884
      _ExtentY        =   1667
      _StockProps     =   0
   End
   Begin VB.Frame fraTB 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   1020
      TabIndex        =   0
      Top             =   4860
      Width           =   13305
   End
   Begin VB.PictureBox Picture8 
      Height          =   795
      Left            =   1800
      ScaleHeight     =   735
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   2250
      Width           =   735
   End
   Begin VB.PictureBox picExplain 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      Picture         =   "etalon.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   840
      Width           =   555
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   690
      Picture         =   "etalon.frx":08CA
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   2040
      Width           =   315
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   360
      Picture         =   "etalon.frx":0E54
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   1770
      Width           =   315
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      Picture         =   "etalon.frx":13DE
      ScaleHeight     =   315
      ScaleWidth      =   345
      TabIndex        =   6
      Top             =   1530
      Width           =   345
   End
   Begin VB.PictureBox picSmile 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1200
      Picture         =   "etalon.frx":1968
      ScaleHeight     =   315
      ScaleWidth      =   345
      TabIndex        =   7
      Top             =   3150
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picRec 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   30
      Picture         =   "etalon.frx":1EF2
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   3570
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picRound 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   300
      Picture         =   "etalon.frx":247C
      ScaleHeight     =   315
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   3750
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSurgName 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   210
      Picture         =   "etalon.frx":2A06
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   3120
      Width           =   315
   End
   Begin VB.PictureBox picTime 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      Picture         =   "etalon.frx":2F90
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   16
      Top             =   2280
      Width           =   525
   End
   Begin VB.PictureBox picCut 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   960
      Picture         =   "etalon.frx":385A
      ScaleHeight     =   675
      ScaleWidth      =   585
      TabIndex        =   18
      Top             =   3750
      Width           =   585
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   600
      Picture         =   "etalon.frx":4124
      ScaleHeight     =   585
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   5700
      Width           =   555
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   1230
      Picture         =   "etalon.frx":49EE
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   22
      Top             =   5190
      Width           =   585
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   600
      Picture         =   "etalon.frx":52B8
      ScaleHeight     =   555
      ScaleWidth      =   585
      TabIndex        =   23
      Top             =   4830
      Width           =   585
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      Picture         =   "etalon.frx":5B82
      ScaleHeight     =   585
      ScaleWidth      =   555
      TabIndex        =   24
      Top             =   4530
      Width           =   555
   End
   Begin VB.Frame fraWhatDone 
      BackColor       =   &H00E0E0E0&
      Caption         =   "What has been done?"
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
      Height          =   2415
      Left            =   1800
      TabIndex        =   14
      Top             =   3780
      Width           =   4695
      Begin VB.PictureBox picMed 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   1995
         TabIndex        =   21
         Top             =   270
         Width           =   1995
      End
      Begin VB.PictureBox picSurgery 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   1185
         TabIndex        =   19
         Top             =   600
         Width           =   1185
      End
      Begin VB.PictureBox picNoth 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   90
         ScaleHeight     =   345
         ScaleWidth      =   1185
         TabIndex        =   17
         Top             =   990
         Width           =   1185
      End
      Begin VB.PictureBox picOther 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   90
         ScaleHeight     =   345
         ScaleWidth      =   1185
         TabIndex        =   15
         Top             =   1380
         Width           =   1185
      End
   End
   Begin VB.Frame fraSurgery 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Who is/was your Doctor?"
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
      Left            =   2070
      TabIndex        =   8
      ToolTipText     =   "Double Click Mouse to Reset Surgeon Name"
      Top             =   6210
      Width           =   4365
      Begin VB.PictureBox picSurg 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   30
         ScaleHeight     =   765
         ScaleWidth      =   4095
         TabIndex        =   25
         Top             =   240
         Width           =   4095
         Begin VB.TextBox txtSurgName 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   28
            Top             =   420
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.OptionButton optSurgRem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "I don't remember their name"
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
            Left            =   60
            TabIndex        =   27
            Top             =   60
            Width           =   3375
         End
         Begin VB.OptionButton optSurgName 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Name "
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
            Height          =   375
            Left            =   60
            TabIndex        =   26
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   4095
      End
   End
   Begin VB.Frame fraApproxExam 
      BackColor       =   &H00FFFFFF&
      Caption         =   "If unsure approximately when?"
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
      Height          =   930
      Left            =   3030
      TabIndex        =   5
      Top             =   1470
      Width           =   3450
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "When?"
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
      Left            =   1110
      TabIndex        =   29
      Top             =   540
      Width           =   765
   End
End
Attribute VB_Name = "etalon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub GIIEPad1_PadData(ByVal abs_x As Long, ByVal abs_y As Long, ByVal abs_z As Long, ByVal fTouch As Boolean, ByVal fTap As Boolean, ByVal bmButtonState As Integer, ByVal dwTimeStamp As Long, ByVal dwSequenceNum As Long)

End Sub


