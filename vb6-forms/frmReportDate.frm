VERSION 5.00
Object = "{2A5B1681-9AFA-4521-B035-B9AD0C75F381}#1.0#0"; "DedicatedControls.ocx"
Begin VB.Form frmReportDate 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Date Range"
   ClientHeight    =   2175
   ClientLeft      =   11850
   ClientTop       =   5130
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4275
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2190
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   1530
      TabIndex        =   2
      Top             =   1680
      Width           =   555
   End
   Begin VB.PictureBox picReportDate 
      Height          =   1280
      Left            =   570
      ScaleHeight     =   1215
      ScaleWidth      =   3105
      TabIndex        =   0
      Top             =   240
      Width           =   3165
      Begin DedicatedControls.ccDate ccToDate 
         Height          =   255
         Left            =   510
         TabIndex        =   4
         Top             =   720
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         Caption         =   "To Date: "
         Text            =   "06/04/2003"
         DateBackColor   =   -2147483643
      End
      Begin DedicatedControls.ccDate ccDate 
         Height          =   255
         Left            =   420
         TabIndex        =   1
         Top             =   210
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         Caption         =   "From Date:"
         Text            =   "06/04/2003"
         DateBackColor   =   -2147483643
      End
   End
End
Attribute VB_Name = "frmReportDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()

    If ccDate.Visible Then
        frmMDI.ReportDate = ccDate.Text
    Else
        frmMDI.ReportDate = ""
    End If
    
    If ccToDate.Visible Then
        frmMDI.ReportDate2 = ccToDate.Text
    Else
        frmMDI.ReportDate2 = ""
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    centerForm Me
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


