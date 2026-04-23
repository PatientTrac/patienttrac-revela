VERSION 5.00
Begin VB.Form frmSaveAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EMR PatientTrac"
   ClientHeight    =   3615
   ClientLeft      =   9555
   ClientTop       =   6750
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   60
      Picture         =   "SaveAlert.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   0
      Width           =   675
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "Select None"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Select All"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ListBox lstComponents 
      Height          =   2535
      Left            =   960
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   660
      Width           =   3465
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save Selected"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Changes to the areas listed below are about to be lost.   Select only the areas you want to save."
      Height          =   450
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   3465
   End
End
Attribute VB_Name = "frmSaveAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mbooCancelled As Boolean

Public Property Get Cancelled() As Boolean

    Cancelled = mbooCancelled
    
End Property


Private Sub cmdAll_Click()

    SetAll True
    
End Sub

Private Sub cmdCancel_Click()

    mbooCancelled = True
    Me.Hide
    
End Sub

Private Sub cmdNone_Click()

    SetAll False
    
End Sub

Private Sub cmdOK_Click()

    Me.Hide
    
End Sub

Private Sub SetAll(ByVal vbooSelected As Boolean)

    Dim lintCounter As Integer
    Dim lintSelected As Integer
    
    lintSelected = lstComponents.ListIndex
    
    For lintCounter = 0 To lstComponents.listcount - 1
    
        lstComponents.Selected(lintCounter) = vbooSelected
        
    Next
    
    lstComponents.ListIndex = lintSelected
        
End Sub


Private Sub Form_Load()

    mbooCancelled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

  Unload Me
  Exit Sub
End Sub


