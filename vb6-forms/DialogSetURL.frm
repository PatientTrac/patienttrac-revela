VERSION 5.00
Begin VB.Form frmDialogSetURL 
   Caption         =   "Set Dashboard URL"
   ClientHeight    =   1845
   ClientLeft      =   5310
   ClientTop       =   2145
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   5880
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4260
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   3
      Top             =   570
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtPath 
      Height          =   330
      Left            =   405
      TabIndex        =   1
      Text            =   "http://"
      Top             =   585
      Width           =   4560
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2100
      TabIndex        =   0
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Enter the Dashboard URL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   405
      TabIndex        =   2
      Top             =   315
      Width           =   2265
   End
End
Attribute VB_Name = "frmDialogSetURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CancelButton_Click()
   Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
   If Trim(txtPath.Text) = vbNullString Then
      MsgBox "URL Required", vbExclamation, App.Title
      Exit Sub
   End If

   SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "DashboardURL", txtPath, REG_SZ
   StartingAddress = txtPath
   Unload Me
End Sub
Private Sub Command1_Click(Index As Integer)
   'Dim lstrPath As String
   
   'lstrPath = ChooseFolder("Set File Path for Consent Forms")
   'If Trim(lstrPath)<> vbnullstring Then txtPath.Text = lstrPath
      
End Sub
Private Sub Form_Load()
    centerForm Me
    
    txtPath.Text = StartingAddress
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
End Sub


