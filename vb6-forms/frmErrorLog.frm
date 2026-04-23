VERSION 5.00
Begin VB.Form frmErrorLog 
   Caption         =   "Error Log"
   ClientHeight    =   5895
   ClientLeft      =   7125
   ClientTop       =   3600
   ClientWidth     =   10365
   Icon            =   "frmErrorLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   10365
   Begin VB.CommandButton cmdPurge 
      Caption         =   "Purge"
      Height          =   405
      Left            =   5130
      TabIndex        =   2
      Top             =   5430
      Width           =   885
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   4110
      TabIndex        =   1
      Top             =   5430
      Width           =   915
   End
   Begin VB.TextBox txtErrors 
      Height          =   5085
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   210
      Width           =   10005
   End
End
Attribute VB_Name = "frmErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstrLogType As String
Private fstrComponentsPath As String
Private fstrPath As String
Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdPurge_Click()
Dim lintResp As Byte

    lintResp = MsgBox("Are you sure you want to delete error log?", vbExclamation + vbYesNo)
    
    If lintResp = vbYes Then
        Kill fstrPath & "\Error_Log.txt"
        txtErrors.Text = vbNullString
        cmdPurge.Enabled = False
    End If

End Sub

Private Sub Form_Load()
Dim fso As New FileSystemObject

    centerForm Me
    fstrComponentsPath = pstrMainPath & "\Components"
    
    
    If pstrLogType = "Component" Then
        fstrPath = fstrComponentsPath
        If fso.FileExists(fstrComponentsPath & "\Error_log.txt") Then
            txtErrors.Text = File_Method(fstrComponentsPath & "\Error_Log.txt", "open")
            Me.Caption = "Local Components Error Log"
            cmdPurge.Enabled = True
        Else
            cmdPurge.Enabled = False
        End If
    Else
        fstrPath = App.Path
        If fso.FileExists(App.Path & "\Error_Log.txt") Then
            txtErrors.Text = File_Method(App.Path & "\Error_Log.txt", "open")
            cmdPurge.Enabled = True
            Me.Caption = "Local Error Log"
        Else
            cmdPurge.Enabled = False
        End If
    End If
    
End Sub


