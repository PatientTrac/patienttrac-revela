VERSION 5.00
Begin VB.Form frmLogFile 
   Caption         =   "Log File"
   ClientHeight    =   6120
   ClientLeft      =   2565
   ClientTop       =   2850
   ClientWidth     =   10530
   Icon            =   "frmLogFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   10530
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   4530
      TabIndex        =   1
      Top             =   5490
      Width           =   915
   End
   Begin VB.TextBox txtLog 
      Height          =   5085
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   150
      Width           =   10185
   End
End
Attribute VB_Name = "frmLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pstrLogType As String
Private fstrComponentsPath As String
Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub Form_Load()
Dim fso As New FileSystemObject

    centerForm Me
    fstrComponentsPath = pstrMainPath & "\Components"
    
    If pstrLogType = "Component" Then
        If fso.FileExists(fstrComponentsPath & "\log.txt") Then
            txtLog.Text = File_Method(fstrComponentsPath & "\log.txt", "open")
            Me.Caption = "Local Components Event Log"
        End If
    Else
        If fso.FileExists(App.Path & "\log.txt") Then
            txtLog.Text = File_Method(App.Path & "\log.txt", "open")
            Me.Caption = "Local Event Log"
        End If
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Exit Sub
End Sub


