VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMessage 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3315
   ClientLeft      =   7905
   ClientTop       =   3765
   ClientWidth     =   4785
   ControlBox      =   0   'False
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4785
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3960
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox cctMessageBox 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.Timer tmrOut 
      Enabled         =   0   'False
      Left            =   4080
      Top             =   120
   End
   Begin PT_XP_Button.PT_XP cmdCancel 
      Height          =   465
      Left            =   1778
      TabIndex        =   2
      Top             =   2400
      Width           =   1125
      _ExtentX        =   1984
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
      Picture         =   "frmMessage.frx":08CA
      AppearanceThemes=   3
      TransparentColor=   8454143
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMessage.frx":12DC
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function Get_PageData(vstrData As String) As String
Dim strSourceCode As String
    On Error Resume Next
    strSourceCode = Inet1.OpenURL(vstrData)
    If strSourceCode <> vbNullString Then
        If TextFound("<--BEGIN RX>", strSourceCode, Len("BEGIN RX"), True) Then
            frmMDI.cmdElectronicNotice.BackColor = &HD6D8FE
            frmMDI.cmdElectronicNotice.BorderColor = &H404080
            frmMDI.cmdElectronicNotice.ForeColor = &H0&
        Else
            frmMDI.cmdElectronicNotice.BackColor = &HFFFFFF
            frmMDI.cmdElectronicNotice.BorderColor = &H743C00
            frmMDI.cmdElectronicNotice.ForeColor = &H800000
        End If
        strSourceCode = Trim(strSourceCode)
        Get_PageData = strSourceCode
        cctMessageBox.Text = Trim(strSourceCode)
    End If
    Exit Function
End Function
Public Sub ShowMessage _
 ( _
 ByVal vstrText As String, _
 ByVal vbooCritical As Boolean, _
 ByVal vlngLeft As Long, _
 ByVal vlngTop As Long _
 )
    On Error GoTo Error_ShowMessage
    lblMessage.Caption = vstrText
'    Move vlngLeft, vlngTop, lblMessage.Width + 90, lblMessage.Height + 60
    tmrOut.Interval = 3000
    If vbooCritical = True Then
        BackColor = &H40C0&
        lblMessage.ForeColor = &HFFFF&
    End If
    tmrOut.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_ShowMessage:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, lblMessage.Caption, "ShowMessage", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Form_Click()
    tmrOut_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  Unload Me
  Exit Sub
End Sub


Private Sub lblMessage_Click()

    tmrOut_Timer
    
End Sub

Private Sub tmrOut_Timer()

    tmrOut.Enabled = False
    Unload Me
    
End Sub
