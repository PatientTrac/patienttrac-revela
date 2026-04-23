VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmDialogSetExt 
   Caption         =   "Set External System Connection"
   ClientHeight    =   4335
   ClientLeft      =   6990
   ClientTop       =   2565
   ClientWidth     =   8745
   ControlBox      =   0   'False
   Icon            =   "DialogSetExt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   8745
   Begin VB.Frame fraClose 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2985
      TabIndex        =   15
      Top             =   3660
      Width           =   2775
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   555
         Left            =   1410
         TabIndex        =   16
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "DialogSetExt.frx":08CA
         BackColorDown   =   15133676
         TransparentColor=   14215660
         Caption         =   "Close"
         DepthEvent      =   1
         PictureDisabled =   "DialogSetExt.frx":11A4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   555
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "DialogSetExt.frx":1DF6
         BackColorDown   =   15133676
         TransparentColor=   14215660
         Caption         =   "Save"
         DepthEvent      =   1
         PictureDisabled =   "DialogSetExt.frx":27F0
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lytec Values"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3495
      Left            =   4320
      TabIndex        =   11
      Top             =   30
      Width           =   4215
      Begin VB.TextBox txtDefLab 
         Height          =   330
         Left            =   1845
         MaxLength       =   9
         TabIndex        =   12
         ToolTipText     =   "Set to same value in Lytec under Settings/Preferences"
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label lblDefLab 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Default Laboratory"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   405
         Width           =   1485
      End
   End
   Begin VB.Frame fraConnection 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Connection Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3495
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4215
      Begin VB.ComboBox cboSystem 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtDatabase 
         Height          =   330
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   2475
      End
      Begin VB.TextBox txtDSN 
         Height          =   330
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   2475
      End
      Begin VB.TextBox txtUserID 
         Height          =   330
         Left            =   1560
         TabIndex        =   2
         Top             =   1800
         Width           =   2475
      End
      Begin VB.TextBox txtPassword 
         Height          =   330
         Left            =   1560
         TabIndex        =   1
         Top             =   2280
         Width           =   2475
      End
      Begin PT_XP_Button.PT_XP cmdTest 
         Height          =   495
         Left            =   1200
         TabIndex        =   14
         Top             =   2850
         Width           =   1845
         _ExtentX        =   3254
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
         Picture         =   "DialogSetExt.frx":3172
         AppearanceMode  =   1
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Test Connection"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "DialogSetExt.frx":350C
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblSystem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "External System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label lblDatabase 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   885
         Width           =   825
      End
      Begin VB.Label lblDSN 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "DSN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1035
         TabIndex        =   8
         Top             =   1365
         Width           =   405
      End
      Begin VB.Label lblUID 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   780
         TabIndex        =   7
         Top             =   1845
         Width           =   660
      End
      Begin VB.Label lblPWD 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   615
         TabIndex        =   6
         Top             =   2325
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmDialogSetExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
   SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "BillInt_System", cboSystem.Text, REG_SZ
   SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "BillInt_DB", txtDatabase, REG_SZ
   SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "BillInt_DSN", txtDSN, REG_SZ
   SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "BillInt_UID", txtUserID, REG_SZ
   SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "BillInt_PWD", txtPassword, REG_SZ
   SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "LytecDefLab", txtDefLab, REG_SZ
   
   psExtSystem = cboSystem.Text
   psExtDB = txtDatabase
   psExtDSN = txtDSN
   psExtUID = txtUserID
   psExtPWD = txtPassword

   Unload Me
End Sub

Private Sub cmdTest_Click()
   Dim sConnect As String
   Dim cn As New ADODB.Connection
   
   On Error GoTo Test_Err
   
   sConnect = "ODBC;DATABASE=" & txtDatabase & ";UID=" & txtUserID & ";PWD=" & txtPassword & ";DSN=" & txtDSN & ";"
   
   cn.open sConnect
    
   cn.Close
   
   MsgBox "Connection successful.", vbInformation
   
   Exit Sub
   
Test_Err:
   MsgBox "Connection Failed. " & vbNewLine & Err.Number & ":" & Err.Description, vbInformation
End Sub

Private Sub Form_Load()
    Dim Idx As Integer
    
    centerForm Me
    
    cboSystem.Clear
    cboSystem.AddItem "Lytec"
    cboSystem.AddItem "Prime Clinical"
    cboSystem.AddItem "Medical Manager"
    
    For Idx = 0 To cboSystem.ListCount - 1
        If cboSystem.List(Idx) = psExtSystem Then
           cboSystem.ListIndex = Idx
           Exit For
        End If
    Next Idx
    
    txtDatabase = psExtDB
    txtDSN = psExtDSN
    txtUserID = psExtUID
    txtPassword = psExtPWD
    
    txtDefLab = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "LytecDefLab")

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
End Sub


