VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PatientTrac User Security"
   ClientHeight    =   4545
   ClientLeft      =   12870
   ClientTop       =   4005
   ClientWidth     =   8190
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   8190
   Begin VB.Frame fraPass 
      BackColor       =   &H00E0E0E0&
      Height          =   2355
      Left            =   2010
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   5625
      Begin VB.TextBox cctVeryPass 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2010
         PasswordChar    =   "*"
         TabIndex        =   22
         Top             =   1290
         Width           =   3405
      End
      Begin VB.TextBox cctNewPass 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2010
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   765
         Width           =   3405
      End
      Begin VB.TextBox cctCurPass 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2010
         TabIndex        =   20
         Top             =   240
         Width           =   3405
      End
      Begin PT_XP_Button.PT_XP cmdClosePass 
         Height          =   465
         Left            =   3195
         TabIndex        =   26
         Top             =   1770
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "Login.frx":038A
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Cancel"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Login.frx":0D9C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdChangePass 
         Height          =   465
         Left            =   1830
         TabIndex        =   27
         Top             =   1770
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "Login.frx":10EE
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Change Password"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Login.frx":1488
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblVeryPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Password"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   150
         TabIndex        =   25
         Top             =   1290
         Width           =   1575
      End
      Begin VB.Label lblNewPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   150
         TabIndex        =   24
         Top             =   870
         Width           =   1350
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Password"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   150
         TabIndex        =   23
         Top             =   270
         Width           =   1740
      End
   End
   Begin VB.Frame fraConnect 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4485
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   8145
      Begin VB.ComboBox cboFacility 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3360
         Locked          =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2340
         Width           =   3405
      End
      Begin VB.ComboBox cboSpecialty 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2940
         Width           =   3405
      End
      Begin VB.PictureBox PicPT 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   210
         Picture         =   "Login.frx":17DA
         ScaleHeight     =   2205
         ScaleWidth      =   1455
         TabIndex        =   11
         Top             =   210
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3330
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1470
         Width           =   3405
      End
      Begin VB.ComboBox cboUser 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3330
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   3405
      End
      Begin PT_XP_Button.PT_XP cmdCancel 
         Height          =   405
         Left            =   4140
         TabIndex        =   6
         Top             =   4050
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "Login.frx":BFC8
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Cancel"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Login.frx":C9DA
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdConnect 
         Height          =   405
         Left            =   2775
         TabIndex        =   3
         Top             =   4050
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "Login.frx":CD2C
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Connect"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Login.frx":D0C6
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdStatus 
         Height          =   315
         Left            =   6840
         TabIndex        =   15
         ToolTipText     =   "Change Bill to Facility"
         Top             =   2340
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Login.frx":D418
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "Login.frx":D9B2
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPass 
         Height          =   315
         Left            =   6870
         TabIndex        =   18
         ToolTipText     =   "Change Password"
         Top             =   1500
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Login.frx":DD04
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "Login.frx":E29E
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblBillFac 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EDI Facility"
         ForeColor       =   &H00800000&
         Height          =   600
         Left            =   2160
         TabIndex        =   17
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblSpecialty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Specialty"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2370
         TabIndex        =   14
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label lblExpDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   5820
         TabIndex        =   13
         Top             =   3450
         Width           =   1695
      End
      Begin VB.Label lblPractice 
         BackStyle       =   0  'Transparent
         Caption         =   "Specialty"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   4650
         Width           =   1005
      End
      Begin VB.Label lblNotice 
         BackStyle       =   0  'Transparent
         Caption         =   $"Login.frx":E5F0
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   585
         Left            =   180
         TabIndex        =   10
         Top             =   3450
         Width           =   7845
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PatientTrac Electronic Medical Records"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2100
         TabIndex        =   9
         Top             =   300
         Width           =   4545
      End
      Begin VB.Label lblUserID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2520
         TabIndex        =   8
         Top             =   1155
         Width           =   690
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2340
         TabIndex        =   7
         Top             =   1515
         Width           =   900
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Password is Case-Sensitive)"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2880
         TabIndex        =   4
         Top             =   1920
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Public mlngReturn As Long
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal Hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long
Private Function Update_Password(vlngProvider As Long)
Dim cnConnect As New ADODB.Connection
Dim rsConnect As New ADODB.Recordset
Dim strSQL As String
Dim lngProvID As Long

    On Error GoTo Error_Update_Password
    Screen.MousePointer = vbHourglass
    lngProvID = vlngProvider
    If cnConnect.State = adStateClosed Then cnConnect.open psConnect(0)
    strSQL = "UPDATE PROVIDERS SET USER_PASSWORD = '" & Trim(cctNewPass) & "'  WHERE PROVIDER_ID = " & lngProvID
    Set rsConnect = cnConnect.Execute(strSQL)
    
    If cnConnect.State = adStateOpen Then
        cnConnect.Close
        Set cnConnect = Nothing
        Set rsConnect = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_Password:
   Screen.MousePointer = vbNormal
    Exit Function
    Resume Next
End Function

Private Function Write_LogOn()
Dim cnLogOn As New ADODB.Connection
Dim rsLogOn As New ADODB.Recordset
Dim strSQL As String
Dim strPC As String
Dim strTime As String
Dim strLogon As String

    If cnLogOn.State = adStateClosed Then cnLogOn.open psConnect(0)
    On Error GoTo Error_Write_LogOn
    strPC = Get_ComputerName
    strTime = Format(Now, "hh:mm AM/PM")
    strLogon = Format(Now, "mm/dd/yyyy")
    strSQL = "Write_User_Logon '" & cboUser.Text & "'," & mlngLogonUser & ",'" & strPC & "','PatientTrac EMR', '" _
        & Format(strTime, "hh:mm AM/PM") & "', '" & strLogon & "'"
    Set rsLogOn = cnLogOn.Execute(strSQL)
    If cnLogOn.State = adStateOpen Then
        cnLogOn.Close
        Set cnLogOn = Nothing
        Set rsLogOn = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_LogOn:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_LogOn", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Load_Specialties()
Dim strSQL As String
Dim ExpDate As String
Dim sDate As String
Dim lngDateDiff As Long
Dim strTermDate As String


    sDate = Format(Now, "mm/dd/yyyy")
    strTermDate = Format("10/15/2015", "mm/dd/yyyy")
    lngDateDiff = DateDiff("d", sDate, strTermDate)
    If lngDateDiff <= 0 Then
        MsgBox "PatientTrac has terminated due to the failure to maintain the License Agreement and is now uninstalled. Your previous sign on have acknowledged that" _
            & " " & "Notice: Logging into this System constitutes an acceptance of the EULA and your adherance to the Copyright Protection for this software." _
            & " " & "Please contact PatientTrac to reinstate your License Agreement! Please be advised PatientTrac Staff will provide you with paper" _
            & " " & "or electronic copies of the Patient Records for a nominal fee.", vbInformation
        pbCancel = True
        License_Expired
        Exit Function
    End If
    If Not pbCancel Then
        Open_DB psConnect(1)
        strSQL = "SELECT EMR_PRACTICE, EMR_SPECIALTY,EMR_GRP_ID FROM EMR_GROUP  "
        Set rs = cn.Execute(strSQL)
        With rs
            Do Until .EOF
                If (!EMR_SPECIALTY) <> vbNullString Then
                    cboSpecialty.AddItem (!EMR_SPECIALTY)
                    cboSpecialty.ItemData(cboSpecialty.NewIndex) = (!EMR_GRP_ID)
                End If
            .MoveNext
            Loop
        End With
        Close_DB psConnect(1)
    End If
    Exit Function
End Function


Private Sub cboFacility_Click()
Dim strSQL As String
Dim rsPatient As New ADODB.Recordset
Dim cnPatient As New ADODB.Connection
Dim Idx%

    On Error GoTo Error_Change_Bill_Facility
    Screen.MousePointer = vbHourglass
    For Idx = 0 To cboFacility.listcount - 1
            If cboFacility.List(Idx) = cboFacility.Text Then
                clsFacility.BillFacility = cboFacility.ItemData(Idx)
                Exit For
            End If
    Next Idx
    'Need to update NPI, Tax ID and other EDI_STATIC Fields from Facility Table
    If cnPatient.State = adStateClosed Then cnPatient.open psConnect(0)
    strSQL = "UPDATE EDI_STATIC SET VALUE_DESCRIPTION  = '" & CStr(clsFacility.BillFacility) & "' WHERE FIELD_DESCRIPTION = 'Bill_To_Provider'"
    Set rsPatient = cnPatient.Execute(strSQL)
    If cnPatient.State = adStateOpen Then
        cnPatient.Close
        Set cnPatient = Nothing
        Set rsPatient = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Change_Bill_Facility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Change_Bill_Facility", App.EXEName & "." & "modBilling"
    Exit Sub
    Resume Next

End Sub

Private Sub cboSpecialty_Click()
Dim i%
Dim strSpecialty As String

    strSpecialty = Trim(cboSpecialty.Text)
    SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "PracticeType", strSpecialty, REG_SZ
    SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "Specialty", strSpecialty, REG_SZ
    SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\FirstNoteBook\PatientProfiler\General", "PracticeType", strSpecialty, REG_SZ
    SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\FirstNoteBook\PatientTracSCHED\General", "PracticeType", strSpecialty, REG_SZ
    mstrPractice = Trim(strSpecialty)
    cmdConnect.Enabled = True
    Exit Sub
End Sub

Private Sub cboSpecialty_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    
    If cboSpecialty.Text <> vbNullString Then strList = cboSpecialty.Text
    cboSpecialty.ListIndex = SendMessage(cboSpecialty.Hwnd, LB_FINDSTRING, -1, _
    ByVal strList)
    Exit Sub
End Sub


Private Sub cboUser_Click()
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
Dim iAnswer As Integer

    iAnswer = MsgBox("Do you want to close PatientTrac?", vbYesNo)
    If iAnswer = vbYes Then
        pbCancel = True
        Unload Me
    Else
        cboUser.ListIndex = -1
        txtPassword.Text = vbNullString
        Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cmdChangePass_Click()
Dim cnConnect As New ADODB.Connection
Dim rsConnect As New ADODB.Recordset
Dim strSQL As String
Dim iAnswer As Integer
Dim strUser As String
Dim strSuffix As String
Dim strTitle As String
Dim strName As String
Dim bErr As Boolean
Dim lngProvID As Long
Dim Idx%

    On Error GoTo Error_Handler
    Screen.MousePointer = vbHourglass
    If pbCancel Then Exit Sub
    If cboUser.ListIndex = -1 Then
       MsgBox "Please enter a valid User Id.", vbExclamation, "PatientTrac EMR Logon"
       cboUser.SetFocus
       Exit Sub
    End If
'    If Trim(txtPassword.Text) = vbNullString And pbCancel = False Then
'       MsgBox "Please enter password.", vbExclamation, "PatientTrac EMR Logon"
'       txtPassword.SetFocus
'        Screen.MousePointer = vbDefault
'       Exit Sub
'    End If
    For Idx = 0 To cboUser.listcount - 1
        If cboUser.List(Idx) = cboUser.Text Then
            lngProvID = cboUser.ItemData(Idx)
            Exit For
        End If
    Next Idx
    If lngProvID > 0 Then
        If cnConnect.State = adStateClosed Then cnConnect.open psConnect(0)
        strSQL = "SELECT USER_PASSWORD FROM PROVIDERS WHERE PROVIDER_ID = " & lngProvID
        Set rsConnect = cnConnect.Execute(strSQL)
        With rsConnect
            If Not .EOF Then
                If Not IsNull((!USER_PASSWORD)) Then
                    If Trim((!USER_PASSWORD)) <> Trim(cctCurPass) Then
                        MsgBox "The password is incorrect, please re-enter correct password!", vbInformation, vbOK
                        cctCurPass = vbNullString
                    Else
                        If Trim(cctNewPass) = Trim(cctVeryPass) Then
                            Update_Password lngProvID
                            MsgBox "Password successfully update, please login with new password.", vbInformation, vbOK
                            fraPass.Visible = False
                        Else
                            MsgBox "The passwords do not match, please re-enter correct password!", vbInformation, vbOK
                            cctVeryPass = vbNullString
                            cctNewPass = vbNullString
                        End If
                    End If
                End If
            End If
        End With
    End If
    If cnConnect.State = adStateOpen Then
        cnConnect.Close
        Set cnConnect = Nothing
        Set rsConnect = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Handler:
   Screen.MousePointer = vbNormal
    MsgBox "Please contact PatientTrac Technical Support, Error Login" & " " & Err.Description, vbOKOnly + vbExclamation + vbApplicationModal + vbDefaultButton1, "PatientTrac EMR Logon"
    Exit Sub
    Resume Next
End Sub

Private Sub cmdClosePass_Click()
    fraPass.Visible = False
    Exit Sub
End Sub

Private Sub cmdConnect_Click()
Dim cnConnect As New ADODB.Connection
Dim rsConnect As New ADODB.Recordset
Dim iAnswer As Integer
Dim strUser As String
Dim strSuffix As String
Dim strTitle As String
Dim strName As String
Dim bErr As Boolean
Dim lngProvID As Long
Dim Idx%

    On Error GoTo Error_Handler
    Screen.MousePointer = vbHourglass
    If pbCancel Then Exit Sub
    If cboUser.ListIndex = -1 Then
       MsgBox "Please enter a valid User Id.", vbExclamation, "PatientTrac EMR Logon"
       cboUser.SetFocus
       Exit Sub
    End If
    If Trim(txtPassword.Text) = vbNullString And pbCancel = False Then
       MsgBox "Please enter password.", vbExclamation, "PatientTrac EMR Logon"
       txtPassword.SetFocus
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    For Idx = 0 To cboUser.listcount - 1
        If cboUser.List(Idx) = cboUser.Text Then
            mlngLogonUser = cboUser.ItemData(Idx)
            Exit For
        End If
    Next Idx
    Get_Providers mlngLogonUser
    piProviderID = clsProvider.ProviderID
    Write_LogOn
    psRegGeneral = "SOFTWARE\Firstnotebook\PatientTrac\General"
    If Trim(clsProvider.Password) <> Trim(txtPassword.Text) Then
        MsgBox "Incorrect Password, Please Re-Enter a Valid Password.", vbCritical, "PatientTrac EMR Logon"
        txtPassword.Text = vbNullString
        txtPassword.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
     If mlngLogonUser > 1 Then
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "CurrentUser", mlngLogonUser, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, "PatientTracSCHED", "CurrentUser", mlngLogonUser, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultProvider", mlngLogonUser, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, "PatientTracSCHED", "DefaultProvider", mlngLogonUser, REG_SZ
        SetKeyValue HKEY_LOCAL_MACHINE, "PatientProfiler", "DefaultProvider", mlngLogonUser, REG_SZ
    End If
    Screen.MousePointer = vbDefault
    Get_Practice
    frmMDI.StatusBar.Panels(6).Text = "Provider: " & clsProvider.Name
    frmMDI.StatusBar.Panels(5).Text = "Visit Date: " & Format(Now, "MMMM DD, YYYY")
    pbCancel = False
    Unload Me
    Exit Sub
Error_Handler:
   Screen.MousePointer = vbNormal
    MsgBox "Please contact PatientTrac Technical Support, Error Login" & " " & Err.Description, vbOKOnly + vbExclamation + vbApplicationModal + vbDefaultButton1, "PatientTrac EMR Logon"
    Exit Sub
    Resume Next
End Sub

Private Sub cmdConnect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
        cmdConnect_Click
    End If
End Sub


Private Sub cmdPass_Click()
    fraPass.Visible = True
    Exit Sub
End Sub

Private Sub cmdStatus_Click()
    cboFacility.Locked = False
    Exit Sub
End Sub

Private Sub Form_Activate()
Dim sConnect As String
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Idx As Integer
Dim strSQL As String

    On Error GoTo Error_Handler
    If mFirst Then mFirst = False
    cboUser.Enabled = True
    If cn.State = adStateClosed Then cn.open psConnect(0)
   strSQL = "SELECT * FROM PROVIDERS WHERE ACTIVE = 1"
    Set rs = cn.Execute(strSQL)
    With rs
       Do Until .EOF
          If Trim((!User_ID) & "") <> "" Then
             cboUser.AddItem (!User_ID) & ""
             cboUser.ItemData(cboUser.NewIndex) = (!Provider_ID)
          End If
          .MoveNext
       Loop
    End With
    Close_DB psConnect(1)
    Load_Specialties
    Get_Practice
    If mstrPractice <> vbNullString Then
        For Idx = 0 To cboSpecialty.listcount - 1
            If cboSpecialty.List(Idx) = mstrPractice Then
                cboSpecialty.Text = cboSpecialty.List(Idx)
                cboSpecialty.Enabled = False
                Exit For
            End If
        Next Idx
    End If
    Load_Facility
    If mlngReturn > 0 Then
        For Idx = 0 To cboUser.listcount - 1
            If cboUser.ItemData(Idx) = mlngReturn Then
                cboUser.Text = cboUser.List(Idx)
                cboUser.Enabled = False
                Exit For
            End If
        Next Idx
    End If
    cboUser.SetFocus
    centerForm Me
    Exit Sub
Error_Handler:
    Exit Sub
    Resume Next
End Sub
Private Function Load_Facility()
Dim strSQL As String
Dim cnFac As New ADODB.Connection
Dim rsFac As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngFacID As Long
Dim Idx%
Dim strName As String
Dim booSkip As Boolean
 

    On Error GoTo Error_Load_Facility
    Screen.MousePointer = vbHourglass
    cboFacility.Clear
    If cnFac.State = adStateClosed Then cnFac.open psConnect(0)
    strSQL = "SELECT DISTINCT FACILITY_ID, FACILITY_NAME FROM FACILITIES " _
        & "WHERE ACTIVE = 1"
    Set rsFac = cnFac.Execute(strSQL)
    With rsFac
        Do Until .EOF
            lngFacID = 0
            strName = vbNullString
            booSkip = False
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "FACILITY_NAME"
                                strName = Fld.Value
                            Case "FACILITY_ID"
                                lngFacID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            strName = Fix_Name(strName)
            strName = Fix_Facility_Name(strName)
            If Trim(strName) <> vbNullString Then
                cboFacility.AddItem strName
                cboFacility.ItemData(cboFacility.NewIndex) = lngFacID
            End If
        Loop
    End With
    Get_Bill_Facility
    If clsFacility.BillFacility > 0 Then
        For Idx = 0 To cboFacility.listcount - 1
            If cboFacility.ItemData(Idx) = clsFacility.BillFacility Then
                cboFacility.Text = cboFacility.List(Idx)
                Exit For
            End If
        Next Idx
    End If
    If cnFac.State = adStateOpen Then
        cnFac.Close
        Set cnFac = Nothing
        Set rsFac = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Facility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Facility", App.EXEName & "." & "modBilling"
    Exit Function
    Resume Next
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
        cmdConnect_Click
    End If
End Sub

Private Sub Form_Load()
    mFirst = True
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    booAppTime = True
    Unload Me
    Exit Sub
End Sub
Private Sub Timer1_Timer()

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
        cmdConnect_Click
    End If
    If GetKeyState(vbKeyCapital) = 1 Then
        MsgBox "Cap Lock Key is ON!", vbInformation, "PatientTrac Logon"
    End If
    Exit Sub
    
End Sub

Private Sub txtPassword_LostFocus()

    cmdConnect.SetFocus

End Sub

