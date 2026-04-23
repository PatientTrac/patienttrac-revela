VERSION 5.00
Begin VB.Form frmSetServers 
   Caption         =   "Set Database Connections"
   ClientHeight    =   4785
   ClientLeft      =   7860
   ClientTop       =   3735
   ClientWidth     =   7275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   7275
   Begin VB.Frame fraClinicalRepos 
      Caption         =   "Clinical Repository"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   0
      TabIndex        =   14
      Top             =   270
      Width           =   3585
      Begin VB.TextBox txtDriver 
         Height          =   330
         Index           =   0
         Left            =   975
         TabIndex        =   20
         Text            =   "SQL Server"
         Top             =   300
         Width           =   2490
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   975
         TabIndex        =   19
         Top             =   2745
         Width           =   1770
      End
      Begin VB.TextBox txtDatabase 
         Height          =   330
         Index           =   0
         Left            =   990
         TabIndex        =   18
         Text            =   "CLINICAL_REPOSITORY"
         Top             =   2160
         Width           =   2475
      End
      Begin VB.TextBox txtPassword 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   990
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   1710
         Width           =   2475
      End
      Begin VB.ComboBox cboServer 
         Height          =   315
         Index           =   0
         Left            =   990
         TabIndex        =   16
         Top             =   780
         Width           =   2475
      End
      Begin VB.ComboBox cboUserID 
         Height          =   315
         Index           =   0
         Left            =   990
         TabIndex        =   15
         Top             =   1230
         Width           =   2475
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   330
         TabIndex        =   25
         Top             =   825
         Width           =   570
      End
      Begin VB.Label lblDriver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
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
         Index           =   0
         Left            =   390
         TabIndex        =   24
         Top             =   315
         Width           =   525
      End
      Begin VB.Label lblDatabase 
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
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   23
         Top             =   2205
         Width           =   870
      End
      Begin VB.Label lblUserID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   22
         Top             =   1305
         Width           =   660
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   21
         Top             =   1755
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Terms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   3630
      TabIndex        =   2
      Top             =   270
      Width           =   3585
      Begin VB.TextBox txtPassword 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1710
         Width           =   2475
      End
      Begin VB.TextBox txtDatabase 
         Height          =   330
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Text            =   "TERMS"
         Top             =   2160
         Width           =   2475
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   975
         TabIndex        =   6
         Top             =   2745
         Width           =   1770
      End
      Begin VB.TextBox txtDriver 
         Height          =   330
         Index           =   1
         Left            =   945
         TabIndex        =   5
         Text            =   "SQL Server"
         Top             =   300
         Width           =   2490
      End
      Begin VB.ComboBox cboServer 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   780
         Width           =   2505
      End
      Begin VB.ComboBox cboUserID 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   1260
         Width           =   2505
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   1755
         Width           =   825
      End
      Begin VB.Label lblUserID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   12
         Top             =   1305
         Width           =   660
      End
      Begin VB.Label lblDatabase 
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
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   2205
         Width           =   870
      End
      Begin VB.Label lblDriver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
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
         Index           =   1
         Left            =   390
         TabIndex        =   10
         Top             =   315
         Width           =   525
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   330
         TabIndex        =   9
         Top             =   825
         Width           =   570
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3810
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Configure Connection to Servers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   7125
   End
End
Attribute VB_Name = "frmSetServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CheckValidConnection As Boolean 'To Check Validity of Connections to DBs
Private fsKeyServer0() As String
Private fsKeyUser0() As String
Private fbUserKeyLoaded0 As Boolean
Private fbServerKeyLoaded0 As Boolean

Private fsKeyServer1() As String
Private fsKeyUser1() As String
Private fbUserKeyLoaded1 As Boolean
Private fbServerKeyLoaded1 As Boolean
Private Sub CancelButton_Click()
Screen.MousePointer = vbHourglass
   Unload Me
   
   If pbSetDB Then
        Unload frmMDI
   End If
 Screen.MousePointer = vbDefault
End Sub
Private Sub cmdTest_Click(Index As Integer)
   Dim lstrConnect As String
'Code Written by LP on 24-Dec-2002
   If checkServerInfo(Index) = True Then
       Screen.MousePointer = vbHourglass
       lstrConnect = UserLogin(txtDriver(Index), cboServer(Index), txtDatabase(Index), cboUserID(Index), txtPassword(Index))
       Screen.MousePointer = vbDefault
        
       If lstrConnect <> vbNullString Then
           MsgBox "Connection to database was successful.", vbInformation
    '      MsgBox_Lang 34, vbOKOnly, vbInformation, App.Title, ""
          psConnect(0) = lstrConnect
       Else
'           MsgBox "Connection to database was failed!", vbCritical
    '      MsgBox_Lang 35, vbOKOnly, vbCritical, App.Title, ""
       End If
   End If
End Sub

'Code Written by LP on 24-Dec-2002
Private Function checkServerInfo(Index) As Boolean
    If Trim(txtDriver(Index)) = "" Then
        MsgBox "Specify The Driver", vbInformation, App.Title
        txtDriver(Index).SelStart = 0
        txtDriver(Index).SelLength = Len(txtDriver(Index))
        txtDriver(Index).SetFocus
        checkServerInfo = False
    ElseIf Trim(cboServer(Index)) = "" Then
        MsgBox "Specify The Server", vbInformation, App.Title
        cboServer(Index).SelStart = 0
        cboServer(Index).SelLength = Len(cboServer(Index))
        cboServer(Index).SetFocus
        checkServerInfo = False
    ElseIf Trim(txtDatabase(Index)) = "" Then
        MsgBox "Specify The Database", vbInformation, App.Title
        txtDatabase(Index).SelStart = 0
        txtDatabase(Index).SelLength = Len(txtDatabase(Index))
        txtDatabase(Index).SetFocus
        checkServerInfo = False
'    ElseIf Trim(cboUserID(Index)) = "" Then
'        MsgBox "Specify The UserID", vbInformation, App.Title
'        cboUserID(Index).SelStart = 0
'        cboUserID(Index).SelLength = Len(cboUserID(Index))
'        cboUserID(Index).SetFocus
'        checkServerInfo = False
'    ElseIf Trim(txtPassword(Index)) = "" Then
'        MsgBox "Specify The Password", vbInformation, App.Title
'        txtPassword(Index).SelStart = 0
'        txtPassword(Index).SelLength = Len(txtPassword(Index))
'        txtPassword(Index).SetFocus
'        checkServerInfo = False
    Else
        checkServerInfo = True
    End If
End Function

Private Sub GenerateDefaultFile()
Dim sText As String
Dim iFileNum As Integer

 sText = sText & "[ODBC]" & vbNewLine _
    & "Driver=SQL Server" & vbNewLine _
    & "ServerA=" & vbNewLine _
    & "UserA=" & vbNewLine _
    & "ServerB=" & vbNewLine _
    & "UserB=" & vbNewLine _
    & "[SERVERS_0]" & vbNewLine _
    & "Server0=" & vbNewLine & vbNewLine _
    & "[USERS_0]" & vbNewLine _
    & "User0=" & vbNewLine & vbNewLine _
    & "[SERVERS_1]" & vbNewLine _
    & "Server0=" & vbNewLine & vbNewLine _
    & "[USERS_1]" & vbNewLine _
    & "User0=" & vbNewLine & vbNewLine
        
    iFileNum = FreeFile

    Open psIniFile For Output As iFileNum
        Print #iFileNum, sText
    Close iFileNum

End Sub


Private Sub Form_Load()
   centerForm Me
   
'----------------------------------------------------

    Load_ComboBoxes

'-------------------------------------------------------
    
   Get_DB_RegisterSettings True
    
   cboServer(0) = psServer(0)
   txtDriver(0) = psDriver(0)
    
   cboServer(1) = psServer(1)
   txtDriver(1) = psDriver(1)
   
   cboUserID(0) = psUID(0)
   cboUserID(1) = psUID(1)
    
   txtPassword(0) = psPWD(0)
   txtPassword(1) = psPWD(1)
    
   txtDatabase(0) = psDBname(0)
   txtDatabase(1) = psDBname(1)
   
End Sub
Private Sub Form_Paint()
   MonochromeBurst Me, 200, 200, 20, 3 'Gold
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim bServerExists0 As Boolean
Dim bUserExists0 As Boolean
Dim bServerExists1 As Boolean
Dim bUserExists1 As Boolean

    IniFile_Write "ODBC", "UserA", cboUserID(0), psIniFile
    IniFile_Write "ODBC", "UserB", cboUserID(1), psIniFile

    IniFile_Write "ODBC", "ServerA", cboServer(0), psIniFile
    IniFile_Write "ODBC", "ServerB", cboServer(1), psIniFile
    

        
    'Add to server list........................
    If fbServerKeyLoaded0 Then
        For i = LBound(fsKeyServer0) To UBound(fsKeyServer0)
            If cboServer(0) = fsKeyServer0(i) Then
                bServerExists0 = True
                Exit For
            End If
        Next i
        
        
        If Not bServerExists0 Then
            If fsKeyServer0(0) = "" Then
                IniFile_Write "SERVERS_0", "Server0", cboServer(0), psIniFile
            Else
                IniFile_Write "SERVERS_0", "Server" & UBound(fsKeyServer0) + 1, cboServer(0), psIniFile
            End If
        End If
    Else
        IniFile_Write "SERVERS_0", "Server0", cboServer(0), psIniFile
    End If
    
    'Add to users list........................
    If fbUserKeyLoaded0 Then
        For i = LBound(fsKeyUser0) To UBound(fsKeyUser0)
            If cboUserID(0) = fsKeyUser0(i) Then
                bUserExists0 = True
                Exit For
            End If
        Next i
        
        
        If Not bUserExists0 Then
            If fsKeyUser0(0) = "" Then
                IniFile_Write "USERS_0", "User0", cboUserID(0), psIniFile
            Else
                IniFile_Write "USERS_0", "User" & UBound(fsKeyUser0) + 1, cboUserID(0), psIniFile
            End If
        End If
    Else
        IniFile_Write "USERS_0", "User0", cboUserID(0), psIniFile
    End If
    
    
'---------------------------------------------------------------------------------

    'Add to server list........................
    If fbServerKeyLoaded1 Then
        For i = LBound(fsKeyServer1) To UBound(fsKeyServer1)
            If cboServer(1) = fsKeyServer1(i) Then
                bServerExists1 = True
                Exit For
            End If
        Next i
        
        
        If Not bServerExists1 Then
            If fsKeyServer1(0) = "" Then
                IniFile_Write "SERVERS_1", "Server0", cboServer(1), psIniFile
            Else
                IniFile_Write "SERVERS_1", "Server" & UBound(fsKeyServer1) + 1, cboServer(1), psIniFile
            End If
        End If
    Else
        IniFile_Write "SERVERS_1", "Server0", cboServer(1), psIniFile
    End If
    
    'Add to users list........................
    If fbUserKeyLoaded1 Then
        For i = LBound(fsKeyUser1) To UBound(fsKeyUser1)
            If cboUserID(1) = fsKeyUser1(i) Then
                bUserExists1 = True
                Exit For
            End If
        Next i
        
        
        If Not bUserExists1 Then
            If fsKeyUser1(0) = "" Then
                IniFile_Write "USERS_1", "User0", cboUserID(1), psIniFile
            Else
                IniFile_Write "USERS_1", "User" & UBound(fsKeyUser1) + 1, cboUserID(1), psIniFile
            End If
        End If
    Else
        IniFile_Write "USERS_1", "User0", cboUserID(1), psIniFile
    End If
    Unload Me
    Exit Sub
End Sub

Private Sub OKButton_Click()
'Code Written by LP on 24-Dec-2002
   Dim lstrConnect As String
   
   Screen.MousePointer = vbHourglass
   
If checkServerInfo(0) = True And checkServerInfo(1) = True Then
    CheckValidConnection = True
    lstrConnect = UserLogin(txtDriver(0), cboServer(0), txtDatabase(0), cboUserID(0), txtPassword(0))
    If CheckValidConnection = False Then Exit Sub
    CheckValidConnection = True
    lstrConnect = UserLogin(txtDriver(1), cboServer(1), txtDatabase(1), cboUserID(1), txtPassword(1))
    If CheckValidConnection = False Then Exit Sub
    '   SaveSetting appname:="PatientProfiler", Section:="General", Key:="Server0", setting:=cboServer(0)
    '   SaveSetting appname:="PatientProfiler", Section:="General", Key:="Server1", setting:=cboServer(1)
    
       SetKeyValue HKEY_LOCAL_MACHINE, psRegDB, "Server0", cboServer(0), REG_SZ
       SetKeyValue HKEY_LOCAL_MACHINE, psRegDB, "Server1", cboServer(1), REG_SZ
    
    '   SaveSetting appname:="PatientProfiler", Section:="General", Key:="Driver0", setting:=txtDriver(0)
    '   SaveSetting appname:="PatientProfiler", Section:="General", Key:="Driver1", setting:=txtDriver(1)
       
       SetKeyValue HKEY_LOCAL_MACHINE, psRegDB, "Driver0", txtDriver(0), REG_SZ
       SetKeyValue HKEY_LOCAL_MACHINE, psRegDB, "Driver1", txtDriver(1), REG_SZ
    
    '   SaveSetting appname:="PatientProfiler", Section:="UsersID", Key:="User0", setting:=cboUserID(0)
    '   SaveSetting appname:="PatientProfiler", Section:="UsersID", Key:="User1", setting:=cboUserID(1)
    
       SetKeyValue HKEY_LOCAL_MACHINE, psRegDB, "User0", cboUserID(0), REG_SZ
       SetKeyValue HKEY_LOCAL_MACHINE, psRegDB, "User1", cboUserID(1), REG_SZ
       
    '   SaveSetting appname:="PatientProfiler", Section:="Passwords", Key:="Passsword0", setting:=txtPassword(0)
    '   SaveSetting appname:="PatientProfiler", Section:="Passwords", Key:="Passsword1", setting:=txtPassword(1)
       
       SetKeyValue HKEY_LOCAL_MACHINE, psRegDB, "Password0", txtPassword(0), REG_SZ
       SetKeyValue HKEY_LOCAL_MACHINE, psRegDB, "Password1", txtPassword(1), REG_SZ
       
    '   SaveSetting appname:="PatientProfiler", Section:="Databases", Key:="Database0", setting:=txtDatabase(0)
    '   SaveSetting appname:="PatientProfiler", Section:="Databases", Key:="Database1", setting:=txtDatabase(1)
       
       SetKeyValue HKEY_LOCAL_MACHINE, psRegDB, "Database0", txtDatabase(0), REG_SZ
       SetKeyValue HKEY_LOCAL_MACHINE, psRegDB, "Database1", txtDatabase(1), REG_SZ
       
       
       'Load all of the changes in the global variables..............
       Get_DB_RegisterSettings
    
       Unload Me
    End If
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub Load_ComboBoxes()
Dim bRet As Boolean
Dim i As Integer


'Setup ini files.......................................
    psIniFile = App.Path & "\" & App.EXEName & ".ini"
    If Not FileExists(psIniFile) Then GenerateDefaultFile
    
    bRet = ReadSectionList("USERS_0", psIniFile, "User", fsKeyUser0())
    
    If bRet Then
        fbUserKeyLoaded0 = True
        For i = LBound(fsKeyUser0) To UBound(fsKeyUser0)
            If fsKeyUser0(i) <> "" Then cboUserID(0).AddItem fsKeyUser0(i)
        Next i
    End If
    
    cboUserID(0) = IniFile_Read("ODBC", "UserA", psIniFile)

    bRet = ReadSectionList("SERVERS_0", psIniFile, "Server", fsKeyServer0())
    
    If bRet Then
        fbServerKeyLoaded0 = True
        For i = LBound(fsKeyServer0) To UBound(fsKeyServer0)
            If fsKeyServer0(i) <> "" Then cboServer(0).AddItem fsKeyServer0(i)
        Next i
    End If
    
    cboServer(0) = IniFile_Read("ODBC", "ServerA", psIniFile)
    
    
'-----------------------------------------------------------------------------

    bRet = ReadSectionList("USERS_1", psIniFile, "User", fsKeyUser1())
    
    If bRet Then
        fbUserKeyLoaded1 = True
        For i = LBound(fsKeyUser1) To UBound(fsKeyUser1)
            If fsKeyUser1(i) <> "" Then cboUserID(1).AddItem fsKeyUser1(i)
        Next i
    End If
    
    cboUserID(1) = IniFile_Read("ODBC", "UserB", psIniFile)

    bRet = ReadSectionList("SERVERS_1", psIniFile, "Server", fsKeyServer1())
    
    If bRet Then
        fbServerKeyLoaded1 = True
        For i = LBound(fsKeyServer1) To UBound(fsKeyServer1)
            If fsKeyServer1(i) <> "" Then cboServer(1).AddItem fsKeyServer1(i)
        Next i
    End If
    
    cboServer(1) = IniFile_Read("ODBC", "ServerB", psIniFile)
    
End Sub

