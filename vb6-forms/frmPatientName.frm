VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmPatientName 
   Caption         =   "PATIENT FULL NAME"
   ClientHeight    =   2340
   ClientLeft      =   11730
   ClientTop       =   4740
   ClientWidth     =   5610
   FillColor       =   &H00800000&
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
   Icon            =   "frmPatientName.frx":0000
   LockControls    =   -1  'True
   ScaleHeight     =   2340
   ScaleWidth      =   5610
   Begin VB.Frame fraPatFullName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5595
      Begin PT_XP_Button.PT_XP cmdTitle 
         Height          =   315
         Left            =   3450
         TabIndex        =   13
         Top             =   240
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
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
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "...."
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPatCancel 
         Height          =   285
         Left            =   4380
         TabIndex        =   6
         Top             =   960
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
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
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "Cancel"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPatOk 
         Height          =   315
         Left            =   4380
         TabIndex        =   5
         Top             =   510
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
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
         Picture         =   "frmPatientName.frx":08CA
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "OK"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatientName.frx":0C64
         ShowFocus       =   -1  'True
      End
      Begin VB.TextBox txtFirstName 
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
         Left            =   1080
         TabIndex        =   1
         Top             =   630
         Width           =   2250
      End
      Begin VB.TextBox txtMiddleName 
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
         Left            =   1080
         TabIndex        =   2
         Top             =   990
         Width           =   2250
      End
      Begin VB.TextBox txtLastName 
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
         Left            =   1080
         TabIndex        =   3
         Top             =   1350
         Width           =   2250
      End
      Begin VB.ComboBox cboSuffix 
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1710
         Width           =   1485
      End
      Begin VB.ComboBox cboTitle 
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
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   2265
      End
      Begin PT_XP_Button.PT_XP cmdClear 
         Height          =   225
         Left            =   2580
         TabIndex        =   14
         ToolTipText     =   "Clear Suffix"
         Top             =   1770
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   397
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
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "...."
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblPatientTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title :"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   360
         TabIndex        =   12
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblFirstName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First :"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   360
         TabIndex        =   11
         Top             =   690
         Width           =   585
      End
      Begin VB.Label lblMiddleName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle :"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   180
         TabIndex        =   10
         Top             =   1020
         Width           =   765
      End
      Begin VB.Label lblLastName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last :"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   420
         TabIndex        =   9
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label lblSuffix 
         BackColor       =   &H00BFE8FD&
         BackStyle       =   0  'Transparent
         Caption         =   "Suffix :"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   270
         TabIndex        =   8
         Top             =   1740
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmPatientName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal Hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long

Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long


Private Function Clear_Fields()
    cboTitle.Clear
    txtFirstName.Text = vbNullString
    txtMiddleName.Text = vbNullString
    txtMiddleName.Text = vbNullString
    cboSuffix.Clear
    Exit Function
End Function
Private Sub cboSuffix_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    If Len(strList) >= 1 Then
    cboSuffix.ListIndex = SendMessage(cboSuffix.Hwnd, LB_FINDSTRING, -1, _
            ByVal strList)
    End If
End Sub


Private Sub cboTitle_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
strList = cboTitle.Text
    SearchList cboTitle, strList, True
  Exit Sub
End Sub


Private Sub cmdClear_Click()
    cboSuffix.ListIndex = -1
    clsPatient.Suffix = vbNullString
    Exit Sub
End Sub

Private Sub cmdPatCancel_Click()
        
    Unload Me
    Exit Sub
End Sub

Private Sub cmdPatOk_Click()
Dim Arr

    With clsPatient
        .Title = vbNullString
        .FirstName = vbNullString
        .LastName = vbNullString
        .MiddleName = vbNullString
        .Suffix = vbNullString
    End With
    clsPatient.Title = cboTitle.Text
    clsPatient.FirstName = txtFirstName.Text
    clsPatient.MiddleName = txtMiddleName.Text
    clsPatient.LastName = txtLastName.Text
    If Trim(cboSuffix.Text) <> vbNullString Then
        clsPatient.Suffix = Trim(cboSuffix.Text)
        clsPatient.Suffix = Replace(clsPatient.Suffix, ",", "")
        If TextFound("Jr", clsPatient.Suffix, 2, True) Or _
            TextFound("Sr", clsPatient.Suffix, 2, True) Then
            clsPatient.Suffix = IsProperCase(clsPatient.Suffix)
        Else
            clsPatient.Suffix = UCase(clsPatient.Suffix)
        End If
    Else
        If cboSuffix.ListIndex = -1 Then
            clsPatient.Suffix = vbNullString
        End If
    End If
    clsPatient.Suffix = Fix_Suffix(clsPatient.Suffix)
    clsPatient.FirstName = Fix_Name(clsPatient.FirstName)
    clsPatient.LastName = Fix_Name(clsPatient.LastName)
    If TextFound("'", clsPatient.LastName, Len("'"), True) Then
        Arr = Split(clsPatient.LastName, "'")
        If UBound(Arr) > 0 Then
            clsPatient.LastName = UCase(Arr(0))
            clsPatient.LastName = clsPatient.LastName & "'" & IsProperCase(Arr(1))
        End If
    End If
    If TextFound("(", clsPatient.FirstName, Len("("), True) Or _
        TextFound("*", clsPatient.FirstName, Len("*"), True) Then
        MsgBox "Unable to save patient name with symbols!", vbInformation, "PatientTrac Patient Information"
        Exit Sub
    End If
    If TextFound("(", clsPatient.MiddleName, Len("("), True) Or _
        TextFound("*", clsPatient.MiddleName, Len("*"), True) Then
        MsgBox "Unable to save patient name with symbols!", vbInformation, "PatientTrac Patient Information"
        Exit Sub
    End If
    If TextFound("(", clsPatient.LastName, Len("("), True) Or _
        TextFound("*", clsPatient.LastName, Len("*"), True) Then
        MsgBox "Unable to save patient name with symbols!", vbInformation, "PatientTrac Patient Information"
        Exit Sub
    End If
    Unload Me
    Exit Sub
End Sub

Private Sub cmdPatOk_GotFocus()
    cmdPatOk_KeyPress 13
End Sub


Private Sub cmdPatOk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       KeyAscii = 0
        cmdPatOk_Click
    End If
End Sub


Private Sub cmdTitle_Click()
    Dim sStr As String
    sStr = cboTitle.Text
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Title"
    frmNurseTemplates.fType = "Title"
    frmNurseTemplates.fSubDomain = "sd_Patient_Title"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    GetTerms cboTitle, "sd_Patient_Title"
    If Trim(sStr) <> vbNullString Then
       If cboTitle.ListIndex = -1 Then
          cboTitle.AddItem sStr
          cboTitle.ListIndex = cboTitle.NewIndex
       End If
    End If
    Exit Sub
End Sub


Private Sub Form_Activate()
Dim i%

    On Error GoTo Error
    Screen.MousePointer = vbHourglass
    
    If mFirst Then
        mFirst = False
    End If
    Clear_Fields
    GetTerms cboTitle, "sd_Patient_Title"
    GetTerms cboSuffix, "sd_Suffix"
    If clsPatient.Title <> vbNullString Then
        For i = 0 To cboTitle.listcount - 1
            If cboTitle.List(i) = clsPatient.Title Then
                cboTitle.Text = cboTitle.List(i)
            End If
          Next i
    End If
    clsPatient.Title = vbNullString
    If clsPatient.FirstName <> vbNullString Then
        txtFirstName.Text = Trim(clsPatient.FirstName)
        clsPatient.FirstName = vbNullString
    End If
    
    If clsPatient.MiddleName <> vbNullString Then
         txtMiddleName.Text = Trim(clsPatient.MiddleName)
        clsPatient.MiddleName = vbNullString
    End If
    
    If clsPatient.LastName <> vbNullString Then
        txtLastName.Text = Trim(clsPatient.LastName)
        clsPatient.LastName = vbNullString
    End If
    If Trim(clsPatient.Suffix) <> vbNullString Then
        For i = 0 To cboSuffix.listcount - 1
             If cboSuffix.List(i) = clsPatient.Suffix Then
                cboSuffix.Text = clsPatient.Suffix
                Exit For
            End If
        Next i
    Else
        cboSuffix.ListIndex = -1
    End If
    txtFirstName.SetFocus
    centerForm Me
    Screen.MousePointer = vbDefault
    Exit Sub
Error:
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
End Sub

Private Sub Form_Load()
    mFirst = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
 
  On Error Resume Next

    Unload Me
    Exit Sub
End Sub


