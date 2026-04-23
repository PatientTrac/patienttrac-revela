VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmPatientAdress 
   Caption         =   "PATIENT ADDRESS"
   ClientHeight    =   3060
   ClientLeft      =   9960
   ClientTop       =   3315
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatientAdress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmPatientAdress.frx":058A
   ScaleHeight     =   3060
   ScaleWidth      =   5610
   Begin VB.Frame fraAdd 
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
      Height          =   3015
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5595
      Begin VB.ComboBox cboZip 
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
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1710
         Width           =   1365
      End
      Begin VB.ComboBox cboCountry 
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
         Left            =   1230
         TabIndex        =   7
         Top             =   2430
         Width           =   2085
      End
      Begin VB.ComboBox cboCounty 
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
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2070
         Width           =   2835
      End
      Begin VB.ComboBox cboState 
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
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1350
         Width           =   2820
      End
      Begin VB.TextBox txtCity 
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
         Height          =   315
         Left            =   1230
         TabIndex        =   2
         Top             =   990
         Width           =   2820
      End
      Begin VB.TextBox txtAddress2 
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
         Height          =   315
         Left            =   1230
         TabIndex        =   1
         Top             =   630
         Width           =   2820
      End
      Begin VB.TextBox txtAddress1 
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
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   270
         Width           =   2820
      End
      Begin VB.TextBox txtZipCodeExt 
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
         Height          =   315
         Left            =   2865
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1710
         Width           =   1035
      End
      Begin PT_XP_Button.PT_XP cmdAddCancel 
         Height          =   315
         Left            =   4500
         TabIndex        =   9
         Top             =   1020
         Width           =   825
         _ExtentX        =   1455
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
         Caption         =   "Cancel"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatientAdress.frx":0CF4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAddSave 
         Height          =   315
         Left            =   4500
         TabIndex        =   8
         Top             =   570
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
         Picture         =   "frmPatientAdress.frx":1046
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "OK"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmPatientAdress.frx":13E0
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblCountry 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   2475
         Width           =   870
      End
      Begin VB.Label lblProvince 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "County:"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   300
         TabIndex        =   17
         Top             =   2115
         Width           =   765
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State:"
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   510
         TabIndex        =   16
         Top             =   1410
         Width           =   570
      End
      Begin VB.Label lblCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City:"
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   630
         TabIndex        =   15
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apt./Suite:"
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   1
         Left            =   60
         TabIndex        =   14
         Top             =   690
         Width           =   1065
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   13
         Top             =   315
         Width           =   840
      End
      Begin VB.Label lblZipCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zip Code:"
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   12
         Top             =   1785
         Width           =   870
      End
      Begin VB.Label lblDash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2685
         TabIndex        =   11
         Top             =   1725
         Width           =   60
      End
   End
End
Attribute VB_Name = "frmPatientAdress"
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


Private Sub LookUp_PostalCity()
Dim Idx As Integer
Dim strSQL As String
Dim cnZip As New ADODB.Connection
Dim rsZip As New ADODB.Recordset

    On Error GoTo Error_LookUp_PostalCity
    Screen.MousePointer = vbHourglass
    cboCounty.Clear
    If Len(cboState.Text) = 2 Then
        Get_StateName (cboState.Text)
    End If
    txtCity.Text = Replace(txtCity.Text, " - ", "-")
    txtCity.Text = Trim(txtCity.Text)
    If cnZip.State = adStateClosed Then cnZip.open psConnect(1)
    strSQL = "Lookup_City_County_State_Zip '" & Trim(cboZip.Text) & "', '" & Trim(cboState.Text) & "'"
    Set rsZip = cnZip.Execute(strSQL)
    Do Until rsZip.EOF
        If Not IsNull(rsZip.Fields("COUNTY_NAME")) Then
            cboCounty.AddItem IsProperCase(rsZip.Fields("COUNTY_NAME"))
            cboCounty.ListIndex = 0
        End If
        rsZip.MoveNext
    Loop
    If cboZip.listcount <= 0 Then
        strSQL = "Lookup_ZipCode_AltCity_State '" & Trim(txtCity.Text) & "', '" & Trim(cboState.Text) & "'"
        Set rsZip = cnZip.Execute(strSQL)
        Do Until rsZip.EOF
            If Not IsNull(rsZip.Fields("ZIPCODE_NO")) Then
                cboZip.AddItem rsZip.Fields("ZIPCODE_NO")
            End If
            rsZip.MoveNext
        Loop
    End If
     If cnZip.State = adStateOpen Then
        cnZip.Close
        Set cnZip = Nothing
        Set rsZip = Nothing
     End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_LookUp_PostalCity:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "LookUp_PostalCity", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub



Private Sub LookUp_Zip()
Dim Idx As Integer
Dim strSQL As String
Dim cnZip As New ADODB.Connection
Dim rsZip As New ADODB.Recordset

    On Error GoTo Error_LookUp_Zip
    Screen.MousePointer = vbHourglass
    cboZip.Clear
    If Len(cboState.Text) = 2 Then
        Get_StateName (cboState.Text)
    End If
    If cnZip.State = adStateClosed Then cnZip.open psConnect(1)
    strSQL = "Lookup_ZipCode_City_State '" & Trim(txtCity.Text) & "', '" & Trim(cboState.Text) & "'"
    Set rsZip = cnZip.Execute(strSQL)
    Do Until rsZip.EOF
        cboZip.AddItem rsZip.Fields("ZIPCODE_NO")
        rsZip.MoveNext
    Loop
    If cboZip.listcount <= 0 Then
        strSQL = "Lookup_ZipCode_AltCity_State '" & Trim(txtCity.Text) & "', '" & Trim(cboState.Text) & "'"
        Set rsZip = cnZip.Execute(strSQL)
        Do Until rsZip.EOF
            cboZip.AddItem rsZip.Fields("ZIPCODE_NO")
            rsZip.MoveNext
        Loop
    End If
     If cnZip.State = adStateOpen Then
        cnZip.Close
        Set cnZip = Nothing
        Set rsZip = Nothing
     End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_LookUp_Zip:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "LookUp_Zip", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Function Clear_Fields()
    txtAddress1.Text = vbNullString
    txtAddress2.Text = vbNullString
    txtCity.Text = vbNullString
    cboZip.Clear
    txtZipCodeExt.Text = vbNullString
    cboState.Clear
    cboCounty.Clear
    cboCountry.Clear
    Exit Function
End Function


Private Sub cboCountry_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    strList = cboCountry.Text
    cboCountry.ListIndex = SendMessage(cboCountry.Hwnd, LB_FINDSTRING, -1, _
            ByVal strList)
  Exit Sub
End Sub


Private Sub cboCounty_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    strList = cboCounty.Text
    If Trim(strList) <> vbNullString Then
        cboCounty.ListIndex = SendMessage(cboCounty.Hwnd, LB_FINDSTRING, -1, _
                ByVal strList)
    End If
  Exit Sub
End Sub


Private Sub cboState_Click()
    If txtCity.Text <> vbNullString Then LookUp_Zip
End Sub

Private Sub cboState_KeyUp(KeyCode As Integer, Shift As Integer)
    If txtCity.Text <> vbNullString Then LookUp_Zip
End Sub


Private Sub cboZip_Click()
    On Error Resume Next
      If cboZip.ListIndex >= 0 Then
        LookUp_PostalCity
        If txtZipCodeExt.Enabled Then txtZipCodeExt.SetFocus
      End If
      Exit Sub
End Sub

Private Sub cmdAddCancel_Click()
    Unload Me
End Sub

Private Sub cmdAddSave_Click()

        With clsPatient
            .Address = vbNullString
            .Address2 = vbNullString
            .City = vbNullString
            .State = vbNullString
            .Zip = vbNullString
            .Zip_Ext = vbNullString
            .County = vbNullString
            .Country = vbNullString
        End With
        If Trim(txtAddress1.Text) <> vbNullString Then
            clsPatient.Address = Trim(txtAddress1.Text)
        Else
            MsgBox "Patient's Address is a mandatory field!, Please enter an address!", vbInformation, "PatientTrac Address"
            Exit Sub
        End If
        If txtAddress2.Text <> vbNullString Then clsPatient.Address2 = IsProperCase(txtAddress2.Text)
        If txtCity.Text <> vbNullString Then clsPatient.City = Trim(txtCity.Text)
        If cboState.Text <> vbNullString Then clsPatient.State = Trim(cboState.Text)
        If cboZip.Text <> vbNullString Then clsPatient.Zip = Trim(cboZip.Text)
        If txtZipCodeExt.Text <> vbNullString Then clsPatient.Zip_Ext = Trim(txtZipCodeExt.Text)
        If cboCounty.Text <> vbNullString Then clsPatient.County = cboCounty.Text
        If cboCountry.Text <> vbNullString Then clsPatient.Country = UCase(cboCountry.Text)
        clsPatient.Address = Fix_Name(clsPatient.Address)
        clsPatient.City = Fix_Name(clsPatient.City)
        clsPatient.City = Replace(clsPatient.City, " - ", "-")
        clsPatient.City = Trim(clsPatient.City)
        clsPatient.County = Fix_Name(clsPatient.County)
        Unload Me
    Exit Sub
End Sub

Private Sub cmdAddSave_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       KeyAscii = 0
'        cmdAddSave_Click
'    End If
End Sub


Private Sub Form_Activate()
Dim rsData As New ADODB.Recordset
Dim cnData As New ADODB.Connection
Dim strSQL As String
Dim strState As String
Dim Idx%

    On Error GoTo Error_Form_Activate
    If mFirst Then
        mFirst = False
    End If
    Clear_Fields
    If cnData.State = adStateClosed Then cnData.open psConnect(1)
    strSQL = "Get_States"
    Set rsData = cnData.Execute(strSQL)
    Do Until rsData.EOF
        If Not IsNull(Trim(rsData.Fields("STATE_NAME"))) Then
            cboState.AddItem Trim(rsData.Fields("STATE_NAME"))
        End If
        rsData.MoveNext
    Loop
    strSQL = "Get_Countries"
    Set rsData = cnData.Execute(strSQL)
    Do Until rsData.EOF
        If Not IsNull(rsData.Fields("COUNTRY_NAME")) Then
            cboCountry.AddItem rsData.Fields("COUNTRY_NAME")
        End If
        rsData.MoveNext
    Loop
    For Idx = 1 To cboCountry.listcount - 1
        If cboCountry.List(Idx) = "United States" Then
            cboCountry.Text = cboCountry.List(Idx)
        End If
    Next Idx
    If Trim(clsPatient.State) = vbNullString Then
        strState = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "DefaultState")
    Else
        strState = clsPatient.State
    End If
    If strState <> vbNullString Then
        For Idx = 0 To cboState.listcount - 1
            If cboState.List(Idx) = strState Then
                 cboState.Text = cboState.List(Idx)
                 Exit For
            End If
        Next Idx
    End If
    If clsPatient.Address <> vbNullString Then txtAddress1.Text = Trim(clsPatient.Address)
    If clsPatient.Address2 <> vbNullString Then txtAddress2.Text = Trim(clsPatient.Address2)
    clsPatient.City = Replace(clsPatient.City, " - ", "-")
    clsPatient.City = Trim(clsPatient.City)
    If clsPatient.City <> vbNullString Then txtCity.Text = Trim(clsPatient.City)
    If clsPatient.State <> vbNullString Then
        For Idx = 0 To cboState.listcount - 1
            If cboState.List(Idx) = Trim(strState) Then
                 cboState.Text = cboState.List(Idx)
                 Exit For
            End If
        Next Idx
    End If
    If Trim(clsPatient.Zip) <> vbNullString Then
        LookUp_Zip
        If cboZip.listcount >= 0 Then
            For Idx = 0 To cboZip.listcount - 1
                If cboZip.List(Idx) = Trim(clsPatient.Zip) Then
                    cboZip.Text = Trim(clsPatient.Zip)
                    Exit For
                End If
            Next Idx
            If cboZip.Text = vbNullString And cboZip.listcount >= 0 Then
                cboZip.BackColor = &HC0E0FF
                cboZip.ToolTipText = "Erroneous Zip Code"
            End If
        End If
    End If
    If clsPatient.Zip_Ext <> vbNullString Then txtZipCodeExt.Text = Trim(clsPatient.Zip_Ext)
    If clsPatient.Country <> vbNullString Then
        For Idx = 1 To cboCountry.listcount - 1
            If cboCountry.List(Idx) = Trim(clsPatient.Country) Then
                cboCountry.Text = cboCountry.List(Idx)
            End If
        Next Idx
    End If
    If clsPatient.County <> vbNullString Then
        For Idx = 1 To cboCounty.listcount - 1
            If cboCounty.List(Idx) = Trim(clsPatient.County) Then
                cboCounty.Text = cboCounty.List(Idx)
                Exit For
            End If
        Next Idx
    End If
    If cnData.State = adStateOpen Then
        cnData.Close
        Set cnData = Nothing
        Set rsData = Nothing
    End If
    centerForm Me
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Form_Activate:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Form_Activate", App.EXEName & "." & Me.Name
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


Private Sub cboZip_KeyUp(KeyCode As Integer, Shift As Integer)
      If cboZip.Text <> vbNullString Then
         If txtZipCodeExt.Enabled Then txtZipCodeExt.SetFocus
      End If
End Sub


Private Sub txtCity_Change()
    If Trim(cboState.Text) <> vbNullString Then LookUp_Zip
    Exit Sub
End Sub

Private Sub txtCity_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(cboState.Text) <> vbNullString Then LookUp_Zip
    Exit Sub
End Sub


