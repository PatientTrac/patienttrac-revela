VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmListPatient 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Patient LookUp"
   ClientHeight    =   6570
   ClientLeft      =   5925
   ClientTop       =   1680
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListPatient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   8670
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7860
      Top             =   6060
   End
   Begin VB.ComboBox cboSearchFilter 
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
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   3615
   End
   Begin VB.Frame fraPic 
      BackColor       =   &H00E0E0E0&
      Height          =   2625
      Left            =   5430
      TabIndex        =   10
      Top             =   2880
      Width           =   2745
      Begin VB.Image imgPatient 
         Height          =   2415
         Left            =   60
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2625
      End
      Begin VB.Label lblPhoto 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NO  PHOTO AVAILABLE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   735
         Left            =   780
         TabIndex        =   11
         Top             =   1110
         Visible         =   0   'False
         Width           =   1665
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   7650
      Picture         =   "frmListPatient.frx":058A
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   8
      Top             =   2370
      Width           =   525
   End
   Begin VB.ListBox lstEncounters 
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
      Height          =   1260
      Left            =   5010
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1020
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.TextBox txtList 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1365
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4350
      Width           =   4875
   End
   Begin VB.TextBox cctSearch 
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   1020
      Width           =   4875
   End
   Begin VB.ListBox lstListBox 
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
      Height          =   2700
      Left            =   60
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1530
      Width           =   4875
   End
   Begin PT_XP_Button.PT_XP cmdAddList 
      Height          =   345
      Left            =   2850
      TabIndex        =   6
      Top             =   6000
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
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
      Picture         =   "frmListPatient.frx":0E54
      AppearanceThemes=   3
      BorderDrawEvent =   1
      TransparentColor=   14737632
      Caption         =   "Select Patient"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmListPatient.frx":13EE
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdCloseList 
      Height          =   345
      Left            =   4695
      TabIndex        =   7
      Top             =   6000
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
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
      Picture         =   "frmListPatient.frx":1740
      AppearanceThemes=   3
      BorderDrawEvent =   1
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmListPatient.frx":2152
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search by:"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   90
      TabIndex        =   13
      Top             =   210
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Patient's Photo"
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
      Left            =   6030
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblSelEncounter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Encounters:"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5040
      TabIndex        =   5
      Top             =   810
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblInstr 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name, First Name"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   780
      Width           =   2415
   End
End
Attribute VB_Name = "frmListPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngtop As Long
Public mlngleft As Long
Public mstrTerms As String
Public mstrSelectedTerms As String
Public mlngListItemData As Long
Public mstrCurrentData As String
Private mbooByPass As Boolean
Public mstrListName As String
Public mFirst As Boolean
Public mbooCaption As Boolean
Public mstrItemClick As String
Public mstrHomePhone As String
Public mstrAreaCode As String
Public mstrCellPhone As String
Private mbooSearch As Boolean
Public mlngListEnctr As Long
Public mbShowEnctr As Boolean
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal Hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long

Private Function Look_UpPatient(vstrData As String, Optional vstrType As String)
Dim strSQL As String
Dim cnICD As New ADODB.Connection
Dim rsICD As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strText As String
Dim i%
Dim Idx As Integer
Dim strPrefName As String
Dim lngCode As Long
Dim strDesc As String
Dim lngPatID As Long

    On Error GoTo ERROR_Look_UpDiag
    Screen.MousePointer = vbHourglass
    LockWindowUpdate Hwnd
    lstListBox.Clear
    vstrData = Trim(cctSearch.Text)
    vstrData = Replace(vstrData, "'", "''")
    If Trim(vstrData) <> vbNullString Then
        If cnICD.State = adStateClosed Then cnICD.open psConnect(0)
            If vstrType = "EXTERNAL ID" Then
                strSQL = "Select * FROM PATIENT WHERE PATIENT_EXT_REF  like '" & vstrData & "%'"
            ElseIf vstrType = "EMR ID" Then
                strSQL = "Select FIRST_NAME, LAST_NAME, PATIENT_ID FROM PATIENT WHERE PATIENT_ID like '" & vstrData & "%'"
            ElseIf vstrType = "SOCIAL SECURITY" Then
                strSQL = "Select FIRST_NAME, LAST_NAME, PATIENT_ID FROM PATIENT WHERE SOCIAL_SECURITY  like '" & vstrData & "%'"
            ElseIf vstrType = "LAST NAME" Then
                strSQL = "Select FIRST_NAME, LAST_NAME, PATIENT_ID FROM PATIENT WHERE LAST_NAME like '" & vstrData & "%'"
            End If
            If strSQL <> vbNullString Then
                Set rsICD = cnICD.Execute(strSQL)
                 With rsICD
                     Do Until .EOF
                        Clear_UdtPatient
                        lngPatID = 0
                        lngPatID = (!Patient_ID)
                        .MoveNext
                        Get_PatientName lngPatID
                        If Trim(udtPatientName.Suffix) <> vbNullString Then
                            udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                        Else
                            udtPatientName.Name = udtPatientName.LastName & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                        End If
                        If Trim(udtPatientName.Name) <> vbNullString Then
                            lstListBox.AddItem Trim(udtPatientName.Name)
                            lstListBox.ItemData(lstListBox.NewIndex) = lngPatID
                        End If
                     Loop
                  End With
            End If
    End If
    If cnICD.State = adStateOpen Then
        cnICD.Close
        Set cnICD = Nothing
        Set rsICD = Nothing
    End If
    Clear_UdtPatient
    LockWindowUpdate 0
    Screen.MousePointer = vbDefault
    Exit Function
ERROR_Look_UpDiag:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Look_UpDiag", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub Load_Enctr(vlngPatientID As Long)
Dim strSQL  As String
Dim i As Integer
Dim lstrRemarks As String
Dim cnEnctr As New ADODB.Connection
Dim rsEnctr As New ADODB.Recordset


    If vlngPatientID = 0 Then Exit Sub
    If cnEnctr.State = adStateClosed Then cnEnctr.open psConnect(0)
    strSQL = "Get_Encounter_Patient_ID " & vlngPatientID
    Set rsEnctr = cnEnctr.Execute(strSQL)
    With rsEnctr
        Do Until .EOF
           If (!ENCNTR_START_DATE) <> vbNullString Then
                lstEncounters.AddItem (!ENCNTR_START_DATE)
                lstEncounters.ItemData(lstEncounters.NewIndex) = (!Encounter_ID)
            End If
        .MoveNext
        Loop
    End With
    If cnEnctr.State = adStateOpen Then
        cnEnctr.Close
        Set cnEnctr = Nothing
        Set rsEnctr = Nothing
    End If
    Exit Sub
End Sub
Private Function FixPatient(vlngPatientID As Long)
Dim i%
Dim strSQL As String
Dim cnFixPat As New ADODB.Connection
Dim rsFixPat As New ADODB.Recordset

    
    If vlngPatientID > 0 Then
        If cnFixPat.State = adStateClosed Then cnFixPat.open psConnect(0)
        strSQL = "update patient set purge = 0 where patient_id = " & vlngPatientID
        Set rsFixPat = cnFixPat.Execute(strSQL)
        Read_PatientData
        If cnFixPat.State = adStateOpen Then
            cnFixPat.Close
            Set cnFixPat = Nothing
        End If
    End If
    Exit Function
End Function


Public Function Read_PatientData()
Dim strSQL As String
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim cnPhys As New ADODB.Connection
Dim rsPhys As New ADODB.Recordset
Dim strSuffix As String
Dim lngRefID As Long
Dim i%
Dim strName As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error_Read_PatientData
    frmListPatient.Caption = "Search Patients"
    If lstListBox.listcount >= 0 Then
        lstListBox.Clear
    End If
    cmdAddList.Caption = "Select Patient"
    If cnPhys.State = adStateClosed Then cnPhys.open psConnect(0)
    strSQL = "SELECT * FROM PATIENT WHERE PURGE = 0 ORDER BY LAST_NAME"
    Set rsPhys = cnPhys.Execute(strSQL)
    With rsPhys
        Do Until .EOF
            strLast = vbNullString
            strFirst = vbNullString
            strMiddle = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            lngRefID = 0
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "PATIENT_ID"
                                lngRefID = Fld.Value
                            Case "LAST_NAME"
                                strLast = Trim(Fld.Value)
                            Case "FIRST_NAME"
                                 strFirst = Trim(Fld.Value)
                            Case "MIDDLE_NAME"
                                 strMiddle = Trim(Fld.Value)
                            Case "SUFFIX"
                                strSuffix = UCase(Fld.Value)
                                strSuffix = Replace(strSuffix, ",", "")
                                If TextFound("Jr", strSuffix, 2, True) Or _
                                    TextFound("Sr", strSuffix, 2, True) Then
                                    strSuffix = IsProperCase(strSuffix)
                                Else
                                    strSuffix = UCase(strSuffix)
                                End If
                                If TextFound("PHD", strSuffix, Len("phd"), True) Or _
                                    TextFound("PHD", strSuffix, 4, True) Then
                                    strSuffix = "Ph.D."
                                End If
                            Case "TITLE"
                                strTitle = Trim(Fld.Value)
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If lngRefID > 0 And Trim(strLast) <> vbNullString Then
                If Trim(strSuffix) = vbNullString Then
                    If Trim(strTitle) <> vbNullString Then
                        strName = Trim(strLast) & ", " & Trim(strFirst) & " " & Trim(strMiddle) & ", " & Trim(strTitle)
                    Else
                        strName = Trim(strLast) & ", " & Trim(strFirst) & " " & Trim(strMiddle)
                    End If
                Else
                    If Trim(strTitle) <> vbNullString Then
                        strName = Trim(strLast) & ", " & Trim(strFirst) & " " & Trim(strMiddle) & ", " & Trim(strSuffix) & " " & Trim(strTitle)
                    Else
                        strName = Trim(strLast) & ", " & Trim(strFirst) & " " & Trim(strMiddle) & ", " & Trim(strSuffix)
                    End If
                End If
                If strName <> vbNullString Then
                    lstListBox.AddItem Trim(strName)
                    lstListBox.ItemData(lstListBox.NewIndex) = lngRefID
                End If
            End If
        Loop
    End With
    cboSearchFilter.Clear
    If cboSearchFilter.listcount <= 0 Then
        If mstrPractice = "Psychiatry" Then
            cboSearchFilter.AddItem "LAST NAME"
            cboSearchFilter.AddItem "EMR ID"
            cboSearchFilter.AddItem "EXTERNAL ID"
           ' cboSearchFilter.AddItem "SOCIAL SECURITY"
            cboSearchFilter.AddItem "SURGERY ID"
        Else
            cboSearchFilter.AddItem "LAST NAME"
            cboSearchFilter.AddItem "EMR ID"
            cboSearchFilter.AddItem "FINANCIAL ACCT NO"
           ' cboSearchFilter.AddItem "SOCIAL SECURITY"
            cboSearchFilter.AddItem "SURGERY ID"
        End If
    End If
    If cboSearchFilter.listcount >= 0 Then
        For i = 0 To cboSearchFilter.listcount - 1
            If cboSearchFilter.List(i) = "LAST NAME" Then
                mbooByPass = True
                cboSearchFilter.Text = cboSearchFilter.List(i)
                Exit For
            End If
        Next i
    End If
    mbooByPass = False
    If cnPhys.State = adStateClosed Then
        cnPhys.Close
        Set cnPhys = Nothing
        Set rsPhys = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_PatientData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_PatientData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Public Function Read_ReferralData()
Dim strSQL As String
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim cnPhys As New ADODB.Connection
Dim rsPhys As New ADODB.Recordset
Dim strSuffix As String
Dim lngRefID As Long
Dim i%

    Screen.MousePointer = vbHourglass
    On Error GoTo Error_Read_ReferralData
    frmListPatient.Caption = "Search Physicians"
    If lstListBox.listcount >= 0 Then
        lstListBox.Clear
    End If
    cmdAddList.Caption = "Select Physician"
    If cnPhys.State = adStateClosed Then cnPhys.open psConnect(0)
    strSQL = "SELECT * FROM PHYSICIAN_REFERRAL ORDER BY LAST_NAME"
    Set rsPhys = cnPhys.Execute(strSQL)
    With rsPhys
        Do Until .EOF
            strLast = vbNullString
            strFirst = vbNullString
            strMiddle = vbNullString
            strSuffix = vbNullString
            strTitle = vbNullString
            lngRefID = 0
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "REFERRER_ID"
                                lngRefID = Fld.Value
                            Case "LAST_NAME"
                                strLast = Fld.Value
                            Case "FIRST_NAME"
                                strFirst = Fld.Value
                            Case "MIDDLE_NAME"
                                strMiddle = Fld.Value
                            Case "SUFFIX"
                                strSuffix = Fld.Value
                            Case "TITLE"
                                strTitle = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
              If lngRefID > 0 Then
                If Trim(strSuffix) = vbNullString Then
                    lstListBox.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strTitle)
                Else
                    lstListBox.AddItem Trim(strLast) & ",  " & strFirst & " " & Trim(strMiddle) & ", " & Trim(strSuffix) & " " & Trim(strTitle)
                End If
                lstListBox.ItemData(lstListBox.NewIndex) = lngRefID
              End If
        Loop
    End With
    If cboSearchFilter.listcount <= 0 Then
        cboSearchFilter.AddItem "PHYSICIAN LAST NAME"
        cboSearchFilter.AddItem "REFERRAL ID"
        cboSearchFilter.AddItem "PRACTICE NAME"
    End If
    If cboSearchFilter.listcount >= 0 Then
        For i = 0 To cboSearchFilter.listcount - 1
            If cboSearchFilter.List(i) = "PHYSICIAN LAST NAME" Then
                mbooByPass = True
                cboSearchFilter.Text = cboSearchFilter.List(i)
                Exit For
            End If
        Next i
    End If
    mbooByPass = False
    If cnPhys.State = adStateClosed Then
        cnPhys.Close
        Set cnPhys = Nothing
        Set rsPhys = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_ReferralData:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_ReferralData", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Public Sub Read_VendorData()
Dim cnSearch As New ADODB.Connection
Dim rsSearch As New ADODB.Recordset
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim i%

    On Error GoTo ERROR_Read_VendorData
    Screen.MousePointer = vbHourglass
    frmListPatient.Caption = "Search Vendors"
    If lstListBox.listcount >= 0 Then
        lstListBox.Clear
    End If
    If cnSearch.State = adStateClosed Then cnSearch.open psConnect(0)
    strSQL = "SELECT * FROM VENDORS ORDER BY VENDOR_NAME"
    Set rsSearch = cnSearch.Execute(strSQL)
    With rsSearch
        Do Until .EOF
            Clear_Vendor
            Set Flds = .Fields
             For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case Fld.Name
                            Case "VENDOR_NAME"
                                clsVendor.Name = Fld.Value
                            Case "VENDOR_ID"
                                clsVendor.VendorID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If clsVendor.VendorID > 0 Then
                If Trim(clsVendor.Name) <> vbNullString Then
                    lstListBox.AddItem Trim(clsVendor.Name)
                    lstListBox.ItemData(lstListBox.NewIndex) = clsVendor.VendorID
                End If
            End If
        Loop
    End With
    cboSearchFilter.Clear
    If cboSearchFilter.listcount <= 0 Then
        cboSearchFilter.AddItem "VENDOR NAME"
        cboSearchFilter.AddItem "VENDOR ID"
        cboSearchFilter.AddItem "CONTACT NAME"
        cboSearchFilter.AddItem "VENDOR TYPE"
    End If
    If cboSearchFilter.listcount >= 0 Then
        For i = 0 To cboSearchFilter.listcount - 1
            If cboSearchFilter.List(i) = "VENDOR NAME" Then
                mbooByPass = True
                cboSearchFilter.Text = cboSearchFilter.List(i)
                Exit For
            End If
        Next i
    End If
    mbooByPass = False
    If rsSearch.State = adStateClosed Then
        cnSearch.Close
        Set cnSearch = Nothing
        Set rsSearch = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERROR_Read_VendorData:
    Write_Error Err.Description, Err.Number, strSQL, "Read_VendorData", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Function Search_Patient()
Dim strSQL As String
Dim strSearch As String
Dim Cnt As Long
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngPatID As Long
Dim strTitle As String
Dim Arr
Dim strFirst As String

    On Error GoTo Error_Search_Patient
    Screen.MousePointer = vbHourglass
    Clear_UdtPatient
    If cctSearch.Text = vbNullString Then
        MsgBox "Please enter a search item first.", vbInformation, "PatientTrac Patient Search"
        Exit Function
    End If
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    If Trim(cctSearch.Text) = vbNullString Then
        MsgBox "Please enter a search term first!", vbInformation, "PatientTrac Patient Search"
    End If
    strSearch = cctSearch.Text
    If TextFound("'", strSearch, Len("'"), True) Then
        strSearch = Replace(strSearch, "'", "%")
    End If
    If cboSearchFilter.Text = "FINANCIAL ACCT NO" Then
        strTitle = "EXTERNAL ID"
    Else
        strTitle = UCase(cboSearchFilter.Text)
    End If
    Select Case Trim(UCase(strTitle))
        Case "LAST NAME"
            If TextFound(",", strSearch, Len(","), True) Then
                Arr = Split(strSearch, ",")
                If UBound(Arr) > 0 Then
                    strFirst = Trim(Arr(1))
                    strSearch = Arr(0)
                    strFirst = Replace(strFirst, " ", "")
                    strSearch = Replace(strSearch, " ", "")
                    strSearch = Trim(strSearch)
                End If
            End If
            strFirst = Replace(strFirst, "'", "''")
            If Trim(strFirst) <> vbNullString Then
                 strSQL = "Search_Patient_FirstName_Last_Name   '" & strSearch & "%', '%" & strFirst & "%'"
            Else
                strSQL = "Search_Patient_Row_Count_Last_Name 25, '" & strSearch & "%'"
            End If
        Case "EMR ID"
            strSQL = "Search_Patient_EMR_ID '" & strSearch & "%'"
        Case "SOCIAL SECURITY"
            strSQL = "Search_Patient_Social_Security '" & strSearch & "%'"
        Case "EXTERNAL ID"
            strSQL = "Search_Patient_External_ID '" & strSearch & "%'"
        End Select
        If strSQL <> vbNullString Then
            Set rsPat = cnPat.Execute(strSQL)
            Do Until rsPat.EOF
                Clear_UdtPatient
                lngPatID = 0
                lngPatID = rsPat.Fields("PATIENT_ID")
                rsPat.MoveNext
                Get_PatientName lngPatID
                If Trim(udtPatientName.Suffix) <> vbNullString Then
                    udtPatientName.Name = udtPatientName.LastName & ", " & udtPatientName.Suffix & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                Else
                    udtPatientName.Name = udtPatientName.LastName & " " & udtPatientName.FirstName & " " & udtPatientName.MiddleName
                End If
                If strTitle = "EXTERNAL ID" Or _
                    strTitle = "EMR ID" And Trim(udtPatientName.Name) <> vbNullString Then
                        lstListBox.AddItem lngPatID & " - " & Trim(udtPatientName.Name)
                        lstListBox.ItemData(lstListBox.NewIndex) = lngPatID
                Else
                    If Trim(udtPatientName.Name) <> vbNullString Then
                        lstListBox.AddItem Trim(udtPatientName.Name)
                        lstListBox.ItemData(lstListBox.NewIndex) = lngPatID
                    End If
                End If
            Loop
            If lstListBox.listcount < 0 Then
                MsgBox "No Patients found!", vbInformation, "PatientTrac Patient Search"
                Clear_UdtPatient
            End If
        End If
    Clear_UdtPatient
    mbooByPass = False
    If cnPat.State = adStateOpen Then
        cnPat.Close
        Set cnPat = Nothing
        Set rsPat = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Search_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Search_Patient", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function
Private Function Search_Physician()
Dim strSQL As String
Dim strSearch As String
Dim Cnt As Long
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strTitle As String

    On Error GoTo Error_Search_Physician
    Screen.MousePointer = vbHourglass
    If cctSearch.Text = vbNullString Then
        MsgBox "Please enter a search item first.", vbInformation, "PatientTrac Physician Search"
        Exit Function
    End If
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    If Trim(cctSearch.Text) = vbNullString Then
        MsgBox "Please enter a search term first!", vbInformation, "PatientTrac Physician Search"
    End If
    strSearch = cctSearch.Text
    strSearch = Replace(strSearch, "'", "''")
    Select Case Trim(strTitle)
        Case "PHYSICIAN LAST NAME"
            strSQL = "Search_Referral_Last_Name'" & strSearch & "%'"
        Case "REFERRAL ID"
            strSQL = "Search_Referral_ID '" & strSearch & "%'"
        Case "PRACTICE NAME"
            strSQL = "Search_Referral_PracticeName '" & strSearch & "%'"
        End Select
        If strSQL <> vbNullString Then
            Set rsPat = cnPat.Execute(strSQL)
            Do Until rsPat.EOF
                Clear_ReferralStrings
                Set Flds = rsPat.Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "FIRST_NAME"
                                    clsReferral.FirstName = Fix_Name(Fld.Value)
                                Case "MIDDLE_NAME"
                                    clsReferral.MiddleName = Fld.Value
                                Case "LAST_NAME"
                                    clsReferral.LastName = Fix_Name(Fld.Value)
                                Case "SUFFIX"
                                    clsReferral.Suffix = Fix_Suffix(Fld.Value)
                                Case "TITLE"
                                    clsReferral.Title = Fix_Title(Fld.Value)
                                Case "REFERRER_ID"
                                    clsReferral.ReferralID = Fld.Value
                            End Select
                        End If
                    End If
                Next
                rsPat.MoveNext
                If Trim(clsReferral.Suffix) <> vbNullString Then
                    clsReferral.Physician = clsReferral.LastName & ", " & clsReferral.Suffix & " " & clsReferral.FirstName & " " & clsReferral.MiddleName
                Else
                    clsReferral.Physician = clsReferral.LastName & ", " & clsReferral.FirstName & " " & clsReferral.MiddleName
                End If
                If strTitle = "REFERRAL ID" Then
                        lstListBox.AddItem clsReferral.ReferralID & " - " & Trim(clsReferral.Physician)
                        lstListBox.ItemData(lstListBox.NewIndex) = clsReferral.ReferralID
                Else
                    If Trim(clsReferral.Physician) <> vbNullString Then
                        lstListBox.AddItem Trim(clsReferral.Physician)
                        lstListBox.ItemData(lstListBox.NewIndex) = clsReferral.ReferralID
                    End If
                End If
            Loop
            If lstListBox.listcount < 0 Then
                MsgBox "No Physicians found!", vbInformation, "PatientTrac Physician Search"
                Clear_ReferralStrings
            End If
        End If
    mbooByPass = False
    Clear_ReferralStrings
    If cnPat.State = adStateOpen Then
        cnPat.Close
        Set cnPat = Nothing
        Set rsPat = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Search_Physician:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Search_Physician", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function
Private Function Search_Vendor()
Dim strSQL As String
Dim strSearch As String
Dim Cnt As Long
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngPatID As Long
Dim strTitle As String

    On Error GoTo Error_Search_Vendor
    Screen.MousePointer = vbHourglass
    Clear_UdtPatient
    If cctSearch.Text = vbNullString Then
        MsgBox "Please enter a search item first.", vbInformation, "PatientTrac Vendor Search"
        Exit Function
    End If
    If cnPat.State = adStateClosed Then cnPat.open psConnect(0)
    If Trim(cctSearch.Text) = vbNullString Then
        MsgBox "Please enter a search term first!", vbInformation, "PatientTrac Vendor Search"
    End If
    strSearch = cctSearch.Text
    strSearch = Replace(strSearch, "'", "''")
        Select Case strSearch
            Case "VENDOR NAME"
                strSQL = "Search_Vendor_LastName '" & strSearch & "%'"
            Case "VENDOR ID"
                strSQL = "Search_Vendor_Vendor_ID '" & strSearch & "%'"
            Case "CONTACT TITLE"
                strSQL = "Search_Vendor_Vendor_Contact '" & strSearch & "%'"
            Case "VENDOR TYPE"
                strSQL = "Search_Vendor_Type '" & strSearch & "%'"
        End Select
        If strSQL <> vbNullString Then
            Set rsPat = cnPat.Execute(strSQL)
            Do Until rsPat.EOF
                Set Flds = rsPat.Fields
                Clear_Vendor
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Trim(Fld.Value) <> vbNullString Then
                            Select Case UCase(Fld.Name)
                                Case "VENDOR_ID"
                                    lngPatID = Fld.Value
                                Case "VENDOR_NAME"
                                    clsVendor.Name = Trim(Fld.Value)
                                Case "VENDOR_CONTACT"
                                    clsVendor.Contact = Trim(Fld.Value)
                                Case "CONTACT_TITLE"
                                    clsVendor.Contact_Title = Trim(Fld.Value)
                                Case "VENDOR_TELEPHONE"
                                    clsVendor.Telephone = Trim(Fld.Value)
                                Case "VENDOR_FACSIMILE"
                                    clsVendor.Facsimile = Trim(Fld.Value)
                            End Select
                        End If
                    End If
                Next
                rsPat.MoveNext
                If Trim(clsVendor.Name) <> vbNullString Then
                    lstListBox.AddItem Trim(clsVendor.Name)
                    lstListBox.ItemData(lstListBox.NewIndex) = lngPatID
                End If
            Loop
            If lstListBox.listcount < 0 Then
                MsgBox "No Vendors found!", vbInformation, "PatientTrac Vendor Search"
                Clear_Vendor
            End If
        End If
    Clear_Vendor
    mbooByPass = False
    If cnPat.State = adStateOpen Then
        cnPat.Close
        Set cnPat = Nothing
        Set rsPat = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Search_Vendor:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Search_Vendor", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function



Private Function SearchBy()
    Select Case udtTelephone.CallType
        Case "Vendors"
            Search_Vendor
        Case "Physician"
            Search_Physician
        Case "Patient"
            Search_Patient
    End Select
    Exit Function
End Function

Private Sub cboSearchFilter_Change()
    If Trim(cboSearchFilter.Text) <> vbNullString Then cctSearch.Text = vbNullString
End Sub

Private Sub cboSearchFilter_Click()
    lstListBox.Clear
    lstEncounters.Clear
    cctSearch.Text = vbNullString
    cctSearch.SetFocus
End Sub


Private Sub cboSearchFilter_KeyUp(KeyCode As Integer, Shift As Integer)
    SearchList cboSearchFilter, Trim(cboSearchFilter.Text), True
    Exit Sub
End Sub

Private Sub cctSearch_Change()
    If Trim(cctSearch) <> vbNullString Then
        Timer1.Enabled = False
        Timer1.Enabled = True
    End If
    Exit Sub
End Sub

Private Sub cmdAddList_Click()
Dim i%
Dim iAnswer As Integer

        mstrSelectedTerms = vbNullString
            For i = 0 To lstListBox.listcount - 1
                If lstListBox.Selected(i) Then
                    mstrSelectedTerms = lstListBox.List(i)
                    mlngListItemData = lstListBox.ItemData(i)
                    Exit For
                End If
            Next i
        If mbShowEnctr Then
            If mlngListEnctr = 0 Then
                iAnswer = MsgBox("Do you want to select an Encounter for this function!", vbOKCancel)
                If iAnswer = vbOK Then
                    Exit Sub
                    lstEncounters.SetFocus
                End If
            End If
        End If
        Unload Me
    Exit Sub
End Sub

Private Sub cmdCloseList_Click()
        mstrSelectedTerms = vbNullString
        mlngListItemData = 0
        txtList.Text = vbNullString
        Unload Me

    Exit Sub
End Sub

Private Sub Form_Activate()
Dim i%
    If mFirst Then
       mFirst = False
        If mbShowEnctr Then
           lstEncounters.Enabled = True
           lstEncounters.Visible = True
           lblSelEncounter.Visible = True
        End If
        Screen.MousePointer = vbDefault
        mbooByPass = False
        centerForm Me
    End If
    udtTelephone.CallType = "Patient"
    If cboSearchFilter.listcount <= 0 Then
        If mstrPractice = "Psychiatry" Then
            cboSearchFilter.AddItem "LAST NAME"
            cboSearchFilter.AddItem "EMR ID"
            cboSearchFilter.AddItem "EXTERNAL ID"
            cboSearchFilter.AddItem "SOCIAL SECURITY"
            cboSearchFilter.AddItem "SURGERY ID"
        Else
            cboSearchFilter.AddItem "LAST NAME"
            cboSearchFilter.AddItem "EMR ID"
            cboSearchFilter.AddItem "FINANCIAL ACCT NO"
            cboSearchFilter.AddItem "SOCIAL SECURITY"
            cboSearchFilter.AddItem "SURGERY ID"
        End If
    End If
    If cboSearchFilter.listcount >= 0 Then
        For i = 0 To cboSearchFilter.listcount - 1
            If cboSearchFilter.List(i) = "LAST NAME" Then
                mbooByPass = True
                cboSearchFilter.Text = cboSearchFilter.List(i)
                Exit For
            End If
        Next i
    End If
    Exit Sub
End Sub

Private Sub Form_Load()

    mFirst = True
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    mstrTerms = vbNullString
    frmListPatient.mstrListName = vbNullString
    Unload Me
    Exit Sub
  Exit Sub
End Sub

Private Sub lstEncounters_ItemCheck(Item As Integer)
Dim i%
    If Item < 0 Then
        Exit Sub
    End If
    For i = 0 To lstEncounters.listcount - 1
       If lstEncounters.Selected(i) Then
            If IsNumeric(lstEncounters.ItemData(i)) Then
                mlngListEnctr = lstEncounters.ItemData(i)
                Exit For
            End If
       End If
    Next i
    Exit Sub
End Sub


Private Sub lstListBox_ItemCheck(Item As Integer)
Dim strSQL As String
Dim lngPatientID As Long
Dim cnPat As New ADODB.Connection
Dim rsPat As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim fso As New FileSystemObject

    On Error GoTo Error_Read_ItemCheck
    Screen.MousePointer = vbHourglass
    If Item < 0 Then
        Exit Sub
    End If
    udtPatientName.SurgeryID = 0
    txtList.Text = vbNullString
    If lstListBox.Selected(Item) Then
        lngPatientID = lstListBox.ItemData(Item)
        Get_PatientName lngPatientID
    End If
'    If mstrPractice = "Plastic Surgery" Then
'        If lngPatientID <> 0 Then
'            strSQL = "SELECT SURGERY_ID FROM PATIENT_SURGERY_ID WHERE PATIENT_ID = " & lngPatientID
'            Set rsPat = cnPat.Execute(strSQL)
'            If Not rsPat.EOF Then
'                udtPatientName.SurgeryID = rsPat.Fields("SURGERY_ID")
'            End If
'        End If
'    End If
    If lngPatientID > 0 Then
        If Right(psImageFilePath, 1) <> "\" Then
            psImageFilePath = psImageFilePath & "\"
        End If
        fraPic.ZOrder 0
        If fso.FileExists(psImageFilePath & lngPatientID & "\Photo.jpg") Then
           psPhoto = psImageFilePath & lngPatientID & "\Photo.jpg"
           imgPatient.Picture = LoadPicture(psPhoto)
           lblPhoto.Visible = False
           imgPatient.Visible = True
        Else
          imgPatient.Visible = False
          lblPhoto.Visible = True
        End If
    End If
    txtList.Text = "PATIENT:" & " " & Trim(udtPatientName.Name) _
        & vbNewLine & "DOB:" & " " & Format(udtPatientName.DOB, "MMMM DD, YYYY") & vbNewLine & "SS:" & " " & udtPatientName.SS _
        & vbNewLine & "EXTERNAL ID:" & " " & udtPatientName.ExtID _
        & vbNewLine & "EMR ID:" & " " & udtPatientName.IntId
    If udtPatientName.Status = "INACTIVE" Or udtPatientName.Status = "DECEASED" Then
        txtList.Text = txtList.Text & vbNewLine & "ALERT:" & " " & UCase(udtPatientName.Status)
    End If
    If lngPatientID <> 0 And mbShowEnctr Then Load_Enctr lngPatientID
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Read_ItemCheck:
   Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_ItemCheck", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub Timer1_Timer()
Dim strList As String
    lstListBox.Clear
    strList = cctSearch.Text
    If TextFound(",", cctSearch.Text, Len(","), True) And lstListBox.listcount > 0 Then
            lstListBox.ListIndex = SendMessage(lstListBox.Hwnd, LB_FINDSTRING, -1, _
            ByVal strList)
    Else
        If Len(strList) > 1 Then SearchBy
        imgPatient.Visible = False
    End If
    Timer1.Enabled = False
  Exit Sub
End Sub


