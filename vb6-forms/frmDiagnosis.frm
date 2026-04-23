VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmDiagnosis 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Diagnosis"
   ClientHeight    =   4860
   ClientLeft      =   12840
   ClientTop       =   5445
   ClientWidth     =   9090
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
   Icon            =   "frmDiagnosis.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   9090
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame fraLookup 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   4905
      Left            =   -30
      TabIndex        =   0
      Top             =   30
      Width           =   9075
      Begin PT_XP_Button.PT_XP cmdTerms 
         Height          =   435
         Left            =   7500
         TabIndex        =   13
         Top             =   4170
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   767
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
         Picture         =   "frmDiagnosis.frx":08CA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Set as Default"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDiagnosis.frx":0E64
         ShowFocus       =   -1  'True
      End
      Begin VB.TextBox txtSelectDiag 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   300
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   3030
         Width           =   8415
      End
      Begin VB.ListBox lstData 
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
         Height          =   1740
         ItemData        =   "frmDiagnosis.frx":11B6
         Left            =   300
         List            =   "frmDiagnosis.frx":11B8
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   6
         ToolTipText     =   "Double Click to Add Diagnostic"
         Top             =   840
         Width           =   5415
      End
      Begin VB.TextBox cctSearch 
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
         Height          =   375
         Left            =   300
         TabIndex        =   5
         Top             =   390
         Width           =   5415
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   330
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3030
         Visible         =   0   'False
         Width           =   4875
      End
      Begin VB.TextBox txtPrefName 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   330
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3390
         Visible         =   0   'False
         Width           =   6675
      End
      Begin VB.TextBox txtCode 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5490
         TabIndex        =   2
         Top             =   3030
         Visible         =   0   'False
         Width           =   1515
      End
      Begin PT_XP_Button.PT_XP cmdAddPreferred 
         Height          =   405
         Left            =   3330
         TabIndex        =   7
         Top             =   4290
         Width           =   1125
         _ExtentX        =   1984
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
         Picture         =   "frmDiagnosis.frx":11BA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Add"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDiagnosis.frx":1754
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   405
         Left            =   4620
         TabIndex        =   8
         Top             =   4290
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "frmDiagnosis.frx":1AA6
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDiagnosis.frx":24B8
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSelect 
         Height          =   585
         Left            =   5880
         TabIndex        =   15
         Top             =   2160
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1032
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
         Picture         =   "frmDiagnosis.frx":280A
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Select Diagnosis"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDiagnosis.frx":2DA4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSearchType 
         Height          =   495
         Left            =   5970
         TabIndex        =   16
         ToolTipText     =   "Select Search Type"
         Top             =   120
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   873
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmDiagnosis.frx":30F6
         BackColorDown   =   15133676
         TransparentColor=   14737632
         Caption         =   "LOOK UP"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDiagnosis.frx":3870
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdReset 
         Height          =   495
         Left            =   7530
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         Picture         =   "frmDiagnosis.frx":3F82
         BackColorDown   =   15133676
         TransparentColor=   14737632
         Caption         =   "Reset ICD"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDiagnosis.frx":431C
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblFloat 
         BackColor       =   &H00FFFFFF&
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
         Height          =   615
         Left            =   5850
         TabIndex        =   14
         Top             =   1170
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.Label lblSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   300
         TabIndex        =   12
         Top             =   60
         Width           =   1125
      End
      Begin VB.Label lblLongDesc 
         BackColor       =   &H00FED8D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Long Description"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   360
         TabIndex        =   11
         Top             =   3000
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label lblCode 
         BackColor       =   &H00FED8D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   5520
         TabIndex        =   10
         Top             =   3030
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblPracticeName 
         BackColor       =   &H00FED8D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Preferred Name"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   390
         TabIndex        =   9
         Top             =   3570
         Visible         =   0   'False
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrConnectStringTerms As String
Public mstrSelectedDiag As String
Public mFirst As Boolean
Private mbooByPass As Boolean
Public mbooSelection As Boolean
Public mstrCurrentSel As String
Public mbooVisible As Boolean
Public mstrMultiSelDiag As String
Private mstrType As String
Private mstrDiag As String
Public mbooAdmin As Boolean
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal Hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long

Private Sub Load_PreferredDiag()
Dim strSQL As String
Dim iRows As Long
Dim strText As String
Dim i%
Dim Idx As Integer
Dim strPrefName As String
Dim strCode As String
Dim lngCode As Long
Dim strDesc As String
Dim cnICD As New ADODB.Connection
Dim rsICD As New ADODB.Recordset

    On Error GoTo ERROR_Load_PreferredDiag
    LockWindowUpdate Hwnd
    lstData.Clear
    If cnICD.State = adStateClosed Then cnICD.open psConnect(1)
    
    strText = Trim(cctSearch.Text)
    strSQL = "select distinct ICD_DIAG, ICD_CODE from ICD_TEMPLATE ORDER BY ICD_DIAG"
    Set rsICD = cnICD.Execute(strSQL)
    With rsICD
        Do Until .EOF
            If Not IsNull((!ICD_Code)) Then
                strCode = (!ICD_Code)
            End If
            If Not IsNull((!icd_Diag)) Then strDesc = StrConv((!icd_Diag), vbProperCase)
            If Trim(strDesc) <> vbNullString And Trim(strCode) <> vbNullString Then
                lstData.AddItem Trim(strDesc) & " ~ " & Trim(strCode)
                If IsNumeric(Trim(!ICD_Code)) Then lstData.ItemData(lstData.NewIndex) = Val(strCode)
            End If
        .MoveNext
        Loop
     End With
    LockWindowUpdate 0
    If cnICD.State = adStateOpen Then
        cnICD.Close
        Set cnICD = Nothing
        Set rsICD = Nothing
    End If
    Exit Sub
ERROR_Load_PreferredDiag:
    Write_Error Err.Description, Err.Number, strSQL, "Load_PreferredDiag", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Function Read_Data()
Dim Arr
Dim Arrdiag
Dim Idx As Integer
Dim strDiag As String
Dim cnICD As New ADODB.Connection
Dim rsICD As New ADODB.Recordset
Dim strDate As String
Dim strSQL As String
Dim i%


    On Error GoTo Error_Read_Data
    Screen.MousePointer = vbHourglass
    If cnICD.State = adStateClosed Then cnICD.open psConnect(0)
    strSQL = "SELECT DIAGNOSIS FROM EVALUATION_MANAGEMENT WHERE PATIENT_ID = " & piPatientID & " AND ENCOUNTER_ID = " & piEncounterID
    Set rsICD = cnICD.Execute(strSQL)
    If Not rsICD.EOF Then
        If rsICD.Fields("DIAGNOSIS") <> vbNullString Then
            Arr = Split(rsICD.Fields("DIAGNOSIS"), vbNewLine)
            If UBound(Arr) > 0 Then
                For Idx = LBound(Arr) To UBound(Arr)
                    Arrdiag = Split(Arr(Idx), " - ")
                    If UBound(Arrdiag) > 0 Then
                        strDiag = Arrdiag(0)
                        For i = 0 To lstData.listcount - 1
                            If TextFound(strDiag, lstData.List(i), Len(strDiag), True) Then
                                 lstData.Selected(i) = True
                                 cmdSelect_Click
                            End If
                        Next i
                    End If
                Next Idx
            Else
                Arrdiag = Split(rsICD.Fields("DIAGNOSIS"), " - ")
                If UBound(Arrdiag) > 0 Then
                    strDiag = Arrdiag(0)
                    For Idx = 0 To lstData.listcount - 1
                        If TextFound(strDiag, lstData.List(Idx), Len(strDiag), True) Then
                             lstData.Selected(Idx) = True
                             cmdSelect_Click
                        End If
                    Next Idx
                End If
            End If
        End If
    End If
    If cnICD.State = adStateOpen Then
        cnICD.Close
        Set cnICD = Nothing
        Set rsICD = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Data", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Get_ICD(vstrType As String, vstrDiag As String)
Dim strSQL As String
Dim cnICD As New ADODB.Connection
Dim rsICD As New ADODB.Recordset
Dim strPrefName As String
Dim strCode As String
Dim lngCode As Long
Dim strDesc As String
Dim Arr

    On Error GoTo ERROR_Get_ICD
    Screen.MousePointer = vbHourglass
    LockWindowUpdate Hwnd
    lstData.Clear
    If cnICD.State = adStateClosed Then cnICD.open psConnect(1)
    If Trim(vstrDiag) = vbNullString Then
        vstrDiag = Trim(cctSearch.Text)
    End If
    If IsNumeric(vstrDiag) Then
        mstrType = "LookUp Diagnosis by Number"
    Else
        mstrType = "LookUp Diagnosis by Name"
    End If
    If Trim(vstrDiag) <> vbNullString Then
        If cnICD.State = adStateClosed Then cnICD.open psConnect(1)
        If mstrType = "LookUp Diagnosis by Name" Then
            strSQL = "select long_desc, code from icd_code Where long_desc like '%" & vstrDiag & "%'"
        ElseIf mstrType = "LookUp Diagnosis by Number" Then
            strSQL = "select long_desc, code from icd_code Where code like '%" & vstrDiag & "%'"
        End If
        If strSQL <> vbNullString Then
            Set rsICD = cnICD.Execute(strSQL)
             With rsICD
                 Do Until .EOF
                     If Not IsNull((!code)) Then strCode = IsProperCase((!code))
                     If Not IsNull((!Long_Desc)) Then strDesc = IsProperCase((!Long_Desc))
                     If Trim(strCode) <> vbNullString And strDesc <> vbNullString Then
                        If mstrType = "LookUp Diagnosis by Name" Then
                            lstData.AddItem Trim(strDesc) & " - " & Trim(strCode)
                            If IsNumeric(Trim(strCode)) Then lstData.ItemData(lstData.NewIndex) = Val(strCode)
                        Else
                            lstData.AddItem Trim(strCode) & " - " & Trim(strDesc)
                            If IsNumeric(Trim(strCode)) Then lstData.ItemData(lstData.NewIndex) = Val(strCode)
                         End If
                     End If
                 .MoveNext
                 Loop
              End With
        End If
    Else
        MsgBox "Please enter Diagnosis in the Search field first!", vbInformation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If cnICD.State = adStateOpen Then
        cnICD.Close
        Set cnICD = Nothing
        Set rsICD = Nothing
    End If
     LockWindowUpdate 0
    Screen.MousePointer = vbDefault
    Exit Function
ERROR_Get_ICD:
    Write_Error Err.Description, Err.Number, strSQL, "Get_ICD", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub Save_Diagnosis()
Dim strSQL As String
Dim cnICD As New ADODB.Connection
Dim rsICD As New ADODB.Recordset
Dim strDate As String
Dim Arr
Dim ArrCode
Dim lngPrimKey As Long
Dim lngAcctID As Long
Dim Idx As Integer


    On Error GoTo Error_Save_Diagnosis
    Screen.MousePointer = vbHourglass
    If cnICD.State = adStateClosed Then cnICD.open psConnect(0)
    If clsEncounter.StartDate <> vbNullString Then
        strDate = Format(clsEncounter.StartDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    Else
        strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    End If
    strSQL = "SELECT ACCT_ID FROM PATIENT_ACCOUNTS WHERE " _
        & "PATIENT_ID = " & piPatientID & " AND PATIENT_TYPE = 'Patient'"
    Set rsICD = cnICD.Execute(strSQL)
    With rsICD
        If Not .EOF Then
            lngAcctID = (!Acct_ID)
        End If
    End With
    clsOpNote.Preop_Diagnosis = vbNullString
    clsOpNote.Preop_Code = vbNullString
     If txtSelectDiag.Text <> vbNullString Then
        Arr = Split(txtSelectDiag.Text, vbNewLine)
        If UBound(Arr) > 0 Then
            For Idx = LBound(Arr) To UBound(Arr)
                ArrCode = Split(Arr(Idx), "~")
                If UBound(ArrCode) > 0 Then
                    If clsOpNote.Preop_Diagnosis = vbNullString Then
                        clsOpNote.Preop_Diagnosis = ArrCode(0)
                    Else
                        clsOpNote.Preop_Diagnosis = clsOpNote.Preop_Diagnosis & vbNewLine & ArrCode(0)
                    End If
                    If clsOpNote.Preop_Code = vbNullString Then
                        clsOpNote.Preop_Code = ArrCode(1)
                    Else
                        clsOpNote.Preop_Code = clsOpNote.Preop_Code & vbNewLine & ArrCode(1)
                    End If
                    If clsReport.Diagnosis = vbNullString Then
                        clsReport.Diagnosis = Replace(Arr(Idx), "~", "-")
                    Else
                        clsReport.Diagnosis = clsReport.Diagnosis & vbNewLine & Replace(Arr(Idx), "~", "-")
                    End If
                End If
            Next Idx
        Else
            ArrCode = Split(Trim(txtSelectDiag.Text), "~")
            If UBound(ArrCode) > 0 Then
                clsOpNote.Preop_Diagnosis = ArrCode(0)
                clsOpNote.Preop_Code = ArrCode(1)
                clsReport.Diagnosis = Replace(Trim(txtSelectDiag.Text), "~", "-")
            End If
        End If
    End If
    If lngAcctID > 0 Then
        strSQL = "INSERT INTO PATIENT_ACCOUNTS (PATIENT_TYPE, PATIENT_ID, ACCOUNT_NO, ICD_CODE) " _
            & "VALUES ('Patient', " & piPatientID & "," & piPatientID & ",'" & clsOpNote.Preop_Code & "')"
        Set rsICD = cnICD.Execute(strSQL)
    Else
         strSQL = "UPDATE PATIENT_ACCOUNTS  SET PATIENT_ID = " & piPatientID & "," _
            & "ICD_CODE  = '" & clsOpNote.Preop_Code & "' WHERE ACCT_ID = " & lngAcctID
        Set rsICD = cnICD.Execute(strSQL)
    End If

    If clsReport.Diagnosis <> vbNullString Then
        strSQL = "SELECT * FROM EVALUATION_MANAGEMENT WHERE PATIENT_ID = " & piPatientID & " AND ENCOUNTER_ID = " & piEncounterID
         Set rsICD = cnICD.Execute(strSQL)
         If Not rsICD.EOF Then
            lngPrimKey = rsICD.Fields("EM_ID")
         End If
        If lngPrimKey > 0 Then
            strSQL = "UPDATE EVALUATION_MANAGEMENT SET DIAGNOSIS = '" & clsReport.Diagnosis & "', UPDATE_DATE = '" & strDate & "' WHERE EM_ID = " & lngPrimKey
            Set rsICD = cnICD.Execute(strSQL)
        Else
            strSQL = "INSERT INTO EVALUATION_MANAGEMENT (PATIENT_ID,ENCOUNTER_ID, DIAGNOSIS, INSERT_DATE) values " _
                & "(" & modSpecific.piPatientID & ", " & piEncounterID & ",'" & clsReport.Diagnosis & "','" & strDate & "')"
            Set rsICD = cnICD.Execute(strSQL)
        End If
    End If
    If cnICD.State = adStateOpen Then
        cnICD.Close
        Set cnICD = Nothing
        Set rsICD = Nothing
    End If
    clsOpNote.Preop_Diagnosis = vbNullString
    clsOpNote.Preop_Code = vbNullString

    Screen.MousePointer = vbDefault
    Exit Sub
Error_Save_Diagnosis:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Save_Diagnosis", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cctSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    If mbooByPass Then mbooByPass = False: Exit Sub
    If Len(cctSearch) > 2 Then
        strList = Trim(cctSearch.Text)
    End If
    If lstData.listcount > 0 Then
        lstData.ListIndex = SendMessage(lstData.Hwnd, LB_FINDSTRING, -1, _
        ByVal strList)
    End If
    Exit Sub
End Sub

Private Sub cmdAddPreferred_Click()
    Set_Diagnosis
    Save_Diagnosis
    Clear_Fields
    Unload Me
End Sub

Private Sub cmdClose_Click()
    mstrMultiSelDiag = Trim(mstrMultiSelDiag)
    Unload Me
End Sub



Private Sub Set_Diagnosis()
Dim strCode As String
Dim strLongDesc As String
Dim strPrefName As String
Dim lngCnt As Long
Dim Arr
Dim i%

    lngCnt = 0
    If txtSelectDiag.Text <> vbNullString Then
        Arr = Split(txtSelectDiag.Text, vbNewLine)
        If UBound(Arr) > 0 Then
            For i = LBound(Arr) To UBound(Arr)
                lngCnt = lngCnt + 1
                If mstrSelectedDiag = vbNullString Then
                    mstrSelectedDiag = lngCnt & ".  " & Trim(Arr(i))
                Else
                    mstrSelectedDiag = mstrSelectedDiag & vbNewLine & lngCnt & ".  " & Trim(Arr(i))
                End If
            Next i
        Else
            mstrSelectedDiag = Trim(txtSelectDiag.Text)
            mstrSelectedDiag = Replace(mstrSelectedDiag, "~", " - ")
        End If
    End If
    mstrMultiSelDiag = Trim(txtSelectDiag.Text)
    mstrMultiSelDiag = Replace(mstrMultiSelDiag, "~", " - ")
    If mstrSelectedDiag = vbNullString Then mstrSelectedDiag = Trim(mstrMultiSelDiag)
    Exit Sub
End Sub




Private Sub cmdReset_Click()
    Load_PreferredDiag
End Sub

Private Sub cmdSearchType_Click()
Dim strList As String
    mbooByPass = True
    If Len(cctSearch) > 2 Then
        strList = Trim(cctSearch.Text)
    End If
    If Trim(strList) <> vbNullString Then Get_ICD "", strList
    mbooByPass = False
    Exit Sub
End Sub


Private Sub cmdSelect_Click()
Dim Arr
Dim i%
Dim booSkip As Boolean
Dim strDiag As String

    If mbooByPass Then mbooByPass = False: Exit Sub
    booSkip = False
    For i = 0 To lstData.listcount - 1
        If lstData.Selected(i) Then
           strDiag = Trim(lstData.List(i))
           mbooByPass = True
           lstData.Selected(i) = False
           mbooByPass = False
           Exit For
        End If
    Next i

    Arr = Split(txtSelectDiag.Text, "vbNewLine")
    If UBound(Arr) > 0 Then
        For i = LBound(Arr) To UBound(Arr)
            If Arr(i) <> vbNullString Then
                If TextFound(Trim(Arr(i)), strDiag, Len(Arr(i)), True) Then
                    booSkip = True
                    Exit For
                End If
            End If
        Next i
    End If
    If Not booSkip Then
        If Trim(txtSelectDiag.Text) = vbNullString Then
            txtSelectDiag.Text = Trim(strDiag)
        Else
            txtSelectDiag.Text = Trim(txtSelectDiag.Text) & vbNewLine & Trim(strDiag)
        End If
    End If
    cmdAddPreferred.Enabled = True
    cctSearch.Text = vbNullString
    lblFloat.Caption = vbNullString
    lblFloat.Visible = False
    Exit Sub
End Sub

Private Sub cmdTerms_Click()
Dim cnWrite As New ADODB.Connection
Dim rsWrite As New ADODB.Recordset
Dim Arr
Dim strDiag As String
Dim strNo As String
Dim strSQL As String


    On Error GoTo Error:
    If Trim(txtSelectDiag.Text) <> vbNullString Then
        strDiag = Trim(txtSelectDiag.Text)
    End If
    If strDiag <> vbNullString Then
        If cnWrite.State = adStateClosed Then cnWrite.open psConnect(1)
        strSQL = "UPDATE SURG_PROCEDURE_CODES SET  PREOP_DIAGNOSIS = '" & strDiag & "', POSTOP_DIAGNOSIS = '" & strDiag & "'  WHERE PROCEDURE_ID = " & clsOpNote.Procedure_ID
            Set rsWrite = cnWrite.Execute(strSQL)
        End If
        If cnWrite.State = adStateOpen Then
            cnWrite.Close
            Set cnWrite = Nothing
            Set rsWrite = Nothing
        End If
    Exit Sub
Error:
    MsgBox Err.Description, vbInformation
    Exit Sub
    Resume
End Sub

Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
        If mstrType = vbNullString Then mstrType = "LookUp Diagnosis by Name"
      If psConnect(1) = vbNullString Then psConnect(1) = psConnect(1)
    End If
    Load_PreferredDiag
    If lstData.listcount > 0 And mstrCurrentSel <> vbNullString Then
        lstData.ListIndex = SendMessage(lstData.Hwnd, LB_FINDSTRING, -1, mstrCurrentSel)
    End If
    cmdTerms.Visible = True
    Read_Data
    centerForm Me
    Exit Sub
End Sub

Private Sub Form_Load()
    mFirst = True
    If psConnect(1) = vbNullString Then psConnect(1) = psConnect(1)
    Load_PreferredDiag
    If lstData.listcount > 0 And mstrCurrentSel <> vbNullString Then
        lstData.ListIndex = SendMessage(lstData.Hwnd, LB_FINDSTRING, -1, mstrCurrentSel)
    End If
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next
    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.objects
      Set obj = Nothing
    Next obj
    mstrType = vbNullString
    Exit Sub
End Sub

Private Sub Clear_Fields()
    txtDesc.Text = vbNullString
    txtPrefName.Text = vbNullString
    txtCode.Text = vbNullString
    cctSearch.Text = vbNullString
    mbooByPass = True
    lstData.Clear
End Sub

Private Sub Save_ICD()
Dim i%
Dim strSQL As String
Dim lngrecs As Long
Dim varRS
Dim booRet As Boolean
Dim booinsert As Boolean
Dim strCode As String
Dim strLongDesc As String
Dim strPrefName As String
Dim iAnswer As Integer
Dim lngICD_ID As Long

    Screen.MousePointer = vbHourglass
    strCode = txtCode.Text
    strLongDesc = txtDesc.Text
    strPrefName = txtPrefName.Text
    On Error GoTo ERROR_Save_ICD
    Open_DB psConnect(1)
    'Insert or Update?......................
    strSQL = "Select ICD_ID From ICD_TEMPLATE  WHERE ICD_CODE  = '" & strCode & "' "
    Set rs = cn.Execute(strSQL)
    If Not rs.EOF Then
        lngICD_ID = rs.Fields("ICD_ID")
        booinsert = False
        iAnswer = MsgBox("Preferred ICD Exists Overwrite with new Preferred Name?", vbYesNo + vbQuestion)
        If iAnswer = vbYes Then
            booinsert = False
        Else
            Exit Sub
        End If
    Else
        booinsert = True
    End If
    If booinsert Then
        strSQL = "Insert into ICD_TEMPLATE (icd_Code, icd_Diag, preferred_name) " _
                & "values ('" & strCode & "','" & strLongDesc & "','" & strPrefName & "')"
    Else
        strSQL = "Update ICD_TEMPLATE set icd_Code = '" & strCode & "', " _
                & "icd_Diag = '" & strLongDesc & "', " _
                & "preferred_name = '" & strPrefName & "', " _
                & "update_Date =  '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "'" _
                & "Where ICD_ID = " & lngICD_ID
    End If
    Set rs = cn.Execute(strSQL)
    Close_DB psConnect(1)
    MsgBox strLongDesc & "Successfully Added to Preferred ICD-9", vbInformation
    Exit Sub
ERROR_Save_ICD:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Save_ICD", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub lstData_ItemCheck(Item As Integer)
Dim Arr
    If mbooByPass Then mbooByPass = False: Exit Sub
    mstrDiag = vbNullString
    lblFloat.Caption = vbNullString
    lblFloat.Visible = False
    If Trim(txtSelectDiag.Text) = vbNullString Then
        txtSelectDiag.Text = lstData.List(Item)
    Else
        txtSelectDiag.Text = txtSelectDiag.Text & vbNewLine & lstData.List(Item)
    End If
    
'    If lstData.List(Item) <> vbNullString Then
'      Arr = Split(lstData.List(Item), "~")
'      If UBound(Arr) > 0 Then
'        If Arr(1) <> vbNullString Then Get_ICD "LookUp Diagnosis by Number", CStr(Arr(1))
'        If mstrDiag <> vbNullString Then
'            lblFloat.Visible = True
'            lblFloat.Caption = mstrDiag
'        End If
'        txtSelectDiag.Text = lstData.List(Item)
'      End If
'    End If
    Exit Sub
End Sub

Private Sub lstData_LostFocus()
    lblFloat.Visible = False
End Sub









