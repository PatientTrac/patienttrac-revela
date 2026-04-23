VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmDSMIV 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "DSM-IV Look-Up"
   ClientHeight    =   4935
   ClientLeft      =   8265
   ClientTop       =   6510
   ClientWidth     =   10185
   Icon            =   "frmDSMIV.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   10185
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame fraLookup 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
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
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10155
      Begin VB.ComboBox cboChapter 
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
         Left            =   6090
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   90
         Width           =   3855
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
         ForeColor       =   &H00000000&
         Height          =   1515
         Left            =   300
         TabIndex        =   11
         Top             =   2730
         Width           =   9645
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
         TabIndex        =   9
         Top             =   3030
         Visible         =   0   'False
         Width           =   1515
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
         TabIndex        =   5
         Top             =   3780
         Visible         =   0   'False
         Width           =   6675
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
         TabIndex        =   2
         Top             =   600
         Width           =   6705
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
         Height          =   1500
         ItemData        =   "frmDSMIV.frx":08CA
         Left            =   300
         List            =   "frmDSMIV.frx":08CC
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   1110
         Width           =   9615
      End
      Begin PT_XP_Button.PT_XP cmdSearchBy 
         Height          =   345
         Left            =   7080
         TabIndex        =   10
         Top             =   600
         Width           =   1635
         _ExtentX        =   2884
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
         Picture         =   "frmDSMIV.frx":08CE
         AppearanceMode  =   1
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Dx by Name"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDSMIV.frx":0E68
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAddPreferred 
         Height          =   375
         Left            =   3645
         TabIndex        =   12
         Top             =   4350
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
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
         Picture         =   "frmDSMIV.frx":11BA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Add and Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDSMIV.frx":1754
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSaveRefresh 
         Height          =   375
         Left            =   10590
         TabIndex        =   13
         Top             =   4470
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
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
         Picture         =   "frmDSMIV.frx":1AA6
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Add and New"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDSMIV.frx":2040
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   375
         Left            =   5595
         TabIndex        =   14
         Top             =   4350
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
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
         Picture         =   "frmDSMIV.frx":2392
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmDSMIV.frx":2DA4
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblSortBy 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Chapter"
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
         Height          =   225
         Left            =   5190
         TabIndex        =   16
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label lblPracticeName 
         BackColor       =   &H00FED8D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Preferred Name"
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
         Height          =   255
         Left            =   390
         TabIndex        =   8
         Top             =   3570
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblCode 
         BackColor       =   &H00FED8D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Height          =   225
         Left            =   5520
         TabIndex        =   7
         Top             =   2820
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblLongDesc 
         BackColor       =   &H00FED8D9&
         BackStyle       =   0  'Transparent
         Caption         =   "Long Description"
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
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   2790
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label lblSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search"
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
         Height          =   255
         Left            =   330
         TabIndex        =   3
         Top             =   300
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmDSMIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrConnectStringTerms As String
Public mstrSelectedDiag As String
Public mFirst As Boolean
Public mbooSelection As Boolean
Public mbooVisible As Boolean
Private mbooByPass As Boolean
Public mstrAxis As String
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long

Private Function Field_Visible()

    If mbooVisible Then
        lblLongDesc.Visible = True
        lblCode.Visible = True
        lblPracticeName.Visible = True
        txtDesc.Visible = True
        txtPrefName.Visible = True
        txtCode.Visible = True
    End If
End Function

Private Sub Set_Diagnosis()
Dim strCode As String
Dim strLongDesc As String
Dim strPrefName As String

    If mstrSelectedDiag = vbNullString Then
        mstrSelectedDiag = txtSelectDiag.Text
    ElseIf Not TextFound(txtSelectDiag.Text, mstrSelectedDiag, Len(txtSelectDiag.Text), True) Then
        mstrSelectedDiag = mstrSelectedDiag & "; " & txtSelectDiag.Text
    End If
    Exit Sub
End Sub

Private Sub cboChapter_Click()
Dim i%
Dim strSQL As String
Dim strChapter As String
Dim strName As String
Dim rsDiag As New ADODB.Recordset
Dim cnDiag As New ADODB.Connection

    On Error GoTo Error_cboChapter_KeyUp
    Screen.MousePointer = vbHourglass
    If cnDiag.State = adStateClosed Then cnDiag.open psConnect(1)
    If Trim(cboChapter.Text) <> vbNullString Then
        lstData.Clear
        lstData.ListIndex = -1
        strChapter = Trim(cboChapter.Text)
        strSQL = "Select Name, Number, Chapter from  [DSM-IV] WHERE CHAPTER = '" & strChapter & "' ORDER BY Name"
        Set rsDiag = cnDiag.Execute(strSQL)
        With rsDiag
            Do Until .EOF
                If Trim((!Name)) <> vbNullString And (!Number) <> vbNullString Then
                    strName = Replace((!Name), "''", "'")
                    lstData.AddItem strName & " - " & Format((!Number), "000.00")
                End If
            .MoveNext
            Loop
        End With
    End If
    If cnDiag.State = adStateOpen Then
        cnDiag.Close
        Set cnDiag = Nothing
        Set rsDiag = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cboChapter_KeyUp:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "cboChapter_KeyUp", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub cctSearch_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strList As String
    If Len(cctSearch.Text) > 2 Then Load_PsySystems cctSearch.Text
    strList = Trim(cctSearch.Text)
    Exit Sub
End Sub

Private Sub cmdAddPreferred_Click()
    Set_Diagnosis
    Clear_Fields
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSaveRefresh_Click()
    
    Set_Diagnosis
    Clear_Fields
End Sub




Private Function Load_PsySystems(vstrType As String)
Dim strSQL As String
Dim strName As String
Dim strType As String
Dim rsDiag As New ADODB.Recordset
Dim cnDiag As New ADODB.Connection


    LockWindowUpdate hwnd
    On Error GoTo Error_Load_PsySystems
    Screen.MousePointer = vbHourglass
    
    Open_DB mstrConnectStringTerms
    Me.Caption = "DSM-IV" & " " & mstrAxis
    If mstrAxis = "AXIS I" Then strType = "I"
    If mstrAxis = "AXIS II" Then strType = "II"
    lstData.Clear
    lstData.ListIndex = -1
    If cnDiag.State = adStateClosed Then cnDiag.open psConnect(1)
    If vstrType <> vbNullString And IsNumeric(vstrType) Then
        strSQL = "select Name, Number, Chapter from  [DSM-IV] Where Number like '%" & vstrType & "%'"
        Set rsDiag = cnDiag.Execute(strSQL)
        With rsDiag
            Do Until .EOF
                If Trim((!Name)) <> vbNullString And (!Number) <> vbNullString Then
                    strName = Replace((!Name), "''", "'")
                    If Trim(vstrType) Like "[A-Za-z]###" Or Trim(vstrType) Like "###" Then
                        If strName <> vbNullString Then lstData.AddItem Format((!Number), "000.00") & " - " & strName
                    Else
                        lstData.AddItem strName & " - " & Format((!Number), "000.00")
                    End If
                End If
            .MoveNext
            Loop
        End With
    ElseIf vstrType <> vbNullString Then
        strSQL = "select Name, Number, Chapter from  [DSM-IV]  Where Name like '%" & vstrType & "%'"
        Set rsDiag = cnDiag.Execute(strSQL)
        With rsDiag
            Do Until .EOF
                If Trim((!Name)) <> vbNullString And (!Number) <> vbNullString Then
                    strName = Replace((!Name), "''", "'")
                    If Trim(vstrType) Like "[A-Za-z]###" Or Trim(vstrType) Like "###" Then
                        If strName <> vbNullString Then lstData.AddItem Format((!Number), "000.00") & " - " & strName
                    Else
                        lstData.AddItem strName & " - " & Format((!Number), "000.00")
                    End If
                End If
            .MoveNext
            Loop
        End With
    End If

    cboChapter.Clear
    strSQL = "Select distinct Chapter from  [DSM-IV]"
    Set rsDiag = cnDiag.Execute(strSQL)
        With rsDiag
            Do Until .EOF
                If Not IsNull(!Chapter) And Trim(!Chapter) <> vbNullString Then cboChapter.AddItem Trim((!Chapter))
            .MoveNext
            Loop
        End With
    If cnDiag.State = adStateOpen Then
        cnDiag.Close
        Set cnDiag = Nothing
        Set rsDiag = Nothing
    End If
    LockWindowUpdate 0
    If lstData.ListCount <= 0 Then
        LookUp_ICD vstrType
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_PsySystems:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_PsySystems", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function LookUp_ICD(vstrType As String)
Dim strSQL As String
Dim cnICD As New ADODB.Connection
Dim rsICD As New ADODB.Recordset
Dim strText As String
Dim i%
Dim idx As Integer
Dim strPrefName As String
Dim strCode As String
Dim lngCode As Long
Dim strDesc As String

    On Error GoTo ERROR_Look_UpDiag
    Screen.MousePointer = vbHourglass
    If vstrType = vbNullString Then Exit Function
    LockWindowUpdate hwnd
    lstData.Clear
    If vstrType = vbNullString Then vstrType = Trim(cctSearch.Text)
    If Trim(vstrType) <> vbNullString Then
        If cnICD.State = adStateClosed Then cnICD.open mstrConnectStringTerms
            If Trim(vstrType) Like "[A-Za-z]###" Or IsNumeric(vstrType) Then
                strSQL = "select long_desc, code from icd_code Where code like '" & vstrType & "%'"
            Else
                strSQL = "select long_desc, code from icd_code Where long_desc like '" & vstrType & "%'"
            End If
            If strSQL <> vbNullString Then
                Set rsICD = cnICD.Execute(strSQL)
                 With rsICD
                     Do Until .EOF
                         If Not IsNull((!code)) Then strCode = IsProperCase((!code))
                         If Not IsNull((!Long_Desc)) Then strDesc = IsProperCase((!Long_Desc))
                         If Trim(strCode) <> vbNullString And strDesc <> vbNullString Then
                            strCode = Format(strCode, "000.00")
                            If Not IsNumeric(vstrType) Then
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
    End If
    If cnICD.State = adStateOpen Then
        cnICD.Close
        Set cnICD = Nothing
        Set rsICD = Nothing
    End If
     LockWindowUpdate 0
    Screen.MousePointer = vbDefault
    Exit Function
ERROR_Look_UpDiag:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Look_UpDiag", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
        If mstrConnectStringTerms = vbNullString Then mstrConnectStringTerms = psConnect(1)
        mstrSelectedDiag = vbNullString
        Clear_Fields
        Load_PsySystems vbNullString
        cctSearch.SetFocus
        centerForm Me
    End If
    Exit Sub
End Sub

Private Sub Form_Load()
    mFirst = True

    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
  Exit Sub

End Sub

Private Sub Clear_Fields()
    txtDesc.Text = vbNullString
    txtPrefName.Text = vbNullString
    txtCode.Text = vbNullString
    cctSearch.Text = vbNullString
    lstData.Clear
    lstData.ListIndex = -1
End Sub

Private Sub Save_Diagnosis()
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

    strCode = txtCode.Text
    strLongDesc = txtDesc.Text
    strPrefName = txtPrefName.Text

    Open_DB mstrConnectStringTerms
    
    'Insert or Update?......................
    strSQL = "Select NAME From [DSM-IV]  " _
        & "Where CHAPTER  = '" & strCode & "' "
    Set rs = cn.Execute(strSQL)
    
    With rs
        If Not .EOF Then
            If IsNull(!Name) Then booinsert = True
            iAnswer = MsgBox("Preferred DSM-IV Exists Overwrite with new Preferred Name?", vbYesNo + vbQuestion)
            If iAnswer = vbYes Then
                booinsert = False
            Else
                Exit Sub
            End If
        End If
    End With
        If booinsert Then
            strSQL = "Insert into [DSM-IV]  (Number, NAME, Preferred_Name) " _
                    & "values ('" & strCode & "','" & strLongDesc & "','" & strPrefName & "' )"
                    
        Else
            strSQL = "Update [DSM-IV] set NAME = '" & strLongDesc & "', " _
                    & "preferred_name = '" & strPrefName & "', " _
                    & "update_Date =  '" & Format(Now, "mm/dd/yyyy hh:nn:ss") & "'" _
                    & "Where Number = '" & strCode & "'"
        End If

        Set rs = cn.Execute(strSQL)
        mbooSelection = True
        MsgBox strLongDesc & "Successfully Added to DSM-IV", vbInformation
        Close_DB mstrConnectStringTerms
        If mbooSelection = True Then
            If mstrSelectedDiag <> vbNullString Then
                If strPrefName = vbNullString Then strPrefName = strLongDesc
                mstrSelectedDiag = strPrefName & " - " & strCode
            Else
                If strPrefName = vbNullString Then strPrefName = strLongDesc
                mstrSelectedDiag = mstrSelectedDiag & "~" & strPrefName & " - " & strCode
            End If
        End If
    Exit Sub
End Sub

Private Sub lstData_ItemCheck(Item As Integer)
Dim i%
Dim strSQL As String
Dim lngrecs As Long
Dim varRS
Dim booRet As Boolean
Dim strDSM As String
Dim Arr
    If mbooByPass Then mbooByPass = False: Exit Sub
    If lstData.Selected(Item) Then
        If txtSelectDiag.Text = vbNullString Then
            txtSelectDiag.Text = lstData.List(Item)
        Else
            txtSelectDiag.Text = Trim(txtSelectDiag.Text) & " ; " & lstData.List(Item)
        End If
    End If
    cctSearch.Text = vbNullString
    mbooByPass = True
    lstData.Clear
    cctSearch.SetFocus
    mbooByPass = False
    Exit Sub
End Sub


