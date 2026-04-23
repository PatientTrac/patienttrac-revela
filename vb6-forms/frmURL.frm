VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmURL 
   Caption         =   "PatientTrac Browser Terms"
   ClientHeight    =   4320
   ClientLeft      =   5340
   ClientTop       =   2205
   ClientWidth     =   6585
   Icon            =   "frmURL.frx":0000
   ScaleHeight     =   4320
   ScaleWidth      =   6585
   Begin VB.Frame fraPatExplorer 
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
      Height          =   3705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      Begin VB.ListBox lstTerms 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   90
         TabIndex        =   4
         Top             =   1860
         Width           =   4605
      End
      Begin VB.TextBox txtValue 
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
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   4605
      End
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Text            =   "¯¯END!"
         Top             =   360
         Width           =   4605
      End
      Begin PT_XP_Button.PT_XP cmdHome 
         Height          =   645
         Left            =   4920
         TabIndex        =   1
         Top             =   330
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1138
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmURL.frx":08CA
         BackColorDown   =   15133676
         TransparentColor=   13160660
         Caption         =   "Set Home"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmURL.frx":11A4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAdd 
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   1410
         Width           =   885
         _ExtentX        =   1561
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
         Picture         =   "frmURL.frx":1DF6
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "Add"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmURL.frx":2390
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdUp 
         Height          =   405
         Left            =   5130
         TabIndex        =   6
         Top             =   2280
         Width           =   375
         _ExtentX        =   661
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
         Picture         =   "frmURL.frx":26E2
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmURL.frx":2E5C
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdDown 
         Height          =   405
         Left            =   5130
         TabIndex        =   7
         Top             =   2760
         Width           =   375
         _ExtentX        =   661
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
         Picture         =   "frmURL.frx":356E
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmURL.frx":3CE8
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdReplace 
         Height          =   345
         Left            =   1050
         TabIndex        =   8
         Top             =   1410
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "frmURL.frx":43FA
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "Replace"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmURL.frx":4994
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRemove 
         Height          =   345
         Left            =   2280
         TabIndex        =   9
         Top             =   1410
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "frmURL.frx":4CE6
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "Remove"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmURL.frx":5280
         ShowFocus       =   -1  'True
      End
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   345
      Left            =   2310
      TabIndex        =   10
      Top             =   3870
      Width           =   915
      _ExtentX        =   1614
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
      Picture         =   "frmURL.frx":55D2
      AppearanceThemes=   3
      TransparentColor=   13160660
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmURL.frx":596C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDone 
      Height          =   345
      Left            =   3360
      TabIndex        =   11
      Top             =   3870
      Width           =   915
      _ExtentX        =   1614
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
      Picture         =   "frmURL.frx":5CBE
      AppearanceThemes=   3
      TransparentColor=   13160660
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmURL.frx":66D0
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Public Sub Reorder_ListBox(lst As Control, bUp As Boolean)
Dim i As Integer
Dim sContent As String
Dim lData As Long
Dim X As Integer

    If bUp Then
        X = 1
    Else
        X = -1
    End If

    i = lst.ListIndex
    sContent = lst.Text
    lData = lst.ItemData(i)
   
   If (i > 0 And bUp) Or (i < lst.listcount - 1 And Not bUp) Then
    lst.RemoveItem i
    lst.AddItem sContent, i - X
    lst.ItemData(i - X) = lData
    lst.Selected(i - X) = True
   End If
End Sub
Private Sub SaveList()
Dim Idx As Integer
Dim lngSeqNo As Long
Dim strAddress As String
Dim cnHome As New ADODB.Connection
Dim rsHome As New ADODB.Recordset
Dim strSQL As String
Dim lngID As Long

    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    
    If cnHome.State = adStateClosed Then cnHome.open psConnect(1)
    For Idx = 0 To lstTerms.listcount - 1
        strAddress = lstTerms.List(Idx)
        lngSeqNo = lstTerms.ItemData(Idx)
        strSQL = "SELECT * FROM URL_TERMS WHERE URL = '" & strAddress & "'"
        Set rsHome = cnHome.Execute(strSQL)
        With rsHome
            If Not .EOF Then
                If (!URL_ID) > 0 Then
                    lngID = (!URL_ID)
                End If
            End If
        End With
        If lngID = 0 Then
            strSQL = "INSERT INTO URL_TERMS (SEQUENCE_NO, URL)" _
                & "values(" & lngSeqNo & ",'" & strAddress & "')"
            Set rsHome = cnHome.Execute(strSQL)
        Else
            strSQL = "UPDATE URL_TERMS SET URL = '" & strAddress & "', " _
                & "SEQUENCE_NO = " & lngSeqNo & " WHERE URL_ID = " & lngID
            Set rsHome = cnHome.Execute(strSQL)
        End If
    Next Idx
    If cnHome.State = adStateOpen Then
        cnHome.Close
        Set cnHome = Nothing
        Set rsHome = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbInformation
    Exit Sub
End Sub
Public Sub LoadListBox()
Dim cnHome As New ADODB.Connection
Dim rsHome As New ADODB.Recordset
Dim strSQL As String
Dim lngID As Long
Dim strAddress As String
 
    If cnHome.State = adStateClosed Then cnHome.open psConnect(1)
    strSQL = "SELECT * FROM URL_TERMS ORDER BY SEQUENCE_NO"
    Set rsHome = cnHome.Execute(strSQL)
    With rsHome
        Do Until .EOF
            lstTerms.AddItem (!URL)
            If (!SEQUENCE_NO) <> vbNullString Then lstTerms.ItemData(lstTerms.NewIndex) = (!SEQUENCE_NO)
            .MoveNext
        Loop
        lstTerms.ListIndex = -1
    End With
    Screen.MousePointer = vbDefault
    If cnHome.State = adStateOpen Then
        cnHome.Close
        Set cnHome = Nothing
        Set rsHome = Nothing
    End If
Exit Sub
End Sub
Private Function Set_Home()
Dim cnHome As New ADODB.Connection
Dim rsHome As New ADODB.Recordset
Dim strSQL As String
Dim lngID As Long
Dim booinsert As Boolean
Dim strAddress As String


    On Error GoTo Error_Set_Home
    Screen.MousePointer = vbHourglass
    booinsert = True
    If cnHome.State = adStateClosed Then cnHome.open psConnect(1)
    strAddress = cboAddress.Text
    strSQL = "SELECT HOME, URL_ID, URL from URL_TERMS "
    Set rsHome = cnHome.Execute(strSQL)
    Do Until rsHome.EOF
        If rsHome.Fields("HOME") <> vbNullString Then
            lngID = rsHome.Fields("URL_ID")
            strAddress = rsHome.Fields("URL")
            Exit Do
        End If
        rsHome.MoveNext
        booinsert = False
    Loop
    
    If lngID = 0 Then
        strSQL = "INSERT INTO URL_TERMS (HOME, URL)" _
            & "values('" & strAddress & "','" & strAddress & "')"
        Set rsHome = cnHome.Execute(strSQL)
    Else
        strSQL = "UPDATE URL_TERMS SET HOME = '" & strAddress & "', " _
            & "URL = '" & strAddress & "' WHERE URL_ID = " & lngID
        Set rsHome = cnHome.Execute(strSQL)
    End If
    cboAddress.Text = strAddress
    If cnHome.State = adStateOpen Then
        cnHome.Close
        Set cnHome = Nothing
        Set rsHome = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Set_Home:
    Screen.MousePointer = vbDefault
    Exit Function
    Resume
End Function

Private Sub cmdAdd_Click()
    If Trim(txtValue.Text) = vbNullString Then
        MsgBox "Enter a term first.", vbExclamation
        txtValue.SetFocus
        Exit Sub
    End If

    lstTerms.AddItem txtValue.Text
   ' bSave = True
    txtValue.Text = vbNullString
    txtValue.SetFocus
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    If lstTerms.ListIndex = -1 Then
        MsgBox "Select an item from the listbox first", vbExclamation
        Exit Sub
    End If

    If lstTerms.ListIndex = lstTerms.listcount - 1 Then Exit Sub
    Reorder_ListBox lstTerms, False
   ' bSave = True
End Sub

Private Sub cmdHome_Click()
    If cboAddress.Text <> vbNullString Then
        Set_Home
    Else
        MsgBox "Please enter URL first!", vbInformation
        Exit Sub
    End If
    
    Exit Sub
End Sub


Private Sub cmdRemove_Click()
Dim strSQL As String
Dim lintIndex As Integer
Dim R As Integer


    If lstTerms.ListIndex = -1 Then
        MsgBox "Select an item from the listbox first", vbExclamation
        Exit Sub
    End If
    
    lintIndex = lstTerms.ListIndex
    lstTerms.RemoveItem lintIndex
    
End Sub

Private Sub cmdReplace_Click()
Dim strSQL As String
Dim IAnswer As Integer
Dim lintIndex As Integer
Dim sItemData As String

    If Trim(txtValue.Text) = vbNullString Then
        MsgBox "Enter a term first.", vbExclamation
        Exit Sub
    End If

    If lstTerms.ListIndex = -1 Then
        MsgBox "Select an item from the listbox first", vbExclamation
        Exit Sub
    End If
    
    lintIndex = lstTerms.ListIndex
    lstTerms.RemoveItem lintIndex
    lstTerms.AddItem Trim(txtValue.Text), lintIndex
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    If lstTerms.listcount > 0 Then
        SaveList
    End If
End Sub

Private Sub cmdUp_Click()
    If lstTerms.ListIndex = -1 Then
        MsgBox "Select an item from the listbox first", vbExclamation
        Exit Sub
    End If
    
    If lstTerms.ListIndex = 0 Then Exit Sub
    Reorder_ListBox lstTerms, True
   ' bSave = True
End Sub


Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
    End If
    If cboAddress.ListIndex <= 0 Then Set_Home
    centerForm Me
End Sub

Private Sub Form_Load()
    mFirst = True
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next
 
    Unload Me
    Exit Sub
End Sub


