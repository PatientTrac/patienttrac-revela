VERSION 5.00
Object = "{13ADDE2B-B303-4618-975F-648031250148}#5.0#0"; "PolarSpellChecker.dll"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmNurseTemplates 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terms Maintanence"
   ClientHeight    =   3915
   ClientLeft      =   8415
   ClientTop       =   3150
   ClientWidth     =   6870
   ControlBox      =   0   'False
   Icon            =   "frmNurseTemplates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6870
   Begin VB.Frame fraTerms 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Terms"
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
      Height          =   3375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6825
      Begin SPELLCHECKERLibCtl.SpellChecker SpellChecker1 
         Left            =   6450
         OleObjectBlob   =   "frmNurseTemplates.frx":058A
         Top             =   900
      End
      Begin VB.ListBox lstTerms 
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
         Height          =   1950
         Left            =   150
         TabIndex        =   4
         Top             =   1200
         Width           =   5985
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
         Height          =   525
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   5985
      End
      Begin PT_XP_Button.PT_XP cmdAdd 
         Height          =   345
         Left            =   180
         TabIndex        =   5
         Top             =   750
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
         Picture         =   "frmNurseTemplates.frx":05AE
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Add"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmNurseTemplates.frx":0B48
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdUp 
         Height          =   405
         Left            =   6300
         TabIndex        =   6
         Top             =   1620
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
         Picture         =   "frmNurseTemplates.frx":0E9A
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmNurseTemplates.frx":1614
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdDown 
         Height          =   405
         Left            =   6300
         TabIndex        =   7
         Top             =   2100
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
         Picture         =   "frmNurseTemplates.frx":1D26
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   ""
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmNurseTemplates.frx":24A0
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdReplace 
         Height          =   345
         Left            =   1380
         TabIndex        =   8
         Top             =   750
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
         Picture         =   "frmNurseTemplates.frx":2BB2
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Replace"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmNurseTemplates.frx":314C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRemove 
         Height          =   345
         Left            =   2580
         TabIndex        =   9
         Top             =   750
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
         Picture         =   "frmNurseTemplates.frx":349E
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Remove"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmNurseTemplates.frx":3A38
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCheckComment 
         Height          =   315
         Left            =   6270
         TabIndex        =   10
         ToolTipText     =   "Select for Quotation!"
         Top             =   210
         Width           =   435
         _ExtentX        =   767
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
         MaskColor       =   128
         Picture         =   "frmNurseTemplates.frx":3D8A
         AppearanceThemes=   3
         BorderWidth     =   2
         TransparentColor=   14737632
         Caption         =   ""
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmNurseTemplates.frx":4324
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   345
      Left            =   2430
      TabIndex        =   0
      Top             =   3480
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
      Picture         =   "frmNurseTemplates.frx":4676
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmNurseTemplates.frx":4A10
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDone 
      Height          =   345
      Left            =   3495
      TabIndex        =   1
      Top             =   3480
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
      Picture         =   "frmNurseTemplates.frx":4D62
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmNurseTemplates.frx":5774
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmNurseTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fType As String
Public fConnectCR As String
Public fConnectTM As String
Public fCaption As String
Public fGroup As String
Public fSubDomain As String
Public mFirst As Boolean
Private bSave As Boolean
Private bAccept As Boolean
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

Private Sub cmdAdd_Click()
    If mlngLogonUser > 0 Then
        Get_SecurityRole mlngLogonUser
        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
            If UCase(clsPracticeRoles.RoleTitle) <> "Administrator" Then
                MsgBox "You are not an Administrator and cannot modify Terms,  please sent suggestions to support@patienttrac.com!", vbInformation, "PatientTrac Security Management"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
    If Trim(txtValue.Text) = vbNullString Then
        MsgBox "Enter a term first.", vbInformation, "PatientTrac Controlled Vocabulary Management"
        txtValue.SetFocus
        Exit Sub
    End If
    lstTerms.AddItem txtValue.Text
    bSave = True
    txtValue.Text = vbNullString
    txtValue.SetFocus
    Exit Sub
End Sub

Private Sub cmdAdd_GotFocus()
   ' cmdAdd_KeyPress 13
End Sub


Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       KeyAscii = 0
'        cmdAdd_Click
'    End If
End Sub

Private Sub cmdCheckComment_Click()
Dim strResultText As String
Dim vExitCode As Variant
Dim fso As New FileSystemObject

    On Error GoTo Error_Spell
    If fso.FileExists(mstrProgram & "\FirstNotebook\PatientTrac EMR\Dictionary\SpellUS.lex") Then
        SpellChecker1.OpenDictionary mstrProgram & "\FirstNotebook\PatientTrac EMR\Dictionary\SpellUS.lex"
        txtValue.SetFocus
        strResultText = SpellChecker1.CheckText(txtValue.Text, vExitCode)
        If Trim(txtValue.Text) = Trim(strResultText) Then
            MsgBox "No spelling errors found!", vbInformation, "PatientTrac Spell Checker"
        End If
        txtValue.Text = Trim(strResultText)
        SpellChecker1.CloseDictionary mstrProgram & "\FirstNotebook\PatientTrac EMR\Dictionary\SpellUS.lex"
    End If
    Exit Sub
Error_Spell:
    Exit Sub
    Resume
End Sub

Private Sub cmdDone_Click()
    Dim iAnswer As Integer
    Dim bReturn As Boolean
    
    If bSave Then
       iAnswer = MsgBox("Do you want to accept changes you made?", vbYesNoCancel + vbQuestion)
       If iAnswer = vbYes Then cmdSave_Click
       If iAnswer = vbCancel Then cmdDone_Click
    End If
    Unload Me
End Sub
Private Sub cmdDown_Click()
    If lstTerms.ListIndex = -1 Then
        MsgBox "Select an item from the listbox first", vbExclamation
        Exit Sub
    End If

    If lstTerms.ListIndex = lstTerms.listcount - 1 Then Exit Sub
    Reorder_ListBox lstTerms, False
    bSave = True
End Sub
Private Sub cmdRemove_Click()
Dim strSQL As String
Dim lintIndex As Integer
Dim R As Integer

    If mlngLogonUser > 0 Then
        Get_SecurityRole mlngLogonUser
        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
            If UCase(clsPracticeRoles.RoleTitle) <> "Administrator" Then
                MsgBox "You are not an Administrator and cannot modify Terms,  please sent suggestions to support@patienttrac.com!", vbInformation, "PatientTrac Security Management"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
    If lstTerms.ListIndex = -1 Then
        MsgBox "Select an item from the listbox first", vbExclamation, "PatientTrac Controlled Vocabulary Management"
        Exit Sub
    End If
    lintIndex = lstTerms.ListIndex
    lstTerms.RemoveItem lintIndex
    txtValue.Text = vbNullString
    bSave = True
    Exit Sub
End Sub
Private Sub cmdReplace_Click()
Dim strSQL As String
Dim iAnswer As Integer
Dim lintIndex As Integer
Dim sItemData As String


    If mlngLogonUser > 0 Then
        Get_SecurityRole mlngLogonUser
        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
            If UCase(clsPracticeRoles.RoleTitle) <> "Administrator" Then
                MsgBox "You are not an Administrator and cannot modify Terms,  please sent suggestions to support@patienttrac.com!", vbInformation, "PatientTrac Security Management"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
    If Trim(txtValue.Text) = Empty Then
        MsgBox "Enter a term first.", vbExclamation, "PatientTrac Controlled Vocabulary Management"
        Exit Sub
    End If

    If lstTerms.ListIndex = -1 Then
        MsgBox "Select an item from the listbox first", vbExclamation, "PatientTrac Controlled Vocabulary Management"
        Exit Sub
    End If
    
    lintIndex = lstTerms.ListIndex
    lstTerms.RemoveItem lintIndex
    lstTerms.AddItem Trim(txtValue.Text), lintIndex
    txtValue.Text = vbNullString
    bSave = True
    
    Exit Sub
End Sub
Private Sub cmdSave_Click()
    SaveList
End Sub
Private Sub cmdUp_Click()
    If lstTerms.ListIndex = -1 Then
        MsgBox "Select an item from the listbox first", vbExclamation
        Exit Sub
    End If
    
    If lstTerms.ListIndex = 0 Then Exit Sub
    Reorder_ListBox lstTerms, True
    bSave = True
End Sub

Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
    End If
End Sub

Private Sub Form_Load()
Dim Idx As Integer
    bSave = False
   mFirst = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next
    Unload Me
  Exit Sub
End Sub

Private Sub lstTerms_Click()
     txtValue.Text = lstTerms.Text
End Sub
Public Sub ReadData()
Dim strSQL As String


    Dim R As Integer
    
    Me.Top = (Screen.Height - Me.Height) / 2 + 300
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Caption = Me.Caption & " >> " & fCaption
    
    LoadListBox fSubDomain
End Sub
Private Sub txtValue_Change()
    bAccept = True
End Sub
Public Function LoadListBox(ByVal vstrSubDomain As String)
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim lngRow As Long

    On Error GoTo Error_LoadListBox
    Screen.MousePointer = vbHourglass
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Select DISTINCT ATTRBTE_VALID_VALUE, valid_value_seq_nbr " _
        & "From ATTR_VALID_VALUE " _
        & "Where VALID_VALUE_SUBDOM = '" & vstrSubDomain & "' ORDER BY valid_value_seq_nbr"
    Set rsTerms = cnTerms.Execute(strSQL)
    With rsTerms
        Do Until .EOF
            If Trim((!ATTRBTE_VALID_VALUE)) <> vbNullString Then
                lstTerms.AddItem (!ATTRBTE_VALID_VALUE)
            End If
            .MoveNext
        Loop
    End With
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set rsTerms = Nothing
        Set cnTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_LoadListBox:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LoadListBox", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function
Private Sub SaveList()
Dim Idx As Integer
    
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
        Open_DB fConnectTM
            
        For Idx = 0 To lstTerms.listcount - 1
            cn.A_ATT_ATTR_VALID_VALUE "X", fSubDomain, Idx + 1, lstTerms.List(Idx), Null
        Next Idx
        cn.A_ATT_ATTR_VALID_VALUE "D", fSubDomain, lstTerms.listcount, "Cleanup", Null
            
        Close_DB fConnectTM
        
        bSave = False
                    
        Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
End Sub

