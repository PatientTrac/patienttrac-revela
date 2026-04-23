VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmAssessment 
   BackColor       =   &H00C8D0D4&
   Caption         =   "Assessment Summary"
   ClientHeight    =   8865
   ClientLeft      =   4380
   ClientTop       =   3690
   ClientWidth     =   12960
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8865
   ScaleWidth      =   12960
   Begin VB.Frame fraOCX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   4320
      TabIndex        =   8
      Top             =   3150
      Visible         =   0   'False
      Width           =   5865
      Begin VB.CommandButton cmdOCX_Action 
         Caption         =   "Action"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1830
         TabIndex        =   10
         Top             =   3210
         Width           =   825
      End
      Begin VB.CommandButton cmdCloseOCX 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2700
         TabIndex        =   9
         Top             =   3210
         Width           =   825
      End
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   7305
      Left            =   3870
      TabIndex        =   5
      Top             =   870
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12885
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   1
      ImageList       =   "ImageList"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   7335
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   12938
      _Version        =   393217
      Indentation     =   485
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   1
      ImageList       =   "ImageList"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2190
      Top             =   8250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssessment.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssessment.frx":00FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssessment.frx":01F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssessment.frx":02EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssessment.frx":0740
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssessment.frx":0B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssessment.frx":0C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssessment.frx":0E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssessment.frx":1720
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssessment.frx":3152
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel pnlAssesment 
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   450
      Width           =   3765
      _Version        =   65536
      _ExtentX        =   6641
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "  Assessment Structure"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Alignment       =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      Begin VB.CheckBox chkShowData 
         Caption         =   "Show Data Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4380
         TabIndex        =   11
         Top             =   90
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox chkPrimaryVal 
         Caption         =   "Primary Values"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   60
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox chkAllEncounters 
         Caption         =   "All Encounters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1230
         TabIndex        =   6
         Top             =   60
         Value           =   1  'Checked
         Width           =   1425
      End
   End
   Begin Threed.SSPanel pnlContent2 
      Height          =   315
      Left            =   3870
      TabIndex        =   3
      Top             =   450
      Width           =   9015
      _Version        =   65536
      _ExtentX        =   15901
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "  Contents of"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Alignment       =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6180
      TabIndex        =   0
      Top             =   8340
      Width           =   975
   End
End
Attribute VB_Name = "frmAssessment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fobjOCX As VBControlExtender
Private fobjLib As PatientTracTypeLib.iClinical
Private fstrCurrentEncounterDate As String
Private fstrLastEncounterDate As String
Private fbooLastEncounterOpen As Boolean
Private flngLastEncounterID As Long
Private fbooDataChanged As Boolean
Private fstrLastNodeKey As String
Private fintLastNodeIndex As Integer
Private fintObsCheckList As Integer
Private fintObsValueList As Integer
Private fintObsGroupingID As Integer
Private fintObs As Integer
Private fintLab As Integer
Private fintMeds As Integer
Private fstrLabDescr As String
Private fbooAlreadyLoaded As Boolean
Private Sub chkAllEncounters_Click()
    TreeView2.Nodes.Clear
End Sub

Private Sub chkPrimaryVal_Click()
    TreeView2.Nodes.Clear
End Sub

Private Sub cmdClose_Click()
    frmMDI.fbooByPass = True
    'frmMDI.cmdToolBar(frmMDI.cmdToolBar.Count - 1).Value = True
    Unload Me
End Sub

Private Sub cmdCloseOCX_Click()
    On Error Resume Next
    
    If fbooDataChanged Then
        Display_PatientData fstrLastNodeKey, fintLastNodeIndex
    End If
    
    Controls.Remove "ctlOCX"
    fraOCX.Visible = False
    fbooDataChanged = False
    
End Sub

Private Sub cmdOCX_Action_Click()

    fobjLib.WriteClinicalData
    
    If cmdOCX_Action.Caption = "Insert" Then
        cmdOCX_Action.Caption = "Update"
    End If
    
    fbooDataChanged = True
End Sub

Private Sub Form_Load()
Dim i%
Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim sGroup As String
Dim sForm As String
Dim sProvider As String
Dim sFacility As String
Dim sAppt As String
Dim sEnd As String
Dim sStart As String
Dim sIntrvn As String
Dim sTxnmy As String
Dim sUser As String
Dim sPatient As String
Dim varVals
Dim lngencounterid As Long
'
'If piPatientID < 1 Then
'    MsgBox "Select a patient first.", vbInformation
'    Unload Me
'    Exit Sub
'End If
'
'
'If piEncounterID < 1 Then
'    MsgBox "Select an Encounter first.", vbInformation
'    Unload Me
'    Exit Sub
'End If

LogFile "Start Log for " & Me.Name, True

LogFile ""
LogFile "********************** Form_Load ***************************"
LogFile ""

'piPatientID = 6977    '3164
'piEncounterID = 2199 '2122    '1618

centerForm Me

strSQL = "Select encounter_id, max(ENCNTR_START_DATE) " _
    & "From Encounter " _
    & "Where patient_id = " & piPatientID & " and ENCNTR_START_DATE = ( Select  max(ENCNTR_START_DATE) From ENCOUNTER  Where patient_id = " & piPatientID & " ) " _
    & "Group By encounter_id"
        
    LogFile "": LogFile strSQL: LogFile ""
    
    bRet = adoSelectQuery(strSQL, psConnect(0), vRS, , iRows)
    If iRows > 0 Then
        lngencounterid = vRS(0, 0)
    End If
    
'Get last encounter and is it open........................
'strSQL = "Select encounter_id, ENCNTR_START_DATE, ENCNTR_END_DATE " & vbNewLine _
    & "From ENCOUNTER " & vbNewLine _
    & "Where encounter_id = (select max(encounter_id)From ENCOUNTER Where patient_id = " & piPatientID & " )"
    
strSQL = "Select encounter_id, ENCNTR_START_DATE, ENCNTR_END_DATE " & vbNewLine _
    & "From ENCOUNTER " & vbNewLine _
    & "Where encounter_id = " & lngencounterid
            
    LogFile "": LogFile strSQL: LogFile ""
    
    bRet = adoSelectQuery(strSQL, psConnect(0), vRS, , iRows)
    If iRows > 0 Then
        flngLastEncounterID = vRS(0, 0)
        fstrLastEncounterDate = vRS(1, 0)
        If IsNull(vRS(2, 0)) Then fbooLastEncounterOpen = True
    End If


strSQL = "Select  " & vbNewLine _
    & "e.APPOINTMENT_DATE, " & vbNewLine _
    & "e.ENCNTR_END_DATE, " & vbNewLine _
    & "e.ENCNTR_START_DATE, " & vbNewLine _
    & "e.ENCOUNT_CPT_INTRVN, " & vbNewLine _
    & "t.CPT_TXNMY_NAME, " & vbNewLine _
    & "f.FACILITY_NAME, " & vbNewLine _
    & "rtrim(p.FIRST_NAME) + ' ' + rtrim(p.MIDDLE_NAME) + ' ' + rtrim(p.LAST_NAME) + ', ' + rtrim(p.TITLE), " & vbNewLine _
    & "rtrim(x.FIRST_NAME) + ' ' + rtrim(x.MIDDLE_NAME) + ' ' + rtrim(x.LAST_NAME), " & vbNewLine _
    & "e.USER_ID " & vbNewLine _
    & "From ENCOUNTER e, FACILITIES f, PROVIDERS p, TERMS.dbo.BT_CPT_TXNMY t, PATIENT x " & vbNewLine _
    & "Where e.facility_id *= f.facility_id and e.provider_id *= p.provider_id And e.ENCOUNT_CPT_TXNMY_ID *= t.cpt_txnmy_id And e.patient_id = x.patient_id And " & vbNewLine _
    & "e.patient_id = " & piPatientID & " And encounter_id = " & piEncounterID
    
    LogFile "": LogFile strSQL: LogFile ""
        
    bRet = adoSelectQuery(strSQL, psConnect(0), vRS, , iRows)
    
    
If iRows > 0 Then
    For i = LBound(vRS, 2) To UBound(vRS, 2)
        If Not IsNull(vRS(0, i)) Then sAppt = vRS(0, i)
        If Not IsNull(vRS(1, i)) Then sEnd = vRS(1, i)
        If Not IsNull(vRS(2, i)) Then sStart = vRS(2, i): fstrCurrentEncounterDate = sStart
        If Not IsNull(vRS(3, i)) Then sIntrvn = vRS(3, i)
        If Not IsNull(vRS(4, i)) Then sTxnmy = Trim(vRS(4, i))
        If Not IsNull(vRS(5, i)) Then sFacility = vRS(5, i)
        If Not IsNull(vRS(6, i)) Then sProvider = vRS(6, i)
        If Not IsNull(vRS(7, i)) Then sPatient = vRS(7, i)
        If Not IsNull(vRS(8, i)) Then sUser = vRS(8, i)
    Next i
End If
    'TreeView_AddItem Obj As Object, iParentLevel%, iLevel%, sKey$, sCap$, iImageType%, Optional bExpand, Optional bBold, Optional lColor
    TreeView_AddItem TreeView, 0, tvwChild, "", "Encounter", 5, True, True
    TreeView_AddItem TreeView, 1, tvwChild, "", "Patient: " & sPatient, 6
    TreeView_AddItem TreeView, 1, tvwChild, "", "Started: " & sStart, 6
    TreeView_AddItem TreeView, 1, tvwChild, "", "Ended: " & sEnd, 6
    TreeView_AddItem TreeView, 1, tvwChild, "", "Appointment: " & sAppt, 6
    TreeView_AddItem TreeView, 1, tvwChild, "", "Provider: " & sProvider, 6
    TreeView_AddItem TreeView, 1, tvwChild, "", "Facility: " & sFacility, 6
    TreeView_AddItem TreeView, 1, tvwChild, "", "Intervention: " & sIntrvn, 6
    TreeView_AddItem TreeView, 1, tvwChild, "", "Taxonomy: " & sTxnmy, 6
    TreeView_AddItem TreeView, 1, tvwChild, "", "User: " & sUser, 6

strSQL = "Select a.emr_grp_name, b.form_caption, b.form_id " & vbNewLine _
    & "From EMR_GROUP a, FORMS b, FORM_TAB c " & vbNewLine _
    & "Where a.emr_grp_id = b.emr_grp_id and " & vbNewLine _
    & "b.form_id = c.form_id and " & vbNewLine _
    & "a.emr_grp_id = 1 " & vbNewLine _
    & "Order By b.form_id, c.tab_id"
    
    LogFile "": LogFile strSQL: LogFile ""
    
    bRet = adoSelectQuery(strSQL, psConnect(1), vRS, , iRows)

If iRows > 0 Then
    For i = LBound(vRS, 2) To UBound(vRS, 2)
        If sGroup <> vRS(0, i) Then
            'TreeView_AddItem Obj As Object, iParentLevel%, iLevel%, sKey$, sCap$, iImageType%, Optional bExpand, Optional bBold, Optional lColor
            TreeView_AddItem TreeView, 0, 0, "Group" & i, CStr(vRS(0, i)) & " Screens", 9, True, True, &HC00000
        Else
            If sForm <> vRS(1, i) Then
                'TreeView_AddItem Obj As Object, iParentLevel%, iLevel%, sKey$, sCap$, iImageType%, Optional bExpand, Optional bBold, Optional lColor
                TreeView_AddItem TreeView, 1, tvwChild, "Form" & CStr(vRS(2, i)), CStr(vRS(1, i)), 7, , , &HC00000
            End If
            sForm = vRS(1, i)
        End If
        sGroup = vRS(0, i)
    Next i
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.Objects
      Set obj = Nothing
    Next obj
End Sub


Private Sub TreeView_Collapse(ByVal Node As MSComctlLib.Node)
Dim sKey

    sKey = Left(TreeView.Nodes.Item(Node.Index).Key, 1)
    
    Select Case sKey
        
        Case ""
            TreeView.Nodes.Item(Node.Index).Image = 4 'Down Arrow
          
        Case "G"
            TreeView.Nodes.Item(Node.Index).Image = 9
            
    End Select
    
End Sub

Private Sub TreeView_Expand(ByVal Node As MSComctlLib.Node)

Dim sKey

    sKey = Left(TreeView.Nodes.Item(Node.Index).Key, 1)
    
    Select Case sKey
        
        Case ""
            TreeView.Nodes.Item(Node.Index).Image = 5  'Up Arrow
            
        Case "G"
            TreeView.Nodes.Item(Node.Index).Image = 9
            
    End Select
   
End Sub


Private Sub TreeView_NodeClick(ByVal Node As MSComctlLib.Node)
Dim sKey As String
Dim iFormID As Integer
Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim r As Integer
Dim sParentTab As String
Dim sChildTab As String
Dim X As Integer
Dim iParent As Integer
Dim iCount As Integer
Dim iData As Integer
Dim sClass As String
Dim sOCX As String

LogFile ""
LogFile "********************** TreeView_NodeClick ***************************"
LogFile ""

Screen.MousePointer = vbHourglass

sKey = Node.Key

If sKey = vbNullString Or Left(sKey, 4) <> "Form" Then Exit Sub

Select Case TreeView.Nodes(Node.Index).Text

    Case "Encounter:"

Case Else
    Display_PatientData Node.Key, Node.Index

End Select

LogFile "Clicked on " & TreeView.Nodes(Node.Index).Text: LogFile ""

Screen.MousePointer = vbDefault
End Sub


Private Function Get_Clinical_Data(iActX As Integer, Optional iTabID As Integer) As Integer

Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim r As Integer
Dim bFirst As Boolean
Dim bFirstEncounter As Boolean
Dim X As Integer
Dim sPrimVal As String
Dim sStartDate As String
Dim iParent As Integer
Dim sControl As String
Dim sOCX As String
Dim sText As String
Dim sLine() As String
Dim k As Integer
Dim c As Integer
Dim lbooGetSummary As Boolean
Dim lvarSumRS

LogFile ""
LogFile "********************** Get_Clinical_Data ***************************"
LogFile ""

'LogFile "Passed Values:  iActX = " & iActX & ", Optional iTabID = " & iTabID
'LogFile ""

sOCX = TreeView2.Nodes.Item(TreeView2.Nodes.Count).Tag
sControl = TreeView2.Nodes.Item(TreeView2.Nodes.Count).Text 'TreeView2.Nodes.Item(TreeView2.Nodes.Count).Tag


LogFile "Passed Values:  iActX = " & iActX & ", Optional iTabID = " & iTabID & ",  OCX: " & sOCX & ", Control: " & sControl
LogFile ""


Select Case sControl

    Case "Follicular"
        lbooGetSummary = True
    
    Case Else
        lbooGetSummary = False
    
End Select


If chkPrimaryVal.Value Then
    sPrimVal = " And pri_value_ind = 1 "
Else
    sPrimVal = ""
End If

If chkAllEncounters.Value Then

    
    'strSQL = "Select a.clinical_value_ref + rtrim(cast(a.sequence_no as char)) as ref, a.clinical_value , a.sequence_no, b.encntr_start_date, a.encounter_id, a.actx_id  " & vbNewLine _
        & "From PATIENT_CLINICAL a, ENCOUNTER b  " & vbNewLine _
        & "Where a.actx_id = " & iActX & " And a.patient_id = b.patient_id and a.encounter_id = b.encounter_id And  " & vbNewLine _
        & "a.patient_id = " & piPatientID & sPrimVal
      
    strSQL = "Select distinct rtrim(case when c.VALID_VALUE_DEF is null then a.clinical_value_ref else c.VALID_VALUE_DEF end) + rtrim(cast(a.sequence_no as char))as ref, " _
            & "a.clinical_value , " & vbNewLine _
            & "a.sequence_no, " & vbNewLine _
            & "b.encntr_start_date, " & vbNewLine _
            & "a.encounter_id, a.actx_id  " & vbNewLine _
        & "From PATIENT_CLINICAL a, ENCOUNTER b, terms.dbo.ATTR_VALID_VALUE c " & vbNewLine _
        & "Where a.actx_id = " & iActX & " And " & vbNewLine _
            & "a.patient_id = b.patient_id and " & vbNewLine _
            & "a.encounter_id = b.encounter_id And " & vbNewLine _
            & "c.VALID_VALUE_SUBDOM =* a.clinical_value_ref And " & vbNewLine _
            & "a.patient_id = " & piPatientID & sPrimVal
          
    strSQL = strSQL & vbNewLine & " Union Select a.clinical_value_ref + rtrim(cast(a.sequence_no as char)) as ref, a.clinical_notes , a.sequence_no, b.encntr_start_date, a.encounter_id, a.actx_id   " & vbNewLine _
        & "From PATIENT_NOTES a, ENCOUNTER b  " & vbNewLine _
        & "Where a.actx_id = " & iActX & " And a.patient_id = b.patient_id and a.encounter_id = b.encounter_id And rtrim(a.clinical_notes)<>'' And " & vbNewLine _
        & "a.patient_id = " & piPatientID & sPrimVal
      
    strSQL = strSQL & vbNewLine & " Union Select a.intrvn_desc + rtrim(cast(a.sequence_no as char)) as ref, a.intrvn_code , a.sequence_no, b.encntr_start_date, a.encounter_id, a.actx_id   " & vbNewLine _
        & "From PATIENT_INTRVN a, ENCOUNTER b  " & vbNewLine _
        & "Where a.actx_id = " & iActX & " And a.encounter_id = b.encounter_id And  " & vbNewLine _
        & "a.patient_id = " & piPatientID & sPrimVal
            
    strSQL = strSQL & vbNewLine & " Union Select a.signature_ref  + rtrim(cast(a.sequence_no as char)) as ref,  'encoded signature' , a.sequence_no, b.encntr_start_date, a.encounter_id, a.actx_id   " & vbNewLine _
        & "From SIGNATURES  a, ENCOUNTER b  " & vbNewLine _
        & "Where a.actx_id = " & iActX & " And a.patient_id = b.patient_id and a.encounter_id = b.encounter_id And  " & vbNewLine _
        & "a.patient_id = " & piPatientID
        
     If fintObsCheckList = iActX Then
     
     'strSQL = strSQL & vbNewLine & " Union Select  rtrim(cast(a.obs_grouping_id as char)) + ', ' + rtrim(cast(a.obs_seq as char)) + ':' + rtrim(cast(obs_checklist_value as char)) as ref, rtrim(cast(a.obs_checklist_id as char)), a.obs_seq, b.encntr_start_date, a.encounter_id, actx_id =' " & iActX & "'  " & vbNewLine _
        & "From PATIENT_OBS_CHECKLIST  a, ENCOUNTER b  " & vbNewLine _
        & " Where a.encounter_id = b.encounter_id And a.patient_id = " & piPatientID
     
     'rtrim(cast(a.obs_checklist_value as char))
    
     strSQL = strSQL & vbNewLine & " Union Select rtrim(c.obs_checklist_text) as ref, Case When obs_checklist_value = 1 then 'yes' else 'no' end, a.obs_grouping_id, b.encntr_start_date, a.encounter_id, actx_id =' " & iActX & "'  " & vbNewLine _
        & "From PATIENT_OBS_CHECKLIST a, ENCOUNTER b, terms.dbo.OBSERVATION_CHECKLIST c " & vbNewLine _
        & " Where a.encounter_id = b.encounter_id And a.OBS_CHECKLIST_ID = c.OBS_CHECKLIST_ID And a.OBS_GROUPING_ID = " & fintObsGroupingID & " And a.patient_id = " & piPatientID
          
     End If
     
     If fintObsValueList = iActX Then
     
     'strSQL = strSQL & vbNewLine & " Union Select  rtrim(cast(a.obs_grouping_id as char)) + ', ' + rtrim(cast(a.obs_seq as char)) + ':' + a.obs_value as ref, rtrim(cast(a.obs_value_id as char)), a.obs_seq, b.encntr_start_date, a.encounter_id, actx_id =' " & iActX & "'  " & vbNewLine _
        & "From PATIENT_OBS_VALUES  a, ENCOUNTER b  " & vbNewLine _
        & " Where a.encounter_id = b.encounter_id And a.patient_id = " & piPatientID
        
     strSQL = strSQL & vbNewLine & " Union Select rtrim(c.OBS_VALUE_LABEL) as ref, rtrim(a.obs_value), a.obs_grouping_id, b.encntr_start_date, a.encounter_id, actx_id =' " & iActX & "'  " & vbNewLine _
        & "From PATIENT_OBS_VALUES  a, ENCOUNTER b, terms.dbo.OBSERVATION_VALUES  c  " & vbNewLine _
        & " Where a.OBS_VALUE_ID = c.OBS_VALUE_ID And a.encounter_id = b.encounter_id And a.OBS_GROUPING_ID = " & fintObsGroupingID & " And a.patient_id = " & piPatientID
        
     End If
     
     If fintObs = iActX Then

     strSQL = strSQL & vbNewLine & " Union Select c.OBS_VALUE_LABEL as ref, rtrim(cast(a.obs_value_id as char)), a.obs_seq, b.encntr_start_date, a.encounter_id, actx_id =' " & iActX & "'  " & vbNewLine _
        & "From PATIENT_OBS_VALUES  a, ENCOUNTER b, terms.dbo.OBSERVATION_VALUES  c   " & vbNewLine _
        & " Where a.OBS_VALUE_ID = c.OBS_VALUE_ID And a.encounter_id = b.encounter_id And a.patient_id = " & piPatientID
        
     strSQL = strSQL & vbNewLine & " Union Select  rtrim(cast(a.obs_grouping_id as char)) + ', ' + rtrim(cast(a.obs_seq as char)) + ':' + rtrim(cast(obs_checklist_value as char)) as ref, rtrim(cast(a.obs_checklist_id as char)), a.obs_seq, b.encntr_start_date, a.encounter_id, actx_id =' " & iActX & "'  " & vbNewLine _
        & "From PATIENT_OBS_CHECKLIST  a, ENCOUNTER b  " & vbNewLine _
        & " Where a.encounter_id = b.encounter_id And a.patient_id = " & piPatientID
     End If
     
     If fintLab = iActX Then
     'strSQL = strSQL & vbNewLine & " Union Select a.cpt_intrvn_code + ',' + rtrim(cast( a.lab_value_id as char)) as ref, a.lab_result_value, a.lab_value_id, b.encntr_start_date, a.encounter_id, actx_id=' " & iActX & "'  " & vbNewLine _
        & " From LAB_RESULTS a , ENCOUNTER b,  terms.dbo.TAB_ACTIVE_X  c   " & vbNewLine _
        & " Where a.encounter_id = b.encounter_id And a.patient_id = b.patient_id And c.actx_intrv_cd = a.cpt_intrvn_code And c.tab_id = " & iTabID & vbNewLine _
        & " And a.patient_id = " & piPatientID
     
     strSQL = strSQL & vbNewLine & " Union Select rtrim(d.value_label) + ' (' + rtrim(a.cpt_intrvn_code) + ')' as ref, a.lab_result_value, a.lab_value_id, b.encntr_start_date, a.encounter_id, actx_id=' " & iActX & "'  " & vbNewLine _
        & " From LAB_RESULTS a, ENCOUNTER b, terms.dbo.TAB_ACTIVE_X  c, terms.dbo.BT_VALUE_PARMS d " & vbNewLine _
        & " Where a.encounter_id = b.encounter_id And " _
        & " a.patient_id = b.patient_id And " & vbNewLine _
        & " c.actx_intrv_cd = a.cpt_intrvn_code And " & vbNewLine _
        & " a.lab_value_id = d.value_id And " & vbNewLine _
        & " c.tab_id = " & iTabID & " And " & vbNewLine _
        & " a.patient_id = " & piPatientID
     End If
     
     
     If fintMeds = iActX Then
     strSQL = strSQL & vbNewLine & " Union Select rtrim(cast(a.patient_med_id as char)) + ', ' + rtrim(cast(a.pref_med_id as char)) as ref, a.ingredient_name, a.listing_seq_no, b.encntr_start_date, a.encounter_id, actx_id=' " & iActX & "'  " & vbNewLine _
        & " From PATIENT_MEDICATIONS a , ENCOUNTER b  " & vbNewLine _
        & " Where a.encounter_id = b.encounter_id And a.patient_id = b.patient_id  " & vbNewLine _
        & " And a.patient_id = " & piPatientID
     
     End If
     
     strSQL = strSQL & vbNewLine & " Order By b.encntr_start_date,ref, 2"
Else

    'strSQL = "Select clinical_value_ref + rtrim(cast(sequence_no as char)) as ref, clinical_value, sequence_no " & vbNewLine _
        & "From PATIENT_CLINICAL  " & vbNewLine _
        & "Where actx_id = " & iActX & " And  " & vbNewLine _
        & "patient_id = " & piPatientID & " And  " & vbNewLine _
        & "encounter_id = " & piEncounterID & sPrimVal
        
    strSQL = "Select distinct rtrim(case when c.VALID_VALUE_DEF is null then a.clinical_value_ref else c.VALID_VALUE_DEF end) + rtrim(cast(a.sequence_no as char))as ref, " _
            & "a.clinical_value , " & vbNewLine _
            & "a.sequence_no " & vbNewLine _
        & "From PATIENT_CLINICAL a, ENCOUNTER b, terms.dbo.ATTR_VALID_VALUE c " & vbNewLine _
        & "Where a.actx_id = " & iActX & " And " & vbNewLine _
            & "a.patient_id = b.patient_id and " & vbNewLine _
            & "encounter_id = " & piEncounterID & " And " & vbNewLine _
            & "c.VALID_VALUE_SUBDOM =* a.clinical_value_ref And " & vbNewLine _
            & "a.patient_id = " & piPatientID & sPrimVal
                
    strSQL = strSQL & vbNewLine & " Union Select clinical_value_ref + rtrim(cast(sequence_no as char)) as ref, clinical_notes, sequence_no  " & vbNewLine _
        & "From PATIENT_NOTES  " & vbNewLine _
        & "Where rtrim(clinical_notes)<>'' And actx_id = " & iActX & " And  " & vbNewLine _
        & "patient_id = " & piPatientID & " And  " & vbNewLine _
        & "encounter_id = " & piEncounterID & sPrimVal
        
    strSQL = strSQL & vbNewLine & " Union Select intrvn_desc + rtrim(cast(sequence_no as char)) as ref, intrvn_code, sequence_no   " & vbNewLine _
        & "From PATIENT_INTRVN  " & vbNewLine _
        & "Where actx_id = " & iActX & " And  " & vbNewLine _
        & "patient_id = " & piPatientID & " And  " & vbNewLine _
        & "encounter_id = " & piEncounterID & sPrimVal
        
    strSQL = strSQL & vbNewLine & " Union Select signature_ref + rtrim(cast(sequence_no as char)) as ref, 'encoded signature', sequence_no   " & vbNewLine _
        & "From SIGNATURES  " & vbNewLine _
        & "Where actx_id = " & iActX & " And  " & vbNewLine _
        & "patient_id = " & piPatientID & " And  " & vbNewLine _
        & "encounter_id = " & piEncounterID
        
        
     If fintObsCheckList = iActX Then
     
     strSQL = strSQL & vbNewLine & " Union Select rtrim(b.obs_checklist_text) as ref, Case When a.obs_checklist_value = 1 then 'yes' else 'no' end, a.obs_grouping_id  " & vbNewLine _
        & "From PATIENT_OBS_CHECKLIST a, terms.dbo.OBSERVATION_CHECKLIST b  " & vbNewLine _
        & "Where a.OBS_CHECKLIST_ID = b.OBS_CHECKLIST_ID And a.OBS_GROUPING_ID = " & fintObsGroupingID & " And a.patient_id =" & piPatientID & " And  " & vbNewLine _
        & "a.encounter_id = " & piEncounterID
        
     End If
     
     If fintObsValueList = iActX Then
     
     'strSQL = strSQL & vbNewLine & " Union Select rtrim(cast(obs_grouping_id as char)) + ', ' + rtrim(cast(obs_seq as char)) + ':' + obs_value as ref, rtrim(cast(obs_value_id as char)), obs_seq  " & vbNewLine _
        & "From PATIENT_OBS_VALUES  " & vbNewLine _
        & "Where patient_id =" & piPatientID & " And  " & vbNewLine _
        & "encounter_id = " & piEncounterID
        
     strSQL = strSQL & vbNewLine & " Union Select rtrim(b.OBS_VALUE_LABEL) as ref, rtrim(cast(a.obs_value_id as char)), a.obs_grouping_id  " & vbNewLine _
        & "From PATIENT_OBS_VALUES a, terms.dbo.OBSERVATION_VALUES b " & vbNewLine _
        & "Where a.OBS_VALUE_ID = b.OBS_VALUE_ID And a.OBS_GROUPING_ID = " & fintObsGroupingID & " And a.patient_id =" & piPatientID & " And  " & vbNewLine _
        & "a.encounter_id = " & piEncounterID
        
     End If
     
     If fintObs = iActX Then
     strSQL = strSQL & vbNewLine & " Union Select  rtrim(cast(obs_grouping_id as char)) + ', ' + rtrim(cast(obs_seq as char)) + ':' + obs_value as ref, rtrim(cast(obs_value_id as char)), obs_seq  " & vbNewLine _
        & "From PATIENT_OBS_VALUES  " & vbNewLine _
        & "Where patient_id =" & piPatientID & " And  " & vbNewLine _
        & "encounter_id = " & piEncounterID
        
     strSQL = strSQL & vbNewLine & " Union Select rtrim(cast(obs_grouping_id as char)) + ', ' + rtrim(cast(obs_seq as char)) + ':' + rtrim(cast(obs_checklist_value as char)) as ref, rtrim(cast(obs_checklist_id as char)), obs_seq  " & vbNewLine _
        & "From PATIENT_OBS_CHECKLIST  " & vbNewLine _
        & "Where patient_id =" & piPatientID & " And  " & vbNewLine _
        & "encounter_id = " & piEncounterID
     End If
     
     If fintLab = iActX Then
     'strSQL = strSQL & vbNewLine & " Union Select cpt_intrvn_code + ',' + rtrim(cast( lab_value_id as char)) as ref, lab_result_value, lab_value_id  " & vbNewLine _
        & " From LAB_RESULTS a, terms.dbo.TAB_ACTIVE_X  c     " & vbNewLine _
        & " Where patient_id =" & piPatientID & " And c.actx_intrv_cd = a.cpt_intrvn_code And c.tab_id = " & iTabID _
        & " And encounter_id = " & piEncounterID
     
     strSQL = strSQL & vbNewLine & " Union Select rtrim(d.value_label) + ' (' + rtrim(a.cpt_intrvn_code) + ')' as ref, lab_result_value, lab_value_id  " & vbNewLine _
        & " From LAB_RESULTS a, terms.dbo.TAB_ACTIVE_X  c, terms.dbo.BT_VALUE_PARMS d " & vbNewLine _
        & " Where patient_id =" & piPatientID & " And " & vbNewLine _
        & " c.actx_intrv_cd = a.cpt_intrvn_code And " & vbNewLine _
        & " a.lab_value_id = d.value_id And " & vbNewLine _
        & " c.tab_id = " & iTabID & " And " & vbNewLine _
        & " encounter_id = " & piEncounterID
        
     End If
     
     
     If fintMeds = iActX Then
     strSQL = strSQL & vbNewLine & " Union Select rtrim(cast(patient_med_id as char)) + ', ' + rtrim(cast(pref_med_id as char)) as ref, ingredient_name, listing_seq_no  " & vbNewLine _
        & " From PATIENT_MEDICATIONS  " & vbNewLine _
        & "Where patient_id =" & piPatientID & " And  " & vbNewLine _
        & "encounter_id = " & piEncounterID
     
     End If
             
    strSQL = strSQL & vbNewLine & " Order By ref, 2"
End If

    LogFile "": LogFile strSQL: LogFile ""
    'Debug.Print strSQL
    bRet = adoSelectQuery(strSQL, psConnect(0), vRS, , iRows)
    
'0 - clinical_value_ref
'1 - clinical_value
'2 - sequence_no
'3 - encntr_start_date
'4 - encounter_id
'5 - actx_id

    If iRows > 0 Then
    
        If lbooGetSummary Then
            GetSummaryData sControl, lvarSumRS
        End If
        
        X = TreeView2.Nodes.Count - 1
        iParent = X + 1
        
        For r = LBound(vRS, 2) To UBound(vRS, 2)
        
'            For c = LBound(vRS, 1) To UBound(vRS, 1)
'                Debug.Print "vRS(" & c & "," & r & ") = " & vRS(c, r)
'            Next c
           
            X = X + 1
            If chkAllEncounters.Value Then
             If sStartDate <> vRS(3, r) Then
                If Not bFirstEncounter Then  'Red Diamond icon
                    'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                    TreeView2.Nodes.Add X, tvwChild, "Encounter" & X & "-" & CStr(vRS(4, r)) & ":" & CStr(iActX), vRS(3, r), 10
                    bFirstEncounter = True
                    LogFile "TreeView2.Nodes.Add Relative=" & X & ", Relationship=Child, Key=Encounter" & X & "-" & CStr(vRS(4, r)) & ":" & CStr(iActX) & ", Text=" & vRS(3, r) & ", Image=10, SelectedImage:(not passed)"
                Else  'Red Diamond icon
                    'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                    TreeView2.Nodes.Add iParent, tvwChild, "Encounter" & X & "-" & CStr(vRS(4, r)) & ":" & CStr(iActX), vRS(3, r), 10
                    bFirst = False
                    LogFile "TreeView2.Nodes.Add Relative=" & iParent & ", Relationship=Child, Key=Encounter" & X & "-" & CStr(vRS(4, r)) & ":" & CStr(iActX) & ", Text=" & vRS(3, r) & ", Image=10, SelectedImage:(not passed)"
                End If
                    X = X + 1
                    TreeView2.Nodes.Item(X).Tag = sOCX   '& "." & sControl
                    TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                    
             End If
            End If
            
            If Not bFirst Then
                'TreeView2.Nodes.Add Relative, Relationship, Key,Text, Image, SelectedImage
                If lbooGetSummary Then
                    sText = ProcessSummaryData(sControl, lvarSumRS, CLng(vRS(4, r)), Format(vRS(3, r), "mm/dd/yyyy"))
                    'If sText = vbnullstring Then x = x - 1
                Else
                    sText = Trim(vRS(0, r)) & " = " & Trim(vRS(1, r))
                End If
                
                If Len(sText) > 70 Then
                    sText = Wrap_Text(sText, sLine(), 70)
                Else
                    ReDim sLine(0)
                    sLine(0) = sText
                End If
                
                If sLine(0) <> vbNullString Then
                    'Bullet icon
                    'TreeView2.Nodes.Add x, tvwChild, "Data" & x, vRS(0, r) & " = " & vRS(1, r), 6
                    'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                    TreeView2.Nodes.Add X, tvwChild, "Data" & X, sLine(0), 6
                    LogFile "TreeView2.Nodes.Add Relative=" & X & ", Relationship=Child, Key=Data" & X & ", Text=" & sText & ", Image=6, SelectedImage:(not passed)"
                    TreeView2.Nodes(TreeView2.Nodes.Count).Tag = vRS(2, r)
                    TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                    
                    For k = 1 To UBound(sLine)  'Bullet icon
                        TreeView2.Nodes.Add X, tvwChild, "SubData" & X & k, sLine(k)
                        TreeView2.Nodes(TreeView2.Nodes.Count).Tag = vRS(2, r)
                        TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                    Next k
                    
                    X = X + UBound(sLine)
                End If
                bFirst = True
            Else
            
                If lbooGetSummary Then
                    sText = ProcessSummaryData(sControl, lvarSumRS, CLng(vRS(4, r)), Format(vRS(3, r), "mm/dd/yyyy"))
                    If sText = vbNullString Then X = X - 1
                Else
                    sText = Trim(vRS(0, r)) & " = " & Trim(vRS(1, r))
                End If
                
                If Len(sText) > 70 Then
                    sText = Wrap_Text(sText, sLine(), 70)
                Else
                    ReDim sLine(0)
                    sLine(0) = sText
                End If
                
                'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                'TreeView2.Nodes.Add x, tvwNext, "Data" & x, vRS(0, r) & " = " & vRS(1, r), 6
                
                If sLine(0) <> vbNullString Then
                    TreeView2.Nodes.Add X, tvwNext, "Data" & X, sLine(0), 6
                    
                    For k = 1 To UBound(sLine)  'Bullet icon
                        TreeView2.Nodes.Add X, tvwChild, "SubData" & X & k, sLine(k)
                        TreeView2.Nodes(TreeView2.Nodes.Count).Tag = vRS(2, r)
                        TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                    Next k
                
                
                    LogFile "TreeView2.Nodes.Add Relative=" & X & ", Relationship=Next, Key=Data" & X & ", Text=" & vRS(0, r) & " = " & vRS(1, r) & ", Image=6, SelectedImage:(not passed)"
                End If
            End If
            
            TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
            
            If chkAllEncounters.Value Then sStartDate = vRS(3, r)
        Next r
        
        Get_Clinical_Data = UBound(vRS, 2) + 1 + UBound(sLine)
        
        
    End If

    
End Function



Private Function Wrap_Text(vstrText As String, vstrLine() As String, vintMaxLength As Integer) As String
Dim strLine As String
Dim strChar As String
Dim strRemainder As String
Dim i As Integer
Dim k As Integer
Dim X As Integer
Dim j As Integer
Dim strSection As String
Dim booReverse As Boolean

'Debug.Print "***************************************************************************"
'Debug.Print vstrText
'Debug.Print "---------------------------------------------------------------------------"

 strRemainder = vstrText
 strChar = ""
 strLine = ""
 k = 0

 Do Until strRemainder = "" And InStr(strRemainder, Chr(13)) = 0 And InStr(strRemainder, Chr(10)) = 0
    strChar = ""
    i = 0
    'Debug.Print
    'Debug.Print "k = " & k & " --------------------------------------------------------------------"
    'Debug.Print strRemainder
    'Debug.Print "--------------------------------------------------------------------"
    'Debug.Print "Len(strRemainder) = " & Len(strRemainder)
    
    While i < vintMaxLength And strChar <> Chr(13) And strChar <> Chr(10)
        i = i + 1
        strChar = Mid(strRemainder, i, 1)
        strChar = Trim(strChar)
    Wend
    
    If i > Len(strRemainder) Then i = Len(strRemainder) + 1
    
    'Debug.Print "i = " & i
    
    'If strChar<> vbnullstring Then
        'Debug.Print "strChar ascii = " & Asc(strChar)
    'End If
    
    If i >= vintMaxLength Then
        i = 0: strChar = ""
        'Start from character position "vintMaxLength" and go backwards examining each character for a space,comma, or period
         While i < Len(strRemainder) And strChar <> Space(1) And strChar <> "," And strChar <> "."
             strChar = Mid(strRemainder, vintMaxLength - i, 1)
             i = i + 1
         Wend
         
         strSection = Left(strRemainder, vintMaxLength - i + 1)
         booReverse = True
    Else
        strSection = Left(strRemainder, i - 1)
    End If
    
    strSection = Trim(strSection)
    
    'Strip leading carriage return.............
    If Left(strSection, 2) = vbCrLf Then
        strSection = Mid(strSection, 3, Len(strSection) - 2)
    End If
    
    'Debug.Print
    'Debug.Print "strSection start ==============="
    'Debug.Print "|" & strSection & "|"
    'Debug.Print "strSection end ==============="
    
    If Not booReverse Then
        strRemainder = Mid(strRemainder, i + 2)
    Else
        strRemainder = Mid(strRemainder, vintMaxLength - i + 1)
    End If
    
    strRemainder = Trim(strRemainder)
    
    'Debug.Print
    'Debug.Print "strRemainder start ==============="
    'Debug.Print "|" & strRemainder & "|"
    'Debug.Print "strRemainder end ==============="
    
    ReDim Preserve vstrLine(k)
    vstrLine(k) = strSection
    strLine = strLine & strSection & vbNewLine
    
    'strip trailing carriage returns......
    If InStr(3, vstrLine(k), Chr(13)) > 0 Or InStr(3, vstrLine(k), Chr(10)) > 0 Then
         While InStr(vstrLine(k), Chr(13)) > 0 Or InStr(vstrLine(k), Chr(10)) > 0
            vstrLine(k) = Left(vstrLine(k), Len(vstrLine(k)) - 1)
            strChar = Right(vstrLine(k), 1)
        Wend
    End If
    
    'Debug.Print
    'Debug.Print "/////////// vstrLine(" & k & ") = |" & vstrLine(k) & "|"
    booReverse = False
    
    k = k + 1

 Loop
 
 
 Wrap_Text = strLine & strRemainder
 
End Function

Private Sub TreeView2_Collapse(ByVal Node As MSComctlLib.Node)
Dim sKey
Dim sTag As String

    'sKey = TreeView2.Nodes.Item(Node.Index).Key
    
    If TreeView2.Nodes.Item(Node.Index).Image = 2 Then
        TreeView2.Nodes.Item(Node.Index).Image = 1
    End If
    
    'sTag = TreeView2.Nodes.Item(Node.Index).Tag
    'sKey = Left(TreeView2.Nodes.Item(Node.Index).Key, 1)
End Sub

Private Sub TreeView2_Expand(ByVal Node As MSComctlLib.Node)
    If TreeView2.Nodes.Item(Node.Index).Image = 1 Then
        TreeView2.Nodes.Item(Node.Index).Image = 2
    End If
End Sub

Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
Dim strOCX_Name As String
Dim strOCX_Control As String
Dim sTag As String
Dim intNextIndex As Integer
Dim strSelectedEncounterDate As String
Dim strControl As String

    On Error GoTo Error_LoadOCX
    
    strSelectedEncounterDate = TreeView2.Nodes.Item(Node.Index).Text

    If Node.Image = 3 Or Node.Image = 10 Then
        fraOCX.Visible = True
        
        'sTag = TreeView2.Nodes.Item(Node.Index).Tag
        strControl = TreeView2.Nodes.Item(Node.Index).Tag
        
        'If Node.Image = 3 Then
            'strOCX_Name = TreeView2.Nodes.Item(Node.Index).Tag
            'strOCX_Control = Node.Text
           ' strControl = TreeView2.Nodes.Item(Node.Index).Tag
        'Else
            'strOCX_Name = Left(sTag, InStr(sTag, ".") - 1)
            'strOCX_Control = Mid(sTag, InStr(sTag, ".") + 1)
           ' strControl = sTag
        'End If
        
        Licenses.Add strControl   'strOCX_Name & "." & strOCX_Control
        
        'Set fobjOCX = Controls.Add(strOCX_Name & "." & strOCX_Control, "ctlOCX", fraOCX)
        Set fobjOCX = Controls.Add(strControl, "ctlOCX", fraOCX)
        fobjOCX.Visible = True
        fobjOCX.Left = 100
        fobjOCX.Top = 150
        fraOCX.Width = fobjOCX.Width + 200
        fraOCX.Height = fobjOCX.Height + 600
        CenterControl fraOCX, Me
        
        'Set fobjLib = fobjOCX.object
   
        'Is there data below?................................
        If Node.Index < TreeView2.Nodes.Count Then
            If chkAllEncounters.Value And TreeView2.Nodes.Item(Node.Index + 1).Image = 10 Then
                intNextIndex = Node.Index + 2
            Else
                intNextIndex = Node.Index + 1
            End If
        Else
            intNextIndex = TreeView2.Nodes.Count
        End If
        
        If intNextIndex > TreeView2.Nodes.Count Then
            cmdOCX_Action.Caption = "Insert"
        Else
            If TreeView2.Nodes.Item(intNextIndex).Image = 6 Then
                cmdOCX_Action.Caption = "Update"
            Else
                cmdOCX_Action.Caption = "Insert"
            End If
        End If
   
   If Not IsDate(strSelectedEncounterDate) Then strSelectedEncounterDate = fstrCurrentEncounterDate
   
        If fstrLastEncounterDate = strSelectedEncounterDate And fbooLastEncounterOpen Then
             cmdCloseOCX.Top = fobjOCX.Height + 175
             cmdCloseOCX.Left = (fraOCX.Width - cmdCloseOCX.Width) / 2 + cmdCloseOCX.Width / 2 + 15
             
             cmdOCX_Action.Visible = True
             cmdOCX_Action.Top = fobjOCX.Height + 175
             cmdOCX_Action.Left = (fraOCX.Width - cmdOCX_Action.Width) / 2 - cmdOCX_Action.Width / 2 - 15
        Else
             cmdOCX_Action.Visible = False
             cmdCloseOCX.Top = fobjOCX.Height + 175
             cmdCloseOCX.Left = (fraOCX.Width - cmdCloseOCX.Width) / 2
        End If
   
        Load_OCX Node.Index, Node.Image
    End If
    
    Exit Sub
    
Error_LoadOCX:
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 734 Then Resume Next
    ErrHandler _
     "frmAssessment.cmdSave_Click", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub Load_OCX(vlngIndex As Long, vintImage As Integer)
Dim intNextIndex As Integer
Dim strData As String
Dim obj As Object
Dim i As Integer
Dim k As Integer
Dim intActX_ID As Integer
Dim lngEnctr As Long
Dim sKey As String
Dim p1 As Integer
Dim p2 As Integer
Dim strCTP As String

On Error GoTo Error_LoadOCX

LogFile ""
LogFile "********************** Load_OCX ***************************"
LogFile ""

LogFile "Passed Values:  vlngIndex = " & vlngIndex & ", vintImage = " & vintImage
LogFile ""

    If vintImage = 3 Then
        intActX_ID = Mid(TreeView2.Nodes.Item(vlngIndex).Key, InStr(TreeView2.Nodes.Item(vlngIndex).Key, "-") + 1)
        lngEnctr = piEncounterID
        
        If chkAllEncounters.Value Then i = 2 Else i = 1
        
        If TreeView2.Nodes.Item(vlngIndex).Tag = "OBS.CheckList" Then
            'strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Text
            'strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
            strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Tag
        End If
        
        If TreeView2.Nodes.Item(vlngIndex).Tag = "OBS.ValueList" Then
            'strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Text
            'strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
            strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Tag
        End If
        
'        If InStr(TreeView2.Nodes.Item(vlngIndex + i - 1).Tag, "OBS") > 0 Then
'            strCTP = TreeView2.Nodes.Item(vlngIndex + i).Text
'            strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
'        End If
        
        If InStr(TreeView2.Nodes.Item(vlngIndex + i - 1).Tag, "EMR_Lab") > 0 Then
            strCTP = TreeView2.Nodes.Item(vlngIndex + i).Text
            p1 = InStr(strCTP, "(") + 1
            p2 = InStr(strCTP, ")")
            strCTP = Mid(strCTP, p1, p2 - p1)
            'strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
        End If
        
'        If TreeView2.Nodes.Item(vlngIndex + i - 1).Tag = "OBS.CheckList" Then
'            strCTP = TreeView2.Nodes.Item(vlngIndex + i).Text
'            strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
'        End If
'
'        If TreeView2.Nodes.Item(vlngIndex + 1).Tag = "OBS.ValueList" Then
'            strCTP = TreeView2.Nodes.Item(vlngIndex + i).Text
'            strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
'        End If
'
'        If TreeView2.Nodes.Item(vlngIndex + 1).Tag = "OBS.Observations" Then
'            strCTP = TreeView2.Nodes.Item(vlngIndex + i).Text
'            strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
'        End If
        
    Else
        sKey = TreeView2.Nodes.Item(vlngIndex).Key
        intActX_ID = Mid(sKey, InStr(sKey, ":") + 1)
        p1 = InStr(sKey, "-") + 1
        p2 = InStr(sKey, ":")
        lngEnctr = Mid(sKey, p1, p2 - p1)
        
        If TreeView2.Nodes.Item(vlngIndex).Tag = "OBS.CheckList" Then
            'strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Text
            'strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
            strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Tag
        End If
        
        If TreeView2.Nodes.Item(vlngIndex).Tag = "OBS.ValueList" Then
            'strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Text
            'strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
            strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Tag
        End If
        
        If TreeView2.Nodes.Item(vlngIndex).Tag = "OBS.Observations" Then
            strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Text
            strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
        End If
        
        If TreeView2.Nodes.Item(vlngIndex).Tag = "EMR_Lab.CPT_Values" Then
            strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Text
            p1 = InStr(strCTP, "(") + 1
            p2 = InStr(strCTP, ")")
            strCTP = Mid(strCTP, p1, p2 - p1)
            'strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
        End If
        
        If TreeView2.Nodes.Item(vlngIndex).Tag = "BusinessGUI.Medications" Then
            strCTP = TreeView2.Nodes.Item(vlngIndex + 1).Text
            strCTP = Left(strCTP, InStr(strCTP, ",") - 1)
            strCTP = ""
        End If
    End If
    

    
    Set fobjLib = fobjOCX.Object
    fobjLib.ConnectStringTerms = psConnect(1)
    fobjLib.ConnectStringClinical = psConnect(0)
    fobjLib.PatientID = piPatientID
    fobjLib.EncounterID = lngEnctr 'piEncounterID
    
    If strCTP <> vbNullString Then
        fobjLib.CPTCode = strCTP
    End If
    'fobjLib.ProviderID = txtProvider
    fobjLib.ComponentID = intActX_ID
    fobjLib.ReadClinicalData
    
'    intNextIndex = vlngIndex + 1  'Start of data
'
'    While TreeView2.Nodes.Item(intNextIndex).Image = 6
'        strData = TreeView2.Nodes.Item(intNextIndex).Text
'
'    Wend
    
    Exit Sub
Error_LoadOCX:
    ErrHandler _
     "frmAssessment.LoadOCX:", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume

End Sub

Private Sub Notes()
'Patient clinical data can be changed (insert/update) only if the last encounter is not closed,
'meaning that the ENCOUNTER:encntr_end_date is NULL

'Selecting a notepad icon always represents the current patient and encounter selected.

'Selecting a red diamond icon reproduces data for that encounter shown.

'Image Icons:
'1 - Closed Folder
'2 - Open Folder
'3 - NotePad
'4 - Down Arrow
'5 - Up Arrow
'6 - Black Square
'7 - Form
'8 - EMR logo
'9 - Yellow Diamond
'10 - Red Diamond

End Sub

Private Sub Display_PatientData(vstrNodeKey As String, vintNodeIndex As Integer)
'Private Sub Display_PatientData(ByVal Node As MSComctlLib.Node)

Dim sKey As String
Dim iFormID As Integer
Dim strSQL As String
Dim vRS
Dim vdata
Dim bRet As Boolean
Dim iRows As Long
Dim r As Integer
Dim sParentTab As String
Dim sChildTab As String
Dim X As Integer
Dim iParent As Integer
Dim iCount As Integer
Dim iData As Integer
Dim sClass As String
Dim sOCX As String
Dim strIntrvnCD As String
Dim sControl As String
Dim sText As String
Dim intTabID As Integer
Dim strObsDef As String
Dim booDataPresent As Boolean
Dim booSubDataPresent As Boolean
Dim iSubParent As Integer

LogFile ""
LogFile "********************** Display_PatientData ***************************"
LogFile ""

LogFile "Passed Values:  vstrNodeKey = " & vstrNodeKey & ", vintNodeIndex = " & vintNodeIndex
LogFile ""

TreeView2.Nodes.Clear

sKey = vstrNodeKey
fstrLastNodeKey = sKey
fintLastNodeIndex = vintNodeIndex

iFormID = Mid(sKey, 5)

strSQL = "Select  " & vbNewLine _
    & "a.tab_caption, " & vbNewLine _
    & "d.tab_seq,  " & vbNewLine _
    & "a.tab_id,  " & vbNewLine _
    & "child = coalesce((Select tab_caption From TAB Where tab_id = b.tab_child_id),''),  " & vbNewLine _
    & "b.tab_child_id,  " & vbNewLine _
    & "actx_class = coalesce((Select actx_name + '.' + actx_class From ACTIVE_X Where actx_id = c.actx_id),''),  " & vbNewLine _
    & "c.actx_id  " & vbNewLine _
& "From TAB a, TAB_RELATIONS b, TAB_ACTIVE_X c, FORM_TAB d  " & vbNewLine _
& "Where a.tab_id *= b.tab_parent_id and  " & vbNewLine _
    & "a.tab_id = c.tab_id And a.tab_id = d.tab_id And form_id = " & iFormID
strSQL = strSQL & vbNewLine & " Union Select  " & vbNewLine _
    & "a.tab_caption, " & vbNewLine _
    & "d.tab_seq,  " & vbNewLine _
    & "a.tab_id,  " & vbNewLine _
    & "child = coalesce((Select tab_caption From TAB Where tab_id = b.tab_child_id),''),  " & vbNewLine _
    & "b.tab_child_id,  " & vbNewLine _
    & "actx_class = coalesce((Select actx_name + '.' + actx_class From ACTIVE_X Where actx_id = c.actx_id),''),  " & vbNewLine _
    & "c.actx_id  " & vbNewLine _
& "From TAB a, TAB_RELATIONS b, TAB_ACTIVE_X c, FORM_TAB d  " & vbNewLine _
& "Where a.tab_id = b.tab_parent_id and  " & vbNewLine _
     & "b.tab_child_id *= c.tab_id And a.tab_id = d.tab_id And form_id = " & iFormID & vbNewLine _
& " Order By d.tab_seq"

'--------------------------------
LogFile ""
LogFile "Clicked on " & TreeView.Nodes.Item(vintNodeIndex).Text
LogFile "": LogFile strSQL: LogFile ""

bRet = adoSelectQuery(strSQL, psConnect(1), vRS, , iRows)

TreeView2.Nodes.Clear

'0 - tab_caption
'1 - tab_seq
'2 - tab_id (parent)
'3 - child (caption)
'4 - tab_child_id
'5 - actx_name.actx_class
'6 - actx_id

If iRows > 0 Then
    For r = LBound(vRS, 2) To UBound(vRS, 2)
    
        If vRS(5, r) <> vbNullString Then  'actx_name.actx_class
            sClass = Mid(vRS(5, r), InStr(vRS(5, r), ".") + 1)
            sOCX = Left(vRS(5, r), InStr(vRS(5, r), ".") - 1)
            sControl = vRS(5, r)
            
            Select Case sClass
                Case "CheckList"
                    fintObsCheckList = vRS(6, r)  'actx_id
                    
                    If IsNull(vRS(4, r)) Then
                        intTabID = vRS(2, r)
                    Else
                        intTabID = vRS(4, r)
                    End If
                    
                    strSQL = "Select ACTX_INTRV_CD " _
                        & "From TAB_ACTIVE_X " _
                        & "Where ACTX_ID = " & vRS(6, r) & " And TAB_ID = " & intTabID
                        
                    If Get_Values(strSQL, psConnect(1), vdata) Then
                        fintObsGroupingID = Trim(CStr(vdata(0, 0)))
                    End If
                    
                    strSQL = "Select OBS_DEFINITION From OBSERVATION_GROUPING Where OBS_GROUPING_ID = " & fintObsGroupingID
                    
                    If Get_Values(strSQL, psConnect(1), vdata) Then
                        sClass = Trim(vdata(0, 0))
                    End If
                    
                    
                    
                Case "Observations"
                    fintObs = vRS(6, r)  'actx_id
                
                Case "ValueList"
                    fintObsValueList = vRS(6, r)   'actx_id
                    
                    If IsNull(vRS(4, r)) Then
                        intTabID = vRS(2, r)
                    Else
                        intTabID = vRS(4, r)
                    End If
                    
                    strSQL = "Select ACTX_INTRV_CD " _
                        & "From TAB_ACTIVE_X " _
                        & "Where ACTX_ID = " & vRS(6, r) & " And TAB_ID = " & intTabID
                        
                    If Get_Values(strSQL, psConnect(1), vdata) Then
                        fintObsGroupingID = Trim(CStr(vdata(0, 0)))
                    End If
                    
                    strSQL = "Select OBS_DEFINITION From OBSERVATION_GROUPING Where OBS_GROUPING_ID = " & fintObsGroupingID
                    
                    If Get_Values(strSQL, psConnect(1), vdata) Then
                        sClass = Trim(vdata(0, 0))
                    End If
                
                Case "CPT_Values"
                    fintLab = vRS(6, r)   'actx_id
                    
                    If IsNull(vRS(4, r)) Then
                        intTabID = vRS(2, r)
                    Else
                        intTabID = vRS(4, r)
                    End If
                    
                    strIntrvnCD = ""
                    'Get intervention code.......................
                    'strSQL = "Select a.cpt_intrvn_code " & vbNewLine _
                        & "From LAB_RESULTS a , terms.dbo.TAB_ACTIVE_X b " & vbNewLine _
                        & "Where b.actx_intrv_cd = a.cpt_intrvn_code And " & vbNewLine _
                        & "b.tab_id = " & vRS(4, r) & " And a.patient_id = " & piPatientID
                        
                    strSQL = "Select actx_intrv_cd " & vbNewLine _
                        & "From TAB_ACTIVE_X  " & vbNewLine _
                        & "Where tab_id = " & vRS(4, r)
                       
                        
                    If Get_Values(strSQL, psConnect(1), vdata) Then
                        strIntrvnCD = Trim(CStr(vdata(0, 0)))
                    Else
                        sClass = "(no intervention applied)"
                        'sText = "(no lab data)"
                    End If
                    
                    'replace sClass variable..............
                    If strIntrvnCD <> "" Then
                        strSQL = "Select CPT_INTRVN_DEF From BT_CPT_INTRVN_LAB Where CPT_INTRVN_CODE = '" & strIntrvnCD & "'"
                        If Get_Values(strSQL, psConnect(1), vdata) Then
                            sClass = Trim(CStr(vdata(0, 0))) & " >> " & strIntrvnCD
                            'fstrLabDescr = ""
                        End If
                    End If
                    
                
                Case "Medications"   'actx_id
                    fintMeds = vRS(6, r)
               
             End Select
           
        End If
        
        
    
        If sParentTab <> vRS(0, r) Then  'First level Folder icon
            LogFile "TreeView2.Nodes.Add Relative:(not passed) , Relationship=First, Key=ParentTab" & CStr(vRS(2, r)) & "-" & r & ", Text=" & CStr(vRS(0, r)) & ", Image=2, SelectedImage:(not passed)"
            'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
            TreeView2.Nodes.Add , tvwFirst, "ParentTab" & CStr(vRS(2, r)) & "-" & r, CStr(vRS(0, r)), 2
            TreeView2.Nodes(TreeView2.Nodes.Count).ForeColor = &HC00000
            TreeView2.Nodes(TreeView2.Nodes.Count).Bold = True
             iParent = TreeView2.Nodes.Count
             X = X + 1
             iCount = iCount + 1
             booDataPresent = False
             booSubDataPresent = False
        End If
       

        If sChildTab <> vRS(3, r) Then  '2nd level (Folder or notepad icon)
            If Not IsNull(vRS(4, r)) Then
                LogFile "TreeView2.Nodes.Add Relative=" & iParent & ", Relationship=Child, Key=ParentTab" & CStr(vRS(4, r)) & "-" & r & ", Text=" & CStr(vRS(3, r)) & ", Image=2, SelectedImage:(not passed)"
                'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                TreeView2.Nodes.Add iParent, tvwChild, "ChildTab" & CStr(vRS(4, r)) & "-" & r, CStr(vRS(3, r)), 2
                TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                TreeView2.Nodes(TreeView2.Nodes.Count).ForeColor = &HC00000
                X = TreeView2.Nodes.Count
                iCount = iCount + 1
                booDataPresent = False
                'booSubDataPresent = False
            Else
                LogFile "TreeView2.Nodes.Add Relative=" & iParent & ", Relationship=Child, Key=ChildTab" & CStr(vRS(2, r)) & "-" & CStr(vRS(6, r)) & ", Text=" & sClass & ", Image=3, SelectedImage:(not passed)"
                'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                TreeView2.Nodes.Add iParent, tvwChild, "ChildTab" & CStr(vRS(2, r)) & "-" & CStr(vRS(6, r)), sClass, 3
                TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                TreeView2.Nodes.Item(TreeView2.Nodes.Count).Tag = sControl 'sOCX
                iData = Get_Clinical_Data(CInt(vRS(6, r)))
                iCount = iCount + 1 + iData
                
                If iData = 0 And chkShowData.Value And Not booDataPresent Then
                    'TreeView2.Nodes.Item(TreeView2.Nodes.Count).Parent.Index
                    TreeView2.Nodes(iParent).Expanded = False
                Else
                    TreeView2.Nodes(iParent).Expanded = True
                    booDataPresent = True
                End If
            End If
        ElseIf IsNull(vRS(4, r)) Then
                LogFile "TreeView2.Nodes.Add Relative=" & iParent & ", Relationship=Child, Key=ChildTab" & CStr(vRS(2, r)) & "-" & CStr(vRS(6, r)) & ", Text=" & sClass & ", Image=3, SelectedImage:(not passed)"
                'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                TreeView2.Nodes.Add iParent, tvwChild, "ChildTab" & CStr(vRS(2, r)) & "-" & CStr(vRS(6, r)), sClass, 3
                TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                TreeView2.Nodes.Item(TreeView2.Nodes.Count).Tag = sControl 'sOCX
                iData = Get_Clinical_Data(CInt(vRS(6, r)))
                iCount = iCount + 1 + iData
                
                If iData = 0 And chkShowData.Value And Not booDataPresent Then
                    'TreeView2.Nodes.Item(TreeView2.Nodes.Count).Parent.Index
                    TreeView2.Nodes(iParent).Expanded = False
                Else
                    TreeView2.Nodes(iParent).Expanded = True
                    booDataPresent = True
                End If
        End If
        
        If Not IsNull(vRS(4, r)) Then
           
            If vRS(6, r) <> vbNullString Then   '3rd level icon
                LogFile "TreeView2.Nodes.Add Relative=" & X & ", Relationship=Child, Key=Class" & r & "-" & CStr(vRS(6, r)) & ", Text=" & sClass & ", Image=3, SelectedImage:(not passed)"
                'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                TreeView2.Nodes.Add X, tvwChild, "Class" & r & "-" & CStr(vRS(6, r)), sClass, 3
                TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                TreeView2.Nodes.Item(TreeView2.Nodes.Count).Tag = sControl  'sOCX
                iData = Get_Clinical_Data(CInt(vRS(6, r)), CInt(vRS(4, r)))  '4th level
                iCount = iCount + 1 + iData
                
                iSubParent = TreeView2.Nodes.Item(TreeView2.Nodes.Count).Parent.Index
                
'                If TreeView2.Nodes.Item(TreeView2.Nodes.Count).Parent.Image = 2 Then
'                    booDataPresent = False
'                End If
                If iData = 0 And chkShowData.Value And Not booDataPresent Then
                    TreeView2.Nodes(iSubParent).Expanded = False
                    
                    If Not booSubDataPresent Then
                        TreeView2.Nodes(iParent).Expanded = False
                    End If
                Else
                    TreeView2.Nodes(iParent).Expanded = True
                    TreeView2.Nodes(iSubParent).Expanded = True
                    booDataPresent = True
                    booSubDataPresent = True
                End If
                
            Else 'Get second order tab nesting........................
                Get_ChildChild_Data CInt(vRS(2, r)), CInt(vRS(4, r)) 'pass tab_id & tab_child_id
            End If
        End If
        
        sParentTab = vRS(0, r)
        sChildTab = vRS(3, r)
        
    Next r
End If

If Left(TreeView.Nodes.Item(vintNodeIndex).Key, 1) = "F" Then
    pnlContent2.Caption = "Contents of " & TreeView.Nodes(vintNodeIndex).Text
Else
    pnlContent2.Caption = "Contents of "
End If

End Sub

Private Function Get_Values(vstrSQL As String, vstrConnect As String, vRS) As Long
Dim bRet As Boolean
Dim iRows As Long

On Error GoTo Error_Get_Values

    bRet = adoSelectQuery(vstrSQL, vstrConnect, vRS, , iRows)
    
    If bRet Then
        If iRows > 0 Then
            Get_Values = iRows
        End If
    End If

Exit Function

Error_Get_Values:
    ErrHandler _
     "frmAssessment.Get_Values:", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function

Private Sub Get_ChildChild_Data(vintTabID As Integer, vintTabChild As Integer)
Dim sKey As String
Dim iFormID As Integer
Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim r As Integer
Dim sParentTab As String
Dim sChildTab As String
Dim X As Integer
Dim iParent As Integer
Dim iCount As Integer
Dim iData As Integer
Dim sClass As String
Dim sOCX As String

LogFile ""
LogFile "********************** Get_ChildChild_Data ***************************"
LogFile ""

LogFile "Passed Values:  vintTabID = " & vintTabID & ", vintTabChild = " & vintTabChild
LogFile ""

strSQL = "Select  " & vbNewLine _
    & "a.tab_caption, " & vbNewLine _
    & "b.tab_seq,  " & vbNewLine _
    & "a.tab_id,  " & vbNewLine _
    & "child = coalesce((Select tab_caption From TAB Where tab_id = b.tab_child_id),''),  " & vbNewLine _
    & "b.tab_child_id,  " & vbNewLine _
    & "actx_class = coalesce((Select actx_name + '.' + actx_class From ACTIVE_X Where actx_id = c.actx_id),''),  " & vbNewLine _
    & "c.actx_id  " & vbNewLine _
& "From TAB a, TAB_RELATIONS b, TAB_ACTIVE_X c  " & vbNewLine _
& "Where a.tab_id = b.tab_parent_id and  " & vbNewLine _
     & "b.tab_child_id = c.tab_id And a.tab_id = " & vintTabChild & vbNewLine _
& " Order By b.tab_seq"


LogFile "": LogFile strSQL: LogFile ""

bRet = adoSelectQuery(strSQL, psConnect(1), vRS, , iRows)

If iRows > 0 Then
    iParent = TreeView2.Nodes.Count
    For r = LBound(vRS, 2) To UBound(vRS, 2)
    
        If vRS(5, r) <> vbNullString Then sOCX = Left(vRS(5, r), InStr(vRS(5, r), ".") - 1)
        sClass = Mid(vRS(5, r), InStr(vRS(5, r), ".") + 1)

        If sChildTab <> vRS(3, r) Then
            If Not IsNull(vRS(4, r)) Then
                'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                TreeView2.Nodes.Add iParent, tvwChild, "ChildTab" & CStr(vRS(4, r)) & "-" & r, CStr(vRS(3, r)), 2
                TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                TreeView2.Nodes(TreeView2.Nodes.Count).ForeColor = &HC00000
                X = TreeView2.Nodes.Count
                iCount = iCount + 1
            Else
                'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                TreeView2.Nodes.Add iParent, tvwChild, "ChildTab" & CStr(vRS(2, r)) & "-" & CStr(vRS(6, r)), sClass, 3
                TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                TreeView2.Nodes.Item(TreeView2.Nodes.Count).Tag = sOCX
                iData = Get_Clinical_Data(CInt(vRS(6, r)))
                iCount = iCount + 1 + iData
            End If
        ElseIf IsNull(vRS(4, r)) Then
                'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                TreeView2.Nodes.Add iParent, tvwChild, "ChildTab" & CStr(vRS(2, r)) & "-" & CStr(vRS(6, r)), sClass, 3
                TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                TreeView2.Nodes.Item(TreeView2.Nodes.Count).Tag = sOCX
                iData = Get_Clinical_Data(CInt(vRS(6, r)))
                iCount = iCount + 1 + iData
        End If
        
        If Not IsNull(vRS(4, r)) Then
            
            If vRS(6, r) <> vbNullString Then
                'TreeView2.Nodes.Add Relative, Relationship, Key, Text, Image, SelectedImage
                TreeView2.Nodes.Add X, tvwChild, "Class" & r & "-" & CStr(vRS(6, r)), sClass, 3
                TreeView2.Nodes(TreeView2.Nodes.Count - 1).Expanded = True
                TreeView2.Nodes.Item(TreeView2.Nodes.Count).Tag = sOCX
                iData = Get_Clinical_Data(CInt(vRS(6, r)))
                iCount = iCount + 1 + iData
                
                If iData = 0 And chkShowData.Value Then
                    'TreeView2.Nodes.Item(TreeView2.Nodes.Count).Parent.Index
                    TreeView2.Nodes(iParent).Expanded = False
                Else
                    TreeView2.Nodes(iParent).Expanded = True
                End If
                
            Else 'Get second order tab nesting........................
                Get_ChildChild_Data CInt(vRS(2, r)), CInt(vRS(4, r)) 'pass tab_id & tab_child_id
            End If
        End If
        
        sParentTab = vRS(0, r)
        sChildTab = vRS(3, r)
    Next r
End If



End Sub

Private Sub GetSummaryData(vstrControl As String, vvarRecSet)

'    Select Case vstrControl
'
'        Case "Follicular"
'            Follicular vvarRecSet
'
'    End Select

End Sub

Private Function ProcessSummaryData(vstrControl As String, vvarRecSet, vlngEncounterID As Long, vstrDate As String) As String
    Select Case vstrControl
    
        Case "Follicular"
            ProcessSummaryData = Sum_Follicular(vvarRecSet, vlngEncounterID, vstrDate)
    
    End Select
    
End Function

Private Function Sum_Follicular(vvarRecSet, vlngEncounterID As Long, vstrDate As String) As String
Dim c As Integer
Dim r As Integer
Static lintRow As Integer
Static lstrDate As String
Static lstrLastSum As String
    
'    If IsDate(lstrDate) Then
'        If lstrDate <> vstrDate Then
'            lintRow = 0
'        End If
'    ElseIf Not IsDate(lstrDate) Then
'        lintRow = 0
'    End If

    'If lstrLastSum = "" Then lintRow = 0: lstrDate = ""
    If lintRow = UBound(vvarRecSet, 2) Then lintRow = 0
    
    For r = LBound(vvarRecSet, 2) + 1 To UBound(vvarRecSet, 2)
    
        If vlngEncounterID = vvarRecSet(0, r) And vstrDate = vvarRecSet(1, r) And (r > lintRow Or r = UBound(vvarRecSet, 2)) Then
            '                Side                               Value                            Total
            Sum_Follicular = vvarRecSet(2, r) & ": Size = " & vvarRecSet(3, r) & ", Total = " & vvarRecSet(4, r)
            lintRow = r
            Exit For
        'ElseIf r > lintRow And lstrDate = vstrDate Then
        ElseIf r > lintRow Then
            r = lintRow
            Exit For
        End If
    
    Next r
    
    
    'If Sum_Follicular = lstrLastSum Then Sum_Follicular = ""
    lstrLastSum = Sum_Follicular
    lstrDate = vstrDate
    
    
    If r = UBound(vvarRecSet, 2) + 1 Then lintRow = 0

End Function
