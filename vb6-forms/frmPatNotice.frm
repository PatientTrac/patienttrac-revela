VERSION 5.00
Object = "{2A5B1681-9AFA-4521-B035-B9AD0C75F381}#1.0#0"; "DedicatedControls.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmPatNotice 
   Caption         =   "Practice Notice Reports"
   ClientHeight    =   6915
   ClientLeft      =   5610
   ClientTop       =   3120
   ClientWidth     =   11940
   Icon            =   "frmPatNotice.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmPatNotice.frx":094A
   ScaleHeight     =   6915
   ScaleWidth      =   11940
   Begin VB.Frame fraNotice 
      Height          =   2355
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   11865
      Begin VB.CommandButton cmdAdmin 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10650
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1950
         Width           =   1155
      End
      Begin FPSpreadADO.fpSpread grdPatNotice 
         Height          =   1815
         Left            =   90
         TabIndex        =   3
         Top             =   300
         Width           =   5535
         _Version        =   458752
         _ExtentX        =   9763
         _ExtentY        =   3201
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmPatNotice.frx":0A92
      End
      Begin VB.Frame fraAdmin 
         Caption         =   "Admin"
         Height          =   1305
         Left            =   5850
         TabIndex        =   2
         Top             =   300
         Visible         =   0   'False
         Width           =   5715
         Begin DedicatedControls.ccMultiCheck lstNotice 
            Height          =   1200
            Left            =   30
            TabIndex        =   4
            Top             =   30
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2117
            Caption         =   "Current Menu Selections"
            CaptionPosition =   1
            ForeColor       =   8388608
         End
      End
   End
   Begin FPSpreadADO.fpSpread grdPatient 
      Height          =   3975
      Left            =   30
      TabIndex        =   0
      Top             =   2340
      Width           =   11895
      _Version        =   458752
      _ExtentX        =   20981
      _ExtentY        =   7011
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmPatNotice.frx":0C90
   End
End
Attribute VB_Name = "frmPatNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim fpimage As Picture
Public piPatientID As Long
Public piEncounterID As Long
Public piProviderID As Long
Public piFacilityID As Long
Private fUpdateMode As String
Private mLocID As Integer
Private mbAdmit As Boolean
Private mbDischarge As Boolean
Private mRoomAssign As String
Private mNewPatient As Boolean
Private mstrSaveData As String
Private mRow As Integer
Public mstrSearchItem As String
Public gbooBypass As Boolean
Public mFirst As Boolean
Private mstrNotice As String
Private mstrClinDesc As String


Private Sub Get_BatchReport()
Dim i
Dim strSQL As String
Dim varRS
Dim lngrecs As Long
Dim booRet As Boolean
Dim lngdelta As Long
Dim varCurrentDate As String
Dim varPapDate As String
Dim sPatientName As String
Dim strClinicalRef As String
    
    mRow = 0
    
If mstrNotice <> vbNullString Then

      strSQL = "Select distinct a.CLINICAL_VALUE,a.Patient_ID, a.encounter_id, b.last_name, b.first_name " _
        & "FROM PATIENT_CLINICAL a, Patient b " _
        & "Where clinical_desc = '" & mstrClinDesc & "' and " _
        & "clinical_value_ref = '" & mstrNotice & "' and a.patient_id = b.patient_id " _
        & "Order by a.CLINICAL_VALUE_REF"
        
    booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
    If booRet Then
        If lngrecs Then
            For i = LBound(varRS, 2) To UBound(varRS, 2)
                If Not IsNull(varRS(0, i)) Then
                    varCurrentDate = Format(Now, "mm/dd/yyyy")
                    varPapDate = varRS(0, i)
                    lngdelta = DateDiff("yyyy", varRS(0, 0), varCurrentDate)
                    If lngdelta <> vbNullString Then
                        If lngdelta >= 1 And lngdelta < 2 Then
            With grdPatient
                    mRow = mRow + 1
                    .Row = mRow
                      .Col = 1 ' Patient ID
                      .SetText 1, .Row, varRS(1, i)
                      .CellType = CellTypeStaticText
                      
                      .Col = 2 ' Encounter ID
                      .SetText 2, .Row, varRS(2, i)
                      .CellType = CellTypeStaticText
                      
                      .Col = 3 ' Patient Name
                      .SetText 3, .Row, varRS(4, i) & " " & varRS(3, i)
                      .CellType = CellTypeStaticText
                      
                      .Col = 4 ' Last Pap Smear
                      .SetText 4, .Row, varRS(0, i)
                      .CellType = CellTypeStaticText
                    
            End With
                        End If
                    End If
                End If
            Next i
        End If
    End If
End If

End Sub

Sub AddTopBorder(Row As Long)
'Add the top row border
  
    'Select a block of cells
    grdPatient.SetCellBorder 1, Row, grdPatient.MaxCols, Row, 4, &H8000000F, CellBorderStyleSolid
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
  Exit Sub
End Sub

Private Sub grdPatient_Click(ByVal Col As Long, ByVal Row As Long)
'Show or hide the notes row when clicking the picture
    'If Col <> 1 Then Exit Sub
Dim lngPatientID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim fpnote As String
Dim varPatID

    grdPatient.MaxRows = mRow
    With grdPatient
        .Col = Col
        .Row = Row
        .BlockMode = True
        .Col = 2
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .FontBold = True
        .BackColor = 16777215
         txtSearch.Text = ""
        .BlockMode = False

    .Col = Col
    .Row = Row


        If Col = 2 Then
            grdPatient.GetText 1, .Row, varPatID
            piPatientID = varPatID
            grdPatient.GetText 2, .Row, varEnctrID
            piEncounterID = CLng(varEnctrID)
        End If
   End With
   Exit Sub
End Sub


Sub InitPatient()

    With grdPatient
        .MaxCols = 5
        .MaxRows = 1
        .ClearRange 0, 0, grdPatient.MaxCols, grdPatient.MaxRows, False
        .RowHeadersShow = False  'turn off row headers
        .EditModePermanent = True   'hide the focus rectangle
        .ScrollBars = ScrollBarsBoth    'Only show vertical scrollbars
        .GridShowHoriz = False  'turn off horizontal grid lines
        .GridShowVert = False   'turn off vertical grid lines
        .BlockMode = True
        .Col = 0
        .Row = 0
        .Col2 = .MaxCols
        .Row2 = 0
        .ColWidth(1) = 7       'set picture col width
        .ColWidth(2) = 60   'Set name, memo col width
        .ColWidth(3) = 12
        .ColWidth(4) = 14
        .ColWidth(5) = 15
        .ShadowColor = &HFFC0FF
        .FontBold = True
        .BlockMode = False
        
        'Set header text
        .SetText 1, 0, "Patient ID"
        .SetText 2, 0, "Encounter_ID"
        .SetText 3, 0, "Patient Name"
        .SetText 4, 0, "Last Pap Smear"
        'init text tips
        .TextTip = TextTipFixed
        .SetTextTipAppearance "Times New Roman", 10, True, False, &HFFFF&, &H800000

    End With
    
End Sub

Sub GetData()
'Populate Spread
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sFirstName As String
Dim sLastName As String
Dim sPatientName As String
Dim strSQL As String
Dim varRS
Dim lngrecs As Long
Dim booRet
Dim i%
Dim r%
Dim varBirth
Dim lngPatExtRef As Long
Dim lngPatID
Dim vartemp
Dim lngRow As Long
Dim lngCol As Long
Dim varFilterTime As Long
Dim varCurrentDate
Dim varTime

    strSQL = "SELECT distinct a.First_Name, a.Last_Name,a.Patient_ID,  a.Patient_Ext_Ref " _
            & "FROM PATIENT a, encounter b, providers c " _
            & "Where a.Patient_ID = b.Patient_ID " _
            & "And b.provider_id = c.provider_id And c.Active = 1 " _
            & "Order by a.Last_Name"
            
            booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
    If booRet Then
        If lngrecs > 0 Then
            lblCount.Caption = lngrecs
            For i = LBound(varRS, 2) To UBound(varRS, 2)
                lstPatient.AddItem varRS(1, i) & ", " & varRS(0, i)
                lstPatient.ItemData(lstPatient.NewIndex) = varRS(2, i)
                lstPatient.ListIndex = i
            Next i
        End If
    End If

    If piPatientID <> 0 Then
        For i = 0 To lstPatient.ListCount - 1
            If lstPatient.ItemData(lstPatient.ListIndex) = piPatientID Then
                lstPatient.Selected(i) = True
            End If
        Next i
    End If
    

  Exit Sub
Error_GetData:
        Screen.MousePointer = vbDefault
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        Exit Sub
        Resume
       
End Sub


Private Sub Form_Load()


    lstNotice.AddItem "Pap Smear", True
    
End Sub


Private Sub grdPatient_DblClick(ByVal Col As Long, ByVal Row As Long)
'Show or hide the notes row when clicking the picture
    'If Col <> 1 Then Exit Sub
Dim lngPatientID As Long
Dim lngRow As Long
Dim lngCol As Long
Dim fpnote As String
Dim varPatID
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strPrintItem As String
Dim strRequest As String
Dim strUSER_ID As String
Dim intPATIENT_ID As Integer

    grdPatient.MaxRows = mRow
    With grdPatient
        .Col = Col
        .Row = Row
        .BlockMode = True
        .Col = 2
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .FontBold = True
        .BackColor = 16777215
         txtSearch.Text = ""
        .BlockMode = False

    .Col = Col
    .Row = Row


        If Col = 2 Then
            grdPatient.GetText 1, .Row, varPatID
            piPatientID = varPatID
            grdPatient.GetText 2, .Row, varEnctrID
            piEncounterID = CLng(varEnctrID)
        End If
   End With
   frmMDIReport.fmstrfile = "C:\Program Files\Firstnotebook\PatientTrac EMR\Reports\PapSmear_1100"
   frmMDIReport.cmdRun_Click
   frmMDIReport.Show
   frmMDIReport.cmdPrint_Click
   cn.open psConnect(0)
   strPrintItem = mstrNotice
   strRequest = "Practice"
   strUSER_ID = PsUserID
   intPATIENT_ID = piPatientID
   If frmMDIReport.mbooPrint Then
        cn.I_PRACTICE_PRINTLOG UCase(strPrintItem), UCase(strRequest), UCase(strUSER_ID), intPATIENT_ID
   End If

    Set rs = Nothing
    
    cn.Close
    Set cn = Nothing
   
   Exit Sub
    
End Sub

Private Sub grdPatNotice_Click(ByVal Col As Long, ByVal Row As Long)
Dim varItem

        With grdPatNotice
            If Col = 1 Then
                .GetText 2, Row, varItem
                mstrNotice = varItem
                Get_BatchReport
            End If
        End With
    Exit Sub
End Sub


Private Sub Init_PatNotice()


    With grdPatNotice
        .MaxCols = 3
        .MaxRows = 1
        .ClearRange 0, 0, grdPatNotice.MaxCols, grdPatient.MaxRows, False
        .RowHeadersShow = False  'turn off row headers
        .EditModePermanent = True   'hide the focus rectangle
        .ScrollBars = ScrollBarsBoth    'Only show vertical scrollbars
        .GridShowHoriz = False  'turn off horizontal grid lines
        .GridShowVert = False   'turn off vertical grid lines
        .BlockMode = True
        .Col = 0
        .Row = 0
        .Col2 = .MaxCols
        .Row2 = 0
        .ColWidth(1) = 7       'set picture col width
        .ColWidth(2) = 60   'Set name, memo col width
        .ColWidth(3) = 12
        .ColWidth(4) = 14
        .ColWidth(5) = 15
        .ShadowColor = &HFFC0FF
        .FontBold = True
        .BlockMode = False
        
        'Set header text
        .SetText 1, 0, " "
        .SetText 2, 0, "Pt Notice ID"
        .SetText 3, 0, "PATIENT_NOTICE"
        'init text tips
        .TextTip = TextTipFixed
        .SetTextTipAppearance "Times New Roman", 10, True, False, &HFFFF&, &H800000

    End With

End Sub

Private Function Get_ObjectName()

    If mstrNotice <> vbNullString Then
        If mstrNotice = "Pap Smear" Then
            Get_ObjectName = ccdlastpap
            mstrClinDesc = "FN_ObGyn.NDHReaVital"
        End If
    End If

End Function


Private Sub Load_grdPatNotice()
Dim i As Integer

    mRow = 0
    With grdPatNotice
        .ClearRange 0, 0, grdPatNotice.MaxCols, grdPatNotice.MaxRows, False
        For i = 0 To lstNotice.ListCount - 1
            mRow = mRow + 1
            .Row = mRow
            If lstNotice.Selected(i) Then
                varNotice = lstNotice.List(i)
                .SetText 2, .Row, varNotice
            End If
        Next i
            
            .BlockMode = True
                .Col = 0
                .Row = 1
                .Col2 = .MaxCols
                .Row2 = .MaxRows
                .CellType = CellTypeStaticText
                .FontBold = True
                .FontName = "Times New Roman"
            .BlockMode = False

        End With

    Exit Sub
End Sub
