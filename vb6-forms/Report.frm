VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{2A5B1681-9AFA-4521-B035-B9AD0C75F381}#1.0#0"; "DedicatedControls.ocx"
Object = "{890AC7CB-5871-4973-95D1-5176E8DFFCFE}#1.0#0"; "DedicatedControlsST.ocx"
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   9600
   ClientLeft      =   6150
   ClientTop       =   2910
   ClientWidth     =   12885
   Icon            =   "Report.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   12885
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1800
      Top             =   8580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DedicatedControlsST.ccAnnotation ccaImage 
      Height          =   3530
      Left            =   3300
      TabIndex        =   4
      Top             =   5550
      Visible         =   0   'False
      Width           =   510
      _ExtentX        =   5715
      _ExtentY        =   5001
   End
   Begin VB.PictureBox picSignature 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7740
      ScaleHeight     =   315
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   3615
   End
   Begin DedicatedControls.ccApprove ccaSignature 
      Height          =   495
      Left            =   2610
      TabIndex        =   2
      Top             =   8600
      Visible         =   0   'False
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   873
      Enabled         =   0   'False
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
   Begin FPSpreadADO.fpSpread grdReport 
      Height          =   8925
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   12645
      _Version        =   458752
      _ExtentX        =   22304
      _ExtentY        =   15743
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   13619374
      MaxCols         =   10
      MaxRows         =   0
      SpreadDesigner  =   "Report.frx":0442
   End
   Begin VB.Label lblHidden 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "hidden"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   -60
      Visible         =   0   'False
      Width           =   885
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFile_ReRun 
         Caption         =   "Re-run"
      End
      Begin VB.Menu mnuFile_ShowGrid 
         Caption         =   "Show Grid"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuFile_Export 
         Caption         =   "Export To Excel..."
      End
      Begin VB.Menu mnuFile_sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fclsPatient As PatientInfo
Private fudtReferral As ReferralInfo
Private fclsEncounter As EncounterInfo
Private fudtAudit() As BooleanAudit
Private fudtGridHeader() As GridHeader
Private fintRowInserts As Integer
Private fbooRun As Boolean
Public fstrTitle As String
Private fstrSavePath As String
Private fstrIncludedObject() As String
Private fstrExcludedObject() As String
Private fbooSkip As Boolean
Private fbooRemRowNoData As Boolean
Private flngRowHidden() As Long
Private fudtRowInsert() As RowInsert
Private fintInsertedRows  As Integer
'Private piPatientID As Long
'Private piEncounterID As Long
Private Type RowInsert
    StartRow As Integer
    InsertedRows As Integer
End Type

Private Type PatientInfo
    Name As String
    DOB As String
    Gender As String
    Address As String
    CityStateZip As String
    Phone As String
    Email As String
    SS As String
    ExtID As String
    IntId As Long
    FCC As String
    Race As String
    Age As String
End Type

Private Type ReferralInfo
    Practice As String
    Physician As String
    Address As String
    CityStateZip As String
    Phone As String
    CellPhone As String
    FAX As String
    Email As String
    UPIN As String
    License As String
    Specialty As String
End Type

Private Type EncounterInfo
    StartDate As String
    EndDate As String
    VisitReason As String
    Facility As String
    Provider As String
    ServicePlace As String
    ServicePlaceCode As String
    EncounterType As String
End Type


Private Type BooleanAudit
    OCX As String
    Control As String
    Object As String
    DataPresent As Boolean
    ShowWhen As Boolean
End Type

Private Type GridHeader
    StartCol As Long
    EndCol As Long
    Row As Long
End Type

Private Sub cmdRun_Click()
    grdReport.MaxCols = 5
    grdReport.MaxRows = 0
    fbooRun = True
    Get_Patient piPatientID
    Get_Referral
    Get_EncounterData
    Open_Script
    fbooRun = False
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbHourglass
    Me.Caption = "Report > " & fstrTitle
    'grdReport.MaxCols = 5
    'grdReport.MaxRows = 0
        
        
    Get_Patient piPatientID
    Get_Referral
    Get_EncounterData

    Open_Script
    'CenterControl grdReport, Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    centerForm Me
    Me.Top = Me.Top + 150
    grdReport.EditModePermanent = True
    grdReport.ProcessTab = True
    grdReport.BackColorStyle = BackColorStyleUnderGrid

End Sub



Private Sub Open_Script()
Dim intNum As Integer
Dim strLine As String

    'If fbooRun Then pstrReportScript = App.Path & "\fpt_TestReport.txt"
    
    intNum = FreeFile
    Erase fudtAudit
    Erase fudtGridHeader
    Erase fudtRowInsert
    fintRowInserts = 0
    fintInsertedRows = 0
    Erase flngRowHidden
    
    Open pstrReportScript For Input As intNum
        Do While Not EOF(intNum)
            Line Input #intNum, strLine
            
            If strLine <> vbNullString And Left(strLine, 1) = "<" Then
                'Debug.Print strLine
                Execute strLine
            End If
        
        Loop
    Close intNum
    
    Remove_ResidualTags

End Sub

Private Sub Execute(vstrLine As String)
Dim strItem() As String
Dim lngCol As Long
Dim lngRow As Long
Dim lngCol2 As Long
Dim lngRow2 As Long
Dim strDate As String
Dim strData As String
Dim intSide As Integer
Dim intStyle As Integer
Dim strEncodedSignature As String
Dim booInvalid As Boolean
Dim strRefData As String
Dim varImage
Dim strPath As String
Dim booData As Boolean
Dim varRS
Dim c%, r%, i%, k%, g%, n1%, n2%, n3%, n4%
Dim intSeq As Integer
Dim strObjType As String
Dim booBooleanType As Boolean
Dim strRowHeader() As String
Dim strColHeader() As String
Dim strColHeaderWidth() As String
Dim intLastCol As Integer
Dim strExcludedObject() As String
Dim booNotEmptyRow As Boolean
Dim booSkip As Boolean
Dim booMissing As Boolean
Dim strType As String
Dim strFontType() As String
Dim strComponentVal As String
Dim strObject As String
Dim booCase As Boolean
Dim strOCX As String
Dim strControl As String
Dim intPos1 As Integer
Dim intPos2 As Integer
Dim booNoteData As Boolean
Dim varNotes
Dim intSpan As Integer
Dim intTotSpan As Integer
Dim booDataPresent As Boolean

'Dynamic Row height..............
Dim varSpan(3)
Dim X%, Y%, z%
Dim intSpanWidth As Integer
Dim booShowData As Boolean
Dim intRowInserts As Integer
Dim booRowAdded As Boolean
Dim intPrevSeq As Integer
Dim intColWidth As Integer
Dim intInsertedRows As Integer
Dim varCell
Dim boogrdSSData As Boolean


    On Error GoTo Error_Execute
    
    strItem() = Split(vstrLine, "~")
 '   Debug.Print vstrLine
    'If IsNumeric(strItem(2)) Then
        'If strItem(2) = 25 Then
           'Debug.Print
'            grdReport.TopRow = strItem(2) - 15
       ' End If
    'End If
'
    booNoteData = False
    
    Select Case strItem(0)
    
        Case "<Alignment>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            lngCol2 = strItem(3)
            lngRow2 = strItem(4)
            intInsertedRows = Get_InsertRowCount(lngRow)
            Set_Block lngCol, lngRow + intInsertedRows, lngCol2, lngRow2 + intInsertedRows
            
            If strItem(5) = "HorizAlignLeft" Then
                grdReport.TypeHAlign = TypeHAlignLeft
            ElseIf strItem(5) = "HorizAlignCenter" Then
                grdReport.TypeHAlign = TypeHAlignCenter
            ElseIf strItem(5) = "HorizAlignRight" Then
                grdReport.TypeHAlign = TypeHAlignRight
            ElseIf strItem(5) = "VertAlignTop" Then
                grdReport.TypeVAlign = TypeVAlignTop
            ElseIf strItem(5) = "VertAlignCenter" Then
                grdReport.TypeVAlign = TypeVAlignCenter
            ElseIf strItem(5) = "VertAlignBottom" Then
                grdReport.TypeVAlign = TypeVAlignBottom
            End If
            
            grdReport.BlockMode = False
            
        Case "<BackColor>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            lngCol2 = strItem(3)
            lngRow2 = strItem(4)
            intInsertedRows = Get_InsertRowCount(lngRow)
            Set_Block lngCol, lngRow + intInsertedRows, lngCol2, lngRow2 + intInsertedRows
            grdReport.BackColor = strItem(5)
            grdReport.BlockMode = False
            
        Case "<Border>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            lngCol2 = strItem(3)
            lngRow2 = strItem(4)
            
            Select Case strItem(5)  'Sides
                Case "None"
                    intSide = 0
                Case "Left"
                    intSide = 1
                Case "Right"
                    intSide = 2
                Case "Top"
                    intSide = 4
                Case "Bottom"
                    intSide = 8
                Case "All"
                    intSide = 16
            End Select
            
            Select Case strItem(7)  'Style
                Case "None"
                    intStyle = 0
                Case "Solid"
                    intStyle = 1
                Case "Dash"
                    intStyle = 13
            End Select
            
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetCellBorder lngCol, lngRow + intInsertedRows, lngCol2, lngRow2 + intInsertedRows, intSide, strItem(6), intStyle
        
        Case "<BorderInterior>"
            
        Case "<CellSpan>"
            lngCol = strItem(1) 'Start Col
            lngRow = strItem(2) 'Start Row
            lngCol2 = strItem(3) 'Quantity
            lngRow2 = strItem(4) 'Quantity
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.AddCellSpan lngCol, lngRow + intInsertedRows, lngCol2, lngRow2 '+ intInsertedRows
            
        Case "<ColWidth>"
            lngCol = strItem(1)
            grdReport.ColWidth(lngCol) = strItem(2)
            
        Case "<ColHeaders>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            strColHeader() = Split(strItem(3), ",")
            'intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.Row = lngRow '+ intInsertedRows
            
            For i = LBound(strColHeader) To UBound(strColHeader)
                grdReport.Col = lngCol + i
                grdReport.SetText lngCol + i, lngRow, CVar(strColHeader(i))
                grdReport.TypeHAlign = TypeHAlignCenter
                grdReport.FontBold = True
            Next i
            
            ReDim Preserve fudtGridHeader(UBound(fudtGridHeader) + 1)
            fudtGridHeader(UBound(fudtGridHeader)).StartCol = lngCol
            fudtGridHeader(UBound(fudtGridHeader)).EndCol = lngCol + i
            fudtGridHeader(UBound(fudtGridHeader)).Row = lngRow
            
        Case "<ColHeaderWidths>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            strColHeaderWidth() = Split(strItem(3), ",")
            
            For i = LBound(strColHeaderWidth) To UBound(strColHeaderWidth)
                grdReport.ColWidth(lngCol + i) = CInt(strColHeaderWidth(i))
            Next i
        
        Case "<Date>", "<DateTime>"
            lngCol = strItem(1) 'Start Col
            lngRow = strItem(2) 'Start Row
            strDate = Format(Now, strItem(3))
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(strDate)
            
        Case "<IncludedObjects>"
            lngCol = strItem(1) 'Start Col
            lngRow = strItem(2) 'Start Row
            
            If InStr(strItem(3), ",") > 0 Then
                fstrIncludedObject() = Split(strItem(3), ",")
            Else
                ReDim fstrIncludedObject(0)
                fstrIncludedObject(0) = strItem(3)
            End If
            
        Case "<EncounterStartDate>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.StartDate)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.StartDate = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterEndDate>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.EndDate)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.EndDate = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterVisitReason>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.VisitReason)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.VisitReason = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterFacility>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.Facility)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.Facility = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterProvider>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.Provider)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.Provider = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterServicePlace>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.ServicePlace)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.ServicePlace = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterServicePlaceCode>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.ServicePlaceCode)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.ServicePlaceCode = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<EncounterType>"
            lngCol = strItem(1): lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsEncounter.EncounterType)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsEncounter.EncounterType = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
        
        Case "<ExcludedObjects>"
            lngCol = strItem(1) 'Start Col
            lngRow = strItem(2) 'Start Row
            
            If InStr(strItem(3), ",") > 0 Then
                fstrExcludedObject() = Split(strItem(3), ",")
            Else
                If strItem(3) = "" Then strItem(3) = "|"
                ReDim fstrExcludedObject(0)
                fstrExcludedObject(0) = strItem(3)
            End If
             
        Case "<FetchData>"  'Returns one value for a specific object on a control
            lngCol = strItem(1)
            lngRow = strItem(2)
'            If lngRow > 40 And lngRow < 44 Then
'                Debug.Print
'            End If
            intInsertedRows = Get_InsertRowCount(lngRow)
            If strItem(4) = "PATIENT_CLINICAL" Then
                '                             OCX       Control      objRef        Sequence
                strData = Get_ClinicalValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
            
                'dnn   6/7/2004.....if vbnullstring try PATIENT_NOTES table............
                If strData = "" Then
                    strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                    If strData <> "" Then booNoteData = True
                End If
                
                'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                If strData = "" Then
                    strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                End If
                
            ElseIf strItem(4) = "PATIENT_NOTES" Then
                strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                If strData <> "" Then booNoteData = True
            ElseIf strItem(4) = "PATIENT_PROFILER" Then
                strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
            ElseIf strItem(4) = "PATIENT_PROFILER_NOTES" Then
                strData = Get_ProfilerNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                If strData <> "" Then booNoteData = True
            End If
            
            strData = Trim(strData)
            grdReport.Col = lngCol
            grdReport.Row = lngRow + intInsertedRows
            'Reproductive Note exception
            If strItem(7) <> "cclNote" Then
            'dnn 6/6/2004  strip carriage returns..(Accrued data types)..........
                If strData <> "" And strItem(7) <> "cctSummary" Or strItem(7) <> "cctSummary2" Or strItem(7) <> "cctSummary3" Then
                    If InStr(strData, vbNewLine) > 0 Then
                        strData = Replace(strData, vbNewLine, " ")
                            If InStr(strData, ";") > 0 Then
                                strData = Replace(strData, ";", ", ")
                            End If
                    End If
                End If
            End If
            grdReport.GetText lngCol, lngRow + intInsertedRows, varCell
            
            If varCell = vbNullString Then
                Select Case strItem(3)
                    Case "StaticText"
                        grdReport.CellType = CellTypeStaticText
                        grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(strData)
                        
                    Case "CheckBox"
                        grdReport.CellType = CellTypeCheckBox
                        If strData = "True" Then strData = "1" Else strData = "0"
                        grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(strData)
                        
                        
                    Case "OptionButton"
                        grdReport.CellType = CellTypeStaticText
                        grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(strData)
                    
                End Select
            End If
            
            grdReport.GetText lngCol, lngRow + intInsertedRows, varCell
            
            If varCell = vbNullString Then
                If HideRow(lngRow + intInsertedRows) And strData = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            End If
            
        Case "<FetchBooleanData>" 'Outputs text based on a checkbox or option button True or False case.................
            lngCol = strItem(1)
            lngRow = strItem(2)
            If strItem(4) = "PATIENT_CLINICAL" Then
                '                             OCX       Control      objRef        Sequence
                strData = Get_ClinicalValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
            
                'dnn   6/7/2004.....if vbnullstring try PATIENT_NOTES table............
                If strData = "" Then
                    strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                    If strData <> "" Then booNoteData = True
                End If
                
                'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                If strData = "" Then
                    strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                End If
            
            ElseIf strItem(4) = "PATIENT_NOTES" Then
                strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                If strData <> "" Then booNoteData = True
            ElseIf strItem(4) = "PATIENT_PROFILER" Then
                strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
            ElseIf strItem(4) = "PATIENT_PROFILER_NOTES" Then
                strData = Get_ProfilerNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                If strData <> "" Then booNoteData = True
            End If
            
            grdReport.Col = lngCol
            grdReport.Row = lngRow
            grdReport.CellType = CellTypeStaticText
            
            Select Case Trim(strData)

                Case "True", "1", "-1", "yes", "Yes", "YES", "y", "Y", "A"
                    booCase = True
                    
                Case "False", "", "no", "N", "NO", "I", "n", "0", "No"
                    booCase = False
                    
                Case Else
                    booCase = False
                
            End Select
            
            If booCase = CBool(strItem(9)) And strData = "True" Then
                grdReport.SetText lngCol, lngRow, CVar(strItem(3))
            End If
            
            'dnn  6/7/2004...........6/24/2004
            If strData <> "False" And strData <> "" And strItem(9) = "True" Then
                grdReport.SetText lngCol, lngRow, CVar(strItem(3))
            ElseIf strData <> "" Then
            
                'dnn 6/6/2004  strip carriage returns..(Accrued data types)..........
                If strData <> "" Then
                    If InStr(strData, vbNewLine) > 0 Then
                        'get data between carriage returns..........
                        intPos1 = InStr(strData, vbNewLine) + 2
                        intPos2 = InStr(intPos1, strData, vbNewLine)
                        strData = Trim(Mid(strData, intPos1 + 1, intPos2 - intPos1 - 1))
                        
                        If InStr(strData, ";") > 0 Then
                            strData = Replace(strData, ";", ", ")
                        End If
                    End If
                End If
                
                If strData <> "False" And strData <> "True" Then
                    grdReport.SetText lngCol, lngRow, CVar(strData)
                    
                End If
            End If
            
            
            '6/24/2004................
            If strData = "" Or strData = "False" And strItem(9) = "False" Then
                grdReport.SetText lngCol, lngRow, ""
                
            End If
            
            'If HideRow(0) > 0 Then strData
                'If HideRow(lngRow) And strItem(3) = "" Then grdReport.RowHeight(lngRow) = 0
                If HideRow(lngRow) And Not booCase Then
                    grdReport.ROWHEIGHT(lngRow) = 0
                Else
                    grdReport.ROWHEIGHT(lngRow) = 9
                    
                End If
            'End If
            
            
        Case "<FetchBooleanComp>" 'Outputs component value based on associated checkbox or option button True or False case.................
            lngCol = strItem(1)
            lngRow = strItem(2)
            
           ' If lngRow = 24 Then
                'Stop
            'End If
            'If booCase = False Then
            If strItem(4) = "PATIENT_CLINICAL" Then
                '                             OCX       Control      objRef        Sequence
                strData = Get_ClinicalValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
            
                'dnn   6/7/2004.....if vbnullstring try PATIENT_NOTES table............
                If strData = "" Then
                    strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                    If strData <> "" Then booNoteData = True
                End If
                
                'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                If strData = "" Then
                    strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                End If
                
            ElseIf strItem(4) = "PATIENT_NOTES" Then
                strData = Get_ClinicalNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID, piEncounterID)
                If strData <> "" Then booNoteData = True
            ElseIf strItem(4) = "PATIENT_PROFILER" Then
                strData = Get_ProfilerValue(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
            ElseIf strItem(4) = "PATIENT_PROFILER_NOTES" Then
                strData = Get_ProfilerNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                If strData <> "" Then booNoteData = True
            End If
            
            grdReport.Col = lngCol
            grdReport.Row = lngRow
            grdReport.CellType = CellTypeStaticText
            
            strObject = Mid(strItem(3), InStr(strItem(3), ":") + 1)
           ' End If
            Select Case Trim(strData)

                Case "True", "1", "-1", "yes", "Yes", "YES", "y", "Y", "A"
                    booCase = True
                    
                Case "False", "", "no", "N", "NO", "I", "n", "0", "No"
                    booCase = False
                
                Case Else
                    booCase = False
                    
            End Select
            
            If booCase = CBool(strItem(9)) Then
                strObjType = strItem(10)
                strOCX = Left(strItem(3), InStr(strItem(3), ".") - 1)
                strControl = Mid(strItem(3), InStr(strItem(3), ".") + 1)
                strControl = Left(strControl, InStr(strControl, ":") - 1)
                
                If strObjType = "ccList" Or strObjType = "ccListGrid" Or strObjType = "ccMultiCheck" Or strObjType = "ListBox" Then
                    fbooSkip = False
                Else
                    If strItem(4) = "PATIENT_CLINICAL" Then
                        '                                     OCX       Control      objRef        Sequence
                        'strComponentVal = Get_ClinicalValue(strItem(5), strItem(6), strObject, CInt(strItem(8)), piPatientID, piEncounterID)
                        strComponentVal = Get_ClinicalValue(strOCX, strControl, strObject, CInt(strItem(8)), piPatientID, piEncounterID)
                        
                        'if vbnullstring try PATIENT_NOTES table............
                        If strComponentVal = "" Then
                            'dnn 6/7/2004............
                            'strComponentVal = Get_ClinicalNote(strItem(5), strItem(6), strObject, CInt(strItem(8)), piPatientID, piEncounterID)
                            strComponentVal = Get_ClinicalNote(strOCX, strControl, strObject, CInt(strItem(8)), piPatientID, piEncounterID)
                            If strComponentVal <> "" Then booNoteData = True
                        End If
                        
                        'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                        If strComponentVal = "" Then
                            strComponentVal = Get_ProfilerValue(strOCX, strControl, strObject, CInt(strItem(8)), piPatientID)
                        End If
                    
                    ElseIf strItem(4) = "PATIENT_NOTES" Then
                        strComponentVal = Get_ClinicalNote(strItem(5), strItem(6), strObject, CInt(strItem(8)), piPatientID, piEncounterID)
                        If strData <> "" Then booNoteData = True
                    ElseIf strItem(4) = "PATIENT_PROFILER" Then
                        strComponentVal = Get_ProfilerValue(strItem(5), strItem(6), strObject, CInt(strItem(8)), piPatientID)
                    ElseIf strItem(4) = "PATIENT_PROFILER_NOTES" Then
                        strComponentVal = Get_ProfilerNote(strItem(5), strItem(6), strItem(7), CInt(strItem(8)), piPatientID)
                        If strData <> "" Then booNoteData = True
                    End If
                
                End If
                If strComponentVal <> "" Then
                    grdReport.SetText lngCol, lngRow, CVar(strComponentVal)
                End If
            Else
                'grdReport.SetText lngCol, lngRow, ""
            End If
            
            ReDim Preserve fudtAudit(UBound(fudtAudit) + 1)
            fudtAudit(UBound(fudtAudit)).OCX = Left(strItem(3), InStr(strItem(3), ".") - 1)
            fudtAudit(UBound(fudtAudit)).Control = Mid(strItem(3), InStr(strItem(3), ".") + 1, (InStr(strItem(3), ":") - InStr(strItem(3), ".") - 1))
            fudtAudit(UBound(fudtAudit)).Object = Mid(strItem(3), InStr(strItem(3), ":") + 1)
            fudtAudit(UBound(fudtAudit)).DataPresent = booCase
            fudtAudit(UBound(fudtAudit)).ShowWhen = strItem(9)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow) And strComponentVal = "" Then grdReport.ROWHEIGHT(lngRow) = 0
            'End If
        
            
                    
        Case "<FetchAnnotation>"

            ccaImage.Image = strItem(7)
            Get_Image ccaImage, strItem(3), strItem(4), strItem(5), CInt(strItem(6)), piPatientID, piEncounterID
            
            'strPath = Get_FilePath(strItem(7))
            strPath = App.Path & "\Images"
            If Not FolderExists(strPath) Then
                MkDir strPath
            End If
            
            ccaImage.Object.SaveFile App.Path & "\Images\image.bmp"
            
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            
            grdReport.Col = lngCol
            grdReport.Row = lngRow + intInsertedRows
            grdReport.CellType = CellTypePicture
            grdReport.TypePictCenter = True
                
            If FileExists(App.Path & "\Images\image.bmp") Then
                grdReport.TypePictPicture = LoadPicture(App.Path & "\Images\image.bmp")
            End If
            
         Case "<FetchLabData>"
        
        
        Case "<FetchList>"
            Select Case strItem(3) 'Table
                Case "PATIENT_CLINICAL"
                    '                            OCX        Control     objRef
                    booData = Get_ClinicalList(strItem(4), strItem(5), strItem(6), varRS, piPatientID, piEncounterID)
                
                    'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                    If Not booData Then
                        '                                   OCX        Control     objRef
                        booData = Get_ProfilerClinicalList(strItem(4), strItem(5), strItem(6), varRS, piPatientID)
                    End If
                
            End Select
            
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            
            If booData Then
                i = 0
                
                'If frmFarpointReports.chkConcatList.Value = 1 Then
                    'For r = LBound(varRS, 2) To UBound(varRS, 2)
                        'If InStr(varRS(0, r), vbNewLine) = 0 Then
                            'strData = strData & varRS(0, r) & ", "
                        'End If
                   'Next r
                    
                    For r = LBound(varRS, 2) To UBound(varRS, 2)
                        If InStr(varRS(0, r), vbNewLine) = 0 Then
                            strData = strData & varRS(0, r) & ", "
                        Else
                            strData = varRS(0, r)
                        End If
                    Next r
                    
                    
                    If Right(strData, 2) = ", " Then
                        strData = Left(strData, Len(strData) - 2)
                    End If
                    
                    strData = Replace(strData, vbNewLine, " ,")
                    
                    grdReport.SetText lngCol, lngRow + intInsertedRows, strData
                    'grdReport.AddCellSpan lngCol, lngRow, strItem(7), 1
                            

            Else
                'grdReport.SetText lngCol, lngRow, ""
            End If
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow) And strData = "" Then grdReport.ROWHEIGHT(lngRow) = 0
            'End If
            
            
        Case "<FetchRecordSet>"
            booShowData = True
            'Check Audit trail....................
            For X = LBound(fudtAudit) To UBound(fudtAudit)
            
                If fudtAudit(X).OCX = strItem(4) _
                    And fudtAudit(X).Control = strItem(5) _
                    And fudtAudit(X).Object = strItem(6) Then
                    'Debug.Print
                 
                    If fudtAudit(X).ShowWhen Then
                       'Show grid when there is data
                        If fudtAudit(X).DataPresent Then
                            booShowData = True
                        Else
                            booShowData = False
                        End If
                    Else
                        'Show grid when there is no data
                        If fudtAudit(X).DataPresent Then
                            booShowData = False
                        Else
                            booShowData = True
                        End If
                    End If
                    Exit For
                End If
            Next X
            
            If Not booShowData Then
                'Destroy grid header..................
                lngCol = strItem(1)
                lngRow = strItem(2)
                'grdReport.RowHeight(lngRow - 1) = 0
                For X = LBound(fudtGridHeader) To UBound(fudtGridHeader)
                    If fudtGridHeader(X).StartCol = lngCol And fudtGridHeader(X).Row = lngRow - 1 Then
                        For Y = fudtGridHeader(X).StartCol To fudtGridHeader(X).EndCol
                            grdReport.SetText Y, lngRow - 1 + fintRowInserts, ""
                        Next Y
                    End If
                Next X
                
            End If
            
            If Not fbooSkip And booShowData Then
                Select Case strItem(3) 'Table
                    Case "PATIENT_CLINICAL"
                    
                        If strItem(6) = "grdSS" Then
                            If strItem(5) = "Medication" Then
                                booData = Get_MedicationData(varRS, piPatientID, piEncounterID, piProviderID)
                            ElseIf strItem(5) = "Lab_Results" Then
                                booData = Get_Lab_ResultsData(piPatientID, piEncounterID, piProviderID)
                                'Set varRS, pudtLabs returns multiple record sets ..............
                                lngCol = strItem(1): lngRow = strItem(2)
                                For n1 = LBound(pudtLabs, 1) To UBound(pudtLabs, 1) 'Encounters
                                    For n2 = LBound(pudtLabs, 2) To UBound(pudtLabs, 2) 'Orders
                                        For n4 = LBound(pudtLabs, 4) To UBound(pudtLabs, 4) 'Lab Rows
                                            ReDim varRS(UBound(pudtLabs, 3), UBound(pudtLabs, 4))
                                            For n3 = LBound(pudtLabs, 3) To UBound(pudtLabs, 3) 'Lab Cols
                                                varCell = CVar(pudtLabs(n1, n2, n3, n4).Value)
                                                varRS(n3, n4) = varCell
                                            Next n3
                                        Next n4
                                        GoSub ProcessRS
                                    Next n2
                                Next n1
                                i = i - 1
                            End If
                        Else
                            '                            OCX        Control
                            booData = Get_ClinicalData(strItem(4), strItem(5), varRS, piPatientID, piEncounterID)
                        End If
                        
                        'dnn   6/7/2004.....if vbnullstring try PATIENT_Profiler table............
                        If Not booData Then
                            '                                    OCX        Control
                            booData = Get_ProfilerClinicalData(strItem(4), strItem(5), varRS, piPatientID)
                        End If
                        
                End Select
                
                lngCol = strItem(1)
                lngRow = strItem(2)
                
                booRowAdded = False
                If booData Then
                    If strItem(5) <> "Lab_Results" Then
                    intLastCol = strItem(8)
                    'grdReport.InsertRows lngRow + 1, 1
                    i = 0
                    
                    'If fintRowInserts > 0 Then
                    '    fintRowInserts = fintRowInserts - 1
                    'Else
                    '    intRowInserts = fintRowInserts
                    'End If
                    intRowInserts = fintRowInserts
                    boogrdSSData = False
                    
                    'If strItem(5) <> "Lab_Results" Then
                    For r = LBound(varRS, 2) To UBound(varRS, 2)
                    
                        intSpan = 0
                        booSkip = False
                        For k = LBound(fstrExcludedObject) To UBound(fstrExcludedObject)
                            If InStr(varRS(0, r), fstrExcludedObject(k)) > 0 Then
                                booSkip = True
                                Exit For
                            End If
                        Next k
                                                
                        'Avoid accrued values..................
                        If InStr(varRS(2, r), vbNewLine) = 0 And Not booSkip Then
                            
                            If strItem(6) = "grdSS" And booData Then
                                If strItem(5) = "Medication" Then
                                    LoadRow_Medication grdReport, varRS, r, lngCol, lngRow + fintRowInserts, fstrIncludedObject(), fstrExcludedObject()
                                'ElseIf strItem(5) = "Lab_Results" Then
                                '    LoadRow_LabResults grdReport, varRS, r, lngCol, lngRow + fintRowInserts, fstrIncludedObject(), fstrExcludedObject()
                                End If
                                    i = i + 1
                                    
                                    If i > 1 Then
                                        grdReport.InsertRows lngRow + i + fintRowInserts, 1
                                    End If
                                    
                                    fintRowInserts = fintRowInserts + 1
                            
                                    If i = UBound(varRS, 2) + 1 Then
                                        i = i - 1
                                    End If
                                
                                
                            Else
                                intSeq = varRS(1, r)
                                strObjType = Left(varRS(0, r), 3)
                            
                                grdReport.Col = IdentifyColumn(lngCol, varRS(0, r)) 'lngCol + c
                            End If
                            
                            'Any cell spans applied to this column?........
                            '                                             ColAnchor,  RowAnchor,  NumCols,    NumRows
                             grdReport.GetCellSpan grdReport.Col, lngRow, varSpan(0), varSpan(1), varSpan(2), varSpan(3)
                            
                            If varSpan(2) > 0 Then 'There is a cell span
                                'If grdReport.Col > grdReport.Col * varSpan(2) And grdReport.Col <= varSpan(0) + varSpan(2) - 1 Then
                                    grdReport.Col = grdReport.Col * varSpan(2) - lngCol ' varSpan(0) + varSpan(2)
                                'End If
                                
                                'Does data exists in this cell?
                                grdReport.GetText grdReport.Col, lngRow, varCell
                                'If varCell<> vbnullstring Then
                                    'grdReport.Col = varSpan(0) + varSpan(2)
                                'End If
                                
                                intSpan = varSpan(2) - 1
                                intTotSpan = intTotSpan + intSpan
                                
                            End If
                            'grdReport.Col = lngCol + c
                            
                            If intPrevSeq <> intSeq And r > 0 And intSeq > 1 Then
                
                                i = i + 1
                                If fintRowInserts > 1 Then
                                    grdReport.InsertRows lngRow + i + fintRowInserts, 1
                                Else
                                    grdReport.InsertRows lngRow + i, 1
                                End If
                                'intTotSpan = 0
                                'ReDim Preserve fudtRowInsert(fintRowInserts)
                                fintRowInserts = fintRowInserts + 1
                                'fudtRowInsert(fintRowInserts).InsertedRows = fintRowInserts
                                'fudtRowInsert(fintRowInserts).StartRowlngRow
                                
                                'grdReport.AddCellSpan grdReport.Col, lngRow + i, varSpan(2), 1
                                'grdReport.InsertRows lngRow + i, 1
                                'intRowInserts = intRowInserts + 1
                                booRowAdded = True
                            End If
                            
                            If booRowAdded Then
                                grdReport.AddCellSpan grdReport.Col, lngRow + i, varSpan(2), 1
                            End If
                            grdReport.Row = lngRow + i + intRowInserts
                            
                            If grdReport.Col > 0 Then
                                Select Case strObjType
                                    Case "chk", "ccc", "cco", "opt"
                                        grdReport.CellType = CellTypeCheckBox
                                        booBooleanType = True
                                        
                                    Case Else
                                        booBooleanType = False
                                End Select
                            
                                If strItem(6) <> "grdSS" Then
                                    grdReport.TypeHAlign = TypeHAlignCenter
                                    grdReport.SetText grdReport.Col, lngRow + i + intRowInserts, varRS(2, r)
                                
                                
                                    Select Case strItem(7) 'Interior Borders
                                    
                                        Case "Vert"
                                            grdReport.SetCellBorder grdReport.Col, lngRow + i + intRowInserts, grdReport.Col + intSpan, lngRow + i, 1, 0, 1  'Left
                                            grdReport.SetCellBorder grdReport.Col, lngRow + i + intRowInserts, grdReport.Col + intSpan, lngRow + i, 2, 0, 1  'Right
                                        Case "Horiz"
                                            grdReport.SetCellBorder grdReport.Col, lngRow + i + intRowInserts, grdReport.Col + intSpan, lngRow + i + intRowInserts, 4, 0, 1 'Top
                                            grdReport.SetCellBorder grdReport.Col, lngRow + i + intRowInserts, grdReport.Col + intSpan, lngRow + i + intRowInserts, 8, 0, 1 'Bottom
                                        Case "Both"
                                            grdReport.SetCellBorder grdReport.Col, lngRow + i + intRowInserts, grdReport.Col + intSpan, lngRow + i, 16, 0, 1
                                    End Select
                                End If
                            
                             End If
                            
                            'If HideRow(0) > 0 Then
                                If HideRow(lngRow + fintRowInserts) And varRS(2, r) = "" Then grdReport.ROWHEIGHT(lngRow + fintRowInserts) = 0
                            'End If
    
                            c = c + 1
                            
                            If r > 0 And r < UBound(varRS, 2) Then
                                If intSeq <> intPrevSeq Then 'varRS(1, r + 1)
                                    c = 0
                                    
                                    'grdReport.InsertRows lngRow + i, 1
                                    'fintRowInserts = fintRowInserts + i
                                    'i = i + 1
                                End If
                            End If
    
                        End If
                       
                       intPrevSeq = intSeq
                    Next r
                    End If
                    'wayne 12/12
                    'Outline .......................
                   ' grdReport.SetCellBorder lngCol, lngRow + intRowInserts, lngCol + intLastCol - 1 + intTotSpan / (i + 1), lngRow + i + intRowInserts, 16, 0, 1
                    grdReport.SetCellBorder lngCol, lngRow + intRowInserts, lngCol + intLastCol - 1 + intTotSpan, lngRow + i + intRowInserts, 16, 0, 1
                    'Outline Header..................................
                   ' grdReport.SetCellBorder lngCol, lngRow - 1 + intRowInserts, lngCol + intLastCol - 1 + intTotSpan / (i + 1), lngRow - 1 + intRowInserts, 16, 0, 1
                   ' grdReport.SetCellBorder lngCol, lngRow - 1 + intRowInserts, lngCol + intLastCol - 1 + intTotSpan / (i + 1), lngRow - 1 + intRowInserts, 2, 0, 1
                    grdReport.SetCellBorder lngCol, lngRow - 1 + intRowInserts, lngCol + intLastCol - 1 + intTotSpan, lngRow - 1 + intRowInserts, 16, 0, 1
                    grdReport.SetCellBorder lngCol, lngRow - 1 + intRowInserts, lngCol + intLastCol - 1 + intTotSpan, lngRow - 1 + intRowInserts, 2, 0, 1
                    
                    
                    ReDim Preserve fudtRowInsert(UBound(fudtRowInsert) + 1)
                    fudtRowInsert(UBound(fudtRowInsert)).InsertedRows = fintRowInserts
                    fudtRowInsert(UBound(fudtRowInsert)).StartRow = lngRow
                    'SuppressUnusedCols grdReport, lngCol, lngRow, lngCol + CLng(intLastCol) - 1, lngRow + CLng(i)
                    
                End If
                fbooSkip = False
            End If
            
        Case "<Font>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            'lngCol2 = strItem(3)
            'lngRow2 = strItem(4)
            'Set_Block lngCol, lngRow, lngCol2, lngRow2
            '    grdReport.FontName = strItem(5)
            '    grdReport.FontSize = CInt(strItem(6))
            '    grdReport.ForeColor = CLng(strItem(7))
                
            'Set_Block lngCol, lngRow + intInsertedRows, lngCol, lngRow + intInsertedRows
             Set_Block lngCol, lngRow, lngCol, lngRow + intInsertedRows
                grdReport.FontName = strItem(3)
                grdReport.FontSize = CInt(strItem(4))
                grdReport.ForeColor = CLng(strItem(5))
                
                'Separate Type into distinct words.................
'                If strItem(8) <> "" Then
'                    strType = Separate_Words(strItem(8))
                If strItem(6) <> "" Then
                    strType = Separate_Words(strItem(6))
                    
                    If InStr(strType, " ") > 0 Then
                        strFontType() = Split(strType, " ")
                        For r = LBound(strFontType) To UBound(strFontType)
                            Select Case strFontType(r)
                                Case "Bold"
                                    grdReport.FontBold = True
                                Case "Italic"
                                    grdReport.FontItalic = True
                                Case "Underline"
                                    grdReport.FontUnderline = True
                            End Select
                        Next r
                    Else
                        Select Case strType
                            Case "Bold"
                                grdReport.FontBold = True
                            Case "Italic"
                                grdReport.FontItalic = True
                            Case "Underline"
                                grdReport.FontUnderline = True
                        End Select
                    End If
                End If
            
            grdReport.BlockMode = False
            
            
'        Case "<ForeColor>"
'            lngCol2 = strItem(3)
'            lngRow2 = strItem(4)
'            Set_Block lngCol, lngRow, lngCol2, lngRow2
'            grdReport.ForeColor = strItem(5)
'            grdReport.BlockMode = False
'
'        Case "<FontName>"
'            lngCol2 = strItem(3)
'            lngRow2 = strItem(4)
'            Set_Block lngCol, lngRow, lngCol2, lngRow2
'            grdReport.FontName = strItem(5)
'            grdReport.BlockMode = False
'
'        Case "<FontSize>"
'            lngCol2 = strItem(3)
'            lngRow2 = strItem(4)
'            Set_Block lngCol, lngRow, lngCol2, lngRow2
'            grdReport.FontSize = strItem(5)
'            grdReport.BlockMode = False
'
'        Case "<FontType>"
'            lngCol2 = strItem(3)
'            lngRow2 = strItem(4)
'            Set_Block lngCol, lngRow, lngCol2, lngRow2
'
'            Select Case strItem(5)
'                Case "Bold"
'                    grdReport.FontBold = True
'                Case "Italic"
'                    grdReport.FontItalic = True
'                Case "Underline"
'                    grdReport.FontUnderline = True
'                Case "Strike"
'                    grdReport.FontStrikethru = True
'            End Select
'
'            grdReport.BlockMode = False
            
        Case "<GridDim>"
            Me.Width = strItem(1) + 400
            Me.Height = strItem(2) + 1000
            centerForm Me
            grdReport.Width = strItem(1)
            grdReport.Height = strItem(2)
            
        Case "<GridLines>"
            grdReport.GridShowVert = strItem(1)
            grdReport.GridShowHoriz = strItem(2)
              
        Case "<Headers>"
            grdReport.ColHeadersShow = strItem(1)
            grdReport.RowHeadersShow = strItem(2)
            
        Case "<Image>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            grdReport.Col = lngCol
            grdReport.Row = lngRow
            grdReport.CellType = CellTypePicture
            grdReport.TypePictCenter = True
            
            strPath = Get_FilePath(strItem(3))
            
            If Not FolderExists(strPath) Then
                MkDir strPath
            End If
            
            If FileExists(strItem(3)) Then
                grdReport.TypePictPicture = LoadPicture(strItem(3))
            Else
                MsgBox "Image " & strItem(3) & " not found.", vbInformation
            End If
            
        Case "<Max>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            'intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.MaxCols = lngCol
            grdReport.MaxRows = lngRow
            
            
        Case "<MaxChars>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.Col = lngCol
            grdReport.Row = lngRow + intInsertedRows
            grdReport.TypeMaxEditLen = strItem(3)
            
            
        Case "<MergeCols>" 'spans cells in that column having the same content
            lngCol = strItem(1)
            lngCol2 = strItem(2)
            For c = lngCol To lngCol2
                grdReport.Col = c
                grdReport.ColMerge = MergeAlways
            Next c
            
        Case "<MergeRows>" 'spans cells in that row having the same content
            lngRow = strItem(1)
            lngRow2 = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            
            For r = lngRow + intInsertedRows To lngRow2 + intInsertedRows
                grdReport.Row = r
                grdReport.RowMerge = MergeAlways
            Next r
        
        Case "<PatientName>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.Name)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.Name = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
        
        Case "<PatientDOB>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.DOB)
            
         'If HideRow(0) > 0 Then
            If HideRow(lngRow + intInsertedRows) And fclsPatient.DOB = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
         'End If
        
        
        Case "<PatientSex>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.Gender)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.Gender = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientAddress>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.Address)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.Address = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientCityStateZip>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.CityStateZip)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.CityStateZip = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientPhone>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.Phone)
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.Phone = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientEmail>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.Email)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.Email = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientSS>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.SS)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.SS = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientFCC>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.FCC)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.FCC = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientRace>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.Race)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.Race = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientAge>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.Age)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.Age = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<PatientIntID>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(piPatientID)
        
        Case "<PatientExtID>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fclsPatient.ExtID)
            
            'If HideRow(0) > 0 Then
                If HideRow(lngRow + intInsertedRows) And fclsPatient.ExtID = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
            'End If
            
        Case "<ReferralPracticeName>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.Practice)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.Practice = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0

        Case "<ReferralPhysician>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.Physician)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.Physician = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
        
        Case "<ReferralAddress>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.Address)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.Address = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
           
        Case "<ReferralCityStateZipCode>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.CityStateZip)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.CityStateZip = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralPhone>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.Phone)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.Phone = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralCellPhone>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.CellPhone)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.CellPhone = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralFAX>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.FAX)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.FAX = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralEmail>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.Email)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.Email = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralUPIN>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.UPIN)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.UPIN = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralLicense>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.License)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.License = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0
               
        Case "<ReferralSpecialty>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(fudtReferral.Specialty)
            If HideRow(lngRow + intInsertedRows) And fudtReferral.Specialty = "" Then grdReport.ROWHEIGHT(lngRow + intInsertedRows) = 0

        Case "<RemoveRow>"
            If strItem(1) = "NoData" Then
                If UBound(strItem) > 1 Then
                    ReDim flngRowHidden(UBound(strItem) - 2) ' the trailing numbers in the <RemoveRow>~NoData tag
                    For i = LBound(flngRowHidden) To UBound(flngRowHidden)
                        flngRowHidden(i) = strItem(i + 2)
                    Next i
                End If
            End If
            
        Case "<RowHeaders>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            strRowHeader() = Split(strItem(3), ",")
            
            For i = LBound(strRowHeader) To UBound(strRowHeader)
                grdReport.SetText lngCol + i, lngRow + intInsertedRows, CVar(strRowHeader(i))
                grdReport.Col = lngCol + i
                grdReport.Row = lngRow + intInsertedRows
                grdReport.TypeHAlign = TypeHAlignCenter
                grdReport.FontBold = True
                grdReport.SetCellBorder lngCol + i, lngRow + intInsertedRows, lngCol + i, lngRow + intInsertedRows, 16, 0, 1
            Next i
        
        Case "<RowHeadersColWidth>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            strRowHeader() = Split(strItem(3), ",")
            
            For i = LBound(strRowHeader) To UBound(strRowHeader)
                grdReport.ColWidth(lngCol + i) = CInt(strRowHeader(i))
            Next i
        
        Case "<RowHeight>"
            lngRow = strItem(1)
            intInsertedRows = Get_InsertRowCount(lngRow)
            grdReport.ROWHEIGHT(lngRow + intInsertedRows) = strItem(2)
            
        Case "<ScrollBars>"
                If strItem(1) = "False" And strItem(2) = "False" Then
                    grdReport.ScrollBars = ScrollBarsNone
                ElseIf strItem(1) = "True" And strItem(2) = "False" Then
                    grdReport.ScrollBars = ScrollBarsVertical
                ElseIf strItem(1) = "False" And strItem(2) = "True" Then
                    grdReport.ScrollBars = ScrollBarsHorizontal
                ElseIf strItem(1) = "True" And strItem(2) = "True" Then
                    grdReport.ScrollBars = ScrollBarsBoth
                End If
        
        Case "<Signature>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            
            '                                    OCX       Control        Ref             Seq
            strEncodedSignature = Get_Signature(strItem(3), strItem(4), strItem(5), CInt(strItem(6)), piPatientID, piEncounterID, strRefData)
            If strEncodedSignature <> vbNullString Then
                'booInvalid = ccaSignature.DecodeSignatureValues(strEncodedSignature, "CurrentEstraceESTRACE2mgBUCCAL0TID04/28/200300621")
                booInvalid = ccaSignature.DecodeSignatureValues(strEncodedSignature, strRefData)
                picSignature.Picture = ccaSignature.Picture
                SavePicture picSignature.Image, App.Path & "\Images\sig.bmp"
                grdReport.Col = lngCol
                grdReport.Row = lngRow + intInsertedRows
                grdReport.CellType = CellTypePicture
                grdReport.TypePictCenter = True
                grdReport.TypePictPicture = LoadPicture(App.Path & "\Images\sig.bmp")
            End If
            
        
        Case "<Text>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            intInsertedRows = Get_InsertRowCount(lngRow)
            'grdReport.TopRow = lngRow - 8
            grdReport.SetText lngCol, lngRow + intInsertedRows, CVar(strItem(3))
            
            
        Case "<WordWrap>"
            lngCol = strItem(1)
            lngRow = strItem(2)
            lngCol2 = strItem(3)
            lngRow2 = strItem(4)
            intInsertedRows = Get_InsertRowCount(lngRow)
            Set_Block lngCol, lngRow + intRowInserts, lngCol2, lngRow2 + intRowInserts
            'grdReport.Col = lngCol
            'grdReport.Row = lngRow
            
            'If strItem(3) = "True" Then
                grdReport.TypeEditMultiLine = True
                grdReport.TypeTextWordWrap = True
            'Else
            '    grdReport.TypeEditMultiLine = False
            '    grdReport.TypeTextWordWrap = False
            'End If
            grdReport.BlockMode = False
            
    End Select
    
    If booNoteData Then
        intSpanWidth = 0
        'Check for cell span...............................
       '                                      ColAnchor,   RowAnchor,  NumCols,    NumRows
        grdReport.GetCellSpan lngCol, lngRow, varSpan(0), varSpan(1), varSpan(2), varSpan(3)
        
        For z = 0 To varSpan(2) - 1
            intSpanWidth = intSpanWidth + grdReport.ColWidth(lngCol + z)
        Next z
        
        If intSpanWidth = 0 Then intSpanWidth = grdReport.ColWidth(lngCol)
        
        'Set cell to Wordwrap................
        Set_Block lngCol, lngRow, lngCol, lngRow
        grdReport.TypeEditMultiLine = True
        grdReport.TypeTextWordWrap = True
        grdReport.BlockMode = False
        
        'Expand row size.................
        lblHidden.Width = intSpanWidth * 121
        grdReport.GetText lngCol, lngRow, varNotes
        lblHidden.Caption = CStr(varNotes)

        If varNotes <> vbNullString Then
        Dim cellheight As Long
        grdReport.BlockMode = True
        grdReport.Col = lngCol
        grdReport.Row = lngRow
        ' Define cell type as edit
        grdReport.CellType = CellTypeStaticText
        grdReport.TypeTextWordWrap = True
        'grdReport.TypeEditMultiLine = True
    ' Provide text for the cell
        grdReport.Text = varNotes
    ' Return the value of the selected cell
        cellheight = grdReport.MaxTextCellHeight
        grdReport.ROWHEIGHT(lngRow) = cellheight
        grdReport.BlockMode = False

        End If
        'grdReport.TopRow = lngRow - 8
        
        
    End If
    
'    For g = 1 To grdReport.MaxCols
'        intColWidth = grdReport.MaxTextColWidth(g)
'        grdReport.ColWidth(g) = intColWidth
'    Next g
            
    grdReport.SetActiveCell 0, 0
    
    Exit Sub
    
ProcessRS: 'Lab Result record sets......
    'booRowAdded = False
    intLastCol = strItem(8)
    'i = 0
    'intRowInserts = fintRowInserts
    'boogrdSSData = False
    
    For r = LBound(varRS, 2) To UBound(varRS, 2)
        intSpan = 0
        booSkip = False
        LoadRow_LabResults grdReport, varRS, r, lngCol, lngRow + fintRowInserts, fstrIncludedObject(), fstrExcludedObject()
        i = i + 1
        
        If i > 1 Then
            grdReport.InsertRows lngRow + i + fintRowInserts, 1
        End If
        
        fintRowInserts = fintRowInserts + 1

        If i = UBound(varRS, 2) + 1 Then
            'i = i - 1
        End If
        
         'Any cell spans applied to this column?........
         '                                             ColAnchor,  RowAnchor,  NumCols,    NumRows
        grdReport.GetCellSpan grdReport.Col, lngRow, varSpan(0), varSpan(1), varSpan(2), varSpan(3)
         
        If varSpan(2) > 0 Then 'There is a cell span
             
            grdReport.Col = grdReport.Col * varSpan(2) - lngCol
    
            'Does data exists in this cell?
            grdReport.GetText grdReport.Col, lngRow, varCell
    
            intSpan = varSpan(2) - 1
            intTotSpan = intTotSpan + intSpan
             
        End If
        
        If intPrevSeq <> intSeq And r > 0 And intSeq > 1 Then
            i = i + 1
            If fintRowInserts > 1 Then
                grdReport.InsertRows lngRow + i + fintRowInserts, 1
            Else
                grdReport.InsertRows lngRow + i, 1
            End If

            fintRowInserts = fintRowInserts + 1
            booRowAdded = True
        End If
        
        If booRowAdded Then
            grdReport.AddCellSpan grdReport.Col, lngRow + i, varSpan(2), 1
        End If
        
        grdReport.Row = lngRow + i + intRowInserts
        booBooleanType = False
        
        If HideRow(lngRow + fintRowInserts) And varRS(2, r) = "" Then grdReport.ROWHEIGHT(lngRow + fintRowInserts) = 0
                         
            c = c + 1
            
            If r > 0 And r < UBound(varRS, 2) Then
                If intSeq <> intPrevSeq Then
                    c = 0
                End If
            End If
    
        'End If
                       
        intPrevSeq = intSeq
                   
    Next r
    Return
    
Error_Execute:
    'Debug.Print
    
    If Err.Number = 9 Then
        
        If strItem(0) = "<FetchBooleanComp>" Then
            ReDim fudtAudit(0): Resume Next
        ElseIf strItem(0) = "<ColHeaders>" Then
            ReDim fudtGridHeader(0): Resume Next
        ElseIf strItem(0) = "<FetchRecordSet>" Then
            ReDim fudtRowInsert(0): Resume Next
        End If
    End If
    Exit Sub
    
  Resume
  
End Sub

Private Sub SuppressUnusedCols(vobjGrid As Object, lngCol As Long, lngRow As Long, lngCol2 As Long, lngRow2 As Long)
Dim r As Integer
Dim c As Integer

    For c = lngCol To lngCol2
        
        vobjGrid.Col = c
        vobjGrid.ColHidden = True
        
        For r = lngRow To lngRow2
            vobjGrid.Row = r
            If vobjGrid.Text <> vbNullString And vobjGrid.Text <> "0" Then
                vobjGrid.ColHidden = False
                Exit For
            End If
            
        Next r
    
    Next c
    
End Sub

Private Sub Set_Block(vlngCol, vlngRow, vlngCol2, vlngRow2)

    With grdReport
        .BlockMode = True
        .Col = vlngCol
        .Row = vlngRow
        .Col2 = vlngCol2
        .Row2 = vlngRow2
    End With
    
End Sub

Private Sub Form_Paint()
    ColorBurst Me, "BLUE", 128, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim booExported As Boolean
Dim strFileName As String
Dim strFilter As String
Dim strSystemDirectory As String
Dim strSystemDrive As String
Dim strLogFile As String
Dim strDefaultFile As String
Dim strPath As String
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next
  
    If fstrSavePath = "" Then
        fstrSavePath = App.Path & "\HTML"
        'Does folder exist?...............................
        'strSystemDirectory = Get_WindowsSystemDirectory
        'strSystemDrive = Left(lstrSystemDirectory, InStr(lstrSystemDirectory, ":"))
        If Not FolderExists(App.Path & "\HTML") Then
            MkDir App.Path & "\HTML"
        End If
    End If
    'strDefaultFile = fstrTitle & "_" & Format(Date, "yyyymmdd") & ".xls"
    strDefaultFile = fstrTitle & "_" & "Reports" & ".html"
    'strFileName = CommonDialogAction(CommonDialog, "Save", "Save as....", fstrSavePath, strDefaultFile)
   strFileName = fstrTitle & "_" & "Reports" & ".html"
   If strFileName <> vbNullString Then
       strPath = fstrSavePath & strFileName
       strLogFile = strPath & "\rpt_log.txt"
    'End If
       ' booExported = grdReport.ExportToExcel(strFileName, fstrTitle, strLogFile)
    booExported = grdReport.ExportToHTML(strFileName, False, strLogFile)
        If booExported Then
           ' MsgBox "Export complete.", vbInformation, "Result"
        Else
            MsgBox "Export did not succeed.", vbExclamation, "Result"
        End If
   End If


    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.Objects
      Set obj = Nothing
    Next obj
  Exit Sub

End Sub

Private Sub mnuFile_Exit_Click()
    Unload Me
End Sub


Private Sub mnuFile_Export_Click()
Dim booExported As Boolean
Dim strFileName As String
Dim strFilter As String
Dim strSystemDirectory As String
Dim strSystemDrive As String
Dim strLogFile As String
Dim strDefaultFile As String
Dim strPath As String

    If fstrSavePath = "" Then
        fstrSavePath = App.Path & "\ExcelFiles"
        'Does folder exist?...............................
        'strSystemDirectory = Get_WindowsSystemDirectory
        'strSystemDrive = Left(lstrSystemDirectory, InStr(lstrSystemDirectory, ":"))
        
        If Not FolderExists(App.Path & "\ExcelFiles") Then
            MkDir App.Path & "\ExcelFiles"
        End If
            
    End If

    strDefaultFile = fstrTitle & "_" & Format(Date, "yyyymmdd") & ".xls"
    'strDefaultFile = fstrTitle & "_" & Format(Date, "yyyymmdd") & ".html"
    
    strFileName = CommonDialogAction(CommonDialog, "Save", "Save as....", fstrSavePath, strDefaultFile)
   
    If strFileName <> vbNullString Then
        strPath = Get_FilePath(strFileName)
        strLogFile = strPath & "\rpt_log.txt"
        
        booExported = grdReport.ExportToExcel(strFileName, fstrTitle, strLogFile)
        'booExported = grdReport.ExportToHTML(strFileName, False, strLogFile)
        
        If booExported Then
            MsgBox "Export complete.", vbInformation, "Result"
        Else
            MsgBox "Export did not succeed.", vbExclamation, "Result"
        End If
    
    End If

End Sub

Private Sub mnuFile_Print_Click()

    On Error GoTo Error_Print
    
    ShowPrinter Me.hwnd
    
    Print_Grid grdReport, "Date: " & Now
    
    Exit Sub

Error_Print:
    MsgBox Err.Description, vbInformation, "Print"
    Exit Sub
    
End Sub


Private Sub Print_Grid(vgrd As fpSpread, vstrHeader As String)
Dim strNote(1) As String
Dim strNotes As String
Dim i As Integer

    With vgrd
   ' .PrintSmartPrint = True
    .PrintAbortMsg = "Printing - Click Cancel to quit"
    .PrintJobName = vstrHeader
   ' .PrintHeader = vstrHeader
    .PrintFooter = "/cPage #/p/n"
    .PrintBorder = False
    .PrintColHeaders = True
    .PrintColor = True
    .PrintGrid = True
    .PrintMarginTop = 500
    .PrintMarginBottom = 25
    .PrintMarginLeft = 50
    .PrintMarginRight = 50

    .PrintOrientation = PrintOrientationPortrait
    .PrintRowHeaders = False
    .PrintShadows = False
    .PrintUseDataMax = True
    
    .PrintSheet

    End With
End Sub


Private Sub mnuFile_ReRun_Click()
    grdReport.MaxCols = 5
    grdReport.MaxRows = 0
    fbooRun = True
    Get_Patient piPatientID
    Get_Referral
    Get_EncounterData
    Open_Script
    fbooRun = False
End Sub

Private Sub mnuFile_ShowGrid_Click()

    If mnuFile_ShowGrid.Checked Then
        grdReport.GridShowHoriz = False
        grdReport.GridShowVert = False
        mnuFile_ShowGrid.Checked = False
    Else
        grdReport.GridShowHoriz = True
        grdReport.GridShowVert = True
        mnuFile_ShowGrid.Checked = True
    End If
    
    
End Sub



Private Function IdentifyColumn(vlngCol As Long, vvarObject) As Integer
Dim i As Integer

    For i = LBound(fstrIncludedObject) To UBound(fstrIncludedObject)
    
        If InStr(vvarObject, fstrIncludedObject(i)) > 0 Then
            IdentifyColumn = vlngCol + i
            Exit For
        End If
        
    Next i
    
End Function

Private Sub Get_EncounterData()

    Get_Encounter piEncounterID

    fclsEncounter.StartDate = clsEncounter.StartDate
    fclsEncounter.EndDate = clsEncounter.EndDate
    fclsEncounter.VisitReason = clsEncounter.VisitReason
    fclsEncounter.Facility = clsEncounter.Facility
    fclsEncounter.Provider = clsEncounter.Provider
    fclsEncounter.ServicePlace = clsEncounter.ServicePlace
    fclsEncounter.ServicePlaceCode = clsEncounter.ServicePlaceCode
    fclsEncounter.EncounterType = clsEncounter.EncounterType
End Sub

Private Sub Remove_ResidualTags()
Dim c!, r!
Dim varCell

    For r = 1 To grdReport.MaxRows
    
        For c = 1 To grdReport.MaxCols
            grdReport.GetText c, r, varCell
            
            If varCell <> vbNullString Then
                If Left(varCell, 1) = "<" And Right(varCell, 2) = "/>" Then
                   grdReport.SetText c, r, ""
                End If
            End If
            
        Next c
        
    Next r
    
End Sub

Private Function HideRow(vlngRow As Long) As Boolean
Dim i As Integer
Dim intRow As Integer
Dim intAbsRow As Integer
Dim intStartRow As Integer


    On Error GoTo Error_HideRow
    
    intAbsRow = vlngRow
    
    If fintRowInserts > 0 Then
        For i = LBound(fudtRowInsert) To UBound(fudtRowInsert)
            
            If vlngRow > fudtRowInsert(i).StartRow And fudtRowInsert(i).StartRow > 0 Then
                intAbsRow = vlngRow - fintRowInserts
                Exit For
            'Else
                'intAbsRow = vlngRow
            End If
        
        Next i
    End If
    
    For i = LBound(flngRowHidden) To UBound(flngRowHidden)
        'If intAbsRow > 33 And intAbsRow < 50 Then
        '    Debug.Print vlngRow 'fudtRowInsert(0).StartRow
        'End If
'
'        If vlngRow > fudtRowInsert(0).StartRow Then
'            intRow = flngRowHidden(i) + fintRowInserts
'        Else
'            intRow = flngRowHidden(i)
'        End If
        
        If flngRowHidden(i) = intAbsRow Then
            HideRow = True
            Exit For
        End If
    Next i
    
    Exit Function
Error_HideRow:
    If Err.Number = 9 Then
        Exit Function
    Else
        ReDim flngRowHidden(0)
    End If
    Exit Function
    
    Resume
End Function

Private Function Get_InsertRowCount(vlngRow As Long) As Integer
Dim g%
     
 '    On Error Resume Next
     On Error GoTo ErrHandler

    For g = LBound(fudtRowInsert) To UBound(fudtRowInsert)
        If vlngRow >= fudtRowInsert(g).StartRow Then
            Get_InsertRowCount = fudtRowInsert(g).InsertedRows
        End If
    Next g
    
    Exit Function
    
ErrHandler:
 Exit Function
    
End Function

Private Sub Get_Referral()
Dim strReferralInfo As String
Dim strReferral() As String

   ' strReferralInfo = Get_ReferralInfo(piPatientID, piEncounterID)
    
    If strReferralInfo = "" Then Exit Sub
    
    strReferral() = Split(strReferralInfo, "~")
    fudtReferral.Practice = strReferral(0)
    fudtReferral.Physician = strReferral(1)
    fudtReferral.Address = strReferral(2)
    fudtReferral.CityStateZip = strReferral(3)
    fudtReferral.Phone = strReferral(4)
    fudtReferral.CellPhone = strReferral(5)
    fudtReferral.FAX = strReferral(6)
    fudtReferral.Email = strReferral(7)
    fudtReferral.UPIN = strReferral(8)
    fudtReferral.License = strReferral(9)
    fudtReferral.Specialty = strReferral(10)

End Sub
