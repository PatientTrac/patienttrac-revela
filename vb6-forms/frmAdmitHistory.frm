VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAdmitHistory 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Patient Admission History"
   ClientHeight    =   6450
   ClientLeft      =   6030
   ClientTop       =   2880
   ClientWidth     =   11295
   ControlBox      =   0   'False
   Icon            =   "frmAdmitHistory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   435
      Left            =   5115
      TabIndex        =   0
      Top             =   5850
      Width           =   1065
      _ExtentX        =   1879
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
      Picture         =   "frmAdmitHistory.frx":08CA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmAdmitHistory.frx":12DC
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdData 
      Height          =   4995
      Left            =   30
      TabIndex        =   1
      Top             =   690
      Width           =   11220
      _Version        =   524288
      _ExtentX        =   19791
      _ExtentY        =   8811
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      ColHeaderDisplay=   0
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   16777215
      GridColor       =   8388608
      MaxCols         =   7
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmAdmitHistory.frx":162E
      UserResize      =   0
      Appearance      =   2
      TextTipDelay    =   30
      ClipboardOptions=   0
      AppearanceStyle =   0
   End
   Begin PT_XP_Button.PT_XP cmdPrint 
      Height          =   465
      Left            =   10170
      TabIndex        =   2
      Top             =   60
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
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
      Picture         =   "frmAdmitHistory.frx":1B3A
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Print"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmAdmitHistory.frx":20D4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdNewPat 
      Height          =   465
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   820
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
      Picture         =   "frmAdmitHistory.frx":2426
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Select Patient"
      Depth           =   5
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmAdmitHistory.frx":29C0
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmAdmitHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mPatID As Long
Private Sub Format_Grid()
Dim lngCol As Long
Dim lngRow As Long

    On Error GoTo Err_Format_Grid
    Screen.MousePointer = vbHourglass
    Add_Header
    With grdData
        .ColHeaderRows = 6
        .MaxCols = 7
        .MaxRows = 0
        If piPatientID > 0 Then Get_Patient piPatientID
        Format_CellTypeEdit grdData, SpreadHeader + 2, 1, "7", "ADMISSION HISTORY FOR" & vbNewLine & "PATIENT:" & " " & UCase(clsPatient.LastName) & ", " & UCase(clsPatient.FirstName) & vbNewLine & "MEDICAL RECORD NUMBER:" & " " & clsPatient.IntId, "TypeHAlignCenter", "ADMISSION_DATE", HeaderColor
        grdData.SetCellBorder 1, SpreadHeader, 7, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdData
        lngCol = 1
        .ColWidth(lngCol) = 17.5
         Format_CellTypeStatic grdData, SpreadHeader + 4, lngCol, "1", "ADMISSION DATE", "TypeHAlignCenter", "ADMISSION_DATE", HeaderColor
         grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        .ColWidth(lngCol) = 17.2
         Format_CellTypeStatic grdData, SpreadHeader + 4, lngCol, "1", "DISCHARGE DATE", "TypeHAlignCenter", "DISCHARGE_DATE", HeaderColor
         grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        .ColWidth(lngCol) = 10
         Format_CellTypeStatic grdData, SpreadHeader + 4, lngCol, "1", "FLOOR", "TypeHAlignCenter", "FLOOR", HeaderColor
         grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        .ColWidth(lngCol) = 10
         Format_CellTypeStatic grdData, SpreadHeader + 4, lngCol, "1", "WING", "TypeHAlignCenter", "WING", HeaderColor
         grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        .ColWidth(lngCol) = 10
         Format_CellTypeStatic grdData, SpreadHeader + 4, lngCol, "1", "ROOM", "TypeHAlignCenter", "ROOM", HeaderColor
         grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        .ColWidth(lngCol) = 10
         Format_CellTypeStatic grdData, SpreadHeader + 4, lngCol, "1", "BED", "TypeHAlignCenter", "BED", HeaderColor
         grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        .ColWidth(lngCol) = 15
         Format_CellTypeStatic grdData, SpreadHeader + 4, lngCol, "1", "TRANS DATE", "TypeHAlignCenter", "TRANS_DATE", HeaderColor
         grdData.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    End With
    Add_Spacer SpreadHeader + 5, grdData
    Screen.MousePointer = vbDefault
    Exit Sub
Err_Format_Grid:
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
End Sub
Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
              .Col = vlngCol
              .Col2 = vlngCol
              .Row = vlngRow
              .Row2 = vlngRow
              grdData.CellType = CellTypeStaticText
              .TypeMaxEditLen = 32000
              .TypeEditMultiLine = True
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              If vlngBackColor > 0 Then .BackColor = vlngBackColor
              If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
              ElseIf Trim(vstrSpan) <> vbNullString Then
                lngSpan = vstrSpan
                If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
              End If
              .Text = Trim(vstrText)
              .SetText vlngCol, vlngRow, Trim(vstrText)
              CellNote = .MaxTextCellHeight
              If .RowHeight(vlngRow) < CellNote Then
                .RowHeight(vlngRow) = CellNote
              End If
              .TypeVAlign = TypeVAlignCenter
                If vstrHAligh = "TypeHAlignCenter" Then
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .TypeHAlign = TypeHAlignLeft
                End If
              If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
End Function

Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)

    With vObjGrid
        vlngRow = vlngRow
       If .MaxRows < vlngRow Then .MaxRows = vlngRow
        .BlockMode = True
            .Col = 0
            .Col2 = 0
            .Row = vlngRow
            .Row2 = vlngRow
            .RowHeight(vlngRow) = 3.5
            .BackColor = 12632256
            .AddCellSpan 0, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Load_Grid()
Dim cnAdmit As New ADODB.Connection
Dim rsAdmit As New ADODB.Recordset
Dim strFirstName As String
Dim strLastName As String
Dim sValue
Dim strAdmit As String
Dim strDisc As String
Dim booSkip As Boolean
Dim lngRow As Long
Dim lngCol As Long
Dim strSQL As String

    On Error GoTo LoadGrid_Error:
    Screen.MousePointer = vbHourglass
    If mPatID = 0 Then mPatID = piPatientID
    If cnAdmit.State = adStateClosed Then cnAdmit.open psConnect(0)
    strSQL = "SELECT PATIENT_ADMISSION.*, PATIENT_ADMISSION_ROOMS.FLOOR, PATIENT_ADMISSION_ROOMS.WING, " _
        & "PATIENT_ADMISSION_ROOMS.ROOM, PATIENT_ADMISSION_ROOMS.BED FROM PATIENT_ADMISSION " _
        & "LEFT OUTER JOIN PATIENT_ADMISSION_ROOMS ON PATIENT_ADMISSION_ROOMS.PATIENT_ADMITID = PATIENT_ADMISSION.PATIENT_ADMITID " _
        & "WHERE PATIENT_ADMISSION.PATIENT_ID = " & mPatID & " ORDER BY PATIENT_ADMISSION.ADMISSION_DATE DESC"
        Set rsAdmit = cnAdmit.Execute(strSQL)
        
    With rsAdmit
       Do Until .EOF
        booSkip = False
        If (!ADMISSION_DATE) <> vbNullString Then strAdmit = Format((!ADMISSION_DATE), "MM/DD/YYYY hh:mm AM/PM")
        If strAdmit <> "01/01/1900" Then
            grdData.MaxRows = grdData.MaxRows + 1
            lngRow = grdData.MaxRows
            lngCol = 1
            Format_CellTypeStatic grdData, lngRow, lngCol, "1", strAdmit, "TypeHAlignCenter", "ADMISSION_DATE"
            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Else
            booSkip = True
        End If
        
        If (!DISCHARGE_DATE) <> vbNullString Then strDisc = Format((!DISCHARGE_DATE), "MM/DD/YYYY hh:mm AM/PM")
        If strDisc <> "01/01/1900" Then
            lngCol = 2
            Format_CellTypeStatic grdData, lngRow, lngCol, "1", strDisc, "TypeHAlignCenter", "DISCHARGE_DATE"
            grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
        If Not booSkip Then
            lngCol = 3
            If (!Floor) <> vbNullString Then
                Format_CellTypeStatic grdData, lngRow, lngCol, "1", (!Floor), "TypeHAlignCenter", "FLOOR"
                grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
            lngCol = 4
            If (!Wing) <> vbNullString Then
                Format_CellTypeStatic grdData, lngRow, lngCol, "1", (!Wing), "TypeHAlignCenter", "WING"
                grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
            lngCol = 5
            If (!room) <> vbNullString Then
                Format_CellTypeStatic grdData, lngRow, lngCol, "1", (!room), "TypeHAlignCenter", "ROOM"
                grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
            lngCol = 6
            If (!Bed) <> vbNullString Then
                Format_CellTypeStatic grdData, lngRow, lngCol, "1", (!Bed), "TypeHAlignCenter", "BED"
                grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
            lngCol = 7
            If (!Trans_Date) <> vbNullString Then
                Format_CellTypeStatic grdData, lngRow, lngCol, "1", Format((!Trans_Date), "MM/DD/YYYY"), "TypeHAlignCenter", "TRANS_DATE"
                grdData.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        End If
          .MoveNext
       Loop
    End With
    If cnAdmit.State = adStateOpen Then
        Set rsAdmit = Nothing
        cnAdmit.Close
        Set cnAdmit = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
LoadGrid_Error:
    Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
                .Col = vlngCol
                .Col2 = vlngCol
                .Row = vlngRow
                .Row2 = vlngRow
                .CellType = CellTypeEdit
                .TypeMaxEditLen = 32000
                .TypeEditMultiLine = True   'Make multi-lines
                .FontName = "Times New Roman"
                .FontSize = 8
                .FontBold = True
                .ForeColor = GridBlue
                If vlngBackColor > 0 Then .BackColor = vlngBackColor
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                    .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                ElseIf Trim(vstrSpan) <> vbNullString Then
                  lngSpan = vstrSpan
                  If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                End If
                .TypeVAlign = TypeVAlignCenter
                If vstrHAligh = "TypeHAlignRight" Then
                    .TypeHAlign = TypeHAlignRight
                ElseIf vstrHAligh = "TypeHAlignLeft" Then
                  .TypeHAlign = TypeHAlignLeft
                Else
                  .TypeHAlign = TypeHAlignCenter
                End If
                .SetText vlngCol, vlngRow, Trim(vstrText)
                .Text = Trim(vstrText)
                CellNote = .MaxTextCellHeight
                If .RowHeight(vlngRow) < CellNote Then
                  .RowHeight(vlngRow) = CellNote + 7
                End If
                If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
End Function

Public Sub ReadData()
    Format_Grid
    Load_Grid
End Sub

Private Sub cmdNewPat_Click()
Dim i%

    On Error GoTo Error_Select_Patient
    Screen.MousePointer = vbHourglass
    frmListPatient.mbShowEnctr = False
    piPatientID = 0
    frmListPatient.mlngListItemData = 0
    frmListPatient.Show vbModal
    If frmListPatient.mlngListItemData <> 0 Then
        piPatientID = Val(frmListPatient.mlngListItemData)
        Format_Grid
        Load_Grid
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Select_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, udtSchedPatient.Name, "Select_Patient", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdPrint_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(1) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Multi Appointment Schedule"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    With grdData
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Admission History" & "_" & clsPatient.IntId
        .PrintColor = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintSmartPrint = True
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "Multiple Appointments"
        .PrintBorder = True
        .PrintColHeaders = True
        .PrintRowHeaders = True
        .PrintScalingMethod = PrintScalingMethodSmartPrint
        .PrintMarginLeft = 650
        .PrintMarginRight = 700
        .PrintMarginBottom = 900
        .PrintMarginTop = 900
        .PrintUseDataMax = True
        .PrintColor = True
        .PrintShadows = False
        .PrintSheet
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdPrint_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Print", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Form_Load()
    Format_Grid
    ReadData
    centerForm Me
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  Unload Me
  Exit Sub
End Sub


Private Sub grdData_DblClick(ByVal Col As Long, ByVal Row As Long)

    Exit Sub
End Sub

Private Sub grdData_GotFocus()
    On Error GoTo Error_Handler
 
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Handler:
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Function Add_Header()
Dim CellNote As Long
Dim strPhoto As String
Dim Idx As Integer
Dim strSQL As String
Dim lngRow As Long
Dim lngCol As Long
Dim Arr
Dim fso As New FileSystemObject

    On Error GoTo Error_Add_Header
    Screen.MousePointer = vbHourglass
    Get_Bill_Facility
    Get_ReportFacility clsFacility.BillFacility
    If udtFacility.Logo <> vbNullString Then
        If fso.FileExists(udtFacility.Logo) Then
            Format_CellTypePicture grdData, SpreadHeader, 1, "1", udtFacility.Logo, "TypeHAlignCenter", "LOGO", "75"
            Format_CellTypeEdit grdData, SpreadHeader, 2, "6", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
        Else
            Format_CellTypeEdit grdData, SpreadHeader, 1, "7", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
        End If
    Else
        Format_CellTypeEdit grdData, SpreadHeader, 1, "7", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
    End If
    Add_Spacer SpreadHeader + 1, grdData
    Screen.MousePointer = vbDefault
    Exit Function
Error_Add_Header:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Add_Header", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Format_CellTypePicture(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long
Dim fso As New FileSystemObject

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            If vlngRowHeight > 0 Then .RowHeight(vlngRow) = vlngRowHeight
            .CellType = CellTypePicture
            If fso.FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
            End If
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(lngSpan) Then
                lngSpan = vstrSpan
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
             End If
             If Trim(vstrCellNote) <> vbNullString Then .CellNote = vstrCellNote
            .TypePictMaintainScale = True
            .TypePictCenter = True
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
End Function

