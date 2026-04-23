VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmSystemError 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Error Log"
   ClientHeight    =   7410
   ClientLeft      =   4095
   ClientTop       =   2730
   ClientWidth     =   13665
   Icon            =   "frmSystemError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   Begin FPSpreadADO.fpSpread grdError 
      Height          =   5955
      Left            =   30
      TabIndex        =   0
      Top             =   750
      Width           =   13605
      _Version        =   458752
      _ExtentX        =   23998
      _ExtentY        =   10504
      _StockProps     =   64
      AutoCalc        =   0   'False
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   8388608
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   9
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmSystemError.frx":0442
      UnitType        =   0
      UserResize      =   0
      Appearance      =   2
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdRun 
      Height          =   345
      Left            =   7080
      TabIndex        =   1
      Top             =   150
      Width           =   1665
      _ExtentX        =   2937
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
      Picture         =   "frmSystemError.frx":0723
      AppearanceMode  =   1
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Run Query"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSystemError.frx":0CBD
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   345
      Left            =   6285
      TabIndex        =   2
      ToolTipText     =   "Close and do not Save"
      Top             =   6840
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmSystemError.frx":100F
      AppearanceThemes=   3
      BorderDrawEvent =   1
      TransparentColor=   15790320
      Caption         =   "Done"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSystemError.frx":1A21
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDelete 
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
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
      Picture         =   "frmSystemError.frx":1D73
      AppearanceMode  =   1
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Clear Log"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSystemError.frx":210D
      ShowFocus       =   -1  'True
   End
   Begin MSComCtl2.DTPicker ccdEndDate 
      Height          =   285
      Left            =   4740
      TabIndex        =   4
      Top             =   390
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   16761024
      CalendarTitleForeColor=   8388608
      CalendarTrailingForeColor=   8421504
      CheckBox        =   -1  'True
      CustomFormat    =   "MM/dd/yyy"
      DateIsNull      =   -1  'True
      Format          =   85721091
      CurrentDate     =   38795.9993055556
   End
   Begin MSComCtl2.DTPicker ccdStartDate 
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   390
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   16761024
      CalendarTitleForeColor=   8388608
      CalendarTrailingForeColor=   8421504
      CheckBox        =   -1  'True
      CustomFormat    =   "MM/dd/yyy "
      DateIsNull      =   -1  'True
      Format          =   85721091
      CurrentDate     =   38795.0006944444
   End
   Begin PT_XP_Button.PT_XP cmdPrint 
      Height          =   345
      Left            =   12330
      TabIndex        =   8
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
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
      Picture         =   "frmSystemError.frx":245F
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Print"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSystemError.frx":29F9
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblFrom 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label lblTo 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Height          =   195
      Left            =   4770
      TabIndex        =   6
      Top             =   150
      Width           =   1125
   End
End
Attribute VB_Name = "frmSystemError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fbooByPassList As Boolean
Private fbooByPassCheck As Boolean
Private fintColWidth(10) As Integer
Private fbooLoading As Boolean
Private Function LookUpCellTag(vstrName As String) As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim booExit As Boolean

    On Error GoTo Error_LookUpCellTag
    Screen.MousePointer = vbHourglass
    With grdError
        LookUpCellTag = vbNullString
            For lngCol = 1 To .MaxCols
                    .Col = lngCol
                    .Row = SpreadHeader + 2
                    If .CellTag <> vbNullString Then
                        If TextFound(vstrName, .CellTag, Len(vstrName), True) Then
                            LookUpCellTag = lngCol & "|" & lngRow
                            booExit = True
                            Exit For
                        End If
                    End If
                If booExit Then Exit For
            Next lngCol
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUpCellTag:
    Screen.MousePointer = vbDefault
   Exit Function
    Resume
End Function

Private Sub Format_Grid()
Dim strSQL As String
Dim cnError As New ADODB.Connection
Dim rsError As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim R%
Dim lngCol As Long
Dim lngRow As Long

    On Error GoTo Error_Format_Grid
    Screen.MousePointer = vbHourglass
    With grdError
        .ClearRange 1, 1, .MaxCols, .MaxRows, False
        .MaxRows = 0
        .ColHeaderRows = 4
        .RowHeadersShow = False
        .MaxCols = 8
        Format_CellTypeEdit grdError, SpreadHeader, 1, "8", "PATIENTTRAC ERROR LOG", "TypeHAlignCenter", "", HeaderColor
        grdError.SetCellBorder lngCol, SpreadHeader, 8, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdError
        lngCol = 1
        .ColWidth(lngCol) = 15
        Format_CellTypeEdit grdError, SpreadHeader + 2, lngCol, "1", "APPLICATION", "TypeHAlignCenter", "APP_NAME", HeaderColor
        grdError.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        .ColWidth(lngCol) = 20
        Format_CellTypeEdit grdError, SpreadHeader + 2, lngCol, "1", "ERROR DESCRIPTION", "TypeHAlignCenter", "ERR_DESC", HeaderColor
        grdError.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        .ColWidth(lngCol) = 24.5
        Format_CellTypeEdit grdError, SpreadHeader + 2, lngCol, "1", "ERROR MESSAGE", "TypeHAlignCenter", "REMARKS", HeaderColor
        grdError.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        .ColWidth(lngCol) = 10
        Format_CellTypeEdit grdError, SpreadHeader + 2, lngCol, "1", "PATIENT ID", "TypeHAlignCenter", "PATIENT_ID", HeaderColor
        grdError.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        .ColWidth(lngCol) = 10
        Format_CellTypeEdit grdError, SpreadHeader + 2, lngCol, "1", "ENCOUNTER ID", "TypeHAlignCenter", "ENCOUNTER_ID", HeaderColor
        grdError.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 6
        .ColWidth(lngCol) = 10
        Format_CellTypeEdit grdError, SpreadHeader + 2, lngCol, "1", "COMPUTER NAME", "TypeHAlignCenter", "COMPUTER", HeaderColor
        grdError.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 7
        .ColWidth(lngCol) = 10
        Format_CellTypeEdit grdError, SpreadHeader + 2, lngCol, "1", "USER NAME", "TypeHAlignCenter", "USER_NAME", HeaderColor
        grdError.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 8
        .ColWidth(lngCol) = 10
        Format_CellTypeEdit grdError, SpreadHeader + 2, lngCol, "1", "DATE", "TypeHAlignCenter", "ERR_DATE", HeaderColor
        grdError.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 3, grdError
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Grid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_Grid", Me.Name
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
              .SetText vlngCol, vlngRow, Trim(vstrText)
              .TypeEditMultiLine = True   'Make multi-lines
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
Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrVAlign As String, Optional lngBackColor As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long
    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeStaticText
            .TypeMaxEditLen = 32000
            .FontName = "Times New Roman"
            .FontSize = 8
            .FontBold = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
            If vstrVAlign = "TypeVAlignCenter" Then
                .TypeVAlign = TypeVAlignCenter
            ElseIf vstrVAlign = "TypeVAlignBottom" Then
                .TypeVAlign = TypeVAlignBottom
            Else
               .TypeVAlign = TypeVAlignTop
            End If
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            ElseIf vstrHAligh = "TypeHAlignRight" Then
                .TypeHAlign = TypeHAlignRight
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If vstrCellNote <> vbNullString Then
                .CellNote = vstrCellNote
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            End If
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
            If lngBackColor > 0 Then .BackColor = lngBackColor
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
            .BackColor = GridBlue
            .AddCellSpan 0, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdDelete_Click()
    grdError.ClearRange 1, 1, grdError.MaxCols, grdError.MaxRows, False
    Exit Sub
End Sub

Private Sub cmdPrint_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer
Dim lngPageCnt As Long
Dim intPrint%

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    mbooCancelPrint = False
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Invoice"
        Exit Sub
    End If
    lngPageCnt = grdError.PrintPageCount
    If lngPageCnt <= 0 Then lngPageCnt = 1
    With grdError
        For intPrint = 1 To Printer.Copies
                .PrintSmartPrint = True
                .PrintType = PrintTypeAll
                .PrintJobName = "Error Log" & "_" & udtTime.StartDate & " " & "TO" & " " & udtTime.EndDate
                .hDCPrinter = Printer.hdc
                .PrintRowHeaders = False
                .PrintRowHeaders = False
                .PrintCenterOnPageH = True
                .PrintCenterOnPageV = False
                .PrintMarginTop = 500
                .PrintMarginBottom = 770
                .PrintBorder = True
                .PrintColor = True
                .PrintOrientation = PrintOrientationPortrait
                .PrintShadows = False
                .PrintSheet
        Next intPrint
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdPrint_Click:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbInformation, "PatientTrac Invoice"
    Exit Sub
End Sub

Private Sub cmdRun_Click()
Dim strStart As String

    Clear_Time
    If Not IsNull(ccdStartDate.Value) Then
        strStart = Format(ccdStartDate, "mm/dd/yyyy")
        udtTime.StartDate = Format(strStart, "mm/dd/yyyy 00:00:00 AM/PM")
    Else
        MsgBox "Please select start date!", vbInformation, "PatientTrac Error Log"
        Exit Sub
    End If
    If Not IsNull(ccdEndDate.Value) Then
        If Format(ccdStartDate, "mm/dd/yyyy") = Format(ccdEndDate.Value, "mm/dd/yyyy") Then
            udtTime.EndDate = DateAdd("h", 23, udtTime.StartDate)
            udtTime.EndDate = Format(udtTime.EndDate, "mm/dd/yyyy 23:59:00 AM/PM")
        Else
            udtTime.EndDate = Format(ccdEndDate.Value, "mm/dd/yyyy 00:00:00 AM/PM")
            udtTime.EndDate = DateAdd("h", 23, udtTime.EndDate)
        End If
    Else
        udtTime.EndDate = Format(udtTime.StartDate, "mm/dd/yyyy 00:00:00 AM/PM")
        udtTime.EndDate = DateAdd("h", 23, udtTime.EndDate)
    End If
    If udtTime.StartDate <> vbNullString And udtTime.EndDate <> vbNullString Then udtTime.Range = True
    Run_Query
End Sub



Private Sub Form_Activate()
Dim strDate As String
    fbooLoading = True
    Format_Grid
    strDate = DateAdd("d", -7, Now)
    ccdStartDate.Value = Format(strDate, "mm/dd/yyyy hh:mm AM/PM")
    ccdEndDate.Value = Format(Now, "mm/dd/yyyy hh:mm AM/PM")
    fbooLoading = False
End Sub

Private Sub Form_Load()
    centerForm Me
    Me.Top = Me.Top + 180
End Sub


Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Unload Me
    Exit Sub
End Sub

Private Sub Run_Query()
Dim strSQL As String
Dim cnError As New ADODB.Connection
Dim rsError As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow  As Long
Dim lngCol As Long
Dim varCellValue
Dim R As Integer
Dim strDate As String
Dim strUser As String
Dim strComputer As String
Dim strApp As String
Dim strFromDate As String
Dim strToDate As String
Dim strSort As String
Dim strTerm As String
Dim strOrderBy As String
Dim strDeleteRow As String
Dim Arr

    On Error GoTo Error_Run_Query
    Screen.MousePointer = vbHourglass
    If cnError.State = adStateClosed Then cnError.open psConnect(0)
    strSQL = "SELECT * FROM ERRORS WHERE INSERT_DATE BETWEEN '" & udtTime.StartDate & "' And '" & udtTime.EndDate & " ' ORDER BY app_name, ERR_DESC, COMPUTER, USER_NAME"
    Set rsError = cnError.Execute(strSQL)
    Do Until rsError.EOF
        lngRow = lngRow + 1
        grdError.MaxRows = lngRow
        Set Flds = rsError.Fields
        For Each Fld In Flds
            If Not IsNull(Fld.Value) Then
                If Trim(Fld.Value) <> vbNullString Then
                    Select Case UCase(Fld.Name)
                        Case "ERR_DESC"
                             Arr = Split(LookUpCellTag("ERR_DESC"), "|")
                            If UBound(Arr) >= 0 Then
                                lngCol = Arr(0)
                                Format_CellTypeEdit grdError, lngRow, lngCol, "1", Trim(Fld.Value), "TypeHAlignCenter", "ERR_DESC"
                                grdError.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        Case "APP_NAME"
                             Arr = Split(LookUpCellTag("APP_NAME"), "|")
                            If UBound(Arr) >= 0 Then
                                lngCol = Arr(0)
                                Format_CellTypeEdit grdError, lngRow, lngCol, "1", Trim(Fld.Value), "TypeHAlignCenter", "APP_NAME"
                                grdError.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        Case "ERR_DATE"
                             Arr = Split(LookUpCellTag("ERR_DATE"), "|")
                            If UBound(Arr) >= 0 Then
                                lngCol = Arr(0)
                                Format_CellTypeEdit grdError, lngRow, lngCol, "1", Trim(Fld.Value), "TypeHAlignCenter", "ERR_DATE"
                                grdError.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        Case "PATIENT_ID"
                             Arr = Split(LookUpCellTag("PATIENT_ID"), "|")
                            If UBound(Arr) >= 0 Then
                                lngCol = Arr(0)
                                Format_CellTypeEdit grdError, lngRow, lngCol, "1", Trim(Fld.Value), "TypeHAlignCenter", "PATIENT_ID"
                                grdError.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        Case "ENCOUNTER_ID"
                             Arr = Split(LookUpCellTag("ENCOUNTER_ID"), "|")
                            If UBound(Arr) >= 0 Then
                                lngCol = Arr(0)
                                Format_CellTypeEdit grdError, lngRow, lngCol, "1", Trim(Fld.Value), "TypeHAlignCenter", "ENCOUNTER_ID"
                                grdError.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        Case "USER_NAME"
                             Arr = Split(LookUpCellTag("USER_NAME"), "|")
                            If UBound(Arr) >= 0 Then
                                lngCol = Arr(0)
                                Format_CellTypeEdit grdError, lngRow, lngCol, "1", Trim(Fld.Value), "TypeHAlignCenter", "USER_NAME"
                                grdError.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        Case "COMPUTER"
                             Arr = Split(LookUpCellTag("COMPUTER"), "|")
                            If UBound(Arr) >= 0 Then
                                lngCol = Arr(0)
                                Format_CellTypeEdit grdError, lngRow, lngCol, "1", Trim(Fld.Value), "TypeHAlignCenter", "COMPUTER"
                                grdError.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                        Case "REMARKS"
                             Arr = Split(LookUpCellTag("REMARKS"), "|")
                            If UBound(Arr) >= 0 Then
                                lngCol = Arr(0)
                                Format_CellTypeEdit grdError, lngRow, lngCol, "1", Trim(Fld.Value), "TypeHAlignCenter", "REMARKS"
                                grdError.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            End If
                    End Select
                End If
            End If
        Next
        rsError.MoveNext
    Loop
    If cnError.State = adStateOpen Then
        Set cnError = Nothing
        Set rsError = Nothing
    End If
     Screen.MousePointer = vbDefault
     Exit Sub
Error_Run_Query:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Run_Query", Me.Name
    Exit Sub
    Resume
End Sub

