VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmPatAppHist 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Patient Appointment History"
   ClientHeight    =   7950
   ClientLeft      =   7665
   ClientTop       =   3840
   ClientWidth     =   13395
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatAppHist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   13395
   Begin PT_XP_Button.PT_XP cmdCloseList 
      Height          =   345
      Left            =   6150
      TabIndex        =   0
      Top             =   7290
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
      Picture         =   "frmPatAppHist.frx":08CA
      AppearanceThemes=   3
      BorderDrawEvent =   1
      TransparentColor=   14737632
      Caption         =   "Close"
      DefaultThemes   =   2
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatAppHist.frx":12DC
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdPatAppt 
      Height          =   6405
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   13275
      _Version        =   458752
      _ExtentX        =   23416
      _ExtentY        =   11298
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
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
      GrayAreaBackColor=   16777215
      GridColor       =   8388608
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   7
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmPatAppHist.frx":162E
      UserResize      =   0
      Appearance      =   2
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdNewPat 
      Height          =   345
      Left            =   90
      TabIndex        =   2
      Top             =   150
      Width           =   1485
      _ExtentX        =   2619
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
      Picture         =   "frmPatAppHist.frx":190A
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Patient"
      Depth           =   5
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatAppHist.frx":1EA4
      ShowFocus       =   -1  'True
   End
   Begin MSComCtl2.DTPicker ccdStartDate 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   150
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
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
      CurrentDate     =   38862.7362384259
      MinDate         =   32874
   End
   Begin PT_XP_Button.PT_XP cmdAdmit 
      Height          =   315
      Left            =   5220
      TabIndex        =   4
      ToolTipText     =   "Enter/Reset Admission Date"
      Top             =   180
      Width           =   315
      _ExtentX        =   556
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
      Picture         =   "frmPatAppHist.frx":21F6
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "OK"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatAppHist.frx":2590
      ShowFocus       =   -1  'True
   End
   Begin MSComCtl2.DTPicker ccdEndDate 
      Height          =   375
      Left            =   6780
      TabIndex        =   6
      Top             =   150
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
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
      CurrentDate     =   38862.7362384259
      MinDate         =   32874
   End
   Begin PT_XP_Button.PT_XP cmdDischarge 
      Height          =   315
      Left            =   8640
      TabIndex        =   7
      ToolTipText     =   "Enter Discharge Date"
      Top             =   180
      Width           =   315
      _ExtentX        =   556
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
      Picture         =   "frmPatAppHist.frx":28E2
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "OK"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatAppHist.frx":2C7C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPrint 
      Height          =   345
      Left            =   10350
      TabIndex        =   9
      Top             =   150
      Width           =   1065
      _ExtentX        =   1879
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
      Picture         =   "frmPatAppHist.frx":2FCE
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Print"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatAppHist.frx":3568
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdExecute 
      Height          =   345
      Left            =   9330
      TabIndex        =   10
      Top             =   150
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
      Picture         =   "frmPatAppHist.frx":38BA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Run"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPatAppHist.frx":3E54
      PictureAlignment=   5
      ShowFocus       =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
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
      Height          =   165
      Left            =   5820
      TabIndex        =   8
      Top             =   210
      Width           =   1035
   End
   Begin VB.Label lblAdmt 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
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
      Height          =   165
      Left            =   2220
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmPatAppHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mFirst As Boolean
Private udtAppt As udtAppt
Private Type udtAppt
    Patient_ID As Long
    Provider_ID As Long
    SchedTime As String
    Duration As String
    SchedDate As String
    VisitType As String
    SchedType As String
    Status As String
    ProviderID As Long
    Reason As String
    PatientType As String
    RoomNumber As String
    RoomTelephone As String
    Last_Visit As String
    Next_Visit As String
End Type
Private Function Clear_UDT_Appt()
    With udtAppt
        .Patient_ID = 0
        .Provider_ID = 0
        .SchedTime = vbNullString
        .Duration = vbNullString
        .SchedDate = vbNullString
        .VisitType = vbNullString
        .SchedType = vbNullString
        .Status = vbNullString
        .ProviderID = 0
        .Reason = vbNullString
        .PatientType = vbNullString
        .RoomNumber = vbNullString
        .RoomTelephone = vbNullString
        .Last_Visit = vbNullString
        .Next_Visit = vbNullString
    End With
End Function


Private Sub Format_Grid()
Dim strList As String
Dim strSQL As String
Dim lngGroup2 As Long
Dim lngCol  As Long
Dim varCellValue
Dim varHeader
Dim lngCnt As Long
Dim sDate As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strTag As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection


  Screen.MousePointer = vbHourglass
    grdPatAppt.MaxRows = 0
    With grdPatAppt
        .ReDraw = False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .FontName = "Times New Roman"
        .GridSolid = True
        .ColHeaderRows = 6
        .GridShowHoriz = False
        .GridShowVert = False
        .RowHeadersShow = False
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .MaxCols = 7
        .ColWidth(1) = 12
        .ColWidth(2) = 12
        .ColWidth(3) = 23.5
        .ColWidth(4) = 12
        .ColWidth(5) = 24
        .ColWidth(6) = 12
        .ColWidth(7) = 12
    End With
    Add_Header
    If piPatientID > 0 Then Get_Patient piPatientID
    Format_CellTypeEdit grdPatAppt, SpreadHeader + 2, 1, "maxcols", "PATIENT APPOINTMENT HISTORY" & " " & "FOR" & " " & UCase(clsPatient.Name), "TypeHAlignCenter", "", False, NormalColor
    sDate = Format(sDate, "MMMM DD YYYY")
    Add_Spacer SpreadHeader + 3
    lngCol = 1
    Format_CellTypeEdit grdPatAppt, SpreadHeader + 4, lngCol, "1", "Appointment Date", "TypeHAlignCenter", "APPT_DATE", False, HeaderColor
    grdPatAppt.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 2
    Format_CellTypeEdit grdPatAppt, SpreadHeader + 4, lngCol, "1", "Appointment Time", "TypeHAlignCenter", "APPT_TIME", False, HeaderColor
    grdPatAppt.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 3
    Format_CellTypeEdit grdPatAppt, SpreadHeader + 4, lngCol, "1", "Purpose of Visit", "TypeHAlignCenter", "PURPOSE_VISIT", False, HeaderColor
    grdPatAppt.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 4
    Format_CellTypeEdit grdPatAppt, SpreadHeader + 4, lngCol, "1", "Duration", "TypeHAlignCenter", "DURATION", False, HeaderColor
    grdPatAppt.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 5
    Format_CellTypeEdit grdPatAppt, SpreadHeader + 4, lngCol, "1", "Provider", "TypeHAlignCenter", "PROVIDER_ID", False, HeaderColor
    grdPatAppt.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 6
    Format_CellTypeEdit grdPatAppt, SpreadHeader + 4, lngCol, "1", "Status", "TypeHAlignCenter", "STATUS", False, HeaderColor
    grdPatAppt.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 7
    Format_CellTypeEdit grdPatAppt, SpreadHeader + 4, lngCol, "1", "Reason", "TypeHAlignCenter", "REASON", False, HeaderColor
    grdPatAppt.SetCellBorder lngCol, SpreadHeader + 4, lngCol, SpreadHeader + 4, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 5
    grdPatAppt.ReDraw = True
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
         Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Grid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_Grid", App.EXEName & "." & Me.Name
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
            Format_CellTypePicture grdPatAppt, SpreadHeader, 1, "1", udtFacility.Logo, "TypeHAlignCenter", "LOGO", "75"
            Format_CellTypeEdit grdPatAppt, SpreadHeader, 2, "6", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
        Else
            Format_CellTypeEdit grdPatAppt, SpreadHeader, 1, "7", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
        End If
    Else
        Format_CellTypeEdit grdPatAppt, SpreadHeader, 1, "7", Trim(udtFacility.FacHeader), "TypeHAlignCenter", "FAC_NAME"
    End If
    Add_Spacer SpreadHeader + 1
    Screen.MousePointer = vbDefault
    Exit Function
Error_Add_Header:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Add_Header", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vbooUnderLine As Boolean, Optional vlngBackColor As Long)
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
            .SetText vlngCol, vlngRow, vstrText
            .TypeEditMultiLine = True   'Make multi-lines
            .FontName = "Times New Roman"
            .FontSize = 8
            .FontBold = True
            If vbooUnderLine Then .FontUnderline = True
            If vlngBackColor > 0 Then .BackColor = vlngBackColor
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
              .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            Else
              lngSpan = vstrSpan
              If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            .Text = vstrText
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
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
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypePicture(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long)
Dim CellNote As Long
Dim lngSpan As Long
Dim fso As New FileSystemObject

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypePicture
            .TypePictCenter = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            Else
               If lngSpan > 0 Then
                    lngSpan = Val(vstrSpan)
                    .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                End If
             End If
            If fso.FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypeCombo(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrString As String, vstrHAligh As String)
Dim CellNote As Long
    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeComboBox
            .TypeMaxEditLen = 32000
            .TypeComboBoxList = vstrString
            .TypeComboBoxWidth = 1
            .TypeComboBoxCurSel = 0
            .FontName = "Times New Roman"
            .FontSize = 8
            .FontBold = True
        .BlockMode = False
    End With
End Function
Private Sub Add_Spacer(vlngRow As Long)

    With grdPatAppt
        vlngRow = vlngRow
       If .MaxRows < vlngRow Then .MaxRows = vlngRow
        .BlockMode = True
            .Col = 1
            .Col2 = 1
            .Row = vlngRow
            .Row2 = vlngRow
            .RowHeight(vlngRow) = 3
            .BackColor = 12632256
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 1, vlngRow, .MaxCols, vlngRow, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub
Private Function Read_Patient_Data()
Dim strSQL As String
Dim lngCol As Long
Dim lngRow As Long
Dim cnAppt As New ADODB.Connection
Dim rsAppt As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strSysName As String

    On Error GoTo Error_Read_Patient_Data
    Screen.MousePointer = vbHourglass
    If piPatientID = 0 Then
        MsgBox "Please select a patient first!", vbInformation, "PatientTrac Appointment History"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    If cnAppt.State = adStateClosed Then cnAppt.open psConnect(0)
    If Not IsNull(ccdStartDate.Value) Then
        udtTime.StartDate = Format(ccdStartDate.Value, "mm/dd/yyyy")
    Else
        udtTime.StartDate = vbNullString
    End If
    If Not IsNull(ccdEndDate.Value) Then
        udtTime.EndDate = Format(ccdEndDate.Value, "mm/dd/yyyy")
    End If
    If Trim(udtTime.StartDate) <> vbNullString Then
        strSQL = "Collect_Patient_Appointment_PT_ID_Range " & piPatientID & ",'" & udtTime.StartDate & "','" & udtTime.EndDate & "'"
    Else
        strSQL = "Get__Patient_Appointment_PT_ID " & piPatientID
    End If
    Set rsAppt = cnAppt.Execute(strSQL)
    grdPatAppt.ReDraw = False
    With rsAppt
        Do Until .EOF
            Clear_UDT_Appt
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "SCHEDULE_DATE"
                                udtAppt.SchedDate = Fld.Value
                            Case "APPT_TIME"
                                udtAppt.SchedTime = Fld.Value
                            Case "PURPOSE_VISIT"
                                udtAppt.VisitType = Fld.Value
                            Case "DURATION"
                                udtAppt.Duration = Fld.Value
                            Case "PROVIDER_ID"
                                udtAppt.Provider_ID = Fld.Value
                            Case "STATUS"
                                udtAppt.Status = Fld.Value
                            Case "REASON"
                                udtAppt.Reason = Fld.Value
                            Case "PATIENT_ID"
                                udtAppt.Patient_ID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtAppt.SchedDate <> vbNullString Then
                lngRow = lngRow + 1
                grdPatAppt.MaxRows = lngRow
                grdPatAppt.SetRowItemData lngRow, udtAppt.Patient_ID
                lngCol = 1
                Format_CellTypeEdit grdPatAppt, lngRow, lngCol, "1", udtAppt.SchedDate, "TypeHAlignCenter", "APPT_DATE"
                lngCol = 2
                Format_CellTypeEdit grdPatAppt, lngRow, lngCol, "1", udtAppt.SchedTime, "TypeHAlignLeft", "APPT_TIME"
                lngCol = 3
                Format_CellTypeEdit grdPatAppt, lngRow, lngCol, "1", udtAppt.VisitType, "TypeHAlignLeft", "PURPOSE_VISIT"
                lngCol = 4
                Format_CellTypeEdit grdPatAppt, lngRow, lngCol, "1", udtAppt.Duration, "TypeHAlignLeft", "DURATION"
                lngCol = 5
                Get_EnctrProviderName udtAppt.Provider_ID
                Format_CellTypeEdit grdPatAppt, lngRow, lngCol, "1", clsEncounter.Provider, "TypeHAlignLeft", "PROVIDER_ID"
                lngCol = 6
                Format_CellTypeEdit grdPatAppt, lngRow, lngCol, "1", udtAppt.Status, "TypeHAlignLeft", "STATUS"
                lngCol = 7
                Format_CellTypeEdit grdPatAppt, lngRow, lngCol, "1", udtAppt.Reason, "TypeHAlignLeft", "REASON"
                For lngCol = 1 To grdPatAppt.MaxCols
                    grdPatAppt.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Next lngCol
            End If
        Loop
    End With
    grdPatAppt.ReDraw = False
    Set Flds = Nothing
    If cnAppt.State = adStateOpen Then
        cnAppt.Close
        Set cnAppt = Nothing
        Set rsAppt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Patient_Data:
   Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Patient_Data", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Private Sub cmdAdmit_Click()
    ccdStartDate.Value = Format(Now, "mm/dd/yyyy")
    Exit Sub
End Sub

Private Sub cmdCloseList_Click()

    Unload Me
    Exit Sub
End Sub

Private Sub cmdDischarge_Click()
    ccdEndDate.Value = Format(Now, "mm/dd/yyyy")
End Sub

Private Sub cmdExecute_Click()
    grdPatAppt.ClearRange 1, 1, grdPatAppt.MaxCols, grdPatAppt.MaxRows, False
    grdPatAppt.MaxRows = 0
    Read_Patient_Data
End Sub


Private Sub cmdNewPat_Click()
Dim i%
    On Error GoTo Error_Select_Patient
    Screen.MousePointer = vbHourglass
    udtTime.EndDate = vbNullString
    udtTime.StartDate = vbNullString
    frmListPatient.mbShowEnctr = False
    piPatientID = 0
    frmListPatient.mlngListItemData = 0
    frmListPatient.Show vbModal
    If frmListPatient.mlngListItemData <> 0 Then
        piPatientID = Val(frmListPatient.mlngListItemData)
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
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer
Dim CellNote As Long
Dim intPrint As Integer
Dim strRef As String

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    mbooCancelPrint = False
    Printer.TrackDefault = True
    Get_Patient piPatientID
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation
        Exit Sub
    End If
    With grdPatAppt
        For intPrint = 1 To Printer.Copies
            .PrintType = PrintTypeAll
            .PrintJobName = clsPatient.Name & "_" & "Appointment Report"
            .hDCPrinter = Printer.hdc
            .PrintColHeaders = True
            .PrintRowHeaders = False
            .PrintScalingMethod = PrintScalingMethodSmartPrint
            .PrintCenterOnPageH = True
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
    Write_Error Err.Description, Err.Number, "", "cmdPrint_Click", App.EXEName & "." & TypeName(Me)
    Exit Sub
End Sub

Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
    End If
     Get_Patient piPatientID
    udtTime.EndDate = vbNullString
    udtTime.StartDate = vbNullString
    Format_Grid
    centerForm Me
End Sub

Private Sub Form_Load()
    mFirst = True
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    grdPatAppt.ClearRange 0, 0, grdPatAppt.MaxCols, grdPatAppt.MaxRows, False
    Unload Me
    Exit Sub
End Sub



Private Sub grdPatAppt_DblClick(ByVal Col As Long, ByVal Row As Long)

    Exit Sub
End Sub


