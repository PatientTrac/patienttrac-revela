VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmMovePatient 
   Caption         =   "Select Patient Relocation"
   ClientHeight    =   3795
   ClientLeft      =   8205
   ClientTop       =   2880
   ClientWidth     =   8355
   FillColor       =   &H00800000&
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
   Icon            =   "frmMovePatient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   8355
   Begin VB.Frame fraPic 
      BackColor       =   &H00E0E0E0&
      Height          =   2235
      Left            =   0
      TabIndex        =   6
      Top             =   420
      Width           =   2775
      Begin VB.Label lblPhoto 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NO  PHOTO AVAILABLE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   555
         Left            =   600
         TabIndex        =   7
         Top             =   900
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Image imgPatient 
         Height          =   2025
         Left            =   60
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   405
      Left            =   3218
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
      Begin PT_XP_Button.PT_XP cmdSelect 
         Height          =   405
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmMovePatient.frx":08CA
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Select"
         CaptionAlignment=   5
         CaptionPosition =   4
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMovePatient.frx":0C64
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   405
         Left            =   960
         TabIndex        =   3
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmMovePatient.frx":0FB6
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMovePatient.frx":19C8
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.TextBox txtSelDate 
      Alignment       =   2  'Center
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
      Height          =   435
      Left            =   5610
      TabIndex        =   0
      Top             =   4320
      Width           =   2745
   End
   Begin FPSpreadADO.fpSpread grdTerms 
      Height          =   2055
      Left            =   4110
      TabIndex        =   4
      ToolTipText     =   "Right click mouse to deselect"
      Top             =   690
      Width           =   4155
      _Version        =   524288
      _ExtentX        =   7329
      _ExtentY        =   3625
      _StockProps     =   64
      ColHeaderDisplay=   0
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      DInformActiveRowChange=   0   'False
      DisplayColHeaders=   0   'False
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
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   2
      MaxRows         =   10
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   16761024
      ShadowText      =   8388608
      SpreadDesigner  =   "frmMovePatient.frx":1D1A
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   100
      ScrollBarTrack  =   1
      ClipboardOptions=   0
      AppearanceStyle =   0
   End
   Begin PT_XP_Button.PT_XP cmdTerms 
      Height          =   375
      Left            =   6930
      TabIndex        =   5
      ToolTipText     =   "Manage Resources"
      Top             =   2850
      Width           =   1335
      _ExtentX        =   2355
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
      Picture         =   "frmMovePatient.frx":2206
      AppearanceThemes=   3
      BorderDrawEvent =   1
      TransparentColor=   14215660
      Caption         =   "Terms"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmMovePatient.frx":2980
      ShowFocus       =   -1  'True
   End
   Begin MSComCtl2.DTPicker ccdAdmit 
      Height          =   375
      Left            =   60
      TabIndex        =   10
      Top             =   2790
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      CustomFormat    =   "MM/dd/yyy hh:mm:ss"
      DateIsNull      =   -1  'True
      Format          =   88408067
      CurrentDate     =   38862.7362384259
      MinDate         =   38353
   End
   Begin VB.Label lblSelectMove 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Select to Transfer Patient "
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4110
      TabIndex        =   9
      Top             =   300
      Width           =   4125
   End
   Begin VB.Label lblPatient 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   30
      TabIndex        =   8
      Top             =   60
      Width           =   2745
   End
End
Attribute VB_Name = "frmMovePatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Private mlngGridTermRow As Long
Public mbMovePatient As Boolean
Public mstrResource As String
Private mlngGridRow As Long
Private udtEnctr As udtEnctr
Private Type udtEnctr
    EnctrID As Long
    StartDate As String
    EndDate As String
    VisitReason As String
    Facility As String
    Provider As String
    Admission As String
    Discharge As String
    ServicePlace As String
    ServicePlaceCode As String
    EncounterType As String
    Appointment As String
    ProviderID As Long
    VisitType As String
    Intrvn_CD As String
End Type

Private Sub Clear_EnctrString()
    With udtEnctr
        .EnctrID = 0
        .StartDate = vbNullString
        .EndDate = vbNullString
        .VisitReason = vbNullString
        .Facility = vbNullString
        .Provider = vbNullString
        .Admission = vbNullString
        .Discharge = vbNullString
        .ServicePlace = vbNullString
        .ServicePlaceCode = vbNullString
        .EncounterType = vbNullString
        .Appointment = vbNullString
        .ProviderID = 0
        .VisitType = vbNullString
        .Intrvn_CD = vbNullString
    End With
End Sub
Private Function Check_Availability(vstrResource As String) As Boolean
Dim cnApptFormat As New ADODB.Connection
Dim rsApptFormat As New ADODB.Recordset
Dim strSQL As String
Dim strApptDate As String
Dim strApptTime As String


    On Error GoTo Error_Check_Availability
    Screen.MousePointer = vbHourglass
    strApptDate = Format(ccdAdmit.Value, "mm/dd/yyyy")
    strApptTime = Format(ccdAdmit.Value, "hh:mm AM/PM")
    If cnApptFormat.State = adStateClosed Then cnApptFormat.open psConnect(0)
    strSQL = "SELECT TIME_ID,APPT_TIME FROM RESOURCES_MGMT WHERE APPT_DATE = '" & strApptDate & "' AND " _
        & "RESOURCE_NAME = '" & vstrResource & "'"
    Set rsApptFormat = cnApptFormat.Execute(strSQL)
    If Not rsApptFormat.EOF Then
        
    End If
    

    
    
    
'    TIME_ID
'          ,APPT_DATE
'          RESOURCE_ID,
'          RESOURCE_NAME
'          ,FACILITY_ID
'          ,PROVIDER_ID
'          ,PATIENT_ID
'          ,APPT_TIME
'          ,DURATION
'          ,STATUS
'          ,REASON
'          ,NOTES
'          ,IN_SERVICE
'          ,USER_ID
'          ,ACTIVE
          

    If cnApptFormat.State = adStateOpen Then
        cnApptFormat.Close
        Set cnApptFormat = Nothing
        Set rsApptFormat = Nothing
    End If

    Screen.MousePointer = vbDefault
    Exit Function
Error_Check_Availability:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Check_Availability", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Format_GridTerms(vstrType As String)
Dim varColName
Dim lngRow As Long
Dim strList As String
Dim lngx As Long
Dim lngY As Long
Dim lngWidth As Long
Dim lngHeight As Long
Dim booEmbed As Boolean

    Screen.MousePointer = vbHourglass
    With grdTerms
        .Redraw = False
        .ClearRange 0, SpreadHeader, grdTerms.MaxRows, grdTerms.MaxCols, False
        .ProcessTab = True
        .FontBold = True
        .ColHeaderRows = 2
        .ColHeadersShow = True
        .RowHeadersShow = False
        .MaxCols = 4
        .ColWidth(1) = 15
'        .ColWidth(2) = 15
'        .ColWidth(3) = 5
'        .ColWidth(4) = 5

'        Format_CellTypeEdit grdTerms, SpreadHeader, 1, "2", "PATIENT STATUS", "TypeHAlignLeft", "RESOURCE"
''           Format_CellTypeEdit grdTerms, SpreadHeader, 3, "2", "CPT CODE", "TypeHAlignLeft", "CPT_CODE"
'        Load_Resources
        .BlockMode = True
            .Col = 1
            .Col2 = 1
            .Row = SpreadHeader + 1
            .Row2 = SpreadHeader + 1
            .RowHeight(SpreadHeader + 1) = 3
            .BackColor = 12632256
            .AddCellSpan 1, SpreadHeader + 1, .MaxCols, 1
            .SetCellBorder 1, SpreadHeader + 1, .MaxCols, SpreadHeader + 1, 16, &H80&, CellBorderStyleSolid
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Function
End Function








Private Function Load_Resources()
Dim cnApptFormat As New ADODB.Connection
Dim rsApptFormat As New ADODB.Recordset
Dim strSQL As String
Dim lngCnt As Long
Dim lngCol As Long
Dim lngRow As Long
Dim fso As New FileSystemObject

    On Error GoTo Error_Load_Resources
    Screen.MousePointer = vbHourglass
    If cnApptFormat.State = adStateClosed Then cnApptFormat.open psConnect(1)
        strSQL = "SELECT RESOURCE_NAME, RESOURCE_ID, SEQUENCE_NO FROM APPOINTMENT_RESOURCE WHERE RESOURCE_HEADER = 'True' Order by SEQUENCE_NO"
        Set rsApptFormat = cnApptFormat.Execute(strSQL)
        With rsApptFormat
           Do Until .EOF
                   If (!RESOURCE_NAME) <> vbNullString Then
                        lngCnt = lngCnt + 1
                        If grdTerms.MaxCols < lngCnt Then grdTerms.MaxCols = lngCnt
                       lngCol = lngCol + 1
                       Format_CellTypeStatic grdTerms, SpreadHeader + 3, lngCol, "1", (!RESOURCE_NAME), "TypeHAlignCenter", (!RESOURCE_ID), "", HeaderColor
                       grdTerms.Col = lngCol
                       grdTerms.ColWidth(lngCol) = 15
                       grdTerms.Row = SpreadHeader + 3
                       grdTerms.ColID = (!RESOURCE_NAME)
                   End If
               .MoveNext
           Loop
        End With
    If cnApptFormat.State = adStateOpen Then
        cnApptFormat.Close
        Set cnApptFormat = Nothing
        Set rsApptFormat = Nothing
    End If
    If piPatientID > 0 Then
        fraPic.ZOrder 0
        If Right(psImageFilePath, 1) <> "\" Then
            psImageFilePath = psImageFilePath & "\"
        End If
        If fso.FileExists(psImageFilePath & piPatientID & "\Photo.jpg") Then
           psPhoto = psImageFilePath & piPatientID & "\Photo.jpg"
           imgPatient.Picture = LoadPicture(psPhoto)
           lblPhoto.Visible = False
           imgPatient.Visible = True
        Else
          imgPatient.Visible = False
          lblPhoto.Visible = True
        End If
    End If

    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Resources:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Resources", App.EXEName & "." & Me.Name
    Exit Function
    Resume
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

Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional vlngBackColor As Long, Optional vbooUnderLine As Boolean)
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
            .FontBold = True
            If vbooUnderLine Then .FontUnderline = True
            If vlngBackColor Then .BackColor = vlngBackColor
            If Trim(vstrCellNote) <> vbNullString Then .CellNote = vstrCellNote
            If vstrSpan <> vbNullString Then
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                    .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                    lngSpan = Val(vstrSpan)
                    If lngSpan > 0 Then
                        .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                    End If
                End If
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote + 7
            .TypeVAlign = TypeVAlignCenter
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            .EditMode = False
          .BlockMode = False
    End With
End Function

Private Sub cmdClose_Click()
    mstrResource = vbNullString
    Unload Me
End Sub

Private Sub cmdSelect_Click()
Dim varCellValue

    If mlngGridRow <= 0 Then
        MsgBox "No Item selected." & vbNewLine & "If no Resource Change desired then select CLOSE, If Resource Change Please Select Desired Resource!", vbInformation
        Exit Sub
    End If
    With grdTerms
        .Col = 1
        .Row = mlngGridRow
        If .BackColor = grdCurItem Then
            .GetText .Col, mlngGridRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
               If TextFound("OPERATING ROOM", Trim(varCellValue), Len("OPERATING ROOM"), True) Then
                   If Check_Availability(CStr(varCellValue)) Then
                      ' 'Get_CurrentEnctr
                       mbMovePatient = True
                       If udtEnctr.EnctrID > 0 Then mbMovePatient = True
                   End If
               Else
                   'Get_CurrentEnctr
                   If udtEnctr.EnctrID > 0 Then mbMovePatient = True
               End If
            End If
        End If
    End With
    mstrResource = mstrResource
    udtAppointment.SchedDate = udtEnctr.StartDate
    Unload Me
    Exit Sub
End Sub
Private Function Get_CurrentEnctr(vstrType As String)
Dim i%
Dim sDate As String
Dim strSQL As String
Dim lReturn As Integer
Dim strEncntrStart As String
Dim strEnctTime As String
Dim strVisitType As String
Dim lngEnctrID As Long
Dim strPatType As String
Dim strEnctrType As String
Dim strSerPlace As String
Dim intServCode As Long
Dim lngFacID As Long
Dim iReturn As Integer
Dim cnStatus As New ADODB.Connection
Dim rsStatus As New ADODB.Recordset
Dim lngDateDiff As Long

    On Error GoTo Error_Get_CurrentEnctr
    Screen.MousePointer = vbHourglass
    mstrResource = vbNullString
    mstrResource = vstrType
    Dump_EncounterStrData
    Clear_EnctrString
    If frmPatientAppointment.mstrApptMode = "Surgery" Then
        If cnStatus.State = adStateClosed Then cnStatus.open psConnect(1)
       ' strSQL = "SELECT  CPT_CODE,  ENCNTR_TYPE, CPT_NAME, SERVICE_PLACE, PLACE_SERVICE_CD, ENCNTR_TYPE FROM RESOURCES_MGMT  WHERE RESOURCE_NAME like 'vstrType%'"
        Set rsStatus = cnStatus.Execute(strSQL)
        With rsStatus
           If Not rsStatus.EOF Then
              If (!CPT_Code) <> vbNullString Then udtEnctr.Intrvn_CD = (!CPT_Code)
              If (!CPT_Name) <> vbNullString Then udtEnctr.VisitType = (!CPT_Name)
              If (!Service_Place) <> vbNullString Then udtEnctr.ServicePlace = (!Service_Place)
              If (!Place_Service_CD) <> vbNullString Then udtEnctr.ServicePlaceCode = (!Place_Service_CD)
              If (!Encntr_Type) <> vbNullString Then udtEnctr.EncounterType = (!Place_Service_CD)
              If (!CPT_Desc) <> vbNullString Then udtEnctr.VisitReason = (!CPT_Desc)
           End If
        End With
        cnStatus.Close
        If cnStatus.State = adStateClosed Then cnStatus.open psConnect(0)
        Get_SurgFacility

        strSQL = "Select * from ENCOUNTER WHERE patient_id = " & frmPatientAppointment.mlngPatientID & " AND PURGE = 0 ORDER BY VISIT_TYPE"
        Set rsStatus = cnStatus.Execute(strSQL)
        sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
        With rsStatus
            Do Until .EOF
                If (!Visit_Type) = vstrType Then
                    If (!ENCNTR_START_DATE) <> vbNullString Then strEnctTime = (!ENCNTR_START_DATE)
                        If strEnctTime <> vbNullString Then
                        strEnctTime = Format(strEnctTime, "mm/dd/yyyy")
                        lngDateDiff = DateDiff("d", strEnctTime, Format(sDate, "mm/dd/yyyy"))
                        If lngDateDiff = 0 Then
                            udtEnctr.EnctrID = (!Encounter_ID)
                            Exit Do
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End With
    End If
    If udtEnctr.EnctrID = 0 Then
        If piProviderID > 0 Then udtEnctr.ProviderID = piProviderID
        udtEnctr.StartDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
        sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
        strSQL = "Insert into ENCOUNTER (PATIENT_ID, PROVIDER_ID, FACILITY_ID, VISIT_TYPE, SERVICE_PLACE," _
             & "ENCNTR_TYPE, ENCNTR_START_DATE,  INSERT_DATE, SERVICE_PLACE_CD, USER_ID, PURGE) " _
             & "values  (" & frmPatientAppointment.mlngPatientID & "," & udtEnctr.ProviderID & "," & udtFacility.FacID & ",'" _
             & udtEnctr.VisitType & "','" _
             & udtEnctr.ServicePlace & "','" _
             & udtEnctr.EncounterType & "','" _
             & udtEnctr.StartDate & "','" _
             & sDate & "'," & Val(udtEnctr.ServicePlaceCode) & ",'" & PsUserID & "',0)"
        Set rsStatus = cnStatus.Execute(strSQL)
        strSQL = "SELECT ENCOUNTER_ID FROM ENCOUNTER WHERE PATIENT_ID = " & frmPatientAppointment.mlngPatientID & " AND " _
            & "VISIT_TYPE = '" & udtEnctr.VisitType & "' AND ENCNTR_START_DATE = '" & udtEnctr.StartDate & "'"
        Set rsStatus = cnStatus.Execute(strSQL)
        If Not rsStatus.EOF Then
             udtEnctr.EnctrID = rsStatus.Fields("ENCOUNTER_ID")
             modSpecific.piEncounterID = udtEnctr.EnctrID
             Get_Encounter udtEnctr.EnctrID
        End If
    Else
        strSQL = "UPDATE ENCOUNTER SET ENCNTR_START_DATE = '" & udtEnctr.StartDate & "', UPDATE_DATE = '" & sDate & "' WHERE ENCOUNTER_ID = " & udtEnctr.EnctrID
        Set rsStatus = cnStatus.Execute(strSQL)
    End If
    Get_Encounter udtEnctr.EnctrID
    modSpecific.piEncounterID = udtEnctr.EnctrID
    piEncounterID = udtEnctr.EnctrID
    If cnStatus.State = adStateOpen Then
        cnStatus.Close
        Set cnStatus = Nothing
        Set rsStatus = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_CurrentEnctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_CurrentEnctr", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function



Private Sub Form_Activate()

    If mFirst Then
       mFirst = False
    End If
    Format_GridTerms "RESOURCES"
    mstrResource = vbNullString
    ccdAdmit.CheckBox = True
    ccdAdmit.Value = Format(Now, "MM/dd/yyy hh:mm:ss")
    mbMovePatient = False
    If frmPatientAppointment.mlngPatientID > 0 Then
        fraPic.ZOrder 0
        If Right(psImageFilePath, 1) <> "\" Then
            psImageFilePath = psImageFilePath & "\"
        End If
        If FileExists(psImageFilePath & frmPatientAppointment.mlngPatientID & "\Photo.jpg") Then
           psPhoto = psImageFilePath & frmPatientAppointment.mlngPatientID & "\Photo.jpg"
           imgPatient.Picture = LoadPicture(psPhoto)
           lblPhoto.Visible = False
           imgPatient.Visible = True
        Else
          imgPatient.Visible = False
          lblPhoto.Visible = True
        End If
    End If
    centerForm (Me)
    Exit Sub
End Sub
Private Sub Form_Load()
    mFirst = True
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

  On Error Resume Next
  
  
    ccdAdmit.CheckBox = False
    Unload Me
    Exit Sub
 
End Sub


Private Sub grdTerms_Click(ByVal Col As Long, ByVal Row As Long)
    With grdTerms
        If Row >= 0 Then
            .BlockMode = True
                .Col = Col
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                 .BackColor = grdCurItem
             .BlockMode = False
        End If
    End With
    Exit Sub
End Sub


Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue
Dim strTag As String
Dim strResource As String

    With grdTerms
        If Row >= 0 Then
            .BlockMode = True
                .Col = Col
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                 .BackColor = grdCurItem
                 mlngGridRow = Row
             .BlockMode = False
        End If
    End With
    Exit Sub
End Sub


