VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmTask 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Message / Tasks"
   ClientHeight    =   8445
   ClientLeft      =   6765
   ClientTop       =   4665
   ClientWidth     =   15645
   Icon            =   "frmTask.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   15645
   Begin VB.ComboBox cboProviders 
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
      Left            =   7260
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   30
      Width           =   3555
   End
   Begin FPSpreadADO.fpSpread grdMessages 
      Height          =   7485
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   15585
      _Version        =   458752
      _ExtentX        =   27490
      _ExtentY        =   13203
      _StockProps     =   64
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   8388608
      GridSolid       =   0   'False
      MaxCols         =   10
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmTask.frx":08CA
      Appearance      =   2
      TextTip         =   4
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   345
      Left            =   7245
      TabIndex        =   1
      Top             =   7950
      Width           =   1155
      _ExtentX        =   2037
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
      Picture         =   "frmTask.frx":0B55
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmTask.frx":1567
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblMessages 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Message/Tasks for:"
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
      Left            =   4830
      TabIndex        =   3
      Top             =   90
      Width           =   2235
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Private mLoadPatient As Boolean
Private mChange As Boolean
Public mBeenHere As Boolean
Private mlngGridRow As Long
Private mlngGridCol As Long
Private mlngMessRow As Long
Private mlngMessCol As Long
Private mFirstLoad As Boolean
Public mlngPatientID As Long
Public mlngProviderID As Long
Private Color As OLE_COLOR
Const LB_FINDSTRING = &H18F
Private Declare Function SendMessage Lib "user32" _
       Alias "SendMessageA" _
   (ByVal Hwnd As Long, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long
Public mstrMessageType As String

Private Function Format_Assignment()
Dim cnMess As New ADODB.Connection
Dim rsMess As New ADODB.Recordset
Dim strSQL As String
Dim lngColWidth As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strHeader As String
Dim lngCnt As Long
Dim sDate As String
Dim bRet As Boolean

    On Error GoTo Error_Format_Assignment
    Screen.MousePointer = vbHourglass
    grdMessages.MaxRows = 0
    With grdMessages
        .ReDraw = False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .FontName = "Times New Roman"
        .GridSolid = False
        .GridShowHoriz = False
        .GridShowVert = False
        .RowHeadersShow = False
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
    End With
 
    If cnMess.State = adStateClosed Then cnMess.open psConnect(1)

    strSQL = "SELECT HEADERS from GRID_HEADERS where " _
        & "form_name = 'STAFF_ASSIGN'  "
        Set rsMess = cnMess.Execute(strSQL)
        With rsMess
            If Not .EOF Then
                bRet = True
            End If
        End With
    If Not bRet Then
        strSQL = "INSERT INTO GRID_HEADERS (FORM_NAME, HEADERS, SEQUENCE_NO ) " _
            & "VALUES ('STAFF_ASSIGN', 'DATE', 1)"
        Set rsMess = cnMess.Execute(strSQL)
        strSQL = "INSERT INTO GRID_HEADERS (FORM_NAME, HEADERS, SEQUENCE_NO ) " _
            & "VALUES ('STAFF_ASSIGN', 'ASSIGNMENT', 2)"
        Set rsMess = cnMess.Execute(strSQL)
        strSQL = "INSERT INTO GRID_HEADERS (FORM_NAME, HEADERS, SEQUENCE_NO ) " _
            & "VALUES ('STAFF_ASSIGN', 'PATIENT_ID', 3)"
        Set rsMess = cnMess.Execute(strSQL)
        strSQL = "INSERT INTO GRID_HEADERS (FORM_NAME, HEADERS, SEQUENCE_NO ) " _
            & "VALUES ('STAFF_ASSIGN', 'STATUS', 4)"
        Set rsMess = cnMess.Execute(strSQL)
        strSQL = "INSERT INTO GRID_HEADERS (FORM_NAME, HEADERS, SEQUENCE_NO ) " _
            & "VALUES ('STAFF_ASSIGN', 'COMPLETED BY', 5)"
        Set rsMess = cnMess.Execute(strSQL)
        strSQL = "INSERT INTO GRID_HEADERS (FORM_NAME, HEADERS, SEQUENCE_NO ) " _
            & "VALUES ('STAFF_ASSIGN', 'COMPLETION DATE', 6)"
        Set rsMess = cnMess.Execute(strSQL)
    End If
    
    strSQL = "SELECT HEADERS, FORM_ID, visible, SEQUENCE_NO from GRID_HEADERS where " _
        & "form_name = 'STAFF_ASSIGN' ORDER BY SEQUENCE_NO"
        Set rsMess = cnMess.Execute(strSQL)
        lngCnt = 0
        Do Until rsMess.EOF
            lngCnt = lngCnt + 1
            rsMess.MoveNext
        Loop
    strSQL = "SELECT HEADERS, FORM_ID, VISIBLE, SEQUENCE_NO from GRID_HEADERS WHERE FORM_NAME = 'STAFF_ASSIGN' ORDER BY SEQUENCE_NO "
    Set rsMess = cnMess.Execute(strSQL)
    With rsMess
        Do Until .EOF
            grdMessages.MaxCols = lngCnt
            If Not IsNull(!Headers) Then
                If Trim((!Headers)) <> vbNullString Then
                    strHeader = (!Headers)
                Else
                    strHeader = vbNullString
                End If
            End If
            If Trim(strHeader) <> vbNullString Then
                lngCol = lngCol + 1
                grdMessages.ColWidth(lngCol) = 21
                lngRow = SpreadHeader + 1
                Format_CellTypeEdit grdMessages, lngRow, lngCol, "1", UCase(strHeader), "TypeHAlignCenter", UCase(strHeader), HeaderColor
                grdMessages.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                If Not IsNull((!Visible)) Then
                    If (!Visible) = 0 Then grdMessages.ColHidden = True
                End If
            End If
        .MoveNext
        Loop
    End With
    sDate = Format(Now, "MMMM DD, YYYY")
    Format_CellTypeEdit grdMessages, SpreadHeader, 1, CStr(lngCnt), "STAFF ASSIGNMENT STATUS", "TypeHAlignCenter", "", HeaderColor
    grdMessages.SetCellBorder 1, SpreadHeader, lngCnt, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 2, grdMessages
    grdMessages.ReDraw = True
    If cnMess.State = adStateOpen Then
        cnMess.Close
        Set cnMess = Nothing
        Set rsMess = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_Assignment:
    Write_Error Err.Description, Err.Number, strSQL, "Format_Assignment", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

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
                'If vbooUnderline Then .FontUnderline = True
                If vlngBackColor > 0 Then .BackColor = vlngBackColor
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                  .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                ElseIf Trim(vstrSpan) <> vbNullString Then
                  lngSpan = vstrSpan
                  If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                End If
                .SetText vlngCol, vlngRow, Trim(vstrText)
                .Text = Trim(vstrText)
                CellNote = .MaxTextCellHeight
                If .RowHeight(vlngRow) < CellNote Then
                  .RowHeight(vlngRow) = CellNote + 7
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
            .ForeColor = GridBlue
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
            If vstrVAlign = "TypeVAlignTop" Then
                .TypeVAlign = TypeVAlignTop
            ElseIf vstrVAlign = "TypeVAlignBottom" Then
                .TypeVAlign = TypeVAlignBottom
            Else
               .TypeVAlign = TypeVAlignCenter
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
            .Col = 1
            .Col2 = .MaxCols
            .Row = vlngRow
            .Row2 = vlngRow
            .BackColor = 12632256
            .RowHeight(vlngRow) = 3.5
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, 16, GridBlue, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub

Private Function Read_Assignment()
Dim cnMess As New ADODB.Connection
Dim rsMess As New ADODB.Recordset
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strDate
Dim booAddRow As Boolean
Dim lngCol As Long
Dim lngRow As Long
Dim i%
Dim lngMesID As Long

    On Error GoTo Error_Read_Assignment
    Screen.MousePointer = vbHourglass
    grdMessages.ClearRange 1, 0, grdMessages.MaxCols, grdMessages.MaxRows, True
    Format_Messages
    If mlngProviderID > 0 Then mlngProviderID = mlngLogonUser
    If mlngLogonUser = 0 And mlngProviderID = 0 Then
        Get_DefaultProvider
        mlngProviderID = piProviderID
    End If
    If cboProviders.Text = vbNullString Then
        For i = 0 To cboProviders.listcount - 1
            If cboProviders.ItemData(i) = mlngProviderID Then
                mbooByPass = True
                cboProviders.Text = cboProviders.List(i)
                Exit For
            End If
        Next i
    End If
    If cnMess.State = adStateClosed Then cnMess.open psConnect(0)
    strSQL = "Read_Assignment_Provider " & mlngProviderID
    Set rsMess = cnMess.Execute(strSQL)
    With rsMess
        Do Until rsMess.EOF
            Clear_Phone
            lngMesID = 0
            Set Flds = rsMess.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "MESSAGE_ID"
                                lngMesID = Fld.Value
                            Case "MESSAGE_TIME"
                                udtTelephone.Time = Format(Fld.Value, "mm/dd/yyyy")
                            Case "MESSAGE_TYPE"
                                udtTelephone.CallType = Fld.Value
                            Case "SENDER_ID"
                                udtTelephone.SenderID = Fld.Value
                            Case "RECIEVOR_ID"
                                udtTelephone.ReceiverID = Fld.Value
                            Case "CALLER_NAME"
                                 udtTelephone.CallerName = Fld.Value
                            Case "MEDICATION_STATUS"
                                udtTelephone.MedStatus = Fld.Value
                            Case "PATIENT_SYMPTOMS"
                                udtTelephone.Symptoms = Fld.Value
                            Case "MESSAGE_NOTE"
                                udtTelephone.MessNote = Fld.Value
                            Case "MESSAGE_ACTION"
                                udtTelephone.MessAction = Fld.Value
                            Case "CONTACT_TYPE"
                                udtTelephone.ContactType = Fld.Value
                            Case "CALLER_ID"
                                udtTelephone.CallerID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtTelephone.CallerID > 0 Then
                lngRow = lngRow + 1
                grdMessages.MaxRows = lngRow
                grdMessages.SetRowItemData lngRow, lngMesID
                Format_CellTypeEdit grdMessages, lngRow, 1, "1", udtTelephone.Time, "TypeHAlignCenter", "MESSAGE_TIME"
                grdMessages.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Format_CellTypeEdit grdMessages, lngRow, 2, "1", udtTelephone.CallerName, "TypeHAlignLeft", CStr(udtTelephone.CallerID)
                grdMessages.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Format_CellTypeEdit grdMessages, lngRow, 3, "1", udtTelephone.MedStatus, "TypeHAlignLeft", "MEDICATION_STATUS"
                grdMessages.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Format_CellTypeEdit grdMessages, lngRow, 4, "1", udtTelephone.Symptoms, "TypeHAlignLeft", "PATIENT_SYMPTOMS"
                grdMessages.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Format_CellTypeEdit grdMessages, lngRow, 5, "1", udtTelephone.MessNote, "TypeHAlignLeft", "MESSAGE_NOTE"
                grdMessages.SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                Format_CellTypeEdit grdMessages, lngRow, 6, "1", udtTelephone.MessAction, "TypeHAlignLeft", "MESSAGE_ACTION"
                grdMessages.SetCellBorder 6, lngRow, 6, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            End If
        Loop
    End With
    If cnMess.State = adStateOpen Then
        cnMess.Close
        Set cnMess = Nothing
        Set rsMess = Nothing
    End If

    If mlngProviderID = 0 Then mlngProviderID = mlngLogonUser
    For i = 0 To cboProviders.listcount - 1
        If cboProviders.ItemData(i) = mlngProviderID Then
            mbooByPass = True
            cboProviders.Text = cboProviders.List(i)
            Exit For
        End If
    Next i
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Assignment:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Assignment", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Set_GridCurSel(vObjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, Optional vstrCellTag As String)
Dim Idx As Integer
Dim booData As Boolean
Dim strSpan As String
Dim strTag As String
Dim strCellNote As String
Dim lngCol As Long
Dim lngSpan As Long

    With vObjGrid
        .ReDraw = False
        .Col = vlngCol
        .Row = vlngRow
        If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        If .CellType = CellTypeComboBox Then
            For Idx = 0 To .TypeComboBoxCount
            .TypeComboBoxIndex = Idx
            If TextFound(Trim(vstrData), Trim(.TypeComboBoxString), Len(Trim(vstrData)), True) Then
                .TypeComboBoxCurSel = Idx
                Exit For
            End If
            Next Idx
        End If
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
    .ReDraw = True
    End With
End Function

Private Sub Format_Messages()
Dim cnMess As New ADODB.Connection
Dim rsMess As New ADODB.Recordset
Dim strSQL As String
Dim lngColWidth As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strHeader As String
Dim lngCnt As Long
Dim sDate As String

    On Error GoTo Error_Format_Messages
    Screen.MousePointer = vbHourglass
    grdMessages.MaxRows = 0
    With grdMessages
        .ReDraw = False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .FontName = "Times New Roman"
        .GridSolid = False
        .GridShowHoriz = False
        .GridShowVert = False
        .RowHeadersShow = False
        .ColHeaderRows = 3
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
 End With
 
    If cnMess.State = adStateClosed Then cnMess.open psConnect(1)
    strSQL = "Get_Grid_Headers 'Telephone Messages' "
        Set rsMess = cnMess.Execute(strSQL)
        lngCnt = 0
        Do Until rsMess.EOF
            lngCnt = lngCnt + 1
            rsMess.MoveNext
        Loop
    strSQL = "Get_Grid_Headers 'Telephone Messages' "
    Set rsMess = cnMess.Execute(strSQL)
    With rsMess
        Do Until .EOF
            grdMessages.MaxCols = lngCnt
            If Not IsNull(!Headers) Then
                If Trim((!Headers)) <> vbNullString Then
                    strHeader = (!Headers)
                Else
                    strHeader = vbNullString
                End If
            End If
            If Trim(strHeader) <> vbNullString Then
                lngCol = lngCol + 1
                grdMessages.ColWidth(lngCol) = 21
                lngRow = SpreadHeader + 1
                Format_CellTypeEdit grdMessages, lngRow, lngCol, "1", UCase(strHeader), "TypeHAlignCenter", UCase(strHeader), HeaderColor
                grdMessages.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                If Not IsNull((!Visible)) Then
                    If (!Visible) = 0 Then grdMessages.ColHidden = True
                End If
            End If
        .MoveNext
        Loop
    End With
    sDate = Format(Now, "MMMM DD, YYYY")
    Format_CellTypeEdit grdMessages, SpreadHeader, 1, CStr(lngCnt), "TELEPHONE MESSAGES FOR" & " " & UCase(sDate), "TypeHAlignCenter", "", HeaderColor
    grdMessages.SetCellBorder 1, SpreadHeader, lngCnt, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 2, grdMessages
    grdMessages.ReDraw = True
    If cnMess.State = adStateOpen Then
        cnMess.Close
        Set cnMess = Nothing
        Set rsMess = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Messages:
    Write_Error Err.Description, Err.Number, strSQL, "Format_Messages", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub Read_Messages()
Dim cnMess As New ADODB.Connection
Dim rsMess As New ADODB.Recordset
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strDate
Dim booAddRow As Boolean
Dim lngCol As Long
Dim lngRow As Long
Dim i%
Dim lngMesID As Long

    On Error GoTo Error_Read_Messages
    Screen.MousePointer = vbHourglass
    grdMessages.ClearRange 1, 0, grdMessages.MaxCols, grdMessages.MaxRows, True
    Format_Messages
    If mlngProviderID = 0 Then mlngProviderID = mlngLogonUser
    If mlngLogonUser = 0 And mlngProviderID = 0 Then
        Get_DefaultProvider
        mlngProviderID = piProviderID
    End If
    If cboProviders.Text = vbNullString Then
        For i = 0 To cboProviders.listcount - 1
            If cboProviders.ItemData(i) = mlngProviderID Then
                mbooByPass = True
                cboProviders.Text = cboProviders.List(i)
                Exit For
            End If
        Next i
    End If

    If cnMess.State = adStateClosed Then cnMess.open psConnect(0)
    strSQL = "Read_Message_Provider " & mlngProviderID
    Set rsMess = cnMess.Execute(strSQL)
    With rsMess
        Do Until rsMess.EOF
            Clear_Phone
            lngMesID = 0
            Set Flds = rsMess.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "MESSAGE_ID"
                                lngMesID = Fld.Value
                            Case "MESSAGE_TIME"
                                udtTelephone.Time = Format(Fld.Value, "mm/dd/yyyy")
                            Case "MESSAGE_TYPE"
                                udtTelephone.CallType = Fld.Value
                            Case "SENDER_ID"
                                udtTelephone.SenderID = Fld.Value
                            Case "RECIEVOR_ID"
                                udtTelephone.ReceiverID = Fld.Value
                            Case "CALLER_NAME"
                                 udtTelephone.CallerName = Fld.Value
                            Case "MEDICATION_STATUS"
                                udtTelephone.MedStatus = Fld.Value
                            Case "PATIENT_SYMPTOMS"
                                udtTelephone.Symptoms = Fld.Value
                            Case "MESSAGE_NOTE"
                                udtTelephone.MessNote = Fld.Value
                            Case "MESSAGE_ACTION"
                                udtTelephone.MessAction = Fld.Value
                            Case "CONTACT_TYPE"
                                udtTelephone.ContactType = Fld.Value
                            Case "CALLER_ID"
                                udtTelephone.CallerID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If udtTelephone.CallerID > 0 And IsProperCase(udtTelephone.ContactType) <> "Assignment" Then
                If udtTelephone.SenderID > 0 And udtTelephone.SenderID = mlngProviderID Or _
                    udtTelephone.SenderID = 0 And udtTelephone.ReceiverID = mlngProviderID Then
                    lngRow = lngRow + 1
                    grdMessages.MaxRows = lngRow
                    grdMessages.SetRowItemData lngRow, lngMesID
                    Format_CellTypeEdit grdMessages, lngRow, 1, "1", udtTelephone.Time, "TypeHAlignCenter", "MESSAGE_TIME"
                    grdMessages.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Format_CellTypeEdit grdMessages, lngRow, 2, "1", udtTelephone.CallerName, "TypeHAlignLeft", CStr(udtTelephone.CallerID)
                    grdMessages.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Format_CellTypeEdit grdMessages, lngRow, 3, "1", udtTelephone.MedStatus, "TypeHAlignLeft", "MEDICATION_STATUS"
                    grdMessages.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Format_CellTypeEdit grdMessages, lngRow, 4, "1", udtTelephone.Symptoms, "TypeHAlignLeft", "PATIENT_SYMPTOMS"
                    grdMessages.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Format_CellTypeEdit grdMessages, lngRow, 5, "1", udtTelephone.MessNote, "TypeHAlignLeft", "MESSAGE_NOTE"
                    grdMessages.SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    Format_CellTypeEdit grdMessages, lngRow, 6, "1", udtTelephone.MessAction, "TypeHAlignLeft", "MESSAGE_ACTION"
                    grdMessages.SetCellBorder 6, lngRow, 6, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    If IsProperCase(udtTelephone.ContactType) = "Inquiry" Then
                        Format_CellTypeEdit grdMessages, lngRow, 1, "1", udtTelephone.Time, "TypeHAlignCenter", "Inquiry"
                        grdMessages.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                End If
            End If
        Loop
    End With
    If cnMess.State = adStateOpen Then
        cnMess.Close
        Set cnMess = Nothing
        Set rsMess = Nothing
    End If

    If mlngProviderID = 0 Then mlngProviderID = mlngLogonUser
    For i = 0 To cboProviders.listcount - 1
        If cboProviders.ItemData(i) = mlngProviderID Then
            mbooByPass = True
            cboProviders.Text = cboProviders.List(i)
            Exit For
        End If
    Next i
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Read_Messages:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Messages", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Function Get_Staff()
Dim lngFacilityID As Long
Dim strSQL As String
Dim cnProv As New ADODB.Connection
Dim rsProv As New ADODB.Recordset
Dim lngProv As Long
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim strSuffix As String
Dim lngProvID As Long

    On Error GoTo Error_Get_Staff
    Screen.MousePointer = vbHourglass
    
    If cnProv.State = adStateClosed Then cnProv.open psConnect(0)
    strSQL = "Load_Active_Providers"
    Set rsProv = cnProv.Execute(strSQL)
    With rsProv
        Do Until .EOF
            strFirst = vbNullString
            strLast = vbNullString
            strMiddle = vbNullString
            strTitle = vbNullString
            strSuffix = vbNullString
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Fld.Value <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "FIRST_NAME"
                                strFirst = Fld.Value
                            Case "MIDDLE_NAME"
                                strMiddle = Fld.Value
                            Case "LAST_NAME"
                                strLast = Fld.Value
                            Case "SUFFIX"
                                strSuffix = Fld.Value
                            Case "TITLE"
                                strTitle = Fld.Value
                            Case "PROVIDER_ID"
                                lngProvID = Fld.Value
                        End Select
                    End If
                End If
            Next
            .MoveNext
            If Trim(strMiddle) <> vbNullString Then
               strMiddle = IsProperCase(strMiddle)
                strFirst = strFirst & " " & Left(strMiddle, 1)
            End If
            If Trim(strLast) <> vbNullString Then
                strFirst = strFirst & " " & strLast
            End If
            If Trim(strSuffix) <> vbNullString Then
                strFirst = strFirst & ", " & Trim(strSuffix)
            End If
            If Trim(strTitle) <> vbNullString Then
                strFirst = strFirst & ", " & Trim(strTitle)
            End If
            If Trim(strFirst) <> vbNullString Then
                cboProviders.AddItem strFirst
                cboProviders.ItemData(cboProviders.NewIndex) = lngProvID
            End If
        Loop
    End With
    LockWindowUpdate 0
    mbooByPass = False
    If cnProv.State = adStateOpen Then
       cnProv.Close
       Set cnProv = Nothing
       Set rsProv = Nothing
       Set Flds = Nothing
    End If
    Exit Function
    Screen.MousePointer = vbDefault
Error_Get_Staff:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Staff", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

 
Private Function Remove_InsertCells()
Dim iCol As Long
Dim iRow As Long
Dim vDate
Dim anchorcol  As Variant
Dim anchorrow As Variant
Dim colspan As Variant
Dim rowspan As Variant
Dim IsSpan As Boolean
Dim Span As Integer
Dim vdate2

    With grdMessages
        .ReDraw = False
        For iCol = 1 To .MaxCols
            For iRow = SpreadHeader To grdMessages.MaxRows
               .GetText 0, iRow, vDate
              If iRow > 0 Then .GetText 0, iRow - 1, vdate2
                If vDate <> vbNullString And vdate2 <> vbNullString Then
                    If vDate = vdate2 Then .DeleteRows (iRow), 1
                End If
            Next iRow
        Next iCol
        .ReDraw = True
    End With
    Exit Function
End Function
Private Function Remove_Span()
Dim Col As Variant
Dim Row As Variant
Dim Col2 As Variant
Dim Row2 As Variant
Dim Span As Integer
Dim anchorcol As Variant
Dim anchorrow As Variant
Dim colspan As Variant
Dim rowspan As Variant
Dim i As Long
Dim j As Long
Dim IsSpan As Boolean

    ' Get the selected cells
    grdMessages.getSelection 0, Col, Row, Col2, Row2

    ' Check to see if cells are already in span
    For i = 1 To 7
        For j = SpreadHeader To grdMessages.MaxRows
            Span = grdMessages.GetCellSpan(i, j, anchorcol, anchorrow, colspan, rowspan)
            ' If cells are in span, remove span
            If Span Then
                If grdMessages.Col = 1 Then anchorcol = "1"
                If j = SpreadHeader Then anchorrow = SpreadHeader
                If j = SpreadHeader + 1 Then anchorrow = SpreadHeader + 1
                If j = SpreadHeader + 2 Then anchorrow = SpreadHeader + 2
                grdMessages.RemoveCellSpan anchorcol, anchorrow
                IsSpan = True
            End If
        Next j
    Next i
    
    Exit Function
    
End Function



 
Private Sub Set_RowHeight()
Dim R%
Dim vNote
Dim Note As Long
Dim iCol%
Dim vEmptyTime

    With grdMessages
        .ReDraw = False
        For iCol = 1 To .MaxCols
            For R = 1 To .MaxRows
                    .BlockMode = True
                        .Col = iCol
                        .Col2 = iCol
                        .Row = R
                        .Row2 = R
                        .TypeEditMultiLine = True   'Make multi-line
                        .TypeTextWordWrap = True  'Make multi-line
                        .TypeMaxEditLen = 20000
                        .TypeHAlign = TypeHAlignCenter
                        .TypeVAlign = TypeVAlignCenter
                        .GetText iCol, R, vNote
                        .Text = vNote
                        Note = .MaxTextCellHeight
                        If .RowHeight(.Row) < Note Then
                            .RowHeight(.Row) = Note
                        End If
                    .BlockMode = False
            Next R
        Next iCol
        .ReDraw = True
    End With
    
    Exit Sub
End Sub

Private Sub cboProviders_Click()
Dim i%
    If mbooByPass Then mbooByPass = False: Exit Sub
    If Trim(cboProviders.Text) <> vbNullString Then
        For i = 0 To cboProviders.listcount - 1
            If cboProviders.List(i) = cboProviders.Text Then
                If cboProviders.ItemData(i) <> 0 Then
                    mlngProviderID = cboProviders.ItemData(i)
                    Read_Messages
                    Exit For
                End If
            End If
        Next i
    End If
    Exit Sub
End Sub

Private Sub cboProviders_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strList As String
    
    If Len(cboProviders.Text) > 2 Then
        strList = cboProviders.Text
        cboProviders.ListIndex = SendMessage(cboProviders.Hwnd, LB_FINDSTRING, -1, _
        ByVal strList)
    End If
    Exit Sub
End Sub


Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub


Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
        centerForm Me
        Get_Staff
        If mstrMessageType = vbNullString Then
            Format_Messages
            Read_Messages
        Else
            Format_Assignment
            Read_Assignment
        End If
        centerForm Me
    End If
    Exit Sub
End Sub

Private Sub Form_Load()
Dim sValue As String
Dim Idx As Long
Dim Arr

    mFirst = True
    mLoadPatient = False
    mbooByPass = False
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)

  On Error Resume Next

    Unload Me
    Exit Sub
End Sub

Private Sub grdMessages_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long

   mlngMessRow = Row
   mlngMessCol = Col
   If mlngMessRow < 1 Then Exit Sub

    With grdMessages
        For lngRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = lngRow
                .Row2 = lngRow
                .BackColor = NormalColor
            .BlockMode = False
        Next lngRow
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = mlngMessRow
            .Row2 = mlngMessRow
            .BackColor = grdCurItem
        .BlockMode = False
         If .CellTag = "Inquiry" Then
             mstrPatContact = "Contact"
         Else
             mstrPatContact = "Phone"
         End If
         mlngMessageID = 0
         mlngMessageID = .GetRowItemData(mlngMessRow)
         If mlngMessageID > 0 Then
             frmTelephone.Show vbModal
             Format_Messages
             Read_Messages
         End If
    End With
    Exit Sub
End Sub


