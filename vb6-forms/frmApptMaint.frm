VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmApptMaint 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Appointment Types"
   ClientHeight    =   4290
   ClientLeft      =   8550
   ClientTop       =   2775
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "Times"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frmApptMaint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   11925
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   345
      Left            =   6060
      TabIndex        =   0
      ToolTipText     =   "Save, Restore, Close"
      Top             =   3840
      Width           =   945
      _ExtentX        =   1667
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
      Picture         =   "frmApptMaint.frx":08CA
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmApptMaint.frx":12DC
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdTerms 
      Height          =   3165
      Left            =   180
      TabIndex        =   1
      Top             =   570
      Width           =   11565
      _Version        =   458752
      _ExtentX        =   20399
      _ExtentY        =   5583
      _StockProps     =   64
      ColHeaderDisplay=   0
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      DInformActiveRowChange=   0   'False
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
      MaxCols         =   5
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmApptMaint.frx":162E
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   100
      ScrollBarTrack  =   1
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdRemoveRow 
      Height          =   345
      Left            =   1485
      TabIndex        =   2
      Top             =   60
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmApptMaint.frx":1912
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Remove"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmApptMaint.frx":1EAC
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdGridRow 
      Height          =   345
      Left            =   210
      TabIndex        =   3
      Top             =   60
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmApptMaint.frx":21FE
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmApptMaint.frx":2798
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdTermSave 
      Height          =   345
      Left            =   4920
      TabIndex        =   4
      Top             =   3840
      Width           =   945
      _ExtentX        =   1667
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
      Picture         =   "frmApptMaint.frx":2AEA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Done"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmApptMaint.frx":2E84
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdUndue 
      Height          =   345
      Left            =   2790
      TabIndex        =   5
      Top             =   60
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmApptMaint.frx":31D6
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Undo"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmApptMaint.frx":3770
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmApptMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Private mlngGridRow As Long

Private Function Load_CPTCode()
Dim i%
Dim strSQL As String
Dim cnCode As New ADODB.Connection
Dim rsCode As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field

    On Error GoTo Error_Load_CPTCode
    Screen.MousePointer = vbHourglass
    
    clsGridFormat.List = vbNullString
    clsGridFormat.Text = vbNullString
    
    If cnCode.State = adStateClosed Then cnCode.open psConnect(1)
    strSQL = "Select CPT_NAME FROM CPT_TEMPLATE ORDER BY PROVIDER"
    Set rsCode = cnCode.Execute(strSQL)
    With rsCode
        Do Until .EOF
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "CPT_NAME"
                                If clsGridFormat.List = vbNullString Then
                                    clsGridFormat.List = Trim(Fld.Value)
                                Else
                                    clsGridFormat.List = clsGridFormat.List & Chr(9) & Trim(Fld.Value)
                                End If
                            Case "PROVIDER"
                                If clsGridFormat.Text = vbNullString Then
                                    clsGridFormat.Text = Trim(Fld.Value)
                                Else
                                    clsGridFormat.Text = clsGridFormat.Text & Chr(9) & Trim(Fld.Value)
                                End If
                        End Select
                    End If
                End If
            
            Next
            .MoveNext
        Loop
    End With
    If cnCode.State = adStateOpen Then
        Set rsCode = Nothing
        cnCode.Close
        Set cnCode = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_CPTCode:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_CPTCode", App.EXEName & "." & "ModPatientData"
    Exit Function
    Resume
End Function
Private Sub Add_Row()
Dim lngRow As Long
Dim lngCol As Long
Dim strList As String

    lngRow = grdTerms.MaxRows + 1
    grdTerms.MaxRows = lngRow
    lngCol = 1
    Format_CellTypeEdit grdTerms, lngRow, lngCol, "2", "", "TypeHAlignCenter", "APPT_CODE"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 3
    Load_CPTCode
    Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", clsGridFormat.List, "TypeHAlignCenter", "TYPE_DESC"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 4
    Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "", "TypeHAlignCenter", "APPT_TIME"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 5
    strList = "New Patient" & Chr(9) & "Established Patient" & Chr(9) & "Business"
    Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", strList, "TypeHAlignCenter", "VISIT_TYPE"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 6
    Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", "", "TypeHAlignCenter", "PROVIDER_TITLE"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Exit Sub
End Sub

Private Function Delete_Row()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim lngPrimKey As Long
Dim varCellValue
Dim lngRow As Long

    On Error GoTo Error_Delete_Row
    Screen.MousePointer = vbHourglass
    With grdTerms
        If mlngGridRow > 0 Then
            .GetText 1, mlngGridRow, varCellValue
            lngPrimKey = .GetRowItemData(mlngGridRow)
            If lngPrimKey > 0 Then
                If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
                strSQL = "Remove_Appointment_Type " & lngPrimKey
                Set rsCodes = cnCodes.Execute(strSQL)
                MsgBox varCellValue & " " & "has been purged", vbInformation
            End If
        End If
    End With
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Delete_Row:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Delete_Row", App.EXEName & " " & Me.Name
    Exit Function
    Resume
End Function

Private Function Format_CellTypeBackColor(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, lngBackColor As Long, _
Optional vstrCellTag As String, Optional vstrCellNote As String)
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
              .Text = vstrText
              If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
              If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
              CellNote = .MaxTextCellHeight
              If vstrSpan <> vbNullString Then
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                    .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                Else
                    lngSpan = Val(vstrSpan)
                    If lngSpan > 0 Then
                        .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                    End If
                End If
              End If
              If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
              .TypeVAlign = TypeVAlignCenter
                If vstrHAligh = "TypeHAlignCenter" Then
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .TypeHAlign = TypeHAlignLeft
                End If
                .BackColor = lngBackColor
          .BlockMode = False

    End With
End Function
Private Function Format_CellTypeCombo(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, _
vstrString As String, vstrHAligh As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, _
Optional vstrCellNote As String, Optional vstrHAlign As String)
Dim CellNote As Long
Dim lngSpan As Long

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
              If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
              If vstrCellTag <> vbNullString Then
                .CellTag = vstrCellTag
              End If
              If vstrSpan <> vbNullString Then
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                    .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                Else
                    lngSpan = Val(vstrSpan)
                    If lngSpan > 0 Then
                        .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                    End If
                End If
              End If
              If vlngBackColor > 0 Then
                 .BackColor = vlngBackColor
              End If
              If Trim(vstrHAlign) = vbNullString Then
                .TypeHAlign = TypeHAlignCenter
              Else
                .TypeHAlign = vstrHAlign
              End If
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
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
           ' .TypePictStretch = True
            .TypePictCenter = True
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            'grdPic.SetCellBorder vlngCol, vlngRow, vlngCol, vlngRow, CellBorderIndexOutline, &H800000, CellBorderStyleSolid
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
Private Sub Add_Spacer(vObjGrid As Object, Optional vlngRow As Long)
Dim iRow As Long

    With vObjGrid
        If vlngRow >= 0 Then
             vlngRow = vlngRow + 1
            .MaxRows = vlngRow
            iRow = vlngRow
        Else
            iRow = vlngRow
        End If
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

Private Sub Format_Grid()

    Screen.MousePointer = vbHourglass
    grdTerms.ProcessTab = True
    grdTerms.FontBold = True
    With grdTerms
        .ReDraw = False
        grdTerms.ClearRange 1, 1, grdTerms.MaxCols, grdTerms.MaxRows, False
        .MaxCols = 6
        .ColHeaderRows = 2
        .ColWidth(1) = 10
        .ColWidth(2) = 10
        Format_CellTypeEdit grdTerms, SpreadHeader, 1, "2", "ACRYONYM", "TypeHAlignCenter", "APPT_CODE", "", HeaderColor
        grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ColWidth(3) = 32.5
        Format_CellTypeEdit grdTerms, SpreadHeader, 3, "1", "DESCRIPTION", "TypeHAlignCenter", "TYPE_DESC", "", HeaderColor
        grdTerms.SetCellBorder 3, SpreadHeader, 3, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ColWidth(4) = 12
        Format_CellTypeEdit grdTerms, SpreadHeader, 4, "1", "TIME VALUE", "TypeHAlignCenter", "APPT_TIME", "", HeaderColor
        grdTerms.SetCellBorder 4, SpreadHeader, 4, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ColWidth(5) = 18
        Format_CellTypeEdit grdTerms, SpreadHeader, 5, "1", "PATIENT TYPE", "TypeHAlignCenter", "VISIT_TYPE", "", HeaderColor
        grdTerms.SetCellBorder 5, SpreadHeader, 5, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ColWidth(6) = 10
        Format_CellTypeEdit grdTerms, SpreadHeader, 6, "1", "PROVIDER TYPE", "TypeHAlignCenter", "PROVIDER_TITLE", "", HeaderColor
        grdTerms.SetCellBorder 5, SpreadHeader, 5, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Add_Spacer grdTerms, SpreadHeader + 1
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Public Function Read_Appt_Type()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long

    On Error GoTo Error_Read_Appt_Type
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    strSQL = "Get_Appointment_Type"
    Set rsCodes = cnCodes.Execute(strSQL)
    Do Until rsCodes.EOF
        If rsCodes.Fields("APPT_CODE") <> vbNullString Then
            Add_Row
            Set Flds = rsCodes.Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    With grdTerms
                        lngRow = grdTerms.MaxRows
                        .SetRowItemData lngRow, rsCodes.Fields("APPT_TYPE_ID")
                        Select Case Fld.Name
                            Case "APPT_CODE"
                                    lngCol = 1
                                    Format_CellTypeEdit grdTerms, lngRow, lngCol, "2", Fld.Value, "TypeHAlignCenter", "APPT_CODE"
                                    grdTerms.SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            Case "TYPE_DESC"
                                    lngCol = 3
                                    Set_GridCurSel grdTerms, Fld.Value, lngCol, lngRow, "1", "TYPE_DESC", 0, "", "TypeHAlignLeft"
                                    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            Case "APPT_TIME"
                                    lngCol = 4
                                    Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", Fld.Value, "APPT_TIME"
                                    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            Case "VISIT_TYPE"
                                    lngCol = 5
                                    Set_GridCurSel grdTerms, Fld.Value, lngCol, lngRow, "1", "VISIT_TYPE"
                                    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                            Case "PROVIDER_TITLE"
                                    lngCol = 6
                                    Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", Fld.Value, "TypeHAlignCenter", "PROVIDER_TITLE"
                                    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        End Select
                    End With
                End If
            Next
        End If
        rsCodes.MoveNext
    Loop
    If cnCodes.State = adStateOpen Then
        cnCodes.Close
        Set cnCodes = Nothing
        Set rsCodes = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Appt_Type:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Appt_Type", App.EXEName & " " & Me.Name
    Exit Function
    Resume
End Function
Private Function Set_GridCurSel(vObjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String, Optional vstrHAlign As String)
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
                If Trim(vstrData) = Trim(.TypeComboBoxString) Then
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
        If Trim(vstrHAlign) = vbNullString Then
          .TypeHAlign = TypeHAlignCenter
        ElseIf vstrHAlign = "TypeHAlignLeft" Then
          .TypeHAlign = TypeHAlignLeft
        End If
        If vlngBackColor > 0 Then .BackColor = vlngBackColor
        If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
    .ReDraw = True
    End With
End Function
Private Function LookUpProvider(vstrCPTName As String)
Dim i%
Dim strSQL As String
Dim cnCode As New ADODB.Connection
Dim rsCode As New ADODB.Recordset

    On Error GoTo Error_LookUpProvider
    Screen.MousePointer = vbHourglass
    clsGridFormat.Text = vbNullString
    If cnCode.State = adStateClosed Then cnCode.open psConnect(1)
    strSQL = "Get_CPT_Template_Name  '" & vstrCPTName & "'"
    Set rsCode = cnCode.Execute(strSQL)
    With rsCode
        Do Until .EOF
            If (!Provider) <> vbNullString Then
                If clsGridFormat.Text = vbNullString Then
                    clsGridFormat.Text = Trim((!Provider))
                Else
                    clsGridFormat.Text = clsGridFormat.Text & Chr(9) & Trim((!Provider))
                End If
            End If
        rsCode.MoveNext
    Loop
    End With
    If cnCode.State = adStateOpen Then
        Set rsCode = Nothing
        cnCode.Close
        Set cnCode = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUpProvider:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "LookUpProvider", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function





Private Function Write_ApptType()
Dim strSQL As String
Dim cnApptType As New ADODB.Connection
Dim rsApptType As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim strCode As String
Dim strDesc As String
Dim strTime As String
Dim strType As String
Dim varCellValue
Dim lngPrimKey As Long
Dim strTitle As String

    On Error GoTo Error_Write_ApptType
    Screen.MousePointer = vbHourglass
    If cnApptType.State = adStateClosed Then cnApptType.open psConnect(1)
    With grdTerms
        For lngRow = 1 To .MaxRows
            lngPrimKey = 0
            strCode = vbNullString
            strDesc = vbNullString
            strTime = vbNullString
            strType = vbNullString
            strTitle = vbNullString
            .GetText 1, lngRow, varCellValue: strCode = Trim(varCellValue)
            .GetText 3, lngRow, varCellValue: strDesc = Trim(varCellValue)
            .GetText 4, lngRow, varCellValue: strTime = Trim(varCellValue)
            .GetText 5, lngRow, varCellValue: strType = Trim(varCellValue)
            .GetText 6, lngRow, varCellValue: strTitle = Trim(varCellValue)
            If Trim(strTitle) = vbNullString And Trim(strCode) <> vbNullString Then
                MsgBox "Please select appropriate provider type! on Row" & " " & CStr(lngRow) & "!", vbInformation, "PatientTrac Appointment Maintanence"
                Exit Function
            End If
            If strCode <> vbNullString Then
                strSQL = "Get_Appointment_Type_ID '" & strCode & "','" & strTitle & "'"
                Set rsApptType = cnApptType.Execute(strSQL)
                If Not rsApptType.EOF Then
                    If rsApptType.Fields("APPT_TYPE_ID") <> vbNullString Then lngPrimKey = rsApptType.Fields("APPT_TYPE_ID")
                End If
                If lngPrimKey = 0 Then
                    strSQL = "Write_Appointment_Type '" & strCode & "','" & strDesc & "','" & strTime & "','" & strType & "','" & strTitle & "'"
                    Set rsApptType = cnApptType.Execute(strSQL)
                Else
                    strSQL = "Update_Appointment_Type '" & strCode & "','" & strDesc & "','" _
                        & strTime & "','" & strType & "','" & strTitle & "'," & lngPrimKey
                    Set rsApptType = cnApptType.Execute(strSQL)
                End If
            End If
        Next lngRow
    End With
    If cnApptType.State = adStateOpen Then
        cnApptType.Close
        Set cnApptType = Nothing
        Set rsApptType = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_ApptType:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_ApptType", App.EXEName & " " & Me.Name
    Exit Function
    Resume
End Function

Private Sub cmdClose_Click()
    grdTerms.ClearRange 1, SpreadHeader, grdTerms.MaxCols, grdTerms.MaxRows, False
    Unload Me
    Exit Sub
End Sub

Private Sub cmdGridRow_Click()
    Add_Row
    Exit Sub
End Sub

Private Sub cmdRemoveRow_Click()
Dim lngPrimKey As Long
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim strDate As String

    On Error GoTo Error_Remove
    Screen.MousePointer = vbDefault
    strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    If mlngGridRow > 0 Then
        lngPrimKey = grdTerms.GetRowItemData(mlngGridRow)
        grdTerms.DeleteRows mlngGridRow, 1
        grdTerms.MaxRows = grdTerms.MaxRows - 1
    End If
    If lngPrimKey > 0 Then
        If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
        strSQL = "UPDATE APPOINTMENT_TYPE SET PURGE = 'True' WHERE APPT_TYPE_ID = " & lngPrimKey
        Set rsCodes = cnCodes.Execute(strSQL)
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Remove:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Remove", App.EXEName & " " & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub cmdTermSave_Click()
    Write_ApptType
    grdTerms.ClearRange 1, SpreadHeader, grdTerms.MaxCols, grdTerms.MaxRows, False
    Unload Me
    Exit Sub
End Sub


Private Sub Form_Activate()
    If mFirst Then mFirst = False
    Format_Grid
    Read_Appt_Type
    Exit Sub
End Sub

Private Sub Form_Load()
    mFirst = True
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Unload Me
    Exit Sub
End Sub


Private Sub grdTerms_Change(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue
Dim lngRow As Long
Dim lngCol As Long


    With grdTerms
        .GetText Col, Row, varCellValue
            lngCol = Col
            lngRow = Row
        Select Case Col
            Case 1
                Format_CellTypeEdit grdTerms, lngRow, lngCol, "2", CStr(varCellValue), "TypeHAlignCenter", "APPT_CODE"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Case 3
                Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", CStr(varCellValue), "TypeHAlignCenter", "TYPE_DESC"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Case 4
                If Not TextFound("minutes", CStr(CStr(varCellValue)), Len("minutes"), True) Then
                    Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", CStr(varCellValue) & " " & "minutes", "TypeHAlignCenter", "TYPE_DESC"
                    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                End If
        End Select
    End With
    Exit Sub
End Sub

Private Sub grdTerms_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
Dim Idx As Integer
Dim strCPT As String
Dim lngRow As Long
Dim lngCol As Long

    With grdTerms
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString Then
            If .CellTag = "TYPE_DESC" Then
                If .CellType = CellTypeComboBox Then
                    For Idx = 0 To .TypeComboBoxCount
                        .TypeComboBoxIndex = Idx
                        If .TypeComboBoxCurSel = Idx Then
                            strCPT = Trim(.TypeComboBoxString)
                            Exit For
                        End If
                    Next Idx
                End If
            End If
        End If
    End With
    If Trim(strCPT) <> vbNullString Then
        LookUpProvider Trim(strCPT)
        lngCol = 6
        lngRow = Row
        grdTerms.ClearRange lngCol, Row, lngCol, Row, False
        If Not TextFound("chr(9)", clsGridFormat.Text, Len(Chr(9)), True) Then
           Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", clsGridFormat.Text, "TypeHAlignCenter", "PROVIDER_TITLE"
        Else
            Format_CellTypeCombo grdTerms, Row, lngCol, "1", clsGridFormat.Text, "TypeHAlignCenter", "PROVIDER_TITLE"
        End If
        grdTerms.SetCellBorder lngCol, Row, lngCol, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    End If
    Exit Sub
    
End Sub

Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
    mlngGridRow = Row
    
    If Row >= 0 Then
        With grdTerms
            .BlockMode = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = Row
                .Row2 = Row
                 .BackColor = grdCurItem
                 .SetCellBorder 1, Row, .MaxCols, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
             .BlockMode = False
         End With
    End If
End Sub


Private Sub grdTerms_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    mlngGridRow = 0
    If Row <= 0 Then Exit Sub
    With grdTerms
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
             .BackColor = NormalColor
         .BlockMode = False
    End With
    Exit Sub
End Sub


