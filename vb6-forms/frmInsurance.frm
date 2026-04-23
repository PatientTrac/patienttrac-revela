VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmApptMaint 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Appointment Types"
   ClientHeight    =   4290
   ClientLeft      =   9900
   ClientTop       =   3420
   ClientWidth     =   6615
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
   Icon            =   "frmInsurance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6615
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   345
      Left            =   3405
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
      Picture         =   "frmInsurance.frx":08CA
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInsurance.frx":12DC
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdTerms 
      Height          =   3045
      Left            =   180
      TabIndex        =   1
      Top             =   690
      Width           =   6135
      _Version        =   458752
      _ExtentX        =   10821
      _ExtentY        =   5371
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
      ShadowDark      =   16761024
      ShadowText      =   8388608
      SpreadDesigner  =   "frmInsurance.frx":162E
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   100
      ScrollBarTrack  =   1
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdRemoveRow 
      Height          =   315
      Left            =   1485
      TabIndex        =   2
      Top             =   300
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmInsurance.frx":18EE
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Remove"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInsurance.frx":1E88
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdGridRow 
      Height          =   315
      Left            =   210
      TabIndex        =   3
      Top             =   300
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmInsurance.frx":21DA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInsurance.frx":2774
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdTermSave 
      Height          =   345
      Left            =   2265
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
      Picture         =   "frmInsurance.frx":2AC6
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Done"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInsurance.frx":2E60
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdUndue 
      Height          =   315
      Left            =   2790
      TabIndex        =   5
      Top             =   300
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmInsurance.frx":31B2
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Undo"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmInsurance.frx":374C
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmApptMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfirst As Boolean
Private mlngGridRow As Long

Private Sub Add_Row()
Dim lngRow As Long
Dim lngCol As Long
Dim strList As String

    lngRow = grdTerms.MaxRows + 1
    grdTerms.MaxRows = lngRow
    Format_CellTypeEdit grdTerms, lngRow, 1, "1", "", "TypeHAlignCenter", "APPT_CODE"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Format_CellTypeEdit grdTerms, lngRow, 2, "1", "", "TypeHAlignCenter", "TYPE_DESC"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Format_CellTypeEdit grdTerms, lngRow, 3, "1", "", "TypeHAlignCenter", "APPT_TIME"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    strList = "New Patient" & Chr(9) & "Established Patient" & Chr(9) & "Business"
    Format_CellTypeCombo grdTerms, lngRow, 4, "1", strList, "TypeHAlignCenter", "VISIT_TYPE"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Exit Sub
End Sub

Private Function Delete_Row()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim lngPrimKey As Long
Dim varCellValue

    On Error GoTo Error_Delete_Row
    Screen.MousePointer = vbHourglass
    With grdTerms
        If mlngGridRow > 0 Then
            .GetText 1, lngRow, varCellValue
            lngPrimKey = .GetRowItemData(mlngGridRow)
            If lngPrimKey > 0 Then
                If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
                strSQL = "UPDATE APPOINTMENT_TYPE ORDER SET WHERE PURGE = 'True' WHERE APPT_TYPE_ID = " & lngPrimKey
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
                    lngSpan = CLng(vstrSpan)
                    If lngSpan > 0 Then
                        .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                    End If
                End If
              End If
              If .ROWHEIGHT(vlngRow) < CellNote Then .ROWHEIGHT(vlngRow) = CellNote
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
vstrString As String, vstrHAligh As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
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
                    lngSpan = CLng(vstrSpan)
                    If lngSpan > 0 Then
                        .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                    End If
                End If
              End If
              If vlngBackColor > 0 Then
                 .BackColor = vlngBackColor
              End If
              .TypeHAlign = TypeHAlignCenter
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
                    lngSpan = CLng(vstrSpan)
                    If lngSpan > 0 Then
                        .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                    End If
                End If
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
            CellNote = .MaxTextCellHeight
            If .ROWHEIGHT(vlngRow) < CellNote Then .ROWHEIGHT(vlngRow) = CellNote + 7
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

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            If vlngRowHeight > 0 Then .ROWHEIGHT(vlngRow) = vlngRowHeight
            .CellType = CellTypePicture
            If FileExists(vstrPhotoPath) Then
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
                lngSpan = CLng(vstrSpan)
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
            If .ROWHEIGHT(vlngRow) < CellNote Then .ROWHEIGHT(vlngRow) = CellNote
            If lngBackColor > 0 Then .BackColor = lngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        .BlockMode = False
    End With
End Function
Private Sub Add_Spacer(vobjGrid As Object, Optional vlngRow As Long)
Dim iRow As Long

    With vobjGrid
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
            .ROWHEIGHT(vlngRow) = 3
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
        .MaxCols = 4
        .ColHeaderRows = 2
        .ColWidth(1) = 15
        Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "ACRYONYM", "TypeHAlignCenter", "APPT_CODE", "", HeaderColor
        grdTerms.SetCellBorder 1, SpreadHeader, 1, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ColWidth(2) = 15
        Format_CellTypeEdit grdTerms, SpreadHeader, 2, "1", "DESCRIPTION", "TypeHAlignCenter", "TYPE_DESC", "", HeaderColor
        grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ColWidth(3) = 15
        Format_CellTypeEdit grdTerms, SpreadHeader, 3, "1", "TIME VALUE", "TypeHAlignCenter", "APPT_TIME", "", HeaderColor
        grdTerms.SetCellBorder 3, SpreadHeader, 3, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ColWidth(4) = 15
        Format_CellTypeEdit grdTerms, SpreadHeader, 5, "1", "PATIENT TYPE", "TypeHAlignCenter", "VISIT_TYPE", "", HeaderColor
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
    strSQL = "SELECT * FROM APPOINTMENT_TYPE ORDER BY APPT_CODE WHERE PURGE = 'False'"
    Set rsCodes = cnCodes.Execute(strSQL)
    Do Until rsCodes.EOF
        If rsCodes.Fields("APPT_CODE") <> vbNullString Then
            Add_Row
            Set Flds = rsCodes.Fields
            For Each Fld In Flds
                With grdTerms
                    lngRow = grdTerms.MaxRows
                    .SetRowItemData lngRow, rsCodes.Fields("APPT_TYPE_ID")
                    Select Case Fld.Name
                        Case "APPT_CODE"
                                Format_CellTypeEdit grdTerms, lngRow, 1, "1", Fld.Value, "TypeHAlignCenter", "APPT_CODE"
                                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Case "TYPE_DESC"
                                Format_CellTypeEdit grdTerms, lngRow, 2, "1", Fld.Value, "TypeHAlignCenter", "TYPE_DESC"
                                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Case "APPT_TIME"
                                Format_CellTypeEdit grdTerms, lngRow, 3, "1", "", "TypeHAlignCenter", "APPT_TIME"
                                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Case "VISIT_TYPE"
                                Set_GridCurSel grdTerms, Fld.Value, lngCol, lngRow, "1", "VISIT_TYPE"
                    End Select
                End With
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
Private Function Set_GridCurSel(vobjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
Dim Idx As Integer
Dim booData As Boolean
Dim strSpan As String
Dim strTag As String
Dim strCellNote As String
Dim lngCol As Long
Dim lngSpan As Long

    With vobjGrid
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
                lngSpan = CLng(vstrSpan)
                If lngSpan > 0 Then
                    .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                End If
            End If
        End If
        If vlngBackColor > 0 Then .BackColor = vlngBackColor
        If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
    .ReDraw = True
    End With
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
Dim varCellValue As String
Dim lngPrimKey As Long

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
            .GetText 1, lngRow, varCellValue: strCode = Trim(varCellValue)
            .GetText 2, lngRow, varCellValue: strDesc = Trim(varCellValue)
            .GetText 3, lngRow, varCellValue: strTime = Trim(varCellValue)
            .GetText 4, lngRow, varCellValue: strType = Trim(varCellValue)
            If strCode <> vbNullString Then
                strSQL = "SELECT APPT_TYPE_ID FROM APPOINTMENT_TYPE WHERE APPT_CODE = '" & strCode & "'"
                Set rsApptType = cnApptType.Execute(strSQL)
                If Not rsApptType.EOF Then
                    If rsApptType.Fields("APPT_TYPE_ID") <> vbNullString Then lngPrimKey = rsApptType.Fields("APPT_TYPE_ID")
                End If
                If lngPrimKey = 0 Then
                    strSQL = "INSERT INTO APPOINTMENT_TYPE (APPT_CODE, TYPE_DESC, APPT_TIME, VISIT_TYPE) " _
                        & "values ('" & strCode & "','" & strDesc & "','" & strTime & "','" & strType & "')"
                    Set rsApptType = cnApptType.Execute(strSQL)
                Else
                    strSQL = "UPDATE APPOINTMENT_TYPE SET APPT_CODE = '" & strCode & "', TYPE_DESC = '" & strDesc & "'," _
                        & "APPT_TIME = '" & strTime & "', VISIT_TYPE = '" & strType & "' WHERE APPT_TYPE_ID = " & lngPrimKey
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

End Sub


Private Sub cmdTermSave_Click()
    Write_ApptType
    grdTerms.ClearRange 1, SpreadHeader, grdTerms.MaxCols, grdTerms.MaxRows, False
    Unload Me
    Exit Sub
End Sub


Private Sub Form_Activate()
    If mfirst Then mfirst = False
    Format_Grid
    Read_Appt_Type
    Exit Sub
End Sub

Private Sub Form_Load()
    mfirst = True
    Exit Sub
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
    mstrTerms = vbNullString
    frmListPatient.mstrListName = vbNullString
  Exit Sub
End Sub


Private Sub grdTerms_Advance(ByVal AdvanceNext As Boolean)

End Sub


Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
    mlngGridRow = Row
    
    If Row >= 0 Then
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
             .BackColor = grdCurItem
             .SetCellBorder 1, Row, .MaxCols, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
         .BlockMode = False
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


