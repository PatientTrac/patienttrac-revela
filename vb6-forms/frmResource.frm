VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmResource 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Resourse Management"
   ClientHeight    =   4290
   ClientLeft      =   8115
   ClientTop       =   3330
   ClientWidth     =   9345
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
   Icon            =   "frmResource.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   9345
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   345
      Left            =   4770
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
      Picture         =   "frmResource.frx":08CA
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmResource.frx":12DC
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdTerms 
      Height          =   3225
      Left            =   180
      TabIndex        =   1
      Top             =   510
      Width           =   8925
      _Version        =   458752
      _ExtentX        =   15743
      _ExtentY        =   5689
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
      SpreadDesigner  =   "frmResource.frx":162E
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   100
      ScrollBarTrack  =   1
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdRemoveRow 
      Height          =   315
      Left            =   1455
      TabIndex        =   2
      Top             =   60
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
      Picture         =   "frmResource.frx":18EE
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Remove"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmResource.frx":1E88
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdGridRow 
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   60
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
      Picture         =   "frmResource.frx":21DA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmResource.frx":2774
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdTermSave 
      Height          =   345
      Left            =   3630
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
      Picture         =   "frmResource.frx":2AC6
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Done"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmResource.frx":2E60
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdUndue 
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   60
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
      Picture         =   "frmResource.frx":31B2
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Undo"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmResource.frx":374C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdProvGrp 
      Height          =   435
      Left            =   7530
      TabIndex        =   6
      ToolTipText     =   "Manage Provider Types"
      Top             =   30
      Width           =   1515
      _ExtentX        =   2672
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
      Picture         =   "frmResource.frx":3A9E
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Provider Type"
      DepthEvent      =   1
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmResource.frx":4038
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Private mlngGridRow As Long

Private Sub Add_Row()
Dim lngRow As Long
Dim lngCol As Long
Dim strList As String

    lngRow = grdTerms.MaxRows + 1
    grdTerms.MaxRows = lngRow
    lngCol = 1
    Format_CellTypeEdit grdTerms, lngRow, lngCol, "2", "", "TypeHAlignCenter", "RESOURCE_NAME"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 3
    Format_CellTypeEdit grdTerms, lngRow, 3, "1", " ", "TypeHAlignCenter", "RESOURCE_DESC "
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 4
    strList = Get_GridTerms("sd_Provider_ProviderType")
    Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", strList, "TypeHAlignCenter", "PRACTICE_GROUP"
    grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 5
    Format_CellTypeEdit grdTerms, lngRow, 5, "1", " ", "TypeHAlignCenter", "SEQUENCE_NO"
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
                strSQL = "UPDATE APPOINTMENT_RESOURCE ORDER SET WHERE PURGE = 'True' WHERE RESOURCE_ID = " & lngPrimKey
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
                    lngSpan = Val(vstrSpan)
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
Public Function Get_GridTerms(vstrSubDomain As String) As String
Dim strSQL As String
Dim cnTerms As New ADODB.Connection
Dim rsTerms As New ADODB.Recordset


    On Error GoTo Error_Get_Appointment_Types
    Screen.MousePointer = vbHourglass
    clsGridFormat.List = vbNullString
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "SELECT * FROM ATTR_VALID_VALUE WHERE VALID_VALUE_SUBDOM = '" & vstrSubDomain & "' ORDER BY  VALID_VALUE_SEQ_NBR"
    Set rsTerms = cnTerms.Execute(strSQL)
    Do Until rsTerms.EOF
        
        If Get_GridTerms = vbNullString Then
            Get_GridTerms = rsTerms.Fields("ATTRBTE_VALID_VALUE")
        Else
            Get_GridTerms = Get_GridTerms & Chr(9) & rsTerms.Fields("ATTRBTE_VALID_VALUE")
        End If
        rsTerms.MoveNext
    Loop
    clsGridFormat.List = Get_GridTerms
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_Appointment_Types:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Appointment_Types", App.EXEName & "." & "ModConst"
    Exit Function
    Resume
End Function
Private Sub Format_Grid()

    Screen.MousePointer = vbHourglass
    grdTerms.ProcessTab = True
    grdTerms.FontBold = True
    With grdTerms
        .ReDraw = False
        grdTerms.ClearRange 1, 1, grdTerms.MaxCols, grdTerms.MaxRows, False
        .MaxCols = 5
        .ColHeaderRows = 2
        .ColWidth(1) = 15
        .ColWidth(2) = 10
        Format_CellTypeEdit grdTerms, SpreadHeader, 1, "2", "RESOURCE NAME", "TypeHAlignCenter", "RESOURCE_NAME", "", HeaderColor
        grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ColWidth(3) = 15
        Format_CellTypeEdit grdTerms, SpreadHeader, 3, "1", "DESCRIPTION", "TypeHAlignCenter", "RESOURCE_DESC", "", HeaderColor
        grdTerms.SetCellBorder 3, SpreadHeader, 3, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ColWidth(4) = 15
        Format_CellTypeEdit grdTerms, SpreadHeader, 4, "1", "USER GROUP", "TypeHAlignCenter", "PRACTICE_GROUP", "", HeaderColor
        grdTerms.SetCellBorder 4, SpreadHeader, 4, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ColWidth(5) = 15
        Format_CellTypeEdit grdTerms, SpreadHeader, 5, "1", "SEQUENCE NO", "TypeHAlignCenter", "SEQUENCE_NO", "", HeaderColor
        grdTerms.SetCellBorder 5, SpreadHeader, 5, SpreadHeader, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        Add_Spacer grdTerms, SpreadHeader + 1
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Public Function Read_Resource()
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long

    On Error GoTo Error_Read_Resource
    Screen.MousePointer = vbHourglass
    If cnCodes.State = adStateClosed Then cnCodes.open psConnect(1)
    strSQL = "SELECT * FROM APPOINTMENT_RESOURCE WHERE PURGE = 'False' ORDER BY APPT_CODE"
    Set rsCodes = cnCodes.Execute(strSQL)
    Do Until rsCodes.EOF
        If rsCodes.Fields("APPT_CODE") <> vbNullString Then
            Add_Row
            Set Flds = rsCodes.Fields
            For Each Fld In Flds
                With grdTerms
                    lngRow = grdTerms.MaxRows
                    .SetRowItemData lngRow, rsCodes.Fields("RESOURCE_ID")
                    Select Case Fld.Name
                        Case "APPT_CODE"
                                lngCol = 1
                                Format_CellTypeEdit grdTerms, lngRow, lngCol, "2", Fld.Value, "TypeHAlignCenter", "APPT_CODE"
                                grdTerms.SetCellBorder lngCol, lngRow, lngCol + 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Case "RESOURCE_DESC"
                                lngCol = 3
                                Set_GridCurSel grdTerms, Fld.Value, lngCol, lngRow, "1", "RESOURCE_DESC"
                                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Case "PRACTICE_GROUP"
                                lngCol = 4
                                Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", Fld.Value, "PRACTICE_GROUP"
                                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        Case "SEQUENCE_NO"
                                lngCol = 5
                                Set_GridCurSel grdTerms, Fld.Value, lngCol, lngRow, "1", "SEQUENCE_NO"
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
Error_Read_Resource:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Resource", App.EXEName & " " & Me.Name
    Exit Function
    Resume
End Function
Private Function Set_GridCurSel(vObjGrid As Object, vstrData As String, vlngCol As Long, vlngRow As Long, Optional vstrSpan As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
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
        If vlngBackColor > 0 Then .BackColor = vlngBackColor
        If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
    .ReDraw = True
    End With
End Function
Private Function Write_Resource()
Dim strSQL As String
Dim cnResource As New ADODB.Connection
Dim rsResource As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRow As Long
Dim lngCol As Long
Dim strName As String
Dim strDesc As String
Dim strGrp As String
Dim strSeq As String
Dim varCellValue
Dim lngPrimKey As Long

    On Error GoTo Error_Write_Resource
    Screen.MousePointer = vbHourglass
    If cnResource.State = adStateClosed Then cnResource.open psConnect(0)
    With grdTerms
        For lngRow = 1 To .MaxRows
            lngPrimKey = 0
            strName = vbNullString
            strDesc = vbNullString
            strGrp = vbNullString
            strSeq = vbNullString
            .GetText 1, lngRow, varCellValue: strName = Trim(varCellValue)
            .GetText 3, lngRow, varCellValue: strDesc = Trim(varCellValue)
            .GetText 4, lngRow, varCellValue: strGrp = Trim(varCellValue)
            .GetText 5, lngRow, varCellValue: strSeq = Trim(varCellValue)
            If strName <> vbNullString Then
                strSQL = "SELECT RESOURCE_ID FROM APPOINTMENT_RESOURCE WHERE RESOURCE_NAME = '" & strName & "'"
                Set rsResource = cnResource.Execute(strSQL)
                If Not rsResource.EOF Then
                    If rsResource.Fields("RESOURCE_ID") <> vbNullString Then lngPrimKey = rsResource.Fields("RESOURCE_ID")
                End If
                If lngPrimKey = 0 Then
                    strSQL = "INSERT INTO APPOINTMENT_RESOURCE (RESOURCE_NAME, RESOURCE_DESC, PRACTICE_GROUP, SEQUENCE_NO) " _
                        & "values ('" & strName & "','" & strDesc & "','" & strGrp & "','" & strSeq & "')"
                    Set rsResource = cnResource.Execute(strSQL)
                Else
                    strSQL = "UPDATE APPOINTMENT_RESOURCE SET APPT_CODE = '" & strName & "', RESOURCE_DESC = '" & strDesc & "'," _
                        & "PRACTICE_GROUP = '" & strGrp & "', SEQUENCE_NO = '" & strSeq & "' WHERE RESOURCE_ID = " & lngPrimKey
                    Set rsResource = cnResource.Execute(strSQL)
                End If
            End If
        Next lngRow
    End With
    If cnResource.State = adStateOpen Then
        cnResource.Close
        Set cnResource = Nothing
        Set rsResource = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Resource:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Resource", App.EXEName & " " & Me.Name
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

Private Sub cmdProvGrp_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Provider Types"
    frmNurseTemplates.fType = "Provider Types"
    frmNurseTemplates.fSubDomain = "sd_Provider_ProviderType"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
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
        strSQL = "UPDATE APPOINTMENT_RESOURCE SET PURGE = 'True' WHERE RESOURCE_ID = " & lngPrimKey
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
    Write_Resource
    grdTerms.ClearRange 1, SpreadHeader, grdTerms.MaxCols, grdTerms.MaxRows, False
    Unload Me
    Exit Sub
End Sub


Private Sub cmdUndue_Click()
Dim lngPrimKey As Long
Dim strSQL As String
Dim cnCodes As New ADODB.Connection
Dim rsCodes As New ADODB.Recordset
Dim strDate As String

    On Error GoTo Error_Remove
    Screen.MousePointer = vbDefault
'    strDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
'
'    If lngPrimKey > 0 Then
'        If cnCodes.State = adStateClosed Then cnCodes.Open psConnect(1)
'        strSQL = "UPDATE APPOINTMENT_RESOURCE SET PURGE = 'False' WHERE RESOURCE_ID = " & lngPrimKey & " AND " _
'            & "UPDATE_DATE = '" & strDate & "'"
'        Set rsCodes = cnCodes.Execute(strSQL)
'    End If
'
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Remove:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Remove", App.EXEName & " " & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
        Format_Grid
        Read_Resource
        centerForm Me
    End If
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


Private Sub grdTerms_ComboCloseUp(ByVal Col As Long, ByVal Row As Long, ByVal SelChange As Integer)
    grdTerms.Col = 1
    
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


