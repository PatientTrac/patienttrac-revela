VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmGridList 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   6510
   ClientLeft      =   9465
   ClientTop       =   4170
   ClientWidth     =   12900
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGridList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   12900
   Begin VB.TextBox cctDefaultNormal 
      ForeColor       =   &H00800000&
      Height          =   945
      Left            =   7080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   270
      Width           =   4965
   End
   Begin VB.TextBox cctSeqNo 
      Alignment       =   2  'Center
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
      Height          =   435
      Left            =   3750
      TabIndex        =   1
      Top             =   270
      Width           =   1275
   End
   Begin VB.TextBox cctVisible 
      Alignment       =   2  'Center
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
      Height          =   435
      Left            =   5220
      TabIndex        =   2
      Top             =   270
      Width           =   1635
   End
   Begin VB.TextBox txtNewColumn 
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
      Height          =   435
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   3465
   End
   Begin PT_XP_Button.PT_XP cmdUp 
      Height          =   555
      Left            =   12330
      TabIndex        =   5
      Top             =   2700
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   979
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
      Picture         =   "frmGridList.frx":08CA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":1044
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDown 
      Height          =   525
      Left            =   12330
      TabIndex        =   6
      Top             =   3390
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   926
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
      Picture         =   "frmGridList.frx":1756
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":1ED0
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdAddList 
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   960
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
      Picture         =   "frmGridList.frx":25E2
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":2B7C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdReplace 
      Height          =   345
      Left            =   1080
      TabIndex        =   11
      Top             =   960
      Width           =   1185
      _ExtentX        =   2090
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
      Picture         =   "frmGridList.frx":2ECE
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Replace"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":3468
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRemove 
      Height          =   345
      Left            =   2310
      TabIndex        =   12
      Top             =   960
      Width           =   1185
      _ExtentX        =   2090
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
      Picture         =   "frmGridList.frx":37BA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Remove"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":3D54
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDone 
      Height          =   465
      Left            =   5768
      TabIndex        =   14
      Top             =   5910
      Width           =   1365
      _ExtentX        =   2408
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
      Picture         =   "frmGridList.frx":40A6
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":4980
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdGrid 
      Height          =   4425
      Left            =   120
      TabIndex        =   15
      Top             =   1380
      Width           =   12015
      _Version        =   458752
      _ExtentX        =   21193
      _ExtentY        =   7805
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
      MaxCols         =   2
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
      SpreadDesigner  =   "frmGridList.frx":55D2
      UserResize      =   0
      Appearance      =   2
      TextTip         =   4
      TextTipDelay    =   100
      ScrollBarTrack  =   1
      ClipboardOptions=   0
   End
   Begin VB.Label lblNormal 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Normal Statement"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7140
      TabIndex        =   13
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label lblSeqNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Row Number"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3750
      TabIndex        =   10
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Visible Property"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5220
      TabIndex        =   9
      Top             =   60
      Width           =   1725
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Row Name"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   1245
   End
   Begin VB.Label lblSort 
      BackStyle       =   0  'Transparent
      Caption         =   "Row Order"
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   12240
      TabIndex        =   7
      Top             =   1950
      Width           =   855
   End
End
Attribute VB_Name = "frmGridList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngtop As Long
Public mlngleft As Long
Public mstrTerms As String
Private mbooByPass As Boolean
Public mstrListName As String
Public mFirst As Boolean
Private mstrItemClick As String
Public mstrFormCaption As String




Private Sub Clear_Fields()
    txtNewColumn.Text = vbNullString
    cctSeqNo.Text = vbNullString
    cctVisible.Text = vbNullString
    cctDefaultNormal.Text = vbNullString
    Exit Sub
End Sub

Private Function Write_Data()
Dim strName As String
Dim lngSeqNo As String
Dim intVisibie As Integer
Dim varCellValue
Dim lngPrimKey As Long
Dim booWrite As Boolean
Dim strNormalState As String
Dim lngRow As Long
Dim cnWrite  As New ADODB.Connection
Dim rsWrite As New ADODB.Recordset
Dim strSQL As String

    On Error GoTo Error_Write_Data
    Screen.MousePointer = vbHourglass
    For lngRow = 1 To grdGrid.MaxRows
        With grdGrid
            lngPrimKey = .GetRowItemData(lngRow)
            .GetText 2, lngRow, varCellValue
            If varCellValue <> vbNullString Then strName = Trim(varCellValue)
            .GetText 3, lngRow, varCellValue
            If varCellValue <> vbNullString Then lngSeqNo = Val(varCellValue)
            .GetText 4, lngRow, varCellValue
            If varCellValue <> vbNullString Then
                If varCellValue = "True" Then
                    intVisibie = 1
                Else
                    intVisibie = 0
                End If
            End If
            .GetText 5, lngRow, varCellValue
            strNormalState = Replace(varCellValue, "'", "''")
        End With
        If cnWrite.State = adStateClosed Then cnWrite.open psConnect(1)
        strSQL = "Get_Physical_Exam_System_ID_System '" & strName & "'"
        Set rsWrite = cnWrite.Execute(strSQL)
        If Not rsWrite.EOF Then
            lngPrimKey = rsWrite.Fields("BODY_SYSTEM_ID")
        Else
            lngPrimKey = 0
        End If
        If lngPrimKey = 0 Then
            strSQL = "Write_Physical_Exam_System '" & Trim(strName) & "'," & lngSeqNo & "," & intVisibie & ",'" & Trim(strNormalState) & "'"
            Set rsWrite = cnWrite.Execute(strSQL)
            strSQL = "Get_Physical_Exam_System_ID_System '" & strName & "'"
            Set rsWrite = cnWrite.Execute(strSQL)
            If Not rsWrite.EOF Then
                lngPrimKey = rsWrite.Fields("BODY_SYSTEM_ID")
            Else
                lngPrimKey = 0
            End If
           If lngPrimKey > 0 Then grdGrid.SetRowItemData lngRow, lngPrimKey
        Else
            strSQL = "Update_Physical_Exam_System '" & Trim(strName) & "'," _
                & lngSeqNo & "," & intVisibie & ",'" & Trim(strNormalState) & "'," & lngPrimKey
            Set rsWrite = cnWrite.Execute(strSQL)
        End If
    Next lngRow
    If cnWrite.State = adStateOpen Then
        cnWrite.Close
        Set cnWrite = Nothing
        Set rsWrite = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Data:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Data", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Renuumber_Grid()
Dim lngRow As Long
Dim iCol%

    With grdGrid
        For lngRow = 1 To .MaxRows
            .SetText 3, lngRow, lngRow
        Next lngRow
    End With
    Write_Data
    Exit Function
End Function


Private Sub Format_Headers()
Dim colID As Variant
Dim lngCnt As Long
Dim lngCol As Long
Dim lngRow As Long
Dim cnExam As New ADODB.Connection
Dim rsExam As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim varHeader
Dim strCount As String
Dim strSQL As String

    On Error GoTo Error_Format_Headers
    Screen.MousePointer = vbHourglass
    With grdGrid
        .ReDraw = False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .FontName = "Times New Roman"
        .GridSolid = False
        .GridShowHoriz = False
        .GridShowVert = False
        .RowHeadersShow = False
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .MaxCols = 5
        .ColHeaderRows = 4
        .ColWidth(1) = 8
        .ColWidth(2) = 17
        .ColWidth(3) = 9
        .ColWidth(4) = 10
        .ColWidth(5) = 52
        Format_CellTypeEdit grdGrid, SpreadHeader, 1, "maxcols", "PHYSICAL EXAM TERMS MAINTENANCE", "TypeHAlignCenter", "BODY_SYSTEM_ID", "", HeaderColor
        grdGrid.SetCellBorder lngCol, SpreadHeader, lngCol, grdGrid.MaxCols, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer grdGrid, SpreadHeader + 1
        lngCol = 1
        Format_CellTypeStatic grdGrid, SpreadHeader + 2, 1, "1", "GRID ID", "TypeHAlignCenter", "BODY_SYSTEM_ID", "", HeaderColor
        grdGrid.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 2
        Format_CellTypeStatic grdGrid, SpreadHeader + 2, lngCol, "1", "BODY SYSTEM", "TypeHAlignLeft", "EXAM_SYSTEM_NAME", "", HeaderColor
        grdGrid.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 3
        Format_CellTypeStatic grdGrid, SpreadHeader + 2, lngCol, "1", "ROW POSITION", "TypeHAlignCenter", "SEQUENCE_NO", "", HeaderColor
        grdGrid.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 4
        Format_CellTypeStatic grdGrid, SpreadHeader + 2, lngCol, "1", "VISIBLE", "TypeHAlignCenter", "VISIBLE", "", HeaderColor
        grdGrid.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngCol = 5
        Format_CellTypeEdit grdGrid, SpreadHeader + 2, lngCol, "1", "EXAM_NORMAL", "TypeHAlignLeft", "EXAM_NORMAL", "", HeaderColor
        grdGrid.SetCellBorder lngCol, SpreadHeader + 2, lngCol, SpreadHeader + 2, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    End With
    If cnExam.State = adStateClosed Then cnExam.open mstrConnectStringTerms
    strSQL = "Load_Physical_Exam_System"
    Set rsExam = cnExam.Execute(strSQL)
    With rsExam
        Do Until .EOF
            lngCnt = lngCnt + 1
            .MoveNext
        Loop
    End With
    lngRow = 0
    strSQL = "Load_Physical_Exam_System"
    Set rsExam = cnExam.Execute(strSQL)
    With rsExam
        Do Until .EOF
            lngRow = lngRow + 1
            grdGrid.MaxRows = lngRow
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                        Select Case UCase(Fld.Name)
                            Case "BODY_SYSTEM_ID"
                                Format_CellTypeStatic grdGrid, lngRow, 1, "1", Trim(Fld.Value), "TypeHAlignCenter", "BODY_SYSTEM_ID"
                            Case "EXAM_SYSTEM_NAME"
                                Format_CellTypeStatic grdGrid, lngRow, 2, "1", Trim(Fld.Value), "TypeHAlignLeft", "EXAM_SYSTEM_NAME"
                            Case "SEQUENCE_NO"
                                Format_CellTypeStatic grdGrid, lngRow, 3, "1", Trim(Fld.Value), "TypeHAlignCenter", "SEQUENCE_NO"
                            Case "VISIBLE"
                                Format_CellTypeStatic grdGrid, lngRow, 4, "1", Trim(Fld.Value), "TypeHAlignCenter", "VISIBLE"
                            Case "EXAM_NORMAL"
                                Format_CellTypeEdit grdGrid, lngRow, 5, "1", Trim(Fld.Value), "TypeHAlignLeft", "EXAM_NORMAL"
                        End Select
                    End If
                End If
            Next
            .MoveNext
            grdGrid.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdGrid.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdGrid.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdGrid.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdGrid.SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Loop
    End With
    Add_Spacer grdGrid, SpreadHeader + 3
    grdGrid.ReDraw = True
    If cnExam.State = adStateOpen Then
        cnExam.Close
        Set cnExam = Nothing
        Set rsExam = Nothing
    End If
    Screen.MousePointer = vbDefault
    mbooHasChanged = False
    mbooByPass = False
    grdGrid.LeftCol = 1
    Exit Sub
Error_Format_Headers:
    Write_Error Err.Description, Err.Number, strSQL, "Format_Headers", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
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
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote + 10
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
            If vlngRowHeight > 0 Then .RowHeight(vlngRow) = vlngRowHeight
            .CellType = CellTypePicture
            lngSpan = Val(vstrSpan)
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf lngSpan > 1 Then
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
             End If
            If FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
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
        If vlngBackColor > 0 Then .BackColor = vlngBackColor
        If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
    .ReDraw = True
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
            .BackColor = GridLine
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 1, vlngRow, .MaxCols, vlngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub

Public Sub Reorder_ListBox(lst As Control, bUp As Boolean)
Dim i As Integer
Dim sContent As String
Dim lData As Long
Dim X As Integer

    If lst.ListIndex = -1 Then
        MsgBox "Select an item from the listbox first", vbExclamation
        Exit Sub
    End If
    If bUp Then
        X = 1
    Else
        X = -1
    End If
    i = lst.ListIndex
    sContent = lst.Text
    lData = lst.ItemData(i)
   If (i > 0 And bUp) Or (i < lst.listcount - 1 And Not bUp) Then
    lst.RemoveItem i
    lst.AddItem sContent, i - X
    lst.ItemData(i - X) = lData
    lst.Selected(i - X) = True
   End If
   
End Sub

Private Sub cctVisible_Click()

    Exit Sub
End Sub


Private Sub cctVisible_KeyUp(KeyCode As Integer, Shift As Integer)
    If cctVisible.Text = "t" Or cctVisible.Text = "T" Then
        cctVisible.Text = "True"
    ElseIf cctVisible.Text = "f" Or cctVisible.Text = "F" Then
        cctVisible.Text = "False"
    End If
    Exit Sub
End Sub


Private Sub cmdAddList_Click()
Dim i%
Dim varCellValue
Dim lngRow As Long
    With grdGrid
        If txtNewColumn.Text = vbNullString Then
            MsgBox "Enter Data in the textboxes first!", vbExclamation
            Exit Sub
        Else
            .MaxRows = .MaxRows + 1
            lngRow = .MaxRows
            Format_CellTypeEdit grdGrid, lngRow, 2, "1", Trim(txtNewColumn.Text), "TypeHAlignCenter", "EXAM_SYSTEM_NAME"
            Format_CellTypeEdit grdGrid, lngRow, 3, "1", Trim(cctSeqNo.Text), "TypeHAlignCenter", "SEQUENCE_NO"
            Format_CellTypeEdit grdGrid, lngRow, 4, "1", Trim(cctVisible.Text), "TypeHAlignCenter", "VISIBLE"
            Format_CellTypeEdit grdGrid, lngRow, 5, "1", Trim(cctDefaultNormal.Text), "TypeHAlignCenter", "EXAM_NORMAL"
            grdGrid.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdGrid.SetCellBorder 2, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdGrid.SetCellBorder 3, lngRow, 3, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdGrid.SetCellBorder 4, lngRow, 4, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            grdGrid.SetCellBorder 5, lngRow, 5, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            Write_Data
        End If
    End With
    Clear_Fields
    Exit Sub
End Sub



Private Sub cmdDone_Click()
     Write_Data
    Unload Me
    Exit Sub
End Sub

Private Sub cmdDown_Click()
Dim lngRow As Integer

    If grdGrid.ActiveRow = 0 Then
        MsgBox "Please select a row first!", vbInformation
    End If
    With grdGrid
        grdGrid.SwapRowRange .SelBlockRow, .SelBlockRow2, grdGrid.ActiveRow + 1
    End With
    Renuumber_Grid
End Sub



Private Sub cmdRemove_Click()
Dim lngRow As Integer

    If grdGrid.ActiveRow = 0 Then
        MsgBox "Select row first!", vbExclamation
        Exit Sub
    End If
    With grdGrid
        .DeleteRows (.ActiveRow), 1
        .MaxRows = .MaxRows - 1
    End With
    Exit Sub
End Sub

Private Sub cmdReplace_Click()
Dim lngRow As Integer
Dim varCellValue

    If grdGrid.ActiveRow = 0 Then
        MsgBox "Select row first!", vbExclamation
        Exit Sub
    End If
    With grdGrid
        .ReDraw = False
        .SetText 2, grdGrid.ActiveRow, Trim(txtNewColumn.Text)
        .SetText 3, grdGrid.ActiveRow, Trim(cctSeqNo.Text)
        .GetText 4, grdGrid.ActiveRow, varCellValue
        If varCellValue = Trim(cctVisible.Text) Then
            For lngRow = grdGrid.ActiveRow + 1 To .MaxRows
                .GetText 3, lngRow, varCellValue
                If varCellValue <= Trim(cctSeqNo.Text) Then
                    varCellValue = Val(varCellValue) + 1
                End If
                .SetText 3, lngRow, varCellValue
            Next lngRow
        End If
        .SetText 4, grdGrid.ActiveRow, Trim(cctVisible.Text)
        .SetText 5, grdGrid.ActiveRow, Trim(cctDefaultNormal.Text)
        
         Write_Data
        .ReDraw = True
    End With
    
End Sub

Private Sub cmdUp_Click()
Dim lngRow As Integer

    If grdGrid.ActiveRow = 0 Then
        MsgBox "Please select a row first!", vbInformation
    End If
    With grdGrid
        grdGrid.SwapRowRange .SelBlockRow, .SelBlockRow2, grdGrid.ActiveRow - 1
    End With
    Renuumber_Grid
End Sub

Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
       centerForm Me
    End If
End Sub

Private Sub Form_Load()
    mFirst = True
    Me.Caption = mstrFormCaption
    Format_Headers
    Exit Sub
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mstrTerms = vbNullString
    frmList.mstrListName = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Exit Sub
End Sub

Private Sub grdGrid_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue
Dim lngRow As Long

    If Row < 1 Then Exit Sub
    With grdGrid
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
            .Row = Row
            .Row2 = Row
            .FontBold = True
            .BackColor = grdCurItem
        .BlockMode = False
        .GetText 2, Row, varCellValue
        If varCellValue <> vbNullString Then txtNewColumn.Text = Trim(varCellValue)
        .GetText 3, Row, varCellValue
        If varCellValue <> vbNullString Then cctSeqNo.Text = Trim(varCellValue)
        .GetText 4, Row, varCellValue
        If varCellValue <> vbNullString Then cctVisible.Text = Trim(varCellValue)
        .GetText 5, Row, varCellValue
        If varCellValue <> vbNullString Then cctDefaultNormal.Text = Trim(varCellValue)
    End With
    Exit Sub
End Sub

