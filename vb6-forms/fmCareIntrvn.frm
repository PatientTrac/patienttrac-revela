VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form fmMseTerms 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Mental Status Selection"
   ClientHeight    =   5250
   ClientLeft      =   10365
   ClientTop       =   4530
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fmCareIntrvn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   8925
   Begin VB.TextBox cctProv 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   1260
      TabIndex        =   8
      Top             =   60
      Width           =   2595
   End
   Begin VB.ComboBox cboGroup 
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
      ItemData        =   "fmCareIntrvn.frx":058A
      Left            =   5430
      List            =   "fmCareIntrvn.frx":0594
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   30
      Width           =   2925
   End
   Begin PT_XP_Button.PT_XP cmdCloseList 
      Height          =   435
      Left            =   5145
      TabIndex        =   0
      Top             =   4650
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
      Picture         =   "fmCareIntrvn.frx":05BE
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmCareIntrvn.frx":0FD0
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdTerms 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   750
      Width           =   8355
      _Version        =   458752
      _ExtentX        =   14737
      _ExtentY        =   6588
      _StockProps     =   64
      AutoCalc        =   0   'False
      ColHeaderDisplay=   0
      ColsFrozen      =   1
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
      GrayAreaBackColor=   16777215
      GridColor       =   8388608
      MaxCols         =   2
      MaxRows         =   0
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "fmCareIntrvn.frx":1322
      Appearance      =   2
      ClipboardOptions=   0
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   435
      Left            =   2715
      TabIndex        =   2
      Top             =   4620
      Width           =   975
      _ExtentX        =   1720
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
      Picture         =   "fmCareIntrvn.frx":15D7
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmCareIntrvn.frx":1971
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdUp 
      Height          =   495
      Left            =   8430
      TabIndex        =   3
      Top             =   1530
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   873
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
      Picture         =   "fmCareIntrvn.frx":1CC3
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmCareIntrvn.frx":243D
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDown 
      Height          =   465
      Left            =   8430
      TabIndex        =   4
      Top             =   2160
      Width           =   405
      _ExtentX        =   714
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
      Picture         =   "fmCareIntrvn.frx":2B4F
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmCareIntrvn.frx":32C9
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdAll 
      Height          =   315
      Left            =   7110
      TabIndex        =   7
      Top             =   390
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
      Picture         =   "fmCareIntrvn.frx":39DB
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Show All"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmCareIntrvn.frx":3F75
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblVisit 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Provider"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblMessages 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Visit Type"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4170
      TabIndex        =   6
      Top             =   90
      Width           =   1185
   End
End
Attribute VB_Name = "fmMseTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbooByPass As Boolean
Private mFirst As Boolean
Private Function Get_MSE_Terms()
Dim strSQL As String
Dim cnCareIntrvn As New ADODB.Connection
Dim rsCareIntrvn As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strTerms As String
Dim strGroup As String
Dim lngCareID As Long
Dim lngRow As Long
Dim lngCol As Long


    On Error GoTo Error_Get_MSE_Terms
    Screen.MousePointer = vbHourglass
    'Provider; Visit Type?
    If cnCareIntrvn.State = adStateClosed Then cnCareIntrvn.open psConnect(1)
    strSQL = "SELECT *  FROM MSE_HEADER_FORMAT WHERE " _
        & "VISIBLE = 'True' ORDER BY SEQUENCE_NO"
    Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
    With rsCareIntrvn
        Do Until .EOF
          strTerms = vbNullString
          strGroup = vbNullString
          lngCareID = 0
          Set Flds = .Fields
          For Each Fld In Flds
              If Not IsNull(Trim(Fld.Value)) Then
                  If Trim(Fld.Value) <> vbNullString Then
                      Select Case UCase(Fld.Name)
                          Case "HEADERS"
                              strTerms = IsProperCase(Trim(Fld.Value))
'                          Case "CARE_GROUP"
'                              strGroup = UCase(Trim(Fld.Value))
                          Case "TERM_ID"
                              lngCareID = Trim(Fld.Value)
                      End Select
                  End If
              End If
          Next
          .MoveNext
          If Trim(strTerms) <> vbNullString Then
                lngRow = grdTerms.MaxRows + 1
                grdTerms.MaxRows = lngRow
                grdTerms.SetRowItemData lngRow, lngCareID
                lngCol = 1
                Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", strTerms, "TypeHAlignLeft", "INTERVENTION"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", Get_GridTerms("sd_CarePlan_Intrvn_Group"), "TypeHAlignCenter", "GROUP"
                Set_GridCurSel grdTerms, strGroup, lngCol, lngRow, "1", "GROUP"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
          End If
        Loop
    End With
    If cnCareIntrvn.State = adStateOpen Then
        cnCareIntrvn.Close
        Set cnCareIntrvn = Nothing
        Set rsCareIntrvn = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_MSE_Terms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_MSE_Terms", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function
Private Function Format_GridTerms()

    Screen.MousePointer = vbHourglass
    With grdTerms
        .ReDraw = False
       .ClearRange 0, SpreadHeader, grdTerms.MaxRows, grdTerms.MaxCols, False
        .ProcessTab = True
        .FontBold = True
        .ColHeaderRows = 2
        .ColHeadersShow = True
        .RowHeadersShow = False
        .ColWidth(1) = 33
        .ColWidth(2) = 32.5
        Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "MSE TERM", "TypeHAlignCenter", "INTERVENTION", "", HeaderColor
        Format_CellTypeEdit grdTerms, SpreadHeader, 2, "maxcols", "VISIBLE", "TypeHAlignCenter", "GROUP", "", HeaderColor
        grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdTerms
        grdTerms.MaxRows = 0
        .ReDraw = True
    End With
End Function
Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)

    With vObjGrid
        vlngRow = vlngRow
       If .MaxRows < vlngRow Then .MaxRows = vlngRow
        .BlockMode = True
            .Col = 1
            .Col2 = 1
            If vlngRow = -998 Then
                .Row = SpreadHeader + 2
                .Row2 = SpreadHeader + 2
            Else
                .Row = vlngRow
                .Row2 = vlngRow
            End If
            .RowHeight(vlngRow) = 3
            .BackColor = GridLine
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 1, vlngRow, .MaxCols, vlngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub
Private Function Format_CellTypeCombo(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, _
vstrString As String, vstrHAligh As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .ReDraw = False
        .BlockMode = True
              .Col = vlngCol
              .Col2 = vlngCol
              .Row = vlngRow
              .Row2 = vlngRow
              .CellType = CellTypeComboBox
              .TypeMaxEditLen = 32000
              .TypeComboBoxList = vstrString
              .TypeComboBoxWidth = 1
              If vlngBackColor > 0 Then
                 .BackColor = vlngBackColor
              End If
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
              .FontName = "Arial"
              .FontSize = 10
              .FontBold = True
          .BlockMode = False
        .ReDraw = True
    End With
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional vlngBackColor As Long, Optional vlngFontSize As Long)
Dim CellNote As Long
Dim lngSpan As Long

        
    With vGrid
        .ReDraw = False
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeEdit
            .TypeMaxEditLen = 32000
            .FontBold = True
            .FontName = "Arial"
            .FontSize = 10
            If vlngFontSize > 0 Then .FontSize = vlngFontSize
            .TypeEditMultiLine = True   'Make multi-lines
            'If vbooUnderLine Then .FontUnderline = Truer
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
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote + 5
            .TypeVAlign = TypeVAlignCenter
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If vlngBackColor Then .BackColor = vlngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            .EditMode = False
          .BlockMode = False
        .ReDraw = True
    End With
End Function
Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional lngBackColor As Long, Optional vstrVAlign As String)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .ReDraw = False
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeStaticText
            .TypeMaxEditLen = 32000
            .FontName = "Arial"
            .FontSize = 10
            .FontBold = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
            .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
             lngSpan = Val(vstrSpan)
            .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
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
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
            If lngBackColor > 0 Then .BackColor = lngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        .BlockMode = False
        .ReDraw = True
    End With
End Function
Private Function Load_Group_Filter()
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim lngRow As Long
Dim varCellValue
Dim Arr
Dim Idx As Integer
Dim lngCol As Long
Dim strTerms As String
Dim bExit As Boolean

    On Error GoTo Error_Load_Group_Filter
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    cboGroup.Clear
    strSQL = "Get_Attr_Valid_Value_SubDom_Sequence 'sd_CarePlan_Intrvn_Group'"
    Set rsTerms = cnTerms.Execute(strSQL)
    With rsTerms
        Do Until .EOF
            bExit = False
            strTerms = vbNullString
            If (!ATTRBTE_VALID_VALUE) <> vbNullString Then
                strTerms = Trim(!ATTRBTE_VALID_VALUE)
            End If
            .MoveNext
            For Idx = 0 To cboGroup.listcount - 1
                If cboGroup.List(Idx) = strTerms Then
                    bExit = True
                    Exit For
                End If
            Next Idx
            If Trim(strTerms) <> vbNullString And Not bExit Then
                cboGroup.AddItem Trim(strTerms)
            End If
        Loop
    End With
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set rsTerms = Nothing
        Set cnTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Group_Filter:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Load_Group_Filter", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function

Private Function Load_Group_Terms()
Dim strSQL As String
Dim cnCareIntrvn As New ADODB.Connection
Dim rsCareIntrvn As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strTerms As String
Dim strGroup As String
Dim lngCareID As Long
Dim lngRow As Long
Dim lngCol As Long


    On Error GoTo Error_Load_Group_Terms
    If cboGroup = vbNullString Then
        Exit Function
    End If
    grdTerms.MaxRows = 0
    Screen.MousePointer = vbHourglass
    If cnCareIntrvn.State = adStateClosed Then cnCareIntrvn.open psConnect(1)
    strSQL = "SELECT *  FROM CARE_PLAN WHERE " _
        & "SPECIALTY = '" & mstrPractice & "' AND CARE_GROUP = '" & cboGroup & "' ORDER BY SEQUENCE_NO"
    Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
    With rsCareIntrvn
        Do Until .EOF
          strTerms = vbNullString
          strGroup = vbNullString
          lngCareID = 0
          Set Flds = .Fields
          For Each Fld In Flds
              If Not IsNull(Trim(Fld.Value)) Then
                  If Trim(Fld.Value) <> vbNullString Then
                      Select Case UCase(Fld.Name)
                          Case "CARE_ITEM"
                              strTerms = IsProperCase(Trim(Fld.Value))
                          Case "CARE_GROUP"
                              strGroup = UCase(Trim(Fld.Value))
                          Case "CAREPLAN_ID"
                              lngCareID = Trim(Fld.Value)
                      End Select
                  End If
              End If
          Next
          .MoveNext
          If Trim(strTerms) <> vbNullString Then
                lngRow = grdTerms.MaxRows + 1
                grdTerms.MaxRows = lngRow
                grdTerms.SetRowItemData lngRow, lngCareID
                lngCol = 1
                Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", strTerms, "TypeHAlignLeft", "INTERVENTION"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", Get_GridTerms("sd_CarePlan_Intrvn_Group"), "TypeHAlignCenter", "GROUP"
                Set_GridCurSel grdTerms, strGroup, lngCol, lngRow, "1", "GROUP"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
          End If
        Loop
    End With
    If cnCareIntrvn.State = adStateOpen Then
        cnCareIntrvn.Close
        Set cnCareIntrvn = Nothing
        Set rsCareIntrvn = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Group_Terms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Group_Terms", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
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

    If vstrData <> vbNullString Then
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
            If vlngBackColor > 0 Then .BackColor = vlngBackColor
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
            If vstrCellNote <> vbNullString Then .CellNote = Trim(vstrCellNote)
        .ReDraw = True
        End With
    End If
End Function

Private Function Write_MSE_Terms()
Dim strSQL As String
Dim cnCareIntrvn As New ADODB.Connection
Dim rsCareIntrvn As New ADODB.Recordset
Dim lngPrimKey As Long
Dim lngRow As Long
Dim lngCol As Long
Dim strTerms As String
Dim strGroup As String
Dim lngCareID As Long
Dim varCellValue

    On Error GoTo Error_Write_Care_Plan
    Screen.MousePointer = vbHourglass
    With grdTerms
        If cnCareIntrvn.State = adStateClosed Then cnCareIntrvn.open psConnect(1)
        For lngRow = 1 To grdTerms.MaxRows
            strTerms = vbNullString
            strGroup = vbNullString
            lngCareID = 0
            .GetText 1, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                strTerms = Trim(varCellValue)
            End If
            .GetText 2, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                strGroup = Trim(varCellValue)
            End If
            strSQL = "SELECT TERM_ID FROM MSE_HEADER_FORMAT WHERE HEADERS = '" & strTerms & "'"
            Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
            If Not rsCareIntrvn.EOF Then
                lngCareID = rsCareIntrvn.Fields("TERM_ID")
            End If
'    [TERM_ID] [int] IDENTITY(1000,1) NOT NULL,
'    [HEADER] [varchar](max) NULL,
'    [INITIAL_PATIENT] [varchar](50) NULL,
'    [SUBSEQUENT_PATIENT] [varchar](50) NULL,
'    [SEQUENCE_NO] [int] NULL,
'    [PROVIDER_ID] [varchar](50) NULL,
'    [INSERT_DATE] [datetime] NULL,
'    [UPDATE_DATE] [datetime] NULL,
'    [ROWGUID] [uniqueidentifier] ROWGUIDCOL  NOT NULL,
'    [VISIBLE] [varchar](50) NULL,
            If lngCareID = 0 Then
                strSQL = "INSERT INTO MSE_HEADER_FORMAT (HEADERS, VISIBLE, SEQUENCE_NO) " _
                    & "values ('" & strTerms & "','" & mstrPractice & "', '" & strGroup & "'," & lngRow & ")"
                Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
            Else
                strSQL = "UPDATE MSE_HEADER_FORMAT SET VISIBLE = '" & strGroup & "', SEQUENCE_NO = " & lngRow & " WHERE TERM_ID = " & lngCareID
                Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
            End If
        Next lngRow
    End With
    MsgBox "Successfully Updated Mental Status", vbInformation, "PatientTrac Mental Status Configuration"
    If cnCareIntrvn.State = adStateOpen Then
        cnCareIntrvn.Close
        Set cnCareIntrvn = Nothing
        Set rsCareIntrvn = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Care_Plan:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Care_Plan", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cmdAll_Click()
Dim strSQL As String
Dim cnCareIntrvn As New ADODB.Connection
Dim rsCareIntrvn As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strTerms As String
Dim strGroup As String
Dim lngCareID As Long
Dim lngRow As Long
Dim lngCol As Long


    On Error GoTo Error_Get_All_MSE
    Screen.MousePointer = vbHourglass
    'show all
    If cnCareIntrvn.State = adStateClosed Then cnCareIntrvn.open psConnect(1)
    strSQL = "SELECT *  FROM MSE_HEADER_FORMAT ORDER BY HEADERS"
    Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
    With rsCareIntrvn
        Do Until .EOF
          strTerms = vbNullString
          strGroup = vbNullString
          lngCareID = 0
          Set Flds = .Fields
          For Each Fld In Flds
              If Not IsNull(Trim(Fld.Value)) Then
                  If Trim(Fld.Value) <> vbNullString Then
                      Select Case UCase(Fld.Name)
                          Case "HEADERS"
                              strTerms = IsProperCase(Trim(Fld.Value))
'                          Case "CARE_GROUP"
'                              strGroup = UCase(Trim(Fld.Value))
                          Case "TERM_ID"
                              lngCareID = Trim(Fld.Value)
                      End Select
                  End If
              End If
          Next
          .MoveNext
          If Trim(strTerms) <> vbNullString Then
                lngRow = grdTerms.MaxRows + 1
                grdTerms.MaxRows = lngRow
                grdTerms.SetRowItemData lngRow, lngCareID
                lngCol = 1
                Format_CellTypeEdit grdTerms, lngRow, lngCol, "1", strTerms, "TypeHAlignLeft", "INTERVENTION"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                lngCol = 2
                Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", Get_GridTerms("sd_CarePlan_Intrvn_Group"), "TypeHAlignCenter", "GROUP"
                Set_GridCurSel grdTerms, strGroup, lngCol, lngRow, "1", "GROUP"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
          End If
        Loop
    End With
    If cnCareIntrvn.State = adStateOpen Then
        cnCareIntrvn.Close
        Set cnCareIntrvn = Nothing
        Set rsCareIntrvn = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_All_MSE:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_All_MSE", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Sub

Private Sub cmdCloseList_Click()
    Unload Me
    Exit Sub
End Sub





Private Function Get_GridTerms(vstrSubDom As String) As String
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim lngRow As Long
Dim varCellValue
Dim Arr
Dim Idx As Integer
Dim lngCol As Long
Dim strTerms As String

    On Error GoTo Error_Get_GridTerms
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    strSQL = "Get_Attr_Valid_Value_SubDom_Sequence '" & vstrSubDom & "'"
    Set rsTerms = cnTerms.Execute(strSQL)
    With rsTerms
        Do Until .EOF
            strTerms = vbNullString
            If (!ATTRBTE_VALID_VALUE) <> vbNullString Then
                strTerms = Trim(!ATTRBTE_VALID_VALUE)
            End If
            .MoveNext
            If Trim(strTerms) <> vbNullString Then
                If Get_GridTerms = vbNullString Then
                    Get_GridTerms = UCase(Trim(strTerms))
                Else
                    Get_GridTerms = Get_GridTerms & Chr(9) & UCase(Trim(strTerms))
                End If
            End If
        Loop
    End With
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set rsTerms = Nothing
        Set cnTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_GridTerms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_GridTerms", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function

Private Sub cmdDown_Click()
Dim lngRow As Integer
    If grdTerms.ActiveRow = 0 Then
        MsgBox "Please select a row first!", vbInformation
        Exit Sub
    End If
    With grdTerms
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                Exit For
            End If
        Next lngRow
        .SwapRowRange lngRow, lngRow, lngRow + 1
    End With
    Exit Sub
End Sub




Private Sub cmdSave_Click()
    Write_Care_Plan
End Sub

Private Sub cmdUp_Click()
Dim lngRow As Integer

    If grdTerms.ActiveRow = 0 Then
        MsgBox "Please select a row first!", vbInformation
        Exit Sub
    End If
    If grdTerms.ActiveRow > 1 Then
        With grdTerms
            For lngRow = 1 To .MaxRows
                .Col = 1
                .Row = lngRow
                If .BackColor = grdCurItem Then
                    Exit For
                End If
            Next lngRow
            .SwapRowRange lngRow, lngRow, lngRow - 1
        End With
    End If
    Exit Sub
End Sub

Private Sub Form_Activate()
    If mFirst Then
        mFirst = False
    End If
    mbooByPass = True
    Load_Group_Filter
    Format_GridTerms
    Get_MSE_Terms
    cctProv.Text = clsProvider.Name
    mbooByPass = False
    centerForm Me
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

Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long


    If Row <= 0 Then Exit Sub
    With grdTerms
        For lngRow = 1 To .MaxRows
            .BlockMode = True
               .Col = Col
               .Col2 = .MaxCols
               .Row = lngRow
               .Row2 = lngRow
               .BackColor = NormalColor
               .SetCellBorder 1, lngRow, .MaxCols, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
            .BlockMode = False
        Next lngRow
        .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = Row
            .Row2 = Row
            .BackColor = grdCurItem
        .BlockMode = False
    End With
End Sub


Private Sub grdTerms_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    With grdTerms
        .BlockMode = True
           .Col = Col
           .Col2 = .MaxCols
           .Row = Row
           .Row2 = Row
           .BackColor = NormalColor
           .SetCellBorder 1, Row, .MaxCols, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .BlockMode = False
    End With
End Sub


