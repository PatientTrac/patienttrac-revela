VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmManage 
   ClientHeight    =   4935
   ClientLeft      =   12885
   ClientTop       =   4755
   ClientWidth     =   7740
   FillColor       =   &H00C00000&
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7740
   Begin PT_XP_Button.PT_XP cmdTerms 
      Height          =   375
      Left            =   2786
      TabIndex        =   0
      ToolTipText     =   "Add Terms"
      Top             =   4440
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmManage.frx":058A
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   " Select "
      CaptionAlignment=   3
      CaptionPosition =   4
      DepthMode       =   1
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   4194304
      ForeColorFocus  =   4194304
      ForeColorDown   =   4194304
      PictureDisabled =   "frmManage.frx":0B24
      PictureAlignment=   5
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdFacility 
      Height          =   4305
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   7665
      _Version        =   524288
      _ExtentX        =   13520
      _ExtentY        =   7594
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
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   16777215
      GridColor       =   8421504
      MaxCols         =   2
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   0
      ShadowText      =   0
      SpreadDesigner  =   "frmManage.frx":0E76
      UserResize      =   0
      Appearance      =   2
      TextTipDelay    =   100
      ScrollBarTrack  =   1
      ClipboardOptions=   0
      CellNoteIndicator=   2
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   375
      Left            =   3919
      TabIndex        =   2
      Top             =   4440
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmManage.frx":1344
      AppearanceThemes=   3
      TransparentColor=   15790320
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmManage.frx":1D56
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Load_Client_Facility()
Dim strSQL As String
Dim rsPatient As New ADODB.Recordset
Dim cnPatient As New ADODB.Connection
Dim Flds  As ADODB.Fields
Dim Fld As ADODB.Field
Dim strNPI As String
Dim strClient As String
Dim lngID As Long
Dim lngRow As Long
Dim varCellValue
Dim lngCol As Long

    On Error GoTo Error_Load_Client_Facility
    Screen.MousePointer = vbHourglass
    Format_Facilities
    If cnPatient.State = adStateClosed Then cnPatient.Open psConnect(0)
    strSQL = "SELECT * FROM CLIENT_DATABASE ORDER BY CLIENT_NAME"
    Set rsPatient = cnPatient.Execute(strSQL)
    With rsPatient
        Do Until .EOF
            Set Flds = .Fields
            For Each Fld In Flds
                If Not IsNull(Fld.Value) Then
                    If Trim(Fld.Value) <> vbNullString Then
                            Select Case (Fld.Name)
                                Case "CLIENT_NAME"
                                    strClient = Trim(Fld.Value)
                                Case "TAX_ID"
                                    strNPI = Fld.Value
                            End Select
                    End If
                End If
            Next
            .MoveNext
            If Trim(strClient) <> vbNullString Then
                With grdFacility
                        lngRow = grdFacility.MaxRows + 1
                        grdFacility.MaxRows = lngRow
                        lngCol = 1
                        Format_CellTypeEdit grdFacility, lngRow, 1, "1", strClient, "TypeHAlignLeft"
                        grdFacility.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                        lngCol = 2
                        Format_CellTypeEdit grdFacility, lngRow, lngCol, "1", CStr(strNPI), "TypeHAlignCenter"
                        grdFacility.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    'End If
                End With
            End If
        Loop
    End With
    If cnPatient.State = adStateOpen Then
        cnPatient.Close
        Set cnPatient = Nothing
        Set rsPatient = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Client_Facility:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Client_Facility", App.EXEName & "." & "ModBilling"
    Exit Function
    Resume Next
End Function
Private Function Format_Facilities()
Dim lngColWidth As Long
Dim varCellValue
Dim varHeader
Dim lngCnt  As Long
Dim lngCol As Long
Dim lngRow As Long
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim strHeaders As String

    On Error GoTo Error_Format_Problem
    Screen.MousePointer = vbHourglass
    With grdFacility
        .RemoveCellSpan 1, SpreadHeader
        .RemoveCellSpan 1, SpreadHeader + 2
        .ClearRange 1, SpreadHeader, 2, .MaxRows, False
        .ClearRange 1, SpreadHeader + 2, 2, .MaxRows, False
        .Redraw = False
        .ProcessTab = True
        .FontBold = True
        .FontName = "Georgia"
        .RowHeadersShow = False
        .ColHeadersShow = True
        .MaxCols = 2
        .MaxRows = 0
        .ColHeaderRows = 4
        .ColHeadersAutoText = DispBlank
    End With
    grdFacility.ColWidth(1) = 42.8
    grdFacility.ColWidth(2) = 18
    lngCol = 1
    lngRow = SpreadHeader
    Format_CellTypeEdit grdFacility, lngRow, lngCol, "2", "FACILITIES", "TypeHAlignCenter", " ", HeaderColor
    grdFacility.SetCellBorder lngCol, lngRow, 2, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 1, grdFacility
    lngCol = 1
    lngRow = SpreadHeader + 2
    Format_CellTypeEdit grdFacility, lngRow, lngCol, "1", "FACILITY  NAME", "TypeHAlignCenter", "FACILITY_NAME", HeaderColor
    grdFacility.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    lngCol = 2
    Format_CellTypeEdit grdFacility, lngRow, lngCol, "1", "TAX ID", "TypeHAlignCenter", "FACILITY_ID", HeaderColor
    grdFacility.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Add_Spacer SpreadHeader + 3, grdFacility
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_Problem:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_Problem", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function

Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional FontSize As Long)
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
                .TypeEditMultiLine = True
                .FontName = "Georgia"
                If FontSize > 0 Then
                    .FontSize = FontSize
                Else
                    .FontSize = 10
                End If
                .FontBold = False
'                .ForeColor = GridBlue
                If vlngBackColor > 0 Then .BackColor = vlngBackColor
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                  .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                ElseIf Trim(vstrSpan) <> vbNullString Then
                  lngSpan = vstrSpan
                  If lngSpan > 1 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
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
            .ClearRange 1, vlngRow, .MaxCols, vlngRow, False
            .RowHeight(vlngRow) = 3.5
            .AddCellSpan 1, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub

Private Sub cmdClose_Click()

    Unload Me
    Exit Sub
End Sub

Private Sub cmdTerms_Click()
Dim lngFac As Long
Dim i%
Dim lngRow As Long
Dim strTax As String
Dim strName As String
Dim varCellValue

    With grdFacility
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                .GetText 2, lngRow, varCellValue
                strTax = varCellValue
                Exit For
            End If
        Next lngRow
    End With
    If Trim(strTax) <> vbNullString Then
        Get_Database_Change Trim(strTax)
        Get_Facilities_By_Tax Trim(strTax)
        piFacilityID = clsFacility.FacID
        Unload Me
    End If
    Exit Sub
End Sub


Private Sub Form_Activate()

    Exit Sub
End Sub

Private Sub Form_Load()
    Load_Client_Facility
    Me.Top = frmMDI.Top + 900
    Exit Sub
End Sub

Private Sub grdFacility_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim varCellValue

    If Row > 0 Then
        With grdFacility
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
                .BackColor = grdCurItem
            .BlockMode = False
        End With
        Screen.MousePointer = vbDefault
    End If
End Sub


