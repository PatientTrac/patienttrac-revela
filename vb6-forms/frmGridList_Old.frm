VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "pt_xp.ocx"
Begin VB.Form frmGridList 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   7050
   ClientLeft      =   10455
   ClientTop       =   3105
   ClientWidth     =   14070
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
   ScaleHeight     =   7050
   ScaleWidth      =   14070
   Begin VB.TextBox Text2 
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
      Height          =   525
      Left            =   2010
      TabIndex        =   16
      Top             =   1050
      Width           =   1395
   End
   Begin VB.TextBox Text1 
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
      Height          =   525
      Left            =   210
      TabIndex        =   14
      Top             =   1050
      Width           =   1395
   End
   Begin FPSpreadADO.fpSpread grdGrid 
      Height          =   2505
      Left            =   210
      TabIndex        =   13
      Top             =   3420
      Width           =   12615
      _Version        =   393216
      _ExtentX        =   22251
      _ExtentY        =   4419
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      ScrollBarMaxAlign=   0   'False
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmGridList.frx":08CA
      VScrollSpecial  =   -1  'True
      Appearance      =   2
   End
   Begin VB.TextBox cctSeqNo 
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
      Height          =   525
      Left            =   4830
      TabIndex        =   9
      Top             =   2340
      Width           =   1395
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
      Height          =   525
      Left            =   6870
      TabIndex        =   5
      Top             =   2340
      Width           =   2325
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
      Height          =   525
      Left            =   1500
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin PT_XP_Button.PT_XP cmdUp 
      Height          =   465
      Left            =   13350
      TabIndex        =   0
      Top             =   4410
      Width           =   375
      _ExtentX        =   661
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
      Picture         =   "frmGridList.frx":0AB7
      AppearanceThemes=   3
      TransparentColor=   12640511
      Caption         =   ""
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":1231
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDown 
      Height          =   435
      Left            =   13350
      TabIndex        =   1
      Top             =   5010
      Width           =   405
      _ExtentX        =   714
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
      Picture         =   "frmGridList.frx":1943
      AppearanceThemes=   3
      TransparentColor=   12640511
      Caption         =   ""
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":20BD
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDone 
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Tag             =   "MEDICAL_ HISTORY"
      Top             =   6270
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AppearanceThemes=   3
      TransparentColor=   12640511
      Caption         =   "Done"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdAddList 
      Height          =   345
      Left            =   240
      TabIndex        =   10
      Top             =   2970
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
      Picture         =   "frmGridList.frx":27CF
      AppearanceThemes=   3
      TransparentColor=   12640511
      Caption         =   "Add"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":2D69
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdReplace 
      Height          =   345
      Left            =   1170
      TabIndex        =   11
      Top             =   2970
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
      Picture         =   "frmGridList.frx":30BB
      AppearanceThemes=   3
      TransparentColor=   12640511
      Caption         =   "Replace"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":3655
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdRemove 
      Height          =   345
      Left            =   2400
      TabIndex        =   12
      Top             =   2970
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
      Picture         =   "frmGridList.frx":39A7
      AppearanceThemes=   3
      TransparentColor=   12640511
      Caption         =   "Remove"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmGridList.frx":3F41
      ShowFocus       =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Column Number"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2040
      TabIndex        =   17
      Top             =   750
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Row Number"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   750
      Width           =   1485
   End
   Begin VB.Label lblSeqNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Row Number"
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   4860
      TabIndex        =   7
      Top             =   2010
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Visible Property"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6870
      TabIndex        =   6
      Top             =   1860
      Width           =   1875
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Row Type"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   210
      Width           =   1875
   End
   Begin VB.Label lblSort 
      BackStyle       =   0  'Transparent
      Caption         =   "Row Order"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   12930
      TabIndex        =   3
      Top             =   3690
      Width           =   1065
   End
End
Attribute VB_Name = "frmGridList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrConnectStringClinical As String
Public mstrConnectStringTerms As String
Public mlngtop As Long
Public mlngleft As Long
Public mstrTerms As String
Public mvarSelectedTerms As String
Public mstrCurrentData As String
Private mbooByPass As Boolean
Public mstrListName As String
Public mFirst As Boolean
Public mbooImplant As Boolean
Public mstrSelection As String
Private mbooGarments As Boolean
Private mbooAbBinder As Boolean
Private mbooSpecialBraus As Boolean
Public mbooCaption As Boolean
Private mstrItemClick As String
Public mstrFormCaption As String
Private mRow As Long
Private mintGridRow As Integer
Private mintGridCol As Integer


Private Function Insert_GridData()
Dim i%
Dim strSQL As String
Dim strName As String
Dim lngSeqNo As String
Dim intVisibie As Integer
Dim varCellValue


    On Error GoTo ERROR_INSERT_GRIDTERMS

        With grdGrid
            .GetText 2, mintGridRow, varCellValue
            If varCellValue <> vbNullString Then strName = Trim(varCellValue)
            .GetText 3, mintGridRow, varCellValue
            If varCellValue <> vbNullString Then lngSeqNo = CLng(varCellValue)
            .GetText 4, mintGridRow, varCellValue
            If varCellValue <> vbNullString Then
                If varCellValue = "Visible" Then
                    intVisibie = 1
                Else
                    intVisibie = 0
                End If
            End If
        End With
        
        Open_DB mstrConnectStringTerms
        
        strSQL = "Insert into Grid_Headers (FORM_NAME, HEADERS, FORM_ID, SEQUENCE_NO, VISIBLE " _
            & "values('" & mstrCurrentData & "','" & strName & "'," & lngSeqNo & ",'" & intVisibie & "')"
            
        Set rs = cn.Execute(strSQL)
        
        Close_DB mstrConnectStringTerms
        
    Exit Function
    
ERROR_INSERT_GRIDTERMS:
    Screen.MousePointer = vbDefault
    MsgBox "Please contact PatientTrac Support with: INSERT GRID TERMS" & " " & Err.Description & " " & Err.Number, vbCritical
    Exit Function
    Resume
End Function
Private Function Update_GridData()
Dim i%
Dim strSQL As String
Dim strName As String
Dim lngSeqNo As String
Dim intVisibie As Integer
Dim varCellValue
Dim lngPrimKey As Long
Dim colID As Variant
Dim iCnt As Integer
Dim iCol As Integer
Dim iRow As Integer
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strFldsName As String
Dim byteChunk() As Byte
Dim NumOfChuncks As Long
Dim CurrentRecPos As Long
Dim FieldSize As Long
Dim varHeader
Dim strCount As String
Dim varSysName
Dim varField

    Open_DB mstrConnectStringTerms
    

    For iRow = 1 To grdGrid.MaxRows
             lngPrimKey = grdGrid.GetRowItemData(iRow)
             grdGrid.GetText 2, mintGridRow, varCellValue
            If varCellValue <> vbNullString Then strName = Trim(varCellValue)
            grdGrid.GetText 3, mintGridRow, varCellValue
            If varCellValue <> vbNullString Then lngSeqNo = CLng(varCellValue)
            grdGrid.GetText 4, mintGridRow, varCellValue
            If varCellValue <> vbNullString Then
                If varCellValue = "Visible" Then
                    intVisibie = 1
                Else
                    intVisibie = 0
                End If
            End If
        strSQL = "UPDATE GRID_HEADERS SET HEADERS = '" & strName & "', SEQUENCE_NO = " & lngSeqNo & ",VISIBLE = " & intVisibie & " " _
            & "WHERE FORM_ID = " & lngPrimKey
        Set rs = cn.Execute(strSQL)
    Next iRow
    
    Close_DB mstrConnectStringTerms
    
    Exit Function
    
End Function
Private Sub Format_Headers()
Dim colID As Variant
Dim iCnt As Integer
Dim iCol As Integer
Dim iRow As Integer
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strFldsName As String
Dim byteChunk() As Byte
Dim NumOfChuncks As Long
Dim CurrentRecPos As Long
Dim FieldSize As Long
Dim varHeader
Dim strCount As String
Dim strSQL As String
Dim varSysName
Dim varField

    
    Screen.MousePointer = vbHourglass
    mRow = 1

    With grdGrid
        .Redraw = False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .FontName = "Times New Roman"
        .GridSolid = False
        .GridShowHoriz = False
        .GridShowVert = False
        .RowHeadersShow = False
        '.RowHeadersAutoText = DispNumbers
        .ColHeadersShow = True
        .ColHeadersAutoText = DispBlank
        .ROWHEIGHT(0) = 15
        .MaxCols = 10
        '.ColHidden = True
        .ColWidth(1) = 10
        .ColWidth(2) = 20
        .ColWidth(3) = 15
        .ColWidth(4) = 10
        .ColWidth(5) = 10
        .ColWidth(6) = 10
        .ColWidth(7) = 10
        .ColWidth(8) = 10
        .ColWidth(9) = 10
        .ColWidth(10) = 10
 End With
    '------------------------------------------------------------
    ' Format Headers
    
    ' Format Headers
    With grdGrid
        .SetText 1, 0, "REPORT ID"
        .AddCellSpan 2, 0, 2, 1
        .Col = 2
        .TypeHAlign = TypeHAlignCenter
        .SetText 2, 0, "REPORT NAME"
        .AddCellSpan 2, 0, 1, 1
        .Col = 3
        .TypeHAlign = TypeHAlignCenter
        .SetText 3, 0, "HEADERS"
        .TypeHAlign = TypeHAlignCenter
        .Col = 4
        .TypeHAlign = TypeHAlignCenter
        .SetText 4, 0, "CELLTYPE"
        .TypeHAlign = TypeHAlignCenter
        .SetText 1, 0, "DATA_CELL"
        .AddCellSpan 2, 0, 2, 1
        .Col = 2
        .TypeHAlign = TypeHAlignCenter
        .SetText 2, 0, "TEXT_CELL"
        .AddCellSpan 2, 0, 1, 1
        .Col = 3
        .TypeHAlign = TypeHAlignCenter
        .SetText 3, 0, "COLUMN_NUMBER"
        .TypeHAlign = TypeHAlignCenter
        .Col = 4
        .TypeHAlign = TypeHAlignCenter
        .SetText 4, 0, "ROW_NUMBER"
        .TypeHAlign = TypeHAlignCenter
        .Col = 3
        .TypeHAlign = TypeHAlignCenter
        .SetText 3, 0, "VISIBLE"
        .TypeHAlign = TypeHAlignCenter
        .Col = 4
        .TypeHAlign = TypeHAlignCenter
        .SetText 4, 0, "ORDER NO"
        .TypeHAlign = TypeHAlignCenter
    End With
    
    
    
    Open_DB mstrConnectStringTerms

    
    strSQL = "Select HEADERS from REPORT_CONFIG where form_name = '" & mstrCurrentData & "' Order by SEQUENCE_NO"
    
    Set rs = cn.Execute(strSQL)
    
    strCount = 0
    Do Until rs.EOF
        strCount = strCount + 1
    rs.MoveNext
    Loop

    iRow = 0
   strSQL = "Select HEADERS, FORM_ID, SEQUENCE_NO, VISIBLE  from REPORT_CONFIG where form_name = '" & mstrCurrentData & "' Order by SEQUENCE_NO "
    Set rs = cn.Execute(strSQL)
    Set Flds = rs.Fields
    With rs
        Do Until .EOF
            grdGrid.MaxRows = strCount
            iRow = iRow + 1
            For Each Fld In Flds
                FieldSize = Fld.ActualSize
                If FieldSize > 0 Then
                    For iCol = 1 To grdGrid.MaxCols
                         grdGrid.GetText iCol, 0, varSysName
                        If varSysName = "REPORT_ID" Then varSysName = "REPORT_ID"
                        If varSysName = "HEADERS" Then varSysName = "HEADERS"
                        If varSysName = "CELLTYPE" Then varSysName = "CELLTYPE"
                        If varSysName = "DATA_CELL" Then varSysName = "DATA_CELL"
                        If varSysName = "COLUMN_NUMBER" Then varSysName = "COLUMN_NUMBER"
                        If varSysName = "ROW_NUMBER" Then varSysName = "ROW_NUMBER"
                        If varSysName = "ROW_NUMBER" Then varSysName = "ROW_NUMBER"
                        If varSysName = "GRID_ORDER" Then varSysName = "SEQUENCE_NO"
                        If varSysName = Fld.Name Then
                                grdGrid.Row = iRow
                                grdGrid.Col = iCol
                                varField = Fld.Value
                                If varField = "True" Then varField = "Visible"
                                grdGrid.SetText iCol, iRow, Trim(varField)
                                If iCol = 1 Then
                                    grdGrid.SetRowItemData (mintGridRow), Fld.Value
                                End If
                                grdGrid.TypeHAlign = TypeHAlignLeft
                        End If
                    Next iCol
                End If
            Next
        .MoveNext
        Loop
    End With
    grdGrid.Redraw = True
    Close_DB mstrConnectStringTerms
    Screen.MousePointer = vbDefault
    mbooHasChanged = False
    mbooByPass = False
    grdGrid.LeftCol = 1

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
   
   If (i > 0 And bUp) Or (i < lst.ListCount - 1 And Not bUp) Then
    lst.RemoveItem i
    lst.AddItem sContent, i - X
    lst.ItemData(i - X) = lData
    lst.Selected(i - X) = True
   End If
   
End Sub

Private Sub cmdAddList_Click()
Dim i%
Dim varCellValue
Dim iRow%


    With grdGrid
        If txtNewColumn.Text = vbNullString Then
            MsgBox "Enter Data in the textboxes first!", vbExclamation
            Exit Sub
        Else
            .MaxRows = .MaxRows + 1
            .InsertRows mintGridRow, 1
            .SetText 2, mintGridRow + 1, Trim(txtNewColumn.Text)
            .SetText 3, mintGridRow + 1, Trim(cctSeqNo.Text)
            .SetText 4, mintGridRow + 1, Trim(cctVisible.Text)
            Insert_GridData
            'Renumber Sequence Numbers
            For iRow = mintGridRow + 2 To .MaxRows
                .GetText 3, iRow, varCellValue
                If varCellValue <> vbNullString Then
                    varCellValue = Val(varCellValue) + 1
                End If
                .setttext 3, iRow, varCellValue
            Next iRow
            Update_GridData
        End If
    End With
    Exit Sub
End Sub



Private Sub cmdDone_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdDown_Click()
'    If lstListBox.ListIndex = lstListBox.ListCount - 1 Then Exit Sub
'    Reorder_ListBox lstListBox, True
End Sub

Private Sub cmdRemove_Click()
Dim iRow As Integer

    If mintGridRow = 0 Then
        MsgBox "Select row first!", vbExclamation
        Exit Sub
    End If
    With grdGrid
        .DeleteRows (mintGridRow), 1
        .MaxRows = .MaxRows - 1
    End With
    
    Exit Sub
End Sub

Private Sub cmdReplace_Click()
Dim iRow As Integer
Dim varCellValue

    If mintGridRow = 0 Then
        MsgBox "Select row first!", vbExclamation
        Exit Sub
    End If
    With grdGrid
        .Redraw = False
        .SetText 2, mintGridRow, Trim(txtNewColumn.Text)
        .SetText 3, mintGridRow, Trim(cctSeqNo.Text)
        ' Check for sequence number change
        .GetText 4, mintGridRow, varCellValue
        If varCellValue = Trim(cctVisible.Text) Then
             'Renumber Sequence Numbers
            For iRow = mintGridRow + 1 To .MaxRows
                .GetText 3, iRow, varCellValue
                If varCellValue <> vbNullString Then
                    varCellValue = Val(varCellValue) + 1
                End If
                .SetText 3, iRow, varCellValue
            Next iRow
        End If
        .SetText 4, mintGridRow, Trim(cctVisible.Text)
        MsgBox "Please set Row Position with up down buttons", vbInformation
        
        Update_GridData
        .Redraw = True
    End With
    
End Sub

Private Sub cmdUp_Click()
Dim iRow As Integer


    If mintGridRow = 0 Then
        MsgBox "Please select a row first!", vbInformation
    End If
    
'    With grdGrid
'        .MaxRows = .MaxRows + 1
'        .MoveRange (mintGridRow - 1), 1
'
'
'    End With

End Sub

Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
        centerForm Me
    End If
End Sub

Private Sub Form_Load()
Dim r%
Dim i%
Dim Arr
    centerForm Me
    mFirst = True
    Me.Caption = mstrFormCaption
    Format_Headers
    Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mstrTerms = vbNullString
    frmList.mstrListName = vbNullString
End Sub

Private Sub grdGrid_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue
Dim iRow%



    mintGridCol = Col
    mintGridRow = Row
    
    With grdGrid
    
        For iRow = 1 To .MaxRows
            .BlockMode = True
                .Col = 1
                .col2 = .MaxCols
                .Row = iRow
                .row2 = iRow
                .BackColor = &H80000005
            .BlockMode = False
        Next iRow
    
        .BlockMode = True
            .Col = 1
            .col2 = .MaxCols
            .Row = mintGridRow
            .row2 = mintGridRow
            .FontBold = True
             .BackColor = &HC0E0FF
        .BlockMode = False
        .GetText 2, mintGridRow, varCellValue
        If varCellValue <> vbNullString Then txtNewColumn.Text = Trim(varCellValue)
        .GetText 3, mintGridRow, varCellValue
        If varCellValue <> vbNullString Then cctSeqNo.Text = Trim(varCellValue)
        .GetText 4, mintGridRow, varCellValue
        If varCellValue <> vbNullString Then cctVisible.Text = Trim(varCellValue)
    End With
    Exit Sub
End Sub

