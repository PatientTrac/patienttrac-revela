VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmSystemErrorLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Error Log"
   ClientHeight    =   8340
   ClientLeft      =   4095
   ClientTop       =   2475
   ClientWidth     =   13635
   Icon            =   "frmSystemErrorLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   13635
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraQuery 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Query Builder "
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
      Height          =   3675
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   13545
      Begin VB.Frame fraShow 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Show Columns"
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
         Height          =   3165
         Left            =   8670
         TabIndex        =   6
         Top             =   330
         Width           =   2805
         Begin VB.CheckBox chkSelectAll_Show 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Select All"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   210
            TabIndex        =   8
            Top             =   2850
            Width           =   1275
         End
         Begin VB.ListBox lstShow 
            Height          =   2535
            ItemData        =   "frmSystemErrorLog.frx":0442
            Left            =   210
            List            =   "frmSystemErrorLog.frx":0467
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   270
            Width           =   2385
         End
      End
      Begin VB.Frame fraSortBy 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Sort By "
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
         Height          =   3165
         Left            =   5970
         TabIndex        =   2
         Top             =   330
         Width           =   2475
         Begin VB.CheckBox chkSelectAll_Sort 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Select All"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   210
            TabIndex        =   9
            Top             =   2820
            Width           =   1125
         End
         Begin VB.CommandButton cmdDown 
            Height          =   315
            Left            =   1890
            Picture         =   "frmSystemErrorLog.frx":04E4
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1590
            Width           =   375
         End
         Begin VB.CommandButton cmdUp 
            Height          =   315
            Left            =   1890
            Picture         =   "frmSystemErrorLog.frx":085B
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1230
            Width           =   375
         End
         Begin VB.ListBox lstSortBy 
            Height          =   2535
            ItemData        =   "frmSystemErrorLog.frx":0BD3
            Left            =   210
            List            =   "frmSystemErrorLog.frx":0BE9
            Style           =   1  'Checkbox
            TabIndex        =   3
            Top             =   270
            Width           =   1545
         End
      End
      Begin VB.Frame fraFilter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filter"
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
         Height          =   3015
         Left            =   210
         TabIndex        =   1
         Top             =   360
         Width           =   5535
         Begin VB.ComboBox cccApplication 
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
            Height          =   330
            Left            =   330
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2100
            Width           =   2025
         End
         Begin VB.ComboBox cccComputer 
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
            Height          =   330
            Left            =   330
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1290
            Width           =   2025
         End
         Begin VB.ComboBox cccUsers 
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
            Height          =   330
            Left            =   330
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   510
            Width           =   2025
         End
         Begin MSComCtl2.DTPicker ccdtoDate 
            Height          =   375
            Left            =   3120
            TabIndex        =   10
            Top             =   1290
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
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
            CustomFormat    =   "MM/dd/yyy hh:mm:ss"
            DateIsNull      =   -1  'True
            Format          =   59244545
            CurrentDate     =   38862.7362384259
            MinDate         =   37257
         End
         Begin MSComCtl2.DTPicker ccdFromDate 
            Height          =   375
            Left            =   3120
            TabIndex        =   11
            Top             =   510
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
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
            CustomFormat    =   "MM/dd/yyy hh:mm:ss"
            DateIsNull      =   -1  'True
            Format          =   59244545
            CurrentDate     =   38862.7362384259
            MinDate         =   37257
         End
         Begin VB.Label lblApp 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Application"
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
            Height          =   285
            Left            =   330
            TabIndex        =   19
            Top             =   1860
            Width           =   1305
         End
         Begin VB.Label lblComputer 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Computer"
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
            Height          =   285
            Left            =   330
            TabIndex        =   17
            Top             =   1020
            Width           =   1305
         End
         Begin VB.Label lblUser 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Users"
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
            Height          =   285
            Left            =   330
            TabIndex        =   15
            Top             =   270
            Width           =   1305
         End
         Begin VB.Label lbglToDate 
            BackColor       =   &H00E0E0E0&
            Caption         =   "To Date"
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
            Height          =   285
            Left            =   3120
            TabIndex        =   13
            Top             =   1020
            Width           =   1305
         End
         Begin VB.Label lblFromDate 
            BackColor       =   &H00E0E0E0&
            Caption         =   "From Date"
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
            Height          =   285
            Left            =   3120
            TabIndex        =   12
            Top             =   270
            Width           =   1305
         End
      End
      Begin PT_XP_Button.PT_XP cmdRun 
         Height          =   585
         Left            =   11700
         TabIndex        =   21
         Top             =   720
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1032
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
         Picture         =   "frmSystemErrorLog.frx":0C31
         AppearanceMode  =   1
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Run Query"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSystemErrorLog.frx":150B
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   525
         Left            =   11700
         TabIndex        =   22
         ToolTipText     =   "Close and do not Save"
         Top             =   1440
         Width           =   1515
         _ExtentX        =   2672
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
         Picture         =   "frmSystemErrorLog.frx":215D
         AppearanceThemes=   3
         BorderDrawEvent =   1
         TransparentColor=   14215660
         Caption         =   "Done"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSystemErrorLog.frx":2A37
         ShowFocus       =   -1  'True
      End
   End
   Begin FPSpreadADO.fpSpread grdError 
      Height          =   4515
      Left            =   0
      TabIndex        =   20
      Top             =   3750
      Width           =   13605
      _Version        =   458752
      _ExtentX        =   23998
      _ExtentY        =   7964
      _StockProps     =   64
      AutoCalc        =   0   'False
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
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
      MaxCols         =   15
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarMaxAlign=   0   'False
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmSystemErrorLog.frx":3689
      UnitType        =   0
      UserResize      =   0
      Appearance      =   2
      ClipboardOptions=   0
   End
End
Attribute VB_Name = "frmSystemErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fbooByPassList As Boolean
Private fbooByPassCheck As Boolean
Private fintColWidth(10) As Integer
Private fbooLoading As Boolean


Private Function AllSelected(vobjListBox As Object) As Boolean
Dim i As Integer

    AllSelected = True

    For i = 0 To vobjListBox.ListCount - 1
        If Not vobjListBox.Selected(i) Then
            AllSelected = False
            Exit For
        End If
    Next i
    Exit Function
End Function

Private Sub Format_Grid()
Dim strSQL As String
Dim cnError As New ADODB.Connection
Dim rsError As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field

    On Error GoTo Error_Format_Grid
     Screen.MousePointer = vbHourglass
     
    With grdError
        .ClearRange 1, 1, .MaxCols, .MaxRows, False
        .MaxRows = 0
    End With
    For r = 0 To lstShow.ListCount - 1
        lngCol = lngCol + 1
        grdError.ColWidth(lngCol) = 15
        Format_CellTypeEdit grdError, SpreadHeader, lngCol, "1", UCase(lstShow.List(r)), "TypeHAlignCenter", lstShow.List(r), HeaderColor
        grdError.SetCellBorder lngCol, SpreadHeader, lngCol, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    Next r
    lstSortBy.Selected(3) = True
    lstSortBy.ListIndex = -1
    Add_Spacer SpreadHeader + 1, grdError
    
    Exit Sub
Error_Format_Grid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_Grid" & "." & Me.Name
    Exit Sub
    Resume
End Sub

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
              .SetText vlngCol, vlngRow, Trim(vstrText)
              .TypeEditMultiLine = True   'Make multi-lines
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              'If vbooUnderline Then .FontUnderline = True
              If vlngBackColor > 0 Then .BackColor = vlngBackColor
              If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
              ElseIf Trim(vstrSpan) <> vbNullString Then
                lngSpan = vstrSpan
                If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
              End If
              .Text = Trim(vstrText)
              CellNote = .MaxTextCellHeight
              If .ROWHEIGHT(vlngRow) < CellNote Then
                .ROWHEIGHT(vlngRow) = CellNote
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
Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)

    With vObjGrid
        vlngRow = vlngRow
       If .MaxRows < vlngRow Then .MaxRows = vlngRow
        .BlockMode = True
            .Col = 0
            .Col2 = 0
            .Row = vlngRow
            .Row2 = vlngRow
            .ROWHEIGHT(vlngRow) = 3.5
            .BackColor = 12632256
            .AddCellSpan 0, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub
Private Sub chkSelectAll_Show_Click()
    fbooByPassList = True
    
    If Not fbooByPassCheck Then
        If chkSelectAll_Show.Value Then
            SelectAll lstShow, True
        Else
            SelectAll lstShow, False
        End If
    End If
    
    fbooByPassList = False
End Sub

Private Function HideCol(vstrName As String)
Dim lngRow As Long
Dim lngPrimKey As Long
Dim lngCol As Long
Dim varCellValue

    With grdEnctrSummary
        LookUpColName = vbNullString
        For lngCol = 1 To .MaxCols
            .Col = lngCol
            .Row = SpreadHeader
            .GetText lngCol, SpreadHeader, varCellValue
            If Trim(varCellValue) <> vbNullString And TextFound(vstrName, Trim(varCellValue), Len(vstrName), True) Then
                strFound = Trim(varCellValue)
                grdError.ColHidden = True
                Exit For
            End If
        Next lngCol
    End With
End Function
Private Sub SelectAll(vobjListBox As Object, vbooSelect As Boolean)
Dim i As Integer

    For i = 0 To vobjListBox.ListCount - 1
        vobjListBox.Selected(i) = vbooSelect
    Next i
    
    vobjListBox.ListIndex = -1
    
End Sub

Private Sub chkSelectAll_Sort_Click()
    fbooByPassList = True
    
    If Not fbooByPassCheck Then
        If chkSelectAll_Sort.Value Then
            SelectAll lstSortBy, True
        Else
            SelectAll lstSortBy, False
        End If
    End If
    fbooByPassList = False
    Exit Sub
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdDown_Click()
     If lstSortBy.ListIndex = lstSortBy.ListCount - 1 Then Exit Sub
    Reorder_ListBox lstSortBy, False
End Sub

Private Sub cmdRun_Click()
     Run_Query
End Sub

Private Sub cmdUp_Click()
    If lstSortBy.ListIndex = 0 Then Exit Sub
    Reorder_ListBox lstSortBy, True
End Sub


Private Sub Form_Activate()
Dim strDate As String
    fbooLoading = True
    Format_Grid
    Load_Users
    Load_Computers
    Load_Applications
    
    strDate = DateAdd("d", -7, Now)
    ccdFromDate.Value = Format(strDate, "mm/dd/yyyy hh:mm AM/PM")
    ccdtoDate.Value = Format(Now, "mm/dd/yyyy hh:mm AM/PM")
    SelectAll lstShow, True
    lstSortBy.ListIndex = -1
    Load_Grid
    fbooLoading = False
End Sub

Private Sub Form_Load()

    LogFile "Start Log System Errors ", True
    centerForm Me
    Me.Top = Me.Top + 180
    
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
  Exit Sub

End Sub

Private Sub lstShow_Click()
Dim i As Integer

    i = lstShow.ListIndex
    
    If i > -1 And Not fbooByPassList Then
    
        If lstShow.Selected(i) = False Then
            fbooByPassCheck = True
            chkSelectAll_Show.Value = 0
            fbooByPassCheck = False
        End If
        
    End If
    
    If AllSelected(lstShow) Then
        fbooByPassCheck = True
        chkSelectAll_Show.Value = 1
        fbooByPassCheck = False
    End If
    
    If Not AnySelected(lstShow) Then
        lstShow.ListIndex = -1
    End If
    
    Exit Sub
End Sub


Private Function AnySelected(vobjListBox As Object) As Boolean
Dim i As Integer

    For i = 0 To vobjListBox.ListCount - 1
        If vobjListBox.Selected(i) Then
            AnySelected = True
            Exit For
        End If
    Next i
End Function

Private Sub lstShow_ItemCheck(Item As Integer)
Dim strList As String
Dim Idx As Integer

    If lstShow.Selected(Item) = False Then
        For Idx = 0 To lstShow.ListCount - 1
            If lstShow.List(Idx) = lstShow.List(Item) Then
                strList = lstShow.List(Idx)
                HideCol strList
                Exit For
            End If
        Next Idx
    End If
    Exit Sub
End Sub

Private Sub lstSortBy_Click()
Dim i As Integer

    i = lstSortBy.ListIndex
    
    If i > -1 And Not fbooByPassList Then
    
        If lstSortBy.Selected(i) = False Then
        
            fbooByPassCheck = True
            chkSelectAll_Sort.Value = 0
            fbooByPassCheck = False
             
        End If
        
    End If
    
    If AllSelected(lstSortBy) Then
        fbooByPassCheck = True
        chkSelectAll_Sort.Value = 1
        fbooByPassCheck = False
    End If
    
    If Not AnySelected(lstSortBy) Then
        lstSortBy.ListIndex = -1
    End If
    
End Sub



Private Sub Load_Users()
Dim strSQL As String
Dim varRS
Dim booRet As Boolean
Dim lngRows As Long
Dim r As Integer
Dim lstrRemarks As String
Dim strUser As String
    
On Error GoTo Error_Load_Users

LogFile "Private Sub Load_Users()-----------------------------------"

    strSQL = "Select distinct user_name From ERRORS Order By user_name"
    
    LogFile strSQL
    lstrRemarks = strSQL
    
    booRet = SelectQuery(strSQL, psConnect(0), varRS, lngRows)
    
    If booRet Then
    
        If lngRows > 0 Then
            cccUsers.AddItem "(all)"
            strUser = Get_UserName()
            
            For r = LBound(varRS, 2) To UBound(varRS, 2)
                If Not IsNull(varRS(0, r)) Or varRS(0, r) <> "" Then
                    cccUsers.AddItem varRS(0, r)
                    If strUser = varRS(0, r) Then cccUsers.ListIndex = r + 1
                End If
            Next r
            
           
        End If
        
    End If
    
Exit Sub

Error_Load_Users:
    ErrHandler _
     Me.Name & " Sub Load_Users", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub Load_Computers()
Dim strSQL As String
Dim cnError As New ADODB.Connection
Dim rsError As New ADODB.Recordset
Dim lngCnt As Long

    
    On Error GoTo Error_Load_Computers
    Screen.MousePointer = vbHourglass
    cccComputer.Clear
    cccComputer.AddItem "(all)"
    If cnError.State = adStateClosed Then cnError.open psConnect(0)
    strSQL = "Select distinct computer From ERRORS "
    Set rsError = cnError.Execute(strSQL)
    Do Until rsError.EOF
        If rsError.Fields("computer") <> vbNullString Then
            strComputer = Get_ComputerName
            cccComputer.AddItem rsError.Fields("app_name")
            If strComputer = rsError.Fields("app_name") Then cccComputer.ListIndex = lngCnt + 1
        End If
        rsError.MoveNext
    Loop
    If cnError.State = adStateOpen Then cnError.Close
        cnError.Close
        Set cnError = Nothing
        Set rsError = Nothing
    End If
     Screen.MousePointer = vbDefault
     Exit Sub
Error_Load_Computers:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Computers" & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Load_Applications()
Dim strSQL As String
Dim cnError As New ADODB.Connection
Dim rsError As New ADODB.Recordset
Dim lngCnt As Long

 
    On Error GoTo Error_Load_Applications
     Screen.MousePointer = vbHourglass
    cccApplication.Clear
    cccApplication.AddItem "(all)"
    If cnError.State = adStateClosed Then cnError.open psConnect(0)
    strSQL = "Select distinct app_name From ERRORS "
    Set rsError = cnError.Execute(strSQL)
    Do Until rsError.EOF
        If rsError.Fields("app_name") <> vbNullString Then
            cccApplication.AddItem rsError.Fields("app_name")
            If App.EXEName = rsError.Fields("app_name") Then cccApplication.ListIndex = lngCnt + 1
        End If
        rsError.MoveNext
    Loop
    If cnError.State = adStateOpen Then cnError.Close
        cnError.Close
        Set cnError = Nothing
        Set rsError = Nothing
    End If
     Screen.MousePointer = vbDefault
Exit Sub
Error_Load_Applications:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Applications" & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Load_Grid()
Dim strSQL As String
Dim cnError As New ADODB.Connection
Dim rsError As New ADODB.Recordset
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim lngRows As Long
Dim r As Integer
Dim lstrRemarks As String
Dim strUser As String
Dim strComputer As String
Dim strApp As String
Dim strFromDate As String
Dim strToDate As String

    On Error GoTo Error_Load_Grid
     Screen.MousePointer = vbHourglass

    strUser = cccUsers.Text  'Get_UserName()
    strComputer = cccComputer.Text 'Get_ComputerName()
    strApp = cccApplication.Text
    strFromDate = Format(ccdFromDate.Value, "mm/dd/yyyy hh:nn")
    strToDate = Format(ccdtoDate.Text, "mm/dd/yyyy hh:nn")

    For r = 0 To lstShow.ListCount - 1
        strField(r) = lstShow.List(r)
    Next r
    If cnError.State = adStateClosed Then cnError.open psConnect(0)

    strSQL = "SELECT * From Errors " _
        & "Where (user_name = '" & strUser & "' or computer = '" & strComputer & "') And " _
                & "app_name = '" & strApp & "' And " _
                & "err_date between '" & strFromDate & "' And '" & strToDate & "'"
    Set rsError = cnError.Execute(strSQL)
    Do Until rsError.EOF
        Set Flds = rsError.Fields
        For Each Fld In Flds
            Format_CellTypeEdit grdError, SpreadHeader, lngCol, "1", UCase(Fld.Name), "TypeHAlignCenter", Fld.Name, HeaderColor
            grdError.SetCellBorder lngCol, SpreadHeader, lngCol, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Next
        rsError.MoveNext
    Loop
    lstSortBy.Selected(3) = True  'Date
    lstSortBy.ListIndex = -1
    Exit Sub
Error_Load_Grid:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Load_Grid" & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub Run_Query()
Dim strSQL As String
Dim varRS
Dim booRet As Boolean
Dim lngRows As Long
Dim r As Integer
Dim lstrRemarks As String
Dim strField(10) As String
Dim strUser As String
Dim strComputer As String
Dim strApp As String
Dim strFromDate As String
Dim strToDate As String
Dim strSort As String
Dim strTerm As String
Dim strOrderBy As String
 
    On Error GoTo Error_Run_Query

    If cccUsers.Text = "(all)" Then
        strUser = vbNullString
    Else
        strUser = "user_name = '" & cccUsers.Text & "' And "
    End If
    
    If cccComputer.Text = "(all)" Then
        strComputer = vbNullString
    Else
        strComputer = "computer = '" & cccComputer.Text & "' And "
    End If
    
    If cccApplication.Text = "(all)" Then
        strApp = vbNullString
    Else
        strApp = "app_name = '" & cccApplication.Text & "' And "
    End If
    
    strFromDate = Format(ccdFromDate.Value, "mm/dd/yyyy")
    strToDate = Format(ccdtoDate.Text, "mm/dd/yyyy")

    For r = 0 To lstSortBy.ListCount - 1
        If lstSortBy.Selected(r) Then
            strTerm = Translate_SortByTerm(lstSortBy.List(r))
            strSort = strSort & strTerm & ","
        End If
    Next r

    strSort = Left(strSort, Len(strSort) - 1)
    strOrderBy = "Order By " & strSort

    For r = 0 To lstShow.ListCount - 1
        strField(r) = lstShow.List(r)
    Next r
    
    strSQL = "Select " _
            & "err_num, " _
            & "err_desc, " _
            & "convert(char(20),err_date,100) as Date, " _
            & "app_name, " _
            & "app_version, " _
            & "app_proc, " _
            & "computer, " _
            & "user_name, " _
            & "remarks, " _
            & "patient_id, " _
            & "encounter_id " _
        & "From Errors " _
        & "Where " & strUser & strComputer & strApp _
                & "err_date between '" & strFromDate & "' And '" & strToDate & "' " _
                & strOrderBy

    booRet = SelectQuery(strSQL, psConnect(0), varRS, lngRows)
    
    If booRet Then
    
        If lngRows > 0 Then
            Load_Grid
        End If
        
    End If
    lstSortBy.ListIndex = -1
        
    Exit Sub
Error_Run_Query:
    ErrHandler _
     Me.Name & " Sub Run_Query", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Function Translate_SortByTerm(vstrSortBy As String) As String

    Select Case vstrSortBy
        
        Case "Application"
            Translate_SortByTerm = "app_name"
        
        Case "Computer"
            Translate_SortByTerm = "computer"
        
        Case "User"
            Translate_SortByTerm = "user_name"
            
        Case "Date"
            Translate_SortByTerm = "err_date"
        
        Case "Error Number"
            Translate_SortByTerm = "err_num"
        
        Case "Error Description"
            Translate_SortByTerm = "err_desc"
    
    End Select
    
End Function


