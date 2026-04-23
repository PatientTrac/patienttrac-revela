VERSION 5.00
Object = "{13ADDE2B-B303-4618-975F-648031250148}#5.0#0"; "PolarSpellChecker.dll"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form fmMseTerms 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Mental Status Selection"
   ClientHeight    =   5250
   ClientLeft      =   7335
   ClientTop       =   3885
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fmMseTerms.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   8970
   Begin VB.Frame fraMseTerms 
      Height          =   3645
      Left            =   1650
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   4965
      Begin SPELLCHECKERLibCtl.SpellChecker SpellChecker1 
         Left            =   240
         OleObjectBlob   =   "fmMseTerms.frx":058A
         Top             =   3090
      End
      Begin VB.TextBox txtValue 
         Height          =   405
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   150
         Width           =   4245
      End
      Begin FPSpreadADO.fpSpread grdMSETerms 
         Height          =   2175
         Left            =   0
         TabIndex        =   11
         Top             =   990
         Width           =   4890
         _Version        =   524288
         _ExtentX        =   8625
         _ExtentY        =   3836
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         ColHeaderDisplay=   0
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
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         GridColor       =   8388608
         MaxCols         =   1
         MaxRows         =   0
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   16761024
         ShadowDark      =   8388608
         ShadowText      =   8388608
         SpreadDesigner  =   "fmMseTerms.frx":05AE
         UnitType        =   0
         UserResize      =   0
         Appearance      =   2
         TextTipDelay    =   30
         ClipboardOptions=   0
         AppearanceStyle =   0
      End
      Begin PT_XP_Button.PT_XP cmdSelectTerm 
         Height          =   315
         Left            =   1620
         TabIndex        =   12
         Top             =   3240
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "fmMseTerms.frx":0AC2
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Select"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "fmMseTerms.frx":0E5C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCloseTerm 
         Height          =   315
         Left            =   2670
         TabIndex        =   13
         Top             =   3240
         Width           =   975
         _ExtentX        =   1720
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
         Picture         =   "fmMseTerms.frx":11AE
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DefaultThemes   =   2
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "fmMseTerms.frx":1BC0
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdAdd 
         Height          =   345
         Left            =   30
         TabIndex        =   15
         Top             =   600
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
         Picture         =   "fmMseTerms.frx":1F12
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Add"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "fmMseTerms.frx":24AC
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdReplace 
         Height          =   345
         Left            =   1260
         TabIndex        =   16
         Top             =   600
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
         Picture         =   "fmMseTerms.frx":27FE
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Replace"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "fmMseTerms.frx":2D98
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRemove 
         Height          =   345
         Left            =   2460
         TabIndex        =   17
         Top             =   600
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
         Picture         =   "fmMseTerms.frx":30EA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Remove"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "fmMseTerms.frx":3684
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCheckComment 
         Height          =   315
         Left            =   4470
         TabIndex        =   18
         ToolTipText     =   "Select for Quotation!"
         Top             =   180
         Width           =   435
         _ExtentX        =   767
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
         MaskColor       =   128
         Picture         =   "fmMseTerms.frx":39D6
         AppearanceThemes=   3
         BorderWidth     =   2
         TransparentColor=   14737632
         Caption         =   ""
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   -2147483629
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "fmMseTerms.frx":3F70
         PictureAlignment=   4
         ShowFocus       =   -1  'True
      End
   End
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
      ItemData        =   "fmMseTerms.frx":42C2
      Left            =   5550
      List            =   "fmMseTerms.frx":42CF
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   30
      Width           =   2925
   End
   Begin PT_XP_Button.PT_XP cmdCloseList 
      Height          =   435
      Left            =   4485
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
      Picture         =   "fmMseTerms.frx":4305
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmMseTerms.frx":4D17
      ShowFocus       =   -1  'True
   End
   Begin FPSpreadADO.fpSpread grdTerms 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   8355
      _Version        =   524288
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
      SpreadDesigner  =   "fmMseTerms.frx":5069
      Appearance      =   2
      ClipboardOptions=   0
      AppearanceStyle =   0
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   435
      Left            =   3375
      TabIndex        =   2
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
      Picture         =   "fmMseTerms.frx":5534
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmMseTerms.frx":58CE
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdUp 
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   1950
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
      Picture         =   "fmMseTerms.frx":5C20
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmMseTerms.frx":639A
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDown 
      Height          =   465
      Left            =   8520
      TabIndex        =   4
      Top             =   2580
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
      Picture         =   "fmMseTerms.frx":6AAC
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   ""
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmMseTerms.frx":7226
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdAll 
      Height          =   315
      Left            =   7230
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
      Picture         =   "fmMseTerms.frx":7938
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Show All"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmMseTerms.frx":7ED2
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPrintRecords 
      Height          =   375
      Left            =   7290
      TabIndex        =   19
      ToolTipText     =   "Print Images"
      Top             =   4620
      Width           =   1155
      _ExtentX        =   2037
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
      Picture         =   "fmMseTerms.frx":8224
      AppearanceThemes=   3
      BorderWidth     =   0
      TransparentColor=   14737632
      Caption         =   "Print"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "fmMseTerms.frx":87BE
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
Private mstrGridTerm As String

Private Function Format_Grid_Field(vstrHeader As String)
Dim strHeader As String
Dim varCellValue

    Screen.MousePointer = vbHourglass
    strHeader = Trim(vstrHeader)
    
    txtValue = vbNullString
    With grdMSETerms
        .Redraw = False
       .ClearRange 0, SpreadHeader, grdMSETerms.MaxRows, grdMSETerms.MaxCols, False
        .ProcessTab = True
        .FontBold = True
        .ColHeaderRows = 2
        .ColHeadersShow = True
        .RowHeadersShow = False
        .ColWidth(1) = 39
        Format_CellTypeEdit grdMSETerms, SpreadHeader, 1, "1", UCase(strHeader), "TypeHAlignCenter", "HEADER", "", HeaderColor
        grdMSETerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdMSETerms
        grdMSETerms.MaxRows = 0
        .Redraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Function
End Function

Private Function Get_MSE_Terms(vstrType As String)
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
Dim strType As String


    On Error GoTo Error_Get_MSE_Terms
    Screen.MousePointer = vbHourglass
    strType = vstrType
    If cnCareIntrvn.State = adStateClosed Then cnCareIntrvn.open psConnect(1)
    Select Case IsProperCase(Trim(vstrType))
        Case "All Visits"
            strSQL = "SELECT *  FROM MSE_HEADER_FORMAT WHERE " _
                & "VISIBLE = 'True' ORDER BY SEQUENCE_NO"
        Case "Initial Patient"
            strSQL = "SELECT *  FROM MSE_HEADER_FORMAT WHERE " _
                & "VISIBLE = 'True' AND INITIAL_PATIENT = 'True' ORDER BY SEQUENCE_NO"
        Case "Subsequent Patient"
            strSQL = "SELECT *  FROM MSE_HEADER_FORMAT WHERE " _
                & "VISIBLE = 'True' AND SUBSEQUENT_PATIENT = 'True' ORDER BY SEQUENCE_NO"
    End Select
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
                          Case "VISIBLE"
                              strGroup = UCase(Trim(Fld.Value))
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
                Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", "True" & Chr(9) & "False", "TypeHAlignCenter", "GROUP"
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
        .Redraw = False
       .ClearRange 0, SpreadHeader, grdTerms.MaxRows, grdTerms.MaxCols, False
        .ProcessTab = True
        .FontBold = True
        .ColHeaderRows = 2
        .ColHeadersShow = True
        .RowHeadersShow = False
        .ColWidth(1) = 38
        .ColWidth(2) = 28
        Format_CellTypeEdit grdTerms, SpreadHeader, 1, "1", "MSE TERM", "TypeHAlignCenter", "INTERVENTION", "", HeaderColor
        Format_CellTypeEdit grdTerms, SpreadHeader, 2, "maxcols", "VISIBLE", "TypeHAlignCenter", "GROUP", "", HeaderColor
        grdTerms.SetCellBorder 1, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        grdTerms.SetCellBorder 2, SpreadHeader, 2, SpreadHeader, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        Add_Spacer SpreadHeader + 1, grdTerms
        grdTerms.MaxRows = 0
        .Redraw = True
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
        .Redraw = False
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
        .Redraw = True
    End With
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional vlngBackColor As Long, Optional vlngFontSize As Long)
Dim CellNote As Long
Dim lngSpan As Long

        
    With vGrid
        .Redraw = False
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
        .Redraw = True
    End With
End Function
Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrCellNote As String, Optional lngBackColor As Long, Optional vstrVAlign As String)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .Redraw = False
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
        .Redraw = True
    End With
End Function
Private Function Remove_Duplicate(vlngPrim As Long)
Dim strSQL As String
Dim cnCareIntrvn As New ADODB.Connection
Dim rsCareIntrvn As New ADODB.Recordset
Dim lngRow As Long
Dim strHeader As String
Dim varCellValue
Dim lngPrim As Long

    On Error GoTo Error_Remove_Duplicate
    lngPrim = vlngPrim
    If cnCareIntrvn.State = adStateClosed Then cnCareIntrvn.open psConnect(1)
    strSQL = "DELETE ATTR_VALID_VALUE WHERE ATTR_ID = " & lngPrim
    Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
    If cnCareIntrvn.State = adStateOpen Then
        cnCareIntrvn.Close
        Set cnCareIntrvn = Nothing
        Set rsCareIntrvn = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Remove_Duplicate:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Remove_Duplicate", App.EXEName & "." & Me.Name
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
            .Redraw = False
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
        .Redraw = True
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
            strSQL = "SELECT TERM_ID FROM MSE_HEADER_FORMAT WHERE HEADER = '" & strTerms & "'"
            Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
            If Not rsCareIntrvn.EOF Then
                lngCareID = rsCareIntrvn.Fields("TERM_ID")
            End If
            Select Case cboGroup
                Case "All Visits"
                    If lngCareID = 0 Then
                        strSQL = "INSERT INTO MSE_HEADER_FORMAT (HEADERS, VISIBLE, INITIAL_PATIENT, SUBSEQUENT_PATIENT, SEQUENCE_NO) " _
                            & "values ('" & strTerms & "','" & mstrPractice & "', '" & strGroup & "','" & strGroup & "','" & strGroup & "'," & lngRow & ")"
                        Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
                    Else
                        strSQL = "UPDATE MSE_HEADER_FORMAT SET VISIBLE = '" & strGroup & "', " _
                            & "INITIAL_PATIENT = '" & strGroup & "', " _
                            & "SUBSEQUENT_PATIENT = '" & strGroup & "', " _
                            & "SEQUENCE_NO = " & lngRow & " WHERE TERM_ID = " & lngCareID
                        Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
                    End If
                Case "Initial Patient", "New Patient"
                    If lngCareID = 0 Then
                        strSQL = "INSERT INTO MSE_HEADER_FORMAT (HEADERS, VISIBLE, INITIAL_PATIENT, SEQUENCE_NO) " _
                            & "values ('" & strTerms & "','" & mstrPractice & "', '" & strGroup & "','" & strGroup & "'," & lngRow & ")"
                        Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
                    Else
                        strSQL = "UPDATE MSE_HEADER_FORMAT SET VISIBLE = '" & strGroup & "', " _
                            & "INITIAL_PATIENT = '" & strGroup & "', " _
                            & "SEQUENCE_NO = " & lngRow & " WHERE TERM_ID = " & lngCareID
                        Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
                    End If
                Case "Subsequent Patient", "Established Patient"
                    If lngCareID = 0 Then
                        strSQL = "INSERT INTO MSE_HEADER_FORMAT (HEADERS, VISIBLE, SUBSEQUENT_PATIENT, SEQUENCE_NO) " _
                            & "values ('" & strTerms & "','" & mstrPractice & "', '" & strGroup & "','" & strGroup & "'," & lngRow & ")"
                        Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
                    Else
                        strSQL = "UPDATE MSE_HEADER_FORMAT SET VISIBLE = '" & strGroup & "', " _
                            & "SUBSEQUENT_PATIENT = '" & strGroup & "', " _
                            & "SEQUENCE_NO = " & lngRow & " WHERE TERM_ID = " & lngCareID
                        Set rsCareIntrvn = cnCareIntrvn.Execute(strSQL)
                    End If
            End Select
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

Private Function Write_Terms()
Dim lngRow As Long
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim varCellValue
Dim lngPrim As Long

    On Error GoTo Error_Write_Terms
    With grdMSETerms
        For lngRow = 1 To .MaxRows
            lngPrim = 0
            .GetText 1, lngRow, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                lngPrim = .GetRowItemData(lngRow)
                If lngPrim > 0 Then
                    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
                    strSQL = "UPDATE ATTR_VALID_VALUE SET VALID_VALUE_SEQ_NBR = " & lngRow & " WHERE ATTR_ID = " & lngPrim
                    Set rsTerms = cnTerms.Execute(strSQL)
                End If
            End If
        Next lngRow
    End With
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Terms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Terms", App.EXEName & "." & Me.Name
    Exit Function
    Resume Next
End Function

Private Sub cboGroup_Click()
    Get_MSE_Terms Trim(cboGroup.Text)
    Exit Sub
End Sub


Private Sub cmdAdd_Click()
Dim lngRow As Long
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim varCellValue

    On Error GoTo Error_Add_Terms
'    If mlngProvID > 0 Then
'        Get_SecurityRole mlngProvID
'        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
'            If UCase(clsPracticeRoles.RoleTitle) <> "Administrator" Then
'                MsgBox "You are not an Administrator and cannot modify Terms,  please sent suggestions to support@patienttrac.com!", vbInformation, "PatientTrac Security Management"
'                Screen.MousePointer = vbDefault
'                Exit Sub
'            End If
'        End If
'    End If
    If Trim(txtValue.Text) = vbNullString Then
        MsgBox "Enter a term first.", vbInformation, "PatientTrac Controlled Vocabulary Management"
        txtValue.SetFocus
        Exit Sub
    End If
    lngRow = grdMSETerms.MaxRows + 1
    grdMSETerms.MaxRows = lngRow
    Format_CellTypeEdit grdMSETerms, lngRow, 1, "1", LCase(txtValue.Text), "TypeHAlignLeft"
    grdMSETerms.SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    txtValue.Text = Replace(txtValue.Text, "'", "''")
    strSQL = "INSERT INTO ATTR_VALID_VALUE (VALID_VALUE_SUBDOM, ATTRBTE_VALID_VALUE, VALID_VALUE_SEQ_NBR) " _
        & "VALUES ('" & mstrGridTerm & "','" & LCase(txtValue.Text) & "', " & lngRow & ")"
    Set rsTerms = cnTerms.Execute(strSQL)
    'MsgBox "Successfully Added" & " " & LCase(txtValue.Text), vbInformation, "PatientTrac Mental Status Configuration"
    txtValue.Text = vbNullString
    txtValue.SetFocus
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Add_Terms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Add_Terms", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

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
Dim Idx As Long


    On Error GoTo Error_Get_All_MSE
    Screen.MousePointer = vbHourglass
    'show all
    If cnCareIntrvn.State = adStateClosed Then cnCareIntrvn.open psConnect(1)
    strSQL = "SELECT *  FROM MSE_HEADER_FORMAT ORDER BY HEADER"
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
                          Case "HEADER"
                              strTerms = IsProperCase(Trim(Fld.Value))
                          Case "VISIBLE"
                              strGroup = UCase(Trim(Fld.Value))
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
                Format_CellTypeCombo grdTerms, lngRow, lngCol, "1", "True" & Chr(9) & "False", "TypeHAlignCenter", "GROUP"
                Set_GridCurSel grdTerms, strGroup, lngCol, lngRow, "1", "GROUP"
                grdTerms.SetCellBorder lngCol, lngRow, lngCol, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
          End If
        Loop
    End With
    For Idx = 0 To cboGroup.listcount - 1
        If cboGroup.List(Idx) = "All Visits" Then
            cboGroup.Text = cboGroup.List(Idx)
            Exit For
        End If
    Next Idx
    If cnCareIntrvn.State = adStateOpen Then
        cnCareIntrvn.Close
        Set cnCareIntrvn = Nothing
        Set rsCareIntrvn = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Get_All_MSE:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_All_MSE", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdCheckComment_Click()
Dim strResultText As String
Dim vExitCode As Variant
Dim fso As New FileSystemObject

    On Error GoTo Error_Spell
    If fso.FileExists(mstrProgram & "\FirstNotebook\PatientTrac EMR\Dictionary\SpellUS.lex") Then
        SpellChecker1.OpenDictionary mstrProgram & "\FirstNotebook\PatientTrac EMR\Dictionary\SpellUS.lex"
        txtValue.SetFocus
        strResultText = SpellChecker1.CheckText(txtValue.Text, vExitCode)
        If Trim(txtValue.Text) = Trim(strResultText) Then
            MsgBox "No spelling errors found!", vbInformation, "PatientTrac Spell Checker"
        End If
        txtValue.Text = Trim(strResultText)
        SpellChecker1.CloseDictionary mstrProgram & "\FirstNotebook\PatientTrac EMR\Dictionary\SpellUS.lex"
    End If
    Exit Sub
Error_Spell:
    Exit Sub
    Resume
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
Dim lngPrim As Long
Dim bExit As Boolean
Dim lngTest As Long

    On Error GoTo Error_Get_GridTerms
    If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
    mstrGridTerm = vstrSubDom
    strSQL = "SELECT ATTRBTE_VALID_VALUE, ATTR_ID, VALID_VALUE_SEQ_NBR " _
        & "FROM ATTR_VALID_VALUE WHERE VALID_VALUE_SUBDOM = '" & mstrGridTerm & "' ORDER " _
        & "BY VALID_VALUE_SEQ_NBR"
    Set rsTerms = cnTerms.Execute(strSQL)
    With rsTerms
        Do Until .EOF
            strTerms = vbNullString
            lngPrim = 0
            bExit = False
            If (!ATTRBTE_VALID_VALUE) <> vbNullString Then
                strTerms = Trim(!ATTRBTE_VALID_VALUE)
            End If
            If (!ATTR_ID) <> vbNullString Then
                lngPrim = Trim(!ATTR_ID)
            End If
            .MoveNext
            If Trim(strTerms) <> vbNullString Then
                With grdMSETerms
                    For lngTest = 1 To .MaxRows
                        .GetText 1, lngTest, varCellValue
                        If Trim(LCase(varCellValue)) = Trim(LCase(strTerms)) Then
                            bExit = True
                            Remove_Duplicate lngPrim
                        End If
                    Next lngTest
                    If Not bExit Then
                        lngRow = lngRow + 1
                        .MaxRows = lngRow
                        grdMSETerms.SetRowItemData lngRow, lngPrim
                        Format_CellTypeEdit grdMSETerms, lngRow, 1, "1", LCase(strTerms), "TypeHAlignLeft"
                        .SetCellBorder 1, lngRow, 1, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
                    End If
                End With
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

Private Sub cmdCloseTerm_Click()
    fraMseTerms.Visible = False
    Exit Sub
End Sub

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

Private Sub cmdPrintRecords_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(1) As String
Dim strHeaders As String
Dim i As Integer

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.Hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation, "PatientTrac Multi Appointment Schedule"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    With grdTerms
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintJobName = "MSE TERMS"
        .PrintColor = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
        .hDCPrinter = Printer.hdc
        .PrintSmartPrint = True
        .PrintAbortMsg = "Printing - Click Cancel to quit"
        .PrintBorder = True
        .PrintColHeaders = True
        .PrintRowHeaders = True
        .PrintScalingMethod = PrintScalingMethodSmartPrint
        .PrintMarginLeft = 650
        .PrintMarginRight = 700
        .PrintMarginBottom = 900
        .PrintMarginTop = 900
        .PrintUseDataMax = True
        .PrintColor = True
        .PrintShadows = False
        .PrintSheet
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdPrint_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Print", App.EXEName & "." & Me.Name
    Exit Sub
    Resume

End Sub

Private Sub cmdRemove_Click()
Dim strSQL As String
Dim cnCareIntrvn As New ADODB.Connection
Dim rsCareIntrvn As New ADODB.Recordset
Dim lngRow As Long
Dim strHeader As String
Dim varCellValue
Dim lngPrim As Long

    On Error GoTo Error_Remove_Terms
'    If mlngProvID > 0 Then
'        Get_SecurityRole mlngProvID
'        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
'            If UCase(clsPracticeRoles.RoleTitle) <> "Administrator" Then
'                MsgBox "You are not an Administrator and cannot modify Terms,  please sent suggestions to support@patienttrac.com!", vbInformation, "PatientTrac Security Management"
'                Screen.MousePointer = vbDefault
'                Exit Sub
'            End If
'        End If
'    End If
    With grdMSETerms
        For lngRow = 1 To .MaxRows
            .Col = 1
            .Row = lngRow
            If .BackColor = grdCurItem Then
                lngPrim = .GetRowItemData(lngRow)
                Remove_Duplicate lngPrim
                .GetText 1, lngRow, varCellValue
                strHeader = Trim(varCellValue)
                .DeleteRows lngRow, 1
                .MaxRows = .MaxRows - 1
                Exit For
            End If
        Next lngRow
    End With
    
    txtValue = vbNullString
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Remove_Terms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Remove_Terms", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdReplace_Click()
Dim strSQL As String
Dim rsTerms As New ADODB.Recordset
Dim cnTerms As New ADODB.Connection
Dim lngRow As Long
Dim varCellValue
Dim iAnswer As Integer
Dim sItemData As String
Dim lngPrim As Long


    On Error GoTo Error_Replace_Terms
'    If mlngProvID > 0 Then
'        Get_SecurityRole mlngProvID
'        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
'            If UCase(clsPracticeRoles.RoleTitle) <> "Administrator" Then
'                MsgBox "You are not an Administrator and cannot modify Terms,  please sent suggestions to support@patienttrac.com!", vbInformation, "PatientTrac Security Management"
'                Screen.MousePointer = vbDefault
'                Exit Sub
'            End If
'        End If
'    End If
    If Trim(txtValue.Text) = vbNullString Then
        MsgBox "Enter a term first.", vbExclamation, "PatientTrac Controlled Vocabulary Management"
        Exit Sub
    End If
    txtValue.Text = Replace(txtValue.Text, "'", "''")
    lngPrim = grdMSETerms.GetRowItemData(grdMSETerms.ActiveRow)
    If lngPrim > 0 Then
        If cnTerms.State = adStateClosed Then cnTerms.open psConnect(1)
        strSQL = "UPDATE ATTR_VALID_VALUE SET ATTRBTE_VALID_VALUE = '" & txtValue.Text & "' WHERE ATTR_ID  = " & lngPrim
        Set rsTerms = cnTerms.Execute(strSQL)
    End If
    grdMSETerms.SetRowItemData grdMSETerms.ActiveRow, lngPrim
    Format_CellTypeEdit grdMSETerms, grdMSETerms.ActiveRow, 1, "1", LCase(txtValue.Text), "TypeHAlignLeft"
    grdMSETerms.SetCellBorder 1, grdMSETerms.ActiveRow, 1, grdMSETerms.ActiveRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    txtValue.Text = vbNullString
    If cnTerms.State = adStateOpen Then
        cnTerms.Close
        Set cnTerms = Nothing
        Set rsTerms = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Replace_Terms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Replace_Terms", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub cmdSave_Click()
    Write_MSE_Terms
End Sub



Private Sub cmdSelectTerm_Click()
    Write_Terms
    fraMseTerms.Visible = False
    Exit Sub
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
    Format_GridTerms
   ' Get_MSE_Terms
    cctProv.Text = clsProvider.Name
    mbooByPass = False
    centerForm Me
    Screen.MousePointer = vbDefault
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

Private Sub grdMSETerms_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim strHeader As String
Dim varCellValue

    If Row <= 0 Then Exit Sub
    With grdMSETerms
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
        .GetText 1, Row, varCellValue
        txtValue = Trim(varCellValue)
    End With
    Exit Sub
End Sub


Private Sub grdTerms_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim strHeader As String
Dim varCellValue

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
        .GetText 1, Row, varCellValue
        strHeader = Trim(varCellValue)
        Format_Grid_Field strHeader
        .GetText 1, Row, varCellValue
        strHeader = Trim(varCellValue)
        txtValue = vbNullString
        If strHeader <> vbNullString Then
            Format_Grid_Field strHeader
            Select Case strHeader
                 Case "Sleep"
                     Get_GridTerms "sd_MSE_Sleep"
                 Case "Voice"
                     Get_GridTerms "sd_MSE_Voice"
                 Case "Patient Appears"
                     Get_GridTerms "sd_MSE_PatientAppears"
                 Case "Nutrional Status"
                     Get_GridTerms "sd_MSE_NuritionStat"
                 Case "Speech"
                     Get_GridTerms "sd_MSE_Speech"
                 Case "Intelligence Estimate"
                     Get_GridTerms "sd_MSE_IntellEstimate"
                 Case "Habitus"
                     Get_GridTerms "sd_MSE_Habitus"
                 Case "Gait"
                     Get_GridTerms "sd_MSE_Gait"
                 Case "Station"
                     Get_GridTerms "sd_MSE_Station"
                 Case "Fund Of Knowledge"
                     Get_GridTerms "sd_MSE_FundKnow"
                 Case "Eye Contact"
                     Get_GridTerms "sd_MSE_EyeContact"
                 Case "Conation"
                     Get_GridTerms "sd_MSE_Conation"
                 Case "Judgment"
                     Get_GridTerms "sd_MSE_Judgment"
                 Case "Association"
                     Get_GridTerms "sd_MSE_Association"
                 Case "Appearance", "Attire"
                     Get_GridTerms "sd_MSE_Appearance"
                 Case "Orientation"
                     Get_GridTerms "sd_MSE_Orientation"
                 Case "Alertness"
                     Get_GridTerms "sd_MSE_Alertness"
                 Case "Remarkable Features"
                     Get_GridTerms "sd_MSE_RemFeatures"
                 Case "Strength Area Assessed"
                     Get_GridTerms "sd_MSE_AssessedIn"
                 Case "Tone Area Assessed"
                     Get_GridTerms "sd_MSE_AssessedIn"
                 Case "Verbal Auditory Content"
                     Get_GridTerms "sd_MSE_VerbalAuditory"
                 Case "Attitude Behavior"
                     Get_GridTerms "sd_MSE_AttituteBehavior"
                 Case "Thought Content"
                     Get_GridTerms "sd_MSE_ThoughtContent"
                 Case "Insight"
                     Get_GridTerms "sd_MSE_Insight"
                 Case "Mood"
                     Get_GridTerms "sd_MSE_Mood"
                 Case "Nutritional Status"
                     Get_GridTerms "sd_MSE_NuritionStat"
                 Case "Affect"
                     Get_GridTerms "sd_MSE_Affect"
                 Case "Speech"
                     Get_GridTerms "sd_MSE_Speech"
                 Case "Associations"
                     Get_GridTerms "sd_MSE_Association"
                 Case "Thought Process"
                     Get_GridTerms "sd_MSE_ThoughtProc"
                 Case "Abnormal Movements"
                     Get_GridTerms "sd_MSE_AbnormalMovements"
                 Case "Hallucination"
                     Get_GridTerms "sd_MSE_Hallucinations"
                 Case "Delusions"
                     Get_GridTerms "sd_MSE_Delusions"
                 Case "Behavior"
                     Get_GridTerms "sd_MSE_Behavior"
             End Select
             fraMseTerms.Visible = True
        End If
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


