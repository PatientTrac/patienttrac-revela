VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{6514F5A0-641C-11D2-9FD0-0020AF131A57}#3.0#0"; "fpFlp30.ocx"
Begin VB.Form frmRoomLocations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Room Locations Maintenance"
   ClientHeight    =   7440
   ClientLeft      =   6255
   ClientTop       =   3150
   ClientWidth     =   11070
   ControlBox      =   0   'False
   Icon            =   "frmRoomLocations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin LpADOLib.fpComboAdo cboLocation 
      Height          =   360
      Left            =   1020
      TabIndex        =   1
      Top             =   330
      Width           =   5025
      _Version        =   196608
      _ExtentX        =   8864
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   0
      Text            =   ""
      Columns         =   0
      Sorted          =   2
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   3
      RowHeight       =   -1
      WrapList        =   -1  'True
      WrapWidth       =   0
      AutoSearch      =   1
      SearchMethod    =   2
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   14737632
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   14737632
      BorderDropShadowWidth=   0
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   1
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   1000
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   1
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   -1  'True
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   14737632
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      DataMemberList  =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   0   'False
      EditAlignH      =   0
      EditAlignV      =   0
      AllowAnimate    =   -1  'True
      ColDesigner     =   "frmRoomLocations.frx":000C
   End
   Begin FPSpreadADO.fpSpread grdData 
      DragIcon        =   "frmRoomLocations.frx":0307
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   780
      Width           =   10995
      _Version        =   458752
      _ExtentX        =   19394
      _ExtentY        =   10081
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      ColHeaderDisplay=   0
      ColsFrozen      =   1
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
      MaxCols         =   1
      MaxRows         =   0
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   12648447
      ShadowDark      =   8388608
      ShadowText      =   8388608
      SpreadDesigner  =   "frmRoomLocations.frx":0749
      UnitType        =   0
      UserResize      =   0
      Appearance      =   2
      TextTip         =   3
      TextTipDelay    =   25
      ScrollBarTrack  =   1
   End
   Begin VB.Label lblLocation 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Facility:"
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
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   825
   End
End
Attribute VB_Name = "frmRoomLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub Form_Load()
        
    On Error GoTo Error_Handler
    
    Screen.MousePointer = vbHourglass
    Format_Grid
    centerForm Me
    Me.Top = Me.Top + 300
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Handler:

End Sub
Private Sub LoadTextValues()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sPhone1 As String
Dim varPhone
Dim vId
Dim Idx As Integer
Dim sPhone As String
    
    On Error GoTo LoadTextValues_Error
    
'    grdData.GetText 8, grdData.ActiveRow, vId
'
'    cn.open psConnect(0)
'
'    cn.s_RoomLocInfo vId, rs
'
'    With rs
'       Do While Not .EOF
'          txtFloor.Text = Trim((!Floor)) & ""
'          txtWing.Text = Trim((!Wing)) & ""
'          txtRoom.Text = Trim((!room)) & ""
'          txtBed.Text = Trim((!Bed)) & ""
'
'          If InStr((!Telephone) & "", "-") Then
'             varPhone = Split((!Telephone), "-")
'             txtPhone1.Text = varPhone(0)
'             txtPhone2.Text = varPhone(1)
'             txtPhone3.Text = varPhone(2)
'             Erase varPhone
'          Else
'             sPhone = Trim((!Telephone) & "")
'             If Len(sPhone) > 0 Then
'                txtPhone1.Text = Left(sPhone, 3)
'                If Len(sPhone) > 3 Then
'                   sPhone = Mid(sPhone, 4)
'                   txtPhone2.Text = Left(sPhone, 3)
'                   If Len(sPhone) > 3 Then
'                      txtPhone3.Text = Mid(sPhone, 4)
'                   End If
'                End If
'             End If
'          End If
'
'          .MoveNext
'       Loop
'       .Close
'    End With
'
'    Set rs = Nothing
'
'    cn.Close
'    Set cn = Nothing
'
'    mFloor = txtFloor.Text
'    mWing = txtWing.Text
'    mRoom = txtRoom.Text
'    mBed = txtBed.Text
'    mPhone = Trim(txtPhone1.Text) & "-" & Trim(txtPhone2.Text) & "-" & Trim(txtPhone3.Text)
    Exit Sub
    
LoadTextValues_Error:
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub



Private Function SaveData() As Boolean
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strFloor As String
Dim strWing As String
Dim strRoom As String
Dim strBed As String
Dim sPhone As String
Dim varID
Dim Idx As Integer
    
    On Error GoTo SaveData_Error
    Screen.MousePointer = vbHourglass
    
'    strFloor = Trim(Replace(txtFloor.Text, "'", "''"))
'    strWing = Trim(Replace(txtWing.Text, "'", "''"))
'    strRoom = Trim(Replace(txtRoom.Text, "'", "''"))
'    strBed = Trim(Replace(txtBed.Text, "'", "''"))
'    sPhone = Trim(txtPhone1.Text) & "-" & Trim(txtPhone2.Text) & "-" & Trim(txtPhone3.Text)
'    sPhone = Replace(sPhone, "'", "''")
'    cn.open psConnect(0)
'    If fbAdd Then
'       cn.i_RoomLocation strFloor, strWing, strRoom, strBed, sPhone
'
'       grdData.MaxRows = grdData.MaxRows + 1
'       grdData.Row = grdData.MaxRows
'
'       grdData.SetText 0, grdData.Row, grdData.MaxRows
'       grdData.SetText 1, grdData.Row, strFloor
'       grdData.SetText 2, grdData.Row, strWing
'       grdData.SetText 3, grdData.Row, strRoom
'       grdData.SetText 4, grdData.Row, strBed
'       grdData.SetText 5, grdData.Row, sPhone
'
'       cn.s_RoomLocationID strFloor, strWing, strRoom, strBed, rs
'       With rs
'          If Not .EOF Then
'             grdData.SetText 8, grdData.Row, (!LocID)
'          Else
'             grdData.SetText 0, grdData.Row, "???"
'             MsgBox "Problem retrieving Room Location ID", vbExclamation
'             Exit Function
'          End If
'          .Close
'       End With
'       txtFloor.SetFocus
'       cmdSave.Enabled = False
'    Else
'       grdData.GetText 8, grdData.ActiveRow, varID
'       cn.u_RoomLocation strFloor, strWing, strRoom, strBed, sPhone, CStr(varID)
'
'       grdData.SetText 1, grdData.ActiveRow, strFloor
'       grdData.SetText 2, grdData.ActiveRow, strWing
'       grdData.SetText 3, grdData.ActiveRow, strRoom
'       grdData.SetText 4, grdData.ActiveRow, strBed
'       grdData.SetText 5, grdData.ActiveRow, sPhone
'       fbAdd = False
'    End If
'    Set rs = Nothing
'    cn.Close
'    Set cn = Nothing
    Screen.MousePointer = vbDefault
    SaveData = True
    Exit Function
SaveData_Error:
    Screen.MousePointer = vbDefault
    SaveData = False
     Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function
Private Sub Format_Grid()
    Dim Idx As Integer
    Dim strSQL As String
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    On Error GoTo Err_Format_Grid
    Screen.MousePointer = vbHourglass
    With grdData
'        .ColsFrozen = 1
'        .MaxCols = 8
'        .MaxRows = 0
'
'        .ColWidth(0) = 5
'        .SetText 0, 0, ""
'
'        .ColWidth(1) = 8
'        .SetText 1, 0, "Floor"
'
'        .ColWidth(2) = 8
'        .SetText 2, 0, "Wing"
'
'        .ColWidth(3) = 8
'        .SetText 3, 0, "Room"
'
'        .ColWidth(4) = 8
'        .SetText 4, 0, "Bed"
'
'        .ColWidth(5) = 12
'        .SetText 5, 0, "Telephone"
'
'        .ColWidth(6) = 12
'        .SetText 6, 0, "Patient ID"
'
'        .ColWidth(7) = 28
'        .SetText 7, 0, "Patient"
'
'        .ColWidth(8) = 7
'        .SetText 8, 0, "ID"
'        .Col = 8
'        .ColHidden = True
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
Err_Format_Grid:
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub
Private Function Validate_Entries() As Boolean
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strFloor As String
Dim strWing As String
Dim strRoom As String
Dim strBed As String
Dim chkRec As Boolean


    On Error GoTo Error_Handler
    
    Validate_Entries = True
    
'    If Trim(txtRoom.Text) = "" Then
'       MsgBox "Room is Required!", vbInformation + vbOKCancel
'       Validate_Entries = False
'       Exit Function
'    End If
    
'    If Trim(txtBed.Text) = "" Then
'       MsgBox "Bed is Required!", vbInformation + vbOKCancel
'       Validate_Entries = False
'       Exit Function
'    End If
'
'    strFloor = Trim(Replace(txtFloor.Text, "'", "''"))
'    strWing = Trim(Replace(txtWing.Text, "'", "''"))
'    strRoom = Trim(Replace(txtRoom.Text, "'", "''"))
'    strBed = Trim(Replace(txtBed.Text, "'", "''"))
    
'    chkRec = True
'    If Not fbAdd And _
'       mFloor = Trim(txtFloor.Text) And _
'       mWing = Trim(txtWing.Text) And _
'       mRoom = Trim(txtRoom.Text) And _
'       mBed = Trim(txtBed.Text) Then
'       chkRec = False
'    End If
'    If chkRec Then
       cn.open psConnect(0)
    
       cn.s_RoomLocationExist UCase(strFloor), UCase(strWing), UCase(strRoom), UCase(strBed), rs
       With rs
          If Not .EOF Then
             MsgBox "Entry already exists!", vbExclamation
             Validate_Entries = False
          End If
           .Close
       End With
       Set rs = Nothing
    
       cn.Close
       Set cn = Nothing
   ' End If
    Exit Function
Error_Handler:
     Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

  Unload Me

  Exit Sub
End Sub

