VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmAssignRoom 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Room Number"
   ClientHeight    =   3990
   ClientLeft      =   7290
   ClientTop       =   4155
   ClientWidth     =   4860
   ControlBox      =   0   'False
   Icon            =   "frmAssignRoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin FPSpreadADO.fpSpread grdData 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4875
      _Version        =   458752
      _ExtentX        =   8599
      _ExtentY        =   5741
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   -2147483645
      OperationMode   =   2
      RowHeaderDisplay=   0
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      SpreadDesigner  =   "frmAssignRoom.frx":08CA
      UserResize      =   0
   End
   Begin PT_XP_Button.PT_XP cmdSelect 
      Height          =   405
      Left            =   1320
      TabIndex        =   1
      Top             =   3450
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   714
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
      Picture         =   "frmAssignRoom.frx":0C13
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Select"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmAssignRoom.frx":0FAD
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdCancel 
      Height          =   405
      Left            =   2490
      TabIndex        =   2
      Top             =   3450
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   714
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
      Picture         =   "frmAssignRoom.frx":12FF
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Cancel"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmAssignRoom.frx":1699
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmAssignRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mFloor As String
Public mWing As String
Public mRoom As String
Public mBed As String
Public mPhone As String
Private Sub Format_Grid()

    Dim Idx As Integer
    Dim strSQL As String
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    On Error GoTo Err_Format_Grid

    Screen.MousePointer = vbHourglass

    With grdData
        .ColsFrozen = 1
        .MaxCols = 5
        .MaxRows = 0
        
        .ColWidth(0) = 5
        .SetText 0, 0, ""
        
        .ColWidth(1) = 8
        .SetText 1, 0, "Floor"
        
        .ColWidth(2) = 8
        .SetText 2, 0, "Wing"
       
        .ColWidth(3) = 8
        .SetText 3, 0, "Room"
        
        .ColWidth(4) = 8
        .SetText 4, 0, "Bed"
        
        .ColWidth(5) = 8
        .SetText 5, 0, "Phone"
        .Col = 5
        .ColHidden = True
    End With
    
    Screen.MousePointer = vbDefault

    Exit Sub

Err_Format_Grid:
    Screen.MousePointer = vbDefault

    ErrHandler _
     "frmAssignRoom.Format_Grid", _
     Err.Number, _
     Err.Description
End Sub
Private Sub cmdCancel_Click()
    Me.Hide
End Sub
Private Sub cmdSelect_Click()
    Dim vValue
    
    On Error GoTo Error_Handler
    
    If grdData.ActiveRow < 1 Then
       MsgBox "Please select row first", vbInformation
       Exit Sub
    End If

    grdData.GetText 1, grdData.ActiveRow, vValue
    mFloor = vValue
    
    grdData.GetText 2, grdData.ActiveRow, vValue
    mWing = vValue
    
    grdData.GetText 3, grdData.ActiveRow, vValue
    mRoom = vValue
    
    grdData.GetText 4, grdData.ActiveRow, vValue
    mBed = vValue
    
    grdData.GetText 5, grdData.ActiveRow, vValue
    mPhone = vValue

    Me.Hide
    
    Exit Sub
    
Error_Handler:
    ErrHandler _
     "frmAssignRoom.cmdSelect_Click", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub Form_Load()
    mFloor = ""
    mWing = ""
    mRoom = ""
    mBed = ""

    Format_Grid
    Load_Grid
End Sub

Private Sub Load_Grid()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim strFirstName As String
    Dim strLastName As String
    
    Dim Cnt As Integer
    
    Dim sValue
    
    On Error GoTo LoadGrid_Error:

    Cnt = 1

    cn.open psConnect(0)
    
    cn.s_OpenRooms rs
    
    With rs
       Do Until .EOF
          grdData.MaxRows = grdData.MaxRows + 1
          grdData.Row = grdData.MaxRows
          grdData.SetText 0, grdData.Row, Cnt
          grdData.Col = 1
          grdData.CellType = CellTypeStaticText
          grdData.SetText 1, grdData.Row, (!Floor) & ""
          grdData.Col = 2
          grdData.CellType = CellTypeStaticText
          grdData.SetText 2, grdData.Row, (!Wing) & ""
          grdData.Col = 3
          grdData.CellType = CellTypeStaticText
          grdData.SetText 3, grdData.Row, (!room) & ""
          grdData.Col = 4
          grdData.CellType = CellTypeStaticText
          grdData.SetText 4, grdData.Row, (!Bed) & ""
          grdData.Col = 5
          grdData.CellType = CellTypeStaticText
          grdData.SetText 5, grdData.Row, (!Telephone) & ""
          
          Cnt = Cnt + 1
          .MoveNext
       Loop
       .Close
    End With
    
    Set rs = Nothing
    
    cn.Close
    Set cn = Nothing
    
    Exit Sub

LoadGrid_Error:
    ErrHandler _
     "frmAssignRoom.Load_Grid", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl

End Sub


