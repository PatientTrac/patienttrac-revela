VERSION 5.00
Begin VB.Form frmSelPractice 
   BackColor       =   &H0065D89E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Practice"
   ClientHeight    =   1845
   ClientLeft      =   7995
   ClientTop       =   2370
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3885
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1530
      TabIndex        =   2
      Top             =   1230
      Width           =   885
   End
   Begin VB.PictureBox picPractice 
      BackColor       =   &H0080FFFF&
      Height          =   675
      Left            =   390
      ScaleHeight     =   615
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   360
      Width           =   3135
      Begin VB.OptionButton optPractice 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   2745
      End
   End
End
Attribute VB_Name = "frmSelPractice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fintPracticeID() As Integer
Public mstrPractice As String
Private Sub cmdOK_Click()
Dim i As Integer

    LogFile "frmSelPractice cmdOK"
    If PracticeSelected(i) Then
        Unload Me
        mstrPractice = PracticeSelected(i)
        Set_Practice fintPracticeID(i), i
        Else
        MsgBox "Select a practice first.", vbExclamation
    End If
    
End Sub

Private Sub Form_Activate()
    Load_Practices
    centerForm Me
End Sub

Private Sub Load_Practices()
Dim strSQL As String
Dim vRS
Dim iRows As Long
Dim R As Integer
Dim c As Integer
Dim bRet As Boolean
Dim lstrRemarks As String

On Error GoTo Error_Load_Practices

strSQL = "Select EMR_PRACTICE, EMR_GRP_ID, EMR_SPECIALTY " _
    & "From EMR_GROUP " _
    & "Where EMR_GRP_NAME = 'Clinical Data' and ACTIVE = 1 " _
    & "Order By EMR_PRACTICE"
    
    lstrRemarks = strSQL
   
    bRet = SelectQuery(strSQL, psConnect(1), vRS, iRows)

    If iRows > 0 Then
        ReDim fintPracticeID(UBound(vRS, 2))
        For R = LBound(vRS, 2) To UBound(vRS, 2)
            If R > 0 Then
                Load optPractice(R)
                optPractice(R).Visible = True
                optPractice(R).Caption = vRS(0, R)
                optPractice(R).Top = optPractice(R - 1).Top + optPractice(R - 1).Height + 30
                picPractice.Height = picPractice.Height + optPractice(R - 1).Height
                cmdOK.Top = cmdOK.Top + optPractice(R - 1).Height
                Me.Height = Me.Height + optPractice(R - 1).Height
                fintPracticeID(R) = vRS(1, R)
            Else
                optPractice(0).Caption = vRS(0, 0)
                fintPracticeID(0) = vRS(1, 0)
            End If
        Next R
    End If
    
    picPractice.Height = picPractice.Height + optPractice(R - 1).Height
    cmdOK.Top = cmdOK.Top + optPractice(R - 1).Height
    Me.Height = Me.Height + optPractice(R - 1).Height

    Exit Sub
    
Error_Load_Practices:
    ErrHandler _
     Me.Name & " Load_Practices", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Function PracticeSelected(vintID As Integer) As Boolean
Dim i As Integer

    For i = optPractice.LBound To optPractice.UBound
        If optPractice(i).Value Then
            vintID = i
            PracticeSelected = True
            Exit For
        End If
    Next i
    
End Function

Private Sub Set_Practice(vintPracticeID As Integer, vintIndex As Integer)
End Sub

Private Sub Form_Paint()
    MonochromeBurst Me, 0, 170, 300, 5
End Sub


Private Sub Form_Unload(Cancel As Integer)
 
  On Error Resume Next

    Unload Me
    Exit Sub
  Exit Sub
End Sub


