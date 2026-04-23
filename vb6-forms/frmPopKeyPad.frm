VERSION 5.00
Object = "{C08D2A54-E5CB-4AFE-8AAD-B3E08A43486C}#1.3#0"; "XButton.ocx"
Begin VB.Form frmPopKeyPad 
   BorderStyle     =   0  'None
   Caption         =   "PopForm"
   ClientHeight    =   1545
   ClientLeft      =   7680
   ClientTop       =   2370
   ClientWidth     =   2595
   Icon            =   "frmPopKeyPad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   0
      Left            =   450
      TabIndex        =   2
      Top             =   1080
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "0"
      Picture         =   "frmPopKeyPad.frx":058A
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "1"
      Picture         =   "frmPopKeyPad.frx":05A6
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   2
      Left            =   450
      TabIndex        =   4
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "2"
      Picture         =   "frmPopKeyPad.frx":05C2
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   3
      Left            =   810
      TabIndex        =   5
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "3"
      Picture         =   "frmPopKeyPad.frx":05DE
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   4
      Left            =   90
      TabIndex        =   6
      Top             =   420
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "4"
      Picture         =   "frmPopKeyPad.frx":05FA
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   5
      Left            =   450
      TabIndex        =   7
      Top             =   420
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "5"
      Picture         =   "frmPopKeyPad.frx":0616
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   6
      Left            =   810
      TabIndex        =   8
      Top             =   420
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "6"
      Picture         =   "frmPopKeyPad.frx":0632
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   7
      Left            =   90
      TabIndex        =   9
      Top             =   750
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "7"
      Picture         =   "frmPopKeyPad.frx":064E
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   8
      Left            =   450
      TabIndex        =   10
      Top             =   750
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "8"
      Picture         =   "frmPopKeyPad.frx":066A
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   9
      Left            =   810
      TabIndex        =   11
      Top             =   750
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "9"
      Picture         =   "frmPopKeyPad.frx":0686
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   10
      Left            =   810
      TabIndex        =   12
      Top             =   1080
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "/"
      Picture         =   "frmPopKeyPad.frx":06A2
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   11
      Left            =   90
      TabIndex        =   13
      Top             =   1080
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   ","
      Picture         =   "frmPopKeyPad.frx":06BE
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   12
      Left            =   1170
      TabIndex        =   14
      Top             =   90
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "+"
      Picture         =   "frmPopKeyPad.frx":06DA
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   13
      Left            =   1170
      TabIndex        =   15
      Top             =   420
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "-"
      Picture         =   "frmPopKeyPad.frx":06F6
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   14
      Left            =   1170
      TabIndex        =   16
      Top             =   750
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   "<"
      Picture         =   "frmPopKeyPad.frx":0712
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdKey 
      Height          =   300
      Index           =   15
      Left            =   1170
      TabIndex        =   17
      Top             =   1080
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   5
      Caption         =   ">"
      Picture         =   "frmPopKeyPad.frx":072E
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdOK 
      Height          =   285
      Left            =   1650
      TabIndex        =   18
      Top             =   90
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8388608
      Bevel           =   6
      Caption         =   "OK"
      Picture         =   "frmPopKeyPad.frx":074A
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdCancel 
      Height          =   285
      Left            =   1650
      TabIndex        =   19
      Top             =   420
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8388608
      Bevel           =   6
      Caption         =   "Cancel"
      Picture         =   "frmPopKeyPad.frx":0766
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdBackspace 
      Height          =   285
      Left            =   1650
      TabIndex        =   0
      Top             =   750
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8388608
      Bevel           =   6
      Caption         =   "Bksp"
      Picture         =   "frmPopKeyPad.frx":0782
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdClear 
      Height          =   285
      Left            =   1650
      TabIndex        =   1
      Top             =   1080
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   8388608
      Bevel           =   6
      Caption         =   "Clear"
      Picture         =   "frmPopKeyPad.frx":079E
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin VB.Shape shpBack 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1515
      Left            =   0
      Top             =   30
      Width           =   2565
   End
End
Attribute VB_Name = "frmPopKeyPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mctlTarget As Control
Private mctlTarget2 As Object
Private mstrPreviousValue As String
Private mintMaxLength As Integer

Public Property Set Target(ByRef rctlNewReference As Control)

    Set mctlTarget = rctlNewReference
    mstrPreviousValue = mctlTarget.Text
    mctlTarget.Text = mstrPreviousValue
    mintMaxLength = mctlTarget.MaxLength
    
End Property


Public Property Set Target2(ByRef rctlNewReference As Object)
    Set mctlTarget2 = rctlNewReference
End Property


Private Sub cmdBackspace_Click()

    If Len(mctlTarget.Text) > 0 Then
        mctlTarget.Text = Left$(mctlTarget.Text, Len(mctlTarget.Text) - 1)
    End If
    
End Sub

Private Sub cmdCancel_Click()

    mctlTarget.Text = mstrPreviousValue
    mctlTarget2.Trigger_AfterUpdate
    Unload Me
    
End Sub

Private Sub cmdClear_Click()

    mctlTarget.Text = ""
    
End Sub

Private Sub cmdKey_Click(Index As Integer)
    If mintMaxLength > 0 And Len(mctlTarget.Text) >= mintMaxLength Then
       Beep
       Exit Sub
    End If

    'Select Case Index
'    Case 10
'        If Left((mctlTarget.Text), 1) = "/" Then
'            mctlTarget.Text = Mid(mctlTarget.Text, 2)
'        'Else
'        '    mctlTarget.Text = "-" & mctlTarget.Text
'        End If
'    Case 11
'        If InStr((mctlTarget.Text), ".") = 0 Then
'            mctlTarget.Text = mctlTarget.Text & cmdKey(Index).Caption
'        End If
'    Case 12
'        If InStr((mctlTarget.Text), "+") = 0 Then
'            mctlTarget.Text = mctlTarget.Text & cmdKey(Index).Caption
'        End If
'    Case 13
'        If InStr((mctlTarget.Text), ".") = 0 Then
'            mctlTarget.Text = mctlTarget.Text & cmdKey(Index).Caption
'        End If
'    Case Else
'        If mctlTarget.Text = "0" Then
'            mctlTarget.Text = ""
'        End If
        mctlTarget.Text = mctlTarget.Text & cmdKey(Index).Caption
   ' End Select
End Sub

Private Sub cmdOK_Click()
    mctlTarget2.Trigger_AfterUpdate
    Unload Me
    
    
End Sub

