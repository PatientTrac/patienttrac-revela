VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmTB 
   Caption         =   "Development Only"
   ClientHeight    =   2325
   ClientLeft      =   1920
   ClientTop       =   4950
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   12045
   Begin VB.Frame fraTB 
      BackColor       =   &H00DDCDCD&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin PT_XP_Button.PT_XP cmdRevEnctr 
         Height          =   495
         Left            =   2220
         TabIndex        =   2
         ToolTipText     =   "Previous, Next, or Current Encounter"
         Top             =   30
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   873
         BackColor       =   14536141
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceMode  =   2
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   14215660
         Caption         =   "Review Encounters"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdPatient 
         Height          =   495
         Left            =   30
         TabIndex        =   1
         ToolTipText     =   "Select / Change Patients"
         Top             =   30
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   873
         BackColor       =   14536141
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   14215660
         Caption         =   "Patient"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP PT_XP1 
         Height          =   495
         Left            =   7950
         TabIndex        =   3
         ToolTipText     =   "Save, Restore, Close"
         Top             =   30
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   873
         BackColor       =   14536141
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceMode  =   2
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   14215660
         Caption         =   "Clinical Data              Close"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP PT_XP2 
         Height          =   495
         Left            =   5640
         TabIndex        =   4
         ToolTipText     =   "Review Scanned Chart Data"
         Top             =   30
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   873
         BackColor       =   14536141
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   14215660
         Caption         =   "History "
         DepthEvent      =   1
      End
   End
   Begin VB.Menu Encounter 
      Caption         =   "Encounter"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu Enctr_Item1 
         Caption         =   "Next Encounter"
      End
      Begin VB.Menu Enctr_Item2 
         Caption         =   "Previous Encounter"
      End
      Begin VB.Menu Enctr_Item3 
         Caption         =   "Current Encounter"
      End
   End
   Begin VB.Menu Clinical_Data 
      Caption         =   "Clinical_Data"
      Visible         =   0   'False
      Begin VB.Menu Clinical_Item1 
         Caption         =   "Save Clinical Data"
      End
      Begin VB.Menu Clinical_Item2 
         Caption         =   "Close Clinical Data"
         Visible         =   0   'False
      End
      Begin VB.Menu Clinical_Item3 
         Caption         =   "Restore Clinical Data"
      End
   End
End
Attribute VB_Name = "frmTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEncounter_Click()
    
    
End Sub


Private Sub cmdEncounter_DropDownClick(Index As Integer)
        
End Sub


Private Sub PT_Toolbar_Click()

End Sub


Private Sub cmdRevEnctr_DropDownClick()

    'PopupMenu Encounter, , cmdRevEnctr.Left + fraTB.Left, cmdRevEnctr.Top + cmdRevEnctrHeight + fraTB.Top
    PopupMenu Encounter, , cmdRevEnctr.Left + fraTB.Left, fraTB.Top + cmdRevEnctr.Height '+ fraTB.Top

End Sub


Private Sub Enctr_Item1_Click()
    MsgBox "Next Encounter Coming", vbInformation
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.Objects
      Set obj = Nothing
    Next obj
  Exit Sub

End Sub


