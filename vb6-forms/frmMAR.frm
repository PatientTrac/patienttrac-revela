VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmMAR 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PatientTrac Medication Adminstration Record"
   ClientHeight    =   7905
   ClientLeft      =   6615
   ClientTop       =   3105
   ClientWidth     =   13845
   ControlBox      =   0   'False
   Icon            =   "frmMAR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMAR.frx":063E
      BackColorDown   =   15133676
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      PictureDisabled =   "frmMAR.frx":0F18
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmMAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private mlngFormID As Long
Public Property Let FormID(ByVal vlngNewValue As Long)
    mlngFormID = vlngNewValue
End Property

Public Property Get FormID() As Long
    FormID = mlngFormID
End Property


Private Sub cmdClose_Click()
    Dim iAnswer As Integer
    
    On Error GoTo Error_Handler
 

    Unload Me
    
    Exit Sub
    
Error_Handler:
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMDI.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
  On Error Resume Next

    Unload Me
    Exit Sub
End Sub


