VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{76F358A5-0FAA-4EA4-8073-9E9D9A484F37}#1.0#0"; "CImageBufferX.ocx"
Begin VB.Form frmImageCapture 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Patient Image"
   ClientHeight    =   5310
   ClientLeft      =   8835
   ClientTop       =   3705
   ClientWidth     =   4005
   ControlBox      =   0   'False
   Icon            =   "ImageCapture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   Begin CIMAGEBUFFERXLib.CImageBufferX cimbx 
      Height          =   2955
      Left            =   90
      TabIndex        =   3
      Top             =   630
      Width           =   3825
      _Version        =   65536
      _ExtentX        =   6747
      _ExtentY        =   5212
      _StockProps     =   0
   End
   Begin PT_XP_Button.PT_XP cmdCapture 
      Height          =   465
      Left            =   1290
      TabIndex        =   1
      ToolTipText     =   "Click to take picture"
      Top             =   60
      Width           =   1425
      _ExtentX        =   2514
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
      Picture         =   "ImageCapture.frx":08CA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Take Picture"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "ImageCapture.frx":11A4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdDone 
      Height          =   345
      Left            =   1515
      TabIndex        =   2
      ToolTipText     =   "Close and do not Save"
      Top             =   4860
      Width           =   975
      _ExtentX        =   1720
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
      Picture         =   "ImageCapture.frx":1DF6
      AppearanceThemes=   3
      BorderDrawEvent =   1
      TransparentColor=   14737632
      Caption         =   "Done"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "ImageCapture.frx":2190
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblPhotoPrivacy 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Photo privacy statement"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   90
      TabIndex        =   0
      Top             =   3840
      Width           =   3795
   End
End
Attribute VB_Name = "frmImageCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pfstrFile As String
Private flngWidth As Long
Private flngHeight As Long
Private mFirst As Boolean


Private Sub cmdCapture_Click()
Dim fso As New FileSystemObject
   
    On Error GoTo Err_Picture
    Screen.MousePointer = vbHourglass
    If fso.FileExists(pfstrFile) Then
       fso.DeleteFile pfstrFile
    End If
    Display_Video cimbx
    Grab_Image cimbx, 0
    Save_Active_Image cimbx, pfstrFile, 0, 0
    Disable_Video cimbx
    Screen.MousePointer = vbNormal
    Exit Sub
Err_Picture:
   Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, pfstrFile, "Err_Picture", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
Dim lstrSQL As String
Dim varRS
Dim lbolRet As Boolean
Dim llngRows As Long
Dim fso As New FileSystemObject

    On Error GoTo Error_Form_Activate
    centerForm Me
    If mFirst Then
       mFirst = False
       Get_PrivacyStatement
       If fso.FileExists(pfstrFile) Then
          Load_Image cimbx, pfstrFile
       Else
         If fso.FileExists(App.Path & "\BlankPhoto.jpg") Then
            Load_Image cimbx, App.Path & "\BlankPhoto.jpg"
         End If
       End If
    End If
    Exit Sub
Error_Form_Activate:
   Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, pfstrFile, "Form_Activate", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub
Private Sub Form_Load()
    mFirst = True
End Sub
Public Sub Get_PrivacyStatement()
Dim cnPriv As New ADODB.Connection
Dim rsPriv As New ADODB.Recordset
    
    On Error GoTo Error_Get_PrivacyStatement
  
    If cnPriv.State = adStateClosed Then cnPriv.open psConnect(1)
   
    'Call stored procedure to retrieve data values for combo
    cnPriv.S_ATTR_VALID_VALUE "photo_privacy_stmnt", rsPriv
    
    lblPhotoPrivacy.Visible = False
    
    With rsPriv
       If Not .EOF Then
          lblPhotoPrivacy.Caption = (!ATTRBTE_VALID_VALUE) & ""
          lblPhotoPrivacy.Visible = True
       End If
       .Close
    End With
    
    If cnPriv.State = adStateOpen Then
       Set rs = Nothing
       cnPriv.Close
        Set cnPriv = Nothing
    End If
    If Not lblPhotoPrivacy.Visible Then
       cmdDone.Top = cimbx.Top + cimbx.Height + 100
       Me.Height = cmdDone.Top + cmdDone.Height + 500
    End If

    Exit Sub
    
Error_Get_PrivacyStatement:
    MsgBox Err.Description, vbInformation
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

  Exit Sub
End Sub


