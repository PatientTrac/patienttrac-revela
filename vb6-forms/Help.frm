VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C08D2A54-E5CB-4AFE-8AAD-B3E08A43486C}#1.3#0"; "XButton.ocx"
Begin VB.Form frmHelp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Help"
   ClientHeight    =   7440
   ClientLeft      =   2850
   ClientTop       =   885
   ClientWidth     =   12405
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   12405
   Begin RichTextLib.RichTextBox txtHelp 
      Height          =   6285
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   11086
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Help.frx":0742
   End
   Begin MSComctlLib.ImageList ImgListText 
      Left            =   10920
      Top             =   6780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":07C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":08D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":0B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":0F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":1066
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":1178
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":128A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":139C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Help.frx":14AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   11790
      Top             =   6810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbarText 
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImgListText"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Color"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Font"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Left"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Center"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Right"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Find"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Select All"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin XBtn.XButton cmdSave 
      Height          =   435
      Left            =   90
      TabIndex        =   2
      Top             =   6840
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      BackColor       =   8438015
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   6
      Caption         =   "Save"
      Picture         =   "Help.frx":15C0
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
   Begin XBtn.XButton cmdClose 
      Height          =   435
      Left            =   4200
      TabIndex        =   3
      Top             =   6840
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   767
      BackColor       =   8438015
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   128
      Bevel           =   6
      Caption         =   "Close"
      Picture         =   "Help.frx":15DC
      PictAlign       =   4
      CaptionAlign    =   4
      PushTimer       =   0   'False
      PushInterval    =   300
      MousePointer    =   0
      PicRaster       =   12583114
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fstrRTFFile As String
Private fstrFind As String
Private flngFind As Long
Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdSave_Click()
    If fstrRTFFile <> "" Then
        txtHelp.SaveFile fstrRTFFile
    End If
End Sub

Private Sub Form_Load()
    centerForm Me
    Me.Caption = "Help for " & psHelpTitle
    
    If Not FileExists(App.Path & "\" & psHelpFile) Then
        txtHelp.SaveFile App.Path & "\" & psHelpFile
    End If
    
    'If psHelpFile<> vbnullstring And FileExists(App.Path & "\" & psHelpFile) Then
    If psHelpFile <> vbNullString Then
        txtHelp.LoadFile App.Path & "\" & psHelpFile
        fstrRTFFile = App.Path & "\" & psHelpFile
    End If
    
    'KeepOnTop Me, True
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    txtHelp.Width = Me.ScaleWidth - (txtHelp.Left * 2)
    txtHelp.Height = Me.ScaleHeight - cmdClose.Height - 600 '(txtHelp.Top * 2)
    cmdClose.Top = txtHelp.Height + (txtHelp.Top) + 100
    cmdClose.Left = (Me.ScaleWidth - cmdClose.Width) / 2
    cmdSave.Top = cmdClose.Top
End Sub


Private Sub tbarText_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    
        Case 1
            If txtHelp.SelBold Then
                txtHelp.SelBold = False
            Else
                txtHelp.SelBold = True
            End If
            
        Case 2
            If txtHelp.SelItalic Then
                txtHelp.SelItalic = False
            Else
                txtHelp.SelItalic = True
            End If
        
        Case 3
            If txtHelp.SelUnderline Then
                txtHelp.SelUnderline = False
            Else
                txtHelp.SelUnderline = True
            End If
            
        Case 4
            On Error GoTo Cancel_Error
            CommonDialog.CancelError = True
            CommonDialog.flags = 0
            CommonDialog.ShowColor
            txtHelp.SelColor = CommonDialog.Color
                
        Case 5
            On Error GoTo Cancel_Error
            CommonDialog.flags = cdlCFBoth
            CommonDialog.CancelError = True
            CommonDialog.ShowFont
            txtHelp.SelFontName = CommonDialog.FontName
            txtHelp.SelFontSize = CommonDialog.FontSize
            
        Case 7
            txtHelp.SelAlignment = rtfLeft
            
        Case 8
            txtHelp.SelAlignment = rtfCenter
            
        Case 9
            txtHelp.SelAlignment = rtfRight
            
            
        Case 11
            fstrFind = txtHelp.SelText
            fstrFind = InputBox$("Search for >", "Find", fstrFind)
            If Len(fstrFind) > 0 Then
                txtHelp.SetFocus
                txtHelp.SelLength = 0
                flngFind = txtHelp.Find(fstrFind)
            End If
        
        Case 12
            If txtHelp.SelLength = 0 Or txtHelp.SelLength < Len(txtHelp.Text) + 2 Then
                txtHelp.SelStart = 0
                txtHelp.SelLength = Len(txtHelp.TextRTF)
            Else
                txtHelp.SelLength = 0
                txtHelp.SelText = ""
            End If
        
            If txtHelp.Enabled Then txtHelp.SetFocus
        
    End Select

Cancel_Error:
    Exit Sub
End Sub


Private Sub txtHelp_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lngStart As Long

    If KeyCode = 114 Then  'F3 key
        lngStart = flngFind + Len(fstrFind)
        
        If lngStart >= Len(txtHelp) Then lngStart = 0
        
        flngFind = txtHelp.Find(fstrFind, lngStart)
        
        If flngFind = -1 Then
            lngStart = 0
            flngFind = txtHelp.Find(fstrFind, lngStart)
        End If
    End If
End Sub


