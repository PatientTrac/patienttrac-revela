VERSION 5.00
Object = "{C08D2A54-E5CB-4AFE-8AAD-B3E08A43486C}#1.3#0"; "XButton.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.2#0"; "ImgEdit.ocx"
Object = "{E1A6B8A3-3603-101C-AC6E-040224009C02}#1.0#0"; "ImgThumb.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmViewer 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Viewer"
   ClientHeight    =   10785
   ClientLeft      =   3000
   ClientTop       =   1485
   ClientWidth     =   15270
   Icon            =   "Viewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   15270
   Begin VB.PictureBox picScale 
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   11445
      ScaleHeight     =   285
      ScaleWidth      =   1785
      TabIndex        =   18
      Top             =   9840
      Width           =   1845
      Begin XBtn.XButton cmd25 
         Height          =   285
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
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
         BorderColor     =   8388608
         Bevel           =   6
         Caption         =   "25%"
         Picture         =   "Viewer.frx":0442
         PictAlign       =   4
         CaptionAlign    =   4
         PushTimer       =   0   'False
         PushInterval    =   300
         MousePointer    =   0
         PicRaster       =   12583114
      End
      Begin XBtn.XButton cmd50 
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Top             =   0
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
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
         BorderColor     =   8388608
         Bevel           =   6
         Caption         =   "50%"
         Picture         =   "Viewer.frx":045E
         PictAlign       =   4
         CaptionAlign    =   4
         PushTimer       =   0   'False
         PushInterval    =   300
         MousePointer    =   0
         PicRaster       =   12583114
      End
      Begin XBtn.XButton cmd75 
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Top             =   0
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
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
         BorderColor     =   8388608
         Bevel           =   6
         Caption         =   "75%"
         Picture         =   "Viewer.frx":047A
         PictAlign       =   4
         CaptionAlign    =   4
         PushTimer       =   0   'False
         PushInterval    =   300
         MousePointer    =   0
         PicRaster       =   12583114
      End
   End
   Begin VB.HScrollBar scrHoriz 
      Height          =   255
      Left            =   2790
      TabIndex        =   17
      Top             =   9420
      Width           =   12345
   End
   Begin VB.VScrollBar scrVert 
      Height          =   9405
      Left            =   2550
      TabIndex        =   16
      Top             =   30
      Width           =   255
   End
   Begin VB.CheckBox chkHideTIF_list 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hide Navigation"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7410
      TabIndex        =   14
      Top             =   9870
      Width           =   1815
   End
   Begin VB.CheckBox chkActualSize 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Actual Size"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   13530
      TabIndex        =   11
      Top             =   9900
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.HScrollBar HVScale 
      Height          =   255
      LargeChange     =   10
      Left            =   11070
      Max             =   200
      Min             =   1
      SmallChange     =   2
      TabIndex        =   10
      Top             =   10260
      Value           =   1
      Width           =   3435
   End
   Begin VB.PictureBox picControls 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   30
      ScaleHeight     =   375
      ScaleWidth      =   9375
      TabIndex        =   7
      Top             =   10170
      Width           =   9435
      Begin PT_XP_Button.PT_XP cmdUpdatePic 
         Height          =   345
         Left            =   1170
         TabIndex        =   22
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "Viewer.frx":0496
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Viewer.frx":0830
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   345
         Left            =   8220
         TabIndex        =   23
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "Viewer.frx":0B82
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Viewer.frx":1594
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSavePic 
         Height          =   345
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "Viewer.frx":18E6
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Save as"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Viewer.frx":1C80
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRename 
         Height          =   345
         Left            =   3525
         TabIndex        =   25
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "Viewer.frx":1FD2
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Rename"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Viewer.frx":256C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdDelete 
         Height          =   345
         Left            =   4695
         TabIndex        =   26
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "Viewer.frx":28BE
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Delete"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Viewer.frx":2E58
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRotate 
         Height          =   345
         Left            =   5865
         TabIndex        =   27
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "Viewer.frx":31AA
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Rotate"
         Depth           =   5
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Viewer.frx":3744
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdPrint 
         Height          =   345
         Left            =   7050
         TabIndex        =   28
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "Viewer.frx":3A96
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Print"
         Depth           =   5
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Viewer.frx":4030
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdCopyTo 
         Height          =   345
         Left            =   2355
         TabIndex        =   29
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
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
         Picture         =   "Viewer.frx":4382
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Copy To"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "Viewer.frx":491C
         ShowFocus       =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6600
      Top             =   9690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picViewer 
      BackColor       =   &H00FFFFFF&
      Height          =   9345
      Left            =   2820
      ScaleHeight     =   9285
      ScaleWidth      =   12240
      TabIndex        =   6
      Top             =   90
      Width           =   12300
      Begin RichTextLib.RichTextBox txtRTF 
         Height          =   975
         Left            =   3330
         TabIndex        =   15
         Top             =   5310
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1720
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"Viewer.frx":4C6E
      End
      Begin ImgeditLibCtl.ImgEdit ImgViewer 
         Height          =   2895
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   4035
         _Version        =   131074
         _ExtentX        =   7117
         _ExtentY        =   5106
         _StockProps     =   96
         ImageControl    =   "ImgEdit"
         BorderStyle     =   0
         UndoBufferSize  =   805306366
         OcrZoneVisibility=   -3680
         AnnotationOcrType=   25801
         ForceFileLinking1x=   -1  'True
         MagnifierZoom   =   25801
         sReserved1      =   -3680
         sReserved2      =   -3680
         lReserved1      =   1902397682
         lReserved2      =   1902397682
         bReserved1      =   -1  'True
         bReserved2      =   -1  'True
      End
      Begin VB.Image imgView 
         Height          =   765
         Left            =   120
         Top             =   6450
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame fraNavigator 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Navigator"
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
      Height          =   9720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2550
      Begin VB.DriveListBox drvNavigator 
         Height          =   315
         Left            =   165
         TabIndex        =   4
         Top             =   345
         Width           =   2220
      End
      Begin VB.DirListBox dirNavigator 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   120
         TabIndex        =   3
         Top             =   750
         Width           =   2280
      End
      Begin VB.FileListBox filNavigator 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   120
         TabIndex        =   2
         Top             =   2700
         Width           =   2310
      End
      Begin ThumbnailLibCtl.ImgThumbnail ImgThumbnail 
         Height          =   4155
         Left            =   120
         TabIndex        =   1
         Tag             =   "7"
         Top             =   5490
         Width           =   2265
         _Version        =   65536
         _ExtentX        =   3995
         _ExtentY        =   7329
         _StockProps     =   97
         BackColor       =   16777215
         BorderStyle     =   1
         BackColor       =   16777215
         ThumbBackColor  =   14737632
         ThumbWidth      =   60
         ThumbHeight     =   82
         HighlightColor  =   12582912
      End
      Begin VB.Label lblTIF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIF Pages"
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
         Left            =   90
         TabIndex        =   5
         Top             =   5250
         Width           =   915
      End
   End
   Begin VB.Label lblZoom 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   14610
      TabIndex        =   13
      Top             =   10230
      Width           =   495
   End
   Begin VB.Label lblLinearScale 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Linear Scaling"
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
      Height          =   180
      Left            =   9570
      TabIndex        =   12
      Top             =   10320
      Width           =   1410
   End
   Begin VB.Label lblFilePath 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   30
      TabIndex        =   9
      Top             =   9780
      Width           =   6165
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   9705
      Left            =   2550
      Top             =   0
      Width           =   12645
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fsFilePath As String
Private fsSavePath As String
Private fsPicFilePath As String
Private fiHorizontal As Long
Private fiVertical As Long
Private fbByPass As Boolean
Private fbActualSizeByPass As Boolean

Private fiActualImageWidth As Long
Private fiActualImageHeight As Long
Private piPatientID As Long
Private fiHorizontalMax As Long
Private fiVerticalMax As Long
Private fbooByPass As Boolean

Private fbooTIF As Boolean
Private fbooTEXT As Boolean
Private fbooIMG As Boolean
Public Sub DisplayImage(vstrFile As String)
Dim lstrRemarks As String

On Error GoTo Error_DisplayImage

    lblFilePath.Caption = vstrFile
    
    lstrRemarks = "Image File Path (vstrFile) = " & vstrFile
    
    
    If InStr(UCase(vstrFile), ".TIF") > 0 Then
        fbooByPass = True
        fraNavigator.Visible = True
        chkHideTIF_list.Visible = True
        ImgThumbnail.Image = lblFilePath.Caption
        picViewer.Width = 11340
        fbooByPass = False
    Else
        ImgThumbnail.Image = ""
        fraNavigator.Visible = False
        chkHideTIF_list.Visible = False
        picViewer.Width = 13700
    End If
    
    ImgViewer.Width = picViewer.Width - 50
    ImgViewer.Height = picViewer.Height - 50
    
    ImgViewer.Image = lblFilePath.Caption
    ImgViewer.ClearDisplay
    

    If chkActualSize.Value Then
       ImgViewer.Zoom = 100
    Else
        If HVScale.Value > 1 Then ImgViewer.Zoom = HVScale.Value
    End If
    
    ImgViewer.display
    
    Exit Sub
    
Error_DisplayImage:
    ErrHandler _
     Me.Name & " Form_Activate", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub


'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal sName$, ByVal Key$, ByVal Def$, ByVal ret$, ByVal Size%, ByVal IniName$) As Long

'Public Function Get_FileExtension(ByVal FileName$) As String
'
''Example:
'' sFile = "system.ini"
'' pFile = Get_FilePrefix(sFile)
'' pFile >> ini
'
'Dim p%
'
' p = InStr(FileName, ".")
' If p > 0 Then
'  Get_FileExtension = Mid$(FileName, p + 1)
' Else
'  Get_FileExtension = FileName
' End If
'
'End Function


'Public Sub IniFile_Write(sSection$, sVar$, sNewVal$, sIniFile$)
''Writes a string value to a variable in an ini file.
'
''Sample Code:
''   IniFile_Write "General", "LogDate", Now, sIniFile
'
'Dim i%
'
'    If Dir$(sIniFile) <> "" Then
'     i = WritePrivateProfileString(sSection, sVar, sNewVal, sIniFile)
'    Else
'     MsgBox "Missing ini File", vbExclamation, "Error"
'    End If
'
'End Sub
'
'Public Function IniFile_Read(sSection$, sVar$, sIniFile$) As String
''Captures a string value assigned to an ini file variable
'
''Sample Code:
''  sIniFile = App.Path & "\" & App.EXEName & ".ini"
''  sLogDate = IniFile_Read("General", "LogDate", sIniFile)
'
'Dim sRet$
'Dim iChars%
'
'    If Dir$(sIniFile) <> "" Then
'
'        sRet = String$(500, 0)
'        iChars = GetPrivateProfileString(sSection, sVar, "", sRet, Len(sRet), sIniFile)
'
'        If iChars > 0 Then
'            IniFile_Read = Left$(sRet, iChars)
'        Else
'            IniFile_Read = ""
'        End If
'    Else
'        MsgBox "Missing File", vbExclamation, "Error"
'    End If
'
'End Function
'
'
'Public Function Get_FilePath(ByVal File$) As String
'Dim p%, Key$
'
''Example:
'' sFile = "c:\windows\system\grid.vbx"
'' sPath = Get_FilePath(sFile)
'' sPath >> c:\windows\system
'
'    p = 1
'    Do Until p = 0
'     p = InStr(File, "\")
'     Key = Key & Left$(File, p)
'     File = Mid$(File, p + 1)
'    Loop
'
'    Get_FilePath = Left$(Key, Len(Key) - 1)
'
'End Function
'
'
'Public Function Get_FileName(ByVal FilePath$) As String
'Dim p%
'
''Example:
'' sPath = "c:\windows\system.ini"
'' sFile = Get_FileName(sPath)
'' sFile >> system.ini
'
'p = 1
'Do Until p = 0
' p = InStr(FilePath, "\")
' FilePath = Mid$(FilePath, p + 1)
'Loop
'
'Get_FileName = FilePath
'
'End Function
'Public Function FileExists(sPath As String) As Boolean
'Dim sFile As String
'
'    sFile = Get_FileName(sPath)
'    If UCase$(Dir$(sPath)) = UCase$(sFile) Then FileExists = True
'
'End Function

Private Sub ClearAllThumbs()
Dim i As Integer

    For i = 1 To ImgViewer.PageCount
        ImgThumbnail.ClearThumbs i
    Next i
    
    ImgThumbnail.Refresh
    
End Sub



Private Sub chkActualSize_Click()
 
    If fbByPass Then Exit Sub
  
    If chkActualSize.Value Then
    
        If ImgViewer.Visible Then
            HVScale.Value = HVScale.Max / 2
            lblZoom.Caption = HVScale.Value
    
            ImgViewer.ClearDisplay
            ImgViewer.Image = lblFilePath.Caption
            ImgViewer.display
            ImgViewer.Zoom = 100
        Else
            imgView.Stretch = False
            'imgView.Width = fiActualImageWidth
            'imgView.Height = fiActualImageHeight
            fiActualImageHeight = imgView.Height
            fiActualImageWidth = imgView.Width
        End If
       
    End If
    
    
End Sub

Private Sub chkHideTIF_list_Click()
    If chkHideTIF_list.Value Then
        fraNavigator.Visible = False
        picViewer.Width = Me.Width - 100
        picViewer.Left = 0
        Shape.Visible = False
    Else
        fraNavigator.Visible = True
        picViewer.Width = Shape.Width - 90 '11340
        picViewer.Left = fraNavigator.Left + fraNavigator.Width + 50
        Shape.Visible = True
    End If
    
    ImgViewer.Width = picViewer.Width - 50
    ImgViewer.Height = picViewer.Height - 50
End Sub

Private Sub cmd25_Click()
    ImgViewer.Zoom = 25
    HVScale.Value = 25
    lblZoom.Caption = 25
    chkActualSize.Value = 0
    picViewer.SetFocus
    
    '---------------------------
    imgView.Stretch = True
    imgView.Width = fiActualImageWidth * 0.25
    imgView.Height = fiActualImageHeight * 0.25
End Sub

Private Sub cmd50_Click()
    ImgViewer.Zoom = 50
    HVScale.Value = 50
    lblZoom.Caption = 50
    chkActualSize.Value = 0
    picViewer.SetFocus
    
    '---------------------------
    imgView.Stretch = True
    imgView.Width = fiActualImageWidth * 0.5
    imgView.Height = fiActualImageHeight * 0.5
    
End Sub


Private Sub cmd75_Click()
    ImgViewer.Zoom = 75
    HVScale.Value = 75
    lblZoom.Caption = 75
    chkActualSize.Value = 0
    picViewer.SetFocus
    
    '---------------------------
    imgView.Stretch = True
    imgView.Width = fiActualImageWidth * 0.75
    imgView.Height = fiActualImageHeight * 0.75
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopyTo_Click()
Dim iFileNum As Integer
Dim iAnswer As Byte

On Error GoTo Open_Error


If fsFilePath = "" Then fsFilePath = App.Path

CommonDialog.CancelError = True
CommonDialog.DialogTitle = "Copy To"
CommonDialog.Filter = "All (*.*)|*.*"
CommonDialog.InitDir = fsFilePath
CommonDialog.FileName = filNavigator.List(filNavigator.ListIndex)

If Trim(CommonDialog.FileName) = vbNullString Then
    MsgBox "Select an image file first", vbExclamation
    Exit Sub
End If


TRY_AGAIN:
  
    CommonDialog.ShowSave
    fsFilePath = Get_FilePath(CommonDialog.FileName)


        If Dir$(CommonDialog.FileName) = CommonDialog.FileTitle Then
            iAnswer = MsgBox("File already exists! Overwrite?", vbExclamation + vbYesNo, "Save As...")
            If iAnswer = vbNo Then GoTo TRY_AGAIN
        End If
     
    FileCopy lblFilePath.Caption, CommonDialog.FileName
Exit Sub

Open_Error:
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim iResp As Byte
Dim iPos As Integer
    
    iResp = MsgBox("Are you sure that you want to delete photo?", vbExclamation + vbYesNo)
    
    If iResp = vbYes Then
        iPos = filNavigator.ListIndex
        'ImgViewer.Picture = LoadPicture()
        ImgViewer.ClearDisplay
        Kill lblFilePath
        filNavigator.Refresh
        filNavigator.ListIndex = iPos
    End If
    
End Sub


Private Sub cmdPrint_Click()
Dim i As Integer
Dim intPages As Integer


    On Error GoTo Print_Error
    'CommonDialog.CancelError = True
    'CommonDialog.flags = 0
    'CommonDialog.ShowPrinter
    ShowPrinter Me.Hwnd
    
    'If InStr(UCase(filNavigator.FileName), ".TIF") > 0 Then
    If fbooTIF Then
        ImgThumbnail.SelectAllThumbs
        intPages = ImgThumbnail.SelectedThumbCount
        ImgViewer.PrintImage 1, intPages, 2
        ImgThumbnail.DeselectAllThumbs
        
'        For i = 1 To intPages
'            ImgThumbnail.ThumbSelected(i) = True
'            ImgViewer.Page = i
'            ImgViewer.Display
'            ImgViewer.PrintImage 1, 1, 2
'            ImgThumbnail.ThumbSelected(i) = False
'        Next i
   
    ElseIf fbooIMG Then
        Printer.PaintPicture imgView.Picture, 0, 0  ', imgView.Width, imgView.Height
        Printer.EndDoc
    ElseIf fbooTEXT Then
        txtRTF.SelPrint Printer.hdc
    End If
    
    Exit Sub

Print_Error:
    If Err.Number = 482 Then MsgBox Err.Description, vbCritical
    Exit Sub
End Sub


Private Sub cmdRename_Click()
Dim sFile As String
Dim sFileName As String
Dim sNewFile As String
Dim sPath As String
Dim i As Integer

On Error GoTo Rename_Error

    sFileName = filNavigator.List(filNavigator.ListIndex)
    sNewFile = InputBox("New File Name:", "Rename FIle", sFileName)
    sPath = filNavigator.Path & "\" & sNewFile
    
    
    If Trim(sNewFile) = vbNullString Then
        MsgBox "Enter a new file name.", vbExclamation
    Else
        sFile = filNavigator.Path & "\" & sFileName
        Name sFile As sPath
        filNavigator.Refresh
        
        For i = 0 To filNavigator.listcount - 1
        
            If filNavigator.List(i) = sNewFile Then
                fbByPass = True
                filNavigator.ListIndex = i
                fbByPass = False
                Exit For
            End If
        
        Next
        
    End If
    
Exit Sub

Rename_Error:
    MsgBox Err.Description, vbExclamation
    Exit Sub
End Sub

Private Sub cmdRotate_Click()
    ImgViewer.RotateRight
End Sub

Private Sub cmdSavePic_Click()
Dim sFileName As String
Dim sFilter As String
Dim i As Integer
Dim sFile As String
Dim intResp As Integer
Dim iFyleType As Integer

    sFile = filNavigator.FileName
    sFilter = "*." & Get_FileExtension(filNavigator.FileName)
    fsSavePath = dirNavigator.Path
    sFileName = CommonDialogAction(CommonDialog, "Save", "Save as....", fsSavePath, sFilter)
    fsSavePath = Get_FilePath(sFileName)
       
    If sFileName <> vbNullString Then
        If fbooTIF Then
        
            Select Case UCase(Right(sFileName, 3))
            
                Case "GIF"
                    iFyleType = 8
                
                Case "BMP"
                    iFyleType = 3
                
                Case "TIF"
                    iFyleType = 1
                
                Case "JPG"
                    iFyleType = 6
                
            End Select
            
            ImgViewer.SaveAs sFileName, iFyleType, , , , True
            fsSavePath = Get_FilePath(sFileName)
            chkActualSize.Value = 1
        

        ElseIf fbooIMG Then
            If InStr(UCase(sFileName), ".BMP") > 0 Then
                SavePicture imgView, sFileName
            Else
                intResp = MsgBox("Save as bitmap file?", vbQuestion + vbYesNo, "Save Image")
                If intResp = vbYes Then
                    'Strip file extension as replace with .bmp.........
                    If InStr(sFileName, ".") > 0 Then
                        sFileName = Left(sFileName, Len(sFileName) - 3)
                        sFileName = sFileName & "bmp"
                        SavePicture imgView, sFileName
                    End If
                End If
            End If
        
        ElseIf fbooTEXT Then
            If InStr(UCase(sFileName), ".TXT") > 0 Then
                txtRTF.SaveFile sFileName, 1
            ElseIf InStr(UCase(sFileName), ".RTF") > 0 Then
                txtRTF.SaveFile sFileName, 0
            End If
        End If
        
    End If
    
    filNavigator.Refresh
    
    If sFileName <> vbNullString Then sFile = Get_FileName(sFileName) Else sFile = Get_FileName(sFile)
        
    
        For i = 0 To filNavigator.listcount - 1
        
            If filNavigator.List(i) = sFile Then
                fbByPass = True
                filNavigator.ListIndex = i
                fbByPass = False
                Exit For
            End If
        
        Next
End Sub


Public Function CommonDialogAction(CommonDialog As Control, sAction As String, sMsg As String, sLastFilePath As String, Optional vFilter, Optional vFont) As Variant

'Sample Code:
'    If Not FileExists(sFile) Then
'        sFile = Get_ActiveX_File(fsActXFile(lstActiveX.ListIndex))
'    End If


Dim iAnswer As Byte

On Error GoTo CommonDialogAction_Error

If sLastFilePath = "" Then sLastFilePath = App.Path

Select Case sAction

    Case "ShowColor"
        CommonDialog.ShowColor
        CommonDialogAction = CommonDialog.Color
    
    Case "ShowFont"
        CommonDialog.Flags = cdlCFBoth
        CommonDialog.ShowFont
        
        If Not IsMissing(vFont) Then
            ReDim vFont(3)
            vFont(0) = CommonDialog.FontName
            vFont(1) = CommonDialog.FontSize
            vFont(2) = CommonDialog.FontBold
            vFont(3) = CommonDialog.FontItalic
        End If
    
    Case "ShowHelp"
        CommonDialog.HelpFile = sLastFilePath '"C:\VB5\Help\Vb5.hlp"
        CommonDialog.HelpCommand = &H3
        CommonDialog.showHelp
        
    Case "ShowPrinter"
        CommonDialog.CancelError = True
        CommonDialog.Flags = 0
        CommonDialog.ShowPrinter
        
        'Printer.FontSize = 10
        'Printer.Print sMsg
        'Printer.Print
        'Printer.EndDoc
        
        CommonDialogAction = Printer
        
    Case "Open"
        CommonDialog.CancelError = True
        CommonDialog.DialogTitle = "Open " & sMsg
        CommonDialog.Filter = "(" & vFilter & ")|" & vFilter & "|All (*.*)|*.*" '"OCX (*.ocx)|*.ocx|All (*.*)|*.*"
        CommonDialog.InitDir = sLastFilePath
        CommonDialog.FileName = vFilter  '"*.ocx"
        CommonDialog.ShowOpen
        sLastFilePath = Get_FilePath(CommonDialog.FileName)

        If sLastFilePath = "" Then
            MsgBox "File not found!", vbExclamation
            Exit Function
        End If
        
        CommonDialogAction = CommonDialog.FileName
    
    Case "Save"
        CommonDialog.CancelError = True
        CommonDialog.DialogTitle = "Save As " & sMsg
        CommonDialog.Filter = "(" & vFilter & ")|" & vFilter & "|All (*.*)|*.*"
        CommonDialog.InitDir = sLastFilePath
        CommonDialog.FileName = vFilter
        
TRY_AGAIN:
        CommonDialog.ShowSave
        sLastFilePath = Get_FilePath(CommonDialog.FileName)
    
            If Dir$(CommonDialog.FileName) = CommonDialog.FileTitle Then
                iAnswer = MsgBox("File already exists! Overwrite?", vbExclamation + vbYesNo, "Save As...")
                If iAnswer = vbNo Then GoTo TRY_AGAIN
            End If
        
        CommonDialogAction = CommonDialog.FileName
        
    Case "Update"
        CommonDialog.CancelError = True
        CommonDialog.DialogTitle = "Save " & sMsg
        CommonDialog.Filter = "(" & vFilter & ")|" & vFilter & "|All (*.*)|*.*"
        CommonDialog.InitDir = sLastFilePath
        CommonDialog.FileName = sMsg

TRY_AGAIN2:
        CommonDialog.ShowSave
        sLastFilePath = Get_FilePath(CommonDialog.FileName)
    
            If Dir$(CommonDialog.FileName) = CommonDialog.FileTitle Then
                iAnswer = MsgBox("File already exists! Overwrite?", vbExclamation + vbYesNo, "Save As...")
                If iAnswer = vbNo Then GoTo TRY_AGAIN2
            End If
        
        CommonDialogAction = CommonDialog.FileName
    
End Select



Exit Function

CommonDialogAction_Error:
    MsgBox Err.Description
    Exit Function
End Function


Private Sub cmdUpdatePic_Click()
Dim sFileName As String
Dim sFilter As String
Dim i As Integer
Dim sFile As String
Dim sPath As String
Dim iFyleType As Integer
Dim intResp As Integer

    sFile = filNavigator.FileName
    'sFilter = "*." & Get_FileExtension(filNavigator.FileName)
    sPath = filNavigator.Path
    'sFileName = CommonDialogAction(CommonDialog, "Update", sFile, sPath, sFilter)
    
    sFileName = sPath & "\" & sFile
       
    If sFileName <> vbNullString Then
        
        If fbooTIF Then
        
            Select Case UCase(Right(sFileName, 3))
            
                Case "GIF"
                    iFyleType = 8
                
                Case "BMP"
                    iFyleType = 3
                
                Case "TIF"
                    iFyleType = 1
                
                Case "JPG"
                    iFyleType = 6
                
            End Select
            
            ImgViewer.SaveAs sFileName, iFyleType, , , , True
            chkActualSize.Value = 1
            
        ElseIf fbooIMG Then
            If InStr(UCase(sFileName), ".BMP") > 0 Then
                SavePicture imgView, sFileName
            Else
                intResp = MsgBox("Save as bitmap file?", vbQuestion + vbYesNo, "Save Image")
                If intResp = vbYes Then
                    'Strip file extension as replace with .bmp.........
                    If InStr(sFileName, ".") > 0 Then
                        sFileName = Left(sFileName, Len(sFileName) - 3)
                        sFileName = sFileName & "bmp"
                        SavePicture imgView, sFileName
                    End If
                End If
            End If
        
        ElseIf fbooTEXT Then
            If InStr(UCase(sFileName), ".TXT") > 0 Or InStr(UCase(sFileName), ".SQL") > 0 Then
                txtRTF.SaveFile sFileName, 1
            ElseIf InStr(UCase(sFileName), ".RTF") > 0 Then
                txtRTF.SaveFile sFileName, 0
            End If
        End If
        
    End If

    
    filNavigator.Refresh
    
    For i = 0 To filNavigator.listcount - 1
    
        If filNavigator.List(i) = sFile Then
            fbByPass = True
            filNavigator.ListIndex = i
            fbByPass = False
            Exit For
        End If
    
    Next
        
End Sub

Private Sub dirNavigator_Change()
    filNavigator.Path = dirNavigator.Path
    cmdSavePic.Enabled = False
End Sub

Private Sub drvNavigator_Change()
Dim strPath As String

    On Error GoTo Error
    strPath = dirNavigator.Path
    dirNavigator.Path = drvNavigator.Drive
    Exit Sub
    
Error:
    MsgBox Err.Description & " " & Err.Number
    dirNavigator.Path = strPath
    Exit Sub
    Resume
    
End Sub

Private Sub filNavigator_Click()
Dim intTIFpages As Integer
Dim i As Integer


On Error GoTo picError

If fbByPass Then Exit Sub

Screen.MousePointer = vbHourglass

cmdSavePic.Enabled = True

'VSlide.Value = 0
'HSlide.Value = 0

If Right$(filNavigator.Path, 1) <> "\" Then
    lblFilePath.Caption = filNavigator.Path & "\" & filNavigator.FileName
Else
    lblFilePath.Caption = filNavigator.Path & filNavigator.FileName
End If

'If chkActualSize.Value Then ImgViewer.Stretch = False

txtRTF.Visible = False

If InStr(UCase(filNavigator.FileName), ".TIF") > 0 Then

    fbooTIF = True
    fbooTEXT = False
    fbooIMG = False
    
    fbByPass = True
    ImgViewer.Visible = True
    imgView.Visible = False
    txtRTF.Visible = False
    
    cmdRotate.Enabled = True
    intTIFpages = ImgViewer.PageCount
    ImgThumbnail.Image = lblFilePath.Caption
    fbByPass = False
    
    ImgViewer.Image = lblFilePath.Caption
    ImgViewer.ClearDisplay
    
    
    If chkActualSize.Value Then
       ImgViewer.Zoom = 100
       ImgViewer.Stretch = False
    Else
        If HVScale.Value > 1 Then ImgViewer.Zoom = HVScale.Value
        ImgViewer.Stretch = True
    End If
    ImgViewer.display
ElseIf InStr(UCase(filNavigator.FileName), ".RTF") > 0 Or InStr(UCase(filNavigator.FileName), ".TXT") > 0 Or InStr(UCase(filNavigator.FileName), ".SQL") > 0 Then
    txtRTF.Visible = True
    ImgViewer.Visible = False
    imgView.Visible = False
    
    If InStr(UCase(lblFilePath.Caption), ".TXT") > 0 Then
        txtRTF.LoadFile lblFilePath.Caption, 1
    Else
        txtRTF.LoadFile lblFilePath.Caption
    End If
    fbooTIF = False
    fbooTEXT = True
    fbooIMG = False
Else
    ImgThumbnail.Image = ""
    ImgViewer.Visible = False
    cmdRotate.Enabled = False
    imgView.Visible = True
    txtRTF.Visible = False
    If chkActualSize.Value Then imgView.Stretch = False Else imgView.Stretch = True
    imgView.Picture = LoadPicture(lblFilePath.Caption)
    fiActualImageHeight = imgView.Height
    fiActualImageWidth = imgView.Width
    fbooTIF = False
    fbooTEXT = False
    fbooIMG = True
End If


  

'    Debug.Print ImgViewer.ImageScaleHeight
'    Debug.Print ImgViewer.ImagePalette
'    Debug.Print ImgViewer.FileType
'    Debug.Print ImgViewer.CompressionType
'    Debug.Print ImgViewer.DisplayICMEnabled
'    Debug.Print ImgViewer.ImageResolutionX
'    Debug.Print ImgViewer.ImageResolutionY
'    Debug.Print ImgViewer.Height
'    Debug.Print ImgViewer.Width
'    Debug.Print ImgViewer.Zoom
    
    fbByPass = False
    Screen.MousePointer = vbDefault

Exit Sub

picError:
    If Err.Number = 438 Or Err.Number = 1002 Then
        If Not fbooTIF Then MsgBox Err.Description, vbExclamation, "Error # " & Err.Number
        Resume Next
    Else
        ImgViewer.ClearDisplay
        imgView.Picture = LoadPicture()
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume Next
End Sub

Private Sub Form_Activate()
Dim iFileNum As Integer
Dim sIniFile As String
Dim sContent As String
Dim sLastPath As String
Dim lstrRemarks As String
Dim lbooImagePathExists As Boolean
Dim sImagePath As String

On Error GoTo Error_Viewer

fbByPass = True
Me.Width = Screen.Width
Me.Height = Screen.Height - 400
Me.Top = 0
Me.Left = 0

lbooImagePathExists = True

filNavigator.Pattern = "*.bmp;*.wmf;*.ico;*.gif;*.jpg;*.tif;*.mbs;*.thm;*.rtf;*.txt;*.sql"
If sLastPath = vbNullString Then
    If psImageFilePath = vbNullString Then
      psImageFilePath = "C:\"
    End If
    drvNavigator.Drive = psImageFilePath
    
    If lbooImagePathExists Then
       If InStr(psImageFilePath, "\") = 0 Then
            sImagePath = psImageFilePath & "\"
       Else
            sImagePath = psImageFilePath
       End If
       
    If piPatientID <= 0 Then
        piPatientID = modSpecific.piPatientID
    Else
        piPatientID = frmPatients.piPatientID
    End If
        If FolderExists(sImagePath & piPatientID) Then
            dirNavigator.Path = sImagePath & piPatientID
        Else
            dirNavigator.Path = sImagePath
        End If
    
    End If
Else
    drvNavigator.Drive = sLastPath
    dirNavigator.Path = sLastPath
End If


ImgViewer.Width = picViewer.Width
ImgViewer.Height = picViewer.Height

cmdSavePic.Enabled = False

  HVScale.Value = HVScale.Max / 2
  lblZoom.Caption = HVScale.Value
  ImgViewer.AutoRefresh = True
    
    ImgViewer.ImagePalette = wiPaletteRGB24
    
    imgView.Left = ImgViewer.Left
    imgView.Top = ImgViewer.Top
    imgView.Width = ImgViewer.Width
    imgView.Height = ImgViewer.Height
    
    txtRTF.Left = ImgViewer.Left
    txtRTF.Top = ImgViewer.Top
    txtRTF.Width = ImgViewer.Width
    txtRTF.Height = ImgViewer.Height
    
    scrVert.Min = 0
    scrVert.Max = ImgViewer.Height
    scrVert.LargeChange = 200
    scrVert.SmallChange = 50
    
    scrHoriz.Min = 0
    scrHoriz.Max = ImgViewer.Width
    scrHoriz.LargeChange = 200
    scrHoriz.SmallChange = 50
  
fbByPass = False

Exit Sub

Error_Viewer:

    If Err.Number = 68 Then
        lstrRemarks = "Image Path set in configuration is not a valid path."
        MsgBox lstrRemarks, vbExclamation
        lbooImagePathExists = False
    End If
   ErrHandler _
     Me.Name & " Form_Activate", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume Next
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    If IsFormLoaded("frmInPatient") Then
        frmInPatient.ZOrder 0
    End If

    Unload Me
    Exit Sub
End Sub

Private Sub HVScale_Change()
Dim lsngWidth As Single
Dim lsngHeight As Single

    If fbByPass Then Exit Sub
    
    If HVScale.Value > 1 Then
        lsngWidth = HVScale.Value / 100
        lsngHeight = HVScale.Value / 100
    End If

    
    If HVScale.Value > 1 Then ImgViewer.Zoom = HVScale.Value: lblZoom.Caption = HVScale.Value
    
    If HVScale.Value > 1 Then imgView.Width = fiActualImageWidth * lsngWidth: imgView.Height = fiActualImageHeight * lsngHeight
    
End Sub

Private Sub HVScale_GotFocus()

        fbByPass = True
        chkActualSize.Value = 0
        fbByPass = False
        imgView.Stretch = True

End Sub


Private Sub HVScale_Scroll()
Dim lsngWidth As Single
Dim lsngHeight As Single

    If fbByPass Then Exit Sub

    If HVScale.Value > 1 Then
        lsngWidth = HVScale.Value / 100
        lsngHeight = HVScale.Value / 100
    End If

    If HVScale.Value > 1 Then ImgViewer.Zoom = HVScale.Value: lblZoom.Caption = HVScale.Value

    If HVScale.Value > 1 Then imgView.Width = fiActualImageWidth * lsngWidth: imgView.Height = fiActualImageHeight * lsngHeight
End Sub

Private Sub ImgThumbnail_Click(ByVal ThumbNumber As Long)
    'On Error Resume Next
    ImgViewer.Page = ThumbNumber
    If ImgViewer.Image <> "" Then
        ImgViewer.display
    End If
End Sub













Private Sub scrHoriz_Change()
    If fbooTIF Then
        ImgViewer.Left = -scrHoriz.Value
    ElseIf fbooTEXT Then
        txtRTF.Left = -scrHoriz.Value
    ElseIf fbooIMG Then
        imgView.Left = -scrHoriz.Value
    End If
End Sub

Private Sub scrHoriz_Scroll()
    If fbooTIF Then
        ImgViewer.Left = -scrHoriz.Value
    ElseIf fbooTEXT Then
        txtRTF.Left = -scrHoriz.Value
    ElseIf fbooIMG Then
        imgView.Left = -scrHoriz.Value
    End If
End Sub


Private Sub scrVert_Change()
    If fbooTIF Then
        ImgViewer.Top = -scrVert.Value
    ElseIf fbooTEXT Then
        txtRTF.Top = -scrVert.Value
    ElseIf fbooIMG Then
        imgView.Top = -scrVert.Value
    End If
End Sub


Private Sub scrVert_Scroll()
    If fbooTIF Then
        ImgViewer.Top = -scrVert.Value
    ElseIf fbooTEXT Then
        txtRTF.Top = -scrVert.Value
    ElseIf fbooIMG Then
        imgView.Top = -scrVert.Value
    End If
End Sub


