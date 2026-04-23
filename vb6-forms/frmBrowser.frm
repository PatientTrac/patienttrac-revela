VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PatientTrac CCD Reports"
   ClientHeight    =   10485
   ClientLeft      =   3915
   ClientTop       =   1995
   ClientWidth     =   12000
   ForeColor       =   &H00400000&
   Icon            =   "frmBrowser.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   12000
   Begin VB.Frame fraFiles 
      Height          =   4935
      Left            =   1320
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
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
         Height          =   1770
         Left            =   240
         Pattern         =   "*.xlm;*.xms"
         TabIndex        =   2
         Top             =   2595
         Width           =   3990
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
         Left            =   240
         TabIndex        =   1
         Top             =   645
         Width           =   3960
      End
      Begin VB.DriveListBox drvNavigator 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2700
      End
      Begin PT_XP_Button.PT_XP cmdCancel 
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   4485
         Width           =   1065
         _ExtentX        =   1879
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
         Picture         =   "frmBrowser.frx":0442
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBrowser.frx":0E54
         ShowFocus       =   -1  'True
      End
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   10455
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11970
      ExtentX         =   21114
      ExtentY         =   18441
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ComboBox cboAddress 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2085
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Text            =   "cboAddress"
      Top             =   -750
      Width           =   6375
   End
   Begin VB.Frame fraFav 
      BackColor       =   &H00E0E0E0&
      Height          =   5565
      Left            =   -4080
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   3735
      Begin VB.ListBox lstFav 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4890
         Left            =   60
         TabIndex        =   5
         Top             =   570
         Width           =   3525
      End
      Begin PT_XP_Button.PT_XP cmdTerms 
         Height          =   345
         Left            =   2820
         TabIndex        =   6
         ToolTipText     =   "Manage Favorites"
         Top             =   180
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmBrowser.frx":11A6
         BackColorDown   =   15133676
         TransparentColor=   14215660
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmBrowser.frx":1740
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   345
         Left            =   3270
         TabIndex        =   7
         Top             =   180
         Width           =   315
         _ExtentX        =   556
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
         Picture         =   "frmBrowser.frx":1A92
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmBrowser.frx":24A4
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6180
      Top             =   1500
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":27F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":309C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":337E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":3660
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin PT_XP_Button.PT_XP cmdGo 
      Height          =   345
      Left            =   8520
      TabIndex        =   8
      Top             =   -750
      Width           =   645
      _ExtentX        =   1138
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
      Picture         =   "frmBrowser.frx":3942
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Go!"
      CaptionAlignment=   3
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmBrowser.frx":3EDC
      PictureAlignment=   5
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdHome 
      Height          =   345
      Left            =   9270
      TabIndex        =   10
      Top             =   -540
      Width           =   885
      _ExtentX        =   1561
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
      Picture         =   "frmBrowser.frx":422E
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Home"
      CaptionAlignment=   3
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmBrowser.frx":47C8
      PictureAlignment=   5
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdFavorite 
      Height          =   345
      Left            =   750
      TabIndex        =   11
      Top             =   -390
      Width           =   1275
      _ExtentX        =   2249
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
      Picture         =   "frmBrowser.frx":4B1A
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Favorites"
      CaptionAlignment=   3
      CaptionPosition =   0
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmBrowser.frx":50B4
      PictureAlignment=   5
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdNext 
      Height          =   345
      Left            =   360
      TabIndex        =   12
      Top             =   -450
      Width           =   315
      _ExtentX        =   556
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
      Picture         =   "frmBrowser.frx":5406
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   "Next"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmBrowser.frx":57A0
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPrev 
      Height          =   345
      Left            =   0
      TabIndex        =   13
      Top             =   -450
      Width           =   315
      _ExtentX        =   556
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
      Picture         =   "frmBrowser.frx":5AF2
      AppearanceThemes=   3
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmBrowser.frx":5E8C
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin VB.Menu mnu_Patient 
      Caption         =   "Select Patient"
   End
   Begin VB.Menu mnu_View 
      Caption         =   "View CCD Report"
   End
   Begin VB.Menu mnu_CCR 
      Caption         =   "Get CCR"
   End
   Begin VB.Menu mnu_Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" _
    (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Private mbooByPass As Boolean
Private lngPatient As Long

 
Private Sub MyNOtes()
Dim FileRead As New C32
'string filepathXML;
'string filePathXSL;
'string filePathOut;
'
'/* ----- CCD with CCD_Rev3.xsl ----- */
'filepathXML = @"C:\UnitTest\GenerateHumanReadable\CCD\Ann-Toliver_2011_08_10_16_55_57.xml";
'filePathXSL = @"C:\UnitTest\GenerateHumanReadable\StyleSheets\CCD_Rev3.xsl";
'filePathOut = @"C:\UnitTest\GenerateHumanReadable\html\Ann-Toliver_CCD_Rev3.xsl.html";

'string docType = c32.DetermineMessageType(filepathXML);
'
'// Perform logic to choose the correct XSL file here
'
'result = c32.GenerateHumanReadable(filepathXML, filePathXSL, filePathOut);
'
'
   
   
' FileRead.DetermineMessageType (filepathXML)
 
'The file "DHIT COM Helper Release Notes.docx" details the commands you will need to execute to generate the various message types. They are typically single line commands. For example to generate a CCD issue the following command:
'
'int GenerateC32SP(string connectionString, string storedProcedureName, string parameterName, string parameterValue, string fileOutPath);
'
'Though displaying a CCD requires two commands - one to determine the type and another to generate the html:
'
'Generating Human Readable HTML from CCD or CCR XML:
'
'CCD and CCR require different stylesheets to perform the transform. So DetermineMessageType is used to determine the message type then you will have to locate the appropriate XSL file to use GenerateHumanReadable.
'
'DetermineMessageType returns "CCD", "CCR", or "Unknown"
'
'Example Usage:


'Either of these two would be a great place to start. The integrity commands are slightly
'more involved but the explanation and more so the examples in the release notes should make it clear.
'Generating Human Readable HTML from CCD or CCR XML:
'
'CCD and CCR require different stylesheets to perform the transform. So DetermineMessageType is used to determine the message type then you will have to locate the appropriate XSL file to use GenerateHumanReadable.
'
'DetermineMessageType returns "CCD", "CCR", or "Unknown"
'
'Example Usage:
'
'string filepathXML;
'string filePathXSL;
'string filePathOut;
'
'/* ----- CCD with CCD_Rev3.xsl ----- */
'filepathXML = @"C:\UnitTest\GenerateHumanReadable\CCD\Ann-Toliver_2011_08_10_16_55_57.xml";
'filePathXSL = @"C:\UnitTest\GenerateHumanReadable\StyleSheets\CCD_Rev3.xsl";
'filePathOut = @"C:\UnitTest\GenerateHumanReadable\html\Ann-Toliver_CCD_Rev3.xsl.html";
'
'string docType = c32.DetermineMessageType(filepathXML);
'
'// Perform logic to choose the correct XSL file here
'
'result = c32.GenerateHumanReadable(filepathXML, filePathXSL, filePathOut);


End Sub

Private Sub brwWebBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Dim Idx As Integer

'    cboAddress.Text = brwWebBrowser.LocationURL
    Exit Sub
End Sub

Private Sub cboAddress_Click()
    On Error GoTo Error_Get_Address
    If mbooByPass Then mbooByPass = False: Exit Sub
    
    If mbDontNavigateNow Then Exit Sub
    brwWebBrowser.navigate cboAddress.Text
    Exit Sub
Error_Get_Address:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, cboAddress.Text, "Get_Address", "PatientTracEMR.frmBrowser"
    Exit Sub
    Resume
End Sub


Private Sub cmdClose_Click()
    fraFav.Visible = False
    Exit Sub
End Sub

Private Sub cmdFavorite_Click()
    On Error GoTo Error_Favorite
''    If lstFav.listcount <= 0 Then Load_Terms
'    fraFav.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Favorite:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, cboAddress.Text, "Favorite", "PatientTracEMR.frmBrowser"
    Exit Sub
    Resume
End Sub

Private Sub cmdGo_Click()
   If mbDontNavigateNow Then Exit Sub
   ' timTimer.Enabled = True
    brwWebBrowser.navigate cboAddress.Text
    Exit Sub
End Sub

Private Sub cmdHome_Click()
 Dim cnHome As New ADODB.Connection
Dim rsHome As New ADODB.Recordset
Dim strSQL As String
Dim strAddress As String

    On Error GoTo Error_Get_Home
    Screen.MousePointer = vbHourglass
    
    If cnHome.State = adStateClosed Then cnHome.open psConnect(1)
    strAddress = cboAddress.Text
    strSQL = "SELECT HOME, URL_ID from URL_TERMS "
    Set rsHome = cnHome.Execute(strSQL)
    If Not rsHome.EOF Then
        If rsHome.Fields("HOME") <> vbNullString Then
            cboAddress.Text = rsHome.Fields("HOME")
        End If
    End If
    If cnHome.State = adStateOpen Then
        cnHome.Close
        Set cnHome = Nothing
        Set rsHome = Nothing
    End If
    brwWebBrowser.navigate cboAddress.Text
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Get_Home:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_Home", "PatientTracEMR.frmBrowser"
    Exit Sub
    Resume
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    brwWebBrowser.GoForward
End Sub

Private Sub cmdPrev_Click()
    On Error Resume Next
    brwWebBrowser.GoBack
    Exit Sub
End Sub


Private Sub cmdTerms_Click()
    On Error GoTo Error_Get_Terms
    Screen.MousePointer = vbDefault
'    frmURL.LoadListBox
'    frmURL.Show vbModal
'    Load_Terms
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Get_Terms:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_Terms", "PatientTracEMR.frmBrowser"
    Exit Sub
    Resume
End Sub

Private Sub Form_Activate()
    On Error GoTo Error_Activate
    Form_Resize
    Exit Sub
Error_Activate:
    Exit Sub
    Resume
End Sub
Private Sub Form_Load()
    On Error Resume Next
    drvNavigator.Drive = "C:\"
    dirNavigator.Path = "C:\ABN_Data"
     Exit Sub
End Sub



Private Sub brwWebBrowser_DownloadComplete()
Dim Idx As Integer

    On Error Resume Next
'    Me.Caption = brwWebBrowser.LocationName
'    cboAddress.Text = brwWebBrowser.LocationURL
    Exit Sub
End Sub
 
Private Sub Form_Resize()
'    Me.Width = frmMDI.Width / 1.2      ' 1.5
'   Me.Height = frmMDI.Height - 1550
'    brwWebBrowser.Width = Me.ScaleWidth - 100
'    brwWebBrowser.Height = frmMDI.Height - 100
    Me.Top = 0
    Me.Left = 0
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next
    timTimer.Enabled = False
    mbDontNavigateNow = True
    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    If IsFormLoaded("frmInPatient") Then
        frmInPatient.ZOrder 0
    End If
    Exit Sub
End Sub

Private Sub lstFav_Click()
    If Trim(lstFav.Text) <> vbNullString Then
        fraFav.Visible = False
        cboAddress.Text = Trim(lstFav.Text)
        brwWebBrowser.navigate cboAddress.Text
    End If
End Sub
Private Sub mnu_CCD_Click()




    'brwWebBrowser.navigate
    
End Sub

Private Sub mnu_Exit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub mnu_Patient_Click()
Dim i%

    On Error GoTo Error_Select_Patient
    Screen.MousePointer = vbHourglass
    frmListPatient.mbShowEnctr = False
    lngPatient = 0
    frmListPatient.mlngListItemData = 0
    frmListPatient.Show vbModal
    If frmListPatient.mlngListItemData <> 0 Then
        lngPatient = Val(frmListPatient.mlngListItemData)
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Select_Patient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, udtSchedPatient.Name, "Select_Patient", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub mnu_View_Click()
Dim docType As String

    fraFiles.Visible = True
    dirNavigator.Path = "C:\ABN_Data\" & lngPatient
 
'    Generating Human Readable HTML from CCD or CCR XML:
'
'    CCD and CCR require different stylesheets to perform the transform. So DetermineMessageType is used to determine the message type then you will have to locate the appropriate XSL file to use GenerateHumanReadable.
'
'    DetermineMessageType returns "CCD", "CCR", or "Unknown"
'
'    Example Usage:
'
'    string filepathXML;
'    string filePathXSL;
'    string filePathOut;
'
'    /* ----- CCD with CCD_Rev3.xsl ----- */
'    filepathXML = @"C:\UnitTest\GenerateHumanReadable\CCD\Ann-Toliver_2011_08_10_16_55_57.xml";
'    filePathXSL = @"C:\UnitTest\GenerateHumanReadable\StyleSheets\CCD_Rev3.xsl";
'    filePathOut = @"C:\UnitTest\GenerateHumanReadable\html\Ann-Toliver_CCD_Rev3.xsl.html";
'
'    docType = C32.DetermineMessageType(filepathXML)
'
'    // Perform logic to choose the correct XSL file here
'
'    result = c32.GenerateHumanReadable(filepathXML, filePathXSL, filePathOut);
    
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
'        Me.Caption = brwWebBrowser.LocationName
    Else
        'Me.Caption = "Working..."
    End If
'    brwWebBrowser_NavigateComplete brwWebBrowser.LocationName
End Sub



