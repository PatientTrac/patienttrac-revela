VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{E1A6B8A3-3603-101C-AC6E-040224009C02}#1.0#0"; "ImgThumb.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmScannedImage 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Viewer"
   ClientHeight    =   9750
   ClientLeft      =   4170
   ClientTop       =   2775
   ClientWidth     =   17400
   Icon            =   "frmScannedImage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   17400
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9645
      Left            =   2520
      TabIndex        =   6
      Top             =   60
      Width           =   14775
      ExtentX         =   26061
      ExtentY         =   17013
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
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
      Begin PT_XP_Button.PT_XP cmdRefresh 
         Height          =   525
         Left            =   150
         TabIndex        =   7
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   926
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmScannedImage.frx":0442
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Refresh Patient"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmScannedImage.frx":0D1C
         ShowFocus       =   -1  'True
      End
      Begin VB.DriveListBox drvNavigator 
         Height          =   315
         Left            =   165
         TabIndex        =   4
         Top             =   765
         Width           =   2220
      End
      Begin VB.DirListBox dirNavigator 
         Height          =   1890
         Left            =   120
         TabIndex        =   3
         Top             =   1170
         Width           =   2280
      End
      Begin VB.FileListBox filNavigator 
         Height          =   2430
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   2310
      End
      Begin ThumbnailLibCtl.ImgThumbnail ImgThumbnail 
         Height          =   3825
         Left            =   120
         TabIndex        =   1
         Top             =   5850
         Width           =   2265
         _Version        =   65536
         _ExtentX        =   3995
         _ExtentY        =   6747
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
         Top             =   5610
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmScannedImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fsFilePath As String
Private fsSavePath As String
Private fsPicFilePath As String
Private fstrNetSerPath As String
Private fiHorizontal As Long
Private fiVertical As Long
Private fbByPass As Boolean
Private fbActualSizeByPass As Boolean
Private mFirst As Boolean
Private fiActualImageWidth As Long
Private fiActualImageHeight As Long
Private piPatientID As Long
Private fiHorizontalMax As Long
Private fiVerticalMax As Long
Private fbooByPass As Boolean
Private mstrDomainServer As String
Private fbooTIF As Boolean
Private fbooTEXT As Boolean
Private fbooIMG As Boolean
Private mbDontNavigateNow As Boolean



Function Get_Path()
Dim Arr
Dim strComputer
Dim WSHShell, WSHNetwork, objDomain, DomainString, UserString, UserObj, Path

    On Error GoTo Error_Windows_Path
    
    Set WSHShell = CreateObject("WScript.Shell")
    Set WSHNetwork = CreateObject("WScript.Network")
    'Automatically find the domain name
    Set objDomain = GetObject("LDAP://rootDse")
    DomainString = objDomain.get("dnsHostName")

    'Grab the user name
    UserString = WSHNetwork.UserName
    'Bind to the user object to get user name and check for group memberships later
    Set UserObj = GetObject("WinNT://" & DomainString & "/" & UserString)
    
    'Grab the computer name for use in add-on code later
    strComputer = WSHNetwork.ComputerName
    
    Arr = Split(DomainString, ".")
    If Arr(0) <> vbNullString Then
        mstrDomainServer = Arr(0)
    Else
        Exit Function
    End If
    fstrNetSerPath = mstrDomainServer & "\Scanned_Files\"
    If mstrDomainServer <> vbNullString Then
        WSHNetwork.RemoveNetworkDrive "P:", True, True
        WSHNetwork.MapNetworkDrive "P:", "\\" & Arr(0) & "\Scanned_Files", True
    End If
    Set UserObj = Nothing
    Set WSHNetwork = Nothing
    Set DomainString = Nothing
    Set WSHShell = Nothing
Exit Function
Error_Windows_Path:
    MsgBox "Windows Error for Server Connection!Ensure your network connections are intact", vbInformation
    Exit Function
End Function


Public Function CommonDialogAction(CommonDialog As Control, sAction As String, sMsg As String, sLastFilePath As String, Optional vFilter, Optional vFont) As Variant
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


Private Sub cmdRefresh_Click()
Dim fso As New FileSystemObject

    If frmPatients.piPatientID <> 0 Then
        piPatientID = frmPatients.piPatientID
        If fso.FolderExists("\\" & fstrNetSerPath & piPatientID) Then
            dirNavigator.Path = "\\" & fstrNetSerPath & piPatientID
        Else
            dirNavigator.Path = "\\" & fstrNetSerPath
        End If
    End If
End Sub

Private Sub dirNavigator_Change()
    filNavigator.Path = dirNavigator.Path
    Me.Caption = vbNullString
    Me.Caption = dirNavigator.Path
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
Dim Arr
Dim lngPatient As Long

On Error GoTo picError

'If fbByPass Then Exit Sub

Screen.MousePointer = vbHourglass

    If Me.Caption <> vbNullString Then
        Me.Caption = dirNavigator.Path
    End If
    
'get_path

    If Right$(filNavigator.Path, 1) <> "\" Then
        If piPatientID <> 0 Then
            Me.Caption = Me.Caption & "\" & filNavigator.FileName
        Else
            Me.Caption = Me.Caption & "\" & filNavigator.FileName
        End If
    Else
        Me.Caption = filNavigator.Path & filNavigator.FileName
    End If


If InStr(UCase(filNavigator.FileName), ".TIF") > 0 Then

    Screen.MousePointer = vbDefault
    
    fbooTIF = True
    fbooTEXT = False
    fbooIMG = False
    
    fbByPass = True
    
ElseIf InStr(UCase(filNavigator.FileName), ".RTF") > 0 Or InStr(UCase(filNavigator.FileName), ".TXT") > 0 Or InStr(UCase(filNavigator.FileName), ".SQL") > 0 Then

    
    'If InStr(UCase(me.Caption), ".TXT") > 0 Then
      '  txtRTF.LoadFile me.Caption, 1
   ' Else
      '  txtRTF.LoadFile me.Caption
  '  End If
    
    fbooTIF = False
    fbooTEXT = True
    fbooIMG = False
Else
    ImgThumbnail.Image = vbNullString

 '   imgView.Picture = LoadPicture(me.Caption)

    
    fbooTIF = False
    fbooTEXT = False
    fbooIMG = True
 
End If

    StartingAddress = Me.Caption
    If Len(StartingAddress) > 0 Then
      '  try to navigate to the starting address
       ' timTimer.Enabled = True
       ' If InStr(StartingAddress, "?") > 0 Then
          WebBrowser1.navigate StartingAddress ' & "&RemoteIP=" & me.Caption
        'Else
          ' brwWebBrowser.Navigate StartingAddress & "?RemoteIP=" & sIP_Addr
        'End If
    End If

    fbByPass = False
    Screen.MousePointer = vbDefault

Exit Sub

picError:
    If Err.Number = 438 Or Err.Number = 1002 Then
        If Not fbooTIF Then MsgBox Err.Description, vbExclamation, "Error # " & Err.Number
        Resume Next
    Else
     '   ImgViewer.ClearDisplay
      '  imgView.Picture = LoadPicture()
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


    On Error GoTo ACTIVATE_ERROR
    If mFirst Then
        mFirst = False
        Get_Path
        Me.Top = frmMDI.Top + 1500
        Me.Left = 0
        Me.Height = frmMDI.Height - 2000
        WebBrowser1.Height = Me.Height
        filNavigator.Pattern = "*.bmp;*.wmf;*.ico;*.gif;*.jpg;*.tif;*.mbs;*.thm;*.rtf;*.txt;*.sql"
        If fstrNetSerPath = vbNullString Then
            drvNavigator.Drive = "P:\"
        Else
            drvNavigator.Drive = fstrNetSerPath
        End If
        If fstrNetSerPath <> vbNullString Then
            cmdRefresh_Click
        End If
    End If
    Exit Sub
    
ACTIVATE_ERROR:
    
    MsgBox Err.Description & " " & "Error Loading Path!" & " " & fstrNetSerPath, vbInformation
    Exit Sub
  
End Sub


Private Sub Form_Load()
Dim iFileNum As Integer
Dim sIniFile As String
Dim sContent As String
Dim sLastPath As String
Dim lstrRemarks As String
Dim lbooImagePathExists As Boolean
Dim sImagePath As String

On Error GoTo Error_Viewer

        mFirst = True

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
    mbDontNavigateNow = True
    If IsFormLoaded("frmInPatient") Then
        frmInPatient.Show
    End If
    Unload Me
    Exit Sub
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
   Dim sIP_Addr As String
   Dim strSQL As String
   Dim vRS
   Dim bRet As Boolean
   Dim iRows As Long
   Dim obj As Object
   Dim vDBArr
   Dim sURL As String
   Dim lngURLPatientID As Long
   Dim lngURLEncounterID As Long
   Dim sDBType As String
   Dim sLinkType As String
   Dim sProcName As String
   Dim sActxName As String
   Dim sActxClass As String
   Dim sCaptionName As String
   Dim cn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim strPractice As String
   Dim sIP_Name As String
   
   mbDontNavigateNow = True
   
 
    
   If InStr(StartingAddress, "?") > 0 Then
      URL = StartingAddress '------------& "&RemoteIP=" & sIP_Addr
   Else
      URL = StartingAddress ' ------------ & "?RemoteIP=" & sIP_Addr
   End If
   mbDontNavigateNow = False
   Exit Sub
    
    
End Sub

