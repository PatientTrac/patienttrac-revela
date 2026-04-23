VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmLoadControl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   11190
   ClientLeft      =   4125
   ClientTop       =   2280
   ClientWidth     =   16335
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   Icon            =   "frmSurgLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   16276.23
   ScaleMode       =   0  'User
   ScaleWidth      =   16335
   Begin VB.Frame fraButton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   6210
      TabIndex        =   0
      Top             =   10380
      Width           =   3045
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   405
         Left            =   1410
         TabIndex        =   1
         ToolTipText     =   "Save, Restore, Close"
         Top             =   30
         Width           =   975
         _ExtentX        =   1720
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
         Picture         =   "frmSurgLog.frx":094A
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSurgLog.frx":135C
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSave 
         Height          =   405
         Left            =   330
         TabIndex        =   2
         ToolTipText     =   "Save, Restore, Close"
         Top             =   30
         Width           =   975
         _ExtentX        =   1720
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
         Picture         =   "frmSurgLog.frx":16AE
         AppearanceThemes=   3
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Save"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmSurgLog.frx":1A48
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmLoadControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private piSelPatientID As Long
Public piSelEnctrID As Long
Private piProviderID As Long
Private piFacilityID As Long
Private piPatientExtID As Long
Private mintActXId As Long
Private mintProviderID As Integer
Private msEncounterDate As String
Private mstrProviderName As String
Private lintComponentCounter As Integer
Private lintTabCounter As Integer
Private lbooMissingClass As Boolean
Private lstrClassName As String
Private lstrLibrary As String
Private mlngCurrentPatientID As Long
Private mlngCurrentEncounterID As Long
Private mstrGridName As String
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private mlngFormID As Long
Private mstrIniFile As String
Private mbooByPass As Boolean
'Dimension collection to keep track of loaded clinical components
Private mcolActxComp As New Collection
Private mbooSelectPat As Boolean
Private mbooRecurssive As Boolean
Private mlngLastPatientId As Long
Private mlngLastEncounterId As Long
Private mObjSelectPatient As PatientTracTypeLib.iClinical
'Private WithEvents PatientList As SelectPatient
Public mintindex As Integer
Private mFirst As Boolean
Private mstrRemarks As String
Private mbooNext As Boolean
Private mbooPrev As Boolean
Public EnctrType As String
Private mbooNoData As Boolean
Private mlngSelFacId As Long
Private mintGridRow As Integer
Private mintGridCol As Integer
Public mstrLoadType As String






Private Sub Load_Control()
Dim strOCX As String
Dim strControl As String
Dim lstrControl As String
Dim lstrSQL As String
Dim strSQL As String
Dim lbooReadOnly As Boolean
Dim cnGetTerms As New ADODB.Connection
Dim rsGetTerms As New ADODB.Recordset
 

    On Error GoTo Error_Load_Control

    If cnGetTerms.State = adStateClosed Then cnGetTerms.Open psConnect(1)
    
    lstrSQL = "Select ACTX_ID FROM ACTIVE_X " _
        & "Where ACTX_NAME = '" & mstrPatOcx & "' And " _
        & "ACTX_CLASS = '" & mstrPatCtl & "'"
    Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
    With rsGetTerms
        If Not .EOF Then
            If Not IsNull(!actx_id) Then
                mintActXId = (!actx_id)
            End If
        End If
    End With
    If mintActXId = 0 Then
            strSQL = "INSERT INTO ACTIVE_X (ACTX_NAME, ACTX_CLASS,ACTX_REL_DIR, ACTX_FILE_NAME, ACTX_VERSION) " _
                & "values('" & mstrPatOcx & "', '" & mstrPatCtl & "', 'Components', '" & mstrPatOcx & "', '2')"
            Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
            
            strSQL = "Select ACTX_ID FROM ACTIVE_X " _
            & "Where ACTX_NAME = '" & mstrPatOcx & "'  And " _
            & "ACTX_CLASS = '" & mstrPatCtl & "'"
            Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
            With rsGetTerms
                If Not .EOF Then
                    If Not IsNull(!actx_id) Then
                        mintActXId = (!actx_id)
                    End If
                End If
            End With
    End If
    If cnGetTerms.State = adStateOpen Then
        cnGetTerms.Close
        Set cnGetTerms = Nothing
        If rsGetTerms.State = adStateOpen Then
            rsGetTerms.Close
           Set rsGetTerms = Nothing
        End If
    End If
    Set fobjOCX = Nothing
    'If Licenses.Count > 0 Then Licenses.Clear
    Licenses.Add mstrPatOcx & "." & mstrPatCtl
    Set fobjOCX = Controls.Add(mstrPatOcx & "." & mstrPatCtl, "ctl" & mstrPatCtl, Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 172
        fobjOCX.Top = 24
        Set fobjLib = fobjOCX.object
        fobjLib.ComponentID = mintActXId
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = modGeneral.piPatientID
        fobjLib.EncounterID = modGeneral.piEncounterID
        fobjLib.ProviderID = modGeneral.piProviderID
        fobjLib.ReadClinicalData
        fraButton.Left = (Me.Width / 2) - (fraButton.Width / 2)
        fraButton.Top = Val(fobjOCX.Height) + 150
    '    Me.Height = (fobjOCX.Height) + 800
        Me.Width = fobjOCX.Width + 300
     Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Control:
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If
    MsgBox Err.Description, vbInformation
    Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub



Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
    fobjLib.WriteClinicalData
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()

   If mFirst = True Then
        mFirst = False
        CenterForm Me
   End If
   Screen.MousePointer = vbDefault
   Exit Sub

End Sub

Private Sub Form_Load()
'Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
'' Size of Form in Pixels at design resolution
'    DesignX = 800
'    DesignY = 600
'    RePosForm = True   ' Flag for positioning Form
'    DoResize = False   ' Flag for Resize Event
'    ' Set up the screen values
'    Xtwips = Screen.TwipsPerPixelX
'    Ytwips = Screen.TwipsPerPixelY
'    Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
'    Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution
'     ScaleFactorX = (Xpixels / DesignX)
'    ScaleFactorY = (Ypixels / DesignY)
'    ScaleMode = 1  ' twips
'    'Exit Sub  ' uncomment to see how Form1 looks without resizing
'    Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'  '  Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
'     "  by " + Str$(Ypixels)
'    MyForm.Height = Me.Height ' Remember the current size
'    MyForm.Width = Me.Width
    mFirst = True
    Load_Control
    Screen.MousePointer = vbDefault
Exit Sub
End Sub

Public Property Get FormID() As Long
    FormID = mlngFormID
End Property

Public Property Let FormID(ByVal vlngNewValue As Long)
    mlngFormID = vlngNewValue
End Property

Private Sub Form_Resize()
'      Dim ScaleFactorX As Single, ScaleFactorY As Single
'
'      If Not DoResize Then  ' To avoid infinite loop
'         DoResize = True
'         Exit Sub
'      End If
'
'      RePosForm = False
'      ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
'      ScaleFactorY = Me.Height / MyForm.Height
'      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'      MyForm.Height = Me.Height ' Remember the current size
'      MyForm.Width = Me.Width

        fraButton.Left = (Me.Width / 2) - (fraButton.Width / 2)
        fraButton.Top = Val(fobjOCX.Height) + 150
     '   Me.Height = (fraButton.Top) + 800
        
        Exit Sub
End Sub

Public Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object

  On Error Resume Next
    Set mcolActxComp = Nothing
    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.Objects
      Set obj = Nothing
    Next obj
  Exit Sub
End Sub












