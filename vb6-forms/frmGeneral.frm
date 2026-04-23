VERSION 5.00
Begin VB.Form frmGeneral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EMR Dash Board General"
   ClientHeight    =   4770
   ClientLeft      =   6960
   ClientTop       =   3390
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraControlPanel 
      Height          =   765
      Left            =   30
      TabIndex        =   1
      Top             =   4020
      Visible         =   0   'False
      Width           =   6000
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3150
         TabIndex        =   3
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2010
         TabIndex        =   2
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame fraControl 
      Height          =   4000
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6000
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngFormID As Long
Private mlngPatientID As Long
Private mlngEncounterID As Long
Private mlngProviderID As Long
Private msEncounterDate As String
Private mstrDBType As String

Dim lintComponentCounter As Integer
Dim lintTabCounter As Integer
Dim lbooMissingClass As Boolean
Dim mlstrClassName As String
Dim mlstrActxName As String
Dim lstrLibrary As String

'dnn.................
Private mintTabs As Long
Private mintTabsPerRow As Long

'*** Note!   This collection below is a rudimentary first attempt to handle this.
'            We need to create a specific collection that allows to manage the
'            loaded controls on a page by page basis, instead of keeping all
'            controls for all pages in the same bag.

'Dimension collection to keep track of loaded clinical components
Private mcolActxComp As New Collection
Private mbooSelectPat As Boolean
Private mbooRecurssive As Boolean
Private mlngLastPatientId As Long
Private mlngLastEncounterId As Long
Private mObjSelectPatient As PatientTracTypeLib.iClinical
'Private WithEvents PatientList As SelectPatient

Private Declare Function ClientToScreen Lib "USER32" (ByVal hWnd As Long, lpPoint As CPOINT) As Long

Private Type CPOINT
        X As Long
        Y As Long
End Type

Dim objOCX As VBControlExtender
Dim liclComponent As PatientTracTypeLib.iClinical


Private Sub SetDlgPosition(ByVal objButton As Control, ByRef objForm As Form)
    Dim Point As CPOINT
    Dim llngReturn As Long
    llngReturn = ClientToScreen(objButton.hWnd, Point)
    objForm.Left = (Point.X - 4) * Screen.TwipsPerPixelX + objButton.Width + 50
    objForm.Top = (Point.Y - 4) * Screen.TwipsPerPixelY - objForm.Height + objButton.Height + 50
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

'Private Sub cmdOptions_Click()
'    Load DialogControlPanel
'    SetDlgPosition cmdOptions, DialogControlPanel
'    DialogControlPanel.Show vbModal
'End Sub

Private Sub cmdOK_Click()
    
    If mlstrActxName = "Encounter" Then
       Unload Me
       Exit Sub
    End If
    
    'Cycle through all loaded components
    'For Each liclComponent In mcolActxComp
        'Test if the componet indicates to be dirty
        If liclComponent.HasChanged Then
           If Not objOCX.Object.ValidateData Then Exit Sub
            'Call the component's write routine
            liclComponent.WriteClinicalData
        End If
    'Next
    Unload Me
End Sub
Public Property Get DBType() As String
    DBType = mstrDBType
End Property

Public Property Let DBType(ByVal vstrNewValue As String)
    mstrDBType = vstrNewValue
End Property
Public Property Get FormID() As Long
    FormID = mlngFormID
End Property

Public Property Let FormID(ByVal vlngNewValue As Long)
    mlngFormID = vlngNewValue
End Property

Public Property Get PatientID() As Long
    PatientID = mlngPatientID
End Property

Public Property Let PatientID(ByVal vlngNewValue As Long)
    mlngPatientID = vlngNewValue
End Property

Public Property Get EncounterID() As Long
    EncounterID = mlngEncounterID
End Property

Public Property Let EncounterID(ByVal vlngNewValue As Long)
    mlngEncounterID = vlngNewValue
End Property

Public Property Get ProviderID() As Long
    ProviderID = mlngProviderID
End Property

Public Property Let ProviderID(ByVal vlngNewValue As Long)
    mlngProviderID = vlngNewValue
End Property

Private Sub RefreshComponents()
    'Dimension reference for component control to add
    
    'Cycle through all loaded components
    'For Each liclComponent In mcolActxComp
    'Establish component clinical properties
    If mlstrActxName <> "Encounter" Then
       liclComponent.ConnectStringTerms = psConnect(1)
       liclComponent.ConnectStringClinical = psConnect(0)
       liclComponent.PatientID = piPatientID
       liclComponent.EncounterID = piEncounterID
       objOCX.Object.DBType = mstrDBType
       'liclComponent.ProviderID = piProviderID
       liclComponent.ComponentID = GetActxID(mlstrClassName, mlstrActxName, psConnect(1))
       'Force component to retrieve any existing clinical data
       liclComponent.ReadClinicalData
       If pbViewOnly Then objOCX.Object.ViewOnly
     Else
       objOCX.ConnectStringTerms = psConnect(1)
       objOCX.ConnectStringClinical = psConnect(0)
       objOCX.PatientID = piPatientID
       objOCX.EncounterID = piEncounterID
       objOCX.ReadClinicalData
     End If
 'Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lbooComponentsChanged As Boolean
    
    Dim lintCounter As Integer
    
    If mlstrActxName = "Encounter" Then Exit Sub

    'Cycle through all loaded components
    'For Each liclComponent In mcolActxComp
        'Test if the componet indicates to be dirty
        If liclComponent.HasChanged Then
        
            'Test if this is the first dirty component
            If lbooComponentsChanged = False Then
                'Load dialog
                lbooComponentsChanged = True
                Load frmSaveAlert
                frmSaveAlert.lstComponents.Clear
            End If
            
            'Add dirty compontent to dialog
            With frmSaveAlert!lstComponents
                .AddItem TypeName(liclComponent)
                .Selected(.NewIndex) = True
                .ItemData(.NewIndex) = liclComponent.ComponentID
            End With
            
        End If
        
    'Next
    
    'Test if any components were found dirty
    If lbooComponentsChanged Then
        
        'Show dialog and wait for user response
        frmSaveAlert!lstComponents.ListIndex = 0
        frmSaveAlert.Show vbModal, Me
        
        'When user closes dialog, test cancel was not hit
        If Not frmSaveAlert.Cancelled Then
            With frmSaveAlert!lstComponents
                'Cycle thru list items
                For lintCounter = 0 To .ListCount - 1
                    'Test if item is selected for save
                    If .Selected(lintCounter) Then
                        'For Each liclComponent In mcolActxComp
                            'If .ItemData(lintCounter) = liclComponent.ComponentID Then
                             If liclComponent.HasChanged Then
                                If Not objOCX.Object.ValidateData Then
                                   Cancel = -1
                                   Exit Sub
                                End If
                                liclComponent.WriteClinicalData
                                'mcolComponents(Format$(.ItemData(lintCounter), "\i\d\_00000")).WriteClinicalData
                            End If
                        'Next
                    End If
                Next
            End With
        
        'Free component reference resources
        Set liclComponent = Nothing
        
        Else
            'Cancel form unload if cancel was hit at the dialog
            Cancel = True
        End If
        'Free dialog's resources
        Unload frmSaveAlert
        Set frmSaveAlert = Nothing
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object

  On Error Resume Next
    If mlstrActxName <> "Encounter" Then Licenses.Remove (0)
    Set mcolActxComp = Nothing
    Set mObjSelectPatient = Nothing
    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
  Exit Sub
End Sub

Private Sub SetPatientInfo()
 ' LblPatientId.Caption = ""
  'If piPatientID > 0 Then LblPatientId.Caption = piPatientID
  'LblPatient.Caption = Trim(psPatientName)
  'LblEncounterDate.Caption = Trim(msEncounterDate)
  'frmMDI.StatusBar.Panels(2).Text = "Encounter: " & piEncounterID
  'frmMDI.StatusBar.Panels(1).Text = "Patient: " & psPatientName & " ID: " & piPatientID
End Sub

Private Sub Form_Activate()

    'Visible property of the basic tab is used as a flag to indicate if
    'the initialization of the page has executed.
    
    If fraControl.Visible = False Then
    
        'Initialization needs to take place.
        DoEvents
        
        Screen.MousePointer = vbHourglass
        
        fraControl.Visible = False  'dnn

        fraControlPanel.Visible = False
        
       ' DistributeAllComponents

        Call RefreshComponents
        
        'Set initial state of basic tab
        'tabBasic.Tab = 0
        'tabBasic.TabIndex = 0
        
        'Show the initialized tab
        '(and prevent from initializing again)
        
        
        'Modified and commented by KP on 17/12/2002
        'Me.Width = tabBasic.Width + fraControlPanel.Width + 900

'        Me.Width = fraControl.Width + fraControlPanel.Width + 500
        
     
 '       Me.Height = fraControl.Height + fraPatientInfo.Height + 500
    
'        fraControl.Top = frPatientInfo.Height + 100
 '       fraControl.Left = 100

        
  '      fraControlPanel.Left = fraControl.Left + fraControl.Width + 100
   

        
        fraControl.Visible = True
        fraControlPanel.Visible = True
        
        
        'SetPatientInfo
        fraControl.Height = objOCX.Height + 200
        fraControl.Width = objOCX.Width + 200
        fraControlPanel.Top = fraControl.Top + fraControl.Height + 50
        Me.Height = fraControlPanel.Top + fraControlPanel.Height + 400
        Me.Width = fraControl.Width + 200
        fraControlPanel.Width = fraControl.Width
        cmdOK.Left = fraControlPanel.Width / 2 - cmdOK.Width
        cmdClose.Left = fraControlPanel.Width / 2 + 50
        
        centerForm Me
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub Form_Load()
    
    mlngLastPatientId = 0
    mlngLastEncounterId = 0
    centerForm Me
    Call SetPatientInfo
    
    'Call InitaliseComp
    
    'Visible property of the basic tab is used as a flag to indicate if
    'the initialization of the page has executed.
    'Make sure is set to false upon load to tell the activate routine
    'that the tab pages need to be initialized.
    fraControl.Visible = False
End Sub


Public Function GetActxID(ByVal lstrClass As String, ByVal lstrActx As String, ByVal lstrTerms As String) As Long
'Added by KP
    Dim lCon As New ADODB.Connection
    Dim lRs As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo GetActxID_Error:
   
    strSQL = "Select ACTX_ID from Active_X where ACTX_CLASS='" & lstrClass & "' and ACTX_NAME= '" & lstrActx & "'"
    lCon.open lstrTerms
    If lCon.State = 1 Then
        lRs.open strSQL, lCon
        
        If Not lRs.EOF And Not lRs.BOF Then
            GetActxID = lRs(0)
        End If
        If lRs.State = 1 Then lRs.Close
        lCon.Close
    End If
    If Not lRs Is Nothing Then Set lRs = Nothing
    If Not lCon Is Nothing Then Set lCon = Nothing
    Exit Function
GetActxID_Error:
    If Not lRs Is Nothing Then Set lRs = Nothing
    If Not lCon Is Nothing Then Set lCon = Nothing
End Function

