VERSION 5.00
Begin VB.Form frmUPMC 
   Caption         =   "PatientTrac TRAUMA PROGRESS NOTE"
   ClientHeight    =   10215
   ClientLeft      =   7050
   ClientTop       =   2340
   ClientWidth     =   13650
   Icon            =   "UPMC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   13650
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0080C0FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   11460
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9420
      Width           =   1575
   End
   Begin VB.CommandButton cmdManagePatients 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add/Manage Patients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11460
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8850
      Width           =   1605
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "..."
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   9150
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblReportDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4470
      TabIndex        =   2
      Top             =   9180
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      Caption         =   "Report"
      Height          =   195
      Left            =   6690
      TabIndex        =   1
      Top             =   9180
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmUPMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPatientID As Long
Private mlngEncounterID As Long
Private mlngProviderID As Long
Private msEncounterDate As String
Private lintComponentCounter As Integer
Private lintTabCounter As Integer
Private lbooMissingClass As Boolean
Private lstrClassName As String
Private lstrLibrary As String
Private mlngCurrentPatientID As Long
Private mlngCurrentEncounterID As Long
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private mlngFormID As Long
'dnn.................
Private mintTabs As Long
Private mintTabsPerRow As Long
Private fstrTermsConnect As String
Private fstrClinicalConnect As String
Private mbooFormLoaded As Boolean
Private mstrDefProvider As String
Private mstrDefFacility As String
Private mstrIniFile As String
Private mbooGet As Boolean
'*** Note!   This collection below is a rudimentary first attempt to handle this.
'            We need to create a specific collection that allows to manage the
'            loaded controls on a page by page basis, instead of keeping all
'            controls for all pages in the same bag.
Private mbooformload As Boolean
'Dimension collection to keep track of loaded clinical components
Private mcolActxComp As New Collection
Private mbooSelectPat As Boolean
Private mbooRecurssive As Boolean
Private mlngLastPatientId As Long
Private mlngLastEncounterId As Long
Private mObjSelectPatient As PatientTracTypeLib.iClinical
'Private WithEvents PatientList As SelectPatient

Private Sub Load_UPMC()
Dim strOCX As String
Dim strControl As String
Dim intActxID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_UPMC
    
    strOCX = "FN_Trauma"
    strControl = "UPMCProg"
    
    intActxID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlUPMCProg", Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 30
        fobjOCX.Top = 30
        
    Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intActxID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        cmdManagePatients.Top = fobjOCX.Height - 1500
        cmdManagePatients.Left = cmdManagePatients.Left + 50
        cmdClose.Top = cmdManagePatients.Top + cmdManagePatients.Height + 50
        cmdClose.Left = cmdManagePatients.Left
        Me.Width = fobjOCX.Width + 200
        Me.Height = fobjOCX.Height + 600
        centerForm Me
    Exit Sub
    
Error_Load_UPMC:

    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If

    ErrHandler _
     Me.Name & " Load_UPMC", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub



Public Function RunReport(vsrtReportFile As String) As Boolean
   pstrReportScript = vsrtReportFile
   
   ' frmreport.piPatientId = fobjLib.PatientID
  '  frmreport.piEncounterID = fobjLib.EncounterID
    
    If piPatientID = 0 Then
        MsgBox "Select a patient first.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If piEncounterID = 0 Then
        MsgBox "Select a encounter first.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    RunReport = True
    Unload Me
  '  frmreport.Show vbModal
End Function

Public Sub MedPrint()
Dim strReportFile As String
Dim mxtalReport As Object
Dim lstrRemarks As String

LogFile "Private Sub MedPrint-----------------------------------"

    strReportFile = "C:\Program Files\FirstNotebook\PatientTrac EMR\Prescript2.rpt"
    Set mxtalReport = CreateObject("Crystal.CrystalReport")
    
    mxtalReport.StoredProcParam(0) = "R"
    mxtalReport.StoredProcParam(1) = mlngPatientID
    mxtalReport.StoredProcParam(2) = mlngEncounterID
    mxtalReport.StoredProcParam(3) = mlngProviderID
    mxtalReport.ReportFileName = strReportFile
    mxtalReport.WindowShowPrintBtn = True
    mxtalReport.WindowShowPrintSetupBtn = True
    mxtalReport.Destination = 0
    mxtalReport.Action = 1

End Sub

Public Sub ReadClinicalData()
   ' Read_Location cboFacility, fstrClinicalConnect
   ' Get_Location cboFacility, cboProvider, fstrClinicalConnect, mlngEncounterID, mlngPatientID
End Sub

Public Function WriteData(vstrAction As String) As Boolean
Dim strSQL As String
Dim strIntrvn As String
Dim lngTxnmy As Long
Dim intFacilityID As Integer
Dim intProviderID As Integer
Dim strUser As String
Dim i As Integer
Dim strAppt As String
Dim strStart As String
Dim strEnd As String
Dim strVisitReason As String
Dim lstrRemarks As String
Dim intServCode As Integer
Dim strServPlace As String
Dim strEncntrType As String
    
'On Error GoTo Error_WriteData

'LogFile "Private Function WriteData()-----------------------------------"


'    Screen.MousePointer = vbHourglass
'
'    strUser = Get_UserName()
'
'    i = IdentifyControl("ctlDuration")
'    strAppt = fctlObj(i).Appointment_Date
'    strStart = fctlObj(i).Start_Date
'    strEnd = fctlObj(i).End_Date
'
'    If cboServicePlace.ListIndex > -1 Then
'        intServCode = cboServicePlace.ItemData(cboServicePlace.ListIndex)
'        strServPlace = Literals(cboServicePlace.Text, "'")
'    End If
'
'    If cboEncounterType.ListIndex > -1 Then
'        strEncntrType = Literals(cboEncounterType.Text, "'")
'    End If
'
''i = IdentifyControl("ctlType")
''strIntrvn = fctlObj(i).Data_Intervention
''lngTxnmy = fctlObj(i).Data_Taxonomy
'
'    strVisitReason = fctlPatient.VisitReason
'
'    i = IdentifyControl("ctlLocation")
'    intFacilityID = fctlObj(i).Data_Facility
'    intProviderID = fctlObj(i).Data_Provider
'
'     If vstrAction = "E" Then
'              strSQL = "exec a_encounter_clinical " _
'              & "@action_cd = '" & vstrAction & "'," _
'              & "@patient_id = " & pipatientid & "," _
'              & "@encounter_id = " & pipatientid & "," _
'              & "@start_date = '" & strStart & "'," _
'              & "@end_date = '" & strEnd & "'," _
'              & "@intrvn_cd = '" & strIntrvn & "'," _
'              & "@txnmy_cd = " & lngTxnmy & "," _
'              & "@facility_id = " & intFacilityID & "," _
'              & "@provider_id = " & intProviderID & "," _
'              & "@appt_date = '" & strAppt & "'," _
'              & "@visit_reason = '" & strVisitReason & "'," _
'              & "@ENCNTR_TYPE = '" & strEncntrType & "'," _
'              & "@SERVICE_PLACE = '" & strServPlace & "'," _
'              & "@SERVICE_PLACE_CD = " & intServCode & "," _
'              & "@user_id = '" & strUser & "'"
'     Else
'
'         If IsDate(strAppt) Then
'             strSQL = "exec a_encounter_clinical " _
'              & "@action_cd = '" & vstrAction & "'," _
'              & "@patient_id = " & pipatientid & "," _
'              & "@encounter_id = " & pipatientid & "," _
'              & "@start_date = '" & strStart & "'," _
'              & "@end_date = NULL," _
'              & "@intrvn_cd = '" & strIntrvn & "'," _
'              & "@txnmy_cd = " & lngTxnmy & "," _
'              & "@facility_id = " & intFacilityID & "," _
'              & "@provider_id = " & intProviderID & "," _
'              & "@appt_date = '" & strAppt & "'," _
'              & "@visit_reason = '" & strVisitReason & "'," _
'              & "@ENCNTR_TYPE = '" & strEncntrType & "'," _
'              & "@SERVICE_PLACE = '" & strServPlace & "'," _
'              & "@SERVICE_PLACE_CD = " & intServCode & "," _
'              & "@user_id = '" & strUser & "'"
'
'        Else
'             strSQL = "exec a_encounter_clinical " _
'              & "@action_cd = '" & vstrAction & "'," _
'              & "@patient_id = " & pipatientid & "," _
'              & "@encounter_id = " & pipatientid & "," _
'              & "@start_date = '" & strStart & "'," _
'              & "@end_date = NULL," _
'              & "@intrvn_cd = '" & strIntrvn & "'," _
'              & "@txnmy_cd = " & lngTxnmy & "," _
'              & "@facility_id = " & intFacilityID & "," _
'              & "@provider_id = " & intProviderID & "," _
'              & "@appt_date = NULL," _
'              & "@visit_reason = '" & strVisitReason & "'," _
'              & "@ENCNTR_TYPE = '" & strEncntrType & "'," _
'              & "@SERVICE_PLACE = '" & strServPlace & "'," _
'              & "@SERVICE_PLACE_CD = " & intServCode & "," _
'              & "@user_id = '" & strUser & "'"
'        End If
'
'    End If
'
'    lstrRemarks = strSQL
'    LogFile strSQL
'
'    WriteData = adoActionQuery(strSQL, fstrClinicalConnect)
'
'    Screen.MousePointer = vbDefault
'
'    Exit Function
'
'Error_WriteData:
'
'    LogErrorHandler _
'     psConnect(0), _
'     Me.Name & " Function WriteData", _
'     Err.Number, _
'     Err.Description, _
'     App.EXEName, _
'     App.Path, _
'     App.Major & "." & App.Minor & "." & App.Revision, _
'     piPatientId, _
'     piEncounterID, _
'     Get_UserName(), _
'     Get_ComputerName(), _
'     lstrRemarks
'     Screen.MousePointer = vbDefault
'     Exit Function

End Function

Private Sub cmdClose_Click()
    mbooFormLoaded = False
    Unload Me
End Sub




Private Sub cmdManagePatients_Click()
    Unload Me
    frmPatients.Show vbModal
    Exit Sub
End Sub



Private Sub Form_Activate()

    'Visible property of the basic tab is used as a flag to indicate if
    'the initialization of the page has executed.
    
    'Me.Visible = False
   
    LogFile "Start Log  UPMC --------------------", True

        
        Load_UPMC

        
        centerForm Me


        Screen.MousePointer = vbDefault
    
    'End If
   
    
    
End Sub

Private Sub Form_Load()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    mlngLastPatientId = 0
    mlngLastEncounterId = 0
    centerForm Me

    

End Sub

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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lbooComponentsChanged As Boolean
    Dim liclComponent As PatientTracTypeLib.iClinical
    Dim lintCounter As Integer
    
    Exit Sub
    '---------------by pass this ----------------------------------------------------
    'Cycle through all loaded components
    For Each liclComponent In mcolActxComp
        'Test if the componet indicates to be dirty
        If liclComponent.HasChanged Then
            'Test if this is the first dirty component
            If lbooComponentsChanged = False Then
                'Load dialog
                lbooComponentsChanged = True
                Load frmSaveAlert
            End If
            
            'Add dirty compontent to dialog
            With frmSaveAlert!lstComponents
                .AddItem TypeName(liclComponent)
                .Selected(.NewIndex) = True
                .ItemData(.NewIndex) = liclComponent.ComponentID
            End With
            
        End If
        
    Next
    
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
                        For Each liclComponent In mcolActxComp
                            'If .ItemData(lintCounter) = liclComponent.ComponentID Then
                             If liclComponent.HasChanged Then
                                liclComponent.WriteClinicalData
                                'mcolComponents(Format$(.ItemData(lintCounter), "\i\d\_00000")).WriteClinicalData
                            End If
                        Next
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

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.Objects
      Set obj = Nothing
    Next obj
  Exit Sub

End Sub





Private Sub PatientList_EncounterClick(ByVal AEncounterId As Long, ByVal AEncounterDate As String)
'   piEncounterID = AEncounterId
'   mlngEncounterID = piEncounterID
'   msEncounterDate = Trim(AEncounterDate)
'   SetPatientInfo
End Sub

Private Sub PatientList_EncounterDblClick()
   'If tabBasic.Tab = 0 Then tabBasic.Tab = 1
End Sub


