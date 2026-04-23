VERSION 5.00
Begin VB.Form frmConsent 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "CONSENTS AND TREATMENT"
   ClientHeight    =   3750
   ClientLeft      =   5175
   ClientTop       =   3120
   ClientWidth     =   8925
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
   Icon            =   "frmConsent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4555.809
   ScaleMode       =   0  'User
   ScaleWidth      =   8925
End
Attribute VB_Name = "frmConsent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public piPatientID As Long
Public piEncounterID As Long
Public piSelEncntrID As Long
Public piProviderID As Long
Public piPatientExtID As Long
Public piFacilityID As Long
Public mintProviderID As Integer
Private msEncounterDate As String
Public mstrProviderName As String
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
Private fstrReportDate As Date
Private fstrReportDate2 As Date
Private mbooFormLoaded As Boolean
Private mstrDefProvider As String
Private mstrDefFacility As String
Private mstrIniFile As String
Private mbooGet As Boolean
Private mstrReportFileName As String
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
Public mintindex As Integer
Private mFirst As Boolean
Private mstrRemarks As String
Private mbooNext As Boolean
Private mbooPrev As Boolean

Public Sub Set_Default(vstrEncounter As String, vstrEntrValue As String)
Dim strSQL As String
Dim varRS
Dim booret As Boolean
Dim lngrecs As Long
Dim strOCX As String
Dim booInsert As Boolean
Dim lngSeqNo As Long
Dim strUser

     On Error GoTo Error_Set_Default
    If lngSeqNo = 0 Then lngSeqNo = 1
     
     strOCX = "frmPatients"
     
                 'Insert or Update?......................
                strSQL = "Select count(*) From DEFAULT_CLINICAL " _
                        & "Where PROVIDER_ID = " & mintProviderID & " And " _
                        & "CLINICAL_VALUE_REF  = '" & vstrEncounter & "' And " _
                        & "CLINICAL_DESC = '" & strOCX & "'"
                        
            
                 booret = SelectQuery(strSQL, psconnect(0), varRS, lngrecs)
                
                If booret Then
                    If lngrecs > 0 Then
                        If varRS(0, 0) = 0 Then booInsert = True
                    End If
                Else
                    Exit Sub
                End If
                
                
                If booInsert Then
                    strSQL = "Insert into DEFAULT_CLINICAL (provider_id, clinical_value, clinical_value_ref, clinical_desc, SEQUENCE_NO ) " _
                        & "Values(" & mintProviderID & ",'" _
                                    & vstrEntrValue & "','" _
                                    & vstrEncounter & "','" _
                                    & strOCX & "'," _
                                    & lngSeqNo & " )"
                Else  'Update
                    strSQL = "Update DEFAULT_CLINICAL " _
                        & "Set clinical_value = '" & vstrEntrValue & "'," _
                        & "clinical_value_ref = '" & vstrEncounter & "'," _
                        & "clinical_desc = '" & strOCX & "'," _
                        & "SEQUENCE_NO = " & lngSeqNo _
                        & " Where provider_id = " & mintProviderID
                        
                End If
                
                booret = ActionQuery(strSQL, psconnect(0))
                
    Exit Sub
Error_Set_Default:

    ErrHandler _
     Me.Name & " Sub Set_Default", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume

     
End Sub

Private Sub Load_Consent()
Dim strOCX As String
Dim strControl As String
Dim intActxID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_Consent

    piPatientID = modSpecific.piPatientID
    piEncounterID = modSpecific.piEncounterID
    piFacilityID = modSpecific.piFacilityID
    piProviderID = modSpecific.piProviderID
    piSelEncntrID = piEncounterID
    Validate_PiValues piPatientID, piSelEncntrID, piProviderID, piFacilityID

    strOCX = "FN_PsyMse"
    strControl = "TreatmentPlans"
    
    intActxID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlTreatmentPlans", Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 150
        fobjOCX.Top = 100
        'fraControlPanel.Left = fobjOCX.Width + 125
        frmConsent.Width = fobjOCX.Width + 300
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intActxID
        fobjLib.ConnectStringTerms = psconnect(1)
        fobjLib.ConnectStringClinical = psconnect(0)
      '  cmdClose.Left = fobjOCX.Width - 2000
        fobjLib.PatientID = piPatientID
        fobjLib.EncounterID = piEncounterID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        frmConsent.Height = (fobjOCX.Height * 0.95) + 200
      '  cmdClose.Top = frmConsent.Height - 50
    Exit Sub
    
Error_Load_Consent:

    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If

    ErrHandler _
     "frmConsent.Load_Consent", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub
Public Property Get ReportDate() As String
    ReportDate = fstrReportDate
End Property

Public Property Let ReportDate(ByVal vstrDate As String)
    fstrReportDate = vstrDate
End Property

Public Property Get ReportDate2() As String
    ReportDate2 = fstrReportDate2
End Property

Public Property Let ReportDate2(ByVal vstrDate As String)
    fstrReportDate2 = vstrDate
End Property



Public Function RunReport(vsrtReportFile As String) As Boolean
Dim strFile As String
Dim strRemarks As String
Dim strTitle As String
Dim strFileName As String
Dim strParamCode As String
Dim booPatientID As Boolean
Dim booEncounterD As Boolean
Dim booDate As Boolean
Dim booDate2 As Boolean
Dim strToDate As String
Dim strFromDate As String
   
   
   
   
   'pstrReportScript = vsrtReportFile
   

    If piPatientID = 0 Then
        'check Consult
        MsgBox "Select a patient first.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If piEncounterID = 0 Then
        MsgBox "Select a encounter first.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    strRemarks = strFile
    strFileName = strFile
    
    If Not FileExists(pstrMainPath & "\Reports\" & strFile) Then
        MsgBox "Report file " & strFile & " does not exist.", vbCritical
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    

    'Decode parameters used.....................
    If InStr(strFileName, "_") > 0 Then
        strParamCode = Mid(strFileName, InStr(strFileName, "_") + 1)
        strParamCode = Left(strParamCode, Len(strParamCode) - 4)
    End If
    
    'Strip Parameter code.....................
    If Right(strFileName, 1) = "8" Then
        strTitle = Left(strFileName, Len(strFileName) - 10)
    Else
        strTitle = Left(strFileName, Len(strFileName) - 9)
    End If
    
    strTitle = Separate_Words(strTitle)
    
    If Not IsNumeric(strParamCode) Then
        MsgBox "Not a valid file name format for this EMR application.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    booPatientID = Left(strParamCode, 1)
    booEncounterD = Mid(strParamCode, 2, 1)
    booDate = Mid(strParamCode, 3, 1)
    booDate2 = Mid(strParamCode, 4, 1)
    
    
    If booPatientID And piPatientID = 0 Then
       MsgBox "Select a Patient First.", vbExclamation
       Screen.MousePointer = vbDefault
       Exit Function
    End If

    If booEncounterD And piEncounterID = 0 Then
       MsgBox "Select an encounter First.", vbExclamation
       Screen.MousePointer = vbDefault
       Exit Function
    End If
    
    frmReportDate.ccDate.Visible = booDate
    frmReportDate.ccToDate.Visible = booDate2
  
    If booDate And Not booDate2 Then

        CenterControl frmReportDate.ccDate, frmReportDate.picReportDate
        frmReportDate.Show vbModal
        strFromDate = Format(fstrReportDate, "yyyy,mm,dd")
        
        If Not IsDate(strFromDate) Then
            MsgBox "Entry " & strFromDate & " is not a date." & vbNewLine & "Try again.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    If booDate And booDate2 Then
 
        frmReportDate.Show vbModal
        
        strFromDate = Format(fstrReportDate, "yyyy,mm,dd")
        
        If Not IsDate(strFromDate) Then
            MsgBox "Entry " & strFromDate & " is not a date." & vbNewLine & "Try again.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        
        strToDate = Format(fstrReportDate2, "yyyy,mm,dd")
        
        If Not IsDate(strToDate) Then
            MsgBox "Entry " & strToDate & " is not a date." & vbNewLine & "Try again.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    ' wayne
    If Not booDate And booDate2 Then
     
        CenterControl frmReportDate.ccToDate, frmReportDate.picReportDate
        frmReportDate.Show vbModal
        strToDate = Format(fstrReportDate2, "yyyy,mm,dd")
        
        If Not IsDate(strToDate) Then
            MsgBox "Entry " & strToDate & " is not a date." & vbNewLine & "Try again.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    ' wayne
    strToDate = Format(Now, "yyyy,mm,dd")
    pstrReportScript = pstrMainPath & "\Reports\" & strFile
    ' frmreport.fstrTitle = strTitle
    
    RunReport = True
    Unload Me
    ' frmreport.Show vbModal

    Exit Function
    
End Function

Public Sub ReadClinicalData()
   ' Read_Location cboFacility, fstrClinicalConnect
   ' Get_Location cboFacility, cboProvider, fstrClinicalConnect, piEncounterID, piPatientId
End Sub

Public Function WriteData(vstrAction As String) As Boolean
Dim strSQL As String
Dim strIntrvn As String
Dim lngTxnmy As Long
Dim intFacilityID As Integer
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
'    LogErrHandler _ _
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
Public Function Get_Defaults(vstrEnctr As String)
Dim strSQL As String
Dim varRS
Dim booret As Boolean
Dim lngrecs As Long
Dim strOCX As String
Dim booInsert As Boolean
Dim i%
Dim strEnctrType As String
Dim strEnctrPlace As String


    Get_Defaults = vbNullString

     strSQL = "select CLINICAL_VALUE_REF, clinical_value " _
        & "From DEFAULT_CLINICAL " _
        & "Where provider_id = " & mintProviderID & " And " _
        & "CLINICAL_DESC = 'frmPatients' And  " _
        & "CLINICAL_VALUE_REF = '" & vstrEnctr & "'"
                
    booret = SelectQuery(strSQL, psconnect(0), varRS, lngrecs)
    
    If booret Then
        If lngrecs > 0 Then
            For i = LBound(varRS, 2) To UBound(varRS, 2)
                If varRS(0, i) = vstrEnctr Then
                  Get_Defaults = varRS(1, i)
                  Exit For
                End If
            Next i
        End If
    End If
    

Exit Function
End Function


'Private Sub cmdPatient_Click()
'
'    Unload Me
'    modSpecific.piPatientID = piPatientID
'    modSpecific.piEncounterID = piEncounterID
'    modSpecific.piFacilityID = piFacilityID
'    modSpecific.piProviderID = piProviderID
'    frmPatients.Show
'    Screen.MousePointer = vbDefault
'End Sub


Private Sub Form_Activate()

   If mFirst = True Then
        mFirst = False
   End If
   Screen.MousePointer = vbDefault
       centerForm Me
   Exit Sub
    
End Sub

Private Sub Form_Load()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    mlngLastPatientId = 0
    mlngLastEncounterId = 0
    mFirst = True

    mbooNext = frmPatients.mbooNext
    mbooPrev = frmPatients.mbooPrev
    
       
    LogFile "Start Log  Consent --------------------", True
 
    Load_Consent
    Screen.MousePointer = vbDefault

Exit Sub

End Sub

Public Property Get FormID() As Long
    FormID = mlngFormID
End Property

Public Property Let FormID(ByVal vlngNewValue As Long)
    mlngFormID = vlngNewValue
End Property

Public Property Get PatientID() As Long
    PatientID = piPatientID
End Property

Public Property Let PatientID(ByVal vlngNewValue As Long)
    piPatientID = vlngNewValue
End Property

Public Property Get EncounterID() As Long
    EncounterID = piEncounterID
End Property

Public Property Let EncounterID(ByVal vlngNewValue As Long)
    piEncounterID = vlngNewValue
End Property

Public Property Get ProviderID() As Long
    ProviderID = piProviderID
End Property

Public Property Let ProviderID(ByVal vlngNewValue As Long)
    piProviderID = vlngNewValue
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

Private Sub Form_Resize()

  ' frmConsent.Top = 500
    'frmmdi.Height =
  '  frmConsent.Height = 11000
   ' centerForm Me
End Sub

Public Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.Objects
      Set obj = Nothing
    Next obj
    modSpecific.piPatientID = piPatientID
    modSpecific.piEncounterID = piEncounterID
    modSpecific.piFacilityID = piFacilityID
    modSpecific.piProviderID = piProviderID
End Sub





Public Function OpenEncounter(vlngPatientID As Long, vlngEncounterID As Long) As Boolean
Dim strSQL As String
Dim varRS
Dim lngRows As Long
Dim booret As Boolean
Dim lstrRemarks As String
    
On Error GoTo Error_OpenEncounter

LogFile "Private Function OpenEncounter()-----------------------------------"

    strSQL = "Select encounter_id, ENCNTR_END_DATE " _
        & "From Encounter " _
        & "Where patient_id = " & vlngPatientID & " And encounter_id = " & vlngEncounterID
        
   ' lstrRemarks = strSQL
  '  LogFile strSQL
   
     booret = SelectQuery(strSQL, psconnect(0), varRS, lngRows)
      
    If booret Then
    
        If lngRows > 0 Then
            If Not IsNull(varRS(1, 0)) Then
                If IsDate(varRS(1, 0)) Then
                    OpenEncounter = False
                Else
                    OpenEncounter = True
                End If
            Else
                OpenEncounter = True
            End If
        End If
    
    End If
    
    Exit Function
    
Error_OpenEncounter:

    ErrHandler _
    "frmConsent.OpenEncounter", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function






Private Sub New_Encounter_Click()
Dim strSQL As String
Dim strIntrvn As String
Dim lngTxnmy As Long
Dim intFacilityID As Integer
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
Dim strType As String
Dim strPlace As String
Dim strFacility As String
Dim strProvider As String
Dim varRS
Dim lngrecs As Long
Dim booret As Boolean
Dim strAction As String


On Error GoTo Error_cmdNewEncounter_Click

    Screen.MousePointer = vbHourglass
    frmEnctr.Create_NewEncounter True


Error_cmdNewEncounter_Click:

    ErrHandler _
     Me.Name & " cmdNewEncounter_Click", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume

End Sub

Private Sub Next_Enctr_Click()
Dim lintIndex As Integer
Dim lintResp As Integer
Dim booSkipSave As Boolean
Dim strSQL As String
Dim varRS
Dim lngrecs As Long
Dim bRet As Boolean
Dim lngNextID As Long
Dim i%

    Unload frmConsent
    frmPatients.Form_Load
    frmPatients.Show
    frmPatients.cmdEncntr_Click

'
'   If mbooNext = False Then
'        MsgBox "Current encounter is the last encounter!", vbInformation
'        Exit Sub
'    End If
'
'    If Not OpenEncounter(piPatientID, piEncounterID) Then
'        'Closed Encounter...............
'        If pbooLockClosedEncounterData Then
'            Screen.MousePointer = vbDefault
'            booSkipSave = True
'            Exit Sub
'        End If
'    End If
'
'
'    strSQL = "select encounter_id from encounter where patient_id = " & piPatientID & " order by encounter_id"
'
'
'    bRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
'
'    If bRet Then
'        If lngrecs > 1 Then
'            For i = LBound(varRS, 2) To UBound(varRS, 2)
'                If varRS(0, i) = piEncounterID Then
'                    lngNextID = varRS(0, i + 1)
'                ElseIf varRS(0, i) <= 0 Then
'                    MsgBox "Current Encounter is the last Encounter!", vbCritical
'                    Exit Sub
'                End If
'            Next i
'        Else
'           MsgBox "There are no additional encounters!", vbCritical
'        End If
'    End If
'    Screen.MousePointer = vbHourglass
'
'    piEncounterID = lngNextID
'    modSpecific.piEncounterID = piEncounterID
'    Get_Encounter
'    fobjLib.UserCallSub
'    fobjLib.PatientID = piEncounterID
'    fobjLib.EncounterID = piEncounterID
'    fobjLib.ReadClinicalData
'    Screen.MousePointer = vbDefault
    Exit Sub
Exit Sub

End Sub


Private Sub Prev_Enctr_Click()
Dim lintIndex As Integer
Dim lintResp As Integer
Dim booSkipSave As Boolean
Dim strSQL As String
Dim varRS
Dim lngrecs As Long
Dim bRet As Boolean
Dim lngNextID As Long
Dim i%
    
    Unload frmConsent
    frmPatients.Form_Load
    frmPatients.Show
    frmPatients.cmdEncntr_Click
    
    
'   If mbooPrev = False Then
'        MsgBox "No Previous Encounters", vbInformation
'        Exit Sub
'    End If
'
'    If Not OpenEncounter(piPatientID, piEncounterID) Then
'        'Closed Encounter...............
'        If pbooLockClosedEncounterData Then
'            Screen.MousePointer = vbDefault
'            booSkipSave = True
'            Exit Sub
'        End If
'    End If
'
'    If Not booSkipSave Then
'
'            If mbooHasChanged Then
'                lintResp = MsgBox("Do You Want To Save Before" & vbNewLine & "Looking At The Previous Encounter?", vbInformation + vbYesNo, "Save ComponentID ")
'                'If lintResp = vbYes Then cmdSave_Click
'                If lintResp = vbYes Then fobjLib.WriteClinicalData
'            End If
'    End If
'
'    Screen.MousePointer = vbHourglass
'
'    If Get_PrevEncounter <> 0 Then
'        If Get_PrevEncounter = piEncounterID Then
'            MsgBox "There are no previous encounters!", vbCritical
'        Else
'            piEncounterID = Get_PrevEncounter
'        End If
'    End If
'    modSpecific.piEncounterID = piEncounterID
'    Get_Encounter
'
'    fobjLib.EncounterID = piEncounterID
'    fobjLib.UserCallSub
'    fobjLib.ReadClinicalData
'    Screen.MousePointer = vbDefault

End Sub

Private Sub Restore_Click()
        fobjLib.ReadClinicalData
End Sub

Private Sub Save_Click()
    fobjLib.WriteClinicalData
End Sub




