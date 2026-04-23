VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmPlastiSurg 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "SURGEON CLINICAL DATA"
   ClientHeight    =   9630
   ClientLeft      =   3525
   ClientTop       =   2625
   ClientWidth     =   12345
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
   Icon            =   "frmPlastiSurg.frx":0000
   ScaleHeight     =   11699.32
   ScaleMode       =   0  'User
   ScaleWidth      =   12345
   Begin VB.Frame fraTB 
      BackColor       =   &H00DDCDCD&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12360
      Begin PT_XP_Button.PT_XP cmdRevEnctr 
         Height          =   735
         Left            =   2220
         TabIndex        =   2
         ToolTipText     =   "Previous, Next, or Current Encounter"
         Top             =   -180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   1296
         BackColor       =   14536141
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceMode  =   2
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Review Encounters"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdPatient 
         Height          =   555
         Left            =   30
         TabIndex        =   3
         ToolTipText     =   "Select / Change Patients"
         Top             =   0
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   979
         BackColor       =   14536141
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Patient"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdClinical 
         Height          =   555
         Left            =   8490
         TabIndex        =   4
         ToolTipText     =   "Save, Restore, Close"
         Top             =   0
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   979
         BackColor       =   14536141
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceMode  =   2
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "Clinical Data              Close"
         DepthEvent      =   1
      End
      Begin PT_XP_Button.PT_XP cmdHistory 
         Height          =   555
         Left            =   5640
         TabIndex        =   0
         ToolTipText     =   "Review Scanned Chart Data"
         Top             =   0
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   979
         BackColor       =   14536141
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   5
         AppearanceThemes=   3
         BackColorDown   =   14536141
         BackColorOver   =   13432063
         BackColorFocus  =   14536141
         BackColorDisabled=   14536141
         BorderDrawEvent =   1
         BorderWidth     =   0
         TransparentColor=   14737632
         Caption         =   "History "
         DepthEvent      =   1
      End
   End
   Begin VB.Menu Encounter 
      Caption         =   "Encounter"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu Prev_Enctr 
         Caption         =   "Previous Encounter"
         Shortcut        =   ^P
      End
      Begin VB.Menu Next_Enctr 
         Caption         =   "Next Encounter"
         Shortcut        =   ^N
      End
      Begin VB.Menu Current_Enctr 
         Caption         =   "Current_Encounter"
         Shortcut        =   ^Z
      End
      Begin VB.Menu New_Encounter 
         Caption         =   "New Encounter"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu Clinical_Data 
      Caption         =   "Clinical Data"
      Visible         =   0   'False
      Begin VB.Menu Save 
         Caption         =   "Save"
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu Restore 
         Caption         =   "Restore"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
         Shortcut        =   ^{F6}
      End
   End
End
Attribute VB_Name = "frmPlastiSurg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public piPatientID As Long
Public piEncounterID As Long
Public piSelEncntrID As Long
Public mbNotLoad As Boolean
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
Dim booRet As Boolean
Dim lngrecs As Long
Dim strOCX As String
Dim booinsert As Boolean
Dim lngSeqNO As Long
Dim strUser

     On Error GoTo Error_Set_Default
    If lngSeqNO = 0 Then lngSeqNO = 1
     
     strOCX = "frmPatients"
     
                 'Insert or Update?......................
                strSQL = "Select count(*) From DEFAULT_CLINICAL " _
                        & "Where PROVIDER_ID = " & mintProviderID & " And " _
                        & "CLINICAL_VALUE_REF  = '" & vstrEncounter & "' And " _
                        & "CLINICAL_DESC = '" & strOCX & "'"
                        
            
                 booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
                
                If booRet Then
                    If lngrecs > 0 Then
                        If varRS(0, 0) = 0 Then booinsert = True
                    End If
                Else
                    Exit Sub
                End If
                
                
                If booinsert Then
                    strSQL = "Insert into DEFAULT_CLINICAL (provider_id, clinical_value, clinical_value_ref, clinical_desc, SEQUENCE_NO ) " _
                        & "Values(" & mintProviderID & ",'" _
                                    & vstrEntrValue & "','" _
                                    & vstrEncounter & "','" _
                                    & strOCX & "'," _
                                    & lngSeqNO & " )"
                Else  'Update
                    strSQL = "Update DEFAULT_CLINICAL " _
                        & "Set clinical_value = '" & vstrEntrValue & "'," _
                        & "clinical_value_ref = '" & vstrEncounter & "'," _
                        & "clinical_desc = '" & strOCX & "'," _
                        & "SEQUENCE_NO = " & lngSeqNO _
                        & " Where provider_id = " & mintProviderID
                        
                End If
                
                booRet = ActionQuery(strSQL, psConnect(0))
                
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

Public Sub Load_ChildForm(vstrOCX As String, vstrControl As String)
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_ChildForm
    mbNotLoad = False


        If frmPlasticSurgery.piPatientID <> 0 Then piPatientID = frmPlasticSurgery.piPatientID
        If piPatientID = 0 Then piPatientID = frmPatients.piPatientID
        If piEncounterID = 0 Then piEncounterID = frmPlasticSurgery.piEncounterID
        If piEncounterID = 0 Then piEncounterID = frmEnctr.piEncounterID
        
    If piPatientID = 0 Then
        MsgBox "Please select Patient!", vbInformation
        Exit Sub
        mbNotLoad = True
        Unload Me
    End If
    
    
    piSelEncntrID = piEncounterID

    
    intACTXID = Get_ComponentID(vstrOCX, vstrControl)
    
    Licenses.Add vstrOCX & "." & vstrControl
    Set fobjOCX = Controls.Add(vstrOCX & "." & vstrControl, "ctl" & vstrControl, Me)
        fobjOCX.Visible = True
        'fraControlPanel.Left = fobjOCX.Width + 125
        frmPlastiSurg.Width = fobjOCX.Width + 1000
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fraTB.Top = 0
        fobjOCX.Left = 150
        fobjOCX.Top = 1000
        fraTB.Width = fobjOCX.Width + 1000
        fobjLib.PatientID = piPatientID
        fobjLib.EncounterID = piEncounterID
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        
        frmPlastiSurg.Height = (fobjOCX.Height * 0.95) + 1000
    Exit Sub
    
Error_Load_ChildForm:

    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If

    ErrHandler _
     "frmPlastiSurg.Load_ChildForm", _
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
Private Sub Close_Click()
    mbooFormLoaded = False
    Unload Me
    modSpecific.piPatientID = piPatientID
    modSpecific.piEncounterID = piEncounterID
    modSpecific.piFacilityID = piFacilityID
    modSpecific.piProviderID = piProviderID
End Sub

Private Sub cmdClinical_Click()
Dim iAnswer As Integer
    iAnswer = MsgBox("Do you want to save changes?", vbYesNoCancel)
    If iAnswer = vbNo Then
        mbooFormLoaded = False
        modSpecific.piPatientID = piPatientID
        modSpecific.piEncounterID = piEncounterID
        modSpecific.piFacilityID = piFacilityID
        modSpecific.piProviderID = piProviderID
        Unload Me
    ElseIf iAnswer = vbYes Then
        fobjLib.WriteClinicalData
        Unload Me
    Else
        Exit Sub
    End If

    Exit Sub
End Sub

Private Sub cmdClinical_DropDownClick()

    PopupMenu Clinical_Data, , cmdClinical.Left + fraTB.Left, fraTB.Top + cmdClinical.Height

End Sub

Private Sub cmdHistory_Click()
    frmViewer.Show vbModal
End Sub


Public Function Get_Defaults(vstrEnctr As String)
Dim strSQL As String
Dim varRS
Dim booRet As Boolean
Dim lngrecs As Long
Dim strOCX As String
Dim booinsert As Boolean
Dim i%
Dim strEnctrType As String
Dim strEnctrPlace As String


    Get_Defaults = vbNullString

     strSQL = "select CLINICAL_VALUE_REF, clinical_value " _
        & "From DEFAULT_CLINICAL " _
        & "Where provider_id = " & mintProviderID & " And " _
        & "CLINICAL_DESC = 'frmPatients' And  " _
        & "CLINICAL_VALUE_REF = '" & vstrEnctr & "'"
                
    booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
    
    If booRet Then
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


Private Sub cmdPatient_Click()

    Unload Me
    modSpecific.piPatientID = piPatientID
    modSpecific.piEncounterID = piEncounterID
    modSpecific.piFacilityID = piFacilityID
    modSpecific.piProviderID = piProviderID
    frmPatients.Show
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdRevEnctr_DropDownClick()

    PopupMenu Encounter, , cmdRevEnctr.Left + fraTB.Left, fraTB.Top + cmdRevEnctr.Height
       
End Sub

Private Sub Current_Enctr_Click()

    piEncounterID = piSelEncntrID
    modSpecific.piEncounterID = piEncounterID
    frmPatients.piPatientID = piPatientID
    frmPatients.piEncounterID = piEncounterID
    fobjLib.EncounterID = piEncounterID
    frmMDI.StatusBar.Panels(1).Text = "Patient: " & clsPatient.Name
    frmMDI.StatusBar.Panels(2).Text = "External ID: " & clsPatient.ExtID
    frmMDI.StatusBar.Panels(3).Text = "Internal ID: " & piPatientID
    frmMDI.StatusBar.Panels(4).Text = "Encounter: " & piEncounterID
    frmPatients.piPatientID = piPatientID
    frmPatients.piEncounterID = piEncounterID
    mbooNext = False
    mbooPrev = False
        'mintindex = lstEncounters.ListIndex
    If Get_LastEncounter > piEncounterID Then
        mbooNext = True
    Else
        mbooNext = False
    End If
    Get_PrevEncounter piEncounterID
    If modSpecific.mlngPrevEnctrID > 0 Then
        mbooPrev = True
    Else
        mbooPrev = False
    End If
    
    MsgBox "Do you want to clear all Reviewed Encounter Data for Current Encounter?", vbYesNo
    If vbYes Then
        fobjLib.ReadClinicalData
    End If

End Sub

Private Sub Form_Activate()

   If mFirst = True Then
        mFirst = False
   End If
   Screen.MousePointer = vbDefault
       centerForm Me
   Exit Sub
    
End Sub

Private Sub Form_Load()

    mlngLastPatientId = 0
    mlngLastEncounterId = 0
    mFirst = True

    mbooNext = frmPatients.mbooNext
    mbooPrev = frmPatients.mbooPrev
    
    If mbNotLoad Then Unload Me
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

  ' frmPlastiSurg.Top = 500
    'frmmdi.Height =
  '  frmPlastiSurg.Height = 11000
   ' centerForm Me
End Sub

Public Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next
    Set mcolActxComp = Nothing
    modSpecific.piPatientID = piPatientID
    modSpecific.piEncounterID = piEncounterID
    modSpecific.piFacilityID = piFacilityID
    modSpecific.piProviderID = piProviderID
    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.objects
      Set obj = Nothing
    Next obj
  Exit Sub

End Sub





Public Function OpenEncounter(vlngPatientID As Long, vlngEncounterID As Long) As Boolean
Dim strSQL As String
Dim varRS
Dim lngRows As Long
Dim booRet As Boolean
Dim lstrRemarks As String
    
On Error GoTo Error_OpenEncounter

LogFile "Private Function OpenEncounter()-----------------------------------"

    strSQL = "Select encounter_id, ENCNTR_END_DATE " _
        & "From Encounter " _
        & "Where patient_id = " & vlngPatientID & " And encounter_id = " & vlngEncounterID
        
   ' lstrRemarks = strSQL
  '  LogFile strSQL
   
     booRet = SelectQuery(strSQL, psConnect(0), varRS, lngRows)
      
    If booRet Then
    
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
    "frmPlastiSurg.OpenEncounter", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function






Private Sub Align_TB()
   ' fraTB.Top = 0
   ' fraTB.Width = Me.Width
    'fobjOCX.Top = 500
    
End Sub

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
Dim strtype As String
Dim strPlace As String
Dim strFacility As String
Dim strProvider As String
Dim varRS
Dim lngrecs As Long
Dim booRet As Boolean
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

    Unload frmPlastiSurg
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
    
    Unload frmPlastiSurg
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


