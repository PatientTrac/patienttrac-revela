VERSION 5.00
Begin VB.Form frmEnctrDefault 
   Caption         =   "Encounter Default"
   ClientHeight    =   4875
   ClientLeft      =   10680
   ClientTop       =   4080
   ClientWidth     =   5010
   Icon            =   "frmEnctrDefault.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   5010
   Begin VB.Frame fraEncounterDefaults 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Encounter Default Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   6585
      Begin VB.TextBox cctReason 
         Height          =   1065
         Left            =   270
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2670
         Width           =   4395
      End
      Begin VB.Frame fraPatientSel 
         BackColor       =   &H00E0E0E0&
         Height          =   675
         Left            =   1470
         TabIndex        =   12
         Top             =   3900
         Width           =   1935
         Begin VB.CommandButton cmdSelCancel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   180
            Width           =   765
         End
         Begin VB.CommandButton cmdSelOK 
            BackColor       =   &H00FFC0C0&
            Caption         =   "OK"
            Height          =   375
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   180
            Width           =   795
         End
      End
      Begin VB.Frame fraService 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   90
         TabIndex        =   5
         Top             =   1200
         Width           =   4005
         Begin VB.ComboBox cboServicePlace 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   2145
         End
         Begin VB.CommandButton cmdServicePlaceTerms 
            Caption         =   "..."
            Height          =   255
            Left            =   3540
            TabIndex        =   8
            ToolTipText     =   "Maintain Service Places"
            Top             =   270
            Width           =   255
         End
         Begin VB.ComboBox cboEncounterType 
            Height          =   315
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   7
            ToolTipText     =   "Applies to encounter selected"
            Top             =   690
            Width           =   2145
         End
         Begin VB.CommandButton cmdEncounterType 
            Caption         =   "..."
            Height          =   255
            Left            =   3540
            TabIndex        =   6
            Top             =   720
            Width           =   255
         End
         Begin VB.Label lblServicePlace 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Place of Service"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   1170
         End
         Begin VB.Label lblEnctrType 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Encounter Type"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   750
            Width           =   1140
         End
      End
      Begin VB.ComboBox cboFacility 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   3465
      End
      Begin VB.ComboBox cboProvider 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   870
         Width           =   2295
      End
      Begin VB.Label lblFacility 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Facility"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   450
         Width           =   480
      End
      Begin VB.Label lblProvider 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Provider"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   30
         TabIndex        =   3
         Top             =   870
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmEnctrDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFacility As String
Private mstrProvider As String
Private mFirst As Boolean
Private mintProviderID As Integer

Private Sub cboFacility_Click()
Dim IntId As Integer
Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim i As Integer
Dim lstrRemarks As String

    On Error GoTo Error_ERROR_HANDLER

        If cboFacility.ListIndex > -1 Then
            IntId = cboFacility.ItemData(cboFacility.ListIndex)
        End If
        
            If IntId > 0 Then
            'Get Providers......................................................
                strSQL = "Select rtrim(first_name) + ' ' + (case when middle_name <> '' then middle_name else '' end) + ' ' + rtrim(last_name) + ', ' + rtrim(title) as PROVIDER_NAME, PROVIDER_ID " _
                        & "From PROVIDERS " _
                        & "Where provider_id in (Select provider_id From FACILITY_PROVIDERS Where facility_id = " & IntId & ") " _
                        & "Order By PROVIDER_NAME"
            Else
                strSQL = "Select rtrim(first_name) + ' ' + (case when middle_name <> '' then middle_name else '' end) + ' ' + rtrim(last_name) + ', ' + rtrim(title) as PROVIDER_NAME, PROVIDER_ID " _
                        & "From PROVIDERS "
            End If
            
            bRet = SelectQuery(strSQL, psConnect(0), vRS, iRows)
            
            If bRet Then
                If iRows > 0 Then
                    LoadCombo cboProvider, vRS
                End If
            End If
            
    Exit Sub
Error_ERROR_HANDLER:

     ErrHandler _
     Me.Name & " cboFacility_Click", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub cboServicePlace_Click()
Dim strServicePlace As String

    strServicePlace = cboServicePlace.Text
    Load_EncounterTypes strServicePlace
    
End Sub


Private Sub cmdEncounterType_Click()
    pstrMaintType = "EncounterType"
    frmTermsMaint.Show vbModal
End Sub

Private Sub cmdSelCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelOK_Click()
Dim strEnctrType As String
Dim strEnctrPlace As String
Dim strFacility As String
Dim strProvider As String
Dim strReasonVisit As String



    
     strEnctrType = "cboEncounterType"
     strEnctrPlace = "cboServicePlace"
     strFacility = "cboFacility"
     strProvider = "cboProvider"
     strReasonVisit = "cctReason"
     clsEncounter.EncounterType = cboEncounterType.Text
     clsEncounter.ServicePlace = cboServicePlace.Text
     mstrFacility = cboFacility.Text
     mstrProvider = cboProvider.Text
     clsEncounter.VisitReason = cctReason.Text
     
    Set_Default strEnctrType, clsEncounter.EncounterType
    Set_Default strEnctrPlace, clsEncounter.ServicePlace
    Set_Default strFacility, mstrFacility
    Set_Default strProvider, mstrProvider
    Set_Default strReasonVisit, clsEncounter.VisitReason
    Unload Me
  
End Sub

Private Sub cmdServicePlaceTerms_Click()
    pstrMaintType = "ServicePlace"
    frmTermsMaint.Show vbModal
End Sub


Private Sub Form_Activate()
If mFirst Then
    frmEnctrDefault.Top = 4800
    frmEnctrDefault.Left = 6700
    Load_Facilities
    Load_Service
    Get_Defaults
    mFirst = False
    Screen.MousePointer = vbDefault
End If
End Sub


Private Sub Load_Facilities()
Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim i As Integer
Dim lstrRemarks As String
    
    On Error GoTo Error_Load_Facilities
    strSQL = "SELECT FACILITY_NAME, FACILITY_ID " _
        & "From FACILITIES WHERE ACTIVE = 1 ORDER BY FACILITY_NAME"
    lstrRemarks = strSQL
    LogFile strSQL
    bRet = adoSelectQuery(strSQL, psConnect(0), vRS, , iRows)
        If bRet Then
            If iRows > 0 Then
                LoadCombo cboFacility, vRS
            End If
        End If
    Screen.MousePointer = vbDefault
   Exit Sub
Error_Load_Facilities:
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
End Sub

Private Sub Load_EncounterTypes(vstrService As String)
Dim lstrSQL As String
Dim lvarRS
Dim lbooRet As Boolean
Dim lintRows As Long
Dim lstrRemarks As String
Dim i As Integer
Dim R%


On Error GoTo Error_Load_EncounterTypes

LogFile "Private Sub Load_EncounterTypes()-----------------------------------"

        lstrSQL = "Select Distinct ENCOUNTER_STATUS, ENCTRTYPE_GROUP_ID " _
                & "From HCFA_CODES WHERE ENCOUNTER_TYPE = '" & vstrService & "'"
            
    lstrRemarks = lstrSQL
    LogFile lstrSQL
    lbooRet = adoSelectQuery(lstrSQL, psConnect(0), lvarRS, , lintRows)
    
    If lintRows > 0 Then
                For R = LBound(lvarRS, 2) To UBound(lvarRS, 2) 'Loop rows
                    If lvarRS(0, R) <> " " Then
                        If IsNull(lvarRS(0, R)) Then lvarRS(0, R) = ""
                        cboEncounterType.AddItem lvarRS(0, R)
                        cboEncounterType.ItemData(cboEncounterType.NewIndex) = lvarRS(1, R)
                    End If
                Next R
    End If
'    If piEncounterID > 0 Then
'        'Get Encounter type..............................
'        lstrSQL = "Select encntr_type From ENCOUNTER Where encounter_id = " & piEncounterID
'
'        lbooRet = adoSelectQuery(lstrSQL, psConnect(0), lvarRS, , lintRows)
'
'        If lintRows > 0 Then
'            If Not IsNull(lvarRS(0, 0)) Then
'               For i = 0 To cboEncounterType.ListCount - 1
'                   If lvarRS(0, 0) = cboEncounterType.List(i) Then
'                      cboEncounterType.Text = lvarRS(0, 0)
'                   End If
'               Next i
'            End If
'        End If
'    End If
    
    Exit Sub
    
Error_Load_EncounterTypes:
     ErrHandler _
     Me.Name & " Sub Load_EncounterTypes", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub Load_Service()
Dim lstrSQL As String
Dim lvarRS
Dim lbooRet As Boolean
Dim lintRows As Long
Dim lstrRemarks As String
Dim R%


On Error GoTo Error_Load_Service

LogFile "Private Sub Load_Load_Service()-----------------------------------"

            lstrSQL = "Select DISTINCT ENCOUNTER_TYPE, ENCTRTYPE_GROUP_ID " _
                & "From HCFA_CODES "
                
    lstrRemarks = lstrSQL
    LogFile lstrSQL
    lbooRet = adoSelectQuery(lstrSQL, psConnect(0), lvarRS, , lintRows)
    
        If lintRows > 0 Then
       ' For i = LBound(lvarRS, 2) To UBound(lvarRS, 2)
            'If lvarRS(0, i) <> " " Then
                For R = LBound(lvarRS, 2) To UBound(lvarRS, 2) 'Loop rows
                    If lvarRS(0, R) <> " " Then
                        If IsNull(lvarRS(0, R)) Then lvarRS(0, R) = ""
                           ' If cboServicePlace.List <> lvarRS(0, r) Then
                            cboServicePlace.AddItem lvarRS(0, R)
                            cboServicePlace.ItemData(cboServicePlace.NewIndex) = lvarRS(1, R)
                            'End If
                    End If
            Next R
        End If
    
    Exit Sub
    
Error_Load_Service:
    ErrHandler _
     Me.Name & " Sub Load_Service", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub


Private Sub Form_Load()
    mFirst = True
    Screen.MousePointer = vbDefault
Exit Sub
End Sub

Public Function Get_Defaults()
Dim strSQL As String
Dim varRS
Dim booret As Boolean
Dim lngrecs As Long
Dim strOCX As String
Dim booinsert As Boolean
Dim i%
Dim strAppt As String
Dim strStart As String
Dim strEnd As String
Dim strVisitReason As String
Dim lstrRemarks As String
Dim intServCode As Integer
Dim strAction As String
Dim strEnctrType As String
Dim strEnctrPlace As String
Dim strFacility As String
Dim strProvider As String
Dim strReasonVisit As String
Dim strUser As String
Dim lngSeqNo As Long
Dim R%


    On Error GoTo Error_Get_Defaults
       
        strUser = PsUserID
        If mintProviderID = 0 Then
        ' Get Provider ID
        strSQL = "select provider_id from providers where " _
            & "USER_ID = '" & strUser & "'"
        booret = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
        If booret Then
            If lngrecs > 0 Then
                mintProviderID = varRS(0, 0)
                piProviderID = mintProviderID
            End If
        End If
        piProviderID = mintProviderID
    End If

     
    Get_Defaults = vbNullString
    
     strEnctrType = "cboEncounterType"
     strEnctrPlace = "cboServicePlace"
     strFacility = "cboFacility"
     strProvider = "cboProvider"
     strReasonVisit = "cctReason"
     
     strSQL = "select CLINICAL_VALUE_REF, sequence_no, clinical_value " _
        & "From DEFAULT_CLINICAL " _
        & "Where provider_id = " & mintProviderID & " And " _
        & "CLINICAL_DESC = 'frmPatients'"
        
        
        
    booret = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
    
    If booret Then
        If lngrecs > 0 Then
            For i = LBound(varRS, 2) To UBound(varRS, 2)
                If varRS(0, i) = strEnctrPlace Then
                    For R = 0 To cboServicePlace.listcount - 1
                    If cboServicePlace.List(R) = varRS(2, i) Then
                            cboServicePlace.Text = varRS(2, i)
                            strEnctrPlace = cboServicePlace.Text
                            Exit For
                        End If
                    Next R
                End If

                If varRS(0, i) = strFacility Then
                    For R = 0 To cboFacility.listcount - 1
                        If cboFacility.List(R) = varRS(2, i) Then
                            cboFacility.Text = varRS(2, i)
                            Exit For
                        End If
                    Next R
                End If
                
                If varRS(0, i) = strProvider Then
                    For R = 0 To cboProvider.listcount - 1
                        If cboProvider.List(R) = varRS(2, i) Then
                            cboProvider.Text = varRS(2, i)
                            Exit For
                        End If
                    Next R
                End If
                
                If varRS(0, i) = strVisitReason Then
                    cctReason.Text = varRS(2, i)
                End If
             Next i
          Else
            Load_Service
        End If
    End If
    
 'Get Encounter Types
       strSQL = "select CLINICAL_VALUE_REF, sequence_no, clinical_value " _
        & "From DEFAULT_CLINICAL " _
        & "Where provider_id = " & mintProviderID & " And " _
        & "CLINICAL_DESC = 'frmPatients'"
        
    booret = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)

    If booret Then
        If lngrecs > 0 Then
          For i = LBound(varRS, 2) To UBound(varRS, 2)
            If cboServicePlace.Text <> vbNullString Then Load_EncounterTypes strEnctrPlace
                If varRS(0, i) = strEnctrType Then
                    For R = 0 To cboEncounterType.listcount - 1
                        If cboEncounterType.List(R) = varRS(2, i) Then
                            cboEncounterType.Text = varRS(2, i)
                            Exit For
                        End If
                    Next R
                End If
           Next i
        End If
End If

    
    Screen.MousePointer = vbDefault
    
    Exit Function
Error_Get_Defaults:
     ErrHandler _
     Me.Name & " Sub Get_Defaults", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function

Public Sub Set_Default(vstrEncntrDesc As String, vstrEntrValue As String)
Dim strSQL As String
Dim varRS
Dim booret As Boolean
Dim lngrecs As Long
Dim strOCX As String
Dim booinsert As Boolean
Dim Arr
Dim lngSeqNo As Long
Dim strUser As String
Dim lngClinicalId As Long

     On Error GoTo Error_Set_Default
    If lngSeqNo = 0 Then lngSeqNo = 1
    booret = False
        strUser = PsUserID
        If mintProviderID = 0 Then
        ' Get Provider ID
            strSQL = "select provider_id from providers where " _
                & "USER_ID = '" & strUser & "'"
            booret = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
            If booret Then
                If lngrecs > 0 Then
                    mintProviderID = varRS(0, 0)
                End If
            End If
        End If
  strOCX = "frmPatients"
   'Insert or Update?......................
             strSQL = "Select CLINICAL_ID From DEFAULT_CLINICAL " _
                     & "Where PROVIDER_ID = " & mintProviderID & " And " _
                     & "CLINICAL_VALUE_REF = '" & vstrEncntrDesc & "' And " _
                     & "CLINICAL_DESC = '" & strOCX & "'"
              booret = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
             
             If booret Then
                 If lngrecs > 0 Then
                    lngClinicalId = varRS(0, 0)
                     booinsert = False
                 Else
                     booinsert = True
                 End If
             Else
                 Exit Sub
             End If

             If booinsert Then
                 strSQL = "Insert into DEFAULT_CLINICAL (provider_id, clinical_value, clinical_value_ref, clinical_desc, SEQUENCE_NO ) " _
                     & "Values(" & mintProviderID & ",'" _
                                 & vstrEntrValue & "','" _
                                 & vstrEncntrDesc & "','" _
                                 & strOCX & "'," _
                                 & lngSeqNo & " )"
             Else  'Update
             
             If vstrEntrValue <> vbNullString Then
                 strSQL = "Update DEFAULT_CLINICAL " _
                     & "Set clinical_value = '" & vstrEntrValue & "'," _
                     & "clinical_value_ref = '" & vstrEncntrDesc & "'," _
                     & "clinical_desc = '" & strOCX & "'," _
                     & "SEQUENCE_NO = " & lngSeqNo _
                     & " Where provider_id = " & mintProviderID & " And  CLINICAL_ID = " & lngClinicalId
              End If
              
             End If
             
             booret = ActionQuery(strSQL, psConnect(0))
                
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
