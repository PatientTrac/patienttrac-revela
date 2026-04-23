VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmNewEnctr 
   Caption         =   "New Visit"
   ClientHeight    =   5655
   ClientLeft      =   4395
   ClientTop       =   4095
   ClientWidth     =   12150
   Icon            =   "frmNewEnctr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   12150
   Begin VB.Frame fraEncounterDefaults 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Create New Visit "
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
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   12075
      Begin VB.TextBox cctReason 
         Height          =   825
         Left            =   150
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   960
         Width           =   6345
      End
      Begin PT_XP_Button.PT_XP cmdVisitTerms 
         Height          =   255
         Left            =   5940
         TabIndex        =   18
         Top             =   780
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   450
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmNewEnctr.frx":094A
         BackColorDown   =   15133676
         TransparentColor=   13160660
         Caption         =   ""
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmNewEnctr.frx":0CE4
         ShowFocus       =   -1  'True
      End
      Begin VB.ListBox lstVisitReason 
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
         Height          =   1110
         Left            =   150
         TabIndex        =   17
         Top             =   1110
         Width           =   6345
      End
      Begin VB.Frame fraProcedure 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Visit Type"
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
         Height          =   2145
         Left            =   180
         TabIndex        =   15
         Top             =   2400
         Width           =   6345
         Begin VB.ListBox lstProcedures 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1845
            Left            =   30
            Style           =   1  'Checkbox
            TabIndex        =   16
            Top             =   270
            Width           =   6195
         End
      End
      Begin VB.Frame fraService 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   6570
         TabIndex        =   5
         Top             =   1200
         Width           =   5415
         Begin VB.ComboBox cboServicePlace 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   3465
         End
         Begin VB.CommandButton cmdServicePlaceTerms 
            Caption         =   "..."
            Height          =   255
            Left            =   6000
            TabIndex        =   8
            ToolTipText     =   "Maintain Service Places"
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox cboEncounterType 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   7
            ToolTipText     =   "Applies to encounter selected"
            Top             =   690
            Width           =   3465
         End
         Begin VB.CommandButton cmdEncounterType 
            Caption         =   "..."
            Height          =   255
            Left            =   6000
            TabIndex        =   6
            Top             =   690
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblServicePlace 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Place of Service"
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
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   1485
         End
         Begin VB.Label lblEnctrType 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Encounter Type"
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
            Left            =   120
            TabIndex        =   10
            Top             =   750
            Width           =   1545
         End
      End
      Begin VB.ComboBox cboFacility 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   3465
      End
      Begin VB.ComboBox cboProvider 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   870
         Width           =   3465
      End
      Begin PT_XP_Button.PT_XP cmdSelCancel 
         Height          =   435
         Left            =   6330
         TabIndex        =   12
         Top             =   5010
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmNewEnctr.frx":1036
         BackColorDown   =   15133676
         TransparentColor=   13160660
         Caption         =   "Close"
         DepthEvent      =   1
         PictureDisabled =   "frmNewEnctr.frx":1A30
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSelOK 
         Height          =   435
         Left            =   4890
         TabIndex        =   13
         Top             =   5010
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmNewEnctr.frx":23B2
         BackColorDown   =   15133676
         TransparentColor=   13160660
         Caption         =   "Save"
         DepthEvent      =   1
         PictureDisabled =   "frmNewEnctr.frx":2DAC
         ShowFocus       =   -1  'True
      End
      Begin MSComCtl2.DTPicker ccdStartDate 
         Height          =   375
         Left            =   1350
         TabIndex        =   20
         Top             =   360
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   8388608
         CalendarTitleBackColor=   16761024
         CalendarTitleForeColor=   8388608
         CalendarTrailingForeColor=   8421504
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/dd/yyy hh:mm:ss"
         DateIsNull      =   -1  'True
         Format          =   59768835
         CurrentDate     =   38862.7362384259
         MinDate         =   37257
      End
      Begin VB.Label lblVisitDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Visit Date"
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
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblFacility 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Facility"
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
         Left            =   6630
         TabIndex        =   4
         Top             =   450
         Width           =   765
      End
      Begin VB.Label lblProvider 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Provider"
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
         Left            =   6630
         TabIndex        =   3
         Top             =   870
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmNewEnctr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFacility As String
Private mstrProvider As String
Public piProviderID As Long
Private mFirst As Boolean
Private mintProviderID As Integer

 
Private Function Get_SurgProc(vlngPatientID As Long)
Dim strSQL As String

    Open_DB psConnect(0)
    strSQL = "Select cosmetic_procedure from COSMETIC_ORDER where status = 'Final' AND patient_id = " & vlngPatientID
    Set rs = cn.Execute(strSQL)
    With rs
        Do Until .EOF
            If (!COSMETIC_ORDER) <> vbNullString Then lstProcedures.AddItem (!cosmetic_procedure)
        .MoveNext
        Loop
    End With
    Close_DB psConnect(0)
    Exit Function
End Function

Private Function WriteData()
Dim strSQL As String
Dim strIntrvn As String
Dim lngTxnmy As Long
Dim strUser As String
Dim i As Integer
Dim strAppt As String
Dim strStart As String
Dim strEnd As String
Dim strVisitReason As String
Dim strVisitType As String
Dim lstrRemarks As String
Dim lngrecs As Long
Dim varRS
Dim booRet As Boolean
Dim lngProviderId As Long
Dim lngFacilityID As Long
Dim sDate As String
Dim varFacilityName

On Error GoTo Error_WriteData

LogFile "Private Function WriteData()-----------------------------------"

    Screen.MousePointer = vbHourglass

    strUser = PsUserID

    If clsEncounter.VisitReason = vbNullString Then
        clsEncounter.VisitReason = lstVisitReason.Text
    Else
        If clsEncounter.VisitReason <> lstVisitReason.Text Then
            clsEncounter.VisitReason = clsEncounter.VisitReason & " " & lstVisitReason.Text
        End If
    End If
    
    If ccdStartDate.Text = vbNullString Then ccdStartDate.Text = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    
    strStart = Format(ccdStartDate.Text, "mm/dd/yyyy hh:mm:ss AM/PM")
    
    
    'If mboonencntr Then
    If cboServicePlace.Text <> vbNullString Then
        clsEncounter.ServicePlace = cboServicePlace.Text
    End If
    If clsEncounter.ServicePlace = vbNullString Then
        clsEncounter.ServicePlace = clsFacility.FacType
    End If
    
    If cboEncounterType.Text <> vbNullString Then
        clsEncounter.EncounterType = cboEncounterType.Text
    End If
    If clsEncounter.EncounterType = vbNullString Then
        clsEncounter.EncounterType = clsEncounter.EncounterType
    End If

    If cboProvider.Text <> vbNullString Then
        For i = 0 To cboProvider.ListCount - 1
            If cboProvider.List(i) = cboProvider.Text Then
                lngProviderId = cboProvider.ItemData(i)
                Exit For
            End If
        Next i
    End If
    
    For i = 0 To lstProcedures.ListCount - 1
        If lstProcedures.Selected(i) Then
            If strVisitType = vbNullString Then
                strVisitType = lstProcedures.List(i)
            Else
                strVisitType = strVisitType & "; " & lstProcedures.List(i)
            End If
        End If
    Next i
   If strAppt = vbNullString Then
        strAppt = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
   Else
        strAppt = Format(strAppt, "mm/dd/yyyy hh:mm:ss AM/PM")
   End If
    If lngProviderId = vbNullString Or lngProviderId = 0 Then
        lngProviderId = piProviderID
    End If
    
    If cboFacility.Text <> vbNullString Then
        For i = 0 To cboFacility.ListCount - 1
            If cboFacility.List(i) = cboFacility.Text Then
                varFacilityName = cboFacility.List(i)
                strSQL = "select facility_id from facilities where facility_name = '" & varFacilityName & "'"
                booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
                If booRet Then
                    If lngrecs > 0 Then
                        lngFacilityID = varRS(0, 0)
                    End If
                End If
                Exit For
            End If
        Next i
    End If
    
    If lngFacilityID = vbNullString Or lngFacilityID = 0 Then
        lngFacilityID = piFacilityID
    End If

        strSQL = "insert into ENCOUNTER(patient_id, start_date, facility_id, provider_id, appt_date, visit_reason, " _
                & "ENCNTR_TYPE, SERVICE_PLACE, Visit_Type, user_id) values (" & piPatientID & "," _
                & "'" & strStart & "'," & lngFacilityID & "," & lngProviderId & ",'" & strAppt & "'," _
                & "'" & clsEncounter.VisitReason & "','" & clsEncounter.EncounterType & "', '" & clsEncounter.ServicePlace & "','" & strVisitType & "','" & strUser & "')"
        Set rs = cn.Execute(strSQL)
        
        strSQL = "select max(encounter_id)  " _
                & "From encounter " _
                & "Where patient_id = " & piPatientID
                
        booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
            
        If booRet Then
            If lngrecs > 0 Then
                piEncounterID = varRS(0, 0)
                modSpecific.piEncounterID = piEncounterID
                frmPatients.piEncounterID = piEncounterID
            End If
        End If

    Close_DB psConnect(0)
    Screen.MousePointer = vbDefault

    Exit Function

Error_WriteData:
    ErrHandler _
    "frmNewEnctr.WriteData", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function



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
    WriteData
    Unload Me
  
End Sub

Private Sub cmdServicePlaceTerms_Click()
    pstrMaintType = "ServicePlace"
    frmTermsMaint.Show vbModal
End Sub


Private Sub cmdVisitTerms_Click()
    Load frmNurseTemplates
    frmNurseTemplates.fCaption = "Visit Reason"
    frmNurseTemplates.fType = "Visit Reason"
    frmNurseTemplates.fSubDomain = "sd_Encounter_VisitReason"
    frmNurseTemplates.fConnectCR = psConnect(0)
    frmNurseTemplates.fConnectTM = psConnect(1)
    frmNurseTemplates.ReadData
    frmNurseTemplates.Show vbModal
    lstVisitReason.Clear
    GetList lstVisitReason, "sd_Encounter_VisitReason"
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


                strSQL = "Select FACILITY_NAME, FACILITY_ID " _
                        & "From FACILITIES " _
                        & "Order By FACILITY_NAME"
            
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

    ErrHandler _
     Me.Name & " Sub Load_Facilities", _
     Err.Number, _
     Err.Description
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
Dim r%


On Error GoTo Error_Load_EncounterTypes

LogFile "Private Sub Load_EncounterTypes()-----------------------------------"

        lstrSQL = "Select Distinct ENCOUNTER_STATUS, ENCTRTYPE_GROUP_ID " _
                & "From HCFA_CODES WHERE ENCOUNTER_TYPE = '" & vstrService & "'"
            
    lstrRemarks = lstrSQL
    LogFile lstrSQL
    lbooRet = adoSelectQuery(lstrSQL, psConnect(0), lvarRS, , lintRows)
    
    If lintRows > 0 Then
                For r = LBound(lvarRS, 2) To UBound(lvarRS, 2) 'Loop rows
                    If lvarRS(0, r) <> " " Then
                        If IsNull(lvarRS(0, r)) Then lvarRS(0, r) = ""
                        cboEncounterType.AddItem lvarRS(0, r)
                        cboEncounterType.ItemData(cboEncounterType.NewIndex) = lvarRS(1, r)
                    End If
                Next r
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
Dim r%


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
                For r = LBound(lvarRS, 2) To UBound(lvarRS, 2) 'Loop rows
                    If lvarRS(0, r) <> " " Then
                        If IsNull(lvarRS(0, r)) Then lvarRS(0, r) = ""
                           ' If cboServicePlace.List <> lvarRS(0, r) Then
                            cboServicePlace.AddItem lvarRS(0, r)
                            cboServicePlace.ItemData(cboServicePlace.NewIndex) = lvarRS(1, r)
                            'End If
                    End If
            Next r
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
Dim booRet As Boolean
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
Dim lngSeqNO As Long
Dim r%


    On Error GoTo Error_Get_Defaults
       
        strUser = PsUserID
        If mintProviderID = vbNullString Then
        ' Get Provider ID
        strSQL = "select provider_id from providers where " _
            & "USER_ID = '" & strUser & "'"
        booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
        If booRet Then
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
        
        
        
    booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
    
    If booRet Then
        If lngrecs > 0 Then
            For i = LBound(varRS, 2) To UBound(varRS, 2)
                If varRS(0, i) = strEnctrPlace Then
                    For r = 0 To cboServicePlace.ListCount - 1
                    If cboServicePlace.List(r) = varRS(2, i) Then
                            cboServicePlace.Text = varRS(2, i)
                            strEnctrPlace = cboServicePlace.Text
                            Exit For
                        End If
                    Next r
                End If

                If varRS(0, i) = strFacility Then
                    For r = 0 To cboFacility.ListCount - 1
                        If cboFacility.List(r) = varRS(2, i) Then
                            cboFacility.Text = varRS(2, i)
                            Exit For
                        End If
                    Next r
                End If
                
                If varRS(0, i) = strProvider Then
                    For r = 0 To cboProvider.ListCount - 1
                        If cboProvider.List(r) = varRS(2, i) Then
                            cboProvider.Text = varRS(2, i)
                            Exit For
                        End If
                    Next r
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
        
    booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)

    If booRet Then
        If lngrecs > 0 Then
          For i = LBound(varRS, 2) To UBound(varRS, 2)
            If cboServicePlace.Text <> vbNullString Then Load_EncounterTypes strEnctrPlace
                If varRS(0, i) = strEnctrType Then
                    For r = 0 To cboEncounterType.ListCount - 1
                        If cboEncounterType.List(r) = varRS(2, i) Then
                            cboEncounterType.Text = varRS(2, i)
                            Exit For
                        End If
                    Next r
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
Dim booRet As Boolean
Dim lngrecs As Long
Dim strOCX As String
Dim booinsert As Boolean
Dim Arr
Dim lngSeqNO As Long
Dim strUser As String
Dim lngClinicalId As Long

     On Error GoTo Error_Set_Default
    If lngSeqNO = vbNullString Then lngSeqNO = 1
    booRet = False
        strUser = PsUserID
        If mintProviderID = vbNullString Then
        ' Get Provider ID
            strSQL = "select provider_id from providers where " _
                & "USER_ID = '" & strUser & "'"
            booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
            If booRet Then
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
              booRet = SelectQuery(strSQL, psConnect(0), varRS, lngrecs)
             
             If booRet Then
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
                                 & lngSeqNO & " )"
             Else  'Update
             
             If vstrEntrValue <> vbNullString Then
                 strSQL = "Update DEFAULT_CLINICAL " _
                     & "Set clinical_value = '" & vstrEntrValue & "'," _
                     & "clinical_value_ref = '" & vstrEncntrDesc & "'," _
                     & "clinical_desc = '" & strOCX & "'," _
                     & "SEQUENCE_NO = " & lngSeqNO _
                     & " Where provider_id = " & mintProviderID & " And  CLINICAL_ID = " & lngClinicalId
              End If
              
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

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    For Each ctl In frm.Controls
      Set ctl = Nothing
    Next ctl
    For Each obj In frm.objects
      Set obj = Nothing
    Next obj
  Exit Sub
End Sub

Private Sub lstVisitReason_Click()
    
    If lstVisitReason.Text = "Surgery" Then
        Get_SurgProc piPatientID
        fraProcedure.Visible = True
        MsgBox "Please select Surgery Procedure.", vbInformation
    End If
End Sub


