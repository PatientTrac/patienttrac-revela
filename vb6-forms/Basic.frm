VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBasic 
   Caption         =   "Basic"
   ClientHeight    =   9510
   ClientLeft      =   960
   ClientTop       =   1560
   ClientWidth     =   13245
   Icon            =   "Basic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   13245
   Begin VB.PictureBox picPatientSel 
      BackColor       =   &H0080C0FF&
      Height          =   7905
      Left            =   840
      ScaleHeight     =   7845
      ScaleWidth      =   8565
      TabIndex        =   20
      Top             =   1020
      Visible         =   0   'False
      Width           =   8625
      Begin VB.Frame fraFilter 
         BackColor       =   &H0080C0FF&
         Caption         =   "Search Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1155
         Left            =   4620
         TabIndex        =   30
         Top             =   240
         Width           =   3735
         Begin VB.OptionButton optByPatientName 
            BackColor       =   &H0080C0FF&
            Caption         =   "By Last Name"
            Height          =   195
            Left            =   150
            TabIndex        =   34
            Top             =   300
            Value           =   -1  'True
            Width           =   1365
         End
         Begin VB.OptionButton optByExternalId 
            BackColor       =   &H0080C0FF&
            Caption         =   "By External ID"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   840
            Width           =   1305
         End
         Begin VB.OptionButton optByInternalId 
            BackColor       =   &H0080C0FF&
            Caption         =   "By Internal ID"
            Height          =   195
            Left            =   150
            TabIndex        =   32
            Top             =   570
            Width           =   1275
         End
         Begin VB.CheckBox chkWithEnctrs 
            BackColor       =   &H0080C0FF&
            Caption         =   "w\Encounters"
            Height          =   285
            Left            =   1680
            TabIndex        =   31
            Top             =   540
            Width           =   1335
         End
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   3195
      End
      Begin VB.Frame fraPatientSel 
         BackColor       =   &H0080C0FF&
         Height          =   675
         Left            =   3360
         TabIndex        =   24
         Top             =   6780
         Width           =   1935
         Begin VB.CommandButton cmdSelCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1020
            TabIndex        =   26
            Top             =   180
            Width           =   765
         End
         Begin VB.CommandButton cmdSelOK 
            Caption         =   "OK"
            Height          =   375
            Left            =   90
            TabIndex        =   25
            Top             =   180
            Width           =   795
         End
      End
      Begin VB.ListBox lstEncounters 
         Height          =   4740
         Left            =   4620
         TabIndex        =   22
         Top             =   1830
         Width           =   3765
      End
      Begin VB.ListBox lstPatient 
         Height          =   5520
         Left            =   120
         TabIndex        =   21
         Top             =   1050
         Width           =   4305
      End
      Begin MSComctlLib.StatusBar sbPatient 
         Height          =   345
         Left            =   0
         TabIndex        =   38
         Top             =   7500
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5468
               MinWidth        =   5468
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5468
               MinWidth        =   5468
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCount 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3420
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl_Count 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Count"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3420
         TabIndex        =   36
         Top             =   150
         Width           =   420
      End
      Begin VB.Label lblHeaders 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Last Name, First Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   780
         Width           =   4275
      End
      Begin VB.Label lblSelEncounter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encounters"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4650
         TabIndex        =   29
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Search:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame frPatientInfo 
      Height          =   465
      Left            =   150
      TabIndex        =   12
      Top             =   30
      Width           =   12735
      Begin VB.Label lbl_EncounterDate 
         AutoSize        =   -1  'True
         Caption         =   "Encounter: "
         Height          =   195
         Left            =   8160
         TabIndex        =   17
         Top             =   180
         Width           =   825
      End
      Begin VB.Label LblEncounterDate 
         AutoSize        =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   9060
         TabIndex        =   18
         Top             =   180
         Width           =   45
      End
      Begin VB.Label LblPatient 
         AutoSize        =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4350
         TabIndex        =   16
         Top             =   180
         Width           =   45
      End
      Begin VB.Label lbl_PatientName 
         AutoSize        =   -1  'True
         Caption         =   "Patient: "
         Height          =   195
         Left            =   3750
         TabIndex        =   15
         Top             =   180
         Width           =   585
      End
      Begin VB.Label LblPatientId 
         AutoSize        =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   930
         TabIndex        =   14
         Top             =   180
         Width           =   45
      End
      Begin VB.Label lbl_PatientID 
         AutoSize        =   -1  'True
         Caption         =   "Patient ID: "
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "..."
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   9120
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Frame fraControlPanel 
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7365
      Left            =   10320
      TabIndex        =   1
      Top             =   600
      Width           =   2580
      Begin VB.Frame fraClinical 
         Caption         =   "Clinical Data"
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
         Height          =   855
         Left            =   150
         TabIndex        =   42
         Top             =   420
         Width           =   2265
         Begin VB.CommandButton cmdRestore 
            Caption         =   "Res&tore"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   44
            Top             =   300
            Width           =   957
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   300
            Width           =   957
         End
      End
      Begin VB.Frame fraSavedEncounters 
         Caption         =   "Encounters"
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
         Height          =   855
         Left            =   120
         TabIndex        =   39
         Top             =   1380
         Width           =   2265
         Begin VB.CommandButton cmdNext 
            Caption         =   "Next >> "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1170
            TabIndex        =   41
            Top             =   330
            Width           =   957
         End
         Begin VB.CommandButton cmdPrev 
            Caption         =   "<< Prev"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   150
            TabIndex        =   40
            Top             =   330
            Width           =   957
         End
      End
      Begin VB.CommandButton cmdPatient 
         Caption         =   "Select Patient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   720
         TabIndex        =   23
         Top             =   3030
         Width           =   1365
      End
      Begin VB.ListBox lstRelStatements 
         Height          =   2790
         Left            =   150
         TabIndex        =   19
         Top             =   3960
         Width           =   2325
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         TabIndex        =   3
         Top             =   6840
         Width           =   957
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1380
         TabIndex        =   2
         Top             =   7590
         Width           =   957
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   210
         TabIndex        =   4
         Top             =   7650
         Width           =   957
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   5
         Top             =   8130
         Width           =   957
      End
      Begin VB.CommandButton cmdHistory 
         Caption         =   "&History..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   2490
         Width           =   1350
      End
      Begin VB.CommandButton cmdNotes 
         Caption         =   "N&otes..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   8130
         Width           =   957
      End
      Begin VB.Label lblRelStatements 
         BackStyle       =   0  'Transparent
         Caption         =   "Related Activities:"
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
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   3675
         Width           =   1755
      End
   End
   Begin TabDlg.SSTab tabBasic 
      Height          =   8280
      Left            =   120
      TabIndex        =   0
      Top             =   570
      Visible         =   0   'False
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   14605
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Basic.frx":094A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Label lblReportDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4470
      TabIndex        =   11
      Top             =   9180
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      Caption         =   "Report"
      Height          =   195
      Left            =   6690
      TabIndex        =   10
      Top             =   9180
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmBasic"
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
Private lintComponentCounter As Integer
Private lintTabCounter As Integer
Private lbooMissingClass As Boolean
Private lstrClassName As String
Private lstrLibrary As String
Private mlngCurrentPatientID As Long
Private mlngCurrentEncounterID As Long

'dnn.................
Private mintTabs As Long
Private mintTabsPerRow As Long
Private fstrClinicalConnect As String


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
'Private mObjSelectPatient As PatientTracTypeLib.iClinical
'Private WithEvents PatientList As SelectPatient
'
'Private Sub lctlComponent_ObjectEvent(Info As EventInfo)
'    Dim i As Long
'    Dim numargs As Long
'    Dim msg As String
'
'    msg = ""
'
'    numargs = Info.EventParameters.Count
'
'    msg = msg & "Incoming event: " & Info.Name & vbCrLf
'    For i = 0 To numargs - 1
'        msg = msg & "Arg n. " & CStr(i) & " name = " & _
'                CStr(Info.EventParameters(i).Name & _
'                " value = " & CStr(Info.EventParameters(i).Value)) & _
'                vbCrLf
'    Next i
'
'    MsgBox msg, vbInformation, "Incoming event!"
'
'End Sub

Public Function RunReport(vsrtReportFile As String) As Boolean
    pstrReportScript = vsrtReportFile
    'piPatientID = CLng(txtPatient)
    'piEncounterID = CLng(txtEncounter)
    
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
 '   frmreport.Show vbModal
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

Private Sub chkWithEnctrs_Click()
    Reload_Patients
End Sub

Private Sub cmdClose_Click()
    Dim liclComponent As PatientTracTypeLib.iClinical
    Dim lintResp As Integer
    Dim booSkipSave As Boolean
    
    If Not OpenEncounter(piPatientID, piEncounterID) And pbooLockClosedEncounterData Then
        booSkipSave = True
    End If
    
    If Not booSkipSave Then
        For Each liclComponent In mcolActxComp
    
            If liclComponent.HasChanged Then
                If lintResp <> vbYes Then
                    lintResp = MsgBox("Data has changed." & vbNewLine & "Save before closing?", vbInformation + vbYesNo, "Save ComponentID " & liclComponent.ComponentID)
                End If
    
                If lintResp = vbYes Then liclComponent.WriteClinicalData
    
            End If
    
        Next
    End If
    
    frmMDI.cboFormList.ListIndex = -1
    Unload Me
    
End Sub


Private Sub cmdHistory_Click()
    frmViewer.Show vbModal
End Sub

Private Sub cmdNew_Click()
    Unload Me
    frmPatients.Show vbModal
End Sub

Private Sub cmdNext_Click()
Dim lintIndex As Integer
    Dim liclComponent As PatientTracTypeLib.iClinical
    Dim lintResp As Integer
    Dim booSkipSave As Boolean
    
    
    If Not OpenEncounter(piPatientID, piEncounterID) Then
        'Closed Encounter...............
        If pbooLockClosedEncounterData Then
            Screen.MousePointer = vbDefault
            booSkipSave = True
            Exit Sub
        End If
    End If
    
    If Not booSkipSave Then
        For Each liclComponent In mcolActxComp
    
            If liclComponent.HasChanged Then
                lintResp = MsgBox("Do You Want To Save Before" & vbNewLine & "Looking At The Next Encounter?", vbInformation + vbYesNo, "Save ComponentID ")
                'If lintResp = vbYes Then cmdSave_Click
                If lintResp = vbYes Then liclComponent.WriteClinicalData
            End If
        Next
    End If
   
    Screen.MousePointer = vbHourglass
    
    lintIndex = lstEncounters.ListIndex
    
    If lintIndex > 0 Then
        lstEncounters.ListIndex = lintIndex - 1
        'cmdNext.Enabled = True
    End If
    

    Load_EncounterData
    
    If lstEncounters.ListIndex = 0 Then
        'If lstEncounters.ListCount > 1 Then
            If lstEncounters.ListCount > 1 Then cmdPrev.Enabled = True
        'Else
            cmdNext.Enabled = False
        'End If
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
        If lstEncounters.ListIndex < lstEncounters.ListCount - 1 Then cmdNext.Enabled = True
    End If
    
    msEncounterDate = Encounter_DateRange
    LblEncounterDate.Caption = Trim(msEncounterDate)
    
    mlngEncounterID = lstEncounters.ItemData(lstEncounters.ListIndex)
    piEncounterID = lstEncounters.ItemData(lstEncounters.ListIndex)
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdPatient_Click()
    Screen.MousePointer = vbHourglass
    picPatientSel.Visible = True
    mlngCurrentPatientID = mlngPatientID
    mlngCurrentEncounterID = mlngEncounterID
    fraControlPanel.Enabled = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrev_Click()
    Dim lintIndex As Integer
    Dim liclComponent As PatientTracTypeLib.iClinical
    Dim lintResp As Integer
    Dim booSkipSave As Boolean
    
    If Not OpenEncounter(piPatientID, piEncounterID) Then
        'Closed Encounter...............
        If pbooLockClosedEncounterData Then
            Screen.MousePointer = vbDefault
            booSkipSave = True
            Exit Sub
        End If
    End If
    
    If Not booSkipSave Then
        For Each liclComponent In mcolActxComp
    
            If liclComponent.HasChanged Then
                lintResp = MsgBox("Do You Want To Save Before" & vbNewLine & "Looking At The Previous Encounter?", vbInformation + vbYesNo, "Save ComponentID ")
                'If lintResp = vbYes Then cmdSave_Click
                If lintResp = vbYes Then liclComponent.WriteClinicalData
            End If
    
        Next
    End If
 
    Screen.MousePointer = vbHourglass
    
    lintIndex = lstEncounters.ListIndex
    
    If lintIndex < lstEncounters.ListCount - 1 Then
        lstEncounters.ListIndex = lintIndex + 1
        cmdPrev.Enabled = True
    End If
    
    Load_EncounterData
    
    If lstEncounters.ListIndex = lstEncounters.ListCount - 1 Then
        cmdPrev.Enabled = False
        If lstEncounters.ListCount > 1 Then cmdNext.Enabled = True
    Else
        cmdNext.Enabled = True
    End If
    
    msEncounterDate = Encounter_DateRange
    LblEncounterDate.Caption = Trim(msEncounterDate)
    
    mlngEncounterID = lstEncounters.ItemData(lstEncounters.ListIndex)
    piEncounterID = lstEncounters.ItemData(lstEncounters.ListIndex)
    
    'Unload Me
    'Controls.Remove tabdlg.SSTab
    'tabBasic.Visible = False
    'Form_Activate
    'Me.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdRestore_Click()
    cmdSelOK_Click
End Sub


Private Sub cmdSave_Click()

    Dim liclComponent As PatientTracTypeLib.iClinical
    Dim lintResp As Integer
    
    Screen.MousePointer = vbHourglass
    
    If Not OpenEncounter(piPatientID, piEncounterID) Then
        'Closed Encounter...............
        If pbooLockClosedEncounterData Then
            MsgBox "Closed Encounter.  Cannot save changed data.", vbInformation, "Closed Encounter"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    'Cycle through all loaded components
    For Each liclComponent In mcolActxComp
        'Test if the componet indicates to be dirty
        If liclComponent.HasChanged Then
            'Call the component's write routine
            liclComponent.WriteClinicalData
        'Else
        '    lintResp = MsgBox("Data has not changed." & vbNewLine & "Save anyway?", vbInformation + vbYesNo, "Save ComponentID " & liclComponent.ComponentID)
        '    LogFile "Save ComponentID " & liclComponent.ComponentID
        '    If lintResp = vbYes Then liclComponent.WriteClinicalData
            
        End If
    Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelCancel_Click()
Dim i As Integer

    If lstPatient.ItemData(lstPatient.ListIndex) <> mlngPatientID Then
        For i = 0 To lstPatient.ListCount - 1
            If lstPatient.ItemData(i) = mlngCurrentPatientID Then
                lstPatient.ListIndex = i
                Get_Encounters lstPatient.ItemData(i), mlngCurrentEncounterID
                Exit For
            End If
        Next i
    End If

    picPatientSel.Visible = False
    fraControlPanel.Enabled = True
End Sub

Private Sub cmdSelOK_Click()
Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long

    If lstEncounters.Text = vbNullString Then
        MsgBox "Select an encounter.", vbExclamation, "Patient Select"
        Exit Sub
    End If

    If lstPatient.ItemData(lstPatient.ListIndex) = mlngPatientID And lstEncounters.ItemData(lstEncounters.ListIndex) = mlngEncounterID Then
        Exit Sub
    End If
    
    fraControlPanel.Enabled = True
    piPatientID = lstPatient.ItemData(lstPatient.ListIndex)
    mlngPatientID = lstPatient.ItemData(lstPatient.ListIndex)
    mlngEncounterID = lstEncounters.ItemData(lstEncounters.ListIndex)
    piEncounterID = lstEncounters.ItemData(lstEncounters.ListIndex)
    picPatientSel.Visible = False
    
    frmMDI.StatusBar.Panels(1).Text = "Patient: " & lstPatient.Text
    frmMDI.StatusBar.Panels(2).Text = "External ID: " & pxPatientID
    frmMDI.StatusBar.Panels(3).Text = "Internal ID: " & piPatientID
    frmMDI.StatusBar.Panels(4).Text = "Encounter: " & piEncounterID
    
    psPatientName = lstPatient.Text
    msEncounterDate = Encounter_DateRange
    
    SetPatientInfo
    
    Unload Me
      
      'frmBasic.FormID = lFormID
      frmBasic.PatientID = piPatientID
      frmBasic.EncounterID = piEncounterID
      frmBasic.ProviderID = piProviderID
                
      frmBasic.Show
    
     
End Sub


Private Function Encounter_DateRange() As String
Dim strSQL As String
Dim varRS
Dim lngRows As Long
Dim bRet As Boolean
Dim strStartDate As String
Dim strEndDate As String
Dim lstrRemarks As String

LogFile "Private Function Encounter_DateRange-----------------------------------"

    strSQL = "Select ENCNTR_START_DATE, ENCNTR_END_DATE " _
        & "From ENCOUNTER " _
        & "Where ENCOUNTER_ID = " & mlngEncounterID

    bRet = adoSelectQuery(strSQL, fstrClinicalConnect, varRS, , lngRows)
    
    LogFile Me.Name & ":Encounter_DateRange: " & vbNewLine & strSQL
    
    If lngRows > 0 Then
        If Not IsNull(varRS(0, 0)) Then strStartDate = Format(varRS(0, 0), "mm/dd/yyyy hh:nn AMPM")
        If Not IsNull(varRS(1, 0)) Then strEndDate = Format(varRS(1, 0), "mm/dd/yyyy hh:nn AMPM")
        Encounter_DateRange = strStartDate & " - " & strEndDate
        'fstrEncounterStartDate = strStartDate
        'fstrEncounterEndDate = strEndDate
    End If

End Function

Private Sub Form_Activate()

    'Visible property of the basic tab is used as a flag to indicate if
    'the initialization of the page has executed.
    
    'Me.Visible = False
   
    LogFile "Start Log  Form Basic --------------------", True
    LogFile "Form:" & frmMDI.cboFormList.Text
    
    fstrClinicalConnect = psConnect(0)
    
    If tabBasic.Visible = False Then
    
        'Initialization needs to take place.
        DoEvents
        
        Screen.MousePointer = vbHourglass
        'tabBasic.Visible = True  'dnn
        tabBasic.Visible = False  'dnn
        fraControlPanel.Visible = False
        
        DistributeAllComponents
        
        'Set initial state of basic tab
        tabBasic.Tab = 0
        tabBasic.TabIndex = 0
        
        'Show the initialized tab
        '(and prevent from initializing again)
        Me.Width = tabBasic.Width + fraControlPanel.Width + 900
        
'        If tabBasic.Height > fraControlPanel.Height Then
'            Me.Height = tabBasic.Height + 1000
'        End If
'
        If tabBasic.Height > fraControlPanel.Height Then
            Me.Height = tabBasic.Height + frPatientInfo.Height + 1000
        End If

        tabBasic.Top = fraControlPanel.Top
        fraControlPanel.Left = tabBasic.Left + tabBasic.Width + 200
        
        centerForm Me
        'Me.Visible = True
        tabBasic.Visible = True
        fraControlPanel.Visible = True
        
        tabBasic.SetFocus
        
        msEncounterDate = Encounter_DateRange
        SetPatientInfo
        Load_Patients
        Get_Encounters piPatientID, piEncounterID
        CenterControl picPatientSel, Me
        Screen.MousePointer = vbDefault
    End If

   
    
    
End Sub

Private Sub Load_Patients()
Dim strSQL As String
Dim lstrRemarks As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim r As Integer
Dim c As Integer

LogFile "Private Sub Load_Patients-----------------------------------"

    Set vRS = Nothing
    strSQL = "Select patient_id, rtrim(last_name) last_name, rtrim(first_name) first_name " _
        & " From Patient " _
        & "Order by rtrim(last_name),rtrim(first_name)"
    
    'iRows = Run_SelectQuery(psDB, strSQL, vData)
    bRet = adoSelectQuery(strSQL, psConnect(0), vRS, , iRows)

    LogFile Me.Name & ":Load_Patients: " & vbNewLine & strSQL
    
If iRows > 0 Then
    For r = LBound(vRS, 2) To UBound(vRS, 2)
    
        lstPatient.AddItem vRS(1, r) & ", " & vRS(2, r)
        If IsNumeric(vRS(0, r)) Then lstPatient.ItemData(r) = vRS(0, r)
        If vRS(0, r) = piPatientID Then lstPatient.ListIndex = lstPatient.NewIndex
    Next r
End If

End Sub


Private Sub Get_Encounters(vlngPatientID As Long, Optional vlngEncounterID As Long)
Dim lstrSQL As String
Dim i As Integer
Dim lstrRemarks As String

LogFile "Private Sub Get_Encounters-----------------------------------"

If piPatientID = 0 Then Exit Sub

    lstrSQL = "Select (cast(encntr_start_date as varchar)  + ' - ' + case when encntr_end_date is null then '' else cast(encntr_end_date as varchar) end) as dur, encounter_id " _
        & "From ENCOUNTER " _
        & "Where patient_id = " & vlngPatientID & " Order By ENCNTR_START_DATE desc"
        
        LogFile Me.Name & ":Get_Encounters: " & vbNewLine & lstrSQL
        
        Load_ComboBox lstEncounters, lstrSQL, psConnect(0)
        
        If vlngEncounterID > 0 Then
            For i = 0 To lstEncounters.ListCount - 1
                If lstEncounters.ItemData(i) = vlngEncounterID Then
                    lstEncounters.ListIndex = i  'lstEncounters.NewIndex
                    Exit For
                End If
            Next i
            
            If lstEncounters.ListIndex = 0 Then cmdNext.Enabled = False
            If lstEncounters.ListIndex = lstEncounters.ListCount - 1 Then cmdPrev.Enabled = False
            
        End If
        
        
End Sub
Private Sub Form_Load()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    mlngLastPatientId = 0
    mlngLastEncounterId = 0
    centerForm Me
    
    'Visible property of the basic tab is used as a flag to indicate if
    'the initialization of the page has executed.
    'Make sure is set to false upon load to tell the activate routine
    'that the tab pages need to be initialized.
    tabBasic.Visible = False
    
    'Load Related Activities ListBox
    
    'Open connection to database
    cn.Open psConnect(1)
    
    cn.S_RELATED_FORMS "S", (mlngFormID), rs

    LogFile Me.Name & ":Form_Load: S_RELATED_FORMS 'S'," & CStr(mlngFormID)
    
    With rs
       If Not .EOF Then
          Do Until .EOF
             lstRelStatements.AddItem (!FORM_CAPTION)
             lstRelStatements.ItemData(lstRelStatements.NewIndex) = (!CHILD_FORM_ID)
             .MoveNext
          Loop
          .Close
       End If
    End With
    
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
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


Private Sub DistributeAllComponents()

    On Error GoTo errDistributeAllComponents
    
    'Dimension data retrieval variables
    Dim lconTerms As New ADODB.Connection
    Dim lrstContainerDefinition As New ADODB.Recordset
    Dim lrstContainerChildren As New ADODB.Recordset
    Dim lstrRemarks As String
    Dim intResp As Byte
    
    LogFile "Private Sub DistributeAllComponents-----------------------------------"
    
    'dnn*..........
    mintTabs = 0
    mintTabsPerRow = 0
    mbooSelectPat = False
    
    'Open connection to database
    lconTerms.Open psConnect(1)
    
    'Call stored procedure with top level page definition passing form ID and
    'retrieving results into local recordset
    lconTerms.S_TAB_FORMS "R", (mlngFormID), lrstContainerDefinition
    
    LogFile Me.Name & ":DistributeAllComponents: S_TAB_FORMS 'R', " & CStr(mlngFormID)

    'Call routine to create children tabs for the form
    CreateBlankTabs Me, lrstContainerDefinition
    
    With lrstContainerDefinition

        If Not .EOF <> 0 Then
        
            'Set form caption to value from first row
            Caption = !FORM_CAPTION
            
            'dnn*..........................
            If Not IsNull((!CHILD_TABS)) Then mintTabs = (!CHILD_TABS)
            If Not IsNull((!CHILD_TABS_PER_ROW)) Then mintTabsPerRow = (!CHILD_TABS_PER_ROW)
            
            'Cycle through recordset populating tabs' contents
            Do Until .EOF
                
                tabBasic.Tab = (!TAB_SEQ) - 1
                lstrRemarks = tabBasic.Tab
                LogFile Me.Name & ":DistributeAllComponents:TabDrillDown (!TAB_ID), lconTerms, tabBasic >> TAB_ID = " & CStr((!TAB_ID))
                TabDrillDown (!TAB_ID), lconTerms, tabBasic
                DoEvents
                .MoveNext
    
            Loop
        
        End If

        .Close

    End With

    'Terminate db references
    Set lrstContainerDefinition = Nothing
    
    lconTerms.Close
    Set lconTerms = Nothing
    
    Exit Sub
    
errDistributeAllComponents:

    If InStr(Err.Description, "network error") > 0 Then
        intResp = MsgBox(Err.Description & vbNewLine & "Tyr Again?", vbYesNo + vbCritical, "Error #" & Err.Number)
        If intResp = vbYes Then Resume Else Screen.MousePointer = vbDefault: Exit Sub
    End If
    
     ErrHandler _
     Me.Name & " DistributeAllComponents", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub


Private Sub RefreshComponents()
Dim lstrRemarks As String

LogFile "Private Sub RefreshComponents-----------------------------------"

On Error Resume Next

  'Dimension reference for component control to add
    Dim liclComponent As PatientTracTypeLib.iClinical
      
  'Cycle through all loaded components
  For Each liclComponent In mcolActxComp
    LogFile "RefreshComponents: actx_id = " & liclComponent.ComponentID
    Debug.Print liclComponent.ComponentID
    'Establish component clinical properties
    liclComponent.ExternalUserCall
    liclComponent.PatientID = mlngPatientID
    liclComponent.EncounterID = mlngEncounterID
    liclComponent.ProviderID = mlngProviderID
    liclComponent.ConnectStringTerms = psConnect(1)
    liclComponent.ConnectStringClinical = psConnect(0)
    liclComponent.ReadClinicalData
  Next
  
End Sub


Private Sub CreateBlankTabs _
 ( _
 ByRef robjContainer As Object, _
 ByRef rrstTabLayout As ADODB.Recordset _
 )

    'Dimension component and tab counters
    Dim lintTabCounter As Integer
    Dim ltabTabControl As Control
    Dim lstrRemarks As String
    
    LogFile "Private Sub CreateBlankTabs-----------------------------------"
    
    If TypeOf robjContainer Is Form Then
    
        'If the container is a form, it must be the basic form.
        'Create reference to top level tab created at design time.
        Set ltabTabControl = tabBasic
        ltabTabControl.Tabs = rrstTabLayout.Fields("TABS").Value
        ltabTabControl.TabsPerRow = rrstTabLayout.Fields("TABS_PER_ROW").Value
        ltabTabControl.Top = rrstTabLayout!TAB_POS_Y + 150
        ltabTabControl.Left = rrstTabLayout!TAB_POS_X
'        ltabTabControl.Left = rrstTabLayout!TAB_POS_X + 200
        ltabTabControl.Width = rrstTabLayout!TAB_WIDTH
        ltabTabControl.Height = rrstTabLayout!TAB_HEIGHT
        ltabTabControl.Font.Name = rrstTabLayout!TAB_FONT_NAME
        ltabTabControl.Font.Size = rrstTabLayout!TAB_FONT_SIZE
        ltabTabControl.ForeColor = rrstTabLayout!TAB_FONT_COLOR
        ltabTabControl.BackColor = rrstTabLayout!TAB_BACK_COLOR
        
        If rrstTabLayout!TAB_FONT_STYLE = "Bold" Then
            ltabTabControl.Font.Bold = True
        Else
            ltabTabControl.Font.Bold = False
        End If
        
        If rrstTabLayout!TAB_FONT_STYLE = "Italic" Then
            ltabTabControl.Font.Italic = True
        Else
            ltabTabControl.Font.Italic = False
        End If
        
        robjContainer.Width = fraControlPanel.Width + 500 + ltabTabControl.Width
        robjContainer.Height = ltabTabControl.Height + 1800
        centerForm Me
    
    Else
    
        'If the container is a tab, we need to create a new tab within
        Set ltabTabControl = Controls.Add("tabdlg.SStab", _
         robjContainer.Name & "_" & Format$(rrstTabLayout!TAB_ID, "0000"), robjContainer)
        
        ltabTabControl.Tabs = rrstTabLayout.Fields("TABS").Value
        ltabTabControl.TabsPerRow = rrstTabLayout.Fields("TABS_PER_ROW").Value
        ltabTabControl.Top = 600
        ltabTabControl.Left = 300
        ltabTabControl.Height = robjContainer.Height - 750
        ltabTabControl.Width = robjContainer.Width - 600
        ltabTabControl.Visible = True
        
        'dnn......................................
        'If mintTabs > 0 Then
        '    ltabTabControl.Tabs = mintTabs
        'End If
        
        'If mintTabsPerRow > 0 Then
        '    ltabTabControl.TabsPerRow = mintTabsPerRow
        'End If
        
    End If
    
    With rrstTabLayout
    
        'Cycle through recordset populating tab control with page definition
        Do Until .EOF
        
            lintTabCounter = lintTabCounter + 1
            
            'Keep track of current tab
            'lintTabCounter = (!TAB_SEQ) - 1
            
            'Add tab page to control
            ltabTabControl.Tabs = lintTabCounter
            
            'Activate destination tab
            ltabTabControl.Tab = lintTabCounter - 1
            DoEvents
                
            'Set page caption according to data value
            ltabTabControl.Caption = (!TAB_CAPTION)
            If (UCase(ltabTabControl.Caption) = "SELECT PATIENT") Then
               mbooSelectPat = True
            End If
            
            DoEvents
            .MoveNext
            
        Loop
        
        'If lintTabCounter <> 0 Then
            'dnn*....................
            'ltabTabControl.TabsPerRow = lintTabCounter
        'End If
        
        'Since we are using a forward only recordset, refresh it to leave it
        'the way it was found.
        .Requery
        
        '*** Note!   We could be changing this for a recordset that allows
        '            MoveFirst to avoid re-querying.
        
    End With
    
    'Set current tab to first tab
    ltabTabControl.Tab = 0

End Sub

Private Sub TabDrillDown _
( _
 ByVal vlngTabID As Long, _
 ByRef rconTerms As ADODB.Connection, _
 ByRef rtabContainer As Control _
 )

    On Error GoTo errTabDrillDown
    
    Dim lrstTabChildren As New ADODB.Recordset
    Dim lrstTabDefinition As New ADODB.Recordset
    Dim lctlChildTab As Control
    Dim lstrRemarks As String
    Dim intResp As Byte
    
    LogFile "Private Sub TabDrillDown-----------------------------------"
    
    'Call stored procedure to find if tab contains more tabs within
    rconTerms.S_TAB_PARENT_CHILD "R", (vlngTabID), lrstTabDefinition
    LogFile "TabDrillDown:S_TAB_PARENT_CHILD 'R', " & CStr(vlngTabID)
    
    With lrstTabDefinition
    
        'Test for no more children tabs
        If IsNull(!TAB_CHILD_ID) Then
            lstrRemarks = "TabID = " & vlngTabID
            'Load components within tab
            TabPopulate vlngTabID, rconTerms, rtabContainer
        
        Else
            
            'Call routine to create blank tabs
             CreateBlankTabs rtabContainer, lrstTabDefinition
        
            'For all children tabs
            Do Until .EOF
            
                Set lctlChildTab = Controls(rtabContainer.Name & "_" & Format$(vlngTabID, "0000"))
                lctlChildTab.Tab = (!TAB_SEQ) - 1
                
                'Recurse to find if children have tabs within
                TabDrillDown (!TAB_CHILD_ID), rconTerms, lctlChildTab
                
                lctlChildTab.Tab = 0
                
                DoEvents
                .MoveNext
                
                '*** Note!  This assumes the database contrains prevent from
                '           defining an infinite relationship of parent and
                '           children tabs.   To make this code more solid we
                '           need to add a limiting counter that will trap and
                '           avoid and infinite recursion.
                
            Loop

        End If
        
    End With
    
    Exit Sub
    
errTabDrillDown:

    If InStr(Err.Description, "network error") > 0 Then
        intResp = MsgBox(Err.Description & vbNewLine & "Tyr Again?", vbYesNo + vbCritical, "Error #" & Err.Number)
        If intResp = vbYes Then Resume Else Screen.MousePointer = vbDefault: Exit Sub
    End If

     ErrHandler _
     Me.Name & " TabDrillDown", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub

Private Sub TabPopulate _
 ( _
 ByVal vlngTabID As Long, _
 ByRef rconTerms As ADODB.Connection, _
 ByRef rtabContainer As Control _
 )
 
    On Error GoTo errTabPopulate
  
    'Dimension component and tab counters
    Dim lintComponentCounter As Integer
    Dim lintTabCounter As Integer
    Dim lbooMissingClass As Boolean
    Dim lstrClassName As String
    Dim lstrLibrary As String
    Dim lstrRemarks As String
    Dim i As Integer
    Dim intResp As Byte
    
    'Dimension reference for component control to add
    Dim lctlComponent As VBControlExtender
    Dim liclComponent As PatientTracTypeLib.iClinical
    
    'Dimension recordset for definition of tab layout
    Dim lrstComponents As New ADODB.Recordset
    
    LogFile "Private Sub TabPopulate-----------------------------------"
    
    'Call stored procedure with tab definition passing tab ID and
    'retrieving results into local recordset
    rconTerms.S_TAB_TAB_ACTIVE_X "R", (vlngTabID), lrstComponents
    LogFile "TabPopulate:S_TAB_TAB_ACTIVE_X 'R', " & CStr(vlngTabID)
    
    With lrstComponents

        If Not IsNull(!ACTX_SEQ) Then
            
            'Cycle through recordset populating tab control with page definition
            Do Until .EOF
            
                lstrClassName = (!ACTX_CLASS)
                lstrLibrary = (!ACTX_NAME)
                lstrRemarks = lstrLibrary & "." & lstrClassName
                'i = 1 / 0
                LogFile Me.Name & ":TabPopulate: " & lstrLibrary & "." & lstrClassName
                
                'Licenses.Add lstrLibrary & "." & lstrClassName
                
                'Establish reference to component control as we add it...
                Set lctlComponent = Controls.Add(lstrLibrary & "." & lstrClassName, _
                 Format$(Controls.Count + 1, "\c\t\l000"), rtabContainer)
    
                'Test if class was not found in library
'                If lbooMissingClass = True Then
'
'                    'Load control place holder with unrecognized class...
'                    Set lctlComponent = Controls.Add("BusinessGUI.Unrecognized", _
'                     Format$(Controls.Count + 1, "\c\t\l000"), rtabContainer)
'                     lbooMissingClass = False
'
'                    'Add missing class name to caption
'                    lctlComponent.ClassName = lstrClassName
'
'                End If
                
                'Place and show component
                If Not lbooMissingClass Then  'dnn 11/20/2003
                    lctlComponent.Top = (!ACTX_POS_Y)
                    lctlComponent.Left = (!ACTX_POS_X)
                    lctlComponent.TabIndex = lintComponentCounter
                    lctlComponent.Visible = True
                    
                    'dnn...................
                    If Not IsNull((!CHILD_TABS)) Then mintTabs = (!CHILD_TABS)
                    If Not IsNull((!CHILD_TABS_PER_ROW)) Then mintTabs = (!CHILD_TABS_PER_ROW)
                    
                'Test if the component is not the place holder for a missing class
                'If lbooMissingClass = False Then
                    
                    'Create reference to clinical interface of the component
                    Set liclComponent = lctlComponent.Object
                    mlngPatientID = piPatientID
                    mlngEncounterID = piEncounterID
                    mlngProviderID = piProviderID
                    'Establish component clinical properties
                    
                    
                    'liclComponent.ConnectStringTerms = psConnect(1)
                    'liclComponent.ConnectStringClinical = psConnect(0)
                    liclComponent.PatientID = mlngPatientID
                    liclComponent.EncounterID = mlngEncounterID
                    liclComponent.ProviderID = mlngProviderID
                    liclComponent.ComponentID = (!actx_id)
                    
                    If (!OBJ_POOL_ID) > 0 Then  'Pool ID  dnn
                        liclComponent.ObjPoolID = (!OBJ_POOL_ID)
                    End If

                    liclComponent.ConnectStringTerms = psConnect(1)
                    liclComponent.ConnectStringClinical = psConnect(0)
                    
                    'xxxx
'                    If mbooSelectPat And (UCase(lstrClassName) = "SELECTPATIENT") Then
'                      If (mObjSelectPatient Is Nothing) Then
'                         Set mObjSelectPatient = liclComponent
'                         Set PatientList = lctlComponent
'                      End If
'                    End If
                    
                    If (!ACTX_INTRV_CD) <> "" Then liclComponent.CPTCode = (!ACTX_INTRV_CD)
                    
                    'Force component to retrieve any existing clinical data
                    liclComponent.ReadClinicalData
                    
                    'Keep a reference of loaded clinical component
                     mcolActxComp.Add liclComponent, Format$(CStr((vlngTabID * 1000) + liclComponent.ComponentID), "\i\d\_00000") & (!ACTX_INTRV_CD)
                
                End If
                
                'Increment component counter and continue loop
                lintComponentCounter = lintComponentCounter + 1
                
                DoEvents
                .MoveNext
                
            Loop
        
        End If
        
        .Close
    
    End With
    
    Exit Sub

errTabPopulate:
    'Or Err.Number = 429
    If InStr(Err.Description, "network error") > 0 Then
        intResp = MsgBox(Err.Description & vbNewLine & "Tyr Again?", vbYesNo + vbCritical, "Error #" & Err.Number)
        If intResp = vbYes Then Resume Else Screen.MousePointer = vbDefault: Exit Sub
    End If
    
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 734 Then Resume Next

    'Test if error is related to a class from the database not founnd in the ocx
    If Err.Number = 711 Then
        'Set flag and resume
        lbooMissingClass = True
        
        MsgBox lstrLibrary & "." & lstrClassName & " component missing.  Add to ActiveX collection using EMR Toolbox.", vbCritical
        Resume Next
    
    'Test if error is related to a missing license
    ElseIf Err = 731 Then
    
        'It must be the FP Spread... add license
        'Hard-coded for now until a data driven scheme is devised
        Licenses.Add lstrLibrary & "." & lstrClassName
        Resume
    
    ElseIf Err = 747 Then
    
        'It must be our license... add license
        'Hard-coded for now until a data driven scheme is devised
        Licenses.Add lstrLibrary & "." & lstrClassName
        If Err.Number = 732 Or Err.Number = 733 Then Resume Next
  
        Resume
   
    Else
        'Until we have a solid error handling scheme, defer the error to previous handler
        'Err.Raise Err.Number
    End If
     
     ErrHandler _
     Me.Name & " TabPopulate", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume

End Sub

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
    
End Sub


Private Sub lstPatient_Click()
Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long

    Get_Encounters lstPatient.ItemData(lstPatient.ListIndex), piEncounterID
    
    strSQL = "Select PATIENT_EXT_REF " _
        & " From Patient " _
        & "Where patient_id = " & lstPatient.ItemData(lstPatient.ListIndex)
        
    bRet = adoSelectQuery(strSQL, psConnect(0), vRS, , iRows)

    If iRows > 0 Then
        If IsNull(vRS(0, 0)) Then
            pxPatientID = "(NULL)"
        Else
            pxPatientID = vRS(0, 0)
        End If
    End If
    
  sbPatient.Panels(1).Text = "External Patient ID: " & pxPatientID
  sbPatient.Panels(2).Text = "Internal Patient ID: " & lstPatient.ItemData(lstPatient.ListIndex)
End Sub

Private Sub lstRelStatements_DBlClick()
   Dim lFormID As Long
   
   If lstRelStatements.ListIndex > -1 Then
      lFormID = lstRelStatements.ItemData(lstRelStatements.ListIndex)
      
      Unload Me
      
      frmBasic.FormID = lFormID
      frmBasic.PatientID = piPatientID
      frmBasic.EncounterID = piEncounterID
      frmBasic.ProviderID = piProviderID
                
      frmBasic.Show
   End If
End Sub

Private Sub optByExternalId_Click()
    lblHeaders = "Last Name, First Name, (External ID)"
    Reload_Patients
End Sub

Private Sub optByInternalId_Click()
    lblHeaders = "Last Name, First Name, (Internal ID)"
    Reload_Patients
End Sub

Private Sub optByPatientName_Click()
    lblHeaders = "Last Name, First Name"
    Reload_Patients
End Sub

Private Sub Reload_Patients()
Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim i As Integer
Dim lsrtChar As String
Dim llngInternalID As Long
Dim lstrInternalID As String
Dim lstrRemarks As String

LogFile "Private Sub Reload_Patients-----------------------------------"

Screen.MousePointer = vbHourglass
    txtSearch = vbNullString
    
    If chkWithEnctrs.Value Then
        
        If optByPatientName.Value Then
            strSQL = "Select rtrim(last_name) + ', ' + rtrim(first_name) , patient_id " _
                & "From PATIENT " _
                & "Where patient_id in (select patient_id From ENCOUNTER) " _
                & "Order by rtrim(last_name),rtrim(first_name)"
            
        ElseIf optByInternalId.Value Then
        
            strSQL = "Select rtrim(last_name) + ', ' + rtrim(first_name) + ' (' + rtrim(cast(patient_id as char)) + ')', patient_id " _
                & "From PATIENT " _
                & "Where patient_id in (select patient_id From ENCOUNTER) " _
                & "Order by rtrim(last_name),rtrim(first_name)"
        
        ElseIf optByExternalId.Value Then
        
            strSQL = "Select rtrim(last_name) + ', ' + rtrim(first_name) + ' (' + rtrim(patient_ext_ref) + ')', patient_id " _
                & "From PATIENT " _
                & "Where patient_id in (select patient_id From ENCOUNTER) " _
                & "Order by rtrim(last_name),rtrim(first_name)"
        End If
    Else
        If optByPatientName.Value Then
            strSQL = "Select rtrim(last_name) + ', ' + rtrim(first_name) , patient_id " _
                & "From PATIENT " _
                & "Order by rtrim(last_name),rtrim(first_name)"
            
        ElseIf optByInternalId.Value Then
        
            strSQL = "Select rtrim(last_name) + ', ' + rtrim(first_name) + ' (' + rtrim(cast(patient_id as char)) + ')', patient_id " _
                & "From PATIENT " _
                & "Order by rtrim(last_name),rtrim(first_name)"
        
        ElseIf optByExternalId.Value Then
        
            strSQL = "Select rtrim(last_name) + ', ' + rtrim(first_name) + ' (' + rtrim(patient_ext_ref) + ')', patient_id " _
                & "From PATIENT " _
                & "Order by rtrim(last_name),rtrim(first_name)"
        End If
    End If

    'Reload Combo box....................................................
    bRet = adoSelectQuery(strSQL, psConnect(0), vRS, , iRows)
    If iRows > 0 Then LoadCombo lstPatient, vRS
    
    lblCount.Caption = lstPatient.ListCount
    lblCount.Refresh
    
    
    'Parse Internal Patient ID from status bar..............................................
    If sbPatient.Panels(2).Text <> vbNullString Then
        While i < Len(sbPatient.Panels(2).Text)
            i = i + 1
            lsrtChar = Mid(sbPatient.Panels(2).Text, i, 1)
            
            If IsNumeric(lsrtChar) Then
                lstrInternalID = lstrInternalID & lsrtChar
            End If
        Wend
    
    End If
    
    'Re-select patient.........................................
    If IsNumeric(lstrInternalID) Then
        llngInternalID = Val(lstrInternalID)
        
        For i = 0 To lstPatient.ListCount - 1
            If lstPatient.ItemData(i) = llngInternalID Then
                lstPatient.ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub tabBasic_Click(PreviousTab As Integer)
Dim iCurrTab As Integer


   iCurrTab = tabBasic.Tab
   If mbooRecurssive Or (Not tabBasic.Visible) Then
     Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   'xxx If (tabBasic.Tab <> 0) And Not (mObjSelectPatient Is Nothing) Then
   If (tabBasic.Tab <> 0) Then
     'piPatientID = mObjSelectPatient.PatientID
     'piEncounterID = mObjSelectPatient.EncounterID
     'piProviderID = mObjSelectPatient.ProviderID
     
     If (piPatientID > 0) And (piEncounterID > 0) And ((mlngLastPatientId <> piPatientID) Or (mlngLastEncounterId <> piEncounterID)) Then
       mbooRecurssive = True
       mlngLastPatientId = piPatientID
       mlngLastEncounterId = piEncounterID
       mlngPatientID = piPatientID
       mlngEncounterID = piEncounterID
       mlngProviderID = piProviderID
       RefreshComponents
       'SetPatientInfo
     End If
   End If
   
   If (tabBasic.Tab <> 0) And (mbooSelectPat) Then
      If piPatientID < 1 Then
        MsgBox "Select a Patient First!", vbInformation
        tabBasic.Tab = 0
        Exit Sub
      End If
                    
     If piEncounterID < 1 Then
       MsgBox "Select a Encounter First", vbInformation
       tabBasic.Tab = 0
       Exit Sub
     End If
   End If
   tabBasic.Tab = iCurrTab
   mbooRecurssive = False
   Screen.MousePointer = vbDefault
End Sub


Private Sub PatientList_PatientClick(ByVal APatientId As Long, ByVal APatientName, AEncounterDate As String)
'   piPatientID = APatientId
'   psPatientName = APatientName
'   msEncounterDate = AEncounterDate
'   mlngPatientID = piPatientID
'   SetPatientInfo
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


Private Sub SetPatientInfo()
Dim lstrRemarks As String

LogFile "Private Sub SetPatientInfo-----------------------------------"

'  LblPatientId.Caption = ""
'  If piPatientID > 0 Then LblPatientId.Caption = "Internal: " & piPatientID & ",  External: " & pxPatientID
'  LblPatient.Caption = Trim(psPatientName) & ", " & frmPatients.txtFirstName & " DOB: " & frmPatients.txtMonth & "/" & frmPatients.txtDay & "/" & frmPatients.txtYear & " "
'  LblEncounterDate.Caption = Trim(msEncounterDate)
'
  'frmMDI.StatusBar.Panels(1).Text = "Patient: " & psPatientName
  'frmMDI.StatusBar.Panels(2).Text = "External ID: " & pxPatientID
  'frmMDI.StatusBar.Panels(3).Text = "Internal ID: " & piPatientID
  'frmMDI.StatusBar.Panels(4).Text = "Encounter: " & piEncounterID
End Sub

Private Sub tabBasic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tabBasic.Top = fraControlPanel.Top
End Sub

Private Sub txtSearch_Change()
Dim k As Integer
Dim lRow As Long
Dim sItem As String

    If optByInternalId.Value Or optByExternalId.Value Then
        SearchList lstPatient, "(" & txtSearch.Text & ")", False
    Else
        SearchList lstPatient, txtSearch.Text, True
    End If

'sItem = UCase$(Trim$(txtSearch.Text))
'
'If sItem <> "" Then
'
'  For lRow = 0 To lstPatient.ListCount - 1
'
'   k = InStr(UCase$(lstPatient.List(lRow)), sItem)
'
'   If k = 1 Then
'    lstPatient.Selected(lRow) = True
'    Exit For
'   End If
'
'  Next lRow
'End If

  
End Sub



Private Sub Load_EncounterData()
Dim lstrRemarks As String

LogFile "Private Sub Load_EncounterData-----------------------------------"

    'piPatientID = lstPatient.ItemData(lstPatient.ListIndex)
    'mlngPatientID = lstPatient.ItemData(lstPatient.ListIndex)
    mlngEncounterID = lstEncounters.ItemData(lstEncounters.ListIndex)
    piEncounterID = lstEncounters.ItemData(lstEncounters.ListIndex)
    'picPatientSel.Visible = False
    
    'frmMDI.StatusBar.Panels(1).Text = "Patient: " & lstPatient.Text
    'frmMDI.StatusBar.Panels(2).Text = "External ID: " & pxPatientID
    'frmMDI.StatusBar.Panels(3).Text = "Internal ID: " & piPatientID
    frmMDI.StatusBar.Panels(4).Text = "Encounter: " & piEncounterID
    
      'Unload Me
      
      'frmBasic.FormID = lFormID
      'frmBasic.PatientID = piPatientID
      frmBasic.EncounterID = piEncounterID
      'frmBasic.ProviderID = piProviderID
      RefreshComponents
End Sub
