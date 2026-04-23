VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmPreview 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Print Preview"
   ClientHeight    =   10395
   ClientLeft      =   9855
   ClientTop       =   3525
   ClientWidth     =   11580
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10395
   ScaleWidth      =   11580
   Begin PT_XP_Button.PT_XP cmdCloseList 
      Height          =   375
      Left            =   5318
      TabIndex        =   0
      Top             =   10020
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
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
      Picture         =   "frmPreview.frx":08CA
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthMode       =   1
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmPreview.frx":12DC
      PictureAlignment=   8
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrtType As String
Public mlngEnctr As Long
Public mlngPat As Long
Public mlngGrdEncr As Long
Public piSelPatientID As Long
Private mFirst As Boolean
Private mintActXId As Integer
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Public mstrLoad_Type As String

Function Call_Print(Optional vstrType As String)
    If vstrType = vbNullString Then
        fobjLib.UserCallFunc "Print"
    Else
        fobjLib.UserCallFunc vstrType
    End If
End Function

Function Load_InPatient_Batch()
Dim strOCX As String
Dim strControl As String
Dim lstrControl As String
Dim lstrSQL As String
Dim strSQL As String
Dim lbooReadOnly As Boolean
Dim cnGetTerms As New ADODB.Connection
Dim rsGetTerms As New ADODB.Recordset

    On Error GoTo Error_Load_InPatient_Batch
    If cnGetTerms.State = adStateClosed Then cnGetTerms.open psConnect(1)
    lstrSQL = "Select ACTX_ID FROM ACTIVE_X " _
        & "Where ACTX_NAME = 'Reporter' And " _
        & "ACTX_CLASS = 'PreviewReport'"
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
                & "values('Reporter', 'PreviewReport', 'Components', 'PreviewReport', '2')"
            Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
            
            strSQL = "Select ACTX_ID FROM ACTIVE_X " _
            & "Where ACTX_NAME = 'Reporter'  And " _
            & "ACTX_CLASS = 'PreviewReport'"
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
    Me.Caption = "Patient Notes"
    Licenses.Add "Reporter.PreviewReport"
    Set fobjOCX = Controls.Add("Reporter.PreviewReport", "ctl" & "PreviewReport", Me)
    fobjOCX.Visible = True
    fobjOCX.Left = 150
    fobjOCX.Top = 100
    Me.Width = fobjOCX.Width + 700
    Set fobjLib = fobjOCX.Object
    fobjLib.ComponentID = mintActXId
    fobjLib.ConnectStringTerms = psConnect(1)
    fobjLib.ConnectStringClinical = psConnect(0)
    fobjLib.PatientID = piSelPatientID
    fobjLib.EncounterID = mlngGrdEncr
    fobjLib.ProviderID = piProviderID
    fobjLib.UserCallSub CVar(udtInPatSched.Appt_Date) & "|" & piProviderID & "|" & piFacilityID
    fobjLib.ReadClinicalData
    cmdCloseList.Left = (Me.Width / 2) - (cmdCloseList.Width / 2)
    cmdCloseList.Top = Val(fobjOCX.Height) + 150
    Me.Height = (fobjOCX.Height) + 1200
     Screen.MousePointer = vbDefault
    Exit Function
Error_Load_InPatient_Batch:
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If
    MsgBox Err.Description, vbInformation
    Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function

Private Function Load_Preview()
Dim strOCX As String
Dim strControl As String
Dim lstrControl As String
Dim lstrSQL As String
Dim strSQL As String
Dim lbooReadOnly As Boolean
Dim cnGetTerms As New ADODB.Connection
Dim rsGetTerms As New ADODB.Recordset

    On Error GoTo Error_Load_Control
    If cnGetTerms.State = adStateClosed Then cnGetTerms.open psConnect(1)
    lstrSQL = "Select ACTX_ID FROM ACTIVE_X " _
        & "Where ACTX_NAME = 'Reporter' And " _
        & "ACTX_CLASS = 'PreviewReport'"
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
                & "values('Reporter', 'PreviewReport', 'Components', 'PreviewReport', '2')"
            Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
            
            strSQL = "Select ACTX_ID FROM ACTIVE_X " _
            & "Where ACTX_NAME = 'Reporter'  And " _
            & "ACTX_CLASS = 'PreviewReport'"
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
    Me.Caption = "Patient Notes"
    Licenses.Add "Reporter.PreviewReport"
    Set fobjOCX = Controls.Add("Reporter.PreviewReport", "ctl" & "PreviewReport", Me)
    fobjOCX.Visible = True
    fobjOCX.Left = 150
    fobjOCX.Top = 100
    Me.Width = fobjOCX.Width + 700
    Set fobjLib = fobjOCX.Object
    fobjLib.ComponentID = mintActXId
    fobjLib.ConnectStringTerms = psConnect(1)
    fobjLib.ConnectStringClinical = psConnect(0)
    fobjLib.PatientID = piSelPatientID
    fobjLib.EncounterID = mlngGrdEncr
    fobjLib.ProviderID = piProviderID
    fobjLib.ReadClinicalData
    cmdCloseList.Left = (Me.Width / 2) - (cmdCloseList.Width / 2)
    cmdCloseList.Top = Val(fobjOCX.Height) + 150
    Me.Height = (fobjOCX.Height) + 1200
     Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Control:
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If
    MsgBox Err.Description, vbInformation
    Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function
Private Function Load_Referral()
Dim strOCX As String
Dim strControl As String
Dim lstrControl As String
Dim lstrSQL As String
Dim strSQL As String
Dim lbooReadOnly As Boolean
Dim cnGetTerms As New ADODB.Connection
Dim rsGetTerms As New ADODB.Recordset

    On Error GoTo Error_Load_Control
    If cnGetTerms.State = adStateClosed Then cnGetTerms.open psConnect(1)
    lstrSQL = "Select ACTX_ID FROM ACTIVE_X " _
        & "Where ACTX_NAME = 'Reporter' And " _
        & "ACTX_CLASS = 'Letter'"
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
                & "values('Reporter', 'Letter', 'Components', 'Letter', '2')"
            Set rsGetTerms = cnGetTerms.Execute(lstrSQL)
            
            strSQL = "Select ACTX_ID FROM ACTIVE_X " _
            & "Where ACTX_NAME = 'Reporter'  And " _
            & "ACTX_CLASS = 'Letter'"
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
    Me.Caption = "Physician Letter"
    Licenses.Add "Reporter.Letter"
    Set fobjOCX = Controls.Add("Reporter.Letter", "ctl" & "Letter", Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 150
        fobjOCX.Top = 100
        Me.Width = fobjOCX.Width + 700
        Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = mintActXId
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        fobjLib.PatientID = piSelPatientID
        fobjLib.EncounterID = mlngGrdEncr
        fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        'Me.Width = fobjOCX.Width + 100
        cmdCloseList.Left = (Me.Width / 2) - (cmdCloseList.Width / 2)
        cmdCloseList.Top = Val(fobjOCX.Height) + 150
        Me.Height = (fobjOCX.Height) + 1200
     Screen.MousePointer = vbDefault
    Exit Function
Error_Load_Control:
    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add "Reporter.Letter"
    End If
    MsgBox Err.Description, vbInformation
    Screen.MousePointer = vbDefault
     Exit Function
     Resume
End Function


Private Sub cmdCloseList_Click()
    Unload Me

End Sub


 
Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
    End If
    centerForm Me
    Exit Sub
End Sub

Sub Form_Load()
    mFirst = True
    mlngPat = piSelPatientID
    Select Case mstrLoad_Type
        Case "Patient Record"
            Load_Preview
        Case "Referral Letter"
            Load_Referral
        Case "InPatient_Records"
            Load_InPatient_Batch
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim strLoad As String
    On Error Resume Next
    strLoad = mstrLoad_Type
'    If strLoad = "Patient Record" Then
'        If IsFormLoaded("frmInPatient") Then
'            frmInPatient.ZOrder 0
'        End If
'    End If
    If Licenses.Count >= 0 Then
        If Licenses.Item(0) <> vbNullString Then
            If Licenses.Item(0) = "PatientTrac_Report.Admission" Then
                If Licenses.Count > 1 Then
                    Licenses.Remove (1)
                End If
            ElseIf Licenses.Count > 0 Then
                Licenses.Remove (0)
            End If
        End If
        Licenses.Clear
    End If
    Set fobjLib = Nothing
    Set fobjOCX = Nothing
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

 





