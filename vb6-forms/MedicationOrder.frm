VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmMedicationOrder 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PatientTrac Medication Administration Record"
   ClientHeight    =   9285
   ClientLeft      =   4185
   ClientTop       =   2610
   ClientWidth     =   13485
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MedicationOrder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   13485
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   525
      Left            =   10140
      TabIndex        =   0
      Top             =   0
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   926
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
      Picture         =   "MedicationOrder.frx":063E
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close Medication Administration Record"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "MedicationOrder.frx":1038
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmMedicationOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private fobjLib As PatientTracTypeLib.iClinical
Private fobjOCX As VBControlExtender
Private fstrLastFilePath As String
Private mstrControlName As String
Private mlngFormID As Long
Private fstrTermsConnect As String
Private fstrClinicalConnect As String
Private fintFormID As Integer

Private Sub IdentifyForm()
Dim strSQL As String
Dim varRS
Dim lngRows As Long
Dim bRet As Boolean

    strSQL = "Select form_id From FORMS Where form_obj_name = '" & Me.Name & "'"
    
    LogFile strSQL
    
    bRet = adoSelectQuery(strSQL, fstrTermsConnect, varRS, , lngRows)
    
    If bRet Then
    
        If lngRows > 0 Then
            fintFormID = varRS(0, 0)
        End If
    
    End If
    
End Sub


Public Property Get FormID() As Long
    FormID = mlngFormID
End Property


Public Property Let FormID(ByVal vlngNewValue As Long)
    mlngFormID = vlngNewValue
End Property

Private Sub LoadRelatedLinks()
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
'    cboRelated.Clear
'
'    cn.Open psConnect(1)
'
'    cn.S_RELATED_FORMS "S", (mlngFormID), rs
'
'    With rs
'       If Not .EOF Then
'          Do Until .EOF
'             cboRelated.AddItem (!FORM_CAPTION)
'             cboRelated.ItemData(cboRelated.NewIndex) = (!CHILD_FORM_ID)
'             .MoveNext
'          Loop
'          .Close
'       End If
'    End With
    
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
End Sub


Private Sub cmdClose_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


Private Sub Form_Load()

Load_MARS

End Sub

Private Sub Load_MARS()
Dim strOCX As String
Dim strControl As String
Dim intACTXID As Integer
Dim strRemarks As String

    On Error GoTo Error_Load_StimSheet
    
    strOCX = "PT_MEDICATIONS"
    strControl = "MARS"
    
    intACTXID = Get_ComponentID(strOCX, strControl)
    
    Licenses.Add strOCX & "." & strControl
    Set fobjOCX = Controls.Add(strOCX & "." & strControl, "ctlMars", Me)
        fobjOCX.Visible = True
        fobjOCX.Left = 30
        fobjOCX.Top = 50
        
    Set fobjLib = fobjOCX.Object
        fobjLib.ComponentID = intACTXID
        fobjLib.ConnectStringTerms = psConnect(1)
        fobjLib.ConnectStringClinical = psConnect(0)
        'fobjLib.PatientID = piPatientID
        'fobjLib.EncounterID = piEncounterID
        'fobjLib.ProviderID = piProviderID
        fobjLib.ReadClinicalData
        
        Me.Width = fobjOCX.Width + 200
        centerForm Me
        'cmdClose.top = fobjOCX.top + 100
    Exit Sub
    
Error_Load_StimSheet:

    If Err.Number = 732 Or Err.Number = 733 Or Err.Number = 711 Then Resume Next
    
    If Err.Number = 731 Or Err.Number = 747 Then
        Licenses.Add strOCX & "." & strControl
    End If

     ErrHandler _
     Me.Name & " Load_StimSheet", _
     Err.Number, _
     Err.Description
     Screen.MousePointer = vbDefault
     Exit Sub
     Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    Unload Me
    Exit Sub
End Sub


