VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmSpecOrder 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Patient Special Order"
   ClientHeight    =   8865
   ClientLeft      =   8805
   ClientTop       =   3270
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpecOrder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   10500
   Begin PatientTrac.ccApprove ccaSignature 
      Height          =   1035
      Left            =   11250
      TabIndex        =   13
      Top             =   3270
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1826
   End
   Begin VB.ComboBox cboProvider 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   420
      Width           =   2565
   End
   Begin VB.PictureBox picSignature 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11250
      ScaleHeight     =   915
      ScaleWidth      =   5655
      TabIndex        =   10
      Top             =   1500
      Visible         =   0   'False
      Width           =   5715
   End
   Begin VB.ListBox lstOrder 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   5520
      TabIndex        =   8
      ToolTipText     =   "Select to Review Previous Order"
      Top             =   420
      Width           =   3045
   End
   Begin VB.PictureBox picLogo 
      BackColor       =   &H00FFFFFF&
      Height          =   6945
      Left            =   60
      ScaleHeight     =   6885
      ScaleWidth      =   10305
      TabIndex        =   5
      Top             =   1260
      Width           =   10365
      Begin VB.PictureBox picPatientTrac 
         Height          =   6075
         Left            =   1260
         Picture         =   "frmSpecOrder.frx":09EA
         ScaleHeight     =   6015
         ScaleWidth      =   7635
         TabIndex        =   6
         Top             =   630
         Width           =   7695
      End
      Begin VB.Label lblPatientTrac 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "PATIENTTRAC SPECIAL ORDERS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   705
         Left            =   360
         TabIndex        =   7
         Top             =   60
         Width           =   9405
      End
   End
   Begin FPSpreadADO.fpSpread grdSpec 
      Height          =   6975
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Left click to select a row and right click to remove an item."
      Top             =   1260
      Width           =   10365
      _Version        =   458752
      _ExtentX        =   18283
      _ExtentY        =   12303
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      ColHeaderDisplay=   0
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoHeadings   =   0   'False
      DAutoSizeCols   =   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   16777215
      GridColor       =   14737632
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   10
      MaxRows         =   10
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      ScrollBarMaxAlign=   0   'False
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   16761024
      ShadowDark      =   8388608
      ShadowText      =   4194304
      SpreadDesigner  =   "frmSpecOrder.frx":EA9C
      UnitType        =   0
      UserResize      =   0
      Appearance      =   1
      TextTipDelay    =   50
      ClipboardOptions=   0
      CellNoteIndicator=   3
   End
   Begin PT_XP_Button.PT_XP cmdPrescript 
      Height          =   345
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
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
      Picture         =   "frmSpecOrder.frx":ED72
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Print Order"
      DepthEvent      =   1
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSpecOrder.frx":F30C
      PictureAlignment=   0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdClose 
      Height          =   345
      Left            =   6323
      TabIndex        =   2
      Top             =   8370
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   609
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
      Picture         =   "frmSpecOrder.frx":F65E
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Close"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSpecOrder.frx":10070
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSave 
      Height          =   345
      Left            =   5153
      TabIndex        =   3
      Top             =   8370
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   609
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
      Picture         =   "frmSpecOrder.frx":103C2
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Save"
      DepthEvent      =   1
      ForeColorDisabled=   16761024
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSpecOrder.frx":1075C
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdPatient 
      Height          =   345
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   609
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
      Picture         =   "frmSpecOrder.frx":10AAE
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Select Patient"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSpecOrder.frx":11048
      PictureAlignment=   0
      ShowFocus       =   -1  'True
   End
   Begin PT_XP_Button.PT_XP cmdSign 
      Height          =   345
      Left            =   3143
      TabIndex        =   14
      Top             =   8370
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   609
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
      Picture         =   "frmSpecOrder.frx":1139A
      AppearanceThemes=   3
      TransparentColor=   14737632
      Caption         =   "Signature Pad"
      DepthEvent      =   1
      ForeColorDisabled=   8388608
      ForeColorOver   =   8388608
      ForeColorFocus  =   8388608
      ForeColorDown   =   8388608
      PictureDisabled =   "frmSpecOrder.frx":11934
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblProvider 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   2700
      TabIndex        =   12
      Top             =   150
      Width           =   855
   End
   Begin VB.Label lblPrevOrder 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Orders"
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
      Height          =   315
      Left            =   5520
      TabIndex        =   9
      Top             =   150
      Visible         =   0   'False
      Width           =   1845
   End
End
Attribute VB_Name = "frmSpecOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Private mlngGridRow As Long
Private mlngGridCol As Long
Private mlngEnctr As Long
Private mstrComment As String
Private mstrSignatureToken As String
Private mlngProviderID As Long
Private mlngOrderNo As Long
Private udtEnctr As udtEnctr
Private Type udtEnctr
    EnctrID As Long
    StartDate As String
    EndDate As String
    VisitReason As String
    FacID As Long
    Provider As String
    Admission As String
    Discharge As String
    ServicePlace As String
    ServicePlaceCode As String
    EncounterType As String
    Appointment As String
    ProviderID As Long
    VisitType As String
    Intrvn_CD As String
    VisitTime As String
End Type

Public Function Get_OrderProv()
Dim strSQL As String
Dim Flds As ADODB.Fields
Dim Fld As ADODB.Field
Dim strMedProvider As String
Dim cnMedProv As New ADODB.Connection
Dim rsMedProv As New ADODB.Recordset
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim strSuffix As String
Dim strCity As String
Dim strState As String
Dim strZip As String
Dim strZipExt As String
Dim booSkip As Boolean
Dim lngProvID As Long

    On Error GoTo Error_Get_OrderProv
    Screen.MousePointer = vbHourglass
    If cnMedProv.State = adStateClosed Then cnMedProv.open psConnect(0)
    strSQL = "Load_Active_Providers"
    Set rsMedProv = cnMedProv.Execute(strSQL)
    With rsMedProv
        Do Until .EOF
            Clear_Provider
            booSkip = True
            strFirst = vbNullString
            strLast = vbNullString
            strMiddle = vbNullString
            strTitle = vbNullString
            strSuffix = vbNullString
            lngProvID = 0
            If Not IsNull(!Title) Then
                strTitle = (!Title)
                If TextFound("MD", strTitle, Len("MD"), True) Or _
                    TextFound("M.D.", strTitle, Len("M.D."), True) Or _
                    TextFound("D.O.", strTitle, Len("D.O."), True) Or _
                    TextFound("DO", strTitle, Len("DO"), True) Then
                        booSkip = False
                    End If
            End If
            If Not booSkip Then
                Set Flds = .Fields
                For Each Fld In Flds
                    If Not IsNull(Fld.Value) Then
                        If Fld.Value <> vbNullString Then
                            Select Case UCase(Fld.Name)
                            Case "FIRST_NAME"
                                strFirst = Fld.Value
                            Case "MIDDLE_NAME"
                                strMiddle = Fld.Value
                            Case "LAST_NAME"
                                strLast = Fld.Value
                            Case "SUFFIX"
                                strSuffix = Fld.Value
                            Case "TITLE"
                                strTitle = Fld.Value
                            Case "PROVIDER_ID"
                                lngProvID = Fld.Value
                            End Select
                        End If
                    End If
                Next
            End If
            .MoveNext
            If Trim(strFirst) <> vbNullString Then
                 clsProvider.Name = Trim(strFirst)
             End If
            If Trim(strMiddle) <> vbNullString Then
                strMiddle = IsProperCase(strMiddle)
                 clsProvider.Name = clsProvider.Name & " " & Left(strMiddle, 1)
             End If
            If Trim(strLast) <> vbNullString Then
                 clsProvider.Name = clsProvider.Name & " " & Trim(strLast)
             End If
            If Trim(strSuffix) <> vbNullString Then
                 clsProvider.Name = clsProvider.Name & ", " & Trim(strSuffix)
             End If
            If Trim(strTitle) <> vbNullString And clsProvider.Name <> vbNullString Then
                 clsProvider.Name = clsProvider.Name & ", " & Trim(strTitle)
                 clsProvider.Title = Trim(strTitle)
             End If
            If Trim(clsProvider.Name) <> vbNullString Then
                cboProvider.AddItem clsProvider.Name
                If IsNumeric(lngProvID) Then cboProvider.ItemData(cboProvider.NewIndex) = lngProvID
            End If
        Loop
    End With
    If cnMedProv.State = adStateOpen Then
        cnMedProv.Close
        Set cnMedProv = Nothing
        Set rsMedProv = Nothing
        Set Flds = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_OrderProv:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_OrderProv", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Function Save_Data()
    If Val(udtPatientName.IntId) > 0 Then
        Write_Order
    End If
    Exit Function
End Function

Private Function Set_Image(vstrPhoto As String, Optional vlngRow As Long, Optional vlngCol As Long, Optional vstrCellSpan As String)
Dim strEncodedSignature As String
Dim strRefData As String
Dim booInvalid As Boolean
Dim lngRow As Long: Dim lngNoteRow As Long
Dim lngCol As Long
Dim strPhoto As String
Dim lngCellSpan As Long
Dim IntAxtX As Integer
Dim Arr
Dim FSO As New FileSystemObject
Dim strPicture As PictureBox
Dim strSQL As String
Dim cnProv As New ADODB.Connection
Dim rsProv As New ADODB.Recordset
Dim lngProv As Long
Dim strFirst As String
Dim strLast As String
Dim strMiddle As String
Dim strTitle As String
Dim strSuffix As String
Dim strName As String

    On Error GoTo Error_SetImage
    Screen.MousePointer = vbHourglass
    If vstrPhoto = "SIGNATURE_MD" Then
        Arr = Split(LookUpCellTag("SIGNATURE_MD"), "|")
        If UBound(Arr) > 0 Then
            lngRow = Arr(1)
            lngCol = Arr(0)
        End If
        If mbooDefaultSign Then
            clsSignature.SignReason = "PrescriptionDefault"
            If clsSignature.SignReason <> vbNullString Then
                If mlngProviderID = 0 Then mlngProviderID = piProviderID
                strEncodedSignature = Get_DefaultSignature(clsSignature.SignReason, mlngProviderID)
            End If
            If Trim(strEncodedSignature) <> vbNullString Then
                picSignature.Width = 4000
                picSignature.Height = 1000
                picSignature.BackColor = NormalColor
                booInvalid = ccaSignature.DecodeSignatureValues(strEncodedSignature, clsSignature.SignReason)
                picSignature.Picture = ccaSignature.Picture
                SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                If FSO.FileExists(App.Path & "\Images\SignNote.bmp") Then
                    SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                    strPhoto = App.Path & "\Images\SignNote.bmp"
                End If
                strPhoto = App.Path & "\Images\SignNote.bmp"
                Format_CellTypeEdit grdSpec, lngRow, lngCol, "4", "", "TypeHAlignCenter", "SIGNATURE_MD", False, NormalColor
                Format_CellTypeSignature grdSpec, lngRow, lngCol, "4", strPhoto, "TypeHAlignCenter", "SIGNATURE_MD", 47
            End If
        End If
        If booInvalid Or mbooDefaultSign = False Then
            If Trim(udtPatientName.IntId) <> vbNullString And Val(udtPatientName.IntId) > 0 Then
                clsSignature.SignReason = mstrComment
                If clsSignature.SignReason <> vbNullString Then
                    strEncodedSignature = Get_Signature(App.EXEName, "Special Order", "ccMDSig", 1, Val(udtPatientName.IntId), mlngEnctr, clsSignature.SignReason)
                Else
                    clsSignature.SignReason = Read_Default(mlngProviderID)
                    If Trim(clsSignature.SignReason) <> vbNullString Then
                        strEncodedSignature = Get_Signature(App.EXEName, "Special Order", "ccMDSig", 1, Val(udtPatientName.IntId), mlngEnctr, clsSignature.SignReason)
                    End If
                End If
                If Trim(strEncodedSignature) <> vbNullString Then
                    picSignature.Width = 4000
                    picSignature.Height = 1000
                    picSignature.BackColor = NormalColor
                    booInvalid = ccaSignature.DecodeSignatureValues(strEncodedSignature, clsSignature.SignReason)
                    picSignature.Picture = ccaSignature.Picture
                    SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                    If FSO.FileExists(App.Path & "\Images\SignNote.bmp") Then
                        SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                        strPhoto = App.Path & "\Images\SignNote.bmp"
                    End If
                    Format_CellTypeSignature grdSpec, lngRow, lngCol, "4", strPhoto, "TypeHAlignCenter", "SIGNATURE_MD", 47
                Else
                    Format_CellTypeStatic grdSpec, lngRow, lngCol, "4", "PHYSICIAN SIGNATURE", "TypeHAlignCenter", "SIGNATURE_MD", False, grdGray
                End If
            End If
        End If
    ElseIf vstrPhoto = "SIGNATURE_NURSE" Then
            clsSignature.SignReason = mstrComment
            If Trim(udtPatientName.IntId) <> vbNullString And Val(udtPatientName.IntId) > 0 Then
                If clsSignature.SignReason <> vbNullString Then
                    strEncodedSignature = Get_Signature(App.EXEName, "Special Order", "ccNurseSig", 1, Val(udtPatientName.IntId), mlngEnctr, clsSignature.SignReason)
                Else
                    clsSignature.SignReason = Read_Default(mlngProviderID)
                    If Trim(clsSignature.SignReason) <> vbNullString Then
                        strEncodedSignature = Get_Signature(App.EXEName, "Special Order", "ccNurseSig", 1, Val(udtPatientName.IntId), mlngEnctr, clsSignature.SignReason)
                    End If
                End If
                If Trim(strEncodedSignature) <> vbNullString Then
                    picSignature.Width = 4000
                    picSignature.Height = 1000
                    picSignature.BackColor = NormalColor
                    booInvalid = ccaSignature.DecodeSignatureValues(strEncodedSignature, clsSignature.SignReason)
                    picSignature.Picture = ccaSignature.Picture
                    SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                    If FSO.FileExists(App.Path & "\Images\SignNote.bmp") Then
                        SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                        strPhoto = App.Path & "\Images\SignNote.bmp"
                    End If
                    Format_CellTypeSignature grdSpec, lngRow, lngCol, "4", strPhoto, "TypeHAlignCenter", "SIGNATURE_NURSE", 47
                Else
                    Format_CellTypeStatic grdSpec, lngRow, lngCol, "4", "RESIDENT/CLINICIAN SIGNATURE", "TypeHAlignCenter", "SIGNATURE_NURSE", False, grdGray
                End If
                'Now we have to find if this has been approved
                strRefData = Get_Second_Signature(App.EXEName, "Special Order", Val(udtPatientName.IntId), mlngEnctr)
                If strRefData <> vbNullString Then
                     Arr = Split(LookUpCellTag("SIGNATURE_MD"), "|")
                    If UBound(Arr) > 0 Then
                        lngRow = Arr(1)
                        lngCol = Arr(0)
                    End If
                    Arr = Split(strRefData, "~")
                    If UBound(Arr) > 0 Then
                        If IsNumeric(Arr(1)) Then
                            lngProv = Arr(1)
                            strEncodedSignature = Arr(0)
                            If cnProv.State = adStateClosed Then cnProv.open psConnect(0)
                            strSQL = "Get_Providers " & lngProv
                            Set rsProv = cnProv.Execute(strSQL)
                            If Not rsProv.EOF Then
                                If rsProv.Fields("FIRST_NAME") <> vbNullString Then
                                    strFirst = rsProv.Fields("FIRST_NAME")
                                ElseIf rsProv.Fields("MIDDLE_NAME") <> vbNullString Then
                                    strMiddle = rsProv.Fields("MIDDLE_NAME")
                                ElseIf rsProv.Fields("LAST_NAME") <> vbNullString Then
                                    strLast = rsProv.Fields("LAST_NAME")
                                ElseIf Not IsNull(rsProv.Fields("SUFFIX")) Then
                                    If rsProv.Fields("SUFFIX") <> vbNullString Then
                                        strSuffix = rsProv.Fields("SUFFIX")
                                    End If
                                ElseIf rsProv.Fields("TITLE") <> vbNullString Then
                                    strTitle = rsProv.Fields("TITLE")
                                End If
                            End If
                            If Trim(strMiddle) <> vbNullString Then strMiddle = Left(strMiddle, 1)
                            strName = strFirst & " " & strMiddle & " " & strLast
                            If Trim(strSuffix) <> vbNullString Then strName = strName & ",  " & Trim(strSuffix)
                            If Trim(strTitle) <> vbNullString Then strName = strName & ",  " & Trim(strTitle)
                            If cnProv.State = adStateOpen Then
                                cnProv.Close
                                Set cnProv = Nothing
                                Set rsProv = Nothing
                            End If
                            If Trim(strEncodedSignature) <> vbNullString Then
                                picSignature.Width = 4000
                                picSignature.Height = 1000
                                picSignature.BackColor = NormalColor
                                booInvalid = ccaSignature.DecodeSignatureValues(strEncodedSignature, clsSignature.SignReason)
                                picSignature.Picture = ccaSignature.Picture
                                SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                                If FSO.FileExists(App.Path & "\Images\SignNote.bmp") Then
                                    SavePicture picSignature.Image, App.Path & "\Images\SignNote.bmp"
                                    strPhoto = App.Path & "\Images\SignNote.bmp"
                                End If
                                If clsSignature.Physician <> vbNullString Then
                                     Arr = Split(LookUpCellTag("PHYSCIAN_NOTE"), "|")
                                    If UBound(Arr) > 0 Then
                                        lngNoteRow = Arr(1)
                                    End If
                                    If lngNoteRow > 0 Then
                                        Format_CellTypeEdit grdSpec, lngNoteRow, lngCol, "4", clsSignature.Physician, "TypeHAlignCenter", "PHYSCIAN_NOTE"
                                        Format_CellTypeSignature grdSpec, lngRow, lngCol, "4", strPhoto, "TypeHAlignCenter", "SIGNATURE_MD", 47
                                    Else
                                        lngNoteRow = lngRow
                                        Format_CellTypeEdit grdSpec, lngNoteRow, lngCol, "4", clsSignature.Physician, "TypeHAlignCenter", "PHYSCIAN_NOTE"
                                        lngNoteRow = lngNoteRow + 1
                                        grdSpec.MaxRows = grdSpec.MaxRows + 1
                                        grdSpec.InsertRows lngRow, 1
                                        Format_CellTypeSignature grdSpec, lngNoteRow, lngCol, "4", strPhoto, "TypeHAlignCenter", "SIGNATURE_MD", 47
                                    End If
                                Else
                                    Format_CellTypeSignature grdSpec, lngRow, lngCol, "4", strPhoto, "TypeHAlignCenter", "SIGNATURE_MD", 47
                                End If
                            End If
                        End If
                    End If
                End If
            End If
    Else
        If vstrCellSpan <> vbNullString Then
            lngCellSpan = Val(vstrCellSpan)
        Else
            lngCellSpan = 1
        End If
        If vlngRow > 0 Then
            lngRow = vlngRow
        Else
            lngRow = 1
        End If
        If vlngCol > 0 Then
            lngCol = vlngCol
        Else
            lngCol = 1
        End If
        If vstrPhoto <> vbNullString Then
            If FSO.FileExists(vstrPhoto) Then
                Format_CellTypePicture grdSpec, lngRow, lngCol, CStr(lngCellSpan), vstrPhoto, "TypeHAlignCenter", "", 75
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_SetImage:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, vstrPhoto, "Collect Report Values", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume
End Function
Private Function Read_Default(vlngProviderID As Long) As String
Dim strSQL As String
Dim cnDefault As New ADODB.Connection
Dim rsDefault As New ADODB.Recordset
Dim strClinicalDesc As String

    On Error GoTo Error_Read_Default
    Screen.MousePointer = vbHourglass
    Read_Default = vbNullString
    If cnDefault.State = adStateClosed Then cnDefault.open psConnect(1)
    strSQL = "Get_Default_Value '" & App.EXEName & "','SpecialOrder' '," & vlngProviderID
    Set rsDefault = cnDefault.Execute(strSQL)
    With rsDefault
        If Not .EOF Then
            If Not IsNull(!CLINICAL_VALUE) Then
                Read_Default = Trim((!CLINICAL_VALUE))
            End If
        End If
    End With
    If cnDefault.State = adStateOpen Then
        cnDefault.Close
        Set cnDefault = Nothing
        Set rsDefault = Nothing
    End If
    Exit Function
    Screen.MousePointer = vbDefault
Error_Read_Default:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Default", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function


Private Function Format_CellTypeSignature(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long)
Dim CellNote As Long
Dim lngSpan As Long
Dim FSO As New FileSystemObject

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypePicture
            '.TypePictMaintainScale = True
            '.TypePictStretch = True
            .TypePictCenter = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf vstrSpan <> vbNullString Then
               lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
             End If
            If FSO.FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            .RowHeight(vlngRow) = picSignature.Height
            .RowHeight(vlngRow) = 49
          .BlockMode = False
    End With
End Function

Private Function Format_SpecOrder()
Dim Data
Dim lngRow As Long
Dim varCellValue
Dim i%
Dim strBlank As String
Dim lngGridRow As Long
Dim lngGridCol As Long
Dim CellNote As Long
Dim strEncodedSignature As String
Dim strRefData As String
Dim booInvalid As Boolean

    On Error GoTo Error_Format_SpecOrder
    Screen.MousePointer = vbHourglass
    Add_ProviderData
    Add_PatientData
    With grdSpec
        .ReDraw = False
         lngGridRow = .MaxRows + 3
        .MaxRows = lngGridRow
        Format_CellTypeEdit grdSpec, lngGridRow, 1, "9", "PHYSICIAN ORDER" & " " & "ORDER NUMBER:", "TypeHAlignCenter", "ORDER_NUMBER", HeaderColor
        grdSpec.SetCellBorder 1, lngGridRow, 9, lngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        lngGridRow = .MaxRows + 3
        .MaxRows = lngGridRow
        Format_CellTypePicture grdSpec, lngGridRow, 1, "1", App.Path & "\Images\RX.bmp", "TypeHAlignCenter"
        Format_CellTypeEdit grdSpec, lngGridRow, 3, "6", "", "TypeHAlignCenter", "COMMENT", grdGray, 35
        grdSpec.SetCellBorder 3, lngGridRow, 8, lngGridRow, CellBorderIndexOutline, GridBlue, CellBorderStyleDot
        lngGridRow = lngGridRow + 2
        If .MaxRows < lngGridRow Then .MaxRows = lngGridRow
        lngGridRow = lngGridRow + 2
        If .MaxRows < lngGridRow Then .MaxRows = lngGridRow + 2
        Format_CellTypeEdit grdSpec, lngGridRow, 1, "4", "PHYSICIAN SIGNATURE", "TypeHAlignCenter", "SIGNATURE_MD"
        lngGridRow = lngGridRow + 3
        If .MaxRows < lngGridRow Then .MaxRows = lngGridRow
        Format_CellTypeEdit grdSpec, lngGridRow, 1, "8", clsProvider.Name, "TypeVAlignLeft"
        lngGridRow = lngGridRow + 1
        If .MaxRows < lngGridRow Then .MaxRows = lngGridRow
        Format_CellTypeEdit grdSpec, lngGridRow, 1, "5", "Physician's Authenticated Signature", "TypeVAlignLeft"
        .MaxRows = lngGridRow + 2
        .SetCellBorder 9, 1, 9, .MaxRows, CellBorderIndexRight, GridBlue, CellBorderStyleSolid
        .ReDraw = True
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_Format_SpecOrder:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Format_SpecOrder", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Get_OrderEnctr()
Dim cnEnctr As New ADODB.Connection
Dim rsEnctr As New ADODB.Recordset
Dim strSQL As String
Dim strEnctr As String
Dim strEnctrDate As String
Dim lngDateDiff As Long
Dim strAppt As String

    On Error GoTo Error_Get_OrderEnctr
    If cnEnctr.State = adStateClosed Then cnEnctr.open psConnect(0)
    If mlngEnctr = 0 Then
        strEnctrDate = Format(Now, "mm/dd/yyyy")
        strSQL = "Get_Encounter_Patient_ID " & Val(udtPatientName.IntId)
        Set rsEnctr = cnEnctr.Execute(strSQL)
        With rsEnctr
            Do Until .EOF
                If (!ENCNTR_START_DATE) <> vbNullString Then
                    strEnctr = Format((!ENCNTR_START_DATE), "mm/dd/yyyy")
                        If strEnctr <> strEnctrDate Then
                        lngDateDiff = DateDiff("d", strEnctrDate, strEnctr)
                        If lngDateDiff <= 0 Then
                            mlngEnctr = (!Encounter_ID)
                            Exit Do
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End With
    Else
        Get_CurrentEnctr
    End If
    If cnEnctr.State = adStateClosed Then
        cnEnctr.Close
        Set cnEnctr = Nothing
        Set rsEnctr = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_OrderEnctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_OrderEnctr", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cboProvider_Click()
Dim Idx As Integer
    For Idx = 0 To cboProvider.listcount - 1
        If cboProvider.List(Idx) = cboProvider.Text Then
            mlngProviderID = cboProvider.ItemData(Idx)
            Exit For
        End If
    Next Idx
    Exit Sub
End Sub


Private Sub cmdClose_Click()
    grdSpec.ClearRange 1, 1, grdSpec.MaxCols, grdSpec.MaxRows, False
    Clear_UdtPatient
    Unload Me
End Sub

Private Sub cmdPatient_Click()

    Clear_UdtPatient
    mstrComment = vbNullString
    frmListPatient.Caption = "Search Patients"
    frmListPatient.Show vbModal
    If frmListPatient.mlngListItemData > 0 Then udtPatientName.IntId = frmListPatient.mlngListItemData
    If Trim(udtPatientName.IntId) <> vbNullString And Val(udtPatientName.IntId) > 0 Then
         Format_SpecOrder
         picLogo.Visible = False
    End If
    Read_Order
    Exit Sub
End Sub
Private Function Add_ProviderData()
Dim lngRow As Long


    Get_Providers piProviderID
    With grdSpec
        .ReDraw = False
        .ProcessTab = True
        .FontBold = True
        .FontSize = 8
        .GridSolid = False
        .BorderStyle = BorderStyleFixedSingle
        .GridShowHoriz = False
        .GridShowVert = False
        .RowHeadersShow = False
        .ColHeadersShow = False
        .MaxCols = 9
        .MaxRows = 0
        .ColWidth(1) = 15
        .ColWidth(2) = 7
        .ColWidth(7) = 12
        lngRow = lngRow + 1
        .MaxRows = lngRow
        Format_CellTypeEdit grdSpec, lngRow, 1, "9", clsProvider.Practice, "TypeHAlignCenter"
        lngRow = lngRow + 1
        .MaxRows = lngRow
        Format_CellTypeEdit grdSpec, lngRow, 1, "9", clsProvider.Name, "TypeHAlignCenter"
        lngRow = lngRow + 1
        .MaxRows = lngRow
        If clsProvider.Address2 <> vbNullString Then
            Format_CellTypeEdit grdSpec, lngRow, 1, "4", clsProvider.Address & ", " & clsProvider.Address2, "TypeHAlignLeft"
        Else
            Format_CellTypeEdit grdSpec, lngRow, 1, "4", clsProvider.Address, "TypeHAlignLeft"
        End If
        If Trim(clsProvider.Facsimile) <> vbNullString Then
            Format_CellTypeEdit grdSpec, lngRow, 7, "2", "DEA Number:" & " " & clsProvider.DEA_Number, "TypeHAlignLeft"
        End If
        lngRow = lngRow + 1
        .MaxRows = lngRow
        Format_CellTypeEdit grdSpec, lngRow, 1, "4", clsProvider.CityStateZip, "TypeHAlignLeft"
        If Trim(clsProvider.Facsimile) <> vbNullString Then
            Format_CellTypeEdit grdSpec, lngRow, 7, "2", "Medical License No.:" & " " & clsProvider.License_No, "TypeHAlignLeft"
        End If
        lngRow = lngRow + 1
        .MaxRows = lngRow
        If Trim(clsProvider.Telephone) <> vbNullString Then
            Format_CellTypeEdit grdSpec, lngRow, 1, "4", "Telephone: " & clsProvider.Telephone, "TypeHAlignLeft"
        End If
        Format_CellTypeEdit grdSpec, 5, 7, "2", "Date:" & " " & Format(Now, "MMMM dd, yyyy"), "TypeHAlignLeft"
        lngRow = lngRow + 1
        .MaxRows = lngRow
        If Trim(clsProvider.Facsimile) <> vbNullString Then
            Format_CellTypeEdit grdSpec, lngRow, 1, "4", "Fax: " & clsProvider.Facsimile, "TypeHAlignLeft"
        End If
        lngRow = lngRow + 1
        .MaxRows = lngRow
        .ReDraw = True
    End With
        Add_Spacer lngRow, grdSpec
    Exit Function
End Function

Private Function Format_CellTypeCombo(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, _
vstrString As String, vstrHAligh As String, Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
              .Col = vlngCol
              .Col2 = vlngCol
              .Row = vlngRow
              .Row2 = vlngRow
              .CellType = CellTypeComboBox
              .TypeMaxEditLen = 32000
              .TypeComboBoxList = vstrString
              .TypeComboBoxWidth = 1
              If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
              If vstrCellTag <> vbNullString Then
                .CellTag = vstrCellTag
              End If
              If vstrSpan <> vbNullString Then
                If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                    .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
                Else
                    lngSpan = Val(vstrSpan)
                    If lngSpan > 0 Then
                        .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                    End If
                End If
              End If
              If vlngBackColor > 0 Then
                 .BackColor = vlngBackColor
              End If
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              .TypeComboBoxAutoSearch = TypeComboBoxAutoSearchSingleChar
          .BlockMode = False
    End With
End Function
Private Function Add_PatientData()
Dim lngRow As Long

    Get_PatientName Val(udtPatientName.IntId)
    With grdSpec
        .ReDraw = False
        lngRow = .MaxRows + 2
        .MaxRows = lngRow
        Format_CellTypeEdit grdSpec, lngRow, 1, "2", "Patient Name", "TypeHAlignLeft"
        Format_CellTypeEdit grdSpec, lngRow, 3, "2", udtPatientName.Name, "TypeHAlignLeft"
        Format_CellTypeEdit grdSpec, lngRow, 7, "1", "Gender", "TypeHAlignLeft"
        Format_CellTypeEdit grdSpec, lngRow, 8, "1", udtPatientName.Gender, "TypeHAlignLeft"
        lngRow = lngRow + 1
        .MaxRows = lngRow
        Format_CellTypeEdit grdSpec, lngRow, 1, "2", "Address", "TypeHAlignLeft"
        Format_CellTypeEdit grdSpec, lngRow, 3, "2", Trim(udtPatientName.Address) & "    " & Trim(udtPatientName.CityStateZip), "TypeHAlignLeft"

        Format_CellTypeEdit grdSpec, lngRow, 7, "1", "DOB", "TypeHAlignLeft"
        Format_CellTypeEdit grdSpec, lngRow, 8, "2", Format(udtPatientName.DOB, "MMMM DD, YYYY"), "TypeHAlignLeft"
        lngRow = lngRow + 1
        .MaxRows = lngRow
        .ReDraw = True
    End With
        Add_Spacer lngRow, grdSpec
    Exit Function
End Function
Private Function Format_CellTypeEdit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngBackColor As Long, Optional vlngRowHeight As Long)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
              .Col = vlngCol
              .Col2 = vlngCol
              .Row = vlngRow
              .Row2 = vlngRow
              .CellType = CellTypeEdit
              .TypeMaxEditLen = 32000
              .SetText vlngCol, vlngRow, Trim(vstrText)
              .TypeEditMultiLine = True   'Make multi-lines
              .FontName = "Times New Roman"
              .FontSize = 8
              .FontBold = True
              'If vbooUnderline Then .FontUnderline = True
              If vlngBackColor > 0 Then .BackColor = vlngBackColor
              If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
              ElseIf Trim(vstrSpan) <> vbNullString Then
                lngSpan = vstrSpan
                If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
              End If
              .Text = Trim(vstrText)
              CellNote = .MaxTextCellHeight
              If .RowHeight(vlngRow) < CellNote Then
                .RowHeight(vlngRow) = CellNote
              End If
              .TypeVAlign = TypeVAlignCenter
                If vstrHAligh = "TypeHAlignCenter" Then
                    .TypeHAlign = TypeHAlignCenter
                Else
                    .TypeHAlign = TypeHAlignLeft
                End If
              If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
              If vlngRowHeight > 0 Then .RowHeight(vlngRow) = vlngRowHeight
          .BlockMode = False
    End With
End Function
Private Function Format_CellTypeSplit(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vbooUnderLine As Boolean, Optional vlngBackColor As Long, Optional vstrCellNote As String, Optional vbooCellHeight As Boolean, _
Optional vlngFontSize As Long)
Dim CellNote As Long
Dim lngSpan As Long

    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeEdit
            .TypeMaxEditLen = 32000
            .TypeEditMultiLine = True   'Make multi-lines
            .FontName = "Times New Roman"
            If vlngFontSize = 0 Then
                .FontSize = 10
            Else
                .FontSize = vlngFontSize
            End If
            .FontBold = True
            If vbooUnderLine Then .FontUnderline = True
            If vlngBackColor > 0 Then .BackColor = vlngBackColor
            If vstrCellNote <> vbNullString Then .CellNote = vstrCellNote
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
            .TypeVAlign = TypeVAlignTop
            If vstrHAligh = "TypeHAlignCenter" Then
              .TypeHAlign = TypeHAlignCenter
            Else
              .TypeHAlign = TypeHAlignLeft
            End If
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
            .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString Then
            lngSpan = vstrSpan
            If lngSpan > 0 Then .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            .SetText vlngCol, vlngRow, Trim(vstrText)
            .Text = Trim(vstrText)
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then
              .RowHeight(vlngRow) = CellNote + 1
            End If
          .BlockMode = False
    End With
End Function

Private Function Format_CellTypeStatic(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrText As String, _
vstrHAligh As String, Optional vstrCellTag As String, Optional vstrVAlign As String, Optional lngBackColor As Long, Optional vstrCellNote As String)
Dim CellNote As Long
Dim lngSpan As Long
    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypeStaticText
            .TypeMaxEditLen = 32000
            .FontName = "Times New Roman"
            .FontSize = 8
            .FontBold = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            ElseIf Trim(vstrSpan) <> vbNullString And IsNumeric(vstrSpan) Then
                lngSpan = Val(vstrSpan)
                .AddCellSpan vlngCol, vlngRow, lngSpan, 1
            End If
            .SetText vlngCol, vlngRow, vstrText
            .Text = vstrText
            If vstrVAlign = "TypeVAlignCenter" Then
                .TypeVAlign = TypeVAlignCenter
            ElseIf vstrVAlign = "TypeVAlignBottom" Then
                .TypeVAlign = TypeVAlignBottom
            Else
               .TypeVAlign = TypeVAlignTop
            End If
            If vstrHAligh = "TypeHAlignCenter" Then
                .TypeHAlign = TypeHAlignCenter
            ElseIf vstrHAligh = "TypeHAlignRight" Then
                .TypeHAlign = TypeHAlignRight
            Else
                .TypeHAlign = TypeHAlignLeft
            End If
            If vstrCellNote <> vbNullString Then
                .CellNote = vstrCellNote
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            End If
            CellNote = .MaxTextCellHeight
            If .RowHeight(vlngRow) < CellNote Then .RowHeight(vlngRow) = CellNote
            If lngBackColor > 0 Then .BackColor = lngBackColor
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
        .BlockMode = False
    End With
End Function
Private Function Format_CellTypePicture(vGrid As Object, vlngRow As Long, vlngCol As Long, vstrSpan As String, vstrPhotoPath As String, vstrHAligh As String, _
Optional vstrCellTag As String, Optional vlngRowHeight As Long)
Dim CellNote As Long
Dim lngSpan As Long
Dim FSO As New FileSystemObject

    lngSpan = vstrSpan
    With vGrid
        .BlockMode = True
            .Col = vlngCol
            .Col2 = vlngCol
            .Row = vlngRow
            .Row2 = vlngRow
            .CellType = CellTypePicture
            .TypePictCenter = True
            If TextFound("maxcols", vstrSpan, Len("maxcols"), True) Then
                .AddCellSpan vlngCol, vlngRow, .MaxCols, 1
            Else
               If lngSpan > 0 And vstrSpan <> vbNullString Then
                    lngSpan = Val(vstrSpan)
                    .AddCellSpan vlngCol, vlngRow, lngSpan, 1
                End If
             End If
            If vlngRowHeight > 0 Then .RowHeight(vlngRow) = vlngRowHeight
            If FSO.FileExists(vstrPhotoPath) Then
                .TypePictPicture = LoadPicture(vstrPhotoPath)
            End If
            If vstrCellTag <> vbNullString Then .CellTag = vstrCellTag
          .BlockMode = False
    End With
End Function
Private Sub Add_Spacer(vlngRow As Long, vObjGrid As Object)

    With vObjGrid
        vlngRow = vlngRow
       If .MaxRows < vlngRow Then .MaxRows = vlngRow
        .BlockMode = True
            .Col = 0
            .Col2 = 0
            .Row = vlngRow
            .Row2 = vlngRow
            .RowHeight(vlngRow) = 3.5
            .BackColor = 12632256
            .AddCellSpan 0, vlngRow, .MaxCols, 1
            .SetCellBorder 0, vlngRow, .MaxCols, vlngRow, 16, &H80&, CellBorderStyleSolid
        .BlockMode = False
    End With
    Exit Sub
End Sub

Private Sub cmdPrescript_Click()
Dim strPatient As String
Dim strProvider As String
Dim strSVP As String
Dim strHeader(0) As String
Dim strHeaders As String
Dim i As Integer
Dim lngPageCnt As Long
Dim intPrint As Integer
Dim strFooter(0) As String
Dim strFoot2 As String

    On Error GoTo Error_cmdPrint_Click
    Screen.MousePointer = vbHourglass
    Write_Order
    mbooCancelPrint = False
    Printer.TrackDefault = True
    CD_ShowPrinter (Me.hwnd)
    If mbooCancelPrint Then
        MsgBox "Print Command Cancelled", vbInformation
        Exit Sub
    End If
    lngPageCnt = grdSpec.PrintPageCount
    With grdSpec
            For intPrint = 1 To Printer.Copies
                    .PrintType = PrintTypeAll
                    .hDCPrinter = Printer.hdc
                    .PrintJobName = "Special_Order" & "_Patient_ID_" & udtPatientName.IntId & "_" & Format(Now, "mm_dd_yyyy")
                    .PrintColHeaders = False
                    .PrintScalingMethod = PrintScalingMethodSmartPrint
                    .PrintRowHeaders = False
                    .PrintMarginLeft = 650
                    .PrintMarginRight = 700
                    .PrintMarginBottom = 900
                    .PrintMarginTop = 900
                    .PrintUseDataMax = True
                    .PrintBorder = False
                    .PrintColor = True
                    .PrintOrientation = PrintOrientationPortrait
                    .PrintShadows = False
                    .PrintSmartPrint = True
                    .PrintSheet
            Next intPrint
    End With
    Write_PracticeReport "Special Order", piPatientID, clsProvider.Name
    Screen.MousePointer = vbDefault
    Exit Sub
Error_cmdPrint_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "cmdPrint_Click", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


Private Sub cmdSave_Click()
    Save_Data
    Clear_UdtPatient
    Unload Me
    Exit Sub
    
End Sub

Private Function Write_Order()
Dim strSQL As String
Dim lngPrimKey As Long
Dim sDate As String
Dim lngRef As Long
Dim strStatus As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset
Dim Arr
Dim lngRow As Long
Dim strOrder As String

    On Error GoTo Error_Write_Order
    Screen.MousePointer = vbHourglass
    If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    If Val(udtPatientName.IntId) <= 0 Then Exit Function
    If mstrComment = vbNullString Then Exit Function
    Get_CurrentEnctr
    sDate = Get_Enctr_Date(mlngEnctr)
    If sDate = vbNullString Then
        sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    Else
        sDate = Format(sDate, "mm/dd/yyyy hh:mm:ss AM/PM")
    End If
    strSQL = "Get_Patient_Order_PT_EnctrID_CPT " & Val(udtPatientName.IntId) & "," _
        & mlngEnctr & ",'99999'"
    Set rsLab = cnLab.Execute(strSQL)
     With rsLab
        If Not .EOF Then
            If (!Order_ID) > 0 Then mlngOrderNo = (!Order_ID)
        End If
    End With
    strStatus = "ORDERED"
    mstrComment = Replace(mstrComment, "'", "''")
    If mlngOrderNo <= 0 Then
        strSQL = "Write_Patient_Orders_SpecialOrder " & Val(udtPatientName.IntId) & "," & mlngEnctr & "," _
            & piProviderID & ",'" & clsProvider.Title & "','Miscellaneous','99999','SPECIAL ORDER','" & mstrComment & "','" _
            & strStatus & "','Routine','" & Format(sDate, "mm/dd/yyyy hh:mm:ss AM/PM") & "'"
        Set rsLab = cnLab.Execute(strSQL)
    Else
        strSQL = "Update_Patient_Order_Special_Order '" & Trim(mstrComment) & "','" _
            & sDate & "','99999','SPECIAL ORDER'," & mlngOrderNo
        Set rsLab = cnLab.Execute(strSQL)
    End If
   strSQL = "Get_Patient_Order_PT_EnctrID_CPT " & Val(udtPatientName.IntId) & "," _
        & mlngEnctr & ",'99999'"
        Set rsLab = cnLab.Execute(strSQL)
        With rsLab
            If Not .EOF Then
                If (!Order_ID) > 0 Then mlngOrderNo = (!Order_ID)
            End If
        End With
    Arr = Split(LookUpCellTag("ORDER_NUMBER"), "|")
    If UBound(Arr) > 0 Then
        lngRow = Arr(1)
    End If
    If lngRow > 1 Then
        strOrder = "PHYSICIAN ORDER" & " " & "ORDER NUMBER:" & " " & mlngOrderNo
        Format_CellTypeEdit grdSpec, lngRow, 1, "9", strOrder, "TypeHAlignCenter", "ORDER_NUMBER", HeaderColor
        grdSpec.SetCellBorder 1, lngRow, 9, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
    End If
    If cnLab.State = adStateOpen Then
        cnLab.Close
        Set cnLab = Nothing
        Set rsLab = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Write_Order:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Write_Order", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Function Get_CurrentEnctr()
Dim cnEncnt As New ADODB.Connection
Dim rsEncnt As New ADODB.Recordset
Dim Fld As ADODB.Field
Dim Flds As ADODB.Fields
Dim strSQL As String

    On Error GoTo Error_Get_CurrentEnctr
    Screen.MousePointer = vbHourglass
    Clear_EnctrString
    If cnEncnt.State = adStateClosed Then cnEncnt.open psConnect(0)
    strSQL = "Get_Encounter_Enctr_ID " & mlngEnctr
    Set rsEncnt = cnEncnt.Execute(strSQL)
    If Not rsEncnt.EOF Then
        Set Flds = rsEncnt.Fields
        For Each Fld In Flds
            If Not IsNull(Fld.Value) Then
                If Trim(Fld.Value) <> vbNullString Then
                    If UCase(Fld.Name) = "FACILITY_ID" Then
                        udtEnctr.FacID = Fld.Value
                    ElseIf UCase(Fld.Name) = "ENCNTR_START_DATE" Then
                         udtEnctr.StartDate = Fld.Value
                    ElseIf UCase(Fld.Name) = "PROVIDER_ID" Then
                         udtEnctr.ProviderID = Fld.Value
                    End If
                End If
            End If
        Next
    End If
    If cnEncnt.State = adStateOpen Then
        cnEncnt.Close
        Set cnEncnt = Nothing
        Set rsEncnt = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_CurrentEnctr:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Get_CurrentEnctr", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function
Private Sub Clear_EnctrString()
    With udtEnctr
        .EnctrID = 0
        .StartDate = vbNullString
        .EndDate = vbNullString
        .VisitReason = vbNullString
        .FacID = 0
        .Provider = vbNullString
        .Admission = vbNullString
        .Discharge = vbNullString
        .ServicePlace = vbNullString
        .ServicePlaceCode = vbNullString
        .EncounterType = vbNullString
        .Appointment = vbNullString
        .ProviderID = 0
        .VisitType = vbNullString
        .Intrvn_CD = vbNullString
        .VisitTime = vbNullString
    End With
End Sub
Private Function Read_Order()
Dim strSQL As String
Dim lngPrimKey As Long
Dim sDate As String
Dim lngRef As Long
Dim strStatus As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset

    On Error GoTo Error_Read_Order
    Screen.MousePointer = vbHourglass
    If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    If udtPatientName.IntId = vbNullString Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    sDate = Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    strSQL = "Get_Patient_Order_PT_CPT " & Val(udtPatientName.IntId) & ",'99999'"
    Set rsLab = cnLab.Execute(strSQL)
     With rsLab
        Do Until .EOF
            If Not IsNull((!INSERT_DATE)) Then
                If (!INSERT_DATE) <> vbNullString Then
                    lstOrder.AddItem Format((!INSERT_DATE), "MMMM DD,YYYY")
                    If (!Order_ID) > 0 Then lstOrder.ItemData(lstOrder.NewIndex) = (!Order_ID)
                End If
            End If
            .MoveNext
        Loop
    End With
    If cnLab.State = adStateOpen Then
        cnLab.Close
        Set cnLab = Nothing
        Set rsLab = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Read_Order:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Read_Order", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub cmdSign_Click()
    Save_Data
    frmMDSign.mstrRef = mstrComment
    If udtPatientName.IntId <> vbNullString Then
        frmMDSign.piSignPatID = Val(udtPatientName.IntId)
    End If
    frmMDSign.piSignEnctr = mlngEnctr
    frmMDSign.mstrRef = mstrComment
    frmMDSign.piSignProv = mlngProviderID
    frmMDSign.vstrSource = "Special Order"
    Me.Hide
    frmMDSign.Show vbModal
    Set_Image "SIGNATURE_MD", 1, 1, "4"
    Me.Show
    Exit Sub
End Sub

Private Sub Form_Activate()
Dim i%

    mlngOrderNo = 0
    If mFirst Then
        centerForm Me
        Get_OrderProv
        If piProviderID > 0 Then
            For i = 0 To cboProvider.listcount - 1
                If cboProvider.ItemData(i) = piProviderID Then
                    cboProvider.ListIndex = i
                End If
            Next i
        End If
    End If
    Exit Sub
End Sub

Private Sub Form_Load()
    mFirst = True
End Sub
Private Function LookUpCellTag(vstrName As String) As String
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim booExit As Boolean

    On Error GoTo Error_LookUpCellTag
    Screen.MousePointer = vbHourglass
    With grdSpec
        LookUpCellTag = vbNullString
        For lngCol = 1 To .MaxCols
            For lngRow = 1 To .MaxRows
                .Col = lngCol
                .Row = lngRow
                If .CellTag <> vbNullString Then
                    If TextFound(vstrName, .CellTag, Len(vstrName), True) Then
                        LookUpCellTag = lngCol & "|" & lngRow
                        booExit = True
                        Exit For
                    End If
                End If
            Next lngRow
            If booExit Then Exit For
        Next lngCol
    End With
    Screen.MousePointer = vbDefault
    Exit Function
Error_LookUpCellTag:
    Screen.MousePointer = vbDefault
   Exit Function
End Function

Private Sub Form_Unload(Cancel As Integer)
    grdSpec.ClearRange 1, 1, grdSpec.MaxCols, grdSpec.MaxRows, False
    ccaSignature.ClearSignature
    Clear_UdtPatient
    Unload Me
    Exit Sub
End Sub


Private Sub grdSpec_Change(ByVal Col As Long, ByVal Row As Long)
Dim lngRow As Long
Dim lngCol As Long
Dim varCellValue
Dim Arr
Dim Idx%
Dim strComment As String
Dim lngCnt As Long

    With grdSpec
        mstrComment = vbNullString
        .Col = 3
        .Row = Row
        If .CellTag <> vbNullString And .CellTag = "COMMENT" Then
            .GetText 3, Row, varCellValue
            If Trim(varCellValue) <> vbNullString Then
                mstrComment = Trim(varCellValue)
                Arr = Split(mstrComment, ".")
                lngRow = Row
                If UBound(Arr) > 0 Then
                    .RowHeight(Row) = 3
                    lngCnt = UBound(Arr)
                    grdSpec.MaxRows = grdSpec.MaxRows + lngCnt
                    For Idx = LBound(Arr) To UBound(Arr)
                        strComment = vbNullString
                        If Trim(Arr(Idx)) <> vbNullString And Len(Arr(Idx)) > 3 Then
                            If Idx > 0 Then
                                strComment = Arr(Idx)
                                lngRow = lngRow + 1
                                grdSpec.InsertRows lngRow, 1
                                Format_CellTypeSplit grdSpec, lngRow, 3, "6", strComment, "TypeHAlignLeft", "COMMENT", False, grdGray, "", 150
                                'grdSpec.SetCellBorder 3, lngRow, lngRow, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleDot
                            Else
                                strComment = Arr(Idx)
                                Format_CellTypeSplit grdSpec, Row, 3, "6", strComment, "TypeHAlignLeft", "COMMENT", False, grdGray, "", 150
                                'grdSpec.SetCellBorder 3, Row, 8, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleDot
                            End If
                        End If
                    Next Idx
                Else
                    Format_CellTypeEdit grdSpec, Row, 3, "6", mstrComment, "TypeHAlignLeft", "COMMENT", grdGray, 150
                    grdSpec.SetCellBorder 3, Row, 8, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleDot
                End If
                'Format_CellTypeSplit
'                Format_CellTypeEdit grdSpec, Row, 3, "6", mstrComment, "TypeHAlignLeft", "COMMENT", grdGray, 150
'                grdSpec.SetCellBorder 3, Row, 8, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleDot
            End If
        End If
    End With
    Exit Sub
End Sub


Private Sub grdSpec_DblClick(ByVal Col As Long, ByVal Row As Long)

    With grdSpec
        .Col = Col
        .Row = Row
        If .CellTag <> vbNullString And .CellTag = "SIGNATURE_MD" Then
            frmMDSign.mstrRef = mstrComment
            frmMDSign.piSignPatID = Val(udtPatientName.IntId)
            frmMDSign.piSignProv = mlngProviderID
            frmMDSign.vstrSource = "Special Order"
            frmMDSign.piSignEnctr = mlngEnctr
            Me.Hide
            frmMDSign.Show vbModal
            Set_Image "SIGNATURE_MD", 1, 1, "4"
        End If
    End With
    Me.Show
    Exit Sub
End Sub


Private Sub lstOrder_Click()
Dim strSQL As String
Dim lngPrimKey As Long
Dim sDate As String
Dim lngRef As Long
Dim strStatus As String
Dim cnLab As New ADODB.Connection
Dim rsLab As New ADODB.Recordset
Dim Idx As Integer
Dim Arr
Dim lngRow As Long
Dim strOrder As String
Dim strComment As String
Dim lngCnt As Long

    On Error GoTo Error_lstOrder_Click
    Screen.MousePointer = vbHourglass
    mlngOrderNo = 0
    For Idx = 0 To lstOrder.listcount - 1
        If lstOrder.List(Idx) = lstOrder.Text Then
            mlngOrderNo = lstOrder.ItemData(Idx)
            Exit For
        End If
    Next Idx
    If cnLab.State = adStateClosed Then cnLab.open psConnect(0)
    If Val(udtPatientName.IntId) <= 0 Then Exit Sub
    strSQL = "Get_Patient_Order_By_Order_ID " & mlngOrderNo
    Set rsLab = cnLab.Execute(strSQL)
    With rsLab
        If Not .EOF Then
            If Not IsNull((!SPECIAL_INSTRUCTIONS)) Then
                If (!SPECIAL_INSTRUCTIONS) <> vbNullString Then mstrComment = (!SPECIAL_INSTRUCTIONS)
            End If
            If Not IsNull((!Encounter_ID)) Then
                If (!Encounter_ID) <> vbNullString Then mlngEnctr = (!Encounter_ID)
            End If
            If Not IsNull(!Provider_ID) Then
                If (!Provider_ID) > 0 Then
                    clsEncounter.ProviderID = (!Provider_ID)
                End If
            End If
        End If
    End With
    For lngRow = 1 To grdSpec.MaxRows
        grdSpec.Col = 3
        grdSpec.Row = lngRow
        If grdSpec.CellTag <> vbNullString And grdSpec.CellTag = "COMMENT" Then
            Format_SpecOrder
            '---------
            Arr = Split(mstrComment, ".")
            If UBound(Arr) > 0 Then
                If UBound(Arr) > 0 Then
                    grdSpec.RowHeight(lngRow) = 3
                    lngCnt = UBound(Arr)
                    grdSpec.MaxRows = grdSpec.MaxRows + lngCnt
                    For Idx = LBound(Arr) To UBound(Arr)
                        strComment = vbNullString
                        If Trim(Arr(Idx)) <> vbNullString And Len(Arr(Idx)) > 3 Then
                            If Idx > 0 Then
                                strComment = Arr(Idx)
                                lngRow = lngRow + 1
                                grdSpec.InsertRows lngRow, 1
                                Format_CellTypeSplit grdSpec, lngRow, 3, "6", strComment, "TypeHAlignLeft", "COMMENT", False, grdGray, "", 150
                                'grdSpec.SetCellBorder 3, lngRow, lngRow, Row, CellBorderIndexOutline, GridBlue, CellBorderStyleDot
                            Else
                                strComment = Arr(Idx)
                                Format_CellTypeSplit grdSpec, lngRow, 3, "6", strComment, "TypeHAlignLeft", "COMMENT", False, grdGray, "", 150
                                'grdSpec.SetCellBorder 3, lngRow, 8, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleDot
                            End If
                        End If
                    Next Idx
                Else
                    Format_CellTypeEdit grdSpec, lngRow, 3, "6", mstrComment, "TypeHAlignLeft", "COMMENT", grdGray, 150
                    grdSpec.SetCellBorder 3, lngRow, 8, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleDot
                End If
            End If
            Exit For
        End If
    Next lngRow
    If cnLab.State = adStateOpen Then
        cnLab.Close
        Set cnLab = Nothing
        Set rsLab = Nothing
    End If
    If mlngOrderNo > 0 Then
        Arr = Split(LookUpCellTag("ORDER_NUMBER"), "|")
        If UBound(Arr) > 0 Then
            lngRow = Arr(1)
        End If
        If lngRow > 1 Then
            strOrder = "PHYSICIAN ORDER" & " " & "ORDER NUMBER:" & " " & mlngOrderNo
            Format_CellTypeEdit grdSpec, lngRow, 1, "9", strOrder, "TypeHAlignCenter", "ORDER_NUMBER", HeaderColor
            grdSpec.SetCellBorder 1, lngRow, 9, lngRow, CellBorderIndexOutline, GridBlue, CellBorderStyleSolid
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_lstOrder_Click:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "lstOrder_Click", App.EXEName & "." & Me.Name
    Exit Sub
    Resume
End Sub


