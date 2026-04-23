VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{C08D2A54-E5CB-4AFE-8AAD-B3E08A43486C}#1.3#0"; "XButton.ocx"
Begin VB.Form frmAdmin 
   ClientHeight    =   9210
   ClientLeft      =   2100
   ClientTop       =   2820
   ClientWidth     =   17745
   Icon            =   "frmAdmin.frx":0000
   ScaleHeight     =   9210
   ScaleWidth      =   17745
   Begin VB.Frame fraInitalTelephone 
      BackColor       =   &H00E0E0E0&
      Height          =   9135
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   17745
      Begin FPSpreadADO.fpSpread grdPatient 
         Height          =   5595
         Left            =   150
         TabIndex        =   6
         Top             =   2430
         Visible         =   0   'False
         Width           =   17265
         _Version        =   458752
         _ExtentX        =   30454
         _ExtentY        =   9869
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowColor     =   16761024
         ShadowDark      =   16761024
         ShadowText      =   8388608
         SpreadDesigner  =   "frmAdmin.frx":0ECA
         Appearance      =   2
      End
      Begin VB.TextBox txtPatInfo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   885
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1440
         Width           =   4215
      End
      Begin PT_XP_Button.PT_XP cmdNewEncrt 
         Height          =   435
         Left            =   15870
         TabIndex        =   19
         Top             =   1710
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   767
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
         Picture         =   "frmAdmin.frx":10C2
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "New Visit"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAdmin.frx":165C
         PictureDown     =   "frmAdmin.frx":19AE
         ShowFocus       =   -1  'True
      End
      Begin VB.TextBox txtEnctr 
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
         Height          =   345
         Left            =   11910
         TabIndex        =   17
         Top             =   900
         Width           =   2865
      End
      Begin VB.TextBox txtEMRID 
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
         Height          =   345
         Left            =   6150
         TabIndex        =   15
         Top             =   930
         Width           =   1725
      End
      Begin PT_XP_Button.PT_XP cmdTelephone 
         Height          =   645
         Left            =   1980
         TabIndex        =   13
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1138
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
         Picture         =   "frmAdmin.frx":1F48
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "Telephone Notes"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAdmin.frx":2822
         ShowFocus       =   -1  'True
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   4215
      End
      Begin PT_XP_Button.PT_XP cmdMessaging 
         Height          =   705
         Left            =   15270
         TabIndex        =   7
         Top             =   150
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1244
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
         Picture         =   "frmAdmin.frx":3474
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Messaging"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAdmin.frx":434E
         ShowFocus       =   -1  'True
      End
      Begin FPSpreadADO.fpSpreadPreview fpSpreadPreview1 
         Height          =   1545
         Left            =   18420
         TabIndex        =   5
         Top             =   3090
         Width           =   4125
         _Version        =   458752
         _ExtentX        =   7276
         _ExtentY        =   2725
         _StockProps     =   96
         BorderStyle     =   1
         AllowUserZoom   =   -1  'True
         GrayAreaColor   =   8421504
         GrayAreaMarginH =   720
         GrayAreaMarginType=   0
         GrayAreaMarginV =   720
         PageBorderColor =   8388608
         PageBorderWidth =   2
         PageShadowColor =   0
         PageShadowWidth =   2
         PageViewPercentage=   100
         PageViewType    =   0
         ScrollBarH      =   1
         ScrollBarV      =   1
         ScrollIncH      =   360
         ScrollIncV      =   360
         PageMultiCntH   =   1
         PageMultiCntV   =   1
         PageGutterH     =   -1
         PageGutterV     =   -1
         ScriptEnhanced  =   0   'False
      End
      Begin PT_XP_Button.PT_XP cmdUpdatePatient 
         Height          =   465
         Left            =   8130
         TabIndex        =   4
         Top             =   870
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   820
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
         Picture         =   "frmAdmin.frx":5EA0
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Update Patient Information"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAdmin.frx":643A
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdActionItems 
         Height          =   615
         Left            =   13260
         TabIndex        =   3
         Top             =   150
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1085
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
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "My Action Items"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdNewPatient 
         Height          =   645
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1138
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
         Picture         =   "frmAdmin.frx":678C
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "New Patient"
         CaptionAlignment=   5
         CaptionPosition =   4
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAdmin.frx":7066
         PictureOver     =   "frmAdmin.frx":7CB8
         PictureOverEffect=   3
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdExistingPatient 
         Height          =   645
         Left            =   3840
         TabIndex        =   2
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1138
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
         Picture         =   "frmAdmin.frx":890A
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   " Existing  Patient"
         CaptionAlignment=   5
         CaptionPosition =   4
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAdmin.frx":91E4
         PictureOver     =   "frmAdmin.frx":9E36
         PictureOverEffect=   3
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSelect 
         Height          =   495
         Left            =   7950
         TabIndex        =   16
         Top             =   8460
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
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
         Picture         =   "frmAdmin.frx":AA88
         AppearanceThemes=   3
         TransparentColor=   13160660
         Caption         =   "Select Patient/Visit"
         CaptionAlignment=   5
         CaptionPosition =   4
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmAdmin.frx":B022
         ShowFocus       =   -1  'True
      End
      Begin XBtn.XButton cmdAdminHistory 
         Height          =   405
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BackColor       =   12640511
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4210816
         Bevel           =   6
         Caption         =   "Admission History"
         Picture         =   "frmAdmin.frx":B374
         PictAlign       =   4
         CaptionAlign    =   4
         PushTimer       =   0   'False
         PushInterval    =   300
         MousePointer    =   0
         PicRaster       =   12583114
      End
      Begin VB.Frame fraPatient 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   6165
         Left            =   1680
         TabIndex        =   9
         Top             =   2280
         Visible         =   0   'False
         Width           =   14775
         Begin PT_XP_Button.PT_XP cmdClose 
            Height          =   405
            Left            =   7410
            TabIndex        =   12
            Top             =   5640
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   714
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
            Picture         =   "frmAdmin.frx":B390
            AppearanceThemes=   3
            TransparentColor=   13160660
            Caption         =   "Close"
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAdmin.frx":BDA2
            ShowFocus       =   -1  'True
         End
         Begin PT_XP_Button.PT_XP cmdSave 
            Height          =   405
            Left            =   6000
            TabIndex        =   11
            Top             =   5640
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
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
            Picture         =   "frmAdmin.frx":C0F4
            AppearanceThemes=   3
            TransparentColor=   13160660
            Caption         =   "Save"
            DepthEvent      =   1
            ForeColorDisabled=   16761024
            ForeColorOver   =   8388608
            ForeColorFocus  =   8388608
            ForeColorDown   =   8388608
            PictureDisabled =   "frmAdmin.frx":C48E
            ShowFocus       =   -1  'True
         End
         Begin PatientTrac.PatNewContact PatNewContact1 
            Height          =   5535
            Left            =   0
            TabIndex        =   10
            Top             =   60
            Width           =   14745
            _ExtentX        =   26009
            _ExtentY        =   9763
         End
      End
      Begin VB.Label lblEnctr 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Selected Visit"
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
         Height          =   285
         Left            =   10470
         TabIndex        =   18
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label lblEMRID 
         BackColor       =   &H00E0E0E0&
         Caption         =   "EMR Record No."
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
         Left            =   4470
         TabIndex        =   14
         Top             =   990
         Width           =   1515
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11820
      Top             =   5820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdmin.frx":C7E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdmin.frx":D6BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdmin.frx":E0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdmin.frx":EAAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdmin.frx":F4A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdmin.frx":FEA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdmin.frx":10D7C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrConnectStringClinical As String
Public mstrConnectStringTerms As String
Public mlngPatientID As Long
Public mlngSelPatient As Long
Public mstrSelectedTerms As String
Public mstrListItemData As String
Private mRow As Integer
Private mintGridRow As Integer
Private mintGridCol As Integer

Private Sub Clear_Fields()
    txtSearch.Text = vbNullString
    txtPatInfo.Text = vbNullString
    txtEMRID.Text = vbNullString
    txtEnctr.Text = vbNullString
    
    With grdPatient
        .ClearRange 1, 0, .MaxCols, .MaxRows, True
        mRow = 0
        .MaxRows = mRow
        .Col = 1
        .Row = mRow
    End With
    
End Sub

Private Sub Set_RowHeight()
Dim r
Dim c
Dim varNote
Dim CellNote As Long

    grdPatient.Col = 1
    grdPatient.Row = 1
    
    grdPatient.Redraw = False
    
    With grdPatient

    For r = 1 To .MaxRows
        For c = 1 To .MaxCols
            .GetText c, r, varNote
            If varNote <> vbNullString Then
            .BlockMode = True
                .Col = c
                .Col2 = c
                .Row = r
                .Row2 = r
                .TypeEditMultiLine = True   'Make multi-line
                .TypeTextWordWrap = True  'Make multi-line
                .TypeMaxEditLen = 20000
                .TypeVAlign = TypeVAlignCenter
                .TypeHAlign = TypeHAlignLeft
                .Text = varNote
                CellNote = .MaxTextCellHeight
                If .ROWHEIGHT(r) < CellNote Then
                     .ROWHEIGHT(r) = CellNote
                End If
                .SetCellBorder c, r, .MaxCols, r, 16, &H80&, CellBorderStyleSolid
                '.SetCellBorder c, r, .MaxCols, r, 16, &H80&, CellBorderStyleSolid
               ' .SetCellBorder c, r, .MaxCols, r, 1, &H80&, CellBorderStyleSolid
                .TypeHAlign = TypeHAlignLeft
                .TypeVAlign = TypeVAlignCenter
                .FontName = "Georgia"
            .BlockMode = False
            End If
        Next c
    Next r
    grdPatient.Redraw = True
    End With
Exit Sub
End Sub


Private Function Read_Data()
    Get_Collection piPatientID
End Function

Private Sub cmdClose_Click()
    fraPatient.Visible = False
End Sub

Sub Patient_Grid(vlngPatientID As Long)
Dim lngRow As Long
Dim lngCol As Long
Dim mfpNote As String
Dim sPatientName As String
Dim strPlan As String
Dim lngLastEncounter As Long
Dim vartemp
Dim r
Dim strSQL As String
Dim lngencounterid As Long
Dim i%
Dim varCellValue As Variant

    On Error GoTo Error_AddRecord
    Get_Collection piPatientID
    
    With grdPatient
         .Redraw = False
         .ProcessTab = True
        ' .Width = 16000
         .FontBold = True
         .FontName = "Georgia"
         .GridSolid = False
         .GridShowHoriz = False
         .GridShowVert = False
         .RowHeadersShow = False
         .ColHeadersShow = True
         .ColHeaderRows = 1
         .MaxCols = 7
         .ColHeadersAutoText = DispBlank
         .TextTip = TextTipFloatingFocusOnly
         .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
         .ColWidth(1) = 7
         .ColWidth(2) = 20
         .ColWidth(3) = 20
         .ColWidth(4) = 25
         .ColWidth(5) = 25
         .ColWidth(6) = 25
         .ColWidth(7) = 20
         
         mRow = 0
        .MaxRows = mRow
        .ClearRange 1, mRow, .MaxCols, .MaxRows, True
        txtPatInfo.Text = vbNullString
    '    'Build the cell note
        txtPatInfo.Text = " Date of Birth: " & clsPatient.DOB & vbNewLine
        txtPatInfo.Text = txtPatInfo.Text & " Phone: " & clsPatient.Phone & vbNewLine
        txtPatInfo.Text = txtPatInfo.Text & "Address" & " " & clsPatient.Address & clsPatient.Address2 & vbNewLine
        txtPatInfo.Text = txtPatInfo.Text & " " & clsPatient.CityStateZip & vbNewLine
        txtPatInfo.Text = txtPatInfo.Text & "Social Security" & " " & clsPatient.SS & vbNewLine

    ' set row headers
        ' Set Headers
        .SetText 1, 0, "Encounter ID"
        .SetText 2, 0, "Visit Date"
        .SetText 3, 0, "Next Appointment"
        .SetText 4, 0, "Reason for Visit"
        .SetText 5, 0, "Visit Type"
        .SetText 6, 0, "Procedure Type"
        .SetText 7, 0, "Procedure Status"
    End With
        Open_DB psConnect(0)
        
        strSQL = "Select Encounter_ID, encntr_start_date, APPOINTMENT_DATE, visit_reason, visit_type " _
                & "FROM ENCOUNTER WHERE Patient_ID = " & vlngPatientID
        
        Set rs = cn.Execute(strSQL)
        
        With rs
            Do Until .EOF
                mRow = mRow + 1
                grdPatient.Row = mRow
                grdPatient.MaxRows = mRow
                grdPatient.SetText 1, mRow, (!ENCOUNTER_ID)
                If (!ENCOUNTER_ID) = piEncounterID Then
                    grdPatient_Click 1, mRow
                End If
                grdPatient.SetText 2, mRow, (!ENCNTR_START_DATE)
                grdPatient.SetText 3, mRow, (!APPOINTMENT_DATE)
                grdPatient.SetText 4, mRow, (!VISIT_REASON)
                grdPatient.SetText 5, mRow, (!VISIT_TYPE)
                grdPatient.SetCellBorder 1, mRow, grdPatient.MaxCols, mRow, 16, &H80&, CellBorderStyleSolid
                grdPatient.SetCellBorder 1, mRow, grdPatient.MaxCols, mRow, 1, &H80&, CellBorderStyleSolid
            .MoveNext
            Loop
        End With
        
        For i = 1 To grdPatient.MaxRows
            grdPatient.GetText 1, i, varCellValue
            lngencounterid = Val(varCellValue)
            strSQL = "Select COSMETIC_PROCEDURE ,  Status,  ORDER_TYPE from COSMETIC_ORDERS where encounter_Id = " & lngencounterid
            Set rs = cn.Execute(strSQL)
            With rs
                If Not .EOF Then
                    grdPatient.SetText 6, i, (!cosmetic_procedure)
                    grdPatient.SetText 7, i, (!Status)
                End If
            End With
        Next i
        Close_DB psConnect(0)
        Set_RowHeight

Exit Sub
Error_AddRecord:
        Screen.MousePointer = vbDefault
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        Exit Sub
        Resume
End Sub


Private Sub cmdExistingPatient_Click()
    Clear_Fields
    frmListPatient.Read_PatientData
    frmListPatient.Show vbModal
    If frmListPatient.mlngListItemData <> 0 Then
        piPatientID = frmListPatient.mlngListItemData
        grdPatient.ZOrder 0
        grdPatient.Visible = True
        Patient_Grid piPatientID
        Get_Patient piPatientID
        txtSearch.Text = clsPatient.Name
        txtEMRID.Text = clsPatient.IntId
    Else
        MsgBox "Please select a patient first!", vbInformation
    End If
    
   cmdSelect.Visible = True
   Exit Sub
End Sub


Private Sub cmdMessaging_Click()
    frmPatientAppointment.Show vbModal
End Sub


Private Sub cmdNewEncrt_Click()
    frmNewEnctr.Show vbModal
    Get_Encounter piEncounterID
    
End Sub

Private Sub cmdNewPatient_Click()
   fraPatient.Visible = True
   grdPatient.Visible = False
   grdPatient.ZOrder 1
   cmdSelect.Visible = False
   Exit Sub
End Sub

Private Sub cmdSelect_Click()
    
    If piPatientID = 0 Then
        piPatientID = modSpecific.piPatientID
        If piPatientID = 0 Then
            MsgBox "Please select a Patient!", vbInformation
            Exit Sub
        End If
    End If
    If piEncounterID = 0 Then
        piEncounterID = modSpecific.piEncounterID
        If piEncounterID = 0 Then
            MsgBox "Please select an encounter! Select an encounter by clicking on desired row!", vbInformation
            Exit Sub
        End If
    End If
    Unload Me
    Exit Sub
    
End Sub


Private Sub cmdTelephone_Click()
    frmTelephone.Show vbModal
    Exit Sub
End Sub




Private Sub cmdSave_Click()
    PatNewContact1.Write_Inquiry
    fraPatient.Visible = False
End Sub

Private Sub cmdUpdatePatient_Click()
    PatientInfo.piProviderID = piProviderID
    If mlngSelPatient = 0 Then
        PatientInfo.piPatientID = piPatientID
    Else
        PatientInfo.piPatientID = mlngSelPatient
    End If
    PatientInfo.mNewPatient = False
    PatientInfo.Show vbModal
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


Private Sub grdPatient_Click(ByVal Col As Long, ByVal Row As Long)
Dim varCellValue As Variant
Dim i%

    mintGridRow = Row
    mintGridCol = Col
    
    'Undo the Backcolor
    With grdPatient
       For i = 1 To .MaxRows
        .BlockMode = True
        .Col = 1
        .Col2 = .MaxCols
        .Row = i
        .Row2 = i
        .BackColor = 16777215
       Next i
    End With
    
    
    If mintGridRow < 0 Then Exit Sub
    With grdPatient
        .GetText 1, mintGridRow, varCellValue
        If varCellValue <> vbNullString Then
            piEncounterID = varCellValue
            modSpecific.piEncounterID = piEncounterID
            .GetText 2, mintGridRow, varCellValue
            .BlockMode = True
            .Col = 1
            .Col2 = .MaxCols
            .Row = mintGridRow
            .Row2 = mintGridRow
            .BackColor = &HFFC0C0
            .BlockMode = False
            txtEnctr.Text = vbNullString
            txtEnctr.Text = varCellValue
        End If
    End With
    
    Exit Sub
End Sub


Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)


    frmListPatient.Show vbModal
    mstrSelectedTerms = frmListPatient.mstrSelectedTerms
    mstrListItemData = frmListPatient.mlngListItemData
    
End Sub


