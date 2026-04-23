VERSION 5.00
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmMedMaint 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medication Maintanence"
   ClientHeight    =   5925
   ClientLeft      =   5325
   ClientTop       =   3000
   ClientWidth     =   11400
   ControlBox      =   0   'False
   Icon            =   "frmMedMaint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   11400
   Begin VB.Frame fraUpdate 
      Caption         =   "Manage Inventory"
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
      Height          =   5880
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   11400
      Begin PT_XP_Button.PT_XP cmdAddUpdate 
         Height          =   495
         Left            =   9510
         TabIndex        =   37
         Top             =   1620
         Width           =   1755
         _ExtentX        =   3096
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
         Picture         =   "frmMedMaint.frx":058A
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Add Update Medication"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMedMaint.frx":0B24
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdSearch 
         Height          =   405
         Left            =   4980
         TabIndex        =   32
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
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
         Picture         =   "frmMedMaint.frx":0E76
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Search"
         DepthEvent      =   1
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMedMaint.frx":1410
         ShowFocus       =   -1  'True
      End
      Begin VB.TextBox cctSearch 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   31
         Top             =   330
         Width           =   4665
      End
      Begin VB.ListBox lstDrugs 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1260
         ItemData        =   "frmMedMaint.frx":1762
         Left            =   120
         List            =   "frmMedMaint.frx":1769
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   30
         Top             =   750
         Width           =   6375
      End
      Begin VB.Frame fraMedDetails 
         Height          =   2955
         Left            =   120
         TabIndex        =   2
         Top             =   2100
         Visible         =   0   'False
         Width           =   11205
         Begin VB.TextBox txtProdId 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1650
            TabIndex        =   38
            Top             =   840
            Width           =   3825
         End
         Begin VB.TextBox txtStrength 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6870
            TabIndex        =   35
            Top             =   150
            Width           =   4215
         End
         Begin VB.ComboBox cboUnitCost 
            Height          =   315
            Left            =   4530
            TabIndex        =   28
            Top             =   2550
            Width           =   1605
         End
         Begin VB.TextBox txtNDCnumber 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1650
            MaxLength       =   24
            TabIndex        =   15
            Top             =   480
            Width           =   3825
         End
         Begin VB.Frame fraType 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   5670
            TabIndex        =   11
            Top             =   510
            Width           =   3855
            Begin VB.OptionButton optBrand 
               Caption         =   "Brand"
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
               Left            =   1200
               TabIndex        =   13
               Top             =   0
               Width           =   1065
            End
            Begin VB.OptionButton optGeneric 
               Caption         =   "Generic"
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
               Left            =   2340
               TabIndex        =   12
               Top             =   30
               Width           =   1215
            End
            Begin VB.Label lblType 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Type"
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
               Height          =   195
               Left            =   -840
               TabIndex        =   14
               Top             =   30
               Width           =   1395
            End
         End
         Begin VB.Frame fraDefaultValues 
            Caption         =   "Default Values"
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
            Height          =   1695
            Left            =   6210
            TabIndex        =   4
            Top             =   1200
            Width           =   4935
            Begin VB.TextBox txtRoute 
               Height          =   285
               Left            =   1050
               TabIndex        =   34
               Top             =   960
               Width           =   3405
            End
            Begin VB.TextBox txtMedUnit 
               Height          =   345
               Left            =   1050
               TabIndex        =   33
               Top             =   570
               Width           =   3405
            End
            Begin VB.TextBox txtRemarks 
               Height          =   285
               Left            =   1050
               MaxLength       =   512
               TabIndex        =   5
               Top             =   1260
               Width           =   3405
            End
            Begin VB.PictureBox ccnDosage 
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   990
               ScaleHeight     =   195
               ScaleWidth      =   915
               TabIndex        =   6
               Top             =   255
               Width           =   975
            End
            Begin VB.Label lblRemarks 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   " Remarks:"
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
               Left            =   60
               TabIndex        =   10
               Top             =   1335
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Route:"
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
               TabIndex        =   9
               Top             =   1020
               Width           =   630
            End
            Begin VB.Label lblUnit 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit:"
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
               TabIndex        =   8
               Top             =   660
               Width           =   495
            End
            Begin VB.Label lblDosage 
               BackStyle       =   0  'Transparent
               Caption         =   "Dosage:"
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
               Height          =   195
               Left            =   150
               TabIndex        =   7
               Top             =   315
               Width           =   675
            End
         End
         Begin VB.TextBox txtMedName 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1650
            MaxLength       =   64
            TabIndex        =   3
            Top             =   120
            Width           =   3825
         End
         Begin VB.PictureBox ccnPackageSize 
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1830
            ScaleHeight     =   195
            ScaleWidth      =   1395
            TabIndex        =   16
            Top             =   1800
            Width           =   1455
         End
         Begin VB.PictureBox ccnUnitAWP 
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4515
            ScaleHeight     =   195
            ScaleWidth      =   1395
            TabIndex        =   17
            Top             =   1800
            Width           =   1455
         End
         Begin VB.PictureBox ccnAWPPerPackage 
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1830
            ScaleHeight     =   195
            ScaleWidth      =   1395
            TabIndex        =   18
            Top             =   2160
            Width           =   1455
         End
         Begin VB.PictureBox ccnUnitCost 
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4515
            ScaleHeight     =   195
            ScaleWidth      =   1395
            TabIndex        =   19
            Top             =   2160
            Width           =   1455
         End
         Begin VB.PictureBox ccnCostUnit 
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1830
            ScaleHeight     =   195
            ScaleWidth      =   1395
            TabIndex        =   27
            Top             =   2550
            Width           =   1455
         End
         Begin VB.Label lblProdID 
            Caption         =   "Our Product ID"
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
            Height          =   225
            Left            =   60
            TabIndex        =   39
            Top             =   900
            Width           =   1455
         End
         Begin VB.Label lblStrength 
            Caption         =   "Strength"
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
            Height          =   225
            Left            =   5730
            TabIndex        =   36
            Top             =   180
            Width           =   975
         End
         Begin VB.Label lblCost 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit:"
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
            Height          =   195
            Left            =   3570
            TabIndex        =   29
            Top             =   2610
            Width           =   795
         End
         Begin VB.Label lblCostUnit 
            BackStyle       =   0  'Transparent
            Caption         =   "Cost Per Unit:"
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
            Height          =   195
            Left            =   60
            TabIndex        =   26
            Top             =   2610
            Width           =   1515
         End
         Begin VB.Label lblMed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Medication:"
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
            Left            =   60
            TabIndex        =   25
            Top             =   210
            Width           =   1140
         End
         Begin VB.Label lblNDCNumber 
            BackStyle       =   0  'Transparent
            Caption         =   "NDC Number:"
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
            Height          =   195
            Left            =   60
            TabIndex        =   24
            Top             =   540
            Width           =   1395
         End
         Begin VB.Label lblPackageSize 
            BackStyle       =   0  'Transparent
            Caption         =   "Package Size:"
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
            Height          =   195
            Left            =   60
            TabIndex        =   23
            Top             =   1860
            Width           =   1395
         End
         Begin VB.Label lblUnitAwp 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit AWP:"
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
            Height          =   195
            Left            =   3600
            TabIndex        =   22
            Top             =   1860
            Width           =   795
         End
         Begin VB.Label lblAWPPerPackage 
            BackStyle       =   0  'Transparent
            Caption         =   "AWP Per Package:"
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
            Height          =   195
            Left            =   60
            TabIndex        =   21
            Top             =   2220
            Width           =   2145
         End
         Begin VB.Label lblUnitCost 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Cost:"
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
            Height          =   195
            Left            =   3570
            TabIndex        =   20
            Top             =   2220
            Width           =   795
         End
      End
      Begin PT_XP_Button.PT_XP cmdDoneDefault 
         Height          =   465
         Left            =   5040
         TabIndex        =   1
         Top             =   5250
         Width           =   1005
         _ExtentX        =   1773
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
         Picture         =   "frmMedMaint.frx":1777
         AppearanceThemes=   3
         TransparentColor=   14737632
         Caption         =   "Done"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmMedMaint.frx":1B11
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMedMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fConnectCR As String
Public fConnectTM As String

Private bAccept As Boolean
Private mType As String

Private Sub Clear_Fields()
    lstDrugs.Clear
    txtMedName.Text = vbNullString
    txtNDCnumber.Text = vbNullString
    optBrand.Value = False
    ccnPackageSize.Text = vbNullString
    ccnAWPPerPackage.Text = vbNullString
    ccnUnitAWP.Text = vbNullString
    ccnUnitCost.Text = vbNullString
    ccnDosage.Text = vbNullString
    txtStrength.Text = vbNullString
    txtRemarks.Text = vbNullString
    cboUnitCost.Text = vbNullString
    txtProdId.Text = vbNullString
    Exit Sub
End Sub

Public Sub Reorder_ListBox(lst As Control, bUp As Boolean)
Dim i As Integer
Dim sContent As String
Dim lData As Long
Dim X As Integer

    If bUp Then
        X = 1
    Else
        X = -1
    End If

    i = lst.ListIndex
    sContent = lst.Text
    lData = lst.ItemData(i)
   
   If (i > 0 And bUp) Or (i < lst.ListCount - 1 And Not bUp) Then
    lst.RemoveItem i
    lst.AddItem sContent, i - X
    lst.ItemData(i - X) = lData
    lst.Selected(i - X) = True
   End If
End Sub

Private Sub cboRefills_Click()
    bAccept = True
End Sub

Private Sub ccnAWPPerPackage_AfterUpdate()
    bAccept = True
End Sub

Private Sub ccnDosage_AfterUpdate()
    bAccept = True
End Sub

Private Sub ccnPackageSize_AfterUpdate()
    bAccept = True
End Sub


Private Sub ccnUnitAWP_AfterUpdate()
    bAccept = True
End Sub

Private Sub ccnUnitCost_AfterUpdate()
    bAccept = True
End Sub
Public Sub ReadData()
'    Me.top = (Screen.Height - Me.Height) / 2 + 300
'    Me.left = (Screen.Width - Me.Width) / 2
'
'    LoadListBox "Select med_id,medication from medications order by medication"
'
'    LoadComboBox cboUnit, "sd_med_dose_unit"
'    LoadComboBox cboRoute, "sd_med_route"
'    LoadComboBox cboRefills, "sd_med_refills"
End Sub

Private Sub cmdAddUpdate_Click()
Dim strSQL As String
Dim booRet As Boolean
Dim sMedType As String
Dim sPackageSize As String
Dim sUnitAWP As String
Dim sAWPperPackage As String
Dim sUnitCost As String
Dim sCostUnit As String
Dim sDosage As String
Dim iMedId As Integer
Dim strStrength As String
Dim booinsert As Boolean
Dim strDrug As String
Dim lngMedID As Long

    fraMedDetails.Visible = False
    sMedType = "Brand"
    If optGeneric.Value = True Then sMedType = "Generic"
    
    sPackageSize = Trim(ccnPackageSize.Text)
    sUnitAWP = Trim(ccnUnitAWP.Text)
    sAWPperPackage = Trim(ccnAWPPerPackage.Text)
    sUnitCost = Trim(ccnUnitCost.Text)
    sDosage = Trim(ccnDosage.Text)
    strStrength = Trim(txtStrength.Text)
    sCostUnit = Trim(cboUnitCost.Text)
    strDrug = Trim(txtMedName.Text)
    
    'If Len(Trim(ccnPackageSize.Text)) = 0 Then sPackageSize = "Null"
    'If Len(Trim(ccnUnitAWP.Text)) = 0 Then sUnitAWP = "Null"
    'If Len(Trim(ccnAWPPerPackage.Text)) = 0 Then sAWPperPackage = "Null"
   ' If Len(Trim(ccnUnitCost.Text)) = 0 Then sUnitCost = "Null"
    'If Len(Trim(ccnDosage.Text)) = 0 Then sDosage = "Null"
    
    If txtMedName.Text = vbNullString Then
        MsgBox "Please add medication first!", vbInformation
        Exit Sub
    End If
    ' Insert or Update
     Open_DB psConnect(1)
     strSQL = "Select Med_ID from Medication where Medication = '" & strDrug & "'"
    Set rs = cn.Execute(strSQL)
    booinsert = True
    With rs
        If Not .EOF Then
            If Not IsNull(!Med_ID) Then booinsert = False
            lngMedID = (!Med_ID)
        End If
    End With
    
    'txtNDCnumber
    If booinsert = False Then
       strSQL = "Update medications set " & _
                "medication = '" & txtMedName.Text & "'," & _
                "ndc_number = '" & txtNDCnumber.Text & "'," & _
                "med_type = '" & sMedType & "'," & _
                "package_size = " & sPackageSize & "," & _
                "unit_awp = " & sUnitAWP & "," & _
                "awp_per_package = " & sAWPperPackage & "," & _
                "unit_cost = " & sUnitCost & "," & _
                "Cost_Unit = " & sCostUnit & "," & _
                "dosage = " & sDosage & "," & _
                "units = '" & txtMedUnit.Text & "'," & _
                "remarks = '" & txtRemarks.Text & "'," & _
                "Strength = " & strStrength & "', " & _
                "route = '" & txtRoute & " " & _
                "Where med_id = " & lngMedID
    Else
       strSQL = "Insert into medication (medication,ndc_number,med_type,package_size,unit_awp, " _
                    & "awp_per_package,unit_cost, Cost_Unit,dosage, units,route, remarks,Strength) " _
                    & "values ('" & txtMedName.Text & "','" & txtNDCnumber.Text & "','" _
                    & sMedType & "','" & sPackageSize & "','" & sUnitAWP & "','" & sAWPperPackage & "','" _
                    & sUnitCost & "','" & sCostUnit & "','" & sDosage & "','" & txtMedUnit.Text & "','" _
                    & txtRoute.Text & "','" _
                    & txtRemarks.Text & "','" _
                    & strStrength & "')"
    End If
    
    Set rs = cn.Execute(strSQL)
     Open_DB fConnectTM
    MsgBox txtMedName.Text & " " & "successfully added to Medication Inventory!", vbInformation
    Clear_Fields
Exit Sub
End Sub

Private Sub cmdDoneDefault_Click()
    Dim iAnswer As Integer
    Dim bReturn As Boolean
    Clear_Fields
    Unload Me
End Sub
Public Sub LoadComboBox(cbo As Object, ByVal vstrSubDomain As String)

    Open_DB fConnectTM
   
'    'Call stored procedure to retrieve data values for combo
'    cn.S_ATT_BT_ATTRBTE_VALID_VALUE "S", vstrSubDomain, rs
'
'    cbo.Clear
'
'    With rs
'        Do Until .EOF
'            cbo.AddItem (!ATTRBTE_VALID_VALUE)
'            .MoveNext
'        Loop
'        cbo.ListIndex = -1
'
'        rs.Close
    'End With
    
    Close_DB fConnectTM
    Exit Sub
End Sub

Private Sub cmdV_Click()
    LoadListBox "select med_id,medication from medications where upper(medication) like 'V%' order by medication"
End Sub

Private Sub cmdW_Click()
    LoadListBox "select med_id,medication from medications where upper(medication) like 'W%' order by medication"
End Sub

Private Sub cmdX_Click()
    LoadListBox "select med_id,medication from medications where upper(medication) like 'X%' order by medication"
End Sub

Private Sub cmdY_Click()
    LoadListBox "select med_id,medication from medications where upper(medication) like 'Y%' order by medication"
End Sub

Private Sub cmdZ_Click()
    LoadListBox "select med_id,medication from medications where upper(medication) like 'Z%' order by medication"
End Sub

Private Sub cmdSearch_Click()
Dim strSQL As String
Dim strText As String
Dim sStr As String


    Clear_Fields
    strText = Trim(cctSearch.Text)
    Open_DB psConnect(1)
    strSQL = "SELECT * FROM Medication where Medication like '" & strText & "'"
    Set rs = cn.Execute(strSQL)
    With rs
       Do Until .EOF
          sStr = Trim((!Medication))
          lstDrugs.AddItem sStr
          lstDrugs.ItemData(lstDrugs.NewIndex) = (!LISTING_SEQ_NO)
          .MoveNext
       Loop
       .Close
    End With
  
    If lstDrugs.ListCount <= 0 Then
        strSQL = "SELECT TRADENAME, LISTING_SEQ_NO FROM ndx_listing where tradename like '" & strText & "%'"
        Set rs = cn.Execute(strSQL)
        With rs
           Do Until .EOF
              sStr = Trim(!tradename)
              lstDrugs.AddItem sStr
              lstDrugs.ItemData(lstDrugs.NewIndex) = (!LISTING_SEQ_NO)
              .MoveNext
           Loop
           .Close
        End With
    End If
    If lstDrugs.ListCount <= 0 Then
        MsgBox strText & " " & "not found. Enter new medication information; if desired.", vbInformation
        fraMedDetails.Visible = True
    End If
     
    Close_DB psConnect(1)
    Screen.MousePointer = vbDefault
    
    Exit Sub
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


Private Sub lstDrugs_Click()
Dim i%
Dim strDrug As String
Dim lngListSeq As String
Dim strSQL As String

    fraMedDetails.Visible = True
    For i = 0 To lstDrugs.ListCount - 1
        If lstDrugs.Selected(i) Then
            strDrug = lstDrugs.List(i)
            lngListSeq = lstDrugs.ItemData(i)
        End If
    Next i
    Open_DB psConnect(1)
    strSQL = "Select * from Medication where Medication = '" & strDrug & "'"
    Set rs = cn.Execute(strSQL)
    With rs
        If Not .EOF Then
              txtMedName.Text = (!tradename)
              txtNDCnumber.Text = (!LISTING_SEQ_NO)
              ccnPackageSize.Text = (!Package_Size)
              If (!Med_Type) = "Generic" Then optGeneric.Value = True
              If (!Med_Type) = "Brand" Then optBrand.Value = True
              ccnAWPPerPackage.Text = (!AWP_Per_Package)
              ccnUnitAWP.Text = (!Unit_AWP)
              ccnUnitCost.Text = (!Unit_Cost)
              cboUnitCost.Text = (!Cost_Unit)
              ccnDosage.Text = (!Dosage)
              txtMedUnit.Text = (!Units)
              txtRoute.Text = (!Route)
              txtRemarks.Text = (!Remarks)
              txtStrength.Text = (!Strength)
        End If
    End With
    
    If txtMedName.Text = vbNullString Then
    strSQL = "Select * from ndx_listing a, NDX_PACKAGES b where a.LISTING_SEQ_NO =  " & lngListSeq & " And a.listing_seq_no = b.listing_seq_no"
    Set rs = cn.Execute(strSQL)
    With rs
        If Not .EOF Then
              txtMedName.Text = (!tradename)
              txtNDCnumber.Text = (!LISTING_SEQ_NO)
              ccnPackageSize.Text = (!PACKSIZE)
             ' If (!Med_Type) = "Generic" Then optGeneric.Value = True
             ' If (!Med_Type) = "Brand" Then optBrand.Value = True
             ' ccnAWPPerPackage.Text = (!AWP_Per_Package)
             ' ccnUnitAWP.Text = (!Unit_AWP)
             ' ccnUnitCost.Text = (!Unit_Cost)
              cboUnitCost.Text = (!UNIT)
              ccnDosage.Text = (!DOSAGE_FORM)
              txtMedUnit.Text = (!Units)
              txtRoute.Text = (!Route)
             ' txtRemarks.Text = (!Remarks)
              txtStrength.Text = (!Strength)
        End If
    End With
    
    End If
    
    
    Close_DB psConnect(1)
    Exit Sub
End Sub


Private Sub optBrand_Click()
    bAccept = True
End Sub

Private Sub optGeneric_Click()
    bAccept = True
End Sub


Private Sub txtNDCnumber_Change()
    bAccept = True
End Sub

Private Sub txtRemarks_Change()
    bAccept = True
End Sub
Public Sub LoadListBox(pSQL As String)
'    Dim varRS
'    Dim booret As Boolean
'    Dim lngRows As Long
'
'    Dim r As Integer
'
'    lstTerms.Clear
'
'    booret = SelectQuery(pSQL, fConnectCR, varRS, lngRows)
'
'    If booret Then
'        If lngRows > 0 Then
'            For r = LBound(varRS, 2) To UBound(varRS, 2)
'                lstTerms.AddItem varRS(1, r)
'                lstTerms.ItemData(lstTerms.NewIndex) = varRS(0, r)
'            Next r
'        End If
'    End If
End Sub


