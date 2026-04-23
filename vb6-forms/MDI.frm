VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   " PatientTrac EMR"
   ClientHeight    =   10515
   ClientLeft      =   4635
   ClientTop       =   2715
   ClientWidth     =   19050
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   210
      Top             =   8850
   End
   Begin MSComctlLib.ImageList PatientTracImages 
      Left            =   12810
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   37
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":2264
            Key             =   "Dashboard"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":2B3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":3418
            Key             =   "Consent"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":3CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":45CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":4EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":5780
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":605A
            Key             =   "Viewer"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6934
            Key             =   "SuperBill"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":720E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":7AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":83C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":8C9C
            Key             =   "Select Patient"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":9576
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":9E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A72A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B004
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B8DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":C1B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":CA92
            Key             =   "Reports"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":D36C
            Key             =   "Psychiatry"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":DC46
            Key             =   "StimSheet	"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":E520
            Key             =   "Surgery Log"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":EDFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":F6D4
            Key             =   "Appointments"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":FFAE
            Key             =   "Therapist"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":10888
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":11162
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":11A3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":12316
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":12BF0
            Key             =   "Plastic Surgery"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":134CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":13DA4
            Key             =   "Medications"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":144AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":15E40
            Key             =   "Inquiry Followup"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":17EBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   10110
      Width           =   19050
      _ExtentX        =   33602
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   6191
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3889
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5203
            MinWidth        =   5203
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "7/18/2015"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:09 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   5055
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   18990
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   19050
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   2070
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   4
         Top             =   0
         Width           =   4095
      End
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5985
         Left            =   -1800
         Picture         =   "MDI.frx":18634
         ScaleHeight     =   5985
         ScaleWidth      =   7635
         TabIndex        =   3
         Top             =   -120
         Width           =   7635
      End
      Begin VB.Label lblResizer 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "resizer"
         Height          =   195
         Left            =   6810
         TabIndex        =   6
         Top             =   810
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19050
      _ExtentX        =   33602
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "PatientTracImages"
      HotImageList    =   "PatientTracImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tablet Notes"
            Object.ToolTipText     =   "Tablet Notes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Select Patient"
            Object.ToolTipText     =   "Select Patient"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "StimSheet"
            Object.ToolTipText     =   "Stim Sheet"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NewInquiry"
            Object.ToolTipText     =   "New Patient Inquiry"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Patient Data"
            Object.ToolTipText     =   "Patient Data"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Viewer"
            Object.ToolTipText     =   "Viewer"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Plastic Surgery"
            Object.ToolTipText     =   "Plastic Surgery"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Psychiatry"
            Object.ToolTipText     =   "Psychiatry"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SuperBill"
            Object.ToolTipText     =   "SuperBill"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Consent"
            Object.ToolTipText     =   "Consents"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Appointments"
            Object.ToolTipText     =   "Appointments"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Therapist"
            Object.ToolTipText     =   "Therapist"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Reports"
            Object.ToolTipText     =   "Reports"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Medications"
            Object.ToolTipText     =   "Prescription Application"
            ImageIndex      =   34
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DSM-IV"
            Object.ToolTipText     =   "DSM-IV LookUp"
            ImageIndex      =   36
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Surgery Log"
            Object.ToolTipText     =   "Surgery Log"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Inquiry FollowUp"
            ImageIndex      =   37
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Contact Notes"
            Object.ToolTipText     =   "Patient Contact Notes"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin PT_XP_Button.PT_XP cmdInterAlert 
         Height          =   435
         Left            =   13920
         TabIndex        =   11
         Top             =   60
         Width           =   2025
         _ExtentX        =   3572
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
         Picture         =   "MDI.frx":266E6
         AppearanceThemes=   3
         TransparentColor=   11250603
         Caption         =   "Practice Alerts"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "MDI.frx":26E60
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdElectronicNotice 
         Height          =   435
         Left            =   11490
         TabIndex        =   10
         Top             =   60
         Width           =   2205
         _ExtentX        =   3889
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
         Picture         =   "MDI.frx":27572
         AppearanceThemes=   3
         TransparentColor=   11250603
         Caption         =   "Electronic Order Pending"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "MDI.frx":27CEC
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdRefresh 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   17580
         TabIndex        =   9
         Top             =   60
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "MDI.frx":283FE
         AppearanceThemes=   3
         TransparentColor=   8421504
         Caption         =   "Refresh"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "MDI.frx":28998
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdBrowserScan 
         Height          =   435
         Left            =   16140
         TabIndex        =   8
         Top             =   60
         Width           =   1245
         _ExtentX        =   2196
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
         Picture         =   "MDI.frx":28CEA
         AppearanceThemes=   3
         TransparentColor=   8421504
         Caption         =   "Scanned Chart"
         DepthEvent      =   1
         ForeColorDisabled=   8388608
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "MDI.frx":29284
         ShowFocus       =   -1  'True
      End
      Begin VB.ComboBox cboBrowserReport 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12210
         TabIndex        =   7
         ToolTipText     =   "Reports, Browser Options"
         Top             =   1140
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.ComboBox cboFormList 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11670
         TabIndex        =   1
         Text            =   "<Assessments & Forms>"
         Top             =   1170
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   180
      Top             =   7290
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnu_Nursing 
         Caption         =   "InPatient Format"
         Index           =   0
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPatient 
      Caption         =   "&Patient"
      Begin VB.Menu mnu_NewInquiry 
         Caption         =   "New Patient Inquiry"
      End
      Begin VB.Menu mnuPatientSelect 
         Caption         =   "&Select Patient"
      End
      Begin VB.Menu mnuEncounter 
         Caption         =   "Select Patient and Encounter"
      End
   End
   Begin VB.Menu mnuScreens 
      Caption         =   "&Patient Data"
      Begin VB.Menu mnu_Patient_Instruction 
         Caption         =   "Patient Check-Out"
      End
      Begin VB.Menu mnu_Intake 
         Caption         =   "Patient Intake"
      End
      Begin VB.Menu mnu_Census 
         Caption         =   "&Patient Census"
      End
      Begin VB.Menu mnu_Immunization 
         Caption         =   "Immunization"
         Index           =   0
      End
      Begin VB.Menu mnu_LabResults 
         Caption         =   "Laboratory Result - Entry"
      End
      Begin VB.Menu mnu_LabResult 
         Caption         =   "Laboratory Results - Complex"
      End
      Begin VB.Menu mnu_Appt_History 
         Caption         =   "Patient Appointment History"
      End
      Begin VB.Menu mnu_Admission_Hist 
         Caption         =   "Patient Admission History"
      End
      Begin VB.Menu mnuScreens_PatientData 
         Caption         =   "Patient Scanned Data"
      End
      Begin VB.Menu mnu_Problem 
         Caption         =   "Patient Problem List"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Patient Forms"
      Begin VB.Menu mnu_Financial 
         Caption         =   "Encounter/Charges"
      End
      Begin VB.Menu mnu_HospNote 
         Caption         =   "InPatientNote"
      End
      Begin VB.Menu mnu_Patient_Response 
         Caption         =   "Inquiry FollowUp"
         Index           =   37
      End
      Begin VB.Menu mnu_Lab 
         Caption         =   "Lab - Radiology Orders"
      End
      Begin VB.Menu mnu_MechRating 
         Caption         =   "Mech Rating"
      End
      Begin VB.Menu mnu_PatientReport 
         Caption         =   "Patient Record"
      End
      Begin VB.Menu mnu_Correspondence 
         Caption         =   "Practice Letters"
      End
      Begin VB.Menu mnu_Prevous 
         Caption         =   "Previous EMR"
      End
      Begin VB.Menu mnu_Referral 
         Caption         =   "Referral/PCP Letter"
      End
      Begin VB.Menu mnu_SpecOrder 
         Caption         =   "Special Order"
      End
   End
   Begin VB.Menu mnu_Elect_Order 
      Caption         =   "&Electronic Orders"
      Begin VB.Menu mnu_EPresc_Script 
         Caption         =   "Write Prescription"
      End
      Begin VB.Menu mnu_EPresc_Auth 
         Caption         =   "Authorizations"
      End
      Begin VB.Menu mnu_EPresc_User 
         Caption         =   "User Administration"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_EPresc_DUR_Admin 
         Caption         =   "DUR Administration"
      End
      Begin VB.Menu mnu_EPresc_Pharmacy 
         Caption         =   "Manage Pharmacy"
      End
      Begin VB.Menu mnu_EPresc_PrefMed 
         Caption         =   "Preferred Meds"
      End
      Begin VB.Menu mnu_EPresc_Utilization 
         Caption         =   "Utilization Report"
      End
   End
   Begin VB.Menu mnu_Mgmt 
      Caption         =   "&Management"
      Begin VB.Menu mnu_ReportCreator 
         Caption         =   "Correspondence Creator"
      End
      Begin VB.Menu mnuReport_DaySheet 
         Caption         =   "Day Sheet"
      End
      Begin VB.Menu mnu_MAR 
         Caption         =   "MAR"
      End
      Begin VB.Menu mnu_Med_Report 
         Caption         =   "Medication Report"
      End
      Begin VB.Menu mnu_Tasks 
         Caption         =   "Messages/Tasks"
      End
      Begin VB.Menu mnu_Patient_Roster 
         Caption         =   "Practice Analysis"
      End
      Begin VB.Menu mnu_PatientCollect 
         Caption         =   "Patient Collection"
      End
      Begin VB.Menu mnu_PrintLog 
         Caption         =   "Practice Print Log"
      End
      Begin VB.Menu mnu_Rec_Request 
         Caption         =   "Record Request"
      End
      Begin VB.Menu mnu_Practice_Report 
         Caption         =   "Practice Reports"
      End
      Begin VB.Menu mnu_RefCensus 
         Caption         =   "&Referral Census"
      End
      Begin VB.Menu mnu_RoomLocations 
         Caption         =   "Room Locations"
      End
   End
   Begin VB.Menu mnu_ContactMgmt 
      Caption         =   "&Data Management"
      Begin VB.Menu mnu_ApptType 
         Caption         =   "Appointment Type"
      End
      Begin VB.Menu mnu_CPT_Compare 
         Caption         =   "CPT Comparison"
      End
      Begin VB.Menu mnu_EM 
         Caption         =   "E/M Maintanence"
      End
      Begin VB.Menu mnu_Family_Contact 
         Caption         =   "&Family Contacts"
      End
      Begin VB.Menu mnu_Reimburse 
         Caption         =   "Fee Reimbursement"
      End
      Begin VB.Menu mnu_Insurance 
         Caption         =   "Insurance Contacts"
      End
      Begin VB.Menu mnu_Legal_Contact 
         Caption         =   "&Legal Contacts"
      End
      Begin VB.Menu mnu_RefPhys 
         Caption         =   "Medical Contacts"
      End
      Begin VB.Menu mnu_MedData 
         Caption         =   "Medication Data"
      End
      Begin VB.Menu mnu_Pharmacy 
         Caption         =   "&Pharmacy"
      End
      Begin VB.Menu mnu_Vendors 
         Caption         =   "Vendors"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "&Configuration"
      Begin VB.Menu mnu_Appt_Type 
         Caption         =   "Appointment Schedule Type"
         Begin VB.Menu mnu_Office 
            Caption         =   "Office Appointments"
         End
         Begin VB.Menu mnu_InPatients 
            Caption         =   "InPatient Appointments"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSysConfig 
         Caption         =   "System Configuration"
      End
      Begin VB.Menu mnuDatabase 
         Caption         =   "Database &Connections"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuImages 
         Caption         =   "&Images Path and Settings"
      End
      Begin VB.Menu mnuDashboard 
         Caption         =   "&Dashboard Address"
      End
      Begin VB.Menu mnuExtDB 
         Caption         =   "External DB Configuration"
      End
   End
   Begin VB.Menu mnu_Patient_Audt 
      Caption         =   "Roster Evaluation Management"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      WindowList      =   -1  'True
      Begin VB.Menu mnuHelp_About 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHelp_SystemInfo 
         Caption         =   "System Info"
      End
      Begin VB.Menu mnuError 
         Caption         =   "Event Log"
      End
   End
   Begin VB.Menu mnu_CCD 
      Caption         =   "&ABN-CCD"
      Begin VB.Menu mnu_Ambulatory 
         Caption         =   "Ambulatory"
      End
      Begin VB.Menu mnu_CCDConnect 
         Caption         =   "CCD Connect"
      End
      Begin VB.Menu mnu_Enctrpt 
         Caption         =   "Encrypt"
      End
      Begin VB.Menu mnu_CCD_EPrescribe_CCD 
         Caption         =   "E-Prescribe"
      End
      Begin VB.Menu mnu_Lab_CCD 
         Caption         =   "Labs"
      End
      Begin VB.Menu mnu_Immunization_CCD 
         Caption         =   "Immunization"
      End
      Begin VB.Menu mnu_PatientDemographic 
         Caption         =   "Patient Demographics"
      End
      Begin VB.Menu mnuCCD_PHS_CCD 
         Caption         =   "PHS"
      End
      Begin VB.Menu mnu_Reports 
         Caption         =   "View Reports"
      End
   End
   Begin VB.Menu mnu_Facility 
      Caption         =   "Facilities"
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" ( _
                    ByVal hWnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2
Private fobjXtal As Object
Public fbooByPass As Boolean
Private fvarCmdLine
Private fstrMainDrive As String
Private fstrDefaultPractice As String
Private fbooFormLoaded As Boolean
Private fobjTerminal As Object
Public fbooPracticeReports As Boolean
Public fbooFarpointPracticeReports As Boolean
Public mlngFormID As Long
Public mbooFirst As Boolean
Private datTimer As Date

Private Function Validate_MD_Provider(Optional vlngProv As Long) As Boolean
Dim rsProvider As New ADODB.Recordset
Dim cnProvider As New ADODB.Connection
Dim strSQL As String
Dim lngProv As Long
Dim Idx%
Dim lngFac As Long

    On Error GoTo Error_Validate_MD_Provider
    Validate_MD_Provider = False
    If cnProvider.State = adStateClosed Then cnProvider.open psConnect(0)
    strSQL = "SELECT * FROM APPOINTMENT_PROVIDERS WHERE  PROVIDER_ID = " & mlngLogonUser
    Set rsProvider = cnProvider.Execute(strSQL)
    With rsProvider
        If Not rsProvider.EOF Then
            If Not IsNull((!Provider_Type)) Then
                If (!Provider_Type) = "PHYSICIAN" Then
                    Validate_MD_Provider = True
                End If
            End If
        End If
    End With
    If cnProvider.State = adStateOpen Then
        cnProvider.Close
        Set cnProvider = Nothing
        Set rsProvider = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Validate_MD_Provider:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Validate_MD_Provider", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Function Get_AppointmentFormat()
Dim cnFormat As New ADODB.Connection
Dim rsFormat As New ADODB.Recordset
Dim strSQL As String
Dim strFormat As String
Dim lngPrimKey As Long

    On Error GoTo Error_Get_AppointmentFormat
    If cnFormat.State = adStateClosed Then cnFormat.open psConnect(1)
    strSQL = "SELECT FORM_ID, ROW_NAME FROM CUSTOM_PRACTICE_FORM WHERE FORM_NAME = 'APPOINTMENT'"
    Set rsFormat = cnFormat.Execute(strSQL)
    With rsFormat
        If Not .EOF Then
            If Not IsNull(!ROW_NAME) Then strFormat = (!ROW_NAME)
            lngPrimKey = (!FORM_ID)
        End If
    End With
    Get_Edi_Static
    If strFormat = "InPatient" Then
        If clsEDI_Client.Site_ID = "1999" Then
            If Validate_MD_Provider Then
                If Not frmNursingHome.mbooCloseNurse Then
                    If Not IsFormLoaded("frmNursingHome") Then
                        Load frmNursingHome
                        frmNursingHome.Show
                        Exit Function
                    End If
                End If
            Else
                If Not IsFormLoaded("frmInPatientAppt") Then
                    Load frmInPatientAppt
                    frmInPatientAppt.Show
                End If
            End If
        Else
            If Not IsFormLoaded("frmPatientAppointment") Then
                'If IsFormLoaded("frmInPatient") Then Unload frmInPatient
                frmPatientAppointment.Show
            Else
                 'If IsFormLoaded("frmInPatient") Then Unload frmInPatient
                Load frmPatientAppointment
                frmPatientAppointment.Show
                frmPatientAppointment.Refresh
            End If
        End If
    ElseIf strFormat = "Internal Medicine" Then
        If Not frmInternalMed.mbooCloseNurse Then
            If Not IsFormLoaded("frmInternalMed") Then
                Load frmInternalMed
                frmInternalMed.Show
                Exit Function
            End If
        End If
    Else
        If IsFormLoaded("frmInPatientAppt") Then Unload frmInPatientAppt
        If clsEDI_Client.Site_ID = "2777" And clsProvider.title = "RN" Or clsProvider.title = "MA" Then
            Get_SuperProv
        End If
        If Not IsFormLoaded("frmPatientAppointment") Then
            frmPatientAppointment.Show
        Else
            frmPatientAppointment.Show
        End If
    End If
    If cnFormat.State = adStateOpen Then
        cnFormat.Close
        Set cnFormat = Nothing
        Set rsFormat = Nothing
    End If
    frmNursingHome.mbooCloseNurse = False
    Screen.MousePointer = vbDefault
    Exit Function
Error_Get_AppointmentFormat:
    If Err.Number = 401 Then Resume Next
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Get_AppointmentFormat", App.EXEName & "." & TypeName(Me)
    Exit Function
    Resume Next
End Function

Private Function Load_DaySheet()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "DaySheet"
    frmLoadControl.Show
    Exit Function
End Function

Public Sub ReleaseObject(frm As Form)
   Set frm = Nothing
End Sub
Private Function Load_Correspondence()

End Function

Private Sub cboBrowserReport_Change()
        
    If cboBrowserReport.Text = "Browser" Then
        frmBrowser.Show
''    ElseIf cboBrowserReport.Text = "Active Report" Then
''        frmMDIReport.Show
'    ElseIf cboBrowserReport.Text = "Fertility & Gynecology" Then
'        frmObGyn.Show
'    ElseIf cboBrowserReport.Text = "Psychiatry" Then
'        frmPsychiatry.Show
'    ElseIf cboBrowserReport.Text = "Appointments" Then
'        frmPatientAppointment.Show
    End If
End Sub

Private Sub cboBrowserReport_Click()
    If cboBrowserReport.Text = "Browser" Then
        frmBrowser.Show
'    ElseIf cboBrowserReport.Text = "Active Report" Then
'        frmMDIReport.Show
'    ElseIf cboBrowserReport.Text = "Fertility & Gynecology" Then
'        frmObGyn.Show
'    ElseIf cboBrowserReport.Text = "Psychiatry" Then
'        frmPsychiatry.Show
'    ElseIf cboBrowserReport.Text = "Appointments" Then
'        frmPatientAppointment.Show
    End If
End Sub

Private Sub cboFormList_Click()
'  Dim iFormID As Integer
'
'  If cboFormList.ListIndex = -1 Then cboFormList.Text = "<Assessments & Forms>": Exit Sub
'
'    piPatientID = frmPatients.piPatientID
'    piEncounterID = frmPatients.piEncounterID
'    piProviderID = frmPatients.piProviderID
'  If piPatientID = 0 Then
'    MsgBox "Select a Patient First.", vbExclamation
'    Exit Sub
'  End If
'
'  If piEncounterID = 0 Then
'    MsgBox "Select an Encounter First.", vbExclamation
'    Exit Sub
'  End If
'
'  'dnn....12/29/2003
''    If InStr(cboFormList.Text, "IVF") > 0 Then
''        Toolbar.Buttons(5).Visible = True
''    Else
''        'Toolbar.Buttons(5).Visible = False
''    End If
'
'
'
'    iFormID = cboFormList.ItemData(cboFormList.ListIndex)
'    'If (iFormID <> 20) And (iFormID <> 6) Then
'    '   If (piPatientID < 1) Then
'    '     MsgBox "Select a Patient First!", vbExclamation
'    '     Exit Sub
'    '   ElseIf (piEncounterID < 1) Then
'    '        MsgBox "Select an Encounter First!", vbExclamation
'    '        Exit Sub
'    '   End If
'    'End If
'    If InStr(cboFormList.List(UCase(cboFormList.ListIndex)), "STIM") > 0 Then
'        frmStimSheet.FormID = iFormID
'        frmStimSheet.Show
'    Else
'       'Set application specific properties for the form
'       frmBasic.FormID = iFormID
'       frmBasic.PatientID = piPatientID
'
'       '*** Note!   Hard coded value for testing purposes...
'       '            Replace when encounter and provider selection work
'       frmBasic.EncounterID = piEncounterID
'       frmBasic.ProviderID = piProviderID
'
'       frmBasic.Show
'   End If
'
End Sub

Private Sub cmdBrowserScan_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "ScanImages"
    Load frmLoadControl
    frmLoadControl.Show vbModal
End Sub

Private Sub cmdElectronicNotice_Click()
    Get_Providers mlngLogonUser
    If clsProvider.Emdeon_User <> vbNullString Then
        clsEmdeonConnect.Load_Type = "Authorizations"
        Connect clsEmdeonConnect.Load_Type
        Rx_InBox
    Else
        MsgBox "You are not enrolled for this feature. Please contact PatientTrac support at (786)350-3900 to complete enrollment!", vbInformation, "PatientTrac Electronic Orders"
        Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cmdInterAlert_Click()
        frmIntraOffice.Show vbModal
    If Check_Scanned_Docs Then
        cmdInterAlert.BackColor = &HD6D8FE
        cmdInterAlert.BorderColor = &H404080
        cmdInterAlert.ForeColor = &H0&
    End If
    Exit Sub
End Sub

Private Sub cmdRefresh_Click()
    frmInPatient.ZOrder 0
    Exit Sub
End Sub

Private Sub MDIForm_Activate()
Dim lstrConnect As String
Dim lstrFileName() As String
Dim lstrDirPath As String
Dim lstrDirPath2 As String
Dim lintNumOfFiles As Integer
Dim i As Integer
Dim lstrSystemDirectory As String
Dim lstrMainAppPath As String
Dim lstrMainDrive As String
Dim lbooRet As Boolean
Dim lstrServer As String
Dim strReportName As String
Dim intTries As Integer
Dim strDisplayStimSheetIcon As String
Dim strEnable As String
Dim fso As New FileSystemObject

    On Error GoTo Error_Main_Load

    If mbooFirst Then
        If mstrPractice = vbNullString Then GetDefaultPractice
        fbooFormLoaded = True
         Me.Caption = "PatientTrac Electronic Medical Records"
         strEnable = ReadRegistryValue("Software\KBNF", "LockData")
         If strEnable = "yes" Then 'Edit closed encounter data
             pbooLockClosedEncounterData = True
         ElseIf strEnable = "no" Then 'Lock closed encounter data
             pbooLockClosedEncounterData = False
         Else
             pbooLockClosedEncounterData = False
         End If
         psIniFile = App.Path & "\" & App.EXEName & ".ini"
         If Not FileExists(psIniFile) Then GenerateDefaultFile
         If mstrPractice = vbNullString Then GetDefaultPractice
         fvarCmdLine = mstrPractice
         lstrSystemDirectory = Get_WindowsSystemDirectory
         lstrMainDrive = Left(lstrSystemDirectory, InStr(lstrSystemDirectory, ":"))
         fstrMainDrive = mstrProgram & "\FirstNotebook\PatientTrac EMR"
         pstrMainPath = fstrMainDrive
         If fso.FileExists(App.Path & "\..\Images\PT_EMR.jpg") Then
             Me.Picture = LoadPicture(App.Path & "\..\Images\PT_EMR.jpg")
         End If
        Get_RegisterSettings
         If psServer(0) = vbNullString Then
             psServer(0) = QueryValue(HKEY_LOCAL_MACHINE, psRegDB, "Server0")
         End If
        If psConnect(0) <> vbNullString Then
              StatusBar.Panels(5).Text = "Server: " & psServer(0)
        End If
         If psConnect(1) <> vbNullString Then
             If fvarCmdLine = "" Or fvarCmdLine = "SetDB" Then
                 fvarCmdLine = GetDefaultPractice()
                 If InStr(fvarCmdLine, " Default") > 0 Then
                     fstrDefaultPractice = Left(fvarCmdLine, InStr(fvarCmdLine, " Default") - 1)
                 End If
             End If
         Else
             Unload Me
         End If
         Get_Edi_Static
         If mstrPractice = "IVF" Or mstrPractice = "Gyn" Then
             strDisplayStimSheetIcon = QueryValue(HKEY_LOCAL_MACHINE, psRegGeneral, "ShowStimSheetIcon")
             If strDisplayStimSheetIcon = "yes" Then
                 Toolbar.buttons(3).Visible = True
             Else
                 Toolbar.buttons(3).Visible = False
             End If
        Else
            Toolbar.buttons(3).Visible = False
         End If
         Get_Providers mlngLogonUser
         If clsProvider.Emdeon_User <> vbNullString Then
            Rx_InBox
         End If
        If mstrPractice <> "Psychiatry" Then
            Toolbar.buttons(4).Visible = True
        Else
            Toolbar.buttons(4).Visible = False
            Toolbar.buttons(16).Visible = False
        End If
        Toolbar.buttons(6).Visible = False
        StatusBar.Panels(5).Text = "Visit Date: " & Format(Now, "MMMM DD, YYYY")
        pbCancel = False
        If Check_Scanned_Docs Then
            cmdInterAlert.BackColor = &HD6D8FE
            cmdInterAlert.BorderColor = &H404080
            cmdInterAlert.ForeColor = &H0&
        End If
    End If
    If mbooFirst Then
        MDIForm_Resize
        Get_AppointmentFormat
    End If
    
    mbooFirst = False
    Exit Sub
Error_Main_Load:
    Exit Sub
    Resume Next
End Sub

Private Sub MDIForm_Load()
    mbooFirst = True
    datTimer = DateAdd("n", 15, Now)   'set it to fire after 5 minutes
    Timer1.Interval = 500
    Timer1.Enabled = True
    Exit Sub
Error_Main_Load:
    Exit Sub
    Resume
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim fso As New FileSystemObject
Dim fsoFolder As Folder
Dim fsoFile As File
Dim Idx%

    On Error Resume Next
    Set fsoFolder = fso.GetFolder("c:\")
    If IsFormLoaded("frmInPatientAppt") Then Unload frmInPatientAppt
    If IsFormLoaded("frmPatientAppointment") Then Unload frmPatientAppointment
    If IsFormLoaded("frmBrowser") Then Unload frmBrowser
    If IsFormLoaded("frmInPatient") Then Unload frmInPatient
    If IsFormLoaded("frmEnctr") Then Unload frmEnctr
    If IsFormLoaded("frmSuperBill") Then Unload frmSuperBill
    If IsFormLoaded("frmLoadControl") Then Unload frmLoadControl
    If IsFormLoaded("frmDSMIV") Then Unload frmDSMIV
    If IsFormLoaded("frmMeds") Then Unload frmMeds
    If IsFormLoaded("frmPatients") Then Unload frmPatients
    If IsFormLoaded("frmTelephone") Then Unload frmTelephone
    If IsFormLoaded("PatientInfo") Then Unload PatientInfo
    Exit Sub
    
End Sub



Private Sub MDIForm_Resize()

    If mbooFirst Then
        SetFrmArea Me
        StretchFrmBKG Me, picOriginal, picStretched
    End If
   Exit Sub
End Sub

Private Function Update_Logon()
Dim cnLogOn As New ADODB.Connection
Dim rsLogOn As New ADODB.Recordset
Dim strSQL As String
Dim strPC As String
Dim lngLogIn As Long
Dim strTime As String
Dim strLogon As String

    If cnLogOn.State = adStateClosed Then cnLogOn.open psConnect(0)
    On Error GoTo Error_Update_Logon
    strTime = Format(Now, "hh:mm AM/PM")
    strLogon = Format(Now, "mm/dd/yyyy")
    strPC = Get_ComputerName
    strSQL = "Get_User_Logon " & mlngLogonUser & ", '" & strPC & "', '" & App.EXEName & "' , '" & Format(strLogon, "mm/dd/yyyy") & "'"
    Set rsLogOn = cnLogOn.Execute(strSQL)
    With rsLogOn
       If Not rsLogOn.EOF Then
            lngLogIn = rsLogOn.Fields("LOGON_ID")
       End If
    End With
    If lngLogIn > 0 Then
        strSQL = "Update_User_Logon '" & Format(strTime, "hh:mm AM/PM") & "', '" & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM") & "', " & lngLogIn
            Set rsLogOn = cnLogOn.Execute(strSQL)
    End If
    If cnLogOn.State = adStateOpen Then
        cnLogOn.Close
        Set cnLogOn = Nothing
        Set rsLogOn = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_Update_Logon:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Update_Logon", App.EXEName & "." & Me.Name
    Exit Function
    Resume
End Function

Private Sub mnu_Admission_Hist_Click()
'   If piPatientID = 0 Then
'        MsgBox "Please select a patient first!", vbInformation, "PatientTrac Electronic Medical Records"
'        Exit Sub
'    End If
    Load frmAdmitHistory
    frmAdmitHistory.mPatID = piPatientID
    frmAdmitHistory.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_Ambulatory_Click()
Dim cnTest As New ADODB.Connection
Dim rsTest As ADODB.Recordset
Dim strSQL As String
Dim iRet As Long
Dim strPath As String
Dim strExe As String
Dim strDB As String
Dim strData As String
Dim fso As New FileSystemObject
Dim strFolder As String



    On Error GoTo Error_Ambulatory
    If mstrProgram = "C:\Program Files (x86)" Then
        strPath = "SOFTWARE\Wow6432Node\FirstNoteBook\PatientTrac\Databases"
    Else
        strPath = "SOFTWARE\FirstNoteBook\PatientTrac\Databases"
    End If
    
    
    strDB = QueryValue(HKEY_LOCAL_MACHINE, strPath, "Server0")
    strPath = "C:\Program Files (x86)\FirstNotebook\PatientTrac EMR\Components\DHITCOMHelper"
    If Not fso.FolderExists("C:\ABN_Data") Then fso.CreateFolder ("C:\ABN_Data")
    If Not fso.FolderExists("C:\ABN_Data\" & piPatientID) Then fso.CreateFolder ("C:\ABN_Data\" & piPatientID)
    strFolder = "C:\ABN_Data\" & piPatientID
    
    
    
'    strDB = "Data Source=" & strDB & ";Initial Catalog=Clinical_Repository;Integrated Security=True"
'    strExe = strPath & "\" & "DHITCOMHelper_Console.exe" & " " & "-c" & " " & strDB _
'        & " " & "-s" & " " & "{call CCDC (?)}" & " " & "-p" _
'        & " " & "@vClientid" & " " & "-v" & " " & piPatientID & " " & "-f" & " " & strFolder & "\Ambulatory" & "_" & Format(Now, "mm_dd_yyyy") & ".xml"
    'iRet = ShellExecute(Me.Hwnd, "cmd.exe", strFolder, strExe, strFolder & "\Ambulatory", 1)
    iRet = ShellExecute(Me.hWnd, "open", strPath, strExe, strFolder & "\Ambulatory", 1)
    Screen.MousePointer = vbDefault
   Exit Sub
Error_Ambulatory:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Ambulatory_CCD", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub mnu_Appt_History_Click()
    frmPatAppHist.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_ApptType_Click()
    frmApptMaint.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_CCD_EPrescribe_Click()

End Sub

Private Sub mnu_CCD_EPrescribe_CCD_Click()
Dim cnTest As New ADODB.Connection
Dim rsTest As ADODB.Recordset
Dim strSQL As String
Dim iRet As Long
Dim strPath As String
Dim strExe As String
Dim strDB As String
Dim strData As String
Dim fso As New FileSystemObject
Dim strFolder As String

    On Error GoTo Error_EPrescribe_CCD
    If mstrProgram = "C:\Program Files (x86)" Then
        strPath = "SOFTWARE\Wow6432Node\FirstNoteBook\PatientTrac\Databases"
    Else
        strPath = "SOFTWARE\FirstNoteBook\PatientTrac\Databases"
    End If

    strDB = QueryValue(HKEY_LOCAL_MACHINE, strPath, "Server0")
    strPath = "C:\Program Files (x86)\FirstNotebook\PatientTrac EMR\Components\DHITCOMHelper"
    If Not fso.FolderExists("C:\ABN_Data") Then fso.CreateFolder ("C:\ABN_Data")
    If Not fso.FolderExists("C:\ABN_Data\" & piPatientID) Then fso.CreateFolder ("C:\ABN_Data\" & piPatientID)
    strFolder = "C:\ABN_Data\" & piPatientID
    strDB = "Data Source=" & strDB & ";Initial Catalog=Clinical_Repository;Integrated Security=True"
    strExe = "DHITCOMHelper_Console.exe" & " " & "-c" & " " & strDB _
        & " " & "-s" & " " & "{call CCDC_ePrescribed (?)}" & " " & "-p" _
        & " " & "@vPatient_ID" & " " & "-v" & " " & piPatientID & " " & "-f" & " " & strFolder & "\E-Prescribe" & "_" & Format(Now, "mm_dd_yyyy") & ".xml"
    iRet = ShellExecute(Me.hWnd, "open", strPath, strExe, "C:|", 1)
    Screen.MousePointer = vbDefault
   Exit Sub
Error_EPrescribe_CCD:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "EPrescribe_CCD", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub mnu_Census_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "PatientRoster"
    frmLoadControl.Caption = "Practice Analysis"
    Load frmLoadControl
    frmLoadControl.Show vbModal
End Sub

Private Sub mnu_Correspondence_Click()
Dim lngDateDiff As Long
Dim strDate As String
Dim strTermDate As String
Dim strFinance As String
Dim strOption As String

    strFinance = "SOFTWARE\Denge"
    'strOption = QueryValue(HKEY_LOCAL_MACHINE, strFinance, "Correspondence")
'    If strOption = "True" Then
'        strOption = QueryValue(HKEY_LOCAL_MACHINE, strFinance, "ApprovedDate")
'        strDate = Format(Now, "mm/dd/yyyy")
'        strTermDate = Format(strOption, "mm/dd/yyyy")
'        lngDateDiff = DateDiff("d", strDate, strTermDate)
'        If lngDateDiff < 10 Then
'            MsgBox "License for this function has expired, Please contact PatientTrac!", vbInformation
'            Exit Sub
'        Else
    Load_Correspondence
       ' End If
    'Else
      '  MsgBox "PatientTrac PMS - You do not have a license for this product, Please contact PatientTrac Sales!", vbInformation
     '   mnu_Correspondence.Enabled = False
      '  Exit Sub
   ' End If
    Exit Sub
End Sub

Private Sub mnu_CPT_Compare_Click()
    mstrPatOcx = "FN_PsyMse"
    mstrPatCtl = "CPT_Compare"
    frmLoadControl.Show
    Exit Sub
End Sub



Private Sub mnu_EM_Click()
'    mstrPatOcx = "FN_RevSys"
'    mstrPatCtl = "CPT_Maint"
'    frmLoadControl.Show
    frmCPTMaint.Show
    Exit Sub
End Sub

Private Sub mnu_Enctrpt_Click()

    'C:\Program Files (x86)\FirstNotebook\PatientTrac EMR\Components\CCDConnect
    
    
    Exit Sub
End Sub

Private Sub mnu_EPresc_Auth_Click()

    If mstrURL = "https://clinician.emdeon.com" Then
        clsEmdeonConnect.Load_Type = "Authorizations"
        Connect clsEmdeonConnect.Load_Type
    End If
    Exit Sub
End Sub

Private Sub mnu_EPresc_Caregiver_Click()

    clsEmdeonConnect.Load_Type = "Caregiver"
    Connect clsEmdeonConnect.Load_Type
    Exit Sub
End Sub

Private Sub mnu_EPresc_DUR_Admin_Click()
     clsEmdeonConnect.Load_Type = "DUR"
     Connect clsEmdeonConnect.Load_Type
    Exit Sub
End Sub


Private Sub mnu_EPresc_Payer_Click()
     clsEmdeonConnect.Load_Type = "Payer"
     Connect clsEmdeonConnect.Load_Type
    Exit Sub
End Sub


Private Sub mnu_EPresc_Pharmacy_Click()

     clsEmdeonConnect.Load_Type = "Pharmacy"
     Connect clsEmdeonConnect.Load_Type
End Sub


Private Sub mnu_EPresc_PrefMed_Click()

     clsEmdeonConnect.Load_Type = "Prescription"
     Connect clsEmdeonConnect.Load_Type
    Exit Sub
End Sub


Private Sub mnu_EPresc_Script_Click()

    If mstrURL = "https://clinician.emdeon.com" Then
        clsEmdeonConnect.Load_Type = "Patient"
        Connect clsEmdeonConnect.Load_Type
    End If
    Exit Sub
End Sub

Private Sub mnu_EPresc_Utilization_Click()

    If mstrURL = "https://clinician.emdeon.com" Then
        clsEmdeonConnect.Load_Type = "Utilization Report"
        Connect clsEmdeonConnect.Load_Type
     End If
    Exit Sub
End Sub

Private Sub mnu_Facility_Click()
    If IsFormLoaded("frmPatientAppointment") Then Unload frmPatientAppointment
    frmManage.Show vbModal
    frmPatientAppointment.Show
    StatusBar.Panels(1).Text = clsFacility.Name
End Sub

Private Sub mnu_Family_Contact_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "Family_Contact"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_Financial_Click()
    mstrPatOcx = "PatientTrac_Report"
    mstrPatCtl = "SpecificPatient"
    frmLoadControl.piSelPatientID = piPatientID
    frmLoadControl.Caption = "PatientTrac Patient Encounter"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_HospNote_Click()
    mstrPatOcx = "PatientTrac_Report"
    mstrPatCtl = "HospNote"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub mnu_Immunization_CCD_Click()
Dim cnTest As New ADODB.Connection
Dim rsTest As ADODB.Recordset
Dim strSQL As String
Dim iRet As Long
Dim strPath As String
Dim strExe As String
Dim strDB As String
Dim strData As String
Dim fso As New FileSystemObject
Dim strFolder As String

    On Error GoTo Error_Immun_CCD
    If mstrProgram = "C:\Program Files (x86)" Then
        strPath = "SOFTWARE\Wow6432Node\FirstNoteBook\PatientTrac\Databases"
    Else
        strPath = "SOFTWARE\FirstNoteBook\PatientTrac\Databases"
    End If

    strDB = QueryValue(HKEY_LOCAL_MACHINE, strPath, "Server0")
    strPath = "C:\Program Files (x86)\FirstNotebook\PatientTrac EMR\Components\DHITCOMHelper"
    If Not fso.FolderExists("C:\ABN_Data") Then fso.CreateFolder ("C:\ABN_Data")
    If Not fso.FolderExists("C:\ABN_Data\" & piPatientID) Then fso.CreateFolder ("C:\ABN_Data\" & piPatientID)
    strFolder = "C:\ABN_Data\" & piPatientID
    strDB = "Data Source=" & strDB & ";Initial Catalog=Clinical_Repository;Integrated Security=True"
    strExe = "DHITCOMHelper_Console.exe" & " " & "-c" & " " & strDB _
        & " " & "-s" & " " & "{call Call_IMMUNIZATION (?)}" & " " & "-p" _
        & " " & "@vClientid" & " " & "-v" & " " & piPatientID & " " & "-f" & " " & strFolder & "\Immunization" & "_" & Format(Now, "mm_dd_yyyy") & ".xml"
    iRet = ShellExecute(Me.hWnd, "open", strPath, strExe, "C:|", 1)
    Screen.MousePointer = vbDefault
   Exit Sub
Error_Immun_CCD:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Immun_CCD", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub mnu_Immunization_Click(Index As Integer)
    mstrPatOcx = "PatientTrac_Report"
    mstrPatCtl = "VaccineRecord"
    frmLoadControl.Caption = "Patient Immunization"
'    Load frmLoadControl
    frmLoadControl.Show vbModal
End Sub

'Private Sub mnu_InPatient_Click()
'Dim cnFormat As New ADODB.Connection
'Dim rsFormat As New ADODB.Recordset
'Dim strSQL As String
'Dim strFormat As String
'Dim lngPrimKey As Long
'
'    On Error GoTo Error_Format_InPatient
'    If cnFormat.State = adStateClosed Then cnFormat.open psConnect(1)
'    strSQL = "SELECT FORM_ID, ROW_NAME FROM CUSTOM_PRACTICE_FORM WHERE FORM_NAME = 'APPOINTMENT'"
'    Set rsFormat = cnFormat.Execute(strSQL)
'        With rsFormat
'            If Not .EOF Then
'                If Not IsNull(!ROW_NAME) Then strFormat = (!ROW_NAME)
'                lngPrimKey = rsFormat.Fields("FORM_ID")
'            End If
'        End With
'    If strFormat <> vbNullString Then
'        strSQL = "UPDATE CUSTOM_PRACTICE_FORM SET ROW_NAME = 'InPatient' WHERE FORM_ID = " & lngPrimKey
'        Set rsFormat = cnFormat.Execute(strSQL)
'    Else
'        strSQL = "INSERT INTO CUSTOM_PRACTICE_FORM (ROW_NAME, FORM_NAME) values ('InPatient', 'APPOINTMENT')"
'        Set rsFormat = cnFormat.Execute(strSQL)
'    End If
'    If cnFormat.State = adStateOpen Then
'        cnFormat.Close
'        Set cnFormat = Nothing
'        Set rsFormat = Nothing
'    End If
'    If IsFormLoaded("frmPatientAppointment") Then Unload frmPatientAppointment
'    Load frmInPatientAppt
'    frmInPatientAppt.Show
'    Screen.MousePointer = vbDefault
'    Exit Sub
'Error_Format_InPatient:
'    Screen.MousePointer = vbDefault
'    Write_Error Err.Description, Err.Number, strSQL, "Format_InPatient", App.EXEName & "." & TypeName(Me)
'    Exit Sub
'    Resume
'End Sub

Private Sub mnu_InPatients_Click(Index As Integer)
Dim cnFormat As New ADODB.Connection
Dim rsFormat As New ADODB.Recordset
Dim strSQL As String
Dim strFormat As String
Dim lngPrimKey As Long

    On Error GoTo Error_Format_InPatient
    If cnFormat.State = adStateClosed Then cnFormat.open psConnect(1)
    strSQL = "SELECT FORM_ID, ROW_NAME FROM CUSTOM_PRACTICE_FORM WHERE FORM_NAME = 'APPOINTMENT'"
    Set rsFormat = cnFormat.Execute(strSQL)
        With rsFormat
            If Not .EOF Then
                If Not IsNull(!ROW_NAME) Then strFormat = (!ROW_NAME)
                lngPrimKey = rsFormat.Fields("FORM_ID")
            End If
        End With
    If strFormat <> vbNullString Then
        strSQL = "UPDATE CUSTOM_PRACTICE_FORM SET ROW_NAME = 'InPatient' WHERE FORM_ID = " & lngPrimKey
        Set rsFormat = cnFormat.Execute(strSQL)
    Else
        strSQL = "INSERT INTO CUSTOM_PRACTICE_FORM (ROW_NAME, FORM_NAME) values ('InPatient', 'APPOINTMENT')"
        Set rsFormat = cnFormat.Execute(strSQL)
    End If
    If cnFormat.State = adStateOpen Then
        cnFormat.Close
        Set cnFormat = Nothing
        Set rsFormat = Nothing
    End If
    If IsFormLoaded("frmPatientAppointment") Then Unload frmPatientAppointment
    Load frmInPatientAppt
    frmInPatientAppt.Show
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_InPatient:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_InPatient", App.EXEName & "." & TypeName(Me)
    Exit Sub
    Resume
End Sub


Private Sub mnu_Insurance_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "Insurers"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub mnu_Intake_Click()
    frmPatientIntake.Show
    Exit Sub
     
End Sub

Private Sub mnu_Lab_CCD_Click()
Dim cnTest As New ADODB.Connection
Dim rsTest As ADODB.Recordset
Dim strSQL As String
Dim iRet As Long
Dim strPath As String
Dim strExe As String
Dim strDB As String
Dim strData As String
Dim fso As New FileSystemObject
Dim strFolder As String

    On Error GoTo Error_Labs_CCD
    If mstrProgram = "C:\Program Files (x86)" Then
        strPath = "SOFTWARE\Wow6432Node\FirstNoteBook\PatientTrac\Databases"
    Else
        strPath = "SOFTWARE\FirstNoteBook\PatientTrac\Databases"
    End If

    strDB = QueryValue(HKEY_LOCAL_MACHINE, strPath, "Server0")
    strPath = "C:\Program Files (x86)\FirstNotebook\PatientTrac EMR\Components\DHITCOMHelper"
    If Not fso.FolderExists("C:\ABN_Data") Then fso.CreateFolder ("C:\ABN_Data")
    If Not fso.FolderExists("C:\ABN_Data\" & piPatientID) Then fso.CreateFolder ("C:\ABN_Data\" & piPatientID)
    strFolder = "C:\ABN_Data\" & piPatientID
    strDB = "Data Source=" & strDB & ";Initial Catalog=Clinical_Repository;Integrated Security=True"
    strExe = "DHITCOMHelper_Console.exe" & " " & "-c" & " " & strDB _
        & " " & "-s" & " " & "{call CCDC_LabResult (?)}" & " " & "-p" _
        & " " & "@vClientid" & " " & "-v" & " " & piPatientID & " " & "-f" & " " & strFolder & "\Labs" & "_" & Format(Now, "mm_dd_yyyy") & ".xml"
    iRet = ShellExecute(Me.hWnd, "open", strPath, strExe, "C:|", 1)
    Screen.MousePointer = vbDefault
   Exit Sub
Error_Labs_CCD:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Labs_CCD", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub mnu_lab_Click()
    mstrPatOcx = "FN_PsyMse"
    mstrPatCtl = "Patient_Orders"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub mnu_LabResult_Click()
    mstrPatOcx = "FN_Orders"
    mstrPatCtl = "Lab_Results"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub mnu_LabResults_Click()
    mstrPatOcx = "FN_PsyMse"
    mstrPatCtl = "LabResult_Simple"
    frmLoadControl.Caption = "Laboratory Result Entry"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub


Private Sub mnu_Legal_Contact_Click()
    mstrPatOcx = "FN_PsyMse"
    mstrPatCtl = "Legal_Contact"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_MAR_Click()
    mstrPatOcx = "PT_Medications"
    mstrPatCtl = "MedicationAdmin"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_MechRating_Click()
    mstrPatOcx = "PatientTrac_Report"
    mstrPatCtl = "Mech_Scales"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub mnu_Med_Report_Click()
    mstrPatOcx = "PT_Medications"
    mstrPatCtl = "MedReport"
    frmLoadControl.piSelPatientID = piPatientID
    frmLoadControl.Caption = "Medication Report"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_MedData_Click()
    mstrPatOcx = "PT_Medications"
    mstrPatCtl = "PsychMeds"
    frmLoadControl.Caption = "Psychotropic Medication Data"
    Load frmLoadControl
    frmLoadControl.Show vbModal
End Sub

Private Sub mnu_NewInquiry_Click()
    mstrPatContact = "Contact"
    mstrPatOcx = "PatientTrac_Report"
    mstrPatCtl = "PatientContact"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub mnu_Nursing_Click(Index As Integer)
    
    If Not IsFormLoaded("frmNursingHome") Then
        Unload frmPatientAppointment
        Load frmNursingHome
        frmNursingHome.Show
        Exit Sub
    End If
End Sub

Private Sub mnu_Office_Click()
Dim cnFormat As New ADODB.Connection
Dim rsFormat As New ADODB.Recordset
Dim strSQL As String
Dim strFormat As String
Dim lngPrimKey As Long

    On Error GoTo Error_Format_Office
    If cnFormat.State = adStateClosed Then cnFormat.open psConnect(1)
    strSQL = "SELECT FORM_ID, ROW_NAME FROM CUSTOM_PRACTICE_FORM WHERE FORM_NAME = 'APPOINTMENT'"
    Set rsFormat = cnFormat.Execute(strSQL)
        With rsFormat
            If Not .EOF Then
                If Not IsNull(!ROW_NAME) Then strFormat = (!ROW_NAME)
                lngPrimKey = rsFormat.Fields("FORM_ID")
            End If
        End With
    If strFormat <> vbNullString Then
        strSQL = "UPDATE CUSTOM_PRACTICE_FORM SET ROW_NAME = 'Office' WHERE FORM_ID = " & lngPrimKey
        Set rsFormat = cnFormat.Execute(strSQL)
    Else
        strSQL = "INSERT INTO CUSTOM_PRACTICE_FORM (ROW_NAME, FORM_NAME) values ('Office', 'APPOINTMENT')"
        Set rsFormat = cnFormat.Execute(strSQL)
    End If
    If cnFormat.State = adStateOpen Then
        cnFormat.Close
        Set cnFormat = Nothing
        Set rsFormat = Nothing
    End If
    If IsFormLoaded("frmInPatientAppt") Then Unload frmInPatientAppt
    frmPatientAppointment.Prepare_Grid
    frmPatientAppointment.Show
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Format_Office:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "Format_Office", App.EXEName & "." & TypeName(Me)
    Exit Sub
    Resume
End Sub

Private Sub mnu_Patient_Audt_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "PatientAudit"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_Patient_Instruction_Click()
    If clsEDI_Client.Site_ID = vbNullString Then Get_Edi_Static
    If clsEDI_Client.Site_ID = "1999" Then
        mstrPatOcx = "Reporter"
        mstrPatCtl = "InPatientNotes"
        frmLoadControl.piSelPatientID = piPatientID
        frmLoadControl.piSelEnctrID = piEncounterID
        frmLoadControl.Caption = "Print Patient Progress Note"
        Load frmLoadControl
        frmLoadControl.Show vbModal
    Else
        mstrPatOcx = "PatientTrac_Report"
        mstrPatCtl = "Patient_Instructions"
        frmLoadControl.piSelPatientID = piPatientID
        frmLoadControl.piSelEnctrID = piEncounterID
        frmLoadControl.Caption = "Patient Orders - Instructions"
        Load frmLoadControl
        frmLoadControl.Show vbModal
    End If
    Exit Sub
End Sub

Private Sub mnu_Patient_Roster_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "PatientRoster"
    frmLoadControl.Caption = "Practice Analysis"
    Load frmLoadControl
    frmLoadControl.Show vbModal
End Sub

Private Sub mnu_PatientCollect_Click()

    On Error GoTo Error_Practice_Collect
    Screen.MousePointer = vbHourglass
    If mlngLogonUser > 0 Then
        Get_SecurityRole mlngLogonUser
        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
            If IsProperCase(clsPracticeRoles.RoleTitle) <> "Administrator" Then
                MsgBox "You are not authorized to access Financial Reports!", vbInformation, "PatientTrac Security Management"
                Write_Error "PatientTrac Financial Report Security Attempt", CStr(mlngLogonUser), "PatientTrac Financial Report Security Attempt", "Practice_Report", Me.Name
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            mstrPatOcx = "FN_RevSys"
            mstrPatCtl = "FinancialRpt"
            Load frmLoadControl
            frmLoadControl.Show vbModal
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Practice_Collect:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Practice_Collect", Me.Name
    Exit Sub
    Resume
End Sub

Private Sub mnu_PatientReport_Click()
    frmPreview.piSelPatientID = piPatientID
    frmPreview.mlngGrdEncr = piEncounterID
    frmPreview.mstrLoad_Type = "Patient Record"
    Load frmPreview
    frmPreview.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_Pharmacy_Click()
    mstrPatOcx = "FN_PP_PS"
    mstrPatCtl = "Pharm_Contact"
    frmLoadControl.Caption = "PatientTrac Pharmacy Management"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_Practice_Report_Click()
    
    On Error GoTo Error_Practice_Report
    
    Screen.MousePointer = vbHourglass
    If mlngLogonUser > 0 Then
        Get_SecurityRole mlngLogonUser
        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
            If IsProperCase(clsPracticeRoles.RoleTitle) <> "Administrator" Then
                MsgBox "You are not authorized to access Financial Reports!", vbInformation, "PatientTrac Security Management"
                Write_Error "PatientTrac Financial Report Security Attempt", CStr(mlngLogonUser), "PatientTrac Financial Report Security Attempt", "Practice_Report", Me.Name
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            mstrPatOcx = "FN_RevSys"
            mstrPatCtl = "FinancialRpt"
            Load frmLoadControl
            frmLoadControl.Show vbModal
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Practice_Report:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "Practice_Report", Me.Name
    Exit Sub
    Resume
End Sub

Private Sub mnu_Prevous_Click()
    mstrPatOcx = "Reporter"
    mstrPatCtl = "SOAP"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub mnu_PrintLog_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "PracticePrint"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub mnu_Providers_Click()
    frmProviders.Show
    Exit Sub
End Sub



Private Sub mnu_Rec_Request_Click()
    mstrPatOcx = "Reporter"
    mstrPatCtl = "PreviewReport"
    Load frmLoadControl
    frmLoadControl.Show vbModal
End Sub

Private Sub mnu_RefCensus_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "PatientRoster"
    frmLoadControl.Caption = "Practice Analysis"
    Load frmLoadControl
    frmLoadControl.Show vbModal
End Sub

Private Sub mnu_Referral_Click()
    frmPreview.piSelPatientID = piPatientID
    frmPreview.mlngGrdEncr = piEncounterID
    frmPreview.mstrLoad_Type = "Referral Letter"
    Load frmPreview
    frmPreview.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_RefPhys_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "ReferPhys"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_Reimburse_Click()
    mstrPatOcx = "FN_PsyMse"
    mstrPatCtl = "FeeSchedule"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub mnu_ReportCreator_Click()
    On Error GoTo Error_ReportCreator
    Screen.MousePointer = vbHourglass
    If mlngLogonUser > 0 Then
        Get_SecurityRole mlngLogonUser
        If UCase(clsPracticeRoles.Permission) <> "ALL" Then
            If IsProperCase(clsPracticeRoles.RoleTitle) <> "Administrator" Then
                MsgBox "You are not authorized to access Financial Reports!", vbInformation, "PatientTrac Security Management"
                Write_Error "PatientTrac Financial Report Security Attempt", CStr(mlngLogonUser), "PatientTrac Financial Report Security Attempt", "ReportCreator", Me.Name
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            mstrPatOcx = "PatientTrac_Report"
            mstrPatCtl = "Letter"
            frmLoadControl.Show
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_ReportCreator:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, "", "ReportCreator", Me.Name
    Exit Sub
    Resume
    Exit Sub
End Sub

Private Sub mnu_Reports_Click()
    frmBrowser.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_SpecOrder_Click()
    frmSpecOrder.Show
    Exit Sub
End Sub

Private Sub mnu_Tasks_Click()
    frmTask.Show vbModal
    Exit Sub
End Sub

Private Sub mnu_Vendors_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "Vendors"
    frmLoadControl.Show
    Exit Sub
End Sub

Private Sub mnuCCD_PHS_CCD_Click()
Dim cnTest As New ADODB.Connection
Dim rsTest As ADODB.Recordset
Dim strSQL As String
Dim iRet As Long
Dim strPath As String
Dim strExe As String
Dim strDB As String
Dim strData As String
Dim fso As New FileSystemObject
Dim strFolder As String

    On Error GoTo Error_EPrescribe_CCD
    If mstrProgram = "C:\Program Files (x86)" Then
        strPath = "SOFTWARE\Wow6432Node\FirstNoteBook\PatientTrac\Databases"
    Else
        strPath = "SOFTWARE\FirstNoteBook\PatientTrac\Databases"
    End If

    strDB = QueryValue(HKEY_LOCAL_MACHINE, strPath, "Server0")
    strPath = "C:\Program Files (x86)\FirstNotebook\PatientTrac EMR\Components\DHITCOMHelper"
    If Not fso.FolderExists("C:\ABN_Data") Then fso.CreateFolder ("C:\ABN_Data")
    If Not fso.FolderExists("C:\ABN_Data\" & piPatientID) Then fso.CreateFolder ("C:\ABN_Data\" & piPatientID)
    strFolder = "C:\ABN_Data\" & piPatientID
    strDB = "Data Source=" & strDB & ";Initial Catalog=Clinical_Repository;Integrated Security=True"
    strExe = "DHITCOMHelper_Console.exe" & " " & "-c" & " " & strDB _
        & " " & "-s" & " " & "{call CCDC_ePrescribed (?)}" & " " & "-p" _
        & " " & "@vPatient_ID" & " " & "-v" & " " & piPatientID & " " & "-f" & " " & strFolder & "\Public_Health" & "_" & Format(Now, "mm_dd_yyyy") & ".xml"
    iRet = ShellExecute(Me.hWnd, "open", strPath, strExe, "C:|", 1)
    Screen.MousePointer = vbDefault
   Exit Sub
Error_EPrescribe_CCD:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "EPrescribe_CCD", App.EXEName & "." & Me.Name
    Exit Sub
    Resume Next
End Sub

Private Sub mnuDashboard_Click()
    frmDialogSetURL.Show vbModal
End Sub

Private Sub mnuDatabase_Click()
 '   frmSetServers.Show vbModal
End Sub

Private Sub mnuEncounter_Click()
    frmPatients.Show vbModal

End Sub

Private Sub mnuError_Click()
    frmSystemError.Show vbModal
End Sub

Private Sub mnuExtDB_Click()
    frmDialogSetExt.Show vbModal
End Sub

Private Sub mnuFile_Exit_Click()
Dim target_hwnd  As Long
    On Error GoTo Error_Close
    If IsFormLoaded("frmPatientAppointment") Then Unload frmPatientAppointment
    If IsFormLoaded("frmBrowser") Then Unload frmBrowser
    If IsFormLoaded("frmInPatientAppt") Then Unload frmInPatientAppt
    If IsFormLoaded("frmPatientAppointment") Then Unload frmPatientAppointment
    If IsFormLoaded("frmBrowser") Then Unload frmBrowser
    If IsFormLoaded("frmInPatient") Then Unload frmInPatient
    If IsFormLoaded("frmEnctr") Then Unload frmEnctr
    If IsFormLoaded("frmSuperBill") Then Unload frmSuperBill
    If IsFormLoaded("frmLoadControl") Then Unload frmLoadControl
    If IsFormLoaded("frmDSMIV") Then Unload frmDSMIV
    If IsFormLoaded("frmMeds") Then Unload frmMeds
    If IsFormLoaded("frmPatients") Then Unload frmPatients
    If IsFormLoaded("frmTelephone") Then Unload frmTelephone
    If IsFormLoaded("PatientInfo") Then Unload PatientInfo
    Unload Me
    Go_Close
    Exit Sub
Error_Close:
    MsgBox Err.Description, Err.Number, vbInformation
    Exit Sub
    Resume
End Sub


Private Sub InitializeToolbar()
    Toolbar.buttons(1).Caption = "Patient" & vbCrLf & "Profile"
End Sub

Private Sub mnuHelp_About_Click()
Dim lstrText As String
Dim lstrVersion As String
Dim lstrLegal As String
Dim lstrTradeMark As String
Dim lstrCreated As String
Dim fso As New FileSystemObject

    If fso.FileExists(App.Path & "\" & App.EXEName & ".exe") Then
        lstrCreated = "Created on " & Format$(FileDateTime(App.Path & "\" & App.EXEName & ".exe"), "mm/dd/yyyy hh:nn:ss AM/PM")
    Else
        lstrCreated = "Executable not found"
    End If
    lstrText = "Version: " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & vbNewLine _
        & lstrCreated & vbNewLine & vbNewLine _
        & "Company: " & App.CompanyName & vbNewLine & vbNewLine _
        & "Copyright: " & App.LegalCopyright & vbNewLine & vbNewLine _
        & "Licensed: " & Get_UserName & " on " & Get_ComputerName & " Workstation" & vbNewLine & vbNewLine _
        & "Trademark: " & App.LegalTrademarks & vbNewLine & vbNewLine _
        & "Psychiatric Content Assisted by: Richard M. Bunt, M.D., Atlanta, Georgia"
    MsgBox lstrText, vbInformation, "About " & App.ProductName
    Exit Sub
End Sub

Private Sub mnuHelp_SystemInfo_Click()
Dim lstrText
Dim lstrOSdir As String
Dim lstrPlatform As String

    lstrOSdir = Get_WindowsSystemDirectory
    
    Select Case SysInfo.OSPlatform
         Case 0
             lstrPlatform = "Unidentified > " & lstrOSdir
         Case 1
             lstrPlatform = "Windows 95, Version " & CStr(SysInfo.OSVersion)
         Case 2
             lstrPlatform = "Windows NT, Version " & CStr(SysInfo.OSVersion) & ", Build " & CStr(SysInfo.OSBuild)
     End Select

    'lstrText = "Platform > " & lstrPlatform & vbNewLine _
        & "Version > " & SysInfo.OSVersion & vbNewLine _
        & "Build > " & SysInfo.OSBuild
        
        Debug.Print SysInfo.Name
        
        
  MsgBox lstrPlatform, vbInformation, "System Information"
End Sub


Private Sub mnuImages_Click()
    frmDialogSetFilePath.Show vbModal
End Sub

Private Sub mnuPatientProfile_Click()
    frmFaceSheet.Show vbModal
    Exit Sub
End Sub

Private Sub mnuPatientSelect_Click()
            
    frmPatients.Show
End Sub

Private Sub mnuReport_DaySheet_Click()


    'strFinance = "SOFTWARE\Denge\Finance"
    'strOption = QueryValue(HKEY_LOCAL_MACHINE, strFinance, "Calendar")
    Load_DaySheet
'    If strOption <> vbNullString Then
'        strDate = Format(Now, "mm/dd/yyyy")
'        strTermDate = Format(strOption, "mm/dd/yyyy")
'        lngDateDiff = DateDiff("d", strDate, strTermDate)
'        If lngDateDiff < 10 Then
'            MsgBox "License for this function has expired, Please contact PatientTrac!", vbInformation
'            Exit Sub
'        Else
'
'        End If
'    Else
'        MsgBox "PatientTrac PMS - You do not have a license for this product, Please contact PatientTrac Sales!", vbInformation
'        Exit Sub
'    End If
    Exit Sub
End Sub




Private Sub mnuRoomLocs_Click()
    frmRoomLocations.Show vbModal
    Exit Sub
End Sub

Private Sub mnuScreens_PatientData_Click()
    mstrPatOcx = "FN_RevSys"
    mstrPatCtl = "ScanImages"
    Load frmLoadControl
    frmLoadControl.Show vbModal
    Exit Sub
End Sub

Private Sub mnuSysConfig_Click()
   frmSysConfig.Show vbModal
End Sub


Private Sub Timer1_Timer()
 
    On Error GoTo Error_Timer
 
'    If booAppTime Then
'        Timer1.Enabled = True
'        Screen.MousePointer = vbDefault
'        If Now <= datTimer Then Exit Sub
'        If IsFormLoaded("frmInPatientAppt") Then frmInPatientAppt.Hide
'        If IsFormLoaded("frmPatientAppointment") Then frmPatientAppointment.Hide
'        If IsFormLoaded("frmNursingHome") Then frmNursingHome.Hide
'        frmLogin.mlngReturn = mlngLogonUser
'        frmLogin.Show vbModal
'        datTimer = DateAdd("n", 5, Now)
'        If IsFormLoaded("frmInPatientAppt") Then frmInPatientAppt.Show
'        If IsFormLoaded("frmPatientAppointment") Then frmPatientAppointment.Show
'        If IsFormLoaded("frmNursingHome") Then frmNursingHome.Show
'    Else
'        datTimer = DateAdd("n", 5, Now)
'    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Timer:
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume Next
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iConnect As Integer
Dim iFormID As Integer
Dim i As Integer
Dim strSQL As String
Dim varRS
Dim lngRows As Long
Dim booRet As Boolean
Dim lstrRemarks As String
Dim strFormName As String
Dim strMARID As String

    On Error GoTo Error_Load_Forms
    Screen.MousePointer = vbHourglass
    
   If IsFormLoaded("frmEnctr") Then Unload frmEnctr
    Select Case Button.Key
        Case "Patient Data"
        If piPatientID < 1 Then
           MsgBox "Select a Patient First!", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Case "Inquiry FollowUp"
            mstrPatContact = "Contact"
            mstrPatOcx = "PatientTrac_Report"
            mstrPatCtl = "PatientContact"
            frmLoadControl.Show
        Case "StimSheet"
            If mstrPractice = "ObGyn" Then
                If piPatientID < 1 Then
                    MsgBox "Select a Patient First!", vbExclamation
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If piEncounterID < 1 Then
                    MsgBox "Select an Encounter first.", vbInformation
                    Exit Sub
                End If
                frmStimSheet.Show
            End If
        Case "Select Patient"
            piPatientID = 0
            StatusBar.Panels(1).Text = vbNullString
            StatusBar.Panels(2).Text = vbNullString
            StatusBar.Panels(3).Text = vbNullString
            StatusBar.Panels(4).Text = vbNullString
            Get_Providers mlngLogonUser
            StatusBar.Panels(6).Text = "Provider: " & clsProvider.Name
            StatusBar.Panels(5).Text = "Visit Date: " & Format(Now, "MMMM DD, YYYY")
            frmPatients.Show
        Case "Contact Notes"
            mstrPatOcx = "PatientTrac_Report"
            mstrPatCtl = "ContactNote"
            frmLoadControl.Show
        Case "Surgery Log"
            mstrPatOcx = "PatientTrac_Report"
            mstrPatCtl = "SurgLog"
            frmLoadControl.Show
        Case "PatientInquiry"
            If mstrPractice = "Psychiatry" Then
                MsgBox "You do not have appropriate license for this functionality.", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                mstrPatOcx = "FN_Orders"
                mstrPatCtl = "PatientContact"
                Load frmLoadControl
                frmLoadControl.Show vbModal
            End If
        Case "SuperBill"
            If piPatientID < 1 Then
                MsgBox "Select a Patient First!", vbExclamation
                Screen.MousePointer = vbDefault
                If Not IsFormLoaded("frmNursingHome") Then
                     frmMDI.Get_AppointmentFormat
                 End If
                Exit Sub
            End If
            Screen.MousePointer = vbDefault
            mstrPatOcx = "FN_PsyMse"
            mstrPatCtl = "SuperBill"
            Load frmLoadControl
            frmLoadControl.Show vbModal
        Case "Create Encounter"
            If piPatientID < 1 Then
                MsgBox "Select a Patient First!", vbExclamation
                Exit Sub
            End If
            frmEnctr.piPatientID = piPatientID
            frmEnctr.Show vbModal
        Case "Print Patient"
            If piPatientID < 1 Then
                MsgBox "Select a Patient First!", vbExclamation
                Exit Sub
            End If
        Case "Tablet Notes"
            If Not IsFormLoaded("frmNursingHome") Then
                Unload frmPatientAppointment
                Load frmNursingHome
                frmNursingHome.Show
                Exit Sub
'            Else
'                Get_Providers mlngLogonUser
'                frmBrowser.Show
            End If
        Case "Viewer"
            mstrPatOcx = "FN_RevSys"
            mstrPatCtl = "ScanImages"
            frmLoadControl.Show
        Case "Serial Communications"
            Launch_Terminal
        Case "MAR"
            mstrPatOcx = "PT_Medications"
            mstrPatCtl = "MedicationAdmin"
            Load frmLoadControl
            frmLoadControl.Show vbModal
         Case "DSM-IV"
            frmDSMIV.mstrConnectStringTerms = psConnect(1)
            If mstrPractice = "ObGyn" Then
                frmDiagnosis.Show
            Else
                frmDSMIV.Show
            End If
        Case "Medications"
            Screen.MousePointer = vbDefault
            frmMeds.mstrSource = "MDI"
            frmMeds.Show
        Case "Appointments"
            Screen.MousePointer = vbDefault
            If IsFormLoaded("frmNursingHome") Then
                Unload frmNursingHome
            End If
            Get_AppointmentFormat
'         If clsEDI_Client.Site_ID = "1125" Then
'            frmInPatientAppt.Show
'        Else
'            frmPatientAppointment.Prepare_Grid
'            frmPatientAppointment.Show
'         End If
    End Select
'    If IsFormLoaded("frmPatientAppointment") Then
'        frmPatientAppointment.ZOrder 0
'    ElseIf IsFormLoaded("frmInPatientAppt") Then
'        frmInPatientAppt.ZOrder 0
'    End If
    Screen.MousePointer = vbDefault
    Exit Sub
Error_Load_Forms:
    Exit Sub
    Resume Next
End Sub



Private Function Insert_Spaces(vstrText As String) As String
Dim lstrChar As String
Dim i As Integer
Insert_Spaces = ""
i = 1
    While i < Len(vstrText)
        i = i + 1
        lstrChar = Mid(vstrText, i, 1)
        'Is lstrChar a Captal letter?...............
        If lstrChar = UCase(lstrChar) Then
            Insert_Spaces = Insert_Spaces & Left(vstrText, i - 1) & Space(1)
            vstrText = Mid(vstrText, i)
            i = 1
        End If
    Wend
     Insert_Spaces = Trim(Insert_Spaces & vstrText)
End Function

Private Sub Notes()

End Sub



Private Sub CreatePatientSignImage(iPatientID As Long, iEncounterID As Long, sImagePath As String)
Dim lconClinical As New ADODB.Connection
Dim lrstClinical As New ADODB.Recordset
Dim lstrSignatureToken As String
Dim lbooInvalidData As Boolean
Dim strSQL As String
Dim vRS
Dim bRet As Boolean
Dim iRows As Long
Dim iActivex_id As Integer
Dim lstrRemarks As String

    On Error GoTo Error_CreatePatientSignImage
    Screen.MousePointer = vbHourglass
    
'    Dim picSignature As PictureBox
'    Dim ccaSignature As ccApprove
'    lconClinical.open psConnect(0)
'    strSQL = "Select ACTX_ID from ACTIVE_X where ACTX_NAME = 'BusinessGUI' and " _
'        & "ACTX_CLASS = 'IUIPatientSignature'"
'    bRet = adoSelectQuery(strSQL, psConnect(1), vRS, , iRows)
'    iActivex_id = 0
'    If iRows > 0 Then iActivex_id = vRS(0, 0)
'    'Call stored procedure to retrieve component's clinical data
'     lconClinical.S_ENC_SIGNATURES _
'     "S", iPatientID, iEncounterID, iActivex_id, lrstClinical
'
'
'    If Not lrstClinical.EOF Then
'       If (lrstClinical!SIGNATURE_REF) = "IUIPatientSignature" Then
'          lstrSignatureToken = (lrstClinical!Signature)
'       End If
'    End If
'    If Len(lstrSignatureToken) = 0 Then Exit Sub 'if no signature exit procedure
'    'This code place the these following controls at runtime to do picture saving work
'    'added new following 2 lines
'    'Changed on 12/10/2002
'    Set picSignature = Controls.Add("VB.PictureBox", "picSignature")
'    Set ccaSignature = Controls.Add("DedicatedControls.ccApprove", "ccAsignature", picSignature)
'    lbooInvalidData = ccaSignature.DecodeSignatureValues(lstrSignatureToken, ".")
'    'picSignature.Visible = True
'    Set picSignature.Picture = ccaSignature.Picture
'    SavePicture picSignature.Image, sImagePath
'    'After finishing saving remove these controls
'    Controls.Remove ccaSignature
'    Controls.Remove picSignature
'    Set lrstClinical = Nothing
'    lconClinical.Close
'    Set lconClinical = Nothing
'    'We have to set nothing only when its hold some object
'    If Not picSignature Is Nothing Then Set picSignature = Nothing
'    If Not ccaSignature Is Nothing Then Set ccaSignature = Nothing
    Screen.MousePointer = vbDefault
    
    Exit Sub
Error_CreatePatientSignImage:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "CreatePatientSignImage", Me.Name
    Exit Sub
    Resume
End Sub


Public Function GetDefaultPractice()
Dim strSQL As String
Dim vRS
Dim iRows As Long
Dim bRet As Boolean
    On Error GoTo Error_GetDefaultPractice
    Screen.MousePointer = vbHourglass
    strSQL = "Select EMR_PRACTICE, EMR_GRP_ID, EMR_SPECIALTY " _
        & "From EMR_GROUP " _
        & "Where EMR_GRP_NAME = 'Clinical Data' " _
        & "And EMR_PRACTICE like '%Default'"
        bRet = SelectQuery(strSQL, psConnect(1), vRS, iRows)
        If iRows > 0 Then
            GetDefaultPractice = vRS(0, 0)
            mstrPractice = vRS(2, 0)
        Else
            frmSelPractice.Show vbModal
        End If
    Screen.MousePointer = vbDefault
    Exit Function
Error_GetDefaultPractice:
    Screen.MousePointer = vbDefault
    Write_Error Err.Description, Err.Number, strSQL, "GetDefaultPractice", Me.Name
    Exit Function
    Resume
End Function
Private Sub GenerateDefaultFile()
Dim sText As String
Dim iFileNum As Integer

    sText = "[General]" & vbNewLine _
    & "DefaultProvider=" & vbNewLine _
    
 'sText = sText & "[ODBC]" & vbNewLine _
    & "Driver=SQL Server" & vbNewLine _
    & "Server=" & vbNewLine _
    & "LastUser=" & vbNewLine _
    & "LastLogin=" & vbNewLine _
    & "Database0=TERMS" & vbNewLine _
    & "Database1=CLINICAL_REPOSITORY" & vbNewLine _
    & "RelativePath=ActiveX" & vbNewLine _
    & "LastPath=" & vbNewLine & vbNewLine _
    & "[SERVERS]" & vbNewLine _
    & "Server0=" & vbNewLine & vbNewLine _
    & "[OCX]" & vbNewLine _
    & "Name=" & vbNewLine _
    & "Control=" & vbNewLine _
    & "PatientID=" & vbNewLine _
    & "EncounterID=" & vbNewLine _
    & "OCX_FilePath="
        
    iFileNum = FreeFile

    Open psIniFile For Output As iFileNum
        Print #iFileNum, sText
    Close iFileNum

End Sub


Public Sub Launch_Terminal()
Dim strExe As String
Dim booRegister As Boolean
Dim strActX As String
Dim fso As New FileSystemObject

    On Error GoTo Launch_Terminal_Error
    Set fobjTerminal = Nothing
    strExe = fstrMainDrive & "\ToolBox\LabCom.exe"
    If fso.FileExists(strExe) Then
        strActX = "LabCom.clsLab"
        Set fobjTerminal = CreateObject(strActX)
        If booRegister Then
            RegisterObject strExe
            Set fobjTerminal = CreateObject(strActX)
            booRegister = False
        End If
        'fobjTerminal.Set_Object Me.txtValue
        'fobjTerminal.FileToTransfer = App.Path & "\NSF.txt"
        'fbObjectLoaded = True
    Else
        MsgBox "LabCom.exe does not exist.", vbInformation
        Exit Sub
    End If
    Exit Sub
Launch_Terminal_Error:
    booRegister = True
    Resume Next

End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

Dim fso As New FileSystemObject
    Select Case ButtonMenu.Key
        Case "LabCorp"
            If fso.FileExists(fstrMainDrive & "\Toolbox\LabCorp.exe") Then
                WinExec fstrMainDrive & "\Toolbox\LabCorp.exe", 0
            Else
                MsgBox "LabCorp application missing.", vbExclamation, "Outside Lab Orders"
            End If
        Case "Quest Diagnostics"
            If fso.FileExists(fstrMainDrive & "\Toolbox\QuestDiagnostics.exe") Then
                WinExec fstrMainDrive & "\Toolbox\QuestDiagnostics.exe", 0
            Else
                MsgBox "Quest Diagnostics application missing.", vbExclamation, "Outside Lab Orders"
            End If
    End Select
    Exit Sub
End Sub



