VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8131EC61-92BA-41BF-8BD9-990128770913}#1.0#0"; "PT_XP.ocx"
Begin VB.Form frmApptDate 
   Caption         =   "Please select desired Appointment Date"
   ClientHeight    =   4185
   ClientLeft      =   3840
   ClientTop       =   2520
   ClientWidth     =   6960
   FillColor       =   &H00800000&
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
   Icon            =   "frmApptDate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   6960
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   405
      Left            =   2543
      TabIndex        =   4
      Top             =   3060
      Width           =   2055
      Begin PT_XP_Button.PT_XP cmdSelect 
         Height          =   405
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmApptDate.frx":08CA
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Select"
         CaptionAlignment=   5
         CaptionPosition =   4
         DepthMode       =   1
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmApptDate.frx":0C64
         ShowFocus       =   -1  'True
      End
      Begin PT_XP_Button.PT_XP cmdClose 
         Height          =   405
         Left            =   1110
         TabIndex        =   6
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
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
         Picture         =   "frmApptDate.frx":0FB6
         AppearanceThemes=   3
         TransparentColor=   14215660
         Caption         =   "Close"
         DepthEvent      =   1
         ForeColorDisabled=   16761024
         ForeColorOver   =   8388608
         ForeColorFocus  =   8388608
         ForeColorDown   =   8388608
         PictureDisabled =   "frmApptDate.frx":19C8
         ShowFocus       =   -1  'True
      End
   End
   Begin VB.TextBox txtSelDate 
      Alignment       =   2  'Center
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
      Height          =   435
      Left            =   3720
      TabIndex        =   3
      Top             =   2310
      Width           =   3165
   End
   Begin VB.ComboBox cboTime 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "h:nn AM/PM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Height          =   330
      Left            =   3750
      TabIndex        =   1
      Top             =   690
      Width           =   2685
   End
   Begin MSComCtl2.MonthView ccSchedDate 
      Height          =   2670
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   4710
      _Version        =   393216
      ForeColor       =   8388608
      BackColor       =   16744448
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   68681729
      TitleBackColor  =   16761024
      TitleForeColor  =   8388608
      TrailingForeColor=   8421504
      CurrentDate     =   38827
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Appointment Times"
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
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   3060
   End
End
Attribute VB_Name = "frmApptDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFirst As Boolean
Public mstrApptDate As String
Public mstrApptTime As String

Private Sub cboTime_Click()

    If cboTime.Text <> vbNullString Then
        txtSelDate.Text = txtSelDate.Text & " " & cboTime.Text
        txtSelDate.Text = Format(txtSelDate.Text, "mmmm dd, yyyy hh:mm AM/PM")
    End If
    
    Exit Sub
End Sub


Private Sub ccSchedDate_DateClick(ByVal DateClicked As Date)
    
    If Not IsNull(ccSchedDate.Value) Then
        txtSelDate.Text = Format(ccSchedDate.Value, "mmmm dd, yyyy")
         Get_OpenAppointment Format(ccSchedDate.Value, "mm/dd/yyyy")
    End If
    
End Sub

Private Function Get_OpenAppointment(vstrDate As String)
Dim i%
Dim strSQL As String
Dim vstrTime As String
Dim lngProviderID As Long
Dim sValue As String
Dim IntCount As Integer
Dim iAnswer As Integer
Dim strTime As String
Dim strInt As String
Dim lngTime As Long
Dim lngIDX As Long
Dim Arr
Dim strDuration As String
Dim intCboIdx As Integer

    Open_DB psConnect(0)
    cboTime.Clear
    Load_ApptTime
    strSQL = "Select APPT_TIME, DURATION from APPOINTMENTS where APPT_DATE = '" & vstrDate & "' AND " _
            & "PROVIDER_ID = " & piProviderID & " AND " _
            & "FACILITY_ID = " & piFacilityID & " AND ACTIVE = 1"
    Open_DB psConnect(0)
    Set rs = cn.Execute(strSQL)
    IntCount = 0
    With rs
        Do Until .EOF
            If (!Appt_Time) <> vbNullString Then
                strTime = (!Appt_Time)
                strTime = Format(strTime, "hh:mm AM/PM")
                If (!Duration) <> vbNullString Then strDuration = (!Duration)
                        strInt = "EVERY 20 MIN"
                        If Trim(UCase(clsApptPref.SchedInterval)) = Trim(strInt) Then
                            lngTime = 20
                        ElseIf UCase(clsApptPref.SchedInterval) = "EVERY 15 MIN" Then
                            lngTime = 15
                        End If
                            If strDuration = "1 Hour" Then strDuration = "60 minutes"
                            Arr = Split(strDuration, " ")
                            If UBound(Arr) > 0 Then
                                If Arr(0) = "15" Then lngIDX = Val(Val(Arr(0)) / lngTime)
                                If Arr(0) = "20" Then lngIDX = Val(Val(Arr(0)) / lngTime)
                                If Arr(0) = "30" Then lngIDX = Val(Val(Arr(0)) / lngTime)
                                If Arr(0) = "45" Then lngIDX = Val(Val(Arr(0)) / lngTime)
                                If Arr(0) = "60" Then lngIDX = Val(Val(Arr(0)) / lngTime)
                            End If
                For i = 0 To cboTime.listcount - 1
                    If cboTime.List(i) = strTime Then
                        cboTime.RemoveItem (i)
                        If lngIDX > 1 Then
                            intCboIdx = lngIDX - 1
                            cboTime.RemoveItem (i + intCboIdx)
                        End If
                    End If
                Next i
            End If
        .MoveNext
        Loop
    End With
    
    Close_DB psConnect(0)
    Exit Function
End Function


Private Function Load_ApptTime()
    If cboTime.listcount <= 0 Then
         If UCase(clsApptPref.SchedInterval) = "EVERY 15 MIN" Then
               cboTime.AddItem "08:00 AM": cboTime.AddItem "08:15 AM": cboTime.AddItem "08:30 AM": cboTime.AddItem "08:45 AM"
               cboTime.AddItem "09:00 AM": cboTime.AddItem "09:15 AM": cboTime.AddItem "09:30 AM": cboTime.AddItem "09:45 AM"
               cboTime.AddItem "10:00 AM": cboTime.AddItem "10:15 AM": cboTime.AddItem "10:30 AM": cboTime.AddItem "10:45 AM"
               cboTime.AddItem "11:00 AM": cboTime.AddItem "11:15 AM": cboTime.AddItem "11:30 AM": cboTime.AddItem "11:45 AM"
               cboTime.AddItem "12:00 PM": cboTime.AddItem "12:15 PM": cboTime.AddItem "12:30 PM": cboTime.AddItem "12:45 PM"
               cboTime.AddItem "01:00 PM": cboTime.AddItem "01:15 PM": cboTime.AddItem "01:30 PM": cboTime.AddItem "01:45 PM"
               cboTime.AddItem "02:00 PM": cboTime.AddItem "02:15 PM": cboTime.AddItem "02:30 PM": cboTime.AddItem "02:45 PM"
               cboTime.AddItem "03:00 PM": cboTime.AddItem "03:15 PM": cboTime.AddItem "03:30 PM": cboTime.AddItem "03:45 PM"
               cboTime.AddItem "04:00 PM": cboTime.AddItem "04:15 PM": cboTime.AddItem "04:30 PM": cboTime.AddItem "04:45 PM"
               cboTime.AddItem "05:00 PM": cboTime.AddItem "05:15 PM": cboTime.AddItem "05:30 PM": cboTime.AddItem "05:45 PM"
               cboTime.AddItem "06:00 PM": cboTime.AddItem "06:15 PM": cboTime.AddItem "06:30 PM": cboTime.AddItem "06:45 PM"
               cboTime.AddItem "07:00 PM": cboTime.AddItem "07:15 PM": cboTime.AddItem "07:30 PM": cboTime.AddItem "07:45 PM"
               cboTime.AddItem "08:00 PM": cboTime.AddItem "08:15 PM": cboTime.AddItem "08:30 PM": cboTime.AddItem "08:45 PM"
        ElseIf UCase(clsApptPref.SchedInterval) = "EVERY 20 MIN" Then
             cboTime.AddItem "08:00 AM": cboTime.AddItem "08:20 AM": cboTime.AddItem "08:40 AM"
             cboTime.AddItem "09:00 AM": cboTime.AddItem "09:20 AM": cboTime.AddItem "09:40 AM"
             cboTime.AddItem "10:00 AM": cboTime.AddItem "10:20 AM": cboTime.AddItem "10:40 AM"
             cboTime.AddItem "11:00 AM": cboTime.AddItem "11:20 AM": cboTime.AddItem "11:40 AM"
             cboTime.AddItem "12:00 PM": cboTime.AddItem "12:20 PM": cboTime.AddItem "12:40 PM"
             cboTime.AddItem "01:00 PM": cboTime.AddItem "01:20 PM": cboTime.AddItem "01:40 PM"
             cboTime.AddItem "02:00 PM": cboTime.AddItem "02:20 PM": cboTime.AddItem "02:40 PM"
             cboTime.AddItem "03:00 PM": cboTime.AddItem "03:20 PM": cboTime.AddItem "03:40 PM"
             cboTime.AddItem "04:00 PM": cboTime.AddItem "04:20 PM": cboTime.AddItem "04:40 PM"
             cboTime.AddItem "05:00 PM": cboTime.AddItem "05:20 PM": cboTime.AddItem "05:40 PM"
             cboTime.AddItem "06:00 PM": cboTime.AddItem "06:20 PM": cboTime.AddItem "06:40 PM"
             cboTime.AddItem "07:00 PM": cboTime.AddItem "07:20 PM": cboTime.AddItem "07:40 PM"
             cboTime.AddItem "08:00 PM": cboTime.AddItem "08:20 PM": cboTime.AddItem "08:40 PM"
          End If
    End If
    Exit Function
End Function


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If Not IsNull(ccSchedDate.Value) Then mstrApptDate = Format(ccSchedDate.Value, "mm/dd/yyyy")
    If cboTime.Text <> vbNullString Then mstrApptTime = cboTime.Text
    frmSchedPatient.Read_Data
    Unload Me
    frmSchedPatient.Show
    Exit Sub
End Sub


Private Sub Form_Activate()
    If mFirst Then
       mFirst = False
    End If
    Load_ApptTime
    centerForm Me
    Exit Sub
End Sub

Private Sub Form_Load()
    mFirst = True
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form, ctl As Control, obj As Object
  On Error Resume Next

    Unload Me
    Exit Sub

End Sub


