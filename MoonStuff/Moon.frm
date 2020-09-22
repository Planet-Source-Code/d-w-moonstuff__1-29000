VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Moon 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Moon Phase"
   ClientHeight    =   6780
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   8940
   DrawWidth       =   2
   Icon            =   "Moon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   596
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox BlueMoon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   8055
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   17
      Top             =   885
      Width           =   750
   End
   Begin VB.TextBox DateSec 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   8070
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Sec"
      Top             =   315
      Width           =   360
   End
   Begin MSComCtl2.UpDown SpinSec 
      Height          =   255
      Left            =   8460
      TabIndex        =   14
      Top             =   330
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox DateMin 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7170
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Min"
      Top             =   315
      Width           =   435
   End
   Begin MSComCtl2.UpDown SpinMin 
      Height          =   255
      Left            =   7620
      TabIndex        =   12
      Top             =   330
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox DateHour 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6285
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Hour"
      Top             =   315
      Width           =   435
   End
   Begin MSComCtl2.UpDown SpinHour 
      Height          =   255
      Left            =   6735
      TabIndex        =   10
      Top             =   330
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox DateDay 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4425
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Day"
      Top             =   315
      Width           =   450
   End
   Begin MSComCtl2.UpDown SpinDay 
      Height          =   255
      Left            =   4890
      TabIndex        =   8
      Top             =   330
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      OrigLeft        =   6075
      OrigTop         =   75
      OrigRight       =   6315
      OrigBottom      =   330
      Enabled         =   -1  'True
   End
   Begin VB.TextBox DateMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Month"
      Top             =   315
      Width           =   885
   End
   Begin MSComCtl2.UpDown SpinMonth 
      Height          =   255
      Left            =   3945
      TabIndex        =   6
      Top             =   330
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      OrigLeft        =   5190
      OrigTop         =   75
      OrigRight       =   5430
      OrigBottom      =   330
      Enabled         =   -1  'True
   End
   Begin VB.TextBox DateYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5370
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Year"
      Top             =   315
      Width           =   495
   End
   Begin MSComCtl2.UpDown SpinYear 
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "Use the enter key to scroll faster through the years."
      Top             =   330
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      OrigLeft        =   3930
      OrigTop         =   75
      OrigRight       =   4170
      OrigBottom      =   330
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton StopGo 
      Appearance      =   0  'Flat
      Caption         =   "STOP"
      Height          =   1020
      Left            =   2760
      Picture         =   "Moon.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Start and stop animation."
      Top             =   2715
      Width           =   885
   End
   Begin VB.PictureBox ThePhase 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   60
      Picture         =   "Moon.frx":09BC
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   1
      Top             =   3390
      Width           =   3000
   End
   Begin VB.PictureBox TheMoon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   0
      Picture         =   "Moon.frx":585E
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   15
      Width           =   3000
   End
   Begin VB.Timer Phase 
      Interval        =   200
      Left            =   1290
      Top             =   1080
   End
   Begin VB.Timer Starter 
      Interval        =   1000
      Left            =   735
      Top             =   1080
   End
   Begin MSComctlLib.ImageList MoonPics 
      Left            =   915
      Top             =   405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":A700
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":A77C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":A7F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":A92C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":AB20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":AD74
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":B038
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":B37C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":B724
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":BB2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":BF9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":C44C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":C960
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":CE98
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":D40C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":D998
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":DF4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":E4C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":EA40
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":EF48
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":F3EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":F85C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":FC5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":1000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":1035C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":1062C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":10898
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":10A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Moon.frx":10C48
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label(1)"
      ForeColor       =   &H80000009&
      Height          =   195
      Index           =   1
      Left            =   4245
      TabIndex        =   16
      Top             =   840
      Width           =   570
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Month                    Day              Year              Hours          Minutes        Seconds"
      ForeColor       =   &H80000009&
      Height          =   240
      Index           =   0
      Left            =   3060
      TabIndex        =   15
      Top             =   30
      Width           =   5730
   End
End
Attribute VB_Name = "Moon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Up As Boolean





Private Sub Calculate()
TheDate = CDate(Str(MonthNumber(DateMonth)) & _
"/" & DateDay & "/" & DateYear & " " & DateHour & _
":" & DateMin & ":" & DateSec)
Label(0) = "Month                    Day              Year              Hours          Minutes        Seconds"
Label(1) = "System Time = " & UTtoLocal(TheDate)
Label(2) = "Universal Time = " & Format(TheDate, "mm/dd/yyyy hh:mm:ss") & " UTC"
Label(3) = "Julian Days = " & Format(UTtoJulianDays(TheDate), "0.00000")
Label(4) = "Time Zone = " & TimeZone
Label(5) = "Weekday = " & WeekdayName(Weekday(TheDate))
Label(6) = "Leap Year = " & IsLeapYear(Year(TheDate))
Label(7) = "Next Leap Year = " & NextLeapYear(TheDate)
Label(8) = "Age Of Lunation = " & Age(TheDate) & " Days"
Label(9) = "Angle Of Illumination = " & Angle(TheDate) & "°"
Label(10) = "Percent Of Illumination = " & Illum(TheDate) & "%"
Label(11) = "Lunation Number = " & Lunation(TheDate)
Label(12) = "Moon Phase = " & MoonDescription(TheDate)
Label(13) = "Previous New Moon = " & Format(PreviousNewMoon(TheDate), "mm/dd/yyyy hh:mm:ss") & " UTC"
Label(14) = "Previous First Quarter = " & Format(PreviousFirstQuarter(TheDate), "mm/dd/yyyy hh:mm:ss") & " UTC"
Label(15) = "Previous Full Moon = " & Format(PreviousFullMoon(TheDate), "mm/dd/yyyy hh:mm:ss") & " UTC"
Label(16) = "Previous Last Quarter = " & Format(PreviousLastQuarter(TheDate), "mm/dd/yyyy hh:mm:ss") & " UTC"
Label(17) = "Next New Moon = " & Format(NextNewMoon(TheDate), "mm/dd/yyyy hh:mm:ss") & " UTC"
Label(18) = "Next First Quarter = " & Format(NextFirstQuarter(TheDate), "mm/dd/yyyy hh:mm:ss") & " UTC"
Label(19) = "Next Full Moon = " & Format(NextFullMoon(TheDate), "mm/dd/yyyy hh:mm:ss") & " UTC"
Label(20) = "Next Last Quarter = " & Format(NextLastQuarter(TheDate), "mm/dd/yyyy hh:mm:ss") & " UTC"
DrawMoonPhase Angle(TheDate), ThePhase
End Sub


Private Sub LoadLabels()
Dim i As Integer
Label(0).ForeColor = vbGreen
Label(1).ForeColor = vbGreen
Label(1).Height = Label(1).Height * Resize
Label(1).Width = Label(1).Width * Resize
Label(1).Left = Label(1).Left * Resize
Label(1).Top = Label(1).Top * Resize
For i = 2 To 20
Load Label(i)
Label(i).AutoSize = True
Label(i).Height = Label(1).Height
Label(i).Width = Label(1).Width
Label(i).Top = Label(i - 1).Top + Label(1).Height * 1.4
Label(i).Left = Label(1).Left
Label(i).ForeColor = vbGreen
Label(i).Visible = True
Next
End Sub

Public Sub SetDate()
TheDate = ConvertToUT(Now)
DateYear = Year(TheDate)
DateMonth = MonthName(Month(TheDate))
DateDay = Day(TheDate)
DateHour = Format(TheDate, "h")
DateMin = Right(Format(TheDate, "h:mm"), 2)
DateSec = Format(TheDate, "s")
Caption = "The Moon Phase Today: " & _
   WeekdayName(Weekday(TheDate)) & ", " & _
   DateMonth & " " & DateDay & ", " & _
   DateYear & " is " & MoonDescription(TheDate)
Calculate
End Sub

Private Sub DateYear_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Up Then
SpinYear_UpClick
Else
SpinYear_DownClick
End If
End If
End Sub







Private Sub Form_Load()
LoadLabels
SetDate
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim ListPic As IPictureDisp
Dim TempPath As String
Dim Phases As Integer
If UnloadMode = 0 Then
Cancel = True
Hide
TitleToTray Me
Stopping = True
Phase.Enabled = False
Phases = Int(Age(ConvertToUT(Now)))
If Phases = 0 Then Phases = 1
If Phases < 15 Then Phases = Phases + 1
If Phases > 29 Then Phases = 29
Set ListPic = MoonPics.ListImages(Phases).ExtractIcon
ChangeIcon Tray, ListPic
End If
End Sub


Private Sub Phase_Timer()
Dim ListPic As IPictureDisp
Dim NewIcon As Object
Static i As Integer
i = i + 1
If i > 29 Then i = 1
Set ListPic = Tray.MoonPics.ListImages(i).ExtractIcon
Set NewIcon = ListPic
BlueMoon.Move Me.ScaleWidth - MoonPics.ImageWidth, _
   Me.ScaleTop, MoonPics.ImageHeight + 10, MoonPics.ImageWidth + 10
MoonPics.ListImages(i).Draw BlueMoon.hDC
BlueMoon.Refresh
Icon = ListPic
ChangeIcon Tray, NewIcon
End Sub









Private Sub SpinDay_DownClick()
Dim i As Integer
Dim n As Integer
n = DaysInMonth(CDate(MonthNumber(DateMonth) & "/" & CInt(DateYear)))
i = CInt(DateDay) + 1
If i > n Then i = 1
DateDay = i
If CInt(DateYear) = 1582 Then
If MonthNumber(DateMonth) = 10 Then
If CInt(DateDay) = 5 Then DateDay = 15
End If
End If
Calculate
End Sub

Private Sub SpinDay_UpClick()
Dim i As Integer
Dim n As Integer
n = DaysInMonth(CDate(MonthNumber(DateMonth) & "/" & CInt(DateYear)))
i = CInt(DateDay) - 1
If i = 0 Then i = n
DateDay = i
If CInt(DateYear) = 1582 Then
If MonthNumber(DateMonth) = 10 Then
If CInt(DateDay) = 14 Then DateDay = 4
End If
End If
Calculate
End Sub


Private Sub SpinHour_DownClick()
Dim i As Integer
i = CInt(DateHour) + 1
If i = 24 Then i = 0
DateHour = i
Calculate
End Sub

Private Sub SpinHour_UpClick()
Dim i As Integer
i = CInt(DateHour) - 1
If i = -1 Then i = 23
DateHour = i
Calculate
End Sub


Private Sub SpinMin_DownClick()
Dim i As Integer
i = CInt(DateMin) + 1
If i = 60 Then i = 1
DateMin = i
Calculate
End Sub


Private Sub SpinMin_UpClick()
Dim i As Integer
i = CInt(DateMin) - 1
If i = 0 Then i = 59
DateMin = i
Calculate
End Sub

Private Sub SpinMonth_DownClick()
Dim i As Integer
Dim x As Integer
x = MonthNumber(DateMonth)
x = x + 1
If x = 13 Then x = 1
i = DaysInMonth(CDate(x & "/" & CInt(DateYear)))
If CInt(DateDay) > i Then DateDay = i
DateMonth = MonthName(x)
If CInt(DateYear) = 1582 Then
If MonthNumber(DateMonth) = 10 Then
DateDay = 1
End If
End If
Calculate
End Sub

Private Sub SpinMonth_UpClick()
Dim i As Integer
Dim x As Integer
x = MonthNumber(DateMonth)
x = x - 1
If x = 0 Then x = 12
i = DaysInMonth(CDate(x & "/" & CInt(DateYear)))
If CInt(DateDay) > i Then DateDay = i
DateMonth = MonthName(x)
If CInt(DateYear) = 1582 Then
If MonthNumber(DateMonth) = 10 Then
DateDay = 1
End If
End If
Calculate
End Sub


Private Sub SpinSec_DownClick()
Dim i As Integer
i = CInt(DateSec) + 1
If i = 60 Then i = 1
DateSec = i
Calculate
End Sub


Private Sub SpinSec_UpClick()
Dim i As Integer
i = CInt(DateSec) - 1
If i = 0 Then i = 59
DateSec = i
Calculate
End Sub


Private Sub SpinYear_DownClick()
Dim i As Integer
DateYear.SetFocus
Up = False
i = DaysInMonth(CDate(MonthNumber(DateMonth) & "/" & CInt(DateYear) + 1))
If CInt(DateDay) > i Then DateDay = i
DateYear = CInt(DateYear) + 1
If CInt(DateYear) = 1582 Then
If MonthNumber(DateMonth) = 10 Then
DateDay = 1
End If
End If
Calculate
End Sub


Private Sub SpinYear_UpClick()
Dim i As Integer
DateYear.SetFocus
Up = True
i = DaysInMonth(CDate(MonthNumber(DateMonth) & "/" & CInt(DateYear) - 1))
If CInt(DateDay) > i Then DateDay = i
DateYear = CInt(DateYear) - 1
If CInt(DateYear) = 1582 Then
If MonthNumber(DateMonth) = 10 Then
DateDay = 1
End If
End If
Calculate
End Sub


Private Sub Starter_Timer()
Static i As Integer
If i = 1 Then
DrawMoonPhase Angle(TheDate), ThePhase
End If
If i = 2 Then
CyclePhases TheMoon
Starter.Enabled = False
i = 0
End If
i = i + 1
End Sub


Private Sub StopGo_Click()
If StopGo.Caption = "STOP" Then
StopGo.Caption = "GO"
Stopping = True
Phase.Enabled = False
Else
Stopping = False
StopGo.Caption = "STOP"
Starter.Enabled = True
Phase.Enabled = True
End If
End Sub






