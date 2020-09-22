VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tray 
   Caption         =   "Moon Stuff"
   ClientHeight    =   3180
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4680
   Icon            =   "Tray.frx":0000
   LinkTopic       =   "Tray"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.ImageList MoonPics 
      Left            =   2010
      Top             =   945
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
            Picture         =   "Tray.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":04BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":0536
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":066E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":0862
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":0AB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":0D7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":10BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":1466
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":186E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":1CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":218E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":26A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":2BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":314E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":36DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":3C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":420A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":4782
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":4C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":512E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":559E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":599E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":5D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":609E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":636E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":65DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":67DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tray.frx":698A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuHide 
         Caption         =   "Hide Stuff"
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "Edit"
      Begin VB.Menu menuMoon 
         Caption         =   "Moon Stuff"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit2 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim ListPic As IPictureDisp
Dim Phases As Integer
Phases = Round(Age(ConvertToUT(Now)))
If Phases = 0 Then Phases = 1
If Phases < 15 Then Phases = Phases + 1
If Phases > 29 Then Phases = 29
Set ListPic = MoonPics.ListImages(Phases).ExtractIcon
Icon = ListPic
PlaceIcon Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Msg As Long
Msg = x / Screen.TwipsPerPixelX
Select Case Msg
Case WM_LBUTTONDOWN
Case WM_LBUTTONUP
If Not Moon.Visible Then
    If Moon.StopGo.Caption = "STOP" Then
    Stopping = False
    Moon.Phase.Enabled = True
    Moon.Starter.Enabled = True
    End If
TrayToTitle Moon
Moon.Show
Moon.SetDate
End If
Case WM_LBUTTONDBLCLICK
Case WM_RBUTTONUP
If Moon.Visible Then
PopupMenu Tray.menuFile
Else
PopupMenu Tray.menuEdit
End If
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DestroyIcon Me
End Sub


Private Sub menuExit_Click()
Unload Moon
Unload Me
End
End Sub


Private Sub menuExit2_Click()
Unload Moon
Unload Me
End
End Sub

Private Sub menuHide_Click()
Dim Phases As Integer
Dim ListPic As IPictureDisp
Moon.Hide
TitleToTray Moon
Stopping = True
Moon.Phase.Enabled = False
Phases = Round(Age(ConvertToUT(Now)))
If Phases = 0 Then Phases = 1
If Phases < 15 Then Phases = Phases + 1
If Phases > 29 Then Phases = 29
Set ListPic = MoonPics.ListImages(Phases).ExtractIcon
ChangeIcon Tray, ListPic
End Sub

Private Sub menuMoon_Click()
If Moon.StopGo.Caption = "STOP" Then
Stopping = False
Moon.Phase.Enabled = True
Moon.Starter.Enabled = True
End If
TrayToTitle Moon
Moon.Show
Moon.SetDate
End Sub




