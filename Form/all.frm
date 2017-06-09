VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Player 2.4"
   ClientHeight    =   1590
   ClientLeft      =   135
   ClientTop       =   -1185
   ClientWidth     =   9255
   Icon            =   "all.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   9255
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   7320
      Top             =   840
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "all.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "all.frx":175C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      LargeChange     =   10
      Left            =   3120
      SmallChange     =   10
      TabIndex        =   0
      Top             =   0
      Width           =   4035
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   3615
      Begin VB.Label Lbl1 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   20
         Top             =   150
         Width           =   495
      End
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   495
         Left            =   0
         TabIndex        =   19
         Top             =   -30
         Width           =   3100
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   5
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   100
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   0   'False
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   5477
         _cy             =   873
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000013&
      Height          =   495
      Left            =   6240
      TabIndex        =   14
      Top             =   0
      Width           =   1095
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   240
         X2              =   480
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   480
         X2              =   720
         Y1              =   360
         Y2              =   120
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   720
         X2              =   480
         Y1              =   360
         Y2              =   120
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H008080FF&
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   16
         Top             =   120
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   720
         MousePointer    =   5  'Size
         Picture         =   "all.frx":25EE
         Stretch         =   -1  'True
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00404040&
         Caption         =   "\/"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6960
      Top             =   840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   3520
      MouseIcon       =   "all.frx":28F8
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   -240
      Width           =   2895
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Back"
         Height          =   225
         Index           =   13
         Left            =   900
         MousePointer    =   1  'Arrow
         TabIndex        =   24
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cradieat"
         Height          =   225
         Index           =   9
         Left            =   660
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Avi"
         Height          =   225
         Index           =   16
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Convert"
         Height          =   225
         Index           =   15
         Left            =   2040
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Setting"
         Height          =   225
         Index           =   14
         Left            =   2040
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Directory"
         Height          =   225
         Index           =   12
         Left            =   705
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ViewList"
         Height          =   225
         Index           =   11
         Left            =   180
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Video"
         Height          =   225
         Index           =   6
         Left            =   1515
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "About"
         Height          =   225
         Index           =   5
         Left            =   2205
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "PlayList"
         Height          =   225
         Index           =   4
         Left            =   1500
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Search"
         Height          =   225
         Index           =   3
         Left            =   1360
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Open"
         Height          =   225
         Index           =   2
         Left            =   150
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   480
         Width           =   450
      End
   End
   Begin VB.Label X 
      Height          =   255
      Left            =   8880
      TabIndex        =   23
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Text1 
      Height          =   255
      Left            =   8520
      TabIndex        =   22
      Top             =   240
      Width           =   255
   End
   Begin VB.Label URL 
      Height          =   255
      Left            =   7800
      TabIndex        =   21
      Top             =   240
      Width           =   255
   End
   Begin VB.Label F 
      Height          =   255
      Left            =   8160
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------
Dim c         As Boolean: Dim g          As Boolean
Dim X1        As Single:  Dim g1         As Single
Dim Y1        As Single:  Dim g2         As Single
Dim n         As Single:  Dim xT         As Long
Dim x_t       As Long: Public For6       As Boolean
Dim Ca        As Boolean: Dim M          As Single
Dim lR        As Long
Private Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
As Long
          If Topmost = True Then 'Make the window topmost
             SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
         Else '-------------------------------
              SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
             SetTopMostWindow = False
         End If '-------------------------------
End Function
Private Sub Form_Activate()
 Activ
   If FForm1.SerTxt.Text = "Min" Then FForm1.SerTxt.Text = "Normal"
End Sub
Private Sub Activ()
On Error Resume Next
   lR = SetTopMostWindow(Form1.hwnd, True)
End Sub

Private Sub Form_Click()
Activ
End Sub

Private Sub Form_GotFocus()
Activ
End Sub

Private Sub Form_Load()
Activ
Me.Left = 5000: Me.Top = 0: c = False
Me.Height = 420: Frame1.Top = -300
Me.Width = 7230: Frame1.Left = 3520
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Lbl1(2).BackColor = &HC0FFC0
End Sub

Private Sub Form_Resize()
   If FForm1.SerTxt.Text = "Min" Then FForm1.SerTxt.Text = "Normal"
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
g = True: g1 = Y
Frame1.MouseIcon = ImageList1.ListImages(2).Picture
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim d As Single, i% '------------------Move Form On the Screen----------------------
d = ((Frame1.Top + Y) - g1)
        If d < 80 And d > -720 Then
        If g = True Then Call Frame1.Move(3520, d)
        End If
6 '------------------Set Labels Orginal Color----------------------------------------
For i = 1 To 18
If (i = 1) Or (i = 7) Or (i = 8) Or (i = 10) Then GoTo 2
Lbl1(i).BackColor = &HC0FFC0
2 Next
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
g = False: Frame1.MouseIcon = ImageList1.ListImages(1).Picture
End Sub

Private Sub HScroll1_Change()
On Error Resume Next
If (((WindowsMediaPlayer1.Controls.currentPosition) - HScroll1.Value) > 5) Or _
   (((WindowsMediaPlayer1.Controls.currentPosition) - HScroll1.Value) < -5) Then _
      WindowsMediaPlayer1.Controls.currentPosition = (HScroll1.Value)
End Sub

Private Sub HScroll1_Scroll()
On Error Resume Next
 WindowsMediaPlayer1.Controls.currentPosition = HScroll1.Value
End Sub

Private Sub Image1_Click()
Activ
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
c = True: X1 = X: Y1 = Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
n = (Me.Left + X - X1)
If c = True Then
    If n > 0 Then
        Call Me.Move(Me.Left + X - X1)
        For1_X = Me.Left
        If For6 = True Then Form6.Move (Me.Left + X - X1)
        DoEvents
    End If
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
c = False
End Sub
Private Sub GoMain()
On Error Resume Next '-------------------------------
LineColor True, Form6.Lin(4).BorderColor, FForm1
        FForm1.Show
        If F.Caption = "Video" Then
        FForm1.F.Caption = "Video"
        If Pla = True Then GoTo 2
        Form3.Show
        Form3.WindowsMediaPlayer1.URL = URL.Caption
        If Timer2.Enabled = True Then
        Form3.WindowsMediaPlayer1.Controls.currentPosition = WindowsMediaPlayer1.Controls.currentPosition
        Form3.WindowsMediaPlayer1.settings.Rate = 1.5
        End If
         Form3.WindowsMediaPlayer1.settings.Rate = 1: FForm1.URL.Caption = URL.Caption
        Else '-------------------------------
        FForm1.WindowsMediaPlayer1.settings.Rate = 1.5
If Timer2.Enabled = True Then FForm1.WindowsMediaPlayer1.settings.autoStart = True
        FForm1.WindowsMediaPlayer1.URL = Me.WindowsMediaPlayer1.URL
        FForm1.WindowsMediaPlayer1.Controls.currentPosition = WindowsMediaPlayer1.Controls.currentPosition
        FForm1.F.Caption = "For": FForm1.URL.Caption = Me.WindowsMediaPlayer1.URL
        FForm1.WindowsMediaPlayer1.settings.Rate = 1
        End If '-------------------------------
2       Unload Form6
        Unload Me
End Sub
Private Sub Lbl1_Click(Index As Integer)
On Error Resume Next
Dim B$
Select Case Index
Case 2: B = "All MediaFiles|*.M3u;*.mp3;*.wav;*.mid;*.avi;*.mpg;*.dat;*.vcd;*.svd;*.ifo;*.mov;*.wmv;*.wma;*.asf;*.mp2;*.m1v;*.swf|All Files|*.*"
        ShowOpen B, Form6, Me, WindowsMediaPlayer1, "Open Media", Pla
Case 3: Lbl1(13).BackColor = &HC0FFC0: Activ: frmBrowser.Show
Case 4: Form6.Visible = False: Lbl1_Click (12): Form6.cmd(12).Value = True
Case 5: Lbl1(15).BackColor = &HC0FFC0: FAbout.Show: Intcnt = 2: Me.Enabled = False
Case 6: Lbl1(16).BackColor = &HC0FFC0
        If F.Caption = "Video" Then '---------------Sho Video Screen---------------
        If Form3.WindowsMediaPlayer1.URL = WindowsMediaPlayer1.URL Then Exit Sub
        Form3.Show '-----------------
        Form3.WindowsMediaPlayer1.URL = WindowsMediaPlayer1.URL
        WindowsMediaPlayer1.Controls.stop
        Pla = True '--------------------
        Else
        MsgBox ("This Media File Not Valid VideoFomat")
        End If
Case 7: FForm1.URL.Caption = URL: FForm1.WindowsMediaPlayer1.URL _
= URL: FForm1.F.Caption = "For": Unload FForm1: End
Case 8: Form6.Hide: Me.WindowState = 1: FForm1.SerTxt.Text = "": FForm1.SerTxt.Text = "Min"
Case 9: Lbl1(9).BackColor = &HC0FFC0: Credit
Case 10: GoMain
Case 11: Lbl1(11).BackColor = &HC0FFC0
         Call lode(Form6)
         Form6.Show
         Form6.SSTab1.Tab = 2
Case 12: Lbl1(12).BackColor = &HC0FFC0
         Call lode(Form6)
         Form6.Show
         Form6.SSTab1.Tab = 1
Case 13: GoMain
Case 14: Lbl1(14).BackColor = &HC0FFC0
         Call lode(Form6)
         Form6.Show
         Form6.SSTab1.Tab = 3
Case 15: If DDE = False Then Shell App.Path + "\Copy.exe", vbNormalFocus
         For i = 0 To 120: FForm1.SerTxt.Text = i: Next
         FForm1.SerTxt.Text = ""
         FForm1.SerTxt.Text = "Convert"
Case 16: Lbl1(16).BackColor = &HC0FFC0
         If Avi = False Then Shell App.Path + "\Pic To Avi.exe", vbNormalFocus
End Select
Activ
End Sub

Private Sub Lbl1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 1:    Lbl1(Index).BackColor = &H80000007
Case 7:    Lbl1(Index).BackColor = &H8080FF
Case 8:    Lbl1(Index).BackColor = &HC0FFFF
Case 10:   Lbl1(Index).BackColor = &H404040
Case Else: Lbl1(Index).BackColor = &HFF00&
End Select
End Sub

Private Sub Timer1_Timer()
On Error Resume Next '-------------------------------
HScroll1.max = WindowsMediaPlayer1.Controls.currentItem.duration
Lbl1(1).Caption = WindowsMediaPlayer1.Controls.currentPositionString
HScroll1.Value = Int(WindowsMediaPlayer1.Controls.currentPosition)
        Ca = False
        If FForm1_WMP_Vol <> 0 Then _
        WindowsMediaPlayer1.settings.Volume = FForm1_WMP_Vol
        If Pla = True Then WindowsMediaPlayer1.Controls.stop
'------------------------Equlizer-----------------------------------------------
If Pla = True Then
Timer2.Enabled = False: Exit Sub: End If
                    X.Caption = WindowsMediaPlayer1.Controls.currentPosition
                    If Text1.Caption = "" Then
                    Text1.Caption = WindowsMediaPlayer1.Controls.currentPosition
                    X.Caption = "23"
                    End If
            If X.Caption = Text1.Caption Then
            Timer2.Enabled = False
            Else
            Timer2.Enabled = True
            Text1.Caption = X.Caption
            End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Pla = True Then Exit Sub
                    EqualizerForm Form6
DoEvents
End Sub

Private Sub WindowsMediaPlayer1_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
On Error Resume Next
   lR = SetTopMostWindow(Form1.hwnd, True)
End Sub
Private Sub WindowsMediaPlayer1_MouseMove(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
On Error Resume Next
    FForm1_WMP_Vol = WindowsMediaPlayer1.settings.Volume
End Sub
