VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15090
   Icon            =   "all.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   15090
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   120
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
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   4035
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   3615
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         ForeColor       =   &H80000013&
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   150
         Width           =   495
      End
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   495
         Left            =   0
         TabIndex        =   23
         Top             =   0
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
         enableContextMenu=   -1  'True
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
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000013&
      Height          =   495
      Left            =   6240
      TabIndex        =   18
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
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H008080FF&
         Height          =   255
         Left            =   480
         TabIndex        =   20
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
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         Caption         =   "\/"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   255
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9600
      Top             =   120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   840
      MouseIcon       =   "all.frx":28F8
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   720
      Width           =   9975
      Begin VB.Label Label19 
         BackColor       =   &H000000FF&
         Caption         =   ">"
         ForeColor       =   &H8000000E&
         Height          =   220
         Left            =   9720
         MousePointer    =   1  'Arrow
         TabIndex        =   17
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label18 
         BackColor       =   &H000000FF&
         Caption         =   "<"
         ForeColor       =   &H8000000E&
         Height          =   220
         Left            =   9000
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Graphic"
         Height          =   220
         Left            =   9120
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ManyCopy"
         Height          =   220
         Left            =   7200
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Convert"
         Height          =   220
         Left            =   6480
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Setting"
         Height          =   220
         Left            =   5760
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HideTaskbar"
         Height          =   220
         Left            =   4680
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Directory"
         Height          =   220
         Left            =   3840
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ViewList"
         Height          =   220
         Left            =   8160
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Menu"
         Height          =   220
         Left            =   3240
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Video"
         Height          =   220
         Left            =   2640
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "About"
         Height          =   220
         Left            =   2040
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "PlayList"
         Height          =   225
         Left            =   1320
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Sink"
         Height          =   220
         Left            =   720
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Open"
         Height          =   220
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   120
         Width           =   450
      End
   End
   Begin VB.Label F 
      Height          =   135
      Left            =   9480
      TabIndex        =   1
      Top             =   360
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
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal X As Long, _
ByVal Y As Long, _
ByVal cx As Long, _
ByVal cy As Long, _
ByVal wFlags As Long) As Long
Dim c As Boolean, X1, Y1 As Single, n As Single, g As Boolean, g1 As Single, g2 As Single
'
'
Private Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
As Long
          If Topmost = True Then 'Make the window topmost
             SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
         Else
              SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
             SetTopMostWindow = False
         End If
End Function


Private Sub Form_Activate()
     Dim lR As Long
   lR = SetTopMostWindow(Form1.hwnd, True)

End Sub

Private Sub Form_Click()
     Dim lR As Long
   lR = SetTopMostWindow(Form1.hwnd, True)

End Sub

Private Sub Form_GotFocus()
     Dim lR As Long
   lR = SetTopMostWindow(Form1.hwnd, True)

End Sub

Private Sub Form_Load()
    Dim lR As Long
    lR = SetTopMostWindow(Form1.hwnd, True)
    Me.Left = 5000: Me.Top = 0
    c = False
Me.Height = 420: Frame1.Top = 50
Me.Width = 7230: Frame1.Left = 3600
'Frame1.MouseIcon = ImageList1.ListImages(1).Picture
'For inactive always on top Properties : lR = SetTopMostWindow(Form1.hwnd, False)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0FFC0

End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
g = True
g1 = X
Frame1.MouseIcon = ImageList1.ListImages(2).Picture
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim d As Single
d = Frame1.Left + X - g1
If d < 3600 And d > -3670 Then
If g = True Then Call Frame1.Move(Frame1.Left + X - g1)
End If
Label1.BackColor = &H80000007: Label3.BackColor = &HC0FFC0: Label5.BackColor = &HC0FFC0: Label7.BackColor = &H8080FF: Label9.BackColor = &HC0FFC0: Label11.BackColor = &HC0FFC0: Label12.BackColor = &HC0FFC0: Label15.BackColor = &HC0FFC0: Label17.BackColor = &HC0FFC0
Label2.BackColor = &HC0FFC0: Label4.BackColor = &HC0FFC0: Label6.BackColor = &HC0FFC0: Label8.BackColor = &HC0FFFF: Label10.BackColor = &H404040: Label13.BackColor = &HC0FFC0: Label14.BackColor = &HC0FFC0: Label16.BackColor = &HC0FFC0: Label18.BackColor = &HFF&: Label19.BackColor = &HFF&
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
g = False
Frame1.MouseIcon = ImageList1.ListImages(1).Picture

End Sub

Private Sub HScroll1_Scroll()
 WindowsMediaPlayer1.Controls.CurrentPosition = HScroll1.Value

End Sub

Private Sub Image1_Click()
     Dim lR As Long
   lR = SetTopMostWindow(Form1.hwnd, True)

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
c = True
X1 = X: Y1 = Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
n = (Me.Left + X - X1)
If c = True Then
If n > 0 Then
1 Call Me.Move(Me.Left + X - X1)
End If
End If
Label7.BackColor = &H8080FF
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
c = False
End Sub

Private Sub Label1_Click()
     Dim lR As Long
   lR = SetTopMostWindow(Form1.hwnd, True)

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0FFC0

End Sub

Private Sub Label10_Click()
Unload Me
FForm1.Show
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.BackColor = vbBlack
End Sub

Private Sub Label11_Click()
Label11.BackColor = &HC0FFC0

End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.BackColor = &HFF00&

End Sub

Private Sub Label12_Click()
Label12.BackColor = &HC0FFC0

End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.BackColor = &HFF00&

End Sub

Private Sub Label13_Click()
Label13.BackColor = &HC0FFC0

End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.BackColor = &HFF00&

End Sub

Private Sub Label14_Click()
Label14.BackColor = &HC0FFC0

End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.BackColor = &HFF00&

End Sub

Private Sub Label15_Click()
Label15.BackColor = &HC0FFC0

End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.BackColor = &HFF00&

End Sub

Private Sub Label16_Click()
Label16.BackColor = &HC0FFC0

End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.BackColor = &HFF00&

End Sub


Private Sub Label2_Click()
Label2.BackColor = &HC0FFC0
     Dim lR As Long, ASD As String
   lR = SetTopMostWindow(Form1.hwnd, True)
                    CommonDialog1.DialogTitle = "Open Media"
                    CommonDialog1.Filter = "All MediaFiles|*.M3u;*.mp3;*.wav;*.mid;*.avi;*.mpg;*.dat;*.vcd;*.svd;*.ifo;*.mov;*.wmv;*.asf;*.mp2;*.m1v;*.swf|All Files|*.*"
                    CommonDialog1.ShowOpen
                    If CommonDialog1.FileName <> "" Then
If WindowsMediaPlayer1.URL <> CommonDialog1.FileName Then
WindowsMediaPlayer1.settings.AutoStart = True
WindowsMediaPlayer1.URL = CommonDialog1.FileName
HScroll1.max = WindowsMediaPlayer1.Controls.currentItem.Duration
Select Case UCase(Right(CommonDialog1.FileName, 3))
                    Case "MPG": ASD = "OK"
                    Case "AVI": ASD = "OK"
                    Case "DAT": ASD = "OK"
                    Case "VCD": ASD = "OK"
                    Case "SVD": ASD = "OK"
                    Case "IFO": ASD = "OK"
                    Case "MOV": ASD = "OK"
                    Case "WMV": ASD = "OK"
                    Case "ASF": ASD = "OK"
                    Case "MP2": ASD = "OK"
                    Case "MV1": ASD = "OK"
                    End Select
                    If ASD = "" Then
                    F.Caption = "For"
                    Exit Sub
                    Else
                    F.Caption = "Video"
                    End If
End If
End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0FFC0
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HFF00&
End Sub

Private Sub Label3_Click()
Label3.BackColor = &HC0FFC0
Dim lR As Long
   lR = SetTopMostWindow(Form1.hwnd, True)
FLoad.Show
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = &HFF00&
End Sub

Private Sub Label4_Click()
Label4.BackColor = &HC0FFC0
Dim I As Integer
'For i = 0 To FForm1.List1.ListCount - 1
'Next i
         Open App.Path + "\NMS\DAT2.M3U" For Output As #1
         Print 1, "#EXTM3U:"
         For I = 0 To FForm1.List1.ListCount - 1
         FForm1.List1.Selected(I) = True
         Print 1, "#EXTNIF:"
         Print 1, FForm1.List1.Text
         If FForm1.List1.Text = "" Then Exit For
         Next I
         Close 1
WindowsMediaPlayer1.URL = App.Path + "\nms\dat2.m3u"
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackColor = &HFF00&

End Sub

Private Sub Label5_Click()
Label5.BackColor = &HC0FFC0
FAbout.Show
Pla = "About"
Me.Enabled = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = &HFF00&

End Sub

Private Sub Label6_Click()
Label6.BackColor = &HC0FFC0
If F.Caption = "Video" Then
Form3.Show
Form3.WindowsMediaPlayer1.URL = WindowsMediaPlayer1.URL
WindowsMediaPlayer1.Controls.Stop
Pla = "Play"
Else
MsgBox ("This Audio File Not Valid VideoFomat")
End If
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackColor = &HFF00&

End Sub

Private Sub Label7_Click()
FForm1.WindowsMediaPlayer1.URL = WindowsMediaPlayer1.URL
Unload FForm1
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackColor = &HFF&: Label8.BackColor = &HC0FFFF
End Sub

Private Sub Label8_Click()
Me.WindowState = 1
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackColor = &HFFFF&: Label7.BackColor = &H8080FF: Label10.BackColor = &H404040
End Sub

Private Sub Label9_Click()
Label9.BackColor = &HC0FFC0

End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BackColor = &HFF00&

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
HScroll1.max = WindowsMediaPlayer1.Controls.currentItem.Duration
Label1.Caption = WindowsMediaPlayer1.Controls.currentPositionString
HScroll1.Value = Int(WindowsMediaPlayer1.Controls.CurrentPosition)
WindowsMediaPlayer1.settings.Rate = 1
If Pla = "Play" Then WindowsMediaPlayer1.Controls.Stop
End Sub


Private Sub WindowsMediaPlayer1_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
     Dim lR As Long
   lR = SetTopMostWindow(Form1.hwnd, True)

End Sub

'Private Sub WindowsMediaPlayer1_MouseMove(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
'     Dim lR As Long
'   lR = SetTopMostWindow(Form1.hwnd, True)
'
'End Sub
