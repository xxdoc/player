VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Show Video"
   ClientHeight    =   6480
   ClientLeft      =   6885
   ClientTop       =   1740
   ClientWidth     =   8550
   Icon            =   "Form3.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   ScaleHeight     =   6480
   ScaleWidth      =   8550
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   8040
      Top             =   5760
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   14208
      _cy             =   11245
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Dim DfA As String

Private Sub Form_Load()
On Error Resume Next
Pla = True
Me.Left = Screen.Width \ 2 - (Me.Width \ 2): Me.Top = Screen.Height \ 2 - (Me.Height \ 2)
End Sub

Private Sub Form_Resize()
On Error Resume Next '--------------------------------------------
        If Me.WindowState <> 1 Then
        If Me.Width < 3240 Then
        Me.Width = 3240: Me.Enabled = False
        End If '----------------------------------------------------
        If Me.Height < 2625 Then
        Me.Height = 2625: Me.Enabled = False
        End If
        Me.WindowsMediaPlayer1.Left = 0: Me.WindowsMediaPlayer1.Top = 0
        Me.WindowsMediaPlayer1.Width = (Me.Width - 100): Me.WindowsMediaPlayer1.Height = (Me.Height - 450)
        Pla = True
        End If '--------------------------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next '------------------------------------
FForm1.F.Caption = "For"
If FForm1.Visible = False Then
    Form1.URL.Caption = WindowsMediaPlayer1.URL
    Form1.WindowsMediaPlayer1.URL = WindowsMediaPlayer1.URL
End If
Pla = False
End Sub

Private Sub Timer1_Timer()
Me.Enabled = True
End Sub

