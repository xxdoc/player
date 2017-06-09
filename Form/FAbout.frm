VERSION 5.00
Begin VB.Form FAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   75
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   5160
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   5040
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   8
      Top             =   240
      Width           =   855
      Begin VB.Image Image1 
         Height          =   855
         Left            =   0
         Picture         =   "FAbout.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   240
      Picture         =   "FAbout.frx":4D48
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Support Site"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "               OK"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   960
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   4800
      Width           =   4215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Home Page"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   1560
      LinkItem        =   "Ghayesh_Rayaneh@Yahoo.com"
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label13 
      Height          =   135
      Left            =   1920
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label7 
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6000
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      X1              =   6000
      X2              =   6000
      Y1              =   120
      Y2              =   5520
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   5295
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "For ShahidMofateh Librery "
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   5655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "For: XP ,98 ,NT ,Windows2000 ,Unix "
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CompanyName:GHAYESH RAYANEH"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Country :   Republic Islamic Of  IRAN"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   5655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BY:         NasserNiazyMobasser"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   5655
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2005-2006"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copyright:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   5535
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xo As Single, yo As Single, c As Boolean
Private Sub Form_Load()
Me.Left = (Screen.Width \ 2) - 3000: Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)
 ' Me.Width = 90
'Picture1.Picture = FForm1.ImageList1.ListImages(2).Picture
Tyf Picture3, "With_Black", Image2: Label1.Caption = "OK"
Label1.Caption = "OK"
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Label10.ForeColor = RGB(0, 0, 150)
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbLeftButton Then
Else: Label10.ForeColor = RGB(0, 255, 255)
End If
Label11.ForeColor = &H8000000D
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Label10.ForeColor = &HFF8080
Shell "explorer http://NasserVB.blogfa.com"
End Sub

Private Sub Label11_Click()
Label1.Caption = "NO"
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = RGB(0, 125, 125)
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbLeftButton Then
Else
Label11.ForeColor = RGB(0, 255, 255)
End If
Label10.ForeColor = &HFF8080
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = &H8000000D
End Sub

Private Sub Label12_Click()
On Error Resume Next
Shell "explorer mailto:nasservb@Gmail.com"

End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = RGB(0, 255, 255)
End Sub

Private Sub Label15_Click()
Shell "explorer http://tcvb.blogfa.com"
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.ForeColor = RGB(0, 0, 150)
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> vbLeftButton Then Label15.ForeColor = RGB(0, 255, 255)
End Sub

Private Sub Label15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.ForeColor = &HFF8080
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xo = X
yo = Y
c = True
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If c = True Then
    Me.Move Me.Left + X - xo, Me.Top + Y - yo
    DoEvents
    End If
    Label11.ForeColor = &H8000000D
Label12.ForeColor = &HFF8080
Label10.ForeColor = &HFF8080
Label15.ForeColor = &HFF8080
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
c = False
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &HFF8080
End Sub
Private Sub Timer1_Timer()
'On Error Resume Next
If Label1.Caption = "OK" Then OK
If Label1.Caption = "NO" Then No
If Label1.Caption = "End" Then EndAbout
 End Sub
 Sub OK()
    If (Me.Width < 6050) Then
    Me.Width = Me.Width + 200
    Call Me.Move(Me.Left - 100)
    Else: Label1.Caption = ""
    End If
 End Sub
Sub No()
    If (Me.Width > 200) Then
    Me.Width = Me.Width - 200
    Call Me.Move(Me.Left + 100)
    Else: Label1.Caption = "End"
    End If
End Sub

Sub EndAbout()
                   If Intcnt = 2 Then
                    Form1.Show: Form1.Enabled = True: Unload Me
                    ElseIf Intcnt = 1 Then
                     FForm1.Show: FForm1.Enabled = True: Unload Me
                    End If
End Sub
Private Sub Tyf(Frm As PictureBox, Color As String, Img As Image)
On Error GoTo B
Dim r%, F%, Heght%, Wath%, X%, i%
Heght = Frm.Height + 200: Wath = 300
F = Heght \ 255
Select Case Color
    Case "With_Black":  GoTo 4
End Select
Exit Sub '---------------------------Main--------------------------------------------
4 '--------------------------------------------------------------------------------
For i = 0 To Heght Step F
    r = r + 1
    If r = 20000 Then Exit For
        For X = i To F + i
           Frm.Line (0, X)-(Wath, X), RGB(255 - r, 255 - r, 255 - r)
        Next X
Next i: GoTo B
B:
Set Frm.Picture = Frm.Image
Img.Picture = Frm.Picture: Frm.Cls
End Sub

