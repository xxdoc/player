VERSION 5.00
Begin VB.Form LOADE 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   2745
   ClientLeft      =   3015
   ClientTop       =   645
   ClientWidth     =   4155
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   3000
      ScaleHeight     =   2715
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1080
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading File ..."
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Wellcom To Player 2.4"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ms.dll"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Text2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "||||||||||||||||||||||||||||||||||||||"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2775
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "LOADE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                Dim DAT1 As ClientRecord
                Dim i As Integer
Private Sub Form_Load()
On Error GoTo AD '--------------------------------------------------------
                ProsesFiles
Dim SD As Integer, A As String, s As String, d As String, F As String
                Open App.Path + "\NMS\DAT1.DLL" For Random Access Read As 1
'                Do Until EOF(1)
                Get 1, , DAT1
'                Loop
                Close 1
Select Case DAT1.Sink
Case "FTAB" '----------------------------------------
            Call SinkP(FForm1): FForm1.Show: Unload Me
Case "FORM" '----------------------------------------
             Call SinkS(FForm1): FForm1.Show: Unload Me
Case "MinS": '---------------------------------------
            Call SinkS(FForm1): FForm1.Show: FormMinSize FForm1: Unload Me
Case "MinP": '---------------------------------------
             Call SinkP(FForm1): FForm1.Show: MinSize FForm1: Unload Me
Case Else '-------------------------------
12          FLoad.Show
            FLoad.Command1.Enabled = False
            Unload Me
End Select '-------------------------------
        Exit Sub
AD:             If Err.Number = 76 Then
                Call KTM
                Else
                GoTo 12 '-----------------------
                End If
End Sub
Private Sub KTM()
Tyf Picture1, "Green_Black", Image1
            Call ASD.CreateFolder(App.Path + "\NMS") 'Crating Folder For Archive
            Open App.Path + "\NMS\User.dll" For Output As 3
            Close 3
            Open App.Path + "\NMS\DAT1.DLL" For Output As 1
            Open App.Path + "\NMS\DAT2.M3U" For Output As 2
                         Write #1, Text1.Text: Write #2, Text1.Text
            Close 1: Close 2 '------------------------------------
            Label2.Visible = True
            Text2.Visible = True
            Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
                Dim strans As String
                Text2.Caption = (Text2.Caption) + "|"
                Label2.Caption = App.Path + "\NMS\DAT" + Str(Int(Rnd() * 100)) + ".DLL "
                If Len(Text2.Caption) > 107 Then
                 FLoad.Show: FLoad.Text1.Text = "A": Unload Me
                End If '--------------------------------------
End Sub
'///////////////////////////////////////////////////////////////////////////////////
Private Sub ProsesFiles()
                For i = 67 To 90
                If ASD.FolderExists(Chr(i) + ":\Windows") = True Then
                If ASD.FolderExists(Chr(i) + ":\Windows\System32") = True Then
                Path = Chr(i) + ":\Windows\System32\": Call Proses(Path): Exit Sub
                End If
                End If
                Next
End Sub
