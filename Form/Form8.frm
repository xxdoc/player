VERSION 5.00
Object = "{8AB3C700-28FF-476A-852B-5E1558F44746}#1.0#0"; "Command.ocx"
Begin VB.Form Form8 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Directory"
   ClientHeight    =   3795
   ClientLeft      =   7815
   ClientTop       =   1425
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin OsenXPCntrl.Command Command1 
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   3380
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "&Crate Folder"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "Form8.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.Command Command3 
      Height          =   300
      Left            =   840
      TabIndex        =   5
      Top             =   3380
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "Form8.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.Command Command2 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   3380
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "Form8.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000008&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.DriveListBox Drive2 
         BackColor       =   &H00404040&
         ForeColor       =   &H80000018&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.DirListBox Dir2 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFC0&
         Height          =   1665
         Left            =   120
         TabIndex        =   1
         Top             =   650
         Width           =   2415
      End
      Begin VB.Label Text1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ASd As New FileSystemObject
Private Sub Command1_Click()
On Error Resume Next
Dim n As String '-------------------------------------------
    n = InputBox("Plase Insert The FolderName", "SelectFolderName", "NewFolder")
If ASd.FolderExists(Text1.Caption + "\" + n) = True Then
    MsgBox ("The Folder Is Exises")
    Exit Sub '----------------------------------------------
ElseIf n = Empty Then
    n = "NewFolder"
    GoTo 2
Else '--------------------------------------------------------
2   ASd.CreateFolder (Text1.Caption + "\" + n): Dir2.Refresh: Dir2.Path = Dir2.Path + "\" + n
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next '------------------------------------
If Len(Text1.Caption) = 3 Then
         Form7.Text5.Text = Text1.Caption
Else:    Form7.Text5.Text = Text1.Caption + "\"
End If '---------------------------------------------------
         Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Dir2_Change()
Text1.Caption = Dir2.Path
End Sub

Private Sub Dir2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dir2.ToolTipText = Dir2.Path
End Sub

Private Sub Drive2_Change()
On Error Resume Next
Dir2.Path = Drive2.Drive: Text1.Caption = Dir2.Path
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Caption = Dir2.Path
Command1.Refresh:: Command2.Refresh: Command3.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Form7.Enabled = True
End Sub
Private Sub NasserNiazyMobasser_Emza(Emza As String)
'
'
'                                  _________________
'                                s                   s
'                             s                         s
'                          s                              s
'                        s                                  s
'                      s               s s                    s
'                    s                s   s                     s
'                  s                 s     s                      s
'                s                   s                             s
'              s   __________________s___________________           s
'             s    s                 s                 s            s
'            s      s                s               s             s
'           s         s              s             s              s
'          s             s           s           s              s
'         s                s         s         s              s
'         s                  s       s       s              s
'         s                    s     s     s              s
'         s                      s   s   s              s
'          s                     ________________     s
'           s                   s                s  s
'            s                s                  ss
'              s            s                  s
'               s  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\  /\
'                \/  \/  \/  \/  \/  \/  \/  \/  \/  \/  \/  \/  \/
'***************************************************************************************
'*       ggggg    hhhh  hhhh      aaa      yy    yy   eeeeee      sss    hhhh  hhhh
'*      g          hh    hh       a a        y  y     e          s        hh    hh
'*     gg          hhhhhhhh      aaaaa        yy      eeee        ss      hhhhhhhh
'*     gg   gg     hh    hh     a     a       yy      e             s     hh    hh
'*      gggggg    hhhh  hhhh  aaa     aaa    yyyy     eeeeee     sss     hhhh  hhhh
'***************************************************************************************

End Sub

