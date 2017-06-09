VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8AB3C700-28FF-476A-852B-5E1558F44746}#1.0#0"; "command.ocx"
Begin VB.Form frmBrowser 
   ClientHeight    =   5400
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   8250
   Icon            =   "frmBrowser1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8250
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   8520
      Top             =   2160
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   8640
      Top             =   1680
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483635
      ForeColor       =   12582912
      TabCaption(0)   =   "Search MyComputer"
      TabPicture(0)   =   "frmBrowser1.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame(2)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search On Line"
      TabPicture(1)   =   "frmBrowser1.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "picAddress"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame(4)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame 
         Caption         =   "/\/\/\/\"
         Height          =   1815
         Index           =   4
         Left            =   5160
         TabIndex        =   26
         Top             =   -1680
         Width           =   3015
         Begin VB.FileListBox File1 
            Height          =   1065
            Left            =   1440
            Pattern         =   "*.htm"
            TabIndex        =   30
            Top             =   600
            Width           =   1455
         End
         Begin VB.DirListBox Dir1 
            Height          =   990
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1215
         End
         Begin VB.DriveListBox Drive2 
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000D&
            Caption         =   "\/\/\/\/"
            Height          =   255
            Left            =   1800
            TabIndex        =   27
            Top             =   1560
            Width           =   735
         End
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Caption         =   "Frame"
         Height          =   4935
         Index           =   3
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   8175
         Begin OsenXPCntrl.Command GoA 
            Height          =   375
            Left            =   7820
            TabIndex        =   25
            Top             =   470
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   ""
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
            MICON           =   "frmBrowser1.frx":047A
            PICN            =   "frmBrowser1.frx":0496
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.Command cmd 
            Height          =   405
            Index           =   10
            Left            =   1440
            TabIndex        =   22
            ToolTipText     =   "www.google.com"
            Top             =   45
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   ""
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
            MICON           =   "frmBrowser1.frx":1294
            PICN            =   "frmBrowser1.frx":12B0
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.Command cmd 
            Height          =   405
            Index           =   7
            Left            =   960
            TabIndex        =   21
            ToolTipText     =   "Refrash"
            Top             =   45
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   ""
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
            MICON           =   "frmBrowser1.frx":1A6E
            PICN            =   "frmBrowser1.frx":1A8A
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.Command cmd 
            Height          =   400
            Index           =   8
            Left            =   480
            TabIndex        =   20
            ToolTipText     =   "Foward"
            Top             =   50
            Width           =   470
            _ExtentX        =   820
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   ""
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
            MICON           =   "frmBrowser1.frx":1E4A
            PICN            =   "frmBrowser1.frx":1E66
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.Command cmd 
            Height          =   400
            Index           =   9
            Left            =   0
            TabIndex        =   19
            ToolTipText     =   "Back"
            Top             =   50
            Width           =   470
            _ExtentX        =   820
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   ""
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
            MICON           =   "frmBrowser1.frx":224B
            PICN            =   "frmBrowser1.frx":2267
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboAddress 
            Height          =   315
            Left            =   915
            TabIndex        =   13
            Text            =   "http://www.vbook.coo.ir"
            Top             =   480
            Width           =   6900
         End
         Begin SHDocVwCtl.WebBrowser brwWebBrowser 
            Height          =   4095
            Left            =   0
            TabIndex        =   14
            Top             =   840
            Width           =   8160
            ExtentX         =   14393
            ExtentY         =   7223
            ViewMode        =   1
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   0
            AutoArrange     =   -1  'True
            NoClientEdge    =   -1  'True
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
         Begin OsenXPCntrl.Command cmd 
            Height          =   405
            Index           =   5
            Left            =   4080
            TabIndex        =   15
            ToolTipText     =   "Home Site"
            Top             =   45
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "HomePage"
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
            MICON           =   "frmBrowser1.frx":2639
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.Command cmd 
            Height          =   405
            Index           =   4
            Left            =   3240
            TabIndex        =   16
            ToolTipText     =   "Support Site"
            Top             =   45
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "Support"
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
            MICON           =   "frmBrowser1.frx":2655
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.Command cmd 
            Height          =   405
            Index           =   3
            Left            =   2520
            TabIndex        =   23
            ToolTipText     =   "Help Page"
            Top             =   45
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "Help"
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
            MICON           =   "frmBrowser1.frx":2671
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblAddress 
            Caption         =   "&Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Tag             =   "&Address:"
            Top             =   540
            Width           =   3075
         End
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Caption         =   "Frame"
         Height          =   5055
         Index           =   2
         Left            =   -75000
         TabIndex        =   1
         Top             =   360
         Width           =   8175
         Begin VB.ListBox Lvfiles 
            Height          =   4740
            Left            =   2160
            TabIndex        =   24
            Top             =   50
            Width           =   5895
         End
         Begin VB.Frame Frame 
            Caption         =   "Search Info"
            Height          =   1215
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   1935
            Begin VB.DriveListBox Drive1 
               Height          =   315
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Width           =   1075
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   1080
               TabIndex        =   4
               Text            =   "mp3"
               Top             =   810
               Width           =   615
            End
            Begin OsenXPCntrl.Command cmd 
               Height          =   330
               Index           =   6
               Left            =   1200
               TabIndex        =   3
               Top             =   360
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   582
               BTYPE           =   3
               TX              =   "Browse"
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
               MICON           =   "frmBrowser1.frx":268D
               UMCOL           =   -1  'True
               SOFT            =   -1  'True
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Left            =   120
               TabIndex        =   5
               Text            =   "d:\"
               Top             =   360
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "FileType:*."
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   800
               Width           =   855
            End
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   480
            Top             =   3600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "Save To Playlist File"
            Filter          =   "All Playlist[*.m3u]|*.M3u"
         End
         Begin OsenXPCntrl.Command cmd 
            Height          =   400
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   " Start Search"
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
            MICON           =   "frmBrowser1.frx":26A9
            PICN            =   "frmBrowser1.frx":26C5
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   1
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.Command cmd 
            Height          =   405
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   2400
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "Save Playlist"
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
            MICON           =   "frmBrowser1.frx":2B5C
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label2 
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1850
            Width           =   1935
         End
         Begin VB.Label Label3 
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2100
            Width           =   1935
         End
      End
      Begin VB.PictureBox picAddress 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   7620
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   960
         Width           =   7620
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mremov 
         Caption         =   "Remov"
      End
      Begin VB.Menu MUp 
         Caption         =   "Up"
      End
      Begin VB.Menu MDown 
         Caption         =   "Down"
      End
      Begin VB.Menu Menu_Pro 
         Caption         =   "Propertis"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 1
    On Error GoTo Hi
    Dim SearchPath As String, FindStr As String
    Dim FileSize As Long
    Dim NumFiles As Integer, NumDirs As Integer
    Dim Magued
    Magued = Text2.Text         '------------File Type--------
    Lvfiles.Clear
    Me.Caption = "Files Searching..."
    SearchPath = Text1.Text     '----------File Path
    FindStr = "*." + Text2.Text '------------File Type--------
    cmd(1).Enabled = False: Frame(2).Enabled = False
    FileSize = FindFiles(SearchPath, FindStr, NumFiles, NumDirs)
    Label2 = NumFiles & " Files Is Founded!"
    Label3.Caption = Format((FileSize \ 1024), "###,###") & " KB"
    Screen.MousePointer = vbDefault: cmd(1).Enabled = True: Frame(2).Enabled = True: Me.Caption = "Search Copmplted!"
Hi:
    Case 6
            Dim lpIDList As Long
            Dim sBuffer As String
            Dim szTitle As String
            Dim tBrowseInfo As BrowseInfo
            szTitle = "Search In"
            With tBrowseInfo
            .hWndOwner = Me.hwnd
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
            End With
            lpIDList = SHBrowseForFolder(tBrowseInfo)
            If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            Text1.Text = sBuffer: Label3.Caption = "Search Path =" + sBuffer
            End If
    Case 2: CommonDialog1.DialogTitle = "Save List"
            CommonDialog1.Filter = "All Playlist[*.m3u]|*.m3u|All files[*.*]|*.*": CommonDialog1.ShowSave
            If CommonDialog1.filename <> "" Then Savelis CommonDialog1.filename
    Case 3: frmBrowser.brwWebBrowser.Navigate App.Path + "\Help\z.html"
    Case 8: frmBrowser.brwWebBrowser.GoForward
    Case 9: frmBrowser.brwWebBrowser.GoBack
    Case 7: frmBrowser.brwWebBrowser.Refreshd
    Case 5: frmBrowser.brwWebBrowser.Navigate "http://NaserVb.Blogfa.com"
    Case 4: frmBrowser.brwWebBrowser.Navigate "http://TcVb.Blogfa.com"
    Case 10: frmBrowser.brwWebBrowser.Navigate "http://google.com"
End Select
If Index > 6 Then timTimer.Enabled = True

End Sub


Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dir1.ToolTipText = Dir1.Path
End Sub

Private Sub Drive1_Change()
Text1.Text = Left(Drive1.Drive, 2) + "\"
End Sub

Private Sub Drive2_Change()
On Error Resume Next
Dir1.Path = Drive2.Drive
End Sub

Private Sub File1_Click()
Dim PAdres As String
PAdres = Dir1.Path & "\" & File1.filename
If Len(Dir1.Path) = 3 Then PAdres = Dir1.Path + File1.filename
        cboAddress.Text = PAdres
        cboAddress.AddItem cboAddress.Text
     brwWebBrowser.Navigate (PAdres)
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
File1.ToolTipText = File1.filename
End Sub

Private Sub Form_Load()
On Error Resume Next
     SetingBrw Me
Drive1_Change
    Form_Resize
    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True '--------------------------------------------
        brwWebBrowser.Navigate StartingAddress
    End If
     brwWebBrowser.Navigate (App.Path + "\Help\z.htm")
Exit Sub
End Sub

Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Dim i As Integer '----------------------------------------------------------
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If '------------------------------------------------------------------
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If '------------------------------------------------------------------------
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    cboAddress.Width = Me.ScaleWidth - (880 + 375)
    GoA.Left = Me.Width - 500
  If Me.Width <= 5250 Then '-----------------------------------------------------
  Me.Enabled = False: Me.Width = 5250: End If
  If Me.Height <= 3700 Then
  Me.Enabled = False: Me.Height = 3700: End If
    brwWebBrowser.Width = Me.Width - 150: SSTab1.Width = Me.Width - 100
    SSTab1.Height = Me.Height - 100: picAddress.Width = SSTab1.Width - 50
    brwWebBrowser.Height = Me.Height - 1750
Lvfiles.Width = Me.Width - 2355: Lvfiles.Height = Me.Height - 975
Frame(2).Height = Me.Height: Frame(2).Width = Me.Width
Frame(3).Height = Me.Height: Frame(3).Width = Me.Width
End Sub


Private Sub Frame_Click(Index As Integer)
If Index = 3 Then
brwWebBrowser.Offline = Not (brwWebBrowser.Offline)
Me.Caption = "Work OffLine:= " + CStr(brwWebBrowser.Offline)
End If
End Sub

Private Sub GoA_Click()
brwWebBrowser.Navigate (cboAddress.Text)
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame(4).Top = 0
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame(4).Top = -1680
End Sub

Private Sub lvFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Lvfiles.ListCount > 0) And (Button = 2) Then PopupMenu Menu
End Sub

Private Sub Lvfiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Lvfiles.ToolTipText = Lvfiles.Text
End Sub

Private Sub MAdd_Click()
With CommonDialog1
     .DialogTitle = "Add Media"
     .Filter = Text2.Text + "|*." + Text2.Text + "|All Files|*.*"
     .ShowOpen
     If .filename <> "" Then Lvfiles.AddItem .filename
End With
End Sub

Private Sub MDown_Click()
LChange Me, Lvfiles, False
End Sub

Private Sub Menu_Pro_Click()
Dim r As Long
    r = ShowFileProperties(Lvfiles.Text, Me.hwnd)
 
End Sub


Private Sub mremov_Click()
Lvfiles.RemoveItem Lvfiles.ListIndex
End Sub

Private Sub MUp_Click()
LChange Me, Lvfiles, True
End Sub

Private Sub Timer1_Timer()
Me.Enabled = True
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else '--------------------------------------
        Me.Caption = "Working..."
    End If
End Sub


