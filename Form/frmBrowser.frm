VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8AB3C700-28FF-476A-852B-5E1558F44746}#1.0#0"; "Command.ocx"
Begin VB.Form frmBrowser 
   ClientHeight    =   5355
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   8250
   Icon            =   "frmBrowser.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
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
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   8760
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":12AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483635
      ForeColor       =   12582912
      TabCaption(0)   =   "Search MyComputer"
      TabPicture(0)   =   "frmBrowser.frx":158E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame(2)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search On Line"
      TabPicture(1)   =   "frmBrowser.frx":15AA
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "picAddress"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Caption         =   "Frame"
         Height          =   5055
         Index           =   2
         Left            =   -75000
         TabIndex        =   10
         Top             =   360
         Width           =   8175
         Begin VB.Frame Frame 
            Caption         =   "Search Directory"
            Height          =   1215
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   1935
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   1080
               TabIndex        =   18
               Text            =   "mp3"
               Top             =   810
               Width           =   615
            End
            Begin VB.TextBox Text1 
               Height          =   315
               Left            =   120
               TabIndex        =   5
               Text            =   "d:\"
               Top             =   360
               Width           =   1400
            End
            Begin VB.DriveListBox Drive1 
               Height          =   315
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label1 
               Caption         =   "FileType:*."
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   800
               Width           =   855
            End
         End
         Begin VB.DirListBox Dir1 
            Height          =   990
            Left            =   120
            TabIndex        =   13
            Top             =   3840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chk_sub_dir 
            Caption         =   "Search sub directories"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2520
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   1320
            Top             =   4080
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "Save To Playlist File"
            Filter          =   "All Playlist[*.m3u]|*.M3u"
         End
         Begin ComctlLib.ListView lvFiles 
            Height          =   4650
            Left            =   2160
            TabIndex        =   11
            ToolTipText     =   "Hold Ctrl to select or unselect file or for multi select hold Shift key"
            Top             =   0
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   8202
            View            =   3
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Files"
               Object.Width           =   14896
            EndProperty
         End
         Begin OsenXPCntrl.Command Cmd 
            Height          =   465
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   45
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   820
            BTYPE           =   3
            TX              =   "Start Search"
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
            MICON           =   "frmBrowser.frx":15C6
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.Command Cmd 
            Height          =   465
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   1920
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   820
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
            MICON           =   "frmBrowser.frx":15E2
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label F 
            Caption         =   "Label1"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   3000
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Caption         =   "Frame"
         Height          =   4935
         Index           =   3
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   8175
         Begin OsenXPCntrl.Command Cmd 
            Height          =   405
            Index           =   0
            Left            =   4080
            TabIndex        =   21
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
            MICON           =   "frmBrowser.frx":15FE
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.Command Cmd 
            Height          =   405
            Index           =   4
            Left            =   3360
            TabIndex        =   20
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
            MICON           =   "frmBrowser.frx":161A
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
            TabIndex        =   7
            Text            =   "http://www.GHAYESH RAYANEH.com"
            Top             =   480
            Width           =   6675
         End
         Begin MSComctlLib.Toolbar tbToolBar 
            Height          =   510
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   900
            ButtonWidth     =   820
            ButtonHeight    =   794
            ImageList       =   "imlIcons"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   6
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Back"
                  Object.ToolTipText     =   "Back"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Forward"
                  Object.ToolTipText     =   "Forward"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Stop"
                  Object.ToolTipText     =   "Stop"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Refresh"
                  Object.ToolTipText     =   "Refresh"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Home"
                  Object.ToolTipText     =   "Home"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Search"
                  Object.ToolTipText     =   "Search"
                  ImageIndex      =   6
               EndProperty
            EndProperty
            Begin OsenXPCntrl.Command Cmd 
               Height          =   405
               Index           =   3
               Left            =   2775
               TabIndex        =   17
               Top             =   45
               Width           =   600
               _ExtentX        =   1058
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
               MICON           =   "frmBrowser.frx":1636
               UMCOL           =   -1  'True
               SOFT            =   -1  'True
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
         End
         Begin SHDocVwCtl.WebBrowser brwWebBrowser 
            Height          =   4095
            Left            =   0
            TabIndex        =   9
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
         Begin VB.Label lblAddress 
            Caption         =   "&Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Tag             =   "&Address:"
            Top             =   540
            Width           =   3075
         End
      End
      Begin VB.PictureBox picAddress 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   0
         ScaleHeight     =   435
         ScaleWidth      =   7620
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1020
         Width           =   7620
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '-------Const For Search-----------------------
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const MAX_PATH = 260
'----------------Typeing For Search------------------------------
Private Type FILETIME
        dwLowDateTime             As Long
        dwHighDateTime            As Long
End Type '-------------------------------------
Private Type WIN32_FIND_DATA
        ftCreationTime            As FILETIME
        ftLastAccessTime          As FILETIME
        ftLastWriteTime           As FILETIME
        dwFileAttributes          As Long
        nFileSizeHigh             As Long
        nFileSizeLow              As Long
        dwReserved0               As Long
        dwReserved1               As Long
        cFileName                 As String * MAX_PATH
        cAlternate                As String * 14
End Type
'-----------------------Declareing API------------------------------------------
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'------------------------------------------------------------------------------
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'------------------------------------------------------------------------------
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
   (ByVal lpLibFileName As String) As Long
'------------------------------------------------------------------------------
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
    (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
'------------------------------------------------------------------------------
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long
'------------------------------------------------------------------------------
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, _
   ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, _
   ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
'------------------------------------------------------------------------------
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
'------------------------------------------------------------------------------
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, _
    lpExitCode As Long) As Long
'------------------------------------------------------------------------------
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private pbMessage As Boolean
'------------------------------------------------------------------------------
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'------------------------------------------------------------------------------
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'------------------------------------------------------------------------------
Dim Reg As String, Success As Long
Dim file_obj As New FileSystemObject
Dim selected_boll As Boolean
Dim mresult '--------------------------------------------------------
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Private Sub Savelis(OutPath As String)
On Error Resume Next '--------------------------------------------------
                    Dim T3 As String, T2, strans As String, L As Single, i As Integer
                    T3 = "": T2 = ""
                    If lvFiles.ListItems(1) = "" Then
                    strans = MsgBox("File Not Found!", vbCritical)
                    Exit Sub '------------------------------------------------------
                    End If
                    If UCase(Right(OutPath, 3)) <> "M3U" Then Exit Sub
            Open OutPath For Output As #1
                    Print #1, "#EXTM3U:"
                For i = 1 To lvFiles.ListItems.Count '----------------------------
                    Print #1, "#EXTNIF:"
                    Print #1, lvFiles.ListItems(i)
                Next i '------------------------------------------------------
            Close #1
End Sub


Private Sub cmd_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 1:         Dim i As Integer
                Dir1.Path = Text1.Text
                Search
Case 2: CommonDialog1.ShowSave '-----------------------------------------------
        If CommonDialog1.filename = "" Then Exit Sub
        Savelis CommonDialog1.filename
Case 3: Me.Caption = "http://www.GHAYESH RAYANEH.com"
        brwWebBrowser.Navigate (App.Path + "\Help\z.htm")
End Select
End Sub

Private Sub Drive1_Change()
Text1.Text = Drive1.Drive + "\": Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Drive1_Change
    Me.Show '----------------------------------------------------------------
    tbToolBar.Refresh
    Form_Resize
    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True '--------------------------------------------
        brwWebBrowser.Navigate StartingAddress
    End If
     brwWebBrowser.Navigate (App.Path + "\Help\z.htm")
     Setng Me, 3, False, FForm1.WindowsMediaPlayer1, FForm1.List1, 3
     lvFiles.BackColor = Setting.ProgramColor: Drive1.BackColor = Setting.ProgramColor
     Text1.BackColor = Setting.ProgramColor: Label1.BackColor = Setting.ProgramColor
     Text2.BackColor = Setting.ProgramColor: cboAddress.BackColor = Setting.ProgramColor
     lblAddress.BackColor = Setting.ProgramColor: SSTab1.BackColor = Setting.ProgramColor
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
    cboAddress.Width = Me.ScaleWidth - 850
  If Me.Width <= 5250 Then '-----------------------------------------------------
  Me.Enabled = False: Me.Width = 5250: End If
  If Me.Height <= 3700 Then
  Me.Enabled = False: Me.Height = 3700: End If
    brwWebBrowser.Width = Me.Width - 150: SSTab1.Width = Me.Width - 100
    SSTab1.Height = Me.Height - 100: picAddress.Width = SSTab1.Width - 50
    brwWebBrowser.Height = Me.Height - 1750
lvFiles.Width = Me.Width - 2355: lvFiles.Height = Me.Height - 975
Frame(2).Height = Me.Height: Frame(2).Width = Me.Width
Frame(3).Height = Me.Height: Frame(3).Width = Me.Width
End Sub


Private Sub Timer1_Timer()
Me.Enabled = True
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else '--------------------------------------
        Me.Caption = "Working..."
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal button As button)
    On Error Resume Next
    timTimer.Enabled = True
    Select Case button.Key
        Case "Back":    brwWebBrowser.GoBack
        Case "Forward": brwWebBrowser.GoForward
        Case "Refresh": brwWebBrowser.Refresh
        Case "Home":    brwWebBrowser.GoHome
        Case "Search":  brwWebBrowser.GoSearch
        Case "Stop":    timTimer.Enabled = False
                        brwWebBrowser.stop
                        Me.Caption = brwWebBrowser.LocationName
    End Select
End Sub
Private Sub Search()
Dim i As Integer
  Dim ar As Variant
  Dim k As Integer '----------------------------------------
  Dim temp As String
lvFiles.ListItems.Clear

  ar = Split("*." + Text2.Text, ";", , vbTextCompare)
  For k = 0 To UBound(ar)
    temp = ar(k) '------------------------------------------
    If chk_sub_dir.Value = 1 Then
      Call getfiles(Dir1.Path, True, temp)
    Else
      Call getfiles(Dir1.Path, False, temp)
    End If '------------------------------------------------
 DoEvents
 Next k

 lvFiles.SetFocus
End Sub
Private Function StripNulls(F As String) As String
    StripNulls = Left$(F, InStr(1, F, Chr$(0)) - 1)
End Function

Private Function AddBackslash(s As String) As String
    If Len(s) Then
       If Right$(s, 1) <> "\" Then
             AddBackslash = s & "\"
       Else: AddBackslash = s
       End If
    Else:    AddBackslash = "\"
    End If
End Function
Public Sub getfiles(Path As String, SubFolder As Boolean, Optional Pattern As String = "*.*")
On Error Resume Next
'   Screen.MousePointer = vbHourglass
    Dim li As ListItem
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, fPath As String, FName As String
    Dim sPattern As String
'-------------------------------------------------------
    fPath = AddBackslash(Path)
    sPattern = Pattern
    FName = UCase(fPath & sPattern)
'-------------------------------------------------------
    WFD.cFileName = UCase(WFD.cFileName)
     hFile = FindFirstFile(FName, WFD)
    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
     ' If Left(fPath & StripNulls(WFD.cFileName), 15) <> Text1.Text + "System Volum" Then
        Set li = lvFiles.ListItems.Add(, , fPath + StripNulls(WFD.cFileName))
        Set li = Nothing
     ' End If
    End If
'--------------------------------------------------------
    If hFile > 0 Then
    While FindNextFile(hFile, WFD)
        If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
        '  If Left(fPath & StripNulls(WFD.cFileName), 15) <> Text1.Text + "System Volum" Then
            Set li = lvFiles.ListItems.Add(, , fPath + StripNulls(WFD.cFileName))
            Set li = Nothing
        End If ': End If
    Wend
    End If
'---------------------------------------------------------
    If SubFolder Then
       hFile = FindFirstFile(fPath & "*.*", WFD)
        If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
        StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
           getfiles fPath & StripNulls(WFD.cFileName), True, sPattern
        End If '-------------------------------------------
        While FindNextFile(hFile, WFD)
            If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And _
                StripNulls(WFD.cFileName) <> "." And StripNulls(WFD.cFileName) <> ".." Then
                getfiles fPath & StripNulls(WFD.cFileName), True, sPattern
            End If
        Wend
    End If '------------------------------------------------
    FindClose hFile
    Set li = Nothing
'   Screen.MousePointer = vbDefault
End Sub
