VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{8AB3C700-28FF-476A-852B-5E1558F44746}#1.0#0"; "command.ocx"
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5295
   ClientLeft      =   1335
   ClientTop       =   2355
   ClientWidth     =   4770
   LinkMode        =   1  'Source
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin OsenXPCntrl.Command Cmd 
      Height          =   400
      Index           =   4
      Left            =   0
      TabIndex        =   38
      Top             =   2700
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Close"
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
      MICON           =   "Form6.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   5280
      Top             =   3000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4895
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   16711680
      TabCaption(0)   =   "Equlizer"
      TabPicture(0)   =   "Form6.frx":001C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame(8)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Directory"
      TabPicture(1)   =   "Form6.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PlayList"
      TabPicture(2)   =   "Form6.frx":0054
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Setting"
      TabPicture(3)   =   "Form6.frx":0070
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame(3)"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame 
         Height          =   2250
         Index           =   3
         Left            =   -74880
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   360
         Width           =   4440
         Begin VB.Frame Frame 
            Caption         =   "Speed"
            Height          =   1815
            Index           =   4
            Left            =   2160
            TabIndex        =   18
            Top             =   240
            Width           =   2175
            Begin MSComctlLib.Slider Slider3 
               Height          =   255
               Left            =   240
               TabIndex        =   19
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   1
               Max             =   9
            End
            Begin MSComctlLib.Slider Slider1 
               Height          =   375
               Left            =   240
               TabIndex        =   20
               ToolTipText     =   "PlaySpeed"
               Top             =   960
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               MouseIcon       =   "Form6.frx":008C
               LargeChange     =   1
               Min             =   1
               SelStart        =   5
               TickStyle       =   1
               Value           =   5
            End
            Begin VB.Label Label2 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "    Play Speed Normal"
               Height          =   255
               Left            =   240
               TabIndex        =   4
               Top             =   1440
               Width           =   1695
            End
            Begin VB.Label Label5 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Larenc"
               Height          =   255
               Left            =   480
               TabIndex        =   21
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.Frame Frame 
            Caption         =   "Speker"
            Height          =   1815
            Index           =   5
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1935
            Begin MSComctlLib.Slider Slider4 
               Height          =   255
               Left            =   120
               TabIndex        =   14
               ToolTipText     =   "Balanc Audio"
               Top             =   1080
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               MouseIcon       =   "Form6.frx":0966
               LargeChange     =   1
               Min             =   1
               Max             =   3
               SelStart        =   2
               Value           =   2
            End
            Begin MSComctlLib.Slider Slider13 
               Height          =   255
               Left            =   120
               TabIndex        =   15
               ToolTipText     =   "Vulome"
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               MouseIcon       =   "Form6.frx":1240
               Max             =   100
               SelStart        =   100
               TickStyle       =   3
               Value           =   100
            End
            Begin VB.Image Image12 
               Height          =   360
               Left            =   1440
               Top             =   600
               Width           =   360
            End
            Begin VB.Image Image11 
               Height          =   360
               Left            =   120
               Top             =   600
               Width           =   360
            End
            Begin VB.Label Label8 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "SpeakerBalanc"
               Height          =   255
               Left            =   360
               TabIndex        =   17
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label3 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Vulome"
               Height          =   255
               Left            =   650
               TabIndex        =   16
               Top             =   480
               Width           =   615
            End
         End
      End
      Begin VB.Frame Frame 
         Height          =   2325
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4435
         Begin VB.ListBox List1 
            Height          =   1620
            ItemData        =   "Form6.frx":1B1A
            Left            =   120
            List            =   "Form6.frx":1B1C
            TabIndex        =   11
            ToolTipText     =   "Plase DblClick Me !!"
            Top             =   240
            Width           =   3375
         End
         Begin OsenXPCntrl.Command Cmd 
            Height          =   300
            Index           =   13
            Left            =   1440
            TabIndex        =   28
            Top             =   1920
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "LoadList"
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
            MICON           =   "Form6.frx":1B1E
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
            Height          =   300
            Index           =   12
            Left            =   120
            TabIndex        =   29
            Top             =   1920
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "QuickSave"
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
            MICON           =   "Form6.frx":1B3A
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
            Height          =   500
            Index           =   9
            Left            =   3960
            TabIndex        =   30
            Top             =   1320
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "\/"
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
            MICON           =   "Form6.frx":1B56
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
            Height          =   300
            Index           =   11
            Left            =   2520
            TabIndex        =   31
            Top             =   1920
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "SaveList"
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
            MICON           =   "Form6.frx":1B72
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
            Height          =   300
            Index           =   10
            Left            =   3600
            TabIndex        =   32
            Top             =   1920
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "Return"
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
            MICON           =   "Form6.frx":1B8E
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
            Height          =   500
            Index           =   8
            Left            =   3600
            TabIndex        =   33
            Top             =   1320
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "/\"
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
            MICON           =   "Form6.frx":1BAA
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
            Height          =   300
            Index           =   6
            Left            =   3600
            TabIndex        =   34
            Top             =   600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "Remove"
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
            MICON           =   "Form6.frx":1BC6
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
            Height          =   300
            Index           =   5
            Left            =   3600
            TabIndex        =   35
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "Add"
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
            MICON           =   "Form6.frx":1BE2
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
            Height          =   300
            Index           =   7
            Left            =   3600
            TabIndex        =   36
            Top             =   960
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "Clear"
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
            MICON           =   "Form6.frx":1BFE
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
      Begin VB.Frame Frame 
         Height          =   2370
         Index           =   1
         Left            =   -74930
         TabIndex        =   6
         Top             =   360
         Width           =   4545
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   4335
         End
         Begin VB.DirListBox Dir1 
            Height          =   1665
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   2055
         End
         Begin VB.FileListBox File1 
            Height          =   1650
            Hidden          =   -1  'True
            Left            =   2280
            Pattern         =   "*.M3u;*.wma;*.mp3;*.wav;*.mid;*.avi;*.mpg;*.dat;*.vcd;*.svd;*.ifo;*.mov;*.wmv;*.asf;*.mp2;*.m1v;*.swf"
            TabIndex        =   7
            ToolTipText     =   "Plase DLBClick Me !!"
            Top             =   600
            Width           =   2055
         End
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00FFFF80&
         Height          =   2250
         Index           =   8
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   4455
         Begin OsenXPCntrl.Command Cmd 
            Height          =   300
            Index           =   18
            Left            =   2640
            TabIndex        =   42
            Top             =   1950
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "Home"
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
            MICON           =   "Form6.frx":1C1A
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
            Height          =   300
            Index           =   16
            Left            =   1800
            TabIndex        =   41
            Top             =   1950
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
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
            MICON           =   "Form6.frx":1C36
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
            Height          =   300
            Index           =   17
            Left            =   960
            TabIndex        =   40
            Top             =   1950
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
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
            MICON           =   "Form6.frx":1C52
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
            Height          =   300
            Index           =   14
            Left            =   0
            TabIndex        =   37
            Top             =   1950
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "Color"
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
            MICON           =   "Form6.frx":1C6E
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   18
            X1              =   1440
            X2              =   1440
            Y1              =   615
            Y2              =   2415
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   17
            X1              =   1320
            X2              =   1320
            Y1              =   615
            Y2              =   2415
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   1
            X1              =   2040
            X2              =   2040
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderStyle     =   3  'Dot
            BorderWidth     =   5
            Index           =   2
            X1              =   1800
            X2              =   1800
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   11
            X1              =   2160
            X2              =   2160
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   12
            X1              =   1920
            X2              =   1920
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   3
            X1              =   1680
            X2              =   1680
            Y1              =   600
            Y2              =   2520
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   4
            X1              =   1560
            X2              =   1560
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   5
            X1              =   840
            X2              =   840
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   6
            X1              =   1080
            X2              =   1080
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   13
            X1              =   1200
            X2              =   1200
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   14
            X1              =   960
            X2              =   960
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   7
            X1              =   480
            X2              =   480
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   8
            X1              =   360
            X2              =   360
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   9
            X1              =   720
            X2              =   720
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   10
            X1              =   240
            X2              =   240
            Y1              =   600
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   15
            X1              =   600
            X2              =   600
            Y1              =   615
            Y2              =   2400
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   16
            X1              =   120
            X2              =   120
            Y1              =   600
            Y2              =   2520
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   19
            X1              =   2280
            X2              =   2280
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   20
            X1              =   2400
            X2              =   2400
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   21
            X1              =   2520
            X2              =   2520
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   22
            X1              =   2760
            X2              =   2760
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   23
            X1              =   2640
            X2              =   2640
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   24
            X1              =   2880
            X2              =   2880
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   25
            X1              =   3000
            X2              =   3000
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   26
            X1              =   3120
            X2              =   3120
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   27
            X1              =   3240
            X2              =   3240
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   28
            X1              =   3360
            X2              =   3360
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   29
            X1              =   3480
            X2              =   3480
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   30
            X1              =   3720
            X2              =   3720
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   31
            X1              =   3600
            X2              =   3600
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   32
            X1              =   3840
            X2              =   3840
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   33
            X1              =   3960
            X2              =   3960
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   34
            X1              =   4080
            X2              =   4080
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   35
            X1              =   4320
            X2              =   4320
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line Lin 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   5
            Index           =   36
            X1              =   4200
            X2              =   4200
            Y1              =   600
            Y2              =   2280
         End
      End
   End
   Begin VB.Frame Frame 
      Height          =   2655
      Index           =   7
      Left            =   -120
      TabIndex        =   0
      Top             =   3120
      Width           =   4815
      Begin OsenXPCntrl.Command Cmd 
         Height          =   375
         Index           =   15
         Left            =   1680
         TabIndex        =   39
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cancel"
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
         MICON           =   "Form6.frx":1C8A
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame 
         Caption         =   "Preview"
         Height          =   1335
         Index           =   6
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         Begin VB.Image Image1 
            Height          =   975
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1095
         End
      End
      Begin OsenXPCntrl.Command Cmd 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   25
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Up/\"
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
         MICON           =   "Form6.frx":1CA6
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
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   26
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Down\/"
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
         MICON           =   "Form6.frx":1CC2
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
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "OK"
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
         MICON           =   "Form6.frx":1CDE
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
   Begin VB.Label FG 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label F 
      Caption         =   "Label4"
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   4200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label URL 
      Caption         =   "Label1"
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu cv 
      Caption         =   "cv"
      Visible         =   0   'False
      Begin VB.Menu SetP 
         Caption         =   "SetTo Play"
      End
      Begin VB.Menu Add 
         Caption         =   "AddTo PlayList"
      End
      Begin VB.Menu Add_All 
         Caption         =   "Add All To List"
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub Add_All_Click()
Dim i As Integer
            If File1.List(0) = "" Then Exit Sub
            If Len(Dir1.Path) < 4 Then
                    For i = 0 To File1.ListCount - 1
                    List1.AddItem (Dir1.Path + File1.List(i))
                    Next
            Else
                    For i = 0 To File1.ListCount - 1
                    List1.AddItem (Dir1.Path + "\" + File1.List(i))
                    Next
            End If
End Sub

Private Sub Add_Click()
On Error Resume Next
Dim j As String '-------------------------------
            If Right(Nasa(j$, "\", True), 1) <> ":" Then
            j$ = Dir1.Path + "\" + File1.filename
            Else: j$ = Dir1.Path + File1.filename
            End If '-------------------------------
                    If j = "" Then Exit Sub
                    If j = FG.Caption Then Exit Sub
                    List1.AddItem j: FG.Caption = j
End Sub
Private Sub Cmd_Click(Index As Integer)
On Error Resume Next
Dim B As String
Select Case Index '--------------------------------------------------------------------
Case 1:     Text1.Text = Val(Text1.Text) + 1
Case 2:     Text1.Text = Val(Text1.Text) - 1
Case 3:     FForm1.Image10.Picture = Image1.Picture: _
                ForIcon = Val(Text1.Text): FForm1.Enabled = True
            FForm1.Show
            Unload Me '----------------------------------------------
Case 4:     Form1.For6 = False: Me.Hide
Case 5:     B = "All MediaFiles|*.M3u;*.mp3;*.wav;*.mid;*.avi;*.mpg;*.dat;*.vcd;*.svd;*.ifo;*.mov;*.wmv;*wma;*.asf;*.mp2;*.m1v;*.swf|All Files|*.*"
            ShowOpen B, Form6, Form1, Form1.WindowsMediaPlayer1, "Add media", Pla
Case 6:     List1.RemoveItem List1.ListIndex
Case 7:     List1.Clear '------------------------------------
Case 8:     LChange Form6, List1, True  'Up
Case 9:     LChange Form6, List1, False 'Down
Case 10: '---------------------------------------------------------------------
            If ListFile = "" Then
            Dim A As Integer: A = MsgBox("File Not Found , Your Not Opened PlayList File" + vbCrLf + "Plase Click " _
            + "<Load> Button And Select PlayList File!", vbInformation): Exit Sub: End If
            RETURNForm Form6, Form1.WindowsMediaPlayer1, False
Case 11:    ShowOpen "All PlayList[*.m3u]|*.m3u", Form6, Form1, Form1.WindowsMediaPlayer1, "Save Playlist", Pla
Case 12:    Form1.WindowsMediaPlayer1.settings.autoStart = True
            QuickSave Form6, Form1.WindowsMediaPlayer1, Pla
            Form1.F.Caption = F.Caption: Form1.URL.Caption = URL.Caption
Case 13:    Form1.WindowsMediaPlayer1.settings.autoStart = True
            ShowOpen "All PlayList[*.m3u]|*.m3u", Form6, Form1, Form1.WindowsMediaPlayer1, "Open Playlist", Pla
Case 14:    LineColor False, 0, Me
Case 15:    Unload Me
Case 16:    frmBrowser.Show: frmBrowser.brwWebBrowser.Navigate "http://TcVb.Blogfa.com"
Case 17:    frmBrowser.Show: frmBrowser.brwWebBrowser.Navigate App.Path + "\Help\z.htm" 'help
Case 18:    frmBrowser.Show: frmBrowser.brwWebBrowser.Navigate "http://NaserVb.Blogfa.com" 'home
End Select
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
DriveChange Me, Drive1, Dir1
End Sub

Private Sub File1_DblClick()
On Error Resume Next
SetFile Form1.WindowsMediaPlayer1, File1, Pla, Form1.F, Form1.URL, Me
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 4 Then
    Add_All_Click
ElseIf Button = 3 Then
    Add_All_Click
    Cmd_Click (12)
ElseIf Button = vbRightButton Then
Call PopupMenu(cv)
End If
End Sub

Private Sub Form_Load()
On Error Resume Next '--------------------------------------
 Setng Form6, 18, True, Form1.WindowsMediaPlayer1, Form6.List1, 8
    If Me.Caption <> "Player Option" Then
        Image1.Picture = FForm1.ImageList1.ListImages(1).Picture
        If ForIcon <> "" & ForIcon <> 0 Then _
        Image1.Picture = FForm1.ImageList1.ListImages(Val(ForIcon)).Picture
        Frame(7).Left = -120: Frame(7).Top = -240
    End If '-------------------------------
        Image11.Picture = FForm1.ImageList2.ListImages(5).Picture
        Image12.Picture = FForm1.ImageList2.ListImages(6).Picture

End Sub

Private Sub Form_Unload(Cancel As Integer)
                FForm1.Enabled = True: FForm1.Show
End Sub

Private Sub List1_DblClick()
SetFile Form1.WindowsMediaPlayer1, List1, Pla, Form1.F, Form1.URL, Form6
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Cmd_Click (6)
ElseIf Button = 3 Then
    Cmd_Click (12)
ElseIf Button = 4 Then
    Cmd_Click (7)
End If
End Sub

Private Sub SetP_Click()
On Error Resume Next
File1_DblClick
End Sub

Private Sub Slider1_Scroll()
On Error Resume Next
Call PlaySpeed(Form6, Slider1, Slider3, Form1.WindowsMediaPlayer1, Label2)
End Sub

Private Sub Slider13_Change()
On Error Resume Next
FForm1_WMP_Vol = Slider13.Value
End Sub

Private Sub Slider13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
FForm1_WMP_Vol = Slider13.Value
End Sub

Private Sub Slider13_Scroll()
On Error Resume Next
FForm1_WMP_Vol = Slider13.Value

End Sub

Private Sub Slider3_Change()
On Error Resume Next
Call SmalSpeed(Form6, Form1.WindowsMediaPlayer1, Slider3, Slider1, Label5)

End Sub

Private Sub Slider3_Click()
On Error Resume Next
Call SmalSpeed(Form6, Form1.WindowsMediaPlayer1, Slider3, Slider1, Label5)
End Sub

Private Sub Slider3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Call SmalSpeed(Form6, Form1.WindowsMediaPlayer1, Slider3, Slider1, Label5)
End Sub

Private Sub Slider4_Click()
On Error Resume Next
Call Balance(Slider4, Form1.WindowsMediaPlayer1, Image11 _
                 , Image12, Form6)

End Sub
Private Sub Text1_Change()
On Error Resume Next '------------------------------------
                    If Text1.Text > 5 Then
                    Text1.Text = 1
                    ElseIf Text1.Text < 1 Then
                    Text1.Text = 5
                    End If '-------------------------------
                    Image1.Picture = FForm1.ImageList1.ListImages(Val(Text1.Text)).Picture
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
                    Slider13.Value = FForm1_WMP_Vol
End Sub
