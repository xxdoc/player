VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8AB3C700-28FF-476A-852B-5E1558F44746}#1.0#0"; "command.ocx"
Begin VB.Form FForm1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Player 2.4"
   ClientHeight    =   7650
   ClientLeft      =   1710
   ClientTop       =   1980
   ClientWidth     =   10530
   Icon            =   "FForm1.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2665.073
   ScaleMode       =   0  'User
   ScaleWidth      =   10733.81
   Begin VB.Frame Frame 
      Caption         =   "Play Setting"
      Height          =   2010
      Index           =   7
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   2160
      Width           =   4435
      Begin VB.Frame Frame 
         Height          =   1815
         Index           =   9
         Left            =   2160
         TabIndex        =   19
         Top             =   120
         Width           =   2175
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            MousePointer    =   99
            LargeChange     =   1
            Min             =   1
            SelStart        =   5
            Value           =   5
         End
         Begin MSComctlLib.Slider Slider3 
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            MousePointer    =   99
            LargeChange     =   1
            Max             =   9
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "    Play Speed Normal"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Larenc"
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Speker"
         Height          =   1695
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1935
         Begin MSComctlLib.Slider Slider2 
            Height          =   315
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            LargeChange     =   1
            Min             =   1
            Max             =   3
            SelStart        =   2
            Value           =   2
         End
         Begin MSComctlLib.Slider Slider4 
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            MousePointer    =   99
            LargeChange     =   10
            Max             =   100
            SelStart        =   100
            TickStyle       =   3
            Value           =   100
            TextPosition    =   1
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
            TabIndex        =   18
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vulome"
            Height          =   255
            Left            =   650
            TabIndex        =   17
            Top             =   480
            Width           =   615
         End
      End
   End
   Begin OsenXPCntrl.Command cmd 
      Height          =   255
      Index           =   15
      Left            =   4238
      TabIndex        =   53
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "<"
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
      MPTR            =   99
      MICON           =   "FForm1.frx":08CA
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
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   35
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "#"
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
      MPTR            =   99
      MICON           =   "FForm1.frx":08E6
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
      Height          =   255
      Index           =   13
      Left            =   4440
      TabIndex        =   32
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   ">"
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
      MPTR            =   99
      MICON           =   "FForm1.frx":0902
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
      Height          =   255
      Index           =   12
      Left            =   4440
      TabIndex        =   31
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "<"
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
      MPTR            =   99
      MICON           =   "FForm1.frx":091E
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
      Height          =   255
      Index           =   20
      Left            =   3960
      TabIndex        =   30
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
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
      MPTR            =   99
      MICON           =   "FForm1.frx":093A
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox SerTxt 
      Height          =   285
      Left            =   1320
      LinkItem        =   "ctext"
      LinkTopic       =   "project1|form7"
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   10
      Left            =   360
      TabIndex        =   22
      Top             =   1200
      Width           =   4455
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   900
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   3600
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   5
         autoStart       =   0   'False
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
         _cx             =   6350
         _cy             =   1588
      End
   End
   Begin VB.Frame Frame 
      Height          =   735
      Index           =   1
      Left            =   6240
      TabIndex        =   0
      ToolTipText     =   "contorl programs"
      Top             =   2400
      Width           =   3375
      Begin OsenXPCntrl.Command cmd 
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   34
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Convert"
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
         MPTR            =   99
         MICON           =   "FForm1.frx":0956
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
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Open Media"
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
         MPTR            =   99
         MICON           =   "FForm1.frx":0972
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
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   57
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
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
         MPTR            =   99
         MICON           =   "FForm1.frx":098E
         PICN            =   "FForm1.frx":09AA
         PICH            =   "FForm1.frx":0D8E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.Command Command2 
         Height          =   375
         Left            =   1320
         TabIndex        =   58
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
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
         MPTR            =   99
         MICON           =   "FForm1.frx":1176
         PICN            =   "FForm1.frx":1192
         PICH            =   "FForm1.frx":1517
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame 
      Height          =   615
      Index           =   3
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      Begin OsenXPCntrl.Command cmd 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   67
         Top             =   195
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
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
         MICON           =   "FForm1.frx":18B6
         PICN            =   "FForm1.frx":18D2
         PICH            =   "FForm1.frx":1B8F
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.Command cmd 
         Height          =   300
         Index           =   11
         Left            =   960
         TabIndex        =   49
         Top             =   195
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "S"
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
         MPTR            =   99
         MICON           =   "FForm1.frx":1FA4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.Command Label4 
         Height          =   300
         Left            =   2880
         TabIndex        =   64
         Top             =   195
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "About"
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
         MICON           =   "FForm1.frx":1FC0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.Command Image1 
         Height          =   300
         Left            =   2040
         TabIndex        =   63
         Top             =   195
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
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
         MPTR            =   99
         MICON           =   "FForm1.frx":1FDC
         PICN            =   "FForm1.frx":1FF8
         PICH            =   "FForm1.frx":2462
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
         Height          =   300
         Index           =   25
         Left            =   3600
         TabIndex        =   51
         Top             =   200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   ">>>>"
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
         MPTR            =   99
         MICON           =   "FForm1.frx":28CC
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
         Height          =   300
         Index           =   22
         Left            =   1200
         TabIndex        =   50
         Top             =   195
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "Go Sink"
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
         MPTR            =   99
         MICON           =   "FForm1.frx":28E8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.Command Option1 
         Height          =   300
         Left            =   120
         TabIndex        =   48
         Top             =   195
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "Option"
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2904
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFF80&
      Height          =   2010
      Index           =   6
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   4455
      Begin OsenXPCntrl.Command cmd 
         Height          =   300
         Index           =   31
         Left            =   960
         TabIndex        =   70
         Top             =   1710
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
         MICON           =   "FForm1.frx":2920
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
         Height          =   300
         Index           =   26
         Left            =   1800
         TabIndex        =   69
         Top             =   1710
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
         MICON           =   "FForm1.frx":293C
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
         Height          =   300
         Index           =   27
         Left            =   2640
         TabIndex        =   68
         Top             =   1710
         Width           =   855
         _ExtentX        =   1508
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
         MICON           =   "FForm1.frx":2958
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
         Height          =   300
         Index           =   24
         Left            =   120
         TabIndex        =   52
         Top             =   1710
         Width           =   735
         _ExtentX        =   1296
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2974
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
         Y1              =   375
         Y2              =   2175
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   17
         X1              =   1320
         X2              =   1320
         Y1              =   375
         Y2              =   2175
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   1
         X1              =   2040
         X2              =   2040
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderStyle     =   3  'Dot
         BorderWidth     =   5
         Index           =   2
         X1              =   1800
         X2              =   1800
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   11
         X1              =   2160
         X2              =   2160
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   12
         X1              =   1920
         X2              =   1920
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   3
         X1              =   1680
         X2              =   1680
         Y1              =   360
         Y2              =   2280
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   4
         X1              =   1560
         X2              =   1560
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   5
         X1              =   840
         X2              =   840
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   6
         X1              =   1080
         X2              =   1080
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   13
         X1              =   1200
         X2              =   1200
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   14
         X1              =   960
         X2              =   960
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   7
         X1              =   480
         X2              =   480
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   8
         X1              =   360
         X2              =   360
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   9
         X1              =   720
         X2              =   720
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   10
         X1              =   240
         X2              =   240
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   15
         X1              =   600
         X2              =   600
         Y1              =   375
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   16
         X1              =   120
         X2              =   120
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   19
         X1              =   2280
         X2              =   2280
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   20
         X1              =   2400
         X2              =   2400
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   21
         X1              =   2520
         X2              =   2520
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   22
         X1              =   2760
         X2              =   2760
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   23
         X1              =   2640
         X2              =   2640
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   24
         X1              =   2880
         X2              =   2880
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   25
         X1              =   3000
         X2              =   3000
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   26
         X1              =   3120
         X2              =   3120
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   27
         X1              =   3240
         X2              =   3240
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   28
         X1              =   3360
         X2              =   3360
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   29
         X1              =   3480
         X2              =   3480
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   30
         X1              =   3720
         X2              =   3720
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   31
         X1              =   3600
         X2              =   3600
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   32
         X1              =   3840
         X2              =   3840
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   33
         X1              =   3960
         X2              =   3960
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   34
         X1              =   4080
         X2              =   4080
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   35
         X1              =   4200
         X2              =   4200
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Line Lin 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Index           =   36
         X1              =   4320
         X2              =   4320
         Y1              =   600
         Y2              =   2040
      End
      Begin VB.Label text1 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3960
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label x 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   11
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   4215
      Begin OsenXPCntrl.Command cmd 
         Height          =   300
         Index           =   30
         Left            =   3240
         TabIndex        =   66
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "AviMaker"
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
         MICON           =   "FForm1.frx":2990
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
         Height          =   300
         Index           =   21
         Left            =   2280
         TabIndex        =   38
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "Cradieat"
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
         MICON           =   "FForm1.frx":29AC
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
         Height          =   300
         Index           =   14
         Left            =   1200
         TabIndex        =   37
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "SelectIcon"
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
         MICON           =   "FForm1.frx":29C8
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
         Height          =   300
         Index           =   16
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "Manycopy"
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
         MICON           =   "FForm1.frx":29E4
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line38 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   0
         X2              =   4200
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "List Of Files"
      Height          =   2009
      Index           =   4
      Left            =   6000
      TabIndex        =   2
      Top             =   5400
      Width           =   4435
      Begin OsenXPCntrl.Command cmd 
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   47
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Set"
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2A00
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
         Height          =   255
         Index           =   19
         Left            =   960
         TabIndex        =   46
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Load"
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2A1C
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
         Height          =   255
         Index           =   17
         Left            =   1800
         TabIndex        =   45
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Save"
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2A38
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
         Height          =   255
         Index           =   10
         Left            =   2760
         TabIndex        =   44
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2A54
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
         Height          =   255
         Index           =   9
         Left            =   3600
         TabIndex        =   43
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2A70
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
         Height          =   255
         Index           =   8
         Left            =   3600
         TabIndex        =   42
         Top             =   1320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2A8C
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
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   41
         Top             =   960
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2AA8
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
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   40
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Remov"
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2AC4
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
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   39
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
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
         MPTR            =   99
         MICON           =   "FForm1.frx":2AE0
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox List1 
         Height          =   1425
         ItemData        =   "FForm1.frx":2AFC
         Left            =   120
         List            =   "FForm1.frx":2AFE
         TabIndex        =   14
         ToolTipText     =   "Plase DblClick Me !!"
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000003&
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   360
      MousePointer    =   99  'Custom
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   26
      ToolTipText     =   "Quick About"
      Top             =   120
      Width           =   735
      Begin VB.Image Image10 
         Height          =   570
         Left            =   0
         MousePointer    =   14  'Arrow and Question
         Stretch         =   -1  'True
         ToolTipText     =   "Quick About"
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   5
      Left            =   3000
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      Begin OsenXPCntrl.Command cmd 
         Height          =   300
         Index           =   23
         Left            =   0
         TabIndex        =   54
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "Open"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "FForm1.frx":2B00
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
         Height          =   375
         Index           =   28
         Left            =   360
         TabIndex        =   55
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "FForm1.frx":2B1C
         PICN            =   "FForm1.frx":2B38
         PICH            =   "FForm1.frx":2F1C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.Command cmd 
         Height          =   375
         Index           =   29
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
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
         MICON           =   "FForm1.frx":3304
         PICN            =   "FForm1.frx":3320
         PICH            =   "FForm1.frx":36A5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line37 
         X1              =   0
         X2              =   4200
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2009
      Index           =   2
      Left            =   5760
      TabIndex        =   10
      Top             =   3240
      Width           =   4435
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   4335
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         Height          =   1455
         Hidden          =   -1  'True
         Left            =   2280
         Pattern         =   "*.M3u;*.mp3;*.wav;*.mid;*.avi;*.mpg;*.dat;*.vcd;*.svd;*.ifo;*.mov;*.wmv;*.wma;*.asf;*.mp2;*.m1v;*.swf"
         TabIndex        =   11
         ToolTipText     =   "Plase DLBClick Me !!"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Line Line39 
         BorderColor     =   &H80000003&
         X1              =   0
         X2              =   0
         Y1              =   2040
         Y2              =   0
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   5280
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Play List File|*.M3U"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   6600
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   5160
      TabIndex        =   4
      Top             =   0
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Equlizer"
      TabPicture(0)   =   "FForm1.frx":3A44
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Drive Serch"
      TabPicture(1)   =   "FForm1.frx":3A60
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7395
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1903
            MinWidth        =   1903
            TextSave        =   "02:16 È.Ù"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1903
            MinWidth        =   1903
            TextSave        =   "2007/09/02"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4800
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   37
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FForm1.frx":3A7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FForm1.frx":3F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FForm1.frx":4E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FForm1.frx":5750
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FForm1.frx":5A6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   2415
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   16711680
      TabCaption(0)   =   "Equalizer"
      TabPicture(0)   =   "FForm1.frx":6944
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Drive Serch"
      TabPicture(1)   =   "FForm1.frx":6960
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "List"
      TabPicture(2)   =   "FForm1.frx":697C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Setting"
      TabPicture(3)   =   "FForm1.frx":6998
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Player Home"
      TabPicture(4)   =   "FForm1.frx":69B4
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame(12)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   2055
         Index           =   12
         Left            =   0
         TabIndex        =   65
         Top             =   360
         Width           =   4455
      End
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   4800
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FForm1.frx":69D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FForm1.frx":40862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FForm1.frx":6DAD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FForm1.frx":6E1E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FForm1.frx":6E8F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FForm1.frx":6F006
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label URL 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5160
      TabIndex        =   28
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label F 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label FG 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   7080
      Width           =   255
   End
   Begin VB.Menu Popup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu O 
         Caption         =   "File"
         Begin VB.Menu OpenAudio 
            Caption         =   "Open Media"
         End
         Begin VB.Menu openL 
            Caption         =   "Open For List"
         End
      End
      Begin VB.Menu rfv 
         Caption         =   "-"
      End
      Begin VB.Menu Settin 
         Caption         =   "Setting"
         Begin VB.Menu Vol 
            Caption         =   "Vulume"
            Begin VB.Menu VolU 
               Caption         =   "Vulome Up      +"
            End
            Begin VB.Menu VolD 
               Caption         =   "Vulome Down  -"
            End
         End
         Begin VB.Menu Balanc 
            Caption         =   "Balanc(Speaker Change)"
            Begin VB.Menu SpeakL 
               Caption         =   "Speaker Left"
            End
            Begin VB.Menu SpeakN 
               Caption         =   "Speaker Normal"
            End
            Begin VB.Menu SpeakR 
               Caption         =   "Speaker Right"
            End
         End
         Begin VB.Menu Speed 
            Caption         =   "PlaySpeed"
            Begin VB.Menu SpeedU 
               Caption         =   "Speed Up      +"
            End
            Begin VB.Menu SpeedD 
               Caption         =   "Speed Down  -"
            End
         End
      End
      Begin VB.Menu bzc 
         Caption         =   "-"
      End
      Begin VB.Menu Size 
         Caption         =   "View"
         Begin VB.Menu minimize 
            Caption         =   "MakeAvi"
         End
         Begin VB.Menu ful 
            Caption         =   "Option"
         End
         Begin VB.Menu max 
            Caption         =   "Max Size"
         End
         Begin VB.Menu min 
            Caption         =   "Min Size"
         End
      End
      Begin VB.Menu jkhhhh 
         Caption         =   "-"
      End
      Begin VB.Menu Tools 
         Caption         =   "Tools"
         Begin VB.Menu openV 
            Caption         =   "Convert"
         End
         Begin VB.Menu Arm 
            Caption         =   "Icon For Sofware"
         End
         Begin VB.Menu Mcopy 
            Caption         =   "ManyCopy"
         End
      End
      Begin VB.Menu fgfdgfdvgf 
         Caption         =   "-"
      End
      Begin VB.Menu Sink 
         Caption         =   "Sink"
         Begin VB.Menu SSink 
            Caption         =   "Select Sink"
         End
         Begin VB.Menu QCS 
            Caption         =   "Quick Change Sink"
         End
      End
      Begin VB.Menu mkw 
         Caption         =   "-"
      End
      Begin VB.Menu Sety 
         Caption         =   "PlayList"
         Begin VB.Menu QSave 
            Caption         =   "Quick Save"
         End
         Begin VB.Menu LoadL 
            Caption         =   "Load List"
         End
         Begin VB.Menu SaveL 
            Caption         =   "Save List"
         End
      End
      Begin VB.Menu dghdfhgfh 
         Caption         =   "-"
      End
      Begin VB.Menu DST 
         Caption         =   "Set"
         Begin VB.Menu Return 
            Caption         =   "Return To The End Load List"
         End
         Begin VB.Menu setD 
            Caption         =   "Set Drive Of Play"
         End
         Begin VB.Menu Set 
            Caption         =   "Set List Of Play"
         End
      End
      Begin VB.Menu hfghjgf 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu pjs 
         Caption         =   "-"
      End
      Begin VB.Menu end 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Cop 
      Caption         =   "Set"
      Visible         =   0   'False
      Begin VB.Menu Play 
         Caption         =   "Set To Play"
      End
      Begin VB.Menu PAll 
         Caption         =   "PlayAll"
      End
      Begin VB.Menu A 
         Caption         =   "Add To List"
      End
      Begin VB.Menu Add_All 
         Caption         =   "Add All To List"
      End
   End
End
Attribute VB_Name = "FForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub A_Click()
On Error Resume Next '-------------------------------------------------------
            If Len(Dir1.Path) < 4 Then
                    If FG.Caption <> (Dir1.Path + File1.filename) Then
                    List1.AddItem (Dir1.Path + File1.filename)
                    FG.Caption = (Dir1.Path + File1.filename)
                    End If
            Else '------------------------------------------
                    If FG.Caption <> (Dir1.Path + "\" + File1.filename) Then
                    List1.AddItem (Dir1.Path + "\" + File1.filename)
                    FG.Caption = (Dir1.Path + "\" + File1.filename)
                    End If
            End If
End Sub

Private Sub about_Click()
                    Label4_Click
End Sub

Private Sub Add_All_Click()
Dim i As Integer
            If File1.List(0) = "" Then Exit Sub
            If Len(Dir1.Path) < 4 Then
                    For i = 0 To File1.ListCount - 1
                    List1.AddItem (Dir1.Path + File1.List(i))
                    Next
            Else '------------------------------------------
                    For i = 0 To File1.ListCount - 1
                    List1.AddItem (Dir1.Path + "\" + File1.List(i))
                    Next
            End If
End Sub

Private Sub Arm_Click()
                    Cmd_Click (14)
End Sub
                    
Private Sub Cmd20()
        On Error Resume Next
        Dim Fil As String
        Form6.Visible = False
        LineColor True, FForm1.Lin(1).BorderColor, Form6
        Form1.Show: Form1.WindowsMediaPlayer1.settings.Rate = 1.7
If For1_X <> 0 Then Form1.Move For1_X
If F.Caption = "Video" Then '------------------------------------------
If Timer2.Enabled = False Then Form1.WindowsMediaPlayer1.settings.autoStart = False
 Swich Form3.WindowsMediaPlayer1.URL, F
 ElseIf UCase(Right(WindowsMediaPlayer1.URL, 3)) = "M3U" Then
                 Form1.Show: Me.Hide
                 DfA = URL.Caption '-----------------------------------
                Call Opening(Form6, DfA): Fil = DfA
                Call ListE(Form1, Form6.List1, Form1.WindowsMediaPlayer1, DfA, False)
                 Swich Fil, Form6.F
Else: Swich URL.Caption, F: End If '-----------------------------------
   Form1.WindowsMediaPlayer1.settings.Rate = 1
End Sub
Private Sub Swich(URL As String, Label As Label)
If Label.Caption = "Video" Then  '-------------If Playing Video-----------------------------
    Form1.WindowsMediaPlayer1.Controls.currentPosition = _
    Form3.WindowsMediaPlayer1.Controls.currentPosition
    Form1.WindowsMediaPlayer1.settings.autoStart = Form3.WindowsMediaPlayer1.settings.autoStart
    Form1.WindowsMediaPlayer1.URL = URL: WindowsMediaPlayer1.URL = ""
    Unload Form3: Form1.F.Caption = "Video": Form1.URL.Caption = URL
Else '-----------------------------------If Playing Audio----------------------------
    Form1.WindowsMediaPlayer1.Controls.currentPosition = _
          WindowsMediaPlayer1.Controls.currentPosition
    Form1.WindowsMediaPlayer1.settings.autoStart = WindowsMediaPlayer1.settings.autoStart
    Form1.WindowsMediaPlayer1.URL = URL: Form1.F.Caption = "For"
    WindowsMediaPlayer1.URL = "": Form1.URL.Caption = URL
End If '-----------------------------------------------------------------------------
    Me.Hide
End Sub

Private Sub Cmd_Click(Index As Integer)
On Error Resume Next
Dim B$
Select Case Index
       Case 0:      Dim FD As Long, X1%, X2%, X3%
                    Randomize Timer
                    X1 = Int(Rnd() * 70): X2 = Int(Rnd() * 70): X3 = Int(Rnd() * 70)
                    FD = RGB(X1 + 170, X2 + 170, X3 + 170)
                    Me.BackColor = FD: Label4.BackColor = FD
                    For i = 1 To 12
                    Frame(i).BackColor = FD: Next
                    Dir1.BackColor = FD: Drive1.BackColor = FD: File1.BackColor = FD: List1.BackColor = FD
                    Setting.ProgramColor = FD
                    Open App.Path + "\NMS\user.dll" For Random Access Write As #2
                    Put #2, , Setting: Close
       Case 1: B = "All MediaFiles|*.M3u;*.mp3;*.wav;*.mid;*.avi;*.mpg;*.dat;*.vcd;*.svd;*.ifo;*.mov;*.wmv;*.wma;*.asf;*.mp2;*.m1v;*.swf|All Files|*.*"
                    ShowOpen B, Me, Me, WindowsMediaPlayer1, "Open Media", True
       Case 2: Unload Me
       Case 3:
                    If FForm1.Height = 5400 Then
                        FForm1.Height = 2955
                        Frame(11).Visible = False
                    Else '---------------------------------
                        FForm1.Height = 5400
                        Frame(11).Visible = True
                    End If '-------------------------------
       Case 4:      If DDE = False Then Shell App.Path + "\Copy.exe", vbNormalFocus
                    For i = 0 To 120: SerTxt.Text = i: Next
                    SerTxt.Text = ""
                    SerTxt.Text = "Convert"
       Case 5: '-------------------------------------------
                    B = "All MediaFiles|*.mp3;*.wav;*.mid;*.avi;*.mpg;*.dat;*.vcd;*.svd;*.ifo;*.mov;*.wmv;*.Wma;*.asf;*.mp2;*.m1v;*.swf|All Files|*.*"
                    ShowOpen B, Me, Me, WindowsMediaPlayer1, "Add media", True
       Case 6: '-------------------------------------------
                    If List1.Text = "" Then
                    MsgBox ("File Not Found")
                    Else
                    List1.RemoveItem (List1.ListIndex)
                    End If
       Case 7: List1.Clear
       Case 8: Call LChange(FForm1, List1, True)
       Case 9: Call LChange(FForm1, List1, False)
       Case 10:     Dim CLK As Integer
                    If ListFile <> "" Then
                    RETURNForm Me, WindowsMediaPlayer1, True
                    Else
                    CLK = MsgBox("File Not Found!,You'r Not Opened List-Plase Click <<Load List>> & Open List", vbInformation)
                    End If
       Case 11: '-----------------------------------------
                    If cmd(15).Visible = True Then
                    Call SinkS(Me)
                    Else: Call SinkP(Me)
                    End If
       Case 12: FormMinSize Me: SinkM = "MinS"
       Case 13: FormSizeChange Me
       Case 14:     Form6.Width = 4335: Form6.Height = 2250
                    FForm1.Enabled = False: Form6.Show
                    Form6.Caption = "Plase Select Icon"
                    Form6.SSTab1.Visible = False
       Case 15: Call MinSize(Me): SinkM = "MinP"
       Case 16: If DDE = False Then Shell App.Path + "\Copy.exe", vbNormalFocus
                    For i = 0 To 120: SerTxt.Text = i: Next
                    SerTxt.Text = "Copy"
       Case 17: ShowOpen "All PlayListFile[*.m3u]|*.m3u", Me, Me, WindowsMediaPlayer1, "Save Playlist", True
       Case 18: Call QuickSave(FForm1, WindowsMediaPlayer1, True)
       Case 19: ShowOpen "All PlayListFile[*.m3u]|*.m3u", Me, Me, WindowsMediaPlayer1, "Open Playlist", True
       Case 20: Cmd20
       Case 21: Credit
       Case 22: FLoad.Show
       Case 23: Cmd_Click (1) 'Code Cmd27 = Cm1
       Case 24: LineColor False, 0, Me
       Case 25:
                    If FForm1.Width = 9360 Then
                          FForm1.Width = 4740: cmd(25).Caption = ">>>>"
                    Else: FForm1.Width = 9360: cmd(25).Caption = "<<<<"
                    End If
       Case 26: frmBrowser.Show: frmBrowser.brwWebBrowser.Navigate App.Path + "\help\z.html"
       Case 27: frmBrowser.Show: frmBrowser.brwWebBrowser.Navigate "http://tcvb.blogfa.com"
       Case 28: Unload Me
       Case 29: Me.WindowState = 1: SerTxt.Text = "": SerTxt.Text = "Min"
       Case 30: Call Shell(App.Path + "\Pic To Avi.exe", vbNormalFocus): Avi = True
       Case 31: frmBrowser.Show: frmBrowser.brwWebBrowser.Navigate "http://naservb.blogfa.com"
End Select
End Sub

Private Sub Command2_Click()
    Me.WindowState = 1: SerTxt.Text = "": SerTxt.Text = "Min"
End Sub
Private Sub end_Click()
    Unload Me
End Sub

Private Sub File1_DblClick()
SetFile WindowsMediaPlayer1, File1, True, F, URL, Me
A_Click
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next '------------------
    If Button = 4 Then
    Add_All_Click
ElseIf Button = 3 Then
    PAll_Click
ElseIf Button = vbRightButton Then
    A_Click
End If
If Button = vbRightButton Then
If File1.filename <> "" Then
Call PopupMenu(Cop)
End If '-------------------------------
End If
End Sub
Private Sub Form_Activate()
If SerTxt.Text = "Min" Then SerTxt.Text = "Normal"
End Sub
Private Sub Form_Load()
'On Error Resume Next
                    LoadForm Me
                    Image11.Picture = ImageList2.ListImages(5).Picture
                    Image12.Picture = ImageList2.ListImages(6).Picture
DefultPath = App.Path
End Sub

Private Sub Form_Resize()
If SerTxt.Text = "Min" Then SerTxt.Text = "Normal"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo TR '----------------------------------------
            SerTxt.Text = "": SerTxt.Text = "Exit"
             Dim strans As Integer, A As Integer, B As Integer, n As Integer, v As Integer, c As Integer, X As Integer
             If ASD.FolderExists(App.Path + "\nms") = False Then GoTo TR
             If Pla = True Then '---------------------------------------
             URL.Caption = Form3.WindowsMediaPlayer1.URL
             Else: If WindowsMediaPlayer1.URL <> "" Then _
             URL.Caption = WindowsMediaPlayer1.URL
             End If '-------------------------------
             Unload FLoad: Unload Form3: Unload Form6
             Unload Form1: Unload frmBrowser
              '----------Proses Setting------------------------------
If cmd(15).Visible = True Then
   If cmd(15).Caption = "<" Then
   SinkM = "FTAB" '-Sink Profshnal-----------
   Else: SinkM = "MinP": End If
Else '-------------Sink Standard-------------
   If Frame(5).Visible = True Then
   SinkM = "MinS"
   Else: SinkM = "FORM": End If
End If: DAT1.Sink = SinkM
            DAT1.WmpURL = URL.Caption + "#"
            DAT1.IconForm = ForIcon
            Open App.Path + "\NMS\DAT1.DLL" For Random Access Write As #4
            Put #4, , DAT1
            Close #4
If ASD.FileExists(App.Path + "\command.ocx") = False Then UnProses (Path)
TR: '------------------------------------------------------------------
                    End
End Sub

Sub Menu(Button As Integer)
On Error Resume Next
If Button = vbRightButton Then
Call PopupMenu(Popup)
End If
End Sub

Private Sub Frame_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then _
Call PopupMenu(Popup)
End Sub

Private Sub ful_Click()
    Option1_Click
End Sub

Private Sub Image1_Click()
frmBrowser.Show
End Sub

Private Sub Image10_Click()
Credit
End Sub

Private Sub List1_DblClick()
SetFile WindowsMediaPlayer1, List1, True, F, URL, Me
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 4 Then
      Cmd_Click (7)
ElseIf Button = vbRightButton Then
      Cmd_Click (6)
ElseIf Button = 3 Then
      Cmd_Click (18)
End If
End Sub

Private Sub LoadL_Click()
    Cmd_Click (19)
End Sub

Private Sub max_Click()
    Cmd_Click (13)
End Sub

Private Sub Mcopy_Click()
    Cmd_Click (16)
End Sub
Private Sub minimize_Click()
    Cmd_Click (30)
End Sub

Private Sub openAudio_Click()
    Cmd_Click (1)
End Sub

Private Sub min_Click()
If SSTab3.Visible = True Then Cmd_Click (15)
If SSTab3.Visible = False Then Cmd_Click (13)
End Sub

Private Sub Dir1_Change()
On Error Resume Next
                    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
                    DriveChange Me, Drive1, Dir1
End Sub
Private Sub Label4_Click()
On Error Resume Next
                    Intcnt = 1: FAbout.Show: FForm1.Enabled = False
End Sub
Private Sub openL_Click()
    Cmd_Click (5)
End Sub

Private Sub Option1_Click()
FLoad.Show: FLoad.Command3.Visible = True
FLoad.Frame4.Visible = True: FLoad.Frame2.Visible = False: FLoad.Caption = "Option Softwar"
Call FLoad.Lod
End Sub
Private Sub PAll_Click()
    Add_All_Click
    Cmd_Click (18)
End Sub

Private Sub Play_Click()
    File1_DblClick
End Sub

Private Sub QCS_Click()
    Cmd_Click (11)
End Sub

Private Sub QSave_Click()
    Cmd_Click (18)
End Sub

Private Sub Return_Click()
    Cmd_Click (10)
End Sub

Private Sub SaveL_Click()
    Cmd_Click (17)
End Sub

Private Sub SerTxt_Change()
Select Case SerTxt.Text
        Case "Load":   DDE = True: SerTxt.Text = ""
        Case "Close":  DDE = False
        Case "Exit":   DDE = False
        Case "Avi":    Avi = True
        Case "Unload": Avi = False
End Select
End Sub
Private Sub SerTxt_LinkClose()
DDE = False
End Sub

Private Sub SerTxt_LinkError(LinkErr As Integer)
DDE = False
End Sub

Private Sub SerTxt_LinkOpen(Cancel As Integer)
DDE = True
End Sub

Private Sub set_Click()
On Error Resume Next
List1_DblClick
End Sub

Private Sub setD_Click()
On Error Resume Next
File1_DblClick
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Menu(Button)
End Sub

Private Sub Slider1_Change()
On Error Resume Next '-------------------------------
Call PlaySpeed(FForm1 _
, Slider1, Slider3, WindowsMediaPlayer1, Label2)
End Sub

Private Sub Slider3_Change()
On Error Resume Next '-------------------------------
Call SmalSpeed(Me, _
WindowsMediaPlayer1, Slider3, Slider1, Label5)
End Sub

Private Sub Slider2_Change()
On Error Resume Next '-------------------------------
Call Balance(Slider2, WindowsMediaPlayer1, Image11, Image12, Me)
End Sub

Private Sub Slider4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider4_Scroll
End Sub

Private Sub Slider4_Scroll()
On Error Resume Next
                    If WindowsMediaPlayer1.settings.mute = False Then
                    WindowsMediaPlayer1.settings.Volume = Slider4.Value
                    FForm1_WMP_Vol = Slider4.Value
                    End If
End Sub

Private Sub SpeakL_Click()
                    Slider2.Value = 1
End Sub

Private Sub SpeakN_Click()
                    Slider2.Value = 2
End Sub

Private Sub SpeakR_Click()
                    Slider2.Value = 3
End Sub

Private Sub SpeedD_Click()
On Error Resume Next
                    Slider1.Value = Slider1.Value - 1
End Sub

Private Sub SpeedU_Click()
On Error Resume Next
                    Slider1.Value = Slider1.Value + 1
End Sub

Private Sub SSink_Click()
                    Cmd_Click (22)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 1 Then
    Frame(2).Visible = True: Frame(6).Visible = False
Else '-------------------------------
    Frame(2).Visible = False: Frame(6).Visible = True
End If '-------------------------------
End Sub

Private Sub SSTab1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Menu(Button)
End Sub

Private Sub SSTab3_Click(PreviousTab As Integer)
On Error Resume Next '-------------------------------
         Frame(6).Visible = False
         Frame(2).Visible = False
         Frame(4).Visible = False
         Frame(7).Visible = False
         Frame(1).Visible = False
         Frame(3).Visible = False
         Frame(11).Visible = False
         Picture1.Visible = False
Select Case SSTab3.Tab
     Case 0: Frame(6).Visible = True
     Case 1: Frame(2).Visible = True
     Case 2: Frame(4).Visible = True
     Case 3: Frame(7).Visible = True
     Case 4: Frame(1).Visible = True
             Frame(3).Visible = True
             Frame(11).Visible = True
             Picture1.Visible = True
End Select
End Sub

Private Sub SSTab3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Menu(Button)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Call Control '-------------------------------
                    X.Caption = WindowsMediaPlayer1.Controls.currentPosition
                    If Text1.Caption = "" Then
                    Text1.Caption = WindowsMediaPlayer1.Controls.currentPosition
                    X.Caption = "23"
                    End If '-------------------------------
                    If X.Caption = Text1.Caption Then
                    Timer2.Enabled = False
                    Else '-------------------------------
                    Timer2.Enabled = True
                    Text1.Caption = X.Caption
                    End If '-------------------------------
                      Slider3_Change
                    If FForm1_WMP_Vol <> 0 Then
                    If WindowsMediaPlayer1.settings.mute = False Then Slider4.Value = FForm1_WMP_Vol
                    End If '-------------------------------
 SerTxt_Change
End Sub

Private Sub Timer2_Timer()
                    EqualizerForm Me
End Sub

Private Sub VolD_Click()
On Error Resume Next
Slider4.Value = Slider4.Value - 25
End Sub

Private Sub VolU_Click()
On Error Resume Next
Slider4.Value = Slider4.Value + 25
End Sub

Private Sub Control()
On Error Resume Next
If F.Caption = "Video" Then
If Timer2.Enabled = False Then
WindowsMediaPlayer1.Controls.stop
End If
End If
End Sub

Private Sub WindowsMediaPlayer1_MouseMove(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
On Error Resume Next
                    If WindowsMediaPlayer1.settings.mute = False Then _
                    FForm1_WMP_Vol = WindowsMediaPlayer1.settings.Volume
End Sub


