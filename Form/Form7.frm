VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{8AB3C700-28FF-476A-852B-5E1558F44746}#1.0#0"; "Command.ocx"
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Can 
      BackColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   795
      TabIndex        =   45
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
      Begin VB.Label Cam 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin OsenXPCntrl.Command Cmd 
         CausesValidation=   0   'False
         Default         =   -1  'True
         Height          =   405
         Index           =   3
         Left            =   7560
         TabIndex        =   58
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&Convert"
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
         MICON           =   "Form7.frx":0442
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin OsenXPCntrl.Command Cmd 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   4
         Left            =   6120
         TabIndex        =   57
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "Copy&Wizard \/"
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
         FOCUSR          =   0   'False
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "Form7.frx":045E
         UMCOL           =   0   'False
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
      End
      Begin OsenXPCntrl.Command Cmd 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   5
         Left            =   4920
         TabIndex        =   56
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&PlayFile"
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
         FOCUSR          =   0   'False
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "Form7.frx":047A
         UMCOL           =   0   'False
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
      End
      Begin OsenXPCntrl.Command Cmd 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   6
         Left            =   3840
         TabIndex        =   55
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&About"
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
         FOCUSR          =   0   'False
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "Form7.frx":0496
         UMCOL           =   0   'False
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
      End
      Begin OsenXPCntrl.Command Cmd 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   7
         Left            =   2760
         TabIndex        =   54
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&Exit"
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
         FOCUSR          =   0   'False
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "Form7.frx":04B2
         UMCOL           =   0   'False
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Directory"
         Height          =   1095
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   8895
         Begin OsenXPCntrl.Command Cmd 
            CausesValidation=   0   'False
            Height          =   300
            Index           =   2
            Left            =   7920
            TabIndex        =   60
            Top             =   720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "Br&ows"
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
            MICON           =   "Form7.frx":04CE
            UMCOL           =   0   'False
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   2
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin OsenXPCntrl.Command Cmd 
            CausesValidation=   0   'False
            Height          =   300
            Index           =   1
            Left            =   7920
            TabIndex        =   59
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            BTYPE           =   3
            TX              =   "&Brows"
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
            MICON           =   "Form7.frx":04EA
            UMCOL           =   0   'False
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   2
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.PictureBox cainButton1 
            Height          =   135
            Left            =   840
            ScaleHeight     =   75
            ScaleWidth      =   555
            TabIndex        =   53
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox CText 
            Height          =   285
            Left            =   0
            LinkItem        =   "SerTxt"
            LinkTopic       =   "Player|Form1"
            TabIndex        =   52
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   0
            Top             =   720
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1200
            TabIndex        =   6
            Top             =   720
            Width           =   5535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form7.frx":0506
            Left            =   6840
            List            =   "Form7.frx":052B
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   240
            Width           =   6615
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OutFileName"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "| FileType |"
            Height          =   255
            Left            =   6840
            TabIndex        =   8
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "InputFileName"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   135
         Left            =   0
         TabIndex        =   2
         Top             =   1080
         Width           =   8895
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   30
         TabIndex        =   1
         Top             =   1560
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Max             =   1000
         Scrolling       =   1
      End
      Begin VB.Label zASD 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   15
         TabIndex        =   10
         Top             =   1350
         Width           =   2640
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   0
      TabIndex        =   14
      Top             =   1560
      Width           =   9255
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   375
         Left            =   3360
         TabIndex        =   40
         Top             =   6550
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   700
         Left            =   4560
         ScaleHeight     =   645
         ScaleWidth      =   195
         TabIndex        =   70
         Top             =   3720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   6255
         Left            =   0
         TabIndex        =   48
         Top             =   120
         Visible         =   0   'False
         Width           =   8775
         Begin VB.Label F 
            Height          =   135
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
            Height          =   960
            Left            =   0
            TabIndex        =   49
            Top             =   360
            Width           =   8775
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
            _cx             =   15478
            _cy             =   1693
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Back /\/\/\"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6960
            TabIndex        =   50
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H80000010&
         Caption         =   "Delet"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7800
         MaskColor       =   &H00000000&
         TabIndex        =   44
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H80000010&
         Caption         =   "Copy"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6840
         MaskColor       =   &H00000000&
         TabIndex        =   43
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H80000010&
         Caption         =   "Cut"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5880
         MaskColor       =   &H00000000&
         TabIndex        =   42
         Top             =   720
         Width           =   855
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H80000012&
         Caption         =   "InputFile Directory"
         ForeColor       =   &H80000018&
         Height          =   2655
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   4335
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   3240
            TabIndex        =   71
            Text            =   "*.*"
            Top             =   2280
            Width           =   855
         End
         Begin VB.DirListBox Dir1 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFC0&
            Height          =   1890
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   1815
         End
         Begin VB.DriveListBox Drive1 
            BackColor       =   &H00404040&
            ForeColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   4095
         End
         Begin VB.FileListBox File1 
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFC0&
            Height          =   1650
            Hidden          =   -1  'True
            Left            =   2040
            TabIndex        =   36
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label12 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TypeFile"
            ForeColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   2040
            TabIndex        =   72
            Top             =   2280
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H80000012&
         Height          =   1215
         Left            =   4800
         TabIndex        =   28
         Top             =   3720
         Width           =   3735
         Begin VB.TextBox Text4 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   240
            TabIndex        =   31
            Text            =   "File"
            Top             =   840
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000010&
            Caption         =   "MainName"
            Height          =   255
            Left            =   2160
            TabIndex        =   30
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000010&
            Caption         =   "Custom"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Output FileName"
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000008&
         ForeColor       =   &H8000000E&
         Height          =   1215
         Left            =   4845
         TabIndex        =   22
         Top             =   2400
         Width           =   3735
         Begin VB.OptionButton Option4 
            BackColor       =   &H80000010&
            Caption         =   "Custom"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H80000010&
            Caption         =   "MainType"
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Top             =   480
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   23
            Text            =   ".AVI"
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "*."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   27
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OuyputFile Type"
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000012&
         Height          =   855
         Left            =   4845
         TabIndex        =   18
         Top             =   5040
         Width           =   3735
         Begin VB.OptionButton Option6 
            BackColor       =   &H80000010&
            Caption         =   "Custom File"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H80000010&
            Caption         =   " Directory Files"
            Height          =   255
            Left            =   2160
            TabIndex        =   19
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Input File Directory"
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000012&
         Height          =   975
         Left            =   4845
         TabIndex        =   15
         Top             =   1320
         Width           =   3735
         Begin OsenXPCntrl.Command Cmd 
            Height          =   255
            Index           =   15
            Left            =   3000
            TabIndex        =   69
            Top             =   150
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "...."
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
            MICON           =   "Form7.frx":057D
            UMCOL           =   -1  'True
            SOFT            =   -1  'True
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   3495
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Output Directory"
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000010&
         Caption         =   "OutputFile Directory"
         ForeColor       =   &H80000018&
         Height          =   2465
         Left            =   240
         TabIndex        =   33
         Top             =   3480
         Width           =   4335
         Begin OsenXPCntrl.Command Cmd 
            Height          =   315
            Index           =   9
            Left            =   3480
            TabIndex        =   68
            Top             =   2010
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Up /\"
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
            MICON           =   "Form7.frx":0599
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
            Height          =   315
            Index           =   10
            Left            =   2640
            TabIndex        =   67
            Top             =   2010
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Down \/"
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
            MICON           =   "Form7.frx":05B5
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
            Height          =   315
            Index           =   11
            Left            =   1920
            TabIndex        =   66
            Top             =   2010
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
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
            MICON           =   "Form7.frx":05D1
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
            Height          =   315
            Index           =   12
            Left            =   1080
            TabIndex        =   65
            Top             =   2010
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
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
            MPTR            =   0
            MICON           =   "Form7.frx":05ED
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
            Height          =   315
            Index           =   8
            Left            =   720
            TabIndex        =   64
            Top             =   2010
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
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
            MICON           =   "Form7.frx":0609
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
            Height          =   315
            Index           =   13
            Left            =   120
            TabIndex        =   63
            Top             =   2010
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
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
            MICON           =   "Form7.frx":0625
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
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFC0&
            Height          =   1620
            ItemData        =   "Form7.frx":0641
            Left            =   120
            List            =   "Form7.frx":0643
            TabIndex        =   34
            Top             =   240
            Width           =   4095
         End
      End
      Begin OsenXPCntrl.Command Cmd 
         Height          =   375
         Index           =   14
         Left            =   5760
         TabIndex        =   61
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Start "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
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
         MICON           =   "Form7.frx":0645
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
         Index           =   16
         Left            =   4800
         TabIndex        =   62
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "/\ &Back"
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
         MICON           =   "Form7.frx":0661
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C&lose"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   345
         Left            =   7920
         MouseIcon       =   "Form7.frx":067D
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label Help 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   0
         TabIndex        =   39
         Top             =   6550
         Width           =   3375
      End
      Begin VB.Image Image2 
         Height          =   735
         Left            =   0
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   9000
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   0
         Stretch         =   -1  'True
         Top             =   120
         Width           =   8775
      End
      Begin VB.Label BNM 
         Caption         =   "Label11"
         Height          =   135
         Left            =   120
         TabIndex        =   47
         Top             =   6120
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Label FG 
      Height          =   135
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
      Height          =   135
      Left            =   -240
      TabIndex        =   12
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Color 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancel As Boolean
Dim DFA As String, D As String, I%
Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 1:  '-------------------------------
        CommonDialog1.DialogTitle = "Open File"
        CommonDialog1.Filter = "All Media[Audio & Video]|*.Wav;*.Mp3;*.Mid;*.Wma;*.Avi;*.Mpeg;*.Mpg;*.dat;*.ifo;*.Wmv;*.Mov|All Files|*.*"
        CommonDialog1.ShowOpen '-----------------------------------------------
        If CommonDialog1.FileName <> "" And CommonDialog1.FileName <> Text1.Text Then
        Text1.Text = CommonDialog1.FileName
        End If '-----------------------------------------
Case 2: CommonDialog1.DialogTitle = "Save File": CommonDialog1.ShowSave
        If CommonDialog1.FileName <> "" And CommonDialog1.FileName <> Text2.Text Then
        Text2.Text = Left$(CommonDialog1.FileName, Len(CommonDialog1.FileName) - 4)
        End If
Case 3: Call Coping(Form7)  '------------------------------------------------------------
Case 4: Label11_Click
        Me.Height = 7200: Frame1.Visible = False: Frame2.Visible = False: Frame3.Visible = False: Frame10.Visible = True: Frame10.Top = -150
Case 5: '---------------------------------------------------------------------
        If Text2.Text = Empty Or Combo1.Text = Empty Then Exit Sub
        Can.Visible = False
        Frame11.Visible = True: DFA = Text2.Text + "." + Right(Combo1.Text, 3)
        Call VideoMod(Form7, WindowsMediaPlayer1, F)
            If F.Caption = "Video" Then '-------------------------------------
            WindowsMediaPlayer1.Height = 5895: Me.Height = 8415: WindowsMediaPlayer1.URL = Text2.Text + "." + Right(Combo1.Text, 3)
            Else: WindowsMediaPlayer1.Height = 960: Me.Height = 3390: WindowsMediaPlayer1.URL = Text2.Text + "." + Right(Combo1.Text, 3)
            End If
Case 6: Credit '---------------------------------------------------
Case 7: Unload Me
Case 8: Dim T As Integer '-----------------------------------------------
        Can.Visible = True
        For T = 0 To File1.ListCount - 1
        If Cancel = True Then Exit For
        File1.Selected(T) = True: List1.AddItem (Dir1.Path + "\" + File1.FileName): DoEvents
        Next T '------------------------------------------------------------
        Cancel = False: Can.Visible = False
Case 9: Call LChange(Form7, List1, True)
Case 10: Call LChange(Form7, List1, False)
Case 11: List1.Clear
Case 12: '----------------------------------------------------------
        If List1.Text <> "" Then
           Help.Caption = "": Help.BackColor = &H8000000F: Color.BackColor = Help.BackColor: Color.Caption = Help.Caption: List1.RemoveItem (List1.ListIndex)
        Else: Help.Caption = "File Not Found!": Help.BackColor = RGB(250, 250, 0): Color.BackColor = Help.BackColor: Color.Caption = Help.Caption
        End If
Case 13: '---------------------------------------------------
        If File1.FileName <> "" And Dir1.Path + "\" + File1.FileName <> FG.Caption Then
           List1.AddItem (Dir1.Path + "\" + File1.FileName): FG.Caption = File1.FileName: Help.BackColor = &H8000000F: Help.Caption = "": Color.BackColor = Help.BackColor: Color.Caption = Help.Caption
        Else: Help.Caption = "File Not Found Or File Already Haved in List": Help.BackColor = RGB(250, 250, 0): Color.BackColor = Help.BackColor: Color.Caption = Help.Caption
        End If
Case 14: '--------------------------------------------
        If Option7.Value = False And Option8.Value = False _
        And Option9.Value = False Then MsgBox ("Plase Select Proses Mode!"): Exit Sub
        '{----------------------------------------------------------------------------}
        If Option7.Value = True Then
        Call Pres("Cut", Me)
        List1.Clear '-------------------------------
        ElseIf Option8.Value = True Then
        Call Pres("Copy", Me)
        ElseIf Option9.Value = True Then
        Call Pres("Delet", Me)
        List1.Clear
        End If '-------------------------------------
Case 15: Form8.Show: Me.Enabled = False
Case 16: Me.Height = 2325: Frame1.Visible = True: Frame2.Visible = True: Frame3.Visible = True: Frame10.Top = 1680
End Select '--------------------------------------------------------------->>>>>>>>>>>
End Sub

Private Sub Cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 3 Then ProsesFile
End Sub

Private Sub Cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 1: zASD.Caption = "To Botton Click LoadFile For InputFile [Alt+B]": zASD.BackColor = RGB(0, 255, 0)
Case 2: zASD.Caption = "To Botton Click LoadFile For OutputFile [Alt+O]": zASD.BackColor = RGB(0, 255, 0)
Case 3: zASD.Caption = "To Botton Click Convert InputFile To OutputFile [Alt+C]": zASD.BackColor = RGB(0, 255, 200)
Case 4: zASD.Caption = "To Botton Click Converting Many File To Onesecond [Alt+W]": zASD.BackColor = RGB(100, 200, 0)
Case 5: zASD.Caption = "To Botton Click Viwe OutputFile If Completed ! [Alt+P]": zASD.BackColor = RGB(255, 255, 255)
Case 6: zASD.Caption = "To Botton Click Viwe Information For Program [Alt+A]": zASD.BackColor = RGB(0, 255, 0)
Case 7: zASD.Caption = "To Botton Click Exit This Program [Alt+E] ": zASD.BackColor = RGB(255, 100, 100)
End Select
End Sub

Private Sub CText_Change()
Select Case (CText.Text)
Case "Min":     Me.WindowState = 1: Me.Show
Case "Exit":    Unload Me
Case "Convert": Cmd_Click (16): Me.Show
Case "Copy":    Cmd_Click (4): Me.Show
Case "Normal":  Me.WindowState = 0
End Select
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo A
Dir1.Path = Drive1.Drive
Exit Sub '-------------------------------
A: Drive1.Drive = "C:"
End Sub

Private Sub File1_DblClick()
Cmd_Click (13)
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 4 Then Cmd_Click (8)
If Button = vbRightButton Then Cmd_Click (13)
End Sub
Private Sub Cam_Click()
Cancel = True
End Sub

Private Sub Can_Click()
Cancel = True
End Sub

Private Sub cText_LinkClose()
DDE = False
End Sub

Private Sub cText_LinkOpen(Cancel As Integer)
DDE = True
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ProsesFile
End Sub
Private Sub Frame10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &HFF&
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ProsesFile
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &HFF&
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ProsesFile
End Sub

Private Sub Label10_Click()
Unload Me
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &HC0C0FF
End Sub

Private Sub Label11_Click()
Can.Visible = False
Frame11.Visible = False
WindowsMediaPlayer1.URL = "": Me.Height = 2325
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ProsesFile '--------------------------------------------
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 4 Then Cmd_Click (11)
If Button = 2 Then Cmd_Click (12)
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
List1.ToolTipText = List1.Text
End Sub

Private Sub Option1_Click()
Text4.Enabled = False
End Sub

Private Sub Option2_Click()
Text4.Enabled = True
End Sub

Private Sub Option3_Click()
Text3.Enabled = False
End Sub

Private Sub Option4_Click()
Text3.Enabled = True
End Sub

Private Sub Option5_Click()
Frame4.Enabled = False: List1.Enabled = False: Frame7.BackColor = &H80000010: Frame4.BackColor = 0
End Sub

Private Sub Option6_Click()
Frame4.Enabled = True: List1.Enabled = True: Frame4.BackColor = &H80000010: Frame7.BackColor = 0
End Sub

Private Sub Option7_Click()
Sho
End Sub

Private Sub Option8_Click()
Sho
End Sub
Sub Sho()
Frame9.Visible = True: Frame6.Visible = True: Frame8.Visible = True
End Sub
Private Sub Option9_Click()
On Error Resume Next
Frame9.Visible = False: Frame6.Visible = False: Frame8.Visible = False
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ProsesFile '----------------------------------------------------------------
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ProsesFile '-------------------------------------------------
End Sub
Private Sub Form_Load()
On Error GoTo c
 If App.PrevInstance = True Then End
zASD.Caption = "Welcom To GhayeshRayaneh Converter": zASD.BackColor = RGB(0, 255, 255)
For I = 1 To 16 '-----------------------------------------
Cmd(I).Refresh: Next
Tyf Picture1, "With_Black", Image1
Tyf Picture1, "Black_With", Image2
Me.Height = 2325: Seting Me '---------------------------
CText.LinkMode = 1
CText.Text = "Load"
CText.LinkPoke: DDE = True
CText_Change
Exit Sub '---------------------------
c: DDE = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim I As Integer
If DDE = True Then
CText = "Close": CText.LinkPoke
End If '-------------------------------
For I = 0 To 50: Me.Caption = I: Next
End Sub

Private Sub Text6_Change()
On Error GoTo D
File1.Pattern = Text6.Text: Exit Sub
D: File1.Pattern = "*.*"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next '-------------------------------
    If ProgressBar1.Value = 1000 Then
    ProgressBar1.Value = 0: ProgressBar1.Visible = False: Timer1.Enabled = False: Frame3.Enabled = True
    Exit Sub
    End If '----------------------------------------
Frame3.Enabled = False: ProgressBar1.Visible = True: ProgressBar1.Value = ProgressBar1.Value + 10
zASD.Caption = "Plase With": zASD.BackColor = &H8000000F
End Sub

Public Sub Coping(Frm As Form7)
'On Error Resume Next
Dim c As Integer '--------------------------------------------
With Frm
.CText.Enabled = False
If ASd.FileExists(.Text1.Text) = True Then
    If .Combo1.Text <> Empty And .Combo1.Text <> "______" Then
        If .Text2.Text = "" Then '-------------------------------
        c = MsgBox("Your Output File Is Invalid ,Do you Want To SelectFile For Output?", vbYesNo)
            If c = vbYes Then Cmd_Click (2)
                If c = vbNo Then Exit Sub
                    End If '---------------------------------------
                        If ASd.FileExists(.Text2.Text + (Right$(.Combo1.Text, Len(.Combo1.Text) - 1))) = False Then
12
Call ASd.CopyFile(.Text1.Text, .Text2.Text + (Right$(.Combo1.Text, Len(.Combo1.Text) - 1)))
                                .Timer1.Enabled = True '-------------------------------
                                Else
                                c = MsgBox("Your Output File Have In Yor Drive , Doyou Want To ReWeith?", vbYesNo + vbExclamation)
                            If c = vbYes Then GoTo 12
                        End If '-------------------------------------------------------
                    Else
                c = MsgBox("Your Output FileType Is Invalid " & "Plase ReSelect OutputFile Type Of Combobox", vbExclamation)
                End If
            Else '-----------------------------------------------------------------------
        c = MsgBox("Your InputFile Is Invalid , Doyou Want To ReSelect?", vbYesNo + vbExclamation)
    If c = vbYes Then Cmd_Click (1)
End If '---------------------------------------------------------------------------------
.CText.Enabled = True
End With
End Sub

