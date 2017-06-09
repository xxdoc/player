VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8AB3C700-28FF-476A-852B-5E1558F44746}#1.0#0"; "command.ocx"
Begin VB.Form FLoad 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Sink "
   ClientHeight    =   6690
   ClientLeft      =   3750
   ClientTop       =   4410
   ClientWidth     =   9405
   LinkMode        =   1  'Source
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin OsenXPCntrl.Command Command1 
      Height          =   300
      Left            =   8280
      TabIndex        =   10
      Top             =   6240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
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
      MICON           =   "FLoad.frx":0000
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
      Default         =   -1  'True
      Height          =   300
      Left            =   7320
      TabIndex        =   9
      Top             =   6240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
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
      MICON           =   "FLoad.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.Command Command3 
      Height          =   300
      Left            =   6360
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "Apply"
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
      MICON           =   "FLoad.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000A&
      Height          =   6255
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Visible         =   0   'False
      Width           =   9495
      Begin VB.Frame Frame12 
         Height          =   975
         Left            =   5040
         TabIndex        =   49
         Top             =   5280
         Width           =   4095
         Begin VB.TextBox Text2 
            BackColor       =   &H00000000&
            ForeColor       =   &H80000014&
            Height          =   285
            Left            =   2760
            TabIndex        =   53
            Text            =   "1"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H8000000E&
            ForeColor       =   &H80000003&
            Height          =   285
            Left            =   2760
            TabIndex        =   50
            Text            =   "5"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Count Media"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "PlayVideo Repeat Numer"
            Height          =   255
            Left            =   600
            TabIndex        =   52
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "PlayAudio Repeat Numer"
            Height          =   255
            Left            =   600
            TabIndex        =   51
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Color"
         Height          =   615
         Left            =   240
         TabIndex        =   46
         Top             =   5520
         Width           =   1815
         Begin VB.Label Label5 
            BackColor       =   &H8000000D&
            Caption         =   " "
            Height          =   255
            Left            =   1320
            TabIndex        =   48
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000D&
            Caption         =   "Program Color"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Button Cintrol"
         Height          =   5175
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   8895
         Begin VB.Frame Frame9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Button Type"
            Height          =   1215
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   8535
            Begin OsenXPCntrl.Command Cm 
               Height          =   135
               Index           =   10
               Left            =   7800
               TabIndex        =   63
               Top             =   720
               Visible         =   0   'False
               Width           =   135
               _ExtentX        =   238
               _ExtentY        =   238
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
               MICON           =   "FLoad.frx":0054
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               BTYPE           =   1
               TX              =   "Window 16-Bit"
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
               MICON           =   "FLoad.frx":0070
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   2
               Left            =   1800
               TabIndex        =   36
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               BTYPE           =   2
               TX              =   "Windows 32-Bit"
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
               MICON           =   "FLoad.frx":008C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   3
               Left            =   3360
               TabIndex        =   37
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Windows XP"
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
               MICON           =   "FLoad.frx":00A8
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   4
               Left            =   4680
               TabIndex        =   38
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               BTYPE           =   4
               TX              =   "Mac"
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
               MICON           =   "FLoad.frx":00C4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   5
               Left            =   5760
               TabIndex        =   39
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               BTYPE           =   5
               TX              =   "Java Metal"
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
               MICON           =   "FLoad.frx":00E0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   6
               Left            =   7200
               TabIndex        =   40
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BTYPE           =   6
               TX              =   "Netscape 6"
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
               MICON           =   "FLoad.frx":00FC
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   41
               Top             =   720
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
               BTYPE           =   7
               TX              =   "Simple Flat"
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
               MICON           =   "FLoad.frx":0118
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   8
               Left            =   1320
               TabIndex        =   42
               Top             =   720
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               BTYPE           =   8
               TX              =   "Flat Highlight"
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
               MICON           =   "FLoad.frx":0134
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   9
               Left            =   2760
               TabIndex        =   55
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               BTYPE           =   9
               TX              =   "Offic XP"
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
               MICON           =   "FLoad.frx":0150
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   11
               Left            =   3840
               TabIndex        =   56
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BTYPE           =   11
               TX              =   "TransParent"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   0   'False
               BCOL            =   14737632
               BCOLO           =   14737632
               FCOL            =   14737632
               FCOLO           =   14737632
               MCOL            =   12632256
               MPTR            =   0
               MICON           =   "FLoad.frx":016C
               UMCOL           =   0   'False
               SOFT            =   -1  'True
               PICPOS          =   0
               NGREY           =   -1  'True
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   12
               Left            =   5160
               TabIndex        =   57
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
               BTYPE           =   12
               TX              =   "3D Hover"
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
               MICON           =   "FLoad.frx":0188
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Cm 
               Height          =   375
               Index           =   13
               Left            =   6240
               TabIndex        =   58
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BTYPE           =   13
               TX              =   "Oval Flat"
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
               MICON           =   "FLoad.frx":01A4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Typ 
               Caption         =   " 3"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   600
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Other"
            Height          =   975
            Left            =   120
            TabIndex        =   27
            Top             =   3960
            Width           =   8655
            Begin OsenXPCntrl.Command Command4 
               Height          =   375
               Left            =   7320
               TabIndex        =   33
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Preview"
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
               MICON           =   "FLoad.frx":01C0
               UMCOL           =   -1  'True
               SOFT            =   -1  'True
               PICPOS          =   0
               NGREY           =   -1  'True
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.CheckBox Chk5 
               BackColor       =   &H00E0E0E0&
               Caption         =   "UseGreyscal"
               Height          =   375
               Left            =   5880
               TabIndex        =   32
               Top             =   360
               Width           =   1335
            End
            Begin VB.CheckBox Chk4 
               BackColor       =   &H00E0E0E0&
               Caption         =   "SoftBevel"
               Height          =   375
               Left            =   4680
               TabIndex        =   31
               Top             =   360
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox Chk3 
               BackColor       =   &H00E0E0E0&
               Caption         =   "ShowFocusRect"
               Height          =   375
               Left            =   3000
               TabIndex        =   30
               Top             =   360
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox Chk2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Behaviour"
               Height          =   375
               Left            =   1800
               TabIndex        =   29
               Top             =   360
               Width           =   1095
            End
            Begin VB.CheckBox Chk1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "CausesValidation"
               Height          =   375
               Left            =   120
               TabIndex        =   28
               Top             =   360
               Value           =   1  'Checked
               Width           =   1575
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Spical Efect"
            Height          =   855
            Left            =   120
            TabIndex        =   22
            Top             =   2880
            Width           =   8655
            Begin OsenXPCntrl.Command Co 
               Height          =   375
               Index           =   4
               Left            =   6480
               TabIndex        =   26
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Shadowed"
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
               MICON           =   "FLoad.frx":01DC
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   3
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Co 
               Height          =   375
               Index           =   3
               Left            =   4320
               TabIndex        =   25
               Top             =   240
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Engraved"
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
               MICON           =   "FLoad.frx":01F8
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   2
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Co 
               Height          =   375
               Index           =   2
               Left            =   2280
               TabIndex        =   24
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Embossed"
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
               MICON           =   "FLoad.frx":0214
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   1
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Co 
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   23
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "None"
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
               MICON           =   "FLoad.frx":0230
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Efect 
               Caption         =   "1"
               Height          =   135
               Left            =   360
               TabIndex        =   44
               Top             =   600
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Color Scheme"
            Height          =   975
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   8655
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   8280
               Top             =   1320
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin OsenXPCntrl.Command Com 
               Height          =   375
               Index           =   4
               Left            =   6480
               TabIndex        =   16
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Use Container"
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
               COLTYPE         =   4
               FOCUSR          =   -1  'True
               BCOL            =   16777215
               BCOLO           =   16777215
               FCOL            =   0
               FCOLO           =   16711680
               MCOL            =   12632256
               MPTR            =   0
               MICON           =   "FLoad.frx":024C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Com 
               Height          =   375
               Index           =   3
               Left            =   4320
               TabIndex        =   15
               Top             =   240
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Force Standard"
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
               COLTYPE         =   3
               FOCUSR          =   -1  'True
               BCOL            =   16777215
               BCOLO           =   16777215
               FCOL            =   0
               FCOLO           =   16711680
               MCOL            =   12632256
               MPTR            =   0
               MICON           =   "FLoad.frx":0268
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Com 
               Height          =   375
               Index           =   2
               Left            =   2280
               TabIndex        =   14
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Custom"
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
               COLTYPE         =   2
               FOCUSR          =   -1  'True
               BCOL            =   8421631
               BCOLO           =   16777152
               FCOL            =   8421376
               FCOLO           =   16711680
               MCOL            =   16761024
               MPTR            =   0
               MICON           =   "FLoad.frx":0284
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin OsenXPCntrl.Command Com 
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   13
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Use Windows"
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
               MICON           =   "FLoad.frx":02A0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Color 
               Caption         =   "1"
               Height          =   255
               Left            =   360
               TabIndex        =   43
               Top             =   600
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label Lab 
               BackColor       =   &H00FFC0C0&
               Caption         =   " "
               Height          =   255
               Index           =   5
               Left            =   3840
               TabIndex        =   21
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Lab 
               BackColor       =   &H00FF0000&
               Caption         =   " "
               Height          =   255
               Index           =   4
               Left            =   3480
               TabIndex        =   20
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Lab 
               BackColor       =   &H00808000&
               Caption         =   " "
               Height          =   255
               Index           =   3
               Left            =   3120
               TabIndex        =   19
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Lab 
               BackColor       =   &H00FFFFC0&
               Caption         =   " "
               Height          =   255
               Index           =   2
               Left            =   2760
               TabIndex        =   18
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Lab 
               BackColor       =   &H00C0C0FF&
               Caption         =   " "
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   17
               Top             =   600
               Width           =   255
            End
         End
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5880
      TabIndex        =   1
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   6375
      Left            =   -120
      TabIndex        =   0
      Top             =   -240
      Width           =   9615
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Professional Sink"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   360
         TabIndex        =   5
         Top             =   5640
         Width           =   3855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Standard Sink"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   5640
         TabIndex        =   4
         Top             =   5640
         Width           =   2895
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Professional Sink"
         Height          =   5175
         Left            =   240
         MouseIcon       =   "FLoad.frx":02BC
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   360
         Width           =   4575
         Begin VB.PictureBox Image4 
            Height          =   3255
            Left            =   120
            ScaleHeight     =   3195
            ScaleWidth      =   4275
            TabIndex        =   59
            Top             =   1800
            Width           =   4335
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "The Sink is Up The Siftwar Mode .And User Frindly Aplication Hase Mode .The Profeshnal Sink is Update Mode Enjoy This Mode."
            Height          =   975
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Standard Sink"
         Height          =   5175
         Left            =   5040
         MouseIcon       =   "FLoad.frx":05C6
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   360
         Width           =   4335
         Begin VB.PictureBox Image3 
            Height          =   2775
            Left            =   120
            ScaleHeight     =   2715
            ScaleWidth      =   4035
            TabIndex        =   60
            Top             =   2040
            Width           =   4095
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "The Sink is Down The Siftwar Mode .And User Frindly Aplication Hase Mode .The Standard Sink is Template Mode Enjoy This Mode."
            Height          =   975
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   4095
         End
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright:GHAYESH RAYANEH All rights Reserved ,2005-2006   By:NasserNiazyMobasser"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   6120
      Width           =   4935
   End
End
Attribute VB_Name = "FLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DfA As String

Private Sub Chk1_Click()
Command4.CausesValidation = Chk1.Value
End Sub

Private Sub Chk2_Click()
Command4.CheckBoxBehaviour = Chk2.Value
End Sub

Private Sub Chk3_Click()
Command4.ShowFocusRect = Chk3.Value
End Sub

Private Sub Chk4_Click()
Command4.SoftBevel = Chk4.Value
End Sub

Private Sub Chk5_Click()
Command4.UseGreyscale = Chk5.Value
End Sub

Private Sub Cm_Click(Index As Integer)
For i = Cm.LBound To Cm.UBound
Cm(i).Enabled = True
Next
Cm(Index).Enabled = False: Typ.Caption = Index
End Sub

Private Sub Co_Click(Index As Integer)
For i = Co.LBound To Co.UBound
Co(i).Enabled = True
Next
Co(Index).Enabled = False: Efect.Caption = Index
End Sub

Private Sub Com_Click(Index As Integer)
For i = Com.LBound To Com.UBound
Com(i).Enabled = True
Next
Com(Index).Enabled = False: Color.Caption = Index
End Sub

Private Sub Command1_Click()
On Error Resume Next
                    Unload Me
End Sub

Private Sub Command2_Click()
        On Error Resume Next
        Dim i As Integer, M As String
If Frame4.Visible = True Then
Command3_Click
FForm1.Show
Unload Me
Else
                If Option1.Value = True Then
                        Call FForm1.Show: Call SinkP(FForm1): Unload Me
                ElseIf Option2.Value = True Then
                        Call FForm1.Show: Call SinkS(FForm1): Unload Me
                Else
                        MsgBox ("Plase Select Sink Mode")
                End If
End If
End Sub

Private Sub Command3_Click()
If Val(Text3.Text) = 0 Then Text3.Text = 5
If Val(Text2.Text) = 0 Then Text3.Text = 1
With Setting
            .AudioCunt = Val(Text3.Text)
            .VideoCunt = Val(Text2.Text)
            .Causes = Chk1.Value
            .Behavi = Chk2.Value
            .Rect = Chk3.Value
            .Soft = Chk4.Value
            .Grey = Chk5.Value
            .ProgramColor = Label5.BackColor
            .Button.BColor = Val(Color.Caption)
            .Button.BEfect = Val(Efect.Caption)
            .Button.BType = Val(Typ.Caption)
             For i = 1 To 5
            .Button.Custom(i) = Lab(i).BackColor: Next
End With
Open App.Path + "\NMS\user.dll" For Random Access Write As #2
Put #2, , Setting
Close
           
Setng FForm1, 30, True, FForm1.WindowsMediaPlayer1, FForm1.List1, 12
End Sub

Private Sub Form_Load()
On Error Resume Next
Option1.Value = False: Option2.Value = False
Me.Left = Screen.Width \ 2 - (Me.Width \ 2): Me.Top = Screen.Height \ 2 - (Me.Height \ 2)
Image3.Picture = FForm1.ImageList2.ListImages(2).Picture
Image4.Picture = FForm1.ImageList2.ListImages(1).Picture
End Sub
Sub Lod()
If Setting.ProgramColor = 0 Then Exit Sub
With Setting
For i = 1 To 5
 Lab(i).BackColor = .Button.Custom(i): Next
      Typ.Caption = .Button.BType
    Color.Caption = .Button.BColor
    Efect.Caption = .Button.BEfect
       Chk1.Value = Val(.Causes)
       Chk2.Value = Val(.Behavi)
       Chk3.Value = Val(.Rect)
       Chk4.Value = Val(.Soft)
       Chk5.Value = Val(.Grey)
 Label5.BackColor = .ProgramColor
 Com(2).BackColor = .Button.Custom(1)
  Com(2).BackOver = .Button.Custom(2)
 Com(2).ForeColor = .Button.Custom(3)
  Com(2).ForeOver = .Button.Custom(4)
 Com(2).MaskColor = .Button.Custom(5)
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Option1.Enabled = False And Option2.Enabled = False And Text1.Text = "A" Then End
If Command1.Enabled = False Then End
If Text1.Text = "A" Then
Call ALoad(FForm1)
End If
FForm1.Show
End Sub

Private Sub Frame1_Click()
On Error Resume Next
      If Option1.Enabled = True Then
      Option1.Value = True
      End If
End Sub

Private Sub Frame1_DblClick()
Command2_Click
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
                    Frame1.BackColor = RGB(0, 250, 0)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Frame1.BackColor = &HFFFF80: Frame3.BackColor = &HFFFF80
End Sub

Private Sub Frame3_Click()
On Error Resume Next
  If Option2.Enabled = True Then
  Option2.Value = True
  End If
End Sub

Private Sub Frame3_DblClick()
Command2_Click
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
                    Frame3.BackColor = RGB(0, 250, 0)
End Sub

Private Sub Image1_Click()
On Error Resume Next
If Option1.Enabled = True Then
                    Option1.Value = True
End If
End Sub

Private Sub Image1_DblClick()
Command2_Click
End Sub

Private Sub Image2_Click()
On Error Resume Next
If Option2.Enabled = True Then
Option2.Value = True
End If
End Sub

Private Sub Image2_DblClick()
Command2_Click
End Sub

Private Sub Image3_Click()
On Error Resume Next
                    If Option2.Enabled = True Then
                    Option2.Value = True
                    End If
End Sub

Private Sub Image3_DblClick()
Command2_Click
End Sub

Private Sub Image4_Click()
On Error Resume Next
    If Option1.Enabled = True Then
    Option1.Value = True
    End If
End Sub

Private Sub Image4_DblClick()
Command2_Click
End Sub

Private Sub Lab_Click(Index As Integer)
CommonDialog1.ShowColor
Select Case Index
Case 1: Com(2).BackColor = CommonDialog1.Color
Case 2:  Com(2).BackOver = CommonDialog1.Color
Case 3: Com(2).ForeColor = CommonDialog1.Color
Case 4:  Com(2).ForeOver = CommonDialog1.Color
Case 5: Com(2).MaskColor = CommonDialog1.Color
End Select
Lab(Index).BackColor = CommonDialog1.Color
End Sub

Private Sub Label2_Click()
Credit
End Sub

Private Sub Label5_Click()
CommonDialog1.ShowColor
Label5.BackColor = CommonDialog1.Color
End Sub

Private Sub Option1_DblClick()
Image4_DblClick
End Sub

Private Sub Option2_DblClick()
Image3_DblClick
End Sub

