VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8AB3C700-28FF-476A-852B-5E1558F44746}#1.0#0"; "command.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcoom PicToAvi"
   ClientHeight    =   5610
   ClientLeft      =   3555
   ClientTop       =   2475
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6825
   Begin OsenXPCntrl.Command Cmd 
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   32
      Top             =   4750
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
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
      MICON           =   "Frmavi.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   5880
      ScaleHeight     =   795
      ScaleWidth      =   555
      TabIndex        =   25
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Frmavi.frx":001C
      Left            =   120
      List            =   "Frmavi.frx":006B
      TabIndex        =   24
      Top             =   4650
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar Pr 
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   4750
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4935
      Left            =   0
      TabIndex        =   4
      Top             =   -230
      Width           =   5415
      Begin VB.ListBox lstDIBList 
         Height          =   1815
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox SurTxt 
         Height          =   285
         Left            =   2640
         LinkItem        =   "SerTxt"
         LinkTopic       =   "Player|Form1"
         TabIndex        =   34
         Top             =   4560
         Visible         =   0   'False
         Width           =   735
      End
      Begin OsenXPCntrl.Command Cmd 
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   31
         Top             =   4560
         Width           =   375
         _ExtentX        =   661
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
         MPTR            =   0
         MICON           =   "Frmavi.frx":0169
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
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   4560
         Width           =   375
         _ExtentX        =   661
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
         MPTR            =   0
         MICON           =   "Frmavi.frx":0185
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
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   29
         Top             =   4560
         Width           =   615
         _ExtentX        =   1085
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
         MPTR            =   0
         MICON           =   "Frmavi.frx":01A1
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
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   28
         Top             =   4560
         Width           =   615
         _ExtentX        =   1085
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
         MPTR            =   0
         MICON           =   "Frmavi.frx":01BD
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
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   27
         Top             =   4560
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
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
         MICON           =   "Frmavi.frx":01D9
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
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   26
         Top             =   4560
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "MakeAvi"
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
         MICON           =   "Frmavi.frx":01F5
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   3960
         TabIndex        =   22
         Top             =   1080
         Width           =   1455
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   960
            TabIndex        =   35
            Text            =   "1"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   225
            Left            =   960
            TabIndex        =   23
            Text            =   "2"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label16 
            Caption         =   "Frame/Sec"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Stop inPic"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Time:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   2160
         Width           =   3015
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Hidden          =   -1  'True
         Left            =   3720
         Pattern         =   "*.bmp;*.jpg"
         System          =   -1  'True
         TabIndex        =   7
         Top             =   2520
         Width           =   1575
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   2160
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ListBox lstDIBList_1 
         Height          =   4155
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin OsenXPCntrl.Command Cmd 
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   33
         Top             =   4560
         Width           =   495
         _ExtentX        =   873
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
         MPTR            =   0
         MICON           =   "Frmavi.frx":0211
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
         Height          =   2055
         Left            =   3960
         TabIndex        =   15
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   850
         Width           =   1455
      End
      Begin VB.Image imgPreview 
         Height          =   1695
         Left            =   2140
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label11 
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         ToolTipText     =   "Start Search By Basic"
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00CAC6C4&
         Height          =   615
         Left            =   4000
         TabIndex        =   14
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label7 
         Height          =   345
         Left            =   4680
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   1110
         Width           =   1460
      End
      Begin VB.Label Label4 
         Caption         =   "Frame=          %=  "
         Height          =   495
         Left            =   3960
         TabIndex        =   10
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00CAC6C4&
         Height          =   615
         Left            =   4000
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   4320
         Width           =   735
      End
   End
   Begin VB.Label Label13 
      Caption         =   "<< Output Size"
      Height          =   375
      Left            =   1440
      TabIndex        =   21
      Top             =   4725
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   4320
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      Caption         =   "Pase Waith Saving Avi  ..."
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Option Base 0


Private Const RASTERCAPS    As Long = 38
Private Const RC_PALETTE    As Long = &H100
Private Const SIZEPALETTE   As Long = 104
Private Type Rect
   left   As Long
   top    As Long
   right  As Long
   bottom As Long
End Type
Dim I1      As Integer
Dim Er      As String
Dim Fil     As String * 255
Dim OutSiz  As String
Dim Indx_   As String
Dim Start   As Boolean
Dim H1      As Integer
Dim W1      As Integer
Dim con1    As Integer
Dim Tim     As Integer
Public del1 As FileSystemObject

Private Sub Cmd_Click(Index As Integer)
On Error GoTo C
Dim v%
If (Index > 2 And Index < 6) Then OutSiz = ""
Select Case Index
Case 1: OutSiz = "": If lstDIBList_1.Text = "" Then Exit Sub
        lstDIBList_1.RemoveItem lstDIBList_1.ListIndex
Case 2: v% = MsgBox("Copyright: GhayeshRayaneh All Right Reserved " & vbCrLf & "            2006 , By:NasserNiazyMobasser", vbInformation)
Case 3: lstDIBList_1.Clear: ASD.DeleteFolder App.Path + "\Temp": ASD.CreateFolder App.Path + "\Temp"
Case 4: If File1.filename = "" Then Exit Sub
        If Len(Dir1.Path) = 3 Then
        lstDIBList_1.AddItem Dir1.Path + File1.List(I1)
        Else: lstDIBList_1.AddItem Dir1.Path + "\" + File1.List(I1)
        End If
Case 5: LChange Me, lstDIBList_1, True
Case 6: LChange Me, lstDIBList_1, False
Case 7: Can = False: Save_bmp lstDIBList_1
Case 8: Can = True
        Rapair
End Select
C:
End Sub

Private Sub Combo1_Change()
Indx_ = Combo1.Text
    EditCustom
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Select Case Combo1.ListIndex 'By Twips : Number + 4 *15 =By Pixel is Standard
Case 0:  Picture1.Height = (320 + 4) * 15: Picture1.Width = (480 + 4) * 15 '480__320
Case 1:  Picture1.Height = (480 + 4) * 15: Picture1.Width = (640 + 4) * 15 '640__480
Case 2:  Picture1.Height = (600 + 4) * 15: Picture1.Width = (800 + 4) * 15 '800__600
Case 3:  Picture1.Height = (768 + 4) * 15: Picture1.Width = (1024 + 4) * 15 '1024_768
Case 4:  Picture1.Height = (90 + 4) * 15:  Picture1.Width = (120 + 4) * 15 '120__90
Case 5:  Picture1.Height = (160 + 4) * 15: Picture1.Width = (120 + 4) * 15 '120__160
Case 6:  Picture1.Height = (120 + 4) * 15: Picture1.Width = (176 + 4) * 15 '176__120
Case 7:  Picture1.Height = (144 + 4) * 15: Picture1.Width = (176 + 4) * 15 '176__144
Case 8:  Picture1.Height = (144 + 4) * 15: Picture1.Width = (192 + 4) * 15 '192__144
Case 9:  Picture1.Height = (180 + 4) * 15: Picture1.Width = (240 + 4) * 15 '240 180
Case 10: Picture1.Height = (200 + 4) * 15: Picture1.Width = (320 + 4) * 15 '320 200
Case 11: Picture1.Height = (240 + 4) * 15: Picture1.Width = (320 + 4) * 15 '320 240
Case 12: Picture1.Height = (240 + 4) * 15: Picture1.Width = (352 + 4) * 15 '352 240
Case 13: Picture1.Height = (288 + 4) * 15: Picture1.Width = (352 + 4) * 15 '352 288
Case 14: Picture1.Height = (488 + 4) * 15: Picture1.Width = (352 + 4) * 15 '352 488
Case 15: Picture1.Height = (576 + 4) * 15: Picture1.Width = (352 + 4) * 15 '352 576
Case 16: Picture1.Height = (288 + 4) * 15: Picture1.Width = (384 + 4) * 15 '384 288
Case 17: Picture1.Height = (480 + 4) * 15: Picture1.Width = (480 + 4) * 15 '480 480
Case 18: Picture1.Height = (576 + 4) * 15: Picture1.Width = (480 + 4) * 15 '480 576
Case 19: Picture1.Height = (576 + 4) * 15: Picture1.Width = (640 + 4) * 15 '640 576
Case 20: Picture1.Height = (480 + 4) * 15: Picture1.Width = (704 + 4) * 15 '704 480
Case 21: Picture1.Height = (576 + 4) * 15: Picture1.Width = (704 + 4) * 15 '704 576
Case 22: Picture1.Height = (480 + 4) * 15: Picture1.Width = (720 + 4) * 15 '720 480
Case 23: Picture1.Height = (576 + 4) * 15: Picture1.Width = (720 + 4) * 15 '720 576
Case 24: Picture1.Height = (576 + 4) * 15: Picture1.Width = (768 + 4) * 15 '768 576
End Select
Indx_ = Combo1.Text
    EditCustom
 Picture1.Refresh
End Sub
Private Sub EditCustom()
Dim A2$
A2 = Indx_
    If (lstDIBList_1.ListCount = 0) Or (A2 = "") Then
    Cmd(7).Enabled = False
    Else: Cmd(7).Enabled = True
    End If
'------------------------------Proses Custome Size----------------------------------
If Num(Nasa(A2, "_", True)) = False Then GoTo Y
If Num(Nasa(A2, "_", False)) = False Then GoTo Y
If Mid(A2, Len(Nasa(A2, "_", True)) + 1, 2) <> "__" Then GoTo Y
1:  Label13.Caption = "Size (" + Nasa(A2, "_", True) _
    + "x" + Nasa(A2, "_", False) + ")" '------------------------------------------------
    Picture1.Height = (Val(Nasa(A2, "_", False)) + 4) * 15: _
    Picture1.Width = ((Val(Nasa(A2, "_", True))) + 4) * 15: Picture1.Refresh: Exit Sub
Y: '-------------------------------Error Typing CustomeSize-------------------------------
Label13.Caption = "Error Typing 'Width__Height' ": Cmd(7).Enabled = False
If OutSiz <> Combo1.Text Then
OutSiz = "": lstDIBList.Clear
End If
End Sub

Private Sub Save_bmp(LST As ListBox)
On Error Resume Next
Dim S1 As Integer, J%, TER As Double
'Combo1_Change
If LST.ListCount = 0 Then Exit Sub
If ASD.FolderExists(App.Path + "\Temp") = True Then GoTo 2
ASD.CreateFolder App.Path + "\Temp" '------------------------
2   Pr.Visible = True: Pr.Max = LST.ListCount: Label3.Caption = "Proscec Your Selected Files And Saving To Temp..."
    Frame2.Enabled = False: Cmd(8).Visible = True: Label13.Visible = False: Frame1.Visible = False
    Label8.Caption = LST.ListCount: Label10.Visible = False: Combo1.Visible = False
    Label6.Caption = "Plase Waith Saving File ...": Label7.Caption = "0": Label6.Visible = True
    If OutSiz = Label13.Caption Then GoTo Skip
3 For I1 = 0 To LST.ListCount - 1 '----------Start Save To Bitmap------
 If Can = True Then GoTo D
   LST.Selected(I1) = True: Label9.Caption = left(LST.List(I1), 14) + "... " + right(LST.List(I1), 14)
    Label7.Caption = Mid((Str((I1 * 10) / Val(LST.ListCount / 10))), 2, 5) + " %"
    Label5.Caption = "Frame=" + Str(I1) + "/" + Str(Label8.Caption): Pr.Value = I1: Label12.Caption = TTim(Pr.Value, LST.ListCount)
    With Picture1
    .Picture = LoadPicture(LST.List(I1))
                    .PaintPicture .Picture, 0, 0, _
                    .Width, .Height
                     Set .Picture = .Image: .Refresh
    End With
    SavePicture Picture1.Picture, App.Path + "\Temp\" + Str(I1) + ".Bmp"
    DoEvents
 Next
Skip:
lstDIBList.Clear '---------------Prosesin Stop In Pictuer Per Secund---------------------
S1 = Val(Text1): If S1 = 0 Then S1 = 2
    For I1 = 0 To LST.ListCount - 1
    For J = 0 To S1 - 1
        Er$ = (App.Path + "\Temp\" + Str(I1) + ".Bmp")
        lstDIBList.AddItem (Er$)
    Next J: Next I1
OutSiz = Label13.Caption: Label11.Caption = ""
make_avi lstDIBList
Exit Sub
D:
Label13.Visible = True: Frame1.Visible = True: Pr.Visible = False
Frame2.Enabled = True: Cmd(8).Visible = False
Label8.Caption = lstDIBList.ListCount: Label10.Visible = True
Label6.Visible = False: Combo1.Visible = True
End Sub

Private Sub Dir1_Change()
On Error GoTo A:
File1.Path = Dir1.Path
A:
End Sub

Private Sub Dir1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
If Dir1.ListIndex < 1 Then Exit Sub
Dir1.ToolTipText = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
DriveChange Me, Drive1, Dir1
End Sub

Private Sub File1_Click()
On Error GoTo A
imgPreview.Picture = LoadPicture(Dir1.Path + "\" + File1.filename)
Image1.Picture = LoadPicture(Dir1.Path + "\" + File1.filename)
Label11.Caption = "W= " + left(Str(Image1.Width / 15), 5) + "    H= " + left(Str(Image1.Height / 15), 5)
A:
End Sub

Private Sub File1_DblClick()
Cmd_Click (4)
End Sub

Private Sub File1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
If button = 4 Then
For I1 = 0 To File1.ListCount - 1
    If Len(Dir1.Path) = 3 Then
    lstDIBList_1.AddItem Dir1.Path + File1.List(I1)
    Else: lstDIBList_1.AddItem Dir1.Path + "\" + File1.List(I1)
    End If
Next
 OutSiz = ""
ElseIf button = 2 Then
Cmd_Click (4)
End If
End Sub

Private Sub File1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo A
File1.ToolTipText = File1.List(File1.ListIndex)
A: ProsesTime
End Sub



Private Sub Form_Load()
On Error GoTo 2
 If App.PrevInstance = True Then End
SurTxt.LinkMode = 1: DDE = True
SurTxt.Text = "Avi": SurTxt.LinkPoke
2
    If ASD.FolderExists(App.Path + "/Temp") = False Then
    ASD.CreateFolder App.Path + "/Temp"
    Else: Rapair
    End If
Me.Width = 5475: Me.Height = 5475
Seting
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo 2
SurTxt.Text = "Unload": ASD.DeleteFolder App.Path + "\temp": ASD.CreateFolder App.Path + "\temp"
SurTxt.LinkPoke
2:
End Sub

Private Sub Frame1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
ProsesTime
End Sub

Private Sub Frame2_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
If (lstDIBList_1.ListCount = 0) Or (Combo1.Text = "") Then
Cmd(7).Enabled = False
Else: Cmd(7).Enabled = True
End If
Combo1_Change: ProsesTime
End Sub
Public Sub ProsesTime()
Dim B1%, C1%, Tim$
If Val(Text1.Text) = 0 Then Text1.Text = 2
B1 = lstDIBList_1.ListCount * Val(Text1.Text)
If B1 < 60 Then
Tim = "FilmTime = 00:" + Fit(B1): Label15.Caption = Tim
Exit Sub
End If
C1 = B1 \ 60
B1 = B1 Mod 60
Tim = "FilmTime = " + Fit(C1) + ":" + Fit(B1)
Label15.Caption = Tim
End Sub


Private Sub Label11_Click()
If File1.ListCount = 0 Then
MsgBox "File Not Found!": Exit Sub
End If
Can = False: H1 = Image1.Height: W1 = Image1.Width: Label6.Visible = True: Label6.Caption = "Sarching Picture By Basic..."
Form1.Caption = Form1.Caption + " - Sarching By Basic...": Combo1.Visible = False
Frame2.Enabled = False: Cmd(8).Visible = True: Pr.Visible = True: Pr.Max = File1.ListCount - 1
For I1 = 0 To File1.ListCount - 1
Pr.Value = I1
If Can = True Then Exit For
Image1.Picture = LoadPicture(Dir1.Path + "\" + File1.List(I1))
    If (Image1.Width = W1) And (Image1.Height = H1) Then
    If Len(Dir1.Path) = 3 Then
    lstDIBList_1.AddItem Dir1.Path + File1.List(I1)
    Else: lstDIBList_1.AddItem Dir1.Path + "\" + File1.List(I1)
    End If:    End If
    File1.ListIndex = I1
    DoEvents
Next
Pr.Visible = False: Cmd(8).Visible = False: Combo1.Visible = True
Frame2.Enabled = True: Label6.Visible = False
Form1.Caption = "Welcoom PicToAvi"
End Sub

Private Sub lstDIBList_1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo A
If button = vbLeftButton Then
If lstDIBList_1.ListIndex = -1 Then Exit Sub
imgPreview.Picture = LoadPicture(lstDIBList_1.Text)
Image1.Picture = LoadPicture(lstDIBList_1.Text)
Label11.Caption = "W= " + left(Str(Image1.Width / 15), 5) + "    H= " + left(Str(Image1.Height / 15), 5)
ElseIf button = vbRightButton Then
Cmd_Click (1)
ElseIf button = 4 Then
Cmd_Click (3)
End If
A: ProsesTime: OutSiz = "": lstDIBList.Clear
End Sub

Private Sub lstDIBList_1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
lstDIBList_1.ToolTipText = lstDIBList_1.Text
End Sub

Public Function Nasa(FName As String, Character As String, Key As Boolean) As String
On Error Resume Next
Dim i As Integer, m As Integer, Z As String
If Key = True Then '------------The Key Is True {Proses String Of Left} -----------
    For i = 1 To Len(FName)
        If Mid$(FName, i, 1) = Character Then
        Nasa = Z
        Exit Function
        End If
        If Mid$(FName, i, 1) <> "" Then
        Z = Z + Mid$(FName, i, 1)
        Else
        Nasa = Z
        Exit Function
        End If
    Next i
Else '------------The Key Is False {Proses String Of Right} -----------------------
    For m = 0 To Len(FName)
        If Z = FName Then
        Z = "Fil"
        Exit Function
        End If
        If Mid$(FName, (Len(FName)) - m, 1) = Character Then
        Nasa = Z
        Exit Function
        End If
        If Mid$(FName, (Len(FName)) - m, 1) <> "" Then
        Z = Mid$(FName, (Len(FName)) - m, 1) + Z
        Else
        Nasa = Z
        Exit Function
        End If
    Next m
End If
Nasa = Z
End Function
Public Function Num(ASTR As String) As Boolean
If Len(ASTR) = 0 Then GoTo W
Dim A As Boolean
For I1 = 1 To Len(ASTR)
A = False
Mp = " " + Mid(ASTR, I1, 1)
For i = 0 To 9
If Mp = Str(i) Then A = True
Next
If A = False Then GoTo W
Next
Num = True
Exit Function
W:
Num = False: Exit Function
End Function


Private Sub SurTxt_Change()
Select Case SurTxt.Text
Case "Min": Me.WindowState = 1
Case "Exit": Unload Me
Case "Normal": Me.WindowState = 0
End Select
End Sub

Private Sub SurTxt_LinkClose()
DDE = False
End Sub

Private Sub SurTxt_LinkOpen(Cancel As Integer)
DDE = True
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 2 Then Text1.Text = right(Text1.Text, 2)
End Sub

Private Sub Text2_Change()
If Val(Text2.Text) < 1 Or Val(Text2.Text) > 30 Then Text2.Text = 1
End Sub
