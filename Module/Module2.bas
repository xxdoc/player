Attribute VB_Name = "M_Sink"
Type BTN
    BType           As Integer
    BColor          As Integer
    BEfect          As Integer
    Custom(1 To 5)  As Long
End Type
Type OBJ
    Button          As BTN
    Causes          As Boolean
    Behavi          As Boolean
    Rect            As Boolean
    Soft            As Boolean
    Grey            As Boolean
    ProgramColor    As Long
    AudioCunt       As Integer
    VideoCunt       As Integer
    AudioLoop       As Boolean
    VideoLoop       As Boolean
End Type
Public Setting             As OBJ
Public SinkM               As String
Public Avi                 As Boolean
Public CommonDialog        As cFileDlg
Public DefultPath          As String
Public OutputFile          As String
Public ListFile            As String
Sub SinkP(Frm As FForm1) 'Profetional Sink For Playing-----------------------------
On Error Resume Next
With Frm
            .Width = 4590
            .Height = 3900
            .SSTab3.Top = 0
            .SSTab3.Tab = 4
            .SSTab3.Left = 0
            .Frame(4).Left = 0
            .Frame(2).Left = 0
            .Frame(6).Left = 0
            .Frame(7).Left = 0
            .Frame(10).Left = 0
            .cmd(3).Top = 3000
            .Label4.Width = 1150
            .Frame(1).Top = 122.323
            .Frame(3).Top = 376.245
            .Frame(2).Top = 125.415
            .Frame(4).Top = 125.415
            .Frame(6).Top = 125.415
            .Picture1.Top = 167.22
            .Frame(1).Left = 125.415
            .Frame(3).Left = 122.323
            .Frame(11).Top = 585.271
            .Frame(10).Top = 836.101
            .Frame(7).Top = 125.415
            .Frame(5).Left = 3587.11
            .Frame(11).Left = 122.323
            .Frame(11).BorderStyle = 1
            .Picture1.Left = 3669.678
            .Label4.Caption = "AboutSoftwar"
            .WindowsMediaPlayer1.Width = 4455
            .min.Visible = False
            .max.Visible = False
            .SSTab3.Visible = True
            .cmd(15).Visible = True
            .cmd(3).Visible = False
            .Frame(5).Visible = False
            .Frame(2).Visible = False
            .Line38.Visible = False
            .Frame(6).Visible = False
            .Frame(4).Visible = False
            .cmd(25).Visible = False
            .cmd(80).Visible = False
            .Frame(7).Visible = False
            .cmd(25).Visible = False
            .cmd(13).Visible = False
            .cmd(12).Visible = False
            .cmd(20).Visible = False
            .StatusBar1.Visible = True
End With
Call TabClick(FForm1)
End Sub
Sub SinkS(Frm As FForm1) 'Standard Sink For Playing-----------------------------
On Error Resume Next
With Frm
            .cmd(20).Left = 4525.936: .cmd(20).Visible = False
            .Width = 4740
            .Height = 2955
            .Frame(1).Top = 0
            .SSTab1.Tab = 1
            .SSTab3.Top = 2000
            .Frame(2).Top = 83.61
            .Frame(5).Left = 3792
            .Label4.Width = 615
            .Frame(10).Left = 120
            .Frame(3).Top = 250.83
            .Frame(6).Top = 83.61
            .Frame(10).Top = 501.66
            .Picture1.Left = 3700
            .Frame(4).Top = 961.516
            .cmd(3).Top = 292.635
            .Frame(3).Left = 366.968
            .Picture1.Top = 41.805
            .Frame(1).Left = 122.323
            .Frame(11).Top = 752.491
            .Frame(7).Top = 961.516
            .Frame(4).Left = 122.323
            .Frame(2).Left = 4892.904
            .Frame(6).Left = 4892.904
            .SSTab1.Left = 4892.904
            .Frame(7).Left = 4892.904
            .Frame(11).BorderStyle = 0
            .Frame(11).Left = 5015.226
            .Frame(6).Height = 699.886
            .Label4.Caption = "About"
            .WindowsMediaPlayer1.Width = 4455
            .max.Visible = True
            .min.Visible = True
            .SSTab1.Visible = True
            .Line38.Visible = True
            .Frame(4).Visible = True
            .Frame(2).Visible = True
            .Frame(4).Visible = True
            .Frame(1).Visible = True
            .cmd(3).Visible = True
            .SSTab3.Visible = False
            .Frame(8).Visible = True
            .cmd(13).Visible = True
            .cmd(12).Visible = True
            .cmd(13).Visible = True
            .Frame(11).Visible = False
            .cmd(15).Visible = False
            .cmd(25).Visible = True
            .Frame(7).Visible = True
End With
End Sub
Sub FormMinSize(Frm As FForm1)
On Error Resume Next
With Frm
           .Height = 1250
           .Width = 4770
           .cmd(20).Visible = True
           .Frame(1).Visible = False
           .Frame(2).Visible = False
           .Frame(3).Visible = False
           .Frame(5).Visible = True
           .cmd(4).Visible = False
           .cmd(3).Visible = False
           .Option1.Visible = False
           .cmd(14).Visible = False
           .cmd(12).Visible = False
           .Image10.Visible = False
           .cmd(25).Visible = False
           .Picture1.Visible = False
           .WindowsMediaPlayer1.Width = 3700
           .Frame(10).Width = 3750
           .Frame(10).Left = 0
           .Frame(10).Top = 0
           .StatusBar1.Visible = False
           .Frame(10).Visible = True
End With
End Sub
Sub MinSize(Frm As FForm1)
On Error Resume Next
If Frm.cmd(15).Caption = "<" Then
    With Frm '---------------------Go To Min Mode---------------------------------------
            .cmd(15).Caption = ">"
            .WindowsMediaPlayer1.Width = 3495
            .cmd(20).Left = 4336.336
            .cmd(15).Top = 83.61
            .Height = 1250
            .cmd(20).Visible = True
            .SSTab3.Visible = False
            .Frame(5).Visible = True
            .Frame(10).Visible = True
            .Frame(6).Visible = False
            .Frame(2).Visible = False
            .Frame(4).Visible = False
            .Frame(1).Visible = False
            .Frame(3).Visible = False
            .Frame(7).Visible = False
            .Frame(10).Width = 3562.646
            .Frame(11).Visible = True
            .Picture1.Visible = False
            .StatusBar1.Visible = False
            .Frame(10).Top = 0
End With
ElseIf Frm.cmd(15).Caption = ">" Then
 With Frm '------------------Back ToThe Standard Mode--------------------------------------
                        .WindowsMediaPlayer1.Width = 4455
                        .Frame(10).Width = 4541.226
                        .cmd(15).Caption = "<"
                        .Frame(10).Top = 836.101
                        .cmd(15).Top = 10
                        .Height = 3900
                        .Width = 4590
                        .SSTab3.Visible = True
                        .Frame(2).Visible = False
                        .Frame(6).Visible = False
                        .Frame(5).Visible = False
                        .cmd(20).Visible = False
                        .StatusBar1.Visible = True
    Call TabClick(FForm1) '-------------------------------
    End With
End If
End Sub

Sub FormSizeChange(Frm As FForm1)
On Error Resume Next '-------------------------------
            With Frm
            If .Height = 2955 And .Width = 4740 Then
                    .Height = 5400
                    .Width = 9360
                    .cmd(25).Caption = "<<<<"
                    .Frame(10).Visible = True
                    .Frame(11).Visible = True
            ElseIf .Width = 9360 And .Height = 5400 Then
                    .Height = 2955
                    .Width = 4740
                    .cmd(25).Caption = ">>>>"
                    .Frame(3).Visible = True
                    .Frame(11).Visible = False
            Else
'{--------------ReSize Form To Standard Mode-----------------------------------------}
                            .Width = 4740
                            .Height = 2955
                            .Frame(10).Top = 501
                            .Frame(10).Left = 120
                            .Frame(10).Width = 4541.226
                            .WindowsMediaPlayer1.Width = 4455
                            .Frame(1).Visible = True
                            .Frame(2).Visible = True
                            .Frame(3).Visible = True
                            .Frame(5).Visible = False
                            .Frame(7).Visible = True
                            .Option1.Visible = True
                            .Image10.Visible = True
                            .Frame(11).Visible = False
                            .Picture1.Visible = True
                            .cmd(3).Visible = True
                            .cmd(4).Visible = True
                            .cmd(12).Visible = True
                            .cmd(14).Visible = True
                            .cmd(20).Visible = False
                            .cmd(25).Visible = True
                            .cmd(25).Caption = ">>>>"
                            .StatusBar1.Visible = True
                            SinkM = "FORM"
            End If
            End With
End Sub

Sub TabClick(Frm As FForm1)
On Error Resume Next
With Frm '-----------------------------------
 .Frame(6).Visible = False: .Frame(2).Visible = False: .Frame(4).Visible = False: .Frame(1).Visible = False: .Frame(3).Visible = False: .Frame(11).Visible = False: .Picture1.Visible = False
    Select Case .SSTab3.Tab
      Case 0: .Frame(6).Visible = True
      Case 1: .Frame(2).Visible = True
      Case 2: .Frame(4).Visible = True
      Case 3: .Frame(7).Visible = True
      Case 4: .Frame(1).Visible = True
              .Frame(3).Visible = True
              .Frame(11).Visible = True
              .Picture1.Visible = True
    End Select '-------------------------------
End With
End Sub
Sub ALoad(Frm As FForm1)
On Error Resume Next
            Frm.WindowsMediaPlayer1.URL = App.Path + "\" + "Test.Mp3"
            Frm.WindowsMediaPlayer1.Controls.Play: Frm.URL.Caption = App.Path + "Test.Mp3"
            Frm.F.Caption = "For"
End Sub
 Sub Setng(Frm As Form, MaxCmd As Integer, Direct As Boolean _
            , Wmp As WindowsMediaPlayer, Lst As ListBox, MaxFrame As Integer)
With Setting
If .AudioCunt = 0 And .VideoCunt = 0 Then Exit Sub
            Wmp.settings.playCount = .AudioCunt
For i = 1 To MaxCmd
            Frm.cmd(i).CausesValidation = .Causes
            Frm.cmd(i).CheckBoxBehaviour = .Behavi
            Frm.cmd(i).ShowFocusRect = .Rect
            Frm.cmd(i).SoftBevel = .Soft
            Frm.cmd(i).UseGreyscale = .Grey
            Frm.cmd(i).ButtonType = .Button.BType
            Frm.cmd(i).SpecialEffect = .Button.BEfect
            Frm.cmd(i).ColorScheme = .Button.BColor
            If .Button.BColor = 2 Then
                Frm.cmd(i).BackColor = .Button.Custom(1)
                Frm.cmd(i).BackOver = .Button.Custom(2)
                Frm.cmd(i).ForeColor = .Button.Custom(3)
                Frm.cmd(i).ForeOver = .Button.Custom(4)
                Frm.cmd(i).MaskColor = .Button.Custom(5)
            End If
Next
FForm1.Label4.CausesValidation = .Causes
FForm1.Label4.CheckBoxBehaviour = .Behavi
FForm1.Label4.ShowFocusRect = .Rect
FForm1.Label4.SoftBevel = .Soft
FForm1.Label4.UseGreyscale = .Grey
FForm1.Label4.ButtonType = .Button.BType
FForm1.Label4.SpecialEffect = .Button.BEfect
FForm1.Label4.ColorScheme = .Button.BColor
  If .Button.BColor = 2 Then
    FForm1.Label4.BackColor = .Button.Custom(1)
    FForm1.Label4.BackOver = .Button.Custom(2)
    FForm1.Label4.ForeColor = .Button.Custom(3)
    FForm1.Label4.ForeOver = .Button.Custom(4)
    FForm1.Label4.MaskColor = .Button.Custom(5)
  End If
End With
For i = 1 To MaxFrame
            Frm.Frame(i).BackColor = Setting.ProgramColor
Next
Frm.BackColor = Setting.ProgramColor
If Direct = True Then
Frm.Dir1.BackColor = Setting.ProgramColor
Frm.Drive1.BackColor = Setting.ProgramColor
Frm.File1.BackColor = Setting.ProgramColor
End If
Lst.BackColor = Setting.ProgramColor: FForm1.Label4.ButtonType = Setting.Button.BType

End Sub
Public Sub ShowOpen(Filter As String, FrmList As Form, Frm As Form, Wmp As WindowsMediaPlayer, Mode As String, ByVal Vhs As Boolean)
On Error Resume Next
Set CommonDialog = New cFileDlg
'{-----------------------------------------------------------
                    CommonDialog.DlgTitle = Mode
                    CommonDialog.Filter = Filter
                    CommonDialog.DefaultExt = DefultPath
                    CommonDialog.OwnerhWnd = Frm.hwnd
If Mode = "Save Playlist" Then '------------------------------------------
      If CommonDialog.VBGetSaveFileName(OutputFile) <> True Then Exit Sub
Else: If CommonDialog.VBGetOpenFileName(OutputFile) <> True Then Exit Sub
  End If '----------------------------------
      If OutputFile = "" Then Exit Sub
Select Case Mode '--------------Start Proses Mode--------------------------------
    Case "Open Media":
                    DfA = OutputFile
                    Call VideoMod(Frm, Wmp, Frm.F, True, Vhs)
                    Plaing Frm, Frm.F, Wmp, OutputFile, Vhs
    Case "Add media": '-------------------------------
                    If Frm.FG.Caption <> (OutputFile) Then
                    Frm.List1.AddItem (OutputFile)
                    Frm.FG.Caption = (OutputFile)
                    End If
    Case "Open Playlist": '-------------------------------
                    If (UCase(Right$(OutputFile, 3)) <> "M3U") _
                    Or OutputFile = "" Then Exit Sub
                Call Opening(FrmList, OutputFile): DfA = OutputFile
                Call ListE(FrmList, FrmList.List1, Wmp, DfA, False): Wmp.URL = OutputFile
                     Plaing Frm, Frm.F, Wmp, OutputFile, Vhs
    Case "Save Playlist": '-------------------------------------
                    If (UCase(Right$(OutputFile, 3)) <> "M3U") _
                    Or OutputFile = "" Then Exit Sub
                Call Saving(FrmList, FrmList.List1, OutputFile): DfA = OutputFile
                Call ListE(FrmList, FrmList.List1, Wmp, DfA, False)
                     Plaing Frm, Frm.F, Wmp, OutputFile, Vhs
End Select
End Sub
Public Sub Plaing(Frm As Form, Lbl As Label, Wmp As WindowsMediaPlayer, OFile As String, ByVal Vhs As Boolean)
Frm.URL.Caption = OFile '-------------------Start Play File-----------------------
            If (Lbl.Caption = "Video") Then
            If Vhs = True Then '------------sho Video---------------
            Form3.Show
            If Form3.WindowsMediaPlayer1.URL = OFile Then Exit Sub
            Wmp.URL = "": Form3.WindowsMediaPlayer1.settings.autoStart = True
            Form3.WindowsMediaPlayer1.URL = OFile
            Else: Unload Form3: Wmp.settings.autoStart = True
            If Wmp.URL = OFile Then Exit Sub
            Wmp.URL = OFile
            End If
            Else '---------Audio File----------------------------
            Unload Form3: If Wmp.URL = OFile Then Exit Sub
            Wmp.URL = OFile
            End If
End Sub
Public Sub Credit()
Dim v As Integer
 v = MsgBox("_________________GhayeshRayaneh Player V2.4__________________" _
& vbCrLf & "The Program Suported This Media Type" _
& vbCrLf & "__________________________________________________________" _
& vbCrLf & "|*.Wav:   Wave Audio Standard For Windndows                     |" _
& vbCrLf & "|*.M3u:   M3u PlayListFile Standard For Windows                    |" _
& vbCrLf & "|*.mov:   QuickTime MediaFile For Windows                             |" _
& vbCrLf & "|*.mp3;*.mp2;*.mpg;*.dat;*.vcd;*.svd:   All Mpeg MediaFile |" _
& vbCrLf & "|*.wmv;*.wma;*.wmf:   All Windows MediaFIle                        |" _
& vbCrLf & "|*.mid:   Sund Not And MidiGame Miusic For Windows               |" _
& vbCrLf & "|*.avi:   Windows Standard Video File                                       |" _
& vbCrLf & "|*.ifo:   DVD MpegFormat Video For Windows                           |" _
& vbCrLf & "__________________________________________________________" _
& vbCrLf & "Copyright : GHAYESH RAYANEH All Right Reserved , 2005-2006" _
& vbCrLf & "BY : NasserNiazyMobasser Of Repubic Islamic Of IRAN" _
& vbCrLf & "E-mail : nasserVB@Gmail.com" _
& vbCrLf & "__________________________" _
& vbCrLf & "HomePage:http://NasserVb.blogfa.com" _
& vbCrLf & "SupportSite:http://TcVb.blogfa.com", vbInformation)

End Sub
Public Sub Tyf(Frm As Control, Color As String, Img As Image)
On Error GoTo B
Dim r%, F%, Heght%, Wath%, x%
Heght = Frm.Height + 200: Wath = 300
F = Heght \ 255
Select Case Color
    Case "Green_Black": GoTo 3
    Case "With_Black":  GoTo 4
End Select
Exit Sub '---------------------------Main--------------------------------------------
3 '--------------------------------------------------------------------------------
For i = 0 To Heght Step F
    r = r + 1
    If r = 20000 Then Exit For
        For x = i To F + i
           Frm.Line (0, x)-(Wath, x), RGB(0, 255 - r, 0)
        Next x
Next i: GoTo B
4 '--------------------------------------------------------------------------------
For i = 0 To Heght Step F
    r = r + 1
    If r = 20000 Then Exit For
        For x = i To F + i
           Frm.Line (0, x)-(Wath, x), RGB(255 - r, 255 - r, 255 - r)
        Next x
Next i: GoTo B
B:
Set Frm.Picture = Frm.Image
Img.Picture = Frm.Picture: Frm.Cls
End Sub
