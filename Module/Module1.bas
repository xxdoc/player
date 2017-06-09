Attribute VB_Name = "PLayer_Cod"
 Option Explicit
'****************************************************************************************************
'*                                                                                                  *
'*                 Copyright's(c) GHAYESH RAYANEH(GhayeshSoft) , 2005 - 2007                       *
'*                                                                                                  *
'*    The Softwar Cod Writed & Drawed For GhayeshRayaneh .By : NasserNiazyMobasser . For            *
'*    Republic Islamic OF IRAN . The Softwar Hase Tow Sink For Playing . ActiveX     For            *
'*    Microsoft WindowsMediaPlayer . Mail:nasservb@gmail.com                                        *
'*    HomePage: http://Nasservb.Blogfa.com  Support Site:TcVB.Blogfa.com                            *
'*                                                                                                  *
'****************************************************************************************************
'{-----------------Type UserData For Save Setting--------------------------------}
Public Type ClientRecord
            IconForm               As Integer
            WmpURL                 As String * 100
            Sink                   As String * 4
End Type
'{-----------------Type Conset For Form1 Api----------------------------------------------}
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
'{------------------Declearing API Windows Setting And Shell-------------------------------}
Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long _
    , ByVal x As Long, ByVal Y As Long, ByVal cx As Long _
    , ByVal cy As Long, ByVal wFlags As Long) As Long
'---------------------------------------------------------------------------------
Public Declare Function GetWindowThreadProcessId Lib "user32" _
                    (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'{-------------------Type Public---------------------------------------------------}
            Public ASD            As New FileSystemObject 'For Prossing Files And Folders
            Public DAT1           As ClientRecord 'For Save Setting
            Public ForSho         As Boolean 'For Form1 Showing Mode
            Public Pla            As Boolean 'For Play Video Time
            Public DDE            As Boolean 'For Conect Tools Time
            Public FForm1_WMP_Vol As Integer 'For Program Audio Control
            Public Intcnt         As Integer 'For About Form Id
            Public ForIcon        As Integer 'For Form1 API
            Public i              As Integer 'For Loop's
            Public DfA            As String  'For Call Funection, File's Adress
            Public Path           As String  'For Windows Path,Copy Dll File Of Softwar Directory In New windows System32 Path
            Public For1_X         As Single  'For Form1 Left Of Screen(form1.left=X)
Sub LoadForm(ByRef Frm As FForm1) 'Loading FForm1 (Main Form)
'On Error Resume Next '-------------------------------
            Dim B As String, n As Integer, x As Integer
            Open App.Path + "\NMS\DAT1.DLL" For Random Access Read As #1
            Get #1, , DAT1
            Close #1 '-------------------------------
    B = Nasa(DAT1.WmpURL, "#", True)
    x = Val(DAT1.IconForm)
            If Len(B) <> 100 Then
            DfA = B: Call VideoMod(FForm1, Frm.WindowsMediaPlayer1, Frm.F, True, True): Frm.URL.Caption = B: FForm1.WindowsMediaPlayer1.Controls.stop
            End If '-----------------------------------
            If x = 0 Or x = 8738 Then x = 1
            Frm.Image10.Picture = Frm.ImageList1.ListImages(Val(x)).Picture: ForIcon = Val(x)
           Open App.Path + "\NMS\User.Dll" For Random Access Read As #3
           Get #3, , Setting
           Close
            Setng Frm, 31, True, Frm.WindowsMediaPlayer1, Frm.List1, 12
End Sub
Sub DriveChange(Frm As Form, Driv As DriveListBox, Dir As DirListBox) 'Cheneging Drive In Drive Box
On Error GoTo errorhandler
                    Dir.Path = Driv.Drive
                    Exit Sub
errorhandler: '-------------------------------(Error For Cd & Floppy Drive)
                    Dim message As String
                    If Err.Number = 68 Then
                    Dim r As Integer
                    message = "drive not invalide": r = MsgBox(message, vbRetryCancel + vbCritical, "Player")
                    If r = vbRetry Then
                    Resume
                    Else
                    Driv.Drive = Driv.List(1)
                    Resume Next
                    End If
                    Else
                    Call MsgBox(Err.Description, vbOKOnly + vbExclamation)
                    Resume Next
                    End If
End Sub
Public Sub lode(Frm As Form) 'For Form6 Load In Form1
On Error Resume Next
Dim i As Integer '-------------------------------
    If ForSho = False Then
    Frm.SSTab1.Visible = True
        For i = 0 To FForm1.List1.ListCount - 1
        Frm.List1.AddItem (FForm1.List1.List(i))
        Next i '-------------------------------
    Frm.Width = 4755: Frm.Height = 3480
    Frm.Drive1.Drive = FForm1.Drive1.Drive: Frm.Dir1.Path = FForm1.Dir1.Path
    Frm.Slider1.Value = FForm1.Slider1.Value
    Frm.Slider13.Value = FForm1.Slider4.Value
    Frm.Slider4.Value = FForm1.Slider2.Value
    Frm.Slider3.Value = FForm1.Slider3.Value
    ForSho = True
    End If '-------------------------------
    Frm.Left = Form1.Left: Frm.Top = 70
    Form1.For6 = True: Frm.Height = 3480
    Frm.Width = 4755: Frm.Caption = "Player Option"
End Sub
Public Sub SmalSpeed(Frm As Form, Wmp As WindowsMediaPlayer _
                , SLD1 As Slider, SLD2 As Slider, Lbl As Label) 'For Control Speed Short(In Form.slider3)
On Error Resume Next '-------------------------------------------
Dim A As Integer, B As Integer, g As Integer, j As Integer, v As Single, s As String
    If SLD2.Value < 5 Then 'Sped Is Up----------------------------------------
            g = (SLD2.Value + 5)
            j = (SLD1.Value)
            s = Str$(g) + Str$(j)
            Lbl.Caption = "Larenc " + Str$(Val(s)) + "%"
            B = SLD1.Value
            A = SLD2.Value + 5
            s = ("0." + (Str$(A)) + (Str$(B)))
            v = Val(s)
            Wmp.settings.Rate = v
    ElseIf SLD2.Value > 4 Then 'Speed Is Down-------------------------------------
            g = (SLD2.Value - 5)
            j = (SLD1.Value)
            s = "1" + Str$(g) + Str$(j)
            Lbl.Caption = "Larenc " + Str$(Val(s)) + "%"
            B = SLD1.Value
            A = SLD2.Value - 5
            s = ("1." + (Str$(A)) + (Str$(B)))
            v = Val(s)
            Wmp.settings.Rate = v
    End If '-------------------------------------------------------------
End Sub

Sub PlaySpeed(Frm As Form, SLD1 As Slider, SLD2 As Slider _
                 , Wmp As WindowsMediaPlayer, Lbl As Label) 'For Control Speed Larg
On Error Resume Next '----------------------------------
                Select Case SLD1.Value
                Case 1 '-------------------------------
                     Wmp.settings.Rate = 0.6
                     Lbl.Caption = "Play Speed  60 %"
                Case 2 '-------------------------------
                     Wmp.settings.Rate = 0.7
                     Lbl.Caption = "Play Speed  70 %"
                Case 3 '-------------------------------
                     Lbl.Caption = "Play Speed  80 %"
                     Wmp.settings.Rate = 0.8
                Case 4 '-------------------------------
                     Lbl.Caption = "Play Speed  90 %"
                     Wmp.settings.Rate = 0.9
                Case 5 '-------------------------------
                     Lbl.Caption = "Play Speed Normal"
                     Wmp.settings.Rate = 1
                Case 6 '-------------------------------
                     Lbl.Caption = "Play Speed  110 %"
                     Wmp.settings.Rate = 1.1
                Case 7 '-------------------------------
                     Lbl.Caption = "Play Speed  120 %"
                     Wmp.settings.Rate = 1.2
                Case 8 '-------------------------------
                     Lbl.Caption = "Play Speed  130 %"
                     Wmp.settings.Rate = 1.3
                Case 9 '-------------------------------
                     Lbl.Caption = "Play Speed  140 %"
                     Wmp.settings.Rate = 1.4
                Case 10 '-------------------------------
                     Lbl.Caption = "Play Speed  150 %"
                     Wmp.settings.Rate = 1.5
                End Select '----------------------------
                    SLD2.Value = 0
End Sub
Sub Saving(Frm As Form, Lst As ListBox, OutPath As String) 'For Saving Playlist
On Error Resume Next '--------------------------------------------------
                    Dim T3 As String, T2, strans As String, L As Single, i As Integer
                    T3 = "": T2 = ""
                    Lst.Selected(0) = True
                    If Lst.ListIndex = -1 Then
                    strans = MsgBox("File Not Found!", vbCritical)
                    Exit Sub
                    End If
            Open OutPath For Output As #1
                    Print #1, "#EXTM3U:"
                For i = 0 To Lst.ListCount - 1 'Start Save In File-----------------------
                    Lst.Selected(i) = True
                    If T2 = "OK" Then
                    If Lst.Text = T3 Then
                    Exit For
                    End If
                    T3 = Lst.Text
                    End If
                    Print #1, "#EXTNIF:"
                    Print #1, Lst.Text
                    If (Lst.Selected(i) = True) = False Then
                    Exit For
                    End If
                    If T2 = "" Then
                    T2 = "OK"
                    T3 = Lst.Text
                    End If
                Next i '------------------------------------------------------
            Close #1
End Sub
Sub QuickSave(Frm As Form, Wmp As WindowsMediaPlayer, ByVal Vhs As Boolean) 'For Savin Quick Of list
On Error Resume Next '---------------------------------------------------------
                    Call Saving(Frm, Frm.List1, App.Path + "\NMS\DAT2.M3U"): DfA = App.Path + "\NMS\DAT2.M3U"
                    Call ListE(Frm, Frm.List1, Wmp, App.Path + "\nms\dat2.m3u", Vhs)
End Sub
Sub ListE(Frm As Form, Lst As ListBox, Wmp As WindowsMediaPlayer, ByVal PFil As String, ByVal Sho As Boolean) 'Proses List For Playing AudioList Or videolist
On Error Resume Next
                    Dim i As Integer
                    Lst.Selected(0) = True
                    If Lst.ListIndex = -1 Then Exit Sub
                For i = 0 To Lst.ListCount - 1
                    Lst.Selected(i) = True
                    DfA = Lst.Text
                    Call VideoMod(Frm, Wmp, Frm.F, False, False)
                       If Frm.F.Caption = "For" Then Exit For
                Next i '--------------------------------------------------------------
                   Plaing Frm, Frm.F, Wmp, PFil, Sho: ListFile = PFil
End Sub
Public Sub VideoMod(Frm As Form, Wmp As WindowsMediaPlayer, _
                Lbl As Label, ByVal Sho As Boolean, Play As Boolean) 'Proses File For Video Or Audio
On Error Resume Next '-----------------------------------------
                    Dim AFD As String, Fil As String
                    If DfA = "" Then Exit Sub
                    Select Case UCase(Right$(DfA, 3))
                    Case "MPG": AFD = "OK"
                    Case "AVI": AFD = "OK"
                    Case "DAT": AFD = "OK"
                    Case "VCD": AFD = "OK"
                    Case "IFO": AFD = "OK"
                    Case "MOV": AFD = "OK"
                    Case "WMV": AFD = "OK"
                    Case "M3U": AFD = "NO"
                    End Select
If Sho = False And Play = False Then
If AFD = "OK" Then Lbl.Caption = "Video"
If AFD = "" Then Lbl.Caption = "For"
Exit Sub: End If
    If AFD = "OK" Then '----------------------------------------
            If Sho = True Then '<<<<<<<<<<<<<<<<<<<<<
                Lbl.Caption = "Video": Plaing Frm, Lbl, Wmp, DfA, True
            ElseIf Play = True And Sho = False Then
                  Lbl.Caption = "For": Plaing Frm, Lbl, Wmp, DfA, False
            End If '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                DfA = ""
                Lbl.Caption = "Video"
    ElseIf AFD = "NO" Then
                Wmp.settings.autoStart = True: ListFile = DfA
                Call Opening(Frm, DfA): Fil = DfA
                Call ListE(Frm, Frm.List1, Wmp, DfA, Sho)
    Else: Lbl.Caption = "For"
    Unload Form3
    If Play = True Then '----------------------------------------
    If Wmp.URL <> DfA Then
     Wmp.settings.autoStart = True: Wmp.URL = DfA: Frm.URL.Caption = DfA
    End If: End If: DfA = "": End If
End Sub
Sub Opening(Frm As Form, InPath) 'Open PlayList
On Error GoTo Handel
                Dim L$, d$, T3$, T4$, CLK%, i%
                T3 = "": T4 = ""
                Open InPath For Input As #5
                Frm.List1.Clear
                        Line Input #5, d$
                For i = 0 To 10000 '----------------------{Strat For
                        Line Input #5, d$
                        If d = "" Then Exit For
                        Line Input #5, L$: Frm.List1.AddItem (L$)
60              Next '-----------------------------------------------}End For
Handel:                Close #5
End Sub
Sub SetFile(Wmp As WindowsMediaPlayer, Lst As Control, ByVal SetVideo As Boolean _
, Lbl1 As Label, LBL2 As Label, Frm As Form) 'For DblClick In List Or directoryFile
                    On Error Resume Next
                    
                    Dim j As String, FName As String
                    FName = Lst.List(Lst.ListIndex): j = FName
            If TypeOf Lst Is FileListBox Then  '-----------------------------------
            If Len(Frm.Dir1.Path) < 4 Then
                    j = Frm.Dir1.Path + FName
            Else:   j = Frm.Dir1.Path + "\" + FName
            End If: End If '------------------------------------------
                    If FName = "" Then
                    MsgBox ("File Not Fund"): Exit Sub
                    Else: DfA = j$: End If
                    LBL2.Caption = DfA
                    Call VideoMod(Frm, Wmp, Lbl1, SetVideo, True)
                         Plaing Frm, Lbl1, Wmp, j$, SetVideo
End Sub
Sub RETURNForm(Frm As Form, Wmp As WindowsMediaPlayer, ByVal Sho As Boolean)
On Error Resume Next 'For Return Playlist Files In List
                    If ListFile = "" Then Exit Sub
                    Call Opening(Frm, ListFile): DfA = ListFile
                    Call ListE(Frm, Frm.List1, Wmp, DfA, False)
                         Plaing Frm, Frm.F, Wmp, ListFile, Sho
End Sub
Sub EqualizerForm(Frm As Form) 'For Digital Equlizer audio
On Error Resume Next
Dim h As Integer, L As Integer, A As Integer, B As Integer, x As Integer, c As Integer
With Frm '{---------------Calling Randomiz Audio Signal---------------------------------}
            h = 800: L = 1000 '--------------------------------------*
Call Randomize: A = (Int(Rnd * (h - L + 1) + L)): .Lin(11).Y1 = A - 260
            h = 700: L = 550 '---------------------------------------*
Call Randomize: B = (Int(Rnd * (h - L + 1) + L)): .Lin(2).Y1 = B:
            h = 650: L = 480 '---------------------------------------*
Call Randomize: x = (Int(Rnd * (h - L + 1) + L)): .Lin(5).Y1 = x + 50
'{----------Start Audio Equlizer----------------------------------------------}
        .Lin(9).Y1 = ((A + B) / 2) - 178: .Lin(10).Y1 = ((A + B) / 2) - 200: .Lin(16).Y1 = A - 200: .Lin(8).Y1 = B: .Lin(7).Y1 = ((x + B) / 2) - 130: .Lin(15).Y1 = ((x + B) / 2)
        .Lin(6).Y1 = ((x + B) / 2) - 120: .Lin(14).Y1 = x: .Lin(12).Y1 = .Lin(10).Y1: .Lin(1).Y1 = .Lin(9).Y1: .Lin(3).Y1 = .Lin(15).Y1: .Lin(4).Y1 = .Lin(6).Y1: .Lin(13).Y1 = .Lin(7).Y1 - 80
        .Lin(17).Y1 = (A + B + x) / 3 - 200: .Lin(18).Y1 = (A + B + x) / 3 - 180: .Lin(19).Y1 = (.Lin(4).Y1 + .Lin(7).Y1) / 2: .Lin(20).Y1 = .Lin(10).Y1 + 60: .Lin(21).Y1 = .Lin(12).Y1 + 50
        .Lin(22).Y1 = .Lin(5).Y1 + 40: .Lin(23).Y1 = ((A + x + c) / 3): .Lin(24).Y1 = .Lin(8).Y1 + 34: .Lin(25).Y1 = .Lin(1).Y1: .Lin(26).Y1 = .Lin(16).Y1: .Lin(27).Y1 = (A + x) \ 2
        .Lin(28).Y1 = (x - 48): .Lin(29).Y1 = A - 100: .Lin(30).Y1 = B: .Lin(31).Y1 = (x + B) \ 2 + 89: .Lin(32).Y1 = (A - x) + B: .Lin(33).Y1 = (B - A) + x + 360: .Lin(34).Y1 = (x - B) + A: .Lin(35).Y1 = (A + x) - B + 32: .Lin(36).Y1 = (x + B) - A + 420
End With: DoEvents
End Sub

Sub LineColor(ModeL As Boolean, Colr As Single, Frm As Form)  '[Model,True =Chang,False=Main]
On Error Resume Next 'For Cheneg Equlizer Color
If ModeL = False Then
Frm.CommonDialog1.DialogTitle = "Select Color": Frm.CommonDialog1.ShowColor
Else: Frm.CommonDialog1.Color = Colr
End If
                With Frm
                For i = 1 To 36
                .Lin(i).BorderColor = .CommonDialog1.Color
                Next: End With
 
End Sub
Public Sub Balance(Sld As Slider, Wmp As WindowsMediaPlayer _
, ImgL1 As Image, ImgR1 As Image, Frm As Form) 'Balans Audio
On Error Resume Next
With Frm
                    Select Case Sld.Value
                    Case 1: Wmp.settings.Balance = -5000 'Left Spaker------------------
                              ImgL1.Picture = FForm1.ImageList2.ListImages(5).Picture
                              ImgR1.Picture = FForm1.ImageList2.ListImages(4).Picture
                    Case 2: Wmp.settings.Balance = 0 'Normal Spaker----------------------
                              ImgL1.Picture = FForm1.ImageList2.ListImages(5).Picture
                              ImgR1.Picture = FForm1.ImageList2.ListImages(6).Picture
                    Case 3: Wmp.settings.Balance = 5000 'Right Spaker-------------------
                              ImgL1.Picture = FForm1.ImageList2.ListImages(3).Picture
                              ImgR1.Picture = FForm1.ImageList2.ListImages(6).Picture
                    End Select
End With
End Sub
Public Sub Proses(InPath As String) 'Prosessing in Instaltion File For Windows
On Error Resume Next
If ASD.FileExists(InPath + "Command.ocx") = False Then ASD.CopyFile App.Path + "\Command.ocx", InPath, True
End Sub
Public Sub UnProses(InPath As String) 'Prosessing in Instaltion File For Program
On Error Resume Next
If ASD.FileExists(App.Path + "\Command.ocx") = False Then ASD.CopyFile InPath + "\command.ocx", App.Path + "\", True
End Sub
Public Sub LChange(Frm As Form, Lst As ListBox, UpDwn As Boolean) 'List file Down Or Up
On Error Resume Next
If UpDwn = True Then '----------------------------The Proses Is Upning In List--------
Dim A As Integer, F As String
        A = Lst.ListIndex
        If A = -1 Then Exit Sub
        If Lst.ListIndex = 0 Then Exit Sub
        F = Lst.Text
        Lst.RemoveItem (Lst.ListIndex)
        Call Lst.AddItem(F, A - 1)
        Lst.Selected(A - 1) = True
Else '----------------------------The Proses Is Downing In List-----------------------
        A = Lst.ListIndex
        If A = -1 Then Exit Sub
        If Lst.ListIndex = Lst.ListCount - 1 Then Exit Sub
        F = Lst.Text
        Lst.RemoveItem (Lst.ListIndex)
        Call Lst.AddItem(F, A + 1)
        Lst.Selected(A + 1) = True
End If
End Sub

Public Function Nasa(FName As String, Character As String, Key As Boolean) As String '[Key,True=Left ,False=Right]
On Error Resume Next 'For Cuting String
Dim i As Integer, M As Integer, Z As String
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
    For M = 0 To Len(FName)
        If Z = FName Then
        Z = "Fil"
        Exit Function
        End If
        If Mid$(FName, (Len(FName)) - M, 1) = Character Then
        Nasa = Z
        Exit Function
        End If
        If Mid$(FName, (Len(FName)) - M, 1) <> "" Then
        Z = Mid$(FName, (Len(FName)) - M, 1) + Z
        Else
        Nasa = Z
        Exit Function
        End If
    Next M
End If
Nasa = Z
End Function

Public Sub NasserNiazyMobasser_Emza(Emza As String) 'Programer Emza!!
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
