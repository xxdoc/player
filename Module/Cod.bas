Attribute VB_Name = "Cod"
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
Public ASd As New FileSystemObject
Public Setting As OBJ
Public DDE As Boolean

Public Sub Pres(Mode As String, Frm As Form7)
On Error Resume Next
Dim ds As Integer '----------------------------------------------
With Frm
.CText.Enabled = False
    If Mode = "Cut" Or Mode = "Delet" Then
    ds = MsgBox("Do You Want To " + Mode + "ing This File", vbYesNo)
    If ds = vbNo Then GoTo 2
    End If '------------------------------------------------------
.Can.Visible = True: Cancel = False
Dim D As Integer, I As Integer, B   As Boolean, e As Boolean, O As Boolean
Dim w As String, NG As String, g As String, q As String, p As Boolean
'---------Analiz Copy Object------------------------------------------------------------------------------------
If Mode <> "Delet" Then
If .Text5.Text = Empty Or ASd.FolderExists(.Text5.Text) = False Then
MsgBox ("The Output Directory Is Invalid !")
GoTo 2
End If
End If '---------------------------------------------------------------------
        B = False: e = False: O = False: p = False: .Frame10.Enabled = False
        .Help.Caption = Mode + "ing File Plase With ... ": .Help.BackColor = &H8000000F
                   If .Option6.Value = True Then
                    e = True
                    If .List1.List(0) = "" Then
                    .Help.Caption = "File Not Found!": .Help.BackColor = RGB(300, 200, 0): GoTo 2
                    End If
                Else '-----------------------------------
                    If .File1.List(0) = "" Then
                    .Help.Caption = "File Not Found!": .Help.BackColor = RGB(300, 200, 0):  GoTo 2
                    End If
                    End If '-------------------------------
        If .Option2.Value = True Then B = True
        If .Option4.Value = True Then O = True
If e = True Then '------------------------------------------
.ProgressBar2.Max = .List1.ListCount
Else: .ProgressBar2.Max = .File1.ListCount
End If
'---------Start Coping ...--------------------------------------------------------------------------------
For I = 0 To .ProgressBar2.Max - 1
    If Cancel = True Then .Help.Caption = Mode + "ing Is Cancel": .Help.BackColor = RGB(250, 250, 0): GoTo 2
    g = .Dir1.Path + .File1.List(I)
    If Len(.Dir1.Path) > 3 Then g = .Dir1.Path + "\" + .File1.List(I)
                If e = True And B = False Then
                NG = Nasa(.List1.List(I), "\", False): .List1.ListIndex = I
                w = Nasa(NG, ".", True): NG = w
                ElseIf e = False And B = False Then
                NG = Nasa(.File1.List(I), ".", True): .File1.ListIndex = I
                If Len(Nasa(NG, ":", False)) = Len(NG) + 1 Then
                NG = .Dir1.Path + NG
                Else: NG = .Dir1.Path + "\" + NG
                End If
                Else
                NG = .Text4.Text + Str$(I)
                End If '-----------------------------------------------
                            If Mode <> "Delet" Then
                            If e = True And O = False Then
                            q = "." + Nasa(.List1.List(I), ".", False)
                            ElseIf e = False And O = False Then
                            q = "." + Nasa(.File1.List(I), ".", False)
                         Else '------------------------------------------
                            If Left$(.Text3.Text, 1) <> "." Or Len(.Text3.Text) < 2 Then
                            MsgBox ("The File Type Is Invalid"): GoTo 2
                            End If
                            q = .Text3.Text
                            End If
                            End If
'--------------------Call Prosessing For File------------------------------------------------
Select Case Mode
Case "Delet" '-------------------------------
        If e = True Then
              Call ASd.DeleteFile(.List1.List(I), True)
        Else: Call ASd.DeleteFile(g, True)
        End If
Case "Cut" '-------------------------------
        If e = True Then
              Call ASd.MoveFile(.List1.List(I), .Text5.Text & NG & q)
        Else: Call ASd.MoveFile(g, .Text5.Text + NG + q)
        End If
Case "Copy" '-------------------------------
        If e = True Then
              Call ASd.CopyFile(.List1.List(I), .Text5.Text & NG & q, True)
        Else: Call ASd.CopyFile(g, .Text5.Text + NG + q, True)
        End If
End Select '-------------------------------
     .ProgressBar2.Value = .ProgressBar2.Value + 1: DoEvents
Next I
'---------End Copy------------------------------------------------------------------------
   .Dir1.Refresh: .File1.Refresh: .Help.Caption = Mode + " File Is Completed": .Help.BackColor = RGB(0, 230, 0)
2: .ProgressBar2.Value = 0: .Frame10.Enabled = True: .Can.Visible = False
 .CText.Enabled = True
 End With
End Sub
Public Sub VideoMod(Frm As Form, WMP As WindowsMediaPlayer, _
                LBL As Label)
On Error Resume Next '-----------------------------------------
                    Dim AFD As String
                    If DFA = "" Then Exit Sub
                    Select Case UCase(Right$(DFA, 3))
                    Case "MPG": AFD = "OK"
                    Case "AVI": AFD = "OK"
                    Case "DAT": AFD = "OK"
                    Case "VCD": AFD = "OK"
                    Case "IFO": AFD = "OK"
                    Case "MOV": AFD = "OK"
                    Case "WMV": AFD = "OK"
                    End Select
    If AFD <> "" Then '----------------------------------------
                DFA = ""
                LBL.Caption = "Video"
    Else: LBL.Caption = "For"
    DFA = "": End If
End Sub
Public Sub LChange(Frm As Form, Lst As ListBox, UpDwn As Boolean)
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

Public Function Nasa(FName As String, Character As String, Key As Boolean) As String
On Error Resume Next
Dim I As Integer, m As Integer, Z As String
If Key = True Then '------------The Key Is True {Proses String Of Left} -----------
    For I = 1 To Len(FName)
        If Mid$(FName, I, 1) = Character Then
        Nasa = Z
        Exit Function
        End If
        If Mid$(FName, I, 1) <> "" Then
        Z = Z + Mid$(FName, I, 1)
        Else
        Nasa = Z
        Exit Function
        End If
    Next I
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
Public Sub Credit()
Dim v As Integer
 v = MsgBox("_________________GhayeshRayaneh Convert V1.0012_________________" _
& vbCrLf & "The Program Suported This Media Type" _
& vbCrLf & "__________________________________________________________" _
& vbCrLf & "|*.Wav:   Wave Audio Standard For Windndows                     |" _
& vbCrLf & "|*.M3u:   M3u PlayListFile Standard For Windows                    |" _
& vbCrLf & "|*.mov:   QuickTime MediaFile For Windows                             |" _
& vbCrLf & "|*.mp3;*.mp2;*.mpg;*.dat;*.vcd;*.svd:   All Mpeg MediaFile       |" _
& vbCrLf & "|*.wmv;*.wma;*.wmf:   All Windows MediaFIle                        |" _
& vbCrLf & "|*.mid:   Sund Not And MidiGame Miusic For Windows             |" _
& vbCrLf & "|*.avi:   Windows Standard Video File                                      |" _
& vbCrLf & "|*.ifo:   DVD MpegFormat Video For Windows                          |" _
& vbCrLf & "__________________________________________________________" _
& vbCrLf & "Copyright : GHAYESH RAYANEH All Right Reserved , 2005-2006" _
& vbCrLf & "BY : NasserNiazyMobasser Of Repubic Islamic Of IRAN" _
& vbCrLf & "E-mail : nasser_mobasser@Gmail.com", vbInformation)

End Sub
Public Sub Tyf(Frm As PictureBox, Color As String, Img As Image)
On Error Resume Next
Dim R%, F%, Heght%, Wath%, X%, I%
Heght = 765: Wath = 200
'If Frm.Width < 100 Then Wath = 6120 + 200
F = Heght \ 255
Select Case Color
    Case "Black_With":  GoTo 1
    Case "With_Black":  GoTo 2
End Select
Exit Sub '---------------------------Main--------------------------------------------
1
For I = 0 To Heght Step 2.5
    R = R + 1
    If R = 20000 Then Exit For
        For X = I To F + I
           Frm.Line (0, X)-(Wath, X), RGB(R, R, R)
        Next X
Next I: GoTo B
2  '--------------------------------------------------------------------------------
For I = 0 To Heght Step 2.5
    R = R + 1
    If R = 20000 Then Exit For
        For X = I To F + I
           Frm.Line (0, X)-(Wath, X), RGB(255 - R, 255 - R, 255 - R)
        Next X
Next I '--------------------------------------------------------------------------------
B:
Set Frm.Picture = Frm.Image
Img.Picture = Frm.Picture
End Sub
Public Sub Seting(Frm As Form7)
On Error GoTo 3
Open App.Path + "\NMS\user.dll" For Random Access Read As 5
Get 5, , Setting
Close 5
If Setting.AudioCunt = 0 And Setting.ProgramColor = 0 And Setting.VideoCunt = 0 Then Exit Sub
With Setting
For I = 1 To 16
            Frm.Cmd(I).CausesValidation = .Causes
            Frm.Cmd(I).CheckBoxBehaviour = .Behavi
            Frm.Cmd(I).ShowFocusRect = .Rect
            Frm.Cmd(I).SoftBevel = .Soft
            Frm.Cmd(I).UseGreyscale = .Grey
            Frm.Cmd(I).ButtonType = .Button.BType
            Frm.Cmd(I).SpecialEffect = .Button.BEfect
            Frm.Cmd(I).ColorScheme = .Button.BColor
            If .Button.BColor = 2 Then
                Frm.Cmd(I).BackColor = .Button.Custom(1)
                Frm.Cmd(I).BackOver = .Button.Custom(2)
                Frm.Cmd(I).ForeColor = .Button.Custom(3)
                Frm.Cmd(I).ForeOver = .Button.Custom(4)
                Frm.Cmd(I).MaskColor = .Button.Custom(5)
            End If
Next
End With
3
End Sub
Public Sub ProsesFile()
On Error Resume Next
If Form7.Text1.Text = Empty Then '-------------------------------
        Form7.zASD.Caption = "Plase Click Brows and Select InputFile": Form7.zASD.BackColor = RGB(255, 255, 0)
ElseIf Form7.Text2.Text = Empty Then '----------------------------
        Form7.zASD.Caption = "Plase Click Brows and Select OutputFile": Form7.zASD.BackColor = RGB(200, 200, 0)
ElseIf Form7.Combo1.Text = Empty Or Form7.Combo1.Text = "______" Then
        Form7.zASD.Caption = "Plase Change ComboBox and Select FileType": Form7.zASD.BackColor = RGB(200, 200, 100)
End If '----------------------------------------------------
        Form7.zASD.Caption = "": Form7.zASD.BackColor = &H8000000F
End Sub



