Attribute VB_Name = "Module1"
Type BTN
    BType           As Integer
    BColor          As Integer
    BEfect          As Integer
    Custom(1 To 5)  As Long
End Type
Type OBJ
    button          As BTN
    Causes          As Boolean
    Behavi          As Boolean
    Recte           As Boolean
    Soft            As Boolean
    Grey            As Boolean
    ProgramColor    As Long
    AudioCunt       As Integer
    VideoCunt       As Integer
    AudioLoop       As Boolean
    VideoLoop       As Boolean
End Type
'Public Declare Function GetPixel Lib "gdi32" _
'                       (ByVal hDC As Long, _
'                        ByVal X As Long, _
'                        ByVal Y As Long) As Long
'Public Declare Function SetWindowRgn Lib "user32" _
'              (ByVal hWnd As Long, _
'               ByVal hRgn As Long, _
'               ByVal bRedraw As Boolean) As Long
'Public Declare Function CreateRectRgn Lib "gdi32" _
'                                     (ByVal X1 As Long, _
'                                      ByVal Y1 As Long, _
'                                      ByVal X2 As Long, _
'                                      ByVal Y2 As Long) As Long
'Public Declare Function CombineRgn Lib "gdi32" _
'      (ByVal hDestRgn As Long, _
'       ByVal hSrcRgn1 As Long, _
'       ByVal hSrcRgn2 As Long, _
'       ByVal nCombineMode As Long) As Long
'Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Public Const RGN_OR = 2
'Public Const WM_NCLBUTTONDOWN = &HA1
'Public Const HTCAPTION = 2
Public Setting    As OBJ
Public A          As String
Public ASD        As New FileSystemObject
Public DDE        As Boolean
' ---------------------------------------------------------------
Sub DriveChange(Frm As Form, Driv As DriveListBox, Dir As DirListBox)
On Error GoTo errorhandler
                    Dir.Path = Driv.Drive
                    Exit Sub
errorhandler: '-------------------------------
                    Dim message As String
                    If Err.Number = 68 Then
                    Dim R As Integer
                    message = "drive not invalide": R = MsgBox(message, vbRetryCancel + vbCritical, "Player")
                    If R = vbRetry Then
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
Public Sub LChange(Frm As Form, LST As ListBox, UpDwn As Boolean)
On Error Resume Next
If UpDwn = True Then '----------------------------The Proses Is Upning In List--------
Dim A As Integer, F As String
        A = LST.ListIndex
        If A = -1 Then Exit Sub
        If LST.ListIndex = 0 Then Exit Sub
        F = LST.Text
        LST.RemoveItem (LST.ListIndex)
        Call LST.AddItem(F, A - 1)
        LST.Selected(A - 1) = True
Else '----------------------------The Proses Is Downing In List-----------------------
        A = LST.ListIndex
        If A = -1 Then Exit Sub
        If LST.ListIndex = LST.ListCount - 1 Then Exit Sub
        F = LST.Text
        LST.RemoveItem (LST.ListIndex)
        Call LST.AddItem(F, A + 1)
        LST.Selected(A + 1) = True
End If
End Sub
Public Sub Seting()
On Error GoTo 3
For i = 1 To 8
Form1.Cmd(i).Refresh: Next
If ASD.FileExists(App.Path + "\NMS\User.Dll") = False Then Exit Sub
Open App.Path + "\NMS\user.dll" For Random Access Read As 5
Get 5, , Setting
Close 5
If Setting.AudioCunt = 0 And Setting.ProgramColor = 0 And Setting.VideoCunt = 0 Then Exit Sub
With Setting
            Form1.BackColor = .ProgramColor
            Form1.Frame1.BackColor = .ProgramColor
            Form1.Frame2.BackColor = .ProgramColor
            Form1.Drive1.BackColor = .ProgramColor
            Form1.File1.BackColor = .ProgramColor
            Form1.Dir1.BackColor = .ProgramColor
            Form1.Label13.BackColor = .ProgramColor
            Form1.Label14.BackColor = .ProgramColor
            Form1.Label15.BackColor = .ProgramColor
            Form1.Label10.BackColor = .ProgramColor
            Form1.Label11.BackColor = .ProgramColor
            Form1.Label16.BackColor = .ProgramColor
            Form1.lstDIBList_1.BackColor = .ProgramColor
For i = 1 To 8
            Form1.Cmd(i).SoftBevel = .Soft
            Form1.Cmd(i).UseGreyscale = .Grey
            Form1.Cmd(i).ShowFocusRect = .Recte
            Form1.Cmd(i).ButtonType = .button.BType
            Form1.Cmd(i).CausesValidation = .Causes
            Form1.Cmd(i).CheckBoxBehaviour = .Behavi
            Form1.Cmd(i).ColorScheme = .button.BColor
            Form1.Cmd(i).SpecialEffect = .button.BEfect
            If .button.BColor = 2 Then
                Form1.Cmd(i).BackOver = .button.Custom(2)
                Form1.Cmd(i).ForeOver = .button.Custom(4)
                Form1.Cmd(i).BackColor = .button.Custom(1)
                Form1.Cmd(i).ForeColor = .button.Custom(3)
                Form1.Cmd(i).MaskColor = .button.Custom(5)
            End If
Next
End With
3:
End Sub
Public Sub Rapair()
Dim Fil As String, L%, W$, H$
Form1.lstDIBList_1.Clear
L = 1
For i = 1 To 10000
    Fil = (App.Path + "\Temp\" + Str(i) + ".bmp")
    If ASD.FileExists(Fil) = True Then
        Form1.lstDIBList.AddItem Fil
            If (i \ 2) = 0 Then
            L = L + 1
            Else: Form1.lstDIBList_1.AddItem Fil: L = 1
            End If
        Else: Exit For
    End If
Next
If Form1.lstDIBList_1.ListCount = 0 Then
ASD.DeleteFolder App.Path + "\Temp": ASD.CreateFolder App.Path + "\Temp"
End If
End Sub
Public Function TTim(Frame As Integer, EndF As Integer) As String '
Dim B1%, C1%
B1 = EndF: B1 = B1 - Frame
If B1 < 60 Then
TTim = "Time = - 00:" + Fit(B1): Exit Function
End If
C1 = B1 \ 60
B1 = B1 Mod 60
TTim = "Time = - " + Fit(C1) + ":" + Fit(B1)
End Function
Public Function Fit(A1 As Integer) As String
        If Len(Str(A1)) = 2 Then
        Fit = "0" + right(Str(A1), 1)
        Else: Fit = right(Str(A1), 2)
        End If
End Function

'*-------------Copyright By:NasserNiazyMobasser----------------*
'*-------------www.vbook.coo.ir-www.tcvb.coo.ir----------------*
'*-------------2005-2007-By:Ghayeshsoft--nasservb@gmail.com----*

