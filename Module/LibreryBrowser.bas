Attribute VB_Name = "Librery"
'------Typing New data For Propertis File---------------------
Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters  As String
    lpDirectory   As String
    nShow As Long
    hInstApp As Long
    lpIDList      As Long
    lpClass       As String
    hkeyClass     As Long
    dwHotKey      As Long
    hIcon         As Long
    hProcess      As Long
End Type
'------Typing New data For Seearch File---------------------
Public Type BrowseInfo
    hWndOwner       As Long
    pIDLRoot        As Long
    pszDisplayName  As Long
    lpszTitle       As Long
    ulFlags         As Long
    lpfnCallback    As Long
    lParam          As Long
    iImage          As Long
End Type
'---------------Conset For Seearch--------------------
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Const ATTR_NORMAL = 0
Public Const ATTR_READONLY = 1
Public Const ATTR_HIDDEN = 2
Public Const ATTR_SYSTEM = 4
Public Const ATTR_VOLUME = 8
Public Const ATTR_DIRECTORY = 16
Public Const ATTR_ARCHIVE = 32
'-----------------------Declareing API------------------------------------------
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
        ByVal lpBuffer As String) As Long '-------------------------------------
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
        ByVal lpString2 As String) As Long '------------------------------------
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Declare Function ShellExecuteEX Lib "shell32.dll" Alias _
        "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long '--------------------
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long
'----------------------------------------------------------------------------------
Public StartingAddress As String
Public mbDontNavigateNow As Boolean

Public Sub SetingBrw(Frm As frmBrowser)
On Error GoTo 3
For i = 1 To 10
Frm.cmd(i).Refresh: Next
If Setting.AudioCunt = 0 And Setting.ProgramColor = 0 And Setting.VideoCunt = 0 Then Exit Sub
With Setting
For i = 1 To 10
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
End With
With Frm
    For i = 1 To 4: .Frame(i).BackColor = Setting.ProgramColor: Next
     .Lvfiles.BackColor = Setting.ProgramColor: .Drive1.BackColor = Setting.ProgramColor
     .Text1.BackColor = Setting.ProgramColor: .Label1.BackColor = Setting.ProgramColor
     .Text2.BackColor = Setting.ProgramColor: .cboAddress.BackColor = Setting.ProgramColor
     .lblAddress.BackColor = Setting.ProgramColor: .SSTab1.BackColor = Setting.ProgramColor
     .Label2.BackColor = Setting.ProgramColor: .Label3.BackColor = Setting.ProgramColor
     .File1.BackColor = Setting.ProgramColor: .Dir1.BackColor = Setting.ProgramColor: .Drive2.BackColor = Setting.ProgramColor
End With
3
End Sub
Public Function ShowFileProperties(filename As String, OwnerhWnd As Long) As Long
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    ShellExecuteEX SEI
    ShowFileProperties = SEI.hInstApp
End Function
Public Function FindFiles(Path As String, SearchStr As String, _
FileCount As Integer, DirCount As Integer)
    Dim filename As String
    Dim DirName As String
    Dim dirNames() As String
    Dim nDir As Integer
    Dim i As Integer
On Error GoTo sysFileERR
If Right(Path, 1) <> "\" Then Path = Path & "\"
nDir = 0 '-------------------------------------------
ReDim dirNames(nDir)
DirName = Dir(Path, vbDirectory Or vbHidden)
Do While Len(DirName) > 0
    If (DirName <> ".") And (DirName <> "..") Then
        If GetAttr(Path & DirName) And vbDirectory Then
        dirNames(nDir) = DirName
        DirCount = DirCount + 1
        nDir = nDir + 1
        ReDim Preserve dirNames(nDir)
        End If
sysFileERRCont:
    End If
    DirName = Dir(): DoEvents
Loop '-------------------------------------------
filename = Dir(Path & SearchStr, vbNormal Or vbHidden Or vbSystem _
Or vbReadOnly)
While Len(filename) <> 0
    FindFiles = FindFiles + FileLen(Path & filename)
    FileCount = FileCount + 1
    frmBrowser.Lvfiles.AddItem Path & filename
    filename = Dir(): DoEvents
Wend '---------------------------------------
If nDir > 0 Then
For i = 0 To nDir - 1
    FindFiles = FindFiles + FindFiles(Path & dirNames(i) & "\", _
    SearchStr, FileCount, DirCount): DoEvents
Next i
End If '------------------------
AbortFunction:
Exit Function
sysFileERR:
If Right(DirName, 4) = ".sys" Then
Resume sysFileERRCont
Else '----------------------------------
MsgBox "Error: " & Err.Number & " - " & Err.Description, , _
"Unexpected Error"
Resume AbortFunction
End If
End Function

Public Sub Savelis(OutPath As String)
On Error Resume Next '--------------------------------------------------
                    Dim T3 As String, T2, strans As String, L As Single, i As Integer
                    T3 = "": T2 = ""
                    If Lvfiles.List(1) = "" Then
                    strans = MsgBox("File Not Found!", vbCritical)
                    Exit Sub '------------------------------------------------------
                    End If
                    If UCase(Right(OutPath, 3)) <> "M3U" Then Exit Sub
            Open OutPath For Output As #1
                    Print #1, "#EXTM3U:"
                For i = 1 To Lvfiles.ListCount '----------------------------
                    Print #1, "#EXTNIF:"
                    Print #1, Lvfiles.List(i)
                Next i '------------------------------------------------------
            Close #1
End Sub

