VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private MFile  As String 'local copy
Private Wmp1   As WindowsMediaPlayer
Private Vid    As Boolean
Public Property Let MediaFile(ByVal Media As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Video = 5
    MFile = Media
End Property

Public Property Get MediaFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Video
    MediaFile = MFile
End Property
Public Property Let Wmp(ByVal Wmpl As WindowsMediaPlayer)
Wmp1 = Wmpl
End Property
Public Property Get Wmp() As WindowsMediaPlayer
Wmp = Wmp1
End Property
'Public Property Let Video(ByVal Bol As Boolean)
'Vid = Bol
'End Property
Public Property Get Video() As Boolean
Video = Vid
End Property
Public Sub PlayFile(UrlLabel As Label, PlaLabel As Label, Vhs As Boolean)
'-------------------Start Play File-----------------------
            UrlLabel.Caption = MediaFile
            VideoMode UrlLabel
            If Video = True Then
            If Vhs = True Then '------------sho Video---------------
            Form3.Show
            If Form3.WindowsMediaPlayer1.URL = MediaFile Then Exit Sub
            Wmp.URL = "": Form3.WindowsMediaPlayer1.settings.autoStart = True
            Form3.WindowsMediaPlayer1.URL = MediaFile
            Else: Unload Form3: Wmp.settings.autoStart = True
2           If Wmp.URL = MediaFile Then Exit Sub
            Wmp.URL = MediaFile
            End If
            Else: GoTo 2 '---------Audio File----------------------------
            End If
End Sub
Private Function VideoMode(UrlLabel As Label) As Boolean   'Proses File For Video Or Audio
On Error Resume Next '-----------------------------------------
                    Dim AFD As String
                    If MediaFile = "" Then Exit Function
                    Select Case UCase(Right$(MediaFile, 3))
                    Case "MPG": AFD = "OK"
                    Case "AVI": AFD = "OK"
                    Case "DAT": AFD = "OK"
                    Case "VCD": AFD = "OK"
                    Case "IFO": AFD = "OK"
                    Case "MOV": AFD = "OK"
                    Case "WMV": AFD = "OK"
                    Case "M3U": AFD = "NO"
                    End Select
If AFD = "OK" Then
    UrlLabel.Caption = "Video": VideoMode = True: Video = True
Else: VideoMode = False: Video = False: End If
If AFD = "" Then UrlLabel.Caption = "For"
If AFD = "No" Then UrlLabel.Caption = "M3u"
End Function


