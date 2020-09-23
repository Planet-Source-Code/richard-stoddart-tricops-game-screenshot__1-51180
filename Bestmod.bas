Attribute VB_Name = "Module2"
'Used for drawing a star pattern on the background - details below
Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'function to play .mid files
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

' Graphics functions and constants
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6    ' Masks
Public Const SRCPAINT = &HEE0086  ' onto masks
Public Const SRCCOPY = &HCC0020   ' backgrounds
Public Const SRCINVERT = &H660046 'Copies and inverts the source over the destination

'This is to hide and show the mouse, True = hide
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'This is to show and hide the taskbar
Declare Function findwindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'taskbar constants
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40

'This is used to play .wav sounds
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'sound constants
Public Const SND_SYNC = &H0 ' Don't return until sound ends (default).
Public Const SND_ASYNC = &H1 ' Return immediately after the sound starts.
Public Const SND_NODEFAULT = &H2 ' If the sound file is not found, do NOT play default sound.
Public Const SND_MEMORY = &H4 ' Play a sound from a buffer in memory.
Public Const SND_LOOP = &H8 ' Loop sound continuously (used with SND_ASYNC)
Public Const SND_NOSTOP = &H10 ' Don't stop current sound to play another.
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Function IsKeyDown(AsciiKeyCode As Byte) As Boolean
If GetKeyState(AsciiKeyCode) < -125 Then IsKeyDown = True
End Function

Public Sub s_Playsound(strName As String)
    'strName = "C:\windows\desktop\" & strName & ".wav"
    strName = App.Path & "\" & strName & ".wav"
    sndPlaySound strName, SND_ASYNC Or SND_NODEFAULT
End Sub
