Attribute VB_Name = "modAPI"


Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetTickCount Lib "kernel32" () As Long
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_SYNC = &H0         '  play synchronously (default)

