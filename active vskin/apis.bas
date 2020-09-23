Attribute VB_Name = "apis"
Public Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long


Public Declare Function SetWindowPos& Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


Const conHwndTopmost = -1

 
Public Const lport = 1300
Public Const cport = 1350
Public Const cip = "127.0.0.1"
Public cuser As String
Public userr As String
Public passs As String
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Type TypeRGB 'This user type is used to convert long decimals to Bytes
    R As Byte
    G As Byte
    b As Byte
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public doit As Boolean
Option Explicit

Dim hwndDest            As Long
Dim mlWidth             As Long
Dim mlHeight            As Long
Dim mbFormSwitch        As Boolean
Dim mbWindows98orHigher As Boolean
'
' Constants used by FlashWindowEx/FLASHWINFO
'
'Stop flashing
Const FLASHW_STOP = &H0
'Flash the caption
Const FLASHW_CAPTION = &H1
'Flash the taskbar button.
Const FLASHW_TRAY = &H2
'Flash both
Const FLASHW_ALL = FLASHW_TRAY Or FLASHW_CAPTION
'Flash continuously until FLASHW_STOP is set
Const FLASHW_TIMER = &H4
'Flash continuously until window comes to foreground
Const FLASHW_TIMERNOFG = &HC

Const VER_PLATFORM_WIN32_WINDOWS = 1

Public Type FLASHWINFO
   cbSize    As Long
   hwnd      As Long
   dwFlags   As Long
   uCount    As Long
   dwTimeout As Long
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Dim XY() As POINTAPI

Public Declare Function GetVersionEx Lib "kernel32" Alias _
    "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function CreateEllipticRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long
    
Public Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, ByVal nCount As Long, _
    ByVal nPolyFillMode As Long) As Long

Public Declare Function SetWindowRgn Lib "User32" _
    (ByVal hwnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Long) As Long
    
Public Declare Function FlashWindow Lib "User32" _
    (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Public Declare Function FlashWindowEx Lib "User32" ( _
    fwi As FLASHWINFO) As Boolean

Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function GetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long



Public Const GWL_EXSTYLE = (-20)
Public Enum TransType
  LWA_OPAQUE = 0
  LWA_COLORKEY = 1
  LWA_ALPHA = 2
End Enum
Public Const RGN_OR = 2
Public Const WS_EX_LAYERED = &H80000
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Const SRCCOPY = &HCC0020
Public Declare Function BitBlt Lib "gdi32" (ByVal hdestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function RedrawWindow Lib "User32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Const WM_SETREDRAW = &HB
Public Const RDW_INVALIDATE = &H1
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_UPDATENOW = &H100
Public Const WM_PAINT = &HF

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const WM_NCLBUTTONUP = &HA2
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

