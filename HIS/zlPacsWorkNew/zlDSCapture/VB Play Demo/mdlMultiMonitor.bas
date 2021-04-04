Attribute VB_Name = "mdlMultiMonitor"
Option Explicit

  Const MONITORINFOF_PRIMARY = &H1
  Const MONITOR_DEFAULTTONEAREST = &H2
  Const MONITOR_DEFAULTTONULL = &H0
  Const MONITOR_DEFAULTTOPRIMARY = &H1
  
  Private Type RECT
          Left   As Long
          Top   As Long
          Right   As Long
          Bottom   As Long
  End Type
  
  '显示器信息
  Private Type MONITORINFO
          cbSize   As Long
          rcMonitor   As RECT
          rcWork   As RECT
          dwFlags   As Long
  End Type
  
  
  Private Type POINT
          x   As Long
          y   As Long
  End Type
  
  Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
  Private Declare Function MonitorFromPoint Lib "user32.dll" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
  Private Declare Function MonitorFromRect Lib "user32.dll" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
  Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
  Public Declare Function EnumDisplayMonitors Lib "user32.dll" (ByVal hdc As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
  Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
  
  
  Public monitor() As Long
  
  
  Private Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, ByVal dwData As Long) As Long
          Dim MI     As MONITORINFO, R       As RECT
          
          ReDim Preserve monitor(UBound(monitor) + 1)
          
          monitor(UBound(monitor) - 1) = hMonitor

'          'initialize   the   MONITORINFO   structure
'          MI.cbSize = Len(MI)
'          'Get   the   monitor   information   of   the   specified   monitor
'          GetMonitorInfo hMonitor, MI
'          'write   some   information   on   teh   debug   window
'
'          'check   whether   Form1   is   located   on   this   monitor
'          If MonitorFromWindow(Form1.hwnd, MONITOR_DEFAULTTONEAREST) = hMonitor Then
'                  Debug.Print "Form1   is   located   on   this   monitor"
'          End If
'          'heck   whether   the   point   (0,   0)   lies   within   the   bounds   of   this   monitor
'          If MonitorFromPoint(0, 0, MONITOR_DEFAULTTONEAREST) = hMonitor Then
'                  Debug.Print "The   point   (0,   0)   lies   wihthin   the   range   of   this   monitor..."
'          End If
'          'check   whether   Form1   is   located   on   this   monitor
'          GetWindowRect Form1.hwnd, R
'          If MonitorFromRect(R, MONITOR_DEFAULTTONEAREST) = hMonitor Then
'                  Debug.Print "The   rectangle   of   Form1   lies   within   this   monitor"
'          End If
'          Debug.Print ""
'          'Continue   enumeration

          '这里必须返回1，以便后续执行
          MonitorEnumProc = 1
  End Function
  
  
  Public Function GetMonitorIndex(ByVal windowHandle As Long) As Long
    
    Dim i As Integer
    
    Dim monitorCount As Integer
    monitorCount = 0
    
    On Error GoTo GetMonitorInf
      monitorCount = UBound(monitor)
GetMonitorInf:
      If monitorCount <= 1 Then
        ReDim Preserve monitor(1)
        monitor(1) = -1
  
        EnumDisplayMonitors ByVal 0&, ByVal 0&, AddressOf MonitorEnumProc, ByVal 0&
      End If
    
        
    For i = 1 To UBound(monitor)
      If MonitorFromWindow(windowHandle, MONITOR_DEFAULTTONEAREST) = monitor(i) Then
        GetMonitorIndex = i - 1
        Exit Function
      End If
    Next i
    
    GetMonitorIndex = -1
    
  End Function
  
  
      
  
  
  
  

















