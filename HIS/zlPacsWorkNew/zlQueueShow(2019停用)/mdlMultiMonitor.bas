Attribute VB_Name = "mdlMultiMonitor"
Option Explicit



  
Public gmonitors() As Monitorinfos


Private Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, ByVal dwData As Long) As Long
     Dim monitorInf As MONITORINFO
     Dim R As RECT
     
     ReDim Preserve gmonitors(UBound(gmonitors) + 1)
     
     'initialize   the   MONITORINFO   structure
     monitorInf.cbSize = Len(monitorInf)
     'Get   the   monitor   information   of   the   specified   monitor
     GetMonitorInfo hMonitor, monitorInf
     'write   some   information   on   teh   debug   window

    
     gmonitors(UBound(gmonitors) - 1).monitorHandle = hMonitor
     gmonitors(UBound(gmonitors) - 1).monitorInf = monitorInf
     
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
'获取当前窗口所在显示器索引
Dim i As Integer
    
    Call InitMonitor
    
    For i = 1 To UBound(gmonitors)
        If MonitorFromWindow(windowHandle, MONITOR_DEFAULTTONEAREST) = gmonitors(i).monitorHandle Then
            GetMonitorIndex = i - 1
            Exit Function
        End If
    Next i

    GetMonitorIndex = -1
  
End Function


Public Sub InitMonitor()
'初始化监视器设置
Dim monitorCount As Integer
    
    monitorCount = 0

On Error GoTo GetMonitorInf
    monitorCount = UBound(gmonitors)
    
GetMonitorInf:
    If monitorCount <= 1 Then
        ReDim Preserve gmonitors(1)
        gmonitors(1).monitorHandle = -1

        EnumDisplayMonitors ByVal 0&, ByVal 0&, AddressOf MonitorEnumProc, ByVal 0&
    End If
End Sub
  

Public Sub SetFullScreenWindow(ByVal objWindow As Object, ByVal lngMonitorIndex As Long)
'全屏设置窗口大小

    '取得屏幕的相对位置
    objWindow.Left = objWindow.ScaleX(gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Left, vbPixels, vbTwips)
    objWindow.Top = objWindow.ScaleY(gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Top, vbPixels, vbTwips)

    '取得屏幕的大小
    objWindow.Width = objWindow.ScaleX(gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Right - gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Left, vbPixels, vbTwips)  'Screen.Width
    objWindow.Height = objWindow.ScaleY(gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Bottom - gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Top, vbPixels, vbTwips) 'Screen.Height
End Sub

Public Sub SetCustomWindow(ByVal objWindow As Object, ByVal lngMonitorIndex As Long, trLCDRect As TRect)
'自定义设置窗口大小

     '取得屏幕的相对位置
    objWindow.Left = objWindow.ScaleX(gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Left, vbPixels, vbTwips) + trLCDRect.lngLeft
    objWindow.Top = objWindow.ScaleY(gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Top, vbPixels, vbTwips) + trLCDRect.lngTop
    
    '取得屏幕的大小
    objWindow.Width = trLCDRect.lngWidth       'Screen.Width
    objWindow.Height = trLCDRect.lngHeight     'Screen.Height
End Sub















