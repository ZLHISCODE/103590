Attribute VB_Name = "mdlPublic"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gobjComLib As Object
Public gobjPlugIn As Object
Public gobjReport As Object

'---------------------------------------------------------------------------------------------------------
Public gstrHwndOLD As String
Public glngPid As Long
Public gblnFinded As Boolean

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim strClass As String * 200 '窗口类名
    Dim strTitle As String * 200
    Dim lngPid As Long
 
    On Error Resume Next
    
    GetWindowText hwnd, strTitle, 200
    GetClassName hwnd, strClass, 200
    
    'ThunderRT6FormDC
    'WindowsForms10.Window.8.app.0.13965fa_r6_ad1
    If InStr(UCase(GetStr(strClass)), UCase("Form")) > 0 _
        And GetStr(strTitle) <> "" And InStr(gstrHwndOLD, "," & hwnd & ",") = 0 Then
        
        Call GetWindowThreadProcessId(hwnd, lngPid)
        
        If lngPid = glngPid And lngPid <> 0 Then
            If IsWindowVisible(hwnd) Then
                'MsgBox GetStr(strClass) & "," & GetStr(strTitle)
                
                SetWindowPos hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
                SetWindowPos hwnd, -2, 0, 0, 0, 0, &H1 Or &H2
                
                BringWindowToTop hwnd
                SetForegroundWindow hwnd
                SetActiveWindow hwnd
                
                'SetWindowsInTaskBar hwnd, True
                
                gblnFinded = True
            End If
        End If
    End If
    EnumChildProc = 1
End Function

Public Function EnumChildProcOld(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim strClass As String * 200
    Dim lngPid As Long
    
    On Error Resume Next
    GetClassName hwnd, strClass, 200

    If InStr(UCase(GetStr(strClass)), UCase("Form")) > 0 Then
        Call GetWindowThreadProcessId(hwnd, lngPid)

        If lngPid = glngPid And lngPid <> 0 Then
             gstrHwndOLD = gstrHwndOLD & "," & hwnd & ","
        End If
    End If
    EnumChildProcOld = 1
End Function

Public Function GetStr(ByVal szString As String) As String
    Dim lngZero As Long
    
    lngZero = InStr(szString, Chr(0))
    
    If lngZero > 0 Then
        GetStr = Left(szString, lngZero - 1)
    Else
        GetStr = szString
    End If
End Function

Public Sub SetWindowsInTaskBar(ByVal lnghwnd As Long, ByVal blnShow As Boolean)
'功能：设置窗体是否在任务条上显示
    Dim lngStyle As Long
    
    lngStyle = GetWindowLong(lnghwnd, GWL_EXSTYLE)
    If blnShow Then
        lngStyle = lngStyle Or &H40000
    Else
        lngStyle = lngStyle And Not &H40000
    End If
    Call SetWindowLong(lnghwnd, GWL_EXSTYLE, lngStyle)
End Sub
