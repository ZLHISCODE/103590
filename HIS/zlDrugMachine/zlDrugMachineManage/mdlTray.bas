Attribute VB_Name = "mdlTray"
Option Explicit

'---------------------------------------------------------------
'说明：Windows托盘的模块
'编制：余智勇
'---------------------------------------------------------------

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1

Private Const trayLBUTTONDOWN = 7695
Private Const trayLBUTTONUP = 7710
Private Const trayLBUTTONDBLCLK = 7725
Private Const trayRBUTTONDOWN = 7740
Private Const trayRBUTTONUP = 7755
Private Const trayRBUTTONDBLCLK = 7770
Private Const trayMOUSEMOVE = 7680

Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_LBUTTONDBLCLK = &H203

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private trayStructure As NOTIFYICONDATA

Private rc As Long
Public mblnVisible As Boolean

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Function AddIcon(pic As PictureBox, tip$)
   trayStructure.szTip = tip$ & Chr$(0)
   trayStructure.uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
   trayStructure.uID = 100
   trayStructure.cbSize = Len(trayStructure)
   
   trayStructure.hwnd = pic.hwnd
   trayStructure.uCallbackMessage = WM_MOUSEMOVE
   trayStructure.hIcon = pic.Picture
   rc = Shell_NotifyIcon(NIM_ADD, trayStructure)
End Function

Public Function DeleteIcon(pic As Control)
   trayStructure.uID = 100
   trayStructure.cbSize = Len(trayStructure)
   trayStructure.hwnd = pic.hwnd
   trayStructure.uCallbackMessage = WM_MOUSEMOVE
   rc = Shell_NotifyIcon(NIM_DELETE, trayStructure)
End Function

Public Sub TrayStatus(ByVal blnVisible As Boolean, ByVal frmMain As Form)
    mblnVisible = blnVisible
    If blnVisible Then
        If frmMain.WindowState = Val("1-最小化") Then frmMain.WindowState = 0
        frmMain.Show
    Else
        frmMain.Hide
    End If
    
    App.TaskVisible = blnVisible
End Sub
