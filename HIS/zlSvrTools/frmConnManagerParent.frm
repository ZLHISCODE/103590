VERSION 5.00
Begin VB.Form frmConnManagerParent 
   Caption         =   "数据连接"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8415
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   8415
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmConnManagerParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SW_HIDE = 0
Private Const SW_SHOWMAXIMIZED = 3


Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Private Sub Form_Load()
    Call frmConnectionsManager.ShowMe(Me, False)
End Sub

Private Sub Form_Resize()
    ShowWindow frmConnectionsManager.hwnd, SW_HIDE
    ShowWindow frmConnectionsManager.hwnd, SW_SHOWMAXIMIZED
End Sub
