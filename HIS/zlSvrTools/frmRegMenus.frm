VERSION 5.00
Begin VB.Form frmRegMenus 
   Caption         =   "用户注册附加窗体"
   ClientHeight    =   2400
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   2895
   Icon            =   "frmRegMenus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   2895
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu TrackMenu 
      Caption         =   "弹出菜单"
      Begin VB.Menu MnuDeleteCur 
         Caption         =   "删除当前日志(&D)"
      End
      Begin VB.Menu MnuDeleteAll 
         Caption         =   "删除所有日志(&A)"
      End
   End
End
Attribute VB_Name = "frmRegMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Bln日志 As Boolean                               '为真,运行日志调用;否则错误日志调用
Public FrmObj As Form

Private Sub MnuDeleteAll_Click()
    Call DeleteAllLog(FrmObj, Bln日志)
End Sub

Private Sub MnuDeleteCur_Click()
    Call DeleteCurLog(FrmObj, Bln日志)
End Sub
