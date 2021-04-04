VERSION 5.00
Begin VB.Form frmQueryDisplay 
   Caption         =   "查询配置效果预览"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11100
   Icon            =   "frmQueryDisplay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   11100
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "frmQueryDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OnFormQueryUnload()
Public Event OnTestFormResize()

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    RaiseEvent OnFormQueryUnload
End Sub

Private Sub Form_Resize()
    RaiseEvent OnTestFormResize
End Sub
