VERSION 5.00
Begin VB.Form frmDockEx 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label lblTest 
      AutoSize        =   -1  'True
      Caption         =   "扩展部件测试卡片页签"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   450
      TabIndex        =   0
      Top             =   1020
      Width           =   3300
   End
End
Attribute VB_Name = "frmDockEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function GetInSideFunc() As String
'功能：获取窗体功能名
'开放的图标对应－－ "新增,3001|修改,3003|保存,3091|删除,3004|打开,100|撤消,3565|关闭,3021|打印,103|预览,102|退出,2613|完成,225|帮助,901|过滤,731|查找,721|设置,181|刷新,791|签名,804"
    GetInSideFunc = "新增,3001|修改,3003|保存,3091|删除,3004|打开,100|撤消,3565|关闭,3021|打印,103|预览,102|退出,2613|完成,225|帮助,901|过滤,731|查找,721|设置,181|刷新,791|签名,804"
End Function

Public Function ExecuteFunc(ByVal strName As String) As Boolean
'功能：执行菜单上的功能。参数可自行添加
'参数：strName 功能名称
    If Not Me.Visible Then Exit Function
    Debug.Print "zlPlugIn/frmDockEx/ExecuteFunc " & strName & "功能被执行！！！"
    ExecuteFunc = True
End Function

Public Sub RefreshInSide()
'功能：窗体刷新。参数可自行添加
    Debug.Print "zlPlugIn/frmDockEx/RefreshInSide 当前卡片被刷新！！！"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "卸载了！！！！！"
End Sub
