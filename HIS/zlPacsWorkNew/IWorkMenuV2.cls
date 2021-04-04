VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWorkMenuV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function zlBaseMenuID() As Long

End Function

Public Function zlExecuteCmd(ByVal lngCmdType As Long)
'执行菜单命令

End Function

Public Function zlIsModuleMenu(ByVal strModuleName As String, objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单

End Function


Public Sub zlCreateMenu(ByVal strModuleName As String, objMenuBar As Object)
'创建影像记录对应的菜单
'objMenuBar为CommandBars的菜单对象

End Sub


Public Sub zlCreateToolBar(ByVal strModuleName As String, objToolBar As Object)
'创建工具栏

End Sub


Public Sub zlClearMenu(ByVal strModuleName As String)
'清除所创建的菜单

End Sub


Public Sub zlClearToolBar(ByVal strModuleName As String)
'清除创建的工具栏

End Sub


Public Sub zlExecuteMenu(ByVal strModuleName As String, ByVal lngMenuId As Long)
'根据菜单ID执行对应功能

End Sub


Public Sub zlUpdateMenu(ByVal strModuleName As String, Control As XtremeCommandBars.ICommandBarControl)
'更新菜单

End Sub


Public Sub zlPopupMenu(ByVal strModuleName As String, objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单

End Sub


Public Sub zlRefreshSubMenu(ByVal strModuleName As String, objCommandBar As Object)
'刷新弹出的子菜单
'objCommandBar为选中的某个菜单对象

End Sub





