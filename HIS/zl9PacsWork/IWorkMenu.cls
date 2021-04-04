VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWorkMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function zlIsModuleMenu(ByVal objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单

End Function


Public Sub zlCreateMenu(objMenuBar As Object)
'创建影像记录对应的菜单
'objMenuBar为CommandBars的菜单对象

End Sub


Public Sub zlCreateToolBar(objToolBar As Object)
'创建工具栏
End Sub


Public Sub zlClearMenu()
'清除所创建的菜单

End Sub


Public Sub zlClearToolBar()
'清除创建的工具栏

End Sub


Public Sub zlExecuteMenu(ByVal lngMenuId As Long)
'根据菜单ID执行对应功能

End Sub


Public Sub zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
'更新菜单

End Sub


Public Sub zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单

End Sub


Public Sub zlRefreshSubMenu(objCommandBar As Object)
'刷新弹出的子菜单
'objCommandBar为选中的某个菜单对象

End Sub





