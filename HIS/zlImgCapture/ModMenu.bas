Attribute VB_Name = "modMenu"
Option Explicit
Public Const ID_Capture_CapPicture = 300
Public Const ID_Capture_CapSaveVido = 301
Public Const ID_Capture_SavePicture = 302
Public Const ID_Capture_DelPicture = 303
Public Const ID_Capture_Setup = 304
Public Const ID_Capture_Setup_Format = 305
Public Const ID_Capture_Setup_Source = 306
Public Const ID_Capture_Setup_Driver = 307
Public Const ID_Capture_Prot = 308
Public Const ID_Capture_Exit = 309
Public Sub CreateMenu(ToolBars As Object)
    '------------------------------------------------
    '功能：                                  创建菜单
    '参数：
    '           IconX                        设置图标X大小
    '           IconY                        设置图标Y大小
    '返回：                                  无
    '上级函数或过程：                        frViewer_load
    '下级函数或过程：                        无
    '引用的外部参数：                        无
    '编制人：                                曾超 2005-6-27
    '------------------------------------------------
    Dim Control As CommandBarControl
    Dim ControlFile As CommandBarPopup
    Dim ControlSelect As CommandBarPopup
    Dim ToolBar As CommandBar
    Dim ControlPopup As CommandBarPopup
    
    '去掉扩展按钮
    ToolBars.Options.ShowExpandButtonAlways = False
    ToolBars.ActiveMenuBar.EnableDocking xtpFlagHideWrap
    
    '去掉菜单
    ToolBars.Item(1).Visible = False
    
    
    Set ToolBar = ToolBars.Add("主工具栏", xtpBarBottom)
    
    With ToolBar.Controls
        .Add xtpControlButton, ID_Capture_CapPicture, "采集"
'        .Add xtpControlButton, ID_Capture_CapSaveVido, "录像"
        .Add xtpControlButton, ID_Capture_SavePicture, "保存"
        .Add xtpControlButton, ID_Capture_DelPicture, "删除"
        
        Set ControlPopup = .Add(xtpControlSplitButtonPopup, ID_Capture_Setup, "设置")
        ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Capture_Setup_Format, "格式"
        ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Capture_Setup_Source, "来源"
        
        .Add xtpControlButton, ID_Capture_Prot, "端口"
        
        .Add xtpControlButton, ID_Capture_Exit, "退出"
    End With
    
    ToolBar.Position = xtpBarTop
    ToolBar.SetIconSize 24, 24
    ToolBar.ShowTextBelowIcons = True
    
    
End Sub


