VERSION 5.00
Begin VB.Form frmMipClient 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "frmMipClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Public Function Initialize() As Boolean
    '******************************************************************************************************************
    '功能：初始化
    '参数：无
    '返回：初始化成功返回True,否则返回False
    '******************************************************************************************************************
    
'    Call InitCommandBar
'    Call AddIcon(picNotify.hwnd, imgIcon(0).Picture, "消息服务平台客户端收发服务")
            
    Initialize = True
    
End Function

'Private Function InitCommandBar() As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim objMenu As CommandBarPopup
'    Dim objBar As CommandBar
'    Dim objPopup As CommandBarPopup
'    Dim objControl As CommandBarControl
'    Dim cbrCustom As CommandBarControlCustom
'
'    '------------------------------------------------------------------------------------------------------------------
'    '初始设置
'    cbsMain.VisualTheme = xtpThemeOffice2003
'
'    With cbsMain.Options
'        .ShowExpandButtonAlways = False
'        .ToolBarAccelTips = True
'        .AlwaysShowFullMenus = False
'        '.UseFadedIcons = True '放在VisualTheme后有效
'        .IconsWithShadow = True '放在VisualTheme后有效
'        .UseDisabledIcons = True
'        .LargeIcons = True
'        .SetIconSize True, 24, 24
'        .SetIconSize False, 16, 16
'    End With
'    cbsMain.EnableCustomization False
'
'    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
'    cbsMain.Options.LargeIcons = True
'
'    '------------------------------------------------------------------------------------------------------------------
'    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值
'
'    cbsMain.ActiveMenuBar.Title = "菜单"
'    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
'    cbsMain.ActiveMenuBar.Visible = False
'
'End Function
'
'Public Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim cbrPopupBar As CommandBar
'    Dim cbrPopupItem As CommandBarControl
'    Dim cbrPopupItem2 As CommandBarControl
'    Dim cbrMenuBar As CommandBarControl
'    Dim cbrControl As CommandBarControl
'    Dim cbrControl2 As CommandBarControl
'
'    '弹出菜单处理
'
'    On Error GoTo errHand
'
'    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
'
'    Select Case bytPlace
'    Case 1
'
'        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 1, "运行日志(&L)")
'        cbrPopupItem.DefaultItem = True
'
'        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 2, "启动服务(&S)")
'        cbrPopupItem.BeginGroup = True
'
'        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 3, "停止服务(&T)")
'
'    End Select
'
'    Set ShowConetneMenu = cbrPopupBar
'
'    Exit Function
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    '
'End Function
'
'Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    Select Case Control.id
'    Case 1
'
'    Case 2
'        MsgBox "启动服务"
'    Case 3
'        MsgBox "停止服务"
'    End Select
'End Sub

Private Sub Form_Unload(Cancel As Integer)
        
'    Call RemoveIcon(picNotify.hwnd)
    
End Sub

'Private Sub picNotify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '--------------------------------------------------------------------------------------------------
'    '功能:  处理picNotify的各种处理事件,主要用于自动提醒相关功能(陈渝编写)
'    '--------------------------------------------------------------------------------------------------
'
'    Select Case Hex(X) '
'        Case "1E3C"     'Right-Button-Down
'        Case "1E4B"     'Right-Button-Up
'            Call ShowConetneMenu(1).ShowPopup
'        Case "1830"     'Right-Button-Down LARGE FONTS '
'        Case "1E1E"     'Left-Button-up
'        Case "1E0F"     'Left-Button-Down '
'        Case "1E2D"     'Left-Button-Double-Click '
'            '
'        Case "1824"     'Left-Button-Double-Click LARGE FONTS
'        Case "1E5A"     'Right-Button-Double-Click '
'    End Select '
'End Sub
