Attribute VB_Name = "mdlPubCommandBars"
Option Explicit
'CommandBars控件常用功能封装模块

Public Function CbsSetting(ByRef cbsMain As CommandBars)
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.ActiveMenuBar.ContextMenuPresent = False    '禁止右键选择工具栏来取消
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap  '禁止移动工具栏
End Function

Public Sub CbsButtonInit(ByRef cbsMain As CommandBars, Buttons As Collection, _
                         Optional blnLargeIcons As Boolean = False, _
                         Optional Position As XTPBarPosition)
    '创建工具栏菜单
    'cbsMain :工具栏对象
    'Buttons :菜单集合,每个元素的格式为 菜单id,标题,是否分组
    'blnLargeIcons :是否大图标
    'Position      :菜单位置
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim strButton As Variant
    Dim varButton As Variant

    Call CbsSetting(cbsMain)
    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.ActiveMenuBar
    cbsMain.Options.LargeIcons = blnLargeIcons  '小图标
    objBar.Position = Position   '工具栏在顶部

    For Each strButton In Buttons
        varButton = Split(strButton, ",")
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, Val(varButton(0)), varButton(1))     '固有
            objControl.STYLE = xtpButtonIconAndCaption
            If UCase(varButton(2)) = "TRUE" Then objControl.BeginGroup = True '固有
        End With
    Next
    cbsMain.RecalcLayout
End Sub


'Public Sub CbsResize3(ByRef cbsMain As CommandBars, ByRef frmParan As Form, _
'                     ByRef objLeft As Control, _
'                     ByRef objWE As Control, _
'                     ByRef objNS As Control, _
'                     ByRef objRightTop As Control, _
'                     ByRef objRightBottom As Control, _
'                     ByVal minWidth As Single, minHeight As Single, _
'                     Optional LeftExpend As Boolean, _
'                     Optional TopExpend As Boolean _
'                     )
'    'cbsMain        :工具栏
'    'frmParan       :窗体
'    'objLeft        :左边控件
'    'objWE          :左右分隔条
'    'objNS          :上下分隔条
'    'objRightTop    :右上
'    'objRightButton :右下
'    'minWidth       :控件最小宽度
'    'minHeight      :控件最小高度
'    'LeftExpend     :初始化时，左边控件拉伸还是右边控件拉伸，默认为右边控件
'    'TopExpend      :初始化时，上边控件拉伸还是下边控件拉伸，默认为下边控件
'    '窗体分为三部分的情况下可以调用
'    '-------------------------------
'    ' 左  ｜ 右上                  |
'    '     ｜-----------------------|
'    '     ｜ 右下                  |
'    '-------------------------------
'
'    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
'    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
'    On Error Resume Next
'
'
'    If frmParan.Width < minWidth * 2 + 100 Then frmParan.Width = minWidth * 2 + 100
'    If frmParan.Height < minHeight * 2 + 100 Then frmParan.Height = minHeight * 2 + 100
'    objWE.MousePointer = 9 '左右形
'    objNS.MousePointer = 7 '上下形
'    objWE.BorderStyle = 0   '无边框
'    objNS.BorderStyle = 0
'    objWE.Caption = ""
'    objNS.Caption = ""
'    '初始化分隔条颜色
'    objWE.BackColor = frmParan.BackColor
'    objNS.BackColor = frmParan.BackColor
'
'    If LeftExpend Then
'        With objLeft
'            .Top = lngTop
'            .Width = minWidth
'            .Left = lngLeft
'            .Height = lngBottom - lngTop
'        End With
'        With objRightTop
'            .Top = lngTop
'            If TopExpend Then
'                .Height = lngBottom - lngTop - objNS.Height - minHeight
'            Else
'                .Height = minHeight
'            End If
'            .Width = (lngRight - lngLeft) - objLeft.Width - objWE.Width
'            .Left = lngRight - lngLeft - .Width
'        End With
'    Else
'        With objRightTop
'            .Top = lngTop
'            .Width = minWidth
'            .Left = lngRight - lngLeft - .Width
'            If TopExpend Then
'                .Height = lngBottom - lngTop - objNS.Height - minHeight
'            Else
'                .Height = minHeight
'            End If
'        End With
'        With objLeft
'            .Left = lngLeft
'            .Top = lngTop
'            .Height = lngBottom - lngTop
'            .Width = (lngRight - lngLeft) - objRightTop.Width - objWE.Width
'        End With
'    End If
'    With objWE
'        .Left = objLeft.Left + objLeft.Width
'        .Top = objLeft.Top
'        .Height = objLeft.Height
'        .Width = 45
'    End With
'    With objNS
'        .Left = objWE.Left + objWE.Width
'        .Width = objRightTop.Width
'        .Height = 45
'        .Top = objRightTop.Top + objRightTop.Height
'    End With
'    With objRightBottom
'        .Left = objWE.Left + objWE.Width
'        .Top = objNS.Top + objNS.Height
'        .Width = lngRight - lngLeft - .Left
'        .Height = lngBottom - lngTop - .Top
'    End With
'
'End Sub

'Public Sub cbsSubResize(vfg As VSFlexGrid, cbsSub As CommandBars)
'    '一个表格一个工具栏的页面的Resize
'    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
'    Call cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
'
'    With vfg
'        .Left = lngLeft + 45
'        .Top = lngTop
'        .Height = lngBottom - lngTop - 45
'        .Width = lngRight - lngLeft - 90
'    End With
'End Sub

'Public Sub CbsResize2(ByRef cbsMain As CommandBars, ByRef frmParan As Form, _
'                     ByRef objLeft As Control, _
'                     ByRef objWE As Control, _
'                     ByRef objRight As Control, _
'                     ByVal minWidth As Single, _
'                     Optional LeftExpend As Boolean _
'                     )
'    'cbsMain        :工具栏
'    'frmParan       :窗体
'    'objLeft        :左边控件
'    'objWE          :左右分隔条
'    'objRight       :右
'    'minWidth       :控件最小宽度
'    'LeftExpend     :初始化时，左边控件拉伸还是右边控件拉伸，默认为右边控件
'    '窗体分为两部分的情况下可以调用
'    '-------------------------------
'    ' 左  ｜ 右                    |
'    '     ｜                       |
'    '     ｜                       |
'    '-------------------------------
'    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
'    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
'    On Error Resume Next
'
'    If frmParan.Width < minWidth * 2 + 100 Then frmParan.Width = minWidth * 2 + 100
'    objWE.MousePointer = 9 '左右形
'    objWE.BorderStyle = 0   '无边框
'    objWE.Caption = ""
'    '初始化分隔条颜色
'    objWE.BackColor = frmParan.BackColor
'
'    If LeftExpend Then
'        With objLeft
'            .Top = lngTop
'            .Width = minWidth
'            .Left = lngLeft
'            .Height = lngBottom - lngTop
'        End With
'        With objRight
'            .Top = lngTop
'            .Height = lngBottom - lngTop
'            .Width = (lngRight - lngLeft) - objLeft.Width - objWE.Width
'            .Left = lngRight - lngLeft - .Width
'        End With
'    Else
'        With objRight
'            .Top = lngTop
'            .Width = minWidth
'            .Left = lngRight - lngLeft - .Width
'            .Height = lngBottom - lngTop
'        End With
'        With objLeft
'            .Left = lngLeft
'            .Top = lngTop
'            .Height = lngBottom - lngTop
'            .Width = (lngRight - lngLeft) - objRight.Width - objWE.Width
'        End With
'    End If
'    With objWE
'        .Left = objLeft.Left + objLeft.Width
'        .Top = objLeft.Top
'        .Height = objLeft.Height
'        .Width = 45
'    End With
'
'End Sub





