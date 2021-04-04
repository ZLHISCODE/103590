Attribute VB_Name = "mdlPubCommandBars"
Option Explicit
'CommandBars�ؼ����ù��ܷ�װģ��

Public Function CbsSetting(ByRef cbsMain As CommandBars)
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
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
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.ActiveMenuBar.ContextMenuPresent = False    '��ֹ�Ҽ�ѡ�񹤾�����ȡ��
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap  '��ֹ�ƶ�������
End Function

Public Sub CbsButtonInit(ByRef cbsMain As CommandBars, Buttons As Collection, _
                         Optional blnLargeIcons As Boolean = False, _
                         Optional Position As XTPBarPosition)
    '�����������˵�
    'cbsMain :����������
    'Buttons :�˵�����,ÿ��Ԫ�صĸ�ʽΪ �˵�id,����,�Ƿ����
    'blnLargeIcons :�Ƿ��ͼ��
    'Position      :�˵�λ��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim strButton As Variant
    Dim varButton As Variant

    Call CbsSetting(cbsMain)
    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.ActiveMenuBar
    cbsMain.Options.LargeIcons = blnLargeIcons  'Сͼ��
    objBar.Position = Position   '�������ڶ���

    For Each strButton In Buttons
        varButton = Split(strButton, ",")
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, Val(varButton(0)), varButton(1))     '����
            objControl.STYLE = xtpButtonIconAndCaption
            If UCase(varButton(2)) = "TRUE" Then objControl.BeginGroup = True '����
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
'    'cbsMain        :������
'    'frmParan       :����
'    'objLeft        :��߿ؼ�
'    'objWE          :���ҷָ���
'    'objNS          :���·ָ���
'    'objRightTop    :����
'    'objRightButton :����
'    'minWidth       :�ؼ���С���
'    'minHeight      :�ؼ���С�߶�
'    'LeftExpend     :��ʼ��ʱ����߿ؼ����컹���ұ߿ؼ����죬Ĭ��Ϊ�ұ߿ؼ�
'    'TopExpend      :��ʼ��ʱ���ϱ߿ؼ����컹���±߿ؼ����죬Ĭ��Ϊ�±߿ؼ�
'    '�����Ϊ�����ֵ�����¿��Ե���
'    '-------------------------------
'    ' ��  �� ����                  |
'    '     ��-----------------------|
'    '     �� ����                  |
'    '-------------------------------
'
'    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
'    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
'    On Error Resume Next
'
'
'    If frmParan.Width < minWidth * 2 + 100 Then frmParan.Width = minWidth * 2 + 100
'    If frmParan.Height < minHeight * 2 + 100 Then frmParan.Height = minHeight * 2 + 100
'    objWE.MousePointer = 9 '������
'    objNS.MousePointer = 7 '������
'    objWE.BorderStyle = 0   '�ޱ߿�
'    objNS.BorderStyle = 0
'    objWE.Caption = ""
'    objNS.Caption = ""
'    '��ʼ���ָ�����ɫ
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
'    'һ�����һ����������ҳ���Resize
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
'    'cbsMain        :������
'    'frmParan       :����
'    'objLeft        :��߿ؼ�
'    'objWE          :���ҷָ���
'    'objRight       :��
'    'minWidth       :�ؼ���С���
'    'LeftExpend     :��ʼ��ʱ����߿ؼ����컹���ұ߿ؼ����죬Ĭ��Ϊ�ұ߿ؼ�
'    '�����Ϊ�����ֵ�����¿��Ե���
'    '-------------------------------
'    ' ��  �� ��                    |
'    '     ��                       |
'    '     ��                       |
'    '-------------------------------
'    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
'    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
'    On Error Resume Next
'
'    If frmParan.Width < minWidth * 2 + 100 Then frmParan.Width = minWidth * 2 + 100
'    objWE.MousePointer = 9 '������
'    objWE.BorderStyle = 0   '�ޱ߿�
'    objWE.Caption = ""
'    '��ʼ���ָ�����ɫ
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





