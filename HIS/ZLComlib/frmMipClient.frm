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
   StartUpPosition =   3  '����ȱʡ
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
    '���ܣ���ʼ��
    '��������
    '���أ���ʼ���ɹ�����True,���򷵻�False
    '******************************************************************************************************************
    
'    Call InitCommandBar
'    Call AddIcon(picNotify.hwnd, imgIcon(0).Picture, "��Ϣ����ƽ̨�ͻ����շ�����")
            
    Initialize = True
    
End Function

'Private Function InitCommandBar() As Boolean
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim objMenu As CommandBarPopup
'    Dim objBar As CommandBar
'    Dim objPopup As CommandBarPopup
'    Dim objControl As CommandBarControl
'    Dim cbrCustom As CommandBarControlCustom
'
'    '------------------------------------------------------------------------------------------------------------------
'    '��ʼ����
'    cbsMain.VisualTheme = xtpThemeOffice2003
'
'    With cbsMain.Options
'        .ShowExpandButtonAlways = False
'        .ToolBarAccelTips = True
'        .AlwaysShowFullMenus = False
'        '.UseFadedIcons = True '����VisualTheme����Ч
'        .IconsWithShadow = True '����VisualTheme����Ч
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
'    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
'
'    cbsMain.ActiveMenuBar.Title = "�˵�"
'    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
'    cbsMain.ActiveMenuBar.Visible = False
'
'End Function
'
'Public Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
'    '******************************************************************************************************************
'    '���ܣ�
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim cbrPopupBar As CommandBar
'    Dim cbrPopupItem As CommandBarControl
'    Dim cbrPopupItem2 As CommandBarControl
'    Dim cbrMenuBar As CommandBarControl
'    Dim cbrControl As CommandBarControl
'    Dim cbrControl2 As CommandBarControl
'
'    '�����˵�����
'
'    On Error GoTo errHand
'
'    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
'
'    Select Case bytPlace
'    Case 1
'
'        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 1, "������־(&L)")
'        cbrPopupItem.DefaultItem = True
'
'        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 2, "��������(&S)")
'        cbrPopupItem.BeginGroup = True
'
'        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 3, "ֹͣ����(&T)")
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
'        MsgBox "��������"
'    Case 3
'        MsgBox "ֹͣ����"
'    End Select
'End Sub

Private Sub Form_Unload(Cancel As Integer)
        
'    Call RemoveIcon(picNotify.hwnd)
    
End Sub

'Private Sub picNotify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '--------------------------------------------------------------------------------------------------
'    '����:  ����picNotify�ĸ��ִ����¼�,��Ҫ�����Զ�������ع���(�����д)
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
