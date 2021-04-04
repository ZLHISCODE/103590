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
    '���ܣ�                                  �����˵�
    '������
    '           IconX                        ����ͼ��X��С
    '           IconY                        ����ͼ��Y��С
    '���أ�                                  ��
    '�ϼ���������̣�                        frViewer_load
    '�¼���������̣�                        ��
    '���õ��ⲿ������                        ��
    '�����ˣ�                                ���� 2005-6-27
    '------------------------------------------------
    Dim Control As CommandBarControl
    Dim ControlFile As CommandBarPopup
    Dim ControlSelect As CommandBarPopup
    Dim ToolBar As CommandBar
    Dim ControlPopup As CommandBarPopup
    
    'ȥ����չ��ť
    ToolBars.Options.ShowExpandButtonAlways = False
    ToolBars.ActiveMenuBar.EnableDocking xtpFlagHideWrap
    
    'ȥ���˵�
    ToolBars.Item(1).Visible = False
    
    
    Set ToolBar = ToolBars.Add("��������", xtpBarBottom)
    
    With ToolBar.Controls
        .Add xtpControlButton, ID_Capture_CapPicture, "�ɼ�"
'        .Add xtpControlButton, ID_Capture_CapSaveVido, "¼��"
        .Add xtpControlButton, ID_Capture_SavePicture, "����"
        .Add xtpControlButton, ID_Capture_DelPicture, "ɾ��"
        
        Set ControlPopup = .Add(xtpControlSplitButtonPopup, ID_Capture_Setup, "����")
        ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Capture_Setup_Format, "��ʽ"
        ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Capture_Setup_Source, "��Դ"
        
        .Add xtpControlButton, ID_Capture_Prot, "�˿�"
        
        .Add xtpControlButton, ID_Capture_Exit, "�˳�"
    End With
    
    ToolBar.Position = xtpBarTop
    ToolBar.SetIconSize 24, 24
    ToolBar.ShowTextBelowIcons = True
    
    
End Sub


