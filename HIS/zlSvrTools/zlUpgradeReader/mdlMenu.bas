Attribute VB_Name = "mdlMenu"
Option Explicit

'********************************************************************
'CommandBar����ID
Public Enum CommandBarIDCond
    conMenu_File = 1
    conMenu_Edit = 2
    conMenu_View = 8
    conMenu_Help = 9
    
    '�ļ��˵�
    conMenu_File_Open = 101
    conMenu_File_Save = 103
    conMenu_File_Login = 107
    conMenu_File_Logout = 108
    conMenu_File_Exit = 109
    
    '�༭�˵�

    
    '�鿴�˵�
    conMenu_View_Expend = 711               'չ��/�۵���(&X)
    conMenu_View_Expend_AllCollapse = 7111     '�۵�������(&L)
    conMenu_View_Expend_AllExpend = 7112       'չ��������(&X)
    conMenu_View_Expend_CurCollapse = 7113     '�۵���ǰ��(&C)
    conMenu_View_Expend_CurExpend = 7114       'չ����ǰ��(&E)

    conMenu_View_ShowPrivewText = 722          '��ʾ����
    conMenu_View_ShowGroupBox = 723            '��ʾ����
    conMenu_View_ShowRelation = 724            '��ʾ��������
    
    conMenu_View_Filter = 802
    conMenu_View_RecordPrev = 803
    conMenu_View_RecordNext = 804
    conMenu_View_Find = 805
    conMenu_View_FindNext = 806
    conMenu_View_Refresh = 809
    conMenu_View_Close = 810
    
    '�����˵�
    conMenu_Help_About = 901
    
    conMenu_Custom_System = 900001
    conMenu_Custom_Icon = 900002
End Enum

'CommandBar���г�������
Public Const XTP_ID_WINDOW_LIST = 35000 '�����б�
Public Const XTP_ID_TOOLBARLIST = 59392 '�������б�
Public Const ID_INDICATOR_CAPS = 59137 '״̬������д��
Public Const ID_INDICATOR_NUM = 59138 '״̬�������֣�
Public Const ID_INDICATOR_SCRL = 59139 '״̬����������

'CommandBar�����ȼ�
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16
'********************************************************************

Public Enum IconID
    ICON_Mail = 0   '�ʼ�ͼ��
    ICON_Importance '��Ҫ��
    ICON_FlagTrain  '��ѵ��־
    ICON_NoRead     'δ��
    ICON_Read       '�Ѷ�
    ICON_Unknown    '��ȷ��
    ICON_Low        '��
    ICON_Center     '��
    ICON_High       '��
    ICON_Train      '����ѵ
    ICON_UnTrain    'δ��ѵ
End Enum
