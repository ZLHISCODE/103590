Attribute VB_Name = "mdlPubDefine"
Option Explicit
'ȫ�ֱ�����ȫ�ֲ���
'----------------------------------------------------------------------------------
Public Const ConMenu_Appfro_AddBill = 101           '���ӷ���
Public Const ConMenu_Appfro_DelBill = 102           'ɾ������
Public Const ConMenu_Appfro_ModifyItem = 103        '�޸���Ŀ
Public Const ConMenu_Appfro_Exit = 104              '�˳�
Public Const ConMenu_Appfro_DeptSel = 105           'ִ�п���ѡ��
Public Const ConMenu_Appfro_Refresh = 106           'ˢ��
Public Const ConMenu_Appfro_ModifyBill = 107        '�޸ķ���
Public Const ConMenu_Appfro_ModifyDept = 108        '�޸�ִ�п���
Public Const ConMenu_Appfor_ItemSort = 401          '����˳��
Public Const ConMenu_Appfor_ClincHelp = 402         '����˳��

Public Const ConMenu_Browse_SelAll = 109            'ȫѡ
Public Const ConMenu_Browse_ClsAll = 110            'ȫ��
Public Const ConMenu_Browse_Refresh = 111           'ˢ��
Public Const ConMenu_Browse_Print = 112             '��ӡ
Public Const ConMenu_Browse_Exit = 113              '�˳�
Public Const ConMenu_Browse_Find = 114              '����


Public Const ConMenu_Browse_Save = 115              '����
Public Const ConMenu_Browse_Cancel = 116            'ȡ��
Public Const ConMenu_Browse_PrintView = 117         '��ӡԤ��
Public Const ConMenu_Browse_PrintSet = 118          '��ӡ����
Public Const ConMenu_Appfro_Group = 119             'ѡ�����
Public Const conFun_Sample_Auditing = 120           '����
Public Const conFun_Sample_unAuditing = 121         'ȡ������
Public Const ConMenu_Browse_unPrint = 122           '���ô�ӡ
Public Const ConMenu_Browse_PrintAll = 123          '��ӡ����
Public Const ConMenu_Browse_PrintViewAll = 124      'Ԥ������
Public Const ConMenu_Browse_PrintSetAll = 125       '��ӡ����

'-------------------------------------------------------------------
Public Const ConTab_Sample_History = 201            '����
Public Const ConTab_Sample_Image = 202              'ͼ��
Public Const ConTab_Sample_Comment = 203            '��ע
'--------------------------------------------------------------------

'---------------------------------------------------------------------
Public Const ConMenu_pop_In = 301             'סԺ��
Public Const ConMenu_pop_bed = 302            '����
Public Const ConMenu_pop_Dept = 303           '�������
Public Const ConMenu_pop_DeptDistrict = 304   '���벡��
Public Const ConMenu_pop_Out = 305            '�����
Public Const ConMenu_pop_PatiCard = 306       '���￨
Public Const ConMenu_pop_SampleCode = 307     '�����

'---------------------------------------------------------------------
'������������
'*********************************************************************
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

'CommandBar�����
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_PAGEUP = &H21
Public Const VK_PAGEDOWN = &H22
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B

Public Const VK_P = &H50

Public Const VsModiBackColor = &HD6FFCA        'vs�ؼ����ɱ༭��Ԫ�ı���ɫ
'*********************************************************************

Public Const conMenu_Tool_PlugIn = 890          '���
Public Const conMenu_Tool_PlugIn_Item = 89000   '�����,ʵ������Ϊ conMenu_Tool_PlugIn_Item + n, 1<=n<=99

'�˵���ť
Public Const CONFUN_UP = 501                            '��һ��
Public Const CONFUN_DOWN = 502                          '��һ��
