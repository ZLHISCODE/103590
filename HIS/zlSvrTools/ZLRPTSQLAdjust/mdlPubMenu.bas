Attribute VB_Name = "mdlPubMenu"
Option Explicit
'�������ݲ˵�ID����
'********************************************************************
Public Const conMenu_FilePopup = 1              '�ļ�
Public Const conMenu_ManagePopup = 2            '����
Public Const conMenu_EditPopup = 3              '�༭
Public Const conMenu_ReportPopup = 4            '����
Public Const conMenu_ViewPopup = 7              '�鿴
Public Const conMenu_ToolPopup = 8              '����
Public Const conMenu_HelpPopup = 9              '����

'�ļ��˵�
Public Const conMenu_File_PrintSet = 101        '��ӡ����(&S)��
Public Const conMenu_File_Preview = 102         'Ԥ��(&V)
Public Const conMenu_File_Print = 103           '��ӡ(&P)
Public Const conMenu_File_Excel = 104           '�����&Excel��
Public Const conMenu_File_Parameter = 181       '��������(&M)
Public Const conMenu_File_LogOut = 190            'ע��(&L)
Public Const conMenu_File_Exit = 191            '�˳�(&X)


'�༭�˵�
Public Const conMenu_Edit_NewParent = 301       '�·���(&N)
Public Const conMenu_Edit_NewItem = 302         '����Ŀ(&A)
Public Const conMenu_Edit_Modify = 303          '�޸�(&M)
Public Const conMenu_Edit_Delete = 304          'ɾ��(&D)
Public Const conMenu_Edit_Audit = 305           '���(&U)
Public Const conMenu_Edit_Blankoff = 306        '����(&B)
Public Const conMenu_Edit_Disuse = 307          'ͣ��(&P)
Public Const conMenu_Edit_Reuse = 308           '����(&R)

'�鿴�˵�
Public Const conMenu_View_ToolBar = 701              '������(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Public Const conMenu_View_StatusBar = 702            '״̬��(&S)
Public Const conMenu_View_Expend = 711               'չ��/�۵���(&X)
Public Const conMenu_View_Expend_AllCollapse = 7111     '�۵�������(&L)
Public Const conMenu_View_Expend_AllExpend = 7112       'չ��������(&X)
Public Const conMenu_View_Expend_CurCollapse = 7113     '�۵���ǰ��(&C)
Public Const conMenu_View_Expend_CurExpend = 7114       'չ����ǰ��(&E)
Public Const conMenu_View_Filter = 721               '����(&G)
Public Const conMenu_View_Find = 722                 '����(&F)
Public Const conMenu_View_FindNext = 723             '������һ��(&N)
Public Const conMenu_View_Refresh = 791              'ˢ��(&R)

Public Const conMenu_View_Navigation = 792              '���ܵ���(&D)

Public Const conMenu_View_Property = 793              '����(&P)


'�����˵�
Public Const conMenu_Help_Help = 901        '��������(&H)
Public Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_Help_Web_Mail = 9022       '���ͷ���(&M)
Public Const conMenu_Help_About = 991       '����(&A)��

'��ݼ�
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

