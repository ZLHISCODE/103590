Attribute VB_Name = "mdlXtremeSuite"
Option Explicit


'''''''''''''''''''Xtreme�ؼ���ض���
''���˵�
Public Const mconMenu_FilePopup = 1 '�ļ�
Public Const mconMenu_ManagePopup = 2 '����
Public Const mconMenu_EditPopup = 3 '�༭
Public Const mconMenu_ReportPopup = 4 '����
Public Const mconMenu_ViewPopup = 7 '�鿴
Public Const mconMenu_ToolPopup = 8 '����
Public Const mconMenu_HelpPopup = 9 '����

''�ļ��˵�
Public Const mconMenu_File_Open = 100               '*��(&O)��
Public Const mconMenu_File_PrintSet = 101           '*��ӡ����(&S)��
Public Const mconMenu_File_Preview = 102            '*Ԥ��(&V)
Public Const mconMenu_File_Print = 103              '*��ӡ(&P)
Public Const mconMenu_File_Excel = 104              '�����&Excel��
Public Const mconMenu_File_BillPrint = 105          '���ݴ�ӡ
Public Const mconMenu_File_BillPreview = 106        '����Ԥ��
Public Const mconMenu_File_Parameter = 181       '*��������(&M)

Public Const mconMenu_File_Exit = 191            '*�˳�(&X)

''�鿴�˵�
Public Const mconMenu_View_ToolBar = 701              '������(&T)
Public Const mconMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Public Const mconMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Public Const mconMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Public Const mconMenu_View_StatusBar = 702            '״̬��(&S)
Public Const mconMenu_View_Append = 703               '������Ϣ(&A)
Public Const mconMenu_View_Expend = 711               'չ��/�۵���(&X)
Public Const mconMenu_View_Expend_CurCollapse = 7111     '�۵���ǰ��(&C)
Public Const mconMenu_View_Expend_CurExpend = 7112       'չ����ǰ��(&E)
Public Const mconMenu_View_Expend_AllCollapse = 7113     '�۵�������(&L)
Public Const mconMenu_View_Expend_AllExpend = 7114       'չ��������(&X)
Public Const mconMenu_View_Find = 721                 '*����(&F)
Public Const mconMenu_View_FindNext = 722             '��������(&N)
Public Const mconMenu_View_FindType = 723             '���ҷ�ʽ(&Y)
Public Const mconMenu_View_ReadIC = 724               '��IC��(&I)
Public Const mconMenu_View_PatInfor = 725             '�鿴������Ϣ
Public Const mconMenu_View_PriceBill = 727
Public Const mconMenu_View_PriceTable = 728
Public Const mconMenu_View_PriceList = 729
Public Const mconMenu_View_FilterView = 730           '�Թ��˷�ʽ��ʾ
Public Const mconMenu_View_Filter = 731               '*���ݹ���(&I),�Ӵ���Ĺ��˹���
Public Const mconMenu_View_Notify = 732               '*ҽ������(&B)
Public Const mconMenu_View_Busy = 733                 '����æ(&M)
Public Const mconMenu_View_ShowAll = 734
Public Const mconMenu_View_ShowHistory = 735
Public Const mconMenu_View_ShowStoped = 736
Public Const mconMenu_View_Hide = 741                 '*����(&H)
Public Const mconMenu_View_Show = 742                 '*��ʾ(&S)
Public Const mconMenu_View_Forward = 743              '*ǰ��(&F)
Public Const mconMenu_View_Backward = 744             '*����(&B)
Public Const mconMenu_View_Dept = 745                '�鿴����
Public Const mconMenu_View_Location = 746            '��λ
Public Const mconMenu_View_LocationItem = 747        '��λ��Ŀ
Public Const mconMenu_View_ColSet = 748              '������
Public Const mconMenu_View_Option = 781               'ѡ��(&O)
Public Const mconMenu_View_Refresh = 791              '*ˢ��(&R)
Public Const mconMenu_View_Jump = 792                 '��ת(&J)

Public Const mconMenu_View_SelAll = 7301              'ȫѡ
Public Const mconMenu_View_ClsAll = 7302              'ȫ��

Public Const mconMenu_View_Navigatebeginning = 7401           '*��һ��(&F)
Public Const mconMenu_View_Navigateleft = 7402                '*��һ��(&F)
Public Const mconMenu_View_Navigateright = 7403               '*��һ��(&F)
Public Const mconMenu_View_Navigateend = 7404                 '*���һ��(&F)

Public Const mconMenu_View_FontSize = 4004         '�ֺ�����
Public Const mconMenu_View_FontSize_1 = 4004         '9����
Public Const mconMenu_View_FontSize_2 = 4004         '11����
Public Const mconMenu_View_FontSize_3 = 4004         '15����

''�����˵�
Public Const mconMenu_Help_Help = 901           '*��������(&H)
Public Const mconMenu_Help_Web = 902            '&WEB�ϵ�����
Public Const mconMenu_Help_Web_Home = 9021      '������ҳ(&H)
Public Const mconMenu_Help_Web_Forum = 9023     '������̳(&F)
Public Const mconMenu_Help_Web_Mail = 9022      '*���ͷ���(&M)
Public Const mconMenu_Help_About = 991          '����(&A)��

'�̵�����
Public Const mconMenu_Edit_AddBill = 3001        '���Ӽ�¼��
Public Const mconMenu_Edit_AddTable = 3002       '�����̵��
Public Const mconMenu_Edit_AddTableAuto = 30021  '�Զ������̵��
Public Const mconMenu_Edit_AddTableTotal = 30022 '���ܼ�¼�������̵��
Public Const mconMenu_Edit_AddTableZero = 30023  'ȫ����Ϊ��
Public Const mconMenu_Edit_AddTableHouseAll = 30024  '�ⷿȫ��ҩƷ�̵�
Public Const mconMenu_Edit_AddTableSpecial = 30025   '����ҩƷ�̵�
Public Const mconMenu_Edit_AddModify = 3003      '�޸�
Public Const mconMenu_Edit_AddDel = 3004         'ɾ��
Public Const mconMenu_Edit_AddVerify = 3005      '���
Public Const mconMenu_Edit_AddStrike = 3006      '����
Public Const mconMenu_Edit_AddAffirmant = 3007   '�Ķ�ȷ��
Public Const mconMenu_Edit_AddDisplay = 3008     '�鿴����
Public Const mconMenu_Edit_CheckTable = 5001     '�̵�����ܼ��

''CommandBar�����ȼ�
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

''CommandBar�����
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
