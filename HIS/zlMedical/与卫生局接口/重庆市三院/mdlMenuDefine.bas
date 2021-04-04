Attribute VB_Name = "mdlMenuDefine"
Option Explicit
'�������ݲ˵�ID����
'*********************************************************************
Public Const conMenu_FilePopup = 1 '�ļ�
Public Const conMenu_ManagePopup = 2 '����
Public Const conMenu_EditPopup = 3 '�༭
Public Const conMenu_ReportPopup = 4 '����
Public Const conMenu_ViewPopup = 7 '�鿴
Public Const conMenu_ToolPopup = 8 '����
Public Const conMenu_HelpPopup = 9 '����

'�ļ��˵�

Public Const conMenu_File_Login = 100            '*��(&O)��
Public Const conMenu_File_PrintSet = 101        '*��ӡ����(&S)��
Public Const conMenu_File_Preview = 102         '*Ԥ��(&V)
Public Const conMenu_File_Print = 103           '*��ӡ(&P)
Public Const conMenu_File_Excel = 104           '�����&Excel��
Public Const conMenu_File_MedRec = 105          '��ҳ��ӡ(&R)
Public Const conMenu_File_MedRecSetup = 1051        '��ӡ����(&S)
Public Const conMenu_File_MedRecPreview = 1052      '��ӡԤ��(&P)
Public Const conMenu_File_MedRecPrint = 1053        '��ӡ��ҳ(&V)
Public Const conMenu_File_RowPrint = 121        '��¼��ӡ(&R)
Public Const conMenu_File_BatPrint = 122        '������ӡ(&B)
Public Const conMenu_File_Parameter = 181       '*��������(&M)
Public Const conMenu_File_Exit = 191            '*�˳�(&X)

'�༭�˵�
Public Const conMenu_Manage_Regist = 211      '*���˹Һ�(&H)
Public Const conMenu_Manage_Bespeak = 212     'ԤԼ�Һ�(&B)
Public Const conMenu_Manage_Transfer = 213    '����ת��(&C)
Public Const conMenu_Manage_Receive = 214     '*���˽���(&Z)
Public Const conMenu_Manage_Cancel = 215      'ȡ������(&Q)
Public Const conMenu_Manage_Finish = 216      '*��ɽ���(&W)
Public Const conMenu_Manage_Redo = 217        '�ָ�����(&R)

'ҽ���˵�����϶�,����ʱ��4λ���,50λ�ֶ�,001-050,051-100,101-150,...
Public Const conMenu_Edit_Dept = 3001    '*����Ŀ(&A)
Public Const conMenu_Edit_Diagnose = 3002     '*����/��¼(&Y)
Public Const conMenu_Edit_Check = 3003     '*�޸�(&M)
Public Const conMenu_Edit_Combo = 3004     '*ɾ��(&D)
Public Const conMenu_Edit_Verify = 3005   '*����(&B)
Public Const conMenu_Edit_Stop = 3006       '*ҽ��ֹͣ(&S)
Public Const conMenu_Edit_ReStop = 3007     '*ȷ��ֹͣ(&C)
Public Const conMenu_Edit_Pause = 3008      '*��ͣ(&P)
Public Const conMenu_Edit_Reuse = 3009      '*����(&U)
Public Const conMenu_Edit_Audit = 3010      '*���/У��(&V)
Public Const conMenu_Edit_Price = 3011      '*�Ƽ۵���(&I)
Public Const conMenu_Edit_ClearUp = 3012    '*ҽ������(&F)
Public Const conMenu_Task_Send = 3013       '*����(&G)
Public Const conMenu_Edit_SendDrug = 30131      '*ҩ��ҽ������(&1)
Public Const conMenu_Edit_SendOther = 30132     '����ҽ������(&2)
Public Const conMenu_Edit_Untread = 3014    '*����(&R)
Public Const conMenu_Edit_SendBack = 3015   '*���ڷ����ջ�(&N)
Public Const conMenu_Edit_Test = 3016       '*Ƥ�Խ��(&T)

'�����˵�
Public Const conMenu_Edit_NewParent = 3051   '*�·���(&N)
Public Const conMenu_Edit_Insert = 3052      '*����(&I)
Public Const conMenu_Edit_MarkMap = 3061     '*ͼƬ(&I)��
Public Const conMenu_Edit_ApplyTo = 3062     '*���ÿ���(&T)
Public Const conMenu_Edit_Request = 3063     '����Ҫ��(&R)
Public Const conMenu_Edit_Compend = 3064     '*���ݹ���(&F)
Public Const conMenu_Edit_Import = 3071      '*��������(&B)��
Public Const conMenu_Edit_Adjust = 3082      '*����(&J)
Public Const conMenu_Edit_Archive = 3083     '*�鵵(&R)
Public Const conMenu_Task_Accept = 3091        '*����
Public Const conMenu_Edit_Sort = 3092        '*���ĵ�����

'����˵�
Public Const conMenu_Report_DrugQuery = 401    'ҩ���շ���ѯ(&H)
Public Const conMenu_Report_Reports = 402      '�������ñ���(&W)
Public Const conMenu_Report_MultiBill = 403    '��ӡ�ಡ�˵���(&K)
Public Const conMenu_Report_ClinicBill = 404   '��ӡ���Ƶ���(&J)��
Public Const conMenu_Report_AdviceBill1 = 405  '����ҽ����(&P)
Public Const conMenu_Report_AdviceBill2 = 406  '��ʱҽ����(&T)
Public Const conMenu_Report_AdviceBill3 = 407  'ҽ����¼��(&B)
Public Const conMenu_Report_WorkLog = 408      '�����ձ�(&O)
Public Const conMenu_Report_Item = 451         'Ԥ��Ϊ�Ժ�Ķ�̬��������ʼ��

'�鿴�˵�
Public Const conMenu_View_ToolBar = 701              '������(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Public Const conMenu_View_StatusBar = 702            '״̬��(&S)
Public Const conMenu_View_Append = 703               '������Ϣ(&A)
Public Const conMenu_View_Expend = 711               'չ��/�۵���(&X)
Public Const conMenu_View_Expend_CurCollapse = 7111     '�۵���ǰ��(&C)
Public Const conMenu_View_Expend_CurExpend = 7112       'չ����ǰ��(&E)
Public Const conMenu_View_Expend_AllCollapse = 7113     '�۵�������(&L)
Public Const conMenu_View_Expend_AllExpend = 7114       'չ��������(&X)
Public Const conMenu_View_Find = 721                 '*����(&F)
Public Const conMenu_View_FindNext = 722             '��������(&N)
Public Const conMenu_View_FindType = 723             '���ҷ�ʽ(&Y)
Public Const conMenu_View_Filter = 731               '*���ݹ���(&I)
Public Const conMenu_View_Notify = 732               '*ҽ������(&B)
Public Const conMenu_View_Busy = 733                 '����æ(&M)
Public Const conMenu_View_Hide = 741                 '*����(&H)
Public Const conMenu_View_Show = 742                 '*��ʾ(&S)
Public Const conMenu_View_Backward = 743             '*����(&B)
Public Const conMenu_View_Forward = 744              '*ǰ��(&F)
Public Const conMenu_View_Option = 781               'ѡ��(&O)
Public Const conMenu_View_Refresh = 791              '*ˢ��(&R)
Public Const conMenu_View_Jump = 792                 '��ת(&J)

'���߲˵�
Public Const conMenu_Tool_Reference = 801       '*�ο�(&R)
Public Const conMenu_Tool_Reference_1 = 8011    '������ϲο�(&D)
Public Const conMenu_Tool_Reference_2 = 8012    '���ƴ�ʩ�ο�(&C)
Public Const conMenu_Tool_MedRec = 802          '*��ҳ����(&M)
Public Const conMenu_Tool_Meet = 803            '*���˻���(&E)
Public Const conMenu_Tool_MeetFinish = 8031         '��ɻ���(&F)
Public Const conMenu_Tool_MeetCancel = 8032         'ȡ�����(&C)
Public Const conMenu_Tool_Sign = 804            '*����ǩ��(&I)
Public Const conMenu_Tool_SignNew = 8041            '����ǩ��(&I)
Public Const conMenu_Tool_SignVerify = 8042         '��֤ǩ��(&V)
Public Const conMenu_Tool_SignEarse = 8043          'ȡ��ǩ��(&E)
Public Const conMenu_Tool_Monitor = 811         '*���(&M)
Public Const conMenu_Tool_Monitor_1 = 81101         'ʱ��Ҫ����(&T)
Public Const conMenu_Tool_Monitor_2 = 81102         '����Ҫ����(&C)
Public Const conMenu_Tool_Assistant = 812       '*����(&A)
Public Const conMenu_Tool_Analyse = 813         '*����(&Y)
Public Const conMenu_Tool_Search = 814          '*����(&S)
Public Const conMenu_Tool_Define = 815          '*����(&D)
Public Const conMenu_Tool_Report = 816          '*����(&P)
Public Const conMenu_Tool_Apply = 817           '*Ӧ��(&A)
Public Const conMenu_Tool_Option = 819          'ѡ��(&O)

'�����˵�
Public Const conMenu_Help_Help = 901        '*��������(&H)
Public Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_Help_Web_Mail = 9022       '*���ͷ���(&M)
Public Const conMenu_Help_About = 991       '����(&A)��

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

'*********************************************************************

