Attribute VB_Name = "mdlXtremeSuite"
Option Explicit


'''''''''''''''''''Xtreme�ؼ���ض���
''��������
'���ŷ�ҩ
Public Const mconPane_Dept_Condition = 1                    '������

'PIVA����
Public Const mconPane_PIVA_Condition = 1                    '������

'������ҩ
Public Const mconPane_Recipe_Condition = 1    '������
Public Const mconPane_Recipe_List = 2         '�����б�

''TabControl��ҳ
'PIVA����
Public Const mconTab_PIVA_Check = 0                 '���˲�
Public Const mconTab_PIVA_Dosage = 1                '����ҩ
Public Const mconTab_PIVA_Send = 2                  '������
Public Const mconTab_PIVA_Return = 3                '�ѷ���

'���ŷ�ҩ
Public Const mconTab_Dept_Send = 0                  'δ��ҩƷ�嵥
Public Const mconTab_Dept_SumSend = 1               '�����嵥
Public Const mconTab_Dept_Shortage = 2              'ȱҩ�嵥
Public Const mconTab_Dept_Reject = 3                '�ܷ�ҩ�嵥
Public Const mconTab_Dept_Return = 4                '�ѷ�ҩ�嵥

'������ҩ
Public Const mconTab_Recipe_DosageOk = 0          '��ҩȷ��
Public Const mconTab_Recipe_Dosage = 1            '��ҩ
Public Const mconTab_Recipe_Abolish = 2           'ȡ����ҩ
Public Const mconTab_Recipe_Send = 3              '����ҩ
Public Const mconTab_Recipe_OverTime = 4          '����δ��
Public Const mconTab_Recipe_Return = 5            '��ҩ

''���˵�
Public Const mconMenu_FilePopup = 1 '�ļ�
Public Const mconMenu_ManagePopup = 2 '����
Public Const mconMenu_EditPopup = 3 '�༭
Public Const mconMenu_ReportPopup = 4 '����
Public Const mconMenu_PlanPopup = 5 '�Ű�����
Public Const mconMenu_ViewPopup = 7 '�鿴
Public Const mconMenu_ToolPopup = 8 '����
Public Const mconMenu_HelpPopup = 9 '����

''�ļ��˵�
Public Const mconMenu_File_Open = 100               '*��(&O)��
Public Const mconMenu_File_PrintSet = 101           '*��ӡ����(&S)��
Public Const mconMenu_File_Preview = 102            '*Ԥ��(&V)
Public Const mconMenu_File_Print = 103              '*��ӡ(&P)
Public Const mconMenu_File_Excel = 104              '�����&Excel��

Public Const mconMenu_File_Parameter = 181          '*��������(&M)

Public Const mconMenu_File_Exit = 191               '*�˳�(&X)
Public Const mconMenu_File_Message = 10000            '��Ϣ�˵�

'PIVA
Public Const mconMenu_File_PIVA_BillPrint = 151              '���ݴ�ӡ
Public Const mconMenu_File_PIVA_BillPrintWait = 152          '��ӡҩƷ��ҩ��
Public Const mconMenu_File_PIVA_BillPrintLable = 153         '��ӡ��ǩ
Public Const mconMenu_File_PIVA_BillPrintTotal = 154         '��ӡ��ҩ�嵥
Public Const mconMenu_File_PIVA_BillPrintReturn = 155        '��ӡ��ҩ�����ʣ��嵥
Public Const mconMenu_File_PIVA_BillPrintNext = 156          '����ƿǩ
Public Const mconMenu_File_PIVA_BillPrintSum = 157         '��ӡ���ܱ���

'���ŷ�ҩ
Public Const mconMenu_File_Dept_BillPrint = 151              '���ݴ�ӡ
Public Const mconMenu_File_Dept_BillPrintTotal = 152         '��ӡ�����嵥
Public Const mconMenu_File_Dept_BillPrintRestore = 153       '��ӡ��ҩ֪ͨ��
Public Const mconMenu_File_Dept_BillPrintWait = 154          '��ӡҩƷ��ҩ��

'������ҩ
Public Const mconMenu_File_Recipe_BillPrintDosage = 151            '��ӡ��ҩ��(&B)-F6
Public Const mconMenu_File_Recipe_BillPrintRecipe = 152            '��ӡ����ǩ(&D)-F4
Public Const mconMenu_File_Recipe_BillPrintReport = 153            '��ӡ��ҩ�嵥(&W)
Public Const mconMenu_File_Recipe_BillPrintReturn = 154            '��ӡ��ҩ֪ͨ��(&R)
Public Const mconMenu_File_Recipe_BillPrintLable = 155             '��ӡҩƷ��ǩ(&L)-F11
Public Const mconMenu_File_Recipe_BillPrintBack = 156              '��ӡ�˷ѵ���(T)
Public Const mconMenu_File_Recipe_BillPrintChange = 157            '��ӡҽ������֪ͨ��


''�༭�˵�
'PIVA
Public Const mconMenu_Edit_PIVA_Check = 3301                '�˲飨�󷽣�
Public Const mconMenu_Edit_PIVA_Prepare = 3302              '��ҩ
Public Const mconMenu_Edit_PIVA_AutoSetBatch = 3303         '�Զ���������
Public Const mconMenu_Edit_PIVA_Dosage = 3304               '��ҩ
Public Const mconMenu_Edit_PIVA_CancelDosage = 3305         'ȡ����ҩ
Public Const mconMenu_Edit_PIVA_Send = 3306                 '����
Public Const mconMenu_Edit_PIVA_CancelSend = 3307           'ȡ������
Public Const mconMenu_Edit_PIVA_Cancel = 3308               'ȡ��
Public Const mconMenu_Edit_PIVA_PASS = 3309                 'PASS
Public Const mconMenu_Edit_PIVA_Delete = 3310               'ɾ������
Public Const mconMenu_Edit_PIVA_ReVerify = 3311             '�������
Public Const mconMenu_Edit_PIVA_Approve = 3312                '���ҽ��
Public Const mconMenu_Edit_PIVA_CancelApprove = 3313         'ȡ�����
Public Const mconMenu_Edit_PIVA_Lock = 3314                 'ȫ������
Public Const mconMenu_Edit_PIVA_UnLock = 3315               'ȫ������
Public Const mconMenu_Edit_PIVA_Beach = 3316                '��������
Public Const MCONMENU_EDIT_PIVA_REFUSE = 3317               'ȷ�Ͼܾ�
Public Const MCONMENU_EDIT_PIVA_SURE = 3318                 'ȷ�ϵ���
Public Const MCONMENU_EDIT_PIVA_SORTSET = 3319              '��������
Public Const MCONMENU_EDIT_PIVA_PLAN = 3320                 '�Űల��
Public Const MCONMENU_EDIT_PIVA_MedicalRecord = 3331        '���Ӳ�������



'�Ű�˵�
'PIVA
Public Const MCONMENU_PLAN_PIVA_DESK = 3501
Public Const MCONMENU_PLAN_PIVA_DESKDRUG = 3502
Public Const MCONMENU_PLAN_PIVA_PERWORK = 3503

'��ҷ�ҩҵ��
Public Const mconMenu_Edit_PlugIn = 3400          '��չ

'���ŷ�ҩ
Public Const mconMenu_Edit_Dept_Verify = 3101             '��ҩ
Public Const mconMenu_Edit_Dept_Desire = 3102             'ȱҩ����
Public Const mconMenu_Edit_Dept_Reject = 3103             '�ܷ�ȷ��
Public Const mconMenu_Edit_Dept_Return = 3104             '��ҩ
Public Const mconMenu_Edit_Dept_ReturnOther = 3105        '������ҩ���Ĵ���
Public Const mconMenu_Edit_Dept_ReVerify = 3106           'ҩƷ��ҩ����
Public Const mconMenu_Edit_Dept_StopFlag = 3107           'ֹͣ��ҩ���
Public Const mconMenu_Edit_Dept_RejectRestore = 3108      '�ܷ��ָ�
Public Const mconMenu_Edit_Dept_EMR = 3109                '������ѯ
Public Const mconMenu_Edit_Dept_Packer = 3110             '�ְ���
Public Const mconMenu_Edit_Dept_VerifySign = 3112       '��֤ǩ��
Public Const mconMenu_Edit_Dept_Hot_IC = 3150             'IC����ť�ȼ�
Public Const mconMenu_Edit_Dept_CustomCheck = 3151             '�Զ�����˹���
Public Const mconMenu_Edit_Dept_MedicalRecord = 3161           '���Ӳ�������

'������ҩ
Public Const mconMenu_Edit_Recipe_Dosage = 3201                 '��ҩģʽ(&D)-^D
Public Const mconMenu_Edit_Recipe_Abolish = 3202                'ȡ��ģʽ(&A)-^A
Public Const mconMenu_Edit_Recipe_Send = 3203                   '��ҩģʽ(&C)-^C
Public Const mconMenu_Edit_Recipe_Return = 3204                 '��ҩģʽ(&H)-^H
Public Const mconMenu_Edit_Recipe_Batch = 3205                  '������ҩ(&B)
Public Const mconMenu_Edit_Recipe_SendOther = 3206              '������ҩ���Ĵ���(&F)
Public Const mconMenu_Edit_Recipe_ReturnBatch = 3207            '������ҩ���Ĵ���(&T)
Public Const mconMenu_Edit_Recipe_SendByBill = 3208             '��Ʊ�ݺŷ�ҩ(&I)
Public Const mconMenu_Edit_Recipe_ReturnByBill = 3209           '��Ʊ�ݺ���ҩ(&R)
Public Const mconMenu_Edit_Recipe_Flag = 3210                   'ֹͣ��ҩ���(&S)
Public Const mconMenu_Edit_Recipe_Cancel = 3211                 'ȡ����ҩ(&Q)-^Q
Public Const mconMenu_Edit_Recipe_Charge = 3212                 '���ﻮ��(&M)-F8
Public Const mconMenu_Edit_Recipe_Stuff = 3213                  '���ķ���(@W)-F9
Public Const mconMenu_Edit_Recipe_Change = 3214                 '�л���ҩ��(&E)
Public Const mconMenu_Edit_Recipe_EMR = 3215                     '������ѯ
Public Const mconMenu_Edit_Recipe_SendHot = 3216                '��ҩ���������ڿ�ݼ�����
Public Const mconMenu_Edit_Recipe_AddSign = 3217                '����ǩ��
Public Const mconMenu_Edit_Recipe_Windows = 3218               '������ҩ����(&N)
Public Const mconMenu_Edit_Recipe_Call = 3219                  '����(&G)
Public Const mconMenu_Edit_Recipe_VerifySign = 3220               '��֤ǩ��
Public Const mconMenu_Edit_Recipe_Cancle = 3221                  'ȡ��ȷ��(&G)
Public Const mconMenu_Edit_Recipe_TakeDrug = 3222                'ȡҩȷ��
Public Const mconMenu_Edit_Recipe_Hot_IC = 3250                 'IC����ť�ȼ�
Public Const mconMenu_Edit_Recipe_AutoSend = 3222                  '�����Զ���ҩ����
Public Const mconMenu_Edit_Recipe_AutoSend_Open = 32221            '���ô����ϴ�
Public Const mconMenu_Edit_Recipe_AutoSend_Set = 32222            '����webService·��
Public Const mconMenu_Edit_Recipe_AutoSend_LoadDrug = 32223           '�ϴ�ҩƷ��������
Public Const mconMenu_Edit_Recipe_AutoSend_LoadStock = 32224          '�ϴ�ҩƷ�������
Public Const mconMenu_Edit_Recipe_MedicalRecord = 3232          '���Ӳ�������

'ҩ���Զ���ҩ����
Public Const mconMenu_AutoSend = 100000

'������ҩ����
Public Const mconMenu_Edit_Recipe_Guide = 4001
Public Const mconMenu_Edit_Recipe_OK = 4002
Public Const mconMenu_Edit_Recipe_Average = 4003
'ˢ��791
'�˳�191


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

''�����˵���¼�����ͣ�
Public Const mconMenu_InputPopup = 1000                  '¼������

'���ŷ�ҩ
Public Const mconMenu_Input_Dept_HosNumber = 1101             'סԺ��
Public Const mconMenu_Input_Dept_Name = 1102                  '����
Public Const mconMenu_Input_Dept_BedNumber = 1103             '����
Public Const mconMenu_Input_Dept_NO = 1104                    '���ݺ�
Public Const mconMenu_Input_Dept_Ident = 1105                 '����ID
Public Const mconMenu_Input_Dept_ReceiveNO = 1106             '��ҩ��
Public Const mconMenu_Input_Dept_BatchSendNO = 1107           '���ܷ�ҩ��
Public Const mconMenu_Input_Dept_Dept = 1108                  '��ҩ����
Public Const mconMenu_Input_Dept_ICCard = 1109                'IC��

'������ҩ
Public Const mconMenu_Input_Recipe_NO = 1201                    '���ݺ�(&1)
Public Const mconMenu_Input_Recipe_OPNO = 1202                  '�����(&2)
Public Const mconMenu_Input_Recipe_Name = 1203                  '����(&3)
Public Const mconMenu_Input_Recipe_IDCard = 1204                '���֤(&4)
Public Const mconMenu_Input_Recipe_ICCard = 1205                'IC��(&5)
Public Const mconMenu_Input_Recipe_MINo = 1206                  'ҽ����(&6)
Public Const mconMenu_Input_Recipe_HosNumber = 1207             'סԺ��(&7)

''�����˵��������б����ݣ�
Public Const mconMenu_ListPopup = 2000
Public Const mconMenu_List_OnlyShowDept = 2001              '����ʾ�����б�
Public Const mconMenu_List_ShowOther = 2002                 '��ʾ��ϸ����
Public Const mconMenu_List_ShowAll = 2010                   '��ʾ���п���
Public Const mconMenu_List_ShowClin = 2011                  '��ʾ�ٴ�����
Public Const mconMenu_List_ShowTech = 2012                  '��ʾҽ������
Public Const mconMenu_List_ShowArea = 2013                  '��ʾ����
Public Const mconMenu_List_ShowReject = 2014                '��ȡ�ܷ�ҩƷ
Public Const mconMenu_List_Sort = 2015                      '���Ű�����ʱ������


''�����˵���PASS��
Public Const mconMenu_PASS = 5000
Public Const mconMenu_PASS_Item = 5100
Public Const mconMenu_PASS_Spec = 5200

'�����˵���PIVA������
Public Const mconMenu_Look = 2000
Public Const mconMenu_Filter = 2100


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
