Attribute VB_Name = "mdlPubDefine"
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
Public Const conMenu_File_Preview = 102         'Ԥ�����Ӱ��¼(&V)
Public Const conMenu_File_Print = 103           '��ӡ���Ӱ��¼(&P)
Public Const conMenu_File_Excel = 104           '�����Excel...
Public Const conMenu_File_TypeManage = 105      '��ι���
Public Const conMenu_File_Exit = 191            '�˳�(&X)

'�༭�˵�
Public Const conMenu_Edit_NewItem = 302         '����Ŀ(&A)
Public Const conMenu_Edit_Modify = 303          '�޸�(&M)
Public Const conMenu_Edit_Delete = 304          'ɾ��(&D)
Public Const conMenu_Edit_FinOut = 305             '��ɽ���(&U)
Public Const conMenu_Edit_FinIn = 306              '��ɽӰ�(&B)
Public Const conMenu_Edit_FinRead = 307          '�������(&P)
'Public Const conMenu_Edit_Out = 308            '����(&U)
'Public Const conMenu_Edit_In = 309              '�Ӱ�(&B)
'Public Const conMenu_Edit_Read = 310          '����(&P)
Public Const conMenu_Edit_CancelOut = 311            'ȡ����ɽ���(&U)
Public Const conMenu_Edit_CancelIn = 312             'ȡ����ɽӰ�(&B)
Public Const conMenu_Edit_CancelRead = 313          'ȡ���������(&P)
Public Const conMenu_Edit_CheckOutSign = 314    '��֤�������ǩ��(&C)
Public Const conMenu_Edit_CheckInSign = 315    '��֤�Ӱ����ǩ��(&C)
'����˵�
Public Const conMenu_Report_Record = 601   '���Ӱ������ѯ


'�鿴�˵�
Public Const conMenu_View_ToolBar = 701              '������(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Public Const conMenu_View_StatusBar = 702            '״̬��(&S)

'�����˵�
Public Const conMenu_Help_Help = 901        '��������(&H)
Public Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_Help_Web_Forum = 9023      '������̳(&F)
Public Const conMenu_Help_Web_Mail = 9022       '���ͷ���(&M)
Public Const conMenu_Help_About = 991       '����(&A)��

'������������
'********************************************************************
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
Public gstrProductName As String            'OEM��Ʒ����
Public gstrSQL As String
Public gcnOracle As New ADODB.Connection
Public gobjRegister As Object               'ע����Ȩ����zlRegister
Public grsUserInfo As ADODB.Recordset
Public glngSys As Long
Public gstrDbaUser As String
'1,'����',2,'����',3,'һ������',4,'����',5,'��ǰ',6,'����',7,'��Ѫ',8,'Σ',9,'����',10,'Σ/��',11,'�ؼ�',12,'����'

Public gstrSysName As String
Public gstrPrivs As String
Public glngModul As Long

Public gobjEmr As Object

