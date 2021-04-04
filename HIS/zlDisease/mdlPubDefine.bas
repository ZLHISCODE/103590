Attribute VB_Name = "mdlPubDefine"
Option Explicit

Public Const conMenu_FilePopup = 1    '�ļ�
Public Const conMenu_ManagePopup = 2    '����
Public Const conMenu_EditPopup = 3    '�༭
Public Const conMenu_ReportPopup = 4    '����
Public Const conMenu_PlugPopup = 6    '��ӳ��򣻼��鼼ʦ����վʹ��6100-6199
Public Const conMenu_ViewPopup = 7    '�鿴
Public Const conMenu_ToolPopup = 8    '����
Public Const conMenu_HelpPopup = 9    '����

'�ļ��˵�
Public Const conMenu_File_Open = 100            '*��(&O)��
Public Const conMenu_File_PrintSet = 101        '*��ӡ����(&S)��
Public Const conMenu_File_Preview = 102         '*Ԥ��(&V)
Public Const conMenu_File_Print = 103           '*��ӡ(&P)
Public Const conMenu_File_RowPrint = 121        '��¼��ӡ(&R)
Public Const conMenu_File_Parameter = 181       '*��������(&M)
Public Const conMenu_File_Modify = 3003         '*�޸�(&M)
Public Const conMenu_File_Delete = 3004         '*ɾ��(&D)
Public Const conMenu_File_Excel = 104           '�����&Excel��
Public Const conMenu_File_ExportToXML = 192     '���ΪXML�ĵ�
Public Const conMenu_File_Exit = 191            '�˳�

'�༭�˵�
Public Const conMenu_Edit_Reuse = 3009
Public Const conMenu_Edit_NewItem = 3001    '*����Ŀ(&A)
Public Const conMenu_Edit_Audit = 3010      '*���/У��(&V)
Public Const conMenu_Edit_Refuse = 3004        '�ܾ�
Public Const conMenu_Edit_NewTable = 3001    '������
Public Const conMenu_Edit_Add = 3002         '������ע
Public Const conMenu_Edit_Modify = 3003     '*�޸�(&M)
Public Const conMenu_Edit_EditInfo = 3564    '*����������Ϣ(&E)
Public Const conMenu_Edit_Delete = 3004     '*ɾ��(&D)
Public Const conMenu_Edit_Send = 3013       '*����(&G)
Public Const conMenu_Edit_Untread = 3014    '*����(&R)
Public Const conMenu_Edit_ApplyTo = 3062     '*���ÿ���(&T)
Public Const conMenu_Edit_Request = 3063     '����Ҫ��(&R)
Public Const conMenu_Edit_Compend = 3064     '*���ݹ���(&F)
Public Const conMenu_Edit_ElementChange = 3065      '*Ҫ����������
Public Const conMenu_Edit_Privacy = 3093     '*������˽��������

'��ͼ�˵�
Public Const conMenu_View_ToolBar = 701              '������(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Public Const conMenu_View_StatusBar = 702            '״̬��(&S)
Public Const conMenu_View_DiseaseRegist = 3031
Public Const conMenu_View_Refresh = 791              '*ˢ��(&R)

'���߲˵�
Public Const conMenu_Tool_Send = 3013            '����(&S)
Public Const conMenu_Tool_Transfer = 213        'ת��
Public Const conMenu_Tool_Finish = 3010          '���
Public Const conMenu_Tool_OK = 225              'ȷ��Ϊ��Ⱦ��
Public Const conMenu_Tool_NO = 3021                 '�Ǵ�Ⱦ��
Public Const conMenu_Tool_ViewReport = 7045          '�鿴�����鱨��
Public Const conMenu_Tool_Cancel = 3565
Public Const conMenu_Tool_Incept = 252    '����(&I)
Public Const conMenu_Tool_Refuse = 3004    '�ܾ�(&R)
Public Const conMenu_Tool_Aduit = 3010

'�����˵�
Public Const conMenu_Help_Help = 901        '*��������(&H)
Public Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_Help_Web_Forum = 9023      '������̳(&F)
Public Const conMenu_Help_Web_Mail = 9022       '*���ͷ���(&M)
Public Const conMenu_Help_About = 991       '����(&A)��


Public Const HWND_TOPMOST = -1              '��ǰ��
Public Const ID_EDIT_COPY = 323                      '����

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


