Attribute VB_Name = "mdlInsure"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
'    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;3-������������GetNextNO();
'    99-���н������Ӹ��Ӳ���(���°�)
Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As String
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_BROWSEFORCOMPUTER = &H1000  'Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000   'Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000 'Browsing for Everything
Private Const CSIDL_NETWORK As Long = &H12

Private Const MAX_PATH = 260
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2
Private Const LVM_SETCOLUMNWIDTH = &H101E

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'���뷨����API----------------------------------------------------------------------------------------------
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Public Const KLF_REORDER = &H8

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

'���ı������м��ܻ���ܵĺ���
Public Declare Function EncryptStr Lib "FTP_Trans.dll" (ByVal SourceStr As String, ByVal Key As String, ByVal IsEncrypt As Boolean) As String

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    վ�� As String
End Type
Public UserInfo As TYPE_USER_INFO

Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public glngSys As Long                      'ϵͳ��Ų���
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSQL As String                    '������Ϊ������ʱSQL���

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrDbUser As String                 '��ǰ���ݿ��û�
Public gstrUserName As String               '��ǰ�û�����
Public gstr��λ���� As String
Public gbln�������� As Boolean              '����ר��,���ڷ����Ƿ�Ϊ��������
Public gstr���ⲡ�� As String               '���ⲡ����
Public gintDebug As Integer                 '�����ע����ж�ȡ�ĵ��Ա�־

Public gstrMatchMethod As String    'ƥ�䷽ʽ:0��ʾ˫��ƥ��
Public gstrDec As String

Public gintInsure As Integer
Public gstrInsure As String         '��¼������ʹ�õ�ҽ���ӿ�
Public gstrҽԺ���� As String * 10               'ҽԺ���
Public gstrҽ���������� As String
Public glngReturn As Long           '�����������ر�־

Public mintOrder As Integer         '��ǰҽ���ӿڶ�������
Public gclsInsure As clsInsure
Public gobjInsure_Obj() As Object   '���������Ѵ򿪵�ҽ����������
Public gobjInsure_Name() As String  '���������Ѵ򿪵�ҽ����������
Public glngInstanceCount As Long    'ʵ��������94352

Public Type T��������
    ����ID       As Long
    ���         As Long
    סԺ����     As Long
    �ʻ��ۼ�����   As Currency
    �ʻ��ۼ�֧��   As Currency
    �ۼƽ���ͳ��   As Currency
    �ۼ�ͳ�ﱨ��   As Currency
    ����         As Currency
    �ⶥ��         As Currency
    ʵ������     As Currency
    �������ý��   As Currency
    ȫ�Էѽ��   As Currency
    �����Ը����   As Currency
    ����ͳ����   As Currency
    �Żݽ��       As Currency
    ͳ�ﱨ�����   As Currency
    �����Ը����   As Currency
    �����ʻ�֧��   As Currency
    ֧��˳���     As String
    ��ҳID         As Long
    ��;����       As Long
    סԺ����       As Long

    ����ͳ���Ը� As Currency
    
    '������(20060711):��������Ϊ�½�������ҽ������
    ����Աͳ��֧�� As Currency
    ����Ա������λ�� As Currency
    ����Ա����GGZF As Currency
    ����Ա�������� As Currency
    ����Ա�������޶� As Currency
    ��Աְ��       As String
    
    ����ҽ��ͳ��֧�� As Currency
    ����ҽ�Ʊ������� As Currency
    ����ҽ�Ʊ���ͳ���Ը� As Currency
    ����ҽ�Ʊ������� As Currency
End Type
Public g�������� As T��������           '����Ԥ����֮�����Ľ������������д���ս����¼
Public gcol������� As New Collection   '����Ԥ����֮�����Ľ������������д���ս������
                                        'ÿ����ԱΪһ�����飬����Ϊ���Ρ�����ͳ���ͳ�ﱨ��������
Public glngModul As Long

Public Enum ҽ��Enum
    TYPE_������ = 10
    TYPE_�������� = 11                  'Modified by ZYB ##2002-10-28
    TYPE_�����ɽ = 12
    TYPE_��������ɽ = 13
    TYPE_���������� = 14
    TYPE_�ɶ��� = 20
    TYPE_�ɶ����� = 21
    type_�ɶ����� = 22
    TYPE_�ɶ��ϳ� = 23
    type_���� = 24
    TYPE_�Ĵ�üɽ = 25
    TYPE_��ɽ = 26
    TYPE_����ʡ = 30
    TYPE_������ = 31
    TYPE_���Ͻ�ˮ = 32
    TYPE_�Թ��� = 40
    TYPE_������ = 43
    TYPE_������ = 50
    TYPE_�Ͻ� = 51
    TYPE_������ = 61
    TYPE_�Ͼ��� = 65
    TYPE_�������� = 70
    
    'Modified by ���� 20031218 ����������
    TYPE_����ʡ = 71
    TYPE_������ = 72
    TYPE_��ƽ�� = 73
    TYPE_������ = 80
    TYPE_ͭ�� = 81
    TYPE_������ = 82
    TYPE_���������� = 83
    
    'modified by �̴���(12,84,85,87,88,89,90,92,93,94,97,98,99)
    TYPE_���� = 84
    TYPE_���� = 85
    TYPE_�ش�У԰�� = 86    '���˺�:200403
    TYPE_��Ԫ = 87
    TYPE_���� = 88
    TYPE_���� = 89
    TYPE_�¶� = 90
    TYPE_���� = 1
    TYPE_�����山 = 91          '���˺�:20040705
    TYPE_�������� = 92
    TYPE_���� = 93
    TYPE_�㽭 = 94
    TYPE_�˰� = 95
    TYPE_��Ҧ = 99
    TYPE_ǭ�� = 96
    
    TYPE_���� = 98
    TYPE_������Ժ = 97
    TYPE_����ũ�� = 100
    TYPE_������ = 102
    TYPE_��ͨ = 103
    
    TYPE_�ɶ���ũҽ = 101
    TYPE_�ɶ��ڽ� = 106
    TYPE_��Ԫ���� = 107     'ͬʱҲ��üɽҽ��
    '���˺�:20050419
    TYPE_�ϳ����� = 108     '���Ԫ����ҽ��һ��.
    
    TYPE_�山ũҽ = 120
    TYPE_��Ϫũҽ = 121
    TYPE_���� = 122
    TYPE_�ɶ����� = 140
    TYPE_�ٲ׷��� = 141
    TYPE_�˳ɺ˹�ҵ = 142
    TYPE_�������� = 143
    TYPE_���� = 181
    '                      ������ɽ��
    TYPE_ɽ�� = 130      '�¶���20050304
    
    '��������Ϊ���¹淶ʵ�к������ӵĽӿ�
    TYPE_���� = 401      '�¶���20050401
    TYPE_���ũҽ = 402
    TYPE_�Ĵ��Թ� = 407 '�°��Թ�ҽ��
    
    TYPE_ͭɽ�� = 408 '�¶� 20050620
    '����Ϊ�������п���
    type_ͭ����ҽ = 602  ' �¶�: 20050530
    TYPE_����ʡũ�� = 669
End Enum

Public Enum ����Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
    Busi_TransferSwap
    Busi_TransferDelSwap
    Busi_WipeoffMoney
    Busi_SettleSwap
    Busi_SettleDelSwap
    Busi_ComeInSwap
    Busi_LeaveSwap
    Busi_TranChargeDetail
    Busi_LeaveDelSwap
    Busi_RegistSwap
    Busi_RegistDelSwap
    Busi_ComeInDelSwap
    Busi_ModiPatiSwap
    Busi_ChooseDisease
    Busi_IdentifyCancel
End Enum

Public Enum ���Enum
    balan���� = 10
    balan��Ժ = 20
    balanԤ�� = 30
    balan���� = 40
End Enum

Public Enum �����֤Enum
    id�����շ� = 0
    id��Ժ�Ǽ� = 1
    id�ʻ����� = 2
    id�Һ� = 3
    id���� = 4
    id����ȷ�� = 5
End Enum

Public Enum ҽԺҵ��
    'Modified by ZYB 2005-08-08 ȡ�������������������ó�����ʹ�ã����Ա����Ŀ����Ϊ��ҽ��������Ҫ��������������������support����������ϡ�supportסԺ��������
    'ԭ�򣺽���������ԭʼ��������һ�µ�����
    '�µĽ���취��ʹ��GetCapability�������м���Ƿ�֧�ֽ������ϣ����strAdvance��Ϊ�գ����ʾ���ĳ���ض��Ľ��㷽ʽ����ҽ���Ƿ�֧��ȫ�ˣ������֧�֣����ʾ�ý��㷽ʽȫ��Ϊ�ֽ�
    support�����˷� = 1
    support�����˸����ʻ� = 3
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    
    support����Ԥ�� = 0
    
    supportԤ���˸����ʻ� = 2
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��

    support���������շ� = 22        '�����������֤�󣬿ɽ��ж���շѲ���
    support�����շ���ɺ���֤ = 23  '�������շ���ɣ��Ƿ��ٴε��������֤
    
    supportҽ���ϴ� = 24            'ҽ����������ʱ�Ƿ�ʵʱ����
    support�ֱҴ��� = 25            'ҽ�������Ƿ���ֱ�
    support��;������������ϴ����� = 26 '�ṩ�����ϴ��������ݵĽ��㹦��
    support��������ѽ��ʵļ��ʵ��� = 27 '�Ƿ�����������ʵ��ݣ�����õ����Ѿ�����
    
    support�����ݳ������� = 28
    support��Ժ��ʵ�ʽ��� = 29      '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���

    '�൥���շ�ʱ��Ҫע�������
    '����ҷ��Ƕ൥�ݣ���ҽ���Ǳ߱���Ϊһ�ŵ��ݵģ���Ҫ�������¿��ǣ�
    '1�����ϵͳ������78-���ŵ����շѷֱ��ӡ��Ϊ�棬��������ж൥���շ�
    '2.1���˷�ʱ��������ս����¼�ı�ע�к����൥���շѡ������߸õ����Ƕ൥���շѣ�Ʊ�ݴ�ӡ���ݴ���1����¼�������ִ��
    '2.2������˷�ʱ�õ����ǵ����շѣ�Ʊ�ݴ�ӡ����С�ڵ���1����¼��������ȡ�ò��˸õǼ�ʱ����ͬ�����е��ݺų�������ʾ����ԱӦ��ͬʱ�˷Ѻ�����ȡ�µķ���
    support�൥���շ� = 30          '�Ƿ�֧�ֶ൥���շ�
    
    support�����շѴ�Ϊ���۵� = 31  '�������շѵ�תΪ���۵����棬�޸���ǰ�̶��ж�ĳ��ҽ���ķ�ʽ
    support�����ֳ�����ϸ = 32    '�������סԺ���ʴ�����ÿ����ϸ���в��ֳ���
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧�ֵ�����ͨ�������ϴ���
    supportסԺ�������� = 34        'HISʼ����ΪסԺ֧�ֽ������ϣ������֧����ҽ���ӿ��ڲ��������ؼټ��ɣ����Ӹò�����Ϊ�����GetCapability�����������ֽ��㷽ʽ�Ƿ�֧��ȫ��
    support�������� = 35            '�Ƿ����������ʣ�����Ա����Ҫӵ�и������ʵ�Ȩ�ޡ��˲���ȱʡΪ�棬��֧�ֵĽӿ��赥������
    support����_ָ��סԺ���� = 36   '�Ƿ�֧��ָ��סԺ��������ҽ������
    support����_ָ�����ڷ�Χ = 37   '�Ƿ�֧��ָ���������ڷ�Χ����ҽ������
    support����_����Ӥ�������� = 38 '�Ƿ���������Ӥ��������
    Support�൥���շѱ���ȫ�� = 39  '�൥���շѱ���ȫ��
    'ֻ�ܸ�ҽ�������ڵ�ҽ����ʼ�������д������Ҫ���ɲ�������Ļ�������Ҫ�����ø�����������ҽ�����ķ���ֵ��
    Support��ʼ��ʧ���ѻ����� = 40  '����������ǰ��ҽ����ʼ��ʧ�ܣ��Ժ��ٽ��г�ʼ���������ѻ�����ֻ����HIS��
    support������� = 41            '�Ƿ�֧������ҽ�����˵ļ��ʷ���ʹ��סԺ���������
    support����_ָ������ = 42           '�Ƿ������ڽ������ý�����ָ������
    support����_ָ��������Ŀ = 43       '�Ƿ������ڽ������ý�����ָ��������Ŀ
    support����_�������ú���ýӿ� = 44 '�Ƿ��ڽ������ú�ŵ���סԺ������㣿
    support����_����������ú���ýӿ� = 49 '�������:�Ƿ��ڽ������ú�ŵ�������������㣿
    support����_ָ���������� = 45       '�Ƿ������ڽ������ý�����ָ����������

        supportҽ���ӿڴ�ӡƱ�� = 46                    'HIS��Ȼ���ϸ����Ʊ�ݵ�����ӡ����ҽ�����Ʊ�ݵĴ�ӡ�������շ�һ��ֻ��ӡһ�ŷ�Ʊ�������շѸ���ϵͳ�趨�������ŷ�Ʊ
        support�൥��һ�ν��� = 47                              '���ҽ��֧������������㽫���е��ݷ��صı����ܶ���η�̯���������ϣ���������ʱҲ����ˡ�
                                                                                '�������̣��������ʱ�����һ�ŵ���ʱ�����ϴ�������ʱ�ڵ�һ�ŵ���ʱ����
        supportҽ��ȷ���������� = 48                    '��Դ�ڱ���ҽ����������ҽ������ʱ������ʾѯ����ҽ���ڻ���ҽ���⴦��������������ڷ��ü�¼��ժҪ��

    support�Һű��봫����ϸ = 61
    support�����Һ� = 62
    support����ķ� = 63
    
    
    supportʵʱ��� = 60                'ָ���Ƿ����ʵʱ�����صĽӿں�����CheckClinicGuideline��CheckSettleGuideline��CheckItem
    '�������������ڿ����Ƿ������׼��Ŀ����
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
    support�����ѽɿ���� = 64            '���շ�ʱ,����շѲ�����"�����нɿ�������ۼƿ���"Ϊtrueʱ,ͬʱ��ҽ������ʱû������ɿ���ʱ�������û�
        support�����˷Ѻ��ӡ�ص� = 65                          '�����˷Ѻ����շ�ģ������Զ��屨����ɻص���ӡ
        support�������Ϻ��ӡ�ص� = 66                          'סԺ�������Ϻ��ɽ���ģ������Զ��屨����ɻص���ӡ

    support�ϴ����ﵵ�� = 70                    '������ҽ������ʱ���Ƿ����TranElecDossier����������ﲡ�˵��Ӿ���/���ӵ������ϴ�
        
    support����_���ֵ��ݽ��� = 80                                   'Ԥ���㡢���㶼ֻ����һ��ҽ������
    support�ҺŲ���ȡ������ = 81    '�ڹҺ�ʱ����ʹ��ҽ����ȡ������
    
    support������ȫ�� = 82 '�����˷�ʱ�������ݽ����˷ѣ�86176
    support�൥�ݷֵ��ݽ��� = 83 '�൥��һ�ν��㰴���ݽ���ҽ��������86321
    support������;���� = 84  'סԺ����ʱ�Ƿ�����ҽ������������;����,�����:81661
End Enum

Public gblnLED As Boolean '�Ƿ�ʹ��LED�����豸
Private rsInsure As New ADODB.Recordset

Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Public gobjRegister As Object               'ע����Ȩ����zlRegister

'--������������ʾ��
Public Function GetNewConnection() As ADODB.Connection
    Dim cnOracleTmp         As ADODB.Connection
    Dim strUser             As String
    Dim strPwd              As String
    Dim strServer           As String
    Dim strMsg              As String
    If gobjRegister Is Nothing Then
        On Error Resume Next
        Set gobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
        If gobjRegister Is Nothing Then
            Set gobjRegister = CreateObject("zlRegister.clsRegister")
            Err.Clear
            If gobjRegister Is Nothing Then
                'MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
                '����ʧ�ܣ�ʹ�þ�ģʽ���д�������
                Set cnOracleTmp = New ADODB.Connection
                cnOracleTmp.Open gcnOracle.ConnectionString
                Set GetNewConnection = cnOracleTmp
                Exit Function
            End If
        End If
    End If
    
    '���´����ӣ���һ��������9����ϵͳ�Զ������ϴ����ӽ��д���
    '����У��ɹ���ʹ����ģʽ���д���
    '    Public Function GetConnection(ByVal strServer As String, ByVal strUserName As String, ByVal strPassword As String, ByVal blnTransPassword As Boolean, _
    '     Optional ByVal bytProvider As enuProvider = MSODBC, Optional ByRef strError As String = "���뷵�ش�����Ϣ", Optional ByVal blnSaveAccount As Boolean = True) As ADODB.Connection
    '���ܣ� ��ָ�������ݿ⣬��������ʵ������ADO���Ӷ���(�����10.35.10��ǰ�����룬���µ�ת�������������),��������������û��������뵽����gstrServer��gstrUserName��gstrPassword
    '������ strServer       :�������������߿���ֱ��ָ��IP:Port/SID
    '       strUserName     :�û���
    '       strPassword     :����
    '       blnTransPassword:�Ƿ��������ת��
    '       bytProvider     :�����ݿ����ӵ����ַ�ʽ,0-msODBC��ʽ,1-OraOLEDB��ʽ
    '       strError        :����ʧ�ܺ����ָ���˴˲������򷵻ش�����Ϣ��δָ��ʱֱ�ӵ�����ʾ��Ϣ��
    '       blnSaveAccount  :�����û��������롢����������ȫ�ֱ�����һ�㣬���ڵ�¼����ʱ���棬���ӿ�ReGetConnection��GetUserName��GetServerName��GetPassword��LoginValidateʹ�ã�
    '���أ� ���ݿ�򿪳ɹ������Ӷ����״̬���Է���adStateOpen(1),ʧ���򷵻�AdStateClosed(0)
    Set GetNewConnection = gobjRegister.ReGetConnection(9, strMsg)
End Function
'================================================================================================================================
'=����˵������ʾ������Ϣ�������ڴ�������Ĺ����м�������ع���
'=��Σ�
'=  1.strErrMsg��������ʾ��Ϣ��
'=  2.mbs��������ʾģʽ��Ĭ��ΪVbMsgBoxStyle.vbInformation��
'=  3.strTitle��������ʾ���⡣
'=  4.blnTran���Ƿ�ִ������ع���
'=���Σ�(��)
'=���أ�(VbMsgBoxResult)��ʾ���ѡ��ֵ��
'=ע�⣺
'=  1.��HIS��û�п�������ĵط���blnTran=False�������������ĵط���blnTran=True��
'=  2.����Һš�����Һų�����������㡢�����������б��봫�����blnTran=True��
'=  3.��Ժ�Ǽǡ���Ժ�Ǽǳ�������Ժ�Ǽǡ���Ժ�Ǽǳ����пɴ������blnTran=True��
'=  4.��סԺ���㡢סԺ��������б��봫�����blnTran=True��
'=  5.������������㡢�����ϴ���סԺ�������ȷ����У����贫��blnTran������blnTran=
'================================================================================================================================
Public Function ErrMsgBox(strErrMsg As String, Optional mbsStyle As VbMsgBoxStyle = vbInformation, Optional strTitle As String = "") As VbMsgBoxResult
    Dim blnTran As Boolean
On Error GoTo ErrH
    '��ȡ����״̬
    blnTran = gclsInsure.zlTranState
    '�ع�����
    If blnTran Then gcnOracle.RollbackTrans
    '��ʾ������Ϣ
    ErrMsgBox = MsgBox(strErrMsg, mbsStyle, strTitle)
    '���¿�������
    If blnTran Then gcnOracle.BeginTrans
    '��������Թ�����ȥ
    DebugTool strErrMsg
    Exit Function
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Function
End Function

Public Function GetErrInfo(strCode As String, ByVal intinsure As Integer) As String
'���ܣ����ݴ�����뷵�ش�����Ϣ
'������bytType=�������,strCode=�������
    Dim rsTmpErr As New ADODB.Recordset
    
    strCode = Trim(strCode)
    Select Case intinsure
        Case TYPE_����ʡ, TYPE_������, TYPE_���Ͻ�ˮ
            Select Case strCode
                Case "0000":      GetErrInfo = "����"
                Case "0001":      GetErrInfo = "�޷���ȡ�����ļ�����رձ�������������ļ���"
                Case "0002":      GetErrInfo = "��Ӧ�ó������������ʧ��(�޷��ҵ�Ӧ�ó��������),��ȷ��Socket Server�Ƿ���������!"
                Case "0003":      GetErrInfo = "Ӧ�ó�������������޷���ɽ���!"
                Case "0004":      GetErrInfo = "�޷��õ�ϵͳ������Ϣ!"
                Case "0005":      GetErrInfo = "�Ҳ����α����������ĵĳ������������������������!"
                Case "0009":      GetErrInfo = "���������ÿ��Ŷ�Ӧ�ķ����ı��"
                Case "1":         GetErrInfo = "�ն��豸��֧�ִ˹���"
                Case "10":        GetErrInfo = "��֤����,����ĸ����������"
                Case "1001":      GetErrInfo = "˳��ų��ȷǷ�"
                Case "1002":      GetErrInfo = "�շ���Ŀ�������Ƿ�"
                Case "1003":      GetErrInfo = "�շ���Ŀ����Ƿ�"
                Case "1004":      GetErrInfo = "������۸���Ϊ��"
                Case "1005":      GetErrInfo = "������۸���С�ڵ���0"
                Case "1006":      GetErrInfo = "����������Ƿ�"
                Case "11":        GetErrInfo = "֧������,����"
                Case "1101":      GetErrInfo = "˳��Ŵ���"
                Case "1102":      GetErrInfo = "�����ѽ��㲻���ٴ��ݷ�����ϸ"
                Case "1103":      GetErrInfo = "û�м�������Ҫ�޸ĵķ�����ϸ����!�����������������ȷ!"
                Case "1104":      GetErrInfo = "�޸ķ�����ϸ���ϳ���!"
                Case "1105":      GetErrInfo = "�ò���סԺ�����ѽ����!������Ŀ����ͬ���ദ��"
                Case "12":        GetErrInfo = "֧������,�û�����ʼ��ʧ��"
                Case "13":        GetErrInfo = "֧������,SAM����ʼ��ʧ��"
                Case "14":        GetErrInfo = "֧������,�û�����֤MAC1ʧ��"
                Case "15":        GetErrInfo = "֧������,SAM����֤MAC2ʧ��"
                Case "16":        GetErrInfo = "������,��ȡ���ʧ��"
                Case "17":        GetErrInfo = "���¶�̬��Ϣ,�û������´���"
                Case "18":        GetErrInfo = "δ֪�����"
                Case "19":        GetErrInfo = "���¶�̬��Ϣ,PSAM����ȡ����"
                Case "2":         GetErrInfo = "���׳�ʼ��,��ⲻ���ն��豸+���豸����"
                Case "20":        GetErrInfo = "�޴�������"
                Case "2001":      GetErrInfo = "�����˻�������ƷǷ�,�����ٽ��н���!"
                Case "21":        GetErrInfo = "֧������,TACУ�����"
                Case "2101":      GetErrInfo = "�����Ѱ����Ժ����,�����ٽ��н���"
                Case "2102":      GetErrInfo = "��������δͨ������Ժ�ڼ�ķ���Ϊȫ�Է�"
                Case "2103":      GetErrInfo = "���ý���ʱ��⵽�洢���̵����������˳��š�λ������ȷ!"
                Case "22":        GetErrInfo = "Ȧ�潻��,MAC1У�����"
                Case "2201":      GetErrInfo = "������ȡ���˵�֧������޷����з��ý���<bnzxx>��"
                Case "2202":      GetErrInfo = "������ȡ���ⲡ�˵�֧������޷����з��ý���<by21bzxx>��"
                Case "2203":      GetErrInfo = "���ý���ʱ��By10cyjsb��д���ݳ���!"
                Case "2204":      GetErrInfo = "���û���ʧ��!"
                Case "2205":      GetErrInfo = "Ԥ����ʧ��!"
                Case "2206":      GetErrInfo = "����Ա�洢����ִ�г���!"
                Case "2207":      GetErrInfo = "�����Ƚ��з��õ������֮���ٽ��н������!"
                Case "2208":      GetErrInfo = "û����Ч��Ԥ�����¼���޷�����!"
                Case "2209":      GetErrInfo = "סԺ�����Ѿ������·����ƣ�ϵͳֻ���屾�µķ���!"
                Case "2210":      GetErrInfo = "δ��ѯ������Ǽǻ�������ڳ���������˵�������ޣ�"
                Case "2211":      GetErrInfo = "��ǰ���˽����¼�������һ�ν��㣬����ѵ�ǰ�����¼֮������н����¼����֮����ܽ��л���ҵ�����!"
                Case "23":        GetErrInfo = "Ȧ�潻��,TACУ�����"
                Case "24":        GetErrInfo = "Ȧ�潻��,�û�����ʼ��ʧ��"
                Case "25":        GetErrInfo = "Ȧ�潻��,�û�����֤MAC2ʧ��"
                Case "29":        GetErrInfo = "��������ʧ��"
                Case "3":         GetErrInfo = "���׳�ʼ��,��ⲻ��PSAM��"
                Case "30":        GetErrInfo = "�޴˽��״���"
                Case "3001":      GetErrInfo = "����ҽԺ����Ƿ������������ļ������ã�"
                Case "3002":      GetErrInfo = "�����Ͳ�����δ�忨���루�أ��忨��"
                Case "3003":      GetErrInfo = "�޷���ȡ������Ϣ�������ԣ�"
                Case "31":        GetErrInfo = "���ͽ�������ʧ��,����ͨѶ�˿�"
                Case "3100":      GetErrInfo = "�����Ѱ���סԺ�Ǽǣ����ܽ��и���ҵ��"
                Case "3101":      GetErrInfo = "������ʹ�õĿ��Ƿ�������ƾ������ҽ��������"
                Case "3102":      GetErrInfo = "����������ҽԺδ���н��㣬�޷����и�ҵ��"
                Case "3103":      GetErrInfo = "δ�����˿��Ļ������ϣ�����������!"
                Case "3104":      GetErrInfo = "���ܻ�ȡ���˵ĳ������ڣ����ܽ���Ժ�Ǽǣ�"
                Case "3105":      GetErrInfo = "�ò�����/סԺ���Ѿ���ռ�ã���������סԺ��/�����ţ�"
                Case "3106":      GetErrInfo = "���ܼ�������ҽԺ�ȼ������ݣ����ʵҽԺ���룡"
                Case "3107":      GetErrInfo = "������鲻��ȷ,����������!"
                Case "3108":      GetErrInfo = "�޷�����������Ⱥ��ı�������!"
                Case "3109":      GetErrInfo = "ִ��������Ⱥ�Ĵ洢���̡������ִ���!"
                Case "3110":      GetErrInfo = "ҽ�����ı�����IC��ʵ�ʼ�¼��ҽ�����ı��벻һ��!"
                Case "3111":      GetErrInfo = "����Ա�������ʱ����!"
                Case "3128":      GetErrInfo = "��ǰ���˵ľ�������Ǽ������ִ�м���תסԺ������"
                Case "32":        GetErrInfo = "������Ӧ���ݳ�ʱ���ױ�ȡ��,����ͨѶ�˿�"
                Case "33":        GetErrInfo = "У����Ӧ����(ETX)����"
                Case "34":        GetErrInfo = "У����Ӧ����(LRC)����"
                Case "35":        GetErrInfo = "У����Ӧ����(STX)����"
                Case "36":        GetErrInfo = "У����Ӧ���ݴ�����Կ����"
                Case "37":        GetErrInfo = "������Ӧ����ʧ��,����ͨѶ�˿�"
                Case "38":        GetErrInfo = "δ֪����,������ȡ��"
                Case "4":         GetErrInfo = "���׳�ʼ��,PSAM����ȡ����"
                Case "40":        GetErrInfo = "ͨѶ����"
                Case "4001":      GetErrInfo = "�ļ����Ƿ�"
                Case "4002":      GetErrInfo = "д�ļ����̳���"
                Case "41":        GetErrInfo = "Ȧ�潻����֤ʧ�ܣ��뽫������ҽ�����Ĵ���"
                Case "4101":      GetErrInfo = "�޷�����ϸ��Ϣ"
                Case "42":        GetErrInfo = "�ſ�����"
                Case "5":         GetErrInfo = "���׳�ʼ��,��ⲻ���û���"
                Case "5001":      GetErrInfo = "֧��ԭ�򳤶ȷǷ�,֧��ʧ��"
                Case "5002":      GetErrInfo = "֧�����Ӧ�ô���0,֧��ʧ��"
                Case "5003":      GetErrInfo = "֧������ڿ������,֧��ʧ��"
                Case "5004":      GetErrInfo = "д��ʧ��"
                Case "5101":      GetErrInfo = "�����ѳ�Ժ��˳��Ŵ���"
                Case "5102":      GetErrInfo = "������Ǽ�ʱ��ʹ�õĿ�����"
                Case "5103":      GetErrInfo = "�޷��õ���Ч�ĸ����˻�֧�����ֽ�֧����"
                Case "5104":      GetErrInfo = "��Ա��δ���з��ô���/���߼��������ò�Ա��֧�����ݡ�"
                Case "6":         GetErrInfo = "���׳�ʼ��,�Ǳ�ϵͳ��"
                Case "6101":      GetErrInfo = "�޷����ܴ���"
                Case "6102":      GetErrInfo = "��֧������ڷ����ܶ�޷�֧��"
                Case "7":         GetErrInfo = "�û�����ȡ����"
                Case "7101":      GetErrInfo = "���������ݿ�����ʧ��,��ȷ�����糩ͨ�Լ�NET8����������ȷ!"
                Case "7102":      GetErrInfo = "��ǰ�û����ݿ�����ʧ��,��ȷ��Socket Server�Ƿ���������!"
                Case "7103":      GetErrInfo = "ʡҽ��������������ʧ��"
                Case "7104":      GetErrInfo = "�����Ѿ���ȡ����"
                Case "7106":      GetErrInfo = "��ҽ��������������ʧ��"
                Case "7107":      GetErrInfo = "ʡ��ҽ��������������ʧ��!"
                Case "8":         GetErrInfo = "��֤���Ų���"
                Case "8001":      GetErrInfo = "��ȡ���˻�����Ϣ����"
                Case "8002":      GetErrInfo = "������Ч��������¼δ���㣬����������"
                Case "8003":      GetErrInfo = "�Ѿ�������Ժ�Ǽ�"
                Case "8004":      GetErrInfo = "����Ա�޷�����ҽ�ƴ���"
                Case "8005":      GetErrInfo = "���ҽ�ƴ��������ʸ�ʱ��ϵͳ����"
                Case "8006":      GetErrInfo = "���ȵ�ҽ�����Ľ����ʸ�����"
                Case "8007":      GetErrInfo = "�������ʱ����"
                Case "8008":      GetErrInfo = "�����ڷ���״̬"
                Case "8009":      GetErrInfo = "���ý���ʱ��ϵͳ����"
                Case "8010":      GetErrInfo = "û����Ч��������¼"
                Case "8011":      GetErrInfo = "��δ�����������¼�����ʵ"
                Case "8012":      GetErrInfo = "ȫ�ԷѲ��ּӹҹ��ԷѲ��ִ��ڷ����ܶ���ʵ"
                Case "8013":      GetErrInfo = "��������¼����δͨ������Ϊ�����ܴ������㣬ȫ���Է�"
                Case "8014":      GetErrInfo = "����Ա����˷����޷�����ҽ�ƴ���"
                Case "8015":      GetErrInfo = "����Ա��ҽ�Ʊ����չ���Ա"
                Case "8016":      GetErrInfo = "����ԱΪҽ�Ʊ����չ���Ա"
                Case "8017":      GetErrInfo = "��ǰ�����ѽ���"
                Case "8080":      GetErrInfo = "ҽ��������δ�ڹ�ҽԺ��ͨ����ҽ��ҵ��,����ҽ��������ϵ"
                Case "8100":      GetErrInfo = "����סԺ��ʱ����С�ڹ涨����,�����������Ժ�Ǽ�"
                Case "9":         GetErrInfo = "��֤����,�û����������뱻��"
                Case "9001":      GetErrInfo = "Ӧ�÷�����ִ�д洢����/�������������"
                Case "9002":      GetErrInfo = "�������ӵ��������ݿ�(hisint/hisintkm),�޷����н��״���!"
                Case "9003":      GetErrInfo = "�򱾵����ݿ����ύ�����޸ĳ���,�޷������ݽ����ύ���߻ع�!"
                Case "9004":      GetErrInfo = "�ò��˵Ļ������ϻ�û�еǼǻ����Ѿ��ύ�ɹ�,�޷��ع�����!"
                Case "9005":      GetErrInfo = "���ݿ���δ�������ò��˵�δ�ύ����������,�޷��ύ����!"
                Case "9006":      GetErrInfo = "�ⲿӦ�ó������������ƺŵ�λ����Ϊ18λ!"
                Case "9201":      GetErrInfo = "��ѯ�ֶη�����ϸ��¼������!"
                Case "9202":      GetErrInfo = "��ѯסԺ���˴��������¼������!"
                Case "9203":      GetErrInfo = "��Ч�Ĳ�ѯ���!"
                Case "9204":      GetErrInfo = "���������������ʧ��,�޷��������µı����Ϣ!"
                Case "9205":      GetErrInfo = "���������Ϣ��ѯ����!"
                Case "9301":      GetErrInfo = "�޷���λ���˵�ҽ���������޷�����!"
                Case "9996":      GetErrInfo = "ʡ�������ݴ���ʧ��!"
                Case "9997":      GetErrInfo = "���������ݴ���ʧ��!"
                Case "9998":      GetErrInfo = "ʡ/���������ݴ���ʧ��!"
                Case "9000":      GetErrInfo = "Ӧ�÷����������쳣����"
                Case Else
                    GetErrInfo = "ҽ��֧�ֲ��ֳ��ִ���"
            End Select
            GetErrInfo = GetErrInfo & "[������롪" & strCode & "]"
        Case TYPE_�ɶ���
'            gstrSQL = "select errtext from errcode where code='" & strCODE & "'"
'            rsTmpErr.CursorLocation = adUseClient
'            rsTmpErr.Open gstrSQL, gcnSybase, adOpenKeyset
'            If Not rsTmpErr.EOF Then
'                GetErrInfo = IIf(IsNull(rsTmpErr!errtext), "δ֪ԭ��Ĵ���", rsTmpErr!errtext)
'            Else
'                GetErrInfo = "δ֪ԭ��Ĵ���"
'            End If
            '������(2006-01-16):���ڶ���ҽԺ��errcode���ǿյ�,���Բ��ܱ�֤������Ч�Ĵ�����Ϣ
            Select Case strCode
                Case "003":       GetErrInfo = "���߿����"
                Case "004":       GetErrInfo = "�����д�"
                Case "005":       GetErrInfo = "����ǰ��һ��"
                Case "007":       GetErrInfo = "�޽�������"
                Case "008":       GetErrInfo = "�޻�����Ժ��Ϣ�����ж�����Ժ�Ǽ���Ϣ"
                Case "009":       GetErrInfo = "�޻���ע����Ϣ"
                Case "020":       GetErrInfo = "���������޴˶���ҽԺ����"
                Case "021":       GetErrInfo = "������������ƽ����"
                Case "022":       GetErrInfo = "Ҫ�޸Ļ�ɾ���ķ�����ϸ�Ѿ�����"
                Case "023":       GetErrInfo = "���ʽ������ʻ����"
                Case "024":       GetErrInfo = "����ʱ����"
                Case "025":       GetErrInfo = "���ؿ����޴˽��ʼ�¼"
                Case "026":       GetErrInfo = "��ͳ������"
                Case "027":       GetErrInfo = "���ܰ�����Ժ�Ǽ�"
                Case "028":       GetErrInfo = "�ò�Աδ��Ժ"
                Case "029":       GetErrInfo = "�ò�Ա�ѽ���"
                Case "030":       GetErrInfo = "����δ�������� , ���ܰ����Ժ"
                Case "031":       GetErrInfo = "�޳�Ժ�Ǽǲ��ܰ�����Ժ�Ǽ�"
                Case "032":       GetErrInfo = "����ǼǺ��ѸĲ��ָܻ���Ժ"
                Case "033":       GetErrInfo = "�ָ���Ժ�Ǽǳ���"
                Case "034":       GetErrInfo = "δȡ�Ծ���ǼǺ�"
                Case "035":       GetErrInfo = "ʱ��ת������"
                Case "036":       GetErrInfo = "�������ݲ�һ��"
                Case "037":       GetErrInfo = "�޽�����Ϣ"
                Case "038":       GetErrInfo = "�����ܶ�Ϊ��"
                Case "039":       GetErrInfo = "���ʽ���С��0"
                Case "041":       GetErrInfo = "�������Ÿ�ʽ����"
                Case "042":       GetErrInfo = "��Ժ���㵱��δ������ݼ���"
                Case "043":       GetErrInfo = "�ÿ���Ч"
                Case "099":       GetErrInfo = "��������"
                Case Else
                                  GetErrInfo = "ҽ������δ֪����"
            End Select
            GetErrInfo = GetErrInfo
        Case TYPE_����������, TYPE_������
               Select Case Val(strCode)
                    Case 0: GetErrInfo = "����"
                    Case -2:      GetErrInfo = "�����ڴ��ϵͳ������������ϵͳ���ܽ��"
                Case -1001, -1003, -1004, -1005 - 1006 - 1007:
                        GetErrInfo = "ϵͳ�������������ȷ�������Ƿ���������!"
                Case -1002:
                        GetErrInfo = "���������Ӵ���,������ԭ�������:" & vbCrLf & _
                                     "    ��1�����粻ͨ" & vbCrLf & _
                                     "    ��2�������������ʧ��" & vbCrLf & _
                                     "    ��3���ͻ������ô���" & vbCrLf & _
                                     "    ��4���ͻ�������ҽԺ�������" & vbCrLf & _
                                     "����취Ϊ:ȷ�������Ƿ�������ȷ�Ϸ����Ƿ���������"
                Case -5555
                    GetErrInfo = "��������������IC���Ƿ�����������Ͳ�ƥ��!"
                '�ܺ�ȫ���� 2003-12-17
                Case -5556
                    GetErrInfo = "���Ų�һ�£�"
                Case -6001, -6002, -6003, -6004, -6005, -6007, -6008
                    GetErrInfo = "ϵͳ�������ݽ���ʱ����,����ϵͳ�ļ�package.dat" & vbCrLf & _
                                 "�ļ��⵽�ƻ��򴫵ݵĲ���ֵ����!"
                Case -6009
                    GetErrInfo = "�����е�ҽԺ��ź�ע���ҽԺ��Ų�һ��!"
                Case -6006, -7001
                    GetErrInfo = "ϵͳ���кϷ�����֤���󣬿�������ϵͳ��������" & vbCrLf & _
                                 "��ע������Ƿ�ʹ�ã�ҽ�����Ľ��������ҽԺ�Ľ���������!"
                Case 1001
                    GetErrInfo = "�����ڸñ���!"
                Case 1002
                    GetErrInfo = "���Ŵ��鿨��!"
                Case 1003
                    GetErrInfo = "ֹ�������֧�������鿨!"
                Case 1004
                    GetErrInfo = "��ҽԺ�ֵ䣬ҽԺ��Ŵ�!"
                Case 1005
                    GetErrInfo = "ҽԺ�ѱ�����!"
                Case 1007
                    GetErrInfo = "����ʱ����ڵ�ǰϵͳʱ�䣬Ӧ�ô���"
                Case 1008
                    GetErrInfo = "��������ظ���IC���ϵ����ݴ������鿨��"
                '�ܺ�ȫ���� 2003-12-17
                Case 1009: GetErrInfo = "У�����ݴ������η�Ϊ��������"
                Case 1011
                    GetErrInfo = "������Ϣ�вα��˵Ļ�����Ϣ�������鿨��"
                Case 1016
                    GetErrInfo = "���ķ�����ͣ�����ڽ��и��£��������Ӻ����ԣ�"
                Case 1020: GetErrInfo = "�Ƿ�����ҵ�����ʹ���Ӧ�ô���"
                Case 1022: GetErrInfo = "���������"
                Case 1023: GetErrInfo = "������סԺ��"
                Case 1024: GetErrInfo = "�������ز���"
                Case 1025: GetErrInfo = "���ķ�������뼰ʱ�����ķ������Ա��ϵ��"
                Case 1026: GetErrInfo = "�������кŴ���ҽԺϵͳ���Ƿ��������⵽�ƻ���"
                Case 1027: GetErrInfo = "�����˻����ⲻһ�£����鿨��"
                Case 1028: GetErrInfo = "�Ƿ�����ʱ�䣬���ķ�����ͣ��"
                Case 1030: GetErrInfo = "ҽԺ��������㷨����"
                Case 1031: GetErrInfo = "ҽԺ������˽����㷨����"
                Case 1032: GetErrInfo = "����������"
                Case 1033: GetErrInfo = "ҽԺסԺ�����㷨����"
                Case 1034: GetErrInfo = "ҽԺסԺ���˽����㷨����"
                Case 1035: GetErrInfo = "ͳ���ۼƴ���"
                Case 1036: GetErrInfo = "д����Ŵ�"
                Case 1037: GetErrInfo = "���Ŵ���"
                Case 1038: GetErrInfo = "�˻�������"
                Case 1039: GetErrInfo = "סԺû��סԺ������"
                Case 1041: GetErrInfo = "ת��û��ת�ﵥ��"
                Case 1042: GetErrInfo = "����������"
                Case 1043: GetErrInfo = "ҽԺ��ͳ�ڱ���Ϊ�㣬������סԺ��"
                Case 1044: GetErrInfo = "�˻���������Ĳ�����"
                Case 1045: GetErrInfo = "�˻����ͬ��"
                Case 1046: GetErrInfo = "ҽԺ����󲡽����㷨����"
                Case 1047: GetErrInfo = "ҽԺ����󲡳��˽����㷨����"
                Case 1048: GetErrInfo = "�����������"
                Case 1049: GetErrInfo = "ҽԺ������˽������ݴ���"
                Case 1050: GetErrInfo = "����ʱ�����������"
                Case 1052: GetErrInfo = "סԺ�����ѵǼ���Ժ��"
                Case 1053: GetErrInfo = "�ò���δ�Ǽǲ���Ժ��"
                Case 1054: GetErrInfo = "�ô������ƣ����Ѿ����"
                Case 1058: GetErrInfo = "�󲡵Ǽ�ʱ�󲡱��벻��Ϊ�գ�"
                Case 1059: GetErrInfo = "��ҽ�޶�ҽԺ����"
                Case 1062: GetErrInfo = "ת�ﵥ00000E�Ļ�������С��70��"
                Case 1063: GetErrInfo = "ת�ﵥ00000E�Ļ������֤�Ŵ���"
                Case 1064: GetErrInfo = "��Ժ�ĵǼ����ں���Ժ���ڴ���"
                Case 1301: GetErrInfo = "�󲡱����Ѵ��ڣ�"
                Case 1302: GetErrInfo = "�޴˴󲡱��룡"
                Case 1303: GetErrInfo = "�����ƴ󲡱����Ѵ��ڣ�"
                Case 1304: GetErrInfo = "�޴������ƴ󲡱��� , û�ö�Ӧ�������ʻ���"
                Case 1305: GetErrInfo = "��ת�ﵥ���Ѵ��ڣ�"
                Case 1306: GetErrInfo = "�޴�ת�ﵥ�ţ�"
                Case 1307: GetErrInfo = "���޶�ҽԺ�Ѵ��ڣ�"
                Case 1308: GetErrInfo = "�޴��޶�ҽԺ��"
                Case 7001, 7002, 7003, 7004, 7005
                    GetErrInfo = "���������Ӵ���ϵͳʹ�õĶ�����̬���ӿ����IC������"
                Case 7006: GetErrInfo = "ϵͳ���󣬵��붯̬���ӿ����"
                Case 7007: GetErrInfo = "д��ʱУ�鿨�������� (���ܿ�������)��"
                Case -8001: GetErrInfo = "������У�����"
                Case -8002: GetErrInfo = "ҽԺ��Ŵ���ϵͳ���ô���"
                Case -8003: GetErrInfo = "ϵͳ�汾������Ҫ���¿ͻ��ĳ���"
                Case -8004: GetErrInfo = "ϵͳ���ڴ�����Ҫ���Ŀͻ������ڣ�"
                Case 1401: GetErrInfo = "������ҽ�����㣡"
                Case 1402: GetErrInfo = "�����ﲻ���ڣ�"
                Case 1403: GetErrInfo = "�Ҵ����Ѵ����"
                Case 1404: GetErrInfo = "�α�����1����2 (ҽ�Ʊ��ղ�����)��"
                Case 1405: GetErrInfo = "����������������գ�"
                Case 1406: GetErrInfo = "��������й��˱��գ�"
                Case 1407: GetErrInfo = "��Ϊ��Ժ״̬���������������"
                Case 1408: GetErrInfo = "��Ч����"
                Case 1409: GetErrInfo = "���ʽ���ͳ���ۼƳ��ָ�ֵ��"
                Case 1410: GetErrInfo = "�Ƿ�ҽԺ , ҽԺ�����ڣ�"
                Case 1411: GetErrInfo = "������ת�ﵥ�ѱ����ã�"
                Case 1412: GetErrInfo = "�Ƿ����ڸ�ʽ Ӧ��'YYYYMMDD'"
                Case 1413: GetErrInfo = "�Ƿ�����ʱ���ʽ Ӧ��'YYYYMMDDhhmmss'��"
                Case 1414: GetErrInfo = "סԺ�Ǽǵ�ʱ��д����Ų���ȣ�"
                Case 1415: GetErrInfo = "���ҽԺ����������סԺ��"
                Case 1416: GetErrInfo = "���ҽԺ�����Թ���סԺ��"
                Case 1417: GetErrInfo = "�������˽������"
                Case 1418: GetErrInfo = "���˳��˽������"
                Case 1419: GetErrInfo = "�������˽������"
                '�ܺ�ȫ���� 2003-12-17
                '�������´������
                Case 1424: GetErrInfo = "�����ܶ���ָ�����"
                Case 1425: GetErrInfo = "���㷽ʽ����Ժ�ǼǷ�ʽ��һ�£�"
                Case 1427: GetErrInfo = "��Ժ���ڴ��ڳ�Ժ���ڻ��߳�Ժ���ڴ��ڽ������ڣ�"
                Case Else
                    GetErrInfo = "ҽ��֧�ֲ��ֳ��ִ���"
                End Select
                '�ܺ�ȫ���� 2003-12-17
                'ͬʱ���ϴ����ţ��Է�����
                GetErrInfo = "�����ţ�" & strCode & vbCr & "����������" & GetErrInfo
        Case TYPE_�ش�У԰��
                Select Case Val(strCode)
                Case 0: GetErrInfo = " �ɹ�"
                Case -1: GetErrInfo = "�򿪴���ʧ��"
                Case -2: GetErrInfo = "��д������ʧ��"
                Case -3: GetErrInfo = "��������"
                Case -4: GetErrInfo = "��ʱ����"
                Case -5: GetErrInfo = "�޿�"
                Case -6: GetErrInfo = "�û�������"
                Case -7: GetErrInfo = "��������"
                Case -8: GetErrInfo = "д������"
                Case -9: GetErrInfo = "��ֵʧ��"
                Case -10: GetErrInfo = "��ֵʧ��"
                Case -11: GetErrInfo = "����Licence����"
                Case -12: GetErrInfo = "Licence����"
                Case -13: GetErrInfo = "ϵͳ������"
                Case -14: GetErrInfo = "����"
                Case -15: GetErrInfo = "����δ����"
                Case -16: GetErrInfo = "����ͨѶ����"
                Case -17: GetErrInfo = "�����ļ�����"
                Case -18: GetErrInfo = "Ӧ�����"
                Case -19: GetErrInfo = "�������еĿ�"
                Case -20: GetErrInfo = "���ѵ���"
                Case -21: GetErrInfo = "���ݿ����ʧ��"
                Case -22: GetErrInfo = "��������ʧ��"
                Case -23: GetErrInfo = "�������"
                Case -24: GetErrInfo = "���ſ�����"
                Case -25: GetErrInfo = "�����������޶�"
                Case -100: GetErrInfo = "�޷�ʶ��Ŀ�"
                Case Else
                    GetErrInfo = "У԰��֧�ֲ��ֳ��ִ���"
                End Select
                GetErrInfo = "�����ţ�" & strCode & vbCr & "����������" & GetErrInfo
               Case Else
    End Select
End Function

Public Function OraDataOpen(cnOracle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional blnMessage As Boolean = True) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strError As String
    
    On Error Resume Next
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
    End With
    If Err <> 0 Then
        If blnMessage = True Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            Else
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
            End If
        End If
        
        Err.Clear
        OraDataOpen = False
        Exit Function
    End If
    OraDataOpen = True
End Function

Public Sub GetUserInfo()
 '���ܣ���ȡ��½�û���Ϣ
    Dim rsUser As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsUser = New ADODB.Recordset
    rsUser.CursorLocation = adUseClient
    'rsUser.Open "Select A.ID,A.����ID,A.���,A.����,A.����,B.�û���,C.���� as ���� from ��Ա�� A,�ϻ���Ա�� B,���ű� C Where A.����ID=C.ID And  B.��ԱID=A.ID AND Upper(B.�û���)=Upper(User)", gcnOracle, adOpenKeyset
    
    strSQL = "select P.ID,P.���,P.����,P.����,D.���� as ���ű���,D.���� as ��������,M.����ID,u.�û��� " & _
                " from �ϻ���Ա�� U,��Ա�� P,���ű� D,������Ա M " & _
                " Where U.��Աid = P.id And P.ID=M.��ԱID and  M.ȱʡ=1 and M.����id = D.id and U.�û���=user"
    rsUser.Open strSQL, gcnOracle, adOpenKeyset
    
    If rsUser.RecordCount <> 0 Then
        UserInfo.ID = rsUser!ID
        UserInfo.��� = rsUser!���
        UserInfo.����ID = IIf(IsNull(rsUser!����ID), 0, rsUser!����ID)
        UserInfo.���� = IIf(IsNull(rsUser!����), "", rsUser!����)
        UserInfo.���� = IIf(IsNull(rsUser!����), "", rsUser!����)
        UserInfo.���� = rsUser!��������
        UserInfo.�û��� = rsUser!�û���
        UserInfo.վ�� = rsUser!�û���
        
        'Ϊ�˲������������ظ�������һ������
        gstrUserName = UserInfo.����
    End If
End Sub

Public Function DateStr() As String
    Dim rsTmp As New ADODB.Recordset

    rsTmp.Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    DateStr = Format(rsTmp.Fields(0).Value, "yyyy-MM-dd HH:mm:ss")
End Function

Public Function TrimStr(ByVal str As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Public Function TruncZero(ByVal StrInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(StrInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(StrInput, 1, lngPos - 1)
    Else
        TruncZero = StrInput
    End If
End Function

Public Function NextNo(intBillID As Integer) As Variant
'���ܣ������ض���������µĺ���,�������£�
'   һ����Ŀ��ţ�
'   1   ����ID         ����
'   2   סԺ��         ����(ZLHIS9/10����ͬ���ݲ�֧��)
'   3   �����         ����(ZLHIS9/10����ͬ���ݲ�֧��)
'   x   �������ݺ�     �ַ�,���ݱ�Ź���˳��������,���Զ���ȱ
'   �������λȷ��ԭ��:
'       ��1990Ϊ���������������������0��9/A��Z��˳����Ϊ��ȱ���

    Dim rsCtrl As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim vntNo As Variant, strSQL As String
    Dim intYear, strYear As String
ReStart:
    Err = 0
    On Error GoTo errHand

    If intBillID = 1 Then '����ID
        With rsCtrl
            If .State = adStateOpen Then .Close
                strSQL = "Select * From ������Ʊ� Where ��Ŀ���=" & intBillID
                Call SQLTest(App.ProductName, "NextNo", strSQL)
                .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
                Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            vntNo = IIf(IsNull(!������), 0, !������)
            strSQL = "Select Nvl(Max(����ID),0)+1 as ����ID From ������Ϣ Where ����ID>=" & vntNo
            
            With rsTmp
                If .State = adStateOpen Then .Close
                Call SQLTest(App.ProductName, "NextNo", strSQL)
                .Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
                Call SQLTest
                If Not (.EOF Or .BOF) Then
                    If Not IsNull(.Fields(0).Value) Then
                        vntNo = .Fields(0).Value
                    End If
                End If
            End With
            
            On Error Resume Next
            .Update "������", IIf(vntNo - 10 > 0, vntNo - 10, 1)
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    Else
        With rsCtrl
            strSQL = "Select C.*,Sysdate as Today From ������Ʊ� C Where C.��Ŀ���=" & intBillID
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "NextNo", strSQL)
            .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
            Call SQLTest
            If .EOF Or .BOF Then
                NextNo = Null
                Exit Function
            End If
            
            intYear = Format(!Today, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            vntNo = IIf(IsNull(!������), "", !������)
            
            If IIf(IsNull(!��Ź���), 0, !��Ź���) = 1 Then
                '����˳����
                If vntNo < strYear & Format(CDate(Format(!Today, "YYYY-MM-dd")) - CDate(Format(!Today, "YYYY") & "-01-01") + 1, "000") & "0000" Then
                    vntNo = strYear & Format(CDate(Format(!Today, "YYYY-MM-dd")) - CDate(Format(!Today, "YYYY") & "-01-01") + 1, "000") & "0000"
                End If
                vntNo = Left(vntNo, 4) & Right(String(4, "0") & CStr(Val(Mid(vntNo, 5)) + 1), 4)
            Else
                '����˳����
                If Left(vntNo, 1) < strYear Then
                    vntNo = strYear & "0000000"
                End If
                vntNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + 1), 7)
            End If
            
            If Not (UCase(strYear) >= "A" And UCase(strYear) <= "Z") Or zlCommFun.ActualLen(vntNo) > 8 Then GoTo ReStart
            
            On Error Resume Next
            .Update "������", vntNo
            If Err <> 0 Then
                .CancelUpdate
                GoTo ReStart
            End If
            NextNo = vntNo
        End With
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    NextNo = Null
End Function

Public Function GetNextNO(ByVal intBillID As Integer, Optional lng���� As Long) As Variant
    'blnUse:���°汾��ʹ�ù��������е�GetNextNO��������������HIS+�汾
    
    If IsZLHIS10 Then
        #If gverControl >= 3 Then
            GetNextNO = zlDatabase.GetNextNO(intBillID, lng����)
        #Else
            GetNextNO = NextNo(intBillID)
        #End If
    Else
        GetNextNO = NextNo(intBillID)
    End If
End Function

 

Public Function Get��Ժ���(lng����ID As Long, lng��ҳID As Long, _
Optional ByVal bln����� As Boolean = True, Optional ByVal bln�������� As Boolean = False) As String
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.������Ϣ as ��Ժ���,B.���� �������� " & _
             " From ������ A,��������Ŀ¼ B " & _
             " Where A.����ID=[1] And A.����ID=B.ID(+) And A.��ҳID=[2] And A.�������=2"
    Set rsInNote = zlDatabase.OpenSQLRecord(strTmp, "Get��Ժ���", lng����ID, lng��ҳID)
    
    If Not rsInNote.EOF Then
        Get��Ժ��� = IIf(IsNull(rsInNote!��Ժ���), "", rsInNote!��Ժ���)
    End If
    If Not bln����� Then
        Get��Ժ��� = Trim(Get��Ժ���)
        If Get��Ժ��� = "" Then Get��Ժ��� = "��"
    End If
    If bln�������� Then
        If Not rsInNote.EOF Then
            Get��Ժ��� = Get��Ժ��� & "|" & Nvl(rsInNote!��������)
        Else
            Get��Ժ��� = Get��Ժ��� & "|"
        End If
    End If
End Function

Public Function BuildPatiInfo(ByVal bytType As Byte, ByVal strInfo As String, ByVal lng����ID As Long, ByVal intinsure As Integer) As Long
'���ܣ����������ʻ���Ϣ
'������bytType=0-����,1-סԺ
'      strInfo='0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
'      8����;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(1,2,3);15����֤��;16�����;17�Ҷȼ�
'      18�ʻ������ۼ�;19�ʻ�֧���ۼ�;20����ͳ���ۼ�;21ͳ�ﱨ���ۼ�;22סԺ�����ۼ�;23�������
'      24��������;25�����ۼ�;26����ͳ���޶�
'���أ�����ID
    Const MAX_BOUND = 26 'Ҫ�������Ϣ����
    
    Dim rsPati As New ADODB.Recordset, str��λ���� As String, lng���� As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, curDate As Date
    Dim lng���� As Long, array��Ϣ As Variant
    Dim lngTemp As Long
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    If Len(Trim(strInfo)) <> 0 Then
        curDate = zlDatabase.Currentdate
        
        '200308z012:��֤�������Ϣ������
        If UBound(Split(strInfo, ";")) < MAX_BOUND Then
            strInfo = strInfo & String(MAX_BOUND - UBound(Split(strInfo, ";")), ";")
        End If
        array��Ϣ = Split(strInfo, ";")
        
        '�ӵ�7��������ȡ����λ����
        '���з��صĵ�λ��ʽ�������ţ��������в����ǵ�λ���룬��������ע����Ϣ������ɱ���������Բ�����
        If intinsure <> TYPE_�ϳ����� Then
            If array��Ϣ(7) Like "*(*" Then
                str��λ���� = Split(array��Ϣ(7), "(")(UBound(Split(array��Ϣ(7), "(")))
                str��λ���� = Mid(str��λ����, 1, Len(str��λ����) - 1)
            End If
        End If
        
        'ȡ����
        If IsDate(array��Ϣ(5)) Then
            lng���� = Int(curDate - CDate(array��Ϣ(5))) / 365
        End If
        
        lng���� = Val(array��Ϣ(8))
        
        '�ṩ�˲�����ݰ󶨵Ĺ��ܣ���˲�����Ҫ�ϲ�
'        If lng����ID > 0 Then
'            '�ò����Ѿ�����
'            gstrSQL = "Select nvl(����ID,0) ����ID from �����ʻ� where ҽ����='" & CStr(array��Ϣ(1)) & "' and ����=" & lng���� & " and ����=" & intInsure
'            Call OpenRecordset(rsTemp, "�����ʻ�")
'            If rsTemp.EOF = False Then
'                If rsTemp("����ID") <> lng����ID Then
'                    '������(2006-01-16):����ҽ��֧�ֲ���Ǽ�ʱ�Զ��Ǽ�
'                    If intInsure = TYPE_�ɶ��� Or intInsure = TYPE_�¶� Or intInsure = type_�ɶ����� Or intInsure = TYPE_��ɽ Or intInsure = TYPE_��Ԫ���� Or intInsure = TYPE_�ɶ����� Or intInsure = TYPE_�ϳ����� Then
'                        If MsgBox("�Ѿ�������ͬҽ���ŵ�����һλ���ˣ�����Ҫ������λ���˺ϲ���", vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then
'                            gcnOracle.RollbackTrans
'                            Exit Function
'                        End If
'                        '�����������˽��кϲ�
'                        lngTemp = MergePatient(lng����ID, rsTemp!����ID)
'                        If lngTemp = 0 Then
'                            gcnOracle.RollbackTrans
'                            Exit Function
'                        End If
'                        lng����ID = lngTemp
'                    Else
'                        MsgBox "�Ѿ�������ͬҽ���ŵ�����һλ���ˣ������ڲ��˹����н�����λ���˺ϲ�", vbInformation, gstrSysName
'                        gcnOracle.RollbackTrans
'                        Exit Function
'                    End If
'                End If
'            End If
'        End If
        
        '�ʻ�Ψһ������,����,ҽ����
        #If gverControl < 6 Then
            strSQL = "Select A.*,B.ҽ���� From ������Ϣ A," & _
                "   (Select * From �����ʻ�" & _
                "   Where ����=[1] And ҽ����=[2] And Nvl(����,0)=[3]) B" & _
                " Where " & IIf(lng����ID = 0, "A.����ID=B.����ID", "A.����ID=B.����ID(+) and A.����ID=[4]") '���ܲ���ID�Ѿ�ȷ��
        #Else
            strSQL = "Select A.����id, A.�����, A.סԺ��, A.���￨��, A.����֤��, A.�ѱ�, A.ҽ�Ƹ��ʽ, A.����, A.�Ա�, A.����, A.��������, A.�����ص�, A.���֤��, A.����֤��, A.���, A.ְҵ, A.����, A.����, A.����, A.ѧ��, A.����״��, A.��ͥ��ַ," & vbNewLine & _
                "      A.��ͥ�绰, A.��ͥ��ַ�ʱ� As �����ʱ�, A.�໤��, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ, A.��ϵ�˵绰, A.��ͬ��λid, A.������λ, A.��λ�绰, A.��λ�ʱ�, A.��λ������, A.��λ�ʺ�, A.������, A.������, A.��������, A.����ʱ��, A.����״̬," & vbNewLine & _
                "      A.��������, A.סԺ����, A.��ǰ����id, A.��ǰ����id, A.��ǰ����, A.��Ժʱ��, A.��Ժʱ��, A.��Ժ, A.Ic����, A.������, A.ҽ����, A.����, A.��ѯ����, A.�Ǽ�ʱ��, A.ͣ��ʱ��, A.����," & vbNewLine & _
                "      B.ҽ���� From ������Ϣ A," & _
                "   (Select * From �����ʻ�" & _
                "   Where ����=[1] And ҽ����=[2] And Nvl(����,0)=[3]) B" & _
                " Where " & IIf(lng����ID = 0, "A.����ID=B.����ID", "A.����ID=B.����ID(+) and A.����ID=[4]") '���ܲ���ID�Ѿ�ȷ��
        #End If
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", intinsure, CStr(array��Ϣ(1)), lng����, lng����ID)
        If rsPati.EOF Then
            '�ޱ����ʻ�����Ϊû�в�����Ϣ
            If lng����ID = 0 Then lng����ID = GetNextNO(1)
            strSQL = "zl_������Ϣ_Insert(" & lng����ID & ",NULL,NULL,'������ҽ�Ʊ���'," & _
                "'" & array��Ϣ(3) & "','" & array��Ϣ(4) & "'," & IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
                "To_Date('" & Format(array��Ϣ(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                "NULL,'" & array��Ϣ(6) & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," & _
                "NULL,NULL,NULL,NULL,NULL,NULL,'" & array��Ϣ(7) & "',NULL,NULL,NULL," & _
                "NULL,NULL,NULL," & intinsure & "," & _
                "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            Call zlDatabase.ExecuteProcedure(strSQL, "�������˵���")
        Else
            '���ò����Ƿ��Ѿ�ͣ��
            If Not IsNull(rsPati!ͣ��ʱ��) Then
                gcnOracle.RollbackTrans
                MsgBox "�ò��˵���Ϣ�Ѿ�ͣ�á�", vbInformation, gstrSysName
                Exit Function
            End If
            
            '�в�����Ϣ�ͱ����ʻ���Ϣ
            If rsPati("����") <> array��Ϣ(3) Then
                If MsgBox("����ԭ�еǼǵ������� " & rsPati("����") & " ����ˢ���õ������� " & array��Ϣ(3) & " ������" & vbCrLf & _
                          "��������²���ԭ�еĵǼ���Ϣ���Ƿ�ȷ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
            End If
            
            '2005-08-13 �ܺ�ȫ
            '��������λ��ԭ����ʽ��������λ����ΪID
            If lng����ID = 0 Then lng����ID = rsPati!����ID
            If intinsure = TYPE_�ɶ��� Or intinsure = TYPE_�¶� Or intinsure = type_�ɶ����� Or intinsure = TYPE_��ɽ Or intinsure = TYPE_��Ԫ���� Or intinsure = TYPE_�ɶ����� Or intinsure = TYPE_�ϳ����� Then
                strSQL = "zl_������Ϣ_Update(" & _
                    lng����ID & "," & IIf(IsNull(rsPati!�����), "NULL", rsPati!�����) & "," & _
                    IIf(IsNull(rsPati!סԺ��), "NULL", rsPati!סԺ��) & ",'" & IIf(IsNull(rsPati!�ѱ�), "", rsPati!�ѱ�) & "'," & _
                    "'" & IIf(IsNull(rsPati!ҽ�Ƹ��ʽ), "", rsPati!ҽ�Ƹ��ʽ) & "'," & _
                    "'" & array��Ϣ(3) & "','" & array��Ϣ(4) & "'," & IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
                    "     To_Date('" & Format(array��Ϣ(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                    "'" & IIf(IsNull(rsPati!�����ص�), "", rsPati!�����ص�) & "','" & array��Ϣ(6) & "'," & _
                    "'" & IIf(IsNull(rsPati!���), "", rsPati!���) & "','" & IIf(IsNull(rsPati!ְҵ), "", rsPati!ְҵ) & "'," & _
                    "'" & IIf(IsNull(rsPati!����), "", rsPati!����) & "','" & IIf(IsNull(rsPati!����), "", rsPati!����) & "'," & _
                    "'" & IIf(IsNull(rsPati!ѧ��), "", rsPati!ѧ��) & "','" & IIf(IsNull(rsPati!����״��), "", rsPati!����״��) & "'," & _
                    "'" & IIf(IsNull(rsPati!��ͥ��ַ), "", rsPati!��ͥ��ַ) & "','" & IIf(IsNull(rsPati!��ͥ�绰), "", rsPati!��ͥ�绰) & "'," & _
                    "'" & IIf(IsNull(rsPati!�����ʱ�), "", rsPati!�����ʱ�) & "','" & IIf(IsNull(rsPati!��ϵ������), "", rsPati!��ϵ������) & "'," & _
                    "'" & IIf(IsNull(rsPati!��ϵ�˹�ϵ), "", rsPati!��ϵ�˹�ϵ) & "','" & IIf(IsNull(rsPati!��ϵ�˵�ַ), "", rsPati!��ϵ�˵�ַ) & "'," & _
                    "'" & IIf(IsNull(rsPati!��ϵ�˵绰), "", rsPati!��ϵ�˵绰) & "'," & IIf(IsNull(rsPati!��ͬ��λID), "NULL", rsPati!��ͬ��λID) & "," & _
                    "'" & array��Ϣ(7) & "','" & IIf(IsNull(rsPati!��λ�绰), "", rsPati!��λ�绰) & "'," & _
                    "'" & IIf(IsNull(rsPati!��λ�ʱ�), "", rsPati!��λ�ʱ�) & "','" & IIf(IsNull(rsPati!��λ������), "", rsPati!��λ������) & "'," & _
                    "'" & IIf(IsNull(rsPati!��λ�ʺ�), "", rsPati!��λ�ʺ�) & "','" & IIf(IsNull(rsPati!������), "", rsPati!������) & "'," & _
                    " " & IIf(IsNull(rsPati!������), "NULL", rsPati!������) & "," & intinsure & ")"
            ElseIf intinsure = TYPE_�������� Then
                strSQL = "zl_������Ϣ_Update(" & _
                    lng����ID & "," & IIf(IsNull(rsPati!�����), "NULL", rsPati!�����) & "," & _
                    IIf(IsNull(rsPati!סԺ��), "NULL", "NULL") & ",'" & IIf(IsNull(rsPati!�ѱ�), "", rsPati!�ѱ�) & "'," & _
                    "'" & IIf(IsNull(rsPati!ҽ�Ƹ��ʽ), "", rsPati!ҽ�Ƹ��ʽ) & "'," & _
                    "'" & array��Ϣ(3) & "','" & array��Ϣ(4) & "'," & IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
                    "To_Date('" & Format(array��Ϣ(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                    "'" & IIf(IsNull(rsPati!�����ص�), "", rsPati!�����ص�) & "','" & array��Ϣ(6) & "'," & _
                    "'" & IIf(IsNull(rsPati!���), "", rsPati!���) & "','" & IIf(IsNull(rsPati!ְҵ), "", rsPati!ְҵ) & "'," & _
                    "'" & IIf(IsNull(rsPati!����), "", rsPati!����) & "','" & IIf(IsNull(rsPati!����), "", rsPati!����) & "'," & _
                    "'" & IIf(IsNull(rsPati!ѧ��), "", rsPati!ѧ��) & "','" & IIf(IsNull(rsPati!����״��), "", rsPati!����״��) & "'," & _
                    "'" & IIf(IsNull(rsPati!��ͥ��ַ), "", rsPati!��ͥ��ַ) & "','" & IIf(IsNull(rsPati!��ͥ�绰), "", rsPati!��ͥ�绰) & "'," & _
                    "'" & IIf(IsNull(rsPati!�����ʱ�), "", rsPati!�����ʱ�) & "','" & IIf(IsNull(rsPati!��ϵ������), "", rsPati!��ϵ������) & "'," & _
                    "'" & IIf(IsNull(rsPati!��ϵ�˹�ϵ), "", rsPati!��ϵ�˹�ϵ) & "','" & IIf(IsNull(rsPati!��ϵ�˵�ַ), "", rsPati!��ϵ�˵�ַ) & "'," & _
                    "'" & IIf(IsNull(rsPati!��ϵ�˵绰), "", rsPati!��ϵ�˵绰) & "'," & IIf(IsNull(rsPati!��ͬ��λID), "NULL", rsPati!��ͬ��λID) & "," & _
                    " " & IIf(IsNull(rsPati!������λ), "NULL", "'" & rsPati!������λ & "'") & ",'" & IIf(IsNull(rsPati!��λ�绰), "", rsPati!��λ�绰) & "'," & _
                    "'" & IIf(IsNull(rsPati!��λ�ʱ�), "", rsPati!��λ�ʱ�) & "','" & IIf(IsNull(rsPati!��λ������), "", rsPati!��λ������) & "'," & _
                    "'" & IIf(IsNull(rsPati!��λ�ʺ�), "", rsPati!��λ�ʺ�) & "','" & IIf(IsNull(rsPati!������), "", rsPati!������) & "'," & _
                    " " & IIf(IsNull(rsPati!������), "NULL", rsPati!������) & "," & intinsure & ")"
            Else
                strSQL = "zl_������Ϣ_Update(" & _
                    lng����ID & "," & IIf(IsNull(rsPati!�����), "NULL", rsPati!�����) & "," & _
                    IIf(IsNull(rsPati!סԺ��), "NULL", rsPati!סԺ��) & ",'" & IIf(IsNull(rsPati!�ѱ�), "", rsPati!�ѱ�) & "'," & _
                    "'" & IIf(IsNull(rsPati!ҽ�Ƹ��ʽ), "", rsPati!ҽ�Ƹ��ʽ) & "'," & _
                    "'" & array��Ϣ(3) & "','" & array��Ϣ(4) & "'," & IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
                    "To_Date('" & Format(array��Ϣ(5), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                    "'" & IIf(IsNull(rsPati!�����ص�), "", rsPati!�����ص�) & "','" & array��Ϣ(6) & "'," & _
                    "'" & IIf(IsNull(rsPati!���), "", rsPati!���) & "','" & IIf(IsNull(rsPati!ְҵ), "", rsPati!ְҵ) & "'," & _
                    "'" & IIf(IsNull(rsPati!����), "", rsPati!����) & "','" & IIf(IsNull(rsPati!����), "", rsPati!����) & "'," & _
                    "'" & IIf(IsNull(rsPati!ѧ��), "", rsPati!ѧ��) & "','" & IIf(IsNull(rsPati!����״��), "", rsPati!����״��) & "'," & _
                    "'" & IIf(IsNull(rsPati!��ͥ��ַ), "", rsPati!��ͥ��ַ) & "','" & IIf(IsNull(rsPati!��ͥ�绰), "", rsPati!��ͥ�绰) & "'," & _
                    "'" & IIf(IsNull(rsPati!�����ʱ�), "", rsPati!�����ʱ�) & "','" & IIf(IsNull(rsPati!��ϵ������), "", rsPati!��ϵ������) & "'," & _
                    "'" & IIf(IsNull(rsPati!��ϵ�˹�ϵ), "", rsPati!��ϵ�˹�ϵ) & "','" & IIf(IsNull(rsPati!��ϵ�˵�ַ), "", rsPati!��ϵ�˵�ַ) & "'," & _
                    "'" & IIf(IsNull(rsPati!��ϵ�˵绰), "", rsPati!��ϵ�˵绰) & "'," & IIf(IsNull(rsPati!��ͬ��λID), "NULL", rsPati!��ͬ��λID) & "," & _
                    " " & IIf(IsNull(rsPati!������λ), "NULL", "'" & rsPati!������λ & "'") & ",'" & IIf(IsNull(rsPati!��λ�绰), "", rsPati!��λ�绰) & "'," & _
                    "'" & IIf(IsNull(rsPati!��λ�ʱ�), "", rsPati!��λ�ʱ�) & "','" & IIf(IsNull(rsPati!��λ������), "", rsPati!��λ������) & "'," & _
                    "'" & IIf(IsNull(rsPati!��λ�ʺ�), "", rsPati!��λ�ʺ�) & "','" & IIf(IsNull(rsPati!������), "", rsPati!������) & "'," & _
                    " " & IIf(IsNull(rsPati!������), "NULL", rsPati!������) & "," & intinsure & ")"
                Call SQLTest(App.ProductName, "ҽ���ӿ�", strSQL)
            End If
            Call zlDatabase.ExecuteProcedure(strSQL, "���²��˵���")
        End If
        
        '�������±����ʻ���Ϣ(�Զ�)
        strSQL = "zl_�����ʻ�_insert(" & lng����ID & "," & intinsure & "," & _
            lng���� & "," & _
            "'" & IIf(array��Ϣ(0) = "-1", array��Ϣ(1), array��Ϣ(0)) & "'," & _
            "'" & array��Ϣ(1) & "'," & _
            "'" & array��Ϣ(2) & "'," & _
            "'" & array��Ϣ(9) & "'," & _
            "'" & array��Ϣ(15) & "'," & _
            "'" & array��Ϣ(10) & "'," & _
            "'" & str��λ���� & "'," & _
            Val(array��Ϣ(11)) & "," & _
            Val(array��Ϣ(12)) & "," & _
            IIf(Val(array��Ϣ(13)) = 0, "NULL", Val(array��Ϣ(13))) & "," & _
            IIf(Val(array��Ϣ(14)) = 0, 1, Val(array��Ϣ(14))) & "," & _
            IIf(Val(array��Ϣ(16)) = 0, lng����, Val(array��Ϣ(16))) & "," & _
            "'" & array��Ϣ(17) & "'," & _
            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSQL, "���������ʻ�")
        
        '���������ʻ������Ϣ(�Զ�)
        '200308z012:�ɶ�:����"24��������=zyjs,25�����ۼ�=tcbxbl,26����ͳ���޶�=zyxe"
        strSQL = "zl_�ʻ������Ϣ_Insert(" & lng����ID & "," & intinsure & "," & Year(curDate) & "," & _
            Val(array��Ϣ(18)) & "," & Val(array��Ϣ(19)) & "," & _
            Val(array��Ϣ(20)) & "," & Val(array��Ϣ(21)) & "," & _
            Val(array��Ϣ(22)) & "," & Val(array��Ϣ(24)) & "," & Val(array��Ϣ(25)) & "," & Val(array��Ϣ(26)) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "���������Ϣ")
    End If
    
    gcnOracle.CommitTrans
    BuildPatiInfo = lng����ID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Public Function GetTextFromCombo(cmbTemp As ComboBox, ByVal blnAfter As Boolean, Optional strSplit As String = ".") As String
'������cmbTemp  ׼����ȡ���ݵ�ComboBox�ؼ�
'      blnAfter ��ʾ��.֮ǰ��֮��ȡֵ
    Dim lngPos As Long
    
    lngPos = InStr(cmbTemp.Text, strSplit)
    If lngPos = 0 Then
        'ֱ�ӷ��������ַ���
        GetTextFromCombo = "'" & cmbTemp.Text & "'"
    Else
        If blnAfter = False Then
            'Բ��֮ǰ
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, 1, lngPos - 1) & "'"
        Else
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, lngPos + 1) & "'"
        End If
    End If
End Function

Public Sub SetComboByText(cmbTemp As ComboBox, ByVal strText As String, ByVal blnAfter As Boolean, Optional strSplit As String = ".")
'������cmbTemp  ׼�����õ�ComboBox�ؼ�
'      blnAfter ��ʾ��.֮ǰ��֮��ȡֵ
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cmbTemp.ListCount - 1
        strTemp = cmbTemp.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            'ֱ�ӷ��������ַ���
            If strText = strTemp Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                'Բ��֮ǰ
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '�Ѿ��ҵ�
        cmbTemp.ListIndex = lngCount
    Else
        cmbTemp.ListIndex = -1
        If blnAfter = True Then
            cmbTemp.AddItem strText
        End If
    End If
End Sub

Public Function MidUni(ByVal strTemp As String, ByVal Start As Long, ByVal Length As Long) As String
'���ܣ������ݿ����õ��ַ������Ӽ���Ҳ���Ǻ��ְ������ַ��㣬����ĸ����һ��
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    'ȥ�����ܳ��ֵİ���ַ�
    MidUni = Replace(MidUni, Chr(0), "")
End Function

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
'���ܣ����ı���Varchar2�ĳ��ȼ��㷽�����нض�
    Dim strText As String
    
    strText = IIf(IsNull(varText), "", varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    'ȥ�����ܳ��ֵİ���ַ�
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function

Public Function GetComputer(frmParant As Form, Optional ByVal strCaption As String = "ѡ������") As String
'���ܣ����ؼ������
   Dim BI As BrowseInfo
   Dim pidl As Long
   Dim sPath As String
   Dim pos As Integer
   
  'obtain the pidl to the special folder 'network'
   If SHGetSpecialFolderLocation(frmParant.hwnd, CSIDL_NETWORK, pidl) = 0 Then
     'fill in the required members, limiting the
     'Browse to the network by specifying the
     'returned pidl as pidlRoot
      With BI
         .hwndOwner = frmParant.hwnd
         .pIDLRoot = pidl
         .pszDisplayName = Space$(MAX_PATH)
         .lpszTitle = lstrcat(strCaption, "")
         .ulFlags = BIF_BROWSEFORCOMPUTER
      End With
         
     'show the browse dialog. We don't need
     'a pidl, so it can be used in the If..then directly.
      If SHBrowseForFolder(BI) <> 0 Then
               
         'a server was selected. Although a valid pidl
         'is returned, SHGetPathFromIDList only return
         'paths to valid file system objects, of which
         'a networked machine is not. However, the
         'BROWSEINFO displayname member does contain
         'the selected item, which we return
          GetComputer = TrimStr(BI.pszDisplayName)
            
      End If  'If SHBrowseForFolder
      
      Call CoTaskMemFree(pidl)
               
   End If  'If SHGetSpecialFolderLocation
   
End Function

Public Sub CenterTableCaption(mshTemp As Object)
'���ܣ����ñ�����ͷ���ж���
    With mshTemp
        .COL = 0
        .Row = .FixedRows - 1
        .ColSel = .Cols - 1
        .RowSel = .Row
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = .FixedRows: .COL = .FixedCols
    End With
End Sub

Public Function GetסԺ����(lng����ID As Long) As Integer
'���ܣ���ȡָ�����˱����סԺ����
'˵��������סԺ��������궼����һ��סԺ��
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select Count(*) as ���� From ������ҳ" & _
        " Where Nvl(��ҳID,0)<>0 And Nvl(��Ժ����,Sysdate)=To_Date(To_Char(Sysdate,'YYYY')||'-01-01','YYYY-MM-DD') And ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ�����˱����סԺ����", lng����ID)
    
    If Not rsTmp.EOF Then GetסԺ���� = IIf(IsNull(rsTmp!����), 0, rsTmp!����)
End Function

Public Function Get�ʻ���Ϣ(ByVal intinsure As Integer, ByVal lng����ID As Long, ByVal str��� As String, intסԺ�����ۼ� As Integer, _
    cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, _
    curͳ�ﱨ���ۼ� As Currency, Optional cur�������� As Currency, Optional cur�����ۼ� As Currency, _
    Optional cur����ͳ���޶� As Currency) As Boolean
'���ܣ��õ��ʻ������Ϣ
'200308z012:�����������ز���
    Dim rsTemp As New ADODB.Recordset
    
    cur�ʻ������ۼ� = 0
    cur�ʻ�֧���ۼ� = 0
    cur����ͳ���ۼ� = 0
    curͳ�ﱨ���ۼ� = 0
    intסԺ�����ۼ� = 0
    cur�������� = 0
    cur�����ۼ� = 0
    cur����ͳ���޶� = 0
    
    '�ʻ������Ϣ
    gstrSQL = "Select * From �ʻ������Ϣ Where ����ID=[1] And ����=[2] And ���=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ������Ϣ", lng����ID, intinsure, str���)
    If rsTemp.EOF = False Then
        cur�ʻ������ۼ� = IIf(IsNull(rsTemp("�ʻ������ۼ�")), 0, rsTemp("�ʻ������ۼ�"))
        cur�ʻ�֧���ۼ� = IIf(IsNull(rsTemp("�ʻ�֧���ۼ�")), 0, rsTemp("�ʻ�֧���ۼ�"))
        cur����ͳ���ۼ� = IIf(IsNull(rsTemp("����ͳ���ۼ�")), 0, rsTemp("����ͳ���ۼ�"))
        curͳ�ﱨ���ۼ� = IIf(IsNull(rsTemp("ͳ�ﱨ���ۼ�")), 0, rsTemp("ͳ�ﱨ���ۼ�"))
        intסԺ�����ۼ� = IIf(IsNull(rsTemp("סԺ�����ۼ�")), 0, rsTemp("סԺ�����ۼ�"))
        cur�������� = IIf(IsNull(rsTemp("��������")), 0, rsTemp("��������"))
        cur�����ۼ� = IIf(IsNull(rsTemp("�����ۼ�")), 0, rsTemp("�����ۼ�"))
        cur����ͳ���޶� = IIf(IsNull(rsTemp("����ͳ���޶�")), 0, rsTemp("����ͳ���޶�"))
    End If

End Function
Public Function Get������Ϣ(ByVal intinsure As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str��� As String, Optional str������Ϣ As String) As Boolean
'���ܣ�������ⲡ�˱�־,�ɶ��ѽ�����ʹ��
    Dim rsTemp As New ADODB.Recordset
    Dim str��Ժ��� As String
    
    str������Ϣ = "0"
    '���ڿ���Ƚ���Ĳ��˱���ȡ��Ժ����������ȵ���Ϣ
    gstrSQL = "Select to_char(��Ժ����,'YYYY') as ��Ժ��� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ڿ���Ƚ���Ĳ��˱���ȡ��Ժ����������ȵ���Ϣ", lng����ID, lng��ҳID)
    str��Ժ��� = rsTemp("��Ժ���")
    If str��Ժ��� <> str��� Then str��� = str��Ժ���
    
    '�ʻ������Ϣ
    gstrSQL = "Select * From �ʻ������Ϣ Where ����ID=[1] And ����=[2] And ���=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ������Ϣ", lng����ID, intinsure, str���)
    If rsTemp.EOF = False Then
        str������Ϣ = IIf(IsNull(rsTemp("������Ϣ")), "0", rsTemp("������Ϣ"))
    End If

End Function

Public Function �����������(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, ByVal intinsure As Integer) As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim rs�㷨 As New ADODB.Recordset
    Dim clsҽ�� As New clsInsure
    Dim rs������� As New ADODB.Recordset
    Dim dblȫ�Է� As Currency, dbl�����Ը� As Currency, dbl����ͳ�� As Currency, dblTemp As Double
    Dim dbl����� As Double
    Dim dbl�����ʻ� As Double
    Dim lng����ID As Long
    Dim rs��׼��Ŀ As New ADODB.Recordset
    Dim dblTemp1 As Double
    

    If rs��ϸ.RecordCount > 0 Then
        rs��ϸ.MoveFirst
        lng����ID = rs��ϸ("����ID")
    End If
    
    gstrSQL = "select A.�շ�ϸĿID from ������׼��Ŀ A,�����ʻ� B " & _
            "where A.����ID=B.����ID and B.����ID=[1] and ����=[2]"
    Set rs��׼��Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID, intinsure)
    
    '2����ͳ��֧����Ŀ�ϼƷ�����������
    '2.1����ʼ����¼��
    Set rs������� = New ADODB.Recordset
    With rs�������
        If .State = adStateOpen Then .Close
        .Fields.Append "���մ���ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 8, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ͳ����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    Err = 0
    On Error Resume Next
    '�������������Ļ��ܶ�
    Do Until rs��ϸ.EOF
        rs��׼��Ŀ.Filter = "�շ�ϸĿID = " & rs��ϸ("�շ�ϸĿID")
    
        If rs��ϸ("�Ƿ�ҽ��") = 1 Or rs��׼��Ŀ.EOF = False Then
            '�������׼��Ŀ��ǿ�н���ͳ��
            If rs�������.RecordCount = 0 Then
                rs�������.AddNew
                rs�������("���մ���ID") = rs��ϸ("����֧������ID")
                rs�������("����") = rs��ϸ("����")
                rs�������("���") = rs��ϸ("ʵ�ս��")
            Else
                rs�������.MoveFirst
                rs�������.Find "���մ���ID=" & rs��ϸ("����֧������ID")
                If rs�������.EOF Then
                    rs�������.AddNew
                    rs�������("���մ���ID") = rs��ϸ("����֧������ID")
                    rs�������("����") = rs��ϸ("����")
                    rs�������("���") = rs��ϸ("ʵ�ս��")
                Else
                    rs�������("����") = rs�������("����") + rs��ϸ("����")
                    rs�������("���") = rs�������("���") + rs��ϸ("ʵ�ս��")
                End If
            End If
            rs�������.Update
        Else
            dblȫ�Է� = dblȫ�Է� + rs��ϸ("ʵ�ս��")
        End If
        dblTemp = dblTemp + rs��ϸ("ʵ�ս��")
        rs��ϸ.MoveNext
    Loop
    g��������.�������ý�� = dblTemp
    
    '2.2���������ͳ����
    gstrSQL = "select ID,�㷨,ͳ��ȶ�,��׼����,��׼����,�Ƿ�ҽ�� FROM ����֧������ where ����=[1]"
    Set rs�㷨 = zlDatabase.OpenSQLRecord(gstrSQL, "�������ͳ����", intinsure)
    
    dblTemp = 0
    If rs�������.RecordCount > 0 Then rs�������.MoveFirst
    g��������.�Żݽ�� = 0
    Do Until rs�������.EOF
        rs�㷨.Filter = "ID=" & rs�������("���մ���ID")
        If rs�㷨.RecordCount > 0 Then
            If rs�㷨("�Ƿ�ҽ��") = 1 Then
                '�㷨:1-�ܶ������Ŀ��2-סԺ�պ˶���Ŀ;3-���õ��μ��㷨
                Select Case rs�㷨("�㷨")
                Case 1          '1-�ܶ������Ŀ
                    If rs�㷨("ͳ��ȶ�") = 0 Then
                        dblȫ�Է� = dblȫ�Է� + rs�������("���")
                    Else
                        dblTemp = dblTemp + rs�������("���") * rs�㷨("ͳ��ȶ�") / 100
                    End If
                Case 2      '2-סԺ�պ˶���Ŀ
                    If Val(rs�������("����")) > Val(rs�㷨("��׼����")) Then
                        '���סԺ�ճ�����׼��������ô�������� ��׼����*��׼���� +  (����-��׼����)*ͳ��ȶ�
                        '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                        dbl����� = rs�㷨("��׼����") * rs�㷨("��׼����") + _
                            (rs�������("����") - IIf(rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0, 0, rs�㷨("��׼����"))) * rs�㷨("ͳ��ȶ�")
                    Else
                        '���סԺ�յ�����׼��������ô�������� ����*��׼���� ���� ����*ͳ��ȶ�
                        '����׼�������׼������һ��Ϊ0ʱ�����൱�ڲ�Ҫ��׼����
                        If rs�㷨("��׼����") = 0 Or rs�㷨("��׼����") = 0 Then
                            dbl����� = rs�������("����") * rs�㷨("ͳ��ȶ�")
                        Else
                            dbl����� = rs�������("����") * rs�㷨("��׼����")
                        End If
                    End If
                    
                    '�ܽ��������С����ȡȫ��������ֻ�����
                    dblTemp = dblTemp + IIf(rs�������("���") < dbl�����, rs�������("���"), dbl�����)
                    
                    If rs�������("���") > dbl����� Then
                        'ȫ������ȫ�Է�
                        dblȫ�Է� = dblȫ�Է� + rs�������("���") - dbl�����
                    End If
                Case Else   '3-���õ��μ��㷨
                    If Nvl(rs�������!���, 0) = 0 Then
                    Else
                        dblTemp1 = ��ȡ���õ��ζ�_����(Nvl(rs�������!���մ���id, 0), Nvl(rs�������!���, 0))
                        dblTemp = dblTemp + dblTemp1
                        g��������.�Żݽ�� = g��������.�Żݽ�� + (Nvl(rs�������!���, 0) - dblTemp1)
                    End If
                End Select
            Else
                dblȫ�Է� = dblȫ�Է� + rs�������("���")
            End If
        Else
            dblȫ�Է� = dblȫ�Է� + rs�������("���")
        End If
        rs�������.MoveNext
    Loop
    
    g��������.����ͳ���� = dblTemp
    g��������.ȫ�Էѽ�� = dblȫ�Է�
    g��������.�����Ը���� = g��������.�������ý�� - dblȫ�Է� - dblTemp - g��������.�Żݽ��
   '20040617���˺�����
    '
    '
    '    Do Until rs��ϸ.EOF
    '        rs��׼��Ŀ.Filter = "�շ�ϸĿID = " & rs��ϸ("�շ�ϸĿID")
    '
    '        If rs��ϸ("�Ƿ�ҽ��") = 1 Or rs��׼��Ŀ.EOF = False Then
    '            '�������׼��Ŀ��ǿ�н���ͳ��
    '            dbl����ͳ�� = dbl����ͳ�� + rs��ϸ("ͳ����")
    '            dbl�����Ը� = dbl�����Ը� + rs��ϸ("ʵ�ս��") - rs��ϸ("ͳ����")
    '        Else
    '            dblȫ�Է� = dblȫ�Է� + rs��ϸ("ʵ�ս��")
    '        End If
    '
    '        rs��ϸ.MoveNext
    '    Loop
    
    If clsҽ��.GetCapability(support�շ��ʻ�ȫ�Է�, 0, intinsure) = True Then
        dbl�����ʻ� = dbl�����ʻ� + dblȫ�Է�
    End If
    
    If Isȫ��ͳ��(lng����ID, intinsure) = True Then
        '�����Ը�Ҳ����ҽ������֧��
        If g��������.�Żݽ�� = 0 Then
            str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";0|ҽ������;" & g��������.����ͳ���� + g��������.�����Ը���� & ";0"
        Else
            str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";0|ҽ������;" & g��������.����ͳ���� + g��������.�����Ը���� & ";0|�Żݽ��;" & g��������.�Żݽ�� & ";0"
        End If
        g��������.ͳ�ﱨ����� = g��������.����ͳ���� + g��������.�����Ը����
    Else
        If clsҽ��.GetCapability(support�շ��ʻ������Ը�, 0, intinsure) = True Then
            dbl�����ʻ� = dbl�����ʻ� + g��������.�����Ը����
        End If
        If g��������.�Żݽ�� = 0 Then
            str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";0|ҽ������;" & g��������.����ͳ���� & ";0"
        Else
            str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";0|ҽ������;" & g��������.����ͳ���� & ";0|�Żݽ��;" & g��������.�Żݽ�� & ";0"
        End If
        g��������.ͳ�ﱨ����� = g��������.����ͳ����
    End If
    ����������� = True
End Function

Public Function Isȫ��ͳ��(ByVal ����ID As Long, ByVal intinsure As Integer) As Boolean
'���ܣ��ж��Ƿ�ȫ��ͳ�ﲡ��(ע�⣺���Ĳ���ID���ܷ�ҽ�����˵�)
    Dim rsTemp As New ADODB.Recordset
    
    If intinsure = TYPE_�Թ��� Then
        '�����Թ�ҽ����ֻҪ������������Ա���Ǿ���ȫ��ͳ��
        gstrSQL = "select ��ְ from �����ʻ� where ����ID=[1] and ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ȫ��ͳ�ﲡ��", ����ID, TYPE_�Թ���)
        If rsTemp.EOF = False Then
            Isȫ��ͳ�� = IIf(rsTemp("��ְ") = 3, True, False)
        End If
    Else
        gstrSQL = _
            "Select Nvl(B.ȫ��ͳ��,0) as ȫ��ͳ��" & _
            " From �����ʻ� A,��������� B" & _
            " Where A.���� = B.���� And Nvl(A.����, 0) = Nvl(B.����, 0)" & _
            " And Nvl(A.��ְ,0)=Nvl(B.��ְ,0)" & _
            " And B.����<=Nvl(A.�����,0) And (A.�����<=B.���� Or B.����=0)" & _
            " And A.����ID=[1] And A.����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", ����ID, intinsure)
        If Not rsTemp.EOF Then Isȫ��ͳ�� = (rsTemp!ȫ��ͳ�� = 1)
    End If
End Function

Public Function AddDate(ByVal strOrin As String, Optional ByVal blnʱ As Boolean = False) As String
'���ܣ�Ϊ��ȫ��������Ϣ��������
    Dim strTemp As String
    Dim intPos As Integer
    
    strTemp = Trim(strOrin)
    
    If strTemp = "" Then
        AddDate = ""
        Exit Function
    End If
    
    intPos = InStr(strTemp, "-")
    If intPos = 0 Then
        intPos = InStr(strTemp, ".")
        If intPos <> 0 Then
            'ʹ�� . ��
            strTemp = Replace(strTemp, ".", "-")
        End If
    End If
    
    If intPos = 0 Then
        'û��"-",�ֹ�����
        intPos = Len(strTemp)
        If intPos <= 8 Then
            If intPos = 8 Then
                strTemp = Mid(strTemp, 1, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7, 2)
            ElseIf intPos > 4 Then
                strTemp = Left(strTemp, intPos - 4) & "-" & Mid(Right(strTemp, 4), 1, 2) & "-" & Right(strTemp, 2)
            ElseIf intPos > 2 Then
                strTemp = Format(Date, "yyyy") & "-" & Left(strTemp, intPos - 2) & "-" & Right(strTemp, 2)
            Else
                strTemp = Format(Date, "yyyy") & "-" & Format(Date, "MM") & "-" & strTemp
            End If
        End If
    Else
        If blnʱ = False Then
            If IsDate(strTemp) Then
                strTemp = Format(CDate(strTemp), "yyyy-MM-dd")
            End If
        Else
            '����Сʱ
            If InStr(strTemp, " ") > 0 Then
                '������Сʱ
                If IsDate(strTemp & ":00") Then
                    strTemp = Format(CDate(strTemp & ":00"), "yyyy-MM-dd HH:ss")
                End If
            Else
                If IsDate(strTemp) Then
                    strTemp = Format(CDate(strTemp), "yyyy-MM-dd HH:ss")
                End If
            End If
        End If
    End If
    
    AddDate = strTemp
End Function

Public Function Insert�����������(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str���㷽ʽ As String) As Boolean
'���ܣ��������������ݱ�������
'���������㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    Dim cnTemp As New ADODB.Connection
    Dim strDate As String
    Dim lngCount As Long, arr���㷽ʽ As Variant, arr��� As Variant
    
    'Ϊ�˷�ֹһ�����Ӵ���ν�������
    Set cnTemp = GetNewConnection
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    cnTemp.BeginTrans
    On Error GoTo errHandle
    
    gstrSQL = "zl_����ģ�����_Clear(" & lng����ID & "," & lng��ҳID & ")"
    cnTemp.Execute gstrSQL, , adCmdStoredProc
    
    arr���㷽ʽ = Split(str���㷽ʽ, "|")
    For lngCount = 0 To UBound(arr���㷽ʽ)
        If arr���㷽ʽ(lngCount) <> "" Then
            arr��� = Split(arr���㷽ʽ(lngCount), ";")
            If UBound(arr���) > 1 Then
                If Val(arr���(1)) <> 0 Then
                    gstrSQL = "zl_����ģ�����_Insert(" & lng����ID & "," & IIf(lng��ҳID = 0, "null", lng��ҳID) & _
                        ",'" & arr���(0) & "'," & Val(arr���(1)) & "," & strDate & ")"
                    cnTemp.Execute gstrSQL, , adCmdStoredProc
                End If
            End If
        End If
    Next
    
    cnTemp.CommitTrans
    Insert����������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    cnTemp.RollbackTrans
End Function

Public Function Clear�����������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��ڽ���֮�󣬽����������������
    
    gstrSQL = "zl_����ģ�����_Clear(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
    
    Clear����������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get��������(ByVal str���֤ As String, ByVal lng���� As Long) As String
'���ܣ��������֤���������õ���������
    Dim strDate As String
    If Len(str���֤) = 15 Then
        '��ʽ�����֤��
        strDate = AddDate(Mid(str���֤, 7, 6))
        strDate = "19" & strDate
    ElseIf Len(str���֤) = 18 Then
        '��ʽ�����֤��
        strDate = AddDate(Mid(str���֤, 7, 8))
    Else
        'û�����֤��
        strDate = Format(DateAdd("yyyy", lng���� * -1, Date), "yyyy-MM-dd")
    End If
    
    If IsDate(strDate) = True Then
        Get�������� = Format(CDate(strDate), "yyyy-MM-dd")
    End If
End Function

Public Function GetOracleFormat(ByVal dat���� As Date)
    GetOracleFormat = "To_Date('" & Format(dat����, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Public Sub RemoveSelect(lvw As ListView)
'���ܣ�ɾ����ǰѡ����
    Dim lngIndex  As Long
    
    With lvw
        If .SelectedItem Is Nothing Then Exit Sub
        
        lngIndex = .SelectedItem.Index
        .ListItems.Remove lngIndex
        
        If .ListItems.Count > 0 Then
            '��������б��������һ��ѡ��
            lngIndex = IIf(.ListItems.Count > lngIndex, lngIndex, .ListItems.Count)
            .ListItems(lngIndex).Selected = True
            .ListItems(lngIndex).EnsureVisible
        End If
    End With

End Sub

Public Function CanסԺ�������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ��жϲ��˵�סԺ���������Ƿ��������ϡ��жϱ�׼�Ǽ�鲡�����µ�סԺ��¼������У��Ͳ��ܽ�����
'������lng����ID     ����ID
'      lng��ҳID     �ý��ʼ�¼���ڵ�סԺ����
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle

    gstrSQL = "SELECT COUNT(*) as סԺ���� FROM ������ҳ WHERE ����ID=[1] AND ��ҳID>[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�жϲ��˵�סԺ���������Ƿ���������", lng����ID, lng��ҳID)
    If rsTemp("סԺ����") > 0 Then
        MsgBox "�ò����Ѿ����µ�סԺ��¼������������ǰסԺ�Ľ������ݡ�", vbInformation, gstrSysName
        Exit Function
    End If

    CanסԺ������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ҽ�������Ѿ���Ժ(ByVal lng����ID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = " Select DECODE(��Ժ����,NULL,0,1) AS ��Ժ״̬ From ������ҳ " & _
              " Where (����ID,��ҳID) IN (Select ����ID,סԺ���� From ������Ϣ Where ����ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж�ҽ�������Ƿ��Ժ", lng����ID)
    ҽ�������Ѿ���Ժ = (rsTmp!��Ժ״̬ = 1)
End Function

Public Function ����δ�����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim rs���� As New ADODB.Recordset
    '���ô�סԺ�Ƿ��з���δ����
    #If gverControl >= 5 Then
        gstrSQL = "Select nvl(�������,0) as ���  from ������� where ����ID=[1] and ����=1 And ����=2"
    #Else
        gstrSQL = "Select nvl(�������,0) as ���  from ������� where ����ID=[1] and ����=1"
    #End If
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ����δ�����", lng����ID)
    If rs����.EOF = True Then
        ����δ����� = False
    Else
        ����δ����� = (rs����("���") <> 0)
    End If
End Function

Public Function ��ȡ���Ժ���(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
Optional ByVal bln��Ժ��� As Boolean = True, Optional ByVal bln����� As Boolean = True, _
Optional ByVal bln�������� As Boolean = False) As String
    
    '1-�������;2-��Ժ���;3-��Ժ���
    Dim rs��� As New ADODB.Recordset
    If bln�������� = False Then
        gstrSQL = " Select A.������Ϣ" & _
                  " From ������ A" & _
                  " Where A.����ID=[1] And A.��ҳID=[2]" & _
                  " And A.�������=[3] And ��ϴ���=1"
    Else
        gstrSQL = " Select A.������Ϣ,B.���� ��������" & _
                  " From ������ A,��������Ŀ¼ B" & _
                  " Where A.����ID=[1] And A.��ҳID=[2]" & _
                  " And A.����ID=B.ID(+) And A.�������=[3]"
    End If
    Set rs��� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���Ժ���", lng����ID, lng��ҳID, IIf(bln��Ժ���, "1", "3"))
    
    ��ȡ���Ժ��� = ""
    If Not rs���.EOF Then
        ��ȡ���Ժ��� = IIf(IsNull(rs���!������Ϣ), "", rs���!������Ϣ)
    End If
    
    ��ȡ���Ժ��� = Trim(��ȡ���Ժ���)
    If Not bln����� And ��ȡ���Ժ��� = "" Then
        ��ȡ���Ժ��� = "��"
    End If
    If bln�������� Then
        If Not rs���.EOF Then
            ��ȡ���Ժ��� = ��ȡ���Ժ��� & "|" & Nvl(rs���!��������, " ")
        Else
            ��ȡ���Ժ��� = ��ȡ���Ժ��� & "| "
        End If
    End If
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDO As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDO = 1 To 12
        strSource = Mid(strOld, intDO, 1)
        strTarget = Mid(strPass, intDO, 1)
        strReturn = strReturn & Chr(asc(strSource) Xor asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function ��������(ByVal int���� As Integer) As Boolean
    Dim rs���� As New ADODB.Recordset
    
    �������� = False
    gstrSQL = "Select Nvl(��������,0) ���� From ������� Where ���=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�������", int����)
    If Not rs����.EOF Then
        �������� = (rs����!���� = 1)
    End If
End Function

Private Function GetPatiInfo(lngID As Long) As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrH
    
'    strSql = "Select * From ������Ϣ A,������ҳ B Where A.����ID=B.����ID(+) And A.����ID=" & lngID & " Order by ��ҳID"
    '��ҳID=0ʱ(����NULL)����ʾԤԼ��Ժ
    strSQL = _
        " Select A.����ID,Decode(B.����ID,NULL,NULL,Nvl(B.��ҳID,0)) as ��ҳID," & _
        " A.����,A.סԺ��,B.��Ժ����,B.��Ժ����" & _
        " From ������Ϣ A,������ҳ B" & _
        " Where A.����ID=B.����ID(+) And A.����ID=[1]" & _
        " Order by Nvl(B.��ҳID,0)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˵�סԺ��Ϣ", lngID)
    If Not rsTmp.EOF Then Set GetPatiInfo = rsTmp
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
        
Private Function MergePatient(ByVal lngOld As Long, ByVal lngInsure As Long) As Long
    Dim i As Integer, j As Integer
    Dim curDate As Date, strSQL As String
    Dim rsPatiS As New ADODB.Recordset
    Dim rsPatiO As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    
    Set rsPatiS = GetPatiInfo(lngOld)
    Set rsPatiO = GetPatiInfo(lngInsure)
        
    'A��B��һ��������ԤԼ��Ժ
    If Not IsNull(rsPatiS!��ҳID) And Nvl(rsPatiS!��ҳID, 0) = 0 Then
        MsgBox "����:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]������ԤԼ��Ժ�Ǽǣ�����ȡ���õǼǡ�", vbInformation, gstrSysName
    End If
    If Not IsNull(rsPatiO!��ҳID) And Nvl(rsPatiO!��ҳID, 0) = 0 Then
        MsgBox "����:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]������ԤԼ��Ժ�Ǽǣ�����ȡ���õǼǡ�", vbInformation, gstrSysName
    End If
        
    'AB��ס��Ժ
    If Not IsNull(rsPatiS!��ҳID) And Not IsNull(rsPatiO!��ҳID) Then
        '1.��סԺ����Ժ,������(�Ⱥ�סԺ����Ϊ����Ժ-��Ժ,��Ժ-��Ժ����������Ժ-��Ժ,��Ժ-��Ժ)
        '��Ϊ�����˺ϲ���,���򲻶��⴦���Զ���Ժ������Ժ
        rsPatiS.MoveLast
        rsPatiO.MoveLast
        If rsPatiS!��Ժ���� <= rsPatiO!��Ժ���� Then
            If IsNull(rsPatiS!��Ժ����) Then
                MsgBox "����:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]���һ��סԺ����Ժ,����ǰδ��Ժ,����ִ�кϲ�������", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If IsNull(rsPatiO!��Ժ����) Then
                MsgBox "����:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]���һ��סԺ����Ժ,����ǰδ��Ժ,����ִ�кϲ�������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '2.ʱ�佻����ʾ�Ƿ����
        curDate = zlDatabase.Currentdate
        rsPatiS.MoveFirst
        For i = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For j = 1 To rsPatiO.RecordCount
                If Not (rsPatiO!��Ժ���� >= IIf(IsNull(rsPatiS!��Ժ����), curDate, rsPatiS!��Ժ����) Or _
                    IIf(IsNull(rsPatiO!��Ժ����), curDate, rsPatiO!��Ժ����) <= rsPatiS!��Ժ����) Then
                    If MsgBox("���ֲ���:" & rsPatiS!���� & "[" & rsPatiS!סԺ�� & "]�� " & rsPatiS!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiS!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiS!��Ժ����), curDate, rsPatiS!��Ժ����), "yyyy-MM-dd") & vbCrLf & _
                        "�벡��:" & rsPatiO!���� & "[" & rsPatiO!סԺ�� & "]�ĵ� " & rsPatiO!��ҳID & " ��סԺ���ڼ�" & Format(rsPatiO!��Ժ����, "yyyy-MM-dd") & "��" & Format(IIf(IsNull(rsPatiO!��Ժ����), curDate, rsPatiO!��Ժ����), "yyyy-MM-dd") & _
                        vbCrLf & "���ཻ�棬Ӧ�ò���ͬһ�����ˣ�ȷʵҪ�ϲ���", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
                rsPatiO.MoveNext
            Next
            rsPatiS.MoveNext
        Next
    End If
    
    '$IF HIS9
    #If gverControl = 0 Then
        strSQL = "zl_������Ϣ_MERGE(" & lngOld & "," & lngInsure & ")"
    #Else
    '$ELSE  HIS+
        strSQL = "zl_������Ϣ_MERGE(" & lngOld & "," & lngInsure & ", 'ҽ������ǼǺϲ�','" & gstrUserName & "')"
    #End If
    
    DoEvents
    Screen.MousePointer = 11
    Call zlDatabase.ExecuteProcedure(strSQL, "������ݺϲ�")
    Screen.MousePointer = 0
    
    '�ϲ���Ӧֻʣһ������
    strSQL = "Select ����ID From ������Ϣ Where ����ID IN([1],[2])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�ϲ���Ӧֻʣһ������", lngOld, lngInsure)
    If Not rsTmp.EOF Then
        If glngSys Like "8??" Then
            MsgBox "�ͻ��ϲ��ɹ�,�ϲ���Ŀͻ�IDΪ " & rsTmp!����ID & "��", vbInformation, gstrSysName
        Else
            MsgBox "���˺ϲ��ɹ�,�ϲ���Ĳ���IDΪ " & rsTmp!����ID & "��", vbInformation, gstrSysName
        End If
        MergePatient = rsTmp!����ID
    End If
End Function

Public Sub DebugTool(ByVal strInfo As String)
    Call LogWrite("ҽ���ӿڵ�����־", glngModul, "ҽ���ӿڷ���", strInfo)
End Sub

Public Function SystemImes() As Variant
'���ܣ���ϵͳ�������뷨���Ʒ��ص�һ���ַ���������
'���أ�����������������뷨,�򷵻ؿմ�
    Dim arrIme(99) As Long, ARRNAME() As String
    Dim lngLen As Long, STRNAME As String * 255
    Dim lngCount As Long, i As Integer, j As Integer
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then
            ReDim Preserve ARRNAME(j)
            lngLen = ImmGetDescription(arrIme(i), STRNAME, Len(STRNAME))
            ARRNAME(j) = Mid(STRNAME, 1, InStr(STRNAME, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, ARRNAME, vbNullString)
End Function

Public Function OpenIme(Optional strIme As String) As Boolean
'����:�����ƴ��������뷨,��ָ������ʱ�ر��������뷨��֧�ֲ������ơ�
    Dim arrIme(99) As Long, lngCount As Long, STRNAME As String * 255
    
    If strIme = "���Զ�����" Then OpenIme = True: Exit Function
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), STRNAME, Len(STRNAME)
            If InStr(1, Mid(STRNAME, 1, InStr(1, STRNAME, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function ���֤��ת��������(ByVal str���֤�� As String, ByRef str��������) As Boolean
    
    Dim intI As Integer
    ���֤��ת�������� = True
    '��֤����Ĳ����Ƿ����Ҫ��
    For intI = 1 To Len(str���֤��)
        If InStr("0123456789", Mid(str���֤��, intI, 1)) <= 0 Then
            If intI = 18 Then
                If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", Mid(str���֤��, intI, 1)) <= 0 Then
                    str�������� = "���֤�����а�����Ч�ַ�!"
                    ���֤��ת�������� = False
                End If
            Else
                str�������� = "���֤�����а�����Ч�ַ�!"
                ���֤��ת�������� = False
            End If
        End If
    Next
    
    If ���֤��ת�������� = True Then
        Select Case Len(str���֤��)
            Case 15
                str�������� = "19" & Mid(str���֤��, 7, 6)
                If IsDate(Mid(str��������, 1, 4) & "-" & Mid(str��������, 5, 2) & "-" & Mid(str��������, 7, 2)) = False Then
                    str�������� = "���֤�����д���!"
                    ���֤��ת�������� = False
                End If
            Case 18
                str�������� = Mid(str���֤��, 7, 8)
                If IsDate(Mid(str��������, 1, 4) & "-" & Mid(str��������, 5, 2) & "-" & Mid(str��������, 7, 2)) = False Then
                    str�������� = "���֤�����д���!"
                    ���֤��ת�������� = False
                End If
            Case Else
                str�������� = "���֤����λ������!"
                ���֤��ת�������� = False
        End Select
    End If
    
End Function

Public Function IsApartComponents(ByVal intinsure As Integer) As Boolean
    On Error GoTo errHand
    'Ϊ�˱����ϲ����Ĳ������ӣ����û��ֲ������ڶ�����ʹ��ҽ���µĹ���ģʽ��������Ӵ˻��ڣ�����Ƿ���Ĳ������򵥶�����
    '���ò����Ƿ��Ƿ���Ĳ���
    If rsInsure.State = 0 Then
        gstrSQL = "Select ���,ҽ������,ҽ���� From ������� "
        Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����������")
    End If
    rsInsure.Filter = "���=" & intinsure
    If rsInsure.RecordCount = 0 Then rsInsure.Filter = 0: Exit Function
    If Nvl(rsInsure!ҽ������) = "" Then rsInsure.Filter = 0: Exit Function
    rsInsure.Filter = 0
    
    IsApartComponents = True
errHand:
End Function

Public Function CreateObject_Insure(ByVal intinsure As Integer, ByRef intOrder As Integer, Optional ByVal intCall As Integer = 0) As Boolean
    Dim blnExist As Boolean
    Dim strObject As String, strBag As String
    Dim intObject As Integer, intCOUNT As Integer
    Dim objTemp As Object
    '����˵��:
    'intCall:0-δָ��ҽ�����򴴽�ҽ������;1-��Identify�⣬����ҵ��ǿ�Ƶ��ø��Ե�ҽ������
    
    On Error GoTo errHand
    '����ҽ���ӿڶ���
    If rsInsure.State = 0 Then
        gstrSQL = " Select ���,ҽ������,ҽ���� From ������� "
        Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����������")
    End If
    rsInsure.Filter = "���=" & intinsure
    If rsInsure.RecordCount = 0 Then
        rsInsure.Filter = 0
        MsgBox "��ҽ���ӿڻ�δע�ᣡ���=" & intinsure, vbInformation, gstrSysName
        Exit Function
    End If
    strBag = Nvl(UCase(rsInsure!ҽ����))
    strObject = UCase(Nvl(rsInsure!ҽ����, rsInsure!ҽ������))
    If intCall = 1 Then strObject = UCase(rsInsure!ҽ������)
    rsInsure.Filter = 0
    
    '����Ƿ���ڸö���
    On Error Resume Next
    intCOUNT = UBound(gobjInsure_Name)
    If Err <> 0 Then intCOUNT = -1
    
    'Ӧ������Ҫ�����Σ���Ϊʹ�����µĿؼ����� 2008-10-17
    'On Error GoTo errHand
    For intObject = 0 To intCOUNT
        If gobjInsure_Name(intObject) = strObject Then
            If Not gobjInsure_Obj(intObject) Is Nothing Then
                intOrder = intObject
                CreateObject_Insure = True
                Exit Function
            Else
                blnExist = True
                Exit For
            End If
        End If
    Next
    
    'ȥ���ļ�����׺
    strObject = Mid(strObject, 1, Len(strObject) - 4)
    '��������
    Set objTemp = CreateObject(strObject & ".Cls" & Mid(strObject, 4))
    If objTemp Is Nothing Then Exit Function
    intObject = intCOUNT + 1
    ReDim Preserve gobjInsure_Name(intObject)
    ReDim Preserve gobjInsure_Obj(intObject)
    gobjInsure_Name(intObject) = strObject & ".DLL"
    Set gobjInsure_Obj(intObject) = objTemp
    intOrder = intObject
    
    'ҽ�����������Ҫ���ó�ʼ������
    If strBag <> "" Then
        If Not gobjInsure_Obj(intObject).InitInsure(gcnOracle, intinsure) Then Exit Function
    End If
    
    CreateObject_Insure = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ChooseInsure_Base(ByVal intinsure As Integer) As Integer
    '��ʽ���£�
    '����ҽ���ӿ�
    '31-������·��ҽ��
    '51-������ҽ��
    '���ܣ����ѡ������ҽ���ӿڡ�������ԭ�������̣����򴴽�ָ���ķ���ҽ����������������CodeMan()
    Dim intSelect As Integer
    Dim rsTemp As New ADODB.Recordset
    
    ChooseInsure_Base = intinsure
    '����Ƿ���ڶ���������ʽʵ�ֵ�ҽ���ӿ�
    gstrSQL = "Select count(*) AS Records From ������� Where ҽ������ Is Not NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���ڶ�������")
    If Nvl(rsTemp!Records, 0) = 0 Then Exit Function
    
    '����ѡ����������Աѡ��
    intSelect = frmѡ��ǰҽ��_Base.ShowSelect
    ChooseInsure_Base = intSelect
End Function

Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '���¼�¼,���������,������
    'strPrimary:�ֶ���,ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Function �൥���շ�_�շѷֱ��ӡ() As Boolean
    #If gverControl >= 4 Then
        �൥���շ�_�շѷֱ��ӡ = (Val(zlDatabase.GetPara(78, glngSys, , 0)) = 1)
    #Else
        �൥���շ�_�շѷֱ��ӡ = (Val(GetPara(78, glngSys, , , 0)) = 1)
    #End If
End Function

Public Sub �൥���շ�_�˷�(ByVal lng����ID As Long)
    '2.1���˷�ʱ��������ս����¼�ı�ע�к����൥���շѡ������Ҹõ����Ƕ൥���շѣ�Ʊ�ݴ�ӡ���ݴ���1����¼�������ִ��
    '2.2������˷�ʱ�õ����ǵ����շѣ�Ʊ�ݴ�ӡ����С�ڵ���1����¼��������ȡ�ò��˸õǼ�ʱ����ͬ�����е��ݺų�������ʾ����ԱӦ��ͬʱ�˷Ѻ�����ȡ�µķ���
    Dim strNO As String, str�����嵥 As String
    Dim lng����ID As Long
    Dim str�Ǽ�ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '��鱣�ս����¼���Ƿ��¼���Ƕ൥���շѣ�ע�⣬��Ȼ���ս����¼�м�¼Ϊ�൥���շѣ��������ڹ�ѡ��ϵͳ������78-���ŵ����շѷֱ��ӡ����HIS��û�е����൥���������˷�ʱ�������ˣ�����Ҫ�ж�
    gstrSQL = " Select ��ע From ���ս����¼ Where ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�Ϊ�൥���շ�", lng����ID)
    If InStr(1, rsTemp!��ע, "�൥���շ�") = 0 Then Exit Sub        '���Ƕ൥���շ�ֱ���˳�
    '��ȡ���ν��ʵ������Ϣ
    gstrSQL = " Select NO,����ID,�Ǽ�ʱ�� From ������ü�¼ Where ����ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ݺ��벡��ID", lng����ID)
    lng����ID = rsTemp!����ID
    strNO = rsTemp!NO
    str�Ǽ�ʱ�� = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss")
    '����Ʊ�ݴ�ӡ�����ж��Ƿ�൥���շ�
    gstrSQL = " Select NO From Ʊ�ݴ�ӡ���� Where ID=(Select ID From Ʊ�ݴ�ӡ���� Where ��������=1 And NO=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ʊ�ݴ�ӡ�����ж��Ƿ�൥���շ�", strNO)
    If rsTemp.RecordCount > 1 Then Exit Sub                       '������¼˵��δ��ѡϵͳ������HIS��Ϊ�Ƕ൥���շ�
    '��ȡ�Ǽ�ʱ�䣬����ID��ͬ�ĵ����嵥����ʾ����Ա
    gstrSQL = " Select Distinct NO From ������ü�¼ " & _
              " Where Mod(��¼����,10)=1 And ��¼״̬=1 And ����ID=[1] And �Ǽ�ʱ��=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�Ǽ�ʱ�䣬����ID��ͬ�ĵ����嵥", lng����ID, CDate(str�Ǽ�ʱ��))
    With rsTemp
        Do While Not .EOF
            If !NO <> strNO Then str�����嵥 = str�����嵥 & "," & !NO
            .MoveNext
        Loop
        If str�����嵥 <> "" Then
            str�����嵥 = Mid(str�����嵥, 2)
            MsgBox "�൥���շѣ��˷�ʱ��һ��������µ��ݵ��˷ѣ�Ȼ�������շѣ�" & vbCrLf & str�����嵥, vbInformation, gstrSysName
        End If
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function
Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModul As Long, _
    Optional ByVal blnPrivate As Boolean, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '-------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����ϵͳ����ֵ
    '����:varPara-�����Ż�������������ֻ��ַ����ʹ�������
    '     lngSys-ϵͳ��(10.20.0�Ժ�汾��Ч)
    '     lngModul-ģ���(10.20.0�Ժ�汾��Ч)
    '     blnPrivate-�Ƿ�˽��ģ��(10.20.0�Ժ�汾��Ч)
    '     strDefault-Ĭ��ֵ
    '     blnNotCache-�Ƿ��л����ж�ȡ(10.20.0�Ժ�汾��Ч)
    '����:����ֵ
    '����:
    '����:2008/01/04
    '-------------------------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If IsToolInPara Then
        GetPara = CallByName(zlDatabase, "GetPara", VbMethod, varPara, IIf(lngSys = 0, glngSys, lngSys), lngModul, blnPrivate, strDefault, blnNotCache)
    Else
        If TypeName(varPara) = "String" Then
            gstrSQL = "Select ����ֵ From ϵͳ������ where ������=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡϵͳ����", CStr(varPara))
        Else
            gstrSQL = "Select ����ֵ From ϵͳ������ where ������=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡϵͳ����", Val(varPara))
        End If
        If rsTemp.RecordCount <> 0 Then
            GetPara = Nvl(rsTemp!����ֵ, strDefault)
        Else
            GetPara = strDefault
        End If
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IsToolInPara() As Boolean
    '---------------------------------------------------------------------------------
    '����:�ж��Ƿ��zlTools�е�zlParameters�ж�ȡ����ֵ
    '����:��,����true,���򷵻�Fasle
    '����:
    '����:2008/01/04
    '---------------------------------------------------------------------------------
    Dim arrVersion
    Dim rsTemp As New ADODB.Recordset
    
    '��ҽ������ֻ��CodeMan()���ܻ�ȡϵͳ�ţ��ڶ�ȡ����ʱ����֪��ϵͳ�ţ���д��ע������ҽ��������Ĭ��Ϊ 100
    glngSys = GetSetting("ZLSOFT", "����ȫ��", "ϵͳ��", 100)

    'ȡϵͳ�汾��
    gstrSQL = "Select �汾�� From zlSystems Where  ��� =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡϵͳ�汾��", glngSys)
    
    '�жϰ汾��
    arrVersion = Split(rsTemp!�汾��, ".")
    If arrVersion(0) = "10" Then
        '���˰汾
        If Val(arrVersion(1)) < 20 Then
            'ֻ�дΰ汾��20���²��ܶ�ȡ
            IsToolInPara = False
        Else
            IsToolInPara = True
        End If
    End If
End Function

Public Function IsZLHIS10() As Boolean
    Dim arrVersion
    Dim rsTemp As New ADODB.Recordset
    
    'ȡϵͳ�汾��
    gstrSQL = "Select �汾�� From zlSystems Where Floor(���/100)=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡϵͳ�汾��", 1)
   
    '�жϰ汾��
    arrVersion = Split(rsTemp!�汾��, ".")
    If arrVersion(0) = "10" Then
        IsZLHIS10 = True
    End If
End Function

Public Sub OpenRecordset_OtherBase(rsTmp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional gcnConnect As ADODB.Connection)
'���ܣ��򿪼�¼��
    If rsTmp.State = adStateOpen Then rsTmp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
'    rsTmp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle, adOpenStatic, adLockReadOnly
    If gcnConnect Is Nothing Then Set gcnConnect = gcnOracle
    Set rsTmp = gcnConnect.Execute(IIf(strSQL = "", gstrSQL, strSQL))
    Call SQLTest
End Sub


Public Sub OpenRecordset(rsTmp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional gcnConnect As ADODB.Connection)
'���ܣ��򿪼�¼��
    If rsTmp.State = adStateOpen Then rsTmp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
'    rsTmp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle, adOpenStatic, adLockReadOnly
    If gcnConnect Is Nothing Then Set gcnConnect = gcnOracle
    Set rsTmp = gcnConnect.Execute(IIf(strSQL = "", gstrSQL, strSQL))
    Call SQLTest
End Sub


