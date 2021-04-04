Attribute VB_Name = "mdlPubDefine"
Option Explicit
Public Const gstrSplitCmb = "-"

'�������ݲ˵�ID����:*��ʾ��ͼ��
'*********************************************************************
Public Const conMenu_FilePopup = 1 '�ļ�
Public Const conMenu_ManagePopup = 2 '����
Public Const conMenu_EditPopup = 3 '�༭
Public Const conMenu_ReportPopup = 4 '����
Public Const conMenu_ViewPopup = 7 '�鿴
Public Const conMenu_ToolPopup = 8 '����
Public Const conMenu_HelpPopup = 9 '����
Public Const gconLockColor = &H80000000
Public Const gconEditColor = &HC0FFC0
'�ļ��˵�
Public Const conMenu_File_Open = 100            '*��(&O)��
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
Public Const conMenu_File_RoomSet = 182         'ִ�м��豸
Public Const conMenu_File_SendImg = 184         '����ͼ��
Public Const conMenu_File_Exit = 191            '*�˳�(&X)
Public Const conMenu_File_ExportToXML = 192     '���ΪXML�ĵ�
Public Const conMenu_File_ImportFromXML = 193   '��XML�ĵ�����
Public Const conMenu_File_BillPrintView = 194   '���ݴ�ӡԤ��
Public Const conMenu_File_BillPrint = 195       '���ݴ�ӡ

'����˵�:����վ����Ĺ��ܲ˵�
Public Const conMenu_Manage_Regist = 211      '*���˹Һ�(&H)
Public Const conMenu_Manage_Bespeak = 212     'ԤԼ�Һ�(&B),ʱ�䰲��(&B)
Public Const conMenu_Manage_Transfer = 213    '*ת�ﴦ��(&C)
Public Const conMenu_Manage_Transfer_Send = 2131      '����ת��(&S)
Public Const conMenu_Manage_Transfer_Cancel = 2132    'ȡ��ת��(&C)
Public Const conMenu_Manage_Transfer_Incept = 2133    'ת�����(&I)
Public Const conMenu_Manage_Transfer_Refuse = 2134    'ת��ܾ�(&R)
Public Const conMenu_Manage_Transfer_Force = 2135     'ǿ������(&F)
Public Const conMenu_Manage_Receive = 214     '*���˽���(&Z)
Public Const conMenu_Manage_Cancel = 215      'ȡ������(&Q)
Public Const conMenu_Manage_Finish = 216      '*��ɽ���(&W)
Public Const conMenu_Manage_Redo = 217        '�ָ�����(&R)

Public Const conMenu_Manage_Call = 218      '����
Public Const conMenu_Manage_CallNext = 21801      '��һ��(&N)
Public Const conMenu_Manage_CallPrevious = 21802     '��һ��(&P)

Public Const conMenu_Manage_Reset = 219     '����˳��
Public Const conMenu_Manage_Up = 21901    '����(&U)
Public Const conMenu_Manage_Down = 21902     '����(&D)
Public Const conMenu_Manage_Discard = 21903      '����(&D)
Public Const conMenu_Manage_Recall = 21904      '�ٻ�(&R)
Public Const conMenu_Manage_Untread = 21905        '�˺�(&R)

Public Const conMenu_Manage_Plan = 221        '*ִ�б���(&P)
Public Const conMenu_Manage_Logout = 222      'ȡ������(&L)
Public Const conMenu_Manage_Refuse = 223      '�ܾ�ִ��(&R)
Public Const conMenu_Manage_ReGet = 224       'ȡ���ܾ�(&G)
Public Const conMenu_Manage_Complete = 225    '*ִ�����(&C)
Public Const conMenu_Manage_Undone = 226      'ȡ�����(&U)
Public Const conMenu_Manage_ThingAdd = 227    '*��¼ִ�����(&A)
Public Const conMenu_Manage_ThingModi = 228   '*����ִ�����(&M)
Public Const conMenu_Manage_ThingDel = 229    '*ɾ��ִ�����(&D)
Public Const conMenu_Manage_ClearUp = 233     '��鱨�沵��(&U)

Public Const conMenu_Manage_Request = 231        '*����(&V)
Public Const conMenu_Manage_RequestView = 2311           '��������(&V)
Public Const conMenu_Manage_RequestPrint = 2312           '��ӡ���Ƶ���(&J)
Public Const conMenu_Manage_RequestBatPrint = 2313           '������ӡ����(&B)
Public Const conMenu_Manage_Report = 232         '*����(&O)
Public Const conMenu_Manage_ReportEdit = 2321        '��д����(&E)
Public Const conMenu_Manage_ReportView = 2322        '���ı���(&W)
Public Const conMenu_Manage_ReportPrint = 2323       '�����ӡ(&P)
Public Const conMenu_Manage_ReportPreview = 2324     'ִ��Ԥ��(&V)
Public Const conMenu_Manage_LeaveMedi = 251 '�Ĵ�ҩƷ

Public Const conMenu_Manage_Audit = 252         '*�������
Public Const conMenu_Manage_UnAudit = 253       '*ȡ�����
Public Const conMenu_Manage_Arrange = 254       '*ִ�а���
Public Const conMenu_Manage_UnArrange = 255     '*ȡ������

'ҽ��(�༭)�˵�����϶�,����ʱ��4λ���,50λ�ֶ�,001-050,051-100,101-150,...
Public Const conMenu_Edit_NewItem = 3001    '*����Ŀ(&A)
Public Const conMenu_Edit_Append = 3002     '*����/��¼(&Y)
Public Const conMenu_Edit_Modify = 3003     '*�޸�(&M)
Public Const conMenu_Edit_Delete = 3004     '*ɾ��(&D)
Public Const conMenu_Edit_Blankoff = 3005   '*����(&B)
Public Const conMenu_Edit_Stop = 3006       '*ҽ��ֹͣ(&S)
Public Const conMenu_Edit_ReStop = 3007     '*ȷ��ֹͣ(&C)
Public Const conMenu_Edit_Pause = 3008      '*��ͣ(&P)
Public Const conMenu_Edit_Reuse = 3009      '*����(&U)
Public Const conMenu_Edit_Audit = 3010      '*���/У��(&V)
Public Const conMenu_Edit_Price = 3011      '*�Ƽ۵���(&I)
Public Const conMenu_Edit_ClearUp = 3012    '*ҽ������(&F)
Public Const conMenu_Edit_Send = 3013       '*����(&G)
Public Const conMenu_Edit_SendDrug = 30131      '*ҩ��ҽ������(&1)
Public Const conMenu_Edit_SendOther = 30132     '����ҽ������(&2)
Public Const conMenu_Edit_Untread = 3014    '*����(&R)
Public Const conMenu_Edit_SendBack = 3015   '*���ڷ����ջ�(&N)
Public Const conMenu_Edit_Test = 3016       '*Ƥ�Խ��(&T)
Public Const conMenu_Edit_ChargeOff = 3017       '*���ó���(&E)
Public Const conMenu_Edit_NoPrint = 3018    '���δ�ӡ(&I)
Public Const conMenu_Edit_ChargeDelApply = 3019 '*��������(&L)
Public Const conMenu_Edit_ChargeDelAudit = 3020 '*�������(&U)

'����(�༭)�˵�
Public Const conMenu_Edit_NewParent = 3051   '*�·���(&N)
Public Const conMenu_Edit_Insert = 3052      '*����(&I)
Public Const conMenu_Edit_ModifyParent = 3053 '*�޸ķ���(&M)
Public Const conMenu_Edit_DeleteParent = 3054 '*ɾ������(&D)
Public Const conMenu_Edit_MarkMap = 3061     '*ͼƬ(&I)��
Public Const conMenu_Edit_ApplyTo = 3062     '*���ÿ���(&T)
Public Const conMenu_Edit_Request = 3063     '����Ҫ��(&R)
Public Const conMenu_Edit_Compend = 3064     '*���ݹ���(&F)
Public Const conMenu_Edit_Import = 3071      '*��������(&B)��
Public Const conMenu_Edit_Adjust = 3082      '*����(&J)
Public Const conMenu_Edit_Archive = 3083     '*�鵵(&R)
Public Const conMenu_Edit_UnArchive = 3084     'ȡ���鵵(&D)
Public Const conMenu_Edit_Save = 3091        '*����
Public Const conMenu_Edit_Sort = 3092        '*���ĵ�����
Public Const conMenu_Edit_Privacy = 3093     '*������˽��������

Public Const conMenu_Edit_Select = 3094      '*ѡ��
Public Const conMenu_Edit_DeSelect = 3095    '*ȡ��ѡ��

Public Const conMenu_Edit_Merge = 3096

'Public Const conMenu_Manage_ThingAdd = 227    '�ӵ�(&A)
'Public Const conMenu_Manage_ThingModi = 228   '*����ִ�����(&M)
Public Const conMenu_Edit_Transf_Delete = 229   '�����ӵ�

'���ϵͳ���� 32��ͷ�ĺ�
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_SaveExit = 3200           '���沢�˳�
Public Const conMenu_Edit_SizeFit = 3201            '��ʽ����
Public Const conMenu_Edit_SourceFit = 3202          '��Դ����
Public Const conMenu_Edit_Camera = 3203             '�����豸
Public Const conMenu_Edit_TakePicture = 3204        '����
Public Const conMenu_Edit_SelAll = 3205             'ȫѡ
Public Const conMenu_Edit_ClsAll = 3206             'ȫ��
Public Const conMenu_Edit_CallBack = 3207           '��������
Public Const conMenu_Edit_Money = 3208              '���÷�ʽ
Public Const conMenu_Edit_Pay = 3209                '֧����ʽ
Public Const conMenu_Edit_CheckItem = 3210          '�����Ŀ
Public Const conMenu_Edit_ChargeItem = 3211         '�շ���Ŀ

'������Ŀ(�༭)�˵� 3501-3530
Public Const conMenu_Edit_Transf_Modify = 3502   '�޸ĵ���
Public Const conMenu_Edit_Transf_Save = 3503     '����
Public Const conMenu_Edit_Transf_Cancle = 3504   'ȡ��

Public Const conMenu_Edit_Transf_UndoEnd = 3505  '�������
Public Const conMenu_Edit_Transf_Negative = 3506 '����(+)
Public Const conMenu_Edit_Transf_Positive = 3507 '����(-)
Public Const conMenu_Edit_Transf_Reprint = 3508  '�ش򵥾�

'������λ(�༭)�˵� 3531-3559
Public Const conMenu_Edit_Seat = 3530        '��λ
Public Const conMenu_Edit_Seat_Add = 3531    '��λ����
Public Const conMenu_Edit_Seat_Modify = 3532 '��λ�޸�
Public Const conMenu_Edit_Seat_Delete = 3533 '��λɾ��
Public Const conMenu_Edit_Seat_Clear = 3534  '���ռ�õ���λ
Public Const conMenu_Edit_Seat_Set = 3535    '������λ
Public Const conMenu_Edit_Seat_Swap = 3536    '������λ

Public Const conMenu_Edit_Seat_View = 3551 '�鿴
Public Const conMenu_Edit_Seat_Icon = 3552 'ͼ�귽ʽ
Public Const conMenu_Edit_Seat_List = 3553 '�б�ʽ
Public Const conMenu_Edit_Seat_Report = 3554 '����ʽ

'�ݴ�ҩƷ(�༭)�˵� 3561 -3579
Public Const conMenu_Edit_Leave_Add = 3561 '����
Public Const conMenu_Edit_Leave_Modify = 3562 '�޸�
Public Const conMenu_Edit_Leave_Delete = 3563 'ɾ��
Public Const conMenu_Edit_Leave_Post = 3564 'ʹ�õǼ�
Public Const conMenu_Edit_Leave_SavePost = 3565 '����Ǽ�����
Public Const conMenu_Edit_Leave_UndoPost = 3565 '�����Ǽ�

Public Const conMenu_Edit_Leave_Repertory = 3571 '����ѯ
Public Const conMenu_Edit_Leave_AccountBook = 3572 '���̨��

'����ϵͳ���� 3580 -  3599
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_CopyNewItem = 3580        '*���Ʋ�����Ŀ
Public Const conMenu_Edit_Default = 3582            'ȱʡ���
Public Const conMenu_Edit_MakeCharge = 3586         '���ɷ���
Public Const conMenu_Edit_Preferences = 3587         '�ο�����

'Ѫ��ϵͳ���� 31��ͷ�ĺ�
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_NewKind = 311             '����Ʒ��
Public Const conMenu_Edit_ModifyKind = 312          '�޸�Ʒ��
Public Const conMenu_Edit_DeleteKind = 313          'ɾ��Ʒ��
Public Const conMenu_Edit_StorgeLimit = 314         '�������
Public Const conMenu_Edit_StorgeDept = 315          '�ⷿ
Public Const conMenu_Edit_StorgePostion = 316       '��λ
Public Const conMenu_Edit_Check = 3101              '�˶�
Public Const conMenu_Edit_View = 3102               '����
Public Const conMenu_Edit_ModifyBill = 3103         '�޸ķ�Ʊ
Public Const conMenu_Edit_Verify = 3104             '�������
Public Const conMenu_Edit_AdjustPrice = 3105        '����

'LISʹ�õĲɵ�
Public Const conMenu_Edit_QCRes = 3650         '�ʿ�Ʒ

'����˵�
Public Const conMenu_Report_DrugQuery = 401    'ҩ���շ���ѯ(&H)
Public Const conMenu_Report_Reports = 402      '�������ñ���(&W)
Public Const conMenu_Report_MultiBill = 403    '��ӡ�ಡ�˵���(&K)
Public Const conMenu_Report_ClinicBill = 404   '��ӡ���Ƶ���(&J)��
Public Const conMenu_Report_AdviceBill1 = 405  '����ҽ����(&P)
Public Const conMenu_Report_AdviceBill2 = 406  '��ʱҽ����(&T)
Public Const conMenu_Report_AdviceBill3 = 407  'ҽ����¼��(&B)
Public Const conMenu_Report_WorkLog = 408      '�����ձ�(&O)


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
Public Const conMenu_View_ReadIC = 724               '��IC��(&I)
Public Const conMenu_View_PatInfor = 725             '�鿴������Ϣ
Public Const conMenu_View_PriceBill = 727
Public Const conMenu_View_PriceTable = 728
Public Const conMenu_View_PriceList = 729
Public Const conMenu_View_FilterView = 730           '�Թ��˷�ʽ��ʾ
Public Const conMenu_View_Filter = 731               '*���ݹ���(&I),�Ӵ���Ĺ��˹���
Public Const conMenu_View_Notify = 732               '*ҽ������(&B)
Public Const conMenu_View_Busy = 733                 '����æ(&M)
Public Const conMenu_View_ShowAll = 734
Public Const conMenu_View_ShowHistory = 735
Public Const conMenu_View_ShowStoped = 736
Public Const conMenu_View_Hide = 741                 '*����(&H)
Public Const conMenu_View_Show = 742                 '*��ʾ(&S)
Public Const conMenu_View_Forward = 743              '*ǰ��(&F)
Public Const conMenu_View_Backward = 744             '*����(&B)
Public Const conMenu_View_Dept = 745                '�鿴����
Public Const conMenu_View_Location = 746            '��λ
Public Const conMenu_View_LocationItem = 747        '��λ��Ŀ
Public Const conMenu_View_Option = 781               'ѡ��(&O)
Public Const conMenu_View_Refresh = 791              '*ˢ��(&R)
Public Const conMenu_View_Jump = 792                 '��ת(&J)

'���ϵͳ����70��ͷ�ĺ�
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_View_Single = 7040             '����
Public Const conMenu_View_Group = 7041              '����
Public Const conMenu_View_LocationMethod = 7042     '��λ����
Public Const conMenu_View_Column = 7043             'ѡ������

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
Public Const conMenu_Tool_Option = 819          'ѡ��(&O),�Ӵ�������ù���


'�ɼ��˵�
Public Const conMenu_Cap_Dynamic = 8100         '��̬��ʾ(&V)
Public Const conMenu_Cap_MarkMap = 8101       'Ӱ��ɼ�(&C)
Public Const conMenu_Cap_Import = 8102        'Ӱ����(&I)
Public Const conMenu_Cap_DevSet = 8103          'Ӱ���豸����(&D)
Public Const comMenu_Cap_Process = 8104         'Ӱ����
Public Const conMenu_Cap_Record = 8105          '¼��(&R)
Public Const conMenu_Cap_Play = 8106          '����(&P)
Public Const conMenu_Cap_Stop = 8107            'ֹͣ(&T)
Public Const conMenu_Cap_Forward = 8108         '���(&F)
Public Const conMenu_Cap_Back = 8109            '����(&B)
Public Const conMenu_Cap_SaveAs = 8110          '����¼��(&S)


Public Const conMenu_Img_Look = 8111        'Ӱ���Ƭ(&S)
Public Const conMenu_Img_Contrast = 8112    '��Ƭ�Ա�(&E)
Public Const conMenu_Img_Delete = 8113        'ͼ��ɾ��(&K)
Public Const conMenu_Img_Query = 8114        'Q/R��ȡͼ��(&Q)



'ͼ����
Public Const conMenu_Process_Window = 501           '���ȶԱȶ�
Public Const conMenu_Process_Zoom = 502             '����
Public Const conMenu_Process_RRotate = 503          '˳ʱ����ת
Public Const conMenu_Process_LRotate = 504          '��ʱ����ת
Public Const conMenu_Process_Sharpness = 505        '��
Public Const conMenu_Process_Filter = 506           'ƽ��
Public Const conMenu_Process_Arrow = 507            '��ͷ��ע
Public Const conMenu_Process_Ellipse = 508          'Բ�α�ע
Public Const conMenu_Process_Text = 509             '���ֱ�ע


'�����˵�
Public Const conMenu_Help_Help = 901        '*��������(&H)
Public Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_Help_Web_Forum = 9023      '������̳(&F)
Public Const conMenu_Help_Web_Mail = 9022       '*���ͷ���(&M)
Public Const conMenu_Help_About = 991       '����(&A)��

Public Const conMenu_Edit_MediAudit = 3564 '*ҩ�����(&U)(������ҩ���)

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

Public Const VsModiBackColor = &HD6FFCA        'vs�ؼ����ɱ༭��Ԫ�ı���ɫ
'*********************************************************************
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Enum REGISTER
    ע����Ϣ
    ˽��ģ��
    ˽��ȫ��
    ����ģ��
    ����ȫ��
End Enum
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
     x As Long
     y As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hwnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
End Sub
Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String) As CommandBarControl
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption)
        
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        
    End With
    
    Set NewCommandBar = objControl
    
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.ID = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function


Public Function DockPannelCreate(ByRef dkpMain As DockingPane, ByVal intIndex As Integer, _
                                    ByVal lngCX As Long, ByVal lngCY As Long, _
                                    ByVal bytDirection As DockingDirection, _
                                    Optional ByVal objNeighbour As Pane = Nothing, _
                                    Optional ByVal strTitle As String, _
                                    Optional ByVal bytOptions As PaneOptions) As Pane
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Set DockPannelCreate = dkpMain.CreatePane(intIndex, lngCX, lngCY, bytDirection, objNeighbour)
    DockPannelCreate.Title = strTitle
    DockPannelCreate.Options = PaneNoCaption
    
End Function

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
End Function

Public Function TabControlInit(ByRef tbc As TabControl, _
                                Optional ByVal bytAppearance As XTPTabAppearanceStyle = xtpTabAppearancePropertyPage2003) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    With tbc
        
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
'            .Position = bytPosition
        End With
        
        Set .Icons = frmPubResource.imgPublic.Icons
        

        
    End With

    TabControlInit = True
    
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto 'xtpSystemThemeBlue
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    cbsMain.Options.LargeIcons = False
    
    CommandBarInit = True
    
End Function

Public Function CopyMenu(ByVal cbsMain As Object, Optional ByVal intNo As Integer = 2) As CommandBar
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '�����˵�����
    
    On Error GoTo errHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Set cbrPopupItem2 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.ID, cbrControl2.Caption)
                cbrPopupItem2.Parameter = cbrControl2.Parameter
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
errHand:
    Err.Clear
    Resume Next
End Function

Public Sub SetDockRight(cbsMain As Object, BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsMain.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsMain.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

Public Sub LocationObj(ByRef objTxt As Object, Optional ByVal blnDoevents As Boolean = False)
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If blnDoevents Then DoEvents
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
    
errHand:
    
End Sub

Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button            '������
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              'ͼ������
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '��ͼ��
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '״̬��
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet              '��ӡ����
    
        Call zlPrintSet
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel               '��ӡ����,Ԥ������,�����Excel
        
        If objPrnVsf Is Nothing Then Exit Function
        
        If Not SearchPrintData(objPrnVsf, frmPubResource.msfPrint) Then
            MsgBox "���ӡ�����粻�������ݣ������¼��ӣ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        '���ô�ӡ��������
        Set objPrint.Body = frmPubResource.msfPrint
        objPrint.Title.Text = strPrintTitle
        Set objAppRow = New zlTabAppRow
        Call objAppRow.Add("")
        Call objAppRow.Add("��ӡʱ��:" & Now())
        Call objPrint.BelowAppRows.Add(objAppRow)

        Select Case Control.ID
        Case conMenu_File_Print
            bytMode = zlPrintAsk(objPrint)
            If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
        Case conMenu_File_Preview
            zlPrintOrView1Grd objPrint, 2
        Case conMenu_File_Excel
            zlPrintOrView1Grd objPrint, 3
        End Select
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '������
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '��ť����
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '��ͼ��
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '״̬��
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
    
    Case conMenu_Help_Help              '��������
    
        Call ShowHelp(App.ProductName, frmMain.hwnd, frmMain.Name, Int((glngSys) / 100))
        
    Case conMenu_Help_Web_Home          'Web�ϵ�����
        
        Call zlHomePage(frmMain.hwnd)
        
    Case conMenu_Help_Web_Forum         'Web�ϵ���̳
    
        Call zlWebForum(frmMain.hwnd)
        
    Case conMenu_Help_Web_Mail          '���ͷ���
        
        Call zlMailTo(frmMain.hwnd)
            
    Case conMenu_Help_About             '����
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_File_Exit              '�˳�
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function


Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "") As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strFormat As String
    Dim lngNotPrintCols As Long
    Dim lngPrintCol As Long
    
    If objPrintVsf.Cols = 0 Then Exit Function
    If strNotPrintCol <> "" Then
        lngNotPrintCols = UBound(Split(strNotPrintCol, ",")) + 1
        strNotPrintCol = "," & strNotPrintCol & ","
    End If
    
    objPrintVsf.Rows = objVsf.Rows
    objPrintVsf.FixedRows = objVsf.FixedRows
    
    objPrintVsf.Cols = 0
    lngPrintCol = -1
    For lngCol = 0 To objVsf.Cols - 1
        
        If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
            
            If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                lngPrintCol = lngPrintCol + 1
                
                objPrintVsf.Cols = lngPrintCol + 1
                
                objPrintVsf.ColWidth(lngPrintCol) = objVsf.ColWidth(lngCol)
                objPrintVsf.ColAlignmentFixed(lngPrintCol) = objVsf.ColAlignment(lngCol)
                If objVsf.ColDataType(lngCol) = flexDTBoolean Then
                    objPrintVsf.ColAlignment(lngPrintCol) = 4
                Else
                    objPrintVsf.ColAlignment(lngPrintCol) = objVsf.ColAlignment(lngCol)
                End If
            End If
        End If
    Next
    
    If objPrintVsf.Cols = 0 Then Exit Function
    
    For lngRow = 0 To objVsf.Rows - 1

        objPrintVsf.RowHeight(lngRow) = IIf(objVsf.RowHeight(lngRow) < objVsf.RowHeightMin, objVsf.RowHeightMin, objVsf.RowHeight(lngRow))
        lngPrintCol = -1
        For lngCol = 0 To objVsf.Cols - 1
            
            If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
                If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                    lngPrintCol = lngPrintCol + 1
                    
                    If objVsf.ColDataType(lngCol) = flexDTBoolean And lngRow >= objVsf.FixedRows Then
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = IIf(Abs(Val(objVsf.TextMatrix(lngRow, lngCol))) = 1, "��", "")
                    Else
                        strFormat = objVsf.ColFormat(lngCol)
                        If strFormat = "" Then
                            objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Trim(objVsf.TextMatrix(lngRow, lngCol))
                        Else
                            objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Format(objVsf.TextMatrix(lngRow, lngCol), strFormat)
                        End If
                    End If
                End If
            End If
        Next
        Call SetMsfForeColor(objPrintVsf, lngRow, Val(objVsf.Cell(flexcpForeColor, lngRow, 1)))
    Next
    SearchPrintData = True
End Function

Public Sub SetMsfForeColor(ByRef msf As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intCol As Integer

    With msf

        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .COL = intCol
            .CellForeColor = lngColor
        Next

    End With
End Sub


Public Sub SendLMouseButton(ByVal lngHwnd As Long, ByVal x As Single, ByVal y As Single)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngX As Long
    Dim lngY As Long
    Dim lngLoop As Long
    Dim lngXY As Long
            
    lngX = x / 15
    lngY = y / 15
        
    lngXY = 2
    For lngLoop = 1 To 15
        lngXY = lngXY * 2
    Next
    
    lngXY = lngXY * lngY + lngX
    
    SendMessage lngHwnd, WM_LBUTTONDOWN, 0, ByVal lngXY
    SendMessage lngHwnd, WM_LBUTTONUP, 0, ByVal lngXY

End Sub

Public Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    GetPersonSet = False
    If Val(GetPara("ʹ�ø��Ի����")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ָ������Ϣ������ע�����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strKeyValue-��ֵ
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case enmRegister
    Case ע����Ϣ
        
        Call SaveSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue)
        
    Case ˽��ģ��

        Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case ˽��ȫ��

        Call SaveSetting("ZLSOFT", "˽��ȫ��\" & UserInfo.�û��� & "\" & strSection, strKey, strKeyValue)
        
    Case ����ģ��

        Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case ����ȫ��
        
        Call SaveSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue)
        
    End Select
    
    SetRegister = True
    
errHand:
    
End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '���ܣ� ��ָ����ע����Ϣ��ȡ����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strDefKeyValue-ȱʡ��ֵ
    '���أ� strKeyValue-��ֵ
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case ע����Ϣ
        
        strValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strDefKeyValue)
        
    Case ˽��ģ��

        strValue = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case ˽��ȫ��

        strValue = GetSetting("ZLSOFT", "˽��ȫ��\" & UserInfo.�û��� & "\" & strSection, strKey, strDefKeyValue)
        
    Case ����ģ��

        strValue = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case ����ȫ��
        
        strValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function


'========================================================================================
'=���Q:���(ChkRsState)
'=��ڲ���:Rs               ����:ADODB.Recordset
'=���ڲ���:ChkRsState       ����:Boolean
'=����:����¼����״̬
'=����:2004-07-08
'=����:л��
'========================================================================================
Function ChkRsState(rs As ADODB.Recordset) As Boolean
On Error GoTo ErrH:
    With rs
        If rs Is Nothing Then
            ChkRsState = True
            Exit Function
        Else
            ChkRsState = False
        End If
        If rs.State = 0 Then
            ChkRsState = True
            Exit Function
        Else
            ChkRsState = False
        End If
        If .RecordCount < 1 Then
            ChkRsState = True
        Else
            ChkRsState = False
        End If
        If .EOF Or .BOF Then
            ChkRsState = True
        Else
            ChkRsState = False
        End If
    End With
    Exit Function
ErrH:
    Err.Clear
End Function


'==================================================================================================
'=����:ȥ���ַ����еĵ�����("'")(ConvertString)
'=��ڲ���:
'=1).sStr          ����:String
'=���ڲ���:��
'=����:ȥ���ַ���(sStr)�еĵ�����
'=����:2010-12-11
'=���:л��
'=˵��:��SQL����в��ܴ�������
'==================================================================================================
Public Function ConvertString(ByVal sStr As String) As String
    Dim i               As Integer
    Dim strReturn       As String
    Dim strSystemChar   As String
On Error GoTo ErrH
    strSystemChar = "'|[]"
    '���ϵͳ����¼���ַ�
    For i = 1 To Len(strSystemChar)
        sStr = Replace(sStr, Mid(strSystemChar, i, 1), "")
    Next
    strReturn = sStr
    ConvertString = strReturn
    Exit Function
ErrH:
    Err.Clear
    ConvertString = ""
End Function

'==================================================================================================
'��ⳤ���Ƿ񳬹�����(�ֽ���)
'==================================================================================================
Public Function ChkStrUniCode(mStr As String, mLen As Long) As String
    Dim strL        As String
On Error GoTo ErrH
    mStr = ConvertString(mStr)
    If mLen <= 0 Then
        ChkStrUniCode = mStr
        Exit Function
    Else
        strL = StrConv(mStr, vbFromUnicode)
        strL = LeftB(strL, mLen)
        ChkStrUniCode = StrConv(strL, vbUnicode)
    End If
    Exit Function
ErrH:
    Err.Clear
    ChkStrUniCode = ""
    Exit Function
End Function

'==================================================================================================
'=����:�õ������б���Text������ȡ��ID(Cmb_ID)
'=��ڲ���:
'=1).�����б��ؼ�         ����:Control
'=���ڲ���:��
'=����:�õ������б���Text������ȡ��ID
'=����:2004-12-11
'=���:л��
'=˵��:��ԭ�����ID�е����ݲ��ܴ�"-"
'==================================================================================================
Function Cmb_ID(Combo As Object, Optional Index As Byte = 1) As String
    Dim xx          As Variant
On Error GoTo ErrH
    If Combo.Text = "" Then
        Cmb_ID = ""
    Else
        xx = Split(Combo.Text, gstrSplitCmb)
        If Index - 1 <= UBound(xx) Then '����±�ֵС������ֵ[֤���н�ȡֵ]
            Cmb_ID = xx(Index - 1)
        Else                        '����±�ֵ���ڵ�������ֵ[֤�����޽�ȡֵ]������
            Cmb_ID = "[��]"
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function


'==================================================================================================
'=����:�õ������б���Text������ȡ��ID(Cmb_EditIndex)
'=��ڲ���:
'=1).�����б��ؼ�         ����:Control
'=���ڲ���:��
'=����:�õ������б���Text������ȡ��ID
'=����:2004-12-11
'=���:л��
'=˵��:��ԭ�����ID�е����ݲ��ܴ�"-"
'==================================================================================================
Function Cmb_EditIndex(Combo As Object, sID As String) As Long
    Dim lngCount    As Long
    Dim lngStep     As Long
    Dim xx          As Variant
On Error GoTo ErrH
    lngCount = Combo.ListCount - 1
    For lngStep = 0 To lngCount
        xx = Split(Combo.List(lngStep) & gstrSplitCmb, gstrSplitCmb)
        If sID = xx(0) Then
            Cmb_EditIndex = lngStep
            Exit For
        End If
    Next
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'========================================================================================
'=���ܣ���ѯ�������ݲ���λ
'=��Σ�1��objVsf VSFlexGrid����
'=      2��strFind �����ַ���
'=      3����ѯ�м��������á�,���ָ�
'========================================================================================
Public Sub vsfSetRow(ByRef objVsf As VSFlexGrid, ByVal strFind As String, ByVal strCols As String)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    Dim varCols     As Variant
    Dim strCol      As String
    Dim blnExit     As Boolean
    
    varCols = Split(strCols, ",")
    blnExit = False
    
    '��ȡ���ڵ�ǰ�еļ�¼����
    For lngLoop = objVsf.Row + 1 To objVsf.Rows - 1
        For intCol = 0 To UBound(varCols)
            strCol = varCols(intCol)
            If InStr(UCase(objVsf.TextMatrix(lngLoop, objVsf.ColIndex(strCol))), UCase(strFind)) > 0 Then
                lngRow = lngLoop
                blnExit = True
                Exit For
            End If
        Next
        If blnExit Then Exit For
    Next
    
    '��ȡС�ڵ�ǰ�еļ�¼����
    If lngRow = 0 Then
        For lngLoop = 0 To objVsf.Row
            For intCol = 0 To UBound(varCols)
                strCol = varCols(intCol)
                If InStr(UCase(objVsf.TextMatrix(lngLoop, objVsf.ColIndex(strCol))), UCase(strFind)) > 0 Then
                    lngRow = lngLoop
                    blnExit = True
                    Exit For
                End If
            Next
            If blnExit Then Exit For
        Next
    End If
    If objVsf.Rows > 1 And lngRow >= 1 Then objVsf.Row = lngRow
    DoEvents
    objVsf.ShowCell lngRow, 1
End Sub





