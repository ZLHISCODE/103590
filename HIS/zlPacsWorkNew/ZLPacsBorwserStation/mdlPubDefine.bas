Attribute VB_Name = "mdlPubDefine"
Option Explicit
'�������ݲ˵�ID����:*��ʾ��ͼ��
'*********************************************************************
'���������������������ҵ�������ɣ�������������еĹ��ܣ�Ӧ��"����"�˵���չ���������ҵ�����Ĺ��ܣ�Ӧ��"�༭"�˵���չ��
Public Const conMenu_FilePopup = 1 '�ļ�
Public Const conMenu_ManagePopup = 2 '����
Public Const conMenu_EditPopup = 3 '�༭
Public Const conMenu_ReportPopup = 4 '����
Public Const conMenu_PlugPopup = 6 '��ӳ��򣻼��鼼ʦ����վʹ��6100-6199
Public Const conMenu_ViewPopup = 7 '�鿴
Public Const conMenu_ToolPopup = 8 '����
Public Const conMenu_HelpPopup = 9 '����


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

'���˷��ò�ѯ
Public Const conMenu_File_PrintMultiBill = 110 '��ӡ���Ŵ߿
Public Const conMenu_File_PrintSingleBill = 732 '��ӡ���Ŵ߿
Public Const conMenu_File_PrintDayDetail = 3554 '��ӡһ���嵥
Public Const conMenu_File_PrintBedCard = 3555 '��ӡ��ͷ��
Public Const conMenu_File_PrintPageSet = 113 '��ӡ��ҳ����

'����˵�:����վ����Ĺ��ܲ˵�
Public Const conMenu_Manage_Monitor = 201     '*�໤��

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

Public Const conMenu_Manage_Call = 218            '����
Public Const conMenu_Manage_CallNext = 21801        '��һ��(&N)
Public Const conMenu_Manage_CallPrevious = 21802    '��һ��(&P)

Public Const conMenu_Manage_Reset = 219     '����˳��
Public Const conMenu_Manage_Up = 21901        '����(&U)
Public Const conMenu_Manage_Down = 21902      '����(&D)
Public Const conMenu_Manage_Discard = 21903   '����(&D)
Public Const conMenu_Manage_Recall = 21904    '�ٻ�(&R)
Public Const conMenu_Manage_Untread = 21905   '�˺�(&R)
Public Const conMenu_Manage_TagEnd = 21906  '���Ϊ����(&M)
Public Const conMenu_Manage_ShowAller = 220  '��ʾ���й�����¼

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
Public Const conMenu_Manage_ReportLisView = 2325     '����LIS����(&L)
Public Const conMenu_Manage_LeaveMedi = 251 '�Ĵ�ҩƷ

Public Const conMenu_Manage_Audit = 252         '*�������
Public Const conMenu_Manage_UnAudit = 253       '*ȡ�����
Public Const conMenu_Manage_Arrange = 254       '*ִ�а���
Public Const conMenu_Manage_UnArrange = 255     '*ȡ������

'��ʿվ�������ת
Public Const conMenu_Manage_Change_In = 2600          '�������
Public Const conMenu_Manage_Change_Turn = 2601      'ת��
Public Const conMenu_Manage_Change_Bed = 2602         '����
Public Const conMenu_Manage_Change_House = 2603       '����
Public Const conMenu_Manage_Change_Out = 2604         '���˳�Ժ
Public Const conMenu_Manage_Change_InPati = 2605      'תΪסԺ����
Public Const conMenu_Manage_Change_BedGrid = 2606     '���Ĵ�λ�ȼ�
Public Const conMenu_Manage_Change_PatiInfo = 2607    '����סԺ��Ϣ
Public Const conMenu_Manage_Change_Baby = 2608        '�������Ǽ�
Public Const conMenu_Manage_Change_ReCalcFee = 2609   '���ѱ��������
Public Const conMenu_Manage_Change_InsureSel = 2610   'ҽ������ѡ��
Public Const conMenu_Manage_Change_Undo = 2611         '��������
Public Const conMenu_Manage_Print_Label = 2612         '��ӡ���

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
Public Const conMenu_Edit_MediAudit = 3564 '*ҩ�����(&U)(������ҩ���)
Public Const conMenu_Edit_UnUse = 3021      '*���δ��(&H)
Public Const conMenu_Edit_Surplus = 3022      '*����Ǽ�

'�����¼��ʹ�õ��ĸ���,ճ���Լ������������
Public Const conMenu_Edit_Copy = 3031      '*����
Public Const conMenu_Edit_PASTE = 3032      '*ճ��
Public Const conMenu_Edit_SPECIALCHAR = 3033      '*�����������
Public Const conMenu_Edit_Clear = 3034      '*�������

'����(�༭)�˵�
Public Const conMenu_Edit_NewParent = 3051   '*�·���(&N)
Public Const conMenu_Edit_Insert = 3052      '*����(&I)
Public Const conMenu_Edit_ModifyParent = 3053 '*�޸ķ���(&M)
Public Const conMenu_Edit_DeleteParent = 3054 '*ɾ������(&D)
Public Const conMenu_Edit_MarkMap = 3061     '*ͼƬ(&I)��
Public Const conMenu_Edit_ApplyTo = 3062     '*���ÿ���(&T)
Public Const conMenu_Edit_Request = 3063     '����Ҫ��(&R)
Public Const conMenu_Edit_Compend = 3064     '*���ݹ���(&F)
Public Const conMenu_Edit_ElementChange = 3065      '*Ҫ����������
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
Public Const conMenu_Edit_Option = 3097      '����ѡ��

'Public Const conMenu_Manage_ThingAdd = 227    '�ӵ�(&A)
'Public Const conMenu_Manage_ThingModi = 228   '*����ִ�����(&M)
Public Const conMenu_Edit_Transf_Delete = 229   '�����ӵ�

'���˷��ò�ѯ
'----------------------------------------------------------------------
Public Const conMenu_Edit_PreBalance = 817 'Ԥ�ᵱǰ����
Public Const conMenu_Edit_PreBalanceAll = 818 'Ԥ�����в���
Public Const conMenu_Edit_Balance = 3011 '����
Public Const conMenu_Edit_Billing = 3003 '����
Public Const conMenu_Edit_ReBilling = 3004 'ֱ������

Public Const conMenu_Edit_ReBillingButton = 3017       '*���ó���(&E)
Public Const conMenu_Edit_ReBillingApply = 3019 '*��������(&L)
Public Const conMenu_Edit_ReBillingAudit = 3020 '*�������(&U)

Public Const conMenu_Edit_FeeAudit = 3564 '���
Public Const conMenu_Edit_FeeUnAudit = 3565 'ȡ�����

'���ѿ�����
'---------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_CardPay = 3811 '����
Public Const conMenu_Edit_CardBathPay = 3812 '��������
Public Const conMenu_Edit_CardBack = 3813 '�˿�
Public Const conMenu_Edit_CardCancelBack = 38131 'ȡ������
Public Const conMenu_Edit_CardCallBack = 3814 '����
Public Const conMenu_Edit_CardCancelCallBack = 38141 'ȡ������

Public Const conMenu_Edit_CardInFull = 3816 '��ֵ
Public Const conMenu_Edit_CardInFullBack = 3817 '��ֵ����
Public Const conMenu_Edit_CardModify = 3818 '�޸Ŀ���Ϣ
Public Const conMenu_Edit_CardResume = 3819 '������
Public Const conMenu_Edit_CardStop = 38191 '��ͣ��
Public Const conMenu_Edit_MoveCard = 3821 '����ʱ���Ƴ���Ƭ
Public Const conMenu_Apply_AllCard = 3822 '����ʱ�����ݵ�ǰ���ݣ�Ӧ����������Ҫ�����ĵ���
Public Const conMenu_Apply_AllColumn = 3823 '����ʱ�����ݵ�ǰ����ָ�����У�Ӧ����������Ҫ�����Ĵ�����Ϣ
Public Const conMenu_COMBOX_INTERFACE = 3820 '���ѿ��ӿ�
Public Const conMenu_Square_BrushCard = 3824 '���ѿ�Ŀ¼+�ӿ����


'�������
'----------------------------------------------------------------------------------
Public Const conMenu_Edit_Triage = 2604  '����
Public Const conMenu_Edit_ModiyPati = 2607  '����������Ϣ
Public Const conmenu_Edit_ChangeNum = 3088 '���
Public Const conmenu_Edit_Leave = 3089 '���˲�����
Public Const conmenu_Edit_Wait = 3090 '���˴���
Public Const conmenu_View_TriagePati = 7101 '��ʾ�ѷ��ﲡ��
Public Const conmenu_View_AdmissionsPati = 7102 '��ʾ�ѽ��ﲡ��
Public Const conmenu_View_OverPati = 7103 '��ʾ����ɲ���
Public Const conmenu_View_Leave = 7104 '��ʾ�����ﲡ��
Public Const conmenu_View_AutoRefresh = 7120 '�Զ�ˢ��

'�ҺŰ���
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_PlanAdd = 6820 '�ƻ�����
Public Const conMenu_Edit_PlanModify = 6821 '�޸ļƻ�����
Public Const conMenu_Edit_PlanDelete = 6822 'ɾ���ƻ�����
Public Const conMenu_Edit_PlanVerify = 6823 '��˼ƻ�����
Public Const conMenu_Edit_PlanCancel = 6824 'ȡ����˼ƻ�
Public Const conMenu_Edit_AllStartNO = 6825  'ȫ�����ùҺ���ſ���
Public Const conMenu_Edit_AllStopNO = 6826 'ȫ��ͣ�ùҺ���ſ���
Public Const conMenu_Edit_StopPlanTimes = 6827  'ͣ�ð��żƻ�
Public Const conMenu_Edit_ClearStopPlan = 6828  '�������ͣ�ð��żƻ�



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

'Ժ�б༭\�鿴�˵�
Public Const conMenu_Edit_DelDayItem = 3802        'ɾ���ձ���ǰ����Ϣ
Public Const conMenu_Edit_BuildConstant = 3803        '���ɳ���������Ŀ
 
'��������\�鿴�˵�
Public Const conMenu_Edit_CfPay = 4000        '����������
Public Const conMenu_Edit_BillPay = 4001        '��Ʊ�ݷ���
Public Const conMenu_Edit_BillBackPay = 4002        '����������
Public Const conMenu_Edit_StopPay = 4003        '��ֹͣ���ϱ��

Public Const conMenu_View_FontSize = 4004         '�ֺ�����
Public Const conMenu_View_FontSize_1 = 4004         '9����
Public Const conMenu_View_FontSize_2 = 4004         '11����
Public Const conMenu_View_FontSize_3 = 4004         '15����



'LISʹ�õĲɵ� 3650-3690
Public Const conMenu_Edit_QCRes = 3650         '*�ʿ�Ʒ
Public Const conMenu_LIS_Cancel = 3651         '*ȡ��
Public Const conMenu_LIS_PatientInfo = 3652    '������Ϣ
Public Const conMenu_LIS_HideList = 3653       '���ز����б�
Public Const conMenu_LIS_TOQC = 3654           '��Ϊ�ʿ�
Public Const conMenu_LIS_SendReport = 3655     '���ͱ��浥
Public Const conMenu_LIS_SignVerify = 3656     '��֤ǩ��
Public Const conMenu_LIS_MB_Connect = 3701     'ø��������
Public Const conMenu_LIS_MB_Disconnect = 3702  'ø���ǶϿ�
Public Const comMenu_LIS_TodayQC = 3703        '�����ʿ�
Public Const comMenu_LIS_History = 3704        '��ʷ�ʿ�
Public Const comMenu_LIS_ShowListHead = 3705   'ѡ��Ҫ��ʾ����
Public Const conMenu_LIS_LJAverage = 3706      '��ֵLJ�ʿ�
Public Const conMenu_LIS_RightMenu = 3707      '�Ҽ��˵�

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
Public Const conMenu_View_FontSize_S = 4041            'ҽ�����壺С����
Public Const conMenu_View_FontSize_L = 4042            'ҽ�����壺������

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
Public Const conMenu_View_PatiInput = 726             '��ʾ������Ϣ����������
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

Public Const conMenu_View_Navigatebeginning = 7401           '*��һ��(&F)
Public Const conMenu_View_Navigateleft = 7402                '*��һ��(&F)
Public Const conMenu_View_Navigateright = 7403               '*��һ��(&F)
Public Const conMenu_View_Navigateend = 7404                 '*���һ��(&F)

'���˷��ò�ѯ
Public Const conMenu_View_Billing = 3551             '�鿴���ʵ�
Public Const conMenu_View_DateType = 781           '��ѯʱ��
Public Const conMenu_View_DetailType = 793          '�嵥����
Public Const conMenu_View_GroupCol = 733            '�����ֶ�

Public Const conMenu_View_ReBalance = 7510  '��ʾ��������
Public Const conMenu_View_ZeroFee = 7511    '��ʾ�����
Public Const conMenu_View_CheckFee = 7512   '��ʾ������
Public Const conMenu_View_Owe = 7513        '��ʾδ���岡��
Public Const conMenu_View_UnAudit = 7514     '��ʾδ��˲���
Public Const conMenu_View_OnePati = 7515    '���סԺֻ��һ�β���

'���ϵͳ����70��ͷ�ĺ�
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_View_Single = 7040             '����
Public Const conMenu_View_Group = 7041              '����
Public Const conMenu_View_LocationMethod = 7042     '��λ����
Public Const conMenu_View_Column = 7043             'ѡ������

Public Const conMenu_View_LocationRange = 7044     '��λ��Χ

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
Public Const conMenu_Tool_Community = 805       '*��������(&U)
Public Const conMenu_Tool_MedRecAudit = 806        '�������(&M)
Public Const conMenu_Tool_MedRecAuditSubmit = 8061      '�ύ���(&S)
Public Const conMenu_Tool_MedRecAuditCancel = 8062      'ȡ���ύ(&C)
Public Const conMenu_Tool_MedRecAuditResponse = 8063    '��鷴��(&M)
Public Const conMenu_Tool_Archive = 807         '*��Ա����(&I)
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

Public Const conMenu_Tool_PlugIn = 890          '���
Public Const conMenu_Tool_PlugIn_Item = 89000   '�����,ʵ������Ϊ conMenu_Tool_PlugIn_Item + n, 1<=n<=99

'PACS����վ�˵�
Public Const conMenu_Manage_Result = 8300       '�����
Public Const conMenu_Manage_Negative = 8301      '���������
Public Const conMenu_Manage_Positive = 8302      '���������
Public Const conMenu_Manage_Quality = 8303       'Ӱ������
Public Const conMenu_Manage_First = 8304      '�׼�
Public Const conMenu_Manage_Second = 8305      '�Ҽ�
Public Const conMenu_Manage_ChangeUser = 8306      '�����û�
Public Const conMenu_Manage_ChangeDevice = 8307     '�����豸Ӱ�����
Public Const conMenu_Manage_ImageInterval = 8308    '��ͼ����
Public Const conMenu_Manage_CopyCheck = 8200        'ͬһ���˵Ǽ���ͬ��Ŀ��ͬ��λ
Public Const conMenu_Manage_GChannel = 8201         '��ɫͨ��
Public Const conMenu_Manage_GChannelOk = 8202       '��ɫͨ�����
Public Const conMenu_Manage_GChannelCancel = 8203   '��ɫͨ��ȡ��
Public Const conMenu_Manage_Review = 8204           '���
Public Const conMenu_Manage_SelectAllImages = 8205      'ȫѡͼ��
Public Const conMenu_Manage_UnSelectAllImages = 8206    'ȫ��ͼ��
Public Const conMenu_Manage_ReverseSelectImages = 8207  '��ѡͼ��
Public Const conMenu_Manage_TechDoctorExecute = 8208    '��ʦִ��
Public Const conMenu_Manage_ReportRelease = 8209        '���淢��
Public Const conMenu_Manage_RelatingPatiet = 8210       '��������
Public Const conMenu_Manage_LocateType = 8211           '��λ��ʽ
Public Const conMenu_Manage_LocateValue = 8212          '��λֵ

'PACS����༭��
Public Const conMenu_PacsReport_SelFormat = 8309    'ѡ�񱨸��ʽ
Public Const conMenu_PacsReport_SelFormat_Item = 8310    'ѡ�񱨸��ʽ
Public Const conMenu_PacsReport_Save = 8311         '���汨��
Public Const conMenu_PacsReport_Sign = 8312         '����ǩ��
Public Const conMenu_PacsReport_DelSign = 8313      '����ǩ��
Public Const conMenu_PacsReport_MoveUp = 8314       'ͼ��ǰ��
Public Const conMenu_PacsReport_MoveDown = 8315     'ͼ�����
Public Const conMenu_PacsReport_DelImage = 8316     'ɾ��ͼ��
Public Const conMenu_PacsReport_DelMarks = 8317     '�����ע
Public Const conMenu_PacsReport_Open = 8318         '�򿪱���༭����
Public Const conMenu_PacsReport_FontSet = 8319      '���ô��ı�������
Public Const conMenu_PacsReport_History = 8320      '�޶���ʷ
Public Const conMenu_PacsReport_Mode_Orig = 8321    'ԭʼ״̬
Public Const conMenu_PacsReport_Mode_Clear = 8322   '����״̬
Public Const conMenu_PacsReport_History_Times = 8323 '��ʷ����
Public Const conMenu_PacsReport_DelMiniImage = 8324     'ɾ����������ͼ
Public Const conMenu_PacsReport_SelMiniImage = 8325     '��ȡ��������ͼ
Public Const conMenu_PacsReport_RptImg2CapImg = 8326    '�ڱ���ͼ���ߺͲɼ����߼��л�
Public Const conMenu_PacsReport_PrivOrder = 8327        '��һ��ҽ��
Public Const conMenu_PacsReport_NextOrder = 8328        '��һ��ҽ��
Public Const conMenu_PacsReport_AddNumber = 8329        '����������������
Public Const conMenu_PacsReport_RepFormat = 8330        '�Զ��屨��ѡ���ʽ
Public Const conMenu_PacsReport_RepFormat_Item = 8331   '�Զ��屨��ѡ��ľ����ʽ��
Public Const conMenu_PacsReport_SaveWord = 8332         '����ʾ�ʾ��

'�ɼ��˵�
Public Const conMenu_Cap_Dynamic = 8100         '��̬��ʾ(&V)
Public Const conMenu_Cap_MarkMap = 8101       'Ӱ��ɼ�(&C)
Public Const conMenu_Cap_Import = 8102        'Ӱ����(&I)
Public Const conMenu_Cap_DevSet = 8103          'Ӱ���豸����(&D)
Public Const comMenu_Cap_Process = 8104         'Ӱ����
Public Const conMenu_Cap_Record = 8105          '¼��(&R)
Public Const conMenu_Cap_Record_Stop = 8099     'ֹͣ¼��(&O)
Public Const conMenu_Cap_Full_Screen = 8098     'ȫ��(&U)
Public Const conMenu_Cap_Play = 8106          '����(&P)
Public Const conMenu_Cap_Stop = 8107            'ֹͣ(&T)
Public Const conMenu_Cap_Forward = 8108         '���(&F)
Public Const conMenu_Cap_Back = 8109            '����(&B)
Public Const conMenu_Cap_SaveAs = 8110          '����¼��(&S)
Public Const conMenu_Cap_OpenStudyList = 8122   '�򿪼���б�
Public Const conMenu_Cap_StudySyncState = 8123  'Ӱ����ͬ��״̬


Public Const conMenu_Img_Look = 8111        'Ӱ���Ƭ(&S)
Public Const conMenu_Img_Contrast = 8112    '��Ƭ�Ա�(&E)
Public Const conMenu_Img_Delete = 8113        'ͼ��ɾ��(&K)
Public Const conMenu_Img_Query = 8114        'Q/R��ȡͼ��(&Q)

'��ά�ؽ��˵�
Public Const conMenu_Img_3D = 8115          '��ά�ؽ�
Public Const conMenu_Img_3D_VA = 8116       '�ݻ��ؽ�
Public Const conMenu_Img_3D_MPR = 8117      'MPR
Public Const conMenu_Img_3D_MMPR = 8118     'MMPR
Public Const conMenu_Img_3D_VE = 8119       '�����ڿ���
Public Const conMenu_Img_3D_SA = 8120       '�����ؽ�
Public Const conMenu_Img_3D_PF = 8121       '��ע����

'�Ŷӽк�ϵͳ
Public Const conMenu_Queue_CallThis = 8250      'ֱ��
Public Const conMenu_Queue_CallNext = 8251      '˳����������һ��
Public Const conMenu_Queue_CallFirst = 8252     '����
Public Const conMenu_Queue_ReInQueue = 8253     '����
Public Const conMenu_Queue_ReCall = 8254        '�غ�
Public Const conMenu_Queue_Abandon = 8255       '����
Public Const conMenu_Queue_Refresh = 8256       'ˢ��
Public Const conMenu_Queue_Setup = 8257         '��������
Public Const conMenu_Queue_Update = 8258        '�޸�
Public Const conMenu_Queue_Broadcast = 8259     '�㲥
Public Const conMenu_Queue_Pause = 8260         '��ͣ
Public Const conMenu_Queue_Finaled = 8261       '��ɾ���
Public Const conMenu_Queue_Find = 8262          '����


'ͼ����
Public Const conMenu_Process_Window = 501           '���ȶԱȶ�
Public Const conMenu_Process_Zoom = 502             '����
Public Const conMenu_Process_Corp = 512             '�϶�
Public Const conMenu_Process_RRotate = 503          '˳ʱ����ת
Public Const conMenu_Process_LRotate = 504          '��ʱ����ת
Public Const conMenu_Process_Sharpness = 505        '��
Public Const conMenu_Process_Filter = 506           'ƽ��
Public Const conMenu_Process_Arrow = 507            '��ͷ��ע
Public Const conMenu_Process_Ellipse = 508          'Բ�α�ע
Public Const conMenu_Process_Text = 509             '���ֱ�ע
Public Const conMenu_Process_RectZoom = 510         '�ü��ɼ�
Public Const conMenu_Process_RectCapture = 511         '�ü���ɼ�
Public Const conMenu_Process_Restore = 8124         '�ָ�


'�����˵�
Public Const conMenu_Help_Help = 901        '*��������(&H)
Public Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_Help_Web_Forum = 9023      '������̳(&F)
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

