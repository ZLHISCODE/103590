Attribute VB_Name = "mdlPubDefine"
Option Explicit
'�������ݲ˵�ID����:*��ʾ��ͼ��
'*********************************************************************
'����������������������ҵ�������ɣ��������������еĹ��ܣ�Ӧ��"����"�˵���չ���������ҵ�����Ĺ��ܣ�Ӧ��"�༭"�˵���չ��
Public Const conMenu_FilePopup = 1    '�ļ�
Public Const conMenu_ManagePopup = 2    '����
Public Const conMenu_EditPopup = 3    '�༭
Public Const conMenu_ReportPopup = 4    '����
Public Const conMenu_PlugPopup = 6    '��ӳ��򣻼��鼼ʦ����վʹ��6100-6199
Public Const conMenu_ViewPopup = 7    '�鿴
Public Const conMenu_ToolPopup = 8    '����
Public Const conMenu_HelpPopup = 9    '����
Public Const conMenu_PlugIn = 10        '�����չ
Public Const conMenu_PlugIn_Menu = 1000001      '�����չ


'�ļ��˵�
Public Const conMenu_File_Open = 100            '*��(&O)��
Public Const conMenu_File_PrintSet = 101        '*��ӡ����(&S)��
Public Const conMenu_File_Preview = 102         '*Ԥ��(&V)
Public Const conMenu_File_Preview_Pati = 10201         '*Ԥ�������б�(&V)    '���ò�ѯ��Ч
Public Const conMenu_File_Print = 103           '*��ӡ(&P)
Public Const conMenu_File_NoAskPrint = 106      '��Ĭ��ӡ
Public Const conMenu_File_Print_Pati = 10301         '*��ӡ�����б�(&P) '���ò�ѯ��Ч
Public Const conMenu_File_Print_PatiPath = 10302         '*��ӡ���˰��ٴ�·����
Public Const conMenu_File_Excel = 104           '�����&Excel��
Public Const conMenu_File_Excel_Pati = 10401         '*�����б������&Excel�� '���ò�ѯ��Ч
Public Const conMenu_File_MedRec = 105          '��ҳ��ӡ(&R)
Public Const conMenu_File_MedRecSetup = 1051        '��ӡ����(&S)
Public Const conMenu_File_MedRecPreview = 1052      '��ӡԤ��(&P)
Public Const conMenu_File_MedRecPrint = 1053        '��ӡ��ҳ(&V)
Public Const conMenu_File_RowPrint = 121        '��¼��ӡ(&R)
Public Const conMenu_File_BatPrint = 122        '������ӡ(&B)
Public Const conMenu_File_Print_Bespeak = 123        'ԤԼ�Һŵ���ӡ
Public Const conMenu_File_Parameter = 181       '*��������(&M)
Public Const conMenu_File_RoomSet = 182         'ִ�м��豸
Public Const ConMenu_File_ShortcutSet = 183     '��ݹ�������
Public Const conMenu_File_SendImg = 184         '����ͼ��
Public Const conMenu_File_SaveJpeg = 185       '����ΪͼƬ
Public Const conMenu_File_Exit = 191            '*�˳�(&X)
Public Const conMenu_File_ExportToXML = 192     '����ΪXML�ĵ�
Public Const conMenu_File_ImportFromXML = 193   '��XML�ĵ�����
Public Const conMenu_File_BillPrintView = 194   '���ݴ�ӡԤ��
Public Const conMenu_File_BillPrint = 195       '���ݴ�ӡ
Public Const conMenu_File_ExportAll = 196       '��������
Public Const conMenu_File_ExportToXMLs = 197     '��������ΪXML�ĵ�
Public Const conMenu_File_ImportFromXMLs = 198   '��XML�ĵ���������
Public Const conMenu_File_BarcodePrint = 199   '�����ӡ
Public Const conMenu_File_BillPrintSet = 200    'Ʊ�ݴ�ӡ����

'���˷��ò�ѯ
Public Const conMenu_File_PrintMultiBill = 110    '��ӡ���Ŵ߿
Public Const conMenu_File_PrintSingleBill = 732    '��ӡ���Ŵ߿
Public Const conMenu_File_SchemeSet = 737    '������������
Public Const conMenu_File_PrintDayDetail = 3554    '��ӡһ���嵥
Public Const conMenu_File_PrintBedCard = 3555    '��ӡ��ͷ��
Public Const conMenu_File_PrintPageSet = 113    '��ӡ��ҳ����

'�����˵�:����վ�����Ĺ��ܲ˵�
Public Const conMenu_Manage_Monitor = 201     '*�໤��
Public Const conMenu_Manage_FeeItemSet = 3065      '������Ŀ��������

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
Public Const conMenu_Manage_ReBack = 2171        '����
Public Const conMenu_Manage_ReBackCancel = 2172        'ȡ������

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

Public Const conMenu_Manage_InQueue = 230     '�Ŷӽк����
Public Const conMenu_Manage_Plan = 221        '*ִ�б���(&P)
Public Const conMenu_Manage_Logout = 222      'ȡ������(&L)
Public Const conMenu_Manage_Refuse = 223      '�ܾ�ִ��(&R)
Public Const conMenu_Manage_ReGet = 224       'ȡ���ܾ�(&G)
Public Const conMenu_Manage_Complete = 225    '*ִ�����(&C)
Public Const conMenu_Manage_Undone = 226      'ȡ�����(&U)
Public Const conMenu_Manage_ThingAdd = 227    '*��¼ִ�����(&A)
Public Const conMenu_Manage_ThingModi = 228   '*����ִ�����(&M)
Public Const conMenu_Manage_ThingDel = 229    '*ɾ��ִ�����(&D)
'Public Const conMenu_Manage_ModifBaseInfo = 230 '������Ϣ����
Public Const conMenu_Manage_ThingAudit = 234    '*�˶�ִ�����(&E)
Public Const conMenu_Manage_ThingDelAudit = 235    '*ȡ���˶�ִ��(&F)
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
Public Const conMenu_Manage_RecipeAuditView = 2326
Public Const conMenu_Manage_ReportPacsView = 2330    '����Pacs����(&P)
Public Const conMenu_Manage_LeaveMedi = 251    '�Ĵ�ҩƷ

Public Const conMenu_Manage_Audit = 252         '*�������
Public Const conMenu_Manage_UnAudit = 253       '*ȡ�����
Public Const conMenu_Manage_Arrange = 254       '*ִ�а���
Public Const conMenu_Manage_UnArrange = 255     '*ȡ������

Public Const ConMenu_pop_Dept = 303           '�������
Public Const ConMenu_pop_DeptDistrict = 304   '���벡��

'��ʿվ�������ת
Public Const conMenu_Manage_Change_In = 2600          '������ס
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
Public Const conMenu_Manage_Change_TurnUnit = 2613      'ת����
Public Const conMenu_Manage_Change_TurnTeam = 2614      'תҽ��С��
Public Const conMenu_Manage_Change_TransposeBed = 2615         '��λ�Ի�
Public Const conMenu_Manage_Change_PaitNote = 2616         '���˱�ע��Ϣ

'ҽ��(�༭)�˵�����϶�,����ʱ��4λ���,50λ�ֶ�,001-050,051-100,101-150,...
Public Const conMenu_Edit_NewItem = 3001    '*����Ŀ(&A)
Public Const conMenu_Edit_NewItemQAdvice = 300101    '*�л���ҽ��
Public Const conMenu_Edit_NewItemQEpr = 300102    '*�л�������
Public Const conMenu_Edit_NewRis = 300103     'RIS�˵���ť
Public Const conMenu_Edit_NewRisSch = 300104   '*ԤԼ
Public Const conMenu_Edit_NewRisDel = 300105   '*ȡ��ԤԼ
Public Const conMenu_Edit_NewRisModi = 300106   '*����ԤԼ
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
Public Const conMenu_Edit_StopAudit = 3050      'ͣ�����(&W)
Public Const conMenu_Edit_BatExecute = 3098       '*ҽ������ִ��
Public Const conMenu_Edit_SendBilling = 30133   '���ͼ��ʵ�
Public Const conMenu_Edit_SendCharge = 30134   '�����շѵ�

Public Const conMenu_Edit_SendDrug = 30131      '*ҩ��ҽ������(&1)
Public Const conMenu_Edit_SendOther = 30132     '����ҽ������(&2)
Public Const conMenu_Edit_SendInfusion = 30135 '*������ҺҩƷҽ��
Public Const conMenu_Edit_AdvicePrice = 30136      '*��ҽ������
Public Const conMenu_Edit_AdvicePay = 30137      '*���֧��
Public Const conMenu_Edit_AdviceBalance = 30138      '*������
Public Const conMenu_Edit_Untread = 5019    '*����(&R)
Public Const conMenu_Edit_SendBack = 3015   '*���ڷ����ջ�(&Z)
Public Const conMenu_Edit_Test = 990056       '*Ƥ�Խ��(&T)
Public Const conMenu_Edit_ChargeOff = 3017       '*���ó���(&E)
Public Const conMenu_Edit_NoPrint = 3018    '���δ�ӡ(&I)
Public Const conMenu_Edit_ChargeDelApply = 3019    '*��������(&L)
Public Const conMenu_Edit_ChargeDelAudit = 3020    '*�������(&U)
Public Const conMenu_Edit_UnUse = 3021      '*���δ��(&H)
Public Const conMenu_Edit_Surplus = 3022      '*����Ǽ�
Public Const conMenu_Edit_MediAudit = 3023          '*ҩ�����(&U)(������ҩ���)
Public Const conMenu_Edit_ExtraFeeMove = 3024       '*����ת��
Public Const conMenu_Edit_ExtraFeeExe = 3025        '*����ִ��
Public Const conMenu_Edit_ExtraFeeUnExe = 3026      '*����ȡ��ִ��
Public Const conMenu_Edit_AdviceUnAudit = 3027      '*����ҩ��ȡ�����
Public Const conMenu_Edit_LISApply = 3028      '*��������
Public Const conMenu_Edit_LISApplyModi = 3029      '*�޸ļ�������
Public Const conMenu_Edit_LISApplyDel = 3030      '*ȡ����������
Public Const conMenu_Edit_CriticalAdvice = 3031 'Σ��ֵҽ��
Public Const conMenu_Edit_ViewDrugExplain = 30821 '�鿴ҩƷ˵����
Public Const conMenu_Edit_Refcom = 30822          '�ܾ��������
Public Const conMenu_Edit_ViewRefcom = 30823      '�����󷽽��
Public Const conMenu_Edit_DrugAuto = 30824        '��ȡ��ҩ�嵥
Public Const conMenu_Edit_DrugGrp = 30825         '��ҩ�嵥һ����ҩ
Public Const conMenu_Edit_DrugList = 30826        '������ҩ�嵥




'�ٴ�·������(�༭)
Public Const conMenu_Edit_Report = 816          '*�����ǼǱ�(&P)

'������¼��ʹ�õ��ĸ���,ճ���Լ������������
Public Const conMenu_Edit_Copy = 3031      '*����
Public Const conMenu_Edit_PASTE = 3032      '*ճ��
Public Const conMenu_Edit_SPECIALCHAR = 3033      '*�����������
Public Const conMenu_Edit_Clear = 3034      '*�������
Public Const conMenu_Edit_PrevPage = 3035   '*��һҳ
Public Const conMenu_Edit_NextPage = 3036   '*��һҳ
Public Const conMenu_Edit_Word = 3037       '*�ʾ�ѡ��
Public Const conMenu_Edit_Brief = 3038      '*С��
Public Const conMenu_Edit_Group_New = 3039          '*�·���
Public Const conMenu_Edit_Group_Append = 3040       '*׷�ӷ���
Public Const conMenu_Edit_Curve = 3041       '*�������߱༭
Public Const conMenu_Edit_CurveTable = 3042       '*���±���༭
Public Const conMenu_Edit_Curve_Show = 3043       '*����������ʾ

'�����ز˵�
Public Const conMenu_Edit_CollectFees = 3588    '�������տ�
Public Const conMenu_Edit_CollectFees_Cancel = 3589  '�������տ�����
Public Const conMenu_Edit_ReprintReceipt = 3591     '�ش��վ�
Public Const conMenu_Edit_RollingCurtain = 3511    '*����(&Z)
Public Const conMenu_Edit_RollingCurtain_Cancel = 3512    '*��������(&D)
Public Const conMenu_Edit_CheckCash = 3513    '*�ֽ�㳮(&E)
Public Const conMenu_Edit_ChargeBook_Reprint = 3514   '*�ش�ɿ���(&E)
Public Const conMenu_Edit_RollingCurtain_Zero = 3515   '*���ʹ���(&C)
Public Const conMenu_Edit_StandbyMoeny_PutOut = 3516   '*���ű��ý�(&L)
Public Const conMenu_Edit_StandbyMoeny_OnWork = 3522   '*�����ϸڱ��ý�
Public Const conMenu_Edit_Personnel_Group = 3517   '*��Ա����(&F)
Public Const conMenu_Edit_StandbyMoeny_PutIn = 3518   '*�ջر��ý�(&H)
Public Const conMenu_Edit_Collect_Other = 3519  '������Ա�տ�
Public Const conMenu_Edit_Collect_Manual = 3520   '�ֹ��տ�(&M)
Public Const conMenu_Edit_Collect_RollingCurtain = 3588   '*�����տ�(&S)
Public Const conMenu_Edit_Collect_Cancel = 3589   '*�տ�����(&S)
Public Const conMenu_Edit_DrawBook_Reprint = 3521   '*�ش����õ�


'�������ֹ��������б�
Public Const conMenu_Edit_WaveSynchro = 3044    '*������Ŀͬ������
Public Const conMenu_Edit_CollectMan = 3045     '*������Ŀ����
Public Const conMenu_Edit_AnimalPart = 3046     '*���²�λ����
Public Const conMenu_Edit_FileMan = 3047        '*�����ļ�����
Public Const conMenu_Edit_WavyMan = 3048        '������Ŀ����
Public Const conMenu_Edit_MOBILE = 3049         '�ƶ���������:1)��Һ������Ŀ����(����Һ������\Һ������Ӧ����Ŀ,�����Զ����������ݸ�����);2)���±�����;3)������Ŀ�������


'����(�༭)�˵�
Public Const conMenu_Edit_NewParent = 3051   '*�·���(&N)
Public Const conMenu_Edit_Insert = 3052      '*����(&I)
Public Const conMenu_Edit_ModifyParent = 3053    '*�޸ķ���(&M)
Public Const conMenu_Edit_DeleteParent = 3054    '*ɾ������(&D)
Public Const conMenu_Edit_MarkMap = 3061     '*ͼƬ(&I)��
Public Const conMenu_Edit_MarkKeyMap = 3070     '�ؼ�ͼ��
Public Const conMenu_Edit_ApplyTo = 3062     '*���ÿ���(&T)
Public Const conMenu_Edit_Request = 3063     '����Ҫ��(&R)
Public Const conMenu_Edit_Compend = 3064     '*���ݹ���(&F)
Public Const conMenu_Edit_Affix = 3069      '����ģ������
Public Const conMenu_Edit_ElementChange = 3065      '*Ҫ����������
Public Const conMenu_Edit_ImportMerge = 3066      '����ϲ�·��
Public Const conMenu_Edit_UnImportMerge = 3067      'ȡ������ϲ�·��
Public Const conMenu_Edit_ViewMergeImport = 3068      '�鿴�ϲ�·����������
Public Const conMenu_Edit_Import = 3071      '*��������(&B)��

Public Const conMenu_Edit_ViewPacs = 4255     '������ͼ��ͱ���

'ҽ��(�༭)�˵�
Public Const conMenu_Edit_BatUnPack = 3072      '�������
Public Const conMenu_Edit_PacsApply = 3073      '*�������
Public Const conMenu_Edit_PacsApplyModi = 3074      '*�޸ļ������
Public Const conMenu_Edit_Apply = 3076      '*�´�����
Public Const conMenu_Edit_ApplyModi = 3077      '*�޸�����
Public Const conMenu_Edit_ApplyDel = 3078      '*ȡ������
Public Const conMenu_Edit_BloodApply = 3079    '��Ѫ����
Public Const conMenu_Edit_TraReaction = 30799    '��Ѫ��Ӧ
Public Const conMenu_Edit_BloodApplyModi = 3080    '�޸���Ѫ����
Public Const conMenu_Edit_ApplyView = 3081    '�鿴���뵥
Public Const conMenu_Edit_ApplyCustom = 30761 '�Զ������뵥
Public Const conMenu_Edit_MeetArrive = 3855    'ȷ��ҽ���μӻ���

'��ҽ��֤����
Public Const conMenu_Edit_ZyAdd = 3856     '��������
Public Const conMenu_Edit_ZyEdit = 3857   '�޸Ĵ���
Public Const conMenu_Edit_ZyView = 3858   '�鿴����
Public Const conMenu_Edit_ZyDel = 3859   'ɾ������

'����(�༭)�˵�
Public Const conMenu_Edit_Adjust = 3082      '*����(&J)
Public Const conMenu_Edit_Archive = 3083     '*�鵵(&R)
Public Const conMenu_Edit_UnArchive = 3084     'ȡ���鵵(&D)
Public Const conMenu_Edit_QCReport = 3085     '�ʿر���(&D)

'�ʿر���
Public Const conMenu_Edit_ItemEdit = 3086      '�༭
Public Const conMenu_Edit_ItemUndo = 3087       'ȡ��
Public Const conMenu_Edit_ItemSave = 3088       '����
Public Const conMenu_Edit_Exit = 3089           '�˳�
Public Const conMenu_Verify_AuditingLogin = 3090    '�鵵
Public Const conMenu_Verify_LogOut = 3099           'ȡ���鵵



Public Const conMenu_Edit_Save = 3503        '*����
Public Const conMenu_Edit_Sort = 3092        '*���ĵ�����
Public Const conMenu_Edit_Privacy = 3093     '*������˽��������
Public Const conMenu_Edit_Select = 3094      '*ѡ��
Public Const conMenu_Edit_DeSelect = 3095    '*ȡ��ѡ��
Public Const conMenu_Edit_Merge = 3096
Public Const conMenu_Edit_Dilute = 3098      '�걾ϡ��

Public Const conMenu_Edit_OperationApply = 3099    '��������
Public Const conMenu_Edit_ConsultationApply = 3100    '��������

'�ٴ�·��Ӧ�� ������Ŀ�ƶ�����
Public Const conMenu_Edit_Up = 3301         '����
Public Const conMenu_Edit_Down = 3302       '����
Public Const conMenu_Edit_SaveSorted = 3303     '��������
'Public Const conMenu_Manage_ThingAdd = 227    '�ӵ�(&A)
'Public Const conMenu_Manage_ThingModi = 228   '*����ִ�����(&M)
Public Const conMenu_Edit_Transf_Delete = 229   '�����ӵ�

'ҽ�ƿ�����
Public Const conMenu_Edit_CardBound = 3839    '�󶨿�
Public Const conMenu_Edit_CancelCardBound = 3842    'ȡ���󶨿�
Public Const conMenu_Edit_CardLoss = 3833    '��ʧ
Public Const conMenu_Edit_CardCancelLoss = 3834    'ȡ����ʧ
Public Const conMenu_Edit_Cardtrade = 3835    '����
Public Const conMenu_Edit_CardFill = 3836    '����
Public Const conMenu_Edit_CardBackMoney = 3837    '�˿�
Public Const conMenu_Edit_ChangePassWord = 3838    '��������
Public Const conMenu_Edit_ChangePassWord_Force = 3843    'ǿ�Ƶ�������
Public Const conMenu_Edit_MzToZy = 3840
Public Const conMenu_Edit_ZyToMz = 3841
Public Const conMenu_Edit_Family = 3844
Public Const conMenu_View_Family = 3845

'���˷��ò�ѯ
'----------------------------------------------------------------------
Public Const conMenu_Edit_PreBalance = 817    'Ԥ�ᵱǰ����
Public Const conMenu_Edit_PreBalanceAll = 818    'Ԥ�����в���
Public Const conMenu_Edit_Balance = 3011    '����
Public Const conMenu_Edit_Billing = 3003    '����
Public Const conMenu_Edit_Billing_Mulit = 3872    '��������

Public Const conMenu_Edit_ReBilling = 3004    'ֱ������

Public Const conMenu_Edit_ReBillingButton = 3017       '*���ó���(&E)
Public Const conMenu_Edit_ReBillingApply = 3019    '*��������(&L)
Public Const conMenu_Edit_ReBillingAudit = 3020    '*�������(&U)

Public Const conMenu_Edit_FeeAudit = 804    '��˻�ʼ���
Public Const conMenu_Edit_FeeUnAudit = 3565    'ȡ�����
Public Const conMenu_Edit_OverFeeAudit = 3566    '������
Public Const conMenu_Edit_PatiMemo = 3567   '��ע��Ϣ�༭
Public Const conMenu_Edit_PrePayMoney = 3568    'Ԥ����


'���ѿ�����
'---------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_CardPay = 3811    '����
Public Const conMenu_Edit_CardBathPay = 3812    '��������
Public Const conMenu_Edit_CardBack = 3813    '�˿�
Public Const conMenu_Edit_CardCancelBack = 38131    'ȡ������
Public Const conMenu_Edit_CardCallBack = 3814    '����
Public Const conMenu_Edit_CardCancelCallBack = 38141    'ȡ������

Public Const conMenu_Edit_CardInFull = 3816    '��ֵ
Public Const conMenu_Edit_CardInFullBack = 3817    '��ֵ����
Public Const conMenu_Edit_CardModify = 3818    '�޸Ŀ���Ϣ
Public Const conMenu_Edit_CardResume = 3819    '������
Public Const conMenu_Edit_CardStop = 38191    '��ͣ��
Public Const conMenu_Edit_MoveCard = 3821    '����ʱ���Ƴ���Ƭ
Public Const conMenu_Apply_AllCard = 3822    '����ʱ�����ݵ�ǰ���ݣ�Ӧ����������Ҫ�����ĵ���
Public Const conMenu_Apply_AllColumn = 3823    '����ʱ�����ݵ�ǰ����ָ�����У�Ӧ����������Ҫ�����Ĵ�����Ϣ
Public Const conMenu_COMBOX_INTERFACE = 3820    '���ѿ��ӿ�
Public Const conMenu_Square_BrushCard = 3824    '���ѿ�Ŀ¼+�ӿ����

'Ʊ�����
Public Const conMenu_Edit_DamnifyAdd = 3831    '��������
Public Const conMenu_Edit_DamnifyDelete = 3832  '����ɾ��
Public Const conMenu_Edit_UserType = 3833          'Ʊ��ʹ�����

'�������
'----------------------------------------------------------------------------------
Public Const conMenu_Edit_Triage = 2604  '����
Public Const conMenu_Edit_ModiyPati = 2607  '����������Ϣ
Public Const conMenu_Edit_ModiyPatiBaseInfo = 2610 '73743:���˻�����Ϣ����
Public Const conmenu_Edit_BackHospitalize = 3086    '����
Public Const conmenu_Edit_BackHospitalizeCancel = 3087    '����ȡ��

Public Const conmenu_Edit_ChangeNum = 3088    '���
Public Const conmenu_Edit_Leave = 3089    '���˲�����
Public Const conmenu_Edit_Wait = 3090    '���˴���

Public Const conmenu_View_TriagePati = 7101    '��ʾ�ѷ��ﲡ��
Public Const conmenu_View_AdmissionsPati = 7102    '��ʾ�ѽ��ﲡ��
Public Const conmenu_View_OverPati = 7103    '��ʾ����ɲ���
Public Const conmenu_View_Leave = 7104    '��ʾ�����ﲡ��
Public Const conmenu_View_AutoRefresh = 7120    '�Զ�ˢ��

'�ҺŰ���
'----------------------------------------------------------------------------------------------------------------------
Public Const conMenu_Edit_PlanAdd = 6820    '�ƻ�����
Public Const conMenu_Edit_PlanModify = 6821    '�޸ļƻ�����
Public Const conMenu_Edit_PlanDelete = 6822    'ɾ���ƻ�����
Public Const conMenu_Edit_PlanVerify = 6823    '��˼ƻ�����
Public Const conMenu_Edit_PlanCancel = 6824    'ȡ����˼ƻ�
Public Const conMenu_Edit_AllStartNO = 6825  'ȫ�����ùҺ���ſ���
Public Const conMenu_Edit_AllStopNO = 6826    'ȫ��ͣ�ùҺ���ſ���
Public Const conMenu_Edit_StopPlanTimes = 6827  'ͣ�ð��żƻ�
Public Const conMenu_Edit_ClearStopPlan = 6828  '�������ͣ�ð��żƻ�
Public Const comMenu_Edit_SetDateSegment = 6829    '�ҺŰ���ʱ�� ʱ�������
Public Const conMenu_Edit_SetPlanDateSeqment = 6830    '�Һżƻ�ʱ�� ʱ�������
Public Const comMenu_Edit_UnitRegModify = 6831      '�ҺŰ��ź�����λ��ŷ���
Public Const ComMenu_Edit_UnitRegArrangeModify = 6832    '�ҺŰ��żƻ�������λ��ŷ���
Public Const ComMenu_Edit_AutoDefaultLimitAppointment = 6833    '�ҺŰ����Զ�Ĭ����Լ��

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
Public Const conMenu_Edit_Transf_Save = 3501     '����
Public Const conMenu_Edit_Transf_Cancle = 3504   'ȡ��

Public Const conMenu_Edit_Transf_UndoEnd = 3505  '�������
Public Const conMenu_Edit_Transf_Negative = 3506    '����(+)
Public Const conMenu_Edit_Transf_Positive = 3507    '����(-)
Public Const conMenu_Edit_Transf_Reprint = 3508  '�ش򵥾�

Public Const conMenu_Edit_Transf_Liquid = 3509   '��Һ����
Public Const conMenu_Edit_Transf_Puncture = 3510    '���̲���


'������λ(�༭)�˵� 3531-3559
Public Const conMenu_Edit_Seat = 3530        '��λ
Public Const conMenu_Edit_Seat_Add = 3531    '��λ����
Public Const conMenu_Edit_Seat_Modify = 3532    '��λ�޸�
Public Const conMenu_Edit_Seat_Delete = 3533    '��λɾ��
Public Const conMenu_Edit_Seat_Clear = 3534  '���ռ�õ���λ
Public Const conMenu_Edit_Seat_Set = 3535    '������λ
Public Const conMenu_Edit_Seat_Swap = 3536    '������λ

Public Const conMenu_Edit_Seat_View = 3551    '�鿴
Public Const conMenu_Edit_Seat_Icon = 3552    'ͼ�귽ʽ
Public Const conMenu_Edit_Seat_List = 3553    '�б���ʽ
Public Const conMenu_Edit_Seat_Report = 3554    '������ʽ

Public Const conMenu_Edit_View_Seat = 3550  '��λͼ��
Public Const conMenu_Edit_View_GBed = 3555    '��ͨ��λ
Public Const conMenu_Edit_View_RBed = 3556    'ռ�ô�λ
Public Const conMenu_Edit_View_YBed = 3557    'ά����λ

Public Const conMenu_Edit_View_Gseat = 3558    '��ͨ��λ
Public Const conMenu_Edit_View_Rseat = 3559    'ռ����λ
Public Const conMenu_Edit_View_Yseat = 3560    'ά����λ


'�ݴ�ҩƷ(�༭)�˵� 3561 -3579
Public Const conMenu_Edit_Leave_Add = 3561    '����
Public Const conMenu_Edit_Leave_Modify = 3562    '�޸�
Public Const conMenu_Edit_Leave_Delete = 3563    'ɾ��
Public Const conMenu_Edit_Leave_Post = 3564    'ʹ�õǼ�
Public Const conMenu_Edit_Leave_SavePost = 3565    '����Ǽ�����
Public Const conMenu_Edit_Leave_UndoPost = 3565    '�����Ǽ�

Public Const conMenu_Edit_Leave_Repertory = 3571    '����ѯ
Public Const conMenu_Edit_Leave_AccountBook = 3572    '���̨��

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

'������λ����(�༭)�˵� 3601-3610
Public Const conMenu_Edit_Bed_Add = 3601            '����
Public Const conMenu_Edit_Bed_Modify = 3602         '����
Public Const conMenu_Edit_Bed_Delete = 3603         '����
Public Const conMenu_Edit_Bed_ToRepair = 3604       'ת����
Public Const conMenu_Edit_Bed_ToEmpty = 3605        'ת�մ�
Public Const conMenu_Edit_SelUnit = 3606            '����ѡ��


'Ժ�б༭\�鿴�˵�
Public Const conMenu_Edit_DelDayItem = 3802        'ɾ���ձ���ǰ����Ϣ
Public Const conMenu_Edit_BuildConstant = 3803        '���ɳ���������Ŀ
'
'��������\�鿴�˵�
Public Const conMenu_Edit_CfPay = 4000        '����������
Public Const conMenu_Edit_BillPay = 4001        '��Ʊ�ݷ���
Public Const conMenu_Edit_BillBackPay = 4002        '����������
Public Const conMenu_Edit_StopPay = 4003        '��ֹͣ���ϱ��

Public Const conMenu_View_FontSize = 4004         '�ֺ�����
Public Const conMenu_View_FontSize_1 = 4004         '9����
Public Const conMenu_View_FontSize_2 = 4004         '11����
Public Const conMenu_View_FontSize_3 = 4004         '15����

Public Const conMenu_Edit_OtherPay = 4005        '�������ⷿ����



'LISʹ�õĲɵ� 3650-3690
Public Const conMenu_Edit_QCRes = 3650         '*�ʿ�Ʒ
Public Const conMenu_LIS_Cancel = 3651         '*ȡ��
Public Const conMenu_LIS_PatientInfo = 3652    '������Ϣ
Public Const conMenu_LIS_HideList = 741       '���ز����б�
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
Public Const conMenu_LIS_SaveSample = 3708   '����걾
Public Const conMenu_LIS_DropSample = 3709   '���ٱ걾

'�����˵�
Public Const conMenu_Report_DrugQuery = 401    'ҩ���շ���ѯ(&H)
Public Const conMenu_Report_Reports = 402      '�������ñ���(&W)
Public Const conMenu_Report_MultiBill = 403    '��ӡ�ಡ�˵���(&K)
Public Const conMenu_Report_ClinicBill = 404   '��ӡ���Ƶ���(&J)��
Public Const conMenu_Report_AdviceBill1 = 405  '����ҽ����(&P)
Public Const conMenu_Report_AdviceBill2 = 406  '��ʱҽ����(&T)
Public Const conMenu_Report_AdviceBill3 = 407  'ҽ����¼��(&B)
Public Const conMenu_Report_WorkLog = 408      '�����ձ�(&O)
Public Const conMenu_Report_ClinicIndexBill = 409      '����ָ����
Public Const conMenu_Report_BloodInstant = 410    '��Ѫִ�е�(&I)

'�鿴�˵�
Public Const conMenu_View_ToolBar = 701              '������(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Public Const conMenu_View_StatusBar = 702            '״̬��(&S)
Public Const conMenu_View_Append = 703               '������Ϣ(&A)
Public Const conMenu_View_Difference = 704              '��ʾ����(&D)
Public Const conMenu_View_Contrast = 705                 '�ԱȲ鿴
Public Const conMenu_View_NoticBoard = 706           '����������
Public Const conMenu_View_AdviceLost = 707           'ҽ��ˢ��ʱ��λ�����

Public Const conMenu_View_FontSize_S = 4041            'ҽ�����壺С����
Public Const conMenu_View_FontSize_M = 4040            'ҽ�����壺������
Public Const conMenu_View_FontSize_L = 4042            'ҽ�����壺������

Public Const conMenu_View_StPath = 4043                '�鿴��׼·��
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
Public Const conMenu_View_ShowDel = 738           '��ʾɾ������
Public Const conMenu_View_Hide = 741                 '*����(&H)
Public Const conMenu_View_Show = 742                 '*��ʾ(&S)
Public Const conMenu_View_Forward = 743              '*ǰ��(&F)
Public Const conMenu_View_Backward = 744             '*����(&B)
Public Const conMenu_View_Dept = 745                '�鿴����
Public Const conMenu_View_Location = 746            '��λ
Public Const conMenu_View_LocationItem = 747        '��λ��Ŀ
Public Const conMenu_View_Option = 781               'ѡ��(&O)
Public Const conMenu_View_Refresh = 791              '*ˢ��(&R)
Public Const conMenu_View_RefreshSpare = 7911        '��ȡ����ҽ��
Public Const conMenu_View_Jump = 792                 '��ת(&J)
Public Const conMenu_View_Warrant = 794              '������Ϣ����
Public Const conMenu_View_Shell = 3818                '��ӳ���


Public Const conMenu_View_Navigatebeginning = 7401           '*��һ��(&F)
Public Const conMenu_View_Navigateleft = 7402                '*��һ��(&F)
Public Const conMenu_View_Navigateright = 7403               '*��һ��(&F)
Public Const conMenu_View_Navigateend = 7404                 '*���һ��(&F)

'�����ز˵�
Public Const conMenu_View_Detail = 7501             '�鿴��ϸ����
Public Const conMenu_View_ChargeAndBilllTotal = 7502             '�鿴�տƱ�ݻ���

'���˽��ʹ���
'----------------------------------------------------------------------
Public Const conMenu_File_CashCount = 4801    '�ֽ�㳮
Public Const conMenu_File_SetInsure = 4802    '�������
Public Const conMenu_Edit_ClinicBalance = 4803
Public Const conMenu_Edit_InHosBalance = 4804
Public Const conMenu_Edit_BatchBalance = 4805
Public Const conMenu_Edit_UnitBalance = 4806
Public Const conMenu_Edit_FeeManage = 4807
Public Const conMenu_Edit_ClinicToHos = 4808
Public Const conMenu_Edit_ToHosCancel = 4809
Public Const conMenu_Edit_ErrReBalance = 4810
Public Const conMenu_Edit_ErrCancelBalance = 4811
Public Const conMenu_Edit_ErrDelBalance = 4812
Public Const conMenu_Edit_CancelBalance = 4813
Public Const conMenu_Edit_PrintAmend = 4814
Public Const conMenu_Edit_PrintDetail = 4815
Public Const conMenu_Edit_PrintAmendByPati = 4816
Public Const conMenu_Edit_WriteCard = 4817
Public Const conMenu_View_RefreshType = 4818
Public Const conMenu_View_RefreshType_No = 4819
Public Const conMenu_View_RefreshType_Ask = 4820
Public Const conMenu_View_RefreshType_Auto = 4821
Public Const conMenu_File_FeeCollect = 4822

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

'��Ա����(����ListView�ؼ�����ʾ��ʽ:��ͼ��;Сͼ��;�б�;��ϸ����)
Public Const conMenu_View_LargeICO = 7610  '��ͼ��
Public Const conMenu_View_MinICO = 7611  'Сͼ��
Public Const conMenu_View_ListICO = 7612  '�б�
Public Const conMenu_View_DetailsICO = 7613  '��ϸ�б�

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
Public Const conMenu_Tool_MeetIdea = 8033         '��д�������(&W)
Public Const conMenu_Tool_Sign = 804            '*����ǩ��(&I)
Public Const conMenu_Tool_SignNew = 8041            '����ǩ��(&I)
Public Const conMenu_Tool_SignVerify = 8042         '��֤ǩ��(&V)
Public Const conMenu_Tool_SignEarse = 8043          'ȡ��ǩ��(&E)
Public Const conMenu_Tool_SignAuditAffirm = 8044         '*�ϼ���ǩ(&V)
Public Const conMenu_Tool_SignAuditCancel = 8045          'ȡ����ǩ(&E)
Public Const conMenu_Tool_Community = 805       '*��������(&U)
Public Const conMenu_Tool_MedRecAudit = 806        '�������(&M)
Public Const conMenu_Tool_MedRecAuditSubmit = 8061      '�ύ���(&S)
Public Const conMenu_Tool_MedRecAuditCancel = 8062      'ȡ���ύ(&C)
Public Const conMenu_Tool_MedRecAuditResponse = 8063    '��鷴��(&M)
Public Const conMenu_Tool_MedRecAuditWriteResponse = 8064    '��д������
Public Const conMenu_Tool_Archive = 807         '*��Ա����(&I) / ����סԺ ���Ӳ�������
Public Const conMenu_Tool_ExaReport = 808 '��������ܼ챨��
Public Const conMenu_Tool_Monitor = 811         '*���(&M)
Public Const conMenu_Tool_Monitor_1 = 81101         'ʱ��Ҫ����(&T)
Public Const conMenu_Tool_Monitor_2 = 81102         '����Ҫ����(&C)
Public Const conMenu_Tool_Assistant = 812       '*����(&A)
Public Const conMenu_Tool_Analyse = 813         '*����(&Y)
Public Const conMenu_Tool_Search = 814          '*����(&S)
Public Const conMenu_Tool_Define = 815          '*����(&D)
Public Const conMenu_Tool_Report = 816          '*����(&P)
Public Const conMenu_Tool_Apply = 817           '*Ӧ��(&A)
Public Const conMenu_Tool_BathSend = 818        '�������͵�����
Public Const conMenu_Tool_Option = 819          'ѡ��(&O),�Ӵ�������ù���
Public Const conMenu_Tool_KssAudit = 820        '*������ҩ���
Public Const conMenu_Tool_OPSAudit = 821        '������˹���
Public Const conMenu_Tool_CISMed = 822        '�ٴ��Թ�ҩ
Public Const conMenu_Tool_TransAudit = 823   '��Ѫ�ּ�����
Public Const conMenu_Tool_MedRatio = 824      'ҩռ��
Public Const conMenu_Tool_HealthCard = 825      '���񽡿���
Public Const conMenu_Tool_OPSEmpower = 826   '������Ȩ����
Public Const conMenu_Tool_UnitSubject = 827     '�����������
Public Const conMenu_Tool_UnitNBoard = 828      '��������������
Public Const conMenu_Tool_RisPrint = 829       '��ӡRISԤԼ��
Public Const conMenu_Tool_RisPrintBat = 830    '������ӡRISԤԼ��

Public Const conMenu_Tool_PlugIn = 890          '���
Public Const conMenu_Tool_PlugIn_Item = 89000   '�����,ʵ������Ϊ conMenu_Tool_PlugIn_Item + n, 1<=n<=99

'PACS����վ�˵�
Public Const conMenu_Manage_CriticalValues = 8342           'Σ��ֵ��¼
Public Const conMenu_Manage_CriticalSituation = 8343        'Σ�����
Public Const conMenu_Manage_Normal = 8344                   '����
Public Const conMenu_Manage_Critical = 8345                 'Σ��

Public Const conMenu_Manage_Result = 8300       '�����
Public Const conMenu_Manage_Negative = 8301      '���������
Public Const conMenu_Manage_Positive = 8302      '���������

Public Const conMenu_Manage_ImageQuality = 8303       'Ӱ������
Public Const conMenu_Manage_ImageFirst = 8304         '��һ��
Public Const conMenu_Manage_ImageSecond = 8305        '�ڶ���
Public Const conMenu_Manage_ImageThird = 8396         '������
Public Const conMenu_Manage_ImageFourth = 8397        '���ļ�

Public Const conMenu_Manage_ReportQuality = 8346       '��������
Public Const conMenu_Manage_ReportFirst = 8347         '��һ��
Public Const conMenu_Manage_ReportSecond = 8348        '�ڶ���
Public Const conMenu_Manage_ReportThird = 8349         '������
Public Const conMenu_Manage_ReportFourth = 8350        '���ļ�

Public Const conMenu_Manage_FuHeLevel = 8220         '�������
Public Const conMenu_Manage_FuHe = 8221             '����
Public Const conMenu_Manage_JiBenFuHe = 8222        '��������
Public Const conMenu_Manage_BuFuHe = 8223           '������

Public Const conMenu_Manage_SwitchUser = 8338       '�л��û�
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
Public Const conMenu_Manage_ReportExecutor = 8228       '����ִ��
Public Const conMenu_Manage_RelatingPatiet = 8210       '��������
Public Const conMenu_Manage_LocateType = 8211           '��λ��ʽ
Public Const conMenu_Manage_LocateValue = 8212          '��λֵ
Public Const conMenu_Manage_DeleteSeries = 8213         'ɾ������ͼ��
Public Const conMenu_Manage_DeleteImage = 8214          'ɾ��ͼ��
Public Const conMenu_Manage_FilmRelease = 8215          '��Ƭ����
Public Const conMenu_Manage_ReportFilmRelease = 8216    '���潺Ƭͬʱ����
Public Const conMenu_Manage_Release = 8217              '����
Public Const conMenu_Manage_Burn = 8218                 '��¼
Public Const conMenu_Manage_RefreshImg = 8219           'ˢ��ͼ��
Public Const conMenu_Manage_Query = 8224                '��ѯ
Public Const conMenu_Manage_ConfigQuery = 8225          '���ò�ѯ����
Public Const conMenu_Manage_CustomQuery = 8226          '�Զ����ѯ
Public Const conMenu_Manage_SetXWParam = 8227             'RISPACS����
Public Const conMenu_Manage_CloseQuery = 8229           '�رղ�ѯ
Public Const conMenu_Manage_FilmPrevew = 8230           '��ƬԤ��
Public Const conMenu_Manage_FilmPrint = 8231            '��Ƭ��ӡ
Public Const conMenu_Manage_FilmDelete = 8232           '��Ƭɾ��
Public Const conMenu_Manage_CheckList = 8233            '�鿴���뵥
Public Const conMenu_Manage_SendArrange = 8234          '���Ͱ���
Public Const conMenu_Manage_PacsPlugIn = 8235           'Pacs���������ܹҽ�
Public Const conMenu_Manage_PacsPlugCfg = 8236          '�������

'PACS����༭��
Public Const conMenu_PacsReport_Group = 8336        '����
Public Const conMenu_PacsReport_SelFormat = 8309    'ѡ�񱨸��ʽ
Public Const conMenu_PacsReport_SelFormat_Item = 8310    'ѡ�񱨸��ʽ
Public Const conMenu_PacsReport_Save = 8311         '���汨��
Public Const conMenu_PacsReport_Sign = 8312         '����ǩ��
Public Const conMenu_PacsReport_Reject = 8340         '���沵��
Public Const conMenu_PacsReport_RejectHistory = 8341  '������ʷ
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
Public Const conMenu_PacsReport_History_Times = 8323    '��ʷ����
Public Const conMenu_PacsReport_DelMiniImage = 8324     'ɾ����������ͼ
Public Const conMenu_PacsReport_SelMiniImage = 8325     '��ȡ��������ͼ
Public Const conMenu_PacsReport_RptImg2CapImg = 8326    '�ڱ���ͼ���ߺͲɼ����߼��л�
Public Const conMenu_PacsReport_PrivOrder = 8327        '��һ��ҽ��
Public Const conMenu_PacsReport_NextOrder = 8328        '��һ��ҽ��
Public Const conMenu_PacsReport_AddNumber = 8329        '�����������������
Public Const conMenu_PacsReport_RepFormat = 8330        '�Զ��屨��ѡ���ʽ
Public Const conMenu_PacsReport_RepFormat_Item = 8331   '�Զ��屨��ѡ��ľ����ʽ��
Public Const conMenu_PacsReport_SaveWord = 8332         '����ʾ�ʾ��
Public Const conMenu_PacsReport_ClearWritingState = 8333    '������浱ǰ������
Public Const conMenu_PacsReport_VerifySign = 8334           '����ǩ����֤
Public Const conMenu_PacsReport_VerifySign_Item = 8335      '����ǩ����֤,�Ծ���汾����֤
Public Const conMenu_PacsReport_Default = 8337              '�ָ�Ĭ�Ͻ���
Public Const conMenu_PacsReport_FinalShowMode = 8339        '����״̬��ʾ

'�ɼ��˵�
Public Const conMenu_Cap_Group = 8090           '�ɼ�
Public Const conMenu_Cap_Dynamic = 8100         '��̬��ʾ(&V)
Public Const conMenu_Cap_MarkMap = 8101         'Ӱ��ɼ�(&C)
Public Const conMenu_Cap_Import = 8102          'Ӱ����(&I)
Public Const conMenu_Cap_DevSet = 8103          'Ӱ���豸����(&D)
Public Const comMenu_Cap_Process = 8104         'Ӱ����
Public Const conMenu_Cap_Record = 8105          '¼��(&R)
Public Const conMenu_Cap_DelImg = 8097          'ɾ��ͼ��
Public Const conMenu_Cap_Full_Screen = 8098     'ȫ��(&U)
Public Const conMenu_Cap_Record_Stop = 8099     'ֹͣ¼��(&O)
Public Const conMenu_Cap_Play = 8106            '����(&P)
Public Const conMenu_Cap_Stop = 8107            'ֹͣ(&T)
Public Const conMenu_Cap_Forward = 8108         '���(&F)
Public Const conMenu_Cap_Back = 8109            '����(&B)
Public Const conMenu_Cap_SaveAs = 8126    '8110          '����¼��(&S)
Public Const conMenu_Cap_OpenStudyList = 8122   '�򿪼���б�
Public Const conMenu_Cap_StudySyncState = 8123  'Ӱ����ͬ��״̬
Public Const conMenu_Cap_RecordAudio = 8125     '¼��
Public Const conMenu_Cap_After_Capture = 8140   '��̨�ɼ�
Public Const conMenu_Cap_After_Record = 8141    '��̨¼��
Public Const conMenu_Cap_After_Tag = 8142       '���±��


'�����˵� 3900-3970
Public Const conMenu_PatholManage = 3900                '��������
Public Const conMenu_Pathol_Antibody_Manage = 3901      '�������
Public Const conMenu_Pathol_MealManage = 3902           '�ײ�ά��
Public Const conMenu_Pathol_Request = 3903              '��������
Public Const conMenu_Pathol_ReportDelay = 3904          '�����ӳ�
Public Const conMenu_Pathol_ConRequest = 3905           '��������
Public Const conMenu_Pathol_ConFeedback = 3906          '���ﷴ��
Public Const conMenu_Pathol_Decalin_Task = 3907         '�Ѹ��������
Public Const conMenu_Pathol_BatSlicesAccept = 3908      '��Ƭ��������
Public Const conMenu_Pathol_BatSlicesSure = 3909        '��Ƭ����ȷ��
Public Const conMenu_Pathol_BatSpeExamAccept = 3910     '�ؼ���������
Public Const conMenu_Pathol_BatSpeExamSure = 3911       '�ؼ�����ȷ��
Public Const conMenu_Pathol_BatProcess = 3912           '��������
Public Const conMenu_Pathol_Quality_Manage = 3913       '������������
Public Const conMenu_Pathol_NumConfig = 3914            '������������
Public Const conMenu_Pathol_WorkModule = 3915           'վ��ģʽ����
Public Const conMenu_PatholSlices_Quality = 3916        '������Ƭ����

'�����걾�˵�
Public Const conMenu_PatholSpecimen = 3940                  '�����걾
Public Const conMenu_PatholSpecimen_LAB = 3941              '��ǩ
Public Const conMenu_PatholSpecimen_PreviewLab = 3942       'Ԥ����ǩ
Public Const conMenu_PatholSpecimen_PrintLab = 3943         '��ӡ��ǩ
Public Const conMenu_PatholSpecimen_ACP = 3944              '���յ�
Public Const conMenu_PatholSpecimen_PrintAccept = 3945      '��ӡ���յ�
Public Const conMenu_PatholSpecimen_PreviewAccept = 3946    'Ԥ�����յ�
Public Const conMenu_PatholSpecimen_Get = 3947              '��ȡ�걾
Public Const conMenu_PatholSpecimen_Del = 3948              'ɾ���걾
Public Const conMenu_PatholSpecimen_Save = 3949             '����걾
Public Const conMenu_PatholSpecimen_Accept = 3950           '���ձ걾
Public Const conMenu_PatholSpecimen_Reject = 3951           '���ձ걾


'������Ƭ�˵�
Public Const conMenu_PatholSlices = 3960                    '������Ƭ
Public Const conMenu_PatholSlices_LAB = 3961                '��ǩ
Public Const conMenu_PatholSlices_PreviewLAB = 3962         'Ԥ����ǩ
Public Const conMenu_PatholSlices_PrintLAB = 3963           '��ӡ��ǩ
Public Const conMenu_PatholSlices_List = 3964               '�嵥
Public Const conMenu_PatholSlices_PreviewList = 3965        'Ԥ���嵥
Public Const conMenu_PatholSlices_PrintList = 3966          '��ӡ�嵥
Public Const conMenu_PatholSlices_RequestView = 3967        '����鿴
Public Const conMenu_PatholSlices_Accept = 3968             '��Ƭ����
Public Const conMenu_PatholSlices_Finish = 3969             '��Ƭ���


'����ȡ�Ĳ˵�
Public Const conMenu_PatholMaterial = 3970                  '����ȡ��
Public Const conMenu_PatholMaterial_PrintAll = 3971         '��ӡ����
Public Const conMenu_PatholMaterial_PreviewAll = 3972       'Ԥ������
Public Const conMenu_PatholMaterial_PrintSingle = 3973      '������ӡ
Public Const conMenu_PatholMaterial_PreviewSingle = 3974    '����Ԥ��
Public Const conMenu_PatholMaterial_RequestView = 3975      '����鿴
Public Const conMenu_PatholMaterial_Get = 3976              '�Ŀ���ȡ
Public Const conMenu_PatholMaterial_Del = 3977              'ɾ���Ŀ�
Public Const conMenu_PatholMaterial_Save = 3978             '����Ŀ�
Public Const conMenu_PatholMaterial_Sure = 3979             'ȷ��ȡ��
Public Const conMenu_PatholMaterial_Decalcification = 3980  '�Ѹ�
Public Const conMenu_PatholMaterial_ChangeVat = 3981        '����
Public Const conMenu_PatholMaterial_CancelVat = 3982        '����
Public Const conMenu_PahtolMaterial_Finish = 3983           '���


'�����ؼ�˵�
Public Const conMenu_PatholSpeExam = 3990                   '�����ؼ�
Public Const conMenu_PatholSpeExam_LAB = 3991               '��ǩ
Public Const conMenu_PatholSpeExam_PreviewLAB = 3992        'Ԥ����ǩ
Public Const conMenu_PatholSpeExam_PrintLab = 3993          '��ӡ��ǩ
Public Const conMenu_PatholSpeExam_List = 3994              '�嵥
Public Const conMenu_PatholSpeExam_PreviewList = 3995       'Ԥ���嵥
Public Const conMenu_PatholSpeExam_PrintList = 3996         '��ӡ�嵥
Public Const conMenu_PatholSpeExam_RequestView = 3997       '����鿴
Public Const conMenu_PatholSpeExam_Accept = 3998            '�ؼ����
Public Const conMenu_PatholSpeExam_Finish = 3999            '�ؼ����

'�������̱���
Public Const conMenu_PatholProRep = 4100                    '�������̱���
Public Const conMenu_PatholProRep_Print = 4101              '��ӡ
Public Const conMenu_PatholProRep_Preview = 4102            'Ԥ��
Public Const conMenu_PatholProRep_Already = 4103            '����
Public Const conMenu_PatholProRep_Back = 4104               '����
Public Const conMenu_PatholProRep_Clear = 4105              '�������
Public Const conMenu_PatholProRep_Input = 4106              '�ؼ���Ŀ¼��
Public Const conMenu_PatholProRep_New = 4107                '��������
Public Const conMenu_PatholProRep_Del = 4108                'ɾ������
Public Const conMenu_PatholProRep_Save = 4109               '���ݱ���

'�����ײ�ά��
Public Const conMenu_PatholMeal_Save = 4110
Public Const conMenu_PatholMeal_Cancel = 4111
Public Const conMenu_PatholMeal_AddRecord = 4112
Public Const conMenu_PatholMeal_ModRecord = 4113
Public Const conMenu_PatholMeal_DelRecord = 4114
Public Const conMenu_PatholMeal_UpRow = 4115
Public Const conMenu_PatholMeal_DownRow = 4116

'���������˵�
Public Const conMenu_Pathol_ArchivesManage = 3920    '��������
Public Const conMenu_Pathol_ArchivesClass = 3921    '�����������
Public Const conMenu_Pathol_ArchivesPlace = 3922    '����λ������
Public Const conMenu_Pathol_ArchivesFile = 3923  '�����ļ��鵵
Public Const conMenu_Pathol_ArchivesMaterial = 3924  '�������Ϲ鵵
Public Const conMenu_Pathol_ArchivesLend = 3925      '�������Ĺ���


'�ղع����˵�
Public Const conMenu_Collection = 3930
Public Const conMenu_Collection_Manage = 3931    '�ղع���
Public Const conMenu_Collection_To = 3932    '�ղص�
Public Const conMenu_Collection_ViewShare = 3933    '�鿴����
Public Const comMenu_Collection_Type = 3934    '��̬�ղز˵�

'���˵�
Public Const comMenu_Petition_Capture = 3935    'ɨ�����뵥
Public Const comMenu_Petition_View = 3936       '�鿴���뵥

'Pacsռ�ò���
Public Const conMenu_Reserve_23 = 4120
Public Const conMenu_Reserve_24 = 4125
Public Const conMenu_Reserve_25 = 4130
Public Const conMenu_Reserve_26 = 4135
Public Const conMenu_Reserve_27 = 4140
Public Const conMenu_Reserve_28 = 4145
Public Const conMenu_Reserve_29 = 4150
Public Const conMenu_Reserve_30 = 4155
Public Const conMenu_Reserve_31 = 4160
Public Const conMenu_Reserve_32 = 4165
Public Const conMenu_Reserve_33 = 4170
Public Const conMenu_Reserve_34 = 4175
Public Const conMenu_Reserve_35 = 4180
Public Const conMenu_Reserve_36 = 4185
Public Const conMenu_Reserve_37 = 4190
Public Const conMenu_Reserve_38 = 4195


Public Const conMenu_Img_Group = 8110           'Ӱ��
Public Const conMenu_Img_Look = 8127            'Ӱ���Ƭ(&S)
Public Const conMenu_Img_Contrast = 8112        '��Ƭ�Ա�(&E)
Public Const conMenu_Img_Look3D = 8111            'Ӱ���Ƭ(&S)
Public Const conMenu_Img_Delete = 8113          'ͼ��ɾ��(&K)
Public Const conMenu_Img_Query = 8114           'Q/R��ȡͼ��(&Q)


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
Public Const conMenu_Queue_Restore = 8253     '�ָ�
Public Const conMenu_Queue_ReCall = 8254        '�غ�
Public Const conMenu_Queue_Abandon = 8255       '����
Public Const conMenu_Queue_Refresh = 8256       'ˢ��
Public Const conMenu_Queue_Setup = 8257         '��������
Public Const conMenu_Queue_Update = 8258        '�޸�
Public Const conMenu_Queue_Broadcast = 8259     '�㲥
Public Const conMenu_Queue_Pause = 8260         '��ͣ
Public Const conMenu_Queue_Finaled = 8261       '��ɾ���
Public Const conMenu_Queue_Find = 8262          '����
Public Const conMenu_Queue_ComeBack = 8263      '����
Public Const conMenu_Queue_RecDiagnose = 8264   '����

Public Const conMenu_Queue_Locate = 8265        '��λ
Public Const conMenu_Queue_LocateValue = 8266    '��λֵ
Public Const conMenu_Queue_LocateType = 8267    '��λ����

Public Const conMenu_Queue_PrintNumber = 8272       '���
Public Const conMenu_Queue_InsertQueue = 8273       '���
Public Const conMenu_Queue_RestartQueue = 8274      '����

'�����طּ�����
Public Const conMenu_Kss_Jurisdiction = 9001    'Ȩ��
Public Const conMenu_Kss_Grant = 9002           '��Ȩ
Public Const conMenu_Kss_Cancellation = 9003    'ȡ����Ȩ
Public Const conMenu_Kss_Adjustment = 9004      '����Ȩ��
Public Const conMenu_Kss_ShowCancel = 9005      '��ʾȡ����Ȩ����Ա

'ҩ���շ���ѯ��������˵�
Public Const conMenu_FontSet = 509
Public Const conMenu_FontSet_FontSize_S = 4041            'ҩ���շ���ѯ����С����
Public Const conMenu_FontSet_FontSize_L = 4042            'ҩ���շ���ѯ���񣺴�����

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

'�ٴ�·���༭�˵�
Public Const conMenu_Edit_OutLogModi = 601    '�޸ĳ����Ǽ�
Public Const conMenu_Edit_OutLogView = 602    '�鿴�����Ǽ�
'��׼·��ά��
Public Const conMenu_Edit_NewCourseItem = 3001    '��������
Public Const conMenu_Edit_ModifyCourseItem = 3003    '�޸Ķ���
Public Const conMenu_Edit_DelCourseItem = 3004   'ɾ������

Public Const conMenu_Edit_ModifyTableContent = 3822    '�޸ı�������

Public Const conMenu_Edit_NewPath = 9002    '����·��
Public Const conMenu_Edit_ModifyPath = 9004    '�޸�·��
Public Const conMenu_Edit_DelPath = 9003   'ɾ��·��

Public Const conMenu_Edit_NewTable = 3051    '��������
Public Const conMenu_Edit_ModifyTable = 3053    '�޸ı���
Public Const conMenu_Edit_DelTable = 3054   'ɾ������
'��׼·���������ݱ༭
Public Const conMenu_NewRow = 100    '����һ��
Public Const conMenu_NewCol = 101    '����һ��
Public Const conMenu_DelCol = 102    'ɾ����
Public Const conMenu_DelRow = 103    'ɾ����
Public Const conMenu_ClearItem = 104    '�������
Public Const conMenu_Save = 107    '����
Public Const conMenu_Edit = 108    '�༭
Public Const conMenu_Exit = 111    '�˳�
'�����˵�
Public Const conMenu_Help_Help = 901        '*��������(&H)
Public Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_Help_Web_Forum = 9023      '������̳(&F)
Public Const conMenu_Help_Web_Mail = 6879       '*���ͷ���(&M)
Public Const conMenu_Help_About = 991       '����(&A)��

'������������
'*********************************************************************
'CommandBar���г�������
Public Const XTP_ID_WINDOW_LIST = 35000    '�����б�
Public Const XTP_ID_TOOLBARLIST = 59392    '�������б�
Public Const ID_INDICATOR_CAPS = 59137    '״̬������д��
Public Const ID_INDICATOR_NUM = 59138    '״̬�������֣�
Public Const ID_INDICATOR_SCRL = 59139    '״̬����������

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

