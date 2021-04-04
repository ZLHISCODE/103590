Attribute VB_Name = "mdlPathology"
Option Explicit


'����ϵͳ���
Public Const G_LNG_PATHOLSYS_NUM = 1294

'����鵵ģ����
Public Const G_LNG_PATHOLARCHIVES_NUM = 1295

'�������ģ����
Public Const G_LNG_PATHOLBORROW_NUM = 1296

'���������ʧ����ģ����
Public Const G_LNG_PATHOLLOSE_NUM = 1297


'ͼ��id����
Public Const G_INT_ICONID_SPECIMEN = 10015
Public Const G_INT_ICONID_MATERIAL = 10016
Public Const G_INT_ICONID_SLICES = 10017
Public Const G_INT_ICONID_SPEEXAM = 10018
Public Const G_INT_ICONID_PROREPORT = 10019
Public Const G_INT_ICONID_SLICESSURE = 8133
Public Const G_INT_ICONID_BATPROCESS = 802


'���������Ͷ���
Public Enum StudyType
    stNormal = 0  '����
    stIce = 1     '����
    stCell = 2    'ϸ��
    stMeet = 3    '����
    stAutopsy = 4 'ʬ��
    stSpeed = 5   '����ʯ��
End Enum



'��������̶���
Public Enum TStudyProcedure
    spReserve = 0
    spMaterial = 1
    spSlices = 2
    spDiagnose = 3
    spMianyi = 4
    spTeran = 5
    spFenzi = 6
    spAgainMaterial = 8
    spAgainSlices = 9
    spFinal = 10
End Enum


'����ִ�в��趨��
Public Enum TExecuteStep
    None = 0        'δִ��
    NeedDo = 1      '��ִ��
    AcceptDo = 2    '�ѽ���
    AlreadDo = 3    '��ִ��
End Enum


'��������
Public Enum TRequestType
    rtMianyi = 0    '�����黯����
    rtTeran = 1     '����Ⱦɫ����
    rtFenzi = 2     '���Ӳ�������
    rtSlices = 3    '����Ƭ����
    rtMaterial = 4  '��ȡ������
End Enum


'�ؼ����Ͷ���
Public Enum TSpeexamType
    stMianyi = 0    '�����黯
    stTeshu = 1     '����Ⱦɫ
    stFenzi = 2     '���Ӳ���
End Enum

'������״̬��Ϣ
Public Type TStudyStateInf
    lngPatholAdviceId As Long   '����ҽ��ID
    lngStudyType As Long        '�������
    lngMaterialStep As Long     'ȡ�Ĺ���
    lngSlicesStep As Long       '��Ƭ����
    lngMianYiStep As Long       '���߹���
    lngFenZiStep As Long        '���ӹ���
    lngTeRanStep As Long        '��Ⱦ����
    strPatholNumber As String   '�����
End Type


'��׼������
Public Const glngStandardRowCount As Long = 51
'��׼�и߶�
Public Const glngStandardRowHeight As Long = 300
'��Ƭ�������ؼ����Ƭ��
Public Const glngSlicesRowCount As Long = 51
'���������
Public Const glngMaxRowCount As Long = 101



'����������ʱ���ʽ�ַ���
Public Const gstrFullDateTimeFormat = "yyyy-mm-dd hh:mm:ss"

'���ڸ�ʽ�ַ���
Public Const gstrDateFormat = "yyyy-mm-dd"

'ʱ���ʽ�ַ���
Public Const gstrTimeFormat = "hh:mm:ss"

'��������ĵ�Ԫ����ɫ
Public Const gCellErrColor As Long = &HC0C0FF



'�ж����ʽΪ��������,�Ƿ�����(Ĭ�ϲ�����),�ɷ�༭(Ĭ�Ͽɱ༭),�Ƿ�Button��ť(Ĭ�ϲ���),���
'
'���������Ϊ���ԡ����ʾ����Ϊ��չ�У���Ҫ���ڿ����еĸ߶�
'
'���������£�
'��ʾ����>�ֶ�����
'hide����ʾ����
'btn����ʾ������button��ť
'read����ʾ����Ϊֻ��
'merge����ʾ����Ϊ�ϲ��У��кϲ���
'check����ʾ�Ƿ���checkbox�ؼ�
'w1600����ʾ���Ϊ1600
'key:��ʾΪ�ؼ��ֶ�
'�������������Ϊ���������ʾ����Ϊ������CheckBox�У�����������
'fulldatetime��yyyy-mm-dd hh:mm:ss
'onlydate��yyyy-mm-dd
'onlytime��hh:mm:ss
'shortdatetime��yyyy-mm-dd hh:mm
'cbx<0-��,1-��,2-δ����>����ʾ����Ϊ��ѡ��
'Align<8,0>������λ������
'colleft,colcenter,colright����ʾ�еĶ��뷽ʽ
'txtleft,txtcenter,txtright����ʾ�ı��Ķ��뷽ʽ
'chkleft,chkcenter,chkright����ʾcheck�Ķ��뷽ʽ
'tdate����ʾʱ������
'tnum����ʾ��������
'tstr����ʾ�ַ�������
'uncfg����ʾ��������������


'��ת���������£�
'���ؼ�����:0-�����黯,1-����Ⱦɫ,2-���Ӳ���,els-����|��ǰ״̬:0-δ����,1-�ѽ���,2-�����|�嵥״̬:0-<nocheck>δ��ӡ,1-<check>�Ѵ�ӡ
'<nocheck>��ʾ��������ʾʱ����Ԫ������δѡ�еĹ�ѡ��
'<check>��ʾ��������ʾʱ����Ԫ��������ѡ�еĹ�ѡ��
'els:��ʾ��������������ʱ��ȡ��ֵ

'============================================================================================================================

Public Const gstrPatholCol_ID           As String = "ID"
Public Const gstrPatholCol_�����       As String = "�����"
Public Const gstrPatholCol_����         As String = "����"
Public Const gstrPatholCol_�Ա�         As String = "�Ա�"
Public Const gstrPatholCol_�Ŀ��       As String = "�Ŀ��"
Public Const gstrPatholCol_�걾����     As String = "�걾����"
Public Const gstrPatholCol_ȡ��λ��     As String = "ȡ��λ��"
Public Const gstrPatholCol_�������     As String = "�������"
Public Const gstrPatholCol_������ϸ     As String = "������ϸ"
Public Const gstrPatholCol_��ʧԭ��     As String = "��ʧԭ��"
Public Const gstrPatholCol_����״̬     As String = "����״̬"
Public Const gstrPatholCol_����״̬     As String = "����״̬"
Public Const gstrPatholCol_���ĺ�       As String = "���ĺ�"
Public Const gstrPatholCol_������       As String = "������"
Public Const gstrPatholCol_��������     As String = "��������"
Public Const gstrPatholCol_�黹����     As String = "�黹����"
Public Const gstrPatholCol_֤������     As String = "֤������"
Public Const gstrPatholCol_֤������     As String = "֤������"
Public Const gstrPatholCol_��ϵ�绰     As String = "��ϵ�绰"
Public Const gstrPatholCol_��ϵ��ַ     As String = "��ϵ��ַ"
Public Const gstrPatholCol_Ѻ��         As String = "Ѻ��"
Public Const gstrPatholCol_��������     As String = "��������"
Public Const gstrPatholCol_��������     As String = "��������"
Public Const gstrPatholCol_����ԭ��     As String = "����ԭ��"
Public Const gstrPatholCol_�黹״̬     As String = "�黹״̬"
Public Const gstrPatholCol_��ע         As String = "��ע"
Public Const gstrPatholCol_ȷ��״̬     As String = "ȷ��״̬"
Public Const gstrPatholCol_�������     As String = "�������"
Public Const gstrPatholCol_�鵵ID       As String = "�鵵ID"
Public Const gstrPatholCol_��������     As String = "��������"
Public Const gstrPatholCol_���λ��     As String = "���λ��"
Public Const gstrPatholCol_��ϸ��ַ     As String = "��ϸ��ַ"
Public Const gstrPatholCol_�黹��       As String = "�黹��"
Public Const gstrPatholCol_�˻�Ѻ��     As String = "�˻�Ѻ��"
Public Const gstrPatholCol_����ҽԺ     As String = "����ҽԺ"
Public Const gstrPatholCol_����ҽʦ     As String = "����ҽʦ"
Public Const gstrPatholCol_�������     As String = "�������"
Public Const gstrPatholCol_�Ǽ���       As String = "�Ǽ���"
Public Const gstrPatholCol_�ڵ�����     As String = "�ڵ�����"
Public Const gstrPatholCol_��ʧ����     As String = "��ʧ����"
Public Const gstrPatholCol_�ɽ�����     As String = "�ɽ�����"
Public Const gstrPatholCol_�ѽ�����     As String = "�ѽ�����"
Public Const gstrPatholCol_�������     As String = "�������"
Public Const gstrPatholCol_��������     As String = "��������"
Public Const gstrPatholCol_ʵ������     As String = "ʵ������"
Public Const gstrPatholCol_��������     As String = "��������"
Public Const gstrPatholCol_�黹����     As String = "�黹����"

Public Const gstrPatholCol_��������    As String = "��������"
Public Const gstrPatholCol_�������    As String = "�������"
Public Const gstrPatholCol_��������    As String = "��������"
Public Const gstrPatholCol_��������    As String = "��������"
Public Const gstrPatholCol_��������    As String = "��������"
Public Const gstrPatholCol_��鷶Χ    As String = "��鷶Χ"
Public Const gstrPatholCol_��ʼ����    As String = "��ʼ����"
Public Const gstrPatholCol_��������    As String = "��������"
Public Const gstrPatholCol_��������    As String = "��������"
Public Const gstrPatholCol_�������    As String = "�������"
Public Const gstrPatholCol_��������    As String = "��������"
Public Const gstrPatholCol_����˵��    As String = "����˵��"
Public Const gstrPatholCol_�鵵ʱ��    As String = "�鵵ʱ��"
Public Const gstrPatholCol_������      As String = "������"
Public Const gstrPatholCol_��������    As String = "��������"

Public Const gstrPatholCol_��ԴID        As String = "��ԴID"
Public Const gstrPatholCol_����ҽ��ID    As String = "����ҽ��ID"
Public Const gstrPatholCol_������Դ      As String = "������Դ"
Public Const gstrPatholCol_����          As String = "����"
Public Const gstrPatholCol_�����Ŀ      As String = "�����Ŀ"
Public Const gstrPatholCol_����          As String = "����"
Public Const gstrPatholCol_�鵵״̬      As String = "�鵵״̬"
Public Const gstrPatholCol_���״̬      As String = "���״̬"
Public Const gstrPatholCol_����ʱ��      As String = "����ʱ��"
Public Const gstrPatholCol_ִ�й���      As String = "ִ�й���"

Public Const gstrArchivesClass_ID           As String = "ID"
Public Const gstrArchivesClass_��������     As String = "��������"
Public Const gstrArchivesClass_��������     As String = "��������"
Public Const gstrArchivesClass_��������     As String = "��������"
Public Const gstrArchivesClass_������       As String = "������"
Public Const gstrArchivesClass_����ʱ��     As String = "����ʱ��"
Public Const gstrArchivesClass_��ע         As String = "��ע"



'������ʧ��ʾ��
Public Const gstrMaterialLoseCols As String = "|�������,read,merge|�����,merge,read,uncfg|ID,key,hide,uncfg|�Ŀ��>���,read,uncfg|�걾����,read|ȡ��λ��,read|������ϸ,read,uncfg|�ڵ�����,read|��ʧ����,read|��ʧԭ��,read|���״̬,read|"
Public Const gstrMaterialLoseConvertFormat As String = "���״̬:0-�浵��,1-������ʧ,2-����ʧ|����״̬:0-δ���,1-���ֽ��,2-�ѽ��"



'���Ĺ�����ʾ��
Public Const gstrMaterialBorrowCols As String = "|ID,hide,uncfg,key|���ĺ�,read,uncfg|������,read|��������>����ʱ��,read,w1600,onlydate|�黹����,read,w1600,onlydate|֤������,read|֤������,read|��ϵ�绰,read|��ϵ��ַ,read|Ѻ��,read|��������,read|��������,read|����ԭ��,read|�黹״̬,read|ȷ��״̬,read|��ע,read|"
Public Const gstrMaterialBorrowConvertFormat As String = "֤������:0-���֤,1-ѧ��֤,2-����֤,3-��ʻ֤,4-����,5-�籣��,6-�м�֤,7-����|��������:0-�ڲ�����,1-�ⲿ����|�黹״̬:0-δ�黹,1-�ѹ黹,2-���ֹ黹,3-��ʧ����|ȷ��״̬:0-δȷ��,1-��ȷ��"



'���Ĳ�����ϸ��
Public Const gstrMaterialBorrowDetailCols As String = "|�������,merge,read|�����,read,merge|�鵵ID,key,hide,uncfg|�Ŀ��>���,read,uncfg|�걾����,read|ȡ��λ��,read|�������,read|������ϸ,read,uncfg|��������,read|�黹����,read|��������>��������,read|���λ��,read|��ϸ��ַ,read|�黹״̬,read|"
Public Const gstrMaterialBorrowDetailConvertFormat As String = "�������:0-����,1-����,2-ϸ��,3-����,4-ʬ��,5-����ʯ��|�黹״̬:0-δ�黹,1-�ѹ黹,2-���ֹ黹,3-��ʧ,4-�����"



'���Ĳ��Ϲ黹��ʾ��
Public Const gstrMaterialBorrowReturnCols As String = "|�������,merge,read|�����,read,merge|�鵵ID,key,hide,uncfg|�Ŀ��>���,rowcheck,uncfg|ʵ������|��������,read|�걾����,read|ȡ��λ��,read|�������,read|������ϸ,read,uncfg|��������>��������,read|���λ��,read|��ϸ��ַ,read|�黹״̬,read|"
Public Const gstrMaterialBorrowReturnConvertFormat As String = "�������:0-����,1-����,2-ϸ��,3-����,4-ʬ��,5-����ʯ��|�黹״̬:0-δ�黹,1-�ѹ黹,2-���ֹ黹,3-��ʧ"



'���Ĺ黹��ʷ��
Public Const gstrMaterialBorrowBackCols As String = "|id,key,hide,uncfg|�黹��,read,uncfg|�黹����,read,w1600,onlydate|�˻�Ѻ��,read|����ҽԺ,read|����ҽʦ,read|�������,read,uncfg|�Ǽ���,read|��ע,read|"
Public Const gstrMaterialBorrowBackConvertFormat As String = ""



'���Ĳ�����ʾ�У����ĵǼǴ��ڣ� ����,merge,read|�Ա�,merge,read|
Public Const gstrMaterialBorrowEnregCols As String = "|�������,merge,read|�����,read,merge|����,merge,read|id,key,read,uncfg,hide|�Ŀ��>���,rowcheck,uncfg|�������|�ɽ�����,read|�걾����,read|ȡ��λ��,read|�������,read|������ϸ,read|��������>��������,read|���λ��,read,w2400|��ϸ��ַ,read,w1600|��ʧ����,read|�ѽ�����,read|���״̬,read|����״̬,read|"
Public Const gstrMaterialBorrowEnregConvertFormat As String = "�������:0-����,1-����,2-ϸ��,3-����,4-ʬ��,5-����ʯ��|���״̬:0-�浵��,1-������ʧ,2-����ʧ|����״̬:0-δ���,1-���ֽ��,2-�ѽ��"

Public Const gstrMaterialBorrowEnregedCols As String = "|�������,merge,read|�����,read,merge|id,key,read,uncfg,hide|�Ŀ��>���,rowcheck,uncfg|�걾����,read|ȡ��λ��,read|�������,read|������ϸ,read|��������,read|��������>��������,read|���λ��,read|��ϸ��ַ,read,w1600|"



'����������ʾ��
Public Const gstrArchivesManageCols As String = "|ID,hide,uncfg,key|��������,read,uncfg|�������,read|��������,read|��������,hide,read,uncfg|��������,hide,read,uncfg|��鷶Χ,read|��ʼ����,read,onlydate,w1600|��������,read,onlydate,w1600|��������,read|�������,read|��������,read|��ϸ��ַ,read,w2400|����˵��,read|����״̬,read|�鵵ʱ��,read,onlydate,w1600|������,read|��������,read,onlydate,w1600|"
Public Const gstrArchivesManageConvertFormat As String = "����״̬:0-δ�鵵,1-�ѹ鵵"



'�����������ѯϸ��ʾ�ж���
Public Const gstrArchivesMaterialCols As String = "|�������,merge,read|�������,merge,read|�����,merge,uncfg,read|����,merge,read,uncfg|�Ա�,merge,read|����,merge,read|" & _
                                                "�����Ŀ,merge,read|����ҽ��id,hide,uncfg|�Ŀ��>���,rowcheck,uncfg|�걾����,read|ȡ��λ��,read|������ϸ,read|" & _
                                                "����,read|��������>��������,read|���λ��,read,w2400|��ϸ��ַ,read,w1600|���״̬,read|����ʱ��,read,onlydate,w1600|ִ�й���,hide,uncfg|��ԴID,key,hide,uncfg|������Դ,hide,uncfg|"
Public Const gstrArchivesMaterialConvertFormat As String = "�������:0-����,1-����,2-ϸ��,3-����,4-ʬ��,5-����ʯ��"

'����������ϸ��ʾ�ж���
Public Const gstrArchivesMaterialDetailCols As String = "|�������,merge,read|�������,merge,read|�����,merge,uncfg,read|����,merge,read,uncfg|�Ա�,merge,read|����,merge,read|" & _
                                                "�����Ŀ,merge,read|����ҽ��id,hide,uncfg|�Ŀ��>���,rowcheck,uncfg|�걾����,read|ȡ��λ��,read|������ϸ,read|" & _
                                                "����,read|���״̬,read|����״̬,read|����ʱ��,read,onlydate,w1600|ִ�й���,hide,uncfg|��ԴID,key,hide,uncfg|������Դ,hide,uncfg|"


'�����൵����ϸ��ʾ�ж���
Public Const gstrArchivesWordCols As String = "|��ԴID,key,hide,uncfg|������Դ,hide,uncfg|�����,rowcheck,uncfg|����,read,uncfg|�Ա�,read|����,read|" & _
                                                "�����Ŀ,read|�������,read|����ʱ��,read,onlydate,w1600|ִ�й���,hide,uncfg|����ҽ��id,hide,uncfg|���״̬,read,uncfg|"
Public Const gstrArchivesWordConvertFormat As String = "�������:0-����,1-����,2-ϸ��,3-����,4-ʬ��,5-����ʯ��"



'������������
Public Const gstrArchivesClassCols As String = "|ID,hide,key,uncfg|��������,uncfg,w1800|��������,cbx<0-������,1-������,2-�������>|��������, w1600|��ע,w2400|������,w1800,read|����ʱ��,onlydate,w1800,read|"
Public Const gstrArchivesClassConvertFormat As String = "��������:0-������,1-������,2-�������"



'���걾�����б���
Public Const gstrSpecimenModuleCols As String = "|ID,hide,key|�걾����,w1800|�걾��λ|�걾����,cbx<0-�����걾,1-����ϸ��,2-����ϸ��,3-Һ��ϸ��>,w1800|Ĭ�ϱ걾��,w1800|Ĭ����Ƭ��,w1800|����|��ע,w2400|"
Public Const gstrSpecimenModuleConvertFormat As String = "�걾����:0-�����걾,1-����ϸ��,2-����ϸ��,3-Һ��ϸ��"

Public Const gstrSpecimenModule_ID         As String = "ID"
Public Const gstrSpecimenModule_�걾����   As String = "�걾����"
Public Const gstrSpecimenModule_�걾��λ   As String = "�걾��λ"
Public Const gstrSpecimenModule_�걾����   As String = "�걾����"
Public Const gstrSpecimenModule_Ĭ�ϱ걾�� As String = "Ĭ�ϱ걾��"
Public Const gstrSpecimenModule_Ĭ����Ƭ�� As String = "Ĭ����Ƭ��"
Public Const gstrSpecimenModule_����       As String = "����"
Public Const gstrSpecimenModule_��ע       As String = "��ע"



'�걾��ʾ����
'Public Const gstrSpecimenCols As String = "|�걾ID,hide,key|�ͼ�ID,hide|�걾����|�걾����,cbx<0-�����걾,1-С�걾,2-����ϸ��,3-����ϸ��,4-Һ��ϸ��>|�ɼ���λ|�걾����>����|�������,cbx<0-�걾,1-����,2-��Ƭ,3-��Ƭ,4-����>|" & _
'                                        "���λ��|ԭ�б��|��ע|��������,fulldatetime,read,w2400|����״̬,read|"
'Public Const gstrSpecimenConvertFormat As String = "�걾����:0-�����걾,1-С�걾,2-����ϸ��,3-����ϸ��,4-Һ��ϸ��|�������:0-�걾,1-����,2-��Ƭ,3-��Ƭ,4-����"

Public Const gstrSpecimenCols As String = "|�걾ID,hide,key,uncfg|�ͼ�ID,hide,uncfg|�걾����,uncfg|�걾����,cbx<0-�����걾,1-����ϸ��,2-����ϸ��,3-Һ��ϸ��>|�ɼ���λ|�걾����>����|�������,cbx<0-�걾,1-����,2-��Ƭ,3-��Ƭ,4-����>|" & _
                                        "���λ��|ԭ�б��|��ע|��������,fulldatetime,read,w2400|����״̬,read|"
                                        
Public Const gstrSpecimenConvertFormat As String = "�걾����:0-�����걾,1-����ϸ��,2-����ϸ��,3-Һ��ϸ��|�������:0-�걾,1-����,2-��Ƭ,3-��Ƭ,4-����"


Public Const gSpecimen_�걾ID   As String = "�걾ID"
Public Const gSpecimen_�ͼ�ID   As String = "�ͼ�ID"
Public Const gSpecimen_�걾���� As String = "�걾����"
Public Const gSpecimen_�걾���� As String = "�걾����"
Public Const gSpecimen_�ɼ���λ As String = "�ɼ���λ"
Public Const gSpecimen_����     As String = "�걾����"
Public Const gSpecimen_������� As String = "�������"
Public Const gSpecimen_���λ�� As String = "���λ��"
Public Const gSpecimen_ԭ�б�� As String = "ԭ�б��"
Public Const gSpecimen_�������� As String = "��������"
Public Const gSpecimen_��ע     As String = "��ע"
Public Const gSpecimen_����״̬ As String = "����״̬"



'�����������
Public Const gstrPatholQualityCols As String = "|ID,hide,key,uncfg|������Ŀ,cbx< ,�걾����,��Ƭ����,�������,�����黯,����Ⱦɫ,���Ӳ���>,uncfg|���۽��,cbx< ,��,��,��,��>,uncfg|�������,w2000|�Ľ�����,w2000|��ע,w2000|������,read|����ʱ��,w1900,fulldatetime,read|"
                                        
Public Const gstrPatholQualityConvertFormat As String = ""


Public Const gstrPatholQuality_ID       As String = "ID"
Public Const gstrPatholQuality_�����   As String = "�����"
Public Const gstrPatholQuality_������Ŀ As String = "������Ŀ"
Public Const gstrPatholQuality_���۽�� As String = "���۽��"
Public Const gstrPatholQuality_������� As String = "�������"
Public Const gstrPatholQuality_�Ľ����� As String = "�Ľ�����"
Public Const gstrPatholQuality_��ע     As String = "��ע"
Public Const gstrPatholQuality_������   As String = "������"
Public Const gstrpatholQuality_����ʱ�� As String = "����ʱ��"


'============================================================================================================================



'����ȡ����ʾ��
Public Const gstrNormalMaterialCols As String = "|�Ŀ�ID,key,read,w1000,hide,uncfg|�Ŀ��>���,read,w800,align<2,0>|ȡ��λ��,w2400|�걾����,uncfg|��״|������|��Ƭ��|�Ƿ�����,cbx<0-�� ,1-��>,uncfg|�Ƿ��Ѹ�,cbx<0-�� ,1-��>|" & _
                                                "��ȡҽʦ,uncfg|��ȡҽʦ|ȡ��ʱ��,fulldatetime,w2400,uncfg|��¼ҽʦ,read|ȡ������,read|ȷ��״̬,read|"

'ϸ��ȡ����ʾ��
Public Const gstrCellMaterialCols As String = "|�Ŀ�ID,key,,read,w1000,hide,uncfg|�Ŀ��>���,read,w800,align<2,0>|�걾����,uncfg|����|��ɫ|�걾��,cbx<, ml, 1��, 2��, 4��, 8��>,uncfg|ϸ������>������|�Ƿ�����,cbx<0-�� ,1-��>,uncfg|��Ƭ��|" & _
                                                "��ȡҽʦ,uncfg|��ȡҽʦ|ȡ��ʱ��,fulldatetime,w2400,uncfg|��¼ҽʦ,read|ȡ������,read|ȷ��״̬,read|"

'����ȡ����ʾ��
Public Const gstrIceMaterialCols As String = "|�Ŀ�ID,key,,read,w1000,hide,uncfg|�Ŀ��>���,read,w800,align<2,0>|ȡ��λ��,w2400|�걾����,uncfg|��״|�Ƿ�����,cbx<0-�� ,1-��>|�Ƿ����,cbx<0-��,1-��>,uncfg|������|��Ƭ��|" & _
                                                "��ȡҽʦ,uncfg|��ȡҽʦ|ȡ��ʱ��,fulldatetime,w2400,uncfg|��¼ҽʦ,read|ȡ������,read|ȷ��״̬,read|"
                                                
Public Const gstrMaterialConvertFormat As String = "�Ƿ����:0-��,1-��|�Ƿ��Ѹ�:0-�� ,1-��|ȷ��״̬:0-δȷ��,1-��ȷ��|�Ƿ�����:0-��,1-��"
                                                
                                                
Public Const gstrMaterial_�Ŀ�ID   As String = "�Ŀ�ID"
Public Const gstrMaterial_�Ŀ��     As String = "�Ŀ��"
Public Const gstrMaterial_�걾���� As String = "�걾����"
Public Const gstrMaterial_ȡ��λ�� As String = "ȡ��λ��"
Public Const gstrMaterial_��״     As String = "��״"
Public Const gstrMaterial_������   As String = "������"
Public Const gstrMaterial_��Ƭ��   As String = "��Ƭ��"
Public Const gstrMaterial_ȡ������ As String = "ȡ������"
Public Const gstrMaterial_ȡ��ʱ�� As String = "ȡ��ʱ��"
Public Const gstrMaterial_��ȡҽʦ As String = "��ȡҽʦ"
Public Const gstrMaterial_��ȡҽʦ As String = "��ȡҽʦ"
Public Const gstrMaterial_��¼ҽʦ As String = "��¼ҽʦ"
Public Const gstrMaterial_����     As String = "����"
Public Const gstrMaterial_��ɫ     As String = "��ɫ"
Public Const gstrMaterial_�걾��   As String = "�걾��"
Public Const gstrMaterial_ϸ������ As String = "ϸ������"
Public Const gstrMaterial_�Ƿ���� As String = "�Ƿ����"
Public Const gstrMaterial_�Ƿ��Ѹ� As String = "�Ƿ��Ѹ�"
Public Const gstrMaterial_�Ƿ����� As String = "�Ƿ�����"
Public Const gstrMaterial_ȷ��״̬ As String = "ȷ��״̬"

                                                
'============================================================================================================================


'�Ѹ���ʾ��
Public Const gstrDecalinCols As String = "|ID,key,hide,uncfg|�걾����,uncfg|��ʼʱ��,fulldatetime,w2400,uncfg|����ʱ��(Сʱ)>����ʱ��,w1600,uncfg|����Ա,uncfg|����ʱ��,fulldatetime,read,w2400|��ǰ�״�,read|��ǰ״̬>���״̬,read|"
Public Const gstrDecalinTaskCols As String = "|ID,key,hide,uncfg|�����|�걾����,uncfg|��ʼʱ��,fulldatetime,w2400,uncfg|����ʱ��(Сʱ)>����ʱ��,w1600|ʣ��ʱ��(��)>ʣ��ʱ��,w1200|����ʱ��,fulldatetime,read,w2400,uncfg|��ǰ�״�,read|����Ա,read,uncfg|��ǰ״̬>���״̬,read|"

Public Const gstrDecalinConvertFormat As String = "��ǰ״̬:0-������,1-�����"

Public Const gstrDecalin_ID       As String = "ID"
Public Const gstrDecalin_�걾ID   As String = "�걾ID"
Public Const gstrDecalin_�걾���� As String = "�걾����"
Public Const gstrDecalin_��ʼʱ�� As String = "��ʼʱ��"
Public Const gstrDecalin_����ʱ�� As String = "����ʱ��(Сʱ)"
Public Const gstrDecalin_ʣ��ʱ�� As String = "ʣ��ʱ��(��)"
Public Const gstrDecalin_����ʱ�� As String = "����ʱ��"
Public Const gstrDecalin_��ǰ�״� As String = "��ǰ�״�"
Public Const gstrDecalin_����Ա   As String = "����Ա"
Public Const gstrDecalin_��ǰ״̬ As String = "��ǰ״̬"

 
'============================================================================================================================



'��Ƭ��ʾ��
Public Const gstrSlicesCols As String = "|��ƬID>ID,key,w1000,hide,uncfg|�Ŀ�ID,hide,uncfg|�Ŀ��>���,w1000,align<2,0>,uncfg|ȡ��λ��|�걾����,uncfg|��Ƭ��|��Ƭ����|��Ƭ��ʽ|��Ƭʱ��,fulldatetime,w2400|��Ƭ��ʦ|��ǰ״̬|�嵥״̬|"
Public Const gstrSlicesConvertFormat As String = "��Ƭ����:0-ʯ����Ƭ,1-������Ƭ,2-ϸ����Ƭ|��Ƭ��ʽ:0-����,1-����,2-����,3-����,4-��Ƭ,5-��Ⱦ,6-��Ƭ|��ǰ״̬:0-δ����,1-�ѽ���,2-�����|�嵥״̬:0-δ��ӡ,1-�Ѵ�ӡ"


Public Const gstrSlices_�Ŀ�ID     As String = "�Ŀ�ID"
Public Const gstrSlices_�Ŀ��     As String = "�Ŀ��"
Public Const gstrSlices_�걾����   As String = "�걾����"
Public Const gstrSlices_��Ƭ��     As String = "��Ƭ��"
Public Const gstrSlices_��Ƭ����   As String = "��Ƭ����"
Public Const gstrSlices_��Ƭʱ��   As String = "��Ƭʱ��"
Public Const gstrSlices_��Ƭ��     As String = "��Ƭ��ʦ"
Public Const gstrSlices_��ǰ״̬   As String = "��ǰ״̬"
Public Const gstrSlices_�嵥״̬   As String = "�嵥״̬"


'��Ƭ������ʾ��
Public Const gstrSlicesQualityCols As String = "|�����,w1200,read|��Ƭ����,w1000,read|�걾����,w1500,read|ȡ��λ��,w1500,read|�Ŀ��,w800,read|ID,key,w1000,hide|��Դ����,hide|��ԴId,hide|�Ŀ�ID,hide|��Ƭ����,w1000,cbx<,��,��,��,��>|������,w1100,read|��������,read,onlydate|"
Public Const gstrSlicesQualityConvertFormat As String = ""


Public Const gstrSlicesQuality_��ƬID     As String = "��ƬID"
Public Const gstrSlicesQuality_�Ŀ�ID     As String = "�Ŀ�ID"
Public Const gstrSlicesQuality_�걾����   As String = "�걾����"
Public Const gstrSlicesQuality_��Ƭ��ʽ   As String = "��Ƭ��ʽ"
Public Const gstrSlicesQuality_��Ƭ���   As String = "��Ƭ���"
Public Const gstrSlicesQuality_��Ƭ����   As String = "��Ƭ����"
Public Const gstrSlicesQuality_��ע       As String = "��ע"
Public Const gstrSlicesQuality_������     As String = "������"
Public Const gstrSlicesQuality_����ʱ��   As String = "����ʱ��"



'��Ƭ�����嵥��ʾ��
Public Const gstrSlicesWorkCols As String = "|�����,rowcheck,merge,w1600,uncfg|����ҽ��ID,hide,uncfg|�������,merge|����,merge|��ƬID>ID,key,w1000,hide,uncfg|�Ŀ�ID,hide,uncfg|�Ŀ��>���,w1000,align<2,0>,uncfg|ȡ��λ��|�걾����,w1600,uncfg|�걾����|��Ƭ����|��Ƭ��ʽ|��Ƭ��|ȡ��ʱ��,fulldatetime,w2400,uncfg|��ǰ״̬|�嵥״̬|"
Public Const gstrSlicesWorkConvertFormat As String = "�����:els-<check><source>|�������:0-����,1-����,2-ϸ��,3-����,4-ʬ��,5-����ʯ��|�걾����:0-���α걾,1-С�걾,2-����ϸ��,3-����ϸ��,4-Һ��ϸ��|��Ƭ����:0-ʯ����Ƭ,1-������Ƭ,2-ϸ����Ƭ|��Ƭ��ʽ:0-����,1-����,2-����,3-����,4-��Ƭ,5-��Ⱦ,6-��Ƭ|��ǰ״̬:0-δ����,1-�ѽ���,2-�����|�嵥״̬:0-δ��ӡ,1-�Ѵ�ӡ"



Public Const gstrSlicesWork_�����    As String = "�����"
Public Const gstrSlicesWork_����ҽ��ID    As String = "����ҽ��ID"
Public Const gstrSlicesWork_����      As String = "����"
Public Const gstrSlicesWork_�������  As String = "�������"
Public Const gstrSlicesWork_�Ŀ�ID    As String = "�Ŀ�ID"
Public Const gstrSlicesWork_�Ŀ��  As String = "�Ŀ��"
Public Const gstrSlicesWork_�걾����  As String = "�걾����"
Public Const gstrSlicesWork_�걾����  As String = "�걾����"
Public Const gstrSlicesWork_��Ƭ����  As String = "��Ƭ����"
Public Const gstrSlicesWork_��Ƭ��    As String = "��Ƭ��"
Public Const gstrSlicesWork_��ǰ״̬  As String = "��ǰ״̬"
Public Const gstrSlicesWork_�嵥״̬  As String = "�嵥״̬"




'��Ƭȷ����ʾ��
Public Const gstrSlicesSureColsWithMaterialNum As String = "|�����,rowcheck,merge,w1600|����,merge,read|�������,merge,read|��ƬId>ID,read,w1000,hide|�Ŀ�ID,hide|�Ŀ��>���,w1000,read,align<2,0>|�걾����,read|��Ƭ����,read|��Ƭ��ʽ,read|����Ƭ��,read|��ȷ����|��ǰ״̬,read|����ҽ��ID,key,hide|"
Public Const gstrSlicesSureConvertFormat = "�����:els-<check><source>|�������:0-����,1-����,2-ϸ��,3-����,4-ʬ��,5-����ʯ��|��Ƭ����:0-ʯ����Ƭ,1-������Ƭ,2-ϸ����Ƭ|��Ƭ��ʽ:0-����,1-����,2-����,3-����,4-��Ƭ,5-��Ⱦ,6-��Ƭ|��ǰ״̬:0-δ����,1-�ѽ���,2-�����"


Public Const gstrSlicesSure_ID         As String = "��ƬID"
Public Const gstrSlicesSure_�����     As String = "�����"
Public Const gstrSlicesSure_����       As String = "����"
Public Const gstrSlicesSure_�������   As String = "�������"
Public Const gstrSlicesSure_��Ƭ״̬   As String = "��Ƭ״̬"
Public Const gstrSlicesSure_�Ŀ�ID     As String = "�Ŀ�ID"
Public Const gstrSlicesSure_�Ŀ��   As String = "�Ŀ��"
Public Const gstrSlicesSure_�걾����   As String = "�걾����"
Public Const gstrSlicesSure_��ǰ״̬   As String = "��ǰ״̬"
Public Const gstrSlicesSure_����Ƭ��   As String = "����Ƭ��"
Public Const gstrSlicesSure_��ȷ����   As String = "��ȷ����"
Public Const gstrSlicesSure_ȷ��״̬   As String = "ȷ��״̬"



'============================================================================================================================

'������Ϣ��ʾ��
Public Const gstrAntibodyCols As String = "|����ID,key,hide,uncfg|��������,uncfg|ʹ���˷�|�����˷�|��������,onlydate,w1600|��Ч��|��������,onlydate,w1600|��¡��|���ö���|������|Ӧ�����|ʹ��״̬|�Ǽ���,uncfg|�Ǽ�ʱ��,fulldatetime,w2400,uncfg|��ע|"
Public Const gstrAntibodyConvertFormat As String = "��¡��:0-����¡��Ũ���ͣ�,1-����¡�������ͣ�,2-���¡��Ũ���ͣ�,3-���¡�������ͣ�|ʹ��״̬:0-��ֹͣ,1-ʹ����"



Public Const gstrAntibody_����ID   As String = "����ID"
Public Const gstrAntibody_�������� As String = "��������"
Public Const gstrAntibody_ʹ���˷� As String = "ʹ���˷�"
Public Const gstrAntibody_�����˷� As String = "�����˷�"
Public Const gstrAntibody_�������� As String = "��������"
Public Const gstrAntibody_��Ч��   As String = "��Ч��"
Public Const gstrAntibody_�������� As String = "��������"
Public Const gstrAntibody_��¡��   As String = "��¡��"
Public Const gstrAntibody_���ö��� As String = "���ö���"
Public Const gstrAntibody_������ As String = "������"
Public Const gstrAntibody_Ӧ����� As String = "Ӧ�����"
Public Const gstrAntibody_ʹ��״̬ As String = "ʹ��״̬"
Public Const gstrAntibody_�Ǽ���   As String = "�Ǽ���"
Public Const gstrAntibody_�Ǽ�ʱ�� As String = "�Ǽ�ʱ��"
Public Const gstrAntibody_��ע     As String = "��ע"
        
        
'���巴����Ϣ��ʾ��
Public Const gstrAntibodyFeedbackCols As String = "|ID,key,hide,uncfg|�ο������,w2400|ʵ������|��������|�������,w3200,uncfg|����ҽ��,uncfg|����ʱ��,fulldatetime,w2400,uncfg|"
Public Const gstrAntibodyFeedbackConvertFormat As String = "ʵ������:0-�����黯,1-����Ⱦɫ,2-���Ӳ���,3-����"


Public Const gstrAntibodyFeedback_ID         As String = "ID"
Public Const gstrAntibodyFeedback_�ο������ As String = "�ο������"
Public Const gstrAntibodyFeedback_ʵ������   As String = "ʵ������"
Public Const gstrAntibodyFeedback_��������   As String = "��������"
Public Const gstrAntibodyFeedback_�������   As String = "�������"
Public Const gstrAntibodyFeedback_����ҽ��   As String = "����ҽ��"
Public Const gstrAntibodyFeedback_����ʱ��   As String = "����ʱ��"
        
        
        
'============================================================================================================================


'�ײ���Ϣ��ʾ��
Public Const gstrAntibodyMealCols As String = "|�ײ�ID,key,hide,uncfg|�ײ�����,uncfg|�ײ����|�ײ�˵��,w3200|����ʱ��,fulldatetime,read,w2400|������,read|"
Public Const gstrAntibodyMealConvertFormat As String = ""


Public Const gstrAntibodyMeal_�ײ�ID   As String = "�ײ�ID"
Public Const gstrAntibodyMeal_�ײ����� As String = "�ײ�����"
Public Const gstrAntibodyMeal_�ײ���� As String = "�ײ����"
Public Const gstrAntibodyMeal_�ײ�˵�� As String = "�ײ�˵��"
Public Const gstrAntibodyMeal_����ʱ�� As String = "����ʱ��"
Public Const gstrAntibodyMeal_������   As String = "������"


'�ײͿ�����ϸ��ʾ��
Public Const gstrAntibodyMealLinkCols As String = "|����ID,hide,uncfg|����ID,key,hide,uncfg|��������,rowcheck,uncfg,w1200|��¡��,read,w1700|������,read|���ö���,read|Ӧ�����,read,w2400|��ע,read|����˳��,hide,uncfg|"
Public Const gstrAntibodyMealLinkConvertFormat As String = "��¡��:0-����¡��Ũ���ͣ�,1-����¡�������ͣ�,2-���¡��Ũ���ͣ�,3-���¡�������ͣ�"


Public Const gstrAntibodyMealLink_����ID   As String = "����ID"
Public Const gstrAntibodyMealLink_����ID   As String = "����ID"
Public Const gstrAntibodyMealLink_�������� As String = "��������"
Public Const gstrAntibodyMealLink_��¡��   As String = "��¡��"
Public Const gstrAntibodyMealLink_������ As String = "������"
Public Const gstrAntibodyMealLink_���ö��� As String = "���ö���"
Public Const gstrAntibodyMealLink_Ӧ����� As String = "Ӧ�����"
Public Const gstrAntibodyMealLink_��ע     As String = "��ע"
Public Const gstrAntibodyMealLink_����˳�� As String = "����˳��"

'============================================================================================================================


'�����ӳ���ʾ��
Public Const gstrReportDelayCols As String = "|ID,key,hide,uncfg|�ӳ�ԭ��,btn,w3200,uncfg|�ӳ�����,uncfg|��ʱ���,w3200|ת����|�Ǽ�ʱ��,fulldatetime,read,w2400|�Ǽ���,read|��ǰ״̬,read|"
Public Const gstrReportDelayConvertFormat As String = "��ǰ״̬:0-δ��ӡ,1-�Ѵ�ӡ"

Public Const gstrReportDelay_ID       As String = "ID"
Public Const gstrReportDelay_�����   As String = "�����"
Public Const gstrReportDelay_�ӳ�ԭ�� As String = "�ӳ�ԭ��"
Public Const gstrReportDelay_�ӳ����� As String = "�ӳ�����"
Public Const gstrReportDelay_��ʱ��� As String = "��ʱ���"
Public Const gstrReportDelay_ת����   As String = "ת����"
Public Const gstrReportDelay_�Ǽ���   As String = "�Ǽ���"
Public Const gstrReportDelay_�Ǽ�ʱ�� As String = "�Ǽ�ʱ��"
Public Const gstrReportDelay_��ǰ״̬ As String = "��ǰ״̬"


'============================================================================================================================


'���̱�����ʾ��
Public Const gstrProcedureRepCols As String = "|ID,key,hide,uncfg|����ͼ��,hide,uncfg|�걾����,uncfg|��������,uncfg|��������|�����,hide,uncfg|������,hide,uncfg|������>����ҽʦ,uncfg|��������,fulldatetime, w2400|��ǰ״̬|��ע,hide,uncfg|"
Public Const gstrProcedureRepConvertFormat = "��������:0-��������,1-���߱���,2-���ӱ���,3-��Ⱦ����|��������:0-��,1-����,2-��ҩ��ҩ,3-ӫ��,4-��ͨ|��ǰ״̬:0-δ��ӡ,1-�Ѳ���,2-�ѳ���,3-�Ѵ�ӡ"

Public Const gstrProcedureRep_ID       As String = "ID"
Public Const gstrProcedureRep_����ͼ�� As String = "����ͼ��"
Public Const gstrProcedureRep_�걾���� As String = "�걾����"
Public Const gstrProcedureRep_�������� As String = "��������"
Public Const gstrProcedureRep_�������� As String = "��������"
Public Const gstrProcedureRep_����� As String = "�����"
Public Const gstrProcedureRep_������ As String = "������"
Public Const gstrProcedureRep_������   As String = "������"
Public Const gstrProcedureRep_�������� As String = "��������"
Public Const gstrProcedureRep_��ǰ״̬ As String = "��ǰ״̬"
Public Const gstrProcedureRep_��ע     As String = "��ע"


'============================================================================================================================


'������ʾ��
Public Const gstrRequisitionCols As String = "|����ID,key,hide,uncfg|������,uncfg|��������,uncfg|����״̬|����ϸĿ|����ʱ��,fulldatetime,w2400|��ǰ״̬>����״̬|��������,w3200|���ʱ��,fulldatetime,w2400|"
Public Const gstrRequisitionViewCols As String = "|����ID,key,hide,uncfg|������,uncfg|����ʱ��,fulldatetime,w2400|����ϸĿ,uncfg|��ǰ״̬>����״̬|��������,w3200|����״̬|���ʱ��,fulldatetime,w2400|"
Public Const gstrRequisitionConvertFormat As String = "��������:0-�����黯,1-����Ⱦɫ,2-���Ӳ���,3-����Ƭ,4-��ȡ��|����״̬:0-��,1-�貹��,2-�Ѳ���|����ϸĿ:0-��,1-����,2-��ҩ��ҩ,3-ӫ��,4-��ͨ|��ǰ״̬:0-������,1-�ѽ���,2-�����"


Public Const gstrRequisition_����ID   As String = "����ID"
Public Const gstrRequisition_������   As String = "������"
Public Const gstrRequisition_�������� As String = "��������"
Public Const gstrRequisition_����״̬ As String = "����״̬"
Public Const gstrRequisition_����ϸĿ As String = "����ϸĿ"
Public Const gstrRequisition_����ʱ�� As String = "����ʱ��"
Public Const gstrRequisition_��ǰ״̬ As String = "��ǰ״̬"
Public Const gstrRequisition_�������� As String = "��������"


'�ؼ�����������ϸ��ʾ��
Public Const gstrRequest_SpeExam_Cols As String = "|ID,key,hide,read,uncfg|����ID,hide,read,uncfg|�Ŀ�ID,hide,read,uncfg|�Ŀ��>���,w1000,read,align<2,0>,uncfg|�걾����,read,w1600,uncfg|��������,btn,read,uncfg|��������,read|��ǰ״̬,read|��Ŀ���,read|���ʱ��,fulldatetime,read,w2400|������>�ؼ�ҽʦ,read|"
Public Const gstrRequest_SpeExamConvertFormat As String = "��������:-1-����,0-����,els-��<source>������|��ǰ״̬:0-������,1-�ѽ���,2-�����"


Public Const gstrRequest_SpeExam_ID       As String = "ID"
Public Const gstrRequest_SpeExam_�Ŀ�� As String = "�Ŀ��"
Public Const gstrRequest_SpeExam_�걾���� As String = "�걾����"
Public Const gstrRequest_SpeExam_����ID   As String = "����ID"
Public Const gstrRequest_SpeExam_�������� As String = "��������"
Public Const gstrRequest_SpeExam_�������� As String = "��������"
Public Const gstrRequest_SpeExam_��ǰ״̬ As String = "��ǰ״̬"
Public Const gstrRequest_SpeExam_��Ŀ��� As String = "��Ŀ���"
Public Const gstrRequest_SpeExam_���ʱ�� As String = "���ʱ��"
Public Const gstrRequest_SpeExam_������   As String = "������"


'��Ƭ����������ϸ��ʾ��
Public Const gstrRequest_Slices_Cols As String = "|ID,key,hide,read,uncfg|�Ŀ�ID,hide,read,uncfg|�Ŀ��>���,w1000,read,align<2,0>,uncfg|�걾����,read,uncfg|��Ƭ����,read|��Ƭ��ʽ,read|��Ƭ����>��Ƭ��,read|��ǰ״̬,read|��Ƭʱ��,fulldatetime,read,w2400|��Ƭ��,read|"
Public Const gstrRequest_SlicesConvertFormat As String = "��Ƭ����:0-ʯ����Ƭ,1-������Ƭ,2-ϸ����Ƭ|��Ƭ��ʽ:0-����,1-����,2-����,3-����,4-��Ƭ,5-��Ⱦ,6-��Ƭ|��ǰ״̬:0-������,1-�ѽ���,2-�����"


Public Const gstrRequest_Slices_ID       As String = "ID"
Public Const gstrRequest_Slices_�Ŀ��   As String = "�Ŀ��"
Public Const gstrRequest_Slices_�걾���� As String = "�걾����"
Public Const gstrRequest_Slices_��Ƭ���� As String = "��Ƭ����"
Public Const gstrRequest_Slices_��Ƭ��ʽ As String = "��Ƭ��ʽ"
Public Const gstrRequest_Slices_��Ƭ���� As String = "��Ƭ����"
Public Const gstrRequest_Slices_��ǰ״̬ As String = "��ǰ״̬"
Public Const gstrRequest_Slices_��Ƭʱ�� As String = "��Ƭʱ��"
Public Const gstrRequest_Slices_��Ƭ��   As String = "��Ƭ��"



'��ȡ��������������ʾ��
Public Const gstrRequest_Material_Cols As String = "|�Ŀ�ID,hide,key,read,uncfg|�Ŀ��>���,w1000,read,align<2,0>,uncfg|�걾����,read,uncfg|�걾��,read|������,read|ȡ��ʱ��,fulldatetime,read,w2400|��ȡҽʦ,read|��ȡҽʦ,read|��¼ҽʦ,read|"
Public Const gstrRequest_MaterialConvertFormat As String = ""



Public Const gstrRequest_Material_�Ŀ��   As String = "�Ŀ��"
Public Const gstrRequest_Material_�걾���� As String = "�걾����"
Public Const gstrRequest_Material_�걾��   As String = "�걾��"
Public Const gstrRequest_Material_������   As String = "������"
Public Const gstrRequest_Material_ȡ��ʱ�� As String = "ȡ��ʱ��"
Public Const gstrRequest_Material_��ȡҽʦ As String = "��ȡҽʦ"
Public Const gstrRequest_Material_��ȡҽʦ As String = "��ȡҽʦ"
Public Const gstrRequest_Material_��¼ҽʦ As String = "��¼ҽʦ"



'�ؼ�����Ŀ�����Ϣ��ʾ��
Public Const gstrRequestAntibodyCols As String = "|����ID,key,hide,uncfg|��������,rowcheck,btn,w1600,uncfg|ʹ���˷�,read|�����˷�,read|��������,onlydate,,read,w1600|��Ч��,read|��������,onlydate,read,w1600|��Ŀ˳��,hide,uncfg|"
Public Const gstrRequestAntibodyConvertFormat As String = ""

Public Const gstrRequestAntibody_����ID   As String = "����ID"
Public Const gstrRequestAntibody_�������� As String = "��������"
Public Const gstrRequestAntibody_ʹ���˷� As String = "ʹ���˷�"
Public Const gstrRequestAntibody_�����˷� As String = "�����˷�"
Public Const gstrRequestAntibody_�������� As String = "��������"
Public Const gstrRequestAntibody_��Ч��   As String = "��Ч��"
Public Const gstrRequestAntibody_�������� As String = "��������"
Public Const gstrRequestAntibody_��Ŀ˳�� As String = "��Ŀ˳��"




'============================================================================================================================



'��������������ϸ��ʾ��
Public Const gstrConsultationCols As String = "|ID,key,hide,uncfg|����ҽʦ,uncfg|���ﵥλ|����ҽʦ,uncfg|��������|����ʱ��,shortdatetime,w2400|��ֹʱ��,shortdatetime,w2400|�������>�������,w2400|��Ͻ��,w3200,uncfg|������,w3200,uncfg|��ǰ״̬|���ʱ��,fulldatetime,w2400|��ע,w3200|"
Public Const gstrConsultationConvertFormat As String = "��������:0-���ڻ���,1-Ժ�����|��ǰ״̬:0-������,1-�ѳ���,2-�ѷ���,3-�Ѳ���"


Public Const gstrConsultation_ID       As String = "ID"
Public Const gstrConsultation_����ҽʦ As String = "����ҽʦ"
Public Const gstrConsultation_���ﵥλ As String = "���ﵥλ"
Public Const gstrConsultation_����ҽʦ As String = "����ҽʦ"
Public Const gstrConsultation_�������� As String = "��������"
Public Const gstrConsultation_����ʱ�� As String = "����ʱ��"
Public Const gstrConsultation_��ֹʱ�� As String = "��ֹʱ��"
Public Const gstrConsultation_������� As String = "�������"
Public Const gstrConsultation_��Ͻ�� As String = "��Ͻ��"
Public Const gstrConsultation_������ As String = "������"
Public Const gstrConsultation_��ǰ״̬ As String = "��ǰ״̬"
Public Const gstrConsultation_���ʱ�� As String = "���ʱ��"
Public Const gstrConsultation_��ע     As String = "��ע"





'============================================================================================================================



'�ؼ���Ϣ��ʾ��
Public Const gstrSpeExamCols As String = "|�ؼ�ID>ID,key,read,w1000,hide,uncfg|�Ŀ�ID,hide,uncfg|�Ŀ��>���,read,w1000,align<2,0>,uncfg|�걾����,read,uncfg|����ID,hide,uncfg|����ID,hide,uncfg|��������,btn,read,uncfg|�ؼ�ϸĿ,read|��������,read|�ؼ켼ʦ,read|����ʱ��,fulldatetime,read,w2400|���ʱ��,fulldatetime,read,w2400|��ǰ״̬,read|�嵥״̬,read|�ؼ�����,hide,uncfg|"
Public Const gstrSpeExamConvertFormat = "��������:-1-����,0-����,els-��<source>������|��ǰ״̬:0-������,1-�ѽ���,2-�����|�ؼ�ϸĿ:0-��,1-����,2-��ҩ��ҩ,3-ӫ��,4-��ͨ|�嵥״̬:0-δ��ӡ,1-�Ѵ�ӡ"


Public Const gstrSpeExam_ID       As String = "ID"
Public Const gstrSpeExam_�Ŀ�ID   As String = "�Ŀ�ID"
Public Const gstrSpeExam_�Ŀ�� As String = "�Ŀ��"
Public Const gstrSpeExam_�걾���� As String = "�걾����"
Public Const gstrSpeExam_����ID   As String = "����ID"
Public Const gstrSpeExam_����ID   As String = "����ID"
Public Const gstrSpeExam_�������� As String = "��������"
Public Const gstrSpeExam_�ؼ�ϸĿ As String = "�ؼ�ϸĿ"
Public Const gstrSpeExam_�������� As String = "��������"
Public Const gstrSpeExam_��ǰ״̬ As String = "��ǰ״̬"
Public Const gstrSpeExam_��Ŀ��� As String = "��Ŀ���"
Public Const gstrSpeExam_����ʱ�� As String = "����ʱ��"
Public Const gstrSpeExam_���ʱ�� As String = "���ʱ��"
Public Const gstrSpeExam_�ؼ�ҽʦ As String = "�ؼ켼ʦ"
Public Const gstrSpeExam_�嵥״̬ As String = "�嵥״̬"
Public Const gstrSpeExam_�ؼ����� As String = "�ؼ�����"


'�ؼ칤���嵥��ʾ��
Public Const gstrSpeExamWorkCols As String = "|�����,rowcheck,merge,w1600,uncfg|����ҽ��ID,hide,uncfg|�������,merge|����,merge|�ؼ�ID>ID,key,read,w1000,hide,uncfg|�Ŀ�ID,hide,uncfg|�Ŀ��>���,w1000,align<2,0>,uncfg|�걾����,w1600,uncfg|�ؼ�����|�ؼ�ϸĿ|����ID,hide,uncfg|��������,uncfg|��������|��ǰ״̬|�嵥״̬|����ʱ��,fulldatetime,read,w2400|���ʱ��,fulldatetime,read,w2400|"
Public Const gstrSpeExamWorkConvertFormat = "�����:els-<check><source>|�������:0-����,1-����,2-ϸ��,3-����,4-ʬ��,5-����ʯ��|�ؼ�����:0-�����黯,1-����Ⱦɫ,2-���Ӳ���|��������:-1-����,0-����,els-��<source>������|�ؼ�ϸĿ:0-��,1-����,2-��ҩ��ҩ,3-ӫ��,4-��ͨ|��ǰ״̬:0-������,1-�ѽ���,2-�����|�嵥״̬:0-δ��ӡ,1-�Ѵ�ӡ"


Public Const gstrSpeExamWork_ID             As String = "ID"
Public Const gstrSpeExamWork_�������       As String = "�������"
Public Const gstrSpeExamWork_�����         As String = "�����"
Public Const gstrSpeExamWork_����ҽ��ID     As String = "����ҽ��ID"
Public Const gstrSpeExamWork_����           As String = "����"
Public Const gstrSpeExamWork_�Ŀ�ID         As String = "�Ŀ�ID"
Public Const gstrSpeExamWork_�Ŀ��         As String = "�Ŀ��"
Public Const gstrSpeExamWork_�걾����       As String = "�걾����"
Public Const gstrSpeExamWork_�ؼ�����       As String = "�ؼ�����"
Public Const gstrSpeExamWork_����ID         As String = "����ID"
Public Const gstrSpeExamWork_��������       As String = "��������"
Public Const gstrSpeExamWork_��������       As String = "��������"
Public Const gstrSpeExamWork_��ǰ״̬       As String = "��ǰ״̬"
Public Const gstrSpeExamWork_�嵥״̬       As String = "�嵥״̬"
Public Const gstrSpeExamWork_����ʱ��       As String = "����ʱ��"
Public Const gstrSpeExamWork_���ʱ��       As String = "���ʱ��"


'============================================================================================================================


'�ؼ�����ȡ��ʾ��
Public Const gstrSpeExamResultGetCols As String = "|ID,Key,hide,uncfg|�Ŀ��>���,rowcheck,w1000,align<2,0>,uncfg|�걾����,read,uncfg|��������,read,uncfg|��Ŀ���,uncfg|�ؼ�ϸĿ,read|��������,read|��Ŀ˳��,hide,uncfg|"
Public Const gstrSpeExamResultGetConvertFormat As String = "��������:-1-����,0-����,els-��<source>������|�ؼ�ϸĿ:0-��,1-����,2-��ҩ��ҩ,3-ӫ��,4-��ͨ"


Public Const gstrSpeExamResultGet_�Ŀ�� As String = "�Ŀ��"
Public Const gstrSpeExamResultGet_�걾���� As String = "�걾����"
Public Const gstrSpeExamResultGet_�������� As String = "��������"
Public Const gstrSpeExamResultGet_��Ŀ��� As String = "��Ŀ���"
Public Const gstrSpeExamResultGet_��Ŀ˳�� As String = "��Ŀ˳��"



'============================================================================================================================






Public Function GetNumber(ByVal str As String) As Long
'��ȡ�ַ����е�����
    Dim strNum As String
    Dim i As Long
        
    For i = 1 To Len(str)
        If IsNumeric(Mid(str, i, 1)) Then
            strNum = strNum & Mid(str, i, 1)
        End If
    Next i
    
    GetNumber = CLng(IIf(strNum = "", -1, strNum))
    
End Function



Public Function CheckPopedom(ByVal strPrivs As String, ByVal strPopedom As String) As Boolean
'���Ȩ��
    Dim strCurPrivs As String
    
    strCurPrivs = ";" & strPrivs & ";"
    
    CheckPopedom = InStr(1, UCase(strCurPrivs), UCase(";" & strPopedom & ";")) > 0
End Function


Public Sub GetPatholStudyState(ByVal lngAdviceID As Long, ByRef recStudy As TStudyStateInf)
'��ȡ����ŵ����״̬��Ϣ
    Dim strSql As String
    Dim rsPatholNum As ADODB.Recordset
    
    
    strSql = "select ����ҽ��ID,�����,�������,ȡ�Ĺ���,��Ƭ����,���߹���,��Ⱦ����,���ӹ��� from ��������Ϣ where ҽ��id=[1]"
    
    Set rsPatholNum = zlDatabase.OpenSQLRecord(strSql, "��ȡ����״̬��Ϣ", lngAdviceID)
    
    If rsPatholNum.RecordCount <= 0 Then
        recStudy.lngPatholAdviceId = -1
        recStudy.lngStudyType = -1
        recStudy.lngMaterialStep = -1
        recStudy.lngSlicesStep = -1
        recStudy.lngMianYiStep = -1
        recStudy.lngFenZiStep = -1
        recStudy.lngTeRanStep = -1
        recStudy.strPatholNumber = ""
        Exit Sub
    End If
    
    recStudy.lngPatholAdviceId = Val(Nvl(rsPatholNum!����ҽ��id))
    recStudy.lngStudyType = Val(Nvl(rsPatholNum!�������))
    recStudy.lngMaterialStep = Val(Nvl(rsPatholNum!ȡ�Ĺ���))
    recStudy.lngSlicesStep = Val(Nvl(rsPatholNum!��Ƭ����))
    recStudy.lngMianYiStep = Val(Nvl(rsPatholNum!���߹���))
    recStudy.lngFenZiStep = Val(Nvl(rsPatholNum!���ӹ���))
    recStudy.lngTeRanStep = Val(Nvl(rsPatholNum!��Ⱦ����))
    recStudy.strPatholNumber = Nvl(rsPatholNum!�����)
End Sub

Public Function GetPatholMenuIndex(objMenuBar As Object) As Long
'��ȡ����˵�����
    Dim cbrPathol As CommandBarControl
    
    Set cbrPathol = objMenuBar.FindControl(, conMenu_PatholManage)
    
    If Not cbrPathol Is Nothing Then
        GetPatholMenuIndex = cbrPathol.Index
    Else
        GetPatholMenuIndex = 3
    End If
End Function


Public Function HasMenu(objMenuBar As Object, ByVal lngMenuId As Long) As Boolean
'�Ƿ����ָ���˵�
    Dim cbrParentMenu As CommandBarControl
    
    Set cbrParentMenu = objMenuBar.FindControl(, lngMenuId)
    
    HasMenu = IIf(cbrParentMenu Is Nothing, False, True)
End Function


Public Function GetHistoryQuerySql(ByVal strSourceSql As String) As String
'ȡ��ת�������ݲ�ѯ���
    Dim strNewSql As String
    
    strNewSql = strSourceSql
    
    strNewSql = Replace(strNewSql, "����ҽ����¼", "H����ҽ����¼")
    strNewSql = Replace(strNewSql, "����ҽ������", "H����ҽ������")
    strNewSql = Replace(strNewSql, "Ӱ�����¼", "HӰ�����¼")
    
    strNewSql = Replace(strNewSql, "���Ӳ�����¼", "H���Ӳ�����¼")
    strNewSql = Replace(strNewSql, "���Ӳ�������", "H���Ӳ�������")
    
    
'    ����������10.32.0֮��ȡ������ת��
'    strNewSql = Replace(strNewSql, "��������Ϣ", "H��������Ϣ")
'    strNewSql = Replace(strNewSql, "����������Ϣ", "H����������Ϣ")
'    strNewSql = Replace(strNewSql, "����걾��Ϣ", "H����걾��Ϣ")
'    strNewSql = Replace(strNewSql, "�����ͼ���Ϣ", "H�����ͼ���Ϣ")
'    strNewSql = Replace(strNewSql, "����ȡ����Ϣ", "H����ȡ����Ϣ")
'    strNewSql = Replace(strNewSql, "�����Ѹ���Ϣ", "H�����Ѹ���Ϣ")
'    strNewSql = Replace(strNewSql, "������Ƭ��Ϣ", "H������Ƭ��Ϣ")
'    strNewSql = Replace(strNewSql, "������̱���", "H������̱���")
'    strNewSql = Replace(strNewSql, "����������Ϣ", "H����������Ϣ")
'    strNewSql = Replace(strNewSql, "�����ؼ���Ϣ", "H�����ؼ���Ϣ")
'    strNewSql = Replace(strNewSql, "�������ӳ�", "H�������ӳ�")
'    strNewSql = Replace(strNewSql, "���������Ϣ", "H���������Ϣ")
'    strNewSql = Replace(strNewSql, "����鵵��Ϣ", "H����鵵��Ϣ")
  
    
    GetHistoryQuerySql = strNewSql
    
End Function




Public Sub InitDebugObject(ByVal lngModuleNum As Long, ByVal frmMain As Object, ByVal strUser As String, ByVal strPwd As String)
'��ʼ������״̬�µ��������
    Set gcnOracle = New ADODB.Connection
    
    Call OraDataOpen("", strUser, strPwd)
    
    glngSys = 100
    gstrPrivs = ";PACS�����ӡ;PACS����ɾ��;PACS������д;PACS�������Ʊ���;PACS�����޶�;PACS���˱���;�ɼ���������;��������;�洢����;��������;����;��鱨��;���Ǽ�;������;��ɫͨ��;�Ŷӽк�;���ͼ��;ȡ������;ȡ��������;ɾ����ʱӰ��;��Ƶ�ɼ�;���;���п���;ͼ�����;δ�ɷѱ���;�ļ�����;�ޱ������;Ӱ���ʿ�;������������;Excel���;"
    glngModul = lngModuleNum
    
    UserInfo.ID = 281
    UserInfo.���� = "������"
    UserInfo.�û��� = "ZLHIS"
    UserInfo.��� = "1123"
    UserInfo.���� = "WGY"
    UserInfo.����ID = "65"
    
    
    Call InitCommon(gcnOracle)
    
    Call RegCheck
        
    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs) '��ʼ��ҽ�����������Ĳ���
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
End Sub


Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSql As String
    Dim strError As String
    
    On Error Resume Next
    err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If err <> 0 Then
            '���������Ϣ
            strError = err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    err = 0
    On Error GoTo errHand
    
    gstrDBUser = UCase(strUserName)
    SetDbUser gstrDBUser
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    err = 0
End Function
