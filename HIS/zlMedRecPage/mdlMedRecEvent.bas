Attribute VB_Name = "mdlMedRecEvent"
Option Explicit
'----------------------------------------------------------
'����    ����ҳ�ؼ��¼����з�װ���Լ�����ؼ��Ĵ���
'������  ����˶
'�������ڣ�2013/10/30
'���̺�����
'
'
'�޸ļ�¼:
'
'----------------------------------------------------------
'----------------------------------------------------------
'�ؼ�ö��
'˵����1��סԺ��ҳ����׼�桢�Ĵ��桢���ϰ桢���ϰ棩��������ҳ
'       �벡����ҳ����׼�桢�Ĵ��桢���ϰ桢���ϰ棩
'       ����ͬ��Ϣ�ı༭�ؼ���������IndexҪ������ͬ��
'       �˴��ؼ�ö�ٽ�9�������ö��������һ��
'      2������������ҳ�ؼ����٣����������ҳ����ö��
'      3����д��OM=������ҳ��PM=������ҳ��IM=סԺ��ҳ��DS=���ѡ����
'               ST=��׼�棬YN=���ϰ棬SC=�Ĵ��棬HN=���ϰ�
'      4��������ţ�!=ֻ�У�&=��,/=����
'      5���ؼ��������˵����+�������ж��������1��   -������û�У�����1��
'                          �޺�׺��һ����һ��
'      6����ͬ��Ϣ�ı��ؼ�����Index������ͬ������ͬʱͨ��������������ʵ����ͬ
'----------------------------------------------------------
   
Public Enum ErrCol
    ERR_ID = 0
    ERR_���� = 1
    ERR_��Ϣ = 2
End Enum

Public Enum Pane_ID
    Pane_���� = 1
    Pane_��ҳ = 2
    Pane_��� = 3
End Enum

'ҳ��ö��(OM /IM / PM)
Public Enum PIC�˵�
    PIC_סԺ��ҳ = 0
    PIC_������Ϣ = 1
    PIC_��ҽ��� = 2
    PIC_��ҽ������ = 3
    PIC_��ҽ��� = 4
    PIC_��ҽ������ = 5
    PIC_ҩ����� = 6
    PIC_��Ѫ��Ϣ = 7
    PIC_ǩ����Ϣ = 8
    PIC_������¼ = 9
    PIC_סԺ���� = 10
    PIC_סԺ��� = 11
    PIC_������Ϣ = 12
    PIC_���Ƽ�¼ = 13
    PIC_����ҩ�� = 14
    PIC_������ = 15
    PIC_��֢�໤ = 16
    PIC_�������� = 17
    PIC_��ҳ1 = 18
    PIC_��ҳ2 = 19
End Enum

'������ؿؼ�(OM /IM / PM)
Public Enum ParaCtrl
    '��ҽ���(IM / PM): optDiag
    PC_XY��������� = 0
    PC_XY�������������� = 1
    '��ҽ���(IM /PM): optDiag
    PC_ZY��������� = 2
    PC_ZY�������������� = 3
    '������Ϣ(OM):optDiag
    PC_��������� = 0
    PC_�������������� = 1
    '����������(IM)/������Ϣ(OM):optAller
    PC_��ҩƷĿ¼���� = 0
    PC_������Դ���� = 1
    '����������(IM / PM):OptParaOPSInfo��chkParaOPSInfo
    PC_��������Ŀ���� = 0
    PC_��ICDCM9�������� = 1
    PC_δ�ҵ�ʱ����¼�� = 0 'chkParaOPSInfo
End Enum

'��Աչʾ��ؿؼ�(IM / PM):lblManInfo,cboManInfo
Public Enum ManCtrl
    'סԺ���(IM)/ҽ��������(PM)
    MC_����ҽʦ = 0
    MC_������ = 1
    MC_���λ����� = 2
    MC_����ҽʦ = 3
    MC_����ҽʦ = 4
    MC_סԺҽʦ = 5
    MC_�о���ҽʦ = 6
    MC_ʵϰҽʦ = 7
    MC_�ʿ�ҽʦ = 8
    MC_�ʿػ�ʿ = 9
    MC_���λ�ʿ = 10
    'סԺ���(PM)
    MC_��ĿԱ = 11
    'סԺ���(IM)/ҽ��������(PM) SC
    MC_����ҽʦ = 12
End Enum

'ʱ��������ؿؼ�(OM/IM/PM)��lblDateInfo��mskDateInfo(+),cmdDateInfo(-)
Public Enum DateCtrl
    '������Ϣ(OM/PM/IM)
    DC_�������� = 0
    '������Ϣ(IM/PM)
    DC_��Ժʱ�� = 2
    DC_��Ժʱ�� = 3
    '��ҽ���(IM/PM)
    DC_ȷ������ = 4
    DC_����ʱ�� = 5
    'סԺ���(IM/PM)/�������(OM)
    DC_�������� = 6
    DC_����ʱ�� = 7
    'סԺ���(IM)/ҽ��������(PM)
    DC_�ʿ����� = 8
    'סԺ���(PM)
    DC_��Ŀ���� = 9
    DC_�ջ����� = 10
End Enum

'��ַ��ؿؼ�(OM/IM/PM):lblAdressInfo,txtAdressInfo,padrInfo(-),cmdAdressInfo(-)
Public Enum AdressCtrl
    '������Ϣ(OM/IM/PM)
    ADRC_�����ص� = 0
    ADRC_���� = 1
    ADRC_��סַ = 2
    ADRC_���ڵ�ַ = 3
    '������Ϣ(IM/PM)
    ADRC_��ϵ�˵�ַ = 4
    '������Ϣ(OM/IM/PM)
    ADRC_�������� = 5
    ADRC_��λ��ַ = 6
End Enum

'�����ֵ��̶����� �����б��Ϳؼ�(OM/IM/PM)��lblBaseInfo,cboBaseInfo
Public Enum BaseCodeCtrl
    '������Ϣ(OM/IM/PM)
    BCC_���ʽ = 0
    BCC_�Ա� = 1
    BCC_���� = 2
    BCC_ְҵ = 3
    BCC_���� = 4
    BCC_���� = 5
    '������Ϣ(IM/PM)
    BCC_��ϵ = 6
    BCC_��Ժ;�� = 7
    '������Ϣ(OM)
    BCC_�Ļ��̶� = 8
    '������Ϣ(OM)
    BCC_ȥ�� = 9
    '��ҽ���(IM/PM)
    BCC_��Ⱦ��������ϵ = 10
    BCC_��Ժ��� = 11
    BCC_�ֻ��̶� = 12
    BCC_���������� = 13
    BCC_�������ԺXY = 14
    BCC_��Ժ���ԺXY = 15
    BCC_��������Ժ = 16
    BCC_��ǰ������ = 17
    BCC_�����벡�� = 18
    BCC_�ٴ��벡�� = 19
    BCC_�����ڼ� = 20 '!PM
    BCC_�ٴ���ʬ�� = 21
    '��ҽ���(IM/PM)
    BCC_�������ԺZY = 22
    BCC_��Ժ���ԺZY = 23
    BCC_��֤ = 24
    BCC_�η� = 25
    BCC_��ҩ = 26
    BCC_������� = 27
    BCC_��ҽ�����豸 = 28
    BCC_���ȷ��� = 29
    BCC_��ҽ���Ƽ��� = 30
    BCC_������ҩ�Ƽ� = 31
    BCC_��֤ʩ�� = 32
    'ҽ��������(PM)/סԺ���(IM)
    BCC_�������� = 33
    'סԺ���(IM/PM)
    BCC_�������� = 34
    BCC_HBsAg = 35
    BCC_Ѫ�� = 36 '������Ϣ(OM)
    BCC_HCVAb = 37
    BCC_RH = 38 '������Ϣ(OM)
    BCC_HIVAb = 39
    BCC_��Һ��Ӧ = 40
    BCC_��Ѫ��Ӧ = 41
    BCC_��Ѫǰ9���� = 42
    BCC_����״�� = 43 '������Ϣ(OM)
    BCC_��Ժ��ʽ = 44
    BCC_����Ժ�ƻ����� = 49
    '����(IM/PM)
    BCC_ѹ�������ڼ� = 45
    BCC_ѹ������ = 46
    BCC_������׹���˺� = 47
    BCC_������׹��ԭ�� = 48
    'YN:��ҳ1��PM)
    BCC_���ϴ�סԺʱ�� = 50
    'YN:��ҳ2��IM/PM)
    BCC_�ط����ʱ�� = 51
    BCC_Լ����ʽ = 52
    BCC_Լ������ = 53
    BCC_Լ��ԭ�� = 54
    BCC_��������Ժ��ʽ = 55
    'HN:������IM / PM),SC:סԺ�����IM / PM)
    BCC_�������� = 56
    'HN:������IM / PM)
    BCC_�ٴ�·������ = 57
    BCC_������Ⱦ�� = 58
    BCC_ʵʩDGRS���� = 59
    BCC_��������ʬ�� = 60
    BCC_���֤ = 61
    BCC_����ԭ�� = 62
    BCC_�������� = 63
End Enum

'CheckBox�ؼ�(IM/PM/OM):chkInfo
Public Enum CheckCtrl
    '������Ϣ(IM/PM)
    CHK_����Ժ = 0
    CHK_��Ժǰ��Ժ���� = 1
    '��ҽ���(IM/PM)
    CHK_�Ƿ�ȷ�� = 2
    CHK_��ԭѧ��� = 3
    'CHK_��������ʬ�� = 4
    CHK_�·����� = 5
    '��ҽ���(IM/PM)
    CHK_Σ�� = 6
    CHK_��֢ = 7
    CHK_���� = 8
    'סԺ���(IM/PM)
    CHK_ʾ�̲��� = 9
    CHK_���в��� = 10
    CHK_���Ѳ��� = 11
    CHK_���� = 12
    '����(IM/PM)
    CHK_CT = 13
    CHK_MRI = 14
    CHK_��ɫ������ = 15
    '������Ϣ(OM)
    CHK_��Ⱦ���ϴ� = 16
    '����������(YN:IM/PM)
    CHK_Χ�������� = 17
    CHK_������� = 18
    'YN:��ҳ1��IM/PM)
    CHK_����·�� = 19
    CHK_���·�� = 20
    CHK_���� = 21
    CHK_סԺ����Σ�� = 22
    'YN:��ҳ1(PM)��SC:��ҳ2 (IM / PM)
    CHK_�Ƿ�ͬһ���� = 23
    'YN:    ��ҳ2 (IM / PM)
    CHK_�˹������ѳ� = 24
    CHK_�ط���֢ҽѧ�� = 25
    CHK_סԺ����Լ�� = 26
    'HN:������IM / PM)
    CHK_�����ֹ��� = 27
    CHK_ϸ���걾�ͼ� = 28
    'SC_סԺ�����IM / PM)
    CHK_������� = 29
    '������Ϣ��OM)
    CHK_�޹�����¼ = 30
End Enum

'�޶�����ؼ�(IM/PM/OM):lblSpecificInfo(+),txtSpecificInfo(+,-)��cmdSpecificInfo(-),cboSpecificInfo(-)
Public Enum SpecificLimitCtrl
    '�������(IM/PM/OM)
    SLC_��λ�绰 = 1
    SLC_��λ�ʱ� = 2
    SLC_��ͥ�绰 = 3
    SLC_��ͥ�ʱ� = 4
    SLC_�����ʱ� = 5
    SLC_��� = 6
    SLC_��ߵ�λ = 7
    SLC_���� = 8
    SLC_���ص�λ = 9
    '�������(OM)
    SLC_���� = 10
    '�������(IM/PM)
    SLC_��Ժ���� = 11
    '�������(OM)
    SLC_����ѹ = 12
    SLC_����ѹ = 13
    '�������(IM/PM)
    SLC_��ϵ�˵绰 = 14
    SLC_���� = 15
    SLC_Ӥ�׶����� = 16
    SLC_�������������� = 17
    SLC_��������Ժ���� = 18
    SLC_סԺ���� = 19
    SLC_סԺ�� = 20
    '��ҽ���(IM/PM)
    SLC_���ȴ��� = 21
    SLC_�ɹ����� = 22
    'ҽ��������(PM)
    SLC_�ػ� = 23
    SLC_һ������ = 24
    SLC_�������� = 25
    SLC_�������� = 26
    SLC_ICU = 27
    SLC_CCU = 28
    'סԺ���(IM/PM)
    SLC_���ϸ�� = 29
    SLC_��ѪС�� = 30
    SLC_��Ѫ�� = 31
    SLC_��ȫѪ = 32
    SLC_������� = 33
    SLC_������ʹ�� = 34
    SLC_����ʱ����Ժǰ_�� = 35
    SLC_����ʱ����Ժǰ_Сʱ = 36
    SLC_����ʱ����Ժǰ_���� = 37
    SLC_����ʱ����Ժ��_�� = 38
    SLC_����ʱ����Ժ��_Сʱ = 39
    SLC_����ʱ����Ժ��_���� = 40
    SLC_�������� = 41
    'סԺ���(PM)
    SLC_���ú� = 42
    'YN:��ҳ2��IM/PM��
    SLC_Լ����ʱ�� = 43
    'HN:������IM/PM��
    SLC_��֢�໤�� = 44
    SLC_��֢�໤Сʱ = 45
    SLC_Apgar = 46
    'SC:������Ϣ��IM/PM��
    SLC_QQ = 47
    'SC:סԺ�����IM/PM��
    SLC_��׵��� = 48
    SLC_Ժ�ڻ��� = 49
    SLC_��Ժ���� = 50
    'SC:����
    SLC_���ϴ�סԺʱ�� = 51 'YN:��ҳ1��PM)BCC_���ϴ�סԺʱ�� = 50
    SLC_Ӥ�׶�����_DAY = 52 'Ӥ�����䵥λΪ��ʱ������������ʽ ���ݴ洢��ʽ:2��15�� ��ĸ�̶�Ϊ30
End Enum

'��ͨ�ؼ���Ϣ(IM/PM/OM)��lblInfo(+),txtInfo(+)
Public Enum GeneralCtrl
    '������Ϣ(PM)
    GC_������ = 0
    GC_������ = 1
    GC_X�ߺ� = 2
    '������Ϣ(IM/PM/OM)
    GC_���� = 3
    GC_����֤�� = 4
    'GC_��λ��ַ = 5  ����Ϊ ADRC_��λ��ַ = 6
    '������Ϣ(IM/PM)
    GC_��ϵ������ = 6
    GC_��Ժ���� = 7
    GC_��Ժ���� = 8
    GC_��Ժ���� = 9
    GC_��Ժ���� = 10
    '������Ϣ(PM)
    GC_ҽ���� = 11
    '������Ϣ(OM)
    GC_ժҪ = 12
    '������Ϣ(OM)
    GC_����� = 14
    GC_�໤�� = 15
    '������Ϣ(OM)
    GC_������ַ = 16
    'סԺ���(IM/PM)/������Ϣ(OM)
    GC_ҽѧ��ʾ = 17
    GC_����ҽѧ��ʾ = 18
    '��ҽ���(IM/PM)
    GC_����� = 19
    GC_����ԭ�� = 20
    GC_��ԭѧ��� = 21
    GC_���Ȳ��� = 22
    'סԺ���(IM/PM)
    GC_������ = 23
    GC_ת��ҽ�ƻ��� = 24
    GC_31������סԺ = 25
    '������Ϣ(IM/PM)
    GC_ת��1 = 27
    GC_ת��2 = 28
    GC_ת��3 = 29
    'YN:��ҳ1��IM/PM��
    GC_�˳�ԭ�� = 30
    GC_����ԭ�� = 31
    'YN:��ҳ2��IM/PM��
    GC_��֢�໤������ = 32
    'HN:������IM/PM��
    GC_����T = 33
    GC_����N = 34
    GC_����M = 35
    'SC:������Ϣ��IM/PM��
    GC_Email = 36
    'SC:סԺ�����IM/PM��
    GC_�������� = 37
    'SC:��ҳ2��IM/PM��
    GC_����ҩ�� = 38
    GC_�ٴ����� = 39
    GC_͸�����ص�ֵ = 40
    '������Ϣ��IM/PM��
    GC_������ϵ = 41
    GC_��Ժת�� = 42
    GC_�໤�����֤�� = 64
End Enum
'OptionButton
Public Enum OPCtrl
    'סԺ���(IM/PM)  optInput
    OP_��סԺ�� = 0
    OP_��סԺ�� = 1
    'OM optState
    OP_���� = 0
    OP_���� = 1
    'HN:����(IM/PM)  optInput
    OP_ICU�� = 2
    OP_ICU�� = 3
End Enum
Public Enum DeptRow
    DR_ת�ƿ��� = 0
    DR_ת��ʱ�� = 1
End Enum

'�Զ���ȡ�ؼ�����cmdAutoLoad
Public Enum AuoLoadCtrl
    ALC_������ = 0
    ALC_���� = 1
    ALC_������¼ = 2
    ALC_�ٴ�·�� = 3
End Enum

'���ؼ���ö��(IM/PM/OM)��������Ϣ��vsAller
Public Enum AllerColsIndex
    AI_����ҩ�� = 0
    AI_������Ӧ = 1
    AI_����ʱ�� = 2
    AI_����Դ���� = 3
    AI_ҩ��ID = 4
    AI_������Դ = 5
End Enum

'���ؼ���ö��(IM/PM/OM/DS)����ҽ��ϣ�vsDiagXY����ҽ��ϣ�vsDiagZY
Public Enum DiagColsIndex
    DI_������� = 0
    DI_���� = 1
    DI_��ϱ��� = 2
    DI_������� = 3
    DI_��ҽ֤�� = 4
    DI_����ʱ�� = 5
    DI_��ע = 6
    DI_��Ժ���� = 7
    DI_��Ժ��� = 8
    DI_ICD���� = 9
    DI_�Ƿ�δ�� = 10
    DI_�Ƿ����� = 11
    DI_���� = 12
    DI_Del = 13
    DI_���ID = 14
    DI_����ID = 15
    DI_֤��ID = 16
    DI_ҽ��IDs = 17 '�뵱ǰ��Ϲ�����ҽ��ID��ɵ��ַ�����ҽ��ID���Զ��ŷָ�
    DI_��Ϸ��� = 18
    DI_�̶����� = 19
    DI_�Ƿ��� = 20
    DI_��Ч���� = 21
    DI_������Ϣ = 22
    DI_����ID = 23
    DI_�����Դ = 24
    DI_�������� = 25
    DI_������� = 26
    DI_֤����� = 27
    DI_��¼���� = 28
    DI_��¼��Ա = 29
End Enum

'������ö��
Public Enum TSJCRow
    'HN��YN,ST
    TR_������4 = 0
    TR_������5 = 1
    TR_������6 = 2
    'SC
    TR_CT = 0
    TR_PETCT = 1
    TR_˫ԴCT = 2
    TR_XƬ = 3
    TR_B�� = 4
    TR_�����Ķ�ͼ = 5
    TR_MRI = 6
    TR_ͬλ�ؼ�� = 7
End Enum
'������ö��
Public Enum OPSColsIndex
    PI_Copy = 0
    PI_�������� = 1
    PI_�������� = 2
    PI_������ҩʱ�� = 3
    PI_������� = 4
    PI_׼������ = 5
    PI_�������� = 6
    PI_�������� = 7
    PI_�ٴ����� = 8
    PI_����ҽʦ = 9
    PI_������ʿ = 10
    PI_����1 = 11
    PI_����2 = 12
    PI_����ʼʱ�� = 13
    PI_�������� = 14 '������������ʽ
    PI_ASA�ּ� = 15
    PI_NNIS�ּ� = 16
    PI_�������� = 17
    PI_����ҽʦ = 18
    PI_�п����� = 19
    PI_�пڲ�λ = 20
    PI_�ط������Ҽƻ� = 21
    PI_�ط�������Ŀ�� = 22
    PI_�пڸ�Ⱦ = 23
    PI_����֢ = 24
    PI_Ԥ���ÿ���ҩ = 25
    PI_����ҩ���� = 26
    PI_��Ԥ�ڵĶ������� = 27
    PI_������֢ = 28
    PI_������������ = 29
    PI_��������֢ = 30
    PI_�����Ѫ��Ѫ�� = 31
    PI_�����˿��ѿ� = 32
    PI_�������Ѫ˨ = 33
    PI_���������л���� = 34
    PI_�������˥�� = 35
    PI_�����˨�� = 36
    PI_�����Ѫ֢ = 37
    PI_�����Źؽڹ��� = 38
    PI_��������ID = 39
    PI_������ĿID = 40
    PI_����ID = 41
    PI_����ʽ = 42 '����������������
    PI_������Դ = 43
End Enum
'������ö��
Public Enum ChemothColsIndex
    CI_��ѧ���Ʊ��� = 0
    CI_��ʼ���� = 1
    CI_�������� = 2
    CI_�Ƴ��� = 3
    CI_���Ʒ��� = 4
    CI_���� = 5
    CI_����Ч�� = 6
    CI_����ID = 7
End Enum
'������ö��
Public Enum RadiothColsIndex
    RI_�������Ʊ��� = 0
    RI_��ʼ���� = 1
    RI_�������� = 2
    RI_��Ұ��λ = 3
    RI_������� = 4
    RI_�ۼ��� = 5
    RI_����Ч�� = 6
    RI_����ID = 7
End Enum

'����ҩƷ��ö��
Public Enum SpiritColsIndex
    SI_ҩ������ = 0
    SI_�Ƴ� = 1
    SI_������� = 2
    SI_���ⷴӦ = 3
    SI_��Ч = 4
    SI_ҩƷID = 5
End Enum
'��֢�໤ö��
Public Enum ICUColsIndex
    UI_��� = 0
    UI_�໤������ = 1
    UI_����ʱ�� = 2
    UI_�˳�ʱ�� = 3
    UI_����ס�ƻ� = 4
    UI_����סԭ�� = 5
End Enum
'��֢�໤��еö��
Public Enum ICUInstruColsIndex
    TI_ICU���� = 0
    TI_��е������ = 1
    TI_��ʼʱ�� = 2
    TI_����ʱ�� = 3
    TI_��Ⱦ�ۼ�Сʱ = 4
End Enum
'ҽԺ��Ⱦö��
Public Enum InfectColsIndex
    FI_ȷ������ = 0
    FI_��Ⱦ��λ = 1
    FI_ҽԺ��Ⱦ���� = 2
    FI_ҽԺ��Ⱦ���� = 3
End Enum
'�걾��Դö��
Public Enum SampleColsIndex
    MI_�걾 = 0
    MI_��ԭѧ���뼰���� = 1
    MI_�ͼ����� = 2
End Enum
'����ҩö��
Public Enum KSSColsIndex
    KI_��� = 0
    KI_����ҩ���� = 1
    KI_��ҩĿ�� = 2
    KI_ʹ�ý׶� = 3
    KI_ʹ������ = 4
    KI_һ���п�Ԥ���� = 5
    KI_DDD�� = 6
    KI_������ҩ = 7
End Enum

Private mblnChk  As Boolean  '�Ƿ�ִ��chk����¼�
Private mobjDiag As Object   '��ϱ�����
Private mblnReturn As Boolean

'--------------------------------------------------------------------------
'�ؼ��¼���װ
'���������������ؼ���+�¼���
'--------------------------------------------------------------------------
'Form�¼�
Public Sub FormActivate()
'Form_Activate�¼�
'    If gclsPros.IsLoad Then
'        Call ChangePage(, 0)
'    End If
'    gclsPros.IsLoad = False
End Sub

Public Sub FormKeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'Form_KeyDown�¼�
    Dim lngIndex As Long, i As Long, lngCount As Long
    Dim monTmp As MonthView

    With gclsPros.CurrentForm
        Select Case intKeyCode
            '���·�
            Case vbKeyPageDown
                Call ChangePage(True)
                Exit Sub
            '���Ϸ�
            Case vbKeyPageUp
                Call ChangePage(False)
                Exit Sub
            '������ǰ
            Case vbKeyHome
                .vsbMain.Value = .vsbMain.Min
                Exit Sub
            '�������
            Case vbKeyEnd
                .vsbMain.Value = .vsbMain.Max
                Exit Sub
            Case vbKeyUp
                If intShift = 2 Then
                    i = .vsbMain.Value
                    If i - 10 < .vsbMain.Min Then
                        i = .vsbMain.Min
                    Else
                        i = i - 10
                    End If
                    .vsbMain.Value = i
                    Exit Sub
                End If
            Case vbKeyDown
                If intShift = 2 Then
                    i = .vsbMain.Value
                    If i + 10 > .vsbMain.Max Then
                        i = .vsbMain.Max
                    Else
                        i = i + 10
                    End If
                    .vsbMain.Value = i
                    Exit Sub
                End If
            Case vbKeyEscape
                '�������Σ���Ҫ������ĳЩ����û��monInfo�ؼ�����������ҳ
                On Error Resume Next
                Set monTmp = .monInfo
                If Err.Number = 0 Then
                    monTmp.Visible = False
                    Err.Clear: On Error GoTo 0
                    Call ShowInfectInfo(False)
                End If
            Case vbKeyF5
                Call ChangePage(True, 1)
            Case vbKeyF6
                Call ChangePage(True, 2)
            Case vbKeyF7
                Call ChangePage(True, 6)
            Case vbKeyF8
                Call ChangePage(True, 8)
            Case vbKeyF9
                Call ChangePage(True, 9)
            Case vbKeyF10
                Call ChangePage(True, 12)
            Case vbKeyF11
                Call ChangePage(True, 14)
            Case vbKeyF12
                Call ChangePage(True, 17)
        End Select
        If gclsPros.FuncType = f������ҳ Then
            If gclsPros.OpenMode = EM_���� Or gclsPros.OpenMode = EM_�༭ Then
                If intShift = 2 And intKeyCode = vbKeyU Then
                    CmdUPClick
                ElseIf intShift = 2 And intKeyCode = vbKeyD Then
                    CmdDownClick
                End If
            End If
        End If
        If intKeyCode = vbKeyS And intShift = 2 Then
            If gclsPros.FuncType = fҽ����ҳ Then
                If gclsPros.InfosChange Then
                    Call menuPageOperate(MOP_ȷ��)
                End If
             ElseIf gclsPros.FuncType = f������ҳ And gclsPros.OpenMode <> EM_���� Then
                If gclsPros.InfosChange Then
                    Call menuPageOperate(MOP_ȷ��)
                End If
            End If
        End If
    End With
End Sub

Public Sub FormKeyPress(ByRef intKeyAscii As Integer)
'Form_KeyPress�¼�
    If intKeyAscii = Asc("'") Then intKeyAscii = 0
End Sub

Public Function FormLoad(Optional ByVal blnChange As Boolean) As Boolean
    Dim i As Integer
'���ܣ���ҳ����Form_Load�¼�
'������blnChange=�Ƿ��ǻ�ȡ��һ�ݻ�����һ�ݲ�������
'���أ�Ture-�ɹ���False-ʧ��
    gclsPros.IsOpen = True
    gclsPros.IsLoad = True
    On Error GoTo errH
    With gclsPros.CurrentForm
        '�����ʼ���Լ����ݼ���
        On Error GoTo errH:
        '����������Ҫ������ҳ�����߻�ȡ��һ��סԺ����һ��סԺ����ҳID
        If gclsPros.FuncType = f������ҳ Then
            If Not ValiAndGet��ҳID Then Exit Function
            Call OpenExtraData
            Select Case gclsPros.OpenMode
                Case EM_�༭
                    gclsPros.NoType = IT_Old
                    gclsPros.IsExistPati = True
                Case EM_��������
                    gclsPros.NoType = IT_New
                Case EM_������ҳ
                    gclsPros.NoType = IT_Old
                    gclsPros.IsExistPati = True
            End Select
        End If
        
        'ҳ���л�ʱ,���ظ�����
        If Not blnChange Then
            If gclsPros.FuncType = fҽ����ҳ Or gclsPros.FuncType = f������ҳ Then
                '����Ƿ�����Ҳ���������ҳ��ҳ
                Call CreatePlugInOK(gclsPros.Module)
                If Not gobjPlugIn Is Nothing Then
                    Err.Clear: On Error Resume Next
                    If gobjPlugIn.gblnfrmMec = True Then
                        '���ò��������Զ��帽ҳ�ӿ�
                        If Err.Number = 0 Then
                            Set gfrmMecCol = gobjPlugIn.GetMeRecFormCol(gclsPros.SysNo, gclsPros.Module, gclsPros.����ID, gclsPros.��ҳID, gclsPros.PatiType)
                        End If
                        If Err.Number = 0 Then
                            gBlnNew = True
                        Else
                            gBlnNew = False
                        End If
                        Call zlPlugInErrH(Err, "GetMeRecFormCol")
                        Err.Clear: On Error GoTo 0
                    End If
                End If
            End If
            
            If gBlnNew = True And (Not gfrmMecCol Is Nothing) Then
                gIntPic = gclsPros.CurrentForm.PicPage.Count - 1
                For i = 1 To gfrmMecCol.Count
                    gPic��Ҹ�ҳ = gclsPros.CurrentForm.PicPage.Count
                    Load gclsPros.CurrentForm.PicPage(gPic��Ҹ�ҳ)
                    gclsPros.CurrentForm.PicPage(gPic��Ҹ�ҳ).Height = gfrmMecCol(i).Height + 50
                    SetParent gfrmMecCol(i).hwnd, gclsPros.CurrentForm.PicPage(gPic��Ҹ�ҳ).hwnd
                    gfrmMecCol(i).Top = IIf(gclsPros.FuncType = f������ҳ, IIf(gclsPros.MedPageSandard = ST_��������׼, -300, -1900), -300): gfrmMecCol(i).Left = 0: gfrmMecCol(i).Tag = gPic��Ҹ�ҳ
                    gfrmMecCol(i).Show
                    Set gclsPros.CurrentForm.PicPage(gPic��Ҹ�ҳ).Container = gclsPros.CurrentForm.picMain
                Next
            End If
        End If
        Call SetAllObject
        If Not InitMedRecEnv Then Exit Function
        If gclsPros.OpenMode <> EM_�������� Then   '�����������ü�������
            If Not LoadMedPageData(gclsPros.����ID, gclsPros.��ҳID, gclsPros.RegistNo, gclsPros.PatiType = PF_����) Then Exit Function
        End If
        Call SetPageVisible '����Ĭ��ҳ���Լ�ҳ��ɼ���
        Call gclsPros.InitFacePara '���ý�������ؼ�״̬
        If Not InitMedRecEnv(True) Then Exit Function '���ݼ��غ�Ľ������
        If gclsPros.PatiType = PF_סԺ And gclsPros.FuncType = fҽ����ҳ Then
            gclsPros.IsSigned = SetSignature
        End If
        Call SetFaceInit                        '������ظ�����ʼ״̬
        Call SetFaceEditable(gclsPros.IsSigned) '���ý���ؼ�������
        .subcMain.hwnd = .hwnd
        .subcMain.Messages(WM_MOUSEWHEEL) = True
        Call SetAllVSF
        Call SetPicPosition(True, True)
        
        Call SetComboBoxProperty(True) '���ε�ComboBox���������¼�
    End With
    FormLoad = True
    gclsPros.LoadFinish = True
    'Ĭ��Ӥ�׶�����
    Call CboSpecificInfoClick(SLC_Ӥ�׶�����)
    gclsPros.InfosChange = False
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub FormUnLoad(ByRef blnCancel As Integer)
    Dim i As Integer
    If (gclsPros.InfosChange And Not gclsPros.IsOK And gclsPros.FuncType = f���ѡ��) Or (gclsPros.FuncType <> f���ѡ�� And gclsPros.InfosChange And gclsPros.OpenMode <> EM_����) Then
        If MsgBox("����˳����ղ����޸ĵ����ݽ����ᱻ���档ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            blnCancel = True: Exit Sub
        End If
    End If
    If gclsPros.FuncType = f���ѡ�� Or gclsPros.FuncType = fҽ����ҳ Then
        If gclsPros.PatiType = PF_���� Then
            Call zlDatabase.SetPara("�����������", gclsPros.DiagInputXY, gclsPros.SysNo, p����ҽ��վ, InStr(gclsPros.Privs, "��������") > 0)
            If gclsPros.FuncType = fҽ����ҳ Then
                Call zlDatabase.SetPara("����������Դ", gclsPros.AllerInput, gclsPros.SysNo, p����ҽ��վ, gclsPros.AllerSource = 0 And gclsPros.PassType = 3 And InStr(gclsPros.Privs, "��������") > 0)
            End If
        Else
            Call zlDatabase.SetPara("��ҽ�������", gclsPros.DiagInputXY, gclsPros.SysNo, pסԺҽ��վ, InStr(gclsPros.Privs, "��������") > 0)
            Call zlDatabase.SetPara("��ҽ�������", gclsPros.DiagInputZY, gclsPros.SysNo, pסԺҽ��վ, InStr(gclsPros.Privs, "��������") > 0)
            If gclsPros.FuncType = fҽ����ҳ Then
                Call zlDatabase.SetPara("����������Դ", gclsPros.AllerInput, gclsPros.SysNo, p����ҽ��վ, gclsPros.AllerSource = 0 And gclsPros.PassType = 3 And InStr(gclsPros.Privs, "��������") > 0)
                Call zlDatabase.SetPara("�����������", gclsPros.OPSInput & IIf(gclsPros.OPSFree, 1, 0), gclsPros.SysNo, pסԺҽ��վ, InStr(gclsPros.Privs, "��������") > 0)
            End If
        End If
    End If
    Call SaveWinState(gclsPros.CurrentForm, App.ProductName)
    gclsPros.IsOpen = False
    If gclsPros.FuncType <> f������ҳ Then
        Call gclsMain.Closed(Not gclsPros.IsOK, gclsPros.DiseaseIDs, gclsPros.DiagIDs, gclsPros.PictureFile)
    End If
    If gclsPros.FuncType = f������ҳ Or gclsPros.FuncType = fҽ����ҳ Then
        With gclsPros.CurrentForm
            .subcMain.Messages(WM_MOUSEWHEEL) = False
        End With
        Call SetComboBoxProperty(False)
    End If
    'ж����Ҹ�ҳ
    If gBlnNew = True And (Not gfrmMecCol Is Nothing) Then
        For i = 1 To gfrmMecCol.Count
            Unload gfrmMecCol(i)
        Next
        gBlnNew = False
        Set gfrmMecCol = Nothing
    End If
End Sub

'cmdCancel�¼�
Public Sub CmdCancelClick()
    Unload gclsPros.CurrentForm
End Sub

Public Sub CmdCancelGotFocus()
'cmdCancel_GotFocus�¼�
    Call ShowInfectInfo(False)
End Sub

'CmdDown�¼�
Public Sub CmdDownClick()
'CmdDown_Click�¼�
    If gclsPros.OpenMode <> EM_���� And gclsPros.InfosChange Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ�ϲ鿴��һ�ݲ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
           Exit Sub
        End If
    End If
    Call ClearPageContent
    Call gclsPros.RefreshPara
    Call gclsPros.InitCacheRecInfo
    gclsPros.��ҳID = Get��ҳIDByCur(gclsPros.��ҳID, True)
    gclsPros.Is��Ŀ = True
    Call FormLoad(True)
End Sub

Public Sub CmdDownGotFocus()
'CmdDown_GotFocus�¼�
    Call ShowInfectInfo(False)
End Sub

'cmdHelp�¼�
Public Sub CmdHelpClick()
'cmdHelp_Click�¼�
    ShowHelp App.ProductName, gclsPros.CurrentForm.hwnd, gclsPros.CurrentForm.Name, gclsPros.SysNo \ 100
End Sub

'cmdDiagMove�¼�
Public Sub CmdDiagMoveClick(ByRef intIndex As Integer)
'cmdDiagMove_Click�¼�
    Call MoveDiagRows(IIf(intIndex \ 2 = 0, gclsPros.CurrentForm.vsDiagXY, gclsPros.CurrentForm.vsDiagZY), IIf(intIndex Mod 2 = 0, -1, 1))
End Sub

Public Sub CmdDiagMoveGotFocus(ByRef intIndex As Integer)
'cmdDiagMove_GotFocus�¼�
    '��ҽ��Ҫ����
    If intIndex \ 2 = 0 Then Call ShowInfectInfo(False)
End Sub

Public Sub CmdHelpGotFocus()
'cmdHelp_GotFocus�¼�
    Call ShowInfectInfo(False)
End Sub

'cmdOPSMove�¼�
Public Sub cmdOPSMoveClick(ByRef intIndex As Integer)
'cmdOPSMove_Click�¼�
    Call MoveOPSRows(gclsPros.CurrentForm.vsOPS, IIf(intIndex Mod 2 = 0, -1, 1))
End Sub

'CmdUp�¼�
Public Sub CmdUPClick()
'CmdUP_Click�¼�
    If gclsPros.OpenMode <> EM_���� And gclsPros.InfosChange Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ�ϲ鿴��һ�ݲ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
           Exit Sub
        End If
    End If
    Call ClearPageContent
    Call gclsPros.RefreshPara
        Call gclsPros.InitCacheRecInfo
    gclsPros.��ҳID = Get��ҳIDByCur(gclsPros.��ҳID, False)
    gclsPros.Is��Ŀ = True
    Call FormLoad(True)
End Sub

Public Sub CmdUPGotFocus()
'CmdUP_GotFocus�¼�
    Call ShowInfectInfo(False)
End Sub

Public Sub cmdDoctorDiagClick(ByVal Index As Integer, ByVal frmParent As Form)
    Dim vPoint As POINTAPI
    If Index = 1 Then
        vPoint = GetCoordPos(gclsPros.CurrentForm.vsDiagXY.hwnd, gclsPros.CurrentForm.vsDiagXY.Left + 15, gclsPros.CurrentForm.vsDiagXY.CellTop)
        frmPublicTable.ShowMe Index, gclsPros.����ID, gclsPros.��ҳID, frmParent, vPoint.X, vPoint.Y, gclsPros.CurrentForm.vsDiagXY.Height
    Else
        vPoint = GetCoordPos(gclsPros.CurrentForm.vsDiagZY.hwnd, gclsPros.CurrentForm.vsDiagZY.Left + 15, gclsPros.CurrentForm.vsDiagZY.CellTop)
        frmPublicTable.ShowMe Index, gclsPros.����ID, gclsPros.��ҳID, frmParent, vPoint.X, vPoint.Y, gclsPros.CurrentForm.vsDiagZY.Height
    End If
End Sub

Public Sub cmdDoctorOPSClick(ByVal frmParent As Form)
    Dim vPoint As POINTAPI
    vPoint = GetCoordPos(gclsPros.CurrentForm.vsOPS.hwnd, gclsPros.CurrentForm.vsOPS.Left + 15, gclsPros.CurrentForm.vsOPS.CellTop)
    frmPublicTable.ShowMe 3, gclsPros.����ID, gclsPros.��ҳID, frmParent, vPoint.X, vPoint.Y, gclsPros.CurrentForm.vsOPS.Height
End Sub

'cmdDeliceryInfo�¼�
Public Sub CmdDeliceryInfoClick(Optional ByVal bytFunc As Byte = 0, Optional ByRef objFrmMain As Object, Optional ByVal lngPatiID As Long, Optional ByVal lngMainID As Long)
'cmdDeliceryInfo_Click�¼�
'����:
'   bytFunc=0 ����ϵͳ����
'   bytFunc=1 �������Ǽǵ���
    Dim LngRow As Long
    Dim str��� As String, dat��Ժ���� As Date, dat��Ժ���� As Date
    Dim strTmp As String, blnOK As Boolean
    
    If bytFunc = 0 Then
        With gclsPros.CurrentForm
            If .txtInfo(GC_������).Text = "" Or .txtSpecificInfo(SLC_סԺ��).Text = "" Or .txtInfo(GC_����).Text = "" Then
               MsgBox "����¼����ҳ�еĲ�����,סԺ��,�����Ȼ�����Ϣ!", vbInformation, gstrSysName
               Exit Sub
            End If
            Call grsDeliceryInfo.AddNew(Array("��Ϣ��", "��Ϣֵ", "����"), Array("������", .txtInfo(GC_������).Text, 1))
            Call grsDeliceryInfo.AddNew(Array("��Ϣ��", "��Ϣֵ", "����"), Array("סԺ��", .txtSpecificInfo(SLC_סԺ��).Text, 1))
            Call grsDeliceryInfo.AddNew(Array("��Ϣ��", "��Ϣֵ", "����"), Array("����", .txtInfo(GC_����).Text, 1))
    
            strTmp = .mskDateInfo(DC_��Ժʱ��).Text
            If Not IsDate(strTmp) Then strTmp = ""
            Call grsDeliceryInfo.AddNew(Array("��Ϣ��", "��Ϣֵ", "����"), Array("��Ժ����", strTmp, 1))
            strTmp = .mskDateInfo(DC_��Ժʱ��).Text
            If Not IsDate(strTmp) Then strTmp = ""
            Call grsDeliceryInfo.AddNew(Array("��Ϣ��", "��Ϣֵ", "����"), Array("��Ժ����", strTmp, 1))
            LngRow = FindDiagRow(DT_��Ժ���XY)
            str��� = .vsDiagXY.TextMatrix(LngRow, DI_��ϱ���) & Space(2) & .vsDiagXY.TextMatrix(LngRow, DI_�������)
            Call grsDeliceryInfo.AddNew(Array("��Ϣ��", "��Ϣֵ", "����"), Array("��Ҫ���", str���, 1))
            Call frmDeliceryInfo.EditDelivery(gclsPros.CurrentForm, gclsPros.����ID, gclsPros.��ҳID, Val(.vsDiagXY.TextMatrix(LngRow, DI_����ID)), gclsPros.OpenMode <> EM_����, grsDeliceryInfo, grsBabyInfo, grsBabyDiag, blnOK)
            If blnOK Then
                Call CheckValueChange
            End If
        End With
    ElseIf bytFunc = 1 Then
        If gclsPros Is Nothing Then
            Set gclsPros = New clsProperty
        End If
        
        Set gclsPros.CurrentForm = objFrmMain
        gclsPros.����ID = lngPatiID
        gclsPros.��ҳID = lngMainID
        Set grsDeliceryInfo = zlDatabase.CopyNewRec(GetPatiAuxiInfoData(lngPatiID, lngMainID, , 2), , "��Ϣ��,��Ϣֵ,��Ϣֵ ��Ϣ��ֵ", Array("����", adInteger, 1, 0, "��¼����", adInteger, 1, Empty))
        Do While Not grsDeliceryInfo.EOF
            grsDeliceryInfo!���� = 0
            grsDeliceryInfo.MoveNext
        Loop
        Set grsBabyDiag = zlDatabase.CopyNewRec(GetBabyDiagData(lngPatiID, lngMainID), , , Array("��¼����", adInteger, 1, Empty))
        Set grsBabyInfo = zlDatabase.CopyNewRec(GetBabyInfoData(lngPatiID, lngMainID), , , Array("��¼����", adInteger, 1, Empty))
        Call frmDeliceryInfo.EditDelivery(objFrmMain, lngPatiID, lngMainID, 0, True, grsDeliceryInfo, grsBabyInfo, grsBabyDiag, , 2)
    End If
End Sub

Public Sub CmdDeliceryInfoGotFocus()
'cmdDeliceryInfo_GotFocus�¼�
    Call ShowInfectInfo(False)
End Sub

'cmdPrint�¼�
Public Sub CmdPrintClick()
'cmdPrint_Click�¼�
    Call PageOperate(MOP_��ӡ)
End Sub

Public Sub CmdPrintGotFocus()
'cmdPrint_GotFocus�¼�
    Call ShowInfectInfo(False)
End Sub

'cmdPrintdown�¼�
Public Sub CmdPrintdownGotFocus()
'cmdPrintdown_GotFocus�¼�
    Call ShowInfectInfo(False)
End Sub

'cmdPriviewDown
Public Sub CmdPriviewDownGotFocus()
'cmdPriviewDown_GotFocus�¼�
    Call ShowInfectInfo(False)
End Sub

'cmdSign�¼�
Public Sub CmdSignClick(ByRef intIndex As Integer)
'���ܣ�cmdSign_Click�¼���ǩ��
    If gclsPros.CurrentForm.cmdSign(intIndex).Caption = "ǩ��" Then
        Call SetSign(intIndex)              'ǩ��
    Else
        Call SetSign(intIndex, True)        'ȡ��ǩ��
    End If
End Sub

'ManInfo�¼���װ
'ManInfo�¼���װ
Public Sub ManInfoClick(ByRef intIndex As Integer)
'���ܣ�cboManInfo_Click�¼���װ
    Dim rsTmp As ADODB.Recordset, rsInput As ADODB.Recordset
    Dim cboTmp As ComboBox
    Dim intIdx As Integer
    Dim blnRestore As Boolean
    Dim strTmp As String

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    If gclsPros.CurrentForm.Visible Then
        If cboTmp.ListIndex <> -1 Then
            If cboTmp.ItemData(cboTmp.ListIndex) = -1 Then
                If gclsPros.FuncType = fҽ����ҳ Then Set gclsPros.ManInfo = Nothing   '��ջ��棬���¶�ȡ����
                Set rsInput = zlDatabase.CopyNewRec(GetManData(intIndex), , "ID,����,���� ƴ������,��ʼ���,���� ����,ȱʡ")
                If rsInput.RecordCount <> 0 Then
                    '��������б�չ����ر������б�
                    If SendMessage(cboTmp.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then
                        SendMessageLong cboTmp.hwnd, CB_SHOWDROPDOWN, False, 0
                    End If
                    If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, cboTmp, rsInput, True, , , rsTmp) Then
                        If cboTmp.ListCount = 0 Then Call SetCboFromRec(intIndex, 1)
                        intIdx = Cbo.FindIndex(cboTmp, rsTmp!ID)
                        If intIdx = -1 Then
                            cboTmp.AddItem rsTmp!���� & "-" & Chr(13) & rsTmp!����, cboTmp.ListCount - 1
                            cboTmp.ItemData(cboTmp.NewIndex) = rsTmp!ID
                            intIdx = cboTmp.NewIndex
                        End If
                        Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
                        cboTmp.Tag = cboTmp.ListIndex
                    Else
                        blnRestore = True
                    End If
                Else
                    MsgBox "û��סԺҽ����ʿ�����ݣ����ȵ�����/��Ա���������á�", vbInformation, gstrSysName
                    blnRestore = True
                End If
            Else
                cboTmp.Tag = cboTmp.ListIndex
            End If
            '�ָ������е���Ա(������Click)
            If blnRestore Then
                If Val(cboTmp.Tag) <> -1 Then
                    Call zlControl.CboSetIndex(cboTmp.hwnd, Val(cboTmp.Tag))
                End If
            Else
                'ҽʦ����,ˢ��ǩ��״̬
                If gclsPros.FuncType = fҽ����ҳ Then
                    gclsPros.IsSigned = SetSignature()
                    Call SetFaceEditable(gclsPros.IsSigned)
                End If
            End If
        End If
    End If
    
    Call CheckValueChange(cboTmp)
End Sub

Public Sub ManInfoDropDown(ByRef intIndex As Integer)
'���ܣ�cboManInfo_DropDown�¼���װ
    Dim strTmp As String
    Dim cboTmp As ComboBox
    Dim intIdx As Integer

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    If cboTmp.Tag = "" Then
       Call ManInfoGotFocus(intIndex)
    End If
    strTmp = cboTmp.Text
    If cboTmp.ListCount = 0 Then
        Call SetCboFromRec(intIndex, 1, , IIf(gclsPros.FuncType = fҽ����ҳ, "[����...]", "NULL"))
    End If
    If strTmp <> "" Then
        intIdx = Cbo.FindIndex(cboTmp, strTmp)
        If intIdx = -1 Then
            Call SetCboFromName(strTmp, cboTmp, "��Ա")
        Else
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
        End If
    End If
    cboTmp.Tag = cboTmp.ListIndex
    Call TxtGotFocus(cboTmp, True, True)
End Sub

Public Function ManInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
    If intIndex = MC_������ Then
        If intKeyAscii = vbKeyPageDown Or intKeyAscii = vbKeyPageUp Then intKeyAscii = 0
    End If
End Function

Public Sub ManInfoGotFocus(ByRef intIndex As Integer)
'���ܣ�cboManInfo_GotFocus�¼���װ
    Dim strTmp As String
    Dim cboTmp As ComboBox
    Dim intIdx As Integer
    Dim blnAdd As Boolean
    
    Call ChangeCtl

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    strTmp = cboTmp.Text
    If cboTmp.ListCount = 0 Then
        blnAdd = True
        Call SetCboFromRec(intIndex, 1, , IIf(gclsPros.FuncType = fҽ����ҳ, "[����...]", "NULL"))
    End If
    If strTmp <> "" Then
        intIdx = Cbo.FindIndex(cboTmp, strTmp)
        If intIdx = -1 Then
            blnAdd = True
            Call SetCboFromName(strTmp, cboTmp, "��Ա", blnAdd)
        Else
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
        End If
    End If
    cboTmp.Tag = cboTmp.ListIndex
    Call TxtGotFocus(cboTmp, True, True)
End Sub

Public Sub ManInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'���ܣ�cboManInfo_KeyPress�¼���װ
    Dim cboTmp As ComboBox
    Dim strInput As String, strFilter As String
    Dim rsTmp As ADODB.Recordset, rsInput As ADODB.Recordset
    Dim lngIdx As Integer
    Dim blnRestore As Boolean
    Dim strTmp As String

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    If intKeyAscii = vbKeyReturn Then
        strTmp = cboTmp.Text
        If cboTmp.ListCount = 0 Then
            Call SetCboFromRec(intIndex, 1, , IIf(gclsPros.FuncType = fҽ����ҳ, "[����...]", "NULL"))
        End If
        If strTmp <> "" Then
            lngIdx = Cbo.FindIndex(cboTmp, strTmp)
            If lngIdx = -1 Then
                Call SetCboFromName(strTmp, cboTmp, "��Ա")
            Else
                Call zlControl.CboSetIndex(cboTmp.hwnd, lngIdx)
                Call CheckValueChange(cboTmp)
            End If
        End If
        '����ƥ����
        strInput = Trim(cboTmp.Text)
        If strInput = "" Then
            '���֮ǰ�����ݣ�����ɾ��֮�󣬱��水ťҲӦ�ñ�ÿ��ã�ǩ����ťҲӦ�÷����仯
            If cboTmp.Tag >= 0 Then
                cboTmp.Tag = cboTmp.ListIndex
                Call CheckValueChange(cboTmp)
                 'ҽʦ����,ˢ��ǩ��״̬
                If gclsPros.FuncType = fҽ����ҳ Then
                    gclsPros.IsSigned = SetSignature()
                    Call SetFaceEditable(gclsPros.IsSigned)
                End If
            End If
            zlCommFun.PressKey vbKeyTab: mblnReturn = True
            Exit Sub
        End If
        '��ͬ����Ŀ�򲻽��д���
        If cboTmp.ListIndex <> -1 Then
            If cboTmp.Tag <> cboTmp.ListIndex Then
                cboTmp.Tag = cboTmp.ListIndex
                'ҽʦ����,ˢ��ǩ��״̬
                If gclsPros.FuncType = fҽ����ҳ Then
                    gclsPros.IsSigned = SetSignature()
                    Call SetFaceEditable(gclsPros.IsSigned)
                End If
            End If
            
            If zlStr.NeedName(strInput) = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex)) Then
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
                Exit Sub
            End If
        End If

        strInput = UCase(strInput)
        strFilter = "���� Like '" & strInput & "*' Or ���� Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*' Or " & IIf(gclsPros.BriefCode = 0, "����", "��ʼ���") & " Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*'"
        Set rsInput = Rec.FilterNew(GetManData(intIndex), strFilter, "ID,����,���� ƴ������,��ʼ���,���� ����,ȱʡ")
        If rsInput.RecordCount = 0 Then
            MsgBox "δ�ҵ���Ӧ��ҽ����ʿ��", vbInformation, gstrSysName
            blnRestore = True
        Else
            '��������б�չ����ر������б�
            If SendMessage(cboTmp.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then
                SendMessageLong cboTmp.hwnd, CB_SHOWDROPDOWN, False, 0
            End If
            If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, cboTmp, rsInput, True, , , rsTmp) Then
                If cboTmp.ListCount = 0 Then Call SetCboFromRec(intIndex, 1)
                lngIdx = Cbo.FindIndex(cboTmp, rsTmp!ID)
                If lngIdx = -1 Then
                    cboTmp.AddItem rsTmp!���� & "-" & Chr(13) & rsTmp!����, cboTmp.ListCount - 1
                    cboTmp.ItemData(cboTmp.NewIndex) = rsTmp!ID
                    lngIdx = cboTmp.NewIndex
                End If

                Call zlControl.CboSetIndex(cboTmp.hwnd, lngIdx)
                cboTmp.Tag = cboTmp.ListIndex
                cboTmp.Text = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex))
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
            Else
                blnRestore = True
            End If
        End If
        '�ָ������е���Ա(������Click)
        If blnRestore Then
            If Val(cboTmp.Tag) <> -1 Then
                Call zlControl.CboSetIndex(cboTmp.hwnd, Val(cboTmp.Tag))
            End If
            cboTmp.Text = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex))
        Else
            Call CheckValueChange(cboTmp)
            'ҽʦ����,ˢ��ǩ��״̬
            If gclsPros.FuncType = fҽ����ҳ Then
                gclsPros.IsSigned = SetSignature()
                Call SetFaceEditable(gclsPros.IsSigned)
            End If
        End If
    End If
End Sub

Public Sub ManInfoLostFocus(ByRef intIndex As Integer)
'���ܣ�cboManInfo_LostFocus�¼���װ
    Dim intIdx As Integer, i As Long
    Dim blnHave As Boolean
    Dim cboTmp As ComboBox

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    If cboTmp.ListIndex >= 0 Then
       If cboTmp.Text <> zlStr.NeedName(cboTmp.List(cboTmp.ListIndex)) Then
           cboTmp.Text = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex))
       End If
    End If
End Sub

Public Sub ManInfoValidate(ByRef intIndex As Integer, ByRef blnCancel As Boolean)
'���ܣ�cboManInfo_Validate�¼���װ
    Dim strInput As String, strFilter As String
    Dim rsTmp As ADODB.Recordset, rsInput As ADODB.Recordset
    Dim cboTmp As ComboBox
    Dim intIdx As Integer
    Dim strTmp As String

    Set cboTmp = gclsPros.CurrentForm.cboManInfo(intIndex)
    strTmp = cboTmp.Text
    If cboTmp.ListCount = 0 Then
        Call SetCboFromRec(intIndex, 1, , IIf(gclsPros.FuncType = fҽ����ҳ, "[����...]", "NULL"))
    End If
    If strTmp <> "" Then
        intIdx = Cbo.FindIndex(cboTmp, strTmp)
        If intIdx = -1 Then
            Call SetCboFromName(strTmp, cboTmp, "��Ա")
        Else
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
        End If
        cboTmp.Tag = cboTmp.ListIndex
    Else
        cboTmp.Tag = cboTmp.ListIndex
        Exit Sub '������
    End If
    If cboTmp.ListIndex <> -1 Then cboTmp.Text = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex)): Exit Sub '��ѡ��

    strInput = UCase(zlStr.NeedName(cboTmp.Text))
    strFilter = "���� Like '" & strInput & "*' Or ���� Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*' Or " & IIf(gclsPros.BriefCode = 0, "����", "��ʼ���") & " Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*'"
    Set rsInput = Rec.FilterNew(GetManData(intIndex), strFilter, "ID,����,���� ƴ������,��ʼ���,���� ����,ȱʡ")
    If rsInput.RecordCount <> 0 Then
        '��������б�չ����ر������б�
        If SendMessage(cboTmp.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then
            SendMessageLong cboTmp.hwnd, CB_SHOWDROPDOWN, False, 0
        End If
        If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, cboTmp, rsInput, True, , , rsTmp) Then
            intIdx = Cbo.FindIndex(cboTmp, rsTmp!ID)
            If intIdx = -1 Then
                cboTmp.AddItem rsTmp!���� & "-" & Chr(13) & rsTmp!����, cboTmp.ListCount - 1
                cboTmp.ItemData(cboTmp.NewIndex) = rsTmp!ID
                intIdx = cboTmp.NewIndex
            End If
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
            cboTmp.Tag = cboTmp.ListIndex
            cboTmp.Text = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex))
        Else
            blnCancel = True: cboTmp.Text = ""
        End If
    Else
        MsgBox "δ�ҵ���Ӧ��ҽ����ʿ��", vbInformation, gstrSysName
        blnCancel = True: cboTmp.Text = ""
    End If
End Sub

'DateInfo�¼���װ
Public Sub DateInfoChange(ByRef intIndex As Integer)
'���ܣ�MskDateInfo_Change
    Select Case intIndex
        Case DC_��������
            Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_����ʱ��), Not IsDate(gclsPros.CurrentForm.mskDateInfo(intIndex).Text))
    End Select
    Call CheckValueChange(gclsPros.CurrentForm.mskDateInfo(intIndex))
End Sub

Public Sub DateInfoClick(ByRef intIndex As Integer)
'���ܣ�cmdDateInfo_Click
    Dim objmonInfo As MonthView  '������ÿؼ�����
    Dim objCmd As CommandButton
    Dim objMSK As MaskEdBox
    Dim datStart As Date
    Dim dateEnd As Date
    Dim datTmp As Date
    On Error GoTo errH
    gclsPros.DateIndex = intIndex
    With gclsPros.CurrentForm
        Set objmonInfo = .monInfo
        Set objCmd = .cmdDateInfo(intIndex)
        Set objMSK = .mskDateInfo(intIndex)
        If IsDate(gclsPros.InTime) Then
            datStart = CDate(gclsPros.InTime)
        End If
        If IsDate(gclsPros.OutTime) Then
            dateEnd = CDate(gclsPros.OutTime)
        Else
            dateEnd = zlDatabase.Currentdate
        End If
        objmonInfo.MinDate = 0
        objmonInfo.MaxDate = zlDatabase.Currentdate
        Select Case intIndex
            Case DC_��������
                objmonInfo.MaxDate = datStart
            Case DC_��Ժʱ��
                objmonInfo.MaxDate = dateEnd
            Case DC_��Ժʱ��
                objmonInfo.MinDate = datStart
            Case DC_ȷ������
                objmonInfo.MinDate = datStart
                objmonInfo.MaxDate = dateEnd
            Case DC_����ʱ��
                objmonInfo.MinDate = datStart
                objmonInfo.MaxDate = zlDatabase.Currentdate
            Case DC_��������
                objmonInfo.MaxDate = dateEnd
            Case DC_��Ŀ����, DC_�ջ�����
                objmonInfo.MinDate = dateEnd
                objmonInfo.MaxDate = zlDatabase.Currentdate
            Case DC_�ʿ�����
                objmonInfo.MinDate = datStart
                objmonInfo.MaxDate = CDate("3000-01-01")
'                objmonInfo.Value = zlDatabase.Currentdate
        End Select
        If IsDate(objMSK.Text) Then
            datTmp = CDate(objMSK.Text)
            If datTmp > objmonInfo.MaxDate Then
                datTmp = objmonInfo.MaxDate
            ElseIf datTmp < objmonInfo.MinDate Then
                datTmp = objmonInfo.MinDate
            End If
            objmonInfo.Value = datTmp
        End If
        objmonInfo.Left = objCmd.Left + objCmd.Width - objmonInfo.Width + objMSK.Container.Left + .PicPage(0).Left
        If intIndex = DC_�������� Then
            objmonInfo.Top = objCmd.Top + objCmd.Height + 20 + objMSK.Container.Top
        Else
            objmonInfo.Top = objCmd.Top - objmonInfo.Height - 20 + objMSK.Container.Top
        End If
        objmonInfo.ZOrder
        objmonInfo.Visible = True
        objmonInfo.SetFocus
    End With
    Exit Sub
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
End Sub

Public Sub DateInfoGotFocus(ByRef intIndex As Integer)
'���ܣ�MskDateInfo_GotFocus
    Dim objMSK As MaskEdBox
    '�������벻����������
    Call ChangeCtl
    zlCommFun.OpenIme False
End Sub

Public Sub DateInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'���ܣ�MskDateInfo_KeyDown
    If intKeyCode = vbKeyF4 Or (intKeyCode = vbKeyDown And intShift = vbAltMask) Then
        Call DateInfoClick(gclsPros.CurrentForm.cmdDateInfo(intIndex))
    End If
End Sub

Public Sub DateInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'���ܣ�MskDateInfo_KeyPress
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

Public Sub DateInfoValidate(ByRef intIndex As Integer, ByRef blnCancel As Boolean)
'���ܣ�MskDateInfo_Validate
    Dim objMSK As MaskEdBox
    Dim str���� As String
    Dim str��Ժʱ�� As String

    With gclsPros.CurrentForm
        Set objMSK = .mskDateInfo(intIndex)
        If Not IsDate(objMSK.Text) And objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
            Call ShowMessage(objMSK, "�������ʱ�䲻����Ч��ʱ�䣬���������롣")
            blnCancel = True
            Exit Sub
        End If
        Select Case intIndex
            Case DC_��������

            Case DC_��������, DC_����ʱ��, DC_����ʱ��, DC_ȷ������
                If objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
                    If intIndex = DC_ȷ������ Then
                        If objMSK.Text <> Replace(objMSK.Tag, "#", "_") Then
                            If Not CheckDateRange(objMSK.Text, True) Then
                                Call ShowMessage(objMSK, "�������ʱ�䲻�ڲ������Ժʱ�䷶Χ�ڣ����������롣")
                                blnCancel = True
                            End If
                        End If
                    End If
                End If
            Case DC_��Ժʱ��, DC_��Ժʱ��
                If gclsPros.InTime = "" And intIndex = DC_��Ժʱ�� Then
                    If IsDate(.mskDateInfo(DC_��Ժʱ��).Text) Then
                        gclsPros.InTime = .mskDateInfo(DC_��Ժʱ��).Text
                    End If
                ElseIf intIndex = DC_��Ժʱ�� And gclsPros.OutTime = "" Then
                    If IsDate(.mskDateInfo(DC_��Ժʱ��).Text) Then
                        gclsPros.InTime = .mskDateInfo(DC_��Ժʱ��).Text
                    End If
                End If
                If objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
                    If intIndex = DC_��Ժʱ�� And gclsPros.InTime <> "" Then
                        If CDate(gclsPros.InTime) > CDate(objMSK.Text) Then
                            Call ShowMessage(objMSK, "������ĳ�Ժʱ��С����Ժʱ�䣬���������롣")
                            Call DateInfoClick(intIndex)
                        ElseIf CDate(objMSK.Text) > zlDatabase.Currentdate Then
                            Call ShowMessage(objMSK, "������ĳ�Ժʱ����ڵ�ǰʱ�䣬���������롣")
                            Call DateInfoClick(intIndex)
                        Else
                            gclsPros.OutTime = objMSK.Text
                        End If
                    ElseIf intIndex = DC_��Ժʱ�� And gclsPros.OutTime <> "" Then
                        If CDate(gclsPros.OutTime) < CDate(objMSK.Text) Then
                            Call ShowMessage(objMSK, "���������Ժʱ����ڳ�Ժʱ�䣬���������롣")
                            Call DateInfoClick(intIndex)
                        ElseIf CDate(objMSK.Text) > zlDatabase.Currentdate Then
                            Call ShowMessage(objMSK, "���������Ժʱ����ڵ�ǰʱ�䣬���������롣")
                            Call DateInfoClick(intIndex)
                        Else
                            gclsPros.InTime = objMSK.Text
                        End If
                    End If
                End If
            Case DC_��Ŀ����
                If objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
                    If CDate(Format(gclsPros.OutTime, "yyyy-mm-dd")) > CDate(objMSK.Text) And gclsPros.OutTime <> "" Then
                        Call ShowMessage(objMSK, "������ı�Ŀ����С�ڳ�Ժʱ�䣬���������롣")
                        Call DateInfoClick(intIndex)
                    ElseIf CDate(objMSK.Text) > zlDatabase.Currentdate Then
                        Call ShowMessage(objMSK, "������ı�Ŀ���ڴ��ڵ�ǰʱ�䣬���������롣")
                        Call DateInfoClick(intIndex)
                    End If
                End If
            Case DC_�ʿ�����
                If objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
                    If IsDate(gclsPros.InTime) Then
                        If CDate(Format(gclsPros.InTime, "yyyy-mm-dd")) > CDate(objMSK.Text) Then
                            Call ShowMessage(objMSK, "��������ʿ�����С����Ժʱ�䣬���������롣")
                            Call DateInfoClick(intIndex)
                        End If
                    End If
                End If
        End Select
    End With
End Sub

'chkInfo�¼�
Public Sub chkInfoClick(ByRef intIndex As Integer)
'���ܣ�chkInfo_Click
    Dim blnCheck As Boolean

    With gclsPros.CurrentForm
         blnCheck = .chkInfo(intIndex).Value = 1
        Select Case intIndex
            Case CHK_�Ƿ�ȷ��
                Call SetCtrlLocked(.mskDateInfo(DC_ȷ������), Not blnCheck, True)
                Call SetCtrlLocked(.cmdDateInfo(DC_ȷ������), Not blnCheck, True)
            Case CHK_����
                Call SetCtrlLocked(.txtSpecificInfo(SLC_��������), Not blnCheck, True)
                Call SetCtrlLocked(.cboSpecificInfo(SLC_��������), Not blnCheck, True)
                If blnCheck Then
                    Call CboSpecificInfoClick(SLC_��������)
                End If
            Case CHK_��ԭѧ���
                If Not blnCheck Then
                    .txtInfo(GC_��ԭѧ���).Tag = ""
                    .cmdInfo(GC_��ԭѧ���).Tag = ""
                End If
                Call SetCtrlLocked(.txtInfo(GC_��ԭѧ���), Not blnCheck, True)
                Call SetCtrlLocked(.cmdInfo(GC_��ԭѧ���), Not blnCheck)
            Case CHK_����
                If gclsPros.PathVCauses Then
                    Call SetCtrlLocked(.cboBaseInfo(BCC_����ԭ��), Not blnCheck, True)
                Else
                    Call SetCtrlLocked(.txtInfo(GC_����ԭ��), Not blnCheck, True)
                End If
            Case CHK_����·��
                Call SetCtrlLocked(.chkInfo(CHK_����), Not blnCheck, True)
                Call SetCtrlLocked(.chkInfo(CHK_���·��), Not blnCheck, True)
                Call SetCtrlLocked(.txtInfo(GC_�˳�ԭ��), Not blnCheck, True)
                If Not blnCheck Then
                    If gclsPros.PathVCauses Then
                        Call SetCtrlLocked(.cboBaseInfo(BCC_����ԭ��), Not blnCheck, True)
                    Else
                        Call SetCtrlLocked(.txtInfo(GC_����ԭ��), Not blnCheck, True)
                    End If
                End If
            Case CHK_���·��
                Call SetCtrlLocked(.txtInfo(GC_�˳�ԭ��), blnCheck, True)
            Case CHK_סԺ����Լ��
                If gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                    Call SetCtrlLocked(.txtSpecificInfo(SLC_Լ����ʱ��), Not blnCheck, True)
                    Call SetCtrlLocked(.cboBaseInfo(BCC_Լ����ʽ), Not blnCheck, True)
                    Call SetCtrlLocked(.cboBaseInfo(BCC_Լ������), Not blnCheck, True)
                    Call SetCtrlLocked(.cboBaseInfo(BCC_Լ��ԭ��), Not blnCheck, True)
                End If
            Case CHK_�������
                Call SetCtrlLocked(.txtSpecificInfo(SLC_Ժ�ڻ���), Not blnCheck, True)
                Call SetCtrlLocked(.txtSpecificInfo(SLC_��Ժ����), Not blnCheck, True)
                Call SetCtrlLocked(.txtInfo(GC_��������), Not blnCheck, True)
            Case CHK_�޹�����¼
                If mblnChk = False Then
                    If .vsAller.TextMatrix(.vsAller.FixedRows, AI_����ҩ��) <> "" And .vsAller.TextMatrix(.vsAller.FixedRows, AI_����ҩ��) <> "��" Then
                        If blnCheck Then
                            MsgBox "�Ѿ��й���ҩ����ܱ��Ϊ�ޡ�", vbInformation, gstrSysName
                            mblnChk = True
                            .chkInfo(intIndex).Value = 0
                            Exit Sub
                        End If
                    End If
                    Call SetCtrlLocked(.vsAller, blnCheck)
                    .vsAller.TextMatrix(.vsAller.FixedRows, AI_����ҩ��) = IIf(blnCheck, "��", "")
                End If
                mblnChk = False
        End Select
        Call CheckValueChange(.chkInfo(intIndex))
    End With
End Sub

Public Sub ChkInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'���ܣ�ChkInfo_KeyPress
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

Public Sub ChkInfoGotFocus(ByRef intIndex As Integer)
'���ܣ�ChkInfo_GotFocus
    'ҽԺ��Ⱦ��ؿؼ��ɼ����Լ�λ��
    Call ChangeCtl
    Call ShowInfectInfo(False)
End Sub

'chkFeeEdit�¼�
Public Sub ChkFeeEditClick()
'���ܣ�ChkFeeEdit_Click
    Call SetCtrlLocked(gclsPros.CurrentForm.vsFees, gclsPros.CurrentForm.chkFeeEdit.Value = 0)
End Sub

Public Sub ChkFeeEditKeyPress(ByRef intKeyAscii As Integer)
'���ܣ�ChkFeeEdit_KeyPress

    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        If gclsPros.CurrentForm.chkFeeEdit.Value = 1 Then
            gclsPros.CurrentForm.vsFees.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        End If
    End If
End Sub

'optInput�¼�
Public Sub OptInputClick(ByRef intIndex As Integer)
'optInput_Click�¼�
    With gclsPros.CurrentForm
        Select Case intIndex
            Case OP_��סԺ��, OP_��סԺ��
                Call SetCtrlLocked(.txtInfo(GC_31������סԺ), intIndex = OP_��סԺ��, True)
            Case OP_ICU��, OP_ICU��
                Call SetCtrlLocked(.txtSpecificInfo(SLC_��֢�໤��), intIndex = OP_ICU��, True)
                Call SetCtrlLocked(.txtSpecificInfo(SLC_��֢�໤Сʱ), intIndex = OP_ICU��, True)
        End Select
    End With
    Call CheckValueChange
End Sub

Public Sub OptInputKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'optInput_KeyPress�¼�
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

'optDiag�¼�
Public Sub optDiagClick(ByRef intIndex As Integer)
    If gclsPros.PatiType = PF_���� Then
        gclsPros.DiagInputXY = intIndex Mod 2
        gclsPros.DiagInputZY = gclsPros.DiagInputXY
    Else
        If intIndex < 2 Then
            gclsPros.DiagInputXY = intIndex Mod 2
        Else
            gclsPros.DiagInputZY = intIndex Mod 2
        End If
    End If
    Call CheckValueChange
End Sub

Public Sub optDiagGotFocus(ByRef intIndex As Integer)
'optDiag_GotFocus�¼�
    Call ChangeCtl
    Call ShowInfectInfo(False)
End Sub

Public Sub optDiagKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'optDiag_KeyPress�¼�
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

'optState�¼�����������л���������ҳ
Public Sub optStateClick(ByRef intIndex As Integer)
    Dim blnDo As Boolean
    Dim rsTmp As ADODB.Recordset

    '����������δ¼�����������Զ���ȡ�ϴ����
    If intIndex = OP_���� Then
        With gclsPros.CurrentForm
            If .chkInfo(CHK_��Ⱦ���ϴ�).Value = 1 Then
                blnDo = .vsDiagXY.Rows = .vsDiagXY.FixedRows + 1 And .vsDiagZY.Rows = .vsDiagZY.FixedRows + 1
                If blnDo Then blnDo = blnDo And .vsDiagXY.TextMatrix(.vsDiagXY.FixedRows, DI_�������) = "" And .vsDiagZY.TextMatrix(.vsDiagZY.FixedRows, DI_�������) = ""
                If blnDo Then
                    Set rsTmp = GetPatiDiagData(gclsPros.����ID, gclsPros.��ҳID, 0, True, , gclsPros.Moved)
                    If rsTmp.RecordCount <> 0 Then gclsPros.Is���� = True: gclsPros.IsLastDiag = True
                    Call CacheLoadVsDiagData(.vsDiagXY, rsTmp, DT_�������XY, , -1)
                    If gclsPros.Have��ҽ Then
                        Call CacheLoadVsDiagData(.vsDiagZY, rsTmp, DT_�������ZY, , -1)
                    End If
                End If
            End If
        End With
    End If
End Sub

'optAller�¼�
Public Sub OptAllerClick(ByRef intIndex As Integer)
'optAller_Click�¼�
    If intIndex = PC_��ҩƷĿ¼���� Then
        gclsPros.AllerInput = 0
        gclsPros.UseTYT = False
    Else
        If Not gobjPass Is Nothing Then
            gclsPros.AllerInput = 1
            gclsPros.UseTYT = True
        Else
            gclsPros.AllerInput = 1
        End If
    End If
    Call CheckValueChange
End Sub

Public Sub OptAllerKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'optAller_KeyPress�¼�
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

'optParaOPSInfo�¼�
Public Sub OptParaOPSInfoClick(ByRef intIndex As Integer)
'optParaOPSInfo_Click�¼�
    If intIndex = PC_��������Ŀ���� Then
        gclsPros.OPSInput = 0
    Else
        gclsPros.OPSInput = 1
    End If
End Sub

Public Sub OptParaOPSInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'OptParaOPSInfo_KeyPress�¼�
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

'chkParaOPSInfo�¼�
Public Sub ChkParaOPSInfoClick(ByRef intIndex As Integer)
'chkParaOPSInfo_Click�¼�
    If gclsPros.CurrentForm.chkParaOPSInfo(PC_δ�ҵ�ʱ����¼��).Value = 1 Then
        gclsPros.OPSFree = True
    Else
        gclsPros.OPSFree = False
    End If
    Call CheckValueChange
End Sub

Public Sub ChkParaOPSInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'chkParaOPSInfo_KeyPress�¼�
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

'SpecificInfo�¼�
Public Sub SpecificInfoChange(ByRef intIndex As Integer)
'���ܣ�txtSpecificInfo_Change
    Dim objTextBox As TextBox
    Dim objComboBox As ComboBox

    With gclsPros.CurrentForm
        Select Case intIndex
            Case SLC_����, SLC_Ӥ�׶�����
            '��������Ŵ���׼���䵥λ
                Set objTextBox = .txtSpecificInfo(intIndex)
                Set objComboBox = .cboSpecificInfo(intIndex)
                If IsNumeric(objTextBox.Text) Or objTextBox.Text = "" Then
                    objComboBox.Visible = True
                    objComboBox.Tag = ""
                    If objComboBox.Container.Name = "fraCbo" Then
                        objComboBox.Container.Visible = True
                    End If
                
                    If intIndex = SLC_���� Then
                        If gclsPros.FuncType = f������ҳ Then
                            objTextBox.Width = 450
                        Else
                            objTextBox.Width = 360
                        End If
                    ElseIf intIndex = SLC_Ӥ�׶����� Then
                        DrawLineCTL objTextBox, 1
                        objTextBox.Width = 360
                        DrawLineCTL objTextBox
                    End If
                    If objComboBox.ListIndex = -1 Then objComboBox.ListIndex = 0
                Else
                    objComboBox.Visible = False
                    objComboBox.Tag = "����"
                    objComboBox.ListIndex = -1
                    If objComboBox.Container.Name = "fraCbo" Then
                        objComboBox.Container.Visible = False
                    End If
                
                    If intIndex = SLC_���� Then
                        If gclsPros.FuncType = f������ҳ Then
                            objTextBox.Width = 1250
                        Else
                            objTextBox.Width = 1150
                        End If
                    ElseIf intIndex = SLC_Ӥ�׶����� Then
                        DrawLineCTL objTextBox, 1
                        objTextBox.Width = 1250
                        DrawLineCTL objTextBox
                    End If
                End If
            Case SLC_���ȴ���
                Set objTextBox = .txtSpecificInfo(intIndex)
                Call SetCtrlLocked(.txtInfo(GC_���Ȳ���), Val(objTextBox.Text) = 0, True)
                Call SetCtrlLocked(.cmdInfo(GC_���Ȳ���), Val(objTextBox.Text) = 0, True)
                Call SetCtrlLocked(.txtSpecificInfo(SLC_�ɹ�����), Val(objTextBox.Text) = 0, True)
                If Val(objTextBox.Text) > 0 Then
                    '��Ҫ��ϵĳ�Ժ�����Ϊ����ʱ,ȱʡ���ɹ�����=���ȴ���
                    If .Visible Then
                        If .vsDiagXY.TextMatrix(FindDiagRow(DT_��Ժ���XY), DI_��Ժ���) <> "����" Then
                            .txtSpecificInfo(SLC_�ɹ�����).Text = objTextBox.Text
                        ElseIf Val(objTextBox.Text) > 1 Then
                            .txtSpecificInfo(SLC_�ɹ�����).Text = Val(objTextBox.Text) - 1
                        End If
                    End If
                End If
        End Select
        Call CheckValueChange(.txtSpecificInfo(intIndex))
    End With
End Sub

Public Sub SpecificInfoClick(ByRef intIndex As Integer, Optional ByVal blnCmdButton As Boolean)
'CmdSpecificInfo_Click�¼�
'������blnCmdButton=True-cmdButton�ؼ���False-��cmdButton�ؼ�
    Dim blnALLPati As Boolean
    Dim arrDate() As String
    Dim str��ȡ���� As String, strIfdate As String
    Dim blnCancel As Boolean
    Dim vRect As RECT
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim int��ҳid As Integer ''��¼��ҳid
    Dim strKEY As String
    Dim objTextסԺ�� As TextBox
    Dim blnEditInfo As Boolean
    Dim blnSign As Boolean

    On Error GoTo errH

    If blnCmdButton Then
        '����ϵͳ����סԺ��ѡ��
        If intIndex = SLC_סԺ�� Then
            ReDim arrDate(2)
            arrDate(0) = zlDatabase.GetPara("��ʼ����", gclsPros.SysNo, gclsPros.Module)
            arrDate(1) = zlDatabase.GetPara("��������", gclsPros.SysNo, gclsPros.Module)
            blnSign = zlDatabase.GetPara("��ǩ���ĳ�Ժ����", gclsPros.SysNo, gclsPros.Module)
            If arrDate(0) = "" Or arrDate(1) = "" Then
                arrDate(0) = "": arrDate(1) = ""
                blnALLPati = True
            Else
                arrDate(0) = Format(arrDate(0), "yyyy-mm-dd")
                arrDate(1) = Format(arrDate(1), "yyyy-mm-dd")
            End If
            If Not blnALLPati Then blnALLPati = Val(zlDatabase.GetPara("��ȡ���г�Ժ����", gclsPros.SysNo, gclsPros.Module)) = 1

            If Val(zlDatabase.GetPara("��ȡ24Сʱ�ڳ�Ժ����", gclsPros.SysNo, gclsPros.Module)) <> 1 Then
                If gclsPros.OutFile = "" Then
                    str��ȡ���� = " And (B.��Ժ����-B.��Ժ����)*24>=24"
                Else
                    str��ȡ���� = "סԺʱ��>=24"
                End If
            End If

            If gclsPros.EditUnrecive = False And gclsPros.OutFile = "" Then
                str��ȡ���� = str��ȡ���� & " And E.����ʱ�� IS NOT NULL"
            End If
            If Not blnALLPati Then
                If gclsPros.OutFile = "" Then
                    strIfdate = " And B.��Ժ���� Between Trunc(To_Date('" & arrDate(0) & "','yyyy-mm-dd')) And Trunc(To_Date('" & arrDate(1) & "','yyyy-mm-dd'))+1-1/24/60/60"
                Else
                    strIfdate = " ��Ժ���� >= #" & Format(arrDate(0), "yyyy-mm-dd 00:00:00") & "# and ��Ժ���� <= #" & Format(arrDate(1), "yyyy-mm-dd 23:59:59") & "#"
                End If
            End If
            Set objTextסԺ�� = gclsPros.CurrentForm.txtSpecificInfo(SLC_סԺ��)
            vRect = zlControl.GetControlRect(objTextסԺ��.hwnd)
            '39906:������,2013-05-07,��Ӳ������ձ�־
            If gclsPros.OutFile = "" Then
                '����26488 by lesfeng 2010-03-18 ����
                strSql = "" & _
                    " Select    to_char(Id) Id ,�ϼ�id,0 as ��ҳID,ĩ��,����,����,�Ա�,��Ժ����,��Ժ����,סԺ����,����,����,���� as ���id " & _
                    "   From (  Select id as id,�ϼ�id,0 as ĩ��,����,���� ,'' as �Ա�,'' as ��Ժ����,'' as ��Ժ����,'' as סԺ����,'' As ����,'' as ����,max(Level) as ����" & _
                    "           From ���ű� " & _
                    "           Start with id in (  Select distinct(b.��Ժ����id)  as �ϼ�id " & _
                    "                               From ������ҳ b,������Ϣ a,�������ռ�¼ E" & _
                    "                               Where a.����id=b.����id And B.����ID=E.����ID(+) And B.��ҳID=E.��ҳID(+) and b.סԺ�� is not null and b.��Ŀ���� is null and b.��Ժ���� is not null and nvl(b.��������,0)=0  " & _
                                                            strIfdate & str��ȡ���� & _
                    "                             )  Connect by prior �ϼ�id = id group by id,�ϼ�id,ĩ��,����,���� " & _
                    "       )"
                If blnSign Then
                    strSql = strSql & vbNewLine & _
                    "Union All" & _
                    " Select a.����id || '-' || b.��ҳid ,c.id,b.��ҳID,1 as ĩ��, to_char(b.סԺ��) as ����,a.���� as ����,a.�Ա�,to_char(b.��Ժ����,'yyyy-mm-dd'),to_char(b.��Ժ����,'yyyy-mm-dd')," & _
                    "         to_char(Zl_��ȡסԺ��������ҳid(a.����id,b.��ҳid,0)) ,decode(D.�������,null,'��',0,'��','��') As ����,Decode(E.����ʱ��,Null,'��','��') As ����,-9999 as ���id" & _
                    " From ������Ϣ a,������ҳ b,���ű� c,������� D,�������ռ�¼ E " & _
                    " Where a.����ID = b.����ID and B.����id = D.����id(+) And D.����(+)=2 And B.����ID=E.����ID(+) And B.��ҳID=E.��ҳID(+) " & _
                    "     and b.��Ŀ���� is null " & _
                    "     and b.��Ժ���� is not null and b.סԺ�� is not null and nvl(b.��������,0)=0 " & _
                    "     and b.��Ժ����id =c.id " & strIfdate & str��ȡ���� & _
                    "And Exists" & _
                    " (Select *" & vbNewLine & _
                    "       From ������ҳ�ӱ�" & vbNewLine & _
                    "       Where ����id = b.����id And ��ҳid = b.��ҳid And ��Ϣ�� In ('������ǩ��', '����ҽʦǩ��', 'סԺҽʦǩ��', 'סԺҽʦǩ��'))" & vbNewLine & _
                    " order by ���ID desc "
                Else
                    strSql = strSql & vbNewLine & _
                    "Union All" & _
                     " Select a.����id || '-' || b.��ҳid ,c.id,b.��ҳID,1 as ĩ��, to_char(b.סԺ��) as ����,a.���� as ����,a.�Ա�,to_char(b.��Ժ����,'yyyy-mm-dd'),to_char(b.��Ժ����,'yyyy-mm-dd')," & _
                    "         to_char(Zl_��ȡסԺ��������ҳid(a.����id,b.��ҳid,0)) ,decode(D.�������,null,'��',0,'��','��') As ����,Decode(E.����ʱ��,Null,'��','��') As ����,-9999 as ���id" & _
                    " From ������Ϣ a,������ҳ b,���ű� c,������� D,�������ռ�¼ E " & _
                    " Where a.����ID = b.����ID and B.����id = D.����id(+) And D.����(+)=2 And B.����ID=E.����ID(+) And B.��ҳID=E.��ҳID(+) " & _
                    "     and b.��Ŀ���� is null " & _
                    "     and b.��Ժ���� is not null and b.סԺ�� is not null and nvl(b.��������,0)=0 " & _
                    "     and b.��Ժ����id =c.id " & strIfdate & str��ȡ���� & _
                    " order by ���ID desc "
                End If

                    '���˺�:���۲��˲��ܽ�����
                    '39906:������,2013-05-07,��Ҫ��ʾ��Ŀ�������ͽ�������
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 2, "����סԺ��", False, "", "����Ŀ������[Count]�ˣ������ѽ��յĲ�������[����='��']��", False, False, True, vRect.Left, vRect.Top, 300, blnCancel, False, True)
                    If rsTmp Is Nothing Then Exit Sub
                    If rsTmp.State <> 1 Or rsTmp.EOF Then Exit Sub
                    objTextסԺ��.Text = rsTmp!���� & ""
                    int��ҳid = Val(rsTmp!��ҳID & "")
                    If gclsPros.OpenMode <> EM_�༭ Then
                        '����õ�סԺ��������Ͳ�����
                        '78747:��α����Ͷ��࣬��ΪLoadPatiByInNo���Ѿ������Ƽ��
                        'If GetסԺ����Or��ҳid(Val(Split(rsTmp!ID & "", "-")(0)), Val(rsTmp!��ҳID & ""), False, True) = False Then: Exit Sub
                        gclsPros.IsSelPati = True
                        '��סԺ�Ÿı����ղ���¼����Ϣ
                        If Val(objTextסԺ��.Text) <> Val(gclsPros.InNo) Then
                            If Not CheckMedPageChange Then
                                gclsPros.InfosChange = False
                            End If
                            If gclsPros.InfosChange = True And Val(gclsPros.InNo) <> 0 Then
                                gclsPros.InfosChange = False
                                If MsgBox("��Ϣ�ѷ����仯���Ƿ�ȷ�ϸ���¼�벡�ˣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                    Call gclsPros.InitCacheRecInfo
                                ElseIf Val(gclsPros.InNo) <> 0 Then
                                    objTextסԺ��.Text = gclsPros.InNo
                                    Exit Sub
                                End If
                            Else
                                gclsPros.InfosChange = False
                            End If
                        End If
                        Call LoadPatiByInNo(objTextסԺ��.Text, int��ҳid)
                        gclsPros.IsSelPati = False
                        Call AfterLoadPatiByNo
                    End If
            Else
                If str��ȡ���� <> "" Then
                    strIfdate = IIf(strIfdate = "", "", strIfdate & " and ") & str��ȡ����
                End If
                With frmPageMedRecNOSel
                    .Top = vRect.Top + 300
                    .Left = vRect.Left
                    strKEY = .ShowMe(gclsPros.CurrentForm, gclsPros.PatiOut, strIfdate)
                    If strKEY = "" Then Exit Sub
                    objTextסԺ��.Text = Split(strKEY, "_")(0)
                    If Val(objTextסԺ��.Text) = 0 Then
                        objTextסԺ��.Text = ""
                        Exit Sub
                    End If
                    int��ҳid = Split(strKEY, "_")(1)
                    If gclsPros.OpenMode <> EM_�༭ Then
                        '��סԺ�Ÿı����ղ���¼����Ϣ
                        gclsPros.IsSelPati = True
                        '��סԺ�Ÿı����ղ���¼����Ϣ
                        If Val(objTextסԺ��.Text) <> Val(gclsPros.InNo) Then
                            If Not CheckMedPageChange Then
                                gclsPros.InfosChange = False
                            End If
                            If gclsPros.InfosChange = True And Val(gclsPros.InNo) <> 0 Then
                                gclsPros.InfosChange = False
                                If MsgBox("��Ϣ�ѷ����仯���Ƿ�ȷ�ϸ���¼�벡�ˣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                    Call gclsPros.InitCacheRecInfo
                                ElseIf Val(gclsPros.InNo) <> 0 Then
                                    objTextסԺ��.Text = gclsPros.InNo
                                    Exit Sub
                                End If
                            Else
                                gclsPros.InfosChange = False
                            End If
                        End If
                        '����õ�סԺ��������Ͳ�����
                        If LoadPatiByInNo(objTextסԺ��.Text, int��ҳid) = False Then gclsPros.IsSelPati = False
                        Call AfterLoadPatiByNo
                    End If
                End With
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SpecificInfoGotFocus(ByRef intIndex As Integer)
'���ܣ�txtSpecificInfo_GotFocus
    'ʹ�ò������Ĺر��������뷨
    Call ChangeCtl
    zlCommFun.OpenIme False
    Call TxtGotFocus(gclsPros.CurrentForm.txtSpecificInfo(intIndex), True, True)
End Sub
    
Public Sub SpecificInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'���ܣ�txtSpecificInfo_KeyDown
    Dim objTextBox As TextBox
    
    Set objTextBox = gclsPros.CurrentForm.txtSpecificInfo(intIndex)
    If intKeyCode = vbKeyReturn Then
        If intIndex = SLC_סԺ�� Then
            gclsPros.IsReturn = True: Exit Sub
        ElseIf (intIndex = SLC_�ɹ����� Or intIndex = SLC_���ȴ���) Then
             Set objTextBox = gclsPros.CurrentForm.txtSpecificInfo(SLC_���ȴ���)
        ElseIf intIndex = SLC_��λ�ʱ� Or intIndex = SLC_��ͥ�ʱ� Or intIndex = SLC_�����ʱ� Then
            If ((Not IsNumeric(objTextBox.Text)) Or Len(objTextBox.Text) > 6 Or InStr(objTextBox.Text, ".") > 0) And objTextBox.Text <> "" Then
                Call SelectYouBian(objTextBox)
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    Else
        gclsPros.IsReturn = False
    End If
End Sub


Public Sub SelectYouBian(objTextBox As TextBox)
    '���ܣ��ʱ�ѡ����
    Dim strInput As String
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI

    strInput = objTextBox.Text
    If strInput <> "" Then
        If zlCommFun.IsCharChinese(strInput) Then
            strSql = strSql & " And A.���� Like [1] "
        Else
            strSql = strSql & " And A.���� Like [1] "
        End If
    Else
        Exit Sub
    End If
    strSql = "Select Rownum as ID,����,����,�ʱ�  From ���� A " & _
             "Where �ʱ� is not null " & strSql & " Order by ����"
    vPoint = GetCoordPos(objTextBox.hwnd, 0, 0)
    Set rsTmp = zlDatabase.ShowSQLSelect(objTextBox.Parent, strSql, 0, "�ʱ�", False, "", "", False, _
        False, True, vPoint.X, vPoint.Y, objTextBox.Height, False, False, False, UCase(strInput) & "%")
    If Not rsTmp Is Nothing Then
        objTextBox.Text = rsTmp!�ʱ� & ""
    End If
End Sub


Public Sub SpecificInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'���ܣ�txtSpecificInfo_KeyPress
    Dim objTextBox As TextBox
    Dim objCboTmp As ComboBox
    Dim blnCBO As Boolean
    Dim strMask As String
    Dim strTmp As String
    Dim blnEditInfo As Boolean

    On Error Resume Next
    Set objTextBox = gclsPros.CurrentForm.txtSpecificInfo(intIndex)
    If Err.Number <> 0 Then
        Set objCboTmp = gclsPros.CurrentForm.cboSpecificInfo(intIndex): blnCBO = True
        Err.Clear: On Error GoTo 0
    Else
        On Error GoTo 0
    End If
    
    Select Case intIndex
        Case SLC_סԺ��
            If gclsPros.OpenMode <> EM_�༭ And gclsPros.IsReturn Then
            '�����Ϣ�Ƿ�仯
                If Not CheckMedPageChange Then
                    gclsPros.InfosChange = False
                End If
                If Val(objTextBox.Text) <> Val(gclsPros.InNo) And Trim(objTextBox.Text) <> "" And IsHavePageNos(CT_סԺ��, False, Val(objTextBox.Text)) Then
                    If gclsPros.InfosChange And Val(gclsPros.InNo) <> 0 Then
                        gclsPros.InfosChange = False
                        If MsgBox("��Ϣ�ѷ����仯���Ƿ�ȷ�ϸ���¼�벡�ˣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYesNo Then
                            objTextBox.Text = gclsPros.InNo
                            Exit Sub '������������������
                        Else
                            Call gclsPros.InitCacheRecInfo
                        End If
                    Else
                        gclsPros.InfosChange = False
                    End If
                End If
                If Val(objTextBox.Text) <> Val(gclsPros.InNo) Or objTextBox.Text = "" Then
                    If LoadPatiByInNo(objTextBox.Text) Then
                        Call AfterLoadPatiByNo
                    Else
                        gclsPros.InNo = ""
                    End If
                    gclsPros.IsReturn = False
                ElseIf Val(objTextBox.Text) = Val(gclsPros.InNo) Then
                    intKeyAscii = 0
                    Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
                End If
            End If
        Case SLC_��λ�ʱ�, SLC_�����ʱ�, SLC_��ͥ�ʱ�
            If Chr(intKeyAscii) = "." Then
                intKeyAscii = 0
            End If
    End Select

    If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
        '�������볤��
        If Not blnCBO Then
            If objTextBox.MaxLength <> 0 Then
                If zlCommFun.ActualLen(objTextBox.Text) > objTextBox.MaxLength Then
                    intKeyAscii = 0: Exit Sub
                End If
            End If
        End If

        Select Case intIndex
            Case SLC_��ͥ�绰, SLC_��λ�绰, SLC_��ϵ�˵绰
                strMask = "1234567890-()"
            Case SLC_סԺ��, SLC_���ȴ���, SLC_�ɹ�����, _
                    SLC_��������, SLC_����ʱ����Ժǰ_Сʱ, SLC_����ʱ����Ժǰ_����, SLC_����ʱ����Ժ��_����, SLC_����ʱ����Ժ��_Сʱ, _
                    SLC_����ʱ����Ժǰ_��, SLC_����ʱ����Ժ��_��, SLC_������ʹ��, SLC_��֢�໤��, SLC_��֢�໤Сʱ, SLC_Apgar, SLC_QQ, SLC_��Ժ����, SLC_Ժ�ڻ���
                strMask = "1234567890"
            Case SLC_���ϸ��, SLC_��ѪС��, SLC_��Ѫ��, SLC_��ȫѪ, SLC_��׵���, SLC_�������, SLC_ICU, _
                    SLC_CCU, SLC_һ������, SLC_��������, SLC_��������, SLC_�ػ�, SLC_Լ����ʱ��, SLC_���, SLC_����, SLC_���ϴ�סԺʱ��
                strMask = "1234567890."
            Case SLC_��������������, SLC_��������Ժ����
                strMask = "1234567890.;"
            Case SLC_Ӥ�׶�����_DAY
                strMask = "0123456789"
        End Select

        If strMask <> "" Then
            If InStr(strMask, Chr(intKeyAscii)) = 0 Then
                intKeyAscii = 0: Exit Sub
            ElseIf intIndex = SLC_Apgar Then
                    '��֤txtApgar����ֵ��0-10֮��
                    If objTextBox.Text <> "" And objTextBox.Text <> "1" Or _
                        objTextBox.Text <> "" And objTextBox.Text = "1" And Chr(intKeyAscii) <> "0" Then
                        intKeyAscii = 0: Exit Sub
                    End If
            End If
        End If
    End If
End Sub

Public Sub SpecificInfoMouseDown(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'���ܣ�txtSpecificInfo_MouseDown
    Call TxtMouseDown(gclsPros.CurrentForm.txtSpecificInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

Public Sub SpecificInfoMouseUp(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'���ܣ�txtSpecificInfo_MouseUp
    Call TxtMouseUp(gclsPros.CurrentForm.txtSpecificInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

Public Sub SpecificInfoValidate(ByRef intIndex As Integer, ByRef blnCancel As Boolean)
'���ܣ�txtSpecificInfo_Validate
    Dim objText As TextBox
    Dim objTextDate As MaskEdBox

    Set objText = gclsPros.CurrentForm.txtSpecificInfo(intIndex)
    Select Case intIndex
        Case SLC_����
            'û�������г�������ʱ����һ������
            Set objTextDate = gclsPros.CurrentForm.mskDateInfo(DC_��������)
            If objText.Text = "" And IsDate(objTextDate.Text) Then
                objTextDate.Tag = ""
'                Call txt��������_Validate(False)
            End If
        Case SLC_���ȴ���, SLC_�ɹ�����, SLC_��������, SLC_���ϸ��, SLC_��ѪС��, SLC_��Ѫ��, SLC_��ȫѪ, SLC_�������
            If objText.Text <> "" Then
                If Not IsNumeric(objText.Text) Then
                    objText.Text = ""
                ElseIf Val(objText.Text) <= 0 And intIndex <> SLC_�ɹ����� Then
                    objText.Text = ""
                ElseIf intIndex = SLC_���ȴ��� Or intIndex = SLC_�ɹ����� Or intIndex = SLC_�������� Then
                    If IsNumeric(objText.Text) Then
                        objText.Text = Int(Val(objText.Text))
                    End If
                End If
            End If
    End Select
End Sub

'CboBaseInfo�¼�
Public Sub CboBaseInfoChange(ByRef intIndex As Integer)
'CboBaseInfo_Change�¼�
    Dim cboTmp As ComboBox
    Dim lngPos As Long, lnglen As Long

    If gclsPros.IsReturn Then Exit Sub
    Select Case intIndex
        Case BCC_���֤
            Set cboTmp = gclsPros.CurrentForm.cboBaseInfo(intIndex)
            gclsPros.IsReturn = True
            If Cbo.FindIndex(cboTmp, cboTmp.Text, True) = -1 Then
                '�����������
                If Not zlStr.CheckCharScope(cboTmp.Text, "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ*") Then
                    cboTmp.Text = ""
                Else
                    If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_����).Text)) = "�й�" Then
                        If zlCommFun.ActualLen(cboTmp.Text) > 18 Then
                            cboTmp.Text = Mid(cboTmp.Text, 1, 18)
                        End If
                    End If
                End If
            End If
            '������ҳ������ʵ
            If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_����).Text)) = "�й�" Then
                If cboTmp.Tag <> "������Change�¼�" Then
                    lngPos = InStr(cboTmp.Text, "*")
                    lnglen = Len(Mid(cboTmp.Text, 13, 2))
                    Select Case lngPos
                        Case 0
                            cboTmp.Tag = cboTmp.Text
                        Case Else 'Is <= 12
                            cboTmp.Tag = Mid(cboTmp.Text, 1, lngPos - 1)
                            cboTmp.Text = cboTmp.Tag
                            cboTmp.SelStart = Len(cboTmp.Text)
                    End Select
                End If '
            Else
                cboTmp.Tag = cboTmp.Text
            End If
            gclsPros.IsReturn = False
    End Select
    Call CheckValueChange
End Sub


Public Sub CboSpecificInfoClick(ByRef intIndex As Integer)
'cboSpecificInfo_Click�¼�
    Dim objPic As Object
    Dim objFra As Object
    Dim lngNum As Long
    
    With gclsPros.CurrentForm
        Select Case intIndex
            Case SLC_��������
                Call SetCtrlLocked(.txtSpecificInfo(intIndex), .cboSpecificInfo(intIndex).Text = "����", True)
                If .cboSpecificInfo(intIndex).Text <> "����" Then
                    If .Visible Then zlControl.ControlSetFocus (.txtSpecificInfo(intIndex))
                End If
            Case SLC_Ӥ�׶�����
                If gclsPros.LoadFinish Then
                    Set objFra = .cboSpecificInfo(intIndex).Container
                    If .cboSpecificInfo(intIndex).Text = "��" Then
                        .txtSpecificInfo(SLC_Ӥ�׶�����_DAY).Visible = True
                        .lblSpecificInfo(SLC_Ӥ�׶�����_DAY).Visible = True
                        DrawLineCTL .txtSpecificInfo(SLC_Ӥ�׶�����_DAY), 1
                        lngNum = .txtSpecificInfo(SLC_Ӥ�׶�����).Left + .txtSpecificInfo(SLC_Ӥ�׶�����).Width + 120
                        .txtSpecificInfo(SLC_Ӥ�׶�����_DAY).Left = lngNum
                        .lblSpecificInfo(SLC_Ӥ�׶�����_DAY).Left = lngNum
                        DrawLineCTL .txtSpecificInfo(SLC_Ӥ�׶�����_DAY)
                        DrawLineCTL objFra, 1
                        objFra.Left = lngNum + .txtSpecificInfo(SLC_Ӥ�׶�����_DAY).Width + 120
                        DrawLineCTL objFra
                    Else
                        .txtSpecificInfo(SLC_Ӥ�׶�����_DAY).Text = ""
                        .txtSpecificInfo(SLC_Ӥ�׶�����_DAY).Visible = False
                        .lblSpecificInfo(SLC_Ӥ�׶�����_DAY).Visible = False
                        
                        If .cboSpecificInfo(intIndex).Tag <> "����" Then
                            DrawLineCTL objFra, 1
                            objFra.Left = .txtSpecificInfo(SLC_Ӥ�׶�����).Left + .txtSpecificInfo(SLC_Ӥ�׶�����).Width + 120
                            DrawLineCTL objFra
                        Else
                            DrawLineCTL objFra, 1  '�������
                            DrawLineCTL .txtSpecificInfo(SLC_Ӥ�׶�����) '�ػ��������ⵥλ�������ʱ��Ӥ�׶����������Ҳ���������
                        End If
                    End If
                End If
        End Select
        Call CheckValueChange(.txtSpecificInfo(intIndex))
    End With
End Sub

Public Sub CboSpecificInfoGotFocus(ByRef intIndex As Integer)
'CboSpecificInfo_GotFocus�¼�
    Call ChangeCtl
    With gclsPros.CurrentForm
        '�޶��������Ŀһ�㲻�����뺺��
        zlCommFun.OpenIme False
        Call ShowInfectInfo(False)
    End With
End Sub

Public Sub CboSpecificInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'CboSpecificInfo_KeyDown�¼�
    With gclsPros.CurrentForm
        If intKeyCode = vbKeyDelete Then
            If .cboSpecificInfo(intIndex).Style = 2 And .cboSpecificInfo(intIndex).ListIndex <> -1 Then
                .cboSpecificInfo(intIndex).ListIndex = -1
            End If
        End If
    End With
End Sub

Public Sub cboSpecificInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'cboSpecificInfo_KeyPress�¼�
    Dim lngIdx As Long
    Dim cboTmp As ComboBox
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        zlCommFun.PressKey vbKeyTab: mblnReturn = True
    Else
        Set cboTmp = gclsPros.CurrentForm.cboSpecificInfo(intIndex)
        If intIndex = SLC_���� And gclsPros.FuncType = f������ҳ Or intIndex = SLC_Ӥ�׶����� Then
          If cboTmp.ListCount + vbKey1 >= intKeyAscii And intKeyAscii >= vbKey1 Then
                If intKeyAscii - vbKey1 <= cboTmp.ListCount Then
                    cboTmp.ListIndex = intKeyAscii - vbKey1
                End If
            End If
        Else
            lngIdx = zlControl.CboMatchIndex(cboTmp.hwnd, intKeyAscii)
            If lngIdx = -1 And cboTmp.ListCount > 0 Then lngIdx = 0
            cboTmp.ListIndex = lngIdx
        End If
    End If
End Sub

Public Sub cboSpecificInfoLostFocus(ByRef intIndex As Integer)
'cboSpecificInfo_LostFocus�¼�
    Dim lngIdx As Long
    Dim cboTmp As ComboBox, txtTmp As TextBox

    Set cboTmp = gclsPros.CurrentForm.cboSpecificInfo(intIndex)
    Set txtTmp = gclsPros.CurrentForm.txtSpecificInfo(intIndex)
    If intIndex = SLC_���� And gclsPros.FuncType = f������ҳ Or intIndex = SLC_Ӥ�׶����� Then
        If Not ValidateAge(txtTmp, cboTmp, IIf(intIndex = SLC_Ӥ�׶�����, 1, 0)) Then Exit Sub
    End If
End Sub

Public Sub txtDateInfoGotFocus(Index As Integer)
    Call ChangeCtl
    Call TxtGotFocus(gclsPros.CurrentForm.txtDateInfo(Index), True, True)
End Sub

'cmdAutoLoad�¼�
Public Sub CmdAutoLoadClick(ByRef intIndex As Integer)
'cmdAutoLoad_Click�¼�
    Dim strSql As String, rsTmp As Recordset
    Dim DateSs As Date          '�ò������������ʱ��
    Dim rsTime As ADODB.Recordset
    Dim vsTmp As VSFlexGrid
    Dim i As Long, j As Long, LngRow As Long
    Dim blnClear As Boolean
    Dim strPrivs As String
    Dim strUseStage As String

    On Error GoTo errH
    Select Case intIndex
        Case ALC_������ '����ҩ�Զ���ȡ
            strSql = "Select Min(NVL(to_date(c.�걾��λ,'yyyy-mm-dd hh24:mi:ss'),c.��ʼִ��ʱ��)) as ʹ��ʱ��" & vbNewLine & _
                    " From ������ĿĿ¼ A, ����ҽ����¼ C" & vbNewLine & _
                    " Where  a.Id = c.������Ŀid and a.���='F' And c.����id = [1] And c.��ҳid = [2] And c.ҽ��״̬=8"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, gclsPros.��ҳID)
            If rsTmp.RecordCount > 0 Then DateSs = CDate(Format(NVL(rsTmp!ʹ��ʱ��, 0), "yyyy-MM-dd"))

            strSql = "Select distinct ID, ҽ��id, �ϼ�id, ����, ����, ��λ, ִ��ʱ�䷽��, Ƶ�ʼ��, �����λ, Ƶ�ʴ���, �ϴ�ִ��ʱ��, ��ʼִ��ʱ��, ����ʱ��," & vbNewLine & _
                    "       Sum(Ddd��) Over(Partition By ID,��ҩĿ��) As Ddd��, Count(1) Over(Partition By ���id) As ������ҩ,ҩ��ID,decode(��ҩĿ��,1,'Ԥ��',2,'����',' ') as ��ҩĿ��" & vbNewLine & _
                    "From   (Select Distinct ID, ҽ��id, �ϼ�id, ����, ����, ��λ, ִ��ʱ�䷽��, Ƶ�ʼ��, �����λ, Ƶ�ʴ���, �ϴ�ִ��ʱ��, ��ʼִ��ʱ��, ����ʱ��," & vbNewLine & _
                    "                Sum(����) Over(Partition By ID, ҽ��id, ���id,��ҩĿ��) * ����ϵ�� / Decode(Dddֵ, 0, Null, Dddֵ) As Ddd��, ���id,ҩ��ID,��ҩĿ��" & vbNewLine & _
                    "         From   (Select z.Id, a.Id As ҽ��id, z.����id As �ϼ�id, z.����, z.����, z.���㵥λ As ��λ, a.ִ��ʱ�䷽��, a.Ƶ�ʼ��, a.�����λ, a.Ƶ�ʴ���," & vbNewLine & _
                    "                         a.�ϴ�ִ��ʱ��, a.��ʼִ��ʱ��, Nvl(a.�ϴ�ִ��ʱ��, Nvl(a.ִ����ֹʱ��, a.��ʼִ��ʱ��)) As ����ʱ��, a.���id, f.����, h.����ϵ��," & vbNewLine & _
                    "                         Nvl((Select e.Dddֵ From �����÷����� E Where e.��Ŀid = a.������Ŀid And e.�÷�id = r.������Ŀid), h.Dddֵ) As Dddֵ,A.������ĿID as ҩ��ID,A.��ҩĿ��" & vbNewLine & _
                    "                  From   ����ҽ����¼ A, ����ҽ����¼ R, סԺ���ü�¼ F, ҩƷ��� H, ҩƷ���� B, ������ĿĿ¼ Z" & vbNewLine & _
                    "                  Where  a.������Ŀid = b.ҩ��id And a.������� In ('5', '6') And" & vbNewLine & _
                    "                         (a.ҽ����Ч = 0 And a.�ϴ�ִ��ʱ�� Is Not Null Or a.ҽ����Ч = 1 And a.ҽ��״̬ = 8) And Nvl(b.������, 0) <> 0 And" & vbNewLine & _
                    "                         a.���id = r.Id And a.Id = f.ҽ����� And f.��¼״̬ <> 0 And f.�շ�ϸĿid = h.ҩƷid And b.ҩ��id = z.Id And" & vbNewLine & _
                    "                         f.��¼���� <> 12 And a.����id = [1] And a.��ҳid = [2]))" & vbNewLine & _
                    "Order  By Ddd�� Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, gclsPros.��ҳID)

            If rsTmp.RecordCount = 0 Then
                MsgBox "û���ҵ��ò��˵Ŀ���ҩ��ʹ�ü�¼��", vbInformation, gstrSysName
                Exit Sub
            End If
            Set vsTmp = gclsPros.CurrentForm.vsKSS
            With vsTmp
                Do While Not rsTmp.EOF
                    LngRow = 0
                    strUseStage = GetKSSUseStage(CDate(Format(rsTmp!��ʼִ��ʱ�� & "", "yyyy-MM-dd")), CDate(Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd")), DateSs)
                    For i = .FixedRows To .Rows - 1
                        '��λ��
                        LngRow = 0
                        If Val(rsTmp!ID & "") = 0 Then
                            Exit For '���ص�ʱ�򲻻���أ�����˳�������һ��ѭ��
                        ElseIf .TextMatrix(i, KI_����ҩ����) & "" = "" Then
                            LngRow = i: Exit For
                        ElseIf Val(rsTmp!ID & "") = Val(.RowData(i) & "") And .TextMatrix(i, KI_��ҩĿ��) = rsTmp!��ҩĿ�� & "" And (.TextMatrix(i, KI_ʹ�ý׶�) = "" Or .TextMatrix(i, KI_ʹ�ý׶�) = strUseStage) Then
                            LngRow = -1 * i: Exit For
                        ElseIf i = .Rows - 1 Then
                            .AddItem ""
                            LngRow = .Rows - 1
                            Exit For
                        End If
                    Next

                    If LngRow > 0 Then
                        .RowData(LngRow) = Val(rsTmp!ҩ��id & "")
                        .TextMatrix(LngRow, KI_����ҩ����) = rsTmp!���� & ""
                        .Cell(flexcpData, LngRow, KI_����ҩ����) = .TextMatrix(LngRow, KI_����ҩ����)
                        .TextMatrix(LngRow, KI_��ҩĿ��) = rsTmp!��ҩĿ�� & ""
                        .TextMatrix(LngRow, KI_DDD��) = FormatEx(Val(rsTmp!DDD�� & ""), 2)
                        .TextMatrix(LngRow, KI_������ҩ) = decode(Val(rsTmp!������ҩ & ""), 1, "����", 2, "����", 3, "����", 4, "����", ">����")
                        .TextMatrix(LngRow, KI_ʹ�ý׶�) = strUseStage
                    Else '��ͬ��¼
                        LngRow = Abs(LngRow)
                        If .TextMatrix(LngRow, KI_DDD��) = "" Then .TextMatrix(LngRow, KI_DDD��) = FormatEx(Val(rsTmp!DDD�� & ""), 2)
                        If decode(.TextMatrix(i, KI_������ҩ), "����", 1, "����", 2, "����", 3, "����", 4, ">����", 999, 0) < Val(rsTmp!������ҩ & "") Then
                            .TextMatrix(i, KI_������ҩ) = decode(Val(rsTmp!������ҩ & ""), 1, "����", 2, "����", 3, "����", 4, "����", ">����")
                        End If
                    End If
                    If LngRow <> 0 Then '��ȡʹ������
                        .TextMatrix(LngRow, KI_ʹ������) = GetKSSUseDay(Val(rsTmp!ҽ��ID), Val(.RowData(LngRow)), NVL(rsTmp!ִ��ʱ�䷽��) & "", CDate(rsTmp!��ʼִ��ʱ��), CDate(rsTmp!����ʱ��), _
                                NVL(rsTmp!Ƶ�ʴ���, 0), NVL(rsTmp!Ƶ�ʼ��, 0), NVL(rsTmp!�����λ), NVL(rsTmp!��ҩĿ��), rsTime) & ""
                    End If
                    rsTmp.MoveNext
                Loop
                Call ChangeVSFHeight(vsTmp, True)
            End With
        Case ALC_���� '�����Զ���ȡ
            strPrivs = GetInsidePrivs(p����ӿ�, , 2400)
            If InStr(strPrivs, "�ڲ��ӿ�") > 0 Then
                gclsPros.CurrentForm.lblAutoInfo.Visible = True
                gclsPros.CurrentForm.lblAutoInfo = "������Դ���������ϵͳ(Ĭ��)"
                Set rsTmp = AutoGetOPSInfo(True, gclsPros.����ID, gclsPros.��ҳID)
            Else
                If gblnHaveOPS Then
                    gclsPros.CurrentForm.lblAutoInfo.Visible = True
                    gclsPros.CurrentForm.lblAutoInfo = "������Դ������ҽ�����(û�С��������ϵͳ-����ӿڹ���-�ڲ��ӿڡ�Ȩ��)"
                Else
                    gclsPros.CurrentForm.lblAutoInfo.Visible = True
                    gclsPros.CurrentForm.lblAutoInfo = "������Դ������ҽ�����(δ��װ�������ϵͳ)"
                End If
                Set rsTmp = AutoGetOPSInfo(False, gclsPros.����ID, gclsPros.��ҳID)
            End If

            If Not rsTmp.EOF Then
                Set vsTmp = gclsPros.CurrentForm.vsOPS
                '����������������Ƿ���������Ϣ
                For i = vsTmp.FixedRows To vsTmp.Rows - 1
                    If vsTmp.TextMatrix(i, PI_��������) <> "" Or vsTmp.TextMatrix(i, PI_��������) <> "" Then
                        If MsgBox("�Ƿ����ԭ�е�������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                            blnClear = True
                        End If
                        Exit For
                    End If
                Next
                rsTmp.MoveFirst
                With vsTmp
                    If blnClear Then .Rows = .FixedRows
                    LngRow = IIf(.TextMatrix(.Rows - 1, PI_��������) <> "", .Rows, .Rows - 1)
                    .Rows = .Rows + rsTmp.RecordCount + IIf(.TextMatrix(.Rows - 1, PI_��������) <> "", 1, 0)
                    Call ChangeVSFHeight(vsTmp, True)
                    For i = LngRow To LngRow + rsTmp.RecordCount - 1
                        .TextMatrix(i, PI_��������) = Format(NVL(rsTmp!������ʼʱ��, rsTmp!��������) & "", "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, PI_��������) = Format(NVL(rsTmp!��������ʱ��, rsTmp!��������) & "", "yyyy-MM-dd HH:mm")
                        .TextMatrix(i, PI_��������) = rsTmp!�������� & ""
                        .TextMatrix(i, PI_��������) = rsTmp!�������� & ""
                        .TextMatrix(i, PI_����ҽʦ) = rsTmp!����ҽʦ & ""
                        .TextMatrix(i, PI_������ʿ) = rsTmp!������ʿ & ""
                        .TextMatrix(i, PI_����1) = rsTmp!��һ���� & ""
                        .TextMatrix(i, PI_����2) = rsTmp!�ڶ����� & ""
                        .TextMatrix(i, PI_����ʽ) = rsTmp!����ʽ & ""
                        .TextMatrix(i, PI_����ҽʦ) = rsTmp!����ҽʦ & ""
                        If rsTmp!�п� & rsTmp!���� & "" <> "" Then
                            .TextMatrix(i, PI_�п�����) = rsTmp!�п� & "/" & rsTmp!����
                        End If
                        .TextMatrix(i, PI_��������ID) = Val(rsTmp!��������ID & "")
                        .TextMatrix(i, PI_������ĿID) = Val(rsTmp!������Ŀid & "")
                        .TextMatrix(i, PI_����ID) = Val(rsTmp!ID & "")
                        .TextMatrix(i, PI_��������) = rsTmp!�������� & ""
                        .TextMatrix(i, PI_�������) = rsTmp!������� & ""
                        .TextMatrix(i, PI_ASA�ּ�) = rsTmp!asa�ּ� & ""
                        .TextMatrix(i, PI_NNIS�ּ�) = rsTmp!NNIS�ּ� & ""
                        .TextMatrix(i, PI_��������) = rsTmp!�������� & ""
                        .TextMatrix(i, PI_�ٴ�����) = IIf(Val(rsTmp!�ٴ����� & "") = 1, -1, 0)
                        .TextMatrix(i, PI_����ʼʱ��) = Format(rsTmp!����ʼʱ�� & "", "yyyy-MM-dd HH:mm")
                        .Cell(flexcpData, i, PI_��������) = rsTmp!����ԭ�� & ""
                        '��¼���ڱ༭�ָ�
                        For j = 0 To .Cols - 1
                            If j = PI_�������� And .TextMatrix(i, PI_��������) <> "" Then
                                If .Cell(flexcpData, i, j) = "" Then
                                    .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                                End If
                            Else
                                .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                            End If
                        Next
                        .Cell(flexcpData, i, PI_��������) = IIf(rsTmp!�������� & "" = "", 0, 1)
                        rsTmp.MoveNext
                    Next
                End With
            End If
        Case ALC_������¼
             If gclsPros.CurrentForm.chkInfo(CHK_�޹�����¼).Value = 1 Then gclsPros.CurrentForm.chkInfo(CHK_�޹�����¼).Value = 0
             strSql = " Select Distinct a.Id, a.��¼��Դ, a.����ʱ��, a.ҩ��id, a.ҩ����, a.������Ӧ, a.����Դ����, a.��¼ʱ��" & vbNewLine & _
                      " From (Select a.Id, a.��¼��Դ, a.����ʱ��, a.ҩ��id, a.ҩ����, a.������Ӧ, a.����Դ����, a.��¼ʱ��" & vbNewLine & _
                      "       From ���˹�����¼ A," & vbNewLine & _
                      "            (Select c.����id, c.��ҳid, c.ҩ��id, Max(c.��¼ʱ��) As ��¼ʱ��" & vbNewLine & _
                      "              From ���˹�����¼ C" & vbNewLine & _
                      "              Where c.��¼��Դ = 2 And c.����id = [1] And c.��ҳid = [2]" & vbNewLine & _
                      "              Group By c.����id, c.��ҳid, c.ҩ��id) B" & vbNewLine & _
                      "       Where a.��� = 1 And a.����id = [1] And a.��ҳid = [2] And" & vbNewLine & _
                      "             ((a.��¼��Դ = 2 And a.��¼ʱ�� = b.��¼ʱ�� And a.ҩ��id = b.ҩ��id) Or a.��¼��Դ in (1,3))" & vbNewLine & _
                      "       Union" & vbNewLine & _
                      "       Select a.Id, a.��¼��Դ, a.����ʱ��, a.ҩ��id, a.ҩ����, a.������Ӧ, a.����Դ����, a.��¼ʱ��" & vbNewLine & _
                      "       From ���˹�����¼ A" & vbNewLine & _
                      "       Where a.��� = 1 And a.����id = [1] And a.��ҳid = [2] And a.��¼��Դ in (1,3) ) A" & vbNewLine & _
                      " Order By Nvl(Trunc(a.����ʱ��), a.��¼ʱ��) Desc,a.��¼��Դ Desc,a.ҩ����"
             
            On Error GoTo errH
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ҳ��ȡ������Ϣ", gclsPros.����ID, gclsPros.��ҳID)
            
            With gclsPros.CurrentForm.vsAller
                If rsTmp.EOF Then
                    MsgBox "û����ȡ���κεĹ�����¼��Ϣ��", vbInformation, gstrSysName
                Else
                    rsTmp.MoveFirst

                    .Rows = .FixedRows
                    For i = 1 To rsTmp.RecordCount
                        LngRow = -1
                        If Not IsNull(rsTmp!ҩ��ID) Then
                            LngRow = .FindRow(rsTmp!ҩ��ID & "", , AI_ҩ��ID, , True)
                        ElseIf Not IsNull(rsTmp!ҩ����) Then
                            LngRow = .FindRow(rsTmp!ҩ���� & "", , AI_����ҩ��, , True)
                        End If
                        If LngRow = -1 Then
                            For j = .FixedRows To .Rows - 1
                                If .TextMatrix(j, AI_����ҩ��) = "" Then
                                    LngRow = j
                                End If
                            Next
                            
                            If LngRow = -1 Then
                                .Rows = .Rows + 1
                                LngRow = .Rows - 1
                            End If
                            
                            .TextMatrix(LngRow, AI_����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd")
                            .TextMatrix(LngRow, AI_����ҩ��) = NVL(rsTmp!ҩ����)
                            .TextMatrix(LngRow, AI_������Ӧ) = NVL(rsTmp!������Ӧ)
                            .TextMatrix(LngRow, AI_����Դ����) = NVL(rsTmp!����Դ����)
                            .TextMatrix(LngRow, AI_ҩ��ID) = rsTmp!ҩ��ID & ""
                            .TextMatrix(LngRow, AI_������Դ) = rsTmp!��¼��Դ & ""
                            '���ݱ��ݴ洢
                            .Cell(flexcpData, LngRow, AI_����ʱ��) = .TextMatrix(LngRow, AI_����ʱ��)
                            .Cell(flexcpData, LngRow, AI_����ҩ��) = .TextMatrix(LngRow, AI_����ҩ��)
                            .Cell(flexcpData, LngRow, AI_������Ӧ) = .TextMatrix(LngRow, AI_������Ӧ)
                            .Cell(flexcpData, LngRow, AI_����Դ����) = .TextMatrix(LngRow, AI_����Դ����)
                            .Cell(flexcpData, LngRow, AI_ҩ��ID) = .TextMatrix(LngRow, AI_ҩ��ID)
                            .RowData(LngRow) = Val(rsTmp!ID & "")
                        End If
                        rsTmp.MoveNext
                    Next
                    .Rows = .Rows + 1   '����һ�п���
                    .Row = .FixedRows
                    .Col = AI_����ҩ��
                    Call ChangeVSFHeight(gclsPros.CurrentForm.vsAller, True, 300, 3)
                End If
            End With
        Case ALC_�ٴ�·��
            strSql = "Select Decode(c.����, 2, c.����, '') As ����,b.״̬" & vbNewLine & _
                "From ����·������ A, �����ٴ�·�� B, ���쳣��ԭ�� C" & vbNewLine & _
                "Where a.·����¼id(+) = b.Id And b.��ǰ���� = a.����(+) And Nvl(b.��ǰ�׶�id, b.ǰһ�׶�id) = a.�׶�id(+) And b.״̬ <> 0 And a.����ԭ�� = c.����(+) And b.����id = [1] And b.��ҳid = [2]"

            On Error GoTo errH
            With gclsPros.CurrentForm
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption, gclsPros.����ID, gclsPros.��ҳID)
                If rsTmp.RecordCount > 0 Then
                    .chkInfo(CHK_����·��).Value = 1
                    If Val(rsTmp!״̬ & "") = 3 Then
                        .chkInfo(CHK_���·��).Value = 0
                        .txtInfo(GC_�˳�ԭ��).Text = rsTmp!���� & ""
                    ElseIf Val(rsTmp!״̬ & "") = 2 Then
                        .chkInfo(CHK_���·��).Value = 1
                    End If
                Else
                    .chkInfo(CHK_����·��).Value = 0
                End If
                '��ȡ�������
                strSql = "Select Count(1) Over(Partition By b.����id, b.��ҳid) As ������, c.���� As ����ԭ��" & vbNewLine & _
                        "From ����·������ A, �����ٴ�·�� B, ���쳣��ԭ�� C" & vbNewLine & _
                        "Where a.·����¼id = b.Id And c.����(+) = a.����ԭ�� And a.������� = -1 And b.����id = [1] And b.��ҳid = [2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, .Caption, gclsPros.����ID, gclsPros.��ҳID)
                If rsTmp.RecordCount > 0 Then
                    .chkInfo(CHK_����).Value = 1
                    If Val(rsTmp!������ & "") = 1 And Not gclsPros.PathVCauses Then
                        .txtInfo(GC_����ԭ��).Text = rsTmp!����ԭ�� & ""
                    End If
                Else
                    .chkInfo(CHK_����).Value = 0
                End If
            End With
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'cboBaseInfo�¼�
Public Sub CboBaseInfoClick(ByRef intIndex As Integer)
'cboBaseInfo_Click�¼�
    Dim rsTmp As ADODB.Recordset
    Dim objTextBox As TextBox
    Dim strTmp As String
    Dim blnLocked As Boolean

    With gclsPros.CurrentForm
        Select Case intIndex
            Case BCC_��Ժ��ʽ
                strTmp = zlStr.NeedName(.cboBaseInfo(intIndex).Text)
                blnLocked = Not (strTmp Like "*תԺ*" Or strTmp Like "*ת����*")
                Call SetCtrlLocked(.txtInfo(GC_ת��ҽ�ƻ���), blnLocked, True)
                Call SetCtrlLocked(.cmdInfo(GC_ת��ҽ�ƻ���), blnLocked, True)
                If gblnSet Then Exit Sub
                Call ChangeOutInfo(strTmp, True) '������ϵĳ�Ժ���
            Case BCC_��ϵ
                strTmp = zlStr.NeedName(.cboBaseInfo(intIndex).Text)
                If strTmp Like "*����*" Then
'                    .txtInfo(GC_������ϵ).Visible = True
                    .picRelation.Visible = True
                Else
'                    .txtInfo(GC_������ϵ).Visible = False
                    .picRelation.Visible = False
                    .txtInfo(GC_������ϵ).Text = ""
                End If
            Case BCC_��Ժ;��
                strTmp = zlStr.NeedName(.cboBaseInfo(intIndex).Text)
                blnLocked = Not (strTmp Like "*ת��*" And Not strTmp Like "*��ת��*")
                Call SetCtrlLocked(.txtInfo(GC_��Ժת��), blnLocked, True)
                Call SetCtrlLocked(.cmdInfo(GC_��Ժת��), blnLocked, True)
            Case BCC_��Һ��Ӧ
                On Error Resume Next
                Set objTextBox = .txtInfo(GC_����ҩ��) '���Ĵ�����
                strTmp = objTextBox.Text
                If Err.Number = 0 Then
                    blnLocked = zlStr.NeedName(.cboBaseInfo(intIndex).Text) <> "��"
                    Call SetCtrlLocked(objTextBox, blnLocked, True)
                    Call SetCtrlLocked(.txtInfo(GC_�ٴ�����), blnLocked, True)
                Else
                    Err.Clear
                End If
                On Error GoTo 0
            Case BCC_��������ʬ��
                If .cboBaseInfo(BCC_��������ʬ��).ListIndex = 1 Then
                    .cboBaseInfo(BCC_�ٴ���ʬ��).Clear
                    .cboBaseInfo(BCC_�ٴ���ʬ��).AddItem "0-δ��"
                    .cboBaseInfo(BCC_�ٴ���ʬ��).AddItem "1-����"
                    .cboBaseInfo(BCC_�ٴ���ʬ��).AddItem "2-������"
                    .cboBaseInfo(BCC_�ٴ���ʬ��).AddItem "3-���϶�"
                Else
                    .cboBaseInfo(BCC_�ٴ���ʬ��).Clear
                    .cboBaseInfo(BCC_�ٴ���ʬ��).AddItem "-"
                End If
                Call SetDiagMatchInfo(BCC_�ٴ���ʬ��)
            Case BCC_����ԭ��
                .txtInfo(GC_����ԭ��).Text = zlStr.NeedName(.cboBaseInfo(intIndex).Text)
        End Select
        Call CheckValueChange(.cboBaseInfo(intIndex))
    End With
End Sub

Public Sub cboBaseInfoDropDown(ByRef intIndex As Integer)
'���ܣ�cboBaseInfo_DropDown�¼���װ
    Dim strTmp As String
    Dim cboTmp As ComboBox
    Dim intIdx As Integer

    Set cboTmp = gclsPros.CurrentForm.cboBaseInfo(intIndex)
    strTmp = cboTmp.Text
    If (intIndex = BCC_���� Or intIndex = BCC_ְҵ Or intIndex = BCC_����) And cboTmp.ListCount = 0 Then
        Call SetCboFromRec(BCC_����, 0)
    End If
    If strTmp <> "" Then
        intIdx = Cbo.FindIndex(cboTmp, strTmp)
        If intIdx <> -1 Then
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
        End If
    End If
    cboTmp.Tag = cboTmp.ListIndex
    'ʹ�ò������Ĺر��������뷨
    zlCommFun.OpenIme False
    Call ShowInfectInfo(intIndex = BCC_��Ⱦ��������ϵ, cboTmp)
    If cboTmp.Style = 0 Then
        Call zlControl.TxtSelAll(cboTmp)
    End If
End Sub

Public Sub CboBaseInfoGotFocus(ByRef intIndex As Integer)
'CboBaseInfo_GotFocus�¼�
    Dim cboTmp As ComboBox
    Dim intIdx As Integer
    Dim strTmp As String

    Call ChangeCtl
    With gclsPros.CurrentForm
        Set cboTmp = .cboBaseInfo(intIndex)
        If (intIndex = BCC_���� Or intIndex = BCC_ְҵ Or intIndex = BCC_����) And cboTmp.ListCount = 0 Then
            Call SetCboFromRec(BCC_����, 0)
        End If
        strTmp = cboTmp.Text
        If strTmp <> "" Then
            intIdx = Cbo.FindIndex(cboTmp, strTmp)
            If intIdx <> -1 Then
                Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
            End If
        End If
         If intIndex <> BCC_���֤ Then cboTmp.Tag = cboTmp.ListIndex
        'ʹ�ò������Ĺر��������뷨
        zlCommFun.OpenIme False
        Call ShowInfectInfo(intIndex = BCC_��Ⱦ��������ϵ, cboTmp)
    End With
End Sub

Public Sub CboBaseInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'CboBaseInfo_KeyDown�¼�
    With gclsPros.CurrentForm
        If intKeyCode = vbKeyDelete Then
            If .cboBaseInfo(intIndex).Style = 0 Then
                .cboBaseInfo(intIndex).ListIndex = -1
                .cboBaseInfo(intIndex).Text = ""
            Else
                If .cboBaseInfo(intIndex).ListIndex <> -1 Then
                    .cboBaseInfo(intIndex).ListIndex = -1
                End If
            End If
        ElseIf intKeyCode = vbKeyEscape And intIndex = BCC_��Ⱦ��������ϵ Then
            Call ShowInfectInfo(False)
        End If
    End With
End Sub

Public Sub CboBaseInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'CboBaseInfo_KeyPress�¼�
    Dim lngIdx As Long, cboTmp As ComboBox
    Dim strInput As String
    Dim strFilter As String
    Dim rsInput As ADODB.Recordset

    Set cboTmp = gclsPros.CurrentForm.cboBaseInfo(intIndex)

    If intKeyAscii = vbKeyReturn And (intIndex = BCC_���� Or intIndex = BCC_ְҵ Or intIndex = BCC_����) And cboTmp.Style = 0 Then
        strInput = Trim(cboTmp.Text)
        If strInput = "" Then zlCommFun.PressKey vbKeyTab: mblnReturn = True: Exit Sub
        '��ͬ����Ŀ�򲻽��д���
        If cboTmp.ListIndex <> -1 Then
            If zlStr.NeedName(strInput) = zlStr.NeedName(cboTmp.List(cboTmp.ListIndex)) Then
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
                Exit Sub
            End If
        End If
        strInput = UCase(strInput)
        'ADO��ͨ�����*��%,ֻ������ͷƥ����βƥ�䣬����˫��ƥ�䣬�������ַ����м�ƥ��
        If zlCommFun.IsCharChinese(strInput) Then
            strFilter = "���� Like '*" & strInput & "*'"
        Else
            strFilter = "���� like '*" & strInput & "*' or ���� like '*" & strInput & "*'"
        End If
        Set rsInput = Rec.FilterNew(GetBaseCode(intIndex), strFilter)
        If rsInput.RecordCount = 0 Then Exit Sub
        '��������б�չ����ر������б�
        If SendMessage(cboTmp.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then
            SendMessageLong cboTmp.hwnd, CB_SHOWDROPDOWN, False, 0
        End If
        If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, cboTmp, rsInput, True, , , rsInput) Then
            lngIdx = Cbo.FindIndex(cboTmp, rsInput!ID)
        Else
            lngIdx = Val(cboTmp.Tag)
        End If
        Call zlControl.CboSetIndex(cboTmp.hwnd, lngIdx)
        cboTmp.Tag = cboTmp.ListIndex
    ElseIf intIndex = BCC_���֤ Then
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            zlCommFun.PressKey vbKeyTab: mblnReturn = True
        Else
            If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_����).Text)) = "�й�" Then
                If zlCommFun.ActualLen(cboTmp.Text) >= 18 And intKeyAscii <> vbKeyBack Then
                    intKeyAscii = 0 '���ֻ������18������
                Else
                    If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
                        intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                            intKeyAscii = 0
                        ElseIf zlCommFun.IsCharChinese(cboTmp.Text) Then
                            cboTmp.Text = "": cboTmp.Tag = ""
                        End If
                        If gclsPros.FuncType = fҽ����ҳ And intKeyAscii <> 0 Then
                            '������ҳ������ʵ
                            Select Case zlCommFun.ActualLen(cboTmp.Text)
                                Case 12
                                    cboTmp.Tag = cboTmp.Text & Chr(intKeyAscii)
                                Case 13
                                    cboTmp.Tag = cboTmp.Tag & Chr(intKeyAscii)
                            End Select
                        End If
                    End If
                End If
            Else
                If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                        intKeyAscii = 0
                    ElseIf zlCommFun.IsCharChinese(cboTmp.Text) Then
                        cboTmp.Text = "": cboTmp.Tag = ""
                    End If
                    If gclsPros.FuncType = fҽ����ҳ And intKeyAscii <> 0 Then
                        cboTmp.Tag = cboTmp.Text & Chr(intKeyAscii)
                    End If
                End If
            End If
        End If
    ElseIf intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
End Sub

Public Sub CboBaseInfoKeyUp(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'CboBaseInfo_KeyUp�¼�
    If intKeyCode = vbKeyDelete Then
        If intIndex = BCC_��ǰ������ Then
            gclsPros.CurrentForm.cboBaseInfo(intIndex).ListIndex = -1
        End If
    End If
End Sub

Public Sub cboBaseInfoValidate(ByRef intIndex As Integer, ByRef blnCancel As Boolean)
'CboBaseInfo_Validate�¼�
    '���ڽ������޸Ŀ�������ƥ�䣬��Ҫ��֤�Ƿ�ѡ�������
    Dim strInput As String, strFilter As String
    Dim rsTmp As ADODB.Recordset, rsInput As ADODB.Recordset
    Dim cboTmp As ComboBox
    Dim intIdx As Integer
    Dim strTmp As String

    Set cboTmp = gclsPros.CurrentForm.cboBaseInfo(intIndex)
    If (intIndex = BCC_���� Or intIndex = BCC_ְҵ Or intIndex = BCC_����) And cboTmp.ListCount = 0 Then
        Call SetCboFromRec(BCC_����, 0)
    End If
    strTmp = cboTmp.Text
    If strTmp <> "" Then
        intIdx = Cbo.FindIndex(cboTmp, strTmp)
        If intIdx <> -1 Then
            Call zlControl.CboSetIndex(cboTmp.hwnd, intIdx)
        End If
    End If
    If intIndex <> BCC_���֤ Then cboTmp.Tag = cboTmp.ListIndex

    If intIndex = BCC_���� Or intIndex = BCC_ְҵ Or intIndex = BCC_���� Then
        If cboTmp.ListIndex <> -1 Then Exit Sub '��ѡ��
        If cboTmp.Text = "" Then
            MsgBox "������" & decode(intIndex, BCC_����, "����", BCC_ְҵ, "ְҵ", BCC_����, "����") & "��", vbInformation, gstrSysName
            blnCancel = True: Exit Sub '������
        End If
        strInput = UCase(zlStr.NeedName(cboTmp.Text))
        'ADO��ͨ�����*��%,ֻ������ͷƥ����βƥ�䣬����˫��ƥ�䣬�������ַ����м�ƥ��
        If zlCommFun.IsCharChinese(strInput) Then
            strFilter = "���� Like '*" & strInput & "*'"
        Else
            strFilter = "���� like '*" & strInput & "*' or ���� like '*" & strInput & "*'"
        End If

        blnCancel = True: cboTmp.Text = ""
        Set rsInput = Rec.FilterNew(GetBaseCode(intIndex), strFilter, "ID,����,����,����,ȱʡ")
        If rsInput.RecordCount <> 0 Then
            If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, cboTmp, rsInput, True, , , rsTmp) Then
                intIdx = Cbo.FindIndex(cboTmp, rsTmp!ID)
                If intIdx <> -1 Then
                    cboTmp.ListIndex = intIdx: blnCancel = False
                End If
            End If
        End If
        If blnCancel Then
            MsgBox "������" & decode(intIndex, BCC_����, "����", BCC_ְҵ, "ְҵ", BCC_����, "����") & "��", vbInformation, gstrSysName
        End If
    End If
End Sub

'monInfo�¼���װ
Public Sub monInfoDateClick(ByVal datDateClicked As Date)
'���ܣ�monInfo_DateClick
    Dim strDate As String, strFMT As String
    Dim objMSK As MaskEdBox
    Dim datCurrent As Date

    Set objMSK = gclsPros.CurrentForm.mskDateInfo(gclsPros.DateIndex)
    '��ȡʱ��������
    If objMSK.MaxLength >= Len("####-##-## ##:##") Then
        'yyyy-MM-dd HH:mm:ss ��ʽʱ��
        If objMSK.MaxLength > Len("####-##-## ##:##") Then
            strFMT = "HH:mm:ss"
        Else
            'yyyy-MM-dd HH:mm ��ʽʱ��
            strFMT = "HH:mm"
        End If
        'ԭʱ����ʱ�����ͣ���ȡ��ʱ���ʱ�������ݣ�����ȡ��ǰʱ���ʱ����
        If IsDate(objMSK.Text) Then
            strDate = " " & Format(objMSK.Text, strFMT)
        Else
            strDate = " " & Format(zlDatabase.Currentdate, strFMT)
        End If
    End If
    '��ȡʱ��
    strDate = Format(datDateClicked, "yyyy-MM-dd") & strDate
    objMSK.Text = strDate
    Select Case gclsPros.DateIndex
        Case DC_ȷ������
            If Not CheckDateRange(strDate, True) Then
                MsgBox "�������ʱ������ڲ��˵�סԺ�ڼ䡣", vbInformation, gstrSysName
                Exit Sub
            End If
        Case DC_��Ժʱ��, DC_��Ժʱ��
            If gclsPros.InTime = "" And gclsPros.DateIndex = DC_��Ժʱ�� Then
                If IsDate(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��).Text) Then
                    gclsPros.InTime = gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��).Text
                End If
            ElseIf gclsPros.DateIndex = DC_��Ժʱ�� And gclsPros.OutTime = "" Then
                If IsDate(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��).Text) Then
                    gclsPros.InTime = gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��).Text
                End If
            End If
            If objMSK.Text <> Replace(objMSK.Mask, "#", "_") Then
                If gclsPros.DateIndex = DC_��Ժʱ�� And gclsPros.InTime <> "" Then
                    If CDate(gclsPros.InTime) > CDate(objMSK.Text) Then
                        MsgBox "������ĳ�Ժʱ��С����Ժʱ�䣬���������롣", vbInformation, gstrSysName
                        Exit Sub
                    ElseIf CDate(objMSK.Text) > zlDatabase.Currentdate Then
                        MsgBox "������ĳ�Ժʱ����ڵ�ǰʱ�䣬���������롣", vbInformation, gstrSysName
                        Exit Sub
                    Else
                        gclsPros.OutTime = objMSK.Text
                    End If
                ElseIf gclsPros.DateIndex = DC_��Ժʱ�� And gclsPros.OutTime <> "" Then
                    If CDate(gclsPros.OutTime) < CDate(objMSK.Text) Then
                        MsgBox "���������Ժʱ����ڳ�Ժʱ�䣬���������롣", vbInformation, gstrSysName
                        Exit Sub
                    ElseIf CDate(objMSK.Text) > zlDatabase.Currentdate Then
                        MsgBox "���������Ժʱ����ڵ�ǰʱ�䣬���������롣", vbInformation, gstrSysName
                        Exit Sub
                    Else
                        gclsPros.InTime = objMSK.Text
                    End If
                End If
            End If
    End Select
    gclsPros.CurrentForm.txtDateInfo(objMSK.Index).Text = objMSK.Text
    gclsPros.CurrentForm.monInfo.Visible = False
    zlControl.ControlSetFocus objMSK
End Sub

Public Sub monInfoKeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'���ܣ�monInfo_KeyDown
    If intKeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
End Sub

Public Sub monInfoKeyPress(ByRef intKeyAscii As Integer)
'���ܣ�monInfo_KeyPress
    If intKeyAscii = 13 Then
        intKeyAscii = 0
        Call monInfoDateClick(gclsPros.CurrentForm.monInfo.Value)
    End If
End Sub

Public Sub monInfoValidate(ByRef blnCancel As Boolean)
'���ܣ�monInfo_Validate
    gclsPros.CurrentForm.monInfo.Visible = False
End Sub
'lstInfectParts�¼���lstAdvEvent�¼���lstInfection�¼�
Public Sub LstGotFocus(ByRef lstInput As ListBox)
'lstInfectParts_GotFocus�¼���lstAdvEvent_GotFocus�¼���lstInfection_GotFocus�¼�
    Call ChangeCtl
    lstInput.ListIndex = 0
End Sub

Public Sub lstLostFocus(ByRef lstInput As ListBox)
'lstInfectParts_LostFocus�¼���lstAdvEvent_LostFocus�¼���lstInfection_LostFocus�¼�
    lstInput.ListIndex = -1
End Sub

Public Sub LstItemCheck(ByRef lstInput As ListBox, ByRef intItem As Integer)
'lstInfectParts_ItemCheck�¼���lstAdvEvent_ItemCheck�¼���lstInfection_ItemCheck�¼�
    Dim cboTmp As ComboBox
    With gclsPros.CurrentForm
        If lstInput.Name = "lstAdvEvent" Then
            If lstInput.List(intItem) = "ѹ��" Then
                Call SetCtrlLocked(.cboBaseInfo(BCC_ѹ�������ڼ�), Not lstInput.Selected(intItem), True)
                Call SetCtrlLocked(.cboBaseInfo(BCC_ѹ������), Not lstInput.Selected(intItem), True)
            ElseIf lstInput.List(intItem) = "ҽԺ�ڵ���/׹��" Then
                Call SetCtrlLocked(.cboBaseInfo(BCC_������׹���˺�), Not lstInput.Selected(intItem), True)
                Call SetCtrlLocked(.cboBaseInfo(BCC_������׹��ԭ��), Not lstInput.Selected(intItem), True)
            End If
        End If
        Call CheckValueChange
    End With
End Sub

Public Sub LstKeyDown(ByRef lstInput As ListBox, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'���ܣ�lvwInfectParts_KeyDown
    If intKeyCode = vbKeyEscape And lstInput.Name = "lstInfectParts" Then
        Call ShowInfectInfo(False)
    End If
End Sub

Public Sub LstKeyPress(ByRef lstInput As ListBox, ByRef intKeyAscii As Integer)
'���ܣ�lvwInfectParts_KeyPress
    With gclsPros.CurrentForm
        If intKeyAscii = vbKeyReturn Then
             intKeyAscii = 0
            If lstInput.ListIndex = lstInput.ListCount - 1 Then
                If lstInput.Name = "lstAdvEvent" Then
                    If Not .cboBaseInfo(BCC_ѹ�������ڼ�).Locked Or Not .cboBaseInfo(BCC_������׹���˺�).Locked Then
                        If Not .cboBaseInfo(BCC_ѹ�������ڼ�).Locked Then
                            zlControl.ControlSetFocus .cboBaseInfo(BCC_ѹ�������ڼ�)
                        Else
                            zlControl.ControlSetFocus .cboBaseInfo(BCC_������׹���˺�)
                        End If
                    Else
                       Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
                    End If
                ElseIf lstInput.Name = "lstInfection" Then
                    Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
                ElseIf lstInput.Name = "lstInfectParts" Then
                    Call ShowInfectInfo(False)
                    zlControl.ControlSetFocus .vsDiagXY
                End If
            Else
                lstInput.ListIndex = lstInput.ListIndex + 1
            End If
        End If
    End With
End Sub

'lvwFee�¼�
Public Sub lvwFeeItemCheck(ByVal Item As MSComctlLib.ListItem)
'lvwFee_ItemCheck�¼�
    Call AddOrDelFreeCols(gclsPros.CurrentForm.vsFees, Item.Text, Item.SubItems(1), Item.Checked)
End Sub

'cmdFeeEdit�¼�
Public Sub cmdFeeEditClick()
'cmdFeeEdit_Click�¼�
    Dim intSeq As Integer
    Dim i As Integer, LngRow As Long, LngCol As Long
    Dim lstTmp As ListItem

    '1.����б��ɼ�
    If Not gclsPros.CurrentForm.lvwFee.Visible Then
    ''1.���˲����Ƿ����ⲿ�ļ��д���
        gclsPros.FeesOut.Filter = "סԺ�� = " & IIf(gclsPros.InNo = "", 0, gclsPros.InNo)

        If gclsPros.FeesOut.State = adStateClosed Then Exit Sub
        If gclsPros.FeesOut.RecordCount = 0 Then
            MsgBox "סԺ��Ϊ" & gclsPros.InNo & "�Ĳ���������������ⲿ�ļ���û�ҵ���", vbInformation, gstrSysName
            Exit Sub
        Else
        ''2.��ʾ�б�
            gclsPros.CurrentForm.lvwFee.ListItems.Clear
            With gclsPros.FeesOut
                Do While Not .EOF
                    Set lstTmp = gclsPros.CurrentForm.lvwFee.ListItems.Add(, "K" & intSeq, IIf(IsNull(!������), "", !������))
                    lstTmp.SubItems(1) = Format(!���, gclsPros.FreeFormat)
                    lstTmp.SubItems(2) = !סԺ����
                    'lstTmp.Checked = !סԺ���� = mlng��ҳID
                    For i = 3 To gclsPros.CurrentForm.vsFees.Rows * 3
                        LngRow = i \ 3: LngCol = (i Mod 3) * 2
                        If GetTextByDot(gclsPros.CurrentForm.vsFees.TextMatrix(LngRow, LngCol)) = lstTmp.Text And _
                                gclsPros.CurrentForm.vsFees.TextMatrix(LngRow, LngCol + 1) = lstTmp.SubItems(1) And gclsPros.��ҳID = lstTmp.SubItems(2) Then
                            lstTmp.Checked = True
                            Exit For
                        End If
                    Next
                    intSeq = intSeq + 1
                    .MoveNext
                Loop
            End With
            gclsPros.CurrentForm.lvwFee.Visible = True
            gclsPros.CurrentForm.lvwFee.Top = gclsPros.CurrentForm.cmdFeeEdit.Top + gclsPros.CurrentForm.cmdFeeEdit.Height
            gclsPros.CurrentForm.lvwFee.Left = gclsPros.CurrentForm.cmdFeeEdit.Left
        End If
    Else
    '2.����ɼ�,�����б�
        gclsPros.CurrentForm.lvwFee.Visible = False
    End If
End Sub

Public Sub ModifyPatiInfo()
'���ܣ��޸Ĳ��˻�����Ϣ
    On Error GoTo errH
    '��ʼ��������Ϣ�ӿ�
    If gobjPatient Is Nothing Then
        On Error Resume Next
        Set gobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo errH
        Call gobjPatient.zlInitCommon(gcnOracle, gclsPros.SysNo, UserInfo.DBUser)
    End If
    If gobjPatient Is Nothing Then
        MsgBox "����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    Call gobjPatient.zlInitCommon(gcnOracle, gclsPros.SysNo, UserInfo.DBUser)
    '�޸Ĳ��˻�����Ϣ��,ˢ�½�������
    If gobjPatient.ModiPatiBaseInfo(gclsPros.CurrentForm, "סԺ��ҳ", gclsPros.����ID, gclsPros.��ҳID, gclsPros.PatiType, False) Then
        Set gclsPros.PatiInfo = GetPatiMainInfoData(gclsPros.����ID, gclsPros.��ҳID) '������ҳ�Լ�������Ϣ
        Call SetCtrlValues("����", gclsPros.PatiInfo!���� & "")
        Call SetCtrlValues("�Ա�", gclsPros.PatiInfo!�Ա� & "")
        Call SetCtrlValues("����", gclsPros.PatiInfo!���� & "")
        Call SetCtrlValues("��������", gclsPros.PatiInfo!�������� & "")
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub cmdModifyDownGotFocus()
'cmdPriviewDown_GotFocus�¼�
    Call ShowInfectInfo(False)
End Sub

'cmdMakeLog�¼�
Public Sub cmdMakeLogClick()
'���ܣ�cmdMakeLog_Click
    Dim strLog As String, i As Long
    Dim vsTmp As VSFlexGrid

    Set vsTmp = gclsPros.CurrentForm.vsDiagXY
    With vsTmp
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_�������) <> "" Then
                strLog = strLog & "��" & .TextMatrix(i, DI_�������) & IIf(.TextMatrix(i, DI_�Ƿ�����) <> "", "(��)", "")
            End If
        Next
    End With

    Set vsTmp = gclsPros.CurrentForm.vsDiagZY
    With vsTmp
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_�������) <> "" Then
                strLog = strLog & "��" & .TextMatrix(i, DI_�������) & IIf(.TextMatrix(i, DI_�Ƿ�����) <> "", "(��)", "")
            End If
        Next
    End With
    With gclsPros.CurrentForm.txtInfo(GC_ժҪ)
        If strLog <> "" Then
            If .SelStart = 0 And .SelLength = Len(.Text) Then
                .SelStart = Len(.Text)
            End If
            i = .SelStart
            .SelText = Mid(strLog, 2)
            .SelStart = i
            .SelLength = Len(Mid(strLog, 2))
        End If
        .SetFocus
    End With
End Sub

'TxtInfo�¼�
Public Sub CmdInfoClick(ByRef intIndex As Integer)
'���ܣ�cmdInfo_KeyPress
    Dim strSql As String, strCaption As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean, blnMultiSel As Boolean
    Dim strMsg As String, strResult As String
    Dim objTXTBox As TextBox
    Dim bytStyle As Byte

    Select Case intIndex
        Case GC_ת��1, GC_ת��2, GC_ת��3, GC_��Ժ����, GC_��Ժ����, GC_��Ժ����, GC_��Ժ����
            'ѡ��ת�ƿ���
            If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
            grsDeptInfo.Filter = "": Set rsTmp = Rec.Distinct(Rec.FilterNew(grsDeptInfo, "��������='�ٴ�' OR ��������='����'", "ID,����,����,����"))
            strCaption = "�ٴ�����": strMsg = "���Ź���"
        Case GC_��֢�໤������
            If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
            grsDeptInfo.Filter = "": Set rsTmp = Rec.FilterNew(grsDeptInfo, "��������='ICU'", "ID,����,����,����")
            strCaption = "ICU��֢�໤��": strMsg = "���Ź���"
        Case GC_��ԭѧ���

        Case GC_����ԭ��
            Set rsTmp = zlDatabase.CopyNewRec(GetBaseCode("סԺ����ԭ��"), , "ID,����,����,����")
            strCaption = "����ԭ��": strMsg = "�ֵ������"
        Case GC_���Ȳ���
            Set rsTmp = zlDatabase.CopyNewRec(GetBaseCode("���Ȳ������"), , "ID,����,����,����")
            strCaption = "����ԭ��": strMsg = "�ֵ������"
        Case GC_ҽѧ��ʾ
            'ѡ��ҽѧ��ʾ
            strSql = "Select Rownum ID,����,����,���� From ҽѧ��ʾ Order by ����"
            strCaption = "ҽѧ��ʾ": strMsg = "�ֵ������": blnMultiSel = True
        Case GC_ת��ҽ�ƻ���
            strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����,ĩ�� From ��Ժת�� Order By ����"
            strCaption = "��Ժת��": strMsg = "�ֵ������": blnMultiSel = False: bytStyle = 2
        Case GC_��Ժת��
            strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����,ĩ�� From ҽ�ƻ��� Order By ����"
            strCaption = "ҽ�ƻ���": strMsg = "�ֵ������": blnMultiSel = False: bytStyle = 2
    End Select
    '���ݴ���
    On Error GoTo errH
    Set objTXTBox = gclsPros.CurrentForm.txtInfo(intIndex)
    If intIndex <> GC_��ԭѧ��� Then
        If strSql <> "" Then
            vPoint = GetCoordPos(objTXTBox.Container.hwnd, objTXTBox.Left, objTXTBox.Top)
            If blnMultiSel Then
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(gclsPros.CurrentForm, strSql, 0, strCaption, True, "", "", True, True, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel, True, True)
            Else
                Set rsTmp = zlDatabase.ShowSelect(gclsPros.CurrentForm, strSql, bytStyle, strCaption, , , , , True, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel)
            End If
        Else
            blnCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, objTXTBox, rsTmp, True, , , rsTmp)
        End If
    Else
        'D-ICD-10��������
        Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, "D", gclsPros.��Ժ����ID, gclsPros.CurrentForm.cboBaseInfo(BCC_�Ա�).Text, False)
    End If

    If rsTmp Is Nothing Then
        If intIndex = GC_��ԭѧ��� Then Exit Sub
        If Not blnCancel Then
            MsgBox "û������""" & strCaption & """���ݣ����ȵ�" & strMsg & "�����á�", vbInformation, gstrSysName
        End If
        objTXTBox.Tag = ""
        zlControl.ControlSetFocus objTXTBox
    Else
        If blnMultiSel Then '��ѡ�Զ��ŷָ�
            While Not rsTmp.EOF
                strResult = strResult & "," & rsTmp!����
                rsTmp.MoveNext
            Wend
            objTXTBox.Text = Mid(strResult, 2)
        Else
            If intIndex = GC_��ԭѧ��� Then
                objTXTBox.Text = IIf(Not IsNull(rsTmp!����), "(" & rsTmp!���� & ")", "") & NVL(rsTmp!����)
                objTXTBox.Tag = objTXTBox.Text
                gclsPros.CurrentForm.cmdInfo(intIndex).Tag = rsTmp!��ĿID
            Else
                objTXTBox.Text = rsTmp!����
                If gclsPros.FuncType = f������ҳ Then
                    If intIndex = GC_��Ժ���� Then
                        '53638:������,2013-05-10,�����ű������
                        If gclsPros.UseFileRules = True And gclsPros.��Ժ����ID <> Val(rsTmp!ID & "") And Val(gclsPros.InNo) <> 0 Then
                            If IsPageNosCodeRule(CT_������) = True Then
                                gclsPros.CurrentForm.txtInfo(GC_������).Text = NVL(GetNextNo(5, , rsTmp!���� & ""))
                            End If
                        End If
                        gclsPros.��Ժ����ID = Val(rsTmp!ID & "")
                    ElseIf intIndex = GC_��Ժ���� Then
                        gclsPros.��Ժ����ID = Val(rsTmp!ID & "")
                    End If
                    Call SetFaceInit(True)
                    Call SetPageVisible
                    Call SetFaceEditable(gclsPros.IsSigned)
                End If
            End If
        End If
        zlControl.ControlSetFocus objTXTBox
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
    If intIndex = GC_��֢�໤������ Then
        Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_�˹������ѳ�), objTXTBox.Text = "", True)
        Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_�ط���֢ҽѧ��), objTXTBox.Text = "", True)
        Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_�ط����ʱ��), objTXTBox.Text = "", True)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'cmdLastDiag�¼�
Public Sub cmdLastDiagClick()
'cmdLastDiag_Click�¼�
    Dim rsTmp As Recordset
    Dim strSql As String

    If gclsPros.��ҳID = 1 Then Exit Sub
    On Error GoTo errH
    strSql = "Select �������, ������� || '(' || ���� || ')' As �������" & vbNewLine & _
                "From ������ϼ�¼ a, ��������Ŀ¼ b" & vbNewLine & _
                "Where a.����id = b.Id(+) And ����id = [1] And ��ҳid = [2] And ������� = 3 And ��ϴ��� = 1 And" & vbNewLine & _
                "      ��¼��Դ = (Select Max(Nvl(��¼��Դ, 0)) From ������ϼ�¼ Where ����id = [1] And ��ҳid = [2] And ��¼��Դ <= 4)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "", gclsPros.����ID, gclsPros.��ҳID - 1)
    gclsPros.CurrentForm.lblDiagInfo.BorderStyle = 1
    gclsPros.CurrentForm.lblDiagInfo.Visible = True
    If Not rsTmp.EOF Then
        gclsPros.CurrentForm.lblDiagInfo.Caption = NVL(rsTmp!�������)
    Else
        gclsPros.CurrentForm.lblDiagInfo.Caption = "δ�ҵ��ϴ�סԺ����Ҫ�����Ϣ"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub TxtInfoChange(ByRef intIndex As Integer)
'���ܣ�txtInfo_Change
    Dim objTXTBox As TextBox
    Dim lngPos As Long, lnglen As Long

    Set objTXTBox = gclsPros.CurrentForm.txtInfo(intIndex)
    Select Case intIndex
        Case GC_ת��1
            If objTXTBox.Text = "" Then
                gclsPros.CurrentForm.txtInfo(GC_ת��2).Text = ""
                gclsPros.CurrentForm.txtInfo(GC_ת��3).Text = ""
            End If
        Case GC_ת��2
            If objTXTBox.Text = "" Then
                gclsPros.CurrentForm.txtInfo(GC_ת��3).Text = ""
            End If
        Case GC_�໤�����֤��
            If gclsPros.IsReturn Then Exit Sub
            gclsPros.IsReturn = True
            If Not zlStr.CheckCharScope(objTXTBox.Text, "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ*") Then
                objTXTBox.Text = ""
            Else
                If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_����).Text)) = "�й�" Then
                    If zlCommFun.ActualLen(objTXTBox.Text) > 18 Then
                        objTXTBox.Text = Mid(objTXTBox.Text, 1, 18)
                    End If
                End If
            End If
            '������ҳ������ʵ
            If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_����).Text)) = "�й�" Then
                If objTXTBox.Tag <> "������Change�¼�" Then
                    lngPos = InStr(objTXTBox.Text, "*")
                    lnglen = Len(Mid(objTXTBox.Text, 13, 2))
                    Select Case lngPos
                        Case 0
                            objTXTBox.Tag = objTXTBox.Text
                        Case Else 'Is <= 12
                            objTXTBox.Tag = Mid(objTXTBox.Text, 1, lngPos - 1)
                            objTXTBox.Text = objTXTBox.Tag
                            objTXTBox.SelStart = Len(objTXTBox.Text)
                    End Select
                End If '
            Else
                objTXTBox.Tag = objTXTBox.Text
            End If
            gclsPros.IsReturn = False
    End Select
    Call CheckValueChange(objTXTBox)
End Sub

Public Sub TxtInfoGotFocus(ByRef intIndex As Integer)
'���ܣ�txtInfo_GotFocus
    Call ChangeCtl
    If Not (intIndex = GC_ժҪ And gclsPros.CurrentForm.txtInfo(intIndex).SelLength <> 0) Then
        Call TxtGotFocus(gclsPros.CurrentForm.txtInfo(intIndex), True, True)
    End If
End Sub

Public Sub TxtInfoKeyDown(ByRef intIndex As Integer, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'���ܣ�txtInfo_KeyDown
    If intKeyCode = vbKeyDelete Then
        If intIndex = GC_ҽѧ��ʾ Then
            gclsPros.CurrentForm.txtInfo(intIndex).Text = ""
        End If
    End If
End Sub

Public Sub TxtInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'���ܣ�txtInfo_KeyPress
    Dim objTXTBox As TextBox
    Dim strSql As String, strFilter As String, strInput As String
    Dim strCaption As String, strSeek As String, strNote As String
    Dim blnĩ�� As Boolean, blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI

    On Error GoTo errH

    Set objTXTBox = gclsPros.CurrentForm.txtInfo(intIndex)
    If intIndex = GC_�໤�����֤�� Then
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            zlCommFun.PressKey vbKeyTab: mblnReturn = True
        Else
            If Trim(zlCommFun.GetNeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_����).Text)) = "�й�" Then
                If zlCommFun.ActualLen(objTXTBox.Text) >= 18 And intKeyAscii <> vbKeyBack Then
                    intKeyAscii = 0 '���ֻ������18������
                Else
                    If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
                        intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                            intKeyAscii = 0
                        ElseIf zlCommFun.IsCharChinese(objTXTBox.Text) Then
                            objTXTBox.Text = "": objTXTBox.Tag = ""
                        End If
                        If gclsPros.FuncType = fҽ����ҳ And intKeyAscii <> 0 Then
                            '������ҳ������ʵ
                            Select Case zlCommFun.ActualLen(objTXTBox.Text)
                                Case 12
                                    objTXTBox.Tag = objTXTBox.Text & Chr(intKeyAscii)
                                Case 13
                                    objTXTBox.Tag = objTXTBox.Tag & Chr(intKeyAscii)
                            End Select
                        End If
                    End If
                End If
            Else
                If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                        intKeyAscii = 0
                    ElseIf zlCommFun.IsCharChinese(objTXTBox.Text) Then
                        objTXTBox.Text = "": objTXTBox.Tag = ""
                    End If
                    If gclsPros.FuncType = fҽ����ҳ And intKeyAscii <> 0 Then
                        objTXTBox.Tag = objTXTBox.Text & Chr(intKeyAscii)
                    End If
                End If
            End If
        End If
    Else
        If intKeyAscii = vbKeyReturn Then
            If objTXTBox.Text <> "" Then
                strInput = UCase(objTXTBox.Text)
                Select Case intIndex
                    Case GC_ת��1, GC_ת��2, GC_ת��3, GC_��Ժ����, GC_��Ժ����, GC_��֢�໤������
                        If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
                        If intIndex = GC_ת��1 Or intIndex = GC_ת��2 Or intIndex = GC_ת��3 Then
                            grsDeptInfo.Filter = "": Set rsTmp = grsDeptInfo: strCaption = "ת�ƿ���"
                        Else
                            grsDeptInfo.Filter = "�������� = '" & IIf(intIndex = GC_��֢�໤������, "ICU", "�ٴ�") & "'": Set rsTmp = grsDeptInfo: strCaption = "�ٴ�����"
                        End If
                        strFilter = "���� Like '" & strInput & "*' Or ���� Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*' Or ���� Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*'"
                        Set rsTmp = Rec.Distinct(Rec.FilterNew(rsTmp, strFilter, "Id,����,����,����,λ��"), "Id,����,����,����,λ��")
                    Case GC_���Ȳ���
                        strFilter = "���� Like '" & strInput & "*' Or ���� Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*' Or ���� Like '" & IIf(gclsPros.LikeString = "%", "*", "") & strInput & "*'"
                        Set rsTmp = Rec.FilterNew(GetBaseCode("���Ȳ������"), strFilter, "ID,����,����,����")
                        strCaption = "����ԭ��"
                    Case GC_ת��ҽ�ƻ���
                        If zlCommFun.IsCharChinese(strInput) Then
                            strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ��Ժת�� Where ���� Like [1]"
                        Else
                            If gclsPros.BriefCode = 1 Then
                                strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ��Ժת�� Where zlWbCode(����) Like [1]"
                            Else
                                strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ��Ժת�� Where ���� Like [1]"
                            End If
                        End If
                        strCaption = "��Ժת��"
                    Case GC_��Ժת��
                        If zlCommFun.IsCharChinese(strInput) Then
                            strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ҽ�ƻ��� Where ���� Like [1]"
                        Else
                            If gclsPros.BriefCode = 1 Then
                                strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ҽ�ƻ��� Where zlWbCode(����) Like [1]"
                            Else
                                strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ҽ�ƻ��� Where ���� Like [1]"
                            End If
                        End If
                        strCaption = " ҽ�ƻ���"
                End Select
                If strSql <> "" Or strFilter <> "" Then
                    If strSql <> "" Then
                        vPoint = GetCoordPos(objTXTBox.Container.hwnd, objTXTBox.Left, objTXTBox.Top)
                        If intIndex = GC_��Ժת�� Or intIndex = GC_ת��ҽ�ƻ��� Then
                            Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, strCaption, blnĩ��, strSeek, strNote, False, _
                                False, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel, False, False, _
                                gclsPros.LikeString & UCase(objTXTBox.Text) & "%")
                        Else
                            Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, strCaption, blnĩ��, strSeek, strNote, False, _
                                False, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel, False, False, _
                                UCase(objTXTBox.Text) & "%", gclsPros.LikeString & UCase(objTXTBox.Text) & "%")
                        End If
                    Else
                        Call zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, objTXTBox, rsTmp, True, , , rsTmp)
                    End If
                    '������������,��һ��Ҫƥ��
                    If Not rsTmp Is Nothing Then
                        objTXTBox.Text = rsTmp!����
                        If gclsPros.FuncType = f������ҳ Then
                            If intIndex = GC_��Ժ���� Then
                                '53638:������,2013-05-10,�����ű������
                                If gclsPros.UseFileRules = True And gclsPros.��Ժ����ID <> Val(rsTmp!ID & "") And Val(gclsPros.InNo) <> 0 Then
                                    If IsPageNosCodeRule(CT_������) = True Then
                                        gclsPros.CurrentForm.txtInfo(GC_������).Text = NVL(GetNextNo(5, , rsTmp!���� & ""))
                                    End If
                                End If
                                gclsPros.��Ժ����ID = Val(rsTmp!ID & "")
                            ElseIf intIndex = GC_��Ժ���� Then
                                gclsPros.��Ժ����ID = Val(rsTmp!ID & "")
                            End If
                            Call SetFaceInit(True)
                            Call SetPageVisible
                            Call SetFaceEditable(gclsPros.IsSigned)
                        End If
                    Else
                        objTXTBox.Tag = ""
                        If gclsPros.GetMedical Then
                            MsgBox "���ֵ����δ�ҵ�������,������¼�룡", vbInformation, gstrSysName
                            objTXTBox.Text = ""
                            objTXTBox.SetFocus
                        End If
                    End If
                End If
            End If
            intKeyAscii = 0
            'ҽ���Ų������У����üӹ�������
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        ElseIf Not (intKeyAscii >= 0 And intKeyAscii < vbKeySpace) Then
            'ѡ���ݼ�
            If intKeyAscii = Asc("*") Then
                'ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
                On Error Resume Next
                strSql = ""
                strSql = gclsPros.CurrentForm.cmdInfo(intIndex).Name
                Err.Clear: On Error GoTo errH
                If strSql <> "" Then
                    intKeyAscii = 0
                    Call CmdInfoClick(intIndex)
                    Exit Sub
                End If
            End If
    
            '�������볤��
            If objTXTBox.MaxLength <> 0 Then
                If zlCommFun.ActualLen(objTXTBox.Text) > objTXTBox.MaxLength Then
                    intKeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub TxtInfoMouseDown(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'���ܣ�txtInfo_MouseDown
    Call TxtMouseDown(gclsPros.CurrentForm.txtInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

Public Sub TxtInfoMouseUp(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'���ܣ�txtInfo_MouseUp
    Call TxtMouseUp(gclsPros.CurrentForm.txtInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

Public Sub TxtInfoValidate(ByRef intIndex As Integer, ByRef blnCancel As Boolean)
'���ܣ�txtInfo_Validate
    Dim objTXTBox As TextBox, objCmdBtn As CommandButton
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim strInput As String, str�Ա� As String
    Dim blnCancelSel As Boolean
    Dim strMsg As String
    
    On Error GoTo errH
    
    Set objTXTBox = gclsPros.CurrentForm.txtInfo(intIndex)
    On Error Resume Next
    Set objCmdBtn = gclsPros.CurrentForm.cmdInfo(intIndex) '����û�ж�Ӧ�İ�ť
    Err.Clear: On Error GoTo errH
    
    Select Case intIndex
        Case GC_��ԭѧ���
            If objTXTBox.Text = "" Then
                objTXTBox.Tag = ""
                objCmdBtn.Tag = ""
            ElseIf objTXTBox.Text = objTXTBox.Tag Then
                'Nothing
            Else
                strInput = UCase(objTXTBox.Text)
                If gclsPros.CurrentForm.cboBaseInfo(BCC_�Ա�).Text Like "*��*" Then
                    str�Ա� = "��"
                ElseIf gclsPros.CurrentForm.cboBaseInfo(BCC_�Ա�).Text Like "*Ů*" Then
                    str�Ա� = "Ů"
                End If
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSql = "���� Like [1] Or ���� Like [2] Or " & IIf(gclsPros.BriefCode = 0, "����", "�����") & " Like [2]"
                End If
                strSql = _
                    " Select ID,ID as ��ĿID,����,����,����," & IIf(gclsPros.BriefCode = 0, "����", "����� as ����") & ",˵��" & _
                    " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & strSql & ")" & _
                    IIf(str�Ա� <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by ����"

                If gclsPros.DiagSourceZY = 1 And zlCommFun.IsCharChinese(strInput) Then
                    '�����ж��룺Y-�����ж����ⲿԭ�򣻲����������M-������̬ѧ���룻������ϣ�D-ICD-10��������
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput & "%", gclsPros.LikeString & strInput & "%", "'D'", str�Ա�)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                Else
                    vPoint = GetCoordPos(objTXTBox.hwnd, 0, 0)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "��ԭѧ���", False, "", "", False, False, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancelSel, False, True, _
                        strInput & "%", gclsPros.LikeString & strInput & "%", "'D'", str�Ա�, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                    If blnCancelSel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        blnCancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing And (gclsPros.DiagSourceZY = 2 Or gclsPros.DiagSourceZY = 3 And gclsPros.InsureType <> 0) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            blnCancel = True
                        End If
                    End If
                End If
                
                If Not blnCancel Then
                    If rsTmp Is Nothing Then
                        objCmdBtn.Tag = ""
                    Else
                        objTXTBox.Text = IIf(Not IsNull(rsTmp!����), "(" & rsTmp!���� & ")", "") & NVL(rsTmp!����)
                        objTXTBox.Tag = objTXTBox.Text
                        objCmdBtn.Tag = rsTmp!��ĿID
                    End If
                End If
            End If
        Case GC_��֢�໤������
            strInput = Trim(objTXTBox.Text)
            If strInput <> "" Then
                If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
                grsDeptInfo.Filter = "": Set rsTmp = Rec.FilterNew(grsDeptInfo, "��������='ICU'", "ID,����,����,����")
                rsTmp.Filter = "����='" & strInput & "'"
                If rsTmp.EOF Then
                    rsTmp.Filter = "���� Like '" & strInput & "*' OR ���� Like '" & strInput & "*' OR ���� Like '" & IIf(gclsPros.LikeString <> "", "*", "") & strInput & "*' "
                    blnCancelSel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, objTXTBox, rsTmp, True, , , rsTmp)
                    If rsTmp Is Nothing Then
                        If Not blnCancel Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            blnCancel = True
                        End If
                    Else
                        objTXTBox.Text = rsTmp!����
                    End If
                End If
            End If
            Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_�˹������ѳ�), objTXTBox.Text = "" Or blnCancel, True)
            Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_�ط���֢ҽѧ��), objTXTBox.Text = "" Or blnCancel, True)
            Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_�ط����ʱ��), objTXTBox.Text = "" Or blnCancel, True)
    Case GC_ת��ҽ�ƻ���, GC_��Ժת��
        strInput = UCase(objTXTBox.Text)
        If strInput = "" Then
            objTXTBox.Tag = ""
        Else
            
            If zlCommFun.IsCharChinese(strInput) Then
                If intIndex = GC_ת��ҽ�ƻ��� Then
                    strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ��Ժת�� Where ���� Like [1]"
                ElseIf intIndex = GC_��Ժת�� Then
                    strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ҽ�ƻ��� Where ���� Like [1]"
                End If
            Else
                If gclsPros.BriefCode = 1 Then
                    If intIndex = GC_ת��ҽ�ƻ��� Then
                        strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ��Ժת�� Where zlWbCode(����) Like [1]"
                    ElseIf intIndex = GC_��Ժת�� Then
                        strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ҽ�ƻ��� Where zlWbCode(����) Like [1]"
                    End If
                Else
                    If intIndex = GC_ת��ҽ�ƻ��� Then
                        strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ��Ժת�� Where ���� Like [1]"
                    ElseIf intIndex = GC_��Ժת�� Then
                        strSql = "Select ���� As ID, ����, �ϼ� As �ϼ�id, ����, ����, ĩ�� From ҽ�ƻ��� Where ���� Like [1]"
                    End If
                End If
            End If
            If strSql <> "" Then
                vPoint = GetCoordPos(objTXTBox.Container.hwnd, objTXTBox.Left, objTXTBox.Top)
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "ҽ�ƻ���", True, False, True, False, _
                    False, True, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel, False, False, _
                      gclsPros.LikeString & UCase(objTXTBox.Text) & "%")
                If Not rsTmp Is Nothing Then
                    objTXTBox.Text = rsTmp!����
                Else
                    objTXTBox.Tag = ""
                    If gclsPros.GetMedical Then
                        MsgBox "���ֵ����δ�ҵ�������,������¼��", vbInformation, gstrSysName
                        objTXTBox.Text = ""
                        objTXTBox.SetFocus
                    End If
                End If
            End If
        End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'AdressInfo�¼�
Public Sub CmdAdressInfoClick(ByRef intIndex As Integer)
'���ܣ�cmdAdressInfo_Click
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim objTXTBox As TextBox
    Dim bytStyle As Byte, strCaption As String, strMsg As String, blnRoot As Boolean, blnNonWin As Boolean

    On Error GoTo errH
    Select Case intIndex
        Case ADRC_��λ��ַ
            'ѡ��λ��Ϣ
            strSql = "Select ID,�ϼ�ID,ĩ��,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ��" & _
                " From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strCaption = "��Լ��λ": strMsg = "��Լ��λ����": bytStyle = 2: blnRoot = True: blnNonWin = True
        Case ADRC_�����ص�, ADRC_��סַ, ADRC_��ϵ�˵�ַ, ADRC_���ڵ�ַ
            'ѡ���������
            strSql = "Select Rownum as ID,����,����,���� From ���� Order by ����"
            strCaption = "����": strMsg = "�ֵ������": bytStyle = 0: blnRoot = False: blnNonWin = True
        Case ADRC_��������, ADRC_����
            'ѡ����������
            strSql = "Select 1  From ���� Where Nvl(����,0)<>0 And RowNum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption)
            If rsTmp.RecordCount > 0 Then bytStyle = 2
            If bytStyle = 2 Then
                strSql = _
                        "Select Id, �ϼ�id, Id ����, ����, ����, ĩ��" & vbNewLine & _
                        "From (Select Rpad(����, 15, '0') As Id, Rpad(Substr(����, 1, Decode(Nvl(����, 0), 0, 0, 1, 2, 4)), 15, '0') As �ϼ�id, ����, ����," & vbNewLine & _
                        "              Decode(Nvl(����, 0), 2, 1, 3, 1, 0) As ĩ��" & vbNewLine & _
                        "       From ����" & vbNewLine & _
                        "       Where Nvl(����, 0) < 3" & vbNewLine & _
                        "       Order By ����)" & vbNewLine & _
                        "Start With �ϼ�id Is Null" & vbNewLine & _
                        "Connect By Prior Id = �ϼ�id"
            Else
                strSql = "Select Rownum as ID,����,����,���� From ���� Order by ����"
            End If
            strCaption = "����": strMsg = "�ֵ������": blnRoot = False: blnNonWin = IIf(bytStyle = 0, True, False)
    End Select

    '���ݴ���
    On Error GoTo errH
    '���ݴ���
    Set objTXTBox = gclsPros.CurrentForm.txtAdressInfo(intIndex)
    vPoint = GetCoordPos(objTXTBox.Container.hwnd, objTXTBox.Left, objTXTBox.Top)
    Set rsTmp = zlDatabase.ShowSelect(gclsPros.CurrentForm, strSql, bytStyle, strCaption, , , , , blnRoot, blnNonWin, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel)

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û������""" & IIf(strCaption = "����", "����", strCaption) & """���ݣ����ȵ�" & strMsg & "�����á�", vbInformation, gstrSysName
        End If
        objTXTBox.Tag = ""
        zlControl.ControlSetFocus objTXTBox
    Else
        If intIndex = ADRC_��λ��ַ Then
            objTXTBox.Text = rsTmp!���� & IIf(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
            If gclsPros.PatiType = PF_���� Then
                If InStr(gclsPros.Privs, "��Լ���˵Ǽ�") > 0 Then objTXTBox.Tag = Val(rsTmp!ID)
            Else
                objTXTBox.Tag = Val(rsTmp!ID)
            End If
            If gclsPros.CurrentForm.txtSpecificInfo(SLC_��λ�绰).Text = "" Then
                gclsPros.CurrentForm.txtSpecificInfo(SLC_��λ�绰).Text = NVL(rsTmp!�绰)
            End If
            objTXTBox.SetFocus
        Else
            objTXTBox.Text = rsTmp!����
            objTXTBox.SetFocus
            If intIndex = ADRC_�����ص� And gclsPros.FuncType = f������ҳ And gclsPros.DefautADD Then
                Call SetPatiAddress(ADRC_��ϵ�˵�ַ, "��ϵ�˵�ַ", rsTmp!����, True)
                Call SetPatiAddress(ADRC_��סַ, "��ͥ��ַ", rsTmp!����, True)
                gclsPros.CurrentForm.txtSpecificInfo(SLC_��ͥ�ʱ�).Text = rsTmp!���� & ""
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub txtAdressInfoGotFocus(ByRef intIndex As Integer)
'���ܣ�txtAdressInfo_GotFocus
    Call ChangeCtl
    Call TxtGotFocus(gclsPros.CurrentForm.txtAdressInfo(intIndex), True, True)
End Sub

Public Sub txtAdressInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'txtAdressInfo_KeyPress�¼�
    Dim objBox As TextBox
    Dim strSql As String, strCaption As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI

    Set objBox = gclsPros.CurrentForm.txtAdressInfo(intIndex)

    If intKeyAscii = vbKeyReturn Then
        If objBox.Text <> "" Then
            Select Case intIndex
                Case ADRC_��λ��ַ
                    'ѡ��λ��Ϣ
                    strSql = "Select ID,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
                        " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                        " And �ϼ�id Is not Null and (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                        " Order by ����"
                    strCaption = "������λ"
                Case ADRC_�����ص�, ADRC_��סַ, ADRC_��ϵ�˵�ַ, ADRC_���ڵ�ַ
                    '�����������
                    strSql = "Select Rownum as ID,����,����,���� From ���� " & _
                        " Where (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                        " Order by ����"
                    strCaption = "����"
                Case ADRC_��������, ADRC_����
                    '������������
                    strSql = "Select Rownum as ID,����,����,���� From ���� " & _
                        " Where (���� Like [1] Or ���� Like [2] Or ���� Like [2]) And Nvl(����, 0) < 3" & _
                        " Order by ����"
                    strCaption = IIf(intIndex = ADRC_��������, "����", "����")
            End Select

            If strSql <> "" Then
                vPoint = GetCoordPos(objBox.Container.hwnd, objBox.Left, objBox.Top)
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, strCaption, False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, objBox.Height, blnCancel, False, False, _
                    UCase(objBox.Text) & "%", gclsPros.LikeString & UCase(objBox.Text) & "%")
                '������������,��һ��Ҫƥ��
                If Not rsTmp Is Nothing Then
                    If intIndex = ADRC_��λ��ַ Then
                        objBox.Text = rsTmp!���� & IIf(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                        If gclsPros.PatiType = PF_���� Then
                            If InStr(gclsPros.Privs, "��Լ���˵Ǽ�") > 0 Then objBox.Tag = Val(rsTmp!ID)
                        Else
                            objBox.Tag = Val(rsTmp!ID)
                        End If
                        If gclsPros.CurrentForm.txtSpecificInfo(SLC_��λ�绰).Text = "" Then
                            gclsPros.CurrentForm.txtSpecificInfo(SLC_��λ�绰).Text = NVL(rsTmp!�绰)
                        End If
                    Else
                        objBox.Text = rsTmp!����
                        If intIndex = ADRC_�����ص� And gclsPros.FuncType = f������ҳ And gclsPros.DefautADD Then
                            Call SetPatiAddress(ADRC_��ϵ�˵�ַ, "��ϵ�˵�ַ", rsTmp!����, True)
                            Call SetPatiAddress(ADRC_��סַ, "��ͥ��ַ", rsTmp!����, True)
                            gclsPros.CurrentForm.txtSpecificInfo(SLC_��ͥ�ʱ�).Text = rsTmp!���� & ""
                        End If
                    End If
                Else
                    objBox.Tag = ""
                End If
                objBox.SetFocus
                Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
            End If
        Else
            intKeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        End If
    Else
        If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
            'ѡ���ݼ�
            If intKeyAscii = Asc("*") Then
                'ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
                On Error Resume Next
                strSql = ""
                strSql = gclsPros.CurrentForm.cmdAdressInfo(intIndex).Name
                Err.Clear: On Error GoTo 0
                If strSql <> "" Then
                    intKeyAscii = 0
                    Call CmdAdressInfoClick(intIndex)
                    Exit Sub
                End If
            End If

            '�������볤��
            If objBox.MaxLength <> 0 Then
                If zlCommFun.ActualLen(objBox.Text) > objBox.MaxLength Then
                    intKeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End If
End Sub

Public Sub txtAdressInfoMouseDown(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'txtAdressInfo_MouseDown�¼�
    Call TxtMouseDown(gclsPros.CurrentForm.txtAdressInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

Public Sub txtAdressInfoMouseUp(ByRef intIndex As Integer, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'txtAdressInfo_MouseUp�¼�
    Call TxtMouseUp(gclsPros.CurrentForm.txtAdressInfo(intIndex), intButton, intShift, sngX, sngY)
End Sub

'vsChemoth�¼�
Public Sub ChemothAfterEdit(ByRef vsChemoth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsChemoth_AfterEdit�¼�
    Dim strInput As String
    With vsChemoth
        If LngCol = CI_��ѧ���Ʊ��� Then
            If .ComboIndex < 0 Then Exit Sub
           .TextMatrix(LngRow, CI_����ID) = .ComboData(.ComboIndex)
        ElseIf LngCol = CI_�������� Or LngCol = CI_��ʼ���� Then
            strInput = zlStr.FullDate(.TextMatrix(LngRow, LngCol), False, gclsPros.InTime, gclsPros.OutTime)
            If Not IsDate(strInput) Then
                .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
            Else
                .TextMatrix(LngRow, LngCol) = strInput
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
            End If
        End If
    End With
End Sub

Public Sub ChemothAfterRowColChange(ByRef vsChemoth As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsChemoth_AfterRowColChange �¼�
'    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
'    Call zlVsGridRowChange(vsChemoth, lngOldRow, lngNewRow, lngOldCol, lngNewCol)
End Sub

Public Sub ChemothGotFocus(ByRef vsChemoth As VSFlexGrid)
'vsChemoth_GotFocus�¼�
'    Call zlVsGridGotFocus(vsChemoth)
End Sub

Public Sub ChemothKeyDown(ByRef vsChemoth As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsChemoth_KeyDown�¼�
   Dim LngCol As Long

    With vsChemoth
        If intKeyCode = vbKeyDelete And .Editable <> flexEDNone Then
            If MsgBox("���Ƿ����Ҫɾ�����еĻ�����Ϣ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = .FixedRows Then
                For LngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, LngCol) = ""
                    .Cell(flexcpData, .Row, LngCol) = ""
                    .RowData(.Row) = 0
                Next
            Else
                .RemoveItem .Row
                 Call ChangeVSFHeight(vsChemoth, True)
            End If
            zlControl.ControlSetFocus vsChemoth, True
        ElseIf intKeyCode = vbKeyReturn Then
            If .TextMatrix(.Row, CI_��ѧ���Ʊ���) = "" And .Col = CI_����Ч�� Then
                zlControl.ControlSetFocus gclsPros.CurrentForm.vsRadioth, True
                Exit Sub
            End If
            
            Select Case .Col
                Case .Cols - 1, CI_����Ч��
                    If Not .Row >= .Rows - 1 Then
                        .Col = 0
                        .Row = .Row + 1
                    Else
                        Call ChemothKeyDownEdit(vsChemoth, .Row, .Col, intKeyCode, intShift)
                    End If
                    .SetFocus
                Case Else
                    zlCommFun.PressKey vbKeyRight
            End Select
        End If
    End With
End Sub

Public Sub ChemothKeyDownEdit(ByRef vsChemoth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyCode As Integer, ByVal intShift As Integer)
'vsChemoth_KeyDownEdit�¼�
    Dim lngCurRow As Long
    Dim strKEY As String

    If intKeyCode <> vbKeyReturn Then Exit Sub

    With vsChemoth
        Call zlVsMoveGridCell(vsChemoth, CI_��ѧ���Ʊ���, .Cols - 1, True, lngCurRow)
        If lngCurRow > 0 Then
            '��ʾ��������һ��,��Ҫ������ص�ȱʡֵ
            strKEY = .ColData(CI_��ѧ���Ʊ���)
            If InStr(1, strKEY, ";") > 0 Then
                .TextMatrix(lngCurRow, CI_��ѧ���Ʊ���) = Mid(strKEY, InStr(1, strKEY, ";") + 1)
                .Cell(flexcpData, lngCurRow, CI_��ѧ���Ʊ���) = Mid(strKEY, 1, InStr(1, strKEY, ";") - 1)
                .TextMatrix(lngCurRow, CI_����ID) = .Cell(flexcpData, lngCurRow, CI_��ѧ���Ʊ���)
                .TextMatrix(lngCurRow, CI_�Ƴ���) = 1
                .Col = CI_��ʼ����
            End If
        End If
    End With
End Sub

Public Sub ChemothKeyPress(ByRef vsChemoth As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsChemoth_KeyPress�¼�
    If vsChemoth.Editable = flexEDNone Then Exit Sub
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
    End If
End Sub

Public Sub ChemothKeyPressEdit(ByRef vsChemoth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsChemoth_KeyPressEdit �¼�
    If intKeyAscii = Asc("'") Then intKeyAscii = 0: Exit Sub
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Exit Sub
    End If
    With gclsPros.CurrentForm.vsChemoth
        Select Case LngCol
            Case CI_��ѧ���Ʊ���, CI_���Ʒ���
                Call VsFlxGridCheckKeyPress(vsChemoth, LngRow, LngCol, intKeyAscii, m�ı�ʽ)
            Case CI_��ʼ����, CI_��������
                If InStr("0123456789-" & Chr(8) & Chr(27), Chr(intKeyAscii)) = 0 Then
                    intKeyAscii = 0
                End If
                Call VsFlxGridCheckKeyPress(vsChemoth, LngRow, LngCol, intKeyAscii, m�ı�ʽ)
            Case CI_�Ƴ���, CI_����
                Call VsFlxGridCheckKeyPress(vsChemoth, LngRow, LngCol, intKeyAscii, m����ʽ)
        End Select
    End With
End Sub

Public Sub ChemothLostFocus(ByRef vsChemoth As VSFlexGrid)
'vsChemoth_LostFocus�¼�
'    Call zlVsGridLostFocus(vsChemoth)
End Sub

Public Sub ChemothValidateEdit(ByRef vsChemoth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsChemoth_ValidateEdit�¼�
    Dim strInput As String

    With vsChemoth
        strInput = Trim(.EditText): strInput = Replace(strInput, Chr(vbKeyReturn), ""): strInput = Replace(strInput, Chr(10), "")
        If strInput = "" Then Exit Sub

        Select Case LngCol
            Case CI_��ѧ���Ʊ���

            Case CI_��ʼ����, CI_��������
                strInput = zlStr.FullDate(strInput, False, gclsPros.InTime, gclsPros.OutTime)
                If IsDate(strInput) Then
                    If Not CheckDateRange(strInput) Then
                        MsgBox "�������ʱ������ڲ��˵�סԺ�ڼ䡣", vbInformation, gstrSysName
                        .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                        blnCancel = True
                        Exit Sub
                    Else
                        .TextMatrix(LngRow, LngCol) = strInput
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                    End If
                    If IsDate(Trim(.TextMatrix(LngRow, IIf(LngCol = CI_��������, CI_��ʼ����, CI_��������)))) Then
                        If LngCol = CI_�������� And CDate(strInput) < CDate(Trim(.TextMatrix(LngRow, CI_��ʼ����))) Or _
                            LngCol = CI_��ʼ���� And CDate(Trim(.TextMatrix(LngRow, CI_��������))) < CDate(strInput) Then
                            MsgBox "�������ڲ���С�ڿ�ʼ����,����!", vbInformation, gstrSysName
                            blnCancel = True
                            Exit Sub
                        End If
                    End If
'                    Call zlVsMoveGridCell(vsChemoth, CI_��ʼ����, .Cols - 1, True, LngCol)
                Else
                    MsgBox IIf(LngCol = CI_��������, "��������", "��ʼ����") & "����Ϊ������,���飡", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
            Case CI_���Ʒ���
                blnCancel = Not zlCommFun.StrIsValid(strInput, 50, 0, "���Ʒ���")
            Case CI_�Ƴ���
                If DblIsValid(strInput, 3, True, False, 0, .TextMatrix(0, LngCol)) = False Then blnCancel = True: Exit Sub
                If strInput = "" Then blnCancel = True: Exit Sub
                .EditText = strInput
            Case CI_����
                If DblIsValid(strInput, 10, True, False, 0, .TextMatrix(0, LngCol)) = False Then blnCancel = True: Exit Sub
                If strInput = "" Then blnCancel = True: Exit Sub
                .EditText = strInput
        End Select
    End With
End Sub

'vsRadioth�¼�
Public Sub RadiothAfterEdit(ByRef vsRadioth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsRadioth_AfterEdit�¼�
    Dim strInput As String
    With vsRadioth
        If LngCol = RI_�������Ʊ��� Then
            If .ComboIndex < 0 Then Exit Sub
            .TextMatrix(LngRow, RI_����ID) = .ComboData(.ComboIndex)
        ElseIf LngCol = RI_�������� Or LngCol = RI_��ʼ���� Then
            strInput = zlStr.FullDate(strInput, False, gclsPros.InTime, gclsPros.OutTime)
            If Not IsDate(strInput) Then
                .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
            Else
                .TextMatrix(LngRow, LngCol) = strInput
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
            End If
        End If
    End With
End Sub

Public Sub RadiothAfterRowColChange(ByRef vsRadioth As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsRadioth_AfterRowColChange�¼�
'    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
'    Call zlVsGridRowChange(vsRadioth, lngOldRow, lngNewRow, lngOldCol, lngNewCol)
End Sub

Public Sub RadiothGotFocus(ByRef vsRadioth As VSFlexGrid)
'vsRadioth_GotFocus�¼�
'    Call zlVsGridGotFocus(vsRadioth)
End Sub

Public Sub RadiothKeyDown(ByRef vsRadioth As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsRadioth_KeyDown�¼�
    Dim LngCol As Long

    With vsRadioth
        If intKeyCode = vbKeyDelete And .Editable <> flexEDNone Then
            If MsgBox("���Ƿ����Ҫɾ�����еķ�����Ϣ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = .FixedRows Then
                For LngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, LngCol) = ""
                    .Cell(flexcpData, .Row, LngCol) = ""
                    .RowData(.Row) = 0
                Next
            Else
                .RemoveItem .Row
                Call ChangeVSFHeight(vsRadioth, True)
            End If
            zlControl.ControlSetFocus vsRadioth, True
        ElseIf intKeyCode = vbKeyReturn Then
            If .TextMatrix(.Row, RI_�������Ʊ���) = "" And .Col = RI_����Ч�� Then
                Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
                Exit Sub
            End If
            
            Select Case .Col
                Case .Cols - 1, RI_����Ч��
                    If Not .Row >= .Rows - 1 Then
                        .Col = RI_�������Ʊ���
                        .Row = .Row + 1
                    Else
                        Call RadiothKeyDownEdit(vsRadioth, .Row, .Col, intKeyCode, intShift)
                    End If
                    .SetFocus
                Case Else
                    zlCommFun.PressKey vbKeyRight
            End Select
        End If
    End With
End Sub

Public Sub RadiothKeyDownEdit(ByRef vsRadioth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyCode As Integer, ByVal intShift As Integer)
'vsRadioth_KeyDownEdit�¼�
    Dim lngCurRow As Long, strKEY As String

    If intKeyCode <> vbKeyReturn Then Exit Sub

    With vsRadioth
        Call zlVsMoveGridCell(vsRadioth, RI_�������Ʊ���, .Cols - 1, True, lngCurRow)
        If lngCurRow > 0 Then
'            ��ʾ��������һ�� , ��Ҫ������ص�ȱʡֵ
            strKEY = .ColData(RI_�������Ʊ���)
            If InStr(1, strKEY, ";") > 0 Then
                .TextMatrix(lngCurRow, RI_�������Ʊ���) = Mid(strKEY, InStr(1, strKEY, ";") + 1)
                .Cell(flexcpData, lngCurRow, RI_�������Ʊ���) = Mid(strKEY, 1, InStr(1, strKEY, ";") - 1)
                .TextMatrix(lngCurRow, RI_����ID) = .Cell(flexcpData, lngCurRow, RI_�������Ʊ���)
                .Col = RI_��ʼ����
            End If
        End If
    End With
End Sub

Public Sub RadiothKeyPress(ByRef vsRadioth As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsRadioth_KeyPress�¼�
    If vsRadioth.Editable = flexEDNone Then Exit Sub
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
    End If
End Sub

Public Sub RadiothKeyPressEdit(ByRef vsRadioth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsRadioth_KeyPressEdit�¼�
    Dim strInput As String
    With vsRadioth
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            If .Row = .Rows - 1 And .TextMatrix(.Row, RI_�������Ʊ���) = "" Then
                Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
            End If
            Exit Sub
        End If
        Select Case LngCol
            Case RI_�������Ʊ���, RI_��Ұ��λ
                Call VsFlxGridCheckKeyPress(vsRadioth, LngRow, LngCol, intKeyAscii, m�ı�ʽ)
            Case RI_��ʼ����, RI_��������
                If InStr("0123456789-" & Chr(8) & Chr(27), Chr(intKeyAscii)) = 0 Then
                    intKeyAscii = 0
                End If
                Call VsFlxGridCheckKeyPress(vsRadioth, LngRow, LngCol, intKeyAscii, m�ı�ʽ)
            Case RI_�������, RI_�ۼ���
                Call VsFlxGridCheckKeyPress(vsRadioth, LngRow, LngCol, intKeyAscii, m����ʽ)
            Case RI_����Ч��
        End Select
    End With
End Sub

Public Sub RadiothLostFocus(ByRef vsRadioth As VSFlexGrid)
'vsRadioth_LostFocus�¼�
'    Call zlVsGridLostFocus(vsRadioth)
End Sub

Public Sub RadiothValidateEdit(ByRef vsRadioth As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsRadioth_ValidateEdit�¼�
    Dim strInput As String

    With vsRadioth
        strInput = Trim(.EditText): strInput = Replace(strInput, Chr(vbKeyReturn), ""): strInput = Replace(strInput, Chr(10), "")
        If strInput = "" Then Exit Sub
        Select Case LngCol
            Case RI_�������Ʊ���
            Case RI_��ʼ����, RI_��������
                strInput = zlStr.FullDate(strInput, False, gclsPros.InTime, gclsPros.OutTime)
                If IsDate(strInput) Then
                    If Not CheckDateRange(strInput) Then
                        MsgBox "�������ʱ������ڲ��˵�סԺ�ڼ䡣", vbInformation, gstrSysName
                        .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                        blnCancel = True
                        Exit Sub
                    Else
                        .TextMatrix(LngRow, LngCol) = strInput
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                    End If
                    If IsDate(Trim(.TextMatrix(LngRow, IIf(LngCol = RI_��������, RI_��ʼ����, RI_��������)))) Then
                        If LngCol = RI_�������� And CDate(strInput) < CDate(Trim(.TextMatrix(LngRow, RI_��ʼ����))) Or _
                            LngCol = RI_��ʼ���� And CDate(Trim(.TextMatrix(LngRow, RI_��������))) < CDate(strInput) Then
                            MsgBox "�������ڲ���С�ڿ�ʼ����,����!", vbInformation, gstrSysName
                            blnCancel = True
                            Exit Sub
                        End If
                    End If
                    Call zlVsMoveGridCell(vsRadioth, RI_��ʼ����, .Cols - 1, True, LngCol)
                Else
                    MsgBox IIf(LngCol = RI_��������, "��������", "��ʼ����") & "����Ϊ������,���飡", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
            Case RI_��Ұ��λ
                blnCancel = Not zlCommFun.StrIsValid(strInput, 50, 0, "��Ұ��λ")
            Case RI_�������, RI_�ۼ���
                If DblIsValid(strInput, 10, True, False, 0, .TextMatrix(0, LngCol)) = False Then blnCancel = True: Exit Sub
                If strInput = "" Then blnCancel = True: Exit Sub
                .EditText = strInput
            Case RI_����Ч��
        End Select
    End With
End Sub

'vsFlxAddICU�¼�
Public Sub FlxAddICUAfterEdit(ByRef vsFlxAddICU As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsFlxAddICU_AfterEdit�¼�
    Dim strList As String, i As Long
    Dim strInput As String

    With vsFlxAddICU
        If LngCol = UI_�໤������ Then
            If gclsPros.MedPageSandard = ST_��������׼ Then
                .ColComboList(0) = "..."
            ElseIf gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, UI_���) = i
                Next
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, UI_�໤������) <> "" Then
                        strList = strList & "|" & .TextMatrix(i, UI_���) & "-" & .TextMatrix(i, UI_�໤������)
                    End If
                Next
                strList = Mid(strList, 2)
                gclsPros.CurrentForm.vsICUInstruments.ColComboList(TI_ICU����) = strList
                gclsPros.CurrentForm.vsICUInstruments.Editable = IIf(strList <> "", flexEDKbdMouse, flexEDNone)
            End If
        ElseIf LngCol = UI_����ʱ�� Or LngCol = UI_�˳�ʱ�� Then
            strInput = zlStr.FullDate(.TextMatrix(LngRow, LngCol), , gclsPros.InTime, gclsPros.OutTime)
            If IsDate(strInput) Then
                .TextMatrix(LngRow, LngCol) = strInput
            End If
        End If
    End With
End Sub

Public Sub FlxAddICUCellButtonClick(ByRef vsFlxAddICU As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsFlxAddICU_CellButtonClick�¼�
    Dim strSql As String, rsTmp As Recordset, vPoint As POINTAPI, blnCancel As Boolean

    With vsFlxAddICU
        Select Case LngCol
            Case UI_�໤������
                strSql = " Select Distinct A.ID,A.����,A.����" & _
                        " From ���ű� A,��������˵�� B" & _
                        " Where B.����ID=A.ID And B.��������='ICU'" & _
                        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " Order by A.����"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "��֢�໤��", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)

                If rsTmp Is Nothing Then
                    If Not blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        MsgBox "û������ICU��֢�໤�ҡ�", vbInformation, gstrSysName
                    End If
                Else
                    .TextMatrix(LngRow, LngCol) = rsTmp!���� & ""
                    .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                End If
            Case Else
        End Select
    End With
End Sub

Public Sub FlxAddICUEnterCell(ByRef vsFlxAddICU As VSFlexGrid)
'vsFlxAddICU_EnterCell�¼�

    Dim datTemp As Date

    With vsFlxAddICU
        If vsFlxAddICU.Cols <> 0 Then
            On Error Resume Next
            If .TextMatrix(.Row, UI_�໤������) <> "" And Trim(.TextMatrix(.Row, UI_����ʱ��)) = "" Then
                '���Ƿ��һ��
                If .Row > 1 Then
                    If .TextMatrix(.Row - 1, UI_�໤������) <> "" And IsDate(.TextMatrix(.Row - 1, UI_�˳�ʱ��)) Then
                        datTemp = CDate(.TextMatrix(.Row - 1, UI_�˳�ʱ��))
                        If Format(datTemp, "yyyy-mm-dd HH:MM") < Format(gclsPros.InTime, "yyyy-mm-dd HH:MM") Then
                            .TextMatrix(.Row, UI_����ʱ��) = ""
                        Else
                            '������Ϊ׼
                            .TextMatrix(.Row, UI_����ʱ��) = Trim(.TextMatrix(.Row - 1, UI_�˳�ʱ��))
                        End If
                        datTemp = DateAdd("d", 1, datTemp)
    
                        If Format(datTemp, "yyyy-mm-dd HH:MM") < Format(gclsPros.InTime, "yyyy-mm-dd HH:MM") Or Format(datTemp, "yyyy-mm-dd HH:MM") > Format(gclsPros.OutTime, "yyyy-mm-dd HH:MM") Then
                            .TextMatrix(.Row, UI_�˳�ʱ��) = ""
                        Else
                            .TextMatrix(.Row, UI_�˳�ʱ��) = Format(datTemp, "yyyy-mm-dd HH:MM")
                        End If
                    Else
                        '�����ԺΪ׼
                        .TextMatrix(.Row, UI_����ʱ��) = Format(gclsPros.InTime, "yyyy-mm-dd HH:MM")
                        .TextMatrix(.Row, UI_�˳�ʱ��) = Format(gclsPros.OutTime, "yyyy-mm-dd HH:MM")
                    End If
                Else
                    If Trim(.TextMatrix(.Row, UI_�˳�ʱ��)) = "" Then
                        '�����ԺΪ׼
                        .TextMatrix(.Row, UI_����ʱ��) = Format(gclsPros.InTime, "yyyy-mm-dd HH:MM")
                        .TextMatrix(.Row, UI_�˳�ʱ��) = Format(gclsPros.OutTime, "yyyy-mm-dd HH:MM")
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub VSFlxGotFocus(ByRef vsFlex As VSFlexGrid)
'vsFlex_GotFocus�¼�
    Call ChangeCtl
    Select Case vsFlex.Name
        Case "vsChemoth"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, CI_��ѧ���Ʊ���, CI_����Ч��, 1, CI_��ѧ���Ʊ���)
        Case "vsRadioth"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, RI_�������Ʊ���, RI_����Ч��, 1, RI_�������Ʊ���)
        Case "vsSpirit"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, SI_ҩ������, SI_��Ч, 1, SI_ҩ������)
        Case "vsKSS"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, KI_����ҩ����, KI_������ҩ, 1, KI_����ҩ����)
        Case "vsFlxAddICU"
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, UI_�໤������, UI_����סԭ��, 1, UI_�໤������)
            Else
                Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, UI_�໤������, UI_�˳�ʱ��, 1, UI_�໤������)
            End If
        Case "vsfMain"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, 0, vsFlex.Cols - 1, 1, 1)
            If vsFlex.TextMatrix(0, vsFlex.Col) = "��Ŀ" Then vsFlex.Col = vsFlex.Col + 1
        Case "vsInfect"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, FI_ȷ������, FI_ҽԺ��Ⱦ����, 1, FI_ȷ������)
        Case "vsSample"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, MI_�걾, MI_�ͼ�����, 1, MI_�걾)
        Case "vsTSJC"
            Call LocateVSFRowCol(vsFlex, 1, vsFlex.Rows - 1, 1, vsFlex.Cols - 1, 0, 1)
    End Select
End Sub

Public Sub FlxAddICUKeyDown(ByRef vsFlxAddICU As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsFlxAddICU_KeyDown�¼�
    Dim i As Long, LngCol As Long
    Dim vsTmp As VSFlexGrid
    Dim int��� As Integer
    Dim lngRevRow As Long
    Dim strType As String

    If vsFlxAddICU.Editable = flexEDNone Then Exit Sub
    With vsFlxAddICU
        If intKeyCode = vbKeyDelete Then
            If MsgBox("���Ƿ����Ҫɾ������������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                '�Ĵ���ɾ����֢�໤ʱͬ���޸���֢�໤�������֢�໤��е��ICU�������
                Set vsTmp = gclsPros.CurrentForm.vsICUInstruments
                int��� = Val(.TextMatrix(.Row, UI_���))
                For i = 1 To vsTmp.Rows - 1
                    If i < vsTmp.Rows Then
                        If Val(vsTmp.Cell(flexcpData, i, TI_ICU����)) = int��� Then
                            If strType = "" Then
                                lngRevRow = i
                                If lngRevRow <> 0 Then vsTmp.RemoveItem lngRevRow
                                i = i - 1
                            ElseIf Split(vsTmp.TextMatrix(i, TI_ICU����), "-")(1) = .TextMatrix(.Row, UI_�໤������) Then
                                lngRevRow = i
                                If lngRevRow <> 0 Then vsTmp.RemoveItem lngRevRow
                                i = i - 1
                            End If
                        ElseIf Val(vsTmp.Cell(flexcpData, i, TI_ICU����)) > int��� Then
                            strType = vsTmp.Cell(flexcpData, i, TI_ICU����)
                            vsTmp.Cell(flexcpData, i, TI_ICU����) = Val(vsTmp.Cell(flexcpData, i, TI_ICU����)) - 1
                            vsTmp.TextMatrix(i, TI_ICU����) = Val(vsTmp.Cell(flexcpData, i, TI_ICU����)) & "-" & Split(vsTmp.TextMatrix(i, TI_ICU����), "-")(1)
                        End If
                    End If
                Next
                'ɾ���Ѿ�ɾ������֢�໤��е������ŵ���
                If vsTmp.Rows = 1 Then vsTmp.Rows = vsTmp.Rows + 1
                Call ChangeVSFHeight(vsFlxAddICU, True)
            End If

            If .Row >= .FixedRows Then
                .RemoveItem .Row
                Call ChangeVSFHeight(vsFlxAddICU, True)
            End If
            If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, UI_���) = i
                Next
            End If
        ElseIf intKeyCode = vbKeyReturn Then
            intKeyCode = 0
            If .TextMatrix(.Row, UI_�໤������) = "" Then
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
            Else
                LngCol = -1
                For i = .FixedCols To .Cols - 1
                    If Not .ColHidden(i) And i > .Col Then
                        LngCol = i: Exit For
                    End If
                Next
                If .Row = .Rows - 1 And LngCol = -1 Then
                    .Rows = .Rows + 1
                    Call ChangeVSFHeight(vsFlxAddICU, True)
                ElseIf LngCol = -1 Then
                    .Row = .Row + 1: .Col = UI_�໤������
                    Call ChangeVSFHeight(vsFlxAddICU, True)
                Else
                    .Col = LngCol
                End If
            End If
        End If
    End With
End Sub

Public Sub FlxAddICUKeyDownEdit(ByRef vsFlxAddICU As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyCode As Integer, ByVal intShift As Integer)
'vsFlxAddICU_KeyDownEdit�¼�

    Dim strKEY As String
    If vsFlxAddICU.Editable = flexEDNone Then Exit Sub
    If intKeyCode <> vbKeyReturn Then Exit Sub

    With vsFlxAddICU
        Select Case LngCol
            Case UI_�໤������
                strKEY = Trim(.EditText)
                strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
                strKEY = Replace(strKEY, Chr(10), "")
                If strKEY = "" Then Exit Sub
            Case UI_����ʱ��, UI_�˳�ʱ��
                If .TextMatrix(.Row, UI_�໤������) = "" Then Exit Sub
        End Select
        '�ƶ����,ֻ�б�׼�桢�Ĵ��������֢�໤��¼
        If LngCol = IIf(gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼, UI_����סԭ��, UI_�˳�ʱ��) Then
            If .Row = .Rows - 1 Then .Rows = .Rows + 1: Call ChangeVSFHeight(vsFlxAddICU, True)
            .ShowCell .Row + 1, UI_�໤������
        Else
            .ShowCell .Row, .Col + 1
        End If
    End With
End Sub

Public Sub FlxAddICUStartEdit(ByRef vsFlxAddICU As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, blnCancel As Boolean)
'vsFlxAddICU_StartEdit
    With vsFlxAddICU
        Select Case LngCol
            Case UI_����ʱ��
                blnCancel = .TextMatrix(LngRow, UI_�໤������) = ""
            Case UI_�˳�ʱ��
                blnCancel = Not IsDate(.TextMatrix(LngRow, UI_����ʱ��))
            Case UI_����ס�ƻ�, UI_����סԭ��
                blnCancel = Not IsDate(.TextMatrix(LngRow, UI_�˳�ʱ��))
        End Select
    End With
End Sub

Public Sub FlxAddICUValidateEdit(ByRef vsFlxAddICU As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
    Dim strSql As String, rsTmp As Recordset, vPoint As POINTAPI, blnCancelTmp As Boolean
    Dim strInput As String
    Dim strKEY As String

    With vsFlxAddICU
        Select Case LngCol
            Case UI_�໤������
                If gclsPros.MedPageSandard = ST_��������׼ Then
                    strInput = UCase(.EditText)
                    If strInput = "" Then Exit Sub

                    strSql = " Select Distinct A.ID,A.����,A.����" & _
                            " From ���ű� A,��������˵�� B" & _
                            " Where B.����ID=A.ID And B.��������='ICU'" & _
                            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                            " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                            " Order by A.����"
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "��֢�໤��", _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancelTmp, False, True, _
                        strInput & "%", gclsPros.LikeString & strInput & "%")
                    If rsTmp Is Nothing Then
                        If Not blnCancelTmp Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                            MsgBox "û������ICU��֢�໤�ҡ�", vbInformation, gstrSysName
                        End If
                    Else
                        .TextMatrix(LngRow, LngCol) = rsTmp!���� & ""
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                    End If
                End If
            Case UI_����ʱ��, UI_�˳�ʱ��
                If .TextMatrix(.Row, UI_�໤������) = "" Then Exit Sub
                If CheckInPutIsDate(vsFlxAddICU, LngRow, LngCol) = False Then
                    blnCancel = True
                    Exit Sub
                End If
                strKEY = Trim(.EditText)
                strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
                strKEY = Replace(strKEY, Chr(10), "")
                strKEY = zlStr.FullDate(strKEY, , gclsPros.InTime, gclsPros.OutTime)
                If strKEY <> "" And strKEY <> "-  -     :" And strKEY <> "____-__-__ __:__" And InStr("0123456789", Mid(strKEY, 1, 1)) > 0 Then
                    If Not CheckDateRange(strKEY) Then
                        MsgBox "�������ʱ������ڲ��˵�סԺ�ڼ䡣", vbInformation, gstrSysName
                        blnCancel = True
                        Exit Sub
                    End If
                End If
            Case UI_����סԭ��
                If LenB(StrConv(.EditText, vbFromUnicode)) > 100 Then
                    MsgBox "���ܳ���50�����ֻ�100���ַ��ĳ��ȡ�", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
        End Select
    End With
End Sub

'vsICUInstruments�¼�
Public Sub vsICUInstrumentsAfterEdit(ByRef vsICUInstruments As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
    Dim strInput As String

    If LngCol = TI_��ʼʱ�� Or LngCol = TI_����ʱ�� Then
        strInput = zlStr.FullDate(vsICUInstruments.TextMatrix(LngRow, LngCol), , gclsPros.InTime, gclsPros.OutTime)
        If IsDate(strInput) Then
            vsICUInstruments.TextMatrix(LngRow, LngCol) = strInput
        End If
    End If
End Sub

Public Sub vsICUInstrumentsAfterRowColChange(ByRef vsICUInstruments As VSFlexGrid, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'vsICUInstruments_AfterRowColChange
    Dim vsFlxAddICU As VSFlexGrid
    Dim i As Long, strList As String
    Set vsFlxAddICU = gclsPros.CurrentForm.vsFlxAddICU
    With vsFlxAddICU
        If NewCol = TI_ICU���� Then
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, UI_���) = i
            Next
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, UI_�໤������) <> "" Then
                    strList = strList & "|" & .TextMatrix(i, UI_���) & "-" & .TextMatrix(i, UI_�໤������)
                End If
            Next
            strList = Mid(strList, 2)
            vsICUInstruments.ColComboList(TI_ICU����) = strList
            vsICUInstruments.Editable = IIf(strList <> "", flexEDKbdMouse, flexEDNone)
        End If
    End With
End Sub

Public Sub vsICUInstrumentsKeyDown(ByRef vsICUInstruments As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsICUInstruments_KeyDown
    Dim LngCol As Long, i As Long
    If vsICUInstruments.Editable = flexEDNone Then
        If intKeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: mblnReturn = True
        Exit Sub
    End If
    If intKeyCode = vbKeyDelete Then
        If MsgBox("���Ƿ����Ҫɾ������������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        With vsICUInstruments
            If .Row = .Rows - 1 Then
                .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
            Else
                .RemoveItem .Row
                Call ChangeVSFHeight(vsICUInstruments, True)
            End If
        End With
    ElseIf intKeyCode = vbKeyReturn Then
        intKeyCode = 0
        With vsICUInstruments
            If .TextMatrix(.Row, TI_ICU����) = "" Then
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
            Else
                LngCol = -1
                For i = .FixedCols To .Cols - 1
                    If Not .ColHidden(i) And i > .Col Then
                        LngCol = i: Exit For
                    End If
                Next
                If .Row = .Rows - 1 And LngCol = -1 Then
                    .Rows = .Rows + 1
                    Call ChangeVSFHeight(vsICUInstruments, True)
                ElseIf LngCol = -1 Then
                    .Row = .Row + 1: .Col = TI_ICU����
                Else
                    .Col = LngCol
                End If
            End If
        End With
    End If
End Sub

Public Sub vsICUInstrumentsKeyDownEdit(ByRef vsICUInstruments As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyCode As Integer, ByVal intShift As Integer)
'vsICUInstruments_KeyDownEdit�¼�
    Dim strKEY As String
    If vsICUInstruments.Editable = flexEDNone Then Exit Sub
    If intKeyCode = vbKeyReturn Then Exit Sub
    With vsICUInstruments
        Select Case LngCol
            Case TI_ICU����
                strKEY = Trim(.EditText)
                strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
                strKEY = Replace(strKEY, Chr(10), "")
                If strKEY = "" Then Exit Sub
            Case TI_��ʼʱ��, TI_����ʱ��
                If .TextMatrix(.Row, TI_��е������) = "" Then Exit Sub
                If CheckInPutIsDate(gclsPros.CurrentForm.vsFlxAddICU, LngRow, LngCol) = False Then
                    intKeyCode = 0
                    zlCommFun.PressKey vbKeySpace
                    .EditSelStart = 1
                    .EditSelLength = 1000
                    Exit Sub
                End If
        End Select
        '�ƶ����,�Ĵ��������֢�໤��е��¼
        If LngCol = TI_��Ⱦ�ۼ�Сʱ Then
            If .Row = .Rows - 1 Then .Rows = .Rows + 1: Call ChangeVSFHeight(vsICUInstruments, True)
            .ShowCell .Row + 1, TI_ICU����
        Else
            .ShowCell .Row, .Col + 1
        End If
    End With
End Sub

Public Sub vsICUInstrumentsStartEdit(ByRef vsICUInstruments As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsICUInstruments_StartEdit�¼�
    With vsICUInstruments
        Select Case LngCol
            Case TI_��е������
                blnCancel = .TextMatrix(LngRow, TI_ICU����) = ""
            Case TI_��ʼʱ��
                blnCancel = .TextMatrix(LngRow, TI_��е������) = ""
            Case TI_����ʱ��
                blnCancel = Not IsDate(.TextMatrix(LngRow, TI_��ʼʱ��))
            Case TI_��Ⱦ�ۼ�Сʱ
                blnCancel = Not IsDate(.TextMatrix(LngRow, TI_����ʱ��))
        End Select
    End With
End Sub

Public Sub vsICUInstrumentsValidateEdit(ByRef vsICUInstruments As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsICUInstruments_ValidateEdit
    Dim strKEY As String
    Dim str��סʱ�� As String, strת��ʱ�� As String
    Dim strTmp As String, i As Long
    Dim lngDif As Long

    With vsICUInstruments
        strKEY = Trim(.EditText)
        strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
        strKEY = Replace(strKEY, Chr(10), "")
        If strKEY = "" Then strKEY = .TextMatrix(LngRow, LngCol)

        Select Case LngCol
            Case TI_ICU����
                If .EditText <> "" Then
                    .Cell(flexcpData, LngRow, TI_ICU����) = Mid(.EditText, 1, InStr(.EditText, "-") - 1)
                    .RowData(LngRow) = Val(.Cell(flexcpData, LngRow, TI_ICU����))
                End If
            Case TI_��е������
                 i = Val(.Cell(flexcpData, LngRow, TI_ICU����))
                str��סʱ�� = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(i, UI_����ʱ��))
                strת��ʱ�� = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(i, UI_�˳�ʱ��))
                'û�п�ʼʱ�䣬��Ϊ����ʱ��+1��
                If .TextMatrix(LngRow, TI_��ʼʱ��) = "" And str��סʱ�� <> "" Then .TextMatrix(LngRow, TI_��ʼʱ��) = Format(CDate(str��סʱ��) + 1 / 24 / 60, "yyyy-mm-dd hh:mm")
                'û�п�ʼʱ�䣬��Ϊ�˳�ʱ��-1��
                If .TextMatrix(LngRow, TI_����ʱ��) = "" And strת��ʱ�� <> "" Then .TextMatrix(LngRow, TI_����ʱ��) = Format(CDate(strת��ʱ��) - 1 / 24 / 60, "yyyy-mm-dd hh:mm")
                 If .TextMatrix(LngRow, TI_����ʱ��) <> "" And .TextMatrix(LngRow, TI_��ʼʱ��) <> "" Then
                    lngDif = DateDiff("n", CDate(.TextMatrix(LngRow, TI_��ʼʱ��)), CDate(.TextMatrix(LngRow, TI_����ʱ��)))
                    .TextMatrix(LngRow, TI_��Ⱦ�ۼ�Сʱ) = Format(lngDif \ 60, "00") & ":" & Format(lngDif Mod 60, "00")
                 End If
            Case TI_��ʼʱ��, TI_����ʱ��
                i = Val(.Cell(flexcpData, LngRow, TI_ICU����))
                strTmp = IIf(LngCol = TI_��ʼʱ��, "��ʼʹ��ʱ��", "����ʹ��ʱ��")
                str��סʱ�� = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(i, UI_����ʱ��))
                strת��ʱ�� = Trim(gclsPros.CurrentForm.vsFlxAddICU.TextMatrix(i, UI_�˳�ʱ��))
                strKEY = zlStr.FullDate(strKEY, , gclsPros.InTime, gclsPros.OutTime)
                If strKEY = "" Then Exit Sub
                If Not IsDate(strKEY) Then
                    MsgBox strTmp & "����Ϊ������,���������룡", vbInformation + vbDefaultButton1, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
                If IsDate(str��סʱ��) Then
                    If CDate(strKEY) < CDate(str��סʱ��) Then
                        .EditText = str��סʱ��
                        ShowMessage vsICUInstruments, "ע:" & vbCrLf & "  " & strTmp & "С������סʱ��,���飡"
                        blnCancel = True
                        Exit Sub
                    End If
                End If
                If IsDate(strת��ʱ��) Then
                    If CDate(strKEY) > CDate(strת��ʱ��) Then
                        .EditText = str��סʱ��
                        ShowMessage vsICUInstruments, "ע:" & vbCrLf & "    " & strTmp & "������ת��ʱ��,���飡"
                        blnCancel = True
                        Exit Sub
                    End If
                End If
                strTmp = .TextMatrix(LngRow, IIf(LngCol = TI_��ʼʱ��, TI_����ʱ��, TI_��ʼʱ��))
                If IsDate(strTmp) Then
                    If CDate(strKEY) >= CDate(strTmp) And LngCol = TI_��ʼʱ�� Then
                        ShowMessage vsICUInstruments, "������Ŀ�ʼʹ��ʱ����ڽ���ʹ��ʱ�䣬���顣"
                        blnCancel = True
                        Exit Sub
                    ElseIf CDate(strKEY) <= CDate(strTmp) And LngCol = TI_����ʱ�� Then
                        ShowMessage vsICUInstruments, "������Ľ���ʹ��ʱ��С�ڿ�ʼʹ��ʱ�䣬���顣"
                        blnCancel = True
                        Exit Sub
                    End If
                    If .TextMatrix(LngRow, TI_����ʱ��) <> "" And .TextMatrix(LngRow, TI_��ʼʱ��) <> "" Then
                        lngDif = DateDiff("n", CDate(.TextMatrix(LngRow, TI_��ʼʱ��)), CDate(.TextMatrix(LngRow, TI_����ʱ��)))
                        .TextMatrix(LngRow, TI_��Ⱦ�ۼ�Сʱ) = Format(lngDif \ 60, "00") & ":" & Format(lngDif Mod 60, "00")
                    End If
                End If
                If Not CheckDateRange(strKEY) Then
                    ShowMessage vsICUInstruments, "�������ʱ������ڲ��˵�סԺ�ڼ䡣"
                    blnCancel = True
                    Exit Sub
                End If
        Case TI_��Ⱦ�ۼ�Сʱ
            If InStr(strKEY, ":") > 0 Then
                If Val(Mid(strKEY, InStr(strKEY, ":") + 1)) >= 60 Then
                    ShowMessage vsICUInstruments, "����ķ��������ܳ���59���ӡ�"
                    blnCancel = True
                    Exit Sub
                End If
            End If
        End Select
    End With
End Sub

'vsInfect�¼�
Public Sub vsInfectAfterEdit(ByRef vsInfect As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsInfect_AfterEdit�¼�
    Dim strInput As String
    If LngCol = FI_ȷ������ Then
        strInput = zlStr.FullDate(vsInfect.TextMatrix(LngRow, LngCol), False, gclsPros.InTime, gclsPros.OutTime)
        If IsDate(strInput) Then
            vsInfect.TextMatrix(LngRow, LngCol) = strInput
        End If
    End If
    Call vsInfectAfterRowColChange(vsInfect, -1, -1, vsInfect.Row, vsInfect.Col)
End Sub

Public Sub vsInfectAfterRowColChange(ByRef vsInfect As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsInfect_AfterRowColChange�¼�
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    vsInfect.ColComboList(FI_ҽԺ��Ⱦ����) = "..."
End Sub

Public Sub vsInfectCellButtonClick(ByRef vsInfect As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsInfect_CellButtonClick�¼�
    Dim blnCancle As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim textTmp As TextBox
    Set rsTmp = zlDatabase.CopyNewRec(GetBaseCode("ҽԺ��ȾĿ¼"), , "ID,����,����,����")
    If rsTmp.RecordCount = 0 Then
        MsgBox "û�и�Ⱦ��Ŀ����ѡ��,�뵽�ֵ�����������ø�Ⱦ��Ŀ��", vbInformation, gstrSysName
    Else
        Set textTmp = GetReplaceObject(vsInfect)
        If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , , rsTmp) Then
            With vsInfect
                .TextMatrix(LngRow, FI_ҽԺ��Ⱦ����) = rsTmp!���� & ""
                .TextMatrix(LngRow, LngCol) = rsTmp!���� & ""
                If LngRow = .Rows - 1 Then
                    .Rows = .Rows + 1
                    Call ChangeVSFHeight(vsInfect, True)
                End If
                .ShowCell .Row + 1, FI_ȷ������
            End With
        End If
    End If
End Sub

Public Sub vsInfectKeyDown(ByRef vsInfect As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsInfect_KeyDown�¼�
    If vsInfect.Editable = flexEDNone Then Exit Sub
    If intKeyCode = vbKeyDelete Then
            If MsgBox("���Ƿ����Ҫɾ������������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            With vsInfect
                If .Rows = .FixedRows Then
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                Else
                    .RemoveItem .Row
                     Call ChangeVSFHeight(vsInfect, True)
                End If
            End With
    ElseIf intKeyCode = vbKeyReturn And Trim(vsInfect.TextMatrix(vsInfect.Row, FI_ȷ������)) = "" Then
        zlCommFun.PressKey vbKeyTab: mblnReturn = True
    ElseIf intKeyCode = Asc("*") Then
        Call vsInfectCellButtonClick(vsInfect, vsInfect.Row, vsInfect.Col)
    Else
         vsInfect.ColComboList(FI_ҽԺ��Ⱦ����) = ""  'ʹ��ť״̬��������״̬
    End If
    Call VsGriedFocuesMove(vsInfect, vsInfect.Row, vsInfect.Col, intKeyCode)
End Sub

Public Sub vsInfectKeyPress(ByRef vsInfect As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsInfect_KeyPress�¼�
    If vsInfect.Editable = flexEDNone Then Exit Sub
    If intKeyAscii = vbKeyReturn Then intKeyAscii = 0
End Sub

Public Sub vsInfectKeyPressEdit(ByRef vsInfect As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsInfect_KeyPressEdit�¼�

    If intKeyAscii = Asc("'") Then intKeyAscii = 0: Exit Sub
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call VsGriedFocuesMove(vsInfect, LngRow, LngCol, vbKeyReturn)
        Exit Sub
    End If
    Select Case LngCol
        Case FI_ȷ������, FI_��Ⱦ��λ, FI_ҽԺ��Ⱦ����
            Call VsFlxGridCheckKeyPress(vsInfect, LngRow, LngCol, intKeyAscii, m�ı�ʽ)
        Case Else
    End Select
End Sub

Public Sub vsInfectStartEdit(ByRef vsInfect As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsInfect_StartEdit�¼�
    Select Case LngCol
        Case FI_��Ⱦ��λ
            blnCancel = Not IsDate(vsInfect.TextMatrix(LngRow, FI_ȷ������))
        Case FI_ҽԺ��Ⱦ����
            blnCancel = vsInfect.TextMatrix(LngRow, FI_��Ⱦ��λ) = ""
    End Select
End Sub

Public Sub vsInfectValidateEdit(ByRef vsInfect As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsInfect_ValidateEdit�¼�
    Dim strKEY As String
    Dim strFilter As String
    Dim rsTmp As ADODB.Recordset
    Dim blnInputCancel As Boolean
    Dim textTmp As TextBox
    With vsInfect
        strKEY = Trim(vsInfect.EditText)
        strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
        strKEY = Replace(strKEY, Chr(10), "")
        Select Case LngCol
            Case FI_ȷ������
                If strKEY <> "" Then
                    strKEY = zlStr.FullDate(strKEY, False, gclsPros.InTime, gclsPros.OutTime)
                    If Not IsDate(strKEY) Then
                        MsgBox "ȷ�����ڱ���Ϊ������,���������룡", vbInformation + vbDefaultButton1, gstrSysName
                        blnCancel = True
                        zlCommFun.PressKey vbKeySpace
                        .EditSelStart = 0
                        .EditSelLength = 1000
                        Exit Sub
                    End If
                    If Not CheckDateRange(strKEY, True) Then
                        MsgBox "�������ʱ������ڲ��˵�סԺ�ڼ䡣", vbInformation, gstrSysName
                        blnCancel = True
                        Exit Sub
                    End If
                End If
            Case FI_��Ⱦ��λ
                 Call VsGriedFocuesMove(vsInfect, LngRow, LngCol, vbKeyReturn)
            Case FI_ҽԺ��Ⱦ����
                If strKEY = "" Then Exit Sub
                Set rsTmp = zlDatabase.CopyNewRec(GetBaseCode("ҽԺ��ȾĿ¼"), , "ID,����,����,����")
                If rsTmp.RecordCount = 0 Then
                    MsgBox "û�и�Ⱦ��Ŀ����ѡ��", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                Else
                    strKEY = UCase$(strKEY)
                    strFilter = "���� Like '" & strKEY & "*' OR ����  Like '" & strKEY & "*'  OR ���� like '" & IIf(gclsPros.LikeString <> "", "*", "") & strKEY & "*' "
                    Set textTmp = GetReplaceObject(vsInfect)
                    blnInputCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, Rec.FilterNew(rsTmp, strFilter), True, , , rsTmp)
                    If rsTmp Is Nothing Then
                        If Not blnInputCancel Then
                            MsgBox "δ�ҵ�ƥ��ĸ�Ⱦ��Ŀ��", vbInformation, gstrSysName
                            blnCancel = True
                            Exit Sub
                        Else
                            blnCancel = True
                            Exit Sub
                        End If
                    Else
                        .TextMatrix(LngRow, FI_ҽԺ��Ⱦ����) = rsTmp!���� & ""
                        .EditText = rsTmp!���� & ""
                        vsInfect.SetFocus
                        Call VsGriedFocuesMove(vsInfect, LngRow, LngCol, vbKeyReturn)
                    End If
                End If
        End Select
    End With
End Sub

'vsSample�¼�
Public Sub vsSampleAfterEdit(ByRef vsSample As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsSample_AfterEdit�¼�
    Dim strInput As String

    If LngCol = MI_�ͼ����� Then
        strInput = zlStr.FullDate(vsSample.TextMatrix(LngRow, LngCol), False, gclsPros.InTime, gclsPros.OutTime)
        If IsDate(strInput) Then
            vsSample.TextMatrix(LngRow, LngCol) = strInput
        End If
    End If
End Sub

Public Sub vsSampleKeyDown(ByRef vsSample As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsSample_KeyDown�¼�
    Dim LngCol As Long, i As Long
    With vsSample
        If vsSample.Editable = flexEDNone Then Exit Sub
        If intKeyCode = vbKeyDelete Then
                If MsgBox("���Ƿ����Ҫɾ������������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                If .Rows = .FixedRows Then
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                Else
                    .RemoveItem .Row
                     Call ChangeVSFHeight(vsSample, True)
                End If
        ElseIf intKeyCode = vbKeyReturn Then
            If .TextMatrix(vsSample.Row, MI_�걾) = "" Then
                zlCommFun.PressKey vbKeyTab: mblnReturn = True
            Else
                LngCol = -1
                For i = .FixedCols To .Cols - 1
                    If Not .ColHidden(i) And i > .Col Then
                        LngCol = i: Exit For
                    End If
                Next
                If .Row = .Rows - 1 And LngCol = -1 Then
                    .Rows = .Rows + 1
                     Call ChangeVSFHeight(vsSample, True)
                ElseIf LngCol = -1 Then
                    .Row = .Row + 1: .Col = MI_�걾
                     Call ChangeVSFHeight(vsSample, True)
                Else
                    .Col = LngCol
                End If
            End If
        End If
    End With
End Sub

Public Sub vsSampleStartEdit(ByRef vsSample As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsSample_StartEdit�¼�
    Select Case LngCol
        Case MI_��ԭѧ���뼰����
            blnCancel = vsSample.TextMatrix(LngRow, MI_�걾) = ""
        Case MI_�ͼ�����
            blnCancel = vsSample.TextMatrix(LngRow, MI_��ԭѧ���뼰����) = ""
    End Select
End Sub

Public Sub vsSampleValidateEdit(ByRef vsSample As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsSample_ValidateEdit�¼�
    Dim strKEY As String
    Dim strFilter As String
    Dim rsTmp As ADODB.Recordset

    With vsSample
        strKEY = Trim(vsSample.EditText)
        strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
        strKEY = Replace(strKEY, Chr(10), "")
        Select Case LngCol
            Case MI_�ͼ�����
                strKEY = zlStr.FullDate(strKEY, False, gclsPros.InTime, gclsPros.OutTime)
                If Not IsDate(strKEY) Then
                    MsgBox "ȷ�����ڱ���Ϊ������,���������룡", vbInformation + vbDefaultButton1, gstrSysName
                    blnCancel = True
                    zlCommFun.PressKey vbKeySpace
                    .EditSelStart = 0
                    .EditSelLength = 1000
                    Exit Sub
                End If
                If Not CheckDateRange(strKEY, True) Then
                    MsgBox "�������ʱ������ڲ��˵�סԺ�ڼ䡣", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
        End Select
    End With
End Sub

'vsTSJC�¼�
Public Sub TSJCAfterEdit(ByRef vsTSJC As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsTSJC_AfterEdit�¼�
    Call TSJCAfterRowColChange(vsTSJC, -1, -1, vsTSJC.Row, vsTSJC.Col)
End Sub

Public Sub TSJCAfterRowColChange(ByRef vsTSJC As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsTSJC_AfterRowColChange�¼�
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    vsTSJC.ComboList = "..."
End Sub

Public Sub TSJCCellButtonClick(ByRef vsTSJC As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsTSJC_CellButtonClick�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strSQLItem As String

    With vsTSJC
        strSQLItem = _
            " From ������ĿĿ¼ A" & _
            " Where A.���='D' And A.������� IN(2,3) And A.����Ӧ��=1" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
        strSql = "Select 0 as ĩ��,Max(Level) as ��ID,ID,�ϼ�ID,����,����,NULL as ��λ" & _
            " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select A.����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID" & _
            " Group by ID,�ϼ�ID,����,����"
        strSql = strSql & " Union ALL" & _
            " Select 1 as ĩ��,1 as ��ID,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ" & _
            strSQLItem & " Order By ĩ��,��ID Desc,����"
        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 2, "������", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�м����Ŀ���ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call TSJCSetDiagInput(LngRow, rsTmp)
            Call TSJCEnterNextCell
        End If
    End With
End Sub

Public Sub TSJCKeyDown(ByRef vsTSJC As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsTSJC_KeyDown�¼�
    'If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    With vsTSJC
        If intKeyCode = vbKeyF4 Then
            Call zlCommFun.PressKey(vbKeySpace)
        ElseIf intKeyCode = vbKeyDelete Then
            If MsgBox("ȷʵҪɾ������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                .TextMatrix(.Row, 1) = ""
            End If
        ElseIf intKeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call TSJCKeyPress(vsTSJC, intKeyCode)
        End If
    End With
End Sub

Public Sub TSJCKeyPress(ByRef vsTSJC As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsTSJC_KeyPress�¼�
    If vsTSJC.Editable = flexEDNone Then Exit Sub
    With vsTSJC
        If intKeyAscii = 13 Then
            intKeyAscii = 0
            Call TSJCEnterNextCell
        ElseIf gclsPros.MedPageSandard <> ST_�Ĵ�ʡ��׼ Then
            If intKeyAscii = Asc("*") Then
                intKeyAscii = 0
                Call TSJCCellButtonClick(vsTSJC, .Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Public Sub TSJCKeyPressEdit(ByRef vsTSJC As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsTSJC_KeyPressEdit�¼�
    If intKeyAscii = vbKeyReturn Then
        gclsPros.IsReturn = True
    Else
        gclsPros.IsReturn = False
    End If
End Sub

Public Sub TSJCSetupEditWindow(ByRef vsTSJC As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef lngEditWindow As Long, ByRef blnIsCombo As Boolean)
'vsTSJC_SetupEditWindow�¼�
    With vsTSJC
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Public Sub TSJCValidateEdit(ByRef vsTSJC As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsTSJC_ValidateEdit�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI

    With vsTSJC
        If .EditText = "" Then
            .EditText = .Cell(flexcpData, LngRow, LngCol)
            If gclsPros.IsReturn Then Call TSJCEnterNextCell
        ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
            If gclsPros.IsReturn Then Call TSJCEnterNextCell
        Else
            strInput = UCase(.EditText)
            If LenB(StrConv(strInput, vbFromUnicode)) > 100 Then
                MsgBox "����������ݲ��ܳ���50�����֡�", vbInformation, gstrSysName
                blnCancel = True
                Exit Sub
            End If
            If zlCommFun.IsCharChinese(strInput) Then
                strSql = "B.���� Like [2]" '���뺺��ʱֻƥ������
            Else
                strSql = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
            End If
            strSql = _
                " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ" & _
                " From ������ĿĿ¼ A,������Ŀ���� B" & _
                " Where A.ID=B.������ĿID And A.���='D' And A.������� IN(2,3)" & _
                " And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And A.����Ӧ��=1 And B.����=[3] And (" & strSql & ")" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.����"
            If zlCommFun.IsCharChinese(strInput) Then
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput & "%", gclsPros.LikeString & strInput & "%", gclsPros.BriefCode + 1)
                If rsTmp.EOF Then
                    Set rsTmp = Nothing
                ElseIf rsTmp.RecordCount > 1 Then
                    Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                End If
                Call TSJCSetDiagInput(LngRow, rsTmp)
                .EditText = .Text
                If gclsPros.IsReturn Then Call TSJCEnterNextCell
            Else
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "������", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                    strInput & "%", gclsPros.LikeString & strInput & "%", gclsPros.BriefCode + 1)
                If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    blnCancel = True
                Else
                    Call TSJCSetDiagInput(LngRow, rsTmp)
                    .EditText = .Text
                    If gclsPros.IsReturn Then Call TSJCEnterNextCell
                End If
            End If
        End If
        gclsPros.IsReturn = False
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'vsfMain�¼�
Public Sub vsfMainEnterCell(ByRef vsfMain As VSFlexGrid)
'vsfMain_EnterCell�¼�
    With vsfMain
        Select Case .Col
            Case 1, 4, 7
                If InStr(.TextMatrix(.Row, .Col + 1), ",") > 0 Then
                    .ColComboList(.Col) = Replace(.TextMatrix(.Row, .Col + 1), ",", "|")
                Else
                    .ColComboList(.Col) = ""
                End If
        End Select
    End With
End Sub

Public Sub vsfMainKeyPress(ByRef vsfMain As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsfMain_KeyPress�¼�
    With vsfMain
        If .Editable = flexEDNone Or .Rows <= 1 Then zlCommFun.PressKey (vbKeyTab): Exit Sub
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            Select Case .Col
                Case 0, 3, 6
                    .Col = .Col + 1
                Case 1, 4, 7
                    If .Col = .Cols - 2 Then
                        If .Row <> .Rows - 1 Then
                            .Col = 1
                            .Row = .Row + 1
                        Else
                            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
                        End If
                    Else
                        .Col = .Col + 3
                    End If
            End Select
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Public Sub vsfMainStartEdit(ByRef vsfMain As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsfMain_StartEdit�¼�
    Select Case LngCol Mod 3
        Case 0
            blnCancel = True
        Case 1
            blnCancel = vsfMain.TextMatrix(LngRow, LngCol - 1) = ""
    End Select
End Sub

Public Sub vsfMainValidateEdit(ByRef vsfMain As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsfMain_ValidateEdit�¼�
    Dim sngNum1, sngNum2 As Single

    With vsfMain
        If InStr(.TextMatrix(LngRow, LngCol + 1), "...") > 0 Then
            sngNum1 = Mid(.TextMatrix(LngRow, LngCol + 1), 1, InStr(.TextMatrix(LngRow, LngCol + 1), "...") - 1)
            sngNum2 = Mid(.TextMatrix(LngRow, LngCol + 1), InStr(.TextMatrix(LngRow, LngCol + 1), "...") + 3)
            If Not IsNumeric(.EditText) Then
                blnCancel = True
            ElseIf CSng(.EditText) < sngNum1 Or CSng(.EditText) > sngNum2 Then
                MsgBox "����Ӧ����" & .TextMatrix(LngRow, LngCol + 1) & "�ķ�Χ����!", vbInformation, gstrSysName
                blnCancel = True
            End If
        ElseIf InStr(.TextMatrix(LngRow, LngCol + 1), "-") > 0 Then
            If InStr(.TextMatrix(LngRow, LngCol + 1), "-") = 1 Then
                sngNum1 = Mid(.TextMatrix(LngRow, LngCol + 1), 2, InStr(2, .TextMatrix(LngRow, LngCol + 1), "-") - 1)
                sngNum2 = Mid(.TextMatrix(LngRow, LngCol + 1), InStr(2, .TextMatrix(LngRow, LngCol + 1), "-") + 1)
            Else
                sngNum1 = Mid(.TextMatrix(LngRow, LngCol + 1), 1, InStr(1, .TextMatrix(LngRow, LngCol + 1), "-") - 1)
                sngNum2 = Mid(.TextMatrix(LngRow, LngCol + 1), InStr(1, .TextMatrix(LngRow, LngCol + 1), "-") + 1)
            End If
            If Not IsNumeric(.EditText) Then
                blnCancel = True
            ElseIf CSng(.EditText) < sngNum1 Or CSng(.EditText) > sngNum2 Then
                MsgBox "����Ӧ����" & .TextMatrix(LngRow, LngCol + 1) & "�ķ�Χ����!", vbInformation, gstrSysName
                blnCancel = True
            End If
        ElseIf .TextMatrix(LngRow, LngCol + 1) = "" Then
            If zlCommFun.ActualLen(.EditText) > gclsPros.ValueLen Then
                MsgBox "���볤�Ȳ��ܴ���" & "[" & gclsPros.ValueLen & "]", vbInformation, gstrSysName
                blnCancel = True
            End If
        End If
    End With
End Sub

'vsFrees�¼�
Public Sub vsFeesComboDropDown(ByVal LngRow As Long, ByVal LngCol As Long)
'vsFees_ComboDropDown�¼�
    Dim vsFees As VSFlexGrid
    Dim i As Long

    Set vsFees = gclsPros.CurrentForm.vsFees
    With vsFees
        If LngCol Mod 2 = 0 Then
            '��λ��ƥ����
            If .TextMatrix(LngRow, LngCol) <> "" Then
                For i = 0 To .ComboCount - 1
                    If zlStr.NeedName(.ComboItem(i)) = .TextMatrix(LngRow, LngCol) Then
                        .ComboIndex = i: Exit For
                    End If
                Next
            End If
        End If
    End With
End Sub

Public Sub vsFeesKeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsFees_KeyDown�¼�
    Dim vsFree As VSFlexGrid

    Set vsFree = gclsPros.CurrentForm.vsFees
    With vsFree
        If intKeyCode = vbKeyDelete And .Editable <> flexEDNone Then
            If Not FreeHaveLowLevel(.Row, IIf(.Col Mod 2 = 0, .Col, .Col - 1)) Then
                If MsgBox("�Ƿ�ɾ���÷��ã�", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                    Call AddOrDelFreeCols(vsFree, .TextMatrix(.Row, IIf(.Col Mod 2 = 0, .Col, .Col - 1)), .TextMatrix(.Row, IIf(.Col Mod 2 = 0, .Col + 1, .Col)), False)
                End If
            End If

        ElseIf intKeyCode = vbKeyReturn Then
            intKeyCode = 0
            If .TextMatrix(.Row, IIf(.Col Mod 2 = 0, .Col, .Col - 1)) = "" Or .Editable = flexEDNone Then
                If gclsPros.CurrentForm.cboBaseInfo(BCC_��������).Enabled And Not gclsPros.CurrentForm.cboBaseInfo(BCC_��������).Locked Then
                    Call gclsPros.CurrentForm.cboBaseInfo(BCC_��������).SetFocus
                End If
            Else
                If IIf(.Col Mod 2 = 0, .Col, .Col - 1) = 4 Then
                    .Col = 0: .Row = .Row + 1
                Else
                    .Col = IIf(.Col Mod 2 = 0, .Col, .Col - 1) + 2
                End If
            End If
        End If
    End With
End Sub

Public Sub vsFeesKeyPressEdit(ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsFees_KeyPressEdit�¼�
    Dim vsFees As VSFlexGrid

    Set vsFees = gclsPros.CurrentForm.vsFees
    If vsFees.Editable = flexEDNone Then Exit Sub
    With vsFees
         If intKeyAscii = vbKeyReturn Then
            gclsPros.IsReturn = True
            If LngCol Mod 2 = 0 Then
                intKeyAscii = 0
                If .ComboIndex <> -1 Then
                    '��ʱ.TextMatrix��δ����,����ȡComboItem
                    .TextMatrix(LngRow, LngCol) = .ComboItem(.ComboIndex)
                    Call EnterNextCellFees(vsFees)
                End If
            End If
         Else
             If LngCol Mod 2 = 1 Then
                 If .EditSelLength <> 0 Then Exit Sub
                 If Len(.EditText) > 17 Then intKeyAscii = 0: Exit Sub
                 Call VsFlxGridCheckKeyPress(vsFees, LngRow, LngCol, intKeyAscii, m���ʽ)
             End If
             gclsPros.IsReturn = False
         End If
    End With
End Sub

Public Sub vsFeesStartEdit(ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsFees_StartEdit�¼�
    '�����Ӽ����ã�������༭
    If FreeHaveLowLevel(LngRow, IIf(LngCol Mod 2 = 0, LngCol, LngCol - 1)) Then blnCancel = True
End Sub

Public Sub vsFeesValidateEdit(ByVal LngRow As Long, ByVal LngCol As Long, byrefCancel As Boolean)
'vsFees_ValidateEdit�¼�
    Dim vsFree As VSFlexGrid
    Dim i As Long, lngTmpRow As Long, lngTmpCol As Long

    Set vsFree = gclsPros.CurrentForm.vsFees
    With vsFree
        If LngCol Mod 2 = 0 Then
            For i = .FixedRows * 3 To (.Rows - 1) * 3
                lngTmpRow = i \ 3: lngTmpCol = (i Mod 3) * 2
                If .TextMatrix(lngTmpRow, lngTmpCol) = .EditText And lngTmpRow <> LngRow And lngTmpCol <> LngCol Then
                    If gclsPros.SameName Then
                        Call AddOrDelFreeCols(vsFree, .TextMatrix(LngRow, LngCol), "", True)
                        Exit Sub
                    End If
                End If
            Next
        Else
            .TextMatrix(LngRow, LngCol) = Format(.EditText, gclsPros.FreeFormat)
            Call SumAndSetFrees
        End If
    End With
End Sub

'vsTransfer�¼�
Public Sub vsTransferAfterRowColChange(ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsTransfer_AfterRowColChange�¼�
    Dim vsTransfer As VSFlexGrid
    Dim blnEdit As Boolean
    Set vsTransfer = gclsPros.CurrentForm.vsTransfer
    With vsTransfer
        If lngNewCol >= .FixedCols Then
            If lngNewRow = DR_ת�ƿ��� Then
                blnEdit = .TextMatrix(DR_ת�ƿ���, lngNewCol - 1) <> ""
            Else
                blnEdit = .TextMatrix(DR_ת�ƿ���, lngNewCol) <> ""
            End If
            If lngNewRow = DR_ת�ƿ��� Then
                .FocusRect = IIf(blnEdit, flexFocusSolid, flexFocusLight)
                .ComboList = IIf(blnEdit, "...", "")
            Else
                .ComboList = ""
                .FocusRect = IIf(blnEdit, flexFocusSolid, flexFocusLight)
            End If
        End If
    End With
End Sub

Public Sub vsTransferCellButtonClick(ByVal LngRow As Long, ByVal LngCol As Long)
'vsTransfer_CellButtonClick�¼�
    Dim rsTmp As ADODB.Recordset
    Dim vsTransfer As VSFlexGrid
    Dim textTmp As TextBox

    Set vsTransfer = gclsPros.CurrentForm.vsTransfer

    With vsTransfer
        If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
        grsDeptInfo.Filter = "��������='�ٴ�'"
        If grsDeptInfo.RecordCount = 0 Then
            MsgBox "δ�ҵ��ٴ��������ݣ����ڻ������ݹ��������ò�������Ϊ�ٴ���", vbInformation, gstrSysName
        Else
            grsDeptInfo.Filter = "��������='�ٴ�'"
            Set textTmp = GetReplaceObject(vsTransfer)
            If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, grsDeptInfo, True, , , rsTmp) Then
                If rsTmp.RecordCount > 0 Then
                    .TextMatrix(LngRow, LngCol) = rsTmp!����
                End If
            End If
        End If
    End With
End Sub

Public Sub vsTransferKeyDown(ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsTransfer_KeyDown�¼�
    Dim vsTransfer As VSFlexGrid
    Dim i As Long

    Set vsTransfer = gclsPros.CurrentForm.vsTransfer
    If vsTransfer.Editable = flexEDNone Then Exit Sub
    With vsTransfer
        If intKeyCode = vbKeyDelete Then
            For i = .Col To .Cols - 2
                .TextMatrix(DR_ת��ʱ��, i) = .TextMatrix(DR_ת��ʱ��, i + 1)
                .TextMatrix(DR_ת�ƿ���, i) = .TextMatrix(DR_ת�ƿ���, i + 1)
            Next
            .TextMatrix(DR_ת��ʱ��, .Cols - 1) = ""
            .TextMatrix(DR_ת�ƿ���, .Cols - 1) = ""
        ElseIf intKeyCode = vbKeyInsert Then
            If .TextMatrix(0, .Col) <> "" Then
                For i = .Cols - 1 To .Col + 1 Step -1
                    .TextMatrix(DR_ת��ʱ��, i) = .TextMatrix(DR_ת��ʱ��, i - 1)
                    .TextMatrix(DR_ת�ƿ���, i) = .TextMatrix(DR_ת�ƿ���, i - 1)
                Next
                .TextMatrix(DR_ת��ʱ��, .Col) = ""
                .TextMatrix(DR_ת�ƿ���, .Col) = ""
            End If
        ElseIf intKeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsTransferKeyPress(intKeyCode)
        End If
    End With
End Sub

Public Sub vsTransferStartEdit(ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsTransfer_StartEdit�¼�
    With gclsPros.CurrentForm.vsTransfer
        If LngRow = DR_ת��ʱ�� And .TextMatrix(DR_ת�ƿ���, LngCol) = "" Then blnCancel = True
        If LngRow = DR_ת�ƿ��� And .TextMatrix(DR_ת�ƿ���, LngCol - 1) = "" Then blnCancel = True
    End With
End Sub

Public Sub vsTransferKeyPress(ByRef intKeyAscii As Integer)
'vsTransfer_KeyPress�¼�
    Dim vsTransfer As VSFlexGrid
    Dim i As Long

    Set vsTransfer = gclsPros.CurrentForm.vsTransfer
    If vsTransfer.Editable = flexEDNone Then Exit Sub
    With vsTransfer
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            If .Col = .Cols - 1 And .Row = DR_ת��ʱ�� Then
                If ControlIsLocked(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��)) Then
                    Call gclsPros.CurrentForm.txtInfo(GC_��Ժ����).SetFocus
                Else
                    Call gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��).SetFocus
                End If
            ElseIf .TextMatrix(DR_ת�ƿ���, .Col) = "" Then
                If ControlIsLocked(gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��)) Then
                    Call gclsPros.CurrentForm.txtInfo(GC_��Ժ����).SetFocus
                Else
                    Call gclsPros.CurrentForm.mskDateInfo(DC_��Ժʱ��).SetFocus
                End If
            ElseIf .Row = DR_ת��ʱ�� Then
                .Col = .Col + 1: .Row = DR_ת�ƿ���
            ElseIf .Row = DR_ת�ƿ��� Then
                .Row = DR_ת��ʱ��
            End If
        Else
            If .Row = DR_ת�ƿ��� Then
                If intKeyAscii = Asc("*") Then
                    intKeyAscii = 0
                    Call vsTransferCellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Public Sub vsTransferValidateEdit(ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsTransfer_ValidateEdit�¼�
    Dim vsTransfer As VSFlexGrid
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strInput As String
    Dim textTmp As TextBox

    Set vsTransfer = gclsPros.CurrentForm.vsTransfer

    With vsTransfer
        If .EditText = "" And .TextMatrix(.Row, .Col) <> "" Then
            If MsgBox("�Ƿ�ɾ������ת����Ϣ��", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                .EditText = .TextMatrix(.Row, .Col + 1)
                For i = .Col To .Cols - 2
                    .TextMatrix(DR_ת��ʱ��, i) = .TextMatrix(DR_ת��ʱ��, i + 1)
                    .TextMatrix(DR_ת�ƿ���, i) = .TextMatrix(DR_ת�ƿ���, i + 1)
                Next
                .TextMatrix(DR_ת��ʱ��, .Cols - 1) = ""
                .TextMatrix(DR_ת�ƿ���, .Cols - 1) = ""
            End If
        Else
            If .Row = DR_ת�ƿ��� Then
                If grsDeptInfo Is Nothing Then Set grsDeptInfo = GetDeptData
                grsDeptInfo.Filter = "��������='�ٴ�' "
                strInput = UCase(Trim(.EditText))
                If strInput = "" Then Exit Sub
                Set rsTmp = Rec.FilterNew(grsDeptInfo, "���� Like '*" & strInput & "*' OR ���� Like '" & strInput & "*' OR ���� Like '" & strInput & "*'")
                If rsTmp.EOF Then
                    blnCancel = True
                Else
                    If rsTmp.RecordCount = 1 Then
                        .EditText = rsTmp!����
                    Else
                        Set textTmp = GetReplaceObject(vsTransfer)
                        If zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , , rsTmp) Then
                            If rsTmp.RecordCount <> 0 Then
                                .EditText = rsTmp!����
                            Else
                                blnCancel = True
                            End If
                        Else
                            blnCancel = True
                        End If
                    End If
                End If
            Else
                strInput = zlStr.FullDate(.EditText, , gclsPros.InTime, gclsPros.OutTime)
                If strInput <> "" Then
                    If IsDate(strInput) Then
                        .EditText = strInput
                    Else
                        MsgBox "��������ȷ��ת��ʱ�䣬���磺""2012-12-21""��""20121221""��", vbInformation, gstrSysName
                        blnCancel = True
                    End If
                ElseIf .EditText <> "" Then
                    blnCancel = True
                End If
            End If
        End If
    End With
End Sub

'vsAller�¼�
Public Sub AllerAfterEdit(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsAller_AfterEdit�¼�
    Dim strDate As String

    With vsAller
        Select Case LngCol
            Case AI_����ʱ��
                strDate = zlStr.FullDate(.TextMatrix(LngRow, LngCol), False)
                If IsDate(strDate) Then
                    .TextMatrix(LngRow, LngCol) = strDate
                End If
        End Select
        Call AllerAfterRowColChange(vsAller, -1, -1, .Row, .Col)
    End With
End Sub

Public Sub AllerAfterRowColChange(ByRef vsAller As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsAller_AfterRowColChange�¼�
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    With vsAller
        If lngNewCol = AI_����ҩ�� Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = IIf(Trim(.TextMatrix(lngNewRow, AI_����ҩ��)) = "", flexFocusLight, flexFocusSolid)
            .ComboList = ""
        End If
    End With
End Sub

Public Sub AllerCellButtonClick(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsAller_CellButtonClick�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim int�Ա� As Integer
    Dim vPoint As POINTAPI

    With vsAller
        If gclsPros.UseTYT Then
            If Not gobjPass Is Nothing Then
                strSql = gobjPass.zlPassInputAllergy()
            End If
            If InStr(strSql, ";") > 0 Then
                Call SetAllerInput(LngRow, , strSql)
                Call AllerEnterNextCell
            End If
        Else
            If gclsPros.Sex Like "*��*" Then
                int�Ա� = 1
            ElseIf gclsPros.Sex Like "*Ů*" Then
                int�Ա� = 2
            End If
            If gclsPros.FuncType <> f������ҳ Then
                If gclsPros.CurrentForm.optAller(PC_��ҩƷĿ¼����).Value = True Then
                    strSql = _
                        " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                        " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                        " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                        " Select ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
                        " NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ��" & _
                        " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                        " Union All" & _
                        " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����,A.����," & _
                        " A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                        " From ������ĿĿ¼ A,ҩƷ���� B" & _
                        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
                        IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[1])", "") & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 2, "����ҩ��", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int�Ա�)
                Else
                    strSql = "Select Rownum As ID, ����, ����, ���� From ����Դ Order By ����"
                    vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "����Դ", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
                End If
            Else
                strSql = _
                    " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                    " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                    " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
                    " Select ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
                    " NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ��" & _
                    " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                    " Union All" & _
                    " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����,A.����," & _
                    " A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                    " From ������ĿĿ¼ A,ҩƷ���� B" & _
                    " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
                    IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[1])", "") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 2, "����ҩ��", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int�Ա�)
            End If
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    If gclsPros.CurrentForm.optAller(PC_��ҩƷĿ¼����).Value = True Then
                        MsgBox "û��ҩƷ���ݿ���ѡ��", vbInformation, gstrSysName
                    Else
                        MsgBox "û�й���Դ���ݿ���ѡ��", vbInformation, gstrSysName
                    End If
                End If
            Else
                Call SetAllerInput(LngRow, rsTmp)
                Call AllerEnterNextCell
            End If
        End If
    End With
End Sub

Public Sub AllerKeyDown(ByRef vsAller As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsAller_KeyDown �¼�
    Dim i As Long
    If vsAller.Editable = flexEDNone Then Exit Sub
    'If gbln��ʿվ Or gblnReadOnly Then Exit Sub

    With vsAller
        If intKeyCode = vbKeyF4 Then
            If .Col = 1 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf intKeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, AI_����ҩ��) <> "" Then
                If MsgBox("ȷʵҪ������й���ҩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    Call ChangeVSFHeight(vsAller, True, 0)
                End If
            End If
        ElseIf intKeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call AllerKeyPress(vsAller, intKeyCode)
        End If
    End With
End Sub

Public Sub AllerKeyPress(ByRef vsAller As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsAller_KeyPress�¼�

    If vsAller.Editable = flexEDNone Then Exit Sub
    With vsAller
        If intKeyAscii = vbKeySpace Then  'Space
            If .Col = AI_����ҩ�� And gclsPros.UseTYT Then intKeyAscii = 0: Exit Sub
        End If
        If intKeyAscii = 13 Then
            intKeyAscii = 0
            Call AllerEnterNextCell
        ElseIf .Col = AI_����ҩ�� Then
            If intKeyAscii = Asc("*") Then
                intKeyAscii = 0
                Call AllerCellButtonClick(vsAller, .Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Public Sub AllerKeyPressEdit(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsAller_KeyPressEdit �¼�
    Dim blnIsNextchr As Boolean
    Dim strChr As String

    If intKeyAscii = 13 Then
        gclsPros.IsReturn = True
    Else
        gclsPros.IsReturn = False
    End If
    With vsAller
        If LngCol = AI_������Ӧ Then
            If intKeyAscii = 13 Then .Col = .Col + 1: .ShowCell LngRow, LngCol: Exit Sub
        ElseIf LngCol = AI_����ҩ�� Then
            If intKeyAscii <> 13 Then
                If gclsPros.UseTYT Then intKeyAscii = 0
            End If
        End If
    End With
End Sub

Public Sub AllerSetupEditWindow(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByVal lngEditWindow As Long, ByVal blnIsCombo As Boolean)
'vsAller_SetupEditWindow �¼�
    With vsAller
        If LngCol = AI_����ҩ�� Or LngCol = AI_����ʱ�� Then
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
        End If
    End With
End Sub

Public Sub AllerStartEdit(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsAller_StartEdit�¼�
    If LngCol = AI_������Ӧ And Trim(vsAller.TextMatrix(LngRow, AI_����ҩ��)) = "" Then blnCancel = True
End Sub

Public Sub AllerValidateEdit(ByRef vsAller As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, blnCancel As Boolean)
'vsAller_ValidateEdit�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim int�Ա�  As Integer
    Dim curDate As Date
    Dim strDate As String

    With vsAller
        If LngCol = AI_����ҩ�� Then
            If .EditText = "" Then
                If .Cell(flexcpData, LngRow, LngCol) <> "" Then
                    If MsgBox("ȷʵҪ������й���ҩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        .RemoveItem .Row
                        If .Rows = .FixedRows Then
                            .Rows = .FixedRows + 1
                            Call ChangeVSFHeight(vsAller, True)
                        End If
                    Else
                        .EditText = .Cell(flexcpData, LngRow, LngCol)
                    End If
                End If
                If gclsPros.IsReturn Then Call AllerEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                If gclsPros.IsReturn Then Call AllerEnterNextCell
            Else
                strInput = UCase(.EditText)
                If gclsPros.Sex Like "*��*" Then
                    int�Ա� = 1
                ElseIf gclsPros.Sex Like "*Ů*" Then
                    int�Ա� = 2
                End If
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                If gclsPros.FuncType <> f������ҳ Then
                    If gclsPros.CurrentForm.optAller(PC_��ҩƷĿ¼����).Value = True Then
                        strSql = _
                            " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ," & _
                            " B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                            " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
                            " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.ID=C.������ĿID" & _
                            " And (A.���� Like [1] Or A.���� Like [2] Or C.���� Like [2] Or C.���� Like [2])" & _
                            IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[3])", "") & _
                            decode(gclsPros.BriefCode, 0, " And C.����=[4]", 1, " And C.����=[4]", "") & _
                            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                            " Order by A.����"
    
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "����ҩ��", False, "", "", False, _
                            False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                            strInput & "%", gclsPros.LikeString & strInput & "%", int�Ա�, gclsPros.BriefCode + 1)
                    Else
                        If zlCommFun.IsCharChinese(strInput) Then
                            strSql = "Select Rownum As ID, ����, ����, ���� From ����Դ Where ���� Like [1] Order By ����"
                        Else
                            If gclsPros.BriefCode = 1 Then
                                strSql = "Select Rownum As ID, ����, ����, ���� From ����Դ Where zlWbCode(����) Like [1] Order By ����"
                            Else
                                strSql = "Select Rownum As ID, ����, ����, ���� From ����Դ Where ���� Like [1] Order By ����"
                            End If
                        End If
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "����Դ", False, "", "", False, _
                            False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                            gclsPros.LikeString & UCase(strInput) & "%")
                    End If
                Else
                    strSql = _
                        " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ," & _
                        " B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                        " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
                        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.ID=C.������ĿID" & _
                        " And (A.���� Like [1] Or A.���� Like [2] Or C.���� Like [2] Or C.���� Like [2])" & _
                        IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[3])", "") & _
                        decode(gclsPros.BriefCode, 0, " And C.����=[4]", 1, " And C.����=[4]", "") & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " Order by A.����"
    
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "����ҩ��", False, "", "", False, _
                        False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gclsPros.LikeString & strInput & "%", int�Ա�, gclsPros.BriefCode + 1)
                End If
                If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    blnCancel = True
                Else
                    Call SetAllerInput(LngRow, rsTmp): .EditText = .Text
                    If gclsPros.IsReturn Then Call AllerEnterNextCell
                End If
            End If
            gclsPros.IsReturn = False
        ElseIf LngCol = AI_����ʱ�� Then
            If .EditText <> "" Then
                strDate = zlStr.FullDate(.EditText, False)
                If IsDate(strDate) Then
                    curDate = zlDatabase.Currentdate
                    If CDate(strDate) > curDate Then
                        MsgBox "����������ڲ��ܴ��ڵ�ǰʱ�䡣��ǰʱ�䣺" & Format(curDate, "yyyy-mm-dd") & "��", vbInformation, gstrSysName
                        blnCancel = True
                        .EditText = .TextMatrix(LngRow, LngCol)
                    End If
                    .EditText = Format(strDate, "yyyy-MM-dd")
                    If .Cell(flexcpData, LngRow, LngCol) <> .EditText Then
                        .Cell(flexcpData, LngRow, LngCol) = .EditText
                    End If
                Else
                    MsgBox "��������ȷ�Ĺ���ʱ�䣬���磺""2012-12-21""��""121221""��", vbInformation, gstrSysName
                    blnCancel = True
                End If
            End If
        End If
    End With
End Sub

'vsDiagXY�¼�,vsDiagZY�¼�
Public Sub DiagAfterEdit(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsDiagXY_AfterEdit�¼�,vsDiagZY_AfterEdit�¼�
    Dim bln��ҽ As Boolean
    Dim i As Long, lngStart As Long, lngEnd As Long

    With vsDiag
        bln��ҽ = .Name = "vsDiagXY"
        If LngCol = DI_��Ժ��� Then
            '��Ҫ����ǻس��뿪:����ComboIndex,ȡ���༭ʱ����
            .TextMatrix(LngRow, LngCol) = zlStr.NeedName(.TextMatrix(LngRow, LngCol))
            If bln��ҽ Then
                If Not DiagCellEditable(vsDiag, LngRow, DI_�Ƿ�δ��) Then
                    .TextMatrix(LngRow, DI_�Ƿ�δ��) = ""
                End If
                If .TextMatrix(LngRow, DI_��Ժ���) = "����" Then
                    lngEnd = FindDiagRow(DT_�������)
                    lngStart = FindDiagRow(DT_��Ժ���XY)
                    For i = lngStart To lngEnd - 1
                        If .TextMatrix(i, DI_�������) <> "" Then .TextMatrix(LngRow, DI_�Ƿ�δ��) = ""
                    Next
                End If
            End If
            Call ChangeOutInfo(zlStr.NeedName(.TextMatrix(LngRow, DI_��Ժ���)))
        ElseIf LngCol = DI_������� Or gclsPros.FuncType = f������ҳ And gclsPros.CNIndent And LngCol = DI_��ϱ��� Then
            ' .EditText = "" �ų���Ԫ�������ݲ����س���״��
            If Not (gclsPros.CNIndent And (LngCol = DI_������� And .TextMatrix(LngRow, DI_��ϱ���) <> "" Or LngCol = DI_��ϱ��� And .TextMatrix(LngRow, DI_�������) <> "")) And .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                '�ڵ���vsDiagXY_KeyDown(vbKeyDelete, 0)���ǿ���ɾ����ǰ�У������ָ�ԭʼ����
                .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                Call DiagKeyDown(vsDiag, vbKeyDelete, 0)
            End If
        End If
        Call DiagAfterRowColChange(vsDiag, -1, -1, .Row, .Col)
         If LngCol = DI_��Ժ��� Then
              .TextMatrix(LngRow, LngCol) = zlStr.NeedName(.TextMatrix(LngRow, LngCol))
        End If
    End With
End Sub

Public Sub DiagAfterRowColChange(ByRef vsDiag As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsDiagZY_AfterRowColChange�¼���vsDiagXY_AfterRowColChange�¼�
    Dim i As Long
    Dim bln��ҽ As Boolean
    Dim vPoint As POINTAPI
    Dim blnEdit As Boolean
    Dim j As Long, arrTmp As Variant

    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    If vsDiag.Editable = flexEDNone Then Exit Sub
    With vsDiag
        bln��ҽ = .Name = "vsDiagXY"
        '���ͼƬ
        For i = .FixedRows To .Rows - 1
            Set .Cell(flexcpPicture, i, DI_����) = Nothing
            Set .Cell(flexcpPicture, i, DI_Del) = Nothing
        Next
        If bln��ҽ And gclsPros.FuncType <> f���ѡ�� Then Call ShowInfectInfo(False)
        If Not DiagCellEditable(vsDiag, lngNewRow, lngNewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .ComboList = ""
            .FocusRect = flexFocusSolid
            blnEdit = True
            Set .CellButtonPicture = Nothing
            If bln��ҽ And gclsPros.FuncType <> f���ѡ�� Then
                If .TextMatrix(lngNewRow, 0) = "Ժ�ڸ�Ⱦ" Then
                    If .TextMatrix(lngNewRow, DI_�������) <> "" Then
                        If lngNewCol = DI_������� Or lngNewCol = DI_��ע Then
                             vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
                             Call ShowInfectInfo(True, , vPoint.X, vPoint.Y)
                        End If
                    End If
                End If
            End If
            Select Case lngNewCol
                Case DI_�������
                    If Not (.TextMatrix(lngNewRow, DI_��ϱ���) <> "" And gclsPros.CNIndent And gclsPros.FuncType = f������ҳ) Then
                        .ComboList = "..."
                    End If
                Case DI_��ϱ���
                    If gclsPros.FuncType = f������ҳ And gclsPros.CNIndent Then .ComboList = "..."
                Case DI_����, DI_Del
                    .ComboList = "..."
                    .FocusRect = flexFocusNone
                    Set .CellButtonPicture = IIf(lngNewCol = DI_����, gclsPros.CurrentForm.imgButtonNew.Picture, gclsPros.CurrentForm.imgButtonDel.Picture)
                Case DI_��Ժ����
                    If blnEdit Then
                        .ComboList = "��|�ٴ�δȷ��|�������|��"
                         If Not gclsPros.IsCheckData Then OS.PressKey vbKeySpace
                    Else
                        .ComboList = ""
                        .FocusRect = flexFocusLight
                    End If
                Case DI_��Ժ���
                    .ComboList = .ColData(lngNewCol)
                    If Trim(.TextMatrix(lngNewRow, lngNewCol)) <> "" Then
                        arrTmp = Split(.ColData(lngNewCol) & "", "|")
                        For j = LBound(arrTmp) To UBound(arrTmp)
                            If zlStr.NeedName(arrTmp(j) & "") = .TextMatrix(lngNewRow, lngNewCol) Then
                                .TextMatrix(lngNewRow, lngNewCol) = arrTmp(j)
                                Exit For
                            End If
                        Next
                    End If
                Case DI_��ҽ֤��
                    If .TextMatrix(lngNewRow, DI_�������) = "" Then
                        .ComboList = ""
                        .FocusRect = flexFocusLight
                    Else
                        .ComboList = "..."
                    End If
                Case DI_ICD����
                    .ComboList = "..."
                Case Else
                    .ComboList = ""
            End Select
        End If
        If lngNewRow >= .FixedRows Then
            '��ʾͼƬ
            If lngNewCol <> DI_���� And .TextMatrix(lngNewRow, DI_�������) <> "" Then
                If .Rows - 1 <> lngNewRow Then
                    '��һ�����Ϊ������������
                    If Not (.TextMatrix(lngNewRow, DI_��Ϸ���) = .TextMatrix(lngNewRow + 1, DI_��Ϸ���) And .TextMatrix(lngNewRow + 1, DI_�������) = "") Then
                         Set .Cell(flexcpPicture, lngNewRow, DI_����) = gclsPros.CurrentForm.imgButtonNew.Picture
                    End If
                Else
                    Set .Cell(flexcpPicture, lngNewRow, DI_����) = gclsPros.CurrentForm.imgButtonNew.Picture
                End If
            End If
            '��ʾͼƬ
            If lngNewCol <> DI_Del Then Set .Cell(flexcpPicture, lngNewRow, DI_Del) = gclsPros.CurrentForm.imgButtonDel.Picture
        End If
    End With
End Sub

Public Sub DiagAfterUserResize(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsDiagZY_BeforeUserResize�¼���vsDiagXY_BeforeUserResize�¼�
    If LngCol = DI_������� And gclsPros.PatiType = PF_���� And gclsPros.FuncType = fҽ����ҳ Then
        If gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_��ҽ֤��) < gclsPros.CurrentForm.vsDiagXY.ColWidth(LngCol) Then
            gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_��ҽ֤��) = False
            gclsPros.CurrentForm.vsDiagZY.ColWidth(LngCol) = gclsPros.CurrentForm.vsDiagXY.ColWidth(LngCol) - gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_��ҽ֤��)
        Else
            gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_��ҽ֤��) = True
            gclsPros.CurrentForm.vsDiagZY.ColWidth(LngCol) = gclsPros.CurrentForm.vsDiagXY.ColWidth(LngCol)
        End If
    Else
        gclsPros.CurrentForm.vsDiagZY.ColWidth(LngCol) = gclsPros.CurrentForm.vsDiagXY.ColWidth(LngCol)
    End If
End Sub

Public Sub DiagBeforeUserResize(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsDiagZY_BeforeUserResize�¼���vsDiagXY_BeforeUserResize�¼�
    If LngCol = DI_���� Or LngCol = DI_Del Or LngCol < DI_��ϱ��� Then blnCancel = True
End Sub

Public Sub DiagCellButtonClick(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsDiagZY_CellButtonClick�¼���vsDiagXY_CellButtonClick�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lngCurRow As Long
    Dim bln��ҽ As Boolean
    
    With vsDiag
        bln��ҽ = .Name = "vsDiagXY"
        Select Case LngCol
            Case DI_�������, DI_��ϱ���
                If IIf(bln��ҽ, gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0 And gclsPros.FuncType <> f������ҳ Then
                    '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                    Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, IIf(bln��ҽ, "1", "2"), gclsPros.��Ժ����ID, , True, False)
                Else
                    'B-��ҽ�������룬7-�����ж���Y-�����ж����ⲿԭ��6-������ϣ�M-������̬ѧ���룻������ϣ�D-ICD-10��������
                    Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, IIf(bln��ҽ, decode(Val(.TextMatrix(LngRow, DI_��Ϸ���)), DT_�����ж���, "Y", DT_�������, IIf(gclsPros.M����, "M", "M,D"), "D"), "B"), gclsPros.��Ժ����ID, gclsPros.Sex, True, True, , gclsPros.SysNo)
                End If
                If Not rsTmp Is Nothing Then
                    Call SetDiagInput(vsDiag, LngRow, rsTmp)
                    Call EnterNextCellDiag(vsDiag)
                End If
            Case DI_ICD���� 'ֻ�в�����ҳ�ɼ����ɼ��Żᴥ���¼�
                'B-��ҽ�������룬7-�����ж���Y-�����ж����ⲿԭ��6-������ϣ�M-������̬ѧ���룻������ϣ�D-ICD-10��������
                Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, IIf(.TextMatrix(LngRow, DI_��ϱ���) Like "S*" Or .TextMatrix(LngRow, DI_��ϱ���) Like "T*", "Y", IIf(.TextMatrix(LngRow, DI_��ϱ���) Like "C*" Or (.TextMatrix(LngRow, DI_��ϱ���) Like "D*" And Val(Mid(.TextMatrix(LngRow, DI_��ϱ���), 2, 2)) <= 48), "M", "D")), gclsPros.��Ժ����ID, gclsPros.Sex, False, gclsPros.SysNo)
                If Not rsTmp Is Nothing Then
                    Call SetDiagInput(vsDiag, LngRow, rsTmp, True)
                    Call EnterNextCellDiag(vsDiag)
                End If
            Case DI_��ҽ֤��
                If gclsPros.DiagInputZY = 0 Then
                    '���������:�Ȳ��Ƿ��ж�Ӧ
                    If Set��ҽ֤��(LngRow, Val(.TextMatrix(LngRow, DI_���ID))) Then Exit Sub
                    Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, "Z", gclsPros.��Ժ����ID, gclsPros.Sex, True, , , gclsPros.SysNo)
                Else
                    'Z-��ҽ��������
                    Set rsTmp = zlDatabase.ShowILLSelect(gclsPros.CurrentForm, "Z", gclsPros.��Ժ����ID, gclsPros.Sex, True, , , gclsPros.SysNo)
                End If
                If Not rsTmp Is Nothing Then
                    Call Set��ҽ֤��(LngRow, 0, rsTmp)
                    Call EnterNextCellDiag(vsDiag)
                End If
            Case DI_����
                If Not .Cell(flexcpPicture, LngRow, DI_����) Is Nothing Or Not .CellButtonPicture Is Nothing Then
                    Call DiagKeyDown(vsDiag, vbKeyInsert, 0)
                    Set .CellButtonPicture = Nothing
                End If
            Case DI_Del
                If Not .Cell(flexcpPicture, LngRow, DI_Del) Is Nothing Or Not .CellButtonPicture Is Nothing Then
                    Call DiagKeyDown(vsDiag, vbKeyDelete, 0)
                End If
        End Select
    End With
End Sub

Public Sub DiagClick(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_Click�¼���vsDiagZY_Click�¼�
    Dim bln��ҽ As Boolean

    With vsDiag
        bln��ҽ = .Name = "vsDiagXY"
        If (.MouseCol = DI_���� Or .MouseCol = DI_Del) And .MouseRow >= .FixedRows Then
            If .MouseCol = DI_���� Then
                If .TextMatrix(.MouseRow, DI_�������) = "" Or .TextMatrix(.MouseRow, 0) = IIf(bln��ҽ, "��Ժ���", "��Ҫ���") Then Exit Sub
            End If
            .Select .MouseRow, .MouseCol
            Call DiagCellButtonClick(vsDiag, .MouseRow, .MouseCol)
        End If
    End With
End Sub

Public Sub DiagComboDropDown(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsDiagXY_ComboDropDown�¼���vsDiagZY_ComboDropDown�¼�
    Dim i As Long

    With vsDiag
        If LngCol = DI_��Ժ��� Or LngCol = DI_��Ժ���� Then
            '��λ��ƥ����
            For i = 0 To .ComboCount - 1
                If zlStr.NeedName(.ComboItem(i)) = .TextMatrix(LngRow, LngCol) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Public Sub DiagDblClick(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_DblClick�¼���vsDiagZY_DblClick�¼�
    Call DiagKeyPress(vsDiag, vbKeySpace)
End Sub

Public Sub DiagGotFocus(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_GotFocus�¼���vsDiagZY_GotFocus�¼�
    If vsDiag.Row >= vsDiag.FixedRows And vsDiag.Col >= vsDiag.FixedCols Then
        Call DiagAfterRowColChange(vsDiag, -1, -1, vsDiag.Row, vsDiag.Col)
    End If
End Sub

Public Sub DiagKeyDown(ByRef vsDiag As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsDiagXY_KeyDown�¼���vsDiagZY_KeyDown�¼�
    Dim i As Long, j As Long
    Dim dtCurRow As DiagType, LngRow As Long
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lngҽ��ID As Long, strMsg As String
    Dim blnDel As Boolean

    On Error GoTo errH
    If vsDiag.Editable = flexEDNone Then Exit Sub
    With vsDiag
        If intKeyCode = vbKeyF4 Then
            If .Col = DI_������� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf intKeyCode = vbKeyDelete Then
            If (.TextMatrix(.Row, DI_�������) <> "" Or gclsPros.PatiType = PF_���� And .Rows = .FixedRows + 1) And Not (gclsPros.FuncType = f������ҳ And gclsPros.CNIndent) Or gclsPros.FuncType = f������ҳ And gclsPros.CNIndent Then
                If .TextMatrix(.Row, DI_ҽ��IDs) <> "" Then
                    strMsg = "��������Ѿ�������ҽ��������ɾ����"
                    lngҽ��ID = Val(Mid(.TextMatrix(.Row, DI_ҽ��IDs), 1, InStr(.TextMatrix(.Row, DI_ҽ��IDs) & ",", ",") - 1))
                    If lngҽ��ID > 0 Then
                        strSql = "Select ҽ������ from ����ҽ����¼ where ����ID = [1] and ��ҳID = [2] and id =[3]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "����ҽ�����ݲ�ѯ", gclsPros.����ID, gclsPros.��ҳID, lngҽ��ID)
                        If rsTmp.RecordCount > 0 Then
                            strMsg = "��������Ѿ�������ҽ��:" & rsTmp!ҽ������ & "������ɾ����"
                        End If
                    End If
                    MsgBox strMsg, vbInformation, gstrSysName
                    Exit Sub
                End If
                If Not DiagCellEditable(vsDiag, .Row, DI_�������) Then Exit Sub
                '������ҳ����λ�������У��Ҹ����в�Ϊ�գ�����ո����У�����ɾ����,  ��Ժ���ͬ��
                If gclsPros.FuncType = f������ҳ Then
                    Select Case .Col
                        Case DI_ICD����
                            If .TextMatrix(.Row, DI_ICD����) <> "" Then
                                .TextMatrix(.Row, DI_ICD����) = ""
                                .TextMatrix(.Row, DI_����ID) = ""
                                .Cell(flexcpData, .Row, DI_ICD����) = ""
                                Exit Sub
                            End If
                        Case DI_��Ժ���
                            If .TextMatrix(.Row, DI_��Ժ���) <> "" Then
                                .TextMatrix(.Row, DI_��Ժ���) = ""
                                Call ChangeOutInfo
                                Exit Sub
                            End If
                        Case DI_��ҽ֤��
                            If .TextMatrix(.Row, DI_��ҽ֤��) <> "" Then
                                .TextMatrix(.Row, DI_��ҽ֤��) = ""
                                .TextMatrix(.Row, DI_֤��ID) = ""
                                .Cell(flexcpData, .Row, DI_��ҽ֤��) = ""
                                Exit Sub
                            End If
                    End Select
                End If

                blnDel = True
                If gclsPros.FuncType <> f������ҳ Then
                    If MsgBox("ȷʵҪ������������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        blnDel = False
                    End If
                End If

                If blnDel Then
                    'ɾ����/��Ҫ��Ϻ������ҽӿ�
                    If gclsPros.FuncType <> f������ҳ Then
                        If CreatePlugInOK(IIf(gclsPros.PatiType = PF_����, p����ҽ��վ, pסԺҽ��վ)) Then
                            If Not gobjPlugIn Is Nothing Then
                                On Error Resume Next
                                Call gobjPlugIn.DiagnosisDeleted(gclsPros.SysNo, IIf(gclsPros.PatiType = PF_����, p����ҽ��վ, pסԺҽ��վ), gclsPros.����ID, gclsPros.��ҳID, IIf(IIf(vsDiag.Name = "vsDiagXY", gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0, Val(.TextMatrix(.Row, DI_���ID)), Val(.TextMatrix(.Row, DI_����ID))), .TextMatrix(.Row, DI_�������))
                                Call zlPlugInErrH(Err, "DiagnosisDeleted")
                                Err.Clear: On Error GoTo 0
                            End If
                        End If
                    End If
                    dtCurRow = Val(.TextMatrix(.Row, DI_��Ϸ���))
                     'Ժ�ڸ�Ⱦ����Ⱦ��λ��������,��ҽ����
                    If dtCurRow = DT_Ժ�ڸ�Ⱦ Then Call ShowInfectInfo(False)
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedCols, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, DI_��Ϸ���) = dtCurRow
                    '�����ͬ�������������
                    If .TextMatrix(.Row, DI_�������) = "" Or gclsPros.PatiType = PF_���� And .Rows <> .FixedRows + 1 Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            If .TextMatrix(.Row + 1, DI_�������) = "" Then
                                For j = .FixedCols To .Cols - 1
                                    .TextMatrix(.Row, j) = .TextMatrix(.Row + 1, j)
                                    .Cell(flexcpData, .Row, j) = .Cell(flexcpData, .Row + 1, j)
                                Next
                                .RowData(.Row) = .RowData(.Row + 1)
                                .RemoveItem .Row + 1
                            End If
                        End If
                    End If
                    Call ChangeVSFHeight(vsDiag, True)
                End If
            ElseIf .TextMatrix(.Row, DI_�������) = "" Or gclsPros.PatiType = PF_���� And .Rows <> .FixedRows + 1 Then
                .RemoveItem .Row
                Call ChangeVSFHeight(vsDiag, True)
            End If
            '������������Ϣ
            If Not (gclsPros.FuncType = f���ѡ�� And gclsPros.PatiType = PF_����) Then
                Call SetDiagReletedInfo(vsDiag)
                If gclsPros.PatiType <> PF_���� Then Call ChangeOutInfo
            End If
        ElseIf intKeyCode = vbKeyInsert Then '������
            LngRow = .Row + 1: .AddItem "", LngRow
            Call ChangeVSFHeight(vsDiag, True)
            .TextMatrix(LngRow, DI_��Ϸ���) = .TextMatrix(LngRow - 1, DI_��Ϸ���)
            If gclsPros.PatiType = PF_���� Then .TextMatrix(LngRow, DI_�������) = .TextMatrix(LngRow - 1, DI_�������)
            .Cell(flexcpData, LngRow, DI_�������) = IIf(.TextMatrix(LngRow - 1, DI_�������) = "", .Cell(flexcpData, LngRow - 1, DI_�������), .TextMatrix(LngRow - 1, DI_�������))
            .Cell(flexcpForeColor, .FixedRows, DI_�Ƿ�����, .Rows - 1, DI_�Ƿ�����) = vbRed
            .Cell(flexcpBackColor, .FixedRows, DI_��ϱ���, .Rows - 1, DI_��ϱ���) = GRD_UNEDITCELL_COLOR      '����ɫ
            .Row = LngRow: .Col = DI_��ϱ���
            .ShowCell .Row, .Col
        ElseIf intKeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call DiagKeyPress(vsDiag, intKeyCode)
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub DiagKeyPress(ByRef vsDiag As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsDiagXY_KeyPress�¼���vsDiagZY_KeyPress�¼�
    If vsDiag.Editable = flexEDNone Then Exit Sub
    With vsDiag
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            Call EnterNextCellDiag(vsDiag)
        Else
            If Not DiagCellEditable(vsDiag, .Row, .Col) Then Exit Sub
            Select Case .Col
                Case DI_�Ƿ�δ��, DI_�Ƿ����� '��ҽ����������
                    If intKeyAscii <> vbKeySpace Then Exit Sub
                    intKeyAscii = 0
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", IIf(.Col = DI_�Ƿ�����, "��", "��"), "")
                Case DI_��ϱ���, DI_�������, DI_��ҽ֤��, DI_ICD���� '��ҽ��ҽ֤������,��ҽ��ICD��������
                    If intKeyAscii = Asc("*") Then
                        intKeyAscii = 0
                        Call DiagCellButtonClick(vsDiag, .Row, .Col)
                    Else
                        .ComboList = "" 'ʹ��ť״̬��������״̬
                    End If
            End Select
        End If
    End With
End Sub

Public Sub DiagKeyPressEdit(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsDiagXY_KeyPressEdit�¼���vsDiagZY_KeyPressEdit�¼�
    Dim bln��ҽ As Boolean

    If intKeyAscii = 13 Then
        gclsPros.IsReturn = True
        With vsDiag
            bln��ҽ = .Name = "vsDiagXY"
            If LngCol = DI_��Ժ��� Or LngCol = DI_��Ժ���� Then
                intKeyAscii = 0
                If .ComboIndex <> -1 Then
                    '��ʱ.TextMatrix��δ����,����ȡComboItem
                    .TextMatrix(LngRow, LngCol) = zlStr.NeedName(.ComboItem(.ComboIndex))
                    If bln��ҽ And LngCol = DI_��Ժ��� Then
                        If Not DiagCellEditable(vsDiag, LngRow, DI_�Ƿ�δ��) Then .TextMatrix(LngRow, DI_�Ƿ�δ��) = ""
                    End If
                    Call EnterNextCellDiag(vsDiag)
                 End If
            End If
        End With
    Else
        gclsPros.IsReturn = False
    End If
End Sub


Public Sub DiagSetupEditWindow(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByVal lngEditWindow As Long, ByVal blnIsCombo As Boolean)
'vsDiagXY_SetupEditWindow�¼���vsDiagZY_SetupEditWindow�¼�
    With vsDiag
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Public Sub DiagStartEdit(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsDiagXY_StartEdit�¼���vsDiagZY_StartEdit�¼�
    If gclsPros.FuncType = f���ѡ�� And gclsPros.IsSigned And LngCol <> DI_���� Then
        blnCancel = True
        MsgBox "�ò��˵���ҳ�Ѿ�ǩ�������޸���ϡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not DiagCellEditable(vsDiag, LngRow, LngCol) Then
        blnCancel = True
    ElseIf LngCol = DI_�Ƿ�δ�� Or LngCol = DI_�Ƿ����� Then '��ҽ�ſ��ܽ���÷�֧
        blnCancel = True '��ֱ�ӱ༭
    End If
End Sub

Public Sub DiagValidateEdit(ByRef vsDiag As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsDiagXY_ValidateEdit�¼���vsDiagZY_ValidateEdit�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim int������� As Integer
    Dim strInput As String, vPoint As POINTAPI
    Dim strDiagType As String
    Dim bln��ҽ As Boolean
    Dim str�Ա� As String

    With vsDiag
        bln��ҽ = .Name = "vsDiagXY"
        Select Case LngCol
            Case DI_�������, DI_��ϱ���
                If bln��ҽ Then
                    strDiagType = decode(Val(.TextMatrix(LngRow, DI_��Ϸ���)), 7, "'Y'", 6, IIf(gclsPros.M����, "'M'", "'M,D'"), "'D'")
                Else
                    strDiagType = IIf(gclsPros.DiagInputZY = 0, "", "B")
                End If
                If gclsPros.FuncType = f������ҳ Then
                     If .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                        If LngCol = DI_������� Or LngCol = DI_��ϱ��� And (.TextMatrix(LngRow, DI_�������) = "" Or gclsPros.DaigFree And Not bln��ҽ) Then
                            .EditText = ""
                        Else
                            .EditText = .Cell(flexcpData, LngRow, LngCol)
                        End If
                    ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                        If gclsPros.IsReturn Then Call EnterNextCellDiag(vsDiag)
                    '�������ƶ���ʱ�����༭�������ʱ�������벻Ϊ�գ���������¼�룬��������ƥ��
                    ElseIf Not (LngCol = DI_������� And .TextMatrix(LngRow, DI_��ϱ���) <> "" And gclsPros.CNIndent) Then
                        If Val(.TextMatrix(LngRow, DI_��Ϸ���)) = IIf(bln��ҽ, DT_�������XY, DT_�������ZY) Then
                            int������� = gclsPros.DiagSourceMZ
                        Else
                            int������� = gclsPros.DiagSourceZY
                        End If
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(IIf(bln��ҽ, 0, 1), strInput, str�Ա�, strDiagType)
                        vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(IIf(bln��ҽ, gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0, "�������", "��������"), _
                            False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                            strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str�Ա�, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.��Ժ����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                        If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                            If Not (gclsPros.DaigFree And Not bln��ҽ And LngCol = DI_�������) Then
                                blnCancel = True
                            End If
                        Else
                            '���������뷽ʽ
                            If rsTmp Is Nothing Then
                                If Not (gclsPros.DaigFree And Not bln��ҽ And LngCol = DI_�������) Then
                                    MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                                    blnCancel = True
                                End If
                            Else
                                Call SetDiagInput(vsDiag, LngRow, rsTmp): .EditText = .Text
                                'If mblnReturn Then Call XYEnterNextCell    '�ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                            End If
                        End If
                    End If
                Else
                    If .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                        .EditText = ""
                    ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                        If gclsPros.IsReturn Then Call EnterNextCellDiag(vsDiag)
                    ElseIf .TextMatrix(LngRow, DI_��ϱ���) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" And .EditText Like "*" & .Cell(flexcpData, LngRow, LngCol) & "*" Then
                        '�жϼ���ǰ׺��������Ƿ������������ϱ���
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(IIf(bln��ҽ, 0, 1), strInput, str�Ա�, strDiagType)
                        On Error GoTo errH
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput, strInput, strDiagType, str�Ա�, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.��Ժ����ID)
                        If rsTmp.RecordCount = 1 Then
                            Call SetDiagInput(vsDiag, LngRow, rsTmp)
                            .EditText = .Text
                        Else
                            '�����ڱ�׼������ǰ�����븽����Ϣ
                            '������.Cell(flexcpData, lngRow, lngCol)���Ա��޸�����ʱ�ٴ�ʹ��like�ж�
                            .TextMatrix(LngRow, DI_�������) = .EditText
                        End If
                    ElseIf .TextMatrix(LngRow, DI_��ϱ���) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" And gclsPros.FreeInput Then
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(IIf(bln��ҽ, 0, 1), strInput, str�Ա�, strDiagType)
                        On Error GoTo errH
                        vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(IIf(bln��ҽ, gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0, "�������", "��������"), _
                            False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                            strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str�Ա�, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.��Ժ����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                        If blnInputCancel Then
                            blnCancel = True
                        Else
                            If rsTmp Is Nothing Then
                                .TextMatrix(LngRow, DI_�������) = .EditText
                            Else
                                 Call SetDiagInput(vsDiag, LngRow, rsTmp): .EditText = .Text
                            End If
                        End If
                    Else
                        If Val(.TextMatrix(LngRow, DI_��Ϸ���)) = IIf(bln��ҽ, DT_�������XY, DT_�������ZY) Then
                            int������� = gclsPros.DiagSourceMZ
                        Else
                            int������� = gclsPros.DiagSourceZY
                        End If
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(IIf(bln��ҽ, 0, 1), strInput, str�Ա�, strDiagType)
                        If False And int������� = 1 And zlCommFun.IsCharChinese(strInput) Then
                            '�����ж��룺Y-�����ж����ⲿԭ�򣻲����������M-������̬ѧ���룻������ϣ�D-ICD-10��������
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str�Ա�, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.��Ժ����ID)
                            If rsTmp.EOF Then
                                Set rsTmp = Nothing
                            ElseIf rsTmp.RecordCount > 1 Then
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput, strInput, strDiagType, str�Ա�, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.��Ժ����ID)
                                If rsTmp.RecordCount <> 1 Then Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                            End If
                            Call SetDiagInput(vsDiag, LngRow, rsTmp)
                            .EditText = .Text
                            If gclsPros.IsReturn And rsTmp Is Nothing Then Call EnterNextCellDiag(vsDiag) '��������¼��ʱ���ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                        Else
                            vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                            Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(IIf(bln��ҽ, gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0, "�������", "��������"), _
                                False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                                strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str�Ա�, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.��Ժ����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                            If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                                blnCancel = True
                            Else
                                '���������뷽ʽ
                                If rsTmp Is Nothing And ((int������� = 2 Or int������� = 3 And gclsPros.InsureType <> 0)) Then
                                    MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                                    blnCancel = True
                                Else
                                    Call SetDiagInput(vsDiag, LngRow, rsTmp): .EditText = .Text
                                    'If mblnReturn Then Call XYEnterNextCell    '�ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                                End If
                            End If
                        End If
                    End If
                End If
                gclsPros.IsReturn = False
            Case DI_��ҽ֤��
                If .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                    .EditText = ""
                    .Cell(flexcpData, LngRow, LngCol) = ""
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If gclsPros.IsReturn Then Call EnterNextCellDiag(vsDiag)
                ElseIf .TextMatrix(LngRow, DI_��ϱ���) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" And gclsPros.FreeInput Then
                    strDiagType = "Z"
                    strInput = UCase(.EditText)
                    strSql = GetMedInputSQL(1, strInput, str�Ա�, strDiagType)
                    vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "��ҽ֤��", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str�Ա�, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.��Ժ����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                    If blnInputCancel Then      '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        blnCancel = True
                    Else
                        If rsTmp Is Nothing Then
                            .TextMatrix(LngRow, DI_��ҽ֤��) = .EditText
                        Else
                            Call Set��ҽ֤��(LngRow, 0, rsTmp, rsTmp Is Nothing)
                        End If
                    End If
                Else
                    strInput = UCase(.EditText)
                    strDiagType = "Z"
                    strSql = GetMedInputSQL(1, strInput, str�Ա�, strDiagType)

                    vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "��ҽ֤��", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str�Ա�, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.��Ժ����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                    If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        blnCancel = True
                    Else
                        If Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_�������ZY Then
                            int������� = gclsPros.DiagSourceMZ
                        Else
                            int������� = gclsPros.DiagSourceZY
                        End If
                        '���������뷽ʽ
                         If rsTmp Is Nothing And (int������� = 2 Or (int������� = 3 And gclsPros.InsureType <> 0)) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            blnCancel = True
                         Else
                            Call Set��ҽ֤��(LngRow, 0, rsTmp, rsTmp Is Nothing)
                         End If
                    End If
                End If
                gclsPros.IsReturn = False
            Case DI_ICD����
                If .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                    .EditText = ""
                    .TextMatrix(LngRow, DI_����ID) = ""
                    .Cell(flexcpData, LngRow, LngCol) = ""
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If gclsPros.IsReturn Then Call EnterNextCellDiag(vsDiag)
                Else
                    strInput = UCase(.EditText)
                    'B-��ҽ�������룬7-�����ж���Y-�����ж����ⲿԭ��6-������ϣ�M-������̬ѧ���룻������ϣ�D-ICD-10��������
                    strDiagType = IIf(.TextMatrix(LngRow, DI_��ϱ���) Like "S*" Or .TextMatrix(LngRow, DI_��ϱ���) Like "T*", "Y", IIf(.TextMatrix(LngRow, DI_��ϱ���) Like "C*" Or (.TextMatrix(LngRow, DI_��ϱ���) Like "D*" And Val(Mid(.TextMatrix(LngRow, DI_��ϱ���), 2, 2)) <= 48), "M", "D"))
                    strSql = GetMedInputSQL(0, strInput, str�Ա�)
                    vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(IIf(bln��ҽ, gclsPros.DiagInputXY, gclsPros.DiagInputZY) = 0, "�������", "��������"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gclsPros.LikeString & strInput & "%", strDiagType, str�Ա�, gclsPros.BriefCode + 1, strInput, UserInfo.ID, gclsPros.��Ժ����ID, "ColSet:�п�����|˵��,2400|������ʾ|˵��")
                    If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        blnCancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            blnCancel = True
                        Else
                            Call SetDiagInput(vsDiag, LngRow, rsTmp, True): .EditText = .Text
                            If gclsPros.IsReturn Then Call EnterNextCellDiag(vsDiag)
                        End If
                    End If
                End If
                gclsPros.IsReturn = False
            Case DI_��Ժ���
                If .EditText <> "" Then
                    If .TextMatrix(.Row, DI_��Ч����) <> "" And InStr(.EditText, .TextMatrix(.Row, DI_��Ч����)) > 0 Then
                        MsgBox "��ע�⣬�ü���ͨ�����ܴﵽ������Ч�ġ�", vbInformation, gstrSysName
                        .EditText = "": blnCancel = True: Exit Sub
                    End If
                End If
            Case DI_����ʱ��
                If .EditText <> "" Then
                    strInput = zlStr.FullDate(.EditText)
                    If IsDate(strInput) Then
                        .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                    Else
                        MsgBox "��������ȷ�ķ���ʱ�䣬���磺""2012-12-21 00:00""��", vbInformation, gstrSysName
                        blnCancel = True
                    End If
                End If
                If LngRow = .FixedRows And gclsPros.FuncType <> f���ѡ�� Then
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_����ʱ��), IsDate(.EditText), True)
                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��������), IsDate(.EditText), True)
                End If
        End Select
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'vsOPS�¼�
Public Sub OPSAfterEdit(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsOPS_AfterEdit�¼�
    Dim strInput As String

    With vsOPS
        Select Case LngCol
            Case PI_��������, PI_��������, PI_������ҩʱ��, PI_����ʼʱ��
                If LngCol <> PI_������ҩʱ�� Then
                    strInput = Format(zlStr.FullDate(.TextMatrix(LngRow, LngCol), , gclsPros.InTime, gclsPros.OutTime), "yyyy-mm-dd hh:mm")
                Else
                    strInput = Format(zlStr.FullDate(.TextMatrix(LngRow, LngCol)), "yyyy-mm-dd hh:mm")
                End If
                If Not IsDate(strInput) Then
                    .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                Else
                    .TextMatrix(LngRow, LngCol) = strInput
                    .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                End If
            Case PI_�п�����, PI_��������
                .TextMatrix(LngRow, LngCol) = zlStr.NeedName(.TextMatrix(LngRow, LngCol))
        End Select
'        Call OPSAfterRowColChange(vsOPS, -1, -1, LngRow, LngCol)
        '������Ϸ������
        Call SetDiagMatchInfo(BCC_��ǰ������)
    End With
End Sub

Public Sub OPSAfterRowColChange(ByRef vsOPS As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsOPS_AfterRowColChange�¼�
    Dim blnEdit As Boolean

    With vsOPS
        If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
        If vsOPS.Editable <> flexEDNone Then Call SetCopyImage(vsOPS)
        If Not OPSCellEditable(lngNewRow, lngNewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Select Case lngNewCol
                Case PI_��������, PI_����ҽʦ, PI_������ʿ, PI_����1, PI_����2, PI_����ʽ, PI_����ҽʦ, PI_�пڲ�λ 'PI_����ʽΪ������
                    .ComboList = "..."
                Case PI_��������
                    If gclsPros.FuncType <> f������ҳ Then
                        '�������Ʋ�������
                        blnEdit = gclsPros.CurrentForm.chkParaOPSInfo(PC_δ�ҵ�ʱ����¼��).Value
                    Else
                        blnEdit = gclsPros.CNIndent
                    End If
                    If blnEdit Then
                        .ComboList = "..."
                    Else
                        .ComboList = ""
                    End If
                Case PI_�������
                    If Not gclsPros.IsCheckData Then OS.PressKey vbKeySpace
                Case PI_�п�����, PI_��������
                    .ComboList = .ColData(lngNewCol)
                Case Else
                    .ComboList = ""
            End Select
        End If
    End With
End Sub

Public Sub OPSBeforeUserResize(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsOPS_BeforeUserResize�¼�
    blnCancel = LngCol = PI_Copy
End Sub

Public Sub OPSCellButtonClick(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsOPS_CellButtonClick�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim int�Ա� As Integer, int�������뷽ʽ As Integer
    Dim vPoint As POINTAPI, strInfoName As String
    Dim textTmp As TextBox

    With vsOPS
        Select Case LngCol
            Case PI_��������, PI_��������
                If gclsPros.Sex Like "*��*" Then
                    int�Ա� = 1
                ElseIf gclsPros.Sex Like "*Ů*" Then
                    int�Ա� = 2
                End If
                int�������뷽ʽ = Val(gclsPros.OPSInput)
                If int�������뷽ʽ = 0 And gclsPros.FuncType <> f������ҳ Then
                    '��������Ŀ����
                    strSql = "Select 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��ģ" & _
                        " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                        " Union ALL " & _
                        " Select 1 as ĩ��,ID,����ID as �ϼ�ID,����,����,�������� as ��ģ" & _
                        " From ������ĿĿ¼" & _
                        " Where ���='F' And ������� IN(2,3) And (վ��='" & gstrNodeNo & "' Or վ�� is Null)" & _
                        IIf(int�Ա� <> 0, " And Nvl(�����Ա�,0) IN(0,[2])", "") & _
                        " And (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)"
                Else
                    '��ICD9-CM3����
                    strSql = _
                        " Select 0 as ĩ��,ID,�ϼ�ID," & _
                        " ���||LPAD(���,3,'0') as ����," & _
                        " NULL as ����,����,����,NULL as ˵��" & _
                        " From ����������� Where ���='S'" & _
                        " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) " & _
                        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                        " Union ALL " & _
                        " Select 1 as ĩ��,ID,����ID as �ϼ�ID,����,����,����,����,˵��" & _
                        " From ��������Ŀ¼ Where ���='S'" & _
                        IIf(int�Ա� <> 0, " And (�Ա�����=[1] Or �Ա����� is NULL)", "") & _
                        " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
                End If
                strInfoName = IIf(int�������뷽ʽ = 0 And gclsPros.FuncType <> f������ҳ, "������Ŀ", "��������")
                Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 2, strInfoName, False, "", "", False, True, False, _
                            0, 0, 0, blnCancel, False, False, decode(int�Ա�, 1, "��", 2, "Ů", ""), int�Ա�)
            Case PI_����ʽ '����Ϊ������
                If gclsPros.FuncType <> f������ҳ Then
                    strSql = "Select 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��������" & _
                    " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                    " Union ALL " & _
                    " Select 1 as ĩ��,ID,����ID as �ϼ�ID,����,����,�������� as ��������" & _
                    " From ������ĿĿ¼ Where ���='G'" & _
                    " And (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                    " And (վ��='" & gstrNodeNo & "' Or վ�� is Null)"
                    strInfoName = "������Ŀ"
                    Set rsTmp = zlDatabase.ShowSelect(gclsPros.CurrentForm, strSql, 2, strInfoName, , , , , True, , , , , blnCancel)
                End If
            Case PI_����ҽʦ, PI_����1, PI_����2, PI_����ҽʦ, PI_������ʿ
                strInfoName = IIf(LngCol = PI_������ʿ, "��ʿ", "ҽ��")
                Set rsTmp = GetManData(strInfoName)
                Set textTmp = GetReplaceObject(vsOPS)
                blnCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , "ȱʡ,ҽ��,��ʿ,��������Ա,����,סԺ,����,����", rsTmp)
            Case PI_�пڲ�λ
                strSql = "Select Rownum As ID, A.����, A.����, A.���� From �пڲ�λ A"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ҳ��ȡ�пڲ�λ")
                Set textTmp = GetReplaceObject(vsOPS)
                blnCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , "�пڲ�λ", rsTmp)
        End Select
        '��Ŀ�������
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û��" & strInfoName & "����ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp)
            Call EnterNextCellOPS(vsOPS)
        End If
    End With
End Sub

Public Sub OPSComboDropDown(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsOPS_ComboDropDown�¼�
    Dim i As Long

    With vsOPS
        If LngCol = PI_�п����� Or LngCol = PI_�������� Or LngCol = PI_������� Then
            For i = 0 To .ComboCount - 1
                If zlStr.NeedName(.ComboItem(i)) = .TextMatrix(LngRow, LngCol) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Public Sub OPSDblClick(ByRef vsOPS As VSFlexGrid)
'vsOPS_DblClick�¼�
    Call OPSKeyPress(vsOPS, vbKeySpace)
End Sub

Public Sub OPSKeyDown(ByRef vsOPS As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsOPS_KeyDown�¼�
    Dim i As Long

    If vsOPS.Editable = flexEDNone Then Exit Sub
    With vsOPS
        If intKeyCode = vbKeyF4 Then
            If .ComboList = "..." Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf intKeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, PI_��������) <> "" Then
                If MsgBox("ȷʵҪɾ������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    Call ChangeVSFHeight(vsOPS, True, 600, 2)
                    '������Ϸ������
                    Call SetDiagMatchInfo(BCC_��ǰ������)
                End If
            End If
        ElseIf intKeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call OPSKeyPress(vsOPS, intKeyCode)
        End If
    End With
End Sub

Public Sub OPSKeyPress(ByRef vsOPS As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsOPS_KeyPress�¼�
    If vsOPS.Editable = flexEDNone Then Exit Sub
    With vsOPS
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            Call EnterNextCellOPS(vsOPS)
        Else
            If .ComboList = "..." Then
                If intKeyAscii = Asc("*") Then
                    intKeyAscii = 0
                    Call OPSCellButtonClick(vsOPS, .Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Public Sub OPSKeyPressEdit(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsOPS_KeyPressEdit�¼�
    Dim strInput As String

    With vsOPS
        If intKeyAscii = vbKeyReturn Then
            gclsPros.IsReturn = True
            Select Case LngCol
                Case PI_�п�����, PI_��������, PI_�������
                    intKeyAscii = 0
                    If .ComboIndex <> -1 Then
                        .TextMatrix(LngRow, LngCol) = zlStr.NeedName(.ComboItem(.ComboIndex))
                        Call EnterNextCellOPS(vsOPS)
                    End If
            End Select
        Else
            gclsPros.IsReturn = False
            If LngCol = PI_�������� Or LngCol = PI_�������� Or LngCol = PI_����ʼʱ�� Or LngCol = PI_������ҩʱ�� Then
                If InStr("0123456789-" & Chr(8) & Chr(27), Chr(intKeyAscii)) = 0 Then
                    intKeyAscii = 0
                End If
            ElseIf LngCol = PI_׼������ Then
                If InStr("0123456789" & Chr(8) & Chr(27), Chr(intKeyAscii)) = 0 Then
                    intKeyAscii = 0
                End If
            ElseIf LngCol = PI_����ҩ���� Then
                If InStr("0123456789" & Chr(8) & Chr(27), Chr(intKeyAscii)) = 0 Then
                    intKeyAscii = 0
                End If
            End If
        End If
    End With
End Sub

Public Sub OPSSetupEditWindow(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByVal lngEditWindow As Long, ByVal blnIsCombo As Boolean)
'vsOPS_SetupEditWindow�¼�
    With vsOPS
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Public Sub OPSStartEdit(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsOPS_StartEdit�¼�
    If Not OPSCellEditable(LngRow, LngCol) Then
        blnCancel = True
    ElseIf LngCol = PI_�пڲ�λ Or LngCol = PI_�ط�������Ŀ�� Then
        vsOPS.EditMaxLength = 100
    End If
End Sub

Public Sub OPSValidateEdit(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsOPS_ValidateEdit�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim str�Ա� As String, int�Ա� As Integer
    Dim strInput As String, vPoint As POINTAPI
    Dim textTmp As TextBox
    Dim strTmp As String

    On Error GoTo errH
    With vsOPS
        Select Case LngCol
            Case PI_��������, PI_��������
                If gclsPros.FuncType = f������ҳ Then
                    If .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                        If LngCol = PI_�������� Or (LngCol = PI_�������� And .TextMatrix(LngRow, PI_��������) = "") Then
                            .EditText = ""
                        Else
                            .EditText = .Cell(flexcpData, LngRow, LngCol)
                        End If
                    ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                        '�������ƶ���ʱ�����༭��������ʱ�������벻Ϊ�գ���������¼�룬��������ƥ��
                    ElseIf Not (LngCol = PI_�������� And .TextMatrix(LngRow, PI_��������) <> "" And gclsPros.CNIndent) Then
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(2, strInput, str�Ա�)
                        If str�Ա� = "��" Then
                            int�Ա� = 1
                        ElseIf str�Ա� = "Ů" Then
                            int�Ա� = 2
                        End If
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(gclsPros.OPSInput = "0", "������Ŀ", "��������"), False, "", "", False, True, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%", str�Ա�, int�Ա�)
                        If rsTmp Is Nothing Then
                            If Not blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                                MsgBox "û���ҵ������ҵ�������Ŀ��", vbInformation, gstrSysName
                                blnCancel = True
                            Else
                                blnCancel = True
                            End If
                        Else
                            Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                            If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                        End If
                    End If
                Else
                    If .EditText = "" Then
                        .EditText = .Cell(flexcpData, LngRow, LngCol)
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                    ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                    ElseIf LngCol = PI_�������� And .TextMatrix(LngRow, PI_��������) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" And .EditText Like "*" & .Cell(flexcpData, LngRow, LngCol) & "*" Then
                        '�жϼ���ǰ׺��������Ƿ������������ϱ���
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(2, strInput, str�Ա�)
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput & "%", gclsPros.LikeString & strInput & "%", str�Ա�, int�Ա�)
                        If rsTmp.RecordCount <> 1 Then
                            '�����ڱ�׼������ǰ�����븽����Ϣ
                            .TextMatrix(LngRow, PI_��������) = .EditText
                        Else
                            Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp)
                            .EditText = .Text '������.Cell(flexcpData, lngRow, lngCol)���Ա��޸�����ʱ�ٴ�ʹ��like�ж�
                        End If
                    ElseIf LngCol = PI_�������� And .TextMatrix(LngRow, PI_��������) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" And gclsPros.FreeInput Then
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(2, strInput, str�Ա�)
                        If str�Ա� = "��" Then
                            int�Ա� = 1
                        ElseIf str�Ա� = "Ů" Then
                            int�Ա� = 2
                        End If
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(gclsPros.OPSInput = "0", "������Ŀ", "��������"), False, "", "", False, True, True, _
                                    vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%", str�Ա�, int�Ա�)
                        If blnInputCancel Then
                            blnCancel = True
                        Else
                            If rsTmp Is Nothing Then
                                .TextMatrix(LngRow, PI_��������) = .EditText
                            Else
                                Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                            End If
                        End If
                    Else
                        strInput = UCase(.EditText)
                        strSql = GetMedInputSQL(2, strInput, str�Ա�)
                        If str�Ա� = "��" Then
                            int�Ա� = 1
                        ElseIf str�Ա� = "Ů" Then
                            int�Ա� = 2
                        End If
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, IIf(gclsPros.OPSInput = "0", "������Ŀ", "��������"), False, "", "", False, True, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%", str�Ա�, int�Ա�)
                        If rsTmp Is Nothing Then
                            If Not blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                                If Not gclsPros.OPSFree Or LngCol = PI_�������� Then
                                    MsgBox "û���ҵ������ҵ�������Ŀ��", vbInformation, gstrSysName
                                    blnCancel = True
                                Else
                                    .TextMatrix(LngRow, PI_��������) = ""
                                    .TextMatrix(LngRow, PI_������ĿID) = ""
                                    .Cell(flexcpData, LngRow, PI_��������) = ""
                                    .TextMatrix(LngRow, PI_��������ID) = ""
                                    '�����ʼ�ձ���һ����
                                    If LngRow = .Rows - 1 Then .AddItem "": Call ChangeVSFHeight(vsOPS, True)
                                End If
                            Else
                                blnCancel = True
                            End If
                        Else
                            Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                            If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                        End If
                    End If
                End If
                gclsPros.IsReturn = False
            Case PI_��������, PI_��������, PI_����ʼʱ��, PI_������ҩʱ��
                If LngCol <> PI_������ҩʱ�� Then
                    strInput = Format(zlStr.FullDate(.EditText, , gclsPros.InTime, gclsPros.OutTime), "yyyy-mm-dd hh:mm")
                Else
                    strInput = Format(zlStr.FullDate(.EditText), "yyyy-mm-dd hh:mm")
                End If
                If IsDate(strInput) Then
                    '������ҩ��������Ժ��ʹ�ã���˲������
                    If Not CheckDateRange(strInput) And LngCol <> PI_������ҩʱ�� Then
                        MsgBox "�������ʱ������ڲ��˵�סԺ�ڼ䡣", vbInformation, gstrSysName
                        .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                        blnCancel = True
                        Exit Sub
                    Else
                        .TextMatrix(LngRow, LngCol) = strInput
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                        Call EnterNextCellOPS(vsOPS)
                    End If
                End If
                If strInput = "" And (LngCol = PI_����ʼʱ�� Or LngCol = PI_������ҩʱ��) Then
                    .TextMatrix(LngRow, LngCol) = strInput
                    .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                    Call EnterNextCellOPS(vsOPS)
                End If
            Case PI_����ʽ
                If .EditText = "" Then
                    .EditText = .Cell(flexcpData, LngRow, LngCol)
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                Else
                    strInput = UCase(.EditText)
                    strSql = _
                        " Select A.ID,A.����,A.����,A.�������� as ��������" & _
                        " From ������ĿĿ¼ A,������Ŀ���� B" & _
                        " Where A.���='G' And A.ID=B.������ĿID" & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                        " And (A.���� Like [1] Or A.���� Like [2] Or B.���� Like [2] Or B.���� Like [2])" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " Order by A.����"

                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "������Ŀ", False, "", "", False, True, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gclsPros.LikeString & strInput & "%")
                    If rsTmp Is Nothing Then
                        If Not blnInputCancel Then
                            MsgBox "û���ҵ�ƥ���������Ŀ��", vbInformation, gstrSysName
                        End If
                        blnCancel = True
                    Else
                        Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                    End If
                End If
                gclsPros.IsReturn = False
            Case PI_����ҽʦ, PI_����1, PI_����2, PI_����ҽʦ
                If (LngCol = PI_����1 Or LngCol = PI_����2) And .EditText = "" Then
                    .TextMatrix(LngRow, LngCol) = "": .Cell(flexcpData, LngRow, LngCol) = ""
                    If LngCol = PI_����1 Then
                        .TextMatrix(LngRow, PI_����2) = "": .Cell(flexcpData, LngRow, PI_����2) = ""
                    End If
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                ElseIf .EditText = "" Then
                    .EditText = .Cell(flexcpData, LngRow, LngCol)
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                Else
                    strInput = UCase(.EditText)
                    strSql = "���� Like '" & strInput & "*' OR ���� Like '*" & strInput & "*' OR ���� Like '*" & strInput & "*' OR ��ʼ��� Like '*" & strInput & "*'"
                    Set rsTmp = Rec.FilterNew(GetManData("ҽ��"), strSql)
                    If rsTmp.RecordCount <> 0 Then
                        Set textTmp = GetReplaceObject(vsOPS)
                        blnInputCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , "ȱʡ,ҽ��,��ʿ,��������Ա,����,סԺ,����,����", rsTmp)
                    Else
                        Set rsTmp = Nothing
                    End If
                    If rsTmp Is Nothing Then
                        If Not blnInputCancel Then
                            If (LngCol = PI_����ҽʦ Or LngCol = PI_����1 Or LngCol = PI_����2 Or LngCol = PI_����ҽʦ) And zlCommFun.IsCharChinese(.EditText) And Not gclsPros.IsOutDocCtrl Then
                                If MsgBox("û���ҵ�ƥ��ı�Ժҽ�����Ƿ�¼��δ�ڱ�Ժ������ҽ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                                    Exit Sub
                                End If
                            Else
                                MsgBox "û���ҵ�ƥ���ҽ����", vbInformation, gstrSysName
                            End If
                        End If
                        blnCancel = True
                    Else
                        Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                    End If
                End If
                gclsPros.IsReturn = False
            Case PI_������ʿ
                If .EditText = "" Then
                    .EditText = .Cell(flexcpData, LngRow, LngCol)
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                Else
                    strInput = UCase(.EditText)
                    strSql = "���� Like '" & strInput & "*' OR ���� Like '*" & strInput & "*' OR ���� Like '*" & strInput & "*' OR ��ʼ��� Like '*" & strInput & "*'"
                    Set rsTmp = Rec.FilterNew(GetManData("��ʿ"), strSql)
                    If rsTmp.RecordCount <> 0 Then
                        Set textTmp = GetReplaceObject(vsOPS)
                        blnInputCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , "ȱʡ,ҽ��,��ʿ,��������Ա,����,סԺ,����,����", rsTmp)
                    Else
                        Set rsTmp = Nothing
                    End If
                    If rsTmp Is Nothing Then
                        If Not blnInputCancel Then
                            MsgBox "û���ҵ�ƥ��Ļ�ʿ��", vbInformation, gstrSysName
                        End If
                        blnCancel = True
                    Else
                        Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                        If gclsPros.IsReturn Then Call EnterNextCellOPS(vsOPS)
                    End If
                End If
                gclsPros.IsReturn = False
            Case PI_�пڲ�λ
                If .EditText <> "" And .EditText <> .Cell(flexcpData, LngRow, LngCol) Then
                    strInput = UCase(.EditText)
                    strSql = "���� Like '" & strInput & "*' OR ���� Like '*" & strInput & "*' OR ���� Like '*" & strInput & "*'"
                    strTmp = "Select Rownum As ID, A.����, A.����, A.���� From �пڲ�λ A"
                    Set rsTmp = Rec.FilterNew(zlDatabase.OpenSQLRecord(strTmp, "��ҳ��ȡ�пڲ�λ"), strSql)
                    If rsTmp.RecordCount <> 0 Then
                        Set textTmp = GetReplaceObject(vsOPS)
                        blnInputCancel = Not zlDatabase.zlShowListSelect(gclsPros.CurrentForm, gclsPros.SysNo, gclsPros.Module, textTmp, rsTmp, True, , "�пڲ�λ", rsTmp)
                    Else
                        Set rsTmp = Nothing
                    End If
                    If Not rsTmp Is Nothing Then
                        Call OPSSetInput(vsOPS, LngRow, LngCol, rsTmp): .EditText = .Text
                    End If
                End If
                gclsPros.IsReturn = False
        End Select
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub picCopyClick()
'picCopy_Click �¼�
    Dim vsOPS As VSFlexGrid
    Dim i As Long, LngRow As Long
    Set vsOPS = gclsPros.CurrentForm.vsOPS

    With vsOPS
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, PI_��������) = "" Then
                LngRow = i: Exit For
            End If
        Next
        If LngRow = 0 Then
            .Rows = .Rows + 1
            LngRow = .Rows - 1
            Call ChangeVSFHeight(vsOPS, True)
        End If
        For i = .FixedCols To .Cols - 1
            If i <> PI_�������� And i <> PI_�������� Then
                .TextMatrix(LngRow, i) = .TextMatrix(.Row, i)
            End If
        Next
    End With
End Sub

Private Sub SetCopyImage(ByRef vsOPS As VSFlexGrid)
    Dim blnShow As Boolean
    Dim lngRowHeight As Long, i As Long, lngHeight As Long
    With vsOPS
        blnShow = .TextMatrix(.Row, PI_��������) <> "" And .Row >= .FixedRows And .ColIsVisible(PI_Copy)
        If blnShow Then
            For i = 0 To .Row - 1
                lngRowHeight = .RowHeight(i)
                If .RowHeightMin <> 0 Then
                    If lngRowHeight < .RowHeightMin Then
                        lngRowHeight = .RowHeightMin
                    End If
                End If
                If .RowHeightMax <> 0 Then
                    If lngRowHeight > .RowHeightMax Then
                        lngRowHeight = .RowHeightMax
                    End If
                End If
                lngHeight = lngHeight + lngRowHeight
            Next
            gclsPros.CurrentForm.picCopy.Left = 0
            gclsPros.CurrentForm.picCopy.Top = lngHeight
        End If
        gclsPros.CurrentForm.picCopy.Visible = blnShow
        gclsPros.CurrentForm.picCopy.Enabled = blnShow
        gclsPros.CurrentForm.picCopy.ZOrder
    End With
End Sub

'vsKSS�¼�
Public Sub KSSAfterEdit(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsKSS_AfterEdit�¼�
    With vsKSS
        Call .Select(.Row, .Col)
    End With
End Sub

Public Sub KSSAfterRowColChange(ByRef vsKSS As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsKSS_AfterRowColChange�¼�
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    If lngNewCol = KI_����ҩ���� Then
        vsKSS.ColComboList(KI_����ҩ����) = "..."
    End If
End Sub

Public Sub KSSCellButtonClick(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long)
'vsKSS_CellButtonClick�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strSQLItem As String
    Dim bln���� As Boolean
    Dim vPoint As POINTAPI

    With vsKSS
        If LngCol = KI_����ҩ���� Then
'            If gclsPros.ReadPages Then
            If gclsPros.ShareMedRec Or gclsPros.FuncType = fҽ����ҳ Then
                bln���� = True
                strSQLItem = _
                    " From ������ĿĿ¼ A,ҩƷ���� B" & _
                    " Where A.ID=B.ҩ��ID And A.���='5' And A.������� IN(2,3) And Nvl(b.������, 0) <> 0" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) "
                strSql = "Select 0 as ĩ��,Max(Level) as ��ID,ID,�ϼ�ID,����,����,NULL as ��λ" & _
                    " From ���Ʒ���Ŀ¼ Where ����=1 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With ID In (Select A.����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID" & _
                    " Group by ID,�ϼ�ID,����,����"
                strSql = strSql & " Union ALL" & _
                    " Select 1 as ĩ��,1 as ��ID,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ" & _
                    strSQLItem & " Order By ĩ��,��ID Desc,����"
            Else
                strSql = "Select Rownum As ID, A.����, A.����, A.����" & vbNewLine & _
                    "From ������ҩ A"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            End If
            Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, IIf(bln����, 2, 0), "����ҩ��", False, "", "", False, True, False, vPoint.X, vPoint.Y, IIf(bln����, 0, .CellHeight), blnCancel, False, Not bln����)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û�п���ҩ�����ݿ���ѡ��", vbInformation, gstrSysName
                End If
            Else
                Call KSSSetDiagInput(vsKSS, LngRow, rsTmp)
                Call KSSEnterNextCell(vsKSS)
            End If
        End If
    End With
End Sub

Public Sub KSSKeyDown(ByRef vsKSS As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsKSS_KeyDown�¼�
    If vsKSS.Editable = flexEDNone Then Exit Sub
    If intKeyCode = vbKeyF4 Then
        Call zlCommFun.PressKey(vbKeySpace)
    ElseIf intKeyCode = vbKeyDelete Then
        If MsgBox("ȷʵҪɾ������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            With vsKSS
                .RemoveItem .Row
                If .Rows < 4 Then .Rows = 4: Call ChangeVSFHeight(vsKSS, True)
                Call SetKSSSerial
            End With
        End If
    ElseIf intKeyCode > 127 Then
        '���ֱ�����뺺�ֵ�����
        Call KSSKeyPress(vsKSS, intKeyCode)
    End If
End Sub

Public Sub KSSKeyPress(ByRef vsKSS As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsKSS_KeyPress�¼�
    With vsKSS
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            Call KSSEnterNextCell(vsKSS)
        ElseIf .Editable <> flexEDNone Then
            If intKeyAscii = Asc("*") Then
                intKeyAscii = 0
                Call KSSCellButtonClick(vsKSS, .Row, .Col)
            Else
                .ColComboList(KI_����ҩ����) = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Public Sub KSSKeyPressEdit(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyAscii As Integer)
'vsKSS_KeyPressEdit�¼�
    gclsPros.IsReturn = intKeyAscii = vbKeyReturn
    With vsKSS
        If LngCol = KI_ʹ������ Then
            If .EditSelLength <> 0 Then Exit Sub
            If intKeyAscii = vbKeyBack Then Exit Sub
            If Len(.EditText) > 18 Then intKeyAscii = 0
        ElseIf LngCol = KI_��ҩĿ�� Then
            If .EditSelLength <> 0 Then Exit Sub
            If intKeyAscii = vbKeyBack Then Exit Sub
            If LenB(StrConv(.EditText, vbFromUnicode)) >= 200 Then intKeyAscii = 0
        End If
    End With
End Sub

Public Sub KSSSetupEditWindow(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByVal lngEditWindow As Long, ByVal blnIsCombo As Boolean)
'vsKSS_SetupEditWindow�¼�
    With vsKSS
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Public Sub KSSValidateEdit(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsKSS_ValidateEdit�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI

    With vsKSS
        If LngCol = KI_����ҩ���� Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, LngRow, LngCol)
                If gclsPros.IsReturn Then Call KSSEnterNextCell(vsKSS)
            ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                If gclsPros.IsReturn Then Call KSSEnterNextCell(vsKSS)
            Else
                strInput = UCase(.EditText)
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "B.���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSql = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                End If

                If gclsPros.ShareMedRec Or gclsPros.FuncType = fҽ����ҳ Then
                    strSql = _
                        " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ" & _
                        " From ������ĿĿ¼ A,������Ŀ���� B,ҩƷ���� C" & _
                        " Where A.ID=B.������ĿID And A.ID=C.ҩ��ID And Nvl(c.������, 0) <> 0" & _
                        " And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " And A.���='5' And A.������� IN(2,3) And B.����=[3] And (" & strSql & ")" & _
                        " Order by A.����"
                Else
                    strSql = "Select Rownum As ID, A.����, A.����, A.����" & vbNewLine & _
                        "From ������ҩ A" & vbNewLine & _
                        "Where " & strSql & vbNewLine & _
                        "Order By ����"
                End If
                If zlCommFun.IsCharChinese(strInput) Then
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, strInput & "%", gclsPros.LikeString & strInput & "%", gclsPros.BriefCode + 1)
                    '�ж��Ƿ�������
                    If rsTmp.RecordCount = 0 Then
                        MsgBox "û���ҵ�ָ���Ŀ���ҩ�", vbInformation, gstrSysName
                        blnCancel = True: .EditText = "": Exit Sub
                    End If
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                    Call KSSSetDiagInput(vsKSS, LngRow, rsTmp)
                    .EditText = .Text
                    If gclsPros.IsReturn Then Call KSSEnterNextCell(vsKSS)
                Else
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "����ҩ��", _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gclsPros.LikeString & strInput & "%", gclsPros.BriefCode + 1)
                    If blnInputCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        blnCancel = True
                    Else
                        '�ж��Ƿ�������
                        If rsTmp Is Nothing Then
                            MsgBox "û���ҵ�ָ���Ŀ���ҩ�", vbInformation, gstrSysName
                            blnCancel = True: .EditText = "": Exit Sub
                        End If
                        Call KSSSetDiagInput(vsKSS, LngRow, rsTmp)
                        .EditText = .Text
                        If gclsPros.IsReturn Then Call KSSEnterNextCell(vsKSS)
                    End If
                End If
            End If
            gclsPros.IsReturn = False
        ElseIf LngCol = KI_ʹ������ Or LngCol = KI_DDD�� Then
            If (Not IsNumeric(.EditText) Or InStr(.EditText, "-") > 0 Or InStr(.EditText, "+") > 0) And .EditText <> "" Then
                MsgBox "��������Ч�����֡�", vbInformation, gstrSysName
                blnCancel = True
            Else
                If Len(.EditText) > 12 Then
                    MsgBox "������12λ���µ����֡�", vbInformation, gstrSysName
                    blnCancel = True
                    Exit Sub
                End If
            End If
        ElseIf LngCol = KI_ʹ�ý׶� Then
            '����û��޸��ˣ�����ȡ��ʱ��Ӱ����һ��
            If .Cell(flexcpData, LngRow, LngCol) = "����" Then .Cell(flexcpData, LngRow, LngCol) = ""
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'vsSpirit�¼�
Public Sub SpiritAfterRowColChange(ByRef vsSpirit As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsSpirit_AfterRowColChange�¼�
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    vsSpirit.FocusRect = flexFocusSolid
End Sub

Public Sub SpiritKeyDown(ByRef vsSpirit As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsSpirit_KeyDown�¼�
    Dim LngCol As Long
    If vsSpirit.Editable = flexEDNone Then Exit Sub
    With vsSpirit
        If intKeyCode = vbKeyDelete Then
            If MsgBox("���Ƿ����Ҫɾ�����еľ���ҩƷ��Ϣ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = .FixedRows Then
                For LngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, LngCol) = ""
                    .Cell(flexcpData, .Row, LngCol) = ""
                    .RowData(.Row) = 0
                Next
            Else
                .RemoveItem .Row
                Call ChangeVSFHeight(vsSpirit, True)
            End If
            zlControl.ControlSetFocus vsSpirit, True
        ElseIf intKeyCode <> vbKeyReturn Then
            Exit Sub
        Else
            If .Row = .Rows - 1 Then
                If Trim(.TextMatrix(.Row, SI_ҩ������)) = "" Then Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
            End If
        End If
    End With
End Sub

Public Sub SpiritKeyDownEdit(ByRef vsSpirit As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef intKeyCode As Integer, ByVal intShift As Integer)
'vsSpirit_KeyDownEdit�¼�
    Dim strInput As String, blnCancel As Boolean
    Dim rsTemp As Recordset, strSql As String
    Dim vPoint  As POINTAPI

    If intKeyCode <> vbKeyReturn Then Exit Sub
    intKeyCode = 0
    With vsSpirit
        If LngCol = SI_ҩ������ Then
            '��ѡ
            .Cell(flexcpData, LngRow, LngCol) = ""
            strInput = UCase(Trim(.EditText))
            If strInput = "" Then Exit Sub
            strInput = gclsPros.LikeString & strInput & "%"
            strSql = "" & _
                "   SELECT S.ҩƷid as ID, I.����, I.����, I.���, I.����, I.���㵥λ AS �ۼ۵�λ,S.��׼�ĺ�, S.��ʶ��,S.GMP��֤,I.����ʱ��, I.����ʱ�� " & _
                "   FROM �շ���ĿĿ¼ I, ҩƷ���  S,ҩƷ���� J  " & _
                "   WHERE I.ID=S.ҩƷid   and I.��� In ('5','6','7') And s.ҩ��id=J.ҩ��ID  And J.������� In ('����I��','����II��') " & _
                "           And (i.����ʱ�� Is Null Or to_char(i.����ʱ��,'yyyy-mm-dd')='3000-01-01') " & _
                "           And (i.���� like [1]  Or i.���� Like [1] Or   Exists(Select 1 From �շ���Ŀ���� Where i.Id=�շ�ϸĿid And ���� = 3))"
            vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTemp = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "ҩƷѡ����", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, False, strInput)
            If blnCancel = True Then
                Exit Sub
            End If
            If rsTemp Is Nothing Then
                MsgBox "û���ҵ������ҵĿ�����ҩ��", vbInformation, gstrSysName
                Exit Sub
            Else
                .EditText = NVL(rsTemp!����)
                .TextMatrix(.Row, SI_ҩ������) = NVL(rsTemp!����)
                .Cell(flexcpData, .Row, SI_ҩ������) = NVL(rsTemp!ID)
            End If
            If .TextMatrix(.Rows - 1, SI_ҩ������) <> "" Then
                .Rows = .Rows + 1
                Call ChangeVSFHeight(vsSpirit, True)
            End If
        End If
        Call EnterNextCellSpirit
    End With
End Sub

Public Sub SpiritKeyPress(ByRef vsSpirit As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsSpirit_KeyPress�¼�
    If vsSpirit.Editable = flexEDNone Then Exit Sub
    If intKeyAscii = vbKeyReturn Then
        intKeyAscii = 0
        Call EnterNextCellSpirit
    End If
End Sub

Public Sub SpiritStartEdit(ByRef vsSpirit As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsSpirit_StartEdit�¼�
    If Trim(vsSpirit.TextMatrix(LngRow, SI_ҩ������)) = "" And LngCol <> SI_ҩ������ Then
        blnCancel = True
    End If
End Sub

Public Sub SpiritValidateEdit(ByRef vsSpirit As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, ByRef blnCancel As Boolean)
'vsSpirit_ValidateEdit�¼�
    Dim intMax As Integer
    If LngCol >= SI_ҩ������ And LngCol <= SI_��Ч Then
        intMax = 50
        If LngCol = SI_ҩ������ Then
            intMax = 200
        ElseIf LngCol = SI_���ⷴӦ Then
            intMax = 100
        End If
        If LenB(StrConv(vsSpirit.EditText, vbFromUnicode)) > intMax Then
            MsgBox "����������ݲ��ܳ���" & intMax \ 2 & "�����֡�", vbInformation, gstrSysName
            blnCancel = True
            Exit Sub
        End If
    End If
End Sub

Public Sub TxtGotFocus(ByRef objTextBox As Object, Optional ByVal blnChineseIn As Boolean, Optional ByVal blnSetInfect As Boolean)
'���ܣ�ʵ�ֿؼ���ý���ȫѡ�ؼ����ݵĹ���
'������objTextBox=����Text���ԣ��Ҿ���Sel����������TextBox,ComboBox�ȿؼ�

    '��Ա�Լ��޶�������Ŀ�����人��
    zlCommFun.OpenIme blnChineseIn
    If gclsPros.PatiType = PF_סԺ And blnSetInfect Then
        'ҽԺ��Ⱦ��ؿؼ��ɼ����Լ�λ��
        If gclsPros.CurrentForm.picInfectInfo.Visible Then Call ShowInfectInfo(False)
    End If
    '�ؼ�����ѡ��
    Call zlControl.TxtSelAll(objTextBox)
End Sub

Public Sub ShowInfectInfo(Optional ByVal blnShow As Boolean = True, Optional ByRef objCtrl As Object, _
                            Optional ByVal lngLeft As Long, Optional ByVal lngTop As Long)
'���ܣ��Ը�Ⱦ��Ϣ��ʾ������λ�û�����
'������
'      blnShow=true,��ʾ��Ⱦ��Ϣ��false,���ظ�Ⱦ��Ϣ
'      objCtrl=�ؼ�����
'      lngLeft,lngTop=��Ⱦ��Ϣ��λ�ã��������ҳ��

    Dim blnExit As Boolean

    blnExit = True
    If gclsPros.FuncType = f������ҳ Or gclsPros.FuncType = fҽ����ҳ And gclsPros.PatiType <> PF_���� Then
        With gclsPros.CurrentForm
            '��ؿؼ�״̬����
            If .picInfectInfo.Visible = blnShow Then Exit Sub
            .picInfectInfo.Visible = blnShow
            .picInfectInfo.Top = lngTop - frmMain.Top - frmMain.PicForm.Top - .picMain.Top - .Top + 150
            .picInfectInfo.Left = lngLeft - frmMain.Left - .picMain.Left - frmMain.PicDirectory.Width - 200
            .picInfectInfo.ZOrder
        End With
    End If
End Sub

Private Sub TxtMouseDown(ByRef objText As Object, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'���ܣ�TextBox��Ĭ���Ҽ���Ϣ��ʾ���޸�
    If intButton = 2 And objText.Locked Then
        gclsPros.TXTProc = GetWindowLong(objText.hwnd, GWL_WNDPROC)
        Call SetWindowLong(objText.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub TxtMouseUp(ByRef objText As Object, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'���ܣ�TextBox��Ĭ���Ҽ���Ϣ��ʾ���޸�
    If intButton = 2 And objText.Locked Then
        Call SetWindowLong(objText.hwnd, GWL_WNDPROC, gclsPros.TXTProc)
    End If
End Sub

Private Sub GetDiagTypeScope(ByRef vsDiag As VSFlexGrid, ByVal lngType As Long, ByRef lngBgn As Long, ByRef lngEnd As Long)
'���ܣ���ȡ��ǰ������͵ķ�Χ
'������lngType=��ǰ����
'���أ�lngBgn=��ǰ���͵���ʼ��
'      lngEnd=��ǰ���ͽ�����
    Dim i As Long
    Dim LngCol As Long

    lngBgn = 0: lngEnd = 0
    With vsDiag
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, DI_��Ϸ���)) = lngType Then
                If lngBgn < .FixedRows Then lngBgn = i
                lngEnd = i
            End If
        Next
    End With
End Sub

Private Function DiagRowCanMove(ByVal intStep As Integer, ByVal lngType As Long, ByVal LngRow As Long) As Boolean
'���ܣ���������ƶ��ؼ�״̬
'������intStep=�ƶ�λ�ã�1-�����ƶ���-1�����ƶ�
'   lngType=��ǰ�������
'   lngRow=���ж����У�һ��Ϊ��ǰ��
    Dim lngBgn As Long, lngEnd As Long
    '���ݵ�ǰ�е�λ�������ƶ���Ͽؼ��Ŀ�����
    Call GetDiagTypeScope(IIf(lngType <= 10, gclsPros.CurrentForm.vsDiagXY, gclsPros.CurrentForm.vsDiagZY), lngType, lngBgn, lngEnd)
    If lngBgn = lngEnd Then 'ֻ��һ����ϣ��򲻿��ƶ�
        DiagRowCanMove = False
    ElseIf LngRow = lngBgn Then '��ǰ���Ǳ������һ�У���ֻ������
        DiagRowCanMove = intStep = 1 And gclsPros.OpenMode <> EM_���� And gclsPros.Module <> pסԺ��ʿվ
    ElseIf LngRow = lngEnd Then '��ǰ���Ǳ��������һ�У���ֻ����
        DiagRowCanMove = intStep = -1 And gclsPros.OpenMode <> EM_���� And gclsPros.Module <> pסԺ��ʿվ
    Else  '��ǰ���Ǳ������м�ĳһ�У�����������ƶ�
        DiagRowCanMove = gclsPros.OpenMode <> EM_���� And gclsPros.Module <> pסԺ��ʿվ
    End If
End Function

Private Sub MoveDiagRows(ByRef vsDiag As VSFlexGrid, ByVal intStep As Integer)
'���ܣ��ƶ������
'������vsDiag=��ǰ��ϱ��
'      intStep=�ƶ�λ�ã�1-�����ƶ���-1�����ƶ�
    Dim strTmp As String
    Dim i As Long, LngRow As Long
    Dim bln��ҽ As Boolean, bln�ֻ��̶� As Boolean
    Dim blnJudge As Boolean

    bln��ҽ = vsDiag.Name = "vsDiagXY"

    With vsDiag
        If Not DiagRowCanMove(intStep, Val(.TextMatrix(.Row, DI_��Ϸ���)), .Row) Then Exit Sub
        If .Row < 0 Then
            Exit Sub
        ElseIf gclsPros.FuncType <> f������ҳ Then '���ɱ༭��λ���йص����
            LngRow = IIf(intStep = 1, .Row, .Row + intStep)
            '������ɵĳ�Ժ��ϲ������
            If gclsPros.PathState = PS_�������� And gclsPros.PathOutTime Then
                If bln��ҽ Then
                    blnJudge = .TextMatrix(.Row, DI_��Ϸ���) = "��Ժ���" And gclsPros.InPath <= DT_��Ժ���XY
                Else
                    blnJudge = .TextMatrix(.Row, DI_��Ϸ���) = "��Ժ���" And gclsPros.InPath >= DT_�������ZY
                End If
                If blnJudge Then Exit Sub
            End If
        End If
        For i = .FixedCols To .Cols - 1
            '������������
            strTmp = .TextMatrix(.Row + intStep, i)
            .TextMatrix(.Row + intStep, i) = .TextMatrix(.Row, i)
            .TextMatrix(.Row, i) = strTmp
            '������������
            strTmp = .Cell(flexcpData, .Row + intStep, i)
            .Cell(flexcpData, .Row + intStep, i) = .Cell(flexcpData, .Row, i)
            .Cell(flexcpData, .Row, i) = strTmp
        Next
        '������������
        strTmp = .RowData(.Row + intStep)
        .RowData(.Row + intStep) = .RowData(.Row)
        .RowData(.Row) = Val(strTmp)
        Call SetDiagReletedInfo(vsDiag)
        .Row = .Row + intStep
    End With
End Sub

Private Function OPSRowCanMove(ByVal intStep As Integer, ByVal LngRow As Long) As Boolean
'���ܣ���������ƶ��ؼ�״̬
'������intStep=�ƶ�λ�ã�1-�����ƶ���-1�����ƶ�
'   lngType=��ǰ�������
'   lngRow=���ж����У�һ��Ϊ��ǰ��
    Dim lngBgn As Long, lngEnd As Long
    Dim vsOPS As VSFlexGrid

    '���ݵ�ǰ�е�λ�������ƶ���Ͽؼ��Ŀ�����
    Set vsOPS = gclsPros.CurrentForm.vsOPS
    lngBgn = vsOPS.FixedRows: lngEnd = vsOPS.Rows - 1
    If lngBgn = lngEnd Then 'ֻ��һ����ϣ��򲻿��ƶ�
        OPSRowCanMove = False
    ElseIf LngRow = lngBgn Then '��ǰ���Ǳ������һ�У���ֻ������
        OPSRowCanMove = intStep = 1 And gclsPros.OpenMode <> EM_���� And gclsPros.Module <> pסԺ��ʿվ
    ElseIf LngRow = lngEnd Then '��ǰ���Ǳ��������һ�У���ֻ����
        OPSRowCanMove = intStep = -1 And gclsPros.OpenMode <> EM_���� And gclsPros.Module <> pסԺ��ʿվ
    Else  '��ǰ���Ǳ������м�ĳһ�У�����������ƶ�
        OPSRowCanMove = gclsPros.OpenMode <> EM_���� And gclsPros.Module <> pסԺ��ʿվ
    End If
End Function

Private Sub MoveOPSRows(ByRef vsOPS As VSFlexGrid, ByVal intStep As Integer)
'���ܣ��ƶ������
'������vsDiag=��ǰ��ϱ��
'      intStep=�ƶ�λ�ã�1-�����ƶ���-1�����ƶ�
    Dim strTmp As String
    Dim i As Long, LngRow As Long
    Dim bln��ҽ As Boolean, bln�ֻ��̶� As Boolean
    Dim blnJudge As Boolean


    With vsOPS
        If Not OPSRowCanMove(intStep, .Row) Then Exit Sub
        If .Row < .FixedRows Then
            Exit Sub
        End If
        For i = .FixedCols To .Cols - 1
            '������������
            strTmp = .TextMatrix(.Row + intStep, i)
            .TextMatrix(.Row + intStep, i) = .TextMatrix(.Row, i)
            .TextMatrix(.Row, i) = strTmp
            '������������
            strTmp = .Cell(flexcpData, .Row + intStep, i)
            .Cell(flexcpData, .Row + intStep, i) = .Cell(flexcpData, .Row, i)
            .Cell(flexcpData, .Row, i) = strTmp
        Next
        '������������
        strTmp = .RowData(.Row + intStep)
        .RowData(.Row + intStep) = .RowData(.Row)
        .RowData(.Row) = Val(strTmp)
        .Row = .Row + intStep
    End With
End Sub

Private Sub OPSSetInput(ByRef vsOPS As VSFlexGrid, ByVal LngRow As Long, ByVal LngCol As Long, rsInput As ADODB.Recordset)
'���ܣ�������������������������ñ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim int�������� As Integer
    Dim blnOPSLevel As Boolean

    With vsOPS
        Select Case LngCol
            Case PI_��������, PI_��������
                If Not rsInput Is Nothing Then
                    int�������� = Val(gclsPros.OPSInput)
                    If gclsPros.CNIndent And gclsPros.FuncType = f������ҳ Then
                        '����:������໥�����ģ���������������ƣ����滻,һ�µ����,����ͬ������
                        If .TextMatrix(LngRow, PI_��������) = .Cell(flexcpData, LngRow, PI_��������) Or Trim(.TextMatrix(LngRow, PI_��������)) = "" Then
                            .TextMatrix(LngRow, PI_��������) = rsInput!����
                        End If
                    Else
                        .TextMatrix(LngRow, PI_��������) = rsInput!����
                    End If
                     .Cell(flexcpData, LngRow, PI_��������) = .TextMatrix(LngRow, PI_��������)
                    .TextMatrix(LngRow, PI_��������) = rsInput!����
                    .AutoSize PI_��������, PI_��������
                    If int�������� = 0 Then
                        .TextMatrix(LngRow, PI_������ĿID) = rsInput!ID
                        .TextMatrix(LngRow, PI_��������ID) = ""
                        strSql = "Select A.����ID as ID, Decode(B.��������, '��', '�ļ�����', '��', '��������', '��', '��������', '��', 'һ������', '�ļ�', '�ļ�����', '����', '��������', '����', '��������', 'һ��', 'һ������', Null) �������� From ������϶��� A, ��������Ŀ¼ B Where A.����ID = B.ID and  A.����ID=[1]"
                    Else
                        .TextMatrix(LngRow, PI_��������ID) = rsInput!ID
                        .TextMatrix(LngRow, PI_������ĿID) = ""
                        strSql = "Select A.����ID as ID, Decode(B.��������, '��', '�ļ�����', '��', '��������', '��', '��������', '��', 'һ������', '�ļ�', '�ļ�����', '����', '��������', '����', '��������', 'һ��', 'һ������', Null) �������� From ������϶��� A, ��������Ŀ¼ B Where A.����ID(+) = B.ID and  B.ID=[1]"
                    End If
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, Val(rsInput!ID))
                    If Not rsTmp.EOF Then
                        If int�������� = 0 Then
                            .TextMatrix(LngRow, PI_��������ID) = Val(rsTmp!ID)
                        Else
                            .TextMatrix(LngRow, PI_������ĿID) = Val(rsTmp!ID & "")
                        End If
                         If NVL(rsTmp!��������) <> "" Then .TextMatrix(LngRow, PI_��������) = NVL(rsTmp!��������)
                         blnOPSLevel = NVL(rsTmp!��������) <> ""
                    End If
                Else
                    .TextMatrix(LngRow, LngCol) = .EditText
                    .TextMatrix(LngRow, PI_��������ID) = ""
                    .TextMatrix(LngRow, PI_������ĿID) = ""
                End If
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)

                '����������ͬʱ��������������Ĭ������һ����ͬ
                If Not rsInput Is Nothing And LngRow > .FixedRows And LngRow = .Rows - 1 Then
                    If .TextMatrix(LngRow, PI_��������) = .TextMatrix(LngRow - 1, PI_��������) Then
                        .TextMatrix(LngRow, PI_����ҽʦ) = .TextMatrix(LngRow - 1, PI_����ҽʦ)
                        .TextMatrix(LngRow, PI_������ʿ) = .TextMatrix(LngRow - 1, PI_������ʿ)
                        .TextMatrix(LngRow, PI_����1) = .TextMatrix(LngRow - 1, PI_����1)
                        .TextMatrix(LngRow, PI_����2) = .TextMatrix(LngRow - 1, PI_����2)
                        .TextMatrix(LngRow, PI_����ʽ) = .TextMatrix(LngRow - 1, PI_����ʽ)
                        .TextMatrix(LngRow, PI_����ҽʦ) = .TextMatrix(LngRow - 1, PI_����ҽʦ)
                        .TextMatrix(LngRow, PI_�п�����) = .TextMatrix(LngRow - 1, PI_�п�����)
                        .TextMatrix(LngRow, PI_����ID) = .TextMatrix(LngRow - 1, PI_����ID)
                        .TextMatrix(LngRow, PI_��������) = .TextMatrix(LngRow - 1, PI_��������)

                        For i = PI_����ҽʦ To .Cols - 1
                            .Cell(flexcpData, LngRow, i) = .TextMatrix(LngRow, i)
                        Next
                    End If
                End If
                .Cell(flexcpData, LngRow, PI_��������) = IIf(blnOPSLevel, 1, 0)
                '������Ϸ������
                Call SetDiagMatchInfo(BCC_��ǰ������)
                '�����ʼ�ձ���һ����
                If LngRow = .Rows - 1 Then .AddItem "": Call ChangeVSFHeight(vsOPS, True)
            Case PI_����ʽ '����Ϊ������
                .TextMatrix(LngRow, LngCol) = rsInput!����
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
                .TextMatrix(LngRow, PI_����ID) = rsInput!ID
                .TextMatrix(LngRow, PI_��������) = NVL(rsInput!��������)
            Case PI_����ҽʦ, PI_������ʿ, PI_����1, PI_����2, PI_����ҽʦ
                .TextMatrix(LngRow, LngCol) = rsInput!����
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
            Case PI_�пڲ�λ
                .TextMatrix(LngRow, LngCol) = rsInput!����
                .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, LngCol)
        End Select
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetSign(ByVal intIndex As Integer, Optional ByVal blnUnSign As Boolean)
'���ܣ�ǩ����ȡ��ǩ��
'������intIndex=ǩ����ť����
'      blnUnSign=Fasle-ǩ����True-ȡ��ǩ��
    Dim strSql As String
    Dim rsTmp As Recordset, i As Long
    Dim bln���� As Boolean    '�Ƿ���д��������¼
    '˵����arrInfos��arrManIdxs��arrSgnIdxs���������Ԫ��һһ��Ӧ����Ա����ӵ͵���
    Dim arrInfos() As Variant '����ǩ������Ϣ��
    Dim arrManIdxs() As Variant 'ǩ����Ա�����б��Index
    Dim arrSgnIdxs() As Variant 'ǩ����ť��Index
    Dim blnSign As Boolean '��ǰ��ǩ��״̬
    Dim blnDiagnose As Boolean
    Dim strTmp As String, cboTmp As ComboBox

    With gclsPros.CurrentForm
        '�ж��Ƿ���������ǩ��
        If gintCA > 0 And CheckSign(1, 0, 0, gclsPros.��Ժ����ID, 2) Then
            If gobjESign Is Nothing Then
                On Error Resume Next
                Set gobjESign = CreateObject("zl9ESign.clsESign")
                Err.Clear: On Error GoTo 0
                If Not gobjESign Is Nothing Then
                    Call gobjESign.Initialize(gcnOracle, gclsPros.SysNo)
                End If
            End If
            If gobjESign Is Nothing Then
                MsgBox "����ǩ������δ����ȷ��װ��ǩ���������ܼ�����", vbInformation, gstrSysName
                Exit Sub
            Else
                If Not gobjESign.CheckCertificate(UserInfo.DBUser) Then Exit Sub
            End If
        End If

        arrInfos = Array("סԺҽʦ", "����ҽʦ", "����ҽʦ", "������")
        arrManIdxs = Array(MC_סԺҽʦ, MC_����ҽʦ, MC_���λ�����, MC_������)
        arrSgnIdxs = Array(SL_סԺҽʦ, SL_����ҽʦ, SL_����ҽʦ, SL_������)
        If blnUnSign Then
            '������鲡���Ƿ��Ŀ����ҳ��������״̬
            If Not CheckMecRed(gclsPros.����ID, gclsPros.��ҳID, .Caption, "ȡ��ǩ��") Then Exit Sub
        Else
            '������鲡���Ƿ��Ŀ����ҳ��������״̬
'            If Not CheckMecRed(gclsPros.����ID, gclsPros.��ҳID, .Caption, "ǩ��") Then Exit Sub
            '��Ҫȷ������ǩ���������
            For i = UBound(arrSgnIdxs) To LBound(arrSgnIdxs) Step -1
                If i = LBound(arrSgnIdxs) Then Exit For '������͵�סԺҽʦ
                If i <> UBound(arrSgnIdxs) Then
                    If .cboManInfo(arrManIdxs(i + 1)).Text = "" Then
                        If strTmp = "" Then
                             strTmp = "û��ȷ��" & arrInfos(i + 1)
                             Set cboTmp = .cboManInfo(arrManIdxs(i + 1))
                        Else
                             strTmp = strTmp & "��" & arrInfos(i + 1)
                        End If
                    End If
                End If
            Next
            If strTmp <> "" Then
                Call ShowMessage(cboTmp, strTmp & "��")
                Exit Sub
            End If
            On Error GoTo errH
            '�����������¼������ʾ�Ƿ����
            bln���� = False
            For i = 1 To .vsOPS.Rows - 1
                If Trim(.vsOPS.TextMatrix(i, PI_��������)) <> "" Then
                    bln���� = True
                End If
            Next

            strSql = "Select Count(1) As ���� From ����ҽ����¼ Where ����ID=[1] And ��ҳID=[2] And ҽ��״̬=8 And �������='F'"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gclsPros.CurrentForm.Caption, gclsPros.����ID, gclsPros.��ҳID)
            If Val(rsTmp!���� & "") > 0 And Not bln���� Then
                .vsOPS.Row = .vsOPS.FixedRows: .vsOPS.Col = PI_��������
                If ShowMessage(.vsOPS, "�ò��˴�������ҽ��������ҳ��û�����������¼���Ƿ������", True) = vbNo Then Exit Sub
            End If

            'ǩ��ǰ�Զ�����
            If Not CheckMedPageData(blnDiagnose) Then
                gclsPros.IsCheckData = False
                Exit Sub
            End If

            If Not gclsMain.IsDiagInput And Not blnDiagnose And gclsPros.MustDiagType <> "" Then
                If MsgBox("Ҫ��������Ϣ��û�����룬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            If Not SaveMedPageData() Then Exit Sub
        End If

        For i = LBound(arrSgnIdxs) To UBound(arrSgnIdxs)
            If arrSgnIdxs(i) = intIndex Then
                strSql = "Zl_������ҳ�ӱ�_��ҳ����(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",'" & arrInfos(i) & "ǩ��'," & IIf(blnUnSign, "Null", "'" & UserInfo.���� & "'") & ")"
                Exit For
            End If
        Next
        Call zlDatabase.ExecuteProcedure(strSql, gclsPros.CurrentForm.Caption)
        Set gclsPros.AuxiInfo = GetPatiAuxiInfoData(gclsPros.����ID, gclsPros.��ҳID)
        blnSign = gclsPros.IsSigned
        gclsPros.IsSigned = SetSignature
        If blnSign And Not gclsPros.IsSigned Then
            Call SetFaceInit(True)
        End If
        Call SetFaceEditable(gclsPros.IsSigned)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function PageOperate(ByVal mopType As MedRec_Operate, Optional ByVal intPage As Integer) As Boolean
'���ܣ���ӡ��Ԥ����ҳ����ȷ��������ҳ
'������intType=2����ӡ����=1��Ԥ����0=���ã�3���ȷ����ť������ҳ
'      intPage=1-4��ӡ��ҳ������ʽ��=5��ӡ����+��ҳ1��=6��ӡ����+��ҳ2
'      blnSavePage=True-���ȷ����ť�Ƿ񱣴���ҳ ��False-��ӡ��Ԥ����ҳ
'���أ��Ƿ�ɹ�
    Dim blnDiagnose As Boolean
    Dim blnPagePrint As Boolean, intPrint As Integer

    If mopType = MOP_ȷ�� And (gclsPros.FuncType = fҽ����ҳ And gclsPros.IsSigned Or gclsPros.OpenMode = EM_����) Then
        gclsPros.IsOK = True
        PageOperate = True
        gclsPros.IsDiagChange = Not blnDiagnose
        Exit Function
    End If

    '������ҳ��Ȼ��ǩ������״̬��סԺ��ҳ����ǩ������״̬�ű���
    If gclsPros.OpenMode <> EM_���� And Not gclsPros.IsSigned Then
        If gclsPros.FuncType = f������ҳ Then
            If Not ValidatePageNos(True) Then Exit Function
        End If
        If gclsPros.FuncType = fҽ����ҳ And gclsPros.PatiInfo!�������� = 1 Then
            If Not Check���� Then
                gclsPros.IsCheckData = False
                Exit Function
            End If
        Else
            If Not CheckMedPageData(blnDiagnose) Then
                gclsPros.IsCheckData = False
                Exit Function
            End If
        End If
        If Not gclsMain.IsDiagInput And Not blnDiagnose And gclsPros.MustDiagType <> "" Then
            If MsgBox("Ҫ��������Ϣ��û�����룬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If

        If Not SaveMedPageData() Then Exit Function
        '������ҳ��Ƭ����
        If gclsPros.FuncType = fҽ����ҳ And gclsPros.PatiType = PF_���� Then
            Call SavePatPicture(gclsPros.����ID)
        End If
    End If
    If gclsPros.FuncType = f������ҳ And mopType = MOP_ȷ�� Then
        If gclsPros.OpenMode = EM_�������� Or gclsPros.OpenMode = EM_������ҳ Then
            If Val(zlDatabase.GetPara("¼�벡�����λ��", gclsPros.SysNo, gclsPros.Module)) = 1 Then
                Call gclsMain.MedRecSaveLocation(gclsPros.����ID, gclsPros.��ҳID)
            End If
        End If
    End If
    '�޸���ҳ�򲡰����ȷ�����˳�
    If mopType = MOP_ȷ�� And gclsPros.OpenMode = EM_�༭ Then
        If gclsPros.FuncType = f������ҳ Then Call gclsMain.SavePage(gclsPros.����ID, gclsPros.��ҳID)
        gclsPros.IsOK = True
        PageOperate = True
        gclsPros.IsDiagChange = Not blnDiagnose
    End If
    ' סԺ��ҳ��ӡ��ҳ��������ҳ��ӡ������ҳ���棬������ҳֻ��ȷ��һ��״̬
    If gclsPros.FuncType = f������ҳ Or mopType <> MOP_ȷ�� And gclsPros.FuncType <> f������ҳ Then
        If gclsPros.FuncType <> f������ҳ Then
            blnPagePrint = True
        Else
            blnPagePrint = InStr(gclsPros.Privs, "��������ӡ") > 0
            If blnPagePrint Then
                intPrint = Val(zlDatabase.GetPara("������������Ƥ��ӡ", gclsPros.SysNo, gclsPros.Module))
                blnPagePrint = intPrint <> 0
                If blnPagePrint And intPrint = 2 Then
                    blnPagePrint = MsgBox("�Ƿ��ӡ���˲�����������Ƥ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
                End If
            End If
        End If
        If blnPagePrint Then
            Call PrintInMedRec(mopType, gclsPros.����ID, gclsPros.��ҳID, gclsPros.��Ժ����ID, intPage)
        End If
    End If
    '�����������룬��Ҫ����������ݣ����³�ʼ��,����������������ҳ���ڲ�������
    If gclsPros.OpenMode = EM_�������� Or gclsPros.OpenMode = EM_������ҳ Then
        Call gclsMain.SavePage(gclsPros.����ID, gclsPros.��ҳID)
        '�������������,������һ���û�
        gclsPros.InNo = ""
        Call ClearPageContent
        Call SetAllVSF(True)
        Call ChangePage(, 0)
    End If
    PageOperate = True
End Function

Private Sub KSSSetDiagInput(ByRef vsKSS As VSFlexGrid, ByVal LngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�����������Ŀ������
    With vsKSS
        If Not rsInput Is Nothing Then
            .TextMatrix(LngRow, KI_����ҩ����) = NVL(rsInput!����)
            .RowData(LngRow) = Val(rsInput!ID)
        Else
            .TextMatrix(LngRow, KI_����ҩ����) = .EditText
        End If
        .Cell(flexcpData, LngRow, KI_����ҩ����) = .TextMatrix(LngRow, KI_����ҩ����)
    End With
End Sub

Private Sub KSSEnterNextCell(ByRef vsKSS As VSFlexGrid)
    With vsKSS
        If .Row = .Rows - 1 Then
            If .TextMatrix(.Row, KI_����ҩ����) = "" Then
                Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True: Exit Sub
            ElseIf .Editable = flexEDNone Then
                Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True: Exit Sub
            Else
                .AddItem ""
                Call ChangeVSFHeight(vsKSS, True)
                Call SetKSSSerial
            End If
        End If

        If .Col = KI_������ҩ Then
            .Col = .FixedCols
            .Row = .Row + 1
        Else
            .Col = .Col + 1
        End If
        If .Row >= 0 And .Row < .Rows - 1 And .Col >= 0 And .Row < .Cols - 1 Then
            .ShowCell .Row, .Col
        End If
    End With
End Sub


Public Function AddOrDelFreeCols(ByRef vsFree As VSFlexGrid, ByVal strFreeNames As String, ByVal strFreeNum As String, ByVal blnAdd As Boolean) As Boolean
'���ܣ�ɾ����������ָ������
'������vsFree=���ñ��
'      strFreeNames=������
'      blnAdd=true-�������ã�False-ɾ������
'      strFreeNum=��������
'���أ��Ƿ��ҵ��÷���
    Dim LngRow As Long, LngCol As Long, i As Long
    Dim blnFind As Boolean, j As Long
    Dim lngPreRow As Long, lngPreCol As Long

    With vsFree
        If blnAdd Then
            If .TextMatrix(.Row, IIf(.Col Mod 2 = 0, .Col, .Col + 1)) = strFreeNames And .TextMatrix(.Row, IIf(.Col Mod 2 = 0, .Col + 1, .Col)) <> "" Then
                If Not gclsPros.SameName Then
                    MsgBox "�÷�����¼�룬����ѡһ�֡�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            '������һ����δ¼���λ��
            For i = (.Rows - 1) * 3 To .Rows * 3 - 1
                LngRow = i \ 3: LngCol = (i Mod 3) * 2
                If .TextMatrix(LngRow, LngCol) = "" Or .TextMatrix(LngRow, LngCol + 1) = "" Then
                    Exit For
                ElseIf LngCol = 4 Then '���һ���Ѿ���д,������һ��
                    .Rows = .Rows + 1: LngRow = LngRow + 1: LngCol = 0: Call ChangeVSFHeight(vsFree, True): Exit For
                End If
            Next
            .TextMatrix(LngRow, LngCol) = strFreeNames
            .TextMatrix(LngRow, LngCol + 1) = strFreeNum
        Else
            '�ҵ���Ҫɾ����λ��
            For i = 3 To .Rows * 3 - 1
                LngRow = i \ 3: LngCol = (i Mod 3) * 2
                If .TextMatrix(LngRow, LngCol) = strFreeNames And .TextMatrix(LngRow, LngCol + 1) = strFreeNum Then
                    blnFind = True: Exit For
                End If
            Next
            '����λ�ú�ķ���ȫ����ǰ�ƶ�
            If blnFind Then
                For j = i + 1 To .Rows * 3 - 1
                    LngRow = j \ 3: LngCol = (j Mod 3) * 2
                    lngPreRow = (j - 1) \ 3: lngPreCol = ((j - 1) Mod 3) * 2
                    .TextMatrix(lngPreRow, lngPreCol) = .TextMatrix(LngRow, LngCol)
                    .TextMatrix(lngPreRow, lngPreCol + 1) = .TextMatrix(LngRow, LngCol + 1)
                Next
                '�����ڶ��У����һ��û����д�����Ƴ����һ��
                If .Rows > 2 Then
                    If .TextMatrix(.Rows - 2, 4) = "" Then
                        .Rows = .Rows - 1
                        Call ChangeVSFHeight(vsFree, True)
                    End If
                End If
            End If
        End If
        Call SumAndSetFrees
    End With

    AddOrDelFreeCols = True
End Function

Public Sub SetDiagInput(ByRef vsDiagTmp As VSFlexGrid, ByVal LngRow As Long, rsInput As ADODB.Recordset, Optional bln���� As Boolean)
'���ܣ����������Ŀ������
'      bln����=�Ƿ��Ǹ�������
    Dim str�Ա� As String
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long, j As Long
    Dim strTmp As String, bln�ֻ��̶� As Boolean
    Dim bln��ҽ As Boolean, blnRCodeIn As Boolean
    Dim lngTmpRow As Long, lng��ԺRow As Long
    Dim lngԭ���ID As Long, int��ϴ��� As Integer
    Dim blnSame������� As Boolean
    Dim rs���� As New ADODB.Recordset
    Dim rsOutPut As ADODB.Recordset
    Dim blnGet���� As Boolean

    blnGet���� = gclsPros.GetExtraCode
    With vsDiagTmp
        bln��ҽ = .Name = "vsDiagXY"
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If gclsPros.FuncType = f������ҳ Then
                    '���������߼���֪ʲôԭ����ʱ����,���������켲���������������
                    If rsInput!���� Like "R*" Then
                        If blnRCodeIn Then
                            Exit For
                        Else
                            If MsgBox("��������ʹ��R������Ϊ��Ҫ���룬�Ƿ����룿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                                blnRCodeIn = True: Exit For
                            End If
                        End If
                    End If
                End If
                '���ǵ����ĸ�������
                If Not bln���� Then
                    If i > 1 Then
                        '���һ����������ҽ��������ϣ���ҽ�������ж���ѡ�����ʱ�Ĵ���
                        lngԭ���ID = 0
                        If LngRow = .Rows - 1 Then
                            .Rows = .Rows + 1
                            Call ChangeVSFHeight(vsDiagTmp, True)
                            .TextMatrix(.Rows - 1, DI_��Ϸ���) = .TextMatrix(LngRow, DI_��Ϸ���)
                            If gclsPros.PatiType = PF_���� Then .TextMatrix(.Rows - 1, DI_�������) = .TextMatrix(LngRow, DI_�������)
                        End If
                        'ȷ����ǰ��ʾ��
                        If Val(.TextMatrix(LngRow + 1, DI_��Ϸ���)) = Val(.TextMatrix(LngRow, DI_��Ϸ���)) Then
                            For j = LngRow + 1 To .Rows - 1
                                If Val(.TextMatrix(j, DI_��Ϸ���)) = Val(.TextMatrix(LngRow, DI_��Ϸ���)) Then
                                    LngRow = j
                                    If .TextMatrix(j, DI_�������) = "" Then Exit For
                                Else
                                    Exit For
                                End If
                            Next
                            If .TextMatrix(LngRow, DI_�������) <> "" Then
                                LngRow = LngRow + 1: .AddItem "", LngRow
                                Call ChangeVSFHeight(vsDiagTmp, True)
                                .TextMatrix(LngRow, DI_��Ϸ���) = .TextMatrix(LngRow - 1, DI_��Ϸ���)
                                If gclsPros.PatiType = PF_���� Then .TextMatrix(LngRow, DI_�������) = .TextMatrix(LngRow - 1, DI_�������)
                            End If
                        Else
                            LngRow = LngRow + 1: .AddItem "", LngRow
                            Call ChangeVSFHeight(vsDiagTmp, True)
                            .TextMatrix(LngRow, DI_��Ϸ���) = .TextMatrix(LngRow - 1, DI_��Ϸ���)
                            If gclsPros.PatiType = PF_���� Then .TextMatrix(LngRow, DI_�������) = .TextMatrix(LngRow - 1, DI_�������)
                        End If
                    Else
                        lngԭ���ID = Val(.TextMatrix(LngRow, DI_���ID))
                    End If

                    .TextMatrix(LngRow, DI_��ϱ���) = rsInput!���� & ""
                    If gclsPros.CNIndent And gclsPros.FuncType = f������ҳ Then
                        '����:������໥�����ģ��������������������滻
                        'һ�µ����,����ͬ�����¼���
                        If .TextMatrix(LngRow, DI_�������) = .Cell(flexcpData, LngRow, DI_�������) Or Trim(.TextMatrix(LngRow, DI_�������)) = "" Then
                            .TextMatrix(LngRow, DI_�������) = rsInput!����
                        End If
                    Else
                        .TextMatrix(LngRow, DI_�������) = rsInput!����
                    End If
                    .Cell(flexcpData, LngRow, DI_�������) = rsInput!���� & ""  '����ԭ��
                    .Cell(flexcpData, LngRow, DI_��ϱ���) = rsInput!���� & ""
                    .AutoSize DI_��ϱ���, DI_�������
                    If .ColWidth(DI_�������) < 3200 Then
                        .ColWidth(DI_�������) = 3200
                    End If
                    If gclsPros.FuncType = f���ѡ�� Then .TextMatrix(LngRow, DI_����) = 1
                    .TextMatrix(LngRow, DI_���ID) = rsInput!���ID & ""
                    .TextMatrix(LngRow, DI_����ID) = rsInput!����id & ""
                    .TextMatrix(LngRow, DI_��Ч����) = rsInput!��Ч���� & ""
                    .TextMatrix(LngRow, DI_������Ϣ) = IIf(Val(rsInput!���� & "") = 1, "1", "")
                    .TextMatrix(LngRow, DI_�Ƿ���) = IIf(Val(rsInput!�Ƿ��� & "") = 1, "1", "")
                    .TextMatrix(LngRow, DI_��������) = rsInput!�������� & ""
                    .TextMatrix(LngRow, DI_�������) = rsInput!������� & ""
                    '����֢��Ժ�ڸ�Ⱦ��Ժ���Ĭ��Ϊ��
                    If Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_����֢ Or Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_Ժ�ڸ�Ⱦ Then
                        .TextMatrix(LngRow, DI_��Ժ����) = "��"
                    End If
                    If blnGet���� Then
                        If Not IsNull(rsInput!����) Then
                            Set rsTmp = GetDiagExtraID(rsInput!���� & "")
                            If rsTmp.RecordCount > 0 Then
                                .TextMatrix(LngRow, DI_����ID) = rsTmp!ID & ""
                            Else
                                .TextMatrix(LngRow, DI_����ID) = ""
                            End If
                        End If
                        .TextMatrix(LngRow, DI_ICD����) = IIf(bln����, rsInput!���� & "", rsInput!���� & "")
                        .Cell(flexcpData, LngRow, DI_ICD����) = .TextMatrix(LngRow, DI_ICD����)
                    End If
                    '��������˲���ICD���������дʱ��¼���C00-D48�򵯳�Ҫ��¼��������̬ѧ���룻
                    If gclsPros.CheckICD���� = 1 And (Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_��Ժ���XY) _
                        And (InStr("C", Left(.TextMatrix(LngRow, DI_��ϱ���), 1)) > 0 Or (InStr("D", Left(.TextMatrix(LngRow, DI_��ϱ���), 1)) > 0 And Val(Mid(.TextMatrix(LngRow, DI_��ϱ���), 2, 2)) <= 48)) And Left(.TextMatrix(LngRow, DI_��ϱ���), 1) <> "" Then
                        If frmZLInPut.ShowMe(gclsPros.CurrentForm, "   ���[" & .Cell(flexcpData, LngRow, DI_�������) & "]Ϊ������ϣ�������������̬ѧ���룡", rsOutPut) Then
                            Call SetDiagInput(vsDiagTmp, LngRow, rsOutPut, True)
                        End If
                    End If
                Else
                    .TextMatrix(LngRow, DI_����ID) = rsInput!��ĿID & ""
                    .TextMatrix(LngRow, DI_ICD����) = rsInput!���� & ""
                    .Cell(flexcpData, LngRow, DI_ICD����) = .TextMatrix(LngRow, DI_ICD����)
                End If
                
             
                
                '������ҳ����Ժ��Ҫ����滻������Ҫ�������Ժ��Ҫ���
                If gclsPros.FuncType = f������ҳ Then
                    '�����Ժ��ϲ����踽��,��ҽ�����ø���
                    If Val(.TextMatrix(LngRow, DI_��Ϸ���)) <> DT_�������XY And Val(.TextMatrix(LngRow, DI_��Ϸ���)) <> DT_��Ժ���XY And bln��ҽ Then
                        '��Ժ��ϸ�����V,W,X,Y,�����������ж�ԭ��,��δ�������ر�������켲������
                        If Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_��Ժ���XY And Val(.TextMatrix(LngRow, DI_����ID)) <> 0 Then
                            If InStr("VWXY", Left(rsInput!���� & "", 1)) > 0 Then
                                If gclsPros.Sex Like "*��*" Then
                                    str�Ա� = "��"
                                ElseIf gclsPros.Sex Like "*Ů*" Then
                                    str�Ա� = "Ů"
                                End If

                                strSql = "Select A.Id,A.Id As ��Ŀid, A.����, A.���, A.����, D.ID ����ID, D.���� ��������, A.����, A.˵��, Null ����, A.����id, " & IIf(gclsPros.BriefCode = 0, "A.����", "A.�����") & " as ����, A.��Ч����, A.����, C.�Ƿ���,A.���� ��������, Null ����id,A.��� �������, Null ���id" & vbNewLine & _
                                        "From ��������Ŀ¼ A, ����������� C, ��������Ŀ¼ D " & vbNewLine & _
                                        "Where A.ID=[1] And A.����=D.����(+)  And A.����id = C.Id(+)" & vbNewLine & _
                                        "  And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & IIf(str�Ա� <> "", " And (A.�Ա�����=[2] Or A.�Ա����� is Null) ", " ")
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��Ժ��ϸ����Ӧ�������ж���", Val(.TextMatrix(LngRow, DI_����ID)), str�Ա�)
                                '��λ�����ж��Ŀ��У�û��������һ��
                                lngTmpRow = 0
                                For j = FindDiagRow(DT_�����ж���) To .Rows - 1
                                    If .TextMatrix(j, DI_�������) = "" Then
                                        lngTmpRow = j
                                        Exit For
                                    End If
                                Next
                                If lngTmpRow = 0 Then .Rows = .Rows + 1: lngTmpRow = .Rows - 1: Call ChangeVSFHeight(vsDiagTmp, True)
                                '���������ж�
                                Call SetDiagInput(vsDiagTmp, lngTmpRow, rsTmp)
                            End If
                        End If
                    End If

                    lng��ԺRow = FindDiagRow(IIf(bln��ҽ, DT_��Ժ���XY, DT_��Ժ���ZY))
                    If LngRow = lng��ԺRow Then
                        '�滻�������
                        lngTmpRow = FindDiagRow(IIf(bln��ҽ, DT_�������XY, DT_�������ZY))
                        If .TextMatrix(lngTmpRow, DI_����ID) = "" Then
                             '����:������໥�����ģ��������������������滻
                            If gclsPros.CNIndent And Trim(.TextMatrix(lngTmpRow, DI_�������)) = "" Or Not gclsPros.CNIndent Then
                                .TextMatrix(lngTmpRow, DI_����ID) = .TextMatrix(LngRow, DI_����ID)
                                .TextMatrix(lngTmpRow, DI_��ϱ���) = .TextMatrix(LngRow, DI_��ϱ���)
                                .TextMatrix(lngTmpRow, DI_�������) = .TextMatrix(LngRow, DI_�������)
                                .Cell(flexcpData, lngTmpRow, DI_�������) = .Cell(flexcpData, LngRow, DI_�������)
                                .Cell(flexcpData, lngTmpRow, DI_��ϱ���) = .Cell(flexcpData, LngRow, DI_��ϱ���)
                                '������������Ϣ
                                Call SetDiagReletedInfo(vsDiagTmp, lngTmpRow)
                            End If
                            .TextMatrix(lngTmpRow, DI_��ע) = .TextMatrix(LngRow, DI_��ע)
                        End If
                        '�滻��Ժ���
                        lngTmpRow = FindDiagRow(IIf(bln��ҽ, DT_��Ժ���XY, DT_��Ժ���ZY))
                        If .TextMatrix(lngTmpRow, DI_����ID) = "" Then
                             '����:������໥�����ģ��������������������滻
                            If gclsPros.CNIndent And Trim(.TextMatrix(lngTmpRow, DI_�������)) = "" Or Not gclsPros.CNIndent Then
                                .TextMatrix(lngTmpRow, DI_����ID) = .TextMatrix(LngRow, DI_����ID)
                                .TextMatrix(lngTmpRow, DI_��ϱ���) = .TextMatrix(LngRow, DI_��ϱ���)
                                .TextMatrix(lngTmpRow, DI_�������) = .TextMatrix(LngRow, DI_�������)
                                .Cell(flexcpData, lngTmpRow, DI_�������) = .Cell(flexcpData, LngRow, DI_�������)
                                .Cell(flexcpData, lngTmpRow, DI_��ϱ���) = .Cell(flexcpData, LngRow, DI_��ϱ���)
                                '������������Ϣ
                                Call SetDiagReletedInfo(vsDiagTmp, lngTmpRow)
                            End If
                            .TextMatrix(lngTmpRow, DI_��ע) = .TextMatrix(LngRow, DI_��ע)
                        End If

                        'ɾ�������ж�ԭ��
                        If bln��ҽ And Not .TextMatrix(LngRow, DI_��ϱ���) Like "S*" And Not .TextMatrix(LngRow, DI_��ϱ���) Like "T*" Then
                            lngTmpRow = FindDiagRow(DT_�����ж���)
                            If lngTmpRow < .Rows - 1 Then
                                .Rows = lngTmpRow + 1
                                Call ChangeVSFHeight(vsDiagTmp, True)
                            End If
                            .Cell(flexcpText, lngTmpRow, .FixedCols, lngTmpRow, .Cols - 1) = ""
                            .Cell(flexcpData, lngTmpRow, .FixedCols, lngTmpRow, .Cols - 1) = ""
                            .TextMatrix(lngTmpRow, DI_��Ϸ���) = DT_�����ж���
                            .RowData(lngTmpRow) = 0
                        End If
                    ElseIf Val(.TextMatrix(LngRow, DI_��Ϸ���)) = DT_�����ж��� Then
                        If .TextMatrix(LngRow, DI_��Ժ���) = "" Then
                            .TextMatrix(LngRow, DI_��Ժ���) = .TextMatrix(lng��ԺRow, DI_��Ժ���)
                        End If
                    End If
                End If

                If Not bln��ҽ Then
                    '��ҽ���ݼ�����ϲο�ȡ֤��
                    Call Set��ҽ֤��(LngRow, Val(.TextMatrix(LngRow, DI_���ID)))
                End If
                If gclsPros.FuncType <> f������ҳ Then
                    If CreatePlugInOK(IIf(gclsPros.PatiType = PF_����, p����ҽ��վ, pסԺҽ��վ)) Then
                        int��ϴ��� = 0
                        If gclsPros.PatiType = PF_סԺ Then
                            For j = .FixedRows To LngRow
                                If .TextMatrix(j, DI_��Ϸ���) = .TextMatrix(LngRow, DI_��Ϸ���) Then
                                    int��ϴ��� = int��ϴ��� + 1
                                End If
                            Next
                        Else
                            int��ϴ��� = IIf(LngRow = .FixedRows, -1, -2)
                        End If
                        On Error Resume Next
                        Select Case int��ϴ���
                            Case -1
                                Call gobjPlugIn.DiagnosisEnter(gclsPros.SysNo, p����ҽ��վ, gclsPros.����ID, gclsPros.��ҳID, Val(rsInput!��ĿID), .TextMatrix(LngRow, DI_�������), lngԭ���ID)
                                Call zlPlugInErrH(Err, "DiagnosisEnter")
                            Case -2
                                Call gobjPlugIn.DiagnosisOtherEnter(gclsPros.SysNo, p����ҽ��վ, gclsPros.����ID, gclsPros.��ҳID, Val(rsInput!��ĿID), .TextMatrix(LngRow, DI_�������), lngԭ���ID)
                                Call zlPlugInErrH(Err, "DiagnosisOtherEnter")
                            Case Else
                                Call gobjPlugIn.DiagnosisEnterIn(gclsPros.SysNo, pסԺҽ��վ, gclsPros.����ID, gclsPros.��ҳID, Val(rsInput!��ĿID), .TextMatrix(LngRow, DI_�������), lngԭ���ID, _
                                    IIf(gclsPros.Is��ʿվ, 1, 0), .TextMatrix(LngRow, DI_��Ϸ���), int��ϴ���)
                                Call zlPlugInErrH(Err, "DiagnosisEnterIn")
                        End Select
                        Err.Clear: On Error GoTo errH
                    End If
                End If
                rsInput.MoveNext
            Next
        Else
            If Not bln���� Then
                If gclsPros.CNIndent And gclsPros.FuncType = f������ҳ Or gclsPros.FuncType <> f������ҳ Then
                    .TextMatrix(LngRow, DI_�������) = .EditText
                    If gclsPros.FuncType <> f������ҳ Then
                        .Cell(flexcpData, LngRow, DI_�������) = .TextMatrix(LngRow, DI_�������)
                        .TextMatrix(LngRow, DI_��ϱ���) = ""
                         .Cell(flexcpData, LngRow, DI_�������) = ""
                        .TextMatrix(LngRow, DI_���ID) = ""
                        .TextMatrix(LngRow, DI_����ID) = ""
                        .TextMatrix(LngRow, DI_֤��ID) = ""
                    End If
                End If
            Else
                .TextMatrix(LngRow, DI_�̶�����) = ""
                .TextMatrix(LngRow, DI_ICD����) = ""
                .Cell(flexcpData, LngRow, DI_ICD����) = ""
                .TextMatrix(LngRow, DI_����ID) = ""
            End If
        End If
        .Cell(flexcpForeColor, .FixedRows, DI_�Ƿ�����, .Rows - 1, DI_�Ƿ�����) = vbRed
        .Cell(flexcpBackColor, .FixedRows, DI_��ϱ���, .Rows - 1, DI_��ϱ���) = GRD_UNEDITCELL_COLOR      '����ɫ
        '������Ϸ������
        If Not (gclsPros.PatiType = PF_���� And gclsPros.FuncType = f���ѡ��) Then
            Call SetDiagReletedInfo(vsDiagTmp, LngRow)
        End If
        If gclsPros.FuncType <> f���ѡ�� Then
            '������������Ϣ
            If gclsPros.Module = p����ҽ��վ Then
                If gclsPros.CurrentForm.optState(OP_����).Value = False Then
                    If PatiReSeeDoctor Then
                        If MsgBox("���˾�����ҡ�ҽ����������ϴ���ͬ��Ҫ���Ϊ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                            gclsPros.CurrentForm.optState(OP_����).Value = True
                        End If
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function Set��ҽ֤��(ByVal LngRow As Long, ByVal lng���ID As Long, Optional ByVal rsInput As Recordset, Optional ByVal blnFreeInput As Boolean) As Boolean
'���ܣ���ҽ���ݼ�����ϲο�ȡ֤��
'������rsInput-�����Ϊ�գ������ָ������ҩ֤���¼��
'      blnFreeInput  true - ����¼��
'���أ��Ƿ��ж�Ӧ��ϵ
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim strTmp As String

    On Error GoTo errH

    With gclsPros.CurrentForm.vsDiagZY
        If blnFreeInput Then
            .TextMatrix(LngRow, DI_֤��ID) = ""
            .TextMatrix(LngRow, DI_֤�����) = ""
            .TextMatrix(LngRow, DI_��ҽ֤��) = .EditText
            .Cell(flexcpData, LngRow, DI_��ҽ֤��) = .TextMatrix(LngRow, DI_��ҽ֤��)
        Else
            If rsInput Is Nothing Then
                If lng���ID = 0 Then Exit Function
                strSql = "Select Distinct A.֤����� As ID, A.֤��id As ��Ŀid, B.����, B.����, A.֤������ ����," & IIf(gclsPros.BriefCode = 0, "B.����", "B.����� As ����") & ", B.˵��" & vbNewLine & _
                            "From ������ϲο� A, ��������Ŀ¼ B" & vbNewLine & _
                            "Where A.֤��id = B.Id(+) And A.���id = [1] And A.֤������ Is Not Null" & vbNewLine & _
                            "Order By A.֤�����"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsInput = zlDatabase.ShowSQLSelect(gclsPros.CurrentForm, strSql, 0, "��ҽ֤��", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng���ID)
                If rsInput Is Nothing Then
                    If Not blnCancel Then Exit Function
                    If .EditText <> "" Then .EditText = .Cell(flexcpData, LngRow, DI_��ҽ֤��)
                    Set��ҽ֤�� = True: Exit Function
                End If
            End If

            .TextMatrix(LngRow, DI_֤��ID) = NVL(rsInput!��ĿID)
            .TextMatrix(LngRow, DI_֤�����) = NVL(rsInput!����)
            If Not IsNull(rsInput!����) Then
                'ȥ�����е�֤��
                If .TextMatrix(LngRow, DI_�������) Like "?*(?*)" Then
                    strTmp = Mid(.TextMatrix(LngRow, DI_�������), 1, InStrRev(.TextMatrix(LngRow, DI_�������), "(") - 1)
                Else
                    strTmp = .TextMatrix(LngRow, DI_�������)
                End If
                .TextMatrix(LngRow, DI_�������) = strTmp
                .Cell(flexcpData, LngRow, DI_�������) = .TextMatrix(LngRow, DI_�������)
                .TextMatrix(LngRow, DI_��ҽ֤��) = NVL(rsInput!����)
                .Cell(flexcpData, LngRow, DI_��ҽ֤��) = .TextMatrix(LngRow, DI_��ҽ֤��)
                If .EditText <> "" Then .EditText = .TextMatrix(LngRow, DI_��ҽ֤��)
            End If

            Set��ҽ֤�� = True
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetDiagReletedInfo(ByRef vsDiagTmp As VSFlexGrid, Optional ByVal LngRow As Long = -1)
'���ܣ�������������Ϣ�����Ը���ĳ����ϣ�������Ϸ������
    Dim bln��ҽ As Boolean
    Dim strDiagTypeName As String
    Dim strTmp As String, bln�ֻ��̶� As Boolean, blnOld�ֻ��̶� As Boolean
    Dim lngTmpRow As Long, i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim blnLockd As Boolean
    Dim n As Integer

    On Error GoTo errH
    With vsDiagTmp
        bln��ҽ = .Name = "vsDiagXY"
        If gclsPros.PatiType <> PF_���� Then
            If LngRow <> -1 Then
                lngBegin = LngRow: lngEnd = LngRow
            Else
                lngBegin = .FixedRows: lngEnd = .Rows - 1
            End If
            
            If gclsPros.FuncType = f������ҳ Then
                If bln��ҽ Then '��������
                    lngTmpRow = FindDiagRow(DT_�������): i = FindDiagRow(DT_��Ժ���XY)
                    gclsPros.CurrentForm.cmdDeliceryInfo.Visible = False
                    gclsPros.CurrentForm.cmdDeliceryInfo.Enabled = False
                    gclsPros.CurrentForm.cmdDeliceryInfo.Tag = ""
                    For j = i To lngTmpRow - 1
                        If .TextMatrix(j, DI_������Ϣ) = "1" Then
                            gclsPros.CurrentForm.cmdDeliceryInfo.Visible = True
                            gclsPros.CurrentForm.cmdDeliceryInfo.Enabled = True
                            gclsPros.CurrentForm.cmdDeliceryInfo.Tag = "1"
                            Exit For
                        End If
                    Next
                Else
                    lngTmpRow = .Rows: i = FindDiagRow(DT_��Ժ���XY)
                End If
                '�����������õ�����Ч��ֻ��������,�����ĳ�Ժ���������
                For j = i To lngTmpRow - 1
                    If .TextMatrix(j, DI_�Ƿ���) <> "1" And Val(.TextMatrix(j, DI_����ID)) <> 0 Then
                        If zlStr.NeedName(.TextMatrix(j, DI_��Ժ���)) <> "����" Then .TextMatrix(j, DI_��Ժ���) = "����"
                    End If
                Next
            End If

            If gclsPros.FuncType <> f���ѡ�� Then
                For i = lngBegin To lngEnd
                    strDiagTypeName = .TextMatrix(i, DI_�������)
                    If strDiagTypeName = "" And i >= 1 Then
                        For j = i To 1 Step -1
                            strDiagTypeName = .TextMatrix(j, DI_�������)
                            If strDiagTypeName <> "" Then Exit For
                        Next
                    End If
                    Select Case strDiagTypeName
                        Case "�ţ����������"
                            Call SetDiagMatchInfo(IIf(bln��ҽ, BCC_�������ԺXY, BCC_�������ԺZY))
                            If bln��ҽ Then Call SetDiagMatchInfo(BCC_��������Ժ)
                        Case "��Ժ���"
                            Call SetDiagMatchInfo(IIf(bln��ҽ, BCC_��Ժ���ԺXY, BCC_��Ժ���ԺZY))
                            If bln��ҽ Then Call SetDiagMatchInfo(BCC_��������Ժ)
                        Case "�������", "��Ժ���"
                            Call SetDiagMatchInfo(IIf(bln��ҽ, BCC_�������ԺXY, BCC_�������ԺZY))
                            Call SetDiagMatchInfo(IIf(bln��ҽ, BCC_��Ժ���ԺXY, BCC_��Ժ���ԺZY))
                            If bln��ҽ And strDiagTypeName = "��Ժ���" Then
                                '���ݳ�Ժ��������Ϣ������ؿؼ�����
                                strTmp = UCase(Trim(.TextMatrix(i, DI_��ϱ���)))
                                If .TextMatrix(i, DI_�������) = "��Ժ���" Then
                                    bln�ֻ��̶� = strTmp Like "C*" Or strTmp Like "D0*" Or strTmp Like "D32.*" Or strTmp Like "D33.*"
                                    blnOld�ֻ��̶� = gclsPros.CurrentForm.cboBaseInfo(BCC_�ֻ��̶�).Locked
                                    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_�ֻ��̶�), Not bln�ֻ��̶�, True)
                                    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_����������), Not bln�ֻ��̶�, True)
                                    If gclsPros.CurrentForm.Visible And bln�ֻ��̶� And blnOld�ֻ��̶� Then
                                        Call SetCboDefaultValue(BCC_�ֻ��̶�)
                                        Call SetCboDefaultValue(BCC_����������)
                                    End If
                                End If
                            End If
                        Case "�������" '��ҽ���
                            Call SetDiagMatchInfo(BCC_�����벡��)
                            Call SetDiagMatchInfo(BCC_�ٴ��벡��)
                            'Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_�����), .TextMatrix(i, DI_�������) = "", True)
                        Case "Ժ�ڸ�Ⱦ" '��ҽ���
                            Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_��ԭѧ���), .TextMatrix(i, DI_�������) = "", True)
                            Call chkInfoClick(CHK_��ԭѧ���)
                    End Select
                Next
            ElseIf gclsPros.FuncType = f���ѡ�� And gclsPros.PatiType = PF_סԺ Then
                For i = lngBegin To lngEnd
                    strDiagTypeName = .TextMatrix(i, DI_�������)
                    If strDiagTypeName = "" And i >= 1 Then
                        For j = i To 1 Step -1
                            strDiagTypeName = .TextMatrix(j, DI_�������)
                            If strDiagTypeName <> "" Then Exit For
                        Next
                    End If

                    Select Case strDiagTypeName
                        Case "�ţ����������"
                            Call SetDiagMatchInfo(IIf(bln��ҽ, BCC_�������ԺXY, BCC_�������ԺZY))
                            If bln��ҽ Then Call SetDiagMatchInfo(BCC_��������Ժ)
                        Case "��Ժ���"
                            Call SetDiagMatchInfo(IIf(bln��ҽ, BCC_��Ժ���ԺXY, BCC_��Ժ���ԺZY))
                            If bln��ҽ Then Call SetDiagMatchInfo(BCC_��������Ժ)
                        Case "�������", "��Ժ���"
                            Call SetDiagMatchInfo(IIf(bln��ҽ, BCC_�������ԺXY, BCC_�������ԺZY))
                            Call SetDiagMatchInfo(IIf(bln��ҽ, BCC_��Ժ���ԺXY, BCC_��Ժ���ԺZY))
                            If bln��ҽ And strDiagTypeName = "��Ժ���" Then
                                '���ݳ�Ժ��������Ϣ������ؿؼ�����
                                strTmp = UCase(Trim(.TextMatrix(i, DI_��ϱ���)))
                                If .TextMatrix(i, DI_�������) = "��Ժ���" Then
                                    bln�ֻ��̶� = strTmp Like "C*" Or strTmp Like "D0*" Or strTmp Like "D32.*" Or strTmp Like "D33.*"
                                    blnOld�ֻ��̶� = gclsPros.CurrentForm.cboBaseInfo(BCC_�ֻ��̶�).Locked
                                    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_�ֻ��̶�), Not bln�ֻ��̶�, True)
                                    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_����������), Not bln�ֻ��̶�, True)
                                    If gclsPros.CurrentForm.Visible And bln�ֻ��̶� And blnOld�ֻ��̶� Then
                                        Call SetCboDefaultValue(BCC_�ֻ��̶�)
                                        Call SetCboDefaultValue(BCC_����������)
                                    End If
                                End If
                            End If
                        Case "�������" '��ҽ���
                            Call SetDiagMatchInfo(BCC_�����벡��)
                            Call SetDiagMatchInfo(BCC_�ٴ��벡��)
                    End Select
                Next
            End If
        Else
            '�����д�˷���ʱ�䣬������ķ���ʱ����������д��
            blnLockd = IsDate(.TextMatrix(.FixedRows, DI_����ʱ��))
            If Not blnLockd Then
                If Not bln��ҽ Then
                    blnLockd = IsDate(gclsPros.CurrentForm.vsDiagXY.TextMatrix(gclsPros.CurrentForm.vsDiagXY.FixedRows, DI_����ʱ��))
                Else
                    blnLockd = IsDate(gclsPros.CurrentForm.vsDiagZY.TextMatrix(gclsPros.CurrentForm.vsDiagZY.FixedRows, DI_����ʱ��))
                End If
            End If
            Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_��������), blnLockd, True)
            Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_����ʱ��), blnLockd, True)
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ChangeOutInfo(Optional ByVal str��Ժ��� As String, Optional ByVal blnSet��Ժ��ʽ As Boolean, Optional ByVal blnCheckAll As Boolean)
'���ܣ�������Ժ�����ͬ������������ϵĳ�Ժ���
'������str��Ժ���=��Ժ���������Ժ���Ϊ��������ϱ༭������Ժ����䶯ʱ���Զ����������ϵĳ�Ժ���
'                      ����1������������ϣ����ô���
'                                2������������������Զ������Ժ���������״̬
'          blnSet��Ժ��ʽ=�Ƿ������ó�Ժ��ʽ
    Dim vsDiagTmp As VSFlexGrid
    Dim i As Long, lngStart As Long, lngEnd As Long
    Dim blnHave���� As Boolean
    Dim intIndex As Integer, blnLocked As Boolean, strTmp As String

     '��Ժ���Ϊ��������ҽԺ��Ⱦ������֢��������ϵĳ�Ժ���Ϊ����
    Set vsDiagTmp = gclsPros.CurrentForm.vsDiagXY
    With vsDiagTmp
        lngEnd = FindDiagRow(DT_�������): lngStart = FindDiagRow(DT_��Ժ���XY)
        If str��Ժ��� = "����" Then
            For i = lngStart To lngEnd - 1
                If .TextMatrix(i, DI_�������) <> "" Then
                    If InStr(gclsPros.CurrentForm.txtInfo(GC_��Ժ����).Text, "����") > 0 Then
                        If .TextMatrix(i, DI_�������) <> "�������" And .TextMatrix(i, DI_�������) <> "��Ժ���" Then
                            If .TextMatrix(i, DI_�������) = "" And .Cell(flexcpData, i, DI_�������) <> "�������" Then
                                .TextMatrix(i, DI_��Ժ���) = "����"
                            End If
                        End If
                    Else
                        .TextMatrix(i, DI_��Ժ���) = "����"
                    End If
                End If
            Next
        ElseIf str��Ժ��� <> "" And Not blnSet��Ժ��ʽ Then '���ǳ�Ժ��ʽ�����ĳ�Ժ����ı�
            For i = lngStart To lngEnd - 1
                If zlStr.NeedName(.TextMatrix(i, DI_��Ժ���)) = "����" Then
                    If InStr(gclsPros.CurrentForm.txtInfo(GC_��Ժ����).Text, "����") > 0 Then
                        If .TextMatrix(i, DI_�������) <> "�������" And .TextMatrix(i, DI_�������) <> "��Ժ���" Then
                            If .TextMatrix(i, DI_�������) = "" And .Cell(flexcpData, i, DI_�������) <> "�������" Then
                                .TextMatrix(i, DI_��Ժ���) = str��Ժ���
                            End If
                        End If
                    Else
                        .TextMatrix(i, DI_��Ժ���) = str��Ժ���
                    End If
                End If
            Next
        ElseIf blnSet��Ժ��ʽ Then '��Ժ��ʽ�����ĳ�Ժ����ı�,�������ĳ�Ժ������
            For i = lngStart To lngEnd - 1
                If zlStr.NeedName(.TextMatrix(i, DI_��Ժ���)) = "����" Then .TextMatrix(i, DI_��Ժ���) = ""
            Next
        Else '��Ժ���Ϊ�գ����Զ�����Ƿ����������Ժ���
            For i = lngStart To lngEnd - 1
                If zlStr.NeedName(.TextMatrix(i, DI_��Ժ���)) = "����" Then blnHave���� = True: Exit For
            Next
        End If
        '�����������õ�����Ч��ֻ��������,�����ĳ�Ժ���������
        If gclsPros.FuncType = f������ҳ And str��Ժ��� <> "����" And str��Ժ��� <> "" Then
            For i = lngStart To lngEnd - 1
                If .TextMatrix(i, DI_�Ƿ���) <> "1" And Val(.TextMatrix(i, DI_����ID)) <> 0 Then
                    .TextMatrix(i, DI_��Ժ���) = "����"
                End If
            Next
        End If
    End With
    '������ҽ��ʱ����ҽ��ϣ���ҽ���ֻ������֮һʱ������ص����ݿ��ܻᱻ���
    If gclsPros.IsTCM Then
        Set vsDiagTmp = gclsPros.CurrentForm.vsDiagZY
        With vsDiagTmp
            lngEnd = .Rows: lngStart = FindDiagRow(DT_��Ժ���ZY)
            If str��Ժ��� = "����" Then
                For i = lngStart To lngEnd - 1
                    If .TextMatrix(i, DI_�������) <> "" Then .TextMatrix(i, DI_��Ժ���) = "����"
                Next
            ElseIf str��Ժ��� <> "" And Not blnSet��Ժ��ʽ Then '���ǳ�Ժ��ʽ�����ĳ�Ժ����ı�
                For i = lngStart To lngEnd - 1
                    If zlStr.NeedName(.TextMatrix(i, DI_��Ժ���)) = "����" Then .TextMatrix(i, DI_��Ժ���) = str��Ժ���
                Next
            ElseIf blnSet��Ժ��ʽ Then '��Ժ��ʽ�����ĳ�Ժ����ı�,�������ĳ�Ժ������
                For i = lngStart To lngEnd - 1
                    If zlStr.NeedName(.TextMatrix(i, DI_��Ժ���)) = "����" Then .TextMatrix(i, DI_��Ժ���) = ""
                Next
            ElseIf Not blnHave���� Then
                For i = lngStart To lngEnd - 1
                    If zlStr.NeedName(.TextMatrix(i, DI_��Ժ���)) = "����" Then blnHave���� = True: Exit For
                Next
            End If
            '�����������õ�����Ч��ֻ��������,�����ĳ�Ժ���������
            If gclsPros.FuncType = f������ҳ And str��Ժ��� <> "����" And str��Ժ��� <> "" Then
                For i = lngStart To lngEnd - 1
                    If .TextMatrix(i, DI_�Ƿ���) <> "1" And Val(.TextMatrix(i, DI_����ID)) <> 0 Then
                        .TextMatrix(i, DI_��Ժ���) = "����"
                    End If
                Next
            End If
        End With
    End If
    '���ѡ��
    If gclsPros.FuncType = f���ѡ�� And gclsPros.PatiType = PF_סԺ Then
        If blnHave���� Then str��Ժ��� = "����"
        Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_����ʱ��), str��Ժ��� <> "����", True)
        Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_����ԭ��), str��Ժ��� <> "����", True)
        Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_����ԭ��), str��Ժ��� <> "����", True)
        Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��), str��Ժ��� <> "����", True)
        If gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).ListIndex = -1 Then
            gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).ListIndex = 0
        End If
        Exit Sub
    End If
    '���ó�Ժ��ʽʱΪ������������Ժ��ʽ���������ʱΪ����������
    If Not blnSet��Ժ��ʽ Then
        If str��Ժ��� <> "" Then blnHave���� = str��Ժ��� = "����"
        gblnSet = True
        Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_��Ժ��ʽ), blnHave����)
        If blnHave���� Then
            '��������������Ժ�������Ϊ����
            intIndex = Cbo.FindIndex(gclsPros.CurrentForm.cboBaseInfo(BCC_��Ժ��ʽ), "����")
            If intIndex = -1 Then
                gclsPros.CurrentForm.cboBaseInfo(BCC_��Ժ��ʽ).AddItem "����"
                intIndex = gclsPros.CurrentForm.cboBaseInfo(BCC_��Ժ��ʽ).NewIndex
            End If
            gclsPros.CurrentForm.cboBaseInfo(BCC_��Ժ��ʽ).ListIndex = intIndex
            blnLocked = True
            str��Ժ��� = "����"
        ElseIf str��Ժ��� <> "" And zlStr.NeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_��Ժ��ʽ).Text) = "����" Then
            gclsPros.CurrentForm.cboBaseInfo(BCC_��Ժ��ʽ).ListIndex = -1
            blnLocked = True
        Else
            str��Ժ��� = zlStr.NeedName(gclsPros.CurrentForm.cboBaseInfo(BCC_��Ժ��ʽ).Text)
            blnLocked = Not (str��Ժ��� Like "*תԺ*" Or str��Ժ��� Like "*ת����*")
        End If
        Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_ת��ҽ�ƻ���), blnLocked, True)
        Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_ת��ҽ�ƻ���), blnLocked, True)
        gblnSet = False
    End If
    '��Ժ����䶯���������Ŀؼ�״̬�ı�
    Call ChangeOutInfoSub(str��Ժ���)
End Sub

Public Sub ChangeOutInfoSub(Optional ByVal str��Ժ��� As String)
    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_����ʱ��), str��Ժ��� <> "����", True)
    Call SetCtrlLocked(gclsPros.CurrentForm.txtInfo(GC_����ԭ��), str��Ժ��� <> "����", True)
    Call SetCtrlLocked(gclsPros.CurrentForm.cmdInfo(GC_����ԭ��), str��Ժ��� <> "����", True)
    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_�����ڼ�), str��Ժ��� <> "����", True)
    Call SetCtrlLocked(gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��), str��Ժ��� <> "����", True)
    Call SetCtrlLocked(gclsPros.CurrentForm.chkInfo(CHK_����), str��Ժ��� = "����", True)
    If str��Ժ��� = "����" Then
        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).Clear
        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).AddItem "��"
        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).AddItem "��"
    Else
        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).Clear
        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).AddItem "-"
    End If
    
    If gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).ListIndex = -1 Then
        gclsPros.CurrentForm.cboBaseInfo(BCC_��������ʬ��).ListIndex = 0
    End If
    Call chkInfoClick(CHK_����)
End Sub

Public Sub EnterNextCellDiag(ByRef vsDiagTmp As VSFlexGrid)
    Dim i As Long, j As Long

    With vsDiagTmp
        '����һ��Ԫ��ʼѭ������
        If .Row < .FixedRows Then .Row = .FixedRows
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, DI_��ϱ���) To DI_Del
                If Not .ColHidden(j) Then
                    If DiagCellEditable(vsDiagTmp, i, j) And .ColWidth(j) <> 0 Then Exit For
                End If
            Next
            If j <= DI_Del Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        ElseIf i = .Rows And j > DI_Del And .TextMatrix(.Rows - 1, DI_�������) <> "" Then
            .Rows = .Rows + 1
            Call ChangeVSFHeight(vsDiagTmp, True)
            .TextMatrix(.Rows - 1, DI_��Ϸ���) = .TextMatrix(.Rows - 2, DI_��Ϸ���)
            If gclsPros.PatiType = PF_���� Then .TextMatrix(.Rows - 1, DI_�������) = .TextMatrix(.Rows - 2, DI_�������)
            .ShowCell i, IIf(gclsPros.FuncType = f������ҳ, DI_��ϱ���, DI_�������)
        Else
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        End If
    End With
End Sub

Private Sub EnterNextCellSpirit()
    '------------------------------------------------------------------------------------------------------
    '����:�ƶ���
    '���:
    '����:
    '����:
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim vsSpiritTmp As VSFlexGrid

    Set vsSpiritTmp = gclsPros.CurrentForm.vsSpirit
    With vsSpiritTmp
        '����һ��Ԫ��ʼѭ������
        If .Row < .FixedRows Then .Row = .FixedRows
         If .Col = .Cols - 1 Then
            .Col = SI_ҩ������
            If .Row = .Rows - 1 Then
                If Trim(.TextMatrix(.Row, SI_ҩ������)) <> "" Then
                    .Rows = .Rows + 1
                    .Row = .Row + 1
                    Call ChangeVSFHeight(vsSpiritTmp, True)
                End If
            Else
                .Row = .Row + 1
            End If
         Else
            .Col = .Col + 1
         End If
         If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
         End If
         If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
         End If
    End With
End Sub

Private Sub EnterNextCellOPS(ByRef vsOPS As VSFlexGrid)
    Dim i As Long, j As Long

    With vsOPS
        '����һ��Ԫ��ʼѭ������
        If .Row < .FixedRows Then .Row = .FixedRows
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, PI_��������) To PI_�п�����
                If OPSCellEditable(i, j) Then Exit For
            Next
            If j <= PI_�п����� Then Exit For
        Next
        If i <= .Rows - 1 Then
            Call .Select(i, j)
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        End If
    End With
End Sub

Private Sub EnterNextCellFees(ByRef vsFree As VSFlexGrid)
    Dim i As Long, j As Long

    With vsFree
        '����һ��Ԫ��ʼѭ������
        If .Row < .FixedRows Then .Row = .FixedRows
        For i = .Row To .Rows - 1
            For j = 0 To 5 Step 2
                If Not (j <= .Col \ 2 * 2 And i = .Row) Then
                    If Not FreeHaveLowLevel(i, j) Then Exit For
                    If .TextMatrix(i, j) <> "" Then Exit For
                End If
            Next
        Next
        If i <= .Rows - 1 Then
            Call .Select(i, j)
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab): mblnReturn = True
        End If
    End With
End Sub

Public Sub FormResize()
    On Error Resume Next
    With gclsPros.CurrentForm
        If .ScaleWidth < .picMain.Width Then
            .hsbMain.Visible = True
            .picMain.Left = .ScaleLeft + ((.ScaleWidth - .picMain.Width) * ((.hsbMain.Value) / 100))
        Else
            .hsbMain.Visible = False
            .picMain.Left = .ScaleLeft + (.ScaleWidth - .picMain.Width) / 2
        End If

        .vsbMain.Move .ScaleWidth - .vsbMain.Width, .ScaleTop, .vsbMain.Width, .ScaleHeight
        .vsbMain.LargeChange = 100
        .vsbMain.SmallChange = .vsbMain.LargeChange / 2
        
        .hsbMain.Top = .vsbMain.Top + .vsbMain.Height - 255
        .hsbMain.Left = .ScaleLeft
        .hsbMain.Width = .ScaleLeft + .ScaleWidth - 255
        .hsbMain.LargeChange = 100 / ((.picMain.Width) / .ScaleWidth)
        .hsbMain.SmallChange = 10

        .cmdTop.Move .ScaleWidth - .cmdTop.Width - .vsbMain.Width, .ScaleHeight - .cmdTop.Height - 400
        Call vsbMainChange
    End With
End Sub

Public Sub vsbMainChange()
    With gclsPros.CurrentForm
        .picMain.Top = 500 - ((.picMain.Height + 1100 - .ScaleHeight) * (.vsbMain.Value / 1000))
        If .vsbMain.Value > 0 Then
            .cmdTop.Visible = True
        Else
            .cmdTop.Visible = False
        End If
    End With
End Sub

Public Sub hsbMainChange()
    With gclsPros.CurrentForm
        .picMain.Left = .ScaleLeft + ((.ScaleWidth - .picMain.Width) * ((.hsbMain.Value) / 100))
    End With
End Sub

Public Sub cmdTopClick()
    gclsPros.CurrentForm.vsbMain.Value = 0
End Sub

Public Sub cmdTopGotFocus()
    Call ShowInfectInfo(False)
End Sub

Public Sub txtAdressInfoChange(ByRef Index As Integer)
    Call CheckValueChange(gclsPros.CurrentForm.txtAdressInfo(Index))
End Sub
Public Sub PicPageResize(ByVal Index As Integer)
    With gclsPros.CurrentForm
        If gclsPros.FuncType = f������ҳ Then
            If Index = PIC_������Ϣ Then
                .PicOut.Move 0, .vsTransfer.Top + .vsTransfer.Height + 200
            ElseIf Index = PIC_������¼ Then
                .PicOPS.Move 0, .vsOPS.Top + .vsOPS.Height + 200
            End If
        ElseIf gclsPros.FuncType = fҽ����ҳ Then
            If gclsPros.MedPageSandard = ST_����ʡ��׼ Then
                If Index = PIC_������¼ Then
                    .PicOPS.Move 0, .vsOPS.Top + .vsOPS.Height + 200
                End If
            End If
        End If
        
        If gclsPros.MedPageSandard = ST_�Ĵ�ʡ��׼ Then
            If Index = PIC_��֢�໤ Then
                .lblICUInstruments.Top = IIf(gclsPros.FuncType = fҽ����ҳ, 150, 250) + .vsFlxAddICU.Top + .vsFlxAddICU.Height
                .vsICUInstruments.Top = IIf(gclsPros.FuncType = fҽ����ҳ, 350, 500) + .vsFlxAddICU.Top + .vsFlxAddICU.Height
            End If
        End If
    End With
End Sub

Public Sub SubCMainWndProc(msg As Long, wParam As Long, lParam As Long, Result As Long)
    '�Զ������Ϣ������
    Dim wzDelta As Integer
    Select Case msg
        Case WM_MOUSEWHEEL   '����
            wzDelta = HIWORD(wParam)
            With gclsPros.CurrentForm
                If wzDelta > 0 Then        '���Ϲ���
                    Call ChangePage(False, , , False)
                Else                        '���¹���
                    Call ChangePage(True, , , False)
                End If
            End With
    End Select
End Sub

Public Function SetErrObjectColor(ByVal strErrID As String, Optional ByVal blnOld As Boolean, Optional ByVal colorBack As Long) As Object
'���ܣ���λ���ؼ����ڵ�ҳ�棬���ÿؼ�����ɫ
    Dim clsErrTmp As clsErrInfo
    Dim i As Long
    Dim objErrArr() As ERROBJ
    Dim objTmp As Object
    Dim vsfTmp As VSFlexGrid
    
    If InStr(strErrID, "Error-") > 0 Then
        Set clsErrTmp = gColErr.Item(strErrID)
    ElseIf InStr(strErrID, "Warn-") > 0 Then
        Set clsErrTmp = gColWarn.Item(strErrID)
    Else
        Exit Function
    End If
    
    If clsErrTmp Is Nothing Then
        Exit Function
    End If
    ReDim objErrArr(UBound(clsErrTmp.GetObjErr()) - LBound(clsErrTmp.GetObjErr()) + 1)
    For i = LBound(clsErrTmp.GetObjErr()) To UBound(clsErrTmp.GetObjErr())
        objErrArr(i) = clsErrTmp.GetObjErr(i)
        
        Select Case objErrArr(i).StrObjName
            Case "txtInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.txtInfo(objErrArr(i).LngObjIndex)
                End If
            Case "txtSpecificInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.txtSpecificInfo(objErrArr(i).LngObjIndex)
                End If
            Case "txtDateInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.txtDateInfo(objErrArr(i).LngObjIndex)
                End If
            Case "txtAdressInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.txtAdressInfo(objErrArr(i).LngObjIndex)
                End If
            Case "cboBaseInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.cboBaseInfo(objErrArr(i).LngObjIndex)
                End If
            Case "cboSpecificInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.cboSpecificInfo(objErrArr(i).LngObjIndex)
                End If
            Case "cboManInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.cboManInfo(objErrArr(i).LngObjIndex)
                End If
            Case "chkInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.chkInfo(objErrArr(i).LngObjIndex)
                End If
            Case "mskDateInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.mskDateInfo(objErrArr(i).LngObjIndex)
                End If
            Case "padrInfo"
                If objErrArr(i).LngObjIndex >= 0 Then
                    Set objTmp = gclsPros.CurrentForm.padrInfo(objErrArr(i).LngObjIndex)
                End If
            Case "lstInfection"
                Set objTmp = gclsPros.CurrentForm.lstInfection
            Case "lstAdvEvent"
                Set objTmp = gclsPros.CurrentForm.lstAdvEvent
            Case "vsDiagXY"
                Set objTmp = gclsPros.CurrentForm.vsDiagXY
            Case "vsDiagZY"
                Set objTmp = gclsPros.CurrentForm.vsDiagZY
            Case "vsAller"
                Set objTmp = gclsPros.CurrentForm.vsAller
            Case "vsOPS"
                Set objTmp = gclsPros.CurrentForm.vsOPS
            Case "vsChemoth"
                Set objTmp = gclsPros.CurrentForm.vsChemoth
            Case "vsRadioth"
                Set objTmp = gclsPros.CurrentForm.vsRadioth
            Case "vsSpirit"
                Set objTmp = gclsPros.CurrentForm.vsSpirit
            Case "vsKSS"
                Set objTmp = gclsPros.CurrentForm.vsKSS
            Case "vsFlxAddICU"
                Set objTmp = gclsPros.CurrentForm.vsFlxAddICU
            Case "vsICUInstruments"
                Set objTmp = gclsPros.CurrentForm.vsICUInstruments
            Case "vsfMain"
                Set objTmp = gclsPros.CurrentForm.vsfMain
            Case "vsTSJC"
                Set objTmp = gclsPros.CurrentForm.vsTSJC
            Case "vsTransfer"
                Set objTmp = gclsPros.CurrentForm.vsTransfer
            Case "vsFees"
                Set objTmp = gclsPros.CurrentForm.vsFees
            Case "vsInfect"
                Set objTmp = gclsPros.CurrentForm.vsInfect
            Case "vsSample"
                Set objTmp = gclsPros.CurrentForm.vsSample
            Case Else
                '������Ҳ�������
                If gBlnNew And (Not gfrmMecCol Is Nothing) Then
                    Err.Clear: On Error Resume Next
                    If objErrArr(i).LngObjIndex = -1 Then
                        Set objTmp = colErrTmp(objErrArr(i).StrObjName & objErrArr(i).PicIndex)
                    Else
                        Set objTmp = colErrTmp(objErrArr(i).StrObjName & objErrArr(i).PicIndex & objErrArr(i).LngObjIndex)
                    End If
                    Err.Clear: On Error GoTo 0
                End If
        End Select
        If Not objTmp Is Nothing Then
            If blnOld Then
                If TypeName(objTmp) = "VSFlexGrid" Then
                    Set vsfTmp = objTmp
                    If vsfTmp.Rows > objErrArr(i).LngRow And vsfTmp.Cols > objErrArr(i).LngCol Then
                        vsfTmp.Cell(flexcpBackColor, objErrArr(i).LngRow, objErrArr(i).LngCol, objErrArr(i).LngRow, objErrArr(i).LngCol) = objErrArr(i).OldColor
                    End If
                Else
                    objTmp.BackColor = objErrArr(i).OldColor
                End If
            Else
                If TypeName(objTmp) = "VSFlexGrid" Then
                    Set vsfTmp = objTmp
                    vsfTmp.Cell(flexcpBackColor, objErrArr(i).LngRow, objErrArr(i).LngCol, objErrArr(i).LngRow, objErrArr(i).LngCol) = colorBack
                    vsfTmp.Row = objErrArr(i).LngRow
                    vsfTmp.Col = objErrArr(i).LngCol
                    gclsPros.CurrentForm.picMain.SetFocus
                    Call LocateObjectPage(vsfTmp)
                Else
                    objTmp.BackColor = colorBack
                    gclsPros.CurrentForm.picMain.SetFocus
                    Call LocateObjectPage(objTmp)
                End If
            End If
        End If
    Next

End Function

Public Sub LocateObjectPage(ByRef objTmp As Object)
'���ܣ����ݿؼ���λ���ÿؼ����ڵ���һҳ
'����: objTmp - Ҫ��λ������ҳ�Ŀؼ�
    Dim intIndex As Integer
    Dim picTmp As PictureBox
    Dim lngObjTop As Long
    Dim strName As String
    Dim i As Integer
    
On Error GoTo errH
    If gclsPros.FuncType = f������ҳ Or gclsPros.FuncType = fҽ����ҳ Then
        lngObjTop = gclsPros.CurrentForm.picMain.Top
        If objTmp.Container.Name = "PicPage" Then
            Set picTmp = objTmp.Container
            lngObjTop = lngObjTop + picTmp.Top + objTmp.Top
        ElseIf objTmp.Container.Container.Name = "PicPage" Then
            Set picTmp = objTmp.Container.Container
            lngObjTop = lngObjTop + picTmp.Top + objTmp.Container.Top + objTmp.Top
        ElseIf objTmp.Container.Container.Container.Name = "PicPage" Then
            Set picTmp = objTmp.Container.Container.Container
            lngObjTop = lngObjTop + picTmp.Top + objTmp.Container.Top + objTmp.Container.Container.Top + objTmp.Top
        End If
        
        If Not picTmp Is Nothing Then
    '        �ڽ����Ͽ��ü��ؼ��Ļ��Ͳ���ҳ
            If Not (lngObjTop > 0 And lngObjTop + objTmp.Height < gclsPros.CurrentForm.Height) Then
                intIndex = picTmp.Index
                Call ChangePage(, intIndex, objTmp)
            Else
                objTmp.SetFocus
            End If
        End If
    End If
    Exit Sub
errH:
    Err.Clear
    '��λ��Ҹ�ҳ
    If gclsPros.FuncType = f������ҳ Or gclsPros.FuncType = fҽ����ҳ Then
        If gBlnNew And (Not gfrmMecCol Is Nothing) Then
            strName = ""
            For i = 1 To gfrmMecCol.Count
                strName = strName & "," & gfrmMecCol(i).Name
            Next
            If InStr(strName, objTmp.Container.Name) > 0 Then
                Set picTmp = gclsPros.CurrentForm.PicPage(Val(objTmp.Container.Tag))
                lngObjTop = lngObjTop + picTmp.Top + objTmp.Top
                
            ElseIf InStr(strName, objTmp.Container.Container.Name) > 0 Then
                Set picTmp = gclsPros.CurrentForm.PicPage(Val(objTmp.Container.Container.Tag))
                lngObjTop = lngObjTop + picTmp.Top + objTmp.Container.Top + objTmp.Top
                
            ElseIf InStr(strName, objTmp.Container.Container.Container.Name) > 0 Then
                Set picTmp = gclsPros.CurrentForm.PicPage(Val(objTmp.Container.Container.Container.Tag))
                lngObjTop = lngObjTop + picTmp.Top + objTmp.Container.Top + objTmp.Container.Container.Top + objTmp.Top
            End If
            
        End If
        
        If Not picTmp Is Nothing Then
    '        �ڽ����Ͽ��ü��ؼ��Ļ��Ͳ���ҳ
            If Not (lngObjTop > 0 And lngObjTop + objTmp.Height < gclsPros.CurrentForm.Height) Then
                intIndex = picTmp.Index
                Call ChangePage(, intIndex, objTmp)
            Else
                objTmp.SetFocus
            End If
        End If
    End If
End Sub

Public Sub VsErrClick(ByVal strErrID As String)
'���ܣ�����������Ĵ�����߾�����Ϣ��λ������ؼ��������ÿؼ��Ķ���ɫ
    Dim clsErrTmp As clsErrInfo
    Dim i As Long
    Dim objErrArr() As ERROBJ
    Dim objTmp As Object
    Dim picTmp As PictureBox
    Dim intIndex As Integer
    Static strOldIndex  As String
    
    If strOldIndex <> "" And strOldIndex <> strErrID Then
        If InStr(strOldIndex, "Error-") > 0 Then
            Set clsErrTmp = gColErr.Item(strOldIndex)
        ElseIf InStr(strOldIndex, "Warn-") > 0 Then
            Set clsErrTmp = gColWarn.Item(strOldIndex)
        End If
        If Not clsErrTmp Is Nothing Then
            Call SetErrObjectColor(strOldIndex, True)
            Set clsErrTmp = Nothing
        End If
    End If
    
    
    If InStr(strErrID, "Error-") > 0 Then
        Set clsErrTmp = gColErr.Item(strErrID)
        strOldIndex = strErrID
    ElseIf InStr(strErrID, "Warn-") > 0 Then
        Set clsErrTmp = gColWarn.Item(strErrID)
        strOldIndex = strErrID
    Else
        strOldIndex = ""
        Exit Sub
    End If
    
    If Not clsErrTmp Is Nothing Then
        Call SetErrObjectColor(strErrID, False, vbRed)
    End If

End Sub

Public Function SetAllObject() As Boolean
'���ܣ�����һЩ�ؼ���״̬����
    Dim objTmp As Object
    Dim strName As String
    
    If gclsPros.FuncType <> fҽ����ҳ And gclsPros.FuncType <> f������ҳ Then
        Exit Function
    End If
    gclsPros.CurrentForm.picMain.Top = gclsPros.CurrentForm.ScaleTop + 500
    For Each objTmp In gclsPros.CurrentForm.Controls
        If InStr(",VScrollBar,HScrollBar,Subclass,Line,Image,PatiAddress,", "," & TypeName(objTmp) & ",") < 1 Then
            objTmp.Appearance = 0
        End If
        If InStr(",Frame,PictureBox,CheckBox,TextBox,OptionButton,", "," & TypeName(objTmp) & ",") > 0 Then
            If TypeName(objTmp) <> "TextBox" Then
                objTmp.BackColor = GPAGECOLOR
            Else
                If Not objTmp.Locked Then objTmp.BackColor = GPAGECOLOR
            End If
            If TypeName(objTmp) = "PictureBox" Then
                objTmp.AutoRedraw = True
            End If
        End If
        If TypeName(objTmp) = "Label" Then
            objTmp.BorderStyle = 0
            objTmp.BackStyle = 0
            objTmp.Appearance = 0
        ElseIf TypeName(objTmp) = "TextBox" Then
            objTmp.BorderStyle = 0
        ElseIf TypeName(objTmp) = "PictureBox" Then
            objTmp.TabStop = False
        End If
        If objTmp.Name = "mskDateInfo" Then
            If objTmp.Index = DC_ȷ������ Then
                strName = zlRegInfo("��λ����")
                If InStr(strName, "ƽ��") > 0 Then
                    objTmp.Mask = "####-##-##"
                    objTmp.Tag = "####-##-##"
                End If
            End If
            If objTmp.Index = DC_����ʱ�� Then
                objTmp.Mask = "####-##-## ##:##"
                objTmp.Tag = "####-##-## ##:##"
            End If
        End If
        If objTmp.Name = "lblTitle" Then
            objTmp.Move 0, 10
        ElseIf objTmp.Name = "lineH" Then
            objTmp.BorderStyle = 3
        End If
    Next
    With gclsPros.CurrentForm
        If gclsPros.FuncType = f������ҳ Then
            .lblSpecificInfo(SLC_סԺ��).ForeColor = vbBlue
        ElseIf gclsPros.FuncType = fҽ����ҳ Then
            .lblAutoInfo.ForeColor = vbBlue
        End If
        .lblEdit(0).ForeColor = vbBlue
        .lblEdit(1).ForeColor = vbBlue
    End With
End Function

Public Function SetAllVSF(Optional blnPic As Boolean) As Boolean
'���ܣ���PictureBox �������VSF�ؼ��Ĵ�С��λ��
    Dim objTmp As Object
    Dim vsfTmp As VSFlexGrid
    Dim strVSFName As String
    
    If gclsPros.FuncType <> fҽ����ҳ And gclsPros.FuncType <> f������ҳ Then
        Exit Function
    End If
    
    For Each objTmp In gclsPros.CurrentForm.Controls
        If TypeName(objTmp) = "VSFlexGrid" Then
            Set vsfTmp = objTmp
            strVSFName = vsfTmp.Name
            
            vsfTmp.SelectionMode = flexSelectionFree
            vsfTmp.FocusRect = flexFocusSolid
            vsfTmp.HighLight = flexHighlightWithFocus
            vsfTmp.BackColorSel = &H404040

            vsfTmp.Left = -10
            If strVSFName = "vsOPS" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, 1000, 2)
                vsfTmp.Width = vsfTmp.Container.Width - 400
            ElseIf strVSFName = "vsfMain" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, 30)
                vsfTmp.Width = vsfTmp.Container.Width + 20
            ElseIf strVSFName = "vsDiagXY" Or strVSFName = "vsDiagZY" Then
                Call ChangeVSFHeight(vsfTmp, blnPic)
                vsfTmp.Width = vsfTmp.Container.Width - 400
            ElseIf strVSFName = "vsTSJC" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, 0)
            ElseIf strVSFName = "vsInfect" Then
                 Call ChangeVSFHeight(vsfTmp, blnPic)
                 vsfTmp.Width = vsfTmp.Container.Width / 2 - 100
            ElseIf strVSFName = "vsSample" Then
                 Call ChangeVSFHeight(vsfTmp, blnPic)
                 vsfTmp.Left = vsfTmp.Container.Width / 2 + 50
                 vsfTmp.Width = vsfTmp.Container.Width / 2 - 40
            ElseIf strVSFName = "vsTransfer" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, 20, 2)
                vsfTmp.Left = 1250
            ElseIf strVSFName = "vsAller" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, 300, 3)
                vsfTmp.Width = vsfTmp.Container.Width + 20
            ElseIf strVSFName = "vsChemoth" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, , 2)
                vsfTmp.Width = vsfTmp.Container.Width + 20
            ElseIf strVSFName = "vsRadioth" Then
                Call ChangeVSFHeight(vsfTmp, blnPic, , 2)
                vsfTmp.Width = vsfTmp.Container.Width + 20
            Else
                Call ChangeVSFHeight(vsfTmp, blnPic)
                vsfTmp.Width = vsfTmp.Container.Width + 20
            End If
        End If
    Next
End Function

Public Function ChangeVSFHeight(ByRef vsfTmp As VSFlexGrid, Optional ByVal blnPic As Boolean, Optional ByVal lngAddheight As Long = -1, Optional ByVal lngMinRows As Long = 3) As Boolean
'���ܣ���PictureBox �������VSF�Ĵ�С
    Dim i As Long
    Dim lngOldVSFHeight As Long
    Dim picContainer As PictureBox
    Dim lngRows As Long
    Dim lngVSFHeight As Long
    Dim lngRowHeight As Long
    Dim lngMaxHeight As Long
    
    If gclsPros.FuncType <> fҽ����ҳ And gclsPros.FuncType <> f������ҳ Then
        Exit Function
    End If
    Call CheckValueChange(vsfTmp)
    lngOldVSFHeight = vsfTmp.Height
    Set picContainer = vsfTmp.Container
    lngRowHeight = IIf(vsfTmp.RowHeightMax < vsfTmp.RowHeightMin, vsfTmp.RowHeightMin, vsfTmp.RowHeightMax)

    lngRows = vsfTmp.Rows
    If lngRows < lngMinRows Then lngRows = lngMinRows: vsfTmp.Rows = lngMinRows
    For i = 0 To vsfTmp.Rows - 1
        lngVSFHeight = lngVSFHeight + vsfTmp.RowHeight(i)
    Next
    lngVSFHeight = IIf(lngVSFHeight < lngRows * lngRowHeight, lngRows * lngRowHeight, lngVSFHeight)
    If lngAddheight = -1 Then lngAddheight = lngRowHeight * 1.5
    vsfTmp.Height = lngVSFHeight + lngAddheight
    
    If vsfTmp.Name = "vsInfect" Or vsfTmp.Name = "vsSample" Then
        lngMaxHeight = IIf(gclsPros.CurrentForm.vsInfect.Height > gclsPros.CurrentForm.vsSample.Height, gclsPros.CurrentForm.vsInfect.Height, gclsPros.CurrentForm.vsSample.Height)
        If vsfTmp.Height - lngOldVSFHeight <> 0 Then
            picContainer.Height = vsfTmp.Top + lngMaxHeight + 300
            If blnPic Then
                Call SetPicPosition(True)
            End If
        End If
    ElseIf vsfTmp.Height - lngOldVSFHeight <> 0 Then
        picContainer.Height = picContainer.Height + (vsfTmp.Height - lngOldVSFHeight)
        If blnPic Then
            Call SetPicPosition(True)
        End If
    End If
End Function


Public Function SetPicPosition(Optional ByVal blnV As Boolean, Optional ByVal blnH As Boolean) As Boolean
'���ܣ���PictureBox �������ÿ��PicPage�ؼ���λ��
    Dim lngLeft As Long
    Dim i As Long, j As Long, lngHeight As Long
    
    With gclsPros.CurrentForm
        lngLeft = .picMain.ScaleLeft + ((.picMain.ScaleWidth - .PicPage(0).Width) / 2)

        For i = .PicPage.LBound To .PicPage.UBound
            If .PicPage(i).Tag = "true" Then
                .PicPage(i).Visible = True
                If i = .PicPage.LBound Then
                    .PicPage(i).Move lngLeft, .picMain.ScaleTop
                Else
                    .PicPage(i).Move lngLeft, .PicPage(j).Top + .PicPage(j).Height
                End If
                j = i
                lngHeight = lngHeight + .PicPage(i).Height
            Else
                .PicPage(i).Visible = False
            End If
        Next
        .picMain.Height = .PicPage(0).ScaleTop + lngHeight + 500
    
        Call DrawLine(blnV, blnH)
    End With
End Function

Private Sub DrawLine(Optional ByVal blnV As Boolean, Optional ByVal blnH As Boolean)
'���ܣ�������ҳ�滭�ϱ߿���ÿһ��PicPage�����滭�Ϸָ��ߣ�TextBox�ؼ������滭ֱ��,��ComboBox�������Frame
'����: blnV -�����ŵı߿���, blnH - ÿһ��PicPage�����滭�Ϸָ��ߣ�TextBox�ؼ������滭ֱ��,��ComboBox�������Frame
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, j As Long
    Dim objText As Object
    Dim objPic As PictureBox
    Dim lngMin As Long, lngMax As Long
    Dim blnFra As Boolean
    Dim cboTmp As ComboBox
    On Error Resume Next
    
    With gclsPros.CurrentForm
         For i = .PicPage.LBound To .PicPage.UBound
            If blnH Then
                .PicPage(i).Cls
            End If
            If .PicPage(i).Tag = "true" Then
                lngMax = i
            End If
        Next
        lngMin = .PicPage.LBound + 1
        '����ҳ��ı߿���
        If blnV Then
            .picMain.Cls
            '����ҳ����������
            .picMain.DrawWidth = 1
            x1 = .PicPage(lngMin).Left - 15
            y1 = .PicPage(lngMin).Top
            x2 = x1
            y2 = .PicPage(lngMax).Top + .PicPage(lngMax).Height
            .picMain.Line (x1, y1)-(x2, y2)
            '����ҳ����ұ�����
            x1 = .PicPage(lngMin).Left + .PicPage(lngMin).Width + 5
            y1 = .PicPage(lngMin).Top
            x2 = x1
            y2 = .PicPage(lngMax).Top + .PicPage(lngMax).Height
            .picMain.Line (x1, y1)-(x2, y2)
        End If
        
        'ÿ��PictureBox���ϱ߻�һ������
        If blnH Then
            For i = lngMin To lngMax
                If .PicPage(i).Tag = "true" Then
                    x1 = .PicPage(i).ScaleLeft
                    y1 = .PicPage(i).ScaleTop
                    x2 = .PicPage(i).ScaleLeft + .PicPage(i).ScaleWidth
                    y2 = y1
                    .PicPage(i).DrawWidth = 1
                    .PicPage(i).Line (x1, y1)-(x2, y2)
                    If i = lngMax Then
                        .PicPage(i).Line (x1, y1 + .PicPage(i).ScaleHeight - 10)-(x2, y2 + .PicPage(i).ScaleHeight - 10)
                    End If
                End If
            Next
'            ����������������ComboBox�������Frame ִֻ��һ��
            blnFra = (.fraCbo.UBound = 0)
            
            For Each objText In .Controls
                '��ÿ��TextBox ���滭һ����
                If TypeName(objText) = "TextBox" Then
                    If objText.Name <> "txtAdressInfo" Then
                        DrawLineCTL objText
                    ElseIf objText.Name = "txtAdressInfo" Then
                        If gclsPros.IsStructAdress Then
                            If objText.Index = ADRC_��λ��ַ Then
                                If gclsPros.MedPageSandard <> ST_�Ĵ�ʡ��׼ Then
                                    DrawLineCTL objText
                                End If
                            ElseIf objText.Index = ADRC_�������� Then
                                DrawLineCTL objText
                            End If
                        Else
                            DrawLineCTL objText
                        End If
                    End If
                ElseIf TypeName(objText) = "ComboBox" Then  '��ÿһ��ComboBox��������һ��Frame��ʹ֮��������ƽ���
                    If blnFra And TypeName(objText.Container) = "PictureBox" Then
                        Set cboTmp = objText
                        j = j + 1
                        Load .fraCbo(j)
                        Set .fraCbo(j).Container = cboTmp.Container
                        .fraCbo(j).Left = cboTmp.Left
                        .fraCbo(j).Top = cboTmp.Top + 25
                        .fraCbo(j).Width = cboTmp.Width
                        .fraCbo(j).Height = IIf(gclsPros.FuncType = f������ҳ, 250, 225)
                        .fraCbo(j).BackColor = GPAGECOLOR
                        If cboTmp.Tag = "����" Then
                            .fraCbo(j).Visible = False
                        Else
                            .fraCbo(j).Visible = True
                        End If
                        Set cboTmp.Container = .fraCbo(j)
                        cboTmp.Width = cboTmp.Width + 50
                        cboTmp.Left = -25
                        cboTmp.Top = -25
                    End If
                End If
            Next
        
            For j = .fraCbo.LBound + 1 To .fraCbo.UBound
                DrawLineCTL .fraCbo(j)
            Next
        End If
    End With
End Sub

Private Sub DrawLineCTL(ByRef objCtl As Object, Optional ByVal bytModel As Byte = 0)
'����:��ָ������һ���߻������ԭ������
'objCtl-����ؼ����󣬸��ݸÿؼ������ȡ��Ӧ����ֵ
'bytModel=0-����;1-�����
    Dim objPic As Object  '����
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    Select Case TypeName(objCtl)
    Case "Frame"
        'FraCbo���滭һ����
        x1 = objCtl.Left
        y1 = objCtl.Top + objCtl.Height + 3
        x2 = objCtl.Left + objCtl.Width - 20
        y2 = y1
    Case "TextBox"
        '��ÿ��TextBox ���滭һ����
        x1 = objCtl.Left
        y1 = objCtl.Top + objCtl.Height + 3
        x2 = objCtl.Left + objCtl.Width
        y2 = y1
    End Select
    Set objPic = objCtl.Container
    objPic.DrawWidth = 1
    If bytModel = 0 Then
        objPic.Line (x1, y1)-(x2, y2)
    Else
        objPic.Line (x1, y1)-(x2, y2), objPic.BackColor '�������
    End If
End Sub

Public Sub LoadVsErrData()
'���ܣ���������Ϣ�;�����Ϣ���ص�������
    Dim clsErr As clsErrInfo
    If gColErr.Count <= 0 And gColWarn.Count <= 0 Then Exit Sub
    frmMain.dkpMain.FindPane(Pane_���).Closed = False
    With frmMain.vsErr
        .OutlineBar = flexOutlineBarCompleteLeaf
        .OutlineCol = 0
        .Rows = .FixedRows
        If gColErr.Count > 0 Then
            .Rows = .Rows + 1
            .MergeCells = flexMergeFree
            .TextMatrix(.Rows - 1, ERR_ID) = "����" & CStr(gColErr.Count) & "����"
            .TextMatrix(.Rows - 1, ERR_����) = "����" & CStr(gColErr.Count) & "����"
            .TextMatrix(.Rows - 1, ERR_��Ϣ) = "����" & CStr(gColErr.Count) & "����"
            .MergeRow(.Rows - 1) = True
            .Cell(flexcpAlignment, .Rows - 1, ERR_ID, .Rows - 1, ERR_��Ϣ) = flexAlignLeftCenter
            .Cell(flexcpFontBold, .Rows - 1, ERR_ID, .Rows - 1, ERR_��Ϣ) = True
            .IsSubtotal(.Rows - 1) = True
            .RowOutlineLevel(.Rows - 1) = 0
            For Each clsErr In gColErr
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, ERR_ID) = "  " & Mid(clsErr.strErrID, InStr(clsErr.strErrID, "-") + 1)
                .Cell(flexcpData, .Rows - 1, ERR_ID) = clsErr.strErrID
                .TextMatrix(.Rows - 1, ERR_����) = "����"
                .TextMatrix(.Rows - 1, ERR_��Ϣ) = clsErr.StrErrInfo
                Set .Cell(flexcpPicture, .Rows - 1, 0) = frmMain.imgError.Picture
                .Cell(flexcpAlignment, .Rows - 1, 0, .Rows - 1, 0) = flexAlignLeftCenter
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 1
            Next
        End If
        
        If gColWarn.Count > 0 Then
            .Rows = .Rows + 1
            .MergeCells = flexMergeFree
            .TextMatrix(.Rows - 1, ERR_ID) = "���棨" & CStr(gColWarn.Count) & "����"
            .TextMatrix(.Rows - 1, ERR_����) = "���棨" & CStr(gColWarn.Count) & "����"
            .TextMatrix(.Rows - 1, ERR_��Ϣ) = "���棨" & CStr(gColWarn.Count) & "����"
            .MergeRow(.Rows - 1) = True
            .Cell(flexcpAlignment, .Rows - 1, ERR_ID, .Rows - 1, ERR_��Ϣ) = flexAlignLeftCenter
            .Cell(flexcpFontBold, .Rows - 1, ERR_ID, .Rows - 1, ERR_��Ϣ) = True
            .IsSubtotal(.Rows - 1) = True
            .RowOutlineLevel(.Rows - 1) = 0
            For Each clsErr In gColWarn
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, ERR_ID) = "  " & Mid(clsErr.strErrID, InStr(clsErr.strErrID, "-") + 1)
                .Cell(flexcpData, .Rows - 1, ERR_ID) = clsErr.strErrID
                .TextMatrix(.Rows - 1, ERR_����) = "����"
                .TextMatrix(.Rows - 1, ERR_��Ϣ) = clsErr.StrErrInfo
                Set .Cell(flexcpPicture, .Rows - 1, 0) = frmMain.imgWarn.Picture
                .Cell(flexcpAlignment, .Rows - 1, 0, .Rows - 1, 0) = flexAlignLeftCenter
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 1
            Next
        End If
        
        .Cell(flexcpForeColor, .FixedRows, ERR_ID, .Rows - 1, ERR_����) = vbRed
        .Cell(flexcpForeColor, .FixedRows, ERR_��Ϣ, .Rows - 1, ERR_��Ϣ) = vbBlue
    End With
End Sub

Public Sub ClearErrCol()
'���ܣ������������Ϣ�;�����Ϣ
    Dim i As Long
    
    Call VsErrClick("")
    
    If gColErr.Count > 0 Then
        For i = 1 To gColErr.Count
            gColErr.Remove 1
        Next
    End If
    If gColWarn.Count > 0 Then
        For i = 1 To gColWarn.Count
            gColWarn.Remove 1
        Next
    End If
    
    frmMain.vsErr.Rows = frmMain.vsErr.FixedRows
    frmMain.dkpMain.FindPane(Pane_���).Closed = True
End Sub

Public Sub menuPageOperate(ByVal mopType As MedRec_Operate, Optional ByVal intPage As Integer)
    Dim strMsg As String
    Select Case mopType
        Case MOP_Ԥ��
            strMsg = "Ԥ��"
        Case MOP_��ӡ
            strMsg = "��ӡ"
        Case MOP_ȷ��
            strMsg = "����"
    End Select
    If PageOperate(mopType, intPage) Then
        strMsg = strMsg & "�ɹ���"
    Else
        If gColErr.Count > 0 Then
            strMsg = strMsg & "ʧ�ܣ�����" & CStr(gColErr.Count) & "������" & CStr(gColWarn.Count) & "�����棡"
        Else
            strMsg = strMsg & " ʧ�ܣ�"
        End If
    End If
    frmMain.stbThis.Panels(2).Text = strMsg
End Sub

Private Sub SetComboBoxProperty(ByVal blnMask As Boolean)
'���ܣ�����������֮������Combobox�ؼ������ݴ���ѡ��״̬��ȡ������ѡ��״̬
'������blnMask - true ���ε�ComboBox�ؼ����������¼���false ȡ����ComboBox�ؼ����������¼�������
    Dim cboTmp As ComboBox
    Dim objTmp As Object
    
    If gclsPros.CboMask = blnMask Then
        Exit Sub
    End If
    gclsPros.CboMask = blnMask
    For Each objTmp In gclsPros.CurrentForm.Controls
        If TypeName(objTmp) = "ComboBox" Then
            Set cboTmp = objTmp
            If blnMask Then
                Call CallHook(cboTmp.hwnd)
                If cboTmp.Style = 0 Then cboTmp.SelLength = 0
            Else
                Call CallUnhook(cboTmp.hwnd)
            End If
        End If
    Next
End Sub



Public Sub SetYoubian(Index As Integer, intLevel As Integer, rsReturn As ADODB.Recordset)
'���ܣ������벡�˽ṹ����ַ��ʱ��,�����ʱ�
    If (Not rsReturn Is Nothing) And intLevel = 2 Then
        If Index = ADRC_��סַ Then
            gclsPros.CurrentForm.txtSpecificInfo(SLC_��ͥ�ʱ�).Text = rsReturn!�ʱ� & ""
        ElseIf Index = ADRC_���ڵ�ַ Then
            gclsPros.CurrentForm.txtSpecificInfo(SLC_�����ʱ�).Text = rsReturn!�ʱ� & ""
        ElseIf Index = ADRC_��λ��ַ Then
            gclsPros.CurrentForm.txtSpecificInfo(SLC_��λ�ʱ�).Text = rsReturn!�ʱ� & ""
        End If
    End If
End Sub

Public Sub DiagMouseDown(ByRef vsDiag As VSFlexGrid, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
    Dim LngRow As Long
    If intButton = 2 Then
        vsDiag.SetFocus
        LngRow = vsDiag.MouseRow
        If LngRow >= vsDiag.FixedRows And LngRow <= vsDiag.Rows - 1 Then
            If Not vsDiag.RowHidden(LngRow) Then vsDiag.Row = LngRow
        End If
    End If
End Sub

Public Sub DiagMouseUp(ByRef vsDiag As VSFlexGrid, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
    Dim objPopup As CommandBarPopup
    Dim blnDo As Boolean

    If intButton = 2 Then
        Set mobjDiag = Nothing
        If frmMain.cbsMain Is Nothing Then Exit Sub
        Set objPopup = frmMain.cbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If gobjPlugIn Is Nothing And blnDo Then Exit Sub '������û�в˵���Ŀʱ����ʾһ���հ�С����
        If Not objPopup Is Nothing Then
            Set mobjDiag = vsDiag
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Public Sub ExeDiagPlugIn(ByVal strName As String)
'���ܣ�ִ�������ҹ���
    Dim lngID As String
    Dim strXML As String
    If CreatePlugInOK(gclsPros.Module) And (Not mobjDiag Is Nothing) Then
        With mobjDiag
            lngID = Val(.RowData(.Row))
            strXML = "<ROOT><���ID>" & .TextMatrix(.Row, DI_���ID) & "</���ID><����ID>" & .TextMatrix(.Row, DI_����ID) & "</����ID></ROOT>"
            On Error Resume Next
            Call gobjPlugIn.ExecuteFunc(gclsPros.SysNo, gclsPros.Module, strName, gclsPros.����ID, gclsPros.��ҳID, lngID, .TextMatrix(.Row, DI_�������), 6, strXML)
            Call zlPlugInErrH(Err, "ExecuteFunc")
            Err.Clear: On Error GoTo 0
        End With
    End If
End Sub

Public Sub ChangeCtl()
    '����ý���Ŀؼ� ������Ļ��ʾλ��
On Error GoTo errH
    If mblnReturn = True Then
        If Not gclsPros.CurrentForm.ActiveControl Is Nothing Then
            If Not gclsPros.CurrentForm.ActiveControl.Container Is Nothing Then
                If Not gclsPros.CurrentForm.ActiveControl.Container Is Nothing Then
                    If gclsPros.CurrentForm.ActiveControl.Container.Name = "PicPage" Or gclsPros.CurrentForm.ActiveControl.Container.Name = "fraCbo" Then
                        Call LocateObjectPage(gclsPros.CurrentForm.ActiveControl)
                        mblnReturn = False
                    End If
                End If
            End If
        End If
    End If
errH:
    Err.Clear
End Sub
