Attribute VB_Name = "mdlPassDefine_MK"
Option Explicit

'PASS�ӿں���������˵���μ�PASS�ӿ��ĵ�����
'--------------------------------------------------------------------------------------------------------------------------------------
'�����ӿڶ���    version = 3.0
'--------------------------------------------------------------------------------------------------------------------------------------
'˵����ShellRunAs.dll��Ҫ��װ�ڳ����ϵͳĿ¼
'      DIFPassDll.dllΪPassϵͳ�Զ���ȡ��ע��·��
'ע�������
Public Declare Function RegisterServer Lib "ShellRunAs.dll" () As Integer
'PASS��ʼ��
Public Declare Function PassInit Lib "DIFPassDll.dll" ( _
                                 ByVal UserName As String, _
                                 ByVal DepartMentName As String, _
                                 ByVal WorkstationType As Integer) As Integer
'PASS����ģʽ����
Public Declare Function PassSetControlParam Lib "DIFPassDll.dll" ( _
                                            ByVal SaveCheckResult As Integer, _
                                            ByVal AllowAllegen As Integer, _
                                            ByVal CheckMode As Integer, _
                                            ByVal DisqMode As Integer, _
                                            ByVal UseDiposeIdea As Integer) As Integer
'AllowAllegen �Ƿ�����˹���ʷ״̬�����������������û����룻�����У��ӣӹ��������У��ӣ�ǿ�ƹ���

'�����˻�����Ϣ
Public Declare Function PassSetPatientInfo Lib "DIFPassDll.dll" ( _
                                           ByVal PatientID As String, _
                                           ByVal VisitID As String, _
                                           ByVal Name As String, _
                                           ByVal Sex As String, _
                                           ByVal Birthday As String, _
                                           ByVal Weight As String, _
                                           ByVal cHeight As String, _
                                           ByVal DepartMentName As String, _
                                           ByVal Doctor As String, _
                                           ByVal LeaveHospitalDate As String) As Integer
'������ҩƷ��Ϣ
Public Declare Function PassSetRecipeInfo Lib "DIFPassDll.dll" ( _
                                          ByVal OrderUniqueCode As String, _
                                          ByVal DrugCode As String, _
                                          ByVal DrugName As String, _
                                          ByVal SingleDose As String, _
                                          ByVal DoseUnit As String, _
                                          ByVal Frequency As String, _
                                          ByVal StartOrderDate As String, _
                                          ByVal StopOrderDate As String, _
                                          ByVal RouteName As String, _
                                          ByVal GroupTag As String, _
                                          ByVal OrderType As String, _
                                          ByVal OrderDoctor As String) As Integer

'���벡�˹���ʷ
Public Declare Function PassSetAllergenInfo Lib "DIFPassDll.dll" _
                                            (ByVal AllergenIndex As String, _
                                             ByVal AllergenCode As String, _
                                             ByVal AllergenDesc As String, _
                                             ByVal AllergenType As String, _
                                             ByVal Reaction As String) As Integer
'����:
'     AllergenIndex-����ԭ��ҽ���е�˳���ţ�Ҫ��Ψһ
'     AllergenCode-����ԭ���룬ҩƷId
'     AllergenDesc-����ԭ����
'     AllergenType-�̶�����DrugName
'     Reaction-����֢״�����˿մ�

'���벡��״̬
Public Declare Function PassSetMedCond Lib "DIFPassDll.dll" _
                                       (ByVal MedCondIndex As String, _
                                        ByVal MedCondCode As String, _
                                        ByVal MedCondDesc As String, _
                                        ByVal MedCondType As String, _
                                        ByVal StartDate As String, _
                                        ByVal EndDate As String) As Integer
'����:
'     MedCondIndex-�����ţ�Ψһ����
'     MedCondCode-��ϱ���
'     MedCondDesc-�������
'     MedCondType-�������(User)
'     StartDate-��ʼ���� ��ǰʱ�䣬 ��ȷ���죬yyyy-mm-dd
'     EndDate-�������� ��ǰʱ�䣬��ȷ���죬yyyy-mm-dd


'������Ҫ���е�ҩ�����ҩƷ
Public Declare Function PassSetWarnDrug Lib "DIFPassDll.dll" (ByVal DrugUniqueCode As String) As Integer
'��Ϣ��ѯҩƷ����
Public Declare Function PassSetQueryDrug Lib "DIFPassDll.dll" ( _
                                         ByVal DrugCode As String, _
                                         ByVal DrugName As String, _
                                         ByVal DoseUnit As String, _
    ByVal RouteName As String) As Integer
'��ȡ�Ҽ��˵��Ƿ����ֵ
Public Declare Function PassGetState Lib "DIFPassDll.dll" (ByVal QueryItemNo As String) As Integer
'PASS���ܵ���
Public Declare Function PassDoCommand Lib "DIFPassDll.dll" (ByVal CommandNo As Integer) As Integer
'��ȡҩƷ��ʾ����
Public Declare Function PassGetWarn Lib "DIFPassDll.dll" (ByVal DrugUniqueCode As String) As Integer
'����ҩƷ��������λ��
Public Declare Function PassSetFloatWinPos Lib "DIFPassDll.dll" ( _
    ByVal Left As Integer, ByVal Top As Integer, _
    ByVal Right As Integer, ByVal Bottom As Integer) As Integer
'PASS�˳�����
Public Declare Function PassQuit Lib "DIFPassDll.dll" () As Integer

'----------------------------------------------------------------------------------------------------------------------------------
'--------------         �����ӿ�����   version 4.0 ------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------
'*******PASS4.0**1-����Ƕ����뿪ʼ��DLL����������*****************************

'1��PASS��ʼ��
Public Declare Function MDC_Init Lib "PASS4Invoke.dll" (ByVal pcCheckMode As String, ByVal pcHisCode As String, ByVal pcDoctorCode As String) As Integer
'�������:
'pcCheckMode: �ַ��������ģʽ������ʹ��ϵͳ���ö����ģʽ�����ݴ���ֵ�Ĳ�ͬ�������ͬ��ά�����ߵ���DLLʱ���գ�����ʾ��������
'pcHisCode: �ַ�����ҽԺ���룬��ҽԺģʽ�����ַ�������his�ṩ��ҽԺ���룬����ģʽ��his�ṩ��ҽԺ���롣
'pcDoctorCode: �ַ�����ҽ�����룬�����¼ҽ�����룬������ҽ���ֵ�����еġ�������¼����ƽ̨��

'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��
'-1-ִ�����ʱ
'-2-����PASS������ʧ��
'-3-��ȡ��顢��ѯ�б����
'-4-��ʼ������������
'-5-������Դ�ļ�����
'����: ϵͳ�������ȵ���MDC_Init�ɹ�����ܵ����������ܺ���


'2����ȡPASSϵͳ���һ�δ�����Ϣ����
Public Declare Function MDC_GetLastError Lib "PASS4Invoke.dll" () As String
'�������:
'��
'����ֵ: �ַ��� -������Ϣ

'3������ຯ��

'3-1 ������������Ϣ�ຯ��
'3-1-1 �����˻�����¼��Ϣ
Public Declare Function MDC_SetPatient Lib "PASS4Invoke.dll" ( _
                    ByVal pcPatCode As String, _
                    ByVal pcInHospNo As String, _
                    ByVal pcVisitCode As String, _
                    ByVal pcName As String, _
                    ByVal pcSex As String, _
                    ByVal pcBirthday As String, _
                    ByVal pcHeightCM As String, _
                    ByVal pcWeighKG As String, _
                    ByVal pcDeptCode As String, _
                    ByVal pcDeptName As String, _
                    ByVal pcDoctorCode As String, _
                    ByVal pcDoctorName As String, _
                    ByVal piPatStatus As Integer, _
                    ByVal piIsLactation As Long, _
                    ByVal piIsPregnancy As Long, _
                    ByVal pcPregStartDate As String, _
                    ByVal piHepDamageDegree As Long, _
                    ByVal piRenDamageDegree As Long) As Integer
'�������:
'pcPatCode���ַ������ͣ���ʾ����ID�������pcVisitCodeΨһȷ��һ�����ˣ��˲�������Ϊ�ա�
'pcInHospNo:�������ͣ���ʾ���˴����Ż�סԺ�ţ��˲�������Ϊ�ա�
'pcVisitCode���ַ������ͣ���ʾ���˾��������סԺ�����������pcPatCodeΨһȷ��һ�����ˣ����HISϵͳû�д���Ϣ����ɴ���"1"��
'pcName���ַ������ͣ���ʾ����������
'pcSex���ַ������ͣ���ʾ�����Ա𣬸�ʽΪ"��"��"Ů"��"����"�����û�и�ֵ����Ӱ�첡��"����"��"����"��"�Ա�"ģ�����顣
'pcBirthday���ַ������ͣ���ʾ���˳������ڣ���ʽΪ"yyyy-mm-dd"�����û�и�ֵ����Ӱ�첡��"����"��"��ͯ����"��"�����˾���"��"����"��"�Ρ�������"ģ�����顣���磺"1976-08-12"
'pcHeightCM���ַ������ͣ���ʾ����������Ϊ��λ�����ֵ������ĳ�������Ϊ175���ף���Ӧ����"175"�����HISϵͳû�й����������Ϣ����Ӧ������ַ�����
'pcWeighKG���ַ������ͣ���ʾ�����Թ���Ϊ��λ������ֵ������ĳ��������Ϊ23.5�����Ӧ����"23.5"�����ڴ������ʱ���ܴ��뵥λ���������HISϵͳ������߲����Թ���Ϊ��λ����Ҫ����뻻��ɹ�����ٴ�����ֵ�����HISϵͳû�й�����������Ϣ����Ӧ������ַ�������������Ρ������������.
'pcDeptCode: �ַ������ͣ���ʾ���ұ��롣
'pcDeptName���ַ������ͣ���ʾ�������ơ�
'pcDoctorCode���ַ������ͣ���ʾ����/�Һ�ҽ�����롣
'pcDoctorName���ַ������ͣ���ʾ����/�Һ�ҽ�����ơ�
'piPatStatus�����ͣ���ʾ����״̬��1��ʾסԺ���ˣ�Ĭ�ϣ���2��ʾ���ﲡ�ˣ�3��ʾ���ﲡ�ˡ�
'piIsLactation�����ͣ���ʾ���˲���״̬��������ͨ��PassSetMedCond������������"������"��ʽ����飬ȡֵ�� -1-�޷���ȡ����״̬��Ĭ�ϣ�;0-����;1-��
'piIsPregnancy�����ͣ���ʾ��������״̬��������ͨ��PassSetMedCond������������"������"��ʽ����飬ȡֵ�� -1-�޷���ȡ����״̬��Ĭ�ϣ�;0-����;1-��
'pcPregStartDate���ַ������ͣ���ʾ���￪ʼ���ڣ���ʽΪyyyy-mm-dd��
'piHepDamageDegree�����ͣ���ʾ���˸��𺦳̶ȣ�������ͨ��PassSetMedCond�������������������ϵ���飬ȡֵ�� -1-��ȷ����Ĭ�ϣ���0-�޸��𺦣�1-�ι��ܲ�ȫ��2-��ȸ��𺦣�3-�жȸ��𺦣�4-�ضȸ���
'piRenDamageDegree���ͣ���ʾ�������𺦳̶ȣ�������ͨ��PassSetMedCond��������������������ϵ���飬ȡֵ�� -1-��ȷ����Ĭ�ϣ���0-�����𺦣�1-�����ܲ�ȫ��2-������𺦣�3-�ж����𺦣�4-�ض�����
'
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��
'���ã����˵Ļ�����Ϣ�����仯֮�󣬵��øýӿڡ�


'3-1-2 ������ҩƷ��¼��Ϣ
Public Declare Function MDC_AddScreenDrug Lib "PASS4Invoke.dll" ( _
                    ByVal pcIndex As String, ByVal piOrderNo As Integer, _
                    ByVal pcDrugUniqueCode As String, ByVal pcDrugName As String, _
                    ByVal pcDosePerTime As String, ByVal pcDoseUnit As String, _
                    ByVal pcFrequency As String, ByVal pcRouteCode As String, _
                    ByVal pcRouteName As String, ByVal pcStartTime As String, _
                    ByVal pcEndTime As String, ByVal pcExecuteTime As String, _
                    ByVal pcGroupTag As String, ByVal pcIsTempDrug As String, _
                    ByVal pcOrderType As String, ByVal pcDeptCode As String, _
                    ByVal pcDeptName As String, ByVal pcDoctorCode As String, _
                    ByVal pcDoctorName As String, _
                    ByVal pcRecipNo As String, ByVal pcNum As String, _
                    ByVal pcNumUnit As String, ByVal pcPurpose As String, _
                    ByVal pcOprCode As String, ByVal pcMediTime As String, ByVal pcRemark As String) As Integer

'�������:
'pcIndex���ַ������ͣ���ʾҽ��Ψһ�룬PASSϵͳ�����ݴ˲�����ʶ������ִ���ĸ���ҽ����¼������HISϵͳֻ��ͨ���˲�������ȡPASS���Ľ��ֵ����ͬһѭ������ʱ��Ҫ�����¼��pcIndexֵ����Ψһ�����磬�ɴ����¼���к�ֵ��
'piOrderNo�����ͣ���ʾҽ�����,��ʾͬһ����鴫��ҩƷ��˳��ţ�����ȷ��������������ĸ�ҩ���������-1������ϵͳ���ݵ��ýӿ�˳���Զ�����
'pcDrugUniqueCode���ַ������ͣ���ʾҩƷΨһ�룬Ҫ����PASSϵͳ���ʱ���õ�ҩƷΨһ����ȫһ�£�����PASSϵͳ�޷�ʶ��ҩƷ��Ϣ���˲�������Ϊ�ա�
'pcDrugName���ַ������ͣ���ʾҩƷ���ơ�
'pcDosePerTime���ַ������ͣ���ʾÿ��ʹ�ü��������ֲ��֣�����˲�����Ҫ����PASS�Բ���ÿ�η��ü�������顣ע�⣺�˴�Ҫ����ת��Ϊ��ҩƷ��Լ�����λ��ȫһ�µ�λ�����ֵ������ҩƷ��Լ�����λΪ"mg"�������˵�ÿ�η��ü���Ϊ"0.5g"����ʱ�Ͳ��ܴ���"0.5"����Ӧ����Ϊ"500mg"�󣬴���"500"���˲������Ϊ�գ�������������
'pcDoseUnit���ַ������ͣ���ʾÿ�η��ü�����λ��Ҫ����ҩƷ��Լ�����λ��ȫһ�£����������ɼ�����鲻��ȷ��
'pcFrequency���ַ������ͣ���ʾҩƷ����Ƶ����Ϣ��ע�⣬Ҫ����PASSϵͳ���ʱ���õ�Ƶ�α�����ȫһ�¡�
'pcRouteCode���ַ������ͣ���ʾ��ҩ;�����롣ע�⣬Ҫ����PASSϵͳ���ʱ���õĸ�ҩ;��������ȫһ�£�����PASSϵͳ������ҩ;����ϵ���У��˲���������󣬽�ֱ�ӵ���������������գ�����PASSϵͳ�޷�������ҩ;����ص������Ŀ��
'pcRouteName���ַ������ͣ���ʾ��ҩ;�����ơ�
'pcStartTime���ַ������ͣ���ʾ����ҽ�����ڡ���ʽΪ"yyyy-mm-dd hh:mm:ss "�����翪������Ϊ1999��3��12�գ���Ӧ����"1999-03-12 00:00:00"��
'pcEndTime���ַ������ͣ������������ʾͣ�����ڣ���ʽΪ"yyyy-mm-dd hh:mm:ss "������ͣ������Ϊ1999��3��12�գ���Ӧ����"1999-03-12 00:00:00"������ͣ�����ڵ��ڿ������ڣ�δͣ����ҽ��ͣ�����ڴ����ַ�����
'pcExecuteTime���ַ������ͣ���ʾִ��ҽ��ʱ�䡣��ʽΪ"yyyy-mm-dd hh:mm:ss"��
'pcGroupTag���ַ������ͣ���ʾ����ҽ����ǡ���Ҫ����PASSϵͳ����ע��������������ʶ��ע����Ƿ�����һ��ʹ�ã���ѭ�������ҽ���У�����˲���ֵ��ͬ�����ʾ��������һ���ã���������²��п��ܴ��������������⡣
'pcIsTempDrug���ַ������ͣ���ʾҽ���ǳ���ҽ��������ʱҽ����'0'-��ʾ����ҽ���� '1'-��ʾ��ʱҽ����
'pcOrderType���ַ������ͣ���ʾҽ�����ȡֵ'0'-���ã�Ĭ�ϣ���'1'-�����ϣ�'2'-��ͣ����'3'-��Ժ��ҩ������ϵͳ���ò�����飩��������ҽԺ��������飬���һ�ɾ�����ҽ��pcindex�йص��������������ͣ����������飬����Ӱ��ͣ��ǰ���������
'pcDeptCode���ַ������ͣ���ʾ�������ұ��롣
'pcDeptName���ַ������ͣ���ʾ�����������ơ�
'pcDoctorCode���ַ������ͣ���ʾ����ҽ�����롣
'pcDoctorName���ַ������ͣ���ʾ����ҽ�����ơ�
'pcRecipNo���ַ������ͣ������ţ����ﴦ��ר�ã�סԺ���ա��˲�����Ҫ�������¹��ܣ�
'(1)������"ͳ�Ʒ���"��ʾ�����ţ����ڲ�ѯ�ͺ˶ԡ�
'��2������������ͬһ���˵Ķദ����飬��������ͬ�Ĳ���ҩ����ҩ���������Ŀ���뼲�����������Ϊ����Ӧ֢����봦������أ�����֢�Ͳ�����Ӧ�봦�����޹ء�
'pcNum���ַ������ͣ�ҩƷ�������������ﴦ�����ר�ã�סԺ���ա�Ϊ��7������Ԥ����
'pcNumUnit���ַ������ͣ�ҩƷ����������λ�����ﴦ�����ר�ã�סԺ���ա�Ϊ��7������Ԥ����
'pcPurpose���ַ������ͣ���ҩĿ��(0Ĭ��, 1����Ԥ����2�������ƣ�3Ԥ����4���ƣ�5Ԥ��+����)
'pcOprCode���ַ������ͣ�������ţ������Ӧ����������'��'��������ʾ��ҩΪ�ñ�Ŷ�Ӧ��������ҩ
'pcMediTime���ַ������ͣ���ҩʱ��  0����������ҩ
'                                  1����ǰ0.5h������ҩ
'                                  2����ǰ0.5-2h��
'3:                                   ��ǰ����2h��ҩ
'4:                                   ������ҩ
'5:                                   ������ҩ
'pcRemark���ַ������ͣ���ʾҽ����ע��Ϣ��
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��
'���ã��������ǰ�����ж�����ҩ��Ϣ��¼ʱʱ��Ҫ��ѭ�����ô��롣

'���벡�˹���ʷ��¼��Ϣ
Public Declare Function MDC_AddAller Lib "PASS4Invoke.dll" ( _
                    ByVal pcIndex As String, _
                    ByVal pcAllerCode As String, _
                    ByVal pcAllerName As String, _
                    ByVal pcAllerSymptom As String) As Integer

'�������:
'    pcIndex���ַ������ͣ���ʾ����Դ��ţ���ͬһѭ������ʱ��Ҫ�����¼��pcIndexֵ����Ψһ��
'pcAllerCode: �ַ������ͣ���ʾ����ԴΨһ�룬Ҫ����PASSϵͳ���ʱ���õĹ���ԴΨһ����ȫһ�£�����PASSϵͳ�޷�ʶ��˹�����Ϣ���˲�������Ϊ�ա�
'pcAllerName���ַ������ͣ���ʾ����Դ���ơ�
'pcAllerSymptom���ַ������ͣ���ʾ����Դ֢״��
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��
'���ã������ǰ�����ж���������Ϣ��¼ʱ��Ҫ��ѭ�����ô��롣

'3-1-4 ���벡����ϼ�¼��Ϣ
Public Declare Function MDC_AddMedCond Lib "PASS4Invoke.dll" ( _
                    ByVal pcIndex As String, ByVal pcDiseaseCode As String, _
                    ByVal pcDiseaseName As String, ByVal pcRecipNo As String) As Integer
'�������:
'pcIndex���ַ������ͣ���ʾ�����ţ���ͬһѭ������ʱ��Ҫ�����¼��pcIndexֵ����Ψһ��
'pcDiseaseCode���ַ������ͣ���ʾ���Ψһ�룬Ҫ����PASSϵͳ���ʱ���õ����Ψһ����ȫһ�£�����PASSϵͳ�޷�ʶ��������Ϣ���˲�������Ϊ�ա�
'pcDiseaseName���ַ������ͣ���ʾ������ơ�
'pcRecipNo�������š�
'����ֵ�����ͣ�2-�ɹ����������ظ������pcDiseaseCode
'1-�ɹ�
'0-ʧ��
'���ã������ǰ�����ж��������Ϣ��¼ʱʱ��Ҫ��ѭ�����ô��롣
                    
'3-1-5 ���벡��������¼��Ϣ
Public Declare Function MDC_AddOperation Lib "PASS4Invoke.dll" ( _
                    ByVal pcIndex As String, _
                    ByVal pcOprCode As String, _
                    ByVal pcOprName As String, _
                    ByVal pcIncisionType As String, _
                    ByVal pcOprStartDateTime As String, _
                    ByVal pcOprEndDateTime As String) As Integer
'�������:
'pcIndex���ַ������ͣ���ʾ������ţ���ͬһѭ������ʱ��Ҫ�����¼��pcIndexֵ����Ψһ��
'pcOprCode���ַ������ͣ���ʾ����Ψһ�룬Ҫ����PASSϵͳ���ʱ���õ�����Ψһ����ȫһ�£�����PASSϵͳ�޷�ʶ���������Ϣ���˲�������Ϊ�ա�
'pcOprName���ַ������ͣ���ʾ�������ơ�
'pcIncisionType���ַ������ͣ���ʾ�����п����͡�
'pcOprStartDateTime���ַ������ͣ���ʾ������ʼʱ�䣬��ʽΪ"yyyy-mm-dd hh:mm:ss"��
'pcOprEndDateTime���ַ������ͣ���ʾ��������ʱ�䣬��ʽΪ"yyyy-mm-dd hh:mm:ss"��
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��
'���ã������ǰ�����ж���������Ϣ��¼ʱʱ��Ҫ��ѭ�����ô��롣
'pcIndex:HIS��ҽ��ID;pcOprCode:��������Ŀ¼.���루���=��S����pcOprName:��������Ŀ¼.���� ��pcOprStartDateTime:����ҽ����¼.����ʱ�䣬pcOprEndDateTime:HIS��ȡ������ֹʱ�������մ���
'��������鲻��������ҩ ���Ƿ��ÿ���ҩ   ������ҩƷ���Ƿ񳬳���������������ں�����ҩ����ֵ���С���׼����û��������ܵġ�

'3-2��麯��

'3-2-1������ҩ��麯��
Public Declare Function MDC_DoCheck Lib "PASS4Invoke.dll" (ByVal piShowMode As Integer, ByVal piIsSave As Integer) As Integer
'�������:
'piShowMode�����ͣ���ʾ�������ʾģʽ�� 0-����ʾ���� 1-��ʾ���档
'piIsSave�����ͣ���ʾ������ɼ�ģʽ��0-���ɼ� 1-�ɼ���
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��

'3-3 ��ȡ���������
'3-3-1 ��ȡҩƷҽ����ʾ����
Public Declare Function MDC_GetWarningCode Lib "PASS4Invoke.dll" (ByVal pcIndex As String) As Integer
'�������:
'    pcIndex���ַ������ͣ���ʾҽ��Ψһ�룬Ҫ�������MDC_AddScreenDrug�������������pcIndex ֵ��ȫһ�¡�
'�ر�ע�⣺pcIndex����ʱ������������ȥ�����������ҽ������ߵľ�ʾ����ֵ��
'����ֵ�����ͣ����庬�����£�
'����ֵС��0����ʾ���ܳ����쳣������ҽ����һЩ������Ϣ��ȡֵ�������£�
'-1-��ҩƷ��PASS�в����ڻ�δ��ԡ�
'-2-��ҩƷ���ڲ������ü���������˵��ˡ�
'-3-ҽ����ͣ���������м�⡣
'-4-ҽ�������ϣ������м�⡣
'                -5-ϵͳ���ó�Ժ��ҩ�����м�⡣
'                -9-�޿�ʼ�ͽ���ʱ�䡣
'        ����ֵ���ڻ����0�����س̶ȣ�ȡֵ����Ϊ���£�
'                0-������⣬�޼���������ơ�
'1-������⣬���Ϊ���ɻ����أ��ڵơ�
'                2-������⣬���Ϊ���Ƽ�����ơ�
'                3-������⣬���Ϊ���ã��ȵơ�
'                4-������⣬���Ϊ��ע���Ƶơ�



'��ȡһ��ҩƷҽ�����������ʾ���ں���
Public Declare Function MDC_ShowWarningHint Lib "PASS4Invoke.dll" (ByVal pcIndex As String) As Integer
'�������:
'pcIndex���ַ������ͣ���ʾҽ��Ψһ�룬Ҫ�������MDC_AddScreenDrug�������������pcIndex ֵ��ȫһ�¡�
'����ֵ�����ͣ�1-�ɹ�

'�ر�һ��ҩƷҽ�����������ʾ���ں���
Public Declare Function MDC_CloseWarningHint Lib "PASS4Invoke.dll" () As Integer
'�������:
'��
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��


'3-3-2��ȡ�������������
Public Declare Function MDC_GetResultItemCount Lib "PASS4Invoke.dll" (ByVal pcIndex As String) As Integer
'�������:
'pcIndex���ַ������ͣ���ʾҽ��Ψһ�룬Ҫ�������MDC_AddScreenDrug�������������pcIndex ֵ��ȫһ�¡�
'����ֵ�����ͣ���ʾ��ҽ��������������������

'3-3-3 ��ȡ�������ϸ��Ϣ����
Public Declare Function MDC_GetResultDetail Lib "PASS4Invoke.dll" (ByVal pcIndex As String) As String
'�������:
'pcIndex���ַ������ͣ���ʾҽ��Ψһ�룬Ҫ�������MDC_AddScreenDrug�������������pcIndex ֵ��ȫһ�¡������շ������м�����������������������Ľ����
'����ֵ���ַ�������XML��ʽ���ظ�ҽ����������������ϸ��Ϣ��
'

'4����Ϣ��ѯ�ຯ��
'4-1 ��ȡ��ѯ��Ŀ��Ч�Ժ���
Public Declare Function MDC_GetDrugRefEnabled Lib "PASS4Invoke.dll" (ByVal pcDrugUniqueCode As String, ByVal piQueryType As Integer) As String
'�������:
'pcDrugUniqueCode���ַ������ͣ���ʾҩƷΨһ�룬Ҫ����PASSϵͳ���ʱ���õ�ҩƷΨһ����ȫһ�£�����PASSϵͳ�޷�ʶ��ҩƷ��Ϣ���˲�������Ϊ�ա�
'piQueryType:����,��ʾ��ѯģ�顣�������£��ر�ע�⣺���������������ģ�飩��
'11-ҩƷ˵����
'21-ҩ��ר��
'31-������ҩ����
'41-�й�ҩ��
'51-��Ҫ��Ϣ(��������)
'61-�໥����
'62-ҩʳ����
'63-��������
'64-����Ũ��
'65-ҩ�����֢
'66-ҩ����Ӧ֢
'67-������Ӧ
'68-���𺦼���
'69-���𺦼���
'70-��ͯ��ҩ
'71-������ҩ
'72-������ҩ
'73-������ҩ
'74-������ҩ
'75-�Ա���ҩ
'76-ϸ����ҩ��
'
'����ֵ:
'1�� �������piQueryType�������򷵻ظ���Ŀ��ѯ��Ϣ�Ƿ���õ�����ֵ��0-������ >0-���á�
'2�� ���û�д���piQueryType���������ذ�ģ��˳����֯�õ��ַ�����

'4-2 ��ѯҩƷ��Ϣ����
Public Declare Function MDC_GetDrugQueryInfo Lib "PASS4Invoke.dll" ( _
                    ByVal pcDrugUniqueCode As String, _
                    ByVal pcDrugName As String, _
                    ByVal piQueryType As Integer, _
                    ByVal X As Integer, _
                    ByVal Y As Integer) As Integer
'�������:
'pcDrugUniqueCode���ַ������ͣ���ʾҩƷΨһ�룬Ҫ����PASSϵͳ���ʱ���õ�ҩƷΨһ����ȫһ�£�����PASSϵͳ�޷�ʶ��ҩƷ��Ϣ���˲�������Ϊ�ա�
'pcDrugName���ַ������ͣ���ʾҩƷ���ơ�
'piQueryType������,��ʾ��ѯģ�顣�������£��ر�ע�⣺���������������ģ�飩��
'                                        11-ҩƷ˵����
'                                        21-ҩ��ר��
'                                        31-������ҩ����
'                                        41-�й�ҩ��
'                                        51-��Ҫ��Ϣ(��������)
'                                        61-�໥����
'                                        62-ҩʳ����
'                                        63-��������
'                                        64-����Ũ��
'                                        65-ҩ�����֢
'                                        66-ҩ����Ӧ֢
'                                        67-������Ӧ
'                                        68-���𺦼���
'                                        69-���𺦼���
'                                        70-��ͯ��ҩ
'                                        71-������ҩ
'                                        72-������ҩ
'                                        73-������ҩ
'                                        74-������ҩ
'                                        75-�Ա���ҩ
'                                        76-ϸ����ҩ��
'X�����ͣ���ʾX���ꡣ
'Y�����ͣ���ʾY���ꡣ
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��

'����һ��ҩƷ��Ϣ����
Public Declare Function MDC_DoSetDrug Lib "PASS4Invoke.dll" (ByVal pcDrugUniqueCode As String, _
                            ByVal pcDrugName As String) As Integer
'�������:
'pcDrugUniqueCode���ַ������ͣ���ʾҩƷΨһ�룬Ҫ����PASSϵͳ���ʱ���õ�ҩƷΨһ����ȫһ�£�����PASSϵͳ�޷�ʶ��ҩƷ��Ϣ���˲�������Ϊ�ա�
'pcDrugName���ַ������ͣ���ʾҩƷ���ơ�
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��


'��ѯ�Ѵ���ҩƷ˵������Ч�Ժ���
Public Declare Function MDC_DoRefDrugEnable Lib "PASS4Invoke.dll" (ByVal piQueryType As Integer) As String
'�������:
'piQueryType: ���� , 11 - ҩƷ˵����
'����ֵ�����ͣ�>0��ʾ��Ч��


'��ѯĳһ��ҩƷ��Ϣ����

Public Declare Function MDC_DoRefDrug Lib "PASS4Invoke.dll" (ByVal piQueryType As Integer) As Integer
'�������:
'piQueryType:     ���� , ��ʾ��ѯģ��?��������:
'11-ҩƷ˵����
'                                        51-��Ҫ��Ϣ(��������)
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��

'4-3�رո������ں���
Public Declare Function MDC_CloseDrugHint Lib "PASS4Invoke.dll" () As Integer
'�������:��
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��

'6������ҩ�о����ں���
Public Declare Function MDC_DoMediStudy Lib "PASS4Invoke.dll" (ByVal pcUseTime As String) As Integer
'�������:
'pcUseTime���ַ�������ʾ������ڣ�Ƕ��ҽ������վʱҪ�󴫿գ����Գ������ҩ�о������ʱ�䡣
'����ֵ�����ͣ�1-�ɹ�
'0-ʧ��

'20 PASS�˳�����
Public Declare Function MDC_Quit Lib "PASS4Invoke.dll" () As Integer

'������Ϣ����
Public Declare Function MDC_AddJsonInfo Lib "PASS4Invoke.dll" (ByVal pcJson As String) As Integer
'�������:pcJson���ַ������ͣ�JSON��ʽ��druginfoΪҩƷ���ٲ�����Ϣ��diseaseinfoΪ��ϲ�����Ϣ
'/*�����ʽ
'{
'            "type":"jsontype",
'            "screentype":"1"
'        },
'        {
'            "type":"druginfo",
'            "index":"drug001",
'            "driprate":"60",
'            "driptime":"120"
'        },
'        {
'            "type":"diseaseinfo",
'            "index":"dis001",
'            "starttime":"2015-12-31 09:11:11",
'            "endtime":"2016-08-02 09:11:11"
'        },
'        {
'            "type":"otherrecipinfo",
'            "hiscode":"his001",
'            "index":" drug001",
'            "recipno":"2016-08-02 09:11:11",
'            "drugsource":"USER",
'            "druguniquecode":"123456",
'            "drugname":"��Ī���ֽ���",
'            "doseunit":"g",
'            "routesource":"USER",
'            "routecode":"1"",
'            "routename":"�ڷ�""
'        }
'*/
'����ֵ�����ͣ�1-�ɹ�;0-ʧ��
'���ã�������֯һС��JSON����ã�Ҳ������֯����JSON���á�



'*******PASS4.0**1-����Ƕ����������DLL����������*****************************


'*******PASS4.0**1-����ҩʦ��Ԥϵͳ��DLL����������*****************************
Public Declare Function MDC_GetTaskStatus Lib "PASS4Invoke.dll" ( _
    ByVal pcPatCode As String, ByVal pcInHospNo As String, _
    ByVal pcVisitCode As String, ByVal pcRecipNo As String, _
    ByVal piTaskType As Integer) As Integer
'�������:
'pcPatCode���ַ������ͣ���ʾ����ID�������pcVisitCodeΨһȷ��һ�����ˣ��˲�������Ϊ�ա�Ҫ����MDC_SetPatient���������pcVisitCode����ֵ��ȫһ�¡�
'pcInHospNo:�������ͣ���ʾ��������Ż�סԺ�ţ��˲�������Ϊ�ա�Ҫ����MDC_SetPatient���������pcInHospNo����ֵ��ȫһ�¡�
'pcVisitCode���ַ������ͣ���ʾ���˾��������סԺ�����������pcPatCodeΨһȷ��һ�����ˣ����HISϵͳû�д���Ϣ����ɴ���"1"��Ҫ����MDC_SetPatient���������pcInHospNo����ֵ��ȫһ�¡�
'pcRecipNo���ַ������ͣ����ﴫ�����ţ�סԺ��ҽ��Ψһ�롣����Ҫ����MDC_AddScreenDrug���������pcRecipNo����ֵ��ȫһ�£�
'           סԺҪ����MDC_AddScreenDrug���������pcindex����ֵ��ȫһ�¡���ע���ò������Դ��գ���ʾȡ�������״̬���������崦����ҽ���ϣ�
'piTaskType�����ͣ���ʾ�������ͣ�1��ʾסԺ���ˣ�Ĭ�ϣ���2��ʾ���ﲡ�ˡ�
'
'����ֵ�����ͣ���ʾҩʦ��Ԥ�����1-ͨ����0-����ͨ��
'���ã��������סԺҽ������վ������PASS4����ҩ���ӿ�MDC_DoCheck����PASS�����ʱ�ᵯ������1��û��PASS�����ʱ�ᵯ������2-1���г�ʱ����ʱ��������2-2��
                                   
'*******PASS4.0**1-����ҩʦ��Ԥϵͳ������DLL����������*****************************
Public Function MK_GetPara() As Boolean
        Dim arrList As Variant
        Dim strPara As String
        
        On Error GoTo errH
100     strPara = zlDatabase.GetPara(90001, glngSys, , "") '��ȡURLs �̶���ȡZLHIS ϵͳĬ��100
        '��ʽ������IP&&�������˿ں�
102     If strPara = "" Then strPara = "0" & G_STR_SPLIT & "" & G_STR_SPLIT & "0"
        '����ҩʦ��Ԥ��1-����;0-�رգ�;ҽԺ����:��Ĭ��Ϊ�հ�վ�㴫��;��Ϊ�մ���ָ��ֵ��;�Ƿ���ʾ����(1-��;0-��);�Ƿ����þ�Ĭʽ���((1-��;0-��))
104     arrList = Split(strPara, G_STR_SPLIT)
106     If UBound(arrList) >= 2 Then
            gblnPharmReview = Val(arrList(0)) = 1
            gstrHOSCODE = arrList(1)     'ҽԺ����
            gblnPrePregnancy = Val(arrList(2)) = 1
            If UBound(arrList) >= 3 Then gblnTEST = Val(arrList(3)) = 1
        Else
            gblnPharmReview = False
            gstrHOSCODE = ""     'ҽԺ����
            gblnPrePregnancy = False
            gblnTEST = False
            Exit Function
        End If
        MK_GetPara = True
        Exit Function
errH:
146     MsgBox "��ȡ����ʧ�ܣ�" & vbNewLine & "HZYY_GetPara:��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function MK_SetPara() As String
    MK_SetPara = IIf(gblnPharmReview, 1, 0) & G_STR_SPLIT & gstrHOSCODE & G_STR_SPLIT & IIf(gblnPrePregnancy, 1, 0) & G_STR_SPLIT & IIf(gblnTEST, 1, 0)
End Function
