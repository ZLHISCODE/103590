Attribute VB_Name = "mdl����ʡ"
Option Explicit

'��������ֵΪ0ʱ��ʾ�ɹ�,��0��ʾʧ��

'===============================================================================================================
'���ܣ�����ҽ�����ݿ�����ҽ��ʹ�õ������շ���ˮ��סԺ��ˮ��,���ڹ���HIS���ü�¼��ҽ�����ü�¼,
'      һ���շѶ�Ӧһ������,һ��סԺ��Ӧһ������
'��ڲ������������ͣ�0����/1סԺ��
'���ڲ������շ���ˮ��
'===============================================================================================================
Public Declare Function FGetRecCode Lib "HInterface.dll" (ByVal intRecType As Long, _
    ByVal strRecCode As String) As Long
    
'===============================================================================================================
'���ܣ���ȡ�α��˵Ļ�����Ϣ���ʻ���֧����Ϣ
'��ڲ��������ͣ�0����/1סԺ��,�շ���ˮ��,����֤��
'���ڲ�����0����ID,1����,2����,3��������,4���֤����,5��λID,6��λ����,7�Ա�(��/Ů),8��Ա���,9��������,10����,
'          11�������,12����ְ��,13�������ⲡ��(�ѻ�����'δ֪'/����δ����ʱ����'δ����'/������������������ⲡ��),
'          14����(����),15�����ۼ�סԺ����,16�ʻ�����,17�ʻ���֧,18֧���汾��,19����ͳ��֧���ۼ�,20����󲡻���֧���ۼ�,
'          21���깫��Ա����/��ҵ����֧���ۼ�,22������ͨ��������ۼ�,23������ͨ����������Χ�ڷ����ۼ�,
'          24������������������Χ�ڷ����ۼ�,25�������סԺ������Χ�ڷ����ۼ�,26������ͨסԺ�����ۼ�,
'          27������ͨסԺ������Χ�ڷ����ۼ�,28�����ͥ����סԺ������Χ�ڷ����ۼ�,29����1,30����2,
'          31���괢���ʻ�֧���ۼ�,32������������֧���ۼ�,33�����ֽ�֧���ۼ�,34�ʻ����
'===============================================================================================================
Public Declare Function FGetCardInfo Lib "HInterface.dll" (ByVal intRecType As Long, ByVal strRecCode As String, _
    ByVal strVoucherID As String, intInsID As Long, ByVal strCardID As String, ByVal STRNAME As String, _
    ByVal strAreaCode As String, ByVal strQueryID As String, ByVal strUnitID As String, ByVal strUnitName As String, _
    ByVal strSex As String, ByVal strKind As String, ByVal strBirthday As String, ByVal strNational As String, _
    ByVal strIndustry As String, ByVal strDuty As String, ByVal strChronic As String, ByVal strOthers1 As String, _
    sngInHosNum As Single, sngAccIn As Single, sngAccOut As Single, sngFeeNO As Single, _
    sngPubPay As Single, sngHelpPay As Single, sngSupplyPay As Single, sngOutpatSum As Single, _
    sngOutpatGen1 As Single, sngOutpatGen2 As Single, sngOutpatGen3 As Single, _
    sngInpatSum As Single, sngInpatGen1 As Single, sngInpatGen2 As Single, _
    sngOther1 As Single, sngOther2 As Single, sngBankAccPay As Single, sngOtrPay As Single, _
    sngCashPay As Single, sngAccLeft As Single) As Long

'===============================================================================================================
'���ܣ�����Һ�
'��ڲ������շ���ˮ��,�Һ����ר�Һţ���ͨ��...��,��������,�Һŷ���Ŀ����,�Һŷ���Ŀ����,�Һŷѽ��,
'          ���Ʒ���Ŀ����,���Ʒ���Ŀ����,���Ʒѽ��,�ѱ�,����Ա,����,����ģʽ(T�俨��/Fˢ��)
'���ڲ�����ͳ��֧��,�ʻ�֧��,�ֽ�֧��
'===============================================================================================================
Public Declare Function FOutpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strRegName As String, _
    ByVal strDepartName As String, ByVal strRegFeeCode As String, ByVal strRegFeeName As String, ByVal sngRegFee As Single, _
    ByVal strDiagFeeCode As String, ByVal strDiagFeeName As String, ByVal sngDiagFee As Single, ByVal strFeeType As String, _
    ByVal strOpCode As String, ByVal strRegDate As String, ByVal pRegMode As String, sngPubPay As Single, _
    sngAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'���ܣ�ȡ���Һţ��˻��Һŷ�
'��ڲ����������,����Ա����
'���ڲ�������
'===============================================================================================================
Public Declare Function FCancleOutpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String) As Long

'===============================================================================================================
'���ܣ����ýӿ���ҽ��ϵͳHT.INS����α����˷�����ҽԺ�����з�����ϸ������ҩƷ��������Ŀ��������ʩ�շѵ�
'��ڲ���������(0����/1סԺ),�շ���ˮ�ţ�סԺ�����ﲻͬ��,��Ŀ����('0'��ҩƷ/'1'ҩƷ),��Ŀ����(HIS����),��ϸ����,
'          ��Ŀ����,��λ,��񡢼��͵�,�ѱ���,����ҩ��־,����,Ӧ�۵���,ʵ�۵���,ÿ������,ʹ��Ƶ��,�÷�,ִ������,
'          �շ�Ա����,���ұ���,����ҽ������,��������
'���ڲ���������֧������,�����Ը����,����׼����
'�ѱ��봫����룺����ҽԺ��ȫ��ҩƷ����Ŀ������պ����ʹ�ã��ѱ�����Բ���
'����ҩ��־������룺0�Ǵ���ҩ��1����ҩ��2�շ���Ŀ
'===============================================================================================================
Public Declare Function FWriteFeeDetail Lib "HInterface.dll" (ByVal intRecType As Long, _
    ByVal strRecCode As String, ByVal strItmFlag As String, ByVal strItmCode As String, ByVal strAliasCode As String, _
    ByVal strItmName As String, ByVal strItmUnit As String, ByVal strItmDesc As String, ByVal strFeeCode As String, _
    ByVal strOTCCode As String, ByVal sngQuantity As Single, ByVal sngPharPrice As Single, ByVal sngFactPrice As Single, _
    ByVal sngDosage As Single, ByVal strFrequency As String, ByVal strUsage As String, ByVal sngDays As Single, _
    ByVal strOpCode As String, ByVal strDepCode As String, ByVal strDocCode As String, ByVal strRecDate As String, _
    sngRate As Single, sngSelfFee As Single, sngDeduct As Single) As Long

'===============================================================================================================
'���ܣ��������¼����Ϻ���Խ��㣻����ʱ�ϴ����ü�¼
'��ڲ������շ���ˮ��,����Ա��,�Ƿ�ʹ���ʻ�(��/��),���ұ���,ҽ������,ҽ�Ʒ�ʽ,ҽ�����,('A'),��������,����1,����2,��ע
'���ڲ�����0������ˮ��,1�ܷ���,2������Χ�ڷ���,3�Ը�����,4�Էѷ���,5�𸶱�׼,6ͳ��֧��,7ͳ���Ը�,8�󲡾�������֧��,
'          9�󲡾��������Ը�,10����Ա����֧��/��ҵ����֧��,11����Ա����֧��/��ҵ�����Ը�,12��������֧��,
'          13����ҽ���ʻ�֧��,14���˴����ʻ�֧��,15�ֽ�֧��
'===============================================================================================================
Public Declare Function FTryOutpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String, _
    ByVal strUseAcc As String, ByVal strDepCode As String, ByVal strDocCode As String, ByVal strMedMode As String, _
    ByVal strRecClass As String, ByVal strICDMode As String, ByVal strICD As String, ByVal sngOther1 As Single, _
    ByVal sngOther2 As Single, ByVal strMemo As String, ByVal strBillCode As String, sngSumFee As Single, _
    sngGenFee As Single, sngFirstPay As Single, sngSelfFee As Single, sngPayLevel As Single, _
    sngPubPay As Single, sngPubSelf As Single, sngHelpPay As Single, sngHelpSelf As Single, _
    sngSupplyPay As Single, sngSupplySelf As Single, sngOtrPay As Single, sngMedAccPay As Single, _
    sngBankAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'���ܣ��������¼����Ϻ���㣬�ڵ��ô˺���ǰ������������Խ��㺯�����������ط��ù�������������¼���
'��ڲ������շ���ˮ��,����Ա��,�Ƿ�ʹ���ʻ�(��/��),���ұ���,ҽ������,ҽ�Ʒ�ʽ,ҽ�����,('A'),��������,����1,����2,��ע
'���ڲ�����0������ˮ��,1�ܷ���,2������Χ�ڷ���,3�Ը�����,4�Էѷ���,5�𸶱�׼,6ͳ��֧��,7ͳ���Ը�,8�󲡾�������֧��,
'          9�󲡾��������Ը�,10����Ա����֧��/��ҵ����֧��,11����Ա����֧��/��ҵ�����Ը�,12��������֧��,
'          13����ҽ���ʻ�֧��,14���˴����ʻ�֧��,15�ֽ�֧��
'===============================================================================================================
Public Declare Function FOutpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String, _
    ByVal strUseAcc As String, ByVal strDepCode As String, ByVal strDocCode As String, ByVal strMedMode As String, _
    ByVal strRecClass As String, ByVal strICDMode As String, ByVal strICD As String, ByVal sngOther1 As Single, _
    ByVal sngOther2 As Single, ByVal strMemo As String, ByVal strBillCode As String, sngSumFee As Single, _
    sngGenFee As Single, sngFirstPay As Single, sngSelfFee As Single, sngPayLevel As Single, _
    sngPubPay As Single, sngPubSelf As Single, sngHelpPay As Single, sngHelpSelf As Single, _
    sngSupplyPay As Single, sngSupplySelf As Single, sngOtrPay As Single, sngMedAccPay As Single, _
    sngBankAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'���ܣ�ȡ�������շѣ���Ҫ�����շ���ˮ��
'��ڲ������շ���ˮ��,������ˮ��,����Ա����
'���ڲ�����0�ܷ���,1������Χ�ڷ���,2�Ը�����,3�Էѷ���,4�𸶱�׼,5ͳ��֧��,6ͳ���Ը�,7�󲡾�������֧��,
'          8�󲡾��������Ը�,9����Ա/��ҵ����֧��,10����Ա/��ҵ�����Ը�,11��������֧��,12����ҽ���ʻ�֧��,
'          13���˴����ʻ�֧��,14�ֽ�֧��
'===============================================================================================================
Public Declare Function FCancelOutpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strBillCode As String, _
    ByVal strOpCode As String, sngSumFee As Single, sngGenFee As Single, sngFirstPay As Single, _
    sngSelfFee As Single, sngPayLevel As Single, sngPubPay As Single, sngPubSelf As Single, _
    sngHelpPay As Single, sngHelpSelf As Single, sngSupplyPay As Single, sngSupplySelf As Single, _
    sngOtrPay As Single, sngMedAccPay As Single, sngBankAccPay As Single, _
    sngCashPay As Single) As Long

'===============================================================================================================
'���ܣ��α�������Ժʱ����Ժ�Ǽǣ����Ǽ���Ϣ��¼���
'��ڲ�����סԺ�ţ��շ���ˮ�ţ�,ҽ�Ʒ�ʽ,ҽ�����,����Ա����,��Ժ����,ICD�������('A'),��Ժ���(ICD10����),
'          ���Ҵ���,��������,��Ժҽ������
'���ڲ������������ۼ�סԺ����
'===============================================================================================================
Public Declare Function FInpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strMedMode As String, _
    ByVal strMedClass As String, ByVal strRegOpCode As String, ByVal strBegDate As String, ByVal strICDMode As String, _
    ByVal strICD As String, ByVal strDepCode As String, ByVal strSecCode As String, ByVal strRegDoc As String, _
    sngInHosNum As Single) As Long

'===============================================================================================================
'���ܣ��α����˵Ǽ�סԺ�󲻴���סԺ����ȡ���ǼǴ���
'��ڲ�����סԺ�ţ��շ���ˮ�ţ�,����Ա����
'���ڲ�������
'===============================================================================================================
Public Declare Function FCancelInpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String) As Long

'===============================================================================================================
'���ܣ��α�����סԺ�еǼ���Ϣ�������ת�Ƶȣ����øú����޸ĵǼ���Ϣ
'��ڲ�����סԺ��,ҽ�Ʒ�ʽ,ҽ�����,����Ա����,���￪ʼ����,ICD�������('A'),��Ժ���(ICD10����),���Ҵ���,
'          ��������,��Ժҽ������
'���ڲ�������
'===============================================================================================================
Public Declare Function FChgInpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strMedMode As String, _
    ByVal strMedClass As String, ByVal strRegOpCode As String, ByVal strBegDate As String, ByVal strICDMode As String, _
    ByVal strICD As String, ByVal strDepCode As String, ByVal strSecCode As String, _
    ByVal strRegDoc As String) As Long

'===============================================================================================================
'���ܣ��α����˳�Ժ�Ĳ�����������Ժ�Ǽǣ������ǽ��˴���
'��ڲ�����סԺ��,����Ա����,��Ժ����,��Ժԭ��,ICD�������('A'),��Ժ���(ICD10����),��Ժҽ������
'���ڲ�������
'��Ժԭ������룺1������2��ת��3δ����4������5תԺ��6ת�⣻9����
'===============================================================================================================
Public Declare Function FInpatLeave Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOutOpCode As String, _
    ByVal strEndDate As String, ByVal strOutCause As String, ByVal strICDMode As String, ByVal strICD As String, _
    ByVal strOutDoc As String) As Long

'===============================================================================================================
'���ܣ�ҽԺ����Ժ���˵ķ��ý���Ԥ���㣬������ȡסԺѺ��
'      ������¿���ֱ�ӵ���סԺ�������㣺�������������ο�����Ϣ����ҽ�����ؿ�
'��ڲ�����סԺ��,����Ա����,�Ƿ�ʹ���ʻ�(��/��),���㷽ʽ,��������,����1,����2,��ע
'���ڲ�����0��'UnKnown'��,1�ܷ���,2������Χ�ڷ���,3�Ը�����,4�Էѷ���,5�𸶱�׼,6ͳ��֧��,7ͳ���Ը�,
'          8�󲡾�������֧��,9�󲡾��������Ը�,10����Ա/��ҵ����֧��,11����Ա/��ҵ�����Ը�,12��������֧��,
'          13����ҽ���ʻ�֧��,14���˴����ʻ�֧��,15�ֽ�֧��
'���㷽ʽ������룺0�������㣻1��;����
'===============================================================================================================
Public Declare Function FTryInpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String, _
    ByVal strUseAcc As String, ByVal intLiquiMode As String, ByVal strRefundID As String, ByVal sngOther1 As Single, _
    ByVal sngOther2 As Single, ByVal strMemo As String, ByVal strBillCode As String, sngSumFee As Single, _
    sngGenFee As Single, sngFirstPay As Single, sngSelfFee As Single, sngPayLevel As Single, _
    sngPubPay As Single, sngPubSelf As Single, sngHelpPay As Single, sngHelpSelf As Single, _
    sngSupplyPay As Single, sngSupplySelf As Single, sngOtrPay As Single, sngMedAccPay As Single, _
    sngBankAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'���ܣ�ҽԺ����Ժ���˵ķ��ý���Ԥ���㣬������ȡסԺѺ��
'      ������¿���ֱ�ӵ���סԺ�������㣺�������������ο�����Ϣ����ҽ�����ؿ�
'��ڲ�����סԺ��,����Ա����,�Ƿ�ʹ���ʻ�(��/��),���㷽ʽ,��������,����1,����2,��ע
'���ڲ�����0������ˮ��,1�ܷ���,2������Χ�ڷ���,3�Ը�����,4�Էѷ���,5�𸶱�׼,6ͳ��֧��,7ͳ���Ը�,
'          8�󲡾�������֧��,9�󲡾��������Ը�,10����Ա/��ҵ����֧��,11����Ա/��ҵ�����Ը�,12��������֧��,
'          13����ҽ���ʻ�֧��,14���˴����ʻ�֧��,15�ֽ�֧��
'���㷽ʽ������룺0�������㣻1��;����
'===============================================================================================================
Public Declare Function FInpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String, _
    ByVal strUseAcc As String, ByVal intLiquiMode As String, ByVal strRefundID As String, ByVal sngOther1 As Single, _
    ByVal sngOther2 As Single, ByVal strMemo As String, ByVal strBillCode As String, sngSumFee As Single, _
    sngGenFee As Single, sngFirstPay As Single, sngSelfFee As Single, sngPayLevel As Single, _
    sngPubPay As Single, sngPubSelf As Single, sngHelpPay As Single, sngHelpSelf As Single, _
    sngSupplyPay As Single, sngSupplySelf As Single, sngOtrPay As Single, sngMedAccPay As Single, _
    sngBankAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'���ܣ�����סԺ������Ϣ���α����˷�����Ժ״̬
'��ڲ�����סԺ��,������ˮ��,����Ա����
'���ڲ�����0�ܷ���,1������Χ�ڷ���,2�Ը�����,3�Էѷ���,4�𸶱�׼,5ͳ��֧��,6ͳ��֧��,7�󲡾�������֧��,
'          8�󲡾�������֧��,9����Ա/��ҵ����֧��,10����Ա/��ҵ����֧��,11��������֧��,12����ҽ���ʻ�֧��,
'          13���˴����ʻ�֧��,14�ֽ�֧��
'===============================================================================================================
Public Declare Function FCancelInpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, _
    ByVal strBillCode As String, ByVal strOpCode As String, sngSumFee As Single, sngGenFee As Single, _
    sngFirstPay As Single, sngSelfFee As Single, sngPayLevel As Single, sngPubPay As Single, _
    sngPubSelf As Single, sngHelpPay As Single, sngHelpSelf As Single, sngSupplyPay As Single, _
    sngSupplySelf As Single, sngOtrPay As Single, sngMedAccPay As Single, _
    sngBankAccPay As Single, sngCashPay As Single) As Long

'===============================================================================================================
'���ܣ�1����ҽԺ�����סԺ����ʱ��������ҽ�����ķ��ص��ʻ���ͳ����Ը�����ֵ�֮�Ͳ���ҽԺHISϵͳ�е��ܷ��ã�
'         ��Ҫ���ô˺������ҽ�����ؿ�����ĵ����ݣ������ϴ���
'      2��ҽԺHISϵͳ��Ҫ���°����иò��˵����ݵ��� '����¼��FWriteFeeDetail'�������뱾��ҽ�����ݿ⣬���
'         ���Ե����ϴ������ϴ���Ҳ��������ӿں����Զ��ϴ���
'��ڲ�����2,�շ���ˮ��
'���ڲ�������
'===============================================================================================================
Public Declare Function FSynData Lib "HInterface.dll" (ByVal intType As Long, ByVal strRecCode As String) As Long

'===============================================================================================================
'���ܣ����ô˺�����ҽ�����ؿ�HT.HIS����δ�ϴ������ݴ�����ҽ������
'��ڲ�����2,�շ���ˮ�ţ���Ϊ*��ʾ�ϴ����У�
'���ڲ�������
'===============================================================================================================
Public Declare Function FUpLoad Lib "HInterface.dll" (ByVal intType As Long, ByVal strRecCode As String) As Long

'===============================================================================================================
'���ܣ��ú���Ϊͨ�õ����ݵ��뺯�������ڴ�ҽԺHIS�������ݵ�ҽ�����ص����ݿ�
'��ڲ�����2,�շ���ˮ�ţ���Ϊ*��ʾ�ϴ����У�
'���ڲ���������(1����/2����Ա/3ҽ��/4ҩƷ�ֵ�/5������Ŀ),������Ϣ,������Ϣ,������Ϣ,������Ϣ,��ע,����״̬(I)
'A������
'   piType�� 1
'   psInfo1�� ���ұ���
'   psInfo2�� ��������
'B������Ա��
'   piType�� 2
'   psInfo1�� ����Ա����
'   psInfo2�� ����
'C��ҽ��:
'   piType�� 3
'   psInfo1�� ҽ������
'   psInfo2�� ҽ������
'   psInfo3���������ұ���
'   psInfo4��ְ��(����ҽʦ/������ҽʦ/����ҽʦ/��)
'D��ҩƷ�ֵ�
'   piType�� 4
'   psInfo1�� ҩƷ�Ա���
'   psInfo2����ҽ������
'   psInfo3�� ����(��Ʒ��)
'   psInfo4�����ۼ�
'   psRemark������|������λ|���|����|����
'   ��˵����psInfo4Ӧ�������λ��Ӧ
'           psRemark���������е���������|�ָ���Ϊ��ʱ����ַ�����
'E��������Ŀ��
'   piType �� 5
'   psInfo1����Ŀ�Ա���
'   psInfo2����ҽ������
'   psInfo3������
'   psInfo4�����ۼ�
'   psRemark����λ
'===============================================================================================================
Public Declare Function FImpInfo Lib "HInterface.dll" (ByVal intType As Long, ByVal strInfo1 As String, _
    ByVal strInfo2 As String, ByVal strInfo3 As String, ByVal strInfo4 As String, ByVal strRemark As String, _
    ByVal strOpStaus As String) As Long

'===============================================================================================================
'���ܣ����ڵ���ҽ�����ؿ����ݣ���ҽ��ҩƷ�ֵ䡢���ձ�ȣ��������Ϊ����������Լ�����͵����ļ�����ΪTXT�ı���
'      �ո�ָ������ļ�������·����·��������ʱ���Զ������ڽӿڶ�̬������·��
'��ڲ���������,�ļ���
'���ڲ�������
'===============================================================================================================
Public Declare Function FExpInfo Lib "HInterface.dll" (ByVal strTable As String, ByVal strFile As String) As Long

'===============================================================================================================
'���ܣ�����������ɺ������Ҫȡ���������㲢����¼�������ϸ������ô˺���
'��ڲ������շ���ˮ��,����Ա����
'���ڲ�������
'===============================================================================================================
Public Declare Function FCancelTryOutpatBalance Lib "HInterface.dll" (ByVal strRecCode As String, _
    ByVal strOpCode As String) As Long
    
'����Ϊ��̬���ӿ⺯�����岿��

Public gstrRecCode As String             '�շ���ˮ��
Public gblnReadCard As Boolean           '�Ƿ�ʹ�ö�����
Public gintҽ�Ʒ�ʽ As Long           '1��ͨ���2��ͨסԺ��3�������4�������ȣ�5���
Private intReturn As Long

'����Ϊҽ���ӿں����ݲ���

Public Function ҽ����ʼ��_����() As Boolean
    ҽ����ʼ��_���� = True
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte = 0, Optional lng����ID As Long = 0) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmIDentified As New frmIdentify����
    Dim strPatiInfo As String, cur��� As Currency, str������ As String
    Dim arr, datCurr As Date, str����� As String
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    strPatiInfo = frmIDentified.GetPatient(bytType, lng����ID)
    
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
                lng����ID = BuildPatiInfo(bytType, strPatiInfo, lng����ID, TYPE_����)

        '���ظ�ʽ:�м���벡��ID
        strPatiInfo = frmIDentified.mstrPatient & lng����ID & ";" & frmIDentified.mstrOther
        'д�������
        If bytType = 1 Then
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'˳���','''" & gstrRecCode & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
        ElseIf bytType = 3 Or bytType = 0 Then
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����֤��','''" & gstrRecCode & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
        End If
        Unload frmIDentified
    Else
        ��ݱ�ʶ_���� = ""
        MsgBox "δ��ȡ������Ϣ��", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    ��ݱ�ʶ_���� = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_���� = ""
End Function

Public Function �ҺŽ���_����(ByVal lng����ID As Long, cur��� As Currency) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strCode1 As String, strName1 As String, sngAcc1 As Single, lng����ID As Long
    Dim strCode2 As String, strName2 As String, sngAcc2 As Single
    Dim sngͳ��֧�� As Single, sng����֧�� As Single, sng�ֽ�֧�� As Single, cur������� As Currency
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    On Error GoTo errHandle:
    gstrSQL = "select * from �շ�ϸĿ where ���� = '�������Ʒ�'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
    strCode2 = rsTemp!����
    strName2 = rsTemp!����
    
    '��ȡ���˹Һ���Ϣ
    gstrSQL = "select a.no,a.id,a.����id,a.����,a.���˿���id,c.���� as ��������,a.ʵ�ս��,a.����Ա����,a.�ѱ�,to_char(a.����ʱ��,'yyyy-mm-dd HH24:MI:SS') as ʱ��,a.�վݷ�Ŀ,b.����,b.���� from ������ü�¼ a,�շ�ϸĿ b,���ű� c where ����id=[1] and b.id=a.�շ�ϸĿid and a.���˿���id=c.id order by �վݷ�Ŀ"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng����ID = rsTemp!����ID
    strCode1 = rsTemp!����
    strName1 = rsTemp!����
    sngAcc1 = rsTemp!ʵ�ս��
    rsTemp.MoveNext
'    strCode2 = rsTemp!����
'    strName2 = rsTemp!����
    sngAcc2 = rsTemp!ʵ�ս��
    '����ҽ���Һ�
    intReturn = FOutpatReg(gstrRecCode, "��ͨ��", rsTemp!��������, strCode1, strName1, sngAcc1, strCode2, strName2, _
        sngAcc2, rsTemp!�ѱ�, UserInfo.���, rsTemp!ʱ��, IIf(gblnReadCard, "F", "T"), sngͳ��֧��, sng����֧��, sng�ֽ�֧��)
        
    If intReturn <> 0 Then
        MsgBox "����ҽ������Һ�ʧ�ܣ�δ��ȡ������Ϣ��", vbInformation, gstrSysName
        �ҺŽ���_���� = False
        Exit Function
    End If
'    rsTemp.MoveFirst
'    gstrSQL = "zl_���˼��ʼ�¼_�ϴ�(" & rsTemp!ID & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
'    rsTemp.MoveNext
'    gstrSQL = "zl_���˷��ü�¼_�ϴ�(" & rsTemp!ID & ")"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '�������֧�������ڸ������򽫶���������ֽ�ʽ֧��
'    cur������� = �������_����(rsTemp!����ID)
'    If cur������� < sng����֧�� Then
'        sng�ֽ�֧�� = sng�ֽ�֧�� + sng����֧�� - cur�������
'        sng����֧�� = cur�������
'    End If
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(Date), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
            "," & TYPE_���� & "," & Year(Date) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + sng����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + sngͳ��֧�� & "," & intסԺ�����ۼ� + 1 & ",0,0,0,0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '���ս����¼
    '====================================ע��===================================='
    '�������Ҫ��gstrRecCode��strBillCode��lng����ID��Ӧ������,���ֵ����ڱ�ע��'
    '============================================================================'
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & _
            lng����ID & "," & Year(Date) & "," & _
            "0" & "," & cur�ʻ�֧���ۼ� + sng����֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + sngͳ��֧�� & "," & intסԺ�����ۼ� + 1 & ",NULL,NULL,NULL," & _
            sngͳ��֧�� + sng�ֽ�֧�� + sng����֧�� & "," & sng�ֽ�֧�� & ",NULL,NULL," & sngͳ��֧�� & ",NULL,NULL," & _
            sng����֧�� & ",NULL,NULL,NULL," & gstrRecCode & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    cur��� = sngͳ��֧��
    �ҺŽ���_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    �ҺŽ���_���� = False
End Function

Public Function �ҺŽ������_����(ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
'===============================================================================================================
'���ܣ�ȡ���Һţ��˻��Һŷ�
'��ڲ����������,����Ա����
'���ڲ�������
'===============================================================================================================
'Public Declare Function FCancleOutpatReg Lib "HInterface.dll" (ByVal strRecCode As String, ByVal strOpCode As String) As Integer
    gstrSQL = "Select ����֤�� From �����ʻ� Where ����id In (Select ����id From ������ü�¼ Where ����id=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "û���ҵ�������¼��ԭ���ݣ��������ܼ���ִ�С�", vbInformation, gstrSysName
        �ҺŽ������_���� = False
        Exit Function
    End If
    
    intReturn = FCancleOutpatReg(rsTemp!����֤��, UserInfo.����)
    If intReturn <> 0 Then
        MsgBox "ҽ���Һ��˷�ʱ��������δ��ô�����Ϣ��", vbInformation, gstrSysName
        �ҺŽ������_���� = False
        Exit Function
    End If
    
    �ҺŽ������_���� = True
End Function

Public Function �������_����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    
    If rsTemp.EOF Then
        �������_���� = 100000
    Else
        �������_���� = IIf(rsTemp("�ʻ����") = 0, 100000, rsTemp("�ʻ����"))
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    Dim cur���� As Currency, curͳ�� As Currency, cur��� As Currency, strBillCode As String
    Dim lng����ID As Long, rsTemp As New ADODB.Recordset, sngArrInfo(20) As Single
    
    On Error GoTo errHandle
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�з������ã����ܽ���Ԥ���㡣", vbInformation, gstrSysName
        �����������_���� = False
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    cur���� = 0: curͳ�� = 0
    gstrSQL = "Select * from �����ʻ� where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��Ԥ����", lng����ID)
    cur��� = rsTemp!�ʻ����
    
    intReturn = FCancelTryOutpatBalance(gstrRecCode, UserInfo.���)
    
    '���ݷ�����ϸ
    If ������ϸ����_����(0, rs��ϸ, 1) = False Then Exit Function
    
    '����Ԥ���㺯����������Ԥ����
    gstrSQL = "select a.����,a.���,a.id,c.���� as ����id,c.���� from ��Ա�� a,������Ա b,���ű� c where a.id=b.��Աid and a.����=[1] and c.id=b.����id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rs��ϸ!������))
    intReturn = FTryOutpatBalance(gstrRecCode, UserInfo.���, "��", rsTemp!����ID, rsTemp!���, CStr(gintҽ�Ʒ�ʽ), _
        IIf(gintҽ�Ʒ�ʽ = 3, "12", "11"), "A", "", 0, 0, "", strBillCode, sngArrInfo(0), sngArrInfo(1), sngArrInfo(2), _
        sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), sngArrInfo(8), sngArrInfo(9), _
        sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14))
'    If intReturn <> 0 Then
'        MsgBox "�ڽ���ҽ������Ԥ����ʱ��������δȡ�ô�����Ϣ��", vbInformation, gstrSysName
'        �����������_���� = False
'        Exit Function
'    End If
'
    cur���� = CCur(sngArrInfo(13) + sngArrInfo(12))
    curͳ�� = CCur(sngArrInfo(0) - sngArrInfo(14)) - cur����
    
    '�������������ʻ�����������ʻ���֧��������Ϊ�ʻ����
'    If cur���� > cur��� Then cur���� = cur���
    
'    MsgBox str������ϸ, vbInformation, "������ϸ"
    
    str���㷽ʽ = "ʡ�����ʻ�;" & cur���� & ";0"
    str���㷽ʽ = str���㷽ʽ & "|" & "ʡͳ�����;" & curͳ�� & ";0"
    �����������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
'        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ����
'        ����һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim curͳ�� As Currency, cur��� As Currency, strBillCode As String, datCurr As Date
    Dim lng����ID As Long, rsTemp As New ADODB.Recordset, sngArrInfo(20) As Single
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency, cur�������� As Currency
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    gstrSQL = "Select ����id From ������ü�¼ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng����ID = rsTemp(0)
    gstrSQL = "Select * from �����ʻ� where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    cur��� = rsTemp!�ʻ����
    
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select ��������id,������,b.����,c.��� from ������ü�¼ a,���ű� b,��Ա�� c where b.id=a.��������id and c.����=a.������ and a.����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    '����Ԥ���㺯����������Ԥ����
    strBillCode = Space(7)
    intReturn = FOutpatBalance(gstrRecCode, UserInfo.���, "��", rsTemp!����, rsTemp!���, CStr(gintҽ�Ʒ�ʽ), _
        IIf(gintҽ�Ʒ�ʽ = 3, "12", "11"), "A", "", 0, 0, "", strBillCode, sngArrInfo(0), sngArrInfo(1), sngArrInfo(2), _
        sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), sngArrInfo(8), sngArrInfo(9), _
        sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14))
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "�ڽ���ҽ���������ʱ��������δȡ�ô�����Ϣ��"
        �������_���� = False
        Exit Function
    End If
    
    '��ȡ�����ʻ�֧���͸����ֽ�֧��
    cur�����ʻ� = CCur(sngArrInfo(13) + sngArrInfo(12))
    curͳ�� = CCur(sngArrInfo(0) - sngArrInfo(14)) - cur�����ʻ�
    
'���ڲ�����0������ˮ��,1�ܷ���,2������Χ�ڷ���,3�Ը�����,4�Էѷ���,5�𸶱�׼,6ͳ��֧��,7ͳ���Ը�,8�󲡾�������֧��,
'          9�󲡾��������Ը�,10����Ա����֧��/��ҵ����֧��,11����Ա����֧��/��ҵ�����Ը�,12��������֧��,
'          13����ҽ���ʻ�֧��,14���˴����ʻ�֧��,15�ֽ�֧��
    
    curȫ�Ը� = CCur(sngArrInfo(13)) + CCur(sngArrInfo(14))
    cur�������� = CCur(sngArrInfo(0))
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
            "," & TYPE_���� & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + curͳ�� & "," & intסԺ�����ۼ� + 1 & "," & sngArrInfo(4) & "," & _
            sngArrInfo(4) & ",0,0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '���ս����¼
    '====================================ע��===================================='
    '�������Ҫ��gstrRecCode��strBillCode��lng����ID��Ӧ������,���ֵ����ڱ�ע��'
    '============================================================================'
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + curͳ�� & "," & intסԺ�����ۼ� + 1 & ",NULL,NULL,NULL," & _
            cur�������� & "," & curȫ�Ը� & ",NULL,NULL,NULL,NULL,NULL," & _
            cur�����ʻ� & ",NULL,NULL,NULL,'" & strBillCode & ";" & gstrRecCode & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    �������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ������ϸ����_����(lng����ID As Long, Optional rs��ϸIN As ADODB.Recordset = Nothing, Optional int�����־ As Integer = 1) As Boolean
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str������ As String, strҽ������ As String, objSystem As New FileSystemObject, objStream As TextStream
    Dim str���ұ�� As String, str�������� As String, lng����ID As Long
    Dim strTemp As String, sngRate As Single, sngSelfFee As Single, sngDeduct As Single
    Dim sng���� As Single, sng���� As Single
    Dim sngʵ�ս�� As Single
    
    On Error GoTo errHandle
    
    Set objStream = objSystem.OpenTextFile("C:\Trans.LOG", ForAppending, True, TristateFalse)
    If rs��ϸIN Is Nothing Then
        gstrSQL = "Select * From " & IIf(int�����־ = 1, "������ü�¼", "סԺ���ü�¼") & " Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
        Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    Else
        Set rs��ϸ = rs��ϸIN.Clone
    End If
    If rs��ϸ.EOF = True Then
'        MsgBox "û����Ҫ�ϴ����շѼ�¼", vbExclamation, gstrSysName
        If int�����־ = 1 Then
            ������ϸ����_���� = False
        Else
            ������ϸ����_���� = True
        End If
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    If int�����־ = 2 Then
        gstrSQL = "Select nvl(˳���,0) as ˳��� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Else
        gstrSQL = "Select nvl(����֤��,0) as ˳��� From �����ʻ� Where ����ID=[1] And ����=[2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    str������ = rsTemp!˳���: gstrRecCode = rsTemp!˳���
    objStream.WriteBlankLines 1
    While Not rs��ϸ.EOF
'        If IsNull(rs��ϸ!�Ƿ��ϴ�) Or rs��ϸ!�Ƿ��ϴ� = 0 Then
'0����ID
'1�շ����
'2�վݷ�Ŀ
'3���㵥λ
'4������
'5�շ�ϸĿID
'6����
'7����
'8ʵ�ս��
'9ͳ����
'10����֧������ID
'11�Ƿ�ҽ��
'12ժҪ
'13�Ƿ���
            On Error Resume Next
            Err = 0
            strTemp = Nvl(rs��ϸ!������)
            If Err.Number <> 0 Then
                strTemp = rs��ϸ!ҽ��
                sngʵ�ս�� = rs��ϸ!���
            Else
                sngʵ�ս�� = rs��ϸ!ʵ�ս��
            End If
            Err.Clear
            On Error GoTo errHandle
            gstrSQL = "select b.���,b.����,c.����,c.���� from ������Ա a,(select id,���,���� from ��Ա�� Where ����=[1]) b,(select id,����,���� from ���ű�) c where a.����id=c.id and a.��Աid=b.id"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, strTemp)
            If Not rsTemp.EOF Then
                strҽ������ = rsTemp!���
                str���ұ�� = rsTemp!����
                str�������� = rsTemp!����
            Else
                strҽ������ = ""
                str���ұ�� = ""
                str�������� = ""
            End If
'            gstrSQL = "Select * From �շ�ϸĿ Where ID=" & rs��ϸ!�շ�ϸĿID
            gstrSQL = "select a.��������,A.����,C.��Ŀ���� as ����,A.���㵥λ,B.����,decode(B.ҩƷ��Դ,'����',1,'����',2,'����',3,null) ��������,B.���" & _
                  " from �շ�ϸĿ A,ҩƷĿ¼ B,����֧����Ŀ C where A.id = C.�շ�ϸĿid and A.id=B.ҩƷid(+) and A.id =[1] And C.����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID), TYPE_����)
            strTemp = IIf(rs��ϸ!�շ���� = 5 Or rs��ϸ!�շ���� = 6 Or rs��ϸ!�շ���� = 7, "1", "0")
            
            If int�����־ <> 1 Then
                sng���� = rs��ϸ!���� * rs��ϸ!����
                sng���� = rs��ϸ!��׼����
            Else
                sng���� = rs��ϸ!����
                sng���� = rs��ϸ!����
            End If
            
            '�ϴ���ϸ
'��ڲ���������(0����/1סԺ),�շ���ˮ�ţ�סԺ�����ﲻͬ��,��Ŀ����('0'��ҩƷ/'1'ҩƷ),��Ŀ����(HIS����),��ϸ����,
'          ��Ŀ����,��λ,��񡢼��͵�,�ѱ���,����ҩ��־,����,Ӧ�۵���,ʵ�۵���,ÿ������,ʹ��Ƶ��,�÷�,ִ������,
'          �շ�Ա����,���ұ���,����ҽ������,��������
            objStream.WriteLine "FWriteFeeDetail(" & IIf(int�����־ = 2, 1, 0) & ",""" & str������ & """,""" & _
                strTemp & """,""" & rsTemp!���� & """,""" & rsTemp!���� & """,""" & rsTemp!���� & """,""" & Nvl(rsTemp!���㵥λ) & """,""" & Nvl(rsTemp!���) & ""","""",""" & _
                IIf(strTemp = "0", "2", IIf(rsTemp!�������� = "����ҩ" Or rsTemp!�������� = "����ҩ", "1", "0")) & """," & _
                sng���� & "," & sng���� & "," & sngʵ�ս�� / sng���� & ",0,"""","""",0,""" & _
                UserInfo.��� & """,""" & str���ұ�� & """,""" & strҽ������ & """,""" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:MM:SS") & """," & _
                sngRate & "," & sngSelfFee & "," & sngDeduct & ")"
            
            If int�����־ = 1 Then
                intReturn = FWriteFeeDetail(IIf(int�����־ = 2, 1, 0), str������, _
                    strTemp, rsTemp!����, rsTemp!����, rsTemp!����, Nvl(rsTemp!���㵥λ), Nvl(rsTemp!���), " ", _
                    IIf(strTemp = "0", "2", IIf(rsTemp!�������� = "����ҩ" Or rsTemp!�������� = "����ҩ", "1", "0")), _
                    sng����, sng����, sngʵ�ս�� / sng����, 0, " ", " ", 0, _
                    UserInfo.���, str���ұ��, strҽ������, Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:MM:SS"), _
                    sngRate, sngSelfFee, sngDeduct)
            Else
                intReturn = FWriteFeeDetail(IIf(int�����־ = 2, 1, 0), str������, _
                    strTemp, rsTemp!����, rsTemp!����, rsTemp!����, Nvl(rsTemp!���㵥λ), Nvl(rsTemp!���), " ", _
                    IIf(strTemp = "0", "2", IIf(rsTemp!�������� = "����ҩ" Or rsTemp!�������� = "����ҩ", "1", "0")), _
                    sng����, sng����, sngʵ�ս�� / sng����, 0, " ", " ", 0, _
                    UserInfo.���, str���ұ��, strҽ������, Format(rs��ϸ!����ʱ��, "yyyy-MM-dd HH:MM:SS"), _
                    sngRate, sngSelfFee, sngDeduct)
            End If
            If intReturn <> 0 Then
                MsgBox "�ڽ������ݴ���ʱ��������δȡ�ô�����Ϣ��", vbInformation, gstrSysName
                ������ϸ����_���� = False
                objStream.Close
                Exit Function
            End If
            
            If int�����־ <> 1 Then
                WriteInfo "NO:" & rs��ϸ!NO & "      ���:" & rs��ϸ!���
                gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rs��ϸ!ID & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            End If
'        End If
        rs��ϸ.MoveNext
    Wend
    If int�����־ = 2 Then
        intReturn = FUpLoad(2, gstrRecCode)
        If intReturn <> 0 Then
            MsgBox "�ڽ������ݴ���ʱ��������", vbInformation, gstrSysName
            ������ϸ����_���� = False
            objStream.Close
            Exit Function
        End If
    End If
    objStream.Close
    ������ϸ����_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    objStream.Close
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, sngArrInfo(20) As Single
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, curƱ���ܽ�� As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From ������ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    lng����ID = rsTemp("����ID")
    
    '��ȡ�ڽ���ʱ������շ���ˮ�źͽ�����ˮ��
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        ����������_���� = False
        Exit Function
    End If
    strRecCode = Mid(rsTemp!��ע, InStr(rsTemp!��ע, ";") + 1)
    strBillCode = Left(rsTemp!��ע, InStr(rsTemp!��ע, ";") - 1)
    '���ýӿ�������
    
'��ڲ������շ���ˮ��,������ˮ��,����Ա����
    intReturn = FCancelOutpatBalance(strRecCode, strBillCode, UserInfo.���, sngArrInfo(0), sngArrInfo(1), sngArrInfo(2), _
        sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), sngArrInfo(8), sngArrInfo(9), _
        sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14))
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "��������������ʱ��������δ��ô�����Ϣ��", vbInformation, gstrSysName
        ����������_���� = False
        Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        cur�����ʻ� * -1 & ",Null,Null,Null,'" & strBillCode & ";" & strRecCode & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    ����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str���� As String, str���ֱ��� As String
    Dim rsTmp As New ADODB.Recordset, str������ As String, datCurr As Date
    Dim lng����ID As Long, sngInHosNum As Single
    
    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,D.���� as ���ұ���,A.��Ժ����,A.����ҽʦ,C.����," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽ�_���� = False
        Exit Function
    End If
    
    '��ȡ��Ժ��ϣ����ֱ��룩
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, True, True) '��Ժ���
    If strInNote <> "" Then
        strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    '��ȡסԺҽʦ����
    gstrSQL = "Select ID,���,����,����,���˼��,������ѵ from ��Ա�� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rsTemp!����ҽʦ))
    
    '����ҽ���Ǽ�
    intReturn = FInpatReg(gstrRecCode, "2", IIf(gintҽ�Ʒ�ʽ = 3, "22", "21"), UserInfo.���, _
        Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:MM:SS"), "A", strInNote, rsTemp!���ұ���, " ", _
        IIf(rsTmp.EOF, " ", rsTmp!���), sngInHosNum)
    If intReturn <> 0 Then
        MsgBox "����ҽ����Ժ�Ǽ�ʱ��������δ��ȡ�ô�����Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽ�_���� = False
        Exit Function
    End If
     
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_���� = False
End Function

Public Function ת��ת��_����(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ���ת��ת����Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str���� As String, str���ֱ��� As String
    Dim rsTmp As New ADODB.Recordset, str������ As String, datCurr As Date
    Dim lng����ID As Long, sngInHosNum As Single
    
    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,D.���� as ���ұ���,A.��Ժ����,A.סԺҽʦ,C.˳���," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
        ת��ת��_���� = False
        Exit Function
    End If
    
    '��ȡ��Ժ��ϣ����ֱ��룩
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, True, True) '��Ժ���
    If strInNote <> "" Then
        strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
    End If
    
    '��ȡסԺҽʦ����
    gstrSQL = "Select ID,���,����,����,���˼��,������ѵ from ��Ա�� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rsTemp!סԺҽʦ))
    
    '����ҽ���Ǽ�
    intReturn = FChgInpatReg(rsTemp!˳���, "2", IIf(gintҽ�Ʒ�ʽ = 3, "22", "21"), UserInfo.���, _
        Format(rsTemp!��Ժ����, "yyyy-MM-dd HH24:MI:SS"), "A", strInNote, rsTemp!���ұ���, rsTemp!���ұ���, _
        rsTmp!���)
    If intReturn <> 0 Then
        MsgBox "����ҽ����Ժ������Ϣ�䶯ʱ��������δ��ȡ�ô�����Ϣ��", vbInformation, gstrSysName
        ת��ת��_���� = False
        Exit Function
    End If
     
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ת��ת��_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ת��ת��_���� = False
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    '��ȡ���������Ϣ
    gstrSQL = "Select * From �����ʻ� Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.EOF Then
        MsgBox "�����ҵ����˵������Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽǳ���_���� = False
        Exit Function
    End If
    
    '���ýӿڽ��г����Ǽ�
    intReturn = FCancelInpatReg(rsTemp!˳���, UserInfo.���)
    If intReturn <> 0 Then
        MsgBox "������Ժ�Ǽ�ʱ��������δ��ȡ������Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽǳ���_���� = False
        Exit Function
    End If
    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ��Ժ�Ǽǳ���_���� = False
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, StrInput As String, sngArrInfo(20) As Single
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String, lng����ID As Long
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, strTemp As String
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, rstTemp As String
    Dim curƱ���ܽ�� As Currency, lng��ҳID As Long
    Dim datCurr As Date, cur�����ʻ� As Currency
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ��,��ҳid From סԺ���ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng��ҳID = rsTemp!��ҳID
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "Select * from �����ʻ� where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    str������ = Nvl(rsTemp!˳���, "0")
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        סԺ�������_���� = False
        Exit Function
    End If
    cur�����ʻ� = rsTemp!�����ʻ�֧��
    strTemp = rsTemp!��ע
    '���ýӿ�������
'��ڲ�����סԺ��,������ˮ��,����Ա����
'���ڲ�����0�ܷ���,1������Χ�ڷ���,2�Ը�����,3�Էѷ���,4�𸶱�׼,5ͳ��֧��,6ͳ��֧��,7�󲡾�������֧��,
'          8�󲡾�������֧��,9����Ա/��ҵ����֧��,10����Ա/��ҵ����֧��,11��������֧��,12����ҽ���ʻ�֧��,
'          13���˴����ʻ�֧��,14�ֽ�֧��
    intReturn = FCancelInpatBalance(Mid(strTemp, InStr(strTemp, ";") + 1), Left(strTemp, InStr(strTemp, ";") - 1), _
        UserInfo.���, sngArrInfo(0), sngArrInfo(1), sngArrInfo(2), sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), _
        sngArrInfo(6), sngArrInfo(7), sngArrInfo(8), sngArrInfo(9), sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), _
        sngArrInfo(13), sngArrInfo(14))
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "סԺ�������ʱ��������", vbInformation, gstrSysName
        סԺ�������_���� = False
        Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� - 1 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� - 1 & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        cur�����ʻ� * -1 & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    סԺ�������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ����_����(lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
'        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ����
'        ����һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str����Ա As String, datCurr As Date, str������ As String
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur���� As Currency, cur����ͳ���޶� As Currency
    Dim cur���ͳ���޶� As Currency, cur�����Ը� As Currency, cur��� As Currency
    Dim cur�������� As Currency, curȫ�Ը� As Currency, cur���Ը� As Currency
    
    Dim cur�����ʻ�֧�� As Currency, cur�����ֽ�֧�� As Currency
    Dim curͳ��֧�� As Currency, curҽ��֧�� As Currency, cur����ҽ�� As Currency
    Dim strBillCode As String, sngArrInfo(20) As Single
    
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From סԺ���ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    If rs��ϸ.EOF = True Then
        Err.Raise 9000, gstrSysName, "û����д�շѼ�¼", vbExclamation, gstrSysName
        סԺ����_���� = False
        Exit Function
    End If
    lng����ID = rs��ϸ("����ID")
    
    gstrSQL = "Select nvl(˳���,0) as ˳���,�ʻ���� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    str������ = rsTemp!˳���
    cur��� = rsTemp!�ʻ����
    
    datCurr = zlDatabase.Currentdate
'��ڲ�����סԺ��,����Ա����,�Ƿ�ʹ���ʻ�(��/��),���㷽ʽ,��������,����1,����2,��ע
'���ڲ�����0��'UnKnown'��,1�ܷ���,2������Χ�ڷ���,3�Ը�����,4�Էѷ���,5�𸶱�׼,6ͳ��֧��,7ͳ���Ը�,
'          8�󲡾�������֧��,9�󲡾��������Ը�,10����Ա/��ҵ����֧��,11����Ա/��ҵ�����Ը�,12��������֧��,
'          13����ҽ���ʻ�֧��,14���˴����ʻ�֧��,15�ֽ�֧��
    strBillCode = Space(7)
    intReturn = FInpatBalance(str������, UserInfo.���, "��", 0, "IA01", 0, 0, "", strBillCode, sngArrInfo(0), _
        sngArrInfo(1), sngArrInfo(2), sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), _
        sngArrInfo(8), sngArrInfo(9), sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14))
    If intReturn <> 0 Then
        Err.Raise 9000, gstrSysName, "סԺ����Ԥ����ʱ��������δ��ô�����Ϣ��", vbInformation, gstrSysName
        סԺ����_���� = False
        Exit Function
    End If

    '��ȡ�����ʻ�֧���͸����ֽ�֧��
    cur�����ʻ�֧�� = CCur(sngArrInfo(13))
    cur�����ֽ�֧�� = CCur(sngArrInfo(14))
    cur����ҽ�� = CCur(sngArrInfo(7))
    curҽ��֧�� = CCur(sngArrInfo(9))
    curͳ��֧�� = CCur(sngArrInfo(5))
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
            "," & TYPE_���� & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur�����ʻ�֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + cur����ҽ�� + curҽ��֧�� + curͳ��֧�� & "," & intסԺ�����ۼ� + 1 & "," & cur���� & "," & _
            cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ�֧�� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� + cur����ҽ�� + curҽ��֧�� + curͳ��֧�� & "," & intסԺ�����ۼ� + 1 & _
            "," & cur����ҽ�� + curҽ��֧�� + curͳ��֧�� & ",NULL,NULL," & _
            cur�������� & "," & curȫ�Ը� & "," & cur���Ը� & ",NULL,NULL,NULL,NULL," & _
            cur�����ʻ�֧�� & ",NULL,NULL,NULL,'" & strBillCode & ";" & str������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    סԺ����_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_����(rs������ϸ As Recordset, lng����ID As Long, strҽ���� As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    Dim cur�����ʻ�֧�� As Currency, cur�����ֽ�֧�� As Currency
    Dim curͳ��֧�� As Currency, curҽ��֧�� As Currency, cur����ҽ�� As Currency
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset, strͬ�� As String
    Dim datCurr As Date, str������ As String, strBillCode As String
    Dim curCount As Currency, sngArrInfo(20) As Single, cur��� As Currency
    
    On Error Resume Next
    Kill "C:\Trans.LOG"
    On Error GoTo errHandle
    WriteInfo vbCrLf & "��ʼסԺԤ����"
    Set rs��ϸ = rs������ϸ.Clone
    If rs��ϸ.EOF = True Then
        MsgBox "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    curCount = 0
    While Not rs��ϸ.EOF
        curCount = curCount + rs��ϸ!���
        rs��ϸ.MoveNext
    Wend
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    strͬ�� = ""
reTrans:
    WriteInfo "��ʼ������ϸ"
    If ���ʴ���_����("", 2, strͬ��, lng����ID) = False Then Exit Function
    
    gstrSQL = "Select nvl(˳���,0) as ˳���,�ʻ���� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    str������ = rsTemp!˳���
    cur��� = rsTemp!�ʻ����
    
    datCurr = zlDatabase.Currentdate
    
    
    
    
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 '�������ӳ�Ժ�������ù���2007-09-28����
    Dim strTemp As String, strInNote As String
    Dim rsTmp As New ADODB.Recordset
    Dim lng��ҳID As Long, ��ҳID As Long


    gstrSQL = "Select * From �����ʻ� Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)

    gstrRecCode = rsTemp!˳���


    If rsTemp.EOF Then
        intReturn = FCancelInpatReg(gstrRecCode, UserInfo.���)

    ElseIf rsTemp(0) = 0 Then
        intReturn = FCancelInpatReg(gstrRecCode, UserInfo.���)


    Else
        gstrSQL = "select max(��ҳid) ��ҳid from ������ҳ where ����id=[1] And ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҳid", lng����ID, lng��ҳID)
        lng��ҳID = rsTemp!��ҳID
        gstrSQL = "select A.��Ժ����,D.���� as ��Ժ����,D.���� as ���ұ���,A.��Ժ����,A.סԺҽʦ," & _
                "A.��Ժ��ʽ,C.˳��� from ������ҳ A,�����ʻ� C,���ű� D Where A.����ID=C.����ID " & _
                "And A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)


        '��ȡ��Ժ��ϣ����ֱ��룩
        strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, True, True) '��Ժ���
        If strInNote <> "" Then
            strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
        End If

        '��ȡסԺҽʦ����
        gstrSQL = "Select ID,���,����,����,���˼��,������ѵ from ��Ա�� Where ����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rsTemp!סԺҽʦ))

        Select Case rsTemp!��Ժ��ʽ
            Case "����"
                strTemp = "1"
            Case "תԺ"
                strTemp = "5"
            Case "����"
                strTemp = "4"
            Case "��ת"
                strTemp = "2"
            Case "δ��"
                strTemp = "3"
            Case "ת��"
                strTemp = "6"
            Case Else
                strTemp = "9"
        End Select

    '��ڲ�����סԺ��,����Ա����,��Ժ����,��Ժԭ��,ICD�������('A'),��Ժ���(ICD10����),��Ժҽ������
        intReturn = FInpatLeave(rsTemp!˳���, UserInfo.���, Format(Nvl(rsTemp!��Ժ����, Date), "yyyy-MM-dd HH:MM:SS"), _
            strTemp, "A", strInNote, IIf(rsTmp.EOF, " ", rsTmp!���))

End If
''
    
    
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��ڲ�����סԺ��,����Ա����,�Ƿ�ʹ���ʻ�(��/��),���㷽ʽ,��������,����1,����2,��ע
'���ڲ�����0��'UnKnown'��,1�ܷ���,2������Χ�ڷ���,3�Ը�����,4�Էѷ���,5�𸶱�׼,6ͳ��֧��,7ͳ���Ը�,
'          8�󲡾�������֧��,9�󲡾��������Ը�,10����Ա/��ҵ����֧��,11����Ա/��ҵ�����Ը�,12��������֧��,
'          13����ҽ���ʻ�֧��,14���˴����ʻ�֧��,15�ֽ�֧��
    strBillCode = Space(7)
    intReturn = FTryInpatBalance(str������, UserInfo.���, "��", 0, "IA01", 0, 0, " ", strBillCode, sngArrInfo(0), _
        sngArrInfo(1), sngArrInfo(2), sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), _
        sngArrInfo(8), sngArrInfo(9), sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14))
    If intReturn <> 0 Then
        MsgBox "סԺ����Ԥ����ʱ��������", vbInformation, gstrSysName
        סԺ�������_���� = ""
        Exit Function
    End If

    '��ȡ�����ʻ�֧���͸����ֽ�֧��
    cur�����ʻ�֧�� = CCur(sngArrInfo(13) + sngArrInfo(12))
    cur�����ֽ�֧�� = CCur(sngArrInfo(14))
    cur����ҽ�� = CCur(sngArrInfo(7))
    curҽ��֧�� = CCur(sngArrInfo(9))
    curͳ��֧�� = CCur(sngArrInfo(5))
'    If curCount <> CCur(sngArrInfo(0)) Then
'        MsgBox "��ע�⣺ҽ�����ؽ������뵱ǰ���ݽ���" & vbCrLf, vbInformation, gstrSysName
'    End If
    WriteInfo "Ԥ���㷵��:" & CCur(sngArrInfo(0)) & "    ҽԺ:" & curCount
    If CCur(sngArrInfo(0)) <> curCount Then
        If MsgBox("��ע�⣺ҽ�����ؽ������뵱ǰ���ݽ���" & vbCrLf & "����Ժ����" & curCount & _
            "���������ķ��أ�" & CCur(sngArrInfo(0)) & vbCrLf & "�Ƿ���Ҫ��������ͬ����", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            intReturn = FSynData(2, str������)                'ȡ��סԺ�Խ���
            WriteInfo "����ͬ��"
            strͬ�� = "1"
            GoTo reTrans
        End If
    End If
    
    סԺ�������_���� = "ʡ�����ʻ�;" & cur�����ʻ�֧�� & ";0" '�������޸ĸ����ʻ�
    If curͳ��֧�� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|ʡͳ�����;" & curͳ��֧�� & ";0" '�������޸�ͳ��֧��
    End If
    If cur����ҽ�� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|ʡ��ͳ��;" & cur����ҽ�� & ";0"
    End If
    If curҽ��֧�� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|ʡ����Ա/��ҵ����֧��;" & curҽ��֧�� & ";0"
    End If
    WriteInfo "���Ԥ��"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    סԺ�������_���� = ""
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    Dim strTemp As String, rsTemp As New ADODB.Recordset, datCurr As Date, strInNote As String
    Dim rsTmp As New ADODB.Recordset
    
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "Select * From �����ʻ� Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.EOF Then
        MsgBox "�����ҵ����˵������Ϣ��", vbInformation, gstrSysName
        ��Ժ�Ǽ�_���� = False
        Exit Function
    End If
    gstrRecCode = rsTemp!˳���
    
    gstrSQL = "Select Sum(ʵ�ս��) From סԺ���ü�¼ Where ����id=[1] And ��ҳid=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If rsTemp.EOF Then
        intReturn = FCancelInpatReg(gstrRecCode, UserInfo.���)
        If intReturn <> 0 Then
'            MsgBox "������Ժ�Ǽ�ʱ��������δ��ȡ������Ϣ��", vbInformation, gstrSysName
            ��Ժ�Ǽ�_���� = False
            Exit Function
        End If
    ElseIf rsTemp(0) = 0 Then
        intReturn = FCancelInpatReg(gstrRecCode, UserInfo.���)
        If intReturn <> 0 Then
'            MsgBox "������Ժ�Ǽ�ʱ��������δ��ȡ������Ϣ��", vbInformation, gstrSysName
            ��Ժ�Ǽ�_���� = False
            Exit Function
        End If
    Else
        gstrSQL = "select A.��Ժ����,D.���� as ��Ժ����,D.���� as ���ұ���,A.��Ժ����,A.סԺҽʦ," & _
                "A.��Ժ��ʽ,C.˳��� from ������ҳ A,�����ʻ� C,���ű� D Where A.����ID=C.����ID " & _
                "And A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
        If rsTemp.EOF Then
            MsgBox "δ�ܻ�ȡ��Ժ���˵������Ϣ��", vbInformation, gstrSysName
            ��Ժ�Ǽ�_���� = False
            Exit Function
        End If
        
        '��ȡ��Ժ��ϣ����ֱ��룩
        strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, True, True) '��Ժ���
        If strInNote <> "" Then
            strInNote = Mid(strInNote, InStr(strInNote, "|") + 1)
        End If
        
        '��ȡסԺҽʦ����
        gstrSQL = "Select ID,���,����,����,���˼��,������ѵ from ��Ա�� Where ����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(rsTemp!סԺҽʦ))
        
        Select Case rsTemp!��Ժ��ʽ
            Case "����"
                strTemp = "1"
            Case "תԺ"
                strTemp = "5"
            Case "����"
                strTemp = "4"
            Case "��ת"
                strTemp = "2"
            Case "δ��"
                strTemp = "3"
            Case "ת��"
                strTemp = "6"
            Case Else
                strTemp = "9"
        End Select
    
    '��ڲ�����סԺ��,����Ա����,��Ժ����,��Ժԭ��,ICD�������('A'),��Ժ���(ICD10����),��Ժҽ������
        intReturn = FInpatLeave(rsTemp!˳���, UserInfo.���, Format(Nvl(rsTemp!��Ժ����, Date), "yyyy-MM-dd HH:MM:SS"), _
            strTemp, "A", strInNote, IIf(rsTmp.EOF, " ", rsTmp!���))
        If intReturn <> 0 Then
            MsgBox "����ҽ�����˳�Ժ�Ǽ�ʱ��������δ�ܻ�ȡ������Ϣ��", vbInformation, gstrSysName
            ��Ժ�Ǽ�_���� = False
            Exit Function
        End If
        
    End If
    
    '��HIS֮�еĻ������ݽ����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_���� = False
End Function

Public Function ҽ������_����() As Boolean
    ҽ������_���� = frmSet����.ShowME(TYPE_����)
End Function

Private Function Get����ID(strҽ���� As String, strҽ�����ı��� As String) As String
'���ܣ�ͨ��ҽ�����ĺ����ҽ�����������ID
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select ����ID from �����ʻ� where ���� = [1] and ҽ���� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, strҽ����)
    If Not rsTmp.BOF Then
        Get����ID = CStr(rsTmp("����ID"))
    Else
        Get����ID = ""
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Get����ID = ""
End Function

Public Function ���ʴ���_����(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset, lng��ҳID As Long
    
    gstrSQL = "Select Max(��ҳID) From ������ҳ Where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng��ҳID = rsTemp(0)
    
    If str��Ϣ = "" Then
        gstrSQL = " Select A.* From סԺ���ü�¼ A,�����ʻ� B" & _
                  " Where A.�����־=2 And Nvl(A.ʵ�ս��,0)<>0 And A.��¼״̬<>0 And Nvl(A.�Ƿ��ϴ�,0)=0 And nvl(A.���ӱ�־,0)<>9 " & _
                  " and A.����id=[1] And A.��ҳid=[2]" & _
                  " and A.����ID=B.����ID And B.����=[3]" & _
                  " order by A.NO,A.���"
    Else
        gstrSQL = " Select A.* From סԺ���ü�¼ A,�����ʻ� B" & _
                  " Where A.�����־=2 And Nvl(A.ʵ�ս��,0)<>0 And A.��¼״̬<>0 And nvl(A.���ӱ�־,0)<>9 " & _
                  " and A.����id=[1] And A.��ҳid=[2]" & _
                  " and A.����ID=B.����ID And B.����=[3]" & _
                  " order by A.NO,A.���"
    End If
    WriteInfo "��ȡ���˷��ü�¼:" & gstrSQL
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSQL, lng����ID, lng��ҳID, TYPE_����)
    If Not rsTemp.EOF Then
        WriteInfo "�ϴ���¼:" & rsTemp.RecordCount & "��"
        ���ʴ���_���� = ������ϸ����_����(0, rsTemp, 2)
    Else
        ���ʴ���_���� = True
        Exit Function
    End If
    If ���ʴ���_���� = True And rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rsTemp("ID") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
            rsTemp.MoveNext
        Wend
    End If
End Function


