Attribute VB_Name = "mdl����"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)
 
Public Type str_Out             '���ڲ����ṹ
    errtext As String           '������Ϣ
    out1 As String              '��1������ֵ
    out2 As String              '..2........
    out3 As String              '  .
    out4 As String              '  .
    out5 As String              '  .
    out6 As String              '  .
    out7 As String              '  .
    out8 As String              '  .
    out9 As String              '  .
    out10 As String             '  .
    out11 As String             '  .
    out12 As String             '  .
    out13 As String             '  .
    out14 As String             '  .
    out15 As String             '  .
    out16 As String             '  .
    out17 As String             '  .
    out18 As String             '  .
    out19 As String             '  .
    out20 As String             '  .
    out21 As String             '  .
    out22 As String             '  .
    out23 As String             '  .
    out24 As String             '  .
    out25 As String             '  .
    out26 As String             '  .
    out27 As String             '  .
    out28 As String             '  .
    out29 As String             '��29������ֵ
    out30 As String             '��30������ֵ
End Type

'������ڲ������;�Ϊ�ַ�������
'���м�ӳ���ֵ��Ӧ��str_Out�ṹ�л�ȡ
'�����漰�����ڻ�ʱ��Ĳ�����ӦдΪ"yyyy-MM-dd HH24:MI:SS"��ʽ���ַ���

'=========================================================================================================
'����˵��:��ѯҩƷ,������Ŀ,��λ,�������Ը�������ɱ����
'��ڲ���:ҽ����������,ҽԺ���,ҽԺ��Ŀ����,��ѯ���,��Ա���
'��ӳ��ڲ���:�Ը�������ɱ�����,��־
'    ��־˵��:1---�Ը�����,2---�ɱ�����(��ʾ���ɱ�����Ϊ�ý��,���������С�ڿɱ���ʱ,Ϊ�������)
'             3---�Ը�����|�ɱ�����(��ʾ�Ը�����Ϊ�������������౨�����ö�Ϊ�ɱ�����,���ڲ���ȫ���Է�)
'             4---û��ƥ��(ȫ���Է�),5---ҽԺû�и���Ŀ(ȫ���Է�)
'��������˵��:
'    ��ѯ���: 1---ҩƷ,2---������Ŀ,3---��λ,4---����
'    ��Ա���: 01---��ְ��Ա,02---������Ա
'=========================================================================================================
Public Declare Function fl_readzfbl Lib "cxybclient.dll" Alias "readzfbl" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal stryyxmbm As String, ByVal strcxlb As String, ByVal strrrlb As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'����˵��:ˢ��IC��,����ʾic������ϸ��Ϣ
'��ڲ���:ҽ����������,ҽԺ���,��ʾ��־
'��ӳ��ڲ���:ҽ����������,�����ʺ�,ic����,���֤����,����,�Ա�,��λ����,��λ����,��������,��Ա���,
'             IC�����,�𸶱�׼,����ҽ������޶�,����ҽ���ۼ�֧��,��ҽ���ۼ�֧��
'��������˵��:
'        �Ա�: 1---Ů,0---��
'    ��Ա���: 01---��ְ,02---����,03---�¸�
'    ��ʾ��־: 1---��ʾ,0---����ʾ(��ʾʱ,�ӿڿͻ��˽������Ի�����ʾIC������Ϣ,����������ֵ��������Ϣ)
'=========================================================================================================
Public Declare Function fl_readicxx Lib "cxybclient.dll" Alias "readicxx" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strxxbz As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'����˵��:ҽ������סԺ�����Ҫ�������Ե��ô˺���,��������Ϣ����ҽ��,��ҽ��������
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,ҽԺ��������,ҽԺ��������,��������,ԭ��
'         �����־��Ϊת��תԺ, ҽ������,�ز���־
'��ӳ��ڲ���:��
'��������˵��:
'    �����־: 0---��,1---��
'    �ز���־: 0---��,1---��
'=========================================================================================================
Public Declare Function fl_request Lib "cxybclient.dll" Alias "request" (ByVal strybjgbm As String, _
    ByVal stryybm As String, ByVal strybjzbh As String, ByVal stryyjbbm As String, _
    ByVal stryyjbmc As String, ByVal strsjrq As String, ByVal strsjyy As String, _
    ByVal strzybz As String, ByVal strysxm As String, ByVal strtbbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:ˢ��IC��,��ʾIC����Ϣ,�����סԺ�Ǽ�,�����ز�����ԺΨһ��ʶ���(ҽ��������)��
'��ڲ���:ҽ����������,ҽԺ���,��־,����Ա����,��������,�Ƿ���Ѫ
'��ӳ��ڲ���:ҽ��������,ҽ����������,�����ʺ�,ic����,���֤����,����,�Ա�,��λ����,��λ����,��������,
'             ��Ա���,IC�����,�𸶱�׼,����ҽ������޶�,����ҽ���ۼ�֧��,��ҽ���ۼ�֧��
'��������˵��:
'        ��־:0---����,1---סԺ
'  �Ƿ���Ѫ:�ӿں���δ˵��
'=========================================================================================================
Public Declare Function fl_reg Lib "cxybclient.dll" Alias "reg" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strbz As String, ByVal strczymc As String, ByVal strscrq As String, ByVal strsfdcx As String, _
    strOut As str_Out) As Boolean
    
'=========================================================================================================
'����˵��:д������Ϣ
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������,ҽԺ��������,ҽԺ��������,�ز���־,ҽ������
'         ¼��Ա����,���ұ��,��������,��������
'��ӳ��ڲ���:��
'��������˵��:
'    �ز���־:0---��,1---��
'=========================================================================================================
Public Declare Function fl_wrecipe Lib "cxybclient.dll" Alias "wrecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal stryyjbbm As String, ByVal stryyjbmc As String, _
    ByVal strtbbz As String, ByVal strysxm As String, ByVal strlryxm As String, ByVal strksbh As String, _
    ByVal strksmc As String, ByVal strcfrq As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'����˵��:�޸Ĵ�����Ϣ����ҽ�������ţ��������Ϊ�����޸�(����)��¼
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������,ҽԺ��������,ҽԺ��������,�ز���־,ҽ������
'         ¼��Ա����,���ұ��,��������,��������
'��ӳ��ڲ���:��
'��������˵��:
'    �ز���־:0---��,1---��
'=========================================================================================================
Public Declare Function fl_urecipe Lib "cxybclient.dll" Alias "urecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal stryyjbbm As String, ByVal stryyjbmc As String, _
    ByVal strtbbz As String, ByVal strysxm As String, ByVal strlryxm As String, ByVal strksbh As String, _
    ByVal strksmc As String, ByVal strcfrq As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:ɾ��������Ϣ
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function fl_drecipe Lib "cxybclient.dll" Alias "drecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:д������ϸ��Ϣ
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������,��ϸ���,ҽԺ��ϸ����,ҽԺ��ϸ����,����,���,���,
'         ��λ,����,����,ʱ��,¼����,��־
'��ӳ��ڲ���:��
'��������˵��:
'        ��־:1---ҩƷ,2---������Ŀ,3---��λ��
'ҽԺ��ϸ����:ΪҽԺҩƷ,������Ŀ,��λ�ѱ���(��Ӧ��־)
'ҽԺ��ϸ����:ΪҽԺҩƷ,������Ŀ,��λ������(��Ӧ��־)
'=========================================================================================================
Public Declare Function fl_wdetails Lib "cxybclient.dll" Alias "wdetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal strmxxh As String, ByVal stryymxbm As String, _
    ByVal stryymxmc As String, ByVal strcd As String, ByVal strgg As String, ByVal strlb As String, _
    ByVal strdw As String, ByVal strdj As String, ByVal strsl As String, ByVal strsj As String, _
    ByVal strlrr As String, ByVal strbz As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'����˵��:�޸Ĵ�����ϸ��Ϣ,��ҽ��������,�������,��ϸ���Ϊ�����޸�
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������,��ϸ���,ҽԺ��ϸ����,ҽԺ��ϸ����,����,���,���,
'         ��λ,����,����,ʱ��,¼����,��־
'��ӳ��ڲ���:��
'��������˵��:(ͬ��)
'=========================================================================================================
Public Declare Function fl_udetails Lib "cxybclient.dll" Alias "udetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal strmxxh As String, ByVal stryymxbm As String, _
    ByVal stryymxmc As String, ByVal strcd As String, ByVal strgg As String, ByVal strlb As String, _
    ByVal strdw As String, ByVal strdj As String, ByVal strsl As String, ByVal strsj As String, _
    ByVal strlrr As String, ByVal strbz As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'����˵��:ɾ��������ϸ��Ϣ
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������,��ϸ���
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function fl_ddetails Lib "cxybclient.dll" Alias "ddetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal strmxxh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:��ҽ�����˵ķ��ý���Ԥ����
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,��Ժ����,����Ա,�����־,��ʾ��־
'��ӳ��ڲ���:���úϼ�,���ⲡ�ַ���,���α����ʻ�֧��,���������ʻ�֧��,�ۼƷֶ��Ը�,ͳ���֧��,�𸶶�֧��,
'             ��λ֧��,�Էѷ���,�ؼ����Ը�,�������Ը�,�ؼ����,���η���,����ҽ�Ʊ���֧��,����ͳ������ۼ�,
'             ����ҽ�Ƽ����ۼ�,����ͳ������ۼ�,δ��������,ҽ��֧��,�����ֽ�֧��
'��������˵��:
'    ��ʾ��־:0---����ʾ,1---��ʾ
'    �����־:1---�Խ���,2---��;����
'=========================================================================================================
Public Declare Function fl_pcalc Lib "cxybclient.dll" Alias "pcalc" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcyrq As String, ByVal strczy As String, ByVal strjsbz As String, _
    ByVal strxsbz As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'����˵��:��ʽ����
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,��Ժ����,����Ա,��ʾ��־
'��ӳ��ڲ���:���úϼ�,���ⲡ�ַ���,���α����ʻ�֧��,���������ʻ�֧��,�ۼƷֶ��Ը�,ͳ���֧��,�𸶶�֧��,
'             ��λ֧��,�Էѷ���,�ؼ����Ը�,�������Ը�,�ؼ����,���η���,����ҽ�Ʊ���֧��,����ͳ������ۼ�,
'             ����ҽ�Ƽ����ۼ�,����ͳ������ۼ�,δ��������,ҽ��֧��,�����ֽ�֧��,�����ʻ����
'��������˵��:
'    ��ʾ��־:0---����ʾ,1---��ʾ
'=========================================================================================================
Public Declare Function fl_calc Lib "cxybclient.dll" Alias "calc" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcyrq As String, ByVal strczy As String, _
    ByVal strxsbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:סԺ�����,ȡ����ʽ����,���ص�����ǰ״̬;���������,���ɺ��ֵ���,��������¼
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,��ʾ��־
'��ӳ��ڲ���:��
'��������˵��:
'    ��ʾ��־:0---����ʾ,1---��ʾ
'=========================================================================================================
Public Declare Function fl_rollbackcalc Lib "cxybclient.dll" Alias "rollbackcalc" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strxsbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:ɾ����ҽ�������ŵ�������Ϣ,������Ժ�Ǽ�,����,������ϸ�ȡ������������ʽ����,����ʹ�øú���
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function fl_dall Lib "cxybclient.dll" Alias "dall" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:���ô����Ƿ���ɾ�����޸�
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function fl_canupdaterecipe Lib "cxybclient.dll" Alias "canupdaterecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:���ô�����ϸ�Ƿ���ɾ�����޸�
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������,��ϸ���
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function fl_canupdatedetails Lib "cxybclient.dll" Alias "canupdatedetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal strmxbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:����Ƿ��ܹ��ع�,סԺ�������ʹ�ô˺����ж�
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function fl_canrollback Lib "cxybclient.dll" Alias "canrollback" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:����Ƿ������ҽ���Ը�����,�����Ƿ��и���ʱ��������һ�μ���ʱ���ҩƷ,������Ŀ,����,��λ
'��ڲ���:ҽ����������,ҽԺ���,���ͱ�־
'��ӳ��ڲ���:��
'��������˵��:
'    ���ͱ�־:1---ҩƷ,2---������Ŀ,3---����,4---��λ
'=========================================================================================================
Public Declare Function fl_havenewzfbl Lib "cxybclient.dll" Alias "havenewzfbl" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strlxbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:����ҽ��������ʱ��
'��ڲ���:��
'��ӳ��ڲ���:ҽ��������ʱ��
'��������˵��:
'=========================================================================================================
Public Declare Function fl_getsystime Lib "cxybclient.dll" Alias "getsystime" (strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:����ҽ����������(��ҽ������ϵͳ������IC��)
'��ڲ���:��
'��ӳ��ڲ���:ҽ����������,ҽԺ����
'��������˵��:
'=========================================================================================================
Public Declare Function fl_getybjgbm Lib "cxybclient.dll" Alias "getybjgbm" (strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:������Ժҽ�Ƶ���(��ҽ������ϵͳ������IC��)
'��ڲ���:ҽ����������,ҽԺ���,�����ʺ�
'��ӳ��ڲ���:��Ժҽ�Ƶ���,ҽԺ����
'��������˵��:
'=========================================================================================================
Public Declare Function fl_getlastzyxx Lib "cxybclient.dll" Alias "getlastzyxx" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strgrzh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:д���޸���ҽ��������ҽԺ������Ϣ
'��ڲ���:���,ҽԺ��Ϣ����,ҽԺ��Ϣ����,����
'��ӳ��ڲ���:��
'��������˵��:
'        ���:1---ҩƷ,2---������Ŀ,3---��λ��,0---����
'        ����:�����Ϊ1,����Ϊ(01---����,02---����,03---����);
'             �����Ϊ2,����Ϊ���ұ���;�����Ϊ����,����Ϊ��
'=========================================================================================================
Public Declare Function fl_wyyglxx Lib "cxybclient.dll" Alias "wyyglxx" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strlb As String, ByVal stryyxxbm As String, _
    ByVal stryyxxmc As String, ByVal strqt As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:�޸��û���IC������
'��ڲ���:��
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function fl_changePassword Lib "cxybclient.dll" Alias "changepassword" (ByVal strybjgbm As String, _
    ByVal stryybm As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:У��ϵͳ��
'��ڲ���:��
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function fl_checkxtk Lib "cxybclient.dll" Alias "checkxtk" (strOut As str_Out) As Boolean

Public gstrOutPara As str_Out
Private mblnReturn As Boolean

Public Sub initType()
    With gstrOutPara
        .errtext = Space(1024)
        .out1 = Space(50)
        .out2 = Space(50)
        .out3 = Space(50)
        .out4 = Space(50)
        .out5 = Space(50)
        .out6 = Space(50)
        .out7 = Space(50)
        .out8 = Space(50)
        .out9 = Space(50)
        .out10 = Space(50)
        .out11 = Space(50)
        .out12 = Space(50)
        .out13 = Space(50)
        .out14 = Space(50)
        .out15 = Space(50)
        .out16 = Space(50)
        .out17 = Space(50)
        .out18 = Space(50)
        .out19 = Space(50)
        .out20 = Space(50)
        .out21 = Space(50)
        .out22 = Space(50)
        .out23 = Space(50)
        .out24 = Space(50)
        .out25 = Space(50)
        .out26 = Space(50)
        .out27 = Space(50)
        .out28 = Space(50)
        .out29 = Space(50)
        .out30 = Space(50)
    End With
End Sub

Public Sub TrimType()
    With gstrOutPara
        .errtext = Trim(.errtext)
        .out1 = Trim(.out1)
        .out2 = Trim(.out2)
        .out3 = Trim(.out3)
        .out4 = Trim(.out4)
        .out5 = Trim(.out5)
        .out6 = Trim(.out6)
        .out7 = Trim(.out7)
        .out8 = Trim(.out8)
        .out9 = Trim(.out9)
        .out10 = Trim(.out10)
        .out11 = Trim(.out11)
        .out12 = Trim(.out12)
        .out13 = Trim(.out13)
        .out14 = Trim(.out14)
        .out15 = Trim(.out15)
        .out16 = Trim(.out16)
        .out17 = Trim(.out17)
        .out18 = Trim(.out18)
        .out19 = Trim(.out19)
        .out20 = Trim(.out20)
        .out21 = Trim(.out21)
        .out22 = Trim(.out22)
        .out23 = Trim(.out23)
        .out24 = Trim(.out24)
        .out25 = Trim(.out25)
        .out26 = Trim(.out26)
        .out27 = Trim(.out27)
        .out28 = Trim(.out28)
        .out29 = Trim(.out29)
        .out30 = Trim(.out30)
    End With
End Sub

Public Function ҽ����ʼ��_����() As Boolean
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", TYPE_����)
    
    With rsTemp
        Do While Not .EOF
            If !������ = "ҽ����������" Then
                gstrҽ���������� = Nvl(!����ֵ)
            ElseIf !������ = "ҽԺ����" Then
                gstrҽԺ���� = Nvl(!����ֵ)
            End If
            .MoveNext
        Loop
    End With
    
    If gstrҽ���������� = "" Then
        MsgBox "�����б������������ñ��ղ�������ʹ�ñ��ӿڣ�[ҽ����������]", vbInformation, gstrSysName
        Exit Function
    End If
    
    ҽ����ʼ��_���� = True
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
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
    If lng����ID = 0 Then
        strTemp = "0"
    Else
        gstrSQL = "Select * From �����ʻ� where ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
        If rsTemp.EOF Then
            strTemp = "0"
        Else
            strTemp = Nvl(rsTemp!����֤��, 0)
        End If
    End If
    
'    WriteInfo "���ô��壬��ȡҽ��������Ϣ"
    strPatiInfo = frmIDentified.GetPatient(bytType, strTemp)
'    WriteInfo "ҽ��������Ϣ��" & strPatiInfo
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
'        WriteInfo "дҽ��������Ϣ"
        lng����ID = BuildPatiInfo(bytType, strPatiInfo, lng����ID, TYPE_����)
'        WriteInfo "дҽ��������Ϣ���"
        '���ظ�ʽ:�м���벡��ID
        strPatiInfo = frmIDentified.mstrPatient & lng����ID & ";" & frmIDentified.mstrOther
        str������ = frmIDentified.mstr������
        'д�������
'        WriteInfo "д�����ʻ���Ϣ"
        If bytType = 0 Or bytType = 5 Then
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'˳���','''" & str������ & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����֤��','''" & CLng(strTemp) + 1 & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
        End If
        gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'�Ҷȼ�','''" & frmIDentified.mstr�𸶱�׼ & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
'        WriteInfo "��������֤"
        Unload frmIDentified
    Else
        ��ݱ�ʶ_���� = ""
        MsgBox "ҽ��������Ϣ��ȡʧ��", vbInformation, gstrSysName
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

Public Function �������_����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, TYPE_����)
    
    If rsTemp.EOF Then
        �������_���� = 0
    Else
        �������_���� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'������rsDetail     ������ϸ(����)
'    ����ID         adBigInt, 19, adFldIsNullable
'    �շ����       adVarChar, 2, adFldIsNullable
'    �վݷ�Ŀ       adVarChar, 20, adFldIsNullable
'    ���㵥λ       adVarChar, 6, adFldIsNullable
'    ������         adVarChar, 20, adFldIsNullable
'    �շ�ϸĿID     adBigInt, 19, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ʵ�ս��       adSingle, 15, adFldIsNullable
'    ͳ����       adSingle, 15, adFldIsNullable
'    ����֧������ID adBigInt, 19, adFldIsNullable
'    �Ƿ�ҽ��       adBigInt, 19, adFldIsNullable
'    ժҪ           adVarChar, 200, adFldIsNullable
'    �Ƿ���       adBigInt, 19, adFldIsNullable
'    str���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim cur�Ը� As Currency, cur���� As Currency, cur��� As Currency, lngErr As Long
    Dim lng����ID As Long, rsTemp As New ADODB.Recordset, str������ϸ As String
    Dim strTemp As String, curTemp As Currency, str�Ը����� As String, str�ɱ����� As String
    
    On Error GoTo errHandle
    If rs��ϸ.RecordCount = 0 Then
        MsgBox "û�з��ã����ܽ���Ԥ���㡣", vbInformation, gstrSysName
        �����������_���� = False
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID"): lngErr = 1
    cur�Ը� = 0: cur���� = 0: lngErr = 2
    gstrSQL = "Select * from �����ʻ� where ����id=[1]"
    lngErr = 3
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��Ԥ����", lng����ID)
    lngErr = 4
    cur��� = rsTemp!�ʻ����: lngErr = 5
    strTemp = IIf(rsTemp!��ְ = "2", "02", "01"): lngErr = 4
    str������ϸ = ""
    While Not rs��ϸ.EOF
        gstrSQL = "select * from �շ�ϸĿ where id=[1]"
        lngErr = 6
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��Ԥ����", CLng(rs��ϸ!�շ�ϸĿID))
        lngErr = 7
        
        '��ȡ�շ�ϸĿ���Ը�����
        initType
        mblnReturn = fl_readzfbl(gstrҽ����������, gstrҽԺ����, rsTemp!��� & "_" & rsTemp!ID, _
            IIf(rsTemp!��� = "5" Or rsTemp!��� = "6" Or rsTemp!��� = "7", "1", IIf(rsTemp!��� = "J", "3", "2")), _
            strTemp, gstrOutPara): lngErr = 8
        TrimType
        
        If mblnReturn = False Then
            MsgBox "�ڻ�ȡ��Ŀ[" & rsTemp!ID & "]���Ը�����ʱ��ҽ���ӿڷ������´���" & Chr(13) & Chr(10) & gstrOutPara.errtext
            �����������_���� = False
            Exit Function
        End If
        Select Case gstrOutPara.out2
            Case "1"            '����Ϊ�Ը�����
                curTemp = rs��ϸ!ʵ�ս�� * (1 - CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0))): lngErr = 9
            Case "2"            '����Ϊ�����޶�
                curTemp = IIf(rs��ϸ!ʵ�ս�� > CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0)), CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0)), rs��ϸ!ʵ�ս��): lngErr = 10
            Case "3"            '���Ը��������㱨���������ڿɱ������ȡ�ɱ�����
                str�Ը����� = Left(gstrOutPara.out1, InStr(gstrOutPara.out1, "|") - 1): lngErr = 11
                str�ɱ����� = Mid(gstrOutPara.out1, InStr(gstrOutPara.out1, "|") + 1): lngErr = 12
                str�Ը����� = IIf(IsNumeric(str�Ը�����), str�Ը�����, 0): lngErr = 13
                str�ɱ����� = IIf(IsNumeric(str�ɱ�����), str�ɱ�����, 0): lngErr = 14
                curTemp = rs��ϸ!ʵ�ս�� * (1 - CCur(str�Ը�����)): lngErr = 15
                curTemp = IIf(curTemp > CCur(str�ɱ�����), CCur(str�ɱ�����), curTemp): lngErr = 16
            Case "4", "5"       '�Ը�����Ϊ100%
                curTemp = 0
        End Select
        str������ϸ = str������ϸ & "��Ŀ����:" & rsTemp!���� & "[" & rsTemp!��� & "_" & rsTemp!ID & "]�����Ը�����:[" & _
            gstrOutPara.out2 & "]" & gstrOutPara.out1 & "�����������:" & curTemp & Chr(13) & Chr(10)
        
        cur���� = cur���� + curTemp: lngErr = 17
        cur�Ը� = rs��ϸ!ʵ�ս�� - curTemp: lngErr = 18
        rs��ϸ.MoveNext: lngErr = 19
    Wend
    
    '�������������ʻ�����������ʻ���֧��������Ϊ�ʻ������ಿ�ּ����ֽ�֧��
    '�����Ҫ����1Ԫ��Ϊ����
    If cur���� > cur��� - 1 Then
        curTemp = cur���� - (cur��� - 1): lngErr = 20
        cur���� = cur��� - 1: lngErr = 21
        cur�Ը� = cur�Ը� + curTemp: lngErr = 22
    End If
    
'    MsgBox str������ϸ, vbInformation, "������ϸ"
    
    str���㷽ʽ = "�����ʻ�;" & cur���� & ";0": lngErr = 23
    �����������_���� = True
    Exit Function
errHandle:
    MsgBox "���������[����Ԥ����]ģ�飬��" & lngErr & "�У�������Ϣ��" & Chr(13) & Chr(10) & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency, Optional ByRef strAdvance As String) As Boolean
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
    Dim str����Ա As String, datCurr As Date, str������ As String, str���㷽ʽ As String
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur���� As Currency, cur����ͳ���޶� As Currency, cur����Ա���� As Currency
    Dim cur���ͳ���޶� As Currency, cur�����Ը� As Currency, cur��� As Currency
    Dim cur�������� As Currency, cur���Ը� As Currency, lng����ID As Long, str����Ա��� As String
    Dim blnOld As Boolean
    
    If gstrҽ���������� = "" Then
        Err.Raise 9000, gstrSysName, "׼����ȡҽ���������룬�����ϵͳ�����˿���"
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstrҽ���������� = gstrOutPara.out1
        gstrҽԺ���� = gstrOutPara.out2
    End If
    
    On Error GoTo errHandle
    gstrSQL = "Select * From ������ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    If rs��ϸ.EOF = True Then
        Err.Raise 9000, gstrSysName, "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    str����Ա = ToVarchar(IIf(IsNull(rs��ϸ("����Ա����")), UserInfo.����, rs��ϸ("����Ա����")), 20)
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_����

    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If rsTemp.State = 1 Then
        lng����ID = rsTemp("ID")
    Else
        �������_���� = False
        Exit Function
    End If
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����ID'," & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")

    '��Ҫ���ϴ�������ϸ
    If ������ϸ����_����(lng����ID) = False Then
        Err.Raise 9000, gstrSysName, "��¼�ϴ�ʱ��������", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select nvl(˳���,0) as ˳���,����id From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_����
    Set rsTemp = gcnOracle.Execute(gstrSQL)
    lng����ID = rsTemp!����ID
    str������ = rsTemp!˳���
    
    'ҽ����������, ҽԺ���, ҽ�������ţ� ��Ժ���ڣ�����Ա����ʾ��־
    datCurr = zlDatabase.Currentdate
    initType
'    mblnReturn = fl_pcalc(gstrҽ����������, gstrҽԺ����, str������, Format(datCurr, "yyyy-MM-dd"), str����Ա, "1", "0", gstrOutPara)
    mblnReturn = fl_calc(gstrҽ����������, gstrҽԺ����, str������, Format(datCurr, "yyyy-MM-dd"), str����Ա, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext, vbInformation, gstrSysName
        �������_���� = False
        Exit Function
    End If
    '��ӳ��ڲ���:1���úϼ�,2���ⲡ�ַ���,3���α����ʻ�֧��,4���������ʻ�֧��,5�ۼƷֶ��Ը�,6ͳ���֧��,7�𸶶�֧��,
    '             8��λ֧��,9�Էѷ���,10�ؼ����Ը�,11�������Ը�,12�ؼ����,13���η���,14����ҽ�Ʊ���֧��,15����ͳ������ۼ�,
    '             16����ҽ�Ƽ����ۼ�,17����ͳ������ۼ�,18δ��������,19ҽ��֧��,20�����ֽ�֧��,21�����ʻ����
    '���ӳ��ڲ���:22����Ա���;23����Ա����
    '��ȡ�����ʻ�֧���͸����ֽ�֧��
    cur�����ʻ� = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
    cur��� = CCur(gstrOutPara.out21)
    curȫ�Ը� = CCur(gstrOutPara.out20) + CCur(cur�����ʻ�)
    cur�������� = CCur(gstrOutPara.out1)
    cur���Ը� = CCur(gstrOutPara.out10) + CCur(gstrOutPara.out11)
    cur����Ա���� = CCur(gstrOutPara.out23)
    str����Ա��� = gstrOutPara.out22
    
    'ǿ�Ƹ�ֵ����Ϊ��Ԥ���㹦�ܣ����������ʾ�����ʻ�֧������ʽ����ʱ���˽������Ԥ����¼������ʽ����ʱ�޸����ʻ�֧��ȷδ���������
    str���㷽ʽ = "||�����ʻ�|" & cur�����ʻ� & "||����Ա����|" & cur����Ա����
    
    '�������
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        #If gverControl < 2 Then
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
        #Else
            strAdvance = str���㷽ʽ
            gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
        #End If
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
    End If
    #If gverControl < 2 Then
        blnOld = True
        frm������Ϣ.ShowME (lng����ID)
    #End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & Get����ID(CStr(strҽ����), CStr(TYPE_����)) & _
            "," & TYPE_���� & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & _
            cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & _
            Get����ID(CStr(strҽ����), CStr(TYPE_����)) & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur�������� & "," & curȫ�Ը� & "," & cur���Ը� & ",NULL,NULL," & cur����Ա���� & ",NULL," & _
            cur�����ʻ� & ",'" & str����Ա��� & "',NULL,NULL,'" & str������ & "'" & IIf(blnOld, "", ",1") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    �������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ������ϸ����_����(lng����ID As Long, Optional rs��ϸIN As ADODB.Recordset = Nothing) As Boolean
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str����Ա As String, cur��������, str������ As String, strBillNO As String
    Dim lng����ID As Long, str�������� As String, str���ֱ��� As String, int�ز���־ As Integer
    Dim str���ұ�� As String, str�������� As String, lng����ID As Long
    Dim str��ϸ���� As String, str��ϸ���� As String, str������ As String
    Dim strTemp As String, iLoop As Long
'    ����ID         adBigInt, 19, adFldIsNullable
'    �շ����       adVarChar, 2, adFldIsNullable
'    �վݷ�Ŀ       adVarChar, 20, adFldIsNullable
'    ���㵥λ       adVarChar, 6, adFldIsNullable
'    ������         adVarChar, 20, adFldIsNullable
'    �շ�ϸĿID     adBigInt, 19, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ����           adSingle, 15, adFldIsNullable
'    ʵ�ս��       adSingle, 15, adFldIsNullable
'    ͳ����       adSingle, 15, adFldIsNullable
'    ����֧������ID adBigInt, 19, adFldIsNullable
'    �Ƿ�ҽ��       adBigInt, 19, adFldIsNullable
'    ժҪ           adVarChar, 200, adFldIsNullable
'    �Ƿ���       adBigInt, 19, adFldIsNullable
'    str���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    
    On Error GoTo errHandle
    
    If gstrҽ���������� = "" Then
        MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstrҽ���������� = gstrOutPara.out1
        gstrҽԺ���� = gstrOutPara.out2
    End If
    If rs��ϸIN Is Nothing Then
        gstrSQL = "Select * From ������ü�¼ Where ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ����ID=[1]"
        Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    Else
        Set rs��ϸ = rs��ϸIN.Clone
    End If
    If rs��ϸ.EOF = True Then
'        MsgBox "û����Ҫ�ϴ����շѼ�¼", vbExclamation, gstrSysName
        ������ϸ����_���� = True
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    str����Ա = ToVarchar(UserInfo.����, 20)
    
'    gstrSQL = "select max(��ҳID) as ��ҳID from ������ҳ where ����ID =" & lng����ID
'    Call OpenRecordset(rsTemp, gstrsysname)
'    strBillNo = CStr(lng����ID) & "_" & CStr(rsTemp("��ҳID"))
    gstrSQL = "Select nvl(˳���,0) as ˳���,����ID,����,����֤�� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    str������ = rsTemp!����֤��
    str������ = rsTemp!˳���
    lng����ID = Nvl(rsTemp!����ID, 0)
'    gstrҽ���������� = rsTemp!����
    gstrSQL = "Select * From ���ղ��� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.EOF Then
        str�������� = "δ֪"
        str���ֱ��� = "0"
        int�ز���־ = 0
    Else
        str�������� = rsTemp!����
        str���ֱ��� = rsTemp!ID
        int�ز���־ = IIf(rsTemp!��� = 2, 1, 0)
    End If
    lng����ID = rs��ϸ!��������ID
    gstrSQL = "Select * From ���ű� where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    str���ұ�� = rsTemp!����
    str�������� = rsTemp!����
    
'    str������ = NVL(rs��ϸ!��ҳID, 0) & Right(rs��ϸ!NO, 2)
    'д������Ϣ
    initType
    mblnReturn = fl_wrecipe(gstrҽ����������, gstrҽԺ����, str������, str������, str���ֱ���, str��������, _
                         int�ز���־, Nvl(rs��ϸ!������, rs��ϸ!������), Nvl(rs��ϸ!����Ա����, UserInfo.����), str���ұ��, _
                         str��������, Format(rs��ϸ!�Ǽ�ʱ��, "yyyy-MM-dd"), gstrOutPara)
    TrimType
    If mblnReturn = False Then
        If InStr(gstrOutPara.errtext, "(YBYY.PRI_QTYL") > 0 Then
            ������ϸ����_���� = True
        Else
            MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
            ������ϸ����_���� = False
            Exit Function
        End If
    End If
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����֤��','" & CLng(str������) + 1 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    iLoop = 1
    'д������ϸ
    Do Until rs��ϸ.EOF
        gstrSQL = "Select * From �շ�ϸĿ Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��ϸ!�շ�ϸĿID))
        str��ϸ���� = rsTemp!ID
        str��ϸ���� = rsTemp!����
        initType
        If InStr(Nvl(rsTemp!���, " "), "��") > 0 Then
            strTemp = Left(rsTemp!���, InStr(rsTemp!���, "��") - 1)
        Else
            strTemp = Nvl(rsTemp!���, " ")
        End If
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������,��ϸ���,ҽԺ��ϸ����,ҽԺ��ϸ����,����,���,���,
'         ��λ,����,����,ʱ��,¼����,��־
        If (IsNull(rs��ϸ!�Ƿ��ϴ�) Or rs��ϸ!�Ƿ��ϴ� = 0) And rs��ϸ!ʵ�ս�� <> 0 And rs��ϸ!���� * rs��ϸ!���� <> 0 Then
            mblnReturn = fl_wdetails(gstrҽ����������, gstrҽԺ����, str������, str������, iLoop, _
                rsTemp!��� & "_" & rsTemp!ID, rsTemp!����, " ", ToVarchar(strTemp, 10), Nvl(rsTemp!��������, " "), Nvl(rsTemp!���㵥λ, " "), Format(rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), "#.####"), _
                rs��ϸ!���� * rs��ϸ!����, Format(rs��ϸ!�Ǽ�ʱ��, "yyyy-MM-dd"), Nvl(rs��ϸ!����Ա����, UserInfo.����), _
                IIf(rsTemp!��� = "5" Or rsTemp!��� = "6" Or rsTemp!��� = "7", "1", IIf(rsTemp!��� = "J", "3", "2")), gstrOutPara)
        End If
        TrimType
        If mblnReturn = False Then
            MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
            ������ϸ����_���� = False
            Exit Function
        End If
        
        gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rs��ϸ!ID & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        rs��ϸ.MoveNext
        iLoop = iLoop + 1
    Loop
    ������ϸ����_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curƱ���ܽ�� As Currency, lngErr As Long
    Dim datCurr As Date
    
    If gstrҽ���������� = "" Then
        Err.Raise 9000, gstrSysName, "׼����ȡҽ���������룬�����ϵͳ�����˿�"
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstrҽ���������� = gstrOutPara.out1
        gstrҽԺ���� = gstrOutPara.out2
    End If
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select ����ID,���ʽ�� From ������ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    lngErr = 1
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "Select * from �����ʻ� where ����ID=[2]"
    lngErr = 2
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    str������ = Nvl(rsTemp!˳���, "0")
'    gstrҽ���������� = rsTemp!����
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[3]"
    lngErr = 3
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]"
    lngErr = 4
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    If IsNull(rsTemp!��ע) Then
        Err.Raise 9000, gstrSysName, "�õ��ݵľ����Ŷ�ʧ���������ϡ�"
        Exit Function
    End If
    str������ = rsTemp!��ע
    '���ýӿ�������
    initType
    mblnReturn = fl_canrollback(gstrҽ����������, gstrҽԺ����, str������, gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, "�ж��Ƿ���Գ���ʱ��ҽ���˷���������Ϣ���˷Ѳ��ܼ�����" & Chr(13) & Chr(10) & gstrOutPara.errtext
        Exit Function
    End If
    initType
    mblnReturn = fl_rollbackcalc(gstrҽ����������, gstrҽԺ����, str������, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext
        Exit Function
    End If
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�): lngErr = 5
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")": lngErr = 6
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        cur�����ʻ� * -1 & ",NULL,NULL,NULL,'" & str������ & "')": lngErr = 7
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    ����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox "��������[����������]ģ�飬��" & lngErr & "�У�������Ϣ��" & Chr(13) & Chr(10) & Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str���� As String, str���ֱ��� As String
    Dim rsTmp As New ADODB.Recordset, str������ As String, datCurr As Date
    Dim lng����ID As Long
    Dim strת�� As String   '0-��;1-��
    
    '������˵������Ϣ
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.��Ժ����,A.��Ժ��ʽ,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    If rsTmp.RecordCount <> 0 Then
        datCurr = rsTmp!��Ժ����
    End If
    strת�� = IIf(Nvl(rsTmp!��Ժ��ʽ, "��ͨ") = "ת��", "1", "0")
    Call DebugTool("��Ժ��ʽ[" & Nvl(rsTmp!��Ժ��ʽ, "��ͨ") & "]")
    
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID)   '��Ժ���
    If rsTmp.BOF Then ��Ժ�Ǽ�_���� = False: Exit Function
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_����
    
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If rsTemp.State = 1 Then
        lng����ID = rsTemp("ID")
        str���� = rsTemp!����
        str���ֱ��� = rsTemp!ID
    Else
        ��Ժ�Ǽ�_���� = False
        Exit Function
    End If
    
    If gstrҽ���������� = "" Then
        MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstrҽ���������� = gstrOutPara.out1
        gstrҽԺ���� = gstrOutPara.out2
    End If

    initType
    mblnReturn = fl_reg(gstrҽ����������, gstrҽԺ����, 1, UserInfo.����, Format(datCurr, "yyyy-MM-dd"), "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        ��Ժ�Ǽ�_���� = False
        Exit Function
    End If
    str������ = gstrOutPara.out1
    
    initType
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,ҽԺ��������,ҽԺ��������,��������,ԭ��
'         �����־, ҽ������,�ز���־
    '������Ժ����
    Call DebugTool("Request������Σ�" & gstrҽ���������� & "," & gstrҽԺ���� & "," & str������ & "," & str���ֱ��� & "," & str���� & "," & Format(datCurr, "yyyy-MM-dd") & "," & _
            strInNote & "," & strת�� & "," & UserInfo.���� & "," & "0")
    mblnReturn = fl_request(gstrҽ����������, gstrҽԺ����, str������, str���ֱ���, str����, Format(datCurr, "yyyy-MM-dd"), _
            strInNote, strת��, UserInfo.����, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        ��Ժ�Ǽ�_���� = False
        Exit Function
    End If
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'˳���'," & str������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'����ID'," & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
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

Public Function סԺ�������_����(lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��

    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String, str������ As String, lng����ID As Long
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curƱ���ܽ�� As Currency
    Dim datCurr As Date, cur�����ʻ� As Currency

    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate

    gstrSQL = "Select ����ID,���ʽ�� From סԺ���ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)

    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop

    gstrSQL = "Select * from �����ʻ� where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    str������ = Nvl(rsTemp!˳���, "0")

    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B" & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng����ID = rsTemp("ID")

    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, lng����ID)

    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If

    If IsNull(rsTemp!��ע) Then
        Err.Raise 9000, gstrSysName, "�õ��ݵľ����Ŷ�ʧ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If

    str������ = rsTemp!��ע
    cur�����ʻ� = rsTemp!�����ʻ�֧��

    '���ýӿ�������
    initType
    mblnReturn = fl_canrollback(gstrҽ����������, gstrҽԺ����, str������, gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If

    initType
    mblnReturn = fl_rollbackcalc(gstrҽ����������, gstrҽԺ����, str������, "0", gstrOutPara)

    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)

    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_���� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        cur�����ʻ� * -1 & ",NULL,NULL,NULL,'" & str������ & "')"
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
    Dim str����Ա As String, datCurr As Date, str������ As String, cur�𸶶�֧�� As Currency
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur�����ʻ� As Currency, cur���� As Currency, cur����ͳ���޶� As Currency
    Dim cur���ͳ���޶� As Currency, cur�����Ը� As Currency, cur��� As Currency
    Dim cur�������� As Currency, curͳ��֧�� As Currency, cur����ҽ�� As Currency, curȫ�Ը� As Currency, cur���Ը� As Currency, cur����Ա���� As Currency
    Dim lng��ҳID As Long, cur�ʻ�֧�� As Currency
    Dim str��Ժ���� As String, str����Ա��� As String
    
    On Error GoTo errHandle
    '��Ҫ���ϴ�������ϸ
'    ������ϸ����_���� lng����ID
    
    If gstrҽ���������� = "" Then
        Err.Raise 9000, gstrSysName, "׼����ȡҽ���������룬�����ϵͳ�����˿���"
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstrҽ���������� = gstrOutPara.out1
        gstrҽԺ���� = gstrOutPara.out2
    End If
    
    '������ʻ�֧�����
    gstrSQL = "Select Nvl(��Ԥ��,0) as ��� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ��¼����=2 And ����ID=" & lng����ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ģ��ҽ��", lng����ID)
    If rsTemp.RecordCount > 0 Then
        cur�ʻ�֧�� = rsTemp("���")
    End If
    
    gstrSQL = "Select * From סԺ���ü�¼ Where ����ID=[1] And Rownum<2"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    lng����ID = rs��ϸ("����ID")
    lng��ҳID = rs��ϸ("��ҳID")
    str����Ա = UserInfo.����
    
    'ȡ������ҳ�еĳ�Ժ����
    gstrSQL = " Select A.��Ժ���� From ������ҳ A,������Ϣ B" & _
              " Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.RecordCount <> 0 Then
        str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    Else
        str��Ժ���� = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End If
    
    gstrSQL = "Select nvl(˳���,0) as ˳��� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    str������ = rsTemp!˳���
    'ҽ����������, ҽԺ���, ҽ�������ţ� ��Ժ���ڣ�����Ա����ʾ��־
    datCurr = zlDatabase.Currentdate
    initType
    mblnReturn = fl_calc(gstrҽ����������, gstrҽԺ����, str������, str��Ժ����, str����Ա, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext
        סԺ����_���� = False
        Exit Function
    End If
'��ӳ��ڲ���:1���úϼ�,2���ⲡ�ַ���,3���α����ʻ�֧��,4���������ʻ�֧��,5�ۼƷֶ��Ը�,6ͳ���֧��,7�𸶶�֧��,
'             8��λ֧��,9�Էѷ���,10�ؼ����Ը�,11�������Ը�,12�ؼ����,13���η���,14����ҽ�Ʊ���֧��,15����ͳ������ۼ�,
'             16����ҽ�Ƽ����ۼ�,17����ͳ������ۼ�,18δ��������,19ҽ��֧��,20�����ֽ�֧��,21�����ʻ����

    '��ȡ�����ʻ�֧���͸����ֽ�֧��
    '�жϵ�23���������β�Ϊ��ʱ,�����ڸ���ҽ�� 20070929������
    str����Ա��� = gstrOutPara.out22
    If gstrOutPara.out23 <> "" Then
        cur�����ʻ� = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
        cur��� = CCur(gstrOutPara.out21)
        curȫ�Ը� = CCur(gstrOutPara.out20) - cur�����ʻ�
        cur�������� = CCur(gstrOutPara.out1)
        cur���Ը� = CCur(gstrOutPara.out10) + CCur(gstrOutPara.out11)
        cur�𸶶�֧�� = CCur(gstrOutPara.out7)
        curͳ��֧�� = CCur(gstrOutPara.out6)
        cur����ҽ�� = CCur(gstrOutPara.out14)
        cur����Ա���� = CCur(gstrOutPara.out23)
        
        'У��
        Dim str���㷽ʽ As String
        If cur�����ʻ� <> cur�ʻ�֧�� Then
            str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
            str���㷽ʽ = str���㷽ʽ & "||����ҽ�Ʊ���|" & curͳ��֧��
            str���㷽ʽ = str���㷽ʽ & "||����ҽ�Ʊ���|" & cur����ҽ��
            str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
            
            '�Խ�������㲻һ��,��Ҫ����
            str���㷽ʽ = Mid(str���㷽ʽ, 3)
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�Խ�������㲻һ��,����Ԥ����¼")
        End If
        
        '�ʻ������Ϣ
        Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
        curͳ�ﱨ���ۼ� = curͳ�ﱨ���ۼ� + curͳ��֧�� + cur����Ա����
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
                "," & TYPE_���� & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
                "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
                curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & _
                cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        '���ս����¼
        '����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN, _
         �ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,�������ý��_IN,ȫ�Ը����_IN, _
         �����Ը����_IN,����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN, _
         ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & _
                lng����ID & "," & Year(datCurr) & "," & _
                cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
                curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�𸶶�֧�� & ",NULL,NULL," & _
                cur�������� & "," & curȫ�Ը� & "," & cur���Ը� & "," & cur����ҽ�� & "," & curͳ��֧�� & "," & cur����Ա���� & ",NULL," & _
                cur�����ʻ� & ",'" & str����Ա��� & "'," & lng��ҳID & ",NULL,'" & str������ & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        '---------------------------------------------------------------------------------------------
    Else
        cur�����ʻ� = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
        cur��� = CCur(gstrOutPara.out21)
        curȫ�Ը� = CCur(gstrOutPara.out20) - cur�����ʻ�
        cur�������� = CCur(gstrOutPara.out1)
        cur���Ը� = CCur(gstrOutPara.out10) + CCur(gstrOutPara.out11)
        cur�𸶶�֧�� = CCur(gstrOutPara.out7)
        curͳ��֧�� = CCur(gstrOutPara.out6)
        cur����ҽ�� = CCur(gstrOutPara.out14)
        'cur����Ա���� = CCur(gstrOutPara.out23)
        
        'У��
       ' Dim str���㷽ʽ As String
        If cur�����ʻ� <> cur�ʻ�֧�� Then
            str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
            str���㷽ʽ = str���㷽ʽ & "||����ҽ�Ʊ���|" & curͳ��֧��
            str���㷽ʽ = str���㷽ʽ & "||����ҽ�Ʊ���|" & cur����ҽ��
            'str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
            
            '�Խ�������㲻һ��,��Ҫ����
            str���㷽ʽ = Mid(str���㷽ʽ, 3)
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�Խ�������㲻һ��,����Ԥ����¼")
        End If
        
        '�ʻ������Ϣ
        Call Get�ʻ���Ϣ(TYPE_����, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
        curͳ�ﱨ���ۼ� = curͳ�ﱨ���ۼ� + curͳ��֧��
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
                "," & TYPE_���� & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
                "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
                curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & _
                cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        '���ս����¼
        '����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN, _
         �ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,�������ý��_IN,ȫ�Ը����_IN, _
         �����Ը����_IN,����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN, _
         ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_���� & "," & _
                lng����ID & "," & Year(datCurr) & "," & _
                cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
                curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�𸶶�֧�� & ",NULL,NULL," & _
                cur�������� & "," & curȫ�Ը� & "," & cur���Ը� & "," & cur����ҽ�� & "," & curͳ��֧�� & ",NULL,NULL," & _
                cur�����ʻ� & ",NULL," & lng��ҳID & ",NULL,'" & str������ & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    End If
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
'20070926����������:cur����Ա����
    
    Dim cur�����ʻ�֧�� As Currency, cur�����ֽ�֧�� As Currency, cur�ʻ���� As Currency
    Dim curͳ��֧�� As Currency, curҽ��֧�� As Currency, cur����ҽ�� As Currency, cur����Ա���� As Currency
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str����Ա As String, datCurr As Date, str������ As String
    Dim curCount As Currency
    
    On Error GoTo errHandle
    '��Ҫ���ϴ�������ϸ
'    ������ϸ����_���� 0, rs������ϸ
'
    If gstrҽ���������� = "" Then
        MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstrҽ���������� = gstrOutPara.out1
        gstrҽԺ���� = gstrOutPara.out2
    End If
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
    If curCount = 0 Then
        MsgBox "���˷���Ϊ��,���ܽ��н���", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    str����Ա = UserInfo.����
    
    If ���ʴ���_����("", 0, "", lng����ID) = False Then
        Exit Function
    End If
    
    gstrSQL = "Select Nvl(�ʻ����,0) AS �ʻ����,nvl(˳���,0) as ˳���,���� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
    str������ = rsTemp!˳���
    cur�ʻ���� = rsTemp!�ʻ����
'    gstrҽ���������� = rsTemp!����
    'ҽ����������, ҽԺ���, ҽ�������ţ� ��Ժ���ڣ�����Ա����ʾ��־
    datCurr = zlDatabase.Currentdate
    initType
    mblnReturn = fl_pcalc(gstrҽ����������, gstrҽԺ����, str������, Format(datCurr, "yyyy-MM-dd"), str����Ա, "1", "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        סԺ�������_���� = ""
        Exit Function
    End If
'��ӳ��ڲ���:1���úϼ�,2���ⲡ�ַ���,3���α����ʻ�֧��,4���������ʻ�֧��,5�ۼƷֶ��Ը�,6ͳ���֧��,7�𸶶�֧��,
'             8��λ֧��,9�Էѷ���,10�ؼ����Ը�,11�������Ը�,12�ؼ����,13���η���,14����ҽ�Ʊ���֧��,15����ͳ������ۼ�,
'             16����ҽ�Ƽ����ۼ�,17����ͳ������ۼ�,18δ��������,19ҽ��֧��,20�����ֽ�֧��,21�����ʻ����
    
    '��ȡ�����ʻ�֧���͸����ֽ�֧��
    '�жϵ�22���������β�Ϊ��ʱ,�����ڸ���ҽ�� 20070929������
    If gstrOutPara.out22 <> "" Then
            cur�����ʻ�֧�� = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
            cur�����ֽ�֧�� = CCur(gstrOutPara.out20)
            curͳ��֧�� = CCur(gstrOutPara.out6)
            curҽ��֧�� = CCur(gstrOutPara.out19)
            cur����ҽ�� = CCur(gstrOutPara.out14)
            cur����Ա���� = CCur(gstrOutPara.out22)
            If curCount <> CCur(gstrOutPara.out1) Then
                MsgBox "��ע�⣺ҽ�����ؽ������뵱ǰ���ݽ���" & vbCrLf & _
                               "ҽԺ�ܶ" & curCount & "    ҽ�����أ�" & CCur(gstrOutPara.out1), vbInformation, gstrSysName
            End If
            סԺ�������_���� = "�����ʻ�;" & cur�����ʻ�֧�� & ";0" '�������޸ĸ����ʻ�
        '    If cur�����ֽ�֧�� <> 0 Then
        '        סԺ�������_���� = סԺ�������_���� & "|�ֽ�;" & cur�����ֽ�֧�� & ";0" '�������޸��ֽ�֧��
        '    End If
            If curͳ��֧�� <> 0 Then
                סԺ�������_���� = סԺ�������_���� & "|����ҽ�Ʊ���;" & curͳ��֧�� & ";0" '�������޸�ͳ��֧��
            End If
            If cur����ҽ�� <> 0 Then
                סԺ�������_���� = סԺ�������_���� & "|����ҽ�Ʊ���;" & cur����ҽ�� & ";0"
            End If
            
            If cur����Ա���� <> 0 Then
                סԺ�������_���� = סԺ�������_���� & "|����Ա����;" & cur����Ա���� & ";0"
            End If

            If cur�ʻ���� < cur�����ʻ�֧�� Then
                gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'�ʻ����'," & cur�����ʻ�֧�� & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ����")
            End If
    Else
    
            cur�����ʻ�֧�� = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
            cur�����ֽ�֧�� = CCur(gstrOutPara.out20)
            curͳ��֧�� = CCur(gstrOutPara.out6)
            curҽ��֧�� = CCur(gstrOutPara.out19)
            cur����ҽ�� = CCur(gstrOutPara.out14)
'            cur����Ա���� = CCur(gstrOutPara.out23)
            If curCount <> CCur(gstrOutPara.out1) Then
                MsgBox "��ע�⣺ҽ�����ؽ������뵱ǰ���ݽ���" & vbCrLf & _
                               "ҽԺ�ܶ" & curCount & "    ҽ�����أ�" & CCur(gstrOutPara.out1), vbInformation, gstrSysName
            End If
            סԺ�������_���� = "�����ʻ�;" & cur�����ʻ�֧�� & ";0" '�������޸ĸ����ʻ�
        '    If cur�����ֽ�֧�� <> 0 Then
        '        סԺ�������_���� = סԺ�������_���� & "|�ֽ�;" & cur�����ֽ�֧�� & ";0" '�������޸��ֽ�֧��
        '    End If
            If curͳ��֧�� <> 0 Then
                סԺ�������_���� = סԺ�������_���� & "|����ҽ�Ʊ���;" & curͳ��֧�� & ";0" '�������޸�ͳ��֧��
            End If
            If cur����ҽ�� <> 0 Then
                סԺ�������_���� = סԺ�������_���� & "|����ҽ�Ʊ���;" & cur����ҽ�� & ";0"
            End If
            
'            If cur����Ա���� <> 0 Then
'                סԺ�������_���� = סԺ̓�����_���� & "|����Ա����:" & cur����Ա���� & ":0"
'            End If
             
            If cur�ʻ���� < cur�����ʻ�֧�� Then
                gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_���� & ",'�ʻ����'," & cur�����ʻ�֧�� & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ʻ����")
            End If
  End If
            Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    סԺ�������_���� = ""
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    Dim str������ As String, rsTemp As New ADODB.Recordset
    Dim bln����ó�Ժ As Boolean
    
    If gstrҽ���������� = "" Then
        MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstrҽ���������� = gstrOutPara.out1
        gstrҽԺ���� = gstrOutPara.out2
    End If
    
    On Error GoTo errHandle
    '���ô�סԺ�Ƿ�û�з��÷���
    gstrSQL = "Select nvl(sum(ʵ�ս��),0) as ���  from סԺ���ü�¼ where nvl(���ӱ�־,0)<>9 and ����ID=[1] and ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���˳�Ժ", lng����ID, lng��ҳID)
    If rsTemp.EOF = True Then
        bln����ó�Ժ = True
    Else
        bln����ó�Ժ = (rsTemp("���") = 0)
    End If
    
    If bln����ó�Ժ = True Then
        gstrSQL = "Select nvl(˳���,0) as ˳��� From �����ʻ� Where ����ID=[1] And ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_����)
        str������ = rsTemp!˳���
        initType
        mblnReturn = fl_dall(gstrҽ����������, gstrҽԺ����, str������, gstrOutPara)
        If mblnReturn = False Then
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
    If ErrCenter() = 1 Then Resume
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
    Dim lng��ҳID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim RSPATIENT As New ADODB.Recordset
    
    
    If str���ݺ� <> "" Then
        '���ܴ���һ�ŵ������ж�����˵����
        gstrSQL = "Select Distinct ����ID From סԺ���ü�¼ Where ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ��¼����=[1] and NO=[2]"
        Set RSPATIENT = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", int����, str���ݺ�)
        With RSPATIENT
            Do While Not .EOF
                '�ж��Ƿ���ҽ������
                gstrSQL = "Select Count(*) Records From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.����ID=[1] And A.��ҳID=B.סԺ���� And A.����=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ҽ������", CLng(!����ID), TYPE_����)
                If (rsTemp!Records = 1) Then
                    '��ȡָ�����ˣ�ָ��������ϸ
                    gstrSQL = "Select * From סԺ���ü�¼ Where ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ��¼����=[1] and NO=[2] And ����ID=[3] order by ��ҳID,���"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", int����, str���ݺ�, CLng(!����ID))
                    ���ʴ���_���� = ������ϸ����_����(0, rsTemp)
                    If Not ���ʴ���_���� Then Exit Function
                Else
                    ���ʴ���_���� = True
                End If
                .MoveNext
            Loop
        End With
    Else
        If lng����ID <> 0 Then
            gstrSQL = "Select Max(��ҳID) From ������ҳ Where ����id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
            lng��ҳID = rsTemp(0)
        End If
        
        gstrSQL = "Select * From סԺ���ü�¼ Where ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And Nvl(ʵ�ս��,0)<>0 And nvl(���ӱ�־,0)<>9 and ����id=[1] And ��ҳid=[2] order by ��ҳID,���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", lng����ID, lng��ҳID)
        ���ʴ���_���� = ������ϸ����_����(0, rsTemp)
    End If
End Function

Public Function LeftStr(sString As String, iLen As Long) As String
    Dim sStr As String, sTmp As String
    sStr = sString
    If LenB(StrConv(sStr, vbFromUnicode)) > iLen Then
        sTmp = StrConv(sStr, vbFromUnicode)
        sStr = StrConv(LeftB(sTmp, iLen), vbUnicode)
        If asc(Right(sStr, 1)) = 0 Then sStr = StrConv(LeftB(sTmp, iLen - 1), vbUnicode)
    End If
    LeftStr = sStr
End Function

