Attribute VB_Name = "mdl��Ԫ"
Option Explicit

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
Public Declare Function gy_readzfbl Lib "cxybclient.dll" Alias "readzfbl" (ByVal strybjgbm As String, ByVal stryybm As String, _
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
Public Declare Function gy_readicxx Lib "cxybclient.dll" Alias "readicxx" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strxxbz As String, strOut As str_Out) As Boolean
    
'=========================================================================================================
'����˵��:ҽ������סԺ�����Ҫ�������Ե��ô˺���,��������Ϣ����ҽ��,��ҽ��������
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,ҽԺ��������,ҽԺ��������,��������,ԭ��
'         �����־, ҽ������,�ز���־
'��ӳ��ڲ���:��
'��������˵��:
'    �����־: 0---��,1---��
'    �ز���־: 0---��,1---��
'=========================================================================================================
Public Declare Function gy_request Lib "cxybclient.dll" Alias "request" (ByVal strybjgbm As String, _
    ByVal stryybm As String, ByVal strybjzbh As String, ByVal stryyjbbm As String, _
    ByVal stryyjbmc As String, ByVal strsjrq As String, ByVal strsjyy As String, ByVal strjzbz As String, _
    ByVal strysxm As String, ByVal strtbbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:ˢ��IC��,��ʾIC����Ϣ,�����סԺ�Ǽ�,�����ز�����ԺΨһ��ʶ���(ҽ��������)��
'��ڲ���:ҽ����������,ҽԺ���,��־,����Ա����,��������,�Ƿ���Ѫ
'��ӳ��ڲ���:ҽ��������,ҽ����������,�����ʺ�,ic����,���֤����,����,�Ա�,��λ����,��λ����,��������,
'             ��Ա���,IC�����,�𸶱�׼,����ҽ������޶�,����ҽ���ۼ�֧��,��ҽ���ۼ�֧��
'��������˵��:
'        ��־:0---����,1---סԺ
'  �Ƿ���Ѫ:�ӿں���δ˵��
'=========================================================================================================
Public Declare Function gy_reg Lib "cxybclient.dll" Alias "reg" (ByVal strybjgbm As String, ByVal stryybm As String, _
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
Public Declare Function gy_wrecipe Lib "cxybclient.dll" Alias "wrecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
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
Public Declare Function gy_urecipe Lib "cxybclient.dll" Alias "urecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal stryyjbbm As String, ByVal stryyjbmc As String, _
    ByVal strtbbz As String, ByVal strysxm As String, ByVal strlryxm As String, ByVal strksbh As String, _
    ByVal strksmc As String, ByVal strcfrq As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:ɾ��������Ϣ
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function gy_drecipe Lib "cxybclient.dll" Alias "drecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
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
Public Declare Function gy_wdetails Lib "cxybclient.dll" Alias "wdetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
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
Public Declare Function gy_udetails Lib "cxybclient.dll" Alias "udetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
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
Public Declare Function gy_ddetails Lib "cxybclient.dll" Alias "ddetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
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
Public Declare Function gy_pcalc Lib "cxybclient.dll" Alias "pcalc" (ByVal strybjgbm As String, ByVal stryybm As String, _
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
Public Declare Function gy_calc Lib "cxybclient.dll" Alias "calc" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcyrq As String, ByVal strczy As String, _
    ByVal strxsbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:סԺ�����,ȡ����ʽ����,���ص�����ǰ״̬;���������,���ɺ��ֵ���,��������¼
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,��ʾ��־
'��ӳ��ڲ���:��
'��������˵��:
'    ��ʾ��־:0---����ʾ,1---��ʾ
'=========================================================================================================
Public Declare Function gy_rollbackcalc Lib "cxybclient.dll" Alias "rollbackcalc" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strxsbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:ɾ����ҽ�������ŵ�������Ϣ,������Ժ�Ǽ�,����,������ϸ�ȡ������������ʽ����,����ʹ�øú���
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function gy_dall Lib "cxybclient.dll" Alias "dall" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:���ô����Ƿ���ɾ�����޸�
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function gy_canupdaterecipe Lib "cxybclient.dll" Alias "canupdaterecipe" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:���ô�����ϸ�Ƿ���ɾ�����޸�
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,�������,��ϸ���
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function gy_canupdatedetails Lib "cxybclient.dll" Alias "canupdatedetails" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, ByVal strcfbh As String, ByVal strmxbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:����Ƿ��ܹ��ع�,סԺ�������ʹ�ô˺����ж�
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function gy_canrollback Lib "cxybclient.dll" Alias "canrollback" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strybjzbh As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:����Ƿ������ҽ���Ը�����,�����Ƿ��и���ʱ��������һ�μ���ʱ���ҩƷ,������Ŀ,����,��λ
'��ڲ���:ҽ����������,ҽԺ���,���ͱ�־
'��ӳ��ڲ���:��
'��������˵��:
'    ���ͱ�־:1---ҩƷ,2---������Ŀ,3---����,4---��λ
'=========================================================================================================
Public Declare Function gy_havenewzfbl Lib "cxybclient.dll" Alias "havenewzfbl" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strlxbz As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:����ҽ��������ʱ��
'��ڲ���:��
'��ӳ��ڲ���:ҽ��������ʱ��
'��������˵��:
'=========================================================================================================
Public Declare Function gy_getsystime Lib "cxybclient.dll" Alias "getsystime" (strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:����ҽ����������(��ҽ������ϵͳ������IC��)
'��ڲ���:��
'��ӳ��ڲ���:ҽ����������,ҽԺ����
'��������˵��:
'=========================================================================================================
Public Declare Function gy_getybjgbm Lib "cxybclient.dll" Alias "getybjgbm" (strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:������Ժҽ�Ƶ���(��ҽ������ϵͳ������IC��)
'��ڲ���:ҽ����������,ҽԺ���,�����ʺ�
'��ӳ��ڲ���:��Ժҽ�Ƶ���,ҽԺ����
'��������˵��:
'=========================================================================================================
Public Declare Function gy_getlastzyxx Lib "cxybclient.dll" Alias "getlastzyxx" (ByVal strybjgbm As String, ByVal stryybm As String, _
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
Public Declare Function gy_wyyglxx Lib "cxybclient.dll" Alias "wyyglxx" (ByVal strybjgbm As String, ByVal stryybm As String, _
    ByVal strlb As String, ByVal stryyxxbm As String, _
    ByVal stryyxxmc As String, ByVal strqt As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:�޸��û���IC������
'��ڲ���:��
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function gy_changePassword Lib "cxybclient.dll" Alias "changepassword" (ByVal strybjgbm As String, _
    ByVal stryybm As String, strOut As str_Out) As Boolean

'=========================================================================================================
'����˵��:У��ϵͳ��
'��ڲ���:��
'��ӳ��ڲ���:��
'��������˵��:
'=========================================================================================================
Public Declare Function gy_checkxtk Lib "cxybclient.dll" Alias "checkxtk" (strOut As str_Out) As Boolean

Private mblnReturn As Boolean

Public Function ҽ����ʼ��_��Ԫ() As Boolean
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", TYPE_��Ԫ)
    
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
    ҽ����ʼ��_��Ԫ = True
End Function

Public Function ��ݱ�ʶ_��Ԫ(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim frmIDentified As New frmIdentify��Ԫ
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
    
    strPatiInfo = frmIDentified.GetPatient(bytType, strTemp)
    
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
        lng����ID = BuildPatiInfo(bytType, strPatiInfo, lng����ID, TYPE_��Ԫ)
        
        '���ظ�ʽ:�м���벡��ID
        strPatiInfo = frmIDentified.mstrPatient & lng����ID & ";" & frmIDentified.mstrOther
        str������ = frmIDentified.mstr������
        'д�������
        If bytType = 0 Or bytType = 5 Then
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ԫ & ",'˳���','''" & str������ & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_��Ԫ")
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ԫ & ",'����֤��','''" & CLng(strTemp) + 1 & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_��Ԫ")
        End If
        Unload frmIDentified
    Else
        ��ݱ�ʶ_��Ԫ = ""
        MsgBox "ҽ��������Ϣ��ȡʧ��", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    ��ݱ�ʶ_��Ԫ = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_��Ԫ = ""
End Function

Public Function �������_��Ԫ(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, TYPE_��Ԫ)
    
    If rsTemp.EOF Then
        �������_��Ԫ = 0
    Else
        �������_��Ԫ = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
    End If
End Function

Public Function �����������_��Ԫ(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
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
        Err.Raise 9000, gstrSysName, "û�з��ã����ܽ���Ԥ���㡣", vbInformation, gstrSysName
        �����������_��Ԫ = False
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID"): lngErr = 1
    cur�Ը� = 0: cur���� = 0: lngErr = 2
    gstrSQL = "Select * from �����ʻ� where ����id=[1]": lngErr = 3
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��Ԥ����", lng����ID): lngErr = 4
    cur��� = rsTemp!�ʻ����: lngErr = 5
    strTemp = rsTemp!��ְ: lngErr = 4
    str������ϸ = ""
    While Not rs��ϸ.EOF
        gstrSQL = "select * from �շ�ϸĿ where id=[1]": lngErr = 6
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��Ԥ����", CLng(rs��ϸ!�շ�ϸĿID)): lngErr = 7
        
        '��ȡ�շ�ϸĿ���Ը�����
        initType
        mblnReturn = gy_readzfbl(gstrҽ����������, gstrҽԺ����, rsTemp!��� & "_" & rsTemp!ID, _
            IIf(rsTemp!��� = "5" Or rsTemp!��� = "6" Or rsTemp!��� = "7", "1", IIf(rsTemp!��� = "J", "3", "2")), _
            strTemp, gstrOutPara): lngErr = 8
        TrimType
        
        If mblnReturn = False Then
            Err.Raise 9000, gstrSysName, "�ڻ�ȡ��Ŀ[" & rsTemp!ID & "]���Ը�����ʱ��ҽ���ӿڷ������´���" & Chr(13) & Chr(10) & gstrOutPara.errtext
            �����������_��Ԫ = False
            Exit Function
        End If
        Select Case gstrOutPara.out2
            Case "1"            '����Ϊ�Ը�����
                curTemp = rs��ϸ!ʵ�ս�� * (1 - CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0))): lngErr = 9
            Case "2"            '����Ϊ�����޶�
                curTemp = IIf(rs��ϸ!ʵ�ս�� > CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0)) * rs��ϸ!����, CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0)) * rs��ϸ!����, rs��ϸ!ʵ�ս��): lngErr = 10
            Case "3"            '���Ը��������㱨���������ڿɱ������ȡ�ɱ�����
                str�Ը����� = Left(gstrOutPara.out1, InStr(gstrOutPara.out1, "|") - 1): lngErr = 11
                str�ɱ����� = Mid(gstrOutPara.out1, InStr(gstrOutPara.out1, "|") + 1): lngErr = 12
                str�Ը����� = IIf(IsNumeric(str�Ը�����), str�Ը�����, 0): lngErr = 13
                str�ɱ����� = IIf(IsNumeric(str�ɱ�����), str�ɱ�����, 0): lngErr = 14
                curTemp = rs��ϸ!ʵ�ս�� * (1 - CCur(str�Ը�����)): lngErr = 15
                curTemp = IIf(curTemp > CCur(str�ɱ�����) * rs��ϸ!����, CCur(str�ɱ�����) * rs��ϸ!����, curTemp): lngErr = 16
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
    If cur���� > cur��� - 1 Then
        curTemp = cur���� - (cur��� - 1): lngErr = 20
        cur���� = cur��� - 1: lngErr = 21
        cur�Ը� = cur�Ը� + curTemp: lngErr = 22
    End If
    
'    MsgBox str������ϸ, vbInformation, "������ϸ"
    
    str���㷽ʽ = "�����ʻ�;" & cur���� & ";0": lngErr = 23
    �����������_��Ԫ = True
    Exit Function
errHandle:
    ErrMsgBox "���������[����Ԥ����]ģ�飬��" & lngErr & "�У�������Ϣ��" & Chr(13) & Chr(10) & Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function �������_��Ԫ(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
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
    Dim cur�������� As Currency, cur���Ը� As Currency, lng����ID As Long
    
    If gstrҽ���������� = "" Then
        Err.Raise 9000, gstrSysName, "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = gy_getybjgbm(gstrOutPara)
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
        MsgBox "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    str����Ա = ToVarchar(IIf(IsNull(rs��ϸ("����Ա����")), UserInfo.����, rs��ϸ("����Ա����")), 20)
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_��Ԫ

    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If rsTemp.State = 1 Then
        lng����ID = rsTemp("ID")
    Else
        �������_��Ԫ = False
        Exit Function
    End If
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ԫ & ",'����ID'," & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_��Ԫ")

    '��Ҫ���ϴ�������ϸ
    ������ϸ����_��Ԫ lng����ID
    
    
    gstrSQL = "Select nvl(˳���,0) as ˳���,����id From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_��Ԫ
    Set rsTemp = gcnOracle.Execute(gstrSQL)
    lng����ID = rsTemp!����ID
    str������ = rsTemp!˳���
    
    'ҽ����������, ҽԺ���, ҽ�������ţ� ��Ժ���ڣ�����Ա����ʾ��־
    datCurr = zlDatabase.Currentdate
    initType
'    mblnReturn = gy_pcalc(gstrҽ����������, gstrҽԺ����, str������, Format(datCurr, "yyyy-MM-dd"), str����Ա, "1", "0", gstrOutPara)
    mblnReturn = gy_calc(gstrҽ����������, gstrҽԺ����, str������, Format(datCurr, "yyyy-MM-dd"), str����Ա, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext, vbInformation, gstrSysName
        �������_��Ԫ = False
        Exit Function
    End If
'��ӳ��ڲ���:1���úϼ�,2���ⲡ�ַ���,3���α����ʻ�֧��,4���������ʻ�֧��,5�ۼƷֶ��Ը�,6ͳ���֧��,7�𸶶�֧��,
'             8��λ֧��,9�Էѷ���,10�ؼ����Ը�,11�������Ը�,12�ؼ����,13���η���,14����ҽ�Ʊ���֧��,15����ͳ������ۼ�,
'             16����ҽ�Ƽ����ۼ�,17����ͳ������ۼ�,18δ��������,19ҽ��֧��,20�����ֽ�֧��,21�����ʻ����
    
    '��ȡ�����ʻ�֧���͸����ֽ�֧��
    cur�����ʻ� = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
    cur��� = CCur(gstrOutPara.out21)
    curȫ�Ը� = CCur(gstrOutPara.out20) + CCur(cur�����ʻ�)
    cur�������� = CCur(gstrOutPara.out1)
    cur���Ը� = CCur(gstrOutPara.out10) + CCur(gstrOutPara.out11)
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��Ԫ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & Get����ID(CStr(strҽ����), CStr(TYPE_��Ԫ)) & _
            "," & TYPE_��Ԫ & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & _
            cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_��Ԫ & "," & _
            Get����ID(CStr(strҽ����), CStr(TYPE_��Ԫ)) & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur�������� & "," & curȫ�Ը� & "," & cur���Ը� & ",NULL,NULL,NULL,NULL," & _
            cur�����ʻ� & ",NULL,NULL,NULL,'" & str������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    �������_��Ԫ = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ������ϸ����_��Ԫ(lng����ID As Long, Optional rs��ϸIN As ADODB.Recordset = Nothing, Optional ByVal bln�����ϴ� As Boolean = False) As Boolean
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
        mblnReturn = gy_getybjgbm(gstrOutPara)
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
    
    If bln�����ϴ� Then
        '���жϱ��ղ�����ʵʱ�ϴ��������Ϊ�٣�ֱ���˳�
        gstrSQL = "Select Nvl(����ֵ,1) AS ����ֵ From ���ղ��� Where ����=[1] And ������='ʵʱ�ϴ�'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�ʵʱ�ϴ�", TYPE_��Ԫ)
        If rsTemp.RecordCount <> 0 Then
            If rsTemp!����ֵ = 0 Then
                ������ϸ����_��Ԫ = True
                Exit Function
            End If
        End If
    End If
    
    If rs��ϸIN Is Nothing Then
        gstrSQL = "Select * From ������ü�¼ Where ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 and ����ID=[1]"
        Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    Else
        Set rs��ϸ = rs��ϸIN.Clone
    End If
    If rs��ϸ.EOF = True Then
'        MsgBox "û����Ҫ�ϴ����շѼ�¼", vbExclamation, gstrSysName
        ������ϸ����_��Ԫ = True
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    str����Ա = ToVarchar(UserInfo.����, 20)
    
'    gstrSQL = "select max(��ҳID) as ��ҳID from ������ҳ where ����ID =" & lng����ID
'    Call OpenRecordset(rsTemp, gstrsysname)
'    strBillNo = CStr(lng����ID) & "_" & CStr(rsTemp("��ҳID"))
    gstrSQL = "Select nvl(˳���,0) as ˳���,����ID,����,����֤�� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��Ԫ)
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
    mblnReturn = gy_wrecipe(gstrҽ����������, gstrҽԺ����, str������, str������, str���ֱ���, str��������, _
                         int�ز���־, Nvl(rs��ϸ!������, rs��ϸ!������), Nvl(rs��ϸ!����Ա����, UserInfo.����), str���ұ��, _
                         str��������, Format(rs��ϸ!����ʱ��, "yyyy-MM-dd"), gstrOutPara)
    TrimType
    If mblnReturn = False Then
        If InStr(gstrOutPara.errtext, "(YBYY.PRI_QTYL42_T)") > 0 Then
            ������ϸ����_��Ԫ = True
        Else
            MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
            ������ϸ����_��Ԫ = False
            Exit Function
        End If
    End If
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ԫ & ",'����֤��','" & CLng(str������) + 1 & "')"
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
        If Nvl(rs��ϸ!�Ƿ��ϴ�, 0) = 0 And Nvl(rs��ϸ!ʵ�ս��, 0) <> 0 Then
            mblnReturn = gy_wdetails(gstrҽ����������, gstrҽԺ����, str������, str������, iLoop, _
                rsTemp!��� & "_" & rsTemp!ID, rsTemp!����, " ", ToVarchar(strTemp, 10), Nvl(rsTemp!��������, " "), Nvl(rsTemp!���㵥λ, " "), rs��ϸ!ʵ�ս�� / (rs��ϸ!���� * rs��ϸ!����), _
                rs��ϸ!���� * rs��ϸ!����, Format(rs��ϸ!����ʱ��, "yyyy-MM-dd"), Nvl(rs��ϸ!����Ա����, UserInfo.����), _
                IIf(rsTemp!��� = "5" Or rsTemp!��� = "6" Or rsTemp!��� = "7", "1", IIf(rsTemp!��� = "J", "3", "2")), gstrOutPara)
'        Else
'            mblnReturn = gy_udetails(gstrҽ����������, gstrҽԺ����, str������, str������, rs��ϸ!���, _
'                rsTemp!��� & "_" & rsTemp!ID, rsTemp!����, " ", strTemp, NVL(rsTemp!��������, " "), NVL(rsTemp!���㵥λ, " "), rs��ϸ!��׼����, _
'                rs��ϸ!���� * rs��ϸ!����, Format(rs��ϸ!�Ǽ�ʱ��, "yyyy-MM-dd"), NVL(rs��ϸ!����Ա����, UserInfo.����), _
'                IIf(rsTemp!��� = "5" Or rsTemp!��� = "6" Or rsTemp!��� = "7", "1", IIf(rsTemp!��� = "J", "3", "2")), gstrOutPara)
        End If
        TrimType
        If mblnReturn = False Then
            MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
            ������ϸ����_��Ԫ = False
            Exit Function
        End If
        gstrSQL = "zl_���˼��ʼ�¼_�ϴ� ('" & rs��ϸ!ID & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        rs��ϸ.MoveNext
        iLoop = iLoop + 1
    Loop
    ������ϸ����_��Ԫ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����������_��Ԫ(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
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
        Err.Raise 9000, gstrSysName, "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = gy_getybjgbm(gstrOutPara)
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
    
    gstrSQL = "Select ����ID,���ʽ�� From ������ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]": lngErr = 1
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        
        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "Select * from �����ʻ� where ����ID=[1]": lngErr = 2
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    str������ = Nvl(rsTemp!˳���, "0")
'    gstrҽ���������� = rsTemp!����
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B" & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]": lngErr = 3
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]": lngErr = 4
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ԫ, lng����ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If IsNull(rsTemp!��ע) Then
        Err.Raise 9000, gstrSysName, "�õ��ݵľ����Ŷ�ʧ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    '���ýӿ�������
    str������ = rsTemp!��ע
    initType
    mblnReturn = gy_canrollback(gstrҽ����������, gstrҽԺ����, str������, gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, "�ж��Ƿ���Գ���ʱ��ҽ���˷���������Ϣ���˷Ѳ��ܼ�����" & Chr(13) & Chr(10) & gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If
    initType
    mblnReturn = gy_rollbackcalc(gstrҽ����������, gstrҽԺ����, str������, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��Ԫ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�): lngErr = 5
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_��Ԫ & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")": lngErr = 6
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_��Ԫ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        cur�����ʻ� * -1 & ",NULL,NULL,NULL,'" & str������ & "')": lngErr = 7
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    ����������_��Ԫ = True
    Exit Function
errHandle:
    ErrMsgBox "��������[����������]ģ�飬��" & lngErr & "�У�������Ϣ��" & Chr(13) & Chr(10) & Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ��Ժ�Ǽ�_��Ԫ(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str���� As String, str���ֱ��� As String
    Dim rsTmp As New ADODB.Recordset, str������ As String, datCurr As Date
    Dim lng����ID As Long
    
    '������˵������Ϣ
    On Error GoTo errHandle
    gstrSQL = "select A.��Ժ����,B.סԺ��,D.���� as סԺ����,A.��Ժ����,A.סԺҽʦ,C.����," & _
            "C.���� from ������ҳ A,������Ϣ B,�����ʻ� C,���ű� D " & _
            "Where A.����ID = B.����ID And A.����ID = C.����ID And " & _
            "A.��Ժ����ID = D.ID And A.��ҳID = [2] And A.����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, lng��ҳID)
    datCurr = rsTmp!��Ժ����
    strInNote = ��ȡ���Ժ���(lng����ID, lng��ҳID)   '��Ժ���
    strInNote = ToVarchar(strInNote, 64)
    If rsTmp.BOF Then ��Ժ�Ǽ�_��Ԫ = False: Exit Function
    'ǿ��ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_��Ԫ
    
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ������")
    If rsTemp.State = 1 Then
        lng����ID = rsTemp("ID")
        str���� = rsTemp!����
        str���ֱ��� = rsTemp!ID
    Else
        ��Ժ�Ǽ�_��Ԫ = False
        Exit Function
    End If
    
    If gstrҽ���������� = "" Then
        MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = gy_getybjgbm(gstrOutPara)
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
    mblnReturn = gy_reg(gstrҽ����������, gstrҽԺ����, 1, UserInfo.����, Format(zlDatabase.Currentdate, "yyyy-MM-dd"), "0", gstrOutPara)
    Call WriteBusinessLOG("Reg", lng����ID & "_" & lng��ҳID & "|" & gstrҽ���������� & "," & gstrҽԺ���� & "," & 1 & "," & UserInfo.���� & "," & Format(zlDatabase.Currentdate, "yyyy-MM-dd") & "," & "0", gstrOutPara.out1)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        ��Ժ�Ǽ�_��Ԫ = False
        Exit Function
    End If
    str������ = gstrOutPara.out1
    
    initType
'��ڲ���:ҽ����������,ҽԺ���,ҽ��������,ҽԺ��������,ҽԺ��������,��������,ԭ��
'         �����־, ҽ������,�ز���־
    '������Ժ����
    mblnReturn = gy_request(gstrҽ����������, gstrҽԺ����, str������, "��Ժ���", strInNote, Format(datCurr, "yyyy-MM-dd"), strInNote, "0", UserInfo.����, "0", gstrOutPara)
    Call WriteBusinessLOG("Request", lng����ID & "_" & lng��ҳID & "|" & gstrҽ���������� & "," & gstrҽԺ���� & "," & str������ & ",��Ժ���," & strInNote & "," & Format(datCurr, "yyyy-MM-dd") & "," & strInNote & "," & "0" & "," & UserInfo.���� & "," & "0", "�ɹ�")
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        ��Ժ�Ǽ�_��Ԫ = False
        Exit Function
    End If
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ԫ & ",'˳���'," & str������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_��Ԫ")
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ԫ & ",'����ID'," & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_��Ԫ")
    
     '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��Ԫ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_��Ԫ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call WriteBusinessLOG(lng����ID & "_" & lng��ҳID & "|" & "��������", "������Ϣ��" & Err.Description, "")
    ��Ժ�Ǽ�_��Ԫ = False
End Function


Public Function סԺ�������_��Ԫ(lng����ID As Long) As Boolean
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
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��Ԫ)
    str������ = Nvl(rsTemp!˳���, "0")

    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B" & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    lng����ID = rsTemp("ID")

    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ԫ, lng����ID)

    If rsTemp.EOF = True Then
        MsgBox "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If

    If IsNull(rsTemp!��ע) Then
        MsgBox "�õ��ݵľ����Ŷ�ʧ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If

    str������ = rsTemp!��ע
    cur�����ʻ� = rsTemp!�����ʻ�֧��

    '���ýӿ�������
    initType
    mblnReturn = gy_canrollback(gstrҽ����������, gstrҽԺ����, str������, gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If

    initType
    mblnReturn = gy_rollbackcalc(gstrҽ����������, gstrҽԺ����, str������, "0", gstrOutPara)

    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��Ԫ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)

    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_��Ԫ & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - Nvl(rsTemp("����ͳ����"), 0) & "," & _
        curͳ�ﱨ���ۼ� - Nvl(rsTemp("ͳ�ﱨ�����"), 0) & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��Ԫ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & curƱ���ܽ�� * -1 & ",0,0," & _
        Nvl(rsTemp("����ͳ����"), 0) * -1 & "," & Nvl(rsTemp("ͳ�ﱨ�����"), 0) * -1 & ",0," & Nvl(rsTemp("�����Ը����"), 0) & "," & _
        cur�����ʻ� * -1 & ",NULL,NULL,NULL,'" & str������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    סԺ�������_��Ԫ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_��Ԫ(lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ�
'        ���������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ����
'        ����һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim lng����ID  As Long, lng��ҳID As Long
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str����Ա As String, datCurr As Date, str������ As String
    Dim intסԺ�����ۼ� As Integer, cur�ʻ������ۼ� As Currency
    Dim cur�ʻ�֧���ۼ� As Currency, cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim cur�����ʻ� As Currency, cur���� As Currency, cur����ͳ���޶� As Currency
    Dim cur���ͳ���޶� As Currency, cur�����Ը� As Currency, cur��� As Currency
    Dim cur�������� As Currency, curȫ�Ը� As Currency, cur���Ը� As Currency
    
    On Error GoTo errHandle
    '��Ҫ���ϴ�������ϸ
'    ������ϸ����_��Ԫ lng����ID
    
    If gstrҽ���������� = "" Then
        Err.Raise 9000, gstrSysName, "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = gy_getybjgbm(gstrOutPara)
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
    
    gstrSQL = "Select * From סԺ���ü�¼ Where nvl(���ӱ�־,0)<>9 and ����ID=[1]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    
    If rs��ϸ.EOF = True Then
        Err.Raise 9000, gstrSysName, "û����д�շѼ�¼", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    lng��ҳID = rs��ϸ("��ҳID")
    str����Ա = UserInfo.����
    
    '�������һ�Ŵ������˴�������ϸ��������¼��Ժ���
    If Not WriteOutDisease(lng����ID, lng��ҳID) Then Exit Function
    
    gstrSQL = "Select nvl(˳���,0) as ˳��� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��Ԫ)
    str������ = rsTemp!˳���
    
    'ȡ���˳�Ժ����
    gstrSQL = "Select ��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˳�Ժ����", lng����ID, lng��ҳID)
    datCurr = Nvl(rsTemp!��Ժ����, zlDatabase.Currentdate())
    
    'ҽ����������, ҽԺ���, ҽ�������ţ� ��Ժ���ڣ�����Ա����ʾ��־
    initType
    mblnReturn = gy_calc(gstrҽ����������, gstrҽԺ����, str������, Format(datCurr, "yyyy-MM-dd"), str����Ա, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        Err.Raise 9000, gstrSysName, gstrOutPara.errtext, vbInformation, gstrSysName
        סԺ����_��Ԫ = False
        Exit Function
    End If
'��ӳ��ڲ���:1���úϼ�,2���ⲡ�ַ���,3���α����ʻ�֧��,4���������ʻ�֧��,5�ۼƷֶ��Ը�,6ͳ���֧��,7�𸶶�֧��,
'             8��λ֧��,9�Էѷ���,10�ؼ����Ը�,11�������Ը�,12�ؼ����,13���η���,14����ҽ�Ʊ���֧��,15����ͳ������ۼ�,
'             16����ҽ�Ƽ����ۼ�,17����ͳ������ۼ�,18δ��������,19ҽ��֧��,20�����ֽ�֧��,21�����ʻ����

    '��ȡ�����ʻ�֧���͸����ֽ�֧��
    cur�����ʻ� = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
    cur��� = CCur(gstrOutPara.out21)
    curȫ�Ը� = CCur(gstrOutPara.out20) - cur�����ʻ�
    cur�������� = CCur(gstrOutPara.out1)
    cur���Ը� = CCur(gstrOutPara.out10) + CCur(gstrOutPara.out11)
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_��Ԫ, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & _
            "," & TYPE_��Ԫ & "," & Year(datCurr) & "," & cur�ʻ������ۼ� & _
            "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & "," & _
            cur���� & "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '���ս����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_��Ԫ & "," & _
            lng����ID & "," & Year(datCurr) & "," & _
            cur��� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",NULL,NULL,NULL," & _
            cur�������� & "," & curȫ�Ը� & "," & cur���Ը� & ",NULL,NULL,NULL,NULL," & _
            cur�����ʻ� & ",NULL,NULL,NULL,'" & str������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    סԺ����_��Ԫ = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_��Ԫ(rs������ϸ As Recordset, lng����ID As Long, strҽ���� As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rs������ϸ-��Ҫ����ķ�����ϸ��¼����
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    
    Dim cur�����ʻ�֧�� As Currency, cur�����ֽ�֧�� As Currency
    Dim curͳ��֧�� As Currency, curҽ��֧�� As Currency, cur����ҽ�� As Currency
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str����Ա As String, datCurr As Date, str������ As String
    Dim curCount As Currency
    
    On Error GoTo errHandle
    '��Ҫ���ϴ�������ϸ
'    ������ϸ����_��Ԫ 0, rs������ϸ
'
    If gstrҽ���������� = "" Then
        MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = gy_getybjgbm(gstrOutPara)
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
        MsgBox "����û�з���סԺ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    str����Ա = UserInfo.����
    
    If ���ʴ���_��Ԫ("", 0, "", lng����ID) = False Then Exit Function
    
    gstrSQL = "Select nvl(˳���,0) as ˳���,���� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��Ԫ)
    str������ = rsTemp!˳���
'    gstrҽ���������� = rsTemp!����
    'ҽ����������, ҽԺ���, ҽ�������ţ� ��Ժ���ڣ�����Ա����ʾ��־
    datCurr = zlDatabase.Currentdate
    initType
    mblnReturn = gy_pcalc(gstrҽ����������, gstrҽԺ����, str������, Format(datCurr, "yyyy-MM-dd"), str����Ա, "1", "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        סԺ�������_��Ԫ = ""
        Exit Function
    End If
'��ӳ��ڲ���:1���úϼ�,2���ⲡ�ַ���,3���α����ʻ�֧��,4���������ʻ�֧��,5�ۼƷֶ��Ը�,6ͳ���֧��,7�𸶶�֧��,
'             8��λ֧��,9�Էѷ���,10�ؼ����Ը�,11�������Ը�,12�ؼ����,13���η���,14����ҽ�Ʊ���֧��,15����ͳ������ۼ�,
'             16����ҽ�Ƽ����ۼ�,17����ͳ������ۼ�,18δ��������,19ҽ��֧��,20�����ֽ�֧��,21�����ʻ����
    

    '��ȡ�����ʻ�֧���͸����ֽ�֧��
    cur�����ʻ�֧�� = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
    cur�����ֽ�֧�� = CCur(gstrOutPara.out20)
    curͳ��֧�� = CCur(gstrOutPara.out6)
    curҽ��֧�� = CCur(gstrOutPara.out19)
    cur����ҽ�� = CCur(gstrOutPara.out14)
    If curCount <> CCur(gstrOutPara.out1) Then
        MsgBox "��ע�⣺ҽ�����ؽ������뵱ǰ���ݽ���" & vbCrLf & _
                       "ҽԺ�ܶ" & curCount & "    ҽ�����أ�" & CCur(gstrOutPara.out1), vbInformation, gstrSysName
    End If
    סԺ�������_��Ԫ = "�����ʻ�;" & cur�����ʻ�֧�� & ";0" '�������޸ĸ����ʻ�
'    If cur�����ֽ�֧�� <> 0 Then
'        סԺ�������_��Ԫ = סԺ�������_��Ԫ & "|�ֽ�;" & cur�����ֽ�֧�� & ";0" '�������޸��ֽ�֧��
'    End If
    If curͳ��֧�� <> 0 Then
        סԺ�������_��Ԫ = סԺ�������_��Ԫ & "|ҽ������;" & curͳ��֧�� & ";0" '�������޸�ͳ��֧��
    End If
    If cur����ҽ�� <> 0 Then
        סԺ�������_��Ԫ = סԺ�������_��Ԫ & "|����ҽ�Ʊ���;" & cur����ҽ�� & ";0"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    סԺ�������_��Ԫ = ""
End Function

Public Function ��Ժ�Ǽ�_��Ԫ(lng����ID As Long, lng��ҳID As Long) As Boolean
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
        mblnReturn = gy_getybjgbm(gstrOutPara)
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
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��Ԫ)
        str������ = rsTemp!˳���
        initType
        mblnReturn = gy_dall(gstrҽ����������, gstrҽԺ����, str������, gstrOutPara)
        If mblnReturn = False Then
            ��Ժ�Ǽ�_��Ԫ = False
            Exit Function
        End If
    End If
    
    '��HIS֮�еĻ������ݽ����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_��Ԫ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_��Ԫ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ��Ժ�Ǽ�_��Ԫ = False
End Function

Private Function WriteOutDisease(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim str������ As String, str������ As String, datCurr As Date
    Dim lng����ID As Long, str�������� As String, str���ֱ��� As String
    Dim int�ز���־ As Integer, lng����ID As Long, str���ұ�� As String, str�������� As String
    Dim rsTemp As New ADODB.Recordset
    '��д��Ժ��ϣ��������һ�Ŵ������˴�������ϸ��������¼��Ժ��ϣ�
    
    On Error GoTo errHand
    
    'ȡ���˳�Ժ����
    gstrSQL = "Select ��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˳�Ժ����", lng����ID, lng��ҳID)
    datCurr = Nvl(rsTemp!��Ժ����, zlDatabase.Currentdate())
    
    gstrSQL = "Select nvl(˳���,0) as ˳���,����ID,����,����֤�� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID, TYPE_��Ԫ)
    str������ = rsTemp!����֤��
    str������ = rsTemp!˳���
    lng����ID = Nvl(rsTemp!����ID, 0)
    
    '�ж��Ƿ�Ϊ���ⲡ
    gstrSQL = "Select * From ���ղ��� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    If rsTemp.EOF Then
        int�ز���־ = 0
    Else
        int�ز���־ = IIf(rsTemp!��� = 2, 1, 0)
    End If
    
    '��Ժ��ϣ����ֱ���̶�ΪҽԺ���룬��������Ϊ��Ժ�������
    str�������� = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, False)
    str���ֱ��� = "��Ժ���"
    
    lng����ID = UserInfo.����ID
    gstrSQL = "Select * From ���ű� where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng����ID)
    str���ұ�� = rsTemp!����
    str�������� = rsTemp!����
    
    'д����ͷ
    initType
    mblnReturn = gy_wrecipe(gstrҽ����������, gstrҽԺ����, str������, str������, str���ֱ���, str��������, _
                         int�ز���־, UserInfo.����, UserInfo.����, str���ұ��, _
                         str��������, Format(datCurr, "yyyy-MM-dd"), gstrOutPara)
    TrimType
    If mblnReturn = False Then
        If InStr(gstrOutPara.errtext, "(YBYY.PRI_QTYL42_T)") > 0 Then
            WriteOutDisease = True
        Else
            MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
            WriteOutDisease = False
            Exit Function
        End If
    End If
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_��Ԫ & ",'����֤��','" & CLng(str������) + 1 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    WriteOutDisease = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ������_��Ԫ() As Boolean
    ҽ������_��Ԫ = frmSet��Ԫ.ShowME(TYPE_��Ԫ)
End Function

Private Function Get����ID(strҽ���� As String, strҽ�����ı��� As String) As String
'���ܣ�ͨ��ҽ�����ĺ����ҽ�����������ID
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select ����ID from �����ʻ� where ���� = [1] and ҽ���� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ԫ, strҽ����)
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

Public Function ���ʴ���_��Ԫ(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0) As Boolean
    Dim lng��ҳID As Long
    Dim rsTemp As New ADODB.Recordset
    
    If str���ݺ� <> "" Then
        gstrSQL = " Select A.* From סԺ���ü�¼ A,�����ʻ� B" & _
                  " Where A.��¼״̬<>0 And Nvl(A.�Ƿ��ϴ�,0)=0 And nvl(A.���ӱ�־,0)<>9 " & _
                  " and A.��¼����=[1] and A.NO=[2]" & _
                  " and A.����ID=B.����ID And B.����=[3]" & _
                  " order by A.��ҳID,A.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, TYPE_��Ԫ, int����, str���ݺ�)
    Else
        '��ȡ�ò��˱���סԺ����ҳID
        gstrSQL = "Select סԺ���� From ������Ϣ Where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��˱���סԺ����ҳID", lng����ID)
        lng��ҳID = Nvl(rsTemp!סԺ����, 1)
        
        '��ȡ����סԺ������ϸ
        gstrSQL = " Select * From סԺ���ü�¼ " & _
                  " Where ��¼״̬<>0 And Nvl(�Ƿ��ϴ�,0)=0 And nvl(���ӱ�־,0)<>9 And Nvl(ʵ�ս��,0)<>0" & _
                  " and ����id=" & lng����ID & " And ��ҳID=" & lng��ҳID & _
                  " order by ��ҳID,���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSQL, lng����ID, lng��ҳID)
    End If
    
    ���ʴ���_��Ԫ = ������ϸ����_��Ԫ(0, rsTemp, (str���ݺ� <> ""))
End Function
