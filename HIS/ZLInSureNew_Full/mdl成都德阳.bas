Attribute VB_Name = "mdl�ɶ�����"
Option Explicit
Private mblnInit As Boolean

Private gobjTest As Object

Public Enum ҵ������_�ɶ�����
    ����籣���� = 0
    ��òα���Ա����
    ��Ժ�Ǽ�
    ȡ����Ժ�Ǽ�
    ��Ժ�Ǽ�
    ȡ����Ժ�Ǽ�
    ���Ӵ�������
    ���Ӵ�����ϸ
    ɾ���������ݼ�����ϸ
    ������������
    ��Ժ����
    ȡ����Ժ����
    
    ��ӡ��Ժ���㱨����
    ��ӡסԺ��Ա������㵥
    ������Ա��������
    ��ȡ��������
    ��ȡסԺ��¼��
    ���κ�����
    �����κ�����
    �Ͽ��κ�����
    ��ȡҩƷ��Ϣ
    ����������Ա��������
End Enum
Private Type InitbaseInfor
    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ҽԺ���� As String                      '��ʼҽԺ����
    �������� As String                      'Ĭ�ϵ��籣��������
    ��ϸʱʵ�ϴ� As Boolean
    ���ݲ��Ȳ��ɽ��� As Boolean
    ��ӡ���㵥      As Boolean
    �Զ���Ժ���    As Boolean
End Type
Public InitInfor_�ɶ����� As InitbaseInfor

Private Type �������
        ��¼��        As String
        ���Ϻ�    As String       '��ҽ����
        ����     As String
        �Ա�     As String
        ��������  As String
        ����        As Integer
        ҽ������    As String
        ���ݹ���    As String
        ��λ����    As String
        ��λ����    As String
        ҽ�Ʊ�־    As String
        ��������    As String
        
        �����ܶ�    As Double
        ����ID      As Long
        ���ֱ���    As String
        ��������    As String
        ����        As Long
End Type
Private Type ��������
    ҽ������ As Double
    �����㸶�� As Double
End Type
Private gblnδ������ As Boolean

Private g������� As ��������
Public g�������_�ɶ����� As �������
Public gcnOracle_�ɶ����� As ADODB.Connection     '�м������
Private gbln������� As Boolean
Private gbln�Ѿ���ʼ As Boolean             '�Ѿ�����ʼ����.
'1.����籣������ź������б�
Private Declare Function GetSBJGLB Lib "CDGK_YB.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION GETSBJGLB:PCHAR
'����: ����籣������ź������б�
'��ڲ���: ��
'���ڲ���: ��
'����: A�籣�������+�м����+A�籣��������+�м����+B�籣�������+�м����+B�籣��������+����
'===============================================================================================================

'2����òα���Ա�Ļ�������
Private Declare Function GETRYJBZL Lib "CDGK_YB.dll" Alias "GetRYJBZL" (ByVal str���Ϻ� As String, ByVal str�籣��� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETRYJBZL(ASBBH,ABXJGBH:PCHAR):PCHAR;
'����: ��òα���Ա�Ļ�������
'��ڲ���: ASBBH   PCHAR   �α���Ա����ᱣ�Ϻ�
'          ABXJGBH PCHAR   �α���Ա���ڵı��ջ������
'���ڲ���: ��
'����: A�籣�������+�м����+A�籣��������+�м����+B�籣�������+�м����+B�籣��������+����
'===============================================================================================================

'3.��Ժ�Ǽ�
Private Declare Function RYDJ Lib "CDGK_YB.dll" (ByVal strסԺ�� As String, ByVal str�������� As String, ByVal str������� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION RYDJ(AZYH,;ARYZL,ABXJGBH:PCHAR):PCHAR;
'����: ���籣����ҽ�����ݿ��ҽԺ����ҽ�����ݿ��ж�סԺ��ҽ�����˽��еǼǡ�
'��ڲ���: strסԺ��   PCHAR   סԺ��
'          str�������� PCHAR   �α���Ա�ĸ�������
'          str������� PCHAR �α���Ա���ڵ��籣�������
'���ڲ���: ��
'����:���ر�־@$��ᱣ�Ϻ�||���˼�¼��||ҽ������||���ݹ���||��λ����||����||�Ա�||�������ڣ���ʽ��YYYY-MM-DD��||��λ���||�μӻ���ҽ�Ʊ�־||��Ժ���ڣ���ʽ��YYYY-MM-DD��||���ֱ��||��������||����
'===============================================================================================================

'4.ȡ��סԺ
Private Declare Function ZYQX Lib "CDGK_YB.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION ZYQX(AZYH:PCHAR):PCHAR
'����: ���籣����ҽ�����ݿ��ҽԺ����ҽ�����ݿ���ɾ��ҽ������סԺ��¼��
'��ڲ���: strסԺ��   PCHAR   סԺ��
'���ڲ���: ��
'����:���ر�־
'===============================================================================================================

'5.��Ժ�Ǽ�
Private Declare Function CYCS Lib "CDGK_YB.dll" (ByVal strסԺ�� As String, ByVal str��Ժ���� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CYCS(AZYH ,CYRQ:PCHAR):PCHAR;
'����: ��ҽ������סԺ���������������ϴ����籣����ҽ�����ݿ⣻�Ա���ҽ�����ݿ���ҽ����������Ժ����
'��ڲ���: strסԺ��   PCHAR   סԺ��
'          str��Ժ���� pchar ��Ժ���ڣ�YYYY-MM-DD��
'���ڲ���: ��
'����:���ر�־
'===============================================================================================================

'6.ȡ����Ժ�Ǽ�
Private Declare Function CYCSQX Lib "CDGK_YB.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CYCSQX (AZYH:PCHAR):PCHAR;
'����:ȡ���α��������籣���Ѿ�����ĳ�Ժ���ݣ��Ա����´��䡣
'��ڲ���: strסԺ��   PCHAR   סԺ��
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================


'7.����һ����������
Private Declare Function AddCFJL Lib "CDGK_YB.dll" (ByVal strסԺ�� As String, ByVal str�������� As String, ByVal strҽ�� As String, ByVal str���� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION ADDCFJL(AZYH,ACFRQ,AYS,AKS:PCHAR):PCHAR
'����:����һ���������ݡ���
'��ڲ���:
'        AZYH    PCHAR   סԺ��
'        ACFRQ   PCHAR   �������ڣ�YYYY-MM-DD��
'        AYS PCHAR   ҽ��
'        AKS PCHAR   ����
'���ڲ���: ��
'����:'OK'+�м����+������¼�Ż������Ϣ
'===============================================================================================================

'7.���Ӵ�����ϸ
Private Declare Function AddCFMX Lib "CDGK_YB.dll" (ByVal str������¼�� As String, ByVal strҽ������ As String, ByVal str���� As String, ByVal str���� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION ADDCFMX(ACFID,AYPBH,ASL,ADJ:PCHAR):PCHAR;
'����:����һ��������ϸ��
'��ڲ���:
'    ACFID   PCHAR   ������¼��
'    AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
'    ASL PCHAR   ����(����Ϊ����)
'    ADJ PCHAR   ����
'���ڲ���: ��
'����:'OK'+�м����+������ϸ��¼��+�м����+�Էѱ���+�м����+�Էѽ��������Ϣ
'===============================================================================================================

'8.ɾ���������ݼ�����ϸ
Private Declare Function DELCFJL Lib "CDGK_YB.dll" (ByVal str������¼�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION DELCFJL(ACFID:PCHAR):PCHAR
'����:ɾ���������ݼ�����������ϸ��¼��
'��ڲ���:
'    ACFID   PCHAR   ������¼��
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================


'9.������������
Private Declare Function CFCS Lib "CDGK_YB.dll" (ByVal strסԺ�� As String, ByVal str������¼�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CFCS(AZYH:PCHAR;ACFID:PCHAR):PCHAR
'����:���籣����ÿ��Ĵ���������籣���������ݿ⴫�䣨ͬһ���������Զ���ظ����䣬��һ�δ�������ݽ�����ǰһ�δ�������ݣ�
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'    ACFID   PCHAR   ������¼�ţ�ͨ������ADDCFJL���ص�ֵ��
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================

'10.��Ժ����
Private Declare Function CYJS Lib "CDGK_YB.dll" (ByVal strסԺ�� As String, ByVal lngԤ���־ As Long) As String
'===============================================================================================================
'ԭ��:FNCTION CYJS(AZYH:PCHAR; ISPREV:INTEGER):PCHAR
'����:סԺ�α����˳�Ժ��סԺ��Ԥ����
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'    ISPREV  PCHAR   Ԥ�����־��'0'����ʾԤ���㣩
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================

'11.ȡ����Ժ����
Private Declare Function CYJSQX Lib "CDGK_YB.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CYJSQX(AZYH:PCHAR):PCHAR
'����:ȡ���α����˳�Ժ����
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================

'12.��ӡ��Ժ���㱨����
Private Declare Function JSReport Lib "CDGK_YB.dll" (ByVal str��ʼסԺ�� As String, ByVal str����סԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION JSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL
'����:��ӡ�籣�����ṩ�Ķ�̬����Ŀǰ�����������ö�̬����"סԺ����ͳ�Ʊ����䣩"��"����סԺ���㵥"��"סԺ����ͳ�Ʊ�"���ű���ʹ��"21����ȡ��������"�������Զ����±��ر���
'��ڲ���:
'    ASTARTZYH   PCHAR   ��ӡ��ʼסԺ��
'    AENDZYH PCHAR      ��ӡ����סԺ��
'   ע��:
'    1 ?����סԺ��֮�����е�סԺ��¼����Ϊͬһ���籣��?
'    2����ֻ��ӡһ��סԺ�ŵı���ʱ����������ֵһ����
'���ڲ���: ��
'����:����ע�ⷵ��ֵ
'===============================================================================================================

'13.��ӡסԺ��Ա������㵥
Private Declare Function CWJSREPORT Lib "CDGK_YB.dll" Alias "CWJSReport" (ByVal str��ʼסԺ�� As String, ByVal str����סԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION CWJSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL;
'����:��ӡסԺ��Ա������㵥��
'��ڲ���:
'    ASTARTZYH   PCHAR   ��ӡ��ʼסԺ��
'    AENDZYH PCHAR      ��ӡ����סԺ��
'   ע��:
'    1 ?����סԺ��֮�����е�סԺ��¼����Ϊͬһ���籣��?
'    2����ֻ��ӡһ��סԺ�ŵı���ʱ����������ֵһ����
'���ڲ���: ��
'����:����ע�ⷵ��ֵ
'===============================================================================================================

'14.��ȡ��������
Private Declare Function GETJCXX Lib "CDGK_YB.dll" Alias "GetJCXX" (ByVal str������� As String, ByVal str���ر�־ As String) As String
'===============================================================================================================
'ԭ��:GETJCXX(SBXJGBH:PCHAR;DOWNALL:INTEGER):PCHAR
'����:��ָ�����籣������ȡ��������
'��ڲ���:
'    SBXJGBH PCHAR   ���ջ������
'    DOWNALL PCHAR   ��ֵΪ0ʱ��ʾ���ر���ҽ�����ݿ���û�еĻ������ϣ�Ϊ����ʱ��ʾȫ����������
'���ڲ���: ��
'����:'OK'�������Ϣ
'===============================================================================================================

'15 ����סԺ�ŵõ�סԺ��¼��
Private Declare Function GETZYIDBYZYBH Lib "CDGK_YB.dll" Alias "GetZYIDByZyBH" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETZYIDBYZYBH(AZYH:PCHAR):PCHAR
'����:����סԺ�ŵõ�סԺ��¼��
'��ڲ���:
'   AZYH    PCHAR   סԺ��'���ڲ���: ��
'����:'OK'@$סԺ��¼�Ż������Ϣ
'===============================================================================================================

'16.���κ������Ƿ����ӳɹ�
Private Declare Function CheckCon Lib "CDGK_YB.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION CHECKCON:PCHAR;
'����:���κ������Ƿ����ӳɹ�
'��ڲ���:
'����:OK�������Ϣ
'===============================================================================================================

'17.�����κ�����
Private Declare Function RasDial Lib "CDGK_YB.dll" (ByVal str�������� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION RASDIAL(SBXJGBH:PCHAR):PCHAR
'����:�κ���ѡ����籣�֣����佨������
'��ڲ���:SBXJGBH PCHAR   ���ջ������
'����:  �ɹ�    ������HIS�κ���״̬����ʾ"����"
'       ʧ�� ������Ϣ
'===============================================================================================================

'18.�Ͽ����籣�ֵ�����
Private Declare Function DisDial Lib "CDGK_YB.dll" () As String
'===============================================================================================================
'ԭ��:FUNCTION DISDIAL:PCHAR
'����:�κ���ѡ����籣�֣����佨������
'��ڲ���:
'����:
'===============================================================================================================


'19 ����ҩƷ��ŵõ�ҩƷ��Ϣ
Private Declare Function GetSINYPXX Lib "CDGK_YB.dll" (ByVal str�������� As String, ByVal strҩƷ���� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETSINYPXX(SBXJGBH,CYPBH:PCHAR):PCHAR
'����:����ҩƷ��ŵõ�ҩƷ��Ϣ
'��ڲ���:
'    SBXJGBH PCHAR   ���ջ������
'    CYPBH   PCHAR   ҩƷ���
'����:OK@$���:ҩƷ||��������:��Ī�����ƣ�����ά��أ�||������λ:֧||��������:0||�Էѱ���:20
'===============================================================================================================

'20.סԺ���˿���������Ա��������
Private Declare Function GETNEWRYZL Lib "CDGK_YB.dll" (ByVal strסԺ�� As String) As String
'===============================================================================================================
'ԭ��:FUNCTION GETNEWRYZL(AZYH:PCHAR):PCHAR;STDCALL;
'����:סԺ���˿���������Ա�������ϡ�
'��ڲ���:
'    AZYH    PCHAR   סԺ��
'����:OK@$������Ϣ
'===============================================================================================================






Public Function ҽ����ʼ��_�ɶ�����() As Boolean
    Dim strReg As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    If mblnInit Then
        ҽ����ʼ��_�ɶ����� = True
        Exit Function
    End If
    '��ʼģ��ӿ�
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        InitInfor_�ɶ�����.ģ������ = True
    Else
        InitInfor_�ɶ�����.ģ������ = False
    End If
    
   Call GetRegInFor(g����ģ��, "ҽ��", "�籣��������", strReg)
   
   InitInfor_�ɶ�����.�������� = strReg
   If strReg = "" Then
        MsgBox "��δ����Ĭ�ϵ��籣�������룬�����������!"
        Exit Function
   End If
   
    'ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", TYPE_�ɶ�����)
    InitInfor_�ɶ�����.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
    
    
    '�м������
    gstrSQL = "select ������,����ֵ from ���ղ��� where  ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ɶ�����ҽ��", TYPE_�ɶ�����)
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "��鲦������"
                 gbln������� = Nvl(rsTemp("����ֵ"), 0) = 1
            Case "��ϸʱʵ�ϴ�"
                InitInfor_�ɶ�����.��ϸʱʵ�ϴ� = Nvl(rsTemp("����ֵ"), 0) = 1
            Case "�ȽϽ�������"
                 InitInfor_�ɶ�����.���ݲ��Ȳ��ɽ��� = IIf(Nvl(rsTemp!����ֵ, 1) = 1, 1, 0)
            Case "��ӡ���㵥"
                 InitInfor_�ɶ�����.��ӡ���㵥 = IIf(Nvl(rsTemp!����ֵ, 1) = 1, 1, 0)
            Case "�Զ���Ժ���"
                 InitInfor_�ɶ�����.�Զ���Ժ��� = IIf(Nvl(rsTemp!����ֵ, 1) = 1, 1, 0)
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_�ɶ����� = New ADODB.Connection

    If OraDataOpen(gcnOracle_�ɶ�����, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    
   '�����κ�����
   If gbln�Ѿ���ʼ = False And gbln������� Then
        If ҵ������_�ɶ�����(�����κ�����, InitInfor_�ɶ�����.��������, strOutput) = False Then
             Exit Function
        End If
   End If
   
   If gbln������� Then
        '���κ�����
        If ҵ������_�ɶ�����(���κ�����, "", strOutput) = False Then
             Exit Function
        End If
    End If
    mblnInit = True
    gbln�Ѿ���ʼ = True
    ҽ����ʼ��_�ɶ����� = True
End Function

Public Function ҽ����ֹ_�ɶ�����() As Boolean
    Dim strOutput As String
    mblnInit = False
    If gcnOracle_�ɶ�����.State = 1 Then
        gcnOracle_�ɶ�����.Close
    End If
    '�����κ�����
   Call ҵ������_�ɶ�����(�Ͽ��κ�����, "", strOutput)
    Err = 0
    On Error Resume Next
    ҽ����ֹ_�ɶ����� = True
End Function

Public Function ��ݱ�ʶ_�ɶ�����(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    '���أ��ջ���Ϣ��
    Err = 0
    On Error GoTo errHand:
    If bytType = 0 Or bytType = 3 Then Exit Function
    
    ��ݱ�ʶ_�ɶ����� = frmIdentify�ɶ�����.GetPatient(bytType, lng����ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_�ɶ����� = ""
End Function


Public Function �������_�ɶ�����(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, TYPE_�ɶ�����)
    
    If rsTemp.EOF Then
        �������_�ɶ����� = 0
    Else
        �������_�ɶ����� = rsTemp("�ʻ����")
    End If
End Function
Public Function �����������_�ɶ�����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    �����������_�ɶ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_�ɶ�����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    �������_�ɶ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����������_�ɶ�����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    ����������_�ɶ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_�ɶ�����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutput As String, StrInput As String
    Dim strArr
    Err = 0: On Error GoTo errHand:
    If InitInfor_�ɶ�����.�������� <> g�������_�ɶ�����.�������� Then
        '�����κ�����
        If gbln�Ѿ���ʼ = False And gbln������� Then
             If ҵ������_�ɶ�����(�����κ�����, g�������_�ɶ�����.��������, strOutput) = False Then
                  Exit Function
             End If
        End If
        
        If gbln������� Then
             '���κ�����
             If ҵ������_�ɶ�����(���κ�����, "", strOutput) = False Then
                  Exit Function
             End If
         End If
    End If
    
'    '��ȡסԺ��
'    gstrSQL = "Select סԺ�� From ������Ϣ where ����id=" & lng����id
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡסԺ��"
'
    Dim strסԺ�� As String
    strסԺ�� = GetסԺ��(lng����ID, lng��ҳID)
    If strסԺ�� = "" Then Exit Function
    
    'סԺ��||��������||�籣�������
    StrInput = strסԺ��
    StrInput = StrInput & "||" & Get��������(lng����ID, lng��ҳID)
    StrInput = StrInput & "||" & g�������_�ɶ�����.��������
    If ҵ������_�ɶ�����(��Ժ�Ǽ�, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    '��ᱣ�Ϻ�||���˼�¼��||ҽ������||���ݹ���||��λ����||����||�Ա�||�������ڣ���ʽ��YYYY-MM-DD��||��λ���||�μӻ���ҽ�Ʊ�־||��Ժ���ڣ���ʽ��YYYY-MM-DD��||���ֱ��||��������||����
    strArr = Split(strOutput, "||")
    '������ص���Ϣ
    ''OK'+�м����+�籣����סԺ��¼��
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�ɶ����� & ",'ҽ��סԺ��','''" & strסԺ�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ҽ��סԺ��")
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�ɶ����� & ",'סԺ��¼��','''" & Val(strArr(0)) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��¼��")
'    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�ɶ����� & ",'��������','''" & strArr(12) & "''')"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
    
    '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ɶ����� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    ��Ժ�Ǽ�_�ɶ����� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_�ɶ����� = False
End Function
Private Function Get��������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As String
    '    ��ᱣ�Ϻ�|���˼�¼��|ҽ������|���ݹ���|��λ����|����|�Ա�|�������ڣ���ʽ��YYYY-MM-DD��
    '    ��λ���|�μӻ���ҽ�Ʊ�־|��Ժ���ڣ���ʽ��YYYY-MM-DD��|���ֱ��|��������|����
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String
    gstrSQL = "" & _
        "   Select  to_char(a.��Ժ����,'yyyy-mm-dd') as ��Ժ����,b.���� as ����" & _
        "   From ������ҳ a,���ű� b " & _
        "   Where A.��Ժ����ID=b.id(+) and a.����id=" & lng����ID & " and a.��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��ҳ��Ϣ"
    With g�������_�ɶ�����
        StrInput = .���Ϻ�
        StrInput = StrInput & vbTab & "|" & .��¼��
        StrInput = StrInput & vbTab & "|" & .ҽ������
        StrInput = StrInput & vbTab & "|" & .���ݹ���
        StrInput = StrInput & vbTab & "|" & .��λ����
        StrInput = StrInput & vbTab & "|" & .����
        StrInput = StrInput & vbTab & "|" & .�Ա�
        StrInput = StrInput & vbTab & "|" & .��������
        StrInput = StrInput & vbTab & "|" & .��λ����
        StrInput = StrInput & vbTab & "|" & .ҽ�Ʊ�־
        StrInput = StrInput & vbTab & "|" & Nvl(rsTemp!��Ժ����)
        StrInput = StrInput & vbTab & "|" & .���ֱ���
        StrInput = StrInput & vbTab & "|" & .��������
        StrInput = StrInput & vbTab & "|" & Nvl(rsTemp!����)
    End With
    Get�������� = StrInput
    
    
End Function
Private Function Get���״���(ByVal intType As ҵ������_�ɶ�����, Optional bln������ As Boolean = False) As String
    '������û��
    Select Case intType
        Case ����籣����
            Get���״��� = IIf(bln������, "����籣����", "01")
        Case ��òα���Ա����
            Get���״��� = IIf(bln������, "��òα���Ա����", "02")
        Case ��Ժ�Ǽ�
            Get���״��� = IIf(bln������, "��Ժ�Ǽ�", "03")
        Case ȡ����Ժ�Ǽ�
            Get���״��� = IIf(bln������, "ȡ����Ժ�Ǽ�", "04")
        Case ��Ժ�Ǽ�
            Get���״��� = IIf(bln������, "��Ժ�Ǽ�", "05")
        Case ȡ����Ժ�Ǽ�
            Get���״��� = IIf(bln������, "ȡ����Ժ�Ǽ�", "06")
        Case ���Ӵ�������
            Get���״��� = IIf(bln������, "���Ӵ�������", "07")
        Case ���Ӵ�����ϸ
            Get���״��� = IIf(bln������, "���Ӵ�����ϸ", "08")
        Case ɾ���������ݼ�����ϸ
            Get���״��� = IIf(bln������, "ɾ���������ݼ�����ϸ", "09")
        Case ������������
            Get���״��� = IIf(bln������, "������������", "10")
        Case ��Ժ����
            Get���״��� = IIf(bln������, "��Ժ����", "11")
        Case ȡ����Ժ����
            Get���״��� = IIf(bln������, "ȡ����Ժ����", "12")
        Case ��ӡ��Ժ���㱨����
            Get���״��� = IIf(bln������, "��ӡ��Ժ���㱨����", "13")
        Case ��ӡסԺ��Ա������㵥
            Get���״��� = IIf(bln������, "��ӡסԺ��Ա������㵥", "14")
        Case ������Ա��������
            Get���״��� = IIf(bln������, "������Ա��������", "15")
        Case ��ȡ��������
            Get���״��� = IIf(bln������, "��ȡ��������", "16")
        Case ��ȡסԺ��¼��
            Get���״��� = IIf(bln������, "��ȡסԺ��¼��", "17")
        Case ���κ�����
            Get���״��� = IIf(bln������, "���κ�����", "18")
        Case �����κ�����
            Get���״��� = IIf(bln������, "�����κ�����", "19")
        Case �Ͽ��κ�����
            Get���״��� = IIf(bln������, "�Ͽ��κ�����", "20")
        Case ��ȡҩƷ��Ϣ
            Get���״��� = IIf(bln������, "��ȡҩƷ��Ϣ", "21")
        Case ����������Ա��������
            Get���״��� = IIf(bln������, "����������Ա��������", "22")
        Case Else
            Get���״��� = IIf(bln������, "����Ľ��״���", "-1")
    End Select
End Function
Public Function ҵ������_�ɶ�����(ByVal intType As ҵ������_�ɶ�����, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:strOutPutString-�����,������˳��,��tab���ָ��ķ��ش�
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, lngReturn As Long, strOutput As String, strReturn As String
    Dim strInValue(0 To 20) As String
    
    Dim str���״��� As String
    Dim i As Integer
    Dim strArr
    
    str���״��� = Get���״���(intType)
    StrInput = str���״��� & "|" & strInputString
    DebugTool "����ҵ��������(ҵ������Ϊ:" & intType & "),�������Ϊ" & vbCrLf & str���״��� & "|" & StrInput
    
    ҵ������_�ɶ����� = False
    If InitInfor_�ɶ�����.ģ������ Then
        '��ȡģ������
        Readģ������ intType, StrInput, strOutPutstring
         ҵ������_�ɶ����� = True
        Exit Function
    End If
    strArr = Split(strInputString, "||")
    For i = 0 To UBound(strArr)
        strInValue(i) = strArr(i)
    Next
    
    Err = 0
    On Error GoTo errHand:
'
'    If gobjTest Is Nothing Then
'        Set gobjTest = CreateObject("cdgk_Yb.clscdgk_Yb")
'    End If
    
    Select Case intType
        Case ����籣����
            strOutput = GetSBJGLB()
            
            If strOutput = "" Then
                MsgBox "��ȡ�籣����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Mid(strOutput, 5)
        Case ��òα���Ա����
            strOutput = GETRYJBZL(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "��òα���Ա����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case ��Ժ�Ǽ�
            '
            strOutput = RYDJ(strInValue(0), Replace(strInValue(1), vbTab & "|", "||"), strInValue(2))
            If strOutput = "" Then
                MsgBox "��Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case ȡ����Ժ�Ǽ�
            strOutput = ZYQX(strInValue(0))
            If strOutput = "" Then
                MsgBox "ȡ����Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ��Ժ�Ǽ�
            strOutput = CYCS(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "��Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ȡ����Ժ�Ǽ�
            strOutput = CYCSQX(strInValue(0))
            If strOutput = "" Then
                MsgBox "ȡ����Ժ�Ǽ�ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ���Ӵ�������
            strOutput = AddCFJL(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            If strOutput = "" Then
                MsgBox "���Ӵ�������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case ���Ӵ�����ϸ
            strOutput = AddCFMX(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            If strOutput = "" Then
                MsgBox "���Ӵ�����ϸʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strOutput = ""
            For i = 1 To UBound(strArr)
                strOutput = strOutput & "||" & strArr(i)
            Next
            If strOutput <> "" Then
                strOutput = Mid(strOutput, 3)
            End If
        Case ɾ���������ݼ�����ϸ
            strOutput = DELCFJL(strInValue(0))
            If strOutput = "" Then
                MsgBox "ɾ���������ݼ�����ϸʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ������������
            strOutput = CFCS(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "������������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ��Ժ����
            strOutput = CYJS(strInValue(0), Val(strInValue(1)))
            If strOutput = "" Then
                MsgBox "��Ժ����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                If InStr(1, strArr(0), "δ�ﵽ�𸶱�׼��") <> 0 Then
                    '�����������Ҫ�����δ�����߽����׳��ô���Ϊ��������ͨ��������ֻ�ܲ�ȡ�˷����ظ�ֵ.
                    strArr = Split("dfds|d", "|")
                    strArr(1) = "0||0||0||" & g�������_�ɶ�����.�����ܶ� & "||0||" & g�������_�ɶ�����.�����ܶ� & "||0||0||0||0||0||0||0||" & Format(zlDatabase.Currentdate, "yyyy") & "||" & "" & "||||||||||||||0||0||0||0||0||0||0||0||0||0||0||0||0||0||0||0||0||0||1||1||0||||�ֽ�||||||||||||||||1922-02-02||0||0||0||0||0||0||0||||||||||||||0||0||0||0||0||0||0||0||||0||"
                    gblnδ������ = True
                Else
                    MsgBox strArr(0), vbInformation, gstrSysName
                    Exit Function
                    gblnδ������ = False
                End If
            Else
                gblnδ������ = False
            End If
            strOutput = strArr(1)
        Case ȡ����Ժ����
            strOutput = CYJSQX(strInValue(0))
            If strOutput = "" Then
                MsgBox "��Ժ����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ��ӡ��Ժ���㱨����
            strOutput = JSReport(strInValue(0), strInValue(1))
            strOutput = ""
        Case ��ӡסԺ��Ա������㵥
            strOutput = CWJSREPORT(strInValue(0), strInValue(1))
            strOutput = ""
        
        Case ��ȡ��������
        
            strOutput = GETJCXX(strInValue(0), strInValue(1))
              If strOutput = "" Then
                MsgBox "��ȡ��������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case ��ȡסԺ��¼��
            strOutput = GETZYIDBYZYBH(strInValue(0))
            If strOutput = "" Then
                MsgBox "��ȡסԺ��¼��ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case ���κ�����
            strOutput = CheckCon()
            If strOutput = "" Then
                MsgBox "���κ�����ʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case �����κ�����
            strOutput = RasDial(strInValue(0))
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = ""
        Case �Ͽ��κ�����
            strOutput = DisDial()
            strOutput = ""
        Case ��ȡҩƷ��Ϣ
             strOutput = GetSINYPXX(strInValue(0), strInValue(1))
            If strOutput = "" Then
                MsgBox "��ȡҩƷ��Ϣʱ,�����˿�ֵ��", vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = Split(strOutput, Chr(0))(0)
            strArr = Split(strOutput, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutput = strArr(1)
        Case ����������Ա��������
'            strOutput = GETNEWRYZL(strInValue(0))
'              If strOutput = "" Then
'                MsgBox "����������Ա��������ʱ,�����˿�ֵ��", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
    End Select
    strOutPutstring = strOutput
    ҵ������_�ɶ����� = True
    DebugTool " �������Ϊ:" & strOutput
     Exit Function
    
errHand:
    DebugTool "ҵ������ʧ��"
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_�ɶ�����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
            
    '���˺�:20040923���ӵ�
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    Err = 0
    On Error GoTo errHand
    
    DebugTool "������Ժ�ǳ����ӿ�"
    
    ��Ժ�Ǽǳ���_�ɶ����� = False
    If ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "����δ����ã����ܳ�����Ժ�Ǽ�"
        Exit Function
    End If
    
    
    '��ȡסԺ��
    gstrSQL = "Select ҽ��סԺ�� סԺ�� From �����ʻ� where ����ID=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��Ժ�Ǽǳ���"
    If ҵ������_�ɶ�����(ȡ����Ժ�Ǽ�, Nvl(rsTemp!סԺ��), strOutput) = False Then Exit Function

    '����ҽ���ʻ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ɶ����� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    DebugTool "ȡ���ɹ�"
    ��Ժ�Ǽǳ���_�ɶ����� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_�ɶ�����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    Err = 0
    On Error GoTo errHand:
    
          
    '�ı䵱ǰ״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ɶ����� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_�ɶ����� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_�ɶ����� = False
End Function
Public Function ��Ժ�Ǽǳ���_�ɶ�����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '��Ժ�Ǽǳ���
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    
     '�ı䲡��״̬
     If Not ����δ�����(lng����ID, lng��ҳID) Then
            ShowMsgbox "�ò����Ѿ���Ժ������,������ȡ����Ժ!"
            Exit Function
     End If
    
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ɶ����� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_�ɶ����� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�ɶ�����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
  '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput As String
    
    Dim lng��ҳID As Long, lng�Ƿ��ϴ� As Long
    Dim dbl�����ܶ� As Double
    Dim strArr
    Dim str���㷽ʽ  As String, strסԺ�� As String
    Dim obj���� As ��������
        
    Err = 0: On Error GoTo errHand:
    Call DebugTool("����סԺ����")
    
    
    If g�������_�ɶ�����.����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣", vbInformation, gstrSysName
        Exit Function
    End If
        
    gstrSQL = "Select ��ǰ״̬,ҽ��סԺ�� סԺ�� From �����ʻ� where ����ID=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�жϵ�ǰ��סԺ״̬!"
    If Nvl(rsTemp!��ǰ״̬, 0) = 1 Then
        Err.Raise 9000, gstrSysName, "��ǰ���˻�������Ժ״̬,���Ժ���ٽ���!"
        Exit Function
    End If
    
    strסԺ�� = Nvl(rsTemp!סԺ��)
    gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    If IsNull(rsTemp("��ҳID")) = True Then
        Err.Raise 9000, gstrSysName, "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    lng��ҳID = rsTemp("��ҳID")
        

    gstrSQL = " " & _
          " Select sum(nvl(���ʽ��,0)) as ʵ�ս�� " & _
          " From סԺ���ü�¼ " & _
          " Where ��¼״̬<>0 and ����ID=" & lng����ID & " and  Nvl(���ӱ�־,0)<>9"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�ܷ���"
    dbl�����ܶ� = Nvl(rsTemp!ʵ�ս��, 0)
    If dbl�����ܶ� <> g�������_�ɶ�����.�����ܶ� Then
        Err.Raise 9000, gstrSysName, "�����ܶ��"
        Exit Function
    End If
    g�������_�ɶ�����.�����ܶ� = dbl�����ܶ�
 
    
    '�ٴν���
  
    'AZYH    PCHAR   סԺ��
    'ISPREV  PCHAR   Ԥ�����־��'0'����ʾԤ���㣩
    StrInput = strסԺ��
    StrInput = StrInput & "||1"
    If ҵ������_�ɶ�����(��Ժ����, StrInput, strOutput) = False Then Exit Function
    strArr = Split(strOutput, "||")
    
    '����ֵ
    'Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||
    'ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||
    '��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||
    'ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||
    '����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||
    '������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�||
    With obj����
        .ҽ������ = Val(strArr(4))
        .�����㸶�� = Val(strArr(6))
    End With
    
    If InsertIntoҽ�������¼(strArr, lng����ID) = False Then Exit Function
    
    
    '�����Ժ�Ǽ�
    gstrSQL = "" & _
          "   Select B.ҽ��סԺ�� סԺ��,to_Char(a.��Ժ����,'yyyy-MM-DD') as ��Ժ����" & _
          "   From ������ҳ A,�����ʻ� B " & _
          "   Where A.����iD=b.����id " & _
          "       and A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡסԺ�źͳ�Ժ����"
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "�޶�Ӧ��סԺ��Ա��Ϣ"
        Exit Function
    End If
    
    lng�Ƿ��ϴ� = 0
    If gblnδ������ Then
        'δ�����ߵ�,�轫ҽ�����ĵĲ�����Ժ�Ǽǳ���
        '��ȡסԺ��
        If ҵ������_�ɶ�����(ȡ����Ժ�Ǽ�, Nvl(rsTemp!סԺ��), strOutput) = False Then Exit Function
    '������(2006-2-20):����������ҽԺҪ���Ժ�Ǽ�(��Ժ���)�ֹ�����
    Else
        If InitInfor_�ɶ�����.�Զ���Ժ��� Then
           StrInput = Nvl(rsTemp!סԺ��)
           StrInput = StrInput & "||" & Nvl(rsTemp!��Ժ����)
           If ҵ������_�ɶ�����(��Ժ�Ǽ�, StrInput, strOutput) = False Then
              Exit Function
           Else
              lng�Ƿ��ϴ� = 1
           End If
        End If
    End If
    
    '��д�����
    Call DebugTool("��д�����¼")
    
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(),סԺ����_IN(��ҳID),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(�����㸶��),�����Ը����_IN(��),�����ʻ�֧��_IN(),"
    '   ֧��˳���_IN(סԺ��),��ҳID_IN(��ҳID),��;����_IN,��ע_IN(δ������),У��_IN(),�Ƿ��ϴ�_IN(�Ƿ��ϴ�)
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
   
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�ɶ����� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL,NULL,NULL," & lng��ҳID & ",0,0,0," & _
            dbl�����ܶ� & ",0,0," & _
            obj����.ҽ������ & "," & obj����.ҽ������ & ",0,0," & obj����.�����㸶�� & ",'" & _
            strסԺ�� & "'," & lng��ҳID & ",NULL," & IIf(gblnδ������, "'δ������'", "NULL") & ",NULL," & lng�Ƿ��ϴ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    '---------------------------------------------------------------------------------------------
    סԺ����_�ɶ����� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Private Function InsertIntoҽ�������¼(ByVal strArr As Variant, ByVal lng����ID As Long) As Boolean
    '����:���м�����ҽ�������¼
    '����:strarr��split(stroutput,"||")����������
    
    Err = 0
    On Error GoTo errHand:
    InsertIntoҽ�������¼ = False
    
    DebugTool "����InsertIntoҽ�������¼"
    'strArr:
    'Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||
    'ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||
    '��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||
    'ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||
    '����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||
    '������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�||
    
    '���̲���
    '����,����ID,
    'Ӧ֧��ͳ���,���ⶥ�Ը�,�����Ը�С��,����֧���ϼ�,ͳ��֧��ͳ���,����Ӧ��֧��,���β����㸶��,���β����������,ʵ�ʿۼ�����,ͳ��ⶥ���,ͳ���𸶽��,סԺ��¼��,���˼�¼��,
    '����,סԺ��,���ֱ��,��������,����,ҽ�ƻ�����,��Ժ����,��Ժ����,�ѽ���ͳ���,����ҽ�Ʒ�С��,����ҩƷ��,��������,�������Ʒ�,����������,�Ը�С��,�Ը�ҩƷ��,�Ը�����,�Ը����Ʒ�,
    '�Ը�������,����ҩƷ��,��������,�������Ʒ�,����������,ͳ��֧������,ͳ�����Ʒ�,ͳ��������,��Ժ��־,�����־,�����־,����ҽ��״̬,���㷽ʽ,��˷�ʽ,��������,��λ���,��λ����,
    '��ᱣ�Ϻ�,����,�Ա�,��������,Ԥ�ɽ��,����Ӧ�����,����ʵ�����,�˿���,����ʵ��֧�����,�籣������,�����������,��������־,���ջ�����,����Ա���,������ȡʱ��,ҽ�Ʊ��ձ��,
    '�籣��������,������������,�������ܱ�־,�����𸶿ۼ���־,����������,����������,�����㸶����,�������㸶��,�����㸶�����,����������ܿ�ʼ����,�㸶�ܶ�
    
    '    ����        number(2),
    gstrSQL = "ZL_ҽ�������¼_INSERT(2"
    '    ����ID      number(18),
    gstrSQL = gstrSQL & "," & lng����ID
    '    Ӧ֧��ͳ���    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(0))
    '    ���ⶥ�Ը�  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(1))
    '    �����Ը�С��    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(2))
    '    ����֧���ϼ�    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(3))
    '    ͳ��֧��ͳ���  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(4))
    '    ����Ӧ��֧��    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(5))
    '    ���β����㸶��  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(6))
    '    ���β����������    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(7))
    '    ʵ�ʿۼ�����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(8))
    '    ͳ��ⶥ���    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(9))
    '    ͳ���𸶽��    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(10))
    
    '    סԺ��¼��  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(11) & "'"
    '    ���˼�¼��  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(12) & "'"
    '    ����        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(13) & "'"
    '    סԺ��      varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(14) & "'"
    '    ���ֱ��        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(15) & "'"
    '    ��������        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(16) & "'"
    '    ����        varchar2(50),
    gstrSQL = gstrSQL & ",'" & strArr(17) & "'"
    '    ҽ�ƻ�����  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(18) & "'"
    '    ��Ժ����        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(19) & "'"
    '    ��Ժ����        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(20) & "'"
      
    '    �ѽ���ͳ���    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(21))
    '    ����ҽ�Ʒ�С��  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(22))
    '    ����ҩƷ��  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(23))
    '    ��������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(24))
    '    �������Ʒ�  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(25))
    '    ����������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(26))
    '    �Ը�С��        number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(27))
    '    �Ը�ҩƷ��  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(28))
    '    �Ը�����  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(29))
    '    �Ը����Ʒ�  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(30))
    '    �Ը�������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(31))
    '    ����ҩƷ��  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(32))
    '    ��������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(33))
    '    �������Ʒ�  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(34))
    '    ����������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(35))
    '    ͳ��֧������    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(36))
    '    ͳ�����Ʒ�  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(37))
    '    ͳ��������  number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(38))
      
    '    ��Ժ��־        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(39) & "'"
    '    �����־        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(40) & "'"
    '    �����־        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(41) & "'"
    '    ����ҽ��״̬    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(42) & "'"
    '    ���㷽ʽ        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(43) & "'"
    '    ��˷�ʽ        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(44) & "'"
    '    ��������        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(45) & "'"
    '    ��λ���        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(46) & "'"
    '    ��λ����        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(47) & "'"
    '    ��ᱣ�Ϻ�  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(48) & "'"
    '    ����        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(49) & "'"
    '    �Ա�        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(50) & "'"
    '    ��������        varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(51) & "'"
        
    '    Ԥ�ɽ��        number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(52))
    '    ����Ӧ�����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(53))
    '    ����ʵ�����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(54))
    '    �˿���        number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(55))
    '    ����ʵ��֧�����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(56))
    '    �籣������    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(57))
            
    '    �����������    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(58) & "'"
    '    ��������־    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(59) & "'"
    '    ���ջ�����  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(60) & "'"
    '    ����Ա���  varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(61) & "'"
    '    ������ȡʱ��    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(62) & "'"
    '    ҽ�Ʊ��ձ��    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(63) & "'"
    '    �籣��������    varchar2(50),
    gstrSQL = gstrSQL & ",'" & strArr(64) & "'"
    '    ������������    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(65) & "'"
    '    �������ܱ�־    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(66) & "'"
    '    �����𸶿ۼ���־    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(67) & "'"
            
    '    ����������    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(68))
    '    ����������    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(69))
    '    �����㸶����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(70))
    '    �������㸶��    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(71))
    '    �����㸶�����    number(16,5),
    gstrSQL = gstrSQL & "," & Val(strArr(72))
            
    '    ����������ܿ�ʼ����    varchar2(30),
    gstrSQL = gstrSQL & ",'" & strArr(73) & "'"
    '    �㸶�ܶ�        number(16,5))
    gstrSQL = gstrSQL & "," & Val(strArr(74)) & ")"
    gcnOracle_�ɶ�����.Execute gstrSQL, , adCmdStoredProc
    InsertIntoҽ�������¼ = True
    DebugTool "����ҽ�������¼�ɹ�"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function סԺ�������_�ɶ�����(lng����ID As Long) As Boolean
     '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim rs�����¼ As New ADODB.Recordset
    
    Dim StrInput As String, strOutput  As String
    Dim lng����ID As Long, strסԺ�� As String
    Dim strArr
    Dim lng����id1 As Long, lng����ID As Long
    Dim intMouse As Integer
    
    Err = 0: On Error GoTo errHand:
    
    
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    
    
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", TYPE_�ɶ�����, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = Nvl(rsTemp!����ID, 0)
    gstrSQL = "select * from ҽ�������¼ where ����=2  and ����ID=" & lng����ID
    Call OpenRecordset_�ɶ�����(rs�����¼, "�������")
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
        
    '�жϲ��˵�סԺ���������Ƿ��������ϡ��жϱ�׼�Ǽ�鲡�����µ�סԺ��¼������У��Ͳ��ܽ�����
    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If ��ݱ�ʶ_�ɶ�����(2, lng����id1) = "" Then
      Screen.MousePointer = intMouse
      Exit Function
    End If
    Screen.MousePointer = intMouse
    If lng����ID <> lng����id1 Then
      Err.Raise 9000, gstrSysName, "���ǵ�ǰҪ��������Ĳ���!"
      Exit Function
    End If
    

    If Nvl(rsTemp!��ע) = "δ������" Then
        'δ�����߲���ȡ����Ժ
        '�����°�����Ժ�Ǽ�
        If ��Ժ�Ǽ�_�ɶ�����(lng����ID, Nvl(rsTemp!��ҳID, 0), g�������_�ɶ�����.���Ϻ�) = False Then Exit Function
    Else
        'ȡ����Ժ
        strסԺ�� = rsTemp("֧��˳���")
        '������(2006-2-20):�ж��Ƿ����˳�Ժ�Ǽ�
        If InitInfor_�ɶ�����.�Զ���Ժ��� = False Then
            If rsTemp!�Ƿ��ϴ� = 1 Then
               Err.Raise 9000, gstrSysName, "�ò����Ѿ���Ժ���,�뵽ҽ����ȡ����Ժ���!"
               Exit Function
            End If
        Else
            If ҵ������_�ɶ�����(ȡ����Ժ�Ǽ�, strסԺ��, strOutput) = False Then Exit Function
        End If
        StrInput = strסԺ��
        If ҵ������_�ɶ�����(ȡ����Ժ����, StrInput, strOutput) = False Then
            Exit Function
        End If
    End If
    
    
    'Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�
    strArr = Split("Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�", "||")
    
    StrInput = ""
    Dim i As Integer
    For i = 0 To UBound(strArr)
        If rs�����¼.Fields(strArr(i)).Type = 131 Then
            StrInput = StrInput & "||" & (Val(Nvl(rs�����¼.Fields(strArr(i)))) * -1)
        Else
            StrInput = StrInput & "||" & Nvl(rs�����¼.Fields(strArr(i)))
        End If
    Next
    StrInput = Mid(StrInput, 3)
    strArr = Split(StrInput, "||")
    If InsertIntoҽ�������¼(strArr, lng����ID) = False Then Exit Function
    
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(),סԺ����_IN(��ҳID),����(��),�ⶥ��_IN(��),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(�����㸶��),�����Ը����_IN(��),�����ʻ�֧��_IN(),"
    '   ֧��˳���_IN(סԺ��),��ҳID_IN(��ҳID),��;����_IN,��ע_IN(),У��_IN,�Ƿ��ϴ�_IN(�Ƿ��ϴ�)

    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�ɶ����� & "," & rsTemp("����ID") & "," & Year(zlDatabase.Currentdate) & "," & _
        "NULL,NULL,NULL,NULL," & Nvl(rsTemp!��ҳID, 0) & ",0,0,0," & _
        rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & "," & Nvl(rsTemp!���Ը����, 0) * -1 & ",0," & _
        "NULL,'" & strסԺ�� & "'," & Nvl(rsTemp!��ҳID, 0) & ",NULL," & Nvl(rsTemp!��ע) & ",NULL,0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ�������¼")
    
    סԺ�������_�ɶ����� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function �����Ǽ�_�ɶ�����(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim str������¼�� As String
    Dim strArr
    Dim strסԺ�� As String
    
    Err = 0
    On Error GoTo errHand:
    
    �����Ǽ�_�ɶ����� = False
    
   '�������ŵ��ݵķ�����ϸ
    gstrSQL = "Select A.ID,A.NO,a.��ʶ�� סԺ��,A.����ID,A.��ҳID,to_char(A.����ʱ��,'yyyy-mm-dd') as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս�� " & _
              "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,M.����,Q.���� as ��������,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,�����ʻ� M,���ű� Q,������Ϣ J" & _
              "  where A.NO=[2] and A.��¼����=[3] and A.��¼״̬ = [4] And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
              "        and A.����ID=J.����ID and a.����id=m.����id and a.��������id=Q.id(+) And M.����=[1]" & _
              "        and A.�շ�ϸĿID=B.ID  " & _
              "  Order by A.����ID,A.NO,A.����ʱ��"
              
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "������ϸ�ϴ�", TYPE_�ɶ�����, str���ݺ�, lng��¼����, lng��¼״̬)
    
    With rs��ϸ
    
        If .RecordCount = 0 Then
            ShowMsgbox "û����ص���ϸ��¼,������Щ��Ŀδ����ҽ������!"
            Exit Function
        End If
        Do While Not .EOF
            gstrSQL = "Select * From ҽ��֧����Ŀ_���� where ����=[1] and ����=[2] and �շ�ϸĿid=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȷ��ҽ��֧����Ŀ", TYPE_�ɶ�����, CLng(Nvl(!����, 0)), CLng(Nvl(!�շ�ϸĿID, 0)))
            If rsTemp.EOF Then
                ShowMsgbox "��Ŀ[" & Nvl(!����) & "]δ���ж���!"
                Exit Function
            End If
            .MoveNext
        Loop
        If InitInfor_�ɶ�����.��ϸʱʵ�ϴ� = False Then
            �����Ǽ�_�ɶ����� = True
            Exit Function
        End If
        lng����ID = 0
        str������¼�� = ""
        Dim strժҪ As String
        .MoveFirst
        Do While Not .EOF
            gstrSQL = "Select * From ҽ��֧����Ŀ_���� where ����=[1] and ����=[2] and �շ�ϸĿid=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȷ��ҽ��֧����Ŀ", TYPE_�ɶ�����, CLng(Nvl(!����, 0)), CLng(Nvl(!�շ�ϸĿID, 0)))
                        
            If lng����ID <> Nvl(!����ID, 0) Then
                 '������һ�ŵ���
                 'AZYH    PCHAR   סԺ��
                 'ACFRQ   PCHAR   �������ڣ�YYYY-MM-DD��
                 'AYS PCHAR   ҽ��
                 'AKS PCHAR   ����
                 lng����ID = Nvl(!����ID, 0)
                 lng��ҳID = Nvl(!��ҳID, 0)
                 strסԺ�� = GetסԺ��(lng����ID, Nvl(!��ҳID, 0), True)
                 StrInput = strסԺ��
                 StrInput = StrInput & "||" & Nvl(!�Ǽ�ʱ��)
                 StrInput = StrInput & "||" & Nvl(!ҽ��)
                 StrInput = StrInput & "||" & Nvl(!��������)
                 If ҵ������_�ɶ�����(���Ӵ�������, StrInput, strOutput) = False Then Exit Function
                 str������¼�� = strOutput
                 
'                 '������������
'                'AZYH    PCHAR   סԺ��
'                'ACFID   PCHAR   ������¼�ţ�ͨ������ADDCFJL���ص�ֵ��
'                 strInput = Lpad(Nvl(!סԺ��, 0), 8)
'                 strInput = strInput & "||" & str������¼��
'                 If ҵ������_�ɶ�����(������������, strInput, stroutput) = False Then
'                    '��ɾ�����ŵ���
'                    Call ҵ������_�ɶ�����(ɾ���������ݼ�����ϸ, str������¼��, stroutput)
'                    Exit Function
'                 End If
            End If
            '���Ӵ�����ϸ
            'ACFID   PCHAR   ������¼��
            'AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
            'ASL PCHAR   ����(����Ϊ����)
            'ADJ PCHAR   ����
            StrInput = str������¼��
            StrInput = StrInput & "||" & Nvl(rsTemp!��Ŀ����)
            StrInput = StrInput & "||" & Nvl(!����)
            StrInput = StrInput & "||" & Nvl(!�۸�)
            
            If ҵ������_�ɶ�����(���Ӵ�����ϸ, StrInput, strOutput) = False Then
                '��ɾ�����ŵ���
                Call ҵ������_�ɶ�����(ɾ���������ݼ�����ϸ, str������¼��, strOutput)
                Exit Function
            End If
           '������ϸ��¼��||�Էѱ���||�Էѽ��
           'ժҪ����ֵ:������¼��||��ϸ��¼��||�Էѱ���||�Էѽ��||סԺ��
            strժҪ = str������¼�� & "||" & strOutput & "||" & strסԺ��
            '�����ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strժҪ & "')"
            zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            .MoveNext
        Loop
    End With
    �����Ǽ�_�ɶ����� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function Readģ������(ByVal intҵ������ As ҵ������_�ɶ�����, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--��  ��:ͨ���ù��ܶ�ȡģ������,�Ա����
    '--�����:
    '--������:
    '--��  ��:�ִ�
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim STRNAME As String
    
    If intҵ������ = ��ȡ�������� Then
        strFile = App.Path & "\������.txt"
    Else
        strFile = App.Path & "\ģ���ύ��.txt"
    End If
    
    
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    STRNAME = Get���״���(intҵ������, True)
    
    Dim blnStart As Boolean
    Dim strArr
    
    Err = 0
    On Error GoTo errHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                If intҵ������ = ��ȡ�������� Then
                    strArr = Split(strText, vbTab)
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab
                            End If
                            strArr = Split(strText, vbTab)
                            
                            If Val(strArr(0)) = 1 Then
                                str = strArr(1)
                                Exit Do
                            End If
                        Else
                             If "<" & STRNAME & ">" = strText Then
                                 blnStart = True
                             End If
                        End If
                        If "</" & STRNAME & ">" = strText Then
                            Exit Do
                        End If
                End If
            Loop
            objText.Close
            strOutPutstring = str
    End If
    If InStr(1, strOutPutstring, "@$") <> 0 Then
        strOutPutstring = Split(strOutPutstring, "@$")(1)
    End If
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function
Private Sub OpenRecordset_�ɶ�����(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_�ɶ�����, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Function סԺ�������_�ɶ�����(rsExse As Recordset, ByVal lng����ID As Long, Optional bln���ʴ� As Boolean = True) As String
    
    'rsExse:�ַ���
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim rsTemp As New ADODB.Recordset
    
    Dim lng��ҳID As Long
    Dim StrInput As String, strOutput   As String
    Dim strArr As Variant
    Dim strסԺ�� As String, str���㷽ʽ As String
    
    Dim intMouse As Integer
    
    Err = 0: On Error GoTo errHand:
    
    g�������_�ɶ�����.����ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
   
    gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    If IsNull(rsTemp("��ҳID")) = True Then
        MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    lng��ҳID = rsTemp("��ҳID")
    
    gstrSQL = "Select to_char(��Ժ����,'yyyyMM') as ��Ժ���� ,to_char(��Ժ����,'yyyyMM') as ��Ժ���� from ������ҳ where ����id=[1] and ��ҳid =[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID, lng��ҳID)
    If Nvl(rsTemp!��Ժ����) <> Nvl(rsTemp!��Ժ����) And Nvl(rsTemp!��Ժ����) <> "" Then
        '�������ͬһ�·�,����������
        '�����κ�����
        gstrSQL = "Select a.����,b.ҽ��סԺ�� as סԺ�� From ��������Ŀ¼ a,�����ʻ� b where a.���=b.���� and a.����=140 and b.����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lng����ID)
        g�������_�ɶ�����.�������� = Nvl(rsTemp!����)
        If gbln������� Then
            '���κ�����
            If ҵ������_�ɶ�����(���κ�����, "", strOutput) = False Then
                 Call ҵ������_�ɶ�����(�Ͽ��κ�����, "", strOutput)
            End If
            If ҵ������_�ɶ�����(�����κ�����, g�������_�ɶ�����.��������, strOutput) = False Then
                 Exit Function
            End If
            If ҵ������_�ɶ�����(���κ�����, "", strOutput) = False Then
                 Exit Function
            End If
        End If
        
        If ҵ������_�ɶ�����(����������Ա��������, GetסԺ��(lng����ID, lng��ҳID, True), strOutput) = False Then Exit Function
        
    End If
    
    
    gstrSQL = "Select ҽ��סԺ�� סԺ��,���� From �����ʻ� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡסԺ��"
    If rsTemp.EOF Then
        ShowMsgbox "�ò��˲���ҽ������!"
        Exit Function
    End If
    
    strסԺ�� = Nvl(rsTemp!סԺ��)
    g�������_�ɶ�����.���� = Nvl(rsTemp!����, 0)

    Screen.MousePointer = vbHourglass
    If ����סԺ��ϸ��¼(lng����ID, lng��ҳID) = False Then Exit Function
    
    With rsExse
        g�������_�ɶ�����.�����ܶ� = 0
        Do While Not .EOF
            g�������_�ɶ�����.�����ܶ� = g�������_�ɶ�����.�����ܶ� + Nvl(!���, 0)
            .MoveNext
        Loop
    End With
    
    'AZYH    PCHAR   סԺ��
    'ISPREV  PCHAR   Ԥ�����־��'0'����ʾԤ���㣩
    StrInput = strסԺ��
    StrInput = StrInput & "||0"
    If ҵ������_�ɶ�����(��Ժ����, StrInput, strOutput) = False Then Exit Function
    strArr = Split(strOutput, "||")
    
    '����ֵ
    'Ӧ֧��ͳ���||���ⶥ�Ը�||�����Ը�С��||����֧���ϼ�||ͳ��֧��ͳ���||����Ӧ��֧��||���β����㸶��||���β����������||ʵ�ʿۼ�����||ͳ��ⶥ���||
    'ͳ���𸶽��||סԺ��¼��||���˼�¼��||����||סԺ��||���ֱ��||��������||����||ҽ�ƻ�����||��Ժ����||��Ժ����||�ѽ���ͳ���||����ҽ�Ʒ�С��||����ҩƷ��||
    '��������||�������Ʒ�||����������||�Ը�С��||�Ը�ҩƷ��||�Ը�����||�Ը����Ʒ�||�Ը�������||����ҩƷ��||��������||�������Ʒ�||����������||ͳ��֧������||
    'ͳ�����Ʒ�||ͳ��������||��Ժ��־||�����־||�����־||����ҽ��״̬||���㷽ʽ||��˷�ʽ||��������||��λ���||��λ����||��ᱣ�Ϻ�||����||�Ա�||��������||Ԥ�ɽ��||
    '����Ӧ�����||����ʵ�����||�˿���||����ʵ��֧�����||�籣������||�����������||��������־||���ջ�����||����Ա���||������ȡʱ��||ҽ�Ʊ��ձ��||�籣��������||
    '������������||�������ܱ�־||�����𸶿ۼ���־||����������||����������||�����㸶����||�������㸶��||�����㸶�����||����������ܿ�ʼ����||�㸶�ܶ�||
    
    
    With g�������
    
        .ҽ������ = Val(strArr(4))
        .�����㸶�� = Val(strArr(6))
    End With
    If Format(strArr(22), "####0.00;-####0.00;0;0") <> Format(g�������_�ɶ�����.�����ܶ�, "####0.00;-####0.00;0;0") Then
        ShowMsgbox "�������ݲ���!" & vbCrLf & "ҽ�����ķ����ܶ�:" & Format(strArr(22), "####0.00;-####0.00;0;0") & vbCrLf & " ҽԺ��Ϊ:" & Format(g�������_�ɶ�����.�����ܶ�, "####0.00;-####0.00;0;0")
        If InitInfor_�ɶ�����.���ݲ��Ȳ��ɽ��� Then
            Exit Function
        End If
    End If
    
    
    
    str���㷽ʽ = "ҽ������;" & g�������.ҽ������ & ";0"
    str���㷽ʽ = str���㷽ʽ & "|�����㸶��;" & g�������.�����㸶�� & ";0"
    סԺ�������_�ɶ����� = str���㷽ʽ
    g�������_�ɶ�����.����ID = lng����ID   '��ʾ�ò����Ѿ��������������
    
    '��ӡ���㵥:�����ڽ������ܴ�ӡ
'    If InitInfor_�ɶ�����.��ӡ���㵥 Then
'        '����ӡ�ӿ�
'        '    ASTARTZYH   PCHAR   ��ӡ��ʼסԺ��
'        '    AENDZYH PCHAR   ��ӡ����סԺ��
'
'        StrInput = strסԺ�� & "||"
'        StrInput = StrInput & strסԺ��
'        Call ҵ������_�ɶ�����(��ӡ��Ժ���㱨����, StrInput, strOutput)
'    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ����סԺ��ϸ��¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '���������ϸ��¼
    Dim cnTemp As New ADODB.Connection
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput  As String, strOutput As String
    Dim strArr
    Dim strסԺ�� As String, str������¼�� As String
    
    Err = 0
    On Error GoTo errHand:
      
    
    Call DebugTool("��������")
    Set cnTemp = GetNewConnection
    cnTemp.Open
    Call DebugTool("�����ӳɹ�����ʼ�����ϸ���ݵĺϷ��ԡ�")
      
    ����סԺ��ϸ��¼ = False
    

    gstrSQL = "Select A.ID,A.��ʶ�� as סԺ��,A.NO,A.��¼����,A.��¼״̬,A.���,A.����ID,A.��ҳID,to_char(A.����ʱ��,'yyyy-mm-dd')  as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս��" & _
              "         ,M.���� as ��������,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,C.��Ŀ����,C.��ע,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ" & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,(Select * From ҽ��֧����Ŀ_���� where ����=[2] and ���� =[1]) C,������ҳ D,���ű� M" & _
              "  where A.����ID=[3] and A.��ҳID=[4] and A.���ʷ���=1 and nvl(A.ʵ�ս��,0)<>0 and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 " & _
              "        and A.��������id =M.id(+) and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=[1]" & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID(+) " & _
              "  Order by A.����ID,A.��¼����,A.No,A.��¼״̬,A.���"
              
              
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "�������", TYPE_�ɶ�����, g�������_�ɶ�����.����, lng����ID, lng��ҳID)
    
    
    
   With rs��ϸ
        Do While Not .EOF
            If Nvl(!��Ŀ����) = "" Then
                ShowMsgbox "��Ŀ[" & Nvl(!����) & "]" & Nvl(!����) & " δ����ҽ������,���ڱ�����Ŀ���������ö����ϵ!"
                Exit Function
            End If
            .MoveNext
        Loop
        
        If .RecordCount <> 0 Then .MoveFirst
        Dim strNO As String
        
        str������¼�� = ""
        strNO = ""
        Dim strժҪ As String
        
        Do While Not .EOF
            If strNO <> Nvl(!��¼����, 0) & "_" & Nvl(!NO) & "_" & Nvl(!��¼״̬, 0) Then
                strNO = Nvl(!��¼����, 0) & "_" & Nvl(!NO) & "_" & Nvl(!��¼״̬, 0)
                 '������һ�ŵ���
                 'AZYH    PCHAR   סԺ��
                 'ACFRQ   PCHAR   �������ڣ�YYYY-MM-DD��
                 'AYS PCHAR   ҽ��
                 'AKS PCHAR   ����
                 strסԺ�� = GetסԺ��(lng����ID, lng��ҳID, True)
                 StrInput = strסԺ��
                 StrInput = StrInput & "||" & Nvl(!�Ǽ�ʱ��)
                 StrInput = StrInput & "||" & Nvl(!ҽ��)
                 StrInput = StrInput & "||" & Nvl(!��������)
                 
                 If ҵ������_�ɶ�����(���Ӵ�������, StrInput, strOutput) = False Then Exit Function
                 str������¼�� = strOutput
            End If
            '���Ӵ�����ϸ
            'ACFID   PCHAR   ������¼��
            'AYPBH   PCHAR   ҩƷ���(�籣ҩƷ���)
            'ASL PCHAR   ����(����Ϊ����)
            'ADJ PCHAR   ����
            StrInput = str������¼��
            StrInput = StrInput & "||" & Nvl(!��Ŀ����)
            StrInput = StrInput & "||" & Nvl(!����)
            StrInput = StrInput & "||" & Nvl(!�۸�)
            
            If ҵ������_�ɶ�����(���Ӵ�����ϸ, StrInput, strOutput) = False Then
                '��ɾ�����ŵ���
                Call ҵ������_�ɶ�����(ɾ���������ݼ�����ϸ, str������¼��, strOutput)
                Exit Function
            End If
           '������ϸ��¼��||�Էѱ���||�Էѽ��
           'ժҪ����ֵ:������¼��||��ϸ��¼��||�Էѱ���||�Էѽ��||סԺ��
            strժҪ = str������¼�� & "||" & strOutput & "||" & strסԺ��
            '�����ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & strժҪ & "')"
             cnTemp.Execute gstrSQL, , adCmdStoredProc
            .MoveNext
        Loop
    End With
    ����סԺ��ϸ��¼ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function ҽ������_�ɶ�����() As Boolean
    ҽ������_�ɶ����� = frmSet�ɶ�����.��������
End Function
Public Sub ExecuteProcedure_�ɶ�����(ByVal strCaption As String)
'���ܣ�ִ��SQL���
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_�ɶ�����.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Public Function ����ҽ����Ժ_�ɶ�����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ����²��˵ĳ�Ժ����������������������ʱ���߻����
    Dim StrInput As String
    Dim strOutput As String
    Dim blnYes  As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    
    On Error GoTo errHandle
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "�ò����Ѿ���Ժ���޷��÷��������ܳ���������ͨ����Ժ�Ǽǽ�����Ժ����!"
        Exit Function
    End If
    
    
    gstrSQL = "Select * From סԺ���ü�¼ where nvl(�Ƿ��ϴ�,0)=1 and rownum<=1 and ����id=" & lng����ID & " and ��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�ж��Ƿ�����ϴ���¼"
        
    If Not rsTemp.EOF Then
        ShowMsgbox "�Ѿ����ϴ���������ϸ���ã��Ƿ����Ҫȡ��ҽ����Ժ?", True, blnYes
        If blnYes = False Then
            Exit Function
        End If
    End If
    
    
    '��ȡסԺ��
    gstrSQL = "Select ҽ��סԺ�� סԺ�� From �����ʻ� where ����ID=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��Ժ�Ǽǳ���"
    If ҵ������_�ɶ�����(ȡ����Ժ�Ǽ�, Nvl(rsTemp!סԺ��), strOutput) = False Then Exit Function

    '����ҽ���ʻ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ɶ����� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    
    gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
    
    DebugTool "ȡ���ɹ�"
    
    ����ҽ����Ժ_�ɶ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetסԺ��(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional bln�ʻ� As Boolean = False) As String
    '����:��ȡסԺ��
    Dim rsTemp As New ADODB.Recordset
    
    If bln�ʻ� Then
        gstrSQL = "Select ҽ��סԺ�� סԺ�� From �����ʻ� where ����ID=[2]"
    Else
        gstrSQL = "Select סԺ�� from ������Ϣ where ����id=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡסԺ��", lng����ID)
    If rsTemp.EOF Then
        ShowMsgbox "δ�ҵ���صĲ�����Ϣ!"
        Exit Function
    End If
    If Nvl(rsTemp!סԺ��) = "" Then
        ShowMsgbox "��������ص�סԺ��,����!"
        Exit Function
    End If
    GetסԺ�� = Nvl(rsTemp!סԺ��) & IIf(bln�ʻ�, "", Left(Lpad(lng��ҳID, 2, "0"), 2))
End Function

