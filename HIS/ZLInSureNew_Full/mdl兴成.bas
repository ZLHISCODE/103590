Attribute VB_Name = "mdl�˳�"
Option Explicit
Private mblnInit As Boolean     '�Ƿ��Ѿ���ʼ��
Private Const SP_STR = "|" & vbTab & "|"
Public blnmxb As Boolean      '�º���20050402��ӱ�����Ϊ���ж����Բ��Ժ��ַ�ʽ��ҽ���н���

Public Enum ҵ������_�˳�
    �˳�_���߻��������� = 0
    �˳�_���߻�����ֹͣ
    �˳�_POS������
    �˳�_POS��ֹͣ
    �˳�_��ȡ�ֿ�����Ϣ
    �˳�_JbylReadIC
    �˳�_�������Ԥ�ֽ�
    �˳�_��ͨ����֧��ȷ��
    �˳�_�����˷Ѻ���
    �˳�_סԺ����Ԥ�ֽ�
    �˳�_סԺ֧��ȷ��
End Enum
Private Type InitbaseInfor

    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ҽԺ���� As String                      '��ʼҽԺ����
    
    strPath_System    As String             'ϵͳĿ¼,
    strPath_Get       As String
    strPath_Put         As String
    strPath_In      As String
    strPath_Out     As String
    ODBC_NAME       As String
    ODBC_UserName   As String
    ODBC_PassWord   As String
    ���ڶ�����      As Boolean
    �����������    As Boolean            '�º�����20050408�޸�����
    ҽԺ����        As String
End Type

Public InitInfor_�˳� As InitbaseInfor

Private Type �������
    IC����              As String
    ��ᱣ�Ϻ�          As String
    ���֤��            As String
    ����                As String
    �Ա�                As String
    ��Ա���            As String
    �����ʻ����        As Double
    ���ձ��            As String
    ����ͳ��֧��        As Double         '����ͳ�����֧���ۼ�
    �������ۼ�          As Double       '�����ν���ۼ�
    �α�����            As String
    ͳ�ﹲ�����ۼ�      As Double       'ͳ����𹲸��ν���ۼ�
    ���Բ������ۼ�      As Double       '�������Բ������ۼ�
    ͳ��֧���ۼ�        As Double      'ͳ��֧���ۼ�
    
    '�º����޸����ӣ�����ҽ������������ؾ�ҽ��
    ��ؿ���־          As String         '1:����ؿ� ;0: ����ؿ�

    �ۼƻ�������ʻ�    As Double       '�ۼƻ�������ʻ����
    �ϴγ俨����        As String
    �ʻ�֧���ۼ�        As Double       '��������ʻ�֧���ۼ�
    ��ǰ���            As Double
    �ϴ���������        As String      '�ϴ������ҽ����
    ����סԺ����        As Long
    סԺ��������        As String       'סԺ��Ϣ��������
    ������ʶ            As String       '���Բ����߱�ʶ
    ���Բ�����          As String       '���Բ�����
    ������Ч����        As String       '���Բ���Ч�ڽ�ֹ����
    ��������            As String
    ����                As Long
    �����ܶ�            As Double
    ���ֱ���            As String
    ��Ժ���            As String
    סԺ���            As String
    ��Ժ���            As String
    ��״̬              As String
    ����ID              As Long         '��ǰ����IDֵ
    ����ID              As Long
    ���÷�������        As String
    ���ҽԺ            As String
    ���ҽԺ����        As String
    
End Type


Private Type ��������
        �ļ���¼����    As Long
        ���׷����ܶ�    As Double
        ����Ӧ���ܶ�    As Double
        �����ʻ����    As Double
        ��ҽ����Χ��    As Double
        ҽ����Χ���    As Double
        ҩƷ������    As Double
        ����ҩƷҽ����  As Double
        ����ҩƷ�Ը���  As Double
        �Է�ҩƷ���    As Double
        ����ҽ�����    As Double
        �ؼ����ν��    As Double
        �ؼ������Ը�    As Double
        ���Ʒ�ҽ����    As Double
        ����ҽ�����    As Double
        �����ҽ����    As Double
        ���ⲡ��ʶ      As String
        'סԺ���
        ͳ��֧�����    As Double
        IC����          As String
        ��Ժ���        As String
        ҽԺ����        As String
        ����            As String
        �𸶶�          As Double
        �����ν��      As Double
        �ⶥ�������Ը����  As Double
        �����ͳ��֧���ۼ�  As Double
        ����ȹ����ζ��ۼ�  As Double
        ����סԺ����        As Long
End Type
Private g�������� As ��������

Public g�������_�˳� As �������
Public gcnOracle_�˳� As ADODB.Connection     '�м������
Public gcnSQLSEVER_�˳� As ADODB.Connection     '���ӵ�ҽ�����ĵ����ݿ�

Private gbln������� As Boolean
Private gbln�Ѿ���ʼ As Boolean             '�Ѿ�����ʼ����.

'1.�������������������������ϵͳ�ϻ���½����
Private Declare Function StartPolicy Lib "XCFYFJXT.DLL" ( _
        ByVal strϵͳĿ¼ As String, ByVal strҽԺ���� As String, _
        ByVal strODBC_Name As String, ByVal ODBC�û��� As String, ByVal ODBC�û����� As String) As Long
'===============================================================================================================
'ԭ��:
'����: �������������������������ϵͳ�ϻ���½����
'��ڲ���:
'       1.ϵͳĿ¼ (����)
'       2.ҽԺ����
'       3.ODBC����Դ����
'���ڲ���: ��
'����: 0    �D�D�������߻��ɹ�
'       100              �D�DϵͳĿ¼����
'       101              �D�Dҽ�ƻ����������
'       102              �D�D�������ݿ����
'       ��11 �D�Dδ����ȷ��Ȩ
'˵��: ������������ˡ����÷ֽ⡢����֧��֮ǰ�����������߻������趨Ϊʵ������������һ�μ��ɡ�
'===============================================================================================================

'2.������ֹ������ֹͣ�������ϵͳ��½����
Private Declare Function StopPolicy Lib "XCFYFJXT.DLL" () As Long
'===============================================================================================================
'ԭ��:
'����: ����ֹͣ
'��ڲ���:��
'���ڲ���: ��
'����:  0   �D�Dֹͣ���߻��ɹ�
'       103         �D�D�Ͽ����ݿ����Ӵ���
'˵��: һ��أ��ر������߻���صĴ���֮ǰ����ô˺���
'===============================================================================================================

'3. POS��������������������POS��������������½����POS������������ʼ����
    Private Declare Function StartPos Lib "XCFYFJXT.DLL" () As Long
'===============================================================================================================
'ԭ��:
'����: ��POS��������������������POS��������������½����POS������������ʼ����
'��ڲ���:��
'���ڲ���: ��
'����: 0   �D�D����POS���ɹ�
'            ��0 �D�D����POS������
'˵��:  ���ú��� StartPolicy�ɹ�������ô˺���
'===============================================================================================================


'4. POS������ֹ������ֹͣPOS�ǼǷ���
Private Declare Function StopPos Lib "XCFYFJXT.DLL" () As Long
'===============================================================================================================
'ԭ��:
'����: ��POS������ֹ������ֹͣPOS�ǼǷ���
'��ڲ���:��
'���ڲ���: ��
'����: ��0   �D�DֹͣPOS���ɹ�
'            ��0 �D�DֹͣPOS������
'˵��:  ���ú��� StartPolicy�ɹ�������ô˺���
'===============================================================================================================

'5. ���������ڻ�ȡPOS����Ƭ״̬������ȡ�ֿ�����Ϣ
Private Declare Function GetPersonCommInfo Lib "XCFYFJXT.DLL" (ByVal str������Ϣ As String) As Long
'===============================================================================================================
'ԭ��:
'����: �����������ڻ�ȡPOS����Ƭ״̬������ȡ�ֿ�����Ϣ
'��ڲ���:��
'���ڲ���:
'       IC����|������ݺ���|����|�Ա�|ҽ�Ʋα���Ա���|�����ʻ����|���ձ��|����ͳ�����֧���ۼ�|�����ν���ۼ�
'����: ��0�D�D�ɹ�
'        8�D�D�ѽ��������
'        9�D�D�����ۼƽ����ڹ����ܶ��Ϊ���ҽ�ƿ�����Ҫû�ս�ҽ�����Ĵ���
'        ������0ֵ:                         ����
'˵��:  �����շѡ�סԺ�Ǽ�֮ǰ��Ҫ���ô˺�����
'===============================================================================================================


'6. ���������ڶ�ȡIC���е�ҽ�Ʊ�����Ϣ
'�޸Ľӿں������ƣ����ڴ˺���������Read_ic�޸�ΪJbylReadIC

Private Declare Function JbylReadIC Lib "XCFYFJXT.DLL" (ByVal str������Ϣ As String) As Long
'===============================================================================================================
'ԭ��:
'����: ���������ڶ�ȡIC���е�ҽ�Ʊ�����Ϣ��
'��ڲ���:��
'���ڲ���: ��ᱣ�Ϻ�|����|��Ա���|�α�����|ͳ����𹲸��ν���ۼ�|�������Բ������ۼ�|ͳ��֧���ۼ�|�ۼƻ�������ʻ����|�ϴγ俨����|��������ʻ�֧���ۼ�|��ǰ���|�ϴ������ҽ����|����סԺ����|סԺ��Ϣ��������|���Բ����߱�ʶ|���Բ�����|���Բ���Ч�ڽ�ֹ����
'����: 0    �D�D�ɹ�
'      ��0 �D�Dʧ��
'˵��:
'===============================================================================================================


'7. �������Ԥ�ֽ�
Private Declare Function Poli_Divide Lib "XCFYFJXT.DLL" () As Long
'===============================================================================================================
'ԭ��:
'����: ���������ڶ�������ϸ��Ϣ���з���Ԥ�ֽ⣬����������Ӧ�����

'��ڲ���:��
'           λ��ָ����Ŀ¼�е�:Poli_Divide.in�ļ�
'���ڲ���: ��
'           λ��ָ��Ŀ¼�е�:Poli_Divide.out��Poli_Divide.out.log�ļ�
'����:  0�D�D�ɹ�
'       1�D�DPoli_Divide.in�ļ���ڲ�������
'       104�D�DPoli_Divide.in��Poli_Divide.out?Poli_Divide.log�ļ��򲻿�
'       105�D�Dд���ڲ�������
'       -11�D�Dδ����ȷ��Ȩ
'       ������0�����D�D����ʧ��
'˵��:
'    1�� ���ô˺����뱣֤���߻��ѳɹ�������
'    2�� ���ô˺�����ȷ����ϵͳĿ¼���ѽ���\in��\outĿ¼��
'    ҽ�ƻ����Ĺ���ϵͳ�У����ô˺���ǰ����\inĿ¼��д��Poli_Divide.in�ļ�����������������������ڲ�����.in�ļ��ж�ȡ��
'    3�� ���óɹ����������ϵͳ��\outĿ¼��дPoli_Divide.out��Poli_Divide.log�ļ���ҽ�ƻ����Ĺ���ϵͳ��ȡPoli_Divide.out�ļ����������ϴ��ӿڱ����ݣ���׼������֧����������ڲ�����
'    4�� ���Ԥ�ֽ�ʧ�ܣ���鿴Poli_Divide.log�ļ���
'===============================================================================================================


'8. ��ͨ����֧��ȷ��
Private Declare Function Reg_Poli Lib "XCFYFJXT.DLL" (ByVal StrInput As String, ByVal strOutput As String) As Long
'===============================================================================================================
'ԭ��:
'����: ���������ڶ�������ϸ��Ϣ���з���Ԥ�ֽ⣬����������Ӧ�����

'��ڲ���:
'   ������ˮ��|IC����|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|�����ʻ�֧�����|�ֽ�֧�����
'���ڲ���:
'          ������ˮ��|IC����|�ն˻����|��������/ʱ��|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|�����ʻ�֧�����|�ֽ�֧�����|�ۼ�������ʻ����|MAC1
'����: 0�D�D�ɹ�
'       1�D�DPoli_Divide.in�ļ���ڲ�������
'       ��11�D�Dδ����ȷ��Ȩ
'       ������0�����D�D����ʧ��
'˵��:
'    1�� ���ô˺����뱣֤���߻��ѳɹ�������
'    2�� ��ڲ����ɵ���Ԥ�ֽ⺯����Ĳ�������ó�
'    3�� ����Ԥ�ֽ⺯������ô˺���
'===============================================================================================================

'9. �����˷Ѻ���
Private Declare Function PoliBackCost Lib "XCFYFJXT.DLL" (ByVal StrInput As String) As Long
'===============================================================================================================
'ԭ��:
'����: ����������������˷ѹ��ܣ����������׼�¼

'��ڲ���:
'   ������ˮ��|IC����|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|�����ʻ�֧�����|�ֽ�֧�����
'���ڲ���:
'          ������ˮ��|IC����|�ն˻����|��������/ʱ��|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|�����ʻ�֧�����|�ֽ�֧�����|�ۼ�������ʻ����|MAC1
'����: 0�D�D�ɹ�
'       ��11�D�Dδ����ȷ��Ȩ
'       ������0�����D�D����ʧ��
'˵��:
'===============================================================================================================

'10. סԺ����Ԥ�ֽ�
Private Declare Function Hosp_Divide Lib "XCFYFJXT.DLL" () As Long
'===============================================================================================================
'ԭ��:
'����: ����������������˷ѹ��ܣ����������׼�¼

'��ڲ���:��
'           λ��ָ����Ŀ¼�е�:Hosp_Divide.in�ļ�
'���ڲ���: ��
'           λ��ָ��Ŀ¼�е�:Hosp_Divide.out, Hosp_Divide_out.log�ļ�
'����:  0   �D�D�ɹ�
'       1�D�DHosp_Divide.in�ļ���ڲ�������
'       104�D�DHosp_Divide.in��Hosp_Divide.out?Hosp_Divide.log�ļ��򲻿�
'       105�D�Dд���ڲ�������
'       ��11�D�Dδ����ȷ��Ȩ
'       ������0�����D�D����ʧ��,��Щ����������ݿ⳧�̷��أ�һ��ֽ������ֽ����߲����й�
'˵��:
'    1)���ô˺����뱣֤���߻��ѳɹ�������
'    2)���ô˺�����ȷ����ϵͳĿ¼���ѽ���\in��\outĿ¼��
'    3)ҽ�ƻ����Ĺ���ϵͳ�У����ô˺���ǰ����\inĿ¼��д��Hosp_Divide.in�ļ�����������������������ڲ�����.in�ļ��ж�ȡ��
'    4)���óɹ����������ϵͳ��\outĿ¼��дHosp_Divide.out��Hosp_Divide.log�ļ���ҽ�ƻ����Ĺ���ϵͳ��ȡHosp_Divide.out�ļ����������ϴ��ӿڱ����ݣ���׼��סԺ֧����������ڲ�����
'    5) ���Ԥ�ֽ�ʧ�ܣ���鿴Hosp_Divide.log�ļ�
'===============================================================================================================



'11. סԺ֧��ȷ��
Private Declare Function Reg_Hospital Lib "XCFYFJXT.DLL" (ByVal StrInput As String, ByVal strOutput As String) As Long
'===============================================================================================================
'ԭ��:
'����: �˺����������סԺ�շѣ����������׼�¼

'��ڲ���:
'       ������ˮ��|IC����|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|ͳ��֧�����|ͳ���Ը����|�����ʻ�֧�����|�ֽ�֧�����
'���ڲ���:
'       ������ˮ��|IC����|�ն˻����|��������/ʱ��|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|ͳ��֧�����|ͳ���Ը����|�����ʻ�֧�����|�ֽ�֧�����|�ۼ�������ʻ����|MAC1|
'����:  0�D�D�ɹ�
'       1�D�DPoli_Divide.in�ļ���ڲ�������
'       ��11�D�Dδ����ȷ��Ȩ
'       ������0�����D�D����ʧ��
'˵��:
'    1�� ���ô˺����뱣֤���߻��ѳɹ�������
'    2�� ��ڲ����ɵ���Ԥ�ֽ⺯����Ĳ�������ó���
'    3�� ����סԺԤ�ֽ⺯������ô˺�����
'===============================================================================================================


Private Function Get����(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as ���� from dual "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����")
    If Not rsTemp.EOF Then
        Get���� = Int(Nvl(rsTemp!����, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function


Public Function ҽ����ʼ��_�˳�() As Boolean
    
    Dim strReg As String, strOutput As String, StrInput As String
    
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    If mblnInit = True Then
        ҽ����ʼ��_�˳� = True
        Exit Function
    End If
    
    gstrSQL = "Select * From ���ղ��� where ����=" & TYPE_�˳ɺ˹�ҵ & " and ������='ҽԺ����'"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡҽԺ����"
        
    If rsTemp.EOF Then
        ShowMsgbox "δ����ҽԺ����,���ڲ���������"
        Exit Function
    End If
    InitInfor_�˳�.ҽԺ���� = Nvl(rsTemp!����ֵ)
    
    If InitInfor_�˳�.ҽԺ���� = "" Then
        ShowMsgbox "δ����ҽԺ����,���ڲ���������"
        Exit Function
    End If
    '��ʼģ��ӿ�
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        InitInfor_�˳�.ģ������ = True
    Else
        InitInfor_�˳�.ģ������ = False
    End If
   
    InitInfor_�˳�.ҽԺ���� = gstrҽԺ����

    InitInfor_�˳�.strPath_Get = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Get"), "C:\xcyb\get")
    InitInfor_�˳�.strPath_Put = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Put"), "C:\xcyb\Put")
    InitInfor_�˳�.strPath_In = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_In"), "C:\xcyb\In")
    InitInfor_�˳�.strPath_Out = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Out"), "C:\xcyb\Out")
    InitInfor_�˳�.strPath_System = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_System"), "C:\")
    
    InitInfor_�˳�.ODBC_NAME = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_NAME"), "")
    InitInfor_�˳�.ODBC_UserName = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_USERNAME"), "")
    InitInfor_�˳�.ODBC_PassWord = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_PASSWORD"), "")
    
    
    InitInfor_�˳�.���ڶ����� = Val(GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("������"), "1")) = 1
    
    '�º�����200500408����
    
    InitInfor_�˳�.����������� = Val(GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("�����������"), "1")) = 1
        
    '�������߷���,��ز���
    '    a. ϵͳĿ¼��������[50] �D�D���ͣ�String       ֵ��һ��ָ����c:\
    '    b. ҽԺ���� [8]         �D�D���ͣ�String       ����Ϊ8λ
    '    c. ODBC����Դ����[ ]   �D�D���ͣ�String        ֵ��һ��ָ����ODBC��DSN
    '    d. ODBC�û���[ ]       �D�D���ͣ�String
    '    e. ODBC�û�����[ ]     �D�D���ͣ�String
    
    StrInput = InitInfor_�˳�.strPath_System
    StrInput = StrInput & SP_STR & InitInfor_�˳�.ҽԺ����
    StrInput = StrInput & SP_STR & InitInfor_�˳�.ODBC_NAME
    StrInput = StrInput & SP_STR & InitInfor_�˳�.ODBC_UserName
    StrInput = StrInput & SP_STR & InitInfor_�˳�.ODBC_PassWord
    
     '�º�����20050408�����޸�,Ŀ��:���סԺ�����Ƿ�����������˶�̬��
      
      If InitInfor_�˳�.����������� Then
      
        If ҵ������_�˳�(�˳�_���߻���������, StrInput, strOutput) = False Then
            Exit Function
        End If
        
      End If
      
    If InitInfor_�˳�.���ڶ����� Then
        '������POS������
        If ҵ������_�˳�(�˳�_POS������, "", "") = False Then Exit Function
    End If
    
    If Open�м��_�˳� = False Then
        Exit Function
    End If
    mblnInit = True
    ҽ����ʼ��_�˳� = True
End Function

Public Function ҽ����ֹ_�˳�() As Boolean
    
    
    '����ʼ����־��Ϊfalse
    mblnInit = False
    
 '�º�����20050408�����޸�,Ŀ��:���סԺ���ʵ�����
 
 If InitInfor_�˳�.����������� Then
 
    Call ҵ������_�˳�(�˳�_���߻�����ֹͣ, "", "")
        
    If gcnOracle_�˳�.State = 1 Then
        gcnOracle_�˳�.Close
    End If
    If gcnSQLSEVER_�˳�.State = 1 Then
        gcnSQLSEVER_�˳�.Close
    End If
    Call ҵ������_�˳�(�˳�_POS��ֹͣ, "", "")
    
    ҽ����ֹ_�˳� = True
    
  End If
  
End Function

Public Function ��ݱ�ʶ_�˳�(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    '���أ��ջ���Ϣ��
    Err = 0
    On Error GoTo errHand:
    'If bytType = 1 Or bytType = 3 or  Then Exit Function
    
    ��ݱ�ʶ_�˳� = frmIdentify�˳�.GetPatient(bytType, lng����ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_�˳� = ""
End Function

Public Function �������_�˳�(ByVal lng����ID As Long) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, TYPE_�˳ɺ˹�ҵ)
    
    If rsTemp.EOF Then
        �������_�˳� = 0
    Else
        �������_�˳� = rsTemp("�ʻ����")
    End If
End Function
Private Function WriteINParaFile(ByVal rs��ϸ As ADODB.Recordset, Optional bln����������� As Boolean = False, Optional blnסԺ As Boolean = False) As Boolean
    '����ϸ����д���ı��ļ���
    Dim str������ As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim strFile As String, StrInput As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim str��ˮ�� As String
    
    WriteINParaFile = False
    If objFile.FolderExists(InitInfor_�˳�.strPath_In) = False Then
        ShowMsgbox "δ�������ļ���(" & InitInfor_�˳�.strPath_In & "),�봴��!"
        Exit Function
    End If
    If Not blnסԺ Then
        strFile = InitInfor_�˳�.strPath_In & "\Poli_Divide.in"
    Else
        strFile = InitInfor_�˳�.strPath_In & "\Hosp_Divide.in"
    End If
    
    Err = 0: On Error GoTo errHand:
    If Not Dir(strFile) <> "" Then
        
        objFile.CreateTextFile strFile
    End If
    Set objText = objFile.OpenTextFile(strFile, ForWriting)
        
  
    If bln����������� Then
        str������ = Int(Rnd * 1000000000000#)
    Else
         rs��ϸ.MoveFirst
        str������ = Nvl(rs��ϸ!NO, 0)
    End If
    
    If blnסԺ Then
        
        '�º�����20050311�޸ģ�Hosp_Divide()��ڲ��������������Hosp_Divide.in�ļ��е�һ��û������
        
        '��һ��:
        'IC����|�ļ���¼����|���ν��׷����ܽ��|��Ժ���|ҽԺ����|����
        StrInput = g�������_�˳�.IC���� & "|"
        StrInput = StrInput & rs��ϸ.RecordCount & "|"
        StrInput = StrInput & Format(g�������_�˳�.�����ܶ� * 100, "#####0.00;-#####0.00;0;0") & "|"
        StrInput = StrInput & g�������_�˳�.��Ժ��� & "|"
        StrInput = StrInput & IIf(g�������_�˳�.��Ժ��� = "3", g�������_�˳�.���ҽԺ����, InitInfor_�˳�.ҽԺ����) & "|"
        StrInput = StrInput & g�������_�˳�.���ֱ���
    Else
        '��һ��:�ļ���¼����|���ν��׷����ܽ��
        StrInput = rs��ϸ.RecordCount & "|"
        StrInput = StrInput & Format(g�������_�˳�.�����ܶ� * 100, "#####0.00;-#####0.00;0;0")
    End If
    objText.WriteLine StrInput
    
    With rs��ϸ
        .MoveFirst
        Do While Not .EOF
            gstrSQL = "" & _
                "   Select a.*,b.����,b.���� " & _
                "   From ����֧����Ŀ a,�շ�ϸĿ B  " & _
                "   where a.�շ�ϸĿid=b.ID and  a.����=[2]" & _
                "           and a.�շ�ϸĿid=[1]"
                
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȷ��ҽ��֧����Ŀ", CLng(Nvl(!�շ�ϸĿID, 0)), TYPE_�˳ɺ˹�ҵ)
            
            '�������Ĳ���Ҫ�ж��Ƿ�ҽ����Ŀ,�º�����20050321�޸�
            
'            If rsTemp.EOF Then
'                ShowMsgbox "�շ���Ŀ��" & Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����) & "����δ����ҽ�����룬���ܽ��н���!"
'                Exit Function
'            End If
            
            If Nvl(!����, 0) - Int(Nvl(!����, 0)) > 0 Then
                ShowMsgbox "�շ���Ŀ��" & Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����) & "�����������ΪС���ˣ����ܽ��н���!"
                Exit Function
            End If
            StrInput = ""
            
            If bln����������� Then
                str��ˮ�� = Rpad(Substr(InitInfor_�˳�.ҽԺ����, 1, 8), 8, " ") & Lpad(.AbsolutePosition, 12, "0")
            Else
                str��ˮ�� = Rpad(Substr(InitInfor_�˳�.ҽԺ����, 1, 8), 8, " ") & Lpad(.AbsolutePosition, 12, "0")
            End If
            
            StrInput = str��ˮ��
            If blnסԺ Then
                '������ˮ��|��Ŀ���|��Ŀ����|��Ŀ����|��Ŀ���|����|����|�����ܽ��|���÷�������
                StrInput = StrInput & "|" & .AbsolutePosition
            Else
                '������ˮ��|������|��Ŀ����|��Ŀ����|��Ŀ���|����|����|�����ܽ��|���ִ���
                StrInput = StrInput & "|" & str������
            End If
            
            '�º�����20050321�޸�,��ҽ����Ŀ��������һ��0
            If rsTemp.EOF Then
                gstrSQL = "" & _
                "   Select b.���,b.����,b.���� " & _
                "   From �շ�ϸĿ B  " & _
                "   where  B.id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȷ����ҽ����Ŀ����", CLng(Nvl(!�շ�ϸĿID, 0)))
                StrInput = StrInput & "|0"
                StrInput = StrInput & "|" & Substr(Nvl(rsTemp!����), 1, 20)
                StrInput = StrInput & "|" & Substr(Decode(rsTemp!���, "5", 0, "6", 0, "7", 0, "J", 2, 1), 1, 1)
            Else
              StrInput = StrInput & "|" & Nvl(rsTemp!��Ŀ����)
              StrInput = StrInput & "|" & Substr(Nvl(rsTemp!����), 1, 20)
              StrInput = StrInput & "|" & Substr(Nvl(rsTemp!��ע), 1, 1)
            End If
              
              StrInput = StrInput & "|" & Format(Nvl(!����, 0) * 10000, "######;-#####;0;0")
              StrInput = StrInput & "|" & Format(Nvl(!����, 0), "######;-#####")
              StrInput = StrInput & "|" & Format(Nvl(!ʵ�ս��, 0) * 100, "######0.00;-#####0.00")
            
            
            If blnסԺ Then
                StrInput = StrInput & "|" & Format(!����ʱ��, "YYYYMMDD")
            Else
                '�º�����20050321�޸ģ�ԭ�����ڴ������Բ�
                           
                If g�������_�˳�.������ʶ = "1" And blnmxb = True Then
                    StrInput = StrInput & "|" & g�������_�˳�.���ֱ���  '���ڵ�ʱ��ѯ��������������������Դ�0000
                Else
                   StrInput = StrInput & "|0000"
                End If
                
            End If
            objText.WriteLine StrInput
            
            If bln����������� Or blnסԺ Then
            Else
                'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                 'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                 gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str��ˮ�� & "')"
                 DebugTool "     ������ϸ��־:SQL=" & gstrSQL
                 zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
                 DebugTool " ������ϸ��־:���²��˷��ü�¼�ɹ�:SQL=" & gstrSQL
            End If
            .MoveNext
        Loop
    End With
    objText.Close
    WriteINParaFile = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    objText.Close
End Function
Private Function ReadOutParaFile(ByRef obj�������� As ��������, _
    Optional bln���� As Boolean = True, Optional bln���� As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Գ����ļ������ݽ��зֽ�
    '--�����:bln����-ֻ��������Ϣ
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strFile As String, StrInput As String, strText As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim lngRow As Long
    Dim strArr As Variant
    Dim strSQL As String
    
    ReadOutParaFile = False
    If objFile.FolderExists(InitInfor_�˳�.strPath_Out) = False Then
        ShowMsgbox "δ�������ļ���(" & InitInfor_�˳�.strPath_Out & "),�봴��!"
        Exit Function
    End If
    
    If bln���� Then
        strFile = InitInfor_�˳�.strPath_Out & "\Poli_Divide.out"
    Else
        strFile = InitInfor_�˳�.strPath_Out & "\Hosp_Divide.out"
    End If
    
    Err = 0: On Error GoTo errHand:
    If Not Dir(strFile) <> "" Then
        ShowMsgbox "û�в�����صĳ����ļ���" & strFile & vbCrLf & " ����!"
        Exit Function
    End If
    Set objText = objFile.OpenTextFile(strFile, ForReading)
    
    lngRow = 1
    Do While Not objText.AtEndOfStream
          strText = Trim(objText.ReadLine)
          If strText = "" Then Exit Function
          
          strArr = Split(strText, "|")
          If lngRow = 1 Then
                
                '��������Ϣ
                '����:                                 �ļ���¼����|���ν��׷����ܽ��             |����Ӧ���ܽ��|�����ʻ����|��ҽ����Χ���|ҽ����Χ���|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|���ⲡ��ʶ
                'סԺ:   IC����|��Ժ���|ҽԺ����|����|�ļ���¼����|���ν��׷����ܽ��|ͳ��֧�����|����Ӧ���ܽ��|�����ʻ����|��ҽ����Χ���|ҽ����Χ���|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|�𸶶�|�����ν��|�ⶥ�������Ը����|�����ͳ�����֧���ۼ�|����ȹ����ν���ۼ�|����סԺ����
                If bln���� Then
                    
                    With obj��������
                        .�ļ���¼���� = Val(strArr(0))
                        .���׷����ܶ� = Format(Val(strArr(1)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����Ӧ���ܶ� = Format(Val(strArr(2)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�����ʻ���� = Format(Val(strArr(3)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .��ҽ����Χ�� = Format(Val(strArr(4)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .ҽ����Χ��� = Format(Val(strArr(5)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .ҩƷ������ = Format(Val(strArr(6)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����ҩƷҽ���� = Format(Val(strArr(7)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����ҩƷ�Ը��� = Format(Val(strArr(8)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�Է�ҩƷ��� = Format(Val(strArr(9)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����ҽ����� = Format(Val(strArr(10)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�ؼ����ν�� = Format(Val(strArr(11)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�ؼ������Ը� = Format(Val(strArr(12)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .���Ʒ�ҽ���� = Format(Val(strArr(13)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����ҽ����� = Format(Val(strArr(14)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�����ҽ���� = Format(Val(strArr(15)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .���ⲡ��ʶ = strArr(16)
                        .IC���� = ""
                        .��Ժ��� = ""
                        .ҽԺ���� = ""
                        .���� = ""
                        .ͳ��֧����� = 0
                        .�𸶶� = 0
                        .�����ν�� = 0
                        .�ⶥ�������Ը���� = 0
                        .�����ͳ��֧���ۼ� = 0
                        .����ȹ����ζ��ۼ� = 0
                        .����סԺ���� = 0
                    End With
                Else
                    'סԺ
                    With obj��������
                        .IC���� = strArr(0)
                        .��Ժ��� = strArr(1)
                        .ҽԺ���� = strArr(2)
                        .���� = strArr(3)
                        .�ļ���¼���� = Val(strArr(4))
                        .���׷����ܶ� = Format(Val(strArr(5)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .ͳ��֧����� = Format(Val(strArr(6)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����Ӧ���ܶ� = Format(Val(strArr(7)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�����ʻ���� = Format(Val(strArr(8)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .��ҽ����Χ�� = Format(Val(strArr(9)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .ҽ����Χ��� = Format(Val(strArr(10)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .ҩƷ������ = Format(Val(strArr(11)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����ҩƷҽ���� = Format(Val(strArr(12)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����ҩƷ�Ը��� = Format(Val(strArr(13)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�Է�ҩƷ��� = Format(Val(strArr(14)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����ҽ����� = Format(Val(strArr(15)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�ؼ����ν�� = Format(Val(strArr(16)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�ؼ������Ը� = Format(Val(strArr(17)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .���Ʒ�ҽ���� = Format(Val(strArr(18)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����ҽ����� = Format(Val(strArr(19)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�����ҽ���� = Format(Val(strArr(20)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�𸶶� = Format(Val(strArr(21)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�����ν�� = Format(Val(strArr(22)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�ⶥ�������Ը���� = Format(Val(strArr(23)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .�����ͳ��֧���ۼ� = Format(Val(strArr(24)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����ȹ����ζ��ۼ� = Format(Val(strArr(25)) / 100, "#####0.00;-#####0.00;0.00;0.00")
                        .����סԺ���� = Format(Val(strArr(26)), "#####0.00;-#####0.00;0.00;0.00")
                        .���ⲡ��ʶ = ""
                    End With
                End If
                If bln���� Then
                    Exit Do
                End If
        End If
        If bln���� Then
            '�������������
            '������صĻ��ܺ���ϸ����
            If InsertIntoYBK(strArr, lngRow - 1, lngRow <> 1, bln����) = False Then
                Exit Function
            End If
        Else
            'סԺ�������,������������.�����Խ���ʱ�Ŵ���.
        End If
        lngRow = lngRow + 1
    Loop
    ReadOutParaFile = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InsertIntoSQLServer_����(ByVal strȷ�ϴ� As String) As Boolean
    '�������ݵ�SQLSErver��.
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Գ����ļ������ݽ��зֽ�
    '--�����:strȷ�ϴ�-����ȷ��֧��ʱ�����Ĵ�
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strFile As String, StrInput As String, strText As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim lngRow As Long
    Dim strArr As Variant
    Dim strArr1 As Variant
    Dim strSQL As String
    Dim str�������� As String
    Dim str����ҽʦ As String
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select b.���� as ��������,������ from ������ü�¼ a,���ű� b where a.��������id=b.id(+) and ����id=" & g�������_�˳�.����ID & " and rownum=1"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������"
    If rsTemp.EOF Then
        str�������� = ""
        str����ҽʦ = ""
    Else
        str�������� = Substr(Nvl(rsTemp!��������), 1, 12)
        str����ҽʦ = Substr(Nvl(rsTemp!������), 1, 4)    '���ڵ��������ݿ�Ľ��Ϊ4λ,���ĵ�����12λ,�뵽ʱ����/
    End If
    
    InsertIntoSQLServer_���� = False
    
    If objFile.FolderExists(InitInfor_�˳�.strPath_Out) = False Then
        ShowMsgbox "δ�������ļ���(" & InitInfor_�˳�.strPath_Out & "),�봴��!"
        Exit Function
    End If
    
  '  MsgBox strȷ�ϴ�, vbOKOnly, "zlsoft"
    
    strArr1 = Split(strȷ�ϴ�, "|")
    
'    MsgBox strArr1(5), vbOKOnly, "zlsoft"
'    MsgBox strArr1(0), vbOKOnly, "zlsoft"
'
'    MsgBox strArr1(6), vbOKOnly, "zlsoft"
    
    strFile = InitInfor_�˳�.strPath_Out & "\Poli_Divide.out"
    
    Err = 0: On Error GoTo errHand:
    If Not Dir(strFile) <> "" Then
        ShowMsgbox "û�в�����صĳ����ļ���" & strFile & vbCrLf & " ����!"
        Exit Function
    End If
    Set objText = objFile.OpenTextFile(strFile, ForReading)
    
    lngRow = 1
    Do While Not objText.AtEndOfStream
          strText = Trim(objText.ReadLine)
          If strText = "" Then Exit Do
          strArr = Split(strText, "|")
          
          If lngRow = 1 Then
                
                '��������Ϣ
                '����:�ļ���¼����|���ν��׷����ܽ��|����Ӧ���ܽ��|�����ʻ����|��ҽ����Χ���|ҽ����Χ���|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|���ⲡ��ʶ
                '���ﴫ��: ������ˮ��|IC����|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|�����ʻ�֧�����|�ֽ�֧�����
                '��ṹ:vc_jyh(�����շ���ˮ��),vc_date(���÷�������),C_miid(ҽ�����),vc_cardid(����),Vc_jzh(�Һ���ˮ��),
                '    C_yljgdm(ҽ�ƻ�������),C_ksid(��������),C_doctorid(����ҽʦ),Vc_bzid(����),
                '    N_sum(�ܽ��),N_nyb(��ҽ����Χ���),N_yb(ҽ����Χ���),N_grzf(�����ʻ�֧�����),N_xzjf(�ֽ�֧�����ʻ�����֧����),
                '    N_drug_a(ҩƷ������),N_drug_byb(����ҩƷҽ�����),N_drug_bzf(����ҩƷ�Ը����),N_drug_zf(�Է�ҩƷ���),
                '    N_drug_mi(����ҽ�����),N_drug_s(�ؼ�����ҽ�����),N_drug_szf(�ؼ������Ը����),N_zlnyb(���Ʒ�ҽ�����),N_fwyb(����ҽ�����),N_fwnyb(�����ҽ�����),C_s_flag(���ⲡ��־),
                '    C_jzys (����ҽʦ), C_fscfh(��ʽ������), C_qzysdm(ǩ��ҽʦ����), C_wpcfyy(���䴦��ҽԺ), C_zt(����״̬)

                strSQL = "insert into HLD_MZJYXX(vc_jyh,vc_date,C_miid,vc_cardid,Vc_jzh,C_yljgdm,C_ksid,C_doctorid,Vc_bzid,N_sum,N_nyb,N_yb,N_grzf,N_xjzf,N_drug_a,N_drug_byb,N_drug_bzf,N_drug_zf,N_drug_mi,N_drug_s,N_drug_szf,N_zlnyb,N_fwyb,N_fwnyb,C_s_flag,C_zt) values("
                strSQL = strSQL & "'" & strArr1(0) & "',"  'vc_jyh(�����շ���ˮ��)
                strSQL = strSQL & "'" & Format(zlDatabase.Currentdate, "yyyymmddHHMMSS") & "'," 'vc_date(���÷�������)
                strSQL = strSQL & "'" & g�������_�˳�.��ᱣ�Ϻ� & "',"  'C_miid(ҽ�����)
                strSQL = strSQL & "'" & g�������_�˳�.IC���� & "',"   'vc_cardid(����)
                strSQL = strSQL & "0,"   'Vc_jzh(�Һ���ˮ��)��û��ʵ�����壬ȱʡΪ��
                strSQL = strSQL & "'" & InitInfor_�˳�.ҽԺ���� & "',"    ' C_yljgdm(ҽ�ƻ�������)
                strSQL = strSQL & "" & IIf(str�������� = "", "NULL", "'" & str�������� & "'") & ","    ' C_ksid(��������)
                strSQL = strSQL & "" & IIf(str����ҽʦ = "", "NULL", "'" & str����ҽʦ & "'") & ","    ' C_doctorid(����ҽʦ)
                
                '�º�����20051231�޸ģ�����ҽ��ǰ�ÿ��е��ֶγ��Ȳ�һ��
                If g�������_�˳�.������ʶ = "1" Then
                    strSQL = strSQL & "" & IIf(g�������_�˳�.���Բ����� = "", "Null", "'" & g�������_�˳�.���Բ����� & "'") & ","    'Vc_bzid(����)
                Else
                    strSQL = strSQL & "" & "Null" & ","    'Vc_bzid(����)
                End If
                
                strSQL = strSQL & "" & Format(Val(strArr(1)) / 100, "####0.00;-####0.00;0;0") & "," 'N_sum(�ܽ��)
                strSQL = strSQL & "" & Format(Val(strArr(4)) / 100, "####0.00;-####0.00;0;0") & "," 'N_nyb(��ҽ����Χ���)
                strSQL = strSQL & "" & Format(Val(strArr(5)) / 100, "####0.00;-####0.00;0;0") & "," 'N_yb(ҽ����Χ���)
                strSQL = strSQL & "" & Format(Val(strArr1(7)) / 100, "####0.00;-####0.00;0;0") & "," 'N_grzf(�����ʻ�֧�����),
                strSQL = strSQL & "" & Format(Val(strArr1(8)) / 100, "####0.00;-####0.00;0;0") & "," 'N_xzjf(�ֽ�֧�����ʻ�����֧����)
                strSQL = strSQL & "" & Format(Val(strArr(6)) / 100, "####0.00;-####0.00;0;0") & "," 'N_drug_a(ҩƷ������)
                strSQL = strSQL & "" & Format(Val(strArr(7)) / 100, "####0.00;-####0.00;0;0") & "," 'N_drug_byb(����ҩƷҽ�����)
                strSQL = strSQL & "" & Format(Val(strArr(8)) / 100, "####0.00;-####0.00;0;0") & "," 'N_drug_bzf(����ҩƷ�Ը����)
                strSQL = strSQL & "" & Format(Val(strArr(9)) / 100, "####0.00;-####0.00;0;0") & "," ',N_drug_zf(�Է�ҩƷ���)
                strSQL = strSQL & "" & Format(Val(strArr(10)) / 100, "####0.00;-####0.00;0;0") & "," ',N_drug_mi(����ҽ�����)
                strSQL = strSQL & "" & Format(Val(strArr(11)) / 100, "####0.00;-####0.00;0;0") & "," ',N_drug_s(�ؼ�����ҽ�����)
                strSQL = strSQL & "" & Format(Val(strArr(12)) / 100, "####0.00;-####0.00;0;0") & "," ',N_drug_szf(�ؼ������Ը����)
                strSQL = strSQL & "" & Format(Val(strArr(13)) / 100, "####0.00;-####0.00;0;0") & "," 'N_zlnyb(���Ʒ�ҽ�����)
                strSQL = strSQL & "" & Format(Val(strArr(14)) / 100, "####0.00;-####0.00;0;0") & "," 'N_fwyb(����ҽ�����)
                strSQL = strSQL & "" & Format(Val(strArr(15)) / 100, "####0.00;-####0.00;0;0") & "," 'N_fwnyb(�����ҽ�����)
                strSQL = strSQL & "'" & strArr(16) & "',"   ''C_s_flag(���ⲡ��־)
                strSQL = strSQL & "0)"   'C_zt 0 ������ 1 ����ȷ���䣬״̬δ֪ 2 ����ȷ���䣬���ɹ����
                
                'MsgBox "���������������B1��" & vbCrLf & "�������ţ�" & str�������� & vbCrLf & "��ᱣ�Ϻţ�" & g�������_�˳�.��ᱣ�Ϻ� & vbCrLf & "IC���ţ�" & g�������_�˳�.IC���� & vbCrLf & "���Բ����֣�" & g�������_�˳�.���Բ�����, vbOKOnly, "�������"
                gcnSQLSEVER_�˳�.Execute strSQL
        Else
        
        '��ϸ:������ˮ��|������|��Ŀ����|��Ŀ����|��Ŀ���|����|����|�����ܽ��|ҽ���ڷ���|ҽ�������|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|MAC2
        '��ṹ:VC_RECEIPTID(�����շ���ˮ��),N_SFSXH(�շ�˳���),VC_ITEM_ID(��Ŀ����),VC_ITEM_NAME(��Ŀ����),C_MIID(���ձ��),VC_JZH(�Һ���ˮ��),C_SSMLJB(����Ŀ¼����),C_ZLXMJB(������Ŀ����),C_FWSSFW(������ʩ��Χ),N_PRICE(����),N_sl(����),N_sum(�����ܽ��),N_DRUG_A(ҩƷ������),N_DRUG_BYB(����ҩƷҽ�����),N_drug_bzf(����ҩƷ�Ը����),N_drug_zf(�Է�ҩƷ���),D_DRUG_ZLYB(����ҽ�����),N_SYB(�ؼ�����ҽ�����),N_SYBZF(�ؼ������Ը����),N_ZLFYB(���Ʒ�ҽ�����),N_fwyb(����ҽ�����),N_fwnyb(�����ҽ�����),C_CF_FLAG(������ʶ),C_zt(����״̬)
        strSQL = "insert into HLD_MZCFZLXX(VC_RECEIPTID,N_SFSXH,VC_ITEM_ID,VC_ITEM_NAME,C_MIID,VC_JZH,C_SSMLJB,C_ZLXMJB,C_FWSSFW,N_PRICE,N_sl,N_sum,N_DRUG_A,N_DRUG_BYB,N_drug_bzf,N_drug_zf,N_DRUG_ZLYB,N_SYB,N_SYBZF,N_ZLFYB,N_fwyb,N_fwnyb,C_CF_FLAG,C_zt) values("
        strSQL = strSQL & "'" & strArr1(0) & "',"        'VC_RECEIPTID (�����շ���ˮ��)
        strSQL = strSQL & "" & lngRow - 1 & ","      'N_SFSXH (�շ�˳���)
        strSQL = strSQL & "'" & strArr(2) & "',"        'VC_ITEM_ID (��Ŀ����)
        strSQL = strSQL & "'" & Substr(strArr(3), 1, 20) & "',"      'VC_ITEM_NAME (��Ŀ����)
        strSQL = strSQL & "'" & g�������_�˳�.���ձ�� & "',"         'C_MIID (���ձ��)
        
        strSQL = strSQL & "0,"        'VC_JZH (�Һ���ˮ��)      'û��ʵ������
        Select Case strArr(4)
            Case "0"    'ҩƷ
                gstrSQL = "Select ssmljb as ���� From YB_YD where xmdm='" & strArr(2) & "'"
            Case "1"    '����
                gstrSQL = "Select tjtzbz as ���� From YB_ZLML where xmdm='" & strArr(2) & "'"
            Case Else   '����
                gstrSQL = "Select fwfw as ���� From YB_FWSS where xmdm='" & strArr(2) & "'"
        End Select
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnSQLSEVER_�˳�
        
        '�º�����20050321�޸�,��Ϊ���ڲ���������Ŀ
        
'        If rsTemp.EOF Then
'            ShowMsgbox "�ڱ�������¼ʱ,δ������ص�ҽ����Ŀ����[" & strArr(2) & "]"
'            Exit Function
'        End If
        
            Select Case strArr(4)
                 Case "0"    'ҩƷ
                    '�º�����20050321�޸�,��Ϊ���ڲ���������Ŀ

                     strSQL = strSQL & "'0',"           'C_SSMLJB (����Ŀ¼����)
                     If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (������Ŀ����)
                     Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (������Ŀ����)
                     End If

                 Case "1"    '����
                    strSQL = strSQL & "'1',"              'C_SSMLJB (����Ŀ¼����)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"           'C_ZLXMJB (������Ŀ����)
                    Else
                       strSQL = strSQL & "'0',"        'C_SSMLJB (����Ŀ¼����)
                    End If

                 Case Else   '����
                    strSQL = strSQL & "'2',"           'C_SSMLJB (����Ŀ¼����)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (������Ŀ����)
                    Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (������Ŀ����)
                    End If
                    
             End Select
             
             strSQL = strSQL & "NULL,"        'C_FWSSFW (������ʩ��Χ)
        
        
'       Select Case strArr(4)
'            Case "0"    'ҩƷ
'                '�º�����20050321�޸�,��Ϊ���ڲ���������Ŀ
'                 If rsTemp.EOF Then
'                    strSql = strSql & "3,"        'C_SSMLJB (����Ŀ¼����)
'                    strSql = strSql & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                    strSql = strSql & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                 Else
'                    strSql = strSql & "'" & Nvl(rsTemp!����) & "',"        'C_SSMLJB (����Ŀ¼����)
'                    strSql = strSql & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                    strSql = strSql & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                 End If
'
'            Case "1"    '����
'                 If rsTemp.EOF Then
'                    strSql = strSql & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                    strSql = strSql & "2,"        'C_ZLXMJB (������Ŀ����)
'                    strSql = strSql & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                 Else
'                    strSql = strSql & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                    strSql = strSql & "'" & Nvl(rsTemp!����) & "',"        'C_ZLXMJB (������Ŀ����)
'                    strSql = strSql & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                 End If
'
'            Case Else   '����
'                 If rsTemp.EOF Then
'                    strSql = strSql & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                    strSql = strSql & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                    strSql = strSql & "1,"        'C_FWSSFW (������ʩ��Χ)
'                 Else
'                    strSql = strSql & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                    strSql = strSql & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                    strSql = strSql & "'" & Nvl(rsTemp!����) & "',"        'C_FWSSFW (������ʩ��Χ)
'                 End If
'
'        End Select
        
        strSQL = strSQL & "" & Format(Val(strArr(5)) / 10000, "####0.00;-####0.00;0;0") & ","         'N_PRICE (����)
        strSQL = strSQL & "" & Format(Val(strArr(6)), "####0;-####0;0;0") & ","          'N_sl (����)
        strSQL = strSQL & "" & Format(Val(strArr(7)) / 100, "####0.00;-####0.00;0;0") & ","        'N_sum (�����ܽ��)
        strSQL = strSQL & "" & Format(Val(strArr(10)) / 100, "####0.00;-####0.00;0;0") & ","        'N_DRUG_A (ҩƷ������)
        strSQL = strSQL & "" & Format(Val(strArr(11)) / 100, "####0.00;-####0.00;0;0") & ","      'N_DRUG_BYB (����ҩƷҽ�����)
        strSQL = strSQL & "" & Format(Val(strArr(12)) / 100, "####0.00;-####0.00;0;0") & ","        'N_drug_bzf (����ҩƷ�Ը����)
        strSQL = strSQL & "" & Format(Val(strArr(13)) / 100, "####0.00;-####0.00;0;0") & ","       'N_drug_zf (�Է�ҩƷ���)
        strSQL = strSQL & "" & Format(Val(strArr(14)) / 100, "####0.00;-####0.00;0;0") & ","       'D_DRUG_ZLYB (����ҽ�����)
        strSQL = strSQL & "" & Format(Val(strArr(15)) / 100, "####0.00;-####0.00;0;0") & ","        'N_SYB (�ؼ�����ҽ�����)
        strSQL = strSQL & "" & Format(Val(strArr(16)) / 100, "####0.00;-####0.00;0;0") & ","        'N_SYBZF (�ؼ������Ը����)
        strSQL = strSQL & "" & Format(Val(strArr(17)) / 100, "####0.00;-####0.00;0;0") & ","       'N_ZLFYB (���Ʒ�ҽ�����)
        strSQL = strSQL & "" & Format(Val(strArr(18)) / 100, "####0.00;-####0.00;0;0") & ","       'N_fwyb (����ҽ�����)
        strSQL = strSQL & "" & Format(Val(strArr(19)) / 100, "####0.00;-####0.00;0;0") & ","       'N_fwnyb (�����ҽ�����)
        strSQL = strSQL & "NULL,"        'C_CF_FLAG (������ʶ)  :��ʵ������
        strSQL = strSQL & "'0')"        'C_zt (����״̬)
        
 '       MsgBox strSql, vbOKOnly, "zlsoft"
        gcnSQLSEVER_�˳�.Execute strSQL
                
        End If
        lngRow = lngRow + 1
    Loop
    InsertIntoSQLServer_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InsertIntoSQLServer_�������(ByVal lng����ID As Long) As Boolean
    '�������ݵ�SQLSErver��.
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������������ݲ���SQLServer��
    '--�����:lng����IDֵ
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim str�������� As String, str����ҽʦ As String
    Dim strSQL As String
    Dim lngRow As Long
    Dim str������ˮ�� As String
    
    
    
    InsertIntoSQLServer_������� = False
        
    Err = 0: On Error GoTo errHand:
    

    gstrSQL = "Select b.���� as ��������,������ from ������ü�¼ a,���ű� b where a.��������id=b.id(+) and ����id=" & g�������_�˳�.����ID & " and rownum=1"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������"
    If rsTemp.EOF Then
        str�������� = ""
        str����ҽʦ = ""
    Else
        str�������� = Substr(Nvl(rsTemp!��������), 1, 12)
        str����ҽʦ = Substr(Nvl(rsTemp!������), 1, 4)
    End If

    gstrSQL = "" & _
        "   Select  ����, ����id, ��Ժ���, ҽԺ����, ����, ��¼����, �����ܶ�, ͳ��֧�����, ����Ӧ���ܶ�," & _
        "           �����ʻ����, ��ҽ����Χ���, ҽ����Χ���, ҩƷ������, ����ҩƷҽ�����, ����ҩƷ�Ը����, " & _
        "           �Է�ҩƷ���, ����ҽ�����, �ؼ�����֧����, �ؼ������Ը���, ���Ʒ�ҽ�����, ����ҽ�����, �����ҽ�����," & _
        "           ���ⲡ��ʶ, �𸶶�, �����ν��, �ⶥ���Ը����, ���ͳ������ۼ�, ��ȹ������ۼ�, ����סԺ���� " & _
        "   from ҽ�������¼ " & _
        "   Where ����=1 and ����ID =" & g�������_�˳�.����ID
        
    OpenRecordset_�˳� rsData, "��ȡҽ�������¼"
    
    gstrSQL = "Select * From ���ս����¼ where ����=1 and ��¼id=" & g�������_�˳�.����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"
    str������ˮ�� = Nvl(rsTemp!֧��˳���)
            
    '��������Ϣ
    '����:�ļ���¼����|���ν��׷����ܽ��|����Ӧ���ܽ��|�����ʻ����|��ҽ����Χ���|ҽ����Χ���|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|���ⲡ��ʶ
    '���ﴫ��: ������ˮ��|IC����|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|�����ʻ�֧�����|�ֽ�֧�����
    '��ṹ:vc_jyh(�����շ���ˮ��),vc_date(���÷�������),C_miid(ҽ�����),vc_cardid(����),Vc_jzh(�Һ���ˮ��),
    '    C_yljgdm(ҽ�ƻ�������),C_ksid(��������),C_doctorid(����ҽʦ),Vc_bzid(����),
    '    N_sum(�ܽ��),N_nyb(��ҽ����Χ���),N_yb(ҽ����Χ���),N_grzf(�����ʻ�֧�����),N_xzjf(�ֽ�֧�����ʻ�����֧����),
    '    N_drug_a(ҩƷ������),N_drug_byb(����ҩƷҽ�����),N_drug_bzf(����ҩƷ�Ը����),N_drug_zf(�Է�ҩƷ���),
    '    N_drug_mi(����ҽ�����),N_drug_s(�ؼ�����ҽ�����),N_drug_szf(�ؼ������Ը����),N_zlnyb(���Ʒ�ҽ�����),N_fwyb(����ҽ�����),N_fwnyb(�����ҽ�����),C_s_flag(���ⲡ��־),
    '    C_jzys (����ҽʦ), C_fscfh(��ʽ������), C_qzysdm(ǩ��ҽʦ����), C_wpcfyy(���䴦��ҽԺ), C_zt(����״̬)

    strSQL = "insert into HLD_MZJYXX(vc_jyh,vc_date,C_miid,vc_cardid,Vc_jzh,C_yljgdm,C_ksid,C_doctorid,Vc_bzid,N_sum,N_nyb,N_yb,N_grzf,N_xjzf,N_drug_a,N_drug_byb,N_drug_bzf,N_drug_zf,N_drug_mi,N_drug_s,N_drug_szf,N_zlnyb,N_fwyb,N_fwnyb,C_s_flag,C_zt) values("
    
    '20051128�޸ģ��º��ã�����ҽ�����Ķ����ϵ�����ˮ�Ź涨Ϊ ҽԺ���루8λ��������������ˮ�ţ�ԭ���ʵ���ˮ�ţ�
    'strSql = strSql & "'" & Rpad(Substr(InitInfor_�˳�.ҽԺ����, 1, 8), 8, " ") & Lpad(Substr(lng����ID, 1, 12), 12, "0") & "',"     'vc_jyh(�����շ���ˮ��)
    
    strSQL = strSQL & "'" & Rpad(Substr(InitInfor_�˳�.ҽԺ����, 1, 8), 8, " ") & "-" & Lpad(Substr(str������ˮ��, 10, 11), 11, "0") & "',"     'vc_jyh(�����շ���ˮ��)
    strSQL = strSQL & "'" & Format(zlDatabase.Currentdate, "yyyymmddHHMMSS") & "'," 'vc_date(���÷�������)
    strSQL = strSQL & "'" & g�������_�˳�.��ᱣ�Ϻ� & "',"  'C_miid(ҽ�����)
    strSQL = strSQL & "'" & g�������_�˳�.IC���� & "',"   'vc_cardid(����)
    strSQL = strSQL & "0,"   'Vc_jzh(�Һ���ˮ��)��û��ʵ�����壬ȱʡΪ��
    strSQL = strSQL & "'" & InitInfor_�˳�.ҽԺ���� & "',"    ' C_yljgdm(ҽ�ƻ�������)
    strSQL = strSQL & "" & IIf(str�������� = "", "NULL", "'" & str�������� & "'") & ","    ' C_ksid(��������)
    strSQL = strSQL & "" & IIf(str����ҽʦ = "", "NULL", "'" & str����ҽʦ & "'") & ","    ' C_doctorid(����ҽʦ)
    
    '�º�����20051231�޸ģ�����ҽ��ǰ�ÿ��е��ֶγ��Ȳ�һ��
    If g�������_�˳�.������ʶ = "1" Then
        strSQL = strSQL & "" & IIf(g�������_�˳�.���Բ����� = "", "Null", "'" & g�������_�˳�.���Բ����� & "'") & ","    'Vc_bzid(����)
    Else
        strSQL = strSQL & "" & "Null" & ","    'Vc_bzid(����)
    End If
    
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!�����ܶ�, 0), "####0.00;-####0.00;0;0") & "," 'N_sum(�ܽ��)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!��ҽ����Χ���, 0), "####0.00;-####0.00;0;0") & "," 'N_nyb(��ҽ����Χ���)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!ҽ����Χ���, 0), "####0.00;-####0.00;0;0") & "," 'N_yb(ҽ����Χ���)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsTemp!�����ʻ�֧��, 0), "####0.00;-####0.00;0;0") & "," 'N_grzf(�����ʻ�֧�����),
    strSQL = strSQL & "" & Format(-1 * Nvl(rsTemp!ȫ�Ը����, 0), "####0.00;-####0.00;0;0") & "," 'N_xzjf(�ֽ�֧�����ʻ�����֧����)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!ҩƷ������, 0), "####0.00;-####0.00;0;0") & "," 'N_drug_a(ҩƷ������)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!����ҩƷҽ�����, 0), "####0.00;-####0.00;0;0") & "," 'N_drug_byb(����ҩƷҽ�����)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!����ҩƷ�Ը����, 0), "####0.00;-####0.00;0;0") & "," 'N_drug_bzf(����ҩƷ�Ը����)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!�Է�ҩƷ���, 0), "####0.00;-####0.00;0;0") & "," ',N_drug_zf(�Է�ҩƷ���)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!����ҽ�����, 0), "####0.00;-####0.00;0;0") & "," ',N_drug_mi(����ҽ�����)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!�ؼ�����֧����, 0), "####0.00;-####0.00;0;0") & "," ',N_drug_s(�ؼ�����ҽ�����)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!�ؼ������Ը���, 0), "####0.00;-####0.00;0;0") & "," ',N_drug_szf(�ؼ������Ը����)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!���Ʒ�ҽ�����, 0), "####0.00;-####0.00;0;0") & "," 'N_zlnyb(���Ʒ�ҽ�����)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!����ҽ�����, 0), "####0.00;-####0.00;0;0") & "," 'N_fwyb(����ҽ�����)
    strSQL = strSQL & "" & Format(-1 * Nvl(rsData!�����ҽ�����, 0), "####0.00;-####0.00;0;0") & "," 'N_fwnyb(�����ҽ�����)
    strSQL = strSQL & "'" & Nvl(rsData!���ⲡ��ʶ, 0) & "',"  ''C_s_flag(���ⲡ��־)
    strSQL = strSQL & "0)"   'C_zt 0 ������ 1 ����ȷ���䣬״̬δ֪ 2 ����ȷ���䣬���ɹ����
    
    gcnSQLSEVER_�˳�.Execute strSQL
                    
  
    gstrSQL = "select ����id, ����, ������ˮ��, ������, ��Ŀ����, ��Ŀ����, ��Ŀ���, ����, ����,  " & _
             "        �����ܽ��, ҽ���ڷ���, ҽ�������, ҩƷ������, ����ҽ�����, �����Ը����,  " & _
             "        �Է�ҩƷ���, ����ҽ�����, �ؼ����ν��, �ؼ������Ը�, ���Ʒ�ҽ����, ����ҽ�����,  " & _
             "        �����ҽ����, ���÷�������, mac2  " & _
             " from ҽ��������ϸ��¼" & _
             " where ����=1 and ����id=" & lng����ID   '�º�����20050315����޸�;��Ϊ�������ŵ���
             
             
    
    Call OpenRecordset_�˳�(rsData, "��ȡ������ϸ��¼", gstrSQL)
    
    lngRow = 1
    Do While Not rsData.EOF
        
        '��ϸ:������ˮ��|������|��Ŀ����|��Ŀ����|��Ŀ���|����|����|�����ܽ��|ҽ���ڷ���|ҽ�������|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|MAC2
        '��ṹ:VC_RECEIPTID(�����շ���ˮ��),N_SFSXH(�շ�˳���),VC_ITEM_ID(��Ŀ����),VC_ITEM_NAME(��Ŀ����),C_MIID(���ձ��),VC_JZH(�Һ���ˮ��),C_SSMLJB(����Ŀ¼����),C_ZLXMJB(������Ŀ����),C_FWSSFW(������ʩ��Χ),N_PRICE(����),N_sl(����),N_sum(�����ܽ��),N_DRUG_A(ҩƷ������),N_DRUG_BYB(����ҩƷҽ�����),N_drug_bzf(����ҩƷ�Ը����),N_drug_zf(�Է�ҩƷ���),D_DRUG_ZLYB(����ҽ�����),N_SYB(�ؼ�����ҽ�����),N_SYBZF(�ؼ������Ը����),N_ZLFYB(���Ʒ�ҽ�����),N_fwyb(����ҽ�����),N_fwnyb(�����ҽ�����),C_CF_FLAG(������ʶ),C_zt(����״̬)
        
        strSQL = "insert into HLD_MZCFZLXX(VC_RECEIPTID,N_SFSXH,VC_ITEM_ID,VC_ITEM_NAME,C_MIID,VC_JZH,C_SSMLJB,C_ZLXMJB,C_FWSSFW,N_PRICE,N_sl,N_sum,N_DRUG_A,N_DRUG_BYB,N_drug_bzf,N_drug_zf,N_DRUG_ZLYB,N_SYB,N_SYBZF,N_ZLFYB,N_fwyb,N_fwnyb,C_CF_FLAG,C_zt) values("
       
       '20051128�޸ģ��º��ã�����ҽ�����Ķ����ϵ�����ˮ�Ź涨Ϊ ҽԺ���루8λ��������������ˮ�ţ�ԭ���ʵ���ˮ�ţ�
        'strSql = strSql & "'" & Rpad(Substr(InitInfor_�˳�.ҽԺ����, 1, 8), 8, " ") & Lpad(Substr(lng����ID, 1, 12), 12, "0") & "',"        'VC_RECEIPTID (�����շ���ˮ��)
        
        strSQL = strSQL & "'" & Rpad(Substr(InitInfor_�˳�.ҽԺ����, 1, 8), 8, " ") & "-" & Lpad(Substr(str������ˮ��, 10, 11), 11, "0") & "',"        'VC_RECEIPTID (�����շ���ˮ��)
        strSQL = strSQL & "" & lngRow & ","        'N_SFSXH (�շ�˳���)
        strSQL = strSQL & "'" & Nvl(rsData!��Ŀ����) & "',"        'VC_ITEM_ID (��Ŀ����)
        strSQL = strSQL & "'" & Nvl(rsData!��Ŀ����) & "',"        'VC_ITEM_NAME (��Ŀ����)
        strSQL = strSQL & "'" & g�������_�˳�.���ձ�� & "',"            'C_MIID (���ձ��)
        
        strSQL = strSQL & "0,"        'VC_JZH (�Һ���ˮ��)      'û��ʵ������
        Select Case Nvl(rsData!��Ŀ���, "0")
            Case "0"    'ҩƷ
                gstrSQL = "Select ssmljb as ���� From YB_YD where xmdm='" & Nvl(rsData!��Ŀ����, "0") & "'"
            Case "1"    '����
                gstrSQL = "Select tjtzbz as ���� From YB_ZLML where xmdm='" & Nvl(rsData!��Ŀ����, "0") & "'"
            Case Else   '����
                gstrSQL = "Select fwfw as ���� From YB_FWSS where xmdm='" & Nvl(rsData!��Ŀ����, "0") & "'"
        End Select
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnSQLSEVER_�˳�
        
'        �º�����20050321�޸�ע�͵�
        
'        If rsTemp.EOF Then
'            ShowMsgbox "�ڱ�������¼ʱ,δ������ص�ҽ����Ŀ����[" & Nvl(rsData!��Ŀ����) & "]"
'            Exit Function
'        End If
        
       Select Case Nvl(rsData!��Ŀ���, "0")

             Case "0"    'ҩƷ
                    '�º�����20050321�޸�,��Ϊ���ڲ���������Ŀ

                     strSQL = strSQL & "'0',"           'C_SSMLJB (����Ŀ¼����)
                     If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (������Ŀ����)
                     Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (������Ŀ����)
                     End If

                 Case "1"    '����
                    strSQL = strSQL & "'1',"              'C_SSMLJB (����Ŀ¼����)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"           'C_ZLXMJB (������Ŀ����)
                    Else
                       strSQL = strSQL & "'0',"        'C_SSMLJB (����Ŀ¼����)
                    End If

                 Case Else   '����
                    strSQL = strSQL & "'2',"           'C_SSMLJB (����Ŀ¼����)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (������Ŀ����)
                    Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (������Ŀ����)
                    End If
                    
      End Select
      
      strSQL = strSQL & "NULL,"        'C_FWSSFW (������ʩ��Χ)
      
'            Case "0"    'ҩƷ
'
'                '�º�����20050321�޸�,��Ϊ���ڲ���������Ŀ
'                 If rsTemp.EOF Then
'                    strSQL = strSQL & "3,"        'C_SSMLJB (����Ŀ¼����)
'                    strSQL = strSQL & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                    strSQL = strSQL & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                 Else
'                    strSQL = strSQL & "'" & Nvl(rsTemp!����) & "',"        'C_SSMLJB (����Ŀ¼����)
'                    strSQL = strSQL & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                    strSQL = strSQL & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                 End If
'
'            Case "1"    '����
'                  If rsTemp.EOF Then
'                    strSQL = strSQL & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                    strSQL = strSQL & "2,"        'C_ZLXMJB (������Ŀ����)
'                    strSQL = strSQL & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                 Else
'                    strSQL = strSQL & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                    strSQL = strSQL & "'" & Nvl(rsTemp!����) & "',"        'C_ZLXMJB (������Ŀ����)
'                    strSQL = strSQL & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                 End If
'
'            Case Else   '����
'                 If rsTemp.EOF Then
'                    strSQL = strSQL & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                    strSQL = strSQL & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                    strSQL = strSQL & "1,"        'C_FWSSFW (������ʩ��Χ)
'                 Else
'                    strSQL = strSQL & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                    strSQL = strSQL & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                    strSQL = strSQL & "'" & Nvl(rsTemp!����) & "',"        'C_FWSSFW (������ʩ��Χ)
'                 End If
         
        strSQL = strSQL & "" & Format(Nvl(rsData!����, 0), "####0.0000;-####0.0000;0;0") & ","        'N_PRICE (����)
        strSQL = strSQL & "" & Format(Nvl(rsData!����, 0), "####0;-####0;0;0") & ","          'N_sl (����)
        strSQL = strSQL & "" & Format(Nvl(rsData!�����ܽ��, 0), "####0.00;-####0.00;0;0") & ","        'N_sum (�����ܽ��)
        strSQL = strSQL & "" & Format(Nvl(rsData!ҩƷ������, 0), "####0.00;-####0.00;0;0") & ","        'N_DRUG_A (ҩƷ������)
        strSQL = strSQL & "" & Format(Nvl(rsData!����ҽ�����, 0), "####0.00;-####0.00;0;0") & ","      'N_DRUG_BYB (����ҩƷҽ�����)
        strSQL = strSQL & "" & Format(Nvl(rsData!�����Ը����, 0), "####0.00;-####0.00;0;0") & ","        'N_drug_bzf (����ҩƷ�Ը����)
        strSQL = strSQL & "" & Format(Nvl(rsData!�Է�ҩƷ���, 0), "####0.00;-####0.00;0;0") & ","       'N_drug_zf (�Է�ҩƷ���)
        strSQL = strSQL & "" & Format(Nvl(rsData!����ҽ�����, 0), "####0.00;-####0.00;0;0") & ","       'D_DRUG_ZLYB (����ҽ�����)
        strSQL = strSQL & "" & Format(Nvl(rsData!�ؼ����ν��, 0), "####0.00;-####0.00;0;0") & ","        'N_SYB (�ؼ�����ҽ�����)
        strSQL = strSQL & "" & Format(Nvl(rsData!�ؼ������Ը�, 0), "####0.00;-####0.00;0;0") & ","        'N_SYBZF (�ؼ������Ը����)
        strSQL = strSQL & "" & Format(Nvl(rsData!���Ʒ�ҽ����, 0), "####0.00;-####0.00;0;0") & ","       'N_ZLFYB (���Ʒ�ҽ�����)
        strSQL = strSQL & "" & Format(Nvl(rsData!����ҽ�����, 0), "####0.00;-####0.00;0;0") & ","       'N_fwyb (����ҽ�����)
        strSQL = strSQL & "" & Format(Nvl(rsData!�����ҽ����, 0), "####0.00;-####0.00;0;0") & ","       'N_fwnyb (�����ҽ�����)
        strSQL = strSQL & "NULL,"        'C_CF_FLAG (������ʶ)  :��ʵ������
        strSQL = strSQL & "'0')"        'C_zt (����״̬)
        gcnSQLSEVER_�˳�.Execute strSQL
        rsData.MoveNext
        
        '�º�����20050402�޸����
        
        lngRow = lngRow + 1
        
    Loop
    InsertIntoSQLServer_������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function InsertIntoData_סԺ(ByVal strȷ�ϴ� As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '�������ݵ�SQLSErver��.
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Գ����ļ������ݽ��зֽ�
    '--�����:strȷ�ϴ�-����ȷ��֧��ʱ�����Ĵ�
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strFile As String, StrInput As String, strText As String
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim lngRow As Long
    Dim strArr As Variant
    Dim strArr1 As Variant
    Dim strSQL As String
    Dim str�������� As String
    Dim str����ҽʦ As String
    Dim str����ʱ�� As String, str��Ժʱ�� As String, str��Ժ���� As String, str������� As String
    Dim lngסԺ����  As Long
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand:
    
    gstrSQL = "Select to_char(�շ�ʱ��,'yyyyMMDD') as ����ʱ�� from ���˽��ʼ�¼ where ID=" & g�������_�˳�.����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"
    
    str����ʱ�� = Nvl(rsTemp!����ʱ��)
    gstrSQL = "" & _
        "   Select to_char(a.��Ժ����,'yyyyMMDD') as ��Ժ����,to_char(a.��Ժ����,'yyyyMMDD') as ��Ժ����," & _
        "           decode(trunc(a.��Ժ����)-trunc(a.��Ժ����),0,1,trunc(a.��Ժ����)-trunc(a.��Ժ����)) as ����,a.��Ժ��ʽ,a.סԺҽʦ,b.���� as סԺ����" & _
        "   from ������ҳ a,���ű� b " & _
        "   where a.��Ժ����ID=b.id(+) and a.����id=" & lng����ID & " and a.��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsData, gstrSQL, "��ȡ������Ϣ"
    
    str��Ժʱ�� = Nvl(rsData!��Ժ����): str��Ժ���� = Nvl(rsData!��Ժ����): lngסԺ���� = Nvl(rsData!����, 0)
    '1����2 ��ת3 δ��4 ת��5 ����

    Select Case Nvl(rsData!��Ժ��ʽ)
      Case "����"
          str������� = 1
      Case "��ת"
          str������� = 2
      Case "δ��"
          str������� = 3
      Case "����"
          str������� = 5
      Case "תԺ"
          str������� = 4
      Case Else
          str������� = 1
      End Select
        
    gstrSQL = "Select b.���� as ��������,������ from סԺ���ü�¼ a,���ű� b where a.��������id=b.id(+) and ����id=" & g�������_�˳�.����ID & " and rownum=1"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������"
    If rsTemp.EOF Then
        str�������� = ""
        str����ҽʦ = ""
    Else
        str�������� = Substr(Nvl(rsTemp!��������), 1, 12)
        str����ҽʦ = Substr(Nvl(rsTemp!������), 1, 4)    '���ڵ��������ݿ�Ľ��Ϊ4λ,���ĵ�����12λ,�뵽ʱ����/
    End If
    
    InsertIntoData_סԺ = False
    
    If objFile.FolderExists(InitInfor_�˳�.strPath_Out) = False Then
        ShowMsgbox "δ�������ļ���(" & InitInfor_�˳�.strPath_Out & "),�봴��!"
        Exit Function
    End If
    strArr1 = Split(strȷ�ϴ�, "|")
    strFile = InitInfor_�˳�.strPath_Out & "\Hosp_Divide.out"
    
    Err = 0: On Error GoTo errHand:
    If Not Dir(strFile) <> "" Then
        ShowMsgbox "û�в�����صĳ����ļ���" & strFile & vbCrLf & " ����!"
        Exit Function
    End If
    Set objText = objFile.OpenTextFile(strFile, ForReading)
    
    lngRow = 1
    Do While Not objText.AtEndOfStream
          strText = Trim(objText.ReadLine)
          If strText = "" Then Exit Do
          strArr = Split(strText, "|")
          
          If lngRow = 1 Then
                
                '��������Ϣ
                'סԺ:   IC����|��Ժ���|ҽԺ����|����|�ļ���¼����|���ν��׷����ܽ��|ͳ��֧�����|����Ӧ���ܽ��|�����ʻ����|��ҽ����Χ���|ҽ����Χ���|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|�𸶶�|�����ν��|�ⶥ�������Ը����|�����ͳ�����֧���ۼ�|����ȹ����ν���ۼ�|����סԺ����
                'סԺת��:������ˮ��|IC����|�ն˻����|��������/ʱ��|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|ͳ��֧�����|ͳ���Ը����||�����ʻ�֧�����|�ֽ�֧�����|�ۼ�������ʻ����|MAC1
                '��ṹ:Field1(סԺ������ˮ��),Field2(סԺ��),Field3(���ձ��),Field4(ҽ�ƻ�������),Field5(סԺ���),Field6(��������),Field7(��Ժ����),Field8(��Ժ����),Field13(��Ա���),Field14(סԺ����),Field15(����סԺ����),Field16(����ͳ�����֧���ۼ�),Field17(�����θ���֧���ۼ�),Field18(��Ժ���),Field19(��Ժ���),Field20(�������),Field21(����),Field22(�ܽ��),Field23(ҩƷ������),Field24(����ҩƷҽ�����),Field25(����ҩƷ�Ը����),Field26(�Է�ҩƷ���),Field27(����ҽ�����),Field28(�ؼ�����ҽ�����),Field30(���Ʒ�ҽ�����),Field31(����ҽ�����),Field32(�����ҽ�����),Field33(�𸶶�),Field34(��ҽ����Χ���),Field35(�ⶥ�������Ը����),Field36(�����ʻ�֧�����),Field37(סԺ����),Field38(����ҽʦ),Field39(ҽԺ����Ա),Field29(�ؼ������Ը����),Field40(ͳ��֧�����),Field41(�����ν��),C_kzt38(��״̬),C_zt39(����״̬),ydyymc(���ҽԺ����),ydyyjb(���ҽԺ����)
'
'                strSql = "insert into  HLD_HOSP_JIESUAN(Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field13,Field14,Field15,Field16,Field17,Field18,Field19,Field20,Field21,Field22,Field23,Field24,Field25,Field26,Field27,Field28,Field30,Field31,Field32,Field33,Field34,Field35,Field36,Field37,Field38,Field39,Field29,Field40,Field41,C_kzt,C_zt,ydyymc,ydyyjb) values("

                strSQL = "insert into  HLD_HOSP_JIESUAN(Field1,Field2,Field3,Field4,Field5,Field6,Field7,Field8,Field13,Field14,Field15,Field16,Field17,Field18,Field19,Field20,Field21,Field22,Field23,Field24,Field25,Field26,Field27,Field28,Field30,Field31,Field32,Field33,Field34,Field35,Field36,Field37,Field38,Field39,Field29,Field40,Field41,C_kzt,C_zt) values("
                strSQL = strSQL & "'" & strArr1(0) & "',"   'Field1(סԺ������ˮ��),
                strSQL = strSQL & "'" & lng����ID & "_" & lng��ҳID & "',"   'Field2(סԺ��),
                strSQL = strSQL & "'" & g�������_�˳�.���ձ�� & "',"   'Field3(���ձ��),
                strSQL = strSQL & "'" & InitInfor_�˳�.ҽԺ���� & "',"   'Field4(ҽ�ƻ�������),
                strSQL = strSQL & "'" & g�������_�˳�.סԺ��� & "',"  'Field5(סԺ���),
                strSQL = strSQL & "'" & str����ʱ�� & "'," 'Field6(��������),
                strSQL = strSQL & "'" & str��Ժʱ�� & "'," 'Field7(��Ժ����),
                strSQL = strSQL & "'" & str��Ժ���� & "'," 'Field8(��Ժ����),
                strSQL = strSQL & "'" & g�������_�˳�.��Ա��� & "',"  'Field13(��Ա���),
                strSQL = strSQL & "" & lngסԺ���� & "," 'Field14(סԺ����),
                strSQL = strSQL & "'" & g�������_�˳�.����סԺ���� + 1 & "'," 'Field15(����סԺ����),
                strSQL = strSQL & "" & Format(Val(strArr(24)) / 100, "####0.00;-####0.00;0;0") & ","  'Field16(����ͳ�����֧���ۼ�),
                strSQL = strSQL & "" & Format(Val(strArr(25)) / 100, "####0.00;-####0.00;0;0") & ","  'Field17(�����θ���֧���ۼ�),
                strSQL = strSQL & "'" & g�������_�˳�.��Ժ��� & "',"  'Field18(��Ժ���),
                strSQL = strSQL & "'" & g�������_�˳�.��Ժ��� & "',"  'Field19(��Ժ���),
                strSQL = strSQL & "'" & str������� & "',"   'Field20(�������),
                strSQL = strSQL & "'" & g�������_�˳�.���ֱ��� & "',"  'Field21(����),
                strSQL = strSQL & "" & Format(Val(strArr(5)) / 100, "####0.00;-####0.00;0;0") & ","  'Field22(�ܽ��),
                strSQL = strSQL & "" & Format(Val(strArr(11)) / 100, "####0.00;-####0.00;0;0") & "," 'Field23(ҩƷ������),
                strSQL = strSQL & "" & Format(Val(strArr(12)) / 100, "####0.00;-####0.00;0;0") & "," 'Field24(����ҩƷҽ�����),
                strSQL = strSQL & "" & Format(Val(strArr(13)) / 100, "####0.00;-####0.00;0;0") & "," 'Field25(����ҩƷ�Ը����),
                strSQL = strSQL & "" & Format(Val(strArr(14)) / 100, "####0.00;-####0.00;0;0") & "," 'Field26(�Է�ҩƷ���),
                strSQL = strSQL & "" & Format(Val(strArr(15)) / 100, "####0.00;-####0.00;0;0") & "," 'Field27(����ҽ�����),
                strSQL = strSQL & "" & Format(Val(strArr(16)) / 100, "####0.00;-####0.00;0;0") & "," 'Field28(�ؼ�����ҽ�����),
                strSQL = strSQL & "" & Format(Val(strArr(18)) / 100, "####0.00;-####0.00;0;0") & "," 'Field30(���Ʒ�ҽ�����),
                strSQL = strSQL & "" & Format(Val(strArr(19)) / 100, "####0.00;-####0.00;0;0") & "," 'Field31(����ҽ�����),
                strSQL = strSQL & "" & Format(Val(strArr(20)) / 100, "####0.00;-####0.00;0;0") & "," 'Field32(�����ҽ�����),
                strSQL = strSQL & "" & Format(Val(strArr(21)) / 100, "####0.00;-####0.00;0;0") & "," 'Field33(�𸶶�),
                strSQL = strSQL & "" & Format(Val(strArr(9)) / 100, "####0.00;-####0.00;0;0") & "," 'Field34(��ҽ����Χ���),
                strSQL = strSQL & "" & Format(Val(strArr(23)) / 100, "####0.00;-####0.00;0;0") & "," 'Field35(�ⶥ�������Ը����),
                strSQL = strSQL & "" & Format(Val(strArr1(9)) / 100, "####0.00;-####0.00;0;0") & "," 'Field36(�����ʻ�֧�����),
                strSQL = strSQL & "'" & Substr(Nvl(rsData!סԺ����), 1, 20) & "'," 'Field37(סԺ����),
                strSQL = strSQL & "'" & Substr(Nvl(rsData!סԺҽʦ), 1, 16) & "',"  'Field38(����ҽʦ),
                strSQL = strSQL & "'" & Substr(gstrUserName, 1, 4) & "'," 'Field39(ҽԺ����Ա),
                strSQL = strSQL & "" & Format(Val(strArr(17)) / 100, "####0.00;-####0.00;0;0") & "," 'Field29(�ؼ������Ը����),
                strSQL = strSQL & "" & Format(Val(strArr1(7)) / 100, "####0.00;-####0.00;0;0") & "," 'Field40(ͳ��֧�����),
                strSQL = strSQL & "" & Format(Val(strArr(22)) / 100, "####0.00;-####0.00;0;0") & "," 'Field41(�����ν��),
                strSQL = strSQL & "'" & g�������_�˳�.��״̬ & "',"     'C_kzt38(��״̬),
'                strSql = strSql & "'0'" & ","  'C_zt39(����״̬)
                strSQL = strSQL & "'0'" & ")"  'C_zt39(����״̬)
'                strSql = strSql & "'" & g�������_�˳�.���ҽԺ & "',"      'C_kzt38(��״̬),
'                strSql = strSql & "'" & g�������_�˳�.���ҽԺ���� & "')"
                gcnSQLSEVER_�˳�.Execute strSQL
                '�����м��:
                If InsertIntoYBK(strArr, lngRow, False, False) = False Then Exit Function
        Else
        
             'סԺ:������ˮ��|��Ŀ���|��Ŀ����|��Ŀ����|��Ŀ���|����|����|�����ܽ��|ҽ���ڷ���|ҽ�������|���÷�������|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|MAC2
             '��ṹ:c_receipt_no(סԺ������ˮ��),sort_it(�շ�˳���),c_insu_id(���ձ��),c_item_code(��Ŀ����),c_item_name(��Ŀ����),c_dir(����Ŀ¼����),c_zlxmjb(������Ŀ����),c_ser(������ʩ��Χ),n_price(����),n_amount(����),n_account(���),n_med_jia(ҩƷ������),n_med_yi(����ҩƷҽ�����),n_med_yi_self(����ҩƷ�Ը����),n_med_self(�Է�ҩƷ���),n_zlyb(����ҽ�����),n_tjtzyb(�ؼ�����ҽ�����),n_tjtzyb_self(�ؼ������Ը����),n_zlfyb(���Ʒ�ҽ�����),n_fwyb(����ҽ�����),n_fwfyb(�����ҽ�����),dc_fyfsrq(��������),c_wzzxm(��ת����Ŀ),C_zt(����״̬)
             
             
             strSQL = "insert into HLD_HOSP(c_receipt_no,sort_ID,c_insu_id,c_item_code,c_item_name,c_dir,c_zlxmjb,c_ser,n_price,n_amount,n_account,n_med_jia,n_med_yi,n_med_yi_self,n_med_self,n_zlyb,n_tjtzyb,n_tjtzyb_self,n_zlfyb,n_fwyb,n_fwfyb,dc_fyfsrq,c_wzzxm,C_zt) values("
             
             strSQL = strSQL & "'" & strArr1(0) & "',"  '    c_receipt_no(סԺ������ˮ��),
             strSQL = strSQL & "" & lngRow - 1 & "," '    sort_it(�շ�˳���),
             strSQL = strSQL & "'" & g�������_�˳�.���ձ�� & "',"  '    c_insu_id(���ձ��),
             strSQL = strSQL & "'" & strArr(2) & "',"  '    c_item_code(��Ŀ����),
             strSQL = strSQL & "'" & strArr(3) & "',"  '    c_item_name(��Ŀ����),
             
             '�º�����200500403�޸���ӣ���Ϊ�˲��ֽӿ��ĵ�������̫�������ȷ�������£�
             '1������Ŀ¼����ӦΪ 0-ҩƷ��1-���ƣ�2-����
             '2��������Ŀ����ӦΪ 0-ҽ����1-��ҽ��
             '3��������ʩ��Χû��ʲô����
             
             Select Case strArr(4)
                 Case "0"    'ҩƷ
                     gstrSQL = "Select ssmljb as ���� From YB_YD where xmdm='" & strArr(2) & "'"
                 Case "1"    '����
                     gstrSQL = "Select tjtzbz as ���� From YB_ZLML where xmdm='" & strArr(2) & "'"
                 Case Else   '����
                     gstrSQL = "Select fwfw as ���� From YB_FWSS where xmdm='" & strArr(2) & "'"
             End Select
             If rsTemp.State = 1 Then rsTemp.Close
             rsTemp.Open gstrSQL, gcnSQLSEVER_�˳�

'             '�º�����20050321�޸�
'
'             If rsTemp.EOF Then
'                 ShowMsgbox "�ڱ�������¼ʱ,δ������ص�ҽ����Ŀ����[" & strArr(2) & "]"
'                 Exit Function
'             End If

              Select Case strArr(4)
                 Case "0"    'ҩƷ
                    '�º�����20050321�޸�,��Ϊ���ڲ���������Ŀ

                     strSQL = strSQL & "'0',"           'C_SSMLJB (����Ŀ¼����)
                     If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (������Ŀ����)
                     Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (������Ŀ����)
                     End If

                 Case "1"    '����
                    strSQL = strSQL & "'1',"              'C_SSMLJB (����Ŀ¼����)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"           'C_ZLXMJB (������Ŀ����)
                    Else
                       strSQL = strSQL & "'0',"        'C_SSMLJB (����Ŀ¼����)
                    End If

                 Case Else   '����
                    strSQL = strSQL & "'2',"           'C_SSMLJB (����Ŀ¼����)
                    If rsTemp.EOF Then
                       strSQL = strSQL & "'1',"        'C_ZLXMJB (������Ŀ����)
                    Else
                       strSQL = strSQL & "'0',"        'C_ZLXMJB (������Ŀ����)
                    End If
                    
             End Select
             
             strSQL = strSQL & "NULL,"        'C_FWSSFW (������ʩ��Χ)

'�º�����20050403�޸�
'            Select Case strArr(4)
'                 Case "0"    'ҩƷ
'                    '�º�����20050321�޸�,��Ϊ���ڲ���������Ŀ
'                     If rsTemp.EOF Then
'                        strSql = strSql & "3,"           'C_SSMLJB (����Ŀ¼����)
'                        strSql = strSql & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                        strSql = strSql & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                     Else
'                        strSql = strSql & "'" & Nvl(rsTemp!����) & "',"        'C_SSMLJB (����Ŀ¼����)
'                        strSql = strSql & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                        strSql = strSql & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                     End If
'
'                 Case "1"    '����
'                    If rsTemp.EOF Then
'                       strSql = strSql & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                       strSql = strSql & "2,"           'C_ZLXMJB (������Ŀ����)
'                       strSql = strSql & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                    Else
'                       strSql = strSql & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                       strSql = strSql & "'" & Nvl(rsTemp!����) & "',"        'C_ZLXMJB (������Ŀ����)
'                       strSql = strSql & "NULL,"        'C_FWSSFW (������ʩ��Χ)
'                    End If
'
'                 Case Else   '����
'                    If rsTemp.EOF Then
'                       strSql = strSql & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                       strSql = strSql & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                       strSql = strSql & "1,"           'C_FWSSFW (������ʩ��Χ)
'                    Else
'                       strSql = strSql & "NULL,"        'C_SSMLJB (����Ŀ¼����)
'                       strSql = strSql & "NULL,"        'C_ZLXMJB (������Ŀ����)
'                       strSql = strSql & "'" & Nvl(rsTemp!����) & "',"        'C_FWSSFW (������ʩ��Χ)
'                    End If
'             End Select
                     
             strSQL = strSQL & "" & Format(Val(strArr(5)) / 10000, "####0.00;-####0.00;0;0") & ","         '    n_price(����),
             strSQL = strSQL & "" & Format(Val(strArr(6)), "####0;-####0;0;0") & ","         '    n_amount(����),
             strSQL = strSQL & "" & Format(Val(strArr(7)) / 100, "####0.00;-####0.00;0;0") & "," '    n_account(���),
             strSQL = strSQL & "" & Format(Val(strArr(11)) / 100, "####0.00;-####0.00;0;0") & "," '    n_med_jia(ҩƷ������),
             strSQL = strSQL & "" & Format(Val(strArr(12)) / 100, "####0.00;-####0.00;0;0") & "," '    n_med_yi(����ҩƷҽ�����),
             strSQL = strSQL & "" & Format(Val(strArr(13)) / 100, "####0.00;-####0.00;0;0") & "," '    n_med_yi_self(����ҩƷ�Ը����),
             strSQL = strSQL & "" & Format(Val(strArr(14)) / 100, "####0.00;-####0.00;0;0") & "," '    n_med_self(�Է�ҩƷ���),
             strSQL = strSQL & "" & Format(Val(strArr(15)) / 100, "####0.00;-####0.00;0;0") & "," '    n_zlyb(����ҽ�����),
             strSQL = strSQL & "" & Format(Val(strArr(16)) / 100, "####0.00;-####0.00;0;0") & ","  '    n_tjtzyb(�ؼ�����ҽ�����),
             strSQL = strSQL & "" & Format(Val(strArr(17)) / 100, "####0.00;-####0.00;0;0") & "," '    n_tjtzyb_self(�ؼ������Ը����),
             strSQL = strSQL & "" & Format(Val(strArr(18)) / 100, "####0.00;-####0.00;0;0") & "," '    n_zlfyb(���Ʒ�ҽ�����),
             strSQL = strSQL & "" & Format(Val(strArr(19)) / 100, "####0.00;-####0.00;0;0") & "," '    n_fwyb(����ҽ�����),
             strSQL = strSQL & "" & Format(Val(strArr(20)) / 100, "####0.00;-####0.00;0;0") & "," '    n_fwfyb(�����ҽ�����),
             
             strSQL = strSQL & "'" & strArr(10) & "'," '    dc_fyfsrq(��������),
             strSQL = strSQL & "'0'," '    c_wzzxm(��ת����Ŀ),
             strSQL = strSQL & "'0')" '    C_zt (����״̬)
             gcnSQLSEVER_�˳�.Execute strSQL
            '�����м��:
            If InsertIntoYBK(strArr, lngRow, True, False) = False Then Exit Function
        End If
        lngRow = lngRow + 1
    Loop
    InsertIntoData_סԺ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function




Private Function InsertIntoData_סԺ�Ǽ�(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '�������ݵ�SQLSErver��.
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������������ݲ���SQLServer��
    '--�����:lng����IDֵ
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim str�������� As String
    Dim str����ҽʦ As String
    
    
    
    
    
    InsertIntoData_סԺ�Ǽ� = False
        
    Err = 0: On Error GoTo errHand:
    
    gstrSQL = "" & _
        "   Select a.��Ժ����ID,b.���� as ��Ժ����,to_char(a.��Ժ����,'yyyy-mm-dd hh24:mi:ss') as ��Ժ����" & _
        "   From ������ҳ a,���ű� b" & _
        "   where a.��Ժ����id=b.id(+) and a.����id=" & lng����ID & " and a.��ҳid =" & lng��ҳID
        
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��Ժ��Ϣ"
    
    '��ṹ:xh����ţ�,yljgdm��ҽ�ƻ������룩,bxbh�����ձ�ţ�,zyh��סԺ�ţ�,ryrq����Ժ���ڣ�,rylb����Ժ���,bz�����֣�,zyks��סԺ���ң�,C_zt������״̬��
    '
    'ע��:�ӿ��������,���ڿ��������
    gstrSQL = "insert into HLD_ZYBRXX(yljgdm,bxbm,zyh,ryrq,rylb,bz,zyks,C_zt) values ("
    
    '����
    'gstrSQL = gstrSQL & "," 'xh����ţ�
    
    gstrSQL = gstrSQL & "'" & InitInfor_�˳�.ҽԺ���� & "',"   'yljgdm��ҽ�ƻ������룩
    gstrSQL = gstrSQL & "'" & g�������_�˳�.���ձ�� & "',"  'bxbh�����ձ�ţ�,
    gstrSQL = gstrSQL & "'" & lng����ID & "_" & lng��ҳID & "',"   'zyh��סԺ�ţ�,
    gstrSQL = gstrSQL & "'" & Nvl(rsTemp!��Ժ����) & "'," 'ryrq����Ժ���ڣ�,
    gstrSQL = gstrSQL & "'" & g�������_�˳�.��Ժ��� & "',"  'rylb����Ժ���,
    gstrSQL = gstrSQL & "'" & g�������_�˳�.���ֱ��� & "',"  'bz�����֣�,
    gstrSQL = gstrSQL & "'" & Nvl(rsTemp!��Ժ����) & "',"  'zyks��סԺ���ң�,
    gstrSQL = gstrSQL & "'0')" 'C_zt������״̬��
    
    gcnSQLSEVER_�˳�.Execute gstrSQL

    InsertIntoData_סԺ�Ǽ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InsertIntoYBK(ByVal strVarr As Variant, ByVal lng��� As Long, Optional bln��ϸ As Boolean = True, Optional bln���� As Boolean = True) As Boolean
    '���ݴ�������������ݱ��浽�м��ȥ
    Dim strSQL As String
    
    Dim i As Long
    Dim str�������� As String
    
    Err = 0: On Error GoTo errHand:
    
    InsertIntoYBK = False
    If bln��ϸ = False Then
        i = 0
        '���̲���
        '    ����_IN,����ID_IN,��Ժ���_IN IN,ҽԺ����_IN,����_IN,
        '    ��¼����_IN,�����ܶ�_IN,ͳ��֧�����_IN,����Ӧ���ܶ�_IN,
        '    �����ʻ����_IN,��ҽ����Χ���_IN ,ҽ����Χ���_IN,ҩƷ������_IN,����ҩƷҽ�����_IN,����ҩƷ�Ը����_IN,
        '    �Է�ҩƷ���_IN,����ҽ�����_IN,�ؼ�����֧����_IN,�ؼ������Ը���_IN,���Ʒ�ҽ�����_IN,
        '    ����ҽ�����_IN,�����ҽ�����_IN,���ⲡ��ʶ_IN,�𸶶�_IN,�����ν��_IN,�ⶥ���Ը����_IN,
        '    ���ͳ������ۼ�_IN,��ȹ������ۼ�_IN,����סԺ����_IN
        strSQL = "ZL_ҽ�������¼_INSERT("
        '����:                                 �ļ���¼����|���ν��׷����ܽ��             |����Ӧ���ܽ��|�����ʻ����|��ҽ����Χ���|ҽ����Χ���|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|���ⲡ��ʶ
        'סԺ:   IC����|��Ժ���|ҽԺ����|����|�ļ���¼����|���ν��׷����ܽ��|ͳ��֧�����|����Ӧ���ܽ��|�����ʻ����|��ҽ����Χ���|ҽ����Χ���|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|�𸶶�|�����ν��|�ⶥ�������Ը����|�����ͳ�����֧���ۼ�|����ȹ����ν���ۼ�|����סԺ����
        
        If bln���� Then
            strSQL = strSQL & 1 & ","
            strSQL = strSQL & g�������_�˳�.����ID & ","
        Else
            strSQL = strSQL & IIf(g�������_�˳�.����ID = 0, 3, 2) & ","
            strSQL = strSQL & IIf(g�������_�˳�.����ID = 0, g�������_�˳�.����ID, g�������_�˳�.����ID) & ","
        End If
        
        If Not bln���� Then
            i = i + 1
            '��Ժ���_IN IN,ҽԺ����_IN,����_IN
            strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
            strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
            strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
        Else
            strSQL = strSQL & "NULL" & ","
            strSQL = strSQL & "NULL" & ","
            strSQL = strSQL & "NULL" & ","
        End If
        
        strSQL = strSQL & Val(strVarr(i)) & ",": i = i + 1
        
        '�º�����20050315�޸�,�������ݿ����ݽ��и���;��ԭ��/100��Ϊ/100
        
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        
        If bln���� Then
            'ͳ��֧�����
            strSQL = strSQL & 0 & ","
        Else
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        End If
        
        '�º�����20050315�޸ģ����ڴ��뵽���ݿ�������뷢���ķ��ò���
        
       ' MsgBox strSql, vbOKOnly, "zlsoft"
       
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        If bln���� Then
            '���ⲡ��ʶ
            strSQL = strSQL & "'" & strVarr(16) & "',"
            strSQL = strSQL & 0 & ","
            strSQL = strSQL & 0 & ","
            strSQL = strSQL & 0 & ","
            strSQL = strSQL & 0 & ","
            strSQL = strSQL & 0 & ","
            strSQL = strSQL & 0 & ")"
        Else
            '�𸶶�|�����ν��|�ⶥ�������Ը����|�����ͳ�����֧���ۼ�|����ȹ����ν���ۼ�|����סԺ����
            '���ⲡ��ʶ_IN,�𸶶�_IN,�����ν��_IN,�ⶥ���Ը����_IN,���ͳ������ۼ�_IN,��ȹ������ۼ�_IN,����סԺ����_IN
            strSQL = strSQL & "NULL,"
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
            strSQL = strSQL & Format(Val(strVarr(i)) / 100, "#####0.00;-####0.00;0.00;0.00") & ",": i = i + 1
        
            '�º����޸���20060109�޸�,����סԺ��������ȷ
            'MsgBox "����ҽ�������¼" & i & "-" & Format(Val(strVarr(i)) + 1, "#####;-####0;0;0"), vbOKOnly, gstrSysName
            
            strSQL = strSQL & Format(Val(strVarr(i)) + 1, "#####;-####0;0;0") & ")"
        End If
        gstrSQL = strSQL
        ExecuteProcedure_�˳� "���뱣�ս����¼"
        InsertIntoYBK = True
        Exit Function
    End If
    '������ϸ��¼
    '����:������ˮ��|  ������|��Ŀ����|��Ŀ����|��Ŀ���|����|����|�����ܽ��|ҽ���ڷ���|ҽ�������             |ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|MAC2
    'סԺ:������ˮ��|��Ŀ���|��Ŀ����|��Ŀ����|��Ŀ���|����|����|�����ܽ��|ҽ���ڷ���|ҽ�������|���÷�������|ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|MAC2
    '���̲���:
    '   ����ID_IN,����_IN,������ˮ��_IN,������_IN,��Ŀ����_IN,��Ŀ����_IN,��Ŀ���_IN������_IN��
    '   ����_IN,�����ܽ��_IN,ҽ���ڷ���_IN,ҽ�������_IN,ҩƷ������_IN,����ҽ�����_IN�������Ը����_IN��
    '   �Է�ҩƷ���_IN������ҽ�����_IN���ؼ����ν��_IN���ؼ������Ը�_IN�����Ʒ�ҽ����_IN������ҽ�����_IN�������ҽ����_IN�����÷�������_IN��MAC2_IN��
    i = 0
    strSQL = "ZL_ҽ��������ϸ��¼_INSERT("
    If bln���� Then
        strSQL = strSQL & g�������_�˳�.����ID & ","
        strSQL = strSQL & "1" & ","
    Else
        strSQL = strSQL & IIf(g�������_�˳�.����ID = 0, g�������_�˳�.����ID, g�������_�˳�.����ID) & ","
        strSQL = strSQL & IIf(g�������_�˳�.����ID = 0, 3, 2) & ","
    End If
    strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
    strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
    strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
    strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
    strSQL = strSQL & "'" & strVarr(i) & "',": i = i + 1
    
    
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 10000, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    
    '�º�����20050315�޸ģ����ڸ��±��ս�����ϸ��¼��ʱ���������ֶβ���Ҫ/100����val(strVarr(i))/100�޸�Ϊval(strVarr(i))
    
    strSQL = strSQL & "" & Format(Val(strVarr(i)), "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    If Not bln���� Then
        'סԺ����÷�������
        str�������� = zlCommFun.AddDate(strVarr(i)): i = i + 1
        If Not IsDate(str��������) Then
            str�������� = ""
        End If
    Else
        str�������� = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    End If
    'ҩƷ������|����ҩƷҽ�����|����ҩƷ�Ը����|�Է�ҩƷ���|����ҽ�����|�ؼ�����ҽ�����|�ؼ������Ը����|���Ʒ�ҽ�����|����ҽ�����|�����ҽ�����|MAC2
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & "" & Format(Val(strVarr(i)) / 100, "#####0.0000;-#####0.0000;0.0000;0.0000") & ",": i = i + 1
    strSQL = strSQL & IIf(str�������� = "", "NULL", "to_date('" & str�������� & "','yyyy-mm-dd')") & ","
    strSQL = strSQL & "'" & strVarr(i) & "',"
    strSQL = strSQL & "" & lng��� & ")"
    gstrSQL = strSQL
    ExecuteProcedure_�˳� "���뱣�ս����¼"
    InsertIntoYBK = True
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    InsertIntoYBK = False
End Function
Private Function ��ȡ�����ʻ�֧��() As Double
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ�����ʻ�ֵ(��Ԥ����¼�л�ȡ)
    '--�����:
    '--������:
    '--��  ��:�ɹ�,���ر��θ����ʻ�֧��,���򷵻�0
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select * From ����Ԥ����¼ where ����ID=[1] and  ���㷽ʽ='�����ʻ�'"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ�֧��", g�������_�˳�.����ID)
    If Not rsTemp.EOF Then
        ��ȡ�����ʻ�֧�� = Nvl(rsTemp!��Ԥ��, 0)
    End If
    
End Function

Public Function �����������_�˳�(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    
    Dim str��ϸ As String
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    
    
    g�������_�˳�.�����ܶ� = 0
    str��ϸ = ""
    
    '��һ��:���ܷ���
    DebugTool "�����������,��һ��:���ܷ���"
    With rs��ϸ
        If rs��ϸ.RecordCount = 0 Then ShowMsgbox "δ������صķ��ü�¼!": Exit Function
        Do While Not .EOF
            g�������_�˳�.�����ܶ� = g�������_�˳�.�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
    
    '�ڶ���:д�����ϸ�ļ�
    'д����ļ�
    DebugTool "�����������,�ڶ���:׼��д�����ϸ�ļ�"
    If WriteINParaFile(rs��ϸ, True, False) = False Then
        DebugTool "         ����ļ���ϸд��ʧ��"
        Exit Function
    End If
    DebugTool "         ���д������ļ���ϸ"
    
    
    '������:�����������ֽ�
    DebugTool "�����������,������:�����������ֽ�"
    If ҵ������_�˳�(�˳�_�������Ԥ�ֽ�, "", "") = False Then
        DebugTool "             �����������ֽ�ʧ��"
        Exit Function
    End If
    DebugTool "             �����������ֽ�ɹ�"
    
    '���Ĳ�:�ֽ���س��ν��
    
    DebugTool "�����������,�ڶ���:�ֽ���س��ν��"
    
    If ReadOutParaFile(g��������, True, True) = False Then
        DebugTool "     �ֽ���س��ν��ʧ��!"
        Exit Function
    End If
    DebugTool "     �ֽ���س��ν���ɹ�!"
    
    If Format(g�������_�˳�.�����ܶ�, "#####0.00;-####0.00;0;0") <> Format(g��������.���׷����ܶ�, "#####0.00;-####0.00;0;0") Then
        ShowMsgbox "�����ܶ��,���ܽ���!" & vbCrLf & _
                " HIS�����ܶ�:" & Format(g�������_�˳�.�����ܶ�, "#####0.00;-####0.00; ;") & _
                " ��������ܶ�:" & Format(g��������.���׷����ܶ�, "#####0.00;-####0.00; ;")
        Exit Function
    End If
    str���㷽ʽ = ""
    
    '�º�������20050321�޸�,�������ĵĽ��㷽ʽ:�ڽ�������ҵ��ʱ:��������ʻ��������ڱ��ν��׶�,�Ӹ����ʻ�֧��
    
    With g��������
      
     '�º�����20050320�޸�,�������Բ��Ľ��㷽ʽ��һ��,�����ֽ�ʽ֧��
     If g�������_�˳�.������ʶ <> "1" Then
        
         '�º�����20050403�޸����ӣ���Ϊ���Ǹ����ʻ������С�ڵ�ǰ�����ʻ�֧���Ľ��
         If (.���׷����ܶ� - .����Ӧ���ܶ� + .����ҩƷ�Ը��� + .�ؼ������Ը�) - .�����ʻ���� <= 0 Then
             str���㷽ʽ = str���㷽ʽ & "�����ʻ�;" & Format(.���׷����ܶ� - .����Ӧ���ܶ� + .����ҩƷ�Ը��� + .�ؼ������Ը�, "####0.00;-####0.00;0;0") & ";1"
         Else
             str���㷽ʽ = str���㷽ʽ & "�����ʻ�;" & Format(.�����ʻ����, "####0.00;-####0.00;0;0") & ";1"
         End If
     Else
      
      '�º�����20050402�޸�,�������Բ����߿��԰������Բ���ʽ��ҽ����
        If blnmxb = True Then
         str���㷽ʽ = str���㷽ʽ & "�����ʻ�;" & Format(.���׷����ܶ� - .����Ӧ���ܶ�, "####0.00;-####0.00;0;0") & ";1"
        Else
          If (.���׷����ܶ� - .����Ӧ���ܶ� + .����ҩƷ�Ը��� + .�ؼ������Ը�) - .�����ʻ���� <= 0 Then
             str���㷽ʽ = str���㷽ʽ & "�����ʻ�;" & Format(.���׷����ܶ� - .����Ӧ���ܶ� + .����ҩƷ�Ը��� + .�ؼ������Ը�, "####0.00;-####0.00;0;0") & ";1"
          Else
             str���㷽ʽ = str���㷽ʽ & "�����ʻ�;" & Format(.�����ʻ����, "####0.00;-####0.00;0;0") & ";1"
          End If
        End If
        
     End If
     
    End With
    
    DebugTool "�����������ɹ�,���㷽ʽ��" & str���㷽ʽ
    �����������_�˳� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function �������_�˳�(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim StrInput As String, strOutput As String
    Dim lng����ID  As Long
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strArr As Variant
    Dim ����������� As ��������
    
    
    Err = 0: On Error GoTo errHandle

    Call DebugTool("�����������")

    gstrSQL = "" & _
        "   Select a.*,a.����*a.���� as ����,a.ʵ�ս��/(nvl(a.����,1)*nvl(a.����,1)) as ���� " & _
        "   From ������ü�¼ a " & _
        "   Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"

    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ϸ��¼", lng����ID)

    If rs��ϸ.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")

    If g�������_�˳�.����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣"
        Exit Function
    End If
    g�������_�˳�.����ID = lng����ID
    
    '�º�����20050310�޸ģ����ڡ�g�������_�˳�.�����ܶ���ϴε��ú�û�н�������
    
    g�������_�˳�.�����ܶ� = 0
    
    '��һ��:���ܷ���
    DebugTool "�������,��һ��:���ܷ���"
    With rs��ϸ
        If rs��ϸ.RecordCount = 0 Then ShowMsgbox "δ������صķ��ü�¼!": Exit Function
        Do While Not .EOF
            g�������_�˳�.�����ܶ� = g�������_�˳�.�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
    gcnOracle_�˳�.BeginTrans
    gcnSQLSEVER_�˳�.BeginTrans
    '�ڶ���:д�����ϸ�ļ�
    'д����ļ�
    DebugTool "�������,�ڶ���:׼��д�����ϸ�ļ�"
    
    '�º�����20050315�޸ģ������������_�˳�()����true �޸�Ϊfalse
    
    If WriteINParaFile(rs��ϸ, False, False) = False Then
        DebugTool "         ����ļ���ϸд��ʧ��"
        gcnOracle_�˳�.RollbackTrans
        gcnSQLSEVER_�˳�.RollbackTrans
        Exit Function
    End If
    DebugTool "         ���д������ļ���ϸ"
    
    
    '������:�����������ֽ�
    DebugTool "�������,������:�������ֽ�"
    If ҵ������_�˳�(�˳�_�������Ԥ�ֽ�, "", "") = False Then
        DebugTool "             �������ֽ�ʧ��"
        gcnOracle_�˳�.RollbackTrans
        gcnSQLSEVER_�˳�.RollbackTrans
        Exit Function
    End If
    DebugTool "             �������ֽ�ɹ�"
    
    '���Ĳ�:�ֽ���س��ν��
    
    DebugTool "�������,���Ĳ�:�ֽ���س��ν��"
    
    If ReadOutParaFile(�����������, True, False) = False Then
        DebugTool "     �ֽ���س��ν��ʧ��!"
        gcnOracle_�˳�.RollbackTrans
        gcnSQLSEVER_�˳�.RollbackTrans
        Exit Function
    End If
    DebugTool "     �ֽ���س��ν���ɹ�!"
    
    
    If Format(g��������.���׷����ܶ�, "#####0.00;-####0.00;0;0") <> Format(�����������.���׷����ܶ�, "#####0.00;-####0.00;0;0") Then
        ShowMsgbox "�����ܶ��,���ܽ���!" & vbCrLf & _
                " �����������ܶ�:" & Format(g��������.���׷����ܶ�, "#####0.00;-####0.00; ;") & _
                " ��ʽ��������ܶ�:" & Format(�����������.���׷����ܶ�, "#####0.00;-####0.00; ;")
        gcnOracle_�˳�.RollbackTrans
        gcnSQLSEVER_�˳�.RollbackTrans
        Exit Function
    End If
    
    '���岽: ����֧��ȷ��
    DebugTool "�������,���岽: ����֧��ȷ��"
    '   ������ˮ��|IC����|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|�����ʻ�֧�����|�ֽ�֧�����

    StrInput = Rpad(Substr(InitInfor_�˳�.ҽԺ����, 1, 8), 8, " ") & Lpad(Substr(g�������_�˳�.����ID, 1, 12), 12, "0")
    StrInput = StrInput & "|" & g�������_�˳�.IC����
    StrInput = StrInput & "|" & Int(Format((�����������.���׷����ܶ� * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((�����������.ҽ����Χ��� * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((�����������.��ҽ����Χ�� * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((cur�����ʻ� * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format(((�����������.���׷����ܶ� - cur�����ʻ�) * 100), "####0.00;-####0.00;0;0"))
    
    g�������� = �����������
    
'    MsgBox strInput, vbOKOnly, "zlsoft"
    
    If ҵ������_�˳�(�˳�_��ͨ����֧��ȷ��, StrInput, strOutput) = False Then
        DebugTool "     ��ͨ����֧��ȷ��ʧ��!"
        gcnOracle_�˳�.RollbackTrans
        gcnSQLSEVER_�˳�.RollbackTrans
        Exit Function
    End If
    DebugTool "     ��ͨ����֧��ȷ�ϳɹ�!"
    
    
    '�ύ��صĽ�������
    
    '�º�����20050402�޸�,Ϊ�˵��Է������
'    MsgBox strOutPut, vbOKOnly, "zlsoft"
       
    If InsertIntoSQLServer_����(strOutput) = False Then
        gcnOracle_�˳�.RollbackTrans
        gcnSQLSEVER_�˳�.RollbackTrans
        Exit Function
    End If
        
    '��д�����
    Call DebugTool("��д�����¼")

    
    '������ˮ��|IC����|�ն˻����|��������/ʱ��|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|�����ʻ�֧�����|�ֽ�֧�����|�ۼ�������ʻ����|MAC1
    strArr = Split(strOutput, "|")
    
'    MsgBox strOutPut, vbOKOnly, "zlsoft"
'    MsgBox strArr(10), vbOKOnly, "zlsoft"
     
    '�º�����20050315�޸���ӣ���Ҫ���ڵ���ExecuteProcedure("��������¼")ʱ����,��˽�ȡstrArr(10)5���ַ���
     
     strArr(10) = Substr(strArr(10), 1, 16)
'     MsgBox strArr(10), vbOKOnly, "zlsoft"

   '���뱣�ս����¼
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(),סԺ����_IN(סԺ:��ҳid),����(),�ⶥ��_IN(),ʵ������_IN(�ۼ�������ʻ����),
    '   �������ý��_IN(�����ܽ��),ȫ�Ը����_IN(�ֽ�֧�����),�����Ը����_IN(ҽ�����ܷ���),
    '   ����ͳ����_IN(ҽ�����ܷ���),ͳ�ﱨ�����_IN(����:סԺ:ͳ��֧�����),���Ը����_IN(סԺ:ͳ���Ը����),�����Ը����_IN(),�����ʻ�֧��_IN(�����ʻ�֧�����),"
    '   ֧��˳���_IN(������ˮ��),��ҳID_IN(��ҳid),��;����_IN,��ע_IN(�ն˻����|��������/ʱ��|MAC1)

    
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
    
    '�º�����20050311�޸�,��������Ҫ����100��,�ڴ����ҷ����ݿ���
'
'    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�˳ɺ˹�ҵ & "," & lng����id & "," & Year(zlDatabase.Currentdate) & "," & _
'            "NULL,NULL,NULL,NULL,null,0,0," & Format(Val(strArr(9)), "#####0.00;-####0.00;0;0") & "," & _
'            Format(Val(strArr(4)), "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(8)), "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(6)), "#####0.00;-####0.00;0;0") & "," & _
'            Format(Val(strArr(5)), "#####0.00;-####0.00;0;0") & " ,0,0,0," & Format(Val(strArr(7)), "#####0.00;-####0.00;0;0") & ",'" & _
'             strArr(0) & "',NULL,NULL,'" & strArr(2) & "|" & strArr(3) & "|" & strArr(10) & "')"
             
       
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�˳ɺ˹�ҵ & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "NULL,NULL,NULL,NULL,null,0,0," & Format(Val(strArr(9)) / 100, "#####0.00;-####0.00;0;0") & "," & _
            Format(Val(strArr(4)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(8)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(6)) / 100, "#####0.00;-####0.00;0;0") & "," & _
            Format(Val(strArr(5)) / 100, "#####0.00;-####0.00;0;0") & " ,0,0,0," & Format(Val(strArr(7)) / 100, "#####0.00;-####0.00;0;0") & ",'" & _
             strArr(0) & "',NULL,NULL,'" & strArr(2) & "|" & strArr(3) & "|" & strArr(10) & "')"

    'MsgBox gstrSQL, vbOKOnly, "zlsoft"

    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    gcnOracle_�˳�.CommitTrans
    gcnSQLSEVER_�˳�.CommitTrans
    �������_�˳� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function


Public Function ����������_�˳�(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��

    Dim intMouse As Integer
    Dim lng����ID  As Long
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsԭ��ϸ As New ADODB.Recordset
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    Dim lng����id1 As Long
    
    ����������_�˳� = False


    '�����֤
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If ��ݱ�ʶ_�˳�(2, lng����id1) = "" Then
        If lng����id1 = 0 Then
            Err.Raise 9000, gstrSysName, "�㲻�ǵ�ǰ�ֿ���!"
            Screen.MousePointer = intMouse
            Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo errHand:
    Screen.MousePointer = intMouse
        
    '��һ��:ȷ������IDֵ
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    lng����ID = rsTemp("����ID")

    '�ڶ���:ȷ��������ԭʼ���ݵ���ϸ��¼

    gstrSQL = "Select * From ������ü�¼ " & _
        " Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������¼", lng����ID)
    g�������_�˳�.�����ܶ� = 0
    g�������_�˳�.����ID = lng����ID
    
    
    gstrSQL = "Select * From ������ü�¼ where  ����ID =[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Set rsԭ��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������¼", lng����ID)

    gcnOracle_�˳�.BeginTrans
    gcnSQLSEVER_�˳�.BeginTrans
    
    
    '������:��ԭʼ��¼�еĽ�Ҫ��Ϊ�������ݵĽ�Ҫ���������ϴ���־
    With rs��ϸ
        Do While Not .EOF
            rsԭ��ϸ.Filter = 0
            rsԭ��ϸ.Filter = "no='" & Nvl(!NO) & "' and ��¼����=" & Nvl(!��¼����, 0) & " and ���=" & Nvl(!���, 0)         '& "' and ִ��״̬=" & Nvl(!ִ��״̬, 0)
            If rsԭ��ϸ.EOF Then
                Err.Raise 9000, gstrSysName, "û���ҵ�ԭʼ��¼!" & "no='" & Nvl(!NO) & "' and ��¼����=" & Nvl(!��¼����, 0) & " and ���='" & Nvl(!���, 0) & "' and ִ��״̬=" & Nvl(!ִ��״̬, 0)
                gcnOracle_�˳�.RollbackTrans
                gcnSQLSEVER_�˳�.RollbackTrans
                Exit Function
            End If
            'д�ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(rsԭ��ϸ!ժҪ) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
            g�������_�˳�.�����ܶ� = g�������_�˳�.�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
        
        
    '���Ĳ�:�����м�����ؼ�¼
    
    '����������ý����¼
    '���̲���:����_IN,ԭ����ID_IN,�ֽ���ID_IN
    gstrSQL = "ZL_ҽ������_����("
    gstrSQL = gstrSQL & "1,"
    gstrSQL = gstrSQL & lng����ID & ","
    gstrSQL = gstrSQL & lng����ID & ")"
    
    ExecuteProcedure_�˳� "�������ý����¼"
    
    
    
    '���岽:��ȡ���ս����¼�е����ݣ�ͬʱ�����˷Ѻ���
    
    gstrSQL = "Select * from ���ս����¼ where ����=1 and ��¼id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���ĵ��ݺ�"
    
    If rsTemp.EOF Then
        gcnOracle_�˳�.RollbackTrans
        gcnSQLSEVER_�˳�.RollbackTrans
        Err.Raise 9000, gstrSysName, "�����ڽ����¼,���ܳ���!"
        Exit Function
    End If
    
    '������ˮ��|IC����|�˷ѽ��
    '�º�����20050315�޸ģ���Ҫ����PoliBackCost()������ڲ�������
    
'     strInput = Nvl(rsTemp!֧��˳���) & "|"
'     strInput = strInput & Nvl(rsTemp!֧��˳���)
'     strInput = strInput & g�������_�˳�.IC����
'     strInput = strInput & Nvl(rsTemp!�����ʻ�֧��)
      
     StrInput = Nvl(rsTemp!֧��˳���) & "|"
'    strInput = strInput & Nvl(rsTemp!֧��˳���)
     StrInput = StrInput & g�������_�˳�.IC���� & "|"
     
     StrInput = StrInput & Nvl(rsTemp!�����ʻ�֧��) * 100
     
    If ҵ������_�˳�(�˳�_�����˷Ѻ���, StrInput, strOutput) = False Then
        gcnOracle_�˳�.RollbackTrans
        gcnSQLSEVER_�˳�.RollbackTrans
        Exit Function
    End If
    
    
    '������:���������˷ѵ���ؼ�¼
    
    '���뱣�ս����¼
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(),סԺ����_IN(סԺ:��ҳid),����(),�ⶥ��_IN(),ʵ������_IN(�ۼ�������ʻ����),
    '   �������ý��_IN(�����ܽ��),ȫ�Ը����_IN(�ֽ�֧�����),�����Ը����_IN(ҽ�����ܷ���),
    '   ����ͳ����_IN(ҽ�����ܷ���),ͳ�ﱨ�����_IN(����:סԺ:ͳ��֧�����),���Ը����_IN(סԺ:ͳ���Ը����),�����Ը����_IN(),�����ʻ�֧��_IN(�����ʻ�֧�����),"
    '   ֧��˳���_IN(������ˮ��),��ҳID_IN(��ҳid),��;����_IN,��ע_IN(�ն˻����|��������/ʱ��|MAC1)
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"

    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�˳ɺ˹�ҵ & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL,NULL,null,null,0,0," & -1 * Nvl(rsTemp!ʵ������, 0) & "," & _
           -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & "," & -1 * -1 * Nvl(rsTemp!�����Ը����, 0) & "," & _
           -1 * Nvl(rsTemp!����ͳ����, 0) & "," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & "," & -1 * Nvl(rsTemp!���Ը����, 0) & ",0," & -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & _
           Nvl(rsTemp!֧��˳���, 0) & " ',NULL,NULL,'" & Nvl(rsTemp!��ע) & "')"

    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    
    '���߲�:�ύ��ؽ���
    If InsertIntoSQLServer_�������(lng����ID) = False Then
        gcnOracle_�˳�.RollbackTrans
        gcnSQLSEVER_�˳�.RollbackTrans
        Exit Function
    End If
    '---------------------------------------------------------------------------------------------
    gcnOracle_�˳�.CommitTrans
    gcnSQLSEVER_�˳�.CommitTrans
    ����������_�˳� = True
    Exit Function
errHand:
    gcnOracle_�˳�.RollbackTrans
    gcnSQLSEVER_�˳�.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Private Function Get���״���(ByVal intType As ҵ������_�˳�, Optional bln������ As Boolean = False) As String
    '������û��
    Select Case intType
        Case �˳�_���߻���������
            Get���״��� = IIf(bln������, "�˳�_���߻���������", "01")
        Case �˳�_���߻�����ֹͣ
            Get���״��� = IIf(bln������, "�˳�_���߻�����ֹͣ", "02")
        Case �˳�_POS������
            Get���״��� = IIf(bln������, "�˳�_POS������", "03")
        Case �˳�_POS��ֹͣ
            Get���״��� = IIf(bln������, "�˳�_POS��ֹͣ", "04")
        Case �˳�_��ȡ�ֿ�����Ϣ
            Get���״��� = IIf(bln������, "�˳�_��ȡ�ֿ�����Ϣ", "05")
        Case �˳�_JbylReadIC
            Get���״��� = IIf(bln������, "�˳�_JbylReadIC", "06")
        Case �˳�_�������Ԥ�ֽ�
            Get���״��� = IIf(bln������, "�˳�_�������Ԥ�ֽ�", "07")
        Case �˳�_��ͨ����֧��ȷ��
            Get���״��� = IIf(bln������, "�˳�_��ͨ����֧��ȷ��", "08")
        Case �˳�_�����˷Ѻ���
            Get���״��� = IIf(bln������, "�˳�_�����˷Ѻ���", "09")
        Case �˳�_סԺ����Ԥ�ֽ�
            Get���״��� = IIf(bln������, "�˳�_סԺ����Ԥ�ֽ�", "10")
        Case �˳�_סԺ֧��ȷ��
            Get���״��� = IIf(bln������, "�˳�_סԺ֧��ȷ��", "11")
        Case Else
            Get���״��� = IIf(bln������, "����Ľ��״���", "-1")
    End Select
End Function
Public Function ҵ������_�˳�(ByVal intType As ҵ������_�˳�, strInputString As String, strOutPutstring As String) As Boolean
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
    
    str���״��� = Get���״���(intType, True)
    
    StrInput = strInputString
    DebugTool "����ҵ��������(ҵ�����ʹ���Ϊ:" & intType & " ҵ�����ƣ�" & str���״��� & ")" & vbCrLf & "        �������Ϊ:" & strInputString
    
    
    ҵ������_�˳� = False
    If InitInfor_�˳�.ģ������ Then
        '��ȡģ������
        Readģ������ intType, StrInput, strOutPutstring
         ҵ������_�˳� = True
        Exit Function
    End If
    strArr = Split(strInputString, SP_STR)
    For i = 0 To UBound(strArr)
        strInValue(i) = strArr(i)
    Next
    
    
    Err = 0
    On Error GoTo errHand:
    
    Select Case intType
        Case �˳�_���߻���������
            '�������߷���,��ز���
            '    a. ϵͳĿ¼��������[50] �D�D���ͣ�String    ֵ��һ��ָ����c:\
            '    b. ҽԺ���� [8]         �D�D���ͣ�String  ����Ϊ8λ
            '    c. ODBC����Դ����[ ]   �D�D���ͣ�String     ֵ��һ��ָ����ODBC��DSN
            '    d. ODBC�û���[ ]       �D�D���ͣ�String
            '    e. ODBC�û�����[ ]     �D�D���ͣ�String
            lngReturn = StartPolicy(strInValue(0), strInValue(1), strInValue(2), strInValue(3), strInValue(4))
            
            Select Case lngReturn
            Case 0          '     �D�D�������߻��ɹ�
            Case 100        '     �D�DϵͳĿ¼����
                ShowMsgbox "ϵͳĿ¼����,������������Ƿ���ȷ!"
                Exit Function
            Case 101        '     �D�Dҽ�ƻ����������
                ShowMsgbox "ҽ�ƻ����������,����ҽԺ�����Ƿ�������ȷ!"
                Exit Function
            Case 102        '     �D�D�������ݿ����
                ShowMsgbox "�������ݿ����,������������е�ODBC�����Ƿ���ȷ!"
                Exit Function
            Case -11        '�D�Dδ����ȷ��Ȩ
                ShowMsgbox "�������ݿ����,����ҽ����������ϵ!"
                Exit Function
            Case Else
                ShowMsgbox "δ֪�Ĵ���,���ص����Ϊ:" & lngReturn
                Exit Function
            End Select
            strOutput = ""
            
        Case �˳�_���߻�����ֹͣ
        
            lngReturn = StopPolicy()
            Select Case lngReturn
            Case 0          '     ֹͣ���߻��ɹ�
            Case 103        '�D�D�Ͽ����ݿ����Ӵ���
                ShowMsgbox "�D�D�Ͽ����ݿ����Ӵ���,�����Ѿ��Ͽ�������!"
                Exit Function
            Case Else
                ShowMsgbox "δ֪�Ĵ���,���ص����Ϊ:" & lngReturn
                Exit Function
            End Select
            strOutput = ""
        Case �˳�_POS������
            lngReturn = StartPos()
            If lngReturn <> 0 Then
                ShowMsgbox "����POS������:" & vbCrLf & "��������:" & lngReturn
                Exit Function
            End If
            strOutput = ""
        Case �˳�_POS��ֹͣ
            lngReturn = StopPos()
            If lngReturn <> 0 Then
                ShowMsgbox "ͣ��POS������:" & vbCrLf & "��������:" & lngReturn
                Exit Function
            End If
            strOutput = ""
        Case �˳�_��ȡ�ֿ�����Ϣ
            strOutput = Space(200)
            lngReturn = GetPersonCommInfo(strOutput)
            
         ' �º�����20050310�޸ģ����ݷ��ز�����ʱ�޸ģ���ԭstrOutPut�޸�ΪstrReturn
         
            Select Case lngReturn
            Case 0      '�D�D�ɹ�
            Case 8      '�D�D�ѽ��������
                ShowMsgbox "�ò����Ѿ����������!"
                Exit Function
            Case 9      '�D�D�����ۼƽ����ڹ����ܶ��Ϊ���ҽ�ƿ�����Ҫû�ս�ҽ�����Ĵ���
                ShowMsgbox "�����ۼƽ����ڹ����ܶ��Ϊ���ҽ�ƿ�����Ҫû�ս�ҽ�����Ĵ���!"
                Exit Function
            
            '�º�����20050321�޸ģ��������Բ��Ĵ��󷵻�ֵ���
                
            Case -90  '--�����Բ���û�������Բ��������赽ҽ�����Ĵ���
                 ShowMsgbox "�����Բ���û�������Բ��������赽ҽ�����Ĵ���"
                 Exit Function
            Case -91  '��ǰҽ�ƻ������Ǹÿ������Բ�����ҽ�ƻ���
                 ShowMsgbox "��ǰҽ�ƻ������Ǹÿ������Բ�����ҽ�ƻ�����"
                 Exit Function
            Case -65536
                 ShowMsgbox "�����ҽ�����˵�ҽ������"
                 Exit Function
            Case Else
                ShowMsgbox "��������," & vbCrLf & "��������Ϊ:" & lngReturn
                Exit Function
            End Select
            strOutput = Trim(strOutput)
        Case �˳�_JbylReadIC
            strOutput = Space(200)
           
            '��ʾ���ز���
            'MsgBox JbylReadIC(strOutPut), vbOKOnly, "zlhis"
            lngReturn = JbylReadIC(strOutput)
        
            '�º�����20050310�޸ģ����ݷ��ز�����ʱ�޸ģ���ԭstrOutPut�޸�ΪstrReturn
        
            Select Case lngReturn
            Case 0      '�D�D�ɹ�
            Case Else
                ShowMsgbox "��ȡICʧ��," & vbCrLf & "��������Ϊ:" & lngReturn
                Exit Function
            End Select
            strOutput = Trim(strOutput)
        Case �˳�_�������Ԥ�ֽ�
            lngReturn = Poli_Divide()
            Select Case lngReturn
                Case 0               '�D�D�ɹ�
                Case 1               '�D�DPoli_Divide.in�ļ���ڲ�������
                    ShowMsgbox "Poli_Divide.in�ļ���ڲ�������,����ӿ�����ϵ!"
                    Exit Function
                Case 104             '�D�DPoli_Divide.in��Poli_Divide.out?Poli_Divide.log�ļ��򲻿�
                    ShowMsgbox "Poli_Divide.in��Poli_Divide.out,Poli_Divide.log�ļ��򲻿�,����ӿ�����ϵ!"
                    Exit Function
                Case 105             '�D�Dд���ڲ�������
                    ShowMsgbox "д���ڲ�����������ӿ�����ϵ!"
                    Exit Function
                Case -11               '�D�Dδ����ȷ��Ȩ
                    ShowMsgbox "δ����ȷ��Ȩ������ӿ�����ϵ!"
                    Exit Function
                Case Else
                    ShowMsgbox "����ʧ��," & vbCrLf & "��������Ϊ:" & strReturn
                    Exit Function
            End Select
            strOutput = ""
        Case �˳�_��ͨ����֧��ȷ��
        
            '���ڵ���Reg_Poli()����ʱ������ǿ���˳�����strOutPut��ʼ����һ�ԣ��º�����20050311�޸�
            
            strOutput = Space(200)
            lngReturn = Reg_Poli(strInValue(0), strOutput)
            Select Case lngReturn
                Case 0               '�D�D�ɹ�
                Case 1               '�D�DPoli_Divide.in�ļ���ڲ�������
                    ShowMsgbox "Poli_Divide.in�ļ���ڲ�������,����ӿ�����ϵ!"
                    Exit Function
                Case -11               '�D�Dδ����ȷ��Ȩ
                    ShowMsgbox "δ����ȷ��Ȩ������ӿ�����ϵ!"
                    Exit Function
                Case Else
                    ShowMsgbox "����ʧ��," & vbCrLf & "��������Ϊ:" & strReturn
                    Exit Function
            End Select
            strOutput = Trim(strOutput)
        Case �˳�_�����˷Ѻ���
            lngReturn = PoliBackCost(strInValue(0))
             Select Case lngReturn
                Case 0               '�D�D�ɹ�
                Case -11               '�D�Dδ����ȷ��Ȩ
                   ShowMsgbox "δ����ȷ��Ȩ������ӿ�����ϵ!"
                    Exit Function
                Case Else
                    ShowMsgbox "����ʧ��," & vbCrLf & "��������Ϊ:" & strReturn
                    Exit Function
            End Select
        Case �˳�_סԺ����Ԥ�ֽ�
            lngReturn = Hosp_Divide()
            Select Case lngReturn
                Case 0               '�D�D�ɹ�
                Case 1               '�D�DPoli_Divide.in�ļ���ڲ�������
                    ShowMsgbox "Hosp_Divide.in�ļ���ڲ�������,����ӿ�����ϵ!"
                    Exit Function
                Case 104             '�D�DPoli_Divide.in��Poli_Divide.out?Poli_Divide.log�ļ��򲻿�
                    ShowMsgbox "Hosp_Divide.in ��Hosp_Divide.out,Hosp_Divide.log�ļ��򲻿�,����ӿ�����ϵ!"
                    Exit Function
                Case 105             '�D�Dд���ڲ�������
                    ShowMsgbox "д���ڲ�����������ӿ�����ϵ!"
                    Exit Function
                Case -11               '�D�Dδ����ȷ��Ȩ
                    ShowMsgbox "δ����ȷ��Ȩ������ӿ�����ϵ!"
                    Exit Function
                Case Else
                    ShowMsgbox "����ʧ��," & vbCrLf & "��������Ϊ:" & strReturn
                    Exit Function
            End Select
            strOutput = ""
        Case �˳�_סԺ֧��ȷ��
             
            strOutput = Space(200)
            
       '�º�����20050318�޸ģ���ԭReg_hospital()�����޸�ΪReg_Hospital()�����ڵ��Ӱ����
       
            lngReturn = Reg_Hospital(strInValue(0), strOutput)
           
'            MsgBox strOutPut, vbOKOnly, "zlsoft"
            
            Select Case lngReturn
                Case 0               '�D�D�ɹ�
                Case 1               '�D�DPoli_Divide.in�ļ���ڲ�������
                    ShowMsgbox "Hosp_Divide.in�ļ���ڲ�������,����ӿ�����ϵ!"
                    Exit Function
                Case -11               '�D�Dδ����ȷ��Ȩ
                    ShowMsgbox "δ����ȷ��Ȩ������ӿ�����ϵ!"
                    Exit Function
                Case Else
                    ShowMsgbox "����ʧ��," & vbCrLf & "��������Ϊ:" & strReturn
                    Exit Function
            End Select
            strOutput = Trim(strOutput)
    End Select
    
    strOutPutstring = strOutput
    ҵ������_�˳� = True
    DebugTool "    �������Ϊ:" & strOutPutstring
     Exit Function
errHand:
    DebugTool "    �������Ϊ:" & strOutPutstring
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function ��Ժ�Ǽ�_�˳�(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    Err = 0
    On Error GoTo errHand:
    gcnSQLSEVER_�˳�.BeginTrans
    If InsertIntoData_סԺ�Ǽ�(lng����ID, lng��ҳID) = False Then
        gcnSQLSEVER_�˳�.RollbackTrans
        Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�˳ɺ˹�ҵ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    gcnSQLSEVER_�˳�.CommitTrans
    ��Ժ�Ǽ�_�˳� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle_�˳�.RollbackTrans
    gcnSQLSEVER_�˳�.RollbackTrans
    ��Ժ�Ǽ�_�˳� = False
End Function

Public Function ��Ժ�Ǽǳ���_�˳�(lng����ID As Long, lng��ҳID As Long) As Boolean
  '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
            
    '���˺�:20040923���ӵ�
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
     Err = 0
    On Error GoTo errHand
    
    DebugTool "������Ժ�ǳ����ӿ�"
    
    ��Ժ�Ǽǳ���_�˳� = False
    
    If ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "����δ����ã����ܳ�����Ժ�Ǽ�"
        Exit Function
    End If
    gstrSQL = "Select * from HLD_ZYBRXX where zyh='" & lng����ID & "_" & lng��ҳID & "' and C_zt<>'0'"
    rsTemp.Open gstrSQL, gcnSQLSEVER_�˳�
    If Not rsTemp.EOF Then
        ShowMsgbox "�ò��˵�סԺ��Ϣ�Ѿ��ϴ�����, ������ɾ��!"
        Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�˳ɺ˹�ҵ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
        
    
    'ɾ��SQLServer�е��������
    gstrSQL = " Delete HLD_ZYBRXX where zyh='" & lng����ID & "_" & lng��ҳID & "' and C_zt='0'"
    gcnSQLSEVER_�˳�.Execute gstrSQL
    
    
    
    '����ҽ���ʻ�
    DebugTool "ȡ���ɹ�"
    ��Ժ�Ǽǳ���_�˳� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_�˳�(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    Err = 0:    On Error GoTo errHand:
    ��Ժ�Ǽ�_�˳� = False
    
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "��ǰ���˲�����δ����ã�������Ժ��������"
        Exit Function
    End If
    If frm��Ժ����_�˳�.ShowCard(lng����ID, lng��ҳID) = False Then Exit Function
    
    '�ı䵱ǰ״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�˳ɺ˹�ҵ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_�˳� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ��Ժ�Ǽǳ���_�˳�(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
  '��Ժ�Ǽǳ���
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    
    ��Ժ�Ǽǳ���_�˳� = False
    
    Err = 0: On Error GoTo errHand:
     
     If Not ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "�ò����Ѿ���Ժ������,������ȡ����Ժ!"
        Exit Function
     End If
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�˳ɺ˹�ҵ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_�˳� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�˳�(lng����ID As Long, ByVal lng����ID As Long) As Boolean
  '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)

    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput As String
    Dim rs��ϸ As New ADODB.Recordset
    
    Dim lng��ҳID As Long
    Dim dbl�����ܶ� As Double
    Dim strArr As Variant, strTmpArr As Variant
    
    Dim str���㷽ʽ  As String, strסԺ�� As String
    Dim obj���� As ��������
    Dim dbl�����ʻ� As Double
    
    סԺ����_�˳� = False


    Err = 0: On Error GoTo errHand:
    Call DebugTool("����סԺ����")


    If g�������_�˳�.����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣"
        Exit Function
    End If

    gstrSQL = "Select ��ǰ״̬ From �����ʻ�  where ����ID=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�жϵ�ǰ��סԺ״̬!"

    If Nvl(rsTemp!��ǰ״̬, 0) = 1 Then
        Err.Raise 9000, gstrSysName, "��ǰ���˻�������Ժ״̬,���Ժ���ٽ���!"
        Exit Function
    End If


    With g��������
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
        If IsNull(rsTemp("��ҳID")) = True Then
            Err.Raise 9000, gstrSysName, "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣"
            Exit Function
        End If
        lng��ҳID = rsTemp("��ҳID")
    End With

   '���½����־
    gstrSQL = "Select ID,���ʽ�� as ʵ�ս�� From סԺ���ü�¼ where ����ID=" & lng����ID
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��֤�����־"
    dbl�����ܶ� = 0
    With rs��ϸ
        Do While Not .EOF
                'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                 'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                 gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,NULL)"
                 DebugTool "     ������ϸ��־:SQL=" & gstrSQL
                 zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
                 DebugTool " ������ϸ��־:���²��˷��ü�¼�ɹ�:SQL=" & gstrSQL
                 dbl�����ܶ� = dbl�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
    
    If dbl�����ܶ� <> g��������.���׷����ܶ� Then
        Err.Raise 9000, gstrSysName, "����������ݵķ����ܶ��뱾�ν���ķ����ܶ�ȣ����鴦���Ƿ���ȷ!"
        Exit Function
    End If
    
 
    
    g�������_�˳�.����ID = lng����ID
    '��һ��:סԺ����֧��ȷ��
    '   ������ˮ��|IC����|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|ͳ��֧�����|ͳ���Ը����|�����ʻ�֧�����|�ֽ�֧�����
  '     ������ˮ��|IC����|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|                         |�����ʻ�֧�����|�ֽ�֧�����
   
    StrInput = Rpad(Substr(InitInfor_�˳�.ҽԺ����, 1, 8), 8, " ") & Lpad(Substr(g�������_�˳�.����ID, 1, 12), 12, "0")
    StrInput = StrInput & "|" & g�������_�˳�.IC����
    StrInput = StrInput & "|" & Int(Format((g��������.���׷����ܶ� * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((g��������.ҽ����Χ��� * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((g��������.��ҽ����Χ�� * 100), "####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format((g��������.ͳ��֧����� * 100), "####0.00;-####0.00;0;0"))
    
    '�º�����20050315�޸�,��(g��������.�����ν�� - g��������.ͳ��֧����� * 100)�޸�����
    
    StrInput = StrInput & "|" & Int(Format(((g��������.�����ν�� - g��������.ͳ��֧�����) * 100), "####0.00;-####0.00;0;0"))
    dbl�����ʻ� = ��ȡ�����ʻ�֧��
    StrInput = StrInput & "|" & Int(Format((dbl�����ʻ� * 100), "#####0.00;-####0.00;0;0"))
    StrInput = StrInput & "|" & Int(Format(((g��������.���׷����ܶ� - dbl�����ʻ� - g��������.ͳ��֧�����) * 100), "####0.00;-####0.00;0;0"))
    
    DebugTool "��һ��:סԺ����֧��ȷ��!"
    
    If ҵ������_�˳�(�˳�_סԺ֧��ȷ��, StrInput, strOutput) = False Then
        DebugTool "     סԺ֧��ȷ��ȷ��ʧ��!"
        Exit Function
    End If
    DebugTool "     סԺ֧��ȷ��ȷ�ϳɹ�!"
    
    Err = 0: On Error GoTo ErrHand1:
    gcnOracle_�˳�.BeginTrans
    gcnSQLSEVER_�˳�.BeginTrans
    
    If InsertIntoData_סԺ(strOutput, lng����ID, lng��ҳID) = False Then
        gcnOracle_�˳�.RollbackTrans
        gcnSQLSEVER_�˳�.RollbackTrans
        Exit Function
    End If
    
    '�ڶ���:�ֽ�ȷ�Ϻ���
    '   ������ˮ��|IC����|�ն˻����|��������/ʱ��|�����ܽ��|ҽ�����ܷ���|ҽ�����ܷ���|ͳ��֧�����|ͳ���Ը����|
    '   �����ʻ�֧�����|�ֽ�֧�����|�ۼ�������ʻ����|MAC1
    strArr = Split(strOutput, "|")
    
    
    '��д�����
    Call DebugTool("��д�����¼")
    DebugTool "�ڶ���:��ʼ���汣�ս����¼"
    

   '���뱣�ս����¼
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(),סԺ����_IN(סԺ:��ҳid),����(),�ⶥ��_IN(),ʵ������_IN(�ۼ�������ʻ����),
    '   �������ý��_IN(�����ܽ��),ȫ�Ը����_IN(�ֽ�֧�����),�����Ը����_IN(ҽ�����ܷ���),
    '   ����ͳ����_IN(ҽ�����ܷ���),ͳ�ﱨ�����_IN(����:סԺ:ͳ��֧�����),���Ը����_IN(סԺ:ͳ���Ը����),�����Ը����_IN(),�����ʻ�֧��_IN(�����ʻ�֧�����),"
    '   ֧��˳���_IN(������ˮ��),��ҳID_IN(��ҳid),��;����_IN,��ע_IN(�ն˻����|��������/ʱ��|MAC1)
    
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
    '�º�����20050315�޸�,���ڱ����ȡstrArr(12)�ַ���
     
    strArr(12) = Substr(strArr(12), 1, 16)
    
    '�º�����20050403�����޸ģ���Ϊ����ҽ�����Ĵ��ڳ����Ը����֣��������ҵҽ��ͳ��
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�˳ɺ˹�ҵ & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL,NULL,NULL,null,0,0," & Format(Val(strArr(11)) / 100, "#####0.00;-####0.00;0;0") & "," & _
            Format(Val(strArr(4)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(10)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(6)) / 100, "#####0.00;-####0.00;0;0") & "," & _
            Format(Val(strArr(5)) / 100, "#####0.00;-####0.00;0;0") & " ," & Format(Val(strArr(7)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(Val(strArr(8)) / 100, "#####0.00;-####0.00;0;0") & "," & Format(g��������.�ⶥ�������Ը����, "#####0.00;-####0.00;0;0") & "," & _
            Format(Val(strArr(9)) / 100, "#####0.00;-####0.00;0;0") & ",'" & _
            strArr(0) & " ',NULL,NULL,'" & strArr(2) & "|" & strArr(3) & "|" & strArr(12) & "')"

    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    '---------------------------------------------------------------------------------------------
    '������:����ʱ���ݸ�������
    '   ԭ����ID_IN,�ֽ���ID_IN
    
    gstrSQL = "ZL_ҽ������_סԺȷ��("
    gstrSQL = gstrSQL & g�������_�˳�.����ID & ","
    gstrSQL = gstrSQL & lng����ID & ")"
    
    DebugTool "������:����ʱ���ݸ�������:" & gstrSQL
    
    ExecuteProcedure_�˳� "������ʱ����!"
 
    gcnOracle_�˳�.CommitTrans
    gcnSQLSEVER_�˳�.CommitTrans

    סԺ����_�˳� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
    Exit Function
ErrHand1:
    gcnOracle_�˳�.RollbackTrans
    gcnSQLSEVER_�˳�.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Public Function סԺ�������_�˳�(lng����ID As Long) As Boolean
     '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    Err.Raise 9000, gstrSysName, "��ҽ����֧��סԺ�������,��������ѯ�ӿ���!"
    סԺ�������_�˳� = False
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Public Function �����Ǽ�_�˳�(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
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
    Dim str������¼�� As String, strժҪ As String
    Dim strArr


    �����Ǽ�_�˳� = False


   '�������ŵ��ݵķ�����ϸ
  gstrSQL = "" & _
              "  Select a.�շ�ϸĿID,b.����,b.����" & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,������ҳ C" & _
              "  where A.NO=[1] and A.��¼����=[2] and A.��¼״̬ = [3]" & _
              "        and A.�շ�ϸĿID=B.ID and A.����ID=C.����ID  and A.��ҳID=C.��ҳID And C.����=[4]" & _
              "  Order by A.����ID,A.NO,A.����ʱ��"

    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "������ϸ�ϴ�", str���ݺ�, lng��¼����, lng��¼״̬, TYPE_�˳ɺ˹�ҵ)
    Err = 0:    On Error GoTo errHand:
    With rs��ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            gstrSQL = "Select * From ����֧����Ŀ where ����=[1] and �շ�ϸĿid=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȷ��ҽ��֧����Ŀ", TYPE_�˳ɺ˹�ҵ, CLng(Nvl(!�շ�ϸĿID, 0)))
            If rsTemp.EOF Then
                ShowMsgbox "ע�⣺" & vbCrLf & "   �շ�ϸĿΪ:[" & Nvl(!����) & "]" & Nvl(!����) & " ��δ����ҽ������!"
            End If
            .MoveNext
        Loop
    End With
    �����Ǽ�_�˳� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    
End Function


Private Function Readģ������(ByVal intҵ������ As ҵ������_�˳�, ByVal strInputString As String, ByRef strOutPutstring As String)
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
                                strText = "" & vbTab & "|"
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
'    If InStr(1, strOutPutstring, "@$") <> 0 Then
'        strOutPutstring = Split(strOutPutstring, "@$")(1)
'    End If
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function
Private Sub OpenRecordset_�˳�(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_�˳�, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Function סԺ�������_�˳�(rsExse As Recordset, ByVal lng����ID As Long, Optional bln���ʴ� As Boolean = True) As String
    'rsExse:�ַ���
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    
    Dim lng��ҳID As Long, StrInput As String, strOutput  As String
    Dim strסԺ�� As String, str���㷽ʽ As String, strSQL As String
    Dim lng����id1 As Long
    Dim intMouse As Integer
    
    Dim strArr As Variant

    Err = 0: On Error GoTo errHand:
    
    g�������_�˳�.����ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    intMouse = Screen.MousePointer


    gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)

    If IsNull(rsTemp("��ҳID")) = True Then
        MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    lng��ҳID = rsTemp("��ҳID")

    If bln���ʴ� Then
        Screen.MousePointer = 1
        If ��ݱ�ʶ_�˳�(4, lng����id1) = "" Then
            Screen.MousePointer = intMouse
            סԺ�������_�˳� = ""
            Exit Function
        End If
        Screen.MousePointer = intMouse
        If lng����ID <> lng����id1 Then
            ShowMsgbox "���ǵ�ǰҪ����Ĳ���!"
            Exit Function
        End If
    End If

    
    Screen.MousePointer = vbHourglass
    
    
    strSQL = "" & _
        "   Select A.�շ�ϸĿID,a.����*a.���� as ����,A.ʵ�ս��/(nvl(a.����,1)*nvl(a.����,1)) as ����,a.ʵ�ս�� " & _
        "   From סԺ���ü�¼ A" & _
        "   where ����id=" & lng����ID & " and ��ҳid =" & lng��ҳID & _
        "       and a.��¼״̬<>0 and  A.���ʷ���=1 and nvl(a.ʵ�ս��,0)<>0  And nvl(A.Ӥ����,0)=0"
    
   strSQL = "" & _
    "   Select '' no ,�շ�ϸĿID, sysdate as ����ʱ��,sum(����) as ����,����,sum(ʵ�ս��) as ʵ�ս�� " & _
    "   From (" & strSQL & " ) " & _
    "   group by �շ�ϸĿID,����" & _
    "   having sum(����)<>0"
    
    zlDatabase.OpenRecordset rs��ϸ, strSQL, "��ȡ��ϸ��¼"
    If rs��ϸ.RecordCount = 0 Then
        ShowMsgbox "�ò���δ�����κη���,���ܽ���!"
        Exit Function
    End If
    
    g�������_�˳�.����ID = 0
    g�������_�˳�.����ID = lng����ID
    
    '��һ��:���ܷ���
    DebugTool "סԺ�������,��һ��:���ܷ���"
    g�������_�˳�.�����ܶ� = 0
    
    With rs��ϸ
        If rs��ϸ.RecordCount = 0 Then ShowMsgbox "δ������صķ��ü�¼!": Exit Function
        Do While Not .EOF
            g�������_�˳�.�����ܶ� = g�������_�˳�.�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
        
  
    '�ڶ���:д�����ϸ�ļ�
    'д����ļ�
    DebugTool "סԺ�������,�ڶ���:׼��д�����ϸ�ļ�"
    If WriteINParaFile(rs��ϸ, False, True) = False Then
        DebugTool "         ����ļ���ϸд��ʧ��"
        Exit Function
    End If
    DebugTool "         ���д������ļ���ϸ"
    
    
    '������:סԺ�������-����ֽ�
    DebugTool "סԺ�������,������:סԺ�������ֽ�"
    If ҵ������_�˳�(�˳�_סԺ����Ԥ�ֽ�, "", "") = False Then
        DebugTool "             סԺ�������ֽ�ʧ��"
        Exit Function
    End If
    DebugTool "             סԺ�������ֽ�ɹ�"
    
    '���Ĳ�:�ֽ���س��ν��
    
    DebugTool "סԺ�������,�ڶ���:�ֽ���س��ν��"
    
    If ReadOutParaFile(g��������, False, True) = False Then
        DebugTool "     �ֽ���س��ν��ʧ��!"
        Exit Function
    End If
    DebugTool "     �ֽ���س��ν���ɹ�!"
    
    If Format(g�������_�˳�.�����ܶ�, "#####0.00;-####0.00;0;0") <> Format(g��������.���׷����ܶ�, "#####0.00;-####0.00;0;0") Then
        ShowMsgbox "�����ܶ��,���ܽ���!" & vbCrLf & _
                " HIS�����ܶ�:" & Format(g�������_�˳�.�����ܶ�, "#####0.00;-####0.00; ;") & _
                " ��������ܶ�:" & Format(g��������.���׷����ܶ�, "#####0.00;-####0.00; ;")
        Exit Function
    End If
    
    str���㷽ʽ = ""
    With g��������
        str���㷽ʽ = str���㷽ʽ & "�����ʻ�;" & .���׷����ܶ� - .ͳ��֧����� - .����Ӧ���ܶ� & ";1"
        str���㷽ʽ = str���㷽ʽ & "|ͳ�����;" & .ͳ��֧����� & ";0"
    End With
    סԺ�������_�˳� = str���㷽ʽ
    g�������_�˳�.����ID = lng����ID   '��ʾ�ò����Ѿ��������������
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function Open�м��_�˳�() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    Dim strConn As String
    
    Open�м��_�˳� = False
    Err = 0: On Error Resume Next
        
    '����ҽ���м��
    strServer = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_NAME"), "")
    strUser = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_USERNAME"), "")
    strPass = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_PASSWORD"), "")
    strConn = "dsn=" & strServer & ";uid=" & strUser & ";pwd=" & strPass & ";"
    
     
    Set gcnSQLSEVER_�˳� = New ADODB.Connection
    gcnSQLSEVER_�˳�.Open strConn
    If Err <> 0 Then
        MsgBox "�����û�������������ָ�������޷�ע��," & vbCrLf & "��������е�����Դ�����Ƿ���ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    Err = 0: On Error GoTo errHand:
    
    '���½�����ҽ���������Ĺ�������
    '�м������
    gstrSQL = "select ������,����ֵ from ���ղ��� where  ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�˳ɺ˹�ҵҽ��", TYPE_�˳ɺ˹�ҵ)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "ҽ���û���"
                strUser = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ��������"
                strServer = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Case "ҽ���û�����"
                strPass = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_�˳� = New ADODB.Connection
    If OraDataOpen(gcnOracle_�˳�, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    Open�м��_�˳� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    
End Function
Public Function ҽ������_�˳�(ByVal lng���� As Long, ByVal lngҽ������ As Integer) As Boolean
    '���ܣ� �÷������ڹ����Ӧ�ò���������������ҽ�����ݷ����������Ӵ�
    '���أ��ӿ����óɹ�������true�����򣬷���false
    
    Dim strConn As String
    Dim blnReturn As Boolean
    
    If frmSet�˳�.�������� = False Then
        Exit Function
    End If
  
    If gcnOracle_�˳� Is Nothing And gcnSQLSEVER_�˳� Is Nothing Then
                blnReturn = True
    Else
        If Open�м��_�˳�() Then
                blnReturn = True
        End If
    End If

    InitInfor_�˳�.strPath_Get = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Get"), "C:\xcyb\get")
    InitInfor_�˳�.strPath_Put = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Put"), "C:\xcyb\Put")
    InitInfor_�˳�.strPath_In = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_In"), "C:\xcyb\In")
    InitInfor_�˳�.strPath_Out = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_Out"), "C:\xcyb\Out")
    InitInfor_�˳�.strPath_System = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("strPath_System"), "C:\")
        
    InitInfor_�˳�.ODBC_NAME = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_NAME"), "")
    InitInfor_�˳�.ODBC_UserName = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_USERNAME"), "")
    InitInfor_�˳�.ODBC_PassWord = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("ODBC_PASSWORD"), "")
    
    
    InitInfor_�˳�.���ڶ����� = Val(GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("������"), "1")) = 1
    InitInfor_�˳�.����������� = Val(GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("�����������"), "1")) = 1
    
        
    ҽ������_�˳� = blnReturn
End Function
Public Sub ExecuteProcedure_�˳�(ByVal strCaption As String)
    '���ܣ�ִ��SQL���
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_�˳�.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub



