Attribute VB_Name = "mdl�ɶ��ڽ�"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

Private mblnInit As Boolean
Public Enum ҵ������_�ɶ��ڽ�
    ��������Ϣ_�ڽ� = 0
    ��������_�ڽ�
    ��ȡ�ʻ����_�ڽ�
    ������ϸд��_�ڽ�
    ��������ȷ��_�ڽ�
    ��������ȡ��_�ڽ�
    סԺ�Ǽ�_�ڽ�
    סԺ�����ϴ�_�ڽ�
    סԺ�����ϴ�ȡ��_�ڽ�
    ��Ժ�Ǽ��ϴ�_�ڽ�
    ��Ժ�Ǽ�ȷ��_�ڽ�
    ��ȡ��λǷ�����_�ڽ�
    ��ʼ������_�ڽ�
    ���϶���_�ڽ� '20051020 �¶�
    ����֢�����ϴ�_�ڽ�
End Enum

Private gInitCard As Boolean                '��ʼ���˿���
Private Type InitbaseInfor
    ҽԺ���� As String                      '��ʼҽԺ����
    ���ź�_�ڽ� As Integer
    ������_�ڽ� As Integer                  '0-����,1-��ɭ��˾
    
    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ������������ As Boolean
End Type
Public InitInfor_�ɶ��ڽ� As InitbaseInfor
Private mblnStartTran   As Boolean '�����������
Private Type �������
        ����       As String
        ���˱��   As String
        ���֤��   As String
        ����       As String
        �Ա�       As String
        �������   As String
        ��������   As String
        ��λ����   As String
        ͳ����   As String
        �ƿ�����   As String
        ����Ч��   As String
        ��������   As String
        �ƿ���λ   As String
        ����        As Integer
        �ʻ����    As Double
        ��ְ���    As String
        
        סԺ��ˮ�� As String
        ������� As String
        lng����ID   As Long
        
        �����ܶ�  As Double
        ���㷽ʽ    As String   '���㷽ʽ��
        ���ֱ���    As String
        ��������    As String
        ��Ժ���    As String
        �����ܷ��� As Double
End Type

Private Type ��������
    ������־ As String
    ҽ��������ˮ�� As String
    ҽ���ڷ���   As Double
    ҽ�������   As Double
    ����ҽ��֧��    As Double
    �߶�ҽ��֧��    As Double
    ����Աҽ�Ʋ���  As Double
    �ʻ��������  As Double
    �ʻ�֧��        As Double
    ����֧��        As Double
    �𸶱�׼        As Double
    �����־        As Byte '0-����,1-סԺ
    ����ID          As Long
    ����֧��      As Double '20051020 ����
    ����ӯ��       As Double '20051118 ����
End Type

Public g�������_�ɶ��ڽ� As �������
Public gcnOracle_�ɶ��ڽ� As ADODB.Connection     '�м������
Private g��������   As ��������



'****************************************************************************************************************************************************************************************************************************************
'1 ��ض����������
'****************************************************************************************************************************************************************************************************************************************
'   0-��������Ϣ����(������)
Private Declare Function GetCardInfo_MW Lib "NeijCard.dll" Alias "GetCardInfo" (ByVal lngPort As Long, ByVal strPassWord As String, str���� As String, _
         str���˱�� As String, str���֤�� As String, STR���� As String, str�Ա� As String, _
        str������� As String, str�������� As String, str��λ���� As String, strͳ���� As String, _
        str�ƿ����� As String, str����Ч�� As String, str�������� As String, str�ƿ���λ As String) As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'����:
'����ԭ��:function GetCardInfo(port: integer;UserPassword:PChar; var CardNum,PersonNum,
'                   IDNum,Name,Sex,PersonKind,Birthday,DeptNum,Zone,MAKEDATE,EXPIREDATE,REISSUE,MAKEDEPT: PChar):integer
'����:  a)  Port�����������ΪͨѶ�˿ںţ�0��1��2��3�ֱ������1��2��3��4;����Ϊ��I/O��ַ����0x378�������齫���������ӵ�����1��
'       b)  UserPassword�����������Ϊ�û����룬Ҫ�󳤶�Ϊ6���ַ�����ֻ�ܰ���0��9�����֣�
'       c)  CardNum�����������Ϊ���ţ�����Ϊ10��
'       d)  PersonNum�����������Ϊ���˱�ţ�ҽ����ţ�������Ϊ8��
'       e)  IDNum�����������Ϊ���֤���룬����Ϊ18��
'       f)  Name�����������Ϊ����������Ϊ20��
'       g)  Sex�����������Ϊ�Ա���룬����Ϊ1������'1'Ϊ�У�'2'ΪŮ��
'       h)  PersonKind�����������Ϊ������𣬳���Ϊ1��
'       i)  Birthday�����������Ϊ�������ڣ�����Ϊ8������1982��6��23�ձ�ʾΪ'19820623'��
'       j)  DeptNum�����������Ϊ��λ���룬����Ϊ6��
'       k)  Zone�����������Ϊͳ��������룬����Ϊ1��
'       l)  MAKEDATE�����������Ϊ�ƿ����ڣ�����Ϊ8����ʾ��ʽͬ�������ڣ�
'       m)  EXPIREDATE�����������Ϊ����Ч���ڣ�������Ч��Ϊ99�꣬���ƿ�����Ϊ20021101������Ч����Ϊ21011101��������Ϊ8����ʾ��ʽͬ�������ڣ�
'       n)  REISSUE�����������Ϊ��������������Ϊ2�����磺�״��ƿ�����������Ϊ'00'����һ�β�������������Ϊ'01'���Դ����ƣ�
'       o)  MAKEDEPT�����������Ϊ�ƿ���λ������Ϊ1�����磺'0'��ʾ�ƿ���Ϊ���ݵ�ɭ��˾��
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'   1-��������Ϣ����(������)
Private Declare Function GetCardInfo_KRQ Lib "NeijCard.dll" Alias "GetCardInfo" (ByVal lngPort As Long, str���� As String, _
        str���˱�� As String, str���֤�� As String, STR���� As String, str�Ա� As String, _
        str������� As String, str�������� As String, str��λ���� As String, strͳ���� As String, _
        str�ƿ����� As String, str����Ч�� As String, str�������� As String, str�ƿ���λ As String) As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'������
'˵��:  ����ϸ���,û�����������
'����ԭ��:function GetCardInfoForKRQ(port: integer; var CardNum,PersonNum,IDNum,Name,Sex,PersonKind,Birthday,DeptNum,Zone,
'           MAKEDATE,EXPIREDATE,REISSUE,MAKEDEPT: PChar):integer;
'����:  a)  Port�����������ΪͨѶ�˿ںţ�0��1��2��3�ֱ������1��2��3��4;����Ϊ��I/O��ַ����0x378�������齫���������ӵ�����1��
'       b)  CardNum�����������Ϊ���ţ�����Ϊ10��
'       c)  PersonNum�����������Ϊ���˱�ţ�ҽ����ţ�������Ϊ8��
'       d)  IDNum�����������Ϊ���֤���룬����Ϊ18��
'       e)  Name�����������Ϊ����������Ϊ20��
'       f)  Sex�����������Ϊ�Ա���룬����Ϊ1������'1'Ϊ�У�'2'ΪŮ��
'       g)  PersonKind�����������Ϊ������𣬳���Ϊ1��
'       h)  Birthday�����������Ϊ�������ڣ�����Ϊ8������1982��6��23�ձ�ʾΪ'19820623'��
'       i)  DeptNum�����������Ϊ��λ���룬����Ϊ6��
'       j)  Zone�����������Ϊͳ��������룬����Ϊ1��
'       k)  MAKEDATE�����������Ϊ�ƿ����ڣ�����Ϊ8����ʾ��ʽͬ�������ڣ�
'       l)  EXPIREDATE�����������Ϊ����Ч���ڣ�������Ч��Ϊ99�꣬���ƿ�����Ϊ20021101������Ч����Ϊ21011101��������Ϊ8����ʾ��ʽͬ�������ڣ�
'       m)  REISSUE�����������Ϊ��������������Ϊ2�����磺�״��ƿ�����������Ϊ'00'����һ�β�������������Ϊ'01'���Դ����ƣ�
'       n)  MAKEDEPT�����������Ϊ�ƿ���λ������Ϊ1�����磺'0'��ʾ�ƿ���Ϊ���ݵ�ɭ��˾��
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'   2-�޸�����
Private Declare Function ChangePassword Lib "NeijCard.dll" (ByVal lngPort As Long, ByVal strOldPassWord As String, ByVal strNewPassWord As String) As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'˵��:  ����ϸ���,û�����������
'����ԭ��:function ChangePassword(port:integer;OldPassword,NewPassword:PChar):integer;
'����:  a)  Port�����������ΪͨѶ�˿ںţ�0��1��2��3�ֱ������1��2��3��4;����Ϊ��I/O��ַ����0x378�������齫���������ӵ�����1��
'       b)  OldPassword�����������Ϊԭ���룬Ҫ�󳤶�Ϊ6���ַ�����ֻ�ܰ���0��9�����֣�
'       c)  NewPassword�����������Ϊ�����룬Ҫ�󳤶�Ϊ6���ַ�����ֻ�ܰ���0��9�����֡�
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'****************************************************************************************************************************************************************************************************************************************
'2 ҵ�����
'****************************************************************************************************************************************************************************************************************************************
Public gobj�ɶ��ڽ� As Object
'
'Public gobj�ɶ��ڽ� As New clsNjYh

Public Function ҽ����ʼ��_�ɶ��ڽ�() As Boolean
    
    Dim strReg As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    If mblnInit Then
        ҽ����ʼ��_�ɶ��ڽ� = True
        Exit Function
    End If
    
    GetRegInFor g����ȫ��, "ҽ��", "������", strReg
    InitInfor_�ɶ��ڽ�.������_�ڽ� = Val(strReg)
    
    
    GetRegInFor g����ȫ��, "ҽ��", "���ں�", strReg
    
    InitInfor_�ɶ��ڽ�.���ź�_�ڽ� = IIf(strReg = "", 1, Val(strReg))
        
    '��ʼģ��ӿ�
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        InitInfor_�ɶ��ڽ�.ģ������ = True
    Else
        InitInfor_�ɶ��ڽ�.ģ������ = False
    End If
    
    Call GetRegInFor(g����ģ��, "����", "������������", strReg)
    If Val(strReg) = 1 Then
        InitInfor_�ɶ��ڽ�.������������ = True
    Else
        InitInfor_�ɶ��ڽ�.������������ = False
    End If
    InitInfor_�ɶ��ڽ�.������������ = InitInfor_�ɶ��ڽ�.������������ Or InitInfor_�ɶ��ڽ�.ģ������
    
    
    '����ҽ������
    If gobj�ɶ��ڽ� Is Nothing Then
        Err = 0
        On Error Resume Next
        Set gobj�ɶ��ڽ� = CreateObject("SocketOcxForNC.SocketOcxForNC")
        
        If Err <> 0 Then
                ShowMsgbox "���ܴ���ҽ���ӿ�,����SocketOcxForNC.ocx�Ƿ�����ע��!"
                Exit Function
        End If
    End If
    
    
    'ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & TYPE_�ɶ��ڽ�
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡҽԺ����")
    InitInfor_�ɶ��ڽ�.ҽԺ���� = Nvl(rsTemp!ҽԺ����)
    If Open�м�� = False Then Exit Function
    
    mblnInit = True
    ҽ����ʼ��_�ɶ��ڽ� = True
End Function
Private Function Open�м��() As Boolean
    '�����м��
    '�м������
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strServer As String, strPass As String, strReg As String
    Dim StrInput As String, strOutput As String
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ������ like 'ҽ��%' and ����=" & TYPE_�ɶ��ڽ�
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡ��ز���ֵ")
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
    Set gcnOracle_�ɶ��ڽ� = New ADODB.Connection

    If OraDataOpen(gcnOracle_�ɶ��ڽ�, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��������Ƿ�ͨ����
          
    'GetRegInFor g����ȫ��, "ҽ��", "ConfigFileName", strReg
    'StrInput = strReg
    'GetRegInFor g����ȫ��, "ҽ��", "HostPort", strReg
    'StrInput = StrInput & vbTab & strReg
    'GetRegInFor g����ȫ��, "ҽ��", "IPAddress", strReg
    'StrInput = StrInput & vbTab & strReg
    
    'If ҵ������_�ɶ��ڽ�(��ʼ������_�ڽ�, StrInput, strOutput) = False Then Exit Function
    
    Open�м�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ҽ����ֹ_�ɶ��ڽ�() As Boolean
    '������д�����
    Dim strReg As String
    mblnInit = False
    Err = 0
    On Error Resume Next
    
    Set gobj�ɶ��ڽ� = Nothing
    If gcnOracle_�ɶ��ڽ�.State = 1 Then
        gcnOracle_�ɶ��ڽ�.Close
    End If
    ҽ����ֹ_�ɶ��ڽ� = True
End Function

Public Function ��ݱ�ʶ_�ɶ��ڽ�(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    '���أ��ջ���Ϣ��
    Err = 0
    On Error GoTo errHand:
    ��ݱ�ʶ_�ɶ��ڽ� = frmIdentify�ɶ��ڽ�.GetPatient(bytType, lng����ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_�ɶ��ڽ� = ""
End Function

Public Function �������_�ɶ��ڽ�(ByVal lng����ID As Long) As Currency
    '����: ��ȡ�α����˸����ʻ����
    '����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(�ʻ����,0) as �ʻ���� from �����ʻ� where ����ID='" & lng����ID & "' and ����=" & TYPE_�ɶ��ڽ�
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡ�����ʻ����")
    
    If rsTemp.EOF Then
        �������_�ɶ��ڽ� = 0
    Else
        If rsTemp("�ʻ����") > 0 Then
        �������_�ɶ��ڽ� = rsTemp("�ʻ����")
        Else
        �������_�ɶ��ڽ� = 0
        End If
    End If
End Function

Public Function �����������_�ɶ��ڽ�(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
  
    Dim StrInput As String, strOutput As String
    Dim strArr
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID  As Long
    Dim str����Ա���� As String
    Dim str�շ�ϸĿID As String '���������շ�ϸĿ��ID�������ж��Ƿ�����ظ�����Ŀ
    Err = 0: On Error GoTo errHand:
    
    Call DebugTool("���������������")
    
    With g��������
        .����֧�� = 0
        .������־ = ""
        .�߶�ҽ��֧�� = 0
        .����Աҽ�Ʋ��� = 0
        .�𸶱�׼ = 0
        .ҽ��������ˮ�� = ""
        .ҽ���ڷ��� = 0
        .ҽ������� = 0
        .�ʻ�������� = 0
        .�ʻ�֧�� = 0
        .����֧�� = 0 '20051021 add
    End With
    
    '����Ƿ�����ظ�����Ŀ
    With rs��ϸ
        Do While Not .EOF
            If InStr(1, str�շ�ϸĿID & ",", "," & !�շ�ϸĿID & ",") = 0 Then
                str�շ�ϸĿID = str�շ�ϸĿID & "," & !�շ�ϸĿID
            Else
                Err.Raise 9000, gstrSysName, "�����ظ����շ�ϸĿ����ϲ����ٽ���Ԥ���㣡"
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With

    lng����ID = rs��ϸ("����ID")
    str����Ա���� = Nvl(rs��ϸ!������)
    If g�������_�ɶ��ڽ�.lng����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣"
        Exit Function
    End If
    
    g��������.�����־ = 0
    g��������.����ID = 0
        
    'д����ϸ
    If ������ϸд��(rs��ϸ, True) = False Then Exit Function
    If ���㷽ʽ����(1, str���㷽ʽ) = False Then Exit Function
    
    �����������_�ɶ��ڽ� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
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
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ�֧��", g��������.����ID)
    If Not rsTemp.EOF Then
        ��ȡ�����ʻ�֧�� = Nvl(rsTemp!��Ԥ��, 0)
    End If
    
End Function


Public Function �������_�ɶ��ڽ�(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency, Optional ByRef strAdvance = "") As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim StrInput As String, strOutput As String
    Dim strArr
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim lng����ID  As Long
    Dim str����Ա���� As String
    Err = 0: On Error GoTo errHand:
    
    Call DebugTool("�����������")
    'Modified by ZYB 22051123 ��������������ϴ���ϸ�����㲻�ٽ��кϷ��Լ�鼰��ϸ�ϴ�
    '#################################################################################
'    gstrSQL = "Select �շ�ϸĿID From ���˷��ü�¼  where ����id=" & lng����ID & " group by �շ�ϸĿiD   having Count(�շ�ϸĿid)>=2 "
'    Call OpenRecordset(rsTemp, "�ж���ϸ�Ƿ��ظ�")
'    If Not rsTemp.EOF Then
'        MsgBox "�����ظ����շ�ϸĿ,��ϲ����ٽ���!"
'        Exit Function
'    End If
'
'
'    With g��������
'        .����֧�� = 0
'        .������־ = ""
'        .�߶�ҽ��֧�� = 0
'        .����Աҽ�Ʋ��� = 0
'        .�𸶱�׼ = 0
'        .ҽ��������ˮ�� = ""
'        .ҽ���ڷ��� = 0
'        .ҽ������� = 0
'        .�ʻ�������� = 0
'        .�ʻ�֧�� = 0
'        .����֧�� = 0 '20051021 add
'    End With
'
    '#################################################################################
    
    gstrSQL = "" & _
    "   Select a.*,a.����*a.���� as ����,a.ʵ�ս��/(nvl(a.����,1)*nvl(a.����,1)) as ���� " & _
    "   From ������ü�¼ a " & _
    "   Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Call zlDatabase.OpenRecordset(rs��ϸ, gstrSQL, "��ȡ��ϸ��¼")
    If rs��ϸ.EOF = True Then
        Err.Raise 9000, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If

    lng����ID = rs��ϸ("����ID")
    str����Ա���� = Nvl(rs��ϸ!����Ա����)
    If g�������_�ɶ��ڽ�.lng����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò��˻�û�о������ ��֤�����ܽ���ҽ�����㡣"
        Exit Function
    End If
    
    g��������.�����־ = 0
    g��������.����ID = lng����ID
        
    'д����ϸ
'    If ������ϸд��(rs��ϸ, False) = False Then Exit Function
    If ���㷽ʽ����(1, strAdvance) = False Then Exit Function
    
    '��ȡ�ʻ�֧��
    strAdvance = ""         'ǿ�Ƹ�Ϊ�գ����ߵ����߲���ҪУ��
    g��������.�ʻ�֧�� = ��ȡ�����ʻ�֧��()
    '���ѽ���ȷ��
    '�������: ���˱��    String(8)   In
    '          �籣������  String(10)  In
    '          ҽԺ����    String(5)   In
    '          ����Ա������    String(10)  In
    '          ͳ���������    String(1)   In
    '          ҽ��������ˮ��  String(20)  In
    '          �������    String(1)   In
    '          �����ʻ�֧��    String(10)  In
    With g�������_�ɶ��ڽ�
        StrInput = Rpad(.���˱��, 8)
        StrInput = StrInput & vbTab & Rpad(.����, 10)
        StrInput = StrInput & vbTab & Rpad(InitInfor_�ɶ��ڽ�.ҽԺ����, 5)
        StrInput = StrInput & vbTab & Substr(Rpad(str����Ա����, 10), 1, 10)
        StrInput = StrInput & vbTab & Rpad(.ͳ����, 1)
        StrInput = StrInput & vbTab & Substr(Rpad(g��������.ҽ��������ˮ��, 20), 1, 20)
        StrInput = StrInput & vbTab & Rpad(.�������, 1)
        StrInput = StrInput & vbTab & Lpad(Round(g��������.�ʻ�֧�� * 100), 10, "0")
    End With
    '���ý���
    Call DebugTool("׼��������������ȷ��")
    If ҵ������_�ɶ��ڽ�(��������ȷ��_�ڽ�, StrInput, strOutput) = False Then Exit Function
    Call DebugTool("������������ȷ�Ͻ���")
    
   '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(��),�ʻ��ۼ�֧��_IN(��),�ۼƽ���ͳ��_IN(��),�ۼ�ͳ�ﱨ��_IN(��),סԺ����_IN(��),����(��),�ⶥ��_IN(�ʻ��������),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(��),�����Ը����_IN(����֧��),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(����ʱ������ˮ��),��ҳID_IN,��;����_IN,��ע_IN
    'Modified by ZYB 22051123 ���ϰ汾��������У����������
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�ɶ��ڽ� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "0,0,0,0,0,0," & g��������.�ʻ�������� & ",0," & _
            g�������_�ɶ��ڽ�.�����ܶ� & "," & g��������.ҽ���ڷ��� & "," & g��������.ҽ������� & "," & _
           "0,0,0," & g��������.����֧�� & "," & g��������.�ʻ�֧�� & ",'" & _
            g��������.ҽ��������ˮ�� & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    '---------------------------------------------------------------------------------------------
    �������_�ɶ��ڽ� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function Get������ˮ��(ByVal str�������� As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ������ˮ��
    '--�����:str��������-��YYMMDD��ʽ����
    '--������:
    '--��  ��:������ˮ��
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select ҽԺ������ˮ��_ID.nextval as ���� From dual"
    OpenRecordset_�ɶ��ڽ� rsTemp, "��ȡ������ˮ��"
    Get������ˮ�� = InitInfor_�ɶ��ڽ�.ҽԺ���� & str�������� & Lpad(Nvl(rsTemp!����), 7, "0")
End Function

Private Function ������ϸд��(ByVal rs��ϸ As ADODB.Recordset, Optional ByVal bln���� As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:rs��ϸ-��ϸ��¼
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    
    Dim StrInput As String, strOutput As String, str��ϸ As String
    Dim strInsert As String
    Dim lngSumLen As Long
    
    Dim str��� As String
    Dim str���ձ��� As String
    Dim str������� As String
    Dim str������ˮ�� As String
    Dim lng�������� As Long
    
    Dim strArr
    
    ������ϸд�� = False
    
    DebugTool "����������ϸ�ϴ��ӿ�"
    
    g�������_�ɶ��ڽ�.�����ܶ� = 0
    g�������_�ɶ��ڽ�.�����ܷ��� = 0
       
    
    Err = 0
    On Error GoTo errHand:
    str��ϸ = ""
    'Ȼ����봦����ϸ
    str������ˮ�� = Get������ˮ��(Format(zlDatabase.Currentdate, "yyyymmdd"))
    
    With rs��ϸ
        lng�������� = 0
        Do While Not .EOF
            
            If Val(Nvl(rs��ϸ("ʵ�ս��"), 0)) <> 0 Then
                '������ϸ
                '1   ������Ŀ���� Varchar2(1)    '1'��ҩƷ����   '2'��������Ŀ
                '2   ������Ŀ����    Varchar2(20)    "ҩƷ����"����"������Ŀ����"
                '3   ����    Varchar2(10)    ʵ������*100�ϴ�
                '4   ���    Varchar2(10)    ������2�ֽ�
                '5   �������    Varchar2(10)    ��ҽԺ�ϴ�(��Ҫ��ʲô?)
                
                gstrSQL = "select A.����,A.����,A.���,A.���,A.���㵥λ,B.��Ŀ����,B.��ע,B.�Ƿ�ҽ��,A.���㵥λ,E.���,G.���� ����,B.������� " & _
                          "from �շ�ϸĿ A," & _
                          "         (   Select a.*,b.������� " & _
                          "             From ����֧����Ŀ a,������Ŀ b" & _
                          "             where a.����=b.���� and a.��Ŀ����=b.���� and A.�շ�ϸĿID =[1] and a.����=[2]) B,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
                          "where A.ID=[1] and A.ID=B.�շ�ϸĿID(+) " & _
                         "        AND A.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����Ŀ", CLng(rs��ϸ!�շ�ϸĿID), TYPE_�ɶ��ڽ�)
                
                If rsTemp.EOF Then
                    Err.Raise 9000, gstrSysName, "����δ�������Ŀ,���ڱ�����Ŀ�����н��ж���!"
                    Exit Function
                End If
                
                str��� = Nvl(rsTemp!���)
                str���ձ��� = Nvl(!���ձ���)
                
                '������ձ���Ϊ�գ�����Ҫ�û�ѡ�����
                If str���ձ��� = "" Then
                    str���ձ��� = GetItemInsure_�ɶ��ڽ�(0, !�շ�ϸĿID, True)
                End If
                If str���ձ��� = "" Then str���ձ��� = Nvl(rsTemp!��Ŀ����)
                
                'ȡ�������
                gstrSQL = "Select ������� From ������Ŀ Where ����=" & TYPE_�ɶ��ڽ� & " And ����='" & str���ձ��� & "'"
                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "ȡ�������")
                str������� = Nvl(rsTemp!�������)
                
                str��ϸ = str��ϸ & Substr(Rpad(str�������, 1), 1, 1)
                str��ϸ = str��ϸ & Rpad(str���ձ���, 20)
                str��ϸ = str��ϸ & Lpad(Nvl(!����) * 100, 10, "0")
                str��ϸ = str��ϸ & Rpad(Nvl(str���), 10)
                str��ϸ = str��ϸ & Lpad(Nvl(!ʵ�ս��) * 100, 10, "0")
                                
                'Beging 20051025 add
                '���ﲿ��:A.�����á��������á�(��������Ŀ����.xls��4��)�������á��ƻ���������
                If str���ձ��� = "M10000116" Or str���ձ��� = "M10000117" Or str���ձ��� = "M10000118" Or str���ձ��� = "M10000119" Then
                    Err.Raise 9000, gstrSysName, "���ﲻ��ʹ����������!"
                    Exit Function
                End If
                'End
                
                'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
                'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                'ժҪֵ:ҽԺ������ˮ��
                'Modified by ZYB 20051123 �������ʱδ���洦�����޷�����
                If Not bln���� Then
                    gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL," & Nvl(str�������, "NULL") & ",NUll,'" & str���ձ��� & "',1,'" & str������ˮ�� & "')"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
                End If
                lng�������� = lng�������� + 1
            End If
            
            'Beging 20051027 �¶�
            If Substr(Rpad(str�������, 1), 1, 1) <> "4" Then
                g�������_�ɶ��ڽ�.�����ܶ� = g�������_�ɶ��ڽ�.�����ܶ� + Nvl(rs��ϸ!ʵ�ս��, 0)
            Else
                g�������_�ɶ��ڽ�.�����ܷ��� = g�������_�ɶ��ڽ�.�����ܷ��� + Nvl(rs��ϸ!ʵ�ս��, 0)
            End If
            'End 20051027 �¶�
            rs��ϸ.MoveNext
        Loop

        If lng�������� > 99 Then
            Err.Raise 9000, gstrSysName, "���ﴦ����ϸ���ܴ���99����Ŀ,��ֳ����Ŵ�������¼��!"
            Exit Function
        End If
        
        If .RecordCount <> 0 Then
            .MoveFirst
            '������������˱��    String(8)   In
                          '       �籣������  String(10)  In
                          '       ҽԺ����    String(5)   In
                          '       ����Ա������    String(10)  In
                          '       ͳ���������    String(1)   In
                          '       ҽԺ������ˮ��  String(20)  In
                          '       �������    String(1)   In
                          '       ��������    String(2)   In
                          '       ������ϸ    String����������51  In
            StrInput = Rpad(g�������_�ɶ��ڽ�.���˱��, 8)
            StrInput = StrInput & vbTab & Rpad(g�������_�ɶ��ڽ�.����, 10)
            StrInput = StrInput & vbTab & Rpad(InitInfor_�ɶ��ڽ�.ҽԺ����, 5)
            If bln���� Then
                StrInput = StrInput & vbTab & Rpad(UserInfo.���, 10)
            Else
                StrInput = StrInput & vbTab & Rpad(Nvl(!����Ա���), 10)
            End If
            StrInput = StrInput & vbTab & Rpad(g�������_�ɶ��ڽ�.ͳ����, 1)
            StrInput = StrInput & vbTab & Rpad(str������ˮ��, 20)
            StrInput = StrInput & vbTab & Rpad(g�������_�ɶ��ڽ�.�������, 1)
            StrInput = StrInput & vbTab & Rpad(lng��������, 2)
            StrInput = StrInput & vbTab & str��ϸ
            If ҵ������_�ɶ��ڽ�(������ϸд��_�ڽ�, StrInput, strOutput) = False Then Exit Function
            
            '�����������
            '    ҽԺ��ˮ��_IN IN ҽ��������Ϣ.ҽԺ��ˮ��%TYPE,
            '    ����ID_IN IN ҽ��������Ϣ.����ID%TYPE,
            '    ҽ����ˮ��_IN IN ҽ��������Ϣ.ҽ����ˮ��%TYPE,
            '    ҽ���ڷ���_IN IN ҽ��������Ϣ.ҽ���ڷ���%TYPE,
            '    ҽ�������_IN IN ҽ��������Ϣ.ҽ�������%TYPE,
            '    �ʻ��������_IN IN ҽ��������Ϣ.�ʻ��������%TYPE,
            '    ��ְ���_IN IN ҽ��������Ϣ.��ְ���%TYPE,
            '    ҽ����Ŀ����_IN IN ҽ��������Ϣ.ҽ����Ŀ����%TYPE,
            '    ҽ����Ŀ����_IN IN ҽ��������Ϣ.ҽ����Ŀ����%TYPE,
            '    ҽ���ڷ���1_IN IN ҽ��������Ϣ.ҽ���ڷ���1%TYPE,
            '    �������_IN IN ҽ��������Ϣ.�������%TYPE,
            '    ��Ŀ����_IN IN ҽ��������Ϣ.��Ŀ����%TYPE
            strArr = Split(strOutput, vbTab)
            
            With g��������
                .ҽ��������ˮ�� = strArr(0)
                .ҽ���ڷ��� = Val(strArr(1))
                .ҽ������� = Val(strArr(2))
                .�ʻ�������� = Val(strArr(3))
                .����֧�� = Val(strArr(6)) '20051020 Add
            End With
            strInsert = "ZL_ҽ��������Ϣ_INSERT("
            strInsert = strInsert & "'" & str������ˮ�� & "',"
            strInsert = strInsert & "" & g�������_�ɶ��ڽ�.lng����ID & ","
            strInsert = strInsert & "'" & strArr(0) & "',"
            strInsert = strInsert & "" & Val(strArr(1)) & ","
            strInsert = strInsert & "" & Val(strArr(2)) & ","
            strInsert = strInsert & "" & Val(strArr(3)) & ","
            strInsert = strInsert & "'" & strArr(5) & "',"
            
            
            '�ֽ���ϸ��¼��¼
                        
            '1   ������Ŀ���� Varchar2(1)    '1'��ҩƷ����   '2'��������Ŀ
            '2   ������Ŀ����    Varchar2(20)    "ҩƷ����"����"������Ŀ����"
            '3   ҽ���ڷ���  Varchar2(10)    ʵ������*100
            '4   ������� Varchar2(10)(��ҩ?��ҩ?�����)
            '5   ��Ŀ����    Varchar2(10)    ʵ������*100
            '6   ����֧�� 20051026
            str��ϸ = strArr(4)
            lngSumLen = zlCommFun.ActualLen(str��ϸ)
            StrInput = ""
            Dim r As Long, i As Integer
            
            For i = 1 To lngSumLen Step 51
                r = i
                StrInput = StrInput & "'" & Substr(str��ϸ, r, 1) & "',"
                r = r + 1
                StrInput = StrInput & "'" & Substr(str��ϸ, r, 20) & "',"
                r = r + 20
                StrInput = StrInput & "" & Val(Substr(str��ϸ, r, 10)) & ","
                r = r + 10
                StrInput = StrInput & "'" & Substr(str��ϸ, r, 10) & "',"
                r = r + 10
                StrInput = StrInput & "" & Val(Substr(str��ϸ, r, 10)) & ","
                StrInput = StrInput & "" & Val(strArr(6)) & ")"
                '���SQL���
                gstrSQL = strInsert & StrInput
                
                'Modified by zyb 20051123 ���ܻ����HIS���д�����Ч����ϸ��������Ԥ���㣩������Ҳһ��
                ExecuteProcedure_ZLNJ "������ϸ���ݵ��м��"
                StrInput = ""
            Next
        End If
    End With
    ������ϸд�� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_�ɶ��ڽ�(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput  As String, str��ˮ�� As String
    Dim lng����ID As Long, lng����id1 As Long
    Dim strArr
    Dim rs��ϸ As New ADODB.Recordset
    Dim i As Long
    Dim intMouse  As Integer
 
    On Error GoTo errHand:

    DebugTool "�����������"
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    '�����֤
    'by 20050123 gzy
    lng����id1 = lng����ID
    If ��ݱ�ʶ_�ɶ��ڽ�(2, lng����id1) = "" Then
        Screen.MousePointer = intMouse
        ����������_�ɶ��ڽ� = False
        If lng����id1 = 0 Then
            Exit Function
        End If
    End If
    Screen.MousePointer = intMouse
    
    DebugTool "�����֤���"
    
    gstrSQL = "Select * From ������ü�¼  " & _
        " Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Call zlDatabase.OpenRecordset(rs��ϸ, gstrSQL, "��ȡ������¼")
    
    
    
    g�������_�ɶ��ڽ�.�����ܶ� = 0
    Do Until rs��ϸ.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")
        str��ˮ�� = Nvl(rs��ϸ!ժҪ)
        g�������_�ɶ��ڽ�.�����ܶ� = g�������_�ɶ��ڽ�.�����ܶ� + Nvl(rs��ϸ("���ʽ��"), 0)
        rs��ϸ.MoveNext
    Loop
    g�������_�ɶ��ڽ�.�����ܶ� = Round(g�������_�ɶ��ڽ�.�����ܶ�, 2)
    
    If lng����ID <> lng����id1 Then
        Err.Raise 9000, gstrSysName, " �鿨���˲��ǵ�ǰҪ�����Ĳ���,���ܳ�������"
        Exit Function
    End If
    
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "����ҽ��")
    lng����ID = rsTemp("����ID")

    

    gstrSQL = "Select * From ������ü�¼ " & _
        " Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡ������¼")
    
    DebugTool "����ժҪ��־"
    Do While Not rsTemp.EOF
        '�����ϴ���־

        gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(rsTemp!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str��ˮ�� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
        rsTemp.MoveNext
    Loop
    DebugTool "����ժҪ��־���"

    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=" & TYPE_�ɶ��ڽ� & " and ��¼ID=" & lng����ID
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡԭ���Ľ����¼")

    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    g��������.ҽ��������ˮ�� = rsTemp("֧��˳���")
    
    '    ���˱��    String(8)   In
    '    �籣������  String(10)  In
    '    ҽԺ����    String(5)   In
    '    ����Ա������    String(10)  In
    '    ͳ���������    String(1)   In
    '    ҽ��������ˮ��  String(20)  In
    '    �������    String(1)   In
    With g�������_�ɶ��ڽ�
        StrInput = Rpad(.���˱��, 8)
        StrInput = StrInput & vbTab & Rpad(.����, 10)
        StrInput = StrInput & vbTab & Rpad(InitInfor_�ɶ��ڽ�.ҽԺ����, 5)
        StrInput = StrInput & vbTab & Rpad(gstrUserName, 10)
        StrInput = StrInput & vbTab & Rpad(.ͳ����, 1)
        StrInput = StrInput & vbTab & Rpad(g��������.ҽ��������ˮ��, 20)
        StrInput = StrInput & vbTab & Rpad(.�������, 1)
    End With
    If ҵ������_�ɶ��ڽ�(��������ȡ��_�ڽ�, StrInput, strOutput) = False Then Exit Function
    DebugTool "ҵ������ɹ�"

 '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN(�ʻ��������),ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(��),�ʻ��ۼ�֧��_IN(��),�ۼƽ���ͳ��_IN(��),�ۼ�ͳ�ﱨ��_IN(��),סԺ����_IN(��),����(��),�ⶥ��_IN(�ʻ��������),ʵ������_IN(��),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(��),�����Ը����_IN(��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(��),�����Ը����_IN(��),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(����ʱ������ˮ��),��ҳID_IN,��;����_IN,��ע_IN
    DebugTool "��������¼"
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�ɶ��ڽ� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
        "0,0,0,0,0,0,0," & -1 * Nvl(rsTemp!�ⶥ��, 0) & "," & _
        Nvl(rsTemp!�������ý��, 0) * -1 & "," & Nvl(rsTemp!ȫ�Ը����, 0) * -1 & "," & Nvl(rsTemp!�����Ը����, 0) * -1 & "," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0,0," & rsTemp("�����ʻ�֧��") * -1 & ",'" & _
       g��������.ҽ��������ˮ�� & "',NULL,0,null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���±��ս�����Ϣ")
    DebugTool "�������������"
    
    ����������_�ɶ��ڽ� = True
    Exit Function
errHand::
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Private Function ������Ժ�ǼǴ���(lng����ID As Long, lng��ҳID As Long) As Boolean
    '��������Ǽ�
    Dim StrInput As String, strOutput As String
    Dim str������ˮ�� As String
    Dim rsTemp As New ADODB.Recordset, rsRydj As New ADODB.Recordset
    Dim strArr
    Err = 0
    On Error GoTo errHand:
    
    gstrSQL = "Select C.סԺ��,C.��ǰ����,to_char(A.ȷ������,'yyyy-MM-dd') as ȷ������,A.�Ǽ��� ������,B.λ�� ��Ժ����,A.סԺҽʦ,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') ��Ժ����ʱ��," & _
        " to_char(A.��Ժ����,'yyyymmdd') ��Ժ����  ,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') ��Ժʱ��,D.��Ժ��ϱ���,D.��Ժ�������,G.ȷ����ϱ���,g.ȷ��������� " & _
        " From ������ҳ A,���ű� B,������Ϣ C, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,1,b.����,'')) AS ��Ժ��ϱ���,max(DECODE(a.��ϴ���,1,b.����,'')) AS ��Ժ������� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =1 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by  ����id,��ҳid)   D," & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,2,b.����,'')) AS ȷ����ϱ���,max(DECODE(a.��ϴ���,2,b.����,'')) AS ȷ��������� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =1 and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by  ����id,��ҳid)   g" & _
        " Where A.����id=C.����id and C.����id=" & lng����ID & _
        "       and A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID " & _
        "       and A.��ҳid=D.��ҳid(+) and a.����id=D.����id(+) " & _
        "       and A.��ҳid=g.��ҳid(+) and a.����id=g.����id(+) " & _
        ""

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ��Ϣ", lng����ID, lng��ҳID)

    With g�������_�ɶ��ڽ�
        '�������
        '    ���˱��    String(8)   In
        '    �籣������  String(10)  In
        '    ҽԺ����    String(5)   In
        '    ����Ա������    String(10)  In
        '    ͳ���������    String(1)   In
        '    ��Ժ����    String(8)   In
        '    ��Ժ�Ʊ�    String(10)  In
        '    ��Ժ����ҽ��    String(10)  In
        '    ��ϱ���    String(20)  In 20051026 �޸�Ϊ string(200)
        
        'Beging 20051026 �¶�
        Dim vat����֢ As Variant, str�������� As String, i As Long
        
        gstrSQL = "Select * from �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_�ɶ��ڽ�
        Call zlDatabase.OpenRecordset(rsRydj, gstrSQL, "ȡ��������")
        str�������� = Nvl(rsRydj!��������)
        If str�������� <> "" Then
            If InStr(str��������, "|") > 0 Then
                vat����֢ = Split(str��������, "|")
                str�������� = ""
                For i = 0 To UBound(vat����֢) - 1
                    str�������� = str�������� & Rpad(Substr(vat����֢(i), 1, 20), 20)
                Next
            End If
        Else
            str�������� = Space(180)
        End If
        'End 20051026 �¶�
        StrInput = Rpad(.���˱��, 8)
        StrInput = StrInput & vbTab & Rpad(.����, 10)
        StrInput = StrInput & vbTab & Rpad(InitInfor_�ɶ��ڽ�.ҽԺ����, 5)
        StrInput = StrInput & vbTab & Rpad(gstrUserName, 10)
        StrInput = StrInput & vbTab & Rpad(.ͳ����, 1)
        StrInput = StrInput & vbTab & Rpad(rsTemp!��Ժ����, 8)
        StrInput = StrInput & vbTab & Rpad(Substr(Nvl(rsTemp!��Ժ����), 1, 10), 10)
        StrInput = StrInput & vbTab & Rpad(Substr(Nvl(rsTemp!סԺҽʦ), 1, 10), 10)
        StrInput = StrInput & vbTab & Rpad(Rpad(Substr(g�������_�ɶ��ڽ�.���ֱ���, 1, 20), 20) & Substr(str��������, 1, 180), 200)
                
        If ҵ������_�ɶ��ڽ�(סԺ�Ǽ�_�ڽ�, StrInput, strOutput) = False Then
            Exit Function
        End If
        
        '�������
        '    סԺ��ˮ��  String(20)  Out
        '    ���ܴ�����־    Small int   Out
        '    �𸶱�׼    Long    Out
        '    ��ְ���    String(1)   Out
        
        strArr = Split(strOutput, vbTab)

        '���潫������ˮ��
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�ɶ��ڽ� & ",'˳���','''" & strArr(0) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���潻����ˮ��")
        '�������ܴ�����־
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�ɶ��ڽ� & ",'���ܴ�����־','''" & Val(strArr(1)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������ܴ�����־")
        '�����𸶱�׼
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�ɶ��ڽ� & ",'�𸶱�׼','''" & Val(strArr(2)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����𸶱�׼")
    End With

    ������Ժ�ǼǴ��� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_�ɶ��ڽ�(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutput As String, StrInput As String
    
    '��ȡסԺ��
    Err = 0
    On Error GoTo errHand:
 
    '�жϵ�λǷ�����
    '    ���˱��    String (8)  IN
    '    �籣������  String (10) IN
    '    ͳ���������    String (1)  IN
    StrInput = g�������_�ɶ��ڽ�.���˱��
    StrInput = StrInput & vbTab & g�������_�ɶ��ڽ�.����
    StrInput = StrInput & vbTab & g�������_�ɶ��ڽ�.ͳ����
    
    If ҵ������_�ɶ��ڽ�(��ȡ��λǷ�����_�ڽ�, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    If Val(strOutput) <> 0 Then
        ShowMsgbox "ע�⣺" & vbCrLf & "    ��λ�Ѿ�Ƿ��!"
        'Exit Function
    End If
    
    '�Ƚ��еǼǴ���
    If ������Ժ�ǼǴ���(lng����ID, lng��ҳID) = False Then
        'Exit Function
    End If
    


    '�����˵�״̬�����޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ɶ��ڽ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_�ɶ��ڽ� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_�ɶ��ڽ� = False
End Function

Public Function ��Ժ�Ǽǳ���_�ɶ��ڽ�(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false

    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strҽ����  As String
    Dim str��Ժ���� As String

    Err = 0
    On Error GoTo errHand
    ShowMsgbox "��ҽ���ӿڲ�֧����Ժ�Ǽǳ���,ֻ�ܰ����Ժ"
    Exit Function
   ��Ժ�Ǽǳ���_�ɶ��ڽ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function ��Ժ�Ǽ�_�ɶ��ڽ�(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    '����״̬���޸�
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ɶ��ڽ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_�ɶ��ڽ� = True
    Exit Function
errHand::
    If ErrCenter() = 1 Then
        Resume
    End If
    ��Ժ�Ǽ�_�ɶ��ڽ� = False
End Function
Public Function ��Ժ�Ǽǳ���_�ɶ��ڽ�(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '��Ժ�Ǽǳ���
     '�ı䲡��״̬
     If Not ����δ�����(lng����ID, lng��ҳID) Then
            ShowMsgbox "�ò����Ѿ���Ժ������,���ܳ�Ժ�Ǽǳ���!"
            Exit Function
     End If
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ɶ��ڽ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_�ɶ��ڽ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function סԺ����_�ɶ��ڽ�(lng����ID As Long, ByVal lng����ID As Long) As Boolean
  '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)

    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput As String

    Dim str����Ա As String
    Dim lng��ҳID As Long
    Dim strArr
    Dim lng��Ժ���� As Long, lng��Ժ���� As Long
    
    Dim i As Integer

    If g�������_�ɶ��ڽ�.lng����ID <> lng����ID Then
        MsgBox "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣", vbInformation, gstrSysName
        Exit Function
    End If
    gstrSQL = "Select Sum(nvl(���ʽ��,0)) as �ܶ� from סԺ���ü�¼ where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "�����ܶ�"
    
    If g�������_�ɶ��ڽ�.�����ܶ� + g�������_�ɶ��ڽ�.�����ܷ��� <> Nvl(rsTemp!�ܶ�, 0) Then
        'Modified by ZYB 20051118 ��ʾ����ʾ����Ϣ����ȷ
        Err.Raise 9000, gstrSysName, "���������ܶ���ڱ��ν����ܶ�:" & vbCrLf & "������ܶ�:" & Format(g�������_�ɶ��ڽ�.�����ܶ� + g�������_�ɶ��ڽ�.�����ܷ���, "#####0.00;-####0.00;0;0") & vbCrLf & "��ǰ�����ܶ�Ϊ:" & Format(Nvl(rsTemp!�ܶ�, 0), "#####0.00;-####0.00;0;0")
        Exit Function
    End If

    Err = 0: On Error GoTo errHand:
    Call DebugTool("����סԺ����")


    With g��������
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=" & lng����ID
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�������")
        If IsNull(rsTemp("��ҳID")) = True Then
            Err.Raise 9000, gstrSysName, "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        lng��ҳID = rsTemp("��ҳID")
    End With
    
'   gstrSQL = "Select A.ID From ���˷��ü�¼ a,ҩƷ�շ���¼ B where A.no=b.No and B.���� in (9,10) and a.id=b.����ID and a.����ID=" & lng����ID & " and b.���� like '_3%' and rownum<=2"
'
'
'    Dim bln��Ժ��ҩ As Boolean
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����Ժ��ҩ"
'    If rsTemp.EOF Then
'        bln��Ժ��ҩ = False
'    Else
'        bln��Ժ��ҩ = True
'    End If
'
'  gstrSQL = "Select c.סԺ��,A.�Ǽ��� ������,B.���� ��Ժ����,A.סԺҽʦ,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') ��Ժ����ʱ��," & _
'        " to_char(A.��Ժ����,'yyyyMMdd') ��Ժ����,J.��ֹʱ��,J.����Ա,D.��ϱ���,A.��Ժ��ʽ,to_Char(a.��Ժ����,'yyyyMMDD') as ��Ժ����,a.��Ժ����,H.���� as ��Ժ����" & _
'        " From ������ҳ A,���ű� B,������Ϣ C,���ű� H, " & _
'        "       (Select ����id,��ҳid max(DECODE(a.��ϴ���,2,b.����,'')) AS ��ϱ��� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =3  and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by ����id,��ҳid)   D" & _
'        " Where A.����id=C.����id and C.����id=" & lng����ID & _
'        "       and A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID and A.��Ժ����ID=H.id(+) " & _
'        "       and A.��ҳid=D.��ҳid(+) and a.����id=D.����id(+) " & _
'        ""
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ��������"
'
'
'
'    '���:
'    '    ���˱��    String(8)   In
'    '    �籣������  String(10)  In
'    '    ҽԺ����    String(5)   In
'    '    ����Ա������    String(10)  In
'    '    ͳ���������    String(1)   In
'    '    ��Ժ����    String(8)   In
'    '    ��Ժ�Ʊ�    String(10)  In
'    '    ��Ժ����ҽ��    String(10)  In
'    '    ��ϱ���    String(20)  In
'    '    ��Ժ��ҩ    String(1)   In
'    '    ��Ժ���    String(1)   In
'    '    סԺ��ˮ��  String(20)  In
'
'    With g�������_�ɶ��ڽ�
'        strInput = Rpad(.���˱��, 8)
'        strInput = strInput & vbTab & Rpad(.����, 10)
'        strInput = strInput & vbTab & Rpad(InitInfor_�ɶ��ڽ�.ҽԺ����, 5)
'        strInput = strInput & vbTab & Rpad(Nvl(rsTemp!����Ա), 10)
'        strInput = strInput & vbTab & Rpad(.ͳ����, 1)
'        strInput = strInput & vbTab & Rpad(Nvl(rsTemp!��Ժ����), 8)
'        strInput = strInput & vbTab & Substr(Rpad(Nvl(rsTemp!��Ժ����), 10), 1, 10)
'        strInput = strInput & vbTab & Substr(Rpad(Nvl(rsTemp!סԺҽʦ), 10), 1, 10)
'        strInput = strInput & vbTab & Substr(Rpad(Nvl(rsTemp!��ϱ���), 20), 1, 20)
'        strInput = strInput & vbTab & IIf(bln��Ժ��ҩ, "1", "0")
'        strInput = strInput & vbTab & Substr(Nvl(rsTemp!��Ժ��ʽ), 1, 1)
'        strInput = strInput & vbTab & Substr(Rpad(.סԺ��ˮ��, 20), 1, 20)
'        If ҵ������_�ɶ��ڽ�(��Ժ�Ǽ��ϴ�_�ڽ�, strInput, stroutput) = False Then Exit Function
'    End With
'    If stroutput = "" Then Exit Function
'    strArr = Split(stroutput, vbTab)
'
'   '����
'    '    TRANSDETIAL��� (���������ϸ)
'    '    ���ܴ�����־    String(1)   Out
'    '    ҽ���ڷ���  String(10)  Out
'    '    ҽ�������  String(10)  Out
'    '    ����ҽ��֧��
'    '    ����μӴ�ҽ������Ϊ��ҽ��֧��  String(10)  Out
'    '    �߶�ҽ��֧��    String(10)  Out
'    '    ����Աҽ�Ʋ���  String(10)  Out
'    '    ���˰�����֧��  String(10)  Out
'    '    TRANSDETIAL����
'    '    �𸶱�׼    String(10)  Out
'    '    �����ʻ��������    String(10)  Out
'    stroutput = strArr(0)
'    With g��������
'        .�����־ = 1
'        .������־ = Substr(stroutput, 1, 1)
'        .ҽ���ڷ��� = Val(Substr(stroutput, 2, 10))
'        .ҽ������� = Val(Substr(stroutput, 12, 10))
'        .����ҽ��֧�� = Val(Substr(stroutput, 22, 10))
'        .�߶�ҽ��֧�� = Val(Substr(stroutput, 32, 10))
'        .����Աҽ�Ʋ��� = Val(Substr(stroutput, 42, 10))
'        .����֧�� = Val(Substr(stroutput, 52, 10))
'        .�𸶱�׼ = Val(strArr(1))
'        .�ʻ�������� = Val(strArr(2))
'
'    End With
'
'    If ���㷽ʽ����(2) = False Then Exit Function
        
     '��ȡ�ʻ�֧��
    g��������.����ID = lng����ID
    g��������.�ʻ�֧�� = ��ȡ�����ʻ�֧��() * 100
    '    ���˱��    String(8)   In
    '    �籣������  String(10)  In
    '    ����Ա������    String(10)  In
    '    ͳ��������    String(1)   In
    '    סԺ��ˮ��  String(20)  In
    '    �����ʻ�֧��    String(10)  In

    With g�������_�ɶ��ڽ�
        StrInput = Rpad(.���˱��, 8)
        StrInput = StrInput & vbTab & Rpad(.����, 10)
        StrInput = StrInput & vbTab & Rpad(Substr(gstrUserName, 1, 10), 10)
        StrInput = StrInput & vbTab & Rpad(.ͳ����, 1)
        StrInput = StrInput & vbTab & Rpad(Substr(Rpad(.סԺ��ˮ��, 20), 1, 20), 20)
        StrInput = StrInput & vbTab & Lpad(Substr(Lpad(g��������.�ʻ�֧��, 10), 1, 10), 10, 0)
    End With
    If ҵ������_�ɶ��ڽ�(��Ժ�Ǽ�ȷ��_�ڽ�, StrInput, strOutput) = False Then
        Exit Function
    End If


  '���뱣�ս����¼
    'ԭ���̲���:
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
    '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
    '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    '��ֵ����
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(��),�ʻ��ۼ�֧��_IN(��),�ۼƽ���ͳ��_IN(��),�ۼ�ͳ�ﱨ��_IN(����֧��),סԺ����_IN(��),����(����֧��),�ⶥ��_IN(�ʻ��������),ʵ������_IN(�𸶱�׼),
    '   �������ý��_IN(�����ܶ�),ȫ�Ը����_IN(ҽ���ڷ���),�����Ը����_IN(ҽ�������),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(�߶�ҽ��֧��),�����Ը����_IN(����Աҽ�Ʋ���),�����ʻ�֧��_IN(�����ʻ�֧��),"
    '   ֧��˳���_IN(����ʱ������ˮ��),��ҳID_IN,��;����_IN,��ע_IN(���ܴ�����־)
    
    DebugTool "���㽻���ύ�ɹ�,����ʼ���汣�ս����¼"
    'Bgeing �¶� 20050601
    gstrSQL = "Select * from �����ʻ� where ����ID=" & lng����ID & " and ����=" & TYPE_�ɶ��ڽ�
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "���Ժ����")
    lng��Ժ���� = Nvl(rsTemp!����ID, 0)
    lng��Ժ���� = Nvl(rsTemp!��Ժ����ID, 0)
    
    'beging 20051026 �¶�
    Dim str��Ժ�������� As String, str��Ժ�������� As String
    str��Ժ�������� = Nvl(rsTemp!��������)
    str��Ժ�������� = Nvl(rsTemp!��Ժ��������)
    'end '20051026 �¶�
    
    'End  �¶� 20050601
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�ɶ��ڽ� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "0,0,0," & g��������.����֧�� & ",0," & g��������.����֧�� & "," & g��������.�ʻ�������� & "," & g��������.�𸶱�׼ & "," & _
            g�������_�ɶ��ڽ�.�����ܶ� & "," & g��������.ҽ���ڷ��� & "," & g��������.ҽ������� & "," & _
           g��������.����ҽ��֧�� & "," & g��������.����ҽ��֧�� & "," & g��������.�߶�ҽ��֧�� & "," & g��������.����Աҽ�Ʋ��� & "," & g��������.�ʻ�֧�� / 100 & ",'" & _
            g�������_�ɶ��ڽ�.סԺ��ˮ�� & "',NULL,NULL,'" & g��������.������־ & "')"
            

    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    '---------------------------------------------------------------------------------------------
    'beging �¶� 20050601
    gstrSQL = "update ���ս����¼ Set ����ID=" & lng��Ժ���� & ",��Ժ����ID=" & lng��Ժ���� & _
             " where ��¼ID=" & lng����ID & " And ����=2 and ����=" & TYPE_�ɶ��ڽ�
             
    gcnOracle.Execute gstrSQL
    'end
    'beging 20051026 �¶�
    If str��Ժ�������� <> "" Then
        gstrSQL = "Update ���ս����¼ Set ��������='" & str��Ժ�������� & "'" & _
                " Where ��¼ID=" & lng����ID & " And ����=2 and ����=" & TYPE_�ɶ��ڽ�
        gcnOracle.Execute gstrSQL
    End If
    If str��Ժ�������� <> "" Then
        gstrSQL = "Update ���ս����¼ Set ��Ժ��������='" & str��Ժ�������� & "'" & _
                " Where ��¼ID=" & lng����ID & " And ����=2 and ����=" & TYPE_�ɶ��ڽ�
        gcnOracle.Execute gstrSQL
    End If
    'end '20051026 �¶�
    סԺ����_�ɶ��ڽ� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Public Function סԺ�������_�ɶ��ڽ�(lng����ID As Long) As Boolean
     '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------

    Err = 0: On Error GoTo errHand:
    Err.Raise 9000, gstrSysName, "��ҽ���Ӳ�֧�ַ�����"
    סԺ�������_�ɶ��ڽ� = False
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function
Private Function ������ϸ������м��(ByVal str������ˮ�� As String, ByVal strOutput As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ϸ������м��
    '--�����:��vbtab����
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strHead As String
    Dim StrInput As String, str��ϸ As String
    Dim strArr
    Dim lngSumLen As Long
    Dim r As Long, i As Integer
    'strOutPut: ҽ��������ˮ��  String(20)  Out
    '           ������ϸ    String����������51  Out
    '           TRANSDETIAL��� (���������ϸ) Out
    
    'TRANSDETIAL���
    '        ���ܴ�����־    String(1)   Out
    '        ҽ���ڷ���  String(10)  Out
    '        ҽ�������  String(10)  Out
    '        ����ҽ��֧�� ����μӴ�ҽ������Ϊ��ҽ��֧��  String(10)  Out
    '        �߶�ҽ��֧��    String(10)  Out
    '        ����Աҽ�Ʋ���  String(10)  Out
    '        ���˰�����֧��  String(10)  Out


    ������ϸ������м�� = False
    Err = 0: On Error GoTo errHand:
    strArr = Split(strOutput, vbTab)
    
    '���̲���
    '    ҽԺ��ˮ��_IN IN ҽ��������Ϣ.ҽԺ��ˮ��%TYPE,
    '    ����ID_IN IN ҽ��������Ϣ.����ID%TYPE,
    '    ҽ����ˮ��_IN IN ҽ��������Ϣ.ҽ����ˮ��%TYPE,
    '    ҽ���ڷ���_IN IN ҽ��������Ϣ.ҽ���ڷ���%TYPE,
    '    ҽ�������_IN IN ҽ��������Ϣ.ҽ�������%TYPE,
    '    �ʻ��������_IN IN ҽ��������Ϣ.�ʻ��������%TYPE:=NULL,
    '    ��ְ���_IN IN ҽ��������Ϣ.��ְ���%TYPE:=NULL,
    '    ҽ����Ŀ����_IN IN ҽ��������Ϣ.ҽ����Ŀ����%TYPE:=NULL,
    '    ҽ����Ŀ����_IN IN ҽ��������Ϣ.ҽ����Ŀ����%TYPE,
    '    ҽ���ڷ���1_IN IN ҽ��������Ϣ.ҽ���ڷ���1%TYPE,
    '    �������_IN IN ҽ��������Ϣ.�������%TYPE:=NULL,
    '    ��Ŀ����_IN IN ҽ��������Ϣ.��Ŀ����%TYPE:=NULL,
    '    ���ܴ�����־_IN IN ҽ��������Ϣ.���ܴ�����־%TYPE:=NULL,
    '    ����ҽ��֧��_IN IN ҽ��������Ϣ.����ҽ��֧��%TYPE:=NULL,
    '    �߶�ҽ��֧��_IN IN ҽ��������Ϣ.�߶�ҽ��֧��%TYPE:=NULL,
    '    ����Ա����_IN IN ҽ��������Ϣ.����Ա����%TYPE:=NULL,
    '    ���˱���֧��_IN IN ҽ��������Ϣ.���˱���֧��%TYPE:=NULL
    '    20051021 Add
    '    ����֧��_IN IN ҽ��������Ϣ.����֧��%TYPE:NULL
    strHead = "ZL_ҽ��������Ϣ_INSERT("
    strHead = strHead & "'" & str������ˮ�� & "',"
    strHead = strHead & "" & g�������_�ɶ��ڽ�.lng����ID & ","
    strHead = strHead & "'" & strArr(0) & "',"
    
    strHead = strHead & "" & Val(Substr(strArr(2), 2, 10)) & ","
    strHead = strHead & "" & Val(Substr(strArr(2), 12, 10)) & ","
    strHead = strHead & "" & 0 & ","
    strHead = strHead & "null,"
    
    str��ϸ = strArr(1)
    lngSumLen = zlCommFun.ActualLen(strArr(1))
    StrInput = ""
    For i = 1 To lngSumLen Step 51
     '    ��ְ���_IN IN ҽ��������Ϣ.��ְ���%TYPE:=NULL,
    '    ҽ����Ŀ����_IN IN ҽ��������Ϣ.ҽ����Ŀ����%TYPE:=NULL,
    '    ҽ����Ŀ����_IN IN ҽ��������Ϣ.ҽ����Ŀ����%TYPE,
    '    ҽ���ڷ���1_IN IN ҽ��������Ϣ.ҽ���ڷ���1%TYPE,
    '    �������_IN IN ҽ��������Ϣ.�������%TYPE:=NULL,
    '    ��Ŀ����_IN IN ҽ��������Ϣ.��Ŀ����%TYPE:=NULL,
    
        r = i
        StrInput = StrInput & "'" & Substr(str��ϸ, r, 1) & "',"
        r = r + 1
        StrInput = StrInput & "'" & Substr(str��ϸ, r, 20) & "',"
        r = r + 20
        StrInput = StrInput & "" & Val(Substr(str��ϸ, r, 10)) & ","
        r = r + 10
        StrInput = StrInput & "'" & Substr(str��ϸ, r, 10) & "',"
        r = r + 10
        StrInput = StrInput & "" & Val(Substr(str��ϸ, r, 10)) & ","
        
        
        '����
        'TRANSDETIAL���
         '        ���ܴ�����־    String(1)   Out
         '        ҽ���ڷ���  String(10)  Out
         '        ҽ�������  String(10)  Out
         '        ����ҽ��֧�� ����μӴ�ҽ������Ϊ��ҽ��֧��  String(10)  Out
         '        �߶�ҽ��֧��    String(10)  Out
         '        ����Աҽ�Ʋ���  String(10)  Out
         '        ���˰�����֧��  String(10)  Out

 '    ���ܴ�����־_IN IN ҽ��������Ϣ.���ܴ�����־%TYPE:=NULL,
    '    ����ҽ��֧��_IN IN ҽ��������Ϣ.����ҽ��֧��%TYPE:=NULL,
    '    �߶�ҽ��֧��_IN IN ҽ��������Ϣ.�߶�ҽ��֧��%TYPE:=NULL,
    '    ����Ա����_IN IN ҽ��������Ϣ.����Ա����%TYPE:=NULL,
    '    ���˱���֧��_IN IN ҽ��������Ϣ.���˱���֧��%TYPE:=NULL
 
        StrInput = StrInput & "'" & Substr(strArr(2), 1, 1) & "',"
        StrInput = StrInput & "" & Val(Substr(strArr(2), 22, 10)) & ","
        StrInput = StrInput & "" & Val(Substr(strArr(2), 32, 10)) & ","
        StrInput = StrInput & "" & Val(Substr(strArr(2), 42, 10)) & ","
         '20051021 add
        StrInput = StrInput & "" & Val(Substr(strArr(2), 52, 10)) & ","
        StrInput = StrInput & "" & Val(Substr(strArr(2), 62, 10)) & ")"
        '���SQL���
        
        gstrSQL = strHead & StrInput
        ExecuteProcedure_ZLNJ "������ϸ���ݵ��м��"
        StrInput = ""
    Next
    ������ϸ������м�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function StartOrCommitorRollbackTransaction(ByVal bytType As Byte, Optional blnGcnoracle As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������ύ���ع�����
    '--�����:byttype-0����,1�ύ,2�ع�
    '         blnGcnoracle-�Ƿ��������(gcnoracle)
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Select Case bytType
        Case 0
            gcnOracle_�ɶ��ڽ�.BeginTrans
            If Not blnGcnoracle Then
                gcnOracle.BeginTrans
            End If
            mblnStartTran = True
        Case 1
            gcnOracle_�ɶ��ڽ�.CommitTrans
            If Not blnGcnoracle Then
                gcnOracle.CommitTrans
            End If
            mblnStartTran = False
        Case Else
            gcnOracle_�ɶ��ڽ�.RollbackTrans
            If Not blnGcnoracle Then
                gcnOracle.RollbackTrans
            End If
            mblnStartTran = False
        End Select
End Function

Private Function �����ϴ�(ByVal lng��¼���� As Long, lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '������ϸ�ϴ�
    '����:�ϴ��²����ļ�����ϸ��ҽ������
    '����:  str���ݺ�   NO
    '       int����     ��¼����
    '       lng����ID  Ĭ��Ϊ0����ʾ�������ŵ��ݣ�����Ϊ������ָ�����˵ġ�����Ҫ����Ϊҽ���ڱ�����ʵ�ʱ���Ƿֲ������ύ���ݶ�����һ���ύ��
    '����:
    Dim rsTemp As New ADODB.Recordset, rs��ϸ As New ADODB.Recordset
    Dim StrInput As String, strOutput As String, strArr As Variant
    Dim lng����ID As Long, str��ϸ As String
    Dim i As Long
    Dim ȡ���׺�_int As Integer
    
    'Beging 20051025 add
    Dim bln����ϸ As Boolean
    Dim vat����֢ As Variant, str�������   As String
    Dim lngҽԺ���볤�� As Long, str��Ժ���� As String
    'End 20051025 add
    
    �����ϴ� = False
    
    Err = 0: On Error GoTo errHand:
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.��ҳID,to_char(A.����ʱ��,'yyyymmdd') as ����ʱ��,to_char(A.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս�� " & _
              "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,Z.λ�� as ��������,A.���ձ���,C.��Ŀ����,C.�������,J.��� as �շ����,C.�Ƿ�ҽ��,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ,E.���,G.���� ����,M.ҽ���� " & _
              "  From סԺ���ü�¼ A,���ű� Z,�շ���� J,�շ�ϸĿ B,�����ʻ� M,(Select O.*,Z.������� From ����֧����Ŀ O,������Ŀ Z where O.����=Z.���� and O.��Ŀ����=Z.���� and O.����=" & TYPE_�ɶ��ڽ� & ") C,������ҳ D,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
              "  where a.����id=M.����id and a.��������ID=Z.iD(+)   and M.����=" & TYPE_�ɶ��ڽ� & " and A.NO='" & str���ݺ� & "' and A.��¼����=" & lng��¼���� & " and A.��¼״̬=" & lng��¼״̬ & "And Nvl(A.�Ƿ��ϴ�,0)=0 " & _
              "        and A.�շ����=J.����(+)  and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=" & TYPE_�ɶ��ڽ� & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID(+) " & _
              "        AND B.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) " & _
              "  Order by A.����ID,A.�Ǽ�ʱ��,C.��Ŀ����"

    Call zlDatabase.OpenRecordset(rs��ϸ, gstrSQL, "������ϸ�ϴ�")
    Dim lng����ID As Long

    '�ȼ���Ƿ�����˵������������ڣ����з��Ӧ�ļ�¼��.
    With rs��ϸ
        '�ϴ���ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '���ۼ��
        
            If Val(!����) < 0 Or Val(!�۸�) < 0 Then
                ShowMsgbox "�ڵ����в������븺����!"
                Exit Function
            End If
            If Nvl(!��Ŀ����) = "" Then
                 MsgBox "����Ŀδ����ҽ������[" & Nvl(!����) & "-" & Nvl(!����) & "]�������ϴ���ϸ!", vbInformation, gstrSysName
                 Exit Function
            End If
            
            .MoveNext
        Loop
        
    End With
    
    If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
    
    Dim str������ˮ�� As String
    Dim blnStarTran As Boolean '��������
    Dim str��Ŀ���� As String, str���ձ��� As String, str������� As String
    
    StrInput = ""
    mblnStartTran = False
    ȡ���׺�_int = 0
    lng����ID = 0
    str��Ŀ���� = "@#$%^&(*)_+_)(*&^%$$#"
    '���з��ô���
    With rs��ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do Until .EOF
                If mblnStartTran = False Then
                    '��������
                    Call StartOrCommitorRollbackTransaction(0)
                End If
                
                str���ձ��� = Nvl(!���ձ���)
                '������ձ���Ϊ�գ�����Ҫ�û�ѡ�����
                If str���ձ��� = "" Then
                    str���ձ��� = GetItemInsure_�ɶ��ڽ�(0, !�շ�ϸĿID, False)
                End If
                '���Ϊ�ձ�ʾû��ȡ��ȱʡ���루ʹ���³��������ǰ����Ŀ�������������������ȡ��ǰ��¼���е���Ŀ���뼴��
                If str���ձ��� = "" Then str���ձ��� = Nvl(!��Ŀ����)
                
                'Begin 20051025
                '��ͨҽ������(��Ժ��ϲ��ǡ�ƽ�����͡��ʹ�����)
                'a.�����á��������á�(��������Ŀ����.xls��4��)�������á��ƻ��������á�(��������Ŀ��        ��.xlsǰ15��)����ͨ���û�ϴ���ϴ�;
                bln����ϸ = True
                gstrSQL = "Select * From ���ղ��� Where id=(select ����ID from �����ʻ� Where ����ID=" & !����ID & ")" & _
                          " And (���� ='ƽ��' or ����='�ʹ���')"
                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "ȡ��Ժ���")
                If rsTemp.EOF = True Then
                    If str���ձ��� = "M10000116" Or str���ձ��� = "M10000117" Or str���ձ��� = "M10000118" Or str���ձ��� = "M10000119" Then
                        ShowMsgbox "��ͨҽ����������ʹ���������ã�" & rs��ϸ!���� & vbCrLf & _
                                   "�շ���Ŀ��" & rs��ϸ!���� & " [" & str���ձ��� & "] ���ϴ�"
                        bln����ϸ = False
                        'Call StartOrCommitorRollbackTransaction(2)
                        'Exit Function
                    End If
                Else
                    gstrSQL = "Select to_char(��Ժ����,'yyyymmdd') as ��Ժ���� from ������ҳ Where ����ID=" & !����ID & " And ��ҳID=" & !��ҳID
                    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�������")
                    str��Ժ���� = rsTemp!��Ժ����
                    gstrSQL = "Select * from ������� where ���=" & TYPE_�ɶ��ڽ�
                    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "ҽԺ���")
                    lngҽԺ���볤�� = Len(Trim(rsTemp!ҽԺ����))
                    
                    gstrSQL = "Select * from �����ʻ� Where ����ID=" & !����ID & " And ����=" & TYPE_�ɶ��ڽ�
                    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�������")
                    If Nvl(rsTemp!�������) = "" Then
                        ' Beging һ����������(��Ժ����ǡ�ƽ�����͡��ʹ�����)
                        '       a.ֻ���ϴ����������á�(��������Ŀ����.xls��4��)�͡��ƻ��������á�(��������Ŀ��        ��.xlsǰ15��),����ʹ��һ����õ����ϴ���
                        '       b.4��������á���(���������̥-M10000118)ÿ����ظ�ʹ�ã���"ƽ��"�Ѻ�"�ʹ���"��ֻ��һ����
                        'If str���ձ��� = "M10000116" Or str���ձ��� = "M10000117" Or str���ձ��� = "M10000118" Or str���ձ��� = "M10000119" Then
                        '������ 20051026 �����üƻ�������Ŀ
                        If Substr(str���ձ���, 1, 1) = "M" Then
                            gstrSQL = "Select ���ձ��� from סԺ���ü�¼ Where ��¼״̬=1 and ժҪ<>'���ϴ�'" & _
                                      " and ���ձ��� IN ('M10000116','M10000117','M10000119') " & _
                                      " And ��ҳid=" & !��ҳID & " and ����id=" & !����ID
                            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�Ƿ����ظ���Ŀ")
                            Do Until rsTemp.EOF
                                If str���ձ��� = rsTemp!���ձ��� Then
                                    ShowMsgbox "ҽ���涨����Ŀ��" & rs��ϸ!���� & " [" & str���ձ��� & "] " & "�����ظ�ʹ��" & vbCrLf & _
                                               "�շ���Ŀ��" & rs��ϸ!���� & " [" & str���ձ��� & "] ���ϴ�"
                                    bln����ϸ = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                                rsTemp.MoveNext
                            Loop
                            
                            '������ 20051026  �Ż����
                            If str���ձ��� = "M10000116" Or str���ձ��� = "M10000117" Then
                                gstrSQL = "Select ���ձ��� from סԺ���ü�¼ Where ��¼״̬=1 and ժҪ<>'���ϴ�'" & _
                                          " And (���ձ���='M10000116' or ���ձ���='M10000117')" & _
                                          " And ��ҳid=" & !��ҳID & " and ����id=" & !����ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�Ƿ��л�����Ŀ")
                                If rsTemp.RecordCount > 0 Then
                                    ShowMsgbox "ҽ���涨M10000116(ƽ��)���ܺ�M10000117(�ʹ���)һͬʹ��" & vbCrLf & _
                                               "�շ���Ŀ��" & rs��ϸ!���� & " [" & str���ձ��� & "] ���ϴ�"
                                    bln����ϸ = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                            End If
                            
                            If str���ձ��� = "M10000118" Then
                                gstrSQL = "Select ���ձ��� from סԺ���ü�¼ Where ��¼״̬=1 and ժҪ<>'���ϴ�'" & _
                                          " And (���ձ���='M10000116' or ���ձ���='M10000117')" & _
                                          " And ��ҳid=" & !��ҳID & " and ����id=" & !����ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�Ƿ��л�����Ŀ")
                                If rsTemp.EOF Then
                                    ShowMsgbox "ҽ���涨ʹ��M10000118(�������̥),��Ҫ��ʹ��M10000116(ƽ��)��M10000117(�ʹ���)" & vbCrLf & _
                                               "�շ���Ŀ��" & rs��ϸ!���� & " [" & str���ձ��� & "] ���ϴ�"
                                    bln����ϸ = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                            End If
                            
                            If str���ձ��� = "M10000119" Then
                                gstrSQL = "Select ���ձ��� from סԺ���ü�¼ Where ��¼״̬=1 and ժҪ<>'���ϴ�'" & _
                                          " And ���ձ���='M10000117'" & _
                                          " And ��ҳid=" & !��ҳID & " and ����id=" & !����ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�Ƿ����ʹ���")
                                If rsTemp.EOF Then
                                    ShowMsgbox "ҽ���涨ʹ��M10000119(����ȫ������),��Ҫ��ʹ��M10000117(�ʹ���)" & vbCrLf & _
                                               "�շ���Ŀ��" & rs��ϸ!���� & " [" & str���ձ��� & "] ���ϴ�"
                                    bln����ϸ = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                            End If
                            
                            'bln����ϸ = True
                        Else
                            bln����ϸ = False
                        End If
                        'End һ����������
                    Else
                        'Beging ����֢��������(��Ժ����ǡ�ƽ�����͡��ʹ�������������˲���֢����)
'                            a.����ʹ�����з��ã�ʹ�õġ��������á�(��������Ŀ����.xls��4��)�͡��ƻ��������á�(��������Ŀ����.xlsǰ15��)��Ҫ�ϴ���
'
'                            c.4��������á���ÿ����ظ�ʹ�ã���"ƽ��"�Ѻ�"�ʹ���"��ֻ��һ��;
                        'If str���ձ��� = "M10000116" Or str���ձ��� = "M10000117" Or str���ձ��� = "M10000118" Or str���ձ��� = "M10000119" Then
                        '������ 20051026 �����üƻ�������Ŀ
                        If Substr(str���ձ���, 1, 1) = "M" Then
                            gstrSQL = "Select ���ձ��� from סԺ���ü�¼ Where ��¼״̬=1 and ժҪ<>'���ϴ�'" & _
                                      " And ���ձ��� IN ('M10000116','M10000117','M10000119')" & _
                                      " And ��ҳid=" & !��ҳID & " and ����id=" & !����ID
                            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�Ƿ����ظ���Ŀ")
                            Do Until rsTemp.EOF
                                If str���ձ��� = rsTemp!���ձ��� Then
                                    ShowMsgbox "ҽ���涨����Ŀ��" & rs��ϸ!���� & "[" & str���ձ��� & "]" & "�����ظ�ʹ��" & vbCrLf & _
                                               "�շ���Ŀ��" & rs��ϸ!���� & " [" & str���ձ��� & "] ���ϴ�"
                                    bln����ϸ = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                                rsTemp.MoveNext
                            Loop
                            
                            '������ 20051026  �Ż����
                            If str���ձ��� = "M10000116" Or str���ձ��� = "M10000117" Then
                                gstrSQL = "Select ���ձ��� from סԺ���ü�¼ Where ��¼״̬=1 and ժҪ<>'���ϴ�'" & _
                                          " And (���ձ���='M10000116' or ���ձ���='M10000117')" & _
                                          " And ��ҳid=" & !��ҳID & " and ����id=" & !����ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�Ƿ��л�����Ŀ")
                                If rsTemp.RecordCount > 0 Then
                                    ShowMsgbox "ҽ���涨M10000116(ƽ��)���ܺ�M10000117(�ʹ���)һͬʹ��" & vbCrLf & _
                                               "�շ���Ŀ��" & rs��ϸ!���� & " [" & str���ձ��� & "] ���ϴ�"
                                    bln����ϸ = False
                                   ' Call StartOrCommitorRollbackTransaction(2)
                                   ' Exit Function
                                End If
                            End If
                            
                            If str���ձ��� = "M10000118" Then
                                gstrSQL = "Select ���ձ��� from סԺ���ü�¼ Where ��¼״̬=1 and ժҪ<>'���ϴ�'" & _
                                          " And (���ձ���='M10000116' or ���ձ���='M10000117')" & _
                                          " And ��ҳid=" & !��ҳID & " and ����id=" & !����ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�Ƿ��л�����Ŀ")
                                If rsTemp.EOF Then
                                    ShowMsgbox "ҽ���涨ʹ��M10000118(�������̥),��Ҫ��ʹ��M10000116(ƽ��)��M10000117(�ʹ���)" & vbCrLf & _
                                               "�շ���Ŀ��" & rs��ϸ!���� & " [" & str���ձ��� & "] ���ϴ�"
                                    bln����ϸ = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                            End If
                            
                            If str���ձ��� = "M10000119" Then
                                gstrSQL = "Select ���ձ��� from סԺ���ü�¼ Where ��¼״̬=1 and ժҪ<>'���ϴ�'" & _
                                          " And ���ձ���='M10000117'" & _
                                          " And ��ҳid=" & !��ҳID & " and ����id=" & !����ID
                                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�Ƿ����ʹ���")
                                If rsTemp.EOF Then
                                    ShowMsgbox "ҽ���涨ʹ��M10000119(����ȫ������),��Ҫ��ʹ��M10000117(�ʹ���)" & vbCrLf & _
                                               "�շ���Ŀ��" & rs��ϸ!���� & " [" & str���ձ��� & "] ���ϴ�"
                                     bln����ϸ = False
                                    'Call StartOrCommitorRollbackTransaction(2)
                                    'Exit Function
                                End If
                            End If
                            
                            'bln����ϸ = True
                        Else
                            'b.ʹ�õ���ͨ��Ŀ������򲢷�֢����Ҫ�ķ�����Ŀ���ϴ�,�����ϴ�;
                          '������ 20051026 �ֹ�ѡ���Ƿ��ϴ�
                            'str������� = Nvl(rsTemp!�������)
                            'If InStr(str�������, "|") > 0 Then
                            '    vat����֢ = Split(str�������, "|")
                            '    For i = 0 To UBound(vat����֢) - 1
                            '        gstrSQL = "Select * From ������׼��Ŀ A,���ղ��� B " & _
                            '                "Where A.����ID=B.Id And B.����='" & Split(vat����֢(i), ";")(0) & "'" & _
                            '                " And A.�շ�ϸĿID=" & !�շ�ϸĿID
                            '        Call OpenRecordset(rsTemp, "�Ƿ�����׼��Ŀ")
                            '        If rsTemp.EOF Then
                            '            bln����ϸ = False
                            '        Else
                            '            bln����ϸ = True
                            '        End If
                            '    Next
                            'End If
                            If MsgBox("�շ���Ŀ��" & rs��ϸ!���� & "    ��" & rs��ϸ!ʵ�ս�� & vbCrLf & _
                                     "��ȷ���Ƿ񲢷�֢ʹ����Ŀ��" & vbCrLf & _
                                     "���� [��] ���ϴ���ҽ������ [��] ���ϴ���", vbYesNo, gstrSysName) = vbYes Then
                                bln����ϸ = True
                            Else
                                bln����ϸ = False
                            End If
                        End If
                        'End ����֢��������
                    End If

                End If
                
                'End 20051025
                
                'ȡ�������
                gstrSQL = "Select ������� From ������Ŀ Where ����=" & TYPE_�ɶ��ڽ� & " And ����='" & str���ձ��� & "'"
                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "ȡ�������")
                str������� = Nvl(rsTemp!�������)
                
               '������ 20051027
            If bln����ϸ Then
               If lng����ID <> Nvl(!����ID, 0) Or i >= 38 Or str��Ŀ���� = str���ձ��� Then
                    'i = 1
                    If ȡ���׺�_int = 0 Then
                        str������ˮ�� = Get������ˮ��(Nvl(!����ʱ��))
                        ȡ���׺�_int = 1
                    End If
                    lng����ID = Nvl(!����ID, 0)
                    str��Ŀ���� = str���ձ���
                    If StrInput <> "" Then
                        StrInput = StrInput & vbTab & Rpad(i - 1, 2) & vbTab & str��ϸ
                        '������ص�ҵ������
                        'Beging 20051025 add
                        'If bln����ϸ = True Then
                            If ҵ������_�ɶ��ڽ�(סԺ�����ϴ�_�ڽ�, StrInput, strOutput) = False Then
                                '�ع�����
                                Call StartOrCommitorRollbackTransaction(2)
                                Exit Function
                            End If
                            If ������ϸ������м��(str������ˮ��, strOutput) = False Then
                                '�ύ����,�м������δ������
                                Call StartOrCommitorRollbackTransaction(1)
                                Exit Function
                            End If
                        'End If
                        'End 20051025 add
                        ȡ���׺�_int = 0
                    End If
                    i = 1
                    If ȡ���׺�_int = 0 Then
                        str������ˮ�� = Get������ˮ��(Nvl(!����ʱ��))
                        ȡ���׺�_int = 1
                    End If
                    
                    Call Get������Ϣ(lng����ID)
                    StrInput = Rpad(g�������_�ɶ��ڽ�.���˱��, 8)
                    
                    StrInput = StrInput & vbTab & Rpad(g�������_�ɶ��ڽ�.����, 10)
                    StrInput = StrInput & vbTab & Rpad(InitInfor_�ɶ��ڽ�.ҽԺ����, 5)
                    StrInput = StrInput & vbTab & Rpad(g�������_�ɶ��ڽ�.ͳ����, 1)
                    StrInput = StrInput & vbTab & Rpad(str������ˮ��, 20)
                    If Nvl(str�������) = "1" Then
                        StrInput = StrInput & vbTab & IIf(IS��Ժ��ҩ(Nvl(!NO), Nvl(!ID, 0)), "1", "0")
                    Else
                        StrInput = StrInput & vbTab & "0"
                    End If
                    StrInput = StrInput & vbTab & Rpad(Substr(Nvl(!��������), 1, 10), 10)
                    StrInput = StrInput & vbTab & Rpad(Substr(!ҽ��, 1, 10), 10)
                    StrInput = StrInput & vbTab & Rpad(Substr(g�������_�ɶ��ڽ�.סԺ��ˮ��, 1, 20), 20)
                    str��ϸ = ""
                    '���˱��    String(8)   In
                    '�籣������  String(10)  In
                    'ҽԺ����    String(5)   In
                    'ͳ���������    String(1)   In
                    'ҽԺ������ˮ��  String(20)  In
                    '��Ժ��ҩ���    String(1)   In
                    
                    '�Ʊ�    String(10)  In
                    'ҽ��    String(10)  In
                    'סԺ��ˮ��  String(20)  In
                    '��������    String(2)   In
                    '������ϸ    String����������51  In
               End If
                
                lng����ID = Nvl(!����ID, 0)
                str��Ŀ���� = str���ձ���
                
                str��ϸ = str��ϸ & Substr(Rpad(Nvl(str�������), 1), 1, 1)
                str��ϸ = str��ϸ & Rpad(str���ձ���, 20)
                str��ϸ = str��ϸ & Lpad(Nvl(!����) * 100, 10, "0")
                str��ϸ = str��ϸ & Rpad(Nvl(!���), 10)
                str��ϸ = str��ϸ & Lpad(Nvl(!ʵ�ս��) * 100, 10, "0")
            
                i = i + 1
            End If
            
            '������ 20051027
            If bln����ϸ Then
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL," & Nvl(str�������, "NULL") & ",NULL,'" & str���ձ��� & "',1,'" & str������ˮ�� & "')"
            Else
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL," & Nvl(str�������, "NULL") & ",NULL,'" & str���ձ��� & "',1,'���ϴ�')"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
            .MoveNext
        Loop
    End With
    
    'Beging 20051025 add
    'If bln����ϸ = True Then
        If StrInput <> "" Then
            StrInput = StrInput & vbTab & Rpad(i - 1, 2) & vbTab & str��ϸ
            '������ص�ҵ������
            If ҵ������_�ɶ��ڽ�(סԺ�����ϴ�_�ڽ�, StrInput, strOutput) = False Then
                '�ύ����,�м������δ������
                Call StartOrCommitorRollbackTransaction(2)
                Exit Function
            End If
            If ������ϸ������м��(str������ˮ��, strOutput) = False Then
                '�ύ����,�м������δ������
                Call StartOrCommitorRollbackTransaction(2)
                Exit Function
            End If
            '�ύ
            StartOrCommitorRollbackTransaction (1)
        Else
            '������ 20051027
            If bln����ϸ Then
              If mblnStartTran Then
                 Call StartOrCommitorRollbackTransaction(2)
              End If
            Else
              StartOrCommitorRollbackTransaction (1)
            End If
        End If
    'End If
    'End 20051025 add

    �����ϴ� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    If mblnStartTran Then
        '�ύ����,�м������δ������
        Call StartOrCommitorRollbackTransaction(2)
    End If
End Function
Private Function IS��Ժ��ҩ(ByVal strNO As String, lng����ID As Long) As Boolean
    '����Ƿ��Ժ��ҩ
    'Dim rsTemp As New ADODB.Recordset
    Dim rsTemp1 As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errHand:
    Set rsTemp1 = New ADODB.Recordset
    gstrSQL = "Select ID From ҩƷ�շ���¼ where NO='" & strNO & "' and ���� IN(9,10) and ����id=" & lng����ID & " and ���� like '_3%'"
    zlDatabase.OpenRecordset rsTemp1, gstrSQL, "��ȡ�Ƿ��Ժ��ҩ"
    If rsTemp1.EOF Then
        IS��Ժ��ҩ = False
        Exit Function
    End If
    IS��Ժ��ҩ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function �����Ǽ�_�ɶ��ڽ�(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
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
    
    Err = 0
    On Error GoTo errHand:


    �����Ǽ�_�ɶ��ڽ� = False
    
    If lng��¼״̬ = 1 Then
        '��������
        If �����ϴ�(lng��¼����, lng��¼״̬, str���ݺ�) = False Then
            Exit Function
        End If
    Else
        '��ʼ����
        Call StartOrCommitorRollbackTransaction(0)
        '��������
        If ��������(lng��¼����, lng��¼״̬, str���ݺ�) = False Then
            '�ύ����,�м������δ������,���Իع�
            Call StartOrCommitorRollbackTransaction(2)
            Exit Function
        End If
        '�ύ����
        Call StartOrCommitorRollbackTransaction(1)
    End If
    �����Ǽ�_�ɶ��ڽ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ��������(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ʴ�������,����¼״̬=2�ļ�¼
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsԭ��ϸ As New ADODB.Recordset
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String, str������ˮ�� As String
    Dim strArr
    Dim lng����ID As Long
    
    �������� = False

    Err = 0: On Error GoTo errHand:

    
    gstrSQL = " Select a.ժҪ,A.ID,a.�շ�ϸĿid,A.���,A.����*nvl(A.����,1) as ����,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4) as ����,a.����id " & _
              " From סԺ���ü�¼ A,�����ʻ� B " & _
              " where a.����id=b.����id and A.NO='" & str���ݺ� & "' and A.��¼����=" & lng��¼���� & " and A.��¼״̬=3 and   Nvl(���ӱ�־,0)<>9  " & _
              " order by A.����id,A.ժҪ"
              
    Call zlDatabase.OpenRecordset(rsԭ��ϸ, gstrSQL, "������ϸ�ϴ�")
    
    
    If rsԭ��ϸ.EOF Then
        ShowMsgbox "�õ���û����Ӧ����ϸ��¼,��������!"
        Exit Function
    End If

    gstrSQL = " Select a.* " & _
              " From סԺ���ü�¼ A,�����ʻ� b" & _
              " where a.����id=b.����id and A.NO='" & str���ݺ� & "' and A.��¼����=" & lng��¼���� & " and A.��¼״̬=2 and  Nvl(���ӱ�־,0)<>9 AND nvl(a.�Ƿ��ϴ�,0)=0 " & _
              " order by A.����ID"
              
    Call zlDatabase.OpenRecordset(rs��ϸ, gstrSQL, "������ϸ�ϴ�")

    lng����ID = 0
    '����ԭ���ݵ�ֵ
    With rs��ϸ
        Do While Not .EOF
            rsԭ��ϸ.Filter = "���=" & Nvl(!���, 0) & "  and �շ�ϸĿid=" & Nvl(!�շ�ϸĿID, 0)
            If rsԭ��ϸ.EOF Then
                ShowMsgbox "����ʱδ�ҵ���Ӧ�ļ�¼,����ʧ��!"
                Exit Function
            End If
            str������ˮ�� = Nvl(rsԭ��ϸ!ժҪ)
            If str������ˮ�� = "" Then
                ShowMsgbox "��ԭ���в����ڽ�����ˮ��,���ܼ�����"
                Exit Function
            End If
            '���������ϸ���н�����ˮ��û��
            '������ 20051027
            If str������ˮ�� <> "���ϴ�" Then
            gstrSQL = "Select ҽ����ˮ�� From ҽ��������Ϣ where ҽԺ��ˮ��='" & str������ˮ�� & "' and ����id=" & Nvl(!����ID, 0)
            OpenRecordset_�ɶ��ڽ� rsTemp, "��ȡҽ��������Ϣ"
            If rsTemp.EOF Then
                ShowMsgbox "������ҽ����������,����ϵͳ����Ա��ϵ!"
                Exit Function
            End If
            If Nvl(rsTemp!ҽ����ˮ��) = "" Then
                ShowMsgbox "������ҽ����������,����ϵͳ����Ա��ϵ!"
                Exit Function
            End If
            End If
            '�����ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(rsԭ��ϸ!ժҪ) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
            .MoveNext
        Loop
    End With
    Dim strժҪ As String
    
    gstrSQL = " Select a.ժҪ,A.ID,a.�շ�ϸĿid,A.���,A.����*nvl(A.����,1) as ����,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4) as ����,a.����id " & _
              " From סԺ���ü�¼ A,�����ʻ� B " & _
              " where a.����id=b.����id and A.NO='" & str���ݺ� & "' and A.��¼����=" & lng��¼���� & " and A.��¼״̬=3 and   Nvl(���ӱ�־,0)<>9  " & _
              " order by A.����id,A.ժҪ"
              
    Call zlDatabase.OpenRecordset(rsԭ��ϸ, gstrSQL, "������ϸ�ϴ�")
    
    lng����ID = 0
    strժҪ = ""
    With rsԭ��ϸ
        .MoveFirst
        Do While Not .EOF
                'lng����ID <> Nvl(!����ID, 0) And
                '������ 20051027
            If Nvl(!ժҪ) <> "���ϴ�" Then
                If strժҪ <> Nvl(!ժҪ) Then
                    If lng����ID <> Nvl(!����ID, 0) Then
                        lng����ID = Nvl(!����ID, 0)
                        '�����»�ȡ��صĲ�����Ϣ
                        If Get������Ϣ(lng����ID) = False Then
                            ShowMsgbox "�ڻ�ȡ������Ϣʱ����˴���,����ϵͳԱ������ϵ!"
                            Exit Function
                        End If
                    End If
                    strժҪ = Nvl(!ժҪ)
                    gstrSQL = "Select ҽ����ˮ�� From ҽ��������Ϣ where ҽԺ��ˮ��='" & strժҪ & "' and ����id=" & lng����ID
                    OpenRecordset_�ɶ��ڽ� rsTemp, "��ȡҽ��������Ϣ"
                    
                    StrInput = Rpad(g�������_�ɶ��ڽ�.���˱��, 8)
                    StrInput = StrInput & vbTab & Rpad(g�������_�ɶ��ڽ�.����, 10)
                    StrInput = StrInput & vbTab & Rpad(Substr(gstrUserName, 1, 10), 10)
                    StrInput = StrInput & vbTab & Rpad(g�������_�ɶ��ڽ�.ͳ����, 1)
                    StrInput = StrInput & vbTab & Rpad(g�������_�ɶ��ڽ�.סԺ��ˮ��, 20)
                    StrInput = StrInput & vbTab & Rpad(Nvl(rsTemp!ҽ����ˮ��), 20)
                    
                    'ȡ��
                    '    ���˱��    String(8)   In
                    '    �籣������  String(10)  In
                    '    ����Ա������    String(10)  In
                    '    ͳ���������    String(1)   In
                    '    סԺ��ˮ��  String(20)  In
                    '    ҽ��������ˮ��  String(20)  In
                    
                    If ҵ������_�ɶ��ڽ�(סԺ�����ϴ�ȡ��_�ڽ�, StrInput, strOutput) = False Then Exit Function
                End If
            End If
            .MoveNext
        Loop
    End With
    �������� = True
    Exit Function
errHand:
   If ErrCenter = 1 Then
        Resume
   End If
End Function
Private Function Readģ������(ByVal intҵ������ As ҵ������_�ɶ��ڽ�, ByVal strInputString As String, ByRef strOutPutstring As String)
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
                    strArr = Split(strText, vbTab & "|")
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab & "|"
                            End If
                            strArr = Split(strText, vbTab & "|")
                            
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
    Exit Function
errHand:
    DebugTool Err.Description
    Exit Function
End Function

Private Function Get������Ϣ(ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strArr

    Get������Ϣ = False
    
    Err = 0
    On Error GoTo errHand:
    'COMMENT ON COLUMN �����ʻ�.����ID   is '����ID';
    'COMMENT ON COLUMN �����ʻ�.����     is '�̶�ֵ:106';
    'COMMENT ON COLUMN �����ʻ�.����     is '0';
    'COMMENT ON COLUMN �����ʻ�.����     is '����';
    'COMMENT ON COLUMN �����ʻ�.ҽ����   is 'ͳ����+���˱��';
    'COMMENT ON COLUMN �����ʻ�.����     is '��';
    'COMMENT ON COLUMN �����ʻ�.��Ա��� is '�������';
    'COMMENT ON COLUMN �����ʻ�.��λ���� is '��λ����';
    '
    'COMMENT ON COLUMN �����ʻ�.˳���   is 'ֻ���סԺ:סԺ��ˮ��';
    'COMMENT ON COLUMN �����ʻ�.����֤�� is 'ͳ���������|�ƿ�����|����Ч����|�ƿ���λ|��ְ���';
    'COMMENT ON COLUMN �����ʻ�.�ʻ���� is '�ʻ����';
    'COMMENT ON COLUMN �����ʻ�.��ǰ״̬ is '0-����,1-��Ժ';
    'COMMENT ON COLUMN �����ʻ�.����ID   is '��';
    'COMMENT ON COLUMN �����ʻ�.��ְ     is 'Ŀǰ�����ֵ��1�����ô�';
    'COMMENT ON COLUMN �����ʻ�.�����   is '��������';
    'COMMENT ON COLUMN �����ʻ�.�Ҷȼ�   is '�������';
    'COMMENT ON COLUMN �����ʻ�.����ʱ�� is '��ǰ�����ʱ��';
    '
    'COMMENT ON COLUMN �����ʻ�.���ܴ�����־ is 'ֻ���סԺ:���ܴ�����־';
    'COMMENT ON COLUMN �����ʻ�.�𸶱�׼ is 'ֻ���סԺ:�𸶱�׼';
    
'    gstrSQL = "select a.*,b.����,b.�Ա�, b.����, b.��������, b.���֤��,b.������λ,c.����,C.���� " & _
'             " from �����ʻ� a,������Ϣ b,���ղ��� C " & _
'             " WHERE  a.����ID=c.ID and a.����id=" & lng����ID & " AND a.����id=b.����id and a.����=" & TYPE_�ɶ��ڽ�
    '2006-4-10  �������޸�
    gstrSQL = "select a.����,A.ҽ����,A.�Ҷȼ�,A.��λ����,A.����֤��,A.�����,A.�ʻ����,A.��Ա���,A.˳���,A.����ID" & _
            ",b.����,b.�Ա�, b.����, b.��������, b.���֤��,b.������λ,c.����,C.���� " & _
            " from �����ʻ� a,������Ϣ b,���ղ��� C " & _
            " WHERE  a.��Ժ����ID=c.ID(+) and a.����id=" & lng����ID & " AND a.����id=b.����id and a.����=" & TYPE_�ɶ��ڽ�

    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ϣ"

    With g�������_�ɶ��ڽ�
        .���� = Nvl(rsTemp!����)
        .���˱�� = Mid(Nvl(rsTemp!ҽ����), 2)
        .���֤�� = Nvl(rsTemp!���֤��)
        .���� = Nvl(rsTemp!����)
        .�Ա� = Decode(Nvl(rsTemp!�Ա�), "��", 1, "Ů", 2, 1)
        .������� = Nvl(rsTemp!�Ҷȼ�)
        .�������� = Format(rsTemp!��������, "yyyy-mm-dd")
        .��λ���� = Nvl(rsTemp!��λ����)
        strArr = Split(Nvl(rsTemp!����֤��) & "|||||", "|")
        .ͳ���� = strArr(0)
        .�ƿ����� = strArr(1)
        .����Ч�� = strArr(2)
        .�������� = Nvl(rsTemp!�����)
        .�ƿ���λ = strArr(3)
        .���� = Nvl(rsTemp!����, 0)
        .�ʻ���� = Nvl(rsTemp!�ʻ����, 0)
        .��ְ��� = strArr(4)
        .������� = Nvl(rsTemp!��Ա���)
        .סԺ��ˮ�� = Nvl(rsTemp!˳���)
        .lng����ID = Nvl(rsTemp!����ID, 0)
        .���ֱ��� = Nvl(rsTemp!����)
        .�������� = Nvl(rsTemp!����)
        
    End With
    Get������Ϣ = True
Exit Function
errHand:
        DebugTool "��ȡ������Ϣʧ��" & vbCrLf & " �����:" & Err.Number & vbCrLf & " ������Ϣ:" & Err.Description
End Function

Private Sub OpenRecordset_�ɶ��ڽ�(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "", Optional cnOracle As ADODB.Connection)
    '���ܣ��򿪼�¼��
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    If cnOracle Is Nothing Then
        rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_�ɶ��ڽ�, adOpenStatic, adLockReadOnly
    Else
        If cnOracle.State <> 1 Then
            rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_�ɶ��ڽ�, adOpenStatic, adLockReadOnly
        Else
            rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), cnOracle, adOpenStatic, adLockReadOnly
        End If
    End If
    Call SQLTest
End Sub


Public Function סԺ�������_�ɶ��ڽ�(rsExse As Recordset, ByVal lng����ID As Long, Optional bln���ʴ� As Boolean = True) As String
    
    'rsExse:�ַ���
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������

    Dim cn�ϴ� As New ADODB.Connection, rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    Dim lng��ҳID As Long
    Dim StrInput As String, strOutput   As String
    Dim strArr As Variant
    Dim intMouse As Integer
    Dim dbl�����ܶ� As Double
    Dim dbl�������� As Double
    Dim str���㷽ʽ  As String
    
    Err = 0: On Error GoTo errHand:

    g�������_�ɶ��ڽ�.lng����ID = lng����ID
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If

    With g��������
        .����֧�� = 0
        .������־ = ""
        .�߶�ҽ��֧�� = 0
        .����Աҽ�Ʋ��� = 0
        .�𸶱�׼ = 0
        .ҽ��������ˮ�� = ""
        .ҽ���ڷ��� = 0
        .ҽ������� = 0
        .�ʻ�������� = 0
        .�ʻ�֧�� = 0
    End With
    
    If Get������Ϣ(lng����ID) = False Then Exit Function

    If bln���ʴ� Then
        Screen.MousePointer = 1
        If ��ݱ�ʶ_�ɶ��ڽ�(4, lng����ID) = "" Then
            Screen.MousePointer = intMouse
            סԺ�������_�ɶ��ڽ� = ""
            Exit Function
        End If
        If lng����ID <> g�������_�ɶ��ڽ�.lng����ID Then
            ShowMsgbox "��Ŀ���������,���ܽ��н���!"
            Exit Function
        End If
        Screen.MousePointer = intMouse
        
    Else
        Call Get������Ϣ(lng����ID)
    End If
    
    '�жϵ�λǷ�����
    '    ���˱��    String (8)  IN
    '    �籣������  String (10) IN
    '    ͳ���������    String (1)  IN
    StrInput = g�������_�ɶ��ڽ�.���˱��
    StrInput = StrInput & vbTab & g�������_�ɶ��ڽ�.����
    StrInput = StrInput & vbTab & g�������_�ɶ��ڽ�.ͳ����
    
    If ҵ������_�ɶ��ڽ�(��ȡ��λǷ�����_�ڽ�, StrInput, strOutput) = False Then
        Exit Function
    End If
    
    If Val(strOutput) <> 0 Then
        ShowMsgbox "ע�⣺" & vbCrLf & "    ��λ�Ѿ�Ƿ��"
    End If

    gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=" & rsExse("����ID")
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�������")
    If IsNull(rsTemp("��ҳID")) = True Then
        MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    lng��ҳID = rsTemp("��ҳID")
    Screen.MousePointer = vbHourglass
    
    '������ϸ
    If ����סԺ��ϸ��¼(lng����ID, lng��ҳID) = False Then Exit Function
        
    
    g�������_�ɶ��ڽ�.�����ܶ� = 0
    g�������_�ɶ��ڽ�.�����ܷ��� = 0
    With rsExse
        Do While Not .EOF
            'g�������_�ɶ��ڽ�.�����ܶ� = g�������_�ɶ��ڽ�.�����ܶ� + Nvl(!���, 0)
            
            'Beging 20051027 �¶�
            Dim str������� As String
            If Nvl(rsExse!���մ���id, 0) = 0 Then
                gstrSQL = "Select Nvl(�������,0) As ������� From ������Ŀ Where ����=" & TYPE_�ɶ��ڽ� & _
                          " And ����=(Select nvl(���ձ���,'0') From סԺ���ü�¼ Where �շ�ϸĿID=" & rsExse!�շ�ϸĿID & _
                          " And NO='" & rsExse!NO & "' And ��¼����=" & rsExse!��¼���� & _
                          " And ��¼״̬=" & rsExse!��¼״̬ & " And  ���=" & rsExse!��� & ")"
                Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "ȡ�������")
                
                str������� = "0"
                If rsTemp.RecordCount > 0 Then
                    str������� = Nvl(rsTemp!�������, "0")
                End If
                
                If Substr(Rpad(str�������, 1), 1, 1) <> "4" And Nvl(rsExse!ժҪ, "") <> "���ϴ�" Then
                    g�������_�ɶ��ڽ�.�����ܶ� = g�������_�ɶ��ڽ�.�����ܶ� + Nvl(rsExse!���, 0)
                End If
                If Substr(Rpad(str�������, 1), 1, 1) = "4" Or Nvl(rsExse!ժҪ, "") = "���ϴ�" Then
                    g�������_�ɶ��ڽ�.�����ܷ��� = g�������_�ɶ��ڽ�.�����ܷ��� + Nvl(rsExse!���, 0)
                End If
            Else
                If Nvl(rsExse!���մ���id, 0) <> "4" And Nvl(rsExse!ժҪ, "") <> "���ϴ�" Then
                    g�������_�ɶ��ڽ�.�����ܶ� = g�������_�ɶ��ڽ�.�����ܶ� + Nvl(rsExse!���, 0)
                End If
                If Nvl(rsExse!���մ���id, 0) = "4" Or Nvl(rsExse!ժҪ, "") = "���ϴ�" Then
                    g�������_�ɶ��ڽ�.�����ܷ��� = g�������_�ɶ��ڽ�.�����ܷ��� + Nvl(rsExse!���, 0)
                End If
            End If
            'End 20051027 �¶�
            
            
            .MoveNext
        Loop
    End With
        
        
        
    gstrSQL = "Select A.ID From סԺ���ü�¼ a,ҩƷ�շ���¼ B where A.no=b.No and B.���� in (9,10) and a.id=b.����ID and a.����id=" & lng����ID & " and ��ҳid=" & lng��ҳID & " and b.���� like '_3%' and rownum<=2"

    
    Dim bln��Ժ��ҩ As Boolean
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ����Ժ��ҩ"
    If rsTemp.EOF Then
        bln��Ժ��ҩ = False
    Else
        bln��Ժ��ҩ = True
    End If
    
    gstrSQL = "Select c.סԺ��,A.�Ǽ��� ������,B.���� ��Ժ����,A.סԺҽʦ,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') ��Ժ����ʱ��," & _
        " to_char(A.��Ժ����,'yyyyMMdd') ��Ժ����,D.��ϱ���,A.��Ժ��ʽ,to_Char(a.��Ժ����,'yyyyMMDD') as ��Ժ����,a.��Ժ����,H.λ�� as ��Ժ����" & _
        " From ������ҳ A,���ű� B,������Ϣ C,���ű� H, " & _
        "       (Select ����id,��ҳid,max(DECODE(a.��ϴ���,2,b.����,'')) AS ��ϱ��� From ������ A ,��������Ŀ¼ B Where a.����ID = b.ID And a.������� =3  and a.��ҳid=" & lng��ҳID & " and a.����id=" & lng����ID & " Group by ����id,��ҳid)   D" & _
        " Where A.����id=C.����id and C.����id=" & lng����ID & _
        "       and A.����ID=" & lng����ID & " And A.��ҳID=" & lng��ҳID & " And A.��Ժ����ID=B.ID and A.��Ժ����ID=H.id(+) " & _
        "       and A.��ҳid=D.��ҳid(+) and a.����id=D.����id(+) " & _
        ""
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȷ��������"
    
    
     
     '���:
     '    ���˱��    String(8)   In
     '    �籣������  String(10)  In
     '    ҽԺ����    String(5)   In
     '    ����Ա������    String(10)  In
     '    ͳ���������    String(1)   In
     '    ��Ժ����    String(8)   In
     '    ��Ժ�Ʊ�    String(10)  In
     '    ��Ժ����ҽ��    String(10)  In
     '    ��ϱ���    String(20)  In
     '    ��Ժ��ҩ    String(1)   In
     '    ��Ժ���    String(1)   In
     '    סԺ��ˮ��  String(20)  In
       
     With g�������_�ɶ��ڽ�
         StrInput = Rpad(.���˱��, 8)
         StrInput = StrInput & vbTab & Rpad(.����, 10)
         StrInput = StrInput & vbTab & Rpad(InitInfor_�ɶ��ڽ�.ҽԺ����, 5)
         StrInput = StrInput & vbTab & Rpad(Nvl(gstrUserName), 10)
         StrInput = StrInput & vbTab & Rpad(.ͳ����, 1)
         
         If Trim(Nvl(rsTemp!��Ժ����)) = "" Then
            '���û�г�Ժ����,�򴫵�ǰ��ʱ����ȥ.
            StrInput = StrInput & vbTab & Format(zlDatabase.Currentdate, "yyyymmdd")
         Else
            StrInput = StrInput & vbTab & Rpad(Nvl(rsTemp!��Ժ����), 8)
         End If
         
         StrInput = StrInput & vbTab & Rpad(Substr(Rpad(Nvl(rsTemp!��Ժ����), 10), 1, 10), 10)
         StrInput = StrInput & vbTab & Rpad(Substr(Rpad(Nvl(rsTemp!סԺҽʦ), 10), 1, 10), 10)
        'Beging 20051026 �¶�
        Dim vat����֢ As Variant, str��Ժ�������� As String, i As Long
        Dim rsCydj As New ADODB.Recordset
        
        gstrSQL = "Select * from �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_�ɶ��ڽ�
        Call zlDatabase.OpenRecordset(rsCydj, gstrSQL, "ȡ��Ժ��������")
        str��Ժ�������� = Nvl(rsCydj!��Ժ��������)
        If str��Ժ�������� <> "" Then
            If InStr(str��Ժ��������, "|") > 0 Then
                vat����֢ = Split(str��Ժ��������, "|")
                str��Ժ�������� = ""
                For i = 0 To UBound(vat����֢) - 1
                    str��Ժ�������� = str��Ժ�������� & Rpad(Substr(vat����֢(i), 1, 20), 20)
                Next
            End If
        Else
            str��Ժ�������� = Space(180)
        End If
        'End 20051026 �¶�
         StrInput = StrInput & vbTab & Rpad(Rpad(Substr(g�������_�ɶ��ڽ�.���ֱ���, 1, 20), 20) & Substr(str��Ժ��������, 1, 180), 200)
         StrInput = StrInput & vbTab & IIf(bln��Ժ��ҩ, "1", "0")
         StrInput = StrInput & vbTab & Rpad(g�������_�ɶ��ڽ�.��Ժ���, 1)
         'strInput = strInput & vbTab & Rpad(Get�������_�ڽ�(lng����ID, lng��ҳID), 1)
         StrInput = StrInput & vbTab & Rpad(Substr(Rpad(.סԺ��ˮ��, 20), 1, 20), 20)
         'Beging �¶� 20051020
         If bln���ʴ� = True Then
            If ҵ������_�ɶ��ڽ�(��Ժ�Ǽ��ϴ�_�ڽ�, StrInput, strOutput) = False Then Exit Function
         Else
            Exit Function
         End If
         'End �¶� 20051020
     End With
     
     If strOutput = "" Then Exit Function
     
     strArr = Split(strOutput, vbTab)
     
    '����
     '    TRANSDETIAL��� (���������ϸ)
     '    ���ܴ�����־    String(1)   Out
     '    ҽ���ڷ���  String(10)  Out
     '    ҽ�������  String(10)  Out
     '    ����ҽ��֧��
     '    ����μӴ�ҽ������Ϊ��ҽ��֧��  String(10)  Out
     '    �߶�ҽ��֧��    String(10)  Out
     '    ����Աҽ�Ʋ���  String(10)  Out
     '    ���˰�����֧��  String(10)  Out
     '    TRANSDETIAL����
     '    �𸶱�׼    String(10)  Out
     '    �����ʻ��������    String(10)  Out
     strOutput = strArr(0)
     With g��������
         .�����־ = 1
         .������־ = Substr(strOutput, 1, 1)
         .ҽ���ڷ��� = Val(Substr(strOutput, 2, 10)) / 100
         .ҽ������� = Val(Substr(strOutput, 12, 10)) / 100
         .����ҽ��֧�� = Val(Substr(strOutput, 22, 10)) / 100
         .�߶�ҽ��֧�� = Val(Substr(strOutput, 32, 10)) / 100
         .����Աҽ�Ʋ��� = Val(Substr(strOutput, 42, 10)) / 100
         .����֧�� = Val(Substr(strOutput, 52, 10)) / 100
         .�𸶱�׼ = Val(strArr(1)) / 100
         .�ʻ�������� = Val(strArr(2)) / 100
         '20051021 add
         .����֧�� = Val(Substr(strOutput, 62, 10)) / 100
     End With
     
     
    dbl�����ܶ� = g��������.ҽ���ڷ��� + g��������.ҽ�������
    '������ 20060809 �ڽ���ԺҪ��ҽ������ÿ������ʻ�֧��
    'g��������.�ʻ�֧�� = dbl�����ܶ� - g��������.ҽ������� - g��������.����ҽ��֧�� - g��������.�߶�ҽ��֧�� - g��������.����Աҽ�Ʋ���
    g��������.�ʻ�֧�� = dbl�����ܶ� - g��������.����ҽ��֧�� - g��������.�߶�ҽ��֧�� - g��������.����Աҽ�Ʋ���
    'by 20050122 gzy
    If g��������.�ʻ�������� >= 0 Then
        If g��������.�ʻ�������� < g��������.�ʻ�֧�� Then
           g��������.�ʻ�֧�� = g��������.�ʻ��������
        End If
    Else
        g��������.�ʻ�֧�� = 0
    End If
    If g��������.�ʻ�֧�� <= 0 Then g��������.�ʻ�֧�� = 0
    
    str���㷽ʽ = "�����ʻ�;" & g��������.�ʻ�֧�� & ";1"
    
    If g��������.����ҽ��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|����ͳ��;" & g��������.����ҽ��֧�� & ";0"
    End If
    If g��������.�߶�ҽ��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|��֧��;" & g��������.�߶�ҽ��֧�� & ";0"
    End If
    If g��������.����Աҽ�Ʋ��� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & g��������.����Աҽ�Ʋ��� & ";0"
    End If
    
    g��������.����ӯ�� = 0
    If g��������.����֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|����֧��;" & g��������.����֧�� & ";0"
        'Modified by ZYB 20051118
        '��ҽ�����İ���֧��������ʵ�ʷ������������ý��бȽϣ����ಿ�ּ�������ӯ����
        g��������.����ӯ�� = g�������_�ɶ��ڽ�.�����ܷ��� - g��������.����֧��
    End If
    
    '������ 20051029 �����������ð���ģʽ
    'Modified by ZYB 20051118
    If g��������.����ӯ�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|����ӯ��;" & g��������.����ӯ�� & ";0"
        ShowMsgbox "��Ժ��������������Ϊ��" & g�������_�ɶ��ڽ�.�����ܷ��� & vbCrLf & _
                  "ҽ�����ı����������ã�" & g��������.����֧�� & vbCrLf & _
                  "�������õ�ǰ���㷽ʽ�� �����ٲ�  " & vbCrLf & _
                  "����ӯ���� " & g��������.����ӯ��
    End If
    
    'If Format(g�������_�ɶ��ڽ�.�����ܶ�, "###0.00;-###0.00;0;0") <> Format(dbl�����ܶ�, "###0.00;-###0.00;0;0") Then
    '    Dim blnYes As Boolean
    '    �����ܶ���ҽ�����ķ����ܶ��,���ܽ��н���
    '    ShowMsgbox "���ν����ܶ�(" & g�������_�ɶ��ڽ�.�����ܶ� & ") ��" & vbCrLf & _
    '                "   ���ķ��ص��ܶ�(" & dbl�����ܶ� & ")����,���ܽ���?"
    '    Exit Function
    'End If
    
    'gzy 20051129 ��ͨ�������˲��ж�ҽ�����
    gstrSQL = "Select * From ���ղ��� Where id=(select ����ID from �����ʻ� Where ����ID=" & lng����ID & ")" & _
              " And (���� ='ƽ��' or ����='�ʹ���')"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "ȡ��Ժ���")
    If rsTemp.EOF = True Then
       If Format(g�������_�ɶ��ڽ�.�����ܶ�, "###0.00;-###0.00;0;0") <> Format(dbl�����ܶ�, "###0.00;-###0.00;0;0") Then
           Dim blnYes As Boolean
           '�����ܶ���ҽ�����ķ����ܶ��,���ܽ��н���
           ShowMsgbox "���ν�������������ܶ�(" & g�������_�ɶ��ڽ�.�����ܶ� & ") ��" & vbCrLf & _
                    "   ���ķ��صķ����������ܶ�(" & dbl�����ܶ� & ")����,���ܽ���?"
           Exit Function
       End If
    Else
       gstrSQL = "Select * from �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_�ɶ��ڽ�
       Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�������")
       If Nvl(rsTemp!�������) <> "" Then
          If Format(g�������_�ɶ��ڽ�.�����ܶ�, "###0.00;-###0.00;0;0") <> Format(dbl�����ܶ�, "###0.00;-###0.00;0;0") Then
              'Dim blnYes As Boolean
              '�����ܶ���ҽ�����ķ����ܶ��,���ܽ��н���
              ShowMsgbox "���ν�������������ܶ�(" & g�������_�ɶ��ڽ�.�����ܶ� & ") ��" & vbCrLf & _
                    "   ���ķ��صķ����������ܶ�(" & dbl�����ܶ� & ")����,���ܽ���?"
              Exit Function
          End If
       Else
          If g�������_�ɶ��ڽ�.�����ܶ� <> 0 Then
             ShowMsgbox "��������������(" & g�������_�ɶ��ڽ�.�����ܶ� & ")�������½��н��㡣"
              Exit Function
          End If
       End If
    End If
    
    סԺ�������_�ɶ��ڽ� = str���㷽ʽ
    g�������_�ɶ��ڽ�.lng����ID = lng����ID   '��ʾ�ò����Ѿ��������������
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Get�������_�ڽ�(lng����ID As Long, lng��ҳID As Long) As String
    '����:��ȡ���������ʶ
    '     A-������B-��ת��C-δ����D-������E-����
    '??49  ���������ʶ    CHAR    439 1   1������2��ת��3δ����4������5������סԺ���� Ժ��
    'A-������B-��ת��C-δ����D-������E-����
    
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.��Ժ���" & _
             " From ������ A,��������Ŀ¼ B " & _
             " Where A.����ID=" & lng����ID & " And A.����ID=B.ID(+) And A.��ҳID=" & lng��ҳID & _
             "       And A.������� in (2,3)" & _
             " Order by A.������� Desc"
    
    rsInNote.CursorLocation = adUseClient
    Call zlDatabase.OpenRecordset(rsInNote, strTmp, "ҽ���ӿ�")
    strTmp = ""
    If Not rsInNote.EOF Then
        strTmp = Nvl(rsInNote!��Ժ���)
    End If
    strTmp = Decode(strTmp, "����", "0", "��ת", "1", "δ��", "2", "����", "3", "�Զ���Ժ", 4, "ת����ͳ�������ҽԺ", 5, "ת����ͳ�������ҽԺ", 6, "����", "3")
    Get�������_�ڽ� = strTmp
End Function

Private Function ����סԺ��ϸ��¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '���������ϸ��¼
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim StrInput  As String, strOutput As String
    Dim strArr, strArrժҪ
    Dim lng����ID As Long
    Err = 0
    On Error GoTo errHand:


    ����סԺ��ϸ��¼ = False

    '����δ�ϴ���ϸ�������Ա����ϴ�����ϸ�����ϴ�����ϸ��
    gstrSQL = "" & _
        "   Select distinct A.NO,A.��¼����,A.��¼״̬ " & _
        "   From סԺ���ü�¼ A " & _
        "   Where A.����ID=" & lng����ID & " and A.��ҳID=" & lng��ҳID & " and A.���ʷ���=1  and nvl(A.ʵ�ս��,0)<>0 and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 " & _
        "   Order by A.NO,A.��¼����,Decode(A.��¼״̬,2,2,1)"
        
    
   zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡ������ϸ��¼"
    '�ȼ���Ƿ�����˵������������ڣ����з��Ӧ�ļ�¼��.
    With rs��ϸ
'        '�ϴ���ϸ
'        If .RecordCount <> 0 Then .MoveFirst
'        Do While Not .EOF
'            If Nvl(!��Ŀ����) = "" Then
'                ShowMsgbox "��Ŀ:[" & Nvl(!����) & "] δ���ö�Ӧ��ҽ����Ŀ,�����ö�Ӧ��ϵ!"
'                Exit Function
'            End If
'            If (Val(!����) < 0 Or Val(!�۸�) < 0) And rs��ϸ!��¼״̬ = 1 Then
'                ShowMsgbox "��Ŀ:[" & Nvl(!����) & "] �������븺����!"
'                Exit Function
'            End If
'            .MoveNext
'        Loop
    End With
    '�ȴ�������
    With rs��ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!��¼״̬, 1) = 1 Or Nvl(!��¼״̬, 1) = 3 Then
                '�ϴ�ָ������
                If �����ϴ�(Nvl(!��¼����, 0), Nvl(!��¼״̬, 0), Nvl(!NO)) = False Then
                    Exit Function
                End If
            Else
                '�ϴ�ָ������
                gcnOracle_�ɶ��ڽ�.BeginTrans
                gcnOracle.BeginTrans
                If ��������(Nvl(!��¼����, 0), Nvl(!��¼״̬, 0), Nvl(!NO)) = False Then
                    gcnOracle.RollbackTrans
                    gcnOracle_�ɶ��ڽ�.RollbackTrans
                    Exit Function
                    
                End If
                gcnOracle.CommitTrans
                gcnOracle_�ɶ��ڽ�.CommitTrans
            End If
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
'Private Function Getԭ����ժҪ(ByVal strNO As String, ByVal int��� As Integer, ByVal int���� As Integer) As Variant
'    '����ָ����ֵ����ȡժҪ�������Ϣ
'    Dim rsTemp As New ADODB.Recordset
'    Dim strTemp As String
'
'
'    gstrSQL = " Select ժҪ From ���˷��ü�¼" & _
'              " Where NO='" & strNO & "' And ���=" & int��� & _
'              " And ��¼����=" & int���� & " And ��¼״̬=3"
'
'    Call OpenRecordset(rsTemp, "ȡԭʼ������ϸ����ˮ��")
'
'    If Not rsTemp.EOF Then
'        strTemp = Nvl(rsTemp!ժҪ) & "|||||||"
'    Else
'        strTemp = "|||||||"
'    End If
'    Getԭ����ժҪ = Split(strTemp, "|")
'End Function

'----200410���˺����
Public Function ҽ������_�ɶ��ڽ�() As Boolean
    ҽ������_�ɶ��ڽ� = frmSet�ɶ��ڽ�.��������
    
End Function
'
'Public Function ���ط�����ĿĿ¼_�ɶ��ڽ�(ByVal bytType As Byte, ByVal objProgss As Object) As Boolean
'    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    '����:���ط�����ĿĿ¼
'    '����:bytType-1-ҩƷ,2-����,3-����,4-�������,5-����Ŀ¼
'    '����:���سɹ�,����true,���򷵻�False
'    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSql As String
'    Dim rsTemp As New ADODB.Recordset
'    Dim strDate As String, strInput As String, strOutput As String
'    Dim lngCount As Long
'    Dim i As Long
'    Dim strArr
'    Dim strTitle As String
'
'    ���ط�����ĿĿ¼_�ɶ��ڽ� = False
'    strTitle = Switch(bytType = 1, "ҩƷ", bytType = 2, "������Ŀ", bytType = 3, "������ʩ", bytType = 4, "�������", True, "����Ŀ¼ȡ")
'
'    Err = 0
'    On Error GoTo ErrHand:
'    strSql = "" & _
'        "   Select to_char(Max(���ʱ��),'yyyy-mm-dd hh24:mi:ss')  as ���ʱ�� " & _
'        "   From ҽ���շ�Ŀ¼ " & _
'        "   where ���=" & bytType
'    zlDatabase.OpenRecordset rsTemp, strSql, "��ȡ�����ʱ��"
'
'    strDate = Nvl(rsTemp!���ʱ��)
'    strDate = IIf(strDate = "", "1977-01-01 00:00:00", strDate)
'
'    If Not objProgss Is Nothing Then
'    Else
'        zlCommFun.ShowFlash "�������ء�" & strTitle & "������,��ȴ�..."
'    End If
'    'Ԥ����
'    strInput = bytType & "|" & strDate
'    If ҵ������_�ɶ��ڽ�(�շ�Ŀ¼����Ԥ����, strInput, strOutput) = False Then Exit Function
'    strArr = Split(strOutput, "|")
'    lngCount = Val(strArr(1))
'
'    If Not objProgss Is Nothing Then
'        objProgss.Max = IIf(lngCount = 0, 1, lngCount)
'        objProgss.Min = 1
'        objProgss.Value = 1
'    End If
'
'   For i = 1 To lngCount
'        '��������
'        If ҵ������_�ɶ��ڽ�(�շ�Ŀ¼���ش���, strInput, strOutput) = False Then Exit Function
'        strArr = Split(strOutput, "|")
'        '�����շ�Ŀ¼
'
'        '����:���,����,����,Ӣ������,�շ����,�շѵȼ�,���õȼ�,ƴ����,��λ,����,���,��ע,���ʱ��,��ά����־,֧����׼
'        strSql = "ZL_ҽ���շ�Ŀ¼_UPDATE("
'        strSql = strSql & bytType & ",'"
'        strSql = strSql & strArr(1) & "','" '����
'        strSql = strSql & strArr(2) & "','" '����
'        Select Case bytType
'        Case 1
'            strSql = strSql & strArr(3) & "','" 'Ӣ������
'            strSql = strSql & strArr(4) & "','" '�շ����
'            strSql = strSql & strArr(5) & "','" '���õȼ�
'            strSql = strSql & strArr(6) & "','" 'ƴ����
'            strSql = strSql & strArr(7) & "','" '��λ
'            strSql = strSql & strArr(8) & "','" '����
'            strSql = strSql & strArr(9) & "','" '����
'            strSql = strSql & strArr(10) & "','" '���
'            strSql = strSql & strArr(11) & "',to_date('" '��ע
'            strSql = strSql & strArr(12) & "','yyyy-mm-dd hh24:mi:ss'),'"  '���ʱ��
'            strSql = strSql & strArr(13) & "','"     '��ά����־
'            strSql = strSql & "" & "')" '֧����׼
'        Case 2
'            strSql = strSql & "" & "','" 'Ӣ������
'            strSql = strSql & strArr(3) & "','" '�շ����
'            strSql = strSql & "" & "','" '���õȼ�
'            strSql = strSql & strArr(4) & "','" 'ƴ����
'            strSql = strSql & strArr(5) & "','" '��λ
'            strSql = strSql & strArr(6) & "','" '����
'            strSql = strSql & "" & "','" '����
'            strSql = strSql & "" & "','" '���
'            strSql = strSql & strArr(7) & "',to_date('" '��ע
'            strSql = strSql & strArr(8) & "','yyyy-mm-dd hh24:mi:ss'),'"  '���ʱ��
'            strSql = strSql & strArr(9) & "','"     '��ά����־
'            strSql = strSql & "" & "')" '֧����׼
'        Case 3
'            strSql = strSql & "" & "','" 'Ӣ������
'            strSql = strSql & strArr(3) & "','" '�շ����
'            strSql = strSql & "" & "','" '���õȼ�
'            strSql = strSql & strArr(6) & "','" 'ƴ����
'            strSql = strSql & "" & "','" '��λ
'            strSql = strSql & strArr(4) & "','" '����
'            strSql = strSql & "" & "','" '����
'            strSql = strSql & "" & "','" '���
'            strSql = strSql & "" & "',to_date('" '��ע
'            strSql = strSql & strArr(7) & "','yyyy-mm-dd hh24:mi:ss'),'"  '���ʱ��
'            strSql = strSql & "" & "',"     '��ά����־
'            strSql = strSql & strArr(5) & "')" '֧����׼
'        Case 4
'            ' ����������|�����������
'            strSql = "ZL_ҽ���շ����_UPDATE("
'
'            strSql = strSql & strArr(1) & "','" '����
'            strSql = strSql & strArr(2) & "')" '����
'        Case Else
'            '���ֱ���|��������|ƴ����|�������
'            strSql = "ZL_ҽ������Ŀ¼_UPDATE("
'            strSql = strSql & strArr(1) & "','" '����
'            strSql = strSql & strArr(2) & "','" '����
'            strSql = strSql & strArr(3) & "',to_date('" '������
'            strSql = strSql & strArr(4) & "','yyyy-mm-dd hh24:mi:ss')" '���ʱ��
'        End Select
'        gcnOracle_�ɶ��ڽ�.Execute strSql, , adCmdStoredProc
'        If Not objProgss Is Nothing Then
'            objProgss.Value = i
'        Else
'            zlCommFun.ShowFlash "�������ء�" & strTitle & "������,������" & i & "/" & lngCount & ""
'        End If
'   Next
'   ���ط�����ĿĿ¼_�ɶ��ڽ� = True
'   Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'End Function

Public Function ��ȡ�α���Ա��Ϣ_�ɶ��ڽ�(ByVal StrInput As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------

    '��ȡ�α���Ա��Ϣ
    Dim strOutput As String
    Dim strArr
    
    Dim str�������� As String
    ��ȡ�α���Ա��Ϣ_�ɶ��ڽ� = False
    
    Err = 0
    On Error GoTo errHand:
    
    If ҵ������_�ɶ��ڽ�(��������Ϣ_�ڽ�, StrInput, strOutput) = False Then Exit Function
    '���ش���:����vbtab���˱��vbtab���֤��vbtab����vbtab�Ա�vbtab�������vbtab��������vbtab��λ����vbtabͳ����vbtab�ƿ�����vbtab����Ч��vbtab��������vbtab�ƿ���λ
    If strOutput = "" Then Exit Function
    strArr = Split(strOutput, vbTab)
    
    With g�������_�ɶ��ڽ�
        .���� = strArr(0)
        .���˱�� = strArr(1)
        .���֤�� = strArr(2)
        .���� = strArr(3)
        .�Ա� = strArr(4)
        .������� = strArr(5)
        '.�������� = zlCommFun.AddDate(strArr(6))
        '�¶� 20050601
        If ���֤��ת��������(strArr(2), str��������) = True Then
            .�������� = zlCommFun.AddDate(str��������)
        Else
            MsgBox str�������� & "�����ó��������ֶε�ֵ", vbInformation, gstrSysName
            .�������� = zlCommFun.AddDate(strArr(6))
        End If
        .��λ���� = strArr(7)
        .ͳ���� = strArr(8)
        .�ƿ����� = strArr(9)
        .����Ч�� = strArr(10)
        .�������� = strArr(11)
        .�ƿ���λ = strArr(12)
        .���� = Get����(.��������)
    End With
    
    '--��ȡ�����ʻ����
    If ��ȡ�ʻ����_�ɶ��ڽ�() = False Then Exit Function
    
    ��ȡ�α���Ա��Ϣ_�ɶ��ڽ� = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Public Function ��ȡ�ʻ����_�ɶ��ڽ�() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ���˵�ǰ�ʻ����
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, strOutput As String
    Dim strArr
    Err = 0
    On Error GoTo errHand:
    ��ȡ�ʻ����_�ɶ��ڽ� = False
    With g�������_�ɶ��ڽ�
        '    ���˱��    String (8)  IN
        '    �籣������  String (10) IN
        '    ͳ���������    String (1)  IN
        StrInput = .���˱��
        StrInput = StrInput & vbTab & .����
        StrInput = StrInput & vbTab & .ͳ����
    End With
    
    If ҵ������_�ɶ��ڽ�(��ȡ�ʻ����_�ڽ�, StrInput, strOutput) = False Then Exit Function
    If strOutput = "" Then Exit Function
    strArr = Split(strOutput, vbTab)
    With g�������_�ɶ��ڽ�
        .�ʻ���� = Val(strArr(0))
        .��ְ��� = strArr(1)
    End With
    ��ȡ�ʻ����_�ɶ��ڽ� = True
    Exit Function
errHand:
        If ErrCenter = 1 Then Resume
End Function

Private Function Get����(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-[1])/365 as ���� from dual "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", CDate(strDate))
    If Not rsTemp.EOF Then
        Get���� = Int(Nvl(rsTemp!����, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function

Private Function GetErrInfor(ByVal strErrCode As String) As String
        Dim strErrMsg As String
        
        Select Case strErrCode
                '---��������ش���
                Case "1 ": strErrMsg = " ����ţ�1 " & vbCrLf & " �������������ͨѶ��ʽ����(chk_baud�쳣����)"
                Case "2 ": strErrMsg = " ����ţ�2 " & vbCrLf & " ����������ʼ���˿ڴ���(auto_init)"
                Case "3 ": strErrMsg = " ����ţ�3 " & vbCrLf & " �����������ر�ͨѶ�ڴ���(ic_exit)"
                Case "4 ": strErrMsg = " ����ţ�4 " & vbCrLf & " ������������д������"
                Case "5 ": strErrMsg = " ����ţ�5 " & vbCrLf & " �����������޷���ʼ��������"
                Case "10": strErrMsg = " ����ţ�10" & vbCrLf & " ���������� ����д�����Ƿ��п�����(get_status)"
                Case "11": strErrMsg = " ����ţ�11" & vbCrLf & " ���������� ���ʹ��󣨷�4428����(chk_4428)"
                Case "12": strErrMsg = " ����ţ�12" & vbCrLf & " ���������� ���������(csc_4428)"
                Case "13": strErrMsg = " ����ţ�13" & vbCrLf & " ���������� �޸Ŀ��������"
                Case "20": strErrMsg = " ����ţ�20" & vbCrLf & " ���������� ����о���ݴ���(srd_4428)"
                Case "21": strErrMsg = " ����ţ�21" & vbCrLf & " ���������� д��о����(�û�����)����(swr_4428)"
                Case "23": strErrMsg = " ����ţ�23" & vbCrLf & " ���������� д��о����(�û�����)����(swr_4428)"
                Case "30": strErrMsg = " ����ţ�30" & vbCrLf & " ���������� �û��������"
                Case "31": strErrMsg = " ����ţ�31" & vbCrLf & " ���������� �û����ݼ��ܴ���(ic_decrypt)"
                Case "32": strErrMsg = " ����ţ�32" & vbCrLf & " ���������� �û����ݽ��ܴ���(ic_decrypt)"
                Case "33": strErrMsg = " ����ţ�33" & vbCrLf & " ���������� �û�������ܴ���(ic_encrypt)"
                Case "34": strErrMsg = " ����ţ�34" & vbCrLf & " ���������� �û�������ܴ���(ic_decrypt)"
                Case "35": strErrMsg = " ����ţ�35" & vbCrLf & " ���������� �û�ԭ���볤��Ϊ����ߴ���6"
                Case "36": strErrMsg = " ����ţ�36" & vbCrLf & " ���������� �û������볤��Ϊ����ߴ���6"
                Case "40": strErrMsg = " ����ţ�40" & vbCrLf & " ����������   ���ܴ����ݿ�"
                Case "41": strErrMsg = " ����ţ�41" & vbCrLf & " ����������   û���ƿ�����"
                Case "42": strErrMsg = " ����ţ�42" & vbCrLf & " ����������   ������Ϣ���������������Ա�������ı���Ϣ��"
                '---ҽ���ӿڷ��ص���ش���
                Case "000": strErrMsg = "ִ�гɹ�"
                Case "001": strErrMsg = " ����ţ� 001" & vbCrLf & " ��������������������Ӧ"
                Case "002": strErrMsg = " ����ţ� 002" & vbCrLf & " ����������û���籣��"
                Case "003": strErrMsg = " ����ţ� 003" & vbCrLf & " �����������籣������Ӧ"
                Case "004": strErrMsg = " ����ţ� 004" & vbCrLf & " ������������������Ӧ"
                Case "051": strErrMsg = " ����ţ� 051" & vbCrLf & " ���������������������"
                Case "052": strErrMsg = " ����ţ� 052" & vbCrLf & " �����������������籣���벻��"
                Case "053": strErrMsg = " ����ţ� 053" & vbCrLf & " ����������������ϸ��ʵ�ʼ�¼��������"
                Case "054": strErrMsg = " ����ţ� 054" & vbCrLf & " ����������û�д˽�����ˮ��"
                Case "055": strErrMsg = " ����ţ� 055" & vbCrLf & " ����������������Ŀ����"
                Case "056": strErrMsg = " ����ţ� 056" & vbCrLf & " ����������û�д�סԺ��ˮ�ţ�����ҽ�����׺�+�����סԺ��ˮ��������ҽ����ˮ�Ŷ�Ӧ��סԺ��ˮ�Ų�һ�£��ȵ�"
                Case "058": strErrMsg = " ����ţ� 058" & vbCrLf & " �����������ظ�ҵ��������磺��סԺ������סԺ�������ѳ�Ժ�����г�Ժ����"
                Case "059": strErrMsg = " ����ţ� 059" & vbCrLf & " ���������������籣����Ͷ�Ӧ������ˮ�Ų�һ��(סԺʱ��סԺ��ˮ�Ŷ�Ӧ)"
                Case "060": strErrMsg = " ����ţ� 060" & vbCrLf & " ������������ˮ�Ų�Ϊ�����ʱ"
                Case "061": strErrMsg = " ����ţ� 061" & vbCrLf & " ��������������δȷ���ϴ�סԺ����ǰ"
                Case "062": strErrMsg = " ����ţ� 062" & vbCrLf & " ����������δ����סԺ�Ǽ�"
                Case "071": strErrMsg = " ����ţ� 071" & vbCrLf & " ����������ҽԺ������ˮ���쳣��HISϵͳ���ɣ��գ����Ȳ�������"
                Case "072": strErrMsg = " ����ţ� 072" & vbCrLf & " �����������ظ����ݰ�����"
                Case "073": strErrMsg = " ����ţ� 073" & vbCrLf & " �����������������ݰ�����"
                Case "074": strErrMsg = " ����ţ� 074" & vbCrLf & " ����������Ӧ���ϴ���ϸ��û���ϴ���ϸ"
                Case "075": strErrMsg = " ����ţ� 075" & vbCrLf & " ���������������Ŀ�����쳣��Ŀ���಻��[1]��[2]֮��"
                Case "077": strErrMsg = " ����ţ� 077" & vbCrLf & " �����������ϴ�ҽ�ƻ��������쳣��KB01֮�в�����"
                Case "078": strErrMsg = " ����ţ� 078" & vbCrLf & " �����������Ƕ���ҽ�ƻ���"
                Case "079": strErrMsg = " ����ţ� 079" & vbCrLf & " �����������籣������ҽ�����׺Ų���Ӧ�������ѳ���ʱ��ֻ���������ѵĿ�����������"
                Case "080": strErrMsg = " ����ţ� 080" & vbCrLf & " ����������ҽ�ƻ�������뽻����ˮ�Ų���Ӧ(סԺʱ��סԺ��ˮ�Ŷ�Ӧ)ֻ������ҽ�ƻ��������Լ��Ľ���"
                Case "081": strErrMsg = " ����ţ� 081" & vbCrLf & " ������������Ӧ��ˮ����ϸ������Kc07,KC08K1,Kc08k2"
                Case "082": strErrMsg = " ����ţ� 082" & vbCrLf & " ����������û�ж�Ӧ��סԺ��ˮ��Kc08"
                Case "083": strErrMsg = " ����ţ� 083" & vbCrLf & " �������������������ѳ�Ժ�Ľ����Ѿ���Ժ�Ľ��ײ�������"
                Case "084": strErrMsg = " ����ţ� 084" & vbCrLf & " ������������ԺҽԺ������ԺҽԺ"
                Case "085": strErrMsg = " ����ţ� 085" & vbCrLf & " ������������������������ȴ������볤����Լ�����Ȳ���"
                Case "086": strErrMsg = " ����ţ� 086" & vbCrLf & " ������������Ժ��Ա����ԭ���Ǹ�סԺ��Ա��ֹһ���Ŷ�Ӧ����˱�ŵ����"
                Case "101": strErrMsg = " ����ţ� 101" & vbCrLf & " ��������������״̬�쳣"
                Case "102": strErrMsg = " ����ţ� 102" & vbCrLf & " �����������籣��Ϊ��������"
                Case "103": strErrMsg = " ����ţ� 103" & vbCrLf & " �����������ʻ�������"
                Case "104": strErrMsg = " ����ţ� 104" & vbCrLf & " ������������������ͳ�����"
                Case "106": strErrMsg = " ����ţ� 106" & vbCrLf & " ���������������ڴ���"
                Case "107": strErrMsg = " ����ţ� 107" & vbCrLf & " ����������û�вμ�ҽ�Ʊ��ջ���������"
                Case "108": strErrMsg = " ����ţ� 108" & vbCrLf & " �����������ʻ���ע��ֻ���������������ӲŴ���ע��"
                Case "109": strErrMsg = " ����ţ� 109" & vbCrLf & " ������������Ժ���ϴ�����δȷ�ϳ�Ժ�ϴ�ʱ"
                Case "110": strErrMsg = " ����ţ� 110" & vbCrLf & " ������������Ժ��ȷ�ϳ�Ժȷ�ϴ�ʱ"
                Case "111": strErrMsg = " ����ţ� 111" & vbCrLf & " ����������û�д���ҽ��������T_cardinfo��"
                Case "112": strErrMsg = " ����ţ� 112" & vbCrLf & " ������������ʧ���Ѿ���ʧ"
                Case "113": strErrMsg = " ����ţ� 113" & vbCrLf & " ������������״̬�쳣"
                Case "114": strErrMsg = " ����ţ� 114" & vbCrLf & " ���������������ڸ����˻�Kc04������"
                Case "115": strErrMsg = " ����ţ� 115" & vbCrLf & " ����������ϵͳû�е������߲���"
                Case "116": strErrMsg = " ����ţ� 116" & vbCrLf & " ����������û�е������ҽ�Ʊ��ձ�������"
                Case "117": strErrMsg = " ����ţ� 117" & vbCrLf & " ����������û�е����ҽ�Ʊ��ձ�������"
                Case "118": strErrMsg = " ����ţ� 118" & vbCrLf & " ����������û�е���߶�ҽ�Ʊ��ձ�������"
                Case "119": strErrMsg = " ����ţ� 119" & vbCrLf & " ����������û�е��깫��Աҽ�Ʊ��ձ�������"
                Case "120": strErrMsg = " ����ţ� 120" & vbCrLf & " ������������Ժʱ����Ч"
                Case "121": strErrMsg = " ����ţ� 121" & vbCrLf & " ����������û�е�λ������ϢKb02"
                Case "150": strErrMsg = " ����ţ� 150" & vbCrLf & " ����������֧�����Ϊ���������֧���Ľ��Ϊ����"
                Case "151": strErrMsg = " ����ţ� 151" & vbCrLf & " ���������������ʻ���֧���ۼ�Ϊ�����˻��۳�������"
                Case "152": strErrMsg = " ����ţ� 152" & vbCrLf & " ��������������ͳ�ﳬ֧"
                Case "153": strErrMsg = " ����ţ� 153" & vbCrLf & " ������������ͳ�ﳬ֧"
                Case "154": strErrMsg = " ����ţ� 154" & vbCrLf & " ��������������Աҽ�Ʋ�����֧"
                Case "155": strErrMsg = " ����ţ� 155" & vbCrLf & " �����������˻�֧�����ó�������Ӧ��֧������"
                Case "255": strErrMsg = " ����ţ� 255" & vbCrLf & " ��������������������"
                Case "998": strErrMsg = " ����ţ� 998" & vbCrLf & " ������������ȡ������ˮ��ʧ��"
                Case "999": strErrMsg = " ����ţ� 999" & vbCrLf & " �������������ݿ�sql���󣬻���δ�ҵ�����"
                Case "800": strErrMsg = " ����ţ� 800" & vbCrLf & " ����������δ�����Ժ������ȴ��ͼ���г�Ժȷ�ϲ���"
                Case "801": strErrMsg = " ����ţ� 801" & vbCrLf & " �����������Ѿ���Ժȷ�ϣ��ٴ���ͼ�����Ժȷ��"
                '20051020 �¶� add
                Case "122": strErrMsg = " ����ţ� 122" & vbCrLf & " ��������������֧����Ŀ��������"
                
                Case "200": strErrMsg = " ����ţ� 200" & vbCrLf & " �������������ʿ�ʼ���ڵ���ֹ����û�з�����ϸ"
                Case "201": strErrMsg = " ����ţ� 201" & vbCrLf & " �����������������ж�����Ϣ�������ٴζ���"
                Case "202": strErrMsg = " ����ţ� 202" & vbCrLf & " ��������������������"
                Case "203": strErrMsg = " ����ţ� 203" & vbCrLf & " ���������� û����ϵͳָ�����ڶ���"
                    
                Case "210": strErrMsg = " ����ţ� 210" & vbCrLf & " �����������ϴ���ϴ����ظ�"
                Case "211": strErrMsg = " ����ţ� 211" & vbCrLf & " �����������ϴ���ϴ��벻����"
                    
                Case "220": strErrMsg = " ����ţ� 220" & vbCrLf & " ����������һ��סԺ���ܶ������"
                Case "221": strErrMsg = " ����ţ� 221" & vbCrLf & " ����������һ�����ﲻ�ܶ������"
                Case "222": strErrMsg = " ����ţ� 222" & vbCrLf & " ����������ҩ�겻���ϴ����������"
                Case "223": strErrMsg = " ����ţ� 223" & vbCrLf & " ����������û�з�������ʱ�����ϴ���̥����"
                Case "224": strErrMsg = " ����ţ� 224" & vbCrLf & " ����������û���ʹ���ʱ�����ϴ�ȫ���������"
                Case "225": strErrMsg = " ����ţ� 225" & vbCrLf & " �������������ﲻ���ϴ��ʹ������û�ȫ���������"
                    
                Case "230": strErrMsg = " ����ţ� 230" & vbCrLf & " ����������������סԺ�����ϴ���������"
                Case "231": strErrMsg = " ����ţ� 231" & vbCrLf & " ��������������סԺû�����벢��֢�����ϴ�ҽ�Ʒ���"
                
            Case Else
                strErrMsg = "����ȷ���Ĵ������,�����Ϊ" & strErrCode
    End Select
    GetErrInfor = strErrMsg
End Function
Public Sub ExecuteProcedure_ZLNJ(ByVal strCaption As String)
'���ܣ�ִ��SQL���
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_�ɶ��ڽ�.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Private Function ���㷽ʽ����(��� As Integer, Optional ByRef strAdvance = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������ʾ������
    '--�����:
    '--������:str���㷽ʽ
    '--��  ��:�ɹ�����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, str������㷵�ش� As String
    Dim dbl�����ܶ� As Double
        
    '�����ܶ�=�����Էѽ��+����ͳ��֧�����+��ͳ����      �˽����������˺�������湫ʽת��������
    
    '�����Էѽ�� = �ܷ��ö� - ����ͳ��֧����� - �� / �߶�ͳ��֧�����
    '�Էѽ��ֽ�֧����ʻ�֧���� (��:��ѡ�����ֽ�����ʻ�֧��)
    '��ͳ����߶�ͳ��������ͬ
    'ͳ��֧��������ҽ���ڷ��ø��ݲ�ͬ���𸶱�׼�ͱ���������ҽ��������
    '��˵�����ݱ��������漼�������ɷ����޹�˾�������Ľ���
    ���㷽ʽ���� = False
    
    Err = 0
    On Error GoTo errHand:
    DebugTool "����(" & "Get���㷽ʽ" & ")"
    
'    If g��������.�����־ = 0 Then
        dbl�����ܶ� = g��������.ҽ���ڷ��� + g��������.ҽ�������
'    End If
'    If ��� = 1 Then
        '������ 20051028
        'g��������.�ʻ�֧�� = dbl�����ܶ� - g��������.ҽ������� - g��������.����ҽ��֧�� - g��������.�߶�ҽ��֧�� - g��������.����Աҽ�Ʋ��� - g��������.����֧��
        g��������.�ʻ�֧�� = dbl�����ܶ� - g��������.ҽ������� - g��������.����ҽ��֧�� - g��������.�߶�ҽ��֧�� - g��������.����Աҽ�Ʋ���
'    Else
'        g��������.�ʻ�֧�� = dbl�����ܶ� - g��������.����ҽ��֧�� - g��������.�߶�ҽ��֧�� - g��������.����Աҽ�Ʋ���
'    End If
    'by 20050122 gzy
    If g��������.�ʻ�������� >= 0 Then
       If g��������.�ʻ�������� < g��������.�ʻ�֧�� Then
          g��������.�ʻ�֧�� = g��������.�ʻ��������
       End If
    Else
        g��������.�ʻ�֧�� = 0
    End If
    str���㷽ʽ = "||�����ʻ�|" & g��������.�ʻ�֧��
    str������㷵�ش� = "�����ʻ�;" & g��������.�ʻ�֧�� & ";1"
    
    If g��������.����ҽ��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||����ͳ��|" & g��������.����ҽ��֧��
        str������㷵�ش� = str������㷵�ش� & "|����ͳ��;" & g��������.����ҽ��֧�� & ";0"
    End If
    If g��������.�߶�ҽ��֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||��֧��|" & g��������.�߶�ҽ��֧��
        str������㷵�ش� = str������㷵�ش� & "|��֧��;" & g��������.�߶�ҽ��֧�� & ";0"
    End If
    If g��������.����Աҽ�Ʋ��� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & g��������.����Աҽ�Ʋ���
        str������㷵�ش� = str������㷵�ش� & "|����Ա����;" & g��������.����Աҽ�Ʋ��� & ";0"
    End If
    '20051020 �¶�
    '>Beging ����֧��
    g��������.����ӯ�� = 0
    If g��������.����֧�� <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "||����֧��|" & g��������.����֧��
        str������㷵�ش� = str������㷵�ش� & "|����֧��;" & g��������.����֧�� & ";0"
    End If
    '>end ����֧��
    If Format(g�������_�ɶ��ڽ�.�����ܶ�, "###0.00;-###0.00;0;0") <> Format(dbl�����ܶ�, "###0.00;-###0.00;0;0") Then
        Dim blnYes As Boolean
        '�����ܶ���ҽ�����ķ����ܶ��,���ܽ��н���
        ShowMsgbox "���ν����ܶ�(" & g�������_�ɶ��ڽ�.�����ܶ� & ") ��" & vbCrLf & _
                    "   ���ķ��ص��ܶ�(" & dbl�����ܶ� & ")���ȣ����ܽ��㣡"
        Exit Function
    End If
    If g��������.����֧�� > g�������_�ɶ��ڽ�.�����ܷ��� Then
       ShowMsgbox "���η����������ã�(" & g�������_�ɶ��ڽ�.�����ܷ��� & ") С��" & vbCrLf & _
                    "   ���ķ��ص�����֧����(" & g��������.����֧�� & ")���ܽ��㣡"
       Exit Function
    End If
    
    strAdvance = str������㷵�ش�
    ���㷽ʽ���� = True
    Exit Function
    
    'Modified by ZYB 20051123 ���ڴ���������㣬���Բ�����У����������
    '�������,�򱣴��Ԥ����¼��
'    If str���㷽ʽ <> "" Then
'        str���㷽ʽ = Mid(str���㷽ʽ, 3)
'        g�������_�ɶ��ڽ�.���㷽ʽ = str���㷽ʽ
'
'        If g��������.�����־ = 0 Then
'            #If gverControl < 2 Then
'                gstrSQL = "zl_���˽����¼_Update(" & g��������.����ID & ",'" & str���㷽ʽ & "', 0)"
'            #Else
'                strAdvance = str���㷽ʽ
'                gstrSQL = "zl_ҽ���˶Ա�_Insert(" & g��������.����ID & ",'" & str���㷽ʽ & "')"
'            #End If
'            Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
''        Else
''                gstrSQL = "zl_���˽����¼_Update(" & g��������.����ID & ",'" & str���㷽ʽ & "',1)"
''                Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
'        End If
'    End If
'
'    '��ʾ������Ϣ
'    '"�����ʻ�:" & g��������.�ʻ�֧��
'    #If gverControl < 2 Then
'        If frm������Ϣ.ShowMe(g��������.����ID, False, , IIf(g��������.�����־ = 0, 0, 1)) = False Then
'            ���㷽ʽ���� = False
'            Exit Function
'        End If
'    #End If
    ���㷽ʽ���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function Get���״���(ByVal intType As ҵ������_�ɶ��ڽ�, Optional bln������ As Boolean = False) As String
    Select Case intType
        Case ��������Ϣ_�ڽ�
            Get���״��� = IIf(bln������, "��������Ϣ", "01")
        Case ��������_�ڽ�
            Get���״��� = IIf(bln������, "��������", "02")
        Case ��ȡ�ʻ����_�ڽ�
            Get���״��� = IIf(bln������, "��ȡ�ʻ����", "03")
        Case ������ϸд��_�ڽ�
            Get���״��� = IIf(bln������, "������ϸд��", "04")
        Case ��������ȷ��_�ڽ�
            Get���״��� = IIf(bln������, "��������ȷ��", "05")
        Case ��������ȡ��_�ڽ�
            Get���״��� = IIf(bln������, "��������ȡ��", "06")
        Case סԺ�Ǽ�_�ڽ�
            Get���״��� = IIf(bln������, "סԺ�Ǽ�", "07")
        Case ��Ժ�Ǽ��ϴ�_�ڽ�
            Get���״��� = IIf(bln������, "��Ժ�Ǽ��ϴ�", "08")
        Case סԺ�����ϴ�_�ڽ�
            Get���״��� = IIf(bln������, "סԺ�����ϴ�", "09")
        Case סԺ�����ϴ�ȡ��_�ڽ�
            Get���״��� = IIf(bln������, "סԺ�����ϴ�ȡ��", "10")
        Case ��Ժ�Ǽ�ȷ��_�ڽ�
            Get���״��� = IIf(bln������, "��Ժ�Ǽ�ȷ��_�ڽ�", "11")
        Case ��ȡ��λǷ�����_�ڽ�
            Get���״��� = IIf(bln������, "��ȡ��λǷ�����", "12")
        Case ��ʼ������_�ڽ�
            Get���״��� = IIf(bln������, "��ʼ������", "13")
        Case ���϶���_�ڽ�
            Get���״��� = IIf(bln������, "���϶���", "14")
        Case ����֢�����ϴ�_�ڽ�
            Get���״��� = IIf(bln������, "����֢�����ϴ�", "15")
        Case Else
            Get���״��� = IIf(bln������, "����Ľ��״���", "-1")
    End Select
End Function

Public Function ҵ������_�ɶ��ڽ�(ByVal intType As ҵ������_�ɶ��ڽ�, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������ҵ�����ҵ������
    '--�����:strinPutString-���봮,������˳��,��tab���ָ��Ĵ��봮
    '--������:strOutPutString-�����,������˳��,��tab���ָ��ķ��ش�
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String, lngReturn As Long, strReturn As String
    Dim strOutput(0 To 20) As String, dblOutPut(0 To 25) As Double, intOutPut(0 To 5) As Integer, lngOutPut(0 To 5) As Long
    Dim strArr1
    Dim strArr(0 To 20) As String
    Dim strҵ�� As String
    Dim strReg As String
    
    Dim strNetInput As String
    Dim lng���� As Long
    Dim i As Integer
    

    strҵ�� = Get���״���(intType, True)
    
    DebugTool "����ҵ��������(ҵ�����ʹ���Ϊ:" & intType & " ҵ�����ƣ�" & strҵ�� & ")" & vbCrLf & "        �������Ϊ:" & strInputString
    
    
    ҵ������_�ɶ��ڽ� = False
    
    StrInput = strInputString
    
    If InitInfor_�ɶ��ڽ�.ģ������ Then
        '��ȡģ������
        Readģ������ intType, strInputString, strOutPutstring
         ҵ������_�ɶ��ڽ� = True
        Exit Function
    End If
   
 'Modify �¶� 20051020
 'Beging ��������Ƿ���ͨ
    GetRegInFor g����ȫ��, "ҽ��", "ConfigFileName", strReg
    strNetInput = strReg
    GetRegInFor g����ȫ��, "ҽ��", "HostPort", strReg
    strNetInput = strNetInput & vbTab & strReg
    GetRegInFor g����ȫ��, "ҽ��", "IPAddress", strReg
    strNetInput = strNetInput & vbTab & strReg
    
    strArr1 = Split(strNetInput, vbTab)
    For i = 0 To UBound(strArr1)
        strArr(i) = strArr1(i)
    Next
    
    lngReturn = gobj�ɶ��ڽ�.SetCommPara(strArr(0), Val(strArr(1)), strArr(2))
    If lngReturn <> 1 Then
         '������ 20051028
         'ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
         ShowMsgbox "ҽ�����粻ͨ"
         Exit Function
    End If
    
    lngReturn = 0
 'End��������Ƿ���ͨ
 
    strArr1 = Split(strInputString, vbTab)
    For i = 0 To UBound(strArr1)
        strArr(i) = strArr1(i)
    Next
         
         
    For i = 0 To 20
        strOutput(i) = Space(100)
    Next
    
    Err = 0
    On Error GoTo errHand:
    
    Select Case intType
        Case ��������Ϣ_�ڽ�
            If InitInfor_�ɶ��ڽ�.������_�ڽ� = 0 Then
                '�������:
                ''����ʱҪ���� gobj�ɶ��ڽ�.
                lngReturn = GetCardInfo_MW(Val(strArr(0)), strArr(1), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), strOutput(5), strOutput(6), strOutput(7), strOutput(8), strOutput(9), strOutput(10), strOutput(11), strOutput(12))
            Else
                lngReturn = GetCardInfo_KRQ(Val(strArr(0)), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), strOutput(5), strOutput(6), strOutput(7), strOutput(8), strOutput(9), strOutput(10), strOutput(11), strOutput(12))
            End If
            For i = 0 To 20
                strOutput(i) = Trim(Split(strOutput(i), Chr(0))(0))
            Next
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(CStr(lngReturn))
                 Exit Function
            End If
           '�������ش�
           strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2) & vbTab & strOutput(3) & vbTab & strOutput(4) & vbTab & strOutput(5) & vbTab & strOutput(6) & vbTab & strOutput(7) & vbTab & strOutput(8) & vbTab & strOutput(9) & vbTab & strOutput(10) & vbTab & strOutput(11) & vbTab & strOutput(12)
        Case ��������_�ڽ�
            lngReturn = ChangePassword(Val(strArr(0)), strArr(1), strArr(2))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(CStr(lngReturn))
                 Exit Function
            End If
        Case ��ȡ�ʻ����_�ڽ�
            '�������:  ���˱��    String (8)  IN
            '           �籣������  String (10) IN
            '           ͳ���������    String (1)  IN
            '�������:  �ʻ����    Long    OUT
            '           ��ְ���    String(1)   OUT
            lngReturn = gobj�ɶ��ڽ�.GetAccountAmountFunc(strArr(0), strArr(1), strArr(2), lngOutPut(0), strOutput(0))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = lngOutPut(0) / 100 & vbTab & strOutput(0)
        Case ������ϸд��_�ڽ�
            '�������:  ���˱��    String(8)   In
            '           �籣������  String(10)  In
            '           ҽԺ����    String(5)   In
            '           ����Ա������    String(10)  In
            '           ͳ���������    String(1)   In
            '           ҽԺ������ˮ��  String(20)  In
            '           �������    String(1)   In
            '           ��������    String(2)   In
            '           ������ϸ    String����������51  In

            '�������:  ҽ����ˮ��  String(20)  Out
            '           ҽ���ڷ���  String(10)  Out
            '           ҽ�������  String(10)  Out
            '           �����ʻ��������    String(10)  Out
            '           ������ϸ    String����������51  Out
            '           ��ְ���    String(1)   Out
            '           '20051020 add
            '           ����֧��    String(10) Out
            Do While lng���� <= 3
                'lngReturn = gobj�ɶ��ڽ�.DoConsumeTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strOutPut(0), strOutPut(1), strOutPut(2), strOutPut(3), strOutPut(4), strOutPut(5))
                lngReturn = gobj�ɶ��ڽ�.DoConsumeTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strOutput(0), strOutput(1), strOutput(2), strOutput(3), strOutput(4), strOutput(5), strOutput(6))
                If lngReturn <> 0 Then
                    'If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "ҽ���ӿ�") = vbCancel Then
                    '    lng���� = 8
                   '     Exit Function
                   ' End If
                   lng���� = lng���� + 1
                   If lng���� > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "ҽ���ӿ�"
                        Exit Function
                   End If
                Else
                    lng���� = 5
                End If
            Loop
            'strReturn = strOutPut(0) & vbTab & Val(strOutPut(1)) / 100 & vbTab & Val(strOutPut(2)) / 100 & vbTab & Val(strOutPut(3)) / 100 & vbTab & strOutPut(4) & vbTab & strOutPut(5)
            strReturn = strOutput(0) & vbTab & Val(strOutput(1)) / 100 & vbTab & Val(strOutput(2)) / 100 & vbTab & Val(strOutput(3)) / 100 & vbTab & strOutput(4) & vbTab & strOutput(5) & vbTab & strOutput(6) / 100
            
        Case ��������ȷ��_�ڽ�
            '�������: ���˱��    String(8)   In
            '          �籣������  String(10)  In
            '          ҽԺ����    String(5)   In
            '          ����Ա������    String(10)  In
            '          ͳ���������    String(1)   In
            '          ҽ��������ˮ��  String(20)  In
            '          �������    String(1)   In
            '          �����ʻ�֧��    String(10)  In
            Do While lng���� <= 3
                lngReturn = gobj�ɶ��ڽ�.DoConsumeAffirmFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "ҽ���ӿ�") = vbCancel Then
'                        lng���� = 8
'                        Exit Function
'                    End If
                   lng���� = lng���� + 1
                   If lng���� > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "ҽ���ӿ�"
                        Exit Function
                   End If
                Else
                    lng���� = 5
                End If
            Loop
            strReturn = ""
        Case ��������ȡ��_�ڽ�
            '�������: ���˱��    String(8)   In
            '        �籣������  String(10)  In
            '        ҽԺ����    String(5)   In
            '        ����Ա������    String(10)  In
            '        ͳ���������    String(1)   In
            '        ҽ��������ˮ��  String(20)  In
            '        �������    String(1)   In
            Do While lng���� <= 3
                lngReturn = gobj�ɶ��ڽ�.DoConsumeCancelFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "ҽ���ӿ�") = vbCancel Then
'                        lng���� = 8
'                        Exit Function
'                    End If
                   lng���� = lng���� + 1
                   If lng���� > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "ҽ���ӿ�"
                        Exit Function
                   End If

                Else
                    lng���� = 5
                End If
            Loop
            strReturn = ""
        Case סԺ�Ǽ�_�ڽ�
            '�������: ���˱��    String(8)   In
            '        �籣������  String(10)  In
            '        ҽԺ����    String(5)   In
            '        ����Ա������    String(10)  In
            '        ͳ���������    String(1)   In
            '        ��Ժ����    String(8)   In
            '        ��Ժ�Ʊ�    String(10)  In
            '        ��Ժ����ҽ��    String(10)  In
            '        ��ϱ���    String(20)  In
            
            '�������:סԺ��ˮ��  String(20)  Out
            '        ���ܴ�����־    Small int   Out
            '        �𸶱�׼    Long    Out
            '        ��ְ���    String(1)   Out

            Do While lng���� <= 3
                lngReturn = gobj�ɶ��ڽ�.DoHospInFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strOutput(0), intOutPut(0), lngOutPut(0), strOutput(1))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "ҽ���ӿ�") = vbCancel Then
'                        lng���� = 8
'                        Exit Function
'                    End If
                   lng���� = lng���� + 1
                   If lng���� > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "ҽ���ӿ�"
                        Exit Function
                   End If

                Else
                    lng���� = 5
                End If
            Loop
            strReturn = strOutput(0) & vbTab & intOutPut(0) & vbTab & lngOutPut(0) & vbTab & strOutput(1)
        Case סԺ�����ϴ�_�ڽ�
            '�������: ���˱��    String(8)   In
            '        �籣������  String(10)  In
            '        ҽԺ����    String(5)   In
            '        ͳ���������    String(1)   In
            '        ҽԺ������ˮ��  String(20)  In
            '        ��Ժ��ҩ���    String(1)   In
            '        �Ʊ�    String(10)  In
            '        ҽ��    String(10)  In
            '        סԺ��ˮ��  String(20)  In
            '        ��������    String(2)   In
            '        ������ϸ    String����������51  In
            '�������:  ҽ��������ˮ��  String(20)  Out
            '           ������ϸ    String����������51  Out
            '           TRANSDETIAL��� (���������ϸ) Out
            Do While lng���� <= 3
                lngReturn = gobj�ɶ��ڽ�.DoHospTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), strArr(10), strOutput(0), strOutput(1), strOutput(2))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "ҽ���ӿ�") = vbCancel Then
'                        lng���� = 8
'                        Exit Function
'                    End If
                   lng���� = lng���� + 1
                   If lng���� > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "ҽ���ӿ�"
                        Exit Function
                   End If

                Else
                    lng���� = 5
                End If
            Loop
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2)
        Case סԺ�����ϴ�ȡ��_�ڽ�
            '�������:���˱��    String(8)   In
            '        �籣������  String(10)  In
            '        ����Ա������    String(10)  In
            '        ͳ���������    String(1)   In
            '        סԺ��ˮ��  String(20)  In
            '        ҽ��������ˮ��  String(20)  In
            '�������:

'            lngReturn = gobj�ɶ��ڽ�.DoHospCancelFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5))
'            If lngReturn <> 0 Then
'                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
'                 Exit Function
'            End If
            Do While lng���� <= 3
                lngReturn = gobj�ɶ��ڽ�.DoHospCancelFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "ҽ���ӿ�") = vbCancel Then
'                        lng���� = 8
'                        Exit Function
'                    End If
                   lng���� = lng���� + 1
                   If lng���� > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "ҽ���ӿ�"
                        Exit Function
                   End If

                Else
                    lng���� = 5
                End If
            Loop
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2)
        Case ��Ժ�Ǽ��ϴ�_�ڽ�
            '�������:���˱��    String(8)   In
            '        �籣������  String(10)  In
            '        ҽԺ����    String(5)   In
            '        ����Ա������    String(10)  In
            '        ͳ���������    String(1)   In
            '        ��Ժ����    String(8)   In
            '        ��Ժ�Ʊ�    String(10)  In
            '        ��Ժ����ҽ��    String(10)  In
            '        ��ϱ���    String(20)  In
            '        ��Ժ��ҩ    String(1)   In
            '        ��Ժ���    String(1)   In
            '        סԺ��ˮ��  String(20)  In
            '�������
            '        TRANSDETIAL��� (���������ϸ)
            '        ���ܴ�����־    String(1)   Out
            '        ҽ���ڷ���  String(10)  Out
            '        ҽ�������  String(10)  Out
            '        ����ҽ��֧�� ����μӴ�ҽ������Ϊ��ҽ��֧��  String(10)  Out
            '        �߶�ҽ��֧��    String(10)  Out
            '        ����Աҽ�Ʋ���  String(10)  Out
            '        ���˰�����֧��  String(10)  Out
            '        TRANSDETIAL����
            '        �𸶱�׼    String(10)  Out
            '        �����ʻ��������    String(10)  Out
'            lngReturn = gobj�ɶ��ڽ�.DoHospOutTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), strArr(10), strArr(11), strOutPut(0), strOutPut(1), strOutPut(2))
'            If lngReturn <> 0 Then
'                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
'                 Exit Function
'            End If
            Do While lng���� <= 3
                lngReturn = gobj�ɶ��ڽ�.DoHospOutTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), strArr(10), strArr(11), strOutput(0), strOutput(1), strOutput(2))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "ҽ���ӿ�") = vbCancel Then
'                        lng���� = 8
'                        Exit Function
'                    End If
                   lng���� = lng���� + 1
                   If lng���� > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "ҽ���ӿ�"
                        Exit Function
                   End If

                Else
                    lng���� = 5
                End If
            Loop
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2)
        Case ��Ժ�Ǽ�ȷ��_�ڽ�
        
            '�������:���˱��    String(8)   In
            '    �籣������  String(10)  In
            '    ����Ա������    String(10)  In
            '    ͳ��������    String(1)   In
            '    סԺ��ˮ��  String(20)  In
            '    �����ʻ�֧��    String(10)  In
'            lngReturn = gobj�ɶ��ڽ�.DoHospOutAffirmFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5))
'            If lngReturn <> 0 Then
'                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
'                 Exit Function
'            End If
            Do While lng���� <= 3
                lngReturn = gobj�ɶ��ڽ�.DoHospOutAffirmFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5))
                If lngReturn <> 0 Then
'                    If MsgBox(GetErrInfor(Lpad(lngReturn, 3, "0")), vbRetryCancel + vbDefaultButton1, "ҽ���ӿ�") = vbCancel Then
'                        lng���� = 8
'                        Exit Function
'                    End If
                   lng���� = lng���� + 1
                   If lng���� > 3 Then
                        MsgBox GetErrInfor(Lpad(lngReturn, 3, "0")), vbInformation, "ҽ���ӿ�"
                        Exit Function
                   End If

                Else
                    lng���� = 5
                End If
            Loop
            strReturn = ""
        
        Case ��ȡ��λǷ�����_�ڽ�
            '�������:���˱��    String (8)  IN
            '        �籣������  String (10) IN
            '        ͳ���������    String (1)  IN
            '�������
            '        ��λǷ�����    String(1)   OUT
            lngReturn = gobj�ɶ��ڽ�.GetArrearInfo(strArr(0), strArr(1), strArr(2), strOutput(0))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
                 Exit Function
            End If

            strReturn = strOutput(0)
        Case ��ʼ������_�ڽ�
            '�������:ConfigFileName
            '        HostPort
            '        IPAddress
            lngReturn = gobj�ɶ��ڽ�.SetCommPara(strArr(0), Val(strArr(1)), strArr(2))
            If lngReturn <> 1 Then
                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = ""
        'Add �¶� 20051020
        Case ���϶���_�ڽ�
        '�����������
            'HOSPID  ҽԺ/ҩ����
            'TCDQBM  ͳ���������
            'DZLB    �������(0:����,1:סԺ, 2ҩ��)
            'KSRQ    ���ʿ�ʼ����
            'ZZRQ    ������ֹ����
            'Count   �ϴ���������
            'je      �ϴ��ܶ�
        '�������
            'DZQK    �������(0�ɹ�,1 ������ , �������� 2 ���� , ������� 3����,��������)
            'DZCOUNT ������������
            'DZJE    ���ʽ��
            lngReturn = gobj�ɶ��ڽ�.CompareTotal(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strOutput(0), strOutput(1), strOutput(2))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = strOutput(0) & vbTab & strOutput(1) & vbTab & strOutput(2)
        Case ����֢�����ϴ�_�ڽ�
        'TCDQBM        'ͳ���������    String (1)  IN
        'ZYLSH        'סԺ��ˮ��  String (20) IN
        'ZDBM        '��ϱ���    String (200)    IN
            lngReturn = gobj�ɶ��ڽ�.DoBFZAffirmFunc(strArr(0), strArr(1), strArr(2))
            If lngReturn <> 1 Then
                 ShowMsgbox GetErrInfor(Lpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = ""
    End Select
    strOutPutstring = strReturn
    ҵ������_�ɶ��ڽ� = True
    DebugTool "  �������Ϊ:" & strReturn
     Exit Function
errHand:
    DebugTool "ҵ������ʧ��  �������Ϊ:" & strReturn
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetItemInsure_�ɶ��ڽ�(lng����ID As Long, lng�շ�ϸĿID As Long, bln���� As Boolean) As String
    Dim strDefault As String            'ȱʡҽ������
    Dim strCurrent As String            '��ǰҽ�����룬����ȡ������룬סԺȡסԺ����
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "Select B.���,A.����,A.����,B.˵�� From ������Ŀ A,ҽ��������ϸ B" & _
        " Where B.����=[1] And A.����=B.���� And A.����=B.��Ŀ���� And B.�շ�ϸĿID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ������", TYPE_�ɶ��ڽ�, lng�շ�ϸĿID)
    rsTemp.Filter = "���=" & IIf(bln����, 1, 2)
    Select Case rsTemp.RecordCount
    Case 0
        'û�����ö�Ӧ���룬ȡȱʡ����
        rsTemp.Filter = "���=0"
        If rsTemp.RecordCount <> 0 Then
            GetItemInsure_�ɶ��ڽ� = rsTemp!����
        End If
    Case 1
        GetItemInsure_�ɶ��ڽ� = rsTemp!����
    Case Else
        '��ѡ
        GetItemInsure_�ɶ��ڽ� = frmҽ����Ŀѡ��.ShowSelect(rsTemp, lng�շ�ϸĿID)
    End Select
    
    rsTemp.Filter = 0
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    rsTemp.Filter = 0
End Function

Public Function ����֢ѡ��_�ɶ��ڽ�(lng����ID As Long, lng��ҳID As Long)
    '���� �ϴ�����֤
    '20051024 �¶�
    Dim rsBfz As New ADODB.Recordset
    gstrSQL = "Select * from ������ҳ Where ��Ժ���� is null and ����ID=[1] And ��ҳID=[2]"
    Set rsBfz = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ���Ժ", lng����ID, lng��ҳID)
    If rsBfz.RecordCount > 0 Then
        gstrSQL = "Select * from �����ʻ� Where ����ID=[1]"
        Set rsBfz = zlDatabase.OpenSQLRecord(gstrSQL, "ȡͳ���������", lng����ID)
    Else
        MsgBox "��Ժ���˲���ִ�д˲���!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If ҽ����ʼ��_�ɶ��ڽ� = False Then Exit Function
    
    If ����δ�����(lng����ID, lng��ҳID) = True Then
        frm����ѡ��_�ɶ��ڽ�.GetCode (lng����ID)
    Else
        MsgBox "�����ѽ��壬����ִ�д˲�����", vbInformation, gstrSysName
        Exit Function
    End If
    
End Function

Public Function �жϲ���Ч��_�ɶ��ڽ�(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim int��Ч���� As Integer
    Dim dbl������� As Double

    gstrSQL = "Select nvl(����ֵ,0) as ����ֵ From ���ղ��� where ����=" & TYPE_�ɶ��ڽ� & " and ������='����������'"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����������"
    If rsTemp.EOF Then
       �жϲ���Ч��_�ɶ��ڽ� = False
       MsgBox "û�з��ֲ�����������������", vbInformation, gstrSysName
       Exit Function
    End If
    int��Ч���� = rsTemp!����ֵ
    
    gstrSQL = "Select trunc(sysdate-��Ժ����,2) as ������� From ������ҳ where ����id =" & lng����ID & " and ��ҳid =" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�������"
    If rsTemp.EOF Then
       �жϲ���Ч��_�ɶ��ڽ� = False
       MsgBox "�޷���ȡ���������", vbInformation, gstrSysName
       Exit Function
    End If
    dbl������� = rsTemp!�������
    
    If int��Ч���� < dbl������� And int��Ч���� > 0 Then
       �жϲ���Ч��_�ɶ��ڽ� = False
       MsgBox "�Ѿ�����������Ĳ���������", vbInformation, gstrSysName
       Exit Function
    End If
    �жϲ���Ч��_�ɶ��ڽ� = True
End Function
