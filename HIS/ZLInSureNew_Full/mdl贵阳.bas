Attribute VB_Name = "mdl����"
Option Explicit
#Const gblnTest = 0     '1-����

'�����޸�˵����
'�޸�ʱ�䣺2005-01-14
'�޸��ˣ�����
'�޸����ݣ����������ӿ�(SetBearingFlag��UploadICD)�������󲿷ֽӿ���������������

'�޸�ʱ�䣺2008-07-03
'�޸��ˣ�����
'�޸����ݣ����Ӿ��񣬹��˱���
'��ϸ�嵥����
'1 ?���ӡ�����ȡ�����϶�����: GETGSINFO ok
'2 ?���ӡ�����ѯ�������ⲡ�ַ�����ϸ����: QUERYCENSPECFEELIST   ��ǰ��ѯ��û��������Ҳ����
'3�����ӡ������˱������������볷����APPRECG��DELRECG    ok
'4������ֻ��������������סԺ������ֻ������ͨ������סԺ��HIS����Ҫ�����𣿲���Ҫ ok
'5����ȡҽԺ�������������ݺ�����QUERYHOSPSINGLEILLNESS��Ϊ��ֻ֧��ְ��ҽ�������ҽ�� ok
'6����ѯҽԺ�����ְ��ɽ���Ŀ¼��QUERYHOSPSINGLEILLNESS_BG��ͬ��     ok
'7�����ӡ�����ѯ�������ⲡ�ַ�����ϸ���ݣ�QUERYCENSPECFEELIST   ��2�ظ���
'8 ?��ͨ��������������: �����϶���� ok
'9������ҽ��ʹ������������㣿����������Ǿ������е���  ok
'10����Ժ�Ǽ�֧�ֹ��˼�����ҽ����ȡ����Σ��α�ǰ����Ժ  ok
'11 ?סԺ�������ȡ����Ժ: ת����־ û�ҵ��˲���
'12��ҽ�Ʊ����걨���㣺APPRECM ��εĲ���������̫�������ھ��� ok
'13 ?����ҽ�Ʊ�������: DELRECM �������: �������   ok
'14 ?�����������������   �ޱ仯��,���Ƶ�ǰ�Աȴ���
'15����Ա���    11����ְ��21�����ݣ�32��ʡ�����ݣ�34���������ݣ�41����ͨ����42���ͱ�����43��������Ա��44���������ͥ��45���ضȲм��� ok
'16��ҽԺ���    01��һ����02��������03��������04��������������05��ҩ�ꣻ06������ҽԺ��09���Ƕ���
'17���������    1:��ҵ����ҽ�Ʊ��գ�2:��ҵ����ҽ�Ʊ��գ�3:������ҵ��λ����ҽ�Ʊ��գ�4����ҵ�������գ�5��������ҵ��λ�������գ�6������ҽ����7�����˱���  ok
'18����Ŀ֧�����    11-��ͨ���� 12-��ͨסԺ 21-�������� 22-����סԺ 31-�������� 32-����סԺ 41-�������� 42-����סԺ

Public mdomInput As MSXML2.DOMDocument
Public mdomOutput As MSXML2.DOMDocument
Public tdomInput As MSXML2.DOMDocument
'========================================================================================
'=��־����
'========================================================================================
Public Enum LogType
    DBConnLTNew = 0
    DBConnLTEdit = 1
    DBConnLTDelete = 2
    DBConnLTSping = 3
End Enum

Public gblnHIS1026 As Boolean 'HIS�汾�Ƿ���10.26�����ϰ汾
Public gblnHIS1029 As Boolean 'HIS�汾�Ƿ���10.29�����ϰ汾����10.29.40��
Public mlngCloseTime As Long '������㴰���Զ��ر�ʱ��
Public gbln������ҩ���� As Boolean, gint�ۼ���ҩ�������׼ As Integer, ging������� As Integer
Public gint����ҽ������ As Integer
Private mblnInit As Boolean
Private mblnFail As Boolean
Public mstr���� As String
Public mstr���� As String
Public mstrҽ���� As String
Private mdbl��� As Double
Private mlng����ID As Long
Private mblnҽ����Ժ As Boolean         '��Ժʱ�Ƿ�ͬ������ҽ����Ժ
Private mbln����ҩƷ��ʾ As Boolean
Private mbln��������� As Boolean
Public gstr�����϶���� As String
Public gint���˿���סԺ As Integer      '0-��;1-��

Public gintType As Integer              '������
Public gstrSNO As String                '�籣����
Public gstrIDNO As String               '���֤��
Public gstrPSAMNO As String             'PSAMоƬ���
Public gstrClientIP As String           '�ͻ���IP

Public dblTOTAL As Double
'�����ڹ���zyq20110812����
Public gstr������Ϣ As String
Public gstr�����־ As String
Public gstr���˱�־ As String
'����������
Private mint���㷽ʽ As Integer
Private mstr�����ֱ��� As String
Public gstr�������� As String

Public gobj���� As Object
Private obj���� As Object
Public Const mstrҽ�����ı���_���� As String = "0101"
Public gcnGYYB As New ADODB.Connection

Private Type balance
    dblҽ������ As Double
    dbl�����ʻ� As Double
    dbl��ͳ�� As Double
    dblҽ�Ʋ��� As Double
    dbl������ As Double
End Type

'�൥���շ�ʱ��ʾ��
Public Type t��������
    dblҽ������ As Double
    dbl�󲡻��� As Double
    dbl����Ա���� As Double
    dbl�����ʻ� As Double
    dbl�ֽ� As Double
    dbl�����ܶ� As Double
    dbl������ As Double
End Type
Public g�������� As t��������

Public gbln����������� As Boolean      '���Ϊtrue,������msgbox��ʾ,������ʾ,��Ҫ�����������������
Private preBalance As balance

Private mnodRowset As MSXML2.IXMLDOMElement

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long) As Long

'--������غ���
Public Declare Function SGZ_IFD_Open Lib "SGZ_SSSReader.dll" (ByVal iReaderPort As Long, ByRef iReaderHandle As Long, ByVal iERRInfo As String) As Long
Public Declare Function SGZ_SAM_ReadNmuber Lib "SGZ_SSSReader.dll" (ByVal iReaderHandle As Long, ByVal iOutFileData As String, ByVal iERRInfo As String) As Long
Public Declare Function SGZ_ICC_ReadCardInfo Lib "SGZ_SSSReader.dll" (ByVal iReaderHandle As Long, ByVal iCardType As Long, ByVal iPassword As String, ByVal iInputFileAddr As String, ByVal iOutFileData As String, ByVal iERRInfo As String) As Long
Public Declare Function SGZ_IFD_GetPIN Lib "SGZ_SSSReader.dll" (ByVal iReaderHandle As Long, ByVal iDevType As Long, ByVal szPasswd As String, ByVal iERRInfo As String) As Long
Public Declare Function SGZ_IFD_Close Lib "SGZ_SSSReader.dll" (ByVal iReaderHandle As Long, ByVal iERRInfo As String) As Long

'���ò������
'========================================================================================
'==������Ľ���ɹ�������δ�ɹ��Ľ����¼���в���
'==XieRong 2010-10-12
'========================================================================================
Public Type typ�������

    blnYn                       As Boolean
    str����                     As String
    strҽ����                   As String
    lng����ID                   As Long
    lng��ҳID                   As Long
    STR����                     As String
    strסԺ��                   As String
    m_ȫ�Ը�                    As Double
    m_�ҹ��Ը�                  As Double
    m_����                    As Double
    m_�����Ը�                  As Double
    m_ͳ��֧��                  As Double
    m_ͳ���Ը�                  As Double
    m_��ͳ��                  As Double
    m_���Ը�                  As Double
    m_�����Ը�                  As Double
    m_ҽ���ܷ���                As Double
    m_����Ա����                As Double
    m_�����ܷ���                As Double
    m_HIS�ܷ���                 As Double
    m_������                  As String
    m_����˳���                As String
    m_��������                  As String
    m_����Ա�����𸶱�׼        As Double
    m_����Ա��������          As Double
    m_��ͨ���﹫��Ա�����ۼ�    As Double
    m_������޶��Ա����      As Double
    m_������㷽ʽ              As String
    m_�������˵��              As String

End Type
Public g�������        As typ�������
Private mlng����ID              As Long


Public Function ҽ����ʼ��_����() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim strUser As String, strPass As String, strServer As String
    Dim rsTemp As New ADODB.Recordset
    
    If mblnInit Then
        ҽ����ʼ��_���� = True
        Exit Function
    End If
    If mblnFail Then Exit Function
    
    On Error Resume Next
    If gstrClientIP = "" Then gstrClientIP = zl_Ip_Address_FromOrc(gstrClientIP)
    
    gint����ҽ������ = GetSetting(appName:="ZLSOFT", Section:="����ȫ��", Key:="����ҽ������", Default:="1")
    If gint����ҽ������ = 1 Then
        '��Ҫ����ҽ��ʱ
        Set mdomInput = New MSXML2.DOMDocument
        If Err <> 0 Then
            ShowMsgbox "���ܴ���XML����������ע��msxml3.dll������"
            Exit Function
        End If
        
        Dim strYBServer As String
        On Error Resume Next
        #If gblnTest = 1 Then
            Set gobj���� = CreateObject("GYSYB.CLSGYSYB")
            If Err <> 0 Then
                ShowMsgbox "���ص��Բ���ʱ����������Ϣ���£�" & vbCrLf & Err.Description
                Exit Function
            End If
            Set obj���� = gobj����
        #Else
            '�����ȫ�ֱ�������ʱ����ʱ��Ⱥܾã�������Դ�����ԭ��
            strYBServer = Get���ղ���_����("ҽ��������")
            If strYBServer = "" Then
                Set gobj���� = CreateObject("HospCOMSvr.HospCOMServer")
                Set obj���� = CreateObject("HospRecSvr.HospRecServer")
            Else
                Set gobj���� = CreateObject("HospCOMSvr.HospCOMServer", strYBServer)
                Set obj���� = CreateObject("HospRecSvr.HospRecServer", strYBServer)
            End If
            If Err <> 0 Then
                mblnFail = True
                ShowMsgbox "�޷�����ҽ���ӿڲ�����HospCOMSvr.HospCOMServer����"
                Exit Function
            End If
        #End If
    End If
    'ȡ���ղ���
    On Error GoTo errHand
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=" & TYPE_������
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ղ���", TYPE_������)
    Do While Not rsTemp.EOF
        If rsTemp!������ = "ҽ���û���" Then
            strUser = Nvl(rsTemp!����ֵ)
        ElseIf rsTemp!������ = "ҽ���û�����" Then
            strPass = Nvl(rsTemp!����ֵ)
        ElseIf rsTemp!������ = "ҽ��������1" Then
            strServer = Nvl(rsTemp!����ֵ)
        ElseIf rsTemp!������ = "��Ժ����" Then
            mblnҽ����Ժ = (Nvl(rsTemp!����ֵ, 0) = 0)
        ElseIf rsTemp!������ = "����ҩƷ��ʾ" Then
            mbln����ҩƷ��ʾ = (Nvl(rsTemp!����ֵ, 0) = 1)
        ElseIf rsTemp!������ = "������㴰�ڹر�ʱ��" Then
            mlngCloseTime = Nvl(rsTemp!����ֵ, 20)
        ElseIf rsTemp!������ = "�����������Լ�����ҩ���ƹ���" Then
            gbln������ҩ���� = (Val(Nvl(rsTemp!����ֵ, 0)) = 1)
        ElseIf rsTemp!������ = "�ۼ���ҩ�������׼" Then
            gint�ۼ���ҩ�������׼ = Nvl(rsTemp!����ֵ, 2)
        ElseIf rsTemp!������ = "�������" Then
            ging������� = Nvl(rsTemp!����ֵ, 0)
        End If
        rsTemp.MoveNext
    Loop
    If Not OraDataOpen(gcnGYYB, strServer, strUser, strPass, True) Then Exit Function
    
    'HIS�汾�� �̳ظ���10.26���ϼ����°汾���� 2011-04-07
    gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
    If Split(rsTemp!�汾��, ".")(0) = 10 And Split(rsTemp!�汾��, ".")(1) >= 29 Then
        '10.29���ϰ汾
        gblnHIS1029 = True
    ElseIf Split(rsTemp!�汾��, ".")(0) = 10 And Split(rsTemp!�汾��, ".")(1) >= 26 Then
        '10.26�����ϸ߰汾
        gblnHIS1026 = True
        gblnHIS1029 = False
    Else
        '10.26�����µͰ汾
        gblnHIS1026 = False
        gblnHIS1029 = False
    End If
         
    mblnInit = True
    ҽ����ʼ��_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim str������� As String, str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, cur�ʻ���� As Currency
    Dim STR���� As String, str�Ա� As String, str���֤���� As String, lng���� As Long
    Dim str�������� As String, str��Ա��� As String, str��λ���� As String, str��λ���� As String, strҽ���չ���Ա As String
    Dim strIdentify As String, str���� As String
    Dim bln������־ As Boolean, str֧������ As String
    Dim lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    
    '��ʼ��һЩ�������ڳ�����;�˳�ʱֵȴ�Ѿ�����
    gstr�����϶���� = ""
    If frmIdentify����.GetIdentify(bytType, str����, strҽ����, str�����ı��, str����, bln������־, lng����ID, str֧������) = False Then
        Exit Function
    End If
    If bytType = 0 Or bytType = 3 Then str֧������ = ""
    '��ԭ����
    str������� = Split(str����, "^")(1)
    str���� = Split(str����, "^")(0)
    
    If bytType = id����ȷ�� Then
        '�÷���ֵ��ʱû�����ã�ֻҪ��Ϊ�վͱ�ʾ�ɹ���
        ��ݱ�ʶ_���� = str���� & ";" & strҽ���� & ";" & str����
        Exit Function
    End If
    
    'ȡ�÷���ֵ
    STR���� = GetElemnetValue("PERSONNAME")
    str�Ա� = GetElemnetValue("SEX")
    str�Ա� = Switch(str�Ա� = "1", "��", str�Ա� = "2", "Ů", str�Ա� = "9", "����", True, str�Ա�)
    str���֤���� = GetElemnetValue("PID")
    
    str�������� = AddDate(GetElemnetValue("BIRTHDAY"))
    If IsDate(str��������) = True Then
        lng���� = DateDiff("yyyy", CDate(str��������), zlDatabase.Currentdate)
    Else
        str�������� = ""
    End If
    
    str��Ա��� = GetElemnetValue("PERSONTYPE")
    str��Ա��� = Switch(str��Ա��� = "11", "��ְ", str��Ա��� = "21", "����", _
                      str��Ա��� = "32", "ʡ������", str��Ա��� = "34", "��������", _
                      str��Ա��� = "41", "��ͨ����", str��Ա��� = "42", "�ͱ�����", _
                      str��Ա��� = "43", "������Ա", str��Ա��� = "44", "���ռ�ͥ", _
                      str��Ա��� = "45", "�ضȲм�", True, "����") '����������ʾ�������ͥ�������ݿ�ֻ��8λ��ֻ�ܱ���Ϊ���ռ�ͥ
    str��λ���� = ToVarchar(GetElemnetValue("DEPTCODE"), 12)
    str��λ���� = ToVarchar(GetElemnetValue("DEPTNAME"), 36) '�ֶγ��ȱ���50�������ڻ�Ҫ������뼰����
    cur�ʻ���� = Val(GetElemnetValue("ACCTBALANCE"))
    str������� = Val(GetElemnetValue("INSURETYPE"))
    strҽ���չ���Ա = Val(GetElemnetValue("CAREPSNFLAG"))
    
    '����;ҽ����;����;����;�Ա�;��������;���֤;������λ
    'ҽ���ŵ�һλΪ������
    '�ѷֺ��滻�ɶ���
    strIdentify = Replace(str����, ";", ",") & ";" & strҽ���� & ";;" & STR���� & ";" & str�Ա� & ";" & str�������� & ";" & str���֤���� & ";" & str��λ���� & "(" & str��λ���� & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    str���� = ";"                                       '8.���Ĵ���
    str���� = str���� & ";"                             '9.˳���  ����ҽ�����ڱ���ҽ�������ı��루���⽨��ҽ�����ģ�
    str���� = str���� & ";" & str��Ա���               '10��Ա���
    str���� = str���� & ";" & cur�ʻ����               '11�ʻ����
    str���� = str���� & ";0"                            '12��ǰ״̬
    str���� = str���� & ";" & IIf(lng����ID <> 0, lng����ID, "")   '13����ID
    str���� = str���� & ";" & IIf(str��Ա��� = "��ְ", 1, 2)      '14��ְ(1,2)
    str���� = str���� & ";"                             '15����֤��
    str���� = str���� & ";" & lng����                   '16�����
    str���� = str���� & ";"                             '17�Ҷȼ�
    str���� = str���� & ";" & cur�ʻ����               '18�ʻ������ۼ�
    str���� = str���� & ";0"                            '19�ʻ�֧���ۼ�
    str���� = str���� & ";"                             '20����ͳ���ۼ�
    str���� = str���� & ";"                             '21ͳ�ﱨ���ۼ�
    str���� = str���� & ";"                             '22סԺ�����ۼ�
    str���� = str���� & ";"                             '23�������� (1����������)
    
    lng����ID = BuildPatiInfo(bytType, strIdentify & str����, lng����ID, TYPE_������)
    '���ظ�ʽ:�м���벡��ID
    If lng����ID <> 0 Then
        ��ݱ�ʶ_���� = strIdentify & ";" & lng����ID & str����
        
        mstr���� = str����
        mstr���� = str����
        
        '���µ�ǰҽ�����˵ı�������Լ�ҽ���չ���Ա��־
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'�������','''" & str������� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���汣�����")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'ҽ���չ���Ա','''" & strҽ���չ���Ա & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ���չ���Ա")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'������־','''" & IIf(bln������־, 1, 0) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����������־")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'֧������','''" & str֧������ & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����֧������")
    
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'�����','''" & gintType & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���濨���")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'IDNO','''" & gstrIDNO & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������֤��")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'PSAM','''" & gstrPSAMNO & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����PSAM")
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��������_������(ByVal str�ſ����� As String, ByVal str���� As String, ByVal str������ As String) As Boolean
    If InitXML = False Then Exit Function
    
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' �����
    Call InsertChild(mdomInput.documentElement, "CARDDATA", str�ſ�����)            ' �ſ����ݣ�����IC���ţ�
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' ��ᱣ�Ϻţ����֤�ţ�
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP��ַ
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM��оƬ
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)                ' ����
    Call InsertChild(mdomInput.documentElement, "NEWPASSWORD", str������)           ' ����
    
    '���ýӿ�
    If CommServer("MODIFYCARD") = False Then Exit Function
    ��������_������ = True
End Function

Public Function �������_����(strSelfNo As String) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: strSelfNO-���˸��˱��
'����: ���ظ����ʻ����Ľ��
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    '�����ݿ��ж�ȡ����Ϊ�ղŲű����˵ģ�Ӧ����׼ȷ�ģ�
    gstrSQL = " Select /*+ rule */  B.�ʻ���� " & _
              " From ҽ�����˹����� A,ҽ�����˵��� B " & _
              " Where A.����=B.���� AND A.����=B.���� AND A.ҽ����=B.ҽ���� AND A.��־=1 " & _
              " And B.����=[1] and B.����=0 and B.ҽ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������, strSelfNo)
    
    If rsTemp.EOF = False Then
        �������_���� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, Optional strAdvance As String) As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim str��Ŀ���� As String
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String, str������� As String
    Dim dbl�����ʻ� As Double, dbl����Ա���� As Double, dblHIS�ܶ� As Double, dbl�����ܷ��� As Double, dblҽ������ As Double, dbl��ͳ�� As Double, dbl��� As Double
    Dim lng����ID As Long, str�������� As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Dim int����_CUR As Integer, int����_MAX As Integer, strҩƷ���� As String
    Dim str�շ�ϸĿIDS      As String   '�շ�ϸĿ����

    On Error GoTo errHandle
    
    '���ӵ������Ժ󣬷��ز������ӣ������ܷ��ã�������Ҫ���Ӹ�����ֶΣ�������֤HIS����������ֽ�֧����ҽ��һ�£���ʽ���£�
    '���=HIS�ܷ���-�����ܷ��ã��ֽ�֧��=HIS�ܷ���-ͳ��֧��-���
    
    If rs��ϸ.RecordCount = 0 Then
        str���㷽ʽ = "�����ʻ�;0;0"
        �����������_���� = True
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    
    '�жϸò����Ƿ�����������
    gstrSQL = "select A.��Ա���,A.�������,B.���� from �����ʻ� A,���ղ��� B where A.����ID=[1] and A.����=[2] and A.����ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", lng����ID, TYPE_������)
    If rsTemp.EOF = False Then
        str�������� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        str��Ա��� = Nvl(rsTemp("��Ա���"), "")
        str������� = Nvl(rsTemp!�������)
        'ת����Ա���
        str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21", _
                          str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", _
                          str��Ա��� = "��ͨ����", "41", str��Ա��� = "�ͱ�����", "42", _
                          str��Ա��� = "������Ա", "43", str��Ա��� = "���ռ�ͥ", "44", _
                          str��Ա��� = "�ضȲм�", "45", True, "11") '����������ʾ�������ͥ�������ݿ�ֻ��8λ��ֻ�ܱ���Ϊ���ռ�ͥ
    End If
    datCurr = zlDatabase.Currentdate
    '�ֽⵥ����������ǰ����
    If strAdvance = "" Then strAdvance = "1|1"
    int����_CUR = Val(Split(strAdvance, "|")(1))
    int����_MAX = Val(Split(strAdvance, "|")(0))
'    If (int����_MAX > 1) And �൥���շ�_�շѷֱ��ӡ Then
'        MsgBox "��ȡ��ϵͳ����������Ʊ��ҳ����Ĳ����������շ�ÿ�ŵ��ݷֱ��ӡ�������ɽ��ж൥���շѣ�", vbInformation, gstrSysName
'        Exit Function
'    End If
    mint���㷽ʽ = 0: mstr�����ֱ��� = ""
    If str�������� = "" Then
ReChoose:
        '��ͨ����Ҫ��ѡ����㷽ʽ�뵥���ֽ���Ŀ¼�����㷽ʽ;�����ֱ��룩
        If int����_CUR = 1 Then
            mstr�����ֱ��� = ���ý��㷽ʽ_����(lng����ID, Nothing, False)
            If mstr�����ֱ��� = "" Then mstr�����ֱ��� = ";"
            mint���㷽ʽ = Val(Split(mstr�����ֱ���, ";")(0))
            mstr�����ֱ��� = Split(mstr�����ֱ���, ";")(1)
        End If
    End If
    
    'If Get��֤_����(0, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    
    
    
    If int����_CUR = 1 Then
        dblTOTAL = 0
        preBalance.dbl������ = 0
        preBalance.dbl��ͳ�� = 0
        preBalance.dbl�����ʻ� = 0
        preBalance.dblҽ������ = 0
        preBalance.dblҽ�Ʋ��� = 0
        
        '��XML DomDocument������г�ʼ��
        If InitXML = False Then Exit Function
        Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' �����
        Call InsertChild(mdomInput.documentElement, "CARDDATA", mstr����)           ' �ſ�����
        Call InsertChild(mdomInput.documentElement, "INSURETYPE", str�������)     ' �������
        Call InsertChild(mdomInput.documentElement, "PASSWORD", mstr����)         ' ����
        Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstrҽ����)      ' ���˱��
        Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' ��ᱣ�Ϻ�
        Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP��ַ
        Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM��оƬ
        If str�������� <> "" Then '��������
            '����8λ����
            str�������� = String(8 - Len(str��������), "0") & str��������
            Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str��������)         '���ֲ�����
            Call InsertChild(mdomInput.documentElement, "CFNO", gstr��������)           '��������
        End If
        If str������� = "7" Then
            Call InsertChild(mdomInput.documentElement, "GSRDBH", gstr�����϶����)          '�����϶����
        End If
        Call InsertChild(mdomInput.documentElement, "ISCAL", 0)         ' �Ƿ����
        Call InsertChild(mdomInput.documentElement, "CALTYPE", mint���㷽ʽ)         ' ���㷽ʽ
        Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", mstr�����ֱ���)         ' �����ֽ������
        Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", "0")     ' �˻�֧����
        Call InsertChild(mdomInput.documentElement, "INVOICENO", "") ' ��Ʊ��
        Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
        Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) ' ��������
        Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) '������ʼ����ʱ��
        Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' ������ϸ
        Set mnodRowset = nodRowset
    Else
        Set nodRowset = mnodRowset
        Set mdomInput = tdomInput
    End If
    str�շ�ϸĿIDS = ""
    Do Until rs��ϸ.EOF
        If Nvl(rs��ϸ!ʵ�ս��, 0) <> 0 Then
            gstrSQL = "SELECT C.ҩƷID,C.���,E.���� AS ����  FROM ҩƷĿ¼ C,ҩƷ��Ϣ D,ҩƷ���� E WHERE C.ҩƷID=" & rs��ϸ("�շ�ϸĿID") & " AND C.ҩ��ID=D.ҩ��ID AND D.����=E.����"
            gstrSQL = "select A.���,A.����,B.��Ŀ����,nvl(A.���,F.���) AS ���,F.����,A.���㵥λ from �շ�ϸĿ A,����֧����Ŀ B,(" & gstrSQL & _
                    ") F where A.ID=[1] and A.ID=B.�շ�ϸĿID  AND A.Id=F.ҩƷID(+) and B.����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", CLng(rs��ϸ("�շ�ϸĿID")), TYPE_������)
            If rsTemp.EOF = True Then
                MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
             'ҩƷʹ��������� 2011-05-21  �̳ظ�
            If gbln������ҩ���� = True Then
              strҩƷ���� = strҩƷ���� & DrugsUsed(TYPE_������, lng����ID, Nvl(rs��ϸ!�շ�ϸĿID), Nvl(rs��ϸ!����, 0))
            End If
        

            Set nodRow = InsertChild(nodRowset, "ROW", "")
            
            On Error Resume Next
            str��Ŀ���� = ""
            str��Ŀ���� = Nvl(rs��ϸ!���ձ���)
            Err = 0
            On Error GoTo errHandle
            Call ��������ҩƷ��ʾ(rs��ϸ!�շ�ϸĿID)
                    '���շ�ϸĿID�ж��Ƿ��ں�����
        
            str�շ�ϸĿIDS = str�շ�ϸĿIDS & "," & Nvl(rs��ϸ!�շ�ϸĿID)

            '������ձ���Ϊ�գ�����Ҫ�û�ѡ�����
            If str��Ŀ���� = "" Then
                str��Ŀ���� = GetItemInsure_����(lng����ID, rs��ϸ!�շ�ϸĿID, True)
            End If
            If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsTemp!��Ŀ����)
            
            Call nodRow.setAttribute("ITEMCODE", ToVarchar(str��Ŀ����, 12))
            Call nodRow.setAttribute("ITEMNAME", ToVarchar(rsTemp("����"), 72))
            Call nodRow.setAttribute("SUBJECT", Subject(rsTemp("���")))
            Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rsTemp("���"), 40))
            Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rsTemp("����"), 20))
            Call nodRow.setAttribute("UNIT", ToVarchar(rsTemp("���㵥λ"), 20))
            Call nodRow.setAttribute("PRICE", Format(rs��ϸ("����"), "0.0000"))
            Call nodRow.setAttribute("QUANTITY", Format(rs��ϸ("����"), "0.00"))
            Call nodRow.setAttribute("FROMOFFICE", ToVarchar(UserInfo.����, 56)) '��������
            Call nodRow.setAttribute("FROMDOCT", Format(UserInfo.����, 20))      '����ҽ��
            Call nodRow.setAttribute("TOOFFICE", ToVarchar(UserInfo.����, 56))  '�ܵ�����
            Call nodRow.setAttribute("TODOCT", Format(UserInfo.����, 20))       '�ܵ�ҽ��
            Call nodRow.setAttribute("DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))        '��������
            Call nodRow.setAttribute("NOTE", ToVarchar(rs��ϸ("ժҪ"), 512))        '��ע
            
            dblTOTAL = dblTOTAL + Round(rs��ϸ!ʵ�ս��, 2)
        End If
        rs��ϸ.MoveNext
    Loop
    If Len(Replace(strҩƷ����, vbNewLine, "")) <> 0 Then ShowMsgbox "��" & int����_CUR & "���շѵ����У�" & vbNewLine & strҩƷ����: Exit Function
    
    gstrSQL = "select distinct ҽ���� FROM  �����ʻ�  WHERE ����=[1] AND ����ID=[2]"
    strҽ���� = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��������", TYPE_������, lng����ID).Fields(0)
    '����շ�ϸĿIDS�Ƿ�����ں�������
    str�շ�ϸĿIDS = Mid(str�շ�ϸĿIDS, 2)
    '��⵱ǰҽ�����Ƿ��ƺ�����������
    gstrSQL = "select Count(1) from ҽ��������_���� where ҽ���� = [1] and ״̬ = 1"
    If Val(zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��������", strҽ����).Fields(0)) = 1 Then
        '����ҽ����״̬�¼���Ƿ�����Ŀ���ں�������
        gstrSQL = "Select A.ҽ����, A.�շ�ϸĿid, C.����, C.����, C.���" & vbCrLf & _
                 "From ҽ����������Ŀ_���� A, Table(F_Str2list2([2])) B, �շ�ϸĿ C" & vbCrLf & _
                 "Where rownum = 1 And a.�շ�ϸĿID = B.C2 And a.�շ�ϸĿID = C.ID And a.ҽ���� = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ����������Ŀ", strҽ����, str�շ�ϸĿIDS)
        '����ǿ����ֻ������������͹��˱��ղ���
        If Not (rsTemp.EOF Or rsTemp.BOF) And (gstr�����־ = "��������" Or gstr���˱�־ = "���˱���") Then
            '���ں�������Ŀ
            MsgBox "��ǰҽ���š�" & strҽ���� & "��" & vbCrLf & _
                   "�շ�ϸĿID��" & rsTemp!�շ�ϸĿID & "��" & vbCrLf & _
                   "�շ�ϸĿ���롾" & rsTemp!���� & "��" & vbCrLf & _
                   "�շ�ϸĿ���ơ�" & rsTemp!���� & "��" & vbCrLf & _
                   "�շ�ϸĿ���" & rsTemp!��� & "��" & vbCrLf & _
                   "������ҽ����������Ŀ�У���ҽ�������Ա��ȡ������ã�", vbCritical, gstrSysName
            Exit Function
        End If
    End If
    
    '���ýӿ�
    Set tdomInput = mdomInput
    If CommServer(IIf(str�������� <> "", "CALSPECCLIN", "CALCLIN")) = False Then Exit Function
    
    '��ͬ����Ⱥ�����ص�XML���ֶα�ʾ���岻ͬ������ֱ��ȡ������Ҫ�ֱ��ж�
    '������Ա��������ͨ�������������ͳһ��ALLOWFUND֧����
    '����ҽ����Ա����������FUND1PAY��FUND2PAY֧������ͨ�����ɸ����ʻ�֧��
    If str��Ա��� = "32" Or str��Ա��� = "34" Then
        dblҽ������ = Val(GetElemnetValue("ALLOWFUND"))
        dblҽ������ = dblҽ������ - preBalance.dblҽ������
        preBalance.dblҽ������ = Val(GetElemnetValue("ALLOWFUND"))
        str���㷽ʽ = "ҽ������;" & dblҽ������ & ";0"
    Else
        dbl�����ʻ� = Val(GetElemnetValue("ACCTPAY"))
        dbl�����ʻ� = dbl�����ʻ� - preBalance.dbl�����ʻ�
         preBalance.dbl�����ʻ� = Val(GetElemnetValue("ACCTPAY"))
        str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";1"   '�����޸ĸ����ʻ�
        '��ǰ����������ŷ��أ����ڸ�Ϊֱ��ȡ����ֵ�ͱ�������Ϊ�����˹��� Modified by ZYB 20080702
        
'        If str�������� <> "" Then
            
            dblҽ������ = Val(GetElemnetValue("FUND1PAY"))
            dblҽ������ = dblҽ������ - preBalance.dblҽ������
            preBalance.dblҽ������ = Val(GetElemnetValue("FUND1PAY"))
            
            dbl��ͳ�� = Val(GetElemnetValue("FUND2PAY"))
            dbl��ͳ�� = dbl��ͳ�� - preBalance.dbl��ͳ��
             preBalance.dbl��ͳ�� = Val(GetElemnetValue("FUND2PAY"))
            
            str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & dblҽ������ & ";0" & _
                         "|��ͳ��;" & dbl��ͳ�� & ";0"
'        End If
    End If
    dbl����Ա���� = Val(GetElemnetValue("FUND3PAY"))
    dbl����Ա���� = dbl����Ա���� - preBalance.dblҽ�Ʋ���
    preBalance.dblҽ�Ʋ��� = Val(GetElemnetValue("FUND3PAY"))
    str���㷽ʽ = str���㷽ʽ & "|ҽ�Ʋ���;" & dbl����Ա���� & ";0"
    
    dbl�����ܷ��� = Val(Format(GetElemnetValue("CALFEEALL"), "#0.00;-#0.00;0;"))
    dblHIS�ܶ� = Val(Format(GetElemnetValue("HOSPFEEALL"), "#0.00;-#0.00;0;"))
    
    '�ȱȽ��ܶ��Ƿ�һ��
    If Format(dblTOTAL, "#0.00") <> Format(dblHIS�ܶ�, "#0.00") Then
        If Abs(Val(Format(dblTOTAL, "#0.00")) - Val(Format(dblHIS�ܶ�, "#0.00"))) > 0.5 Then
            MsgBox "HIS�ܶ���ҽ�����յ����ܷ��ò�һ�£���������㣡" & vbCrLf & _
                "HIS:" & Format(dblTOTAL, "#0.00") & Space(10) & "ҽ��:" & Format(dblHIS�ܶ�, "#0.00"), vbInformation, gstrSysName
            Exit Function
        Else
            If MsgBox("HIS�ܶ���ҽ�����յ����ܷ��ò�һ�£������ǵ��۾��Ȳ�����������Ƿ���㣿" & vbCrLf & _
                "HIS:" & Format(dblTOTAL, "#0.00") & Space(10) & "ҽ��:" & Format(dblHIS�ܶ�, "#0.00"), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    dbl��� = dblHIS�ܶ� - dbl�����ܷ���
    dbl��� = dbl��� - preBalance.dbl������
    preBalance.dbl������ = dblHIS�ܶ� - dbl�����ܷ���
    If dbl��� <> 0 And mint���㷽ʽ = 1 Then
        '���=HIS�ܷ���-�����ܷ��ã��ֽ�֧��=HIS�ܷ���-ͳ��֧��-���
        str���㷽ʽ = str���㷽ʽ & "|������;" & dbl��� & ";0"
    End If
    �����������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, Optional ByVal bln�Һ� As Boolean = False, Optional ByRef strAdvance As String = "") As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim str��Ŀ���� As String
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String, str������� As String, str������㷽ʽ As String
    Dim strҽ�� As String, str���� As String, cur�������� As Double, curҽ���ܷ��� As Double, datCurr As Date
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim lng����ID  As Long, str��������   As String, lng��Ŀ�� As Long, cur�ʻ���� As Currency
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str���㷽ʽ     As String
    Dim intCurr As Integer, intMAX As Integer, dbl�ֽ� As Double
    Dim blnOld As Boolean               '�Ƿ����ϰ�HISϵͳ
    
    '---------------------------------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
            
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, cur���� As Double
    Dim dblHIS�ܶ� As Double, dbl�����ܷ��� As Double, dbl��� As Double
    Dim cur����Ա�����𸶱�׼ As Double, cur����Ա�������� As Double, cur��ͨ���﹫��Ա�����ۼ� As Double
    Dim cur����Ա���� As Double, cur������޶��Ա���� As Double
    Dim str����˳��� As String, str������ As String, str�������˵�� As String
    
    On Error GoTo errHandle
    
    cur�ʻ���� = �������_����(strSelfNo)
    lng��Ŀ�� = Val(Get���ղ���_����("���������Ŀ��"))
    
     '�൥�ݴ��� 2011-4-7 �̳ظ�
    #If gverControl < 2 Then
        blnOld = True
    #End If
    strAdvance = Decode(Trim(strAdvance), "", "1|1", strAdvance)
    intCurr = Split(strAdvance, "|")(1) '��ǰ����
    intMAX = Split(strAdvance, "|")(0) '��󵥾�
    If intCurr = 1 Then
        g��������.dbl�󲡻��� = 0
        g��������.dbl�����ܶ� = 0
        g��������.dbl�����ʻ� = 0
        g��������.dbl����Ա���� = 0
        g��������.dbl�ֽ� = 0
        g��������.dblҽ������ = 0
        g��������.dbl������ = 0
    End If
    
    gstrSQL = "SELECT Nvl(��������,Nvl(�۸񸸺�,���)) AS ����� FROM " & IIf(gblnHIS1026 = True, "������ü�¼", "���˷��ü�¼") & _
             " WHERE ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0" & _
             " GROUP BY Nvl(��������,Nvl(�۸񸸺�,���))"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    If rs��ϸ.RecordCount > lng��Ŀ�� Then
        Err.Raise 9000, gstrSysName, "�����շѵ���Ŀ�����ܳ���" & lng��Ŀ�� & "��"
        Exit Function
    End If
    
    gstrSQL = "Select A.ID,A.���,A.�շ�ϸĿID,A.��¼����,A.��¼״̬,A.����ID,A.NO,A.�Ǽ�ʱ��,A.������ as ҽ��,A.�Ǽ�ʱ��," & _
            "   A.����*A.���� as ����,A.��׼���� as ʵ�ʼ۸�,A.���ʽ��," & _
            "   A.�շ����,A.���ձ���,D.��Ŀ����,B.���� as ��Ŀ����,C.���� as ��������,nvl(B.���,F.���) AS ���,F.����,B.���㵥λ,A.ժҪ " & _
            " From (Select * From " & IIf(gblnHIS1026 = True, "������ü�¼", "���˷��ü�¼") & " Where ����ID=[1] And Nvl(ʵ�ս��,0)<>0 And Nvl(���ӱ�־,0)<>9) A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D " & _
            "     ,(SELECT C.ҩƷID,C.���,E.���� AS ����  FROM " & IIf(gblnHIS1026 = True, "������ü�¼", "���˷��ü�¼") & " A,ҩƷĿ¼ C,ҩƷ��Ϣ D,ҩƷ���� E WHERE A.����ID=[1] AND A.�շ�ϸĿID=C.ҩƷID AND C.ҩ��ID=D.ҩ��ID AND D.����=E.����) F " & _
            " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID  AND A.ID=F.ҩƷID(+) And D.����=[2] And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.��¼״̬,0)<>0" & _
            " Order by A.ID"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID, TYPE_������)
    If rs��ϸ.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    strҽ�� = ToVarchar(IIf(IsNull(rs��ϸ("ҽ��")), UserInfo.����, rs��ϸ("ҽ��")), 20)
    str���� = ToVarchar(IIf(IsNull(rs��ϸ("��������")), UserInfo.����, rs��ϸ("��������")), 56)
    datCurr = zlDatabase.Currentdate
    
    'һ��������ϸ����
    
    '�жϸò����Ƿ�����������
    gstrSQL = "select A.��Ա���,A.�������,B.���� from �����ʻ� A,���ղ��� B where A.����ID=[1] and A.����=[2] and A.����ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", lng����ID, TYPE_������)
    If rsTemp.EOF = False Then
        str�������� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        str��Ա��� = Nvl(rsTemp("��Ա���"), "")
        str������� = Nvl(rsTemp!�������)
        'ת����Ա���
        str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21", _
                          str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", _
                          str��Ա��� = "��ͨ����", "41", str��Ա��� = "�ͱ�����", "42", _
                          str��Ա��� = "������Ա", "43", str��Ա��� = "���ռ�ͥ", "44", _
                          str��Ա��� = "�ضȲм�", "45", True, "11") '����������ʾ�������ͥ�������ݿ�ֻ��8λ��ֻ�ܱ���Ϊ���ռ�ͥ
    End If
    
    'If Get��֤_����(0, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
        
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' �����
    Call InsertChild(mdomInput.documentElement, "CARDDATA", mstr����)           ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", str�������)   ' �������
    Call InsertChild(mdomInput.documentElement, "PASSWORD", mstr����)         ' ����
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstrҽ����)      ' ���˱��
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' ��ᱣ�Ϻ�
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP��ַ
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM��оƬ
    If str�������� <> "" Then '��������
        '����8λ����
        str�������� = String(8 - Len(str��������), "0") & str��������
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str��������)         '���ֲ�����
        Call InsertChild(mdomInput.documentElement, "CFNO", gstr��������)          '��������
    End If
    If str������� = "7" Then
        Call InsertChild(mdomInput.documentElement, "GSRDBH", gstr�����϶����)          '�����϶����
    End If
    Call InsertChild(mdomInput.documentElement, "ISCAL", 1)               ' �Ƿ����
    Call InsertChild(mdomInput.documentElement, "CALTYPE", mint���㷽ʽ)         ' ���㷽ʽ
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", mstr�����ֱ���)         ' �����ֽ������
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", Format(cur�����ʻ�, "0.00")) ' �˻�֧����
    Call InsertChild(mdomInput.documentElement, "INVOICENO", "M" & "_" & rs��ϸ!��¼���� & "_" & rs��ϸ("NO")) ' ��Ʊ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")) ' ��������
    Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")) '������ʼ����ʱ��
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' ������ϸ
    
    dblHIS�ܶ� = 0
    Do Until rs��ϸ.EOF
        cur�������� = cur�������� + rs��ϸ("���ʽ��")

            
        Set nodRow = InsertChild(nodRowset, "ROW", "")
        
        str��Ŀ���� = Nvl(rs��ϸ!���ձ���)
        
        '������ձ���Ϊ�գ�����Ҫ�û�ѡ�����
        If str��Ŀ���� = "" Then
            str��Ŀ���� = GetItemInsure_����(lng����ID, rs��ϸ!�շ�ϸĿID, False)
        End If
        If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rs��ϸ!��Ŀ����)
        
        Call nodRow.setAttribute("ITEMCODE", ToVarchar(str��Ŀ����, 12))
        Call nodRow.setAttribute("ITEMNAME", ToVarchar(rs��ϸ("��Ŀ����"), 72))
        Call nodRow.setAttribute("SUBJECT", Subject(rs��ϸ("�շ����")))
        Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rs��ϸ("���"), 40))
        Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rs��ϸ("����"), 20))
        Call nodRow.setAttribute("UNIT", ToVarchar(rs��ϸ("���㵥λ"), 20))
        Call nodRow.setAttribute("PRICE", Format(rs��ϸ("ʵ�ʼ۸�"), "0.0000"))
        Call nodRow.setAttribute("QUANTITY", Format(rs��ϸ("����"), "0.00"))
        
        dblHIS�ܶ� = dblHIS�ܶ� + Format(rs��ϸ("ʵ�ʼ۸�") * rs��ϸ("����"), "0.0000")
        
        Call nodRow.setAttribute("FROMOFFICE", str����)    '��������
        Call nodRow.setAttribute("FROMDOCT", strҽ��)      '����ҽ��
        Call nodRow.setAttribute("TOOFFICE", str����)     '�ܵ�����
        Call nodRow.setAttribute("TODOCT", strҽ��)       '�ܵ�ҽ��
        
        '����ʱ��ʱ��Ϊ�˱�֤ͬһ������Ŀ�ĵ��շ�ʱ�䲻ͬ������ڵǼ�ʱ���ϰ���ż�������
        Call nodRow.setAttribute("DODATE", Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss"))    '��������
        Call nodRow.setAttribute("NOTE", ToVarchar(rs��ϸ("ժҪ"), 512))         '��ע
        
        rs��ϸ.MoveNext
    Loop
     
    '��������¼
    '---------------------------------------------------------------------------------------------
    
    If g�������.blnYn And g�������.str���� = str���� And g�������.strҽ���� = strҽ���� And g�������.lng����ID = lng����ID Then
        g�������.blnYn = False
        curȫ�Ը� = g�������.m_ȫ�Ը�
        cur�ҹ��Ը� = g�������.m_�ҹ��Ը�
        cur���� = g�������.m_����
        cur�����Ը� = g�������.m_�����Ը�
        
        If str��Ա��� = "32" Or str��Ա��� = "34" Then
            curͳ��֧�� = g�������.m_ͳ��֧��
            cur��ͳ�� = 0
        Else
            curͳ��֧�� = g�������.m_ͳ��֧��
            cur��ͳ�� = g�������.m_��ͳ��
        End If
        curͳ���Ը� = g�������.m_ͳ���Ը�
        cur���Ը� = g�������.m_���Ը�
        cur�����Ը� = g�������.m_�����Ը�
        cur����Ա���� = g�������.m_����Ա����
        
        cur����Ա�����𸶱�׼ = g�������.m_����Ա�����𸶱�׼
        cur����Ա�������� = g�������.m_����Ա��������
        cur��ͨ���﹫��Ա�����ۼ� = g�������.m_��ͨ���﹫��Ա�����ۼ�
        cur������޶��Ա���� = g�������.m_������޶��Ա����
        curҽ���ܷ��� = g�������.m_ҽ���ܷ���
        dbl�����ܷ��� = g�������.m_�����ܷ���
        
        If str�������� = "" Then
            dblHIS�ܶ� = g�������.m_HIS�ܷ���
            dbl��� = dblHIS�ܶ� - dbl�����ܷ���
        End If
        
        str������㷽ʽ = g�������.m_������㷽ʽ
        str�������˵�� = g�������.m_�������˵��
        str������ = g�������.m_������
        str����˳��� = g�������.m_����˳���
    Else
        '���ýӿ�
        If CommServer(IIf(str�������� <> "", "CALSPECCLIN", "CALCLIN")) = False Then Exit Function
    
            
        curȫ�Ը� = Val(GetElemnetValue("FEEOUT"))
        cur�ҹ��Ը� = Val(GetElemnetValue("FEESELF"))
        cur���� = Val(GetElemnetValue("STARTFEE"))
        cur�����Ը� = Val(GetElemnetValue("ENTERSTARTFEE"))
        
        If str��Ա��� = "32" Or str��Ա��� = "34" Then
            curͳ��֧�� = Val(GetElemnetValue("ALLOWFUND"))
            cur��ͳ�� = 0
        Else
            curͳ��֧�� = Val(GetElemnetValue("FUND1PAY"))
            cur��ͳ�� = Val(GetElemnetValue("FUND2PAY"))
        End If
        curͳ���Ը� = Val(GetElemnetValue("FUND1SELF"))
        cur���Ը� = Val(GetElemnetValue("FUND2SELF"))
        cur�����Ը� = Val(GetElemnetValue("FEEOVER"))
        
        cur����Ա�����𸶱�׼ = Val(GetElemnetValue("STARTFEE2STD"))
        cur����Ա�������� = cur����
        cur��ͨ���﹫��Ա�����ۼ� = Val(GetElemnetValue("ENTERLMT3"))
        cur����Ա���� = Val(GetElemnetValue("FUND3PAY"))
        cur������޶��Ա���� = Val(GetElemnetValue("FUND3OVER"))
        curҽ���ܷ��� = Val(GetElemnetValue("FEEALL"))
        dbl�����ܷ��� = Val(GetElemnetValue("CALFEEALL"))
        
        If str�������� = "" Then
            dblHIS�ܶ� = Val(GetElemnetValue("HOSPFEEALL"))
            dbl��� = dblHIS�ܶ� - dbl�����ܷ���
        End If
        
        str������㷽ʽ = GetElemnetValue("SPECCALFLAG")
        str�������˵�� = GetElemnetValue("SPECCALFLAGTXT")
        str������ = GetElemnetValue("BALANCEID")
        str����˳��� = GetElemnetValue("BILLNO")
        
    End If
    
    Call SaveBalanceLog(1, lng����ID, lng����ID, GetElemnetValue("BILLNO"), str������, IIf(str�������� <> "", "18", "11"))
    
    If str�������� <> "" Then
        str����˳��� = "����" & str����˳��� '�Ѽ������������˳�������һ��
    Else
        str����˳��� = "��ͨ" & str����˳���         '��ʾ��ͨ����
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� + curͳ���Ը� + cur�����Ը� + cur�����Ը� + cur��ͳ�� + cur���Ը� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� + cur��ͳ�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ���� & "," & cur�ʻ���� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & ",0," & cur�����Ը� & "," & cur�������� & "," & _
        curȫ�Ը� & "," & cur�ҹ��Ը� & "," & curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & _
        cur�����ʻ� & ",'" & str������ & "',null,null,'" & str����˳��� & "',0,'" & AnalyseComputer & "','" & gstrVersion & "','" & IIf(str�������� <> "", "18", "11") & "','" & Mid(str����˳���, 3) & "'," & _
            "NULL,'" & str�������� & "','" & str������� & "',to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ﲻ�����������㷽ʽ���������㷽ʽ���㣨��ȷ��ȡֵ��Χ�Ǵ�1��ʼ��
      '20110812����ǿ���ӹ���Ϣ����gstr�����϶���Ÿ�Ϊgstr������Ϣ
    gstrSQL = "zl_���㸽����Ϣ_Insert (" & lng����ID & "," & cur����Ա�����𸶱�׼ & "," & cur����Ա�������� & "," & cur��ͨ���﹫��Ա�����ۼ� & "," & cur����Ա���� & "," & cur������޶��Ա���� & ",0,0," & _
        "'" & mstr�����ֱ��� & "'," & mint���㷽ʽ & ",NULL,0," & dbl�����ܷ��� & "," & curҽ���ܷ��� & ",'" & gstr������Ϣ & "',0,'" & gstr�������� & "','" & str������㷽ʽ & "','" & str�������˵�� & "')"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    '---------------------------------------------------------------------------------------------
    
    '��̨ǿ�н���ҽ��У�����̳ظ� 2011-04-05 Begin
    If str��Ա��� = "32" Or str��Ա��� = "34" Then
        
        str���㷽ʽ = "ҽ������|" & Format(curͳ��֧��, "0.00")
    Else
        str���㷽ʽ = "�����ʻ�|" & Format(cur�����ʻ�, "0.00") & _
                    "||ҽ������|" & Format(curͳ��֧��, "0.00") & _
                    "||��ͳ��|" & Format(cur��ͳ��, "0.00")
    End If
    str���㷽ʽ = str���㷽ʽ & "||ҽ�Ʋ���|" & Format(cur����Ա����, "0.00")
    If mint���㷽ʽ = 1 And dbl��� > 0 Then
        str���㷽ʽ = str���㷽ʽ & "||������|" & Format(dbl���, "0.00")
    End If
    gstrSQL = "select sum(ʵ�ս��) from " & IIf(gblnHIS1026 = True, "������ü�¼", "���˷��ü�¼") & " where ����ID=[1]"
    dblHIS�ܶ� = zlDatabase.OpenSQLRecord(gstrSQL, "����HIS�ܷ���", lng����ID).Fields(0)
    dbl�ֽ� = dblHIS�ܶ� - curͳ��֧�� - cur��ͳ�� - cur����Ա���� - cur�����ʻ�
    If gblnHIS1029 Then
        If Not bln�Һ� Then
            strAdvance = str���㷽ʽ
        End If
    Else
        If Not bln�Һ� Then
            '���ﲿ���ɽӿ��ڲ�ֱ��У���������ɷ���ϵͳУ��
            'У�Բ��践���ֽ�
            'str���㷽ʽ = str���㷽ʽ & "||�ֽ�|" & dbl�ֽ�
            'gstrSQL = "zl_�����շѽ���_Update(" & lng����ID & ",'" & "" & "'," & 0 & ",'" & str���㷽ʽ & "'," & 0 & ")"
            gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
            gcnOracle.Execute gstrSQL
        Else
            '�ҺŲ����ɹҺų���У�����ӿڲ��ô���
        End If
        
        '�������ۼ�
        g��������.dbl�󲡻��� = g��������.dbl�󲡻��� + cur��ͳ��
        g��������.dbl�����ܶ� = g��������.dbl�����ܶ� + dblHIS�ܶ�
        g��������.dbl�����ʻ� = g��������.dbl�����ʻ� + cur�����ʻ�
        g��������.dbl����Ա���� = g��������.dbl����Ա���� + cur����Ա����
        g��������.dbl�ֽ� = g��������.dbl�ֽ� + dbl�ֽ�
        g��������.dblҽ������ = g��������.dblҽ������ + curͳ��֧��
        
        If mint���㷽ʽ = 1 And dbl��� > 0 Then '�����ֲ��ż��Ա�����
            g��������.dbl������ = g��������.dbl������ + dbl���
        Else '������ڲ����ǵ����֣��򽫲���ۼƵ��ֽ𲿷֡�
            g��������.dbl�ֽ� = g��������.dbl�ֽ� + dbl���
        End If
        
        '���һ�ŵ�����ʾ���յ��Ҳ����' And intMax > 1
        If intMAX = intCurr And bln�Һ� = False Then
            Call frm�����������.ShowForm(intMAX)
        End If
        strAdvance = "" '�ӿ��ڲ��Ѿ�У����������ҪHISУ��
    End If
    
    �������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function


Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    'Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Dim str������ As String, str����˳��� As String, curDate As Date, rsTemp As New ADODB.Recordset
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, intסԺ�����ۼ� As Integer
    Dim lng����ID As Long
    Dim bln���� As Boolean
    Dim str֧������ As String
    
    On Error GoTo errHandle
    
    '�˷�
    '�ж��Ƿ��н��ʼ�¼�������˵����סԺ����ʵ�ֵ�
    gstrSQL = "Select 1 from ���˽��ʼ�¼ where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��н��ʼ�¼�������˵����סԺ����ʵ�ֵ�", lng����ID)
    If rsTemp.RecordCount = 0 Then
        gstrSQL = "select distinct A.����ID from " & IIf(gblnHIS1026 = True, "������ü�¼", "���˷��ü�¼") & " A," & _
             IIf(gblnHIS1026 = True, "������ü�¼", "���˷��ü�¼") & " B " & _
                  " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Else
        gstrSQL = "select distinct A.ID AS ����ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
            " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����˷�", lng����ID)
    lng����ID = rsTemp("����ID")
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����˷�", TYPE_������, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    lng����ID = rsTemp!����ID
    cur�����ʻ� = Nvl(rsTemp!�����ʻ�֧��, 0)
    str������ = Nvl(rsTemp("֧��˳���"), "")
    str����˳��� = Nvl(rsTemp("��ע"), "")
    If str����˳��� = "" Then
        Err.Raise 9000, gstrSysName, "�õ���û�б������˳��ţ��������ϡ�"
        Exit Function
    End If
'    If Left(str����˳���, 2) = "����" Then
'        MsgBox "Ŀǰ��֧��������������ϡ�", vbInformation, gstrSysName
'        Exit Function
'    End If
    str֧������ = Nvl(rsTemp!ҽ�����)
    If str֧������ = "" Then str֧������ = IIf(Left(str����˳���, 2) = "����", "18", "11")
    str����˳��� = Mid(str����˳���, 3)
    curDate = zlDatabase.Currentdate
    
    'If Get��֤_����(0, str����, strҽ����, str�����ı��, str����, lng����ID, True) = False Then Exit Function
    Call �൥���շ�_�˷�(lng����ID)
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", str����˳���)     ' ����˳���
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str������)    ' ������
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str֧������)   ' ֧�����
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)    ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(curDate, "yyyy-MM-dd HH:mm:ss"))  ' ��������
    
    '���ýӿ�
    bln���� = IS����(lng����ID)
    If MsgBox("�Ƿ���Ʊ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    If CommServer("RETBALANCE", IIf(bln����, 1, 0)) = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") & "," & _
        cur�����ʻ� * -1 & ",'" & str������ & "',null,null,'" & Nvl(rsTemp("��ע")) & "'," & _
        "0,'" & AnalyseComputer & "','" & gstrVersion & "','" & str֧������ & "','" & Nvl(rsTemp!������ˮ��) & "'," & _
        "NULL,'" & Nvl(rsTemp!��������) & "','" & Nvl(rsTemp!����֢) & "',to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "Select * From ���㸽����Ϣ Where ����ID=" & lng����ID
    Call OpenRecordset_OtherBase(rsTemp, "��ȡ���㸽�Ӽ�¼", gstrSQL, gcnGYYB)
    If rsTemp.RecordCount <> 0 Then
        gstrSQL = "zl_���㸽����Ϣ_Insert (" & lng����ID & "," & -1 * Nvl(rsTemp!����Ա�����𸶱�׼, 0) & "," & -1 * Nvl(rsTemp!����Ա��������, 0) & "," & -1 * Nvl(rsTemp!��ͨ���﹫��Ա�����ۼ�, 0) & "," _
            & -1 * Nvl(rsTemp!����Ա����, 0) & "," & -1 * Nvl(rsTemp!������Ա����, 0) & ",0,0,'" & Nvl(rsTemp!�����ֱ���_����) & "'," & Nvl(rsTemp!���㷽ʽ, 0) & ",'" & Nvl(rsTemp!������) & "'," & _
            Nvl(rsTemp!���㷽ʽ, 0) & "," & -1 * Nvl(rsTemp!�����ܷ���, 0) & "," & -1 * Nvl(rsTemp!ҽ���ܷ���, 0) & ",'" & Nvl(rsTemp!�����϶����) & "',0,'" & Nvl(rsTemp!��������) & "')"
        gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    End If
    
    ����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function �����ʻ�תԤ��_����(lngԤ��ID As Long, cur�����ʻ� As Currency, strSelfNo As String, str˳��� As String, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    
    �����ʻ�תԤ��_���� = False
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim str���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str�������� As String
    Dim strTemp As String, str��ʾ As String, str��� As String, lng�α�ǰ��Ժ As Long
    Dim str֧������ As String, dbl�ʻ���� As Double
    Dim str������㷽ʽ As String, str���㷽ʽ As String, str���㲡�ֱ��� As String, str���㲡������ As String, str������� As String, str�������˵�� As String
    On Error GoTo errHandle
    
    'If Get��֤_����(1, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    
    'Modified by ZYB 20080703 �ӿ��ĵ�ȡ��������
'    '�жϸò����Ƿ�α�ǰ��Ժ
'    lng�α�ǰ��Ժ = 0
'    If Get���ղ���_����("��Ժʱѡ��α�ǰ��Ժ") = "1" Then
'        If MsgBox("�ò��˲α�ǰ�Ƿ��Ѿ���Ժ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
'            lng�α�ǰ��Ժ = 1
'        End If
'    End If
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.������Ϣ from ������ A where A.����ID=[1] and A.��ҳID=[2]" & _
              " and A.�������=1 and A.��ϴ���=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    If rsTemp.EOF = False Then
        str��� = ToVarchar(IIf(IsNull(rsTemp("������Ϣ")), "����", rsTemp("������Ϣ")), 128)
    Else
        gstrSQL = "select A.������Ϣ from ������ A where A.����ID=[1] and A.��ҳID=[2]" & _
                  " and A.�������=2 and A.��ϴ���=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
        If rsTemp.RecordCount <> 0 Then
            str��� = ToVarchar(Nvl(rsTemp!������Ϣ, "����"), 128)
        Else
            str��� = "����"   '��ϲ�����β���Ϊ��
        End If
    End If
    
    '���������Ժ��Ϣ
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.��Ժ��ʽ,nvl(A.����Ժת��,0) as ����Ժת��,A.����ҽʦ,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ����," & _
              "     C.סԺ��,D.�������,D.������־,D.�ʻ����,D.֧������ " & _
              " from ������ҳ A,���ű� B,������Ϣ C,�����ʻ� D " & _
              " Where A.����ID=D.����ID And D.����=[1] And A.����ID=C.����ID and A.��Ժ����ID = B.ID And A.����ID =[2] And A.��ҳID = [3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", TYPE_������, lng����ID, lng��ҳID)
    dbl�ʻ���� = Nvl(rsTemp!�ʻ����, 0)
    str������� = Nvl(rsTemp!�������)
    str֧������ = Nvl(rsTemp!֧������, "31")
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' �����
    Call InsertChild(mdomInput.documentElement, "CARDDATA", mstr����)           ' �ſ�����
    Call InsertChild(mdomInput.documentElement, "PASSWORD", mstr����)         ' ����
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstrҽ����)      ' ���˱��
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                      ' ��ᱣ�Ϻ�
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP��ַ
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM��оƬ
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", str�������)   ' �������
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str֧������)
    Call InsertChild(mdomInput.documentElement, "HOSPNO", ToVarchar(rsTemp("סԺ��"), 20))     ' סԺ��
'    Call InsertChild(mdomInput.documentElement, "ISINHOSP", lng�α�ǰ��Ժ)     ' �α�ǰ����Ժ 1���ǣ�0����
    Call InsertChild(mdomInput.documentElement, "DIAGNOSES", str���) ' ���
    Call InsertChild(mdomInput.documentElement, "DOCTOR", ToVarchar(rsTemp("����ҽʦ"), 20)) ' ���ҽ��
    Call InsertChild(mdomInput.documentElement, "OFFICE", ToVarchar(rsTemp("��Ժ����"), 20)) ' ����
    Call InsertChild(mdomInput.documentElement, "REGDATE", Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss")) ' ��Ժʱ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.����) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))  ' ��������
    Call InsertChild(mdomInput.documentElement, "GSRDBH", gstr�����϶����)  ' �����϶����
    Call InsertChild(mdomInput.documentElement, "KFZYBZ", gint���˿���סԺ)  ' ���˿���סԺ
    
    '���ýӿ�
    If CommServer("HOSPREG") = False Then Exit Function
    
    Dim intסԺ�����ۼ� As Integer
    Dim cur�������� As Currency
    Dim cur�����ۼ� As Currency
    Dim cur����ͳ���޶� As Currency
    Dim curͳ�ﱨ���ۼ� As Currency
    Dim cur���ͳ���޶� As Currency
    Dim cur���ͳ���ۼ� As Currency
    Dim str������Ϣ As String
    
    intסԺ�����ۼ� = Val(GetElemnetValue("HOSPTIMES"))
    cur�������� = Val(GetElemnetValue("STARTFEE"))
    cur�����ۼ� = Val(GetElemnetValue("STARTFEEPAID"))
    cur����ͳ���޶� = Val(GetElemnetValue("FUND1LMT"))
    curͳ�ﱨ���ۼ� = Val(GetElemnetValue("FUND1PAID"))
    cur���ͳ���޶� = Val(GetElemnetValue("FUND2LMT"))
    cur���ͳ���ۼ� = Val(GetElemnetValue("FUND2PAID"))
    
    str������Ϣ = GetElemnetValue("LOCKINFO")
    Do Until str������Ϣ = ""
        strTemp = Left(str������Ϣ, 2)
        str������Ϣ = Mid(str������Ϣ, 41)
        
        str��ʾ = str��ʾ & Switch(strTemp = "11", "�����������", strTemp = "21", "��������", strTemp = "31", "������ͳ��Ƿ��", _
                                   strTemp = "32", "�����ͳ��δ�ɷ�", strTemp = "41", "��ͣ��", strTemp = "51", "���˱�")
    Loop
    str��ʾ = str��ʾ & GetElemnetValue("NOTE")
    If str��ʾ <> "" Then
        MsgBox "��ע���ҽ�����������" & Mid(str��ʾ, 2) & "��", vbInformation, gstrSysName
    End If
    '<SPECCALFLAG>��������־</SPECCALFLAG>
    '<RECKONINGTYPE>���㷽ʽ</RECKONINGTYPE>
    '<SINGLEILLNESSCODE>�����ֱ���</SINGLEILLNESSCODE>
    str������㷽ʽ = GetElemnetValue("SPECCALFLAG")
    str�������˵�� = GetElemnetValue("SPECCALFLAGTXT")
    str���㷽ʽ = GetElemnetValue("RECKONINGTYPE")
    str���㲡�ֱ��� = GetElemnetValue("SINGLEILLNESSCODE")
    str���㲡������ = GetElemnetValue("SINGLEILLNESSNAME")
    If str������㷽ʽ = "" Then str������㷽ʽ = "00"
    
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
        dbl�ʻ���� & ",0,0," & curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & _
         "," & cur����ͳ���޶� & "," & cur���ͳ���޶� & "," & cur���ͳ���ۼ� & ",'" & ToVarchar(str��ʾ, 100) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'˳���','''" & GetElemnetValue("BILLNO") & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�˳���")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'�����϶����','''" & gstr�����϶���� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���湤���϶����")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'���˿���סԺ','" & gint���˿���סԺ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���湤�˿���סԺ��־")
    '�������õ����㷽ʽ
    gstrSQL = "ZL_ҽ������סԺ��Ϣ_INSERT(" & _
              lng����ID & "," & lng��ҳID & ",'" & gstrUserName & "',2," & str������� & ",'" & str���㲡������ & "',NULL,NULL," & _
              "NULL,NULL,NULL,NULL,NULL,'" & str���㲡�ֱ��� & "',NULL,NULL,NULL,NULL,'" & str���㷽ʽ & "','" & str������㷽ʽ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�������ص����㷽ʽ")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'������','''" & str���㲡�ֱ��� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���浥���ֱ���")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'���㷽ʽ','''" & str���㷽ʽ & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���浥���ֱ���")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, Optional ByVal bln��Ժ As Boolean = False) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
            'ȡ��Ժ�Ǽ���֤�����ص�˳���
            
    '�޸�˵��
    'ʱ�䣺2005-01-14
    '�޸��ˣ�����
    '�޸����ݣ���Ժ�Ǽǽӿ�������Ρ�ICD���룬Ҳ�����ṩ���ϴ�ICD����Ľӿڣ����뿭��ϵ���ݶ����ӿڲ��ϴ�ICD���룬���ϴ�ICD�������
    
    Dim strҽ���� As String
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str��� As String, str������� As String
    Dim str������ As String, str��Ժת�� As String, lngPos As Long
    Dim str��Ժ���� As String
    
    On Error GoTo errHandle
    
    If mblnҽ����Ժ Or bln��Ժ Then
        '����ҽ��Ҫ���Ժ���ڱ�����ڽ������ڣ�������ҽԺ���ȳ�Ժ����㣬��ˣ��ڵ�ҽ����Ժ����ʱ��ȡ���һ�ν��������+1����Ϊ��Ժ���ڴ���ȥ
        gstrSQL = "SELECT to_Char(�շ�ʱ��,'yyyy-MM-dd hh24:mi:ss') AS �շ�ʱ�� FROM ���˽��ʼ�¼ " & _
                 " WHERE ID=( " & _
                 "     SELECT MAX(��¼ID)  " & _
                 "     FROM ���ս����¼  " & _
                 "     WHERE ����ID=[1] AND ����=[2])"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���һ�ν��ʵ�ʱ��", lng����ID, TYPE_������)
        If rsTemp.RecordCount <> 0 Then str��Ժ���� = Format(DateAdd("s", 1, rsTemp!�շ�ʱ��), "yyyy-MM-dd HH:mm:ss")
        
        '�����ݿ��ж����Ѵ洢��ֵ
        gstrSQL = "select ����,ҽ����,˳��� from �����ʻ� where ����ID=[1] and ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, TYPE_������)
        
        strҽ���� = IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����"))
        
        '��ò��˳�Ժ��Ϣ
        gstrSQL = "SELECT A.��Ժ��ʽ,nvl(C.������,B.סԺ��) AS ������  " & _
                 " FROM ������ҳ A,������Ϣ B,סԺ������¼ C " & _
                 " WHERE A.����ID=[1] AND A.��ҳid=[2] AND A.����id=B.����id AND A.����id=C.����id(+)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
        str������ = Nvl(rsTemp("������"), lng����ID)
        Select Case rsTemp("��Ժ��ʽ")
            Case "����", "����"
                str��Ժת�� = "1"
            Case "��ת"
                str��Ժת�� = "2"
            Case "����"
                str��Ժת�� = "3"
            Case Else
                str��Ժת�� = "9"
        End Select
        
        '��ò��˳�Ժ���
        gstrSQL = "select A.������Ϣ from ������ A where A.����ID=[1] and A.��ҳID=[2]" & _
                  " and A.�������=3 and A.��ϴ���=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
        If rsTemp.EOF = False Then
            str��� = Nvl(rsTemp("������Ϣ"), "����")
            '����ͬ��ʽ�ķָ���ͳһ
            str��� = Replace(str���, "��", ",")
            str��� = Replace(str���, "��", ",")
            str��� = Replace(str���, "��", ",")
            str��� = Replace(str���, ";", ",")
            lngPos = InStr(str���, ",")
            If lngPos > 0 Then
                str������� = Mid(str���, lngPos + 1)
                str��� = Mid(str���, 1, lngPos - 1)
            End If
        Else
            str��� = "����"   '��ϲ�����β���Ϊ��
        End If
            
        '���������Ժ��Ϣ
        datCurr = zlDatabase.Currentdate
        gstrSQL = "select A.סԺҽʦ,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ���� from ������ҳ A,���ű� B " & _
                 " Where A.��Ժ����ID = B.ID And A.����ID =[1] And A.��ҳID = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
        If str��Ժ���� = "" Then str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss")
        If str��Ժ���� = "" Then str��Ժ���� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        '��XML DomDocument������г�ʼ��
        If InitXML = False Then Exit Function
        Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
        Call InsertChild(mdomInput.documentElement, "DOCNO", str������)          ' ������
        Call InsertChild(mdomInput.documentElement, "DIAGNOSES", ToVarchar(str���, 128))          ' ���
        Call InsertChild(mdomInput.documentElement, "OTHERDIAGNOSES", ToVarchar(str�������, 128)) ' �������
        Call InsertChild(mdomInput.documentElement, "OUTTYPE", str��Ժת��)                        ' ת�����
        Call InsertChild(mdomInput.documentElement, "ICD", "")                       ' ICD��������
        Call InsertChild(mdomInput.documentElement, "DOCTOR", ToVarchar(Nvl(rsTemp("סԺҽʦ"), "ZLHIS"), 20))  ' ���ҽ��
        Call InsertChild(mdomInput.documentElement, "OFFICE", ToVarchar(rsTemp("��Ժ����"), 20))   ' ����
        Call InsertChild(mdomInput.documentElement, "REGDATE", str��Ժ����) ' ��Ժ����
        Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.����) ' ����Ա
        Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))  ' ��������
        
        '���ýӿ�
        If CommServer("HOSPOUT") = False Then Exit Function
    End If
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    MsgBox IIf(mblnҽ����Ժ Or bln��Ժ, "�ɹ�����HIS��ҽ����Ժ��", "��������HIS��Ժ��"), vbInformation, gstrSysName
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim str˳��� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    gstrSQL = " Select ˳��� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ˮ��", TYPE_������, lng����ID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "û���ҵ��ò��˵�ҽ��������", vbInformation, gstrSysName
        Exit Function
    End If
    str˳��� = Nvl(rsTemp!˳���)

    '�˴�����ҽ�����óɹ������м�飬���Ժʱ���ܽ�������HIS��Ժ
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", str˳���) ' ��Ժʱ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.����) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss"))  ' ��������
    Call CommServer("RETHOSPOUT")

    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")

    ��Ժ�Ǽǳ���_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long, Optional ByVal bln���� As Boolean = True) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim cn�ϴ� As New ADODB.Connection, rsTemp As New ADODB.Recordset, rs���� As ADODB.Recordset, rs��� As New ADODB.Recordset, rs���� As New ADODB.Recordset
    Dim lng����ID As Long, lng��ҳID As Long, str��Ŀ���� As String
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String
    Dim cur�����ʻ� As Double, curͳ��֧�� As Double, cur��ͳ�� As Double, cur�������� As Double
    Dim dbl�����ܷ���  As Double, dblHIS�ܷ��� As Double, dbl��� As Double
    Dim cur����Ա���� As Double, curҽ���չ˹���Ա���� As Double, cur������Ա���� As Double, int������־ As Integer
    Dim strҽ�� As String, str���� As String, str����ҩƷ As String
    Dim bln����ҩƷ As Boolean, bln����ҩƷ����״̬ As Boolean, bln����δ����������ҩƷ As Boolean
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim rsTmp       As ADODB.Recordset
    On Error GoTo errHandle
    mlng����ID = 0         '��ʼ����ֻҪһѡ���ˣ��ͻ���ñ����̣�Ҳ�ͻ����0
    
    If rsExse.RecordCount = 0 Then
        ShowMsgbox "�ò���û���з������ã��޷����н��������"
        Exit Function
    End If
    
    rsExse.MoveFirst
    '������һ�����Ӵ����Դﵽ���ܵ�ǰ��������Ŀ���
    Set cn�ϴ� = GetNewConnection
    '�˴�����ȷ���ǵò����ģ�����Ҫǿ��ˢ��
    Screen.MousePointer = vbDefault
    
    'ȡ�ò��˵Ļ�����Ϣ
    gstrSQL = " Select A.��Ա���,A.�����϶����,A.���˿���סԺ,B.����,C.סԺ���� from �����ʻ� A,���ղ��� B,������Ϣ C" & _
              " Where A.����ID=[1] And A.����ID=C.����ID and A.����=[2]  and A.����ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺԤ��", lng����ID, TYPE_������)
    If rsTemp.EOF = False Then
        lng��ҳID = rsTemp!סԺ����
        gstr�����϶���� = Nvl(rsTemp!�����϶����)
        gint���˿���סԺ = Nvl(rsTemp!���˿���סԺ, 0)
        str��Ա��� = Nvl(rsTemp("��Ա���"), "")
        'ת����Ա���
        str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21", _
                          str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", _
                          str��Ա��� = "��ͨ����", "41", str��Ա��� = "�ͱ�����", "42", _
                          str��Ա��� = "������Ա", "43", str��Ա��� = "���ռ�ͥ", "44", _
                          str��Ա��� = "�ضȲм�", "45", True, "11") '����������ʾ�������ͥ�������ݿ�ֻ��8λ��ֻ�ܱ���Ϊ���ռ�ͥ
    End If
    
    mstr���� = ""
    mstr���� = ""
    If Get��֤_����(1, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    Screen.MousePointer = vbHourglass
    
    mbln��������� = False
    If bln���� Then
        If MsgBox("�Ƿ������������㣨�������ڲ����ӷѵ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then mbln��������� = True
    End If
    
    '����¼����
    If gbln����������� = False And str���� = "1" Then
        If ging������� <> 0 Then
            '0�����,1�������,2��ֹ����
            gstrSQL = "Select ��ע From �����ǼǼ�¼_���� Where ����=[1] And ����ID=[2] And ��ҳID=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSQL, TYPE_������, lng����ID, lng��ҳID)
            If rsTemp.RecordCount = 0 Then
                If ging������� = 1 Then
                    If MsgBox("�ò���δ�Ǽǲ�����¼���Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
                Else
                    ShowMsgbox "�ò���δ�Ǽǲ�����¼�����ȵǼǣ�"
                    Exit Function
                End If
            Else
                '�Ѳ����Ǽǵı�ע�����ʾ����
                ShowMsgbox "�ò��������±�ע��Ϣ�����ʵ��" & vbNewLine & Nvl(rsTemp!��ע)
            End If
        End If
    End If
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDTYPE", gintType)                ' �����
    Call InsertChild(mdomInput.documentElement, "CARDDATA", str����)                 ' ��������
    Call InsertChild(mdomInput.documentElement, "SNO", gstrIDNO)                     ' ��ᱣ�Ϻ�
    Call InsertChild(mdomInput.documentElement, "IPADDR", gstrClientIP)              ' IP��ַ
    Call InsertChild(mdomInput.documentElement, "PSAMNO", gstrPSAMNO)                ' PSAM��оƬ
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", strҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "ISCAL", 0)         ' �Ƿ����
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", "0")     ' �˻�֧����
    Call InsertChild(mdomInput.documentElement, "INVOICENO", "") ' ��Ʊ��
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) ' ��������
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' ������ϸ
    
    rsExse.Sort = " �Ǽ�ʱ�� asc"
    Do Until rsExse.EOF
        '����Ƿ�������Ŀ
        bln����ҩƷ = False                     '��ǰ�Ƿ�����ҩƷ
        bln����ҩƷ����״̬ = False             '��ǰ����ҩƷ�Ƿ�����
        
        If IIf(IsNull(rsExse("�Ƿ��ϴ�")), "0", rsExse("�Ƿ��ϴ�")) = "0" And Nvl(rsExse!���, 0) <> 0 Then
            gstrSQL = "SELECT C.ҩƷID,C.���,E.���� AS ����  FROM ҩƷĿ¼ C,ҩƷ��Ϣ D,ҩƷ���� E WHERE C.ҩƷID=" & rsExse("�շ�ϸĿID") & " AND C.ҩ��ID=D.ҩ��ID AND D.����=E.����"
            gstrSQL = "select A.���,A.����,B.��Ŀ����,nvl(A.���,F.���) AS ���,F.����,A.���㵥λ from �շ�ϸĿ A,����֧����Ŀ B,(" & gstrSQL & _
                    ") F where A.ID=[1] and A.ID=B.�շ�ϸĿID  AND A.Id=F.ҩƷID(+) and B.����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺԤ��", CLng(rsExse("�շ�ϸĿID")), TYPE_������)
            If rsTemp.EOF = True Then
                ShowMsgbox "����Ŀδ����ҽ�����룬���ܽ��㡣"
                Exit Function
            End If
            
            'ֻ�ϴ�ֻ���ݹ�������
            strҽ�� = ToVarchar(IIf(IsNull(rsExse("ҽ��")), UserInfo.����, rsExse("ҽ��")), 20)
            str���� = ToVarchar(IIf(IsNull(rsExse("��������")), UserInfo.����, rsExse("��������")), 56)
            
             On Error Resume Next
            '20100301 BY ZYB
            '�������ҩƷ��Ŀ�����δ�������ϴ���Ҳ������HIS�ܽ��Ա�������ĺ˶��ܽ����˳��ͨ��
            '���������ҩƷ��Ŀδ�����ģ�������ʾ�������������ִ��
            str��Ŀ���� = ""
            str��Ŀ���� = Nvl(rsExse!���ձ���)      'ȱʡ�Ե�ǰ��������Ϊ׼
            
            If mbln����ҩƷ��ʾ Then
                gstrSQL = " Select b.��־ From " & IIf(gblnHIS1026 = True, "סԺ���ü�¼", "���˷��ü�¼") & " a,����ҩƷ�շ� b"
                gstrSQL = gstrSQL & vbCrLf & "Where a.ID = b.����id And a.����ID = [1] And a.��ҳID = [2]"
                Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҩƷ", lng����ID, lng��ҳID)
                
                gstrSQL = " Select 1 From �շ�ϸĿ Where ˵�� Is Not NULL And ��� IN ('5','6','7') And ID=[1]"
                Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ��Ϣ", CLng(rsExse!�շ�ϸĿID))
                If rs����.RecordCount <> 0 Then
                    bln����ҩƷ = True
                    gstrSQL = " Select ��־ From ����ҩƷ�շ� " & _
                              " Where ����ID=[1] And ��ҳID=[2] And ����ID=" & _
                              " (Select ID from " & IIf(gblnHIS1026 = True, "סԺ���ü�¼", "���˷��ü�¼") & " where ��¼����=[1] and ��¼״̬=[2] and No=[3] and ���=[4])"
                              
                    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ŀ��Ϣ", lng����ID, lng��ҳID, CLng(rsExse!��¼����), CLng(rsExse!��¼״̬), CStr(rsExse!NO), CLng(rsExse!���))
                    If rs����.RecordCount <> 0 Then
                        bln����ҩƷ����״̬ = True
                        If Nvl(rs����!��־, 0) = 0 Then
                           '��ҩȫ�Էѱ��룺81900090009
                            '20110201��ʼ�Է���ҩ����:   810851900099,����ǿ�޸�
                            '20110201��ʼ�Է��г�ҩ����: 820851900099,����ǿ�޸�
                            '�г�ҩ���в�ҩȫ�Էѱ���:   829000900099,����ǿ�޸�
                            If rsExse!�շ���� = "5" Then
                            str��Ŀ���� = "810851900099"
                            ElseIf rsExse!�շ���� = "6" Then
                            str��Ŀ���� = "820851900099"
                            Else
                            str��Ŀ���� = "829000900099"
                            End If
                           ' str��Ŀ���� = IIf(rsExse!�շ���� = "5", "81900090009", "82900090009")
                        End If
                    Else
                        bln����δ����������ҩƷ = True  '�Ƿ����δ����������ҩƷ
                    End If
                End If
            End If
            
            Err = 0
            On Error GoTo errHandle
            
            '�����Ӥ���ѣ����Էѱ���
            If Nvl(rsExse!Ӥ����, 0) <> 0 Then
               '20110201��ʼ�Է���ҩ����:   810851900099,����ǿ�޸�
                 '20110201��ʼ�Է��г�ҩ����: 820851900099,����ǿ�޸�
                 '�г�ҩ���в�ҩȫ�Էѱ���:   829000900099 ,����ǿ�޸�
                '����ȫ�Էѱ��룺34900099
                If rsExse!�շ���� = "5" Then
                    str��Ŀ���� = "810851900099"
                ElseIf rsExse!�շ���� = "6" Then
                     str��Ŀ���� = "820851900099"
                ElseIf rsExse!�շ���� = "7" Then
                    str��Ŀ���� = "829000900099"
                Else
                    str��Ŀ���� = "34900099"
                End If
            Else
                '������ձ���Ϊ�գ�����Ҫ�û�ѡ�����
                If str��Ŀ���� = "" Then
                    str��Ŀ���� = GetItemInsure_����(lng����ID, rsExse!�շ�ϸĿID, False)
                End If
                If str��Ŀ���� = "" Then str��Ŀ���� = Nvl(rsExse!ҽ����Ŀ����)
            End If
            
            If Not bln����ҩƷ Or (bln����ҩƷ And bln����ҩƷ����״̬) Then
                Set nodRow = InsertChild(nodRowset, "ROW", "")
                Call nodRow.setAttribute("ITEMSERIAL", ToVarchar(rsExse("NO") & "_" & rsExse("���") & "_" & rsExse("��¼����") & "_" & rsExse("��¼״̬"), 20)) '�������ţ�����Ψһ��������
                Call nodRow.setAttribute("ITEMCODE", ToVarchar(str��Ŀ����, 12))
                Call nodRow.setAttribute("ITEMNAME", ToVarchar(rsExse("�շ�����"), 72))
                Call nodRow.setAttribute("SUBJECT", Subject(rsTemp("���")))
                Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rsTemp("���"), 40))
                Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rsTemp("����"), 20))
                Call nodRow.setAttribute("UNIT", ToVarchar(rsTemp("���㵥λ"), 20))
                Call nodRow.setAttribute("PRICE", Format(rsExse("�۸�"), "0.0000"))
                Call nodRow.setAttribute("QUANTITY", Format(rsExse("����"), "0.00"))
                Call nodRow.setAttribute("FROMOFFICE", str����)   '��������
                Call nodRow.setAttribute("FROMDOCT", strҽ��)     '����ҽ��
                Call nodRow.setAttribute("TOOFFICE", str����)    '�ܵ�����
                Call nodRow.setAttribute("TODOCT", strҽ��)      '�ܵ�ҽ��
                '����ʱ��ʱ��Ϊ�˱�֤ͬһ������Ŀ�ĵ��շ�ʱ�䲻ͬ������ڵǼ�ʱ���ϰ���ż�������
                Call nodRow.setAttribute("DODATE", Format(DateAdd("s", rsExse("���") + IIf(rsExse!��¼״̬ = 2, 1, 0), rsExse("�Ǽ�ʱ��")), "yyyy-MM-dd HH:mm:ss")) '��������
                Call nodRow.setAttribute("NOTE", ToVarchar(rsExse("ժҪ"), 512))     '��ע
            End If
        End If
        
        If Not bln����ҩƷ Or (bln����ҩƷ And bln����ҩƷ����״̬) Then
            cur�������� = cur�������� + Round(rsExse("���"), 2)
        Else
            'J12345671100001
            str����ҩƷ = str����ҩƷ & "," & rsExse("NO") & rsExse("��¼����") & rsExse("��¼״̬") & String(5 - Len(CStr(rsExse("���"))), "0") & rsExse("���")
        End If
        'XieRong 2010-10-22 ����ҽ������
        If Nvl(rsExse!ҽ����Ŀ����) = "" Then
            Dim rsID        As ADODB.Recordset
            gstrSQL = "select ID from " & IIf(gblnHIS1026 = True, "סԺ���ü�¼", "���˷��ü�¼") & " where NO=[1] and ���=[2] and ��¼����=[3] and ��¼״̬=[4]"
            Set rsID = zlDatabase.OpenSQLRecord(gstrSQL, "����ID", rsExse("NO"), rsExse("���"), rsExse("��¼����"), rsExse("��¼״̬"))
            If rsID.RecordCount > 0 Then
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rsID!ID & ",0," & TYPE_������ & ",,'" & str��Ŀ���� & "')"
                zlDatabase.ExecuteProcedure gstrSQL, "����ҽ����Ϣ"
            End If
        End If
        rsExse.MoveNext
    Loop
    
    '���ýӿ�
    If CommServer("CALHOSP", IIf(mbln���������, "1", "0")) = False Then Exit Function
    '����ǿ�������ٴ������Ե�ҽ����������ȷ���غ��ٴ��ϱ��
    str����ҩƷ = str����ҩƷ & ","
    If rsExse.RecordCount > 0 Then rsExse.MoveFirst
    Do Until rsExse.EOF
        If IIf(IsNull(rsExse("�Ƿ��ϴ�")), "0", rsExse("�Ƿ��ϴ�")) = "0" Then
            'Ϊ�������ü�¼�����ϴ���־���ϴ�һ������һ��
            If InStr(1, str����ҩƷ, "," & rsExse("NO") & rsExse("��¼����") & rsExse("��¼״̬") & String(5 - Len(CStr(rsExse("���"))), "0") & rsExse("���") & ",") = 0 Then
                gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
                cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
            End If
        End If
        rsExse.MoveNext
    Loop
    
    cur�����ʻ� = Val(GetElemnetValue("ACCTPAY"))
    If str��Ա��� = "32" Or str��Ա��� = "34" Then
        curͳ��֧�� = Val(GetElemnetValue("ALLOWFUND"))
    Else
        curͳ��֧�� = Val(GetElemnetValue("FUND1PAY"))
    End If
    cur��ͳ�� = Val(GetElemnetValue("FUND2PAY"))
    
'    <FUND3PAY>����Ա����֧��</FUND3PAY>
'    <CAREPAY>ҽ���չ���Ա�����Ա����</CAREPAY>
'    <FUND3OVER>������޶��Ա����</ FUND3OVER >
'    <BEARINGFLAG>������־</BEARINGFLAG>
    cur����Ա���� = Val(GetElemnetValue("FUND3PAY"))
    dbl�����ܷ��� = Val(Format(GetElemnetValue("CALFEEALL"), "#0.00;-#0.00;0;"))
    dblHIS�ܷ��� = Val(Format(GetElemnetValue("HOSPFEEALL"), "#0.00;-#0.00;0;"))
    dbl��� = dblHIS�ܷ��� - dbl�����ܷ���
    
    '�ȱȽ��ܶ��Ƿ�һ��
    If gbln����������� = False Then
        If Format(cur��������, "#0.00") <> Format(dblHIS�ܷ���, "#0.00") Then
            If Abs(Val(Format(cur��������, "#0.00")) - Val(Format(dblHIS�ܷ���, "#0.00"))) > 0.5 Then
                MsgBox "HIS�ܶ���ҽ�����յ����ܷ��ò�һ�£���������㣡" & vbCrLf & _
                    "HIS:" & Format(cur��������, "#0.00") & Space(10) & "ҽ��:" & Format(dblHIS�ܷ���, "#0.00"), vbInformation, gstrSysName
                Exit Function
            Else
                If MsgBox("HIS�ܶ���ҽ�����յ����ܷ��ò�һ�£������ǵ��۾��Ȳ�����������Ƿ���㣿" & vbCrLf & _
                    "HIS:" & Format(cur��������, "#0.00") & Space(10) & "ҽ��:" & Format(dblHIS�ܷ���, "#0.00"), vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
            End If
        End If
    End If
    
    If gbln����������� = False Then
        If bln����δ����������ҩƷ Then
            ShowMsgbox "����δ����������ҩƷ������������ܽ��㣡"
        End If
    End If
    
    '���没�˸����ʻ����
    mstrҽ���� = strҽ����
    mdbl��� = cur�����ʻ�
    
    '������ʱ���ݣ�Ϊ���������׼��
    With g��������
        .�������ý�� = cur��������
    End With
    Screen.MousePointer = 0
    If gbln����������� = False Then frm����������Ϣ.Show 1
    
    Dim str������� As String, strMsg As String
    '��ʾ�ò��˵����������㷽ʽ��������Ա�˶�
    '����Ƿ�ѡ�����㷽ʽ����㷽ʽ��û��ѡ��ȱʡΪ������
    gstrSQL = " Select A.����Ա, A.����, NVL(A.��Ч��־,2) AS ��Ч��־,Nvl(A.�������,B.�������) AS �������, A.��������, A.��ʼ����, A.��������, " & _
              "        A.�����ֱ���_����, A.�����޶�, A.ͳ�����,A.���˸���, NVL(A.���㷽ʽ,'0') AS ���㷽ʽ, " & _
              "        A.������, A.����ͳ���嵥��׼, A.����ͳ��ֵ�����, A.���������׼, A.�����ֵ�����, Nvl(A.���㷽ʽ,'1') AS ���㷽ʽ,NVL(A.������㷽ʽ,'00') AS ������㷽ʽ " & _
              " From ҽ������סԺ��Ϣ A,�����ʻ� B,������Ϣ C" & _
              " Where B.����ID=C.����ID And C.����ID=A.����ID(+) And C.סԺ����=A.��ҳID(+) And C.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ�ѡ�����㷽ʽ����㷽ʽ", lng����ID)
    If Mid(rsTemp!������㷽ʽ, 2, 1) = "0" Then    'ϵͳ�̶��Ļ����ڶ�λ���㣬�����Լ����õĲ���ʾ����������Ա��飬���򲻼��
        Select Case rsTemp!�������
        Case 2
            strMsg = "���������ҵ����ҽ�Ʊ���"
        Case 3
            strMsg = "������𣺻�����ҵ��λ����ҽ�Ʊ���"
        Case 4
            strMsg = "���������ҵ��������"
        Case 5
            strMsg = "������𣺻�����ҵ��λ��������"
        Case 6
            strMsg = "������𣺾�����"
        Case 7
            strMsg = "������𣺹��˱���"
        Case Else
            strMsg = "���������ҵ����ҽ�Ʊ���"
        End Select
        strMsg = "���ڽ���ǰ��ϸ�˶Ըò��˵�סԺ��Ϣ��" & vbCrLf & strMsg & vbCrLf
        If rsTemp!��Ч��־ = 1 Then
            '���㷽ʽ
            strMsg = strMsg & "���㷽ʽ��" & rsTemp!���㷽ʽ
            If rsTemp!���㷽ʽ = "1" Then
                strMsg = strMsg & vbCrLf & _
                    "������Ϣ����" & Nvl(rsTemp!�����ֱ���_����) & "��" & Nvl(rsTemp!��������) & vbCrLf & _
                    "�����޶" & Nvl(rsTemp!�����޶�) & "��ͳ����ɣ�" & Nvl(rsTemp!ͳ�����) & "�����˸�����" & Nvl(rsTemp!���˸���)
            End If
        Else
            '���㷽ʽ
            Select Case rsTemp!���㷽ʽ
            Case 2
                strMsg = strMsg & "���㷽ʽ����֢��������"
            Case 3
                strMsg = strMsg & "���㷽ʽ�������ְ��˴ζ���������㷽ʽ"
            Case 4
                strMsg = strMsg & "���㷽ʽ�������ְ��ն���������㷽ʽ"
            Case 5
                strMsg = strMsg & "���㷽ʽ���������հ�������"
            Case 6
                strMsg = strMsg & "���㷽ʽ�������ְ�������"
            Case Else
                strMsg = strMsg & "���㷽ʽ�����������㷽ʽ������������Ϊ�ǰ��ɷ�ʽ��"
            End Select
            If rsTemp!���㷽ʽ <> "1" Then
                strMsg = strMsg & vbCrLf & _
                    "������Ϣ����" & Nvl(rsTemp!������) & "��" & Nvl(rsTemp!��������)
            End If
        End If
    End If
    Dim str��� As String, str������� As String, lngPos As Long
    '��ò��˳�Ժ���
    gstrSQL = "select A.������Ϣ from ������ A where A.����ID=[1] and A.��ҳID=[2]" & _
              " and A.�������=3 and A.��ϴ���=1"
    Set rs��� = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    If rs���.EOF = False Then
        str��� = Nvl(rs���("������Ϣ"), "����")
        '����ͬ��ʽ�ķָ���ͳһ
        str��� = Replace(str���, "��", ",")
        str��� = Replace(str���, "��", ",")
        str��� = Replace(str���, "��", ",")
        str��� = Replace(str���, ";", ",")
        lngPos = InStr(str���, ",")
        If lngPos > 0 Then
            str������� = Mid(str���, lngPos + 1)
            str��� = Mid(str���, 1, lngPos - 1)
        End If
    Else
        str��� = "����"   '��ϲ�����β���Ϊ��
    End If
    strMsg = strMsg & vbCrLf & "��Ժ��ϣ�" & str���
    If Trim(str�������) <> "" Then strMsg = strMsg & "������ϣ�" & str�������
    '��Ժ��ʽ Modify By �̳ظ� 2010-01-16 ����ҽѧԺҪ��
    Dim rs��Ժ��ʽ As New ADODB.Recordset
    gstrSQL = "SELECT ��Ժ��ʽ  FROM ������ҳ where ����=[1] And ����ID=[2] And ��ҳID=[3]"
    Set rs��Ժ��ʽ = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ��ʽ", TYPE_������, lng����ID, lng��ҳID)
    If Nvl(rs��Ժ��ʽ!��Ժ��ʽ, "��") <> "��" Then strMsg = strMsg & vbCrLf & "��Ժ��ʽ��" & Nvl(rs��Ժ��ʽ!��Ժ��ʽ)
    If gbln����������� = False Then ShowMsgbox strMsg
    
    
    '����Ԥ������
    סԺ�������_���� = "ҽ������;" & curͳ��֧�� & ";0"
    If cur�����ʻ� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|�����ʻ�;" & cur�����ʻ� & ";1" '�����޸ĸ����ʻ�
    End If
    If InStr(1, "4,5", rsTemp!�������) <> 0 Then   '�����������Ƹ�Ϊ��ǰ����
        סԺ�������_���� = סԺ�������_���� & "|��ǰ����;" & cur��ͳ�� & ";0"
    Else
        סԺ�������_���� = סԺ�������_���� & "|��ͳ��;" & cur��ͳ�� & ";0"
    End If
    סԺ�������_���� = סԺ�������_���� & "|ҽ�Ʋ���;" & Format(cur����Ա����, "#0.00;-#0.00;0;") & ";0"
    'ֻ�н��㷽ʽΪ�����ְ��ɵĲŴ��ڲ����ʣ���������ֽ�
    If rsTemp!��Ч��־ = 1 And rsTemp!���㷽ʽ = "�����ְ��ɽ���" Then סԺ�������_���� = סԺ�������_���� & "|������;" & dbl��� & ";0"
    mlng����ID = lng����ID  '��ʾ�ò����Ѿ��������������
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp   As ADODB.Recordset
    Dim lng��ҳID As Long
    Dim curȫ�Ը� As Double, cur�ҹ��Ը� As Double, curͳ��֧�� As Double, cur�ʻ���� As Currency
    Dim curͳ���Ը� As Double, cur�����Ը� As Double, cur�����Ը� As Double
    Dim cur��ͳ�� As Double, cur���Ը� As Double, cur�����ʻ� As Double, cur���� As Currency
    Dim cur����Ա���� As Double, curҽ���չ˹���Ա���� As Double, cur������Ա���� As Double, int������־ As Integer
    Dim dblHIS�ܷ��� As Double, dbl�����ܷ��� As Double, dbl��� As Double, dblҽ���ܷ��� As Double
    
    Dim int���㷽ʽ As Integer, str�����ֱ��� As String
    Dim int���㷽ʽ As Integer, str������ As String
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, datCurr As Date, strNO As String
    Dim str����˳��� As String, str������ As String
    Dim str֧������ As String, str������� As String
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String
    Dim str����XML As String, str���㷽ʽ As String, str���㲡�ֱ��� As String, str�������˵�� As String, str��������־ As String
    
    If mlng����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣"
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    'ȡ��ҳID
    gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
    If Split(rsTmp!�汾��, ".")(0) = 10 And Split(rsTmp!�汾��, ".")(1) >= 34 Then
        gstrSQL = " Select A.ҽ����,A.�������,B.סԺ���� AS ��ҳID,A.������,A.���㷽ʽ,A.�����ֱ���_����,A.���㷽ʽ,C.��Ժ��ʽ,A.֧������ " & _
              " From �����ʻ� A,������Ϣ B,������ҳ C " & _
              " Where A.����ID=B.����ID And B.����ID=C.����ID And B.��ҳID=C.��ҳID And A.����ID=[1]"
    Else
        gstrSQL = " Select A.ҽ����,A.�������,B.סԺ���� AS ��ҳID,A.������,A.���㷽ʽ,A.�����ֱ���_����,A.���㷽ʽ,C.��Ժ��ʽ,A.֧������ " & _
              " From �����ʻ� A,������Ϣ B,������ҳ C " & _
              " Where A.����ID=B.����ID And B.����ID=C.����ID And B.סԺ����=C.��ҳID And A.����ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID", lng����ID)
    cur�ʻ���� = �������_����(rsTemp!ҽ����)
    lng��ҳID = rsTemp!��ҳID
    str������� = Nvl(rsTemp!�������)
    str������ = Nvl(rsTemp!������)
    int���㷽ʽ = Nvl(rsTemp!���㷽ʽ, 1)
    str�����ֱ��� = Nvl(rsTemp!�����ֱ���_����)
    int���㷽ʽ = Nvl(rsTemp!���㷽ʽ, 0)
    str֧������ = Nvl(rsTemp!֧������, "31")
    
    '����Ƿ����δ�趨����ҩƷ�����ݣ������������
   If mbln����ҩƷ��ʾ Then
        gstrSQL = " Select distinct A.ID" & _
                  " From " & IIf(gblnHIS1026 = True, "סԺ���ü�¼", "���˷��ü�¼") & " A,�շ�ϸĿ B" & _
                  " Where Nvl(A.ʵ�ս��,0)<>0 And Nvl(A.����,0)<>0 And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.��¼״̬,0)<>0 And Nvl(A.Ӥ����,0)=0" & _
                  " And B.��� IN ('5','6','7') And B.˵�� Is Not NULL And A.�շ�ϸĿID+0=B.ID And A.����ID=[1] And A.��ҳID=[2]" & _
                  " And Nvl(A.�Ƿ��ϴ�, 0) = 0 And nvl(A.ʵ�ս��,0)<>0 " & _
                  " MINUS" & _
                  " Select distinct ����ID From ����ҩƷ�շ�" & _
                  " Where ����ID=[1] And ��ҳID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δ�趨����ҩƷ�����ݣ������������", lng����ID, lng��ҳID)
        If rsTemp.RecordCount <> 0 Then
            Err.Raise 9000, gstrSysName, "�����ڲ�������ҩƷδ�����������������㣡"
            Exit Function
        End If
    End If
    
    str����XML = mdomInput.xml
    'ȡ���㷽ʽ�����㲡�֣�������Ϊ������δ���ȥ
'    <RECKONINGTYPE>���㷽ʽ</RECKONINGTYPE>
'    <SINGLEILLNESSCODE>�����ֱ���</SINGLEILLNESSCODE>
    str���㷽ʽ = GetElemnetValue("RECKONINGTYPE")
    str���㲡�ֱ��� = GetElemnetValue("SINGLEILLNESSCODE")
    str��������־ = GetElemnetValue("SPECCALFLAG")
    str�������˵�� = GetElemnetValue("SPECCALFLAGTXT")
    
    '����ǿ��ˢ��
    mstr���� = ""
    mstr���� = ""
    '��ȡ������Ϣ
    If Get��֤_����(1, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    '��ȡ�ſ���Ϣ
    str���� = frmˢ��.ShowME
    If str���� = "" Then Exit Function
    str���� = Split(str����, "|")(1)
    str���� = Split(str����, "|")(0)
    
    Screen.MousePointer = vbHourglass
    
    '������ʻ�֧�����
    gstrSQL = "Select Nvl(��Ԥ��,0) as ��� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ��¼���� Not In (11,1) And ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺ����", lng����ID)
    If Not rsTemp.EOF Then cur�����ʻ� = rsTemp("���")
    '�󵥾ݺ�
    gstrSQL = "Select NO,�շ�ʱ�� From ���˽��ʼ�¼ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺ����", lng����ID)
    
    'XML�ĵ��Ѿ���ɳ�ʼ������ʱֻ��Ҫ���²���ֵ
    Call InitXML
    Call mdomInput.loadXML(str����XML)
    Call SetElemnetValue("CARDTYPE", gintType)
    Call SetElemnetValue("CARDDATA", IIf(gintType = 3, "", str����))
    Call SetElemnetValue("SNO", gstrIDNO)
    Call SetElemnetValue("PSAMNO", gstrPSAMNO)
    Call SetElemnetValue("ISCAL", "1")
    Call SetElemnetValue("ACCTWANTTOPAY", Format(cur�����ʻ�, "0.00"))
    Call SetElemnetValue("INVOICENO", "Z_" & rsTemp("NO"))
    Call SetElemnetValue("DODATE", Format(rsTemp("�շ�ʱ��"), "yyyy-MM-dd HH:mm:ss"))
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str����)     ' ��ʽ����ʱ��������
'    <RECKONINGTYPE>���㷽ʽ</RECKONINGTYPE>
'    <SINGLEILLNESSCODE>�����ֱ���</SINGLEILLNESSCODE>
    Call InsertChild(mdomInput.documentElement, "RECKONINGTYPE", str���㷽ʽ) ' ���㷽ʽ
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", str���㲡�ֱ���) ' �����ֱ���
    
    'Ԥ��ʱ�Ѿ����ݣ����ʲ���Ҫ�ٴ�����ϸ����
    Call SetElemnetValue("ROWSET", "")
    '���ýӿ�
    If CommServer("CALHOSP", IIf(mbln���������, "1", "0")) = False Then Exit Function
    
    curȫ�Ը� = Val(GetElemnetValue("FEEOUT"))
    cur�ҹ��Ը� = Val(GetElemnetValue("FEESELF"))
    cur���� = Val(GetElemnetValue("STARTFEE"))
    cur�����Ը� = Val(GetElemnetValue("ENTERSTARTFEE"))
    curͳ��֧�� = Val(GetElemnetValue("FUND1PAY")) + Val(GetElemnetValue("ALLOWFUND"))
    curͳ���Ը� = Val(GetElemnetValue("FUND1SELF"))
    cur��ͳ�� = Val(GetElemnetValue("FUND2PAY"))
    cur���Ը� = Val(GetElemnetValue("FUND2SELF"))
    cur�����Ը� = Val(GetElemnetValue("FEEOVER"))
    
'    <FUND3PAY>����Ա����֧��</FUND3PAY>
'    <CAREPAY>ҽ���չ���Ա�����Ա����</CAREPAY>
'    <FUND3OVER>������޶��Ա����</ FUND3OVER >
'    <BEARINGFLAG>������־</BEARINGFLAG>
    dblҽ���ܷ��� = Val(GetElemnetValue("FEEALL"))
    cur����Ա���� = Val(GetElemnetValue("FUND3PAY"))
    dbl�����ܷ��� = Val(GetElemnetValue("CALFEEALL"))
    dblHIS�ܷ��� = Val(GetElemnetValue("HOSPFEEALL"))
    dbl��� = dblHIS�ܷ��� - dbl�����ܷ���
    
    str������ = GetElemnetValue("BALANCEID")
    str����˳��� = GetElemnetValue("BILLNO")
    
    '��д�����
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� + curͳ���Ը� + cur�����Ը� + cur�����Ը� + cur��ͳ�� + cur���Ը� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� + cur��ͳ�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ���� & "," & cur�ʻ���� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & ",NULL," & cur�����Ը� & "," & _
        g��������.�������ý�� & "," & curȫ�Ը� & "," & cur�ҹ��Ը� & "," & _
        curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & "," & cur���Ը� & "," & cur�����Ը� & "," & cur�����ʻ� & "," & _
        "'" & str������ & "'," & lng��ҳID & ",null,'" & str����˳��� & "',0,'" & AnalyseComputer & "','" & gstrVersion & "','" & str֧������ & "','" & str����˳��� & "'," & _
            "NULL,'" & str�����ֱ��� & "','" & str������� & "',to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "zl_���㸽����Ϣ_Insert (" & lng����ID & ",0,0,0," & cur����Ա���� & "," & cur������Ա���� & "," & curҽ���չ˹���Ա���� & "," & int������־ & "," & _
        "'" & str�����ֱ��� & "'," & int���㷽ʽ & ",'" & str������ & "'," & int���㷽ʽ & "," & dbl�����ܷ��� & "," & dblҽ���ܷ��� & ",'" & gstr�����϶���� & "'," & gint���˿���סԺ & ",NULL,'" & str��������־ & "','" & str�������˵�� & "')"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & curͳ��֧�� + curͳ���Ը� & "," & curͳ��֧�� & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '������㷽ʽ���ǰ����嵥����Ա�����������Ա�����ҵ�ǰ��Ժ��ҽ�����ˣ�����ʾ����ԱΪ�ò��˰����Ժ����
    gstrSQL = "Select ������,��Ա��� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���㷽ʽ", TYPE_������, lng����ID)
    If Right(rsTemp!������, 1) <> 4 And Not (rsTemp!��Ա��� = "��������" Or rsTemp!��Ա��� = "ʡ������") And ҽ�������Ѿ���Ժ(lng����ID) = False Then
        MsgBox "��Ϊ�òα���Ա�����Ժ������"
    End If
    
    סԺ����_���� = True
    
    '����ҽ����Ժ���������������HIS��Ժͬʱ����ҽ����Ժ�Ļ�������Ҫ�ڽ���ɹ������ҽ����Ժ���������ʧ�ܣ����Ա����ʻ����ٴΰ���ҽ����Ժ��
    If mblnҽ����Ժ = False And ҽ�������Ѿ���Ժ(lng����ID) Then
        Call ��Ժ�Ǽ�_����(lng����ID, lng��ҳID, True)
    End If
    
    '�����Ǽ���Ϣ����
    gstrSQL = "Zl_�����ǼǼ�¼_����_���ʸ���(" & TYPE_������ & "," & lng����ID & "," & lng��ҳID & ",1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    Dim lng����ID As Long, lng����ID As Long, lng��ҳID As Long
    Dim str�������� As String, str��ǰ���� As String
    Dim rsTemp  As New ADODB.Recordset, rsCheck As New ADODB.Recordset
    
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str��Ա��� As String
    Dim str����˳��� As String, str������ As String
    Dim cur�����ʻ� As Currency
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curDate As Date    '�˷�
    Dim bln���� As Boolean
    Dim str֧������ As String
    
    On Error GoTo errHand
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
        " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", TYPE_������, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    lng����ID = rsTemp!����ID: lng��ҳID = Nvl(rsTemp!��ҳID)
    cur�����ʻ� = Nvl(rsTemp!�����ʻ�֧��, 0)
    str������ = IIf(IsNull(rsTemp!֧��˳���), "", rsTemp!֧��˳���)
    str����˳��� = IIf(IsNull(rsTemp!��ע), "", rsTemp!��ע)
    str֧������ = Nvl(rsTemp!ҽ�����)
    If str֧������ = "" Then
        'Modify By �̳ظ� 2010-01-16 �ӱ����ʻ�����ȡ֧������
        gstrSQL = "Select ֧������ From �����ʻ� Where ����=[1] And ����ID=[2]"
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ��ʽ", TYPE_������, lng����ID)
        str֧������ = Nvl(rsCheck!֧������, "31")      ' ֧����� 31��סԺ��37��תԺ
    End If
'
'    '�ж��Ƿ�Ϊ������Ա
'    gstrSQL = "Select ��Ա��� From �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_������
'    Call OpenRecordset(rsCheck, "�ж��Ƿ�Ϊ������Ա")
'    If Not (rsCheck!��Ա��� = "ʡ������" Or rsCheck!��Ա��� = "��������") Then
'        MsgBox "����ҽ�Ʋ����Ľ��ʼ�¼�����������", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    '�Ǳ��½��ʵĵ��ݣ����������
    gstrSQL = "select to_char(�շ�ʱ��,'yyyy-MM-dd') ����ʱ�� From ���˽��ʼ�¼ Where ID=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng����ID)
    str�������� = Format(rsCheck!����ʱ��, "yyyyMM")
    str��ǰ���� = Format(zlDatabase.Currentdate, "yyyyMM")
    If str��ǰ���� <> str�������� Then
        Err.Raise 9000, gstrSysName, "ֻ�ܳ������µĽ��ʵ��ݣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '----׼����������----
    '��ȡҽ�����˵Ļ�����Ϣ
    gstrSQL = "Select ����,ҽ����,˳��� ����,��Ա���,���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����˵Ļ�����Ϣ", TYPE_������, lng����ID)
    str���� = rsCheck!����
    strҽ���� = rsCheck!ҽ����
    str��Ա��� = rsCheck!��Ա���
    str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21", _
                    str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", _
                    str��Ա��� = "��ͨ����", "41", str��Ա��� = "�ͱ�����", "42", _
                    str��Ա��� = "������Ա", "43", str��Ա��� = "���ռ�ͥ", "44", _
                    str��Ա��� = "�ضȲм�", "45", True, "11") '����������ʾ�������ͥ�������ݿ�ֻ��8λ��ֻ�ܱ���Ϊ���ռ�ͥ
    str���� = IIf(IsNull(rsCheck!����), "", rsCheck!����)
    bln���� = (str��Ա��� = "32" Or str��Ա��� = "34")
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", str����˳���)            ' ����˳���
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str������)           ' ������
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str֧������)           ' ֧������
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)           ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) ' ��������
    
    '���ýӿ�
    If CommServer("RETBALANCE", IIf(bln����, 1, 0)) = False Then Exit Function
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, rsTemp("����ID"), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_������ & "," & rsTemp("����ID") & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") & "," & _
        -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & Nvl(rsTemp!֧��˳���) & "',null,null,'" & Nvl(rsTemp!��ע) & "'," & _
        "0,'" & AnalyseComputer & "','" & gstrVersion & "','" & str֧������ & "','" & Nvl(rsTemp!������ˮ��) & "'," & _
        "NULL,'" & Nvl(rsTemp!��������) & "','" & Nvl(rsTemp!����֢) & "',to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'))"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "Select * From ���㸽����Ϣ Where ����ID=" & lng����ID
    Call OpenRecordset_OtherBase(rsTemp, "��ȡ���㸽�Ӽ�¼", gstrSQL, gcnGYYB)
    If rsTemp.RecordCount <> 0 Then
        gstrSQL = "zl_���㸽����Ϣ_Insert (" & lng����ID & "," & -1 * Nvl(rsTemp!����Ա�����𸶱�׼, 0) & "," & -1 * Nvl(rsTemp!����Ա��������, 0) & "," & -1 * Nvl(rsTemp!��ͨ���﹫��Ա�����ۼ�, 0) & "," _
            & -1 * Nvl(rsTemp!����Ա����, 0) & "," & -1 * Nvl(rsTemp!������Ա����, 0) & "," & -1 * Nvl(rsTemp!ҽ���չ���Ա�����Ա����, 0) & "," & rsTemp!������־ & "," & _
            "'" & Nvl(rsTemp!�����ֱ���_����) & "'," & Nvl(rsTemp!���㷽ʽ, 0) & ",'" & Nvl(rsTemp!������) & "'," & Nvl(rsTemp!���㷽ʽ, 1) & "," & -1 * Nvl(rsTemp!�����ܷ���, 0) & "," & -1 * Nvl(rsTemp!ҽ���ܷ���, 0) & ",'" & Nvl(rsTemp!�����϶����) & "'," & Nvl(rsTemp!���˿���סԺ, 0) & ")"
        gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    End If
    '�����Ǽ���Ϣ����
    gstrSQL = "Zl_�����ǼǼ�¼_����_���ʸ���(" & TYPE_������ & "," & lng����ID & "," & lng��ҳID & ",0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    סԺ�������_���� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Sub ��ѯǷ�ѵ�λ_����(ByVal str��λ���� As String, ByVal str������� As String)
'���ܣ����ýӿڲ�ѯǷ�ѵ�λ
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str��ʾ As String
    
    If str��λ���� = "" Then Exit Sub
'    str��λ���� = String(12 - Len(str��λ����), "0") & str��λ����
    
    On Error GoTo errHandle
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "DEPTCODE", str��λ����)                '��λ����
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", str�������)              '�������
    
    '���ýӿ�
    If CommServer("QUERYARREARDEPT") = False Then Exit Sub
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then
        MsgBox "���˵�λ��Ƿ�������", vbInformation, gstrSysName
        Exit Sub
    End If
    '���ݱ���õ���������
    For Each nodRow In nodRowset.childNodes
        Select Case GetAttributeValue(nodRow, "INSUREKIND")
            Case "3"
                str��ʾ = str��ʾ & "������ҽ��"
            Case "8"
                str��ʾ = str��ʾ & "�����ҽ��"
            Case "10"
                str��ʾ = str��ʾ & "������Ա����"
        End Select
    Next
    
    If str��ʾ <> "" Then
        MsgBox "���˵�λ����������Ƿ�������" & Mid(str��ʾ, 2) & "��", vbInformation, gstrSysName
    Else
        MsgBox "���˵�λ��Ƿ�������", vbInformation, gstrSysName
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function ������Ϣ_����(ByVal lngErrCode As Long) As String
'���ܣ����ݴ���ŷ��ش�����Ϣ

End Function

Public Function ҽ����Ŀ_����(rsTemp As ADODB.Recordset, Optional ByVal str��� As String = "12") As Boolean
'���ܣ�ҽ������ҩƷĿ¼��ѯ
'��ǰ�����Ĳ�ѯ���ָ�Ϊ����Ŀ֧������ѯ��41-�������� 42-����סԺ 21-�������� 22-����סԺ 11-��ͨ���� 12-��ͨסԺ 31-�������� 32-����סԺ��
'����ͨסԺ������Ŀ�嵥�����գ�����ģʽ����ǰһ����ֻ���ṩ����ѯ�Ľ��棬�ɰ��û�Ҫ���ѯĳ������µ���Ŀ��֧����������Ϣ
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str���� As String, str���� As String, str����, str��ע As String
    Dim str��ʼ���� As String, str�������� As String, str��ǰ���� As String
        
    On Error GoTo errHandle
    
    If ҽ����ʼ��_���� = False Then Exit Function
    
    str��ǰ���� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "ITEMCODE", "")         ' ҽ������
    Call InsertChild(mdomInput.documentElement, "ITEMPAYTYPE", str���) ' ��Ŀ֧�����
    
    '���ýӿ�
    If CommServer("QUERYSERVICE") = False Then Exit Function
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then Exit Function
    For Each nodRow In nodRowset.childNodes
        str���� = GetAttributeValue(nodRow, "ITEMCODE")
        str���� = ToVarchar(Replace(GetAttributeValue(nodRow, "ITEMNAME"), "'", ""), 40)
        str���� = ToVarchar(zlCommFun.SpellCode(str����), 10)
        str��ʼ���� = Mid(GetAttributeValue(nodRow, "STARTDATE"), 1, 10)
        str�������� = Mid(GetAttributeValue(nodRow, "ENDDATE"), 1, 10)
'        PRICELMT           '����޼�
'        SELFRATE           '�Ը�����
'        BEARINGITEMFLAG    '������Ŀ��־
'        GSITEMFLAG         '������Ŀ��־
'        SPECPAYFLAG        '���ⱨ����Ŀ��־
'        BGITEMTYPE         '���ɽ�����Ŀ���
        str��ע = Val(GetAttributeValue(nodRow, "PRICELMT")) & "|" & Val(GetAttributeValue(nodRow, "SELFRATE")) & "|" & _
                  Val(GetAttributeValue(nodRow, "BEARINGITEMFLAG")) & "|" & Val(GetAttributeValue(nodRow, "GSITEMFLAG")) & "|" & _
                  Val(GetAttributeValue(nodRow, "SPECPAYFLAG")) & "|" & Val(GetAttributeValue(nodRow, "BGITEMTYPE"))
        
        If str���� <> "" And str��ǰ���� >= str��ʼ���� And str��ǰ���� <= str�������� Then
            rsTemp.AddNew Array("CLASSCODE", "CODE", "NAME", "PY", "MEMO"), Array("1", str����, str����, str����, str��ע)
            rsTemp.Update
        End If
    Next
    
    ҽ����Ŀ_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Public Function InitXML() As Boolean
'���ܣ���ʼ��XML�����������͸��ڵ�
    Dim pi As MSXML2.IXMLDOMProcessingInstruction
    Dim nodData As MSXML2.IXMLDOMElement
    
    On Error Resume Next
    
    Set mdomInput = New MSXML2.DOMDocument
    Set mdomOutput = New MSXML2.DOMDocument
    If Err <> 0 Then
        Err.Clear
        Exit Function
    End If
    
'    'XML����
'    Set pi = mdomInput.createProcessingInstruction("xml", "version=""1.0"" encoding=""GB2312"" standalone=""yes""")
'    mdomInput.appendChild pi
    
    '���ڵ�
    Set nodData = mdomInput.createElement("DATA")
    Set mdomInput.documentElement = nodData
    
    InitXML = True
End Function

Public Function InsertChild(nodParent As MSXML2.IXMLDOMElement, ByVal Name As String, ByVal Value As String) As MSXML2.IXMLDOMElement
'���ܣ���ָ��XMLԪ����������Ԫ��
    Set InsertChild = mdomInput.createElement(Name)
    InsertChild.Text = Value
    
    nodParent.appendChild InsertChild
End Function

Public Sub InsertAttrib(nodParent As MSXML2.IXMLDOMElement, ByVal Name As String, ByVal Value As String)
'���ܣ���ָ��XMLԪ������������
    Dim attTemp As MSXML2.IXMLDOMAttribute
    
    Set attTemp = mdomInput.createAttribute(Name)
    attTemp.Text = Value
    
    nodParent.setAttributeNode attTemp
End Sub

Public Function CommRecServer(ByVal strFunction As String) As Boolean
'���ܣ�����ҽ������
    Dim InvokeServer As String '����ǰ�÷������ķ���ֵ
    Dim StrInput As String
    
    '�����Ĵ���
    StrInput = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?>" & vbCrLf & mdomInput.xml
    Call DebugTool(StrInput)
    
    Select Case strFunction
        Case "APPRECM"
            InvokeServer = obj����.APPRECM("ZFRJ", StrInput)
        Case "DELRECM"
            InvokeServer = obj����.DELRECM("ZFRJ", StrInput)
        Case "APPRECB"
            InvokeServer = obj����.APPRECB("ZFRJ", StrInput)
        Case "DELRECB"
            InvokeServer = obj����.DELRECB("ZFRJ", StrInput)
        Case "APPRECG"
            InvokeServer = obj����.APPRECG("ZFRJ", StrInput)
        Case "DELRECG"
            InvokeServer = obj����.DELRECG("ZFRJ", StrInput)
        Case "QUERYREC"
            InvokeServer = obj����.QUERYREC("ZFRJ", StrInput)
        Case Else
            ShowMsgbox "����ҽ���ӿڷ����仯���޷�����ִ�н��ף���������ṩ����ϵ��"
            Exit Function
    End Select
    
    '�ϵ����ô�
    If InvokeServer = "" Then
        '����ʧ�ܣ����ع̶��Ĵ�����Ϣ
        InvokeServer = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?><DATA><RETCODE>-1</RETCODE><INFO>ҽ������������ʧ��</INFO></DATA>"
    End If
            
    If mdomOutput.loadXML(InvokeServer) = False Then
        ShowMsgbox "ҽ������������ֵ��ʽ����ȷ��"
    Else
        '�ٶ����������Ƿ�ɹ����з���
        If Val(GetElemnetValue("RETCODE")) = 0 Then
            '���óɹ�
            CommRecServer = True
        Else
            '����ʧ��
            InvokeServer = GetElemnetValue("INFO")
            If InvokeServer = "" Then InvokeServer = "����������ʧ�ܡ�"
            ShowMsgbox "ҽ�����������ش���" & vbCrLf & vbCrLf & InvokeServer
        End If
    End If
End Function

Public Function CommServer(ByVal strFunction As String, Optional ByVal strAdvance As String = "") As Boolean
'���ܣ�����ҽ������
'strOutPut:ҽ���ӿڷ�����Ϣ  2011-05-16�̳ظ�����
    Dim InvokeServer        As String '����ǰ�÷������ķ���ֵ
    Dim StrInput            As String
    Dim strDetailLog        As String
    On Error GoTo errHand
    '�����Ĵ���
    StrInput = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?>" & vbCrLf & mdomInput.xml
    
    strDetailLog = ""
    strDetailLog = strDetailLog & vbCrLf & "����Ա��" & UserInfo.����
    strDetailLog = strDetailLog & vbCrLf & "����վ��" & AnalyseComputer
    
    Select Case strFunction
        Case "GETPSNINFO"
            InvokeServer = gobj����.GETPSNINFO("ZFRJ", StrInput)
        Case "GETGSINFO"
            InvokeServer = gobj����.GETGSINFO("ZFRJ", StrInput)
        Case "MODIFYCARD"               '�޸Ŀ�����
            InvokeServer = gobj����.MODIFYCARD("ZFRJ", StrInput)
        Case "GETCLINNO"                '����Һ�
            
            InvokeServer = gobj����.GETCLINNO("ZFRJ", StrInput)
        Case "CALCLIN"                  '��ͨ����֧��
            strDetailLog = strDetailLog & vbCrLf & "֧����ʽ����ͨ����֧��"
            strDetailLog = strDetailLog & vbCrLf & "���ܣ�GETCLINNO"
            strDetailLog = strDetailLog & vbCrLf & "��Σ�" & vbCrLf & StrInput
            Call DebugTool(StrInput)
            InvokeServer = gobj����.CALCLIN("ZFRJ", StrInput)
            strDetailLog = strDetailLog & vbCrLf & "���Σ�" & vbCrLf & InvokeServer
            
        Case "CALSPECCLIN"              '��������֧��
            strDetailLog = strDetailLog & vbCrLf & "֧����ʽ����������֧��"
            strDetailLog = strDetailLog & vbCrLf & "���ܣ�CALSPECCLIN"
            strDetailLog = strDetailLog & vbCrLf & "��Σ�" & vbCrLf & StrInput
            Call DebugTool(StrInput)
            InvokeServer = gobj����.CALSPECCLIN("ZFRJ", StrInput)
            strDetailLog = strDetailLog & vbCrLf & "���Σ�" & vbCrLf & InvokeServer
            
        Case "RETBALANCE"               '��Ʊ
            If strAdvance = "1" Then    '������Ʊ
                InvokeServer = gobj����.RETLX("ZFRJ", StrInput)
            Else
                InvokeServer = gobj����.RETBALANCE("ZFRJ", StrInput)
            End If
        Case "HOSPREG"                  'סԺ�Ǽ�
            InvokeServer = gobj����.HOSPREG("ZFRJ", StrInput)
        Case "HOSPOUT"                  '��Ժ�Ǽ�
            InvokeServer = gobj����.HOSPOUT("ZFRJ", StrInput)
        Case "CALHOSP"                  'סԺ֧��
            If strAdvance = "1" Then    '�޿����㣬�������ڲ����ӷѵ����
                InvokeServer = gobj����.CALHOSPSP("ZFRJ", StrInput)
            Else
                InvokeServer = gobj����.CALHOSP("ZFRJ", StrInput)
            End If
        Case "SETRECKONINGTYPE"         '�������㷽ʽ
            InvokeServer = gobj����.SETRECKONINGTYPE("ZFRJ", StrInput)
        Case "QUERYHOSPSINGLEILLNESS"   '��������������
            InvokeServer = gobj����.QUERYHOSPSINGLEILLNESS("ZFRJ", StrInput)
        Case "QUERYHOSPSINGLEILLNESS_BG"   '�����ֽ���Ŀ¼
            InvokeServer = gobj����.QUERYHOSPSINGLEILLNESS_BG("ZFRJ", StrInput)
        Case "QUERYSERVICE"              'ҽ������ҩƷĿ¼��ѯ
            InvokeServer = gobj����.QUERYSERVICE("ZFRJ", StrInput)
        Case "QUERYARREARDEPT"          '��ѯǷ�ѵ�λ
            InvokeServer = gobj����.QUERYARREARDEPT("ZFRJ", StrInput)
        Case "GETHOSPSINGLEILLNESS"     '���ص�������������
            InvokeServer = gobj����.GETHOSPSINGLEILLNESS("ZFRJ", StrInput)
        Case "GETHOSPSINGLEILLNESS_BG"  '���ص����ֽ���Ŀ¼
            InvokeServer = gobj����.GETHOSPSINGLEILLNESS_BG("ZFRJ", StrInput)
        Case "SETBEARINGFLAG"           '����������־
            InvokeServer = gobj����.SETBEARINGFLAG("ZFRJ", StrInput)
        Case "UPLOADICD"                '�ϴ�ICD����
            InvokeServer = gobj����.UPLOADICD("ZFRJ", StrInput)
        Case "SETCALTYPE"
            InvokeServer = gobj����.SETCALTYPE("ZFRJ", StrInput)
        Case "RETHOSPOUT"
            InvokeServer = gobj����.RETHOSPOUT("ZFRJ", StrInput)
        Case "GETSPECILLNESS"
            InvokeServer = gobj����.GETSPECILLNESS("ZFRJ", StrInput)
        Case "QUERYSPECILLNESS"
            InvokeServer = gobj����.QUERYSPECILLNESS("ZFRJ", StrInput)
        Case "UPLOADBYBILLNO"
            InvokeServer = gobj����.UPLOADBYBILLNO("ZFRJ", StrInput)
        Case Else
            ShowMsgbox "����ҽ���ӿڷ����仯���޷�����ִ�н��ף���������ṩ����ϵ��"
            Exit Function
    End Select
    
    '�ϵ����ô�
    If InvokeServer = "" Then
        '����ʧ�ܣ����ع̶��Ĵ�����Ϣ
        InvokeServer = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?><DATA><RETCODE>-1</RETCODE><INFO>ҽ������������ʧ��</INFO></DATA>"
    End If
            
    If mdomOutput.loadXML(InvokeServer) = False Then
        ShowMsgbox "ҽ������������ֵ��ʽ����ȷ��"
    Else
        '�ٶ����������Ƿ�ɹ����з���
        If Val(GetElemnetValue("RETCODE")) = 0 Then
            '���óɹ�
            CommServer = True
        Else
            '����ʧ��
            InvokeServer = GetElemnetValue("INFO")
            If InvokeServer = "" Then InvokeServer = "����������ʧ�ܡ�"
            ShowMsgbox "ҽ�����������ش���" & vbCrLf & vbCrLf & InvokeServer
            
        End If
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function Get���ղ���_����(ByVal str������ As String) As String
'���ܣ���ñ��ղ���
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.������,A.����ֵ from ���ղ��� A " & _
              " where A.������=[1] and A.����=[2] and A.���� is null "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", str������, TYPE_������)
    
    If rsTemp.EOF = False Then
        Get���ղ���_���� = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
    End If
End Function

Public Function SetElemnetValue(ByVal Name As String, ByVal Value As String) As Boolean
'���ܣ��õ�ָ��Ԫ�ص�ֵ
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set xmlElement = mdomInput.documentElement.selectSingleNode(Name)
    If Not xmlElement Is Nothing Then
        '�ҵ�ָ����Ԫ��
        xmlElement.nodeTypedValue = Value
        SetElemnetValue = True
    End If
End Function

Public Function GetElemnetValue(ByVal Name As String) As String
'���ܣ��õ�ָ��Ԫ�ص�ֵ
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set xmlElement = mdomOutput.documentElement.selectSingleNode(Name)
    If Not xmlElement Is Nothing Then
        '�ҵ�ָ����Ԫ��
        GetElemnetValue = xmlElement.Text
'    Else
'        'ȡ��
'        Debug.Assert False
    End If
End Function

Public Function GetAttributeValue(xmlElement As MSXML2.IXMLDOMElement, ByVal Name As String) As String
'���ܣ��õ�ָ�����Ե�ֵ
    Dim varAttribute As Variant
    
    varAttribute = xmlElement.getAttribute(Name)
    If IsNull(varAttribute) = False Then
        GetAttributeValue = varAttribute
    End If
End Function

Public Function Get��֤_����(bytType As Byte, str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, _
                ByVal lng����ID As Long, Optional blnǿ��ˢ�� As Boolean = False) As Boolean
'���ܣ��õ�ҽ�����˵Ļ������������֤��Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    '�����ݿ��ж����Ѵ洢��ֵ
    gstrSQL = " select ����,ҽ����,˳���,����,NVL(�����,0) AS �����,IDNO,PSAM " & _
              " from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID, TYPE_������)
    If rsTemp.EOF = False Then
        str���� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        str���� = Replace(str����, ",", ";")
        strҽ���� = IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����"))
        str�����ı�� = mstrҽ�����ı���_����
        gintType = rsTemp!�����
        gstrIDNO = Nvl(rsTemp!IDNO)
        gstrPSAMNO = Nvl(rsTemp!PSAM)
    End If
    If blnǿ��ˢ�� = False And lng����ID > 0 Then
        Get��֤_���� = True
        Exit Function
    End If
    
    If frmIdentify����.GetIdentify(bytType, str����, strҽ����, str�����ı��, str����) = False Then
        Exit Function
    Else
        'ˢ����Ȼ��ȷ����Ҫ����Ƿ���ǵ�ǰ���˵�
        str���� = Split(str����, "^")(0)
'        If lng����ID > 0 Then
'            '�����ݿ��ж����Ѵ洢��ֵ
'            gstrSQL = "select ����,ҽ����,˳��� from �����ʻ� where ����ID=" & lng����ID & " and ����=" & TYPE_������
'            Call OpenRecordset(rsTemp, "����ҽ��")
'
'            If str���� <> Replace(rsTemp("����"), ",", ";") Or strҽ���� <> IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����")) Then
'                MsgBox "��ǰʹ�õĿ��벡�˲�����", vbInformation, gstrSysName
'                Exit Function
'            End If
'        End If
    
    End If
    
    Get��֤_���� = True
End Function

Public Function OwnerUser(ByVal strUserName As String) As Boolean
    Dim RecUser As New ADODB.Recordset
    '�жϵ�ǰ�û��ǲ���������
    OwnerUser = True
    With RecUser
        If .State = 1 Then .Close
        .Open "Select Count(*) ������ From ZlSystems Where ������='" & strUserName & "'", gcnOracle
        
        If Not .EOF Then
            If Not IsNull(!������) Then
                If !������ = 0 Then OwnerUser = False
            End If
        End If
    End With
End Function

Public Function Subject(ByVal strData As String) As String
    Dim rsSubject As New ADODB.Recordset
    '���ض�Ӧ�Ĺ�����Ŀ����
    gstrSQL = "" & _
             " Select B.����,B.���,A.����ֵ ������Ŀ����   " & _
             " From ���ղ��� A,�շ���� B " & _
             " Where A.���>=6 And A.����=[1] And A.������=B.���� And B.����=[2]"
    Set rsSubject = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ӧ�Ĺ�����Ŀ����", TYPE_������, strData)
    
    If rsSubject.EOF Then
        Subject = "11"  '�޶�Ӧ��Ŀ���ض�Ӧ�Ĺ�����Ŀ����'11',��ʾ����
    Else
        Subject = rsSubject!������Ŀ����
    End If
End Function

Public Function ����Һ�_����(ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    
    Dim datCurr As Date
    Dim str���㷽ʽ As String, arr���㷽ʽ
    Dim intTotal  As Integer, intStart As Integer
    Dim cur�����ʻ� As Currency
    Dim curҽ������ As Currency, cur���ͳ�� As Currency, cur����Ա���� As Currency, cur������ As Currency
    Dim str���� As String, strҽ���� As String, str�����ı�� As String, str���� As String, str����˳��� As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    datCurr = zlDatabase.Currentdate()
    gstrSQL = " Select ����ID From " & IIf(gblnHIS1026 = True, "������ü�¼", "���˷��ü�¼") & " Where ����ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����ID)
    lng����ID = rsTemp!����ID
    'If Get��֤_����(0, str����, strҽ����, str�����ı��, str����, lng����ID) = False Then Exit Function
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    '��XML����ֵ
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstrҽ����)     ' ���˱���
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) ' ��������
    '���ýӿ�
    If CommServer("GETCLINNO") = False Then Exit Function
    str����˳��� = GetElemnetValue("BILLNO")
    
    gstrSQL = "Select ����ID,�շ�ϸĿID,����*NVL(����,1) AS ����,��׼���� AS ����,Nvl(ʵ�ս��,0) AS ʵ�ս��,���ձ���,'  ' AS ժҪ" & _
        " From " & IIf(gblnHIS1026 = True, "������ü�¼", "���˷��ü�¼") & _
        " Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    If Not �����������_����(rsTemp, str���㷽ʽ, "") Then Exit Function
    
    '�ֽ���ֽ��㷽ʽ
    arr���㷽ʽ = Split(str���㷽ʽ, "|")
    intTotal = UBound(arr���㷽ʽ)
    For intStart = 0 To intTotal
        Select Case Split(arr���㷽ʽ(intStart), ";")(0)
        Case "�����ʻ�"
            cur�����ʻ� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "ҽ������"
            curҽ������ = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "��ͳ��"
            cur���ͳ�� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "ҽ�Ʋ���"
            cur����Ա���� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "������"
            cur������ = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        End Select
    Next
    
    If Not �������_����(lng����ID, cur�����ʻ�, "", True, "1|1") Then Exit Function
    
   '��Ҫ����������
    str���㷽ʽ = ""
    If cur�����ʻ� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
    If curҽ������ <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curҽ������
    If cur���ͳ�� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||��ͳ��|" & cur���ͳ��
    If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ�Ʋ���|" & cur����Ա����
    If cur������ <> 0 Then str���㷽ʽ = str���㷽ʽ & "||������|" & cur������ & ";0"
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
        gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
    End If
    
    ����Һ�_���� = True
    mlng����ID = lng����ID
    
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ���ý��㷽ʽ_����(ByVal lng����ID As Long, ByVal frmParent As Object, Optional ByVal blnסԺ As Boolean = False) As String
    Dim lng��ҳID As Long
    Dim rsTemp As New ADODB.Recordset
    '���ؽ��㷽ʽ�뵥���ֱ���
    
    '���˱��ղ��������ý��㷽ʽ
    gstrSQL = " Select A.�������,B.סԺ���� From �����ʻ� A,������Ϣ B" & _
              " Where A.����ID=B.����ID And A.����ID=[1] And A.����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", lng����ID, TYPE_������)
    If rsTemp!������� = "7" Then Exit Function
    If blnסԺ Then lng��ҳID = Nvl(rsTemp!סԺ����, 0)
    
    ���ý��㷽ʽ_���� = frm���ý��㷽ʽ.ShowSelect(lng����ID, lng��ҳID, TYPE_������, frmParent)
End Function

Public Function �������㷽ʽ_����(ByVal lng����ID As Long, ByVal frmParent As Object, Optional ByVal blnסԺ As Boolean = False) As Boolean
    Dim lng��ҳID As Long
    Dim rsTemp As New ADODB.Recordset
    
    '���˱��ղ������������㷽ʽ
    gstrSQL = " Select A.�������,B.סԺ���� From �����ʻ� A,������Ϣ B" & _
              " Where A.����ID=B.����ID And A.����ID=[1] And A.����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", lng����ID, TYPE_������)
    If rsTemp!������� = "7" Then Exit Function
    If blnסԺ Then lng��ҳID = Nvl(rsTemp!סԺ����, 0)
    
    �������㷽ʽ_���� = frm�������㷽ʽ.ShowSelect(lng����ID, lng��ҳID, TYPE_������, frmParent)
End Function

Public Sub ����ѡ��_����(ByVal lng����ID As Long)
    Dim lng����ID As Long
    Dim str���� As String
    Dim rs���� As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    '��ȡ�ò�����ǰ�Ĳ�����Ϣ
    gstrSQL = " select B.����,B.���� from �����ʻ� A,���ղ��� B " & _
              " where A.����ID=[1] and A.����=[2] and A.����ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺԤ��", lng����ID, TYPE_������)
    If rsTemp.RecordCount <> 0 Then
        str���� = "[" & rsTemp!���� & "]" & rsTemp!����
    End If
    
    '��סԺ����ѡ����
    gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_������
    Set rs���� = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "ҽ�����֡���" & IIf(str���� = "", "��", str����))
    If Not rs���� Is Nothing Then
        lng����ID = rs����("ID")
    End If
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����ID','''" & lng����ID & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���没��")
End Sub

Public Function ����ICD����_����(ByVal lng����ID As Long) As Boolean
    Dim strICD As String
    Dim rsTemp As New ADODB.Recordset
'    <BILLNO>����˳���</BILLNO>
'    <ICD>ICD����</ICD>
'    <DODATE>����ʱ��</DODATE>
    
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "��ҽ�����˻�δ��Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ѡ��ICD����
    strICD = frm����ѡ��_����.ChooseDisease(lng����ID)
    If strICD = "" Then Exit Function
    
    '�ϴ����˵�ICD����
    gstrSQL = "Select ҽ����,˳��� From �����ʻ� Where ����=[1] ANd ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ò��˵�ҽ����", TYPE_������, lng����ID)
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "BILLNO", Nvl(rsTemp!˳���))   '˳���
    Call InsertChild(mdomInput.documentElement, "ICD", strICD)                  '����
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) '��������
    If CommServer("UPLOADICD") = False Then Exit Function
    
    ����ICD����_���� = True
End Function

Public Function GetItemInsure_����(lng����ID As Long, lng�շ�ϸĿID As Long, bln���� As Boolean) As String
    'ҽ����������в���һ����¼
    'insert into ҽ���������
    '(����,����,����,˵��)
    'Values
    '(50,'1','����','��')
    '����ʷ�������ݲ��뵽ҽ��������ϸ��
    'insert into ҽ��������ϸ
    'select ����,1,�շ�ϸĿID,��Ŀ����,''
    'From ����֧����Ŀ
    'Where ���� = 50
    Dim strDefault As String            'ȱʡҽ������
    Dim strCurrent As String            '��ǰҽ�����룬����ȡ������룬סԺȡסԺ����
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "Select B.���,A.����,A.����,B.˵�� From ������Ŀ A,ҽ��������ϸ B" & _
        " Where B.����=[1] And A.����=B.���� And A.����=B.��Ŀ���� And B.�շ�ϸĿID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ������", TYPE_������, lng�շ�ϸĿID)
    rsTemp.Filter = "���=1"
    Select Case rsTemp.RecordCount
    Case 0
        'û�����ö�Ӧ���룬ȡȱʡ����
        rsTemp.Filter = "���=0"
        If rsTemp.RecordCount <> 0 Then
            GetItemInsure_���� = rsTemp!����
        End If
    Case 1
        GetItemInsure_���� = rsTemp!����
    Case Else
        '��ѡ
        GetItemInsure_���� = frmҽ����Ŀѡ��.ShowSelect(rsTemp, lng�շ�ϸĿID)
    End Select
    
    rsTemp.Filter = 0
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    rsTemp.Filter = 0
End Function

Public Function IS����(ByVal lng����ID As Long) As Boolean
    Dim str��Ա��� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = " Select ��Ա��� From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ա���", lng����ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    str��Ա��� = Nvl(rsTemp!��Ա���)
    str��Ա��� = Switch(str��Ա��� = "��ְ", "11", str��Ա��� = "����", "21" _
                  , str��Ա��� = "ʡ������", "32", str��Ա��� = "��������", "34", True, "11")
    IS���� = (str��Ա��� = "32" Or str��Ա��� = "34")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetNextID(ByVal strTable As String, ByVal cnCustom As ADODB.Connection) As Long
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select " & strTable & "_ID.Nextval From Dual"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, cnCustom
    GetNextID = rsTemp.Fields(0).Value
End Function

Public Sub ��������ҩƷ��ʾ(ByVal lngItemID As Long)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
'20100607����ǿ�ڴ˴������˷�������Ϊ��ҩƷȫ�Էѡ�����Ŀʱ��Ҳ��ʾ
    If Not mbln����ҩƷ��ʾ Then Exit Sub
    gstrSQL = " Select ���,����,˵��,nvl(��������,'ҩƷȫ�Է�') as  �������� From �շ�ϸĿ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ˵��,���������ֶ�", lngItemID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    If InStr(1, "5,6,7", rsTemp!���) = 0 Then Exit Sub
    
    If Nvl(rsTemp!˵��) = "" And rsTemp!�������� <> "ҩƷȫ�Է�" Then Exit Sub
    If Nvl(rsTemp!˵��) <> "" Then
    MsgBox rsTemp!���� + rsTemp!˵��, vbInformation, gstrSysName
    End If
    If rsTemp!�������� = "ҩƷȫ�Է�" Then
    MsgBox rsTemp!���� + rsTemp!��������, vbInformation, gstrSysName
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SaveBalanceLog(ByVal int���� As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal str����˳��� As String, ByVal str������ As String, ByVal str֧����� As String)
    Dim cnLog As New ADODB.Connection
    Set cnLog = GetNewConnection
    With cnLog
        
        gstrSQL = "ZL_������־_����_INSERT(" & lng����ID & "," & lng����ID & ",'" & str����˳��� & "','" & str������ & "'," & _
                "'" & str֧����� & "','" & gstrUserName & "','" & AnalyseComputer & "',1)"
        .Execute gstrSQL, , adCmdStoredProc
        .Close
    End With
End Sub
 
Private Function logID() As String
    gstrSQL = "select decode(max(��־����),null,to_char(sysdate,'yyyymmdd') || '000001',to_char(sysdate,'yyyymmdd') ||lpad(to_number(substr(max(��־����),9,6))+1,6,'0')) from ҽ��������־ where ��־���� like to_char(sysdate,'yyyymmdd') || '%'"
    logID = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��־����").Fields(0)
End Function
'========================================================================================
'=��¼�޸���־(ģ��\����\����(�޸�|ɾ��|����)\�޸�����(�ַ���)\�޸��ļ���־\����1\����2\����3)
'========================================================================================
'=����:
'=     vstrBM     ģ��(40)
'=     vstrGN     ģ��(40)
'=     vstrType   ģ��(1)
'=     vstrTxt    �ַ�����ʽ����־����
'=     vstrFile   ��  ����ʽ����־����
'=     vstrKey1   ����1
'=     vstrKey2   ����2
'=     vstrKey3   ����3
'=     vstrKey4   ����4
'=     vstrSource ���ݱ���,�����|�ֿ�,����:DEF_CUSTID_M|DEF_CUSTID_M
'=     vblnKillFile �Ƿ�ɾ����־�ļ�(True��ɾ��,False��ɾ��)
'========================================================================================
'=ע��:
'=     1.vstrBM   ��ģ��,vstrGN ����,��:vstrBM=Ӫ��ϵͳ,vstrGN=�ͻ�����
'=     2.vstrType ��־����:1=�޸�,2=ɾ��,3=����
'=     3.vstrTxt,vstrFileֻ��ѡ�����е�һ������,����ļ������ܿ�,vstrFile����
'=     4.vstrKey1������¼������ֵ1,vstrKey2������¼������ֵ2,vstrKey3������¼������ֵ3
'========================================================================================
Function AddLog(vStrBM As String, vstrGN As String, vstrType As LogType, _
                Optional vstrTxt As String, Optional vstrFile As String, _
                Optional vstrKey1 As String, Optional vstrKey2 As String, Optional vstrKey3 As String, _
                Optional vstrSource As String, Optional vstrKey4 As String, _
                Optional ByVal vblnKillFile As Boolean = False) As Boolean

    Dim AdoRsMs         As ADODB.Recordset
    Dim strUser         As String
    Dim strDate         As String
    Dim strWS           As String
    Dim strFilePath     As String
    Dim strSQL          As String
On Error GoTo ErrH
    strWS = AnalyseComputer                 'ȡ�ù���վ������
    strUser = UserInfo.����                 'ȡ���û�ID
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")    'ȡ�÷���������
    
    strSQL = "select * from ҽ��������־ where 0=1"
    Set AdoRsMs = New ADODB.Recordset
    With AdoRsMs
        .CursorLocation = adUseClient
        .Properties("Initial Fetch Size") = 50
        .PageSize = 50
        .Open strSQL, gcnGYYB, adOpenKeyset, adLockOptimistic, adAsyncFetch     '���п��޸��α�
    End With
    With AdoRsMs
        .AddNew
        .Fields("��־����") = logID     '�Զ����
        .Fields("ģ��") = vStrBM
        .Fields("����") = vstrGN
        .Fields("����") = vstrType
        .Fields("����") = strDate
        .Fields("����1") = vstrKey1
        .Fields("����2") = vstrKey2
        .Fields("����3") = vstrKey3
        .Fields("����4") = vstrKey4
        .Fields("�û�") = strUser
        .Fields("����վ") = strWS
        .Fields("������Դ") = vstrSource
        If Trim(vstrFile) <> "" And Len(Trim(Dir(vstrFile))) > 0 Then
            .Fields("��־����") = "2"
            .Fields("��־����") = "<���ı�>"
            strFilePath = vstrFile
            Call WriteToDB(AdoRsMs.Fields("��־����"), strFilePath)
            If vblnKillFile Then Kill strFilePath
        Else
            .Fields("��־����") = "1"
            .Fields("��־����") = vstrTxt
        End If
        .Update
    End With
    Set AdoRsMs = Nothing
    Exit Function
ErrH:
    Err.Clear
    Resume Next
    Exit Function
End Function

'========================================================================================
'=��    ��:�ӱ��ж������е���Ϣ��д�뵽�ļ���(�����޸Ĵ���ʱ)
'=��(1) ��:TableName ����,�÷ֺŸ���
'=��(2) ��:TableCondition����,�÷ֺŸ���,�����Ӧ,ͬ���Ŀ���ֻ��һ������
'=��(3) ��:vFile�ļ���,���޸ĺ��룬�������޸�ǰ
'=��(4) ��:vstrTilte������,����������ǽ�Ҫɾ���ļ�¼
'=�� �� ֵ:�����ļ�·��
'========================================================================================
Public Function EditFormerWriteFileA(ByVal TABLEName As String, Optional ByVal TableCondition As String = "", Optional vFile As String = "", Optional vstrTilte As String) As String
    Dim Sql         As String
    Dim rs          As ADODB.Recordset
    Dim PrefixRs    As ADODB.Recordset
    Dim i           As Integer
    Dim j           As Integer
    Dim strFile     As String
    Dim a           As TextStream
    Dim fs          As FileSystemObject
    Dim FileName    As String
    Dim str()       As String
    Dim strW()      As String
    Dim strV        As Variant
On Error GoTo ErrH

    str = Split(TABLEName, ";")
    strW = Split(TableCondition, ";")
    
    If Trim(vFile) = "" Then
        EditFormerWriteFileA = logID
        FileName = App.Path & "\" & EditFormerWriteFileA & ".betry"
        Set fs = New FileSystemObject
        Set a = fs.CreateTextFile(FileName, True)
        If Trim(vstrTilte) <> "" Then
            a.WriteLine vbCrLf & Trim(vstrTilte)
        Else
            a.WriteLine vbCrLf & "�޸�ǰ"
        End If
    Else
        FileName = vFile
        Set fs = New FileSystemObject
        Set a = fs.OpenTextFile(FileName, ForAppending)
        If Trim(vstrTilte) <> "" Then
            a.WriteLine vbCrLf & Trim(vstrTilte)
        Else
            a.WriteLine vbCrLf & "�޸ĺ�"
        End If
    End If
        
    For j = 0 To UBound(str)
        strFile = ""
        Sql = "SELECT Table_Name as TableName,Column_Name as FieldName,Column_Name as FieldNote FROM ALL_TAB_COLUMNS WHERE TABLE_NAME = '" & str(j) & "'"
        Set PrefixRs = zlDatabase.OpenSQLRecord(Sql, "��־")
        With PrefixRs
            While Not .EOF
                strFile = strFile & IIf(Len(Trim(!FieldNote & "")), !FieldNote, !FieldName) & vbTab
                .MoveNext
            Wend
        End With
        Set PrefixRs = Nothing
        a.Write vbCrLf & "������" & str(j) & "��" & vbCrLf
        a.Write strFile
        Sql = "select * from " & str(j)
        If j <= UBound(strW) Then i = j
        If Trim(strW(i)) <> "" Then Sql = Sql & " WHERE " & strW(i)
        Set rs = zlDatabase.OpenSQLRecord(Sql, "��־")
        If rs.RecordCount > 0 Then
            strV = rs.GetString(adClipString, , vbTab, vbCrLf, "")
            a.Write vbCrLf & strV
            Set rs = Nothing
        End If
    Next j
    
    a.Close
    EditFormerWriteFileA = FileName
    Exit Function
ErrH:
    Err.Clear
    Resume Next
    EditFormerWriteFileA = ""
    Exit Function
End Function

'========================================================================================
'=�洢�ļ������ݿ�
'========================================================================================
Private Function WriteToDB(ByRef COL As ADODB.Field, ByVal FileName As String) As Boolean
    Dim mStream As ADODB.Stream
    Dim Lines As String
    Dim NextLine As String
On Error GoTo ErrH
    Open FileName For Input As #1
    Do Until EOF(1)
        Line Input #1, NextLine
        Lines = Lines & NextLine & Chr(13) & Chr(10)
    Loop
    Close #1
    COL.Value = Lines
    Exit Function
ErrH:
    
    Err.Clear
    Exit Function
End Function

Public Function DrugsUsed(ByVal intinsure As Integer, ByVal lng����ID As Long, ByVal lngҩƷID As Long, ByVal dbl�������� As Double) As String
    '���أ�����ʱ���س�����Ϣ�����򷵻ؿմ�
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Zl_ҩƷ�������_����(" & intinsure & "," & lng����ID & "," & lngҩƷID & "," & _
                dbl�������� & "," & gint�ۼ���ҩ�������׼ & "," & IIf(gblnHIS1026 = True, 1, 0) & ") As ������Ϣ From Dual "
    Call OpenRecordset_OtherBase(rsTemp, gstrSQL)
    If Nvl(rsTemp!������Ϣ, "") <> "" Then
        DrugsUsed = Nvl(rsTemp!������Ϣ, "") & vbNewLine
    Else
        DrugsUsed = ""
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    DrugsUsed = ""
End Function

Public Function zl_Ip_Address_FromOrc(Optional strDefaultIp_Address As String = "") As String
    '-----------------------------------------------------------------------------------------------------------
    '����:ͨ��oracle��ȡ�ļ������IP��ַ
    '���:strDefaultIp_Address-ȱʡIP��ַ
    '����:
    '����:����IP��ַ
    '����:���˺�
    '����:2009-01-21 11:08:47
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strIp_Address As String, strSQL As String
    Err = 0: On Error GoTo errHand:
     strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡIP��ַ")
    If rsTemp.EOF = False Then
        strIp_Address = zlCommFun.Nvl(rsTemp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = strDefaultIp_Address
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    zl_Ip_Address_FromOrc = strIp_Address
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

'����˵�����Խ��׽���ȷ�ϻ�ȡ��
'strBusiness���������ƣ���Ӧ��ö�ٱ��� ����Enum
'blnResult��TRUE��ʾ�ύ�ɹ���FALSE��ʾ�����쳣����Ҫȡ��ҽ������

'ҽ���ӿ������ӷ�����BusinessAffirm������ȷ�ϻ�ȡ��ĳ�����ף��������̣�
'1 ?HIS����
'2 ?ҽ�����׳ɹ�
'3 ?HIS�ύ
'4 ?����BusinessAffirmȷ��ҽ������
'
'��Ҫ���Ǵӵ�3����ʼ����������쳣���͵���ȷ�Ͻ��ף�����FALSE��ʾ��Ҫȡ��ҽ������
'��3����ǰ�����κ��쳣����HIS���Ƿ�Χ��?
'
'�����޸Ľ�Ҫ��������㡢����������ϡ�סԺ���㡢סԺ�����������ĸ����״���HIS������Ӧ�޸ġ�

Public Sub BusinessAffirm_������(ByVal intBusiness As Integer, ByVal blnResult As Boolean, Optional ByVal intinsure As Integer = 0, _
    Optional ByVal strAdvance As String)
    '�����׺����ʾ��Ϣ
    If blnResult Then
        '���׳ɹ�
        Select Case intBusiness
            Case ����Enum.Busi_RegistSwap '����Һ�
                Call frm������Ϣ.ShowME(mlng����ID)
            Case ����Enum.Busi_RegistDelSwap '����Һų���
            Case ����Enum.Busi_ClinicSwap '�������
            Case ����Enum.Busi_ClinicDelSwap '����������
            Case ����Enum.Busi_ComeInSwap '��Ժ�Ǽ�
            Case ����Enum.Busi_ComeInDelSwap '��Ժ�Ǽǳ���
            Case ����Enum.Busi_SettleSwap 'סԺ����
            Case ����Enum.Busi_SettleDelSwap 'סԺ�������
        End Select
    Else
        '����ʧ��
        Select Case intBusiness
            Case ����Enum.Busi_RegistSwap '����Һ�
            Case ����Enum.Busi_RegistDelSwap '����Һų���
            Case ����Enum.Busi_ClinicSwap '�������
            Case ����Enum.Busi_ClinicDelSwap '����������
            Case ����Enum.Busi_ComeInSwap '��Ժ�Ǽ�
            Case ����Enum.Busi_ComeInDelSwap '��Ժ�Ǽǳ���
            Case ����Enum.Busi_SettleSwap 'סԺ����
            Case ����Enum.Busi_SettleDelSwap 'סԺ�������
        End Select
    End If
End Sub

