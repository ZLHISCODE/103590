Attribute VB_Name = "mdl����"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

'API��������

'1���ӿڳ�ʼ��������������绷���Ƿ�ͨ������ҽԺ�ͻ�����ǰ�û���ǰ�û������ķ������䡣
Private Declare Function dy_Init Lib "SiInterface" Alias "INIT" () As Long

'2 ҵ��������ִ��ҽ��ҵ������Ҫ�Ĵ���
Private Declare Function dy_Business_Handle Lib "SiInterface" Alias "BUSINESS_HANDLE" _
    (ByVal InputData As String, ByVal OutputData As String) As Long

'Private gobj�����ж��� As New clsT_CQDRYB
Private mstr������ֹ���� As String                  '���汾��סԺ�������Աѡ��ķ�����ֹ����
Private mstrҽ���� As String
Private mdbl��� As Double
Private mlng����ID As Long
Private mstr����� As String
Private mstr���ı�� As String

Private mblnIint As Boolean
Private mblnFail As Boolean                         '��ʼ��ʧ���ѻ�����

Private gstr����ʱ�� As String
Private mdbl�����ܶ� As Double
Private mbln�൥���շ� As Boolean
Private mbln���� As Boolean                         '�൥��������ֻ����ҽ��һ�ν��㣬�ô����ж�
Private mstr������ˮ�� As String                    '���ڶ൥���շѵ�ÿ����¼��������ͬ����ˮ�ţ����ڲ�֤
Private mcnYB As New ADODB.Connection

'���½ṹ�����ڼ�¼������������Ա��ڽ���ʱ�˶�
Private Type typBalance
    cur�����ʻ� As Double
    curͳ��֧�� As Double
    cur���ͳ�� As Double
    cur����Ա���� As Double
    cur����Ա���� As Double
    curҽԺ��֧ As Double
    cur�������� As Double
    HIS����ҽԺ��֧ As Boolean          '�Ƿ���HIS������ҽԺ��֧
End Type
Private pre_Balance As typBalance

'###############################################################################
'20061113,zyb:ȡ��TrackRecordInsure()�ĵ��ã������ҽ��ǰ�û����������ࣨ��������������������dy_initҲ���һ�����ӣ�
'###############################################################################

Public Function ҽ����ʼ��_����() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim lngReturn As Long
    
    If mblnIint = True Then
        'ֻ��Ҫ����һ��
        ҽ����ʼ��_���� = True
        Exit Function
    End If
    
    On Error Resume Next
    
    If gclsInsure.GetCapability(Support��ʼ��ʧ���ѻ�����, 0, TYPE_������) And mblnFail Then
        Exit Function
    Else
        lngReturn = dy_Init
'        lngReturn = gobj�����ж���.dy_Init
    End If
    If Err <> 0 Then
        MsgBox "������ȷ����ҽ���ӿڳ���", vbInformation, gstrSysName
        mblnFail = True
        Exit Function
    End If
    
    If lngReturn = -1 Then
        mblnFail = True
        MsgBox "�������ҽ���ӿڳ�ʼ�������������������绷���Ƿ�ͨ��������" & vbCrLf & vbCrLf & _
          "1��ҽԺ�ͻ�����ҽԺǰ�û�Ӧ�÷�����֮�䣻" & vbCrLf & _
          "2��ҽԺǰ�û�Ӧ�÷�������ҽ������Ӧ�÷�����֮�䡣", vbInformation, gstrSysName
    Else
        ҽ����ʼ��_���� = True
        mblnIint = True
    End If
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    'alter table �����ʻ� add (������Ϣ varchar2(200),�������� varchar2(250));
    Dim strҽ���� As String, StrInput As String, arrOutput  As Variant, int��� As Integer
    Dim STR���� As String, str�Ա� As String, str���֤���� As String, lng���� As Long
    Dim str�������� As String, str��Ա��� As String, str��λ���� As String, str��λ���� As String
    Dim strIdentify As String, str���� As String, str����� As String, str������Ϣ As String
    Dim datCurr As Date
    Dim lngTemp As Long, str�ϴξ���ʱ�� As String
    Dim bln���ݾ���¼����� As Boolean
    Dim intҵ������ As Integer
    Dim str���� As String, str�������� As String, str����֢ As String, str�������� As String
    Dim rsTemp As New ADODB.Recordset
    Dim rs���� As ADODB.Recordset
    
    Call DebugTool("���������֤")
    '��ʼ��һЩ����
    mlng����ID = 0
    mstr����� = ""
    mstrҽ���� = ""
    mdbl��� = 0
    
    '����ǹҺţ���ֱ�ӵ������ﴦ��
    If bytType = 3 Then bytType = 0
    int��� = bytType
    If frmIdentify����.GetIdentify(strҽ����, int���) = False Then
        Exit Function
    End If
    
    Call DebugTool("�ر������֤����")
    'ȡ���ղ��������ݾ���¼����ϡ�
    gstrSQL = "Select Nvl(����ֵ,0) AS ����ֵ From ���ղ��� Where ����=[1] And ������='���ݾ���¼�����'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ղ��������ݾ���¼����ϡ�", TYPE_������)
    If rsTemp.RecordCount <> 0 Then
        bln���ݾ���¼����� = rsTemp!����ֵ
    End If
    Call DebugTool("ȡ���ղ���")
    
    '��������ѽ��й��շѣ��򲻱��ٴ�ˢ����������ˮ�����ϴεĺ�Ϊ׼
    If bytType = 0 Then lngTemp = GetRegisted(strҽ����, str�ϴξ���ʱ��)
    datCurr = zlDatabase.Currentdate
    Call DebugTool("�жϵ����Ƿ����չ���")
    
    '��Ȼ���������������������Ȼ�ǲ������µ���ˮ�Ų�����ҽ���Ǽ�
    If lngTemp <> 0 Then
        Call DebugTool("lngTemp<>0")
        
        lng����ID = lngTemp
        '��ȡ�ϴξ����ҵ������
        gstrSQL = "Select ҵ������,�ʻ����,����֤��,��������,����֢ From �����ʻ� Where ����=[1] ANd ����ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҵ������", TYPE_������, lng����ID)
        intҵ������ = rsTemp!ҵ������
        mdbl��� = Nvl(rsTemp!�ʻ����, 0)
        str�������� = Nvl(rsTemp!����֤��)
        str�������� = Nvl(rsTemp!��������)
        str����֢ = Nvl(rsTemp!����֢)
        '��ȡ������Ϣ��һ���������������ǼǼ�¼������ҽ����ˮ���ǲ���ģ�������ȡһ����
        gstrSQL = " Select A.����,A.�Ա�,A.����,A.���֤��,A.��������,B.��Ա��� AS ��Ա���,A.������λ AS ��λ����,B.ҵ������,B.���ı��,C.HIS��ˮ�� AS ˳��� " & _
                  " From ������Ϣ A,�����ʻ� B,����ǼǼ�¼ C" & _
                  " Where C.��ҳID=0 And C.��¼ID Is Not NULL And C.����ID=A.����ID And A.����ID=B.����ID And B.����=C.����" & _
                  " And C.����=[1] And C.����ʱ��=[2] And C.����ID=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", TYPE_������, CDate(str�ϴξ���ʱ��), lng����ID)
        
        STR���� = rsTemp!����
        str�Ա� = rsTemp!�Ա�
        lng���� = Val(Nvl(rsTemp!����, 0))
        str���֤���� = Nvl(rsTemp!���֤��)
        str�������� = Format(rsTemp!��������, "yyyy-MM-dd")
        
        str��Ա��� = Nvl(rsTemp!��Ա���)
        str��λ���� = ""
        str��λ���� = Nvl(rsTemp!��λ����) '50�ĳ��ȣ���Ҫ�۳�2������
        mstr���ı�� = Nvl(rsTemp!���ı��)
        str����� = Nvl(rsTemp!˳���)
        
        mlng����ID = lng����ID
        mstr����� = str�����
        mstrҽ���� = strҽ����
        
        Call DebugTool("�ѳɹ���ȡ������Ϣ")
    Else
        Call DebugTool("׼������01����")
        '���ýӿ�
        StrInput = "01|" & strҽ����
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
        
        'ȡ�÷���ֵ
        STR���� = arrOutput(1)
        str�Ա� = arrOutput(2)
        lng���� = Val(arrOutput(3))
        str���֤���� = arrOutput(4)
        str�������� = Get��������(str���֤����, lng����)
        
        str��Ա��� = ToVarchar(arrOutput(7), 8) 'VARCHAR2 (20) ��ְ����ְפ�⣬��ʱ�ù�������ְҵ��ת�ɣ����ݣ�������ؾ�ס����ְ����ְ��ؾ�ס��
        'arrOutput(8)   ����Ա��־               'VARCHAR2 (3)
        str��λ���� = ""
        str��λ���� = ToVarchar(arrOutput(9), 48) '50�ĳ��ȣ���Ҫ�۳�2������
        mstr���ı�� = arrOutput(10)
        
        If arrOutput(11) = "2" Then
            MsgBox "�ò���ҽ�������ܼ���ʹ�á�" & arrOutput(12)
            Exit Function
        End If
        
        str������Ϣ = ""
        If arrOutput(11) <> "0" Then
            'סԺʱҪ����
            str������Ϣ = arrOutput(12)
            MsgBox str������Ϣ, vbInformation, gstrSysName
        End If
        Call DebugTool("01���׵��óɹ���")
    End If
    
    '����;ҽ����;����;����;�Ա�;��������;���֤;������λ
    'ҽ���ŵ�һλΪ������
    strIdentify = strҽ���� & ";" & strҽ���� & ";;" & STR���� & ";" & str�Ա� & ";" & str�������� & ";" & str���֤���� & ";" & str��λ���� & "(" & str��λ���� & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    str���� = ";"                                       '8.���Ĵ���
    str���� = str���� & ";"                             '9.˳���
    str���� = str���� & ";" & str��Ա���               '10��Ա���
    str���� = str���� & ";0"                            '11�ʻ����
    str���� = str���� & ";0"                            '12��ǰ״̬
    str���� = str���� & ";"                             '13����ID
    str���� = str���� & ";" & IIf(Left(str��Ա���, 1) = "��", 2, 1)     '14��ְ(1,2)
    str���� = str���� & ";"                             '15����֤��
    str���� = str���� & ";" & lng����                   '16�����
    str���� = str���� & ";"                             '17�Ҷȼ�
    str���� = str���� & ";0"                            '18�ʻ������ۼ�
    str���� = str���� & ";0"                            '19�ʻ�֧���ۼ�
    str���� = str���� & ";"                             '20����ͳ���ۼ�
    str���� = str���� & ";"                             '21ͳ�ﱨ���ۼ�
    str���� = str���� & ";"                             '22סԺ�����ۼ�
    str���� = str���� & ";" & IIf(int��� = 14, 1, "")  '23�������� (1����������)
    
'    If lngTemp = 0 Then
        '����ǵ�һ�ξ���
        lng����ID = BuildPatiInfo(bytType, strIdentify & str����, lng����ID, TYPE_������)
        str����� = ToVarchar(lng����ID & Format(datCurr, "yyMMddHHmmss"), 18)
'    End If
    Call DebugTool("�ɹ��������˵���")
    
    If bytType = 0 Then        '��������ͬʱ���о���Ǽ�
'        '��������ⲡ�������ȣ���Ҫѡ���˼���
'        If lngTemp <> 0 Then
'            '��������
'            If intҵ������ <> int��� Or int��� = 13 Or int��� = 14 Or Mid(strҽ����, 1, 1) = "2" Then
'                If int��� = 13 Or int��� = 14 Or Mid(strҽ����, 1, 1) = "2" Then
'                    If int��� = 13 Then
'                        '���������Ϣ
'                        strInput = "07|" & strҽ����
'                        If HandleBusiness(strInput, arrOutput) = False Then Exit Function
'
'                        str���� = "���ⲡ"
'                        If frm����ѡ������.GetCode(arrOutput, str����, str��������, str��������, str����֢) = False Then Exit Function
'                    ElseIf int��� = 14 Then
'                        str���� = "����"
'                        If frm����ѡ������.GetCode("", str����, str��������, str��������, str����֢) = False Then Exit Function
'                    Else
'                        str���� = "��Ժ"
'                        If frm����ѡ������.GetCode("", str����, str��������, str��������, str����֢) = False Then Exit Function
'                    End If
'                Else
'                    str�������� = ""
'                    str�������� = ""
'                    str����֢ = ""
'                End If
'
'                '��Ҫ���þ�����Ϣ�޸ģ������������н���
'                '���ýӿڸ��£�סԺ��(�����)|���±�־|ҽ�����|����|ҽ��|��Ժ����|��Ժ���|��Ժ����|ȷ�Ｒ������
'                              '|��Ժԭ��|������|����֢
'                strInput = "03|" & str����� & "|100000" & IIf(str�������� <> "", "1", "0") & "00" & IIf(str����֢ <> "", "1", "0") & "|" & int��� & "||||||" & str�������� & "|||" & str����֢
'                If HandleBusiness(strInput, arrOutput) = False Then Exit Function
'
'                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֤��','''" & str�������� & "''')"
'                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
'                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'��������','''" & str�������� & "''')"
'                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
'                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֢','''" & str����֢ & "''')"
'                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���֢")
'            End If
'        Else
    
            Call DebugTool("׼����������ѡ����")
            If int��� = 13 Or int��� = 15 Or int��� = 14 Or (mstr���ı�� = "20" And bln���ݾ���¼�����) Then
                If int��� = 13 Then
                    '���������Ϣ
                    StrInput = "07|" & strҽ����
                    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
                    
                    str���� = "���ⲡ"
                    If frm����ѡ������.GetCode(arrOutput, str����, str��������, str��������, str����֢) = False Then Exit Function
                ElseIf int��� = 14 Then
                    str���� = "����"
                    If frm����ѡ������.GetCode("", str����, str��������, str��������, str����֢) = False Then Exit Function
                Else
                    str���� = "��Ժ"
                    If frm����ѡ������.GetCode("", str����, str��������, str��������, str����֢) = False Then Exit Function
                End If
                
                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֤��','''" & str�������� & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'��������','''" & str�������� & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֢','''" & str����֢ & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���֢")
            End If
            
            Call DebugTool("׼�����þ���Ǽ�")
            StrInput = "02|" & str����� & "|" & int��� & "|" & strҽ���� & _
                       "|����|" & ToVarchar(UserInfo.����, 20) & "|" & _
                       Format(datCurr, "yyyy-MM-dd") & "|" & str�������� & "|" & ToVarchar(UserInfo.����, 20) & "|" & str����֢
            If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
            
            mlng����ID = lng����ID
            mstr����� = str�����
            mstrҽ���� = strҽ����
            mdbl��� = Val(arrOutput(2))
'        End If
    End If
     
     'Modified by ZYB 2006-02-6
    '���ýӿڸ��£�סԺ��(�����)|���±�־|ҽ�����|����|ҽ��|��Ժ����|��Ժ���|��Ժ����|ȷ�Ｒ������
                  '|��Ժԭ��|������|����֢
    If int��� = 15 Or int��� = 14 Then
        Call DebugTool("׼������03����")
        StrInput = "03|" & str����� & "|0000001001||" & _
            "|||||" & str�������� & "|||" & str����֢
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    End If
    
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
        mdbl��� & ",0,0,0,0,0,0,0,0,0,'')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    If lngTemp = 0 Then
        '���·�����Ϣ
        gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'������Ϣ','''" & str������Ϣ & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
        
        '����Ų����µ������ʻ��У���������ǼǼ�¼
        '����סԺ����Ҫ���²������ơ�����֢����Ϣ���ܽ���ģ������Щ������Ȼ�����ڱ����ʻ��У�����ǼǼ�¼ֻ�Ǹ�������д
'        gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'˳���','''" & str����� & "''')"
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
        
    End If
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'���ı��','''" & mstr���ı�� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������ı��")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'���䲡��','''" & "0" & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ȡ�����䲡����־")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'ҵ������','''" & int��� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
    
    '����Ų����µ������ʻ��У���������ǼǼ�¼
    '����סԺ����Ҫ���²������ơ�����֢����Ϣ���ܽ���ģ������Щ������Ȼ�����ڱ����ʻ��У�����ǼǼ�¼ֻ�Ǹ�������д
    '��д����ǼǼ�¼���������£�
    '����_IN             ����ǼǼ�¼.����%TYPE,
    '����ID_IN           ����ǼǼ�¼.����ID%TYPE,
    '��ҳID_IN           ����ǼǼ�¼.��ҳID%TYPE,
    '����ʱ��_IN         ����ǼǼ�¼.����ʱ��%TYPE,
    '״̬_IN             ����ǼǼ�¼.״̬%TYPE:= 0,
    'ҽ�����_IN         ����ǼǼ�¼.ҽ�����%TYPE:=NULL,
    '�ʻ����_IN         ����ǼǼ�¼.�ʻ����%TYPE:=0,
    '����ID_IN           ����ǼǼ�¼.����ID%TYPE:=NULL,
    '��������_IN         ����ǼǼ�¼.��������%TYPE:=NULL,
    '����֢_IN           ����ǼǼ�¼.����֢%TYPE:=NULL,
    'IC����Ϣ_IN         ����ǼǼ�¼.IC����Ϣ%TYPE:=NULL,
    'HIS��ˮ��_IN        ����ǼǼ�¼.HIS��ˮ��%TYPE:=NULL,
    'YB��ˮ��_IN         ����ǼǼ�¼.YB��ˮ��%TYPE:=NULL,
    '��ע_IN             ����ǼǼ�¼.��ע%TYPE:=NULL
    gstrSQL = "zl_����ǼǼ�¼_UPDATE(" & _
        TYPE_������ & "," & lng����ID & ",0,to_date('" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & _
        "1," & int��� & "," & mdbl��� & ",NULL," & _
        "'" & str�������� & "-" & str�������� & "','" & str����֢ & "','" & str������Ϣ & "','" & str����� & "','" & str����� & "','" & mstr���ı�� & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    gstr����ʱ�� = Format(datCurr, "yyyy-MM-dd HH:mm:ss")
    
    g��������.�����Ը���� = int��� '������ʱ���棬�������
    Call DebugTool("ִ�гɹ���")
    
    '���ظ�ʽ:�м���벡��ID
    If lng����ID <> 0 Then
        ��ݱ�ʶ_���� = strIdentify & ";" & lng����ID & str����
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(strSelfNo As String) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: strSelfNO-���˸��˱��
'����: ���ظ����ʻ����Ľ��
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHandle
    
    '�����ݿ��ж�ȡ����Ϊ�ղŲű����˵ģ�Ӧ����׼ȷ�ģ�
    If mstrҽ���� = "" Or strSelfNo <> mstrҽ���� Then
        gstrSQL = "Select �ʻ���� From �����ʻ� where ����=[1] and ����=0 and ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������, strSelfNo)
        
        If rsTemp.EOF = False Then
            �������_���� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
        End If
    Else
        �������_���� = mdbl���
    End If
    'ֻ����һ��
    mstrҽ���� = ""
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, Optional ByRef strAdvance As String = "") As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Static str�����Pre As String
    Dim strҽ���� As String, StrInput As String, arrOutput  As Variant
    Dim strMessage As String
    Dim lng����ID As Long, str��� As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    
    Dim int����_CUR As Integer, int����_MAX As Integer
    Dim dblҽ������_CUR As Double, dbl�����ʻ�_CUR As Double, dbl����Ա����_CUR As Double, dbl���ͳ��_CUR As Double, dbl����Ա����_CUR As Double, dblҽԺ��֧_CUR As Double, dbl��������_CUR As Double
    Dim dblҽ������ As Double, dbl�����ʻ� As Double, dbl����Ա���� As Double, dbl���ͳ�� As Double, dbl����Ա���� As Double, dblҽԺ��֧ As Double, dbl�������� As Double
    
    On Error GoTo errHandle
    
    If rs��ϸ.RecordCount = 0 Then
        str���㷽ʽ = "�����ʻ�;0;0"
        �����������_���� = True
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    datCurr = zlDatabase.Currentdate
    
    '�ֽⵥ����������ǰ����
    If strAdvance = "" Then strAdvance = "1|1"
    int����_CUR = Val(Split(strAdvance, "|")(1))
    int����_MAX = Val(Split(strAdvance, "|")(0))
    mbln�൥���շ� = (int����_MAX > 1)              '�����������ڱ��ս����¼�У�������ʶ�Ƿ��Ƕ൥���շ�
    mbln���� = False
    Call DebugTool("�൥���շѱ�־:" & mbln�൥���շ� & ";�����־:" & mbln����)
    
    If mlng����ID <> lng����ID Then
        MsgBox "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mbln�൥���շ� And �൥���շ�_�շѷֱ��ӡ Then
        MsgBox "��ȡ��ϵͳ����������Ʊ��ҳ����Ĳ����������շ�ÿ�ŵ��ݷֱ��ӡ�������ɽ��ж൥���շѣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '������Ƕ൥���շѣ����߶൥���շ��еĵ�һ�ŵ��ݣ���ɾ�����ϴ���ϸ��Ȼ�����������㱣��������Ľṹ��
    'ɾ���������ϴ�����ϸ��ָ����ҵ����δ����ķϼ�¼��
    If int����_CUR = 1 Then
        '��ʽ����ʱ��strAdvanceʼ�մ���"1|1"
        mdbl�����ܶ� = 0
        pre_Balance.cur�����ʻ� = 0
        pre_Balance.curͳ��֧�� = 0
        pre_Balance.cur���ͳ�� = 0
        pre_Balance.cur����Ա���� = 0
        pre_Balance.cur����Ա���� = 0
        pre_Balance.curҽԺ��֧ = 0
        pre_Balance.cur�������� = 0  '20101028������������
        
        '�����˵���ǰ����������δ��ķ��ã��������ִ��Ԥ����
        StrInput = "10|" & mstr����� & "|" & mstr�����
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
        Call DebugTool("������ϴ�δ����Ĵ�����ϸ")
    End If
    
    '�����ֵ
    str�����Pre = mstr�����
    
    'Ȼ����봦����ϸ
    Do Until rs��ϸ.EOF
        gstrSQL = "select A.����,A.����,A.���,A.���,A.���㵥λ,B.��Ŀ����,B.��ע,A.���㵥λ,E.���,G.���� ���� " & _
                  "from �շ�ϸĿ A,����֧����Ŀ B,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
                  "where A.ID=[1] and A.ID=B.�շ�ϸĿID and B.����=[2]" & _
                 "        AND A.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", CLng(rs��ϸ("�շ�ϸĿID")), TYPE_������)
        If rsTemp.EOF = True Then
            MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
            Exit Function
        End If
        mdbl�����ܶ� = mdbl�����ܶ� + Nvl(rs��ϸ!ʵ�ս��, 0)
        If Val(Nvl(rs��ϸ("ʵ�ս��"), 0)) <> 0 Then
            StrInput = "04|" & mstr����� & "|" & mstr����� & "|" & Format(datCurr, "yyyy-MM-dd HH:mm:ss")
            StrInput = StrInput & "|" & ToVarchar(rsTemp("��Ŀ����"), 10)  'ҽ����ˮ��
            StrInput = StrInput & "|" & ToVarchar(rsTemp("����"), 20)      'ҽԺ����
            StrInput = StrInput & "|" & ToVarchar(rsTemp("����"), 50)      '��Ŀ����
            StrInput = StrInput & "|" & Format(rs��ϸ!ʵ�ս�� / Round(rs��ϸ!����, 2), "0.0000") '����   ���ܴ��ڴ���,����¼���е���Ϊԭʼ����
            StrInput = StrInput & "|" & Format(rs��ϸ("����"), "0.00")     '����
            StrInput = StrInput & "|" & IIf(rs��ϸ("�Ƿ���") = 1, 1, 0)  '�����־
            StrInput = StrInput & "|" & Format(Nvl(rs��ϸ!������, UserInfo.����), 20)         '����ҽ��
            StrInput = StrInput & "|" & Format(UserInfo.����, 20)          '������
            StrInput = StrInput & "|" & ToVarchar(rsTemp("���㵥λ"), 20)     '��λ
            StrInput = StrInput & "|" & ToVarchar(rsTemp("���"), 14)         '���
            StrInput = StrInput & "|" & ToVarchar(rsTemp("����"), 20)         '����
            StrInput = StrInput & "|"                                         '������ϸ��ˮ��
            StrInput = StrInput & "|" & Format(rs��ϸ("ʵ�ս��"), "#####0.0000")         '���
            
            If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
            Call AddMessage(strMessage, arrOutput, ToVarchar(rsTemp("����"), 50), Format(rs��ϸ!ʵ�ս�� / Round(rs��ϸ!����, 2), "0.0000"), False)
        End If
        rs��ϸ.MoveNext
    Loop
    
    If strMessage <> "" Then
        strMessage = "���˷�����ϸ��������еõ�ҽ���������·�����Ϣ���Ƿ������" & vbCrLf & vbCrLf & strMessage
        If MsgBox(strMessage, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            '�û�ѡ��ȡ�������˵���ϸ
            StrInput = "10|" & mstr����� & "|" & mstr�����
            If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
            
            Call DebugTool("�û�ѡ��ȡ�����˵���ϸ")
            Exit Function
        End If
    End If
    
    '����Ԥ����
    StrInput = "06|" & mstr����� & "|||" & Format(mdbl�����ܶ�, "#0.00;-#0.00;0;")
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    '��ֵ
    dbl�����ʻ� = Val(arrOutput(2))
    dblҽ������ = Val(arrOutput(1))
    dbl����Ա���� = Val(arrOutput(3))
    dbl���ͳ�� = Val(arrOutput(5))
    dbl����Ա���� = Val(arrOutput(6))
    If UBound(arrOutput) > 7 Then
        dblҽԺ��֧ = Val(arrOutput(8))
    End If
    If UBound(arrOutput) > 8 Then   '20101028������������
        dbl�������� = Val(arrOutput(9))
        dblҽ������ = dblҽ������ - dbl��������
    End If
    
    Call DebugTool("��ȡ����Ԥ������")
    
    '���㱾����ʵ�Ľ�����������൥���շѵ������
    dbl�����ʻ�_CUR = dbl�����ʻ� - pre_Balance.cur�����ʻ�
    dblҽ������_CUR = dblҽ������ - pre_Balance.curͳ��֧��
    dbl���ͳ��_CUR = dbl���ͳ�� - pre_Balance.cur���ͳ��
    dbl����Ա����_CUR = dbl����Ա���� - pre_Balance.cur����Ա����
    dbl����Ա����_CUR = dbl����Ա���� - pre_Balance.cur����Ա����
    dblҽԺ��֧_CUR = dblҽԺ��֧ - pre_Balance.curҽԺ��֧
    dbl��������_CUR = dbl�������� - pre_Balance.cur��������  '20101028������������
    Call DebugTool("�õ�����Ԥ�������ʵ���")
    
    '���ؽ��������൥�ݷ��ز������շѱ��εĽ�����ǲ�
    str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ�_CUR & ";0"   '�����޸ĸ����ʻ�����Ϊ����ʱ�Ѿ����ٴ���ǰ�û���
    If Val(Format(dblҽ������_CUR, "#0.00")) <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & dblҽ������_CUR & ";0"
    End If
    If Val(Format(dbl����Ա����_CUR, "#0.00")) <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & dbl����Ա����_CUR & ";0"
    End If
    If Val(Format(dbl���ͳ��_CUR, "#0.00")) <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|���ͳ��;" & dbl���ͳ��_CUR & ";0"
    End If
    If Val(Format(dbl����Ա����_CUR, "#0.00")) <> 0 Then
        str���㷽ʽ = str���㷽ʽ & "|����Ա����;" & dbl����Ա����_CUR & ";0"
    End If
    If Val(Format(dbl��������_CUR, "#0.00")) <> 0 Then    '20101028������������
        str���㷽ʽ = str���㷽ʽ & "|��������;" & dbl��������_CUR & ";0"
    End If
    If UBound(arrOutput) > 7 Then
        str���㷽ʽ = str���㷽ʽ & "|ҽԺ��֧;" & dblҽԺ��֧_CUR & ";0"
    End If
    Call DebugTool("���ؽ�������:" & str���㷽ʽ)
    
    '�����ۼ�ֵ
    pre_Balance.cur�����ʻ� = dbl�����ʻ�
    pre_Balance.curͳ��֧�� = dblҽ������
    pre_Balance.cur���ͳ�� = dbl���ͳ��
    pre_Balance.cur����Ա���� = dbl����Ա����
    pre_Balance.cur����Ա���� = dbl����Ա����
    pre_Balance.curҽԺ��֧ = dblҽԺ��֧
    pre_Balance.cur�������� = dbl��������   '20101028������������
    Call DebugTool("�����ۼ�ֵ")
    
    �����������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String, Optional strAdvance As String, Optional bln���� As Boolean = True) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim strҽ���� As String, StrInput As String
    Dim lng����ID  As Long
    Dim str����Ա As String, arrOutput  As Variant
    Dim datCurr As Date
    Dim str���� As String, str����֢ As String, str������Ϣ As String, str��ע As String
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Dim bln�ʻ��ۼ� As Boolean
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    
    Dim curͳ��֧�� As Double
    Dim cur����Ա���� As Double
    Dim cur���ͳ�� As Double
    Dim curҽԺ��֧ As Double
    Dim cur����Ա���� As Double, cur�������� As Double, cur�����ʻ���� As Double '20101028������������
    Dim cur�������� As Currency
    
    Dim str���㷽ʽ As String
    Dim blnOld As Boolean
    Dim blnRevise As Boolean
    Dim bln�ʻ�֧�� As Boolean
    Dim int�ʻ�֧����ʽ As Integer '0-֧��;1-סԺѯ��;2-����ѯ��;3-��֧��
    
    On Error GoTo errHandle
    
    Call DebugTool("�����������")
    strAdvance = ""
    gstrSQL = "Select * From ������ü�¼ Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    If rs��ϸ.EOF = True Then
        Err.Raise 9000 + VbMsgBoxStyle.vbExclamation, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")
    str����Ա = ToVarchar(IIf(IsNull(rs��ϸ("����Ա����")), UserInfo.����, rs��ϸ("����Ա����")), 20)
    
    If mlng����ID <> lng����ID Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣"
        Exit Function
    End If
    
''   ����ͳ�Ʒ����ܶ��ˣ���Ԥ�����ۼƵ�Ϊ׼
'    Do Until rs��ϸ.EOF
'        cur�������� = cur�������� + rs��ϸ("���ʽ��")
'        Call TrackRecordInsure(rs��ϸ!ID, rs��ϸ!�շ�ϸĿID)
'        rs��ϸ.MoveNext
'    Loop
'    cur�������� = Val(Format(cur��������, "#0.00;-#0.00;0;"))
    
    '��ȡ�α����˵Ĳ���֢��������Ϣ
    gstrSQL = "Select ����֤�� As ���ֱ���,��������,����֢,������Ϣ" & _
        " From �����ʻ� " & _
        " Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�α����˵Ĳ���֢��������Ϣ", lng����ID, TYPE_������)
    If Not rsTemp.EOF Then
        str���� = Nvl(rsTemp!���ֱ���)
        If str���� <> "" Then str���� = "[" & str���� & "]"
        str���� = str���� & Nvl(rsTemp!��������)
        str����֢ = Nvl(rsTemp!����֢)
        str������Ϣ = Nvl(rsTemp!������Ϣ)
    End If
    str��ע = str���� & "||" & str����֢ & "||" & str������Ϣ & "@@������" & IIf(mbln�൥���շ�, "�൥���շ�", "�����շ�")
    Call DebugTool("�ɹ���ȡ���������Ϣ�������ֱ��������ơ�����֢��������Ϣ")
    
    '���ý���
    bln�ʻ�֧�� = True
    If mbln���� = False Then
        '������Ƕ൥���շѣ���ѯ���Ƿ���и����ʻ�֧��
        If Not mbln�൥���շ� Then
            'ȡ���ղ��������ݾ���¼����ϡ�
            gstrSQL = "Select Nvl(����ֵ,0) AS ����ֵ From ���ղ��� Where ����=[1] And ������='�����ʻ�'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ղ��������ݾ���¼����ϡ�", TYPE_������)
            If rsTemp.RecordCount <> 0 Then
                int�ʻ�֧����ʽ = rsTemp!����ֵ
            End If
            If int�ʻ�֧����ʽ < 2 Then
                bln�ʻ�֧�� = True
            ElseIf int�ʻ�֧����ʽ = 2 Then
                If MsgBox("���ʱ���Ҫ���и����ʻ�֧����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    bln�ʻ�֧�� = False
                Else
                    bln�ʻ�֧�� = True
                End If
            Else
                bln�ʻ�֧�� = False
            End If
        End If
        
        Call DebugTool("׼�������������")
        StrInput = "05|" & mstr����� & "|1||" & str����Ա & "|" & IIf(bln�ʻ�֧��, "0", "1") & "||" & mdbl�����ܶ� '���ʻ����֧��
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
        mbln���� = True
        bln�ʻ��ۼ� = True
        
        '��������¼
        '---------------------------------------------------------------------------------------------
        mstr������ˮ�� = arrOutput(1)
        curͳ��֧�� = Val(arrOutput(2))
        cur�����ʻ� = Val(arrOutput(3))
        cur����Ա���� = Val(arrOutput(4))
        cur����Ա���� = Val(arrOutput(7))
        cur���ͳ�� = Val(arrOutput(6))
        If UBound(arrOutput) > 8 Then
            curҽԺ��֧ = Val(arrOutput(9))
        End If
        If UBound(arrOutput) > 9 Then   '20101028������������
            cur�������� = Val(arrOutput(10))
            curͳ��֧�� = curͳ��֧�� - cur��������
            cur�����ʻ���� = Val(arrOutput(11))
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'���������������','''" & cur�����ʻ���� & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�������������������")
        End If
        
        
        Call DebugTool("�õ�������")
        
        '���ս����¼�ı�ע�б��汾���Ƿ�൥���շѣ��Լ��Ƿ��ǵ�һ�ŵ���
        str��ע = str��ע & "||1"
    
        '���¾���ǼǼ�¼�ļ�¼ID���Ա��뱣�ս����¼����
        gstrSQL = "zl_����ǼǼ�¼_����(" & TYPE_������ & "," & lng����ID & ",0," & _
            "to_date('" & gstr����ʱ�� & "','yyyy-MM-dd hh24:mi:ss')," & lng����ID & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        '������Ƕ൥���շ�
        If Not mbln�൥���շ� And bln���� Then
            'ֻ����ָ����ʻ�ʵ��֧����������㲻���������ԭ�����£�
            '1�����Ԥ�����ʱ�䲻���㣬�п������Ľ����ʻ����������Ᵽ����������������
            '2����һ�ν���ʧ�ܣ����ʻ����£��ٴε�ȷ��ʱ�����ʻ����Ϊ�㵼�µڶ����¿�Ϊ�㣬�Ӷ����ֶ������������ʵ�ʽ��㲻�������
            If pre_Balance.cur�����ʻ� <> cur�����ʻ� Or pre_Balance.curͳ��֧�� <> curͳ��֧�� Then
                blnRevise = True
                
                str���㷽ʽ = "�����ʻ�|" & cur�����ʻ�
                str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curͳ��֧��
                If cur���ͳ�� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||���ͳ��|" & cur���ͳ��
                If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
                If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
                If cur�������� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||��������|" & cur��������    '20101028������������
                If curҽԺ��֧ <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽԺ��֧|" & curҽԺ��֧
                If str���㷽ʽ <> "" Then
                    #If gverControl < 2 Then
                        blnOld = True
                        gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
                    #Else
                        strAdvance = str���㷽ʽ
                        gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
                    #End If
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
                End If
            End If
        End If
    Else
        '��һ�ŵ��ݱ������е��ܶ���Ķ��ŵ��ݣ������ܶҽ������Ϊ��
        If mbln�൥���շ� = False Then
            MsgBox "���棺�����Ա��ͼ����¼���β������̣���������������˾������ϵ��", vbInformation, gstrSysName
        End If
        
        mdbl�����ܶ� = 0
        curͳ��֧�� = 0
        cur�����ʻ� = 0
    End If
    
    '�ʻ������Ϣ
    datCurr = zlDatabase.Currentdate
    
    If bln�ʻ��ۼ� Then
        Call DebugTool("�����ʻ������Ϣ")
        Call Get�ʻ���Ϣ(TYPE_������, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
            cur����ͳ���ۼ� + curͳ��֧�� & "," & _
            curͳ�ﱨ���ۼ� + curͳ��֧�� & "," & intסԺ�����ۼ� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    End If
    
    'g��������.�����Ը�����б���������ﲡ�˾������ͣ�������ⲡ�������ͨ����������¼�ı�ע������ǲ��ֵ�����
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)'�����Ը����������ʱ���棬�������
    Call DebugTool("���汣�ս����¼")
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & mdbl�����ܶ� & ",0,0," & _
        curͳ��֧�� & "," & curͳ��֧�� & "," & cur�������� & "," & g��������.�����Ը���� & "," & cur�����ʻ� & ",'" & mstr������ˮ�� & "',NULL,NULL,'" & str��ע & "'" & _
        IIf(blnOld, "", IIf(blnRevise, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    �������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim curƱ���ܽ�� As Currency
    Dim curDate As Date
        
    On Error GoTo errHandle
    curDate = zlDatabase.Currentdate
    
    '�˴��벻��ע�ͣ�Ҫȡ����ID����Ȼ�������ʻ������ϢʱҪ����
    gstrSQL = "Select * From ������ü�¼ " & _
        " Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    Do Until rsTemp.EOF
        If lng����ID = 0 Then lng����ID = rsTemp("����ID")

        curƱ���ܽ�� = curƱ���ܽ�� + rsTemp("���ʽ��")
        rsTemp.MoveNext
    Loop
    curƱ���ܽ�� = Val(Format(curƱ���ܽ��, "#####0.00"))
    
    '�˷�
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    lng����ID = rsTemp("����ID")
'
'    '����ҽ����Ŀ��״̬������
'    gstrSQL = "Select * From ���˷��ü�¼ " & _
'        " Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
'    Call OpenRecordset(rsTemp, "����ҽ��")
'    Do While Not rsTemp.EOF
'        Call TrackRecordInsure(rsTemp!ID, rsTemp!�շ�ϸĿID)
'        rsTemp.MoveNext
'    Loop
    
    Call �൥���շ�_�˷�(lng����ID)
    
    gstrSQL = "select * from ���ս����¼ where ����=1 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", TYPE_������, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    str��ˮ�� = rsTemp("֧��˳���")
    cur�����ʻ� = Nvl(rsTemp!�����ʻ�֧��, 0)
    
    '����Ƕ൥���շѣ��Ҳ��ǵ�һ�ŵ��ݣ���ֱ�ӷ�����
    If InStr(rsTemp!��ע, "@@") <> 0 Then
        If UBound(Split(Split(rsTemp!��ע, "@@")(1), "||")) > 0 Then
            If Val(Split(Split(rsTemp!��ע, "@@")(1), "||")(1)) = 0 Then
                ����������_���� = True
            End If
        Else
            ����������_���� = True
        End If
    End If
    
    If ����������_���� = False Then
        StrInput = "99|" & str��ˮ�� & "|" & ToVarchar(UserInfo.����, 20)
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, lng����ID, Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - cur�����ʻ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp!�������ý�� * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0," & rsTemp("�����Ը����") & "," & _
        cur�����ʻ� * -1 & ",'" & str��ˮ�� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")

    ����������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
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
    Dim StrInput As String, arrOutput  As Variant
    Dim datCurr As Date, rsTemp As New ADODB.Recordset
    Dim str���� As String, str˳��� As String
    Dim strTemp As String, str��ʾ As String, str��� As String
    Dim strסԺ����� As String
    Dim strҽ����� As String
    
    On Error GoTo errHandle
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.������Ϣ from ������ A where A.����ID=[1] and A.��ҳID=[2]" & _
              " and A.�������=1 and A.��ϴ���=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    If rsTemp.EOF = False Then
        str��� = ToVarchar(rsTemp("������Ϣ"), 40)
    End If
    
    '���ҽ����
    gstrSQL = "select ����,ҽ���� from �����ʻ� where ����=[1] and ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", TYPE_������, lng����ID)
    str���� = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
    strҽ���� = rsTemp("ҽ����")
    
    '��������������������ֶΡ�סԺ����š�������ֵΪ�գ���ǿ�ƽ�סԺ�������Ϊ��
    strסԺ����� = ""
    If GetMode(lng����ID, lng��ҳID, strסԺ�����) = False Then
        gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'סԺ�����','NULL')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
    End If
    
    '���������Ժ��Ϣ
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.��Ժ��ʽ,nvl(A.����Ժת��,0) as ����Ժת��,A.����ҽʦ,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ���� from ������ҳ A,���ű� B " & _
             " Where A.��Ժ����ID = B.ID And A.����ID =[1] And A.��ҳID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    'ȡҽ�����
    strҽ����� = IIf(rsTemp!��Ժ��ʽ = "�Ҵ�", "23", IIf(rsTemp!��Ժ��ʽ = "ת��", 22, 21))
'    If Is��ͥ����(lng����ID, lng��ҳID) Then
'        strҽ����� = "23"
'    Else
'        strҽ����� = IIf(rsTemp!��Ժ��ʽ = "ת��", 22, 21)
'    End If
    
    '������Ժ�ӿ�
    StrInput = "02|" & GetIdentify(lng����ID, lng��ҳID, True) & "|" & strҽ����� & "|" & strҽ���� & "|" & _
               ToVarchar(rsTemp("��Ժ����"), 30) & "|" & ToVarchar(rsTemp("����ҽʦ"), 20) & "|" & _
               Format(rsTemp("��Ժ����"), "yyyy-MM-dd") & "|" & ToVarchar(str���, 40) & "|" & ToVarchar(UserInfo.����, 20) & "|0"
    Call DebugTool("��������Ժ�ӿڣ�" & StrInput)
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    str˳��� = arrOutput(1)
    mdbl��� = Val(arrOutput(2))
    Call DebugTool("��������ˮ�����ʻ���" & str˳��� & "��" & mdbl���)
    
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
        mdbl��� & ",0,0,0,0,0,0,0,0,0,'')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    'ǿ�ưѵǼ�˳��š����µ�ҽ��������
    gstrSQL = "ZL_�����ʻ�_�޸�ҽ����(" & lng����ID & "," & TYPE_������ & _
                ",'" & str���� & "','" & strҽ���� & "','" & str˳��� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
            'ȡ��Ժ�Ǽ���֤�����ص�˳���
    Dim datCurr As Date, rsTemp As New ADODB.Recordset
    Dim StrInput As String, arrOutput  As Variant, bln����ó�Ժ As Boolean
    Dim str��� As String
    Dim strסԺ����� As String
    
    On Error GoTo errHandle
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.������Ϣ from ������ A where A.����ID=[1] and A.��ҳID=[2]" & _
              " and A.�������=3 and A.��ϴ���=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    If rsTemp.EOF = False Then
        str��� = Nvl(rsTemp("������Ϣ"), "��")
    Else
        str��� = "��"   '��ϲ�����β���Ϊ��
    End If
    str��� = ToVarchar(str���, 40)
    
    '���������Ժ��Ϣ
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.סԺҽʦ,A.��Ժ����,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ���� from ������ҳ A,���ű� B " & _
             " Where A.��Ժ����ID = B.ID And A.����ID =[1] And A.��ҳID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    '���ýӿڣ�������סԺ��Ϣ
    StrInput = "03|" & GetIdentify(lng����ID, lng��ҳID) & "|0000010010|21|||" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd") & "||" & _
                Format(rsTemp("��Ժ����"), "yyyy-MM-dd") & "|||" & ToVarchar(UserInfo.����, 20) & "|0"
    
    '���ô�סԺ�Ƿ�û�з��÷���
    gstrSQL = "Select nvl(sum(ʵ�ս��),0) as ���  from סԺ���ü�¼ where ����ID=[1] and ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���˳�Ժ", lng����ID, lng��ҳID)
    If rsTemp.EOF = True Then
        bln����ó�Ժ = True
    Else
        bln����ó�Ժ = (rsTemp("���") = 0)
    End If
    
    If bln����ó�Ժ = True Then
        '��������ó�Ժ���ͽ��䴦��Ϊ����Ժ�������ø�����סԺ��Ϣ
        gstrSQL = "Select ˳��� from �����ʻ� where ����ID=[1] and ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���˳�Ժ", lng����ID, TYPE_������)
        StrInput = "99|" & rsTemp("˳���") & "|" & ToVarchar(UserInfo.����, 20)
        If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    End If
    
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '����Ƿ����δ����ã�������δ����õĲ��˲���������Ժ
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        MsgBox "�ò��˲�����δ����ã��������������Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If MsgBox("��ȷ��Ҫ���ò��˳�����Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_������ & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    ��Ժ�Ǽǳ���_���� = True
End Function

Public Function ת�����ͥ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '���ýӿڸ��£�סԺ��(�����)|���±�־|ҽ�����|����|ҽ��|��Ժ����|��Ժ���|��Ժ����|ȷ�Ｒ������
                  '|��Ժԭ��|������|����֢
    Dim StrInput As String
    Dim arrOutput
    
    StrInput = "03|" & GetIdentify(lng����ID, lng��ҳID) & "|1000000000|24|||||||||"
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'���䲡��','''" & "1" & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���䲡����־����Ϊ1")
End Function

Public Function ���³�Ժ����_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'���ܣ����²��˵ĳ�Ժ����������������������ʱ���߻����
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str���� As String, str����֢ As String, str�������� As String, str�������� As String
    Dim StrInput As String, arrOutput  As Variant
    Dim str���� As String, strҽ�� As String, bln���䲡�� As Boolean
    Dim str��Ժ���� As String, str��Ժ���� As String, str��� As String, strҽ����� As String
    
    On Error GoTo errHandle
    
    '��ò��˳�Ժ���ּ�����֢
    gstrSQL = "Select ����֤�� ���ֱ���,��������,����֢,���䲡�� From �����ʻ� Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˳�Ժ���ּ�����֢", lng����ID)
    str�������� = Nvl(rsTemp!���ֱ���)
    str����֢ = Nvl(rsTemp!����֢)
    str�������� = Nvl(rsTemp!��������)
    bln���䲡�� = (Nvl(rsTemp!���䲡��, "0") = "1")
    
    str���� = "��Ժ"
    If frm����ѡ������.GetCode("", str����, str��������, str��������, str����֢) = False Then
        Exit Function
    End If
    str�������� = ToVarchar(str��������, 20)
    str����֢ = ToVarchar(str����֢, 200)
    str�������� = TrimStr(str��������)
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.������Ϣ from ������ A where A.����ID=[1] and A.��ҳID=[2]" & _
              " and A.�������=1 and A.��ϴ���=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    If rsTemp.EOF = False Then
        str��� = ToVarchar(rsTemp("������Ϣ"), 40)
    End If
    
    'ȡ���˵���Ժ��Ϣ
    gstrSQL = "select A.��Ժ��ʽ,nvl(A.����Ժת��,0) as ����Ժת��,A.����ҽʦ,A.��Ժ����,A.��Ժ����,A.��Ժ����,B.���� as ��Ժ���� from ������ҳ A,���ű� B " & _
             " Where A.��Ժ����ID = B.ID And A.����ID = [1] And A.��ҳID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���˵���Ժ��Ϣ", lng����ID, lng��ҳID)
    If bln���䲡�� = False Then
        strҽ����� = IIf(rsTemp!��Ժ��ʽ = "�Ҵ�", "23", IIf(rsTemp!��Ժ��ʽ = "ת��", 22, 21))
'        If Is��ͥ����(lng����ID, lng��ҳID) Then
'            strҽ����� = "23"
'        Else
'            strҽ����� = IIf(rsTemp!��Ժ��ʽ = "ת��", 22, 21)
'        End If
    Else
        strҽ����� = "24"
    End If
    str���� = ToVarchar(rsTemp("��Ժ����"), 30)
    strҽ�� = ToVarchar(rsTemp("����ҽʦ"), 20)
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    If IsNull(rsTemp!��Ժ����) Then
        str��Ժ���� = ""
    Else
        str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    End If
    
    'Modified by ZYB 2004-05-10
    '���ýӿڸ��£�סԺ��(�����)|���±�־|ҽ�����|����|ҽ��|��Ժ����|��Ժ���|��Ժ����|ȷ�Ｒ������
                  '|��Ժԭ��|������|����֢
    StrInput = "03|" & GetIdentify(lng����ID, lng��ҳID) & "|1110111001|" & strҽ����� & "|" & str���� & _
               "|" & strҽ�� & "|" & str��Ժ���� & "|" & str��� & "|" & str��Ժ���� & "|" & str�������� & "|||" & str����֢
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֤��','''" & str�������� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'��������','''" & str�������� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'����֢','''" & str����֢ & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���֢")
    
    ���³�Ժ����_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ����ҽ����Ժ_����(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str˳��� As String) As Boolean
'���ܣ����²��˵ĳ�Ժ����������������������ʱ���߻����
    Dim StrInput As String, arrOutput  As Variant
    
    On Error GoTo errHandle
    
    '���ýӿ�
    StrInput = "99|" & str˳��� & "|" & ToVarchar(UserInfo.����, 20)
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    gstrSQL = "ZL_������ҳ_����ҽ����Ժ(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ժ")
    
    ����ҽ����Ժ_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long, ByVal strҽ���� As String, _
        Optional ByVal bln���ò�ѯ As Boolean = False, Optional ByRef strAdvance As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim cn�ϴ� As New ADODB.Connection, rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    
    Dim str������Ϣ As String
    Dim StrInput As String, arrOutput   As Variant, arrTemp As Variant
    Dim cur�������� As Double
    Dim str�ܽ��ҽԺ As String, str�ܽ��ҽ�� As String, str������ϸ��ˮ�� As String
    Dim strҽ�� As String, datCurr As Date, intMsg As Integer
    Dim bln������ As Boolean, bln��;����ֻ�������ϴ����� As Boolean
    Dim bln����Ԥ���� As Boolean
    
    Dim str��Ժ���� As String, str�ϴν������� As String
    
    On Error GoTo errHandle
    If strAdvance = "" Then strAdvance = "0"
    bln����Ԥ���� = (Val(Split(strAdvance, ";")(0)) = 1)
    strAdvance = ""         '�Դ˲����ĸ�ֵ����������Ԥ������ɺ���ʾ��ǰ̨�����Դ˴�����Ϊ��
    
    mlng����ID = 0         '��ʼ����ֻҪһѡ���ˣ��ͻ���ñ����̣�Ҳ�ͻ����0
    mstr������ֹ���� = ""
    pre_Balance.HIS����ҽԺ��֧ = False

    If rsExse.RecordCount = 0 Then
        strAdvance = MessageInfo("�ò���û���з������ã��޷����н��������", bln����Ԥ����)
        Exit Function
    End If
    rsExse.MoveFirst
    
    datCurr = zlDatabase.Currentdate
    With g��������
        .����ID = rsExse("����ID")
        
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", CLng(rsExse("����ID")))
        If IsNull(rsTemp("��ҳID")) = True Then
            strAdvance = MessageInfo("ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", bln����Ԥ����)
            Exit Function
        End If
        .��ҳID = rsTemp("��ҳID")
        .��� = Int(Format(datCurr, "yyyy"))
        
        '�ж���;�����Ƿ�ֻ�������ϴ�����
        bln��;����ֻ�������ϴ����� = True
        gstrSQL = " Select Nvl(����ֵ,1) AS ����ֵ From ���ղ��� Where ����=[1] And ������='��;����'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж���;�����Ƿ�ֻ�������ϴ�����", TYPE_������)
        If rsTemp.RecordCount <> 0 Then
            bln��;����ֻ�������ϴ����� = (rsTemp!����ֵ = 1)
        End If
    End With
    
    'Modified by ZYB 2004-05-10
    '��ȡ���˵Ļ�����Ϣ��������ڷ���ԭ������ʾ
    gstrSQL = "Select ҽ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ò��˵�ҽ����", TYPE_������, g��������.����ID)
    
    StrInput = "01|" & rsTemp!ҽ����
    If HandleBusiness(StrInput, arrTemp, bln����Ԥ����) = False Then Exit Function
    str������Ϣ = ""
    If Val(arrTemp(11)) <> 0 Then
        str������Ϣ = arrTemp(12)
        If Not bln����Ԥ���� Then MsgBox str������Ϣ, vbInformation, gstrSysName
    End If
    '���·�����Ϣ
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'������Ϣ','''" & str������Ϣ & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
    
    Screen.MousePointer = vbHourglass
    '1.2 �������˵���Ժʱ��
    ''������ֹ���ڣ�����ѳ�Ժ�����ǳ�Ժ���ڣ������Ǳ��ν��������ķ�������
    gstrSQL = "select ��Ժ����,nvl(��Ժ����,to_date('3000-01-01','yyyy-MM-dd')) as ��Ժ���� " & _
              "from ������ҳ where ����ID=[1] and ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", g��������.����ID, g��������.��ҳID)
    str��Ժ���� = Format(rsTemp!��Ժ����, "yyyy-MM-dd")
    If rsTemp("��Ժ����") = CDate("3000-01-01") Then
        g��������.��;���� = 1
        With rsExse
            Do While Not .EOF
                If Format(!����ʱ��, "yyyy-MM-dd") > mstr������ֹ���� Then mstr������ֹ���� = Format(!����ʱ��, "yyyy-MM-dd")
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
        End With
    Else
        g��������.��;���� = 0
        mstr������ֹ���� = Format(rsTemp("��Ժ����"), "yyyy-MM-dd")
    End If
    
    '1.3 ��������סԺ�ϴ��н�ʱ�䣬���Ϊ�գ��ϴ��н����ʱ��Ϊ��Ժ����
     gstrSQL = " Select Max(����ʱ��) AS �������� From סԺ���ü�¼" & _
               " Where ����Id=(" & _
               "    Select Max(ID) As ����ID" & _
               "    From ���˽��ʼ�¼" & _
               "    Where ����ID=[1] And ��¼״̬ = 1 " & _
               "    And �շ�ʱ��>[2]" & _
               "    And Nvl(���ӱ�־,0)<>9"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������סԺ�ϴ��н�ʱ��", g��������.����ID, CDate(str��Ժ����))
    If rsTemp.RecordCount = 0 Then
        str�ϴν������� = str��Ժ����
    Else
        If IsNull(rsTemp!��������) Then
            str�ϴν������� = str��Ժ����
        Else
            str�ϴν������� = Format(rsTemp!��������, "yyyy-MM-dd")
        End If
    End If
    
    '1.4 ��;��������������ֱ�����������������������Ժ�������1��
    '���ν��ʷ��������ķ������ڼ�ȥ�ϴν��ʷ��������ķ������ڣ�����סԺ����
    g��������.סԺ���� = DateDiff("d", CDate(str�ϴν�������), CDate(mstr������ֹ����))
    'If g��������.��;���� = 0 Then g��������.סԺ���� = g��������.סԺ���� + 1
    If g��������.סԺ���� < 1 Then g��������.סԺ���� = 1 '������һ��
    Call DebugTool("������ֹ���ڣ�" & mstr������ֹ���� & "���ϴν������ڻ���Ժ���ڣ�" & str�ϴν������� & "��סԺ���գ�" & g��������.סԺ����)
    
    Do Until rsExse.EOF
        cur�������� = cur�������� + rsExse("���")
        rsExse.MoveNext
    Loop
    cur�������� = Val(Format(cur��������, "#####0.00"))
    
    'ֻ�г�Ժ�����ʹ�ò��˷��ò�ѯ�����ã����ϴ�����δ�ϴ���ϸ����;����ֻ�����ϴ����ݽ��н���
    If g��������.��;���� = 0 Or bln���ò�ѯ Or Not bln��;����ֻ�������ϴ����� Then
        '����δ�ϴ���ϸ�������Ա����ϴ�����ϸ�����ϴ�����ϸ��
        gstrSQL = "Select A.ID,A.NO,A.��¼����,A.��¼״̬,A.���,A.����ID,A.��ҳID,A.����ʱ�� as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս��" & _
                  "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
                  "         ,C.��Ŀ����,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ,E.���,G.���� ���� " & _
                  "  From סԺ���ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C,������ҳ D,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
                  "  where A.����ID=" & lng����ID & " and A.��ҳID=" & g��������.��ҳID & _
                  "        and A.���ʷ���=1 and A.��¼����<>1 and Nvl(A.ʵ�ս��,0)<>0 and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 " & _
                  "        and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=[1]" & _
                  "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����=D.���� " & _
                  "        AND B.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) " & _
                  "  Order by A.����ʱ��,A.��¼����,Decode(A.��¼״̬,2,2,1)"
        Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "�������", TYPE_������)
'
'        '����ҽ����Ŀ����Ϣ������
'        Do While Not rs��ϸ.EOF
'            Call TrackRecordInsure(rs��ϸ!ID, rs��ϸ!�շ�ϸĿID)
'            rs��ϸ.MoveNext
'        Loop
        
        '������һ�����Ӵ����Դﵽ���ܵ�ǰ��������Ŀ���
        Set cn�ϴ� = GetNewConnection
        cn�ϴ�.Open
        
        intMsg = 0
        Call DebugTool("��ʼ�ϴ���ϸ")
        If rs��ϸ.RecordCount <> 0 Then rs��ϸ.MoveFirst
        Do Until rs��ϸ.EOF
            If rs��ϸ!��¼״̬ = 1 Then
                If Val(rs��ϸ!����) < 0 Or Val(rs��ϸ!�۸�) < 0 Then
                    '����ȡһ��������¼����ˮ�ţ���Ϊ������ˮ��
                    '��Ϊ��־���ñ�����Ϊ��
                    str������ϸ��ˮ�� = "������¼"
                Else
                    str������ϸ��ˮ�� = ""
                End If
            Else
                str������ϸ��ˮ�� = GetDetailSequence(rs��ϸ!NO, rs��ϸ!���, rs��ϸ!��¼����, rs��ϸ!��¼״̬)
            End If
            
            If rs��ϸ!��¼״̬ = 1 And str������ϸ��ˮ�� <> "" Then
                Call UploadNegative(rs��ϸ!����ID, rs��ϸ!��ҳID, rs��ϸ!ID, rs��ϸ!�շ�ϸĿID)
            Else
                If rs��ϸ!��¼״̬ <> 2 Then
                    strҽ�� = ToVarchar(IIf(IsNull(rs��ϸ("ҽ��")), UserInfo.����, rs��ϸ("ҽ��")), 20)
                    
                    StrInput = "04|" & GetIdentify(lng����ID, g��������.��ҳID)
                    StrInput = StrInput & "|" & rs��ϸ("NO") & "_" & rs��ϸ("��¼����")
                    StrInput = StrInput & "|" & Format(rs��ϸ("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
                    StrInput = StrInput & "|" & ToVarchar(rs��ϸ("��Ŀ����"), 10) '���ı���
                    StrInput = StrInput & "|" & ToVarchar(rs��ϸ("����"), 20) 'ҽԺ����
                    StrInput = StrInput & "|" & ToVarchar(rs��ϸ("����"), 50)     '��Ŀ����
                    StrInput = StrInput & "|" & Format(rs��ϸ("�۸�"), "0.0000")      '����
                    StrInput = StrInput & "|" & Format(rs��ϸ("����"), "0.00")        '����
                    StrInput = StrInput & "|" & IIf(rs��ϸ("�Ƿ���") = 1, 1, 0)     '�����־
                    StrInput = StrInput & "|" & strҽ��                               'ҽ��
                    StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 20)          '������
                    StrInput = StrInput & "|" & ToVarchar(rs��ϸ("���㵥λ"), 20)     '��λ
                    StrInput = StrInput & "|" & ToVarchar(rs��ϸ("���"), 14)         '���
                    StrInput = StrInput & "|" & ToVarchar(rs��ϸ("����"), 20)         '����
                    StrInput = StrInput & "|" & str������ϸ��ˮ��                     '������ϸ��ˮ��
                    StrInput = StrInput & "|" & Format(rs��ϸ("ʵ�ս��"), "#####0.0000")         '���
                Else
                    '���ʱ�ͼ��ʵ��������ֳ�������Ҫһ��һ�ʳ���
                    StrInput = "99|" & str������ϸ��ˮ�� & "|" & ToVarchar(UserInfo.����, 20)
                End If
                
                'Modified by ZYB 20040511 ��������
                '�������ڸ������ʣ������ĳ�����¼����Ϊ���������Ǳ����ڽӿ����ƣ��϶�������ȥ����˱��������������ĳ�����¼���ϴ�
                If HandleBusiness(StrInput, arrOutput, bln����Ԥ����) = False Then
                    '�����ϴ�ʧ��
                    If Not bln����Ԥ���� Then
                        If MsgBox("[����ID:" & lng����ID & "]���ݡ�" & rs��ϸ("NO") & "����" & rs��ϸ("����") & "�����ϴ�ʧ�ܣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                        If intMsg = 0 Then
                            If MsgBox("�ϴ�����ʧ�ܣ��Ƿ�ֹͣ�����ϴ���ֱ�ӽ��н��ʣ�", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                                intMsg = 1
                                Exit Do
                            Else
                                intMsg = -1
                            End If
                        End If
                    Else
                        strAdvance = MessageInfo("[����ID:" & lng����ID & "]���ݡ�" & rs��ϸ("NO") & "����" & rs��ϸ("����") & "�����ϴ�ʧ��", bln����Ԥ����)
                        Exit Function
                    End If
                Else
                    '�����ϴ��ɹ������ϱ��
                    If rs��ϸ!��¼״̬ <> 2 Then
                        gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rs��ϸ("ID") & "," & Val(arrOutput(2)) * rs��ϸ("����") & ",'" & arrOutput(1) & "')"
                    Else
                        gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rs��ϸ("ID") & ")"
                    End If
                    cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
                End If
            End If
            
            rs��ϸ.MoveNext
        Loop
    End If
    
    '����ǳ�Ժ����ǰ��������㣬����Ƿ����δ�ϴ��ķ�����ϸ
    If g��������.��;���� = 0 Then
        If Not frm��ѯδ�ϴ�������ϸ.ShowME(lng����ID, g��������.��ҳID, TYPE_������) Then Exit Function
    End If
    
    '����Ԥ����
    Call DebugTool("׼������Ԥ����")
    '�����ڽ��г�Ժ����ǰ���������ʱ��������ֹ���ڴ���ǰ���ڣ�����ǽ�����;�������������򴫷��ü�¼�����ĵǼ����ڣ������ں˶����ߵķ����Ƿ�һ��
    'ԭ�򣬳�Ժ����ʱ���ݵķ�����ֹ������Ч��������Ȼͳ�Ƶ����з��ý��еĽ���
    If g��������.��;���� = 0 Then
        mstr������ֹ���� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    End If
    
    'סԺ(����)��|��ֹ����|סԺ���գ�������������ԵǼ�ʱ��Ϊ׼��ȡ���ĵǼ�ʱ����Ϊ��ֹ����
    StrInput = "06|" & GetIdentify(lng����ID, g��������.��ҳID) & "|" & Format(mstr������ֹ����, "yyyyMMdd") & "|" & g��������.סԺ���� & "|" & Format(cur��������, "#0.00")
    If HandleBusiness(StrInput, arrOutput, bln����Ԥ����) = False Then Exit Function
    
    pre_Balance.cur�����ʻ� = Val(arrOutput(2))
    pre_Balance.curͳ��֧�� = Val(arrOutput(1))
    pre_Balance.cur���ͳ�� = Val(arrOutput(5))
    pre_Balance.cur����Ա���� = Val(arrOutput(3))
    pre_Balance.cur����Ա���� = Val(arrOutput(6))
    pre_Balance.curҽԺ��֧ = 0
    '���½ӿ�����ҽԺ��֧ʱ������ֵ����ʱ�����ܶ�Ƚ�
    If UBound(arrOutput) > 7 Then
        pre_Balance.curҽԺ��֧ = Val(arrOutput(8))
    End If
    If UBound(arrOutput) > 8 Then       '20101028������������
        pre_Balance.cur�������� = Val(arrOutput(9))
        pre_Balance.curͳ��֧�� = pre_Balance.curͳ��֧�� - pre_Balance.cur��������
    End If
    
    '���没�˸����ʻ����
    mstrҽ���� = strҽ����
    mdbl��� = Val(arrOutput(7)) + Val(arrOutput(2))
    
    '������ʱ���ݣ�Ϊ���������׼��
    With g��������
        .�������ý�� = cur��������
    End With
    
    str�ܽ��ҽԺ = Format(cur��������, "#####0.00")
    '�����ص��ֽ�֧��Ӧ�ü�ȥ����Ա�������֣��������յ��ֽ�֧����
    str�ܽ��ҽ�� = Format(pre_Balance.curͳ��֧�� + pre_Balance.cur�����ʻ� + pre_Balance.cur����Ա���� + pre_Balance.cur���ͳ�� + pre_Balance.curҽԺ��֧ + pre_Balance.cur�������� + Val(arrOutput(4)), "#####0.00")
    If str�ܽ��ҽԺ <> str�ܽ��ҽ�� Then
        If Not bln����Ԥ���� Then
            If MsgBox("ҽԺ�ķ����ܽ��(" & str�ܽ��ҽԺ & ")��ҽ�����ĵķ����ܶ�(" & str�ܽ��ҽ�� & ")���ȣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            If MsgBox("�����ֺ;������������ԭ�򣬿��ܻ���֣�ҽԺ��ҽ�����ĵ���ϸ�ܶ�һ�£��������ܶ��" & vbCrLf & _
            "����㡰�ǡ����������ּ����񲡴������ò�������㷽ʽ��ҽԺ��֧���У��㡰������ͨ���˽��㴦��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                bln������ = True
            End If
        Else
            strAdvance = "[����ID:" & lng����ID & "]HIS�ܷ�����ҽ���ܷ��ò��ȣ�" & vbCrLf & _
                "HIS:" & str�ܽ��ҽ�� & Space(5) & "ҽ��:" & str�ܽ��ҽ��
        End If
    End If
    
    סԺ�������_���� = "ҽ������;" & pre_Balance.curͳ��֧�� & ";0"
    If pre_Balance.cur�����ʻ� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|�����ʻ�;" & pre_Balance.cur�����ʻ� & ";0" '�������޸ĸ����ʻ�
    End If
    If pre_Balance.cur���ͳ�� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|���ͳ��;" & pre_Balance.cur���ͳ�� & ";0"
    End If
    If pre_Balance.cur����Ա���� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|����Ա����;" & pre_Balance.cur����Ա���� & ";0"
    End If
    If pre_Balance.cur����Ա���� <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|����Ա����;" & pre_Balance.cur����Ա���� & ";0"
    End If
    If pre_Balance.cur�������� <> 0 Then      '20101028������������
        סԺ�������_���� = סԺ�������_���� & "|��������;" & pre_Balance.cur�������� & ";0"
    End If
    
    pre_Balance.curҽԺ��֧ = 0
    If UBound(arrOutput) > 7 Then
        pre_Balance.curҽԺ��֧ = Val(arrOutput(8))
    End If
    If bln������ Then
        '�ӿڷ��������ҷ����㣬������HIS�ܶ��ȥYB�ܶ������ҽԺ��֧
        If Not (UBound(arrOutput) > 7) Then
            '��Ҫ�ҷ����㣨��Ϊ�ӿ�δ����ҽԺ��֧��˵����û�����У���˼���ҽ���ܶ�ʱ���ܼ�ҽԺ��֧������Ҫ����
            pre_Balance.HIS����ҽԺ��֧ = True
            pre_Balance.curҽԺ��֧ = cur�������� - (pre_Balance.curͳ��֧�� + pre_Balance.cur�����ʻ� + pre_Balance.cur����Ա���� + pre_Balance.cur���ͳ�� + Val(arrOutput(4)))
        End If
    End If
    If pre_Balance.curҽԺ��֧ <> 0 Then
        סԺ�������_���� = סԺ�������_���� & "|ҽԺ��֧;" & pre_Balance.curҽԺ��֧ & ";0"
    End If
    
    '����Ԥ������ڽ���ʱ�ٱȽ�һ�Σ�������ֲ��
    With g��������
        .ͳ�ﱨ����� = pre_Balance.curͳ��֧��       '1
        .�����ʻ�֧�� = pre_Balance.cur�����ʻ�       '2
        .�ۼƽ���ͳ�� = pre_Balance.cur����Ա����     '3
        .ȫ�Էѽ�� = Val(arrOutput(4))   '4
        .����ͳ���� = pre_Balance.cur���ͳ��       '5
        .�ۼ�ͳ�ﱨ�� = Val(arrOutput(6)) '6
    End With
    
    mlng����ID = lng����ID  '��ʾ�ò����Ѿ��������������
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����ID As Long, Optional ByRef strAdvance As String) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim str����Ա As String, lng�����־ As Long
    Dim str���㷽ʽ As String
    Dim curͳ��֧�� As Double, cur�����ʻ� As Double
    Dim cur���ͳ�� As Double, cur����Ա���� As Double, cur����Ա���� As Double, curҽԺ��֧ As Double, cur�������� As Double, cur�����ʻ���� As Double
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, datCurr As Date, strNO As String
    Dim strFormat As String
    Dim str���� As String, str����֢ As String, str������Ϣ As String, str��ע As String
    Dim blnOld As Boolean, blnRevise As Boolean '�Ƿ���Ҫ��дУ���ֶΣ��Ƿ���ҪУ��������
    Dim bln�ʻ�֧�� As Boolean
    Dim int�ʻ�֧����ʽ As Integer              '0-֧��;1-סԺѯ��;2-����ѯ��;3-��֧��
    
    '��������������
    Dim str���ֱ��� As String, bln������ As Boolean, int��� As Integer    '0-δ����н�;1-����н���...
    Const int��ٿ�ʼ�н� As Integer = 40
    Const int��ٽ����н� As Integer = 50
    
    If mlng����ID <> lng����ID Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣"
        Exit Function
    End If
    
    On Error GoTo errHandle
    Call DebugTool("����סԺ����")
    '��ȡ�α����˵Ĳ���֢��������Ϣ
    gstrSQL = "Select ����֤�� As ���ֱ���,��������,����֢,������Ϣ,Nvl(��ٱ�־,0) AS ��ٱ�־" & _
        " From �����ʻ� " & _
        " Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�α����˵Ĳ���֢��������Ϣ", lng����ID, TYPE_������)
    If Not rsTemp.EOF Then
        str���� = Nvl(rsTemp!���ֱ���)
        str���ֱ��� = Nvl(rsTemp!���ֱ���)
        If str���� <> "" Then str���� = "[" & str���� & "]"
        str���� = str���� & Nvl(rsTemp!��������)
        str����֢ = Nvl(rsTemp!����֢)
        str������Ϣ = Nvl(rsTemp!������Ϣ)
        int��� = rsTemp!��ٱ�־       '���ݱ����ʻ��е���ٱ�־���жϵ�ǰ�Ƿ�������У���ˣ���������������Ҫ���øñ�־
    End If
    str��ע = str���� & "||" & str����֢ & "||" & str������Ϣ
    
    '�ж��Ƿ��ǵ����ֽ���
    If Trim(str���ֱ���) <> "" Then
        '������
        If mcnYB.State = 0 Then
            If Not OpenDatabase Then Exit Function
        End If
        gstrSQL = "Select 1 From BZML Where bzfl=5 And upper(BZBM)='" & UCase(str���ֱ���) & "'"
        Call OpenRecordset_OtherBase(rsTemp, "��ȡ���ַ��࣬���ж��Ƿ��ǵ�����", gstrSQL, mcnYB)
        bln������ = (rsTemp.RecordCount <> 0)
    End If
    
    '������ʻ�֧�����
    gstrSQL = "Select Nvl(��Ԥ��,0) as ��� From ����Ԥ����¼ Where ���㷽ʽ='�����ʻ�' And ��¼����=2 And ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺ����", lng����ID)
    If Not rsTemp.EOF Then cur�����ʻ� = rsTemp("���")
    
    'ѯ���Ƿ�ʹ�ø����ʻ�֧��
    gstrSQL = "Select Nvl(����ֵ,0) AS ����ֵ From ���ղ��� Where ����=[1] And ������='�����ʻ�'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ղ��������ݾ���¼����ϡ�", TYPE_������)
    If rsTemp.RecordCount <> 0 Then
        int�ʻ�֧����ʽ = rsTemp!����ֵ
    End If
    If int�ʻ�֧����ʽ = 0 Or int�ʻ�֧����ʽ = 2 Then
        bln�ʻ�֧�� = True
    ElseIf int�ʻ�֧����ʽ = 1 Then
        If MsgBox("���ʱ���Ҫ���и����ʻ�֧����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            bln�ʻ�֧�� = False
        Else
            bln�ʻ�֧�� = True
        End If
    Else
        bln�ʻ�֧�� = False
    End If
    
    '���ý���
    With g��������
        If .��;���� = 1 Then
            '�����������=5�����񲡣�������ʾ�Ƿ������ٿ�ʼ�н�
            lng�����־ = 10
            If bln������ Then
                If int��� = 0 Then
                    If MsgBox("�ò������ڵ����֣��Ƿ���С���ٿ�ʼ�нᡱ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        lng�����־ = int��ٿ�ʼ�н�
                        int��� = 1
                    End If
                Else
                    If MsgBox("�ò������ڵ����֣����ν��С���ٽ����нᡱ������򲻽��н��㣿", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        lng�����־ = int��ٽ����н�
                        int��� = 0
                    Else
                        Exit Function
                    End If
                End If
            End If
    '            If MsgBox("�ò����Ƿ����ת��ͥ�������㣿", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
    '                lng�����־ = 20 '��Ժת��ͥ����
    '            Else
'                    lng�����־ = 10 '��;����
    '            End If
        Else
            lng�����־ = 0 '��������
        End If
        
        StrInput = "05|" & GetIdentify(lng����ID, .��ҳID) & "|" & lng�����־ & "|" & g��������.סԺ���� & "|" & UserInfo.���� & _
                   "|" & IIf(bln�ʻ�֧��, "0", "1") & "|" & Format(mstr������ֹ����, "yyyyMMdd") & "|" & Format(g��������.�������ý��, "#0.00")
    End With
    
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    '��д�����
    Call DebugTool("��д�����¼")
    datCurr = zlDatabase.Currentdate
    cur�����ʻ� = Val(arrOutput(3))
    curͳ��֧�� = Val(arrOutput(2))
    cur����Ա���� = Val(arrOutput(4))
    cur���ͳ�� = Val(arrOutput(6))
    cur����Ա���� = Val(arrOutput(7))
    If UBound(arrOutput) > 9 Then           '20101028������������
        cur�������� = Val(arrOutput(10))
        curͳ��֧�� = curͳ��֧�� - cur��������
        cur�����ʻ���� = Val(arrOutput(11))
        gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'���������������','''" & cur�����ʻ���� & "''')"      '20101028������������
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������������������")
    End If
    If pre_Balance.HIS����ҽԺ��֧ = False Then
        If UBound(arrOutput) > 8 Then
            curҽԺ��֧ = Val(arrOutput(9))
        End If
    Else
        '�ҷ��ٴΰ�������м���
        curҽԺ��֧ = g��������.�������ý�� - (cur�����ʻ� + curͳ��֧�� + cur����Ա���� + cur���ͳ�� + Val(arrOutput(5)))
    End If
    
    '�Ƚ���ʽ�������Ƿ������������һ�£���һ������ҪУ��
    If Not (cur�����ʻ� = pre_Balance.cur�����ʻ� And curͳ��֧�� = pre_Balance.curͳ��֧�� And _
        cur����Ա���� = pre_Balance.cur����Ա���� And cur���ͳ�� = pre_Balance.cur���ͳ�� And _
        cur����Ա���� = pre_Balance.cur����Ա���� And curҽԺ��֧ = pre_Balance.curҽԺ��֧ And cur�������� = pre_Balance.cur��������) Then
        
        blnRevise = True
        
        str���㷽ʽ = "�����ʻ�|" & cur�����ʻ�
        If curͳ��֧�� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curͳ��֧��
        If cur���ͳ�� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||���ͳ��|" & cur���ͳ��
        If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
        If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
        If curҽԺ��֧ <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽԺ��֧|" & curҽԺ��֧
        If cur�������� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||��������|" & cur��������
        If str���㷽ʽ <> "" Then
            #If gverControl < 2 Then
                blnOld = True
                gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
            #Else
                strAdvance = str���㷽ʽ
                gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
            #End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
        End If
    End If
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
    cur�ʻ������ۼ� = mdbl���
    cur�ʻ�֧���ۼ� = cur�����ʻ�
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_������ & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & _
        cur����ͳ���ۼ� + curͳ��֧�� & "," & _
        curͳ�ﱨ���ۼ� + curͳ��֧�� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_������ & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,NULL,0," & g��������.�������ý�� & ",0,0," & _
        curͳ��֧�� & "," & curͳ��֧�� & ",0,0," & cur�����ʻ� & ",'" & arrOutput(1) & "'," & g��������.��ҳID & "," & g��������.��;���� & ",'" & str��ע & "'" & _
        IIf(blnOld, "", IIf(blnRevise, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս������
    gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & curͳ��֧�� & "," & curͳ��֧�� & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '��Ժ����ʱ����סԺ����ŵ�������Ϊ��
    If g��������.��;���� = 0 Then
        Dim strסԺ����� As String
        If GetMode(lng����ID, g��������.��ҳID, strסԺ�����) = False Then
            '��סԺ�������Ϊ�գ�Ϊ�´�סԺ��׼��
            gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'סԺ�����','''" & "" & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��סԺ�������Ϊ��")
        End If
    End If
    
    '������ٽ����־
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'��ٱ�־','''" & int��� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������н��־")
    
    #If gverControl < 2 Then
        Call frm������Ϣ.ShowME(lng����ID)
    #End If
    
    סԺ����_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant
    Dim lng����ID As Long, str��ˮ�� As String
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim lng����ID As Long
    Dim int��� As Integer
    Dim curDate As Date
        
    On Error GoTo errHandle
    int��� = 0
    curDate = zlDatabase.Currentdate
    
    '�˷�
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "select * from ���ս����¼ where ����=2 and ����=[1] and ��¼ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", TYPE_������, lng����ID)
    If rsTemp.EOF = True Then
        Err.Raise 9000 + VbMsgBoxStyle.vbInformation, gstrSysName, "ԭ���ݵ�ҽ����¼�����ڣ��������ϡ�"
        Exit Function
    End If
    lng����ID = rsTemp!����ID
    If CanסԺ�������(rsTemp("����ID"), rsTemp("��ҳID")) = False Then Exit Function
    
    str��ˮ�� = rsTemp("֧��˳���")
    
    StrInput = "99|" & str��ˮ�� & "|" & ToVarchar(UserInfo.����, 20)
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_������, rsTemp("����ID"), Year(curDate), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
            
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & rsTemp("����ID") & "," & TYPE_������ & "," & Year(curDate) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� - rsTemp("����ͳ����") & "," & _
        curͳ�ﱨ���ۼ� - rsTemp("ͳ�ﱨ�����") & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_������ & "," & rsTemp("����ID") & "," & _
        Year(curDate) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� - rsTemp("�����ʻ�֧��") & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & rsTemp("�������ý��") * -1 & ",0,0," & _
        rsTemp("����ͳ����") * -1 & "," & rsTemp("ͳ�ﱨ�����") * -1 & ",0,0," & _
        rsTemp("�����ʻ�֧��") * -1 & ",'" & str��ˮ�� & "'," & rsTemp("��ҳID") & "," & rsTemp("��;����") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ��")
    
    '������ٽ����־
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'��ٱ�־','''" & int��� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������н��־")

    סԺ�������_���� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ������Ϣ_����(ByVal lngErrCode As Long) As String
'���ܣ����ݴ���ŷ��ش�����Ϣ

End Function

Public Function ҽԺ����_����() As String
'���ܣ��õ�ҽԺ��ҽ������
    Dim StrInput As String, arrOutput As Variant
    
    On Error GoTo errHandle
    
    StrInput = "11"
    If HandleBusiness(StrInput, arrOutput) = False Then Exit Function
    ҽԺ����_���� = arrOutput(1)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function HandleBusiness(ByVal StrInput As String, varOut As Variant, Optional ByVal bln����Ԥ���� As Boolean = False) As Boolean
'���ܣ�����ҽ������������ҵ����
    Dim strInfo As String '����ǰ�÷������ķ���ֵ
    Dim lngReturn As Long
    Dim varArray As Variant, lngCount As Long
    
    On Error Resume Next
    varOut = ""
    Screen.MousePointer = vbHourglass
    strInfo = Space(1024)
    lngReturn = dy_Business_Handle(StrInput, strInfo)
'    lngReturn = gobj�����ж���.dy_Business_Handle(StrInput, strInfo)
    If Err <> 0 Or lngReturn = -1 Then
        varArray = Split(strInfo, "|")
        
        If UBound(varArray) > 0 Then
            strInfo = "ҽ���ӿڵ���ʧ�ܡ�" & vbCrLf & varArray(1)
        Else
            strInfo = "ҽ���ӿڵ���ʧ�ܡ�" & vbCrLf & strInfo
        End If
        Screen.MousePointer = vbDefault
        If Not bln����Ԥ���� Then MsgBox strInfo, vbExclamation, gstrSysName
        Exit Function
    End If
    strInfo = TruncZero(strInfo)
    
    varArray = Split(strInfo, "|")
    If varArray(0) = "-1" Then
        'ҵ�����ʧ��
        If UBound(varArray) > 0 Then
            strInfo = "ҽ���ӿڳ��־��档" & vbCrLf & varArray(1)
        Else
            strInfo = "ҽ��ҵ����ʧ�ܡ�"
        End If
        
        Screen.MousePointer = vbDefault
        If Not bln����Ԥ���� Then MsgBox strInfo, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '���׳ɹ�
    varOut = Split(strInfo, "|")
    
    HandleBusiness = True
    Screen.MousePointer = vbDefault
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

Public Function �۸��ж�_����(ByVal dblҽԺ As Double, ByVal dblҽ�� As Double, ByVal str�޼۷�ʽ As String, _
                              ByVal bln�ؼ� As Boolean, ByVal dbl�ؼ� As Double) As Boolean
'���ܣ��ж�ҽԺ�ļ۸��Ƿ񳬹�ҽ���涨�ĵ���
    Dim strҽԺ��� As String
    
    On Error GoTo errHandle
    
    If InStr(str�޼۷�ʽ, "����") > 0 Then
        strҽԺ��� = Get���ղ���_����("ҽԺ�ȼ�")
        '�����ı�׼�۸�Ϊ����ҽԺ������޼ۣ�����ҽԺ������޼��ڴ˻����Ͽ����ϸ�10%��һ��ҽԺ������޼��ڴ˻������µ�5%
        
        Select Case strҽԺ���
            Case "����"
                dblҽ�� = dblҽ�� * 1.1
            Case "һ��"
                dblҽ�� = dblҽ�� * 0.95
        End Select
    End If
    
    If bln�ؼ� = True And dbl�ؼ� > dblҽ�� Then
        '����ʹ���ؼ�
        dblҽ�� = dbl�ؼ�
    End If
    
    If dblҽԺ > dblҽ�� Then
        If MsgBox("ҽԺ����" & Format(dblҽԺ, "0.000") & " ����ҽ�����ĺ�׼�ļ۸�" & Format(dblҽ��, "0.000") & "���Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    �۸��ж�_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ���ʴ���_����(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String, Optional ByVal lng����ID As Long = 0) As Boolean
'����:�ϴ��²����ļ�����ϸ��ҽ������
'����:  str���ݺ�   NO
'       int����     ��¼����
'       str��Ϣ    �����������������ѣ�����ǰ̨������ɣ����ⳤʱ���������
'       lng����ID  Ĭ��Ϊ0����ʾ�������ŵ��ݣ�����Ϊ������ָ�����˵ġ�����Ҫ����Ϊҽ���ڱ�����ʵ�ʱ���Ƿֲ������ύ���ݶ�����һ���ύ��
'����:
    Dim rsTemp As New ADODB.Recordset
    Dim rsTest As New ADODB.Recordset
    Dim cn�ϴ� As New ADODB.Connection
    Dim StrInput As String, arrOutput, arrTemp  As Variant, curͳ���� As Currency
    Dim strҽ�� As String, str������ As String
    Dim col���� As New Collection, lngPre����ID As Long, var���� As Variant, bln�ɹ� As Boolean
    Dim str������ϸ��ˮ�� As String, str������Ϣ As String, str����������Ϣ As String
    '��ע�⣺����ҽ�����ڼ��ʵ�������ٵ��ô�����̵ġ�
    
    On Error GoTo errHandle
    
    Set cn�ϴ� = GetNewConnection
    cn�ϴ�.Open
    
    '�������ŵ��ݵķ�����ϸ
    gstrSQL = "Select A.ID,A.NO,A.����ID,A.��ҳID,A.����ʱ�� as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս�� " & _
              "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,C.��Ŀ����,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ,E.���,G.���� ���� " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C,������ҳ D,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
              "  where A.NO=[1] and A.��¼����=[2] and A.��¼״̬=1 And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.ʵ�ս��,0)<>0 " & _
              "        and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=" & TYPE_������ & IIf(lng����ID = 0, "", " and A.����ID=[3]") & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����=D.���� " & _
              "        AND B.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) " & _
              "  Order by A.����ID,A.����ʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ʴ���", str���ݺ�, int����, lng����ID)
    
    '��ǰ�ļ�鲻��Ҫ���������Ǳ�����ϴ������Ҳû������
'    Do While Not rsTemp.EOF
'        If Val(rsTemp!����) < 0 Or Val(rsTemp!�۸�) < 0 Then
'            '����ȡһ��������¼����ˮ�ţ���Ϊ������ˮ��
'            str������ϸ��ˮ�� = GetSequence(rsTemp!����ID, rsTemp!��ҳID, rsTemp!�շ�ϸĿID)
'            If Trim(str������ϸ��ˮ��) = "" Then
'                MsgBox "û���ҵ����Գ����ļ�¼��[" & rsTemp!���� & "]" & rsTemp!����, vbInformation, gstrSysName
'                Exit Function
'            End If
'        Else
'            str������ϸ��ˮ�� = ""
'        End If
'        rsTemp.MoveNext
'    Loop
'    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
'
'    '����ҽ����Ŀ����Ϣ������
'    Do While Not rsTemp.EOF
'        Call TrackRecordInsure(rsTemp!ID, rsTemp!�շ�ϸĿID)
'        rsTemp.MoveNext
'    Loop
    
    '���з�����ϸ�Ĵ���
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do Until rsTemp.EOF
        If Val(rsTemp!����) < 0 Or Val(rsTemp!�۸�) < 0 Then
            '����ȡһ��������¼����ˮ�ţ���Ϊ������ˮ��
            str������ϸ��ˮ�� = "������¼"
        Else
            str������ϸ��ˮ�� = ""
        End If
        
        If str������ϸ��ˮ�� = "" Then
            Call DebugTool("׼���ϴ�����������ϸ")
            strҽ�� = ToVarchar(IIf(IsNull(rsTemp("ҽ��")), UserInfo.����, rsTemp("ҽ��")), 20)
            str������ = ToVarchar(IIf(IsNull(rsTemp("����Ա����")), UserInfo.����, rsTemp("����Ա����")), 20)
            
            StrInput = "04|" & GetIdentify(rsTemp("����ID"), rsTemp("��ҳID"))
            StrInput = StrInput & "|" & rsTemp("NO") & "_" & int����
            StrInput = StrInput & "|" & Format(rsTemp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
            StrInput = StrInput & "|" & ToVarchar(rsTemp("��Ŀ����"), 10)     '���ı���
            StrInput = StrInput & "|" & ToVarchar(rsTemp("����"), 20)         'ҽԺ����
            StrInput = StrInput & "|" & ToVarchar(rsTemp("����"), 50)         '��Ŀ����
            StrInput = StrInput & "|" & Format(rsTemp("�۸�"), "0.0000")      '����
            StrInput = StrInput & "|" & Format(rsTemp("����"), "0.00")        '����
            StrInput = StrInput & "|" & IIf(rsTemp("�Ƿ���") = 1, 1, 0)     '�����־
            StrInput = StrInput & "|" & strҽ��                               'ҽ��
            StrInput = StrInput & "|" & str������                             '������
            StrInput = StrInput & "|" & ToVarchar(rsTemp("���㵥λ"), 20)     '��λ
            StrInput = StrInput & "|" & ToVarchar(rsTemp("���"), 14)         '���
            StrInput = StrInput & "|" & ToVarchar(rsTemp("����"), 20)         '����
            StrInput = StrInput & "|" & str������ϸ��ˮ��                     '������ϸ��ˮ��
            StrInput = StrInput & "|" & Format(rsTemp("ʵ�ս��"), "#####0.0000")         '���
            
            If HandleBusiness(StrInput, arrOutput) = False Then
                If bln�ɹ� = True Then
                    MsgBox "�����ϴ���;�������󣬲����Ѿ������Ѿ��ϴ�������Ԥ���㴦���ʣ�����ݵ��ϴ���", vbInformation, gstrSysName
                Else
                    MsgBox "�����ϴ���������û�гɹ��ϴ��ļ�¼������Ԥ���㴦���ʣ�����ݵ��ϴ���", vbInformation, gstrSysName
                End If
                ���ʴ���_���� = True
                Exit Function
            End If
            Call AddMessage(str��Ϣ, arrOutput, rsTemp("����"), rsTemp("�۸�"))  '���Բ�����������Ϣ
            
            '�ڷ��ü�¼�ϴ��ϱ�ǣ�˵���Ѿ��ϴ��������淵�صĽ��
            If arrOutput(3) = 2 Then
                'δͨ������
                curͳ���� = 0
            Else
                '��׼���� * ����
                curͳ���� = Val(arrOutput(2)) * rsTemp("����")
            End If
            
            '������¼�ϴ�������¼������ˮ��
            gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rsTemp("ID") & "," & curͳ���� & ",'" & arrOutput(1) & "')"
            cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
            bln�ɹ� = True
        Else
            str����������Ϣ = UploadNegative(rsTemp!����ID, rsTemp!��ҳID, rsTemp!ID, rsTemp!�շ�ϸĿID)
            If str����������Ϣ <> "" Then str��Ϣ = str��Ϣ & str����������Ϣ & vbCrLf
            bln�ɹ� = (str����������Ϣ = "")
        End If
        
        If lngPre����ID <> rsTemp("����ID") Then '�ж�ʱû�п�����ҳID������Ϊͬһ���˲�����ͬʱ������סԺ����ϸ
            'Modified by ZYB 2004-05-10
            '��ȡ���˵Ļ�����Ϣ��������ڷ���ԭ������ʾ
            Call DebugTool("��ȡ���˷�����Ϣ������")
            gstrSQL = "Select ҽ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
            Set rsTest = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ò��˵�ҽ����", TYPE_������, CLng(rsTemp!����ID))
            
            StrInput = "01|" & rsTest!ҽ����
            If HandleBusiness(StrInput, arrTemp) Then
                str������Ϣ = ""
                If Val(arrTemp(11)) <> 0 Then
                    str������Ϣ = arrTemp(12)
                    MsgBox str������Ϣ, vbInformation, gstrSysName
                End If
                '���·�����Ϣ
                gstrSQL = "ZL_�����ʻ�_������Ϣ(" & rsTemp!����ID & "," & TYPE_������ & ",'������Ϣ','''" & str������Ϣ & "''')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���·�����Ϣ")
            End If
            
            '���Ѿ��ϴ��Ĳ�����Ϣ��¼��������Ϊ���ʱ��Ƕಡ�˵ģ�
            col����.Add rsTemp("����ID") & "_" & rsTemp("��ҳID")
            lngPre����ID = rsTemp("����ID")
        End If
        
        rsTemp.MoveNext
    Loop
    
    If str��Ϣ <> "" Then
        str��Ϣ = "���˷�����ϸ��������еõ�ҽ���������·�����Ϣ����Ŀǰ�����Ѿ����档" & vbCrLf & "����кβ��ף������ѡ�����ϸõ��ݡ�" & vbCrLf & vbCrLf & str��Ϣ
    End If
        
    ���ʴ���_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��������_����(ByVal str���ݺ� As String, ByVal int���� As Integer, str��Ϣ As String) As Boolean
'����:�����Ѿ��ϴ���ҽ�����ĵļ�����ϸ
'����:  str���ݺ�   NO
'       int����     ��¼����
'       str��Ϣ    �����������������ѣ�����ǰ̨������ɣ����ⳤʱ���������
'����:
    
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, arrOutput As Variant
    Dim lngPre����ID As Long
    Dim strҽ�� As String, str������ As String, str������ϸ��ˮ�� As String
    Dim bln�ɹ� As Boolean
    Dim cn�ϴ� As New ADODB.Connection
    
    On Error GoTo errHandle
    
    Set cn�ϴ� = GetNewConnection
    cn�ϴ�.Open
    
    '�������ŵ��ݵķ�����ϸ����δ�ϴ��ļ�¼��ȡԭʼ���ݣ�
    gstrSQL = "Select nvl(count(A.ID),0) as ����,nvl(sum(A.�Ƿ��ϴ�),0) �ϴ��� " & _
              "  From סԺ���ü�¼ A,������ҳ B,����֧����Ŀ C" & _
              "  where A.NO=[1] and A.��¼����=[3] and A.��¼״̬<>2 And Nvl(A.��¼״̬,0)<>0 and nvl(A.ʵ�ս��,0)<>0  " & _
              "        and A.����ID=B.����ID and A.��ҳID=B.��ҳID And B.����=[1] and A.�շ�ϸĿID=C.�շ�ϸĿID and B.����=C.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������", TYPE_������, str���ݺ�, int����)
    
    If rsTemp.EOF = True Then
        MsgBox "�õ�����û�п��ϴ������Ϸ�����ϸ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rsTemp("�ϴ���") = 0 Then
        '��ϸ������û���ϴ�������Ҳ�Ͳ���Ҫ��������
        ��������_���� = True
        Exit Function
    End If
    
    If rsTemp("�ϴ���") < rsTemp("����") Then
        MsgBox "�õ����ﻹ��δ�ϴ��ķ�����ϸ���������ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����õ����ڲ����������Ϊ���ʱ��Ƕಡ�˵ģ�
    gstrSQL = " Select A.ID,A.�շ�ϸĿID,A.NO,A.��¼����,A.��¼״̬,A.���,A.����ID,A.��ҳID,A.����ʱ�� as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս��" & _
              "         ,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,C.��Ŀ����,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ,E.���,G.���� ���� " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C,������ҳ D,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
              "  where A.NO=[1] and A.��¼����=[2] and A.��¼״̬=2 and nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.ʵ�ս��,0)<>0" & _
              "        and A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=[3]" & _
              "        and A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����=D.���� " & _
              "        AND B.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) " & _
              "  Order by A.����ʱ��,A.��¼����,Decode(A.��¼״̬,2,2,1)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������", str���ݺ�, int����, TYPE_������)
    
    '�Ƚ�ҽ����Ŀ��Ϣ��������������
    Do While Not rsTemp.EOF
        Call TrackRecordInsure(rsTemp!ID, rsTemp!�շ�ϸĿID)
        rsTemp.MoveNext
    Loop
    
    '���з�����ϸ�Ĵ���
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do Until rsTemp.EOF
        '���ʱ�ͼ��ʵ��������ֳ�������Ҫһ��һ�ʳ���
        str������ϸ��ˮ�� = GetDetailSequence(rsTemp!NO, rsTemp!���, rsTemp!��¼����, rsTemp!��¼״̬)
        str������ = ToVarchar(IIf(IsNull(rsTemp("����Ա����")), UserInfo.����, rsTemp("����Ա����")), 20)
        StrInput = "99|" & str������ϸ��ˮ�� & "|" & str������
        If HandleBusiness(StrInput, arrOutput) = False Then
            '�����ϴ�ʧ��
            If bln�ɹ� = True Then
                MsgBox "�����ϴ���;�������󣬲����Ѿ������Ѿ��ϴ�������Ԥ���㴦���ʣ�����ݵ��ϴ���", vbInformation, gstrSysName
            Else
                MsgBox "�����ϴ���������û�гɹ��ϴ��ļ�¼������Ԥ���㴦���ʣ�����ݵ��ϴ���", vbInformation, gstrSysName
            End If
            ��������_���� = True
            Exit Function
        Else
            '�ڲ��������Ϸ��ü�¼�ϴ��ϱ�ǣ�˵���Ѿ��ϴ�
            gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & rsTemp("ID") & ")"
            '������һ�����Ӵ�ִ�У��ѳɹ��ϴ��Ĵ����ϴ���־
            cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
        End If
        
        rsTemp.MoveNext
        bln�ɹ� = True
    Loop
    
    ��������_���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub AddMessage(strMessage As String, arrOutput As Variant, ByVal str��Ŀ As String, ByVal dbl���� As Currency, Optional ByVal blnסԺ As Boolean = True)
'���ܣ��ڲ��˷�����ϸ��������п��ܲ���һЩ��Ҫ���Ѳ�����Ա����Ϣ
    Dim strTemp As String
    
    If dbl���� > Val(arrOutput(2)) And Val(arrOutput(2)) > 0 Then
        strTemp = "��    " & str��Ŀ & "��ҽԺ�۸� " & Format(dbl����, "0.0000") & " �������ķ��ؼ۸� " & Format(Val(arrOutput(2)), "0.0000") & vbCrLf
    End If
    If arrOutput(3) = 2 And blnסԺ Then
        strTemp = "��    " & str��Ŀ & "��Ҫ��������û��������¼��ֻ����Ϊ�Է���Ŀ" & vbCrLf
    End If
    
    If InStr(strMessage, strTemp) = 0 Then
        strMessage = strMessage & strTemp
    End If
    
End Sub

'ժҪ�ĸ�ʽ����������¼��ԭʼ��ˮ��|�ѳ�����������������¼��ԭʼ��ˮ��|��������ˮ��
Private Function GetDetailSequence(ByVal strNO As String, ByVal int��� As Integer, _
        ByVal int���� As Integer, ByVal int״̬ As Integer) As String
    Dim rsTemp As New ADODB.Recordset
    '����ʱʹ�ã�����NO����¼���ʡ���¼״̬�����ȡԭʼ��¼����ˮ��
    GetDetailSequence = ""
    If int״̬ <> 2 Then Exit Function
    
    gstrSQL = " Select ժҪ From סԺ���ü�¼" & _
              " Where NO=[1] And ���=[2]" & _
              " And ��¼����=[3] And ��¼״̬=3"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡԭʼ������ϸ����ˮ��", strNO, int���, int����)
    If Not rsTemp.EOF Then
        GetDetailSequence = Split(Nvl(rsTemp!ժҪ, "|"), "|")(0)
    Else
        Call DebugTool("δ�ҵ�ԭʼ������ϸ[NO:" & strNO & "|���:" & int��� & "|��¼����:" & int���� & "|��¼״̬:" & int״̬)
    End If
End Function

Private Function UploadNegative(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngID As Long, ByVal lng�շ�ϸĿID As Long) As String
    '��������ʱʹ�ã�����Ϊa)��ȡ�����㹻�ģ�b)��ȡ��ϸ�ۼ��㹻��
    '��ͳ�ƣ��ֳ������ֱ������ݣ��ָ����ѳ�������
    Dim arrOutput
    Dim curͳ���� As Currency, curͳ�����ۼ� As Currency    'һ����������������һ�����ڷֱʳ�����¼�ۼ�ֵ
    Dim StrInput As String
    Dim strҽ�� As String, str������ As String
    
    Dim strժҪ As Double     '����,��ˮ��|����,��ˮ�ţ����ڱ�����ϸ��¼ʱʹ��
    Dim str��������ˮ�� As String, str��ˮ�� As String
    Dim dbl�������� As Double, dbl���������� As Double, dbl�ѳ������� As Double
    Dim rsTemp As New ADODB.Recordset
    Dim rsSource As New ADODB.Recordset
    Dim rsFilter As New ADODB.Recordset
    On Error GoTo errHand
    
    '��ȡ���δ���������
    gstrSQL = "Select A.ID,A.NO,A.��¼����,A.��¼״̬,A.����ID,A.��ҳID,A.����ʱ�� as �Ǽ�ʱ��,Round(A.ʵ�ս��,4) ʵ�ս�� " & _
              "         ,A.�շ�ϸĿID,A.����*nvl(A.����,1) as ����,Decode(A.����*nvl(A.����,1),0,0,Round(A.ʵ�ս��/(A.����*nvl(A.����,1)),4)) as �۸� " & _
              "         ,C.��Ŀ����,B.����,B.����,A.�Ƿ���,nvl(A.������,A.����Ա����) as ҽ��,A.����Ա����,B.���㵥λ,E.���,G.���� ���� " & _
              "  From סԺ���ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C,������ҳ D,ҩƷĿ¼ E ,ҩƷ��Ϣ F,ҩƷ���� G " & _
              "  where A.����ID=D.����ID and A.��ҳID=D.��ҳID And D.����=[1]" & _
              "        And A.�շ�ϸĿID=B.ID and A.�շ�ϸĿID=C.�շ�ϸĿID and C.����=D.���� And Nvl(A.�Ƿ��ϴ�,0)=0" & _
              "        AND B.ID=E.ҩƷID(+) AND E.ҩ��ID=F.ҩ��ID(+) AND F.����=G.����(+) And A.ID=[2]"
    Set rsSource = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���δ���������", TYPE_������, lngID)
    If rsSource.RecordCount <> 0 Then
        dbl���������� = Abs(rsSource!����)
    End If
    Call DebugTool("��ȡ���δ���������")
    
    '��ȡ����¼�ѳ���������
    gstrSQL = " Select SUM(����) AS �ѳ������� From ����������ϸ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����¼�ѳ���������", lngID)
    dbl���������� = dbl���������� - Nvl(rsTemp!�ѳ�������, 0)
    '�����ѳ�����ɵ�δ�����ϴ���־�Ĵ���
    If dbl���������� = 0 Then Exit Function
    
    '��ȡ����ʣ�������ɳ�����ԭʼ��¼��ȡժҪ�м�¼���ѳ���������
    gstrSQL = " Select A.ID,A.����*A.���� AS ����,A.ժҪ," & _
              "     To_Number(Nvl(Substr(A.ժҪ,Decode(Instr(A.ժҪ,'|',1,1),0,Length(A.ժҪ),Instr(A.ժҪ,'|',1,1))+1),0)) as �ѳ�������" & _
              " From סԺ���ü�¼ A" & _
              " Where A.��¼״̬=1 And Nvl(A.ʵ�ս��,0)<>0 And Nvl(A.���ӱ�־,0)<>9 " & _
              " And Nvl(A.�Ƿ��ϴ�,0)=1 And A.�շ�ϸĿID=" & lng�շ�ϸĿID & _
              " And A.����ID=[1] And A.��ҳID=[2]"
    gstrSQL = " Select ID,����-�ѳ������� AS ʣ������,����,�ѳ�������,ժҪ From (" & gstrSQL & ") Where ����-�ѳ�������>0 Order by ʣ������"
    Set rsFilter = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ʣ�������ɳ�����ԭʼ��¼", lng����ID, lng��ҳID)
    Call DebugTool("��ȡ����ʣ�������ɳ�����ԭʼ��¼")
    
    '��һ����ƥ��ԭ��ȡ���ڵ��ڱ��γ��������ļ�¼�����Ͼ������˳�
    With rsFilter
        .Filter = "ʣ������>=" & dbl����������
        If .RecordCount <> 0 Then
            dbl�������� = dbl����������
            str��������ˮ�� = Split(Nvl(!ժҪ, "|"), "|")(0)
            
            '�ϴ�������¼
            strҽ�� = ToVarchar(IIf(IsNull(rsSource("ҽ��")), UserInfo.����, rsSource("ҽ��")), 20)
            str������ = ToVarchar(IIf(IsNull(rsSource("����Ա����")), UserInfo.����, rsSource("����Ա����")), 20)
            
            StrInput = "04|" & GetIdentify(rsSource("����ID"), rsSource("��ҳID"))
            StrInput = StrInput & "|" & rsSource("NO") & "_" & rsSource!��¼����
            StrInput = StrInput & "|" & Format(rsSource("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
            StrInput = StrInput & "|" & ToVarchar(rsSource("��Ŀ����"), 10)     '���ı���
            StrInput = StrInput & "|" & ToVarchar(rsSource("����"), 20)         'ҽԺ����
            StrInput = StrInput & "|" & ToVarchar(rsSource("����"), 50)         '��Ŀ����
            StrInput = StrInput & "|" & Format(rsSource("�۸�"), "0.0000")      '����
            StrInput = StrInput & "|" & Format(-1 * dbl��������, "0.00")       '����
            StrInput = StrInput & "|" & IIf(rsSource("�Ƿ���") = 1, 1, 0)     '�����־
            StrInput = StrInput & "|" & strҽ��                               'ҽ��
            StrInput = StrInput & "|" & str������                             '������
            StrInput = StrInput & "|" & ToVarchar(rsSource("���㵥λ"), 20)     '��λ
            StrInput = StrInput & "|" & ToVarchar(rsSource("���"), 14)         '���
            StrInput = StrInput & "|" & ToVarchar(rsSource("����"), 20)         '����
            StrInput = StrInput & "|" & str��������ˮ��                     '������ϸ��ˮ��
            StrInput = StrInput & "|" & Format(-1 * rsSource("�۸�") * dbl��������, "#0.0000")     '���
            
            Call DebugTool("׼�������ϴ�����������ˮ��:" & str��������ˮ�� & ";��������:" & dbl�������� & ";����������:" & dbl����������)
            If HandleBusiness(StrInput, arrOutput) = False Then
                UploadNegative = "����ID=" & lngID & "�ĸ���������¼�ϴ�ʧ�ܣ�����������¼��ˮ��=" & str��������ˮ�� & "����������=" & dbl�������� & "��"
                Exit Function
            Else
                '�ɹ��ͱ������ϸ���������ѳ�������
                '�ڷ��ü�¼�ϴ��ϱ�ǣ�˵���Ѿ��ϴ��������淵�صĽ��
                If arrOutput(3) = 2 Then
                    'δͨ������
                    curͳ���� = 0
                Else
                    '��׼���� * ����
                    curͳ���� = Val(arrOutput(2)) * rsSource("����")
                End If
                '����ԭʼ�������ѳ�������
                dbl�ѳ������� = !�ѳ������� + dbl��������
                gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & !ID & ",NULL,'" & str��������ˮ�� & "|" & dbl�ѳ������� & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                '���ϴ���־
                gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & lngID & "," & curͳ���� & ",'" & arrOutput(1) & "|" & str��������ˮ�� & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                '����������ϸ
                gstrSQL = "zl_����������ϸ_Insert(" & lng����ID & "," & lng��ҳID & "," & lngID & "," & dbl�������� & ",'" & arrOutput(1) & "','" & str��������ˮ�� & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                Call DebugTool("�����ϴ���־����������ϸʣ��������������������ϸ��¼")
            End If
            dbl���������� = 0
        End If
        .Filter = 0
        
        If dbl���������� <> 0 Then
            '˵�����㹻�������ɳ�������ʣ���ۼ��Ƿ�����㹻������
            Do While Not .EOF
                If dbl���������� > !ʣ������ Then
                    dbl�������� = !ʣ������
                Else
                    dbl�������� = dbl����������
                End If
                str��������ˮ�� = Split(Nvl(!ժҪ, "|"), "|")(0)
                
                '�ϴ�������¼
                strҽ�� = ToVarchar(IIf(IsNull(rsSource("ҽ��")), UserInfo.����, rsSource("ҽ��")), 20)
                str������ = ToVarchar(IIf(IsNull(rsSource("����Ա����")), UserInfo.����, rsSource("����Ա����")), 20)
                
                StrInput = "04|" & GetIdentify(rsSource("����ID"), rsSource("��ҳID"))
                StrInput = StrInput & "|" & rsSource("NO") & "_" & rsSource!��¼����
                StrInput = StrInput & "|" & Format(rsSource("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
                StrInput = StrInput & "|" & ToVarchar(rsSource("��Ŀ����"), 10)     '���ı���
                StrInput = StrInput & "|" & ToVarchar(rsSource("����"), 20)         'ҽԺ����
                StrInput = StrInput & "|" & ToVarchar(rsSource("����"), 50)         '��Ŀ����
                StrInput = StrInput & "|" & Format(rsSource("�۸�"), "0.0000")      '����
                StrInput = StrInput & "|" & Format(-1 * dbl��������, "0.00")       '����
                StrInput = StrInput & "|" & IIf(rsSource("�Ƿ���") = 1, 1, 0)     '�����־
                StrInput = StrInput & "|" & strҽ��                               'ҽ��
                StrInput = StrInput & "|" & str������                             '������
                StrInput = StrInput & "|" & ToVarchar(rsSource("���㵥λ"), 20)     '��λ
                StrInput = StrInput & "|" & ToVarchar(rsSource("���"), 14)         '���
                StrInput = StrInput & "|" & ToVarchar(rsSource("����"), 20)         '����
                StrInput = StrInput & "|" & str��������ˮ��                     '������ϸ��ˮ��
                StrInput = StrInput & "|" & Format(-1 * rsSource("�۸�") * dbl��������, "#0.0000")     '���
                
                If HandleBusiness(StrInput, arrOutput) = False Then
                    UploadNegative = "����ID=" & lngID & "�ĸ���������¼�ϴ�ʧ�ܣ�����������¼��ˮ��=" & str��������ˮ�� & "����������=" & dbl�������� & "��"
                    Exit Function
                Else
                    '�ɹ��ͱ������ϸ���������ѳ�������
                    '�ڷ��ü�¼�ϴ��ϱ�ǣ�˵���Ѿ��ϴ��������淵�صĽ��
                    If arrOutput(3) = 2 Then
                        'δͨ������
                        curͳ���� = 0
                    Else
                        '��׼���� * ����
                        curͳ���� = -1 * Val(arrOutput(2)) * dbl��������
                    End If
                    curͳ�����ۼ� = curͳ�����ۼ� + curͳ����
                    dbl�ѳ������� = !�ѳ������� + dbl��������
                    dbl���������� = dbl���������� - dbl��������
                    
                    '����ԭʼ�������ѳ�������
                    gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & !ID & ",NULL,'" & str��������ˮ�� & "|" & dbl�ѳ������� & "')"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    '����������ϸ
                    gstrSQL = "zl_����������ϸ_Insert(" & lng����ID & "," & lng��ҳID & "," & lngID & "," & dbl�������� & ",'" & arrOutput(1) & "','" & str��������ˮ�� & "')"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    Call DebugTool("���±�������ϸʣ��������������������ϸ��¼")
                End If
                
                If dbl���������� = 0 Then
                    '���ϴ���־
                    gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & lngID & "," & curͳ�����ۼ� & ",'" & arrOutput(1) & "|" & str��������ˮ�� & "')"
                    gcnOracle.Execute gstrSQL, , adCmdStoredProc
                    '�ѳ�����ϣ���������
                    Call DebugTool("�����ϴ���־")
                    Exit Function
                End If
                
                .MoveNext
            Loop
        End If
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function TrackRecordInsure(ByVal lng����ID As Long, ByVal lng�շ�ϸĿID As Long) As Boolean
    Dim str��ˮ�� As String, str�������� As String, str�շ���� As String, str��ע As String
    Dim dbl��׼���� As Double, dbl�Ը����� As Double
    Dim rsTemp As New ADODB.Recordset

    '20061113,zyb:ȡ��TrackRecordInsure()�ĵ��ã������ҽ��ǰ�û����������ࣨ��������������������dy_initҲ���һ�����ӣ�
    Exit Function
    
    '��¼ҽ����Ŀ��ʱ�Ļ�����Ϣ��ҽ����Ŀ���룬�������ͣ���׼����,�Ը�������
    Call DebugTool("����TrackRecordInsure")
    gstrSQL = "Select A.���,B.��Ŀ���� " & _
        " From �շ�ϸĿ A,����֧����Ŀ B" & _
        " Where A.ID=B.�շ�ϸĿID And B.����=[1] And A.ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����Ŀ���������", TYPE_������, lng�շ�ϸĿID)
    If rsTemp.RecordCount = 0 Then Exit Function
    str��ˮ�� = Nvl(rsTemp!��Ŀ����)
    str�շ���� = rsTemp!���
    Call DebugTool("��ǰҽ����Ŀ���룺" & str��ˮ��)
    If str��ˮ�� = "" Then Exit Function
    
    '������
    If mcnYB.State = 0 Then
        If Not OpenDatabase Then Exit Function
    End If
    
    If InStr(1, "5,6,7", str�շ����) <> 0 Then
        'ҩƷ
        gstrSQL = "select YPLSH  ҽ������,YPBM ҩƷ����,TYM ͨ������,SPM ��Ʒ��,SPMZJM ��Ʒ������,YCMC ҩ������,decode(FYDJ,1,'����',2,'����','�Է�') ���õȼ� " & _
                  "      ,PFJ ������,BZDJ ��׼����,ZFBL �Ը�����,JX ����,BZSL ��װ����,BZDW ��װ��λ,HL ����,HLDW ������λ,RL ����,RLDW ������λ " & _
                  "      ,DECODE(CFYBZ,1,'��') ����ҩ��־,decode(GMP,1,'��') GMP��־,decode(YPXJFS,1,'�޼�') �޼�,TQFYDJ ��Ⱥ��Ŀ�ȼ�,TQZFBL ��Ⱥ�Ը�����,TQBZDJ ��Ⱥ��׼���� " & _
                  "  FROM YPML WHERE YPLSH='" & str��ˮ�� & "'"
    Else
        '����
        gstrSQL = "Select XMLSH ҽ������,XMBM ���Ʊ���,XMMC ��Ŀ����,ZJM ����,decode(FYDJ,1,'����',2,'����','�Է�') ���õȼ�,DW ��λ " & _
                 "       ,TPJ ������,BZJ ��׼����,ZZBL ��ְ�Ը�����,TXBL �����Ը�����,decode(XJFS,1,'ͳһ�޼�',2,'��ҽԺ�ȼ�����',3,'������ҽԺ��׼��������') �޼� " & _
                 "       ,TQFYDJ ��Ⱥ��Ŀ�ȼ�,TQZFBL ��Ⱥ�Ը�����,TQBZDJ ��Ⱥ��׼����,decode(TPXMBZ,1,'��') ������Ŀ��־,BZ ��ע " & _
                 "   FROM ZLXM WHERE XMLSH='" & str��ˮ�� & "'"
    End If
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, mcnYB
        If .RecordCount = 0 Then
            Call DebugTool("δ�ҵ���ҽ����Ŀ")
            Exit Function
        End If
    End With
    
    str�������� = Nvl(rsTemp!���õȼ�)
    dbl��׼���� = Nvl(rsTemp!��׼����, 0)
    If InStr(1, "5,6,7", str�շ����) <> 0 Then
        dbl�Ը����� = Nvl(rsTemp!�Ը�����, 0)
        str��ע = "||||" & Nvl(rsTemp!��Ⱥ�Ը�����, 0)
    Else
        dbl�Ը����� = Nvl(rsTemp!��ְ�Ը�����, 0)
        str��ע = Nvl(rsTemp!������, 0) & "||" & Nvl(rsTemp!�����Ը�����, 0) & "||" & Nvl(rsTemp!��Ⱥ�Ը�����, 0)
    End If
    
    '�����¼���������жϣ�������ڼ�¼��������£�������룩
    '����ID,ҽ����Ŀ����,��������,��׼����,�Ը�����,��ע
    gstrSQL = "zl_ҽ����Ŀ��Ϣ_INSERT(" & lng����ID & ",'" & str��ˮ�� & "','" & str�������� & "'," & _
        dbl��׼���� & "," & dbl�Ը����� & ",'" & str��ע & "')"
    Call DebugTool("����ҽ����Ŀ��Ϣ��" & gstrSQL)
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ŀ��Ϣ��¼")
    TrackRecordInsure = True
End Function

Private Function OpenDatabase() As Boolean
    Dim strServer As String, strUser As String, strPass As String, strTemp As String
    Dim rsTemp As New ADODB.Recordset
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ղ���", TYPE_������)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽ��������"
                strServer = strTemp
            Case "ҽ���û���"
                strUser = strTemp
            Case "ҽ���û�����"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
        Exit Function
    End If
    OpenDatabase = True
End Function

Public Function �ҺŽ���_����(lng����ID As Long) As Boolean
    Dim intTotal As Integer, intStart As Integer
    Dim str���㷽ʽ As String, arr���㷽ʽ
    Dim cur�����ʻ� As Currency, curҽ������ As Currency, cur����Ա���� As Currency, cur���ͳ�� As Currency, cur����Ա���� As Currency, cur�������� As Currency
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    gstrSQL = "Select ����ID,�շ�ϸĿID,����*NVL(����,1) AS ����,��׼���� As ����,ʵ�ս��,������," & IIf(g��������.�����Ը���� = 14, "1", "0") & " As �Ƿ���" & _
        " From ������ü�¼ " & _
        " Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    If Not �����������_����(rsTemp, str���㷽ʽ) Then Exit Function
    If Not �������_����(lng����ID, 0, "", "", False) Then Exit Function
    
    '�ֽ���ֽ��㷽ʽ
    arr���㷽ʽ = Split(str���㷽ʽ, "|")
    intTotal = UBound(arr���㷽ʽ)
    For intStart = 0 To intTotal
        Select Case Split(arr���㷽ʽ(intStart), ";")(0)
        Case "�����ʻ�"
            cur�����ʻ� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "ҽ������"
            curҽ������ = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "����Ա����"
            cur����Ա���� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "���ͳ��"
            cur���ͳ�� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        Case "����Ա����"
            cur����Ա���� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
                Case "��������"
            cur�������� = Val(Split(arr���㷽ʽ(intStart), ";")(1))
        End Select
    Next
    
   '��Ҫ����������
    str���㷽ʽ = ""
    If cur�����ʻ� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & cur�����ʻ�
    If curҽ������ <> 0 Then str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curҽ������
    If cur���ͳ�� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||���ͳ��|" & cur���ͳ��
    If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
    If cur����Ա���� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||����Ա����|" & cur����Ա����
        If cur�������� <> 0 Then str���㷽ʽ = str���㷽ʽ & "||��������|" & cur��������
    If str���㷽ʽ <> "" Then
        str���㷽ʽ = Mid(str���㷽ʽ, 3)
    Else
        str���㷽ʽ = "�����ʻ�|0"
    End If
    gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
    
    �ҺŽ���_���� = True
    Call frm������Ϣ.ShowME(lng����ID)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetIdentify(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal bln���� As Boolean = False) As String
    Dim strסԺ����� As String
    'MODIFIED BY ZYB 20040626 '��Ժ�޸ģ�����������ϵͳ�ǰ�����ID�ϴ��ģ���ˣ���ˮ�Ÿ�Ϊ��
    '��������ʻ��д���סԺ����ţ���ֵ��Ϊ�գ���������Ϊ���˱�ʶ�ϴ�����������ģʽ�ϴ�
    If GetMode(lng����ID, lng��ҳID, strסԺ�����) Then
        GetIdentify = lng����ID & "_" & lng��ҳID & "_" & Get�������(lng����ID, bln����)
    Else
        GetIdentify = strסԺ�����
    End If
End Function

Private Function Get�������(ByVal lng����ID As Long, ByVal bln���� As Boolean) As Integer
    Dim int��� As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    gstrSQL = " Select ������� From �����ʻ� " & _
              " Where ����=[1] ANd ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡסԺ�����", TYPE_������, lng����ID)
    If Err <> 0 Then
        MsgBox "�����ʻ���Ľṹ����ȷ����Ҫ�����ֶΡ�������š���", vbInformation, gstrSysName
        Exit Function
    End If
    
    int��� = Val(Nvl(rsTemp!�������, 0))
    If bln���� Then
        int��� = int��� + 1
        gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_������ & ",'�������','''" & int��� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
    End If
    Get������� = int���
End Function

Private Function GetMode(ByVal lng����ID As Long, ByVal lng��ҳID As Long, strסԺ����� As String) As Boolean
    'MODIFIED BY ZYB 20040626 '��Ժ�޸ģ�����������ϵͳ�ǰ�����ID�ϴ��ģ���ˣ���ˮ�Ÿ�Ϊ��
    '��������ʻ��д���סԺ����ţ���ֵ��Ϊ�գ���������Ϊ���˱�ʶ�ϴ�����������ģʽ�ϴ�
    Dim blnģʽ As Boolean              'Ϊ�棬������ID_��ҳID_������ŷ��أ���������ز���ID
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    gstrSQL = " Select סԺ����� From �����ʻ� " & _
              " Where ����=[1] ANd ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡסԺ�����", TYPE_������, lng����ID)
    If Err <> 0 Then
        '˵�������ڸ��ֶ�
        blnģʽ = True
    Else
        blnģʽ = (Nvl(rsTemp!סԺ�����) = "")
        If Not blnģʽ Then strסԺ����� = Nvl(rsTemp!סԺ�����)
    End If
    GetMode = blnģʽ
End Function
'
'Private Function GetRegisted(ByVal strҽ���� As String) As Long
'    Dim strDate As String, strStart As String, strEnd As String
'    Dim rsTemp As New ADODB.Recordset
'
'    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
'    strStart = strDate & " 00:00:00"
'    strEnd = strDate & " 23:59:59"
'    '��������ڴ��ھ����¼(�ҺŻ��շ�)���򷵻ز���ID�����򷵻���
'    gstrSQL = " Select A.����ID From ���˷��ü�¼ A,���ս����¼ B" & _
'              " Where A.��¼���� In (1,4) And A.����ID Is Not NULL" & _
'              " And A.�Ǽ�ʱ�� Between to_date('" & strStart & "','yyyy-MM-dd hh24:mi:ss')" & _
'              " And to_date('" & strEnd & "','yyyy-MM-dd hh24:mi:ss')" & _
'              " And A.����ID+0 =(Select ����ID From �����ʻ� Where ����=" & TYPE_������ & " ANd ҽ����='" & strҽ���� & "')" & _
'              " And A.����ID=B.��¼ID And B.����=1"
'    Call OpenRecordset(rsTemp, "ȡ����ID")
'    If rsTemp.RecordCount = 0 Then Exit Function
'    GetRegisted = rsTemp!����ID
'End Function

Private Function GetRegisted(ByVal strҽ���� As String, ByRef str����ʱ�� As String) As Long
    '���������ھ���ǼǼ�¼��˵�������
    Dim strDate As String, strStart As String, strEnd As String
    Dim rsTemp As New ADODB.Recordset

    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    strStart = strDate & " 00:00:00"
    strEnd = strDate & " 23:59:59"
    
    'һ���������������ǼǼ�¼������ҽ����ˮ���ǲ���ģ�������ȡһ��
    gstrSQL = " Select A.����ID,A.����ʱ��" & _
              " From ����ǼǼ�¼ A,�����ʻ� B" & _
              " Where A.��ҳID=0 And A.��¼ID Is Not NULL And A.����ID=B.����ID And A.����=B.���� " & _
              " And B.����=[1] And B.ҽ����=[2]" & _
              " And A.����ʱ�� Between [3] And [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�жϸò��˵����Ƿ�����", TYPE_������, strҽ����, CDate(strStart), CDate(strEnd))
    If rsTemp.RecordCount = 0 Then Exit Function
    
    str����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    GetRegisted = rsTemp!����ID
End Function

Public Function ȡ������_������(ByVal bytType As Byte, ByVal lng����ID As Long)
    gstrSQL = "zl_����ǼǼ�¼_DELETE(" & TYPE_������ & "," & lng����ID & ",0,to_date('" & gstr����ʱ�� & "','yyyy-MM-dd hh24:mi:ss'))"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
End Function

Private Function MessageInfo(ByVal strInfo As String, ByVal bln����Ԥ���� As Boolean) As String
    '�����ʾ����MSGBOX������ֵΪ�գ�������ʾ������ֵ���ڽ�Ҫ��ʾ����Ϣ
    If Not bln����Ԥ���� Then
        MsgBox strInfo, vbInformation, gstrSysName
    Else
        MessageInfo = strInfo
    End If
End Function

Private Function Is��ͥ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select ��Ժ���� From ������ҳ Where ��ǰ����ID is not null And ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��ͥ����", lng����ID, lng��ҳID)
    Is��ͥ���� = IsNull(rsTemp!��Ժ����)
End Function
