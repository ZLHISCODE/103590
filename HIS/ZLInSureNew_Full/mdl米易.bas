Attribute VB_Name = "mdl����"
Option Explicit
Private objCom As Object
Public gintComPort As Integer
Private mblnCreated As Boolean
Public Enum �ֶ�
    ���˱�� = 1
    ����
    ����������
    ����
    ������
    ֧�����
    ������
    ����ʱ��
    ������
    ������
    ֧�����
    ��Ա���
    ������ˮ��
    ҽ������
    ��Ŀ����
    ��Ŀ����
    ����
    �����ܶ�
    ��������
    ����ҽ��
    �ܵ�����
    �ܵ�ҽ��
    Ȧ������
    Ȧ����
    Ȧ��ʱ��
    Ȧ����ˮ��
    ���ֱ���
    ������ȡʱ��
    ��Ժ����
    ��Ժ���
    ��Ժ����
    ��Ժ������
    ��Ժ����ʱ��
    ����ʱ��
    ����Ƚ����־
    ��ʵ�����־
    ��Ժԭ��
    ��Ժ����
    ��Ժ���
    ��Ժ����
    ��Ժ������
    ��Ժ����ʱ��
    �˵����
    ������_��ʼ
    ������_��ֹ
    ������ˮ��_��ʼ
    ������ˮ��_��ֹ
    ������_��ʼ
    ������_��ֹ
End Enum

Public gstrPara_���� As String           '���ýӿڵĲ�����
Private Type ComInfo_����
    ϵͳʱ�� As String
    ���˱�� As String
    ���� As String
    ���������� As String
    ����  As String
    ������ As String
    ֧����� As Double
    ������ As String
    ������ As String
    ֧����� As String
    Ȧ������ As String
    ��Ȧ���� As Double
    Ȧ���� As Double
    Ȧ����ˮ�� As String
    ���ֱ��� As String
    ������ȡʱ�� As String
    ��Ժ���� As String
    ��Ժ��� As String
    ��Ժ���� As String
    ������ˮ�� As String
    ��Ժԭ�� As String
    ��Ժ���� As String
    ��Ժ��� As String
    ��Ժ���� As String
    ������ As String
    �������� As String
    ������� As Long
    ����������Ϣ As String
    ִ�н�� As Long                '0��ʾ��������1��ʾ�������
    ���� As String
    �Ա� As String
    �������� As String
    ���֤�� As String
    ��Ա��� As String
    ��λ���� As String
    ���� As Long
    �ʻ���� As Double
    �������� As Double
    ����ͳ���޶� As Double
    ����ͳ�ﱨ������ As Double
    �ҹ��Ը� As Double
    ����ͳ�� As Double
    �����Ը� As Double
    ͳ���Ը� As Double
    ͳ��֧�� As Double
    �����Ը� As Double
    �������� As Double
'���±������ں˶�
    �˶�_��¼�� As Double
    �˶�_�ʻ�֧���ܶ� As Double
    �˶�_ҽ�Ʒ��ܶ� As Double
    �˶�_ȫ�Է��ܶ� As Double
    �˶�_�ҹ��Է��ܶ� As Double
    �˶�_����ͳ���ܶ� As Double
    �˶�_�ֽ�֧���ܶ� As Double
    �˶�_�����Ը��ܶ�  As Double
    �˶�_��Ժ���� As Long
    �˶�_��Ժ���� As Long
    �˶�_���� As Long
    �˶�_���� As Double
    �˶�_�Ը����� As Double
    �˶�_����֧����� As Double
    �˶�_ͳ���Ը���� As Double
    �˶�_ͳ��֧����� As Double
End Type
Public gComInfo_���� As ComInfo_����      '���浱ǰ����������
Private Const gintȦ����ˮ�� As Integer = 1
Private Const gint������ˮ�� As Integer = 2
Private Const gint������ As Integer = 3
'���붨��ΪLONG
Private Const glngȦ�� As Long = 1
Private Const glng���� As Long = 0

Public Declare Function Card_Sale Lib "jpmyyy.dll" Alias "card_sale" _
(ByVal comport As Integer, ByVal userpsd As String, ByVal jystr As String, ByVal jymode As Integer) As Integer
'comport as ���ں�(1-com1,2-com2,3-com3...)
'userpsd as �α���ʹ�����룬6λ�����ַ���,
'jystr as  41���ֽڳ��� as jyje as 15λǰ̨�����¼��,2λ����Ա����, 8λҩ��/ҽԺ����,8λ���׽��,8λˢ�����,
'����0 as ��ȷ,
'���� as Ȧ�桢���ʽӿڳ���,
Public Declare Function Card_ChangePsd Lib "jpmyyy.dll" Alias "change_psd" _
(ByVal comport As Integer, ByVal oldpsd As String, ByVal newpsd As String) As Integer
'comport as ���ں�,oldpsd as ���û�����,newpsd as �޸ĺ����,
'����ֵ as 0 as �޸���ȷ��4 as  д������, 3 as ת�����ݴ���,
' 2 as  ��������,1 as ����ԭ�������ȷ,
Public Declare Function Card_userinfo Lib "jpmyyy.dll" Alias "re_userinfo" _
(ByVal comport As Integer, ByVal userpsd As String, recode As Integer) As String
'��Ƭ��ȫ��֤������10���ţ�15����˳��ţ�8�����ֽ��(�Է�Ϊ��λ)
'����ֵ��comport as ��д�����ںţ�userpsd as ʹ�����룬
'recode as ����ֵ as 0�ɹ�������ʧ��
'���ش� as recodeΪ0���ص�ֵ��ȷ�������Ǻ����ַ���������Ϣ,

'__________________________________________________________________________________
'��Ҫ���棺�ӿ��ڶ������籣�ţ�����˱�š����Ų�ͬ������������ʱ���ߣ�����������
'��Ҫ�������֤ʱ: ��ʾ����������Ϣ����
'����ʱ�������ˣ�aae011���̶�Ϊ"900090009000999"
'סԺ��֧����;���� , ͨ��ģ�����ʵ��
'����������ķ�ʽ�¸����ʻ�

Public Function ҽ����ʼ��_����() As Boolean
    Dim rs����������  As New ADODB.Recordset
    Dim rs���ղ���  As New ADODB.Recordset
    '���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
    '���أ���ʼ���ɹ�������true�����򣬷���false
    
    On Error Resume Next
    gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
    Set rs���������� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽԺ����", type_����)
    
    '����Ȧ�����������������һ�£�����ͬʱ������ֵ
    gComInfo_����.���������� = Nvl(rs����������!ҽԺ����, "")
    gComInfo_����.Ȧ������ = gComInfo_����.����������
    
    gstrSQL = "Select ����ֵ From ���ղ��� Where ����=[1]"
    Set rs���ղ��� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", type_����)
    gintComPort = 1
    If Not rs���ղ���.EOF Then gintComPort = Nvl(rs���ղ���!����ֵ, 1)
    
    ҽ����ʼ��_���� = True
End Function

Public Function ҽ����ֹ_����() As Boolean
    Debug.Print -1
    Set objCom = Nothing
    mblnCreated = False
    ҽ����ֹ_���� = True
End Function

Public Function ����() As Boolean
    If mblnCreated Then Call ҽ����ֹ_����
    
    mblnCreated = ��������
    If Not mblnCreated Then
        MsgBox "�޷�����COM+����ҽ����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    ���� = True
End Function

Public Function ҽ������_����() As Boolean
    ҽ������_���� = frmSet����.ShowME(type_����)
End Function

Public Function ��ݱ�ʶ_����(Optional bytType As Byte, Optional lng����ID As Long) As String
    Dim str��ע As String, RSPATIENT As New ADODB.Recordset
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    If Not ���� Then Exit Function
    ��ݱ�ʶ_���� = frmIdentify����.GetPatient(bytType, lng����ID)
    Call ҽ����ֹ_����
    If ��ݱ�ʶ_���� = "" Then Exit Function
    
    'ǿ�ưѱ�����������������
    gstrSQL = "Select ����ID From �����ʻ� Where ����=[1] And ҽ����=[2]"
    Set RSPATIENT = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����˵Ļ�����Ϣ", type_����, gComInfo_����.���˱��)
    lng����ID = RSPATIENT!����ID
    
    str��ע = Val(gComInfo_����.��������) & ";" & Val(gComInfo_����.����ͳ�ﱨ������) & ";" & Val(gComInfo_����.����ͳ���޶�)
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & lng����ID & "," & type_���� & ",'��ע','''" & str��ע & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ݱ�ʶ_����")
End Function

Public Function �������_����(Optional ByVal bln���� As Boolean = False, Optional ByVal strSelfNo As String = "", Optional ByVal blnסԺ As Boolean = False) As Currency
    '����: ֱ�Ӷ������ڽ��
    '����: �Ƿ����
    '����: ���ظ����ʻ����
    Dim lng����ID As Long
    Dim rsAcc As New ADODB.Recordset
    '����ʧ�����˳�
    gstrSQL = "Select Nvl(�ʻ����,0) �ʻ���� From �����ʻ� Where ����=[1]"
    If bln���� Then
        lng����ID = ReadICCard(blnסԺ)
        If lng����ID = 0 Then Exit Function
        'ֱ�ӷ���
        �������_���� = gComInfo_����.�ʻ���� + gComInfo_����.��Ȧ����
        Exit Function
        
        gstrSQL = gstrSQL & " And ����ID=[2]"
    Else
        gstrSQL = gstrSQL & " And ҽ����=[3]"
    End If
    Set rsAcc = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ����", type_����, lng����ID, strSelfNo)
    
    gComInfo_����.�ʻ���� = rsAcc!�ʻ����
    �������_���� = gComInfo_����.�ʻ���� + gComInfo_����.��Ȧ����
End Function

Private Function ReadICCard(Optional ByVal blnסԺ As Boolean = False) As Long
    '��ȡ������Ϣ��ͬʱ���½ṹ���в��������Ϣ���ʻ������ز���ID
    Dim recode As Integer, strResult As String, str������ As String
    Dim lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    strResult = Card_userinfo(gintComPort, gComInfo_����.����, recode)
    If recode <> 0 Then
        MsgBox strResult, vbInformation, gstrSysName
        Exit Function
    End If
    
    gComInfo_����.���˱�� = Mid(strResult, 11, 15)
    gComInfo_����.���� = Mid(strResult, 1, 10)
    gComInfo_����.�ʻ���� = Val(Mid(strResult, 26, 8)) / 100                   '����Ϊ��λ��¼�Ľ��ת��Ϊ��ԪΪ��λ�Ľ��
    
    str������ = gComInfo_����.������
    '�����֤
    gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����, gComInfo_����.����) & _
        GetParaCode(����������, gComInfo_����.����������)
    If Not ���ýӿ�_����("identifyinfogetting") Then
        Exit Function
    End If
    gComInfo_����.������ = str������
    
    gstrSQL = "Select ����ID From �����ʻ� Where ����=[1] And ҽ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҽ�����˵Ĳ���ID", type_����, gComInfo_����.���˱��)
    If rsTemp.EOF Then Exit Function
    ReadICCard = rsTemp!����ID
End Function

Public Function WriteICCard(ByVal lng����ID As Long, ByVal curMoney As Currency, Optional ByVal blnסԺ As Boolean = False) As Boolean
    Dim blnRead As Boolean, blnErr As Boolean
    Dim lngReturn As Long
    Dim StrInput As String
    
    On Error GoTo errHand
    'д�������ʻ����ʣ�
    If curMoney = 0 Then Exit Function
    
    '���ö���
    blnRead = True
    Do While blnRead
        'ReadICCard:���¶�ȡ��������Ȧ����
        If lng����ID <> ReadICCard(blnסԺ) Then
            MsgBox "�������ڵĿ����ǵ�ǰ���˵ģ��������ȷ�Ŀ��󣬰��س�����", vbInformation, gstrSysName
        Else
            blnRead = False
        End If
    Loop
    
    'Ҫ���±����ʻ�����ͬʱ���½ṹ���ڵ�ֵ
    '15λǰ̨�����¼��,2λ����Ա����, 8λҩ��/ҽԺ����,8λ���׽��,8λˢ�����
    StrInput = Abs(curMoney) * 100 'ת��Ϊ�ֵĸ�ʽ
    If Len(StrInput) < 8 Then StrInput = String(8 - Len(StrInput), "0") & StrInput
    StrInput = Right(gComInfo_����.������, 15) & "01" & gComInfo_����.���������� & StrInput & StrInput
    If curMoney > 0 Then
        '����
        lngReturn = Card_Sale(gintComPort, gComInfo_����.����, StrInput, glngȦ��)
    Else
        '����
        lngReturn = Card_Sale(gintComPort, gComInfo_����.����, StrInput, glng����)
    End If
    blnErr = (lngReturn = 0)
    gComInfo_����.�ʻ���� = gComInfo_����.�ʻ���� + curMoney
'    gComInfo_����.Ȧ���� = 0         '������Ϊ�㣬�����¼���ܷ�Ӧ�����ڽ��
    
    WriteICCard = blnErr
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function �����������_����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    Dim curTotal As Currency, cur�����ʻ� As Currency
    Dim rsTemp As New ADODB.Recordset
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ�
    
    '�����֤�󣬷��ص��ǿ����Ȧ����
    With rs��ϸ
        'ȡ�����η������õĽ��ϼ�
        Do While Not .EOF
            '���ж��Ƿ�������ҽ����Ӧ��Ŀ����
            gstrSQL = " Select ��Ŀ���� From ����֧����Ŀ" & _
                      " Where ����=[1] And �շ�ϸĿID=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ������˶�Ӧ��ҽ������", type_����, CLng(!�շ�ϸĿID))
            If rsTemp.EOF = True Then
                MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            
            curTotal = curTotal + CCur(Format(!ʵ�ս��, "#####0.00;-#####0.00;0;"))
            .MoveNext
        Loop
        
        gComInfo_����.֧����� = curTotal            '�ݴ�����ܶ�
        If curTotal > gComInfo_����.�ʻ���� + gComInfo_����.��Ȧ���� Then
            cur�����ʻ� = gComInfo_����.�ʻ���� + gComInfo_����.��Ȧ����
        Else
            cur�����ʻ� = curTotal
        End If
        str���㷽ʽ = "�����ʻ�;" & cur�����ʻ� & ";1"   '�����޸�
    End With
    
    '����Ƿ���ڱ��ξ����ţ���������ʾ���ܽ��㣬�����˳�����ˢ��
    Dim strDate As String
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    gstrSQL = "Select count(*) Records From ���ս����¼ Where ����=1 And ����=[1] And ����ʱ�� Between [2] And [3] And ֧��˳���=[4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���ڱ��ξ�����", type_����, CDate(strDate), CDate(strDate & " 23:59:59"), gComInfo_����.������)
    If rsTemp!Records <> 0 Then Exit Function
    
    �����������_���� = True
End Function

Public Function �������_����(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    Dim curTotal As Currency
    Dim int�ϴ� As Integer
    Dim lng����ID As Long
    Dim bln���ֲ� As Boolean, blnError As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    '�����ʻ�����֧��ȫ�Էѡ������Ը����֣���ˣ�ֻҪ�������㹻�Ľ�����ȫ��ʹ�ø����ʻ�֧��
    'ע�⣺�ӿڹ涨��������ϸ�������ϴ���סԺ��ϸ��Ԥ����ʱ�ϴ���������ڽ��㣬����ʹ��Ȧ��ӿڣ����������Ǯ���������ڣ������ӿ��ڽ��
    '���������Ҫͨ��������������ȡ����Ȧ�����ǽӿڷ��أ���Ҫ�޸�
    On Error GoTo errHand
    If Not ���� Then Exit Function
    If Not ���ýӿ�_����("getsysdate") Then Exit Function
    
    Call �������_����(True, strSelfNo)
    
    '֧�������ҽ�������Ƿ������ֲ��й�
    bln���ֲ� = False
    gstrSQL = " Select A.����ID,Nvl(B.���,0) �ز��� " & _
              " From �����ʻ� A,(Select * From ���ղ��� Where ����=" & type_���� & ") B " & _
              " Where A.����=[1] And A.ҽ����=[2] And A.����ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ������ֲ�", type_����, strSelfNo)
    lng����ID = rsTemp!����ID
    bln���ֲ� = (rsTemp!�ز��� <> 0)
    '�������ֲ�����ҽԺ��֧��
    gComInfo_����.֧����� = "0201" 'IIf(bln���ֲ�, "0205", "0201")
    gComInfo_����.������ = GetSequence(gint������)
    gComInfo_����.������ˮ�� = GetSequence(gint������ˮ��)
    
    '����Ȧ��ӿ�
    If cur�����ʻ� > gComInfo_����.�ʻ���� Then
        gComInfo_����.Ȧ���� = cur�����ʻ� - gComInfo_����.�ʻ����
        gComInfo_����.Ȧ����ˮ�� = Right(gComInfo_����.����������, 3) & GetSequence(gintȦ����ˮ��)
        
        gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����, gComInfo_����.����) & _
            GetParaCode(Ȧ������, gComInfo_����.Ȧ������) & GetParaCode(Ȧ����, gComInfo_����.Ȧ����) & _
            GetParaCode(����ʱ��, gComInfo_����.ϵͳʱ��) & GetParaCode(Ȧ����ˮ��, gComInfo_����.Ȧ����ˮ��)
        If Not ���ýӿ�_����("qc") Then Exit Function
        If Not WriteICCard(lng����ID, gComInfo_����.Ȧ����) Then
        '�ظ����ó����ӿڣ�ֱ���ɹ�Ϊֹ
            Do While True
                If ���ýӿ�_����("qcrollback") Then Exit Do
            Loop
            Exit Function
        End If
    End If
    
    '��д�����¼
    '�ۼƽ���ͳ��=����Ȧ����ʻ��ۼ�����=ԭ���ڽ��ʻ��ۼ�֧��=�����ʻ�֧��
    '˳��ű�������ţ���ע�б��浱�εĽ�����
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & type_���� & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & gComInfo_����.�ʻ���� - gComInfo_����.Ȧ���� & "," & cur�����ʻ� & "," & gComInfo_����.Ȧ���� & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_����.֧����� & "," & gComInfo_����.֧����� - cur�����ʻ� & "," & 0 & "," & 0 & "," & 0 & ",0," & _
        0 & "," & cur�����ʻ� & ",'" & gComInfo_����.������ & "',null,null,'" & gComInfo_����.������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���������շ�����")
    
    '���ʹ���˸����ʻ�����Ҫ��ʹ������ϴ������ڸýӿ��޳�����������������ϴ���
    If cur�����ʻ� <> 0 Then
        If Not WriteICCard(lng����ID, cur�����ʻ� * -1) Then Exit Function
        gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & _
            GetParaCode(����������, gComInfo_����.����������) & GetParaCode(֧�����, cur�����ʻ�) & _
            GetParaCode(����ʱ��, gComInfo_����.ϵͳʱ��) & GetParaCode(������, gComInfo_����.������) & _
            GetParaCode(������, gComInfo_����.������) & GetParaCode(֧�����, gComInfo_����.֧�����)
        If Not ���ýӿ�_����("dataupload") Then
            '�˳�ǰ�����ڽ�ԭ
            'Modified by zyb 2003-10-25
            Do While 1
                If WriteICCard(lng����ID, cur�����ʻ�) Then
                    Exit Do
                Else
                    Err.Raise 9000, gstrSysName, "д�����ɹ�����忨��", vbInformation, gstrSysName
                End If
            Loop
            Exit Function
        End If
    End If
    
    '�ϴ�������ϸ��¼
    blnError = False
    gstrSQL = "Select Rownum ��ʶ��,A.ID,A.����ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ��,A.������ as ҽ��," & _
            "   A.����*A.���� as ����,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ��," & _
            "   A.�շ����,B.���� as ��Ŀ����,B.���� as ��Ŀ����,D.��Ŀ���� ҽ������," & _
            "   C.���� ��������,E.���� �ܵ�����" & _
            " From (Select * From ������ü�¼ Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0) A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,���ű� E " & _
            " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID(+) And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����=[1]" & _
            " Order by A.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ν��ʷ�����ϸ", type_����)
    With rsTemp
        Do While Not .EOF
            gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����������, gComInfo_����.����������) & _
                GetParaCode(������, gComInfo_����.������) & GetParaCode(������ˮ��, gComInfo_����.������ˮ�� & !��ʶ��) & _
                GetParaCode(������, gComInfo_����.������) & GetParaCode(֧�����, gComInfo_����.֧�����) & _
                GetParaCode(ҽ������, !ҽ������) & GetParaCode(��Ŀ����, !��Ŀ����) & GetParaCode(��Ŀ����, !��Ŀ����) & _
                GetParaCode(����, !����) & GetParaCode(�����ܶ�, !���ʽ��) & GetParaCode(��������, Nvl(!��������, "")) & _
                GetParaCode(����ҽ��, Nvl(!ҽ��, "")) & GetParaCode(�ܵ�����, Nvl(!�ܵ�����, "")) & GetParaCode(�ܵ�ҽ��, "") & _
                GetParaCode(����ʱ��, gComInfo_����.ϵͳʱ��)
            
            If ���ýӿ�_����("recipeinfotran") Then
                int�ϴ� = 1
            Else
                int�ϴ� = 0
                blnError = True
            End If
            
            'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
            'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rsTemp("ID") & ",NULL,NULL,NULL,NULL," & int�ϴ� & ",'" & gComInfo_����.������ˮ�� & !��ʶ�� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
            .MoveNext
        Loop
    End With

    Call ҽ����ֹ_����
    If blnError Then
        Err.Raise 9000, gstrSysName, "���ַ�����ϸδ��ȷ�ϴ����뵽�����ʻ������������ϴ���", vbInformation, gstrSysName
    End If
    �������_���� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    ����������_���� = False
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim str��Ժ����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    On Error GoTo errHand
    If Not ���� Then Exit Function
    
    gstrSQL = "Select A.�Ǽ��� ������,B.���� ��Ժ����,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') ��Ժ����ʱ��," & _
            " to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') ��Ժ����" & _
            " From ������ҳ A,���ű� B" & _
            " Where A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ��Ϣ", lng����ID, lng��ҳID)
    gComInfo_����.������ = rsTemp!������
    gComInfo_����.��Ժ���� = rsTemp!��Ժ����
    gComInfo_����.��Ժ���� = rsTemp!��Ժ����
    str��Ժ����ʱ�� = rsTemp!��Ժ����ʱ��
    gComInfo_����.��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, True, False)
    
    gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & _
        GetParaCode(����, gComInfo_����.����) & GetParaCode(����, gComInfo_����.����) & _
        GetParaCode(����������, gComInfo_����.����������) & GetParaCode(������, gComInfo_����.������) & _
        GetParaCode(֧�����, gComInfo_����.֧�����) & GetParaCode(���ֱ���, gComInfo_����.���ֱ���) & _
        GetParaCode(��Ժ����, gComInfo_����.��Ժ����) & GetParaCode(��Ժ���, gComInfo_����.��Ժ���) & _
        GetParaCode(��Ժ����, gComInfo_����.��Ժ����) & GetParaCode(��Ժ������, gComInfo_����.������) & _
        GetParaCode(��Ժ����ʱ��, str��Ժ����ʱ��)
    If Not ���ýӿ�_����("enterhospital") Then Exit Function
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & type_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    
    Call ҽ����ֹ_����
    ��Ժ�Ǽ�_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim str��Ժ����ʱ�� As String
    Dim blnҽ����Ժ As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ��
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
    If Not ���� Then Exit Function
    
    blnҽ����Ժ = False
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        '����ҽ���ĳ�Ժ�ӿ�
        blnҽ����Ժ = True
        gComInfo_����.֧����� = "0301"
        gComInfo_����.��Ժ��� = ��ȡ���Ժ���(lng����ID, lng��ҳID, False, False)
        'ȡ��Ժԭ��
        gstrSQL = "select decode(��Ժ��ʽ,'����',1,'תԺ',2,'����',3,9) ��Ժ��ʽ From ������ҳ " & _
                " Where ����ID = [1] And ��ҳID = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ��ʽ", lng����ID, lng��ҳID)
        gComInfo_����.��Ժԭ�� = rsTemp!��Ժ��ʽ

        gstrSQL = "select b.���� ��Ժ����,����,��ֹʱ��,����Ա����  " & _
                 " from ���˱䶯��¼ A,���ű� B  " & _
                 " where ����ID=[1] and ��ֹԭ��=1 " & _
                 " and A.����ID=B.ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ���", lng����ID)
        str��Ժ����ʱ�� = Format(rsTemp!��ֹʱ��, "yyyy-MM-dd HH:mm:ss")
        gComInfo_����.��Ժ���� = Format(rsTemp!��ֹʱ��, "yyyy-MM-dd")
        gComInfo_����.��Ժ���� = ToVarchar(rsTemp!��Ժ����, 20)
        gComInfo_����.������ = ToVarchar(rsTemp!����Ա����, 20)
        gComInfo_����.��Ժ��� = ToVarchar(��ȡ���Ժ���(lng����ID, lng��ҳID, False, False), 100)
        
        gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����, gComInfo_����.����) & _
                GetParaCode(����������, gComInfo_����.����������) & GetParaCode(������, gComInfo_����.������) & _
                GetParaCode(֧�����, gComInfo_����.֧�����) & GetParaCode(��Ժԭ��, gComInfo_����.��Ժԭ��) & _
                GetParaCode(��Ժ����, gComInfo_����.��Ժ����) & GetParaCode(��Ժ���, gComInfo_����.��Ժ���) & _
                GetParaCode(��Ժ����, gComInfo_����.��Ժ����) & GetParaCode(��Ժ������, gComInfo_����.������) & _
                GetParaCode(��Ժ����ʱ��, str��Ժ����ʱ��)
        If Not ���ýӿ�_����("leavehospital") Then Exit Function
    End If
    
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & type_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    
    Call ҽ����ֹ_����
    MsgBox IIf(blnҽ����Ժ, "ҽ����Ժ����ɹ���", "HIS��Ժ����ɹ���"), vbInformation, gstrSysName
    ��Ժ�Ǽ�_���� = True
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
    If Not ���� Then Exit Function
    gstrSQL = " Select Count(*) Records From סԺ���ü�¼ " & _
              " Where ����ID=[1] And ��ҳID=[2] And Nvl(��¼״̬,0)<>0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ժ���", lng����ID, lng��ҳID)
    If rsTemp!Records <> 0 Then
        MsgBox "�Ѿ����ڷ��ü�¼���������������Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & _
        GetParaCode(����������, gComInfo_����.����������) & _
        GetParaCode(������, gComInfo_����.������)
    If Not ���ýӿ�_����("enterhospitalrollback") Then Exit Function
    
    Call ҽ����ֹ_����
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & type_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_���� = True
End Function

Public Function ��Ժ�Ǽǳ���_����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '����δ����õĲ��˲�������HIS��Ժ��������Ϊ�Ѱ���ҽ����Ժ���������ٰ���HIS��Ժ
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        MsgBox "ҽ���ѳ�Ժ�Ĳ��˲���������Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & type_���� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_���� = True
End Function

Public Function סԺ�������_����(rsExse As Recordset, ByVal lng����ID As Long) As String
    Dim curTotal As Currency
    Dim lng��ҳID As Long
    Dim cur�����Ը� As Currency, cur�����ʻ� As Currency
    Dim str��Ժ��� As String, str������� As String
    Dim str����ʱ�� As String, str����ʱ�� As String
    Dim blnUpload As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    '�ӿڷ��صı������ȥ����סԺ�ڼ�����������Ļ��ܽ��󣬲��Ǳ��ε�ʵ�ʱ�����
    'rsExse��¼���е��ֶ��嵥
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    On Error GoTo errHand
    If Not ���� Then Exit Function
    
    If Not ���ýӿ�_����("getsysdate") Then Exit Function
    Call ��ȡ���������Ϣ(lng����ID)
    Call �������_����(True, "", True)
    cur�����ʻ� = gComInfo_����.�ʻ����
    
    gstrSQL = " Select B.סԺ���� ��ҳID,to_char(A.��Ժ����,'yyyy') ��Ժ��� " & _
              " From ������ҳ A,������Ϣ B" & _
              " Where B.����ID=[1] And A.��ҳID=B.סԺ���� And A.����ID=B.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ժʱ��", lng����ID)
    str��Ժ��� = rsTemp!��Ժ���
    lng��ҳID = rsTemp!��ҳID
    
    gComInfo_����.������ = 0               '�ϴ�������ϸʱ�������ű���Ҫ��Ϊ��
    gComInfo_����.֧����� = "0301"
    gComInfo_����.������ˮ�� = GetSequence(gint������ˮ��)
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str����ʱ�� = Format(gComInfo_����.ϵͳʱ��, "yyyy-MM-dd HH:mm:ss")
    str������� = Format(gComInfo_����.ϵͳʱ��, "yyyy")
    
    With rsExse
        Do While Not .EOF
            If IsNull(!ҽ����Ŀ����) Then
                MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            .MoveNext
        Loop
        .MoveFirst
        
        '�ϴ�������ϸ
        curTotal = 0
        Do While Not .EOF
            curTotal = curTotal + !���
            
            blnUpload = True
            If Not IsNull(!�Ƿ��ϴ�) Then
                blnUpload = (!�Ƿ��ϴ� = 0)
            End If
            
            If blnUpload Then
                
                'ȡ���շ�ϸĿ�ı���������
                gstrSQL = "Select ���� ��Ŀ���� ,���� ��Ŀ���� From �շ�ϸĿ Where ID = [1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���շ�ϸĿ�ı���������", CLng(!�շ�ϸĿID))
                
                gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����������, gComInfo_����.����������) & _
                    GetParaCode(������, gComInfo_����.������) & GetParaCode(������ˮ��, gComInfo_����.������ˮ�� & .AbsolutePosition) & _
                    GetParaCode(������, gComInfo_����.������) & GetParaCode(֧�����, gComInfo_����.֧�����) & _
                    GetParaCode(ҽ������, !ҽ����Ŀ����) & GetParaCode(��Ŀ����, rsTemp!��Ŀ����) & GetParaCode(��Ŀ����, rsTemp!��Ŀ����) & _
                    GetParaCode(����, !����) & GetParaCode(�����ܶ�, !���) & GetParaCode(��������, Nvl(!��������, "")) & _
                    GetParaCode(����ҽ��, Nvl(!ҽ��, "")) & GetParaCode(�ܵ�����, "") & GetParaCode(�ܵ�ҽ��, "") & _
                    GetParaCode(����ʱ��, Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
                If ���ýӿ�_����("recipeinfotran") Then
                    '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
                    'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                    gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
                Else
                    Exit Function
                End If
            End If
            .MoveNext
        Loop
        
        gComInfo_����.֧����� = curTotal
        '������㣨���ص����ݼ�ȥ���ν������ݣ��͵��ڱ��ε���ʵ�������ݣ�
        gComInfo_����.������ = GetSequence(gint������)
        gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����, gComInfo_����.����) & _
            GetParaCode(����, gComInfo_����.����) & GetParaCode(������, gComInfo_����.������) & _
            GetParaCode(������, gComInfo_����.������) & GetParaCode(�����ܶ�, gComInfo_����.֧�����) & _
            GetParaCode(����������, gComInfo_����.����������) & GetParaCode(֧�����, gComInfo_����.֧�����) & _
            GetParaCode(���ֱ���, gComInfo_����.���ֱ���) & GetParaCode(������, "900090009000999") & _
            GetParaCode(����ʱ��, str����ʱ��) & GetParaCode(����ʱ��, str����ʱ��) & _
            GetParaCode(����Ƚ����־, IIf(str��Ժ��� <> str�������, 1, 0)) & GetParaCode(��ʵ�����־, 0)
        If Not ���ýӿ�_����("ExpenseReckoning") Then Exit Function
        
        Call ���÷ָ�(lng����ID, lng��ҳID)
    End With
    
    '���ؽ��㷽ʽ
    cur�����Ը� = gComInfo_����.֧����� - gComInfo_����.ͳ��֧�� 'gComInfo_����.ͳ���Ը� + gComInfo_����.�����Ը� + gComInfo_����.�����Ը� + gComInfo_����.�ҹ��Ը�
    If gComInfo_����.ͳ��֧�� <> 0 Then
        סԺ�������_���� = "ҽ������;" & gComInfo_����.ͳ��֧�� & ";0"
    End If
    'ֻ�г�Ժ���������ʹ�ø����ʻ�
    If ҽ�������Ѿ���Ժ(lng����ID) Then
        If cur�����ʻ� <> 0 Then
            If cur�����ʻ� > cur�����Ը� Then
                cur�����ʻ� = cur�����Ը�
            End If
            If cur�����ʻ� < 0 Then cur�����ʻ� = 0
            סԺ�������_���� = סԺ�������_���� & IIf(סԺ�������_���� = "", "", "|") & "�����ʻ�;" & cur�����ʻ� & ";1"
        End If
    End If
    
    Call ҽ����ֹ_����
    If סԺ�������_���� = "" Then סԺ�������_���� = "�����ʻ�;0;1"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    Dim cur�����ʻ� As Currency
    Dim lng��ҳID As Long
    Dim blnError As Boolean
    Dim str��Ժ��� As String, str������� As String
    Dim str����ʱ�� As String, str����ʱ�� As String
    Dim str������ As String
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
        '������㣨���ص����ݼ�ȥ���ν������ݣ��͵��ڱ��ε���ʵ�������ݣ�
    On Error GoTo errHand
    If Not ���� Then Exit Function
    If Not ���ýӿ�_����("getsysdate") Then Exit Function
    Call �������_����(True, "", True)
    cur�����ʻ� = gComInfo_����.�ʻ����
    
    gstrSQL = " Select B.סԺ���� ��ҳID,to_char(A.��Ժ����,'yyyy') ��Ժ��� " & _
              " From ������ҳ A,������Ϣ B" & _
              " Where B.����ID=[1] And A.��ҳID=B.סԺ���� And A.����ID=B.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ժʱ��", lng����ID)
    str��Ժ��� = rsTemp!��Ժ���
    lng��ҳID = rsTemp!��ҳID
    
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str����ʱ�� = Format(gComInfo_����.ϵͳʱ��, "yyyy-MM-dd HH:mm:ss")
    str������� = Format(gComInfo_����.ϵͳʱ��, "yyyy")
    
    'סԺ����
    gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����, gComInfo_����.����) & _
        GetParaCode(����, gComInfo_����.����) & GetParaCode(������, gComInfo_����.������) & _
        GetParaCode(������, gComInfo_����.������) & GetParaCode(�����ܶ�, gComInfo_����.֧�����) & _
        GetParaCode(����������, gComInfo_����.����������) & GetParaCode(֧�����, gComInfo_����.֧�����) & _
        GetParaCode(���ֱ���, gComInfo_����.���ֱ���) & GetParaCode(������, 1E+21) & _
        GetParaCode(����ʱ��, str����ʱ��) & GetParaCode(����ʱ��, str����ʱ��) & _
        GetParaCode(����Ƚ����־, IIf(str��Ժ��� <> str�������, 1, 0)) & GetParaCode(��ʵ�����־, 1)
    If Not ���ýӿ�_����("ExpenseReckoning") Then Exit Function
    Call ���÷ָ�(lng����ID, lng��ҳID)
    
    '��ȡ���θ����ʻ�֧����
    gstrSQL = "Select Nvl(A.��Ԥ��,0) �����ʻ� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=" & type_���� & _
        " And A.���㷽ʽ in ('�����ʻ�') And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���θ����ʻ�֧����", lng����ID)
    cur�����ʻ� = 0
    If Not rsTemp.EOF Then
        cur�����ʻ� = rsTemp!�����ʻ�
    End If
    
    '��д���ս����¼
    '�ۼƽ���ͳ��=����Ȧ����ʻ��ۼ�����=ԭ���ڽ��ʻ��ۼ�֧��=�����ʻ�֧��
    '˳��ű�������ţ���ע�б��浱�εĽ�����
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & type_���� & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & gComInfo_����.�ʻ���� & "," & cur�����ʻ� & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & gComInfo_����.�����Ը� & "," & _
        gComInfo_����.֧����� & "," & 0 & "," & gComInfo_����.�ҹ��Ը� & "," & gComInfo_����.ͳ��֧�� + gComInfo_����.ͳ���Ը� & "," & gComInfo_����.ͳ��֧�� & ",0," & _
        gComInfo_����.�����Ը� & "," & cur�����ʻ� & ",'" & gComInfo_����.������ & "',null,null,'" & gComInfo_����.������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��������")
    
    '���ʹ���˸����ʻ�����Ҫ��ʹ������ϴ������ڸýӿ��޳��������������������ϴ���
    'ʹ�ø����ʻ��ĵط���ע��д��
    If cur�����ʻ� <> 0 Then
        '����������ķ�ʽ�¸����ʻ�
        blnError = False
        If Not WriteICCard(lng����ID, cur�����ʻ� * -1, True) Then blnError = True
        
        str������ = gComInfo_����.������
        gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����, gComInfo_����.����) & _
            GetParaCode(����������, gComInfo_����.����������)
        If ���ýӿ�_����("identifyinfogetting") Then
            gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & _
                GetParaCode(����������, gComInfo_����.����������) & GetParaCode(֧�����, cur�����ʻ�) & _
                GetParaCode(����ʱ��, gComInfo_����.ϵͳʱ��) & GetParaCode(������, gComInfo_����.������) & _
                GetParaCode(������, gComInfo_����.������) & GetParaCode(֧�����, "0201")
            If blnError = False Then
                If Not ���ýӿ�_����("dataupload") Then
                    'Modified by zyb 2003-10-25
                    Do While 1
                        If WriteICCard(lng����ID, cur�����ʻ�, True) Then
                            Exit Do
                        Else
                            Err.Raise 9000, gstrSysName, "д�����ɹ�����忨��", vbInformation, gstrSysName
                        End If
                    Loop
                    blnError = True
                End If
            End If
        Else
            blnError = True
        End If
        
        gComInfo_����.������ = str������
        If blnError Then
            '�����˳�����ʾ���ֽ�֧��
            gcnOracle.Execute "Delete ���ս����¼ Where ����=2 And ����=" & type_����
            
            gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & type_���� & "," & lng����ID & "," & _
                Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
                0 & "," & "NULL" & "," & 0 & "," & 0 & "," & gComInfo_����.�����Ը� & "," & _
                gComInfo_����.֧����� & "," & 0 & "," & gComInfo_����.�ҹ��Ը� & "," & gComInfo_����.ͳ��֧�� + gComInfo_����.ͳ���Ը� & "," & gComInfo_����.ͳ��֧�� & ",0," & _
                gComInfo_����.�����Ը� & "," & 0 & ",'" & gComInfo_����.������ & "',null,null,'" & gComInfo_����.������ & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��������")
            Err.Raise 9000, gstrSysName, "�¸����ʻ�ʧ�ܣ�ע����ȡ�ֽ�" & Format(cur�����ʻ�, "#####0.00;-#####0.00; ;") & "��", vbInformation, gstrSysName
        End If
    End If
    Call ҽ����ֹ_����
    
    'ֻ�г�Ժ���������ʹ�ø����ʻ�
    If ҽ�������Ѿ���Ժ(lng����ID) Then
        Call ��Ժ�Ǽ�_����(lng����ID, lng��ҳID)
    End If
    סԺ����_���� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_����(lng����ID As Long) As Boolean
    Dim lng����ID As Long
    Dim str�˵���� As String
    Dim rsTemp As New ADODB.Recordset
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '      4)ֻ�����ϵ�����������Ա�Ľ��ʵ���
    '----------------------------------------------------------------
    On Error GoTo errHand
    If Not ���� Then Exit Function
    
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    
    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "Select * " & _
              "  From ���ս����¼ Where ����=2 and ��¼ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    
    Call ��ȡ���������Ϣ(rsTemp!����ID)
    gComInfo_����.������ = GetSequence(gint������)   'ȡ�µĽ�����
    gComInfo_����.������ = Nvl(rsTemp!˳���, "")      'ȡ��ʱ�ľ�����
    str�˵���� = Nvl(rsTemp!��ע, "")              'ȡ��ʱ�Ľ�����
    
    '��д�����¼
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & type_���� & "," & rsTemp!����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & Nvl(rsTemp!�ʻ��ۼ�����, 0) * -1 & "," & Nvl(rsTemp!�ʻ��ۼ�֧��, 0) * -1 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & Nvl(rsTemp!ʵ������, 0) * -1 & "," & _
        Nvl(rsTemp!�������ý��, 0) * -1 & "," & 0 & "," & Nvl(rsTemp!�����Ը����, 0) * -1 & "," & Nvl(rsTemp!����ͳ����, 0) * -1 & "," & Nvl(rsTemp!ͳ�ﱨ�����, 0) * -1 & ",0," & _
        Nvl(rsTemp!�����Ը����, 0) * -1 & "," & Nvl(rsTemp!�����ʻ�֧��, 0) * -1 & ",'" & gComInfo_����.������ & "',null,null,'" & gComInfo_����.������ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��������")
    
    '���ó�������ӿ�
    gstrPara_���� = GetParaCode(������, gComInfo_����.������) & GetParaCode(������, gComInfo_����.������) & _
            GetParaCode(�˵����, str�˵����) & GetParaCode(����������, gComInfo_����.����������)
    If Not ���ýӿ�_����("expenserollback") Then Exit Function
    
    Call ҽ����ֹ_����
    סԺ�������_���� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Private Sub ���÷ָ�(ByVal lng����ID As Long, lng��ҳID As Long)
    Dim cur�����ܶ� As Currency, curͳ���ܶ� As Currency, curͳ���Ը��ܶ� As Currency '������;��������ܶ�;������;����ͳ��֧���ܶ�
    Dim cur�����ܶ� As Currency, cur�ҹ��Ը��ܶ� As Currency, cur�����Ը��ܶ� As Currency
    Dim rsTemp As New ADODB.Recordset
    
    'ȡ����סԺ�ڼ䣬������;����ķ����ܶͳ���ܶ�
    gstrSQL = "SELECT SUM(�������ý��) �������ý��,SUM(����ͳ����) ����ͳ����,SUM(ͳ�ﱨ�����) ͳ�ﱨ�����, " & _
             " SUM(�����Ը����) �����Ը����,SUM(ʵ������) ʵ������,SUM(�����Ը����) �����Ը����" & _
             " FROM  " & _
             "      (SELECT Distinct ����ID,����ID FROM סԺ���ü�¼ " & _
             "      WHERE ����ID=[1] AND ��ҳID= [2]" & _
             "      ) A,���ս����¼ B " & _
             " WHERE A.����ID=B.����ID AND B.��¼ID=A.����ID AND B.����=[3] AND B.����=2 " & _
             " GROUP BY A.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����סԺ�ڼ���������ķ����ܶͳ�ﱨ���ܶ�", lng����ID, lng��ҳID, type_����)
    If Not rsTemp.EOF Then
        cur�����ܶ� = rsTemp!�������ý��
        curͳ���ܶ� = rsTemp!ͳ�ﱨ�����
        curͳ���Ը��ܶ� = rsTemp!����ͳ���� - rsTemp!ͳ�ﱨ�����
        cur�����ܶ� = rsTemp!ʵ������
        cur�ҹ��Ը��ܶ� = rsTemp!�ҹ��Ը����
        cur�����Ը��ܶ� = rsTemp!�����Ը����
    Else
        cur�����ܶ� = 0
        curͳ���ܶ� = 0
        curͳ���Ը��ܶ� = 0
        cur�����ܶ� = 0
        cur�ҹ��Ը��ܶ� = 0
        cur�����Ը��ܶ� = 0
    End If
    
    gComInfo_����.�ҹ��Ը� = CCur(Format(gComInfo_����.�ҹ��Ը� - cur�ҹ��Ը��ܶ�, "#####0.00;-#####0.00;0;"))
    gComInfo_����.����ͳ�� = CCur(Format(gComInfo_����.����ͳ�� - (curͳ���ܶ� + curͳ���Ը��ܶ�), "#####0.00;-#####0.00;0;"))
    gComInfo_����.�����Ը� = CCur(Format(gComInfo_����.�����Ը� - cur�����ܶ�, "#####0.00;-#####0.00;0;"))
    gComInfo_����.ͳ��֧�� = CCur(Format(gComInfo_����.ͳ��֧�� - curͳ���ܶ�, "#####0.00;-#####0.00;0;"))
    gComInfo_����.ͳ���Ը� = CCur(Format(gComInfo_����.ͳ���Ը� - curͳ���Ը��ܶ�, "#####0.00;-#####0.00;0;"))
    gComInfo_����.�����Ը� = CCur(Format(gComInfo_����.�����Ը� - cur�����Ը��ܶ�, "#####0.00;-#####0.00;0;"))
End Sub

Public Function ���ýӿ�_����(ByVal strFunction As String) As Boolean
    '���ýӿڹ���
    On Error GoTo errHand
    
    Select Case strFunction
    Case "getsysdate"                   '��ȡϵͳʱ��
        Dim strSysdate As Date
        Call objCom.getsysdate(strSysdate, gComInfo_����.ִ�н��)
        gComInfo_����.ϵͳʱ�� = Format(strSysdate, "yyyy-MM-dd HH:mm:ss")
    Case "identifyinfogetting"          '�����֤
        Call objCom.identifyinfogetting(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, _
        gComInfo_����.ִ�н��, gComInfo_����.������, gComInfo_����.����, gComInfo_����.�Ա�, strSysdate, gComInfo_����.���֤��, _
        gComInfo_����.��Ա���, gComInfo_����.��λ����, gComInfo_����.����, gComInfo_����.��Ȧ����)
        gComInfo_����.�������� = Format(strSysdate, "yyyy-MM-dd")
    Case "modifypassword"               '�޸�����
        Call objCom.modifypassword(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��)
    Case "dataupload"                   '�ϴ��������ݣ��ʻ���
        Call objCom.dataupload(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��)
    Case "recipeinfotran"               '�ϴ�������ϸ
        Call objCom.recipeinfotran(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��)
    Case "qc"                           'IC��Ȧ��
        Call objCom.qc(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��)
    Case "qcrollback"                   'IC��Ȧ�泷��
        Call objCom.qcrollback(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��)
    Case "audittreatment"               '�ʸ����������˶�
        Call objCom.audittreatment(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��, _
        gComInfo_����.������, gComInfo_����.����, gComInfo_����.�Ա�, strSysdate, gComInfo_����.���֤��, gComInfo_����.��Ա���, _
        gComInfo_����.����, gComInfo_����.��������, gComInfo_����.����ͳ���޶�, gComInfo_����.����ͳ�ﱨ������, gComInfo_����.��λ����)
        gComInfo_����.�������� = Format(strSysdate, "yyyy-MM-dd")
    Case "enterhospital"                '��Ժ����
        Call objCom.enterhospital(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��)
    Case "enterhospitalrollback"        '��Ժ������
        Call objCom.enterhospitalrollback(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��)
    Case "ExpenseReckoning"             'סԺ����/�������
        Call objCom.ExpenseReckoning(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��, _
        gComInfo_����.�ҹ��Ը�, gComInfo_����.����ͳ��, gComInfo_����.�����Ը�, gComInfo_����.ͳ��֧��, gComInfo_����.ͳ���Ը�, gComInfo_����.�����Ը�, gComInfo_����.��������)
    Case "leavehospital"                '��Ժ����
        Call objCom.leavehospital(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��)
    Case "expenserollback"              'סԺ���㳷��
        Call objCom.expenserollback(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��)
    Case "checkaccount"                 '�˶Ը����ʻ�֧��
        Call objCom.checkaccount(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��, _
        gComInfo_����.�˶�_��¼��, gComInfo_����.�˶�_�ʻ�֧���ܶ�)
    Case "checkexpense"                 '�˶����������Ϣ
        Call objCom.checkexpense(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��, _
        gComInfo_����.�˶�_��¼��, gComInfo_����.�˶�_ҽ�Ʒ��ܶ�, gComInfo_����.�˶�_ȫ�Է��ܶ�, gComInfo_����.�˶�_�ҹ��Է��ܶ�, _
        gComInfo_����.�˶�_����ͳ���ܶ�, gComInfo_����.�˶�_�ʻ�֧���ܶ�, gComInfo_����.�˶�_�ֽ�֧���ܶ�, gComInfo_����.�˶�_�����Ը��ܶ�)
    Case "checkenterleavehosptinfo"     '�˶�סԺ�˴�
        Call objCom.checkenterleavehosptinfo(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��, _
        gComInfo_����.�˶�_��Ժ����, gComInfo_����.�˶�_��Ժ����)
    Case "checkrecipeinfo"              '�˶Է�����ϸ
        Call objCom.checkrecipeinfo(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��, _
        gComInfo_����.�˶�_����, gComInfo_����.�˶�_����, gComInfo_����.�˶�_ҽ�Ʒ��ܶ�, gComInfo_����.�˶�_�Ը�����, gComInfo_����.�˶�_ȫ�Է��ܶ�, _
        gComInfo_����.�˶�_�ҹ��Է��ܶ�, gComInfo_����.�˶�_����ͳ���ܶ�)
    Case "checkexpensereckoninginfo"    '�˶�סԺ���ý�����
        Call objCom.checkexpensereckoninginfo(gstrPara_����, gComInfo_����.��������, gComInfo_����.�������, gComInfo_����.����������Ϣ, gComInfo_����.ִ�н��, _
        gComInfo_����.�˶�_����֧�����, gComInfo_����.�˶�_ͳ���Ը����, gComInfo_����.�˶�_ͳ��֧�����)
    End Select
    
    If gComInfo_����.ִ�н�� = 0 Then
        MsgBox gComInfo_����.�������� & "|" & gComInfo_����.����������Ϣ & "|������룺" & gComInfo_����.�������, vbInformation, gstrSysName
        Exit Function
    End If
    
    ���ýӿ�_���� = True
    Exit Function
errHand:
    MsgBox "����ִ��ʧ�ܣ�", vbInformation, gstrSysName
End Function

Public Function ��������() As Boolean
    On Error GoTo errHand
    
    Set objCom = CreateObject("pb80.n_center_interface.1.0")
    If objCom Is Nothing Then
        Exit Function
    End If
    
    �������� = True
    Exit Function
errHand:
End Function

Public Function ��ȡ���������Ϣ(ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '��ȡҽ�����������Ϣ�������¹��ýṹ��
    gstrSQL = " Select ����,����,ҽ���� ���˱��,˳��� ������,Nvl(����ID,0) ����ID From �����ʻ�" & _
            " Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����˵������Ϣ", type_����, lng����ID)
    If rsTemp.EOF Then Exit Function
    
    gComInfo_����.���˱�� = rsTemp!���˱��
    gComInfo_����.���� = rsTemp!����
    gComInfo_����.������ = rsTemp!������
    gComInfo_����.���� = rsTemp!����
    Call ��ȡ���ֱ���(rsTemp!����ID)
    ��ȡ���������Ϣ = True
End Function

Public Sub ��ȡ���ֱ���(ByVal lng����ID As Long)
    Dim rsTemp As New ADODB.Recordset
    '�жϲ��������������ز������ֱ���="900001"������="900002"
    gComInfo_����.���ֱ��� = "900002"
    gstrSQL = "Select nvl(���,0) ��� From ���ղ��� Where ����=[1] And ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", type_����, lng����ID)
    If Not rsTemp.EOF Then
        If rsTemp!��� <> 0 Then
            gComInfo_����.���ֱ��� = "900001"
        End If
    End If
End Sub

Public Function GetSequence(Optional ByVal intType As Integer = 1) As String
    'intTYPE�ĺ��壺
    '1=Ȧ����ˮ��=ҽԺ���+7λ��ˮ�ţ�����Ȧ�漰Ȧ��ʧ��ʱʹ�ã���Ҫ���棩
    '2=������ˮ��=15λ������ˮ�ţ����Բ����������ݿ��У�
    '3=������=J+YYMMDD+ҽԺ���+4λ��ˮ�ţ�ͬһ�����Ų������ظ�����Ҫ���棩
    Dim strSequence As String
    Dim strHour As String, strMinute As String, strSecond As String
    Dim rsTemp As New ADODB.Recordset
    
    '������������Ե���ˮ��
'    gComInfo_����.ϵͳʱ�� = Now
    strHour = Format(gComInfo_����.ϵͳʱ��, "HH")
    strMinute = Mid(gComInfo_����.ϵͳʱ��, 15, 2) ' Format(gComInfo_����.ϵͳʱ��, "mm")
    strSecond = Format(gComInfo_����.ϵͳʱ��, "ss")
    Select Case intType
'    Case 1      '�͵�����ˮ�Ų�������һ�£�A(��)+A(��)+A(��)+HHmm��
'        strSequence = ��ȡһλ��ʾ��(1) & ��ȡһλ��ʾ��(2) & ��ȡһλ��ʾ��(3) & ��ȡһλ��ʾ��(4, strHour) & ��ȡһλ��ʾ��(4, strMinute) & strSecond
    Case 1, 2     '�Ե�ǰϵͳʱ���yyMMddHHmmss������λ�Ե�ǰ��¼������ֶε�ֵ���
        strSequence = Format(gComInfo_����.ϵͳʱ��, "yyMMddHHmmss")
    Case 3      '4λ��ˮ���Ե�ʱ��HHmmΪ��ʶ
        strSequence = "J" & Format(gComInfo_����.ϵͳʱ��, "yyMMdd") & gComInfo_����.���������� & ��ȡһλ��ʾ��(4, strHour) & ��ȡһλ��ʾ��(4, strMinute) & strSecond
    End Select
    GetSequence = strSequence
End Function

Public Function ��ȡһλ��ʾ��(Optional ByVal intType As Integer = 1, Optional ByVal lngData As Long = 0) As String
    Dim lngMid As Long
    '����һλ����ݡ��·ݻ����ӵı�ʾ�ַ�����1-���;2-�·�;3-����
    Select Case intType
    Case 1
        lngMid = Format(gComInfo_����.ϵͳʱ��, "yyyy")
        lngMid = lngMid - 2000
    Case 2
        lngMid = Format(gComInfo_����.ϵͳʱ��, "MM")
    Case 3
        lngMid = Format(gComInfo_����.ϵͳʱ��, "dd")
    Case Else
        lngMid = lngData
    End Select
    If lngMid >= 10 Then
        ��ȡһλ��ʾ�� = Chr(lngMid - 10 + 65)
    Else
        ��ȡһλ��ʾ�� = lngMid
    End If
End Function

Public Function GetParaCode(ByVal intType As Integer, ByVal strData As Variant) As String
'��Ŀ����    ����ֵ       ��Ŀ����        ��Ŀ����
'AKA123         0       ���ⲡ�ֱ�־    �����ⲡ��
'AKA123         1       ���ⲡ�ֱ�־    ���ⲡ��
'AKC021         11      ҽ����Ա���    ��ְ
'AKC021         12      ҽ����Ա���    ��ְ����פ��
'AKC021         21      ҽ����Ա���    ����
'AKC021         22      ҽ����Ա���    ������ذ���
'AKC021         31      ҽ����Ա���    ����
'AAC004         1       �Ա�            ��
'AAC004         2       �Ա�            Ů
'AKA130         0101    ֧�����        ҩ�깺ҩ
'AKA130         0201    ֧�����        ��ͨ����
'AKA130         0205    ֧�����        ���ⲡ������
'AKA130         0301    ֧�����        ��ͨסԺ
'AKA130         0302    ֧�����        �Ǿ�ס��סԺ
'AKA130         0304    ֧�����        ͳ������תԺ
'AKA130         0305    ֧�����        ת��סԺ
'AKA130         0401    ֧�����        �����סԺ(ֻ����ka10k1��ʹ��)
'AKA130         0901    ֧�����        ����ҽ�ƻ���Ԥ��
'AKA130         0902    ֧�����        ����ҽ�ƻ�������
'AKA130         0701    ֧�����        ����Ա�������
'YKA002         2000001 ҽ����Ŀ����    ����
'YKA002         2000002 ҽ����Ŀ����    ����
'YKA002         2000003 ҽ����Ŀ����    �Է�
'YKA002         2000004 ҽ����Ŀ����    һ��ҽԺ��λ��
'YKA002         2000005 ҽ����Ŀ����    ����ҽԺ��λ��
'YKA002         2000006 ҽ����Ŀ����    ����ҽԺ��λ��
'YKA026         900001  ���ֱ���        ҽ���涨סԺ���֣����ߣ�0��
'YKA026         900002  ���ֱ���        �ǹ涨סԺ���֣��������ߣ�
'AKC195         1       ��Ժԭ��        ����
'AKC195         2       ��Ժԭ��        תԺ
'AKC195         3       ��Ժԭ��        ����
'AKC195         9       ��Ժԭ��        ����
    
    Dim strValue As String
    Select Case intType
    Case ���˱��
        strValue = "aac001"
    Case ����
        strValue = "yac005"
    Case ����������
        strValue = "akb020"
    Case ����
        strValue = "ykc005"
    Case ������
        strValue = "new_ykc005"
    Case ֧�����
        strValue = "defrayamount"
    Case ������
        strValue = "aae011"
    Case ����ʱ��
        strValue = "aae036"
    Case ������
        strValue = "akc190"
    Case ������
        strValue = "yka103"
    Case ֧�����
        strValue = "aka130"
    Case ������ˮ��
        strValue = "yka105"
    Case ҽ������
        strValue = "yka002"
    Case ��Ŀ����
        strValue = "yka094"
    Case ��Ŀ����
        strValue = "yka095"
    Case ����
        strValue = "akc226"
    Case �����ܶ�
        strValue = "yka055"
    Case ��������
        strValue = "yka098"
    Case ����ҽ��
        strValue = "yka099"
    Case �ܵ�����
        strValue = "yka101"
    Case �ܵ�ҽ��
        strValue = "yka102"
    Case Ȧ������
        strValue = "yka151"
    Case Ȧ����
        strValue = "yka152"
    Case Ȧ��ʱ��
        strValue = "yka153"
    Case Ȧ����ˮ��
        strValue = "ykc019"
    Case ���ֱ���
        strValue = "yka026"
    Case ������ȡʱ��, ����ʱ��, ��Ժ����
        strValue = "akc194"
    Case ��Ժ����
        strValue = "akc192"
    Case ��Ժ���
        strValue = "akc193"
    Case ��Ժ����
        strValue = "ykc011"
    Case ��Ժ������
        strValue = "ykc013"
    Case ��Ժ����ʱ��
        strValue = "ykc014"
    Case ����Ƚ����־
        strValue = "ykc007"
    Case ��ʵ�����־
        strValue = "mnjs"
    Case ��Ժԭ��
        strValue = "akc195"
    Case ��Ժ���
        strValue = "akc196"
    Case ��Ժ����
        strValue = "ykc015"
    Case ��Ժ������
        strValue = "ykc017"
    Case ��Ժ����ʱ��
        strValue = "ykc018"
    Case �˵����
        strValue = "yka198"
    Case ������_��ʼ
        strValue = "akc190_Begin"
    Case ������_��ֹ
        strValue = "akc190_End"
    Case ������ˮ��_��ʼ
        strValue = "yka105_begin"
    Case ������ˮ��_��ֹ
        strValue = "yka105_end"
    Case ������_��ʼ
        strValue = "yka103_begin"
    Case ������_��ֹ
        strValue = "yka103_end"
    Case ��Ա���
        strValue = "akc021"
    End Select
    
    GetParaCode = "<" & strValue & ">" & strData & "</" & strValue & ">"
End Function

Public Sub �˶��ʻ�֧��_����()
    Dim cur�ʻ�֧���ܶ� As Currency
    Dim str��ʼ���� As String, str�������� As String
    Dim str��ʼ������ As String, str���������� As String
    Dim str��ʼ������ As String, str���������� As String
    Dim rsAccount As New ADODB.Recordset
    
    On Error GoTo errHand
    '��ȡ��ѯ����
    If frm���ڷ�Χ_����.Show_ME(str��ʼ����, str��������) = False Then Exit Sub
    '������ȡ�ʻ�֧���ܶ�
    gstrSQL = "Select SUM(��Ԥ��) �����ʻ� " & _
        " From ����Ԥ����¼ " & _
        " Where ���㷽ʽ in ('�����ʻ�') " & _
        " And �տ�ʱ�� Between [1] And [2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "ͳ���ʻ�֧���ܶ�", CDate(str��ʼ����), CDate(str��������))
    If Not rsAccount.EOF Then
        cur�ʻ�֧���ܶ� = Nvl(rsAccount!�����ʻ�, 0)
    End If
    
    If Not ���� Then Exit Sub
    '��ȡָ�����ڷ�Χ�ڵĿ�ʼ������������
    If Not FUNC_������(str��ʼ����, str��������, str��ʼ������, str����������, _
    str��ʼ������, str����������, False) Then Exit Sub
    '��ȡҽ�����ķ��ص��ʻ�֧���ܶ�
    gstrPara_���� = GetParaCode(����������, gComInfo_����.����������) & _
        GetParaCode(������_��ʼ, str��ʼ������) & GetParaCode(������_��ֹ, str����������)
    Call ���ýӿ�_����("checkaccount")
    Call ҽ����ֹ_����
    
    If Format(cur�ʻ�֧���ܶ�, "#####0.00;-#####0.00;0;") <> Format(gComInfo_����.�˶�_�ʻ�֧���ܶ�, "#####0.00;-#####0.00;0;") Then
        MsgBox "�����أ��ʻ�֧���ܶ" & cur�ʻ�֧���ܶ� & String(4, " ") & "��ҽ�����ʻ�֧���ܶ" & gComInfo_����.�˶�_�ʻ�֧���ܶ�
    Else
        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub �˶��������_����()
    Dim cur�������ý�� As Currency, cur�����Ը���� As Currency, cur�����ʻ�֧�� As Currency
    Dim str��ʼ���� As String, str�������� As String
    Dim str��ʼ������ As String, str���������� As String
    Dim str��ʼ������ As String, str���������� As String
    Dim rsAccount As New ADODB.Recordset
    
    On Error GoTo errHand
    '��ȡ��ѯ����
    If frm���ڷ�Χ_����.Show_ME(str��ʼ����, str��������) = False Then Exit Sub
    '������ȡ�ʻ�֧���ܶ�
    gstrSQL = "Select SUM(�������ý��) �������ý��,SUM(�����Ը����) �����Ը���� ,SUM(�����ʻ�֧��) �����ʻ�֧��" & _
        " From ���ս����¼ " & _
        " Where ����=1 " & _
        " And ����ʱ�� Between [1] And [2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "ͳ���������", CDate(str��ʼ����), CDate(str��������))
    If Not rsAccount.EOF Then
        cur�������ý�� = Nvl(rsAccount!�������ý��, 0)
        cur�����Ը���� = Nvl(rsAccount!�����Ը����, 0)
        cur�����ʻ�֧�� = Nvl(rsAccount!�����ʻ�֧��, 0)
    End If
    
    If Not ���� Then Exit Sub
    '��ȡָ�����ڷ�Χ�ڵĿ�ʼ������������
    If Not FUNC_������(str��ʼ����, str��������, str��ʼ������, str����������, _
    str��ʼ������, str����������, False) Then Exit Sub
    '��ȡҽ�����ķ��ص�ͳ��֧����ͳ���Ը��ܶ�
    gstrPara_���� = GetParaCode(����������, gComInfo_����.����������) & _
        GetParaCode(������_��ʼ, str��ʼ������) & GetParaCode(������_��ֹ, str����������) '& _
        GetParaCode(������_��ʼ, str��ʼ������) & GetParaCode(������_��ֹ, str����������)
    Call ���ýӿ�_����("checkexpense")
    Call ҽ����ֹ_����
    
    If Not (Format(cur�������ý��, "#####0.00;-#####0.00;0;") = Format(gComInfo_����.�˶�_ҽ�Ʒ��ܶ�, "#####0.00;-#####0.00;0;") _
    And Format(cur�����Ը����, "#####0.00;-#####0.00;0;") = Format(gComInfo_����.�˶�_�ҹ��Է��ܶ�, "#####0.00;-#####0.00;0;") _
    And Format(cur�����ʻ�֧��, "#####0.00;-#####0.00;0;") = Format(gComInfo_����.�˶�_�ʻ�֧���ܶ�, "#####0.00;-#####0.00;0;")) Then
        MsgBox "�����أ�ҽ�Ʒ��ܶ" & cur�������ý�� & String(4, " ") & "��ҽ����ҽ�Ʒ��ܶ" & gComInfo_����.�˶�_ҽ�Ʒ��ܶ� & vbCrLf & _
               "�����أ������Ը��ܶ" & cur�����Ը���� & String(4, " ") & "��ҽ���������Ը��ܶ" & gComInfo_����.�˶�_�ҹ��Է��ܶ� & vbCrLf & _
               "�����أ������ʻ�֧���ܶ" & cur�����ʻ�֧�� & String(4, " ") & "��ҽ���������ʻ�֧���ܶ" & gComInfo_����.�˶�_�ʻ�֧���ܶ�
    Else
        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub �˶�סԺ����_����()
    Dim curͳ��֧���ܶ� As Currency, curͳ���Ը��ܶ� As Currency
    Dim str��ʼ���� As String, str�������� As String
    Dim str��ʼ������ As String, str���������� As String
    Dim str��ʼ������ As String, str���������� As String
    Dim rsAccount As New ADODB.Recordset
    
    On Error GoTo errHand
    '��ȡ��ѯ����
    If frm���ڷ�Χ_����.Show_ME(str��ʼ����, str��������) = False Then Exit Sub
    '������ȡ�ʻ�֧���ܶ�
    gstrSQL = "Select SUM(ͳ�ﱨ�����) ͳ��֧�����,SUM(����ͳ����-ͳ�ﱨ�����) ͳ���Ը���� " & _
        " From ���ս����¼ " & _
        " Where ����=2 " & _
        " And ����ʱ�� Between [1] And [2]"
    Set rsAccount = zlDatabase.OpenSQLRecord(gstrSQL, "ͳ��ͳ��֧���ܶ�", CDate(str��ʼ����), CDate(str��������))
    If Not rsAccount.EOF Then
        curͳ��֧���ܶ� = Nvl(rsAccount!ͳ��֧�����, 0)
        curͳ���Ը��ܶ� = Nvl(rsAccount!ͳ���Ը����, 0)
    End If
    '��ȡָ�����ڷ�Χ�ڵĿ�ʼ������������
    If Not FUNC_������(str��ʼ����, str��������, str��ʼ������, str����������, _
    str��ʼ������, str����������, True) Then Exit Sub
    
    If Not ���� Then Exit Sub
    '��ȡҽ�����ķ��ص�ͳ��֧����ͳ���Ը��ܶ�
    gstrPara_���� = GetParaCode(����������, gComInfo_����.����������) & _
        GetParaCode(������_��ʼ, str��ʼ������) & GetParaCode(������_��ֹ, str����������) & _
        GetParaCode(������_��ʼ, str��ʼ������) & GetParaCode(������_��ֹ, str����������)
    Call ���ýӿ�_����("checkexpensereckoninginfo")
    Call ҽ����ֹ_����
    
    If Not (Format(curͳ��֧���ܶ�, "#####0.00;-#####0.00;0;") = Format(gComInfo_����.�˶�_ͳ��֧�����, "#####0.00;-#####0.00;0;") _
    And Format(curͳ���Ը��ܶ�, "#####0.00;-#####0.00;0;") = Format(gComInfo_����.�˶�_ͳ���Ը����, "#####0.00;-#####0.00;0;")) Then
        MsgBox "�����أ�ͳ��֧���ܶ" & curͳ��֧���ܶ� & String(4, " ") & "��ҽ����ͳ��֧���ܶ" & gComInfo_����.�˶�_ͳ��֧����� & vbCrLf & _
               "�����أ�ͳ���Ը��ܶ" & curͳ���Ը��ܶ� & String(4, " ") & "��ҽ����ͳ���Ը��ܶ" & gComInfo_����.�˶�_ͳ���Ը����
    Else
        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function FUNC_������(ByVal str��ʼ���� As String, ByVal str�������� As String, _
    str��ʼ������ As String, str���������� As String, _
    str��ʼ������ As String, str���������� As String, ByVal bln���� As Boolean) As Boolean
    Dim rs˳��� As New ADODB.Recordset
    '˳��ű�������ţ���ע�б��浱�εĽ�����
    gstrSQL = "Select Min(A.֧��˳���) ��ʼ������,Max(A.֧��˳���) ���������� "
    If bln���� Then gstrSQL = gstrSQL & ",Min(A.��ע) ��ʼ������,Max(A.��ע) ����������"
    gstrSQL = gstrSQL & _
             " From ���ս����¼ A" & _
             " Where A.����=[1] And ����ʱ�� between [2] And [3]"
    Set rs˳��� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������", type_����, CDate(str��ʼ����), CDate(str��������))
    
    If IsNull(rs˳���!��ʼ������) Then Exit Function
    
    str��ʼ������ = rs˳���!��ʼ������
    str���������� = rs˳���!����������
    If bln���� Then
        str��ʼ������ = rs˳���!��ʼ������
        str���������� = rs˳���!����������
    End If
    FUNC_������ = True
End Function

Public Sub �����ϴ�������ϸ()
    On Error GoTo errHand
    Dim bln���ֲ� As Boolean, int�ϴ� As Integer, blnError As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If Not ���� Then Exit Sub
    If Not ���ýӿ�_����("getsysdate") Then Exit Sub
    
    gstrSQL = "Select A.ID,A.����ID,A.NO,A.���,A.��¼����,A.��¼״̬,A.�Ǽ�ʱ��,A.������ as ҽ��," & _
            "   A.����*A.���� as ����,Round(A.���ʽ��/(A.����*A.����),2) as ʵ�ʼ۸�,A.���ʽ��," & _
            "   A.�շ����,B.���� as ��Ŀ����,B.���� as ��Ŀ����,D.��Ŀ���� ҽ������,Nvl(H.���,0) �ز���," & _
            "   C.���� ��������,E.���� �ܵ�����,A.ժҪ ������ˮ��,F.֧��˳��� ������,F.��ע ������,G.ҽ���� ���˱��" & _
            " From (Select * From ������ü�¼ Where ��¼����=1 And Nvl(�Ƿ��ϴ�,0)=0 And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0) A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,���ű� E,���ս����¼ F,�����ʻ� G,(Select * From ���ղ��� Where ����=[1]) H " & _
            " Where A.�շ�ϸĿID=B.ID And A.��������ID=C.ID(+) And A.ִ�в���ID=E.ID(+) And A.�շ�ϸĿID=D.�շ�ϸĿID And D.����=[1]" & _
            " And A.����ID=F.��¼ID And F.����=1 And A.����ID=G.����ID And G.����=D.���� And G.����ID=H.ID(+)" & _
            " Order by A.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ν��ʷ�����ϸ", type_����)
    With rsTemp
        Do While Not .EOF
            '֧�������ҽ�������Ƿ������ֲ��й�
            bln���ֲ� = (!�ز��� <> 0)
            gComInfo_����.֧����� = 201 'IIf(bln���ֲ�, "0205", "0201")
            gstrPara_���� = GetParaCode(���˱��, !���˱��) & GetParaCode(����������, gComInfo_����.����������) & _
                GetParaCode(������, !������) & GetParaCode(������ˮ��, !������ˮ��) & _
                GetParaCode(������, !������) & GetParaCode(֧�����, gComInfo_����.֧�����) & _
                GetParaCode(ҽ������, !ҽ������) & GetParaCode(��Ŀ����, !��Ŀ����) & GetParaCode(��Ŀ����, !��Ŀ����) & _
                GetParaCode(����, !����) & GetParaCode(�����ܶ�, !���ʽ��) & GetParaCode(��������, Nvl(!��������, "")) & _
                GetParaCode(����ҽ��, Nvl(!ҽ��, "")) & GetParaCode(�ܵ�����, Nvl(!�ܵ�����, "")) & GetParaCode(�ܵ�ҽ��, "") & _
                GetParaCode(����ʱ��, gComInfo_����.ϵͳʱ��)
            
            If ���ýӿ�_����("recipeinfotran") Then
                int�ϴ� = 1
            Else
                int�ϴ� = 0
                blnError = True
            End If
            
            'Ϊ���˷��ü�¼���ϱ�ǣ��Ա���ʱ�ϴ�
            'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & rsTemp("ID") & ",NULL,NULL,NULL,NULL," & int�ϴ� & ",'" & !������ˮ�� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
            .MoveNext
        Loop
    End With

    Call ҽ����ֹ_����
    If blnError Then
        MsgBox "���ַ�����ϸδ��ȷ�ϴ����뵽�����ʻ������������ϴ���", vbInformation, gstrSysName
    Else
        MsgBox "��ϸ�ϴ��ɹ���", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Function �����Ǽ�_����(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '��д�뵥��ͷ����д�뵥����
    '��¼״̬��1-����;����Ϊɾ���������ô�����ֻ�����ŵ���ɾ�����ٲ����µ���
    On Error GoTo errHand
    �����Ǽ�_���� = False

    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        gstrSQL = " Select A.����ID,A.NO,A.���,A.��¼����,A.��¼״̬,to_char(A.�Ǽ�ʱ��,'yyyy-MM-dd hh24:mi:ss') �Ǽ�ʱ��," & _
                  " A.������ ҽ��,B.���� ��������,A.�շ�ϸĿID,C.��Ŀ���� ҽ����Ŀ���� ,A.ʵ�ս�� ���,A.����*Nvl(A.����,1) ����,Nvl(A.�Ƿ��ϴ�,0) �Ƿ��ϴ�" & _
                  " From סԺ���ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=[1]) C " & _
                  " Where A.��¼����=[2] And A.��¼״̬=[3] And A.NO=[4]" & _
                  " And A.��������ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0" & _
                  " Order by A.����ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����Ǽ�", type_����, lng��¼����, lng��¼״̬, str���ݺ�)
        If .RecordCount = 0 Then
            MsgBox "δ�ҵ�������¼����ҽ����������������ʧ�ܣ�[�����Ǽ�]", vbInformation, gstrSysName
            Exit Function
        End If
    End With

    If Not ���� Then Exit Function
    If Not �ϴ�����_����(rsTemp) Then Exit Function
    Call ҽ����ֹ_����

    �����Ǽ�_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function �ϴ�����_����(ByVal rsExse As ADODB.Recordset) As Boolean
    Dim lng����ID As Long
    Dim curTotal As Currency
    Dim blnUpload As Boolean, blnInsure As Boolean
    Dim rsTemp As New ADODB.Recordset, rsInsure As New ADODB.Recordset
    
    If Not ���ýӿ�_����("getsysdate") Then Exit Function
    
    gComInfo_����.������ = 0               '�ϴ�������ϸʱ�������ű���Ҫ��Ϊ��
    gComInfo_����.֧����� = "0301"
    gComInfo_����.������ˮ�� = GetSequence(gint������ˮ��)
    
    With rsExse
        Do While Not .EOF
            If IsNull(!ҽ����Ŀ����) Then
                MsgBox "����Ŀδ����ҽ�����룬���ܽ��㡣", vbInformation, gstrSysName
                Exit Function
            End If
            .MoveNext
        Loop
        .MoveFirst
        
        '�ϴ�������ϸ
        Do While Not .EOF
            If lng����ID <> !����ID Then
                '��鱾���Ƿ���ҽ�������Ժ
                gstrSQL = "Select Count(*) Records From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.����ID=[1] And A.��ҳID=B.סԺ���� And A.����=[2]"
                Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ҽ������", CLng(!����ID), type_����)
                blnInsure = (rsInsure!Records = 1)
                If blnInsure Then
                    blnInsure = ��ȡ���������Ϣ(!����ID)
                    If blnInsure Then lng����ID = !����ID
                End If
            End If
            
            If blnInsure Then
                blnUpload = False
                If Not IsNull(!�Ƿ��ϴ�) Then
                    blnUpload = (!�Ƿ��ϴ� = 0)
                End If
                
                If blnUpload Then
                    
                    'ȡ���շ�ϸĿ�ı���������
                    gstrSQL = "Select ���� ��Ŀ���� ,���� ��Ŀ���� From �շ�ϸĿ Where ID = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���շ�ϸĿ�ı���������", CLng(!�շ�ϸĿID))
                    
                    gstrPara_���� = GetParaCode(���˱��, gComInfo_����.���˱��) & GetParaCode(����������, gComInfo_����.����������) & _
                        GetParaCode(������, gComInfo_����.������) & GetParaCode(������ˮ��, gComInfo_����.������ˮ�� & .AbsolutePosition) & _
                        GetParaCode(������, gComInfo_����.������) & GetParaCode(֧�����, gComInfo_����.֧�����) & _
                        GetParaCode(ҽ������, !ҽ����Ŀ����) & GetParaCode(��Ŀ����, rsTemp!��Ŀ����) & GetParaCode(��Ŀ����, rsTemp!��Ŀ����) & _
                        GetParaCode(����, Abs(!����)) & GetParaCode(�����ܶ�, !���) & GetParaCode(��������, Nvl(!��������, "")) & _
                        GetParaCode(����ҽ��, Nvl(!ҽ��, "")) & GetParaCode(�ܵ�����, "") & GetParaCode(�ܵ�ҽ��, "") & _
                        GetParaCode(����ʱ��, Format(!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss"))
                    If ���ýӿ�_����("recipeinfotran") Then
                        '�����ϴ���־����Ϊ��ϸ������ȷ�ϴ��󣬲��ܱ�֤�������ȷ��
                        'ID_IN,ͳ����_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,�Ƿ��ϴ�_IN,ժҪ_IN
                        gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ")"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
                    Else
                        Exit Function
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    
    �ϴ�����_���� = True
End Function


