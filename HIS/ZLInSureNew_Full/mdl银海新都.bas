Attribute VB_Name = "mdl�����¶�"
Option Explicit
'API��������

'1�������ϴ�
Private Declare Function DataUnloading Lib "yhybReckoning.dll" Alias "_DataUnloading@12" _
        (ByVal str_UploadData As String, ByVal str_UploadLsh As String, ByVal str_Fzxbm As String) As String

'2���ʻ�֧��
Private Declare Function reckoning Lib "yhybReckoning.dll" Alias "_reckoning@64" (ByVal str���� As String, _
        ByVal strҽ���� As String, ByVal str������ As String, ByVal str���� As String, _
        ByVal str����˳��� As String, ByVal str֧����� As String, ByVal strҽԺ���� As String, _
        ByVal str��Ժ���� As String, ByVal dbl�ʻ�֧�� As String, ByVal dat֧��ʱ�� As String, _
        ByVal dbl�ܶ� As String, ByVal dblȫ�Է� As String, ByVal dbl�ҹ��Ը� As String, _
        ByVal dbl������ As String, ByVal str������ As String, ByVal str������ As String) As String

'3����ȡ��ǰҽԺ������Ϣ
Private Declare Function GetHospitalInfo Lib "yhybReckoning.dll" Alias "_GetHospitalInfo@0" () As String

'4��������ϸ�ָ�
'Private Declare Function DivideUp Lib "yhybDivideUp.dll" Alias "_DivideUp@24" _
        (ByVal str�����ı�� As String, ByVal strҽ����Ŀ���� As String, ByVal str֧����� As String, _
        ByVal strҽ����Ա��� As String, ByVal dbl�ָ��� As Double) As String
Private Declare Function DivideUp Lib "yhybReckoning.dll" Alias "_DivideUp@24" _
        (ByVal str�����ı�� As String, ByVal strҽ����Ŀ���� As String, ByVal str֧����� As String, _
        ByVal strҽ����Ա��� As String, ByVal dbl�ָ��� As Double) As String

'5�������֧�����
Private Declare Function GetPayCount Lib "yhybReckoning.dll" Alias "_GetPayCount@48" _
        (ByVal str�����ı�� As String, ByVal str֧����� As String, _
        ByVal dbl�����Ը� As Double, ByVal dblȫ�Է� As Double, ByVal dbl�ҹ��Է� As Double, _
        ByVal dbl���� As Double, ByVal dbl�ʻ���� As Double) As String

'6�����ý���
Private Declare Function CalculateFeeCD Lib "yhybBill.dll" Alias "_CalculateFeeCD@84" _
        (ByVal dbl�����ܶ� As Double, ByVal dbl���� As Double, ByVal dblͳ���޶� As Double, _
        ByVal dblͳ��֧���ۼ� As Double, ByVal intʵ������ As Integer, ByVal dbl�ѽ������� As Double, _
        ByVal dbl�ѽ���ҹ��Ը� As Double, ByVal dbl���������� As Double, ByVal dblȫ�Է� As Double, _
        ByVal dbl�ҹ��Է� As Double, ByVal dblͳ�ﱨ������ As Double) As String
'7��ҽ������Ŀ¼�ļ�
Private Declare Function MakeTxt Lib "yhybReckoning.dll" Alias "_MakeTxt@8" (ByVal str����Ŀ¼�ļ� As String, _
        ByVal str����Ŀ¼�ļ� As String) As String

'8������������
Private Declare Function GetKard Lib "yhybReckoning.dll" Alias "_GetKard@4" (ByVal str_UploadData As String) As String

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public mint���õ���_�¶� As Integer
Public mintIC�������� As Integer

Private mstrҽ���� As String
Private mstr���� As String
Private mlng����ID As Long
Private mstr����� As String
Private mstrInfo As String                      '������Ϣ�����ڲ�����־�ļ�
Private mstr������ˮ�� As String                '����סԺ�����������ҵ���������˳���δ���µ������ʻ��У��������סԺ��˳���
Private mcol����ϸ As New Collection

Public Function ҽ����ʼ��_�¶�() As Boolean
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
'    Dim rsTmp As New ADODB.Recordset
    
    '��ȡ��ǰ�ӿ����õ���
    mint���õ���_�¶� = 0
    '�������´���,���ܻ��ж������ʹ�ñ��ӿ�
    
    '�����ж�
'    If rsTmp.State = 1 Then rsTmp.Close
'    gstrSQL = "select count(*) as �к� from dual where sysdate>=TO_DATE('2006-09-20 00:00:00','YYYY-MM-DD HH24:MI:SS')"
'    Call OpenRecordset(rsTmp, "���������ж�")
'
'    rsTmp.Close
'    If rsTmp!�к� = 1 Then
'        MsgBox "���������ѵ�������ɶ�������˾��ϵ��", vbInformation, gstrSysName
'        ҽ����ʼ��_�¶� = false
'        Exit Function
'    End If


    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ�ӿ����õ���", TYPE_�¶�)
    Do Until rsTemp.EOF
       Select Case rsTemp!������
          Case "���õ���"
            mint���õ���_�¶� = Nvl(rsTemp!����ֵ, 0)
          Case "������"
            mintIC�������� = Nvl(rsTemp!����ֵ, 0)
        End Select
        rsTemp.MoveNext
    Loop
    
    ҽ����ʼ��_�¶� = True
End Function

Public Function ҽ������_�¶�() As Boolean
'���ܣ� �÷������ڹ����Ӧ�ò���������������ҽ�����ݷ����������Ӵ�
'���أ��ӿ����óɹ�������true�����򣬷���false
    Dim strConn As String
    
    ҽ������_�¶� = frmSet�¶�����.ShowSet
End Function

Public Function ��ݱ�ʶ_�¶�(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ��ջ���Ϣ��
'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
'      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
'      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    Dim str���� As String, strҽ���� As String, str���� As String
    Dim STR���� As String, str�Ա� As String, str���֤���� As String, lng���� As Long
    Dim str�������� As String, str��Ա��� As String, str��λ���� As String, str��λ���� As String
    Dim strIdentify As String, str���� As String, str����� As String
    Dim datCurr As Date, strҽԺ���� As String
    Dim strReturn As String, str��ˮ�� As String, strסԺ˳��� As String, str���ı�� As String, StrInput As String, arrOutput As Variant
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency, cur�����ʻ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency, cur�������� As Currency, cur�����ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, bln��ȡ�ʻ������Ϣ As Boolean, curͳ���޶� As Currency
    bln��ȡ�ʻ������Ϣ = False
    
    '��ʼ��һЩ����
    mlng����ID = 0
    mstr����� = ""
    mstrҽ���� = ""
    mstr���� = ""
    
    '��ò���ҽ���š������ı�ŵ���Ϣ
    If frmIdentify�ɶ�����.GetIdentify(TYPE_�¶�, str����, strҽ����, str���ı��, str����) = False Then Exit Function
    
    '���ò����Ƿ���ҽ���������סԺ
    Dim rsTemp As New ADODB.Recordset
    '���ò����Ƿ���Ժ,����IC����������ҽ�����뷵�ص�ҽ���Ų�һ��,����ʹ�ÿ��Ž����ж�
    gstrSQL = "select nvl(��ǰ״̬,0) as ��ǰ״̬,˳��� from �����ʻ� where ����=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, str����, TYPE_�¶�)
    
    If rsTemp.EOF = False Then
        If rsTemp("��ǰ״̬") = 1 Then
            '������Ժ�ڼ䷢������ҵ��
            strסԺ˳��� = Nvl(rsTemp!˳���)
'            If mint���õ���_�¶� = 1 Then
'                MsgBox "�ò�������ҽ�������Ժ�������ٽ��������֤��", vbInformation, gstrSysName
'                Exit Function
'            End If
        End If
    End If
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '���������֤
    If Get��ˮ��("A", strҽԺ����, str��ˮ��) = False Then Exit Function
    '����|���˱���|�����ı��|����|��ȡ������ˮ��#
    StrInput = str���� & "|" & strҽ���� & "|" & str���ı�� & "|" & str���� & "|" & IIf(bytType = 1, "31", "11") & "#"
    Call WriteLog("DataUnloading(" & StrInput & "," & str��ˮ�� & "," & str���ı�� & ")")
    strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '�жϣ��������Ϊ111111��˵���ǳ�ʼ���룬����Ҫ���û��޸ģ����˳����ν���
'    If mint���õ���_�¶� = 1 Then
'        If str���� = "111111" Then
'            MsgBox "������Ϊ�籣�ֳ�ʼ���룬�������������룡", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    
    'ȡ�÷���ֵ
    str���� = arrOutput(1)
    strҽ���� = arrOutput(3)
    
    STR���� = arrOutput(4)
    str�Ա� = IIf(arrOutput(5) = "2", "Ů", "��")
    str���֤���� = arrOutput(6)
    str�������� = arrOutput(7)
    If IsDate(str��������) = False Then
        str�������� = Get��������(str���֤����, 0)
    End If
    If IsDate(str��������) Then
        lng���� = DateDiff("yyyy", CDate(str��������), zlDatabase.Currentdate)
        str�������� = Format(CDate(str��������), "yyyy-MM-dd")
    Else
        str�������� = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End If
    
    str��Ա��� = arrOutput(8)
    str��λ���� = arrOutput(9)
    str��λ���� = arrOutput(10)
    '������Ժ�ڼ䷢������ҵ����ˣ��ڽ�������ҵ��ʱ�����סԺ˳��Ų�Ϊ�գ�˵����Ժ��������˳���
    str��ˮ�� = arrOutput(12)
    mstr������ˮ�� = arrOutput(12)
    If strסԺ˳��� <> "" Then str��ˮ�� = strסԺ˳���
    
    '����;ҽ����;����;����;�Ա�;��������;���֤;������λ
    'ҽ���ŵ�һλΪ������
    '������(2006-3-27):Ϊ��ֹ����й©,���뱣�����Ϊԭ����*3
    strIdentify = str���� & ";" & strҽ���� & ";" & str���� * 3 & ";" & STR���� & ";" & str�Ա� & ";" & str�������� & ";" & str���֤���� & ";" & str��λ���� & "(" & str��λ���� & ")"
    strIdentify = Replace(strIdentify, " ", "")
    cur�����ʻ� = arrOutput(11)
    
    str���� = ";"                                       '8.���Ĵ���
    str���� = str���� & ";" & str��ˮ��                 '9.˳���
    str���� = str���� & ";" & str��Ա���               '10��Ա���
    str���� = str���� & ";" & arrOutput(11)             '11�ʻ����
    str���� = str���� & ";" & IIf(strסԺ˳��� <> "", "1", "0")                       '12��ǰ״̬
    str���� = str���� & ";"                             '13����ID
    str���� = str���� & ";" & IIf(Left(str��Ա���, 1) = "��", 2, 1)     '14��ְ(1,2)
    str���� = str���� & ";" & str���ı��               '15����֤�� ����ҽ�����ڱ���ҽ�������ı��루���⽨��ҽ�����ģ�
    str���� = str���� & ";" & lng����                   '16�����
    str���� = str���� & ";"                             '17�Ҷȼ�
    str���� = str���� & ";" & cur�����ʻ�             '18�ʻ������ۼ�
    str���� = str���� & ";0"                            '19�ʻ�֧���ۼ�
    str���� = str���� & ";"                             '20����ͳ���ۼ�
    str���� = str���� & ";"                             '21ͳ�ﱨ���ۼ�
    str���� = str���� & ";"                             '22סԺ�����ۼ�
    str���� = str���� & ";"                             '23�������� (1����������)
    
        '�������˵�����Ϣ�������ʽ��
        '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
        '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
        '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
    
    '������(2006-01-13):ԭ�����λ�ô���
    lng����ID = BuildPatiInfo(bytType, strIdentify & str����, lng����ID, TYPE_�¶�)
    
    gstrSQL = "Select * From �����ʻ� Where ҽ����=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", strҽ����, TYPE_�¶�)
    If Not rsTemp.EOF Then
        lng����ID = rsTemp!����ID
    End If
    datCurr = zlDatabase.Currentdate
    If lng����ID <> 0 Then          '��������Ѵ��ڣ����ȡ�ʻ������Ϣ
        '�ʻ������Ϣ
        Call Get�ʻ���Ϣ(TYPE_�¶�, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�, cur��������, cur�����ۼ�, curͳ���޶�)
        bln��ȡ�ʻ������Ϣ = True
    End If
    

    
    If bln��ȡ�ʻ������Ϣ = True Then          '�����ȡ���ʻ������Ϣ��������д��
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�¶� & "," & Year(datCurr) & "," & _
            cur�����ʻ� & ",0," & _
            cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & "," & curͳ���޶� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    End If
    
    '���ظ�ʽ:�м���벡��ID
    If lng����ID <> 0 Then
        ��ݱ�ʶ_�¶� = strIdentify & ";" & lng����ID & str����
        
        mstrҽ���� = strҽ����
        mstr���� = str����
    Else
        mstr������ˮ�� = ""
    End If
    
    '������������ʾ���
    If gblnLED And bytType = 0 Then
        zl9LedVoice.Speak "#26" & arrOutput(11)
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_�¶�(strSelfNo As String, ByVal bytPlace As Byte) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: strSelfNO-���˸��˱��
'����: ���ظ����ʻ����Ľ��
    Dim rsTemp As New ADODB.Recordset, str���� As String, strҽ���� As String, str���� As String
    Dim strReturn As String, str��ˮ�� As String, str���ı�� As String, StrInput As String, arrOutput  As Variant
    Dim strҽԺ���� As String
    
    On Error GoTo errHandle
    
    
    If bytPlace = balanԤ�� Then
        '�ڲ�����Ժ���Ԥ��֮��ɱ仯�����Ե��²�������Ѿ���׼ȷ��
        '��ò���ҽ���š������ı�ŵ���Ϣ
        If frmIdentify�ɶ�����.GetIdentify(TYPE_�¶�, str����, strҽ����, str���ı��, str����) = False Then Exit Function
        
        If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
        
        '���������֤
        If Get��ˮ��("A", strҽԺ����, str��ˮ��) = False Then Exit Function
        StrInput = str���� & "|" & strҽ���� & "|" & str���ı�� & "|" & str���� & "|11#"
        Call WriteLog("DataUnloading(" & StrInput & "," & str��ˮ�� & "," & str���ı�� & ")")
        strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        mstrҽ���� = strҽ����
        mstr���� = str����
        �������_�¶� = Val(arrOutput(11))
    Else
        '�����ݿ��ж�ȡ����Ϊ�ղŲű����˵ģ�Ӧ����׼ȷ�ģ�
        gstrSQL = "Select �ʻ���� From �����ʻ� where ����=[1] and ����=0 and ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_�¶�, strSelfNo)
        
        If rsTemp.EOF = False Then
            �������_�¶� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �����������_�¶�(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim strҽ���� As String, StrInput As String, arrOutput  As Variant, strReturn As String
    Dim dbl�����ʻ� As Double
    Dim lng����ID As Long, datCurr As Date, lng��� As Long
    Dim str���ı�� As String, str����˳��� As String, str��Ա��� As String, strҽԺ���� As String
    Dim dbl�ܽ�� As Double, dblȫ�Է� As Double, dbl�ҹ��Ը� As Double, dbl���� As Double, dbl��� As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If rs��ϸ.RecordCount = 0 Then
        str���㷽ʽ = "�����ʻ�;0;0"
        �����������_�¶� = True
        Exit Function
    End If
    rs��ϸ.MoveFirst
    lng����ID = rs��ϸ("����ID")
    datCurr = zlDatabase.Currentdate
    
    '�ӱ����ʻ���õǼ���Ϣ
    gstrSQL = "select ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���  " & _
              "from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", lng����ID, TYPE_�¶�)
    'str����˳��� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
    str����˳��� = mstr������ˮ��
    str���ı�� = IIf(IsNull(rsTemp("���ı��")), "", rsTemp("���ı��"))
    str��Ա��� = IIf(IsNull(rsTemp("��Ա���")), "", rsTemp("��Ա���"))
    strҽ���� = rsTemp("ҽ����")
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '��������Ѿ�����ķ�����ϸ
    Set mcol����ϸ = Nothing
    
    'Ȼ����봦����ϸ
    Do Until rs��ϸ.EOF
        '�õ�������ϸ
        gstrSQL = "select A.����,A.����,A.���,A.���,A.���㵥λ,B.��Ŀ����,B.��ע,C.��� as ���� " & _
                 " from �շ�ϸĿ A,����֧����Ŀ B,�շ���� C " & _
                 " where A.���=C.���� and  A.ID=[1] and A.ID=B.�շ�ϸĿID and B.����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", CLng(rs��ϸ("�շ�ϸĿID")), TYPE_�¶�)
        
        '���з��÷ָ�
        strReturn = DivideUp(str���ı��, ToVarchar(rsTemp("��Ŀ����"), 12), "11", str��Ա���, Val(rs��ϸ("����")))
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        '�ڶ�������˳�������Ϊ���㵥��
        StrInput = str����˳��� & "|" & str����˳���
        StrInput = StrInput & "|" & str����˳��� & "_" & lng���      '���
        StrInput = StrInput & "|" & strҽ���� & "|" & str���ı�� & "|" & strҽԺ���� & "|000"
        StrInput = StrInput & "|" & ToVarchar(rsTemp("��Ŀ����"), 12)  'ҽ����ˮ��
        StrInput = StrInput & "|" & ToVarchar(rsTemp("����"), 10)      '�շѴ�������
        StrInput = StrInput & "|" & Format(rs��ϸ("����"), "0.00")
        StrInput = StrInput & "|" & Format(rs��ϸ("����"), "0.00")
        StrInput = StrInput & "|" & Format(rs��ϸ("ʵ�ս��"), "0.00")
        StrInput = StrInput & "|" & arrOutput(4)                       '�Ը�����
        StrInput = StrInput & "|" & Format(Val(arrOutput(1)) * rs��ϸ("����"), "#0.00") 'ȫ�ԷѲ���
        StrInput = StrInput & "|" & Format(Val(arrOutput(2)) * rs��ϸ("����"), "#0.00") '�ҹ��ԷѲ���
        StrInput = StrInput & "|" & Format(Val(arrOutput(3)) * rs��ϸ("����"), "#0.00") '����������
        StrInput = StrInput & "||11"                                   '�����־��֧�����
        StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 56)       '������������
        StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 20)       '��������ҽ��
        StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 56)       '�ܵ���������
        StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 20)       '�ܵ�����ҽ��
        StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 20)        '������
        StrInput = StrInput & "|" & Format(datCurr + lng��� / 24 / 3600, "yyyy-MM-dd HH:mm:ss") '����ʱ��
        StrInput = StrInput & "|" & ToVarchar(rsTemp("����"), 200)       '�շ���Ŀ
        StrInput = StrInput & "|" & ToVarchar(rsTemp("���"), 200)       '���
        StrInput = StrInput & "|"                                        '����
        StrInput = StrInput & "|" & ToVarchar(rsTemp("���㵥λ"), 30)    '��λ
        StrInput = StrInput & "|||"                                      'Ӣ��������ѧ��
        StrInput = StrInput & lng��� & "#"                             '���
        Call WriteLog(StrInput)
        mcol����ϸ.Add StrInput  '���Ƚ���ϸ���棬������ʱ���ϴ�
        
        lng��� = lng��� + 1
        dbl�ܽ�� = dbl�ܽ�� + Val(rs��ϸ("ʵ�ս��"))
        dblȫ�Է� = dblȫ�Է� + Val(arrOutput(1)) * rs��ϸ("����")
        dbl�ҹ��Ը� = dbl�ҹ��Ը� + Val(arrOutput(2)) * rs��ϸ("����")
        dbl���� = dbl���� + Val(arrOutput(3)) * rs��ϸ("����")    'Ŀǰʹ�����������֡�����
        
        rs��ϸ.MoveNext
    Loop
    
    '�õ��������
    dbl��� = �������_�¶�(strҽ����, balan����)
    With g��������
        .�������ý�� = dbl�ܽ��
        .ȫ�Էѽ�� = dblȫ�Է�
        .�����Ը���� = dbl�ҹ��Ը�
        .����ͳ���� = dbl����
        .֧��˳��� = str����˳���
    End With
    '����Ԥ����
    strReturn = GetPayCount(str���ı��, "11", dbl�ܽ��, dblȫ�Է�, dbl�ҹ��Ը�, dbl����, dbl���)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    dbl�����ʻ� = Val(arrOutput(1))                 'ȡ�ӿ������ʻ�֧���Ľ��
    '������(2005-12-28):�¶��Ͷ����߶�����ȫ��ʹ�ø����ʻ�����,�������ж�
    dbl�����ʻ� = IIf(dbl��� < dbl�ܽ��, dbl���, dbl�ܽ��)
    str���㷽ʽ = "�����ʻ�;" & dbl�����ʻ� & ";1"   '�����޸ĸ����ʻ�
    �����������_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function �������_�¶�(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur֧�����   �Ӹ����ʻ���֧���Ľ��
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim strҽ���� As String, StrInput As String, arrOutput  As Variant, strReturn As String
    Dim lng����ID  As Long, rs��ϸ As New ADODB.Recordset
    Dim datCurr As Date, var��ϸ As Variant, rsTemp As New ADODB.Recordset
    Dim str���ı�� As String, str����˳��� As String, strҽԺ���� As String, str���� As String, str��ˮ�� As String
    Dim dbl��� As Double
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From ������ü�¼ Where ����ID=[1]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", lng����ID)
    If rs��ϸ.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "û����д�շѼ�¼"
        Exit Function
    End If
    lng����ID = rs��ϸ("����ID")
    datCurr = rs��ϸ("�Ǽ�ʱ��")
    
    If mstrҽ���� <> strSelfNo Then
        Err.Raise 9000, gstrSysName, "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣"
        Exit Function
    End If
    
    '����ʻ������Ϣ
    gstrSQL = "select ����,ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���  " & _
              "from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "����Ԥ��", lng����ID, TYPE_�¶�)
    str����˳��� = mstr������ˮ��
    str���ı�� = IIf(IsNull(rs��ϸ("���ı��")), "", rs��ϸ("���ı��"))
    str���� = IIf(IsNull(rs��ϸ("����")), "", rs��ϸ("����")) '���뿨��û�п���
    strҽ���� = rs��ϸ("ҽ����")
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '�ϴ�������ϸ��ͳһ��һ����ˮ�ţ�������
    If Get��ˮ��("G", strҽԺ����, str��ˮ��) = False Then Exit Function
    For Each var��ϸ In mcol����ϸ
        Call WriteLog("�ϴ�:" & var��ϸ)
        strReturn = DataUnloading(var��ϸ, str��ˮ��, str���ı��)
        Call WriteLog("����:" & strReturn)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    Next
    
    '���ý���
    With g��������
    Call WriteLog("����(" & str���� & "," & strҽ���� & "," & str���ı�� & "," & mstr���� & "," & str����˳��� & "," & "11" & "," & strҽԺ���� & "," & "000" & "," & CStr(cur�����ʻ�) & "," & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "," & _
               CStr(.�������ý��) & "," & CStr(.ȫ�Էѽ��) & "," & CStr(.�����Ը����) & "," & CStr(.����ͳ����) & "," & ToVarchar(UserInfo.����, 20) & "," & ToVarchar(.֧��˳���, 20) & ")")
    strReturn = reckoning(str����, strҽ����, str���ı��, mstr����, str����˳���, "11", strҽԺ����, "000", Format(cur�����ʻ�, "0.##"), Format(datCurr, "yyyy-MM-dd HH:mm:ss"), _
               Format(.�������ý��, "0.##"), Format(.ȫ�Էѽ��, "0.##"), Format(.�����Ը����, "0.##"), Format(.����ͳ����, "0.##"), ToVarchar(UserInfo.����, 20), ToVarchar(.֧��˳���, 20))
    Call WriteLog("����:" & strReturn)
    End With
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    
    '��������¼
    '---------------------------------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
    Dim cur�����ۼ� As Currency, cur�������� As Currency, curͳ���޶� As Currency
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_�¶�, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�, cur��������, cur�����ۼ�, curͳ���޶�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�¶� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur�������� & "," & cur�����ۼ� & "," & curͳ���޶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�¶� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & g��������.�������ý�� & ",0,0," & _
        0 & "," & 0 & ",0,0," & cur�����ʻ� & ",'')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------
    '������(2006-4-25)����������ʾ
    dbl��� = �������_�¶�(strҽ����, balan����)
    If gblnLED Then
       zl9LedVoice.Speak "#25 " & g��������.�������ý��
       If cur�����ʻ� < g��������.�������ý�� Then
          zl9LedVoice.Speak "#27 " & g��������.�������ý�� - cur�����ʻ�
       Else
          zl9LedVoice.Speak "#26 " & dbl���
       End If
    End If
    
    �������_�¶� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ����������_�¶�(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
'���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
'      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��

    ����������_�¶� = True
End Function

Public Function �����ʻ�תԤ��_�¶�(lngԤ��ID As Long, cur�����ʻ� As Currency, strSelfNo As String, str˳��� As String, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
'������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
    Dim strҽ���� As String, StrInput As String, arrOutput  As Variant, strReturn As String
    Dim datCurr As Date, var��ϸ As Variant, rs��ϸ As New ADODB.Recordset
    Dim str���ı�� As String, str����˳��� As String, strҽԺ���� As String, str���� As String, str��ˮ�� As String
    
    On Error GoTo errHandle
    
    datCurr = zlDatabase.Currentdate
    
    If mstrҽ���� <> strSelfNo Then
        MsgBox "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    
    '����ʻ������Ϣ
    gstrSQL = "select ����,ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���  " & _
              "from �����ʻ� where ����ID=[1] and ����=[2]"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��Ԥ����", lng����ID, TYPE_�¶�)
    str����˳��� = IIf(IsNull(rs��ϸ("�������")), "", rs��ϸ("�������"))
    str���ı�� = IIf(IsNull(rs��ϸ("���ı��")), "", rs��ϸ("���ı��"))
    str���� = IIf(IsNull(rs��ϸ("����")), "", rs��ϸ("����")) '���뿨û�п���
    strҽ���� = rs��ϸ("ҽ����")
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '�����жϽ���Ƿ����ʹ��
    strReturn = GetPayCount(str���ı��, "31", cur�����ʻ�, 0, 0, cur�����ʻ�, cur�����ʻ�)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    If Val(arrOutput(1)) < cur�����ʻ� Then
        MsgBox "�����ʻ���������֧��Ԥ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ý���
    Call WriteLog("reckoning(" & str���� & "," & strҽ���� & "," & str���ı�� & "," & mstr���� & "," & str����˳��� & "," & "31" & "," & strҽԺ���� & "," & "000" & "," & CStr(cur�����ʻ�) & "," & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "," & _
               CStr(cur�����ʻ�) & "," & CStr(0) & "," & CStr(0) & "," & CStr(cur�����ʻ�) & "," & ToVarchar(UserInfo.����, 20) & "," & ToVarchar(str����˳���, 20) & ")")
    strReturn = reckoning(str����, strҽ����, str���ı��, mstr����, str����˳���, "31", strҽԺ����, "000", CDbl(cur�����ʻ�), Format(datCurr, "yyyy-MM-dd HH:mm:ss"), _
               CDbl(cur�����ʻ�), CDbl(0), CDbl(0), CDbl(cur�����ʻ�), ToVarchar(UserInfo.����, 20), ToVarchar(str����˳���, 20))
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    
    '��������¼
    '---------------------------------------------------------------------------------------------
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
            
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_�¶�, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
                
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�¶� & "," & Year(datCurr) & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� + cur�����ʻ� & "," & _
        cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(3," & lngԤ��ID & "," & TYPE_�¶� & "," & lng����ID & "," & _
        Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
        curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & ",0,0,0," & cur�����ʻ� & ",0,0," & _
        0 & "," & 0 & ",0,0," & cur�����ʻ� & ",'')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------

    �����ʻ�תԤ��_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_�¶�(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String, Optional ByVal blnFirst As Boolean = True) As Boolean
'���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
    Dim StrInput As String, arrOutput  As Variant, arrTmp As Variant
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str��ˮ�� As String, strReturn As String
    Dim str���ı�� As String, str����˳��� As String, strҽԺ���� As String, str���� As String
    Dim str��Ժ��� As String, str��Ժ��� As String, strҽԺ��� As String, str��Ժ��� As String
    Dim intValue As Integer
    Dim dblͳ���޶� As Double, dblͳ���ۼ� As Double, dbl�������� As Double, dblסԺ���� As Double
    Dim str�����־ As String
    
    On Error GoTo errHandle
    
    '��ȡ���ղ���ֵ���Ծ���ҽ��������Ժʱ���Ƿ�ͬʱ�ϴ���Ժ��Ϣ
    intValue = 1
'    gstrSQL = "Select Nvl(����ֵ,0) Value From ���ղ��� Where ����=" & TYPE_�¶� & " And ������='�ϴ���Ժ��Ϣ'"
'    Call OpenRecordset(rsTemp, "��ȡ�ϴ���Ժ��Ϣ����ֵ")
'
'    If Not rsTemp.EOF Then
'        intValue = rsTemp!Value
'    End If
    
    '���ҽ����
    gstrSQL = "select ҽ����,����,˳��� as �������,����֤�� as ���ı�� from �����ʻ� where ����=[1] and ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", TYPE_�¶�, lng����ID)
    
    str���� = IIf(IsNull(rsTemp("����")), "", rsTemp("����")) '��������뿨,���ž�Ϊ��
    strҽ���� = rsTemp("ҽ����")
    str����˳��� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
    str���ı�� = IIf(IsNull(rsTemp("���ı��")), "", rsTemp("���ı��"))
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.�������,A.������Ϣ from ������ A where A.����ID=[1] and A.��ҳID=[2]" & _
              " and A.������� in (1,3) and A.��ϴ���=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    Do Until rsTemp.EOF
        If rsTemp("�������") = 1 Then
            str��Ժ��� = ToVarchar(IIf(IsNull(rsTemp("������Ϣ")), "����", rsTemp("������Ϣ")), 128)
        Else
            str��Ժ��� = ToVarchar(IIf(IsNull(rsTemp("������Ϣ")), "����", rsTemp("������Ϣ")), 128)
        End If
        rsTemp.MoveNext
    Loop
    If str��Ժ��� = "" Then str��Ժ��� = "����" '��ϲ�����β���Ϊ��
    If str��Ժ��� = "" Then str��Ժ��� = "����" '��ϲ�����β���Ϊ��
    
    '���������Ժ��Ϣ
    datCurr = zlDatabase.Currentdate
    gstrSQL = " select A.��Ժ����,A.�Ǽ�ʱ��,B.���� ��Ժ���� " & _
              " from ������ҳ A,���ű� B " & _
              " Where A.��Ժ����ID=B.ID  And A.����ID = [1] And A.��ҳID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    str��Ժ��� = Year(rsTemp!��Ժ����)
    
    '��������2006��08��31�����ѽ�������Ҫ���������־
    str�����־ = "0"
    '�ѽ�������Ҫ���������־
    If mint���õ���_�¶� = 2 Then
       Dim str������Ϣ As String
       
       If blnFirst Then
          If MsgBox("�ò����Ƿ�Ϊ�������α����ˣ�", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbYes Then
             str�����־ = "PJ"
          End If
       Else
          Call Get������Ϣ(TYPE_�¶�, lng����ID, lng��ҳID, Year(datCurr), str������Ϣ)
          str�����־ = str������Ϣ
       End If
    End If
    
    '���ҽԺ���
    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Function
    
    '�������
    If Get��ˮ��("C", strҽԺ����, str��ˮ��) = False Then Exit Function
    StrInput = str���� & "|" & strҽ���� & "|" & str���ı�� & "|" & mstr���� & _
                "|" & str����˳��� & "|" & strҽԺ���� & _
                "|000|0|000|31|" & str�����־ & _
                "|" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss") & _
                "|" & ToVarchar(UserInfo.����, 20) & _
                "|" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "#"
    Call WriteLog("DataUnloadint(" & StrInput & "," & str��ˮ�� & "," & str���ı�� & ")")
    strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    dblͳ���޶� = Val(arrOutput(6))
    dblͳ���ۼ� = Val(arrOutput(8))
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer
            
    '�ʻ������Ϣ   ע����ֶ���ʵ���ô�֮��Ķ�Ӧ��ϵ
    '��������    ----   סԺ����
    '�����ۼ�    ----   ����ͳ��֧���ۼ�
    '����ͳ���޶�  ----   סԺͳ���޶�
    '���ͳ���޶�  ----   ʵ������
    '���ͳ���ۼ�  ----   ͳ�ﱨ������
    '������Ϣ      ----   ���ɶ��ѽ����������ⲡ�˱�־
    Call Get�ʻ���Ϣ(TYPE_�¶�, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�¶� & "," & str��Ժ��� & "," & _
        cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & _
        arrOutput(5) & "," & arrOutput(8) & "," & arrOutput(6) & "," & arrOutput(3) & "," & arrOutput(11) & ",'" & str�����־ & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        
    '������Ժ�ӿ�
    If blnFirst Then

        If Val(arrOutput(6)) = 0 Then
            MsgBox "����ͳ���޶�Ϊ�㣬����������ҽ�������Ժ���밴��ͨ���˰�����Ժ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        '�ϴ���Ժ�Ǽ�
        If intValue = 1 Then
            If Get��ˮ��("E", strҽԺ����, str��ˮ��) = False Then Exit Function
            StrInput = str����˳��� & "|" & strҽ���� & "|" & strҽԺ���� & "|000|" & strҽԺ��� & "|31|0"
            StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 20)    '��Ժ������
            StrInput = StrInput & "|" & ToVarchar(rsTemp("��Ժ����"), 20)  '��Ժ����
            StrInput = StrInput & "|" & str��Ժ���
            StrInput = StrInput & "|" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss")
            StrInput = StrInput & "|" & Format(rsTemp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss") & "|||��Ժ�Ǽ�|" & Format("2000-01-01", "yyyy-MM-dd HH:mm:ss") & "|" & Format("2000-01-01", "yyyy-MM-dd HH:mm:ss") & "|9#"
            Call WriteLog("�ϴ���Ժ�Ǽ�(" & StrInput & "," & str��ˮ�� & "," & str���ı�� & ")")
            strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
            Call WriteLog("����:" & strReturn)
            If JudgeReturn(strReturn, arrTmp) = False Then Exit Function
        End If

        '����״̬���޸�
        gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�¶� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        
        '������(2005-12-28):��ʾ���������Ϣ
        '������ͳ���޶���ͳ��֧���ۼ���ʾ����������Ա
        dblסԺ���� = Val(arrOutput(5))
        dbl�������� = Val(arrOutput(11))
        MsgBox "�òα����˵�סԺ�����Ϣ��" & vbCrLf & _
                   "    סԺ����  ����" & Format(dblסԺ����, "#0.00") & "Ԫ     " & vbCrLf & _
                   "    ����ͳ���޶��" & Format(dblͳ���޶�, "#0.00") & "Ԫ     " & vbCrLf & _
                   "    ͳ��֧���ۼƣ���" & Format(dblͳ���ۼ�, "#0.00") & "Ԫ     " & vbCrLf & _
                   "    ͳ�ﱨ��������  " & dbl�������� * 100 & "%", vbInformation, gstrSysName
    End If
    
    ��Ժ�Ǽ�_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_�¶�(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
'������lng����ID-����ID��lng��ҳID-��ҳID
'���أ����׳ɹ�����true�����򣬷���false
            'ȡ��Ժ�Ǽ���֤�����ص�˳���
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str��Ժ��� As String, str��Ժ��� As String
    Dim StrInput As String, arrOutput  As Variant, str��ˮ�� As String, strReturn As String
    Dim str���ı�� As String, str����˳��� As String, strҽԺ���� As String, strҽ���� As String
    Dim strҽԺ��� As String
    
    On Error GoTo errHandle
    
    '���ҽ����
    gstrSQL = "select ҽ����,����,˳��� as �������,����֤�� as ���ı�� from �����ʻ� where ����=[1] and ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", TYPE_�¶�, lng����ID)
    strҽ���� = rsTemp("ҽ����")
    str����˳��� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
    str���ı�� = IIf(IsNull(rsTemp("���ı��")), "", rsTemp("���ı��"))
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    '��ò��˳�Ժ���
    gstrSQL = "select A.�������,A.������Ϣ from ������ A where A.����ID=[1] and A.��ҳID=[2]" & _
              " and A.������� in (1,3) and A.��ϴ���=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    Do Until rsTemp.EOF
        If rsTemp("�������") = 1 Then
            str��Ժ��� = ToVarchar(IIf(IsNull(rsTemp("������Ϣ")), "����", rsTemp("������Ϣ")), 128)
        Else
            str��Ժ��� = ToVarchar(IIf(IsNull(rsTemp("������Ϣ")), "����", rsTemp("������Ϣ")), 128)
        End If
        rsTemp.MoveNext
    Loop
    If str��Ժ��� = "" Then str��Ժ��� = "����" '��ϲ�����β���Ϊ��
    If str��Ժ��� = "" Then str��Ժ��� = "����" '��ϲ�����β���Ϊ��
        
    '���������Ժ��Ϣ
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.����ҽʦ,A.סԺҽʦ,A.�Ǽ�ʱ��,A.��Ժ����,A.��Ժ����,A.��Ժ��ʽ,B.���� as ��Ժ����,C.���� as ��Ժ���� " & _
             " from ������ҳ A,���ű� B,���ű� C " & _
             " Where A.��Ժ����ID = B.ID And A.��Ժ����ID = C.ID And A.����ID = [1] And A.��ҳID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ժ�Ǽ�", lng����ID, lng��ҳID)
    
    '���ҽԺ���
    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Function

    '���ó�Ժ�ӿ�
    If Get��ˮ��("E", strҽԺ����, str��ˮ��) = False Then Exit Function
    StrInput = str����˳��� & "|" & strҽ���� & "|" & strҽԺ���� & "|000|" & strҽԺ��� & "|31|" & _
                IIf(Format(rsTemp("��Ժ����"), "yyyy") = Format(rsTemp("��Ժ����"), "yyyy"), "0", "1")
    StrInput = StrInput & "|" & ToVarchar(rsTemp("����ҽʦ"), 20)  '��Ժ������
    StrInput = StrInput & "|" & ToVarchar(rsTemp("��Ժ����"), 20)  '��Ժ����
    StrInput = StrInput & "|" & str��Ժ���
    StrInput = StrInput & "|" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss")
    StrInput = StrInput & "|" & Format(rsTemp("�Ǽ�ʱ��"), "yyyy-MM-dd HH:mm:ss")
    StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 20)       '��Ժ������
    StrInput = StrInput & "|" & ToVarchar(rsTemp("��Ժ����"), 20)  '��Ժ����
    StrInput = StrInput & "|" & str��Ժ���
    StrInput = StrInput & "|" & Format(rsTemp("��Ժ����"), "yyyy-MM-dd HH:mm:ss")
    StrInput = StrInput & "|" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") '��Ժ����ʱ��
    StrInput = StrInput & "|" & Switch(rsTemp("��Ժ��ʽ") = "����", 0, rsTemp("��Ժ��ʽ") = "����", 1, rsTemp("��Ժ��ʽ") = "תԺ", 2, True, 9) & "#"
    
    Call WriteLog("DataUnloadint(" & StrInput & "," & str��ˮ�� & "," & str���ı�� & ")")
    strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '����״̬���޸�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�¶� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    ��Ժ�Ǽ�_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ���ʴ���_�¶�(strNO As String, int���� As Integer, int״̬ As Integer, Optional lng����ID As Long) As Boolean
'���ܣ���סԺ���˵ļ��ʵ����ϴ���ҽ��ǰ�÷�����
'������lng����ID=�Ƿ�ֻ�ϴ�������ָ�����˵ķ���
    Dim StrInput As String, arrOutput   As Variant, strReturn As String
    Dim rsBill As New ADODB.Recordset, rsTemp As New ADODB.Recordset, rs�շ���� As New ADODB.Recordset
    Dim lng��ǰ���� As Long
    '���ô���ʹ�õı���
    Dim str���ı�� As String, str����˳��� As String, str��Ա��� As String, strҽԺ���� As String
    Dim str��ˮ�� As String, str�շ���� As String, strҽ���� As String, strժҪ As String
    Dim dbl���Ϸ�Χ As Double
    
    ���ʴ���_�¶� = True '���ȱ�֤�����ܵõ����档��ʹ�����ϴ��ܣ�Ҳ�������Ժ�����ϴ���
    On Error GoTo errHandle
    
    '�г������շ����
    gstrSQL = "Select ����,��� as ���� From �շ����"
    Set rs�շ���� = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
    
    '��ȡ������ϸ(ҽ����,˳���,�Ǽ�ʱ��,��Ŀ����,��Ŀ����,����,���,����,����,���,ҽ��,��������)
    '�����зǸ�ҽ���ķ��ò���,δ����ҽ������Ĳ���,��˳��ŵĲ���,Ӥ���Ѳ��ϴ�������������
    gstrSQL = _
        "Select Nvl(A.�۸񸸺�,���) as ���," & _
        " A.����ID,A.��ҳID,F.ҽ����,F.˳���,A.�Ǽ�ʱ��,D.��Ŀ����,B.���� as ��Ŀ����,A.�շ����, " & _
        " Decode(Instr(B.���,'��'),0,B.���,Substr(B.���,1,Instr(B.���,'��')-1)) as ���," & _
        " Decode(Instr(B.���,'��'),0,'',Substr(B.���,Instr(B.���,'��')+1)) as ����," & _
        " Avg(Nvl(A.����,1)*A.����) as ����,Sum(A.��׼����) as ����,Sum(A.ʵ�ս��) as ���," & _
        " A.������ as ҽ��,C.���� as ��������" & _
        " From סԺ���ü�¼ A,�շ�ϸĿ B,���ű� C,����֧����Ŀ D,������ҳ E,�����ʻ� F" & _
        " Where A.��¼״̬<>0 And Nvl(A.�Ƿ��ϴ�,0)=0 And A.�շ�ϸĿID=B.ID And A.��������ID=C.ID And A.�շ�ϸĿID=D.�շ�ϸĿID" & _
        " And A.����ID=E.����ID And A.��ҳID=E.��ҳID And A.����ID=F.����ID" & _
        " And F.˳��� is Not NULL And Nvl(A.Ӥ����,0)=0" & _
        " And D.����=[1] And E.����=[1] And F.����=[1]" & _
        " And A.NO=[2] And A.��¼����=[3] And A.��¼״̬=[4]" & _
        IIf(lng����ID = 0, "", " And A.����ID=[5]") & _
        " Group by Nvl(A.�۸񸸺�,���),A.����ID,A.��ҳID,F.ҽ����,F.˳���," & _
        " A.�Ǽ�ʱ��,D.��Ŀ����,B.����,A.�շ����,B.���,A.������,C.����" & _
        " Order by ����ID,���"
    Set rsBill = zlDatabase.OpenSQLRecord(gstrSQL, "���ʴ���", TYPE_�¶�, strNO, int����, int״̬, lng����ID)
    
    Do Until rsBill.EOF
        '���ʵ����ж������,Ҫ�ֱ���
        If lng��ǰ���� <> rsBill("����ID") Then
            '�Ըò�������Ӧ�ĳ�ʼ������-------------------------------------------------
            lng��ǰ���� = rsBill("����ID")
            
            '�õ���Ժ������Ϣ���Ѿ�������֤�ģ�
            gstrSQL = "Select ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���  " & _
                      "       ,NVL(A.��������,0) as סԺ����,NVL(A.�����ۼ�,0) as ����ͳ��֧���ۼ�" & _
                      "       ,NVL(A.����ͳ���޶�,0) as סԺͳ���޶�,NVL(A.���ͳ���޶�,0) as ʵ������,NVL(A.���ͳ���ۼ�,0) as ͳ�ﱨ������" & _
                      "  From �ʻ������Ϣ A,������ҳ B,�����ʻ� C " & _
                      "  where B.����ID=[1] and B.��ҳID=[2] and A.����ID=B.����ID and A.����=[3] and A.���=to_char(B.��Ժ����,'yyyy')" & _
                      "     and C.����ID=A.����ID and C.����=A.����"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ʴ���", lng��ǰ����, CLng(rsBill("��ҳID")), TYPE_�¶�)
            str����˳��� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
            str���ı�� = IIf(IsNull(rsTemp("���ı��")), "", rsTemp("���ı��"))
            strҽ���� = IIf(IsNull(rsTemp("ҽ����")), "", rsTemp("ҽ����"))
            str��Ա��� = IIf(IsNull(rsTemp("��Ա���")), "", rsTemp("��Ա���"))
            
            If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
            If Get��ˮ��("G", strҽԺ����, str��ˮ��) = False Then Exit Function
        End If
            
        '���з��÷ָ�
        Call WriteLog("DivideUp(" & str���ı�� & "," & ToVarchar(rsBill!��Ŀ����, 12) & "," & "31" & "," & str��Ա��� & "," & Val(rsBill!����) & ")")
        strReturn = DivideUp(str���ı��, ToVarchar(rsBill("��Ŀ����"), 12), "31", str��Ա���, Val(rsBill("����")))
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        '������(2006-3-20):ժҪ�����ʽ ȫ�ԷѲ���|�ҹ��ԷѲ���|���Ϸ�Χ����
        strժҪ = "'" & Format(Val(arrOutput(1)) * rsBill("����"), "#0.00") & "|" & Format(Val(arrOutput(2)) * rsBill("����"), "#0.00") & "|" & Format(Val(arrOutput(3)) * rsBill("����"), "#0.00") & "'"
        dbl���Ϸ�Χ = Val(arrOutput(3)) * rsBill("����")
        
        rs�շ����.Filter = "���� = '" & rsBill("�շ����") & "'"
        If rs�շ����.EOF = False Then str�շ���� = rs�շ����("����")
        
        '�ڶ�������˳�������Ϊ���㵥��
        StrInput = str����˳��� & "|" & str����˳���
        StrInput = StrInput & "|" & strNO & "_" & rsBill("���") & "_" & int���� & "_" & int״̬  '���
        StrInput = StrInput & "|" & strҽ���� & "|" & str���ı�� & "|" & strҽԺ���� & "|000"
        StrInput = StrInput & "|" & ToVarchar(rsBill("��Ŀ����"), 12)  'ҽ����ˮ��
        StrInput = StrInput & "|" & ToVarchar(str�շ����, 10)      '�շѴ�������
        StrInput = StrInput & "|" & Format(rsBill("����"), "0.00")
        StrInput = StrInput & "|" & Format(rsBill("����"), "0.00")
        StrInput = StrInput & "|" & Format(rsBill("���"), "0.00")
        StrInput = StrInput & "|" & arrOutput(4)                       '�Ը�����
        StrInput = StrInput & "|" & Format(Val(arrOutput(1)) * rsBill("����"), "#0.00") 'ȫ�ԷѲ���
        StrInput = StrInput & "|" & Format(Val(arrOutput(2)) * rsBill("����"), "#0.00") '�ҹ��ԷѲ���
        StrInput = StrInput & "|" & Format(Val(arrOutput(3)) * rsBill("����"), "#0.00") '����������
        StrInput = StrInput & "||31"                                   '�����־��֧�����
        StrInput = StrInput & "|" & ToVarchar(rsBill("��������"), 56)  '������������
        StrInput = StrInput & "|" & ToVarchar(rsBill("ҽ��"), 20)      '��������ҽ��
        StrInput = StrInput & "|" & ToVarchar(rsBill("��������"), 56)  '�ܵ���������
        StrInput = StrInput & "|" & ToVarchar(rsBill("ҽ��"), 20)      '�ܵ�����ҽ��
        StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 20)        '������
        StrInput = StrInput & "|" & Format(rsBill("�Ǽ�ʱ��") + rsBill("���") / 24 / 3600, "yyyy-MM-dd HH:mm:ss") '����ʱ��
        StrInput = StrInput & "|" & ToVarchar(rsBill("��Ŀ����"), 200)       '�շ���Ŀ
        StrInput = StrInput & "|" & ToVarchar(rsBill("���"), 200)       '���
        StrInput = StrInput & "|" & ToVarchar(rsBill("����"), 200)       '����
        StrInput = StrInput & "|"                                        '��λ
        StrInput = StrInput & "||"                                      'Ӣ��������ѧ��
        'modify by ccy ,Ψһ
'        StrInput = StrInput & Format(rsBill("�Ǽ�ʱ��"), "yyyyMMddHHmmss") & rsBill("���") & "#"
        '������(2006-02-16):����п���żȻ�ظ�,�޸�Ϊ���·�ʽ
        StrInput = StrInput & strNO & "_" & rsBill("���") & "_" & int���� & "_" & int״̬ & "#"      '���
        Call WriteLog("DataUnloading(" & StrInput & "," & str��ˮ�� & "," & str���ı�� & ")")
        strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        
        gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & strNO & "'," & rsBill("���") & "," & int���� & "," & int״̬ & ",null," & dbl���Ϸ�Χ & "," & strժҪ & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        rsBill.MoveNext
    Loop
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_�¶�(rsExse As Recordset, ByVal lng����ID As Long, ByVal strҽ���� As String) As String
'���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
'������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
'���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    Dim cn�ϴ� As New ADODB.Connection, rsTemp As New ADODB.Recordset, rs�շ���� As New ADODB.Recordset

    Dim StrInput As String, arrOutput   As Variant, strReturn As String
    Dim str���ı�� As String, str����˳��� As String, str��Ա��� As String, strҽԺ���� As String
    Dim cur�����ʻ� As Double, curͳ��֧�� As Double
    Dim dbl�ܽ�� As Double, dblȫ�Է� As Double, dbl�ҹ��Ը� As Double, dbl������ As Double
    Dim dblסԺ���� As Double, dbl����ͳ��֧���ۼ� As Double, dblסԺͳ���޶� As Double, lngʵ������ As Long, dblͳ�ﱨ������ As Double
    Dim strҽ�� As String, datCurr As Date, str��ˮ�� As String, str�շ���� As String
    '������(2006-01-16):���ӱ���
    Dim i As Integer, strժҪ As String
    Dim dbl���Ϸ�Χ As Double
    
    On Error GoTo errHandle
    mlng����ID = 0         '��ʼ����ֻҪһѡ���ˣ��ͻ���ñ����̣�Ҳ�ͻ����0
    
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    rsExse.MoveFirst
    
    datCurr = zlDatabase.Currentdate
    With g��������
        .����ID = rsExse("����ID")
        
        gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", CLng(rsExse("����ID")))
        If IsNull(rsTemp("��ҳID")) = True Then
            MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
            Exit Function
        End If
        .��ҳID = rsTemp("��ҳID")
    End With
    
    '���½��д������
    Dim str����_New As String, strҽ����_New As String, str���ı��_New As String, str����_New As String
    If frmIdentify�ɶ�����.GetIdentify(TYPE_�¶�, str����_New, strҽ����_New, str���ı��_New, str����_New) = False Then
        '�����֤δͨ��
        Exit Function
    End If
    
'    If strҽ���� <> strҽ����_New Then
'        MsgBox "�ÿ����ǵ�ǰ���˵ģ�����һ�¡�", vbInformation, gstrSysName
'        Exit Function
'    End If
    If ��Ժ�Ǽ�_�¶�(g��������.����ID, g��������.��ҳID, strҽ����, False) = False Then
        Exit Function
    End If
    
    '�õ���Ժ������Ϣ���Ѿ�������֤�ģ�
    gstrSQL = "Select ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���  " & _
              "       ,NVL(A.��������,0) as סԺ����,NVL(A.�����ۼ�,0) as ����ͳ��֧���ۼ�" & _
              "       ,NVL(A.����ͳ���޶�,0) as סԺͳ���޶�,NVL(A.���ͳ���޶�,0) as ʵ������,NVL(A.���ͳ���ۼ�,0) as ͳ�ﱨ������" & _
              "  From �ʻ������Ϣ A,������ҳ B,�����ʻ� C " & _
              "  where B.����ID=[1] and B.��ҳID=[2] and A.����ID=B.����ID and A.����=[3] and A.���=to_char(B.��Ժ����,'yyyy')" & _
              "     and C.����ID=A.����ID and C.����=A.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺԤ��", lng����ID, g��������.��ҳID, TYPE_�¶�)
    str����˳��� = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))
    str���ı�� = IIf(IsNull(rsTemp("���ı��")), "", rsTemp("���ı��"))
    str��Ա��� = IIf(IsNull(rsTemp("��Ա���")), "", rsTemp("��Ա���"))
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    
    dblסԺ���� = rsTemp("סԺ����")
    dbl����ͳ��֧���ۼ� = rsTemp("����ͳ��֧���ۼ�")
    dblסԺͳ���޶� = rsTemp("סԺͳ���޶�")
    lngʵ������ = rsTemp("ʵ������")
    dblͳ�ﱨ������ = rsTemp("ͳ�ﱨ������")
    
    '������(2005-12-28):��ʾ�α����˵�סԺ�����Ϣ
    MsgBox "�òα����˵�סԺ�����Ϣ��" & vbCrLf & _
           "    סԺ����  ����" & Format(dblסԺ����, "#0.00") & "Ԫ     " & vbCrLf & _
           "    ����ͳ���޶��" & Format(dblסԺͳ���޶�, "#0.00") & "Ԫ     " & vbCrLf & _
           "    ͳ��֧���ۼƣ���" & Format(dbl����ͳ��֧���ۼ�, "#0.00") & "Ԫ     " & vbCrLf & _
           "    ͳ�ﱨ��������  " & dblͳ�ﱨ������ * 100 & "%", vbInformation, gstrSysName
        
    
    '�г������շ����
    gstrSQL = "Select ����,��� as ���� From �շ����"
    Set rs�շ���� = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
    '������һ�����Ӵ����Դﵽ���ܵ�ǰ��������Ŀ���
    Set cn�ϴ� = GetNewConnection
    
    Screen.MousePointer = vbHourglass
    
    
    If Get��ˮ��("G", strҽԺ����, str��ˮ��) = False Then Exit Function
'    If mint���õ���_�¶� = 1 Then
    '������(2006-01-16):��ʼ����¼����
    i = 1
    Do Until rsExse.EOF
        '������(2006-01-16):��ʾ��ʾ����
        g�ɶ�������Ϣ = "���ڴ��������ϸ�����Ժ" & vbCrLf & _
                        "��" & i & "����ϸ����" & rsExse.RecordCount & "����ϸ��"
        frm�ɶ�������ʾ.Show 1
        
       If g��������.��ҳID = rsExse("��ҳID") Then
       'ֻ������δ�ϴ�������,���ڱ�����ǰ�ķ��ñ����������������ҽԺ�����ķ����ܶһ�µ����
    
         '���з��÷ָ�
            Call WriteLog("���÷ָ�(" & str���ı�� & "," & ToVarchar(rsExse!ҽ����Ŀ����, 12) & ",31," & str��Ա��� & "," & Val(rsExse!�۸�) & ")")
            strReturn = DivideUp(str���ı��, ToVarchar(rsExse("ҽ����Ŀ����"), 12), "31", str��Ա���, Val(rsExse("�۸�")))
            If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
            
            dbl�ܽ�� = dbl�ܽ�� + rsExse("���")
            dblȫ�Է� = dblȫ�Է� + Val(arrOutput(1)) * rsExse("����")
            dbl�ҹ��Ը� = dbl�ҹ��Ը� + Val(arrOutput(2)) * rsExse("����")
            dbl������ = dbl������ + Val(arrOutput(3)) * rsExse("����")
        
            '������(2006-3-20):ժҪ�����ʽ ȫ�ԷѲ���|�ҹ��ԷѲ���|���Ϸ�Χ����
            strժҪ = "'" & Format(Val(arrOutput(1)) * rsExse("����"), "#0.00") & "|" & Format(Val(arrOutput(2)) * rsExse("����"), "#0.00") & "|" & Format(Val(arrOutput(3)) * rsExse("����"), "#0.00") & "'"
            dbl���Ϸ�Χ = Val(arrOutput(3)) * rsExse("����")
                
            If IIf(IsNull(rsExse("�Ƿ��ϴ�")), "0", rsExse("�Ƿ��ϴ�")) = "0" Then
            
                rs�շ����.Filter = "���� = '" & rsExse("�շ����") & "'"
                If rs�շ����.EOF = False Then str�շ���� = rs�շ����("����")
    
                '�ڶ�������˳�������Ϊ���㵥��
                StrInput = str����˳��� & "|" & str����˳���
                StrInput = StrInput & "|" & rsExse("NO") & "_" & rsExse("���") & "_" & rsExse("��¼����") & "_" & rsExse("��¼״̬")
                StrInput = StrInput & "|" & strҽ���� & "|" & str���ı�� & "|" & strҽԺ���� & "|000"
                StrInput = StrInput & "|" & ToVarchar(rsExse("ҽ����Ŀ����"), 12)  'ҽ����ˮ��
                StrInput = StrInput & "|" & ToVarchar(str�շ����, 10)      '�շѴ�������
                StrInput = StrInput & "|" & Format(rsExse("����"), "0.00")
                StrInput = StrInput & "|" & Format(rsExse("�۸�"), "0.00")
                StrInput = StrInput & "|" & Format(rsExse("���"), "0.00")
                StrInput = StrInput & "|" & arrOutput(4)                       '�Ը�����
                StrInput = StrInput & "|" & Format(Val(arrOutput(1)) * rsExse("����"), "0.00") 'ȫ�ԷѲ���
                StrInput = StrInput & "|" & Format(Val(arrOutput(2)) * rsExse("����"), "0.00") '�ҹ��ԷѲ���
                StrInput = StrInput & "|" & Format(Val(arrOutput(3)) * rsExse("����"), "0.00") '����������
                StrInput = StrInput & "||31"                                   '�����־��֧�����
                StrInput = StrInput & "|" & ToVarchar(rsExse("��������"), 56)  '������������
                StrInput = StrInput & "|" & ToVarchar(rsExse("ҽ��"), 20)      '��������ҽ��
                StrInput = StrInput & "|" & ToVarchar(rsExse("��������"), 56)  '�ܵ���������
                StrInput = StrInput & "|" & ToVarchar(rsExse("ҽ��"), 20)      '�ܵ�����ҽ��
                StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 20)        '������
                StrInput = StrInput & "|" & Format(rsExse("�Ǽ�ʱ��") + rsExse("���") / 24 / 3600, "yyyy-MM-dd HH:mm:ss") '����ʱ��
                StrInput = StrInput & "|" & ToVarchar(rsExse("�շ�����"), 200)       '�շ���Ŀ
                StrInput = StrInput & "|" & ToVarchar(rsExse("���"), 200)       '���
                StrInput = StrInput & "|" & ToVarchar(rsExse("����"), 200)       '����
                StrInput = StrInput & "|"                                        '��λ
                StrInput = StrInput & "|||"                                      'Ӣ��������ѧ��
                'modify by ccy ,Ψһ
                'StrInput = StrInput & Format(rsExse("�Ǽ�ʱ��"), "yyyyMMddHHmmss") & rsExse("���") & "#"
                '������(2006-02-16):����п���żȻ�ظ�,�޸�Ϊ���·�ʽ
                StrInput = StrInput & rsExse("NO") & "_" & rsExse("���") & "_" & rsExse("��¼����") & "_" & rsExse("��¼״̬") & "#"      '���
    
                Call WriteLog("DataUnloading(" & StrInput & "," & str��ˮ�� & "," & str���ı�� & ")")
                strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
                If JudgeReturn(strReturn, arrOutput) = False Then
                   MsgBox "�ϴ�" & rsExse("No") & "�ĵ�" & rsExse("���") & "��(��¼״̬Ϊ" & rsExse("��¼״̬") & ")���ü�¼ʱ���ִ�����֪ͨ����Ա��飡"
                   Exit Function
                End If
            End If
            '�Ѿ��ϴ������Լ��ϴ��ɹ��ģ�����Ҫ���±��ձ���ȱ�־
            gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ",'" & rsExse!ҽ����Ŀ���� & "'," & dbl���Ϸ�Χ & "," & strժҪ & ")"
            cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
        Else
            If IIf(IsNull(rsExse("�Ƿ��ϴ�")), "0", rsExse("�Ƿ��ϴ�")) = "0" Then
                MsgBox "�ò��˿��ܴ��ڱ�����ǰδ�����δ�ϴ��ķ��ã�" & vbCrLf & _
                       "���ҽ�����ص��ܽ����ҽԺ�ڲ����ܽ�һ�£������Щ���ý������ʴ���", vbInformation, gstrSysName
                gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsExse("NO") & "'," & rsExse("���") & "," & rsExse("��¼����") & "," & rsExse("��¼״̬") & ",'" & rsExse!ҽ����Ŀ���� & "'," & dbl���Ϸ�Χ & "," & strժҪ & ")"
                cn�ϴ�.Execute gstrSQL, , adCmdStoredProc
            End If
        End If
        '������(2006-01-16):��������
        i = i + 1
        rsExse.MoveNext
    Loop

    '������(2006-01-16):��ʾ��ʾ����
    g�ɶ�������Ϣ = "���ڽ���Ԥ���㣬���Ժ�!"
    frm�ɶ�������ʾ.Show 1
    
    '����Ԥ����
    '2107,404.2,44020,0,37,0,0,1604,103,400,.824
    Call WriteLog("Ԥ����:" & dbl�ܽ�� & "," & dblסԺ���� & "," & dblסԺͳ���޶� & "," & dbl����ͳ��֧���ۼ� & "," & lngʵ������ & "," & 0 & "," & 0 & "," & _
                dbl������ & "," & dblȫ�Է� & "," & dbl�ҹ��Ը� & "," & dblͳ�ﱨ������)
    strReturn = CalculateFeeCD(dbl�ܽ��, dblסԺ����, dblסԺͳ���޶�, dbl����ͳ��֧���ۼ�, lngʵ������, 0, 0, _
                dbl������, dblȫ�Է�, dbl�ҹ��Ը�, dblͳ�ﱨ������)
    Call WriteLog("����:" & strReturn)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    curͳ��֧�� = Val(arrOutput(2))
    
    '������ʱ���ݣ�Ϊ���������׼��
    With g��������
        .�������ý�� = dbl�ܽ��
        .ʵ������ = Val(arrOutput(1))
        .ͳ�ﱨ����� = curͳ��֧��
        .�����Ը���� = Val(arrOutput(4))
    
        .����ͳ���� = dbl������
        .ȫ�Էѽ�� = dblȫ�Է�
        .�����Ը���� = dbl�ҹ��Ը�
        .�����ʻ�֧�� = Val(arrOutput(3)) '����ͳ���Ը�����
    End With
    
    סԺ�������_�¶� = "ҽ������;" & curͳ��֧�� & ";0"
    
    mlng����ID = lng����ID  '��ʾ�ò����Ѿ��������������
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_�¶�(lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
'����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
'���أ����׳ɹ�����true�����򣬷���false
'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
'      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
'      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset, StrInput As String, arrOutput  As Variant, str��ˮ�� As String, strReturn As String
    Dim strҽԺ��� As String, strҽԺ���� As String, strҽ���� As String
    
    Dim cur�ʻ������ۼ� As Currency, cur�ʻ�֧���ۼ� As Currency
    Dim cur����ͳ���ۼ� As Currency, curͳ�ﱨ���ۼ� As Currency
    Dim intסԺ�����ۼ� As Integer, datCurr As Date, cur���� As Currency
    
    If mlng����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò���û�����ҽ����Ԥ������������ܽ��н��㡣"
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    datCurr = zlDatabase.Currentdate
    
    '�õ���Ժ������Ϣ
    gstrSQL = "Select ҽ����,˳��� as �������,����֤�� as ���ı��,��Ա��� as ��Ա���,C.��λ����  " & _
              " ,D.����,D.�Ա�,D.�������� " & _
              " ,nvl(A.��������,0) as סԺ����,nvl(A.�����ۼ�,0) as ����ͳ��֧���ۼ�,nvl(A.����ͳ���޶�,0) as סԺͳ���޶�,nvl(A.���ͳ���޶�,0) as ʵ������,nvl(A.���ͳ���ۼ�,0) as ͳ�ﱨ������" & _
              "  From �ʻ������Ϣ A,������ҳ B,�����ʻ� C,������Ϣ D " & _
              "  where B.����ID=[1] and B.��ҳID=[2] and A.����ID=B.����ID and A.����=[3] and A.���=to_char(B.��Ժ����,'yyyy')" & _
              "     and C.����ID=A.����ID and C.����=A.����   and B.����ID=D.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "סԺԤ��", lng����ID, g��������.��ҳID, TYPE_�¶�)
    If GetҽԺ����(strҽԺ����, rsTemp("���ı��")) = False Then Exit Function
    If GetҽԺ����(strҽԺ���, rsTemp("���ı��"), True) = False Then Exit Function
    
    cur���� = rsTemp("סԺ����")
    '���ý���
    If Get��ˮ��("F", strҽԺ����, str��ˮ��) = False Then Exit Function
    StrInput = IIf(IsNull(rsTemp("�������")), "", rsTemp("�������"))  '����˳���
    StrInput = StrInput & "|" & ToVarchar(rsTemp("���ı��"), 4)        '�����ı��
    StrInput = StrInput & "|" & ToVarchar(rsTemp("ҽ����"), 20)         '���˱���
    StrInput = StrInput & "|" & ToVarchar(rsTemp("��λ����"), 12)        '��λ����
    StrInput = StrInput & "|" & ToVarchar(rsTemp("����"), 20)            '����
    StrInput = StrInput & "|" & ToVarchar(IIf(rsTemp("�Ա�") = "Ů", "2", "1"), 4)         '�Ա�
    StrInput = StrInput & "|" & Format(rsTemp("��������"), "yyyy-MM-dd") '��������
    StrInput = StrInput & "|" & Format(rsTemp("ʵ������"), "0")         'ʵ������
    StrInput = StrInput & "|"                                           '�ɷ�����
    StrInput = StrInput & "|" & strҽԺ����
    StrInput = StrInput & "|000"                                        '��Ժ����
    StrInput = StrInput & "|" & strҽԺ���                             'ҽԺ���
    StrInput = StrInput & "|31"                                         '֧�����
    StrInput = StrInput & "|0"                                          '���ֲ���־
    StrInput = StrInput & "|000"                                        '���ֲ�����
    StrInput = StrInput & "|" & ToVarchar(rsTemp("�������"), 20)       '������
    StrInput = StrInput & "|"                                           '�˵����
    StrInput = StrInput & "|" & ToVarchar(rsTemp("��Ա���"), 20)       'ҽ����Ա���
    With g��������
        StrInput = StrInput & "|" & Format(cur����, "0.00")        '����
        StrInput = StrInput & "|" & Format(.�������ý��, "0.00")    '�����ܶ�
        StrInput = StrInput & "|" & Format(.ȫ�Էѽ��, "0.00")      'ȫ�ԷѲ���
        StrInput = StrInput & "|" & Format(.�����Ը����, "0.00")    '�ҹ��Ը�����
        StrInput = StrInput & "|" & Format(.����ͳ����, "0.00")    '����������
        StrInput = StrInput & "|" & Format(.ʵ������, "0.00")      '�������߲���
        StrInput = StrInput & "|" & Format(.ͳ�ﱨ�����, "0.00")    '����ҽ��ͳ��֧������
        StrInput = StrInput & "|" & Format(.�����ʻ�֧��, "0.00")    '����ҽ��ͳ���Ը�����
        StrInput = StrInput & "|" & Format(0, "0.00")                '����ͳ��֧������
        StrInput = StrInput & "|" & Format(0, "0.00")                '����ͳ���Ը�����
        StrInput = StrInput & "|" & Format(.�����Ը����, "0.00")    '�����Ը����
        StrInput = StrInput & "|" & Format(0, "0.00")                '�����˻�֧�����
    End With
    StrInput = StrInput & "|"                                              '��Ʊ��
    StrInput = StrInput & "|" & ToVarchar(UserInfo.����, 20)               '������
    StrInput = StrInput & "|" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "#"    '����ʱ��

    Call WriteLog("DataUnloading(" & StrInput & "," & str��ˮ�� & "," & rsTemp!���ı�� & ")")
    strReturn = DataUnloading(StrInput, str��ˮ��, rsTemp("���ı��"))
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    
    '��д�����
    
    '�ʻ������Ϣ
    Call Get�ʻ���Ϣ(TYPE_�¶�, lng����ID, Year(datCurr), intסԺ�����ۼ�, cur�ʻ������ۼ�, cur�ʻ�֧���ۼ�, cur����ͳ���ۼ�, curͳ�ﱨ���ۼ�)
    If intסԺ�����ۼ� = 0 Then intסԺ�����ۼ� = GetסԺ����(lng����ID)
            
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    With g��������
        gstrSQL = "zl_�ʻ������Ϣ_insert(" & lng����ID & "," & TYPE_�¶� & "," & Year(datCurr) & "," & _
            cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & _
            cur����ͳ���ۼ� + .����ͳ���� & "," & _
            curͳ�ﱨ���ۼ� + .ͳ�ﱨ����� & "," & intסԺ�����ۼ� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        
        gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�¶� & "," & lng����ID & "," & _
            Year(datCurr) & "," & cur�ʻ������ۼ� & "," & cur�ʻ�֧���ۼ� & "," & cur����ͳ���ۼ� & "," & _
            curͳ�ﱨ���ۼ� & "," & intסԺ�����ۼ� & "," & cur���� & ",NULL," & .ʵ������ & "," & g��������.�������ý�� & _
            "," & .ȫ�Էѽ�� & "," & .�����Ը���� & "," & .����ͳ���� & "," & .ͳ�ﱨ����� & ",0," & .�����Ը���� & ",0,''," & .��ҳID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        
        '���ս������
        gstrSQL = "zl_���ս������_insert(" & lng����ID & ",0," & .����ͳ���� & "," & .ͳ�ﱨ����� & ",NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    End With
        
    סԺ����_�¶� = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function סԺ�������_�¶�(lng����ID As Long) As Boolean
    '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    
    סԺ�������_�¶� = False
End Function

Private Function GetҽԺ����(ByRef strҽԺ���� As String, ByVal str�����ı��� As String, Optional ByVal blnҽԺ��� As Boolean) As Boolean
'���ܣ��õ�ҽԺ��ҽ������
    Dim strReturn As String, arrOutput As Variant
    Dim strTemp As String, varList As Variant, lngIndex As Long, strHospital As String
    
    On Error GoTo errHandle
    
    strReturn = GetHospitalInfo()
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '���Ƚ��ִ���ԭ
    strTemp = ""
    For lngIndex = 1 To UBound(arrOutput)
        strTemp = strTemp & "|" & arrOutput(lngIndex)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2) '֧����һ�����ӵ�|
    If Right(strTemp, 1) = "#" Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1) '֧������#
    
    varList = Split(strTemp, "$")
    
    For lngIndex = 0 To UBound(varList)
        arrOutput = Split(varList(lngIndex), "|")
        
        If UBound(arrOutput) > 3 Then
            If arrOutput(3) = str�����ı��� Then
                If blnҽԺ��� = True Then
                    strHospital = arrOutput(2) 'ҽԺ���
                Else
                    strHospital = arrOutput(0) 'ҽԺ����
                End If
            End If
        End If
    Next
    
    If strHospital <> "" Then
        strҽԺ���� = strHospital
        GetҽԺ���� = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get���ı���() As String
'���ܣ��õ�ҽԺ��ҽ������
    Dim strReturn As String, arrOutput As Variant
    Dim strTemp As String, varList As Variant, lngIndex As Long, strHospital As String
    Dim strҽԺ���� As String, rsTmp As New ADODB.Recordset
        
    On Error GoTo errHandle
    '��ȡҽԺ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_�¶�)
    
    If IsNull(rsTmp("ҽԺ����")) = True Then
        MsgBox "����δ����ҽԺ��ţ��޷�ִ��ҽ�����ף�", vbExclamation, gstrSysName
        Exit Function
    End If
    strҽԺ���� = rsTmp!ҽԺ����
    
    strReturn = GetHospitalInfo()
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '���Ƚ��ִ���ԭ
    strTemp = ""
    For lngIndex = 1 To UBound(arrOutput)
        strTemp = strTemp & "|" & arrOutput(lngIndex)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2) '֧����һ�����ӵ�|
    If Right(strTemp, 1) = "#" Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1) '֧������#
    
    varList = Split(strTemp, "$")
    
    For lngIndex = 0 To UBound(varList)
        arrOutput = Split(varList(lngIndex), "|")
        
        If UBound(arrOutput) > 3 Then
            If arrOutput(0) = strҽԺ���� Then
                Get���ı��� = arrOutput(3) '���ı���
                Exit For
            End If
        End If
    Next
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function JudgeReturn(ByRef strReturn As String, ByRef varOut As Variant) As Boolean
'���ܣ��жϷ���ֵ�Ƿ�Ϸ���
    Dim varArray As Variant, lngReturn As Long, lngPos As Long
    Dim strSuggest
    
    strReturn = TruncZero(strReturn)
    lngPos = InStr(strReturn, "#")
    If lngPos > 0 Then
        strReturn = Mid(strReturn, 1, lngPos - 1)
    End If
    
    varArray = Split(strReturn, "|")
    
    lngReturn = Val(varArray(0))
    If lngReturn < 0 Then
        'ҵ�����ʧ��
        If UBound(varArray) > 0 Then
            strReturn = "ҽ��ҵ����ʧ�ܡ�" & vbCrLf & "�����:" & lngReturn & vbCrLf & varArray(1)
        Else
            strReturn = "ҽ��ҵ����ʧ�ܡ�"
        End If
        
        Select Case lngReturn
            Case -1101
                strSuggest = "�����������ʶ�𲢻�ȡ�µ���ˮ�š�"
            Case -1102, -1210, -1216, -1404, -1405, -1502
                strSuggest = "��Ҫ������˾��顣"
            Case -1103
                strSuggest = "֧������������ȷ��"
            Case -1201, -1203, -1204, -1205, -1213, -1215, -1217, -1220, -1804
                strSuggest = "��Ҫ���籣��ȷ�ϡ�"
            Case -1206
                strSuggest = "�����õ������뿨���������뿨�����루�ſ������µ����룩��"
            Case -1207
                strSuggest = "�ò��˳ֵĿ�������Ч������Ҫʱ���籣��ȷ�ϡ�"
            Case -1208
                strSuggest = "�����������뿨�����ɲ��˵Ĵſ�����ˢ��"
            Case -1209, -1212, -1301
                strSuggest = "������ȷ���롣"
            Case -1214
                strSuggest = "���볤��Ϊ6�����롣"
            Case -1302
                strSuggest = "�����޸����롣"
            Case -1402
                strSuggest = "���ܶԴ˲���ʹ������ͬ�ľ���˳��Ž������ˡ�"
            Case -1501, -1601
                strSuggest = "�����Ѵ�����ͬ��¼��"
                JudgeReturn = True
                Exit Function
        End Select
       
        If strSuggest <> "" Then
            strReturn = strReturn & vbCrLf & vbCrLf & "���鴦������" & strSuggest
        End If
        
        Screen.MousePointer = vbDefault
        MsgBox lngReturn & ":" & strReturn, vbExclamation, gstrSysName
        Exit Function
    End If
    
    varOut = varArray
    JudgeReturn = True
End Function

Private Function Get��ˮ��(ByVal str��־ As String, ByVal strҽԺ���� As String, ByRef str��ˮ�� As String) As Boolean
    Dim datCurr As Date
    
    datCurr = zlDatabase.Currentdate
    '[��Ϣ��־+ҽԺ����+YYMMDD+6λ��ˮ��]
    str��ˮ�� = str��־ & strҽԺ���� & Format(datCurr, "yyMMddHHmmss")
    Get��ˮ�� = True
End Function

Public Function ҽ����Ŀ_�¶�(rsTemp As ADODB.Recordset) As Boolean
'���ܣ�ҽ������ҩƷĿ¼��ѯ
    Dim str���� As String, str���� As String, str���� As String
    Dim strPath As String, strFile As String, strReturn As String, arrOutput As Variant
    Dim lngFile  As Long, str���ı�� As String
    
    
    str���ı�� = Get���ı���
    If str���ı�� = "" Then Exit Function
    
    '���ýӿڣ������ļ�
    strFile = Space(255)
    GetTempPath 255, strFile
    strPath = TrimStr(strFile)
    strFile = strPath & "MakeTxt.txt"
    
    strReturn = MakeTxt(strFile, strPath & "Temp.txt") '����Ŀ¼��Ȼ��Ҫ,��Ҳ���봫
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    lngFile = FreeFile
    Open strFile For Input Access Read As lngFile
    
    On Error GoTo errHandle
    Do Until EOF(lngFile)
        Line Input #lngFile, strReturn
        
        arrOutput = Split(strReturn, vbTab)
        If UBound(arrOutput) >= 11 Then
            str���� = arrOutput(0)
            str���� = ToVarchar(arrOutput(1), 40)
            str���� = ToVarchar(zlCommFun.SpellCode(arrOutput(1)), 10)
        End If
        If str���� <> "" And arrOutput(11) = str���ı�� Then
            'ֻȡ��ǰ���ĵ�ҽ������,�������ĵı�����ܲ�ͬ
            rsTemp.AddNew Array("CLASSCODE", "CODE", "NAME", "PY"), Array("1", str����, str����, str����)
            rsTemp.Update
        End If
    Loop
    Close #lngFile
    Kill strFile
    Kill strPath & "Temp.txt"
    
    ҽ����Ŀ_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Close #lngFile
    
End Function

Public Function ������_�¶�(ByVal str������ As String, strҽ���� As String, str���� As String, str���ı�� As String) As Boolean
'���ܣ����ſ����ݽ��н���
    Dim strReturn  As String, arrOutput As Variant
    
    On Error GoTo errHandle
    
    If str������ = "" Then
        MsgBox "���Ƚ���ˢ��������", vbInformation, gstrSysName
        Exit Function
    End If
    
    strReturn = GetKard(str������)  '����Ϊҽ���š����š�ҽԺ���롢�����ı��
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    strҽ���� = arrOutput(1)
    str���� = arrOutput(2)
    str���ı�� = arrOutput(3)
    ������_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ��������_�¶�(ByVal str���� As String, ByVal strҽ���� As String, ByVal str���ı�� As String, _
            ByVal strԭ���� As String, ByVal str������ As String) As Boolean

'���ܣ��޸��û�����
    Dim StrInput As String, arrOutput   As Variant, strReturn As String
    Dim strҽԺ���� As String, str��ˮ�� As String
    
    On Error GoTo errHandle
    
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Function
    If Get��ˮ��("B", strҽԺ����, str��ˮ��) = False Then Exit Function
    
    StrInput = str���� & "|" & strҽ���� & "|" & str���ı�� & "|" & strԭ���� & "|" & str������ & "#"
    
    strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    MsgBox "�����뱣��ɹ���", vbInformation, gstrSysName
    ��������_�¶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub �˶��ʻ�֧��_�¶�(ByVal lng����ID As Long)
    Dim int��¼��_OUT As Integer, cur���_OUT As Currency
    Dim int��¼��_Client As Integer, cur���_Client As Currency
    Dim lng��ҳID As Long
    Dim StrInput As String, strReturn As String, arrOutput
    Dim str���ı�� As String, str������� As String, strҽԺ���� As String, strҽԺ��� As String, str��ˮ�� As String
    Dim rsTemp As New ADODB.Recordset
    '���Գ�Ժ���˽��м��
    On Error GoTo errHand
    
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "�ò��˻�δ��Ժ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ȡ�ϴ�סԺ����ҳID����Ϊ�ù�����Ҫ���ڳ�Ժ��ʹ�ã���˼ٶ��ò���δ�ٴ���Ժ
    gstrSQL = "Select nvl(סԺ����,1) ��ҳID From ������Ϣ Where ����ID=" & lng����ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ϴ�סԺʱ����ҳID", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    
    'ȡ�ʻ�֧����¼����֧�����
    gstrSQL = "Select Sum(A.��Ԥ��) �ʻ�֧��,Count(*) ��¼��  " & _
             " From ����Ԥ����¼ A, " & _
             "      (Select ����ID,����ID  " & _
             "      From סԺ���ü�¼ " & _
             "      Where ����ID=[1] And ��ҳID=[2]) B " & _
             " Where A.����ID=B.����ID And A.���㷽ʽ='�����ʻ�'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ʻ�֧�����¼��", lng����ID, lng��ҳID)
    int��¼��_Client = Nvl(rsTemp!��¼��, 0)
    cur���_Client = Nvl(rsTemp!�ʻ�֧��, 0)
    
    '��ȡ������Ϣ
    gstrSQL = " Select ����֤�� ���ı��,˳��� ������� From �����ʻ� " & _
            " Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lng����ID, TYPE_�¶�)
    str������� = rsTemp!�������
    str���ı�� = rsTemp!���ı��
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Sub
    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Sub
    
    '���ú˶Խӿ�
    If Get��ˮ��("H", strҽԺ����, str��ˮ��) = False Then Exit Sub
    StrInput = ToVarchar(str���ı��, 4)
    StrInput = StrInput & "|" & ToVarchar(strҽԺ����, 8)
    StrInput = StrInput & "|" & str�������
    StrInput = StrInput & "|" & str������� & "|%#"
    
'    MsgBox "�˶��ʻ�֧����DataUnloading" & strInput
    strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
    
    '�����ҽ�����Ľ��յ��Ĳ�����������ʾ��1-��¼��;2-֧���
    int��¼��_OUT = arrOutput(1)
    cur���_OUT = arrOutput(2)
    
    If Format(cur���_OUT, "#####0.00;-#####0.00;0;") <> Format(cur���_Client, "#####0.00;-#####0.00;0;") Then
        MsgBox "�����ʻ�֧������ҽ�����ķ��صĲ�һ�£����飡" & vbCrLf & _
               "����ʵ���ʻ�֧����" & cur���_Client & String(4, " ") & "ҽ������ͳ�Ƴ����ʻ�֧����" & cur���_OUT & vbCrLf & _
               "�����ʻ�֧��������" & int��¼��_Client & String(4, " ") & "ҽ������ͳ�Ƴ���֧��������" & int��¼��_OUT
    Else
        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub �˶����Ժ_�¶�(ByVal lng����ID As Long)
    '���Գ�Ժ���˽��м��
    Dim int��¼��_OUT As Integer, cur���_OUT As Currency
    Dim int��¼��_Client As Integer, cur���_Client As Currency
    Dim lng��ҳID As Long
    Dim StrInput As String, strReturn As String, arrOutput
    Dim str���ı�� As String, str������� As String, strҽԺ���� As String, strҽԺ��� As String, str��ˮ�� As String
    Dim rsTemp As New ADODB.Recordset
    '���Գ�Ժ���˽��м��
    On Error GoTo errHand
    
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "�ò��˻�δ��Ժ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    int��¼��_Client = 1
    
    '��ȡ������Ϣ
    gstrSQL = " Select ����֤�� ���ı��,˳��� ������� From �����ʻ� " & _
            " Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lng����ID, TYPE_�¶�)
    str������� = rsTemp!�������
    str���ı�� = rsTemp!���ı��
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Sub
    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Sub
    
    '���ú˶Խӿ�
    If Get��ˮ��("I", strҽԺ����, str��ˮ��) = False Then Exit Sub
    StrInput = ToVarchar(str���ı��, 4)
    StrInput = StrInput & "|" & ToVarchar(strҽԺ����, 8)
    StrInput = StrInput & "|" & str�������
    StrInput = StrInput & "|" & str������� & "|%#"
    
'    MsgBox "�˶����Ժ��¼��DataUnloading" & strInput
    strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
    
    '�����ҽ�����Ľ��յ��Ĳ�����������ʾ��1-��¼����
    int��¼��_OUT = arrOutput(1)
    
    If int��¼��_OUT <> int��¼��_Client Then
        MsgBox "�������Ժ��¼��ҽ�����ķ��صĲ�һ�£����飡" & vbCrLf & _
               "�������Ժ��¼����" & int��¼��_Client & String(4, " ") & "ҽ������ͳ�Ƴ������Ժ��¼����" & int��¼��_OUT
    Else
        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub �˶Է��ý���_�¶�(ByVal lng����ID As Long)
    Dim int��¼��_OUT As Integer, cur���_OUT As Currency
    Dim cur����_OUT As Currency, curȫ�Է�_OUT As Currency
    Dim cur�����Ը�_OUT As Currency, curʵ������_OUT As Currency
    Dim curͳ��֧��_OUT As Currency, curͳ���Ը�_OUT As Currency
    Dim cur�����Ը�_OUT As Currency, cur�ʻ�֧��_OUT As Currency
    Dim int��¼��_Client As Integer, cur���_Client As Currency
    Dim cur����_Client As Currency, curȫ�Է�_Client As Currency
    Dim cur�����Ը�_Client As Currency, curʵ������_Client As Currency
    Dim curͳ��֧��_Client As Currency, curͳ���Ը�_Client As Currency
    Dim cur�����Ը�_Client As Currency, cur�ʻ�֧��_Client As Currency
    Dim lng��ҳID As Long
    Dim StrInput As String, strReturn As String, arrOutput
    Dim str���ı�� As String, str������� As String, strҽԺ���� As String, strҽԺ��� As String, str��ˮ�� As String
    Dim rsTemp As New ADODB.Recordset
    '���Գ�Ժ���˽��м��
    On Error GoTo errHand
    
    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
        MsgBox "�ò��˻�δ��Ժ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'ȡ�ϴ�סԺ����ҳID����Ϊ�ù�����Ҫ���ڳ�Ժ��ʹ�ã���˼ٶ��ò���δ�ٴ���Ժ
    gstrSQL = "Select nvl(סԺ����,1) ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ϴ�סԺʱ����ҳID", lng����ID)
    lng��ҳID = rsTemp!��ҳID
    
    'ȡ�ʻ�֧����¼����֧�����
    gstrSQL = "SELECT SUM(�������ý��) ��������,SUM(����ͳ����) ����ͳ��,SUM(ͳ�ﱨ�����) ͳ�ﱨ��, " & _
             " SUM(�����Ը����) �����Ը�,SUM(����) ����,SUM(ʵ������) ʵ������," & _
             " SUM(�����Ը����) �����Ը�,SUM(�����ʻ�֧��) �����ʻ�֧��,Count(*) ��¼�� " & _
             " FROM  " & _
             "      (SELECT ����ID,����ID FROM סԺ���ü�¼ " & _
             "      WHERE ����ID=[1] AND ��ҳID= [2]" & _
             "      ) A,���ս����¼ B " & _
             " WHERE A.����ID=B.����ID AND B.��¼ID=A.����ID AND B.����=[3] AND B.����=2 " & _
             " GROUP BY A.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ʻ�֧�����¼��", lng����ID, lng��ҳID, TYPE_�¶�)
    int��¼��_Client = Nvl(rsTemp!��¼��, 0)
    cur���_Client = Nvl(rsTemp!��������, 0)
    curͳ��֧��_Client = Nvl(rsTemp!ͳ�ﱨ��, 0)
    cur�ʻ�֧��_Client = Nvl(rsTemp!�����ʻ�֧��, 0)
    
    '��ȡ������Ϣ
    gstrSQL = " Select ����֤�� ���ı��,˳��� ������� From �����ʻ� " & _
            " Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lng����ID, TYPE_�¶�)
    str������� = rsTemp!�������
    str���ı�� = rsTemp!���ı��
    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Sub
    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Sub
    
    '���ú˶Խӿ�
    If Get��ˮ��("J", strҽԺ����, str��ˮ��) = False Then Exit Sub
    StrInput = ToVarchar(str���ı��, 4)
    StrInput = StrInput & "|" & ToVarchar(strҽԺ����, 8)
    StrInput = StrInput & "|" & str�������
    StrInput = StrInput & "|" & str������� & "|%|%|%#"
    
'    MsgBox "�˶Է��ý��㣺DataUnloading" & strInput
    strReturn = DataUnloading(StrInput, str��ˮ��, str���ı��)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
    
    '�����ҽ�����Ľ��յ��Ĳ�����������ʾ��1-��¼��;2-֧���
    int��¼��_OUT = arrOutput(1)
    cur����_OUT = arrOutput(2)
    cur���_OUT = arrOutput(3)
    curȫ�Է�_OUT = arrOutput(4)
    cur�����Ը�_OUT = arrOutput(5)
    'cur����ͳ��_OUT = arrOutput(6)
    curʵ������_OUT = arrOutput(7)
    curͳ��֧��_OUT = arrOutput(8)
    curͳ���Ը�_OUT = arrOutput(9)
    cur�����Ը�_OUT = arrOutput(10)
    cur�ʻ�֧��_OUT = arrOutput(11)
    
    'ֻҪͳ��֧�����ʻ�֧���������ܶ�һ�¼���
    If Not (Format(cur���_OUT, "#####0.00;-#####0.00;0;") = Format(cur���_Client, "#####0.00;-#####0.00;0;") _
    And Format(curͳ��֧��_OUT, "#####0.00;-#####0.00;0;") = Format(curͳ��֧��_Client, "#####0.00;-#####0.00;0;") _
    And Format(cur�ʻ�֧��_OUT, "#####0.00;-#####0.00;0;") = Format(cur�ʻ�֧��_Client, "#####0.00;-#####0.00;0;")) Then
        MsgBox "���ؽ���������ҽ�����ķ��صĲ�һ�£����飡" & vbCrLf & _
               "��ҽ���������ܶ" & cur���_OUT & String(4, " ") & "ͳ��֧����" & curͳ��֧��_OUT & String(4, " ") & "�ʻ�֧����" & cur�ʻ�֧��_OUT & vbCrLf & _
               "�����أ������ܶ" & cur���_Client & String(4, " ") & "ͳ��֧����" & curͳ��֧��_Client & String(4, " ") & "�ʻ�֧����" & cur�ʻ�֧��_Client
    Else
        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub �˶Է�����ϸ_�¶�(ByVal lng����ID As Long)
'    Dim int��¼��_OUT As Integer, cur���_OUT As Currency
'    Dim int��¼��_Client As Integer, cur���_Client As Currency
'    Dim lng��ҳID As Long
'    Dim strInput As String, strReturn As String, arrOutput
'    Dim str���ı�� As String, str������� As String, strҽԺ���� As String, strҽԺ��� As String, str��ˮ�� As String
'    Dim rsTemp As New ADODB.Recordset
'    '���Գ�Ժ���˽��м��
'    On Error GoTo ErrHand
'
'    If Not ҽ�������Ѿ���Ժ(lng����ID) Then
'        MsgBox "�ò��˻�δ��Ժ��", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'    'ȡ�ϴ�סԺ����ҳID����Ϊ�ù�����Ҫ���ڳ�Ժ��ʹ�ã���˼ٶ��ò���δ�ٴ���Ժ
'    gstrSQL = "Select nvl(סԺ����,1) ��ҳID From ������Ϣ Where ����ID=" & lng����ID
'    Call OpenRecordset(rsTemp, "���ϴ�סԺʱ����ҳID")
'    lng��ҳID = rsTemp!��ҳID
'
'    'ȡ�ʻ�֧����¼����֧�����
'    gstrSQL = "Select Sum(A.��Ԥ��) �ʻ�֧��,Count(*) ��¼��  " & _
'             " From ����Ԥ����¼ A, " & _
'             "      (Select ����ID,����ID  " & _
'             "      From ���˷��ü�¼ " & _
'             "      Where ����ID=1 And ��ҳID=1) B " & _
'             " Where A.����ID=B.����ID And A.���㷽ʽ='�����ʻ�'"
'    Call OpenRecordset(rsTemp, "ȡ�ʻ�֧�����¼��")
'    int��¼��_Client = NVL(rsTemp!��¼��, 0)
'    cur���_Client = NVL(rsTemp!�ʻ�֧��, 0)
'
'    '��ȡ������Ϣ
'    gstrSQL = " Select ����֤�� ���ı��,˳��� ������� From �����ʻ� " & _
'            " Where ����ID=" & lng����ID & " And ����=" & TYPE_�¶�
'    Call OpenRecordset(rsTemp, "��ȡ������Ϣ")
'    str������� = rsTemp!�������
'    str���ı�� = rsTemp!���ı��
'    If GetҽԺ����(strҽԺ����, str���ı��) = False Then Exit Sub
'    If GetҽԺ����(strҽԺ���, str���ı��, True) = False Then Exit Sub
'
'    '���ú˶Խӿ�
'    If Get��ˮ��("H", strҽԺ����, str��ˮ��) = False Then Exit Sub
'    strInput = ToVarchar(str���ı��, 4)
'    strInput = strInput & "|" & ToVarchar(strҽԺ����, 8)
'    strInput = strInput & "|" & str�������
'    strInput = strInput & "|" & str������� & "|%#"
'
'    strReturn = DataUnloading(strInput, str��ˮ��, str���ı��)
'    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
'
'    '�����ҽ�����Ľ��յ��Ĳ�����������ʾ��1-��¼��;2-֧���
'    int��¼��_OUT = arrOutput(1)
'    cur���_OUT = arrOutput(2)
'
'    If Format(cur���_OUT, "#####0.00;-#####0.00;0;") <> Format(cur���_Client, "#####0.00;-#####0.00;0;") Then
'        MsgBox "�����ʻ�֧������ҽ�����ķ��صĲ�һ�£����飡" & vbCrLf & _
'               "����ʵ���ʻ�֧����" & cur���_Client & String(4, " ") & "ҽ������ͳ�Ƴ����ʻ�֧����" & cur���_OUT & vbCrLf & _
'               "�����ʻ�֧��������" & int��¼��_Client & String(4, " ") & "ҽ������ͳ�Ƴ���֧��������" & int��¼��_OUT
'    Else
'        MsgBox "������ȷ���󣬺˶Գɹ���", vbInformation, gstrSysName
'    End If
'    Exit Sub
'ErrHand:
'    If ErrCenter = 1 Then Resume
End Sub

Private Sub WriteLog(ByVal strInfo As String)
    Call LogWrite("ҽ���ӿڵ�����־", glngModul, "ҽ���ӿڷ���", strInfo)
End Sub
