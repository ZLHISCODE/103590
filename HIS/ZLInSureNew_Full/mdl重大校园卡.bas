Attribute VB_Name = "mdl�ش�У԰��"
Option Explicit
#Const gverControl = 99
Public gintComport_�ش�У԰�� As Long
Private Type �������
    �����豸            As Long
    ������              As Long
    ����ˮ��            As Long
    ����                As String
    �����ѷ���          As Long
    ����Ч��            As String
    ����                As String
    ע������            As String
    ���֤��            As String
    ����Ǯ��1���       As Long   '�Է�Ϊ��λ
    ����Ǯ��2���       As Long   '�Է�Ϊ��λ
    ���������          As Long
    �ϴν�����ˮ��      As Integer
    �ϴν��׽��        As Long
    �ϴν���ʱ��        As String
    �ս����ۼƽ��      As Long
    �ϴν����ն˺�      As Integer
    ���ȴ�ʱ��          As Integer
    �Ա�                As String
    ����                As Long
    ��������            As String
End Type
Public g�������_�ش�У԰�� As �������
Public gdbl�����޶�_�ش�У԰�� As Double
Const mbln���� As Boolean = False
'--У԰������
Public Declare Function CloseComm Lib "cqcardtsgl.dll" (ByVal icdev As Long) As Integer
Public Declare Function OpenComm Lib "cqcardtsgl.dll" (ByVal CommPort As Long) As Long

Public Declare Function Query_Pos_UserCard Lib "cqcardtsgl.dll" (ByVal icdev As Long, ByRef CardType As Long, ByRef CardSerno As Long, _
             ByVal Cardno As String, ByRef CardGroup As Long, ByVal CardDate As String, ByVal Name As String, _
                 ByVal RegDate As String, ByVal Passport As String, ByRef Account0 As Long, ByRef Account1 As Long, _
                 ByRef Serno As Long, ByRef LastSerno As Integer, ByRef LastAccount As Long, ByVal LastTime As String, _
                 ByRef DayAccount As Long, ByRef LastTermno As Integer, ByVal WAITTIME As Integer) As Integer
                 
Public Declare Function extSys_IsInBlackList Lib "cqcardtsgl.dll" (ByVal ulCardSerno As Long) As Integer
Public Declare Function extSys_WithDraw Lib "cqcardtsgl.dll" (ByVal icdev As Long, ByVal ulCardSerno As Long, ByVal ulPurseNo As Long, ByVal p_TransCode As String, ByVal p_Amount As Long) As Integer
Public Declare Function rf_beep Lib "cqcardtsgl.dll" (ByVal icdev As Long, ByVal WAITTIME As Integer) As Integer

Public Const G_WAIT_TIME = 30
Public Const G_WAIT_TIME1 = 10
Private mblnInit As Boolean '�Ƿ��ʼ��

Public Function ҽ����ʼ��_�ش�У԰��() As Boolean
    '���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
    '���أ���ʼ���ɹ�������true�����򣬷���false
  
    Dim strReg As String
    Dim lngHandle As Long
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    If mblnInit = True Then
        ҽ����ʼ��_�ش�У԰�� = True
        Exit Function
    End If
    
    '���ö˿ں�
    Call GetRegInFor(g����ģ��, "����", "�˿ں�", strReg)


    If Val(strReg) = 0 Then
        gintComport_�ش�У԰�� = 0
    Else
        gintComport_�ش�У԰�� = IIf(Val(strReg) > 99, 1, Val(strReg))
    End If
    gstrSQL = "Select * From ���ղ��� where ������ ='�����޶�' and ����=" & TYPE_�ش�У԰��
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����"
    If Not rsTemp.EOF Then
            gdbl�����޶�_�ش�У԰�� = Val(Nvl(rsTemp!����ֵ))
    Else
        
    End If
    gdbl�����޶�_�ش�У԰�� = IIf(gdbl�����޶�_�ش�У԰�� <= 0, 2000, gdbl�����޶�_�ش�У԰��)
    If mbln���� Then
        ҽ����ʼ��_�ش�У԰�� = True
        Exit Function
    End If
    If g�������_�ش�У԰��.�����豸 <> 0 Then
       lngHandle = CloseComm(g�������_�ش�У԰��.�����豸)
    End If
    lngHandle = OpenComm(gintComport_�ش�У԰��)
    If lngHandle < 0 Then
        ShowMsgbox ("���ڴ�ʧ��,�����豸�Ƿ���������!")
        Exit Function
    End If
    g�������_�ش�У԰��.�����豸 = lngHandle
    mblnInit = True
    
    ҽ����ʼ��_�ش�У԰�� = True
End Function

Public Function ҽ����ֹ_�ش�У԰��() As Boolean
    Dim intReutn As Integer
    mblnInit = False
    '�ر�
    If mbln���� Then
        ҽ����ֹ_�ش�У԰�� = True
        Exit Function
    End If
    intReutn = CloseComm(g�������_�ش�У԰��.�����豸)
    If intReutn <> 0 Then
     '  ShowMsgbox GetErrInfo(CStr(intReutn))
       Exit Function
    End If
  ҽ����ֹ_�ش�У԰�� = True
End Function
Public Function ҽ������_�ش�У԰��(ByVal lng���� As Long, ByVal lngҽ������ As Integer) As Boolean
    ҽ������_�ش�У԰�� = frmSet�ش�У԰��.ShowME(lng����, lngҽ������)
End Function
Public Function ��ݱ�ʶ_�ش�У԰��2(ByVal strCard As String, ByVal strPass As String, Optional lng����ID As Long) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    '/**?
    ��ݱ�ʶ_�ش�У԰��2 = frmIdentify�ش�У԰��.GetPatient(3, lng����ID, True)
End Function
Public Function ��ݱ�ʶ_�ش�У԰��(Optional bytType As Byte, Optional lng����ID As Long) As String
    Dim str��ע As String, RSPATIENT As New ADODB.Recordset
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-���1-סԺ
    '���أ��ջ���Ϣ��
    'ע�⣺1)��Ҫ���ýӿڵ����ʶ���ף�
    '      2)���ʶ������ڴ˺�����ֱ����ʾ������Ϣ��
    '      3)ʶ����ȷ����������Ϣȱ��ĳ������Կո���䣻
    '/**?
    ��ݱ�ʶ_�ش�У԰�� = frmIdentify�ش�У԰��.GetPatient(bytType, lng����ID)
End Function


Public Function �������_�ش�У԰��(ByVal lng����ID As Long, ByVal bytplance As Byte) As Currency
    '����: ���ݲ���idȡ�����
    '����: ����id
    '����: ���ظ����ʻ����
    '����ʧ�����˳�
    'bytplance=10����
    Dim rsTmp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    
    #If gverControl >= 5 Then
        gstrSQL = "select  ������� from ������� where ����id= " & lng����ID & " And ����=1 And ����=" & IIf(bytplance = 10, "1", "2")
    #Else
        gstrSQL = "select  ������� from ������� where ����id= " & lng����ID & " And ����=1 "
    #End If
    zlDatabase.OpenRecordset rsTmp, gstrSQL, "��ȡ���"
    If Not rsTmp.EOF Then
        �������_�ش�У԰�� = Nvl(rsTmp!�������, 0)
    End If
    
    If �������_�ش�У԰�� < (g�������_�ش�У԰��.����Ǯ��1��� + g�������_�ش�У԰��.����Ǯ��2���) / 100 Then
        �������_�ش�У԰�� = (g�������_�ش�У԰��.����Ǯ��1��� + g�������_�ش�У԰��.����Ǯ��2���) / 100
    End If
    If bytplance = 10 Then
        If �������_�ش�У԰�� < gdbl�����޶�_�ش�У԰�� Then
            �������_�ش�У԰�� = gdbl�����޶�_�ش�У԰��
        End If
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
        �������_�ش�У԰�� = 0
End Function

Public Function �����������_�ش�У԰��(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    Dim curTotal As Currency
    Dim dbl��� As Double
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '��ϸ�ֶ�
    '   ����ID,�շ����,�վݷ�Ŀ,���㵥λ,������,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��,ժҪ,�Ƿ���
    With rs��ϸ
        'ȡ�����η������õĽ��ϼ�
        Do While Not .EOF
            '���ж��Ƿ�������ҽ����Ӧ��Ŀ����
            curTotal = curTotal + Round(Nvl(!ʵ�ս��, 0), 2)
            .MoveNext
        Loop
    End With
    
    dbl��� = g�������_�ش�У԰��.����Ǯ��1��� + g�������_�ش�У԰��.����Ǯ��2���
    
    If curTotal * 100 > dbl��� Then
        dbl��� = dbl��� / 100
    Else
        dbl��� = curTotal
    End If
    
    str���㷽ʽ = "�����ʻ�;" & Format(dbl���, "###0.00;-###0.00;0;0") & ";1"  '�������޸�
    �����������_�ش�У԰�� = True
End Function
Public Function �������_�ش�У԰��(lng����ID As Long, cur�����ʻ� As Currency, strSelfNo As String) As Boolean
    Dim lng����ID As Long
    �������_�ش�У԰�� = Set�����������(False, lng����ID, cur�����ʻ�, lng����ID, strSelfNo)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Private Function Set�����������(ByVal bln���� As Boolean, lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long, strSelfNo As String) As Boolean
  '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID��
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    
    Dim curTotal As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim rs��ϸ As New ADODB.Recordset
    
    Dim strInfor As String  '�������ķ��ش�
    Dim strTmp As String
    Dim intҵ�� As Integer
    Dim lng����ID As Long
    Dim strNO As String
    Dim lng��¼���� As Long
    Dim lngTmp As Long

    intҵ�� = IIf(bln����, 1, 0)
     Set����������� = False

    If bln���� Then
        '���¶���
'
'        If GetUserCardInfor = False Then
'            Exit Function
'        End If
        
        '��֤�Ƿ�Ϊ�ò��˵�IC��
        gstrSQL = "Select * From  �����ʻ� where ����id=" & lng����ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵Ŀ���ˮ��"
        If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "�ò����ڱ����ʻ����޼�¼!"
            Exit Function
        End If
        
        'ȷ���˷Ѽ�¼
        '�˷�
          gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
                    " where A.NO=B.NO and A.��¼����=B.��¼����  and A.��¼״̬=2 and B.����ID=[1]"
          Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����˷�", lng����ID)
          If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "�����ڲ��˷��ó�����¼!"
            Exit Function
          Else
            lng����ID = rsTemp("����ID")
          End If
          
    End If
    
    '�򿪱��ν�����ϸ��¼
    gstrSQL = " " & _
        "  Select A.�շ����,a.����ID,sum(nvl(A.���ʽ��,0)) as ʵ�ս��" & _
        "  From ������ü�¼  A" & _
        "  Where A.��¼״̬<>0 and A.����ID=" & IIf(bln����, lng����ID, lng����ID) & " and  Nvl(A.���ӱ�־,0)<>9 " & _
        "  Group by A.�շ����,a.����id" & _
        "  Order by A.�շ����"
        
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡ���ν��ʷ�����ϸ"
    With rs��ϸ
        Do While Not .EOF
            '�����ܶ�,����
            '�ۿ�.
            lng����ID = Nvl(!����ID, 0)
            curTotal = curTotal + Round(Nvl(!ʵ�ս��, 0), 2)
            .MoveNext
        Loop
    End With
    curTotal = IIf(bln����, -1, 1) * cur�����ʻ�
    If bln���� Then
    Else
        If Not �ۿ�_�ش�У԰��("0", Val(Format(curTotal, "####0.00;-####0.00;0.00;0.00")) * 100, True) Then
            
            '�����ж��޷��ع�������ʹ���ܶ���пۿ�.
             Set����������� = False
            Exit Function
        End If
    End If
    '����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,
    '�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '�������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN
    '�����ʻ�֧��_IN,֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    gstrSQL = "zl_���ս����¼_insert(1," & IIf(bln����, lng����ID, lng����ID) & "," & TYPE_�ش�У԰�� & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curTotal & "," & curTotal & ",0,0,0,0," & _
        curTotal & ",NULL,NULL," & curTotal & "," & curTotal & ",Null," & 0 & "," & _
        curTotal & ",0,NULL,null,null" & _
         " )"
    zlDatabase.ExecuteProcedure gstrSQL, "���������շ�����"
    Set����������� = True
End Function
Public Function ����������_�ش�У԰��(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��
    Err = 0
    On Error GoTo errHand:
    ����������_�ش�У԰�� = Set�����������(True, lng����ID, cur�����ʻ�, lng����ID, "")
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function �����ʻ�תԤ��_�ش�У԰��(lngԤ��ID As Long, curMoney As Currency, rsԤ����¼ As ADODB.Recordset) As Boolean
    '���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    
    Dim cur��� As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long
    gstrSQL = "select ����id,nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "У԰��תԤ��", TYPE_�ش�У԰��, g�������_�ش�У԰��.����ˮ��)
    If rsTemp.RecordCount > 0 Then
        If rsTemp("״̬") <> 1 Then
            MsgBox "��ҽ��������δ��Ժ,����ִ�и����ʻ�תԤ�����ף�", vbInformation, gstrSysName
            Exit Function
        End If
        lng����ID = Nvl(rsTemp!����ID, 0)
    End If
    Err = 0
    On Error GoTo errHand:
    If curMoney <> 0 Then
        '�ۿ�
        If �ۿ�_�ش�У԰��(" ", curMoney * 100, False) = False Then
            Exit Function
        End If
    End If
   '����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,
    '�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '�������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN
    '�����ʻ�֧��_IN,֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    gstrSQL = "zl_���ս����¼_insert(3," & lngԤ��ID & "," & TYPE_�ش�У԰�� & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curMoney & "," & curMoney & ",0,0,0,0," & _
        curMoney & ",NULL,NULL," & curMoney & "," & curMoney & ",Null," & 0 & "," & _
        curMoney & ",0,NULL,null,null" & _
         " )"
    zlDatabase.ExecuteProcedure gstrSQL, "�����ʻ�תԤ����"
    
    �����ʻ�תԤ��_�ش�У԰�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    �����ʻ�תԤ��_�ش�У԰�� = False
End Function


Public Function �����ʻ�תԤ������_�ش�У԰��(lngԤ��ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ�����Ҫ�Ӹ����ʻ����ת��Ԥ��������ݼ�¼����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lngԤ��ID-��ǰԤ����¼��ID����Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    '����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,
    '�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '�������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN
    '�����ʻ�֧��_IN,֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    Dim curMoney As Currency
    curMoney = cur�����ʻ�
    gstrSQL = "zl_���ս����¼_insert(3," & lngԤ��ID & "," & TYPE_�ش�У԰�� & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curMoney & "," & curMoney & ",0,0,0,0," & _
        curMoney & ",NULL,NULL," & curMoney & "," & curMoney & ",Null," & 0 & "," & _
        curMoney & ",0,NULL,null,null" & _
         " )"
    zlDatabase.ExecuteProcedure gstrSQL, "�����ʻ�תԤ����"

    �����ʻ�תԤ������_�ش�У԰�� = True
End Function

Public Function סԺ�������_�ش�У԰��(rsExse As Recordset, ByVal lng����ID As Long) As String
    '���ܣ���ȡ�ò���ָ���������ݵĿɱ�����
    '������rsExse-��Ҫ����ķ�����ϸ��¼���ϣ�strSelfNO-ҽ���ţ�strSelfPwd-�������룻
    '      �ֶ�:��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID, _
    '           �շ����,�շ�ϸĿID,�շ�����,��������,���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��, _
    '           �Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
    
    '���أ��ɱ�����:"������ʽ;���;�Ƿ������޸�|...."
    'ע�⣺1)�ú�����Ҫʹ��ģ����㽻�ף���ѯ������ػ�ȡ�������
    '�ӿڷ��صı������ȥ����סԺ�ڼ�����������Ļ��ܽ��󣬲��Ǳ��ε�ʵ�ʱ�����
    'rsExse��¼���е��ֶ��嵥
    '��¼����,��¼״̬,NO,���,����ID,��ҳID,Ӥ����,ҽ����Ŀ����,���մ���ID,
    '�շ����,�շ�ϸĿID,B.���� as �շ�����,X.���� as ��������
    '���,����,����,�۸�,���,ҽ��,�Ǽ�ʱ��,�Ƿ��ϴ�,�Ƿ���,������Ŀ��,ժҪ
  
   Dim str���㷽ʽ  As String
    Dim dbl��� As String
    dbl��� = 0
    If GetUserCardInfor = False Then
        סԺ�������_�ش�У԰�� = ""
         Exit Function
    End If
    Do While Not rsExse.EOF
        dbl��� = dbl��� + Nvl(rsExse!���, 0)
        rsExse.MoveNext
    Loop
    
    Dim dbl��� As Double
    dbl��� = g�������_�ش�У԰��.����Ǯ��1��� '+ g�������_�ش�У԰��.����Ǯ��2���
    
    If dbl��� * 100 > dbl��� Then
        dbl��� = dbl��� / 100
    Else
        dbl��� = dbl���
    End If
    
    str���㷽ʽ = "�����ʻ�;" & Format(dbl���, "###0.00;-###0.00;0;0") & ";1" '���λ��������ʻ�֧��,�������޸�

  
   סԺ�������_�ش�У԰�� = str���㷽ʽ
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function סԺ����_�ش�У԰��(lng����ID As Long, ByVal lng����ID As Long) As Boolean

    Dim lng��ҳID As Long
    Dim rsTemp As New ADODB.Recordset
    '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)
    '������㣨���ص����ݼ�ȥ���ν������ݣ��͵��ڱ��ε���ʵ�������ݣ�
     '���¶���
    If GetUserCardInfor() = False Then
        Exit Function
    End If
     
    On Error GoTo errHand
   
    gstrSQL = " Select B.סԺ���� ��ҳID,to_char(A.��Ժ����,'yyyy') ��Ժ��� " & _
              " From ������ҳ A,������Ϣ B" & _
              " Where B.����ID=[1] And A.��ҳID=B.סԺ���� And A.����ID=B.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ժʱ��", lng����ID)
    lng��ҳID = rsTemp!��ҳID
   סԺ����_�ش�У԰�� = סԺ���㼰����_�ش�У԰��(False, lng����ID, lng����ID, lng����ID, lng��ҳID)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function


Private Function סԺ���㼰����_�ش�У԰��(ByVal bln���� As Boolean, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal ԭ����id As Long, ByVal lng��ҳID As Long) As Boolean

    Dim rs��ϸ As New ADODB.Recordset
    Dim curTotal As Double
    Dim intҵ�� As Integer
    Dim cur�ʻ�֧�� As Double
    Dim dblTmp As Double
    Dim rsTemp As New ADODB.Recordset
    intҵ�� = IIf(bln����, 1, 0)
    
  
    '��ȡ�ʻ�֧����
    gstrSQL = "Select Nvl(A.��Ԥ��,0) �����ʻ� " & _
        " From ����Ԥ����¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And not( a.��¼����  in(11,1)) and  B.����=[1]" & _
        " And A.���㷽ʽ='�����ʻ�' And A.����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ʻ�֧����", TYPE_�ش�У԰��, lng����ID)
    cur�ʻ�֧�� = 0
    If Not rsTemp.EOF Then
        cur�ʻ�֧�� = Nvl(rsTemp!�����ʻ�, 0)
    End If
    
    Err = 0
    On Error GoTo errHand:
    
    'סԺӦ�ñ���֧�������е�סԺ�ȶ�
    gstrSQL = " " & _
        "  Select A.�շ����,sum(nvl(A.���ʽ��,0)) as ʵ�ս��" & _
        "  From סԺ���ü�¼  A" & _
        "  Where a.��¼״̬<>0 and A.����ID=" & lng����ID & " and  Nvl(A.���ӱ�־,0)<>9 " & _
        "  Group by A.�շ����" & _
        "  Order by A.�շ����"
        
    zlDatabase.OpenRecordset rs��ϸ, gstrSQL, "��ȡסԺ������ϸ"
    With rs��ϸ
        dblTmp = 0
        Do While Not .EOF
            '�����ܶ�,����
            '�ۿ�
'
'            dblTmp = dblTmp + NVL(!ʵ�ս��)
'            If dblTmp > cur�ʻ�֧�� Then
'                'ȷ���Ƿ��Ѿ����������ʻ�
'                If cur�ʻ�֧�� - curTotal > 0 Then
'                    If Not �ۿ�_�ش�У԰��(NVL(!�շ����, "0"), (cur�ʻ�֧�� - curTotal) * 100, True) Then
'                        '
'                    End If
'                End If
'                curTotal = cur�ʻ�֧��
'                Exit Do
'            Else
                curTotal = curTotal + Round(Nvl(!ʵ�ս��, 0), 2)
''            End If
            .MoveNext
        Loop
    End With
    curTotal = cur�ʻ�֧��
    If bln���� = False Then
            If Not �ۿ�_�ش�У԰��("0", Val(Format(cur�ʻ�֧��, "####0.00;-####0.00;0.00;0.00")) * 100, True) Then
                '
                סԺ���㼰����_�ش�У԰�� = False
                Exit Function
            End If
    End If
    '��������¼
     '����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,
    '�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '�������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN
    '�����ʻ�֧��_IN,֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�ش�У԰�� & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curTotal & "," & curTotal & "," & lng��ҳID & ",0,0,0," & _
        curTotal & ",NULL,NULL," & curTotal & "," & curTotal & ",Null," & 0 & "," & _
        cur�ʻ�֧�� & ",0,NULL,null,null" & _
         " )"
           
        zlDatabase.ExecuteProcedure gstrSQL, "����סԺ�����շ�����"
        סԺ���㼰����_�ش�У԰�� = True
        Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_�ش�У԰��(lng����ID As Long) As Boolean
    Dim lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    
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
    gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
              " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    lng����ID = rsTemp("ID") '�������ݵ�ID
    '���¶���
    If GetUserCardInfor() = False Then
        Exit Function
    End If

    'Ϊ�˽���ʱд���Ľ����������ٴη��ʼ�¼
    gstrSQL = "Select * " & _
              "  From ���ս����¼ Where ����=2 and ��¼ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "�ڱ��ս����¼���޸ý����¼!"
        Exit Function
    End If
    lng����ID = Nvl(rsTemp!����ID, 0)
    lng��ҳID = Nvl(rsTemp!��ҳID, 0)
        
        
    
    '��֤�Ƿ�Ϊ�ò��˵�IC��
    gstrSQL = "Select * From  �����ʻ� where ����id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���˵�ҽ����"
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "�ò����ڱ����ʻ����޼�¼!"
        Exit Function
    End If
    
    If g�������_�ش�У԰��.���� <> Nvl(rsTemp!����) Then
        Err.Raise 9000, gstrSysName, "�ò��˵�IC���������,�����ǲ����������˵�IC��!"
        Exit Function
    End If
    
    '���ó�������ӿ�
    סԺ�������_�ش�У԰�� = סԺ���㼰����_�ش�У԰��(True, lng����ID, lng����ID, lng����ID, lng��ҳID)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Public Function ��Ժ�Ǽ�_�ش�У԰��(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    
    On Error GoTo errHand
    
    '��ȡ���˵���ر�����Ϣ

    gstrSQL = "select * From �����ʻ� where  ����=" & TYPE_�ش�У԰�� & "  and ����id=" & lng����ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��Ժ��ȡ�����ʻ���Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ����ʻ����޸ò��˵ı�����Ϣ!"
        Exit Function
    End If
    
    '�ı䲡��״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ش�У԰�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    ��Ժ�Ǽ�_�ش�У԰�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function ��Ժ�Ǽǳ���_�ش�У԰��(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
                'ȡ��Ժ�Ǽ���֤�����ص�˳���
                
    Dim str��Ժ����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    
    gstrSQL = " Select Count(*) Records From סԺ���ü�¼ " & _
              " Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ժ���", lng����ID, lng��ҳID)
    
    If rsTemp!Records <> 0 Then
        MsgBox "�Ѿ����ڷ��ü�¼���������������Ժ�Ǽǣ�", vbInformation, gstrSysName
        Exit Function
    End If

    
    On Error GoTo errHand
    
    '��ȡ���˵���ر�����Ϣ

    gstrSQL = "select * From �����ʻ� where  ����=" & TYPE_�ش�У԰�� & "  and ����id=" & lng����ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "������Ժ��ȡ�����ʻ���Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�ڱ����ʻ����޸ò��˵ı�����Ϣ!"
        Exit Function
    End If
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ش�У԰�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_�ش�У԰�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function ��Ժ�Ǽ�_�ش�У԰��(lng����ID As Long, lng��ҳID As Long) As Boolean
    '����HIS��Ժ
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ش�У԰�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    ��Ժ�Ǽ�_�ش�У԰�� = True
End Function
Public Function ��Ժ�Ǽǳ���_�ش�У԰��(lng����ID As Long, lng��ҳID As Long) As Boolean
    '����δ����õĲ��˲�������HIS��Ժ��������Ϊ�Ѱ���ҽ����Ժ���������ٰ���HIS��Ժ
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        MsgBox "ҽ���ѳ�Ժ�Ĳ��˲���������Ժ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�ش�У԰�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_�ش�У԰�� = True
End Function


Public Function �ҺŽ���_�ش�У԰��(ByVal lng����ID As Long) As Boolean
    Dim curTotal As Currency '�ϴ������ܶ�
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long
    
    '�ȵ�����Ԥ����,��Ҫ��ȡ�����ʻ�֧����,�ٵ��������
    '���ղ����б�����������ʻ�֧����������ĿID���ҺŽ���ʱ�жϣ����δ���ã�����ȫ�Ը��������ϴ���������ϴ���һ����ϸ
    
    On Error GoTo errHand
    
  gstrSQL = " " & _
        "  Select A.�շ����,A.����id,sum(nvl(A.���ʽ��,0)) as ʵ�ս��" & _
        "  From ������ü�¼  A" & _
        "  Where A.��¼״̬<>0 and  A.����ID=" & lng����ID & " and  Nvl(A.���ӱ�־,0)<>9 " & _
        "  Group by A.�շ����,A.����id" & _
        "  Order by A.�շ����"
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���ν��ʷ�����ϸ"
    
    With rsTemp
        If .RecordCount = 0 Then
            ShowMsgbox "���κη��ü�¼����!"
            Exit Function
        End If
        Do While Not .EOF
            '�����ܶ�,����
            '�ۿ�.
            lng����ID = Nvl(!����ID, 0)
            curTotal = curTotal + Round(Nvl(!ʵ�ս��, 0), 2)
            .MoveNext
        Loop
    End With
    If Not �ۿ�_�ش�У԰��("0", Val(Format(curTotal, "####0.00;-####0.00;0.00;0.00")) * 100, True) Then
        '
        �ҺŽ���_�ش�У԰�� = False
        Exit Function
    End If
 
   '��������¼
     '����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,
    '�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '�������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN
    '�����ʻ�֧��_IN,֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�ش�У԰�� & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curTotal & "," & curTotal & ",0,0,0,0," & _
        curTotal & ",NULL,NULL," & curTotal & "," & curTotal & ",Null," & 0 & "," & _
        curTotal & ",0,NULL,null,null" & _
         " )"
    zlDatabase.ExecuteProcedure gstrSQL, "���뱣�ռ�¼"
     �ҺŽ���_�ش�У԰�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function �Һų���_�ش�У԰��(ByVal lng����ID As Long) As Boolean
   Dim curTotal As Currency '�ϴ������ܶ�
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long
    
    '�ȵ�����Ԥ����,��Ҫ��ȡ�����ʻ�֧����,�ٵ��������
    '���ղ����б�����������ʻ�֧����������ĿID���ҺŽ���ʱ�жϣ����δ���ã�����ȫ�Ը��������ϴ���������ϴ���һ����ϸ
    
    On Error GoTo errHand
    
  gstrSQL = " " & _
        "  Select A.�շ����,A.����id,sum(nvl(A.���ʽ��,0)) as ʵ�ս��" & _
        "  From ������ü�¼  A" & _
        "  Where A.��¼״̬<>0 and  A.����ID=" & lng����ID & " and  Nvl(A.���ӱ�־,0)<>9 " & _
        "  Group by A.�շ����,A.����id" & _
        "  Order by A.�շ����"
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���ν��ʷ�����ϸ"
    
    With rsTemp
        If .RecordCount = 0 Then
            ShowMsgbox "���κη��ü�¼����!"
            Exit Function
        End If
        Do While Not .EOF
            '�����ܶ�,����
            '�ۿ�.
            lng����ID = Nvl(!����ID, 0)
            curTotal = curTotal + Round(Nvl(!ʵ�ս��, 0), 2)
            .MoveNext
        Loop
    End With
  
 
   '��������¼
     '����_IN,��¼ID_IN,����_IN,����ID_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,
    '�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
    '�������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN
    '�����ʻ�֧��_IN,֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�ش�У԰�� & "," & lng����ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curTotal & "," & curTotal & ",0,0,0,0," & _
        curTotal & ",NULL,NULL," & curTotal & "," & curTotal & ",Null," & 0 & "," & _
        curTotal & ",0,NULL,null,null" & _
         " )"
      zlDatabase.ExecuteProcedure gstrSQL, "���뱣�ռ�¼"
         
    �Һų���_�ش�У԰�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function GetUserCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------
    '����:��ȡ����Ϣ(�û�����ѯ)
    '����:��ȡ�ɹ�,����True,���򷵻�False
    '---------------------------------------------------------------------------------------------------
    Dim intReturn As Integer
    GetUserCardInfor = False
    '��

    With g�������_�ش�У԰��
        .���� = Space(17)
        .ע������ = Space(7)
        .���� = Space(21)
        .����Ч�� = Space(7)
        .���֤�� = Space(19)
        .�ϴν���ʱ�� = Space(13)
        
        '����
        If mbln���� Then
            intReturn = testdATA
        Else
             intReturn = Query_Pos_UserCard(.�����豸, .������, .����ˮ��, .����, .�����ѷ���, .����Ч��, _
                    .����, .ע������, .���֤��, .����Ǯ��1���, .����Ǯ��2���, .���������, .�ϴν�����ˮ��, _
                    .�ϴν��׽��, .�ϴν���ʱ��, .�ս����ۼƽ��, .�ϴν����ն˺�, .���ȴ�ʱ��)
                    
        End If
        If intReturn <> 0 Then
            '��������
            Call rf_beep(g�������_�ش�У԰��.�����豸, G_WAIT_TIME)
            ShowMsgbox GetErrInfo(CStr(intReturn), TYPE_�ش�У԰��)
            Exit Function
        End If
        
        .���� = Trim(Replace(.����, Chr(0), "", 1))
        .ע������ = Trim(Replace(.ע������, Chr(0), "", 1))
        .���� = Trim(Replace(.����, Chr(0), "", 1))
        .����Ч�� = Trim(Replace(.����Ч��, Chr(0), "", 1))
        .���֤�� = Trim(Replace(.���֤��, Chr(0), "", 1))
        .�ϴν���ʱ�� = Trim(Replace(.�ϴν���ʱ��, Chr(0), "", 1))
        
        '����������Ϣ
        Dim int�Ա� As Integer
        int�Ա� = Val(IIf(Len(Trim(.���֤��)) = 18, Mid(Trim(.���֤��), 17, 1), Right(Trim(.���֤��), 1))) Mod 2
        '�������֤ȡ����Ӧ���Ա�
        .�Ա� = IIf(int�Ա� = 0, "Ů", "��")
        .�������� = zlCommFun.GetIDCardDate(Trim(.���֤��))
        '��������
        If IsDate(.��������) And .�������� <> "" Then
            .���� = Abs(Int((zlDatabase.Currentdate - CDate(.��������)) / 365))
        Else
            .���� = 0
        End If
        '�жϸÿ��Ƿ����
        If "20" & Trim(.����Ч��) < Format(zlDatabase.Currentdate, "yyyymmdd") Then
            Call rf_beep(g�������_�ش�У԰��.�����豸, G_WAIT_TIME)
            ShowMsgbox "�ÿ��Ѿ�����,��������(��Ч��Ϊ:20" & .����Ч�� & ")!"
            Exit Function
        End If
        If Not mbln���� Then
            '�жϸÿ��Ƿ���Ĭ������
            intReturn = extSys_IsInBlackList(.����ˮ��)
            If intReturn <> 0 Then
                Call rf_beep(g�������_�ش�У԰��.�����豸, G_WAIT_TIME)
               ShowMsgbox "�ÿ�Ϊ��������,����ʹ�øÿ�!"
               Exit Function
            End If
            Call rf_beep(g�������_�ش�У԰��.�����豸, G_WAIT_TIME1)
        End If
    End With
    GetUserCardInfor = True
End Function

Public Function �ۿ�_�ش�У԰��(ByVal str���ô��� As String, ByVal lng��� As Long, Optional bln���� As Boolean) As Boolean
    '����:�Է��ý��пۿ�
    Dim intRetun As Integer
    Dim lngTmp As Long
    �ۿ�_�ش�У԰�� = False
    Err = 0
    On Error GoTo errHand:
    '�����շ���Ҫ�ȿ�1Ǯ��,����۳�0
    lngTmp = lng���
    If mbln���� Then
        �ۿ�_�ش�У԰�� = True
        Exit Function
    End If
    If gdbl�����޶�_�ش�У԰�� < lng��� / 100 Then
        ShowMsgbox "���������޶�,���ܽ���!"
        �ۿ�_�ش�У԰�� = False
        Exit Function
    End If
    If bln���� Then
        '�ȿ۵�Ǯ��2��Ǯ
        If g�������_�ش�У԰��.����Ǯ��2��� <> 0 Then
            If lng��� < g�������_�ش�У԰��.����Ǯ��2��� Then
                intRetun = extSys_WithDraw(g�������_�ش�У԰��.�����豸, g�������_�ش�У԰��.����ˮ��, 1, "200", lng���)
                lngTmp = 0
            Else
                intRetun = extSys_WithDraw(g�������_�ش�У԰��.�����豸, g�������_�ش�У԰��.����ˮ��, 1, "200", g�������_�ش�У԰��.����Ǯ��2���)
                lngTmp = lng��� - g�������_�ش�У԰��.����Ǯ��2���
            End If
        End If
        If g�������_�ش�У԰��.����Ǯ��1��� <> 0 And lngTmp > 0 Then
            '�ٿ۵���Ǯ��1��Ǯ
            intRetun = extSys_WithDraw(g�������_�ش�У԰��.�����豸, g�������_�ش�У԰��.����ˮ��, 0, "200", lngTmp)
        End If
        
'        If lng��� <= g�������_�ش�У԰��.����Ǯ��2��� And g�������_�ش�У԰��.����Ǯ��2��� <> 0 Then
'            intRetun = extSys_WithDraw(g�������_�ش�У԰��.�����豸, g�������_�ش�У԰��.����ˮ��, 1, "200", lngTmp)
'        Else
'            intRetun = extSys_WithDraw(g�������_�ش�У԰��.�����豸, g�������_�ش�У԰��.����ˮ��, 0, "200", lngTmp)
'        End If
    Else
        intRetun = extSys_WithDraw(g�������_�ش�У԰��.�����豸, g�������_�ش�У԰��.����ˮ��, 0, "200", lngTmp)
    End If
    
    If intRetun <> 0 Then
        Call rf_beep(g�������_�ش�У԰��.�����豸, G_WAIT_TIME)
        If intRetun = -14 Then
            ShowMsgbox "����Ǯ�����(" & g�������_�ش�У԰��.����Ǯ��1��� / 100 & ")����,���ܽ��н���!"
        Else
            MsgBox GetErrInfo(CStr(intRetun), TYPE_�ش�У԰��)
        End If
        Exit Function
    Else
        Call rf_beep(g�������_�ش�У԰��.�����豸, G_WAIT_TIME1)
    End If
    �ۿ�_�ش�У԰�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function �ۿ�_���2_�ش�У԰��() As Boolean
    Dim intRetun As Integer
    '�������Ǯ��2��Ǯ.
    �ۿ�_���2_�ش�У԰�� = False
    If GetUserCardInfor = False Then Exit Function
    intRetun = extSys_WithDraw(g�������_�ش�У԰��.�����豸, g�������_�ش�У԰��.����ˮ��, 1, "200", g�������_�ش�У԰��.����Ǯ��2���)
    If intRetun <> 0 Then
        Call rf_beep(g�������_�ش�У԰��.�����豸, G_WAIT_TIME)
        MsgBox GetErrInfo(CStr(intRetun), TYPE_�ش�У԰��)
        Exit Function
    Else
        Call rf_beep(g�������_�ش�У԰��.�����豸, G_WAIT_TIME1)
    End If
    �ۿ�_���2_�ش�У԰�� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function testdATA() As Long
    '��������
    With g�������_�ش�У԰��
        .�����豸 = 1234
        .������ = 3
        .����ˮ�� = 1
        .���� = "13424"
        .�����ѷ��� = 1
        .����Ч�� = "041236"
        .���� = "����"
        .ע������ = "20020303"
        .���֤�� = "510221197404282859"
        .����Ǯ��1��� = 3000
        .����Ǯ��2��� = 4000
        .��������� = 1
        .�ϴν�����ˮ�� = 2
        .�ϴν��׽�� = 1000
        .�ϴν���ʱ�� = "20040402"
        .�ս����ۼƽ�� = 1090
        .�ϴν����ն˺� = 1
        .���ȴ�ʱ�� = 200
    End With
    testdATA = 0
End Function






