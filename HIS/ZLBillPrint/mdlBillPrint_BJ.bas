Attribute VB_Name = "mdlBillPrint_BJ"
Option Explicit


Public Function Init() As Boolean
'���ܣ����е�����Ʊ�ݴ�ӡ�ӿڵĳ�ʼ�����¼�ȵ���
'���أ�ִ�гɹ�/ʧ��



    '�ο���ͨ��ע����д��ʱ��¼����
    'Call SaveSetting("ZLSOFT", "����ȫ��\Ʊ�ݴ�ӡ", "��ǰƱ�ݺ�", 'XXXX')
    'Call GetSetting("ZLSOFT", "����ȫ��\Ʊ�ݴ�ӡ", "��ǰƱ�ݺ�", "")
    
    Init = True
End Function

Public Function Term() As Boolean
'���ܣ���ɵ�����Ʊ�ݴ�ӡ�ӿڵ���Դ�ͷš��Ͽ����ӵȵ���
'���أ�ִ�гɹ�/ʧ��
    
    
    Term = True
End Function


Public Function SYSConfigure() As Boolean
'���ܣ���������,��HIS"ģ���������"(�ļ�/��������)�е��ã����ڱ��ӿ�����ɵ�����Ʊ�ݴ�ӡ�ӿڵĲ������á����ø��ĵȵ���
'���أ�ִ�гɹ�/ʧ��
    
    
    SYSConfigure = True
End Function

Public Function DiscardBill(ByVal lng����ID As Long, ByVal lngƱ�� As Long, ByVal strƱ��ǰ׺ As String, _
    ByVal str��ʼƱ�� As String, ByVal str����Ʊ�� As String, ByVal DateAdd As Date, ByVal str������ As String) As Boolean
'���ܣ�Ʊ�ݱ���
'���أ�ִ�гɹ�/ʧ��

    DiscardBill = True
End Function

Public Function PrintBillOut(ByVal strNOs As String) As Boolean
'���ܣ������շ�Ʊ�ݴ�ӡ
'������strNOs=�����շѣ��Զ��ŷָ��Ĵ����ŵĶ�����ݺ�(һ�δ�ӡ���Ż���ŵ���):'F0000001','F0000002',...
'���أ�ִ�гɹ�/ʧ��
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH
    
   '�ο�����ȡ���ݺ���ط�������
   'ʹ��f_Str2list(Ϊ��ʹ�ð󶨱���,zltools���ṩ�Ľ��ַ���ת��Ϊ��ʱ�ڴ��ĺ���),
   '��Ҫ��SQL����е�һ��Select�ؼ��ֺ���롰/*+ Rule*/����ʾ����ΪCbo����ʱ�ڴ��û��ͳ�����ݣ�����ᵼ�´��ȫ��ɨ��
'   strSQL = "Select/*+ Rule*/ �վݷ�Ŀ as ��Ʊ��Ŀ,Sum(ʵ�ս��) as ���" & _
'            " From ������ü�¼" & _
'            " Where ��¼����=1 and NO In (Select * From Table(f_Str2list([1])))" & _
'            " Group By �վݷ�Ŀ"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ӡ����", Replace(strNOs, "'", ""))
'    If rstmp.RecordCount = 0 Then
'        Exit Function
'    End If

    

    PrintBillOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function PrintBillIn(ByVal lngBalanceId As Long) As Boolean
'���ܣ�סԺ����Ʊ�ݴ�ӡ
'������lngBalanceId=���ʵ�ID
'���أ�ִ�гɹ�/ʧ��
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH
    
'   strSQL = "Select/*+ Rule*/ �վݷ�Ŀ as ��Ʊ��Ŀ,Sum(���ʽ��) as ���" & _
'            " From סԺ���ü�¼" & _
'            " Where ����id=[1]" & _
'            " Group By �վݷ�Ŀ"
'   Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ӡ����", lngBalanceId)

'   strSQL = "Select L.ʵ��Ʊ��,I.סԺ��,L.����Ա����" & _
'            " From ���˽��ʼ�¼ L,������Ϣ I" & _
'            " Where L.����id=I.����id And L.id=[1]"


    PrintBillIn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function RePrintBillOut(ByVal strNOs As String, ByVal strInvoice As String) As Boolean
'���ܣ��ش������շ�Ʊ��
'������strNOs=�����շѣ��Զ��ŷָ��Ĵ����ŵĶ�����ݺ�(һ�δ�ӡ���Ż���ŵ���):'F0000001','F0000002',...
'      strInvoice=�����ش�ʹ�õ���ʼƱ�ݺ�
'���أ�ִ�гɹ�/ʧ��
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH
    
'    strSQL = "Select ����,ʹ����" & _
'                 " From Ʊ��ʹ����ϸ" & _
'                 " Where Id = (" & _
'                 "       Select Max(Id)" & _
'                 "       From Ʊ��ʹ����ϸ" & _
'                 "       Where ���� = 2 And ��ӡid In (" & _
'                 "             Select Id From Ʊ�ݴ�ӡ���� Where ��������=1 And No In (Select * From Table(f_Str2list([1])))))"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡƱ�ݺ�", replace(strNOs,"'",""))

    RePrintBillOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function RePrintBillIn(ByVal lngBalanceId As Long, ByVal strInvoice As String) As Boolean
'���ܣ��ش�סԺ����Ʊ��
'������lngBalanceId=���ʵ�ID
'      strInvoice=�����ش�ʹ�õ���ʼƱ�ݺ�
'���أ�ִ�гɹ�/ʧ��
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH
    
'    strSQL = "Select ����,ʹ����" & _
'                 " From Ʊ��ʹ����ϸ" & _
'                 " Where Id In (" & _
'                 "       Select Max(Id)" & _
'                 "       From Ʊ��ʹ����ϸ" & _
'                 "       Where ���� = 2 And ��ӡid In (" & _
'                 "             Select Id From Ʊ�ݴ�ӡ���� Where ��������=3 And No In (" & _
'                 "             Select No From ���˽��ʼ�¼ Where ID=[1])))"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡƱ�ݺ�", lngBalanceId)
                 
    RePrintBillIn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function EraseBillOut(ByVal strNOs As String) As Boolean
'���ܣ����������շ�Ʊ��
'������strNOs=�����շѣ��Զ��ŷָ��Ĵ����ŵĶ�����ݺ�(һ�δ�ӡ���Ż���ŵ���):'F0000001','F0000002',...
'���أ�ִ�гɹ�/ʧ��
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH

'   ���ڲ����˷��ٴ�ӡ���൥���޸ĵ����,���ܸ����ϵ��������·���Ʊ�ݣ�����Ҫ��max(id)
'    strSQL = "Select/*+ Rule*/ ����,ʹ����" & _
'                 " From Ʊ��ʹ����ϸ" & _
'                 " Where Id = (" & _
'                 "       Select Max(Id)" & _
'                 "       From Ʊ��ʹ����ϸ" & _
'                 "       Where ���� = 2 And ��ӡid In (" & _
'                 "             Select Id From Ʊ�ݴ�ӡ���� Where ��������=1 And No In (Select * From Table(f_Str2list([1]))))"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡƱ�ݺ�", replace(strNOs,"'",""))
    
    EraseBillOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function EraseBillIn(ByVal lngBalanceId As Long) As Boolean
'���ܣ�����סԺ����Ʊ��
'������lngBalanceId=���ʵ�ID
'���أ�ִ�гɹ�/ʧ��
    Dim strSQL As String, rstmp As ADODB.Recordset
    On Error GoTo errH
    
'    strSQL = "Select L.ʵ��Ʊ��,L.����Ա����" & _
'             " From ���˽��ʼ�¼ L Where L.id=[1]"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡƱ�ݺ�", lngBalanceId)
                
    EraseBillIn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
