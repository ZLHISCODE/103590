Attribute VB_Name = "Mdl��������"
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;99-���н������Ӹ��Ӳ���(���°�)

'Modified by ���� 20031218 ���������� ʡ��ҽ��Ҫ�ֿ�����˽�����TYPE_����������ΪintInsure��ע���ύ����

'----------------------------------------------�ֶγ���----------------------------------------------
'������˽�б�������
Private glng����ID As Long                                  '������Ժ�Ǽ�ʱʹ��
Private bln������Ժ As Boolean
Private mstrFields As String
Private mstrValues As String
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����
Private Const mintStyle As Integer = 1                       '1-������ı�;2-�������Ļ

Public mgstrPatientInfo As String                           '������Ϣ��
Public Const mstrPath_�������� As String = "C:\HIS"
Public Const mstrSearch_�������� As String = "��ӡ.avi"
Public Const mstrReply_�������� As String = "Reply.txt"
Public Const mstrRequest_�������� As String = "Request.txt"
Public Const mstrTemp_�������� As String = "Temp.txt"

Enum ��������
    ��ֵ�� = 1
    �ַ��� = 2
    ������ = 3
    ������ = 4
    ʱ���� = 5
End Enum
Enum ������ʽ
    ��¼ = 1
    ��Ժ = 2
    �Һ� = 3
    �շ� = 4
    ���� = 5
    ��Ժ = 6
    ��֤ = 7
End Enum
Enum ����Ŀ��
    ���� = 1
    ���� = 2
    ˢ�� = 3
    ��ϸ = 4
    ��ѯ = 5    'ר���ڵ�¼��ѯ
    Ԥ���� = 6
End Enum

Public mrsIniItems As New ADODB.Recordset                   'ini�����ļ��е���Ŀ
Private mrsIniSection As New ADODB.Recordset                 'ini�����ļ��еĽ�
Private mrsDetail As New ADODB.Recordset
Private curTotalMoney As Currency                           '����Ԥ����ʱ�ϴ��Ľ���ܶ�

'------------------------------------ҽ�����溯��------------------------------------
Public Function ҽ����ʼ��_��������() As Boolean
    Static gbln��ʼ�� As Boolean
    Dim rsTemp As New ADODB.Recordset
'���ܣ�����Ӧ�ò����Ѿ�������ORacle���ӣ�ͬʱ����������Ϣ������ҽ�������������ӡ�
'���أ���ʼ���ɹ�������true�����򣬷���false
    
    ҽ����ʼ��_�������� = False
    If gbln��ʼ�� Then
        ҽ����ʼ��_�������� = True
        Exit Function
    End If
    
    If Not InitStruc Then Exit Function
    
    ҽ����ʼ��_�������� = True 'frm�ȴ���Ӧ.ShowME(������ʽ.��¼, ����Ŀ��.����)
    If ҽ����ʼ��_�������� Then gbln��ʼ�� = True
End Function

'Modified by ���� 20031218 ����������
Public Function ҽ������_��������(ByVal lng���� As Long) As Boolean
'���ܣ� �÷������ڹ����Ӧ�ò���������������ҽ�����ݷ����������Ӵ�
'���أ��ӿ����óɹ�������true�����򣬷���false
    Dim strConn As String
    
    ҽ������_�������� = FrmSet����.ShowSet(lng����)
End Function

Public Function ҽ����ֹ_��������() As Boolean
    'ҽ��ȡ����¼
    'If Not frm�ȴ���Ӧ.ShowME(������ʽ.��¼, ����Ŀ��.����) Then Exit Function
    
    '����ڲ���¼��
    Set mrsIniItems = Nothing
    Set mrsIniSection = Nothing
    
    ҽ����ֹ_�������� = True
End Function

Public Function ��ݱ�ʶ_��������(ByVal intinsure As Integer, ByVal ������ʽ_IN As Integer, Optional ByVal lng����ID As Long = 0) As String
    Dim str˳��� As String, str���� As String, arrReturn
    ��ݱ�ʶ_�������� = ""
    
    '����ǵ��������ʻ������н�����֤����֤�꼴�˳����ṩ���û��޸ı��ղ��ֵ�һ��;����
    If ������ʽ_IN = ������ʽ.��֤ Then
        If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ_IN, ����Ŀ��.ˢ��) Then Exit Function
        ��ݱ�ʶ_�������� = "С��"  '����ǿ��ˢ��
        Exit Function
    End If
    
    '�������Ĳ���ID��Ϊ�գ����ʾ���еĲ����ǲ�����Ժ�Ǽ�
    bln������Ժ = (lng����ID <> 0)
    If bln������Ժ Then glng����ID = lng����ID
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ_IN, ����Ŀ��.ˢ��) Then Exit Function
    ��ݱ�ʶ_�������� = mgstrPatientInfo
    
    '   ��ͬһ���˿��ܴ��ڹҶ�����ҵĺű����ԣ������շ�ˢ��ʱ���᷵��һ��ʱ���ڣ�������һ������죩��
    '������˹ҺŵĶ����ˮ�ż�����������ƣ��ԷֺŸ�������Ҫ����Աѡ������һ������Ϊ����ʹ�õ���ˮ��
    Call Record_Locate(mrsIniItems, "����,Mzlsh0")
    str˳��� = Nvl(mrsIniItems!ֵ, "")
    If InStr(1, str˳���, ";") <> 0 Then
        '���ڶ���Һſ��Ҽ��Һ���ˮ��
        Call Record_Locate(mrsIniItems, "����,Ghksmc")
        str���� = Nvl(mrsIniItems!ֵ, "")
        arrReturn = Split(frmShowList.ShowME(str˳��� & "||" & str����), ";")
        str˳��� = arrReturn(0)
        str���� = arrReturn(1)
        Call UpdateData("Mzlsh0", str˳���)
        Call UpdateData("Ghksmc", str����)
    Else
        Call Record_Locate(mrsIniItems, "����,Ghksmc")
        str���� = Nvl(mrsIniItems!ֵ, "")
    End If
    ��ݱ�ʶ_�������� = ��ݱ�ʶ_�������� & ";" & str����
End Function

Public Function �������_��������(ByVal lng����ID As Long, ByVal intinsure As Integer) As Currency
'����: ��ȡ�α����˸����ʻ����
'����: ���ظ����ʻ����
    Dim rsTemp As New ADODB.Recordset

    
    gstrSQL = "select A.�ʻ���� from �����ʻ� A where A.����ID=[1] and A.����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ����", lng����ID, intinsure)
    
    If rsTemp.EOF Then
        �������_�������� = 0
    Else
        �������_�������� = IIf(IsNull(rsTemp("�ʻ����")), 0, rsTemp("�ʻ����"))
    End If

End Function

Public Function ��Ժ�Ǽ�_��������(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    Dim strObj As String, strסԺ�� As String, blnExist As Boolean
    Dim strLine As TextStream, FileSys As New FileSystemObject
    Dim str˳��� As String
    '�ȷ���ˢ�������ٷ�����Ժ���󣬵õ�Ӧ���ļ���������Ժ�����
    
    On Error GoTo errHand
    ��Ժ�Ǽ�_�������� = False
    
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.����, lng����ID) Then Exit Function
    
    '��Ժ�Ǽ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
    '������Ժ��ˮ��
    Call Record_Locate(mrsIniItems, "����,Zylsh0")
    str˳��� = Nvl(mrsIniItems!ֵ, "")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure & ",'˳���','''" & str˳��� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ��ˮ��")
    
    ��Ժ�Ǽ�_�������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽ�_��������(ByVal intinsure As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal bln������Ժ As Boolean = False) As Boolean
    '�Ƚ�����Ժ:
    '   �����������--ҽ����Ժ��HIS��Ժ
    '   δ��������  --ҽ��������Ժ��HIS��Ժ
    '�ȳ�Ժ�����:
    '   �����������--HIS��Ժ
    '   δ��������  --��ҽ��������Ժ���޷��÷���������������ҽ���Ľ���ӿڣ�����ֱ�ӳ�Ժ��
    
    'bln������Ժ=TRUE:��ʾ�����Ժ�����Գ�����Ժ�ķ�ʽ�ڵ��ã�����ҽ����֧��HIS���泷����Ժ��
    '���ۺ��ַ�ʽ����HIS���涼��ӳΪ��Ժ����ҽ���˷�ӳΪ������Ժ���Ժ
    On Error GoTo errHand
    ��Ժ�Ǽ�_�������� = False
    
    If bln������Ժ Then
        MsgBox "��֧�ָù��ܣ���Ϊ���˰����Ժ������", vbInformation, gstrSysName
        Exit Function
    End If
    
    If ���ڷ��ü�¼(lng����ID, lng��ҳID) Then
        If ��������(lng����ID, lng��ҳID) Then
            If ����ģʽ(intinsure) = 0 Then
                '�Ƚ�����Ժ��˵���ǰ���ҽ����Ժ����
                If ����δ�����(lng����ID, lng��ҳID) Then
                    MsgBox "�ò��˻�����δ����ã����ܳ�Ժ��", vbInformation, gstrSysName
                    Exit Function
                End If
                If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.ˢ��) Then Exit Function
                If lng����ID <> ��ȡ����ID(intinsure) Then
                    MsgBox "������Ϣ������", vbInformation, gstrSysName
                    Exit Function
                End If
                If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.����, lng����ID) Then Exit Function
                MsgBox "�ò�����ҽ�����ĳɹ������Ժ������", vbInformation, gstrSysName
            Else
                '�ȳ�Ժ����㣬�������δ����ã����HIS��Ժ�������ȵ�ҽ����Ժ���ٵ�HIS��Ժ
                If Not ����δ�����(lng����ID, lng��ҳID) Then
                    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.ˢ��) Then Exit Function
                    If lng����ID <> ��ȡ����ID(intinsure) Then
                        MsgBox "������Ϣ������", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.����, lng����ID) Then Exit Function
                    MsgBox "�ò�����ҽ�����ĳɹ������Ժ������", vbInformation, gstrSysName
                End If
            End If
        Else
            If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.����, lng����ID) Then Exit Function
        End If
    Else
        If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.����, lng����ID) Then Exit Function
    End If
    
    '��Ժ�Ǽ�
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
    
    ��Ժ�Ǽ�_�������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ��Ժ�Ǽǳ���_��������(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    '�Ƚ�����Ժ:
    '   �����������--ҽ��������Ժ
    '   δ��������  --��ҽ����Ժ
    '�ȳ�Ժ�����:
    '   ����δ�����--HIS��Ժ
    '   ������δ�����
    '       �����������--ҽ��������Ժ��HIS��Ժ
    '       δ��������  --ҽ����Ժ
    '   ����ʱ��ҽ�����㡢ҽ����Ժ���޷���ʱ��ҽ����������㣩
    On Error GoTo errHand
    Dim lng��ҳID As Long
    Dim rsסԺ���� As New ADODB.Recordset
    ��Ժ�Ǽǳ���_�������� = False
    
    'ȡ����ҳID
    gstrSQL = "Select Nvl(סԺ����,0) ��ҳID From ������Ϣ Where ����ID=[1]"
    Set rsסԺ���� = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID", lng����ID)
    lng��ҳID = rsסԺ����!��ҳID
    
    If ���ڷ��ü�¼(lng����ID, lng��ҳID) Then
        If ��������(lng����ID, lng��ҳID) Then
            If ����ģʽ(intinsure) = 0 Then    '�Ƚ�����Ժ
                If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.����, lng����ID) Then Exit Function
            Else
                If Not ����δ�����(lng����ID, lng��ҳID) Then
                    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.����, lng����ID) Then
                        '���ڷֱ治���Ƿ�õ���ҽ���ĳ�����Ժ�ӿڣ�����δ�ɹ�ʱ����ʾ����Ա�Ƿ������HIS�ĳ�����Ժ
                        If MsgBox("����ҽ���ĳ�����Ժ�ӿ�ʱ�������󣬿������ڸò���δ��ҽ�����İ����Ժ������" & vbCrLf & _
                                "���Ƿ���Ըô��󣬼�������HISϵͳ�еĳ�����Ժ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                    End If
                End If
            End If
        
            gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & intinsure & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "��Ժ�Ǽ�")
        Else
            If Not ��Ժ�Ǽ�_��������(lng����ID, intinsure) Then Exit Function
        End If
    Else
        If Not ��Ժ�Ǽ�_��������(lng����ID, intinsure) Then Exit Function
    End If
    
    ��Ժ�Ǽǳ���_�������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �ҺŽ���_��������(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    Dim lng����ID As Long
    Dim cur�����ʻ� As Currency, cur�ֽ� As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    �ҺŽ���_�������� = False
    
    lng����ID = ��ȡ����ID(intinsure)
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.�Һ�, ����Ŀ��.����, lng����ID, lng����ID) Then Exit Function
'    If intInsure = TYPE_��ƽ�� Then
'        gstrSQL = "Select A.���㷽ʽ,Nvl(A.��Ԥ��,0) ��� " & _
'                    " From ����Ԥ����¼ A,�����ʻ� B " & _
'                    " Where A.����ID=B.����ID And B.����=" & intInsure & " And A.����ID=" & lng����ID & " And ���㷽ʽ='�ֽ�' And not (��¼����=1 or ��¼����=11)"
'        '�϶�ֻ���ֽ�֧��
'        Call OpenRecordset(rsTemp, "��ȡ�ֽ�֧����")
'        cur�ֽ� = Nvl(rsTemp!���, 0)
'
'        cur�����ʻ� = �������_��������(lng����ID, intInsure)
'        If cur�����ʻ� <> 0 Then
'            If cur�����ʻ� > cur�ֽ� Then cur�����ʻ� = cur�ֽ�
'            gstrSQL = " insert into ����Ԥ����¼(ID,��¼����,NO,��¼״̬,����ID,��ҳID,����ID,�ɿλ," & _
'                     " ��λ������,��λ�ʺ�,ժҪ,���,���㷽ʽ,�������,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID) " & _
'                     " select ����Ԥ����¼_ID.nextval ID,��¼����,NO,��¼״̬,����ID,��ҳID,����ID, " & _
'                     " �ɿλ,��λ������,��λ�ʺ�,ժҪ,���,'�����ʻ�',�������,�տ�ʱ��,����Ա���, " & _
'                     " ����Ա����," & cur�����ʻ� & ",����ID " & _
'                     " from ����Ԥ����¼" & _
'                     " Where ����ID=" & lng����ID & " And ���㷽ʽ='�ֽ�' And not (��¼����=1 or ��¼����=11)"
'            gcnOracle.Execute gstrSQL
'            '�����ֽ�֧����
'            cur�ֽ� = Val(Format(cur�ֽ� - cur�����ʻ�, "#####0.00"))
'            If cur�ֽ� <> 0 Then
'                '�޸��ֽ�֧����
'                gstrSQL = " Update ����Ԥ����¼ Set ��Ԥ��= " & cur�ֽ� & _
'                          " Where ����ID=" & lng����ID & " And ���㷽ʽ='�ֽ�' And not (��¼����=1 or ��¼����=11)"
'            Else
'                '���ֽ�֧���ɾ����Ԥ����¼
'                gstrSQL = " Delete ����Ԥ����¼ " & _
'                          " Where ����ID=" & lng����ID & " And ���㷽ʽ='�ֽ�' And not (��¼����=1 or ��¼����=11)"
'            End If
'            gcnOracle.Execute gstrSQL
'        End If
'    End If
    If Not ����ҺŽ����¼(intinsure, lng����ID, lng����ID) Then Exit Function
'    If intInsure = TYPE_��ƽ�� Then frm������Ϣ.ShowME (lng����ID)
    
    �ҺŽ���_�������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �ҺŽ������_��������(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    Dim lng����ID As Long
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    �ҺŽ������_�������� = False
    
    '�˴���ʹ�ò���Ԥ����¼������Ϊ���ǵ�����õ�ʱ��Ԥ����¼��������
    gstrSQL = "Select Distinct B.����ID,B.����,B.ҽ����,B.����,B.˳���" & _
        " From ������ü�¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=[1]" & _
        " And A.����ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ���ӿ�", intinsure, lng����ID)
    lng����ID = rsTmp!����ID
    
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.�Һ�, ����Ŀ��.����, lng����ID, lng����ID) Then Exit Function
    If Not ����ҺŽ����¼(intinsure, lng����ID, lng����ID, True) Then Exit Function
    
    �ҺŽ������_�������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ����ҺŽ����¼(ByVal intinsure As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, Optional ByVal bln���� As Boolean = False) As Boolean
    Dim curGhzfy As Currency, strGhlsh As String                            '�Һ��ܷ���,�Һ���ˮ��
    Dim curZhzfe As Currency, curJjzfe As Currency '�ʻ�֧�������֧����
    Dim rsTemp As New ADODB.Recordset
    
    Dim strAdvance As String
    On Error GoTo errHand
    ����ҺŽ����¼ = False
    
    '��ȡ�Һ��ܷ��ã��Һ���ˮ��
    Call Record_Locate(mrsIniItems, "����,Ghfy00")
    curGhzfy = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Ghlsh0")
    strGhlsh = Nvl(mrsIniItems!ֵ, 0)
    Call Record_Locate(mrsIniItems, "����,Zhzfe0")
    curZhzfe = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Jjzfe0")
    curJjzfe = Val(Nvl(mrsIniItems!ֵ, 0))
    
    'ȡ����ID
    If bln���� Then
        gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
                  " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����ID)
        lng����ID = rsTemp("����ID")
    Else
        'У�����
        strAdvance = "�����ʻ�|" & curZhzfe & "||ҽ������|" & curJjzfe
        gstrSQL = " zl_���˽����¼_Update(" & lng����ID & ",'" & strAdvance & "')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
    End If
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        curGhzfy & "," & curGhzfy - curZhzfe - curJjzfe & "," & 0 & "," & curJjzfe & "," & curJjzfe & ",0," & _
        0 & "," & curZhzfe & ",'" & strGhlsh & "')"
'        .��� & "," & .�ʻ��ۼ����� & "," & .�ʻ��ۼ�֧�� & "," & .�ۼƽ���ͳ�� & "," & _
'        .�ۼ�ͳ�ﱨ�� & "," & IIf(.��ҳID = 0, "NULL", .��ҳID) & "," & .���� & "," & .�ⶥ�� & "," & .ʵ������ & "," & _
'        .���ڷ��ü�¼��� & "," & .ȫ�Էѽ�� & "," & .�����Ը���� & "," & .����ͳ���� & "," & .ͳ�ﱨ����� & ",0," & _
'        .�����Ը���� & "," & cur�����ʻ� & ",'')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����Һ�����")
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
    
    ����ҺŽ����¼ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �����������_��������(ByVal rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, ByVal intinsure As Integer) As Boolean
    Dim curMoney As Currency
    '����Ԥ��������
    On Error GoTo errHand
    �����������_�������� = False
    
    Set mrsDetail = rs��ϸ.Clone
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.�շ�, ����Ŀ��.Ԥ����, ��ȡ����ID(intinsure)) Then Exit Function
    
    Call Record_Locate(mrsIniItems, "����,Zhzfe0")
    curMoney = Val(Nvl(mrsIniItems!ֵ, 0))
    str���㷽ʽ = "�����ʻ�;" & curMoney & ";0"
    Call Record_Locate(mrsIniItems, "����,Jjzfe0")
    curMoney = Val(Nvl(mrsIniItems!ֵ, 0))
    str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & curMoney & ";0"
    
    �����������_�������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function �������_��������(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal intinsure As Integer, Optional ByRef strAdvance As String) As Boolean
    Dim curMoney As Currency
    Dim str���㷽ʽ As String
    Dim rsTemp As New ADODB.Recordset
    '������������
    On Error GoTo errHand
    �������_�������� = False
    
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.�շ�, ����Ŀ��.����, ��ȡ����ID(intinsure), lng����ID) Then Exit Function
    
    'Modified by ���� 20031218 ����������
    '�����ʡ��ҽ����������������㣬��Ҫ��������Ϣ����
    If intinsure <> TYPE_�������� Then
        Call Record_Locate(mrsIniItems, "����,Zhzfe0")
        curMoney = Val(Nvl(mrsIniItems!ֵ, 0))
        If curMoney <> 0 Then
            str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & curMoney
        End If
        Call Record_Locate(mrsIniItems, "����,Jjzfe0")
        curMoney = Val(Nvl(mrsIniItems!ֵ, 0))
        If curMoney <> 0 Then
            str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curMoney
        End If
        
        '�������
        If str���㷽ʽ <> "" Then
            str���㷽ʽ = Mid(str���㷽ʽ, 3)
            #If gverControl < 2 Then
                gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',0)"
            #Else
                strAdvance = str���㷽ʽ
                gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
            #End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
        End If
   End If
    
    If Not ���������շѽ����¼(intinsure, lng����ID, lng����ID) Then Exit Function
    
    'Modified by ���� 20031218 ����������
    '�����ʡ��ҽ����������������㣬��Ҫ��������Ϣ��ʾ����
    If intinsure <> TYPE_�������� Then
        #If gverControl < 2 Then
            frm������Ϣ.ShowME (lng����ID)
        #End If
    End If
    
    �������_�������� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_��������(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    On Error GoTo errHand
    ����������_�������� = False
    
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.�շ�, ����Ŀ��.����, lng����ID, lng����ID) Then Exit Function
    If Not ���������շѽ����¼(intinsure, lng����ID, lng����ID, True) Then Exit Function
    
    ����������_�������� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ���������շѽ����¼(ByVal intinsure As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, Optional ByVal bln���� As Boolean = False) As Boolean
    Dim curMzzfy As Currency, curMzzhzfe As Currency, curMzjjzfe As Currency, curMzgrzfe As Currency, curDbgrzf As Currency, strMzlsh As String     '�����ܷ���,�ʻ�֧��,����֧��,�����Ը�,�󲡸����Ը�,������ˮ��
    Dim rsTemp As New ADODB.Recordset
    Dim blnOld As Boolean
    
    On Error GoTo errHand
    ���������շѽ����¼ = False
    #If gverControl < 2 Then
        blnOld = True
    #End If
    
    '��ȡ�Һ��ܷ��ã��Һ���ˮ��
    Call Record_Locate(mrsIniItems, "����,Zhzfe0")
    curMzzhzfe = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Grzfe0")
    curMzgrzfe = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Jjzfe0")
    curMzjjzfe = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Bcbxf0")
    curMzzfy = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Dbgrzf")
    curDbgrzf = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Djlsh0")
    strMzlsh = Nvl(mrsIniItems!ֵ, 0)
    
    'ȡ����ID
    If bln���� Then
        gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
                  " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����ID)
        lng����ID = rsTemp("����ID")
    End If
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        curMzzfy & "," & curMzgrzfe + curDbgrzf & "," & curDbgrzf & "," & curMzjjzfe & "," & curMzjjzfe & ",0," & _
        0 & "," & curMzzhzfe & ",'" & strMzlsh & "',NULL,NULL,NULL" & IIf(blnOld, "", IIf(intinsure <> TYPE_��������, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���������շ�����")
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
    
    ���������շѽ����¼ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_��������(ByVal lng����ID As Long, ByVal intinsure As Integer) As String
    Dim curMoney As Currency, str���㷽ʽ As String, lngPatient As Long
    '������������
    On Error GoTo errHand
    סԺ�������_�������� = ""
    
    'Modified by ���� 20031218 ���������� ʡ��ҽ����֧��Ԥ����
    If Not (intinsure = TYPE_��������) Then
        סԺ�������_�������� = "�����ʻ�;0;0"
        Exit Function
    End If
    
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.����, ����Ŀ��.ˢ��) Then Exit Function
    lngPatient = ��ȡ����ID(intinsure)
    If lngPatient <> lng����ID Then
        MsgBox "������Ϣ������", vbInformation, gstrSysName
        Exit Function
    End If

    '�ȳ�Ժ����㣨������Ա�ĸо�������ˢ������ʾ��
    If ����ģʽ(intinsure) = 1 Then
        If Not ҽ�������Ѿ���Ժ(lngPatient) Then
            MsgBox "���ڸ�ҽ��������Ժ�����ν��㽫��Ϊ��;���㣡", vbInformation, gstrSysName
        Else
            MsgBox "���ν��㽫��Ϊ��Ժ���㣨��������Զ���ҽ���ĳ�Ժ�ӿڣ���", vbInformation, gstrSysName
        End If
    End If
    
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.����, ����Ŀ��.Ԥ����, lng����ID) Then Exit Function
    
    Call Record_Locate(mrsIniItems, "����,Zhzfe0")
    curMoney = Nvl(mrsIniItems!ֵ, 0)
    str���㷽ʽ = "�����ʻ�;" & curMoney & ";0"
    Call Record_Locate(mrsIniItems, "����,Jjzfe0")
    curMoney = Nvl(mrsIniItems!ֵ, 0)
    str���㷽ʽ = str���㷽ʽ & "|ҽ������;" & curMoney & ";0"
    
    סԺ�������_�������� = str���㷽ʽ
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ����_��������(ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal intinsure As Integer, Optional ByRef strAdvance As String) As Boolean
    Dim curMoney As Currency
    Dim lngԤ��ID As Long
    Dim lngPatient As Long
    Dim bln��Ժ As Boolean '��¼�Ƿ���ó�Ժ�ӿ�
    Dim bln��Ժ�ɹ� As Boolean '��¼���ó�Ժ�ӿ��Ƿ�ɹ�
    Dim str���㷽ʽ As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    סԺ����_�������� = False
    bln��Ժ = False: bln��Ժ�ɹ� = False
    
    'Modified by ���� 20031218 ���������� ��Ԥ����ӿڣ��˴�����ý���ˢ����Ϊ��ʽ������׼��
    If intinsure <> TYPE_�������� Then
        If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.����, ����Ŀ��.ˢ��) Then Exit Function
        lngPatient = ��ȡ����ID(intinsure)
        If lngPatient <> lng����ID Then
            Err.Raise 9000, gstrSysName, "������Ϣ������"
            Exit Function
        End If
    
        '�ȳ�Ժ����㣨������Ա�ĸо�������ˢ������ʾ��
        If ����ģʽ(intinsure) = 1 Then
            If Not ҽ�������Ѿ���Ժ(lngPatient) Then
                Err.Raise 9000, gstrSysName, "���ڸ�ҽ��������Ժ�����ν��㽫��Ϊ��;���㣡"
            Else
                Err.Raise 9000, gstrSysName, "���ν��㽫��Ϊ��Ժ���㣨��������Զ���ҽ���ĳ�Ժ�ӿڣ���"
            End If
        End If
    End If
    
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.����, ����Ŀ��.����, lng����ID, lng����ID) Then Exit Function
    
    'Modified by ���� 20031218 ����������
    '�����ʡ��ҽ����������������㣬��Ҫ��������Ϣ����
    If intinsure <> TYPE_�������� Then
        Call Record_Locate(mrsIniItems, "����,Zhzfe0")
        curMoney = Val(Nvl(mrsIniItems!ֵ, 0))
        If curMoney <> 0 Then
            str���㷽ʽ = str���㷽ʽ & "||�����ʻ�|" & curMoney
        End If
        Call Record_Locate(mrsIniItems, "����,Jjzfe0")
        curMoney = Val(Nvl(mrsIniItems!ֵ, 0))
        If curMoney <> 0 Then
            str���㷽ʽ = str���㷽ʽ & "||ҽ������|" & curMoney
        End If
        
        '�������
        If str���㷽ʽ <> "" Then
            str���㷽ʽ = Mid(str���㷽ʽ, 3)
            #If gverControl < 2 Then
                gstrSQL = "zl_���˽����¼_Update(" & lng����ID & ",'" & str���㷽ʽ & "',1)"
            #Else
                strAdvance = str���㷽ʽ
                gstrSQL = "zl_ҽ���˶Ա�_Insert(" & lng����ID & ",'" & str���㷽ʽ & "')"
            #End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
        End If
    End If
    
    If Not ����סԺ�����¼(intinsure, lng����ID, lng����ID) Then Exit Function
    סԺ����_�������� = True
    
    'Modified by ���� 20031218 ����������
    '�����ʡ��ҽ����������������㣬��Ҫ��������Ϣ��ʾ����
    If intinsure <> TYPE_�������� Then
        #If gverControl < 2 Then
            frm������Ϣ.ShowME (lng����ID)
        #End If
    End If
    
    '�ȳ�Ժ�����
    If ����ģʽ(intinsure) = 1 Then
        If ҽ�������Ѿ���Ժ(lng����ID) Then
            If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.ˢ��) Then Exit Function
            If lng����ID <> ��ȡ����ID(intinsure) Then
                Err.Raise 9000, gstrSysName, "������Ϣ������"
                Exit Function
            End If
            
            bln��Ժ = True '���ó�Ժ�ӿ�
            bln��Ժ�ɹ� = frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.��Ժ, ����Ŀ��.����, lng����ID)
        End If
        
        '��ʾ���ýӿڸ�����Ա���Ա��˽⵱ǰ������ҽ�������Ƿ����������Ժ����
        If bln��Ժ Then
            If Not bln��Ժ�ɹ� Then
                Err.Raise 9000, gstrSysName, "��Ժ�������ʧ�ܣ��뵽�����ʻ��в����Ժ����"
            Else
                Err.Raise 9000, gstrSysName, "�ò�����ҽ�����ĳɹ������Ժ������"
            End If
        End If
    End If
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function סԺ�������_��������(ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    Dim lng����ID As Long
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    סԺ�������_�������� = False
    
    gstrSQL = "Select B.����ID,B.����,B.ҽ����,B.����,B.˳���" & _
        " From ���˽��ʼ�¼ A,�����ʻ� B " & _
        " Where A.����ID=B.����ID And B.����=[2]" & _
        " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ���ӿ�", lng����ID, intinsure)
    lng����ID = rsTmp!����ID
    
    If Not frm�ȴ���Ӧ.ShowME(intinsure, ������ʽ.����, ����Ŀ��.����, lng����ID, lng����ID) Then Exit Function
    If Not ����סԺ�����¼(intinsure, lng����ID, lng����ID, True) Then Exit Function
    
    סԺ�������_�������� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Private Function ����סԺ�����¼(ByVal intinsure As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, Optional ByVal bln���� As Boolean = False) As Boolean
    Dim lng��ҳID As Long
    Dim curJszfy As Currency, curJszhzfe As Currency, curJsjjzfe As Currency, curJsgrzfe As Currency, curDbgrzf As Currency, strJslsh As String     '�����ܷ���,�ʻ�֧��,����֧��,�����Ը�,�󲡸����Ը�,������ˮ��
    Dim rsTemp As New ADODB.Recordset
    Dim blnOld As Boolean
    
    On Error GoTo errHand
    ����סԺ�����¼ = False
    #If gverControl < 2 Then
        blnOld = True
    #End If
    
    '��ȡ�Һ��ܷ��ã��Һ���ˮ��
    Call Record_Locate(mrsIniItems, "����,Zhzfe0")
    curJszhzfe = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Grzfe0")
    curJsgrzfe = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Jjzfe0")
    curJsjjzfe = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Bcbxf0")
    curJszfy = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Dbgrzf")
    curDbgrzf = Val(Nvl(mrsIniItems!ֵ, 0))
    Call Record_Locate(mrsIniItems, "����,Djlsh0")
    strJslsh = Nvl(mrsIniItems!ֵ, 0)
    
    'ȡ����ID
    If bln���� Then
        gstrSQL = "select distinct A.ID from ���˽��ʼ�¼ A,���˽��ʼ�¼ B " & _
                  " where A.NO=B.NO and  A.��¼״̬=2 and B.ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����ID)
        lng����ID = rsTemp("ID") '�������ݵ�ID
    End If
    
    gstrSQL = "Select סԺ���� From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID", lng����ID)
    lng��ҳID = rsTemp!סԺ����
    
    '���ս����¼(��Ϊ"����,��¼ID"Ψһ,���Ա����½���ID�϶�Ϊ����)
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & intinsure & "," & lng����ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng��ҳID & "," & 0 & "," & 0 & "," & 0 & "," & _
        curJszfy & "," & curJsgrzfe + curDbgrzf & "," & curDbgrzf & "," & curJsjjzfe & "," & curJsjjzfe & ",0," & _
        0 & "," & curJszhzfe & ",'" & strJslsh & "',NULL,NULL,NULL" & IIf(blnOld, "", IIf(intinsure <> TYPE_��������, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����סԺ��������")
    
    gstrSQL = "zl_���˽��ʼ�¼_�ϴ�(" & lng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
    
    ����סԺ�����¼ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ���ڷ��ü�¼(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim rs���� As New ADODB.Recordset
    '���ô�סԺ�Ƿ�û�з��÷���
    gstrSQL = "Select nvl(count(����ID),0) as ��� " & _
             " From סԺ���ü�¼ " & _
             " Where ����ID=[1] and ��ҳID=[2]" & _
             " And Nvl(��¼״̬,0)<>0"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�סԺ�ڼ�ǹ���", lng����ID, lng��ҳID)
    If rs����.EOF = True Then
        ���ڷ��ü�¼ = False
    Else
        ���ڷ��ü�¼ = (rs����("���") <> 0)
    End If
End Function
'
'Public Function ����δ�����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
'    Dim rs���� As New ADODB.Recordset
'    '���ô�סԺ�Ƿ��з���δ����
'    gstrSQL = "Select nvl(���,0) as ���  from ����δ����� where ����ID=" & lng����ID & " and ��ҳID=" & lng��ҳID
'    Call OpenRecordset(rs����, "�Ƿ����δ�����")
'    If rs����.EOF = True Then
'        ����δ����� = False
'    Else
'        ����δ����� = (rs����("���") <> 0)
'    End If
'End Function

Private Function ��������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    Dim rs���� As New ADODB.Recordset
    '���ô�סԺ�Ƿ��з���δ����
    gstrSQL = "Select Sum(nvl(Ӧ�ս��,0)) as ��� " & _
              "From סԺ���ü�¼ " & _
              " Where ����ID=[1] and ��ҳID=[2] And Nvl(��¼״̬,0)<>0"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ����δ�����", lng����ID, lng��ҳID)
    If rs����.EOF = True Then
        �������� = False
    Else
        �������� = (rs����("���") <> 0)
    End If
End Function

Private Sub ���������Ϣ(ByVal lng����ID As Long, ByVal ������ʽ_IN As Integer, ByVal ����Ŀ��_IN As Integer, ByVal intinsure As Integer)
    Dim strSQL  As String, strValue As String
    Dim rsInfo As New ADODB.Recordset
    Dim cur�ʻ���� As Currency, intסԺ���� As Integer, str����״̬ As String, cur���ҽ�������ۼ� As Currency
    On Error GoTo errHand
    
    'ȡ������ԭ����ֵ
    gstrSQL = " Select ��λ���� As ����״̬,��Ա��� As סԺ����,����֤�� As ���ҽ�������ۼ�,�ʻ���� From �����ʻ�" & _
              " Where ����=[1] And ����ID=[2]"
    Set rsInfo = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����������Ϣ", intinsure, lng����ID)
    cur�ʻ���� = IIf(IsNull(rsInfo.Fields("�ʻ����").Value), 0, rsInfo.Fields("�ʻ����").Value)
    intסԺ���� = IIf(IsNull(rsInfo.Fields("סԺ����").Value), 0, rsInfo.Fields("סԺ����").Value)
    str����״̬ = IIf(IsNull(rsInfo.Fields("����״̬").Value), "", rsInfo.Fields("����״̬").Value)
    cur���ҽ�������ۼ� = IIf(IsNull(rsInfo.Fields("���ҽ�������ۼ�").Value), 0, rsInfo.Fields("���ҽ�������ۼ�").Value)
    strSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure
    
    '�����ʻ����
    Call Record_Locate(mrsIniItems, "����,Grzhye")
    strValue = Val(Nvl(mrsIniItems!ֵ, 0))
    cur�ʻ���� = Val(strValue)
    If cur�ʻ���� < 0 Then cur�ʻ���� = 0
    gstrSQL = strSQL & ",'�ʻ����'," & cur�ʻ���� & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    'סԺ����
    Call Record_Locate(mrsIniItems, "����,Bckbcs")
    strValue = Val(Nvl(mrsIniItems!ֵ, 0))
    intסԺ���� = IIf(intסԺ���� < Val(strValue), Val(strValue), intסԺ����)
    gstrSQL = strSQL & ",'��Ա���'," & intסԺ���� & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '����״̬
    Call Record_Locate(mrsIniItems, "����,Gzztmc")
    strValue = Nvl(mrsIniItems!ֵ, "")
    str����״̬ = IIf(strValue = "", str����״̬, strValue)
    gstrSQL = strSQL & ",'��λ����','''" & str����״̬ & "''')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '���ҽ�������ۼ�
    If ������ʽ_IN = ������ʽ.��Ժ And ����Ŀ��_IN = ����Ŀ��.���� Then
        Call Record_Locate(mrsIniItems, "����,Ndfylj")
        strValue = Val(Nvl(mrsIniItems!ֵ, 0))
        cur���ҽ�������ۼ� = Val(strValue)
        strSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & intinsure
        gstrSQL = strSQL & ",'����֤��'," & cur���ҽ�������ۼ� & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
    End If

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



'------------------------------------���������뺯��------------------------------------
Public Function SendRequest(ByVal ������ʽ_IN As Integer, ByVal ����Ŀ��_IN As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal intinsure As Integer) As Boolean
    Dim objFileSys As New FileSystemObject, objStream As TextStream
    Dim bln��ϸ As Boolean, curMoney As Currency, cur���ʽ�� As Currency
    Dim str�շ�ϸĿ As String, bln�շ� As Boolean
    Dim str�վݷ�Ŀ As String
    Dim rsSecond As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    '�������󣬲��������ļ�
    SendRequest = False
    
    '��������ļ��Ѵ��ڣ�����ɾ���󷢳������ļ�
    If objFileSys.FileExists(mstrPath_�������� & intinsure & "\" & mstrRequest_��������) Then Call objFileSys.DeleteFile(mstrPath_�������� & intinsure & "\" & mstrRequest_��������, True)
    If objFileSys.FileExists(mstrPath_�������� & intinsure & "\" & mstrTemp_��������) Then Call objFileSys.DeleteFile(mstrPath_�������� & intinsure & "\" & mstrTemp_��������, True)
    If objFileSys.FileExists(mstrPath_�������� & intinsure & "\" & mstrReply_��������) Then Call objFileSys.DeleteFile(mstrPath_�������� & intinsure & "\" & mstrReply_��������, True)
    Set objStream = objFileSys.CreateTextFile(mstrPath_�������� & intinsure & "\" & mstrTemp_��������)
    '��д��ͷ
    With mrsIniSection
        .MoveFirst
        .Find "����='" & ������ʽ_IN & ����Ŀ��_IN & "'"
        If .EOF Then
            MsgBox "�ò���δ���壬��ӿڹ������仯����������ṩ����ϵ��", vbInformation, gstrSysName
            GoTo ClearFiles
        End If
        Call OutputData(objStream, Pack(!����))
    End With
    '��д�����־
    Call OutputData(objStream, "Request=TRUE")
    If ����Ŀ��_IN = ����Ŀ��.ˢ�� Then
        Call Record_Clear(mrsIniItems, True)
        GoTo ReturnCall
    End If
    
    bln��ϸ = False
    Select Case ����Ŀ��_IN
    Case ����Ŀ��.����, ����Ŀ��.Ԥ����
        '��д������Ϣ
        If ������ʽ_IN = ������ʽ.��¼ Then
            Call ��ȡ��¼��Ϣ(intinsure)
            Call Record_Locate(mrsIniItems, "����,UserID")
            Call OutputData(objStream, "UserID=" & Nvl(mrsIniItems!ֵ, "supervious"))
            Call Record_Locate(mrsIniItems, "����,Password")
            Call OutputData(objStream, "Password=" & Nvl(mrsIniItems!ֵ, "yb"))
        Else
            Call OutputData(objStream, "Success=")
            Call OutputData(objStream, "Error=")
            Call Record_Locate(mrsIniItems, "����,Cardno")
            Call OutputData(objStream, "Cardno=" & Nvl(mrsIniItems!ֵ, ""))
        End If
        
        Select Case ������ʽ_IN
        Case ������ʽ.��Ժ
            Call ��ȡ��Ժ��Ϣ(lng����ID)
            Call Record_Locate(mrsIniItems, "����,Ryrq00")
            Call OutputData(objStream, "Ryrq00=" & Nvl(mrsIniItems!ֵ, ""))
            Call Record_Locate(mrsIniItems, "����,Rysj00")
            Call OutputData(objStream, "Rysj00=" & Nvl(mrsIniItems!ֵ, ""))
            Call Record_Locate(mrsIniItems, "����,Ryksmc")
            Call OutputData(objStream, "Ryksmc=" & Nvl(mrsIniItems!ֵ, ""))
            Call Record_Locate(mrsIniItems, "����,Rylb00")
            Call OutputData(objStream, "Rylb00=" & Nvl(mrsIniItems!ֵ, ""))
        Case ������ʽ.�Һ�
            Call ��ȡ�Һ���Ϣ(lng����ID, lng����ID)
            Call Record_Locate(mrsIniItems, "����,Ghksmc")
            Call OutputData(objStream, "Ghksmc=" & Nvl(mrsIniItems!ֵ, ""))
            Call Record_Locate(mrsIniItems, "����,Ghfy00")
            Call OutputData(objStream, "Ghfy00=" & Nvl(mrsIniItems!ֵ, ""))
        Case ������ʽ.�շ�
            Call ��ȡ������Ϣ(lng����ID, intinsure)
            Call OutputData(objStream, "Mzlsh0=" & Get��ˮ��(������ʽ_IN, ����Ŀ��_IN, lng����ID, lng����ID, intinsure))
            Call Record_Locate(mrsIniItems, "����,Bqbm00")
            Call OutputData(objStream, "Bqbm00=" & Nvl(mrsIniItems!ֵ, ""))
            bln��ϸ = True
        Case ������ʽ.����
            'ȡ����ҳID
            gstrSQL = "Select Nvl(סԺ����,0) ��ҳID From ������Ϣ Where ����ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ҳID", lng����ID)
            g��������.��ҳID = rsTmp!��ҳID
            Call ��ȡ������Ϣ(lng����ID, intinsure)
            Call OutputData(objStream, "Zylsh0=" & Get��ˮ��(������ʽ_IN, ����Ŀ��_IN, lng����ID, lng����ID, intinsure))
            Call Record_Locate(mrsIniItems, "����,Bqbm00")
            Call OutputData(objStream, "Bqbm00=" & Nvl(mrsIniItems!ֵ, ""))
            bln��ϸ = True
        Case ������ʽ.��Ժ
            Dim rsTemp As New ADODB.Recordset
            Call ��ȡ��Ժ��Ϣ(lng����ID)
            Call OutputData(objStream, "Zylsh0=" & Get��ˮ��(������ʽ_IN, ����Ŀ��_IN, lng����ID, lng����ID, intinsure))
            Call Record_Locate(mrsIniItems, "����,Cyrq00")
            Call OutputData(objStream, "Cyrq00=" & Nvl(mrsIniItems!ֵ, ""))
            Call Record_Locate(mrsIniItems, "����,Cysj00")
            Call OutputData(objStream, "Cysj00=" & Nvl(mrsIniItems!ֵ, ""))
            
            'Modified by ���� 20031218 ����������
            '�����ʡ��ҽ������Ҫ���ݳ�Ժ״̬��������������
            If intinsure = TYPE_����ʡ Or intinsure = TYPE_������ Then
                gstrSQL = "Select A.��Ժ��ʽ From ������ҳ A,������Ϣ B Where A.����ID=B.����ID And A.��ҳID=B.סԺ���� And A.����ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ״̬", lng����ID)
                Call OutputData(objStream, "cyztlx=" & IIf(rsTemp!��Ժ��ʽ = "����", "����", rsTemp!��Ժ��ʽ))
            ElseIf intinsure = TYPE_��ƽ�� Then
                gstrSQL = "Select סԺ���� ��ҳID From ������Ϣ Where ����ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ״̬", lng����ID)
                Call OutputData(objStream, "Cyzd00=" & ��ȡ���Ժ���(lng����ID, rsTemp!��ҳID, False, False))
            End If
        End Select
    Case ����Ŀ��.����
        '��д������Ϣ
        If ������ʽ_IN <> ������ʽ.��¼ Then
            Call ��ȡ������Ϣ(lng����ID)
            Call Record_Locate(mrsIniItems, "����,Cardno")
            Call OutputData(objStream, "Cardno=" & Nvl(mrsIniItems!ֵ, ""))
        End If
        
        Select Case ������ʽ_IN
        Case ������ʽ.��Ժ
            'Ҫ������סԺ��ˮ��
            Call OutputData(objStream, "Cxlsh0=" & Get��ˮ��(������ʽ_IN, ����Ŀ��_IN, lng����ID, lng����ID, intinsure))
        Case ������ʽ.�Һ�
            'Ҫ�����ĹҺ���ˮ��
            Call OutputData(objStream, "Ghlsh0=" & Get��ˮ��(������ʽ_IN, ����Ŀ��_IN, lng����ID, lng����ID, intinsure))
        Case ������ʽ.�շ�, ������ʽ.����
            'Modified by ���� 20031218 ���������� ֻ�и�����ҽ�����ڸ����շ�
'            If ������ʽ_IN = ������ʽ.�շ� And intInsure = TYPE_������ Then
'                Call ��ȡ������Ϣ(lng����ID)
'                Call Record_Locate(mrsIniItems, "����,Bqbm00")
'                If Trim(NVL(mrsIniItems!ֵ)) = "" Then
'                    Call OutputData(objStream, "Cxdjh0=" & Get��ˮ��(������ʽ_IN, ����Ŀ��_IN, lng����ID, lng����ID))
'                Else
'                    Call OutputData(objStream, "gzdjh0=" & Get��ˮ��(������ʽ_IN, ����Ŀ��_IN, lng����ID, lng����ID))
'                End If
'            Else
                Call OutputData(objStream, "Cxdjh0=" & Get��ˮ��(������ʽ_IN, ����Ŀ��_IN, lng����ID, lng����ID, intinsure))
'            End If
        Case ������ʽ.��Ժ
            'Ҫȡ����Ժ��סԺ��ˮ��
            Call OutputData(objStream, "Zylsh0=" & Get��ˮ��(������ʽ_IN, ����Ŀ��_IN, lng����ID, lng����ID, intinsure))
        End Select
    End Select
    
    '���bln��ϸΪ�棬����ȡ����
    If bln��ϸ Then
        If ������ʽ_IN = ������ʽ.���� Then
            bln�շ� = False
            If ����Ŀ��_IN = ����Ŀ��.Ԥ���� Then
                gstrSQL = "Select A.����ID,A.��ҳID,A.Ӥ����,C.��Ŀ���� as ҽ����Ŀ���,  " & _
                         "  A.���մ���ID,A.�շ����,A.�շ�ϸĿID,B.���� as ҽ����Ŀ����,substrb(B.���,1,20) ���, " & _
                         "  A.���㵥λ ��λ, sum(A.����) ����,sum(A.���) ���,'HIS' ҽ��, " & _
                         " decode(C.�Ƿ�ҽ��,1,'Y','N') �Ƿ�ҽ��,A.�վݷ�Ŀ ��Ʊ��Ŀ  " & _
                         "  From (  " & _
                         "       Select Mod(A.��¼����,10) as ��¼����,A.��¼״̬,A.NO,Nvl(A.�۸񸸺�,���) as ���,A.����ID, " & _
                         "       A.��ҳID,Nvl(A.Ӥ����,0) as Ӥ����, A.������ as ҽ��,A.��������ID,A.�շ����,A.�շ�ϸĿID, " & _
                         "       Nvl(A.���մ���ID,0) as ���մ���ID,Avg(Nvl(A.����,1)*A.����) as ����, A.��׼����, " & _
                         "       A.���㵥λ,A.�վݷ�Ŀ,Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0)) as ���,A.����ʱ��, " & _
                         "       Nvl(A.�Ƿ��ϴ�,0) as �Ƿ��ϴ�,Nvl(A.�Ƿ���,0) as �Ƿ���,A.ժҪ  " & _
                         "       From סԺ���ü�¼ A,������Ŀ B  " & _
                         "       Where A.���ʷ��� = 1 And Nvl(A.��¼״̬,0)<>0 And A.������ĿID = B.ID And A.����ID =" & lng����ID & " And A.��ҳID=" & g��������.��ҳID & _
                         "       Group by Mod(A.��¼����,10),A.��¼״̬,A.NO,Nvl(A.�۸񸸺�,���),A.����ID,A.��ҳID, " & _
                         "       Nvl(A.Ӥ����,0),A.������,A.���㵥λ,A.��׼����,A.�վݷ�Ŀ, A.��������ID , A.�շ����, A.�շ�ϸĿID,  " & _
                         "       NVL(A.���մ���ID, 0), A.����ʱ��, NVL(A.�Ƿ��ϴ�, 0), NVL(A.�Ƿ���, 0), A.ժҪ  " & _
                         "       Having Sum(Nvl(A.ʵ�ս��,0))-Sum(Nvl(A.���ʽ��,0))<>0 " & _
                         "       ) A, �շ�ϸĿ B,�շ���� D,���ű� X, " & _
                         "       (Select * From ����֧����Ŀ Where ����=" & intinsure & ") C  " & _
                         "  Where A.�շ�ϸĿID = B.ID And B.ID = C.�շ�ϸĿID(+) And A.��������ID = x.ID And D.���� = A.�շ���� " & _
                         " Group By A.����ID,A.��ҳID,A.Ӥ����,C.��Ŀ���� ,  " & _
                         "  A.���մ���ID,A.�շ����,A.�շ�ϸĿID,B.���� ,substrb(B.���,1,20) , " & _
                         "  A.���㵥λ,decode(C.�Ƿ�ҽ��,1,'Y','N'),A.�վݷ�Ŀ" & _
                         " Having Sum(A.����)<>0"
            Else
                gstrSQL = "Select 'HIS' ҽ��,A.�շ�ϸĿID,C.��Ŀ���� ҽ����Ŀ���,decode(C.�Ƿ�ҽ��,1,'Y','N') �Ƿ�ҽ��,A.�վݷ�Ŀ ��Ʊ��Ŀ,E.���� ҽ����Ŀ����,  " & _
                    " substrb(E.���,1,20) ���,A.���㵥λ ��λ,Sum(A.���ʽ��) ���,Sum(A.����*Nvl(A.����,1)) ����  " & _
                    " From סԺ���ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=" & intinsure & ") C,�շ���� D,�շ�ϸĿ E  " & _
                    " Where A.ִ�в���ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And A.����ID=" & lng����ID & " ANd A.����ID=" & lng����ID & _
                    " And A.�շ����=D.���� And A.�շ�ϸĿID=E.ID And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 And Nvl(A.���ӱ�־,0)<>9 " & _
                    " Group By A.�շ�ϸĿID,C.��Ŀ����,C.�Ƿ�ҽ��,A.�վݷ�Ŀ,E.����,substrb(E.���,1,20),A.���㵥λ  " & _
                    " Having Sum(A.����*Nvl(A.����,1))<>0"
            End If
            Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "�ϴ����ʷ���")
        Else
            'ȡ�շ�ϸĿID
            str�շ�ϸĿ = ",,"
            bln�շ� = True
            
            If ����Ŀ��_IN = ����Ŀ��.���� Then
                '���������Ҫ������ȡ��¼��
                gstrSQL = "Select A.������,A.�շ�ϸĿID,C.��Ŀ���� ҽ����Ŀ���,decode(C.�Ƿ�ҽ��,1,'Y','N') �Ƿ�ҽ��,A.�վݷ�Ŀ,E.���� ҽ����Ŀ����,  " & _
                    " substrb(E.���,1,20) ���,A.���㵥λ,A.ʵ�ս��,A.��׼���� ����,A.����*Nvl(A.����,1) ����  " & _
                    " From ������ü�¼ A,���ű� B,(Select * From ����֧����Ŀ Where ����=" & intinsure & ") C,�շ���� D,�շ�ϸĿ E  " & _
                    " Where A.ִ�в���ID+0=B.ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID(+) And A.����ID=[2] ANd A.����ID=[1]" & _
                    " And A.�շ����=D.���� And A.�շ�ϸĿID=E.ID And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 And Nvl(A.���ӱ�־,0)<>9"
                Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", lng����ID, lng����ID)
            End If
            
            With mrsDetail
                Do While Not .EOF
                    If InStr(1, str�շ�ϸĿ, "," & !�շ�ϸĿID & ",") = 0 Then
                        str�շ�ϸĿ = str�շ�ϸĿ & !�շ�ϸĿID & ","
                    End If
                    .MoveNext
                Loop
                str�շ�ϸĿ = Mid(str�շ�ϸĿ, 3)
                str�շ�ϸĿ = Mid(str�շ�ϸĿ, 1, Len(str�շ�ϸĿ) - 1)
                If .RecordCount <> 0 Then .MoveFirst
            End With
            
            gstrSQL = "Select D.ID �շ�ϸĿID,D.���� ҽ����Ŀ����,substrb(D.���,1,20) ���,C.��Ŀ���� ҽ����Ŀ���" & _
                    " From �շ�ϸĿ D,(Select * From ����֧����Ŀ Where ����=[1]) C" & _
                    " Where D.ID=C.�շ�ϸĿID(+) And D.ID IN ([2])"
            Set rsSecond = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", intinsure, str�շ�ϸĿ)
        End If
        
        With mrsDetail
            'Modified by ���� 20031218 ����������
            '��д������ϸ��¼��
            If ����Ŀ��_IN = ����Ŀ��.Ԥ���� Or (intinsure <> TYPE_��������) Then curTotalMoney = 0
            Call OutputData(objStream, "Cfxms0=" & .RecordCount)
            '��дƱ����Ŀ�����Cfxms0=0���򲻻���᱾�����ݣ������д�Ļ�����˵�������Ǽ��շѣ�
            
            '��д������ϸ��һ����¼ʮ�У�
            Do While Not .EOF
                If .AbsolutePosition = 1 Then
                    With mrsIniSection
                        .MoveFirst
                        .Find "����='" & ������ʽ_IN & ����Ŀ��.��ϸ & "'"
                        If .EOF Then
                            MsgBox "�ò���δ���壬��ӿڹ������仯����������ṩ����ϵ��", vbInformation, gstrSysName
                            GoTo ClearFiles
                        End If
                        'д��ͷ
                        Call OutputData(objStream, Pack(!����))
                    End With
                End If
                
                'Modified by ���� 20031218 ���������� ֻ����·ҽ��������Ŀ
                '�ж��շ�ϸĿ�Ƿ��Ѿ���ˣ����û�У�����ʾ���˳�
'                If intInsure = TYPE_�������� Then If �����Ŀ("", !�շ�ϸĿID) = False Then GoTo ClearFiles
                
                '����Ϊʮ������
                If Not bln�շ� Then
                    Call OutputData(objStream, Nvl(!ҽ����Ŀ���, ""))
                    Call OutputData(objStream, !�Ƿ�ҽ��)
                    str�վݷ�Ŀ = Nvl(!��Ʊ��Ŀ, "")
                    If intinsure <> TYPE_�������� Then     'ʡ��ҽ����������
                        If str�վݷ�Ŀ = "���Ʒ�" Then str�վݷ�Ŀ = "����"
                    End If
                    Call OutputData(objStream, str�վݷ�Ŀ)
                    Call OutputData(objStream, Nvl(!ҽ����Ŀ����, ""))
                    Call OutputData(objStream, Nvl(!���, "��"))
                    Call OutputData(objStream, Nvl(!��λ, "��"))
                    curMoney = Nvl(!���, 0)
                    Call OutputData(objStream, Format(curMoney / !����, "#####0.0000;-#####0.0000;0;0"))
                    Call OutputData(objStream, Format(!����, "#####0.00;-#####0.00;0;0"))
                    Call OutputData(objStream, Format(curMoney, "#####0.00;-#####0.00;0;0"))
                    Call OutputData(objStream, Nvl(!ҽ��, ""))
                    'Modified by ���� 20031218 ����������
                    If ����Ŀ��_IN = ����Ŀ��.Ԥ���� Or (intinsure <> TYPE_��������) Then curTotalMoney = curTotalMoney + Nvl(!���, 0)
                Else
                    rsSecond.MoveFirst
                    rsSecond.Find "�շ�ϸĿID=" & !�շ�ϸĿID
                    Call OutputData(objStream, Nvl(rsSecond!ҽ����Ŀ���, ""))
                    If ����Ŀ��_IN = ����Ŀ��.Ԥ���� Then
                        Call OutputData(objStream, IIf(!�Ƿ�ҽ�� = 1, "Y", "N"))
                    Else
                        Call OutputData(objStream, !�Ƿ�ҽ��)
                    End If
                    str�վݷ�Ŀ = Nvl(!�վݷ�Ŀ, "")
                    If intinsure <> TYPE_�������� Then
                        If str�վݷ�Ŀ = "���Ʒ�" Then str�վݷ�Ŀ = "����"
                    End If
                    Call OutputData(objStream, str�վݷ�Ŀ)
                    Call OutputData(objStream, Nvl(rsSecond!ҽ����Ŀ����, ""))
                    Call OutputData(objStream, Nvl(rsSecond!���, "��"))
                    Call OutputData(objStream, Nvl(!���㵥λ, "��"))
                    curMoney = Nvl(!ʵ�ս��, 0)
                    Call OutputData(objStream, Format(curMoney / !����, "#####0.0000;-#####0.0000;0;0"))
                    Call OutputData(objStream, Format(!����, "#####0.00;-#####0.00;0;0"))
                    Call OutputData(objStream, Format(curMoney, "#####0.00;-#####0.00;0;0"))
                    Call OutputData(objStream, Nvl(!������, ""))
                    'Modified by ���� 20031218 ����������
                    If ����Ŀ��_IN = ����Ŀ��.Ԥ���� Or (intinsure <> TYPE_��������) Then curTotalMoney = curTotalMoney + Nvl(!ʵ�ս��, 0)
                End If
                If curMoney < 0 Then
                    MsgBox "ҽ����֧�ָ����ϴ������飡", vbInformation, gstrSysName
                    GoTo ClearFiles
                End If
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
            
            '����ϴ��������ʽ��ȣ����ֹ���ʣ����ڸ���������ɵģ�
            Dim rs���� As New ADODB.Recordset
            If lng����ID <> 0 Then
                gstrSQL = " Select Sum(���ʽ��) ���ʽ�� From " & IIf(������ʽ_IN = ������ʽ.����, "סԺ���ü�¼", "������ü�¼") & _
                          " Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"
                Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "���ν��ʽ��", lng����ID)
                cur���ʽ�� = Nvl(rs����!���ʽ��, 0)
                
                If Format(cur���ʽ��, "#####.00;-#####.00;0;") <> Format(curTotalMoney, "#####.00;-#####.00;0;") Then
                    MsgBox "�ϴ��������ʽ��������������ڸ��������ĵ�����ɵģ����飡", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End With
    End If
ReturnCall:
    objStream.Close
    objFileSys.GetFile(mstrPath_�������� & intinsure & "\" & mstrTemp_��������).Name = mstrRequest_��������
    SendRequest = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
ClearFiles:
    On Error Resume Next
    objStream.Close
    If objFileSys.FileExists(mstrPath_�������� & intinsure & "\" & mstrTemp_��������) Then Call objFileSys.DeleteFile(mstrPath_�������� & intinsure & "\" & mstrTemp_��������, True)
    If objFileSys.FileExists(mstrPath_�������� & intinsure & "\" & mstrRequest_��������) Then Call objFileSys.DeleteFile(mstrPath_�������� & intinsure & "\" & mstrRequest_��������, True)
    If objFileSys.FileExists(mstrPath_�������� & intinsure & "\" & mstrReply_��������) Then Call objFileSys.DeleteFile(mstrPath_�������� & intinsure & "\" & mstrReply_��������, True)
End Function

Private Function Get��ˮ��(ByVal ������ʽ_IN As Integer, ByVal ����Ŀ��_IN As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal intinsure As Integer) As String
    Dim RSPATIENT As New ADODB.Recordset
    Dim int���� As Integer
    '�Һ���������ˮ�ţ���Ӧ���ļ����عҺ���ˮ��
    '�Һų���������Һ���ˮ�ţ�Ӧ���ļ����س�����ˮ��
    '����ˢ��������������ˮ��
    '�������󣺴���������ˮ�ţ�Ӧ���ļ����ص�����ˮ��
    '�������������������ݺţ�Ӧ���ļ����ص�����ˮ��
    '��Ժ��������ˮ�ţ�Ӧ���ļ�����סԺ��ˮ��
    '��Ժ����������סԺ��ˮ�ţ�Ӧ���ļ����س�����ˮ��
    '�������󣺴���סԺ��ˮ�ţ�Ӧ���ļ����ص�����ˮ��
    '���ʳ���������������ݺţ�Ӧ���ļ����ص�����ˮ��
    '��Ժ���󣺴���סԺ��ˮ��
    '��Ժ����������סԺ��ˮ��
    '����Ϊ-1����ʾ����
    '���ս����¼�����ʺ��壺1-����;2-סԺ
    
    Get��ˮ�� = ""
    
    '����Դ����ﻹ��סԺ
    If ������ʽ_IN = ������ʽ.���� Then
        int���� = 2
    Else
        int���� = 1
    End If
    
    '���ݲ�����ʽ������Ŀ�ģ�ȡ����Ҫ����ˮ��
    Select Case ������ʽ_IN
    Case ������ʽ.��Ժ
        Select Case ����Ŀ��_IN
        'Case ����Ŀ��.����
        Case ����Ŀ��.����
            gstrSQL = " Select ˳��� as ��ˮ�� From �����ʻ�" & _
                     " Where ����ID=" & lng����ID & " And ����=" & intinsure
        End Select
    Case ������ʽ.�Һ�
        Select Case ����Ŀ��_IN
        'Case ����Ŀ��.����
        Case ����Ŀ��.����
            'ȡԭʼ���ʼ�¼����ˮ��
            gstrSQL = " Select ֧��˳��� as ��ˮ�� From ���ս����¼ " & _
                      " Where ��¼ID=" & lng����ID & " And ����=" & int����
        End Select
    Case ������ʽ.�շ�
        Select Case ����Ŀ��_IN
        Case ����Ŀ��.����, ����Ŀ��.Ԥ����
            Call Record_Locate(mrsIniItems, "����,Mzlsh0")
            Get��ˮ�� = Nvl(mrsIniItems!ֵ, "")
            Exit Function
        Case ����Ŀ��.����
            'ȡԭʼ���ʼ�¼����ˮ��
            gstrSQL = " Select ֧��˳��� as ��ˮ�� From ���ս����¼ " & _
                      " Where ��¼ID=" & lng����ID & " And ����=" & int����
        'Case ����Ŀ��.ˢ��
        End Select
    Case ������ʽ.����
        Select Case ����Ŀ��_IN
        Case ����Ŀ��.����, ����Ŀ��.Ԥ����
            gstrSQL = " Select ˳��� as ��ˮ�� From �����ʻ�" & _
                     " Where ����ID=" & lng����ID & " And ����=" & intinsure
        Case ����Ŀ��.����
            'ȡԭʼ���ʼ�¼����ˮ��
            gstrSQL = " Select ֧��˳��� as ��ˮ�� From ���ս����¼ " & _
                      " Where ��¼ID=" & lng����ID & " And ����=" & int����
        End Select
    Case ������ʽ.��Ժ
        gstrSQL = " Select ˳��� as ��ˮ�� From �����ʻ�" & _
                 " Where ����ID=" & lng����ID & _
                 " And ����=" & intinsure
    End Select
    Set RSPATIENT = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����˵���ˮ��")
    Get��ˮ�� = RSPATIENT!��ˮ��
End Function

Public Function AnalyseReply(ByVal ������ʽ_IN As Integer, ByVal ����Ŀ��_IN As Integer, ByVal intinsure As Integer) As Integer
    Dim objFileSys As New FileSystemObject, objStream As TextStream
    Dim strCompare As String, strLine As String, strSection As String
    Dim strField As String, strValue As String
    Dim strError As String, strIdentify As String, strAddition As String, str˳��� As String
    Dim lng����ID As Long, intסԺ���� As Integer, lng����ID As Long
    Dim str���� As String, strҽ���� As String
    Dim rsTmp As New ADODB.Recordset
    '������Ӧ�ļ�
    '����ע�͵ĸ�ʽΪ���Ⱥź��Ǹ��ֶζ�Ӧ����������
'    �ӿ�Ӧ���ļ�����ʱ���вα�����Ϣ�����вα��˵ĸ�����Ϣ�磺�������Ա����䡢��λ��ic��״̬������״̬�������˻��������������ĵȣ�����Ľӿ�˵���о���"<<�α���������Ϣ>>"��������
'            xming0=����
'            xbie00=�Ա�
'            brnl00=����
'            dwmc00=��λ����
'            icztmc=IC��״̬
'            gzztmc=����״̬
'            grzhye=�����ʻ����
'            dqmc00=Ͷ����������������
'            fzxmc0=Ͷ������������������
    '�ӿ�Ӧ���ļ�����ʱ���д�����ϸ��Ϣ�������շ���Ŀ�ĸ�����Ϣ�磺���ơ����ȣ�����Ľӿ�˵���о���"<<������ϸ��Ϣ>>"��������
'            ҽԺ�շ���Ŀ��ҽ�����ĵı��=ҽ����Ŀ���
'            �Ƿ�ҽ����Ŀ=�Ƿ�ҽ����Ŀ
'            ҽԺ�շ���Ŀ��ҽ�����ĵķ�Ʊ��Ŀ����=��Ʊ��Ŀ
'            ҽԺ�շ���Ŀ��ҽ�����ĵ�����=ҽ����Ŀ����
'            ҽԺ�շ���Ŀ��ҽԺ�Ĺ��=���
'            ҽԺ�շ���Ŀ��ҽԺ�ĵ�λ=��λ
'            ҽԺ�շ���Ŀ��ҽԺ�ĵ���=����
'            ҽԺ�շ���Ŀ������=����
'            ҽԺ�շ���Ŀ�Ľ��=���
'            ҽԺ�շ���Ŀ��ҽ������=ҽ������
'        ���⣬�ӿڷ��ص��շ��ļ���<<������ϸ��Ϣ>>����������Ϣ�⣬
'        ������һ����Ϣ��ΪҽԺ�շ���Ŀ��ҽ�����ĵĸ����Ը�������0 ��1����
'            �Ը�����=�Ը�����
    '�����ļ��еķ�Ʊ��Ŀ���ֽ⵽[yb0000]��[fyb000]����С���У�
    '�ֱ��������ҽ����Ŀ���úͰ����߹涨�����Ը���Ŀ���á�

    On Error Resume Next
    AnalyseReply = 0
    
    '��������־����������
    Call Record_Clear(mrsIniItems, False)
    
    '��������[fpxmbm]��[mzsfmx]��[yb0000]��[fyb000]��[zysfmx]���˳�
    strCompare = UCase("[fpxmbm]��[mzsfmx]��[yb0000]��[fybfy0]��[zysfmx]��[tsxm00]��[ybgr00]")
    If Not objFileSys.FileExists(mstrPath_�������� & intinsure & "\" & mstrReply_��������) Then Exit Function
    Err = 0
    Set objStream = objFileSys.OpenTextFile(mstrPath_�������� & intinsure & "\" & mstrReply_��������, ForReading, False, TristateMixed)
    If Err = 70 Then Err = 0: Exit Function '�ܾ���Ȩ�ޣ�˵��ҽ���������ڶ������з���
    
    On Error GoTo errHand
    Err = 0
    With objStream
        Do While Not .AtEndOfStream
            strLine = UCase(.ReadLine)
            If InStr(1, strCompare, strLine) <> 0 Then
                Exit Do
            Else
                '������ǽ���������¼�¼��
                If Mid(strLine, 1, 1) = "[" Then
                    '�ж��Ƿ�Ͳ���һ��
                    With mrsIniSection
                        .MoveFirst
                        .Find "����='" & ������ʽ_IN & ����Ŀ��_IN & "'"
                        If .EOF Then
                            MsgBox "�ò���δ���壬��ӿڹ������仯����������ṩ����ϵ��", vbInformation, gstrSysName
                            Exit Function
                        End If
                        strSection = Pack(!����)
                    End With
                    If strSection <> strLine Then
                        If .Line = 2 Then
                            'MsgBox "�����Ӧ���ļ���������Ҫ����Һŵ�Ӧ���ļ�����������ȴ��סԺ�Ǽǵ�Ӧ���ļ���", vbInformation, gstrSysName
                            Exit Function
                        Else
                            If MsgBox("�ý���" & strLine & "δ���壬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                Exit Do
                            Else
                                AnalyseReply = 2
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    On Error Resume Next
                    Err = 0
                    strField = UCase(Split(strLine, "=")(0))
                    strValue = Trim(UCase(Split(strLine, "=")(1)))
                    If Err = 0 Then
                        If strField = "REPLY" Then
                            If strValue <> "TRUE" Then Exit Function
                        End If
                        If Not UpdateData(strField, strValue) Then Exit Function
                    End If
                End If
            End If
        Loop
        .Close
    End With
    
    '�ж��Ƿ����
    Call Record_Locate(mrsIniItems, "����,Success")
    If Nvl(mrsIniItems!ֵ, "FALSE") = "FALSE" Then
        Call Record_Locate(mrsIniItems, "����,Error")
        strError = Nvl(mrsIniItems!ֵ, "")
        MsgBox strError, vbInformation, gstrSysName
        AnalyseReply = 2
        Exit Function
    End If
    If (������ʽ_IN = ������ʽ.�Һ� Or ������ʽ_IN = ������ʽ.��Ժ) And ����Ŀ��_IN = ����Ŀ��.ˢ�� Then
        Call Record_Locate(mrsIniItems, "����,Valid0")
        If Nvl(mrsIniItems!ֵ, "FALSE") = "FALSE" Then
            Call Record_Locate(mrsIniItems, "����,Bnghyy")
            If Nvl(mrsIniItems!ֵ, "") <> "" Then strError = Nvl(mrsIniItems!ֵ, "")
            Call Record_Locate(mrsIniItems, "����,Bndjyy")
            If Nvl(mrsIniItems!ֵ, "") <> "" Then strError = Nvl(mrsIniItems!ֵ, "")
            MsgBox strError, vbInformation, gstrSysName
            AnalyseReply = 2
            Exit Function
        End If
    End If

    
    '���û�и�ҽ�����˵���Ϣ���򴴽���������²�����Ϣ
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    If ����Ŀ��_IN = ����Ŀ��.ˢ�� And (������ʽ_IN = ������ʽ.�Һ� Or ������ʽ_IN = ������ʽ.�շ� Or ������ʽ_IN = ������ʽ.��Ժ) Then
        If lng����ID <> 0 And lng����ID <> ��ȡ����ID(intinsure) Then
            MsgBox "������Ϣ���������ܰ���", vbInformation, gstrSysName
            AnalyseReply = 2
            Exit Function
        End If
        
        Call Record_Locate(mrsIniItems, "����,Cardno")
        str���� = Nvl(mrsIniItems!ֵ, "")
        strIdentify = str����                              '0����
        Call Record_Locate(mrsIniItems, "����,ID0000")
        strҽ���� = Nvl(mrsIniItems!ֵ, "")
        strIdentify = strIdentify & ";" & strҽ����          '1ҽ����
        strIdentify = strIdentify & ";"                                    '2����
        Call Record_Locate(mrsIniItems, "����,Xming0")
        strIdentify = strIdentify & ";" & Nvl(mrsIniItems!ֵ, "")     '3����
        Call Record_Locate(mrsIniItems, "����,Xbie00")
        strValue = Nvl(mrsIniItems!ֵ, "")
        'Modified by ���� 20031218 ����������
        If intinsure = TYPE_�������� Then
            strValue = IIf(strValue = "1", "��", IIf(strValue = "2", "Ů", ""))
        Else
            strValue = IIf(strValue = "0", "��", IIf(strValue = "1", "Ů", ""))
        End If
        strIdentify = strIdentify & ";" & strValue '4�Ա�
        Call Record_Locate(mrsIniItems, "����,Brnl00")
        If Len(strҽ����) = 18 Then
            strIdentify = strIdentify & ";" & Mid(strҽ����, 7, 4) & "-" & Mid(strҽ����, 11, 2) & "-" & Mid(strҽ����, 13, 2)  '5��������
        Else
            strIdentify = strIdentify & ";19" & Mid(strҽ����, 7, 2) & "-" & Mid(strҽ����, 9, 2) & "-" & Mid(strҽ����, 11, 2) '5��������
        End If
        'strIdentify = strIdentify & ";" & DateAdd("YYYY", -1 * Nvl(mrsIniItems!ֵ, 0), zlDatabase.Currentdate)  '5��������
        strIdentify = strIdentify & ";"   '6���֤
        Call Record_Locate(mrsIniItems, "����,Dwmc00")
        strIdentify = strIdentify & ";" & Nvl(mrsIniItems!ֵ, "")  '7.��λ����(����)
        strAddition = ";"                                  '8.���Ĵ���
        strAddition = strAddition & ";"                             '9.˳���
        '����סԺ����������ʱ�����أ���ˣ������ݿ���ȡ��סԺ���������С�ڵ����������ݿ��е�����Ϊ׼
        intסԺ���� = 0: lng����ID = 0
        If lng����ID = 0 Then lng����ID = ��ȡ����ID(intinsure)
        If lng����ID <> 0 Then
            gstrSQL = "Select Nvl(��Ա���,0) סԺ���� From �����ʻ� Where ����ID=[1] And ����=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡסԺ����", lng����ID, intinsure)
            intסԺ���� = rsTmp!סԺ����
            
            gstrSQL = "Select Nvl(����ID,0) ����ID From �����ʻ� Where ����ID=[1] And ����=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ID", lng����ID, intinsure)
            lng����ID = rsTmp!����ID
        End If
        Call Record_Locate(mrsIniItems, "����,Bckbcs")
        If intסԺ���� <= Val(Nvl(mrsIniItems!ֵ, 0)) Then intסԺ���� = Val(Nvl(mrsIniItems!ֵ, 0))
        strAddition = strAddition & ";" & intסԺ����        '10��Ա���
        Call Record_Locate(mrsIniItems, "����,Grzhye")
        strAddition = strAddition & ";" & Nvl(mrsIniItems!ֵ, 0)   '11�ʻ����
        strAddition = strAddition & ";0"                            '12��ǰ״̬
        Call Record_Locate(mrsIniItems, "����,Bqbm00")
        strAddition = strAddition & ";" & IIf(lng����ID = 0, "'NULL'", lng����ID) '13����ID
        strAddition = strAddition & ";" & 1 '14��ְ(1,2,3)
        strAddition = strAddition & ";"     '15����֤��
        Call Record_Locate(mrsIniItems, "����,Brnl00")
        strAddition = strAddition & ";" & Nvl(mrsIniItems!ֵ, 0) '16�����
        strAddition = strAddition & ";"                             '17�Ҷȼ�
        Call Record_Locate(mrsIniItems, "����,Grzhye")
        strAddition = strAddition & ";" & Nvl(mrsIniItems!ֵ, 0)      '18�ʻ������ۼ�
        strAddition = strAddition & ";0"       '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0;0"       '20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�
        strAddition = strAddition & ";" & intסԺ���� & ";"       '22סԺ�����ۼ�
        
        '���в�����Ժʱ����������ʻ��д��ڸò��˵���Ϣ�����ʻ��еĲ���ID�봫��Ĳ���ID��������ʾ����Ա�Ⱥϲ����ٰ�������Ժ�Ǽ�
        If bln������Ժ Then
            gstrSQL = "Select ����ID From �����ʻ� Where ����=[1] And ҽ����=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��˵ı����ʻ��еĲ���ID", str����, strҽ����)
            If Not rsTmp.EOF Then
                If rsTmp!����ID <> glng����ID Then
                    MsgBox "���Ƚ�������ݺϲ����ٰ�������Ժ�Ǽǣ�" & vbCrLf & _
                    "��ǰ����ID[" & glng����ID & "]����ǰ�Ĳ���ID[" & rsTmp!����ID & "]", vbInformation, gstrSysName
                    AnalyseReply = 2
                    Exit Function
                End If
            End If
            If lng����ID = 0 Then lng����ID = glng����ID
        Else
            '����ò����Ѿ����ڱ����ʻ����򲻲���������Ϣ
            gstrSQL = "Select ����ID From �����ʻ� Where ����=[1] And ҽ����=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ���ڱ����ʻ�", str����, strҽ����)
            If Not rsTmp.EOF Then
                lng����ID = rsTmp!����ID
            End If
        End If
        lng����ID = BuildPatiInfo(1, strIdentify & strAddition, lng����ID, intinsure)
        '���ظ�ʽ:�м���벡��ID
        If lng����ID > 0 Then
            mgstrPatientInfo = strIdentify & ";" & lng����ID & strAddition
        End If
    End If
    
    '����ҽ�������ʻ����(�ҺŽ��㡢������㡢סԺ���㡢��Ժ�Ǽ�)
    Call ���������Ϣ(��ȡ����ID(intinsure), ������ʽ_IN, ����Ŀ��_IN, intinsure)
    
    '�����סԺ�շ�ˢ����������ʻ��е���Ժ˳��ţ�������������ԭ�������˳���Ϊ�յ����⣩
    If ����Ŀ��_IN = ����Ŀ��.ˢ�� And ������ʽ_IN = ������ʽ.���� Then
        Call Record_Locate(mrsIniItems, "����,Zylsh0")
        str˳��� = Nvl(mrsIniItems!ֵ, "")
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & ��ȡ����ID(intinsure) & "," & intinsure & ",'˳���','''" & str˳��� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ��ˮ��")
    End If
    
    '��ȡ���ɾ��Ӧ���ļ�
'    If objFileSys.FileExists(mstrPath_�������� & intInsure & "\" & mstrRequest_��������) Then
'        Call objFileSys.DeleteFile(mstrPath_�������� & intInsure & "\" & mstrRequest_��������, True)
'    End If
    Call objFileSys.DeleteFile(mstrPath_�������� & intinsure & "\" & mstrReply_��������, True)
    AnalyseReply = 1
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitStruc() As Boolean
    '��ʼ����¼��
    mstrFields = "����" & "," & adLongVarChar & "," & "200" & "|" & _
                 "��������" & "," & adLongVarChar & "," & "200" & "|" & _
                 "����" & "," & adLongVarChar & "," & "500" & "|" & _
                 "˵��" & "," & adLongVarChar & "," & "2000" & "|" & _
                 "ֵ" & "," & adLongVarChar & "," & "500" & "|" & _
                 "����" & "," & adDouble & "," & "2" & "|" & _
                 "�̶���" & "," & adDouble & "," & "2"
    Call Record_Init(mrsIniItems, mstrFields)
    mstrFields = "����" & "," & adLongVarChar & "," & "20" & "|" & _
                 "��������" & "," & adLongVarChar & "," & "200" & "|" & _
                 "����" & "," & adLongVarChar & "," & "2"
    Call Record_Init(mrsIniSection, mstrFields)
    'װ���ʼ����
    InitStruc = Record_Prepare
End Function

Private Function Pack(ByVal strSection As String) As String
    'Ϊ�������ϰ�װ �磺[Section]
    Pack = UCase("[" & strSection & "]")
End Function

Private Function UnPack(ByVal strSection As String) As String
    '��ԭΪԭʼ����
    UnPack = UCase(Mid(strSection, 2, Len(strSection) - 2))
End Function

Private Sub ��ȡ��¼��Ϣ(ByVal intinsure As Integer)
    Dim rsInfo As New ADODB.Recordset
    
    gstrSQL = "Select * From ���ղ��� Where ����=[1] Order by ���"
    Set rsInfo = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ������", intinsure)
    
    Call UpdateData("UserID", Nvl(rsInfo!����ֵ, "supervisor"))
    rsInfo.MoveNext
    Call UpdateData("Password", Nvl(rsInfo!����ֵ, "yb"))
    rsInfo.Close
End Sub

Private Sub ��ȡ��Ժ��Ϣ(ByVal lng����ID As Long)
    Dim strValue As String
    Dim rs��Ժ As New ADODB.Recordset
    
    gstrSQL = "Select to_char(A.��Ժʱ��,'yyyy-MM-dd hh24:mi:ss') ��Ժʱ��,B.���� ���� " & _
            " From ������Ϣ A,���ű� B " & _
            " Where A.��ǰ����ID=B.ID And A.����ID=[1]"
    Set rs��Ժ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ��Ϣ", lng����ID)
    
    With rs��Ժ
        strValue = Format(!��Ժʱ��, "yyyyMMdd")
        Call UpdateData("Ryrq00", strValue)
        strValue = Format(!��Ժʱ��, "HHmm")
        Call UpdateData("Rysj00", strValue)
        Call UpdateData("Ryksmc", IIf(IsNull(!����), "", !����))
        Call UpdateData("Rylb00", "��ͨ")
    End With
End Sub

Private Sub ��ȡ�Һ���Ϣ(ByVal lng����ID As Long, ByVal lng����ID As Long)
    Dim strValue As String
    Dim rs�Һ� As New ADODB.Recordset
    
    gstrSQL = "Select Sum(A.ʵ�ս��) ����,B.���� ���� " & _
            " From ������ü�¼ A,���ű� B " & _
            " Where A.ִ�в���ID=B.ID ANd A.����ID=[1] And ��¼����=4 And ����ID=[2]" & _
            " Group by B.����"
    Set rs�Һ� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�Һ���Ϣ", lng����ID, lng����ID)
    
    With rs�Һ�
        strValue = Nvl(!����, "")
        Call UpdateData("Ghksmc", strValue)
        strValue = Format(!����, "#####0.00;-#####0.00;0;")
        Call UpdateData("Ghfy00", strValue)
    End With
End Sub

Private Function ��ȡ������Ϣ(ByVal lng����ID As Long, ByVal intinsure As Integer) As ADODB.Recordset
    Dim strValue As String
    Dim rs�շ� As New ADODB.Recordset
    
    gstrSQL = "Select substr(B.����,1,instr(B.����,'@@')-1) ���� From �����ʻ� A,���ղ��� B " & _
                    " Where A.����=B.����(+) And A.����ID=B.ID(+) And A.����ID=[1] And A.����=[2]"
    Set rs�շ� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����˲�����Ϣ", lng����ID, intinsure)
    
    With rs�շ�
        strValue = Nvl(!����, "")
        Call UpdateData("Bqbm00", strValue)
    End With
End Function

Private Sub ��ȡ��Ժ��Ϣ(ByVal lng����ID As Long)
    Dim strValue As String
    Dim rs��Ժ As New ADODB.Recordset
    
    gstrSQL = "Select to_char(��Ժʱ��,'yyyy-MM-dd hh24:mi:ss') ��Ժʱ�� " & _
            " From ������Ϣ " & _
            " Where ����ID=[1]"
    Set rs��Ժ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ժ��Ϣ", lng����ID)
    
    With rs��Ժ
        strValue = Format(!��Ժʱ��, "yyyyMMdd")
        Call UpdateData("Cyrq00", strValue)
        strValue = Format(!��Ժʱ��, "HHmm")
        Call UpdateData("Cysj00", strValue)
    End With
End Sub

Private Sub ��ȡ������Ϣ(ByVal lng����ID As Long)
    Dim strValue As String
    Dim rs���� As New ADODB.Recordset
    
    gstrSQL = "Select ���� " & _
            " From �����ʻ� " & _
            " Where ����ID=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����˵Ŀ���", lng����ID)
    
    With rs����
        strValue = Nvl(!����, "")
        Call UpdateData("Cardno", strValue)
    End With
End Sub

Public Function ��ȡ����ID(ByVal intinsure As Integer) As Long
    Dim rsTmp As New ADODB.Recordset
    
    ��ȡ����ID = 0
    Call Record_Locate(mrsIniItems, "����,ID0000")
    gstrSQL = " Select ����ID From �����ʻ� Where ҽ����=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҽ�����˵Ļ�����Ϣ", CStr(mrsIniItems!ֵ), intinsure)
    If Not rsTmp.EOF Then
        ��ȡ����ID = Nvl(rsTmp!����ID, 0)
    End If
End Function

Private Function ����ģʽ(ByVal intinsure As Integer) As Long
    Dim intValue As Integer
    Dim rsTmp As New ADODB.Recordset
    
    '��ȡ����ֵ(0-�Ƚ���,���Ժ;1-�ȳ�Ժ,�����)
    intValue = 0
    gstrSQL = "Select Nvl(����ֵ,0) Value From ���ղ��� Where ����=[1] And ������='����ģʽ'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ֵ", intinsure)
    
    If Not rsTmp.EOF Then
        intValue = rsTmp!Value
    End If
    ����ģʽ = intValue
End Function

Private Function ҽ�������Ѿ���Ժ(ByVal lng����ID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select Nvl(��ǰ״̬,0) ״̬ From �����ʻ� Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж�ҽ�������Ƿ��Ժ", lng����ID)
    
    ҽ�������Ѿ���Ժ = (rsTmp!״̬ = 0)
End Function

Private Function �����Ŀ(ByVal strNO As String, ByVal lng�շ�ϸĿID As Long) As Boolean
    Dim strCode As String, intVerify As Integer
    Dim rsCheck As New ADODB.Recordset
    �����Ŀ = False
    
    intVerify = 0
    strCode = ""
    
    'ȡ������
    gstrSQL = "Select Nvl(����ֵ,'') Value From ���ղ��� Where ������='����������'"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����������")
    With rsCheck
        If Not .EOF Then
            If Not IsNull(!Value) Then
                strCode = !Value
            End If
        End If
    End With
    If Trim(strCode) = "" Then
        MsgBox "�������÷��������ţ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȡ��˱�־
    gstrSQL = "Select nvl(Sfsh00,0) As Value From yydy.Yy_Yydyb0 Where FWWDBH=[1] And YYXMBH=[2]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��˱�־", strCode, lng�շ�ϸĿID)
    With rsCheck
        If Not .EOF Then
            If Not IsNull(!Value) Then
                intVerify = !Value
            End If
        End If
    End With
    If intVerify = 0 Then
        gstrSQL = "Select '['||����||']'||���� ��Ŀ From �շ�ϸĿ Where ID=[1]"
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ŀ����", lng�շ�ϸĿID)
        If strNO <> "" Then
            MsgBox "��Ŀ" & rsCheck!��Ŀ & "��δͨ����ˣ�NO��" & strNO & "�������ܽ��н��������", vbInformation, gstrSysName
        Else
            MsgBox "��Ŀ" & rsCheck!��Ŀ & "��δͨ����ˣ����ܽ��н��������", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    �����Ŀ = True
End Function

Private Function UpdateData(ByVal strColumn As String, ByVal strValue As String) As Boolean
    On Error Resume Next
    
    UpdateData = False
    Call Record_Locate(mrsIniItems, "����," & strColumn)
    With mrsIniItems
        !ֵ = strValue
        .Update
    End With
    
    '20030417     ��������ȡ��
'    If Err <> 0 Then
'        MsgBox "Ӧ���ļ��г���δ֪�Ľӿ���Ŀ��", vbInformation, gstrSysName
'        Exit Function
'    End If
    UpdateData = True
End Function







'------------------------------------�����ǹ��ڼ�¼���Ĺ����뺯��------------------------------------
Private Function Record_Prepare() As Boolean
    '��ʼ�������ڲ�ӳ���¼��
    On Error Resume Next
    Record_Prepare = False
    
    '-----------------------------mrsIniItems-----------------------------
    mstrFields = "����|��������|����|˵��|ֵ|����|�̶���"
    mstrValues = "UserID|�û���||��¼ҽ�����ݿ���û���||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Password|����||�û�������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "connected|����״̬||����״̬||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "fwwdmc|ҽ������||ҽ������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "czyuan|����Ա||����Ա||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Request|����|TRUE or FALSE|����ҵ��ӿ������ļ��Ŀ�ʼ�����־��TRUEʱ��ʾ�����ļ����Կ�ʼ����ȡ||" & ��������.������ & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Reply|��Ӧ|TRUE or FALSE|����ҵ��ӿڷ����ļ��Ļش��־��TRUEʱ��ʾӦ���ļ����Կ�ʼ��ȡ||" & ��������.������ & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Success|�ɹ�|TRUE or FALSE|�����ɹ���||" & ��������.������ & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Error|����|C400|����ʧ��ԭ��||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cardno|����|C12|ҽ��IC����||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "ID0000|ID|C19|ҽ�Ʊ��պ�||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Xming0|����|C8|����||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Xbie00|�Ա�|C1 1��;2Ů ����Ϊ��|�Ա�||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Brnl00|����|N3|����||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Dwmc00|��λ����|C30|��λ����||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Icztmc|IC��״̬|C20|IC��״̬����||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Gzztmc|����״̬|C30|����״̬����||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Dqmc00|��������|C20|Ͷ����������������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Fzxmc0|����������|C20|Ͷ������������������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ghksmc|�Һſ���|C10|�Һſ�������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ghfy00|�Һŷ���|N(5,2)|�Һŷ���||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ghlsh0|�Һ���ˮ��|C16|�Һ���ˮ��||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ghrq00|�Һ�����|C8|�Һ�����||" & ��������.������ & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ghsj00|�Һ�ʱ��|C4|�Һ�ʱ��||" & ��������.ʱ���� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cxlsh0|������ˮ��|C16|������ˮ��||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Grzhye|�����ʻ����|N(8,2)|�����ʻ����||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Bqbm00|���ֱ���|C20|���ֱ���||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cfxms0|�շ���Ŀ��|N(3)|�շ���Ŀ��||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Djlsh0|������ˮ��|C16|������ˮ��||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Mzlsh0|������ˮ��|C16|������ˮ��||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Bckbcs|סԺ����|N(3)|���ο�������(ͬסԺ����)||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Sftsmz|��������|C1 Y��;N��|�Ƿ���������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Sftsbz|���ⲡ��|C1 Y��;N��|�Ƿ����ⲡ��||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Zhzfe0|�ʻ�֧����|N(8,2)|�ʻ�֧����||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Grzfe0|����֧����|N(8,2)|�����ֽ�֧����||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Dbgrzf|�󲡸����Ը�|N(8,2)|�����ֽ�֧����||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Jjzfe0|����֧����|N(8,2)|����֧����||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Bcbxf0|�ܷ���|N(8,2)|�ܷ���||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Sfrq00|�շ�����|C8|�շ�����||" & ��������.������ & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Sfsj00|�շ�ʱ��|C4|�շ�ʱ��||" & ��������.ʱ���� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Sfrxm0|�շѲ���Ա|C8|�շ�������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cxdjh0|��������|C16|�������ݺ�||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ryrq00|��Ժ����|C8|��Ժ����||" & ��������.������ & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Rysj00|��Ժʱ��|C4|��Ժʱ��||" & ��������.ʱ���� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ryksmc|��Ժ����|C10|��Ժ��������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Rylb00|סԺ���|C8 ��ͨ���ͥ����|סԺ���||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ptbcts|��ͨ��������|N10 ��ͨ��������|��ͨ��������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Crbcts|��Ⱦ��������|N10 ��Ⱦ��������|��Ⱦ��������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Zylsh0|סԺ��ˮ��|C16|��Ժ�Ǽ���ˮ��||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Rydjr0|��Ժ����Ա|C8|��Ժ�Ǽ���||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Zyksmc|סԺ����|C10|סԺ��������||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cydjr0|��Ժ�Ǽ���|C8|��Ժ�Ǽ���||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cyrq00|��Ժ����|C8|��Ժ����||" & ��������.������ & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cysj00|��Ժʱ��|C4|��Ժʱ��||" & ��������.ʱ���� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ndfylj|���ҽ�������ۼ�|N8,2|���ҽ�������ۼ�||" & ��������.��ֵ�� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Valid0|�Ƿ��������|True Or False|�Ƿ�Ҫ����Ժ�Ǽǻ��Ƿ���ԹҺ�||" & ��������.������ & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Bnghyy|���ܹҺŵ�ԭ��|C400|���˲��ܹҺŵ�ԭ��||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Bndjyy|������Ժ��ԭ��|C400|���˲�����Ժ�Ǽǵ�ԭ��||" & ��������.�ַ��� & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)

    '-----------------------------mrsIniItems-----------------------------
    mstrFields = "����|��������|����"
    mstrValues = "cydj|��Ժ�Ǽ�|" & "61"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "cydjcx|��Ժ�Ǽǳ���|" & "62"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "cydjsk|��Ժ�Ǽ�ˢ��|" & "63"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "fpxmbm|����Ʊ����Ŀ|" & "70"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "fybfy0|��ҽ������|" & "70"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "login|��¼|" & "11"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "logout|�˳�|" & "12"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzgh|����Һ�|" & "31"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzghcx|����Һų���|" & "32"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzghsk|����Һ�ˢ��|" & "33"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzsf|�����շ�|" & "41"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzsfcx|�����շѳ���|" & "42"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzsfmx|�����շ���ϸ|" & "44"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzsfsk|�����շ�ˢ��|" & "43"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzsfyjs|�����շ�Ԥ����|" & "46"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "query|��¼��ѯ|" & "15"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "rydj|סԺ�Ǽ�|" & "21"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "rydjcx|סԺ�Ǽǳ���|" & "22"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "rydjsk|סԺ�Ǽ�ˢ��|" & "23"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "yb0000|ҽ������|" & "70"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "zysf|סԺ�շ�|" & "51"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "zysfcx|סԺ�շѳ���|" & "52"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "zysfmx|סԺ�շ���ϸ|" & "54"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "zysfsk|סԺ�շ�ˢ��|" & "53"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "zysfyjs|סԺ�շ�Ԥ����|" & "56"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    
    If Err <> 0 Then
        MsgBox "��ʼ���ڲ����ݽṹʱ������δ֪����", vbInformation, gstrSysName
        Exit Function
    End If
    Record_Prepare = True
End Function

Private Sub Record_Clear(ByRef rsObj As ADODB.Recordset, Optional ByVal blnAll As Boolean = False)
    '�����¼����ֵΪ��
    
    With rsObj
        If .RecordCount = 0 Then Exit Sub
        Do While Not .EOF
            If blnAll Then
                !ֵ = ""
                .Update
            Else
                If Nvl(!�̶���, 0) = 0 Then
                    !ֵ = ""
                    .Update
                End If
            End If
            .MoveNext
        Loop
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

End Sub

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '���¼�¼,���������,������
    'strPrimary:�ֶ���,ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, ",")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Private Function Record_Count(ByRef rsObj As ADODB.Recordset, Optional ByVal blnDelete As Boolean = False) As Long
    '�����ܼ�¼��
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Count = 0
    
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        If blnDelete = False Then Record_Count = .RecordCount: Exit Function
        .Filter = "ɾ��=0"
        If .RecordCount = 0 Then Exit Function
        Record_Count = .RecordCount
    End With
End Function



'------------------------------------�����õ��Ĺ����뺯��------------------------------------
'������״̬����ʱ���轫��غ������������Ϊprivate��������ɺ��뻹ԭ�������͹��̵�����
Private Sub ������������(ByVal ������ʽ_IN As Integer, ByVal ����Ŀ��_IN As Integer, ByVal intinsure As Integer)
    Call InitStruc
    Call SendRequest(������ʽ_IN, ����Ŀ��_IN, 0, 0, intinsure)
End Sub

Private Sub ��ʾ��¼������()
    Dim intField As Integer, intFields As Integer, strMsg As String
    With mrsIniItems
        intFields = .Fields.Count - 1
        
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            strMsg = ""
            For intField = 0 To intFields
                strMsg = strMsg & "[" & .Fields(intField).Name & "]" & IIf(IsNull(.Fields(intField).Value), "", .Fields(intField).Value)
            Next
            Debug.Print strMsg
            .MoveNext
        Loop
    End With
End Sub

Private Sub OutputData(Optional ByVal objStream As TextStream, Optional ByVal strData As String)
    '�����ı�����
    If mintStyle = 1 Then
        objStream.WriteLine strData
    Else
        Debug.Print strData
    End If
End Sub


