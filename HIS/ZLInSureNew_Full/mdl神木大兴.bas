Attribute VB_Name = "mdl��ľ����"
Option Explicit
Private mblnInit As Boolean     '�Ƿ��Ѿ���ʼ��

Private Type InitbaseInfor
    ģ������ As Boolean                     '��ǰ�Ƿ���ģ���ȡҽ���ӿ�����
    ҽԺ���� As String                      '��ʼҽԺ����
    ����Ŀ¼ As String
    
End Type


Public InitInfor_��ľ���� As InitbaseInfor
Private Type �������
    IC����              As String
    ����                As String
    �Ա�                As String
    
    ����ID              As Long         '��ǰ����IDֵ
    ����ID              As Long
    ��ǰ�����          As String       'right(����ID1,6)-yyyyMMDDHHMMSS
    �����ܶ�            As Double
    
    �������            As Boolean  '�Ѿ��������,�����
End Type


Private Type ��������
    �����ܶ�          As Double       '�����סԺ
    �����ʻ�֧��    As Double       '�����סԺ
    ͳ��֧��        As Double       '�����סԺ
    ����Ա����      As Double       'סԺ
    Ѻ���ܶ�        As Double       'סԺ
    Ӧ���ֽ��      As Double       'סԺ
    ���Ѵ�λ��      As Double       'סԺ
    �ԷѴ�λ��      As Double       'סԺ
    ���ѵ��·�      As Double       'סԺ
    �Էѵ��·�      As Double       'סԺ
    ����ǰ�������  As Double       '����
    ����������  As Double       '����
End Type
Private g�������� As ��������

Public g�������_��ľ���� As �������
Public gcnOracle_��ľ���� As ADODB.Connection     '�м������

Public Function ҽ����ʼ��_��ľ����() As Boolean
    
    Dim strReg As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim strUser As String, strPass As String, strServer As String
    
    If mblnInit = True Then
        ҽ����ʼ��_��ľ���� = True
        Exit Function
    End If
    
    '��ʼģ��ӿ�
    Call GetRegInFor(g����ģ��, "����", "ģ��ӿ�", strReg)
    If Val(strReg) = 1 Then
        InitInfor_��ľ����.ģ������ = True
    Else
        InitInfor_��ľ����.ģ������ = False
    End If
   
    InitInfor_��ľ����.ҽԺ���� = gstrҽԺ����
    InitInfor_��ľ����.����Ŀ¼ = GetSetting("ZLSOFT", "����ģ��\zl9Insure", UCase("����Ŀ¼"), App.Path) & "\ReadYbInfo.INI"
    
    
    If Open�м��_��ľ���� = False Then
        Exit Function
    End If
    mblnInit = True
    ҽ����ʼ��_��ľ���� = True
End Function

Public Function ҽ����ֹ_��ľ����() As Boolean
    
    '����ʼ����־��Ϊfalse
    mblnInit = False
    If gcnOracle_��ľ����.State = 1 Then
        gcnOracle_��ľ����.Close
    End If
    ҽ����ֹ_��ľ���� = True
End Function

Public Function ��ݱ�ʶ_��ľ����(Optional bytType As Byte, Optional lng����ID As Long) As String
    '���ܣ�ʶ��ָ����Ա�Ƿ�Ϊ�α����ˣ����ز��˵���Ϣ
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    '���أ��ջ���Ϣ��
    Err = 0
    On Error GoTo errHand:
    ��ݱ�ʶ_��ľ���� = frmIdentify��ľ����.GetPatient(bytType, lng����ID)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    ��ݱ�ʶ_��ľ���� = ""
End Function


Public Function �������_��ľ����(ByVal lng����ID As Long, ByRef dbl͸֧�� As Currency) As Currency
    '����: ��ȡ�α����˸����ʻ����
    '����: ���ظ����ʻ����
    dbl͸֧�� = 10000000000000#
    �������_��ľ���� = 0
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
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ʻ�֧��", g�������_��ľ����.����ID)
    If Not rsTemp.EOF Then
        ��ȡ�����ʻ�֧�� = Nvl(rsTemp!��Ԥ��, 0)
    End If
End Function

Private Function Checkҽ����Ŀ(ByVal lng�շ�ϸĿID As Long, ByRef str��� As String, ByRef strҽ������ As String, ByRef strҽ������ As String, ByRef strƴ������ As String, Optional bln���� As Boolean = False) As Boolean
    '����:��ȡ��ص�ҽ����Ŀ��Ϣ
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "" & _
              "   Select a.��ע,a.��Ŀ����,a.��Ŀ����,b.����,b.���� " & _
              "   From ����֧����Ŀ a,�շ�ϸĿ B  " & _
              "   where a.�շ�ϸĿid=b.ID and  a.����=" & TYPE_�������� & _
              "           and a.�շ�ϸĿid=" & lng�շ�ϸĿID
    
    Checkҽ����Ŀ = False
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ŀ"
    If rsTemp.EOF Then
        If bln���� = False Then
                gstrSQL = "Select ����,���� From �շ�ϸĿ where ID=" & lng�շ�ϸĿID
                zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ������Ŀ"
                ShowMsgbox "�շ���Ŀ��" & Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����) & "����δ����ҽ�����룬���ܼ�������!"
        End If
        Exit Function
    End If
    str��� = Nvl(rsTemp!��ע)
    strҽ������ = Nvl(rsTemp!��Ŀ����)
    strҽ������ = Nvl(rsTemp!��Ŀ����)

    gstrSQL = "select pybm from yy_ypfzb  where lb='" & str��� & "' and bm='" & strҽ������ & "' and mc='" & strҽ������ & "'"
    
    Call OpenRecordset_��ľ����(rsTemp, "��ȡƴ����", gstrSQL)
    If Not rsTemp.EOF Then
        strƴ������ = Nvl(rsTemp!pybm)
    Else
        strƴ������ = ""
    End If
    Checkҽ����Ŀ = True
End Function
Public Function �����������ȡ��_��ľ����(ByVal bytType As Byte, ByVal lng����ID As Long) As Boolean
   '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ����ť
    '--�����:
    '--������:
    '--��  ��:�ɹ�,true
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim blnYes As Boolean
    
    �����������ȡ��_��ľ���� = False
    If g�������_��ľ����.������� = False Then Exit Function
    
    gstrSQL = "Select YBKH,CFBH,YXBZ From MZ_FYMXB where YBKH='" & g�������_��ľ����.IC���� & "' and rownum<=1 "
    Call OpenRecordset_��ľ����(rsTemp, "����Ƿ��Ѿ������", gstrSQL)
    If rsTemp.RecordCount <> 0 Then
        ShowMsgbox "�ò���ҽ���Ѿ��������,��������Ϊ��ȡ���˸ò���," & vbCrLf & "���������������ҽԺ���ݲ���,���Ҫǿ���˳���?", True, blnYes
        If Not blnYes Then
            Exit Function
        End If
    End If
    �����������ȡ��_��ľ���� = True

End Function
Public Function �����������_��ľ����(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String) As Boolean
    '������rsDetail     ������ϸ(����)
    '      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
    '�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��

    Dim str��ϸ As String
    Dim rsTemp As New ADODB.Recordset
    Dim str��� As String, strҽ������ As String, strҽ������ As String, strƴ���� As String
    
    
    g�������_��ľ����.�����ܶ� = 0
    
    If g�������_��ľ����.������� = True Then
        ShowMsgbox "�Ѿ�����������,�밴���㰴ť!"
        Exit Function
    End If
    
    str��ϸ = ""
    If rs��ϸ.RecordCount <> 0 Then
        g�������_��ľ����.��ǰ����� = Lpad(Right(CStr(rs��ϸ!����ID), 6), 6, "0") & Format(zlDatabase.Currentdate, "yyyymmddHHMMSS") 'right(����ID1,6)-yyyyMMDDHHMMSS
    Else
        g�������_��ľ����.��ǰ����� = ""
    End If
    
    
    '��һ��:�ж��Ƿ����δ�������ϸ����
    
    Err = 0: On Error GoTo errHand:
    gcnOracle_��ľ����.BeginTrans
    
    DebugTool "�����������,��һ��:�ж��Ƿ����δ�������ϸ����"
    
    gstrSQL = "Select YBKH,CFBH,YXBZ From MZ_FYMXB where YBKH='" & g�������_��ľ����.IC���� & "' and rownum<=1 "
    Call OpenRecordset_��ľ����(rsTemp, "����Ƿ��Ѿ������", gstrSQL)
    If rsTemp.RecordCount <> 0 Then
        gstrSQL = "Select YBKH,CFBH,YXBZ From MZ_FYMXB where YBKH='" & g�������_��ľ����.IC���� & "' and YXBZ<>'F' and rownum<=1 "
        Call OpenRecordset_��ľ����(rsTemp, "����Ƿ��Ѿ������", gstrSQL)
        If Not rsTemp.EOF Then
            Dim blnYes As Boolean
            ShowMsgbox "��ҽ�������ϴ��Ѿ��ύ����ϸ,��δ���," & vbCrLf & "(������" & Nvl(rsTemp!cfbh) & "),����ԭ������:" & vbCrLf & "    1.���ܲ���Ա�ڽ����������;��ֹ��!" & vbCrLf & "    2.�������ϴ���ϸ��ɺ�,������ַ���ʽ�˳�!" & vbCrLf & "    3.�����Ѿ����������,��HIS��δ��ʽ����!" & vbCrLf & " �Ƿ�Ҫ����ǿ�����?", True, blnYes
            If blnYes = False Then
                gcnOracle_��ľ����.RollbackTrans
                Exit Function
            End If
        End If
        gstrSQL = "ZL_����_Clear('" & g�������_��ľ����.IC���� & "')"
        ExecuteProcedure_��ľ���� "�������"
    End If
    
    
    
    '�ڶ���:�ϴ���ϸ����
     DebugTool "�����������,�ڶ���:�ϴ���ϸ����"
    
    With rs��ϸ
        If rs��ϸ.RecordCount = 0 Then ShowMsgbox "δ������صķ��ü�¼!": Exit Function
        Do While Not .EOF
        
            '�жϱ���
            If Checkҽ����Ŀ(Nvl(!�շ�ϸĿID, 0), str���, strҽ������, strҽ������, strƴ����) = False Then
                gcnOracle_��ľ����.RollbackTrans
                Exit Function
            End If
            '��ǰ�����          As String       'right(����ID1,6)-yyyyMMDDHHMMSS
            '��ϸ����:ҽ������_IN����ˮ��_IN���������_IN�����_IN��ҽ������_IN��ҽ������_IN��ƴ������_IN���۸�IN������_IN����Ч��־_IN
            
            gstrSQL = "ZL_MZ_FYMXB_INSERT("
            gstrSQL = gstrSQL & "'" & g�������_��ľ����.IC���� & "',"       'ҽ������_IN
            gstrSQL = gstrSQL & rs��ϸ.AbsolutePosition & ","   '��ˮ��_IN
            gstrSQL = gstrSQL & "'" & g�������_��ľ����.��ǰ����� & "',"    '�������_IN
            gstrSQL = gstrSQL & "'" & str��� & "',"    '���_IN
            gstrSQL = gstrSQL & "'" & strҽ������ & "',"   'ҽ������_IN
            gstrSQL = gstrSQL & "'" & strҽ������ & "',"   'ҽ������_IN
            If strƴ���� = "" Then
                gstrSQL = gstrSQL & "'" & zlCommFun.zlGetSymbol(strҽ������, 0) & "',"     'ƴ������_IN
            Else
                gstrSQL = gstrSQL & "'" & strƴ���� & "',"     'ƴ������_IN
            End If
            gstrSQL = gstrSQL & Format(Nvl(!ʵ�ս��, 0) / Nvl(!����, 0), "####0.0000;-####0.0000;0;0") & "," '�۸�IN
            gstrSQL = gstrSQL & Format(Nvl(!����, 0), "####0.00;-####0.00;0;0") & ","      '����_IN
            gstrSQL = gstrSQL & "'F')"                                                              '��Ч��־_IN :ҽԺMIS��ֵΪF,ֻ����ҽ���޸����־����ҽ������������ɺ�����ΪT,�쳣����ʱ������ֵΪX�Ա�ҽԺMIS��ѯ?ҽԺMIS��Ȩ�޸����־���������Ը�
            DebugTool gstrSQL
            ExecuteProcedure_��ľ���� "������ϸ��Ϣд��"
            g�������_��ľ����.�����ܶ� = g�������_��ľ����.�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
    
    'д������Ϣ:
    '  ����:  YBKH_IN IN MZ_JSLSB.YBKH%TYPE,--ҽ������     "
    '    CFBH_IN IN MZ_JSLSB.CFBH%TYPE,--�������
    '    FYHJ_IN IN MZ_JSLSB.FYHJ%TYPE,--���úϼ�
    '    XM_IN IN MZ_JSLSB.XM%TYPE   --����
    g�������_��ľ����.�����ܶ� = Round(g�������_��ľ����.�����ܶ�, 2)
    
    gstrSQL = "ZL_MZ_JSLSB_INSERT("
    gstrSQL = gstrSQL & "'" & g�������_��ľ����.IC���� & "',"
    gstrSQL = gstrSQL & "'" & g�������_��ľ����.��ǰ����� & "')"
    
    ExecuteProcedure_��ľ���� "���������Ϣд��"
   DebugTool "�����������,д������Ϣ"
    
    gcnOracle_��ľ����.CommitTrans
    
    
    '������:�ȴ�������Ϣ
     DebugTool "�����������,������:�ȴ�������Ϣ"
    
    If frm����ȴ�_��ľ����.ShowWait(0, g�������_��ľ����.IC����) = False Then
        gcnOracle_��ľ����.BeginTrans
        gstrSQL = "ZL_����_Clear('" & g�������_��ľ����.IC���� & "')"
        ExecuteProcedure_��ľ���� "�������"
        gcnOracle_��ľ����.CommitTrans
        Exit Function
    End If
        
    gstrSQL = "" & _
       "   Select  ybkh ҽ������, cfbh �������, jssj ����ʱ��, jsbz ҽ�������־, " & _
       "           fyhj �����ܷ���, kszf ����֧��, tczf ͳ��֧��, ybje Ӧ���ֽ��, xm ��������,jsqksye ����ǰ�������,jshksye ���������� " & _
       "   From MZ_JSLSB  " & _
       "   Where ybkh='" & g�������_��ľ����.IC���� & "' and jsbz='T'"
    
    OpenRecordset_��ľ���� rsTemp, "��ȡ������Ϣ", gstrSQL
    str���㷽ʽ = ""
    With g��������
        .�����ʻ�֧�� = Format(Nvl(rsTemp!����֧��, 0), "####0.00;-####0.00;0;0")
        .ͳ��֧�� = Format(Nvl(rsTemp!ͳ��֧��, 0), "####0.00;-####0.00;0;0")
        .Ӧ���ֽ�� = Format(Nvl(rsTemp!Ӧ���ֽ��, 0), "####0.00;-####0.00;0;0")
        .�����ܶ� = Format(Nvl(rsTemp!�����ܷ���, 0), "####0.00;-####0.00;0;0")
        .����ǰ������� = Format(Nvl(rsTemp!����ǰ�������, 0), "####0.00;-####0.00;0;0")
        .���������� = Format(Nvl(rsTemp!����������, 0), "####0.00;-####0.00;0;0")
        .���Ѵ�λ�� = 0
        .���ѵ��·� = 0
        .����Ա���� = 0
        .Ѻ���ܶ� = 0
        .�ԷѴ�λ�� = 0
        .�Էѵ��·� = 0
        str���㷽ʽ = "�����ʻ�;" & .�����ʻ�֧�� & ";0"
        str���㷽ʽ = str���㷽ʽ & "|" & "ͳ��֧��;" & .ͳ��֧�� & ";0"
    End With
    DebugTool "�����������ɹ�,���㷽ʽ��" & str���㷽ʽ
    
    g�������_��ľ����.������� = True
    �����������_��ľ���� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle_��ľ����.RollbackTrans
End Function
Public Function �������_��ľ����(lng����ID As Long, cur�����ʻ� As Currency, strҽ���� As String, curȫ�Ը� As Currency) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur֧�����   �Ӹ����ʻ���֧���Ľ��
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ��������Ǳ�֤�˸����ʻ���������ڸ����ʻ�����˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
        '��ʱ�����շ�ϸĿ��Ȼ�ж�Ӧ��ҽ������
    Dim lng����ID  As Long
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHandle

    Call DebugTool("�����������")

    gstrSQL = "" & _
        "   Select a.*,a.����*a.���� as ����,a.ʵ�ս��/(nvl(a.����,1)*nvl(a.����,1)) as ���� " & _
        "   From ������ü�¼ a " & _
        "   Where ����ID=[1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0"

    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ϸ��¼", lng����ID)

    If rs��ϸ.EOF = True Then
        Err.Raise 9000 + vbExclamation, gstrSysName, "û����д�շѼ�¼!"
        Exit Function
    End If
    
    lng����ID = rs��ϸ("����ID")


    If g�������_��ľ����.����ID <> lng����ID Then
        Err.Raise 9000, gstrSysName, "�ò��˻�û�о��������֤�����ܽ���ҽ�����㡣"
        Exit Function
    End If
    g�������_��ľ����.����ID = lng����ID
    
    Dim dbl�����ܶ� As Double
    
    dbl�����ܶ� = 0
    '��һ��:���ܷ���
    DebugTool "�������,��һ��:���ܷ���,�����ϱ�־"
    With rs��ϸ
        If rs��ϸ.RecordCount = 0 Then ShowMsgbox "δ������صķ��ü�¼!": Exit Function
        Do While Not .EOF
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & g�������_��ľ����.��ǰ����� & "')"
            DebugTool "     ������ϸ��־:SQL=" & gstrSQL
            zlDatabase.ExecuteProcedure gstrSQL, "�����ϴ���־"
            DebugTool " ������ϸ��־:���²��˷��ü�¼�ɹ�:SQL=" & gstrSQL
            dbl�����ܶ� = dbl�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With
    
    If Format(dbl�����ܶ�, "#####0.00;-####0.00;0;0") <> Format(g��������.�����ܶ�, "#####0.00;-####0.00;0;0") Then
        Err.Raise 9000, gstrSysName, "�����ܶ��,���ܽ���!" & vbCrLf & _
                " �����������ܶ�:" & Format(g��������.�����ܶ�, "#####0.00;-####0.00; ;") & vbCrLf & _
                " ��ʽ��������ܶ�:" & Format(dbl�����ܶ�, "#####0.00;-####0.00; ;")
        Exit Function
    End If
    
    
    '�ڶ���:���������Ϣ
    DebugTool "�ڶ���:����ʼ���汣�ս����¼"

   '���뱣�ս����¼
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(����ǰ�������),�ۼ�ͳ�ﱨ��_IN(����������),סԺ����_IN(סԺ:��ҳid),����(Ѻ���ܶ�),�ⶥ��_IN(���Ѵ�λ��),ʵ������_IN(�ԷѴ�λ��),
    '   �������ý��_IN(�����ܽ��),ȫ�Ը����_IN(���ѵ��·�),�����Ը����_IN(Ӧ���ֽ��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(����Ա����),�����Ը����_IN(�Էѵ��·�),�����ʻ�֧��_IN(�����ʻ�֧�����),"
    '   ֧��˳���_IN(����:������),��ҳID_IN(��ҳid),��;����_IN,��ע_IN()
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�������� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "NULL,NULL," & g��������.����ǰ������� & "," & g��������.���������� & ",null,0,0,0," & _
            g��������.�����ܶ� & ",0," & g��������.Ӧ���ֽ�� & "," & _
            g��������.ͳ��֧�� & " ," & g��������.ͳ��֧�� & ",0,0," & g��������.�����ʻ�֧�� & ",'" & _
             g�������_��ľ����.��ǰ����� & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
  
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & lng����ID & "," & TYPE_�������� & ",'�ʻ����','" & g��������.���������� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�������һ�εĿ������")
  
    '������:����м�������Ϣ
    DebugTool "������:����м�������Ϣ"
    
    gstrSQL = "ZL_����_Clear('" & g�������_��ľ����.IC���� & "')"
    ExecuteProcedure_��ľ���� "�������"
    
    �������_��ľ���� = True
    g�������_��ľ����.������� = False
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function ����������_��ľ����(lng����ID As Long, cur�����ʻ� As Currency, lng����ID As Long) As Boolean
    '���ܣ��������շѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID     �շѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
    '      cur�����ʻ�   �Ӹ����ʻ���֧���Ľ��

    Dim intMouse As Integer
    Dim lng����ID  As Long
    Dim str��� As String, strҽ������ As String, strҽ������ As String, strƴ���� As String
    Dim rs��ϸ As New ADODB.Recordset
    Dim rsԭ��ϸ As New ADODB.Recordset

    Dim rsTemp As New ADODB.Recordset
    Dim lng����id1 As Long
On Error GoTo errHand:

    ����������_��ľ���� = False

    '�����֤
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If ��ݱ�ʶ_��ľ����(2, lng����id1) = "" Then
        If lng����id1 = 0 Then
            Err.Raise 9000, gstrSysName, "�㲻�ǵ�ǰ�ֿ���!"
            Screen.MousePointer = intMouse
            Exit Function
        End If
    End If
    If lng����ID <> lng����id1 Then
        Screen.MousePointer = intMouse
        Err.Raise 9000, gstrSysName, "�㲻�ǵ�ǰ�ֿ���!"
        Exit Function
    End If

    Err = 0:
    Screen.MousePointer = intMouse

    gcnOracle_��ľ����.BeginTrans


    '��һ��:ȷ������IDֵ
    DebugTool "��һ��:ȷ������IDֵ"
    gstrSQL = "select distinct A.����ID from ������ü�¼ A,������ü�¼ B " & _
              " where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=2 and B.����ID=" & lng����ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ҽ��", lng����ID)
    lng����ID = rsTemp("����ID")

    '�ڶ���:ȷ��������ԭʼ���ݵ���ϸ��¼
    DebugTool "ȷ��������ԭʼ���ݵ���ϸ��¼"

    gstrSQL = "Select * From ������ü�¼ " & _
        " Where ����ID=" & lng����ID & " And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0 and nvl(ʵ�ս��,0)<>0"

    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������¼")
    g�������_��ľ����.�����ܶ� = 0
    g�������_��ľ����.����ID = lng����ID
     


    gstrSQL = "Select * From ������ü�¼ where  ����ID = [1] And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0 and nvl(ʵ�ս��,0)<>0"
    Set rsԭ��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������¼", lng����ID)
    g�������_��ľ����.��ǰ����� = Nvl(rsԭ��ϸ!ժҪ)



    '������:��ԭʼ��¼�еĽ�Ҫ��Ϊ�������ݵĽ�Ҫ���������ϴ���־
    DebugTool "������:��ԭʼ��¼�еĽ�Ҫ��Ϊ�������ݵĽ�Ҫ���������ϴ���־"
    With rs��ϸ
        Do While Not .EOF
        
            '�жϱ���
            If Checkҽ����Ŀ(Nvl(!�շ�ϸĿID, 0), str���, strҽ������, strҽ������, strƴ����) = False Then
                gcnOracle_��ľ����.RollbackTrans
                Exit Function
            End If
        
            '��ǰ�����          As String       'right(����ID1,6)-yyyyMMDDHHMMSS
            '��ϸ����:ҽ������_IN����ˮ��_IN���������_IN�����_IN��ҽ������_IN��ҽ������_IN��ƴ������_IN���۸�IN������_IN����Ч��־_IN
            
            gstrSQL = "ZL_MZ_FYMXB_INSERT("
            gstrSQL = gstrSQL & "'" & g�������_��ľ����.IC���� & "',"       'ҽ������_IN
            gstrSQL = gstrSQL & .AbsolutePosition & ","   '��ˮ��_IN
            gstrSQL = gstrSQL & "'" & g�������_��ľ����.��ǰ����� & "',"    '�������_IN
            gstrSQL = gstrSQL & "'" & str��� & "',"    '���_IN
            gstrSQL = gstrSQL & "'" & strҽ������ & "',"   'ҽ������_IN
            gstrSQL = gstrSQL & "'" & strҽ������ & "',"   'ҽ������_IN
            If strƴ���� = "" Then
                gstrSQL = gstrSQL & "'" & zlCommFun.zlGetSymbol(strҽ������, 0) & "',"     'ƴ������_IN
            Else
                gstrSQL = gstrSQL & "'" & strƴ���� & "',"     'ƴ������_IN
            End If
            gstrSQL = gstrSQL & Format(Nvl(!ʵ�ս��, 0) / (Nvl(!����, 1) * Nvl(!����, 1)), "####0.0000;-####0.0000;0;0") & "," '�۸�IN
            gstrSQL = gstrSQL & Format((Nvl(!����, 1) * Nvl(!����, 1)), "####0.00;-####0.00;0;0") & ","      '����_IN
            gstrSQL = gstrSQL & "'F')"                                                              '��Ч��־_IN :ҽԺMIS��ֵΪF,ֻ����ҽ���޸����־����ҽ������������ɺ�����ΪT,�쳣����ʱ������ֵΪX�Ա�ҽԺMIS��ѯ?ҽԺMIS��Ȩ�޸����־���������Ը�
            DebugTool gstrSQL
            ExecuteProcedure_��ľ���� "������ϸ��Ϣд��"
            
            'д�ϴ���־
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & g�������_��ľ����.��ǰ����� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "�����ϴ���־")
            g�������_��ľ����.�����ܶ� = g�������_��ľ����.�����ܶ� + Nvl(!ʵ�ս��, 0)
            .MoveNext
        Loop
    End With

    
    '���Ĳ�:��������ؼ�¼
    DebugTool "���Ĳ�:������ؼ�¼"
    
    gstrSQL = "Select * from ���ս����¼ where ����=1 and ��¼id=" & lng����ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���ĵ��ݺ�"


   '���뱣�ս����¼
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(),סԺ����_IN(סԺ:��ҳid),����(Ѻ���ܶ�),�ⶥ��_IN(���Ѵ�λ��),ʵ������_IN(�ԷѴ�λ��),
    '   �������ý��_IN(�����ܽ��),ȫ�Ը����_IN(���ѵ��·�),�����Ը����_IN(Ӧ���ֽ��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(����Ա����),�����Ը����_IN(�Էѵ��·�),�����ʻ�֧��_IN(�����ʻ�֧�����),"
    '   ֧��˳���_IN(����:������),��ҳID_IN(��ҳid),��;����_IN,��ע_IN()
    
    gstrSQL = "zl_���ս����¼_insert(1," & lng����ID & "," & TYPE_�������� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL,NULL,null,null," & -1 * Nvl(rsTemp!����, 0) & "," & -1 * Nvl(rsTemp!�ⶥ��, 0) & "," & -1 * Nvl(rsTemp!ʵ������, 0) & "," & _
           -1 * Nvl(rsTemp!�������ý��, 0) & "," & -1 * Nvl(rsTemp!ȫ�Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & _
           -1 * Nvl(rsTemp!����ͳ����, 0) & "," & -1 * Nvl(rsTemp!ͳ�ﱨ�����, 0) & "," & -1 * Nvl(rsTemp!���Ը����, 0) & "," & -1 * Nvl(rsTemp!�����Ը����, 0) & "," & -1 * Nvl(rsTemp!�����ʻ�֧��, 0) & ",'" & _
           Nvl(rsTemp!֧��˳���, 0) & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
    gcnOracle_��ľ����.CommitTrans
    
    ����������_��ľ���� = True
    Exit Function
errHand:
    gcnOracle_��ľ����.RollbackTrans
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function


Public Function ��Ժ�Ǽ�_��ľ����(lng����ID As Long, lng��ҳID As Long, ByRef strҽ���� As String) As Boolean
    '���ܣ�����Ժ�Ǽ���Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    
    
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select to_char(��Ժ����,'yyyy-mm-dd hh24:mi:ss') as ��Ժ���� From ������ҳ where ����id= " & lng����ID & " and ��ҳid=" & lng��ҳID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��Ժ����"
     
   gcnOracle_��ľ����.BeginTrans
    
    '���̲���
    '   ҽ������_IN,סԺ���_IN,��Ժʱ��_IN,�����־_IN
    gstrSQL = "ZL_ZY_JSLSB_INSERT("
    gstrSQL = gstrSQL & "'" & g�������_��ľ����.IC���� & "',"
    gstrSQL = gstrSQL & "'" & lng����ID & "_" & lng��ҳID & "',"
    gstrSQL = gstrSQL & "to_date('" & Nvl(rsTemp!��Ժ����) & "','yyyy-mm-dd hh24:mi:ss'),"
    gstrSQL = gstrSQL & "'F')"
    ExecuteProcedure_��ľ���� "���²�����Ժ"
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
   gcnOracle_��ľ����.CommitTrans
    ��Ժ�Ǽ�_��ľ���� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle_��ľ����.RollbackTrans
    ��Ժ�Ǽ�_��ľ���� = False
End Function

Public Function ��Ժ�Ǽǳ���_��ľ����(lng����ID As Long, lng��ҳID As Long) As Boolean
  '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ����û�������ã������Ժ�Ǽǳ����ӿڣ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
            
    '���˺�:20040923���ӵ�
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
     Err = 0
    On Error GoTo errHand
    
    DebugTool "������Ժ�ǳ����ӿ�"
    
    ��Ժ�Ǽǳ���_��ľ���� = False
    
    If ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "����δ����ã����ܳ�����Ժ�Ǽ�"
        Exit Function
    End If
    gstrSQL = "Select ҽ���� From �����ʻ� where ����id=" & lng����ID & " and ����=" & TYPE_��������
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���յ������Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�����ڸ�ҽ������!"
        Exit Function
    End If

    
    '����Ϊ:ҽ������_IN
    gstrSQL = "ZL_ZY_JSLSB_DELETE("
    gstrSQL = gstrSQL & "'" & Nvl(rsTemp!ҽ����) & "')"
    ExecuteProcedure_��ľ���� "ɾ����Ժ��Ϣ"
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������Ժ�Ǽ�")
        
    '����ҽ���ʻ�
    DebugTool "ȡ���ɹ�"
    ��Ժ�Ǽǳ���_��ľ���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ��Ժ�Ǽ�_��ľ����(lng����ID As Long, lng��ҳID As Long) As Boolean
    '���ܣ�����Ժ��Ϣ����ҽ��ǰ�÷�����ȷ�ϣ�����ֻ��Գ�����Ժ�Ĳ��ˣ�������������Լ�
    '������lng����ID-����ID��lng��ҳID-��ҳID
    '���أ����׳ɹ�����true�����򣬷���false
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    
    Err = 0:    On Error GoTo errHand:
    ��Ժ�Ǽ�_��ľ���� = False
    
    If Not ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "��ǰ���˲�����δ����ã�������Ժ��������"
        Exit Function
    End If
    
    gstrSQL = "Select ҽ���� From �����ʻ� where ����id=" & lng����ID & " and ����=" & TYPE_��������
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���յ������Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�����ڸ�ҽ������!"
        Exit Function
    End If
    
    '����:ҽ������_IN,�����־_IN
    gstrSQL = "ZL_ZY_JSLSB_UPDATE("
    gstrSQL = gstrSQL & "'" & Nvl(rsTemp!ҽ����) & "',"
    gstrSQL = gstrSQL & "'T')"
    ExecuteProcedure_��ľ���� "���²��˳�Ժ��־"
    '�ı䵱ǰ״̬
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    ��Ժ�Ǽ�_��ľ���� = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ��Ժ�Ǽǳ���_��ľ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
  '��Ժ�Ǽǳ���
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant
    
    ��Ժ�Ǽǳ���_��ľ���� = False
    
    Err = 0: On Error GoTo errHand:
     
     If Not ����δ�����(lng����ID, lng��ҳID) Then
        ShowMsgbox "�ò����Ѿ���Ժ������,������ȡ����Ժ!"
        Exit Function
     End If
    
    gstrSQL = "Select ҽ���� From �����ʻ� where ����id=" & lng����ID & " and ����=" & TYPE_��������
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ���յ������Ϣ"
    If rsTemp.EOF Then
        ShowMsgbox "�����ڸ�ҽ������!"
        Exit Function
    End If
    
    '����:ҽ������_IN,�����־_IN
    gstrSQL = "ZL_ZY_JSLSB_UPDATE("
    gstrSQL = gstrSQL & "'" & Nvl(rsTemp!ҽ����) & "',"
    gstrSQL = gstrSQL & "'F')"
    ExecuteProcedure_��ľ���� "���²��˳�Ժ��־"
    
    gstrSQL = "zl_�����ʻ�_��Ժ(" & lng����ID & "," & TYPE_�������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ժ�Ǽ�")
    ��Ժ�Ǽǳ���_��ľ���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function GetMax��ˮ��(ByVal str���� As String) As Long
    Dim rsTemp As New ADODB.Recordset
    
    '��ȡ�����ˮ��
    gstrSQL = "Select nvl(max(ID),0)+1  as ��� From ZY_FYMXB where YBKH='" & str���� & "'"
    OpenRecordset_��ľ���� rsTemp, "��ȡ����", gstrSQL
    If rsTemp.EOF Then
        GetMax��ˮ�� = 1
    Else
       GetMax��ˮ�� = Nvl(rsTemp!���, 1)
    End If
    
    
End Function
Public Function �����Ǽ�_��ľ����(ByVal lng��¼���� As Long, ByVal lng��¼״̬ As Long, ByVal str���ݺ� As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�ϴ�������ϸ����
    '--�����:
    '--������:
    '--��  ��:�ϴ��ɹ�����True,����False
    '-----------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str��ˮ�� As String
    Dim str��� As String, strҽ������ As String, strҽ������ As String, strƴ���� As String


    �����Ǽ�_��ľ���� = False


   '�������ŵ��ݵķ�����ϸ
   gstrSQL = "" & _
              "  Select A.*,M.ҽ����" & _
              "  From סԺ���ü�¼ A,������ҳ C,�����ʻ� M" & _
              "  where a.NO=[1] and A.��¼����=[2] and A.��¼״̬ = [3]" & _
              "        and A.����ID=C.����ID and nvl(a.ʵ�ս��,0)<>0 and A.��ҳID=C.��ҳID  And a.����id=M.����id and M.����=[4] and  C.����=[4]" & _
              "  Order by A.����ID,A.NO,A.����ʱ��"
    
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "������ϸ�ϴ�", str���ݺ�, lng��¼����, lng��¼״̬, TYPE_��������)
    Err = 0:    On Error GoTo errHand:
    With rs��ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
        
            If Checkҽ����Ŀ(Nvl(!�շ�ϸĿID, 0), str���, strҽ������, strҽ������, strƴ����) = False Then Exit Function
                        
            '��ϸ����:ҽ������_IN,��ˮ��_IN,סԺ���_IN,���_IN,ҽ������_IN,ҽ������_IN,ƴ������_IN,�۸�_IN,����_IN,��Ч��־_IN
            str��ˮ�� = GetMax��ˮ��(Nvl(!ҽ����))
            gstrSQL = "ZL_ZY_FYMXB_INSERT("
            gstrSQL = gstrSQL & "'" & Nvl(!ҽ����) & "',"       'ҽ������_IN
            gstrSQL = gstrSQL & str��ˮ�� & ","   '��ˮ��_IN
            gstrSQL = gstrSQL & "'" & Nvl(!����ID, 0) & "_" & Nvl(!��ҳID, 0) & "',"   'סԺ���_IN
            gstrSQL = gstrSQL & "'" & str��� & "',"    '���_IN
            gstrSQL = gstrSQL & "'" & strҽ������ & "',"   'ҽ������_IN
            gstrSQL = gstrSQL & "'" & strҽ������ & "',"   'ҽ������_IN
            If strƴ���� = "" Then
                gstrSQL = gstrSQL & "'" & zlCommFun.zlGetSymbol(strҽ������, 0) & "',"     'ƴ������_IN
            Else
                gstrSQL = gstrSQL & "'" & strƴ���� & "',"     'ƴ������_IN
            End If
            gstrSQL = gstrSQL & Format(Nvl(!ʵ�ս��, 0) / (Nvl(!����, 1) * Nvl(!����, 1)), "####0.0000;-####0.0000;0;0") & "," '�۸�IN
            gstrSQL = gstrSQL & Format((Nvl(!����, 1) * Nvl(!����, 1)), "####0.00;-####0.00;0;0") & ","   '����_IN
            gstrSQL = gstrSQL & "'F')"                                                              '��Ч��־_IN :ҽԺMIS��ֵΪF,ֻ����ҽ���޸����־����ҽ������������ɺ�����ΪT,�쳣����ʱ������ֵΪX�Ա�ҽԺMIS��ѯ?ҽԺMIS��Ȩ�޸����־���������Ը�
            DebugTool gstrSQL
            ExecuteProcedure_��ľ���� "סԺ��ϸ��Ϣд��"
            gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str��ˮ�� & "')"
            zlDatabase.ExecuteProcedure gstrSQL, "������ϸ����"
            .MoveNext
        Loop
    End With
    �����Ǽ�_��ľ���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function ����סԺ��ϸ��¼_��ľ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional bln���� As Boolean = False) As Boolean

    '���������ϸ��¼
    Dim cnTemp As New ADODB.Connection
    Dim rs��ϸ As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str��ˮ�� As String
    Dim str��� As String, strҽ������ As String, strҽ������ As String, strƴ���� As String
    
    Err = 0:    On Error GoTo errHand:
      
      
    ����סԺ��ϸ��¼_��ľ���� = False
    
   gstrSQL = "" & _
              "  Select A.*,M.ҽ����" & _
              "  From סԺ���ü�¼ A,������ҳ C,�����ʻ� M" & _
              "  where Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(���ӱ�־,0)<>9  and a.����id=[1] and A.��ҳid= [2]" & _
              "        and  A.����ID=C.����ID and nvl(a.ʵ�ս��,0)<>0 and A.��ҳID=C.��ҳID  And a.����id=M.����id and M.����=[3] and  C.����=[3]" & _
              "  Order by A.����ID,A.NO,A.����ʱ��"
    Set rs��ϸ = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID, lng��ҳID, TYPE_��������)

   With rs��ϸ
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            
            If Checkҽ����Ŀ(Nvl(!�շ�ϸĿID, 0), str���, strҽ������, strҽ������, strƴ����, bln����) = False Then
                If bln���� = False Then Exit Function
            Else
                        
                '��ϸ����:ҽ������_IN,��ˮ��_IN,סԺ���_IN,���_IN,ҽ������_IN,ҽ������_IN,ƴ������_IN,�۸�_IN,����_IN,��Ч��־_IN
                str��ˮ�� = GetMax��ˮ��(Nvl(!ҽ����))
                gstrSQL = "ZL_ZY_FYMXB_INSERT("
                gstrSQL = gstrSQL & "'" & Nvl(!ҽ����) & "',"       'ҽ������_IN
                gstrSQL = gstrSQL & str��ˮ�� & ","   '��ˮ��_IN
                gstrSQL = gstrSQL & "'" & lng����ID & "_" & lng��ҳID & "',"     'סԺ���_IN
                gstrSQL = gstrSQL & "'" & str��� & "',"    '���_IN
                gstrSQL = gstrSQL & "'" & strҽ������ & "',"   'ҽ������_IN
                gstrSQL = gstrSQL & "'" & strҽ������ & "',"   'ҽ������_IN
                If strƴ���� = "" Then
                    gstrSQL = gstrSQL & "'" & zlCommFun.zlGetSymbol(strҽ������, 0) & "',"     'ƴ������_IN
                Else
                    gstrSQL = gstrSQL & "'" & strƴ���� & "',"     'ƴ������_IN
                End If
                gstrSQL = gstrSQL & Format(Nvl(!ʵ�ս��, 0) / (Nvl(!����, 1) * Nvl(!����, 1)), "####0.0000;-####0.0000;0;0") & "," '�۸�IN
                gstrSQL = gstrSQL & Format((Nvl(!����, 1) * Nvl(!����, 1)), "####0.00;-####0.00;0;0") & ","   '����_IN
                gstrSQL = gstrSQL & "'F')"                                                              '��Ч��־_IN :ҽԺMIS��ֵΪF,ֻ����ҽ���޸����־����ҽ������������ɺ�����ΪT,�쳣����ʱ������ֵΪX�Ա�ҽԺMIS��ѯ?ҽԺMIS��Ȩ�޸����־���������Ը�
                DebugTool gstrSQL
                ExecuteProcedure_��ľ���� "סԺ��ϸ��Ϣд��"
                gstrSQL = "ZL_���˷��ü�¼_����ҽ��(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str��ˮ�� & "')"
                zlDatabase.ExecuteProcedure gstrSQL, "������ϸ����"
            End If
            .MoveNext
        Loop
    End With
    ����סԺ��ϸ��¼_��ľ���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function סԺ�������_��ľ����(rsExse As Recordset, ByVal lng����ID As Long, Optional bln���ʴ� As Boolean = True) As String
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

    g�������_��ľ����.����ID = 0
    If rsExse.RecordCount = 0 Then
        MsgBox "�ò���û���з������ã��޷����н��������", vbInformation, gstrSysName
        Exit Function
    End If
    'intMouse = Screen.MousePointer
    gstrSQL = "Select a.*,b.* From �����ʻ� a,������Ϣ b where a.����id=b.����id and a.����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)
    If rsTemp.EOF Then
        ShowMsgbox "����ص�ҽ��֤��!"
        Exit Function
    End If
    g�������_��ľ����.IC���� = Nvl(rsTemp!����)
    g�������_��ľ����.���� = Nvl(rsTemp!����)
    g�������_��ľ����.�Ա� = Nvl(rsTemp!�Ա�)
    

    gstrSQL = "select MAX(��ҳID) AS ��ҳID from ������ҳ where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", lng����ID)

    If IsNull(rsTemp("��ҳID")) = True Then
        MsgBox "ֻ��סԺ���˲ſ���ʹ��ҽ�����㡣", vbInformation, gstrSysName
        Exit Function
    End If
    lng��ҳID = rsTemp("��ҳID")


'    If bln���ʴ� Then
'        Screen.MousePointer = 1
'        If ��ݱ�ʶ_��ľ����(4, lng����id1) = "" Then
'            Screen.MousePointer = intMouse
'            סԺ�������_��ľ���� = ""
'            Exit Function
'        End If
'        Screen.MousePointer = intMouse
'        If lng����ID <> lng����id1 Then
'            ShowMsgbox "���ǵ�ǰҪ����Ĳ���!"
'            Exit Function
'        End If
'    End If

    
    'Screen.MousePointer = vbHourglass

    
    g�������_��ľ����.����ID = 0
    g�������_��ľ����.����ID = lng����ID
    
    
    '��һ��:���ܷ���
    DebugTool "סԺ�������,��һ��:���ܷ���"
    g�������_��ľ����.�����ܶ� = 0
    With rsExse
        If rsExse.RecordCount = 0 Then ShowMsgbox "δ������صķ��ü�¼!": Exit Function
        Do While Not .EOF
            g�������_��ľ����.�����ܶ� = g�������_��ľ����.�����ܶ� + Nvl(!���, 0)
            .MoveNext
        Loop
    End With
    
    '�ڶ���:������ϸ����
    DebugTool "סԺ�������,�ڶ���:������ϸ����"
    If ����סԺ��ϸ��¼_��ľ����(lng����ID, lng��ҳID) = False Then Exit Function
    
    '������:�����ܷ���
    
    '������:�ȴ�������Ϣ
     DebugTool "סԺ�������,������:�ȴ�������Ϣ"
    
    If frm����ȴ�_��ľ����.ShowWait(1, g�������_��ľ����.IC����) = False Then
        Exit Function
    End If

    
    
    '���Ĳ�:�ֽ���ؽ��
     DebugTool "סԺ�������,���Ĳ�:�ֽ���ؽ��"

     gstrSQL = "" & _
        "   Select  ybkh ҽ������, zybh סԺ���, rysj ��Ժʱ��, cysj ����ʱ��, jsbz ҽ�������־, tpbz ҽ����Ʊ��־, " & _
        "           yybz ҽԺ�����־, fyhj �����ܷ���, kszf ����֧��, tczf ͳ��֧��, gwycb ����Ա����," & _
        "           yj Ѻ���ܶ�, ybje Ӧ���ֽ��, gfcwf ���Ѵ�λ��, zfcwf �ԷѴ�λ��, gftwf ���ѵ��·�, zftwf �Էѵ��·� " & _
        "   from zy_jslsb   " & _
        "   where ybkh='" & g�������_��ľ����.IC���� & "'"
     Call OpenRecordset_��ľ����(rsTemp, "��ȡסԺ������Ϣ", gstrSQL)
    str���㷽ʽ = ""
    With g��������
        .�����ʻ�֧�� = Format(Nvl(rsTemp!����֧��, 0), "####0.00;-####0.00;0;0")
        .ͳ��֧�� = Format(Nvl(rsTemp!ͳ��֧��, 0), "####0.00;-####0.00;0;0")
        .Ӧ���ֽ�� = Format(Nvl(rsTemp!Ӧ���ֽ��, 0), "####0.00;-####0.00;0;0")
        .���Ѵ�λ�� = Format(Nvl(rsTemp!���Ѵ�λ��, 0), "####0.00;-####0.00;0;0")
        .���ѵ��·� = Format(Nvl(rsTemp!���ѵ��·�, 0), "####0.00;-####0.00;0;0")
        .����Ա���� = Format(Nvl(rsTemp!����Ա����, 0), "####0.00;-####0.00;0;0")
        .Ѻ���ܶ� = Format(Nvl(rsTemp!Ѻ���ܶ�, 0), "####0.00;-####0.00;0;0")
        .�ԷѴ�λ�� = Format(Nvl(rsTemp!�ԷѴ�λ��, 0), "####0.00;-####0.00;0;0")
        .�Էѵ��·� = Format(Nvl(rsTemp!�Էѵ��·�, 0), "####0.00;-####0.00;0;0")
        .�����ܶ� = Format(Nvl(rsTemp!�����ܷ���, 0), "####0.00;-####0.00;0;0")
        
        
        str���㷽ʽ = "�����ʻ�;" & .�����ʻ�֧�� & ";0"
        str���㷽ʽ = str���㷽ʽ & "|" & "ͳ��֧��;" & .ͳ��֧�� & ";0"
        str���㷽ʽ = str���㷽ʽ & "|" & "����Ա����;" & .����Ա���� & ";0"
    End With
    DebugTool "סԺ�������ɹ�,���㷽ʽ��" & str���㷽ʽ


    סԺ�������_��ľ���� = str���㷽ʽ
    g�������_��ľ����.����ID = lng����ID   '��ʾ�ò����Ѿ��������������
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function




Public Function סԺ����_��ľ����(lng����ID As Long, ByVal lng����ID As Long) As Boolean
    '���ܣ�����Ҫ���ν��ʵķ�����ϸ�ͽ������ݷ���ҽ��ǰ�÷�����ȷ�ϣ�
    '����: lng����ID -���˽��ʼ�¼ID, ��Ԥ����¼�п��Լ���ҽ���ź�����
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫ���ýӿڵķ�����ϸ���佻�׺͸������㽻�ף�
    '      2)�����ϣ���������ͨ��ģ�������ȡ�˻��������֤��ҽ��������������ȷ�ԣ���˽��ױ�Ȼ�ɹ������Ӱ�ȫ�Ƕȿ��ǣ����������㽻��ʧ��ʱ����Ҫʹ�÷���ɾ�����״�������������㽻�׳ɹ��������÷ָ��������Ǵ�������һ�£���Ҫִ�лָ����㽻�׺ͷ���ɾ�����ס��������ܱ�֤���ݵ���ȫͳһ��
    '      3)���ڽ���֮�󣬿���ʹ�ý������Ͻ��ף���ʱ��Ҫ����ʱִ�н��㽻�׵Ľ��׺ţ����������Ҫͬʱ���ʽ��׺š�(���������շ�����ʱ���Ѿ����ٺ�ҽ���й�ϵ�����Բ���Ҫ������ʵĽ��׺�)

    Dim rsTemp As New ADODB.Recordset, StrInput As String, strOutput As String

    Dim lng��ҳID As Long
    Dim dbl�����ܶ� As Double
    Dim strArr As Variant, strTmpArr As Variant

    Dim str���㷽ʽ  As String, strסԺ�� As String
    Dim obj���� As ��������
    Dim dbl�����ʻ� As Double

    סԺ����_��ľ���� = False


    Err = 0: On Error GoTo errHand:
    Call DebugTool("����סԺ����")

    
    '��һ��:������ݵ�����
    DebugTool "��һ��:������ݵ�����"
    If g�������_��ľ����.����ID <> lng����ID Then
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

    gstrSQL = " " & _
          " Select sum(round(nvl(���ʽ��,0),2)) as ʵ�ս�� " & _
          " From סԺ���ü�¼ " & _
          " Where ��¼״̬<>0 and ����ID=" & lng����ID & " and  Nvl(���ӱ�־,0)<>9"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ�ܷ���"

    dbl�����ܶ� = Round(Val(Nvl(rsTemp!ʵ�ս��, 0)), 2)
    If dbl�����ܶ� <> Round(g��������.�����ܶ�, 2) Then
        If dbl�����ܶ� - Round(g��������.�����ܶ�, 2) <= 0.1 Then
            If MsgBox("����������ݵķ����ܶ�(" & Format(g��������.�����ܶ�, "####0.00;-###0.00;0;0") & ")" & vbCrLf & "�뱾�ν���ķ����ܶ�(" & Format(dbl�����ܶ�, "####0.00;-###0.00;0;0") & ")���ȣ��Ƿ��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            Err.Raise 9000, gstrSysName, "����������ݵķ����ܶ�(" & Format(g��������.�����ܶ�, "####0.00;-###0.00;0;0") & ")" & vbCrLf & "�뱾�ν���ķ����ܶ�(" & Format(dbl�����ܶ�, "####0.00;-###0.00;0;0") & ")���ȣ����鴦���Ƿ���ȷ!"
            Exit Function
        End If
    End If
    g�������_��ľ����.����ID = lng����ID
    
    
    '�ڶ���:���������Ϣ
    DebugTool "�ڶ���:����ʼ���汣�ս����¼"

   '���뱣�ս����¼
    '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
    "   �ʻ��ۼ�����_IN(),�ʻ��ۼ�֧��_IN(),�ۼƽ���ͳ��_IN(),�ۼ�ͳ�ﱨ��_IN(),סԺ����_IN(סԺ:��ҳid),����(Ѻ���ܶ�),�ⶥ��_IN(���Ѵ�λ��),ʵ������_IN(�ԷѴ�λ��),
    '   �������ý��_IN(�����ܽ��),ȫ�Ը����_IN(���ѵ��·�),�����Ը����_IN(Ӧ���ֽ��),
    '   ����ͳ����_IN(ͳ��֧��),ͳ�ﱨ�����_IN(ͳ��֧��),���Ը����_IN(����Ա����),�����Ը����_IN(�Էѵ��·�),�����ʻ�֧��_IN(�����ʻ�֧�����),"
    '   ֧��˳���_IN(����:������),��ҳID_IN(��ҳid),��;����_IN,��ע_IN()
    
    gstrSQL = "zl_���ս����¼_insert(2," & lng����ID & "," & TYPE_�������� & "," & lng����ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "NULL,NULL,NULL,NULL," & lng��ҳID & "," & g��������.Ѻ���ܶ� & "," & g��������.���Ѵ�λ�� & "," & g��������.�ԷѴ�λ�� & "," & _
            g��������.�����ܶ� & "," & g��������.���ѵ��·� & "," & g��������.Ӧ���ֽ�� & "," & _
            g��������.ͳ��֧�� & " ," & g��������.ͳ��֧�� & "," & g��������.����Ա���� & "," & g��������.�Էѵ��·� & "," & g��������.�����ʻ�֧�� & ",'" & _
             lng����ID & "_" & lng��ҳID & "',NULL,NULL,NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������¼")
  
    '������:����м�������Ϣ
    DebugTool "������:����м�������Ϣ"
    gstrSQL = "ZL_סԺ_Clear('" & g�������_��ľ����.IC���� & "')"
    ExecuteProcedure_��ľ���� "�������"
    סԺ����_��ľ���� = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Public Function סԺ�������_��ľ����(lng����ID As Long) As Boolean
     '----------------------------------------------------------------
    '���ܣ���ָ�������漰�Ľ��ʽ��׺ͷ�����ϸ��ҽ��������ɾ����
    '������lng����ID-��Ҫ���ϵĽ��ʵ�ID�ţ�
    '���أ����׳ɹ�����true�����򣬷���false
    'ע�⣺1)��Ҫʹ�ý��ʻָ����׺ͷ���ɾ�����ף�
    '      2)�й�ԭ���㽻�׺ţ��ڲ��˽��ʼ�¼�и��ݽ��ʵ�ID���ң�ԭ������ϸ���佻�׵Ľ��׺ţ��ڲ��˷��ü�¼�и��ݽ���ID���ң�
    '      3)���ϵĽ��ʼ�¼(��¼����=2)�佻�׺ţ���д���ν��ʻָ����׵Ľ��׺ţ���������϶������ķ��ü�¼�Ľ��׺źţ���дΪ���η���ɾ�����׵Ľ��׺š�
    '----------------------------------------------------------------
    Err.Raise 9000, gstrSysName, "��ҽ����֧��סԺ�������,��������ѯ�ӿ���!"
    סԺ�������_��ľ���� = False
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function ҽ������_��ľ����(ByVal lng���� As Long, ByVal lngҽ������ As Integer) As Boolean
    '���ܣ� �÷������ڹ����Ӧ�ò���������������ҽ�����ݷ����������Ӵ�
    '���أ��ӿ����óɹ�������true�����򣬷���false
    
    Dim strConn As String
    Dim blnReturn As Boolean
    
    If frmSet��ľ����.�������� = False Then
        Exit Function
    End If
  
    If gcnOracle_��ľ���� Is Nothing Then
                blnReturn = True
    Else
        If Open�м��_��ľ����() Then
                blnReturn = True
        End If
    End If
    ҽ������_��ľ���� = blnReturn
End Function
Public Sub ExecuteProcedure_��ľ����(ByVal strCaption As String)
    '���ܣ�ִ��SQL���
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_��ľ����.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub
Public Function Open�м��_��ľ����() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    Dim strConn As String
    
    Open�м��_��ľ���� = False
    Err = 0: On Error Resume Next
        
    Err = 0: On Error GoTo errHand:
    
    '���½�����ҽ���������Ĺ�������
    '�м������
    gstrSQL = "select ������,����ֵ from ���ղ��� where  ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ľ���˺˹�ҵҽ��", TYPE_��������)
    
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
    
    Set gcnOracle_��ľ���� = New ADODB.Connection
    If OraDataOpen(gcnOracle_��ľ����, strServer, strUser, strPass, False) = False Then
        MsgBox "�޷����ӵ�ҽ���м�⣬���鱣�ղ����Ƿ�������ȷ��", vbInformation, gstrSysName
        Exit Function
    End If
    Open�м��_��ľ���� = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub OpenRecordset_��ľ����(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    Call SQLTest(App.ProductName, strCaption, IIf(strSQL = "", gstrSQL, strSQL))
    rsTemp.Open IIf(strSQL = "", gstrSQL, strSQL), gcnOracle_��ľ����, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

