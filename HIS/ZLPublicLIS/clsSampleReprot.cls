VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSampleReprot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InitSampleReprot(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal lngModule As Long, strErr As String) As Boolean
    '���ܣ��ӿڲ�����ʼ��
    '����
    'cnMain          ������
    'lngSys          ϵͳ��
    'lngModule       ϵͳģ��
    'strErr          ���ش�����־
    '����ֵ           true �ɹ���false ʧ��
    On Error GoTo errH
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = GetUnitName
    glngSys = lngSys
    glngModule = lngModule
    If GetUserInfo = False Then
        strErr = "δ�ҵ���Ӧ���û�"
        Exit Function
    End If
    InitSampleReprot = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function binInit() As Boolean
    If gcnOracle = "" Then
        binInit = False
        MsgBox "���ʼ��LIS�����ӿ�!", vbInformation, "LIS�����ӿڳ�ʼ��"
        Exit Function
    Else
        binInit = True
    End If
End Function


Public Function GetLaboratoryReportList(ByVal lngPatientID As Long, ByVal lngMainID As Long, Optional lngApplyID As Long) As String
    '���ܣ���ȡ���˵ļ��鱨���б�
    '����
    'lngPatientID           ����id
    'lngMainID              ��ҳid
    'lngApplyID             ����id
    '����                   XML��ʽ���ִ�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim objXML As New zl9ComLib.clsXML
    Dim strXMLOldLIS As String
    Dim strXMLNewLIS As String
    Dim i As Long
    On Error GoTo errH
    If binInit = False Then Exit Function
     InitObjLis
     Call gobjLISInsideComm.GetLabNewReportList(lngPatientID, lngMainID, strXMLNewLIS, lngApplyID)
     If strXMLNewLIS = "" Then
        strSQL = "Select id ���鱨��id,ҽ��id ����id,���� ������־,������Ŀ,�걾���,΢����걾 �Ƿ�΢����,������ �������,������,�����,���ʱ��,����ʱ�� " & vbNewLine & _
                 "  from ����걾��¼ where ����id = [1] and ��ҳid =[2] and ҽ��id is not null"
                 
        If lngApplyID > 0 Then
            strSQL = strSQL & " and ҽ��id =[3]"
        End If
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "���˼��鱨���б�", lngPatientID, lngMainID, lngApplyID)
        If rsTmp.RecordCount > 0 Then
            With objXML
                .ClearXmlText
                .AppendNode "���鱨���б�" ', True '���ڵ�[���鱨���б�]
                For i = 1 To rsTmp.RecordCount
                    .AppendData "���鱨��id", rsTmp!���鱨��id '<���鱨��id>���ͣ�
                    .AppendData "����id", rsTmp!����id '<����id>���ͣ�
                    .AppendData "������־", rsTmp!������־ & "" '<������־>���ͣ�
                    .AppendData "������Ŀ", rsTmp!������Ŀ & ""  '<������Ŀ>���ͣ�
                    .AppendData "�걾���", rsTmp!�걾��� '<�걾���>���ͣ�
                    .AppendData "�Ƿ�΢����", rsTmp!�Ƿ�΢���� & ""  '<΢����걾>���ͣ�
                    .AppendData "�������", Val(rsTmp!������� & "") '< ������ > ����:
                    .AppendData "������", rsTmp!������ & "" '<������>���ͣ�
                    .AppendData "�����", rsTmp!����� & ""  '<�����>���ͣ�
                    .AppendData "���ʱ��", rsTmp!���ʱ�� & "" '<���ʱ��>���ͣ�
                    .AppendData "����ʱ��", rsTmp!����ʱ�� & ""  '<����ʱ��>���ͣ�
                    rsTmp.MoveNext
                Next
                .AppendNode "���鱨���б�", True
                If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
            End With
        End If
        GetLaboratoryReportList = strXMLOldLIS
    Else
        GetLaboratoryReportList = strXMLNewLIS
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Function GetLaboratoryReportResultList(ByVal lngRepottID As Long) As String
    '���ܣ���ȡ���˵ļ��鱨����
    '����
    'lngRepottID            ����id
    '����                   XML��ʽ���ִ�
    
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsNewTmp As ADODB.Recordset
    Dim objXML As New zl9ComLib.clsXML
    Dim strXMLOldLIS As String
    Dim strXMLNewLIS As String
    Dim strBH As String
    Dim i As Long
    On Error GoTo errH
    If binInit = False Then Exit Function
    InitObjLis
    Call gobjLISInsideComm.GetLabNewReportResultList(lngRepottID, strXMLNewLIS)
    If strXMLNewLIS = "" Then
        '�°�û�����ݣ��鿴�ϰ��Ƿ��м�¼
        strSQL = "select  id,΢����걾 from ����걾��¼ where id = [1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "���˼������б�", lngRepottID)
        If rsTmp.RecordCount > 0 Then
            If Val(rsTmp("΢����걾") & "") = 1 Then
                strSQL = "Select distinct a.ϸ��id, b.������ ϸ����, a.�������� ����, a.��ҩ����, e.������ ������, c.��� �����ؽ��, c.������� ��ҩ��, c.ҩ������, e.�÷�����1, e.�÷�����2, e.ѪҩŨ��1," & vbNewLine & _
                         "          e.ѪҩŨ��2 , e.��ҩŨ��1, e.��ҩŨ��2" & vbNewLine & _
                         "   From ������ͨ��� A, ����ϸ�� B, ����ҩ����� C, �����ÿ����� E" & vbNewLine & _
                         "   Where a.ϸ��id = b.Id And c.ϸ�����id = a.Id And c.������id = e.Id and a.����걾id =[1] order by b.������"
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "���˼������б�", lngRepottID)
                If rsTmp.RecordCount > 0 Then
                    With objXML
                        .ClearXmlText
                        .AppendNode "΢������Ŀ" ', True '���ڵ�[��ͨ��Ŀ]
                        For i = 1 To rsTmp.RecordCount
                            If strBH <> rsTmp!ϸ���� & "" Then
                                If strBH <> "" Then
                                    .AppendNode "�����ؽ���б�", True
                                End If
                                strBH = rsTmp!ϸ���� & ""
                                .AppendData "ϸ��id", rsTmp!ϸ��id & "" '<ϸ��id>���ͣ�
                                .AppendData "ϸ����", rsTmp!ϸ���� & "" '<ϸ����>���ͣ�
                                .AppendData "����", rsTmp!���� & "" '<����>���ͣ�
                                .AppendData "��ҩ����", rsTmp!��ҩ���� & ""  '<��ҩ����>���ͣ�
                                .AppendNode "�����ؽ���б�" ', True '���ڵ�[ָ������]
                            End If
                        
                            .AppendData "������", rsTmp!������ & ""  '<������>���ͣ�
                            .AppendData "�����ؽ��", rsTmp!�����ؽ�� & ""  '<�����ؽ��>���ͣ�
                            .AppendData "��ҩ��", rsTmp!��ҩ�� & "" '<��ҩ��>���ͣ�
                            .AppendData "ҩ������", rsTmp!ҩ������ & ""  '<ҩ������>���ͣ�
                            .AppendData "�÷�����1", rsTmp!�÷�����1 & "" '<�÷�����1>���ͣ�
                            .AppendData "�÷�����2", rsTmp!�÷�����2 & ""  '<�÷�����2>���ͣ�
                            .AppendData "ѪҩŨ��1", rsTmp!ѪҩŨ��1 & "" '< ѪҩŨ��1 > ����:
                            .AppendData "ѪҩŨ��2", rsTmp!ѪҩŨ��2 & "" '<ѪҩŨ��2>���ͣ�
                            .AppendData "��ҩŨ��1", rsTmp!��ҩŨ��1 & ""  '<��ҩŨ��1>���ͣ�
                            .AppendData "��ҩŨ��2", rsTmp!��ҩŨ��2 & ""  '<��ҩŨ��2>���ͣ�
                            rsTmp.MoveNext
                        Next
                        .AppendNode "�����ؽ���б�", True
                        .AppendNode "΢������Ŀ", True
                        If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
                    End With
                End If
            Else
                strSQL = "Select a.������Ŀid ָ��id, b.Ӣ���� ָ�����, b.Ӣ���� ָ��Ӣ����, b.������ ָ��������, a.������, a.�����־, a.����ο�, a.�������, c.��˽��Ŀ" & vbNewLine & _
                         "   From ������ͨ��� A, ����������Ŀ B, ������Ŀ C" & vbNewLine & _
                         "   Where a.������Ŀid = b.Id And b.Id = c.������Ŀid and  a.����걾id =[1] "
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "���˼������б�", lngRepottID)
                If rsTmp.RecordCount > 0 Then
                    With objXML
                        .ClearXmlText
                        .AppendNode "��ͨ��Ŀ" ', True '���ڵ�[��ͨ��Ŀ]
                        .AppendNode "ָ������" ', True '���ڵ�[ָ������]
                        For i = 1 To rsTmp.RecordCount
                            .AppendData "ָ��id", rsTmp!ָ��id & "" '<ָ��id>���ͣ�
                            .AppendData "ָ�����", rsTmp!ָ����� & "" '<ָ�����>���ͣ�
                            .AppendData "ָ��Ӣ����", rsTmp!ָ��Ӣ���� & "" '<ָ��Ӣ����>���ͣ�
                            .AppendData "ָ��������", rsTmp!ָ�������� & ""  '<ָ��������>���ͣ�
                            .AppendData "������", rsTmp!������ & "" '<������>���ͣ�
                            .AppendData "�����־", rsTmp!�����־ & ""  '<�����־>���ͣ�
                            .AppendData "����ο�", rsTmp!����ο� & "" '< ����ο� > ����:
                            .AppendData "�������", rsTmp!������� & "" '<�������>���ͣ�
                            .AppendData "��˽��Ŀ", rsTmp!��˽��Ŀ & ""  '<��˽��Ŀ>���ͣ�
                            rsTmp.MoveNext
                        Next
                        .AppendNode "ָ������", True
                        .AppendNode "��ͨ��Ŀ", True
                        If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
                    End With
                End If
            End If
        End If
        If strXMLOldLIS <> "" Then GetLaboratoryReportResultList = strXMLOldLIS
    Else
        GetLaboratoryReportResultList = strXMLNewLIS
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Function GetBloodBankItem(ByVal lngApplyID As Long) As String
    '���ܣ���ȡ���˵ļ��鱨����
    '����
    'lngApplyID            ���id ,ҽ��id
    '����                   XML��ʽ���ִ�

    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsNewTmp As ADODB.Recordset
    Dim objXML As New zl9ComLib.clsXML
    Dim strXMLOldLIS As String
    Dim strXMLNewLIS As String
    Dim strBH As String
    Dim i As Long
    On Error GoTo errH
    If binInit = False Then Exit Function
    InitObjLis
    Call gobjLISInsideComm.GetNewBloodBankItem(lngApplyID, strXMLNewLIS)
    If strXMLNewLIS = "" Then
        '�鿴�ϰ��Ƿ��м�¼
        strSQL = "Select a.������Ŀid ָ��id, b.Ӣ���� ָ�����, b.Ӣ���� ָ��Ӣ����, b.������ ָ��������, a.������, a.�����־, a.����ο�" & vbNewLine & _
                 "   From ������ͨ��� A, ����������Ŀ B,����걾��¼ c" & vbNewLine & _
                 "   Where a.������Ŀid = b.Id  and a.����걾id =  c.id and   c.ҽ��id =[1] "
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "���˼������б�", lngApplyID)
        If rsTmp.RecordCount > 0 Then
            With objXML
                .ClearXmlText
                .AppendNode "��ͨ��Ŀ" ', True '���ڵ�[��ͨ��Ŀ]
                .AppendNode "ָ������" ', True '���ڵ�[ָ������]
                For i = 1 To rsTmp.RecordCount
                    .AppendData "ָ��id", rsTmp!ָ��id & "" '<ָ��id>���ͣ�
                    .AppendData "ָ�����", rsTmp!ָ����� & "" '<ָ�����>���ͣ�
                    .AppendData "ָ��Ӣ����", rsTmp!ָ��Ӣ���� & "" '<ָ��Ӣ����>���ͣ�
                    .AppendData "ָ��������", rsTmp!ָ�������� & ""  '<ָ��������>���ͣ�
                    .AppendData "������", rsTmp!������ & "" '<������>���ͣ�
                    .AppendData "�����־", rsTmp!�����־ & ""  '<�����־>���ͣ�
                    .AppendData "����ο�", rsTmp!����ο� & "" '<����ο�> ����:
                    rsTmp.MoveNext
                Next
                .AppendNode "ָ������", True
                .AppendNode "��ͨ��Ŀ", True
                If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
            End With
        End If
        If strXMLOldLIS <> "" Then GetBloodBankItem = strXMLOldLIS
    Else
        GetBloodBankItem = strXMLNewLIS
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Function GetTransfusionApplyFor(strItemCodeing As String, lngPatientID As Long, intPatientType As Integer, lngHomePageID As Long, Optional strRegistrationBill As String, _
                                        Optional intBaby As Integer, Optional intType As Integer, Optional ByVal intDay As Integer) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����                  ���ݴ���ҽ��ID���ؽ��
    '����
    '                       strItemCodeing ������Ŀ���루�ɴ�������ʹ�ö��ŷָ���
    '                       lngPatientID ����ID
    '                       intPatientType ������Դ 1-���2-סԺ
    '                       lngHomePageID ��ҳID ��������Դ=2ʱ��ѯ)
    '                       lngRegistrationBill �Һŵ�NO��������Դ<>2ʱ��ѯ���ξ��
    '                       intBaby           �Ƿ�Ӥ��
    '                       intType           ���ַ�ʽ��1=�ٴ˲�7���ڵġ�0 = ����ѯ 2=ָ����ѯ������intDay����������= �ݶ�
    '                       intDay            ��intType=2ʱ���˲�������Ч��ָ��Ҫ��ѯ�����������
    '
    '�걾��ɸ�ʽ
    '
    '                   ָ��1<split1>���Ʊ���1<split1>��λ1<split1>��˽��Ŀ1<split1>ָ�����1<split1>������1<split1>Ӣ����1<split1>ȡֵ����1<split1>
    '                       ������1<split2>�����־1<split2>����ο�1<split2>�������1<split2>�걾����1<split3>
    '                   ָ��2<split1>���Ʊ���2<split1>��˽��Ŀ2<split1>ָ�����2<split1>������2<split1>Ӣ����2<split1>ȡֵ����2<split1>
    '                       ������2<split2>�����־2<split2>����ο�2<split2>�������2<split2>�걾����2<split3>
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim rsTmpRuest As New ADODB.Recordset
    Dim lngID As Long
    Dim lngSampleId As Long
    Dim lngSampleGroup As Long
    Dim lngMicroID As Long
    Dim strSampleOne As String
    Dim strSampleTwo As String
    Dim strErr As String
    Dim varItemCodeing As Variant
    Dim strItemcodeOne As String
    Dim i As Integer
    Dim strStartTime As String
    Dim strEndTime As String
    Dim strTmp As String
    
    strEndTime = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    If intType = 1 Then
        strStartTime = Format(zlDatabase.Currentdate - 7, "yyyy-mm-dd 00:00:00")
    ElseIf intType = 2 Then
        strStartTime = Format(zlDatabase.Currentdate - intDay, "yyyy-mm-dd 00:00:00")
    End If
    
    '�ָ��ĳ���
    Const conSplit1 As String = "<split1>"                        '���ڷָ���Ŀ��Ϣ,ʹ�á�<split1>���ָ�
    Const conSplit2 As String = "<split2>"                        '���ڷָ������Ŀ,ʹ�á�<split2>���ָ�
    Const conSplit3 As String = "<split3>"                        '���ڷָ���Ŀ��Ϣ,ʹ�á�<split1>���ָ�
    Const conSplit4 As String = "<split4>"                        '���ڷָ������Ŀ,ʹ�á�<split2>���ָ�
        
    'ֻ����ͨ��Ŀ
    On Error GoTo errH
    If binInit = False Then Exit Function
    strTmp = ""
    InitObjLis
    If Not gobjLISInsideComm Is Nothing Then
        GetTransfusionApplyFor = gobjLISInsideComm.GetNewTransFusionApplyFor(strItemCodeing, lngPatientID, intPatientType, lngHomePageID, strRegistrationBill, intBaby, intType, intDay)
        strTmp = Replace(GetTransfusionApplyFor, "<split1>", "")
        strTmp = Replace(strTmp, "<split2>", "")
        strTmp = Replace(strTmp, "<split3>", "")
        strTmp = Replace(strTmp, "<split4>", "")
        strTmp = Trim(strTmp)
    End If
    If strTmp = "" Then
        varItemCodeing = Split(strItemCodeing, ",")
        For i = LBound(varItemCodeing) To UBound(varItemCodeing)
            strItemcodeOne = varItemCodeing(i)
            strSQL = "  Select d.id ָ��id, d.������ || '(' || d.Ӣ���� || ')' ָ��, b.��λ, b.��˽��Ŀ, d.���� ָ�����, d.������, d.Ӣ����, b.ȡֵ����" & vbNewLine & _
                     "   From ���鱨����Ŀ A, ������Ŀ B, ������ĿĿ¼ C, ����������Ŀ D" & vbNewLine & _
                     "   Where b.������Ŀid = d.Id And a.������Ŀid = c.Id And  a.������Ŀid = d.Id And c.���� = [1]"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ���", strItemcodeOne)
            Do Until rsTmp.EOF
                strSampleOne = strSampleOne & conSplit3 & rsTmp("ָ��") & conSplit1 & strItemcodeOne & conSplit1 & rsTmp("��λ") & _
                                conSplit1 & rsTmp("��˽��Ŀ") & conSplit1 & rsTmp("ָ�����") & _
                                conSplit1 & rsTmp("������") & conSplit1 & rsTmp("Ӣ����") & conSplit1 & rsTmp("ȡֵ����") & conSplit1
                strSQL = " Select *" & vbNewLine & _
                         "   From (Select  ���ʱ��, c.������, Decode(c.�����־, 1, '', 2, '��', 3, '��', 4, '�쳣', 5, '����', 6, '����', '') �����־, c.����ο�," & vbNewLine & _
                         "                 c.������� , a.�걾����" & vbNewLine & _
                         "          From ����걾��¼ A, ������Ŀ B, ������ͨ��� C, ����������Ŀ D" & vbNewLine & _
                         "          Where a.Id = c.����걾id And b.������Ŀid = d.Id And c.������Ŀid = d.Id " & vbNewLine & _
                         "          [����]  and a.���ʱ�� is not null and a.������Դ=[4] and d.id= [5]" & vbNewLine & _
                         "          Order By ���ʱ�� desc ) E" & vbNewLine & _
                         "   Where Rownum = 1"
                If intPatientType = 2 Then
                    If intBaby <> 0 Then
                        strSQL = Replace(strSQL, "[����]", " and A.����ID = [1] and a.��ҳid=[2]  and a.Ӥ��=[6] ")
                    Else
                        strSQL = Replace(strSQL, "[����]", " and A.����ID = [1] and a.��ҳid=[2]  and nvl(a.Ӥ��,0)= 0 ")
                    End If
                    Set rsTmpRuest = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ���", lngPatientID, lngHomePageID, strRegistrationBill, intPatientType, Val(rsTmp("ָ��id")), intBaby)
                    If rsTmpRuest.RecordCount > 0 Then
                        strSampleOne = strSampleOne & rsTmpRuest("������") & conSplit2 & rsTmpRuest("�����־") & conSplit2 & rsTmpRuest("����ο�") & conSplit2 & rsTmpRuest("�������") & conSplit2 & rsTmpRuest("�걾����") & conSplit2 & rsTmpRuest("���ʱ��")
                    Else
                        If intType = 1 Or intType = 2 Then
                            strSQL = " Select *" & vbNewLine & _
                                    "   From (Select  ���ʱ��, c.������, Decode(c.�����־, 1, '', 2, '��', 3, '��', 4, '�쳣', 5, '����', 6, '����', '') �����־, c.����ο�," & vbNewLine & _
                                    "                 c.������� , a.�걾����" & vbNewLine & _
                                    "          From ����걾��¼ A, ������Ŀ B, ������ͨ��� C, ����������Ŀ D" & vbNewLine & _
                                    "          Where a.Id = c.����걾id And b.������Ŀid = d.Id And c.������Ŀid = d.Id " & vbNewLine & _
                                    "          [����]  and a.���ʱ�� is not null  and d.id= [4]" & vbNewLine & _
                                    "          Order By ���ʱ�� desc ) E" & vbNewLine & _
                                    "   Where Rownum = 1"
                            If intBaby <> 0 Then
                                strSQL = Replace(strSQL, "[����]", " and A.����ID = [1] and  a.���ʱ�� between [2] and [3]  and a.Ӥ��=[5]  ")
                            Else
                                strSQL = Replace(strSQL, "[����]", " and A.����ID = [1] and  a.���ʱ�� between [2] and [3]  and nvl(a.Ӥ��,0)= 0  ")
                            End If
                            Set rsTmpRuest = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ���", lngPatientID, CDate(strStartTime), CDate(strEndTime), Val(rsTmp("ָ��id")), intBaby)
                            If rsTmpRuest.RecordCount > 0 Then
                                strSampleOne = strSampleOne & rsTmpRuest("������") & conSplit2 & rsTmpRuest("�����־") & conSplit2 & rsTmpRuest("����ο�") & conSplit2 & rsTmpRuest("�������") & conSplit2 & rsTmpRuest("�걾����") & conSplit2 & rsTmpRuest("���ʱ��")
                            End If
                        End If
                    End If
                Else
                    strSQL = Replace(strSQL, "[����]", " and A.����ID = [1] and  a.�Һŵ�=[3] ")
                    Set rsTmpRuest = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ���", lngPatientID, lngHomePageID, strRegistrationBill, intPatientType, Val(rsTmp("ָ��id")))
                    If rsTmpRuest.RecordCount > 0 Then
                        strSampleOne = strSampleOne & rsTmpRuest("������") & conSplit2 & rsTmpRuest("�����־") & conSplit2 & rsTmpRuest("����ο�") & conSplit2 & rsTmpRuest("�������") & conSplit2 & rsTmpRuest("�걾����") & conSplit2 & rsTmpRuest("���ʱ��")
                    End If
                End If

                rsTmp.MoveNext
            Loop
        Next
        If strSampleOne <> "" Then
            strSampleOne = Mid(strSampleOne, Len(conSplit3) + 1)
        End If
        GetTransfusionApplyFor = strSampleOne
    End If
    Exit Function
errH:
    strErr = "��ȡ���������" & Err.Number & " " & Err.Description
End Function

Public Function ReviseDate(ByVal strDate As String) As String
'���ܣ���ʱ��ת��Ϊͳһ��24Сʱ��ʱ��
    ReviseDate = Format(strDate, "yyyy-mm-dd hh:mm:ss")
End Function

Private Sub Class_Initialize()
    If gobjComLib Is Nothing Then Set gobjComLib = CreateObject("zl9Comlib.clsComlib")
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/5/10
'��    ��:��ʼ��ͨѶ����
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Function Init() As Boolean
    Init = funInit
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/4/24
'��    ��:ʹ��winsock����������������Ϣ
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Sub SendMessage(ByVal strSend As String)
    If gintStart = 0 Then Exit Sub  '���δ����ͨѶ����ֱ���˳�
    Call frmComm.funSendMessage(strSend)
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/4/24
'��    ��:���ⲿ�ṩ��ȡ�����Ľӿ�
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Function GetIniValue(ByVal strKey As String, Optional ByVal strDefault As String) As String
    GetIniValue = funGetIniValue(strKey, strDefault)
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/4/24
'��    ��:�ⲿ�򿪲������ý���
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Sub ShowParaSet(objFrm As Object)
    frmParaSet.Show 1, objFrm
End Sub
