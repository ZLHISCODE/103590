Attribute VB_Name = "mdlCISOutPath"
Option Explicit

Public Function ExportOutPathToXML(ByVal lng·��ID As Long, ByVal int�汾�� As Integer, ByVal strFile As String) As Boolean
'���ܣ����������ٴ�·����XML�ļ�
'������strFile=����·�����ļ���
'˵������������·����Ϣ��ָ���汾����Ϣ
    Dim xPath As DOMDocument
    Dim xRoot As IXMLDOMElement
    Dim xNode As IXMLDOMNode
    Dim xSubNode1 As IXMLDOMNode
    Dim xSubNode2 As IXMLDOMNode
    Dim xSubNode3 As IXMLDOMNode
    Dim xSubNode4 As IXMLDOMNode
    Dim xSubNode5 As IXMLDOMNode
    Dim xPI As IXMLDOMProcessingInstruction
    
    Dim rsTmp As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    Dim rsItemAdvice As ADODB.Recordset
    Dim rsItemEPR As ADODB.Recordset
    Dim rsEvalMark As ADODB.Recordset
    Dim rsEvalCond As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    Set xPath = New DOMDocument
    
    'ע��
    xPath.appendChild xPath.createComment(gstrSysName & "  ����Ա:" & UserInfo.���� & ",����:" & UserInfo.������ & ",ʱ��:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
    
    '�����
    Set xRoot = xPath.createElement("ClinicalPathways")
    Set xPath.documentElement = xRoot
    Call xRoot.setAttribute("ID", lng·��ID)
    Call xRoot.setAttribute("Version", int�汾��)

    '�����ٴ�·����Ϣ
    strSql = "Select A.����,A.����,A.����,A.ͨ��,A.���°汾," & _
        " A.�����Ա�,A.��������,A.˵��,B.��׼����ʱ��,B.��׼����," & _
        " B.�汾˵��,B.������,B.����ʱ��,B.�����,B.���ʱ��,B.ͣ����,B.ͣ��ʱ��,A.�����ʱ�� " & _
        " From ����·��Ŀ¼ A,����·���汾 B Where A.ID=B.·��ID And A.ID=[1] And B.�汾��=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID, int�汾��)
    
    Set xNode = CreateNode(1, xRoot, "PathInfo", NODE_ELEMENT, "")
        CreateNode 2, xNode, "����", , rsTmp!����
        CreateNode 2, xNode, "����", , rsTmp!����
        CreateNode 2, xNode, "����", , rsTmp!����
        CreateNode 2, xNode, "ͨ��", , NVL(rsTmp!ͨ��)
        CreateNode 2, xNode, "���°汾", , NVL(rsTmp!���°汾)
        CreateNode 2, xNode, "�����Ա�", , NVL(rsTmp!�����Ա�)
        CreateNode 2, xNode, "��������", , NVL(rsTmp!��������)
        CreateNode 2, xNode, "˵��", , NVL(rsTmp!˵��)
        CreateNode 2, xNode, "��׼����ʱ��", , NVL(rsTmp!��׼����ʱ��)
        CreateNode 2, xNode, "��׼����", , NVL(rsTmp!��׼����)
        CreateNode 2, xNode, "�汾˵��", , NVL(rsTmp!�汾˵��)
        CreateNode 2, xNode, "������", , NVL(rsTmp!������)
        CreateNode 2, xNode, "����ʱ��", , Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "�����", , NVL(rsTmp!�����)
        CreateNode 2, xNode, "���ʱ��", , Format(NVL(rsTmp!���ʱ��), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "ͣ����", , NVL(rsTmp!ͣ����)
        CreateNode 2, xNode, "ͣ��ʱ��", , Format(NVL(rsTmp!ͣ��ʱ��), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "�����ʱ��", , NVL(rsTmp!�����ʱ��)
    '����·������
    strSql = "Select B.ID,B.����,B.���� From ����·������ A,���ű� B Where A.·��ID=[1] And A.����ID=B.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathDepts", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathDept", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "����ID", , rsTmp!ID
                CreateNode 3, xSubNode1, "����", , rsTmp!����
                CreateNode 3, xSubNode1, "����", , rsTmp!����
            rsTmp.MoveNext
        Loop
    End If
    
    '����·������
    strSql = "Select A.����ID,B.���� as ������,B.���� as ������," & _
        " A.���ID,C.���� as �����,C.���� as ����� " & _
        " From ����·������ A,��������Ŀ¼ B,�������Ŀ¼ C" & _
        " Where Nvl(A.����ID,0)=B.ID(+) And Nvl(A.���ID,0)=C.ID(+) And A.·��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathDiseases", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathDisease", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "����ID", , NVL(rsTmp!����id)
                CreateNode 3, xSubNode1, "������", , NVL(rsTmp!������)
                CreateNode 3, xSubNode1, "������", , NVL(rsTmp!������)
                CreateNode 3, xSubNode1, "���ID", , NVL(rsTmp!���id)
                CreateNode 3, xSubNode1, "�����", , NVL(rsTmp!�����)
                CreateNode 3, xSubNode1, "�����", , NVL(rsTmp!�����)
            rsTmp.MoveNext
        Loop
    End If
    
    '��������
    strSql = "Select B.��������,B.�׶�ID,A.ID,A.����ָ��,A.ָ������,A.ָ����" & _
        " From ����·������ָ�� A,����·������ B" & _
        " Where A.����ID=B.ID And B.·��ID=[1] And �汾��=[2]" & _
        " Order by B.��������,B.�׶�ID,A.���"
    Set rsEvalMark = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID, int�汾��)
    
    strSql = "Select B.��������,B.�׶�ID,A.ָ��ID,A.��ĿID,A.��ϵʽ,A.����ֵ,A.�������" & _
        " From ����·���������� A,����·������ B" & _
        " Where A.����ID=B.ID And B.·��ID=[1] And �汾��=[2]" & _
        " Order by B.��������,B.�׶�ID"
    Set rsEvalCond = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID, int�汾��)
    
    rsEvalMark.Filter = "��������=1"
    rsEvalCond.Filter = "��������=1"
    If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
        Set xNode = CreateNode(1, xRoot, "ImportEval", NODE_ELEMENT, "")
            If Not rsEvalMark.EOF Then
                Set xSubNode1 = CreateNode(2, xNode, "Marks", NODE_ELEMENT, "")
                Do While Not rsEvalMark.EOF
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Mark", NODE_ELEMENT, "")
                        CreateNode 4, xSubNode2, "ID", , rsEvalMark!ID
                        CreateNode 4, xSubNode2, "����ָ��", , rsEvalMark!����ָ��
                        CreateNode 4, xSubNode2, "ָ������", , rsEvalMark!ָ������
                        CreateNode 4, xSubNode2, "ָ����", , rsEvalMark!ָ����
                    rsEvalMark.MoveNext
                Loop
            End If
            If Not rsEvalCond.EOF Then
                Set xSubNode1 = CreateNode(2, xNode, "Conditions", NODE_ELEMENT, "")
                Do While Not rsEvalCond.EOF
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Condition", NODE_ELEMENT, "")
                        CreateNode 4, xSubNode2, "ָ��ID", , rsEvalCond!ָ��ID
                        CreateNode 4, xSubNode2, "��ϵʽ", , rsEvalCond!��ϵʽ
                        CreateNode 4, xSubNode2, "����ֵ", , rsEvalCond!����ֵ
                        CreateNode 4, xSubNode2, "�������", , rsEvalCond!�������
                    rsEvalCond.MoveNext
                Loop
            End If
    End If
    
    '����·��ҽ������
    strSql = "Select Distinct A.ID,A.���ID,A.���,A.��Ч,A.������ĿID,D.���� as ���Ʊ���,D.���� as ��������," & _
        " A.�շ�ϸĿID,E.���� as �շѱ���,E.���� as �շ�����,A.ҽ������,A.��������,A.�ܸ�����," & _
        " A.�걾��λ,A.��鷽��,A.ҽ������,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ," & _
        " A.ִ������,A.ִ�п���ID,F.���� as ִ�п�����,F.���� as ִ�п�����,A.ʱ�䷽��,A.�Ƿ�ȱʡ,A.�Ƿ�ѡ,A.�䷽ID,A.�����ĿID" & _
        " From ����·��ҽ������ A,����·��ҽ�� B,����·����Ŀ C,������ĿĿ¼ D,�շ���ĿĿ¼ E,���ű� F" & _
        " Where A.ID=B.ҽ������ID And B.·����ĿID=C.ID And C.·��ID=[1] And C.�汾��=[2]" & _
        " And Nvl(A.������ĿID,0)=D.ID(+) And Nvl(A.�շ�ϸĿID,0)=E.ID(+) And Nvl(A.ִ�п���ID,0)=F.ID(+)" & _
        " Order by A.���,A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID, int�汾��)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathAdvices", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathAdvice", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "ID", , rsTmp!ID
                CreateNode 3, xSubNode1, "���ID", , NVL(rsTmp!���id)
                CreateNode 3, xSubNode1, "���", , rsTmp!���
                CreateNode 3, xSubNode1, "��Ч", , rsTmp!��Ч
                CreateNode 3, xSubNode1, "������ĿID", , NVL(rsTmp!������ĿID)
                CreateNode 3, xSubNode1, "���Ʊ���", , NVL(rsTmp!���Ʊ���)
                CreateNode 3, xSubNode1, "��������", , NVL(rsTmp!��������)
                CreateNode 3, xSubNode1, "�շ�ϸĿID", , NVL(rsTmp!�շ�ϸĿID)
                CreateNode 3, xSubNode1, "�շѱ���", , NVL(rsTmp!�շѱ���)
                CreateNode 3, xSubNode1, "�շ�����", , NVL(rsTmp!�շ�����)
                CreateNode 3, xSubNode1, "ҽ������", , NVL(rsTmp!ҽ������)
                CreateNode 3, xSubNode1, "��������", , NVL(rsTmp!��������)
                CreateNode 3, xSubNode1, "�ܸ�����", , NVL(rsTmp!�ܸ�����)
                CreateNode 3, xSubNode1, "�걾��λ", , NVL(rsTmp!�걾��λ)
                CreateNode 3, xSubNode1, "��鷽��", , NVL(rsTmp!��鷽��)
                CreateNode 3, xSubNode1, "ҽ������", , NVL(rsTmp!ҽ������)
                CreateNode 3, xSubNode1, "ִ��Ƶ��", , NVL(rsTmp!ִ��Ƶ��)
                CreateNode 3, xSubNode1, "Ƶ�ʴ���", , NVL(rsTmp!Ƶ�ʴ���)
                CreateNode 3, xSubNode1, "Ƶ�ʼ��", , NVL(rsTmp!Ƶ�ʼ��)
                CreateNode 3, xSubNode1, "�����λ", , NVL(rsTmp!�����λ)
                CreateNode 3, xSubNode1, "ִ������", , NVL(rsTmp!ִ������)
                CreateNode 3, xSubNode1, "ִ�п���ID", , NVL(rsTmp!ִ�п���ID)
                CreateNode 3, xSubNode1, "ִ�п�����", , NVL(rsTmp!ִ�п�����)
                CreateNode 3, xSubNode1, "ִ�п�����", , NVL(rsTmp!ִ�п�����)
                CreateNode 3, xSubNode1, "ʱ�䷽��", , NVL(rsTmp!ʱ�䷽��)
                CreateNode 3, xSubNode1, "�Ƿ�ȱʡ", , NVL(rsTmp!�Ƿ�ȱʡ, 0)
                CreateNode 3, xSubNode1, "�Ƿ�ѡ", , NVL(rsTmp!�Ƿ�ѡ, 0)
                CreateNode 3, xSubNode1, "�䷽ID", , NVL(rsTmp!�䷽ID)
                CreateNode 3, xSubNode1, "�����ĿID", , NVL(rsTmp!�����ĿID)
            rsTmp.MoveNext
        Loop
    End If
    
    '����·������
    strSql = "Select ���� From ����·������ Where ·��ID=[1] And �汾��=[2] Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID, int�汾��)
    
    Set xNode = CreateNode(1, xRoot, "PathCategorys", NODE_ELEMENT, "")
    Do While Not rsTmp.EOF
        Set xSubNode1 = CreateNode(2, xNode, "PathCategory", NODE_ELEMENT, NVL(rsTmp!����))
        CreateNode 2, xSubNode1, "����", NODE_ELEMENT, NVL(rsTmp!����)
        rsTmp.MoveNext
    Loop
    
    '����·���׶�/��Ŀ
    strSql = "Select ID,Nvl(��ID,0) as ��ID,���,����,��ʼ����,��������,����,˵��" & _
        " From ����·���׶� Where ·��ID=[1] And �汾��=[2] Order by Nvl(��ID,0) Desc,���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID, int�汾��)
    
    strSql = "Select ID,�׶�ID,����,��Ŀ���,��Ŀ����,ִ�з�ʽ,��Ŀ���,ͼ��ID,����Ҫ��" & _
        " From ����·����Ŀ Where ·��ID=[1] And �汾��=[2] Order by �׶�ID,����,��Ŀ���"
    Set rsItem = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID, int�汾��)
    
    strSql = "Select A.·����ĿID,A.ҽ������ID From ����·��ҽ�� A,����·����Ŀ B" & _
        " Where A.·����ĿID=B.ID And B.·��ID=[1] And �汾��=[2]"
    Set rsItemAdvice = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID, int�汾��)
    
    strSql = "Select A.��ĿID,A.�ļ�ID,C.���,C.���� From ����·������ A,����·����Ŀ B,�����ļ��б� C" & _
        " Where A.��ĿID=B.ID And A.�ļ�ID=C.ID And B.·��ID=[1] And �汾��=[2]"
    Set rsItemEPR = zlDatabase.OpenSQLRecord(strSql, "ExportOutPathToXML", lng·��ID, int�汾��)
    
    Set rsClone = rsTmp.Clone: rsTmp.Filter = "��ID=0"
    
    Set xNode = CreateNode(1, xRoot, "PathTimeSteps", NODE_ELEMENT, "")
    Do While Not rsTmp.EOF
        'ȱʡ��֧
        Set xSubNode1 = CreateNode(2, xNode, "PathTimeStep", NODE_ELEMENT, "")
            CreateNode 3, xSubNode1, "ID", , rsTmp!ID
            CreateNode 3, xSubNode1, "��ID", , ""
            CreateNode 3, xSubNode1, "���", , rsTmp!���
            CreateNode 3, xSubNode1, "����", , rsTmp!����
            CreateNode 3, xSubNode1, "��ʼ����", , NVL(rsTmp!��ʼ����)
            CreateNode 3, xSubNode1, "��������", , NVL(rsTmp!��������)
            CreateNode 3, xSubNode1, "˵��", , NVL(rsTmp!˵��)
            CreateNode 3, xSubNode1, "����", , NVL(rsTmp!����)
            
            '�׶ε���Ŀ
            rsItem.Filter = "�׶�ID=" & rsTmp!ID
            Set xSubNode2 = CreateNode(3, xSubNode1, "Items", NODE_ELEMENT, "")
            Do While Not rsItem.EOF
                Set xSubNode3 = CreateNode(4, xSubNode2, "Item", NODE_ELEMENT, "")
                    CreateNode 5, xSubNode3, "ID", , rsItem!ID
                    CreateNode 5, xSubNode3, "����", , rsItem!����
                    CreateNode 5, xSubNode3, "��Ŀ���", , rsItem!��Ŀ���
                    CreateNode 5, xSubNode3, "��Ŀ����", , rsItem!��Ŀ����
                    CreateNode 5, xSubNode3, "ִ�з�ʽ", , NVL(rsItem!ִ�з�ʽ)
                    CreateNode 5, xSubNode3, "��Ŀ���", , NVL(rsItem!��Ŀ���)
                    CreateNode 5, xSubNode3, "ͼ��ID", , NVL(rsItem!ͼ��ID)
                    CreateNode 5, xSubNode3, "����Ҫ��", , NVL(rsItem!����Ҫ��, 0)

                    '��Ŀ��Ӧ��ҽ��
                    rsItemAdvice.Filter = "·����ĿID=" & rsItem!ID
                    If Not rsItemAdvice.EOF Then
                        Set xSubNode4 = CreateNode(5, xSubNode3, "Advices", NODE_ELEMENT, "")
                        Do While Not rsItemAdvice.EOF
                            CreateNode 6, xSubNode4, "Advice", , rsItemAdvice!ҽ������ID
                            rsItemAdvice.MoveNext
                        Loop
                    End If
                    '��Ŀ��Ӧ�Ĳ���
                    rsItemEPR.Filter = "��ĿID=" & rsItem!ID
                    If Not rsItemEPR.EOF Then
                        Set xSubNode4 = CreateNode(5, xSubNode3, "EPRFiles", NODE_ELEMENT, "")
                        Do While Not rsItemEPR.EOF
                            Set xSubNode5 = CreateNode(6, xSubNode4, "EPRFile", NODE_ELEMENT, "")
                                CreateNode 7, xSubNode5, "�ļ�ID", , rsItemEPR!�ļ�ID
                                CreateNode 7, xSubNode5, "�ļ����", , rsItemEPR!���
                                CreateNode 7, xSubNode5, "�ļ�����", , rsItemEPR!����
                            rsItemEPR.MoveNext
                        Loop
                    End If
                    
                rsItem.MoveNext
            Loop
        
            '�׶ε�����
            rsEvalMark.Filter = "��������=2 And �׶�ID=" & rsTmp!ID
            rsEvalCond.Filter = "��������=2 And �׶�ID=" & rsTmp!ID
            If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
                Set xSubNode2 = CreateNode(3, xSubNode1, "StepEval", NODE_ELEMENT, "")
                    If Not rsEvalMark.EOF Then
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Marks", NODE_ELEMENT, "")
                        Do While Not rsEvalMark.EOF
                            Set xSubNode4 = CreateNode(5, xSubNode3, "Mark", NODE_ELEMENT, "")
                                CreateNode 6, xSubNode4, "ID", , rsEvalMark!ID
                                CreateNode 6, xSubNode4, "����ָ��", , rsEvalMark!����ָ��
                                CreateNode 6, xSubNode4, "ָ������", , rsEvalMark!ָ������
                                CreateNode 6, xSubNode4, "ָ����", , rsEvalMark!ָ����
                            rsEvalMark.MoveNext
                        Loop
                    End If
                    If Not rsEvalCond.EOF Then
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Conditions", NODE_ELEMENT, "")
                        Do While Not rsEvalCond.EOF
                            Set xSubNode4 = CreateNode(5, xSubNode3, "Condition", NODE_ELEMENT, "")
                                CreateNode 6, xSubNode4, "ָ��ID", , NVL(rsEvalCond!ָ��ID)
                                CreateNode 6, xSubNode4, "��ĿID", , NVL(rsEvalCond!��ĿID)
                                CreateNode 6, xSubNode4, "��ϵʽ", , rsEvalCond!��ϵʽ
                                CreateNode 6, xSubNode4, "����ֵ", , rsEvalCond!����ֵ
                                CreateNode 6, xSubNode4, "�������", , rsEvalCond!�������
                            rsEvalCond.MoveNext
                        Loop
                    End If
            End If
        
        '��ѡ��֧
        rsClone.Filter = "��ID=" & rsTmp!ID
        If Not rsClone.EOF Then
            Do While Not rsClone.EOF
                Set xSubNode1 = CreateNode(2, xNode, "PathTimeStep", NODE_ELEMENT, "")
                    CreateNode 3, xSubNode1, "ID", , rsClone!ID
                    CreateNode 3, xSubNode1, "��ID", , rsClone!��ID
                    CreateNode 3, xSubNode1, "���", , rsClone!���
                    CreateNode 3, xSubNode1, "����", , rsClone!����
                    CreateNode 3, xSubNode1, "��ʼ����", , NVL(rsClone!��ʼ����)
                    CreateNode 3, xSubNode1, "��������", , NVL(rsClone!��������)
                    CreateNode 3, xSubNode1, "˵��", , NVL(rsClone!˵��)
                
                    '�׶ε���Ŀ
                    rsItem.Filter = "�׶�ID=" & rsClone!ID
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Items", NODE_ELEMENT, "")
                    Do While Not rsItem.EOF
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Item", NODE_ELEMENT, "")
                            CreateNode 5, xSubNode3, "ID", , rsItem!ID
                            CreateNode 5, xSubNode3, "����", , rsItem!����
                            CreateNode 5, xSubNode3, "��Ŀ���", , rsItem!��Ŀ���
                            CreateNode 5, xSubNode3, "��Ŀ����", , rsItem!��Ŀ����
                            CreateNode 5, xSubNode3, "ִ�з�ʽ", , NVL(rsItem!ִ�з�ʽ)
                            CreateNode 5, xSubNode3, "��Ŀ���", , NVL(rsItem!��Ŀ���)
                            CreateNode 5, xSubNode3, "ͼ��ID", , NVL(rsItem!ͼ��ID)
                            
                            '��Ŀ��Ӧ��ҽ��
                            rsItemAdvice.Filter = "·����ĿID=" & rsItem!ID
                            If Not rsItemAdvice.EOF Then
                                Set xSubNode4 = CreateNode(5, xSubNode3, "Advices", NODE_ELEMENT, "")
                                Do While Not rsItemAdvice.EOF
                                    CreateNode 6, xSubNode4, "Advice", , rsItemAdvice!ҽ������ID
                                    rsItemAdvice.MoveNext
                                Loop
                            End If
                            '��Ŀ��Ӧ�Ĳ���
                            rsItemEPR.Filter = "��ĿID=" & rsItem!ID
                            If Not rsItemEPR.EOF Then
                                Set xSubNode4 = CreateNode(5, xSubNode3, "EPRFiles", NODE_ELEMENT, "")
                                Do While Not rsItemEPR.EOF
                                    Set xSubNode5 = CreateNode(6, xSubNode4, "EPRFile", NODE_ELEMENT, "")
                                        CreateNode 7, xSubNode5, "�ļ�ID", , rsItemEPR!�ļ�ID
                                        CreateNode 7, xSubNode5, "�ļ����", , rsItemEPR!���
                                        CreateNode 7, xSubNode5, "�ļ�����", , rsItemEPR!����
                                    rsItemEPR.MoveNext
                                Loop
                            End If
                            
                        rsItem.MoveNext
                    Loop
                    
                    '�׶ε�����
                    rsEvalMark.Filter = "��������=2 And �׶�ID=" & rsClone!ID
                    rsEvalCond.Filter = "��������=2 And �׶�ID=" & rsClone!ID
                    If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
                        Set xSubNode2 = CreateNode(3, xSubNode1, "StepEval", NODE_ELEMENT, "")
                            If Not rsEvalMark.EOF Then
                                Set xSubNode3 = CreateNode(4, xSubNode2, "Marks", NODE_ELEMENT, "")
                                Do While Not rsEvalMark.EOF
                                    Set xSubNode4 = CreateNode(5, xSubNode3, "Mark", NODE_ELEMENT, "")
                                        CreateNode 6, xSubNode4, "ID", , rsEvalMark!ID
                                        CreateNode 6, xSubNode4, "����ָ��", , rsEvalMark!����ָ��
                                        CreateNode 6, xSubNode4, "ָ������", , rsEvalMark!ָ������
                                        CreateNode 6, xSubNode4, "ָ����", , rsEvalMark!ָ����
                                    rsEvalMark.MoveNext
                                Loop
                            End If
                            If Not rsEvalCond.EOF Then
                                Set xSubNode3 = CreateNode(4, xSubNode2, "Conditions", NODE_ELEMENT, "")
                                Do While Not rsEvalCond.EOF
                                    Set xSubNode4 = CreateNode(5, xSubNode3, "Condition", NODE_ELEMENT, "")
                                        CreateNode 6, xSubNode4, "ָ��ID", , NVL(rsEvalCond!ָ��ID)
                                        CreateNode 6, xSubNode4, "��ĿID", , NVL(rsEvalCond!��ĿID)
                                        CreateNode 6, xSubNode4, "��ϵʽ", , rsEvalCond!��ϵʽ
                                        CreateNode 6, xSubNode4, "����ֵ", , rsEvalCond!����ֵ
                                        CreateNode 6, xSubNode4, "�������", , rsEvalCond!�������
                                    rsEvalCond.MoveNext
                                Loop
                            End If
                    End If
                
                rsClone.MoveNext
            Loop
        End If
        
        rsTmp.MoveNext
    Loop
    
    'XML��Ϣ
    Set xPI = xPath.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call xPath.insertBefore(xPI, xPath.childNodes(0))
    
    '������ļ�
    xPath.Save strFile
    Set xPath = Nothing
    
    ExportOutPathToXML = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set xPath = Nothing
End Function

Public Function ImportOutPathFromXML(ByVal strFile As String, _
    Optional ByVal lng·��ID As Long, Optional ByVal int�汾�� As Integer, _
    Optional ByVal intLimit As Integer, Optional ByRef blnLimit As Boolean) As Boolean
'���ܣ�����ָ���������ٴ�·��XML�ļ�
'������lng·��ID,int�汾��=���ָ������ֻ����汾��ز�����Ϣ�����û��ָ��������ݸ���XML�е���Ϣ����·������������ȫ����
'      intLimit=�������Ƶ����·������,Ϊ0��ʾ������
'      blnLimit=�Ƿ���������·�����������Ƶ���ʧ��
    Dim rsTmp As ADODB.Recordset
    Dim rsIcon As ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    
    Dim arrSQL As Variant, strSql As String
    Dim colItemID As Collection
    Dim colStepID As Collection
    Dim colMarkID As Collection
    Dim colAdviceID As Collection
    Dim colAdviceOriginalID As Collection
    Dim colBranchID As Collection
    Dim colPreID As Collection
    
    Dim xPath As DOMDocument
    Dim xRoot As IXMLDOMElement
    Dim xNode As IXMLDOMNode
    Dim xSubNode1 As IXMLDOMNode
    Dim xSubNode2 As IXMLDOMNode
    Dim xSubNode3 As IXMLDOMNode
    Dim xSubNode4 As IXMLDOMNode
    Dim xSubNode5 As IXMLDOMNode
    
    Dim str���� As String, lng�׶�ID As Long
    Dim strValue As String, strTemp1 As String
    Dim strTemp2 As String, strTemp3 As String
    Dim blnDo As Boolean, blnTran As Boolean
    Dim i As Long, k As Long, n As Long, m As Long
    Dim strPreStep As String
    Dim strtemp4 As String
    Dim strImportRef As String
    Dim lng������ As Long '��¼ͬһ·����Ŀҽ���ĵ���״̬0��ȫ��δ���룬1��ȫ�����룬2�����ֵ���
    Dim lngCount As Long, str��IDs As String, arrID As Variant, lng��ID As Long, strFilter As String
    Dim lng��ĿID As Long
    
    On Error GoTo errH
    
    rsAdvice.Fields.Append "ID", adBigInt
    rsAdvice.Fields.Append "���ID", adBigInt, , adFldIsNullable
    rsAdvice.Fields.Append "����ο�", adVarChar, 200, adFldIsNullable
    rsAdvice.Fields.Append "��ĿID", adBigInt, , adFldIsNullable
    rsAdvice.Fields.Append "����״̬", adInteger
    
    rsAdvice.CursorLocation = adUseClient
    rsAdvice.LockType = adLockOptimistic
    rsAdvice.CursorType = adOpenStatic
    rsAdvice.Open
    
    blnLimit = False
    
    Set xPath = New DOMDocument
    xPath.Load strFile
    
    '����������κ�Ԫ�أ����˳�
    If xPath.documentElement Is Nothing Then
        Set xPath = Nothing
        Screen.MousePointer = 0
        Exit Function
    End If
    
    arrSQL = Array()
    
    '��ȡXML����
    Set xRoot = xPath.selectSingleNode("ClinicalPathways")
    Set xNode = xRoot.selectSingleNode("PathInfo")
    If lng·��ID = 0 Then
        '��ȡӦ�ÿ��ҵ����
        strTemp1 = ""
        If Val(GetNodeValue(xNode, "ͨ��")) = 2 Then
            Set xSubNode1 = xRoot.selectSingleNode("PathDepts")
            If Not xSubNode1 Is Nothing Then
                strSql = "Select A.ID,A.����,A.����" & _
                    " From ���ű� A,��������˵�� C" & _
                    " Where A.ID=C.����ID And C.������� IN(1,3) And C.��������='�ٴ�'" & _
                    " Order by A.����"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML")
                
                For Each xSubNode2 In xSubNode1.childNodes
                    rsTmp.Filter = "����='" & GetNodeValue(xSubNode2, "����") & "' And ����='" & GetNodeValue(xSubNode2, "����") & "'"
                    If Not rsTmp.EOF Then strTemp1 = strTemp1 & "," & rsTmp!ID
                Next
            
                strTemp1 = Mid(strTemp1, 2)
            End If
        End If
        
        '��ȡӦ�ü��������
        strValue = ""
        Set xSubNode1 = xRoot.selectSingleNode("PathDiseases")
        If Not xSubNode1 Is Nothing Then
            strTemp2 = "": strTemp3 = ""
            For Each xSubNode2 In xSubNode1.childNodes
                If Val(GetNodeValue(xSubNode2, "����ID")) <> 0 Then
                    strSql = "Select ID From ��������Ŀ¼ Where ����=[1] And ����=[2]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode2, "������"), GetNodeValue(xSubNode2, "������"))
                    If Not rsTmp.EOF Then strTemp2 = strTemp2 & "," & rsTmp!ID
                ElseIf Val(GetNodeValue(xSubNode2, "���ID")) <> 0 Then
                    strSql = "Select ID From �������Ŀ¼ Where ����=[1] And ����=[2]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode2, "�����"), GetNodeValue(xSubNode2, "�����"))
                    If Not rsTmp.EOF Then strTemp3 = strTemp3 & "," & rsTmp!ID
                End If
            Next
            If strTemp2 <> "" Or strTemp3 <> "" Then
                strValue = Mid(strTemp2, 2) & ";" & Mid(strTemp3, 2)
            End If
        End If
        
        '�����ٴ�·����Ϣ
        strSql = "Select ID,����,���°汾 From ����·��Ŀ¼ Where ����=[1] And ����=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xNode, "����"), GetNodeValue(xNode, "����"))
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        If Not rsTmp.EOF Then
            '�����汾���߸��ǰ汾
            lng·��ID = rsTmp!ID
            int�汾�� = NVL(rsTmp!���°汾, 0) + 1 '���ܸ���δ��˰汾
            str���� = rsTmp!����
            arrSQL(UBound(arrSQL)) = "zl_����·��Ŀ¼_Update(" & _
                lng·��ID & ",'" & GetNodeValue(xNode, "����") & "','" & str���� & "'," & _
                "'" & GetNodeValue(xNode, "����") & "','" & GetNodeValue(xNode, "˵��") & "'," & _
                Val(GetNodeValue(xNode, "�����Ա�")) & ",'" & GetNodeValue(xNode, "��������") & "'," & _
                Val(GetNodeValue(xNode, "ͨ��")) & "," & Val(GetNodeValue(xNode, "�����ʱ��")) & ",'" & strTemp1 & "','" & strValue & "')"
        
        Else
            '�����Ȩ����
            If intLimit > 0 Then
                strSql = "Select Nvl(Count(*),0) as ���� From ����·��Ŀ¼"
                Set rsTmp = New ADODB.Recordset
                Call zlDatabase.OpenRecordset(rsTmp, strSql, "ImportOutPathFromXML")
                If rsTmp!���� >= intLimit Then
                    blnLimit = True
                    Set xPath = Nothing
                    Screen.MousePointer = 0
                    Exit Function
                End If
            End If
            
            '����·��
            lng·��ID = zlDatabase.GetNextId("����·��Ŀ¼")
            int�汾�� = 1
            str���� = GetNextCode(GetNodeValue(xNode, "����"), 1)
            arrSQL(UBound(arrSQL)) = "zl_����·��Ŀ¼_Insert(" & _
                "'" & GetNodeValue(xNode, "����") & "','" & str���� & "'," & _
                "'" & GetNodeValue(xNode, "����") & "','" & GetNodeValue(xNode, "˵��") & "'," & _
                Val(GetNodeValue(xNode, "�����Ա�")) & ",'" & GetNodeValue(xNode, "��������") & "'," & _
                Val(GetNodeValue(xNode, "ͨ��")) & "," & Val(GetNodeValue(xNode, "�����ʱ��")) & ",'" & strTemp1 & "','" & strValue & "'," & lng·��ID & ")"
        End If
    End If
    
    'ɾ���汾��ص����ݣ����²���
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_����·���汾_Delete(" & lng·��ID & "," & int�汾�� & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_����·���汾_Update(" & lng·��ID & "," & int�汾�� & "," & _
        "'" & GetNodeValue(xNode, "��׼����ʱ��") & "','" & GetNodeValue(xNode, "��׼����") & "'," & _
        "'" & GetNodeValue(xNode, "�汾˵��") & "')"
    
    '��������
    Set xNode = xRoot.selectSingleNode("ImportEval")
    If Not xNode Is Nothing Then
        Set xSubNode1 = xNode.selectSingleNode("Marks")
        If Not xSubNode1 Is Nothing Then
            k = 1
            Set colItemID = New Collection
            For Each xSubNode2 In xSubNode1.childNodes
                strValue = zlDatabase.GetNextId("����·������ָ��")
                colItemID.Add strValue, "_" & GetNodeValue(xSubNode2, "ID")
                            
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����·������ָ��_Insert(" & lng·��ID & "," & int�汾�� & ",NULL,1," & _
                    strValue & "," & k & ",'" & GetNodeValue(xSubNode2, "����ָ��") & "'," & _
                    Val(GetNodeValue(xSubNode2, "ָ������")) & ",'" & GetNodeValue(xSubNode2, "ָ����") & "')"
                
                k = k + 1
            Next
        End If
        Set xSubNode1 = xNode.selectSingleNode("Conditions")
        If Not xSubNode1 Is Nothing Then
            For Each xSubNode2 In xSubNode1.childNodes
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����·����������_Insert(" & lng·��ID & "," & int�汾�� & ",NULL,1," & _
                    colItemID("_" & GetNodeValue(xSubNode2, "ָ��ID")) & ",NULL,'" & GetNodeValue(xSubNode2, "��ϵʽ") & "'," & _
                    "'" & GetNodeValue(xSubNode2, "����ֵ") & "','" & GetNodeValue(xSubNode2, "�������") & "')"
            Next
        End If
    End If
    
    '����·��ҽ������
    Set xNode = xRoot.selectSingleNode("PathAdvices")
    If Not xNode Is Nothing Then
        Set colAdviceID = New Collection
        Set colAdviceOriginalID = New Collection
        For Each xSubNode1 In xNode.childNodes
            strValue = zlDatabase.GetNextId("����·��ҽ������")
            strTemp1 = GetNodeValue(xSubNode1, "ID")
            colAdviceID.Add strValue, "_" & strTemp1
            colAdviceOriginalID.Add strTemp1, "_" & strValue
        Next
        k = 1
        For Each xSubNode1 In xNode.childNodes
            blnDo = True: strTemp1 = "": strTemp2 = "": strTemp3 = ""
                
            '��֤������ĿID
            If Val(GetNodeValue(xSubNode1, "������ĿID")) <> 0 Then
                strSql = "Select ����,ID From ������ĿĿ¼ Where ����=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode1, "��������"))
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "����='" & GetNodeValue(xSubNode1, "���Ʊ���") & "'"
                    If rsTmp.RecordCount > 0 Then
                        strTemp1 = rsTmp!ID
                    Else
                        rsTmp.Filter = ""
                        strTemp1 = rsTmp!ID
                    End If
                Else
                    blnDo = False
                End If
            End If
            '��֤�շ�ϸĿID
            If blnDo And Val(GetNodeValue(xSubNode1, "�շ�ϸĿID")) <> 0 Then
                strSql = "Select ����,ID From �շ���ĿĿ¼ Where ����=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode1, "�շ�����"))
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "����='" & GetNodeValue(xSubNode1, "�շѱ���") & "'"
                    If rsTmp.RecordCount > 0 Then
                        strTemp2 = rsTmp!ID
                    Else
                        rsTmp.Filter = ""
                        strTemp2 = rsTmp!ID
                    End If
                Else
                    blnDo = False
                End If
            End If
            '��ȡ����ο�
            strImportRef = IIf(Val(GetNodeValue(xSubNode1, "������ĿID")) <> 0, Trim(GetNodeValue(xSubNode1, "��������")) & _
                IIf(Val(GetNodeValue(xSubNode1, "�շ�ϸĿID")) <> 0, "(" & Trim(GetNodeValue(xSubNode1, "�շ�����")) & ")", ""), "" & _
                IIf(Val(GetNodeValue(xSubNode1, "�շ�ϸĿID")) <> 0, Trim(GetNodeValue(xSubNode1, "�շ�����")), ""))
            '����·��ҽ���ĵ���״��������ʱ��¼��
            rsAdvice.AddNew
            rsAdvice!ID = Val(GetNodeValue(xSubNode1, "ID"))
            rsAdvice!���id = Val(GetNodeValue(xSubNode1, "���ID"))
            rsAdvice!����ο� = strImportRef
            rsAdvice!����״̬ = IIf(blnDo, 1, 0)
            rsAdvice.Update
            
            If blnDo Then
                '��ִ֤�п���ID
                If Val(GetNodeValue(xSubNode1, "ִ�п���ID")) <> 0 Then
                    strSql = "Select ����,ID From ���ű� Where ����=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode1, "ִ�п�����"))
                    If Not rsTmp.EOF Then
                        rsTmp.Filter = "����='" & GetNodeValue(xSubNode1, "ִ�п�����") & "'"
                        If rsTmp.RecordCount > 0 Then
                            strTemp3 = rsTmp!ID
                        Else
                            rsTmp.Filter = ""
                            strTemp3 = rsTmp!ID
                        End If
                    End If
                End If
                
                strValue = GetNodeValue(xSubNode1, "���ID")
                If strValue <> "" Then strValue = colAdviceID("_" & strValue)
                                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����·��ҽ������_Insert(" & _
                    colAdviceID("_" & GetNodeValue(xSubNode1, "ID")) & "," & ZVal(strValue) & "," & _
                    k & "," & Val(GetNodeValue(xSubNode1, "��Ч")) & "," & ZVal(strTemp1) & "," & _
                    "'" & GetNodeValue(xSubNode1, "ҽ������") & "'," & ZVal(GetNodeValue(xSubNode1, "��������")) & "," & _
                    ZVal(GetNodeValue(xSubNode1, "�ܸ�����")) & "," & ZVal(strTemp2) & "," & _
                    "'" & GetNodeValue(xSubNode1, "�걾��λ") & "','" & GetNodeValue(xSubNode1, "��鷽��") & "'," & _
                    "'" & GetNodeValue(xSubNode1, "ִ��Ƶ��") & "'," & ZVal(GetNodeValue(xSubNode1, "Ƶ�ʴ���")) & "," & _
                    ZVal(GetNodeValue(xSubNode1, "Ƶ�ʼ��")) & ",'" & GetNodeValue(xSubNode1, "�����λ") & "'," & _
                    "'" & GetNodeValue(xSubNode1, "ҽ������") & "'," & Val(GetNodeValue(xSubNode1, "ִ������")) & "," & _
                    ZVal(strTemp3) & ",'" & GetNodeValue(xSubNode1, "ʱ�䷽��") & "',Null,Null," & GetNodeValue(xSubNode1, "�Ƿ�ȱʡ", 0) & "," & _
                    GetNodeValue(xSubNode1, "�Ƿ�ѡ", 0) & ",Null," & ZVal(GetNodeValue(xSubNode1, "�䷽ID", 0)) & "," & ZVal(GetNodeValue(xSubNode1, "�����ĿID", 0)) & ")"
                k = k + 1
            Else
                '��������IDΪ��ҽ���ģ�����Щҽ����Ӧ����
                strValue = GetNodeValue(xSubNode1, "ID")
                For n = 0 To UBound(arrSQL)
                    If arrSQL(n) <> "" Then
                        If Split(arrSQL(n), ",")(1) = colAdviceID("_" & strValue) Then
                            '������ҽ��������
                            strTemp1 = Split(Split(arrSQL(n), ",")(0), "(")(1)
                            colAdviceID.Remove "_" & colAdviceOriginalID("_" & strTemp1)
                            colAdviceID.Add "0", "_" & colAdviceOriginalID("_" & strTemp1)
                            arrSQL(n) = ""
                        End If
                    End If
                Next
                '������ҽ��������
                colAdviceID.Remove "_" & strValue
                colAdviceID.Add "0", "_" & strValue
            End If
        Next
    End If
    
    '����·������
    Set xNode = xRoot.selectSingleNode("PathCategorys")
    k = 1
    For Each xSubNode1 In xNode.childNodes
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_����·������_Insert(" & lng·��ID & "," & int�汾�� & "," & k & ",'" & IIf(GetNodeValue(xSubNode1, "����") = "", xSubNode1.Text, GetNodeValue(xSubNode1, "����")) & "')"
        k = k + 1
    Next
    
    '����·���׶�
    Set xNode = xRoot.selectSingleNode("PathTimeSteps")
    k = 1
    Set colStepID = New Collection
    For Each xSubNode1 In xNode.childNodes
        lng�׶�ID = zlDatabase.GetNextId("����·���׶�")
        colStepID.Add lng�׶�ID, "_" & GetNodeValue(xSubNode1, "ID")
        
        strTemp1 = GetNodeValue(xSubNode1, "��ID")
        If strTemp1 <> "" Then strTemp1 = colStepID("_" & strTemp1)
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        If strPreStep <> "" Then
            If InStr("," & strPreStep & ",", "," & GetNodeValue(xSubNode1, "ID") & ",") > 0 Then
                strtemp4 = colPreID("_" & GetNodeValue(xSubNode1, "ID"))
            End If
        End If
        arrSQL(UBound(arrSQL)) = "Zl_����·���׶�_Insert(" & _
            lng�׶�ID & "," & lng·��ID & "," & int�汾�� & "," & ZVal(strTemp1) & "," & _
            IIf(strTemp1 = "", k, GetNodeValue(xSubNode1, "���")) & ",'" & GetNodeValue(xSubNode1, "����") & "'," & _
            ZVal(GetNodeValue(xSubNode1, "��ʼ����")) & "," & ZVal(GetNodeValue(xSubNode1, "��������")) & "," & _
            "'" & GetNodeValue(xSubNode1, "˵��") & "'," & _
            "'" & GetNodeValue(xSubNode1, "����") & "')"
        If strTemp1 = "" Then k = k + 1
        strtemp4 = ""
        
        '�׶��е�����·����Ŀ
        Set xSubNode2 = xSubNode1.selectSingleNode("Items")
        If Not xSubNode2 Is Nothing Then
            Set colItemID = New Collection
            For Each xSubNode3 In xSubNode2.childNodes
                strTemp1 = "": strTemp2 = ""
                '��Ŀ����ҽ��
                lng��ĿID = Val(GetNodeValue(xSubNode3, "ID"))
                Set xSubNode4 = xSubNode3.selectSingleNode("Advices")
                If Not xSubNode4 Is Nothing Then
                    For Each xSubNode5 In xSubNode4.childNodes
                        '����ʱ�ṹ��¼��������ҽ������Ŀ�Ĺ���
                        rsAdvice.Filter = "ID=" & Val(xSubNode5.Text)
                        If rsAdvice.RecordCount <> 0 Then
                            Call rsAdvice.Update("��ĿID", lng��ĿID)
                        End If
                        rsAdvice.Filter = ""
                        
                        If Val(colAdviceID("_" & xSubNode5.Text)) <> 0 Then
                            strTemp1 = strTemp1 & "," & colAdviceID("_" & xSubNode5.Text)
                        End If
                    Next
                    strTemp1 = Mid(strTemp1, 2)
                End If
                
                '��Ŀ��������
                Set xSubNode4 = xSubNode3.selectSingleNode("EPRFiles")
                i = 1
                If Not xSubNode4 Is Nothing Then
                    For Each xSubNode5 In xSubNode4.childNodes
                        '��֤�����ļ�ID
                        strSql = "Select ID From �����ļ��б� Where ���=[1] And ����=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML", GetNodeValue(xSubNode5, "�ļ����"), GetNodeValue(xSubNode5, "�ļ�����"))
                        If Not rsTmp.EOF Then strTemp2 = strTemp2 & ";" & rsTmp!ID & ",," & GetNodeValue(xSubNode5, "�ļ�����") & "," & i + 1
                    Next
                    strTemp2 = Mid(strTemp2, 2)
                End If
                
                'ͼ�����֤��ֻ֧�ֹ���ͼ��
                strTemp3 = GetNodeValue(xSubNode3, "ͼ��ID")
                If strTemp3 <> "" Then
                    If rsIcon Is Nothing Then
                        strSql = "Select ID,Nvl(����,0) as ���� From �ٴ�·��ͼ��"
                        Set rsIcon = zlDatabase.OpenSQLRecord(strSql, "ImportOutPathFromXML")
                    End If
                    rsIcon.Filter = "ID=" & strTemp3 & " And ����=1"
                    If rsIcon.EOF Then strTemp3 = ""
                End If
                
                strValue = zlDatabase.GetNextId("����·����Ŀ")
                colItemID.Add strValue, "_" & GetNodeValue(xSubNode3, "ID")
                
                rsAdvice.Filter = "��ĿID=" & lng��ĿID
                
                lngCount = rsAdvice.RecordCount
                strImportRef = ""
                lng������ = 1
                str��IDs = ""
                
                rsAdvice.Filter = rsAdvice.Filter & " And ����״̬=0"
                '��ȡ����״̬
                If rsAdvice.RecordCount <> 0 Then
                    lng������ = IIf(rsAdvice.RecordCount = lngCount, 0, 2)
                    '��ȡδ����ɹ�ҽ������ID
                    For n = 1 To rsAdvice.RecordCount
                        lng��ID = rsAdvice!���id
                        If lng��ID = 0 Then lng��ID = rsAdvice!ID
                        If InStr(str��IDs & ",", "," & lng��ID & ",") = 0 Then
                            str��IDs = str��IDs & "," & lng��ID
                        End If
                        rsAdvice.MoveNext
                    Next
                End If
                If Len(str��IDs) > 0 Then str��IDs = Mid(str��IDs, 2)

                arrID = Split(str��IDs, ",")
                '��ȡ����ο�
                For m = LBound(arrID) To UBound(arrID)
                    '����δ�����ͬһ��ҽ��
                    strFilter = "(��ĿID = " & lng��ĿID & " AND ���ID = " & Val(arrID(m)) & ") OR (��ĿID = " & lng��ĿID & " AND ID=" & Val(arrID(m)) & ")"
                    rsAdvice.Filter = strFilter
                    rsAdvice.Sort = "���ID,ID"
                    If rsAdvice.RecordCount <> 0 Then
                        For n = 1 To rsAdvice.RecordCount
                            If n = 1 And strImportRef = "" Then
                                strImportRef = rsAdvice!����ο�
                            ElseIf n = 1 And strImportRef <> "" Then
                                strImportRef = strImportRef & Chr(10) & Chr(13) & rsAdvice!����ο� '�Ѿ���������ҽ���Ѿ�������strImportRef
                            Else
                                strImportRef = strImportRef & ";" & rsAdvice!����ο�
                            End If
                            rsAdvice.MoveNext
                        Next
                    End If
                Next
   
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����·����Ŀ_Insert(" & _
                    strValue & "," & lng·��ID & "," & int�汾�� & "," & lng�׶�ID & "," & _
                    "'" & GetNodeValue(xSubNode3, "����") & "'," & GetNodeValue(xSubNode3, "��Ŀ���") & "," & _
                    "'" & GetNodeValue(xSubNode3, "��Ŀ����") & "'," & Val(GetNodeValue(xSubNode3, "ִ�з�ʽ")) & _
                    ",'" & GetNodeValue(xSubNode3, "��Ŀ���") & "'," & _
                    ZVal(strTemp3) & ",'" & strTemp1 & "','" & strTemp2 & "'," & GetNodeValue(xSubNode3, "����Ҫ��", 0) & _
                    ",'" & Trim(strImportRef) & "'," & IIf(Trim(strImportRef) = "" And lng������ = 1, "Null", lng������) & ")"
            Next
        End If
        
        Set xSubNode2 = xSubNode1.selectSingleNode("StepEval")
        If Not xSubNode2 Is Nothing Then
            '����ָ��
            Set xSubNode3 = xSubNode2.selectSingleNode("Marks")
            If Not xSubNode3 Is Nothing Then
                i = 1
                Set colMarkID = New Collection
                For Each xSubNode4 In xSubNode3.childNodes
                    strValue = zlDatabase.GetNextId("����·������ָ��")
                    colMarkID.Add strValue, "_" & GetNodeValue(xSubNode4, "ID")
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·������ָ��_Insert(" & _
                        lng·��ID & "," & int�汾�� & "," & lng�׶�ID & ",2," & _
                        strValue & "," & i & ",'" & GetNodeValue(xSubNode4, "����ָ��") & "'," & _
                        Val(GetNodeValue(xSubNode4, "ָ������")) & ",'" & GetNodeValue(xSubNode4, "ָ����") & "')"
                    i = i + 1
                Next
            End If
            'ָ������
            Set xSubNode3 = xSubNode2.selectSingleNode("Conditions")
            If Not xSubNode3 Is Nothing Then
                For Each xSubNode4 In xSubNode3.childNodes
                    strTemp1 = GetNodeValue(xSubNode4, "ָ��ID")
                    If strTemp1 <> "" Then strTemp1 = colMarkID("_" & strTemp1)
                    strTemp2 = GetNodeValue(xSubNode4, "��ĿID")
                    If strTemp2 <> "" Then strTemp2 = colItemID("_" & strTemp2)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·����������_Insert(" & _
                        lng·��ID & "," & int�汾�� & "," & lng�׶�ID & ",2," & _
                        ZVal(strTemp1) & "," & ZVal(strTemp2) & ",'" & GetNodeValue(xSubNode4, "��ϵʽ") & "'," & _
                        "'" & GetNodeValue(xSubNode4, "����ֵ") & "'," & Val(GetNodeValue(xSubNode4, "�������")) & ")"
                Next
            End If
        End If
    Next
    
    'ִ���ύ����
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "ImportOutPathFromXML"
        End If
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    Set xPath = Nothing
    ImportOutPathFromXML = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set xPath = Nothing
End Function

Public Function CheckNotFinishPath(ByVal lng����ID As Long, ByVal lng�Һ�ID As Long, ByRef lngPathID As Long, ByRef strMsg As String) As Boolean
'����Ƿ��������ִ�е��ٴ�·��
    Dim strSql As String, rsPati As Recordset

    On Error GoTo errH

    strSql = " Select ID, �Һ�ID, ����ʱ��, ״̬" & vbNewLine & _
             " From (Select ID, �Һ�ID, ����ʱ��, ״̬ From ��������·�� Where ����id = [1] Order By ����ʱ�� Desc)" & vbNewLine & _
             " Where Rownum < 2"

    Set rsPati = zlDatabase.OpenSQLRecord(strSql, "CheckNotFinishpath", lng����ID)
    If rsPati.RecordCount > 0 Then
        If Val(NVL(rsPati!�Һ�ID)) = lng�Һ�ID Then
            CheckNotFinishPath = False                          '�ò����Ѿ��е������ٴ�·��
            strMsg = "�ò����Ѿ��е������ٴ�·��"
        ElseIf Val(NVL(rsPati!״̬)) = 1 Then
            lngPathID = Val(NVL(rsPati!ID))
            CheckNotFinishPath = True                           '�ò��˵�ǰ����Ϊ��ɵ��ٴ�·��
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckContinuePath(ByVal lngPathID As Long, ByVal lng����ID As Long, ByVal strDiagIDs As String, ByRef strMsg As String) As Boolean
'����Ƿ���Լ���·��
'1����εĵ�һ��ϣ���ҽ������ҽ�������·�����������ͬ��
'2�����ʱ������Ч�ڼ���
    Dim strSql As String, rsTmp As Recordset
    Dim lng�׶μ�� As Long
    
    On Error GoTo errH

    strSql = " Select Nvl(a.����id, a.���id) As ���id,A.����ID" & vbNewLine & _
             " From ��������·�� A, ����·���汾 B" & vbNewLine & _
             " Where ID = [1] And a.·��id = b.·��id And a.�汾�� = b.�汾��"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckContinuePath", lngPathID)
    If rsTmp.RecordCount > 0 Then
        lng�׶μ�� = GetIntervalTime(lngPathID)
        
        '�����ϲ�ͬ�����ܼ���ԭ�е��ٴ�·��
        If InStr("," & strDiagIDs & ",", "," & rsTmp!���id & ",") < 0 Then
            CheckContinuePath = False
            strMsg = "��ξ������Ҫ��Ϻ�·���ĵ�����ϲ�ͬ"
'        ElseIf lng�׶μ�� <> 0 And Val(NVL(rsTmp!���׶μ��)) > lng�׶μ�� Then
'            CheckContinuePath = False
'            strMsg = "���������׶μ��ʱ��"
'        ElseIf Val(NVL(rsTmp!���׶μ��)) <> lng����ID Then
'            CheckContinuePath = False
'            strMsg = "�ϴ��ٴ�·���Ŀ��Һ͵�ǰ�������ڵĿ��Ҳ�ͬ"
        Else
            CheckContinuePath = True
        End If
        
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetIntervalTime(ByVal lngPathID As Long) As Long
'�����ϴ�ִ��·��������Ϊֹ�ļ��ʱ��
    Dim strSql As String, rsTmp As Recordset
    Dim datLatTime As Date
    
    On Error GoTo errH
    strSql = "Select Max(ִ��ʱ��) as ִ��ʱ�� from ��������·��ִ�� where ·����¼ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetIntervalTime", lngPathID)
    
    If rsTmp.RecordCount > 0 And rsTmp!ִ��ʱ�� & "" <> "" Then
        datLatTime = CDate(NVL(rsTmp!ִ��ʱ��))
    End If
    
    If datLatTime <> CDate(0) Then
        GetIntervalTime = DateDiff("d", datLatTime, Now)
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckOutPathSend(ByVal lng�Һ�ID As Long) As Boolean
'���ܣ����ò����Ƿ����ɹ���Ŀ
'���أ�true=���ɹ���false=δ���ɹ�
    Dim strSql As String, rsPati As Recordset
    
    strSql = " Select Max(a.״̬) As ״̬" & vbNewLine & _
             " From ��������·�� A, ��������·����¼ B" & vbNewLine & _
             " Where a.Id = b.·����¼id And b.�Һ�id = [1]"
             
    On Error GoTo errH
    Set rsPati = zlDatabase.OpenSQLRecord(strSql, "CheckOutPathSend", lng�Һ�ID)
    If rsPati.RecordCount > 0 Then
        If Val(NVL(rsPati!״̬)) <> 0 Then
            CheckOutPathSend = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetNextPhaseOut(ByVal lng�׶�ID As Long) As Long
'���ܣ���ȡָ���׶εĺ����׶�ID
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select ��ID From ����·���׶� Where id = [1] And ��ID is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��һ�׶�", lng�׶�ID)
    If rsTmp.RecordCount > 0 Then
        lng�׶�ID = Val(rsTmp!��ID)
    End If
    
    strSql = "Select b.ID From ����·���׶� a,����·���׶� b " & _
            "Where a.·��ID= b.·��ID And a.�汾��= b.�汾�� And b.���>a.��� And a.ID = [1] And b.��ID Is Null And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��һ�׶�", lng�׶�ID)
    
    If rsTmp.RecordCount > 0 Then GetNextPhaseOut = Val(rsTmp!ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEPRDefineTextOut(Optional ByVal str����IDs As String, Optional ByVal lng��ĿID As Long) As String
'���ܣ���ȡ·����Ŀ��Ӧ�Ĳ�����������������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    If lng��ĿID <> 0 Then '�°���Ӳ������ϰ�ͬʱ
        strSql = " Select Nvl(a.����, b.����) as ���� From ����·������ A, �����ļ��б� B Where a.��Ŀid = [2] And a.�ļ�id = b.Id(+)" & vbNewLine & _
                 " Order by a.���"
    ElseIf str����IDs <> "" And lng��ĿID = 0 Then '�ϰ�
        strSql = " Select /*+ Rule*/ ���� From �����ļ��б�" & _
                 " Where ID IN(Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
                 " Order by ���"
    Else     '�°�
        strSql = "select ���� from ����·������ t where t.��Ŀid=[2] and t.�ļ�id is null and t.ԭ��id IN (Select Column_Value From Table(Cast(f_Str2list([1]) As zlTools.t_Strlist))) order by ���"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetEPRDefineTextOut", str����IDs, lng��ĿID)
    
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = strSql & "��" & rsTmp!����
        rsTmp.MoveNext
    Loop
    
    GetEPRDefineTextOut = Mid(strSql, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Checkҽ����ĿOut(ByVal lngִ��ID As Long) As Boolean
'���ܣ����ָ����ִ����Ŀ�Ƿ�����ҽ����
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select 1 From ��������·��ҽ�� Where ·��ִ��ID = [1] And Rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Checkҽ����ĿOut", lngִ��ID)
    
    Checkҽ����ĿOut = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function CheckSameDayOfPhaseOut(ByVal lngPhase As Long, ByVal lngDay As Long) As Boolean
'���ܣ���鵱���Ƿ������õ����������׶�(��ǰ�׶μ���֧����)
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    '�����ǰ�Ƿ�֧�׶Σ���ȡ�丸ID
    strSql = "Select ��ID From ����·���׶� Where ID = [1] And ��ID is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�׶�", lngPhase)
    If rsTmp.RecordCount > 0 Then lngPhase = rsTmp!��ID
    
    strSql = "Select 1" & vbNewLine & _
            "From ����·���׶� A, ����·���׶� B" & vbNewLine & _
            "Where a.Id = [1] And a.·��id = b.·��id And a.�汾�� = b.�汾�� And b.��� > a.���" & vbNewLine & _
            "  And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�׶�", lngPhase, lngDay)
    If rsTmp.RecordCount > 0 Then CheckSameDayOfPhaseOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInPathOut(ByVal lng����·��Id As Long) As Date
'���ܣ���ȡ���˵Ľ���·���Ŀ�ʼʱ��
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select a.��ʼʱ�� From ��������·�� a Where a.Id =[1] "
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�뾶ʱ��", lng����·��Id)
    If IsNull(rsTmp!��ʼʱ��) Then
        GetPatiInPathOut = zlDatabase.Currentdate
    Else
        GetPatiInPathOut = rsTmp!��ʼʱ��
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInfoOut(lng����ID As Long, lng�Һ�ID As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = " Select Nvl(b.����, a.����) ����, Nvl(b.�Ա�, a.�Ա�) �Ա�, Nvl(b.����, a.����) ����, To_Char(a.��������, 'yyyy-mm-dd hh24:mi:ss') ��������, b.�����," & vbNewLine & _
             "       b.ִ��״̬,B.����ʱ��, B.���ʱ��, c.���� As ����" & vbNewLine & _
             " From ������Ϣ A, ���˹Һż�¼ B, ���ű� C" & vbNewLine & _
             " Where a.����id = b.����id And b.����id = [1] And b.Id = [2] And b.ִ�в���id = c.Id"
    On Error GoTo errH
    Set GetPatiInfoOut = zlDatabase.OpenSQLRecord(strSql, "��ȡ��������", lng����ID, lng�Һ�ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceOut(strIDs As String) As ADODB.Recordset
'���ܣ���ȡ·����Ŀ��Ӧ��ҽ����¼��
    Dim strSql As String
 
    strSql = " Select /*+ rule*/ a.·����ĿID,a.ҽ������ID,b.��Ч,Nvl(b.���ID,b.ID) ���ID,b.������ĿID" & vbNewLine & _
             " From ����·��ҽ�� A,����·��ҽ������ B,(Select Column_Value As ID From Table(f_Num2list([1]))) C" & vbNewLine & _
             " Where a.ҽ������id=b.id And a.·����Ŀid = c.Id" & vbNewLine & _
             " Order by b.���"
    On Error GoTo errH
    Set GetAdviceOut = zlDatabase.OpenSQLRecord(strSql, "��ȡҽ����¼", strIDs)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMustDayOut(ByVal lng����·��Id As Long, ByVal lng��ǰ���� As Long, Optional ByVal blnIsNotMinus As Boolean) As Long
'���ܣ���ȡ����·��ִ�������ϵĵ�ǰ���� (=��ǰʵ������-�����ӳٵ�����+��ǰ����(�п���һ����ǰ����))
'������blnIsNotMinus=�Ƿ񲻼�ȥ�ӳ�ʱ�䣨����ʱ��ǰ������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lng�ӳ����� As Long
    Dim lng��ǰ���� As Long
    Dim i As Long
    Dim lng�׶�ʵ������ As Long
    Dim lng�׶ο�ʼ���� As Long
    Dim byt��ǰ���� As Byte
    
    On Error GoTo errH

    strSql = " Select Max(Decode(A.ʱ�����, 1, 1, 2, 2, 0)) As �׶��Ƿ���ǰ, C.��ʼ����, Nvl(C.��������, C.��ʼ����) As ��������," & vbNewLine & _
             "        Sum(Decode(A.ʱ�����, -1, 1, 0)) As �׶��Ӻ�����, Count(1) As �׶�ʵ������" & vbNewLine & _
             " From ��������·������ A, ����·���׶� C, ����·���׶� D" & vbNewLine & _
             " Where a.�׶�id = c.Id And c.��id = d.Id(+) And" & vbNewLine & _
             "      a.·����¼id = [1]" & vbNewLine & _
             " Group By c.��ʼ����, Nvl(c.��������, c.��ʼ����), a.�׶�id,d.���, c.���" & vbNewLine & _
             " Order By Nvl(d.���, c.���) "

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetMustDayOut", lng����·��Id)

    For i = 0 To rsTmp.RecordCount - 1
        '�ӳ�����
        lng�ӳ����� = lng�ӳ����� + Val(rsTmp!�׶��Ӻ����� & "")
        '��ǰ����
        If Val(rsTmp!�׶��Ƿ���ǰ & "") = 1 Or Val(rsTmp!�׶��Ƿ���ǰ & "") = 2 Then
            '���һ���׶�����ǰ�����1�죬��Ϊ����֪�������ѡ��һ���׶�
            If i = rsTmp.RecordCount - 1 Or rsTmp!��ʼ���� & "" = rsTmp!�������� & "" Then
                If Val(rsTmp!�׶��Ƿ���ǰ & "") = 1 Then
                    lng��ǰ���� = lng��ǰ���� + 1
                ElseIf Val(rsTmp!�׶��Ƿ���ǰ & "") = 2 Then
                    '��һ�׶���ǰ������,��ʱ����Ҫ����һ�׶���ǰ�����족�ٶ����һ��
                End If
                rsTmp.MoveNext
            Else
                '�ȼ�¼�½׶�ʵ�������Ϳ�ʼ����
                lng�׶ο�ʼ���� = Val(rsTmp!��ʼ���� & "")
                lng�׶�ʵ������ = Val(rsTmp!�׶�ʵ������ & "")
                byt��ǰ���� = Val(rsTmp!�׶��Ƿ���ǰ & "")
                rsTmp.MoveNext
                lng��ǰ���� = lng��ǰ���� + (Val(rsTmp!��ʼ���� & "") - lng�׶ο�ʼ���� - lng�׶�ʵ������ + IIf(byt��ǰ���� = 2, 0, 1))
            End If
        Else
            rsTmp.MoveNext
        End If
    Next
    
    GetMustDayOut = lng��ǰ���� - IIf(blnIsNotMinus, 0, lng�ӳ�����) + lng��ǰ����
        
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPathOutLogOut() As Boolean
'���ܣ�����Ƿ���ڲ��˳����Ǽ���Ŀ
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From ����·������ṹ Where ����ID = 2 And Rownum=1"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����·������ṹ")
    CheckPathOutLogOut = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckDelOutPathItem(ByVal lngִ��ID As Long) As Boolean
'���ܣ����ָ����ҽ����·����Ŀִ�м�¼�Ƿ����ɾ������������
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strIDs As String
    Dim i As Long

    '���ǵ������ɵĳ������������ɺ��Զ�ֹͣ�������Ƿ��ͣ�
    '�ǵ������ɵĳ�������У�Ե�δ���ϣ�������ȡ��(��ֹͣ��Ҳ������)��δУ�Եģ�ȡ��ʱ�Զ�ɾ����Ӧ��ҽ����
    strSql = "Select 1 From ��������·��ҽ�� A, ��������·��ҽ�� B" & vbNewLine & _
             "Where a.·��ִ��id = [1] And a.����ҽ��id = b.����ҽ��id And b.·��ִ��id <> a.·��ִ��id  And rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ҽ��", lngִ��ID)
    If rsTmp.RecordCount = 0 Then '��������
        strSql = "Select 1 From ��������·��ҽ�� B, ��������ҽ����¼ C Where b.·��ִ��id = [1] And b.����ҽ��id = c.Id And c.ҽ��״̬ > 1 And c.ҽ��״̬ <> 4 And rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ҽ��", lngִ��ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "����Ŀ������У�Ե�δ���ϵ�ҽ������������ҽ������ִ�д˲�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckDelOutPathItem = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get���ﲡ��ID(ByVal lng����ID As Long, ByVal lng�Һ�ID As Long, ByVal lngType As Long, Optional ByVal lng����ID As Long, Optional ByRef bln��ҽ As Boolean = False) As ADODB.Recordset
'������ lngType =0  ������ȡ�����ҽ�������;�������ҽ��ʱ, ���ȼ�����ҽ�������
'               =1  ȡ���˳���Ҫ���֮�����ϣ�������Ϸǵ�һ���
'               =2  ������ȡ�����ҽ�������;�������ҽ��ʱ, ���ȼ�����ҽ������������Ҫ��ϣ�ͬʱ����϶�Ӧ����Ҫ·����
'˵��:  ���ų�����¼������
    Dim rsTmp As ADODB.Recordset, strSql As String

    If lngType = 0 Then                                                             'ȡȫ��������
        bln��ҽ = Sys.DeptHaveProperty(lng����ID, "��ҽ��")
        If bln��ҽ Then
            strSql = " Select ����id,���id,�������,�������,��¼��Դ,��ϴ���" & vbNewLine & _
                     " From ������ϼ�¼" & vbNewLine & _
                     " Where ��¼��Դ In (1,3) And ������� In (1,11) And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And " & vbNewLine & _
                     "      Nvl(�Ƿ�����, 0) = 0 And Not (NVl(����ID,0)=0 and NVl(���ID,0)=0) " & vbNewLine & _
                     " Order By Decode(�������, 11, 1, 1, 2), Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc,��ϴ���"
        Else
            strSql = " Select ����id,���id,�������,�������,��¼��Դ,��ϴ���" & vbNewLine & _
                     " From ������ϼ�¼" & vbNewLine & _
                     " Where ��¼��Դ In (1,3) And ������� In (1,11) And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And " & vbNewLine & _
                     "       Nvl(�Ƿ�����,0) = 0 And Not (NVl(����ID,0)=0 and NVl(���ID,0)=0) " & vbNewLine & _
                     " Order By �������, Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc,��ϴ���"
        End If
    ElseIf lngType = 1 Then                                                             'ȡ����Ҫ���
        strSql = " Select ����id, ���id, �������, �������, ��¼��Դ,��ϴ���" & vbNewLine & _
                 " From ������ϼ�¼ " & vbNewLine & _
                 " Where ��¼��Դ In (1,3) And ������� In (1,11) And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� <> 1 Or" & vbNewLine & _
                 "      Nvl(�Ƿ�����, 0) = 0 And Not (NVl(����ID,0)=0 and NVl(���ID,0)=0) " & vbNewLine & _
                 " Order By �������, Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc,��ϴ���"
    ElseIf lngType = 2 Then
'        bln��ҽ = Sys.DeptHaveProperty(lng����ID, "��ҽ��")
'        If bln��ҽ Then
'            strSql = " Select Distinct a.Id, k.����id, k.���id, k.�������, K.�������, K.��¼��Դ,k.���� " & vbNewLine & _
'                     " From ����·��Ŀ¼ A, ����·������ B, ����·���汾 C," & vbNewLine & _
'                     "     (Select Rownum As ����, ����id, ���id, �������, �������, ��¼��Դ " & vbNewLine & _
'                     "       From ������ϼ�¼" & vbNewLine & _
'                     "       Where ��¼��Դ In (1, 3) And ������� In (1,11) And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� <> 1 And" & vbNewLine & _
'                     "             Nvl(�Ƿ�����, 0) = 0 And Not (Nvl(����id, 0) = 0 And Nvl(���id, 0) = 0)" & vbNewLine & _
'                     "       Order By Decode(�������,11, 3, 1, 4), Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc, ��ϴ���) K" & vbNewLine & _
'                     " Where a.Id = b.·��id And a.Id = b.·��id And a.Id = c.·��id And a.���°汾 = c.�汾�� And a.���� = 0 And b.���� = 0 And" & vbNewLine & _
'                     "      (b.����id = k.����id Or b.���id = k.���id) And" & vbNewLine & _
'                     "      (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From ����·������ D Where a.Id = d.·��id And d.����id = [3]))" & vbNewLine & _
'                     " Order By k.����"
'        Else
'            strSql = " Select Distinct a.Id, k.����id, k.���id, k.�������,K.�������, K.��¼��Դ,k.���� " & vbNewLine & _
'                     " From ����·��Ŀ¼ A, ����·������ B, ����·���汾 C," & vbNewLine & _
'                     "     (Select Rownum As ����, ����id, ���id, �������, �������, ��¼��Դ " & vbNewLine & _
'                     "       From ������ϼ�¼" & vbNewLine & _
'                     "       Where ��¼��Դ In (1,3) And ������� In (1,11) And ȡ��ʱ�� Is Null And ����id = [1] And ��ҳid = [2] And ��ϴ��� <> 1 And" & vbNewLine & _
'                     "             Nvl(�Ƿ�����, 0) = 0 And Not (Nvl(����id, 0) = 0 And Nvl(���id, 0) = 0)" & vbNewLine & _
'                     "       Order By Sign(������� - 10), ������� Desc, Decode(��¼��Դ, 1, 4, ��¼��Դ) Desc, ��ϴ���) K" & vbNewLine & _
'                     " Where a.Id = b.·��id And a.Id = b.·��id And a.Id = c.·��id And a.���°汾 = c.�汾�� And a.���� = 0 And b.���� = 0 And" & vbNewLine & _
'                     "      (b.����id = k.����id Or b.���id = k.���id) And" & vbNewLine & _
'                     "      (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From ����·������ D Where a.Id = d.·��id And d.����id = [3]))" & vbNewLine & _
'                     " Order By k.����"
'        End If
    End If
    '��¼��Դ:1-������3-��ҳ����
    '�������:1-��ҽ�������;11-��ҽ�������
    '�ж����ϵ�����£�������ϴ���ֻȡ��һ����Ҫ���
    '���������������ȣ���Ҫ��Ϊ��֧��������ϡ�
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����", lng����ID, lng�Һ�ID, lng����ID)
    Set Get���ﲡ��ID = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInDateOut(t_pati As TYPE_Pati) As Date
'���ܣ���ȡ���˵ľ���ʱ��
'���أ�����ʱ��
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = " Select ִ��ʱ�� As ��ʼʱ�� " & vbNewLine & _
             "       From ���˹Һż�¼" & vbNewLine & _
             "       Where ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����ʱ��", t_pati.�Һ�ID)
    GetPatiInDateOut = CDate(rsTmp!��ʼʱ��)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetOutPathTable(ByVal lng����ID As Long, ByVal lng���ID As Long, ByVal lng����ID As Long, Optional ByVal str����IDs As String, _
                                Optional ByVal str���IDs As String, Optional ByVal lng����ID As Long, Optional ByVal lng�Һ�ID As Long) As ADODB.Recordset
    Dim strSql As String
    
    If str����IDs = "" And str���IDs = "" Then
        '�����Distinct����Ϊ�����id�ͼ���id���˰󶨶�Ӧ�����Բ���������ظ�ֵ
        strSql = " Select Distinct a.Id, a.����, a.����, a.����, a.˵��, a.�����Ա�, a.��������, a.���°汾, c.��׼����ʱ�� " & vbNewLine & _
                 " From ����·��Ŀ¼ A, ����·������ B,����·���汾 C" & vbNewLine & _
                 " Where a.Id = b.·��id And (b.����id = [1] Or b.���id = [2]) And a.���°汾 is not null And a.id = b.·��ID And a.���°汾 = c.�汾��" & vbNewLine & _
                 " And a.Id = c.·��id And (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From ����·������ D Where a.Id = d.·��id And d.����id = [3]))"
    Else
        strSql = " Select Distinct a.Id, a.����, a.����, a.����, a.˵��, a.�����Ա�, a.��������, a.���°汾, c.��׼����ʱ�� " & vbNewLine & _
                 " From ����·��Ŀ¼ A, ����·������ B,����·���汾 C" & vbNewLine & _
                 " Where a.Id = b.·��id And (instr(',' || [4] || ',',',' || b.����ID || ',')>0 " & vbNewLine & _
                 " And [4] is not null Or instr(',' || [5] || ',',',' || b.���ID || ',')>0 and [5] is not null) " & vbNewLine & _
                 " And a.���°汾 is not null And a.id = b.·��ID And a.���°汾 = c.�汾��" & vbNewLine & _
                 " And a.Id = c.·��id And (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From ����·������ D Where a.Id = d.·��id And d.����id = [3]))" & _
                 " And Not Exists(Select 1 From ��������·�� D Where a.ID=d.·��ID And d.����ID=[6] And D.�Һ�ID=[7])"
    End If
    On Error GoTo errH
    Set GetOutPathTable = zlDatabase.OpenSQLRecord(strSql, "��ȡ·��Ŀ¼", lng����ID, lng���ID, lng����ID, str����IDs, str���IDs, lng����ID, lng�Һ�ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckPatiPathOutLogOut(ByVal lng·����¼ID As Long) As Boolean
'���ܣ�����Ƿ���ڲ��˳�����¼
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From �������������¼ Where ·����¼ID=[1] And Rownum=1"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˳�����¼", lng·����¼ID)
    CheckPatiPathOutLogOut = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get�׶η���Out(Optional ByVal lng·����¼ID As Long, Optional ByVal lng�׶�ID As Long) As String
'���ܣ���ȡ����ʹ�ù��Ľ׶εķ��ֻ࣬�з�֧·�����з��࣬���ʹ���˸÷��࣬��������·���ڼ�ֻ��ѡ��÷��࣬����ֻ������һ������
'������lng�׶�ID=ָ���ò���ʱ����ȡָ���׶εķ���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH
    If lng�׶�ID <> 0 Then
        strSql = "Select ���� From ����·���׶� Where id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�׶η���", lng�׶�ID)
    Else
        strSql = " Select a.����" & vbNewLine & _
                 " From ����·���׶� A, (Select Distinct �׶�id From ��������·��ִ�� Where ·����¼id = [1]) B" & vbNewLine & _
                 " Where a.Id = b.�׶�id And a.���� Is Not Null And rownum<2"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�׶η���", lng·����¼ID)
    End If
    If rsTmp.RecordCount > 0 Then
        Get�׶η���Out = "" & rsTmp!����
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPhaseNOOut(ByVal lng�׶�ID As Long) As Long
'���ܣ���ȡָ���׶ε����(����ý׶��Ƿ�֧����ȡ���׶ε����)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select ��ID From ����·���׶� Where id = [1] And ��ID is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�׶����", lng�׶�ID)
    If rsTmp.RecordCount > 0 Then
        lng�׶�ID = Val(rsTmp!��ID)
    End If
    
    strSql = "Select ��� From ����·���׶� Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�׶����", lng�׶�ID)
    If rsTmp.RecordCount > 0 Then
        GetPhaseNOOut = Val(rsTmp!���)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLastPhaseNOOut(ByVal lng����·��Id As Long, ByVal lng·��ID As Long)
'���ܣ���ȡ����ָ��·�����һ���׶ε����
Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = " Select Max(Nvl(c.���, b.���)) ���" & vbNewLine & _
             " From ��������·��ִ�� A, ����·���׶� B, ����·���׶� C" & vbNewLine & _
             " Where a.·����¼id = [1] And a.�׶�id = b.Id And b.·��id = [2] And b.��id = c.Id(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�׶����", lng����·��Id, lng·��ID)
    
    GetLastPhaseNOOut = Val("" & rsTmp!���)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathDiagOut(ByVal lng�Һ�ID As Long, ByVal lng�����Դ As Long, ByVal lngDiagType As Long, _
    ByVal lngDiag As Long, ByVal lng���ID As Long) As Boolean
'���ܣ��������·����Ӧ����ϲ����޸�
'������lngDiagType���������,lngDiag=����ID
'����ֵ:F-�������޸�;T-�����޸�
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select a.�������, a.����id, a.���id, a.�����Դ" & vbNewLine & _
            "From ��������·�� A, ��������·����¼ B" & vbNewLine & _
            "Where a.Id = b.·����¼id And b.�Һ�id = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, gstrSysName, lng�Һ�ID)
    Do While Not rsTmp.EOF
        If lngDiagType = Val(rsTmp!������� & "") And lng�����Դ = Val(rsTmp!�����Դ & "") And (lngDiag = Val(rsTmp!����id & "") Or lng���ID = Val(rsTmp!���id & "")) Then
            Exit Function
        End If
        rsTmp.MoveNext
    Loop
    CheckPathDiagOut = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
