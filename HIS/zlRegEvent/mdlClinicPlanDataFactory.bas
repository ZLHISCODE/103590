Attribute VB_Name = "mdlClinicPlanDataFactory"
Option Explicit

Public Function GetSignalSourceObject(ByVal rsRecord As ADODB.Recordset) As �����Դ
    '����:����¼���ĵ�ǰ��¼ת��Ϊ"�����Դ"����
    '��Σ�
    '   rsRecord - ��Դ��¼��
    Dim obj�����Դ As New �����Դ
    
    Err = 0: On Error GoTo errHandler
    If Not rsRecord.EOF Then
        With obj�����Դ
            .ID = Val(Nvl(rsRecord!ID))
            .���� = Nvl(rsRecord!����)
            .���� = Nvl(rsRecord!����)
            .����ID = Val(Nvl(rsRecord!����ID))
            .�������� = Nvl(rsRecord!��������)
            .��ĿID = Val(Nvl(rsRecord!��ĿID))
            .��Ŀ���� = Nvl(rsRecord!��Ŀ����)
            .ҽ��ID = Val(Nvl(rsRecord!ҽ��ID))
            .ҽ������ = Nvl(rsRecord!ҽ������)
            .ҽ��ְ�� = Nvl(rsRecord!ҽ��ְ��)
            .�Ƿ񽨲��� = Val(Nvl(rsRecord!�Ƿ񽨲���)) = 1
            .ԤԼ���� = Val(Nvl(rsRecord!ԤԼ����))
            .����Ƶ�� = Val(Nvl(rsRecord!����Ƶ��))
            .���տ���״̬ = Val(Nvl(rsRecord!���տ���״̬))
            .�Ƿ��ٴ��Ű� = Val(Nvl(rsRecord!�Ƿ��ٴ��Ű�)) = 1
            .�Ű෽ʽ = Val(Nvl(rsRecord!�Ű෽ʽ))
            .�Ƿ�ɾ�� = Val(Nvl(rsRecord!�Ƿ�ɾ��)) = 1
            .����ʱ�� = Format(Nvl(rsRecord!����ʱ��), "yyyy-mm-dd hh:mm:ss")
            .����ʱ�� = Format(Nvl(rsRecord!����ʱ��), "yyyy-mm-dd hh:mm:ss")
            .վ�� = Nvl(rsRecord!վ��)
        End With
    End If
    Set GetSignalSourceObject = obj�����Դ
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetWorkTimesObjects(ByVal rsRecord As ADODB.Recordset) As �ϰ�ʱ�μ�
    '����:����¼��ת��Ϊ"�ϰ�ʱ�μ�"����
    '��Σ�
    '   rsRecord - �ϰ�ʱ�μ�¼
    Dim obj�ϰ�ʱ�μ� As New �ϰ�ʱ�μ�, obj�ϰ�ʱ�� As �ϰ�ʱ��
    Err = 0: On Error GoTo errHandler
        
    '���⴦��
    'ȡ��ͬʱ�εĵ�һ��
    rsRecord.Sort = "վ�� Desc,���� Desc"
    If rsRecord.RecordCount > 0 Then rsRecord.MoveFirst
    Do While Not rsRecord.EOF
        If obj�ϰ�ʱ�μ�.Exits("K" & Nvl(rsRecord!ʱ���)) = False Then
            Set obj�ϰ�ʱ�� = New �ϰ�ʱ��
            With obj�ϰ�ʱ��
                .ʱ��� = Nvl(rsRecord!ʱ���)
                .��ʼʱ�� = Format(Nvl(rsRecord!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
                .����ʱ�� = Format(Nvl(rsRecord!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
                .ȱʡԤԼʱ�� = Format(Nvl(rsRecord!ȱʡʱ��), "yyyy-mm-dd hh:mm:ss")
                .��ǰ�Һ�ʱ�� = Format(Nvl(rsRecord!��ǰʱ��), "yyyy-mm-dd hh:mm:ss")
                .����Ԥ��ʱ�� = Val(Nvl(rsRecord!����Ԥ��ʱ��))
                .��Ϣʱ�� = Nvl(rsRecord!��Ϣʱ��)
            End With
            obj�ϰ�ʱ�μ�.AddItem obj�ϰ�ʱ��, "K" & Nvl(rsRecord!ʱ���)
        End If
        rsRecord.MoveNext
    Loop
    Set GetWorkTimesObjects = obj�ϰ�ʱ�μ�
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetUnitsObjects(ByVal rsRecord As ADODB.Recordset) As ������λ���Ƽ�
    '����:����¼��ת��Ϊ"������λ���Ƽ�"����
    '��Σ�
    '   rsRecord - ������λ���Ƽ�¼
    Dim obj������λ���Ƽ� As New ������λ���Ƽ�, obj������λ���� As ������λ����
    
    Err = 0: On Error GoTo errHandler
    Do While Not rsRecord.EOF
        Set obj������λ���� = New ������λ����
        With obj������λ����
            .���� = Nvl(rsRecord!����)
            .������λ���� = Nvl(rsRecord!����)
        End With
        obj������λ���Ƽ�.AddItem obj������λ����, "K" & Nvl(rsRecord!����)
        rsRecord.MoveNext
    Loop
    Set GetUnitsObjects = obj������λ���Ƽ�
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetVisitPlanObjects(ByVal rsRecord As ADODB.Recordset) As ���ﰲ��
    '����:����¼��ת��Ϊ"���ﰲ��"����
    '���
    Dim obj���ﰲ�� As New ���ﰲ��
    
    Err = 0: On Error GoTo errHandler
    If Not rsRecord.EOF Then
        With obj���ﰲ��
            .����ID = Val(Nvl(rsRecord!����ID))
            .������� = Nvl(rsRecord!ҽ������)
            .�Ű෽ʽ = Val(Nvl(rsRecord!�Ű෽ʽ))
            .��� = Val(Nvl(rsRecord!���))
            .�·� = Val(Nvl(rsRecord!�·�))
            .���� = Val(Nvl(rsRecord!����))
            .Ӧ�÷�Χ = Val(Nvl(rsRecord!Ӧ�÷�Χ))
            .����ID = Val(Nvl(rsRecord!����ID))
            .��ע = Nvl(rsRecord!��ע)
            .������ = Nvl(rsRecord!������)
            .����ʱ�� = Format(Nvl(rsRecord!����ʱ��), "yyyy-mm-dd hh:mm:ss")
            .ģ������ = Val(Nvl(rsRecord!ģ������))
            
            .����ID = Val(Nvl(rsRecord!����ID))
            .��ĿID = Val(Nvl(rsRecord!��ĿID))
            .��Ŀ���� = Nvl(rsRecord!��Ŀ����)
            .ҽ��ID = Val(Nvl(rsRecord!ҽ��ID))
            .ҽ������ = Nvl(rsRecord!ҽ������)
            .ҽ��ְ�� = Nvl(rsRecord!ҽ��ְ��)
            .�Ű���� = Val(Nvl(rsRecord!�Ű����))
            .���������� = Val(Nvl(rsRecord!�Ƿ���������)) = 0
            .���ղ����� = Val(Nvl(rsRecord!�Ƿ����ճ���)) = 0
            .��ʼʱ�� = Format(Nvl(rsRecord!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
            .��ֹʱ�� = Format(Nvl(rsRecord!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
            .����Ա���� = Nvl(rsRecord!����Ա����)
            .�Ǽ�ʱ�� = Format(Nvl(rsRecord!�Ǽ�ʱ��), "yyyy-mm-dd hh:mm:ss")
            .�Ƿ���ʱ���� = Val(Nvl(rsRecord!�Ƿ���ʱ����)) = 1
        End With
    End If
    Set GetVisitPlanObjects = obj���ﰲ��
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitTimesObjects(ByVal rsRecord As ADODB.Recordset) As �����¼��
    '����:����¼��ת��Ϊ"�����¼��"����
    '��Σ�
    Dim obj�����¼�� As New �����¼��, obj�����¼ As �����¼
    
    Err = 0: On Error GoTo errHandler
    Do While Not rsRecord.EOF
        Set obj�����¼ = GetVisitTimesObject(rsRecord)
        obj�����¼��.AddItem obj�����¼, "K" & obj�����¼.ʱ���
        rsRecord.MoveNext
    Loop
    Set GetVisitTimesObjects = obj�����¼��
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitTimesObject(ByVal rsRecord As ADODB.Recordset) As �����¼
    '����:����¼��ת��Ϊ"�����¼��"����
    '��Σ�
    Dim obj�����¼ As New �����¼
    
    Err = 0: On Error GoTo errHandler
    If Not rsRecord.EOF Then
        With obj�����¼
            .��¼ID = Val(Nvl(rsRecord!��¼ID))
            .ʱ��� = Nvl(rsRecord!�ϰ�ʱ��)
            .�Ƿ��ʱ�� = Val(Nvl(rsRecord!�Ƿ��ʱ��)) = 1
            .�Ƿ���ſ��� = Val(Nvl(rsRecord!�Ƿ���ſ���)) = 1
            .�޺��� = Val(Nvl(rsRecord!�޺���))
            'ԤԼ���ƣ�0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ;
            If Val(Nvl(rsRecord!ԤԼ����)) = 1 Then
                .��Լ�� = 0
            Else
                .��Լ�� = IIf(Val(Nvl(rsRecord!��Լ��)) = 0, Val(Nvl(rsRecord!�޺���)), Val(Nvl(rsRecord!��Լ��)))
            End If
            .ԤԼ���� = Val(Nvl(rsRecord!ԤԼ����))
            .���﷽ʽ = Val(Nvl(rsRecord!���﷽ʽ))
            .ԤԼ���� = Val(Nvl(rsRecord!ԤԼ����))
            .�������� = Format(Nvl(rsRecord!��������), "yyyy-mm-dd hh:mm:ss")
            .�ѹ��� = Val(Nvl(rsRecord!�ѹ���))
            .��Լ�� = Val(Nvl(rsRecord!��Լ��))
            .����ҽ�� = Nvl(rsRecord!����ҽ������)
            .����ID = Val(Nvl(rsRecord!����ID))
            .��ĿID = Val(Nvl(rsRecord!��ĿID))
            .��Ŀ���� = Nvl(rsRecord!��Ŀ����)
            .ҽ��ID = Val(Nvl(rsRecord!ҽ��ID))
            .ҽ������ = Nvl(rsRecord!ҽ������)
            .��ʼʱ�� = Format(Nvl(rsRecord!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
            .��ֹʱ�� = Format(Nvl(rsRecord!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
            .�Ƿ��ռ = Val(Nvl(rsRecord!�Ƿ��ռ)) = 1
            .ͣ�￪ʼʱ�� = Format(Nvl(rsRecord!ͣ�￪ʼʱ��), "yyyy-mm-dd hh:mm:ss")
            .ͣ����ֹʱ�� = Format(Nvl(rsRecord!ͣ����ֹʱ��), "yyyy-mm-dd hh:mm:ss")
            .ͣ��ԭ�� = Nvl(rsRecord!ͣ��ԭ��)
            .�Ƿ���ʱ���� = Val(Nvl(rsRecord!�Ƿ���ʱ����)) = 1
        End With
    End If
    Set GetVisitTimesObject = obj�����¼
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitRoomsObjects(ByVal rsRecord As ADODB.Recordset) As �������Ҽ�
    '����:����¼��ת��Ϊ"�������Ҽ�"����
    '��Σ�
    Dim obj�������Ҽ� As New �������Ҽ�, obj�������� As ��������
    
    Err = 0: On Error GoTo errHandler
    Do While Not rsRecord.EOF
        Set obj�������� = New ��������
        With obj��������
            .����ID = Nvl(rsRecord!����ID)
            .�������� = Nvl(rsRecord!����)
        End With
        obj�������Ҽ�.AddItem obj��������, "K" & Nvl(rsRecord!����ID)
        rsRecord.MoveNext
    Loop
    Set GetVisitRoomsObjects = obj�������Ҽ�
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetTimeIntervalObjects(ByVal rsRecord As ADODB.Recordset) As ������Ϣ��
    '����:����¼��ת��Ϊ"������Ϣ��"����
    '��Σ�
    Dim obj������Ϣ�� As New ������Ϣ��, obj������Ϣ As ������Ϣ

    On Error GoTo errHandler
    Do While Not rsRecord.EOF
        Set obj������Ϣ = GetTimeIntervalObject(rsRecord)
        If Not obj������Ϣ Is Nothing Then
            obj������Ϣ��.AddItem obj������Ϣ
        End If
        rsRecord.MoveNext
    Loop
    Set GetTimeIntervalObjects = obj������Ϣ��
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetTimeIntervalObject(ByVal rsRecord As ADODB.Recordset) As ������Ϣ
    '����:����¼��ת��Ϊ"������Ϣ"����
    '��Σ�
    Dim obj������Ϣ As ������Ϣ

    On Error GoTo errHandler
    If rsRecord.EOF Then Exit Function
    Set obj������Ϣ = New ������Ϣ
    With obj������Ϣ
        .��� = Nvl(rsRecord!���)
        .��ʼʱ�� = Format(Nvl(rsRecord!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
        .��ֹʱ�� = Format(Nvl(rsRecord!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
        .���� = Val(Nvl(rsRecord!����))
        .�Ƿ�ԤԼ = Val(Nvl(rsRecord!�Ƿ�ԤԼ)) = 1
        .�Ƿ�ͣ�� = Val(Nvl(rsRecord!�Ƿ�ͣ��)) = 1
    End With
    Set GetTimeIntervalObject = obj������Ϣ
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ChangeCurPlan(obj���ﰲ�� As ���ﰲ��, ByVal strNewItem As String, _
    Optional ByVal blnRuleChanged As Boolean)
    '��ǰ������Ŀ�Ѹı䣬����δ������ﰲ�ż��ϣ�����ȡ��ǰ����
    Dim ObjItem As �����¼��, strKey As String
    Dim objTemp As �����¼��
    
    On Error GoTo Errhand
    If obj���ﰲ��.δ������ﰲ�� Is Nothing Then Set obj���ﰲ��.δ������ﰲ�� = New ���ﰲ��
    If obj���ﰲ��.�ѱ�����ﰲ�� Is Nothing Then Set obj���ﰲ��.�ѱ�����ﰲ�� = New ���ﰲ��
    'ģ��ʱ������仯ʱҪ���δ���氲��
    If blnRuleChanged Then
        obj���ﰲ��.δ������ﰲ��.RemoveAll
    Else
        For Each ObjItem In obj���ﰲ��
            Set objTemp = ObjItem.Clone
            If objTemp.�������� <> "" Then  '����Ŀ�Ĳ�����
                strKey = GetPlanKey(objTemp.��������)
                
                '��δ�����¼���д��ڣ�����ɾ��
                If obj���ﰲ��.δ������ﰲ��.Exits(strKey) Then obj���ﰲ��.δ������ﰲ��.Remove strKey
                
                'һ��ʱ�ζ�û�еĲ�����
                If objTemp.Count = 0 Then
                    '������ѱ����¼���д��ڣ���������ʱ�α�ʾ��ɾ��
                    If obj���ﰲ��.�ѱ�����ﰲ��.Exits(strKey) Then
                        obj���ﰲ��.�ѱ�����ﰲ��(strKey).�Ƿ�ɾ�� = True
                    End If
                ElseIf objTemp.�Ƿ��޸� Then 'δ�޸ĵĲ�����
                    obj���ﰲ��.δ������ﰲ��.AddItem objTemp, strKey
                End If
            End If
        Next
    End If
    
    '�л���ǰ���ﰲ��
    obj���ﰲ��.RemoveAll
    strKey = GetPlanKey(strNewItem)
    If obj���ﰲ��.δ������ﰲ��.Exits(strKey) Then
        obj���ﰲ��.AddItem obj���ﰲ��.δ������ﰲ��(strKey).Clone, strKey
    ElseIf obj���ﰲ��.�ѱ�����ﰲ��.Exits(strKey) Then
        If obj���ﰲ��.�ѱ�����ﰲ��(strKey).�Ƿ�ɾ�� Then
            Set ObjItem = New �����¼��
            ObjItem.�������� = strNewItem
            obj���ﰲ��.AddItem ObjItem, strKey
        Else
            obj���ﰲ��.AddItem obj���ﰲ��.�ѱ�����ﰲ��(strKey).Clone, strKey
        End If
    ElseIf obj���ﰲ��.�ѱ�����ﰲ��.Count > 0 _
        And obj���ﰲ��.�ѱ�����ﰲ��.�Ű���� = obj���ﰲ��.�Ű���� _
        And InStr("4,5", obj���ﰲ��.�Ű����) > 0 Then
        '���ı䰲�ţ�ֻ�ı������Ŀ
        Set ObjItem = obj���ﰲ��.�ѱ�����ﰲ��(1).Clone
        ObjItem.�������� = strNewItem
        obj���ﰲ��.AddItem ObjItem, strKey
    Else
        Set ObjItem = New �����¼��
        ObjItem.�������� = strNewItem
        obj���ﰲ��.AddItem ObjItem, strKey
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetStopVisitObjects(ByVal rsRecord As ADODB.Recordset) As ͣ���¼��
    '����:��ͣ���¼��ת��Ϊ"ͣ���¼��"����
    '��Σ�
    Dim objͣ���¼�� As New ͣ���¼��, objͣ���¼ As ͣ���¼

    On Error GoTo errHandler
    Do While Not rsRecord.EOF
        Set objͣ���¼ = New ͣ���¼
        With objͣ���¼
            .��ʼʱ�� = Format(Nvl(rsRecord!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
            .��ֹʱ�� = Format(Nvl(rsRecord!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
            .ͣ��ԭ�� = Nvl(rsRecord!ͣ��ԭ��)
            .���� = Nvl(rsRecord!����)
        
            objͣ���¼��.AddItem objͣ���¼
        End With
        rsRecord.MoveNext
    Loop
    Set GetStopVisitObjects = objͣ���¼��
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
