Attribute VB_Name = "mdlClinicPlanData"
Option Explicit
Public grsWorkTime As ADODB.Recordset '�����ϰ�ʱ�Σ�����
Public grsUnit As ADODB.Recordset '���к�����λ��ԤԼ��ʽ������
Public Function LpadTime(ByVal strStartTime As String, ByVal strEndTime As String, ByRef dtStartDate As Date, ByRef dtEndDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:strStartTime-��ʼʱ��,��ʽΪHH:MM:SS
    '     strEndTime-��ʼʱ��,��ʽΪHH:MM:SS
    '����:dtStartDate-��ʼʱ��,��yyyy-mm-dd hh:mm:ss
    '     dtEndDate-��ֹʱ��,��yyyy-mm-dd hh:mm:ss
    '����:�������ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2016-03-24 14:50:32
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCurDate As String
    Dim dtStart As Date, dtEnd As Date
    
    On Error GoTo errHandle
    
    strCurDate = Format(Date, "yyyy-mm-dd")
    If strStartTime = "" Or strEndTime = "" Then Exit Function
    strStartTime = Format(strStartTime, "HH:MM:SS")
    strEndTime = Format(strEndTime, "HH:MM:SS")
    
    dtStart = CDate(strCurDate & " " & strStartTime)
    dtEnd = CDate(strCurDate & " " & strEndTime)
    If dtStart >= dtEnd Then dtEnd = dtEnd + 1
    
    dtStartDate = dtStart
    dtEndDate = dtEnd
    LpadTime = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetWorkTimeRange(strʱ��� As String, ByVal strվ�� As String, ByVal str���� As String) As �ϰ�ʱ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ�������ȡ�ϰ�ʱ�ζ�����Ϣ
    '���:strʱ���-ʱ�������
    '����:�����ϰ�ʱ����Ϣ
    '����:���˺�
    '����:2016-03-24 16:03:25
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj�ϰ�ʱ�� As �ϰ�ʱ��, strStart As String, strEnd As String
    Dim rsWorkTime As ADODB.Recordset
    On Error GoTo errHandle
    
    'ʱ���, ��ʼʱ��, ��ֹʱ��, ȱʡʱ��, ��ǰʱ��, ��ǰ��ɫ, nvl(վ��,'-') as վ��, nvl(����,'-') as ����, ���, ����Ԥ��ʱ��, ��Ϣʱ��
    Set rsWorkTime = GetWorkTimeRec
    '1.�����ޱ���Դ���õ�ʱ��
    If strվ�� <> "" And str���� <> "" Then
        rsWorkTime.Filter = "ʱ���='" & strʱ��� & "' And վ��='" & strվ�� & "' And ����='" & str���� & "'"
        If rsWorkTime.EOF Then
            rsWorkTime.Filter = "ʱ���='" & strʱ��� & "' And վ��='" & strվ�� & "' And ����='-'"
            If rsWorkTime.EOF Then
                rsWorkTime.Filter = "ʱ���='" & strʱ��� & "' And վ��='-' And ����='" & str���� & "'"
                If rsWorkTime.EOF Then rsWorkTime.Filter = "ʱ���='" & strʱ��� & "' And վ��='-' And ����='-'"
            End If
        End If
    ElseIf strվ�� <> "" And str���� = "" Then
        rsWorkTime.Filter = "ʱ���='" & strʱ��� & "' And վ��='" & strվ�� & "' And ����='-'"
        If rsWorkTime.EOF Then rsWorkTime.Filter = "ʱ���='" & strʱ��� & "' And վ��='-' And ����='-'"
    ElseIf strվ�� = "" And str���� <> "" Then
        rsWorkTime.Filter = "ʱ���='" & strʱ��� & "' And վ��='-' And ����='" & str���� & "'"
        If rsWorkTime.EOF Then rsWorkTime.Filter = "ʱ���='" & strʱ��� & "' And վ��='-' And ����='-'"
    Else
        rsWorkTime.Filter = rsWorkTime.Filter = "ʱ���='" & strʱ��� & "' And վ��='-' And ����='-'"
    End If
    '����վ�������
    If Not rsWorkTime.EOF Then
        Set obj�ϰ�ʱ�� = New �ϰ�ʱ��
        With obj�ϰ�ʱ��
            .ʱ��� = strʱ���
            .����Ԥ��ʱ�� = Val(Nvl(rsWorkTime!����Ԥ��ʱ��))
            .��ʼʱ�� = Format(rsWorkTime!��ʼʱ��, "yyyy-mm-dd HH:MM:SS")
            .����ʱ�� = Format(rsWorkTime!��ֹʱ��, "yyyy-mm-dd HH:MM:SS")
            .ȱʡԤԼʱ�� = Format(rsWorkTime!ȱʡʱ��, "yyyy-mm-dd HH:MM:SS")
            .��ǰ�Һ�ʱ�� = Nvl(rsWorkTime!��ǰʱ��)
            .��Ϣʱ�� = Nvl(rsWorkTime!��Ϣʱ��)
        End With
        Set GetWorkTimeRange = obj�ϰ�ʱ��
        Exit Function
    End If
   Set GetWorkTimeRange = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 
End Function

Public Function GetClinicRecordFromSignalSource(ByVal lng��ԴId As Long) As �����¼��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݺ�ԴID��ȡ�����¼��
    '���:lng��ԴID-��ԴID
    '����:���س����¼��
    '����:���˺�
    '����:2016-03-22 17:49:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj�����¼�� As �����¼��, obj�����¼ As �����¼
    Dim obj�������Ҽ� As �������Ҽ�, obj�������� As ��������
    Dim obj������Ϣ�� As ������Ϣ��, obj������Ϣ As ������Ϣ
    Dim obj������λ���Ƽ� As ������λ���Ƽ�, obj������λ���� As ������λ����
    Dim obj�ϰ�ʱ��  As �ϰ�ʱ��
    Dim rsControl As ADODB.Recordset, rsWorkTime As ADODB.Recordset, rsNum As ADODB.Recordset, rsUnitControl As ADODB.Recordset
    Dim dtDate As Date, strTemp As String, strSQL As String
    Dim rsRoom As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set obj�����¼�� = New �����¼��
    Set obj�����¼ = New �����¼
    Set rsWorkTime = GetWorkTimeRec
    strSQL = " " & _
    "   Select a.Id,C.����,C.����Ƶ��,C.����ID,C.ҽ��ID,C.ҽ������," & vbNewLine & _
    "          a.�ϰ�ʱ��, a.�޺���, a.��Լ��, a.�Ƿ���ſ���, a.�Ƿ��ʱ��, a.ԤԼ����," & vbNewLine & _
    "          a.�Ƿ��ռ, a.���﷽ʽ, a.����id, b.���� As ��������, d.վ�� " & vbNewLine & _
    "   From �ٴ������Դ C, �ٴ������Դ���� A, �������� B, ���ű� D" & vbNewLine & _
    "   Where c.Id = a.��Դid And a.����id = b.Id(+) And c.����ID = d.ID And c.Id = [1]"
    Set rsControl = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Դ����������Ϣ", lng��ԴId)

    strSQL = " " & _
    "   Select b.����id, b.����id, c.���� As �������� " & _
    "   From �ٴ������Դ���� A, �ٴ������Դ���� B, �������� C " & _
    "   Where a.��Դid = [1] And a.Id = b.����id And B.����id = c.Id"
    Set rsRoom = zlDatabase.OpenSQLRecord(strSQL, "���ٴ������Դ������Ϣ", lng��ԴId)

    strSQL = " " & _
    "   Select b.����id, b.���, b.��ʼʱ��, b.��ֹʱ��, b.��������, b.�Ƿ�ԤԼ " & _
    "   From �ٴ������Դ���� A, �ٴ������Դʱ�� B " & _
    "   Where a.��Դid = [1] And a.Id = b.����id"
    Set rsNum = zlDatabase.OpenSQLRecord(strSQL, "���ٴ������Դʱ����Ϣ", lng��ԴId)
    

    strSQL = "" & _
    "Select  b.����id, b.����, b.����, b.����, b.���, b.���Ʒ�ʽ, b.����, c.��ʼʱ��, c.��ֹʱ�� " & _
    "   From �ٴ������Դ���� A, �ٴ������Դ���� B, �ٴ�����ʱ�� C" & _
    "   Where a.��Դid = [1] And a.Id = b.����id And b.����ID = c.����ID(+) And b.��� = c.���(+)"
 
    Set rsUnitControl = zlDatabase.OpenSQLRecord(strSQL, "���ٴ������Դ������Ϣ", lng��ԴId)
    
    dtDate = zlDatabase.Currentdate
    With rsControl
        Do While Not .EOF
            If Nvl(rsControl!�ϰ�ʱ��) <> "" Then
                Set obj�ϰ�ʱ�� = GetWorkTimeRange(Nvl(rsControl!�ϰ�ʱ��), Nvl(rsControl!վ��), Nvl(rsControl!����))
                Set obj�����¼ = New �����¼
                rsWorkTime.Filter = "ʱ���='" & Nvl(rsControl!�ϰ�ʱ��) & "'"
                If Not rsWorkTime.EOF Then
                    obj�����¼.�������� = Format(obj�ϰ�ʱ��.��ʼʱ��, "yyyy-mm-dd")
                    obj�����¼.���﷽ʽ = Val(Nvl(rsControl!���﷽ʽ))
                    obj�����¼.��¼ID = Val(Nvl(rsControl!id))
                   Set obj�����¼.�ϰ�ʱ�� = obj�ϰ�ʱ��
                    obj�����¼.��ʼʱ�� = CDate(Format(dtDate, "yyyy-mm-dd") & " " & Format(rsWorkTime!��ʼʱ��, "HH:MM:SS"))
                    If Format(rsWorkTime!��ʼʱ��, "yyyy-mm-dd HH:MM:SS") >= Format(rsWorkTime!��ֹʱ��, "yyyy-mm-dd HH:MM:SS") Then
                        obj�����¼.��ֹʱ�� = CDate(Format(dtDate + 1, "yyyy-mm-dd") & " " & Format(rsWorkTime!��ֹʱ��, "HH:MM:SS"))
                    Else
                        obj�����¼.��ֹʱ�� = CDate(Format(dtDate, "yyyy-mm-dd") & " " & Format(rsWorkTime!��ֹʱ��, "HH:MM:SS"))
                    End If
                    obj�����¼.ʱ��� = Nvl(rsControl!�ϰ�ʱ��)
                    obj�����¼.�Ƿ��ʱ�� = Val(Nvl(rsControl!�Ƿ��ʱ��)) = 1
                    obj�����¼.�Ƿ��ռ = Val(Nvl(rsControl!�Ƿ��ռ)) = 1
                    obj�����¼.�Ƿ���ſ��� = Val(Nvl(rsControl!�Ƿ���ſ���)) = 1
                    obj�����¼.����ҽ�� = ""
                    obj�����¼.����ID = Val(Nvl(rsControl!����ID))
                    obj�����¼.ҽ��ID = Val(Nvl(rsControl!ҽ��ID))
                    obj�����¼.ҽ������ = Nvl(rsControl!ҽ������)
                    obj�����¼.�޺��� = Val(Nvl(rsControl!�޺���))
                    obj�����¼.��Լ�� = Val(Nvl(rsControl!��Լ��))
                    obj�����¼.�ѹ��� = 0
                    obj�����¼.��Լ�� = 0
                    obj�����¼.ԤԼ���� = Val(Nvl(rsControl!ԤԼ����))
                    Set obj�������Ҽ� = New �������Ҽ�
                    obj�������Ҽ�.���﷽ʽ = Val(Nvl(rsControl!���﷽ʽ))
                    obj�������Ҽ�.ҽ������ = Nvl(rsControl!ҽ������)
                    '1.��������
                    rsRoom.Filter = "����ID=" & Val(Nvl(rsControl!id))
                    If rsRoom.RecordCount <> 0 Then rsRoom.MoveFirst
                    Do While Not rsRoom.EOF
                        Set obj�������� = New ��������
                        obj��������.����ID = Val(Nvl(rsRoom!����ID))
                        obj��������.�������� = Nvl(rsRoom!��������)
                        
                        obj�������Ҽ�.AddItem obj��������, "K" & obj��������.����ID
                        
                        rsRoom.MoveNext
                    Loop
                   Set obj�����¼.�����������Ҽ� = obj�������Ҽ�
                   
                   '2.���غ�����Ϣ��
                    Set obj������Ϣ�� = New ������Ϣ��
                    rsNum.Filter = "����ID=" & Val(Nvl(rsControl!id))
                    If rsNum.RecordCount <> 0 Then rsNum.MoveFirst
                    Do While Not rsNum.EOF
                        Set obj������Ϣ = New ������Ϣ
                        obj������Ϣ.��� = Val(Nvl(rsNum!���))
                        obj������Ϣ.��ʼʱ�� = Format(rsNum!��ʼʱ��, "yyyy-mm-dd HH:MM:SS")
                        obj������Ϣ.��ֹʱ�� = Format(rsNum!��ֹʱ��, "yyyy-mm-dd HH:MM:SS")
                        obj������Ϣ.�Ƿ�ԤԼ = Val(Nvl(rsNum!�Ƿ�ԤԼ)) = 1
                        obj������Ϣ.���� = Val(Nvl(rsNum!��������))
                       
                        obj������Ϣ��.AddItem obj������Ϣ
                        rsNum.MoveNext
                    Loop
                    'Set obj������Ϣ��.�ϰ�ʱ�� = obj�ϰ�ʱ��
                    obj������Ϣ��.����Ƶ�� = Val(Nvl(rsControl!����Ƶ��))
                    obj������Ϣ��.ʱ��� = obj�����¼.ʱ���
                    obj������Ϣ��.�Ƿ��ʱ�� = obj�����¼.�Ƿ��ʱ��
                    obj������Ϣ��.�Ƿ���ſ��� = obj�����¼.�Ƿ���ſ���
                    obj������Ϣ��.�޺��� = obj�����¼.�޺���
                    obj������Ϣ��.��Լ�� = obj�����¼.��Լ��
                    obj������Ϣ��.ԤԼ���� = obj�����¼.ԤԼ����
                    
                   Set obj�����¼.������Ϣ�� = obj������Ϣ��
                   '3.����������λ����
                    Set obj������λ���Ƽ� = New ������λ���Ƽ�
                    obj������λ���Ƽ�.�Ƿ��ռ = Val(Nvl(rsControl!�Ƿ��ռ))
                    Set obj������Ϣ�� = Nothing
                    strTemp = ""
                    
                    
                    rsUnitControl.Filter = "����ID=" & Val(Nvl(rsControl!id))
                    rsUnitControl.Sort = "����,����,����,���"
                    If rsUnitControl.RecordCount <> 0 Then
                        rsUnitControl.MoveFirst
                        obj������λ���Ƽ�.ԤԼ���Ʒ�ʽ = Val(Nvl(rsUnitControl!���Ʒ�ʽ))
                    End If
                    Do While Not rsUnitControl.EOF
                        If strTemp <> Nvl(rsUnitControl!����) & "-" & Nvl(rsUnitControl!����) & "-" & Nvl(rsUnitControl!����) Then
                            If Not obj������Ϣ�� Is Nothing Then
                                Set obj������λ����.������Ϣ�� = obj������Ϣ��
                                obj������λ���Ƽ�.AddItem obj������λ����, "K" & obj������λ����.������λ����
                            End If
                            Set obj������λ���� = New ������λ����
                            obj������λ����.������λ���� = Nvl(rsUnitControl!����)
                            obj������λ����.���� = Val(Nvl(rsUnitControl!����))
                            obj������λ����.ԤԼ���Ʒ�ʽ = Val(Nvl(rsUnitControl!���Ʒ�ʽ))
                            Set obj������Ϣ�� = New ������Ϣ��
                            
                            strTemp = Nvl(rsUnitControl!����) & "-" & Nvl(rsUnitControl!����) & "-" & Nvl(rsUnitControl!����)
                        End If
                        Set obj������Ϣ = New ������Ϣ
                        obj������Ϣ.��� = Val(Nvl(rsUnitControl!���))
                        
                        obj������Ϣ.��ʼʱ�� = Format(Nvl(rsUnitControl!��ʼʱ��), "yyyy-mm-dd HH:MM:SS")
                        obj������Ϣ.��ֹʱ�� = Format(Nvl(rsUnitControl!��ֹʱ��), "yyyy-mm-dd HH:MM:SS")
                        obj������Ϣ.���� = Val(Nvl(rsUnitControl!����))
                        obj������Ϣ.�Ƿ�ԤԼ = 1 '����ԤԼ��
                        obj������Ϣ��.AddItem obj������Ϣ
                        rsUnitControl.MoveNext
                    Loop
                    If Not obj������Ϣ�� Is Nothing Then
                        Set obj������λ����.������Ϣ�� = obj������Ϣ��
                        obj������λ���Ƽ�.AddItem obj������λ����, "K" & obj������λ����.������λ����
                        
                    End If
                   
                    Set obj�����¼.������λ���Ƽ� = obj������λ���Ƽ�
                   obj�����¼��.AddItem obj�����¼, "K" & obj�����¼.ʱ���
                End If
            End If
            .MoveNext
        Loop
        obj�����¼��.�������� = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    End With
    
    Set GetClinicRecordFromSignalSource = obj�����¼��
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function GetWorkTimeRec() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ϰ�ʱ��εļ�¼��
    '���:
    '����:�ϰ�ʱ��μ�¼��
    '����:���˺�
    '����:2016-03-22 16:18:06
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    
    strSQL = "Select ʱ���, To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'HH24:MI:SS') As ��ʼʱ��," & vbNewLine & _
            "        To_Char(Sysdate + Case When To_Char(��ʼʱ��, 'hh24:mi:ss') >= To_Char(��ֹʱ��, 'hh24:mi:ss') Then 1 Else 0 End, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'HH24:MI:SS') As ��ֹʱ��," & vbNewLine & _
            "        To_Char(Sysdate + Case When To_Char(��ʼʱ��, 'hh24:mi:ss') > To_Char(Nvl(ȱʡʱ��,��ʼʱ��), 'hh24:mi:ss') Then  1 Else 0 End, 'yyyy-mm-dd') || ' ' || To_Char(Nvl(ȱʡʱ��,��ʼʱ��), 'HH24:MI:SS') As ȱʡʱ��," & vbNewLine & _
            "        To_Char(Sysdate + Case When To_Char(��ʼʱ��, 'hh24:mi:ss') < To_Char(Nvl(��ǰʱ��,��ʼʱ��), 'hh24:mi:ss') Then -1 Else 0 End, 'yyyy-mm-dd') || ' ' || To_Char(Nvl(��ǰʱ��,��ʼʱ��), 'HH24:MI:SS') As ��ǰʱ��," & vbNewLine & _
            "        Nvl(վ��, '-') As վ��, Nvl(����, '-') As ����, ����Ԥ��ʱ��, ��Ϣʱ��" & vbNewLine & _
            " From ʱ���"

    If grsWorkTime Is Nothing Then
        Set grsWorkTime = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ϰ�ʱ��")
    ElseIf grsWorkTime.State <> adStateOpen Then
        Set grsWorkTime = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ϰ�ʱ��")
    End If
    
    Set GetWorkTimeRec = grsWorkTime
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetUnitAll() As ADODB.Recordset
    '���ܣ���ȡ���йҺź�����λ��ԤԼ��ʽ
    '��Σ�
    '   strStationNo:վ����
    '   strSignalType:����
    Dim strSQL As String
    
    On Error GoTo errHandler
    strSQL = "Select 1 As ����, ���� From �Һź�����λ" & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select 2 As ����, ���� From ԤԼ��ʽ"
    If grsUnit Is Nothing Then
        Set grsUnit = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���йҺź�����λ��ԤԼ��ʽ")
    ElseIf grsUnit.State <> adStateOpen Then
        Set grsUnit = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���йҺź�����λ��ԤԼ��ʽ")
    End If
    grsUnit.MoveFirst
    Set GetUnitAll = grsUnit
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDoctorRooms(ByVal lng����ID As Long) As ADODB.Recordset
    '���ܣ��������ÿ���ID��ȡ��������
    '��Σ�
    '   lng����ID:����ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandler
    strSQL = "Select a.Id As ����ID, a.����, a.����, b.ȱʡ��־" & vbNewLine & _
            " From �������� A, �����������ÿ��� B" & vbNewLine & _
            " Where a.Id = b.����id And b.����id = [1]" & vbNewLine & _
            "       And (a.վ�� Is Null Or a.վ��=(Select վ�� From ���ű� Where id = [1]))" & vbNewLine & _
            " Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������", lng����ID)
    
    If rsTemp.RecordCount = 0 Then
        strSQL = "Select a.Id As ����ID, a.����, a.����, 0 As ȱʡ��־" & vbNewLine & _
                " From �������� A" & vbNewLine & _
                " Where a.վ�� Is Null Or a.վ��=(Select վ�� From ���ű� Where id = [1])" & vbNewLine & _
                " Order By ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������", lng����ID)
    End If
    
    Set GetDoctorRooms = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSignalSource(Optional ByVal str�Ű෽ʽ As String, _
    Optional ByVal lng��ԴId As Long) As ADODB.Recordset
    '���ܣ���ȡ�ٴ������Դ
    '��Σ�
    '   str�Ű෽ʽ:0-�̶��Ű�;1-�����Ű�;2-�����Ű� ����ö��ŷָ�
    '   lng��ԴID:�ٴ������Դ.ID
    Dim strSQL As String, strWhere As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandler
    If str�Ű෽ʽ <> "" Then strWhere = " And Instr([1], a.�Ű෽ʽ) > 0"
    If lng��ԴId <> 0 Then strWhere = " And a.ID=[2]"
    strSQL = "Select a.Id, a.����, a.����, a.����id, b.���� As ��������, a.��Ŀid, c.���� As ��Ŀ����, a.ҽ��id, a.ҽ������," & vbNewLine & _
            "        d.רҵ����ְ�� As ҽ��ְ��, a.�Ƿ񽨲���, a.ԤԼ����, a.����Ƶ��, a.���տ���״̬, a.�Ƿ���ջ���," & vbNewLine & _
            "        a.�Ƿ��ٴ��Ű�, a.�Ű෽ʽ, a.�Ƿ�ɾ��, a.����ʱ��, a.����ʱ��, b.վ��" & vbNewLine & _
            " From �ٴ������Դ A, ���ű� B, �շ���ĿĿ¼ C, ��Ա�� D" & vbNewLine & _
            " Where a.����id = b.Id And a.��Ŀid = c.Id And a.ҽ��ID = d.ID(+)" & strWhere & vbNewLine & _
            "       And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & vbNewLine & _
            "       And Nvl(Nvl(b.վ��, [4]), Nvl([3], '-')) = Nvl([3], '-')" & vbNewLine & _
            " Order By a.���� asc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ٴ������Դ", str�Ű෽ʽ, lng��ԴId, gstrNodeNo, gVisitPlan_ModulePara.str��Դά��վ��)
    
    Set GetSignalSource = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetNextPlanDate(frmParent As Object, ByVal byt�Ű෽ʽ As Byte, _
    ByRef intYear As Integer, ByRef intMonth As Integer, Optional ByRef intWeek As Integer, _
    Optional ByVal lng��ԱID As Long, Optional ByVal blnShowSelect As Boolean = True) As Boolean
    '���ܣ�ȷ����һ���Ű��������
    '������
    '   byt�Ű෽ʽ��1-�����Ű�;2-�����Ű�
    '   lng��ԱID����"���п���"Ȩ��ʱ����
    '���أ����飬0-��ݣ�1-�·ݣ�2-����
    '˵��������ԱIDȷ����Դʱ��ͨ����ӵ�к�Դ�������һ���������ȷ��ʱ��
    Dim strSQL As String, rsData As ADODB.Recordset
    Dim dtStart As Date, dtEnd As Date, dtCur As Date
    Dim blnFind As Boolean, strWhere As String
    Dim intYearTemp As Integer, intMonthTemp As Integer, intWeekTemp As Integer
    
    Err = 0: On Error GoTo errHandler
    intYear = 0: intMonth = 0: intWeek = 0
'    If lng��ԱID <> 0 Then
'        strWhere = "Exists" & vbNewLine & _
'                " (Select 1" & vbNewLine & _
'                "       From �ٴ����ﰲ�� M, �ٴ������Դ N, ���ű� P" & vbNewLine & _
'                "       Where m.����id = a.Id And m.��Դid = n.Id And n.����id = p.Id" & vbNewLine & _
'                "             And Nvl(n.�Ƿ��ٴ��Ű�, 0) = 1 And Nvl(n.�Ű෽ʽ, 0) = [1]" & vbNewLine & _
'                "             And Nvl(n.�Ƿ�ɾ��, 0) = 0 And (n.����ʱ�� Is Null Or n.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
'                "             And (p.վ�� = '" & gstrNodeNo & "' Or p.վ�� Is Null)" & vbNewLine & _
'                "             And Exists (Select 1 From ������Ա Where ����id = n.����id And ��Աid = [2]))"
'        '����Ҫ�������һ���ѷ����ĳ����
'        strWhere = " And (" & strWhere & " Or a.����ʱ�� Is Not Null)"
'    End If
    '�Ű෽ʽ��0-�̶��Ű�;1-�����Ű�;2-�����Ű�;3-ģ��
    strSQL = "Select a.���, a.�·�, a.����" & vbNewLine & _
            " From �ٴ������ A" & vbNewLine & _
            " Where a.�Ű෽ʽ=[1] " & strWhere & vbNewLine & _
            "       And Nvl(a.վ��,'-') = Nvl([3],'-')" & vbNewLine & _
            " Order By a.��� Desc, a.�·� Desc, a.���� Desc"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ڵ�ǰʱ���������", byt�Ű෽ʽ, lng��ԱID, gstrNodeNo)
    
    dtStart = zlDatabase.Currentdate()
    If rsData.EOF Then
        '���ݿ����޳�������û�ȷ������������
        GoTo SetNewDate:
    Else
        intYear = Val(Nvl(rsData!���))
        intMonth = Val(Nvl(rsData!�·�))
        If byt�Ű෽ʽ = 2 Then intWeek = Val(Nvl(rsData!����))
    End If
        
    If byt�Ű෽ʽ = 1 Then
        If intMonth = 12 Then '��ǰΪ��������һ����,������Ϊ��һ���1��
            intYear = intYear + 1: intMonth = 1
        Else
            intMonth = intMonth + 1
        End If
    Else
        If GetWeekCount(intYear, intMonth) = intWeek Then '��ǰΪ���µ����һ��,������Ϊ���µĵ�һ��
            If intMonth = 12 Then '��ǰΪ��������һ����,������Ϊ��һ���1��
                intYear = intYear + 1: intMonth = 1
            Else
                intMonth = intMonth + 1
            End If
            intWeek = 1
        Else
            intWeek = intWeek + 1
        End If
    End If
    
    'С�ڵ�ǰʱ��ʱ�����û�ȷ������������
    intYearTemp = Year(dtStart): intMonthTemp = Month(dtStart): intWeekTemp = GetDateWeek(dtStart)
    If intYear < intYearTemp Then
        GoTo SetNewDate:
    ElseIf intYear = intYearTemp Then
        If intMonth < intMonthTemp Then
            GoTo SetNewDate:
        ElseIf intMonth = intMonthTemp Then
            If intWeek < intWeekTemp And byt�Ű෽ʽ = 2 Then
                GoTo SetNewDate:
            End If
        End If
    End If
    GetNextPlanDate = True
    Exit Function
    
SetNewDate:
    intYear = 0: intMonth = 0: intWeek = 0
    '���û�ȷ������������
    If blnShowSelect Then
        Dim frm As New frmClinicSetNewPlanDate
        If frm.ShowMe(frmParent, byt�Ű෽ʽ, dtStart, intYear, intMonth, intWeek) = False Then Exit Function
    End If
    GetNextPlanDate = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetWorkTimes(Optional ByVal strStationNo As String, _
    Optional ByVal strSignalType As String) As ADODB.Recordset
    '���ܣ���ȡ�ϰ�ʱ��
    '��Σ�
    '   strStationNo:վ����
    '   strSignalType:����
    Dim rsTemp As New ADODB.Recordset, strFilter As String
    Dim strʱ��� As String, lngCount As Long

    On Error GoTo errHandler
    '��������
    strFilter = "(վ��='-' And ����='-')"
    If strStationNo <> "" Then
        strFilter = strFilter & " OR (վ��='" & strStationNo & "' And ����='-')"
    End If
    If strSignalType <> "" Then
        strFilter = strFilter & " OR (վ��='-' And ����='" & strSignalType & "')"
    End If
    If strStationNo <> "" And strSignalType <> "" Then
        strFilter = strFilter & " OR (վ��='" & strStationNo & "' And ����='" & strSignalType & "')"
    End If
    
    Set rsTemp = GetWorkTimeRec() 'ȡ��������
    rsTemp.Filter = strFilter

    Set GetWorkTimes = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitPlan(ByVal lng����ID As Long, Optional ByVal lng����ID As Long) As ADODB.Recordset
    '���ܣ���ȡ�ٴ����ﰲ��
    '��Σ�
    '   lng����ID:����ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If lng����ID = 0 Then
        strSQL = "Select b.Id As ����id, b.�Ű෽ʽ, b.�������, b.���, b.�·�, b.����, b.Ӧ�÷�Χ, b.����id, b.��ע, b.������, b.����ʱ��," & vbNewLine & _
                "        b.ģ������,Null As ����id, Null As ��Դid, Null As ��Ŀid, Null As ��Ŀ����, Null As ҽ��id," & _
                "        Null As ҽ������, Null As ҽ��ְ��, Null As �Ű����, Null As �Ƿ���������, Null As �Ƿ����ճ���," & vbNewLine & _
                "        null As ��ʼʱ��, null As ��ֹʱ��, Null As ����Ա����, Null As �Ǽ�ʱ��, Null As �Ƿ���ʱ����" & vbNewLine & _
                " From �ٴ������ B" & vbNewLine & _
                " Where b.Id = [2] And Rownum < 2"
    Else
        strSQL = "Select b.Id As ����id, b.�Ű෽ʽ, b.�������, b.���, b.�·�, b.����," & vbNewLine & _
                "        b.Ӧ�÷�Χ, b.����id, b.��ע, b.������, b.����ʱ��,b.ģ������," & vbNewLine & _
                "        a.Id As ����id, a.��Դid, a.��Ŀid, c.���� As ��Ŀ����, a.ҽ��id, a.ҽ������,d.רҵ����ְ�� As ҽ��ְ��," & vbNewLine & _
                "        a.�Ű����, a.�Ƿ���������, a.�Ƿ����ճ���, a.��ʼʱ��, a.��ֹʱ��, a.����Ա����, a.�Ǽ�ʱ��, a.�Ƿ���ʱ����" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ������ B, �շ���ĿĿ¼ C, ��Ա�� D" & vbNewLine & _
                " Where a.����id = b.Id And a.��ĿID = c.ID And a.ҽ��ID = d.ID(+) And a.Id = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������", lng����ID, lng����ID)
    Set GetVisitPlan = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitItems(ByVal lng����ID As Long, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '���ܣ���ȡ�ٴ��������ƻ��¼��Ŀ
    '��Σ�
    '   lng����ID:����ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        strSQL = "Select Distinct To_Char(b.��������,'yyyy-mm-dd') As ��������" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ������¼ B" & vbNewLine & _
                " Where a.Id = b.����id And a.Id = [1] And b.�ϰ�ʱ�� Is Not Null"
    Else
        strSQL = "Select Distinct b.������Ŀ As ��������" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ��������� B" & vbNewLine & _
                " Where a.Id = b.����id And a.Id = [1] And b.�ϰ�ʱ�� Is Not Null"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ŀ", lng����ID)
    Set GetVisitItems = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitTimes(ByVal lng����ID As Long, Optional ByVal str��Ŀ As String, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '���ܣ���ȡ�ٴ�����ʱ��
    '��Σ�
    '   lng����ID:����ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        strSQL = "Select b.ID As ��¼ID,To_Char(b.��������,'yyyy-mm-dd') As ��������, b.�ϰ�ʱ��, b.�Ƿ��ʱ��, b.�Ƿ���ſ���, b.��ʼʱ��, b.��ֹʱ��," & vbNewLine & _
                "        b.�޺���, b.�ѹ���, b.��Լ��, b.��Լ��, b.���﷽ʽ, b.ԤԼ����, b.����ҽ������, b.����ID" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ������¼ B" & vbNewLine & _
                " Where a.Id = b.����id And a.Id = [1] And b.�ϰ�ʱ�� Is Not Null"
        If str��Ŀ <> "" Then
            strSQL = strSQL & " And b.��������=to_date([2],'yyyy-mm-dd')"
        End If
    Else
        strSQL = "Select b.ID as ��¼ID,b.������Ŀ As ��������, b.�ϰ�ʱ��, b.�Ƿ��ʱ��, b.�Ƿ���ſ���, NULL as ��ʼʱ��, NULL as ��ֹʱ��," & vbNewLine & _
                "        b.�޺���, 0 as �ѹ���, b.��Լ��, 0 as ��Լ��, b.���﷽ʽ, b.ԤԼ����, '' As ����ҽ������,0 As ����ID" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ��������� B" & vbNewLine & _
                " Where a.Id = b.����id(+) And a.Id = [1] And b.�ϰ�ʱ�� Is Not Null"
        If str��Ŀ <> "" Then
            strSQL = strSQL & " And b.������Ŀ=[2]"
        End If
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", lng����ID, str��Ŀ)
    Set GetVisitTimes = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPreviousVisitTimes(ByVal lng��ԴId As Long, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '���ܣ���ȡ�ϴ���Ч�ٴ�����ʱ��
    '��Σ�
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        strSQL = "Select ��¼id, ��������, �ϰ�ʱ��, �Ƿ��ʱ��, �Ƿ���ſ���, ��ʼʱ��, ��ֹʱ��, �޺���," & vbNewLine & _
                "        �ѹ���, ��Լ��, ��Լ��, ���﷽ʽ, ԤԼ����, ����ҽ������, ����ID" & vbNewLine & _
                " From (Select b.Id As ��¼id, b.��������, b.�ϰ�ʱ��, b.�Ƿ��ʱ��, b.�Ƿ���ſ���," & vbNewLine & _
                "              b.��ʼʱ��, b.��ֹʱ��, b.�޺���, b.�ѹ���, b.��Լ��, b.��Լ��, b.���﷽ʽ, b.ԤԼ����," & vbNewLine & _
                "              b.����ҽ������, b.����ID, Row_Number() Over(Partition By b.�ϰ�ʱ�� Order By b.Id Desc) As ���" & vbNewLine & _
                "        From �ٴ����ﰲ�� A, �ٴ������¼ B, �ٴ������ C" & vbNewLine & _
                "        Where a.Id = b.����id And a.����id = c.Id And a.��ԴID=[1] And c.����ʱ�� Is Not Null)" & vbNewLine & _
                " Where ��� = 1"
    Else
        strSQL = "Select ��¼id, ��������, �ϰ�ʱ��, �Ƿ��ʱ��, �Ƿ���ſ���, ��ʼʱ��, ��ֹʱ��, �޺���," & vbNewLine & _
                "        �ѹ���, ��Լ��, ��Լ��, ���﷽ʽ, ԤԼ����, ����ҽ������, ����ID" & vbNewLine & _
                " From (Select b.Id As ��¼id, b.������Ŀ As ��������, b.�ϰ�ʱ��, b.�Ƿ��ʱ��, b.�Ƿ���ſ���," & vbNewLine & _
                "              NULL as ��ʼʱ��, NULL as ��ֹʱ��, b.�޺���, 0 As �ѹ���, b.��Լ��, 0 As ��Լ��, b.���﷽ʽ, b.ԤԼ����," & vbNewLine & _
                "              '' As ����ҽ������, 0 As ����ID, Row_Number() Over(Partition By b.�ϰ�ʱ�� Order By b.Id Desc) As ���" & vbNewLine & _
                "        From �ٴ����ﰲ�� A, �ٴ��������� B, �ٴ������ C" & vbNewLine & _
                "        Where a.Id = b.����id And a.����id = c.Id And a.��ԴID=[1] And c.����ʱ�� Is Not Null)" & vbNewLine & _
                " Where ��� = 1"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", lng��ԴId)
    Set GetPreviousVisitTimes = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitTime(ByVal lng��¼ID As Long, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '���ܣ���ȡ�ٴ����ﵥ��ʱ��
    '��Σ�
    '   lng��¼ID:��¼ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        strSQL = "Select b.ID As ��¼ID,b.��������, b.�ϰ�ʱ��, b.�Ƿ��ʱ��, b.�Ƿ���ſ���, b.��ʼʱ��, b.��ֹʱ��," & vbNewLine & _
                "        b.�޺���, b.�ѹ���, b.��Լ��, b.��Լ��, b.���﷽ʽ, b.ԤԼ����, b.����ҽ������, " & vbNewLine & _
                "        b.����ID, b.��ĿID, c.���� As ��Ŀ����, b.ҽ��ID, b.ҽ������, b.�Ƿ��ռ, " & vbNewLine & _
                "        b.ͣ�￪ʼʱ��, b.ͣ����ֹʱ��, b.ͣ��ԭ��, b.�Ƿ���ʱ����" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ������¼ B, �շ���ĿĿ¼ C" & vbNewLine & _
                " Where a.Id = b.����id And b.��ĿID = c.ID And b.Id = [1] And b.�ϰ�ʱ�� Is Not Null"
    Else
        strSQL = "Select b.ID as ��¼ID,b.������Ŀ As ��������, b.�ϰ�ʱ��, b.�Ƿ��ʱ��, b.�Ƿ���ſ���, NULL as ��ʼʱ��, NULL as ��ֹʱ��," & vbNewLine & _
                "        b.�޺���, 0 as �ѹ���, b.��Լ��, 0 as ��Լ��, b.���﷽ʽ, b.ԤԼ����, NULL As ����ҽ������," & vbNewLine & _
                "        0 As ����ID, 0 As ��ĿID, '' As ��Ŀ����, 0 As ҽ��ID, '' As ҽ������, 0 As �Ƿ��ռ, " & vbNewLine & _
                "        NULL As ͣ�￪ʼʱ��, NULL As ͣ����ֹʱ��, NULL As ͣ��ԭ��, 0 As �Ƿ���ʱ����" & vbNewLine & _
                " From �ٴ����ﰲ�� A, �ٴ��������� B" & vbNewLine & _
                " Where a.Id = b.����id(+) And b.Id = [1] And b.�ϰ�ʱ�� Is Not Null"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", lng��¼ID)
    Set GetVisitTime = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitRooms(ByVal lngID As Long, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '���ܣ���ȡ�ٴ���������
    '��Σ�
    '   lngID:��¼ID/����ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandler
    If blnRecord Then
        strSQL = "Select a.����ID, b.����" & vbNewLine & _
                " From �ٴ��������Ҽ�¼ A, �������� B" & vbNewLine & _
                " Where a.����id = b.Id And a.��¼ID = [1]"
    Else
        strSQL = "Select a.����ID, b.����" & vbNewLine & _
                " From �ٴ��������� A, �������� B" & vbNewLine & _
                " Where a.����id = b.Id And a.����id = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ٴ���������", lngID)
    Set GetVisitRooms = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetTimeInterval(ByVal lngID As Long, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '���ܣ���ȡ������Ϣ
    '��Σ�
    '   lngID:��¼ID/����ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        '"       And (a.��ʼʱ�� <> a. ��ֹʱ�� Or a.��ʼʱ�� Is Null And a. ��ֹʱ�� Is Null)" & vbNewLine & _'��ʼʱ������ֹʱ����ȵ��ǼӺŵ����
        strSQL = "Select a.���, a.��ʼʱ��, a. ��ֹʱ��, a.����, a.�Ƿ�ԤԼ, a.�Ƿ�ͣ��" & vbNewLine & _
                " From �ٴ�������ſ��� A,�ٴ������¼ B" & vbNewLine & _
                " Where a.��¼ID =b.ID And b.ID=[1] " & vbNewLine & _
                "       And (a.��ʼʱ�� <> a. ��ֹʱ�� Or a.��ʼʱ�� Is Null And a. ��ֹʱ�� Is Null)" & vbNewLine & _
                "       And (Not(Nvl(b.�Ƿ��ʱ��,0)=1 And Nvl(b.�Ƿ���ſ���,0)=0) Or Nvl(b.�Ƿ��ʱ��,0)=1 And Nvl(b.�Ƿ���ſ���,0)=0 And a.ԤԼ˳��� IS NULL)"
    Else
        strSQL = "Select a.���, a.��ʼʱ��, a. ��ֹʱ��, a.��������  As ����, a.�Ƿ�ԤԼ, 0 As �Ƿ�ͣ��" & vbNewLine & _
                " From �ٴ�����ʱ�� A" & vbNewLine & _
                " Where a.����ID = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lngID)
    Set GetTimeInterval = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetUnitReg(ByVal lngID As Long, ByVal str������λ As String, ByVal byt���� As Byte, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '���ܣ���ȡ������λ������Ϣ
    '��Σ�
    '   lngID:��¼ID/����ID
    '   str������λ:������λ����
    '   byt����:���ͣ�0-������λ��1-ԤԼ��ʽ
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        '"       And (b.��ʼʱ�� <> b. ��ֹʱ�� Or b.��ʼʱ�� Is Null And b. ��ֹʱ�� Is Null)" & vbNewLine & _'��ʼʱ������ֹʱ����ȵ��ǼӺŵ����
        strSQL = "Select a.���Ʒ�ʽ, a.���, b.��ʼʱ��, b.��ֹʱ��, a.����, b.�Ƿ�ԤԼ, b.�Ƿ�ͣ��" & vbNewLine & _
                " From �ٴ�����Һſ��Ƽ�¼ A, �ٴ�������ſ��� B,�ٴ������¼ C" & vbNewLine & _
                " Where a.��¼id = b.��¼id(+) And a.��� = b.���(+)" & vbNewLine & _
                "       And (b.��ʼʱ�� <> b. ��ֹʱ�� Or b.��ʼʱ�� Is Null And b. ��ֹʱ�� Is Null)" & vbNewLine & _
                "       And a.��¼id = c.ID And c.ID=[1] And a.���� = [2] and Nvl(a.����,0) = [3]" & vbNewLine & _
                "       And (Not(Nvl(c.�Ƿ��ʱ��,0)=1 And Nvl(c.�Ƿ���ſ���,0)=0) Or Nvl(c.�Ƿ��ʱ��,0)=1 And Nvl(c.�Ƿ���ſ���,0)=0 And b.ԤԼ˳��� IS NULL)"
    Else
        strSQL = "Select a.���Ʒ�ʽ, a.���, b.��ʼʱ��, b.��ֹʱ��, a.����, b.�Ƿ�ԤԼ, 0 As �Ƿ�ͣ��" & vbNewLine & _
                " From �ٴ�����Һſ��� A, �ٴ�����ʱ�� B" & vbNewLine & _
                " Where a.����ID = b.����ID(+) And a.��� = b.���(+) " & vbNewLine & _
                "       And a.����ID = [1] And a.���� = [2] and Nvl(a.����,0) = [3]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������λ�Һſ���", lngID, str������λ, byt����)

    Set GetUnitReg = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetStopVisit(ByVal lng��ԴId As Long, ByVal dt��ʼʱ�� As Date, dt��ֹʱ�� As Date, _
    Optional ByVal blnAllHoliday As Boolean = True) As ADODB.Recordset
    '���ܣ���ȡͣ���¼���ڼ��ա�ͣ�ﰲ�ţ�
    '��Σ�
    '   lng��ԴID:��ԴID
    '   dt��ʼʱ�䡢dt��ֹʱ��:��ѯʱ�䷶Χ
    '   blnAllHoliday:���м��գ������ݡ����տ���״̬���ж�
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    'ͣ�ﰲ��
    strSQL = "Select 1 As ����, a.��ʼʱ��, a.��ֹʱ��, a.ͣ��ԭ��" & vbNewLine & _
            " From �ٴ�����ͣ���¼ A, �ٴ������Դ B" & vbNewLine & _
            " Where a.������ = b.ҽ������ And a.��¼id Is Null And a.����ʱ�� Is Not Null And a.ȡ���� Is Null" & vbNewLine & _
            "       And b.Id = [1] And b.ҽ��id Is Not Null And Not (a.��ʼʱ�� > [3] Or a.��ֹʱ�� < [2])" & vbNewLine
    '�ڼ���(ȫ�������ﲻ����"�ٴ������Դ.���տ���״̬"�жϣ��ڷ�������ʱ����)
    strSQL = strSQL & _
            " Union All" & vbNewLine & _
            " Select 2 As ����, ��ʼ����, ��ֹ����, ��������" & vbNewLine & _
            " From �������ձ�" & vbNewLine & _
            " Where ���� = 0" & vbNewLine & _
            "       And Not (��ʼ���� > [3] Or ��ֹ���� < [2])"
    If blnAllHoliday = False Then
        strSQL = strSQL & vbNewLine & _
                "   And Exists (Select 1 From �ٴ������Դ Where ID = [1] And Nvl(���տ���״̬, 0) = 0)"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͣ���¼", lng��ԴId, dt��ʼʱ��, dt��ֹʱ��)

    Set GetStopVisit = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get�����¼(ByVal lng��ԴId As Long, ByVal lng��¼ID As Long, ByVal blnRecord As Boolean, _
    ByRef obj�����Դ As �����Դ, ByRef obj�����¼ As �����¼) As Boolean
    Dim rsSignalSource As ADODB.Recordset
    Dim rsRecord As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim obj������λ As ������λ����, obj������λ���� As ������λ����
    Dim obj���з������Ҽ� As �������Ҽ�
    Dim obj�����ϰ�ʱ�μ� As �ϰ�ʱ�μ�, obj�����¼�� As �����¼��
    Dim obj���к�����λ As ������λ���Ƽ�
    
    Err = 0: On Error GoTo errHandler
    '��Դ��Ϣ,���ҡ�ҽ������Ŀȡ�����е�
    Set rsSignalSource = GetSignalSource("", lng��ԴId)
    If rsSignalSource.RecordCount = 0 Then
        MsgBox "δ���ֺ�Դ��Ϣ����ˢ�����ݺ����ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    Set obj�����Դ = GetSignalSourceObject(rsSignalSource)
    If obj�����Դ Is Nothing Then
        MsgBox "��ȡ��Դ��Ϣ���������ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
     Set obj�����ϰ�ʱ�μ� = GetWorkTimesObjects(GetWorkTimes(obj�����Դ.վ��, obj�����Դ.����))
    '�����¼
    Set obj�����¼�� = GetVisitTimesObjects(GetVisitTime(lng��¼ID, blnRecord))
    If obj�����¼��.Count = 0 Then
        MsgBox "��ȡ�����¼���������ԣ�", vbInformation, gstrSysName: Exit Function
    End If
    
    Set obj�����¼ = obj�����¼��(1)
    With obj�����¼
        '��Դ��Ϣ,���ҡ�ҽ������Ŀȡ�����е�
        obj�����Դ.��ĿID = .��ĿID
        obj�����Դ.��Ŀ���� = .��Ŀ����
        obj�����Դ.ҽ��ID = .ҽ��ID
        obj�����Դ.ҽ������ = .ҽ������
    
        If obj�����ϰ�ʱ�μ�.Exits("K" & .ʱ���) Then
            Set .�ϰ�ʱ�� = obj�����ϰ�ʱ�μ�("K" & .ʱ���).Clone
        Else
            Set .�ϰ�ʱ�� = New �ϰ�ʱ��
        End If
        
        '��������
        Set .�����������Ҽ� = GetVisitRoomsObjects(GetVisitRooms(.��¼ID, blnRecord))
        .�����������Ҽ�.���﷽ʽ = .���﷽ʽ
        .�����������Ҽ�.ҽ������ = obj�����Դ.ҽ������
        
        '������Ϣ
        Set .������Ϣ�� = GetTimeIntervalObjects(GetTimeInterval(.��¼ID, blnRecord))
        
        .������Ϣ��.����Ƶ�� = obj�����Դ.����Ƶ��
        .������Ϣ��.�Ƿ��ʱ�� = .�Ƿ��ʱ��
        .������Ϣ��.�Ƿ���ſ��� = .�Ƿ���ſ���
        .������Ϣ��.�޺��� = .�޺���
        .������Ϣ��.��Լ�� = .��Լ��
        .������Ϣ��.ԤԼ���� = .ԤԼ����
        .������Ϣ��.ʱ��� = .ʱ���
        
        '������λ����
        Set .������λ���Ƽ� = New ������λ���Ƽ�
        .������λ���Ƽ�.�Ƿ��ռ = .�Ƿ��ռ
        Set obj���к�����λ = GetUnitsObjects(GetUnitAll())
        For Each obj������λ In obj���к�����λ
            Set rsTemp = GetUnitReg(.��¼ID, obj������λ.������λ����, obj������λ.����, blnRecord)
            If Not rsTemp.EOF Then
                Set obj������λ���� = New ������λ����
                obj������λ����.���� = obj������λ.����
                obj������λ����.������λ���� = obj������λ.������λ����
                obj������λ����.ԤԼ���Ʒ�ʽ = Val(Nvl(rsTemp!���Ʒ�ʽ))
                Set obj������λ����.������Ϣ�� = GetTimeIntervalObjects(rsTemp)
                .������λ���Ƽ�.AddItem obj������λ����
            End If
        Next
    End With
    Get�����¼ = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetԤԼ����(ByVal lng����ID As Long, Optional lng��ԴId As Long) As Integer
    '�̶��������ԤԼ����
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    '1.��Դ�����ԤԼ����
    If lng��ԴId = 0 Then
        strSQL = "Select Nvl(Max(b.ԤԼ����), 0) As ԤԼ����" & vbNewLine & _
                "        From �ٴ����ﰲ�� A, �ٴ������Դ B" & vbNewLine & _
                "        Where a.��Դid = b.Id And a.����id = [1]" & vbNewLine
    Else
        strSQL = "Select Nvl(Max(b.ԤԼ����), 0) As ԤԼ����" & vbNewLine & _
                "        From �ٴ������Դ B" & vbNewLine & _
                "        Where B.id = [2]" & vbNewLine
    End If
    '2.ԤԼ��ʽ�����ԤԼ����
    strSQL = strSQL & _
            " Union All" & vbNewLine & _
            " Select Max(ԤԼ����) As ԤԼ���� From ԤԼ��ʽ" & vbNewLine
    
    strSQL = "Select ԤԼ���� From (" & strSQL & ") Where ԤԼ���� > 0 And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡԤԼ����", lng����ID, lng��ԴId)
    If Not rsTemp.EOF Then
        GetԤԼ���� = Val(Nvl(rsTemp!ԤԼ����))
    End If
    
    '3.ϵͳ����"�Һ�����ԤԼ����"
    If GetԤԼ���� = 0 Then GetԤԼ���� = Val(zlDatabase.GetPara(Val("66-�Һ�����ԤԼ����"), glngSys))
    '4.ȱʡ����
    If GetԤԼ���� = 0 Then GetԤԼ���� = 7
    
    '104266
    '�԰���Ϊ��λ,�����������Դ����ʱ�䡱��12:00:00-23:59:59�ڼ�ģ��򿪷�ԤԼ����+1��
    strSQL = "Select Zl_Fun_Getappointmentdays As ��ԤԼ���� From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�жϴ˿��Ƿ�࿪��һ��")
    If Not rsTemp.EOF Then
        GetԤԼ���� = GetԤԼ���� + Val(Nvl(rsTemp!��ԤԼ����))
    End If
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetVisitedRecord(ByVal lng��ԴId As Long, _
    ByVal str��ʼʱ�� As String, ByVal str��ֹʱ�� As String) As ADODB.Recordset
    '��ȡ��ǰ��Դ�������˵ĳ�������
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    '�ų��������ð��ŵ���ģ��
    strSQL = "Select b.����ID,b.ID as ����ID,b.��ԴID,a.��������" & vbNewLine & _
            " From �ٴ������¼ A,�ٴ����ﰲ�� B,�ٴ������ C" & vbNewLine & _
            " Where a.����ID=b.ID And a.��ԴID=[1] And a.�������� Between [2] And [3]" & vbNewLine & _
            "       And c.ID=b.����ID And Nvl(c.�Ű෽ʽ,0) In (0,1,2)"
    Set GetVisitedRecord = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��Դ�������˵ĳ����¼", lng��ԴId, _
        CDate(str��ʼʱ��), CDate(str��ֹʱ��))
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitedRecordByDate(ByVal lng����ID As Long, ByVal str�������� As String) As ADODB.Recordset
    '��ȡ��ǰ������ָ�����ڵĳ����¼
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandler
    strSQL = "Select a.Id, a.�ϰ�ʱ��," & vbNewLine & _
            "        Max(Decode(b.Id, Null, 0, 1)) As ��ʹ��," & vbNewLine & _
            "        Max(Decode(a.ͣ�￪ʼʱ��, Null, 0, 1)) As ��ͣ��" & vbNewLine & _
            " From �ٴ������¼ A, ���˹Һż�¼ B" & vbNewLine & _
            " Where a.Id = b.�����¼id(+) And a.����id = [1] And a.�������� = [2]" & vbNewLine & _
            " Group By a.Id, a.�ϰ�ʱ��"
    Set GetVisitedRecordByDate = _
        zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ǰ���������ڹҺ�ԤԼ�ĳ����¼", lng����ID, CDate(str��������))
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function CheckExistRecord(ByVal lng��ԴId As Long, ByVal strApply As String, _
    Optional ByVal obj���ﰲ�� As ���ﰲ��, Optional ByVal blnMonthTemplet As Boolean, _
    Optional ByVal lng����ID As Long) As Boolean
    '��鱻Ӧ�õ��������Ƿ����г����¼
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim ObjItem As �����¼��
    Dim varDate As Variant, i As Integer
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo errHandler
    If strApply = "" Then Exit Function
    If blnMonthTemplet Then
        strSQL = "Select /*+cardinality(b,10)*/ 1" & _
                " From �ٴ��������� A, Table(f_Str2list([2], '|')) B" & _
                " Where a.����ID = [1] And a.������Ŀ = b.Column_Value And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鱻Ӧ�õ��������Ƿ����г����¼", lng����ID, strApply)
        CheckExistRecord = Not rsTemp.EOF
        Exit Function
    End If
    
    If lng��ԴId <> 0 Then
        strSQL = "Select /*+cardinality(b,10)*/ 1" & _
                " From �ٴ������¼ A, Table(f_Str2list([2], '|')) B" & _
                " Where a.��ԴID = [1] And a.�������� = To_Date(b.Column_Value, 'yyyy-mm-dd') And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鱻Ӧ�õ��������Ƿ����г����¼", lng��ԴId, strApply)
        CheckExistRecord = Not rsTemp.EOF
        Exit Function
    End If
    
    If obj���ﰲ�� Is Nothing Then Exit Function
    varDate = Split(strApply, "|")
    For i = 0 To UBound(varDate)
        If blnFind Then Exit For
        If Not obj���ﰲ��.�ѱ�����ﰲ�� Is Nothing Then
            For Each ObjItem In obj���ﰲ��.�ѱ�����ﰲ��
                If obj���ﰲ��(1).�������� <> ObjItem.�������� Then
                    If IsDate(varDate(i)) Then
                        If DateDiff("d", ObjItem.��������, varDate(i)) = 0 Then blnFind = True: Exit For
                    Else
                        If ObjItem.�������� = varDate(i) Then blnFind = True: Exit For
                    End If
                End If
            Next
        End If
        If blnFind Then Exit For
        If Not obj���ﰲ��.δ������ﰲ�� Is Nothing Then
            For Each ObjItem In obj���ﰲ��.δ������ﰲ��
                If obj���ﰲ��(1).�������� <> ObjItem.�������� Then
                    If IsDate(varDate(i)) Then
                        If DateDiff("d", ObjItem.��������, varDate(i)) = 0 Then blnFind = True: Exit For
                    Else
                        If ObjItem.�������� = varDate(i) Then blnFind = True: Exit For
                    End If
                End If
            Next
        End If
    Next
    CheckExistRecord = blnFind
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SearchStopVisitReason(frmFrom As Object, objControl As Object, ByVal strInput As String) As String
    '����:ģ�����ң�����ͣ��ԭ��ѡ���б�
    '����:
    Dim strSQL As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput = "" Then Exit Function
    vRect = zlControl.GetControlRect(objControl.Hwnd)
    'ȥ��"'"
    strInput = Replace(strInput, "'", " ")
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If IsNumeric(strInput) Then '����ȫ������ʱֻƥ�����
            strWhere = " Where ���� Like Upper([1])"
        ElseIf zlCommFun.IsCharAlpha(strInput) Then '����ȫ����ĸʱֻƥ�����
            strWhere = " Where ���� Like Upper([1])"
        Else
            strWhere = " Where ���� Like Upper([1]) Or ���� Like [1] Or ���� Like Upper([1])"
        End If
    End If
    
    strSQL = "Select RowNum As ID, ����, ���� From ����ͣ��ԭ��" & strWhere
    Set rsTemp = zlDatabase.ShowSQLSelect(frmFrom, strSQL, 0, "ͣ��ԭ��", False, _
                   "", "", False, False, True, vRect.Left, vRect.Top, objControl.Height, blnCancel, True, False, strKey)
                   
    If blnCancel Then Exit Function
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> adStateOpen Then Exit Function
    
    SearchStopVisitReason = Nvl(rsTemp!����)
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetԤԼ�Һż�¼(ByVal lng��ԴId As Long, _
    ByVal dtStartTime As Date, ByVal dtEndTime As Date) As ADODB.Recordset
    '��ȡָ����Դ��ָ��ʱ�䷶Χ�ڵ�ԤԼ�Һ����ݣ���ͬ���ڵ�ֻ��ȡһ����¼
    Dim strSQL As String
    
    On Error GoTo errHandler
    strSQL = "Select Decode(To_Char(��������, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) As ������Ŀ," & vbNewLine & _
            "        ID As ��¼id, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, �Ƿ��ռ" & vbNewLine & _
            " From (Select a.Id, a.��������, a.�ϰ�ʱ��, a.��ʼʱ��, a.��ֹʱ��, �Ƿ��ռ," & vbNewLine & _
            "               Row_Number() Over(Partition By To_Char(a.��������, 'D'), a.�ϰ�ʱ�� Order By a.��������) As �к�" & vbNewLine & _
            "        From �ٴ������¼ A, ���˹Һż�¼ B" & vbNewLine & _
            "        Where a.Id = b.�����¼id And a.�ϰ�ʱ�� Is Not Null And a.��Դid = [1] And a.�������� Between [2] And [3])" & vbNewLine & _
            " Where �к� < 2" & vbNewLine & _
            " Order By To_Char(��������, 'D'), To_Date(To_Char(��ʼʱ��, 'hh24:mi:ss'), 'hh24:mi:ss')"
    Set GetԤԼ�Һż�¼ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ԤԼ�Һ����ݵĳ����¼", lng��ԴId, _
        CDate(Format(dtStartTime, "yyyy-mm-dd")), CDate(Format(dtEndTime, "yyyy-mm-dd")))
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
