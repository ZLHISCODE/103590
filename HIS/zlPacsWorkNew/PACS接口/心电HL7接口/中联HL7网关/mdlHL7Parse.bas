Attribute VB_Name = "mdlHL7Parse"
Option Explicit

Public Function GetMsgACK(strData As String, strACK As String) As Boolean
'-----------------------------------------------------------------------------
'����:��HL7��Ϣ�У���ȡ��Ϣ����֯����Ӧ��Ϣ
'����:
'       strData ��IN��  ---��Ϣ�ı�
'       strACK  ��OUT��---- ACK��Ϣ
'���أ�True -- �ɹ����գ�False -- ����ܾ�
'-----------------------------------------------------------------------------
    Dim strField() As String
    Dim i As Integer
    Dim strSendingApp As String
    Dim strSendingFac As String
    Dim strMsgControlID As String
    Dim strErrMsg As String
    Dim strMSH As String
    Dim strMSA As String
    Dim lngMsgFullType As Long
    
    On Error GoTo err
    
    'ACK��Ϣ����ԭ��Ϣ����ȡMSH-3��MSH-4��MSH-10���������ǹ̶�ֵ
    'MSH-3  �����ͳ�������
    'MSH-4  �������豸����
    'MSH-10 ����Ϣ����ID
    
    '�����Ϣ��������,'0 -- ��������Ϣ��1 -- ����Ϣͷ��2 -- ����Ϣβ��3 -- ����Ϣ�м�Σ�4 -- ����
    lngMsgFullType = funMsgFullType(strData)
    
    '��¼������־
    Call WriteProcessLog("GetMsgACK", "��֯��Ӧ��Ϣ", "��Ҫ��Ӧ����Ϣ�ǣ�" & strData & vbCrLf & "�����Ϣ���������ǣ�" & lngMsgFullType, 2)
    
    If lngMsgFullType = 0 Or lngMsgFullType = 1 Then
        '������Ϣ������Ϣͷ�����н���
        '��ȡԭ��Ϣ�е�MSH-3��MSH-4��MSH-10
        '��Ϊֻ�ǽ�����Ϣ�е�MSH�Σ����ֱ��ʹ�á�|����Ϊ�ָ���
        strField = Split(strData, "|")
        If Trim(strField(0)) = Chr(11) & "MSH" Then
            If UBound(strField) > 10 Then
                strSendingApp = strField(2)
                strSendingFac = strField(3)
                strMsgControlID = strField(9)
                
                '��֯ACK������Ϣ
                strMSH = Chr(11) & "MSH|^~\&|ZLHIS|HIS001|" & strSendingApp & "|" & strSendingFac & "|" & _
                            getDateTimeString(Now) & "||ACK|" & getDateTimeString(Now) & "|P|2.4|" & Chr(13)
                If lngMsgFullType = 0 Then
                    strMSA = "MSA|AA|" & strMsgControlID & "|" & Chr(28) & Chr(13)
                    GetMsgACK = True
                Else
                    strMSA = "MSA|AR|" & strMsgControlID & "|" & Chr(28) & Chr(13)
                    GetMsgACK = False
                End If
                strACK = strMSH & strMSA
                '��֯��ACK��Ӧ��ֱ���˳�
                Exit Function
            Else
                'MSH����Ϣ����
                strErrMsg = "MSH�����Ҳ�����Ϣ����ID"
            End If
        Else
            'MSH�δ���
            strErrMsg = "�Ҳ���MSH��"
        End If
    Else
        strErrMsg = "��Ϣ������"
    End If
    
    '�������󣬵�����֯һ��MSH��Ϣͷ
    strMSH = Chr(11) & "MSH|^~\&|ZLHIS|HIS001|SendingApp|SendingFac|" & _
                            getDateTimeString(Now) & "||ACK|" & getDateTimeString(Now) & "|P|2.4|" & Chr(13)
    strMSA = "MSA|AR||" & strErrMsg & Chr(28) & Chr(13)
    strACK = strMSH & strMSA
    GetMsgACK = False
    
    Exit Function
err:
    '���������ؿ���Ϣ
    '��¼������־
    Call WriteLog(1001, err.Number, "����ACK��Ӧ�������������Ϣ�ǣ� " & strData & "�����������ǣ�" & err.Description)
End Function

Public Function getDateTimeString(strDateTime As String) As String
'-----------------------------------------------------------------------------
'����:���ظ�ʽ���õ�ʱ���ַ����������뼶��
'����:
'       strDateTime  ---����ʱ���ı�
'���أ���ʽ�����ı�
'-----------------------------------------------------------------------------
    
    getDateTimeString = Format(strDateTime, "YYYYMMDDHHMMSS")
    
End Function



Public Function getXPN(strXPN As String) As String
'-----------------------------------------------------------------------------
'����:��PN����XPN�ֶ��У���ȡ����
'����:
'       strXPN  ---���յ���PN����XPN���͵��ַ���
'���أ�����
'-----------------------------------------------------------------------------
    Dim arrName() As String
    Dim strName As String
    Dim i As Integer
    
    '��ʽ����^����
    On Error GoTo err
    arrName = Split(strXPN, "^")
    
    If UBound(arrName) >= 0 Then
        strName = arrName(0)
        For i = 1 To UBound(arrName) - 1
            strName = strName & arrName(i)
        Next i
    End If
    
    Exit Function
err:
    '��¼������־
    Call WriteLog(1002, err.Number, "getXPN,�������������������������ǣ� " & strXPN & "�����������ǣ�" & err.Description)
End Function

Public Function getCMName(strCMName As String) As String
'-----------------------------------------------------------------------------
'����:��CM���͵��ֶ��У���ȡ������������OBR-32��OBR-33
'����:
'       strCMName  ---���յ���CM���͵������ַ���
'���أ�����
'-----------------------------------------------------------------------------
    Dim arrName() As String
    Dim strName As String
    Dim i As Integer
    Dim iEnd As Integer
    
    '��ʽ��^��^��^^^^^^^^1^��,"����^^^^^^^^^^^"
    
    On Error GoTo err
    arrName = Split(strCMName, "^")
    iEnd = UBound(arrName) - 1
    If iEnd > 4 Then iEnd = 4
    
    If UBound(arrName) >= 1 Then
        strName = arrName(0)
        For i = 1 To iEnd
            strName = strName & arrName(i)
        Next i
        getCMName = strName
    End If
    
    Exit Function
err:
    '��¼������־
    Call WriteLog(1003, err.Number, "getCMName,�������������������������ǣ� " & strCMName & "�����������ǣ�" & err.Description)

End Function

Public Function funParseInMsg(strMsg As String) As Long
'-----------------------------------------------------------------------------
'����:�����ʹ�����յ���HL7��Ϣ
'����:  strMsg -- ��Ϣ�ı�
'���أ� 0 -- �ɹ���1 -- ʧ��,��Ϣ���Ͳ�֧��
'-----------------------------------------------------------------------------
    Dim strSegments() As String
    Dim strFields() As String
    Dim i As Integer
    Dim strMsgType As String
    Dim strPatientID As String
    Dim strOrderID As String
    Dim strDoctor As String
    Dim strResultURL As String
    Dim strResultDiag As String
    Dim strSQL As String
    
    On Error GoTo err
    
    '��ʱֻ����ORU-R01��Ϣ
    
    '����Ϣ���ջس��ֶ�
    strSegments = Split(strMsg, Chr(13))
    
    '���ݶα�־��ѭ������ÿһ�ε���Ϣ����ȡ���е���Ϣ
    For i = 0 To UBound(strSegments) - 1
        strFields = Split(strSegments(i), "|")
        
        If UBound(strFields) > -1 Then
            If Trim(strFields(0)) = Chr(11) & "MSH" Then
                'MSH��
                '��ȡMSH-9����Ϣ���ͣ��ж��Ƿ� ��ORU-R01����Ϣ
                If UBound(strFields) >= 8 Then
                    strMsgType = strFields(8)
                    If strMsgType <> "ORU^R01" Then
                        Call WriteProcessLog("funParseInMsg", "��Ϣ���Ͳ�֧��", "��Ϣ�����ǣ�" & strMsgType & "����������޷�����������Ϣ��", 2)
                        funParseInMsg = 1   '��Ϣ���Ͳ�֧��
                        Exit Function
                    End If
                End If
            ElseIf strFields(0) = vbLf & "PID" Or strFields(0) = "PID" Then
                'PID��
                '��ȡPID-2,Patient ID ����ID ��Ϊ����ID���ж�����
                If UBound(strFields) >= 2 Then
                    strPatientID = strFields(2)
                End If
            ElseIf strFields(0) = vbLf & "OBR" Or strFields(0) = "OBR" Then
                'OBR��
                '��ȡOBR-2��������ҽ������,ҽ��ID����Ϊ��ѯ����¼���������
                '��ȡOBR-32�������Ҫ������+����Ϊִ�з��ò���Ա����
                If UBound(strFields) >= 32 Then
                    strOrderID = strFields(2)
                    strDoctor = getCMName(strFields(32))
                End If
            ElseIf strFields(0) = vbLf & "OBX" Or strFields(0) = "OBX" Then
                'OBX��
                'OBX-11"�۲���״̬"��ȱʡֵ��F����ʾ����������ֵ�п��ܱ�ʾ�����Ҫ�����»����滻�����ж����ֵ�����н��ֱ���滻��
                
                '��ȡ�ĵ緵�ص�URL�������
                '��ȡOBX-2��Value Type ֵ����,����ֵ=��RP�����Ƿ��ص�URL����
                '��ȡOBX-3��Observation Identifier �۲��ʶ��,��ʶ��=��MUSEWebURL������URL����
                '��ȡOBX-5��Observation Value �۲�ֵ���۲�ֵ�����ݾ���URL����
                'URL������Ҫע�⣬���������ƿ����ǻ���������Ҫת��IP��ַ�����Ӵ��е�\T\��Ҫת���&
                
                '��ȡ�ĵ練���ı�������
                'OBX|88|FT|ECGMEASANDDIAG||Test Reason : ~Blood Pressure : ***/*** mmHG~Vent. Rate : 079 BPM     Atrial Rate : 079 BPM~   P-R Int : 150 ms          QRS Dur : 086 ms~    QT Int : 394 ms       P-R-T Axes : 065 013 034 degrees~   QTc Int : 451 ms~~������� ~~Referred By:             //�ĵ�ͼ��ʾ�ļ��Ĳ���
                'Overread By: �¾� ��||||||D|        //��ҽ��ҽ��
                '��ȡOBX-2��Value Type ֵ����,����ֵ=��FT�����Ƿ��صı�������
                '��ȡOBX-3��Observation Identifier �۲��ʶ��,��ʶ��=��ECGMEASANDDIAG�����Ǳ�������
                '��ȡOBX-5��Observation Value �۲�ֵ���۲�ֵ�����ݾ��Ǳ�������������ֵ��Ϊһ���ı�������ʹ�ûس��滻��~�����ţ����浽���浥�ġ��������С�
                
                If UBound(strFields) >= 5 Then
                    '�����URL
                    If strFields(2) = "RP" And strFields(3) = "MUSEWebURL" Then
                        strResultURL = strFields(5)
                        strResultURL = Replace(strResultURL, "\T\", "&")
                        strResultURL = Replace(strResultURL, "'", "��")
                    End If
                    
                    'ECG���
                    If strFields(2) = "FT" And strFields(3) = "ECGMEASANDDIAG" Then
                        strResultDiag = strFields(5)
                        '����������ݣ�ʹ�ûس��滻��~������
                        strResultDiag = Replace(strResultDiag, "~", vbCrLf)
                        '�������ݴ�����ʹ��˫�ֽڵġ��������浥�ֽڵ�"'"
                        strResultDiag = Replace(strResultDiag, "'", "��")
                    End If
                End If
            End If
        End If
    Next i
    
    '�����ȡ����Ϣ�Ƿ���ȷ����ȷ�򱣴浽���ݿ���
    If strResultURL <> "" Then
        '�ĵ�ͼ�ļ���������浽������ҽ������.ִ��˵�����С�
        strSQL = "zlhis.b_Hl7interface.Recevieresult(" & strOrderID & ", '" & strDoctor & "','" & strResultURL & "')"
        
        '��¼������־
        Call WriteProcessLog("funParseInMsg", "׼�������ĵ�����", "���ô洢���� =" & strSQL, 3)
        
        gzlDatabase.ExecuteProcedure strSQL, "���յ��ĵ�����"
                
        '��¼��Ϣ��¼
        Call WriteMessageLog("�����ĵ�����", "ҽ��ID = " & strOrderID & "�����ҽ��=" & strDoctor & "���������=" & strResultURL)
    End If
    If strResultDiag <> "" Then
        '�ĵ�������������浽���浥�ġ������������
        strSQL = "zlhis.b_Hl7interface.SendReport(" & strOrderID & ",'" & strResultDiag & "',NULL,'" & strDoctor & "')"
        
        '��¼������־
        Call WriteProcessLog("funParseInMsg", "׼�������ĵ籨������", "���ô洢���� =" & strSQL, 3)
        
        gzlDatabase.ExecuteProcedure strSQL, "�����ĵ籨������"
        
        '��¼��Ϣ��¼
        Call WriteMessageLog("�����ĵ籨������", "ҽ��ID = " & strOrderID & "�����ҽ��=" & strDoctor & "����������=" & strResultDiag)
    End If
    
    
    Exit Function
err:
    '��¼������־
    Call WriteLog(1004, err.Number, "funParseInMsg������Ϣ������������ǰ�����Ϣ�ǣ� " & Left(strMsg, 250) & "�����������ǣ�" & err.Description)
    Call WriteLog(1004, err.Number, "funParseInMsg������Ϣ����ҽ��ID = " & strOrderID & "�����ҽ��=" & strDoctor & "���������=" & strResultURL & "����������=" & strResultDiag)
End Function

Public Function funParseACK(strMsg As String, strACK As String) As Long
'-----------------------------------------------------------------------------
'����:�����ʹ�����յ���ACK��Ϣ��ͨ����Ϣ����ID�ж�ACK�Ƿ���ȷ������
'����:  strMsg -- ���͵���Ϣ�ı�
'       strACK -- ���յ���ACK�ı�
'���أ� 0 -- �ɹ���1 -- ʧ��,���͵���Ϣ����ȷ;2 -- ���յ��Ĳ���ACK��Ϣ;3 -- �յ�ACK��Ϣ������û�б��Է�����
'-----------------------------------------------------------------------------
    Dim i As Integer
    Dim strSegments() As String
    Dim strFields() As String
    Dim strMsgControlID As String
    Dim blnSendMsgOK As Boolean
    Dim blnIsACK As Boolean
    Dim blnACKOK As Boolean
    
    On Error GoTo err
    
    '����Ϣ���ջس��ֶ�
    strSegments = Split(strMsg, Chr(13))
    
    '��ȡ��Ϣ����ID
    If UBound(strSegments) <> -1 Then
        strFields = Split(strSegments(i), "|")
        If UBound(strFields) > 10 Then
            If strFields(0) = Chr(11) & "MSH" Then
                strMsgControlID = strFields(9)
                blnSendMsgOK = True
            End If
        End If
    End If
    
    If blnSendMsgOK = True Then
        '����ACK��Ϣ
        strSegments = Split(strACK, Chr(13))
        
        For i = 0 To UBound(strSegments) - 1
            strFields = Split(strSegments(i), "|")
            If UBound(strFields) > -1 Then
                If Trim(strFields(0)) = Chr(11) & "MSH" Then
                    'MSH�Σ�MSG-9,��Ϣ����
                    If UBound(strFields) > 8 Then
                        If strFields(8) = "ACK" Then
                            blnIsACK = True
                        Else
                            Exit For
                        End If
                    End If
                ElseIf strFields(0) = vbLf & "MSA" Or strFields(0) = "MSA" Then
                    'MSA�Σ�MSA-1 ȷ�ϴ��룻MSA-2 ��Ϣ����ID
                    If UBound(strFields) >= 2 Then
                        If strFields(1) = "AA" And (strFields(2) = strMsgControlID Or strFields(2) = strMsgControlID & Chr(28)) Then
                            'AA��ʾ������ACK
                            blnACKOK = True
                        ElseIf strFields(1) = "AE" And (strFields(2) = strMsgControlID Or strFields(2) = strMsgControlID & Chr(28)) And (UCase(strFields(3)) = UCase("Duplicate Order Record")) Then
                            'AE��ʾ�����ACK�������������������Duplicate Order Record��˵�����ҽ���Ѿ����ͳɹ��ˣ����Ƿ��ͳɹ����������·����ˡ�
                            blnACKOK = True
                        End If
                    End If
                End If
            End If
        Next i
        
        If blnIsACK = True Then
            If blnACKOK = True Then
                '��������
                funParseACK = 0
            Else
                '���ճ��ִ���
                Call WriteLog(1005, err.Number, "funParseACK�����յ�ACK��Ϣ������û�гɹ���ACK��Ϣ�ǣ� " & strACK)
                funParseACK = 3
                Exit Function
            End If
        Else
            '���յ�����Ϣ����ACK��Ϣ
            Call WriteLog(1006, err.Number, "funParseACK�����յ�����Ϣ����ACK��Ϣ������Ϣ�ǣ� " & strACK)
            funParseACK = 2
            Exit Function
        End If
        
    Else
        '���͵���Ϣ����Ͳ���ȷ����¼������־
        Call WriteLog(1007, err.Number, "funParseACK�����͵���Ϣ��ʽ����ȷ�����͵���Ϣ�ǣ� " & strMsg)
        '���ش�����Ϣ
        funParseACK = 1
        Exit Function
    End If
    
    Exit Function
err:
    '��¼������־
    Call WriteLog(1008, err.Number, "funParseACK,������Ϣ��������������Ϣ�ǣ� " & strMsg & "�����������ǣ�" & err.Description)
End Function
