Attribute VB_Name = "mdlPrint"
Option Explicit

Public Function ReportPrint(ByVal lngKey As Long, ByVal frmPrint As Form, ByVal objPrint As Object, ByVal blnPrint As Boolean) As String
        '����:  ���������ӡ
        '����:
        '       lngKey          �걾ID
        '
        '       blnPrint        True=��ӡ
        Dim strReportCode As String         '�����ʽ���
        Dim strReportParaNo As String       '���뵥��
        Dim bytReportParaMode As Byte       'ҽ����¼����
        Dim rsTmp As New ADODB.Recordset
        Dim blnCurrMoved As Boolean         '�����Ƿ���ת��
        Dim strSQL As String
        Dim strChart(0 To 9) As String
        Dim intLoop As Integer
        Dim strErr As String
        Dim lngAdviceID As Long, lngPatiID As Long
        On Error GoTo errH
    
     
         '����ͼ�ι��Զ��屨�����
100     strSQL = "Select ID,ҽ��ID,����ID from ����걾��¼ where id = [1] "
102     Set rsTmp = ComOpenSQL(strSQL, "ReportPrint", lngKey)
    
104     Do Until rsTmp.EOF
106         lngAdviceID = Val("" & rsTmp!ҽ��ID)
108         lngPatiID = Val("" & rsTmp!����ID)
        
110         If ReadSampleImage(lngKey, strChart, strErr) = False Then
                'ʧ�ܣ�����ʾ������ûͼ��
112             If strErr <> "" Then ShowLog LOG_PRINTSVR, LOG_WARNING, "��ӡ����", 100, "��ȡͼ��ʧ�ܣ�" & strErr
            End If
        
114         If GetReportCode(lngAdviceID, 0, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
                If strReportCode = "" Then
                    ReportPrint = "���鱨���ʽδ�趨��"
                    Exit Function
                Else
116             Call objPrint.ReportOpen(gcnOracle, 100, strReportCode, frmPrint, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & lngAdviceID, _
                                "����ID=" & lngPatiID, "�걾ID=" & lngKey, "���ҽ��=" & lngAdviceID, "����걾=" & lngKey, _
                                "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), "ͼ��4=" & strChart(3), _
                                "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                                "ͼ��9=" & strChart(8), IIf(blnPrint, 2, 1))
                End If
            End If
    
    
            On Error GoTo errH

118         strSQL = "ZL_����걾��¼_�걾�ʿ�(" & rsTmp("ID") & ",'',1)"
120         ComExecuteProc strSQL, "��ӡ"
122         rsTmp.MoveNext
        Loop
        ReportPrint = "OK"
        On Error Resume Next
        'ɾ��ͼ���ļ�
126     For intLoop = 1 To 9
128         Kill strChart(intLoop)
        Next
130
        Exit Function
errH:
132     ReportPrint = "ReportPrint " & CStr(Erl()) & "�У�" & Err.Description
134     ShowLog LOG_PRINTSVR, LOG_ERR, "��ӡ����", Err.Number, ReportPrint
End Function

Private Function GetReportCode(ByVal lngAdviceID As Long, ByVal lng���ͺ� As Long, _
                               ByRef strCode As String, ByRef strNo As String, _
                               ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False) As Boolean
    
        '����;  ��ȡ������
    
        Dim rs As New ADODB.Recordset
        Dim strSQL As String
    
        On Error GoTo errH
    
100     If lngAdviceID = 0 And lng���ͺ� = 0 Then Exit Function
    
    '    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2' AS ������," & _
                           "A.NO," & _
                           "A.��¼���� " & _
                    "FROM ����ҽ������ A,�����ļ��б� C,����ҽ����¼ D,��������Ӧ�� E " & _
                    "Where E.�����ļ�id = C.ID " & _
                            "AND D.������ĿID=E.������ĿID " & _
                          "AND A.ҽ��ID=D.ID AND E.Ӧ�ó���=Decode(D.������Դ,2,2,4,4,1) " & _
                          " AND D.���id= [1] "
                      
102     strSQL = "Select Distinct 'ZLCISBILL' || Trim(To_Char(C.���, '00000')) || '-2' As ������, A.NO, Nvl(A.��¼����,1) as ��¼����, F.ID, F.����" & vbNewLine & _
                "From ����ҽ������ A, �����ļ��б� C, ����ҽ����¼ D, ��������Ӧ�� E, ������ĿĿ¼ F" & vbNewLine & _
                "Where E.�����ļ�id = C.ID And D.������Ŀid = E.������Ŀid And D.������Ŀid = F.ID And A.ҽ��id = D.ID And" & vbNewLine & _
                "      E.Ӧ�ó��� = Decode(D.������Դ, 2, 2, 4, 4, 1) And D.���id = [1] " & vbNewLine & _
                "Order By F.���� "
                          
104     If DataMoved Then
106         strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
108         strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        End If

110     Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lngAdviceID, lng���ͺ�)
                      
    
112     If rs.BOF = False Then
114         strCode = Trim("" & rs("������"))
116         strNo = Trim("" & rs("NO"))
118         bytMode = Val("" & rs("��¼����"))
        End If
    
120     GetReportCode = True
        Exit Function
errH:
122     ShowLog LOG_PRINTSVR, LOG_ERR, "ȡ������", Err.Number, CStr(Erl()) & "��," & Err.Description

End Function
