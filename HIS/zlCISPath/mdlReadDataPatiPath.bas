Attribute VB_Name = "mdlReadDataPatiPath"
Option Explicit
'---------------------------------------------------------------------------------------
'��ģ��ֻ�������ݲ���ķ���
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : ReadPathPhase
' Author    : YWJ
' Date      : 2019-04-29
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function ReadPathPhase(ByVal lngPatiPathID As Long, ByVal lngPhaseBranchId As Long) As ADODB.Recordset
'����:
'lngPatiPathID-����·����¼ID
'lngPhaseBranchId-�׶η�֧ID
    Dim strSQL As String
    
    On Error GoTo errH
    '�׶�����ʱ�� NVL(c.���,b.���) ��Ϊ�˴����÷�֧������������⣬ȡֵb.��� ����Ϊ��������Ҫ��ʾ�ǵڼ�����֧����ȡ��֧·�������ʱ��ȡ����һ�׶ε���ż��Ϸ�֧·������ţ�
    If lngPhaseBranchId = 0 Then
        strSQL = _
        "Select a.�׶�id, a.����, To_Char(a.����, 'yyyy-mm-dd') ����, To_Char(a.����, 'day') ����, b.���� As �׶���, b.���, b.˵��, b.��id,b.·��ID,b.��ʼ���� " & vbNewLine & _
                 "From (Select a.�׶�id, a.����, a.����,a.·����¼id " & vbNewLine & _
                 "       From ����·��ִ�� A" & vbNewLine & _
                 "       Where a.·����¼id = [1]" & vbNewLine & _
                 "       Group By a.�׶�id, a.����, a.����,a.·����¼id) A, �ٴ�·���׶� B,�ٴ�·���׶� C,����·������ G" & vbNewLine & _
                 "Where a.�׶�id = b.Id And b.��id=c.id(+) And g.·����¼id(+) = a.·����¼id And g.�׶�id(+) = a.�׶�id And g.����(+) = a.���� " & vbNewLine & _
                 "Order By ����,g.�Ǽ�ʱ��,NVL(c.���,b.���)"
    Else
        strSQL = _
        "Select a.�׶�id, a.����, To_Char(a.����, 'yyyy-mm-dd') ����, To_Char(a.����, 'day') ����, b.���� As �׶���, b.���, b.˵��, b.��id,b.·��ID,b.��ʼ���� " & vbNewLine & _
                 "From (Select a.�׶�id, a.����, a.����,a.·����¼id " & vbNewLine & _
                 "       From ����·��ִ�� A" & vbNewLine & _
                 "       Where a.·����¼id = [1]" & vbNewLine & _
                 "       Group By a.�׶�id, a.����, a.����,a.·����¼id) A, �ٴ�·���׶� B,�ٴ�·���׶� C,�ٴ�·����֧ D,�ٴ�·���׶� E,�ٴ�·���׶� F,����·������ G" & vbNewLine & _
                 "Where a.�׶�id = b.Id And b.��id=c.id(+) And b.��֧id=d.id(+) and d.ǰһ�׶�id=e.id(+) And e.��id=f.id(+) And g.·����¼id(+) = a.·����¼id And g.�׶�id(+) = a.�׶�id And g.����(+) = a.���� " & vbNewLine & _
                 "Order By ����,g.�Ǽ�ʱ��, Decode(b.��֧ID,Null,NVL(c.���,b.���),NVL(c.���,b.���)+NVL(f.���,e.���))"
    End If

    Set ReadPathPhase = zlDatabase.OpenSQLRecord(strSQL, "ReadPathPhase", lngPatiPathID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReadPathPhaseNoEvaluate
' Author    : YWJ
' Date      : 2019-04-29
' Purpose   :��ȡ��������·���׶�
'---------------------------------------------------------------------------------------
'
Public Function ReadPathPhaseNoEvaluate(ByVal lngPatiPathID As Long, ByVal lngPhaseBranchId As Long) As ADODB.Recordset
'����:
'lngPatiPathID-·����¼ID
'lngPhaseBranchId-�׶η�֧ID
    Dim strSQL As String
    
    On Error GoTo errH
    '�׶�����ʱ�� NVL(c.���,b.���) ��Ϊ�˴����÷�֧������������⣬ȡֵb.��� ����Ϊ��������Ҫ��ʾ�ǵڼ�����֧����ȡ��֧·�������ʱ��ȡ����һ�׶ε���ż��Ϸ�֧·������ţ�
    '����·������
    If lngPhaseBranchId = 0 Then
        strSQL = "Select a.�׶�id, a.����, To_Char(a.����, 'yyyy-mm-dd') ����" & vbNewLine & _
                "From (Select a.�׶�id, a.����, a.����, a.·����¼id" & vbNewLine & _
                "       From ����·��ִ�� A" & vbNewLine & _
                "       Where a.·����¼id = [1]" & vbNewLine & _
                "       Group By a.�׶�id, a.����, a.����, a.·����¼id) A, �ٴ�·���׶� B, �ٴ�·���׶� C, ����·������ G" & vbNewLine & _
                "Where a.�׶�id = b.Id And b.��id = c.Id(+) And g.·����¼id(+) = a.·����¼id And g.�׶�id(+) = a.�׶�id And g.����(+) = a.���� And" & vbNewLine & _
                "      Not Exists" & vbNewLine & _
                " (Select 1 From ����·������ P Where p.·����¼id = a.·����¼id And p.�׶�id = a.�׶�id And p.���� = a.����)" & vbNewLine & _
                "Order By ����, g.�Ǽ�ʱ��, Nvl(c.���, b.���)"
                 
    Else
        strSQL = _
        "Select a.�׶�id, a.����, To_Char(a.����, 'yyyy-mm-dd') ����" & vbNewLine & _
                 "From (Select a.�׶�id, a.����, a.����,a.·����¼id " & vbNewLine & _
                 "       From ����·��ִ�� A" & vbNewLine & _
                 "       Where a.·����¼id = [1]" & vbNewLine & _
                 "       Group By a.�׶�id, a.����, a.����,a.·����¼id) A, �ٴ�·���׶� B,�ٴ�·���׶� C,�ٴ�·����֧ D,�ٴ�·���׶� E,�ٴ�·���׶� F,����·������ G" & vbNewLine & _
                 "Where a.�׶�id = b.Id And b.��id=c.id(+) And b.��֧id=d.id(+) and d.ǰһ�׶�id=e.id(+) And e.��id=f.id(+) And g.·����¼id(+) = a.·����¼id And g.�׶�id(+) = a.�׶�id And g.����(+) = a.���� And" & vbNewLine & _
                "      Not Exists" & vbNewLine & _
                " (Select 1 From ����·������ P Where p.·����¼id = a.·����¼id And p.�׶�id = a.�׶�id And p.���� = a.����)" & vbNewLine & _
                 "Order By ����,g.�Ǽ�ʱ��, Decode(b.��֧ID,Null,NVL(c.���,b.���),NVL(c.���,b.���)+NVL(f.���,e.���))"
    End If

    Set ReadPathPhaseNoEvaluate = zlDatabase.OpenSQLRecord(strSQL, "ReadPathPhaseNoEvaluate", lngPatiPathID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDiagType
' Author    : YWJ
' Date      : 2019-05-08
' Purpose   : ��ȡ���˱��ξ���������
'---------------------------------------------------------------------------------------
Public Function GetDiagType(ByVal lngPatiID As Long, ByVal lngVisitID As Long) As ADODB.Recordset
'����:
'   lngPatiID   -����ID
'   lngVisitID  -��ҳID
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select Distinct Nvl(a.�������, 'D') As ������� From ������ϼ�¼ A Where a.����id = [1] And a.��ҳid = [2]"
    Set GetDiagType = zlDatabase.OpenSQLRecord(strSQL, "GetDiagType", lngPatiID, lngVisitID)
   
   Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
