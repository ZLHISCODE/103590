VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mblnLogin As Boolean    '�Ƿ��Ѿ���¼
Private mfrmPrint As frmMain    '���ô�ӡ����Ҫ�õĴ���
Private mobjPrint As Object     '��ӡ����

Public Function zlLogin(ByVal strUserName As String, ByVal strPwd As String, ByVal strSvr As String) As String
    '����:      ��¼
    '����:
    '       UserName    �û���
    '       strPwd      ��¼��������̨������
    '       strSvr      ������
    '����
    '       OK ��ʾ�ɹ�;������Ϊ��ʾ��Ϣ
    Dim strErr As String
    
    mblnLogin = False
    zlLogin = OraDataOpen(strSvr, strUserName, TranPasswd(strPwd))
    
    If zlLogin = "OK" Then
        strErr = ""
        If ComInitComLib(strErr) Then
            mblnLogin = True
        Else
            If strErr = "" Then strErr = "��ʼ����������ʧ�ܣ�"
            zlLogin = strErr
        End If
    End If
    
End Function
Public Function zlGetZyPrintList(ByVal strInPatientNo As String, ByVal strOutDate As String) As String
    '���ܣ�  ��ȡסԺ���߿ɴ�ӡ����
    '������  strInPatientNo    String  סԺ��
    '        strOutDate        strdate ��Ժ���ڣ�����ȷ������סԺ����
    
    '���أ�  strOut          string  OK|�걾ID|�걾��|����|�Ա�|����|�������|
    '                                   ������|����ʱ��|ִ�п���|��������|������Ŀ|��ӡ����;
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim dateS As Date, dateE As Date, strReturn As String
    On Error GoTo hErr
    ShowLog LOG_PRINTSVR, LOG_INFO, "��ȡסԺ�����嵥", 100, "��� " & strInPatientNo & "," & strOutDate
    If Not mblnLogin Then
        zlGetZyPrintList = "δ��¼!"
        Exit Function
    End If
    If strOutDate = "" Then
        zlGetZyPrintList = "��Ժ����Ϊ��!"
        Exit Function
    End If
    If Not IsDate(strOutDate) Then
        zlGetZyPrintList = "����ĳ�Ժ���ڲ�����������!"
        Exit Function
    Else
        dateS = Format(CDate(strOutDate), "yyyy-MM-dd 00:00:00")
        dateE = Format(CDate(strOutDate), "yyyy-MM-dd 23:59:59")
    End If
    strSQL = "Select a.Id, a.�걾���, a.����, a.�Ա�, a.����, b.���� As �������, a.������, a.����ʱ��, c.���� As ִ�п���, a.���ʱ��, a.����ʱ��, a.������Ŀ, a.��ӡ���� " & vbNewLine & _
            "From ����걾��¼ A, ���ű� B, ���ű� C, ������ҳ D" & vbNewLine & _
            "Where a.�������id = b.Id(+) And a.ִ�п���id = c.Id(+) And a.���ʱ�� Is Not Null And ������Դ = 2 And" & vbNewLine & _
            "      a.��ҳid = d.��ҳid And a.����id = d.����id And d.סԺ�� = [1] And d.��Ժ���� Between [2] And [3]"

    Set rsTmp = ComOpenSQL(strSQL, "", strInPatientNo, dateS, dateE)
    strReturn = ""
    Do Until rsTmp.EOF
        strReturn = strReturn & "^" & rsTmp!ID & "|" & rsTmp!�걾��� & "|" & rsTmp!���� & "|" & rsTmp!�Ա� & "|" & rsTmp!���� & "|" & rsTmp!������� & "|" & rsTmp!������
        strReturn = strReturn & "|" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & rsTmp!ִ�п��� & "|" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & rsTmp!������Ŀ
        strReturn = strReturn & "|" & Val("" & rsTmp!��ӡ����)
        rsTmp.MoveNext
    Loop
    If strReturn <> "" Then
        zlGetZyPrintList = "OK" & strReturn
    Else
        zlGetZyPrintList = "��סԺ��Ϊ" & strInPatientNo & ", ��Ժ����Ϊ" & strOutDate & "�ı��棡"
    End If
    
    ShowLog LOG_PRINTSVR, LOG_INFO, "��ȡסԺ�����嵥", 100, "���� " & zlGetZyPrintList
    Exit Function
hErr:
    strReturn = CStr(Erl()) & "��," & Err.Description
    zlGetZyPrintList = "zlGetZyPrintList " & strReturn
    ShowLog LOG_PRINTSVR, LOG_ERR, "��ȡסԺ�����嵥", Err.Number, strReturn
    
End Function

Public Function zlGetPrintList(ByVal strOutPatientNO As String, Optional ByVal intDays As Integer = 7) As String
    '���ܣ�  ��ȡ���߿ɴ�ӡ����
    '������  OutPatientNO    String  �����
    '        intDays         integer ��������ʾȡ��������ڵı��棬Ĭ��ȡ7���ڵı���
    '���أ�  strOut          string  OK|�걾ID|�걾��|����|�Ա�|����|�������|
    '                                   ������|����ʱ��|ִ�п���|��������|������Ŀ;
    Dim strSQL As String, dNow As Date, dateS As Date, dateE As Date
    Dim strKeyName As String, rsTmp As ADODB.Recordset
    Dim strReturn As String
    On Error GoTo hErr
    
    ShowLog LOG_PRINTSVR, LOG_INFO, "��ȡ�����嵥", 100, "��� " & strOutPatientNO & "," & intDays
    If Not mblnLogin Then
        zlGetPrintList = "δ��¼��"
        Exit Function
    End If
    If Not (intDays >= 1 And intDays <= 100) Then intDays = 7
    
    dNow = ComCurrDate
    
    dateS = Format(dNow - intDays, "yyyy-MM-dd")
    dateE = Format(dNow, "yyyy-MM-dd 23:59:59")
    
     
    strSQL = "Select a.Id, a.�걾���, a.����, a.�Ա�, a.����, b.���� As �������, a.������, a.����ʱ��, c.���� As ִ�п���, a.���ʱ��, a.����ʱ��, a.������Ŀ" & vbNewLine & _
            "From ����걾��¼ A, ���ű� B, ���ű� C" & vbNewLine & _
            "Where a.�������id = b.Id(+) And a.ִ�п���id = c.Id(+) And a.���ʱ�� Is Not Null And Nvl(A.��ӡ����,0)<=0 And ������Դ<>2 And a.����� = [1] And" & vbNewLine & _
            "      a.����ʱ�� Between [2] And [3]"

    '
    Set rsTmp = ComOpenSQL(strSQL, "", Val(strOutPatientNO), dateS, dateE)
    strReturn = ""
    Do Until rsTmp.EOF
        strReturn = strReturn & "^" & rsTmp!ID & "|" & rsTmp!�걾��� & "|" & rsTmp!���� & "|" & rsTmp!�Ա� & "|" & rsTmp!���� & "|" & rsTmp!������� & "|" & rsTmp!������
        strReturn = strReturn & "|" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & rsTmp!ִ�п��� & "|" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & rsTmp!������Ŀ
        rsTmp.MoveNext
    Loop
    If strReturn <> "" Then
        zlGetPrintList = "OK" & strReturn
    Else
        zlGetPrintList = "���" & intDays & "���ڣ�û�������Ϊ" & strOutPatientNO & "�Ŀɴ�ӡ���棡"
    End If
    ShowLog LOG_PRINTSVR, LOG_INFO, "��ȡ�����嵥", 100, "���� " & zlGetPrintList
    Exit Function
hErr:
    strReturn = CStr(Erl()) & "��," & Err.Description
    zlGetPrintList = "zlGetPrintList " & strReturn
    ShowLog LOG_PRINTSVR, LOG_ERR, "��ȡ�����嵥", Err.Number, strReturn
    
End Function

Public Function zlPrintReport(ByVal lngSampleID As Long) As String
    '����:   ��ӡ���ߵ�ָ������
    '������  SampleID        long       �����zlGetPrintList�з��صı걾ID
    '���أ�  strOut          string     OK-��ӡ�ɹ�;����Ϊʧ�ܵ���ʾ��Ϣ
    If mfrmPrint Is Nothing Then Set mfrmPrint = New frmMain
    If mobjPrint Is Nothing Then Set mobjPrint = CreateObject("zl9Report.clsReport")
    If Not mblnLogin Then
        zlPrintReport = "δ��¼��"
        Exit Function
    End If
    
    zlPrintReport = ReportPrint(lngSampleID, mfrmPrint, mobjPrint, True)

End Function

Public Function zlLoginOut() As String
    '����:   �˳�����ϵͳ
    '����:   ��
    '���أ�  strOut      string  OK-�ɹ�;����Ϊʧ�ܵ���ʾ��Ϣ
    
    mblnLogin = False
    Call OraDataClose
    Call FreeImageObj
    
    If Not mfrmPrint Is Nothing Then Unload mfrmPrint
    zlLoginOut = "OK"
End Function
