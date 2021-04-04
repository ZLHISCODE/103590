VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicPacs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InitInterface(ByVal cnOracle As ADODB.Connection, ByVal strDBUser As String) As Boolean
'�Խӿ��������ݽ��г�ʼ�����ɹ�����TRUE
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    
    Set gcnOracle = cnOracle
    
    If gcnOracle Is Nothing Then Exit Function
    
    If gobjComLib Is Nothing Then
        Set gobjComLib = CreateObject("zl9ComLib.clsComLib")

        Call gobjComLib.InitCommon(gcnOracle)
        Call gobjComLib.SetDbUser(UCase(strDBUser))
        Call gobjComLib.RegCheck
    End If
    
    InitInterface = True
End Function

Public Function GetReportImage(ByVal lngAdviceId As Long, ByVal strTmpImgDir As String) As Collection
'����:�����ҽ������Ӧ�����е�����ͼ�񣬱��浽ָ��Ŀ¼,��ȡPACS�����е�����ͼ��
'lngAdviceId:ҽ��id
'strTmpImgFolder:ͼ�񻺴�Ŀ¼
'objImgFileName:����ͼ���ļ�������
    Dim objFSO As New Scripting.FileSystemObject
    Dim strSQL As String
    Dim rsData As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strTmpImgFile As String
    Dim objImgFileName As New Collection
    Dim objFtp As New clsFtp
    Dim strFtpDir As String
    Dim strIP As String
    Dim strFtpPswd As String
    Dim strFTPUser As String
    Dim strSubDir As String
     
    If lngAdviceId <= 0 Then Exit Function
    
    '����ָ��ĿĿ¼
    If Not objFSO.FolderExists(strTmpImgDir) Then MkLocalDir strTmpImgDir
    
    strSQL = "Select D.�豸��,D.FTP�û���,D.FTP����,D.IP��ַ,'/'||D.FtpĿ¼||'/' As Root, " & _
             "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')||C.���UID As URL,d.�豸�� as �豸�� " & _
             "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
             "Where  C.λ��һ=D.�豸��(+) and C.ҽ��ID=[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡFTP��Ϣ", lngAdviceId)
    
    If rsTemp.EOF Then
        MsgBox "û�б���ͼ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rsTemp.RecordCount > 0 Then
        If Trim(Nvl(rsTemp("�豸��"))) = "" Then Exit Function
        
        strFtpDir = Nvl(rsTemp("Root"))
        strIP = Nvl(rsTemp("IP��ַ"))
        strFtpPswd = Nvl(rsTemp("FTP����"))
        strFTPUser = Nvl(rsTemp("FTP�û���"))
        strSubDir = Nvl(rsTemp("URL"))
    End If
    
    '����FTP
    If objFtp.FuncFtpConnect(strIP, strFTPUser, strFtpPswd) = 0 Then
        MsgBox "FTP�����������ӣ������������á�", vbExclamation, gstrSysName
        Exit Function
    End If
    
    strSQL = "select c.�������� " & _
             "from ����ҽ������ a,���Ӳ������� b,���Ӳ������� c " & _
             "Where a.ҽ��ID =[1] And a.����id = b.�ļ�id And b.Id = c.��id " & _
             "and b.�������� = 3 And Substr(b.��������, Instr(b.��������, ';', 1, 18) + 1, 1) = '2' " & _
             "and Nvl(c.��ֹ��, 0) = 0 "

    Set rsData = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��������", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function

    Do While Not rsData.EOF
        strTmpImgFile = Split(rsData!��������, ";")(11)
        
        If Trim(strTmpImgFile) <> "" Then
            '������ͼ���ļ����뼯��
            objImgFileName.Add strTmpImgFile
            
            '�ļ�������ɾ��
            If objFSO.FileExists(strTmpImgDir & "\" & strTmpImgFile) Then objFSO.DeleteFile (strTmpImgDir & "\" & strTmpImgFile)
            
            'ͨ��FTP���ر���ͼ��ָ��Ŀ¼
            Call objFtp.FuncDownloadFile(strFtpDir & strSubDir, strTmpImgDir & "\" & strTmpImgFile, strTmpImgFile)
        End If
        
        rsData.MoveNext
    Loop
    
    Set GetReportImage = objImgFileName
    
    '�ر�FTP����
    objFtp.FuncFtpDisConnect
End Function

Public Function GetReportList(ByVal lngPatId As Long, ByVal lngPageId As Long) As ADODB.Recordset
'����:������ʾ�����嵥������Ԥ��
'lngPatID:����id
'lngPageID:��ҳid
'����ADO��¼������������ID�����ƣ�����ˣ����ʱ��
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
       
    If lngPatId <= 0 Then Exit Function
    
    strSQL = "select B.����ID AS ����ID,A.ҽ������ AS ����,C.������ AS �����,C.���ʱ�� " & _
             "from ����ҽ����¼ A,����ҽ������ B,���Ӳ�����¼ C,Ӱ�����¼ D " & _
             "Where A.����ID = [1] And A.��ҳID = [2] And A.Id = B.ҽ��ID And B.����ID = C.Id " & _
             "And A.�������='D' and ���ID is null And C.���ʱ�� Is Not NULL " & _
             "And A.ID=D.ҽ��ID(+) and A.ҽ����Ч=1 And A.ҽ��״̬ IN (3,5,6,7,8) " & _
             "order by C.���ʱ�� desc "
    
    Set rsData = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ʾ�����嵥", lngPatId, lngPageId)
    
    Set GetReportList = rsData
End Function

Public Function GetReportFormHandle(ByVal lngReportID As Long) As Long
'����:������ʾ�������ݣ�����Ԥ��
'lngReportID:����id
'���ش��ھ��
    Dim objfrmReport As New frmReport
    
    GetReportFormHandle = objfrmReport.zlInitModule(lngReportID)
End Function

Public Sub ShowImage(ByVal lngAdviceId As Long, ByVal frmParent As Object, Optional ByVal blnMoved As Boolean = False)
'����:���ݴ����ҽ�����򿪶�Ӧ��Ӱ���Ƭվ���й�Ƭ
'lngAdviceId: ҽ��ID
'frmParent:������
'blnMoved:�Ƿ�ת�ƹ�
    Dim objPacsCore As Object
    Dim strFtpHost As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strSDPath As String
    Dim strSDUser As String
    Dim strSDPwd As String
    Dim intImageLocation As Long
    Dim objFSO As New Scripting.FileSystemObject
    Dim lng����ID As Long
    
    On Error GoTo DBError
    If lngAdviceId <= 0 Then Exit Sub
    
    strSQL = "Select ����ID From ����ҽ������ Where ҽ��ID=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lngAdviceId)
    If Not rsTmp.EOF Then lng����ID = rsTmp!����id
    
    '���ж��Ƿ���Լ�������
    If CheckEPRReport(lngAdviceId, IIf(lng����ID = 0, -1, lng����ID)) <> 1 Then
        MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)�����ڲ��ܲ鿴ͼ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���ж��Ƿ����ͼ��û��ͼ������ʾ���˳�
    strSQL = "Select A.���UID,Count(B.����UID) as �������� From Ӱ�����¼ A,Ӱ�������� B Where A.���UID=B.���UID And A.ҽ��ID=[1] Group by A.���UID"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��Ƭ����", lngAdviceId)
    If rsTmp.EOF Then
        MsgBox "û�п����ڹ�Ƭ�ı���ͼ��", vbInformation, gstrSysName
        Exit Sub
    End If

    '������Ҫ�򿪵�����ͼ����Ϣ
    strSQL = "Select /*+RULE*/ D.IP��ַ As Host1,d.�豸�� as �豸��1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'\')" & _
        "||C.���UID||'\' As Path,E.IP��ַ As Host2,e.�豸�� as �豸��2, " & _
        "D.����Ŀ¼ AS ����Ŀ¼1, E.����Ŀ¼ AS ����Ŀ¼2,D.����Ŀ¼�û��� as ����Ŀ¼�û���1, " & _
        "E.����Ŀ¼�û��� AS ����Ŀ¼�û���2,D.����Ŀ¼���� AS ����Ŀ¼����1,E.����Ŀ¼���� AS ����Ŀ¼����2 " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) And C.ҽ��ID=[1] "
        
    '�����ת����־�����ȡת������ʷ��
    If blnMoved Then
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ŀ¼��Ϣ", lngAdviceId)
    
    If rsTmp.RecordCount > 0 Then
        '�������صĻ���Ŀ¼����Ҫ�ڵ��ù�Ƭվ֮ǰ�ȴ������Ŀ¼����Ƭվ��ֻ�����أ����������ػ���Ŀ¼
        If Not objFSO.FolderExists(App.Path & "\TmpImage\" & rsTmp("Path")) Then
            MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        End If
        
        ClearCacheFolder App.Path & "\TmpImage\"
        
        '��ȡFTP�����������û��������룬IP��ַ��
        If rsTmp("�豸��1") <> "" Then
            strFtpHost = rsTmp("Host1")
            strSDPath = Nvl(rsTmp("����Ŀ¼1"))
            strSDUser = Nvl(rsTmp("����Ŀ¼�û���1"))
            strSDPwd = Nvl(rsTmp("����Ŀ¼����1"))
        ElseIf Nvl(rsTmp("�豸��2")) <> "" Then
            strFtpHost = rsTmp("Host2")
            strSDPath = Nvl(rsTmp("����Ŀ¼2"))
            strSDUser = Nvl(rsTmp("����Ŀ¼�û���2"))
            strSDPwd = Nvl(rsTmp("����Ŀ¼����2"))
        End If
        
        '�жϹ���Ŀ¼�Ƿ��Ѿ����ӣ����û�����ӣ����������
        On Error Resume Next
        If strSDPath <> "" Then
            Call funcConnectShardDir("\\" & strFtpHost & "\" & strSDPath, strSDUser, strSDPwd)
        End If
        
        Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
        objPacsCore.CallOpenViewer "", lngAdviceId, frmParent, gcnOracle, blnMoved, False
    End If

    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Function CheckEPRReport(ByVal lngҽ��ID As Long, Optional lng����ID As Long, Optional blnBySign As Boolean, Optional ByVal intִ��״̬ As Integer = -999) As Integer
'���ܣ�����Ӧ��Ŀ�ı�����д���
'������lngҽ��ID=�ɼ��е�ҽ��ID
'      lng����ID=���Դ��룬��Ҫ���ڷ��ر��没��ID
'      intִ��״̬=���ڼ������ʱ�������ۺϵ�ִ��״̬
'������blnBySign=�����Ƿ����ͨ��ǩ�������ж�(����ҽ������վ)
'���أ�0-���滹û����д
'      1-��������д���(��ǩ��,�����޶���ǩ��,����ִ�����)
'      2-����δ��д���(δǩ��,���޶���δǩ��,��δִ�����)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '��鱨���Ƿ�����д
    If lng����ID = 0 Then
        strSQL = "Select ����ID From ����ҽ������ Where ҽ��ID=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lngҽ��ID)
        If Not rsTmp.EOF Then lng����ID = rsTmp!����id
    End If
    If lng����ID = 0 Then
        CheckEPRReport = 0: Exit Function
    End If
    
    If Not blnBySign Then
        '��鱨��ִ�й���(5-���;6-�������)��״̬(1-���)
        '���鱨���ǹ������ɼ���ʽ����ģ����ɼ���ʽ����Ϊ����δ�������ͼ�¼
        strSQL = _
            " Select 2 as ����,ҽ��ID,ִ�й���,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
            " Union ALL" & _
            " Select ����,ҽ��ID,ִ�й���,Decode([2],-999,ִ��״̬,[2]) as ִ��״̬,����ʱ��" & _
            " From (" & _
                " Select 1 as ����,B.ҽ��ID,B.ִ�й���,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
                " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                    " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
                " Order by A.���" & _
            " ) Where Rownum=1" & _
            " Order by ����,����ʱ�� Desc"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lngҽ��ID, intִ��״̬)
        If Nvl(rsTmp!ִ�й���, 0) >= 5 Or Nvl(rsTmp!ִ��״̬, 0) = 1 Then
            CheckEPRReport = 1
        Else
            CheckEPRReport = 2
        End If
    Else
        'ͨ��ǩ���汾�жϱ�����ɵķ�ʽ
        strSQL = "Select B.�ļ�ID,Max(B.��ʼ��) as ǩ���汾 From ���Ӳ������� B Where B.�ļ�ID=[1] And B.��������=8 Group by B.�ļ�ID"
        strSQL = "Select B.���ʱ��,B.���汾,C.ǩ���汾 From ���Ӳ�����¼ B,(" & strSQL & ") C Where B.ID=[1] And B.ID=C.�ļ�ID(+)"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng����ID)
            
        '(ǩ������ֱ���޸ģ������޶������ǩ�������汾Ӧ��ǩ���汾һ��)
        If IsNull(rsTmp!���ʱ��) Or Nvl(rsTmp!���汾, 0) <> Nvl(rsTmp!ǩ���汾, 0) Then
            '���ҽ�������Ѿ�ִ��,��ʹû��ǩ���򲻷�Ҳ��ͬ���
            strSQL = _
                " Select 2 as ����,ҽ��ID,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
                " Union ALL" & _
                " Select ����,ҽ��ID,Decode([2],-999,ִ��״̬,[2]) as ִ��״̬,����ʱ��" & _
                " From (" & _
                    " Select 1 as ����,B.ҽ��ID,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
                    " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                        " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
                    " Order by A.���" & _
                " ) Where Rownum=1" & _
                " Order by ����,����ʱ�� Desc"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lngҽ��ID, intִ��״̬)
            If Nvl(rsTmp!ִ��״̬, 0) = 1 Then
                CheckEPRReport = 1
            Else
                CheckEPRReport = 2
            End If
        Else
            CheckEPRReport = 1
        End If
    End If
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function