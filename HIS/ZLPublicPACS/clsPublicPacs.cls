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

Private mobjReport As Object  'zlRichEPR.cDockReport  '���Ӳ����������
Private mblnIsInit As Boolean   '�жϴ˽ӿ��Ƿ���г�ʼ������
Private mfrmReport As frmReport '�����ĵ��༭������

Public Function InitInterface(ByVal cnOracle As ADODB.Connection, ByVal strDBUser As String, Optional ByVal blnUseAlone As Boolean = False) As Boolean
'�Խӿ��������ݽ��г�ʼ�����ɹ�����TRUE
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    
    Set gcnOracle = cnOracle
    
    If gcnOracle Is Nothing Then Exit Function
    
    If gcnOledb Is Nothing Then
        Call InitOledbConn(blnUseAlone)
    End If
    
    If gobjComLib Is Nothing Then
        Set gobjComLib = GetObject("", "zl9ComLib.clsComLib")
    End If
    
    gstrPrivs = ";" & gobjComLib.GetPrivFunc(100, 9004) & ";"
        
    InitInterface = True
    mblnIsInit = True
End Function

Public Function zlHasImage(lngOrderID As Long) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ�����һ��ҽ��ID���ӿڷ���,�������ҽ��ID�����ж��Ƿ���Ӱ��ͼƬ
'������lngOrderID--����ҽ��ID���
'���أ�true--�ò�����Ӱ��ͼƬ  false--�ò���û��Ӱ��ͼƬ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error GoTo errHandle
    
    zlHasImage = mdlPublic.HasImage(lngOrderID)
    
    Exit Function
errHandle:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetReportImage(ByVal lngAdviceID As Long, ByVal strDocReportId As String, _
    ByVal strTmpImgDir As String, Optional ByRef strMsg As String = "") As Collection
'����:�����ҽ������Ӧ�����е�����ͼ�񣬱��浽ָ��Ŀ¼,��ȡPACS�����е�����ͼ��
'lngAdviceId:ҽ��id
'strDocReportId:PACS���ܱ���༭���ĵ�ID���������PACS���ܱ���༭����д�ı��棬�򲻴��ݸò���
'strTmpImgFolder:ͼ�񻺴�Ŀ¼
'objImgFileName:����ͼ���ļ�������

'˵�����˹���Ŀǰֻ�������ã�һ��ҽ����Ӧһ�ݱ���
    Dim objFSO As New Scripting.FileSystemObject
    Dim strSql As String
    Dim rsData As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strTmpImgFile As String
    Dim objImgFileName As New Collection
    Dim objFtp As New clsFtp
    Dim strFtpDir As String
    Dim strIp As String
    Dim strFtpPswd As String
    Dim strFTPUser As String
    Dim strSubDir As String
    Dim lngPatFrom As Long
    
    If mblnIsInit = False Then
        Call ShowErrInfo
        Exit Function
    End If
     
    If lngAdviceID <= 0 Then Exit Function
    
    '����ָ��ĿĿ¼
    If Not objFSO.FolderExists(strTmpImgDir) Then MkLocalDir strTmpImgDir
  
    If Trim(strDocReportId) <> "" Then '˵���Ǳ����ĵ��༭���ı���
        Set GetReportImage = frmReport.GetReportImage(lngAdviceID, strDocReportId, strTmpImgDir, strMsg)
        Exit Function
    End If
    
    strSql = "Select D.�豸��,D.FTP�û���,D.FTP����,D.IP��ַ,'/'||D.FtpĿ¼||'/' As Root, " & _
             "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')||C.���UID As URL,d.�豸�� as �豸�� " & _
             "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
             "Where  C.λ��һ=D.�豸��(+) and C.ҽ��ID=[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡFTP��Ϣ", lngAdviceID)
    
    If rsTemp.EOF Then
        strMsg = "�˱���û��ͼ��"
        Exit Function
    End If
    
    If rsTemp.RecordCount > 0 Then
        If Trim(Nvl(rsTemp("�豸��"))) = "" Then Exit Function
        
        strFtpDir = Nvl(rsTemp("Root"))
        strIp = Nvl(rsTemp("IP��ַ"))
        strFtpPswd = Nvl(rsTemp("FTP����"))
        strFTPUser = Nvl(rsTemp("FTP�û���"))
        strSubDir = Nvl(rsTemp("URL"))
    End If
    
    strSql = "select c.�������� " & _
             "from ����ҽ������ a,���Ӳ������� b,���Ӳ������� c " & _
             "Where a.ҽ��ID =[1] And a.����id = b.�ļ�id And b.Id = c.��id " & _
             "and b.�������� = 3 And Substr(b.��������, Instr(b.��������, ';', 1, 18) + 1, 1) = '2' " & _
             "and Nvl(c.��ֹ��, 0) = 0 "

    Set rsData = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��������", lngAdviceID)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    '����FTP
    If objFtp.FuncFtpConnect(strIp, strFTPUser, strFtpPswd) = 0 Then
        strMsg = "FTP�����������ӣ������������á�"
        Exit Function
    End If

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
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    If mblnIsInit = False Then
        Call ShowErrInfo
        Exit Function
    End If
       
    If lngPatId <= 0 Then Exit Function
    
    strSql = "select B.����ID AS ����ID,A.ҽ������ AS ����,C.������ AS �����,C.���ʱ�� " & _
             "from ����ҽ����¼ A,����ҽ������ B,���Ӳ�����¼ C,Ӱ�����¼ D " & _
             "Where A.����ID = [1] And A.��ҳID = [2] And A.Id = B.ҽ��ID And B.����ID = C.Id " & _
             "And A.�������='D' and ���ID is null And C.���ʱ�� Is Not NULL " & _
             "And A.ID=D.ҽ��ID and A.ҽ����Ч=1 And A.ҽ��״̬ IN (3,5,6,7,8) " & _
             "order by C.���ʱ�� desc "
    
    Set rsData = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ʾ�����嵥", lngPatId, lngPageId)
    
    Set GetReportList = rsData
End Function

Public Function GetReportFormHandle(ByVal lngReportID As Long) As Long
'����:������ʾ�������ݣ�����Ԥ��
'lngReportID:����id
'���ش��ھ��
    Dim lngAdviceID As Long
    
    If mblnIsInit = False Then
        Call ShowErrInfo
        Exit Function
    End If

    Set mobjReport = CreateObject("zlRichEPR.cDockReport")

    If mobjReport Is Nothing Then Exit Function

    lngAdviceID = GetAdviceID(lngReportID)
    mobjReport.zlRefresh 0, 0
    mobjReport.zlRefresh lngAdviceID, 0

    GetReportFormHandle = mobjReport.zlGetForm().hWnd
End Function

Public Sub ShowImage(ByVal lngAdviceID As Long, ByVal frmParent As Object, Optional ByVal blnMoved As Boolean = False)
'����:���ݴ����ҽ�����򿪶�Ӧ��Ӱ���Ƭվ���й�Ƭ
'lngAdviceId: ҽ��ID
'frmParent:������
'blnMoved:�Ƿ�ת�ƹ�
    Call mdlPublic.ViewImage(lngAdviceID, frmParent, blnMoved)
    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Sub ShowPatientImage(ByVal lngAdviceID As Long, ByVal frmParent As Object, Optional ByVal blnMoved As Boolean = False)
'����:���ݴ����ҽ��ID�����ҵ���Ӧ�Ĳ���ID����רҵ��PACS�ļ���б���Ƭ
'lngAdviceId: ҽ��ID
'frmParent:������
'blnMoved:�Ƿ�ת�ƹ�
    Call mdlPublic.ViewPatientImage(lngAdviceID, frmParent, blnMoved)
    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Sub ShowStaticImage(ByVal lngAdviceID As Long, ByVal frmParent As Object, Optional ByVal blnMoved As Boolean = False)
'����:�鿴�ؼ�ͼ��
'lngAdviceId: ҽ��ID
'frmParent:������
'blnMoved:�Ƿ�ת�ƹ�
    Call mdlPublic.ViewStaticImage(lngAdviceID, frmParent, blnMoved)
    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Function Show3DImage(ByVal lngҽ��ID As Long, frmParent As Object) As Long
'���ܣ�3D��Ƭ
'0   �ɹ�
'-121    ���ò�������
'-122    δ��ȷ��װPACS���ӿ��ļ�
'-102    δ��ȷ��װPACS���ӿ��ļ�
'-103    �������Ӵ���
'-104    ���ݿ����
'-101    ��������
    Show3DImage = mdlPublic.View3DImage(lngҽ��ID, frmParent)
End Function


Public Function GetEditerType(ByVal lngDeptId As Long) As Integer
'���������ҽ����Ӧ��ִ�п���ID
'���ܣ���ȡ��ǰʹ�õı༭��ģʽ
'���أ�0--���Ӳ����༭��, 1--PACS�����������2--�����ĵ��༭��
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    If mblnIsInit = False Then
        Call ShowErrInfo
        Exit Function
    End If
    
    GetEditerType = 0
    
    On Error GoTo ErrH
    '��ȡ��ǰ�༭��ģʽ
    strSql = "Select ����ֵ From Ӱ�����̲��� Where ����ID=[1] And ������=[2]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetReportContent", lngDeptId, "����༭��")
    
    If rsTmp.RecordCount > 0 Then GetEditerType = Val(Nvl(rsTmp!����ֵ))
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
    GetEditerType = 0
End Function

Public Function zlDocGetList(ByVal lngPatId As Long, Optional ByVal lngPageId As Long, Optional ByVal strRegNo As String) As Recordset
'����Pacs�ĵ��༭����д�ı����б�
'lngPatID:����ID
'lngPageID:��ҳID
'strRegNo:�Һŵ�
On Error GoTo ErrH
    
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Set zlDocGetList = mfrmReport.zlDocGetList(lngPatId, lngPageId, strRegNo)
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Set zlDocGetList = Nothing
End Function

Public Function zlDocGetListWithAdvice(ByVal strAdviceId As String) As Recordset
'����Pacs�ĵ��༭����д�ı����б�
On Error GoTo ErrH
    
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Set zlDocGetListWithAdvice = mfrmReport.zlDocGetListWithAdvice(strAdviceId)
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Set zlDocGetListWithAdvice = Nothing
End Function

Public Function zlDocGetForm(Optional blnShow As Boolean = False, Optional blnRIS As Boolean = False, Optional blnAutoView As Boolean = True) As Object
'���ܣ�����Pacs�ĵ��༭���ı���鿴���ڶ���
'������     blnShow -- �Ƿ���ʾ��ߵĹ�������
'           blnRIS --  �Ƿ�RIS����RIS����ʾ�˳��Ȱ�ť
'           blnAutoView -- �Ƿ��Զ�����
    
On Error GoTo ErrH
    If mfrmReport Is Nothing Then
        Set mfrmReport = New frmReport
        
        '�����Ƿ���ʾ�����������˳���ť�����������Ҫ����BorderStyleǰ�棬ȷ��FormLoadʱblnShow��������ȷ��
        '����BorderStyleʱ�ᴥ��FormLoad
        mfrmReport.mblnShow = blnShow
        mfrmReport.mblnRIS = blnRIS
        mfrmReport.mblnAutoView = blnAutoView
        
        mfrmReport.BorderStyle = 0
        mfrmReport.Caption = mfrmReport.Caption
    End If
    
    Set zlDocGetForm = mfrmReport
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Set zlDocGetForm = Nothing
End Function

Public Sub zlDocRefresh(ByVal strReportId As String)
'ˢ�¼�鱨��id��Ӧ�ı�������
On Error GoTo ErrH
    
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Call mfrmReport.RefreshReportContent(strReportId)
    
    Exit Sub
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Public Sub zlDocViewStateUpdate(ByVal blnState As Boolean, ByVal lngAdviceID As Long)
'����pacs�ĵ��༭������Ĳ���״̬,���blnState=True����ʾ��Ӧ�����б���Ϊ���ģ�blnState=False,��ʾΪδ��

On Error GoTo ErrH
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Call mfrmReport.zlDocViewStateUpdate(blnState, lngAdviceID)
    Exit Sub
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub


Public Sub zlDocShowReport(ByVal lngAdviceID As Long, Optional ByVal strReportId As String = "", Optional ByVal blnAutoView As Boolean = True, Optional objParent As Object, Optional blnShowModal As Boolean = False)
'���ܣ���ʹ�ñ����ĵ��༭��ʱ������ҽ��ID��ȡ�������ݲ���ʾ
'objParent:������
'lngAdviceId:ҽ��ID
'blnAutoView:�Ƿ��Զ����ģ�True--��
    On Error GoTo ErrH
    
    If mblnIsInit = False Then
        Call ShowErrInfo
        Exit Sub
    End If
    
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Call mfrmReport.ShowMe(lngAdviceID, strReportId, blnAutoView, objParent, blnShowModal)
    
    Exit Sub
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Public Sub PrintReport(ByVal strDocId As String, Optional ByVal strPrinterName As String = "", Optional ByVal blnIsPreview As Boolean = False)
'���ܣ���ӡ����
'����˵��:
'strDocId---����ID
'strPrinterName---��ӡ�����ƣ�Ϊ��ʱ������ӡ���ÿ�
'blnIsPreview---ΪTrueʱ��ʾԤ������
    On Error GoTo ErrH
    
    Call frmPrintReport.PrintReport(strDocId, strPrinterName, blnIsPreview)
    Exit Sub
    
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Public Function zlDocRefreshForm(ByVal lngAdviceID As Long, Optional ByVal strReportId As String = "", Optional objParent As Object) As Boolean
    '���ܣ� ˢ�±����ĵ��༭�����ڣ�ʹ����RIS�ӿ��У���RIS�д�PACS����ʷ�����ͼ��
    '������ lngAdviceId -- ҽ��ID
    '       strReportId -- ����ID���ɿ�
    '       objParent -- ������
    On Error GoTo err
    
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Call mfrmReport.RefreshForm(lngAdviceID, strReportId, objParent)
    
    zlDocRefreshForm = True
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Private Sub ShowErrInfo()
    MsgBox "�ӿڻ�δ��ʼ�������ܽ�����Ӧ������", vbInformation, "ϵͳ��Ϣ"
End Sub

Private Sub Class_Initialize()
    mblnIsInit = False
End Sub

Private Sub Class_Terminate()
    mblnIsInit = False
    
    If Not mfrmReport Is Nothing Then Set mfrmReport = Nothing
    If Not mobjReport Is Nothing Then Set mobjReport = Nothing
End Sub