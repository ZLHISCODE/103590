VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrImageShareDir As String                  '�ϰ��Ӱ�����洢Ŀ¼
Private mblnXWLog As Boolean                        '�Ƿ��¼ͨѶ��־
Private mlngStudySchemeNo As Long                    '��鷽����
Private mlngSeriesSchemeNo As Long                   '���з�����
Private mblnUseXinWangView As Boolean
Private mConnectedShardDir() As String   '�Ѿ����ӹ��Ĺ���Ŀ¼���豸������
Private mobjPacsCore As Object      '��Ƭվ����

Private mobjGetImage As Object
Private mlngSys As Long

Public Function OpenViewer(ByVal lngSys As Long, cnOracle As ADODB.Connection, ByVal strDbUser As String, ByVal lngViewerType As Long, lngAdviceID As Long, _
        blnAddImage As Boolean, Optional ByVal strSerials As String = "", _
        Optional ByVal blnMoved As Boolean = False, Optional ByVal blnLocalizerBackward As Boolean = False, _
        Optional ByVal intImageInterval As Integer = 0, Optional ByVal strImageString As String = "") As Long
'------------------------------------------------
'���ܣ����ݴ����ҽ��ID�ͷ��ͺţ���objPacsCoreָ��Ĺ�Ƭվ
'������
'       lngViewerType -- չ��ͼ���Viewer���ͣ�1-�����ר��Viewer��2-�ٴ������Viewer
'       objPacsCore ������Ƭվ����
'       lngAdviceID ����ҽ��ID
'       blnAddImage--True ��ԭ��ͼ����������ӵ�ǰͼ��Falseɾ��ԭ��ͼ�񣬴򿪵�ǰͼ��
'       objParent -- ������
'       strSerials������ѡ������UID���ƴ����ö��ŷָ�����������룬��ѡ��ȫ������
'       blnMoved������ѡ���Ƿ�ת��
'       blnLocalizerBackward--��ѡ����λ�����,��strImageString����
'       intImageInterval ---��ѡ����ͼ��ļ��������5����ʾÿ5��ͼ��һ��ͼ,��strImageString����
'       strImageString --- ��ѡ��ÿ����������Ҫ�򿪵�ͼ�����ϣ���intImageInterval��blnLocalizerBackward���⣬
'                           ��strImageStringΪ��
'                           �����ǡ�����UID1|1-3;5-27;33-100+����UID2|ȫ����,ȫ����ʾ��ȫ��ͼ��
'���أ�ͼ���ļ���������
'------------------------------------------------
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim strFTPHost As String, strFtpPath As String, strFTPUser As String, strFTPPswd As String
    Dim strSDPath As String, strSDUser As String, strSDPwd As String
    Dim strDeviceNO As String
    Dim i As Integer
    Dim blnConnectDS As Boolean         '�Ƿ����ӵ�ǰ�Ĺ���Ŀ¼
    Dim oneMessage As TGetImgMsg        'Ԥȡͼ�����Ϣ�ṹ
    Dim intImageLocation As Integer
    Dim strXWViewerFilter As String
    
    On Error GoTo DBError
    
    Call InitComLib(cnOracle, strDbUser)
    
    mblnXWLog = (Val(zlDatabase.GetPara("XW��¼�ӿ���־", mlngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1) '�Ƿ��¼�ӿ���־
    mstrImageShareDir = zlDatabase.GetPara("XW��ʷͼ����Ŀ¼", mlngSys, G_LNG_XWPACSVIEW_MODULE, "DCMSHARE")
    mlngStudySchemeNo = Val(zlDatabase.GetPara("XW��鷽����", mlngSys, G_LNG_XWPACSVIEW_MODULE, "1"))
    mlngSeriesSchemeNo = Val(zlDatabase.GetPara("XW���з�����", mlngSys, G_LNG_XWPACSVIEW_MODULE, "2"))
    mblnUseXinWangView = IsUseXwViewer
    
    BUGEX "mblnUseXinWangView = " & mblnUseXinWangView, True
    
    mlngSys = lngSys
    
    '��ѯͼ��������PACS����������PACS
    strSql = "Select ͼ��λ��,Ӱ����� from Ӱ�����¼ where ҽ��ID =[1]"
    
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ�����ڵ�λ��", lngAdviceID)
    
    If rsTmp.RecordCount <> 0 Then
        intImageLocation = Nvl(rsTmp!ͼ��λ��, 0)
    End If
    
    BUGEX "lngAdviceID = " & lngAdviceID, True
    BUGEX "intImageLocation = " & intImageLocation, True
    BUGEX "mblnXWLog = " & mblnXWLog, True
    
    'ͼ�����������ݿ⣬��������DLL��ʾͼ��
    If intImageLocation = 1 Then
        strXWViewerFilter = lngAdviceID & IIf(strSerials <> "", "[;]" & strSerials, "")
        
        If mblnXWLog = True Then
            Call WriteCommLog("OpenViewer", "����XWShowImage�ӿ�", "��ѯ���˲���Ϊ��" & strXWViewerFilter)
        End If
        
        Call XWShowImage(lngViewerType, strXWViewerFilter, IIf(strSerials <> "", mlngSeriesSchemeNo, mlngStudySchemeNo))
        
        OpenViewer = -1
        
        Exit Function
    End If
    
    '�ж��Ƿ��������°�pacs��Ƭ
    If mblnUseXinWangView = True Then
        Call OpenViewerWithInXWPacs(lngAdviceID, Nvl(rsTmp!Ӱ�����), blnMoved)
        
        OpenViewer = -1
        Exit Function
    End If
    
    
    'ͼ�����������ݿ⣬��������zl9PacsCore��ʾͼ��
    strFTPHost = ""
           
    '������Ҫ�򿪵�����ͼ����Ϣ
    strSql = "Select D.IP��ַ As Host1,d.�豸�� as �豸��1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/' As Path,E.IP��ַ As Host2,e.�豸�� as �豸��2, " & _
        "D.����Ŀ¼ AS ����Ŀ¼1, E.����Ŀ¼ AS ����Ŀ¼2,D.����Ŀ¼�û��� as ����Ŀ¼�û���1, " & _
        "E.����Ŀ¼�û��� AS ����Ŀ¼�û���2,D.����Ŀ¼���� AS ����Ŀ¼����1,E.����Ŀ¼���� AS ����Ŀ¼����2, " & _
        "D.FTPĿ¼ as FTPĿ¼1, E.FTPĿ¼ as FTPĿ¼2,D.FTP�û��� as FTP�û���1, E.FTP�û��� AS FTP�û���2,  " & _
        "D.FTP���� as FTP����1, E.FTP���� AS FTP����2 " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) And C.ҽ��ID=[1] "
    
    '�����ת����־�����ȡת������ʷ��
    If blnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����Ŀ¼��Ϣ", lngAdviceID)
    
    If rsTmp.RecordCount > 0 Then
        '�������صĻ���Ŀ¼����Ҫ�ڵ��ù�Ƭվ֮ǰ�ȴ������Ŀ¼����Ƭվ��ֻ�����أ����������ػ���Ŀ¼
        MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        
        '��ȡFTP�����������û��������룬IP��ַ��
        If rsTmp("�豸��1") <> "" Then
            strDeviceNO = rsTmp("�豸��1")
            strFTPHost = rsTmp("Host1")
            strFtpPath = Nvl(rsTmp("FTPĿ¼1"))
            strFTPUser = Nvl(rsTmp("FTP�û���1"))
            strFTPPswd = Nvl(rsTmp("FTP����1"))
            strSDPath = Nvl(rsTmp("����Ŀ¼1"))
            strSDUser = Nvl(rsTmp("����Ŀ¼�û���1"))
            strSDPwd = Nvl(rsTmp("����Ŀ¼����1"))
        ElseIf Nvl(rsTmp("�豸��2")) <> "" Then
            strDeviceNO = rsTmp("�豸��2")
            strFTPHost = rsTmp("Host2")
            strFtpPath = Nvl(rsTmp("FTPĿ¼2"))
            strFTPUser = Nvl(rsTmp("FTP�û���2"))
            strFTPPswd = Nvl(rsTmp("FTP����2"))
            strSDPath = Nvl(rsTmp("����Ŀ¼2"))
            strSDUser = Nvl(rsTmp("����Ŀ¼�û���2"))
            strSDPwd = Nvl(rsTmp("����Ŀ¼����2"))
        End If
        
        '�жϹ���Ŀ¼�Ƿ��Ѿ����ӣ����û�����ӣ����������
        blnConnectDS = True
        For i = 1 To UBound(mConnectedShardDir)
            If mConnectedShardDir(i) = strDeviceNO Then
                blnConnectDS = False
                Exit For
            End If
        Next i
        If blnConnectDS = True And strSDPath <> "" Then
            If funcConnectShardDir(Nothing, "\\" & strFTPHost & "\" & strSDPath, strSDUser, strSDPwd) = 0 Then
                ReDim Preserve mConnectedShardDir(UBound(mConnectedShardDir) + 1) As String
                mConnectedShardDir(UBound(mConnectedShardDir)) = strDeviceNO
            End If
        End If
        
        If mobjPacsCore Is Nothing Then Set mobjPacsCore = DynamicCreate("zl9PacsCore.clsViewer", "��Ƭվ")
        
        '�򿪹�Ƭվ
        If mobjPacsCore Is Nothing Then
            Exit Function
        Else
            OpenViewer = mobjPacsCore.OpenViewer(strImageString, lngAdviceID, cnOracle, blnMoved, blnAddImage, intImageInterval, lngSys)
        End If
        
        '�ȴ򿪹�Ƭվ����Ԥȡ
        oneMessage.strSubDir = rsTmp("Path")
        oneMessage.strDestMainDir = App.Path & "\TmpImage\"
        oneMessage.strIP = strFTPHost
        oneMessage.strFtpDir = strFtpPath
        oneMessage.strFTPUser = strFTPUser
        oneMessage.strFTPPswd = strFTPPswd
        oneMessage.strSDDir = strSDPath
        oneMessage.strSDUser = strSDUser
        oneMessage.strSDPswd = strSDPwd
        
        If CBool(GetSetting("ZLSOFT", "����ģ��", "UseActiveLoad", "True")) Then
            oneMessage.strFtpDir = "/" & oneMessage.strFtpDir & "/"
            Call funPreDownLoadImagesEx(Me, oneMessage)
        Else
            Call funPreDownLoadImages(oneMessage)
        End If
        
    Else    'û�в��ҵ�ͼ���¼����ر�ԭ���Ѿ��򿪵Ĺ�Ƭ����
        If Not mobjPacsCore Is Nothing Then mobjPacsCore.Closefrom
    End If
    
    Exit Function
DBError:
    If gobjComlib.ErrCenter() = 1 Then Resume
    BUGEX "OpenViewer err= " & err.Description, True
End Function

Private Function funPreDownLoadImages(thisMsg As TGetImgMsg) As Boolean
'------------------------------------------------
'���ܣ���̨����ͼ��
'������ thisMsg  -- Ҫ���ص�ͼ����Ϣ
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    Dim lngWinHandle As Long        '��Ҫ������Ϣ�ġ�����ͼ�����ء�����Ĵ��ھ��
    Dim strMsg As String
    Dim wParam As Long
    Dim lResult As Long
    Dim strTemp As String
    Dim buf(1 To 1024) As Byte
    Dim dss As COPYDATASTRUCT
    
    On Error Resume Next
    
    '��֯��Ϣ
    strMsg = thisMsg.strSubDir & "||" & thisMsg.strDestMainDir & "||" & thisMsg.strIP & "||" _
            & thisMsg.strFtpDir & "||" & thisMsg.strFTPUser & "||" & thisMsg.strFTPPswd & "||" _
            & thisMsg.strSDDir & "||" & thisMsg.strSDUser & "||" & thisMsg.strSDPswd
    
    '����COPYDATA��Ϣ
    
    On Error GoTo err
    
    'ʹ��BUF������ʹ��lstrcpy�������������������ַ���Ϣ
   '��Ϣ���壺wParam = 123��dss��dwData = 3
    wParam = 123
   
    Call CopyMemory(buf(1), ByVal strMsg, LenB(StrConv(strMsg, vbFromUnicode)))
    dss.dwData = 3               '�����Ϣ���ã�3ֻ��˫�������һ����Ƕ���
    dss.cbData = LenB(StrConv(strMsg, vbFromUnicode)) + 1
    
    dss.lpData = VarPtr(buf(1))                    'ʹ��buf���ͣ����Կ�����Ϣ��1024֮��
'    dss.lpData = lstrcpy(strMsg, strMsg)            '����������͵���Ϣ��Ҳ����ȷ�ġ�
    
    
    '����ͼ�����ش���
    Shell App.Path & "\zlGetImage.exe"
        
    '���ش����ʱ�򣬲���ͼ�����س���
    lngWinHandle = FindWindow(vbNullString, "����ͼ������")
    
    lResult = SendMessage(lngWinHandle, WM_COPYDATA, wParam, dss)
    
    funPreDownLoadImages = True
    Exit Function
err:
    '�ݲ�����
End Function

Private Function funPreDownLoadImagesEx(ByVal objParent As Object, thisMsg As TGetImgMsg) As Boolean
'------------------------------------------------
'���ܣ���̨����ͼ��
'������ thisMsg  -- Ҫ���ص�ͼ����Ϣ
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    Dim blnUseActivexLoad As Boolean
    Dim objImgInfo As Object
    
    On Error Resume Next
    
    If mobjGetImage Is Nothing Then
        Set mobjGetImage = DynamicCreate("zlGetImageEx.clsImageTransfer", "zlGetImageEx.exe")
        
        If mobjGetImage Is Nothing Then Exit Function
        
        Call mobjGetImage.RegEventObj(objParent)
        Call mobjGetImage.zlInitModule(False, 0)
    End If
    
    Set objImgInfo = mobjGetImage.ImgInfo
    With objImgInfo
        .SubDir = thisMsg.strSubDir
        .DestMainDir = thisMsg.strDestMainDir
        .IP = thisMsg.strIP
        .FtpDir = thisMsg.strFtpDir
        .FtpUser = thisMsg.strFTPUser
        .FTPPswd = thisMsg.strFTPPswd
        .SDDir = thisMsg.strSDDir
        .SDUser = thisMsg.strSDUser
        .SDPswd = thisMsg.strSDPswd
    End With
    
    Call mobjGetImage.MsgInQueue(objImgInfo)
    Call mobjGetImage.zlLoadImage
    
    Exit Function
err:
    '�ݲ�����
End Function

Private Function OpenViewerWithInXWPacs(ByVal lngAdviceID As Long, ByVal strModalityType As String, ByVal blnMoved As Boolean)
'���°�pacs�д򿪹�Ƭ
    Dim rsTemp As ADODB.Recordset

    Dim strFtpUrl As String
    Dim strImages As String
    
    Set rsTemp = GetStudyImageData(lngAdviceID, blnMoved)
    
    strImages = ""

    While Not rsTemp.EOF
        If Nvl(rsTemp!�豸��1) <> "" Then
            strFtpUrl = "\\" & Nvl(rsTemp!Host1) & "\" & mstrImageShareDir & Nvl(rsTemp!Root1) & Nvl(rsTemp!Url)
        Else
            strFtpUrl = "\\" & Nvl(rsTemp!Host2) & "\" & mstrImageShareDir & Nvl(rsTemp!Root2) & Nvl(rsTemp!Url)
        End If
        
        If strImages <> "" Then strImages = strImages & "[;]"
        
        strFtpUrl = Replace(strFtpUrl, "//", "/")
        strImages = strImages & Replace(strFtpUrl, "/", "\")
        
        rsTemp.MoveNext
    Wend
    
    '��Զ��Ŀ¼�ļ����жԱȹ�Ƭ
    Call OEMViewOpen(0, strImages, 0, strModalityType)
End Function

Private Function GetStudyImageData(ByVal lngAdviceID As Long, ByVal blnMoved As Boolean) As ADODB.Recordset
'��ȡ���ͼ������

    Dim strSql As String
        
    strSql = "Select rownum as ˳���, A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
        "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, " & _
        "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
        "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2," & _
        "e.�豸�� as �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) and c.ҽ��ID=[1] "
        

    If blnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set GetStudyImageData = zlDatabase.OpenSQLRecord(strSql, "��ѯͼ����Ϣ", lngAdviceID)
End Function

'�ж��Ƿ����ù�Ƭ
Private Function IsUseXwViewer() As Boolean
On Error GoTo errHandle
    Dim lngPhkResult As Long
    Dim lngKey As Long
    Dim blnResult As Boolean
    Dim strValue As String
    Dim lngLen As Long
    
    blnResult = IIf(RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Silver\Silver Pacs\General", 0, 1, lngPhkResult) = 0&, True, False)
    BUGEX "0 IsUseXwViewer = " & blnResult, True
    '�����װ��xw viewer����ע����еİ�װ��Ϣ
    If blnResult = False Then Exit Function
    
    strValue = Space(255)
    lngLen = Len(strValue)
    
    Call RegQueryValueEx(lngPhkResult, "CompanyName", 0, 0, strValue, lngLen)
    
    Call RegCloseKey(lngKey)
    
    If InStr(strValue, "����ҽѧӰ��ϵͳ") <= 0 Then blnResult = False
    BUGEX "1 IsUseXwViewer = " & blnResult, True
    IsUseXwViewer = blnResult
Exit Function
errHandle:
    IsUseXwViewer = False
    BUGEX "2 err IsUseXwViewer = " & blnResult & "   err= " & err.Description, True
End Function

Private Function XWADViewerOpen(ByVal strFilter As String, ByVal lngPlanId As Long) As Long
'--------------------------------------------
'���ܣ� ��ָ��ͼ�������ɲ���ָ�����ұ����������ļ��е����������
'       ͼ���ʱ��ADViewer ��ǰģʽ�йأ�����ǵ���¼ģʽ�����������Զ��ر�ԭ����ͼ������ǶԱ�
'       ģʽ��������ӵ�ADViewer�С�
'������
'       lngOrderID -- ҽ��ID
'���أ�
'--------------------------------------------
    Dim strRev As String
    Dim lngFunction As Long
    Dim strXwPrivs As String
    
    'XWViewerOpen ����˵����
    'lPlanID��  ����ID����ID ������INI �ļ���һ�£��ڼ����������£�ͨ����ֵΪ1������Ѹ�ID ��Ϊһ���������ʱ��ȡ������롣
    'cpFilter�� �ô���Ҫ��ͼ�������ֵ��������š�����ŵȣ����Դ�����ֵ��
    '           ��ֵ֮ͬ���÷ָ���[;]�������ò��������弰˳����INI �ļ������ã�������lPlanID��Ӧ��
    'lFunc��    ����Ȩ�ޡ�ÿһλ����һ��ܣ�������ж���Ȩ�ޣ���λ���򡱼��ɣ����幦������:
    '           0x00000002�� �ؽ�ͼ�񱣴棬���磺��Ӱ��ͼ��ƴ��ͼ���
    '           0 x00000200: ��Ƭ��ӡ
    '           0 x00040000: ͼ�񵼳�?����Ϊ������ʽ
    '           0 x00080000: GSPS ����
    'cpReserved��   ��������ΪNULL
    
    On Error GoTo err
    
    '��¼�ӿ���־
    If mblnXWLog = True Then
        Call WriteCommLog("XWADViewerOpen", "XW�ӿ�", "��ADViewer������ʾͼ��ҽ��ID= " & strFilter)
    End If
    
    '����RIS�е�Ȩ�ޣ���֯Ȩ�޴�
    lngFunction = 0
    strXwPrivs = gobjComlib.GetPrivFunc(mlngSys, G_LNG_XWPACSVIEW_MODULE)
    
    If InStr(strXwPrivs, "PACS�����ؽ�ͼ��") <> 0 Then
        lngFunction = lngFunction Or &H2
    End If
    
    If InStr(strXwPrivs, "PACS��Ƭ��ӡ") <> 0 Then
        lngFunction = lngFunction Or &H200
    End If
    
    If InStr(strXwPrivs, "PACSͼ�񵼳�") <> 0 Then
        lngFunction = lngFunction Or &H40000
    End If
    
    If InStr(strXwPrivs, "PACS GSPS����") <> 0 Then
        lngFunction = lngFunction Or &H80000
    End If
        
    XWADViewerOpen = OEMViewOpen(lngPlanId, strFilter, lngFunction, "")
    
    If XWADViewerOpen <> 0 Then
        MsgBox "ADViewer�򿪴��󣬷��ص���Ϣ�ǣ�" & XWADViewerOpen
    End If
    
    Exit Function
err:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

'-------------------------------------------------------------------------------------------------------
'ADViewer�鿴ͼ��Ӧ�ú���
'-------------------------------------------------------------------------------------------------------
Private Function XWShowImage(ByVal lngViewerType As Long, ByVal strFilter As String, Optional ByVal lngPlanId As Long = 1) As Long
''--------------------------------------------
''���ܣ� ��������ADViewer����WEB Viewer
''������    lngViewerType -- ��Viewer�ķ�ʽ��1-�����ADViewer��2-�ٴ�WEB Viewer
'           lngOrderID -- ҽ��ID
''���أ�0-�ɹ�;1-����
''--------------------------------------------
    On Error GoTo err
    
    '��¼�ӿ���־
    If mblnXWLog = True Then
        Call WriteCommLog("XWShowImage", "XW�ӿ�", "����ADViewer����WEB��Ƭ����Ƭ��ʽ�ǣ� " & IIf(lngViewerType = 1, "ADViewer", "WEB"))
    End If
    
    If lngViewerType = 1 Then
        Call XWADViewerOpen(strFilter, lngPlanId)
    ElseIf lngViewerType = 2 Then
        Call XWWebViewerOpen(strFilter)
    End If
    
    Exit Function
err:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function XWWebViewerOpen(ByVal lngOrderID As Long) As Long
''--------------------------------------------
''���ܣ� ��������WEB Viewer
'           lngOrderID -- ҽ��ID
''���أ�0-�ɹ�;1-����
''--------------------------------------------
    Dim strIP As String
    Dim strURL As String
    
    On Error GoTo err
    
    strIP = zlDatabase.GetPara("XWWEB������IP", mlngSys, G_LNG_XWPACSVIEW_MODULE, "")
    
    If strIP <> "" Then
        strURL = "C:\Program Files\Internet Explorer\iexplore.exe http://" & strIP & ":8080/imageweb/imageAction.action?ColID0=22&ColValue0=" & lngOrderID
        
        '��¼�ӿ���־
        If mblnXWLog = True Then
            Call WriteCommLog("XWWebViewerOpen", "XW�ӿ�", "ͨ��WEB��ʽ��Ƭ�� " & strURL)
        End If
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerOpen = 0
    Else
        '��¼�ӿ���־
        If mblnXWLog = True Then
            Call WriteCommLog("XWWebViewerOpen", "XW�ӿ�", "ͨ��WEB��ʽ��Ƭ��WEB������IP��ַΪ�ա�")
        End If
        
        MsgBox "WEB������IP��ַΪ�գ��������ú�WEB��������", vbOKOnly, "��ʾ��Ϣ"
        
        XWWebViewerOpen = 1
    End If
    
    Exit Function
err:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub OnState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean, ByVal lngThreadId As Long)
'blnLoadFinish,ͼ�����������ture:ͼ��������ɣ�false:ͼ����������
'blnUpLoad��true:�ϴ���false:����
'lngThreadId����ǰ����ͼ��Ľ��̱��
    
On Error GoTo err
    'ͼ���ϴ��ɹ����ͷ������
    If blnLoadFinish Then Set mobjGetImage = Nothing
    
    Exit Sub
err:

End Sub

Private Sub Class_Initialize()
    ReDim mConnectedShardDir(0) As String   '��ʼ������Ŀ¼���Ӵ�
End Sub

Private Sub Class_Terminate()
    If Not mobjPacsCore Is Nothing Then
        mobjPacsCore.Closefrom
        Set mobjPacsCore = Nothing
    End If
End Sub