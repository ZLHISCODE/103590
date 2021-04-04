Attribute VB_Name = "mdlGetImage"
Option Explicit

Public plngHookHandle As Long       '��¼��Ϣhook��handle
Public dss As COPYDATASTRUCT        '�����ַ�����Ϣ���ڴ�ṹ
Public pblnQueueBusy As Boolean     '��¼��ǰ�����Ƿ��ڴ�����У�һ��ֻ����һ�����������
Public Const TIME_OUT = 60          '��ʱʱ�����ã�60��
Public pintQueueIndex As Integer    '��¼�����е�һ����Ϣ������
Public pConnectedSharedDir() As String  '��¼�Ѿ����ӹ��Ĺ���Ŀ¼
Public pftpConnect As clsFtp        '����һ�������ӵ�FTP��
Public pblnMsgProcessing As Boolean '���ڴ�����Ϣ

'��־��صĲ�������ע������ж�ȡ��־�����������־·��Ϊ�գ���ʹ��exe��ͬĿ¼�µġ�GetImgLog����Ϊ��־·��
Public pblnLogEnable As Boolean     '�Ƿ�������־
Public pstrLogPath As String        '��¼��־��·��,
Public plngLogLevel As String       '��¼��־�ļ��𣬷ֳ�1,2������1��ֻ��¼��Ϣ�������־��2����¼ÿһ�����ص���־


'������Ϣ���ݵĽṹ
Public Type TGetImgMsg
    strSubDir As String          'ͼ�����ڵ���Ŀ¼
    strDestMainDir As String            '����ͼ���Ŀ��Ŀ¼������Ŀ¼
    strIP As String                 'ͼ���������IP��ַ
    strFTPDir As String             'FTPĿ¼
    strFTPUser As String            'FTP�û���
    strFTPPswd As String            'FTP����
    strSDDir As String              '����Ŀ¼����
    strSDUser As String             '����Ŀ¼�û���
    strSDPswd As String             '����Ŀ¼����
    blnEnable As Boolean            '����Ϣ����
End Type
Public curMsg As TGetImgMsg


'����ͼ�����Ϣ����
Public pstrMsgQueue() As String          '��Ϣ���ݶ���

'��ϢHook����
Public plngPreWndProc As Long       'ԭ������Ϣ�������


Public Function MsgInQueue(strMsg As String) As Boolean
'------------------------------------------------
'���ܣ���Ϣ��ӣ��Ⱥ򱻴���
'������ strMsg ������Ҫ��ӵ���Ϣ���ݣ���Ϣ�����ɴ洢��ַ���洢�û������洢���룬�洢Ŀ¼������Ŀ¼���洢�������
'���أ�True������ӳɹ���False�������ʧ��
'-----------------------------------------------
    Dim Timer As Long
    Dim intCount As Integer
    
    On Error GoTo err
    
    MsgInQueue = False
    
    If pblnQueueBusy = True Then
        '�������æ�����ڽ��ж��д�����ȴ�������ȴ���ʱ������Ϣ�ŵ����ö���
        
        
    Else
        '������п��У��������Ӵ������ұ�Ƕ���æ
        pblnQueueBusy = True
        
        '������Ϣ���
        intCount = UBound(pstrMsgQueue) + 1
        ReDim Preserve pstrMsgQueue(intCount) As String
        pstrMsgQueue(intCount) = strMsg
        
        '���д�����ɣ���Ƕ���Ϊ��
        pblnQueueBusy = False
    End If
    MsgInQueue = True
    Exit Function
err:
    '������˳�����ʱ��������
End Function

Public Function MsgOutQueue() As String
'------------------------------------------------
'���ܣ���Ϣ���ӣ����ù��̸�������ӵ���Ϣ
'������
'���أ����س��ӵ���Ϣ����
'-----------------------------------------------
    Dim iCount As Integer
    Dim strMsg As String
    
    On Error GoTo err
    MsgOutQueue = ""        '��ʼ��Ϊ����Ϣ
    
    If pblnQueueBusy = True Then
        '�������æ�����ڽ��ж��д�����ȴ�������ȴ���ʱ���˳����Ӳ���
        
    Else
        '������п��У�����г��Ӵ������ұ�Ƕ���æ
        pblnQueueBusy = True
        
        '��Ϣ���Ӵ���
        iCount = UBound(pstrMsgQueue)
        If iCount = 0 Then
            pblnQueueBusy = False
            Exit Function     '����Ϊ�գ����ó���
        End If
        
        '�Ӷ�������ȡ��Ϣ
        strMsg = pstrMsgQueue(pintQueueIndex)
        
        '�������ָ��
        pintQueueIndex = pintQueueIndex + 1
        If pintQueueIndex > iCount Then
            '�����ǰȡ�������Ƕ����е����һ����Ϣ���������Ϣ����
            ReDim Preserve pstrMsgQueue(0) As String
            pintQueueIndex = 1
        End If
        
        MsgOutQueue = strMsg
        
        '���Ӵ�����ɣ���Ƕ�����
        pblnQueueBusy = False
    End If
    
    Exit Function
err:
    '������˳�����ʱ��������
End Function


Public Function funConnectAndSaveSheardDir(strSharedDir As String, strUser As String, strPswd As String) As Boolean
'------------------------------------------------
'���ܣ����ӹ���Ŀ¼��ʹ���û��������ܵ�¼������
'������ strSharedDir -- ��Ҫ���ӵĹ���Ŀ¼����

'���أ�True-- �ɹ���False -- ʧ��
'-----------------------------------------------
    Dim i As Integer
    
    funConnectAndSaveSheardDir = False

    On Error GoTo err
    
    '�жϹ���Ŀ¼�Ƿ��Ѿ����ӣ����û�����ӣ����������
    For i = 1 To UBound(pConnectedSharedDir)
        If pConnectedSharedDir(i) = strSharedDir Then
            funConnectAndSaveSheardDir = True
            Exit Function
        End If
    Next i
    
    '���ӹ���Ŀ¼
    If strSharedDir <> "" Then
        If funConnectSharedDir(strSharedDir, strUser, strPswd) = True Then
            '���ӳɹ�����¼�ɹ������Ӵ�
            ReDim Preserve pConnectedSharedDir(UBound(pConnectedSharedDir) + 1) As String
            pConnectedSharedDir(UBound(pConnectedSharedDir)) = strSharedDir
            funConnectAndSaveSheardDir = True
        End If
    End If
    
    Exit Function
err:
    '�ݲ�����
End Function


Public Function funConnectSharedDir(strShareRemoteDir As String, strUserName As String, _
    strPassWord As String) As Boolean
'------------------------------------------------
'���ܣ�����������Դ
'������ strShareRemoteDir -- ����Ŀ¼
'       strUserName -- ����Ŀ¼�û���
'       strPassWord -- ����Ŀ¼����
'���أ�True--���ӳɹ��� False -- ����ʧ��
'------------------------------------------------
    
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    On Error GoTo err
    
    funConnectSharedDir = False
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
   ' If lngResult = 0 Then
        funConnectSharedDir = True
   ' End If
    Exit Function
err:
    '�ݲ�����
End Function


Public Function funDownLoadSharedDir(strSourceDir As String, strDestDir As String) As Boolean
'------------------------------------------------
'���ܣ�ͨ������Ŀ¼�ķ�ʽ����һ��Ŀ¼��ͼ��strSourceDirĿ¼���ڵļ������¼��ͨ������Ĺ���ʵ��
'������ strSourceDir -- ��Ҫ�����ļ���ԴĿ¼����Զ�̷������еĹ���Ŀ¼
'       strDestDir  --  �ļ����Ƶ�Ŀ�ĵأ���������ı���Ŀ¼
'���أ�True-- �ɹ���False -- ʧ��
'-----------------------------------------------
    Dim fs As New Scripting.FileSystemObject
    Dim fsFolder As Folder
    Dim fsFiles As Files
    Dim fsFile As File
    
    On Error Resume Next
    
    funDownLoadSharedDir = False
    
    '���ԴĿ¼�����ڣ����˳�
    If Dir(strSourceDir, vbDirectory) = "" Then
        '��¼��־
        Call WriteCommLog("funDownLoadSharedDir", "ͨ������Ŀ¼��ʽ����ͼ��", "ԴĿ¼������" & strSourceDir, 1)
        Exit Function
    End If
    
    '���Ŀ��Ŀ¼�����ڣ��򴴽�Ŀ¼
    If Dir(strDestDir, vbDirectory) = "" Then
        Call MkLocalDir(strDestDir)
    End If
    
    '����Ŀ¼�е������ļ���һ�������أ�����ȷ���������Ŀ¼���Ѿ�����ĳ���ļ����������ļ���������������
    Set fsFolder = fs.GetFolder(strSourceDir)
    Set fsFiles = fsFolder.Files
                
    '��ֹOverWrite����ֹ��FTP�е��ļ����Ǳ���Ŀ¼���Ѿ����˵��ļ���
    For Each fsFile In fsFiles
        '��¼��־
        Call WriteCommLog("funDownLoadSharedDir", "����ͼ��", "����ͼ�� " & strDestDir & "\" & fsFile.Name, 2)
        
        Call fs.CopyFile(strSourceDir & "\" & fsFile.Name, strDestDir & "\" & fsFile.Name, False)
    Next fsFile
    
    funDownLoadSharedDir = True

End Function


Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function funGetAMessage() As Boolean
'------------------------------------------------
'���ܣ�����Ϣ������ȡһ����Ϣ�������Ϣ�ṹ
'������
'���أ�True -- �ɹ� �� False -- ʧ��
'-----------------------------------------------
    Dim strMsg As String
    Dim strOldMsg As String
    Dim strMsgArr() As String
    
    funGetAMessage = False
    
    On Error GoTo err
    
    '��¼��һ����Ϣ�����������ͬ��������Ϣ���򲻴�������Ϣ
    strOldMsg = strMsg
    
    '����Ϣ������ȡһ����Ϣ
    strMsg = MsgOutQueue
    
    '�����Ϣ��Ϊ�գ����������Ϣ
    If strMsg <> "" And strOldMsg <> strMsg Then
        '������Ϣ
        strMsgArr = Split(strMsg, "||")
        'ֻ���������9���ε���Ϣ���еĶο���Ϊ��
        If UBound(strMsgArr) = 8 Then
            curMsg.strSubDir = strMsgArr(0)
            curMsg.strDestMainDir = strMsgArr(1)
            curMsg.strIP = strMsgArr(2)
            curMsg.strFTPDir = strMsgArr(3)
            curMsg.strFTPUser = strMsgArr(4)
            curMsg.strFTPPswd = strMsgArr(5)
            curMsg.strSDDir = strMsgArr(6)
            curMsg.strSDUser = strMsgArr(7)
            curMsg.strSDPswd = strMsgArr(8)
            curMsg.blnEnable = True
            
            '����·����ȡ��·���е�һ�������һ����\�����ߡ�/������
            curMsg.strSubDir = funRemoveSlash(curMsg.strSubDir)
            curMsg.strDestMainDir = funRemoveSlash(curMsg.strDestMainDir)
            curMsg.strFTPDir = funRemoveSlash(curMsg.strFTPDir)
            
            '��Ϣ�������Ӳ�д����Ϣ�ṹ��������Ϣ������ɱ��
            funGetAMessage = True
        Else
            curMsg.blnEnable = False
        End If
    End If
    
    Exit Function
err:
    '�ݲ�����
End Function

Public Function funRemoveSlash(strDir As String) As String
'------------------------------------------------
'���ܣ�ȥ��·���еĵ�һ�������һ��б�ߣ����߷�б��
'������ strDir  -- ��Ҫ�����·��
'���أ�����֮���·��
'-----------------------------------------------
    Dim strTemp As String
    
    strTemp = strDir
    funRemoveSlash = strTemp
    
    On Error GoTo err
    
    'ȥ��·���󲿵�б��
    If Right(strTemp, 1) = "/" Or Right(strTemp, 1) = "\" Then
        strTemp = Left(strTemp, Len(strTemp) - 1)
    End If
    
    'ȥ��·��ǰ��б��
    If Left(strTemp, 1) = "/" Or Left(strTemp, 1) = "\" Then
        strTemp = Mid(strTemp, 2)
    End If
    
    funRemoveSlash = strTemp
    
    Exit Function
err:
    '�ݲ�����,������ֱ�ӷ���δ������ַ���
End Function


Public Function funDownLoadFTP(thisMsg As TGetImgMsg) As Boolean
'------------------------------------------------
'���ܣ���������ͼ�����Ϣ����FTP������ָ��Ŀ¼�е�ȫ��ͼ��
'������ thisMsg  -- ��Ҫ����ͼ�����Ϣ
'���أ�True -- �ɹ��� False -- ʧ��
'-----------------------------------------------
    Dim lngResult As String
    
    funDownLoadFTP = False
    On Error GoTo err
    
    '�ȼ����Ϣ�Ƿ����
    If thisMsg.blnEnable = False Or thisMsg.strIP = "" Then
        
        '��¼��־
        Call WriteCommLog("funDownLoadFTP", "��Ϣ������", "FTP��ʽ����ͼ����Ϣ�����û���IP��ַΪ�գ��޷����ء� IP��ַ�ǣ�" & thisMsg.strIP, 1)
        
        Exit Function
    End If
    
    '����FTP
    '�жϵ�ǰ��Ϣ�е�FTP���Ӹ����е�FTP�����Ƿ���ͬ
    If thisMsg.strIP = pftpConnect.IPAddress And thisMsg.strFTPUser = pftpConnect.User _
        And thisMsg.strFTPPswd = pftpConnect.PassWord Then
        '����Ҫ��������FTP
    Else
        '��������FTP
        Call pftpConnect.FuncFtpDisConnect
        lngResult = pftpConnect.FuncFtpConnect(thisMsg.strIP, thisMsg.strFTPUser, thisMsg.strFTPPswd)
        
        '�������ʧ�ܣ��˳�����
        If lngResult = 0 Then
            '��¼��־
            Call WriteCommLog("funDownLoadFTP", "FTP����ʧ��", "FTP����ʧ�ܣ� " & thisMsg.strIP, 1)
            
            Exit Function
        End If
    End If
    
    '��������·��
    Call MkLocalDir(thisMsg.strDestMainDir & "\" & thisMsg.strSubDir)
    
    '��¼��־
    Call WriteCommLog("funDownLoadFTP", "����ͼ��", "ͨ��FTP����ȫ��ͼ�� ", 1)
    
    '����ͼ��
    Call pftpConnect.funcDownLoadAllFiles("\" & thisMsg.strFTPDir & "\" & thisMsg.strSubDir, thisMsg.strDestMainDir & "\" & thisMsg.strSubDir, False)
    
    funDownLoadFTP = True
    Exit Function
err:
    '�ݲ�����
End Function

Public Function funMsgProcess() As Boolean
'------------------------------------------------
'���ܣ��Զ�������Ϣ
'������ ��
'���أ�True -- �ɹ��� False -- ʧ��
'-----------------------------------------------
    
    On Error GoTo err
    
    funMsgProcess = False
    
    '����Ѿ�������Ϣ�����ˣ����˳�
    If pblnMsgProcessing = True Then Exit Function
    
    '������Ϣ�����ǣ���ֹ������̱���ε���
    pblnMsgProcessing = True
    
    '��Ϣ���Ӳ�������Ϣ
    While funGetAMessage = True
        '��¼��־
        Call WriteCommLog("funMsgProcess", "���յ���������Ϣ", "׼�����ش�Ŀ¼�µ�ͼ��" & curMsg.strSubDir, 1)

        '������Ϣ
        Call funDownLoadImages(curMsg)
    Wend
    
    '��Ϣ������ɣ��˳�����
    pblnMsgProcessing = False
    
    funMsgProcess = True
    Exit Function
err:
    '�ݲ�����
End Function

Public Function funDownLoadImages(thisMsg As TGetImgMsg) As Boolean
'------------------------------------------------
'���ܣ���������ͼ�����Ϣ���ӹ���Ŀ¼����FTP����ͼ��
'������ thisMsg  -- ��Ҫ����ͼ�����Ϣ
'���أ�True -- �ɹ��� False -- ʧ��
'-----------------------------------------------
Dim blnResult As Boolean
    
    On Error GoTo err
    
    '�����ǰ��Ϣ���ã�������Ϣ
    If thisMsg.blnEnable = True Then
        '�й���Ŀ¼����ʹ�ù���Ŀ¼����ͼ��û����ʹ��FTP����ͼ��
        If thisMsg.strSDDir <> "" Then
            '���ӹ���Ŀ¼
            If funConnectAndSaveSheardDir("\\" & thisMsg.strIP & "\" & thisMsg.strSDDir, thisMsg.strSDUser, thisMsg.strSDPswd) = True Then
                '��¼��־
                Call WriteCommLog("funDownLoadImages", "����Ŀ¼��ʽ����ͼ��", "ͨ������Ŀ¼��ʽ����ͼ�񣬴�Ŀ¼�� " & "\\" & thisMsg.strIP & "\" & thisMsg.strSDDir & "\" & thisMsg.strSubDir & " �����ص���Ŀ¼�� " & thisMsg.strDestMainDir & "\" & thisMsg.strSubDir, 1)
                
                '����ͼ��
                blnResult = funDownLoadSharedDir("\\" & thisMsg.strIP & "\" & thisMsg.strSDDir & "\" & thisMsg.strSubDir, thisMsg.strDestMainDir & "\" & thisMsg.strSubDir)
            End If
        Else
            '��¼��־
            Call WriteCommLog("funDownLoadImages", "FTP��ʽ����ͼ��", "ͨ��FTP��ʽ����ͼ�񣬴�Ŀ¼�� \" & thisMsg.strFTPDir & "\" & thisMsg.strSubDir & " �����ص���Ŀ¼�� " & thisMsg.strDestMainDir & "\" & thisMsg.strSubDir, 1)
                
            'ʹ��FTP����
            blnResult = funDownLoadFTP(thisMsg)
        End If
    End If
    
    funDownLoadImages = blnResult
    Exit Function
err:
    '�ݲ�����

End Function

Public Sub Main()
'------------------------------------------------
'���ܣ������򣬸�������ͼ�����س���
'������
'���أ���
'-----------------------------------------------
    Dim strRegPath As String
    
    '����������Ѿ�������һ�Σ���������
    If App.PrevInstance Then
        Exit Sub
    End If
    
    
    
    On Error Resume Next
    
    '��ע����ȡ��־����
    strRegPath = "����ģ��\zlPacsGetImage"
    pblnLogEnable = (Val(GetSetting("ZLSOFT", strRegPath, "��¼��־", 0)) = 1)
    pstrLogPath = GetSetting("ZLSOFT", strRegPath, "��־·��", "")
    plngLogLevel = Val(GetSetting("ZLSOFT", strRegPath, "��־����", 1))
    '�����������־�������־·���Ƿ����
    If pblnLogEnable = True Then
    
        '���û��������־·������ʹ��Ĭ��·��
        If pstrLogPath = "" Then
            pstrLogPath = App.Path & "\GetImgLog"
        End If
        
        '�����־·�������ڣ��򴴽�
        If Dir(pstrLogPath, vbDirectory) = "" Then
            'Ĭ��·�������ڣ��������Ŀ¼
            If Dir(pstrLogPath, vbDirectory) = "" Then
                Call MkLocalDir(pstrLogPath)
            End If
        End If
    End If
    
    '��һ���������򣬼��ش��壬������
    frmMain.Show
    frmMain.WindowState = vbMinimized
    frmMain.Hide
    
End Sub


Private Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String, lngLogLevel As Long)
'------------------------------------------------
'���ܣ���¼ͨѶ��־
'������ logSubName  --  ������־�ĺ�����
'       logTitle   -- ��־����
'       logDesc   --  ��־����
'       lngLogLevel -- ��־����ͨ����־����ȷ����ǰ��־�Ƿ���Ҫ��¼
'���أ���
'------------------------------------------------
    Dim strLog As String
    Dim strFileName As String
    Dim intHour As Integer
    
    On Error GoTo err
    
    If pblnLogEnable = True Then        '�����˼�¼��־���ż�¼��ǰ����־
        '�ж���־����ȷ��������־�Ƿ���Ҫ��¼
        If plngLogLevel >= lngLogLevel Then
            'ͨ����ǰʱ�䣬������־�ļ�����ÿ����Сʱ����һ����־�ļ�
            intHour = Hour(Time)
            intHour = intHour / 2
            intHour = intHour * 2
            strFileName = pstrLogPath & "\" & Date & "-" & intHour & ".log"
            
            '������־����
            strLog = Now() & " ��־���� " & lngLogLevel & " ���⣺ " & logTitle & vbCrLf & "      ������ " & logSubName & vbCrLf & "     ��־���ݣ�" & logDesc & vbCrLf
            
            '����־�ļ�����¼��־
            Open strFileName For Append As #1
            Print #1, strLog
            Close #1
        End If
    End If
    Exit Sub
err:
    Close #1
End Sub
