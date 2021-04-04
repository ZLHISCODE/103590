VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInetFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjInet As InetCtlsObjects.iNet
Attribute mobjInet.VB_VarHelpID = -1

Private mstrData As String
Private mstrFtpVirtualPath As String
Private mblnExecuteState As Boolean
Private mlngLastState As Long

Private mobjAryInet() As InetCtlsObjects.iNet

'Public Event OnStateChanged(ByVal State As Integer)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



'ע�����ʹ����CERN����������������ʹ��execute���첽ִ�з�ʽ



'���ִ��״̬
Property Get LastState() As Long
    LastState = mlngLastState
End Property


'inet�Ƿ��ڿ���״̬
Property Get FreeState() As Boolean
    FreeState = Not mobjInet.StillExecuting
End Property


'��ȡ��ǰftp������Ŀ¼
Property Get FtpVirtualPath() As String
    FtpVirtualPath = mstrFtpVirtualPath
End Property


'��ȡ�Ƿ�����inet����
Property Get IsConnectionInet() As Boolean
    IsConnectionInet = IIf(mobjInet Is Nothing, False, True)
End Property





Private Function ExecuteFtpCMD(ByVal strCMD As String) As Boolean
'ִ��ftp����
    ExecuteFtpCMD = False
    
    mblnExecuteState = True
    mstrData = ""
    
    '�����ǰ������δ��������ֱ���˳�����
    If mlngLastState = icReceivingResponse Then Exit Function
    If mobjInet.StillExecuting Then Exit Function
    
    Call mobjInet.Execute(, strCMD)
    
    Call WaitExecute
    
    If mblnExecuteState Then ExecuteFtpCMD = True
End Function


Private Sub WaitExecute()
'�ȴ�ִ��
    While mobjInet.StillExecuting
'        Call Sleep(1)
        
        DoEvents
    Wend
End Sub


Public Function FormatPath(ByVal strPath As String, ByVal blnIsFtpDir As Boolean) As String
'��ʽ��·��
    Dim strCurPath As String
    Dim strPathChar As String
    
    FormatPath = ""
    If Trim(strPath) = "" Then Exit Function
    
    strPathChar = "\"
    If blnIsFtpDir Then strPathChar = "/"
    
    strCurPath = strPath
    
    '��ʽ������Ŀ¼��������Ŀ¼ǰʹ����"/\VirtualPath\"����ʽʱ�������ܴ�ftp�ϻ�ȡ�ļ���Ϣ
    If blnIsFtpDir Then
        strCurPath = Replace(strCurPath, "\", strPathChar)
    Else
        strCurPath = Replace(strCurPath, "/", strPathChar)
    End If
    
    strCurPath = Replace(strCurPath, "\\", strPathChar)
    strCurPath = Replace(strCurPath, "//", strPathChar)

    If InStrRev(strCurPath, strPathChar) <> Len(strCurPath) Then strCurPath = strCurPath & strPathChar
    
    FormatPath = strCurPath
End Function


Public Function GetFilePath(ByVal strFile As String, ByVal blnFtpFile As Boolean) As String
'��ȡ�ļ���·��
    Dim strCurPath As String
    Dim strPathChar As String
    
    GetFilePath = ""
    If Trim(strFile) = "" Then Exit Function
    
    If blnFtpFile Then
        strPathChar = "/"
        strCurPath = Replace(strFile, "\", "/")
    Else
        strPathChar = "\"
        strCurPath = Replace(strFile, "/", "\")
    End If
    
    '���������·������ֱ���˳�
    If InStr(strCurPath, strPathChar) <= 0 Then Exit Function
    
    GetFilePath = Mid(strCurPath, 1, InStrRev(strCurPath, strPathChar))
    
End Function


Public Function GetFileName(ByVal strFlie As String, ByVal blnFtpFile As Boolean) As String
'��ȡ�ļ���
    Dim strCurFile As String
    
    GetFileName = ""
    If Trim(strFlie) = "" Then Exit Function
    
    strCurFile = strFlie
    
    If blnFtpFile Then
        strCurFile = Replace(strCurFile, "\", "/")
    Else
        strCurFile = Replace(strCurFile, "/", "\")
    End If
    
    If InStr(strFlie, "/") > 0 Then
        GetFileName = Mid(strCurFile, InStrRev(strCurFile, "/") + 1, 255)
    ElseIf InStr(strFlie, "\") > 0 Then
        GetFileName = Mid(strCurFile, InStrRev(strCurFile, "\") + 1, 255)
    Else
        GetFileName = strCurFile
    End If
    
End Function











Public Sub ConnectionInet(objInet As Object)
'����Microsoft Internet Transfer control��Inet����

    Set mobjInet = objInet
    
    mobjInet.Protocol = icFTP
    
End Sub


Public Sub InitFtp(ByVal strHost As String, ByVal lngPort As Long, _
    ByVal strUser As String, ByVal strPwd As String, Optional ByVal strVirtualPath As String = "/")
'��ʼ��ftp������Ϣ
    mobjInet.RemoteHost = strHost
    mobjInet.RemotePort = lngPort
    mobjInet.UserName = strUser
    mobjInet.Password = strPwd
    
    mstrFtpVirtualPath = FormatPath(strVirtualPath, True)
    
    '�л���ָ��������Ŀ¼��
    If mstrFtpVirtualPath <> "/" Then Call ChangeDir(mstrFtpVirtualPath)
End Sub


Public Function OpenTest() As Boolean
'�򿪲���
'ʹ��PWD�����ж��Ƿ��ܹ�����
'˵������ftp����
    
    OpenTest = ExecuteFtpCMD("PWD")

End Function


Public Function ExecuteFtpCommand(ByVal strFtpCMD As String, ByRef strResult As String) As Boolean
'ִ��ftp����
    Dim blnResult As Boolean
    
    ExecuteFtpCommand = False
    
    blnResult = ExecuteFtpCMD(strFtpCMD)
    
    If Not blnResult Then Exit Function
    
    strResult = mstrData
    ExecuteFtpCommand = blnResult
End Function

'Public Sub OpenFtp()
''��ftp����
''OPEN����
''��ʽ��OPEN host[port]
''˵������ftp����
'
'    Call ExecuteFtpCMD("OPEN " & mobjInet.RemoteHost & " " & mobjInet.RemotePort)
'End Sub



Public Function QuitFtp() As Boolean
'�ر�ftp
'QUIT����
'��ʽ��QUIT
'˵��������FTP����,������ftp
'Terminates the current user

    QuitFtp = ExecuteFtpCMD("QUIT")
End Function


Public Function CloseFtp() As Boolean
'�ر�ftp
'CLOSE����
'��ʽ��CLOSE
'˵�����رպ�Զ�����������ӣ�������ftp
'Closes the current FTP connection.

    CloseFtp = ExecuteFtpCMD("CLOSE")
End Function


Public Sub Cancel()
'��������
    Call mobjInet.Cancel
    
    Call WaitExecute
End Sub

Public Function ChangeDir(ByVal strFtpDir As String) As Boolean
'�ı�ftpĿ¼
'CD����
'��ʽ��CD file1
'˵�����ı�ftp��ǰĿ¼��file1��ָ����Ŀ¼��
    Dim strCurPath As String
    Dim blnResult As Boolean
    
    ChangeDir = True
    strCurPath = FormatPath(strFtpDir, True)
    
    '�������Ŀ¼��ͬ���򲻽����л�
    If mstrFtpVirtualPath = strCurPath Then Exit Function
    
    blnResult = ExecuteFtpCMD("CD " & strCurPath)
    If blnResult Then
        'ִ�гɹ����������������Ŀ¼
        mstrFtpVirtualPath = strCurPath
    End If
    
    ChangeDir = blnResult
End Function


Public Function ChangeParentDir() As Boolean
'�л����ϼ�Ŀ¼
'CDUP����
'��ʽ��CDUP
'˵�����л�����ǰftpĿ¼���ϼ�Ŀ¼
    Dim blnResult As Boolean
    
    ChangeParentDir = True
    If mstrFtpVirtualPath = "/" Then Exit Function
    
    blnResult = ExecuteFtpCMD("CD ..")
    
    If blnResult Then
        mstrFtpVirtualPath = Mid(mstrFtpVirtualPath, 1, Len(mstrFtpVirtualPath) - 1)
        
        If InStr(mstrFtpVirtualPath, "/") > 0 Then
            mstrFtpVirtualPath = Mid(mstrFtpVirtualPath, 1, InStrRev(mstrFtpVirtualPath, "/"))
        Else
            mstrFtpVirtualPath = "/"
        End If
    End If
    
    ChangeParentDir = blnResult
End Function


Public Function ChangeRootDir() As Boolean
'�л����ϼ�Ŀ¼
'CDUP����
'��ʽ��CDUP
'˵�����л�����ǰftpĿ¼���ϼ�Ŀ¼
    Dim blnResult As Boolean
    
    ChangeRootDir = True
    If mstrFtpVirtualPath = "/" Then Exit Function
    
    blnResult = ExecuteFtpCMD("CD /")
    
    If blnResult Then
        mstrFtpVirtualPath = "/"
    End If
    
    ChangeRootDir = blnResult
End Function


Public Function DelFile(ByVal strFtpFile As String) As Boolean
'ɾ���ļ�
'strFtpFile��Ҫɾ����ftp�ļ������ָ�����ļ�·������ɾ����Ӧ·���µ��ļ�������Ĭ��ɾ����ǰftp·���µ��ļ�

'DELETE����
'��ʽ��DELETE file1
'˵����ɾ����file1ָ�����ļ�
    Dim strCurFile As String
    
    strCurFile = GetFilePath(strFtpFile, True)
    strCurFile = FormatPath(strCurFile, True) & GetFileName(strFtpFile, True)
    
    DelFile = ExecuteFtpCMD("DELETE " & strCurFile)
End Function



Public Function UpLoadFile(ByVal strLocalFile As String, ByVal strFtpPath As String, _
    Optional ByVal strFtpFileName As String = "") As Boolean
'�ϴ��ļ�
'strLocalFile��     ���ϴ����ļ�ȫ·����
'strFtpPath��       ftpĿ¼
'strFtpFileName��   ftp�ļ���

'PUT����
'��ʽ��PUT file1 file2
'˵��������file1ָ�����ļ������Ƶ�file2ָ����Զ�������ļ���
    Dim strCurFileName As String
    
    strCurFileName = strFtpFileName
    If strCurFileName = "" Then
        '��ȡ�����ļ�����
        strCurFileName = GetFileName(strLocalFile, False)
    End If
    
    '��֯ftp�Ĵ��λ��
    strCurFileName = FormatPath(strFtpPath, True) & strCurFileName

    'ִ���ļ��ϴ�
    UpLoadFile = ExecuteFtpCMD("PUT " & strLocalFile & " " & strCurFileName)
End Function

Public Function DownLoadFile(ByVal strFtpFile As String, ByVal strLocalPath As String, _
    Optional ByVal strLocalFileName As String = "") As Boolean
'�����ļ�
'GET����
'��ʽ��GET file1 file2
'˵������ftp��������file1ָ�����ļ���file2��
    Dim strCurFileName As String
    
    strCurFileName = strLocalFileName
    If strCurFileName = "" Then
        '��ȡftp���ļ�����
        strCurFileName = GetFileName(strFtpFile, True)
    End If
    
    '��֯�����ļ��Ĵ��λ��
    strCurFileName = FormatPath(strLocalPath, False) & strCurFileName
    
    DownLoadFile = ExecuteFtpCMD("GET " & strFtpFile & " " & strCurFileName)
End Function


Public Function BatchDownLoad(ByVal strFtpPath As String, ByVal strLocalPath As String) As Boolean
'��������Ŀ¼�е��ļ�
    Dim strCurPath As String
    Dim strFileData As String
    Dim aryFile() As String
    Dim i As Long
    Dim blnResult As Boolean
    Dim strFormatLocalPath As String
    
    BatchDownLoad = False
    If Trim(strFtpPath) = "" Then Exit Function
    
    strCurPath = FormatPath(strFtpPath, True)
    strFileData = GetFtpFileList(strCurPath)
    If strFileData = "" Then Exit Function
    
    strFormatLocalPath = FormatPath(strLocalPath, False)
    aryFile = Split(strFileData, ";")
    
    blnResult = True
    'ѭ������ftpĿ¼�е��ļ�
    For i = LBound(aryFile) To UBound(aryFile)
        strFileData = aryFile(i)
        If Trim(strFileData) <> "" Then
            If Not ExecuteFtpCMD("GET " & strCurPath & strFileData & " " & strFormatLocalPath & GetFileName(strFileData, True)) Then blnResult = False
        End If
    Next i
    
    BatchDownLoad = blnResult
End Function



Public Function GetFtpList(Optional ByVal strFtpPath As String = "") As String
'��ȡftp�ļ��б�
'DIR����
'��ʽ��DIR /
'˵������file1ָ����Ŀ¼�����ļ�,���ص�����ͨ��GetChunk������ȡ
    Dim strCurPath As String
    
    GetFtpList = ""
    strCurPath = FormatPath(strFtpPath, True)
    
    '���ִ��ʧ�ܣ���ֱ���˳�
    If Not ExecuteFtpCMD("DIR " & IIf(strCurPath = "", mstrFtpVirtualPath, strCurPath)) Then Exit Function
    
    If mstrData = "" Then Exit Function
    
    GetFtpList = ";" & Replace(mstrData, vbCrLf, ";")
End Function


Public Function GetFtpDirList(Optional ByVal strFtpPath As String = "") As String
'��ȡftpĿ¼�б�
    Dim strList As String
    Dim aryList() As String
    Dim i As Long
    
    GetFtpDirList = ""
    
    strList = GetFtpList(strFtpPath)
    If strList = "" Then Exit Function
    
    aryList = Split(strList, ";")
    
    For i = LBound(aryList) To UBound(aryList)
        If aryList(i) <> "" Then
            If Mid(aryList(i), Len(aryList(i)), 1) = "/" Then GetFtpDirList = GetFtpDirList & ";" & aryList(i)
        End If
    Next i
    
    If GetFtpList <> "" Then GetFtpDirList = GetFtpDirList & ";"
    
End Function

Public Function GetFtpFileList(Optional ByVal strFtpPath As String = "") As String
'��ȡftp�ļ��б�
    Dim strList As String
    Dim aryList() As String
    Dim i As Long
    
    GetFtpFileList = ""
    
    strList = GetFtpList(strFtpPath)
    If strList = "" Then Exit Function
    
    aryList = Split(strList, ";")
    
    For i = LBound(aryList) To UBound(aryList)
        If aryList(i) <> "" Then
            If Mid(aryList(i), Len(aryList(i)), 1) <> "/" Then GetFtpFileList = GetFtpFileList & ";" & aryList(i)
        End If
    Next i
    
    If GetFtpFileList <> "" Then GetFtpFileList = GetFtpFileList & ";"
End Function


Public Function SerachFile(ByVal strFtpFileName As String, Optional ByVal strSerachPath As String = "") As Boolean
'�����ļ�
'strSerachPath:��ָ����strSerachPath ·���м����ļ������û��ָ������ʹ�õ�ǰftp·�����м���

'������ڣ��򷵻�true
    Dim strList As String
    
    SerachFile = False
    
    strList = GetFtpList(strSerachPath)
    If strList = "" Then Exit Function
    
    If InStr(strList, ";" & strFtpFileName & ";") > 0 Then SerachFile = True
End Function



Public Function RenameFile(ByVal strOldFile As String, ByVal strNewName As String) As Boolean
'����Ŀ¼
'RENAME����
'��ʽ��RENAME file1 file2
'˵�����ļ��������������Ƿ��ܹ��ɹ�ִ�У�ȡ�����û���Զ�������ϵ�Ȩ��
    Dim strFormatOldFile As String
    Dim strFormatNewFile As String
    
    RenameFile = False
    If Trim(strOldFile) = "" Or Trim(strNewName) = "" Then Exit Function
    
    strFormatOldFile = GetFilePath(strOldFile, True)
    strFormatNewFile = FormatPath(strFormatOldFile, True) & strNewName
    strFormatOldFile = FormatPath(strFormatOldFile, True) & GetFileName(strOldFile, True)

    RenameFile = ExecuteFtpCMD("RENAME " & strFormatOldFile & " " & strFormatNewFile)
End Function


Public Function MoveFile(ByVal strOldFile As String, ByVal strNewFile As String) As Boolean
'�ƶ��ļ�
'�ƶ��ļ�ʱ��ʹ��RENAME����
    Dim strFormatOldFile As String
    Dim strFormatNewFile As String
    
    MoveFile = False
    If Trim(strOldFile) = "" Or Trim(strNewFile) = "" Then Exit Function
    
    strFormatOldFile = GetFilePath(strOldFile, True)
    strFormatNewFile = GetFilePath(strNewFile, True)
    
    strFormatOldFile = FormatPath(strFormatOldFile, True) & GetFileName(strOldFile, True)
    strFormatNewFile = FormatPath(strFormatNewFile, True) & GetFileName(strNewFile, True)
    
    MoveFile = ExecuteFtpCMD("RENAME " & strFormatOldFile & " " & strFormatNewFile)
End Function



Public Function CreateDir(ByVal strNewDir As String, Optional ByVal strFtpDir As String = "") As Boolean
'����Ŀ¼
'MKDIR����
'��ʽ��MKDIR file1
'˵�����ļ��������������Ƿ��ܹ��ɹ�ִ�У�ȡ�����û���Զ�������ϵ�Ȩ��
    Dim strCurPath As String
    Dim aryPath() As String
    Dim strCreatedDir As String
    Dim i As Long
    Dim blnResult As Boolean
    
    CreateDir = False
    If strNewDir = "" Then Exit Function
    
    strCurPath = FormatPath(strNewDir, True)
    aryPath = Split(strCurPath, "/")
    strCreatedDir = FormatPath(strFtpDir, True)
    
    For i = LBound(aryPath) To UBound(aryPath)
        If aryPath(i) <> "" Then
            If Not SerachFile(aryPath(i) & "/", strCreatedDir) Then
                blnResult = ExecuteFtpCMD("MKDIR " & strCreatedDir & aryPath(i) & "/")
                If Not blnResult Then Exit For
            End If
            
            strCreatedDir = strCreatedDir & aryPath(i) & "/"
        End If
    Next i

    CreateDir = blnResult
End Function



Public Function RemoteDir(ByVal strDelDir As String, Optional ByVal strFptDir As String = "") As String
'�Ƴ�Ŀ¼
'RMDIR����
'��ʽ��RMDIR file1
'˵����ɾ��Ŀ¼�������Ƿ��ܹ��ɹ�ִ�У�ȡ�����û���Զ�������ϵ�Ȩ��
    Dim strCurPath As String
    
    RemoteDir = False
    If strDelDir = "" Then Exit Function
    
    strCurPath = FormatPath(strDelDir, True)
    
    RemoteDir = ExecuteFtpCMD("RMDIR " & IIf(strFptDir = "", "", FormatPath(strFptDir, True)) & strCurPath)
End Function


Public Function GetFileSize(ByVal strFtpFile As String) As Double
'��ȡָ��ftp�ļ���С
'����ɹ������ļ����ֽ���,���򷵻�0 1k=1024�ֽ�

'SIZE����
'��ʽ��SIZE file1
'˵����������file1ָ�����ļ���С
    Dim strCurFile As String
    Dim blnResult As Boolean
    
    GetFileSize = 0
    
    strCurFile = GetFilePath(strFtpFile, True)
    strCurFile = FormatPath(strCurFile, True) & GetFileName(strFtpFile, True)

    blnResult = ExecuteFtpCMD("SIZE " & strCurFile)
    
    If blnResult Then
        GetFileSize = mstrData
    End If
End Function


Public Function GetCodeInf(ByVal lngState As Long) As String
    GetCodeInf = ""

    Select Case lngState
        Case icNone
            GetCodeInf = "��״̬�ɱ��档"
        Case icResolvingHost
            GetCodeInf = "�ÿؼ����ڲ�ѯ��ָ����������IP��ַ��"
        Case icHostResolved
            GetCodeInf = "�ÿؼ��ѳɹ����ҵ���ָ����������IP��ַ��"
        Case icConnecting
            GetCodeInf = "�ÿؼ��������������ӡ�"
        Case icConnected
            GetCodeInf = "�ÿؼ������������ӳɹ���"
        Case icRequesting
            GetCodeInf = "�ÿؼ�������������������"
        Case icRequestSent
            GetCodeInf = "�ÿؼ����������ѳɹ���"
        Case icReceivingResponse
            GetCodeInf = "�ÿؼ����ڽ�����������Ӧ��"
        Case icResponseReceived
            GetCodeInf = "�ÿؼ��ѳɹ��ؽ��յ���������Ӧ��"
        Case icDisconnecting
            GetCodeInf = "�ÿؼ����ڽ�������������ӡ�"
        Case icDisconnected
            GetCodeInf = "�ÿؼ��ѳɹ�����������������ӡ�"
        Case icError
            GetCodeInf = "������ͨѶʱ�����˴���"
        Case icResponseCompleted
            GetCodeInf = "�������Ѿ���ɣ������������ݾ��ѽ��յ���"
    End Select
End Function




Private Sub Class_Initialize()
    mblnExecuteState = False
    mlngLastState = -1
    
    Set mobjInet = Nothing
End Sub

Private Sub mobjInet_StateChanged(ByVal State As Integer)
On Error GoTo errHandle
    mlngLastState = State
    
    Select Case State
        Case icResponseCompleted
            '��ȡִ�к������
            mstrData = mobjInet.GetChunk(0, icString)
        Case icError
            '����ִ�гɹ�״̬
            mblnExecuteState = False
    End Select
    

'   ���ﲻ��ִ�������¼��������ڵ���GetChunk����ʱ�����ײ�����һ��������ִ�����������쳣
'    RaiseEvent OnStateChanged(State)
Exit Sub
errHandle:
    Debug.Print err.Description
End Sub