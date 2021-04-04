VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ͼ������"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer tmMsg 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "��������ͼ������رա�                                    ͼ��������ɺ���Զ��رա�����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

'�����ļ���
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Type NETRESOURCE ' ������Դ
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
Private Const RESOURCETYPE_ANY = &H0

Public Enum TMediaType
    imgTag = 0   'ͼ����
    MULFRAMETAG = 1 '����ͼ
    VIDEOTAG = 2 '��Ƶ���
    AUDIOTAG = 3 '��Ƶ���
End Enum

Private Const G_STR_HINT_TITLE As String = "��ʾ"

'Private WithEvents mobjIcon As clsTaskIcon  '������

Private curMsg As clsImgInfo

'��־��صĲ�������ע������ж�ȡ��־�����������־·��Ϊ�գ���ʹ��exe��ͬĿ¼�µġ�GetImgLog����Ϊ��־·��
Public mblnLogEnable As Boolean     '�Ƿ�������־
Public mstrLogPath As String        '��¼��־��·��,
Public mlngLogLevel As String       '��¼��־�ļ��𣬷ֳ�1,2������1��ֻ��¼��Ϣ�������־��2����¼ÿһ�����ص���־

Private mftpConnect As clsFtp        '����һ�������ӵ�FTP��
Private mftpConnectBak As clsFtp
Private mblnIsUpload As Boolean
Private mlngThreadID As Long

Private mConnectedSharedDir() As String  '��¼�Ѿ����ӹ��Ĺ���Ŀ¼
Private mlngRetriesn As Long        '�ϴ������غ����Դ���
Private mblnOpenDebug As Boolean
Public mobjDataQueue As clsDataQueue

Public Event OnComPlete(ByVal curMsg As Object)
Public Event OnState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean, ByVal lngThreadID As Long)

Public Sub DoComPlete(ByVal curMsg As Object)
BUGEX "DoComPlete 0"
    RaiseEvent OnComPlete(curMsg)
BUGEX "DoComPlete 1"
End Sub

Public Sub DoState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean, ByVal lngThreadID As Long)
BUGEX "DoState 0"
    RaiseEvent OnState(blnLoadFinish, blnUpLoad, lngThreadID)
BUGEX "DoState 1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo err
    '�Ͽ�FTP����
    Call mftpConnect.FuncFtpDisConnect
    Call mftpConnectBak.FuncFtpDisConnect
    
    Set mobjDataQueue = Nothing
    Set curMsg = Nothing
    
    Exit Sub
err:
    
End Sub

Public Sub zlInitModule(ByVal blnIsUpload As Boolean, ByVal lngThreadID As Long)
'------------------------------------------------
'���ܣ������򣬸�������ͼ�����س���
'������
'���أ���
'-----------------------------------------------
    Dim strRegPath As String
    
    Set mftpConnect = New clsFtp
    Set mftpConnectBak = New clsFtp
    
    mblnIsUpload = blnIsUpload
    mlngThreadID = lngThreadID
    
    '����������Ѿ�������һ�Σ���������
    If App.PrevInstance Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    Set mobjDataQueue = New clsDataQueue
    
    '��ע����ȡ��־����
    strRegPath = "����ģ��\zlGetImage"
    mblnLogEnable = CBool(GetSetting("ZLSOFT", strRegPath, "��¼��־", "True"))
    mstrLogPath = GetSetting("ZLSOFT", strRegPath, "��־·��", App.Path & "\GetImgLog")
    mlngLogLevel = Val(GetSetting("ZLSOFT", strRegPath, "��־����", 1))
    mblnOpenDebug = CBool(GetSetting("ZLSOFT", strRegPath, "IsOpenDebug", "True"))
    mlngRetriesn = Val(GetSetting("ZLSOFT", strRegPath, "���Դ���", 3))
BUGEX "mblnLogEnable=" & mblnLogEnable & "  mlngLogLevel=" & mlngLogLevel & "   mstrLogPath=" & mstrLogPath
    SaveSetting "ZLSOFT", strRegPath, "��¼��־", mblnLogEnable
    SaveSetting "ZLSOFT", strRegPath, "��־·��", mstrLogPath
    SaveSetting "ZLSOFT", strRegPath, "��־����", mlngLogLevel
    SaveSetting "ZLSOFT", strRegPath, "IsOpenDebug", mblnOpenDebug
    SaveSetting "ZLSOFT", strRegPath, "���Դ���", mlngRetriesn
    
    '�����������־�������־·���Ƿ����
    If mblnLogEnable = True Then
        '���û��������־·������ʹ��Ĭ��·��
        If mstrLogPath = "" Then
            mstrLogPath = App.Path & "\GetImgLog"
        End If
        
        '�����־·�������ڣ��򴴽�
        If Dir(mstrLogPath, vbDirectory) = "" Then
            'Ĭ��·�������ڣ��������Ŀ¼
            If Dir(mstrLogPath, vbDirectory) = "" Then
                Call MkLocalDir(mstrLogPath)
            End If
        End If
    End If

    '��ʼ������
    ReDim mConnectedSharedDir(0) As String
End Sub

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
    For i = 1 To UBound(mConnectedSharedDir)
        If mConnectedSharedDir(i) = strSharedDir Then
            funConnectAndSaveSheardDir = True
            Exit Function
        End If
    Next i
    
    '���ӹ���Ŀ¼
    If strSharedDir <> "" Then
        If funConnectSharedDir(strSharedDir, strUser, strPswd) = True Then
            '���ӳɹ�����¼�ɹ������Ӵ�
            ReDim Preserve mConnectedSharedDir(UBound(mConnectedSharedDir) + 1) As String
            mConnectedSharedDir(UBound(mConnectedSharedDir)) = strSharedDir
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

Public Function funDownLoadSharedDirSingle(strSourceDir As String, strDestDir As String) As Boolean
'------------------------------------------------
'���ܣ�ͨ������Ŀ¼�ķ�ʽ����һ��Ŀ¼�µ�һ��ͼ��strSourceDirĿ¼���ڵļ������¼��ͨ������Ĺ���ʵ��
'������ strSourceDir -- ��Ҫ�����ļ���ԴĿ¼����Զ�̷������еĹ���Ŀ¼
'       strDestDir  --  �ļ����Ƶ�Ŀ�ĵأ���������ı���Ŀ¼
'���أ�True-- �ɹ���False -- ʧ��
    Dim objFSO As New Scripting.FileSystemObject
    
    funDownLoadSharedDirSingle = False
    
BUGEX "funDownLoadSharedDirSingle S"

    If Dir(strDestDir, vbDirectory) = "" Then
        Call MkLocalDir(objFSO.GetParentFolderName(strDestDir))
    End If
    
BUGEX "strSourceDir=" & strSourceDir & "  strDestDir=" & strDestDir
    
    Call objFSO.CopyFile(strSourceDir, strDestDir, False)
    
BUGEX "funDownLoadSharedDirSingle=true"

    funDownLoadSharedDirSingle = True
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
    
BUGEX "funDownLoadSharedDir strSourceDir =" & strSourceDir & "  strDestDir=" & strDestDir
    
    '��ֹOverWrite����ֹ��FTP�е��ļ����Ǳ���Ŀ¼���Ѿ����˵��ļ���
    For Each fsFile In fsFiles
        '��¼��־
        Call WriteCommLog("funDownLoadSharedDir", "����ͼ��", "����ͼ�� " & strDestDir & "\" & fsFile.Name, 2)

BUGEX "Source=" & strSourceDir & "\" & fsFile.Name & "    Destination=" & strDestDir & "\" & fsFile.Name
        
        Call fs.CopyFile(strSourceDir & "\" & fsFile.Name, strDestDir & "\" & fsFile.Name, False)
    Next fsFile
    
    funDownLoadSharedDir = True
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

Public Function funDownLoadFTPSingle(ByVal thisMsg As clsImgInfo) As Boolean
'------------------------------------------------
'���ܣ���������ͼ�����Ϣ����FTP������ָ��Ŀ¼�е�һ��ͼ��
'������ thisMsg  -- ��Ҫ����ͼ�����Ϣ
'���أ�True -- �ɹ��� False -- ʧ��
'-----------------------------------------------
    Dim lngResult As String
    Dim objFile As New FileSystemObject
    Dim strLocalFileName As String
    Dim strVirtualPath As String
BUGEX "funDownLoadFTPSingle 0"
    funDownLoadFTPSingle = False
    On Error GoTo err
    
    '�ȼ����Ϣ�Ƿ����
    If thisMsg.Enable = False Or thisMsg.IP = "" Then
        
        '��¼��־
        Call WriteCommLog("funDownLoadFTP", "��Ϣ������", "FTP��ʽ����ͼ����Ϣ�����û���IP��ַΪ�գ��޷����ء� IP��ַ�ǣ�" & thisMsg.IP, 1)
        
        Exit Function
    End If

    '����FTP
    If mftpConnect.hConnection = 0 Then
        lngResult = mftpConnect.FuncFtpConnect(thisMsg.IP, thisMsg.FTPUser, thisMsg.FTPPswd)
BUGEX "����FTP"
        '�������ʧ�ܣ������ӱ����豸
        If lngResult = 0 And mftpConnectBak.hConnection = 0 Then
            lngResult = mftpConnectBak.FuncFtpConnect(thisMsg.BakIP, thisMsg.BakFTPUser, thisMsg.BakFTPPswd)
            
            If lngResult = 0 Then
            '��¼��־
            Call WriteCommLog("funDownLoadFTP", "FTP����ʧ��", "FTP����ʧ�ܣ� " & thisMsg.IP, 1)
    BUGEX "FTP����ʧ��"
                Exit Function
            End If
        End If
    End If
    
    '��������·��
    Call MkLocalDir(thisMsg.DestMainDir & objFile.GetParentFolderName(thisMsg.SubDir))
    
    '��¼��־
    Call WriteCommLog("funDownLoadFTP", "����ͼ��", "ͨ��FTP���ص���ͼ�� ", 1)
    
    strLocalFileName = Replace(thisMsg.DestMainDir & thisMsg.SubDir, "/", "\")
    strVirtualPath = Replace(thisMsg.FTPDir & objFile.GetParentFolderName(thisMsg.SubDir), "\", "/")
    
BUGEX "strVirtualPath=" & strVirtualPath & "  strLocalFileName=" & strLocalFileName & " strRemoteFileName=" & objFile.GetFileName(strLocalFileName)
    '�Ӵ洢�豸����ͼ��
    If mftpConnect.FuncDownloadFile(strVirtualPath, strLocalFileName, objFile.GetFileName(strLocalFileName)) <> 0 Then
        '����ʧ����ӱ����豸����ͼ��
        If mftpConnectBak.FuncDownloadFile(strVirtualPath, strLocalFileName, objFile.GetFileName(strLocalFileName)) <> 0 Then
            Exit Function
        End If
    End If
    
    funDownLoadFTPSingle = True
    Exit Function
err:
    '�ݲ�����
BUGEX "funDownLoadFTPSingle Error"
End Function

Public Function funDownLoadFTP(ByVal thisMsg As clsImgInfo) As Boolean
'------------------------------------------------
'���ܣ���������ͼ�����Ϣ����FTP������ָ��Ŀ¼�е�ȫ��ͼ��
'������ thisMsg  -- ��Ҫ����ͼ�����Ϣ
'���أ�True -- �ɹ��� False -- ʧ��
'-----------------------------------------------
    Dim lngResult As String
    
    funDownLoadFTP = False
    On Error GoTo err
    
    '�ȼ����Ϣ�Ƿ����
    If thisMsg.Enable = False Or thisMsg.IP = "" Then
        
        '��¼��־
        Call WriteCommLog("funDownLoadFTP", "��Ϣ������", "FTP��ʽ����ͼ����Ϣ�����û���IP��ַΪ�գ��޷����ء� IP��ַ�ǣ�" & thisMsg.IP, 1)
        
        Exit Function
    End If
BUGEX "����FTP"
    '����FTP
    If mftpConnect.hConnection = 0 Then
        lngResult = mftpConnect.FuncFtpConnect(thisMsg.IP, thisMsg.FTPUser, thisMsg.FTPPswd)
BUGEX "����FTP"
        '�������ʧ�ܣ������ӱ����豸
        If lngResult = 0 And mftpConnectBak.hConnection = 0 Then
            lngResult = mftpConnectBak.FuncFtpConnect(thisMsg.BakIP, thisMsg.BakFTPUser, thisMsg.BakFTPPswd)
            
            If lngResult = 0 Then
            '��¼��־
            Call WriteCommLog("funDownLoadFTP", "FTP����ʧ��", "FTP����ʧ�ܣ� " & thisMsg.IP, 1)
    BUGEX "FTP����ʧ��"
                Exit Function
            End If
        End If
    End If
    
    '��������·��
    Call MkLocalDir(thisMsg.DestMainDir & thisMsg.SubDir)
    
    '��¼��־
    Call WriteCommLog("funDownLoadFTP", "����ͼ��", "ͨ��FTP����ȫ��ͼ�� ", 1)
BUGEX "thisMsg.SubDir=" & Replace(thisMsg.FTPDir & thisMsg.SubDir, "\", "/") & "  strLocalPath = " & Replace(thisMsg.DestMainDir & thisMsg.SubDir, "/", "\")
    '����ͼ��
    If mftpConnect.funcDownLoadAllFiles(Replace("\" & thisMsg.FTPDir & thisMsg.SubDir, "\", "/"), Replace(thisMsg.DestMainDir & thisMsg.SubDir, "/", "\"), False) = 0 Then
        If mftpConnectBak.funcDownLoadAllFiles(Replace("\" & thisMsg.FTPDir & thisMsg.SubDir, "\", "/"), Replace(thisMsg.DestMainDir & thisMsg.SubDir, "/", "\"), False) = 0 Then
            Exit Function
        End If
    End If
    
    funDownLoadFTP = True
    Exit Function
err:
    '�ݲ�����
BUGEX "funDownLoadFTP Error"
End Function

Public Function funMsgProcess() As Boolean
'------------------------------------------------
'���ܣ��Զ�������Ϣ
'������ ��
'���أ�True -- �ɹ��� False -- ʧ��
'-----------------------------------------------
    Dim i As Integer
    Dim blnResult As Boolean
   
    On Error GoTo err

    funMsgProcess = False
    
    Call Me.DoState(False, mblnIsUpload, mlngThreadID)
    
    '��Ϣ���Ӳ�������Ϣ
    While funGetAMessage
BUGEX curMsg.SubDir
        If mblnIsUpload Then            '�ϴ�
BUGEX "UpLoadImages"
            '��¼��־
            Call WriteCommLog("funMsgProcess", "���յ���������Ϣ", "׼���ϴ���Ŀ¼�µ�ͼ��" & curMsg.SubDir, 1)
                
            '������Ϣ���ɹ�����true,ʧ�ܷ���false
            '���ϴ�ʧ�������³����ϴ�
            For i = 0 To mlngRetriesn
                blnResult = funUpLoadImages(curMsg)
                
                '�ϴ��ɹ����˳�
                If blnResult Then Exit For
            Next
            
        Else                                '����
BUGEX "DownLoadImages"
            '��¼��־
            Call WriteCommLog("funMsgProcess", "���յ���������Ϣ", "׼�����ش�Ŀ¼�µ�ͼ��" & curMsg.SubDir, 1)
    
            '������Ϣ���ɹ�����true,ʧ�ܷ���false
            '������ʧ�������³�������
            For i = 0 To mlngRetriesn
                blnResult = funDownLoadImages(curMsg)
                
                '���سɹ����˳�
                If blnResult Then Exit For
            Next
        End If
        
        If mblnIsUpload Then
BUGEX "UpLoadImages = " & blnResult
            If blnResult Then
                Call DoComPlete(curMsg)
            Else
                MsgBox "�ļ��ϴ�ʧ�ܣ������������粻�ȶ���ɡ�", vbExclamation, "��ʾ"
            End If
        Else
BUGEX "DownLoadImages = " & blnResult
            If blnResult Then Call DoComPlete(curMsg)
        End If
    Wend
    
    Call Me.DoState(True, mblnIsUpload, mlngThreadID)
    funMsgProcess = True
    Exit Function
err:
    '�ݲ�����
BUGEX "funMsgProcess Err: " & err.Description
End Function

Public Function funGetAMessage() As Boolean
'------------------------------------------------
'���ܣ�����Ϣ������ȡһ����Ϣ�������Ϣ�ṹ
'������
'���أ�True -- �ɹ� �� False -- ʧ��
'-----------------------------------------------
    Dim objMsg As Object
    Dim objOldMsg As Object
    Dim strMsgArr() As String
    
    funGetAMessage = False
    
    On Error GoTo err
    
    '��¼��һ����Ϣ�����������ͬ��������Ϣ���򲻴�������Ϣ
    Set objOldMsg = objMsg

    '����Ϣ������ȡһ����Ϣ
    Set objMsg = mobjDataQueue.MsgOutQueue

    '�����Ϣ��Ϊ�գ����������Ϣ
    If Not objMsg Is Nothing And Not objOldMsg Is objMsg Then
        Set curMsg = objMsg
        curMsg.Enable = True
        
        funGetAMessage = True
    End If
    
    Exit Function
err:
    '�ݲ�����
BUGEX "funGetAMessage err=" & err.Description
End Function

Public Function funDownLoadImages(ByVal thisMsg As clsImgInfo) As Boolean
'------------------------------------------------
'���ܣ���������ͼ�����Ϣ���ӹ���Ŀ¼����FTP����ͼ��
'������ thisMsg  -- ��Ҫ����ͼ�����Ϣ
'���أ�True -- �ɹ��� False -- ʧ��
'-----------------------------------------------
    Dim blnResult As Boolean
    
    On Error GoTo err
BUGEX "funDownLoadImages Start"
    '�����ǰ��Ϣ���ã�������Ϣ
    If thisMsg.Enable = True Then
        '�й���Ŀ¼����ʹ�ù���Ŀ¼����ͼ��û����ʹ��FTP����ͼ��
BUGEX "thisMsg.SDDir=" & thisMsg.SDDir
        If thisMsg.SDDir <> "" Then
BUGEX "funDownLoadImages SheardDir "
            '���ӹ���Ŀ¼
            If funConnectAndSaveSheardDir("\\" & thisMsg.IP & "\" & thisMsg.SDDir, thisMsg.SDUser, thisMsg.SDPswd) = True Then
                '��¼��־
                Call WriteCommLog("funDownLoadImages", "����Ŀ¼��ʽ����ͼ��", "ͨ������Ŀ¼��ʽ����ͼ�񣬴�Ŀ¼�� " & "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir & " �����ص���Ŀ¼�� " & thisMsg.DestMainDir & thisMsg.SubDir, 1)
                
                '����ͼ��
                If thisMsg.IsLoadSingleFile Then       '����Ŀ¼�µ�һ���ļ�
                    blnResult = funDownLoadSharedDirSingle("\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir, thisMsg.DestMainDir & thisMsg.SubDir)
                Else                '����Ŀ¼�µ������ļ�
                    blnResult = funDownLoadSharedDir("\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir, thisMsg.DestMainDir & thisMsg.SubDir)
                End If
            End If
        Else
BUGEX "funDownLoadImages FTP"
            '��¼��־
            Call WriteCommLog("funDownLoadImages", "FTP��ʽ����ͼ��", "ͨ��FTP��ʽ����ͼ�񣬴�Ŀ¼�� \" & thisMsg.FTPDir & "\" & thisMsg.SubDir & " �����ص���Ŀ¼�� " & thisMsg.DestMainDir & thisMsg.SubDir, 1)
                
            'ʹ��FTP����
            If thisMsg.IsLoadSingleFile Then           '����Ŀ¼�µ�һ���ļ�
                blnResult = funDownLoadFTPSingle(thisMsg)
            Else                    '����Ŀ¼�µ������ļ�
                blnResult = funDownLoadFTP(thisMsg)
            End If
        End If
    End If
    
    funDownLoadImages = blnResult
    
    Exit Function
err:
    '�ݲ�����
BUGEX "funDownLoadImages Err"
End Function

Public Function funUpLoadImages(ByVal thisMsg As clsImgInfo) As Boolean
'------------------------------------------------
'���ܣ���������ͼ�����Ϣ���ӹ���Ŀ¼����FTP����ͼ��
'������ thisMsg  -- ��Ҫ����ͼ�����Ϣ
'���أ�True -- �ɹ��� False -- ʧ��
'-----------------------------------------------
    Dim blnResult As Boolean
    
    On Error GoTo err
BUGEX "funUpLoadImages Start=" & thisMsg.Enable
BUGEX "funUpLoadImages thisMsg.BakIP=" & thisMsg.BakIP

    '�����ǰ��Ϣ���ã�������Ϣ
    If thisMsg.Enable = True Then
        '�й���Ŀ¼����ʹ�ù���Ŀ¼����ͼ��û����ʹ��FTP����ͼ��
        If thisMsg.SDDir <> "" Then
BUGEX "funUpLoadImages Sheard"
            '���ӹ���Ŀ¼
            If funConnectAndSaveSheardDir("\\" & thisMsg.IP & "\" & thisMsg.SDDir, thisMsg.SDUser, thisMsg.SDPswd) = True Then
                '��¼��־
                Call WriteCommLog("funUpLoadImages", "����Ŀ¼��ʽ�ϴ�ͼ��", "ͨ������Ŀ¼��ʽ�ϴ�ͼ�񣬴�Ŀ¼�� " & "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir & " ���ϴ�����Ŀ¼�� " & thisMsg.DestMainDir & thisMsg.SubDir, 1)

                '�ϴ�ͼ��
                If thisMsg.MediaType = VIDEOTAG Then
                    blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir & ".avi", "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir)
                ElseIf thisMsg.MediaType = AUDIOTAG Then
                    blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir & ".wav", "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir)
                Else
                    blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir, "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir)
                    blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir & ".jpg", "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir & ".jpg")
                End If
            End If
        Else
BUGEX "funUpLoadImages FTP thisMsg.IP=" & thisMsg.IP
            '��¼��־
            Call WriteCommLog("funUpLoadImages", "FTP��ʽ�ϴ�ͼ��", "ͨ��FTP��ʽ�ϴ�ͼ�񣬴�Ŀ¼�� \" & thisMsg.FTPDir & "\" & thisMsg.SubDir & " ���ϴ�����Ŀ¼�� " & thisMsg.DestMainDir & thisMsg.SubDir, 1)
                
            'ʹ��FTP�ϴ�
            blnResult = funUpLoadFTP(thisMsg, False)
        End If
        
        If thisMsg.BakIP <> "" Then '����ͼ��
            '�й���Ŀ¼����ʹ�ù���Ŀ¼����ͼ��û����ʹ��FTP����ͼ��
            If thisMsg.BakSDDir <> "" Then
BUGEX "funUpLoadImages Bak Sheard"
                '���ӹ���Ŀ¼
                If funConnectAndSaveSheardDir("\\" & thisMsg.BakIP & "\" & thisMsg.BakSDDir, thisMsg.BakSDUser, thisMsg.BakSDPswd) = True Then
                    '��¼��־
                    Call WriteCommLog("funUpLoadImages", "����Ŀ¼��ʽ����ͼ��", "ͨ������Ŀ¼��ʽ�ϴ�ͼ�񣬴�Ŀ¼�� " & "\\" & thisMsg.IP & "\" & thisMsg.SDDir & "\" & thisMsg.SubDir & " ���ϴ�����Ŀ¼�� " & thisMsg.DestMainDir & thisMsg.SubDir, 1)
    
                    '�ϴ�ͼ��
                    If thisMsg.MediaType = VIDEOTAG Then
                        blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir & ".avi", "\\" & thisMsg.BakIP & "\" & thisMsg.BakSDDir & "\" & thisMsg.SubDir)
                    ElseIf thisMsg.MediaType = AUDIOTAG Then
                        blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir & ".wav", "\\" & thisMsg.BakIP & "\" & thisMsg.BakSDDir & "\" & thisMsg.SubDir)
                    Else
                        blnResult = funUpLoadSharedDir(thisMsg.DestMainDir & thisMsg.SubDir, "\\" & thisMsg.BakIP & "\" & thisMsg.BakSDDir & "\" & thisMsg.SubDir)
                    End If
                End If
            Else
                '��¼��־
                Call WriteCommLog("funUpLoadImages", "FTP��ʽ����ͼ��", "ͨ��FTP��ʽ�ϴ�ͼ�񣬴�Ŀ¼�� \" & thisMsg.FTPDir & "\" & thisMsg.SubDir & " ���ϴ�����Ŀ¼�� " & thisMsg.DestMainDir & thisMsg.SubDir, 1)
BUGEX "funUpLoadImages Bak FTP"
                'ʹ��FTP�ϴ�
                blnResult = funUpLoadFTP(thisMsg, True)
            End If
        End If
    End If

BUGEX "funUpLoadImages End"
    funUpLoadImages = blnResult
    
    Exit Function
err:
    '�ݲ�����
BUGEX "funUpLoadImages Err"
End Function

Public Function funUpLoadSharedDir(strSourceDir As String, strDestDir As String) As Boolean
'------------------------------------------------
'���ܣ�ͨ������Ŀ¼�ķ�ʽ�ϴ�ͼ��strSourceDirĿ¼���ڵļ������¼��ͨ������Ĺ���ʵ��
'������ strSourceDir -- ��Ҫ�����ļ���ԴĿ¼����Զ�̷������еĹ���Ŀ¼
'       strDestDir  --  �ļ����Ƶ�Ŀ�ĵأ���������ı���Ŀ¼
'���أ�True-- �ɹ���False -- ʧ��
    Dim objFSO As New Scripting.FileSystemObject
    
BUGEX "funUpLoadSharedDir strSourceDir=" & strSourceDir & "   strDestDir=" & strDestDir
    funUpLoadSharedDir = False
    
    If Dir(objFSO.GetParentFolderName(strDestDir), vbDirectory) = "" Then
        Call MkLocalDir(objFSO.GetParentFolderName(strDestDir))
    End If

    Call objFSO.CopyFile(strSourceDir, strDestDir, True)
    
    funUpLoadSharedDir = True
End Function

Public Function funUpLoadFTP(ByVal thisMsg As clsImgInfo, Optional ByVal blnBakImg As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ������ϴ�ͼ�����Ϣ
'������ thisMsg  -- ��Ҫ����ͼ�����Ϣ
'       blnBakImg --True:����ͼ��,False:�洢ͼ��
'���أ�True -- �ɹ��� False -- ʧ��
    Dim lngResult As String
    Dim strSrcFileName As String
    Dim strVirtualPath As String
    Dim objFSO As New Scripting.FileSystemObject
    
    funUpLoadFTP = False
    On Error GoTo err
    
BUGEX "funUpLoadFTP blnBakImg=" & blnBakImg & "  thisMsg.BakIP=" & thisMsg.BakIP & "  thisMsg.IP= " & thisMsg.IP & "   thisMsg.Enable=" & thisMsg.Enable
    '�ȼ����Ϣ�Ƿ����
    If thisMsg.Enable = False Or IIf(blnBakImg, thisMsg.BakIP = "", thisMsg.IP = "") Then
        '��¼��־
        Call WriteCommLog("funUpLoadFTP", "��Ϣ������", "FTP��ʽ�ϴ�ͼ����Ϣ�����û���IP��ַΪ�գ��޷��ϴ��� IP��ַ�ǣ�" & IIf(blnBakImg, thisMsg.BakIP = "", thisMsg.IP = ""), 1)
BUGEX "funUpLoadFTP Exit"
        Exit Function
    End If
    
    '����FTP
    If blnBakImg Then   '���ӱ����豸
        If mftpConnectBak.hConnection = 0 Then
            If mftpConnectBak.FuncFtpConnect(thisMsg.BakIP, thisMsg.BakFTPUser, thisMsg.BakFTPPswd) = 0 Then
                Exit Function
            End If
        End If
    Else                '���Ӵ洢�豸
        If mftpConnect.hConnection = 0 Then
            If mftpConnect.FuncFtpConnect(thisMsg.IP, thisMsg.FTPUser, thisMsg.FTPPswd) = 0 Then
                Exit Function
            End If
        End If
    End If
    
    If thisMsg.MediaType = VIDEOTAG Then
        strSrcFileName = thisMsg.DestMainDir & thisMsg.SubDir & ".avi"
    ElseIf thisMsg.MediaType = AUDIOTAG Then
        strSrcFileName = thisMsg.DestMainDir & thisMsg.SubDir & ".wav"
    Else
        strSrcFileName = thisMsg.DestMainDir & thisMsg.SubDir
    End If
    
    strSrcFileName = Replace(strSrcFileName, "/", "\")
    strVirtualPath = Replace(IIf(blnBakImg, thisMsg.BakFTPDir, thisMsg.FTPDir) & objFSO.GetParentFolderName(thisMsg.SubDir), "\", "/")
    
    '��¼��־
    Call WriteCommLog("funUpLoadFTP", "�ϴ�ͼ��", "ͨ��FTP�ϴ�ͼ�� ", 1)
    
    '�ϴ�ͼ��
BUGEX "strVirtualPath=" & strVirtualPath
BUGEX "strSrcFileName=" & strSrcFileName
BUGEX "strRemoteFileName=" & objFSO.GetFileName(strSrcFileName)

    If blnBakImg Then
        Call mftpConnectBak.FuncFtpMkDir("/", strVirtualPath)
        Call mftpConnectBak.FuncUploadFile(strVirtualPath, strSrcFileName, objFSO.GetFileName(thisMsg.SubDir))
    Else
        Call mftpConnect.FuncFtpMkDir("/", strVirtualPath)
        Call mftpConnect.FuncUploadFile(strVirtualPath, strSrcFileName, objFSO.GetFileName(thisMsg.SubDir))
        
        If thisMsg.MediaType <> VIDEOTAG And thisMsg.MediaType <> AUDIOTAG Then
            Call mftpConnect.FuncUploadFile(strVirtualPath, strSrcFileName & ".jpg", objFSO.GetFileName(strSrcFileName) & ".jpg")
        End If
    End If
    
    funUpLoadFTP = True
    Exit Function
err:
    '�ݲ�����
BUGEX "funUpLoadFTP Err " & err.Description
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

    If mblnLogEnable = True Then        '�����˼�¼��־���ż�¼��ǰ����־
        '�ж���־����ȷ��������־�Ƿ���Ҫ��¼
        If mlngLogLevel >= lngLogLevel Then
            'ͨ����ǰʱ�䣬������־�ļ�����ÿ����Сʱ����һ����־�ļ�
            intHour = Hour(Time)
            intHour = intHour / 2
            intHour = intHour * 2
            strFileName = mstrLogPath & "\" & Format(Date, "YYYYMMDD") & "-" & intHour & ".log"
BUGEX "WriteCommLog strFileName=" & strFileName
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
BUGEX "WriteCommLog Err=" & err.Description
End Sub

Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    If mblnOpenDebug Or blnIsForce Then
        OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
    End If
End Sub

Private Sub tmMsg_Timer()
    If funMsgProcess Then
        Unload Me
    End If
End Sub

Public Sub zlLoadImage()
    tmMsg.Enabled = True
End Sub
