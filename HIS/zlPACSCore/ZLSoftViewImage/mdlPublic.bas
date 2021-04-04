Attribute VB_Name = "mdlPublic"
Option Explicit

'���ò����� {+}1405{+}ZLHIS[+]ZLHIS[+]HIS[+]0{+}false{+}false{+}0{+}0{+}false

Public gstrLogPath As String        '��־�ļ�
Public gstrImages As String         '��Ϣ���� strImages
Public glngOrderID As Long          '��Ϣ���� lngOrderID
Public gstrDBConnection As String   '��Ϣ���� strDBConnection
Public gblnMoved As Boolean         '��Ϣ���� blnMoved
Public gbAdd As Boolean             '��Ϣ���� bAdd
Public gintImageInterval As Integer '��Ϣ���� intImageInterval
Public glngSys As Long              '��Ϣ���� lngSys
Public gblnReconnectDB As Boolean   '��Ϣ���� blnReconnectDB
Public gstrDBServer As String       '��Ϣ���� strDBServer
Public gstrDBUser As String         '��Ϣ���� strDBUser
Public gstrDBPassword As String     '��Ϣ���� strDBPassword
Public gblnTransPassword As Boolean '��Ϣ���� blnTransPassword
Public gfrmViewImage As frmViewImage    '��Ϣѭ����������
Public gobjPacsCore As Object       '��Ƭ����
Public glngPreWndProc As Long       'ԭ������Ϣ�������
Public glngLog As Long              '�Ƿ��¼��־��0---����δ��ֵ��1---��¼��־��2---����¼��־

Public Const HIS_CAPTION = "����Ӱ���Ƭ����"
Public Const MSG_SPLIT = "{+}"

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Enum LogType
    ltError = 0
    ltDebug = 1
End Enum

Public Function errHandle(errSubName As String, errTitle As String, Optional errDesc As String = "") As Long
'------------------------------------------------
'���ܣ�������
'������ logSubName  --  ��������ĺ�����
'       logTitle   -- ��������
'       logDesc   --  ��������
'���أ�1-�������Resume��0-�����˳�
'------------------------------------------------
    
    errHandle = 0
    
    '��¼������־
    Call WriteCommLog("zlSoftViewImage,����--" & errSubName, errTitle & "���������= " & err.Number, errDesc & "����������=" & err.Description, ltError)
    
    '��ʾ����
    MsgBox errTitle & errDesc, vbOKOnly, "��Ƭ�ӿ�zlSoftViewImage���ִ���"
    
    '�������
    err.Clear
    
End Function

Public Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String, ByVal ltLogType As LogType)
'------------------------------------------------
'���ܣ���¼ͨѶ��־
'������ logSubName  --  ������־�ĺ�����
'       logTitle   -- ��־����
'       logDesc   --  ��־����
'       ltLogType --  ��־����
'���أ���
'------------------------------------------------
    Dim strLog As String
    Dim strFileName As String
    
    On Error GoTo err
    
    If glngLog = 0 Then
        glngLog = Val(GetSetting("ZLSOFT", "����ģ��\zl9PacsCore\zlSoftViewImage\", "Log", 2))
    End If
    
    'Log=1���ż�¼��־
    If glngLog <> 1 And ltLogType <> ltError Then Exit Sub
    
    strFileName = gstrLogPath & "\Interface" & Format(Date, "YYYY-MM-DD") & ".log"
    
    strLog = Now() & " ���⣺ " & logTitle & vbCrLf & "   ������ " & logSubName & vbCrLf & "   ��־���ݣ�" & logDesc & vbCrLf
    
    '������־���ӱ�ǣ�����鿴����
    If ltLogType = ltError Then
        strLog = "�������������" & strLog
    End If
    
    Open strFileName For Append As #1
    Print #1, strLog
    Close #1
    
    Exit Sub
err:
    Close #1
End Sub

Public Function GetLogDir() As String
'------------------------------------------------
'���ܣ���ȡ��־Ŀ¼�����Ŀ¼�����ڣ��򴴽�Ŀ¼
'��������
'���أ���־����Ŀ¼
'------------------------------------------------
    Dim strLogPath As String
    Dim strBackupPath As String
    
    On Error GoTo err
    
    strLogPath = App.Path & "\zlViewImageLog"
    
    Call MkLocalDir(strLogPath + "\")
    
    GetLogDir = strLogPath
   
    Exit Function
err:
    GetLogDir = App.Path & "\XWInterfaceLog"
    Call MkLocalDir(GetLogDir + "\")
End Function

Public Function ProcessMessage(strMsg As String) As Long
'------------------------------------------------
'���ܣ�������յ�����Ϣ
'������strMsg -- ����exeʱ����Ĳ�����
'���أ���
'------------------------------------------------
    
    Dim lngPartType As Long
    Dim strDBUser As String
    Dim lngPatientID As Long
    Dim lngClinicID As Long
    Dim lngDeptID As Long
    Dim lngOrderID As Long
    
    On Error GoTo err
    ProcessMessage = 1
    
    '����Ĳ������壬���������ӷ��������ַ���{+}��
    '������ʽ��strImages{+}lngOrderID{+}strDBConnection{+}blnMoved{+}bAdd{+}intImageInterval{+}lngSys{+}blnReconnectDB
    '�������ͣ� strImages --- ͼ���,�����ǡ�����UID1|1-3;5-27;33-100+����UID2|ȫ����,ȫ����ʾ��ȫ��ͼ��
    '           lngOrderID --- ҽ��ID
    '           strDBConnection --- ���ݿ����Ӵ���������������[+]�û���[+]����[+]�����Ƿ�ת���������ӷ��������ַ���[+]��
    '                          �������롱���û���¼����ʱ���������Ƿ�ת����=1���������롱�����ݿ��¼����ʱ���������Ƿ�ת����=0
    '           blnMoved --- �����Ƿ�ת��
    '           bAdd --- ��ѡ������Ĭ��ֵFalse����ͼ�������ӽ���Ƭվ�������滻ԭ��Ƭվ����ͼ��TrueΪ���ӣ�FasleΪ�滻
    '           intImageInterval --- ��ѡ������Ĭ��ֵ0����ͼ��ļ����ֻ�Դ�ȫ������,��������ͼ������>100ʱ��Ч
    '           lngSys --- ��ѡ������Ĭ��,100��ϵͳ���
    '           blnReconnectDB --- ��ѡ������Ĭ��ֵFalse���Ƿ������������ݿ⡣��һ�δ򿪹�Ƭʱ�Զ��������ݿ⣬֮���ٴ򿪹�Ƭ��
    '                           ��blnReconnectDB���������Ƿ������������ݿ⡣
    '                           =True��ʹ��strDBConnection���������������ݿ⣻=False�����������������ݿ⣬ʹ�ù�Ƭ�������ڵ����ݿ�����
    '
    
    '�ȴ���̶�����
    If UBound(Split(strMsg, MSG_SPLIT)) >= 3 Then
        gstrImages = Split(strMsg, MSG_SPLIT)(0)
        glngOrderID = Val(Split(strMsg, MSG_SPLIT)(1))
        gstrDBConnection = Split(strMsg, MSG_SPLIT)(2)
        gblnMoved = (UCase(Split(strMsg, MSG_SPLIT)(3)) = "TRUE")
    Else
        Call WriteCommLog("����--zlSoftShowHisForms.ProcessMessage", "��������", "������������������������4��������Ϊ��" & strMsg, ltError)
        Exit Function
    End If
    
    '�ٴ����ѡ����
    If UBound(Split(strMsg, MSG_SPLIT)) >= 4 Then
        gbAdd = (UCase(Split(strMsg, MSG_SPLIT)(4)) = "TRUE")
    Else
        gbAdd = False
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) >= 5 Then
        gintImageInterval = Val(Split(strMsg, MSG_SPLIT)(5))
    Else
        gintImageInterval = 0
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) >= 6 Then
        glngSys = Val(Split(strMsg, MSG_SPLIT)(6))
    Else
        glngSys = 100
    End If
    
    If UBound(Split(strMsg, MSG_SPLIT)) = 7 Then
        gblnReconnectDB = (UCase(Split(strMsg, MSG_SPLIT)(7)) = "TRUE")
    Else
        gblnReconnectDB = False
    End If
    
    If CreatePacsCore = False Then
        Exit Function
    End If
    
    Call WriteCommLog("zlSoftShowHisForms.ProcessMessage", "���ù�Ƭ", "��Ƭ�Ĳ����ǣ�gstrImages=" & gstrImages & ",glngOrderID=" & glngOrderID _
        & ",gstrDBConnection=" & gstrDBConnection & ",gblnMoved=" & gblnMoved & ",gbAdd=" & gbAdd & ",gintImageInterval=" & gintImageInterval _
        & ",glngSys=" & glngSys & ",gblnReconnectDB=" & gblnReconnectDB, ltDebug)
    
    Call gobjPacsCore.CallOpenViewerSimple(gstrImages, glngOrderID, gstrDBConnection, gblnMoved, gbAdd, gintImageInterval, glngSys, gblnReconnectDB)
    
    ProcessMessage = 0
    Exit Function
    
err:
    Call WriteCommLog("����--zlSoftShowHisForms.ProcessMessage", "������յ�����Ϣ�����ִ����յ�����Ϣ�ǣ�" & strMsg & "���������= " & err.Number, "����������=" & err.Description, ltError)
End Function

'******************************************************************************************************************
'���ܣ�����PACS��Ƭ����
'��������
'���أ������ɹ�,����true,���򷵻�False
'˵����
'******************************************************************************************************************
Private Function CreatePacsCore() As Boolean

    err = 0: On Error Resume Next
    If Not gobjPacsCore Is Nothing Then CreatePacsCore = True: Exit Function
    
    Set gobjPacsCore = CreateObject("zl9PacsCore.clsViewer")
    
    If err <> 0 Then
        MsgBox "δ�ҵ� zl9PacsCore �����������ǳ���汾��֧�֣������վ���Ƿ����˴˲���!", vbInformation + vbOKOnly, "��ʾ��Ϣ"
        Exit Function
    End If
    
    CreatePacsCore = True
    
End Function

Public Function CloseAllForms() As Boolean

    On Error GoTo err
    
    '�ر���Ϣѭ��������
    If Not gfrmViewImage Is Nothing Then
        Unload gfrmViewImage
        Set gfrmViewImage = Nothing
    End If
    
    CloseAllForms = True
    
    Exit Function
err:
    Call WriteCommLog("����--zlSoftViewImage.CloseAllForms", "�˳����򣬹ر����д��ڣ����ִ��󣬴������= " & err.Number, "����������=" & err.Description, ltError)
    Resume Next
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
