VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChatNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Public Function InitChat(ByRef objMain As Object, ByVal strChatURL As String, ByVal strMyChatUrl As String, _
    ByVal strUser As String, Optional ByRef strMsg As String) As Boolean
      '����:��ʼ��
      '����:
      'objMain-����̨����
      'strChatURL-�������۵�ַ
      'strMyChatUrl-�Ҳ�������۵�ַ
      'strUser-��ǰ����Ա
      'strMsg             -���ش�����Ϣ(���ⵯ��ģ̬��ʾ���������̹���,����ʾ��Ϣ���ظ������̴�����)
      '                    ���ظ�ʽ:��ʾ����[,]��ʾ���
          Dim strFilePath As String
          
1         On Error GoTo ErrH
          '�������û����־�ļ��½���־�ļ�
2         strFilePath = App.Path & "\" & App.EXEName & ".ini"
3         If Dir(strFilePath) = "" Then
              '[DEBUG]
              'OutPutDebugInfo = 0 ;���������Ϣ��zlCISRule.log�ļ���0-�����,1-������нӿڲ�����Ϣ
4             WriteIni "DEBUG", "OutPutDebugInfo", "0 ;���������Ϣ��" & App.EXEName & ".log�ļ�,0-�����,1-���", strFilePath
5         End If
6         gblnLog = (Val(ReadIni("Debug", "OutPutDebugInfo", strFilePath)) = 1)
          
7         WriteLog "������InitChat ��ʼ" & vbNewLine & _
                         "��Σ�objMain=" & (Not objMain Is Nothing) & ",�������۵�ַ=" & strChatURL & ",�ҵ����۵�ַ=" & strMyChatUrl & _
                         ",��ǰ�û�=" & strUser & vbNewLine
8         gstrSysName = "ͨ������"
9         gstrChatURL = strChatURL
10        gstrMyChatUrl = strMyChatUrl
11        gstrUser = strUser
          
12        Set gobjMain = objMain
13        Set gfrmMain = New frmMain
14        Set gcolChat = New Collection
          
15        gfrmMain.Visible = False
          '����ϵͳ����
16        Call gfrmMain.SetNotifyIcon(0)
          

17        InitChat = True
18        WriteLog "������InitChat ����" & vbNewLine
                  
19        Exit Function

ErrH:
20        strMsg = vbExclamation & "[,]" & "��zlChatNotify.clsChatNotify.InitChat�ĵ�" & Erl() & "�г�����" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description
21        WriteLog strMsg & vbNewLine
End Function

Public Function SendChat(ByVal strSubject As String, ByVal strSysCode As String, _
    ByVal strMainCode As String, ByVal dblMainId As Double, ByVal strSender As String, _
    ByVal strReceivers As String, Optional ByRef strMsg As String) As Boolean
'����:
'strSubject         -���۱���
'strSysCode         -ϵͳ����
'strMainCode        -�������
'dblMainId          -����ID
'strSender        -������
'strReceivers    -������(����������÷ָ���";"�ֿ�)
'strMsg             -���ش�����Ϣ(���ⵯ��ģ̬��ʾ���������̹���,����ʾ��Ϣ���ظ������̴�����)
'                    ���ظ�ʽ:��ʾ����[,]��ʾ���
    SendChat = OpenChatRoom("", strSubject, strSysCode, strMainCode, dblMainId, strSender, strReceivers, strMsg)
End Function

Public Function AddWaitList(ByVal rsList As ADODB.Recordset, Optional ByRef strMsg As String) As Boolean
'����:��δ����Ϣ����ȴ��嵥
'strMsg             -���ش�����Ϣ(���ⵯ��ģ̬��ʾ���������̹���,����ʾ��Ϣ���ظ������̴�����)
'                    ���ظ�ʽ:��ʾ����[,]��ʾ���
          Dim i As Long
          
1         On Error GoTo ErrH

2         If grsList Is Nothing Then Call InitRsList
          '1-��δ����Ϣ����ȴ��嵥
3         For i = 1 To rsList.RecordCount
4             Call grsList.AddNew(Array("ID", "Url", "Sys_Code", "Main_Code", "Main_ID", "Subject"), _
                  Array(Val(rsList!Id), rsList!Url & "", rsList!Sys_Code & "", rsList!Main_Code & "", rsList!Main_ID & "", rsList!Subject & ""))
5             rsList.MoveNext
6         Next
7         grsList.Filter = ""
8         If gblnShow Then
9             If Not frmChatList.RefreshList Then Exit Function
10        Else
              '2-����ϵͳ������˸
11            Call gfrmMain.SetIcon(1)
12        End If
13        AddWaitList = True

14        Exit Function

ErrH:
15        strMsg = vbExclamation & "[,]" & "��zlChatNotify.clsChatNotify.AddWaitList�ĵ�" & Erl() & "�г�����" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description
16        WriteLog strMsg & vbNewLine
End Function

Public Function ShowMyChat(ByVal lngUserId As Long, Optional ByRef strMsg As String) As Boolean
'����:�Ҳ��������
'strMsg   -���ش�����Ϣ(���ⵯ��ģ̬��ʾ���������̹���,����ʾ��Ϣ���ظ������̴�����)
'         ���ظ�ʽ:��ʾ����[,]��ʾ���
    Dim strKey As String
    Dim objChat As frmChat
    WriteLog "������ShowMyChat ��ʼ" & vbNewLine & _
                         "��Σ��û�ID=" & lngUserId & vbNewLine
    strKey = "K_USER_" & lngUserId
    On Error Resume Next
    Set objChat = gcolChat(strKey)
    On Error GoTo 0
    If objChat Is Nothing Then
        Set objChat = New frmChat
        gcolChat.Add objChat, strKey
    End If
    ShowMyChat = objChat.ShowMyChat(lngUserId, strMsg)
    WriteLog "������ShowMyChat ����" & vbNewLine
End Function

Private Sub Class_Terminate()
    Dim i As Long
    
    On Error Resume Next
    Unload gfrmMain
    
    Set gfrmMain = Nothing
    Set gobjMain = Nothing
    Set grsList = Nothing
    For i = 1 To gcolChat.Count
        Unload gcolChat(i)
        gcolChat.Remove i
    Next
    Set gcolChat = Nothing
End Sub

 