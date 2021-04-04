Attribute VB_Name = "mdlEZCA"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'��������ʱ���webservice��ʽ
Private mobjTSA As Object
Private strUrl As String
Private userid As String
Private userkey As String
Private lngUSETSA As Long

Public Function TSAWEB_initObj() As Boolean
    On Error Resume Next
    Set mobjTSA = Nothing
    Set mobjTSA = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        Exit Function
    End If
    strUrl = ReadIni("TSA", "URL", App.Path & "\TSA.ini")
    userid = ReadIni("TSA", "USERID", App.Path & "\TSA.ini")
    userkey = ReadIni("TSA", "USERKEY", App.Path & "\TSA.ini")
    lngUSETSA = ReadIni("TSA", "USETSA", App.Path & "\TSA.ini")
    If strUrl = "" Or userid = "" Or userkey = "" Then
        Err.Raise -1, , "TSA.ini�ļ������ڻ����ô���"
        Exit Function
    End If
    If lngUSETSA = 0 Then
        Exit Function  'TSA.ini�ļ��е�USETSAֵΪ0��������ʱ���
    End If
    mobjTSA.MSSoapInit ReadIni("TSA", "URL", App.Path & "\TSA.ini")
    TSAWEB_initObj = True
End Function

Public Function TSAWEB_UnloadObj()
    'ጷŌ���
    If Not mobjTSA Is Nothing Then Set mobjTSA = Nothing
End Function

Private Function GetReturnInfo(ByVal strSign As String) As String
    'ʱ���������Ϣת������
    If strSign = "0001" Then
        GetReturnInfo = "����ͨ���쳣"
    ElseIf strSign = "0002" Then
        GetReturnInfo = "ϵͳ�쳣"
    ElseIf strSign = "0003" Then
        GetReturnInfo = "ϵͳ��æ"
    ElseIf strSign = "0004" Then
        GetReturnInfo = "���ݲ������Ϸ�"
    ElseIf strSign = "0005" Then
        GetReturnInfo = "�û������������"
    ElseIf strSign = "0006" Then
        GetReturnInfo = "���ݿ��쳣"
    ElseIf strSign = "1001" Then
        GetReturnInfo = "������Ӧʧ��"
    ElseIf strSign = "1002" Then
        GetReturnInfo = "���������ѼӸǹ�ʱ���"
    ElseIf strSign = "1003" Then
        GetReturnInfo = "�������ݵȴ��Ӹ�ʱ���"
    ElseIf strSign = "2000" Then
        GetReturnInfo = "��֤�ɹ�"
    ElseIf strSign = "2001" Then
        GetReturnInfo = "δ����ʱ���"
    ElseIf strSign = "2002" Then
        GetReturnInfo = "ǩ��У��ʧ��"
    ElseIf strSign = "3010" Then
        GetReturnInfo = "��ʱ�������֤�ɹ�"
    ElseIf strSign = "3020" Then
        GetReturnInfo = "��ʱ����ļ�������֤�ɹ�"
    ElseIf strSign = "3030" Then
        GetReturnInfo = "��ʱ�������ʱ����ļ�����֤�ɹ�"
    Else
        GetReturnInfo = strSign
    End If
    If GetReturnInfo <> "" Then
        GetReturnInfo = "ʱ����ӿڷ�����ʾ��" & GetReturnInfo
    End If
End Function

Public Function TimesWEB_Tamp(ByVal strSource As String, ByRef strTimeStamp As String) As Boolean
        'ȡʱ���
        Dim intCount As Integer, strSign As String
        Dim sz������Ϣ
        On Error GoTo hErr
        
        If mobjTSA Is Nothing Then Exit Function
        
100     strSign = mobjTSA.applyTimeStamp(userid, userkey, "sha1", StringSHA1(strSource))(0)
102     If strSign <> "1000" And strSign <> "1002" And strSign <> "1003" Then
104         strSign = GetReturnInfo(strSign)
106         MsgBoxEx "����ʱ���ʧ�ܣ�" & strSign, vbExclamation, gstrSysName
            TimesWEB_Tamp = False
            Exit Function
        Else
108         intCount = 0
110         Do While intCount <= 100
                sz������Ϣ = mobjTSA.GetTimeStamp(userid, userkey, 1, "sha1", StringSHA1(strSource))
112             strSign = sz������Ϣ(0)
                'ǩ���л���ʱ��
114             If strSign = 3010 Then
                    strTimeStamp = sz������Ϣ(1)
118                 If IsDate(strTimeStamp) Then
120                     strTimeStamp = Format(CDate(strTimeStamp), "yyyy-MM-dd HH:mm:ss")
                        TimesWEB_Tamp = True
                        Exit Function
                    Else
122                     MsgBoxEx "��ȡ��ʱ�������һ�����ڣ�" & strTimeStamp, vbExclamation, gstrSysName
                    End If
124             ElseIf strSign <> "1003" And strSign <> "2001" Then
126                 strSign = GetReturnInfo(strSign)
128                 MsgBoxEx "��ȡʱ���ʧ�ܣ�" & strSign, vbExclamation, gstrSysName
                    Exit Function
                End If
130             intCount = intCount + 1
            Loop
        End If
132     TimesWEB_Tamp = True
        Exit Function
hErr:
134    MsgBoxEx "ȡʱ���-��" & CStr(Erl()) & "��," & Err.Description, vbExclamation, gstrSysName
End Function

Public Function verifyWEB_Timestamp(ByVal strSource As String) As Boolean
    '��֤ʱ���
    Dim strData As String
    If mobjTSA Is Nothing Then Exit Function
    strData = mobjTSA.verifyTimeStamp(userid, userkey, "sha1", StringSHA1(strSource))(0)
    If strData <> "2010" Then
        MsgBoxEx "��֤ʱ���ʧ�ܣ�" & GetReturnInfo(strData), vbExclamation, gstrSysName
        Exit Function
    End If
    verifyWEB_Timestamp = True
End Function

Public Function verifyWEB_getTimestamp(ByVal strSource As String) As String
    '��ȡʱ���
    Dim strData As String
    Dim strTimeStamp As String
    If mobjTSA Is Nothing Then Exit Function
    
    strData = mobjTSA.GetTimeStamp(userid, userkey, 1, "sha1", StringSHA1(strSource))(0)
    strTimeStamp = mobjTSA.GetTimeStamp(userid, userkey, 1, "sha1", StringSHA1(strSource))(1)
    If strData = "2001" Then
        MsgBoxEx "��ȡ��֤ʱ���ʧ�ܣ�" & GetReturnInfo(strData), vbExclamation, gstrSysName
        verifyWEB_getTimestamp = "��"
        Exit Function
    End If

    If IsDate(strTimeStamp) Then
        strTimeStamp = Format(CDate(strData), "yyyy-MM-dd HH:mm:ss")
    Else
        MsgBoxEx "��ȡ��ʱ�������һ�����ڣ�" & strData, vbExclamation, gstrSysName
        verifyWEB_getTimestamp = "��"
        Exit Function
    End If

    verifyWEB_getTimestamp = strTimeStamp
    
End Function

Public Function ReadIni(strItem As String, strKey As String, strPath As String) As String
    Dim GetStr As String
    On Error GoTo errH

    GetStr = VBA.String(128, 0)
    GetPrivateProfileString strItem, strKey, "", GetStr, 256, strPath
    GetStr = VBA.Replace(GetStr, VBA.Chr(0), "")
    ReadIni = GetStr
    Exit Function
errH:
    Err.Clear
    ReadIni = ""
End Function

