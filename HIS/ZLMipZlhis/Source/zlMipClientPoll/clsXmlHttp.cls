VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXmlHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'readyState��ʾXMLHttpRequest����Ĵ���״̬��
    '0:XMLHttpRequest����û����ɳ�ʼ����
    '1:XMLHttpRequest����ʼ��������
    '2:XMLHttpRequest�������������ɡ�
    '3:XMLHttpRequest����ʼ��ȡ����������Ӧ��
    '4:XMLHttpRequest�����ȡ��������Ӧ������
 
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private mobjHttp As XMLHTTP
Private mstrError As String
Private mstrPostData As String
Private mobjXMLData As Object

Public Function GetCode(ByVal strGetURL As String, pm, Optional ByVal strReferer As String, Optional ByVal strCookie As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngWaitTimeOut As Long
    Dim lngSecondNumber As Long
    
    On Error Resume Next
    
    If LCase(Left(strGetURL, 4)) <> "http" Then strGetURL = "http://" & strGetURL
    
    With mobjHttp
        Call .Open("GET", strGetURL, True)
        
        If strReferer <> "" Then Call .setRequestHeader("Referer", strReferer)
        If strCookie <> "" Then Call .setRequestHeader("Cookie", strCookie)
        
        .send
    End With
        
    lngWaitTimeOut = 0
    lngSecondNumber = 30 '��ʱ������
    
    Do
        DoEvents
        Call Wait(10)
        lngWaitTimeOut = lngWaitTimeOut + 1
    Loop Until (mobjHttp.readyState = 4 Or lngWaitTimeOut >= 100 * lngSecondNumber)
    
    If mobjHttp.readyState = 4 Then
        GetCode = BytesToBstr(mobjHttp.responseBody, pm)
    Else
        GetCode = ""
    End If
End Function

Public Function ChangeString(ByVal strText As String) As String
    '1. + URL ��+�ű�ʾ�ո� %2B
    '2. �ո� URL�еĿո������+�Ż��߱��� %20
    '3. / �ָ�Ŀ¼����Ŀ¼ %2F
    '4. ? �ָ�ʵ�ʵ� URL �Ͳ��� %3F
    '5. % ָ�������ַ� %25
    '6. # ��ʾ��ǩ %23
    '7. & URL ��ָ���Ĳ�����ķָ��� %26
    '8. = URL ��ָ��������ֵ %3D
    
    strText = Replace(strText, "%", "%25")
    strText = Replace(strText, "&", "%26")
    strText = Replace(strText, "=", "%3D")
    strText = Replace(strText, "/", "%2F")
    
    ChangeString = strText
End Function

Public Function PostData(ByVal strPostURL As String, _
                        ByVal strPostContent As String, _
                        Optional ByVal strPostCookie As String, _
                        Optional ByVal strPostRef As String, _
                        Optional strCharset As String, _
                        Optional ByRef strHeader As String, _
                        Optional ByVal intSendWaitTime As Integer = 30) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngWaitTimeOut As Long
    Dim oXmlDoc As New DOMDocument
    
    On Error GoTo errHand
        
    If UCase(Left(strPostURL, 4)) <> "HTTP" Then strPostURL = "http://" & strPostURL
    If strPostCookie = "" Then strPostCookie = "ASPSESSIONIDAQACTAQB=HKFHJOPDOMAIKGMPGBJJDKLJ;"
    
    strPostCookie = Replace(strPostCookie, Chr(32), "%20")
    
    With mobjHttp
        Call .Open("POST", strPostURL, True)
        Call .setRequestHeader("Content-Length", LenB(strPostContent))
'        Call .setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        Call .setRequestHeader("Content-Type", "text/xml; charset=utf-8")
'        Call .setRequestHeader("SOAPAction", "http://WebXml.com.cn" & "/getWeatherbyCityName")
        If strPostRef <> "" Then Call .setRequestHeader("Referer", strPostRef)
        If strPostCookie <> "" Then Call .setRequestHeader("Cookie", strPostCookie)
        Call .send(strPostContent)
    End With
    
    lngWaitTimeOut = 0
'    lngSecondNumber = 30 '��ʱ������
    Do
        DoEvents
        Call Wait(10)
        lngWaitTimeOut = lngWaitTimeOut + 1
    Loop Until (mobjHttp.readyState = 4 Or lngWaitTimeOut >= 100 * intSendWaitTime)
    
    If mobjHttp.readyState = 4 Then
        Set oXmlDoc = New DOMDocument
        oXmlDoc.Load mobjHttp.responseXML
        
        If oXmlDoc.xml = "" Then
            mstrError = mobjHttp.responseText
            PostData = False
        Else
            mstrPostData = oXmlDoc.xml
            PostData = True
        End If
        
    Else
        mstrError = mobjHttp.responseText
'        mstrError = BytesToBstr(mobjHttp.responseBody, strCharset)
        PostData = False
    End If
    
    strHeader = mobjHttp.getAllResponseHeaders()
    
    Exit Function
    
errHand:
    mstrError = Err.Description
End Function

Public Function GetPostDataReturn() As String
    GetPostDataReturn = mstrPostData
End Function

Public Function GetLastError() As String
    GetLastError = mstrError
End Function

Private Sub Wait(tt)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim t, t1, t2, i
    t = tt
    If t > 10 Then
        t1 = Int(t / 10)
        t2 = t - t1 * 10
        For i = 1 To t1
            Call Sleep(10)
            DoEvents
        Next i
        If t2 > 0 Then Call Sleep(t2)
    Else
        If t > 0 Then Call Sleep(t)
    End If
End Sub

Private Function BytesToBstr(varBody, ByVal strCharset As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    On Error Resume Next
    
    Dim objStream As Object
    
    Set objStream = CreateObject("adodb.stream")
    objStream.Type = 1
    objStream.Mode = 3
    objStream.Open
    objStream.Write varBody
    objStream.Position = 0
    objStream.Type = 2
    objStream.Charset = strCharset
    BytesToBstr = objStream.ReadText
    objStream.Close
    Set objStream = Nothing
    
End Function

Private Sub Class_Initialize()
    Set mobjHttp = CreateObject("msxml2.serverXMLHTTP")
End Sub

Private Sub Class_Terminate()
    Set mobjHttp = Nothing
End Sub


