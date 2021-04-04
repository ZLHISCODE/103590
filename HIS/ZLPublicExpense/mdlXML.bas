Attribute VB_Name = "mdlXML"
Option Explicit
Public gobjXml As MSXML2.DOMDocument
Private mintDebug As Integer
Public Function zlXML_Init(Optional ByVal strNode As String = "DATA", _
    Optional blnNotMsg As Boolean = False, Optional strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��XML,���������͸��ڵ�
    '���:strNode-�ӵ�
    '����:strErrMsg-���ش�����Ϣ
    '����: ��ʼ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2011-05-27 10:58:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim nodData As MSXML2.IXMLDOMElement
    mintDebug = -1
    On Error Resume Next
    Set gobjXml = New MSXML2.DOMDocument
    If Err <> 0 Then
        DebugTools "��ʼ��XML����ʧ��(InitXML)�� " & strNode & "��"
        strErrMsg = "��ʼ��XML����ʧ��(InitXML)�� " & strNode & "��"
        Err.Raise vbObjectError + 1, , strErrMsg
    End If
    '���ڵ�
    Err = 0: On Error GoTo errHand:
    Set nodData = gobjXml.createElement(strNode)
    Set gobjXml.documentElement = nodData
    zlXML_Init = True: Exit Function
errHand:
    If blnNotMsg Then strErrMsg = Err.Description: Exit Function
    If ErrCenter = 1 Then Resume
End Function
Public Function zlXML_InsertNodes(nodParent As MSXML2.IXMLDOMElement, _
    ByVal cllData As Collection) As Boolean
    '------------------------------------------------------------------s---------------------------------------------------------------------------
    '����:����ӵ���
    '���:nodParent-���ӵ�
    '        cllData-����Array(�ӵ���,�ӵ�ֵ)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-27 11:03:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    On Error GoTo errHandle
    For i = 1 To cllData.Count
        Call zlXML_InsertNode(nodParent, cllData(i)(0), cllData(i)(1))
    Next
    zlXML_InsertNodes = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlXML_InsertNode(nodParent As MSXML2.IXMLDOMElement, _
    ByVal Name As String, ByVal value As String, Optional ByRef OutNod As MSXML2.IXMLDOMElement) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ��XMLԪ����������Ԫ��
    '���:nodParent-���ӵ�
    '        Name-�ӵ���
    '        Value-�ӵ�ֵ
    '����:OutNod-���ؽӵ����
    '����:���ӳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-27 11:26:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHandle
    Set OutNod = gobjXml.createElement(Name)
    OutNod.Text = value
    nodParent.appendChild OutNod
    zlXML_InsertNode = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlXML_GetXMLString(Optional blnHead As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡXML�ַ���
    '���:blnHead-�Ƿ����ͷ����
    '����:������XML��
    '����:���˺�
    '����:2011-05-27 11:31:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnHead Then
        zlXML_GetXMLString = gobjXml.xml
    Else
        zlXML_GetXMLString = "<?xml version=""1.0"" encoding=""gb2312""?>" & vbCrLf & gobjXml.xml
    End If
End Function

Public Function zlXML_GetRows(ByVal strNodeName As String, ByRef lngOutRows As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡXML����
    '���:strNodeName-�ӵ���
    '����:lngOutRows-����XML����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-27 10:51:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    lngOutRows = 0
    Err = 0: On Error GoTo errHand:
    If mintDebug = 5 Then
       Call zlTempLoadXML
    End If
    lngOutRows = gobjXml.getElementsByTagName(strNodeName).length
    DebugTools "��ȡXML�ļ�¼����(GetOutXMLRows)�� " & strNodeName & "��:" & lngOutRows
    zlXML_GetRows = True
    Exit Function
errHand:
    DebugTools "��ȡXML�ļ�¼����(GetOutXMLRows)�� " & strNodeName & "��" & vbCrLf & "�������:" & vbCrLf & "   " & Err.Description
    If ErrCenter = 1 Then Resume
End Function
Private Sub zlTempLoadXML()
    'J������:��ʱ����XML�ļ�
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Set objText = objFile.OpenTextFile(App.Path & "\xml.txt", ForReading)
    Call zlXML_LoadXMLToDOMDocument(objText.ReadAll)
End Sub

Public Function zlXML_GetNodeValue( _
    ByVal strNodeName As String, Optional ByVal lngRow As Long = 0, _
    Optional ByRef strOutPut As String, Optional ByRef strErrMsg As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�õ�ָ��Ԫ�ص�ֵ
    '���:strNodeName-�ӵ���
    '       lngRow-ָ������
    '       strErrMsg-������Ϣ
    '����:strOutPut-����ֵ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-27 10:52:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim xmlElement As MSXML2.IXMLDOMElement
    Err = 0: On Error GoTo errHand:
    strOutPut = "": strErrMsg = ""
    If lngRow >= 0 Then
        Set xmlElement = gobjXml.getElementsByTagName(strNodeName).Item(lngRow)
    Else
        Set xmlElement = gobjXml.documentElement.selectSingleNode(strNodeName)
    End If
    If Not xmlElement Is Nothing Then
        '�ҵ�ָ����Ԫ��
        strOutPut = Replace(xmlElement.Text, Chr(10), "")
    Else
        strErrMsg = strNodeName & "�����ڣ�����!"
        DebugTools strErrMsg
        If Not gobjXml Is Nothing Then
            DebugTools gobjXml.xml
        Else
            DebugTools "gobjXml.xml=nothing"
        End If
    End If
    zlXML_GetNodeValue = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlXML_LoadXMLToDOMDocument(ByVal strXMLInPut As String, _
    Optional blnAddHead As Boolean = True, Optional blnNotMsg As Boolean = False, _
    Optional ByRef strErrMsg As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����XML����صĶ���
    '���:strXMLInPut-��ص�XML��
    '����:strErrMsg-���ش�����Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-03-13 11:07:00
    '-----------------------------------------------------------------------------------------------------------
    Dim strXml As String
    Err = 0: On Error GoTo errHand:
    If Not blnAddHead Then
        strXml = strXMLInPut
    Else
        strXml = Replace("'<?xml version=''1.0'' encoding=''gb2312''?>'", "'", Chr(34)) & vbCrLf & strXMLInPut
    End If
    If Not gobjXml.loadXML(strXml) Then
        strErrMsg = "XML��������(LoadXML):" & vbCrLf & strXml
        DebugTools strErrMsg
        Err.Raise vbObjectError + 1, , strErrMsg
    End If
    zlXML_LoadXMLToDOMDocument = True
    Exit Function
errHand:
    strErrMsg = Err.Description
    If blnNotMsg Then Exit Function
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub DebugTools(ByVal strInfo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ٵ�����Ϣ
    '���:strInfo-������Ϣ
    '����:���˺�
    '����:2011-05-27 11:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call LogWrite("һ��ͨ�ӿڵ�����־", glngModul, "�����ӿڷ���", strInfo)
End Sub
 
Public Function zlXML_GetChildNodeValue(strParentName As String, strChildName As String, Optional ByVal lngParentRow As Long = 0, Optional ByVal lngChildRow As Long = 0, Optional ByRef strOutPut As String, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ӧ���ڵ����ӽڵ��ֵ
    '���:strParentName-���ڵ����� strChildName-�ӽڵ����� lngParentRow - ָ�������� lngChildRow - ָ��������
    '����:strOutPut-����ֵ strErrMsg -������Ϣ
    '����:����
    '����:2012-12-12 11:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim xmlParentElement As MSXML2.IXMLDOMElement
    Dim xmlChildElement As MSXML2.IXMLDOMElement
    Err = 0: On Error GoTo errHand:
    If lngParentRow >= 0 Then
        Set xmlParentElement = gobjXml.documentElement.getElementsByTagName(strParentName).Item(lngParentRow)
    Else
        Set xmlParentElement = gobjXml.documentElement.selectSingleNode(strParentName)
    End If
    If lngChildRow >= 0 Then
        Set xmlChildElement = xmlParentElement.getElementsByTagName(strChildName).Item(lngChildRow)
    Else
        Set xmlChildElement = xmlParentElement.selectSingleNode(strChildName)
    End If
    If Not xmlChildElement Is Nothing Then
        '�ҵ�ָ����Ԫ��
        strOutPut = Replace(xmlChildElement.Text, Chr(10), "")
    Else
        strErrMsg = strChildName & "�����ڣ�����!"
        DebugTools strErrMsg
        If Not gobjXml Is Nothing Then
            DebugTools gobjXml.xml
        Else
            DebugTools "gobjXml.xml=nothing"
        End If
    End If
    zlXML_GetChildNodeValue = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlXML_GetChildRows(ByVal strParentName, ByVal strChildName As String, ByRef lngOutRows As Long, Optional ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ӧ���ڵ����ӽڵ������
    '���:strParentName-���ڵ��� ,strChildName-�ӽڵ���
    '����:lngOutRows-����XML����
    '����:�ɹ�,����true,���򷵻�False
    '����:����
    '����:2011-12-12 10:51:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    lngOutRows = 0
    Err = 0: On Error GoTo errHand:
    If mintDebug = 5 Then
       Call zlTempLoadXML
    End If
    lngOutRows = gobjXml.getElementsByTagName(strParentName).Item(lngRow).selectNodes(strChildName).length
    DebugTools "��ȡXML�ļ�¼����(GetOutXMLRows)�� " & strChildName & "��:" & lngOutRows
    zlXML_GetChildRows = True
    Exit Function
errHand:
    DebugTools "��ȡXML�ļ�¼����(GetOutXMLRows)�� " & strChildName & "��" & vbCrLf & "�������:" & vbCrLf & "   " & Err.Description
    If ErrCenter = 1 Then Resume
End Function

Public Function zlXML_GetChildNodes(ByVal strParentName As String) As IXMLDOMNodeList
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ӧ���ڵ����ӽڵ㼯��
    '���:strParentName-���ڵ���
    '����:lngOutRows-����XML����
    '����:�ɹ�,����true,���򷵻�False
    '����:����
    '����:2011-12-13 14:51:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim xmlParentElement As MSXML2.IXMLDOMElement

    Set xmlParentElement = gobjXml.documentElement.selectSingleNode(strParentName)
    If Not xmlParentElement Is Nothing Then
       Set zlXML_GetChildNodes = xmlParentElement.childNodes
    Else
       Set zlXML_GetChildNodes = Nothing
    End If
End Function

Public Function zlXML_ExistNode(ByVal strXMLIn As String, ByVal strNodeName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:XML�����Ƿ����ĳ�ڵ�
    '���:strXMLInPut-��ص�XML��
    '   strNodeName-�ڵ���
    '����:�ڵ����,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    zlXML_ExistNode = InStr(strXMLIn, "<" & strNodeName & ">") > 0
    Exit Function
ErrHandler:
    zlXML_ExistNode = False
End Function

