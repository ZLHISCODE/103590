VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public gobjXml As MSXML2.DOMDocument
Private mintDebug As Integer
Public Function zlXML_Init(Optional ByVal strNode As String = "DATA", _
    Optional ByVal blnNotErrMsg As Boolean, Optional ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��XML,���������͸��ڵ�
    '���:strNode-�ӵ�
    '����:strErrMsg-���صĴ�����Ϣ
    '����:��ʼ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-27 10:58:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim nodData As MSXML2.IXMLDOMElement
    '���ڵ�
    Set nodData = gobjXml.createElement(strNode)
    Set gobjXml.documentElement = nodData
    zlXML_Init = True
End Function

Public Function zlXML_InsertNodes(nodParent As MSXML2.IXMLDOMElement, _
    ByVal cllData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlXML_InsertNode(nodParent As MSXML2.IXMLDOMElement, _
    ByVal Name As String, ByVal Value As String, _
    Optional ByRef OutNod As MSXML2.IXMLDOMElement) As Boolean
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
    OutNod.Text = Value
    nodParent.appendChild OutNod
    zlXML_InsertNode = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
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
         ' "<?xml version=""1.0"" encoding=""utf-8"" ?>"
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
    lngOutRows = gobjXml.getElementsByTagName(strNodeName).Length
    DebugTools "��ȡXML�ļ�¼����(GetOutXMLRows)�� " & strNodeName & "��:" & lngOutRows
    zlXML_GetRows = True
    Exit Function
errHand:
    DebugTools "��ȡXML�ļ�¼����(GetOutXMLRows)�� " & strNodeName & "��" & vbCrLf & "�������:" & vbCrLf & "   " & Err.Description
    If gobjComLib.ErrCenter() = 1 Then Resume
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
    Optional ByRef strOutput As String, Optional ByRef strErrMsg As String = "") As Boolean
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
    strOutput = "": strErrMsg = ""
    If lngRow >= 0 Then
        Set xmlElement = gobjXml.getElementsByTagName(strNodeName).Item(lngRow)
    Else
        Set xmlElement = gobjXml.documentElement.selectSingleNode(strNodeName)
        
    End If
    If Not xmlElement Is Nothing Then
        '�ҵ�ָ����Ԫ��
        strOutput = Replace(xmlElement.Text, Chr(10), "")
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
    DebugTools strErrMsg
End Function
Public Function zlXML_LoadXMLToDOMDocument(ByVal strXMLInPut As String, _
     Optional blnAddHead As Boolean = True, Optional strErrMsg As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����XML����صĶ���
    '���:strXMLInPut-��ص�XML��
    '����:
    '       strErrMsg-���ش�����Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-03-13 11:07:00
    '-----------------------------------------------------------------------------------------------------------
    Dim strXML As String
    Err = 0: On Error GoTo errHand:
    If Not blnAddHead Then
        strXML = strXMLInPut
    Else
        strXML = Replace("'<?xml version=''1.0'' encoding=''gb2312''?>'", "'", Chr(34)) & vbCrLf & strXMLInPut
    End If
    DebugTools "�������XML������:"
    If Not gobjXml.loadXML(strXML) Then
        strErrMsg = "XML��������"
        DebugTools "XML��������" & strXML
        Exit Function
    End If
    DebugTools "����XML������ɹ�:"
    zlXML_LoadXMLToDOMDocument = True
    Exit Function
errHand:
    strErrMsg = zlTranErrInfor(Err.Description)
    DebugTools "����XML������ʧ��:" & strErrMsg
End Function

Private Sub DebugTools(ByVal strInfo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ٵ�����Ϣ
    '���:strInfo-������Ϣ
    '����:���˺�
    '����:2011-05-27 11:36:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call gobjComLib.LogWrite("һ��ͨ�ӿڵ�����־", glngModul, "�����ӿڷ���", strInfo)
End Sub
Public Property Get zlGetXMLHeadString() As String
    zlGetXMLHeadString = "<?xml version=""1.0"" encoding=""utf-8"" ?>"
End Property
Private Function TranErrInfor(ByVal strErrMsg As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʽ��������Ϣ
    '����:���˺�
    '����:2012-06-21 10:49:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If InStr(UCase(strErrMsg), "[ZLSOFT]") > 0 Then
        TranErrInfor = Split(strErrMsg, "[ZLSOFT]")(1)
    Else
        TranErrInfor = strErrMsg
    End If
    Exit Function
errHandle:
      TranErrInfor = strErrMsg
End Function
Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    Set gobjXml = New MSXML2.DOMDocument
    mintDebug = -1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set gobjXml = Nothing
    mintDebug = -1
End Sub