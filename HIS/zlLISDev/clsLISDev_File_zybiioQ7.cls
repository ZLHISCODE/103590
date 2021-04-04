VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_zybiioQ7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:zybiio-Q7���߶���������
'����:���
'�û�:���������ҽԺ

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'��׼���������ļ���ȡ������
'   strFile������·�����ļ���
'   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
'   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
'   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
'   ���صļ���������(������������밴���±�׼��֯���)
'   Ԫ��֮����|�ָ�
'   ��0��Ԫ�أ�����ʱ��
'   ��1��Ԫ�أ��������
'   ��2��Ԫ�أ�������
'   ��3��Ԫ�أ��걾
'   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
'   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
Dim XMLDoc As DOMDocument, XMLָ�� As DOMDocument, objFile As New Scripting.FileSystemObject
Dim strFilePath As String, strFileName As String
Dim xList  As IXMLDOMNodeList, yList As IXMLDOMNodeList
Dim aResults() As String, firstnode As IXMLDOMNode
Dim str�걾�� As String, str������ As String
Dim str���� As String
Dim i As Long, j As Long
On Error GoTo errHandle

    Set XMLDoc = New DOMDocument
    Set XMLָ�� = New DOMDocument
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFileName = Dir(strFilePath & "\*.xml")
    Do While strFileName <> ""
        If XMLDoc.Load(strFilePath & "\" & strFileName) Then
            str���� = Replace(Replace(XMLDoc.selectNodes("Export/When").Item(0).Text, "/", " "), ".", ":")
            Set xList = XMLDoc.getElementsByTagName("Cassette")
            For i = 0 To xList.Length - 1
                For j = 0 To xList.Item(i).childNodes.Length - 1
                    If xList.Item(i).childNodes.Item(j).nodeName = "SampleId" Then
                        str�걾�� = xList.Item(i).childNodes.Item(j).Text
                    End If
                    If xList.Item(i).childNodes.Item(j).nodeName = "Analyte" Then
                        str������ = str������ & "|" & xList.Item(i).childNodes.Item(j).childNodes(1).Text & "|" & xList.Item(i).childNodes.Item(j).childNodes(2).Text
                    End If
                Next j
                If Val(str�걾��) > 0 And str������ <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |ѪҺ|0" & str������
                    str������ = ""
                End If
            Next i
        End If
        If Not objFile.FolderExists(strFilePath & "\" & Format(str����, "yyyy-mm-dd")) Then objFile.CreateFolder (strFilePath & "\" & Format(str����, "yyyy-mm-dd"))
        objFile.CopyFile strFilePath & "\" & strFileName, strFilePath & "\" & Format(str����, "yyyy-mm-dd") & "\" & strFileName, True
        objFile.DeleteFile strFilePath & "\" & strFileName
        strFileName = Dir(strFilePath & "\*.xml")
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_FILE_XN1000.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function
