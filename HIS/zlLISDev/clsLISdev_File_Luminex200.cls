VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_File_Luminex200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:��ʽӫ�����߼���ǣ��๦����ʽ�����ǣ� Luminex 200
'����:
'����:�¶�
'�û�:��úҽԺ
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
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strDate As String, strFilePath As String
    
    Dim objStream As TextStream, strLine As String, varItemName As Variant, varItemValue As Variant
    Dim strNO As String, strItem As String, i As Integer
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If strLine Like "Date:*" Then
            strDate = Mid$(strLine, 6)
            If IsDate(strDate) Then
                strDate = Format(CDate(strDate), "yyyy-MM-dd")
            Else
                strDate = Format(Now, "yyyy-MM-dd")
            End If
        ElseIf strLine Like "Time:*" Then
            strDate = strDate & " " & Mid$(strLine, 6)
        ElseIf strLine Like "Start*" Then
            
        ElseIf strLine Like "ItemName*" Then
            varItemName = Split(strLine, "#")
            
        ElseIf InStr(strLine, "#") > 0 Then
            varItemValue = Split(strLine, "#")
            strItem = "": strNO = ""
            For i = LBound(varItemName) To UBound(varItemName)
                If i = LBound(varItemName) Then
                    strNO = Val(varItemValue(i))
                    If strNO = "0" Then strNO = ""
                Else
                    If Trim(varItemName(i)) <> "" And Trim(varItemValue(i)) <> "" Then
                        strItem = strItem & "|" & varItemName(i) & "|" & varItemValue(i)
                    End If
                End If
            Next
            If strNO <> "" And strItem <> "" Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = strDate & "|" & strNO & "| |B|0" & strItem
            End If
        ElseIf strLine Like "End*" Then
        End If
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "MB5", strFile & " " & strLine, Err.Number & " " & Err.Description
End Function


