VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_IMARK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'����:ø���� BIO_IMARK
'����:����
'����:������
'�û�:������ҽԺ
Private mstrCode As String, str_Old As String

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
     

End Sub

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
    Dim str���� As String
    Dim strFilePath As String
    
    Dim objStream As textStream
    Dim strLine As String, varItem As Variant
    Dim str�걾�� As String
    Dim lngn, Lngm As Long
    Dim FindDate As String
    On Error GoTo errHandle
    
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    lngn = CDate(Format(Now, "YYYY-MM-DD")) - dtStart
    For Lngm = 0 To lngn
        FindDate = dtStart + Lngm
        strFile = Mid(strFile, 1, InStrRev(strFile, "\")) + Format(FindDate, "yymmdd") + ".txt"
        If Dir(strFile) <> "" Then
                strFilePath = objFile.GetParentFolderName(strFile)
            If Len(strFilePath) = 0 Then strFilePath = App.Path
            Set objStream = objFile.OpenTextFile(strFile, ForReading)
            Do Until objStream.AtEndOfStream
                strLine = objStream.ReadLine
                varItem = Split(strLine, Chr(&H9))
                If strLine <> "" Then
                    str���� = Format(FindDate, "yyyy-mm-dd")
                    str�걾�� = varItem(1)
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    If str�걾�� Like "QC*" Then '�Ƿ����ʿر걾
                        str�걾�� = 997 + Mid(str�걾��, InStr(str�걾��, ":") + 1)
                        aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|1|" & varItem(2) & "|" & "" & "^" & Mid(varItem(3), 1, 5) & "^^"
                    ElseIf Val(str�걾��) > 0 Then
                        aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |Ѫ��|0|" & varItem(2) & "|" & varItem(4) & "^" & varItem(5) & "^" & varItem(6) & "^" & varItem(5) / varItem(6)
                    End If
                End If
            Loop
            ResultFromFile = aResults
            
            objStream.Close
        End If
    Next
    Exit Function
errHandle:
    WriteLog "clsLISDev_IMARK.ResultFromFile", strFile & " " & strLine, Err.Number & " " & Err.Description
End Function

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function