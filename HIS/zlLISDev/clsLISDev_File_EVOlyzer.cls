VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EVOlyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:ȫ�Զ�ø�����߹���վ��Tecan Freedom EVOlyzer��
'����:��ʿ�ۿ�TECAN
'����:�¶�
'�û�:��������ҽԺ
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
    Dim strChenel As String, str����Ա As String, str���� As String
    Dim strFilePath As String, strFileName As String
    
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim iRow As Integer '�к�
    Dim varline As Variant
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    str���� = Format(dtStart, "yyyy-MM-dd")
    If frm����.ShowMe(str����) = False Then Exit Function
        

    strFileName = Dir(strFilePath & "\*.asc")
    Do While strFileName <> ""
        Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFileName, ForReading)
        Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            varline = Split(strLine, Chr(&H9))
            If UBound(varline) >= 5 Then
                '�걾�ţ�������
                
                str�걾�� = Val(Trim(varline(1)))
                strChenel = Trim(varline(0))
                str������ = Trim(Trim(varline(2))) & "^" & Trim(Trim(varline(3))) & "^" & Trim(Trim(varline(5))) & "^" & Trim(Trim(varline(4)))
                
                If Not IsDate(str����) Then str���� = Format(Now, "yyyy-MM-dd")
                
                If Val(str�걾��) > 0 And str������ <> "" And strChenel <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |ѪҺ|0|" & strChenel & "|" & str������
                End If
            End If
        Loop
        objStream.Close
        Set objStream = Nothing
        
        If Not objFile.FolderExists(strFilePath & "\" & Replace(str����, "-", "")) Then objFile.CreateFolder (strFilePath & "\" & Replace(str����, "-", ""))
        objFile.CopyFile strFilePath & "\" & strFileName, strFilePath & "\" & Replace(str����, "-", "") & "\" & strFileName, True
        objFile.DeleteFile strFilePath & "\" & strFileName
        strFileName = Dir(strFilePath & "\*.asc")
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_EVOlyzer.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function


