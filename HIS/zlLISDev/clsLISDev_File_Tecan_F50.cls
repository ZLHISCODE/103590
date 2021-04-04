VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Tecan_F50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:ȫ�Զ�ø�����߹���վ��Tecan Freedom EVOlyzer��
'����:��ʿ�ۿ�TECAN
'����:���
'�û�:��������ҽԺ
'�ͺ�:F50


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
    Dim strFilePath As String, strFilename As String
    Dim i As Integer
    Dim objStream As textStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str�걾�� As String, str������ As String, blnReturn As Boolean
    Dim iRow As Integer '�к�
    Dim varLine As Variant
    On Error GoTo errHandle
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    str���� = Format(dtStart, "yyyy-MM-dd")

    strFilename = Dir(strFilePath & "\*.asc")
    Do While strFilename <> ""
        Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename, ForReading)
        Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            varLine = Split(strLine, Chr(&H9))
            If InStr(1, Trim(varLine(0)), "mth") > 0 Then
                strChenel = Mid(Trim(varLine(0)), 1, InStr(1, Trim(varLine(0)), "mth") - 2)
            End If
        Loop
        objStream.Close
        Set objStream = Nothing
        Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename, ForReading)
        Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            varLine = Split(strLine, Chr(&H9))
            
            If Trim(varLine(0)) <> "" And InStr(1, Trim(varLine(0)), "mth") = 0 Then
                '�걾�ţ�������
                str�걾�� = Trim(varLine(3))
                '������ģʽ  str���Խ�� & "^" & str������� & "^" & strCO & "^0"
                 
                str������ = Trim(Trim(varLine(2))) & "^" & Trim(Trim(varLine(0))) & "^" & Trim(Trim(varLine(4))) & "^" & Trim(Trim(varLine(1)))
                
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
        objFile.CopyFile strFilePath & "\" & strFilename, strFilePath & "\" & Replace(str����, "-", "") & "\" & strFilename, True
        objFile.DeleteFile strFilePath & "\" & strFilename
        strFilename = Dir(strFilePath & "\*.asc")
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_EVOlyzer.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function