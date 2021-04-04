VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_XN2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:sysmex Ѫ����
'����:sysmex
'����:���
'�û�:��������������ҽԺ
'�ͺ�:XN2000

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
    Dim str_RetuImage As String
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

    strFilename = Dir(strFilePath & "\*.CDF")
    Do While strFilename <> ""
        Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename, ForReading)
        i = 1
        str_RetuImage = ""
        str������ = ""
        Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            varLine = Split(strLine, ",")
            If varLine(0) = "0" Then
                str�걾�� = varLine(3)
            End If
            If varLine(0) = "00" Then
                str���� = varLine(19)
            End If
            If varLine(0) = 1 Then
                str������ = str������ & "|" & varLine(1) & "|" & varLine(3)
            End If
            If varLine(0) = 3 Then
                str_RetuImage = str_RetuImage & "^ " & varLine(2) & ";103;0;" & varLine(3)
            End If
            i = i + 1
        Loop
        If Val(str�걾��) > 0 And str������ <> "" Then
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |ѪҺ|0" & str������ & vbCrLf & Mid(str_RetuImage, 2)
        End If
        objStream.Close
        Set objStream = Nothing
        If Not objFile.FolderExists(strFilePath & "\" & Trim(Replace(Mid(str����, 1, 10), "-", ""))) Then objFile.CreateFolder (strFilePath & "\" & Trim(Replace(Mid(str����, 1, 10), "-", "")))
        objFile.CopyFile strFilePath & "\" & strFilename, strFilePath & "\" & Trim(Replace(Mid(str����, 1, 10), "-", "")) & "\" & strFilename, True
        objFile.DeleteFile strFilePath & "\" & strFilename
        strFilename = Dir(strFilePath & "\*.CDF")
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_EVOlyzer.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function