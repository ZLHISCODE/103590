VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_File_FJ2008PS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�÷������߼�����FJ-2008PS
'����:��Ӫ��������
'����:�ܺ�ȫ
'�û�:����������ҽԺ
'ʱ��:2012-03-11
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
'       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������("^" & ODֵ & "^" & COֵ & "^" & S/Co)
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim str�걾 As String, str����Ա As String, str���� As String, strDate As String
    Dim strFilePath As String, strNewFile As String
    
    Dim objStream As TextStream, aItem() As String
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str������ As String, str������ As String, blnReturn As Boolean
    
    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    '��ʼ�걾����
    ReDim aItem(0): aItem(0) = ""
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    '2012-03-11 ZHQ ���豸���ļ����̶�ΪSendData.txt
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    str������ = ""
    str������ = ""
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        
        '��׼��ʽ�����£�
        '    1    ��2MG              3.27     120221    1.75--2.61   mg/L
        '    2    ��2MG              2.37     120221    1.75--2.61   mg/L
        '    3    ��2MG              4.57     120221    1.75--2.61   mg/L

        str���� = Trim(Mid(strLine, 36, 8))
        str���� = Mid(Format(date, "YYYY"), 1, 2) & str����
        str���� = Mid(str����, 1, 4) & "-" & Mid(str����, 5, 2) & "-" & Mid(str����, 7, 2)
            
        str������ = Trim(Mid(strLine, 1, 5))
        If Not IsNumeric(str������) Then str������ = 9999
        
        varItem = Trim(Mid(strLine, 6, 8))
        varTmp = Trim(Mid(strLine, 16, 18))
        str������ = "|" & varItem & "|" & varTmp
            
        If aResults(UBound(aResults)) <> "" Then
            ReDim Preserve aResults(UBound(aResults) + 1)
        End If
        If str������ <> "" Then
            aResults(UBound(aResults)) = str���� & "|" & str������ & "| |ѪҺ|0" & str������
        End If
        DoEvents
    Loop

    ResultFromFile = aResults
End Function

