VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_iRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:iRead
'����:δ֪
'����:�¶�
'�û�:����ҽԺ

'�ļ���ΪiRead.txt,�ļ����̶�
'

'�ļ���������
'1001,CRP500,<2.50,2014-11-05 19:23:17
'1002,CRP500,<2.50,2014-11-05 20:02:29
'1003,CRP500,<2.50,2014-11-05 20:08:14
'1004,CRP500,<2.50,2014-11-05 20:30:13
'100000159175,CRP500,<5.0,2014-11-06 11:56:22
'100000159175,HsCRP,<2.50,2014-11-06 11:56:22
'100000173653,CRP500,49.45,2014-11-08 11:13:14
'100000173653,HsCRP,>5.0,2014-11-08 11:13:14
'100000173655,CRP500,<5.0,2014-11-08 11:16:56
'100000173655,HsCRP,<2.50,2014-11-08 11:16:56
'1001,CRP500,<5.0,2014-11-08 11:17:56
'1001,HsCRP,<2.50,2014-11-08 11:17:56


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


Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
        Dim strChenel As String, strDate As String
        Dim strFilePath As String, strTmp As String
    
        Dim objStream As textStream
        Dim strLine As String, varItem As Variant, varTmp As Variant
        Dim strSampleNo As String, strFileName As String
        Dim lngRow As Long, varLine As Variant
    
 
        On Error GoTo errHandle
        '��������±����
100     ReDim aResults(0)
102     aResults(0) = ""
104     ResultFromFile = aResults

        
106     strFilePath = objFile.GetParentFolderName(strFile)
108     If Len(strFilePath) = 0 Then strFilePath = App.Path
110     If Right$(strFilePath, 1) = "\" Then strFilePath = Mid$(strFilePath, 1, Len(strFilePath) - 1)
        
112     strFileName = strFile
114     If objFile.FileExists(strFileName) Then
116         Set objStream = objFile.OpenTextFile(strFileName, ForReading)
        
118         Do Until objStream.AtEndOfStream
120             strLine = objStream.ReadLine
122             If InStr(strLine, ",") > 0 Then
124             varLine = Split(strLine, ",")

    
126                 strSampleNo = Trim(varLine(0))
128                 If Val(strSampleNo) > 0 And UBound(varLine) >= 3 Then
130                     strDate = Trim$("" & varLine(3))
132                     If Not IsDate(strDate) Then strDate = Format(Now, "yyyy-MM-dd")
                    
134                     strChenel = Trim$(varLine(1))
                        
136                     strTmp = Trim$(varLine(2))
                        
138                     If strChenel <> "" And strTmp <> "" Then
140                         strChenel = strChenel & "|" & strTmp
    
142                         If aResults(UBound(aResults)) <> "" And strDate <> "" Then
144                             ReDim Preserve aResults(UBound(aResults) + 1)
                            End If
146                         aResults(UBound(aResults)) = strDate & "|" & strSampleNo & "| |Boold|0|" & strChenel
                        End If
     
                    End If
                End If
148             lngRow = lngRow + 1
150             DoEvents
            Loop
152         ResultFromFile = aResults
        Else
154          WriteLog " iRead,�ļ�" & strFileName & "������!", "", ""
        End If
        Exit Function
errHandle:
 
156     WriteLog " iRead," & Erl() & Err.Description, "", ""
End Function
