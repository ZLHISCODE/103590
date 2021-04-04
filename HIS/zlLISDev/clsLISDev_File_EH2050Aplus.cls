VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EH2050Aplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:���� EH-2050a plus �����
'����:����
'����:����
'�û�:˫���صڶ�����ҽԺ
'����:2015-01-06

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
    Dim strCode As String, strValue() As String, strDate As String, strIteam As String, strSample As String, strLine As String, strData As String, strChnnl As String
    Dim lngS As Long, lngE As Long, i As Integer
    Dim objFile As New FileSystemObject
    Dim strFileName As String, strPath As String, objStream As textStream, strFilePath As String
    Dim rsResults() As String, strImage As String, strImageName As String
    Dim strStat As String '�жϼ����ʶ
On Error GoTo hErr
    
   strPath = Mid(strFile, 1, InStrRev(strFile, "\"))
    ReDim rsResults(0)
    rsResults(0) = ""
    strFileName = Dir(strPath & "\", vbDirectory)       'ȡ��һ��
    Do While strFileName <> ""                    '��Ϊ�վ�ѭ��
       If strFileName <> "." And strFileName <> ".." Then   '������� ��ǰĿ¼����һ��Ŀ¼ . �� ..
             If GetAttr(strPath & "\" & strFileName) = vbDirectory Or GetAttr(strPath & "\" & strFileName) = vbArchive Then         '�Ƿ���Ŀ¼
'                If Format(Now, "yyyymmdd") = Left(Split(strFileName, ".")(0), 8) Then '�������ݽ϶࣬Ŀǰֻ������ǰ���ڵ�����
                If InStr(strFileName, Format(Now, "yyyymmdd")) Then
                    If objFile.FileExists(strPath & "\" & strFileName) Then  '�����Ƿ���ڴ��ļ�
                        Set objStream = objFile.OpenTextFile(strPath & "\" & strFileName)
                        Do Until objStream.AtEndOfStream
                            strLine = objStream.ReadLine
                            If InStr(strLine, "eBaseInfo") > 0 And InStr(strLine, "������") > 0 Then 'ȡ������
                                strSample = Split(strLine, ",")(3)
                                strSample = Replace$(strSample, ";", "")
                                strSample = Replace$(strSample, Chr(&H9), "")
                                strSample = Val(strSample)
                            End If
                            If InStr(strLine, "eBaseInfo") > 0 And InStr(strLine, "����ʱ��") > 0 Then 'ȡʱ��
                                strDate = Split(strLine, ",")(3)
                                strDate = Replace$(strDate, ";", "")
                                strDate = Replace$(strDate, Chr(&H9), "")
                                strDate = Format(strDate, "yyyy-mm-dd hh:mm:ss")
                            End If
                            If InStr(strLine, "eBaseInfo") > 0 And InStr(strLine, "����") > 0 Then '�жϼ���
                                strStat = Split(strLine, ",")(3)
                                strStat = Replace$(strStat, ";", "")
                                strStat = Replace$(strStat, Chr(&H9), "")
                                strStat = Trim(strStat)
                            End If
                            If InStr(strLine, "eBaseInfo") > 0 And InStr(strLine, "��ɫ") > 0 Then
                                strChnnl = "��ɫ"
                                strData = " " & Trim(Replace(Replace(Split(strLine, ",")(3), ";", ""), Chr(&H9), ""))
                                strIteam = strIteam & "|" & strChnnl & "|" & strData
                            End If
                            If InStr(strLine, "eBaseInfo") > 0 And InStr(strLine, "�Ƕ�") > 0 Then
                                strChnnl = "�Ƕ�"
                                strData = " " & Trim(Replace(Replace(Split(strLine, ",")(3), ";", ""), Chr(&H9), ""))
                                strIteam = strIteam & "|" & strChnnl & "|" & strData
                            End If
                            If InStr(strLine, "eDryInfo") > 0 And UBound(Split(strLine, ",")) >= 5 Then '��ȡ��ɻ�ѧ����
                                strChnnl = Trim(Replace(Split(strLine, ",")(2), Chr(&H9), ""))
                                strData = " " & Trim(Replace(Split(strLine, ",")(3), Chr(&H9), ""))
                                strIteam = strIteam & "|" & Replace(strChnnl, "��ϸ��", "��ϸ��1") & "|" & strData
                            End If
                            If InStr(strLine, "eNCZInfo") > 0 And UBound(Split(strLine, ",")) >= 5 Then '��ȡ�򾵼���
                                strChnnl = Trim(Replace(Split(strLine, ",")(2), Chr(&H9), ""))
                                strData = " " & Trim(Replace(Split(strLine, ",")(3), Chr(&H9), ""))
                                strIteam = strIteam & "|" & strChnnl & "|" & strData
                            End If
                            If InStr(strLine, "eImage") > 0 Then '��ȡ��ͼ������
                                strImageName = Split(Split(strLine, ",")(1), "=")(1)
                                strImage = strImage & "^" & strImageName & ";113;0;" & Replace(Split(strLine, ",")(2), ";", "")
                            End If
                        Loop
                        If strIteam <> "" And strImage <> "" And strSample <> "" Then
                            strIteam = Replace(strIteam, "leu/uL", "")
                            strIteam = Replace(strIteam, "mmol/L", "")
                            strIteam = Replace(strIteam, "g/L", "")
                            strIteam = Replace(strIteam, "ery/uL", "")
                            strIteam = Replace(strIteam, "��/��L", "")
                            If rsResults(UBound(rsResults)) <> "" Then ReDim Preserve rsResults(UBound(rsResults) + 1)
                            rsResults(UBound(rsResults)) = strDate & "|" & IIf(strStat = "False", strSample, strSample & "^1^") & "| |��Һ|0" & strIteam
                            strIteam = ""
                            If strImage <> "" Then '����ͼƬ
                                strImage = Mid(strImage, 2)
                                strImage = Replace$(strImage, "jpg", "bmp")
                                rsResults(UBound(rsResults)) = rsResults(UBound(rsResults)) & vbCrLf & Replace(strImage, "Image", "")
                                Call writelislog(rsResults(UBound(rsResults)))
                                strImage = ""
                            End If
                        End If
                    End If
                End If
            End If
         End If
         strFileName = Dir      'ȡ��һ��
    Loop
    ResultFromFile = rsResults
    Exit Function
hErr:
    MsgBox "zlCDLISDev.clsLISDev_File_EH-2050Aplus|ResultFromFile" & Err.Description, vbInformation, "�ӿڴ���"
    Err.Clear
End Function

Public Function writelislog(strLog As String)
    Const strFile As String = "LisLog_EH2080C_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFileName As String
    strFileName = App.Path & "\" & strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function

