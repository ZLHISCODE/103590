VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_OLYMPUS_BX35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'����:���ְ�˹(OLYMPUS) BX35 ��΢��
'����:���ְ�˹(OLYMPUS)
'����:����
'�û�:������ҽ������ҽԺ
'����:2014-12-22 09:53:13
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
    Dim OBJFILE As New FileSystemObject
    Dim STRFILENAME As String, strPath As String, objStream As textStream, strFilePath As String
    Dim rsResults() As String, strImage As String
    Dim pathstr As String
On Error GoTo hErr

    If Not OBJFILE.FileExists(App.Path & "\" & "OLYMPUS_BX35.ini") Then '��ȡ���ݿ���Ϣ
        Set objStream = OBJFILE.CreateTextFile(App.Path & "\" & "OLYMPUS_BX35.ini")
            objStream.WriteLine "[OLYMPUS_BX35]"
            objStream.WriteLine "filePath=" & App.Path
            objStream.Close
            Set objStream = Nothing
    End If
    strPath = ReadIni("OLYMPUS_BX35", "filePath", App.Path & "\" & "OLYMPUS_BX35.ini")
    If strPath = "" Or Right$(strPath, 1) = "\" Then
        MsgBox "�����ļ����ô����������С�\�����ţ�", vbInformation, "�ӿ�"
    End If
    ReDim rsResults(0)
    rsResults(0) = ""
    STRFILENAME = Dir(strPath & "\", vbDirectory)       'ȡ��һ��
    Do While STRFILENAME <> ""                    '��Ϊ�վ�ѭ��
       If STRFILENAME <> "." And STRFILENAME <> ".." Then   '������� ��ǰĿ¼����һ��Ŀ¼ . �� ..
             If (GetAttr(strPath & "\" & STRFILENAME) And vbDirectory) = vbDirectory Then        '�Ƿ���Ŀ¼
'                Debug.Print strFileName          'ȷ����Ŀ¼,���ӵ��б�,��������Լ�����
                strDate = Split(STRFILENAME, "_")(0)
                strDate = Format(Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Right(strDate, 2) & " " & Time, "yyyy-mm-dd HH:mm:ss")
                strSample = Split(STRFILENAME, "_")(1)
                If OBJFILE.FileExists(strPath & "\" & STRFILENAME & "\report.txt") Then '�����Ƿ���ڴ��ļ�
                    Set objStream = OBJFILE.OpenTextFile(strPath & "\" & STRFILENAME & "\report.txt")
                    Do Until objStream.AtEndOfStream
                        strLine = objStream.ReadLine
                        If InStr(strLine, "<pcode") > 0 Then
                            strChnnl = Split(Split(strLine, ";")(2), "=")(1)
                            strData = Split(Split(strLine, ";")(3), "=")(1)
                            strIteam = strIteam & "|" & strChnnl & "|" & strData
                        End If
                        If InStr(strLine, "<image") > 0 Then
                            strLine = Replace$(strLine, "<", "")
                            strLine = Replace$(strLine, ">", "")
                            strImage = strImage & "^  ;113;0;" & Split(Split(strLine, " ")(1), "=")(1)
                            strImage = Replace(strImage, vbCr, "")
                        End If
                    Loop
                    If strIteam <> "" And strImage <> "" And strSample <> "" Then
                        If rsResults(UBound(rsResults)) <> "" Then ReDim Preserve rsResults(UBound(rsResults) + 1)
                        rsResults(UBound(rsResults)) = strDate & "|" & strSample & "| |Ѫ��|0" & strIteam
                        strIteam = ""
                        If strImage <> "" Then '����ͼƬ
                            rsResults(UBound(rsResults)) = rsResults(UBound(rsResults)) & vbCrLf & strImage
                            strImage = ""
                        End If
                    End If
                End If
            End If
         End If
         STRFILENAME = Dir      'ȡ��һ��
    Loop
    ResultFromFile = rsResults
    Exit Function
hErr:
    MsgBox "zlCDLISDev|clsLISDev_File_OLYMPUS_BX35|" & Err.Description, vbInformation, "�ӿڴ���"
    Err.Clear
End Function