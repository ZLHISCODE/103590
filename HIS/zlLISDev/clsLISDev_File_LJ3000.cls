VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LJ3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
Option Explicit
'����:����� LJ-3000
'����:�����������＼�����޹�˾
'����:�����
'�û�:��������ҽԺ
'ʱ��:2015-07-21

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
    
    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long
    Dim strIniFile As String, strImg As String, lngPath As Long, strPath As String, strFileName As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.textStream
    Dim SERVERName As String, DatabaseName As String, USERName As String, PWD As String, PICADRESS As String 'SQL SERVER�����ַ���
    
    On Error GoTo errHandle
    strIniFile = App.Path & "\LJ3000.ini"
    
    If Not objFileSystem.FileExists(strIniFile) Then
        Set objStream = objFileSystem.CreateTextFile(strIniFile)
        objStream.WriteLine "[Options]"
        objStream.WriteLine "[DATABASE]"
        objStream.WriteLine "SERVERNAME=."
        objStream.WriteLine "DATABASENAME=jnlj_db"
        objStream.WriteLine "USERNAME=SA"
        objStream.WriteLine "PASSWORD=123456"
        objStream.WriteLine "PICADRESS=D:\Program Files\���������������������\Pictures"
        objStream.Close
        Set objStream = Nothing
    End If
    
    '-----SQL SERVER�����ַ���
    SERVERName = ReadIni("DATABASE", "SERVERNAME", strIniFile)
    DatabaseName = ReadIni("DATABASE", "DATABASENAME", strIniFile)
    USERName = ReadIni("DATABASE", "USERNAME", strIniFile)
    PWD = ReadIni("DATABASE", "PASSWORD", strIniFile)
    PICADRESS = ReadIni("DATABASE", "PICADRESS", strIniFile)
    
    
    If (Trim(SERVERName) = "") Or (Trim(DatabaseName) = "") Or (Trim(USERName) = "") Then
         WriteLog "BH5500S_sql_ResultFromFile", strIniFile, "����SQL SERVER���ݿ����!����ԭ�򣺷�������Ϊ�ջ����ݿ���Ϊ�ջ��û���Ϊ��!"
         Exit Function
    End If
    
    With cnAccess
        .ConnectionString = "Provider=SQLOLEDB.1;Initial Catalog=" & DatabaseName & ";Password=" & PWD & _
        ";Persist Security Info=True;User ID=" & USERName & ";Data Source=" & SERVERName
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "select id as  �걾��,Date  As ��������,YinXue  as ��Ѫ������," & vbNewLine & _
            " Nxb, Zfd , Dfkl, Mj, Zwxb,Jj,Spxb, Jscyc, Jdzz ,Jmyj ,Jscl ,Jxw, Lzbd,rbc_value , wbc_value As wbc_value, tsbc_value As tsbc_value" & vbNewLine & _
            "from jnlj_table where Date>='" & Format(dtStart, "yyyy-mm-dd") & "' and Date<='" & Format(dtEnd, "yyyy-mm-dd") & "' and  yq_do=1"
''yq_do   �����Ƿ���    0��δ��� 1���Ѽ��
''ys_do   ҽ���Ƿ����    0��δ��� 1�������
''yx_test �걾�Ƿ�����Ѫ����  0�������� 1�� ����
''lzbd_test   �걾�Ƿ�����״��������  0�������� 1�� ����
      
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
             .Close
             Exit Function
        End If
        lngResultCount = -1
        
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            
            aResults(lngResultCount) = Format(rsTmp("��������"), "yyyy-MM-dd hh:mm:ss") & "|" & Trim(rsTmp("�걾��")) & "| |���|0|Nxb|" & Decode(rsTmp("Nxb"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Nxb"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Zfd|" & Decode(rsTmp("Zfd"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Zfd"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Dfkl|" & Decode(rsTmp("Dfkl"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Dfkl"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Mj|" & Decode(rsTmp("Mj"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Mj"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Zwxb|" & Decode(rsTmp("Zwxb"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Zwxb"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jj|" & Decode(rsTmp("Jj"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jj"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Spxb|" & Decode(rsTmp("Spxb"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Spxb"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jscyc|" & Decode(rsTmp("Jscyc"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jscyc"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jdzz|" & Decode(rsTmp("Jdzz"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jdzz"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jmyj|" & Decode(rsTmp("Jmyj"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jmyj"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jscl|" & Decode(rsTmp("Jscl"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jscl"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jxw|" & Decode(rsTmp("Jxw"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jxw"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Lzbd|" & Decode(rsTmp("Lzbd"), "0", "δ��", "1", "ż��", "2", "����", "3", "ȫ��Ұ", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Lzbd"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|rbc_value|" & rsTmp("rbc_value") & "|wbc_value|" & rsTmp("wbc_value") & "|tsbc_value|" & rsTmp("tsbc_value")
            'ȡͼƬ·�� strFilePath
            
            strFilePath = PICADRESS & "\" & rsTmp("��������") & "\" & rsTmp("�걾��") & "\" '"2015-07-22" & "\" & 1 & "\" '
            lngPath = InStrRev(strFilePath, "\") '�ҵ����һ��\
            strPath = Mid$(strFilePath, 1, lngPath)
            strFileName = Dir(strPath, vbDirectory)
            Do While strFileName <> ""
                If strFileName <> "." And strFileName <> ".." Then
                    If GetAttr(strPath & "\" & strFileName) = vbDirectory Or GetAttr(strPath & "\" & strFileName) = vbArchive Then '�Ƿ���Ŀ¼
                        If strFileName Like "*.bmp" Then    '����ͼƬ
                            strImg = strImg & "^" & "P" & Replace(Split(strFileName & ".", ".")(0), "-", "") & ";103;1;" & strPath & strFileName
                        End If
                    End If
                End If
                strFileName = Dir
            Loop
            If strImg <> "" Then
                aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & Mid(strImg, 2)
                strImg = ""
            End If
            
            Debug.Print aResults(UBound(aResults))
            Call WriteLog(" aResults(lngResultCount)", aResults(lngResultCount), "")
            rsTmp.MoveNext
        Loop
    End With
    
    If lngResultCount = -1 Then
        ReDim Preserve aResults(0)
    End If
    
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "LJ-3000_ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function


Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

