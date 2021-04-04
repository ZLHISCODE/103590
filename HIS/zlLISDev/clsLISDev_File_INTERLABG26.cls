VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_INTERLABG26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����: ��Ӿ��
'����:
'����: ����
'�û�: �󰲸��ױ���Ժ ��ȡ����������

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
    
    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection, rsTmp As New ADODB.Recordset, strSQL As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject, strFilename As String, txtStream As textStream
    Dim strLine As String, arrItems() As String
    Dim strID As String, strDate As String, strValue As String, strCode As String, strItem As String
    Dim intDay As Integer, intLop As Integer, i As Integer
    Dim strIniFile As String, Mode As Integer
    Dim str�걾�� As String
    
    On Error GoTo hErr
    ReDim aResults(0)
    aResults(0) = ""
    strIniFile = App.Path & "\INTERLABG26.ini"
    If Not objFile.FileExists(strIniFile) Then
       Set txtStream = objFile.CreateTextFile(strIniFile)
       txtStream.WriteLine "[Options]"
       txtStream.WriteLine "TYPE=1      ;   1:��ȡ��������,2:��ȡMDB���ݿ�"
       txtStream.Close
       Set txtStream = Nothing
    End If
    Mode = Val(ReadIni("Options", "TYPE", strIniFile, 0))
    If strFile <> "" Then strFilePath = objFile.GetParentFolderName(strFile)
    intDay = Abs(DateDiff("d", dtStart, dtEnd))
    If Mode = 1 Then
        For intLop = 0 To intDay
            strFilename = Format(dtStart + intLop, "yymmdd") & ".txt"
            If Not objFile.FileExists(strFilePath & "\" & strFilename) Then GoTo nextLine:
            Set txtStream = objFile.OpenTextFile(strFilePath & "\" & strFilename)
            Do Until txtStream.AtEndOfStream
                strLine = txtStream.ReadLine
                If InStr(strLine, Format(dtStart + intLop, "yyyy-mm-dd")) > 0 Then
                    arrItems = Split(strLine, Chr(&H9))
                    strDate = Format(arrItems(0) & " " & Time, "yyyy-mm-dd hh:mm:ss")
                    strID = arrItems(1)
                    strCode = Split(arrItems(2), ":")(1)
                    If strCode = "" Then GoTo nextLine1:
                    strValue = Replace(Split(Split(arrItems(3), ":")(1), " ")(0), "δע��", "")
                    strItem = strItem & "|" & strCode & "|" & strValue
                End If
                If InStr(strLine, "���㵱ǰ�걾") > 0 Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = strDate & "|" & strID & "| |Ѫ��|0" & strItem
                    strItem = ""
                End If
nextLine1:
            Loop
    '        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
    '        aResults(UBound(aResults)) = strDate & "|" & strID & "| |Ѫ��|0" & strItem
    '        strItem = ""
nextLine:
        Next
    Else
        If strFile <> "" Then strFilePath = objFile.GetParentFolderName(strFile)
        With cnAccess
            .CursorLocation = adUseClient
            .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
            .Open
            If .State = 0 Then
                Call WriteLog("������־", "", "δ�ܴ�MDB���ݿ����ӣ�")
                Exit Function
            End If
        End With
        strSQL = "select idpaz,[data esame] as date1,fraz1 as HbA,fraz2 as HbA2,fraz3 as HbF,fraz4 as HbS from archivio where [Data esame] " & _
            "between format('" & Format(dtStart, "yyyy-mm-dd") & "','yyyy-mm-dd') and  format('" & Format(dtEnd, "yyyy-mm-dd") & "','yyyy-mm-dd')"
        Set rsTmp = cnAccess.Execute(strSQL)
        Do Until rsTmp.EOF
            str�걾�� = rsTmp!idpaz
            strDate = Format(rsTmp!date1, "yyyy-mm-dd hh:mm:ss")
            strItem = "HbA|" & rsTmp!HbA & "|HbA2|" & rsTmp!HbA2 & "|HbF|" & rsTmp!HbF & "|HbS|" & rsTmp!HbS
            If str�걾�� <> "" And strItem <> "" Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = strDate & "|" & str�걾�� & "| |Ѫ��|0|" & strItem
            End If
            rsTmp.MoveNext
        Loop
    End If
    For i = 0 To UBound(aResults)
        Debug.Print aResults(i)
    Next
    ResultFromFile = aResults
    Exit Function
hErr:
    Call WriteLog("������־", "", Err.Description)
    Err.Clear
End Function


