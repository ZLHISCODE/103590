VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_FIle_LTS_E100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:LTS-E100����������
'����:�麣����
'����:�콭
'�û�:�����еڶ�����ҽԺ
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo errHand '�򿪴�������
    '------------------------------------------------
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
    Dim strSQL As String, rsNO As ADODB.Recordset, rsItem As ADODB.Recordset, rsImage As ADODB.Recordset
    Dim strBh As String, strAnti As String, str��� As String, strImage As String, str���� As String, str�걾��
    Dim lngResultCount As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};pwd=gsOmlipo"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = "SELECT ��Ӧ��Ʒ��,����ʱ��,��Ʒ���"
        strSQL = strSQL & " FROM Cases "
        strSQL = strSQL & " Where  Cases.����ʱ�� between CDate('" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "') and cdate( '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') Order BY ��Ӧ��Ʒ��"
                      
        Set rsNO = cnAccess.Execute(strSQL)
        If rsNO.EOF Then
            .Close
            Exit Function
        End If
        
         ReDim aResults(0)
        lngResultCount = 0
        
        Do While (Not rsNO.EOF)
           
            If Val("" & rsNO("��Ӧ��Ʒ��")) > 0 Then '�±걾
                
                str���� = Replace("" & rsNO("����ʱ��"), "/", "-") '����ʱ��
                str��� = ""
                
                strSQL = "SELECT ��Ʒ���,ʵ����Ŀ����,���"
                strSQL = strSQL & " FROM SpecimentCheckResult "
                strSQL = strSQL & " Where  SpecimentCheckResult.��Ʒ��� ='" & rsNO("��Ʒ���") & " '"
                              
                Set rsItem = cnAccess.Execute(strSQL)
                If rsItem.EOF Then
                    GoTo nextno:
                End If
                Do While (Not rsItem.EOF)
                    str��� = str��� & "|" & rsItem("ʵ����Ŀ����") & "|" & rsItem("���")
                    rsItem.MoveNext
                Loop
               
                strBh = "***"
        
                strBh = "" & rsNO("��Ӧ��Ʒ��")
                If strBh <> "***" And str��� <> "" Then
                    str�걾�� = "" & rsNO("��Ӧ��Ʒ��")
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str����
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(str�걾��) '�������
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|Ѫ��"
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & str���
                End If
                strSQL = "SELECT ��Ʒ���,ͼƬ����,ͼƬ·��"
                strSQL = strSQL & " FROM SpecimentResult "
                strSQL = strSQL & " Where  SpecimentResult.��Ʒ��� ='" & rsNO("��Ʒ���") & " '"
                              
                Set rsImage = cnAccess.Execute(strSQL)
                If rsImage.EOF Then
                    GoTo nextno:
                End If
                
                strImage = ""
                Do While (Not rsImage.EOF)
                
                    strImage = strImage & "^" & rsImage("ͼƬ����") & ";110;0;" & rsImage("ͼƬ·��")
                    
                    rsImage.MoveNext
                Loop
                If strImage <> "" Then
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strImage, 2)
                End If
            End If
nextno:
            rsNO.MoveNext
        Loop
        
    End With
    Dim i As Integer
    For i = 0 To UBound(aResults)
        Debug.Print aResults(i)
    Next
    ResultFromFile = aResults
   
    '------------------------------------------------
    Exit Function
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\YFErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_FIle_LTS_E100"
    Close #lngErrFileNo
End Function










