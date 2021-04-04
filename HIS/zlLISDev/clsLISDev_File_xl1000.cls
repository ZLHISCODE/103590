VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_xl1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'����:Ѫ����
'ҽԺ:���ݲ�óҽԺ
'����:����ǿ
'ʱ��:2015/1/17
'����:�����ڳ� xl-100 ���ݿ�Ϊsqllite


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
        On Error GoTo ResultFromFile_Err


        Dim aResults() As String
        Dim cnsql As New ADODB.Connection
        Dim cnstring As String
        Dim strFilePath As String, objFile As New Scripting.FileSystemObject
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset
        Dim lngResultCount As Long
        Dim blnUnion As Boolean
        Dim lngSampleID As Long
        Dim blnNew As Boolean           '�Ƿ��±걾��
        Dim lngSampleNumber As Long     '�걾ID
        Dim lngloop As Long
        Dim intLoop As Integer
        Dim lng_count As Long
        Dim st���� As String, en���� As String, str���� As String
        Dim strSer As String, strDB As String, strPWD As String
        Dim strItem As String, strValue As String
        Dim str�걾�� As String, strTmp�걾�� As String
        Dim str��� As String


        '��������±����
100     ReDim aResults(0)
102     aResults(0) = ""
104     lng_count = -1
106     ResultFromFile = aResults
        '85827385
108     With cnsql
110         Set cnsql = New ADODB.Connection
112         If cnsql.State = adStateOpen Then cnsql.Close
            On Error Resume Next
114         cnstring = "DSN=lis;UID="";PSW="""    'dsnϵͳodbc,uid �û���,psw����
116         cnsql.Open cnstring
118         st���� = Format("2000-01-01", "yyyy-mm-dd")
120         en���� = Format(dtEnd, "yyyy-mm-dd")
122         If lngSampleNO > 0 Then
124             strSQL = "select a.id,a.sampleno,a.date,b.item,b.subitem,b.abbr,b.code,b.result,b.unit,b.flag,b.displayorder,lowermarklimit,uppermarklimit" & _
                       " from patient a,result b " & _
                       " where a.id=b.patientid and b.result is not null and a.Date between '" & st���� & "' and '" & en���� & "' and a.SampleCode=" & lngSampleNO & " order by a.date,a.id"
            Else
126             strSQL = "select a.id,a.sampleno,a.date,b.item,b.subitem,b.abbr,b.code,b.result,b.unit,b.flag,b.displayorder,lowermarklimit,uppermarklimit" & _
                       " from patient a,result b " & _
                       " where a.id=b.patientid and  b.result is not null and a.Date between '" & st���� & "' and '" & en���� & "' order by a.date,a.id"
            End If

128         Set rsTmp = cnsql.Execute(strSQL)
130         str�걾�� = ""
132         If rsTmp Is Nothing Then Exit Function
134         Do Until rsTmp.EOF
136             strTmp�걾�� = Val(rsTmp!sampleno)
138             If str�걾�� <> strTmp�걾�� And str�걾�� <> "" Then
140                 str���� = Format(rsTmp!date, "YYYY-MM-DD")
142                 str��� = str���� & "|" & str�걾�� & "| |Ѫ��|0" & str���
144                 If aResults(UBound(aResults)) <> "" Then
146                     ReDim Preserve aResults(UBound(aResults) + 1)
                    End If
148                 aResults(UBound(aResults)) = str���
150                 str�걾�� = strTmp�걾��
152                 strItem = rsTmp!abbr
154                 strValue = rsTmp!result
156                 str��� = "|" & Nvl(strItem, "1") & "|" & Nvl(strValue, " ")
                Else
158                 strItem = rsTmp!abbr
160                 strValue = rsTmp!result
162                 str��� = str��� & "|" & Nvl(strItem, "1") & "|" & Nvl(strValue, " ")
164                 str�걾�� = strTmp�걾��
                End If
166             rsTmp.MoveNext
            Loop
168         ResultFromFile = aResults
        End With

        Exit Function

ResultFromFile_Err:
    If Error = 1 Then
        Resume
    End If
End Function












