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


'仪器:血凝仪
'医院:徐州财贸医院
'开发:陈玉强
'时间:2015/1/17
'厂家:北京众驰 xl-100 数据库为sqllite


Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    '标准函数：从文件获取检验结果
    '   strFile：包含路径的文件名
    '   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
    '   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
    '   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
    '   返回的检验结果数组(各解析程序必须按以下标准组织结果)
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
        On Error GoTo ResultFromFile_Err


        Dim aResults() As String
        Dim cnsql As New ADODB.Connection
        Dim cnstring As String
        Dim strFilePath As String, objFile As New Scripting.FileSystemObject
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset
        Dim lngResultCount As Long
        Dim blnUnion As Boolean
        Dim lngSampleID As Long
        Dim blnNew As Boolean           '是否新标本号
        Dim lngSampleNumber As Long     '标本ID
        Dim lngloop As Long
        Dim intLoop As Integer
        Dim lng_count As Long
        Dim st日期 As String, en日期 As String, str日期 As String
        Dim strSer As String, strDB As String, strPWD As String
        Dim strItem As String, strValue As String
        Dim str标本号 As String, strTmp标本号 As String
        Dim str结果 As String


        '避免出现下标错误
100     ReDim aResults(0)
102     aResults(0) = ""
104     lng_count = -1
106     ResultFromFile = aResults
        '85827385
108     With cnsql
110         Set cnsql = New ADODB.Connection
112         If cnsql.State = adStateOpen Then cnsql.Close
            On Error Resume Next
114         cnstring = "DSN=lis;UID="";PSW="""    'dsn系统odbc,uid 用户名,psw密码
116         cnsql.Open cnstring
118         st日期 = Format("2000-01-01", "yyyy-mm-dd")
120         en日期 = Format(dtEnd, "yyyy-mm-dd")
122         If lngSampleNO > 0 Then
124             strSQL = "select a.id,a.sampleno,a.date,b.item,b.subitem,b.abbr,b.code,b.result,b.unit,b.flag,b.displayorder,lowermarklimit,uppermarklimit" & _
                       " from patient a,result b " & _
                       " where a.id=b.patientid and b.result is not null and a.Date between '" & st日期 & "' and '" & en日期 & "' and a.SampleCode=" & lngSampleNO & " order by a.date,a.id"
            Else
126             strSQL = "select a.id,a.sampleno,a.date,b.item,b.subitem,b.abbr,b.code,b.result,b.unit,b.flag,b.displayorder,lowermarklimit,uppermarklimit" & _
                       " from patient a,result b " & _
                       " where a.id=b.patientid and  b.result is not null and a.Date between '" & st日期 & "' and '" & en日期 & "' order by a.date,a.id"
            End If

128         Set rsTmp = cnsql.Execute(strSQL)
130         str标本号 = ""
132         If rsTmp Is Nothing Then Exit Function
134         Do Until rsTmp.EOF
136             strTmp标本号 = Val(rsTmp!sampleno)
138             If str标本号 <> strTmp标本号 And str标本号 <> "" Then
140                 str日期 = Format(rsTmp!date, "YYYY-MM-DD")
142                 str结果 = str日期 & "|" & str标本号 & "| |血清|0" & str结果
144                 If aResults(UBound(aResults)) <> "" Then
146                     ReDim Preserve aResults(UBound(aResults) + 1)
                    End If
148                 aResults(UBound(aResults)) = str结果
150                 str标本号 = strTmp标本号
152                 strItem = rsTmp!abbr
154                 strValue = rsTmp!result
156                 str结果 = "|" & Nvl(strItem, "1") & "|" & Nvl(strValue, " ")
                Else
158                 strItem = rsTmp!abbr
160                 strValue = rsTmp!result
162                 str结果 = str结果 & "|" & Nvl(strItem, "1") & "|" & Nvl(strValue, " ")
164                 str标本号 = strTmp标本号
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













