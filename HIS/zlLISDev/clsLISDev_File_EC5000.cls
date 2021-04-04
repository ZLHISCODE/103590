VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EC5000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:尿分仪 EC-5000
'厂商:重庆云帆
'作者:陈东
'用户:曲沃县人民医院(太原)

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

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
    
        Dim aResults() As String
        Dim cnAccess As New ADODB.Connection
        Dim strFilePath As String, objFile As New Scripting.FileSystemObject
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim strBh As String, strItem As String, i As Integer
        On Error GoTo hErr
    
100     ReDim aResults(0)
102     strFilePath = objFile.GetParentFolderName(strFile)
104     If Len(strFilePath) = 0 Then strFilePath = App.Path
    
106     If Dir(strFile) = "" Then
108         ResultFromFile = aResults
            Exit Function
        End If
110     With cnAccess
112         .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
114         .Open
116         If .State = adStateClosed Then
118             ResultFromFile = aResults
                Exit Function
            End If
120         strSQL = "" & _
                    "SELECT Mid(A.标本号,1,8) AS 日期, Mid(A.标本号,9) AS 标本号, A.红细胞, A.白细胞, A.精子计数, A.上皮细胞, A.结晶, A.滴虫数量, A.真菌计数, A.管型, A.粘液丝, A.颜色, A.清晰度, A.尿胆原, A.胆红素, A.酮体, A.潜血, A.蛋白质, A.亚硝酸盐, A.[白细胞(干化)], A.葡萄糖, A.比重, A.PH值" & _
                    " FROM TestResult AS A Where Mid(A.标本号,1,8)>='" & Format(dtStart, "yyyyMMdd") & "' And Mid(A.标本号,1,8)<= '" & Format(dtEnd, "yyyyMMdd") & "' " & _
                    IIf(lngSampleNO = -1, "", " And val(Mid(A.标本号,9))=" & lngSampleNO)
                    
122         Set rsTmp = cnAccess.Execute(strSQL)
124         If rsTmp.EOF Then
126             .Close
128             ResultFromFile = aResults
                Exit Function
            End If
 
130         strItem = ""
132         Do While Not rsTmp.EOF
134             strItem = ""
136             strItem = Mid("" & rsTmp("日期"), 1, 4) & "-" & Mid("" & rsTmp("日期"), 5, 2) & "-" & Mid("" & rsTmp("日期"), 7, 2) '检验时间
138             strItem = strItem & "|" & Val("" & rsTmp("标本号"))   '样本序号
140             strItem = strItem & "| "  '检验人
142             strItem = strItem & "|血液"  '标本
144             strItem = strItem & "|0"
                
146             For i = 2 To rsTmp.Fields.count - 1
148                 strBh = Trim("" & rsTmp.Fields(i).Value)
150                 If strBh <> "" Then
152                     If InStr(strBh, "+") > 0 Or InStr(strBh, "-") > 0 Then
154                         strBh = Trim(Mid(strBh, 1, 2))
                        End If
156                     strItem = strItem & "|1" & Format(i, "00") & "|" & strBh
                    End If
                Next
158             If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
160             aResults(UBound(aResults)) = strItem
                
162             rsTmp.MoveNext
            Loop
                
        End With
164     ResultFromFile = aResults
        Exit Function
hErr:
166     WriteLog "EC5000", CStr(Erl()) & ":" & Err.Description, ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    
End Sub






