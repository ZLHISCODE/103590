VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Fascow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
        Dim lngResultCount As Long, strBh As String, strLine_Man As String, strLine_Wman As String, str项目单位 As String
        Dim str低 As String, str中 As String, str高 As String, strimg As String
        Dim blnUnion As Boolean
        Dim str日期 As String
        Dim str仪器型号 As String, strItems As String, strQC As String
        Dim arrLow(2) As Double, arrMid(2) As Double, arrHigh(2) As Double, arrNianDu(2) As Double
    
        On Error GoTo hErr
100     strFilePath = objFile.GetParentFolderName(strFile)
102     If Len(strFilePath) = 0 Then strFilePath = App.Path
    
104     str仪器型号 = Trim(ReadIni("FASCOW", "仪器型号", App.Path & "\FASCOW.ini"))
106     If str仪器型号 = "" Then str仪器型号 = "默认": Call WriteIni("FASCOW", "仪器型号", str仪器型号, App.Path & "\FASCOW.ini")
    
108     With cnAccess
110         .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
112         .Open
114         If .State = adStateClosed Then Exit Function
116         strSQL = "Select 项目名称,项目单位,参女上,参女下,参男上,参男下 From Item Where 项目名称='全血低切' "
118         Set rsTmp = cnAccess.Execute(strSQL)
120         strLine_Man = "": strLine_Wman = "": str项目单位 = ""
122         Do Until rsTmp.EOF
124             str项目单位 = "" & rsTmp!项目单位
126             str项目单位 = Mid(str项目单位, InStr(str项目单位, "(") + 1, InStr(str项目单位, "/") - InStr(str项目单位, "(") - 1)
128             str低 = str项目单位
130             strLine_Wman = "" & rsTmp!参女上 & "," & str项目单位 & "|" & rsTmp!参女下 & "," & str项目单位
132             strLine_Man = "" & rsTmp!参男上 & "," & str项目单位 & "|" & rsTmp!参男下 & "," & str项目单位
134             rsTmp.MoveNext
            Loop
        
136         strSQL = "Select 项目名称,项目单位,参女上,参女下,参男上,参男下 From Item Where 项目名称='全血高切' "
138         Set rsTmp = cnAccess.Execute(strSQL)
140         Do Until rsTmp.EOF
142             str项目单位 = "" & rsTmp!项目单位
144             str项目单位 = Mid(str项目单位, InStr(str项目单位, "(") + 1, InStr(str项目单位, "/") - InStr(str项目单位, "(") - 1)
146             str高 = str项目单位
148             strLine_Wman = Split(strLine_Wman, "|")(0) & "," & rsTmp!参女上 & "," & str项目单位 & "|" & Split(strLine_Wman, "|")(1) & "," & rsTmp!参女下 & "," & str项目单位
150             strLine_Man = Split(strLine_Man, "|")(0) & "," & rsTmp!参男上 & "," & str项目单位 & "|" & Split(strLine_Man, "|")(1) & "," & rsTmp!参男下 & "," & str项目单位
152             rsTmp.MoveNext
            Loop
        
154         strSQL = "Select 项目名称,项目单位,参女上,参女下,参男上,参男下 From Item Where 项目名称='全血中切' "
156         Set rsTmp = cnAccess.Execute(strSQL)
158         Do Until rsTmp.EOF
160             str项目单位 = "" & rsTmp!项目单位
162             str项目单位 = Mid(str项目单位, InStr(str项目单位, "(") + 1, InStr(str项目单位, "/") - InStr(str项目单位, "(") - 1)
164             str中 = str项目单位
166             strLine_Wman = Split(strLine_Wman, "|")(0) & "," & rsTmp!参女上 & "," & str项目单位 & "|" & Split(strLine_Wman, "|")(1) & "," & rsTmp!参女下 & "," & str项目单位
168             strLine_Man = Split(strLine_Man, "|")(0) & "," & rsTmp!参男上 & "," & str项目单位 & "|" & Split(strLine_Man, "|")(1) & "," & rsTmp!参男下 & "," & str项目单位
170             rsTmp.MoveNext
            Loop
        
172         If str仪器型号 = "2010D" Then
174             strSQL = "Select 项目名称,项目单位,参女上,参女下,参男上,参男下 From Item Where 项目名称='血浆粘度' "
176             Set rsTmp = cnAccess.Execute(strSQL)
178             Do Until rsTmp.EOF
180                 str项目单位 = "" & rsTmp!项目单位
182                 str项目单位 = Mid(str项目单位, InStr(str项目单位, "(") + 1, InStr(str项目单位, ".") - InStr(str项目单位, "(") - 1)
184                 str中 = str项目单位
186                 strLine_Wman = Split(strLine_Wman, "|")(0) & "," & rsTmp!参女上 & "," & str项目单位 & "|" & Split(strLine_Wman, "|")(1) & "," & rsTmp!参女下 & "," & str项目单位
188                 strLine_Man = Split(strLine_Man, "|")(0) & "," & rsTmp!参男上 & "," & str项目单位 & "|" & Split(strLine_Man, "|")(1) & "," & rsTmp!参男下 & "," & str项目单位
190                 rsTmp.MoveNext
                Loop
            End If
        
192         strSQL = "": blnUnion = False
194         If dtStart < date Then '要查询往日
196             strSQL = "SELECT a.编号,A.性别, Format(A.日期,'yyyy-mm-dd') As 日期, a.检验医师, '血液' As 标本, 0 As 质控,a.全血高切,a.全血中切," & _
                "a.全血低切,血浆粘度,红细胞压积,高切还原粘度,中切还原粘度,低切还原粘度,红细胞聚集指数,红细胞变形指数,血沉,血沉方程K值,红细胞刚性指数,红细胞计数,红细胞电泳时间,卡松粘度,卡松屈服应力 " & _
                    " FROM record a WHERE  A.全血低切<>0 and A.全血中切<>0 and A.全血高切<>0 And " & _
                    "  Format(A.日期,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                        " And Format(A.日期,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'" & _
                        IIf(lngSampleNo = -1, "", " And a.编号=" & lngSampleNo)
            End If
198         If dtEnd >= date Then '要查当日
200             If Len(strSQL) > 0 Then blnUnion = True '要Union All
            
202             strSQL = IIf(Len(strSQL) = 0, "", strSQL & " Union All ") & _
                    "SELECT a.编号,A.性别, Format(A.日期,'yyyy-mm-dd') As 日期, a.检验医师, '血液' As 标本, 0 As 质控,a.全血高切,a.全血中切," & _
                    "a.全血低切,血浆粘度,红细胞压积,高切还原粘度,中切还原粘度,低切还原粘度,红细胞聚集指数,红细胞变形指数,血沉,血沉方程K值,红细胞刚性指数,红细胞计数,红细胞电泳时间,卡松粘度,卡松屈服应力 " & _
                    " FROM record a WHERE  A.全血低切<>0 and A.全血中切<>0 and A.全血高切<>0 And " & _
                    "  Format(A.日期,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                        " And Format(A.日期,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'" & _
                    IIf(lngSampleNo = -1, "", " And a.编号=" & lngSampleNo)
            End If
204         strSQL = strSQL & IIf(blnUnion, " Order By 日期 Desc , 编号", " Order By Format(A.日期,'yyyy-mm-dd') Desc , 编号")
206         Set rsTmp = cnAccess.Execute(strSQL)
208         If rsTmp.EOF Then
210             .Close
                Exit Function
            End If
        
212         lngResultCount = -1
214         strBh = ""
216         ReDim aResults(0)
            'rsTmp.MoveFirst
218         Do Until rsTmp.EOF
                    
220                 strBh = rsTmp("编号")
222                 str日期 = Format(rsTmp("日期"), "yyyy-MM-dd HH:mm:ss")
224                 strQC = Val("" & rsTmp("质控"))

226                 strItems = ""
228                 If Trim("" & rsTmp("全血高切")) <> "" Then strItems = strItems & "|01|" & Trim("" & rsTmp("全血高切"))
230                 If Trim("" & rsTmp("全血中切")) <> "" Then strItems = strItems & "|02|" & Trim("" & rsTmp("全血中切"))
232                 If Trim("" & rsTmp("全血低切")) <> "" Then strItems = strItems & "|03|" & Trim("" & rsTmp("全血低切"))
234                 If Trim("" & rsTmp("血浆粘度")) <> "" Then strItems = strItems & "|04|" & Trim("" & rsTmp("血浆粘度"))
236                 If Trim("" & rsTmp("红细胞压积")) <> "" Then strItems = strItems & "|05|" & Trim("" & rsTmp("红细胞压积"))
238                 If Trim("" & rsTmp("高切还原粘度")) <> "" Then strItems = strItems & "|06|" & Trim("" & rsTmp("高切还原粘度"))
240                 If Trim("" & rsTmp("中切还原粘度")) <> "" Then strItems = strItems & "|07|" & Trim("" & rsTmp("中切还原粘度"))
242                 If Trim("" & rsTmp("低切还原粘度")) <> "" Then strItems = strItems & "|08|" & Trim("" & rsTmp("低切还原粘度"))
244                 If Trim("" & rsTmp("红细胞聚集指数")) <> "" Then strItems = strItems & "|09|" & Trim("" & rsTmp("红细胞聚集指数"))
246                 If Trim("" & rsTmp("红细胞变形指数")) <> "" Then strItems = strItems & "|10|" & Trim("" & rsTmp("红细胞变形指数"))
248                 If Trim("" & rsTmp("血沉")) <> "" Then strItems = strItems & "|11|" & Trim("" & rsTmp("血沉"))
250                 If Trim("" & rsTmp("血沉方程K值")) <> "" Then strItems = strItems & "|12|" & Trim("" & rsTmp("血沉方程K值"))
252                 If Trim("" & rsTmp("红细胞刚性指数")) <> "" Then strItems = strItems & "|13|" & Trim("" & rsTmp("红细胞刚性指数"))
254                 If Trim("" & rsTmp("红细胞计数")) <> "" Then strItems = strItems & "|14|" & Trim("" & rsTmp("红细胞计数"))
256                 If Trim("" & rsTmp("红细胞电泳时间")) <> "" Then strItems = strItems & "|15|" & Trim("" & rsTmp("红细胞电泳时间"))
258                 If Trim("" & rsTmp("卡松粘度")) <> "" Then strItems = strItems & "|16|" & Trim("" & rsTmp("卡松粘度"))
260                 If Trim("" & rsTmp("卡松屈服应力")) <> "" Then strItems = strItems & "|17|" & Trim("" & rsTmp("卡松屈服应力"))
                    
262                 If strItems <> "" And strBh <> "" And IsDate(str日期) Then
264                     If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
266                     aResults(UBound(aResults)) = str日期 & "|" & strBh & "| |XY|" & strQC & strItems
                    End If
268                 If str仪器型号 <> "2010D" Then
270                     strimg = vbCrLf & "粘度特征曲线;2;20,210|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|3-3,6-6,9-9,12-12,15-15;"
272                     If "" & rsTmp!性别 = "男" Then
274                         strimg = strimg & strLine_Man
                        Else
276                         strimg = strimg & strLine_Wman
                        End If
278                     strimg = strimg & "~" & str低 & "-" & rsTmp!全血低切 & "," & str中 & "-" & rsTmp!全血中切 & "," & str高 & "-" & rsTmp!全血高切 '描点
280                     strimg = strimg & ";(mPa.s),15,20~(1/S),205,1" '坐标轴
                    Else
                        Dim varTmp() As String
                    
282                     If "" & rsTmp!性别 = "男" Then
                            '下限
284                         varTmp = Split(Split(strLine_Man, "|")(0), ",")
286                         arrLow(0) = varTmp(0): arrMid(0) = varTmp(4): arrHigh(0) = varTmp(2): arrNianDu(0) = varTmp(6)
                            '检验值
288                         arrLow(1) = Val(rsTmp!全血低切): arrMid(1) = Val(rsTmp!全血中切): arrHigh(1) = Val(rsTmp!全血高切): arrNianDu(1) = Val(rsTmp!血浆粘度)
                            '上限
290                         varTmp = Split(Split(strLine_Man, "|")(1), ",")
292                         arrLow(2) = varTmp(0): arrMid(2) = varTmp(4): arrHigh(2) = varTmp(2): arrNianDu(2) = varTmp(6)
                        Else
                             '下限
294                         varTmp = Split(Split(strLine_Wman, "|")(0), ",")
296                         arrLow(0) = varTmp(0): arrMid(0) = varTmp(4): arrHigh(0) = varTmp(2): arrNianDu(0) = varTmp(6)
                            '检验值
298                         arrLow(1) = Val(rsTmp!全血低切): arrMid(1) = Val(rsTmp!全血中切): arrHigh(1) = Val(rsTmp!全血高切): arrNianDu(1) = Val(rsTmp!血浆粘度)
                            '上限
300                         varTmp = Split(Split(strLine_Wman, "|")(1), ",")
302                         arrLow(2) = varTmp(0): arrMid(2) = varTmp(4): arrHigh(2) = varTmp(2): arrNianDu(2) = varTmp(6)
                        End If
                    
304                     strimg = "粘度特征曲线;110;1;" & frmGraph.Draw_2010D(arrHigh, arrMid, arrLow, arrNianDu, App.Path, rsTmp!编号)
                    End If
306                 If strimg <> "" Then
308                         aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & strimg
                    End If
310             rsTmp.MoveNext
            Loop
        End With
312     ResultFromFile = aResults
hErr:
314     WriteLog "Fascow", "第" & CStr(Erl()) & "行", Err.Description
End Function
