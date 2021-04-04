VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ATB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:VITEK ATB 细菌分析仪
'厂商:梅里埃 VITEK  北京东方旗云开发的软件
'作者:陈东
'用户:晋煤集团总医院

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
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
        Dim strBh As String, strAnti As String, str结果 As String, str细菌 As String
        Dim lngResultCount As Long
        Dim strType As String  '2009-07-19 增加 陈辉哪边有个用户的数据是从Result表中提取数据，而原来是从ReqToExe表中提数据,
                               'ATBFile.ini
        Dim strBZ As String '是否加备注，老版的liscomm不支持检验备注，要报错。
        
        On Error GoTo errH
100     strType = ReadIni("设置", "TYPE", App.Path & "\ATBFile.ini")
102     strBZ = ReadIni("设置", "备注", App.Path & "\ATBFile.ini")
104     If strType = "" Then strType = "REQ"
        
106     strFilePath = objFile.GetParentFolderName(strFile)
108     If Len(strFilePath) = 0 Then strFilePath = App.Path

110     cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd=scsl"
112     cnAccess.Open
114     If cnAccess.State = adStateClosed Then Exit Function
116     strSQL = "Select A.ReqNO as 编号,A.ExeDate as 日期,A.GermID as 细菌,B.samNo as 标本号 "
118     If strType <> "Result" Then
120         strSQL = strSQL & " From ReqToExe A,RequistionBill B "
        Else
122         strSQL = strSQL & " From ResultExe A,ResultBill B "
        End If
124     strSQL = strSQL & " Where A.ReqNO=B.ReqNO And   A.ExeDate >=#" & Format(dtStart, "yyyy-MM-dd") & "#"
126     strSQL = strSQL & " And  A.ExeDate <=#" & Format(dtEnd, "yyyy-MM-dd") & "#" & _
        IIf(lngSampleNO = -1, "", " And val(A.ReqNo)=" & lngSampleNO)
128     WriteLog "File_ATB", strSQL, ""
    
130     strBh = "***"
132     str细菌 = ""
134     ReDim aResults(0)
    
136     Set rsTmp = cnAccess.Execute(strSQL)
138     If rsTmp.EOF Then
140         cnAccess.Close
142         WriteLog "File_ATB", "无记录！", ""
            Exit Function
        End If
        
144     Do Until rsTmp.EOF
146         If strBh <> rsTmp("编号") Or str细菌 <> Trim("" & rsTmp("细菌")) Then '新标本
148             If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            
150             aResults(UBound(aResults)) = Format(rsTmp("日期"), "yyyy-mm-dd hh:mm:ss") '检验时间
152             aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                    Val(rsTmp("标本号")) '样本序号
154             aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
156             aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                    IIf(IsNull(rsTmp("细菌")), " ", rsTmp("细菌")) '标本
158             aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
160             str细菌 = Trim("" & rsTmp("细菌"))
162             strBh = rsTmp("编号")
164             strAnti = ""
166             If strType <> "Result" Then
168                 strSQL = "Select AntiID,MicExplain,MicResult,ResShow From MicResult Where val(ReqNo)=" & strBh
                Else
170                 strSQL = "Select AntiID,MicExplain,MicResult,ResShow From ResultMic Where val(ReqNo)=" & strBh
                End If
172             Set rsAnti = cnAccess.Execute(strSQL)
174             Do Until rsAnti.EOF
176                 strAnti = strAnti & "|" & rsAnti!AntiID & "|" & rsAnti!MicExplain & "^1^" & rsAnti!ResShow
178                 rsAnti.MoveNext
                Loop
180             aResults(UBound(aResults)) = aResults(UBound(aResults)) & strAnti
182             WriteLog "File_ATB", aResults(UBound(aResults)), ""
            End If
184         rsTmp.MoveNext
        Loop
        
186     ResultFromFile = aResults
        Exit Function
errH:
188     WriteLog "File_ATB", CStr(Erl()) & ":" & Err.Description, ""
End Function
