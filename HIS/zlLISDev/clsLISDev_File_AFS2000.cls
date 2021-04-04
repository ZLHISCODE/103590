VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AFS2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:飞测 CRP检查
'厂商:飞测 AFS2000
'作者:罗云龙
'用户:奇台县中医医院
'日期:2015-12-11

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
'标准函数：从文件获取检验结果
'   strFile：包含路径的文件名
'   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
'   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
'   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
'   返回的检验结果数组(各解析程序必须按以下标准组织结果)
'   元素之间以||分隔
'   第0个元素：检验时间
'   第1个元素：样本序号
'   第2个元素：检验人
'   第3个元素：标本
'   第4个元素：是否质控品
'   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
'       如：第5i个元素为检验项目，第5i + 1个元素为检验结果

'微生物接口此处填 抗生素|结果, 抗生素的结果格式为
'      药敏结果^药敏方法^结果值

'微生物结果示例:
'    药敏结果取值为:R,S,I ,ESBL,BLAC,R*。R-耐药，S-敏感，I-中介，
'　　　　       ESBL β—内酰胺酶,BLAC 未知，R＊结果不明确
'              药敏方法取值为:1,2,3。1-MIC，2-DISK ，3-K-B


    
    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, yp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String, strlog As String, i As Integer
    Dim blnUnion As Boolean
    Dim strDns As String, strUser As String, strPassWord As String
    
    On Error GoTo errHand
    
    strUser = "admin": strPassWord = "afs2000"
    
    Err = 0
    
    '写日志
    strlog = "已经调用部件"
    Call writelislog(strlog)
    ReDim Preserve aResults(0)
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";uid=" & strUser & ";pwd=" & strPassWord & ";Driver={Microsoft Access Driver (*.mdb)}"
'        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        
        If .State = adStateClosed Then Exit Function

        '写日志
        strlog = "已经打开文件"
        Call writelislog(strlog)
        lngSampleNO = -1
        If .State = adStateClosed Then Exit Function
        strSQL = "": blnUnion = False
        strSQL = "SELECT TDate+' '+ TTime As 检验日期,SampleID As 标本号,OutName As 通道号,Result As 检验结果 from SamplesHis A where " & vbNewLine & _
                IIf(lngSampleNO = -1, "  Format(TDate,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                " And Format(TDate,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'", _
                " SampleID='" & Format(lngSampleNO, "00") & "'") & _
                " Union all SELECT TDate+' '+ TTime As 检验日期,SampleID As 标本号,OutName As 通道号,Result As 检验结果 from Samples A where " & vbNewLine & _
                IIf(lngSampleNO = -1, "  Format(TDate,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                " And Format(TDate,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'", _
                " SampleID='" & Format(lngSampleNO, "00") & "'") & _
                " order by 检验日期,标本号 desc"

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("标本号") Then  '新标本
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("检验日期"), "yyyy-mm-dd")  '检验时间
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val(rsTmp("标本号")) '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    " " '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    "血液" '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    "0" '质控
                    
                strBh = rsTmp("标本号")
            End If
            '定性结果[^OD^CutOff^sco]
            aResults(lngResultCount) = aResults(lngResultCount) & _
                 IIf(rsTmp("检验结果") = "", "", "|" & UCase(rsTmp("通道号")) & "|" & rsTmp("检验结果"))
            rsTmp.MoveNext
        Loop
    End With
   '写日志
    strlog = "最后一条数据：" & aResults(lngResultCount)
    Call writelislog(strlog)
    
    ResultFromFile = aResults
    
    Exit Function
errHand:
    WriteLog "ResultFromFile", "", Err.Number & " " & Err.Description
End Function

Public Function writelislog(strlog As String)
    Const strFile As String = "c:\lislog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFilename As String
    strFilename = strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFilename) Then Call objFile.CreateTextFile(strFilename)
    Set objStream = objFile.OpenTextFile(strFilename, ForAppending)
    objStream.WriteLine (strlog)
    objStream.Close
    Set objStream = Nothing

End Function



