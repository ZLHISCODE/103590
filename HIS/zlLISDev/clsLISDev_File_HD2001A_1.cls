VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_HD2001A_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:生物芯片检测仪
'型号:HD2001A
'厂商:上海数康生物科技有限公司
'作者:李建飞
'用户:
'日期:2009-08-21

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
    'strFile：包含路径的文件名
    'lngSampleNo：标本号。为-1表示选取所有时间范围内的标本
    'dtStart：开始时间，如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
    'dtEnd：结束时间。只有选取多个标本(lngSample=-1)时该参数有效。如果不指定该参数，则时间范围为>=dtStart
    '返回的检验结果数组（各解析程序必须按照以下标准组织结果）
    '元素之间以 | 分隔
    '第 0 个元素：检验时间
    '第 1 个元素：样本序号
    '第 2 个元素：检验人
    '第 3 个元素：标本
    '第 4 个元素：是否质控品
    '从第 5 个元素开始为检验结果，每两个元素表示一个检验项目
    '如：第 5i 个元素为检验项目，第 5i+1 个元素表示检验结果
    
    Dim aResults() As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim cnErs As New ADODB.Connection, strFileErs As String, rsErs As ADODB.Recordset, strImgErs As String
    Dim str_检验项目 As String
        
    strFilePath = objFile.GetParentFolderName(strFile)  '取路径
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFileErs = strFile
    
    If Trim(Dir(strFileErs)) <> "" Then
        With cnErs
            .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};"
            .Open
            If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
    
    ReDim aResults(0)
    strImgErs = ""
    
    If strFileErs <> "" Then
        strSql = "select testdate as 检验时间,ID as 样本序号,testdoctor as 检验人,biaobenClass as 标本, " & _
                 "CA19_9,NSE,CEA,CA242,Ferritin,Beta_HCG,AFP,Free_PSA,PSA,CA125,HGH,CA15_3 from TestData " & _
                 "as a where format(cdate(testdate),'yyyy-mm-dd')>= '" & Format(dtStart, "yyyy-mm-dd") & "' " & _
                 "and format(cdate(testdate),'yyyy-mm-dd')<= '" & Format(dtEnd, "YYYY-MM-DD") & "'" & _
                 IIf(lngSampleNo = -1, "", " and VAL(ID)=" & lngSampleNo)
        Set rsErs = cnErs.Execute(strSql)
        
        While Not rsErs.EOF
            '根据文件结构读取检验信息
            If Val("" & rsErs!样本序号) <> 0 Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                '第 0 个元素：检验时间
                '第 1 个元素：样本序号
                '第 2 个元素：检验人
                '第 3 个元素：标本
                '第 4 个元素：是否质控品
                '第 5 个元素开始为检验结果，每两个元素表示一个检验项目
                '如：第 5i 个元素为检验项目，第 5i+1 个元素表示检验结果
                aResults(UBound(aResults)) = Format(CDate(rsErs!检验时间), "yyyy-mm-dd")                        '检验时间
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsErs!样本序号)        '样本序号
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "                                  '检验人
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsErs!标本                      '标本
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                                  '是否质控品
                
                '检验项目
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|CA19_9|" & Trim(rsErs!CA19_9)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|NSE|" & Trim(rsErs!NSE)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|CEA|" & Trim(rsErs!CEA)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|CA242|" & Trim(rsErs!CA242)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|Ferritin|" & Trim(rsErs!Ferritin)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|Beta_HCG|" & Trim(rsErs!Beta_HCG)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|AFP|" & Trim(rsErs!AFP)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|Free_PSA|" & Trim(rsErs!Free_PSA)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|PSA|" & Trim(rsErs!PSA)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|CA125|" & Trim(rsErs!CA125)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|HGH|" & Trim(rsErs!HGH)
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|CA15_3|" & Trim(rsErs!CA15_3)
            End If

            rsErs.MoveNext
        Wend
    End If
    ResultFromFile = aResults
End Function
