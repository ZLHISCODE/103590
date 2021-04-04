VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_HALO_F180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:F180粪便分析仪
'厂商:
'作者:王冬建
'用户:六安市中医院

Public Function GetStartCmd() As String
    '第一次启动时调用
End Function

Public Function GetEndCmd() As String
    '接口关闭的时候调用
End Function

Public Function GetAnswerCmd() As String
    '定时发送的指令
End Function

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo ErrHand '打开错误陷阱
    '------------------------------------------------
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
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select format(riqi,'yyyy-mm-dd hh:mm:ss') as 日期,BianHao as 标本号,Fenbianyanse as 颜色,Xingzhuang as 性状,Hongxibao as 红细胞,Baixibao as 白细胞,Shiwuxibao as 嗜物细胞,Buxiaohuashiwu as 不消化食物,"
        strSQL = strSQL & "Gouchongluan as 钩虫卵,Yuanchong as 原虫,Baonang as 包囊,Huichongluan as 蛔虫卵,Meijunbaozi as 酵母菌,Qita as 其它,Bianchongluan as 鞭虫卵,Jushixibao as 吞噬细胞,Yinxiehuaxue as 隐血化验,"
        strSQL = strSQL & "Qitachongluan as 其它虫卵,Jiadishi as 贾第氏,Yinxiemianyi as 脓细胞,Changdichong as 肠滴虫,Xialeidengshi as 夏雷登氏,Xianbingdu as 腺病毒,Lunzhuanbingdu as 轮状病毒,Kesaqi as 科萨奇,"
        strSQL = strSQL & "Youmenluoxuan as 幽门螺旋,Zhifangqiu as 脂肪球,Zhuantiedanbai as 转铁蛋白,Dangqianzhaopian as 照片 From report where Format(riqi,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "'"
        strSQL = strSQL & " And Format(riqi,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "' and wancheng = 1" & _
                 IIf(lngSampleNO = -1, "", " And val(bianhao)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If

        rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            aResults(UBound(aResults)) = rsTmp("日期") & "|" & rsTmp("标本号") & "| |粪便|0|FBYS|" & Trim(rsTmp("颜色")) & "|FBXZ|" & Trim(rsTmp("性状")) & "|HXB|" & Trim(rsTmp("红细胞")) & _
                     "|BXB|" & Trim(rsTmp("白细胞")) & "|SWXB|" & Trim(rsTmp("嗜物细胞")) & "|BXHSW|" & Trim(rsTmp("不消化食物")) & "|GCL|" & Trim(rsTmp("钩虫卵")) & "|YC|" & Trim(rsTmp("原虫")) & _
                     "|BN|" & Trim(rsTmp("包囊")) & "|HCL|" & Trim(rsTmp("蛔虫卵")) & "|JMJ|" & Trim(rsTmp("酵母菌")) & "|QT|" & IIf(Trim(rsTmp("其它虫卵")) = "", "未查见", Trim(rsTmp("其它"))) & "|BCL|" & Trim(rsTmp("鞭虫卵")) & _
                     "|TSXB|" & Trim(rsTmp("吞噬细胞")) & "|YXSY|" & Trim(rsTmp("隐血化验")) & "|QTCL|" & IIf(Trim(rsTmp("其它虫卵")) = "", "未查见", Trim(rsTmp("其它虫卵"))) & "|JDS|" & Trim(rsTmp("贾第氏")) & "|NXB|" & Trim(rsTmp("脓细胞")) & _
                     "|CDC|" & Trim(rsTmp("肠滴虫")) & "|XLDS|" & Trim(rsTmp("夏雷登氏")) & "|XBD|" & Trim(rsTmp("腺病毒")) & "|LZBD|" & Trim(rsTmp("轮状病毒")) & "|KSQ|" & Trim(rsTmp("科萨奇")) & _
                     "|YMLX|" & Trim(rsTmp("幽门螺旋")) & "|ZFQ|" & Trim(rsTmp("脂肪球")) & "|ZTDB|" & rsTmp("转铁蛋白")
NextDo:
            rsTmp.MoveNext
        Loop
    End With
    '解析完成，关闭文件
    If rsTmp.EOF Then
        cnAccess.Close
    End If
    ResultFromFile = aResults

    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_File_HALO_F180"
    Close #lngErrFileNo
End Function
