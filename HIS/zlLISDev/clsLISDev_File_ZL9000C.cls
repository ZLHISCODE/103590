VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Zl9000c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'血流变
'仪器型号：ZL9000C
'用户：灵丘县人民医院
'作战中：郝立功
'日期：2014-07-08
'
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
    On Error GoTo ErrHand '打开错误陷阱
    '------------------------------------------------
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

    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, yp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String, strlog As String
    Dim testid As Long
    Dim blnUnion As Boolean
    Dim lngCount As Long, str点 As String, varTmp As Variant
    Dim str男_低1 As String, str男_低2 As String, str女_低1 As String, str女_低2 As String
    Dim str男_高1 As String, str男_高2 As String, str女_高1 As String, str女_高2 As String
    Dim strIMG As String
    Dim lngRecord As String
    '写日志
    strlog = "已经调用部件"
    Call writelislog(strlog)
    ReDim Preserve aResults(0)
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        '写日志
        strlog = "已经打开文件"
        Call writelislog(strlog)

        strSQL = "Select ShearRate as X,MaleNormal As 男,FemaleNormal As 女 From ShearRate order by ShearRate"
        Set rsTmp = cnAccess.Execute(strSQL)
        lngRecord = 0
        Do Until rsTmp.EOF
            lngRecord = lngRecord + 1
            rsTmp.MoveNext
        Loop
        lngCount = 0
        If lngRecord > 0 Then rsTmp.MoveFirst
        Do Until rsTmp.EOF
            lngCount = lngCount + 1
            If lngRecord > 3 Then
                If lngCount = 1 Then
                    str男_低1 = "" & Split("" & rsTmp!男, "-")(0) & "," & rsTmp!x
                    str男_低2 = "" & Split("" & rsTmp!男, "-")(1) & "," & rsTmp!x
                    str女_低1 = "" & Split("" & rsTmp!女, "-")(0) & "," & rsTmp!x
                    str女_低2 = "" & Split("" & rsTmp!女, "-")(1) & "," & rsTmp!x
                End If
            Else
                If lngCount = 1 Then
                    str男_低1 = "" & Split("" & rsTmp!男, "-")(0) & "," & rsTmp!x
                    str男_低2 = "" & Split("" & rsTmp!男, "-")(1) & "," & rsTmp!x
                    str女_低1 = "" & Split("" & rsTmp!女, "-")(0) & "," & rsTmp!x
                    str女_低2 = "" & Split("" & rsTmp!女, "-")(1) & "," & rsTmp!x
                End If
            End If
            If lngCount = lngRecord Then
                str男_高1 = "" & Split("" & rsTmp!男, "-")(0) & "," & rsTmp!x
                str男_高2 = "" & Split("" & rsTmp!男, "-")(1) & "," & rsTmp!x
                str女_高1 = "" & Split("" & rsTmp!女, "-")(0) & "," & rsTmp!x
                str女_高2 = "" & Split("" & rsTmp!女, "-")(1) & "," & rsTmp!x
            End If
            rsTmp.MoveNext
        Loop

        strSQL = "": blnUnion = False
        'If dtStart < date Then '要查询往日

        '用日期和序号来进行定位
        strSQL = " select testdataid,序号,性别,日期 as Rq,检验医生 as bg,'血液' as 标本,0 as 质控," & _
                 " '01|' +  cstr(iif(isnull(全血粘度),0,全血粘度)) + '|02|' + cstr(iif(isnull(血浆粘度),0,血浆粘度)) + '|03|' + cstr(iif(isnull(压积),0,压积)) + '|04|' + cstr(iif(isnull(血沉),0,血沉)) + '|05|' +  cstr(format(全血低切相对指数,'0.00')) " & _
                 " + '|06|' + cstr(format(全血高切相对指数,'0.00')) + '|07|' + cstr(format(血沉方程K值,'0.00')) + '|08|' + cstr(format(红细胞聚集指数,'0.00'))" & _
                 " + '|09|' + cstr(iif(isnull(红细胞变形指数),0,红细胞变形指数)) + '|10|' + cstr(format(全血低切还原粘度,'0.00')) + '|11|' + cstr(format(全血高切还原粘度,'0.00'))" & _
                 " + '|12|' + cstr(format(红细胞变形指数TK,'0.00')) + '|13|' + cstr(format(红细胞刚性指数,'0.00'))as 结果1," & _
                 "'|14|' + cstr(format(卡松粘度,'0.00')) + '|15|' + cstr(iif(isnull(纤维蛋白原),0,纤维蛋白原))" & _
                 " + '|16|' + cstr(iif(isnull(血胆固醇),0,血胆固醇)) + '|17|' + cstr(iif(isnull(甘油三脂),0,甘油三脂)) + '|18|' + cstr(iif(isnull(高密脂蛋白),0,高密脂蛋白)) + '|19|' + cstr(iif(isnull(血糖),0,血糖)) + '|20|' + cstr(iif(isnull(血小板粘附率),0,血小板粘附率))" & _
                 " + '|21|' + cstr(iif(isnull(体外血栓干重),0,体外血栓干重)) + '|22|' + cstr(iif(isnull(红细胞电泳),0,红细胞电泳)) + '|23|' + cstr(iif(isnull(血小板聚集率),0,血小板聚集率)) + '|24|' + cstr(iif(isnull(体外血栓长度),0,体外血栓长度))" & _
                 " + '|25|' + cstr(format(中切流阻,'0.00')) + '|26|' + cstr(format(高切流阻,'0.00')) + '|27|' + cstr(format(红细胞聚集系数,'0.00')) + '|28|' + cstr(format(红细胞内粘度,'0.00')) " & _
                 " + '|29|' + cstr(iif(isnull(红细胞计数),0,红细胞计数)) + '|30|' + cstr(format(低切流阻,'0.00')) + '|31|' + cstr(iif(isnull(屈服应力),0,屈服应力)) + '|32|' + cstr(iif(isnull(全血中切还原粘度),0,全血中切还原粘度)) as 结果2" & _
                 " from testdata " & _
                 IIf(lngSampleNO = -1, "where ", " where 序号=" & lngSampleNO & " and ") & _
                 " 日期>=cdate('" & Format(dtStart, "yyyy-mm-dd") & "')" & _
                 " and 日期<=cdate('" & Format(dtEnd, "yyyy-mm-dd 23:59:59") & "')"
        '" + '|25|' + cstr(iif(isnull(血沉最大沉降率),0,血沉最大沉降率)) + '|26|' + cstr(iif(isnull(血沉最大沉降率时间),0,血沉最大沉降率时间)) as 结果2" &

        strSQL = strSQL & IIf(blnUnion, "", " Order By 日期 Desc")

        Set rsTmp = cnAccess.Execute(strSQL)
        'cnAccess.

        If rsTmp.EOF Then
            .Close
            '写日志
            strlog = "没有找到数据"
            Call writelislog(strlog)
            ResultFromFile = aResults
            Exit Function
        End If
        '写日志
        strlog = "yp:已经找到数据"
        Call writelislog(strlog)

        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("序号") Then '新标本

                lngResultCount = lngResultCount + 1
                If lngResultCount > 0 Then
                    ReDim Preserve aResults(lngResultCount)
                End If
                aResults(lngResultCount) = Format(rsTmp("Rq"), "yyyy-mm-dd") & " " & Format(Now, "HH:mm:ss") '检验时间
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         Val(rsTmp("序号")) '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         IIf(Trim("" & rsTmp("Bg")) = "", " ", rsTmp("Bg")) '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         IIf(IsNull(rsTmp("标本")), " ", rsTmp("标本")) '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         Val(rsTmp("质控")) '质控

                strBh = rsTmp("序号")

            End If
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                     Replace(rsTmp("结果1"), Chr(0), "") + Replace(rsTmp("结果2"), Chr(0), "") '结果
            '得到测试id，以从visc中得到shearRate的Visc值
            testid = rsTmp("testdataid")
            strSQL = "select testdataid,shearRate,Visc from visc where testdataid=" & testid & ""
            ' If .State = 1 Then .Close
            Set yp = cnAccess.Execute(strSQL)
            str点 = ""
            Do While Not yp.EOF
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                         yp("shearRate") & "|" & yp("visc")

                str点 = str点 & "," & yp("shearRate") & "-" & yp("visc")
                yp.MoveNext
            Loop
            If str点 <> "" Then
                strIMG = vbCrLf & "粘度特征曲线;2;20,200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;"
                If Trim("" & rsTmp("性别")) = True Then
                    '男
                    strIMG = strIMG & str男_低1 & "," & str男_高1 & "|" & str男_低2 & "," & str男_高2
                Else
                    strIMG = strIMG & str女_低1 & "," & str女_高1 & "|" & str女_低2 & "," & str女_高2
                End If
                strIMG = strIMG & "~" & Mid(str点, 2)
                strIMG = strIMG & ";VIS(mPa.s),30,20~SHR(1/S),195,1" '坐标轴
            End If
            If strIMG <> "" Then aResults(lngResultCount) = aResults(lngResultCount) & strIMG
            '                If str点 <> "" Then
            '                     varTmp = Split(str点, ",")
            '                     If UBound(varTmp) < 4 Then
            '                        str点 = varTmp(0) & "," & varTmp(1) & "," & varTmp(2) & "," & varTmp(3) & "," & varTmp(3)
            '
            '                     End If
            '                    '根据公式画图
            '                    If Trim("" & rsTmp!性别) = True Then    '男
            '                        strImg = "粘度特征曲线;110;1;200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20,;"
            '
            '                        strImg = strImg & str男_低1 & "," & str男_高2 & "|" & str男_低2 & "," & str男_高2
            '                    Else
            '                        strImg = "粘度特征曲线;110;1;200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20,;"
            '
            '                        strImg = strImg & str女_低1 & "," & str女_高2 & "|" & str女_低2 & "," & str女_高2
            '                    End If
            '
            '                    strImg = strImg & "~" & Mid(str点, 2)
            '                    strImg = strImg & ";VIS(mPa.s),30,20~SHR(1/S),195,1" '坐标轴
            '
            '                End If
            '                If strImg <> "" Then aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & strImg
            rsTmp.MoveNext
        Loop
    End With
    '写日志
    strlog = "最后一条数据：" & aResults(lngResultCount)
    Call writelislog(strlog)

    ResultFromFile = aResults
    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",writelislog,clsLISDev_File_Zl9000c"
    Close #lngErrFileNo
End Function

Public Function writelislog(strlog As String)
    On Error GoTo ErrHand '打开错误陷阱
    '------------------------------------------------
    Const strFile As String = "c:\lislog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFileName As String
    strFileName = strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strlog)
    objStream.Close
    Set objStream = Nothing

    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",writelislog,clsLISDev_File_Zl9000c"
    Close #lngErrFileNo
End Function
