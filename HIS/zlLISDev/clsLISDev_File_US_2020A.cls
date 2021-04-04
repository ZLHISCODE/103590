VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_US_2020A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:US-2020A 尿沉渣检测分析仪
'厂商:重庆天海医疗设备有限公司
'作者:王振涛
'用户:河津铝厂职工医院
'时间:2012-10-18
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
                               ByVal dtStart As Date, Optional dtEnd As Date = CDate("3000-12-31")) As String()
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
    ' Dim DBConn As New ADODB.Connection
    '打开SQL数据库
    ' DBConn.Open "Provider   =   sqloledb;Data   Source= '服务器名 ';
    'Initial   Catalog= '数据库名称 ';User   Id= '用户名 ';Password= '密码 ' "
    '如不想将连接字符串写在程序里的话，还可以通过读INI文件，例如配置内容如下:
    '检验技师工作站指定文件时只需指定这个ini文件即可
    '[LX2000]
    'Database Name = ExamDir
    'SERVER Name = (local)
    'USER Name = sa
    'PASSWORD=

    Dim aResults() As String
    Dim ConnSQL As New ADODB.Connection                                 '数据库连接
    Dim rsSql As ADODB.Recordset                                        '返回的记录集
    Dim strSQL As String                                                '执行的SQL语句
    Dim strDate As String                                               '找出的时间
    Dim strNo As String                                                 '标本号
    Dim strItems As String                                              '通道码及结果
    Dim lngResultCount As Long, strid As String
    Dim strLog As String                                                '记录的日志
    
    Dim strIniFile As String, strImageFileName As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream
    Dim strBH As String, intloop As Integer                                                 '判断是否是新的标本号
    Dim strPicPath As String, strPicFix As String                       '图片文件
    Dim intDebug As Integer

    Dim DataBaseName As String
    Dim serverName As String
    Dim userName As String
    Dim passWord As String
    Dim objFile As New Scripting.FileSystemObject                       '判断是否有图形文件
    Dim strImgBmp As String, strTmp As String                           '保存图形数据
    Dim strYear As String, strYearMonth As String, strYearMonthDay As String
    Dim strImg As String
    Dim varItem As Variant, strItem As String
    On Err GoTo errH
    strIniFile = App.Path & "\US_2020A.ini"

    '------------------------------------读取配置文件--------------------
    If Not objFileSystem.FileExists(strFile) Then
        Set objStream = objFileSystem.CreateTextFile(strIniFile)
        objStream.WriteLine "[Debug]"
        objStream.WriteLine "Text=0"    '1＝调试模式会产生日志 0=正常模式不会产生日志

        objStream.WriteLine "[Conn]"
        objStream.WriteLine "DatabaseName= ExamDir"
        objStream.WriteLine "serverName=(local)"
        objStream.WriteLine "userName=sa"
        objStream.WriteLine "passWord="

        objStream.WriteLine "[Options]"
        objStream.WriteLine "PicPath=D:\Program Files\US-2020A\US-2020A\pic"
        objStream.WriteLine "PicFix=U100N"
        objStream.Close
        Set objStream = Nothing
        strPicPath = "D:\Program Files\US-2020A\US-2020A\pic"
        intDebug = 0
        DataBaseName = "ExamDir"  '"th_us2020(1.5)"
        serverName = "(local)" ' "YF_WZT\SQLEXPRESS"
        userName = "sa"
        passWord = ""

    Else
        strPicPath = ReadIni("Options", "PicPath", strIniFile)
        strPicFix = ReadIni("Options", "PicFix", strIniFile)
        If (Len(strPicPath) > 0) Then
            If Mid(strPicPath, Len(strPicPath)) <> "\" Then strPicPath = strPicPath + "\"
        End If
        intDebug = ReadIni("Debug", "Text", strIniFile, 0)
        DataBaseName = ReadIni("Conn", "DatabaseName", strIniFile, "ExamDir")    ' "th_us2020(1.5)")
        serverName = ReadIni("Conn", "serverName", strIniFile, "(local)")  '"YF_WZT\SQLEXPRESS")
        userName = ReadIni("Conn", "userName", strIniFile, "sa")
        passWord = ReadIni("Conn", "passWord", strIniFile, "")
    End If
    '----------------------------------END-------------------------------
    With ConnSQL
        .ConnectionString = "Provider=SQLOLEDB;Data Source=" & serverName & ";user id=" & userName & ";password=" & passWord & ";initial catalog=" & DataBaseName
        .Open
        If .State = adStateClosed Then
            If intDebug = 1 Then
                strLog = "连接SQL SERVER数据库失败"
                Call writelislog(strLog)
            End If
            Exit Function
        End If

        If intDebug = 1 Then
            strLog = "连接SQL SERVER数据库成功"
            Call writelislog(strLog)
        End If
    
        strSQL = "select a.id, b.sxh as sxh, a.ifprint,a.caption as caption ,a.val as val ,b.name as name ,b.sjsj as sjsj " & _
                    ",b.sjrq as sjrq from dbo.jz_item as a,dbo.jzxx as b where " & vbNewLine & _
                    " b.rq>= '" & Format(dtStart, "YYYY-MM-DD") & "' " & vbNewLine & _
                    " and b.sjrq<= '" & Format(dtEnd, "YYYY-MM-DD 23:59:59") & "' and  " & _
                    IIf(lngSampleNO = -1, " ", "b.sxh =" & lngSampleNO & "and") & _
                    " a.id= b.id and a.val is not null order by a.id  "
                    
        Set rsSql = ConnSQL.Execute(strSQL)

        If rsSql.EOF Then
            .Close
            ReDim Preserve aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If

        rsSql.MoveFirst
        lngResultCount = -1
        strBH = "***"
        Do While Not rsSql.EOF
            If strBH <> rsSql("ID") Then    '新样本
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                aResults(lngResultCount) = Format(rsSql("sjrq") & " " & rsSql("sjsj"), "YYYY-MM-DD hh:mm:ss")
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           Val(rsSql("sxh"))                                    '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                         " "  '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           "尿液"  '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"    '& _

                If strItems <> "" And strTmp <> "" Then
                    aResults(lngResultCount - 1) = aResults(lngResultCount - 1) & vbCrLf & Mid(strTmp, 2)
                    strTmp = ""
                End If
                If strPicPath <> "" Then
                  '读取本地选定图片
                    strImgBmp = ""
                    strTmp = ""
                    strYear = Format(rsSql("sjrq"), "yyyy")
                    strYearMonth = Val(Format(rsSql("sjrq"), "MM"))
                    strYearMonthDay = Val(Format(rsSql("sjrq"), "dd"))
                    strImgBmp = strPicPath & strYear & "\" & strYearMonth & "\" & strYearMonthDay & "\pic\"
                    strImageFileName = Dir(strImgBmp & "\*.jpg", vbHidden + vbDirectory + vbNormal + vbReadOnly + vbSystem + vbVolume)
                    Do While strImageFileName <> ""
                        strid = Mid(strImageFileName, 1, InStr(strImageFileName, "_") - 1)
                        If strid = rsSql("id") Then
                            strTmp = strTmp & "^" & strImageFileName & ";113;0;" & strImgBmp & strImageFileName
                            If intDebug = 1 Then
                                strLog = "找到选定图片: " & strImgBmp & strImageFileName
                                Call writelislog(strLog)
                            End If
                        End If
'                          strImg = rsSql("ID") & "_00" & intloop & ".jpg"
'
'                          If Dir(strImgBmp & strImg, vbHidden) Then
'                              strTmp = strTmp & "^" & CStr(intloop) & ";113;0;" & strImgBmp & strImg
'                              If intDebug = 1 Then
'                                  strLog = "找到选定图片: " & strImgBmp & strImg
'                                  Call writelislog(strLog)
'                              End If
'                          End If
                        strImageFileName = Dir
                    Loop
                    
                    
                    If intDebug = 1 Then
                        strLog = "当前数据: " & aResults(lngResultCount)
                        Call writelislog(strLog)
                    End If
                End If
                
                strBH = rsSql("id")
            
            End If

            If Trim(rsSql("caption")) <> "顺序号" Then
                strItems = rsSql("caption")
                strItems = ReplaceItem(strItems, rsSql)                                         '将中文通道码替换成英文
                
                If strItems <> "" Or IsNull(strItems) = False Then
                    aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                        Trim(strItems) & "|" & IIf(IsNull(rsSql("val")), " ", Split(rsSql("val"), " ")(0))
                End If
            End If
            If intDebug = 1 Then
                strLog = "当前数据: " & aResults(lngResultCount)
                Call writelislog(strLog)
            End If
            rsSql.MoveNext
        Loop
        If strTmp <> "" Then aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & Mid(strTmp, 2)
    End With
    If rsSql.EOF Then
        ConnSQL.Close
    End If
    ResultFromFile = aResults
    Exit Function
errH:
    If intDebug = 1 Then
        Call writelislog("错误号(" & Err.Number & "):(" & Err.Description & ")")
    End If
End Function


Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_US_2020A_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & strLog)
    objStream.Close
    Set objStream = Nothing
End Function
 
 '将中文通道码替换成英文
Private Function ReplaceItem(ByVal strItems As String, ByVal rsSql As Recordset) As String
        strItems = Replace(strItems, "葡萄糖", "GLU")
        strItems = Replace(strItems, "蛋白质", "PRO")
        strItems = Replace(strItems, "胆红素", "BIL")
        strItems = Replace(strItems, "尿胆原", "URO")
        strItems = Replace(strItems, "PH值", "PH")
        strItems = Replace(strItems, "比重", "SG")
        strItems = Replace(strItems, "隐血", "BLD")

        strItems = Replace(strItems, "酮体", "KET")
        strItems = Replace(strItems, "亚硝酸盐", "NIT")
        If InStr(rsSql("val"), "+") > 0 Or InStr(rsSql("val"), "-") > 0 Then
            strItems = Replace(strItems, "白细胞", "LUE")
        Else
            strItems = Replace(strItems, "白细胞", "BXB")
        End If
        strItems = Replace(strItems, "红细胞", "HXB")
        strItems = Replace(strItems, "上皮细胞", "SPXB")

        strItems = Replace(strItems, "小圆上皮", "XYSP")
        strItems = Replace(strItems, "红细胞管型", "HXBGX")
        strItems = Replace(strItems, "白细胞管型", "BXBGX")
        strItems = Replace(strItems, "上皮细胞管型", "SPXBGX")
        strItems = Replace(strItems, "蜡样管型", "CYGX")
        strItems = Replace(strItems, "脂肪管型", "ZFGX")
        strItems = Replace(strItems, "颗粒管型", "KLGX")

        strItems = Replace(strItems, "混合管型", "HHGX")
        strItems = Replace(strItems, "宽形管型", "KXGX")
        strItems = Replace(strItems, "尿酸结晶", "LSJJ")
        strItems = Replace(strItems, "草酸钙结晶", "CSGJJ")
        strItems = Replace(strItems, "亮氨酸结晶", "LASJJ")
        strItems = Replace(strItems, "胆红素结晶", "DHSJJ")
        strItems = Replace(strItems, "胆固醇结晶", "DGCJJ")

        strItems = Replace(strItems, "磺胺结晶", "HAJJ")
        strItems = Replace(strItems, "其它细胞1", "QTXB1")
        strItems = Replace(strItems, "其它细胞2", "QTXB2")
        strItems = Replace(strItems, "其它细胞3", "QTXB3")
        strItems = Replace(strItems, "其它细胞4", "QTXB4")

        strItems = Replace(strItems, "管型", "GX")
        strItems = Replace(strItems, "其它", "QT")
        strItems = Replace(strItems, "霉菌", "MJ")
        strItems = Replace(strItems, "精子", "JZ")
        If rsSql("ifprint") = True Then
            strItems = Replace(strItems, "结晶", "JJZ")
        Else
            strItems = Replace(strItems, "结晶", "JJ")
        End If
        strItems = Replace(strItems, "细菌", "XJ")
        ReplaceItem = strItems
End Function






