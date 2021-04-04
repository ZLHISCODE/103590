VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LJ3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
Option Explicit
'仪器:粪便仪 LJ-3000
'厂商:济南兰洁生物技术有限公司
'作者:蒋淑君
'用户:德阳人民医院
'时间:2015-07-21

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
    Dim lngResultCount As Long
    Dim strIniFile As String, strImg As String, lngPath As Long, strPath As String, strFileName As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.textStream
    Dim SERVERName As String, DatabaseName As String, USERName As String, PWD As String, PICADRESS As String 'SQL SERVER连接字符串
    
    On Error GoTo errHandle
    strIniFile = App.Path & "\LJ3000.ini"
    
    If Not objFileSystem.FileExists(strIniFile) Then
        Set objStream = objFileSystem.CreateTextFile(strIniFile)
        objStream.WriteLine "[Options]"
        objStream.WriteLine "[DATABASE]"
        objStream.WriteLine "SERVERNAME=."
        objStream.WriteLine "DATABASENAME=jnlj_db"
        objStream.WriteLine "USERNAME=SA"
        objStream.WriteLine "PASSWORD=123456"
        objStream.WriteLine "PICADRESS=D:\Program Files\济南兰洁第三代大便分析仪\Pictures"
        objStream.Close
        Set objStream = Nothing
    End If
    
    '-----SQL SERVER连接字符串
    SERVERName = ReadIni("DATABASE", "SERVERNAME", strIniFile)
    DatabaseName = ReadIni("DATABASE", "DATABASENAME", strIniFile)
    USERName = ReadIni("DATABASE", "USERNAME", strIniFile)
    PWD = ReadIni("DATABASE", "PASSWORD", strIniFile)
    PICADRESS = ReadIni("DATABASE", "PICADRESS", strIniFile)
    
    
    If (Trim(SERVERName) = "") Or (Trim(DatabaseName) = "") Or (Trim(USERName) = "") Then
         WriteLog "BH5500S_sql_ResultFromFile", strIniFile, "连接SQL SERVER数据库错误!错误原因：服务器名为空或数据库名为空或用户名为空!"
         Exit Function
    End If
    
    With cnAccess
        .ConnectionString = "Provider=SQLOLEDB.1;Initial Catalog=" & DatabaseName & ";Password=" & PWD & _
        ";Persist Security Info=True;User ID=" & USERName & ";Data Source=" & SERVERName
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "select id as  标本号,Date  As 检验日期,YinXue  as 隐血检验结果," & vbNewLine & _
            " Nxb, Zfd , Dfkl, Mj, Zwxb,Jj,Spxb, Jscyc, Jdzz ,Jmyj ,Jscl ,Jxw, Lzbd,rbc_value , wbc_value As wbc_value, tsbc_value As tsbc_value" & vbNewLine & _
            "from jnlj_table where Date>='" & Format(dtStart, "yyyy-mm-dd") & "' and Date<='" & Format(dtEnd, "yyyy-mm-dd") & "' and  yq_do=1"
''yq_do   仪器是否检测    0：未检测 1：已检测
''ys_do   医生是否审核    0：未审核 1：已审核
''yx_test 标本是否做隐血测试  0：不测试 1： 测试
''lzbd_test   标本是否做轮状病毒测试  0：不测试 1： 测试
      
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
             .Close
             Exit Function
        End If
        lngResultCount = -1
        
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            
            aResults(lngResultCount) = Format(rsTmp("检验日期"), "yyyy-MM-dd hh:mm:ss") & "|" & Trim(rsTmp("标本号")) & "| |粪便|0|Nxb|" & Decode(rsTmp("Nxb"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Nxb"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Zfd|" & Decode(rsTmp("Zfd"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Zfd"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Dfkl|" & Decode(rsTmp("Dfkl"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Dfkl"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Mj|" & Decode(rsTmp("Mj"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Mj"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Zwxb|" & Decode(rsTmp("Zwxb"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Zwxb"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jj|" & Decode(rsTmp("Jj"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jj"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Spxb|" & Decode(rsTmp("Spxb"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Spxb"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jscyc|" & Decode(rsTmp("Jscyc"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jscyc"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jdzz|" & Decode(rsTmp("Jdzz"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jdzz"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jmyj|" & Decode(rsTmp("Jmyj"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jmyj"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jscl|" & Decode(rsTmp("Jscl"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jscl"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Jxw|" & Decode(rsTmp("Jxw"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Jxw"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|Lzbd|" & Decode(rsTmp("Lzbd"), "0", "未见", "1", "偶见", "2", "少许", "3", "全视野", "4", "+", "5", "++", "6", "+++", "7", "++++", rsTmp("Lzbd"))
            aResults(lngResultCount) = aResults(lngResultCount) & "|rbc_value|" & rsTmp("rbc_value") & "|wbc_value|" & rsTmp("wbc_value") & "|tsbc_value|" & rsTmp("tsbc_value")
            '取图片路径 strFilePath
            
            strFilePath = PICADRESS & "\" & rsTmp("检验日期") & "\" & rsTmp("标本号") & "\" '"2015-07-22" & "\" & 1 & "\" '
            lngPath = InStrRev(strFilePath, "\") '找到最后一个\
            strPath = Mid$(strFilePath, 1, lngPath)
            strFileName = Dir(strPath, vbDirectory)
            Do While strFileName <> ""
                If strFileName <> "." And strFileName <> ".." Then
                    If GetAttr(strPath & "\" & strFileName) = vbDirectory Or GetAttr(strPath & "\" & strFileName) = vbArchive Then '是否是目录
                        If strFileName Like "*.bmp" Then    '查找图片
                            strImg = strImg & "^" & "P" & Replace(Split(strFileName & ".", ".")(0), "-", "") & ";103;1;" & strPath & strFileName
                        End If
                    End If
                End If
                strFileName = Dir
            Loop
            If strImg <> "" Then
                aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & Mid(strImg, 2)
                strImg = ""
            End If
            
            Debug.Print aResults(UBound(aResults))
            Call WriteLog(" aResults(lngResultCount)", aResults(lngResultCount), "")
            rsTmp.MoveNext
        Loop
    End With
    
    If lngResultCount = -1 Then
        ReDim Preserve aResults(0)
    End If
    
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "LJ-3000_ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function


Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function


