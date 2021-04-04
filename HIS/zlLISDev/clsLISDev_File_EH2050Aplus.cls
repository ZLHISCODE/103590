VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EH2050Aplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:迈瑞 EH-2050a plus 尿沉渣
'厂商:迈瑞
'作者:敬秦
'用户:双流县第二人民医院
'日期:2015-01-06

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
    Dim strCode As String, strValue() As String, strDate As String, strIteam As String, strSample As String, strLine As String, strData As String, strChnnl As String
    Dim lngS As Long, lngE As Long, i As Integer
    Dim objFile As New FileSystemObject
    Dim strFileName As String, strPath As String, objStream As textStream, strFilePath As String
    Dim rsResults() As String, strImage As String, strImageName As String
    Dim strStat As String '判断急诊标识
On Error GoTo hErr
    
   strPath = Mid(strFile, 1, InStrRev(strFile, "\"))
    ReDim rsResults(0)
    rsResults(0) = ""
    strFileName = Dir(strPath & "\", vbDirectory)       '取第一个
    Do While strFileName <> ""                    '不为空就循环
       If strFileName <> "." And strFileName <> ".." Then   '如果不是 当前目录及上一级目录 . 及 ..
             If GetAttr(strPath & "\" & strFileName) = vbDirectory Or GetAttr(strPath & "\" & strFileName) = vbArchive Then         '是否是目录
'                If Format(Now, "yyyymmdd") = Left(Split(strFileName, ".")(0), 8) Then '由于数据较多，目前只处理当前日期的数据
                If InStr(strFileName, Format(Now, "yyyymmdd")) Then
                    If objFile.FileExists(strPath & "\" & strFileName) Then  '查找是否存在此文件
                        Set objStream = objFile.OpenTextFile(strPath & "\" & strFileName)
                        Do Until objStream.AtEndOfStream
                            strLine = objStream.ReadLine
                            If InStr(strLine, "eBaseInfo") > 0 And InStr(strLine, "样本号") > 0 Then '取样本号
                                strSample = Split(strLine, ",")(3)
                                strSample = Replace$(strSample, ";", "")
                                strSample = Replace$(strSample, Chr(&H9), "")
                                strSample = Val(strSample)
                            End If
                            If InStr(strLine, "eBaseInfo") > 0 And InStr(strLine, "报告时间") > 0 Then '取时间
                                strDate = Split(strLine, ",")(3)
                                strDate = Replace$(strDate, ";", "")
                                strDate = Replace$(strDate, Chr(&H9), "")
                                strDate = Format(strDate, "yyyy-mm-dd hh:mm:ss")
                            End If
                            If InStr(strLine, "eBaseInfo") > 0 And InStr(strLine, "急诊") > 0 Then '判断急诊
                                strStat = Split(strLine, ",")(3)
                                strStat = Replace$(strStat, ";", "")
                                strStat = Replace$(strStat, Chr(&H9), "")
                                strStat = Trim(strStat)
                            End If
                            If InStr(strLine, "eBaseInfo") > 0 And InStr(strLine, "颜色") > 0 Then
                                strChnnl = "颜色"
                                strData = " " & Trim(Replace(Replace(Split(strLine, ",")(3), ";", ""), Chr(&H9), ""))
                                strIteam = strIteam & "|" & strChnnl & "|" & strData
                            End If
                            If InStr(strLine, "eBaseInfo") > 0 And InStr(strLine, "浊度") > 0 Then
                                strChnnl = "浊度"
                                strData = " " & Trim(Replace(Replace(Split(strLine, ",")(3), ";", ""), Chr(&H9), ""))
                                strIteam = strIteam & "|" & strChnnl & "|" & strData
                            End If
                            If InStr(strLine, "eDryInfo") > 0 And UBound(Split(strLine, ",")) >= 5 Then '获取尿干化学数据
                                strChnnl = Trim(Replace(Split(strLine, ",")(2), Chr(&H9), ""))
                                strData = " " & Trim(Replace(Split(strLine, ",")(3), Chr(&H9), ""))
                                strIteam = strIteam & "|" & Replace(strChnnl, "白细胞", "白细胞1") & "|" & strData
                            End If
                            If InStr(strLine, "eNCZInfo") > 0 And UBound(Split(strLine, ",")) >= 5 Then '获取尿镜检结果
                                strChnnl = Trim(Replace(Split(strLine, ",")(2), Chr(&H9), ""))
                                strData = " " & Trim(Replace(Split(strLine, ",")(3), Chr(&H9), ""))
                                strIteam = strIteam & "|" & strChnnl & "|" & strData
                            End If
                            If InStr(strLine, "eImage") > 0 Then '获取尿图像数据
                                strImageName = Split(Split(strLine, ",")(1), "=")(1)
                                strImage = strImage & "^" & strImageName & ";113;0;" & Replace(Split(strLine, ",")(2), ";", "")
                            End If
                        Loop
                        If strIteam <> "" And strImage <> "" And strSample <> "" Then
                            strIteam = Replace(strIteam, "leu/uL", "")
                            strIteam = Replace(strIteam, "mmol/L", "")
                            strIteam = Replace(strIteam, "g/L", "")
                            strIteam = Replace(strIteam, "ery/uL", "")
                            strIteam = Replace(strIteam, "个/μL", "")
                            If rsResults(UBound(rsResults)) <> "" Then ReDim Preserve rsResults(UBound(rsResults) + 1)
                            rsResults(UBound(rsResults)) = strDate & "|" & IIf(strStat = "False", strSample, strSample & "^1^") & "| |尿液|0" & strIteam
                            strIteam = ""
                            If strImage <> "" Then '加上图片
                                strImage = Mid(strImage, 2)
                                strImage = Replace$(strImage, "jpg", "bmp")
                                rsResults(UBound(rsResults)) = rsResults(UBound(rsResults)) & vbCrLf & Replace(strImage, "Image", "")
                                Call writelislog(rsResults(UBound(rsResults)))
                                strImage = ""
                            End If
                        End If
                    End If
                End If
            End If
         End If
         strFileName = Dir      '取下一个
    Loop
    ResultFromFile = rsResults
    Exit Function
hErr:
    MsgBox "zlCDLISDev.clsLISDev_File_EH-2050Aplus|ResultFromFile" & Err.Description, vbInformation, "接口错误"
    Err.Clear
End Function

Public Function writelislog(strLog As String)
    Const strFile As String = "LisLog_EH2080C_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim strFileName As String
    strFileName = App.Path & "\" & strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function


