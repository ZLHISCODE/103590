VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_File_FJ2008PS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:γ放射免疫计数器FJ-2008PS
'厂商:国营二六二厂
'作者:周海全
'用户:丹东市中心医院
'时间:2012-03-11
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
'       如：第5i个元素为检验项目，第5i+1个元素为检验结果("^" & OD值 & "^" & CO值 & "^" & S/Co)
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim str标本 As String, str操作员 As String, str日期 As String, strDate As String
    Dim strFilePath As String, strNewFile As String
    
    Dim objStream As TextStream, aItem() As String
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str样本号 As String, str检验结果 As String, blnReturn As Boolean
    
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    '初始标本数组
    ReDim aItem(0): aItem(0) = ""
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    '2012-03-11 ZHQ 此设备的文件名固定为SendData.txt
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    str样本号 = ""
    str检验结果 = ""
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        
        '标准格式串如下：
        '    1    β2MG              3.27     120221    1.75--2.61   mg/L
        '    2    β2MG              2.37     120221    1.75--2.61   mg/L
        '    3    β2MG              4.57     120221    1.75--2.61   mg/L

        str日期 = Trim(Mid(strLine, 36, 8))
        str日期 = Mid(Format(date, "YYYY"), 1, 2) & str日期
        str日期 = Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7, 2)
            
        str样本号 = Trim(Mid(strLine, 1, 5))
        If Not IsNumeric(str样本号) Then str样本号 = 9999
        
        varItem = Trim(Mid(strLine, 6, 8))
        varTmp = Trim(Mid(strLine, 16, 18))
        str检验结果 = "|" & varItem & "|" & varTmp
            
        If aResults(UBound(aResults)) <> "" Then
            ReDim Preserve aResults(UBound(aResults) + 1)
        End If
        If str样本号 <> "" Then
            aResults(UBound(aResults)) = str日期 & "|" & str样本号 & "| |血液|0" & str检验结果
        End If
        DoEvents
    Loop

    ResultFromFile = aResults
End Function


