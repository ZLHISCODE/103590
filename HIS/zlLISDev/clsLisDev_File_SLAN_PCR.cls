VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_File_SLAN_PCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:SLAN荧光定量PCR检测系统
'厂商:宏石
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
    
    Dim xlApp As Object            'EXCEL文件
    Dim xlBook As Object
    Dim xlSheet As Object
    
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str样本号 As String, str检验结果 As String, blnReturn As Boolean
    Dim lngRow As Long, varLine As Variant
    Dim strFileImage As String, strImage As String
    
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    Set xlApp = CreateObject("Excel.Application")
    
    '2012-03-11 ZHQ 设备提供文件为XLS文件格式
    strNewFile = strFilePath & "\slan.xls"
    If Dir(strNewFile) = "" Then
        Set xlBook = xlApp.Workbooks.Open(strFile)
    Else
        Set xlBook = xlApp.Workbooks.Open(strNewFile)
    End If
    Set xlSheet = xlBook.Worksheets(1)
    
    '孔号   通道    检测类型    检测项目(4)     属性    循环阈值    测量结果(7) 平均浓度    分析类型    阴性阈值    单位    CV值(%) 标本唯一标识(13)    姓名    性别    年龄
    'B3     1       阴性对照    HBV-DNA         No Ct   <1000       定量分析    1000        Copies/ml
    'B4     1       阳性对照    HBV-DNA         28.65   1.04E+05    定量分析    1000        Copies/ml
    'B5     1       样品        HBV-DNA         24.62   1.50E+06    定量分析    1000        Copies/ml                               120229HBV01         姜福海  男      51
    'B6     1       样品        HBV-DNA         24.96   1.20E+06    定量分析    1000        Copies/ml
    'B7     1       样品        HBV-DNA         21.14   1.51E+07    定量分析    1000        Copies/ml                               120229HBV02         孔繁奇  女      17
    'B8     1       样品        HBV-DNA         No Ct   <1000       定量分析    1000        Copies/ml
    
    str样本号 = ""
    str检验结果 = ""
    lngRow = 2
    Do While True
        strImage = ""
        If xlSheet.cells(lngRow, 4) <> "" And xlSheet.cells(lngRow, 7) <> "" And xlSheet.cells(lngRow, 13) <> "" Then
            If aResults(UBound(aResults)) <> "" Then
                ReDim Preserve aResults(UBound(aResults) + 1)
            End If
            str日期 = Mid(xlSheet.cells(lngRow, 13), 1, 6)
            str日期 = Mid(Format(date, "YYYY"), 1, 2) & str日期
            str日期 = Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7, 2)
            str样本号 = xlSheet.cells(lngRow, 13)
            If str样本号 <> "" Then
                str检验结果 = "|" & xlSheet.cells(lngRow, 4) & "|" & GetENum(Trim(xlSheet.cells(lngRow, 7)))
                aResults(UBound(aResults)) = str日期 & "|" & str样本号 & "| |血液|0" & str检验结果
                
            End If
        ElseIf xlSheet.cells(lngRow, 4) = "" And xlSheet.cells(lngRow, 7) = "" And xlSheet.cells(lngRow, 13) = "" Then
            Exit Do
        End If
        
        lngRow = lngRow + 1
        DoEvents
    Loop
    xlBook.Close
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing

'    If str样本号 <> Mid(aResults(UBound(aResults)), 12, 3) And 1 = 1 Then
'        ReDim Preserve aResults(UBound(aResults) + 1)
'        aResults(UBound(aResults)) = str日期 & "|" & str样本号 & "|" & str操作员 & "|" & str标本 & "|0" & str检验结果
'    End If
    ResultFromFile = aResults

End Function


Public Function GetENum(ByVal nNumber As Long) As String
'功能：根据传入的纯数字返回带相应小数的10的n次幂
'参数：nNumber：如：15120000,要求返回：1.512E+07
'返回：数值型数据
    
    Dim dbl10 As Double
    Dim i As Integer, j As Integer
    
    If nNumber >= 1 Then
        i = Len(CStr(nNumber))
        
        If i > 1 Then
            dbl10 = Round(nNumber / (10 ^ (i - 1)), 2)
            
            GetENum = dbl10 & "E+" & Right("00" & CStr(i - 1), 2)
        Else
            GetENum = nNumber
        End If
    Else
        '小于零的数据
        j = 1
        Do While True
            dbl10 = nNumber * 10 ^ j
            If dbl10 >= 1 Then
                dbl10 = Round(dbl10, 2)
                GetENum = dbl10 & "E-" & Right("00" & CStr(j), 2)
            End If
            j = j + 1
        Loop
    End If
End Function
