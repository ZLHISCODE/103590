VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_blotray886"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:雷度自动蛋白印记仪
'厂商:雷度
'用户:汨罗人民医院
'作者：刘潇宇

'文件名为*.txt,文件名最好固定，通过配置文件来读取结果。
'雷度自动蛋白印记仪数据
'id , Hospital, HospitalDepartment, TestName, PatientNumber, Bed, PatientName, Sex, Age, PatientAddress, PatientTelephone, SourceTxt, SampleCode, SampleType, DataCollectd, DataReceived, DataTested, Doctor, Checker, ClinicalSymptom, SenderDepartment, SenderDoctor, zuhenumber, pic, T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, Name1, Name2, Name3, Name4, Name5, Name6, Name7, Name8, Name9, Name10, Name11, Name12, Name13, Name14, Name15, Name16, U1, U2, U3, U4, U5, U6, U7, U8, U9, U10, U11, U12, U13, U14, U15, U16
'10,汨罗市人民医院,JYK,,,,,,,,,,1,,,2016-05-17,2016-05-17,,,,,,综合组14项,,,,,0,0,1,0,0,0,0,0,0,0,0,0,0,阳性质控,阴性质控,总lgE,屋尘螨/粉尘螨,屋尘,柏树/榆树/梧桐/柳树/杨树,苦艾/艾蒿/豚草,点青霉/分枝孢霉/烟曲霉/交链孢霉/根霉/毛霉,猫毛皮屑/狗毛皮屑,蛋清/蛋黄,牛奶,鳕鱼/鲑鱼/鲈鱼,虾/蟹/扇贝,牛肉/羊肉,芒果/菠萝/苹果/桃子/草莓,花生/开心果/腰果/榛子,,,70.31,0.31,0.27,0.48,0.03,0.07,0.17,0.14,0.09,0.06,0.08,0.14,0.12,0.02
'11,汨罗市人民医院,JYK,,,,,,,,,,2,,,2016-05-17,2016-05-17,,,,,,综合组14项,,,,,0,0,0,0,2,1,0,0,0,0,0,0,0,阳性质控,阴性质控,总lgE,屋尘螨/粉尘螨,屋尘,柏树/榆树/梧桐/柳树/杨树,苦艾/艾蒿/豚草,点青霉/分枝孢霉/烟曲霉/交链孢霉/根霉/毛霉,猫毛皮屑/狗毛皮屑,蛋清/蛋黄,牛奶,鳕鱼/鲑鱼/鲈鱼,虾/蟹/扇贝,牛肉/羊肉,芒果/菠萝/苹果/桃子/草莓,花生/开心果/腰果/榛子,,,18.29,0.16,0.13,0.17,0.02,0.99,0.39,0.03,0.13,0.04,0.01,0.23,0.07,0.04
'12,汨罗市人民医院,JYK,,,,,,,,,,3,,,2016-05-17,2016-05-17,,,,,,综合组14项,,,,,0,0,0,0,0,1,0,0,0,0,0,0,0,阳性质控,阴性质控,总lgE,屋尘螨/粉尘螨,屋尘,柏树/榆树/梧桐/柳树/杨树,苦艾/艾蒿/豚草,点青霉/分枝孢霉/烟曲霉/交链孢霉/根霉/毛霉,猫毛皮屑/狗毛皮屑,蛋清/蛋黄,牛奶,鳕鱼/鲑鱼/鲈鱼,虾/蟹/扇贝,牛肉/羊肉,芒果/菠萝/苹果/桃子/草莓,花生/开心果/腰果/榛子,,,101.66,0.31,0.08,0.12,0.03,0.04,0.38,0.09,0.23,0.14,0.14,0.26,0.02,0.12
'13,汨罗市人民医院,JYK,,,,,,,,,,4,,,2016-05-17,2016-05-17,,,,,,综合组14项,,,,,0,0,0,0,0,0,0,0,0,0,0,0,0,阳性质控,阴性质控,总lgE,屋尘螨/粉尘螨,屋尘,柏树/榆树/梧桐/柳树/杨树,苦艾/艾蒿/豚草,点青霉/分枝孢霉/烟曲霉/交链孢霉/根霉/毛霉,猫毛皮屑/狗毛皮屑,蛋清/蛋黄,牛奶,鳕鱼/鲑鱼/鲈鱼,虾/蟹/扇贝,牛肉/羊肉,芒果/菠萝/苹果/桃子/草莓,花生/开心果/腰果/榛子,,,74.33,0.12,0.07,0.06,0.06,0.07,0.02,0.12,0.12,0.06,0.06,0.02,0.08,0.04
'14,汨罗市人民医院,JYK,,,,,,,,,,1,,,2016-05-17,2016-05-17,,,,,,综合组14项,,,,,0,0,1,0,0,0,0,0,0,0,0,0,0,阳性质控,阴性质控,总lgE,屋尘螨/粉尘螨,屋尘,柏树/榆树/梧桐/柳树/杨树,苦艾/艾蒿/豚草,点青霉/分枝孢霉/烟曲霉/交链孢霉/根霉/毛霉,猫毛皮屑/狗毛皮屑,蛋清/蛋黄,牛奶,鳕鱼/鲑鱼/鲈鱼,虾/蟹/扇贝,牛肉/羊肉,芒果/菠萝/苹果/桃子/草莓,花生/开心果/腰果/榛子,,,73.80,0.33,0.30,0.50,0.12,0.12,0.19,0.15,0.13,0.06,0.12,0.16,0.02,0.06
'15,汨罗市人民医院,JYK,,,,,,,,,,2,,,2016-05-17,2016-05-17,,,,,,综合组14项,,,,,0,0,0,0,2,1,0,0,0,0,0,0,0,阳性质控,阴性质控,总lgE,屋尘螨/粉尘螨,屋尘,柏树/榆树/梧桐/柳树/杨树,苦艾/艾蒿/豚草,点青霉/分枝孢霉/烟曲霉/交链孢霉/根霉/毛霉,猫毛皮屑/狗毛皮屑,蛋清/蛋黄,牛奶,鳕鱼/鲑鱼/鲈鱼,虾/蟹/扇贝,牛肉/羊肉,芒果/菠萝/苹果/桃子/草莓,花生/开心果/腰果/榛子,,,16.30,0.14,0.02,0.15,0.11,0.93,0.39,0.06,0.09,0.02,0.03,0.22,0.08,0.12
'16,汨罗市人民医院,JYK,,,,,,,,,,3,,,2016-05-17,2016-05-17,,,,,,综合组14项,,,,,1,0,0,0,0,1,0,0,0,0,0,0,0,阳性质控,阴性质控,总lgE,屋尘螨/粉尘螨,屋尘,柏树/榆树/梧桐/柳树/杨树,苦艾/艾蒿/豚草,点青霉/分枝孢霉/烟曲霉/交链孢霉/根霉/毛霉,猫毛皮屑/狗毛皮屑,蛋清/蛋黄,牛奶,鳕鱼/鲑鱼/鲈鱼,虾/蟹/扇贝,牛肉/羊肉,芒果/菠萝/苹果/桃子/草莓,花生/开心果/腰果/榛子,,,103.20,0.36,0.03,0.02,0.05,0.05,0.41,0.08,0.27,0.14,0.15,0.29,0.10,0.04
'17,汨罗市人民医院,JYK,,,,,,,,,,4,,,2016-05-17,2016-05-17,,,,,,综合组14项,,,,,0,0,0,0,0,0,0,0,0,0,0,0,0,阳性质控,阴性质控,总lgE,屋尘螨/粉尘螨,屋尘,柏树/榆树/梧桐/柳树/杨树,苦艾/艾蒿/豚草,点青霉/分枝孢霉/烟曲霉/交链孢霉/根霉/毛霉,猫毛皮屑/狗毛皮屑,蛋清/蛋黄,牛奶,鳕鱼/鲑鱼/鲈鱼,虾/蟹/扇贝,牛肉/羊肉,芒果/菠萝/苹果/桃子/草莓,花生/开心果/腰果/榛子,,,74.71,0.07,0.07,0.04,0.13,0.03,0.05,0.02,0.08,0.03,0.06,0.11,0.16,0.04

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
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strChenel As String, str操作员 As String, str日期 As String
    Dim strFilePath As String, strTmp As String

    Dim objStream As textStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str标本号 As String, str检验结果 As String, blnReturn As Boolean
    Dim lngRow As Long, varLine As Variant

    Dim intCol_U3 As Integer
    Dim intCol_U4 As Integer
    Dim intCol_U5 As Integer
    Dim intCol_U6 As Integer
    Dim intCol_U7 As Integer
    Dim intCol_U8 As Integer
    Dim intCol_U9 As Integer
    Dim intCol_U10 As Integer
    Dim intCol_U11 As Integer
    Dim intCol_U12 As Integer
    Dim intCol_U13 As Integer
    Dim intCol_U14 As Integer
    Dim intCol_U15 As Integer
    Dim intCol_U16 As Integer

    '日期，标本号
    Dim intCol_No As Integer
    Dim intDateS As Integer, intDateLen As Integer
    Dim intCol_Date As Integer

    On Error GoTo errHandle
        
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    '---读取配置
    intCol_No = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "标本号列", "12"))
    intCol_Date = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "日期列", "16"))
    intDateS = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "日期开始位置", "1"))
    intDateLen = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "日期长度", "10"))
    intCol_U3 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U3", "58"))
    intCol_U4 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U4", "59"))
    intCol_U5 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U5", "60"))
    intCol_U6 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U6", "61"))
    intCol_U7 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U7", "62"))
    intCol_U8 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U8", "63"))
    intCol_U9 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U9", "64"))
    intCol_U10 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U10", "65"))
    intCol_U11 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U11", "66"))
    intCol_U12 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U12", "67"))
    intCol_U13 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U13", "68"))
    intCol_U14 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U14", "69"))
    intCol_U15 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U15", "70"))
    intCol_U16 = Val(GetIniKeyValue(App.Path & "\bh5100.ini", "设置", "U16", "71"))
    
    '-----
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    Set objStream = objFile.OpenTextFile(strFilePath & "\all.csv", ForReading)
    
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If InStr(strLine, ",") > 0 Then
            varLine = Split(strLine, ",")
            If lngRow = 0 Then
                str日期 = Format(Now, "yyyy-MM-dd")
            Else
                str检验结果 = ""
                str日期 = Trim(varLine(intCol_Date))
                If Len(str日期) > intDateS + intDateLen Then
                    str日期 = Mid(str日期, intDateS, intDateLen)
                End If
                If Not IsDate(str日期) Then str日期 = Format(Now, "yyyy-MM-dd")
        
                str标本号 = Trim(varLine(intCol_No))
                If Val(str标本号) > 0 Then
                   
                    strTmp = Trim$(varLine(intCol_U3))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U3|" & strTmp
                    
                       strTmp = Trim$(varLine(intCol_U4))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U4|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U5))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U5|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U6))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U6|" & strTmp
                       
                    strTmp = Trim$(varLine(intCol_U7))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U7|" & strTmp
                       
                    strTmp = Trim$(varLine(intCol_U8))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U8|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U9))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U9|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U10))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U10|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U11))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U11|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U12))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U12|" & strTmp
                       
                    strTmp = Trim$(varLine(intCol_U13))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U13|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U14))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U14|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U15))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U15|" & strTmp
                    
                    strTmp = Trim$(varLine(intCol_U16))
                    If strTmp <> "" Then str检验结果 = str检验结果 & "|U16|" & strTmp
                    
                    If str检验结果 <> "" Then
                        If aResults(UBound(aResults)) <> "" Then
                            ReDim Preserve aResults(UBound(aResults) + 1)
                        End If
                        aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| | |0" & str检验结果
                    End If
               End If
            End If
        End If
        lngRow = lngRow + 1
        DoEvents
    Loop
    ResultFromFile = aResults
    
    Exit Function
errHandle:
'     WriteLog " BH5100," & Erl() & Err.Description
End Function

Private Function ChangeValue(ByVal strInValue As String) As String
    '转换结果值，规则为
    '值>0，结果有效，按原值返回；如果结果<=0,显示为'未检出'
    
    If Trim(strInValue) = "" Then
        ChangeValue = ""
    ElseIf Val(strInValue) > 0 Then
        ChangeValue = Trim(strInValue)
    Else
        ChangeValue = "未检出"
    End If
    
End Function



