VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_XE5000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:带流水线的 XE-5000
'厂商:希森美康(SysMex)
'作者:陈东
'用户:晋煤集团医院

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
    
    Dim aResults() As String
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String

    Dim str日期 As String, str项目 As String, str标本号 As String, strImg As String
    Dim strFilePath As String, strFilename As String, strImgPath As String, strImgFilePath As String
    Dim arrLine
    Call ReadConf(App.Path & "\XE5000.ini", "IMG_PATH", strImgPath)
    
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    strFilename = Dir(strFilePath & "\*.CDF")
    ReDim aResults(0)
    
    Do While strFilename <> ""
        
        If objFile.FileExists(strFilePath & "\" & strFilename) Then
            Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename)
            
            Do While Not objStream.AtEndOfStream
                strLine = UCase(objStream.ReadLine)
                arrLine = Split(strLine, ",")
                If arrLine(0) = "0" Then
                    str项目 = "": strImg = ""
                    str日期 = Split(strLine, ",")(1) & " " & Split(strLine, ",")(2)
                    If CDate(str日期) >= dtStart And CDate(str日期) <= dtEnd Then
                        str标本号 = Split(strLine, ",")(3)
                        If lngSampleNO <> -1 Then
                            If lngSampleNO <> Val(str标本号) Then
                                str日期 = "": str标本号 = ""
                                Exit Do
                            End If
                        End If
                    Else
                        str日期 = "": str标本号 = ""
                        Exit Do
                    End If
                    
                ElseIf arrLine(0) = "1" Then
                    '检验结果
                    str项目 = str项目 & "|" & Split(strLine, ",")(1) & "|" & Split(strLine, ",")(3)
                ElseIf arrLine(0) = "3" Then
                    '替换图形路径
                    strImgFilePath = Split(strLine, ",")(3)
                    If InStr(strImgPath, ",") > 0 Then
                        strImgFilePath = Replace(strImgFilePath, Trim(Split(strImgPath, ",")(0)), Trim(Split(strImgPath, ",")(1)))
                    End If
                    If InStr("HRBC,HPLT", Split(strLine, ",")(2)) > 0 Then
                        strImg = strImg & "^" & Split(strLine, ",")(2) & ";105;0;" & strImgFilePath
                    Else
                        strImg = strImg & "^" & Split(strLine, ",")(2) & ";104;0;" & strImgFilePath
                    End If
                ElseIf arrLine(0) = "9" Then
                    Exit Do
                End If
            Loop
            
            If str项目 <> "" And str日期 <> "" And str标本号 <> "" Then
                str项目 = Mid(str项目, 2)
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                If Len(str标本号) >= 12 Then str标本号 = "^^" & str标本号
                aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血清|0|" & str项目
                If strImg <> "" Then
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strImg, 2)
                End If
                
                str项目 = "": str日期 = "": str标本号 = "": strImg = ""
            End If
            
            objStream.Close
            Set objStream = Nothing
            
            '读了就把文件放到BAK下
            If Not objFile.FolderExists(strFilePath & "\Bak") Then Call objFile.CreateFolder(strFilePath & "\Bak")
            objFile.CopyFile strFilePath & "\" & strFilename, strFilePath & "\Bak\" & strFilename, True
            objFile.DeleteFile strFilePath & "\" & strFilename, True
            
        End If
        
        strFilename = Dir
    Loop
    ResultFromFile = aResults
End Function

Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
    'iSendStep：当前发送操作的步骤。如果处理完后发送结束，则返回0，否则自动＋1返回。
    'strResponse：当前仪器的响应内容
    'strSampleInfo：发送的标本信息。
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否急诊
    '   第5个无素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。

    Dim var_标本 As Variant, i_Count As Integer, i_Loop As Integer, str_记录 As String, str_标本号 As String
    Dim i_急诊 As Integer, str_条码 As String, str_检验项目 As String, str_标本类型 As String
    Dim str_SendItem As String, str_EndString As String, str_Tmp As String, lngRow As Long
    Dim str参数 As String, arr参数(3) As String, strCDF As String
    
    If strSampleInfo <> "" Then WriteLog "双向通讯(Xe5000 File Begin)：", strSampleInfo, ""
    var_标本 = Split(strSampleInfo, "||")
    SendSample = ""
   
    Call ReadConf(App.Path & "\XE5000.ini", "SAVECDF_PATH,RET_PATH", str参数)
    If UBound(Split(str参数, ",")) = 1 Then
        arr参数(0) = Split(str参数, ",")(0)
        arr参数(1) = Split(str参数, ",")(1)

    Else
        arr参数(0) = App.Path
        arr参数(1) = App.Path
    End If
    For i_Count = LBound(var_标本) To UBound(var_标本)
        str_记录 = var_标本(i_Count)
        
        If Val(Split(str_记录, "|")(1)) = 0 Then
            str_标本号 = Trim(Split(str_记录, "|")(5))
        Else
            str_标本号 = Val(Split(str_记录, "|")(1))
        End If
        str_标本类型 = Split(str_记录, "|")(3)
        i_急诊 = Val(Split(str_记录, "|")(4))
        str_条码 = Trim(Split(str_记录, "|")(5))
        '固定
        str_检验项目 = ""
        
        For i_Loop = 10 To UBound(Split(str_记录, "|"))
            If Trim(Split(str_记录, "|")(i_Loop)) <> "" Then
                str_Tmp = Trim(Split(str_记录, "|")(i_Loop))
                Call ReadConf(App.Path & "\XE5000.ini", str_Tmp, str参数)
                If str参数 <> "" Then
                    str_检验项目 = str_检验项目 & Trim(str参数) & "+"
                End If
            End If
        Next
        
        If str_检验项目 <> "" And str_条码 <> "" Then
            str_检验项目 = Mid(str_检验项目, 1, Len(str_检验项目) - 1)
            SendSample = SendSample & str_条码 & "," & str_检验项目 & "|"
            
           '标本日期(19)*,仪器代号(20)*,标本号(20)*,标本类型,费别,送检科室,送检医生,检验医生,病历号*,病人姓名*,
           '性别 (2),病人类型,床号,年龄,年龄单位,申请号,送检日期,报告日期,打印标志,结果标志,
           '错误标志,诊断,备注,保留字段,姓名拼音码,接收时间/采样日期,病区,条形码
           strCDF = Split(str_记录, "|")(0) & ",XE-5000," & str_条码 & ",,,,,," & str_条码 & "," & str_标本号 & ","
           strCDF = strCDF & ",,,,,,,,,," & ",,,,,,," & str_条码
           str_Tmp = Trim(arr参数(0) & "\" & Format(Now, "yyyyMMdd") & str_条码 & Format(Now, "HHmmss") & ".CDF")
           Call WritTxt(str_Tmp, strCDF)
        End If
        
    Next
    If SendSample <> "" Then
        str_Tmp = Trim(arr参数(1) & "\RET.TXT")
        Call WritTxt(str_Tmp, SendSample)
        SendSample = ""
    End If
        
End Function

Private Function WritTxt(ByVal strFilename As String, strInfo As String)
    Dim objFileSystem As New FileSystemObject
    Dim objStream As TextStream
    
    If Not objFileSystem.FileExists(strFilename) Then Call objFileSystem.CreateTextFile(strFilename)
    Set objStream = objFileSystem.OpenTextFile(strFilename, ForAppending)
    objStream.WriteLine (strInfo)
    objStream.Close
    Set objStream = Nothing
End Function

Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '读配置文件
    'strIniFile :ini配置文件名
    'strItems   :要读取的字段，如：DATABASE NAME=th_xlb中的 DATABASE NAME,多个字段以,分隔
    'strReturn  :读取的返回值，如:DATABASE NAME=th_xlb中的 th_xlb ，多个反回值以,分隔
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As TextStream

    Dim intloop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '从配置文件中读取保存在文件中的配置cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intloop = 0 To intCount
        arrItems(intloop, 0) = Split(strItems, ",")(intloop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intloop = 0 To intCount
                    If Trim(arrLine(0)) = arrItems(intloop, 0) Then
                        arrItems(intloop, 1) = arrLine(1)
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intloop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intloop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If
    
    
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub


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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim lng_Stx As Long, lng_Etx As Long, bln_QC As Boolean
    Dim str_标本号 As String, str_日期 As String, int_Index As Integer
    Dim str_Code As String, str_Line As String, str_Item As String
    
    Dim str_Path As String, str_File  As String, str_imgName As String, str_RetuImage As String
    Dim bln_Rec As Boolean
    strCmd = Chr(&H6)
    strReserved = strSource
    If Not strReserved Like "*DI*" & Chr(&H3) & "DR*" & Chr(&H3) & "*" Then Exit Sub
    lng_Stx = InStr(strReserved, "DI")
    lng_Etx = InStr(lng_Stx, strReserved, Chr(&H3))
    
    str_Code = Mid(strReserved, lng_Stx + 1, lng_Etx - lng_Stx - 1)
    
    lng_Stx = InStr(strReserved, "DR")
    lng_Etx = InStr(lng_Stx, strReserved, Chr(&H3))
    
    str_Code = str_Code & Mid(strReserved, lng_Stx + 1, lng_Etx - lng_Stx - 1)
    
    strReserved = Mid(strReserved, lng_Etx + 1)
    strResult = ""
    If Left(str_Code, 2) = "DI" Then
        bln_QC = False
    ElseIf ((Left(str_Code, 3) = "D1C") Or (Left(str_Code, 3) = "D2C")) Then
        bln_QC = True
    End If
    If bln_QC = False And Left(str_Code, 2) = "DI" Then
        str_标本号 = Mid(str_Code, 65, 15)
        If Val(str_标本号) <= 0 Then Exit Sub
        
        str_日期 = Mid(str_Code, 43, 4)
        str_日期 = str_日期 & "-" & Mid(str_Code, 47, 2)
        str_日期 = str_日期 & "-" & Mid(str_Code, 49, 2)
        If IsDate(str_日期) Then
            str_日期 = Format(str_日期, "yyyy-MM-dd")
        Else
            str_日期 = Format(Now, "yyyy-MM-dd")
        End If
        
        str_Line = Mid(str_Code, InStr(str_Code, "D2U"), 216)
        str_Line = str_Line & Mid(str_Code, InStr(str_Code, "D6U"), 255)
        
        If Len(str_Line) >= 471 Then
            str_Item = "|WBC|" & Val(Mid(str_Line, 11, 6)) / 10 ^ 3
            str_Item = str_Item & "|RBC|" & GetValues(Mid(str_Line, 17, 5), 3)
            str_Item = str_Item & "|HGB|" & GetValues(Mid(str_Line, 22, 5), 1)
            str_Item = str_Item & "|HCT|" & GetValues(Mid(str_Line, 27, 5), 2)
            str_Item = str_Item & "|MCV|" & GetValues(Mid(str_Line, 32, 5), 2)
            str_Item = str_Item & "|MCH|" & GetValues(Mid(str_Line, 37, 5), 2)
            str_Item = str_Item & "|MCHC|" & GetValues(Mid(str_Line, 42, 5), 1)
            str_Item = str_Item & "|PLT|" & GetValues(Mid(str_Line, 47, 5), 1)
            str_Item = str_Item & "|LYMPH%|" & GetValues(Mid(str_Line, 52, 5), 2)
            str_Item = str_Item & "|MONO%|" & GetValues(Mid(str_Line, 57, 5), 2)
            str_Item = str_Item & "|NEUT%|" & GetValues(Mid(str_Line, 62, 5), 2)
            str_Item = str_Item & "|EO%|" & GetValues(Mid(str_Line, 67, 5), 2)
            str_Item = str_Item & "|BASO%|" & GetValues(Mid(str_Line, 72, 5), 2)
            str_Item = str_Item & "|LYMPH#|" & GetValues(Mid(str_Line, 77, 6), 3)
            str_Item = str_Item & "|MONO#|" & GetValues(Mid(str_Line, 83, 6), 3)
            str_Item = str_Item & "|NEUT#|" & GetValues(Mid(str_Line, 89, 6), 3)
            str_Item = str_Item & "|EO#|" & GetValues(Mid(str_Line, 95, 6), 3)
            str_Item = str_Item & "|BASO#|" & GetValues(Mid(str_Line, 101, 6), 3)
            str_Item = str_Item & "|RDW-CV|" & GetValues(Mid(str_Line, 107, 5), 2)
            str_Item = str_Item & "|RDW-SD|" & GetValues(Mid(str_Line, 112, 5), 2)
            str_Item = str_Item & "|PDW|" & GetValues(Mid(str_Line, 117, 5), 2)
            str_Item = str_Item & "|MPV|" & GetValues(Mid(str_Line, 122, 5), 2)
            str_Item = str_Item & "|P-LCR|" & GetValues(Mid(str_Line, 127, 5), 2)
            
            bln_Rec = Not (Val(Mid(str_Line, 132, 5)) / 10 ^ 3 = 0)
            str_Item = str_Item & "|RET%|" & GetValues(Mid(str_Line, 132, 5), 3)
            str_Item = str_Item & "|RET#|" & GetValues(Mid(str_Line, 137, 5), 5)
            str_Item = str_Item & "|IRF|" & GetValues(Mid(str_Line, 142, 5), 2)
            str_Item = str_Item & "|LFR|" & GetValues(Mid(str_Line, 147, 5), 2)
            str_Item = str_Item & "|MFR|" & GetValues(Mid(str_Line, 152, 5), 2)
            str_Item = str_Item & "|HFR|" & GetValues(Mid(str_Line, 157, 5), 2)
            
            str_Item = str_Item & "|PCT|" & GetValues(Mid(str_Line, 162, 5), 3)
            str_Item = str_Item & "|NRBC%|" & GetValues(Mid(str_Line, 167, 6), 1)
            str_Item = str_Item & "|NRBC#|" & GetValues(Mid(str_Line, 173, 6), 2)
            
            str_Item = str_Item & "|IG#|" & GetValues(Mid(str_Line, 179, 6), 3)
            str_Item = str_Item & "|IG%|" & GetValues(Mid(str_Line, 185, 5), 2)
            str_Item = str_Item & "|HPC#|" & GetValues(Mid(str_Line, 190, 6), 2)
            str_Item = str_Item & "|RET-He|" & GetValues(Mid(str_Line, 196, 5), 2)
            str_Item = str_Item & "|IPF|" & GetValues(Mid(str_Line, 201, 5), 2)
            
            str_Item = str_Item & "|WBC-B|" & GetValues(Mid(str_Line, 228, 5), 3)
            str_Item = str_Item & "|WBC-D|" & GetValues(Mid(str_Line, 234, 5), 3)
            str_Item = str_Item & "|NEUT#&|" & GetValues(Mid(str_Line, 240, 5), 3)
            str_Item = str_Item & "|NEUT%&|" & GetValues(Mid(str_Line, 246, 4), 2)
            str_Item = str_Item & "|LYMP#&|" & GetValues(Mid(str_Line, 251, 5), 3)
            str_Item = str_Item & "|LYMP%&|" & GetValues(Mid(str_Line, 257, 4), 2)
            str_Item = str_Item & "|HFLC#|" & GetValues(Mid(str_Line, 262, 5), 3)
            str_Item = str_Item & "|HFLC%|" & GetValues(Mid(str_Line, 268, 4), 2)
            str_Item = str_Item & "|Area#|" & GetValues(Mid(str_Line, 273, 5), 3)
            str_Item = str_Item & "|Area%|" & GetValues(Mid(str_Line, 279, 4), 2)
            str_Item = str_Item & "|NRBC+W|" & GetValues(Mid(str_Line, 284, 5), 3)
            str_Item = str_Item & "|NEUT-X|" & GetValues(Mid(str_Line, 289, 5), 2)
            str_Item = str_Item & "|RBC-O|" & GetValues(Mid(str_Line, 295, 4), 3)
            str_Item = str_Item & "|PLT-O|" & GetValues(Mid(str_Line, 300, 4), 1)
            str_Item = str_Item & "|PLT-I|" & GetValues(Mid(str_Line, 305, 4), 1)
            str_Item = str_Item & "|HF-BF#|" & GetValues(Mid(str_Line, 310, 6), 1)
            str_Item = str_Item & "|HF-BFB|" & GetValues(Mid(str_Line, 317, 4), 2)
            str_Item = str_Item & "|TC-BF#|" & GetValues(Mid(str_Line, 322, 6), 1)
            str_Item = str_Item & "|FRC#|" & GetValues(Mid(str_Line, 340, 4), 4)
            str_Item = str_Item & "|FRC%|" & GetValues(Mid(str_Line, 345, 4), 3)
            str_Item = str_Item & "|RBC-He|" & GetValues(Mid(str_Line, 350, 4), 2)
            str_Item = str_Item & "|Delta-He|" & GetValues(Mid(str_Line, 355, 5), 2)
            str_Item = str_Item & "|RET-Y|" & GetValues(Mid(str_Line, 360, 5), 2)
            str_Item = str_Item & "|RBC-Y|" & GetValues(Mid(str_Line, 365, 5), 2)
            str_Item = str_Item & "|IRF-Y|" & GetValues(Mid(str_Line, 370, 5), 2)
            str_Item = str_Item & "|RPI|" & GetValues(Mid(str_Line, 376, 4), 1)
            str_Item = str_Item & "|LScRBC|" & GetValues(Mid(str_Line, 381, 4), 2)
            str_Item = str_Item & "|HScRBC|" & GetValues(Mid(str_Line, 386, 4), 2)
            str_Item = str_Item & "|MicroRBC|" & GetValues(Mid(str_Line, 391, 4), 2)
            str_Item = str_Item & "|MacroRBC|" & GetValues(Mid(str_Line, 396, 4), 2)
            str_Item = str_Item & "|H-IPF|" & GetValues(Mid(str_Line, 401, 4), 2)
            str_Item = str_Item & "|IPF#|" & GetValues(Mid(str_Line, 406, 4), 2)
            str_Item = str_Item & "|PLT-X|" & GetValues(Mid(str_Line, 411, 4), 2)
            
        End If
        
        
        If Val(str_标本号) > 0 And str_Item <> "" Then
            str_标本号 = Val(str_标本号)
            strResult = str_日期 & "|" & IIf((Len(str_标本号) >= 12), "^^", "") & str_标本号 & "| |血清|0" & str_Item
        Else
            Exit Sub
        End If
        
        '图形数据
        str_Path = App.Path & "\Gram"
        If Dir(str_Path, vbDirectory) = "" Then
            MkDir str_Path
        End If
        If InStr(str_Code, "D3U") > 0 Then
            str_imgName = "HRBC"
            str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D3U"))
        End If
        If InStr(str_Code, "D4U") > 0 Then
            str_imgName = "HPLT"
            str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D4U"))
        End If
        If InStr(str_Code, "D5U") > 0 Then
            str_imgName = "WBC DISCRI"
            str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D5U"))
        End If
        If InStr(str_Code, "D1G") > 0 Then
            str_imgName = "SDIFF"
            str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D1G"))
        End If
        If InStr(str_Code, "D2G") > 0 Then
            str_imgName = "SBASO"
            str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D2G"))
        End If
        If InStr(str_Code, "D3G") > 0 Then
            str_imgName = "SRET"
            str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D3G"))
        End If
        If InStr(str_Code, "D4G") > 0 Then
'            If bln_Rec Then
                str_imgName = "SPLT"
                str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D4G"))
'            End If
        End If
        If InStr(str_Code, "D5G") > 0 Then
            str_imgName = "SIMI"
            str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D5G"))
        End If
        If InStr(str_Code, "D6G") > 0 Then
            str_imgName = "SRET-E"
            str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D6G"))
        End If
        If InStr(str_Code, "D7G") > 0 Then
            str_imgName = "SNRBC"
            str_File = Replace(str_日期, "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, str_Code, InStr(str_Code, "D7G"))
        End If
        If str_RetuImage <> "" Then strResult = strResult & vbCrLf & Mid(str_RetuImage, 2)
    Else
        '没数据，暂时不处理
    End If
    
End Sub

Private Function Save_image(ByVal str_img As String, ByVal str_FileName As String, ByVal str_Data As String, ByVal lng_Start As Long) As String
    Dim lng_Lens As Long, str_ImgData As String, T_ImgInfo As typHistGramInfo
    
    If lng_Start <= 0 Then Exit Function
    
    If str_img = "HRBC" Or str_img = "HPLT" Or str_img = "WBC DISCRI" Then
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 12
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 41, lng_Lens)
        T_ImgInfo.lnglower = Mid(str_Data, lng_Start + 29, 4)
        T_ImgInfo.lngupper = Mid(str_Data, lng_Start + 33, 4)
        T_ImgInfo.lngmaxx = Mid(str_Data, lng_Start + 16, 3)
        T_ImgInfo.lngmaxy = Mid(str_Data, lng_Start + 19, 3)
        T_ImgInfo.lngresver1 = 0: T_ImgInfo.lngresver2 = 0
        T_ImgInfo.lngstoppos = IIf(str_img = "PLT DISCRI", 40, 46)
        
        
        If uf_xehist(str_ImgData, lng_Lens, T_ImgInfo, "c:\temp1.bmp", str_FileName) = 1 Then
             Save_image = "^" & str_img & ";105;1;" & str_FileName
        End If
    Else
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 1
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 29, lng_Lens)
        If uf_xescat(str_ImgData, lng_Lens, "c:\Temp2.bmp", str_FileName) = 1 Then
            Save_image = "^" & str_img & ";104;1;" & str_FileName
        End If
    End If
    
End Function

Private Function GetValues(ByVal str_in As String, int_i As Integer) As String
    '功能:格式化返回值
    'str_in:待格式化的字符串
    'int_i:当入参是数字时,保留的小数位数
    
    If str_in = "*0000" Then
        GetValues = "----"
    ElseIf str_in = "     " Then
        GetValues = " "
    Else
        GetValues = Val(str_in) / 10 ^ int_i
    End If
    
End Function
