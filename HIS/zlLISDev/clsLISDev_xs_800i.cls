VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_xs_800i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:xs-800i
'厂商:Sysmex
'作者:陈东
'修改人:李建飞
'修改日期:2009-12-24
'修改内容:加入图形解析
'用户:成都儿童专科医院(东城根)

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = Chr(6)
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    Dim varLine As Variant, strItem As String, str_日期 As String, str_标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, str_Path As String, str_imgName As String, str_File As String
    Dim str_RetuImage As String, blnQc As Boolean
    Dim bln_Rec As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "***" Then
        bln_Rec = False
        strReserved = strSource
        Do While Not bln_Rec
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            strLine = Mid(strReserved, InStr(strReserved, ""), InStr(strReserved, ""))
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            
            str_日期 = "": str_标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            '解析
            varLine = Split(strLine, vbCrLf)
            strItem = ""
            For iRow = LBound(varLine) To UBound(varLine)
                If varLine(iRow) Like "DI*" Then
                    str_日期 = Mid(varLine(iRow), 44, 4) & "-" & Mid(varLine(iRow), 48, 2) & "-" & Mid(varLine(iRow), 50, 2)
                    str_日期 = str_日期 & " " & Mid(varLine(iRow), 52, 2) & ":" & Mid(varLine(iRow), 54, 2) & ":" & Mid(varLine(iRow), 56, 2)
                    str_标本号 = Trim(Mid(varLine(iRow), 66))
                    If Val(str_标本号) = 0 Then blnQc = True
                    str检验人标本质控 = " |血液|" & IIf(blnQc, "1", "0")
                ElseIf varLine(iRow) Like "D2U*" Then
                        
                    strItem = strItem & "|WBC|" & Val(Mid(varLine(iRow), 11, 6)) / (10 ^ 3) 'WBC,11,6,3
                    
                    strItem = strItem & "|RBC|" & Val(Mid(varLine(iRow), 17, 5)) / (10 ^ 3) 'RBC,17,5,3
    
                    strItem = strItem & "|HGB|" & Val(Mid(varLine(iRow), 22, 5)) / (10 ^ 1) 'pitem3=HGB,22,5,2
    
                    strItem = strItem & "|HCT|" & Val(Mid(varLine(iRow), 27, 5)) / (10 ^ 2) 'pitem4=HCT,27,5,2
    
                    strItem = strItem & "|MCV|" & Val(Mid(varLine(iRow), 32, 5)) / (10 ^ 2) 'pitem5=MCV,32,5,2
    
                    strItem = strItem & "|MCH|" & Val(Mid(varLine(iRow), 37, 5)) / (10 ^ 2) 'pitem6=MCH,37,5,2
    
                    strItem = strItem & "|MCHC|" & Val(Mid(varLine(iRow), 42, 5)) / (10 ^ 1) 'pitem7=MCHC,42,5,2
    
                    strItem = strItem & "|PLT|" & Val(Mid(varLine(iRow), 47, 5)) / (10 ^ 1) 'pitem8=PLT,47,5,1
    
                    strItem = strItem & "|LYMPHP|" & Val(Mid(varLine(iRow), 52, 6)) / (10 ^ 3) 'pitem9=LYMPH%,52,6,3
    
                    strItem = strItem & "|MONOP|" & Val(Mid(varLine(iRow), 57, 6)) / (10 ^ 3) 'pitem10=MONO%,57,6,3
    
                    strItem = strItem & "|NEUTP|" & Val(Mid(varLine(iRow), 62, 6)) / (10 ^ 3) 'pitem11=NEUT%,62,6,3
    
                    strItem = strItem & "|EOP|" & Val(Mid(varLine(iRow), 67, 6)) / (10 ^ 3) 'pitem12=EO%,67,6,3
    
                    strItem = strItem & "|BASOP|" & Val(Mid(varLine(iRow), 72, 6)) / (10 ^ 3) 'pitem13=BASO%,72,6,3
    
                    strItem = strItem & "|LYMPHA|" & Val(Mid(varLine(iRow), 77, 6)) / (10 ^ 3) 'pitem14=LYMPH#,77,6,3
    
                    strItem = strItem & "|MONOA|" & Val(Mid(varLine(iRow), 83, 6)) / (10 ^ 3) 'pitem15=MONO#,83,6,3
    
                    strItem = strItem & "|NEUTA|" & Val(Mid(varLine(iRow), 89, 6)) / (10 ^ 3) 'pitem16=NEUT#,89,6,3
    
                    strItem = strItem & "|EOA|" & Val(Mid(varLine(iRow), 95, 6)) / (10 ^ 3) 'pitem17=EO#,95,6,3
    
                    strItem = strItem & "|BASOA|" & Val(Mid(varLine(iRow), 101, 6)) / (10 ^ 3) 'pitem18=BASO#,101,6,3
    
                    strItem = strItem & "|RDWCV|" & Val(Mid(varLine(iRow), 107, 5)) / (10 ^ 2) 'pitem19=RDW-CV,107,5,2
    
                    strItem = strItem & "|RDWSD|" & Val(Mid(varLine(iRow), 112, 5)) / (10 ^ 2) 'pitem20=RDW-SD,112,5,2

                        
                    If Val(Mid(varLine(iRow), 117, 5)) = 0 Then
                        strItem = strItem & "|PDW|----"
                    Else
                        strItem = strItem & "|PDW|" & Val(Mid(varLine(iRow), 117, 5)) / (10 ^ 2) 'pitem21=PDW,117,5,2
                    End If

                    If Val(Mid(varLine(iRow), 122, 5)) = 0 Then
                        strItem = strItem & "|MPV|----"
                    Else
                        strItem = strItem & "|MPV|" & Val(Mid(varLine(iRow), 122, 5)) / (10 ^ 2) 'pitem22=MPV,122,5,2
                    End If
                    If Val(Mid(varLine(iRow), 127, 5)) = 0 Then
                        strItem = strItem & "|PLCR|----"
                    Else
                        strItem = strItem & "|PLCR|" & Val(Mid(varLine(iRow), 127, 5)) / (10 ^ 2) 'pitem23=P-LCR,127,5,2
                    End If
                    strItem = strItem & "|RETP|" & Val(Mid(varLine(iRow), 132, 5)) / (10 ^ 3) 'pitem24=RET%,132,5,3

                    strItem = strItem & "|RET|" & Val(Mid(varLine(iRow), 137, 5)) / (10 ^ 5) 'pitem25=RET#,137,5,5

                    strItem = strItem & "|IRF|" & Val(Mid(varLine(iRow), 142, 5)) / (10 ^ 2) 'pitem26=IRF,142,5,2

                    strItem = strItem & "|LFR|" & Val(Mid(varLine(iRow), 147, 5)) / (10 ^ 2) 'pitem27=LFR,147,5,2

                    strItem = strItem & "|MFR|" & Val(Mid(varLine(iRow), 152, 5)) / (10 ^ 2) 'pitem28=MFR,152,5,2

                    strItem = strItem & "|HFR|" & Val(Mid(varLine(iRow), 157, 5)) / (10 ^ 2) 'pitem29=HFR,157,5,2

                    If Val(Mid(varLine(iRow), 162, 5)) = 0 Then
                        strItem = strItem & "|PCT|----"
                    Else
                        strItem = strItem & "|PCT|" & Val(Mid(varLine(iRow), 162, 5)) / (10 ^ 3) 'pitem30=PCT,162,5,3
                    End If
                    strItem = strItem & "|NRBCP|" & Val(Mid(varLine(iRow), 167, 6)) / (10 ^ 1) 'pitem31=NRBC%,167,6,1

                    strItem = strItem & "|NRBCA|" & Val(Mid(varLine(iRow), 173, 6)) / (10 ^ 2) 'pitem32=NRBC#,173,6,2
                    
                    
                     '图形数据
                    str_Path = App.Path & "\Gram"
                    If Dir(str_Path, vbDirectory) = "" Then
                        MkDir str_Path
                    End If
                    If InStr(strLine, "D3U") > 0 Then
                        str_imgName = "RBC DISCRI"
                        str_File = Replace(Mid(str_日期, 1, 10), "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                        str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, strLine, InStr(strLine, "D3U"))
                    End If
                    If InStr(strLine, "D4U") > 0 Then
                        str_imgName = "PLT DISCRI"
                        str_File = Replace(Mid(str_日期, 1, 10), "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                        str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, strLine, InStr(strLine, "D4U"))
                    End If
                    If InStr(strLine, "D5U") > 0 Then
                        str_imgName = "WBC DISCRI"
                        str_File = Replace(Mid(str_日期, 1, 10), "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                        str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, strLine, InStr(strLine, "D5U"))
                    End If
                    If InStr(strLine, "D1G") > 0 Then
                        str_imgName = "DIFF SCAT"
                        str_File = Replace(Mid(str_日期, 1, 10), "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                        str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, strLine, InStr(strLine, "D1G"))
                    End If
                    If InStr(strLine, "D2G") > 0 Then
                        str_imgName = "BASO SCAT"
                        str_File = Replace(Mid(str_日期, 1, 10), "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                        str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, strLine, InStr(strLine, "D2G"))
                    End If
                    If InStr(strLine, "D3G") > 0 Then
                        str_imgName = "SIMI SCAT"
                        str_File = Replace(Mid(str_日期, 1, 10), "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                        str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, strLine, InStr(strLine, "D3G"))
                    End If
                    If InStr(strLine, "D4G") > 0 Then
                        If bln_Rec Then
                            str_imgName = "RET SCAT"
                            str_File = Replace(Mid(str_日期, 1, 10), "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                            str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, strLine, InStr(strLine, "D4G"))
                        End If
                    End If
                    If InStr(strLine, "D5G") > 0 Then
                        str_imgName = "SPLT-O"
                        str_File = Replace(Mid(str_日期, 1, 10), "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                        str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, strLine, InStr(strLine, "D5G"))
                    End If
                    If InStr(strLine, "D6G") > 0 Then
                        str_imgName = "SRET-E"
                        str_File = Replace(Mid(Mid(str_日期, 1, 10), 1, 10), "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                        str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, strLine, InStr(strLine, "D6G"))
                    End If
                    If InStr(strLine, "D7G") > 0 Then
                        str_imgName = "SNRBC"
                        str_File = Replace(Mid(str_日期, 1, 10), "-", "_") & "_" & Val(str_标本号) & "_" & str_imgName & ".gif"
                        str_RetuImage = str_RetuImage & Save_image(str_imgName, str_Path & "\" & str_File, strLine, InStr(strLine, "D7G"))
                    End If
        
                End If
            Next
            
            If strItem <> "" And Trim(str_标本号) <> "" Then
                strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "|" & str检验人标本质控 & strItem
            End If
            If str_RetuImage <> "" Then strResult = strResult & vbCrLf & Mid(str_RetuImage, 2)
            If Not strReserved Like "***" Then
                bln_Rec = True
            End If
        Loop
    Else
        '收到的信息不全,返回,等下次再解析
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    strCmd = Chr(6)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '           元素之间以|分隔
    '           第0个元素：检验时间
    '           第1个元素：样本序号
    '           第2个元素：检验人
    '           第3个元素：标本
    '           第4个元素：是否质控品
    '           第5～9元素：系统保留
    '           从第10个元素开始为需要的检验项目。
    '   iSendStep  : 步骤
    '   blnSuccess : 是否传输失败
    '   strResponse: 仪器的应答信息
    '   blnUndo    : 是否取消.
    '
    '   返回值：发到仪器的字串或指令。
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function


Private Function Save_image(ByVal str_img As String, ByVal str_FileName As String, ByVal str_Data As String, ByVal lng_Start As Long) As String
    Dim lng_Lens As Long, str_ImgData As String, T_ImgInfo As typHistGramInfo
    
    If lng_Start <= 0 Then Exit Function
    
    If str_img = "RBC DISCRI" Or str_img = "PLT DISCRI" Or str_img = "WBC DISCRI" Then
        lng_Lens = Val(Mid(str_Data, lng_Start + 22, 6)) - 12
        If lng_Lens <= 0 Then Exit Function
        str_ImgData = Mid(str_Data, lng_Start + 41, lng_Lens)
        T_ImgInfo.lnglower = Mid(str_Data, lng_Start + 29, 4)
        T_ImgInfo.lngupper = Mid(str_Data, lng_Start + 33, 4)
        T_ImgInfo.lngmaxx = Mid(str_Data, lng_Start + 16, 3)
        T_ImgInfo.lngmaxy = Mid(str_Data, lng_Start + 19, 3)
        T_ImgInfo.lngresver1 = 0: T_ImgInfo.lngresver2 = 0
        T_ImgInfo.lngstoppos = IIf(str_img = "PLT DISCRI", 40, 46)
        T_ImgInfo.lngstoppos = IIf(str_img = "WBC DISCRI", 40, 46)
        
        If uf_xehist(str_ImgData, lng_Lens, T_ImgInfo, "c:\temp1.bmp", str_FileName) = 1 Then
            Save_image = "^" & str_img & ";105;1;" & str_FileName
        ElseIf uf_ufHist(str_ImgData, lng_Lens, "c:temp1.bmp", str_FileName) = 1 Then
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
