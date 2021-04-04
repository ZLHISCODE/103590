VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BC5500_bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:BC-5500
'厂商:深圳迈瑞
'作者:陈东
'用户:商丘第一人民医院

Private mstrIn As String
Const strRetuItems = ",WBC,Neu#,Lymph#,Mon#,Eos#,Bas#,Neu%,Lymph%,Mon%,Eos%,Bas%,RBC,HGB,HCT,MCV,MCH,MCHC,RDW-CV,RDW-SD,PLT,MPV,PDW,PCT,"  '要解析的项目

Dim StrInput As String

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
    
    Dim strData As String, lngS As Long, lngE As Long
    Dim str标本号 As String, str日期 As String, str项目 As String
    Dim str质控 As String, strItemName As String

    Dim strSegment As String '存一个 Segment的内容
    Dim lngSD_FD As Long '存SD和FD的分隔符位置
    Dim lngFD_V As Long '存 FD和V的分隔符位置
    Dim strFD_Name As String '存FD的说明
    
    Dim strSegName As String '放segment的说明
    
    Dim lngImgDataLen As Long, lngImgStart As Long, lngI As Long
    Dim strImgRBC As String, strImgPLT As String, strImgDiff As String, strImgBASO As String
    Dim strImgData As String, strTmp标本号 As String, i As Integer
    If strSource = ",02" Then
        strCmd = ",06" ' Chr(&H6)
    Else
        StrInput = StrInput & strSource
    End If
    
    If StrInput Like "*,05,##,##,##,03,*,08,04,0A*" Then
        '---------
        str标本号 = "": str日期 = "": str项目 = "": strResult = "": strData = ""
        strImgRBC = "": strImgPLT = "": strImgDiff = "": strImgBASO = ""
        lngImgDataLen = 0: lngImgStart = 0: lngI = 0
        '-----
        lngS = Instr_Bin(StrInput, Chr(&H5)) '--开始符 MS
        If Mid_bin(StrInput, lngS + 4, 1) <> Chr(&H3) Then
            '不是开始位置，退出
            StrInput = Mid_bin(StrInput, lngS + 1, , False)
            Exit Sub
        End If
        
        lngE = Instr_Bin(StrInput, Chr(&H8) & Chr(&H4) & Chr(&HA))        '--结束符         chr(10)
        
        If lngE <= lngS Then Exit Sub
        
        strData = Mid_bin(StrInput, lngS, lngE - lngS + 3, False)
        StrInput = Mid_bin(StrInput, lngE + 3, , False)
        'If Instr_Bin(strInput, Chr(&H5)) <= 0 Then strReserved = ""
        
        lngS = 1
        
        str质控 = UCase(Mid_bin(strData, lngS + 1, 3))
        If str质控 = "CTR" Then
            str质控 = "0"
            strData = Mid_bin(strData, lngS + 5, , False)
        ElseIf str质控 = "QCR" Then
            str质控 = "1"
            strData = Mid_bin(strData, lngS + 5, , False)
        ElseIf str质控 = "QCC" Then
            '这个数据不解析
            Exit Sub
        End If
        
        Do While Len(strData) > 0
            '读一段数据
            lngE = Instr_Bin(strData, Chr(&H4))
            If lngE > 0 Then
                strSegment = Mid_bin(strData, 1, lngE - 1, False)
            Else
                '最后一段
                strSegment = strData
            End If
            
            lngSD_FD = Instr_Bin(strSegment, Chr(&HC))
            
            If lngSD_FD > 0 Then
                strSegName = Mid_bin(strSegment, 1, lngSD_FD - 1)
                strSegment = Mid_bin(strSegment, lngSD_FD + 1, , False)
                If InStr(strRetuItems, "," & strSegName & ",") > 0 Then     '检验项目
                    
                    If lngE > 0 Then
                        strData = Mid_bin(strData, lngE + 1, , False)
                    Else
                        strData = ""
                    End If    '文本数据，可以直接用strsegment的数据
                    
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            '格式不对,不解析
                            strSegment = ""
                        End If
                        
                        If strFD_Name = "Val" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngE > 0 Then
                                str项目 = str项目 & "|" & strSegName & "|" & Mid_bin(strSegment, 1, lngE - 1)
                            Else
                                str项目 = str项目 & "|" & strSegName & "|" & Mid_bin(strSegment, 1)
                            End If
                            strSegment = ""
                            
                        End If
                    Loop
                ElseIf strSegName = "PatInfo" Then                          '标本号
                    If lngE > 0 Then
                        strData = Mid_bin(strData, lngE + 1, , False)
                    Else
                        strData = ""
                    End If    '文本数据，可以直接用strsegment的数据
                    
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            '格式不对,不解析
                            strSegment = ""
                        End If

                        
                        If strFD_Name = "SampleID" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngE > 0 Then
                                str标本号 = Mid_bin(strSegment, 1, lngE - 1)
                            Else
                                str标本号 = Val(Mid_bin(strSegment, 1))
                            End If
                            If Val(str标本号) = 0 And Len(str标本号) > 0 Then
                                
                                strTmp标本号 = str标本号
                                str标本号 = ""
                                For i = 1 To Len(strTmp标本号)
                                    If IsNumeric(Mid(strTmp标本号, i, 1)) Then
                                        str标本号 = str标本号 & Mid(strTmp标本号, i, 1)
                                    End If
                                Next
                                str标本号 = 1000 + Val(str标本号)
                                str质控 = "1"
                            End If
                            strSegment = ""
                        End If
                    Loop
                ElseIf strSegName = "PatTime" Then                         '日期
                    If lngE > 0 Then
                        strData = Mid_bin(strData, lngE + 1, , False)
                    Else
                        strData = ""
                    End If    '文本数据，可以直接用strsegment的数据
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            '格式不对,不解析
                            strSegment = ""
                        End If

                        
                        If strFD_Name = "TestTime" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngE > 0 Then
                                str日期 = Mid_bin(strSegment, 1, lngE - 1)
                            Else
                                str日期 = Mid_bin(strSegment, 1)
                            End If
                            If Not IsDate(str日期) Then str日期 = Format(Now, "yyyy-MM-dd")
                            strSegment = ""
                        End If
                    Loop
                    
                ElseIf strSegName = "RBCHisto" Then                     'rbc直方图
                    '二进制数据
                    
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            strSegment = ""
                        End If
                        
                        '----- 数据长度
                        If strFD_Name = "DataLen" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngFD_V > 0 Then
                               lngImgDataLen = Val(Mid_bin(strSegment, 1, lngE - 1))
                            End If
                            
                            strSegment = ""
                        End If
                        '---- 取图形数据
                        lngImgStart = Instr_Bin(strData, "RHistoData")
                        If lngImgStart > 0 Then
                            lngImgStart = lngImgStart + 11
                            strImgRBC = Graph_Text(Mid_bin(strData, lngImgStart, lngImgDataLen, False))
                            strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                            strSegment = ""
                        End If
                    Loop
                ElseIf strSegName = "PLTHisto" Then                     'plt直方图
                    '二进制数据
                    
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            strSegment = ""
                        End If
                        
                        '----- 数据长度
                        If strFD_Name = "DataLen" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngFD_V > 0 Then
                               lngImgDataLen = Val(Mid_bin(strSegment, 1, lngE - 1))
                            End If
                            
                            strSegment = ""
                        End If
                        '---- 取图形数据
                        lngImgStart = Instr_Bin(strData, "PHistoData")
                        If lngImgStart > 0 Then
                            lngImgStart = lngImgStart + 11
                            strImgPLT = Graph_Text(Mid_bin(strData, lngImgStart, lngImgDataLen, False))
                            strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                            strSegment = ""
                        End If
                    Loop
'                ElseIf strSegName = "SepLine" Then                      '直方图分类线
                ElseIf strSegName = "BASOType" Or strSegName = "DIFFType" Then
                    '二进制数据
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            strSegment = ""
                        End If
                        
                        '----- 数据长度
                        If strFD_Name = "DataLen" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngFD_V > 0 Then
                               lngImgDataLen = Val(Mid_bin(strSegment, 1, lngE - 1))
                            End If
                            
                            strSegment = ""
                        End If
                        
                        If strSegName = "BASOType" Then
                            lngImgStart = Instr_Bin(strData, "BASOTypeData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 13
                                Dim strBASOType As String
                                strBASOType = Mid_bin(strData, lngImgStart, lngImgDataLen, False)
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        Else
                            lngImgStart = Instr_Bin(strData, "DIFFTypeData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 13
                                Dim strDiffType As String
                                strDiffType = Mid_bin(strData, lngImgStart, lngImgDataLen, False)
                                
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        
                        
                        End If
                    Loop
                ElseIf strSegName = "BASO" Or strSegName = "DIFF" Then  '散点图
                    '二进制数据
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            strSegment = ""
                        End If
                        
                        '----- 数据长度
                        If strFD_Name = "DataLen" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngFD_V > 0 Then
                               lngImgDataLen = Val(Mid_bin(strSegment, 1, lngE - 1))
                            End If
                            
                            strSegment = ""
                        End If
                        '---- 取图形数据
                        If strSegName = "BASO" Then
                            lngImgStart = Instr_Bin(strData, "BASOData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 9
                                strImgBASO = DrawBmpToFile(Mid_bin(strData, lngImgStart, lngImgDataLen, False), "BASO", str标本号, strBASOType)
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        Else
                            lngImgStart = Instr_Bin(strData, "DIFFData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 9
                                strImgDiff = DrawBmpToFile(Mid_bin(strData, lngImgStart, lngImgDataLen, False), "DIFF", str标本号, strDiffType)
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        End If
                    Loop
                ElseIf strSegName = "BASOBMP" Or strSegName = "DIFFBMP" Then
                    '二进制数据
                    Do While Len(strSegment) > 0
                        lngFD_V = Instr_Bin(strSegment, Chr(&H16))
                        If lngFD_V > 0 Then
                            strFD_Name = Mid_bin(strSegment, 1, lngFD_V - 1)
                            strSegment = Mid_bin(strSegment, lngFD_V + 1, , False)
                        Else
                            strSegment = ""
                        End If
                        
                        '----- 数据长度
                        If strFD_Name = "DataLen" Then
                            lngE = Instr_Bin(strSegment, Chr(&H8))
                            If lngFD_V > 0 Then
                               lngImgDataLen = Val(Mid_bin(strSegment, 1, lngE - 1))
                            End If
                            
                            strSegment = ""
                        End If
                        '---- 取图形数据
                        If strSegName = "BASOBMP" Then
                            lngImgStart = Instr_Bin(strData, "BASOBmpData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 12
                                strImgBASO = SaveBmpToFile(Mid_bin(strData, lngImgStart, lngImgDataLen, False), "BASO", str标本号)
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        Else
                            lngImgStart = Instr_Bin(strData, "DIFFBmpData")
                            If lngImgStart > 0 Then
                                lngImgStart = lngImgStart + 12
                                strImgBASO = SaveBmpToFile(Mid_bin(strData, lngImgStart, lngImgDataLen, False), "DIFF", str标本号)
                                strData = Mid_bin(strData, lngImgStart + lngImgDataLen, , False)
                                strSegment = ""
                            End If
                        End If
                    Loop
                Else
                    If lngE > 0 Then
                        strData = Mid_bin(strData, lngE + 1, , False)
                    Else
                        strData = ""
                    End If
                End If
                
            Else
                If lngE > 0 Then
                    strData = Mid_bin(strData, lngE + 1, , False)
                Else
                    strData = ""
                End If
            End If
        Loop
        
        If Val(str标本号) > 0 And IsDate(str日期) And str项目 <> "" Then
            strResult = str日期 & "|" & Val(str标本号) & "| |手指血|" & str质控 & str项目
            strImgData = ""
            If strImgDiff <> "" Then strImgData = strImgData & "^DIFF;107;1;" & strImgDiff
            If strImgBASO <> "" Then strImgData = strImgData & "^BASO;107;1;" & strImgBASO
            If strImgRBC <> "" Then strImgData = strImgData & "^RBC;0;260,300,0-0,50-50,100-100,150-150,200-200,250-250,300-fL" & strImgRBC
            If strImgPLT <> "" Then strImgData = strImgData & "^PLT;0;260,150,0-0,21.333-5,42.66-10,63.99-15,85.32-20,106.65-25,127.98-fL" & strImgPLT
            If strImgData <> "" Then strResult = strResult & vbCrLf & Mid(strImgData, 2)
        End If
        
    End If
    
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function

Private Function DrawBmpToFile(ByVal str_bin As String, ByVal strType As String, ByVal str标本号 As String, ByVal strDataType As String) As String
    Dim strBinData As String
    Dim lngCount As Long
    Dim strfilename As String
    
    Dim frmPic As New frmGraph
    strBinData = str_bin
    lngCount = Len_Bin(strBinData)
    If lngCount <= 0 Then Exit Function
    
    strfilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    If frmPic.Draw_Bc5500(strBinData, strfilename, strDataType) Then
        DrawBmpToFile = strfilename
    End If
End Function


Private Function SaveBmpToFile(ByVal str_bin As String, ByVal strType As String, ByVal str标本号 As String) As String
    '保存二进制数据为BMP图形文件
    Dim strfilename As String, lngFileNum As Long, lngloop As Long
    Dim strByte As String, ByteData() As Byte
    Dim strBinData As String
    Dim lngCount As Long
    
    strBinData = str_bin
    lngCount = Len_Bin(strBinData)
    If lngCount <= 0 Then Exit Function
    ReDim ByteData(lngCount)
    
    For lngloop = 0 To lngCount - 1
        strByte = Left(strBinData, 3)
        strBinData = Mid(strBinData, 4)
        ByteData(lngloop) = CByte("&H" & Mid(strByte, 2))
    Next
    
    strfilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strfilename For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    SaveBmpToFile = strfilename
End Function

Private Function Graph_Text(ByVal str_bin As String) As String
    '生成图形文本数据
    Dim lng_len As Long, lng_i As Long, str_bit As String
    Dim str_Return As String
    lng_len = Len_Bin(str_bin)

    If lng_len < 1 Then
        Graph_Text = ""
        Exit Function
    End If

    For lng_i = 1 To lng_len
        str_bit = Mid(Mid_bin(str_bin, lng_i, 1, False), 2)
    
        If IsNull(str_bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & ";" & Format(CByte("&H" & str_bit), "000")
        End If
    Next
    Graph_Text = str_Return
End Function
