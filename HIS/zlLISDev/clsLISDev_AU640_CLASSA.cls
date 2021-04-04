VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AU640_CLASSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AU640
'厂商:奥林巴斯
'作者:陈东
'用户:南京江北人民医院

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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, lngLength As Long, iRows As Integer
    Dim blnEnd As Boolean '是否已解析完
    Dim intloop As Integer
    
    ''
    Dim dblTP As Double, dblALB As Double, dblTBIL As Double, dblDBIL As Double
    Dim intE As Integer, intP As Integer, intQ As Integer
    
    intE = Val(ReadIni("AU640_A", "E", App.Path & "\AU640_A.ini"))
    If intE = 0 Then intE = 8000:   Call WriteIni("AU640_A", "E", 8000, App.Path & "\AU640_A.ini")
    
    intQ = Val(ReadIni("AU640_A", "Q", App.Path & "\AU640_A.ini"))
    If intQ = 0 Then intE = 6000:   Call WriteIni("AU640_A", "Q", 6000, App.Path & "\AU640_A.ini")
    
    intP = Val(ReadIni("AU640_A", "P", App.Path & "\AU640_A.ini"))
    If intP = 0 Then intE = 3000:   Call WriteIni("AU640_A", "P", 3000, App.Path & "\AU640_A.ini")
    
    
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "***" Or strSource Like "***" Then
        blnEnd = False
        strReserved = strSource
        strReserved = Replace(strReserved, "", "")
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            
            lngLength = InStr(strReserved, "") - InStr(strReserved, "") + 1
            
            strLine = Mid(strReserved, InStr(strReserved, ""), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            strLine = Replace(strLine, "", "")
            If strLine Like "RB*" Or strLine Like "RE*" Then  '请求开始和请求结束，不处理
                ' #
                strCmd = ""
            ElseIf strLine Like "R*" Then '条码扫描请求
                '# 读取日期，条码返回
                'strCmd="1|" & 日期^^条码
            ElseIf strLine Like "D *" Or strLine Like "DQ*" Then
                strCmd = Chr(&H6)
                str日期 = Format(Now, "yyyy-MM-dd HH:mm:ss")
                strTmp = Mid(strLine, 13, 1)
                If strTmp = "E" Then  '手册上是紧急抽样
                    
                    str标本号 = intE + Val(Replace(Mid(strLine, 14, 3), " ", ""))
                    str检验人标本质控 = " |血清|0"
                ElseIf strTmp = "Q" Then  '急诊 '手册上指质控抽样
                    str标本号 = intQ + Val(Replace(Mid(strLine, 14, 3), " ", "")) + Val(Replace(Mid(strLine, 39, 2), " ", "")) * 10 & "^1"
                    str检验人标本质控 = " |血清|1"
                ElseIf strTmp = "P" Then     '手册上指统计抽样
                    str标本号 = intP + Val(Replace(Mid(strLine, 14, 3), " ", "")) & "^1"
                    str检验人标本质控 = " |血清|0"
                'A-校准样品 R-试剂空白
                ElseIf IsNumeric(strTmp) Then
                    '常规样本
                    str标本号 = Val(Replace(Mid(strLine, 13, 4), " ", "")) & "^0"
                    str检验人标本质控 = " |血清|0"
                End If
                str项目 = Mid(strLine, 42, InStr(strLine, ""))
                If Len(str项目) >= 13 Then
                    '手工发送
                    If InStr(str项目, "r") Then
                        varItem = Split(str项目, "r")
                        For intloop = LBound(varItem) To UBound(varItem) - 1
                            strTmp = Trim(varItem(intloop))
                            
                            If Trim(Mid(strTmp, 1, 2)) = "04" Then
                                dblTP = Val(Trim(Mid(strTmp, 3)))
                            ElseIf Trim(Mid(strTmp, 1, 2)) = "05" Then
                                dblALB = Val(Trim(Mid(strTmp, 3)))
                            ElseIf Trim(Mid(strTmp, 1, 2)) = "06" Then
                                dblTBIL = Val(Trim(Mid(strTmp, 3)))
                            ElseIf Trim(Mid(strTmp, 1, 2)) = "07" Then
                                dblDBIL = Val(Trim(Mid(strTmp, 3)))
                            End If
                            
                            If intloop = 0 And Right(Trim(str检验人标本质控), 1) = 1 Then
                                strItem = strItem & "|" & Trim(Mid(strTmp, 3, 2)) & "|" & Val(Trim(Mid(strTmp, 5)))
                            Else
                                strItem = strItem & "|" & Trim(Mid(strTmp, 1, 2)) & "|" & Val(Trim(Mid(strTmp, 3)))
                            End If
                        
                        Next
                    Else
                        '自动发送
                        iRows = Len(str项目) / 13
                        
                        For iRow = 1 To iRows
                            If Right(Trim(str检验人标本质控), 1) = 1 Then
                                strItem = strItem & "|" & Mid(str项目, 3, 2) & "|" & Val(Trim(Mid(str项目, 5, 9)))
                            Else
                                strItem = strItem & "|" & Mid(str项目, 1, 2) & "|" & Val(Trim(Mid(str项目, 3, 9)))
                            End If
                            str项目 = Mid(str项目, 14)
                            
                            If Mid(str项目, 1, 2) = "04" Then
                                dblTP = Val(Trim(Mid(str项目, 3, 9)))
                            ElseIf Mid(str项目, 1, 2) = "05" Then
                                dblALB = Val(Trim(Mid(str项目, 3, 9)))
                            ElseIf Mid(str项目, 1, 2) = "06" Then
                                dblTBIL = Val(Trim(Mid(str项目, 3, 9)))
                            ElseIf Mid(str项目, 1, 2) = "07" Then
                                dblDBIL = Val(Trim(Mid(str项目, 3, 9)))
                            End If
                        Next
                    End If
                    
                    ' 球蛋白GLO=总蛋白TP-白蛋白ALB
                    ' 白球比A/G=白蛋白ALB/球蛋白GLO
                    If dblTP <> 0 And dblALB <> 0 Then
                        strItem = strItem & "|GLO|" & Format(dblTP - dblALB, "0.0")
                        If dblTP - dblALB <> 0 Then
                            strItem = strItem & "|AG|" & Format(dblALB / (dblTP - dblALB), "0.0")
                        End If
                    End If
                    '间接胆红素IDBIL=总胆红素TBIL-直接胆红素DBIL
                    If dblDBIL <> 0 And dblTBIL <> 0 Then
                        strItem = strItem & "|IDBIL|" & Format(dblTBIL - dblDBIL, "0.0")
                    End If
                End If
                '解析
            End If
            
            If strItem <> "" And Val(str标本号) <> 0 Then
                strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
            End If
            
            If Not strReserved Like "***" Then
                blnEnd = True
            ElseIf strReserved Like "*R*" Then '有请求信息，停止解析结果,返回。
                blnEnd = True
            End If
        Loop
    Else
        '收到的信息不全,返回,等下次再解析
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:解析需要向仪器发送的标本信息
    
    '解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5个元素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    Dim blnEnd As Boolean
    Dim strBuffer As String, strSampleID As String
    Dim i As Long, aItem() As String
    Dim strRackNO As String, iPos As Long
    
    blnSuccess = True: SendSample = ""
    If Val(strSampleID) <= 0 Then Exit Function '无条码退出
    
    If strResponse Like "*R*" Or Len(Trim(strResponse)) = 0 Then '允许发送
        If strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then
            '请求结束
            '生成设备串
            aItem = Split(strSample, "|")
            '标本或SampleID
            strBuffer = aItem(1)
            strSampleID = aItem(5)
            '读取反馈的架号和杯号
            strRackNO = ""
            If Len(Trim(strResponse)) > 0 Then
                strResponse = Replace(strResponse, "RB", "")
                iPos = InStr(strResponse, "R")
                strRackNO = Mid(strResponse, iPos + 5, 11) 'rack no 4位或5位 ,Cup 2位，类型,空格－血,U-尿,X-其他,sampleNO 4位，共11位或12位．
            
                SendSample = "S " & Right(Space(11) & strRackNO, 11) & Right(Space(20) & strSampleID, 20) & Space(4) & "E"

                For i = 10 To UBound(aItem)
                    SendSample = SendSample & aItem(i)
                Next
                SendSample = SendSample & "SE"
            End If
            iSendStep = 0
            
        Else
            SendSample = Chr(&H6)
            iSendStep = iSendStep + 1
        End If
    Else
        SendSample = Chr(&H6)
        iSendStep = 0
        blnSuccess = False '发送失败
    End If
End Function


