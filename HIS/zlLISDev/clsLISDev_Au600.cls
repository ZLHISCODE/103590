VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AU600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AU600
'厂商:奥林巴斯
'作者:陈东
'用户:重庆市三院

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
    
    Dim varline As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, lngLength As Long, iRows As Integer
    Dim blnEnd As Boolean '是否已解析完
    ''
    Dim dblTP As Double, dblALB As Double, dblTBIL As Double, dblDBIL As Double

    
    Dim intE As Integer, intP As Integer, intQ As Integer
    Dim intEmergency As Integer
    
    intE = Val(ReadIni("AU600", "E", App.Path & "\AU600.ini"))
    If intE = 0 Then intE = 8000:   Call WriteIni("AU600", "E", 8000, App.Path & "\AU600.ini")
    
    intQ = Val(ReadIni("AU600", "Q", App.Path & "\AU600.ini"))
    If intQ = 0 Then intE = 6000:   Call WriteIni("AU600", "Q", 6000, App.Path & "\AU600.ini")
    
    intP = Val(ReadIni("AU600", "P", App.Path & "\AU600.ini"))
    If intP = 0 Then intE = 3000:   Call WriteIni("AU600", "P", 3000, App.Path & "\AU600.ini")
    
    intEmergency = Val(ReadIni("AU600", "是否使用急诊", App.Path & "\AU600.ini"))
    If intEmergency = 0 Then Call WriteIni("AU600", "是否使用急诊", CStr(intEmergency), App.Path & "\AU600.ini")
    
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "***" Or strSource Like "***" Then
        blnEnd = False
        strReserved = strSource
        strReserved = Replace(strReserved, "", "")
        strReserved = Replace(strReserved, "", "")
        
        strReserved = Replace(strReserved, "", "")
        strReserved = Replace(strReserved, "", "")
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            
            
            lngLength = InStr(strReserved, "") - InStr(strReserved, "") + 1
            
            strLine = Mid(strReserved, InStr(strReserved, ""), lngLength)
            If InStr(2, strLine, "") > 0 Then
                strLine = Mid(strLine, InStr(2, strLine, ""))
            End If
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            
            If strLine Like "RB*" Or strLine Like "RE*" Then  '请求开始和请求结束，不处理
                ' #
                strCmd = ""
            ElseIf strLine Like "R*" Then '条码扫描请求
                '# 读取日期，条码返回
                'strCmd="1|" & 日期^^条码
            ElseIf strLine Like "D *" Or strLine Like "DQ*" Then
                strCmd = Chr(&H6)
                str日期 = Format(Now, "yyyy-MM-dd HH:mm:ss")
                strTmp = Mid(strLine, 7, 1)
                If strTmp = "E" Then '质控
                    str标本号 = intE + Val(Replace(Mid(strLine, 8, 3), " ", "")) & IIf(intEmergency = 0, "^1", "^0")
                    str检验人标本质控 = " |血清|0"
                ElseIf strTmp = "Q" Then  '急诊
                    str标本号 = intQ + Val(Mid(strLine, 8, 3)) * 10 + Val(Mid(strLine, 11, 5)) & IIf(intEmergency = 0, "^1", "^0")
                    str检验人标本质控 = " |血清|1"
                ElseIf strTmp = "P" Then
                    str标本号 = intP + Val(Replace(Mid(strLine, 8, 3), " ", "")) & IIf(intEmergency = 0, "^1", "^0")
                    str检验人标本质控 = " |血清|0"
                Else
                    str标本号 = Val(Replace(Mid(strLine, 7, 4), " ", "")) & "^0"
                    str检验人标本质控 = " |血清|0"
                End If
                str项目 = Mid(strLine, 16, InStr(strLine, ""))
                If Len(str项目) >= 13 Then
                    iRows = Len(str项目) / 13
                    For iRow = 1 To iRows
                        strItem = strItem & "|" & Mid(str项目, 1, 2) & "|" & Val(Trim(Mid(str项目, 3, 9)))
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
        SendSample = ""

End Function


