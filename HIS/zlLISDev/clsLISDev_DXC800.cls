VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_DXC800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'仪器:临床生化分析系 DXC800
'厂商:美国贝克曼库尔特 Beckman coulter
'作者:陈东
'用户:新疆塔里木医院

Private mblnStart As Boolean '开始
Private mblnEnd As Boolean   '结束
Private mstrData As String '他待解析串

Private mstrLastResult As String '上次返回的串
Private mIntLOG As Integer '重复解析相同串的次数

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
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    
    '----- 定义变量
    Dim varIn As Variant
    Dim varLin As Variant
    Dim strLine As String, i As Integer
    Dim iStartRow As Integer, iEndRow As Integer, str日期 As String, str值 As String
    Dim StrInput As String, str指标 As String
    Dim curTP As Currency, curALB As Currency, curTBIL As Currency, curDBIL As Currency
    
    On Error GoTo errHandle
    
    StrInput = Replace(strSource, vbCrLf, "☆" & vbCrLf)
    '----  初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    iStartRow = -1: iEndRow = -1
    
    If StrInput Like "*[[]*[]]*" Or StrInput Like "[[]*[]]*" Then
        varIn = Split(StrInput, vbCrLf)
        
        For i = LBound(varIn) To UBound(varIn)
            strLine = varIn(i)
            If strLine Like "*[[]*,802,03,*[]]*" Or strLine Like "[[]*,802,03,*[]]*" Then
                
                strLine = Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1)
                varLin = Split(strLine, ",")
                str日期 = Mid(varLin(3), 5, 4) & "-" & Mid(varLin(3), 3, 2) & "-" & Mid(varLin(3), 1, 2)
                
                str值 = IIf(Trim(varLin(15)) Like "*[#]*", " ", Trim(varLin(15)))
                str指标 = Trim(varLin(10))
                If str指标 = "" Then str指标 = " "
                If str指标 = "07D" Then curTP = Val(str值)
                If str指标 = "08D" Then curALB = Val(str值)
                If str指标 = "11A" Then curTBIL = Val(str值)
                If str指标 = "11B" Then curDBIL = Val(str值)
                strResult = strResult & str日期 & "|" & Val(varLin(9)) & "| |血液|0|" & str指标 & "|" & IIf(str值 = "", " ", str值) & "||"
                
                If curTP <> 0 And curALB <> 0 Then
                    strResult = strResult & str日期 & "|" & Val(varLin(9)) & "| |血液|0|GLO|" & Format((curTP - curALB), "0.00") & "|A/G|" & Format(curALB / (curTP - curALB), "0.00") & "||"
                    curTP = 0
                    curALB = 0
                End If
                
                If curTBIL <> 0 And curDBIL <> 0 Then
                    strResult = strResult & str日期 & "|" & Val(varLin(9)) & "| |血液|0|IBIL|" & Format((curTBIL - curDBIL), "0.00") & "||"
                    curTBIL = 0
                    curDBIL = 0
                End If
                
            ElseIf strLine Like "[[]*,802,01,*[]]*" Or strLine Like "*[[]*,802,01,*[]]*" Then
                '一个标本开始
                iStartRow = i
            ElseIf strLine Like "*[[]*,802,05,*[]]*" Or strLine Like "[[]*,802,05,*[]]*" Then
                '记录已解析行
                iEndRow = i
            Else
                If Not strLine Like "*[[]*[]]*" Or strLine Like "[[]*[]]*" Then
                    '返回不完整的数据行
                    If strLine <> "" Then
                        strReserved = strReserved & Replace(strLine, "☆", vbCrLf)
                    End If
                End If
            End If
        Next
        
    Else
        '数据不完整,返回传入值
        strReserved = strSource
    End If
    
    If Len(strResult) > 0 Then
        If Mid(strResult, 1, 2) = "||" Then strResult = Mid(strResult, 3)
        If Right(strResult, 2) = "||" Then strResult = Mid(strResult, 1, Len(strResult) - 2)
    End If
    
    If strResult <> "" Then mIntLOG = 0 '有解析结果,就把重复解析相同串的次数清零
    
    If mstrLastResult <> "" And strReserved <> "" Then
        If mstrLastResult = strReserved And mstrLastResult = strSource And strResult = "" Then
            mIntLOG = mIntLOG + 1
        End If
    End If
    mstrLastResult = strReserved
    
    '--- 重复解析了100次相同的串,没且没得结果,放弃解析
'    If mIntLOG > 10000 Then
'        strReserved = ""
'        mIntLOG = 0
'    End If
    
    If mIntLOG < 5 Then
        '不重复写日志
        Call WriteLog("DXC800 analyse", "strsource:" & strSource & vbNewLine & "strReserved:" & strReserved, "strResult:" & strResult)
    End If
    Exit Sub
errHandle:
    Call WriteLog("DXC800 analyse", "解码错误：" & Err.Description, "")
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


