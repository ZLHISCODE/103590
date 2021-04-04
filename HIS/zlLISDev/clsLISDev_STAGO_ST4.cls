VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_STAGO_ST4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:STAGO -ST4凝血仪
'厂商:法国 STAGO
'作者: 陈东
'用户:重庆江陵医院


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
    '   每组检验结果以||分隔,元素之间以 分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    
    Dim iStrPoint As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim varIn As Variant '存放输入数据
    Dim strTime As String
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strLine As String, intCount As Integer, i As Integer
    Dim str标本号 As String, str日期 As String, str检验人标本质控 As String, str指标 As String
    Dim strIn As String
    On Error GoTo errHandle
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    If strSource = "" Then Exit Sub
    
    If InStr(strSource, Chr(3)) = 0 Or InStr(strSource, Chr(2)) = 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    str指标 = Trim(Mid(strSource, 11, 2))
    If str指标 = "1" Then str指标 = "PT"
    If str指标 = "2" Then str指标 = "APTT"
    If str指标 = "" Then Exit Sub
    str检验人标本质控 = " |血液|0"
    
    'str指标 = str指标 & " " & Trim(Mid(strSource, 13, 5))
    WriteLog "STAGO ST4", "准备解码", strSource
    
    strSource = Replace(strSource, "}", "")
    If InStr(strSource, Chr(3)) > 0 And InStr(strSource, Chr(2)) > 0 Then
        '一次性传完了一次结果
        strIn = strSource
        For iStrPoint = 1 To Len(strSource)
            
            If Mid(strSource, iStrPoint, 1) = Chr(2) Then
                strIn = Mid(strSource, iStrPoint, 1)
            Else
                strIn = strIn & Mid(strSource, iStrPoint, 1)
            End If
            
            If Mid(strSource, iStrPoint, 1) = Chr(3) Then
                strReserved = Mid(strSource, iStrPoint + 1)
                Exit For
            End If
        Next
    Else
        '没传完,不解析,直接返回
        strReserved = strSource
        Exit Sub
    End If
    
    If InStr(strIn, Chr(2)) > 0 And InStr(strIn, Chr(3)) > 0 Then
        iStrPoint = 0
        varIn = Split(strIn, Chr(3))
        For iStrPoint = LBound(varIn) To UBound(varIn) - 1
            intCount = Val(Trim(Mid(varIn(iStrPoint), 26, 2)))
            If intCount <= 0 Then Exit Sub
            str日期 = Format(Now, "yyyy-MM-dd HH:mm:ss")
            str标本号 = Trim(Mid(varIn(iStrPoint), 9, 2))
            If Len(varIn(iStrPoint)) < (intCount * 37 + 27) Then Exit Sub
            
            For i = 1 To intCount
                strLine = Mid(varIn(iStrPoint), 28 + (i - 1) * 37, 37)
                strResult = strResult & "||" & str日期 & "|" & str标本号 & Trim(Mid(strLine, 1, 12)) & "|" & str检验人标本质控 & "|" & _
                              str指标 & "|" & IIf(Trim(Mid(strLine, 14, 8)) = "", "－－", Trim(Mid(strLine, 14, 8)))
                If str指标 <> "APTT" Then
                    strResult = strResult & "|" & Trim(Mid(varIn(iStrPoint), 13, 5)) & "|" & IIf(Trim(Mid(strLine, 22, 8)) = "", "－－", Trim(Mid(strLine, 22, 8))) & "|" & _
                                  "INR|" & IIf(Trim(Mid(strLine, 30, 8)) = "", "－－", Trim(Mid(strLine, 30, 8)))
                End If
            Next
        Next
        If strResult <> "" Then strResult = Mid(strResult, 3)
    End If
    Exit Sub
    
errHandle:
    WriteLog "STAGO ST4", "解码错误", Err.Description
    'Resume
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


