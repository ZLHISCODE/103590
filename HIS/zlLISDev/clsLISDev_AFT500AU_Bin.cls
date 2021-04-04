VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AFT500AU_Bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AFT-500AU 电解质分析仪
'厂商:梅州康立
'作者:李建飞
'用户:成都陈辉

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
        
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号^是否急诊^条码号  可以只有样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
   
    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim mstrCode As String
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim str_标本号 As String, str_日期 As String, str_检验项目 As String
    Dim varTmp() As String
    strResult = ""
    
    If Len(strSource) < 33 * 3 Then
        strReserved = strSource
        Exit Sub
    Else
        str_Code = Mid(strSource, 2)
    End If
    
    While str_Code <> ""
    
        varTmp = Split(str_Code, ",")
        str_日期 = Format(varTmp(4) & "-" & varTmp(3) & "-" & varTmp(2) & " " & varTmp(5) & ":" & varTmp(6) & ":" & varTmp(1), "yyyy-MM-dd hh:mm:ss")
        str_标本号 = CInt("&H" & varTmp(11))
        str_检验项目 = "|K|" & CalcResult(varTmp(13), varTmp(14)) & "|Na|" & CalcResult(varTmp(16), varTmp(17)) & "|Cl|" & _
                        CalcResult(varTmp(19), varTmp(20)) & "|nCa|" & CalcResult(varTmp(22), varTmp(23)) & "|tCa|" & CalcResult(varTmp(28), varTmp(29))
        
        If Val(str_标本号) <> 0 And IsDate(str_日期) And str_检验项目 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血液|0" & str_检验项目
        End If
        
        If Len(mstrCode) < 33 * 3 Then
            strReserved = mstrCode
            str_Code = ""
        Else
            str_Code = Mid(mstrCode, 2)
            mstrCode = ""
        End If
    Wend
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Debug.Print strResult
    
    
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
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

Public Function CalcResult(ByVal strH As String, ByVal strL As String) As String
    CalcResult = Round((CDbl("&h" & strH) * 256 + CDbl("&H" & strL)) / 256, 2)
End Function




