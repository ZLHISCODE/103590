VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_AFT_TCO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AFT-TCO2二氧化碳分析仪
'厂商:深圳康立
'作者:陈德荣
'用户:广元市中医院
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
    Dim str_Code As String, mstrCode As String
    Dim var_Code As Variant
    Dim lng_count As Integer
    
    Dim str_Segment As String '一个消息段(以chr(&HA)分隔)
    
    Dim lng_S As Integer
    Dim lng_E As Integer
    
    Dim str_标本号 As String, str_日期 As String
    Dim str_项目 As String, str_结果 As String
    Dim str_Return As String, strLog As String
    
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    If Len(strSource) < 45 Then strReserved = strSource: Exit Sub
    
    strLog = Format(Now, "yyyy-MM-dd HH:mm:ss") & " 接收原始数据：" & strSource
    Call writelislog(strLog)
    
    mstrCode = Replace(mstrCode, "血样报告编号", Chr(&HB))
    mstrCode = Replace(mstrCode, "MMOL/L", Chr(&H1C))
    
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
                 
        var_Code = Split(str_Code, Chr(&HA))
        str_Return = "": str_标本号 = "": str_日期 = ""
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            Select Case lng_count
                Case 0
                   str_标本号 = Trim(str_Segment)
                   str_标本号 = Trim(Replace(str_标本号, Chr(&HB), ""))
                Case 1
                   str_Segment = Replace(str_Segment, Chr(&HA), "")
                   str_Segment = Replace(str_Segment, Chr(&HD), "")
                   str_日期 = Trim(str_Segment) & ":" & Format(Now, "SS")
                Case 3
                   str_Segment = Replace(str_Segment, Chr(&H1C), "")
                   str_Segment = Replace(str_Segment, Chr(&HA), "")
                   str_项目 = Trim(Mid(str_Segment, 1, InStr(str_Segment, "=") - 1))
                   str_结果 = Trim(Mid(str_Segment, InStr(str_Segment, "=") + 1))
                   str_结果 = Trim(Replace(str_结果, Chr(&H20), ""))
                   str_Return = str_Return & "|" & str_项目 & "|" & str_结果
            End Select
        Next
        If Val(str_标本号) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_Return
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
'    writelislog (strResult)
End Sub

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_CO2_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function


