VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MQ2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:糖化血红蛋白仪
'型号:MQ2000
'厂商:上海惠中医疗科技有限公司
'作者:陈德荣
'用户:成都东区医院
'日期:2012-06-04
Private mstrCode As String  '待解析数据
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
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    Dim str_Code As String '1个完整的消息
    Dim var_Code As Variant
    Dim lng_count As Long
    Dim str_Segment As String '一个消息段(以chr(&HD)分隔)
    Dim var_Filed As Variant   '字段
    Dim strTmp As String
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_标本号 As String, str_日期 As String
    Dim str_项目 As String, str_结果 As String
    Dim str_Return As String
    
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = InStr(mstrCode, Chr(&H3))
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 1)
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If InStr(str_Segment, "sample|") > 0 Then
                var_Filed = Split(str_Segment, "|")
                str_日期 = Format(var_Filed(1), "yyyy-MM-dd hh:mm") & ":" & Format(Now, "ss")
                str_标本号 = Trim(var_Filed(2))
            ElseIf InStr(UCase(str_Segment), "HBA") > 0 Then
                '项目,结果
                str_项目 = Trim(Mid(str_Segment, 1, InStr(str_Segment, "|") - 1))
                str_结果 = Trim(Mid(str_Segment, InStr(str_Segment, "|") + 1))
                str_Return = str_Return & "|" & str_项目 & "|" & str_结果
            End If
        Next
        
        If Val(str_标本号) > 0 And str_项目 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_Return
        End If
        str_项目 = "": str_Return = "": str_结果 = ""
        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = InStr(mstrCode, Chr(&H3))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "MQ2000_Analyse", strSource, Err.Number & " " & Err.Description
End Sub

Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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

