VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_KC100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'产品名称: KC-100酶标仪
'产品型号: KC-100
'厂商:深圳凯特生物医疗电子有限公司
'作者:李建飞
'日期:2010-08-25
'通讯参数：19200 N 8 1
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
    '2010-01-14|1| |血清|0|Alb|44|ALT|6|TP|72|T－Bil|10|D－Bil|2|AST|9|ALP|118|IBil|8|GLB|28|A/G|1.57

    Dim str_Code As String '1个完整的消息
    Dim var_Code As Variant
    Dim lng_count As Long
    Dim var_Filed As Variant   '字段
    
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_标本号 As String, str_日期 As String, str_标本类型 As String
    Dim str_质控 As String
    Dim str_检验项目 As String
    Dim str_Return As String
    Dim str_Line As String
    
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
 
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
                 
        var_Code = Split(str_Code, Chr(&HD))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(lng_count)
            
            If Mid(str_Line, 2, 3) = "MSH" Then
                str_检验项目 = "": str_标本号 = "": str_日期 = ""
            ElseIf Mid(str_Line, 1, 3) = "OBR" Then
                '标本号，日期
                var_Filed = Split(str_Line, "|")
                str_标本号 = var_Filed(2)
                str_日期 = Format(var_Filed(4), "yyyy-MM-dd")
                str_标本类型 = var_Filed(8)
                str_质控 = var_Filed(11)
            ElseIf Mid(str_Line, 1, 3) = "OBX" Then
                var_Filed = Split(str_Line, "|")
                str_检验项目 = str_检验项目 & "|" & var_Filed(3) & "|" & var_Filed(4) & "^" & var_Filed(8) & "^" & var_Filed(9) & "^" & 0
            End If
        Next
        
        
        If Val(str_标本号) > 0 And str_检验项目 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |" & str_标本类型 & "|" & IIf(str_质控 = "0", 0, 1) & str_检验项目
        End If
        str_检验项目 = ""
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "KC100_analyse", strSource, Err.Number & " " & Err.Description
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





