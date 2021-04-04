VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_NFT400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: 生化分析仪
'厂商: 东软
'作者: 陈东
'用户: 山西中条山医院
Private mstrCode As String
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
    Dim lng_S As Long, lng_E As Long, str_Code As String
    Dim var_Code As Variant, str_Line As String, i As Integer
    Dim str_NO As String, str_Date As String, str_Item As String, str_Value As String
    
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    If InStr(strSource, ">") > 0 Then strCmd = "<Saved|1F3>"
    
    lng_S = InStr(mstrCode, "<")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, ">")
    Else
        lng_E = InStr(mstrCode, ">")
    End If
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------
        str_Item = Trim$(Split(str_Code, "|")(3))
        str_Value = Trim$(Split(str_Code, "|")(4))
        str_NO = Trim$(Split(str_Code, "|")(2))
        If Val(str_NO) > 0 And str_Value <> "" And str_Item <> "" Then
            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm") & "|" & str_NO & "| |血清|0|" & str_Item & "|" & str_Value
        ElseIf Trim("" & str_NO) <> "" And str_Value <> "" And str_Item <> "" Then
            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm") & "|" & str_NO & "| |血清|1|" & str_Item & "|" & str_Value
        End If
        '--------------------------------
        lng_S = InStr(mstrCode, "<")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, ">")
        Else
            lng_E = InStr(mstrCode, ">")
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
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



