VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XL600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:德国欧宝XL-600全自动生化分析仪
'型号:XL-600
'厂商:德国欧宝
'Author:李建飞
'User:安徽常君
'Date:2010-05-13

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
    
    Dim lng_S As Long, lng_E As Long
    Dim str_标本号 As String, str_标本 As String
    Dim str_检验日期 As String
    Dim str_检验结果 As String
    Dim var_Tmp() As String
    Dim str_Data As String
    Dim blnEnd As Boolean
    
    strReserved = strReserved & strSource
    lng_S = InStr(strReserved, "<")
    lng_E = InStr(strReserved, ">")
    
    If lng_S > 0 And lng_E > 0 Then
        str_Data = Mid(strReserved, lng_S, lng_E)
        strReserved = Mid(strReserved, lng_E + 1)
        
        str_Data = Replace(Replace(str_Data, "<", ""), ">", "")
        
        var_Tmp = Split(str_Data, "|")
        str_标本号 = var_Tmp(2)
        
        str_标本 = "血清"
        str_检验日期 = Format(Now, "yyyy-MM-dd hh:mm:ss")
        str_检验结果 = str_检验结果 & "|" & Trim(var_Tmp(3)) & "|" & Trim(var_Tmp(4))
        
        If Val(str_标本号) > 0 And Trim(str_检验结果) <> "" Then
            strResult = str_检验日期 & "|" & str_标本号 & "| |" & str_标本 & "|0" & str_检验结果
        End If
        strCmd = "  " & Chr(&H3C) & Chr(&H53) & Chr(&H61) & Chr(&H76) & Chr(&H65) & Chr(&H64) & Chr(&H7C) & Chr(&H31) & Chr(&H46) & Chr(&H33) & Chr(&H3E)
    End If
End Sub


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
    'iSendStep：当前发送操作的步骤。如果处理完后发送结束，则返回0，否则自动＋1返回。
    'strResponse：当前仪器的响应内容
    'strSampleInfo：发送的标本信息。
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否急诊
    '   第5个无素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
End Function
