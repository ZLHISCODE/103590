VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Fully"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:BSI Fully 全自动生化分析仪
'厂商:
'作者:陈东
'用户:贵阳玉屏县人民医院

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
   
    Dim lng_S As Long, lng_E As Long
    Dim str_Code As String '一个完整记录
    Dim str_标本号 As String, str_项目 As String, str_结果 As String, str_日期 As String
    Dim str_条码 As String

    strResult = ""
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, "R")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&HD))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '----------------------------------
        If Len(str_Code) >= 154 Then
            str_日期 = Mid(str_Code, 147, 8)
            If IsNumeric(str_日期) Then
                str_日期 = Mid(str_日期, 5, 4) & "-" & Mid(str_日期, 3, 2) & "-" & Mid(str_日期, 1, 2)
                str_标本号 = Trim(Mid(str_Code, 3, 5))
                If Val(str_标本号) > 0 Then
                    str_项目 = Trim(Mid(str_Code, 57, 6))
                    str_结果 = Trim(Mid(str_Code, 63, 18))
                    If str_结果 <> "" And str_项目 <> "" Then
                        strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0|" & str_项目 & "|" & str_结果
                    End If
                End If
            Else
                str_日期 = ""
            End If
        End If
        '----------------------------------
        lng_S = InStr(mstrCode, "R")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&HD))
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)

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







