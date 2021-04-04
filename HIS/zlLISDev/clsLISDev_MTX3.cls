VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MTX3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:MTX3血凝仪
'厂商:爱尔兰Trinity Biotech
'作者:陈德荣
'用户:四川电力职工医院
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
    Dim i As Integer
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
    Dim str_Code As String '1个完整的消息
    Dim var_Code As Variant
    Dim lng_count As Long
    Dim str_Segment As String '一个消息段(以chr(&HD)分隔)
    Dim var_Filed As Variant   '字段
    Dim strTmp As String
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_标本号 As String, str_日期 As String
    Dim str_项目 As String
    Dim str_结果 As String
    
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
'    If InStr(mstrCode, Chr(&H2)) > 0 Then                 '收到02 应答 ETX(03)
'        strCmd = Chr(&H6)
'    End If
    
    If InStr(mstrCode, Chr(&H3)) > 0 Then                 '收到ETX(03)
        strCmd = Chr(&H6)
    End If
    
    'Call OriData("MTX3", strSource)
    
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = InStr(mstrCode, Chr(&H3))
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        str_日期 = Format(Now, "yyyy-MM-dd HH:mm:ss")
        str_标本号 = Val(Trim(Mid(str_Code, 4, 5)))
        If Val(Trim(Mid(str_Code, 31, 8))) <> 0 Then
           str_项目 = Trim(Mid(str_Code, 27, 3)) & "|" & Val(Trim(Mid(str_Code, 31, 8)))
        End If
        If Val(str_标本号) > 0 And str_项目 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0|" & str_项目
        End If
        str_项目 = "": str_标本号 = ""
        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = InStr(mstrCode, Chr(&H3))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "MTX3_Analyse", strSource, Err.Number & " " & Err.Description
End Sub

Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    SendSample = ""
    iSendStep = 0
End Function
    

