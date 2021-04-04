VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_XD_683"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  XD683 电解质 迅达
'医院:  南桐总医院
'作者:  陈东
'时间:  2008/12/12

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
    '   微生物结果格式 通道码|结果类型^方式^结果值
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令

    Dim str_Code As String
    Dim str_Line As String               '当前要解析的字串
    Dim lng_S As Long                  '开始位置
    Dim lng_E As String                '结束位置
    
    Dim str_日期 As String, str_标本号 As String, str_项目 As String
    
    mstrCode = mstrCode & strSource
    On Error GoTo errH
    strResult = ""
    If mstrCode Like "*XD  683*Cl*" & Chr(&HA) & "*" Then
        
        Do While InStr(mstrCode, "Cl") > 0
            lng_S = InStr(mstrCode, "XD  683")
            lng_E = InStr(lng_S, mstrCode, "Cl")
            lng_E = InStr(lng_E, mstrCode, Chr(&HA))
            str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            
            str_日期 = "": str_标本号 = "": str_项目 = ""
            Do While str_Code <> ""
                lng_S = 1
                lng_E = InStr(str_Code, Chr(&HA))
                str_Line = Mid(str_Code, lng_S, lng_E - lng_S + 1)
                str_Code = Mid(str_Code, lng_E + 1)
                
                If str_Line Like "DATE*" Then
                    str_日期 = Trim(Mid(str_Line, 5))
                ElseIf str_Line Like "Sample No*" Then
                    str_标本号 = Val(Mid(str_Line, 12))
                ElseIf str_Line Like "*=*" Then
                    str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 1, InStr(str_Line, "=") - 1)) & "|" & Trim(Mid(str_Line, InStr(str_Line, "=") + 1, 5))
                End If
                
            Loop
            
            If Val(str_标本号) <> 0 And IsDate(str_日期) And str_项目 <> "" Then
                strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血液|0" & str_项目
            End If
        Loop
    End If
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    WriteLog "err", Err.Description, strResult
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    
End Function







