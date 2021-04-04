VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Cobas400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Cobas400
'厂商:罗氏
'作者:陈东
'用户:晋煤集团总医院
Private mstrCode As String '待解析串

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = Chr(&H1) & Chr(&HA) & "14 COBAS INTEGRA400 09" & Chr(&HA) & Chr(&H2) & Chr(&HA) & "10  7" & Chr(&HA) & Chr(&H3) & Chr(&HA) & Chr(&H4) & Chr(&HA)
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
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant, str_INTEG As String, int_Loop As Integer
    Dim str_Line As String
    Dim str_标本号 As String, str_日期 As String, str_项目 As String, str_结果 As String
    
    
    'str_INTEG = Chr(&H1) & Chr(&HA) & "14 COBAS INTEGRA400 00" & Chr(&HA) & Chr(&H2) & Chr(&HA) & Chr(&H3) & Chr(&HA) & "1" & Chr(&HA) & "428" & Chr(&HA) & Chr(&H4) & Chr(&HA)
    
    strResult = ""
    mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, Chr(&H1) & Chr(&HA))
    lng_E = InStr(mstrCode, Chr(&H4) & Chr(&HA))
    
    Do While lng_S > 0 And lng_E > 0
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        
        mstrCode = Mid(mstrCode, lng_E + 2)
        lng_S = InStr(mstrCode, Chr(&H1) & Chr(&HA))
        lng_E = InStr(mstrCode, Chr(&H4) & Chr(&HA))
        
        '用于保持通讯联络
        Call Pause(1) '延时一秒再应答
        strCmd = Chr(&H1) & Chr(&HA) & "14 COBAS INTEGRA400 09" & Chr(&HA) & Chr(&H2) & Chr(&HA) & "10  7" & Chr(&HA) & Chr(&H3) & Chr(&HA) & Chr(&H4) & Chr(&HA)
    
        str_标本号 = "": str_日期 = "": str_项目 = "": str_结果 = ""
        var_Code = Split(str_Code, Chr(&HA))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If Left(str_Line, 3) = "53 " Then
                str_标本号 = Trim(Mid(str_Line, 4, 5))
                str_日期 = Mid(str_Line, 26, 4) & "-" & Mid(str_Line, 23, 2) & "-" & Mid(str_Line, 20, 2)
                
            ElseIf Left(str_Line, 3) = "55 " Then
                str_项目 = Trim(Mid(str_Line, 4))
            ElseIf Left(str_Line, 3) = "00 " Then
                If InStr(Trim(Mid(str_Line, 4, 14)), ">") > 0 And InStr(Trim(Mid(str_Line, 4, 14)), "<") > 0 Then
                    str_结果 = Trim(Mid(str_Line, 4, 14))
                Else
                    str_结果 = Val(Trim(Mid(str_Line, 4, 14)))
                End If
            End If
        Next
        If str_日期 = "" Then str_日期 = Format(Now, "yyyy-MM-dd")
        If Val(str_标本号) > 0 And IsDate(CDate(str_日期)) And str_项目 <> "" And str_结果 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血液|0|" & str_项目 & "|" & str_结果
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    '   标本之间以||分隔
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

