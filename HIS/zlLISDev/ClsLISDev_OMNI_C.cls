VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_OMNI_C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血气分析仪 OMNI C
'厂商:罗氏(Roche)
'作者:陈东
'用户:塘厦

Private mstrCode As String

Private mstrWaitSend As String  '待发送的命令

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
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim str_标本号 As String, str_项目 As String, str_结果 As String, str_日期 As String
    Dim str_条码 As String, str_SendLine As String
    Dim str_质控 As String
    
    strResult = ""
    str_in = strSource
    '-------------------------------------------------------------------------------------------------------------
    If InStr(str_in, Chr(&H5)) > 0 Then                                         '收到开始符 05 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(str_in, Chr(&HA) & Chr(&HD)) > 0 Then                                            '收到行结束符 03 ,应答 06
         strCmd = Chr(&H6)
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    If mstrCode Like "*1H|*L|1|N*" & Chr(&H4) & "*" Then                         '收到一个完整的记录,开始解析
        WriteLog "开始 解析", mstrCode, ""
        Do While InStr(mstrCode, "L|1|N") > 0
            lng_S = InStr(mstrCode, "1H|")
            lng_E = InStr(mstrCode, "L|1|N")
            lng_E = InStr(lng_E, mstrCode, Chr(&H4))
            
            str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            
            Do While str_Code <> ""
                lng_S = InStr(str_Code, Chr(&H2))
                lng_E = InStr(str_Code, Chr(&H3))
                
                If lng_S > 0 And lng_E > 0 Then
                    str_Line = Mid(str_Code, lng_S, lng_E - lng_S + 1)
                    str_Code = Mid(str_Code, lng_E + 1)
                    If str_Line Like Chr(&H2) & "?O|*" Then                           '标本号,项目 行
                        str_标本号 = Split(str_Line, "|")(3)
                        str_质控 = 0
                        If UCase(Split(str_标本号, "^")(0)) = "QC" Then
                            str_质控 = 1
                        End If
                        str_标本号 = Split(str_标本号, "^")(1)
                        If Len(str_标本号) = 13 Then                            '条码长度固定为13
                            str_标本号 = "^^" & str_标本号
                        End If
                        If str_标本号 = 0 Then str_标本号 = ""
                    ElseIf InStr(str_Line, "R|") > 0 Then                       '项目 结果 行
                        If str_标本号 <> "" Then
                            str_项目 = Split(str_Line, "|")(2)
                            str_项目 = Split(str_项目, "^")(3)
                            
                            str_结果 = Split(str_Line, "|")(3)
                            If str_项目 <> "" And str_结果 <> "" Then
                                If str_日期 = "" Then
                                    str_日期 = Split(str_Line, "|")(12)
                                    If str_日期 <> "" Then
                                        str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2) & ":" & Mid(str_日期, 13, 2)
                                    Else
                                        str_日期 = Format(Now, "yyyy-MM-dd")
                                    End If
                                End If
                                
                                strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|" & str_质控 & "|" & str_项目 & "|" & str_结果

                            End If
 
                        End If
                    ElseIf InStr(str_Line, "L|1|N") > 0 Then
                        str_标本号 = ""
                    End If
                Else
                    str_Code = ""
                End If
            Loop
        Loop
    End If
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        WriteLog "完成解析", strResult, ""
    End If
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






