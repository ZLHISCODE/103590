VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AIA360"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:化学发光分析仪
'型号:AIA360
'厂商:日本东曹株式会社
'作者:陈德荣
'用户:彭州市精神病医院
'修改日期:2012-05-25

Private mstrCode As String    '待解析串

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

    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim var_Code As Variant, int_Loop As Integer
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim str_标本号 As String, str_项目 As String, str_结果 As String, str_日期 As String
    Dim str_条码 As String, str_SendLine As String, str_Item As String, cTdh As String
    Dim str_质控 As String, str_Tmp As String, aTmp As String, bTmp As String

    On Error GoTo hErr
    strReserved = "": strResult = "": strCmd = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    If InStr(mstrCode, Chr(&H5)) > 0 Then                                         '收到结束符 05 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&HA)) > 0 Then                                         '收到行结束符 10 ,应答 06
         strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&H3)) > 0 Then                                         '收到开始符 03 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&H4)) > 0 Then                                         '收到开始符 04 ,应答 06
        strCmd = Chr(&H6)
    End If
    '-------------------------------------------------------------------------------------------------------------
 
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = InStr(mstrCode, Chr(&H4))
    
    If lng_S <= 0 And lng_E <= 0 Then
       strReserved = mstrCode
    End If
    
    Do While lng_S > 0 And lng_E > 0    '收到一个完整的记录,开始解析
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        '--------------------------------------------------------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)
            If (InStr(str_Line, Chr(&H17)) > 0) And (InStr(str_Line, Chr(&H2)) > 0) Then '处理特殊符号
               aTmp = Mid(str_Line, 1, InStr(str_Line, Chr(&H17)) - 1)
               bTmp = Mid(str_Line, InStr(str_Line, Chr(&H2)) + 2)
               str_Line = aTmp + bTmp
            End If
            str_Line = Replace(str_Line, Chr(&H2), "")
            
            If Mid(str_Line, 2, 1) = "O" Then
               str_Tmp = Trim(Split(str_Line, "|")(2))
               str_Tmp = Val(str_Tmp)
               
               str_质控 = "0"
               str_质控 = IIf(Trim(Split(str_Line, "|")(11)) = "Q" Or Trim(Split(str_Line, "|")(11)) = "X\Q", "1", "0")
               
               If str_标本号 <> str_Tmp Then
                    If str_Item <> "" And Val(str_标本号) > 0 And str_日期 <> "" Then
                        strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|" & str_质控 & str_Item
                    End If
                    str_标本号 = str_Tmp
                    str_Item = ""
               End If
            ElseIf Mid(str_Line, 2, 1) = "R" Then
               If UBound(Split(str_Line, "|")) >= 3 Then
                 str_日期 = Split(str_Line, "|")(12)
                 If str_日期 <> "" Then
                    str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2) & ":" & Mid(str_日期, 13, 2)
                    If Not IsDate(str_日期) Then str_日期 = ""
                  End If
                  str_项目 = Split(str_Line, "|")(2)
                  str_项目 = Trim(Replace(str_项目, "^", ""))
                  str_结果 = Trim(Split(str_Line, "|")(3))
                  If str_结果 <> "" Then
                    str_Item = str_Item & "|" & Replace(str_项目, "^", "") & "|" & str_结果
                  End If
               End If
            End If
        Next
        If str_Item <> "" And Val(str_标本号) > 0 And str_日期 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|" & str_质控 & str_Item
        End If
        lng_S = InStr(mstrCode, Chr(&H5))
        lng_E = InStr(mstrCode, Chr(&H4))
    Loop
    '--------------------------------------------------------------------------------------------------------------
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
hErr:
    WriteLog "AIA360", strSource, Err.Number & " " & Err.Description
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

    Dim var_标本 As Variant, i_Count As Integer, i_Loop As Integer, str_记录 As String, str_标本号 As String
    Dim i_急诊 As Integer, str_条码 As String, str_检验项目 As String, str_标本类型 As String
    Dim str_SendItem As String, str_Line As String
    
    SendSample = ""

    
End Function

