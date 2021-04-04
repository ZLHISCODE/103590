VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CentaurXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器名称: 化学发光免疫分析仪
'仪器型号: ADVIA centaur XP
'厂商:拜尔
'作者:陈德荣
'日期:2012-02-15

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

    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim var_Code As Variant, int_Loop As Integer
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim str_标本号 As String, str_项目 As String, str_结果 As String, str_日期 As String
    Dim str_条码 As String, str_SendLine As String, str_Item As String
    Dim str_质控 As String, str_tmp As String, aTmp As String, bTmp As String

    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    If InStr(mstrCode, Chr(&H5)) > 0 Then                                         '收到结束符 05 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(mstrCode, Chr(&HA) & Chr(&HD)) > 0 Then                              '收到行结束符 03 ,应答 06
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
 
    Do While lng_S > 0 And lng_E > 0    '收到一个完整的记录,开始解析
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        str_Code = Replace(str_Code, Chr(&HD) & Chr(&HA) & Chr(&H2), Chr(&H2))
        '--------------------------------------------------------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD))
        For int_Loop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(int_Loop)
'            writelislog ("str_Line : " & str_Line)
            If (InStr(str_Line, Chr(&H17)) > 0) And (InStr(str_Line, Chr(&H2)) > 0) Then '处理特殊符号
               aTmp = Mid(str_Line, 1, InStr(str_Line, Chr(&H17)) - 1)
               bTmp = Mid(str_Line, InStr(str_Line, Chr(&H2)) + 2)
               str_Line = aTmp + bTmp
            End If
            
            If InStr(str_Line, "O|") > 0 Then
               str_tmp = Split(str_Line, "|")(2)
               str_tmp = Split(str_tmp, "^")(0)
               str_tmp = Val(str_tmp)
               If str_标本号 <> str_tmp Then
                    If str_Item <> "" And Val(str_标本号) > 0 And str_日期 <> "" Then
                        strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_Item
                    End If
                    str_标本号 = str_tmp
                    str_Item = ""
               End If
            ElseIf InStr(str_Line, "R|") > 0 And InStr(str_Line, "^DOSE") > 0 Then
               str_日期 = Split(str_Line, "|")(12)
               If str_日期 <> "" Then
                 str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2) & ":" & Mid(str_日期, 13, 2)
                 If Not IsDate(str_日期) Then str_日期 = ""
               End If
               str_项目 = Split(str_Line, "|")(2)
               str_项目 = Trim(Split(str_项目, "^")(3))
               str_结果 = Trim(Split(str_Line, "|")(3))
               If str_结果 <> "" Then
                 str_Item = str_Item & "|" & Replace(str_项目, "^", "") & "|" & str_结果
               End If
            End If
        Next
        If str_Item <> "" And Val(str_标本号) > 0 And str_日期 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_Item
        End If
        '--------------------------------------------------------------------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H5))
        lng_E = InStr(mstrCode, Chr(&H4))
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

Public Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strLog)
    objStream.Close
    Set objStream = Nothing
End Function

