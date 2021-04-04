VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_FUS200_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器名称: FUS200全自动尿沉渣分析仪
'仪器型号: FUS200
'厂商:长春迪瑞医疗器械有限公司
'作者:曾超
'日期:2012-10-17



Private mstrCode As String    '缓存接受到的数据


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


    Dim lng_S As Long                                   '开始的位置
    Dim lng_E As Long                                   '结束的位置
    Dim strDate As String                               '分析出的时间
    Dim strNo As String                                 '分析出的标本号
    Dim str_Items As String                             '分析出的通道码及结果
    Dim str_Sample As String
    Dim var_Tmp As Variant
    Dim intloop As Integer
    Dim strLine As String
    
    '缓存接受到得数据
    mstrCode = mstrCode & strSource

    '判断是否收到了一个完整的数据
    lng_S = InStr(mstrCode, Chr(&HB))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H1C))
    Else
        lng_E = InStr(mstrCode, Chr(&H1C))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        '收到完整的数据，取出
        str_Sample = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        '取出日期，标本号，检验结果
        var_Tmp = Split(str_Sample, Chr(&HD))
        For intloop = LBound(var_Tmp) To UBound(var_Tmp)
            strLine = var_Tmp(intloop)
            If InStr(strLine, Chr(&HB)) > 0 Then
                strDate = Split(strLine, "|")(6)
                strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                strDate = Format(strDate, "YYYY-MM-DD hh:mm:ss")
            ElseIf InStr(strLine, "PID") > 0 Then
                strNo = Split(strLine, "|")(3)
            ElseIf InStr(strLine, "OBX") > 0 And InStr(strLine, "NM") > 0 And InStr(strLine, "Chemistry") > 0 Then
                If Split(strLine, "|")(3) = "WBC" Then
                    str_Items = str_Items & "|" & Split(strLine, "|")(3) & "S|" & Split(Split(strLine, "|")(5), "^")(1)
                Else
                    If Split(Split(strLine, "|")(5), "^")(1) <> "" Then
                        str_Items = str_Items & "|" & Split(strLine, "|")(3) & "|" & Split(Split(strLine, "|")(5), "^")(1)
                    Else
                        str_Items = str_Items & "|" & Split(strLine, "|")(3) & "|" & Split(Split(strLine, "|")(5), "^")(2)
                    End If
                End If
            ElseIf InStr(strLine, "OBX") > 0 And InStr(strLine, "NM") > 0 And InStr(strLine, "Sediment") > 0 Then
                str_Items = str_Items & "|" & Split(strLine, "|")(3) & "|" & Split(strLine, "|")(5)
            End If
        Next
        '保存取出的结果
        If strNo <> 0 And str_Items <> "" Then
            If strResult <> "" Then
                strResult = strResult & "||" & strDate & "|" & strNo & "| |尿液|0" & str_Items
            Else
                strResult = strDate & "|" & strNo & "| |尿液|0" & str_Items
            End If
        End If
        '判断缓存中的数据是否取完
        lng_S = InStr(mstrCode, Chr(&HB))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H1C))
        Else
            lng_E = InStr(mstrCode, Chr(&H1C))
        End If
    Loop
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
