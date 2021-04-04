VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_NovePHOx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  血气分析仪 Nove pHox
'医院:  广西民族医院
'作者:  陈东
'时间:  2009/4/16

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
    Dim str_标本号 As String, str_日期 As String, str_项目 As String, str_结果 As String, str_Items As String
    
    strReserved = "": strResult = "": strCmd = ""
    
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    End If
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    Else
        lng_E = InStr(mstrCode, Chr(&H4))
    End If
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------
        var_Code = Split(str_Code, Chr(&HA))
        For i = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(i)
            If str_Line Like Chr(&H5) & Chr(&H2) & "?H|*" Then
                str_日期 = Trim(Split(str_Line, "|")(13))
                If IsNumeric(str_日期) And Len(str_日期) >= 12 Then
                    str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2)
                Else
                    str_日期 = Format(Now, "yyyy-MM-dd")
                End If
                str_标本号 = ""
                str_Items = ""
            ElseIf str_Line Like Chr(&H2) & "?O|*" Then
                str_标本号 = Trim(Split(str_Line, "|")(2))
                
            ElseIf str_Line Like Chr(&H2) & "?R|*" Then
                str_结果 = Trim(Split(str_Line, "|")(3))
                str_项目 = Trim(Split(Split(str_Line, "|")(2), "^")(3))
                If str_结果 <> "" And str_项目 <> "" Then
                    str_Items = str_Items & "|" & str_项目 & "|" & str_结果
                End If
            End If
        Next
        
        If Val(str_标本号) > 0 And str_Items <> "" And IsDate(str_日期) Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|0" & str_Items
        End If
        '--------------------------------
        lng_S = InStr(mstrCode, Chr(&H5))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
            lng_E = InStr(mstrCode, Chr(&H4))
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
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
    SendSample = ""
End Function



