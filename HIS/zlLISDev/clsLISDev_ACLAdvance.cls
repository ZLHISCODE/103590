VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ACLAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:ACLAdvance
'厂商:贝克曼—库尔特
'作者:陈东
'用户:重庆市第三人民医院

Private mstrCode As String

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = Chr(6)
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '           元素之间以|分隔
    '           第0个元素：检验时间
    '           第1个元素：样本序号
    '           第2个元素：检验人
    '           第3个元素：标本
    '           第4个元素：是否质控品
    '           第5～9元素：系统保留
    '           从第10个元素开始为需要的检验项目。
    '   iSendStep  : 步骤
    '   blnSuccess : 是否传输失败
    '   strResponse: 仪器的应答信息
    '   blnUndo    : 是否取消.
    '
    '   返回值：发到仪器的字串或指令。
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim var_Code As Variant, intLoop As Integer
    Dim str_标本号 As String, str_项目 As String, str_结果 As String, str_日期 As String
    Dim lng_S As Long, lng_E As Long
    strResult = ""
    
    
    If InStr(strSource, Chr(&H5)) > 0 Then                                         '收到开始符 05 ,应答 06
        strCmd = Chr(&H6)
    End If
    If InStr(strSource, Chr(&HA)) > 0 Then                                         '收到行结束符 0A ,应答 06
        strCmd = Chr(&H6)
    End If
    
   
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))

    If lng_S > 0 And lng_E > 0 Then WriteLog "解析 ACLAdvance", mstrCode, ""
    Do While lng_S > 0 And lng_E > 0                       '收到一个完整的记录,开始解析
        
        
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '-------------------------------------------------------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For intLoop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(intLoop)
            If str_Line Like Chr(&H2) & "?O|*" Then                           '标本号,项目 行
                str_标本号 = Val(Split(str_Line, "|")(2))
                If str_标本号 = 0 Then str_标本号 = ""
                str_项目 = Split(str_Line, "|")(4)
               
            ElseIf str_Line Like Chr(&H2) & "?R|*" Then                           '项目 结果 行
                If str_标本号 <> "" And str_项目 <> "" Then
                    If str_项目 = Split(str_Line, "|")(2) Then
                        str_日期 = Split(str_Line, "|")(12)
                        str_结果 = Trim(Split(str_Line, "|")(3))
                        'str_项目 = Replace(str_项目, "^^^", "") & Trim(Split(str_Line, "|")(1))
                        If str_日期 <> "" And str_结果 <> "" Then
                            str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2)
                            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|0|" & Replace(str_项目, "^^^", "") & Trim(Split(str_Line, "|")(1)) & "|" & str_结果
                        End If
                    End If
                End If
              End If
        Next
        '-------------------------------------------------------------------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H5))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
            lng_E = InStr(mstrCode, Chr(&H4))
        End If
    Loop
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        WriteLog "解析 ACLAdvance End", strResult, ""
    End If
        
End Sub

