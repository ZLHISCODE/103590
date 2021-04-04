VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CA1500_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:CA-1500
'厂商:日本东亚 sysmex
'用户:广西民族医院
'开发:陈东
Private mstrCode As String
Private mstr架号  As String
Private mstr项目 As String

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
    
    Dim lng_S As Long, lng_E As Long, str_Code As String, str_Line As String, i As Integer
    Dim str_日期 As String, str_架号 As String, str_质控 As String, str_条码 As String, str_类型 As String
    Dim str_结果 As String, str_项目 As String, strDateType As String, intDataStart As Integer
    Dim intNoStart As Integer, intNoLen As Integer
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    '-- 配置文件读取基础数据
    
    strDateType = Trim(ReadIni("CA1500New", "日期格式", App.Path & "\CA1500New.ini"))
    If strDateType = "" Then Call WriteIni("CA1500New", "日期格式", "MM-DD-YY", App.Path & "\CA1500New.ini")

    intDataStart = Val(ReadIni("CA1500New", "结果位置", App.Path & "\CA1500New.ini"))
    If intDataStart <= 0 Then
        intDataStart = 60
        Call WriteIni("CA1500New", "结果位置", CStr(intDataStart), App.Path & "\CA1500New.ini")
    End If
    
    intNoStart = Val(ReadIni("CA1500New", "标本位置", App.Path & "\CA1500New.ini"))
    If intNoStart <= 0 Then
        intNoStart = 29
        Call WriteIni("CA1500New", "标本位置", CStr(intNoStart), App.Path & "\CA1500New.ini")
    End If
    
    intNoLen = Val(ReadIni("CA1500New", "标本长度", App.Path & "\CA1500New.ini"))
    If intNoLen <= 0 Then
        intNoLen = 15
        Call WriteIni("CA1500New", "标本长度", CStr(intNoLen), App.Path & "\CA1500New.ini")
    End If
    '------
    If InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    End If
    
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '--------------------------------------------------
        If Mid(str_Code, 2, 1) = "R" Then       '双向：仪器扫条码后发的 指令
            str_日期 = "20" & Mid(str_Code, 11, 2) & "-" & Mid(str_Code, 13, 2) & "-" & Mid(str_Code, 15, 2) & " " & Mid(str_Code, 17, 2) & ":" & Mid(str_Code, 19, 2)
            
            str_条码 = Trim(Mid(str_Code, 27, 13))
            str_类型 = Mid(str_Code, 40, 1) ' M"手动输入,"A"仪器赋值,"B"条形码读取器读取,"C"由主计算机指令信息设定
            
            mstr架号 = Mid(str_Code, 21, 6)
            mstr项目 = Mid(str_Code, 52)
            mstr项目 = Replace(mstr项目, Space(6) & Chr(&H3), "")
            mstr项目 = Replace(mstr项目, Space(6), ",")
            If Len(str_条码) >= 13 Then
                strResult = strResult & "||" & str_日期 & "|^^" & str_条码
                strCmd = "1|"
            End If
            
        ElseIf Mid(str_Code, 2, 1) = "D" Then   '      仪器传回的检验结果
            str_质控 = Mid(str_Code, 10, 1)
            If str_质控 = "C" Then
                str_质控 = "1"
            Else
                str_质控 = "0"
            End If
            If strDateType = "YY-MM-DD" Then
                
                str_日期 = "20" & Mid(str_Code, 11, 2) & "-" & Mid(str_Code, 13, 2) & "-" & Mid(str_Code, 15, 2) & " " & Mid(str_Code, 17, 2) & ":" & Mid(str_Code, 19, 2)
            ElseIf strDateType = "DD-MM-YY" Then
                str_日期 = "20" & Mid(str_Code, 15, 2) & "-" & Mid(str_Code, 13, 2) & "-" & Mid(str_Code, 11, 2) & " " & Mid(str_Code, 17, 2) & ":" & Mid(str_Code, 19, 2)
            Else
                '"MM-DD-YY"
                str_日期 = "20" & Mid(str_Code, 15, 2) & "-" & Mid(str_Code, 11, 2) & "-" & Mid(str_Code, 13, 2) & " " & Mid(str_Code, 17, 2) & ":" & Mid(str_Code, 19, 2)
            End If
            
            str_架号 = Mid(str_Code, 21, 6)
            str_条码 = Trim(Mid(str_Code, intNoStart, intNoLen))   '27,13
            
            str_类型 = Mid(str_Code, 40, 1) ' M"手动输入,"A"仪器赋值,"B"条形码读取器读取,"C"由主计算机指令信息设定
            str_项目 = ""
            str_Code = Mid(str_Code, intDataStart) '52 或  60
            'If str_条码 = 48 Then Stop
            For i = 0 To (Len(str_Code) - 1) / 9
                If Len(str_Code) < 9 Then Exit For
                str_结果 = Trim(Mid(str_Code, 4, 6))
                If str_结果 <> "" Then
                    str_项目 = str_项目 & "|" & Mid(str_Code, 1, 3) & "|" & Trim(Mid(str_Code, 4, 5))
                End If
                str_Code = Mid(str_Code, 10)
            Next
            If Len(str_条码) >= 13 Then
                strResult = strResult & "||" & str_日期 & "|^^" & str_条码 & "| |血清|" & str_质控 & str_项目
            ElseIf Val(str_条码) > 0 Then
                strResult = strResult & "||" & str_日期 & "|" & str_条码 & "| |血清|" & str_质控 & str_项目
            End If
        End If
        
        '--------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
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
    
    Dim var_标本 As Variant, i_Count As Integer, i_Loop As Integer, str_记录 As String, str_标本号 As String
    Dim i_急诊 As Integer, str_条码 As String, str_检验项目 As String, str_标本类型 As String, str_tmp As String
    Dim str_SendItem As String, str_Line As String
    Dim str项目 As String
    
    '-------------------
    'str项目 = "040,050,060,080,090,120,150,170,180,190,200,210,220,300,310,320,330,510,600,610"
    str项目 = mstr项目
    '----------------------------
    WriteLog "开始", strSampleInfo, ""
    
    var_标本 = Split(strSampleInfo, "||")
    SendSample = ""
    For i_Count = LBound(var_标本) To UBound(var_标本)
        str_记录 = var_标本(i_Count)
        
        If Val(Split(str_记录, "|")(1)) = 0 Then
            str_标本号 = Trim(Split(str_记录, "|")(5))
        Else
            str_标本号 = Val(Split(str_记录, "|")(1))
        End If
        str_标本类型 = Split(str_记录, "|")(3)
        i_急诊 = Val(Split(str_记录, "|")(4))
        str_条码 = Trim(Split(str_记录, "|")(5))
        str_检验项目 = ""
        str_tmp = ""
        For i_Loop = 10 To UBound(Split(str_记录, "|"))
            If Val(Split(str_记录, "|")(i_Loop)) > 0 Then
                str_tmp = str_tmp & "," & Mid(Split(str_记录, "|")(i_Loop), 1, 2)
            End If
        Next
        
        For i_Loop = 0 To UBound(Split(str项目, ","))
            If InStr(str_tmp, Mid(Split(str项目, ",")(i_Loop), 1, 2)) > 0 Then
                str_检验项目 = str_检验项目 & Split(str项目, ",")(i_Loop) & Space(6)
            Else
                str_检验项目 = str_检验项目 & "000" & Space(6)
            End If
        Next
        If str_检验项目 <> "" And str_条码 <> "" Then
            
            str_Line = Chr(&H2) & "S2210101U" & Format(Now, "yyMMddHHmm") & mstr架号 & str_条码 & "B" & Space(11) & str_检验项目 & Chr(&H3)

        End If
    Next
    iSendStep = 0
    WriteLog "结束", "发送到仪器的信息", SendSample
    
End Function


