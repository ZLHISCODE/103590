VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CS1300A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器: CS1300A 全自动化学发光仪
'厂商: 希森美康（长春迪瑞）
'作者: 黄金桥
'用户: 名山区人民医院
'时间: 2013-08-30
'集成时间 2014-1-28

Private mstrCode As String  '待解析数据
'--------------------------------------------------------------------------------------------------------------------
'以下是必须的接口函数
Public Function GetAnswerCmd() As String
    '自动应答指令
    GetAnswerCmd = Chr(&H6)
End Function

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
    Dim str_Code As String '1个完整的消息
    Dim var_Code As Variant
    Dim lng_count As Long
    
    Dim str_字段分隔符 As String
    Dim str_组件分隔符 As String
    Dim str_子组件隔符 As String
    Dim str_重复分隔符 As String
    Dim str_转义字符 As String
    
    Dim str_Segment As String '一个消息段(以chr(&HD)分隔)
    Dim var_Filed As Variant   '字段
    Dim str_唯一标识 As String '用于标识一个消息
    Dim str_版本 As String
    
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_标本号 As String, str_日期 As String
    Dim str_质控 As String
    Dim str_项目 As String
    Dim str_结果 As String
    Dim str_Return As String
      
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & Replace(strSource, Chr(&H20), "")
    End If
    strCmd = Chr(&H6)
    
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C))
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        var_Code = Split(str_Code, Chr(&HD))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If Mid(str_Segment, 2, 3) = "MSH" Then
                '消息头
                str_字段分隔符 = Mid(str_Segment, 5, 1)
                str_组件分隔符 = Mid(str_Segment, 6, 1)
                str_子组件隔符 = Mid(str_Segment, 7, 1)
                str_重复分隔符 = Mid(str_Segment, 8, 1)
                str_转义字符 = Mid(str_Segment, 9, 1)
                
                var_Filed = Split(str_Segment, str_字段分隔符)
                str_唯一标识 = var_Filed(9)
                str_质控 = var_Filed(10)
                str_版本 = var_Filed(11)
                If str_版本 <> "2.3.1" Then
                    WriteLog "版本不符", "", ""
                End If
                str_Return = "": str_标本号 = "": str_日期 = ""
            ElseIf InStr(str_Segment, "OBR|") > 0 Then
                '标本号，日期
                If Val(str_标本号) > 0 And str_Return <> "" Then
                    strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|" & IIf(str_质控 = "Q", 1, 0) & str_Return
                End If
                var_Filed = Split(str_Segment, str_字段分隔符)
                str_标本号 = var_Filed(3)
                str_日期 = var_Filed(7)
                str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2) & ":" & Mid(str_日期, 13, 2)
                str_Return = ""
            ElseIf InStr(str_Segment, "OBX|") > 0 Then
                var_Filed = Split(str_Segment, str_字段分隔符)
                str_项目 = var_Filed(3)
                str_Return = str_Return & "|" & Split(str_项目, str_组件分隔符)(0) & "|" & Trim(var_Filed(4))
            End If
        Next
        If Val(str_标本号) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|" & IIf(str_质控 = "Q", 1, 0) & str_Return
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "CS1300_analyse", strSource, Err.Number & " " & Err.Description
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




