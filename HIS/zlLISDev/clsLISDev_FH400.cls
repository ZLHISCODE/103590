VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_FH400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:全自动生化仪
'型号:FH-400
'厂商:上海丰汇
'用户:钦北区人民医院
'时间：2013/12/19
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

Public Function GetAnswerCmd() As String
'定时发送的指令
    GetAnswerCmd = ""
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
    Dim str日期 As String
    Dim str结果 As String, str标本号 As String
    Dim str通道码 As String, lng_S As Long
    Dim lng_E As Long, iRow As Long, str_Code As String
    Dim varLine As Variant, strLine As String
    Dim aResult As String, varline2 As Variant
    Dim strline2 As String, j As Long
    Dim str_检验人标本 As String

    str_检验人标本 = " |血清|0"
    On Error GoTo errH
    
    If strSource <> "" Then
        strCmd = Chr(&H6&)
    End If
    mstrCode = mstrCode & strSource
    
    If InStr(Mid(mstrCode, 1, 1), Chr(&H4E)) > 0 Then     '判断普通标本以H4E开头,返回完整段落的开始位置
        lng_S = InStr(mstrCode, Chr(&H4E))
    ElseIf InStr(Mid(mstrCode, 1, 1), Chr(&H45)) > 0 Then    '判断急诊标本以H45开头,返回完整段落的开始位置
        lng_S = InStr(mstrCode, Chr(&H45))
    ElseIf InStr(Mid(mstrCode, 1, 1), Chr(&H43)) > 0 Then    '判断质控标本以H43开头,返回完整段落的开始位置
        lng_S = InStr(mstrCode, Chr(&H43))
    End If
    
    If lng_S > 0 Then
        lng_E = InStr(mstrCode, Chr(&HD) & Chr(&HA))      '判断以HD或者HA结束,返回完整段落的结束位置
    End If
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
        varLine = Split(str_Code, Chr(&HD))  '截取每行数据
        For iRow = LBound(varLine) To UBound(varLine)
            strLine = varLine(iRow)
            
            If InStr(strLine, Chr(&H4E)) > 0 Then
                varline2 = Split(str_Code, Chr(&H9))(0)  '每行数据以H9分组
                str日期 = Format(Now, "yyyy-mm-dd hh:mm:ss")
                str标本号 = Mid(Split(str_Code, Chr(&H9))(0), 2)    'Replace(Mid(strLine, 2, 3), Chr(&H9), "")
                str通道码 = Split(str_Code, Chr(&H9))(1)
                str结果 = Split(str_Code, Chr(&H9))(3)
            ElseIf InStr(strLine, Chr(&H45)) > 0 Then
                varline2 = Split(str_Code, Chr(&H9))(0)
                str日期 = Format(Now, "yyyy-mm-dd hh:mm:ss")
                str标本号 = Mid(Split(str_Code, Chr(&H9))(0), 2) + 1000
                str通道码 = Split(str_Code, Chr(&H9))(1)
                str结果 = Split(str_Code, Chr(&H9))(3)
            ElseIf InStr(strLine, Chr(&H43)) > 0 Then
                varline2 = Split(str_Code, Chr(&H9))(0)
                str日期 = Format(Now, "yyyy-mm-dd hh:mm:ss")
                str标本号 = Mid(Split(str_Code, Chr(&H9))(0), 2)
                str通道码 = Split(str_Code, Chr(&H9))(1)
                str结果 = Split(str_Code, Chr(&H9))(3)
            End If
            
        Next
        
        If str结果 <> "" And str标本号 <> "" Then
            '最后拼结果
            aResult = aResult & "||" & str日期 & "|" & str标本号 & "|" & str_检验人标本 & "|" & str通道码 & "|" & str结果
        End If
        
        If InStr(Mid(mstrCode, 1, 1), Chr(&H4E)) > 0 Then
            lng_S = InStr(mstrCode, Chr(&H4E))
        ElseIf InStr(Mid(mstrCode, 1, 1), Chr(&H45)) > 0 Then
            lng_S = InStr(mstrCode, Chr(&H45))
        ElseIf InStr(Mid(mstrCode, 1, 1), Chr(&H43)) > 0 Then
            lng_S = InStr(mstrCode, Chr(&H43))
        End If
        
        If lng_S > 0 Then
           lng_E = InStr(mstrCode, Chr(&HD) & Chr(&HA))
        End If

    Loop
    strResult = Mid(aResult, 3)    '返回检验结果
    Exit Sub
errH:
    WriteLog "FH400", "FH", Err.Number & " : " & Err.Description
End Sub





