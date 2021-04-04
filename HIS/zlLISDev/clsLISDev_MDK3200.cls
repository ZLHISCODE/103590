VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MDK3200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'产品名称: 血流变MDK3200
'产品型号: MDK3200
'作者:李建飞
'日期:2010-09-15
'用户:核工业四一六医院——陈辉
'注意事项：该接口必须配合总公司 ReadTxtSend 程序使用

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

    Dim str_Code As String '1个完整的消息
    Dim var_Code As Variant
    Dim var_Line As Variant
    Dim i As Integer, j As Integer
    Dim var_Filed As Variant   '字段
    Dim lng_S As Long, lng_E As Long
    Dim str_标本号 As String, str_日期 As String, str_标本类型 As String
    Dim str_质控 As String, str_检验项目 As String
    Dim str_Line As String, str_Data As String, str_Tmp As String
    
    On Error GoTo errHandle

    mstrCode = mstrCode & strSource
 
    lng_S = InStr(mstrCode, Chr(&H2))
    lng_E = InStr(mstrCode, Chr(&H3))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S + 1, lng_E - 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
        
        var_Code = Split(str_Code, "  " & Chr(&HD) & Chr(&HA))
        For i = LBound(var_Code) To UBound(var_Code)
            str_Data = Replace(Replace(var_Code(i), "[", ""), "]", "")
            If str_Data Like "Base*" Then
                var_Line = Split(str_Data, Chr(&HD) & Chr(&HA))
                
                For j = LBound(var_Line) + 1 To UBound(var_Line) - 1
                    str_Line = var_Line(j)
                    
                    var_Filed = Split(str_Line, "=")
                    
                    If var_Filed(0) = "TESTDATE" Then
                        str_日期 = Format(var_Filed(1), "yyyy-MM-dd")
                    ElseIf var_Filed(0) = "SAMPLENO" Then
                        str_标本号 = Val(var_Filed(1))
                    End If
                    
                Next
            ElseIf str_Data Like "DATA*" Then
                var_Line = Split(str_Data, Chr(&HD) & Chr(&HA))
                
                For j = LBound(var_Line) + 1 To UBound(var_Line) - 1
                    str_Line = var_Line(j)
                    
                    var_Filed = Split(str_Line, "=")
                    
                    If var_Filed(0) = "RESULT" Then
                        str_Tmp = Val(var_Filed(1))
                    ElseIf var_Filed(0) = "NO" Then
                        str_Tmp = Val(var_Filed(1)) & "|" & str_Tmp
                    End If
                    
                Next
                str_检验项目 = str_检验项目 & "|" & str_Tmp
            ElseIf str_Data Like "GRAPH*" Then
                
            End If
        Next
        
        str_标本类型 = "血清"
        
        If Val(str_标本号) > 0 And str_检验项目 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |" & str_标本类型 & "|0" & str_检验项目
        End If
        
        str_检验项目 = ""
        lng_S = InStr(mstrCode, Chr(&H2))
        lng_E = InStr(mstrCode, Chr(&H3))
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "MDK3200_analyse", strSource, Err.Number & " " & Err.Description
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
