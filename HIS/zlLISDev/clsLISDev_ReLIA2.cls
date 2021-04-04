VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ReLIA2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器类别：生化类
'设备厂商：未知
'设备型号：RELIA
'使用医院:成都慢性病医院

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
    Dim str_S As Long, str_E As Long, str_SS As Long
    Dim str_Code As String, str_Line As String
    Dim VAR_DAT As Variant
    Dim i As Integer, j As Integer
    Dim str日期 As String, str标本号 As String, str项目 As String, str结果 As String, str检验项目 As String, str检验结果 As String
    
    str日期 = "": str_S = 0: str_E = 0
    str标本号 = "": str项目 = "": str结果 = ""
    mstrCode = strSource
    '判断字符串的完整性
    str_S = InStr(mstrCode, "BEGIN REPORT")

    If str_S > 0 Then
       str_E = InStr(mstrCode, "END REPORT")
    Else
       str_E = 0
    End If

    If str_S > 0 And str_E > 0 Then
        Do While str_S > 0 And str_E > 0  '判断一个数据中有多个开始和结束的开始
             str_Code = Mid(mstrCode, str_S, str_E - str_S + 12)
             mstrCode = Mid(mstrCode, str_E + 15)
              '如果数据中含有完整的标本就开始进行数据的解析
             VAR_DAT = Split(str_Code, Chr(&HD) + Chr(&HA))
             For i = LBound(VAR_DAT) To UBound(VAR_DAT)
                 str_Line = VAR_DAT(i)
                 If str_Line Like "*Specimen ID:*" Then
                    str标本号 = Val(Mid(str_Line, InStr(str_Line, ":") + 1))
                ElseIf str_Line Like "*Test finish time:*" Then
                    str日期 = Trim(Mid(str_Line, InStr(str_Line, ":") + 1))
                    str日期 = Format(CDate(str日期), "yyyy-MM-dd HH:MM") & ":" & Format(Now, "SS")
                    If Not IsDate(str日期) Then str日期 = Format(Now, "yyyy-MM-dd HH:MM:SS")
                 ElseIf str_Line Like "*Test name:*" Then
                    str检验项目 = Trim(Mid(str_Line, InStr(str_Line, ":") + 1))
                 ElseIf InStr(str_Line, str标本号) > 0 And InStr(str_Line, str检验项目) And InStr(str_Line, ":") > 0 And str检验项目 <> "" And str标本号 <> "" Then
                    str检验结果 = Trim(Replace(Replace(Replace(Replace(Mid(str_Line, InStr(str_Line, ":") + 1), "ng/ml", ""), "pg/ml", ""), "↑", ""), "↓", ""))
                 End If
             Next
             If str标本号 <> "" And str日期 <> "" And str检验项目 <> "" And str检验结果 <> "" Then
                str结果 = IIf(Trim(str结果) = "", "", str结果 & "||") & IIf(Trim(str检验项目 & "|" & str检验结果) = "", "", str日期 & "|" & str标本号 & "| |血清|0|" & str检验项目 & "|" & str检验结果)
                str日期 = ""
                str标本号 = ""
                str检验项目 = ""
                str检验结果 = ""
             End If
        '循环条件判断
             str_S = InStr(mstrCode, "BEGIN REPORT")
             If str_S > 0 Then
                str_E = InStr(mstrCode, "END REPORT")
             Else
                str_E = 0
             End If
            
        Loop
    End If
    strResult = str结果
    strReserved = mstrCode
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


