VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_JuniorII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:朱丽亚尿机 Junior II
'厂商:
'作者:陈东
'用户:江陵医院
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
    
    Dim iStrPoint As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim varIn As Variant '存放输入数据
    Dim StrTime As String
    
    Dim lng_S As Long, lng_E As Long
    Dim strLine As String
    Dim str标本号 As String, str日期 As String, str检验人标本质控 As String, str指标 As String
    Dim str格式 As String, strItem As String
    Dim strCode As String
    On Error GoTo errHandle
    str检验人标本质控 = " |尿液|0"
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource = "" Then Exit Sub
    mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    If Not (lng_S > 0 And lng_E > 0) Then
        lng_S = InStr(mstrCode, "URINALYSIS")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&HC))
        Else
            lng_E = InStr(mstrCode, Chr(&HC))
        End If
    End If
    
    If lng_S > 0 And lng_E > 0 Then
        strCode = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        varIn = Split(strCode, vbNewLine)
        strLine = ""
        For iStrPoint = LBound(varIn) To UBound(varIn)
        
            strLine = varIn(iStrPoint)
            
            '--- 标本号"               Seq.no. 12"
            If InStr(strLine, "Seq.no.") > 0 Then
                str标本号 = Val(Trim(Mid(strLine, InStr(strLine, ".no.") + 4)))
                If str标本号 = "0" Then Exit Sub
            End If
            
            '--- 日期  "07.03.20           10:54"
            
            str格式 = Trim(ReadIni("SET", "日期格式", App.Path & "\JuniorII.ini"))
            If str格式 = "" Then
                WriteIni "SET", "日期格式", "年.月.日", App.Path & "\JuniorII.ini"
                str格式 = "年.月.日"
            End If
            
            If strLine Like "*.*.*:*" And InStr(strLine, "Seq.no.") <= 0 Then
                strLine = Mid(strLine, 1, 10)
                If str格式 = "年.月.日" Then
                    str日期 = Format(CDate(Split(strLine, ".")(0) & "-" & Split(strLine, ".")(1) & "-" & Split(strLine, ".")(2)), "yyyy-MM-dd")
                ElseIf str格式 = "月.日.年" Then
                    str日期 = Format(CDate(Split(strLine, ".")(2) & "-" & Split(strLine, ".")(0) & "-" & Split(strLine, ".")(1)), "yyyy-MM-dd")
                ElseIf str格式 = "日.月.年" Then
                    str日期 = Format(CDate(Split(strLine, ".")(2) & "-" & Split(strLine, ".")(1) & "-" & Split(strLine, ".")(0)), "yyyy-MM-dd")
                End If
            ElseIf strLine Like "*Date:*.*.*" Then
                strLine = Split(strLine, "Date:")(1)
                If str格式 = "年.月.日" Then
                    str日期 = Format(CDate(Split(strLine, ".")(0) & "-" & Split(strLine, ".")(1) & "-" & Split(strLine, ".")(2)), "yyyy-MM-dd")
                ElseIf str格式 = "月.日.年" Then
                    str日期 = Format(CDate(Split(strLine, ".")(2) & "-" & Split(strLine, ".")(0) & "-" & Split(strLine, ".")(1)), "yyyy-MM-dd")
                ElseIf str格式 = "日.月.年" Then
                    str日期 = Format(CDate(Split(strLine, ".")(2) & "-" & Split(strLine, ".")(1) & "-" & Split(strLine, ".")(0)), "yyyy-MM-dd")
                End If
            End If
            
            '-- 颜色
            If InStr(strLine, "Color:") > 0 And strResult <> "" Then
                strResult = strResult & "|Color|" & Trim(Mid(strLine, 11))
                Exit Sub
            End If
            
            '-- 指标 "  SG :       1.015 "
            str指标 = ""
            If InStr(strLine, ":") = 6 Then
                str指标 = strLine
            End If
            
            If strResult = "" And str标本号 <> "" And str日期 <> "" And str指标 = "" Then
                strResult = str日期 & "|" & str标本号 & "|" & str检验人标本质控
            ElseIf strResult <> "" And str标本号 <> "" And str日期 <> "" And str指标 <> "" Then
                strItem = IIf(Trim(Mid(str指标, 7, 12)) = "", "－－", Trim(Mid(str指标, 7, 12)))
                strItem = Replace(strItem, "/ul", "")
                strItem = Replace(strItem, "mg/dl", "")
                strResult = strResult & "|" & Trim(Mid(str指标, 2, 4)) & "|" & Trim(strItem)
            End If
        Next
    End If
    
    Exit Sub
errHandle:
    WriteLog "JuniorII", Err.Description, ""

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


