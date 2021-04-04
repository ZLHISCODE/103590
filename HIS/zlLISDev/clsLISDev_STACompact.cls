VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_STACompact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:STACompact－血凝仪－STAGO
'厂商:法国 STAGO
'作者:陈东
'用户:重庆中山医院,海南省人民医院

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
    
    Dim lngS As Long, lngE As Long, strCode As String, varTmp As Variant, iRow As Integer, strLine As String
    Dim strDate As String, strSampleNo As String, strItem As String, strValue As String, strItems As String
    On Error GoTo hErr
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    End If
    mstrCode = mstrCode & strSource
    lngE = 0: lngS = InStr(mstrCode, Chr(&H5))
    If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H4))
    
    Do While lngS > 0 And lngE > 0
        strCode = Mid(mstrCode, lngS, lngE - lngS + 1)
        mstrCode = Mid(mstrCode, lngE + 1)
        varTmp = Split(strCode, Chr(&HD) & Chr(&HA) & Chr(&H2))
        
        strDate = "": strSampleNo = "": strItems = ""
        
        For iRow = LBound(varTmp) To UBound(varTmp)
            strLine = Trim$("" & varTmp(iRow))
            If Left(strLine, 5) = Chr(&H5) & Chr(&H2) & "1H|" Then
                strDate = Split(strLine, "|")(13)
                If Len(strDate) >= 14 Then
                    strDate = Mid$(strDate, 1, 4) & "-" & Mid$(strDate, 5, 2) & "-" & Mid$(strDate, 7, 2) & " " & Mid$(strDate, 9, 2) & ":" & Mid$(strDate, 11, 2) & ":" & Mid$(strDate, 13, 2)
                End If
                If Not (strDate <> "" And IsDate(strDate)) Then strDate = ""
            ElseIf strLine Like "?O|1|*|R*" Then
                strSampleNo = Trim$("" & Split(strLine, "|")(2))
            ElseIf strLine Like "?O|1|*|S*" Then
                strSampleNo = Trim$("" & Split(strLine, "|")(2))
                If strSampleNo <> "" Then strSampleNo = strSampleNo & "^1^"
            ElseIf strLine Like "?R|*|^^^*" Then
                strItem = Trim(Replace$(Split(strLine, "|")(2), "^^^", ""))
                strValue = Trim(Split(strLine, "|")(3))
                If strItem <> "" And strValue <> "" Then
                    strItems = strItems & "|" & strItem & "|" & strValue
                End If
            End If
        Next
        
        If strDate <> "" And strSampleNo <> "" And strItems <> "" Then
            strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |血清|0" & strItems
        End If
        
        lngE = 0: lngS = InStr(mstrCode, Chr(&H5))
        If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H4))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
hErr:
    WriteLog "STACompact", "第" & CStr(Erl()) & "行", Err.Description
    'Resume
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



