VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_H100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:H-100
'厂商:长春瑞迪
'作者:陈东
'用户:泽州县医院
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
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
    Dim strCode As String, varTmp As Variant, iRow As Integer, strLine As String
    Dim strDate As String, strValue As String, strItem As String, strSampleNo As String
    
    On Error GoTo errH
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    strCode = PStr_CutCode(mstrCode, Chr(&H2), Chr(&H3))
    Do While strCode <> ""
        varTmp = Split(strCode, Chr(&HD) & Chr(&HA))
        strItem = "": strDate = "": strSampleNo = ""
        For iRow = LBound(varTmp) To UBound(varTmp)
            strLine = varTmp(iRow)
            If Mid(strLine, 2) Like "Date:####-##-##*" Then
                strDate = Mid(strLine, 7)
            ElseIf Mid(strLine, 2) Like "Seq:*" Then
                strSampleNo = Trim(Replace(Mid(strLine, 6), "-", ""))
            ElseIf InStr(",UBG,BIL,KET,BLD,PRO,NIT,LEU,GLU,SG,pH,VC,", "," & Trim(Mid(strLine, 2, 3)) & ",") > 0 Then
                strValue = Trim(Mid(strLine, 5))
                strValue = Replace(strValue, "umol/L", "")
                strValue = Replace(strValue, "mmol/L", "")
                strValue = Replace(strValue, "cells/uL", "")
                strValue = Trim(Replace(strValue, "g/L", ""))
                If strValue <> "" And Trim(Mid(strLine, 2, 3)) <> "" Then
                    strItem = strItem & "|" & Trim(Mid(strLine, 2, 3)) & "|" & strValue
                End If
            End If
        Next
        If strDate <> "" And strSampleNo <> "" And strItem <> "" Then
            strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Boold|0" & strItem
        End If
        strCode = PStr_CutCode(mstrCode, Chr(&H2), Chr(&H3))
    Loop
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    Exit Sub
errH:
    WriteLog "H100", "", Err.Description
End Sub

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


Private Function MidUni(ByVal strTemp As String, ByVal Start As Long, Optional ByVal Length As Long) As String
'功能：按数据库规则得到字符串的子集，也就是汉字按两个字符算，而字母仍是一个
    If Length = 0 Then Length = LenB(strTemp)
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
    '去掉可能出现的半个字符
    MidUni = Replace(MidUni, Chr(0), "")
End Function


