VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_URIT_180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:URIT_180 尿液分析
'厂商:优利特电子集团有限公司
'作者:敬秦
'用户:宜宾市第四人民医院
'日期：2014-12-12

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
    Dim strCode As String, strValue() As String, strLine As String, strData As String
    Dim lngS As Long, lngE As Long
    Dim intLoop As Integer
    Dim strSample As String, strDate As String, strItems As String
    
On Error GoTo hErr
    If strSource <> "" Then strReserved = strReserved & strSource
    Call writelislog("收到：" & strReserved)
    lngS = InStr(strReserved, Chr(&H2))
    Call writelislog("开始符：" & lngS)
    lngE = InStr(strReserved, Chr(&H3))
    Call writelislog("结束符：" & lngE)
    If lngS > 0 And lngE > 0 Then
        strCode = strReserved
        Do Until strCode = ""
            strData = Mid$(strCode, lngS, lngE - 1)
            strCode = Mid(strCode, lngE + 1)
            strValue = Split(strData, vbCrLf)
            For intLoop = LBound(strValue) To UBound(strValue)
                strLine = strValue(intLoop)
                If InStr(strLine, "NO") Then '获取时间
                    strSample = Mid(strLine, InStr(strLine, ".") + 1, 6)
                    strDate = Format(Mid(strLine, 12) & " " & Format(Now, "hh:mm:ss"), "yyyy-mm-dd HH:mm:ss")
                    
                End If
                If intLoop >= 4 And strLine <> "" Then '获取检验数据
                    strLine = Replace(strLine, "mmol/L", "")
                    strLine = Replace(strLine, "Cell/uL", "")
                    strLine = Replace(strLine, "g/L", "")
                    strLine = Replace(strLine, "umol/L", "")
'                    strResult = strDate & "|" & Val(strSample) & "| |尿液|0"
                    strItems = strItems & "|" & Trim(Mid(strLine, 2, 4)) & "|" & Trim(Mid(strLine, 6))
                End If
            Next
            If strSample <> "" And strItems <> "" Then
                strResult = strResult & "||" & strDate & "|" & Val(strSample) & "| |尿液|0" & strItems
            End If
        Loop
        If strResult <> "" Then
            strResult = Mid(strResult, 3)
'            strResult = Replace(strResult, "Norma", "Normal")
            Call writelislog("strResult：" & strResult)
        End If
    End If
    Exit Sub
hErr:
    MsgBox "zlCDLISDev.clsLISdev_URIT_180.Analyse" & Err.Description, vbInformation, "接口错误"
    Err.Clear
End Sub

Private Function writelislog(strLog As String)
    Const strFile As String = "C:\LisLog_URIT180_"
    Dim OBJFILE As New Scripting.FileSystemObject
    Dim objStream As textStream
    Dim STRFILENAME As String
    STRFILENAME = strFile & Format(Date, "yyyy-mm-dd") & ".txt"
    If Not OBJFILE.FileExists(STRFILENAME) Then Call OBJFILE.CreateTextFile(STRFILENAME)
    Set objStream = OBJFILE.OpenTextFile(STRFILENAME, ForAppending)
    objStream.WriteLine (Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & strLog)
    objStream.Close
    Set objStream = Nothing
End Function
