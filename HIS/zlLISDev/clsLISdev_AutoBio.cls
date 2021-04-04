VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_AutoBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: 安图酶标仪
'厂商:安图
'作者:陈东
'用户:本钢总医院
'2012-10-31
Private mstrCode As String
 
Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    'GetStartCmd = Chr(6)
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
    Dim lngS As Long, lngE As Long, strCode As String
    Dim varData As Variant, strLine As String, iRow As Integer
    Dim strSampleNo As String, strDate As String, strChannel As String, strValue As String
    Dim varTmp As Variant
    On Error GoTo hErr
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    lngS = InStr(mstrCode, Chr(&H5))
    lngE = 0
    lngE = InStr(lngS, mstrCode, Chr(&H4))
    
    strResult = ""
    
    Do While lngS > 0 And lngE > 0
        strCode = Mid$(mstrCode, lngS, lngE - lngS + 1)
        mstrCode = Mid$(mstrCode, lngE + 1)
        '-----------------------------------------------------------------
        varData = Split(strCode, Chr(&HD) & Chr(&HA))
        For iRow = LBound(varData) To UBound(varData)
            strLine = varData(iRow)
            varTmp = Split(strLine, Chr(&H9))
            
            If UBound(varTmp) >= 3 Then
                strSampleNo = varTmp(0)
                If IsNumeric(strSampleNo) Then
                    strDate = Left(strSampleNo, 6)
                    strDate = "20" & Mid$(strDate, 1, 2) & "-" & Mid$(strDate, 3, 2) & "-" & Mid$(strDate, 5, 2)
                    strSampleNo = Mid(strSampleNo, 7)
                    strChannel = varTmp(1)
                    strChannel = Replace$(strChannel, "列", "")
                    strValue = varTmp(2) ' & "^" & varTmp(14) & "^" & varTmp(16)
                    
                    If IsDate(strDate) And strSampleNo <> "" And strChannel <> "" And strValue <> "" Then
                       strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Boold|0|" & strChannel & "|" & strValue
                    Else
                        Call SaveLog(iRow & "行数据不正确，标本号" & strSampleNo & ",日期" & strDate & "，通道" & strChannel & ",结果" & strValue)
                    End If
                End If
            End If
        Next
        '-----------------------------------------------------------------
        lngS = InStr(mstrCode, Chr(&H5))
        lngE = 0
        lngE = InStr(lngS, mstrCode, Chr(&H4))
    Loop
    If strResult <> "" Then strResult = Mid$(strResult, 3)
    Exit Sub
hErr:
    Call SaveLog("第" & CStr(Erl()) & "行，" & Err.Description)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    
End Function

Private Sub SaveLog(ByVal strOutput As String)
    '------------------------------------------------------
    '--  功能:根据调试标志,写日志到当前目录
    '------------------------------------------------------
    
    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If Dir(App.Path & "\调试.TXT") = "" Then Exit Sub
    strFileName = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"
    
    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    strOutput = Replace(strOutput, Chr(&HD), "<CR>")
    strOutput = Replace(strOutput, Chr(&HA), "<LF>")
    objStream.WriteLine (strDate & ":   " & strOutput)
    objStream.Close
    Set objStream = Nothing
End Sub




