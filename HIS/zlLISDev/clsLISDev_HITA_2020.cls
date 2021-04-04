VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA_2020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:日立 2020
'厂商:日立
'作者:曾超
'用户:吉林

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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, strimg As String
    Dim lngS As Long, lngE As Long
    Dim strNow As String
    Dim aLine() As String, aItem() As String
    Dim intLoop As Integer
    Dim strData As String
    
    
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
'    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    If InStr(strSource, Chr(&H5)) > 0 Or InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    End If
    
    mstrCode = mstrCode & strSource
    
    '没有开始符时退出
    If InStr(mstrCode, Chr(&H5)) = 0 Then
        Exit Sub
    End If
    
    '没有结束符时退出
    If InStr(mstrCode, Chr(&H4)) = 0 Then
        Exit Sub
    End If
    
    strNow = Mid(mstrCode, InStr(mstrCode, Chr(&H5)))
    strNow = Mid(mstrCode, 1, InStr(mstrCode, Chr(&H4)))
    
    mstrCode = Mid(mstrCode, InStr(mstrCode, Chr(&H4)) + 1)
    
    '开始解析
    aLine = Split(strNow, vbCrLf)
    
    For intLoop = 0 To UBound(aLine)
        If InStr(aLine(intLoop), "|") > 0 Then
            aItem = Split(aLine(intLoop), "|")
            strTmp = aItem(0)
            strTmp = Replace(Replace(strTmp, vbCr, ""), vbLf, "")
            If Right(strTmp, 1) = "O" Then
                '标本号和时间
                strData = Mid(aItem(6), 1, 4) & "-" & Mid(aItem(6), 5, 2) & "-" & Mid(aItem(6), 7, 2) & " " & _
                                Mid(aItem(6), 9, 2) & ":" & Mid(aItem(6), 11, 2) & ":" & Mid(aItem(6), 13, 2) & _
                          "|" & aItem(2) & "| |血液|0"
            End If
            
            If Right(strTmp, 1) = "R" Then
                '结果
                strData = strData & "|" & Split(aItem(2), "^")(3) & "|" & aItem(3)
            End If
        End If
    Next
    strData = Replace(strData, "||", "| |")
    strResult = strData
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
    
End Function




