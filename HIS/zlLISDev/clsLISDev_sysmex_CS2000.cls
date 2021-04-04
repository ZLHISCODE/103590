VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_sysmex_CS2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  sysmex CS-2000i
'医院:  南阳中医院
'作者:  lwf
'时间:  2014/11/24

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
    '   微生物结果格式 通道码|结果类型^方式^结果值
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令

    Dim str_Code As String
    Dim str_Line As String               '当前要解析的字串
    Dim lng_S As Long                  '开始位置
    Dim lng_E As String                '结束位置
    Dim varLine As Variant
    Dim ii As Long
    Dim dec_bd As Double
    Dim str_日期 As String, str_标本号 As String, str_项目 As String

    mstrCode = strSource
    On Error GoTo errH
    strResult = ""

    lng_S = InStr(strSource, "")
    lng_E = InStr(lng_S, strSource, "")

    If lng_S > 0 And lng_E > 0 And lng_S < lng_E Then
        varLine = Split(mstrCode, "M")
        str_日期 = "": str_标本号 = "": str_项目 = ""
        str_日期 = Mid(Split(varLine(0), "U")(1), 1, 6)
        str_日期 = "20" & Mid(str_日期, 1, 2) & "-" & Mid(str_日期, 3, 2) & "-" & Mid(str_日期, 5, 2)
        str_标本号 = Val(Right(varLine(0), 4))

        For ii = 0 To 10 Step 1
            If ii = 0 Then
                dec_bd = 0.1
            ElseIf ii = 3 Then
                dec_bd = 0.01
            ElseIf ii = 5 Then
                dec_bd = 0.1
            ElseIf ii = 7 Then
                dec_bd = 0.001
            ElseIf ii = 8 Then
                dec_bd = 0.1
            ElseIf ii = 10 Then
                dec_bd = 0.01
            Else
                dec_bd = 1
            End If
            str_项目 = str_项目 & "|" & Mid(Trim(varLine(1)), ii * 9 + 1, 3) & "|" & Val(Trim(Mid(Trim(varLine(1)), ii * 9 + 4, 6))) * dec_bd
        Next
        If Val(str_标本号) > 0 Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血液|0" & str_项目
        End If
    Else
        strReserved = mstrCode
    End If

    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errH:
    WriteLog "err", Err.Description, strResult
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    
End Function
