VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CLIMTEKSTATUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:CLIMTEKSTATUS 尿液分析仪
'厂商:
'作者:陈东
'用户:陈辉,青羊妇幼尿机
'2011-05-11
Dim mstrCode As String

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
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant
    Dim str_Line As String, i As Integer
    Dim str_Date As String, str_SampleNo As String, str_Item As String
    mstrCode = mstrCode & strSource
    var_Code = Split(mstrCode, ",")
    If UBound(var_Code) >= 41 Then
        str_SampleNo = var_Code(0)
        str_Date = var_Code(1)
        str_Date = Mid(str_Date, 7, 4) & "-" & Mid(str_Date, 4, 2) & "-" & Mid(str_Date, 1, 2) & " " & var_Code(2)
        str_Date = Format(CDate(str_Date), "yyyy-MM-dd HH:mm")
        For i = 9 To 41 Step 3
            If Trim(var_Code(i)) <> "" And Trim(var_Code(i + 1)) <> "" Then
                str_Item = str_Item & "|" & Trim(var_Code(i)) & "|" & Trim(var_Code(i + 1))
            End If
        Next
        If Val(str_SampleNo) > 0 And IsDate(str_Date) And Trim(str_Item) <> "" Then
            strResult = strResult & "||" & str_Date & "|" & str_SampleNo & "| |尿液|0" & Replace(str_Item, "Trace", "微量")
        End If
        str_Code = ""
        For i = 42 To UBound(var_Code)
            str_Code = str_Code & "," & var_Code(i)
            
        Next
        mstrCode = str_Code
    End If
        
    If strResult <> "" Then strResult = Mid(strResult, 3)
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

