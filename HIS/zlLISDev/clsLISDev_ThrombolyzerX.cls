VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ThrombolyzerX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血凝仪BE ThrombolyzerX
'厂商:
'作者:lwf
'用户:扶沟永善妇科医院

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
    Dim arrstr As Variant
    Dim arr项目 As Variant              '一组内容
    Dim ii As Long, jj As Long

    Dim str_日期 As String, str_标本号 As String, str_项目 As String

    mstrCode = mstrCode & strSource
    On Error GoTo errH
    strResult = ""
    If (mstrCode Like "*" & Chr(&H1E) & "*" & Chr(&HD) & Chr(&HA) & "*") Then
        arrstr = Split(mstrCode, Chr(&HD) & Chr(&HA))
        For ii = 0 To UBound(arrstr) Step 1
            If Len(arrstr(ii)) > 50 Then
                arr项目 = Split(arrstr(ii), "|")
                str_日期 = "": str_标本号 = "": str_项目 = ""
                str_日期 = Format(Now(), "yyyy-mm-dd")
                str_标本号 = Val(arr项目(1))
                If str_标本号 <> "" Then
                    str_项目 = str_项目 & "|" & Trim(arr项目(2)) & "|" & Trim(arr项目(4))
                    If Trim(arr项目(2)) = 0 Then
                        str_项目 = str_项目 & "|1001|" & Trim(arr项目(5))
                    End If
                    If Trim(arr项目(2)) = 3 Then
                        str_项目 = str_项目 & "|1002|" & Trim(arr项目(5))
                    End If
                End If

                If Val(str_标本号) <> 0 And IsDate(str_日期) And str_项目 <> "" Then
                    strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血液|0" & str_项目
                End If
                mstrCode = ""
            Else
                mstrCode = arrstr(ii)
            End If
        Next
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
