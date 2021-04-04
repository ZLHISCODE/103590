VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_VITEK32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  微生物 VITEK32
'医院:  烟台山医院
'作者:  陈东
'时间:  2008/08/19

Dim mstrCode  As String

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

    
    Dim aData() As String
    Dim intloop As Integer
    Dim strDate As String               '日期
    Dim strSampleNO As String           '标本号
    Dim strLine As String               '当前要解析的字串
    Dim lngBegin As Long                '开始位置
    Dim lngEnd As String                '结束位置
    
    Dim str细菌 As String, str药敏 As String
    
    On Error GoTo errH
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    '没有找到开始位置时退出
    
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H3)) > 0 Then
        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        strCmd = Chr(&H6)
    End If
    
    '没有找到结束位置时退出
    lngBegin = InStr(mstrCode, Chr(&H5))
    lngEnd = InStr(mstrCode, Chr(&H4))
     
    Do While lngBegin > 0 And lngEnd > 0
    
        '取当前一次的结果，其他的返回
        strLine = Mid(mstrCode, lngBegin, lngEnd - lngBegin)
        mstrCode = Mid(mstrCode, lngEnd + 1)
        
        '开始解析
        aData = Split(strLine, "|")
        '日期
        strDate = Mid(aData(1), 3, 4) & "-" & Mid(aData(1), 7, 2) & "-" & Mid(aData(1), 9, 2)
        '标本号
        strSampleNO = Mid(aData(3), 3)
        If Val(strSampleNO) > 0 Then
            str细菌 = "": str药敏 = ""
            
            '开始解析结果
            For intloop = 4 To UBound(aData)
                If Left(aData(intloop), 2) = "a1" Then
                    '因返回数据中无药敏方法信息,所以均返回为   MIC
                    str药敏 = str药敏 & "|" & Mid(aData(intloop), 3) & "|" & Mid(aData(intloop + 3), 3) & "^1^" & Mid(aData(intloop + 2), 3)
                ElseIf Left(aData(intloop), 2) = "o1" Then
                    str细菌 = Mid(aData(intloop), 3)
                End If
            Next
            If str细菌 <> "" And str药敏 <> "" Then
                strResult = strResult & "||" & strDate & "|" & Val(strSampleNO) & "| |" & str细菌 & "|0" & str药敏
            End If
        End If
        lngBegin = InStr(mstrCode, Chr(&H5))
        lngEnd = InStr(mstrCode, Chr(&H4))
    Loop
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





