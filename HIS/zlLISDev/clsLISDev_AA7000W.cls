VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AA7000W"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'仪器:微量元素分析仪 AA7000W
'厂商:东西
'作者:曾超
'用户:南桐医院
Private mblnStart As Boolean '开始
Private mblnEnd As Boolean   '结束
Private mstrData As String '他待解析串

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
    
    Dim lng_Stx As String, lng_Etx As Long, str_Code As String
    Dim varIn As Variant '存放输入数据
    Dim strLine As String
    Dim str标本号 As String, str日期 As String, str检验人标本质控 As String, str指标 As String, str值 As String
    Dim str要提取指标 As String, var要提取指标 As Variant
    Dim iStrPoint As Integer
    Dim aData() As String, aItem() As String
    Dim intLoop  As Integer, intItem As Integer
    Dim strData As String
    
    
'    aData = Split(strSource, ",")
'    If UBound(aData) < 2 Then
'        strResult = strSource
'    End If
'
'    For intLoop = 0 To UBound(aData)
'        If intItem < 4 Then
'            strData = strData & "," & aData(intLoop)
'            intItem = intItem + 1
'        Else
'            strData = strData & ";" & aData(intLoop)
'            intItem = 1
'        End If
'    Next
    
    If InStr(strSource, ";") <= 0 Then
        strResult = strSource
    End If
    
    aData = Split(strSource, ";")
    For intLoop = 0 To UBound(aData)
        If aData(intLoop) <> "" Then
            aItem = Split(aData(intLoop), ",")
            If UBound(aItem) >= 3 And Len(Trim(aItem(3))) > 0 Then
                strResult = strResult & "||" & Now & "|" & aItem(0) & "| |血液|0|" & aItem(3) & "|" & aItem(4)
            Else
                strReserved = strReserved & "," & aData(intLoop)
            End If
        Else
            strReserved = strReserved & "," & aData(intLoop)
        End If
    Next
    strReserved = Mid(strReserved, 2)
    strResult = Mid(strResult, 3)
    
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




