VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MI921"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:MI-921
'厂商:越华
'作者:曾超
'用户:南京江北人民医院
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
    Dim i As Integer
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

    Dim iStrPoint As Integer, intSpot As Integer
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim iPos As Long, strTmp As String, dblValue As Double, strItemValue As String
    Dim strBarcode As String
    Dim strLine As String
    Dim intStart As Integer, intEnd As Integer
    Dim adata() As String
    Dim intloop As Integer
    Dim strItem As String
    
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    adata = Split(strSource, vbCrLf)
    
    For intloop = 0 To UBound(adata)
        strItem = adata(intloop)
        If Len(strItem) >= 49 Then
            If Val(Mid(strItem, 1, 3)) > 0 Then
                strData = strData & "||" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "|" & Val(Mid(strItem, 1, 3)) & "| |血清|0" & _
                       "|3|" & Val(Mid(strItem, 18, 4)) & "|4|" & Val(Mid(strItem, 23, 5)) & _
                       "|5|" & Val(Mid(strItem, 29, 5)) & "|6|" & Val(Mid(strItem, 35, 4)) & _
                       "|7|" & Val(Mid(strItem, 40, 4)) & "|8|" & Val(Mid(strItem, 46, 4)) & _
                       "|9|" & Val(Mid(strItem, 35, 4)) * 1.95 & "|10|" & Val(Mid(strItem, 23, 5)) - Val(Mid(strItem, 29, 5)) - Val(Mid(strItem, 46, 4))
            End If
        Else
            strReserved = strReserved & strItem
        End If
'        If UBound(strItem) >= 11 Then
'            strData = strData & "||" & Now & "|" & strItem(0) & "| |血清|0"
'            strData = strData & "|1|" & Val(strItem(2)) & "|2|" & Val(strItem(4)) & "|3|" & Val(strItem(5)) & _
'                      "|4|" & Val(strItem(6)) & "|5|" & Val(strItem(7)) & "|6|" & Val(strItem(8)) & _
'                      "|7|" & Val(strItem(9)) & "|8|" & Val(strItem(11))
'        End If
    Next
    
    strResult = Mid(strData, 3)
    strCmd = Chr(&H6&)
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

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    '指针定位到下一个换行之后
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = "]" Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '读取子串，同时指针移动到读取的字符串之后
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function




