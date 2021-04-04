VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_diff5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strDeviceString '发送信息的设备串
Private strLastString As String '上次发送的设备串

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
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    Dim aData() As String, iDataPoint As Integer, aItem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim str标本信息 As String
    Dim dbl总蛋白 As Double, dbl白蛋白 As Double, dbl低密度脂蛋白 As Double, dbl高密度脂蛋白 As Double
    Dim dbl载脂蛋白A As Double, dbl载脂蛋白B As Double, dbl总胆红素 As Double, dbl直接胆红素 As Double
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1: blnReadItem = False
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '查找一组数据的开始位置
            iStrPoint = InStr(iStrPoint, UCase(strSource), "RESULT")
            '判断一组结果是否开始
            If iStrPoint > 0 Then
                '记录一组结果的入口
                iCurrGrpEntry = iStrPoint
                iStrPoint = iStrPoint + 6
                '数据是否完整
                If InStr(iStrPoint, UCase(strSource), "ACT5DIFF") > 0 Then
                    aData = Split(Mid(strSource, iStrPoint), vbCr)
                    '数据是否完整
                    If UBound(aData) > 7 Then
                        '检验时间
                        '时间格式必须为：yyyy-mm-dd hh:mm:ss
                        strData = Format(Split(aData(47), " ")(1) & " " & Time, "yyyy-mm-dd hh:mm:ss")
                                
                        '样本序号，必须为数字
                        strData = strData & "|" & Val(Split(aData(50), " ")(1))
                                
                        '检验人
                        strData = strData & "| " '本仪器不支持
                        '标本
                        strData = strData & "|血液" '本仪器不支持
                        '质控品
                        strData = strData & "|0"
                        str标本信息 = strData
                        '准备接收
                        iPos = InStr(iStrPoint, strSource, "!")
                        If iPos > 0 Then
                            aData = Split(Mid(strSource, iPos), vbCr)
                            iDataPoint = 0: iStartPoint = 0
                        Else
                            iDataPoint = 8: iStartPoint = 8
                        End If
                        blnReadItem = True
                    End If
                End If
            Else
                Exit Do
            End If
        Else
            If InStr(UCase(aData(iDataPoint)), "ACT5DIFF") > 0 Then   '结果结束
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                
                blnReadItem = False
                '该组结果已解析
                iCurrGrpEntry = 0
            Else
                '检验项目
                strChenal = Format(iDataPoint - iStartPoint + 1, "00")
                If strChenal > "26" Then '图像数据 ALT分布|0| | | |001001231231223
'                   Select Case Mid(aData(iDataPoint), 1, 1)
'                        Case 19
'                            strData = strData & Chr(13) & Chr(10) & Trim(Mid(aData(iDataPoint), 1, 2)) & "|0| | | |" & to_Hex(Mid(aData(iDataPoint), 3))
'                        Case 20, 21
'                            strData = strData & "||" & str标本信息 & Chr(13) & Chr(10) & Trim(Mid(aData(iDataPoint), 1, 2)) & "|0| | | |" & to_Hex(Mid(aData(iDataPoint), 3))
'                        End Select
                       
                        iCurrGrpEntry = 0
                Else
                    strData = strData & "|" & strChenal
                    Select Case strChenal
                        Case 17, 21
                            dblValue = Val(Mid(aData(iDataPoint), 3)) * 10
                        Case 19
                            dblValue = Val(Mid(aData(iDataPoint), 3)) * 1.1
                        Case Else
                            dblValue = Val(Mid(aData(iDataPoint), 3))
                        End Select
                    strData = strData & "|" & dblValue
                End If
                iDataPoint = iDataPoint + 1
            End If
        End If
    Loop
            
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    Else
        '向设备发送‘数据收到’命令
        strCmd = ""
    End If
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    '解析需要向仪器发送的标本信息
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    blnSuccess = True
    iSendStep = 0
    SendSample = ""
End Function

Private Function GetCheckCode(ByVal strSource As String) As String
'获取校验码
    Dim lngAsc As Long, I As Integer
    
    For I = 1 To Len(strSource)
        lngAsc = lngAsc + Asc(Mid(strSource, I, 1))
    Next
    GetCheckCode = CStr(Hex(lngAsc Mod 256))
End Function

Private Function ReadSubString(ByVal strSource As String, iPoint As Long, Optional ByVal Length As Integer = 0) As String
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

Private Function ReadLine(ByVal strSource, iPoint As Long) As String
    '从指定位置开始读取一行，同时指针定位到下一个换行之后
    Dim iPos As Long
    
    iPos = iPoint
    Do While iPos <= Len(strSource)
        If Mid(strSource, iPos, 1) = Chr(13) Then
            Exit Do
        End If
        iPos = iPos + 1
    Loop
    If iPos > Len(strSource) Then
        ReadLine = Mid(strSource, iPoint)
        iPoint = iPos
    Else
        ReadLine = Mid(strSource, iPoint, iPos - iPoint)
        iPoint = iPos + 1
    End If
End Function

Private Function to_Hex(str记录 As String) As String
Dim I  As Integer
    to_Hex = ""
    For I = 1 To Len(str记录)
            to_Hex = to_Hex & Hex(Asc(Mid(str记录, I, 1)))
    Next
End Function












