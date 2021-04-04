VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_TBA_40FR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    
    Dim strChenal As String, dblValue As Double, iPos As Long
    
    Dim dbl总蛋白 As Double, dbl白蛋白 As Double, dbl低密度脂蛋白 As Double, dbl高密度脂蛋白 As Double
    Dim dbl载脂蛋白A As Double, dbl载脂蛋白B As Double, dbl总胆红素 As Double, dbl直接胆红素 As Double
    Dim blnEnd As Boolean, strLine As String, str日期 As String, str标本号 As String
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            
            strLine = GetStr_Section(strReserved, Chr(2), Chr(3))
            strReserved = Mid(strReserved, InStr(strReserved, Chr(3)) + 1)
    
            iStrPoint = 1: blnReadItem = False: str日期 = "": str标本号 = ""
            Do While iStrPoint <= Len(strLine)
                If Not blnReadItem Then '查找一组数据的开始位置
                    iStrPoint = InStr(iStrPoint, strLine, "")
                    '判断一组结果是否开始
                    If iStrPoint > 0 Then
                        '记录一组结果的入口
                        iStrPoint = iStrPoint + 1
                        iCurrGrpEntry = iStrPoint - 1
                        '数据是否完整
                        If InStr(iStrPoint, strLine, "") > 0 Then
                            '检验时间
                            '时间格式必须为：yyyy-mm-dd hh:mm:ss
                            iStrPoint = iStrPoint + 1
                            strData = Format(Year(Date) & "-" & Format(ReadSubString(strLine, iStrPoint, 4), "00-00"), "yyyy-mm-dd hh:mm:ss")
                            str日期 = strData
                            '样本序号，必须为数字
                            iStrPoint = iStrPoint + 9
                            
                            str标本号 = Val(ReadSubString(strLine, iStrPoint, 5))
                            strData = strData & "|" & str标本号
                            
                            '检验人
                            strData = strData & "| " '本仪器不支持
                            '标本
                            strData = strData & "|血液" '本仪器不支持
                            '质控品
                            strData = strData & "|0"
                            
                            '准备接收
                            iStrPoint = iStrPoint + 24
                            blnReadItem = True
                        
                            dbl总蛋白 = 0
                            dbl白蛋白 = 0
                            dbl低密度脂蛋白 = 0
                            dbl高密度脂蛋白 = 0
                            dbl载脂蛋白A = 0
                            dbl载脂蛋白B = 0
                            dbl总胆红素 = 0
                            dbl直接胆红素 = 0
                            
                        End If
                    Else
                        Exit Do
                    End If
                Else
                    strBuffer = ReadSubString(strLine, iStrPoint, 1)
                    If strBuffer = "" Then '结果结束
                        If dbl总蛋白 <> 0 And dbl白蛋白 <> 0 Then
                            '球蛋白 总蛋白 - 白蛋白
                            strData = strData & "|50"
                            strData = strData & "|" & Format(Round((dbl总蛋白 - dbl白蛋白), 2), "#0.00")
                            'A/G 白蛋白/球蛋白
                            strData = strData & "|51"
                            strData = strData & "|" & Format(Round((dbl白蛋白 / IIf(dbl总蛋白 - dbl白蛋白 = 0, 1, dbl总蛋白 - dbl白蛋白)), 2), "#0.00")
                        End If
                        '间接胆红素 总 - 直接
                        If dbl总胆红素 <> 0 And dbl直接胆红素 <> 0 Then
                            strData = strData & "|52"
                            strData = strData & "|" & Format(Round((dbl总胆红素 - dbl直接胆红素), 2), "#0.00")
                        End If
                        '动脉粥样硬化指数 低密度脂蛋白 / 高密度脂蛋白
                        If dbl高密度脂蛋白 <> 0 And dbl低密度脂蛋白 <> 0 Then
                            strData = strData & "|53"
                            strData = strData & "|" & Format(Round((dbl低密度脂蛋白 / IIf(dbl高密度脂蛋白 = 0, 1, dbl高密度脂蛋白)), 2), "#0.00")
                        End If
                        '载脂蛋白A/B 载脂蛋白A-1/载脂蛋白B
                        If dbl载脂蛋白B <> 0 And dbl载脂蛋白A <> 0 Then
                            strData = strData & "|54"
                            strData = strData & "|" & Format(Round((dbl载脂蛋白A / IIf(dbl载脂蛋白B = 0, 1, dbl载脂蛋白B)), 2), "#0.00")
                        End If
                        
                        'strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                        
                        blnReadItem = False
                        '该组结果已解析
                        iCurrGrpEntry = 0
                    Else
                        '检验项目
                        iStrPoint = iStrPoint - 1
                        strChenal = ReadSubString(strLine, iStrPoint, 2)
                        strData = strData & "|" & strChenal
                        dblValue = Round(Val(ReadSubString(strLine, iStrPoint, 7)), 2)
                        strData = strData & "|" & Format(dblValue, "#0.00")
                        Select Case strChenal
                            Case "01"
                                dbl总蛋白 = dblValue
                            Case "02"
                                dbl白蛋白 = dblValue
                            Case "16"
                                dbl低密度脂蛋白 = dblValue
                            Case "15"
                                dbl高密度脂蛋白 = dblValue
                            Case "20"
                                dbl载脂蛋白A = dblValue
                            Case "21"
                                dbl载脂蛋白B = dblValue
                            Case "03"
                                dbl总胆红素 = dblValue
                            Case "04"
                                dbl直接胆红素 = dblValue
                        End Select
                    End If
                End If
            Loop
                    
'            '返回未解析的原始串
'            If iCurrGrpEntry > 0 Then
'                strReserved = Mid(strSource, iCurrGrpEntry)
'            Else
'                '向设备发送‘数据收到’命令
'                strCmd = ""
'            End If
            'strResult = strResult & "||" & strData
            If IsDate(str日期) And Val(str标本号) > 0 Then
                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
            End If
            If Not strReserved Like "*" & Chr(&H2) & "*" & Chr(&H3) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
        strReserved = strSource
        'strCmd = Chr(&H6&)
    End If
    
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If

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
        If Mid(strSource, iPoint, 1) = Chr(10) Then
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
