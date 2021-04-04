VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MK3_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const LABITEMS As String = ",pH,PCO2,PO2,pHt,PCO2t,PO2t,ctHb,O2Hb,HHb,COHb,MetHb,SulfHb,pHst,cHCO3st,cHCO3,ctCO2(P),ctCO2(B),BE,BEecf,BB,SO2,PAO2t,AaDO2t,a/AO2t,RI,"
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
    '补充图像的方式：
    '                   1.图像数据跟随指标数据后，使用回车换行符来分隔。
    '                   2.有多个图像数据时使用"^"来分隔
    '                   3.单个图像数据格式: 图像画法 0=直方图  1=散点图
    '                     a) 直方图: 图像名称;图像画法(0=直方图  1=散点图);X1;X2;X3;X4;X5...
    '                     b) 散点图: 图像名称;图像画法(0=直方图  1=散点图):
    '                        例:00000100001000010000100010;00000100001000010000100010;
    '                        说明:1.散点图以点阵方式保存每一行使用分号来分隔.
    '                             2.有多少个分号就有多少行
    '                             3.每一行有多少个点由每一行的长度来确定
    '                             3.画图的方向是从最上边向下画，如有65*65的图就是从65行开始画(最上边开始画)
    
    Dim iStrPoint As Integer, iStrPointBK As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    Dim aData() As String, iDataPoint As Integer, aItem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngloop As Long, intloop As Integer
    Dim strDate As String
    Dim i As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '标本号
    Dim strBarcode As String            '条码号
    Dim strImageData As String          '图像信息
    Dim strLine As String               '当前要解析的字串
    Dim lngBegin As Long                '开始位置
    Dim lngEnd As String                '结束位置
    
    '没有这么长时退出
    If UBound(Split(strSource, vbCrLf)) < 8 Then
        strReserved = strSource
        Exit Sub
    End If
    
    '开始解析
    strResult = strSource
    For intloop = 5 To 2 Step -1
        strResult = Replace(strResult, Space(intloop), Space(1))
    Next
    strResult = Replace(strResult, "OK" & vbCrLf, "")
    strResult = Replace(strResult, " ", ";")
    strResult = Replace(strResult, "-", ";-")
    strResult = Mid(strResult, 2)
    strResult = Replace(strResult, vbCrLf, "|")
    strResult = Replace(strResult, "|;", "|")
    
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

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '读取一行字符，并将字符指针移到下一行开始位置
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '读取一个检验项目名称，并将字符指针移到项目名称之后
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function



Public Function CmdAnalyse(strCmdType As String) As String
    '功能:  分解命令为仪器执行的命令
    
'    strControl = Me.cbo波长.Text & ";" & Me.cbo参考波长.Text & ";" & Me.cbo振板频率.Text & ";" & Me.txt振板时间 & _
'                 ";" & Me.cbo进板方式.Text & ":" & Me.cbo空白形式.Text
    
    Dim lngTime As Long
    Select Case strCmdType
        Case "Begin"                   '开始
            CmdAnalyse = "R" & vbCrLf
        Case "End"                      '结束
            CmdAnalyse = "Q" & vbCrLf
        Case "Revert"                   '回复命令
            CmdAnalyse = "OK" & vbCrLf
        Case "ReadData"                  '整板测量
            CmdAnalyse = "P" & vbCrLf
        Case "450"
            CmdAnalyse = "F2" & vbCrLf  '450波片
        Case "1(450)"                   '1号波片
            CmdAnalyse = "F1" & vbCrLf
        Case "2(500)"                   '2号波片
            CmdAnalyse = "F2" & vbCrLf
        Case "3(750)"                   '3号波片
            CmdAnalyse = "F3" & vbCrLf
        Case "2(450)"
            CmdAnalyse = "F2" & vbCrLf  '2与波片为"450"
        Case "F1"                        '1号波片
            CmdAnalyse = "F1" & vbCrLf
        Case "F2"                        '2号波片
            CmdAnalyse = "F2" & vbCrLf
        Case "F3"                        '3号波片
            CmdAnalyse = "F3" & vbCrLf
        Case "F4"                        '4号波片
            CmdAnalyse = "F4" & vbCrLf
        Case "F5"                        '5号波片
            CmdAnalyse = "F5" & vbCrLf
        Case "连续"                     '进板方式连续
            CmdAnalyse = "E0" & vbCrLf
        Case "逐步"                     '进板方式逐步
            CmdAnalyse = "E1" & vbCrLf
        Case "1X"                       '振板频率 1X
            CmdAnalyse = "X1" & vbCrLf
        Case "2X"
            CmdAnalyse = "X2" & vbCrLf
        Case "3X"
            CmdAnalyse = "X3" & vbCrLf
        Case "X1"                       '振板频率 1X
            CmdAnalyse = "X1" & vbCrLf
        Case "X2"
            CmdAnalyse = "X2" & vbCrLf
        Case "X3"
            CmdAnalyse = "X3" & vbCrLf
        Case "空气"
            CmdAnalyse = "A" & vbCrLf
        Case "试剂"
            CmdAnalyse = "B" & vbCrLf
        Case "Z5_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case "Z10_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case "Z15_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case "Z20_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case "Z25_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case "Z30_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5 & vbCrLf
        Case Else
            If Left(strCmdType, 1) = "Z" Then
                '振板时间
                lngTime = Val(Mid(strCmdType, 2))
                If lngTime >= 0 And lngTime <= 60 Then
                    CmdAnalyse = "Z" & Format(lngTime, "00") & vbCrLf
                End If
            End If
    End Select
    
End Function







