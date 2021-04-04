VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Tecan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Tecan Sunrise
'厂商:Tecan(瑞士帝肯)
'作者:陈东
'用户:德阳人民医院
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
    Dim adata() As String, iDataPoint As Integer, aitem() As String
    
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
    
    Dim varLine As Variant
    '没有这么长时退出
    If UBound(Split(strSource, vbCr)) < 9 Then
        strReserved = strSource
        Exit Sub
    End If
    
    '开始解析
    strResult = ""
    For intloop = 1 To 8
        strLine = ""
        varLine = Split(Trim(Split(strSource, vbCr)(intloop)), " ")
        For i = LBound(varLine) To UBound(varLine)
            strLine = strLine & ";" & Format(Val(varLine(i)) / 1000, "0.000")
        Next
        strResult = strResult & Mid(strLine, 2) & "|"
    Next
    
    
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

Public Function CmdAnalyse(strCmdType As String) As String
    '功能:  分解命令为仪器执行的命令
    
'    strControl = Me.cbo波长.Text & ";" & Me.cbo参考波长.Text & ";" & Me.cbo振板频率.Text & ";" & Me.txt振板时间 & _
'                 ";" & Me.cbo进板方式.Text & ":" & Me.cbo空白形式.Text
    
    Select Case strCmdType
        Case "Begin"                                                    '开始 出板
            CmdAnalyse = "" 'Chr(&H2) & "TR0" & Chr(&H3) & "7" & Chr(&HD) & "|+|-"
        Case "End"                                                      '结束 进板
            CmdAnalyse = Chr(&H2) & "TR1" & Chr(&H3) & "6" & Chr(&HD) & "|+|-"
        Case "Revert"                                                       '回复命令
            CmdAnalyse = "+" 'Chr(&H2) & "+" & Chr(&H3) & "*" & Chr(&HD) '+*
        Case "Play"                     '整板测量
            CmdAnalyse = Chr(&H2) & "SM" & Chr(&H3) & Chr(&H1F) & Chr(&HD) & "|+|D" '发SM 返回D时，可以取数
        Case "ReadData"    '取数
            CmdAnalyse = Chr(&H2) & "DA" & Chr(&H3) & Chr(&H4) & Chr(&HD) 'SM
        Case "Out"
            CmdAnalyse = Chr(&H2) & "TR0" & Chr(&H3) & "7" & Chr(&HD) & "|+|-"
        Case "In"
            CmdAnalyse = Chr(&H2) & "TR1" & Chr(&H3) & "6" & Chr(&HD) & "|+|-"
        '----------------------------------------------------------------------------------
        Case "1(405)"                   '波长一
            CmdAnalyse = Chr(&H2) & "SW405" & Chr(&H3) & "4" & Chr(&HD) & "|+"
        Case "2(450)"                   '波长二
            CmdAnalyse = Chr(&H2) & "SW450" & Chr(&H3) & "4" & Chr(&HD) & "|+"
        Case "3(492)"                   '波长三
            CmdAnalyse = Chr(&H2) & "SW492" & Chr(&H3) & ":" & Chr(&HD) & "|+"
        Case "4(620)"                   '波长四
            CmdAnalyse = Chr(&H2) & "SW620" & Chr(&H3) & "1" & Chr(&HD) & "|+"
        '----------------------------------------------------------------------------------
        Case "连续"                     '进板方式连续
            CmdAnalyse = ""
        Case "单步"                     '进板方式逐步
            CmdAnalyse = Chr(&H2) & "GD450 620" & Chr(&H3) & "'" & Chr(&HD) & "|" & Chr(&H2) & "3"
        '----------------------------------------------------------------------------------
        Case "1X"                       '振板频率 1X
            CmdAnalyse = ""
        Case "2X"
            CmdAnalyse = ""
        Case "3X"
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        Case "空气"
            CmdAnalyse = ""
        Case "试剂"
            CmdAnalyse = ""
        Case Else
'            If Val(strCmdType) > 0 Then '时间
'                CmdAnalyse = Chr(&H2) & "SH2 " & Val(strCmdType) * 10 & Chr(&H3) & Chr(&HD) & "|+"
'            End If
    End Select
    
End Function


