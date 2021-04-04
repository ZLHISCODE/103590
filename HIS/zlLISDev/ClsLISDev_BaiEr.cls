VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLISDev_BaiEr_500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  拜耳500
'医院:  成都德阳
'作者:  曾超
'时间:  2008/1/16
Private str_Count As String '记录收到的字符，用于延时应答

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
    Dim lngLoop As Long, intloop As Integer
    Dim strDate As String
    Dim i As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '标本号
    Dim strBarcode As String            '条码号
    Dim strImageData As String          '图像信息
    Dim strLine As String               '当前要解析的字串
    Dim lngBegin As Long                '开始位置
    Dim lngEnd As String                '结束位置
    
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    '没有找到开始位置时退出
    
    
    'strTmp = Replace(strSource, str_Count, "")
    If Len(strSource) > 0 Then str_Count = str_Count & strSource
    
    If InStr(strSource, Chr(&H5)) > 0 Or InStr(strSource, Chr(&H4)) > 0 Then
        strCmd = Chr(&H6)
    End If
    If InStr(strSource, Chr(&HD) & Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    End If
        
    lngBegin = InStr(str_Count, Chr(&H5))
    If lngBegin > 0 Then
        lngEnd = InStr(lngBegin, str_Count, Chr(&H4))
    Else
        lngEnd = InStr(str_Count, Chr(&H4))
    End If
    
    Do While lngBegin > 0 And lngEnd > 0
    '取当前一次的结果，其他的返回
        strLine = Replace(Mid(str_Count, lngBegin, lngEnd - lngBegin + 1), Chr(&HA), "")
        str_Count = Mid(str_Count, lngEnd + 1)
        
        '开始解析
        aData = Split(strLine, vbCr)
        
        '数据长度不够时退出
        For i = LBound(aData) To UBound(aData)
            If aData(i) Like Chr(&H2) & "?P|?|*" Then
                aItem = Split(aData(2), "|")                '日期
                strData = Mid(aItem(3), 1, 4) & "-" & Mid(aItem(3), 5, 2) & "-" & Mid(aItem(3), 7, 2) & " "                '时间
                strData = strData & Mid(aItem(3), 9, 2) & ":" & Mid(aItem(3), 11, 2) & ":" & Mid(aItem(3), 13, 2)
                strData = strData & "|" & Val(aItem(2)) & "| |尿|0"                 '标本号
            
            ElseIf aData(i) Like Chr(&H2) & "?R|?|*" Or aData(i) Like "R|*" Then
                '解析结果
                aItem = Split(aData(i), "|")
                If UBound(aItem) >= 5 Then
                    strTmp = Replace(aItem(5), vbLf, "")
                    If InStr(strTmp, "^") > 0 Then strTmp = Split(strTmp, "^")(0)
                    strData = strData & "|" & aItem(3) & "|" & strTmp
                End If
            End If
        Next
        '替换英文为中文
        strData = Replace(strData, "CLEAR", "透明")
        strData = Replace(strData, "CLOUDY", "浑浊")
        strData = Replace(strData, "NEGATIVE", "阴性")
        strData = Replace(strData, "POSITIVE", "阳性")
        
        strData = Replace(strData, "TRACE-LYSED", "微量")
        strData = Replace(strData, "TRACE-INTACT", "微量")
        strData = Replace(strData, "TRACE", "弱阳性")
        
        strData = Replace(strData, "LT. YELLOW", "淡黄色")
        strData = Replace(strData, "DK. YELLOW", "深黄色")
        strData = Replace(strData, "YELLOW", "黄色")
        
        strData = Replace(strData, "LT. RED", "浅红色")
        strData = Replace(strData, "DK. RED", "深红色")
        strData = Replace(strData, "RED", "红色")
        
        
        strData = Replace(strData, "DK. ORANGE", "深橙色")
        strData = Replace(strData, "LT. ORANGE", "浅橙色")
        strData = Replace(strData, "ORANGE", "橙色")
        
        strData = Replace(strData, "DK. BROWN", "深褐色")
        strData = Replace(strData, "BROWN", "褐色")
        
        strData = Replace(strData, "LT. GREEN", "浅绿色")
        strData = Replace(strData, "GREEN", "绿色")
        If strData <> "" Then strResult = strResult & "||" & strData
        
        lngBegin = InStr(str_Count, Chr(&H5))
        If lngBegin > 0 Then
            lngEnd = InStr(lngBegin, str_Count, Chr(&H4))
        Else
            lngEnd = InStr(str_Count, Chr(&H4))
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    'strCmd = Chr(&H6)
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






