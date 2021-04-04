VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HYCEL_DIANA5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:五分类血球仪
'厂商:法国HYCEL
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
    Dim intLoopY As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '标本号
    Dim strBarcode As String            '条码号
    Dim strImageData As String          '图像信息
    Dim intStart As Integer             '开始位置
    Dim intEnd As Integer               '结束位置
    Dim strDate As String               '检验时间
    
    strCmd = Chr(&H6)
    
    '没有找到结束符"&H03"时退出
    If InStr(strSource, Chr(&H3)) <= 0 Then strReserved = strSource: Exit Sub
    
    
    intStart = InStr(strSource, "") + 1
    intEnd = InStr(strSource, "")
    
    '只取第一段解析后面的先返回下次解析
    strTmp = Mid(strSource, intStart, intEnd - intStart)
    
    strReserved = Mid(strSource, intEnd + 1)
    If InStr(strReserved, Chr(&H2)) <= 0 Then
        '没有开始符时没有返回值
        strReserved = ""
    End If
    
    '开始解析
    
    strDate = Mid(strTmp, 4, 4) & "-" & Mid(strTmp, 8, 2) & "-" & Mid(strTmp, 10, 2)
    strData = strDate & "|" & Val(Mid(strTmp, 30, 4)) & "| |血清|0"
    
    strData = strData & "|1|" & Val(Mid(strTmp, 100, 17)) & "|2|" & Val(Mid(strTmp, 118, 4))
    strData = strData & "|3|" & Val(Mid(strTmp, 122, 5)) & "|4|" & Val(Mid(strTmp, 127, 5))
    strData = strData & "|5|" & Val(Mid(strTmp, 132, 4)) & "|6|" & Val(Mid(strTmp, 136, 4))
    strData = strData & "|7|" & Val(Mid(strTmp, 140, 4)) & "|8|" & Val(Mid(strTmp, 144, 4))
    strData = strData & "|9|" & Val(Mid(strTmp, 148, 4)) & "|10|" & Val(Mid(strTmp, 152, 4))
    strData = strData & "|11|" & Val(Mid(strTmp, 156, 5)) & "|12|" & Val(Mid(strTmp, 161, 5))
    strData = strData & "|13|" & Val(Mid(strTmp, 166, 5)) & "|14|" & Val(Mid(strTmp, 171, 5))
    strData = strData & "|15|" & Val(Mid(strTmp, 176, 5)) & "|16|" & Val(Mid(strTmp, 181, 5))
    strData = strData & "|17|" & Val(Mid(strTmp, 186, 5)) & "|18|" & Val(Mid(strTmp, 191, 4))
    strData = strData & "|19|" & Val(Mid(strTmp, 195, 4)) & "|20|" & Val(Mid(strTmp, 199, 4))
    strData = strData & "|21|" & Val(Mid(strTmp, 203, 4)) & "|22|" & Val(Mid(strTmp, 207, 4))
    
    strResult = strData

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
Private Function ReplaceSpace(strData) As String
    Dim intloop As Integer
    Dim strTmp As String
    For intloop = 1 To Len(strData)
        If Mid(strData, intloop, 1) = " " Then
            If Mid(strTmp, Len(strTmp), 1) <> " " Then
                strTmp = strTmp & Mid(strData, intloop)
            End If
        Else
            strTmp = strTmp & Mid(strData, intloop)
        End If
    Next
End Function

