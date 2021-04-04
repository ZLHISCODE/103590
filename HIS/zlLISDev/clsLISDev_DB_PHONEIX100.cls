VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_DB_PHONEIX100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  微生物 DB PHONEIX100
'医院:  成都德阳
'作者:  曾超
'时间:  2008/1/29

Dim strTmp  As String
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
    
    On Error GoTo errH
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    '没有找到开始位置时退出
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
    End If
    
    If InStr(strSource, Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    End If
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
     
    '没有找到结束位置时退出
    lngBegin = InStr(mstrCode, Chr(&H5))
    If lngBegin > 0 Then
        lngEnd = InStr(lngBegin, mstrCode, Chr(&H4))
    Else
        lngEnd = InStr(mstrCode, Chr(&H4))
    End If
    
    Do While lngBegin > 0 Or lngEnd > 0
    
        '取当前一次的结果，其他的返回
        strLine = Mid(mstrCode, lngBegin, lngEnd - lngBegin + 1)
        mstrCode = Mid(mstrCode, lngEnd + 1)
    
        '去掉０２,１７字符
        strLine = DropEccCode(strLine)
        
        '开始解析
        aData = Split(strLine, Chr(&HD))
        aItem = Split(aData(0), "|")
        '日期
        strData = Mid(aItem(13), 1, 4) & "-" & Mid(aItem(13), 5, 2) & "-" & Mid(aItem(13), 7, 2) & _
                  " " & Mid(aItem(13), 9, 2) & ":" & Mid(aItem(13), 11, 2) & ":" & Mid(aItem(13), 13, 2)
        '标本号^标本
        aItem = Split(aData(2), "|")
        aItem = Split(aItem(2), "^")
        strData = strData & "|" & Val(aItem(0)) & "| |" & aItem(2) & "|0"
        
        '开始解析结果
        For intloop = 0 To UBound(aData)
            If UBound(Split(aData(intloop), "|")) >= 1 Then
                If Split(aData(intloop), "|")(0) = "R" And Val(Split(aData(intloop), "|")(1)) > 1 Then
                    aItem = Split(aData(intloop), "|")
                    If UBound(aItem) < 4 Then
                        aItem = Split(Replace(aData(intloop) & aData(intloop + 1), vbLf, ""), "|")
                    End If
                    If UBound(aItem) >= 3 Then
                        If UBound(Split(aItem(3), "^")) >= 2 Then
                        
                            strData = strData & "|" & Replace(Replace(Split(aItem(2), "^")(5), Chr(&H17), ""), Chr(&H2), "") & _
                                        "|" & Split(aItem(3), "^")(2)
                            If InStr(Split(aItem(2), "^")(3), "MIC") > 0 Then
                                strData = strData & "^1"
                            ElseIf InStr(Split(aItem(2), "^")(3), "DISK") > 0 Then
                                strData = strData & "^2"
                            Else
                                strData = strData & "^3"
                            End If
                            strData = strData & "^" & Split(aItem(3), "^")(1)
                        End If
                    End If
                End If
            End If
        Next
        
        strResult = strResult & "||" & strData
        lngBegin = InStr(mstrCode, Chr(&H5))
        If lngBegin > 0 Then
            lngEnd = InStr(lngBegin, mstrCode, Chr(&H4))
        Else
            lngEnd = InStr(mstrCode, Chr(&H4))
        End If

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




Private Function DropEccCode(ByVal strCode As String) As String
    '去掉传入的校验码信息
    Dim lngS As Long, lngE As Long
    Dim strIn As String, strReturn As String
    strIn = strCode
    strReturn = ""
    Do While Len(strIn) > 0
        lngS = InStr(strIn, Chr(&H2))
        If lngS > 0 Then
            strReturn = strReturn & Mid(strIn, 1, lngS - 1)
            strIn = Mid(strIn, lngS + 2)
        Else
            strReturn = strReturn & strIn
            strIn = ""
        End If
    Loop
    
    strIn = strReturn
    strReturn = ""
    
    Do While Len(strIn) > 0
        lngS = InStr(strIn, Chr(&H17))
        If lngS > 0 Then
            strReturn = strReturn & Mid(strIn, 1, lngS - 1)
            strIn = Mid(strIn, lngS + 5)
        Else
            strReturn = strReturn & strIn
            strIn = ""
        End If
    Loop
    DropEccCode = strReturn
End Function





