VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA_7060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  日立7060
'医院:  成都德阳
'作者:  曾超
'时间:  2008/1/16
Const str_MOR = ">"
Const str_REP = "?"

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
    Dim aData() As String, iDataPoint As Integer, aitem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngLoop As Long, intloop As Integer
    Dim strDate As String

    Dim strLine As String               '当前要解析的字串
    Dim lngBegin As Long                '开始位置
    Dim lngEnd As String                '结束位置
    Dim str_日期  As String
    Dim str_急 As String                '急诊标志
    Dim str_质控 As String              '质控
    Dim str_Tmp As String
    Dim curALB As Currency, curGLO As Currency, curTBIL As Currency, curDBIL As Currency, curIBIL As Currency
    Dim curTP As Currency, curAST As Currency, curALT As Currency
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    '没有找到开始位置时退出
    
    If Not strSource Like "***" Then
        strReserved = strSource
        Exit Sub
    End If
    
    lngBegin = InStr(strSource, Chr(&H2))
    lngEnd = InStr(strSource, Chr(&H3))
    
    '取当前一次的结果，其他的返回
    strLine = Mid(strSource, lngBegin, lngEnd - lngBegin + 1)
    strReserved = Mid(strSource, lngEnd + 1)
    If InStr(strReserved, Chr(&H2)) <= 0 Then
        strReserved = ""    '没有找到开始位置时退出
    End If
    '收到一串数据，应答
    strCmd = Chr(&H6)

    '开始解析
    If InStr(",:,1,2,", Mid(strLine, 2, 1)) > 0 Then
        If UCase(Mid(strLine, 3, 1)) = "G" Then
            str_急 = "1"
        End If
        
        str_质控 = "0"
        If Mid(strLine, 3, 1) = "K" Then
            str_质控 = "1"
        End If
        
        
        If Trim(Mid(strLine, 33, 2)) = "" Or Trim(Mid(strLine, 29, 2)) = "" Or Trim(Mid(strLine, 31, 2)) = "" Then
            str_日期 = Format(Now, "yyyy-MM-dd")
        Else
            str_日期 = Mid(strLine, 33, 2) & "-" & Mid(strLine, 29, 2) & "-" & Mid(strLine, 31, 2)
        End If
        
        strData = Format(CDate(str_日期), "yyyy-MM-dd") & "|" & Val(Mid(strLine, 5, 3)) & IIf(str_急 = "1", "^1^", "") & "| |血清|" & str_质控
        
        '开始处理结果
        strLine = Replace(Mid(strLine, 41), "", "")
        
        For intloop = 0 To Len(strLine) / 9
            If Len(strLine) < 9 Then Exit For
            str_Tmp = Mid(strLine, 1, 9)
            strData = strData & "|" & Trim(Mid(str_Tmp, 1, 2)) & "|" & Val(Trim(Mid(str_Tmp, 3)))
            
            strLine = Mid(strLine, 10)
        Next
        
        strResult = strData
        If strData <> "" Then
            aData = Split(strData, "|")
            For intloop = 5 To UBound(aData)
                If intloop Mod 2 = 1 Then
                    If Val(Trim(aData(intloop))) = 7 Then curTP = Val(Trim(aData(intloop + 1)))
                    If Val(Trim(aData(intloop))) = 8 Then curALB = Val(Trim(aData(intloop + 1)))
                    
                    If Val(Trim(aData(intloop))) = 5 Then curTBIL = Val(Trim(aData(intloop + 1)))
                    If Val(Trim(aData(intloop))) = 6 Then curDBIL = Val(Trim(aData(intloop + 1)))
                    
                    If Val(Trim(aData(intloop))) = 1 Then curALT = Val(Trim(aData(intloop + 1)))
                    If Val(Trim(aData(intloop))) = 2 Then curAST = Val(Trim(aData(intloop + 1)))
                End If
            Next
            
            '--计算 球蛋白，A/G
            If curALB <> 0 And curTP <> 0 Then
                If curTP >= curALB Then
                    curGLO = curTP - curALB
                    strResult = strResult & "|GLO|" & Format(curGLO, "0.000")
                End If
                If curGLO <> 0 Then
                    strResult = strResult & "|A/G|" & Format(curALB / curGLO, "0.000")
                End If
            End If
            '-- 计算 间接胆红素，DBIL/TBIL
            If curTBIL <> 0 And curDBIL <> 0 Then
                curIBIL = curTBIL - curDBIL
                strResult = strResult & "|IBIL|" & Format(curIBIL, "0.000")
                strResult = strResult & "|DBIL/TBIL|" & Format(curDBIL / curTBIL, "0.000")
            End If
            '计算 AST/ALT
            If curAST <> 0 And curALT <> 0 Then
                strResult = strResult & "|AST/ALT|" & Format(curAST / curALT, "0.000")
            End If
        End If
    End If
    

    '
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
    iSendStep = 0
    
End Function

