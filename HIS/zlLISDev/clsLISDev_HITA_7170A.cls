VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA_7170A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  日立7180
'医院:  成都德阳
'作者:  曾超
'时间:  2008/1/16
Const str_MOR = ">"
Const str_REP = "?"
Private str_REC As String
Private mstrSend As String
Private mLastSend As String

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
    
    Dim strData As String
    Dim aData() As String
    Dim intloop As Integer

    Dim strSampleNO As String           '标本号
    Dim strBarcode As String            '条码号
    Dim strLine As String               '当前要解析的字串
    Dim lngBegin As Long                '开始位置
    Dim lngEnd As String                '结束位置
    Dim str_Date  As String
    Dim str_Ji As String                '急诊标志
    Dim str_QC As String              '质控
    Dim str_Tmp As String
    Dim curALB As Currency, curGLO As Currency, curHBA1C As Currency, curHB As Currency
    
    Dim curTP As Currency, curTBIL As Currency, curDBIL As Currency, curALT As Currency, curAST As Currency
    Dim curIBIL As Currency, curASTALT As Currency
    '初始化返回值
    strReserved = "": strResult = "": strCmd = "": strSampleNO = ""
    '没有找到开始位置时退出
'    If InStr(strSource, Chr(&H2)) <= 0 Then
'        strReserved = strSource
'        'strCmd = Chr(&H6)
'        Exit Sub
'    End If
'    '没有找到结束位置时退出
'    If InStr(strSource, Chr(&H3)) <= 0 Then
'        strReserved = strSource
'        'strCmd = Chr(&H6)
'        Exit Sub
'    End If
    str_Date = "": str_Ji = "": str_QC = ""
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
    
    If str_REC = "" Then
        
        Call ReadConf(App.Path & "\HITA7170A.ini", "应答", str_REC)
        If str_REC = "" Then
            str_REC = "A"
            Call WriteIni("7170A", "应答", "A", App.Path & "\HITA7170A.ini")
        End If
    End If
    
    If strLine = str_MOR Or strLine = str_REP Then
        '双向通信的联机请求，原样应答即可
        If mstrSend = "" Then
            strCmd = str_MOR
        End If
    ElseIf strLine Like ";N1*" And Len(strLine) = 42 Then
        
        strCmd = "1|" & str_REC  '收到扫描的标本号请求
    Else
        '开始解析(没有找到时间，暂时使用本机时间)
        strCmd = str_REC
        
        If InStr(",:,1,2,", Mid(strLine, 2, 1)) > 0 Then
        
            If UCase(Mid(strLine, 3, 1)) = "Q" Then
                str_Ji = "1"
            ElseIf UCase(Mid(strLine, 3, 1)) = "G" Then
                Exit Sub
            End If
            
            str_QC = "0"
            If Mid(strLine, 2, 1) = "1" Then
                str_QC = "1"
            ElseIf Mid(strLine, 3, 1) = "F" Then
                str_QC = "F1"
            End If
            If Val(Mid(strLine, 5, 5)) = 0 Then '取条码
                
                If Val(Mid(strLine, 15, 2)) <> 0 Then
                    '条码前两位不为0 ，才是条码
                    strData = Now & "|^" & IIf(str_Ji = "1", "1", "") & "^" & Val(Mid(strLine, 15, 13)) & "| |血清|" & str_QC
                Else
                    '条码前两位为0,就把条码作为标本号处理
                    strData = Now & "|" & Val(Mid(strLine, 15, 13)) & IIf(str_Ji = "1", "^1^", "") & "| |血清|" & str_QC
                End If
            Else
                str_Date = Trim(Mid(strLine, 32, 10))
                If IsNumeric(str_Date) And Len(str_Date) = 10 Then
                    str_Date = "20" & Mid(str_Date, 5, 2) & "-" & Mid(str_Date, 1, 2) & "-" & Mid(str_Date, 3, 2) & " " & Mid(str_Date, 7, 2) & ":" & Mid(str_Date, 9, 2)
                Else
                    str_Date = Format(Now, "yyyy-MM-dd hh:MM:ss")
                End If
                If str_QC = "F1" Then
                    If Trim(Mid(strLine, 14, 13)) = "CONT.01" Then '晋煤的
                        strData = Format(Now, "yyyy-MM-dd HH:mm") & "|1014" & IIf(str_Ji = "1", "^1^", "") & "| |血清|1"
                    ElseIf Trim(Mid(strLine, 14, 13)) = "CONT.02" Then '晋煤的
                        strData = Format(Now, "yyyy-MM-dd HH:mm") & "|2014" & IIf(str_Ji = "1", "^1^", "") & "| |血清|1"
                    ElseIf IsNumeric(Trim(Mid(strLine, 14, 13))) Then   '广西的
                        strData = Format(Now, "yyyy-MM-dd HH:mm") & "|" & Trim(Mid(strLine, 14, 13)) & "^1^" & "| |血清|0"
                    Else
                        strData = Format(Now, "yyyy-MM-dd HH:mm") & "|" & Trim(Mid(strLine, 5, 5)) & IIf(str_Ji = "1", "^1^", "") & "| |血清|1"
                    End If
                Else
                    strData = str_Date & "|" & Val(Mid(strLine, 5, 5)) & IIf(str_Ji = "1", "^1^", "") & "| |血清|" & str_QC
                End If
            End If
            '开始处理结果
            strLine = Replace(Mid(strLine, 45), "", "")
            
            For intloop = 0 To Len(strLine) / 10
                If Len(strLine) < 10 Then Exit For
                str_Tmp = Mid(strLine, 1, 10)
                If InStr(str_Tmp, "%") > 0 Then str_Tmp = Replace(str_Tmp, "%", "")
                strData = strData & "|" & Trim(Mid(str_Tmp, 1, 3)) & "|" & Val(Trim(Mid(str_Tmp, 4)))
                
                strLine = Mid(strLine, 11)
            Next
            
            strResult = strData
            If strData <> "" Then
                aData = Split(strData, "|")
                For intloop = 5 To UBound(aData)
                    If intloop Mod 2 = 1 Then
                        If Val(Trim(aData(intloop))) = 12 Then curTP = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 13 Then curALB = Val(Trim(aData(intloop + 1)))
                        
                        If Val(Trim(aData(intloop))) = 28 Then curHBA1C = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 29 Then curHB = Val(Trim(aData(intloop + 1)))
                        
                        If Val(Trim(aData(intloop))) = 15 Then curTBIL = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 16 Then curDBIL = Val(Trim(aData(intloop + 1)))
                        
                        If Val(Trim(aData(intloop))) = 1 Then curALT = Val(Trim(aData(intloop + 1)))
                        If Val(Trim(aData(intloop))) = 30 Then curAST = Val(Trim(aData(intloop + 1)))
                    End If
                Next
                If curALB <> 0 And curTP <> 0 Then
                    If curTP >= curALB Then
                        curGLO = curTP - curALB
                        strResult = strResult & "|X002|" & Format(curGLO, "0.000")
                    End If
                    If curGLO <> 0 Then
                        strResult = strResult & "|110|" & Format(curALB / curGLO, "0.000")
                    End If
                End If
                If curHBA1C <> 0 And curHB <> 0 Then
                    strResult = strResult & "|1111|" & Format(curHBA1C * 87.6 / curHB + 2.27, "0.000")
                End If
                
                If curTBIL <> 0 And curDBIL <> 0 Then
                    curIBIL = curTBIL - curDBIL
                    strResult = strResult & "|94|" & Format(curIBIL, "0.000")
                End If
                
                If curAST <> 0 And curALT <> 0 Then
                    strResult = strResult & "|X006|" & Format(curAST / curALT, "0.000")
                End If
            
            End If
        End If
    
    End If
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
    Dim varItem As Variant, strItem As String, i As Integer, strSampleNO As String
    Dim str盘号 As String, str杯号  As String, lng_S As Long, lng_E As Long
    '2008-12-29 11:39:00|110|闫黑明|血液|0| |0,10| | | |29|44
    On Error GoTo errHandle
    
    blnSuccess = True
    If iSendStep = 0 Then
        
        If strSample <> "" Then
            strItem = String(88, "0")
            
            varItem = Split(strSample, "|")
            
            strSampleNO = Val(varItem(1))
            strSampleNO = Space(5 - Len(strSampleNO)) & strSampleNO
            
            If InStr(varItem(6), ",") > 0 Then
                blnSuccess = True
                
                str盘号 = Split(varItem(6), ",")(0)
                str杯号 = Split(varItem(6), ",")(1)
                str杯号 = Space(3 - Len(str杯号)) & str杯号
                
                SendSample = ";N1" & strSampleNO & str盘号 & str杯号 & Space(29)
        
                        
                For i = 10 To UBound(varItem)
                    If varItem(i) <= 88 And varItem(i) > 0 Then
                        strItem = Mid(strItem, 1, varItem(i) - 1) & "1" & Mid(strItem, varItem(i) + 1)
                    End If
                Next
                If strItem <> String(88, "0") Then
                    SendSample = SendSample & "88" & strItem & "00000"
                End If
            Else
                WriteLog "sendSample", "盘号,杯号错误,应为0,1的格式", varItem(6)
            End If
            If SendSample <> "" Then
                If mstrSend <> "" Then
                    lng_S = InStr(mstrSend, Chr(&H2))
                    lng_E = InStr(mstrSend, Chr(&H3))
                    SendSample = Mid(mstrSend, lng_S, lng_E - lng_S + 1)
                    mstrSend = Mid(mstrSend, lng_E + 1)
                End If
                mstrSend = mstrSend & SendSample
            
                If mLastSend = SendSample Then
                    SendSample = ""
                Else
                    mLastSend = SendSample
                End If
                iSendStep = 1
            End If
        End If
    Else
        If strResponse = str_MOR Or strResponse = str_REP Then
            If mstrSend <> "" Then
                lng_S = InStr(mstrSend, Chr(&H2))
                lng_E = InStr(mstrSend, Chr(&H3))
                SendSample = Mid(mstrSend, lng_S, lng_E - lng_S + 1)
                mstrSend = Mid(mstrSend, lng_E + 1)
                If mLastSend = SendSample Then
                    SendSample = ""
                    If mstrSend <> "" Then
                        lng_S = InStr(mstrSend, Chr(&H2))
                        lng_E = InStr(mstrSend, Chr(&H3))
                        SendSample = Mid(mstrSend, lng_S, lng_E - lng_S + 1)
                        mLastSend = SendSample
                    End If
                Else
                    mLastSend = SendSample
                End If
            End If
        End If
        
        If mstrSend = "" Then iSendStep = 0
    End If
    WriteLog "SendSample", strSample, "SendSample=" & SendSample & vbNewLine & "iStep=" & iSendStep & vbNewLine & "mStrsend=" & mstrSend & vbNewLine & "mLastSend=" & mLastSend
    Exit Function
errHandle:
    WriteLog "sendSample", strSample, Err.Description
End Function


Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '读配置文件
    'strIniFile :ini配置文件名
    'strItems   :要读取的字段，如：DATABASE NAME=th_xlb中的 DATABASE NAME,多个字段以,分隔
    'strReturn  :读取的返回值，如:DATABASE NAME=th_xlb中的 th_xlb ，多个反回值以,分隔
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As TextStream

    Dim intloop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '从配置文件中读取保存在文件中的配置cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intloop = 0 To intCount
        arrItems(intloop, 0) = Split(strItems, ",")(intloop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intloop = 0 To intCount
                    If Trim(arrLine(0)) = arrItems(intloop, 0) Then
                        arrItems(intloop, 1) = arrLine(1)
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intloop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intloop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If
    
    
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub
