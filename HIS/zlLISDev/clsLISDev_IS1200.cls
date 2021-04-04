VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_IS1200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:IS 1200 化学发光仪
'厂商:四川迈克生物
'作者:李纪军
'用户:府谷县医院
'时间:2013-12-06
Private str_Count As String '用于延时
Private mstrSendCmd As String  '用于延时应答上架指令
Private mstrRack As String    '扫条码时传的盘号,架号等需要返回的信息.
Private mstrWaitSend As String '待发送标本

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = Chr(6)
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
    Dim adata() As String, iDataPoint As Integer, aItem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngLoop As Long, intLoop As Integer
    Dim strDate As String, strItem As String, msStrSource As String
    Dim i As Long, str_Q As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '标本号
    Dim strBarcode As String            '条码号
    Dim strImageData As String          '图像信息
    Dim strNow As String, lngS As Long, lngE As Long
    Dim str_02 As Long '记录02的次数
    strItem = ""
    str_02 = 0
    str_Q = 0
    '没有开始符时退出
    lngS = InStr(strSource, "1H")
    lngE = InStr(strSource, "L|1")
    
    On Error GoTo errHand

    If InStr(strSource, Chr(&H5)) > 0 Then
'        '应答
        strCmd = Chr(&H6)
'        Exit Sub
    End If
    
    '没有结束符时退出
    If InStr(strSource, Chr(&H4)) > 0 Then
'        '应答
        strCmd = Chr(&H6)

    End If
    

    '取当前段来解析
    '开始解析
    If lngS > 0 And lngE > 0 And lngE > lngS Then
      Do While lngE > 0 And lngS > 0
            msStrSource = Trim(Mid(strSource, lngS, lngE - lngS + 4))
            strSource = Trim(Mid(strSource, lngE + 4, Len(strSource) - lngE))
            lngS = InStr(strSource, "1H")
            lngE = InStr(strSource, "L|1")
            ''查找并替换错误符号
            adata = Split(msStrSource, Chr(&HD))
            For i = LBound(adata) To UBound(adata)
                strTmp = Trim(adata(i))
                'Debug.Print strTmp
                If InStr(strTmp, Chr(&H2)) > 0 Then str_02 = Mid(strTmp, InStr(strTmp, Chr(&H2)) + 1, IIf(str_02 > 9, 2, 1))
                If strTmp <> "" Then
                    '标本号
                    If Split(strTmp, "|")(0) = "O" And strSampleNO = "" Then strSampleNO = Split(Split(strTmp, "|")(3), "^")(0)
                    If Split(strTmp, "|")(0) = "O" And strDate = "" Then
                        strDate = Mid(Split(Split(strTmp, "|")(22), "^")(0), 1, 4) & "-" & Mid(Split(Split(strTmp, "|")(22), "^")(0), 5, 2) & "-" & Mid(Split(Split(strTmp, "|")(22), "^")(0), 7, 2) & " " & Mid(Split(Split(strTmp, "|")(22), "^")(0), 9, 2) & ":" & Mid(Split(Split(strTmp, "|")(22), "^")(0), 11, 2) & ":" & Mid(Split(Split(strTmp, "|")(22), "^")(0), 13, 2)
                    End If
                    If Split(strTmp, "|")(0) = "O" Then str_Q = IIf(Split(strTmp, "|")(11) = "Q", 1, 0)
                    If Split(strTmp, "|")(0) = "R" Or Split(strTmp, "|")(0) Like "*R" Then
                        If InStr(strTmp, "") > 0 Then
                             If InStr(adata(i + 1), "") > 0 Then
                                If InStr(adata(i + 1), Chr(&H2)) > 0 Then str_02 = Mid(adata(i + 1), InStr(adata(i + 1), Chr(&H2)) + 1, IIf(str_02 > 9, 2, 1))
                                strTmp = Replace(Replace(strTmp, Mid(strTmp, InStr(strTmp, "")), ""), Chr(&HA) & Chr(&HD), "") & Replace(Trim(adata(i + 1)), Chr(&H2) & str_02, "")
                                strItem = strItem & "|" & Val(Split(Split(strTmp, "|")(2), "^")(3)) & "|" & Trim(Split(strTmp, "|")(3))
                             End If
                        Else
                            strItem = strItem & "|" & Val(Split(Split(strTmp, "|")(2), "^")(3)) & "|" & Trim(Split(strTmp, "|")(3))
                        End If
                    End If
                End If
            Next i
            If strDate <> "" And strSampleNO <> "" And strItem <> "" Then
                strResult = strResult & "||" & strDate & "|" & IIf(str_Q = 1, 8000 + Val(strSampleNO), Val(strSampleNO)) & "| |全血|" & str_Q & strItem
            End If
            strDate = ""
            strSampleNO = ""
            strItem = ""
            str_Q = 0
        Loop
        strSource = ""
        strReserved = ""
        '应答
        strCmd = Chr(&H6)
    Else
        strReserved = strReserved & strSource
    End If
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
    
errHand:
    WriteLog "IS1200", Err.Number & ":" & Err.Description, ""
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
Public Function SendSample11111(ByVal strSampleinfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    'iSendStep：当前发送操作的步骤。如果处理完后发送结束，则返回0，否则自动＋1返回。
    'strResponse：当前仪器的响应内容
    'strSampleInfo：发送的标本信息。
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否急诊
    '   第5个无素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。

    Dim var_标本 As Variant, i_Count As Integer, i_Loop As Integer, str_记录 As String, str_标本号 As String
    Dim i_急诊 As Integer, str_条码 As String, str_检验项目 As String, str_标本类型 As String
    Dim str_SendItem As String, str_Line As String

    WriteLog "开始", strSampleinfo, ""

    var_标本 = Split(strSampleinfo, "||")
    SendSample11111 = ""
    For i_Count = LBound(var_标本) To UBound(var_标本)
        str_记录 = var_标本(i_Count)

        If Val(Split(str_记录, "|")(1)) = 0 Then
            str_标本号 = Trim(Split(str_记录, "|")(5))
        Else
            str_标本号 = Val(Split(str_记录, "|")(1))
        End If
        str_标本类型 = Split(str_记录, "|")(3)
        i_急诊 = Val(Split(str_记录, "|")(4))
        str_条码 = Trim(Split(str_记录, "|")(5))
        str_检验项目 = ""

        For i_Loop = 10 To UBound(Split(str_记录, "|"))
            If Val(Split(str_记录, "|")(i_Loop)) > 0 Then
                str_检验项目 = str_检验项目 & "^^^" & Split(str_记录, "|")(i_Loop) & "^1\"
            End If
        Next

        If str_检验项目 <> "" Then
            str_检验项目 = Mid(str_检验项目, 1, Len(str_检验项目) - 1)

            str_Line = "1H|\^&|||IS11200^1|||||host|TSDWN^REAL|P|1" & Chr(&HD) & Chr(&H3)
            str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","
            'P|1|||||||U||||||
            str_Line = "2P|1|" & str_条码 & "||||||U||||||" & Chr(&HD) & Chr(&H3)
            str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","

            str_Line = "3O|1|" & str_条码 & "|" & mstrRack & "|" & str_检验项目 & "|" & IIf(i_急诊 = 1, "S", "R")
            str_Line = str_Line & "||" & Split(str_记录, "|")(0) & "||||N|||1|||||||||||O" & Chr(&HD) & Chr(&H3)
            str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","

            str_Line = "4L|1|N" & Chr(&HD) & Chr(&H3)
            str_SendItem = str_SendItem & Chr(&H2) & str_Line & GetCheckCode(str_Line) & Chr(&HD) & Chr(&HA) & ","

            If SendSample11111 = "" Then
                SendSample11111 = SendSample11111 & Chr(&H5)                              '加上 消息结束标志(EOT)
                mstrWaitSend = mstrWaitSend & "," & str_SendItem & Chr(&H4)
            Else
                mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & str_SendItem & Chr(&H4)
            End If
        End If
    Next
    iSendStep = 0
    WriteLog "结束", "发送到仪器的信息", SendSample11111
End Function

'获取校验码
Private Function GetCheckCode(ByVal str_in As String) As String
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 100)), 2)
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


Private Function GetImageStr(strType As String, strData As String) As String
    Dim intLoop As Integer
    Dim strTmp As String
    
    Select Case strType
        Case "WBC"
            GetImageStr = GetWBCstr(Trim$(strData))
            GetImageStr = strType & ";1;" & Mid(GetImageStr, 2)
        Case Else
            For intLoop = 1 To Len(strData)
        '        Debug.Print Mid(strData, intloop, 1)
                strTmp = Hex(Asc(Mid(strData, intLoop, 1)))
                If Len(strTmp) >= 3 Then
                    GetImageStr = GetImageStr & ";" & Val("&H" & Mid(strTmp, 1, 2))
                    GetImageStr = GetImageStr & ";" & Val("&H" & Mid(strTmp, 3))
                Else
                    GetImageStr = GetImageStr & ";" & Val("&H" & Mid(strTmp, 1, 2))
                End If
            Next
            
            '写入一些图形固定的大小和标志位置
            Select Case strType
                Case "RBC"
                    GetImageStr = strType & ";0;32,176,11-50,35-100,59-,83-200,104-,128-300,152-;" & Mid(GetImageStr, 2)
                Case "PLT"
                    GetImageStr = strType & ";0;32,170,7-2,27-,54-20,81-,108-40,135-,162-60;" & Mid(GetImageStr, 2)
                Case "BAS"
                    GetImageStr = strType & ";0;32,130,12-,25-100,38-,51-200,64-,77-300,90-,103-400,116-" & Mid(GetImageStr, 2)
            End Select
            
            
    End Select
    
End Function

Private Function GetWBCstr(strData) As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intLoop As Integer
    Dim strVal As String
    
    
    '写入数据
    For intLoop = 1 To Len(strData)
        '计算Y坐标
        If intLoop <= 32 Then
            intY = 1
        Else
            intY = CInt(intLoop / 32)
        End If
        intY = 128 - intY + 1
'        If intY <> intCheckY Then
'            strData = strData & ";" & vbCrLf
''            Debug.Print strData
'        End If
        
        '计算X坐标
        intX = intLoop Mod 32
        If intX = 0 Then intX = 32
        
        strVal = Mid(strData, intLoop, 1)
        
        Select Case strVal
            Case "1", "A"       '4
                GetWBCstr = GetWBCstr & "0001"
            Case "2", "B"       '3
                GetWBCstr = GetWBCstr & "0010"
            Case "3", "C"       '3,4v
                GetWBCstr = GetWBCstr & "0011"
            Case "4", "D"       '2
                GetWBCstr = GetWBCstr & "0100"
            Case "5", "E"       '2,4
                GetWBCstr = GetWBCstr & "0101"
            Case "6", "F"       '2,3
                GetWBCstr = GetWBCstr & "0110"
            Case "7"            '2,3,4
                GetWBCstr = GetWBCstr & "0111"
            Case "8"            '1
                GetWBCstr = GetWBCstr & "1000"
            Case "9"            '1,4
                GetWBCstr = GetWBCstr & "1001"
            Case "A"            '4
                GetWBCstr = GetWBCstr & "0001"
            Case "0"            '0
                GetWBCstr = GetWBCstr & "0000"
        End Select
        
        If intX = 32 Then
            GetWBCstr = GetWBCstr & ";"
        End If
    Next
End Function




