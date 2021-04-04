VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABX_P120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血常规
'厂商:??
'作者:曾超
'用户:重庆市二院
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
    Dim lngloop As Long, intLoop As Integer
    Dim strDate As String
    Dim i As Integer
    Dim strTmp As String
    Dim strSampleNO As String           '标本号
    Dim strBarcode As String            '条码号
    Dim strImageData As String          '图像信息
    
    Dim dbl总蛋白 As Double, dbl白蛋白 As Double, dbl低密度脂蛋白 As Double, dbl高密度脂蛋白 As Double
    Dim dbl载脂蛋白A As Double, dbl载脂蛋白B As Double, dbl总胆红素 As Double, dbl直接胆红素 As Double
    
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    If InStr(strSource, "") > 0 Then
        
        aData = Split(strSource, "")
        
        For lngloop = 0 To UBound(aData)
            
            aItem = Split(aData(lngloop), vbCr)
            
            '没有传完数据时退出
            If UBound(aItem) < 71 And lngloop = UBound(aData) Then
                If strData <> "" Then
                    strResult = strData & vbCrLf & Mid(strImageData, 2)
                End If
                strReserved = aData(lngloop)
                Exit Sub
            End If
            
            strTmp = ""

            i = -1
            For intLoop = 0 To UBound(aItem)
                '指标
                If InStr(aItem(intLoop), "! ") > 0 Then
                    i = i + 1
                    '新记录
                    If strData <> "" Then
                        strData = strData & vbCrLf & Mid(strImageData, 3) & "||"
                    End If
                End If
                If i < 30 And i >= 0 Then
                    i = i + 1
                    strTmp = strTmp & "|" & i & "|" & Val(Mid(aItem(intLoop), 2))
                End If
                
                '图像信息
                If InStr(aItem(intLoop), "X ") > 0 Then         'RBC
                    strImageData = strImageData & "^" & GetImageStr("RBC", Mid(Replace(Replace(aItem(intLoop), vbCr, ""), vbLf, ""), 2))
                End If
                
                If InStr(aItem(intLoop), "Y ") > 0 Then         'PLT
                    strImageData = strImageData & "^" & GetImageStr("PLT", Mid(Replace(Replace(aItem(intLoop), vbCr, ""), vbLf, ""), 2))
                End If

                If InStr(aItem(intLoop), "Z ") > 0 Then         'BAS
                    strImageData = strImageData & "^" & GetImageStr("BAS", Mid(Replace(Replace(aItem(intLoop), vbCr, ""), vbLf, ""), 2))
                End If

                If InStr(aItem(intLoop), "[ ") > 0 Then         'WBC
                    strImageData = strImageData & "^" & GetImageStr("WBC", Mid(Replace(Replace(aItem(intLoop), vbLf, ""), vbCr, ""), 2))
                End If
'
                '标本信息
                If InStr(aItem(intLoop), "q ") > 0 Then
                    i = intLoop
                    strSampleNO = Trim(Replace(Replace(Replace(aItem(i + 1), "r", ""), vbCr, ""), vbLf, ""))
                    strSampleNO = Val("9" & Mid(strSampleNO, 1, 3) & Right(strSampleNO, 2))
                    strBarcode = Trim(Val(Trim(Replace(Mid(aItem(i + 4), 2), "u", ""))))
                    strData = strData & Now & "|" & IIf(strBarcode = 0, Trim(Val(Replace(Trim(Mid(aItem(i + 1), 3)), "r", ""))), strSampleNO & "^0") & _
                    IIf(Len(strBarcode) > 6, "^" & strBarcode, "") & "| |血液|0" & strTmp

                End If
            Next
        Next
    Else
        strReserved = strSource
    End If
    
    If strData <> "" Then
        strResult = strData & vbCrLf & strImageData
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
