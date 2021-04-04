VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HITA_7020"
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
    Dim aData() As String, iDataPoint As Integer, aItem() As String
    
    Dim strChenal As String, dblValue As Double, iPos As Long, iStartPoint As Integer
    Dim lngLoop As Long, intLoop As Integer
    Dim strDate As String

    Dim strLine As String               '当前要解析的字串
    Dim lngBegin As Long                '开始位置
    Dim lngEnd As String                '结束位置
    Dim str_日期  As String
    Dim str_急 As String                '急诊标志
    Dim str_质控 As String              '质控
    Dim str_tmp As String
    Dim curALB As Currency, curGLO As Currency, curTBIL As Currency, curDBIL As Currency, curIBIL As Currency
    Dim curTP As Currency, curAST As Currency, curALT As Currency, cur9 As Currency, cur10 As Currency
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
    If InStr(",:,1,2,", Mid(strLine, 2, 1)) > 0 And InStr(",N,Q,F,", UCase(Mid(strLine, 3, 1))) > 0 Then
    
        If UCase(Mid(strLine, 3, 1)) = "Q" Then
            str_急 = "1"
        End If
        
        str_质控 = "0"
        If Mid(strLine, 3, 1) = "F" Then
            str_质控 = "1"
        End If
        
        
        If Trim(Mid(strLine, 33, 2)) = "" Or Trim(Mid(strLine, 29, 2)) = "" Or Trim(Mid(strLine, 31, 2)) = "" Then
            str_日期 = Format(Now, "yyyy-MM-dd")
        Else
            str_日期 = Mid(strLine, 33, 2) & "-" & Mid(strLine, 29, 2) & "-" & Mid(strLine, 31, 2)
        End If
        
        strData = Format(CDate(str_日期), "yyyy-MM-dd") & "|" & Val(Mid(strLine, 5, 5)) & IIf(str_急 = "1", "^1^", "") & "| |血清|" & str_质控
        
        '开始处理结果
        strLine = Replace(Mid(strLine, 45), "", "")
        
        For intLoop = 0 To Len(strLine) / 10
            If Len(strLine) < 10 Then Exit For
            str_tmp = Replace(Mid(strLine, 1, 10), "%", "")
            
            strData = strData & "|" & Trim(Mid(str_tmp, 1, 3)) & "|" & Val(Trim(Mid(str_tmp, 4)))
            
            strLine = Mid(strLine, 11)
        Next
        
        strResult = strData
        If strData <> "" Then
            aData = Split(strData, "|")
            For intLoop = 5 To UBound(aData)
                If intLoop Mod 2 = 1 Then
                    If Val(Trim(aData(intLoop))) = 7 Then curTP = Val(Trim(aData(intLoop + 1)))
                    If Val(Trim(aData(intLoop))) = 8 Then curALB = Val(Trim(aData(intLoop + 1)))
                    
                    If Val(Trim(aData(intLoop))) = 5 Then curTBIL = Val(Trim(aData(intLoop + 1)))
                    If Val(Trim(aData(intLoop))) = 6 Then curDBIL = Val(Trim(aData(intLoop + 1)))
                    
                    If Val(Trim(aData(intLoop))) = 1 Then curALT = Val(Trim(aData(intLoop + 1)))
                    If Val(Trim(aData(intLoop))) = 2 Then curAST = Val(Trim(aData(intLoop + 1)))
                    
                    If Val(Trim(aData(intLoop))) = 9 Then cur9 = Val(Trim(aData(intLoop + 1)))
                    If Val(Trim(aData(intLoop))) = 10 Then cur10 = Val(Trim(aData(intLoop + 1)))
                End If
            Next
            
            '--计算 球蛋白，A/G
            If curALB <> 0 And curTP <> 0 Then
                If curTP >= curALB Then
                    curGLO = curTP - curALB
                    strResult = strResult & "|101|" & Format(curGLO, "0.000")
                End If
                If curGLO <> 0 Then
                    strResult = strResult & "|102|" & Format(curALB / curGLO, "0.000")
                End If
            End If
            '-- 计算 间接胆红素，DBIL/TBIL
            If curTBIL <> 0 And curDBIL <> 0 Then
                curIBIL = curTBIL - curDBIL
                strResult = strResult & "|103|" & Format(curIBIL, "0.000")
                strResult = strResult & "|105|" & Format(curDBIL / curTBIL, "0.000")
            End If
            '计算 AST/ALT
            If curAST <> 0 And curALT <> 0 Then
                strResult = strResult & "|106|" & Format(curAST / curALT, "0.000")
            End If
            
            '计算9/10
            If cur9 <> 0 And cur10 <> 0 Then
                '尿素/肌酐
                strResult = strResult & "|104|" & Format(cur9 / cur10, "0.000")
            End If
            
        End If
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
