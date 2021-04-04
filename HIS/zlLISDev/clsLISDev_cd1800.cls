VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_cd1800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Cell Dyn 1800
'厂商:美国雅培
'用户:重庆市二院,德阳人民医院
'修改:2007-03-28 陈东
'+ 图形数据解析

Private mstrRBC As String, mstrWBC As String, mstrPLT As String '存图形数据。

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
    Dim i As Integer
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

    Dim iStrPoint As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strValues As String
    Dim lng_S As Long, lng_E As Long, str_Record As String, Str_Img As String
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    If strSource Like "***" Then
        
        strReserved = strSource
        
        Do
            lng_S = InStr(strReserved, "")
            lng_E = InStr(strReserved, "")
            str_Record = Mid(strReserved, lng_S, lng_E - lng_S + 1)
            
            strReserved = Mid(strReserved, lng_E + 1)
            
            If Mid(str_Record, 3, 3) = Space(3) Then
                '---- 解析数据
                
                iStrPoint = 1
                Do While iStrPoint <= Len(str_Record)
                    strBuffer = ReadSubString(str_Record, iStrPoint, 1)
                    '判断一组结果是否开始
                    If strBuffer = "" Then
                        '记录一组结果的入口
                        iCurrGrpEntry = iStrPoint - 1 '63530918
                        
                        iStrPoint = iStrPoint + 8
                        strBuffer = ReadSubString(str_Record, iStrPoint, 6)
                        If strBuffer = "CD1800" Then
                            If Len(str_Record) - iStrPoint >= 280 Then
                                '检验时间 '060520061632
                                iStrPoint = iStrPoint + 32
                                strData = ReadSubString(str_Record, iStrPoint, 2)   '月
                                iStrPoint = iStrPoint + 1
                                strData = strData & "-" & ReadSubString(str_Record, iStrPoint, 2) '日
                                iStrPoint = iStrPoint + 1
                                strData = ReadSubString(str_Record, iStrPoint, 2) & "-" & strData '年
                                iStrPoint = iStrPoint + 3
                                strData = strData & " " & ReadSubString(str_Record, iStrPoint, 5) '时分
                                '时间格式必须为：yyyy-mm-dd hh:mm:ss
                                strData = Format(strData, "yyyy-mm-dd hh:mm:ss")
                                
                                '样本序号，必须为数字 8位
                                iStrPoint = iStrPoint + 3 '定位到标本位
                               ' strSampleID = Trim(ReadSubString(strSource, iStrPoint, 11))
                                strData = strData & "|" & Val(ReadSubString(str_Record, iStrPoint, 16))
                                '检验人
                                strData = strData & "| " '本仪器不支持
                                '标本
                                strData = strData & "|血液" '本仪器不支持
                                '质控品
                                strData = strData & "|0" '本仪器不支持，设为非质控品
                                '检验项目
                                iStrPoint = iStrPoint + 95
                                For i = 1 To 21
                                    strData = strData & "|" & i
                                    iStrPoint = iStrPoint + 1
                                    Select Case i
                                        Case 1, 2, 3, 4, 5, 6, 9, 10, 11, 15, 17, 13
                                            strValues = ReadSubString(str_Record, iStrPoint, 5)
                                            If Val(strValues) <> 0 Then
                                                strData = strData & "|" & Val(strValues) / 10
                                            Else
                                                strData = strData & "|----"
                                            End If
                                        Case 8, 12, 14, 18
                                            strValues = ReadSubString(str_Record, iStrPoint, 5)
                                            If Val(strValues) <> 0 Then
                                                strData = strData & "|" & Val(strValues)
                                            Else
                                                strData = strData & "|----"
                                            End If
                                           
                                        Case 19, 21, 20
                                           
                                            strValues = ReadSubString(str_Record, iStrPoint, 5)
                                            If Val(strValues) <> 0 Then
                                                strData = strData & "|" & Val(strValues) / 1000
                                            Else
                                                strData = strData & "|----"
                                            End If
                                         Case 7, 16
                                            strValues = ReadSubString(str_Record, iStrPoint, 5)
                                            If Val(strValues) <> 0 Then
                                                strData = strData & "|" & Val(strValues) / 100
                                            Else
                                                strData = strData & "|----"
                                            End If
                                    End Select
                                    '该组结果已解析
                                    iCurrGrpEntry = 0
                                Next
                                strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                            End If
                        Else
                            '该组结果无效
                            If Len(strBuffer) = 2 Then iCurrGrpEntry = 0
                        End If
                        '指针定位到下一个换行之后
                        MoveToNextLine str_Record, iStrPoint
                    End If
                Loop   '解析数据
                
                
                If strResult <> "" Then
                    ''有结果，把图形数据附在后面。
                    Str_Img = ""
                    
                    If mstrWBC <> "" Then
                        If Split(mstrWBC, "|")(0) = Mid(str_Record, 46, 38) Then
                            '日期，标本号相同，附加图形数据
                            Str_Img = Str_Img & "^" & Split(mstrWBC, "|")(1)
                            mstrWBC = ""
                        End If
                    End If
                    
                    If mstrRBC <> "" Then
                        If Split(mstrRBC, "|")(0) = Mid(str_Record, 46, 38) Then
                            '日期，标本号相同，附加图形数据
                            Str_Img = Str_Img & "^" & Split(mstrRBC, "|")(1)
                            mstrRBC = ""
                        End If
                    End If
                    
                    If mstrPLT <> "" Then
                        If Split(mstrPLT, "|")(0) = Mid(str_Record, 46, 38) Then
                            '日期，标本号相同，附加图形数据
                            Str_Img = Str_Img & "^" & Split(mstrPLT, "|")(1)
                            mstrPLT = ""
                        End If
                    End If
                    If Str_Img <> "" Then strResult = strResult & vbCrLf & Mid(Str_Img, 2)
                End If
            
            ElseIf Mid(str_Record, 3, 3) = "WBC" Then
                mstrWBC = Mid(str_Record, 46, 38) & "|WBC;0;120,262,43.33-,86.66-100,129.996-,173.326-200,216.656-,259.986-300;" & Replace(Mid(str_Record, 179, 1029), ",", ";")
            ElseIf Mid(str_Record, 3, 3) = "RBC" Then
                mstrRBC = Mid(str_Record, 46, 38) & "|RBC;0;120,262,52-,104-100,156-,208-200,260-;" & Replace(Mid(str_Record, 179, 1029), ",", ";")
            ElseIf Mid(str_Record, 3, 3) = "PLT" Then
                mstrPLT = Mid(str_Record, 46, 38) & "|PLT;0;120,262,65-10,130-20,195-30,260-;" & Replace(Mid(str_Record, 179, 1029), ",", ";")
            End If
            
            If InStr(strReserved, "") <= 0 Then Exit Do
        Loop '传入串已全部解析完
        'strCmd = Chr(&H6)
    Else                    '数据不全，下次解析
        'strCmd = Chr(&H6)
        strReserved = strSource
        Exit Sub
    End If
    
    
''    '没有找到结束位时退出
''    If InStr(strSource, "") = 0 Then
''        strReserved = strSource
''        Exit Sub
''    End If
''
''    '返回未解析的原始串
''    If iCurrGrpEntry > 0 Then
''        strReserved = Mid(strSource, iCurrGrpEntry)
''    Else
''        '向设备发送‘数据收到’命令
''         strCmd = Chr(&H6&)
''    End If
    
    
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
        If Mid(strSource, iPoint, 1) = "]" Then
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
