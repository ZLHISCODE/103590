VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_AU680_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'仪器:AU680
'厂商:
'作者:王冬建
'用户:六安市中医院
Private mstrCode As String, mstrWaitSend As String, strRackNo As String

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
    
    Dim str_No As String, str_Date As String '标本号以及时间
    Dim str_Item As String, str_Value As String '项目以及结果
    Dim m As Long, n As Long, lngA As Long '解析结果时做循环用
    Dim str_Hc As String '存储临时数据，使解析界面清晰
    Dim str_Jz As String '急诊标志
    Dim str_条码 As String '条码
    Dim str_Tm As String
    
    str_Jz = ""
    
    '接收仪器传过来的数据
    mstrCode = mstrCode & strSource
'    If mstrCode Like "*" & Chr(&H20) & Chr(&H3) Then
'        strCmd = Chr(&H6)
'    End If
    
    If mstrCode Like Chr(&H2) & "RE" & Chr(&H3) & "*" Or mstrCode Like Chr(&H2) & "DE" & Chr(&H3) & "*" Then
        mstrCode = Mid(mstrCode, 5)
    End If
    
    If mstrCode Like Chr(&H2) & "DB" & Chr(&H3) & "*" Or mstrCode Like Chr(&H2) & "DB" & Chr(&H3) & "*" Then
        mstrCode = Mid(mstrCode, 5)
    End If
    
    '判断是否是一条完整的数据，否则不进行解析
    If mstrCode Like Chr(&H2) & "*" & Chr(&H3) Then
        '判断是仪器申请还是检验结果
        If mstrCode Like Chr(&H2) & "R*" Then
             '取出条码，并通知程序进行双向通讯
            lngA = UBound(Split(mstrCode, Chr(&H3)))
            For m = 0 To lngA
                str_Tm = Split(mstrCode, Chr(&H3))(m)
                If Len(str_Tm) > 10 Then
                    strRackNo = Mid(str_Tm, 4, 11)
                    str_No = Split(str_Tm, Chr(&H20))(2)
                    str_条码 = Trim(Mid(str_Tm, 15))
                    If str_No Like "E*" Then
                        str_Jz = 1
                        str_No = Val(Mid(str_No, 2))
                    End If
                    If str_条码 <> "" Then
                        str_No = ""
                    End If
                    strResult = strResult & "||" & Format(Now, "yyyy-mm-dd") & "|" & str_No & "^" & str_Jz & "^" & str_条码
                End If
            Next
            strCmd = "1|"
        Else
            '按照多组数据进行解析
            lngA = UBound(Split(mstrCode, Chr(&H3)))
            For m = 0 To lngA - 1
                str_Hc = Trim(Mid(Split(mstrCode, Chr(&H3))(m), 2))
                If Len(str_Hc) > 10 Then
                    '获取样本编号
                    str_No = Mid(str_Hc, 10, 4)
                    str_条码 = Trim(Mid(str_Hc, 14, 20))
'                    If InStr(str_No, Chr(&H20)) > 0 Then
'                        str_No = Trim(Mid(str_No, 1, InStr(str_No, Chr(&H20))))
'                        str_条码 = Trim(Mid(str_No, InStr(str_No, Chr(&H20))))
'                    End If
                    If Mid(str_No, 1, 1) = Chr(&H45) Then
                        str_Jz = 1
                        '去掉标本号前面的急诊标志
                        str_No = Mid(str_No, 2)
                        str_Hc = Mid(str_Hc, 39)
                    Else
                        str_Jz = 0
                        str_Hc = Mid(str_Hc, 39)
                    End If
                    
                    If str_条码 <> "" Then
                        str_No = ""
                    End If
                    strResult = strResult & "||" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "|" & Val(str_No) & "^" & str_Jz & "^" & str_条码 & "| |静脉血|0"
                    
                    '对结果进行解析
                    For n = 1 To 100
                        '每个结果总长度为11，由于截取时去空格了，所以当字符串剩余长度小于9时，解析结束
                        If Len(str_Hc) >= 9 Then
                            str_Item = Mid(str_Hc, 1, 11)
                            str_Hc = Mid(str_Hc, 12)
                            
                            strResult = strResult & "|" & Trim(Mid(str_Item, 1, 3)) & "|" & Val(Trim(Mid(str_Item, 4)))
                        Else
                            Exit For
                        End If
                    Next
                End If
            Next
        End If
        strResult = Mid(strResult, 3)
        '解析结束，清空缓存
        mstrCode = ""
    End If
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
    Dim blnEnd As Boolean
    Dim strBuffer As String, strSampleID As String
    Dim i As Long, aItem() As String
    SendSample = ""
    
    blnSuccess = True
    If iSendStep = 0 And strRackNo <> "" Then
        If strResponse Like "*R*" Or strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then '允许发送
        'Debug.Print "允许发送"
            If strResponse Like "**" Or Len(Trim(strResponse)) = 0 Then
                '请求结束
                '生成设备串
                aItem = Split(strSample, "|")
                '标本或SampleID
                strBuffer = aItem(1)
                strSampleID = aItem(5)
                '读取反馈的架号和杯号
                If Len(Trim(strSampleID)) > 0 Then
                    'SampleID                          'strData = strData & "|^^" & Mid(strSource, InStr(strSource, "R") + 23, 12)
                    SendSample = "S " & Right(Space(11) & strRackNo, 11) & Right(Space(20) & strSampleID, 20) & Space(4) & "E"
                    'Debug.Print strRackNO
                    
                    'SendSample = "S " & Right(Space(11) & strRackNO, 11) & Right(Space(20) & Mid(strResponse, InStr(strResponse, "R") + 24, 12), 20) & Space(4) & "E"
                Else
                    If strRackNo <> "" Then
                        SendSample = "S " & Right(Space(11) & strRackNo, 11) & Space(24) & "E"
                    Else
                        SendSample = "S         " & Right(String(4, "0") & strBuffer, 4) & Space(24) & "E"
                    End If
                End If
                'Debug.Print SendSample
                For i = 10 To UBound(aItem)
                    SendSample = SendSample & aItem(i)
                Next
                SendSample = SendSample & ""
                'SendSample = SendSample & "SE"
            'Debug.Print SendSample
            
                iSendStep = 0
                strRackNo = ""
            Else
                SendSample = Chr(6)
                iSendStep = iSendStep + 1
            End If
        Else
            SendSample = Chr(6)
            iSendStep = 0
            blnSuccess = False '发送失败
        End If
    Else
        SendSample = ""
    End If
End Function

Private Function GetCheckCode(ByVal str_in As String) As String
    '获取校验码
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function




