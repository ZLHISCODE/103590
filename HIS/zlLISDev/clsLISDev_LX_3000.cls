VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LX_3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'仪器:LX3000
'厂商:
'作者:曾超
'用户:永州人民医院

Private mstrCode As String    '待解析串
Private mlngStep_In As Long   '输入字符串解析步骤
Private mstrACK As String     '是否应答ACK,是则应答ACK(06),否-应答 ETX(03)

Private mlngStep_Out As Long  '传双向消息步骤
Private mstr仪器号 As String

Private mstrWaitSend As String '待发送标本


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
    Dim lng_S As Long, lng_E As Long, i_Count As Integer
    Dim str_in As String, str_Line As String, str_条码 As String
    Dim str_日期 As String, str_标本号 As String, str_项目 As String, str_结果 As String
    Dim str_WaitSendLine As String
    Dim strNow As String
    Dim strData As String
    Dim strImage As String
    Dim intLoop As Integer
    Dim aLine() As String
    Dim aItem() As String
    Dim str标本号 As String
    Dim strImagePath As String
    Dim strLine As String
    Dim strFilePaht As String
    Dim strValue As String
    Dim strItem As String
    Dim str日期 As String

    mstrCode = mstrCode & strSource
    If InStr(mstrCode, Chr(&H38) & Chr(&H0) & Chr(&HCCCC) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0)) > 0 Then
        strCmd = "OK"
    End If
    
    '开始符
    If InStr(mstrCode, Chr(&H9) & Chr(&H0) & Chr(&H0) & Chr(&H0)) <= 0 Then
        Exit Sub
    End If
    
    '结束符
    If InStr(mstrCode, Chr(&H2E) & Chr(&H6A) & Chr(&H70) & Chr(&H67)) <= 0 Then
        Exit Sub
    End If
    
    strImagePath = ReadIni("LX_3000", "图像路径", App.Path & "\LX_3000.ini")
    If strImagePath = "" Then strImagePath = "c:\LX3000\":   Call WriteIni("LX_3000", "图像路径", strImagePath, App.Path & "\LX_3000.ini")
    
    
    '处理结果数据
    '首先转换为字符格式
'    strData = Mid(strSource, 1, Len(strSource) - 1)
    lng_E = InStr(mstrCode, Chr(&H2E) & Chr(&H6A) & Chr(&H70) & Chr(&H67)) + 3
    lng_S = InStr(mstrCode, Chr(&H9) & Chr(&H0) & Chr(&H0) & Chr(&H0))
    
    strData = Mid(mstrCode, lng_S, lng_E - lng_S)
    mstrCode = Mid$(mstrCode, lng_E + 1)
    
    aLine = Split(strData, vbCrLf)
    For intLoop = 0 To UBound(aLine)
        If InStr(aLine(intLoop), "EXAMINEDATE:") > 0 And UBound(Split(aLine(intLoop), ":")) > 5 Then
            aItem = Split(aLine(intLoop), ":")
            
            
            strItem = Trim(Replace(aItem(5), "ITEMVALUE", ""))
            If strItem <> "" Then
                strResult = strResult & "||" & Trim(Mid(Trim(aItem(1)), 1, 10)) & "|" & Trim(Mid(Trim(aItem(3)), 1, 5)) & "| |尿|0|"
                strResult = strResult & Trim(Replace(aItem(5), "ITEMVALUE", "")) & "|"
                
                If Trim(Replace(aItem(7), "ITEMDESC", "")) = "" Then
                    strValue = Trim(Replace(aItem(6), "ITEMCHAR", " "))
                Else
                    strValue = Trim(Replace(aItem(7), "ITEMDESC", " "))
                End If
                If strValue = "" Then strValue = " "
                strResult = strResult & strValue
                
'                strImage = Trim(Replace(aItem(9), "DELETEMARK", "")) & ".jpg"
                str标本号 = Trim(Mid(Trim(aItem(3)), 1, 5))
                str日期 = Trim(Mid(Trim(aItem(1)), 1, 10))
            End If
        End If
    Next
    strResult = Mid(strResult, 3)

    If str标本号 <> "" Then
        '处理图像
        strImage = Mid(strData, InStr(strData, ".jpg") - 20, 25)
        strImage = Mid(strImage, InStr(strImage, "2"))
        strImage = Mid(strImage, 1, InStr(strImage, "g"))
        strImage = strImagePath & Format(str日期, "yyyymm") & "\" & Format(str日期, "yyyymmdd") & "\" & Trim(strImage)
    End If

    If strImage <> "" Then
        strResult = strResult & vbCrLf & "A;110;1;" & strImage
    End If
    
    strReserved = ""
End Sub


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
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
    Dim str_SendItem As String
    
    
    
End Function


Private Function ACKChar() As String
    If mstrACK = Chr(&H6) Then
        mstrACK = Chr(&H3)
    Else
        mstrACK = Chr(&H6)
    End If
    ACKChar = mstrACK
End Function

Private Function SaveBase64Bmp(ByVal str标本号 As String, strCode As String) As String

' ******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
' ascii character equivalent. Stop converting at the end of the input string
' or when the first '=' (equal sign) is encountered.
'
' ******************************************************************************

    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim bits(1 To 4)    As Byte
    
    Dim ByteData() As Byte, lngCount As Long, strfilename As String, lngFileNum
    Dim aTmp() As String
    
    aTmp = Split(strCode, ",")
    
    ReDim ByteData(UBound(aTmp))
    
    For lngCount = 0 To UBound(aTmp) - 1
        ByteData(lngCount) = "&H" & aTmp(lngCount)
    Next
        
    
    
    strfilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_1" & ".jpg"
    lngFileNum = FreeFile
    Open strfilename For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    SaveBase64Bmp = strfilename

End Function
