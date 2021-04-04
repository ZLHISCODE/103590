VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_Bc5300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrCode As String  '待解析数据
Const mstrItem  As String = ",WBC,BAS#,BAS%,NEU#,NEU%,EOS#,EOS%,LYM#,LYM%,MON#,MON%,*ALY#,*ALY%,*LIC#,*LIC%,RBC,HGB,MCV,MCH,MCHC,RDW-CV,RDW-SD,HCT,PLT,MPV,PDW,PCT,"
' base 64 encoder string
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

'--------------------------------------------------------------------------------------------------------------------
'以下是必须的接口函数

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
    Dim str_Code As String '1个完整的消息
    Dim var_Code As Variant
    Dim lng_count As Long
    
    Dim str_字段分隔符 As String
    Dim str_组件分隔符 As String
    Dim str_子组件隔符 As String
    Dim str_重复分隔符 As String
    Dim str_转义字符 As String
    
    Dim str_Segment As String '一个消息段(以chr(&HD)分隔)
    Dim var_Filed As Variant   '字段
    Dim str_唯一标识 As String '用于标识一个消息
    Dim str_版本 As String
    
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_标本号 As String, str_日期 As String
    Dim str_质控 As String
    Dim str_项目 As String
    Dim str_结果 As String
    Dim str_Return As String
    
    Dim str_Base64 As String, str_WBC As String, str_RBC As String, str_PLT As String, str_img As String
    Dim str_WBCDIFF As String
    Dim str_消息类型 As String '用于5380的R01应答
    Dim str_仪器型号 As String
    Dim str_消息标识  As String
    
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
                 
        var_Code = Split(str_Code, Chr(&HD))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If Mid(str_Segment, 2, 3) = "MSH" Then
                '消息头
                str_字段分隔符 = Mid(str_Segment, 5, 1)
                str_组件分隔符 = Mid(str_Segment, 6, 1)
                str_子组件隔符 = Mid(str_Segment, 7, 1)
                str_重复分隔符 = Mid(str_Segment, 8, 1)
                str_转义字符 = Mid(str_Segment, 9, 1)
                
                var_Filed = Split(str_Segment, str_字段分隔符)
                str_唯一标识 = var_Filed(9)
                str_质控 = var_Filed(10)
                str_版本 = var_Filed(11)
                
                str_仪器型号 = var_Filed(3)
                str_消息标识 = var_Filed(9)
                str_消息类型 = var_Filed(10)
                
                If str_版本 <> "2.3.1" Then
                    WriteLog "版本不符", "", ""
                End If
                str_Return = "": str_标本号 = "": str_日期 = ""
                
                If str_仪器型号 = "BC-5380" And var_Filed(8) = "ORU^R01" Then
                    strCmd = Chr(&HB) & "MSH|^~\&|LIS||||" & Format(Now, "yyyyMMddHHmmss") & "||ORU^R02|" & str_消息标识 & "|" & str_消息类型 & "|2.3.1||||||UNICODE" & Chr(&HD) _
                           & "MSA|AA|" & str_消息标识 & Chr(&H1C) & Chr(&HD)
                End If
            ElseIf Left(str_Segment, 3) = "OBR" Then
                '标本号，日期
                If Val(str_标本号) > 0 And str_Return <> "" Then
                    strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|" & IIf(str_质控 = "Q", 1, 0) & str_Return
                    str_img = ""
                    If str_WBC <> "" Then str_img = str_img & "^WBC;103;1;" & str_WBC
                    If str_RBC <> "" Then str_img = str_img & "^RBC;103;1;" & str_RBC
                    If str_PLT <> "" Then str_img = str_img & "^PLT;103;1;" & str_PLT
                    If str_WBCDIFF <> "" Then str_img = str_img & "^DIFF;103;1;" & str_WBCDIFF
                    If str_img <> "" Then strResult = strResult & vbCrLf & Mid(str_img, 2)
                End If
                var_Filed = Split(str_Segment, str_字段分隔符)
                str_标本号 = var_Filed(3)
                str_日期 = var_Filed(7)
                str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2) & ":" & Mid(str_日期, 13, 2)
                str_Return = ""
            ElseIf Left(str_Segment, 3) = "OBX" Then
                var_Filed = Split(str_Segment, str_字段分隔符)
                str_项目 = var_Filed(3)
                If InStr(mstrItem, "," & Split(str_项目, str_组件分隔符)(1) & ",") > 0 Then
                    str_Return = str_Return & "|" & Split(str_项目, str_组件分隔符)(1) & "|" & Trim(var_Filed(5))
                ElseIf InStr(",15008,", "," & Split(str_项目, str_组件分隔符)(0) & ",") > 0 Then
                'wbc
                    str_Base64 = Split(Trim(var_Filed(5)), str_组件分隔符)(4)
                    str_WBC = SaveBase64Bmp("WBC", str_标本号, str_Base64)
                    
                ElseIf InStr(",15056,", "," & Split(str_项目, str_组件分隔符)(0) & ",") > 0 Then
                'RBC
                    str_Base64 = Split(Trim(var_Filed(5)), str_组件分隔符)(4)
                    str_RBC = SaveBase64Bmp("RBC", str_标本号, str_Base64)
                ElseIf InStr(",15116,", "," & Split(str_项目, str_组件分隔符)(0) & ",") > 0 Then
                'PLT
                    str_Base64 = Split(Trim(var_Filed(5)), str_组件分隔符)(4)
                    str_PLT = SaveBase64Bmp("PLT", str_标本号, str_Base64)
                ElseIf InStr(",15200,", "," & Split(str_项目, str_组件分隔符)(0) & ",") > 0 Then
                'WBC DIFF
                    str_Base64 = Split(Trim(var_Filed(5)), str_组件分隔符)(4)
                    str_WBCDIFF = SaveBase64Bmp("DIFF", str_标本号, str_Base64)
                End If
            End If
        Next
        If Val(str_标本号) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|" & IIf(str_质控 = "Q", 1, 0) & str_Return
            str_img = ""
            If str_WBC <> "" Then str_img = str_img & "^WBC;113;1;" & str_WBC
            If str_RBC <> "" Then str_img = str_img & "^RBC;113;1;" & str_RBC
            If str_PLT <> "" Then str_img = str_img & "^PLT;113;1;" & str_PLT
            If str_WBCDIFF <> "" Then str_img = str_img & "^DIFF;113;1;" & str_WBCDIFF
            If str_img <> "" Then strResult = strResult & vbCrLf & Mid(str_img, 2)
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "Bc5300_analyse", strSource, Err.Number & " " & Err.Description
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

'---------------------------------------------------------------------------------------------------------------
'以下是内部函数
Private Function SaveBase64Bmp(ByVal strType As String, ByVal str标本号 As String, ByVal str2Decode As String) As String

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
    Dim frmPic As New frmGraph
    Dim ByteData() As Byte, lngCount As Long, strFileBmp As String, lngFileNum
    Dim strFileJpg As String
    
    lngCount = Len(str2Decode)
    ReDim ByteData(lngCount / 4 * 3)
    lngCount = 0
    ' for each 4 character group....
    For lPtr = 1 To Len(str2Decode) Step 4
        iLen = 4
        For iCtr = 0 To 3
            ' retrive the base 64 value, 4 at a time
            iValue = InStr(1, BASE64CHR, Mid$(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
                Case 1 To 64: bits(iCtr + 1) = iValue - 1
                ' =
                Case 65
                    iLen = iCtr
                    Exit For
                ' not found
                Case 0: Exit Function
            End Select
        Next

        ' convert the 4, 6 bit values into 3, 8 bit values
        bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
        bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
        bits(3) = (bits(3) And &H3) * &H40 + bits(4)

        ' add the three new characters to the output string
        For iCtr = 1 To iLen - 1
            ByteData(lngCount) = bits(iCtr)
            lngCount = lngCount + 1
        Next
    Next
    
    strFileBmp = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    strFileJpg = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str标本号) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFileBmp For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    '-- 2012-05-08 图片反色处理
    frmPic.Picture1.AutoSize = True
    frmPic.Picture1.Picture = LoadPicture(strFileBmp)
    Call PicInvertAndSave(frmPic.Picture1, strFileJpg, "bmp")
    SaveBase64Bmp = strFileJpg

End Function
