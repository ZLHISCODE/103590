VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BF6500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:DIRUI BF-6500
'厂商:长春迪瑞
'作者:周海全
'用户:
'日期:

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
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令


    Dim strCode As String, varLine As Variant
    Dim lngS As Long, lngE As Long
    Dim varData As Variant, strDay As String, strNo As String
    Dim iRow As Integer, strItem As String, strResultTmp As String
    Dim strImgRBC  As String, strImgPLT As String, strImgWBC As String, strImgDiff As String, strImgWBCD As String, strImgData As String
    Dim strNum As String
    '初始化返回值
    strResult = "": strReserved = "": strCmd = "": strDay = "": strNo = "":  strItem = ""
    strImgRBC = "": strImgPLT = "": strImgWBC = "": strImgDiff = "": strImgWBCD = ""
     
    mstrCode = mstrCode & strSource
    lngS = InStr(mstrCode, Chr(&HB))
    lngE = 0
    If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H1C))
    
    Do While lngS > 0 And lngE > 0
        strCode = Mid$(mstrCode, lngE - lngS + 1)
        mstrCode = Mid$(mstrCode, lngE + 1)
        '--------------------------------------------------------------------
        varData = Split(strCode, Chr(&HD))
        For iRow = 0 To UBound(varData)
            If varData(iRow) Like "OBR*Count Results*" Then
                varLine = Split(varData(iRow), "|")
                If strDay = "" Then
                    strDay = Mid(varLine(7), 1, 4) & "-" & Mid(varLine(7), 6, 2) & "-" & Mid(varLine(7), 9, 2) & " " & Mid(varLine(7), 11, 2) & ":" & Mid(varLine(7), 13, 2) & ":" & Mid(varLine(7), 15, 2)
                End If
                If strNo = "" Then
                    strNo = varLine(2)
                    strNum = strNo
                    If strNo Like "0-*" Then
                        strNo = Mid(strNo, InStr(strNo, "-") + 1) & "^1"
                        strNum = Mid(strNo, 1, InStr(strNo, "^") - 1)
                    End If
                End If
            ElseIf varData(iRow) Like "OBX*" Then
               varLine = Split(varData(iRow), "|")
               If varLine(1) > 5 And varLine(1) < 30 Then
                  If strItem = "" Then
                    strItem = varLine(1) & "|" & varLine(5)
                  Else
                    strItem = strItem & "|" & varLine(1) & "|" & varLine(5)
                  End If
               End If
               If varLine(1) >= 30 And varLine(1) <= 34 Then
                  If varLine(1) = 30 Then strImgRBC = Base64Decode(StrConv(varLine(5), vbNarrow), "RBC", strNum)
                  If varLine(1) = 31 Then strImgPLT = Base64Decode(StrConv(varLine(5), vbNarrow), "PLT", strNum)
                  If varLine(1) = 32 Then strImgWBC = Base64Decode(StrConv(varLine(5), vbNarrow), "WBC", strNum)
                  If varLine(1) = 33 Then strImgDiff = Base64Decode(StrConv(varLine(5), vbNarrow), "DIFF", strNum)
                  If varLine(1) = 34 Then strImgWBCD = Base64Decode(StrConv(varLine(5), vbNarrow), "WBCD", strNum)
               End If
               
               If varLine(1) = 34 And strNo <> "" And strDay <> "" And strItem <> "" Then
                  strImgData = ""
                  If strImgRBC <> "" Then strImgData = strImgData & "^RBC;100;1;" & strImgRBC
                  If strImgPLT <> "" Then strImgData = strImgData & "^PLT;100;1;" & strImgPLT
                  If strImgWBC <> "" Then strImgData = strImgData & "^WBC;100;1;" & strImgWBC
                  If strImgDiff <> "" Then strImgData = strImgData & "^DIFF;100;1;" & strImgDiff
                  If strImgWBCD <> "" Then strImgData = strImgData & "^WBCD;100;1;" & strImgWBCD
                  If strImgData <> "" Then strResultTmp = strDay & "|" & strNo & "| |Boold|0|" & strItem & vbCrLf & Mid(strImgData, 2)
                  
                  strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strResultTmp
                  strDay = "": strNo = "":  strItem = ""
               End If
            End If
        Next
            
        '--------------------------------------------------------------------
        lngS = InStr(mstrCode, Chr(&HB))
        lngE = 0
        If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H1C))
    Loop
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    ''   strSample:解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '           元素之间以|分隔
    '           第0个元素：检验时间
    '           第1个元素：样本序号
    '           第2个元素：检验人
    '           第3个元素：标本
    '           第4个元素：是否质控品
    '           第5～9元素：系统保留
    '           从第10个元素开始为需要的检验项目。
    '   iSendStep  : 步骤
    '   blnSuccess : 是否传输失败
    '   strResponse: 仪器的应答信息
    '   blnUndo    : 是否取消.
    '
    '   返回值：发到仪器的字串或指令。
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function

Public Function Base64Decode(InStr1 As String, ByVal strType As String, strNo As String) As String
   '生成图形数据
   Dim mInByte(4) As Byte, mOutByte(3) As Byte
   Dim i As Integer, LenArray As Integer, j As Integer
   Dim myBArray() As Byte
   Dim OutStr1 As String
   Dim tmpArray() As Byte
   Dim strFileName As String, lngFileNum As Long

   myBArray() = StrConv(InStr1, vbFromUnicode)
   LenArray = UBound(myBArray)
   ReDim tmpArray(((LenArray + 1) / 4) * 3)
    j = 0
    
    For i = 0 To LenArray Step 4
       If LenArray - i = 0 Then
           Exit For
       Else
            mInByte(0) = myBArray(i)
            mInByte(1) = myBArray(i + 1)
            mInByte(2) = myBArray(i + 2)
            mInByte(3) = myBArray(i + 3)
            Base64DecodeByte mInByte, mOutByte, 4
        End If
        tmpArray(j * 3) = mOutByte(0)
        tmpArray(j * 3 + 1) = mOutByte(1)
        tmpArray(j * 3 + 2) = mOutByte(2)
        j = j + 1
    Next i
    strFileName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & strNo & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFileName For Binary Access Write As lngFileNum
    Put lngFileNum, , tmpArray
    Close lngFileNum
    
    'SaveBmpToFile = strfilename
    Base64Decode = strFileName 'BinaryToString(tmpArray)
End Function

Private Sub Base64DecodeByte(mInByte() As Byte, mOutByte() As Byte, ByteNum As Integer)
    Dim tByte As Byte
    Dim i As Integer
    ByteNum = 0
    For i = 0 To 3
        If mInByte(i) >= Asc("A") And mInByte(i) <= Asc("Z") Then
            mInByte(i) = mInByte(i) - Asc("A")
        ElseIf mInByte(i) >= Asc("a") And mInByte(i) <= Asc("z") Then
            mInByte(i) = mInByte(i) - Asc("a") + 26
        ElseIf mInByte(i) >= Asc("0") And mInByte(i) <= Asc("9") Then
            mInByte(i) = mInByte(i) - Asc("0") + 52
        ElseIf mInByte(i) = Asc("+") Then
            mInByte(i) = 62
        ElseIf mInByte(i) = Asc("/") Then
            mInByte(i) = 63
        Else '"="
            ByteNum = ByteNum + 1
            mInByte(i) = 0
        End If
    Next i
    '取前六位
    tByte = (mInByte(0) And &H3F) * 4 + (mInByte(1) And &H30) / 16
    '0的六位和1的前两位
    mOutByte(0) = tByte
    tByte = (mInByte(1) And &HF) * 16 + (mInByte(2) And &H3C) / 4
    '1的后四位和2的前四位
    mOutByte(1) = tByte
    tByte = (mInByte(2) And &H3) * 64 + (mInByte(3) And &H3F)
    mOutByte(2) = tByte
    '2的后两位和3的六位
End Sub
Private Function BinaryToString(ByVal BinaryStr As Variant) As String '二进制转换为字符串
  Dim lnglen As Long
  Dim tmpBin As Variant
  Dim strC As String
  Dim skipflag As Long
  Dim i As Long
  skipflag = 0
  strC = ""
  
  If Not IsNull(BinaryStr) Then
      lnglen = LenB(BinaryStr)
      For i = 1 To lnglen
          If skipflag = 0 Then
            tmpBin = MidB(BinaryStr, i, 1)
            If AscB(tmpBin) > 127 Then
                strC = strC & Chr(AscW(MidB(BinaryStr, i + 1, 1) & tmpBin))
                skipflag = 1
            Else
                strC = strC & Chr(AscB(tmpBin))
            End If
          Else
            skipflag = 0
          End If
      Next
    End If
    BinaryToString = strC
End Function
