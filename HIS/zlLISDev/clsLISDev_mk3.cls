VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MK3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const LABITEMS As String = ",pH,PCO2,PO2,pHt,PCO2t,PO2t,ctHb,O2Hb,HHb,COHb,MetHb,SulfHb,pHst,cHCO3st,cHCO3,ctCO2(P),ctCO2(B),BE,BEecf,BB,SO2,PAO2t,AaDO2t,a/AO2t,RI,"
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
    strCmd As String, strMBChannel As String, lngMBNo As Long, strMBPosition As String)
   ' Analyse(strInput, strResult, strReserved, strCmd, strMBChannel, lngMBNo, strMBPosition)
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
    Dim CutOff As Double '临界值
    Dim str标志 As String '大于小于标志
    Dim aItemData As Variant, blnReadItem As Boolean
    Dim i As Integer, j As Integer, intRow As Integer, intCol As Integer
    Dim iPos As Integer
    Dim strDate As String, lngSampleNO As Long, strItemValue As String
    Dim strTmp As String
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    strTmp = "    1     2     3     4     5     6     7     8     9    10    11    12 " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
  '  strSource = Replace(strSource, Chr(13) & Chr(10) & Chr(13) & Chr(10), "")
    strSource = Replace(strSource, strTmp, "")
    intRow = Asc(Mid(strMBPosition, 1, 1)) - 65 + 1
    intCol = Val(Mid(strMBPosition, 2, 2))
    
    '查找起始位置，NEGATIVE CONTROL作为开始,iStrPoint = 1:
    iStrPoint = InStr(strSource, "NEGATIVE CONTROL")
    strSource = Mid(strSource, iStrPoint)
    blnReadItem = False: str标志 = ""
    strDate = "": lngSampleNO = -1
    If Len(strSource) > 780 Then '数据未传完
        Do While iStrPoint <= Len(strSource)
            If Not blnReadItem Then '查找标本基本信息
                strBuffer = ReadLine(strSource, iStrPoint)
                Do While Len(strBuffer) > 0
                    Select Case True
                        Case InStr(UCase(strBuffer), UCase("CUT OFF")) > 0 '临界值
                            iPos = InStr(UCase(strBuffer), UCase("CUT OFF"))
                            CutOff = Val(Mid(strBuffer, 17, 5))
                            strBuffer = ""
                        Case InStr(UCase(strBuffer), UCase("INTERPRETATION")) > 0
                            If InStr(strBuffer, ">") > InStr(strBuffer, "+") Then
                                 str标志 = "阳性"
                            Else
                                 str标志 = "阴性"
                            End If
                            strBuffer = ""
                        Case InStr(UCase(strBuffer), UCase("ABSORBANCES")) > 0
                            blnReadItem = True
                            Exit Do
                        Case Else
                            strBuffer = ""
                    End Select
                Loop
            Else
                If str标志 <> "" Then  '基本信息查找完成，定位到检验项目开始位置
                    strBuffer = ReadLine(strSource, iStrPoint) '第九行，检验时间
                    strDate = Format(Replace(strBuffer, ".", ""), "yyyy-mm-dd hh:mm:ss")
                    strBuffer = ReadLine(strSource, iStrPoint) '第十行
               
                '标本信息开始，初始标本号为参数传入
                   For i = intRow To 8
                      strBuffer = ReadLine(strSource, iStrPoint) '第十一行，数据开始
                      aItemData = Split(strBuffer, " ")
                      For j = intCol To 12
                        lngSampleNO = lngMBNo + (i - intRow) * 12 + j - intCol
                       '时间格式必须为：yyyy-mm-dd hh:mm:ss
                       strData = strDate
                       '样本序号，必须为数字
                       strData = strData & "|" & lngSampleNO
                       '检验人
                       strData = strData & "| " '本仪器不支持
                       '标本
                       strData = strData & "|血液" '本仪器不支持
                       '质控品
                       strData = strData & "|0" '本仪器不支持，设为非质控品
                       
                       strData = strData & "|" & strMBChannel
                       
                       '考虑误差
                       If Abs(Val(aItemData(j))) < 0.004 Then Exit Do
                       If str标志 = "阳性" Then
                            If (Val(aItemData(j)) - CutOff) > 0 Then
                               strItemValue = "阳性"
                            Else
                               strItemValue = "阴性"
                            End If
                       Else
                            If (Val(aItemData(j)) - CutOff) > 0 Then
                               strItemValue = "阴性"
                            Else
                               strItemValue = "阳性"
                            End If
                       End If
                       
                       strData = strData & "|" & strItemValue
                       strResult = IIf(Len(strResult) > 0, strResult & "||", "") & strData
                   Next
                Next
              End If
           End If
        Loop
   Else
      iCurrGrpEntry = 1
   End If
    '返回未解析的原始串
    If iCurrGrpEntry > 0 Then
        strReserved = Mid(strSource, iCurrGrpEntry)
    Else
        '向设备发送‘数据收到’命令
        strCmd = ""
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

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '读取一行字符，并将字符指针移到下一行开始位置
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '读取一个检验项目名称，并将字符指针移到项目名称之后
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function









