VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_OMNI"
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
    
    Dim iStrPoint As Integer '字符指针
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer '当前这组结果的入口字符位置
    
    Dim blnReadItem As Boolean '当前是否正在读取检验项目
    Dim iPos As Integer
    Dim strDate As String, lngSampleNO As Long, strItemName As String, dblItemValue As Double
    '初始化返回值
    strReserved = "": strResult = "": strCmd = ""
    
    iStrPoint = 1: blnReadItem = False
    strDate = "": lngSampleNO = -1
    Do While iStrPoint <= Len(strSource)
        If Not blnReadItem Then '查找标本基本信息
            strBuffer = ReadLine(strSource, iStrPoint)
            Do While Len(strBuffer) > 0
                Select Case True
                    Case InStr(UCase(strBuffer), UCase("Meas Date/Time")) > 0 '时间
                        iPos = InStr(UCase(strBuffer), UCase("Meas Date/Time"))
                        strDate = Trim(Mid(strBuffer, iPos + 15, 21))
                        '日期格式为：dd.mm.yyyy hh:mm
                        strDate = Format(CDate(Replace(strDate, ".", "/")), "yyyy-mm-dd hh:mm:ss")
                        
                        strBuffer = Mid(strBuffer, iPos + 36)
                    Case InStr(UCase(strBuffer), UCase("Patient ID")) > 0 '标本号
                        iPos = InStr(UCase(strBuffer), UCase("Patient ID"))
                        lngSampleNO = Val(Mid(strBuffer, iPos + 13, 10))
                        
                        strBuffer = Mid(strBuffer, iPos + 23)
                    Case Else
                        strBuffer = ""
                End Select
            Loop
            If strDate <> "" And lngSampleNO > 0 Then '基本信息查找完成，定位到检验项目开始位置
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
                
                Do While iStrPoint <= Len(strSource)
                    strBuffer = ReadLine(strSource, iStrPoint)
                    If InStr(UCase(strBuffer), UCase("Meas. val. at")) > 0 Then Exit Do
                Loop
                blnReadItem = True
            End If
        Else '检验项目
            strBuffer = ReadLine(strSource, iStrPoint)
            Do While Len(strBuffer) > 0
                iPos = 1
                strItemName = ReadItemName(strBuffer, iPos)
                strBuffer = Replace(LTrim(Mid(strBuffer, iPos)), "%", "")
                
                If InStr(UCase(LABITEMS), "," & UCase(strItemName) & ",") > 0 Then '合法的检验项目
                    strData = strData & "|" & strItemName
                    dblItemValue = Val(strBuffer)
                    strData = strData & "|" & dblItemValue
                    
                    strBuffer = Replace(strBuffer, CStr(dblItemValue), "", 1, 1)
                End If
            Loop
        End If
    Loop
            
    strResult = strData
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
    
    iPos = InStr(iPoint, strSource, Chr(13) + Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 2
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '读取一个检验项目名称，并将字符指针移到项目名称之后
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

