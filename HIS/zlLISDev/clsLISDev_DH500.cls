VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_DH500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:DH500
'厂商:武汉中达
'作者:黄金桥
'用户:成都东区医院

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
    
    Dim varLine As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, iCount As Integer, lngLength As Long
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = "": strValues = ""
    
    If Not (strSource Like "20*/*/*" & Chr(&HD) & "*" & Chr(&HA) & "*") Then
       strReserved = strSource
       Exit Sub
    End If
    
    If InStr(strSource, Chr(&HA)) > 0 Then
        blnEnd = False
        strValues = strSource & Chr(&HA)
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            lngLength = InStr(strValues, Chr(&HA)) - InStr(strValues, Chr(&H32))
            strLine = Mid(strValues, InStr(strValues, Chr(&H32)), lngLength)
            strValues = Mid(strValues, InStr(strValues, Chr(&HA)) + 1)       '一条完整的传输数据取完，后面的数据存到strReserved中下次再解析
            
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            '解析
            strLine = Format(Mid(strLine, 1, 16), "yyyy-MM-dd HH:MM") & ":" & Format(Now, "SS") & Mid(strLine, 17)
            strLine = Replace(strLine, "/", "-")
            strLine = Replace(strLine, "(QC-1) ", "| |血清|1|")
            strLine = Replace(strLine, "(B) ", "| |血清|0|")
            strLine = Replace(strLine, "(U) ", "| |血清|0|")
            strLine = Replace(strLine, "(H)", "")
            strLine = Replace(strLine, "(L)", "")
            strLine = Replace(strLine, " ID=", "|")
            strLine = Replace(strLine, "= ", "|")
            strLine = Replace(strLine, "=", "|")
            strLine = Replace(strLine, "(?)", "")
            strLine = Replace(strLine, "< ", "|<")
            strLine = Replace(strLine, "> ", "|>")
            strLine = Replace(strLine, " Na", "|Na")
            strLine = Replace(strLine, " Cl", "|Cl")
            strLine = Replace(strLine, " Ca", "|Ca")
            strLine = Replace(strLine, " pH", "|pH")
            strLine = Replace(strLine, " n-Ca", "|n-Ca")
            strLine = Replace(strLine, " T-Ca", "|T-Ca")
            strResult = Trim(strResult) & "||" & Trim(strLine)
            If InStr(strValues, Chr(&HA)) = 0 Or Len(strValues) < 5 Then
                blnEnd = True
            End If
        Loop
   
        strResult = Replace(strResult, Chr(&HA), "")
        strResult = Replace(strResult, Chr(&HD), "")
        If strResult <> "" Then
           If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
        End If
    End If
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
