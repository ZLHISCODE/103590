VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_IQ200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:全自动尿沉渣分析仪 IQ200
'厂商:美国Iris
'作者:李建飞
'用户:南京市江宁区中医院(陈玉强)

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
    
    Dim varline As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim varTmp As Variant, intLoop As Integer
    Dim iRow As Integer, strTmp As String, strLine As String, str项目 As String
    Dim lng_S As Long, lng_E As Long
    Dim strData As String, strNode As String
    Dim strValues As String, lngLength As Long
    Dim str标本类型 As String
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)

    mstrCode = mstrCode & strSource
    
    If mstrCode Like "***" Then
        blnEnd = False
'        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            lngLength = InStr(mstrCode, "") - InStr(mstrCode, "")
            strData = Mid(mstrCode, InStr(mstrCode, ""), lngLength)
            mstrCode = Mid(mstrCode, InStr(mstrCode, "") + 1)
            
            str日期 = "": str标本号 = "":  strItem = "": strTmp = ""
            '解析
            '第一步：将收到的16进制数据转为 XML 格式
            varline = Split(strData, vbCrLf)
            strItem = "": strData = ""
            For iRow = LBound(varline) To UBound(varline) - 1
                lng_S = InStr(varline(iRow), "") + 2
                lng_E = IIf(InStr(varline(iRow), "") > 0, InStr(varline(iRow), ""), InStr(varline(iRow), ""))
                
                strLine = Mid(varline(iRow), lng_S, lng_E - lng_S)
                strData = strData & GetStr(strLine)
            Next
            
            '第二步：解析 XML 文件
            strData = Replace(Replace(strData, Chr(&HD), ""), Chr(&HA), "")     '替换回车换行符
            
            '解析标本号和检验日期
            strData = Mid(strData, InStr(strData, "<SA"))
            strNode = Mid(strData, 1, InStr(strData, ">"))
                
            varTmp = Split(strNode, " ")
            For intLoop = LBound(varTmp) To UBound(varTmp)
                varItem = Split(varTmp(intLoop), "=")
                If varItem(0) = "ID" Then str标本号 = Replace(varItem(1), """", "")
                If varItem(0) = "REDTS" Then str日期 = Format(Replace(varItem(1), """", ""), "yyyy-MM-dd")
            Next
            
            '解析检验项目和检验结果
            strData = Mid(strData, InStr(strData, "<AR "))
            lng_S = InStr(strData, "<AR ")
            lng_E = InStr(strData, "</AR>") + 4
            
            While lng_S > 0 And lng_E > 0
                strData = Mid(strData, lng_S)
                strNode = Mid(strData, 1, lng_E)
                strData = Mid(strData, lng_E + 1)
                
                '分两部分处理
                '项目
                strTmp = Mid(strNode, InStr(strNode, "<"), InStr(strNode, ">") - InStr(strNode, "<") + 1)
                strNode = Mid(strNode, InStr(strNode, ">") + 1)
                varTmp = Split(strTmp, " ")
                For intLoop = LBound(varTmp) To UBound(varTmp)
                    varItem = Split(varTmp(intLoop), "=")
                    If varItem(0) = "Key" Then strItem = strItem & "|" & Replace(varItem(1), """", "")
                Next
                '结果
                strTmp = Mid(strNode, 1, InStr(strNode, "<") - 1)
                If strTmp = "[none]" Then
                    strItem = strItem & "|无"
                ElseIf InStr(strTmp, "/") > 0 Then
                    strItem = strItem & "|" & Trim(Replace(Replace(Split(strTmp, "/")(0), "&gt;", ">"), "&lt;", "<"))
                Else
                    strItem = strItem & "|几乎没有"
                End If
                
                lng_S = InStr(strData, "<AR ")
                lng_E = InStr(strData, "</AR>") + 4
            Wend
            
            If strItem <> "" And Val(str标本号) <> 0 Then
                strResult = strResult & "||" & str日期 & "|" & str标本号 & "| |尿液|0" & strItem
            End If
            If Not mstrCode Like "" Then
                blnEnd = True
            End If
        Loop

    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
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

'将16进制字符串转换为字符格式
Public Function GetStr(ByVal strBin As String) As String
    Dim strTmp As String
    Dim strResult As String
    Dim intLoop As Integer
    
    For intLoop = 1 To Len(strBin) Step 4
        strTmp = Mid(strBin, intLoop, 2)
        strResult = strResult & Chr("&H" & strTmp)
    Next
    GetStr = strResult
End Function


