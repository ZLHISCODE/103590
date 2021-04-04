VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ADVIA_2120i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:ADVIA 2120i全自动血细胞分析仪
'厂商:拜尔
'作者:周海全
'用户:

Private mstrCode As String
Private mblnSendInit As Boolean

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    
    mblnSendInit = True
    
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function
Public Function GetAnswerCmd() As String
    
    Dim strCheckBit As String
    If mblnSendInit Then
        GetAnswerCmd = Chr(&H30) & Chr(&H49) & Space(1) & Chr(&H7C) & "001" & Chr(&H7C) & Chr(&H2F) & Chr(&H78) & Chr(&H2F) & Chr(&H7C) & Chr(&HD) & Chr(&HA)
        strCheckBit = GetCheckSum(GetAnswerCmd)
        GetAnswerCmd = Chr(&H2) & GetAnswerCmd & Chr("&H" & strCheckBit) & Chr(&H3)
        mblnSendInit = False
    End If
End Function
Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
        
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号^是否急诊^条码号  可以只有样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令

    Dim strResponseFirst As String
    Dim strResponseThird As String
    Dim strCheckBit As String
    Dim strResponse As String
    Dim strMsgTag As String
    Dim strTmp As String, varTmp As Variant
    Dim row As Integer, Rows As Integer
    Dim lngS As Long, lngE As Long, strLine As String
    
    Dim str_标本号 As String, str_检验日期 As String, str_标本类型 As String
    Dim str_检验项目 As String
    

    mstrCode = mstrCode & strSource
    
    If Trim(strSource) = Chr(&H30) Then
        '等于0时发送固定字符串
        strCmd = "1S" & Space(10) & Chr(&HD) & Chr(&HA)
        strCheckBit = GetCheckSum(strCmd)
        strCmd = Chr(&H2) & strCmd & Chr("&H" & strCheckBit) & Chr(&H3)
    ElseIf InStr(strSource, Chr(&H15)) > 0 Then
        mblnSendInit = True
    ElseIf mstrCode Like "***" Then
        
        If mblnSendInit = True Then mblnSendInit = False
        
        lngS = InStr(mstrCode, Chr(&H2))
        lngE = 0
        If lngS > 0 Then lngE = InStr(mstrCode, Chr(&H3))
        If lngS > 0 And lngE > 0 Then
            strLine = Mid(mstrCode, lngS, lngE - lngS + 1)
            mstrCode = Mid(mstrCode, lngE + 1)
            If Mid(strLine, 3, 11) = "S" & Space(10) Then
                '应答
                strCmd = Chr(Asc(Mid(strLine, 2, 1)) + 1)
                If strCmd = "Q" Then
                    strCmd = "P"
                    mblnSendInit = True
                Else
                    strCmd = strCmd & "S" & Space(10) & Chr(&HD) & Chr(&HA)
                    strCheckBit = GetCheckSum(strCmd)
                    strCmd = Mid(strLine, 2, 1) & Chr(&H2) & strCmd & Chr("&H" & strCheckBit) & Chr(&H3)
                End If
            ElseIf Mid(strLine, 3, 1) = "R" Then
                '检验结果
                str_标本号 = Val(Mid(strLine, 6, 14))
    
                '04/11/10 17:23:23
                str_检验日期 = Mid(strLine, 44, 2) & "-" & Mid(strLine, 38, 2) & "-" & Mid(strLine, 41, 2)
                str_检验日期 = Format(str_检验日期, "yyyy-MM-dd")
    
                varTmp = Split(strLine, Chr(&HD) & Chr(&HA))
                strTmp = varTmp(1): Rows = Len(strTmp) / 9
    
                For row = 1 To Rows
                    str_检验项目 = str_检验项目 & "|" & Trim(Mid(strTmp, 1, 3)) & "|" & Trim(Mid(strTmp, 4, 6))
                    strTmp = Mid(strTmp, 10)
                Next
                
                strResponseThird = Chr(Asc(Mid(strLine, 2, 1)) + 1)
                strResponseFirst = Mid(strLine, 2, 1)
                
                strCmd = strResponseThird & Chr(&H5A) & Space(18) & Chr(&H30) & Chr(&HD) & Chr(&HA)
                strCheckBit = GetCheckSum(strCmd)
                strCmd = strResponseFirst & Chr(&H2) & strCmd & Chr("&H" & strCheckBit) & Chr(&H3)
            
            End If
        
        End If

    End If
    
    If Val(str_标本号) <> 0 And str_检验项目 <> "" Then
        strResult = str_检验日期 & "|" & str_标本号 & "| |全血|0" & str_检验项目
    End If
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '   strSample:解析需要向仪器发送的标本信息
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
    SendSample = ""
End Function

Private Function GetCheckSum(ByVal strResponse As String) As String
    Dim i As Integer
    Dim lngLength As Long
    Dim lngResult As Long

    lngLength = Len(strResponse)
    lngResult = Asc(Left(strResponse, 1))
    
    For i = 2 To lngLength
        lngResult = lngResult Xor Asc(Mid(strResponse, i, 1))
    Next
    
    If lngResult = 3 Then
        lngResult = 127
    End If
    
    GetCheckSum = Hex(lngResult)
End Function


