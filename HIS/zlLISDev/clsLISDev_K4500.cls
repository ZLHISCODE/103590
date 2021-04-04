VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_K4500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:K-4500
'厂商:希森美康
'作者:陈东
'用户:南京江北人民医院

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = Chr(6)
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
    
    Dim varLine As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, lngLength As Long, strRecord As String, lngByte As Long
    Dim str_Code As String, str_Day As String, StrWbc As String, StrRbc As String, StrPLt As String, strBit As String
    
    Dim blnEnd As Boolean '是否已解析完
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    
    strCmd = ",06" '应答为二进制
    '----  取一个完整的串
    If Len_Bin(strSource) < 456 Then
        strReserved = strSource
        Exit Sub
    End If
    
    For lngByte = 1 To Len_Bin(strSource) - 4
        str_Code = Mid_bin(strSource, lngByte, 1)
        str_Day = Mid_bin(strSource, lngByte + 3, 1)
    
        If str_Code = "" And str_Day = "/" Then
    
            If Len_Bin(strSource) - lngByte < 456 Then
                strReserved = strSource
                Exit Sub
            End If
    
            strLine = Mid_bin(strSource, lngByte, 457, False)
            strReserved = Mid_bin(strSource, lngByte + 457, False)
            Exit For
        End If
    Next
    
    If lngByte = Len_Bin(strSource) - 3 Then
        Exit Sub
    End If
    '-- 开始解析 strLine
    str标本号 = Mid_bin(strLine, 15, 13)
    str日期 = Mid_bin(strLine, 2, 2)
    str日期 = str日期 & "-" & Mid_bin(strLine, 5, 2)
    str日期 = str日期 & "-" & Mid_bin(strLine, 8, 2)
    If str标本号 <> 0 And IsDate(str日期) Then
        strResult = "||" & Format(str日期, "yyyy-MM-dd") & "|" & Val(str标本号) & "| |血液|0"
        For lngByte = 1 To 21
            strItem = Mid_bin(strLine, 78 + lngByte * 5, 5)
        
            If ((((lngByte = 8) Or (lngByte = 9)) Or (lngByte = 10)) Or (lngByte = 14)) Then
                strItem = Format(Val(strItem) * 100, "##0.0")
            End If
        
            strResult = strResult & "|" & lngByte & "|" & strItem
        Next
        
        'WBC
        StrWbc = ""
        For lngByte = 243 To 243 + 50
            strBit = Mid_bin(strLine, lngByte, 1, False)
            StrWbc = StrWbc & ";" & Format(CByte("&H" & Replace(strBit, ",", "")), "000")
        Next
        'RBC
        For lngByte = 293 To 293 + 50
            strBit = Mid_bin(strLine, lngByte, 1, False)
            StrRbc = StrRbc & ";" & Format(CByte("&H" & Replace(strBit, ",", "")), "000")
        Next

        'PLT
        For lngByte = 343 To 343 + 40
            strBit = Mid_bin(strLine, lngByte, 1, False)
            StrPLt = StrPLt & ";" & Format(CByte("&H" & Replace(strBit, ",", "")), "000")
        Next
    End If
    
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3) & _
           Chr(13) & Chr(10) & _
          "WBC;0;100,50,8-,17-100,25-,33-200,42-,50-300" & StrWbc & "^" & _
          "RBC;0;100,50,10-,20-100,30-,40-200" & StrRbc & "^" & _
          "PLT;0;100,40,10-10,20-20,30-30" & StrPLt
        
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



