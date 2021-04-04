VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_RAC100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:RAC100 血凝分析仪
'厂商:丹麦雷度
'作者:陈德荣
'用户:汉源县人民医院

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
    Dim str_Code As String, mstrCode As String
    Dim var_Code As Variant
    Dim lng_count As Integer
    Dim blnEnd As Boolean '是否已解析完
    Dim str_Segment As String '一个消息段
    
    Dim lng_S As Integer
    Dim lng_E As Integer
    
    Dim str_标本号 As String, str_日期 As String, strRet As String
    Dim str_项目 As String, str_结果 As String
    Dim str_Return As String, citem13 As String
    Dim sampleno As Integer, sampledate As Integer, Item As Integer
    Dim citem14 As String, citem15 As String, citem16 As String
    Dim citem17 As String, citem18 As String, strTmp As String
    Dim item14 As Integer, item15 As Integer, item16 As Integer
    Dim item17 As Integer, item18 As Integer, item13 As Integer
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream
    
    On Error GoTo errHandle
    strResult = "": strCmd = "": strRet = ""
    strIniFile = App.Path & "\RAC100.ini"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "[Options]"
       objStream.WriteLine "sampleno=1  ;标本号位置"
       objStream.WriteLine "sampledate=2 ;日期位置"
       objStream.WriteLine "Item=11      ;项目通道号位置"
       objStream.WriteLine "item13=12,       ;s结果位置"
       objStream.WriteLine "item14=13,(RATIO);ratio结果位置"
       objStream.WriteLine "item15=14,(INR)  ;INR结果位置"
       objStream.WriteLine "item16=15,(CON)  ;CON结果位置"
       objStream.WriteLine "item17=16,(%)    ;%结果"
       objStream.WriteLine "item18=17,(DFBG) ;dfbg结果"
       objStream.Close
       Set objStream = Nothing
    End If
    sampleno = Val(ReadIni("Options", "sampleno", strIniFile, 1))
    sampledate = Val(ReadIni("Options", "sampledate", strIniFile, 2))
    
    Item = Val(ReadIni("Options", "Item", strIniFile, 11))
    citem13 = ReadIni("Options", "item13", strIniFile, "12,")
    citem14 = ReadIni("Options", "item14", strIniFile, "13,(RATIO)")
    
    citem15 = ReadIni("Options", "item15", strIniFile, "14,(INR)")
    citem16 = ReadIni("Options", "item16", strIniFile, "15,(CON)")
    
    citem17 = ReadIni("Options", "item17", strIniFile, "16,(%)")
    citem18 = ReadIni("Options", "item18", strIniFile, "17,(DFBG)")
    
    citem13 = Trim(Replace(citem13, ";s结果位置", ""))
    citem14 = Trim(Replace(citem14, ";ratio结果位置", ""))
    citem15 = Trim(Replace(citem15, ";INR结果位置", ""))
    citem16 = Trim(Replace(citem16, ";CON结果位置", ""))
    citem17 = Trim(Replace(citem17, ";%结果", ""))
    citem18 = Trim(Replace(citem18, ";dfbg结果", ""))
    
    item13 = Val(Mid(citem13, 1, InStr(citem13, ",")))
    citem13 = Trim(Mid(citem13, InStr(citem13, ",") + 1))
    
    item14 = Val(Mid(citem14, 1, InStr(citem14, ",")))
    citem14 = Trim(Mid(citem14, InStr(citem14, ",") + 1))
    item15 = Val(Mid(citem15, 1, InStr(citem15, ",")))
    citem15 = Trim(Mid(citem15, InStr(citem15, ",") + 1))
    
    item16 = Val(Mid(citem16, 1, InStr(citem16, ",")))
    citem16 = Trim(Mid(citem16, InStr(citem16, ",") + 1))
    item17 = Val(Mid(citem17, 1, InStr(citem17, ",")))
    citem17 = Trim(Mid(citem17, InStr(citem17, ",") + 1))
    
    item18 = Val(Mid(citem18, 1, InStr(citem18, ",")))
    citem18 = Trim(Mid(citem18, InStr(citem18, ",") + 1))
    
    mstrCode = strSource
    
label_0:
    lng_S = InStr(mstrCode, "R,")
    If lng_S > 1 Then
      mstrCode = Mid(mstrCode, lng_S)
    End If

    If Not IsComplate(mstrCode, str_Code) Then
      strReserved = mstrCode
      Exit Sub
    End If
    
    '截取剩余字符
    mstrCode = Mid(mstrCode, Len(str_Code) + 1)

    If str_Code Like "*R,*" Then
       var_Code = Split(str_Code, ",")
       str_Return = "": str_标本号 = "": str_日期 = "": strTmp = ""
       For lng_count = LBound(var_Code) To UBound(var_Code)
          str_Segment = var_Code(lng_count)
          Select Case lng_count
             Case sampleno
                str_标本号 = Trim(str_Segment)
             Case sampledate
                str_日期 = Trim(str_Segment) & " " & Format(Now, "HH:MM:SS")
             Case Item
                strTmp = Trim(str_Segment)
             Case item13
                str_项目 = strTmp & citem13
                str_结果 = Trim(str_Segment)
                If Val(str_结果) <> 0 Then
                  str_Return = str_Return & "|" & str_项目 & "|" & str_结果
                End If
             Case item14
                str_项目 = strTmp & citem14
                str_结果 = Trim(str_Segment)
                If (Trim(citem14) <> "") And (Val(str_结果) <> 0) Then
                  str_Return = str_Return & "|" & str_项目 & "|" & str_结果
                End If
             Case item15
                str_项目 = strTmp & citem15
                str_结果 = Trim(str_Segment)
                If (Trim(citem15) <> "") And (Val(str_结果) <> 0) Then
                  str_Return = str_Return & "|" & str_项目 & "|" & str_结果
                End If
             Case item16
                str_项目 = strTmp & citem16
                str_结果 = Trim(str_Segment)
                If (Trim(citem16) <> "") And (Val(str_结果) <> 0) Then
                  str_Return = str_Return & "|" & str_项目 & "|" & str_结果
                End If
             Case item17
                str_项目 = strTmp & citem17
                str_结果 = Trim(str_Segment)
                If (Trim(citem17) <> "") And (Val(str_结果) <> 0) Then
                  str_Return = str_Return & "|" & str_项目 & "|" & str_结果
                End If
             Case item18
                str_项目 = strTmp & citem18
                str_结果 = Trim(str_Segment)
                If (Trim(citem18) <> "") And (Val(str_结果) <> 0) Then
                  str_Return = str_Return & "|" & str_项目 & "|" & str_结果
                End If
          End Select
       Next
       If Val(str_标本号) > 0 And str_Return <> "" Then
          strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血清|0" & str_Return
       End If
    End If
    If mstrCode Like "*R,*" Then
      GoTo label_0
    End If
    If strResult <> "" Then
      If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    Exit Sub
errHandle:
    WriteLog "RAC00_Analyse", strSource, Err.Number & " " & Err.Description
End Sub

Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    SendSample = ""
    iSendStep = 0
End Function
    

Private Function IsComplate(ByVal as_data As String, Optional as_ret As String) As Boolean
  Dim i As Integer, li_len As Integer, li_cnt As Integer, li_pos As Integer
  IsComplate = False
  li_len = Len(as_data)
  For i = 1 To li_len
    If Mid(as_data, i, 1) = "," Then
        li_cnt = li_cnt + 1
        If li_cnt = 17 Then
            Exit For
        End If
    End If
  Next
  If li_cnt < 17 Then
    Exit Function
  End If

  li_pos = InStr(i, as_data, "R")
  If li_pos > 0 Then
     as_ret = Left(as_data, li_pos - 1)
     If Right(as_ret, 1) = "T" Then
        as_ret = Left(as_ret, Len(as_ret) - 1)
     End If
  Else
     If li_len > i + 4 Then
        as_ret = Left(as_data, i + 5)
        If Right(as_ret, 1) = "R" Then
            as_ret = Left(as_data, i + 4)
        End If
     Else
        as_ret = as_data
     End If
  End If
  IsComplate = True
End Function
