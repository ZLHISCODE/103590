VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_KX21_bin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  sysmex KX21
'厂商:  希森美康 sysmex
'医院:  成都德阳
'作者:  陈东
'时间:  2008/3/31
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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    '设备标准接口
    '功能：解析数据
    Dim str_Code As String, lng_S As Long, lng_len As Long
    Dim str_Day As String, str_日期 As String, str标本号 As String, strValues As String
    Dim int_i As Integer, strRecord As String
    Dim str_WBC As String, str_RBC As String, str_PLT As String
    Dim str_Data As String
    strReserved = strSource
    lng_len = Len_Bin(strReserved)
    If lng_len < 457 Then Exit Sub
    
    For lng_S = 1 To lng_len - 4
        str_Code = Mid_bin(strReserved, lng_S, 1)
        str_Day = Mid_bin(strReserved, lng_S + 3, 1)
        
        If str_Code = Chr(&H2) And str_Day = "/" Then
            If lng_len - lng_S < 456 Then Exit Sub
            str_Data = Mid_bin(strReserved, lng_S, 457, False)
            strReserved = Mid_bin(strReserved, lng_S + 457, 0, False)
            If strReserved = ",00" Then strReserved = ""
            str_日期 = Format(CDate(Trim(Mid_bin(str_Data, 2, 2)) & "-" & Trim(Mid_bin(str_Data, 5, 2)) & "-" & Trim(Mid_bin(str_Data, 8, 2))), "yyyy-MM-dd")
            Exit For
        Else
            If str_Code = "" Then
                str_Day = Mid_bin(strReserved, lng_S + 1, 8)
                If IsDate(Left(str_Day, 4) & "-" & Mid(str_Day, 5, 2) & "-" & Right(str_Day, 2)) Then
                    If lng_len - lng_S < 456 Then Exit Sub
                    str_Data = Mid_bin(strReserved, lng_S, 457, False)
                    strReserved = Mid_bin(strReserved, lng_S + 457, 0, False)
                End If
                str_日期 = Format(CDate(Left(str_Day, 4) & "-" & Mid(str_Day, 5, 2) & "-" & Right(str_Day, 2)), "yyyy-MM-dd")
                Exit For
            End If
        End If
        
    Next
    
    
    If lng_S = lng_len - 3 Then Exit Sub
    If str_日期 = "" Then Exit Sub
    str标本号 = Mid_bin(str_Data, 15, 13)
    strRecord = ""
    For int_i = 1 To 21
        strValues = Mid_bin(str_Data, 78 + int_i * 5, 5)
        If ((((int_i = 9) Or (int_i = 10)) Or (int_i = 11)) Or (int_i = 15)) Then
            If InStr(strValues, "--") > 0 Then
                strValues = strValues
            Else
                strValues = Format(Val(strValues) * 100, "0.0")
            End If
        End If
        strRecord = strRecord & "|" & int_i & "|" & IIf(InStr(strValues, "--") > 0, strValues, Val(strValues))
    Next
    
    If Val(str标本号) > 0 And str_日期 <> "" And strRecord <> "" Then
        strResult = str_日期 & "|" & Val(str标本号) & "| |血清|0" & strRecord & vbCrLf
        
    
        str_WBC = "WBC;0;120,50,8.33-,16.66-100,24.99-,33.32-200,41.65-,49.98-300" & Graph_Text(Mid_bin(str_Data, 243, 50, False))
        str_RBC = "^RBC;0;120,50,10-,20-100,30-,40-200,50-" & Graph_Text(Mid_bin(str_Data, 293, 50, False))
        str_PLT = "^PLT;0;120,50,12.5-10,25-20,37.5-30,50-" & Graph_Text(Mid_bin(str_Data, 343, 40, False))
        strResult = strResult & str_WBC & str_RBC & str_PLT
        
        'strCmd = Chr(&H6)
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



Private Function Graph_Text(ByVal str_bin As String) As String
    '生成图形文本数据
    Dim lng_len As Long, lng_i As Long, str_bit As String
    Dim str_Return As String
    lng_len = Len_Bin(str_bin)

    If lng_len < 1 Then
        Graph_Text = ""
        Exit Function
    End If

    For lng_i = 1 To lng_len
        str_bit = Replace(Mid_bin(str_bin, lng_i, 1, False), ",", "")
    
        If IsNull(str_bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & ";" & Format(CByte("&H" & str_bit), "000")
        End If
    Next
    Graph_Text = str_Return
End Function
