VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_geb600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:GEB600
'厂商:希森美康
'作者:
'用户:湛江二院

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
        
    Dim strCode As String
    Dim lngS As Long, lngE As Long, i As Integer
    Dim str标本号 As String, str项目 As String, str结果 As String, str日期 As String, str质控 As String
    Dim strItems As String
    Dim varline As Variant
    Dim iRow As Integer
    
    strReserved = strSource
    
    If strReserved Like "*" & Chr(2) & "*" & Chr(3) & "*" Then
        lngS = InStr(strReserved, Chr(2))
        lngE = InStr(strReserved, Chr(3))
        strCode = Mid(strReserved, lngS, lngE - lngS + 1)
        strReserved = Mid(strReserved, lngE + 1)
    
        varline = Split(strCode, vbCrLf)
        
        str日期 = Replace(Mid(varline(1), 6), "/", "-")
        If Not IsDate(str日期) Then
            str日期 = Format(Now, "yyyy-MM-dd hh:mm:ss")
        Else
            str日期 = Format(CDate(str日期), "yyyy-MM-dd hh:mm:ss")
        End If
        'str质控 = Mid(varline(2), 7, 1)
        str质控 = "0"
        str标本号 = Val(Mid(varline(1), 2, 7))
        strItems = ""
        For iRow = 3 To UBound(varline)
            str项目 = Mid(varline(iRow), 1, 5)
            str项目 = Replace(str项目, "*", "")
            str项目 = UCase(Replace(str项目, " ", ""))
            If str项目 = "LEU" Then
                strItems = strItems & "|" & str项目 & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str项目 = "NIT" Then
                strItems = strItems & "|" & str项目 & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str项目 = "URO" Then
                strItems = strItems & "|" & str项目 & "|" & Replace(Mid(varline(iRow), 6, 12), " ", "")
            End If
            If str项目 = "BIL" Then
                strItems = strItems & "|" & str项目 & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str项目 = "VC" Then
                strItems = strItems & "|" & str项目 & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str项目 = "PRO" Then
                strItems = strItems & "|" & str项目 & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str项目 = "BLD" Then
                strItems = strItems & "|" & str项目 & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str项目 = "PH" Then
                strItems = strItems & "|" & str项目 & "|" & Val(Mid(varline(iRow), 6))
            End If
            If str项目 = "SG" Then
                strItems = strItems & "|" & str项目 & "|" & Val(Mid(varline(iRow), 6))
            End If
            If str项目 = "GLU" Then
                strItems = strItems & "|" & str项目 & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
            If str项目 = "KET" Then
                strItems = strItems & "|" & str项目 & "|" & Replace(Mid(varline(iRow), 6, 4), " ", "")
            End If
        Next
        

        If Val(str标本号) > 0 And strItems <> "" Then
            strResult = str日期 & "|" & Val(str标本号) & "| |血液|" & str质控 & strItems
        End If
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




