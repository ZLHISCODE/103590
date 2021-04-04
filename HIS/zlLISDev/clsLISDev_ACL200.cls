VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ACL200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: ACL200 血凝仪
'厂商:
'作者: 陈东
'用户: 山西
Private mstr_Code As String

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
    Dim strCode As String, varTmp As Variant, strLine As String
    Dim strDate As String, strSampleNo As String, strItem As String, strValue As String
    On Error GoTo hErr
    If strSource <> "" Then mstr_Code = mstr_Code & strSource
    strResult = ""
    strCode = PStr_CutCode(mstr_Code, "1B05C", "21" & Chr(&HD) & Chr(&HA) & Chr(&HD) & Chr(&HA))
    Do While strCode <> ""
        varTmp = Split(strCode, Chr(&HD) & Chr(&HA) & Chr(&HD) & Chr(&HA))
        
        If UBound(varTmp) >= 5 Then
            '日期
            strLine = varTmp(0)
            strDate = "20" & Mid$(strLine, 16, 2) & "-" & TrunMonth(Mid$(strLine, 11, 3)) & "-" & Mid$(strLine, 7, 2) & " " & Mid$(strLine, 21, 2) & ":" & Mid$(strLine, 25, 2)
            
            strLine = varTmp(5)
            strSampleNo = Mid$(strLine, 10, 3)
            strItem = ""
            
            strLine = varTmp(3)
            
            strValue = Trim(Mid$(strLine, 1, 4))   'pT
            If strValue <> "" Then strItem = strItem & "|PT|" & strValue
            strValue = Mid$(strLine, 6, 3)   'hd
            If strValue <> "" Then strItem = strItem & "|HD|" & strValue
            strValue = Mid$(strLine, 10, 4)  'ptinr
            If strValue <> "" Then strItem = strItem & "|PTINR|" & strValue
            strValue = Mid$(strLine, 15, 4)  'FIB
            If strValue <> "" Then strItem = strItem & "|FIB|" & strValue
            strLine = varTmp(4)
            strValue = Mid$(strLine, 1, 4) 'APTT
            If strValue <> "" Then strItem = strItem & "|APTT|" & strValue
            strValue = Mid$(strLine, 5, 3) 'APTTR
            If strValue <> "" Then strItem = strItem & "|APTTR|" & strValue
            If strDate <> "" And strItem <> "" And strSampleNo <> "" Then
                strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Boold|0" & strItem
            End If
        End If

        strCode = PStr_CutCode(mstr_Code, "1B05C", "21" & Chr(&HD) & Chr(&HA) & Chr(&HD) & Chr(&HA))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
hErr:
    WriteLog "ACL200", strSource, Err.Description
End Sub



Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5个元素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    
End Function






Private Function TrunMonth(ByVal strMonth As String) As String
    '转换英文的月份为数字
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select

End Function



