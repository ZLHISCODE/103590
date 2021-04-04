VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Mp280"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:MP-280化学发光免疫分析仪
'厂商:北京泰格科信生物科技有限公司
'作者:陈东
'用户:万荣县医院（太原）

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

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    
        Dim aResults() As String
        Dim strFilePath As String, objFile As New Scripting.FileSystemObject
        Dim objStream As Scripting.TextStream, strLine As String, varTmp As Variant, strTmp As String
        Dim iLoop As Integer
        Dim strNo As String, strDate As String, strItem As String, strValue As String
    
        On Error GoTo hErr
    
100     If Dir(strFile) = "" Then Exit Function
102     Set objStream = objFile.OpenTextFile(strFile, ForReading)

104     ReDim aResults(0)
106     Do Until objStream.AtEndOfStream
108         strLine = objStream.ReadLine
110         If strLine Like "####/##/##*" Then
112             varTmp = Split(strLine, vbTab)
114             If UBound(varTmp) >= 6 Then
116                 strDate = Replace(varTmp(0), "/", "-")
118                 strNo = Trim(varTmp(1))
120                 strItem = Trim(varTmp(2))
122                 strValue = Trim(varTmp(4))
                
124                 If IsDate(strDate) And Val(strNo) > 0 And strItem <> "" And strValue <> "" Then
126                     If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
128                     aResults(UBound(aResults)) = strDate & "|" & strNo & "| |血清|0|" & strItem & "|" & strValue
                    End If
                End If
            End If
        Loop
130     ResultFromFile = aResults
        Exit Function
hErr:
132     WriteLog "MP280", CStr(Erl()) & "," & Err.Description, ""
End Function



