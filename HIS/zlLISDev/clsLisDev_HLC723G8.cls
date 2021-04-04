VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_HLC723G8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:HLC-723G8
'厂商:
'作者:周海全
'用户:
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
        Dim lngS As Long, lngE As Long, strLine As String
        Dim varTmp As Variant, intloop As Integer, intCount As Integer
        Dim strTmp As String, strSampleNo As String, strDate As String, strItem As String
        On Error GoTo hErr
    
100     mstrCode = mstrCode & strSource
102     lngE = InStr(mstrCode, Chr(&HD))
104     strResult = ""
106     strReserved = ""
108     strDate = Format(Now, "yyyy-MM-dd HH:mm:ss")
110     Do While lngE > 0
112         strLine = Mid$(mstrCode, 1, lngE - 1)
114         mstrCode = Mid$(mstrCode, lngE + 1)
            
116         If Mid$(strLine, 1, 10) <> "9999999999" Then
118             If strLine Like "*####-##-## ##:##:##*Read*" Then
120                 strTmp = Mid$(strLine, InStr(strLine, "-") - 4, 19)
122                 strLine = Mid$(strLine, InStr(strLine, "Read") + 7)
124                 strDate = CDate(strTmp)
                End If
                
126             If strLine Like "########*##" Then
128                 strSampleNo = Mid(strLine, 8, 3)

130                 intCount = 0
132                 strItem = ""
134                 For intloop = 0 To 9
136                     strTmp = Trim(Mid$(strLine, 11 + intloop * 5, 5))
138                     If strTmp <> "" Then
140                         intCount = intCount + 1
142                         strItem = strItem & "|" & Format(intCount, "000") & "|" & strTmp
                        End If
                    Next
                    
144                 If IsDate(strDate) And Val(strSampleNo) > 0 And strItem <> "" Then
146                     strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Boold|0" & strItem
                    End If
                End If
            End If
148         lngE = InStr(mstrCode, Chr(&HD))
        Loop
150     If strResult <> "" Then strResult = Mid(strResult, 3)
        Exit Sub
hErr:
152     WriteLog "HLC723G8 Line:" & CStr(Erl()), mstrCode, strResult & vbNewLine & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String

    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function




