VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_G7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:HLC-723G7
'厂商:
'作者:陈东
'用户:晋煤集团总医院
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
        Dim varTmp As Variant, intLoop As Integer, intCount As Integer
        Dim strTmp As String, strSampleNo As String, strDate As String, strItem As String
        On Error GoTo hErr
    
100     mstrCode = mstrCode & strSource
102     lngE = InStr(mstrCode, Chr(&HD))
104     strResult = ""
106     strReserved = ""
    
108     Do While lngE > 0
110         strLine = Mid$(mstrCode, 1, lngE - 1)
112         mstrCode = Mid$(mstrCode, lngE + 1)
114         If Mid$(strLine, 1, 10) <> "9999999999" Then
116             strTmp = Mid$(strLine, 1, 10)
118             strLine = Mid$(strLine, 11)
120             strSampleNo = Right$(strTmp, 4)
122             strDate = Format(Now, "yyyy") & "-" & Mid$(strTmp, 3, 2) & "-" & Mid$(strTmp, 5, 2)
            
124             varTmp = Split(strLine, " ")
126             intCount = 0
128             strItem = ""
130             For intLoop = LBound(varTmp) To UBound(varTmp)
132                 strTmp = Trim$(varTmp(intLoop))
134                 If strTmp <> "" Then
136                     intCount = intCount + 1
138                     strItem = strItem & "|" & Format(intCount, "000") & "|" & strTmp
                    End If
                Next
            
140             If IsDate(strDate) And Val(strSampleNo) > 0 And strItem <> "" Then
142                 strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |血清|0" & strItem
                End If
            End If
144         lngE = InStr(mstrCode, Chr(&HD))
        Loop
146     If strResult <> "" Then strResult = Mid(strResult, 3)
        Exit Sub
hErr:
148     WriteLog "G7 Line:" & CStr(Erl()), mstrCode, strResult & vbNewLine & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String

    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function


