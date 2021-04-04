VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_BN_ProSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'山西长治中研所
'BN ProSpec 蛋白分析仪  单向
'2011-11-21 陈东
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
    Dim lng_S As Long, lng_E As Long
    Dim strCode As String, strLine As String, strTmp As String, strItem As String
    Dim lngSampleS As Long, lngSampleLen As Long
    Dim lngItemS As Long, lngItemLen As Long, strFlg As String, strValue As String, strChannels As String
    Dim strSampeNo As String, strDate As String
    
    lngSampleS = 3     '标本号开始位置
    lngSampleLen = 3    '标本号长度
    lngItemS = 32       '项目开始位置
    lngItemLen = 55     '一个项目结果的长度
    
    
    If InStr(strSource, Chr(&H3)) > 0 Then strCmd = Chr(&H6)
    mstrCode = mstrCode & strSource
    
    strCode = PStr_CutCode(mstrCode, Chr(&H2), Chr(&H3))
    Do While strCode <> ""
        If strCode Like Chr(&H2) & "D*" Then
            
            strDate = Format(Now, "yyyy-MM-dd HH:mm:ss")
            strSampeNo = Mid(strCode, lngSampleS, lngSampleLen)
            strLine = Mid$(strCode, lngItemS)
            strItem = ""
            Do While Len(strLine) > lngItemLen
                
                strTmp = Mid(strLine, 1, lngItemLen)
                strLine = Mid(strLine, lngItemLen + 1)
                
                strChannels = Mid$(strTmp, 1, 2)
                
                strFlg = Mid$(strTmp, 6, 1)
                If strFlg = "0" Then
                    strFlg = ""
                ElseIf strFlg = "2" Then
                    strFlg = "<"
                Else
                    strFlg = ""
                End If
                strValue = Trim(Mid(strTmp, 12, 10))
                If Trim(strChannels) <> "" And Trim(strValue) <> "" Then
                    strItem = strItem & "|" & strChannels & "|" & strFlg & Val(strValue)
                End If
                
            Loop
            If Val(strSampeNo) > 0 And strItem <> "" Then
                strResult = strResult & "||" & strDate & "|" & strSampeNo & "| |血清|0" & strItem
            End If
        End If
        strCode = PStr_CutCode(mstrCode, Chr(&H2), Chr(&H3))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String

End Function




