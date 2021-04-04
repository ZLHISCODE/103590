VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_AU680"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
'山西长治中研所
'au680 生化仪  单向
'2011-11-21 陈东
'修改：王振涛
'湖南龙山县人民医院，晋城合聚心脑血管病医院,天水第二人民医院
'2012-12-26

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
    Dim lngItemS As Long, lngItemLen As Long, lngItemNoLen As Long
    Dim strSampeNo As String, strDate As String
    Dim str_仪器类型 As String
    Dim intE As Integer, intQ As Integer, intP As Integer
    lngSampleS = 11     '标本号开始位置
    lngSampleLen = 4    '标本号长度
    lngItemS = 20       '项目开始位置
    lngItemLen = 11     '一个项目结果的长度
    lngItemNoLen = 3    '通道码长度
    
    
        
    intE = Val(ReadIni("AU680", "E", App.Path & "\AU680.ini"))
    If intE < 0 Then intE = 8000:   Call WriteIni("AU680", "E", 8000, App.Path & "\AU680.ini")
    
    intQ = Val(ReadIni("AU680", "Q", App.Path & "\AU680.ini"))
    If intQ < 0 Then intQ = 6000:   Call WriteIni("AU680", "Q", 6000, App.Path & "\AU680.ini")
    
    intP = Val(ReadIni("AU680", "P", App.Path & "\AU680.ini"))
    If intP < 0 Then intP = 3000:   Call WriteIni("AU680", "P", 3000, App.Path & "\AU680.ini")
    str_仪器类型 = Trim(ReadIni("仪器类型", "类型", App.Path & "\AU680.ini"))
    If str_仪器类型 = "" Then
        WriteIni "仪器类型", "类型", "AU680", App.Path & "\AU680.ini"
        str_仪器类型 = "AU680"
    End If
    
    If str_仪器类型 = "AU680_LSXRMYY" Then
        lngSampleS = 13
        lngItemS = 48
        lngItemLen = 14
    ElseIf str_仪器类型 = "AU680_TSRMYY" Then
        lngItemS = 46
        lngItemLen = 14
    End If

    If InStr(strSource, Chr(&H3)) > 0 Then strCmd = Chr(&H6)
    mstrCode = mstrCode & strSource
    
    strCode = PStr_CutCode(mstrCode, Chr(&H2), Chr(&H3))
    Do While strCode <> ""
'        If strCode Like Chr(&H2) & "DE*" Then
            
        strDate = Format(Now, "yyyy-MM-dd HH:mm:ss")
        strSampeNo = Mid(strCode, lngSampleS, lngSampleLen)
        If Not IsNumeric(Mid(strSampeNo, 1, 1)) Then
            Select Case Mid(strSampeNo, 1, 1)
                Case "P"
                    strSampeNo = intP + Mid(strSampeNo, 2)
                Case "E"
                    strSampeNo = intE + Mid(strSampeNo, 2)
            End Select
        End If
        strLine = Mid$(strCode, lngItemS)
        strItem = ""
        Do While Len(strLine) > lngItemLen
            
            strTmp = Mid(strLine, 1, lngItemLen)
            strLine = Mid(strLine, lngItemLen + 1)
            If Trim(Mid(strTmp, 1, lngItemNoLen)) <> "" And Trim(Mid(strTmp, lngItemNoLen + 1)) <> "" Then
                strItem = strItem & "|" & Mid(strTmp, 1, lngItemNoLen) & "|" & Val(Mid$(strTmp, lngItemNoLen + 1))
            End If
            
        Loop
        If Val(strSampeNo) > 0 And strItem <> "" Then
            strResult = strResult & "||" & strDate & "|" & strSampeNo & "| |血清|0" & strItem
        End If
'        End If
        strCode = PStr_CutCode(mstrCode, Chr(&H2), Chr(&H3))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String

End Function













