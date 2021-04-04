VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_HC9885"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:HC-9885 电解质分析仪
'厂商:深圳航创
'作者:陈东
'用户:贵阳肺科医院

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
    Dim lng_S As Long, lng_E As Long, str_Code As String
    Dim str_标本号 As String, str_项目 As String, str_Tmp As String
    Dim aItem() As String
    
    mstrCode = mstrCode & Replace(strSource, "!", "")
    mstrCode = Replace(mstrCode, Chr(&H0), "")
    
    lng_E = InStr(mstrCode, Chr(&HA))
    strResult = ""
    Do While lng_E > 0
        
        str_Code = Trim(Mid(mstrCode, 1, lng_E))
        mstrCode = Mid(mstrCode, lng_E + 1)
        WriteLog "开始解析", str_Code, ""
        '----------------------------------------
        str_标本号 = "": str_项目 = "": str_Tmp = ""
        str_Code = DelSapce(str_Code)
        aItem = Split(str_Code, " ")
        If UBound(aItem) >= 6 Then
            If Trim(aItem(0)) <> "" Then str_标本号 = aItem(0)
            If Trim(aItem(1)) <> "" Then str_项目 = str_项目 & "|01|" & aItem(1)
            If Trim(aItem(2)) <> "" Then str_项目 = str_项目 & "|02|" & aItem(2)
            If Trim(aItem(3)) <> "" Then str_项目 = str_项目 & "|03|" & aItem(3)
            If Trim(aItem(4)) <> "" Then str_项目 = str_项目 & "|04|" & aItem(4)
            If Trim(aItem(5)) <> "" Then str_项目 = str_项目 & "|05|" & aItem(5)
            If Trim(aItem(6)) <> "" Then str_项目 = str_项目 & "|06|" & aItem(6)
            
            If UBound(aItem) >= 7 Then _
                If Trim(aItem(7)) <> "" Then str_项目 = str_项目 & "|07|" & aItem(7)
            If UBound(aItem) >= 8 Then _
                If Trim(aItem(8)) <> "" Then str_项目 = str_项目 & "|08|" & aItem(8)
        End If
        
        
'        str_标本号 = Mid(str_Code, 1, 3)
'        str_Tmp = Trim(Mid(str_Code, 7, 4))
'        If str_Tmp <> "" Then str_项目 = str_项目 & "|01|" & str_Tmp
'        str_Tmp = Trim(Mid(str_Code, 14, 5))
'        If str_Tmp <> "" Then str_项目 = str_项目 & "|02|" & str_Tmp
'        str_Tmp = Trim(Mid(str_Code, 22, 4))
'        If str_Tmp <> "" Then str_项目 = str_项目 & "|03|" & str_Tmp
'        str_Tmp = Trim(Mid(str_Code, 29, 4))
'        If str_Tmp <> "" Then str_项目 = str_项目 & "|04|" & str_Tmp
'        str_Tmp = Trim(Mid(str_Code, 36, 4))
'        If str_Tmp <> "" Then str_项目 = str_项目 & "|05|" & str_Tmp
'        str_Tmp = Trim(Mid(str_Code, 43, 4))
'        If str_Tmp <> "" Then str_项目 = str_项目 & "|06|" & str_Tmp
        
        If Val(str_标本号) > 0 And str_项目 <> "" Then
            strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm") & "|" & str_标本号 & "| |血清|0" & str_项目
            WriteLog "完成解析", "", strResult
        End If
        str_项目 = ""
        str_标本号 = ""
        '----------------------------------------
        lng_E = InStr(mstrCode, Chr(&HA))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
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
