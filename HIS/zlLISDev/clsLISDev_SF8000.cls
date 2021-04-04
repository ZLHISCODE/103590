VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_SF8000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:SF-8000全自动凝血测试仪
'厂商:赛科希德
'作者:陈东
'用户:贵阳市肺科医院
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
    Dim str_标本号 As String, str_项目 As String, str_结果 As String
    mstrCode = mstrCode & strSource
    lng_E = InStr(mstrCode, Chr(&HA))
    
    Do While lng_E > 0
        str_Code = Mid(mstrCode, 1, lng_E)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '----------------------------------------
        If Len(str_Code) > 12 Then
            str_标本号 = Trim(Mid(str_Code, 1, 4))
            str_项目 = Trim(Mid(str_Code, 6, 6))
            str_结果 = Trim(Mid(str_Code, 12))
            str_项目 = Replace(str_项目, "g/L", "")
            
            If Val(str_标本号) > 0 And str_项目 <> "" And str_结果 <> "" Then
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd HH:mm") & "|" & str_标本号 & "| |血清|0|" & str_项目 & "|" & str_结果
            End If
        End If
        
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
    
End Function


