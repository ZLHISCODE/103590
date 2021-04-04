VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ADVIACentaurCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器: ADVIA Centaur CP 全自动化学发光仪
'厂商: 德国西门子
'作者: 陈东
'用户: 万荣县医院

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
    
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant, str_Line As String
    Dim intLoop As Integer, str_SampleNo As String, str_Date As String, str_Item As String, str_Value As String
    Dim str_tmp As String
    strCmd = ""
    If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
    If InStr(strSource, Chr(&HA)) > 0 Then strCmd = strCmd & Chr(&H6)

    mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H5))
    lng_E = 0
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    End If
    strResult = ""
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        str_Item = "": str_SampleNo = "": str_Date = ""
        
        var_Code = Split(str_Code, Chr(&HA))
        For intLoop = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(intLoop)
            If str_Line Like "?O|*" Then
                str_SampleNo = Trim(Split(str_Line, "|")(2))
            ElseIf str_Line Like "?R|1|*" Then
                If str_Date = "" Then
                    str_Date = Trim(Split(str_Line, "|")(12))
                    str_Date = Mid(str_Date, 1, 4) & "-" & Mid(str_Date, 5, 2) & "-" & Mid(str_Date, 7, 2) & " " & Mid(str_Date, 9, 2) & ":" & Mid(str_Date, 11, 2) & ":" & Mid(str_Date, 13, 2)
                End If
                str_tmp = Trim(Split(str_Line, "|")(2))
                str_tmp = Trim(Split(str_tmp, "^")(3))
                str_Value = Trim(Split(str_Line, "|")(3))
                If str_tmp <> "" And str_Value <> "" Then str_Item = str_Item & "|" & str_tmp & "|" & str_Value
            End If
        Next
        
        If str_Date <> "" And Val(str_SampleNo) > 0 And str_Item <> "" Then _
            strResult = strResult & "||" & str_Date & "|" & str_SampleNo & "| |血清|0" & str_Item
        lng_S = InStr(mstrCode, Chr(&O5))
        lng_E = 0
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub


Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False) As String
    '解析需要向仪器发送的标本信息
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    blnSuccess = True
    iSendStep = 0
    SendSample = ""
End Function

