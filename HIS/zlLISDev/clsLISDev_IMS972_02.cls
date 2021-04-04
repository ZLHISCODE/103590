VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_IMS972_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:IMS 972 电解质分析仪
'厂商:
'作者:王振涛
'用户:太原
'日期：2012-12-17

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
        
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号^是否急诊^条码号  可以只有样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令
   
    Dim lng_S As Long, lng_E As Long, str_in As String
    Dim str_Code As String '一个完整记录
    Dim str_Line As String '一行数据
    Dim str_标本号 As String, str_项目 As String, str_结果 As String, str_日期 As String
    Dim str_条码 As String, str_SendLine As String
    
    Dim lng_Loop As Long, var_Code As Variant
    strResult = ""
    str_in = strSource
    
    If Len(str_in) > 0 Then mstrCode = mstrCode & str_in
    
    If mstrCode Like "*PAT_ID*" & Chr(&HD) & "*" Then                          '收到一个完整的记录,开始解析
        WriteLog "解析 IMS 972", mstrCode, ""
        Do While InStr(mstrCode, "PAT_ID") > 0
'            lng_S = InStr(mstrCode, "nCa")
             lng_S = 1

            lng_E = InStr(mstrCode, "PAT_ID")
            lng_E = InStr(lng_E, mstrCode, Chr(&HD))
            
            str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
            mstrCode = Mid(mstrCode, lng_E + 1)
            var_Code = Split(str_Code, Chr(&HD))
            str_项目 = ""
            For lng_Loop = LBound(var_Code) To UBound(var_Code)

                str_Line = var_Code(lng_Loop)
                
                If InStr(str_Line, "PAT_ID") > 0 Then                       '项目 结果 行
                    str_标本号 = Val(Mid(str_Line, 7))
                ElseIf InStr(",pH,TCa,iCa,nCa,Cl,K,Na,", "," & Trim(Mid(str_Line, 1, 3)) & ",") > 0 Then
                    str_结果 = Val(Mid(str_Line, 4))
                    If Val(str_结果) <> 0 Then
                        str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 1, 3)) & "|" & str_结果
                    End If
                End If

            Next
            If Val(str_标本号) <> 0 And str_项目 <> "" Then
                strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|" & Val(str_标本号) & "| |血清|0" & str_项目
            End If
        Loop
    End If
    
    If strResult <> "" Then
        strResult = Mid(strResult, 3)
        WriteLog "解析 IMS 972", strResult, ""
    End If
End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    SendSample = ""

End Function






