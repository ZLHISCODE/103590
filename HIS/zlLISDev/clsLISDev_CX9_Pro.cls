VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CX9_Pro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:CX9 Pro
'厂商:美国贝克曼库尔特(Beckman coulter)
'用户:成都儿童专科医院

Private mstr_日期 As String
Private mstr_标本号 As String
Private mstr_Items As String
Private curTP As Currency '07C
Private curALB As Currency '08A

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    
    Dim lng_S As Long, lng_E As Long
    Dim str_Line As String, varTmp As Variant
    Dim str_标本号 As String, str_日期 As String, str_项目 As String, str_结果 As String
    
    strReserved = strSource
    If strReserved Like "*[[]*[]]*" Then
        
        lng_S = InStr(strReserved, "[")
        lng_E = InStr(strReserved, "]")
        str_Line = Mid(strReserved, lng_S, lng_E - lng_S + 1)
        strReserved = Mid(strReserved, lng_E + 1)
        If InStr(strReserved, "[") <= 0 Then strReserved = ""
        
        varTmp = Split(str_Line, ",")
        If UBound(varTmp) > 3 Then
            If varTmp(1) = "702" And varTmp(2) = "01" Then
                mstr_日期 = varTmp(3)
                mstr_日期 = Format(CDate(Mid(mstr_日期, 5, 2) & "-" & Mid(mstr_日期, 3, 2) & "-" & Mid(mstr_日期, 1, 2)), "yyyy-MM-dd")
                mstr_标本号 = Val(varTmp(12))
                
                mstr_Items = ""
                
            ElseIf varTmp(1) = "702" And varTmp(2) = "03" Then
                str_日期 = varTmp(3)
                str_日期 = Format(CDate(Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 3, 2) & "-" & Mid(str_日期, 1, 2)), "yyyy-MM-dd")
                
                str_标本号 = Val(varTmp(9))
                str_项目 = Trim(varTmp(10))
                str_结果 = Val(varTmp(15))
                
                If Val(str_标本号) = Val(mstr_标本号) And str_日期 = mstr_日期 Then
                    mstr_Items = mstr_Items & "|" & str_项目 & "|" & str_结果
                    
                    If str_项目 = "07C" Then curTP = Val(str_结果)
                    If str_项目 = "08A" Then curALB = Val(str_结果)
                    
                    
                End If
            ElseIf varTmp(1) = "702" And varTmp(2) = "05" Then
                str_标本号 = Val(varTmp(6))
                If mstr_日期 <> "" And Val(mstr_标本号) = Val(str_标本号) And mstr_Items <> "" Then
                    strResult = mstr_日期 & "|" & Val(mstr_标本号) & "| |血|" & IIf(Val(mstr_标本号) >= 1000, "1", "0") & mstr_Items
                    
                    If curTP <> 0 And curALB <> 0 Then
                        strResult = strResult & "|GLB|" & curTP - curALB & "|A/G|" & Format(curALB / (curTP - curALB), "0.000")
                    End If
                End If
                mstr_日期 = "": mstr_标本号 = "": mstr_Items = ""
            End If
        End If
        
    End If
End Sub

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
