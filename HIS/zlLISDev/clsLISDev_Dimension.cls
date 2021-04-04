VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Dimension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
    Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant
    Dim str_标本号 As String, str_日期 As String, str_项目 As String, str_结果 As String, i As Integer
    
    mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    
    If InStr(strSource, Chr(&H3)) > 0 Then strCmd = Chr(&H6)
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '---------------------------------------------------
        var_Code = Split(str_Code, Chr(&H1C))
        str_标本号 = Val(var_Code(3))
        str_日期 = var_Code(7)
        If Len(str_日期) >= 12 And IsNumeric(str_日期) Then
            str_日期 = "20" & Mid(str_日期, 11, 2) & "-" & Mid(str_日期, 9, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 5, 2) & ":" & Mid(str_日期, 3, 2)
        Else
            str_日期 = Format(Now, "yyyy-MM-dd HH:mm")
        End If
        str_项目 = ""
        
        If UBound(var_Code) > 11 Then
            For i = 11 To UBound(var_Code)
                If InStr(",TBIL,TP,XADA,XCK,TGL,CHOL,ALP,AST,URCA,XALT,DBIL,GLU,ALB,BUN,LDH,GGT,CREA,", "," & Trim(var_Code(i)) & ",") > 0 And Trim(var_Code(i)) <> "" Then
                    str_结果 = Trim(var_Code(i + 1))
                    If str_结果 <> "" Then
                        str_项目 = str_项目 & "|" & Trim(var_Code(i)) & "|" & str_结果
                    End If
                End If
            Next
        End If
        If Val(str_标本号) > 0 And IsDate(str_日期) And str_项目 <> "" Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|0" & str_项目
        End If
        
        '---------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
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

