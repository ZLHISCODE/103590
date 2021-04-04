VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Ca530"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:CA530
'厂商:麦道尼克(MEDONIC)
'作者:陈东
'用户:成都XX医院
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
        
    Dim str_Code As String, var_Code As Variant, str_Line As String
    Dim lng_S As Long, lng_E As Long, i As Integer, str_Tmp As String
    Dim str_Item As String, l As Integer
    Dim str_RBC As String, str_PLT As String, str_WBC As String, str_WBCM As String
    Dim str_项目 As String, str_日期 As String, str_标本号 As String, str_img As String
    
    If Len(strSource) > 0 Then mstrCode = mstrCode & strSource
    
    str_Item = "RBC,MCV,HCT,PLT,MPV,WBC,HGB,MCH,MCHC,TRBC,TWBC,LYMF,GRAN,MID,LPR,GPR,MPR,RDWR,RDWA,PDW,PCT,LPCR"
    
    lng_S = InStr(mstrCode, "+++++")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, "#####")
    Else
        lng_E = InStr(mstrCode, "#####")
    End If
    str_项目 = "": str_WBC = ""
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E + 4 - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 5)
        '--------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        For i = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(i)
            If Trim(Split(str_Line, "=")(0)) = "DATE" Then
                str_Tmp = Trim(Split(str_Line, "=")(1))
                str_日期 = Split(Split(str_Tmp, "/")(2), "-")(0) & "-" & Format(Split(str_Tmp, "/")(1), "00") & "-" & Format(Split(str_Tmp, "/")(0), "00") & " " & Split(Split(str_Tmp, "/")(2), "-")(1)
            ElseIf Trim(Split(str_Line, "=")(0)) = "SEQ" Then
                str_标本号 = Trim(Split(str_Line, "=")(1))
            ElseIf InStr(str_Item, Trim(Split(str_Line, "=")(0))) > 0 And Trim(Mid(str_Line, 6, 8)) <> "" Then
                str_项目 = str_项目 & "|" & Trim(Split(str_Line, "=")(0)) & "|" & Trim(Mid(str_Line, 6, 8))
            ElseIf Trim(Split(str_Line, "=")(0)) = "CRBC" Then
                str_RBC = ""
                For l = 1 To 10
                    str_RBC = str_RBC & Trim(var_Code(i + l))
                Next
            ElseIf Trim(Split(str_Line, "=")(0)) = "CPLT" Then
                str_PLT = ""
                For l = 1 To 10
                    str_PLT = str_PLT & Trim(var_Code(i + l))
                Next
            ElseIf InStr("CWBCL,CWBCM,CWBCG", Trim(Split(str_Line, "=")(0))) > 0 Then
                
                For l = 1 To 10
                    If Trim(Split(str_Line, "=")(0)) = "CWBCM" Then
                        str_WBCM = Trim(var_Code(i + l))
                    Else
                        str_WBC = str_WBC & Trim(var_Code(i + l))
                    End If
                Next
                
            End If
        Next
        If IsDate(str_日期) And str_项目 <> "" And Val(str_标本号) > 0 Then
            strResult = strResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|0" & str_项目
            str_img = ""
            If str_RBC <> "" Then str_img = str_img & "^RBC;0;300,100,8-20,40-100,80-200" & GetImgData(str_RBC)
            If str_PLT <> "" Then str_img = str_img & "^PLT;0;120,100,8-2,40-10,80-20" & GetImgData(str_PLT)
            If str_WBC <> "" Then str_img = str_img & "^WBC;0;120,100,5.33-20,26.66-100,53.33-200,80-300" & Replace(GetImgData(str_WBC & str_WBCM), ";000", "")
            If str_img <> "" Then strResult = strResult & vbCrLf & Mid(str_img, 2)
            
        End If
        '--------------------------------------
        lng_S = InStr(mstrCode, "+++++")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, "#####")
        Else
            lng_E = InStr(mstrCode, "#####")
        End If
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
End Sub

Private Function GetImgData(ByVal strIn As String) As String
    Dim str_Data As String
    Dim str_Bit As String
    Dim str_Retu As String
    str_Data = strIn
    If str_Data = "" Then Exit Function
    Do While str_Data <> ""
        str_Bit = Left(str_Data, 3)
        str_Data = Mid(str_Data, 4)
        str_Retu = str_Retu & ";" & Format(CByte("&H" & Mid(str_Bit, 2)), "000")
    Loop
    If str_Retu <> "" Then GetImgData = str_Retu
End Function

Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function


