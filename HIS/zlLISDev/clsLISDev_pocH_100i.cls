VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_pocH_100i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:  pocH-100i 三分类血球仪 日本东亚sysmex
'医院:  湖南韶阳正骨医院
'作者:  陈东
'时间:  2008/09/05

Private mstrCode As String
Private mLastReturn As String

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
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    '   微生物结果格式 通道码|结果类型^方式^结果值
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令

    Dim str_Code As String
    Dim str_Line As String               '当前要解析的字串
    Dim lng_S As Long                  '开始位置
    Dim lng_E As String                '结束位置
    
    Dim str_日期 As String, str_标本号 As String, str_项目 As String
    Dim StrWbc As String, StrRbc As String, StrPLt As String
    If InStr(strSource, Chr(&H3)) > 0 Then strCmd = Chr(&H6)
    
    mstrCode = mstrCode & strSource
    On Error GoTo errH
    strResult = ""
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = 0
    End If

    
        
    Do While lng_S > 0 And lng_E > 0

        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        str_日期 = "": str_标本号 = "": str_项目 = ""
        If UCase(Mid(str_Code, 2, 3)) = "D1U" Then
            str_日期 = Trim(Mid(str_Code, 45, 4) & "-" & Mid(str_Code, 49, 2) & "-" & Mid(str_Code, 51, 2))
            
            str_标本号 = Val(Mid(str_Code, 54, 15))
            
            str_项目 = "|WBC|" & GetValues(Mid(str_Code, 76, 4), 1)
            str_项目 = str_项目 & "|RBC|" & GetValues(Mid(str_Code, 81, 4), 2)
            str_项目 = str_项目 & "|HGB|" & GetValues(Mid(str_Code, 86, 4), 0)
            str_项目 = str_项目 & "|HCT|" & GetValues(Mid(str_Code, 91, 4), 1)
            str_项目 = str_项目 & "|MCV|" & GetValues(Mid(str_Code, 96, 4), 1)
            str_项目 = str_项目 & "|MCH|" & GetValues(Mid(str_Code, 101, 4), 1)
            str_项目 = str_项目 & "|MCHC|" & GetValues(Mid(str_Code, 106, 4), 0)
            str_项目 = str_项目 & "|PLT|" & GetValues(Mid(str_Code, 111, 4), 0)
            str_项目 = str_项目 & "|LYM%|" & GetValues(Mid(str_Code, 116, 4), 1)
            str_项目 = str_项目 & "|MXD%|" & GetValues(Mid(str_Code, 121, 4), 1)
            str_项目 = str_项目 & "|NEUT%|" & GetValues(Mid(str_Code, 126, 4), 1)
            str_项目 = str_项目 & "|LYM#|" & GetValues(Mid(str_Code, 131, 4), 1)
            str_项目 = str_项目 & "|MXD#|" & GetValues(Mid(str_Code, 136, 4), 1)
            str_项目 = str_项目 & "|NEUT#|" & GetValues(Mid(str_Code, 141, 4), 1)
            str_项目 = str_项目 & "|RDW-SD|" & GetValues(Mid(str_Code, 146, 4), 1)
            str_项目 = str_项目 & "|RDW-CV|" & GetValues(Mid(str_Code, 151, 4), 1)
            str_项目 = str_项目 & "|PDW|" & GetValues(Mid(str_Code, 156, 4), 1)
            str_项目 = str_项目 & "|MPV|" & GetValues(Mid(str_Code, 161, 4), 1)
            str_项目 = str_项目 & "|P-LCR|" & GetValues(Mid(str_Code, 166, 4), 1)
            
            If Val(str_标本号) <> 0 And IsDate(str_日期) And str_项目 <> "" Then
                mLastReturn = str_日期 & "|" & Val(str_标本号) & "| |血液|0"
                strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |血液|0" & str_项目
            End If
            
        ElseIf UCase(Mid(str_Code, 2, 2)) = "D2" Then
            If mLastReturn <> "" Then
                str_Code = Mid(str_Code, 4)
                StrWbc = Left(str_Code, 100)
                StrRbc = Mid(str_Code, 101, 100)
                StrWbc = HexToData(StrWbc)
                StrRbc = HexToData(StrRbc)
                
                If mLastReturn <> "" Then
                    strResult = strResult & "||" & mLastReturn & vbCrLf & "WBC;0;280,40,0-0,9-,18-100,27-,36-200" & StrWbc
                    strResult = strResult & "^" & "RBC;0;280,50,0-0,10-50,20-100,30-150,40-200" & StrRbc
                End If
            End If
        ElseIf UCase(Mid(str_Code, 2, 2)) = "D3" Then
            If mLastReturn <> "" Then
                StrPLt = Mid(str_Code, 4, 70)
                StrPLt = HexToData(StrPLt)
                
                strResult = strResult & "||" & mLastReturn & vbCrLf & "PLT;0;280,38,0-0,10-10,20-20,30-30" & StrPLt
                
                mLastReturn = ""
            End If
        End If
        
        
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = 0
        End If
        
    Loop

    
    If strResult <> "" Then strResult = Mid(strResult, 3)
    
    Exit Sub
errH:
    WriteLog "err", Err.Description, strResult
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
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
    
End Function

Private Function HexToData(ByVal strIn As String) As String
    Dim i As Integer, strData As String
    
    For i = 1 To Len(strIn) Step 2
        strData = strData & ";" & CByte("&H" & Mid(strIn, i, 2)) * 2
    Next
    HexToData = strData
End Function

Private Function GetValues(ByVal str_in As String, int_i As Integer) As String
    '功能:格式化返回值
    'str_in:待格式化的字符串
    'int_i:当入参是数字时,保留的小数位数
    
    If str_in = "*000" Then
        GetValues = "----"
    ElseIf str_in = "     " Then
        GetValues = " "
    Else
        GetValues = Val(str_in) / 10 ^ int_i
    End If
    
End Function

