VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_OP155"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:OP-155 尿液分析仪
'厂商:无锡市欧普兰科技有限公司
'作者:陈东
'用户:万荣县医院（太原）
 
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
        '设备标准接口
        '功能：解析数据
        '参数：
        'strSource：要解析的原始串
        'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
        '   每组检验结果以||分隔,元素之间以|分隔
        '   第0个元素：检验时间
        '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
        '   第2个元素：检验人
        '   第3个元素：标本
        '   第4个元素：是否质控品
        '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
        '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
        'strReserved：最后不能完成解析的原始串，需要返回作后续处理
        'strCmd：如果需要，可返回向设备发送的命令
    
        '请在这个过程中写解析程序，上面的注释行请不要删除。
        Dim str_时间 As String, str_样本号 As String, str_项目 As String, str_结果 As String
        Dim lng_S As Long, lng_E As Long, str_Tmp As String
        Dim str_Data As String
        Dim varline As Variant
        Dim i As Integer, str_Line As String
        On Error GoTo errH
    
100     mstrCode = mstrCode & strSource
        
102     lng_S = InStr(mstrCode, "-")
104     lng_S = lng_S - 4
106     If lng_S > 0 Then
108         If IsNumeric(Mid(mstrCode, lng_S, 4)) Then
110             lng_E = InStr(lng_S, mstrCode, "VC")
112             lng_E = InStr(lng_E, mstrCode, Chr(&HD) & Chr(&HA))
            Else
114             lng_S = 0
            End If
        End If
    
116     strResult = ""
118     Do While lng_S > 0 And lng_E > 0
            
120         str_Data = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
122         mstrCode = Mid(mstrCode, lng_E + 2)
124         str_项目 = ""
126         str_时间 = ""
128         varline = Split(str_Data, Chr(&HD) & Chr(&HA))
130         For i = LBound(varline) To UBound(varline)
132             str_Line = varline(i)
134             If str_Line Like "####-##-##*" Then
136                 str_时间 = Format(CDate(str_Line), "yyyy-MM-dd") & " " & Format(Now, "HH:mm:ss")
138             ElseIf UCase(str_Line) Like "###*" Then
140                 str_样本号 = Val(str_Line)

                
142             ElseIf Not (UCase(str_Line) Like "*:*") Then
144                     str_Tmp = Trim(Mid(str_Line, 6))
146                     If str_Tmp <> "" And Trim(Mid(str_Line, 1, 3)) <> "" Then
148                         str_项目 = str_项目 & "|" & Trim(Mid(str_Line, 1, 3)) & "|" & Trim(str_Tmp)
                        End If
                End If
            
            Next
150         If Val(str_样本号) > 0 And IsDate(str_时间) And str_项目 <> "" Then
152             strResult = str_时间 & "|" & str_样本号 & "|" & " |" & "尿液" & "|0" & str_项目
            End If
            
154         lng_S = InStr(mstrCode, "-")
156         lng_S = lng_S - 4
158         If lng_S > 0 Then
160             If IsNumeric(Mid(mstrCode, lng_S, 4)) Then
162                 lng_E = InStr(lng_S, mstrCode, "VC")
164                 lng_E = InStr(lng_E, mstrCode, Chr(&HD) & Chr(&HA))
                Else
166                 lng_S = 0
                End If
            End If
            
        Loop
        Exit Sub
errH:
168     WriteLog "op155", CStr(Erl()) & ":" & Err.Description, ""
End Sub


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






