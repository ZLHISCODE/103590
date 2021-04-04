VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_FA100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:FA-100 尿液分析仪
'厂商:山西亚森
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
102     lng_S = InStr(mstrCode, Chr(&H2))
104     lng_E = InStr(mstrCode, Chr(&H3))
    
106     strResult = ""
108     If lng_S > 0 And lng_E > 0 Then
110         str_Data = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
112         mstrCode = Mid(mstrCode, lng_E + 1)
114         str_项目 = ""
116         str_时间 = ""
118         varline = Split(str_Data, Chr(&HD) & Chr(&HA))
120         For i = LBound(varline) To UBound(varline)
122             str_Line = varline(i)
124             If UCase(str_Line) Like "*NO.*" Then
                    '--FA200格式
126                 str_Line = UCase(str_Line)
128                 str_样本号 = Val(Mid(str_Line, InStr(str_Line, "NO.") + 7))
130             ElseIf UCase(str_Line) Like "*FA-11*" Then
                
132                 str_Line = UCase(str_Line)
134                 str_样本号 = Mid(str_Line, InStr(str_Line, "FA-11") + 5)
136                 If InStr(str_样本号, "-") > 0 Then
138                     str_样本号 = Val(Split(str_样本号, "-")(1))
                    Else
140                     str_样本号 = Val(str_样本号)
                    End If
142                 str_时间 = Format(Now, "yyyy-MM-dd HH:MM")
144             ElseIf str_Line Like "####/##/##*" Then
146                 str_时间 = Format(CDate(str_Line), "yyyy-MM-dd HH:MM")
                
148             ElseIf Not (UCase(str_Line) Like "*:*") Then
150                     str_Tmp = Trim(Mid(str_Line, 6))
152                     str_Tmp = Replace(UCase(str_Tmp), "MMOL/L", "")
153                     str_Tmp = Replace(UCase(str_Tmp), "CELLS/UL", "")
154                     str_Tmp = Replace(UCase(str_Tmp), "CELL/UL", "")
156                     str_Tmp = Replace(UCase(str_Tmp), "UMOL/L", "")
158                     str_Tmp = Replace(UCase(str_Tmp), "MG/L", "")
160                     str_Tmp = Replace(UCase(str_Tmp), "LEU/UL", "")
162                     str_Tmp = Replace(UCase(str_Tmp), "G/L", "")
                        
164                     'If str_Line Like "*+*" Then
166                     '    str_Tmp = Trim(Mid(str_Line, 6, InStr(1, Mid(str_Line, 6), "+")))
                        'End If
168                     If str_Tmp <> "" And Trim(Mid(str_Line, 2, 4)) <> "" Then
170                         str_项目 = str_项目 & "|" & Replace(Trim(Mid(str_Line, 2, 4)), "*", "") & "|" & Trim(str_Tmp)
                        End If
                End If
            
            Next
172         If Val(str_样本号) > 0 And IsDate(str_时间) And str_项目 <> "" Then
174             strResult = str_时间 & "|" & str_样本号 & "|" & " |" & "尿液" & "|0" & str_项目
            End If
        End If
        Exit Sub
errH:
176     WriteLog "FA100", CStr(Erl()) & ":" & Err.Description, ""
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


