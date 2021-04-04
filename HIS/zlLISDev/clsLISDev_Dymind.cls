VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Dymind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'产品名称: Dx-CRP 血细胞分析仪
'产品型号: Dx-CRP
'厂商:深圳市帝迈生物技术有限公司
'作者:陈东
'日期:2016-01-05

'用户:妇儿医院

Private mstrCode As String  '待解析数据
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
        '2010-01-14|1| |血清|0|Alb|44|ALT|6|TP|72|T－Bil|10|D－Bil|2|AST|9|ALP|118|IBil|8|GLB|28|A/G|1.57

        Dim str_Code As String '1个完整的消息
        Dim var_Code As Variant
        Dim lng_count As Long
    
        Dim str_字段分隔符 As String
        Dim str_组件分隔符 As String
        Dim str_子组件隔符 As String
        Dim str_重复分隔符 As String
        Dim str_转义字符 As String
    
        Dim str_Segment As String   '一个消息段(以chr(&HD)分隔)
        Dim var_Filed As Variant    '字段
        Dim str_唯一标识 As String  '用于标识一个消息
        Dim str_版本 As String
    
        Dim str_消息类型 As String  '用于 应答
        Dim str_仪器型号 As String
        Dim str_消息标识  As String
    
        Dim lng_S As Long
        Dim lng_E As Long
    
        Dim str_标本号 As String, str_日期 As String
        Dim str_质控 As String
        Dim str_项目 As String
        Dim strItem As String
        Dim strValue As String
        
        On Error GoTo errHandle
100     strResult = ""
102     If strSource <> "" Then mstrCode = mstrCode & strSource
 
104     lng_S = InStr(mstrCode, Chr(&HB)): lng_E = 0
106     If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&HD) & Chr(&H1C))
        
108     Do While lng_S > 0 And lng_E > 0
110         str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
112         mstrCode = Mid(mstrCode, lng_E + 2)
                 
114         var_Code = Split(str_Code, Chr(&HD))
116         For lng_count = LBound(var_Code) To UBound(var_Code)
118             str_Segment = var_Code(lng_count)
120             If Mid(str_Segment, 2, 3) = "MSH" Then
                    '消息头
122                 str_字段分隔符 = Mid(str_Segment, 5, 1)
124                 str_组件分隔符 = Mid(str_Segment, 6, 1)
126                 str_子组件隔符 = Mid(str_Segment, 7, 1)
128                 str_重复分隔符 = Mid(str_Segment, 8, 1)
130                 str_转义字符 = Mid(str_Segment, 9, 1)
                
132                 var_Filed = Split(str_Segment, str_字段分隔符)
134                 str_唯一标识 = var_Filed(9)
136                 str_质控 = var_Filed(10) 'P-标本，Q-质控
138                 str_版本 = var_Filed(11)
                
140                 strCmd = "MSH|^~\&|LIS|ZLHIS||||" & Format(Now, "yyyyMMddHHmmss") & "||ACK^R01|HIS" & Format(Now, "yyyyMMddHHmmss") & "|P|2.3.1||||||UNICODE" & Chr(&HD)
142                 strCmd = strCmd & "MSA|AA|" & str_唯一标识 & Chr(&HD) & Chr(&H1C) & Chr(&HD)
                    
'144                 If str_版本 <> "2.3.1" Then WriteLog "版本不是2.3.1"
                
146                 str_项目 = "": str_标本号 = "": str_日期 = ""
148             ElseIf Mid(str_Segment, 1, 3) = "OBR" Then
                    '标本号，日期
150                 var_Filed = Split(str_Segment, str_字段分隔符)
152                 str_标本号 = var_Filed(3)
154                 str_日期 = var_Filed(7)
156                 str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2)
                
158                 If str_质控 = "Q" Then
                        '质控
160                     str_项目 = "|" & var_Filed(3) & "|" & var_Filed(20)       '项目编号,结果
162                     str_标本号 = var_Filed(12)  '质控液编号
 
                        
                    End If
164             ElseIf Mid(str_Segment, 1, 3) = "OBX" Then
166                 var_Filed = Split(str_Segment, str_字段分隔符)
                     
168                 If var_Filed(2) = "NM" And var_Filed(5) <> "" Then
170                     strItem = var_Filed(3)
172                     strItem = Trim(Split(strItem, "^")(1))
174                     strValue = Trim$(var_Filed(5))
176                     If strItem <> "" And strValue <> "" Then str_项目 = str_项目 & "|" & strItem & "|" & strValue
                    End If
                End If
            Next
        
 
178         If Val(str_标本号) > 0 And str_项目 <> "" Then
180             strResult = strResult & "||" & str_日期 & "|" & Val(str_标本号) & "| |Boold|" & IIf(str_质控 = "2", 1, 0) & str_项目
            End If
      
182         str_项目 = ""
184         lng_S = InStr(mstrCode, Chr(&HB)): lng_E = 0
186         If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&HD) & Chr(&H1C))
        Loop
188     If strResult <> "" Then strResult = Mid(strResult, 3)
        Exit Sub
errHandle:
'190     WriteLog "Dy " & Erl() & "行, " & Err.Description
End Sub


Public Function SendSample(strSampleInfo As String, _
                           iSendStep As Integer, _
                           blnSuccess As Boolean, _
                           Optional ByVal strResponse As String = "", _
                           Optional ByVal blnUndo As Boolean = False, _
                           Optional ByVal iType As Integer = 0) As String
        '解析需要向仪器发送的标本信息
        '   元素之间以|分隔
        '   第0个元素：检验时间
        '   第1个元素：样本序号
        '   第2个元素：检验人
        '   第3个元素：标本
        '   第4个元素：是否质控品
        '   第5～9元素：系统保留
        '   从第10个元素开始为需要的检验项目。
    

End Function





