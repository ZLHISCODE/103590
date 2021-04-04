VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_BC1800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:BC-1800
'厂商:深圳迈瑞
'作者:欧海胜
'用户:东莞高埗医院

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    'GetStartCmd = Chr(6)
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
        '   图形数据放在检验结果后，用回车换行符(vbcrlf)分隔
        'strReserved：最后不能完成解析的原始串，需要返回作后续处理
        'strCmd：如果需要，可返回向设备发送的命令
 
        On Error GoTo Analyse_Err

  
        Dim strItem As String, str日期 As String, str标本号 As String
        Dim str检验人标本质控 As String, strLine As String
        Dim lngLength As Long
        Dim blnEnd As Boolean '是否已解析完
   
    
        '初始化返回值
100     strResult = "": strReserved = "": strCmd = ""
102     If strSource Like "**" & Chr(&H1A&) & "*" Then
104         blnEnd = False
106         strReserved = strSource
108         Do While Not blnEnd
                '取一个完整的记录,剩余的存到strReserved中,下次解析
110             lngLength = InStr(strReserved, Chr(&H1A&)) - InStr(strReserved, "") + 1
112             strLine = Mid(strReserved, InStr(strReserved, ""), lngLength)
114             strReserved = Mid(strReserved, InStr(strReserved, Chr(&H1A&)) + 1)
            
116             str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = ""
                '解析
            
118             strItem = ""
120             If strLine Like "A*" Then
122                 str日期 = Mid(strLine, 16, 4) & "-" & Mid(strLine, 12, 2) & "-" & Mid(strLine, 14, 2) & " " & Mid(strLine, 20, 2) & ":" & Mid(strLine, 22, 2)
124                 str标本号 = Val(Mid(strLine, 3, 8))
126                 str检验人标本质控 = " |血液|0"
128                 strItem = strItem & "|WBC|" & Val(Mid(strLine, 24, 4)) / (10 ^ 1)    'WBC,24,4,3
130                 strItem = strItem & "|LYMPHA|" & Val(Mid(strLine, 28, 4)) / (10 ^ 1) 'LYMPH#,28,4,3
132                 strItem = strItem & "|MIDA|" & Val(Mid(strLine, 32, 4)) / (10 ^ 1)   'Mid#,32,4,3
134                 strItem = strItem & "|GRANA|" & Val(Mid(strLine, 36, 4)) / (10 ^ 1)  'Gran#,36,4,3
136                 strItem = strItem & "|LYMPHP|" & Val(Mid(strLine, 40, 3)) / (10 ^ 1) 'Lymph%#,40,3,1
138                 strItem = strItem & "|MIDP|" & Val(Mid(strLine, 43, 3)) / (10 ^ 1)   'Mid%,43,3,1
140                 strItem = strItem & "|GRANP|" & Val(Mid(strLine, 46, 3)) / (10 ^ 1)  'Gran%,46,3,1
142                 strItem = strItem & "|RBC|" & Val(Mid(strLine, 49, 3)) / (10 ^ 2)    'RBC,49,3,2
144                 strItem = strItem & "|HGB|" & Val(Mid(strLine, 52, 4)) / (10 ^ 1)    'HGB,52,4,1
146                 strItem = strItem & "|MCHC|" & Val(Mid(strLine, 56, 3))              'MCHC,56,4,1
148                 strItem = strItem & "|MCV|" & Val(Mid(strLine, 59, 4)) / (10 ^ 1)    'MCV,60,4,2
150                 strItem = strItem & "|MCH|" & Val(Mid(strLine, 63, 4)) / (10 ^ 1)    'MCH,64,3,1
152                 strItem = strItem & "|RDWCV|" & Val(Mid(strLine, 67, 3)) / (10 ^ 1)  'RDW-CV,67,3,1
154                 strItem = strItem & "|HCT|" & Val(Mid(strLine, 70, 4)) / (10 ^ 4)    'HCT,70,4,2
156                 strItem = strItem & "|PLT|" & Val(Mid(strLine, 74, 3))               'PLT,74,3,0
158                 strItem = strItem & "|MPV|" & Val(Mid(strLine, 77, 3)) / (10 ^ 1)    'MPV,77,3,1
160                 strItem = strItem & "|PDW|" & Val(Mid(strLine, 80, 3)) / (10 ^ 1)    'PDW,80,3,1
162                 strItem = strItem & "|PCT|" & Val(Mid(strLine, 83, 4)) / (10 ^ 4)    'PCT,83,4,4
164                 strItem = strItem & "|RDWSD|" & Val(Mid(strLine, 87, 3)) / (10 ^ 1)  'RDW-SD,87,3,1
                End If
            
166             If strItem <> "" And Val(str标本号) <> 0 Then
168                 strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
                End If
            
170             If Not strReserved Like "**" & Chr(&H1A&) & "*" Then
172                 blnEnd = True
                End If
            Loop
        Else
            '收到的信息不全,返回,等下次再解析
174         strReserved = strSource
        End If
176     If strResult <> "" Then
178         If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
        End If

        Exit Sub

Analyse_Err:
    If Error = 1 Then
        Resume
    End If
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
    strSample = ""
    iSendStep = 0
    blnSuccess = True
End Function




