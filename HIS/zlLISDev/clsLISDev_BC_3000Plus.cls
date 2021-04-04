VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_BC_3000Plus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:BC-3000plus
'厂商:深圳迈瑞
'作者:陈东
'用户:成都儿童专科医院(东城根)

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
    '格式一(成都儿童)
    'A000000170080220070853006300210007003533211355547712802960907026816543201591031681630530000000000000
    '格式二（上海曲阳）
    'AAAI10P19000000000137010302007145401570019001101271190738083691160324097003141613570189074157139054000000000000000000000

    Dim varLine As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, lngLength As Long, lngType As Long
    Dim blnEnd As Boolean '是否已解析完
    Dim StrRbc As String, StrWbc As String, StrPLt As String
    
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    If strSource Like "**" & Chr(&H1A&) & "*" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            lngLength = InStr(strReserved, Chr(&H1A&)) - InStr(strReserved, "") + 1
            strLine = Mid(strReserved, InStr(strReserved, ""), lngLength)
            strReserved = Mid(strReserved, InStr(strReserved, Chr(&H1A&)) + 1)
            
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            '解析
            
            strItem = ""
           
            If strLine Like "A*" Then
                str日期 = Mid(strLine, 16, 4) & "-" & Mid(strLine, 12, 2) & "-" & Mid(strLine, 14, 2)
                
                str标本号 = Val(Mid(strLine, 3, 8))
                If str标本号 <> 0 And IsDate(str日期) Then
                    lngType = 1
                Else
                    '另一种格式
                    str标本号 = Val(Mid(strLine, 11, 12))
                    str日期 = Mid(strLine, 28, 4) & "-" & Mid(strLine, 24, 2) & "-" & Mid(strLine, 26, 2)
                    If str标本号 <> 0 And IsDate(str日期) Then
                        lngType = 2
                    End If
                End If
                str检验人标本质控 = " |血液|0"
                Select Case lngType
                Case 1
                    strItem = strItem & "|WBC|" & Val(Mid(strLine, 24, 4)) / (10 ^ 1)    'WBC,24,4,3
                    strItem = strItem & "|LYMPHA|" & Val(Mid(strLine, 28, 4)) / (10 ^ 1) 'LYMPH#,28,4,3
                    strItem = strItem & "|MIDA|" & Val(Mid(strLine, 32, 4)) / (10 ^ 1)   'Mid#,32,4,3
                    strItem = strItem & "|GRANA|" & Val(Mid(strLine, 36, 4)) / (10 ^ 1)  'Gran#,36,4,3
                    strItem = strItem & "|LYMPHP|" & Val(Mid(strLine, 40, 3)) / (10 ^ 1) 'Lymph%#,40,3,1
                    strItem = strItem & "|MIDP|" & Val(Mid(strLine, 43, 3)) / (10 ^ 1)   'Mid%,43,3,1
                    strItem = strItem & "|GRANP|" & Val(Mid(strLine, 46, 3)) / (10 ^ 1)  'Gran%,46,3,1
                    strItem = strItem & "|RBC|" & Val(Mid(strLine, 49, 3)) / (10 ^ 2)    'RBC,49,3,2
                    strItem = strItem & "|HGB|" & Val(Mid(strLine, 52, 4)) / (10 ^ 1)    'HGB,52,4,1
                    strItem = strItem & "|MCHC|" & Val(Mid(strLine, 56, 3))              'MCHC,56,4,1
                    strItem = strItem & "|MCV|" & Val(Mid(strLine, 59, 4)) / (10 ^ 1)    'MCV,60,4,2
                    strItem = strItem & "|MCH|" & Val(Mid(strLine, 63, 4)) / (10 ^ 1)    'MCH,64,3,1
                    strItem = strItem & "|RDWCV|" & Val(Mid(strLine, 67, 3)) / (10 ^ 1)  'RDW-CV,67,3,1
                    strItem = strItem & "|HCT|" & Val(Mid(strLine, 70, 4)) / (10 ^ 4)    'HCT,70,4,2
                    strItem = strItem & "|PLT|" & Val(Mid(strLine, 74, 3))               'PLT,74,3,0
                    strItem = strItem & "|MPV|" & Val(Mid(strLine, 77, 3)) / (10 ^ 1)    'MPV,77,3,1
                    strItem = strItem & "|PDW|" & Val(Mid(strLine, 80, 3)) / (10 ^ 1)    'PDW,80,3,1
                    strItem = strItem & "|PCT|" & Val(Mid(strLine, 83, 4)) / (10 ^ 4)    'PCT,83,4,4
                    strItem = strItem & "|RDWSD|" & Val(Mid(strLine, 87, 3)) / (10 ^ 1)  'RDW-SD,87,3,1
                    StrWbc = Mid(strLine, 162, 795)
                    StrRbc = Mid(strLine, 957, 750)
                    StrPLt = Mid(strLine, 1707, 660)
                Case 2
                    strItem = strItem & "|WBC|" & Val(Mid(strLine, 36, 4)) / (10 ^ 1)    'WBC,36,4,1
                    strItem = strItem & "|LYMPHA|" & Val(Mid(strLine, 40, 4)) / (10 ^ 1) 'LYMPH#,40,4,1
                    strItem = strItem & "|MONA|" & Val(Mid(strLine, 44, 4)) / (10 ^ 1)   'Mon#,32,4,3
                    strItem = strItem & "|GRANA|" & Val(Mid(strLine, 48, 4)) / (10 ^ 1)  'Gran#,36,4,3
                    strItem = strItem & "|LYMPHP|" & Val(Mid(strLine, 52, 3)) / (10 ^ 1) 'Lymph%#,40,3,1
                    strItem = strItem & "|MONP|" & Val(Mid(strLine, 55, 3)) / (10 ^ 1)   'Mon%,43,3,1
                    strItem = strItem & "|GRANP|" & Val(Mid(strLine, 58, 3)) / (10 ^ 1)  'Gran%,46,3,1
                    strItem = strItem & "|RBC|" & Val(Mid(strLine, 61, 3)) / (10 ^ 2)    'RBC,49,3,2
                    strItem = strItem & "|HGB|" & Val(Mid(strLine, 64, 4)) / (10 ^ 1)    'HGB,52,4,1
                    strItem = strItem & "|MCHC|" & Val(Mid(strLine, 68, 3))              'MCHC,56,3,1
                    strItem = strItem & "|MCV|" & Val(Mid(strLine, 71, 4)) / (10 ^ 1)    'MCV,60,4,2
                    strItem = strItem & "|MCH|" & Val(Mid(strLine, 75, 4)) / (10 ^ 1)    'MCH,64,3,1
                    strItem = strItem & "|RDWCV|" & Val(Mid(strLine, 79, 3)) / (10 ^ 1)  'RDW-CV,67,3,1
                    strItem = strItem & "|HCT|" & Val(Mid(strLine, 82, 4)) / (10 ^ 4)    'HCT,70,4,2
                    strItem = strItem & "|PLT|" & Val(Mid(strLine, 86, 3))               'PLT,74,3,0
                    strItem = strItem & "|MPV|" & Val(Mid(strLine, 89, 3)) / (10 ^ 1)    'MPV,77,3,1
                    strItem = strItem & "|PDW|" & Val(Mid(strLine, 92, 3)) / (10 ^ 1)    'PDW,80,3,1
                    strItem = strItem & "|PCT|" & Val(Mid(strLine, 95, 4)) / (10 ^ 4)    'PCT,83,4,4
                    strItem = strItem & "|RDWSD|" & Val(Mid(strLine, 99, 3)) / (10 ^ 1)  'RDW-SD,87,3,1
                    
                    StrWbc = Mid(strLine, 162, 795)
                    StrRbc = Mid(strLine, 957, 750)
                    StrPLt = Mid(strLine, 1707, 660)
                End Select

            End If
            
            If strItem <> "" And Val(str标本号) <> 0 Then
                strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
                
                strResult = strResult & vbCrLf & "WBC;0;256,340,25-,60-100,95-,130-200,165-,200-300,235-,270-400" & BC3000(StrWbc)
                strResult = strResult & "^RBC;0;256,370,20-50,70-100,120-150,170-200,220-,270-300,320-" & BC3000(StrRbc)
                strResult = strResult & "^PLT;0;256,370,30-2,90-,140-20,190-,240-40,300-,360-60" & BC3000(StrPLt)
            End If
            
            If Not strReserved Like "**" & Chr(&H1A&) & "*" Then
                blnEnd = True
            End If
        Loop
    Else
        '收到的信息不全,返回,等下次再解析
        strReserved = strSource
    End If
    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    'strCmd = Chr(6)
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

Private Function BC3000(ByVal strIn As String) As String
    Dim strImg As String
    Dim lngI As Long
    
    For lngI = 0 To Len(strIn) / 3 - 1
        strImg = strImg & ";" & Mid(strIn, 3 * lngI + 1, 3)
    Next
    BC3000 = strImg
    
End Function


