VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_DM2_Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'DM2  流水线
'临汾人民医院


Private mstrCode As String    '待解析串
Private mlngStep_In As Long   '输入字符串解析步骤
Private mstrACK As String     '是否应答ACK,是则应答ACK(06),否-应答 ETX(03)

Private mlngStep_Out As Long  '传双向消息步骤
Private mstr仪器号 As String

Private mstrWaitSend As String '待发送标本
Private mblnSendCmd  As Boolean

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
        'strReserved：最后不能完成解析的原始串，需要返回作后续处理
        'strCmd：如果需要，可返回向设备发送的命令
        
        mstrCode = mstrCode & strSource
        Dim blnSendCmd   As Boolean
        Dim lngS  As Long
        Dim strSginSample  As String
        Dim lngE  As Long
        Dim strSample() As String   '以回车换行结束
        Dim strRsultTemp  As String  '检验指标结果
        Dim i As Integer
        Dim strDate  As String       '检验日期
        Dim strNO    As String       '标本序号
        Dim intEmergency   As Integer '急诊
        Dim strSampleType  As String  '标本类型
'        If InStr(strSource, Chr(&H4) & Chr(&H1)) > 0 Then
'            strCmd = "0|" & Chr(&H6)
'            Exit Sub
'        ElseIf InStr(Mid(strSource, 4, 8), ",702,01,") > 0 Then
'            trCmd = "0|" & Chr(&H3)
'            Exit Sub
'        ElseIf InStr(Mid(strSource, 4, 8), ",702,01,") > 0 Then
'
'        End If
        If Not mblnSendCmd Then
            strCmd = "0|" & Chr(&H6)
            mblnSendCmd = Not mblnSendCmd
        Else
            strCmd = "0|" & Chr(&H3)
            mblnSendCmd = Not mblnSendCmd
        End If
        lngS = InStr(mstrCode, Chr(&H4) & Chr(&H1))
        If lngS > 0 Then
            lngE = InStr(lngS + 2, mstrCode, Chr(&H4))
        End If
        If lngS > 0 And lngE > 0 Then
            strSginSample = Mid(mstrCode, lngS, lngE)
            mstrCode = Mid(mstrCode, lngE + 1)
            strSginSample = Replace(Replace(strSginSample, Chr(&H4), ""), Chr(&H1), "")
            strSample = Split(strSginSample, vbCrLf)
            For i = 0 To UBound(strSample)
                If InStr(Mid(strSample(i), 4, 8), ",702,01,") > 0 Then
                    strDate = "20" & Mid(Split(strSample(i), ",")(3), 5, 2) & "-" & Mid(Split(strSample(i), ",")(3), 3, 2) & "-" & Mid(Split(strSample(i), ",")(3), 1, 2)
                    strNO = Trim(Split(strSample(i), ",")(12))
                    intEmergency = IIf(Split(strSample(i), ",")(9) = "ST", 1, 0)
                    strSampleType = GetSampleClass(Split(strSample(i), ",")(9))
                ElseIf InStr(Mid(strSample(i), 4, 8), ",702,03,") > 0 Then
                    '指标结果信息
                    strRsultTemp = strRsultTemp & Split(strSample(i), ",")(10) & "|" & Split(strSample(i), ",")(15) & "|"
                End If
            Next
            
            strResult = strResult & "||" & strDate & "|^" & intEmergency & "^" & strNO & "| |" & strSampleType & "|0|" & Mid(strRsultTemp, 1, Len(strResult) - 1)
        End If
        
End Sub


Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
        'iSendStep：当前发送操作的步骤。如果处理完后发送结束，则返回0，否则自动＋1返回。
        'strResponse：当前仪器的响应内容
        'strSampleInfo：发送的标本信息。
        '   '   第0个元素：检验时间
            '   第1个元素：样本序号
            '   第2个元素：病人姓名
            '   第3个元素：标本类型
            '   第4个元素：急诊标志
            '   第5个元素：样本条码
            '   第6个元素：盘号，杯号
            '   第7个元素：病人ID^性别^出生日期^年龄数字^年龄单位^姓名全拼^稀释倍数^科室
            '   第8～9元素：系统保留
            '   从第10个元素开始为需要的检验项目。
        Dim strSampleItem()  As String  '标本信息
        Dim strTemp()        As String  '每个标本以竖线风格的段
        Dim strPatient()     As String  '病人各个字段信息，包含开单科室，临床诊断等
        Dim i                As Integer
        Dim j                As Integer
        Dim intItemCount     As Integer    ' 检验项目数
        Dim strItem          As String     ' 检验项目串
        Dim strCheck         As String     '校验串
        
        blnSuccess = True
        If iSendStep = 0 Then
            strSampleItem = Split(strSampleInfo, "||")
            For i = LBound(strSampleItem) To UBound(strSampleItem)
                strTemp = Split(strSampleItem(i), "|")
                strPatient = Split(strTemp(7), "^")
                If UBound(strTemp) >= 10 Then
                    mstrWaitSend = mstrWaitSend & Chr(&H4) & Chr(&H1) & "|"
                    strCheck = "[00" & ",701,01,0000,00,1," & IIf(Val(strTemp(4)) = 1, "ST", "RO") & "," & _
                                GetSampleClass(strTemp(3)) & "," & strTemp(5) & "," & Space(20) & "诊断" & Space(25 - Len("诊断")) & "," & _
                                Space(25) & "," & strTemp(2) & Space(16 - Len(strTemp(2))) & "," & Space(15) & " ," & strPatient(0) & Space(12 - Len(strPatient(0))) & "," & _
                                Space(18) & "," & "000000," & "0000," & "科室" & Space(16 - Len("科室")) & "," & Space(3 - Len(strPatient(3))) & strPatient(3) & "," & _
                                GetAge(strPatient(4)) & ",000000," & GetSex(strPatient(3)) & "," & Space(25) & "," & Space(7) & ",    ,          ,"
                    intItemCount = UBound(strTemp) - 9
                    For j = 10 To UBound(strTemp)
                        strItem = strItem & GetTDM(strTemp(j)) & ",0"
                    Next
                    strCheck = strCheck & intItemCount & "," & strItem & "]"
                    mstrWaitSend = mstrWaitSend & GetCheckCode(strCheck) & vbCrLf & "|" & Chr(&H4) & "|"
                End If
                intItemCount = 0
                strCheck = ""
            Next
            
            '进入双向
            If InStr(mstrWaitSend, "|") > 0 Then
                SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
                iSendStep = 1
                mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|") + 1)
            End If
        Else
            If InStr(strResponse, Chr(&H6)) > 0 Or InStr(strResponse, Chr(3)) > 0 Then
                SendSample = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, "|") - 1)
                mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|") + 1)
            End If
            If mstrWaitSend = "" Then
                iSendStep = 0
            End If
       End If
        
End Function


Private Function ACKChar() As String
    If mstrACK = Chr(&H6) Then
        mstrACK = Chr(&H3)
    Else
        mstrACK = Chr(&H6)
    End If
    ACKChar = mstrACK
End Function

Private Function GetTDM(ByVal strIn As String) As String
    GetTDM = strIn & Space(4 - Len(strIn))
End Function


Private Function GetCheckCode(ByVal str_in As String) As String
    '计算校验码
    Dim ls_ch As String
    Dim ll_i As Long
    Dim ll_len As Long
    Dim ll_sum As Long

    If ((IsNull(str_in)) Or (str_in = "")) Then
        GetCheckCode = ""
    End If

    ll_len = Len(str_in)

    For ll_i = 1 To ll_len
        ls_ch = Mid(str_in, ll_i, 1)
        ll_sum = ll_sum + Asc(ls_ch)
    Next

    ll_sum = 256 - (ll_sum Mod 256)
    GetCheckCode = Hex(ll_sum)
    If Len(GetCheckCode) = 1 Then GetCheckCode = "0" & GetCheckCode
    
End Function


Public Function GetSampleClass(ByVal strClass As String) As String
    Select Case Trim(strClass)
        Case "血清"
            GetSampleClass = "SE"
        Case "脑脊液"
            GetSampleClass = "SF"
        Case "尿液"
            GetSampleClass = "UR"
        Case "血浆"
            GetSampleClass = "PL"
        Case "血液"
            GetSampleClass = "BL"
        Case "其他"
            GetSampleClass = "OT"
        Case "SE"
            GetSampleClass = "血清"
        Case "SF"
            GetSampleClass = "脑脊液"
        Case "UR"
            GetSampleClass = "尿液"
        Case "PL"
            GetSampleClass = "血浆"
        Case "BL"
            GetSampleClass = "血液"
        Case "OT"
            GetSampleClass = "其他"
    End Select
End Function


Private Function GetSex(ByVal strSex As String) As String
    '性别
    Select Case Trim(strSex)
        Case "男"
            GetSex = "M"
        Case "女"
            GetSex = "F"
        Case "不区分"
            GetSex = "B"
    End Select
End Function

Private Function GetAge(ByVal strAge As String) As Integer
    '根据年龄单位转为数字
    If Mid(strAge, 1, 1) = "" Or Mid(strAge, 1, 1) = "岁" Then
        GetAge = 5
    ElseIf Mid(strAge, 1, 1) = "月" Then
        GetAge = 4
    ElseIf Mid(strAge, 1, 1) = "周" Then
        GetAge = 3
    ElseIf Mid(strAge, 1, 1) = "天" Then
        GetAge = 2
    ElseIf Mid(strAge, 1, 1) = "时" Then
        GetAge = 1
    End If
End Function






