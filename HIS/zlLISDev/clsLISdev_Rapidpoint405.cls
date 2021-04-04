VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_RapidPoint405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血气分析系 Rapidpoint405
'厂商:SMES 西门子
'作者:陈东
'用户:重庆市第三人民医院
Private mstrCode As String
Private mstrWaitSend As String

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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
'设备标准接口
'功能：解析数据
'参数：
'strSource：要解析的原始串
'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
'   每组检验结果以||分隔,元素之间以|分隔
    Dim lngS As Long, lngE As Long, strCode As String, strTmp As String
    Dim varGroup As Variant, iRow As Integer, strLine As String, varItem As Variant, strItem As String, strValue As String
    Dim strDate As String, strNO As String, strReturnItem As String, strType As String
    'strMsg1 = "ID_REQ13"
    'strMsg2 = "0B"
    'strMsg3 = Chr(&H2) & Chr(&H53) & Chr(&H4D) & Chr(&H50) & Chr(&H5F) & Chr(&H52) & Chr(&H45) & Chr(&H51) & Chr(&H1C) & Chr(&H1E) & Chr(&H61) & Chr(&H4D) & Chr(&H4F) & Chr(&H44) & Chr(&H1D) & Chr(&H30) & Chr(&H34) & Chr(&H30) & Chr(&H35) & Chr(&H1D) & Chr(&H1D) & Chr(&H1D) & Chr(&H1C) & Chr(&H69) & Chr(&H49) & Chr(&H49) & Chr(&H44) & Chr(&H1D) & Chr(&H31) & Chr(&H31) & Chr(&H31) & Chr(&H37) & Chr(&H38) & Chr(&H1D) & Chr(&H1D) & Chr(&H1D) & Chr(&H1C) & Chr(&H72) & Chr(&H53) & Chr(&H45) & Chr(&H51) & Chr(&H1D) & Chr(&H33) & Chr(&H36) & Chr(&H30) & Chr(&H1D) & Chr(&H1D) & Chr(&H1D) & Chr(&H1C) & Chr(&H1E) & Chr(&H3) & Chr(&H38) & Chr(&H33) & Chr(&H4)
    'ID_DATA<FS><RS>aMOD<GS>LIS<GS><GS><GS><FS>iIID<GS>333<GS><GS><GS><FS><RS>84
    'strMsg3 = "ID_DATA" & Chr(&H1C) & Chr(&H1E) & "aMOD" & Chr(&H1D) & "LIS" & Chr(&H1D) & Chr(&H1D) & Chr(&H1D) & Chr(&H1C) & "iIID" & Chr(&H1D) & "333" & Chr(&H1D) & Chr(&H1D) & Chr(&H1D) & Chr(&H1C) & Chr(&H1E) & "84"
    On Error GoTo hErr
100 mstrCode = mstrCode & strSource

102 strCode = GetCode(mstrCode, Chr(&H2), Chr(&H4))
104 Do While strCode <> ""

106     strCode = Replace(strCode, Chr(&H17), "")

108     If strCode = "ID_REQ" & Chr(&H1C) & Chr(&H1E) & "13" Then        '联机应答,分成二步
110         strCmd = "1|0B"
112         strResult = "||" & Format(Now, "YYYY-MM-DD") & "|^^0"
114         mstrWaitSend = mstrWaitSend & "|" & "ID_DATA" & Chr(&H1C) & Chr(&H1E) & "aMOD" & Chr(&H1D) & "LIS" & Chr(&H1D) & Chr(&H1D) & Chr(&H1D) & Chr(&H1C) & "iIID" & Chr(&H1D) & "333" & Chr(&H1D) & Chr(&H1D) & Chr(&H1D) & Chr(&H1C) & Chr(&H1E) & "84"

116     ElseIf strCode Like "SMP_NEW_AV*" Then    '应答,分成二步
118         strCmd = "1|0B"
120         strResult = "||" & Format(Now, "YYYY-MM-DD") & "|^^0"
122         strTmp = Chr(&H2) & "SMP_REQ" & Mid$(strCode, 12, InStr(strCode, Chr(&H3)) - 11)
124         mstrWaitSend = mstrWaitSend & "|" & strTmp & GetCheckCode(strTmp) & Chr(&H4)

126     ElseIf strCode Like "QC_NEW_AV*" Then    '质控应答,分成二步
128         strCmd = "1|0B"
130         strResult = "||" & Format(Now, "YYYY-MM-DD") & "|^^0"
132         strTmp = Chr(&H2) & "QC_REQ" & Mid$(strCode, 11, InStr(strCode, Chr(&H3)) - 10)
134         mstrWaitSend = mstrWaitSend & "|" & strTmp & GetCheckCode(strTmp) & Chr(&H4)

136     ElseIf strCode Like "SMP_NEW_DATA*" Or strCode Like "QC_NEW_DATA*" Then '添加质控解析
138         strCmd = "0B"
140         varGroup = Split(strCode, Chr(&H1C))
142         strNO = "": strReturnItem = "": strDate = "": strType = ""
144         For iRow = LBound(varGroup) To UBound(varGroup)
146             strLine = varGroup(iRow)
148             If InStr(strLine, Chr(&H1D)) > 0 Then
150                 varItem = Split(strLine, Chr(&H1D))
152                 strItem = varItem(0)
154                 If strItem = "rDATE" Then
156                     strDate = varItem(1)
158                     strDate = Mid$(strDate, 6, 4) & "-" & TrunMonth(Mid$(strDate, 3, 3)) & "-" & Mid$(strDate, 1, 2)
160                 ElseIf strItem = "rTIME" Then
162                     strDate = strDate & " " & varItem(1)
164                 ElseIf strItem = "iPID" Then
166                     strNO = varItem(1)
                        ' cP50，cSO2
                    ElseIf strItem = "rTYPE" Then    '质控数据没有标本号,则定义为1000号
                        strNO = IIf(varItem(1) = "QC", 1000, strNO)
                        strType = IIf(varItem(1) = "QC", "| |Blod|1", "| |Blod|0")
                        '500添加三个项目mnBili,iOBF,mLactate
168                 ElseIf InStr(",mpH,mPCO2,mPO2,mNa+,mK+,mCa++,mCl-,mGlucose,mLactate,mtHb,mO2Hb,mCOHb,mMetHb,mHHb,mnBili,iTEMP,cHCO3act,cHCO3std,cBE(vt),cBE(vv),ctCO2,cCa++,cAnGap,cctO2(a),cpH,cPO2,cP50,cPCO2,msO2,cHct,cO2CAP,cSO2,", "," & strItem & ",") > 0 Then
170                     strValue = Trim$(varItem(1))
172                     If strValue <> "" Then strReturnItem = strReturnItem & "|" & strItem & "|" & strValue
                    End If
                End If
            Next
174         If strNO <> "" And strDate <> "" And strReturnItem <> "" Then
176             strResult = strResult & "||" & strDate & "|" & strNO & strType & strReturnItem
            End If
178     ElseIf strCode Like "SMP_EDIT_DATA*" Then
180         strCmd = "0B"
        Else
182         strCmd = "0B"
        End If
184     strCode = GetCode(mstrCode, Chr(&H2), Chr(&H4))
    Loop
186 If strCmd <> "" Then WriteLog "405 Analyse：", "Send", strCmd
188 If strResult <> "" Then strResult = Mid$(strResult, 3)
    Exit Sub
hErr:
190 WriteLog "405 Analyse：", CStr(Erl()) & "行", Err.Description
End Sub

Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
'解析需要向仪器发送的标本信息
'   元素之间以|分隔
'   第0个元素：检验时间
'   第1个元素：样本序号
'   第2个元素：检验人
'   第3个元素：标本
'   第4个元素：是否质控品
'   第5～9元素：系统保留
'   从第10个元素开始为需要的检验项目。
    Dim intFor As Integer, intTmp As Integer
    On Error GoTo hErr
100 SendSample = ""
102 If mstrWaitSend <> "" Then


104     If Left(mstrWaitSend, 1) = "|" Then mstrWaitSend = Mid$(mstrWaitSend, 2)
106     If InStr(mstrWaitSend, "|") > 0 Then
108         SendSample = Split(mstrWaitSend, "|")(0)
110         mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, "|") + 1)
        Else
112         SendSample = mstrWaitSend
114         mstrWaitSend = ""
        End If

116     If SendSample <> "" Then
118         WriteLog "405 SendSample ", "", SendSample
120         intTmp = 0
122         For intFor = 1 To 5000
124             intTmp = intTmp + 1
            Next
        End If
    Else
126     WriteLog "405 :", "", "没有需要发送的双向指令。"
    End If
    Exit Function
hErr:
128 WriteLog "405 SendSample：", CStr(Erl()) & "行", Err.Description
End Function


Private Function GetCode(ByRef strInCode As String, ByVal strS As String, ByVal strE As String) As String
'获取一个完整的数据段
    Dim lngS As Long, lngE As Long, strCode As String
    lngE = 0
    lngS = InStr(strInCode, strS)
    If lngS > 0 Then lngE = InStr(lngS, strInCode, strE)
    If lngS > 0 And lngE > 0 Then
        strCode = Mid$(strInCode, lngS, lngE - lngS + Len(strE))
        strInCode = Mid$(strInCode, lngE + Len(strE))
    End If
    GetCode = strCode
End Function

Private Function GetCheckCode(ByVal str_in As String) As String

'获取校验码
    Dim lngAsc As Long, i As Integer

    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function

Private Function TrunMonth(ByVal strMonth As String) As String
'转换英文的月份为数字
    Select Case Mid(UCase(strMonth), 1, 3)
    Case "JAN": TrunMonth = "01"
    Case "FEB": TrunMonth = "02"
    Case "MAR": TrunMonth = "03"
    Case "APR": TrunMonth = "04"
    Case "MAY": TrunMonth = "05"
    Case "JUN": TrunMonth = "06"
    Case "JUL": TrunMonth = "07"
    Case "AUG": TrunMonth = "08"
    Case "SEP": TrunMonth = "09"
    Case "OCT": TrunMonth = "10"
    Case "NOV": TrunMonth = "11"
    Case "DEC": TrunMonth = "12"
    End Select

End Function

