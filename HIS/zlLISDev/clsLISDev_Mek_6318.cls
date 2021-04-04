VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Mek_6318"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Mek 6318
'厂商:日本光电
'作者:陈东
'用户:贵阳肿瘤医院
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
    
    Dim varline As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String
    Dim blnEnd As Boolean '是否已解析完
    Dim strImage As String  '图像数据
    Dim varImage As Variant '图像数组
    Dim lng_S As Long, lng_E As Long
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, "")
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, "")
    Else
        lng_E = InStr(mstrCode, "")
    End If
    
    Do While lng_S > 0 And lng_E > 0

        '取一个完整的记录,剩余的存到strReserved中,下次解析
        strLine = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        strImage = Mid(mstrCode, lng_E + 1)
        
        If InStr(strImage, "") > 0 Then
            strImage = Mid(strImage, 1, InStr(strImage, ""))
            strImage = Replace(strImage, "", "")
        End If
        
        varline = Split(strLine, Chr(13))
        If strImage <> "" Then
            varImage = Split(strImage, Chr(13))
            If UBound(varImage) < 768 Then
                strLine = ""
                lng_S = 0: lng_E = 0
                Exit Do
            End If
        End If
        WriteLog "收到完整数据", strLine & strImage, Replace(mstrCode, strImage, "")
        
        mstrCode = Mid(mstrCode, lng_E + 1)
        
        str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
        '解析
        
        '处理结果
        For iRow = 0 To UBound(varline)
            Select Case iRow
            Case 0 '日期
                If varline(iRow) Like "*##-*-##*" Then
                    str日期 = Mid(varline(iRow), InStr(varline(iRow), "-") - 2, 2)
                    strTmp = Mid(varline(iRow), InStr(varline(iRow), "-") + 1)
                    strTmp = Mid(strTmp, InStr(strTmp, "-") + 1, 2) & "-" & TrunMonth(Mid(strTmp, 1, InStr(strTmp, "-") - 1))
                    str日期 = Format(CDate(strTmp & "-" & str日期), "yyyy-mm-dd")
                Else
                    str日期 = Format(Now, "yyyy-MM-dd")
                End If
            Case 1 '标本号
                str标本号 = Val(varline(iRow))
                str检验人标本质控 = " |血液|0"
            Case 2 'WBC  白细胞总数
                strItem = "|WBC|" & Format_6318(varline(iRow))
            Case 3 'LYM% 淋巴细胞百分比
                strItem = strItem & "|LYMP|" & Format_6318(varline(iRow))
            Case 4 'MON% 中值细胞百分比
                strItem = strItem & "|MONP|" & Format_6318(varline(iRow))
            Case 5 'GRA% 中性细胞百分比
                strItem = strItem & "|GRAP|" & Format_6318(varline(iRow))
            Case 6 'LYM# 淋巴细胞绝对值
                strItem = strItem & "|LYMA|" & Format_6318(varline(iRow))
            Case 7 'MON# 中值细胞绝对值
                strItem = strItem & "|MONA|" & Format_6318(varline(iRow))
            Case 8 'GRA# 中性细胞绝对值
                strItem = strItem & "|GRAA|" & Format_6318(varline(iRow))
            Case 9 'EOS# 嗜酸性粒细胞绝对值
                strItem = strItem & "|EOS|" & Format_6318(varline(iRow))
            Case 10 'RBC  红细胞总数
                strItem = strItem & "|RBC|" & Format_6318(varline(iRow))
            Case 11 'HGB 血红蛋白
                strItem = strItem & "|HGB|" & Format_6318(varline(iRow))
            Case 12 'HCT  红细胞压积
                strItem = strItem & "|HCT|" & Format_6318(varline(iRow))
            Case 13 'MCV  平均红细胞体积
                strItem = strItem & "|MCV|" & Format_6318(varline(iRow))
            Case 14 'MCH  平均血红蛋白含量
                strItem = strItem & "|MCH|" & Format_6318(varline(iRow))
            Case 15 'MCHC 平均血红蛋白浓度
                strItem = strItem & "|MCHC|" & Format_6318(varline(iRow))
            Case 16 'RDW  红细胞平均宽度
                strItem = strItem & "|RDW|" & Format_6318(varline(iRow))
            Case 17 'PLT 血小板总数
                strItem = strItem & "|PLT|" & Format_6318(varline(iRow))
            Case 18 'PCT  血小板压积
                strItem = strItem & "|PCT|" & Format_6318(varline(iRow))
            Case 19 'MPV  平均血小板体积
                strItem = strItem & "|MPV|" & Format_6318(varline(iRow))
            Case 20 'PDW  血小板平均宽度
                strItem = strItem & "|PDW|" & Format_6318(varline(iRow))
            End Select
        Next
        
        '处理图形
        If strImage <> "" Then
            For iRow = 0 To UBound(varImage)
                If iRow >= 0 And iRow <= 256 Then
                    If iRow = 0 Then strImage = "^WBC;0;0,271,15-50,51-100,87-,123-200,159-,195-300,235-,271-400"
                    strImage = strImage & ";" & Val(varImage(iRow))
                End If
                If iRow > 256 And iRow <= 512 Then
                    If iRow = 257 Then strImage = strImage & "^RBC;0;0,256,18-,56-100,94-,132-200,170-,208-300,246-"
                    strImage = strImage & ";" & Val(varImage(iRow))
                End If
                If iRow > 512 And iRow <= 768 Then
                    If iRow = 513 Then strImage = strImage & "^PLT;0;0,220,9-2,42-,75-20,108-,141-40,174-,207-60"
                    strImage = strImage & ";" & Val(varImage(iRow))
                End If
            Next
        End If
        If strItem <> "" And Val(str标本号) <> 0 Then
            strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
        End If

        lng_S = InStr(mstrCode, "")
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, "")
        Else
            lng_E = InStr(mstrCode, "")
        End If
    Loop


    If strResult <> "" Then
        If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
        If strImage <> "" Then
            strResult = strResult & vbCrLf & Mid(strImage, 2)
        End If
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

Private Function Format_6318(ByVal strVal As String) As String
    '格式转换
    '有>,<号原样反回，否则转为数字返回
    If InStr(strVal, ">") > 0 Or InStr(strVal, "<") > 0 Then
        Format_6318 = Trim(strVal)
    Else
        Format_6318 = Val(Trim(strVal))
    End If
End Function


