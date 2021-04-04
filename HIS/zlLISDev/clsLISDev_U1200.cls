VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_U1200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:优利特1200
'厂商:
'作者:尹祖前
'用户:慈溪市桥头卫生院

Private mstrSource As String
Private mstr标本号 As String
Private mstr日期 As String

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
    Dim strLine As String, str日期 As String, str标本号 As String, str项目 As String, str结果 As String, str项目结果 As String
    Dim str质控 As String
    Dim i As Long
    Dim brr()  As String
    Dim arr()  As String
    
    mstrSource = mstrSource & strSource
    
    On Error GoTo errHand
    Do While InStr(mstrSource, "MSH") > 0
        If InStr(mstrSource, "MSH") > 0 Then
            strLine = Mid(mstrSource, 1, InStr(mstrSource, "MSH") - 1) '截取标本开始前数据 防止是上一个标本未完全解析数据
            If Len(strLine) > 0 Then
                If mstr日期 <> "" Then
                    arr = Split(strLine, Chr(13)) '数据以chr(13)分割
                    For i = LBound(arr) To UBound(arr)
                        brr = Split(arr(i), "|")
                        If UBound(brr) > 2 Then
                            Select Case brr(0)
                                Case "MSH" '肯定无此段数据
                                    str日期 = brr(6)
                                    str日期 = Mid(brr(6), 1, 4) & "-" & Mid(brr(6), 5, 2) & "-" & Mid(brr(6), 7, 2) & " " & Mid(brr(6), 9, 2) & ":" & Mid(brr(6), 11, 2) & ":" & Mid(brr(6), 13, 2)
                                    mstr日期 = str日期
                                Case "OBR"
                                    str标本号 = brr(3)
                                    mstr标本号 = str标本号
                                Case "OBX"
                                    str项目 = Mid(brr(3), InStr(brr(3), "^") + 1)
                                    str结果 = Trim(brr(5))
                                   If InStr("D_Vc,D_WBC,D_KET,D_BIL,D_PRO,D_GLU,D_BLD", str项目) > 0 Then
                                       str结果 = Trim(Mid(str结果, 1, 2))
                                       str结果 = Replace(Replace(Replace(Replace(Replace(str结果, "+-", "±"), "+2", "++"), "+3", "+++"), "+4", "++++"), "+5", "+++++")
                                   End If
                                   str项目结果 = str项目结果 & "|" & str项目 & "|" & str结果
                                Case Else
                                '不处理
                            End Select
                        End If
                    Next i
                    If str项目结果 <> "" Then
                        str质控 = "0"
                        If Val(str标本号) = "999" Then str质控 = "1"  '质控标本用指定标本号做质控
                        
                        strResult = strResult & "||" & mstr日期 & "|" & mstr标本号 & "| |尿|" & str质控 & str项目结果
                    End If
                End If
    
            End If
            mstrSource = Mid(mstrSource, InStr(mstrSource, "MSH")) '截取新的标本开始
            mstr日期 = "": mstr标本号 = "" '清空上个标本的信息
            If InStr(3, mstrSource, "MSH") > 0 Then '判断剩下数据是否多个
                strLine = Mid(mstrSource, 1, InStr(3, mstrSource, "MSH") - 1)  '多个时只取一个完整的标本
                mstrSource = Mid(mstrSource, InStr(3, mstrSource, "MSH"))  '剩余的数据下个循环解析
            Else
                strLine = mstrSource '取当前信息
                mstrSource = ""
            End If
                arr = Split(strLine, Chr(13))
                For i = LBound(arr) To UBound(arr)
                    brr = Split(arr(i), "|")
                    If UBound(brr) > 2 Then
                        Select Case brr(0)
                            Case "MSH"
                                str日期 = brr(6)
                                str日期 = Mid(brr(6), 1, 4) & "-" & Mid(brr(6), 5, 2) & "-" & Mid(brr(6), 7, 2) & " " & Mid(brr(6), 9, 2) & ":" & Mid(brr(6), 11, 2) & ":" & Mid(brr(6), 13, 2)
                                mstr日期 = str日期
                            Case "OBR"
                                str标本号 = brr(3)
                                mstr标本号 = str标本号
                            Case "OBX"
                                str项目 = Mid(brr(3), InStr(brr(3), "^") + 1)
                                str结果 = Trim(brr(5))
                               If InStr("D_Vc,D_WBC,D_KET,D_BIL,D_PRO,D_GLU,D_BLD", str项目) > 0 Then
                                   str结果 = Trim(Mid(str结果, 1, 2))
                                   str结果 = Replace(Replace(Replace(Replace(Replace(str结果, "+-", "±"), "+2", "++"), "+3", "+++"), "+4", "++++"), "+5", "+++++")
                               End If
                               str项目结果 = str项目结果 & "|" & str项目 & "|" & str结果
                            Case Else
                            '不处理
                        End Select
                    End If
                Next i
                If str项目结果 <> "" Then
                    str质控 = "0"
                    If Val(str标本号) = "999" Then str质控 = "1"  '质控标本用指定标本号做质控
                    strResult = strResult & "||" & mstr日期 & "|" & mstr标本号 & "| |尿|" & str质控 & str项目结果
                End If
        End If
    Loop
    If strResult <> "" Then
        strResult = Mid(strResult, 3) '去除最前面的 ||
        strResult = Replace(Replace(strResult, "|-|", "|- |"), "|+|", "|+ |")
    End If
    Exit Sub
errHand:
    WriteLog "clsLISDev_U1200", Err.Number & ":" & Err.Description, ""
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












