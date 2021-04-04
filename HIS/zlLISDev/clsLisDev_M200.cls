VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clslisdev_M200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:M200 血凝仪
'厂商:北京世帝科学仪器公司
'作者:
'用户:
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
        Dim strCode As String, lngS As Long, lngE As Long
        Dim varData As Variant, strSampleNo As String, strDate As String, strItem As String, strValue As String
        Dim iRow As Integer
        On Error GoTo hErr
100     mstrCode = mstrCode & strSource
102     lngS = InStr(mstrCode, Chr(&H2))
104     lngE = 0
106     If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H3))
    
108     Do While lngS > 0 And lngE > 0
110         strCode = Mid$(mstrCode, lngS, lngE - lngS + 1)
112         mstrCode = Mid$(mstrCode, lngE + 1)
114         strItem = ""
            '--------------------------------------------------------------------
116         varData = Split(strCode, ";")
118         If UBound(varData) > 23 Then
120             strSampleNo = varData(20)
122             strDate = varData(5)
124             For iRow = 21 To UBound(varData) Step 4
126                 If iRow + 1 > UBound(varData) Then Exit For
128                 If varData(iRow + 1) = "INR" Then
130                     If Trim(varData(iRow)) <> "" Then strItem = strItem & "|INR|" & varData(iRow)
132                 ElseIf varData(iRow - 1) = "D-DIMER" Then
134                     If Trim(varData(iRow)) <> "" Then strItem = strItem & "|D-DIMER|" & varData(iRow)
                    Else
136                     If Trim$(varData(iRow)) <> "" And Trim(varData(iRow + 1)) <> "" Then strItem = strItem & "|" & varData(iRow) & "|" & varData(iRow + 1)
                    End If
                Next
138             If strItem <> "" And strSampleNo <> "" And strDate <> "" Then
140                 strResult = strResult & "||" & strDate & "|" & strSampleNo & "| |Blood|0" & strItem
                End If
            End If
            '--------------------------------------------------------------------
142         lngS = InStr(mstrCode, Chr(&H2))
144         lngE = 0
146         If lngS > 0 Then lngE = InStr(lngS, mstrCode, Chr(&H3))
        Loop
148     If strResult <> "" Then strResult = Mid$(strResult, 3)
        Exit Sub
hErr:
150     WriteLog "M200", strSource, CStr(Erl()) & "行," & Err.Description
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



