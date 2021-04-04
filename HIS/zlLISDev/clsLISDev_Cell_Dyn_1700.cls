VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Cell_Dyn_1700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Cell Dyn 1700
'厂商:美国雅培
'作者:陈东
'用户:北京朝阳二院
Private mstrRBC As String, mstrWBC As String, mstrPLT As String '存图形数据。

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
    
    Dim varLine As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str条码 As String
    Dim blnEnd As Boolean '是否已解析完
    Dim str_Img As String, i As Integer
    '初始化返回值
    strResult = "": strReserved = "": strCmd = ""
    
    If strSource Like "***" Then
        blnEnd = False
        strReserved = strSource
        Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
            strLine = Mid(strReserved, InStr(strReserved, ""), InStr(strReserved, "") - 1)
            strReserved = Mid(strReserved, InStr(strReserved, "") + 1)
            varLine = Split(strLine, ",")
            
            str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
            If UBound(varLine) >= 73 Then
                If varLine(0) Like "*""   ""*" Then
                    '是数据行
                    strTmp = Replace(varLine(6), """", "")
                    str日期 = Format(CDate("20" & Right(strTmp, 2) & "-" & Replace(Left(strTmp, 5), "/", "-")), "yyyy-MM-dd")
                    strTmp = Replace(varLine(7), """", "")
                    str日期 = str日期 & " " & strTmp
                    
                    str标本号 = Val(Replace(varLine(8), """", ""))
                    str检验人标本质控 = " |血液|" & IIf(str标本号 >= 2000, 1, 0)
                    
                    For iRow = 16 To 37 '从16号到37号 是检测结果
                        If Val(varLine(iRow)) <> 0 Then
                            strItem = strItem & "|" & iRow & "|" & Val(varLine(iRow))
                        End If
                    Next
                    
                    If strItem <> "" And Val(str标本号) <> 0 Then
                        strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
                        
                        ''有结果，把图形数据附在后面。
                        str_Img = ""
                        
                        If mstrWBC <> "" Then
                            If Split(mstrWBC, "|")(0) = varLine(6) & varLine(7) & varLine(8) Then
                                '日期，标本号相同，附加图形数据
                                str_Img = str_Img & "^" & Split(mstrWBC, "|")(1)
                                mstrWBC = ""
                            End If
                        End If
                        
                        If mstrRBC <> "" Then
                            If Split(mstrRBC, "|")(0) = varLine(6) & varLine(7) & varLine(8) Then
                                '日期，标本号相同，附加图形数据
                                str_Img = str_Img & "^" & Split(mstrRBC, "|")(1)
                                mstrRBC = ""
                            End If
                        End If
                        
                        If mstrPLT <> "" Then
                            If Split(mstrPLT, "|")(0) = varLine(6) & varLine(7) & varLine(8) Then
                                '日期，标本号相同，附加图形数据
                                str_Img = str_Img & "^" & Split(mstrPLT, "|")(1)
                                mstrPLT = ""
                            End If
                        End If
                        If str_Img <> "" Then strResult = strResult & vbCrLf & Mid(str_Img, 2)

                    End If
                    
                ElseIf varLine(0) Like "*""PLT""*" Then
                    'PLT图形 暂不解析
                    If UBound(varLine) >= 272 Then
                        mstrPLT = varLine(6) & varLine(7) & varLine(8) & "|PLT;0;120,262,65-10,130-20,195-30,260-;"
                        For i = 16 To 272
                            mstrPLT = mstrPLT & varLine(i) & ";"
                        Next
                    End If
                ElseIf varLine(0) Like "*""RBC""*" Then
                    
                    'RBC图形 暂不解析
                    If UBound(varLine) >= 272 Then
                        mstrRBC = varLine(6) & varLine(7) & varLine(8) & "|RBC;0;120,262,52-,104-100,156-,208-200,260-;"
                        For i = 16 To 272
                            mstrRBC = mstrRBC & varLine(i) & ";"
                        Next
                    End If
                ElseIf varLine(0) Like "*""WBC""*" Then
                    'WBC图形 暂不解析
                    If UBound(varLine) >= 272 Then
                        mstrWBC = varLine(6) & varLine(7) & varLine(8) & "|WBC;0;120,262,43.33-,86.66-100,129.996-,173.326-200,216.656-,259.986-300;"
                        
                        For i = 16 To 272
                            mstrWBC = mstrWBC & varLine(i) & ";"
                        Next
                    End If
                End If
            End If
            If Not strReserved Like "***" Then
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



