VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_RT6100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:RT6000
'厂商:雷杜
'作者:
'用户:吴川市人民医院
Private mintStp As Integer


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public str标志 As String




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
    
    Dim intLoop As Integer, i As Integer
    Dim varline As Variant
    Dim strLine As String
    Dim strData(8, 12) As String, str结果 As String, lngMod As Long, lng_E As Long, lng_S As Long, j As Long
    Dim str结果1 As String, strDate As String
    
    
    
    If mintStp = 1 Then
'        该仪器一次就将两次检测的数据发回 , 这里分成两步, 把数据返回
        strReserved = strSource
        '开始解析
        strResult = ""
        If InStr(strSource, "TRANSFER FINISH") > 0 Then
            varline = Split(strSource, "B,1,")  'B,1,的1即为在酶标仪传输时输入的1
            lng_S = LBound(varline)
            lng_E = UBound(varline)
            For i = lng_S + 1 To lng_E
                
                strLine = varline(i)
                str结果 = Split(strLine, ",")(7)
                
                lngMod = i Mod 12
                If i > 10 And lngMod = 0 Then
                
                    str结果1 = str结果1 & ";" & str结果 & "|"
                    str结果 = ""
                Else
                    str结果1 = str结果1 & ";" & str结果
                    str结果 = ""
                End If
                
            Next
            If lng_E < 96 Then
                For j = lng_E + 1 To 96
                    lngMod = j Mod 12
                    If lngMod = 0 Then
                        str结果1 = str结果1 & ";" & "0.000" & "|"
                        str结果 = ""
                    Else
                        str结果1 = str结果1 & ";" & "0.000"
                        str结果 = ""
                    End If
                Next
            End If
       End If
       
       If CheckTheChar("|", str结果1) = 8 Then
            strResult = Replace(Mid(str结果1, 1, Len(str结果1) - 1), "*", "")
            varline = Split(strResult, "|")
            For intLoop = LBound(varline) To UBound(varline)
                strDate = strDate & "|" & Mid(varline(intLoop), 2)
            Next
        strResult = Mid(strDate, 2)
        strReserved = ""
        strSource = ""
       End If
    End If
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

Public Function CmdAnalyse(strCmdType As String) As String
    '功能:  分解命令为仪器执行的命令
        
    'MsgBox strCmdType
    Select Case strCmdType
        Case "Begin"                    '开始 出板
            CmdAnalyse = ""
        Case "End"                      '结束 进板
            CmdAnalyse = ""
        Case "Revert"                   '回复命令
            CmdAnalyse = ""
        Case "Play"                     '整板测量
            If mintStp = 0 Then
                CmdAnalyse = ""
            Else
                CmdAnalyse = ""
            End If
        Case "ReadData_TimeOut"
            CmdAnalyse = 90
        Case "ReadData"                 '读板
            If mintStp = 0 Then
                CmdAnalyse = " |"
                mintStp = 1
            Else
                Sleep (500)
                CmdAnalyse = " |"
                mintStp = 1
            End If
        Case "Out"                      '进板
            CmdAnalyse = ""
        Case "In"                       '出板
            CmdAnalyse = ""
        '以上指令为固定项,必须保留,如仪器没有对应指令,设为空串即可
        '----------------------------------------------------------------------------------
        '                               '波长,按检验仪器参数设置中的波长设定对应的指令
        Case "450"                      '波长一
            CmdAnalyse = ""
        Case "620"                      '波长二
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '进板方式,按检验仪器参数设置中的进板方式设定对应的指令
        Case "连续"                     '连续
            CmdAnalyse = ""
        Case "单步"                     '逐步
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '振板频率,按检验仪器参数设置中的振板频率设定对应的指令
        Case "低"                       '
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '空白形式,按检验仪器参数设置中的空白形式设定对应的指令
        Case "空气"
            CmdAnalyse = ""
        Case "试剂"
            CmdAnalyse = ""
        Case Else

    End Select
    
End Function





Function CheckTheChar(TheChar, TheString) As Long
    '\'TheChar=\"要检测的字符串\"
    '\'TheString=\"待检测的字符串\"
    Dim n As Long
    If InStr(TheString, TheChar) Then
        For n = 1 To Len(TheString)
            If Mid(TheString, n, Len(TheChar)) = TheChar Then
                CheckTheChar = CheckTheChar + 1
            End If
        Next
    'CheckTheChar=\"这个字符\"&CheckTheChar&\"次\"
    'Else
    'CheckTheChar=\"0次\"
    End If
End Function



