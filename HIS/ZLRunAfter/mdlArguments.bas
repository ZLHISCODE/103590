Attribute VB_Name = "mdlArguments"
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2018/12/25
'模块           mdlArguments
'说明           命令行处理模块
'==================================================================================================
Private Const mstrCurModule     As String = "mdlArguments"          '当前模块名称
Private mcllArguments           As Collection                       '命令行集合
Private mstrCommandLine         As String                           '命令行

Public Property Get CommandLine() As String
    CommandLine = mstrCommandLine
End Property

Public Property Let CommandLine(strNewCommandLine As String)
    Dim strCommandLine As String
    strCommandLine = Trim$(strNewCommandLine)
    If mstrCommandLine <> strCommandLine Then
        mstrCommandLine = strCommandLine
        GetArguments
    End If
End Property

Public Property Get CommandArgumentsCount() As Long
    InitArguments
    CommandArgumentsCount = mcllArguments.Count
End Property

Public Property Get CommandArgument(ByVal lngIndex As Long, Optional ByVal blnReducedQuotes As Boolean) As String
    InitArguments
    If blnReducedQuotes Then
        CommandArgument = ReduceQuotes(mcllArguments(lngIndex))
    Else
        CommandArgument = mcllArguments(lngIndex)
    End If
End Property

Public Property Get CommandSwitch(ByVal strSwitch As String, Optional ByVal blnReducedQuotes As Boolean) As Variant
    Dim i As Integer, strArgument As String, strCommandSwitch As String
    
    InitArguments
    For i = 1 To mcllArguments.Count
        strArgument = mcllArguments(i)
        Select Case Left$(strArgument, 1)
            Case "-", "/"
                If Mid$(UCase$(strArgument), 2, Len(strSwitch)) = UCase$(strSwitch) Then
                    If blnReducedQuotes Then
                        strCommandSwitch = ReduceQuotes(Mid$(strArgument, Len(strSwitch) + 2))
                    Else
                        strCommandSwitch = Mid$(strArgument, Len(strSwitch) + 2)
                    End If
                    If Left$(strCommandSwitch, 1) = "=" Then
                        CommandSwitch = Trim$(Mid$(strCommandSwitch, 2))
                    Else
                        CommandSwitch = Trim$(strCommandSwitch)
                    End If
                    Exit Property
                End If
        End Select
    Next
End Property

'--------------------------------------------------------------------------------------------------
'方法           ReduceQuotes
'功能           两个双引号替换为单个双引号
'返回值         String
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Public Function ReduceQuotes(strArg As String) As String
    Dim strCurArg As String
    
    ReduceQuotes = strArg
    strCurArg = strArg
    If Left$(strCurArg, 1) = Chr$(34) Then
        If Right$(strCurArg, 1) = Chr$(34) Then
            strCurArg = Replace$(strArg, Chr$(34) & Chr$(34), Chr$(34))
            ReduceQuotes = Mid$(strCurArg, 2, Len(strCurArg) - 2)
        End If
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           GetArguments
'功能           格式化命令行
'返回值
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Sub GetArguments()
    Dim strCommandLine      As String
    Dim arrCmdParts() As String, i As Integer
    
    If Len(mstrCommandLine) = 0 Then
        mstrCommandLine = Trim$(VBA.Command$)
    End If
    If Len(mstrCommandLine) = 0 Then
        Set mcllArguments = New Collection
        Exit Sub
    End If
    strCommandLine = " " & Replace(mstrCommandLine, Chr$(34) & Chr$(34), Chr$(1)) & " "
    arrCmdParts = Split(strCommandLine, Chr$(34))
    For i = 0 To UBound(arrCmdParts)
        If i And 1 Then     '奇数进行替换
            arrCmdParts(i) = Replace$(arrCmdParts(i), " ", Chr$(2))
            arrCmdParts(i) = Replace$(arrCmdParts(i), "/", Chr$(3))
            arrCmdParts(i) = Replace$(arrCmdParts(i), "-", Chr$(4))
            arrCmdParts(i) = Chr$(34) & arrCmdParts(i) & Chr$(34)
        End If
    Next
    strCommandLine = Trim$(Join(arrCmdParts, ""))
    strCommandLine = Replace$(strCommandLine, "/", " /")
    strCommandLine = Replace$(strCommandLine, "-", " -")
    arrCmdParts = Split(strCommandLine, " ")
    Set mcllArguments = New Collection
    '还原特殊参数
    For i = 0 To UBound(arrCmdParts)
        If Len(arrCmdParts(i)) > 0 Then
            arrCmdParts(i) = Replace$(arrCmdParts(i), Chr$(1), Chr$(34) & Chr$(34))
            arrCmdParts(i) = Replace$(arrCmdParts(i), Chr$(2), " ")
            arrCmdParts(i) = Replace$(arrCmdParts(i), Chr$(3), "/")
            arrCmdParts(i) = Replace$(arrCmdParts(i), Chr$(4), "-")
            If Not InCollection(mcllArguments, arrCmdParts(i)) Then
                mcllArguments.Add arrCmdParts(i), arrCmdParts(i)
            End If
        End If
    Next
End Sub

'--------------------------------------------------------------------------------------------------
'方法           InitArguments
'功能           初始化命令行
'返回值
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Sub InitArguments()
    If mcllArguments Is Nothing Then
        GetArguments
    End If
End Sub
