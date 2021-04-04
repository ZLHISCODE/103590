Attribute VB_Name = "mdlErrCenter"
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/1/14
'模块           mdlErrCenter
'说明           该程序是后台程序，不会弹出错误提示，因此需要将错误记录。
'==================================================================================================
'Private mobjLog                 As clsLog
Private Const M_MAX_LOG_COUNT   As Long = 8000                      '日志最多记录8000行，每次开始清除，或者满8000行清除
Private mlngRecCount            As Long                             '日志记录条数
Private Const mlngStackLen      As Integer = 40                     '调用堆栈的长度
Private mcllMethodStack         As New Collection                   '调用堆栈集合
Private mstrText                As String
Private mlngIndex               As Integer
'--------------------------------------------------------------------------------------------------
'方法           LogName
'功能           设置日志名称
'返回值
'入参列表:
'参数名         类型                    说明
'strLogName     String                  初次调用生成日志。当传入的名称为空，关闭日志
'-------------------------------------------------------------------------------------------------
Public Property Let LogName(strLogName As String)
'    If mobjLog Is Nothing Then
'        If strLogName <> "" Then
'            Set mobjLog = New clsLog
'            Call mobjLog.OpenLog(strLogName, , False)
'        End If
'    ElseIf strLogName = "" Then
'        mobjLog.CloseLog
'        Set mobjLog = Nothing
'    End If
End Property
'--------------------------------------------------------------------------------------------------
'方法           ErrorCenter
'功能           错误处理中心
'返回值         Integer                 0-忽略继续执行，1-重试(Resume),2-中止程序
'入参列表:
'参数名         类型                    说明
'strMethod      String                  错误发生的过程
'-------------------------------------------------------------------------------------------------
Public Function ErrorCenter(Optional ByRef strMethod As String) As Integer
'    mobjLog.WriteOperate String((mcllMethodStack.Count - 1) * 2, " ") & "┣" & strMethod, Err.Number & "-" & Err.Description
'    mobjLog.WriteListTitle String((mcllMethodStack.Count - 1) * 2, " ") & "┣" & "调用堆栈："
'    For mlngIndex = 1 To mcllMethodStack.Count
'        mobjLog.WriteList String((mcllMethodStack.Count - 1) * 2, " ") & "┣" & mcllMethodStack(mlngIndex)
'    Next
'    Call PopMethod(strMethod)
'    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------
'方法           WarnInfo
'功能           警告处理。该种错误可能只生成警告，并不进行错误捕获。
'返回值
'入参列表:
'参数名         类型                    说明
'strWarnInfo    String                  警告信息
'strMethod      String                  错误发生的过程
'-------------------------------------------------------------------------------------------------
Public Sub LogInfo(ByRef strWarnInfo As String, ParamArray arrPars() As Variant)
'    Dim arrInfo()       As Variant
'    arrInfo = arrPars
'    mobjLog.WriteOperateArray String((mcllMethodStack.Count - 1) * 2, " ") & "┣" & strWarnInfo, arrInfo()
End Sub
'--------------------------------------------------------------------------------------------------
'方法           PushMethod
'功能           将调用方法推入堆栈
'返回值
'入参列表:
'参数名         类型                    说明
'strMethod      String                  方法名
'arrPars        String                  参数列表
'-------------------------------------------------------------------------------------------------
Public Sub PushMethod(ByRef strMethod As String, ParamArray arrPars() As Variant)
'    mstrText = ""
'    For mlngIndex = LBound(arrPars) To UBound(arrPars)
'        mstrText = mstrText & "," & DisPlayOneValue(arrPars(mlngIndex))
'    Next
'    mstrText = Mid(mstrText, 2)
'    With mcllMethodStack
'        If .Count = 0 Then
'            If mstrText = "" Then
'                .Add strMethod
'            Else
'                .Add strMethod & "(" & mstrText & ")"
'            End If
'        Else
'            If mstrText = "" Then
'                .Add strMethod, , 1
'            Else
'                .Add strMethod & "(" & mstrText & ")", , 1
'            End If
'        End If
'        If .Count > mlngStackLen Then .Remove .Count
'    End With
'    If mstrText = "" Then
'        mobjLog.WriteOperate String((mcllMethodStack.Count - 1) * 2, " ") & "┏" & strMethod
'    Else
'        mobjLog.WriteOperate String((mcllMethodStack.Count - 1) * 2, " ") & "┏" & strMethod & "(" & mstrText & ")"
'    End If
End Sub
'--------------------------------------------------------------------------------------------------
'方法           PopMethod
'功能           将最近的入栈的方法移除，或者将指定方法之前入堆栈的方法移除（包含指定的方法）
'返回值
'入参列表:
'参数名         类型                    说明
'strMethod      String                  方法名称，不传时弹出最近入堆栈的方法
'-------------------------------------------------------------------------------------------------
Public Sub PopMethod(ByRef strMethod As String, ParamArray arrPars() As Variant)
'    mstrText = ""
'    For mlngIndex = LBound(arrPars) To UBound(arrPars)
'        mstrText = mstrText & "," & DisPlayOneValue(arrPars(mlngIndex))
'    Next
'
'    If mstrText = "" Then
'        mobjLog.WriteOperate String((mcllMethodStack.Count - 1) * 2, " ") & "┗" & strMethod
'    Else
'        mstrText = Mid(mstrText, 2)
'        mobjLog.WriteOperate String((mcllMethodStack.Count - 1) * 2, " ") & "┗" & strMethod & "(" & mstrText & ")"
'    End If
'    With mcllMethodStack
'        If strMethod <> "" Then
'            For mlngIndex = 1 To .Count
'                If mcllMethodStack(mlngIndex) Like strMethod & "*" Then
'                    Exit For
'                End If
'            Next
'            If mlngIndex > .Count Then
'                If .Count > 0 Then  '没有找到任何匹配，则删除一个即可
'                    mlngIndex = 1
'                Else                '没有数据则不删除
'                    mlngIndex = 0
'                End If
'            End If
'        Else
'            mlngIndex = 1  '传空则只删除一个
'        End If
'
'        Do While mlngIndex > 0
'            .Remove 1
'            mlngIndex = mlngIndex - 1
'        Loop
'        mlngIndex = 1
'    End With
End Sub
