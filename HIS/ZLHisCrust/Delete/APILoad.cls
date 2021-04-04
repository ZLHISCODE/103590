VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APILoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
'数据类型定义
'==============================================================================
Private Type VariableBuffer
   VariableParameter() As Byte
End Type
'==============================================================================
'API 函数声明
'==============================================================================
'==============================================================================
'成员定义
'==============================================================================
'类中的全局变量
Private m_opIndex As Long
Private m_OpCode() As Byte
Private mcllModule As New Collection '动态库句柄，索引为动态库地址，值为库的句柄
Private mcllFuncs As New Collection '动态库句柄，索引为动态库地址，值为库的句柄
'********************************************************************************
'**   函 数 名 ：    ExecuteAPI
'**   输    入 ：    LIBPath(String)         -  刷新的目标窗口句柄，可为0
'**            ：    APIScript(String)       -  场景图像的宽度
'**   返    回 ：    (Long)                  -  返回零表示失败，非零表示成功
'**   功能描述 ：    动态执行类库里的API函数
'**   创建日期 ：    2013-09-17
'**   修 改 人 ：
'**   修改日期 ：
'********************************************************************************
Public Function ExecuteAPI(ByVal strLibPath As String, ByVal strFuncName As String, ParamArray arrParas() As Variant) As Long
    Dim hProcAddress As Long, i As Long
    Dim LongCount As Long, StrByteArray() As VariableBuffer
    Dim StringSize As Long, ByteArray() As Byte, IsHaveParameter As Boolean
    Dim OutputArray() As Long
    
    If Not GetAPIFunction(strLibPath, strFuncName, hProcAddress) Then
        ExecuteAPI = 0   'Library 或函数读取失败
        Exit Function
    End If
    
    StringSize = 0
    ReDim StrByteArray(StringSize)
   
    '是否带参数
    IsHaveParameter = UBound(arrParas) <> -1
    If IsHaveParameter Then
        '带参数的函数
        IsHaveParameter = True
        '初始化函数内存大小
        ReDim OutputArray(UBound(arrParas))
        '格式化函数参数
        Dim strNumber As String
        For i = 0 To UBound(arrParas)
            If TypeName(arrParas(i)) = "String" Then
                If Len(arrParas(i)) = 0 Then
                    OutputArray(i) = CLng(VarPtr(Null))
                Else
                    ReDim Preserve StrByteArray(StringSize)
                    ByteArray = StrConv(arrParas(i), vbFromUnicode)
                    ReDim Preserve StrByteArray(StringSize).VariableParameter(UBound(ByteArray) + 1)
                    CopyMemory StrByteArray(StringSize).VariableParameter(0), ByteArray(0), UBound(ByteArray) + 1
                    OutputArray(i) = CLng(VarPtr(StrByteArray(StringSize).VariableParameter(0)))
                    StringSize = StringSize + 1
                End If
            Else
                LongCount = CLng(arrParas(i))
                OutputArray(i) = LongCount
            End If
        Next
        ReDim m_OpCode(400 + 6 * UBound(OutputArray)) '保留用来写m_OpCode
    End If
    If IsHaveParameter = True Then
       '带参数的情况在此执行
       ExecuteAPI = CallWindowProc(GetCodeStart(hProcAddress, OutputArray), 0, 1, 2, 3)
    Else
       '不带参数的情况在此执行
       ExecuteAPI = CallWindowProc(hProcAddress, 0, 1, 2, 3)
    End If
End Function

Private Function GetAPIFunction(ByVal strLibPath As String, ByVal strFuncName As String, ByRef lngProc As Long) As Boolean
'功能：获取函数地址，并缓存
    Dim strFunInfo As String, arrTmp As Variant
    Dim lngModule As Long, blnNewFun As Boolean, blnNewModule As Boolean
    
    '判断API库是否缓存
    On Error Resume Next
    strFunInfo = mcllFuncs("K_" & strLibPath & "_" & strFuncName)
    If Err.Number <> 0 Then
        Err.Clear
        blnNewFun = True
        strFunInfo = mcllModule("K_" & strLibPath)
        If Err.Number <> 0 Then Err.Clear: blnNewModule = True
    Else
        arrTmp = Split(strFunInfo, ",")
        lngProc = CLng(arrTmp(1))
        lngModule = CLng(arrTmp(0))
        strFunInfo = mcllModule("K_" & strLibPath)
        If Err.Number <> 0 Then
            Err.Clear: blnNewModule = True: blnNewFun = True
            Call mcllFuncs.Remove("K_" & strLibPath & "_" & strFuncName)
        ElseIf lngModule <> Val(strFunInfo) Then
            blnNewFun = True
            Call mcllFuncs.Remove("K_" & strLibPath & "_" & strFuncName)
        End If
    End If
    If blnNewModule Or blnNewFun Then
        If blnNewModule Then
            '读取API库
            lngModule = LoadLibrary(ByVal strLibPath)
            If lngModule = 0 Then Exit Function
            mcllModule.Add lngModule, "K_" & strLibPath
        End If
        If blnNewFun Then
            '取得函数地址
            lngProc = GetProcAddress(lngModule, ByVal strFuncName)
            If lngProc = 0 Then
               Exit Function
            End If
            mcllFuncs.Add lngModule & "," & lngProc, "K_" & strLibPath & "_" & strFuncName
        End If
    End If
    GetAPIFunction = True
End Function

Public Function FreeAPI(Optional ByVal strLibPath As String) As Boolean
'功能：手工释放API资源
    Dim arrItem As Variant
    If strLibPath = "" Then
        For Each arrItem In mcllModule
            '释放库空间
            FreeLibrary CLng(arrItem)
        Next
        Set mcllModule = New Collection
    Else
        On Error Resume Next
        arrItem = mcllModule(strLibPath)
        If Err.Number <> 0 Then Err.Clear: FreeAPI = True: Exit Function
        FreeLibrary CLng(arrItem)
        Call mcllModule.Remove(strLibPath)
    End If
End Function


Private Function GetCodeStart(ByVal lngProc As Long, arrParams() As Long) As Long
    Dim lngIndex As Long, lngCodeStart As Long
    lngCodeStart = (VarPtr(m_OpCode(0)) Or &HF) + 1
    m_opIndex = lngCodeStart - VarPtr(m_OpCode(0))
    For lngIndex = 0 To m_opIndex - 1
        m_OpCode(lngIndex) = &HCC
    Next lngIndex
    For lngIndex = UBound(arrParams) To 0 Step -1
       AddByteToCode &H68
       AddLongToCode arrParams(lngIndex)
    Next lngIndex
    AddByteToCode &HE8
    AddLongToCode lngProc - VarPtr(m_OpCode(m_opIndex)) - 4
    AddByteToCode &HC2
    AddByteToCode &H10
    AddByteToCode &H0
    GetCodeStart = lngCodeStart
End Function

Private Sub AddLongToCode(lData As Long)
    CopyMemory m_OpCode(m_opIndex), lData, 4
    m_opIndex = m_opIndex + 4
End Sub

Private Sub AddIntToCode(iData As Integer)
    CopyMemory m_OpCode(m_opIndex), iData, 2
    m_opIndex = m_opIndex + 2
End Sub

Private Sub AddByteToCode(bData As Byte)
    m_OpCode(m_opIndex) = bData
    m_opIndex = m_opIndex + 1
End Sub

Private Sub Class_Terminate()
    Call FreeAPI
End Sub
