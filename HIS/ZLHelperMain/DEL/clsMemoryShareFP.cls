VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemoryShareFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/2/22
'模块           clsMemoryShareFP
'说明           FileMapping方式的内存共享
'==================================================================================================
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function OpenFileMapping Lib "kernel32.dll" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SECTION_MAP_EXECUTE As Long = &H8
Private Const SECTION_MAP_READ As Long = &H4
Private Const SECTION_MAP_WRITE As Long = &H2
Private Const SECTION_QUERY As Long = &H1
Private Const SECTION_EXTEND_SIZE As Long = &H10
Private Const SECTION_ALL_ACCESS As Long = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Private Const FILE_MAP_ALL_ACCESS As Long = SECTION_ALL_ACCESS
Private Const FILE_MAP_COPY As Long = SECTION_QUERY
Private Const FILE_MAP_READ As Long = SECTION_MAP_READ
Private Const FILE_MAP_WRITE As Long = SECTION_MAP_WRITE
Private Const PAGE_READWRITE As Long = &H4
Private Const ERROR_ALREADY_EXISTS = 183&

Private mlngFile        As Long
Private mlngMapPtr      As Long             '映射的内存地址
Private mstrShareName   As String           '共享名称

Private Type MemHeader
    MaxSize     As Long                         '该区域的最大尺寸
    Writed      As Byte                         '该区域的写标志
    Readed      As Byte                         '该区域的读标志
    Reserved    As Integer                      '为了四字节对齐，因此增加该保留区
    ProcessID   As Long                         '进程ID
    Param       As Long                         '状态的参数
    State       As Long                         '状态的值
    DataSize    As Long                         '使用的总长度
End Type

Private mmhCur                                  As MemHeader
Private mstrData                                As String
Private mblnCreate                              As Boolean

'ShareName(R):当前共享区域的全局唯一名称，建议使用GUID
Public Property Get ShareName() As String
    ShareName = mstrShareName
End Property
'MaxSize(R):当前共享区域的最大值
Public Property Get MaxSize() As Long
    MaxSize = mmhCur.MaxSize
End Property
'Writed(R):写入状态。0-未写入。1-已经写入
Public Property Get Writed() As Long
    Writed = CLng(mmhCur.Writed)
End Property
'Readed(R):读取状态。0-未读取。>0-读取次数
Public Property Get Readed() As Long
    Readed = CLng(mmhCur.Readed)
End Property
'CurrentDataSize(R):当前共享域存储的字符串大小
Public Property Get CurrentDataSize() As Long
    CurrentDataSize = mmhCur.DataSize
End Property
'ProcessID(R):当前共享区域的写入进程ID
Public Property Get ProcessID() As String
    ProcessID = mmhCur.ProcessID
End Property
'Param(R):当前共享域的进程交互的参数
Public Property Get Param() As Long
    Param = mmhCur.Param
End Property
'State(R):当前共享域的进程交互的状态值
Public Property Get State() As Long
    State = mmhCur.State
End Property
'Data(R):当前共享区域的内容
Public Property Get Data() As String
    Data = mstrData
End Property

'--------------------------------------------------------------------------------------------------
'方法           CreateMemoryShare
'功能           建立共享内存
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'strShareName   String                  共享名，建议使用GUID
'lngSize        Long                    64K-1,共享内存大小
'-------------------------------------------------------------------------------------------------
Public Function CreateMemoryShare(ByVal strShareName As String, Optional ByVal lngSize As Long = 65535) As Boolean
    Dim lngErr      As Long
    '建立内存共享
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.CreateMemoryShare", "FILEMAPPING_" & strShareName, lngSize)
    Call CloseMemory
    mlngFile = CreateFileMapping(-1, ByVal 0&, PAGE_READWRITE, 0, lngSize, "FILEMAPPING_" & strShareName)
    If mlngFile = 0 Then
        lngErr = Err.LastDllError
        mlngFile = OpenFileMapping(FILE_MAP_READ Or FILE_MAP_WRITE, 0, "FILEMAPPING_" & strShareName)   '打开命名内存文件映射对象
    End If
    
    If mlngFile = 0 Then
        Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.CreateMemoryShare", "mlngFile = 0", CreateMemoryShare)
        Call CloseMemory
        Exit Function
    End If
    mlngMapPtr = MapViewOfFile(mlngFile, FILE_MAP_READ Or FILE_MAP_WRITE, 0, 0, 0)
    If mlngMapPtr = 0 Then
        Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.CreateMemoryShare", " mlngMapPtr = 0", CreateMemoryShare)
        Call CloseMemory
        Exit Function
    Else
        mmhCur.MaxSize = lngSize
        mmhCur.DataSize = Len(mmhCur)
        '缓存区域头
        CopyMemory ByVal mlngMapPtr, mmhCur, LenB(mmhCur)
        mblnCreate = True
    End If
    mstrShareName = strShareName
    CreateMemoryShare = True
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.CreateMemoryShare", CreateMemoryShare)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.CreateMemoryShare") = 1 Then
        Resume
    End If
    Call CloseMemory
End Function

'--------------------------------------------------------------------------------------------------
'方法           OpenMemoryShare
'功能           打开已经存在的内存共享
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'strShareName   String                  共享名，建议使用GUID
'-------------------------------------------------------------------------------------------------
Public Function OpenMemoryShare(ByVal strShareName As String) As Boolean
    Dim objMutex    As New clsMutex
    '打开内存共享
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.OpenMemoryShare", "FILEMAPPING_" & strShareName)
    Call CloseMemory
    mlngFile = OpenFileMapping(FILE_MAP_READ Or FILE_MAP_WRITE, 0, "FILEMAPPING_" & strShareName) '打开命名内存文件映射对象
    If mlngFile = 0 Then
        Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.OpenMemoryShare", "mlngFile = 0", OpenMemoryShare)
        Exit Function
    End If
    mlngMapPtr = MapViewOfFile(mlngFile, FILE_MAP_READ Or FILE_MAP_WRITE, 0, 0, 0)
    If mlngMapPtr = 0 Then
        Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.OpenMemoryShare", "mlngMapPtr = 0", OpenMemoryShare)
        Call CloseMemory
        Exit Function
    Else
        '缓存区域头
        CopyMemory mmhCur, ByVal mlngMapPtr, LenB(mmhCur)
    End If
    mstrShareName = strShareName
    OpenMemoryShare = True
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.OpenMemoryShare", OpenMemoryShare)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.OpenMemoryShare") = 1 Then
        Resume
    End If
    Call CloseMemory
End Function
'--------------------------------------------------------------------------------------------------
'方法           CloseMemory
'功能           关闭内存映射。该方法中中断需要重启VB
'返回值         无
'-------------------------------------------------------------------------------------------------
Private Sub CloseMemory()
    Dim mhCur       As MemHeader
    Dim objMutex    As New clsMutex
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.CloseMemory")
    If Not objMutex.CheckMutex(mstrShareName) Then '没有互斥体，则自动创建，该函数结束时，会自动释放互斥体
        '清空状态与头部缓存
        If mblnCreate Then
            If mlngMapPtr <> 0 Or mlngFile <> 0 Then
                CopyMemory mmhCur, mhCur, LenB(mhCur)
            End If
            mstrData = ""
        End If
        If mlngMapPtr <> 0 Then
            UnmapViewOfFile mlngMapPtr
            mlngMapPtr = 0
        End If
        If mlngFile <> 0 Then
            CloseHandle mlngFile
            mlngFile = 0
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.CloseMemory")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.CloseMemory") = 1 Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------------------------------------
'方法           WriteMemory
'功能           向内存映射写入数据。所有数据同时写入，不能逐一写入
'返回值         Long                    写入后的总字节长度
'入参列表:
'参数名         类型                    说明
'strInfo        String                  写入的内容
'lngProcessPID  Long                    写入共享的进程ID。即当前进程ID
'lngParam       Long                    共享状态参数。
'lngState       Long                    共享状态值
'-------------------------------------------------------------------------------------------------
Public Function WriteMemory(ByVal strInfo As String, Optional ByVal lngProcessPID As Long, Optional ByVal lngParam As Long, Optional ByVal lngState As Long) As Boolean
    Dim arrData()       As Byte
    Dim lngTotal        As Long
    Dim arrByteTmp()    As Byte
    Dim objMutex        As New clsMutex
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.WriteMemory", strInfo, lngProcessPID, lngParam, lngState)
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then '没有互斥体，则自动创建，该函数结束时，会自动释放互斥体
            If SubReadMemory(False, True) Then
                '尚未被读取
                If Not (mmhCur.Writed = 1 And mmhCur.Readed = 0) Then
                    mmhCur.Writed = 1
                    mmhCur.Readed = 0
                    mmhCur.ProcessID = lngProcessPID
                    mmhCur.Param = lngParam
                    mmhCur.State = lngState
                    If strInfo <> "" Then
                        arrByteTmp = StrConv(strInfo, vbFromUnicode)
                        lngTotal = UBound(arrByteTmp) + 1
                    End If
                    mmhCur.DataSize = lngTotal + LenB(mmhCur)
                    ReDim Preserve arrData(mmhCur.DataSize - 1)
                    '将头部放在临时区域
                    Call CopyMemory(arrData(0), mmhCur, LenB(mmhCur))
                    '字符串放在临时存储区域
                    If lngTotal > 0 Then
                        Call CopyMemory(arrData(mmhCur.DataSize - lngTotal), arrByteTmp(0), lngTotal)
                    End If
                    '将临时缓存放在共享内存
                    Call CopyMemory(ByVal mlngMapPtr, arrData(0), mmhCur.DataSize)
                    WriteMemory = True
                End If
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.WriteMemory", WriteMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.WriteMemory") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           ReadMemory
'功能           从内存映射文件内读数据。该方法中中断需要重启VB。
'返回值         String
'入参列表:
'参数名         类型                    说明
'-------------------------------------------------------------------------------------------------
Public Function ReadMemory() As Boolean
    Dim objMutex        As New clsMutex
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.ReadMemory")
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then  '没有互斥体，则自动创建，该函数结束时，会自动释放互斥体
            ReadMemory = SubReadMemory()
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.ReadMemory", ReadMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.ReadMemory") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           ReadMemoryOnce
'功能           从内存映射文件内读数据。调用时成功后自动释放该内存区域。该方法中中断需要重启VB。
'返回值         String
'入参列表:
'参数名         类型                    说明
'-------------------------------------------------------------------------------------------------
Public Function ReadMemoryOnce() As Boolean
    Dim objMutex        As New clsMutex
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.ReadMemoryOnce")
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then  '没有互斥体，则自动创建，该函数结束时，会自动释放互斥体
            ReadMemoryOnce = SubReadMemory(True)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.ReadMemoryOnce", ReadMemoryOnce)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.ReadMemoryOnce") = 1 Then
        Resume
    End If
End Function
'--------------------------------------------------------------------------------------------------
'方法           ReleaseMemory
'功能           释放内存。该方法中中断需要重启VB
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'-------------------------------------------------------------------------------------------------
Public Function ReleaseMemory() As Boolean
    Dim objMutex        As New clsMutex
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.ReleaseMemory")
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then
            ReleaseMemory = SubReleaseMemory
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.ReleaseMemory", ReleaseMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.ReleaseMemory") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           ReleaseMemory
'功能           释放内存。
'返回值         Boolean
'入参列表:
'参数名         类型                    说明
'-------------------------------------------------------------------------------------------------
Private Function SubReleaseMemory() As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.SubReleaseMemory")
    If mlngMapPtr <> 0 Then
        '将清空后的状态缓存
        mmhCur.Writed = 0
        mmhCur.Readed = 0
        mmhCur.DataSize = LenB(mmhCur)
        '将清空的状态放入内存共享
        Call CopyMemory(ByVal mlngMapPtr, mmhCur, LenB(mmhCur))
    End If
    SubReleaseMemory = True
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.SubReleaseMemory", SubReleaseMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.SubReleaseMemory") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           SubReadMemory
'功能           从内存映射文件内读数据。SubReadMemory与SubReadMemoryOnce的子函数
'返回值         String
'入参列表:
'参数名         类型                    说明
'blnOnceCall    Boolean                 是否只调用一次，若是，调用时成功后自动释放该内存区域
'blnOnlyHeader  Boolean                 是否只读取头部
'-------------------------------------------------------------------------------------------------
Private Function SubReadMemory(Optional ByVal blnOnceCall As Boolean, Optional ByVal blnOnlyHeader As Boolean) As Boolean
    Dim arrData()       As Byte
    Dim lngTotal        As Long
    Dim arrByteTmp()    As Byte
    Dim objMutex        As New clsMutex
    Dim mhCur           As MemHeader
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.SubReadMemory", blnOnceCall, blnOnlyHeader)
    '读取共享内存头部，并对头部合法性进行检查
    Call CopyMemory(mhCur, ByVal mlngMapPtr, LenB(mhCur))
    If mhCur.MaxSize <> mmhCur.MaxSize Or mhCur.DataSize < 0 Or mhCur.DataSize > mhCur.MaxSize Then
        Logger.DebugEx "头部信息", "ST_MaxSize", mmhCur.MaxSize, "MaxSize", mhCur.MaxSize, "DataSize", mhCur.DataSize, "Writed", mhCur.Writed, "Readed", mhCur.Readed, "ProcessID", mhCur.ProcessID, "Param", mhCur.Param, "State", mhCur.State
    Else
        Call CopyMemory(mmhCur, mhCur, LenB(mhCur))
        If blnOnlyHeader Then
            SubReadMemory = True
            Logger.DebugEx "头部信息", "MaxSize", mmhCur.MaxSize, "DataSize", mmhCur.DataSize, "Writed", mmhCur.Writed, "Readed", mmhCur.Readed, "ProcessID", mmhCur.ProcessID, "Param", mmhCur.Param, "State", mmhCur.State
        Else
            ReDim Preserve arrData(mmhCur.DataSize - 1)
            lngTotal = mmhCur.DataSize - LenB(mmhCur)
            Call CopyMemory(arrData(0), ByVal mlngMapPtr, mmhCur.DataSize)
            If mmhCur.Writed <> 0 Then
                If lngTotal > 0 Then
                    ReDim Preserve arrByteTmp(lngTotal - 1)
                    Call CopyMemory(arrByteTmp(0), arrData(mmhCur.DataSize - lngTotal), lngTotal)
                    mstrData = StrConv(arrByteTmp, vbUnicode)
                End If
                Logger.DebugEx "数据", "Data", mstrData, "MaxSize", mmhCur.MaxSize, "DataSize", mmhCur.DataSize, "Writed", mmhCur.Writed, "Readed", mmhCur.Readed, "ProcessID", mmhCur.ProcessID, "Param", mmhCur.Param, "State", mmhCur.State
                mmhCur.Readed = (mmhCur.Readed Mod 255) + 1
                Call CopyMemory(ByVal mlngMapPtr, mmhCur, LenB(mmhCur))         '回写状态
                If blnOnceCall Then Call SubReleaseMemory
                SubReadMemory = True
            Else
                Logger.DebugEx "头部信息", "MaxSize", mmhCur.MaxSize, "DataSize", mmhCur.DataSize, "Writed", mmhCur.Writed, "Readed", mmhCur.Readed, "ProcessID", mmhCur.ProcessID, "Param", mmhCur.Param, "State", mmhCur.State
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.SubReadMemory", SubReadMemory)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.SubReadMemory") = 1 Then
        Resume
    End If
End Function
Private Sub Class_Initialize()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.Class_Initialize")
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.Class_Initialize")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.Class_Initialize") = 1 Then
        Resume
    End If
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.Class_Terminate")
    Call CloseMemory
    Call Logger.PopMethod("ZLHelperMain.clsMemoryShareFP.Class_Terminate")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsMemoryShareFP.Class_Terminate") = 1 Then
        Resume
    End If
End Sub
