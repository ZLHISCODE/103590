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
'��д           lshuo
'����           2019/2/22
'ģ��           clsMemoryShareFP
'˵��           FileMapping��ʽ���ڴ湲��
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
Private mlngMapPtr      As Long             'ӳ����ڴ��ַ
Private mstrShareName   As String           '��������

Private Type MemHeader
    MaxSize     As Long                         '����������ߴ�
    Writed      As Byte                         '�������д��־
    Readed      As Byte                         '������Ķ���־
    Reserved    As Integer                      'Ϊ�����ֽڶ��룬������Ӹñ�����
    ProcessID   As Long                         '����ID
    Param       As Long                         '״̬�Ĳ���
    State       As Long                         '״̬��ֵ
    DataSize    As Long                         'ʹ�õ��ܳ���
End Type

Private mmhCur                                  As MemHeader
Private mstrData                                As String
Private mblnCreate                              As Boolean

'ShareName(R):��ǰ���������ȫ��Ψһ���ƣ�����ʹ��GUID
Public Property Get ShareName() As String
    ShareName = mstrShareName
End Property
'MaxSize(R):��ǰ������������ֵ
Public Property Get MaxSize() As Long
    MaxSize = mmhCur.MaxSize
End Property
'Writed(R):д��״̬��0-δд�롣1-�Ѿ�д��
Public Property Get Writed() As Long
    Writed = CLng(mmhCur.Writed)
End Property
'Readed(R):��ȡ״̬��0-δ��ȡ��>0-��ȡ����
Public Property Get Readed() As Long
    Readed = CLng(mmhCur.Readed)
End Property
'CurrentDataSize(R):��ǰ������洢���ַ�����С
Public Property Get CurrentDataSize() As Long
    CurrentDataSize = mmhCur.DataSize
End Property
'ProcessID(R):��ǰ���������д�����ID
Public Property Get ProcessID() As String
    ProcessID = mmhCur.ProcessID
End Property
'Param(R):��ǰ������Ľ��̽����Ĳ���
Public Property Get Param() As Long
    Param = mmhCur.Param
End Property
'State(R):��ǰ������Ľ��̽�����״ֵ̬
Public Property Get State() As Long
    State = mmhCur.State
End Property
'Data(R):��ǰ�������������
Public Property Get Data() As String
    Data = mstrData
End Property

'--------------------------------------------------------------------------------------------------
'����           CreateMemoryShare
'����           ���������ڴ�
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'strShareName   String                  ������������ʹ��GUID
'lngSize        Long                    64K-1,�����ڴ��С
'-------------------------------------------------------------------------------------------------
Public Function CreateMemoryShare(ByVal strShareName As String, Optional ByVal lngSize As Long = 65535) As Boolean
    Dim lngErr      As Long
    '�����ڴ湲��
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.CreateMemoryShare", "FILEMAPPING_" & strShareName, lngSize)
    Call CloseMemory
    mlngFile = CreateFileMapping(-1, ByVal 0&, PAGE_READWRITE, 0, lngSize, "FILEMAPPING_" & strShareName)
    If mlngFile = 0 Then
        lngErr = Err.LastDllError
        mlngFile = OpenFileMapping(FILE_MAP_READ Or FILE_MAP_WRITE, 0, "FILEMAPPING_" & strShareName)   '�������ڴ��ļ�ӳ�����
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
        '��������ͷ
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
'����           OpenMemoryShare
'����           ���Ѿ����ڵ��ڴ湲��
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'strShareName   String                  ������������ʹ��GUID
'-------------------------------------------------------------------------------------------------
Public Function OpenMemoryShare(ByVal strShareName As String) As Boolean
    Dim objMutex    As New clsMutex
    '���ڴ湲��
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.OpenMemoryShare", "FILEMAPPING_" & strShareName)
    Call CloseMemory
    mlngFile = OpenFileMapping(FILE_MAP_READ Or FILE_MAP_WRITE, 0, "FILEMAPPING_" & strShareName) '�������ڴ��ļ�ӳ�����
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
        '��������ͷ
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
'����           CloseMemory
'����           �ر��ڴ�ӳ�䡣�÷������ж���Ҫ����VB
'����ֵ         ��
'-------------------------------------------------------------------------------------------------
Private Sub CloseMemory()
    Dim mhCur       As MemHeader
    Dim objMutex    As New clsMutex
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.CloseMemory")
    If Not objMutex.CheckMutex(mstrShareName) Then 'û�л����壬���Զ��������ú�������ʱ�����Զ��ͷŻ�����
        '���״̬��ͷ������
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
'����           WriteMemory
'����           ���ڴ�ӳ��д�����ݡ���������ͬʱд�룬������һд��
'����ֵ         Long                    д�������ֽڳ���
'����б�:
'������         ����                    ˵��
'strInfo        String                  д�������
'lngProcessPID  Long                    д�빲���Ľ���ID������ǰ����ID
'lngParam       Long                    ����״̬������
'lngState       Long                    ����״ֵ̬
'-------------------------------------------------------------------------------------------------
Public Function WriteMemory(ByVal strInfo As String, Optional ByVal lngProcessPID As Long, Optional ByVal lngParam As Long, Optional ByVal lngState As Long) As Boolean
    Dim arrData()       As Byte
    Dim lngTotal        As Long
    Dim arrByteTmp()    As Byte
    Dim objMutex        As New clsMutex
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.WriteMemory", strInfo, lngProcessPID, lngParam, lngState)
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then 'û�л����壬���Զ��������ú�������ʱ�����Զ��ͷŻ�����
            If SubReadMemory(False, True) Then
                '��δ����ȡ
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
                    '��ͷ��������ʱ����
                    Call CopyMemory(arrData(0), mmhCur, LenB(mmhCur))
                    '�ַ���������ʱ�洢����
                    If lngTotal > 0 Then
                        Call CopyMemory(arrData(mmhCur.DataSize - lngTotal), arrByteTmp(0), lngTotal)
                    End If
                    '����ʱ������ڹ����ڴ�
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
'����           ReadMemory
'����           ���ڴ�ӳ���ļ��ڶ����ݡ��÷������ж���Ҫ����VB��
'����ֵ         String
'����б�:
'������         ����                    ˵��
'-------------------------------------------------------------------------------------------------
Public Function ReadMemory() As Boolean
    Dim objMutex        As New clsMutex
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.ReadMemory")
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then  'û�л����壬���Զ��������ú�������ʱ�����Զ��ͷŻ�����
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
'����           ReadMemoryOnce
'����           ���ڴ�ӳ���ļ��ڶ����ݡ�����ʱ�ɹ����Զ��ͷŸ��ڴ����򡣸÷������ж���Ҫ����VB��
'����ֵ         String
'����б�:
'������         ����                    ˵��
'-------------------------------------------------------------------------------------------------
Public Function ReadMemoryOnce() As Boolean
    Dim objMutex        As New clsMutex
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.ReadMemoryOnce")
    If mlngMapPtr <> 0 Then
        If Not objMutex.CheckMutex(mstrShareName) Then  'û�л����壬���Զ��������ú�������ʱ�����Զ��ͷŻ�����
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
'����           ReleaseMemory
'����           �ͷ��ڴ档�÷������ж���Ҫ����VB
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
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
'����           ReleaseMemory
'����           �ͷ��ڴ档
'����ֵ         Boolean
'����б�:
'������         ����                    ˵��
'-------------------------------------------------------------------------------------------------
Private Function SubReleaseMemory() As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.SubReleaseMemory")
    If mlngMapPtr <> 0 Then
        '����պ��״̬����
        mmhCur.Writed = 0
        mmhCur.Readed = 0
        mmhCur.DataSize = LenB(mmhCur)
        '����յ�״̬�����ڴ湲��
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
'����           SubReadMemory
'����           ���ڴ�ӳ���ļ��ڶ����ݡ�SubReadMemory��SubReadMemoryOnce���Ӻ���
'����ֵ         String
'����б�:
'������         ����                    ˵��
'blnOnceCall    Boolean                 �Ƿ�ֻ����һ�Σ����ǣ�����ʱ�ɹ����Զ��ͷŸ��ڴ�����
'blnOnlyHeader  Boolean                 �Ƿ�ֻ��ȡͷ��
'-------------------------------------------------------------------------------------------------
Private Function SubReadMemory(Optional ByVal blnOnceCall As Boolean, Optional ByVal blnOnlyHeader As Boolean) As Boolean
    Dim arrData()       As Byte
    Dim lngTotal        As Long
    Dim arrByteTmp()    As Byte
    Dim objMutex        As New clsMutex
    Dim mhCur           As MemHeader
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsMemoryShareFP.SubReadMemory", blnOnceCall, blnOnlyHeader)
    '��ȡ�����ڴ�ͷ��������ͷ���Ϸ��Խ��м��
    Call CopyMemory(mhCur, ByVal mlngMapPtr, LenB(mhCur))
    If mhCur.MaxSize <> mmhCur.MaxSize Or mhCur.DataSize < 0 Or mhCur.DataSize > mhCur.MaxSize Then
        Logger.DebugEx "ͷ����Ϣ", "ST_MaxSize", mmhCur.MaxSize, "MaxSize", mhCur.MaxSize, "DataSize", mhCur.DataSize, "Writed", mhCur.Writed, "Readed", mhCur.Readed, "ProcessID", mhCur.ProcessID, "Param", mhCur.Param, "State", mhCur.State
    Else
        Call CopyMemory(mmhCur, mhCur, LenB(mhCur))
        If blnOnlyHeader Then
            SubReadMemory = True
            Logger.DebugEx "ͷ����Ϣ", "MaxSize", mmhCur.MaxSize, "DataSize", mmhCur.DataSize, "Writed", mmhCur.Writed, "Readed", mmhCur.Readed, "ProcessID", mmhCur.ProcessID, "Param", mmhCur.Param, "State", mmhCur.State
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
                Logger.DebugEx "����", "Data", mstrData, "MaxSize", mmhCur.MaxSize, "DataSize", mmhCur.DataSize, "Writed", mmhCur.Writed, "Readed", mmhCur.Readed, "ProcessID", mmhCur.ProcessID, "Param", mmhCur.Param, "State", mmhCur.State
                mmhCur.Readed = (mmhCur.Readed Mod 255) + 1
                Call CopyMemory(ByVal mlngMapPtr, mmhCur, LenB(mmhCur))         '��д״̬
                If blnOnceCall Then Call SubReleaseMemory
                SubReadMemory = True
            Else
                Logger.DebugEx "ͷ����Ϣ", "MaxSize", mmhCur.MaxSize, "DataSize", mmhCur.DataSize, "Writed", mmhCur.Writed, "Readed", mmhCur.Readed, "ProcessID", mmhCur.ProcessID, "Param", mmhCur.Param, "State", mmhCur.State
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