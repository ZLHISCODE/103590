VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImapi2Burn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


  Private Type TGUID
    D1 As Long
    D2 As Long
    D3 As Long
    D4(7) As Byte
  End Type
  
  Private Type tagSTAT
    pwcsName As String
    type As Long
    cbSize As Double
    mtime As Date
    ctime As Date
    atime As Date
    grfMode As Long
    grfLocksSupported As Long
    clsid As TGUID
    grfStateBits As Long
    reserved As Long
  End Type
  

Private Const C_STR_BURN_CLIENTNAME As String = "ZLSoftBurn"
Private Const C_STR_DISC_DIR_SPLIT  As String = ">"


Private WithEvents mobjDiscs       As MsftDiscMaster2
Attribute mobjDiscs.VB_VarHelpID = -1
Private WithEvents mobjFileImage   As MsftFileSystemImage
Attribute mobjFileImage.VB_VarHelpID = -1
Private WithEvents mobjDiscFormat  As MsftDiscFormat2Data
Attribute mobjDiscFormat.VB_VarHelpID = -1

Private mobjCurRecord   As MsftDiscRecorder2
Private mobjDiscEarse   As MsftDiscFormat2Erase '光盘擦除对象

Private mstrBurnDeviceName  As String   '当前需要使用的刻录设备名称
Private mstrBurnVolumeName As String    '盘符名称
Private mlngWriteSpeed As Long          '刻录速度
Private mblnIsOverWrite As Boolean      '是否覆盖
Private mlngVerificationLevel As TIntergrityVerificationLevel   '完整性检测级别
Private mblnOnceMedia As Boolean        '是否封闭光盘，不允许再次写入
Private mblnIncludeBaseDir As Boolean   '是否包含基本目录
Private mblnIsCancelWrite As Boolean    '是否撤销刻录操作
Private mlngReserveSize As Long         '光盘预留大小（单位K）

Private mobjBurnDirResources As Collection  '保持待刻录的目录
Private mobjBurnFileResources As Collection '保存待刻录的文件

Private mstrBurnPath As String

Public Event OnBurnEvent(ByVal strCurState As String, ByRef args As clsImapi2BurnArgs)

Public Event OnReadyFileEvent(ByVal strFile As String, ByVal strSourcePath As String, _
    ByVal copiedSectors As Long, ByVal totalSectors As Long)

Public Event OnBurnProcedureEvent(ByVal strState As String)





'***********************************************************************************************************
'
'封装使用imapi2进行刻录的功能
'
'***********************************************************************************************************




'判断该系统中是否存在可用的刻录设备
Property Get HasBurnDeviceInSystem() As Boolean
    HasBurnDeviceInSystem = mobjDiscs.IsSupportedEnvironment
End Property


'获取光驱数量
Property Get DeviceCount() As Long
    DeviceCount = mobjDiscs.Count
End Property


'获取刻录设备的名称
Private Function GetBurnDeviceName(objRecord As MsftDiscRecorder2) As String
    GetBurnDeviceName = TransformArrayToOneDimension(objRecord.VolumePathNames)(1) & objRecord.ProductId
End Function


'返回设备名称(卷标+设备名)
Property Get DeviceName(ByVal lngDeviceIndex As Long) As String
On Error Resume Next
    Dim i As Long
    Dim objRecord As New MsftDiscRecorder2
       
    Call objRecord.InitializeDiscRecorder(mobjDiscs.Item(lngDeviceIndex))
    
    DeviceName = GetBurnDeviceName(objRecord)
    
    Set objRecord = Nothing

End Property



'当前可使用的刻录设备名称
Property Get CurBurnDevice() As String
    CurBurnDevice = mstrBurnDeviceName
End Property


Property Let CurBurnDevice(ByVal value As String)
    mstrBurnDeviceName = value

    Set mobjCurRecord = Nothing
    Set mobjCurRecord = GetDiscDevice(value)
    
    If mobjCurRecord Is Nothing Then Exit Property
    
    If Not mobjDiscFormat.IsRecorderSupported(mobjCurRecord) Then
        '不支持刻录操作
        Set mobjCurRecord = Nothing
        Exit Property
    End If
    
    If Not mobjDiscFormat.IsCurrentMediaSupported(mobjCurRecord) Then
        '判断光盘是否准备就绪
        Set mobjCurRecord = Nothing
        Exit Property
    End If
    
    mobjDiscFormat.Recorder = mobjCurRecord
'    mobjDiscFormat.ClientName = C_STR_BURN_CLIENTNAME

End Property


'刻录使用的卷标名称
Property Get BurnVolumeName() As String
    BurnVolumeName = mstrBurnVolumeName
End Property

Property Let BurnVolumeName(value As String)
    If Trim(value) = "" Then Exit Property
    
    mstrBurnVolumeName = value
End Property


'设置写入速度
Property Get WriteSpeed() As Long
    WriteSpeed = mlngWriteSpeed
End Property


Property Let WriteSpeed(value As Long)
    mlngWriteSpeed = value
End Property


'是否对源文件进行覆盖
Property Get IsOverWrite() As Boolean
    IsOverWrite = mblnIsOverWrite
End Property

Property Let IsOverWirte(value As Boolean)
    mblnIsOverWrite = value
End Property


'是否包含基本刻录目录
Property Get IsIncludeBaseDir() As Boolean
    IsIncludeBaseDir = mblnIncludeBaseDir
End Property

Property Let IsIncludeBaseDir(value As Boolean)
    mblnIncludeBaseDir = value
End Property


'完整性检测级别
Property Get VerificationLevel() As TIntergrityVerificationLevel
    VerificationLevel = mlngVerificationLevel
End Property

Property Let VerificationLevel(value As TIntergrityVerificationLevel)
    mlngVerificationLevel = value
End Property


'保留大小(单位为K)
Property Get ReserveKBSize() As Long
    ReserveKBSize = mlngReserveSize
End Property


Property Let ReserveKBSize(value As Long)
    mlngReserveSize = value
End Property


'是否为一次性媒体
Property Get OnceMedia() As Boolean
    OnceMedia = mblnOnceMedia
End Property

Property Let OnceMedia(value As Boolean)
    mblnOnceMedia = value
End Property



'**************************************************************************************************************************









'获取当前刻录设备
Public Function GetCurDiscDevice() As MsftDiscRecorder2
    Set GetCurDiscDevice = mobjCurRecord
End Function


'根据设备名取得指定的刻录设备
Public Function GetDiscDevice(ByVal strDeviceName As String) As MsftDiscRecorder2
    Dim i As Integer
    Dim objRecord As MsftDiscRecorder2
    
    Set GetDiscDevice = Nothing
    For i = 0 To mobjDiscs.Count - 1
        Set objRecord = New MsftDiscRecorder2
        Call objRecord.InitializeDiscRecorder(mobjDiscs.Item(i))
        
        If UCase(GetBurnDeviceName(objRecord)) = UCase(strDeviceName) Then
            Set GetDiscDevice = objRecord
            Exit Function
        Else
            Set objRecord = Nothing
        End If
    Next i
End Function


'检查设备是否支持刻录
Public Function CheckingDeviceIsBurn(ByVal strDeviceName As String) As Boolean
    Dim objRecord As MsftDiscRecorder2
    
    Set objRecord = GetDiscDevice(strDeviceName)

    '判断是否支持刻录
    CheckingDeviceIsBurn = mobjDiscFormat.IsRecorderSupported(objRecord)
End Function


'检查光盘是否准备就绪
Public Function CheckingDeviceIsReady(ByVal strDeviceName As String) As Boolean
    Dim objRecord As MsftDiscRecorder2
        
    Set objRecord = GetDiscDevice(strDeviceName)
      
    CheckingDeviceIsReady = mobjDiscFormat.IsCurrentMediaSupported(objRecord)
    
    Set objRecord = Nothing
End Function


'获取磁盘名称
Public Function GetDiscName(ByVal strDeviceName As String)
    Dim strVolumnName  As String
    Dim FileSystemName As String
    Dim lngLength As Long
    Dim SerialNumber As Long

    
    strVolumnName = String(255, Chr(0))
    FileSystemName = String(255, Chr(0))


    Call GetVolumeInformation(Mid(strDeviceName, 1, 3), strVolumnName, 255, SerialNumber, 0, 0, FileSystemName, 255)
    
    GetDiscName = Replace(Trim(strVolumnName), Chr(0), "")
End Function


'获取指定媒体类型别名
Private Function GetDiscMediaTypeAlias(ByVal lngMediaType As Long) As String
    GetDiscMediaTypeAlias = ""
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "GetCurDiscMediaTypeAlias", "当前设备无效，不能执行该操作。")
        Exit Function
    End If
    
    Select Case lngMediaType
        Case 1
            GetDiscMediaTypeAlias = "CD-ROM"
        Case 2
            GetDiscMediaTypeAlias = "CD-R"
        Case 3
            GetDiscMediaTypeAlias = "CD-RW"
        Case 4
            GetDiscMediaTypeAlias = "DVD-ROM"
        Case 5
            GetDiscMediaTypeAlias = "DVD-RAM"
        Case 6
            GetDiscMediaTypeAlias = "DVD+R"
        Case 7
            GetDiscMediaTypeAlias = "DVD+RW"
        Case 8
            GetDiscMediaTypeAlias = "DVD+R DUAL LAYER"
        Case 9
            GetDiscMediaTypeAlias = "DVD-R"
        Case 10
            GetDiscMediaTypeAlias = "DVD-RW"
        Case 11
            GetDiscMediaTypeAlias = "DVD-R DUAL LAYER"
        Case 12
            GetDiscMediaTypeAlias = "RANDOMLY WRITABLE"
        Case 13
            GetDiscMediaTypeAlias = "DVD+RW DUAL LAYER"
        Case 14
            GetDiscMediaTypeAlias = "HD DVD-ROM"
        Case 15
            GetDiscMediaTypeAlias = "HD DVD-R"
        Case 16
            GetDiscMediaTypeAlias = "HD DVD-RAM"
        Case 17
            GetDiscMediaTypeAlias = "BD-ROM"
        Case 18
            GetDiscMediaTypeAlias = "BD-R"
        Case 19
            GetDiscMediaTypeAlias = "MAX"
        Case Else
            GetDiscMediaTypeAlias = "未知"
    End Select
End Function


'取得当前媒体类型别名
Public Function GetCurDiscMediaTypeAlias() As String
    GetCurDiscMediaTypeAlias = ""
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "GetCurDiscMediaType", "当前设备无效，不能执行该操作。")
        Exit Function
    End If
    
    GetCurDiscMediaTypeAlias = GetDiscMediaTypeAlias(mobjDiscFormat.CurrentPhysicalMediaType)
End Function

'获取当前磁盘媒体类
Public Function GetCurDiscMediaType() As IMAPI_MEDIA_PHYSICAL_TYPE
    GetCurDiscMediaType = IMAPI_MEDIA_TYPE_UNKNOWN
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "GetCurDiscMediaType", "当前设备无效，不能执行该操作。")
        Exit Function
    End If
    
    GetCurDiscMediaType = mobjDiscFormat.CurrentPhysicalMediaType
End Function




'取得光盘刻录时所支持的速率数
Public Function GetCurSupportedSpeedCount() As Long
    GetCurSupportedSpeedCount = -1
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "EjectDisc", "当前设备无效，不能执行该操作。")
        Exit Function
    End If

    GetCurSupportedSpeedCount = UBound(TransformArrayToOneDimension(mobjDiscFormat.SupportedWriteSpeeds))
End Function

'取得光盘刻录所支持的速率
Public Function GetCurSupportedSpeed(ByVal lngSpeedIndex As Long) As Long
    GetCurSupportedSpeed = -1
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "EjectDisc", "当前设备无效，不能执行该操作。")
        Exit Function
    End If
    
    GetCurSupportedSpeed = TransformArrayToOneDimension(mobjDiscFormat.SupportedWriteSpeeds)(lngSpeedIndex + 1)
End Function


'获取支持的媒体格式数量
Public Function GetSupportedMediaTypeCount() As Long
    GetSupportedMediaTypeCount = -1
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "EjectDisc", "当前设备无效，不能执行该操作。")
        Exit Function
    End If
    
    GetSupportedMediaTypeCount = UBound(TransformArrayToOneDimension(mobjDiscFormat.SupportedMediaTypes))
End Function

'取得当前设备所支持的格式类型
Public Function GetCurSupportedMediaTypeAlias(ByVal lngSpeedIndex As Long) As String
    GetCurSupportedMediaTypeAlias = ""
    
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "EjectDisc", "当前设备无效，不能执行该操作。")
        Exit Function
    End If
    
    GetCurSupportedMediaTypeAlias = GetDiscMediaTypeAlias(TransformArrayToOneDimension(mobjDiscFormat.SupportedMediaTypes)(lngSpeedIndex + 1))
End Function







'取得可用光盘大小
Public Function GetDiscFreeSize() As Double
    GetDiscFreeSize = 0
    
    If mobjCurRecord Is Nothing Then
        Exit Function
    End If
    
    GetDiscFreeSize = CDbl(mobjDiscFormat.FreeSectorsOnMedia) * CDbl(2048) - CDbl(mlngReserveSize * 1024)
End Function


'取得光盘总的大小
Public Function GetDiscTotalSize() As Double
    GetDiscTotalSize = 0
    
    If mobjCurRecord Is Nothing Then
        Exit Function
    End If

    GetDiscTotalSize = CDbl(mobjDiscFormat.TotalSectorsOnMedia) * CDbl(2048)
End Function


'弹出设备
Public Sub EjectDisc()
    If mobjCurRecord Is Nothing Then
'        Call Err.Raise(-1, "EjectDisc", "当前设备无效，不能执行该操作。")
        Exit Sub
    End If
    
    mobjCurRecord.EjectMedia
End Sub


'擦除光盘数据（暂时不做该功能，无条件测试.....）
Public Sub EraseDisc()

'    Call Err.Raise(-1, "EraseDisc", "未经实际测试，改功能尚未完善......")
    
    
    If mobjCurRecord Is Nothing Then
        Call err.Raise(-1, "EraseDisc", "当前设备无效，不能执行该操作。")
        Exit Sub
    End If

    Set mobjDiscEarse = Nothing
    Set mobjDiscEarse = New MsftDiscFormat2Erase

    '判断磁盘是否支持数据擦除
    If Not mobjDiscEarse.IsCurrentMediaSupported(mobjCurRecord) Then
        Call err.Raise(-1, "EraseDisc", "当前磁盘不支持数据擦除操作。")
        Exit Sub
    End If

    mobjDiscEarse.Recorder = mobjCurRecord
    mobjDiscEarse.ClientName = C_STR_BURN_CLIENTNAME


    Call mobjDiscEarse.EraseMedia
End Sub


'获取目录下的文件大小
Public Function GetPathSize(ByVal strBurnPath As String) As Double
On Error Resume Next
    Dim objFileSystem As New FileSystemObject
    Dim objDir As Folder
    
    Set objDir = objFileSystem.GetFolder(strBurnPath)
    
    GetPathSize = objDir.Size
    
    Set objDir = Nothing
    Set objFileSystem = Nothing
End Function


'撤销刻录操作
Public Sub CancelBurn()
    mblnIsCancelWrite = True
'    If mobjCurRecord Is Nothing Then
'        Exit Sub
'    End If
'
'    Call mobjDiscFormat.CancelWrite
End Sub


'判断是否存在相同的需要刻录的目录
Private Function HasBurnDirTree(ByVal strDir As String) As Boolean
    Dim i As Long
    
    HasBurnDirTree = False
    
    For i = 1 To mobjBurnDirResources.Count
        If mobjBurnDirResources.Item(i) = strDir Then
            HasBurnDirTree = True
            Exit Function
        End If
    Next i
End Function


'判断是否存在相同的刻录文件
Private Function HasBurnFile(ByVal strFile As String) As Boolean
    Dim i As Long
    
    HasBurnFile = False
    
    For i = 1 To mobjBurnFileResources.Count
        If InStr(mobjBurnFileResources.Item(i), strFile & C_STR_DISC_DIR_SPLIT) > 0 Then
            HasBurnFile = True
            Exit Function
        End If
    Next i
End Function


'添加刻录目录树
Public Sub AddBurnDirTree(ByVal strDir As String)
    '如果存在相同的需要刻录的目录，则退出添加
    If HasBurnDirTree(strDir) Then Exit Sub
    
    Call mobjBurnDirResources.Add(strDir)
End Sub


'添加刻录文件
Public Sub AddBurnFile(ByVal strFile As String, Optional ByVal strToDiscDir As String)
    '如果存在相同的需要刻录的文件，则退出添加
    If HasBurnFile(strFile) Then Exit Sub
    
    Call mobjBurnFileResources.Add(strFile & C_STR_DISC_DIR_SPLIT & strToDiscDir)
End Sub



'移除刻录目录树
Public Sub RemoveBurnDirTree(ByVal strDir As String)
    Dim i As Long
    
    For i = 1 To mobjBurnDirResources.Count
        If mobjBurnDirResources.Item(i) = strDir Then
            Call mobjBurnDirResources.Remove(i)
            Exit Sub
        End If
    Next i
End Sub

'移除不需要刻录的文件
Public Sub RemoveBurnFile(ByVal strFile As String)
    Dim i As Long
    
    For i = 1 To mobjBurnFileResources.Count
        If InStr(mobjBurnFileResources.Item(i), strFile & C_STR_DISC_DIR_SPLIT) > 0 Then
            Call mobjBurnFileResources.Remove(i)
            Exit Sub
        End If
    Next i
End Sub


'清除所有待刻录的目录
Public Sub ClearBurnDir()
    Dim i As Long
    
    For i = mobjBurnDirResources.Count To 1 Step -1
        Call mobjBurnDirResources.Remove(i)
    Next i
End Sub


'清除所有待刻录的文件
Public Sub ClearBurnFile()
    Dim i As Long
    
    For i = mobjBurnFileResources.Count To 1 Step -1
        Call mobjBurnFileResources.Remove(i)
    Next i
End Sub


'获取刻录资源的总大小(单位为字节)
Public Function GetBurnResourceTotalSize() As Double
    Dim i As Long
    Dim lngDirSplitIndex As Long
    Dim strFilePath As String
    
    GetBurnResourceTotalSize = 0
    
    For i = 1 To mobjBurnDirResources.Count
        GetBurnResourceTotalSize = GetBurnResourceTotalSize + GetPathSize(mobjBurnDirResources.Item(i))
    Next i
    
    For i = 1 To mobjBurnFileResources.Count
        strFilePath = mobjBurnFileResources.Item(i)
        
        lngDirSplitIndex = InStr(strFilePath, C_STR_DISC_DIR_SPLIT)
        If lngDirSplitIndex > 0 Then strFilePath = Mid(strFilePath, 1, lngDirSplitIndex - 1)
        
        GetBurnResourceTotalSize = GetBurnResourceTotalSize + FileSystem.FileLen(strFilePath)
    Next i
    
End Function




'读取需要刻录的目录
Private Sub ReadBurnDir(objRootDir As IFsiDirectoryItem)
    Dim i As Long
        
    For i = 1 To mobjBurnDirResources.Count
        mstrBurnPath = mobjBurnDirResources.Item(i)
        
        Call DoBurnProcedureEvent("读取刻录目录：" & mstrBurnPath)
        
        Call objRootDir.AddTree(mobjBurnDirResources.Item(i), mblnIncludeBaseDir)
    Next i
End Sub


'读取需要刻录的文件
Private Sub ReadBurnFile(objRootDir As IFsiDirectoryItem)
    Dim i As Long
    Dim j As Long
    Dim lngSplitIndex As Long
    
    Dim aryDir() As String
    Dim strFilePath As String
    Dim strDiscDir As String
    Dim strFileName As String
    Dim objFsiNewDirItem As IFsiDirectoryItem
    Dim objFsiSubDirItem As IFsiDirectoryItem
    Dim objFileStream As Stream
    
    For i = 1 To mobjBurnFileResources.Count
        strDiscDir = mobjBurnFileResources.Item(i)
        
        '取的刻录文件与磁盘目录的分割所在字符索引
        lngSplitIndex = InStr(strDiscDir, C_STR_DISC_DIR_SPLIT)
        
        strFilePath = Mid(strDiscDir, 1, IIf(lngSplitIndex > 0, lngSplitIndex - 1, 255))
        strFileName = Mid(strFilePath, InStrRev(strFilePath, "\") + 1, 255)
        
        mstrBurnPath = strFilePath
        
        Call DoBurnProcedureEvent("获取刻录文件信息：" & strFilePath)
        
        '判断是否需要在光盘上创建目录
        If lngSplitIndex > 0 Then
            strDiscDir = "\" & Mid(strDiscDir, lngSplitIndex + 1, 255) & "\"
            aryDir = Split(strDiscDir, "\")
            
            Call DoBurnProcedureEvent("创建光盘目录：" & Replace(strDiscDir, "\\", ""))
            
            Set objFsiNewDirItem = Nothing
            
            '在磁盘上创建指定目录
            For j = LBound(aryDir) To UBound(aryDir)
                If Trim(aryDir(j)) <> "" Then
                    On Error Resume Next
                    
                    If objFsiNewDirItem Is Nothing Then
'                        On Error GoTo errCreateDir
                        Set objFsiNewDirItem = objRootDir.Item(aryDir(j))
'errCreateDir:
                        
                        '如果没有获取到指定的光盘目录，则创建
                        If objFsiNewDirItem Is Nothing Then
                            Call objRootDir.AddDirectory(aryDir(j))
                            Set objFsiNewDirItem = objRootDir.Item(aryDir(j))
                        End If
                    Else
'                        On Error GoTo errCreateSubDir
                        Set objFsiSubDirItem = Nothing
                        Set objFsiSubDirItem = objFsiNewDirItem.Item(aryDir(j))
'errCreateSubDir:
                        If objFsiSubDirItem Is Nothing Then
                            Call objFsiNewDirItem.AddDirectory(aryDir(j))
                            Set objFsiNewDirItem = objFsiNewDirItem.Item(aryDir(j))
                        Else
                            Set objFsiNewDirItem = objFsiSubDirItem
                        End If
                    End If
                End If
            Next j
        End If
        
        If objFsiNewDirItem Is Nothing Then Set objFsiNewDirItem = objRootDir
        
        Call DoBurnProcedureEvent("读取文件流：" & strFilePath)
        
        '读取文件流
        Set objFileStream = New Stream
        
        '必须设置为二进制类型，才能保证刻录的文件正确
        objFileStream.type = adTypeBinary
        
        Call objFileStream.Open
        Call objFileStream.LoadFromFile(strFilePath)
        
        Call objFsiNewDirItem.AddFile(strFileName, objFileStream)
        
'        '这个地方不能对流进行关闭，否则在创建镜像时，将产生不能读取流错误
'        Call DoBurnProcedureEvent("关闭文件流：" & strFilePath)
'        Call objFileStream.Close
        
        Set objFileStream = Nothing
        Set objFsiNewDirItem = Nothing
    Next i
End Sub


'准备刻录需要的资源
Private Function ReadyBurnResource(ByRef objFileImage As MsftFileSystemImage, _
    ByVal lngMediaType As IMAPI_MEDIA_PHYSICAL_TYPE) As FsiStream
    
    Dim objRootDir As FsiDirectoryItem
    
    Set ReadyBurnResource = Nothing
    
    Call objFileImage.ChooseImageDefaultsForMediaType(lngMediaType)
    
    '设置刻录光盘盒制作iso的通用属性
On Error GoTo errHandle
    '当对同一光盘进行多次刻录时，设置该属性将会出错，应为该对象已从光盘中读取了系统类型
    objFileImage.FileSystemsToCreate = FsiFileSystemISO9660 Or FsiFileSystemJoliet
errHandle:

On Error GoTo errRead
    objFileImage.VolumeName = mstrBurnVolumeName

    Set objRootDir = objFileImage.Root
    
    '添加需要刻录的目录树
    Call DoBurnProcedureEvent("加载待刻录目录...")
    Call ReadBurnDir(objRootDir)
    
    
    '添加需要刻录的文件
    Call DoBurnProcedureEvent("加载待刻录文件...")
    Call ReadBurnFile(objRootDir)
    
    
    Call DoBurnProcedureEvent("创建镜像数据...")
    
    Set ReadyBurnResource = objFileImage.CreateResultImage.ImageStream
    
    Exit Function
errRead:
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function


'开始刻录
Public Sub StartBurn()
    Dim objBurnVerification As IBurnVerification
    Dim objFileStream As FsiStream
    Dim objIMultiSession    '必须声明无类型对象，才能对MultisessionInterfaces接口赋值
    
    mblnIsCancelWrite = False
    
    If mobjCurRecord Is Nothing Then
        Call DoBurnProcedureEvent("当前设备无效，不能执行该操作...")
        Call err.Raise(-1, "StartBurn", "当前设备无效，不能执行该操作。")
        Exit Sub
    End If
    
    If mobjBurnDirResources.Count <= 0 And mobjBurnFileResources.Count <= 0 Then
        Call DoBurnProcedureEvent("未找到需要刻录的资源，不能执行该操作。")
        Call err.Raise(-1, "StartBurn", "未找到需要刻录的资源，不能执行该操作。")
        Exit Sub
    End If
    
    Call DoBurnProcedureEvent("配置刻录对象...")
    
    Set mobjFileImage = Nothing
    Set mobjFileImage = New MsftFileSystemImage
    
    '选择刻录设备
    Call mobjFileImage.ChooseImageDefaults(mobjCurRecord)
    
    '设置刻录速率
    If mlngWriteSpeed > 0 Then
        Call mobjDiscFormat.SetWriteSpeed(mlngWriteSpeed, True)
    End If
    
    '设置友好名称
    mobjDiscFormat.ClientName = C_STR_BURN_CLIENTNAME
    
    mobjDiscFormat.ForceOverwrite = mblnIsOverWrite
    
    mobjDiscFormat.ForceMediaToBeClosed = mblnOnceMedia
    
    '设置完整性检查级别
    Set objBurnVerification = mobjDiscFormat
    objBurnVerification.BurnVerificationLevel = mlngVerificationLevel
    
    mobjFileImage.FreeMediaBlocks = mobjDiscFormat.FreeSectorsOnMedia - Fix(mlngReserveSize * 1024 / 2048) '添加磁盘的预留大小
    
    '多区段配置(读取光盘上已存在的数据)
    If Not mobjDiscFormat.MediaHeuristicallyBlank Then
        Call DoBurnProcedureEvent("加载光盘数据...")
        
         Set objIMultiSession = mobjFileImage
        
        objIMultiSession.MultisessionInterfaces = mobjDiscFormat.MultisessionInterfaces
        

        mobjFileImage.ImportFileSystem
    End If

    '准备刻录文件
On Error GoTo errHandle
    Call DoBurnProcedureEvent("准备刻录数据...")
    
    Set objFileStream = ReadyBurnResource(mobjFileImage, IMAPI_MEDIA_TYPE_DISK) 'mobjDiscFormat.CurrentPhysicalMediaType
    
    If objFileStream Is Nothing Then
        Call DoBurnProcedureEvent("未能成功获取需要刻录的数据流，操作终止...")
        Call err.Raise(-1, "StartBurn", "未能成功获取需要刻录的数据流，操作终止。")
        Exit Sub
    End If
    
    DoEvents
    
    If mblnIsCancelWrite Then
        Call DoBurnProcedureEvent("刻录操作已被撤销...")
        Call err.Raise(-1, "StartBurn", "刻录操作已被撤销。")
        Exit Sub
    End If
    
    '开始刻录文件
    Call DoBurnProcedureEvent("执行刻录操作...")
    
    Call mobjDiscFormat.Write(objFileStream)
    
    
    Call DoBurnProcedureEvent("完成刻录并弹出光盘...")
    Call mobjCurRecord.EjectMedia
    
    Set objBurnVerification = Nothing
    Set objFileStream = Nothing
    Set mobjFileImage = Nothing
    
    
    Call DoBurnProcedureEvent("已完成数据刻录操作...")
    
    Exit Sub
errHandle:
    Set objBurnVerification = Nothing
    Set objFileStream = Nothing

    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Sub



Private Sub DoBurnProcedureEvent(ByVal strState As String)
On Error Resume Next
'    DoEvents
    
    RaiseEvent OnBurnProcedureEvent(strState)
    
'    DoEvents
End Sub




'刻录ISO文件
Public Sub BurnIso(ByVal strIsoFile As String)
    Dim objBurnVerification As IBurnVerification
    Dim objIsoStream As Stream
    
    mblnIsCancelWrite = False
    
    If mobjCurRecord Is Nothing Then
        Call DoBurnProcedureEvent("当前设备无效，不能执行该操作...")
        Call err.Raise(-1, "BurnIso", "当前设备无效，不能执行该操作。")
        Exit Sub
    End If
    
    Call DoBurnProcedureEvent("配置刻录对象...")
        
    '设置刻录速率
    If mlngWriteSpeed > 0 Then
        Call mobjDiscFormat.SetWriteSpeed(mlngWriteSpeed, True)
    End If
        
    '设置友好名称
    mobjDiscFormat.ClientName = C_STR_BURN_CLIENTNAME
    
    mobjDiscFormat.ForceOverwrite = mblnIsOverWrite
    
    mobjDiscFormat.ForceMediaToBeClosed = mblnOnceMedia
    
    '设置完整性检查级别
    Set objBurnVerification = mobjDiscFormat
    objBurnVerification.BurnVerificationLevel = mlngVerificationLevel
    
    '多区段配置
    If Not mobjDiscFormat.MediaHeuristicallyBlank Then
        Call DoBurnProcedureEvent("刻录ISO数据时，请插入空白光盘。")
        Call err.Raise(-1, "BurnIso", "刻录ISO数据时，请插入空白光盘。")
        
        Exit Sub
    End If
    
    '准备刻录文件
On Error GoTo errHandle
    Call DoBurnProcedureEvent("读取ISO数据...")
    
    Set objIsoStream = New Stream
    objIsoStream.type = adTypeBinary
    
    Call objIsoStream.Open
    Call objIsoStream.LoadFromFile(strIsoFile)
    
    
    If mblnIsCancelWrite Then
        Call DoBurnProcedureEvent("刻录操作已被撤销...")
        Call err.Raise(-1, "StartBurn", "刻录操作已被撤销。")
        Exit Sub
    End If
    
    '开始刻录文件
    Call DoBurnProcedureEvent("执行刻录操作...")
    
    Call mobjDiscFormat.Write(objIsoStream)
    
    
    Call DoBurnProcedureEvent("完成刻录并弹出光盘...")
    Call mobjCurRecord.EjectMedia
    
    Call objIsoStream.Close
    
    Set objIsoStream = Nothing
    Set objBurnVerification = Nothing
    
    Exit Sub
errHandle:
    Set objIsoStream = Nothing
    Set objBurnVerification = Nothing

    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Sub

'保存为ISO
Public Sub SaveIso(ByVal strIsoPath As String)
    Dim objIsoStream As Stream
    Dim objFileStream As FsiStream
    Dim dblInSize As Double
    Dim dblOutSize As Double
    
    Dim tagS As tagSTAT
    Dim objVaraint
    Dim tagVaraint
    Dim bojImageStream As Stream
    
'    'vb对流支持不好，尚不能保存为ISO文件...
'    Exit Sub
    
    Set mobjFileImage = Nothing
    Set mobjFileImage = New MsftFileSystemImage
    
    Call DoBurnProcedureEvent("准备ISO镜像数据...")
    
    Set objFileStream = ReadyBurnResource(mobjFileImage, IMAPI_MEDIA_TYPE_DISK)
    
    If objFileStream Is Nothing Then
        Call DoBurnProcedureEvent("未能成功获取所需的ISO数据流，操作终止...")
        Call err.Raise(-1, "StartBurn", "未能成功获取所需的ISO数据流，操作终止。")
        Exit Sub
    End If
    
    Exit Sub
    
    Set objIsoStream = New Stream
    
    objIsoStream.type = adTypeBinary
    objIsoStream.Mode = adModeWrite
    
    
    objVaraint = objFileStream
'
'    Set bojImageStream = objVaraint
    
'    Call objVaraint.Stat(tagVaraint, 1)
    Call objVaraint.RemoteCopyTo(objIsoStream, 255, dblInSize, dblOutSize)
'    Call objFileStream.Clone(objIsoStream)
    
    Call objIsoStream.SaveToFile(strIsoPath, adSaveCreateOverWrite)
    
    objIsoStream.Close
    
    Set objIsoStream = Nothing
    Set objFileStream = Nothing
End Sub



Private Sub Class_Initialize()
    Set mobjDiscs = New MsftDiscMaster2
    Set mobjDiscFormat = New MsftDiscFormat2Data
    Set mobjFileImage = Nothing
    Set mobjCurRecord = Nothing
    Set mobjDiscEarse = Nothing
    Set mobjBurnDirResources = New Collection
    Set mobjBurnFileResources = New Collection
    
    mstrBurnDeviceName = ""
    mstrBurnVolumeName = Format(zlDatabase.Currentdate, "yyyymmdd")
    mblnIsOverWrite = True
    mlngVerificationLevel = ivlFull
    mblnOnceMedia = False
    mblnIncludeBaseDir = False
    mblnIsCancelWrite = False
    mstrBurnPath = ""
    mlngReserveSize = 0 '10 * 1024 '(10M)
End Sub

Private Sub Class_Terminate()
    Set mobjBurnFileResources = Nothing
    Set mobjBurnDirResources = Nothing
    Set mobjDiscEarse = Nothing
    Set mobjCurRecord = Nothing
    Set mobjFileImage = Nothing
    Set mobjDiscFormat = Nothing
    Set mobjDiscs = Nothing
End Sub


'获取写入状态
Private Function GetWriteAction(iFormat2DataArgs As IDiscFormat2DataEventArgs)
    Select Case iFormat2DataArgs.CurrentAction
        Case 1
            GetWriteAction = "格式化媒体..."
        Case 2
            GetWriteAction = "初始化设备..."
        Case 3
            GetWriteAction = "优化写入介质..."
        Case 4
            GetWriteAction = "写入数据到媒体..."
        Case 5
            GetWriteAction = "同步缓冲并关闭磁道..."
        Case 6
            GetWriteAction = "成功写入数据..."
        Case 7
            GetWriteAction = "检查完整性..."
        Case Else
            GetWriteAction = "确认媒体支持性..."
    End Select
End Function

Private Sub mobjDiscFormat_Update(ByVal Object As Object, ByVal Progress As Object)
On Error Resume Next
    Dim iFormat2DataArgs As IDiscFormat2DataEventArgs
    Dim strState As String
    Dim args As New clsImapi2BurnArgs
    
    If mblnIsCancelWrite Then
        Call Object.CancelWrite
    End If
    
    Set iFormat2DataArgs = Progress

    strState = GetWriteAction(iFormat2DataArgs)
    
    args.ElapsedTime = iFormat2DataArgs.ElapsedTime
    args.FreeSystemBuffer = iFormat2DataArgs.FreeSystemBuffer
    args.LastReadLba = iFormat2DataArgs.LastReadLba
    args.LastWrittenLba = iFormat2DataArgs.LastWrittenLba
    args.RemainingTime = iFormat2DataArgs.RemainingTime
    args.SectorCount = iFormat2DataArgs.SectorCount
    args.StartLba = iFormat2DataArgs.StartLba
    args.TotalSystemBuffer = iFormat2DataArgs.TotalSystemBuffer
    args.TotalTime = iFormat2DataArgs.TotalTime
    args.UsedSystemBuffer = iFormat2DataArgs.UsedSystemBuffer
    
'    DoEvents
    
    RaiseEvent OnBurnEvent(strState, args)
    
'    DoEvents
    
    Set args = Nothing
End Sub

Private Sub mobjDiscs_NotifyDeviceAdded(ByVal Object As Object, ByVal uniqueId As String)
    '......
End Sub

Private Sub mobjDiscs_NotifyDeviceRemoved(ByVal Object As Object, ByVal uniqueId As String)
    '......
End Sub

Private Sub mobjFileImage_Update(ByVal Object As Object, ByVal currentFile As String, ByVal copiedSectors As Long, ByVal totalSectors As Long)
On Error Resume Next
    Dim strCurBurnPath As String
'    Dim dblCurFileSize As Double
    
    strCurBurnPath = IIf(mblnIncludeBaseDir, Mid(mstrBurnPath, 1, InStrRev(mstrBurnPath, "\") - 1), strCurBurnPath)
'
'    dblCurFileSize = FileSystem.FileLen(strCurBurnPath & currentFile)
    
'    DoEvents
    
    RaiseEvent OnReadyFileEvent(currentFile, strCurBurnPath, copiedSectors, totalSectors)
    
'    DoEvents
End Sub
