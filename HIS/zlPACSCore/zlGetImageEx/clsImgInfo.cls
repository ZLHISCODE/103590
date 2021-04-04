VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImgInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrSubDir As String
Private mstrDestMainDir As String           '复制图像的目的目录，本机目录
Private mblnEnable As Boolean               '本消息可用
Private mblnUpLoad As Boolean               '值为True时上传，为False时下载
Private mlngMediaType As Long               '动态图
Private mblnAfterCapture As Boolean         '是否为后台采集
Private mblnLoadSingleFile As Boolean       '是否单个载入，true-是，false-否

Private mdcmTmpImg As Object                'DicomImage

Private mlngThreadID As Long

Private Type TFtpDeviceInf
    strDeviceId As String
    strFTPIP As String
    strFTPUser As String
    strFTPPwd As String
    strFtpDir As String
    strSDDir As String
    strSDUser As String
    strSDPswd As String
End Type

Private mobjFtp As TFtpDeviceInf
Private mobjBakFtp As TFtpDeviceInf

Private mlngAdviceId As Long
Private mstrStudyUid As String
Private mdtNowTime As Date
Private mstrSeriesUID As String
Private mstrInstanceUID As String
Private mstrTag As String
Private mstrSeriesDescription As String
Private mstrEncoderName As String
Private mlngRecordTimeLen As Long
Private mblnIsFirstImage As Boolean
Private mlngSendNo As Long
Private mstrDTReceived As String

'后台采集信息
Private mstrAfterModality As String
Private mstrDeviceId As String
Private mstrAfterStationName As String

Private mstrTmpFilePath As String
Private mstrCaptureTime As String

'''''''''''''''''''''''''''''
Public Property Get SubDir() As String               '图像所在的子目录
    SubDir = mstrSubDir
End Property

Public Property Let SubDir(ByVal vNewValue As String)
    mstrSubDir = vNewValue
End Property

Public Property Get DestMainDir() As String         '复制图像的目的目录，本机目录
    DestMainDir = mstrDestMainDir
End Property
 
Public Property Let DestMainDir(ByVal vNewValue As String)
    mstrDestMainDir = vNewValue
End Property


'存储设备
Public Property Get IP() As String                   '图像服务器的IP地址
    IP = mobjFtp.strFTPIP
End Property

Public Property Let IP(ByVal vNewValue As String)
    mobjFtp.strFTPIP = vNewValue
End Property

Public Property Get FTPDir() As String               'FTP目录
    FTPDir = mobjFtp.strFtpDir
End Property

Public Property Let FTPDir(ByVal vNewValue As String)
    mobjFtp.strFtpDir = vNewValue
End Property

Public Property Get FTPUser() As String              'FTP用户名
    FTPUser = mobjFtp.strFTPUser
End Property

Public Property Let FTPUser(ByVal vNewValue As String)
    mobjFtp.strFTPUser = vNewValue
End Property

Public Property Get FTPPswd() As String              'FTP密码
    FTPPswd = mobjFtp.strFTPPwd
End Property

Public Property Let FTPPswd(ByVal vNewValue As String)
    mobjFtp.strFTPPwd = vNewValue
End Property

Public Property Get SDDir() As String                '共享目录名称
    SDDir = mobjFtp.strSDDir
End Property

Public Property Let SDDir(ByVal vNewValue As String)
    mobjFtp.strSDDir = vNewValue
End Property

Public Property Get SDUser() As String               '共享目录用户名
    SDUser = mobjFtp.strSDUser
End Property

Public Property Let SDUser(ByVal vNewValue As String)
    mobjFtp.strSDUser = vNewValue
End Property

Public Property Get SDPswd() As String               '共享目录密码
    SDPswd = mobjFtp.strSDPswd
End Property

Public Property Let SDPswd(ByVal vNewValue As String)
    mobjFtp.strSDPswd = vNewValue
End Property



'备份设备
Public Property Get BakIP() As String                   '图像服务器的IP地址
    BakIP = mobjBakFtp.strFTPIP
End Property

Public Property Let BakIP(ByVal vNewValue As String)
    mobjBakFtp.strFTPIP = vNewValue
End Property

Public Property Get BakFTPDir() As String               'FTP目录
    BakFTPDir = mobjBakFtp.strFtpDir
End Property

Public Property Let BakFTPDir(ByVal vNewValue As String)
    mobjBakFtp.strFtpDir = vNewValue
End Property

Public Property Get BakFTPUser() As String              'FTP用户名
    BakFTPUser = mobjBakFtp.strFTPUser
End Property

Public Property Let BakFTPUser(ByVal vNewValue As String)
    mobjBakFtp.strFTPUser = vNewValue
End Property

Public Property Get BakFTPPswd() As String              'FTP密码
    BakFTPPswd = mobjBakFtp.strFTPPwd
End Property

Public Property Let BakFTPPswd(ByVal vNewValue As String)
    mobjBakFtp.strFTPPwd = vNewValue
End Property

Public Property Get BakSDDir() As String                '共享目录名称
    BakSDDir = mobjBakFtp.strSDDir
End Property

Public Property Let BakSDDir(ByVal vNewValue As String)
    mobjBakFtp.strSDDir = vNewValue
End Property

Public Property Get BakSDUser() As String               '共享目录用户名
    BakSDUser = mobjBakFtp.strSDUser
End Property

Public Property Let BakSDUser(ByVal vNewValue As String)
    mobjBakFtp.strSDUser = vNewValue
End Property

Public Property Get BakSDPswd() As String               '共享目录密码
    BakSDPswd = mobjBakFtp.strSDPswd
End Property

Public Property Let BakSDPswd(ByVal vNewValue As String)
    mobjBakFtp.strSDPswd = vNewValue
End Property



'''''''''''''''''''''''''''
Public Property Get Enable() As Boolean             '本消息可用
    Enable = mblnEnable
End Property

Public Property Let Enable(ByVal vNewValue As Boolean)
    mblnEnable = vNewValue
End Property

Public Property Get IsUpLoad() As Boolean              '值为True时上传，为False时下载
    IsUpLoad = mblnUpLoad
End Property

Public Property Let IsUpLoad(ByVal vNewValue As Boolean)
    mblnUpLoad = vNewValue
End Property

Public Property Get MediaType() As Long        '
    MediaType = mlngMediaType
End Property

Public Property Let MediaType(ByVal vNewValue As Long)
    mlngMediaType = vNewValue
End Property

Public Property Get IsAfterCapture() As Boolean        '是否为后台采集
    IsAfterCapture = mblnAfterCapture
End Property

Public Property Let IsAfterCapture(ByVal vNewValue As Boolean)
    mblnAfterCapture = vNewValue
End Property

Public Property Get IsLoadSingleFile() As Boolean      '是否单个载入，true-是，false-否
    IsLoadSingleFile = mblnLoadSingleFile
End Property

Public Property Let IsLoadSingleFile(ByVal vNewValue As Boolean)
    mblnLoadSingleFile = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
Public Property Get AdviceId() As Long
    AdviceId = mlngAdviceId
End Property

Public Property Let AdviceId(ByVal vNewValue As Long)
    mlngAdviceId = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get StudyUid() As String
    StudyUid = mstrStudyUid
End Property

Public Property Let StudyUid(ByVal vNewValue As String)
    mstrStudyUid = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get NowTime() As Date
    NowTime = mdtNowTime
End Property

Public Property Let NowTime(ByVal vNewValue As Date)
    mdtNowTime = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get SeriesUID() As String
    SeriesUID = mstrSeriesUID
End Property

Public Property Let SeriesUID(ByVal vNewValue As String)
    mstrSeriesUID = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get InstanceUID() As String
    InstanceUID = mstrInstanceUID
End Property

Public Property Let InstanceUID(ByVal vNewValue As String)
    mstrInstanceUID = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get Tag() As String
    Tag = mstrTag
End Property

Public Property Let Tag(ByVal vNewValue As String)
    mstrTag = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get SeriesDescription() As String
    SeriesDescription = mstrSeriesDescription
End Property

Public Property Let SeriesDescription(ByVal vNewValue As String)
    mstrSeriesDescription = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get EncoderName() As String
    EncoderName = mstrEncoderName
End Property

Public Property Let EncoderName(ByVal vNewValue As String)
    mstrEncoderName = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get RecordTimeLen() As Long
    RecordTimeLen = mlngRecordTimeLen
End Property

Public Property Let RecordTimeLen(ByVal vNewValue As Long)
    mlngRecordTimeLen = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get IsFirstImage() As Boolean
    IsFirstImage = mblnIsFirstImage
End Property

Public Property Let IsFirstImage(ByVal vNewValue As Boolean)
    mblnIsFirstImage = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get SendNo() As Long
    SendNo = mlngSendNo
End Property

Public Property Let SendNo(ByVal vNewValue As Long)
    mlngSendNo = vNewValue
End Property

'''''''''''''''''''''''''''''''''''
Public Property Get DTReceived() As String
    DTReceived = mstrDTReceived
End Property

Public Property Let DTReceived(ByVal vNewValue As String)
    mstrDTReceived = vNewValue
End Property

''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''

Public Property Get AfterModality() As String
    AfterModality = mstrAfterModality
End Property

Public Property Let AfterModality(ByVal vNewValue As String)
    mstrAfterModality = vNewValue
End Property

''''''''''''''''''''''''''''''''''''
Public Property Get DeviceId() As String
    DeviceId = mstrDeviceId
End Property

Public Property Let DeviceId(ByVal vNewValue As String)
    mstrDeviceId = vNewValue
End Property

''''''''''''''''''''''''''''''''''''
Public Property Get AfterStationName() As String
    AfterStationName = mstrAfterStationName
End Property

Public Property Let AfterStationName(ByVal vNewValue As String)
    mstrAfterStationName = vNewValue
End Property

''''
Public Property Get TmpFilePath() As String
    TmpFilePath = mstrTmpFilePath
End Property

Public Property Let TmpFilePath(ByVal vNewValue As String)
    mstrTmpFilePath = vNewValue
End Property

Public Property Get CaptureTime() As String
    CaptureTime = mstrCaptureTime
End Property

Public Property Let CaptureTime(ByVal vNewValue As String)
    mstrCaptureTime = vNewValue
End Property

Public Property Get dcmTmpImg() As Object
    Set dcmTmpImg = mdcmTmpImg
End Property

Public Property Set dcmTmpImg(ByVal vNewValue As Object)
    Set mdcmTmpImg = vNewValue
End Property

Public Property Get ThreadID() As Long
    ThreadID = mlngThreadID
End Property

Public Property Let ThreadID(ByVal vNewValue As Long)
    mlngThreadID = vNewValue
End Property
