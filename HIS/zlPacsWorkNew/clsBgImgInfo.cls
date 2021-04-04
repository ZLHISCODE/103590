VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBgImgInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'图像命令
Public Enum TImageCommand
    icReadly = 0      '读取显示
    icDownload = 1    '下载命令
    icUpLoad = 2      '上传命令
End Enum

'图像格式
Public Enum TImageFmt
    ifDcm = 0
    ifBmp = 1
    ifJpg = 2
    ifAvi = 3
    ifWav = 4
End Enum

'状态图类型
Public Enum TStateImageType
    sitNul = 0
    sitDown = 1
    sitAvi = 2
    sitWav = 3
    sitErr = 4
End Enum

'加载状态
Public Enum TLoadState
    lsNone = 0  '默认，没有进行任何操作
    lsMedia = 1     '是多媒体文件
    lsLocal = 2     '已存在本地文件
    lsSent = 3      '已发送传输命令
    lsError = 4     '加载异常
    lsRedo = 5      '重试
End Enum
 
Public Key As String           '关键字，如图像UID
Public PatientName  As String   '患者名称
Public AdviceId As Long         '医嘱ID
Public StudyUID As String       '检查UID
Public RecFmtDate As String     '接受日期 已格式化  YYYYMMDD
Public AdviceDes As String      '医嘱描述

Public Filename As String      '文件名称
Public FilePath As String      '文件路径

Public DeviceNo As Long         '设备号
Public FtpIp As String             'IP
Public FtpPort As Long             '端口
Public FtpUser As String           'User
Public FtpPwd As String            '密码
Public FtpVirtualPath As String    '虚拟目录
Public FtpShareDir As String
Public FtpShareUser As String
Public FtpSharePwd As String

Public BakDeviceNo As Long          '备份设备号
Public BakIp As String             '备份FTPIP
Public BakPort As Long             '备份FTP端口
Public BakUser As String           '备份FTPUser
Public BakPwd As String            '备份FTP密码
Public BakVirtualPath As String    '备份虚拟目录

Public FtpFile As String        'Ftp文件

Public SeriesNoTag As String
Public ImageOrder As Long       '绘制序号
Public DrawHint As String       '绘制提示文本
Public ImgTag As String
Public ImgCommand As TImageCommand '图像命令
Public IsBackGround As Boolean  '是否使用后台传输方式

'Public EncoderName As String   '编码名称
'Public RecordTime As String    '录制时间
'Public RecordLen As Long       '录制长度
Public Format As TImageFmt      '格式
Public LoadState As TLoadState  '图像加载状态
Public IsReset As Boolean       '为true表示重新下载

Public StartTime As Date        '开始处理时间
Public EndTime As Date          '结束处理时间

Public ErrorInfo As String      '错误信息
Public Redo As Long             '重试次数
Public IsReDrawed As Boolean    '是否已经重绘

Public JpgConvert As Boolean    '是否进行JPG转换
Public IsCompress As Boolean    '是否进行压缩后传输

Public IsCreateCmdFile As Boolean   '是否已经创建了命令文件


Public Function CopyNew() As clsBgImgInfo
    Dim objCurBgImgInfo As New clsBgImgInfo
    
    With objCurBgImgInfo
        .Key = Key
        .PatientName = PatientName
        .AdviceId = AdviceId
        .AdviceDes = AdviceDes
        .StudyUID = StudyUID
    
        .Filename = Filename
        .FilePath = FilePath
    
        .FtpIp = FtpIp
        .FtpPort = FtpPort
        .FtpUser = FtpUser
        .FtpPwd = FtpPwd
        .FtpVirtualPath = FtpVirtualPath
        .FtpShareDir = FtpShareDir
        .FtpShareUser = FtpShareUser
        .FtpSharePwd = FtpSharePwd
        
        .BakIp = BakIp
        .BakPort = BakPort
        .BakUser = BakUser
        .BakPwd = BakPwd
        .BakVirtualPath = BakVirtualPath
    
        .FtpFile = FtpFile
    
        .SeriesNoTag = SeriesNoTag
        .ImageOrder = ImageOrder
        .DrawHint = DrawHint
        .ImgTag = ImgTag
        .ImgCommand = ImgCommand
        .IsBackGround = IsBackGround
     
        .Format = Format
        .LoadState = LoadState
        .IsReset = IsReset
    
        .StartTime = StartTime
        .EndTime = EndTime
    
        .ErrorInfo = ErrorInfo
        .Redo = Redo
        .IsReDrawed = IsReDrawed
        
        .JpgConvert = JpgConvert
        .IsCompress = IsCompress
    End With
    
    Set CopyNew = objCurBgImgInfo
End Function







