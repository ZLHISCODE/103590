VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSeriesInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SeriesUID As String          '序列UID（主关键字）
Public SeriesNo As String           '序列号，(0020,0011) : Series Number
Public StudyUID As String           '检查UID
Public ImageInfos As New Collection '图像信息的集合
Public lngSource As Long            '图像来源，0-从PACS图像服务器下载；1-直接打开文件；2---混合；3-重新生成的序列，类似矢冠状位重建、图像拼接、伪彩生成的图像
Public strHostIP As String          '存储的主机IP，即用于FTP，也用于共享目录
Public strFTPUser As String         'FTP 用户名
Public strFTPPasw As String         'FTP 密码
Public strFTPDir As String          'FTP目录，FTP的子目录，格式是"/FTP目录/"
Public strSaveDir As String         '保存图像的具体路径，由"接收日期\检查UID\"组合而成。读取FTP目录和保存到本地缓存时，都使用这个参数
Public strShareDir As String        '共享目录方式的目录,以下这个结构：""\\"+strHostIP+"\"+ strShareDir+"\"+ strSaveDir" 就是保存图像的具体路径，结构是"\\IP 地址\共享目录名\文件目录"
Public strShareDirUser As String    '共享目录方式的用户名
Public strShareDirPasw As String    '共享目录方式的密码
Public intSortType As Integer       '记录当前序列的排序方式：0--图像号；1--床位正序；2--床位逆序；3--采集时间；4--图像时间，仅在ZLShowSeriesInfos中使用。
Public strModality As String        '影像类别 (0008,0060) : Modality
Public lngWinWidth As Long          '当前序列的窗宽，创建序列的时候，从"影像预设窗宽窗位"表中读取预设的窗宽，手工设置F3-F12的窗宽后，记录设置后的窗宽，仅在ZLShowSeriesInfos中使用。
Public lngWinLevel As Long          '当前序列的窗位，创建序列的时候，从"影像预设窗宽窗位"表中读取预设的窗位，手工设置F3-F12的窗位后，记录设置后的窗位，仅在ZLShowSeriesInfos中使用。
Public intRow As Integer            '当前序列所在的行数。仅在ZLShowSeriesInfos中使用。
Public intCol As Integer            '当前序列所在的列数。仅在ZLShowSeriesInfos中使用。
Public blnImageSyn As Boolean       '序列内的图像是否同步。仅在ZLShowSeriesInfos中使用。
Public StretchToFit As Boolean      '自适应缩放。仅在ZLShowSeriesInfos中使用。
Public ScrollX As Long              'X方向移动。仅在ZLShowSeriesInfos中使用。
Public ScrollY As Long              'Y方向移动。仅在ZLShowSeriesInfos中使用。
Public Zoom As Double               '缩放倍数。仅在ZLShowSeriesInfos中使用。
Public FlipState As Integer         '镜像状态。仅在ZLShowSeriesInfos中使用。
Public RotateState As Integer       '旋转状态。仅在ZLShowSeriesInfos中使用。
Public FilterLength As Integer      '平滑。仅在ZLShowSeriesInfos中使用。
Public UnsharpEnhancement As Double '边缘增强强度。仅在ZLShowSeriesInfos中使用。
Public UnsharpLength As Integer     '边缘增强幅度。仅在ZLShowSeriesInfos中使用。
Public MultiFrame As Integer        '是否多帧图像。0-单帧图像，1-多帧图像，2-视频录像，3-音频数据（不使用）。
Public Selected As Boolean          '是否被选择，用于图像序列同步。仅在ZLShowSeriesInfos中使用。
Public strCName As String           '图像在数据库中的中文名，直接打开的文件此字段为空
Public strEName As String           '图像在数据库中的拼音名，直接打开的文件此字段为空
Public strAge As String             '图像在数据库中的年龄，直接打开的文件此字段为空
Public strStudyID As String         '图像在数据库中的检查号，直接打开的文件此字段为空
Public strSex As String             '图像在数据库中的性别，直接打开的文件此字段为空
Public strOrderID As String         '图像在数据库中的医嘱ID，直接打开的文件此字段为空

Private Sub Class_Initialize()
    Set ImageInfos = New Collection
End Sub
