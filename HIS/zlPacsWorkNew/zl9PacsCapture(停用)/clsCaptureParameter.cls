VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaptureParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public VideoDirverType As TVideoDriverType      '视频驱动类型
Public IsSountHint As Boolean                   '声音提示
Public IsWindowHint As Boolean                  '弹出提示
Public ScanDeviceTmpDir As String               '扫描设备临时目录
Public CustomDevicePath As String                '专用视频采集部件路径

Public ComPortType As String                    '端口类型
Public ComPortName As String                    '使用的端口名称
Public ComInterval As Long                      '脚踏采集间隔时长
Public CaptureWay As Long                       '脚踏采集方式

Public ShowBigImage As Integer                '是否显示大图，0-不显示，1-鼠标移动时显示，2-单击时显示
Public ImageZoom As Double                        '图像放大倍数
Public IsAllowChangeSize As Boolean              '是否允许改变采集区域大小

Public IsUseAfterCapture As Boolean             '是否启用后台采集
Public AfterCaptureModality As String           '后台影像类别
Public CurStorageDeviceNo As String             '存储设备号
Public BakStorageDeviceNo As String             '备份设备号
Public IsUseCaptureLock As Boolean              '是否启用采集锁定

Public strCaptureHot As String                  '采集热键


Public Sub ReadParameter()
'初始化参数设置
    Dim rsTmp As New ADODB.Recordset
    Dim intVideoCapture As Integer
    Dim strSQL As String
    
BUGEX "ReadParameter 1", True
    IsUseAfterCapture = False
    AfterCaptureModality = "OT"

    
BUGEX "ReadParameter 2", True
    '读取驱动类型
    VideoDirverType = zlDatabase.GetPara("视频驱动类型", glngSys, glngModule, "0")
    
    
    '读取提示类型
    IsSountHint = Val(zlDatabase.GetPara("采集后声音提示", glngSys, glngModule, True))
    IsWindowHint = Val(zlDatabase.GetPara("采集后弹窗提示", glngSys, glngModule, True))
    
    '读取扫描设备临时存储的图像目录
    ScanDeviceTmpDir = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "扫描设备临时目录", "C:\Documents and Settings\All Users\Application Data\Microsoft\WIA")
  
    '专用视频采集
    CustomDevicePath = GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "专用视频采集")
    
    '读取串口的参数
    ComPortName = zlDatabase.GetPara("脚踏端口", glngSys, glngModule, "1")
    If IsNumeric(ComPortName) Then
        '使用Com采集端口
        ComPortName = CLng(ComPortName)
        ComPortType = "COM"
        
        CaptureWay = zlDatabase.GetPara("脚踏采集方式", glngSys, glngModule, "1")
        If CaptureWay < 0 Or CaptureWay > 2 Then CaptureWay = 1
        
        '读取脚踏间隔时间
        ComInterval = zlDatabase.GetPara("脚踏时间间隔", glngSys, glngModule, "1")
    ElseIf ComPortName = "无" Or Trim(ComPortName) = "" Then
        '不使用采集端口
        ComPortType = "无"
    Else
        '使用HID采集设备
        ComPortType = "HID"
    End If
BUGEX "ReadParameter 3", True

    'strCaptureHot = zlDatabase.GetPara("采集热键", glngSys, glngModule, "F8")
    strCaptureHot = GetSetting("ZLSOFT", "公共模块", "采集热键", "F8")
    
    '是否自动显示大图
    ShowBigImage = Val(zlDatabase.GetPara("鼠标移动时显示大图", glngSys, glngModule, "0"))
    ImageZoom = zlDatabase.GetPara("采集大图放大倍数", glngSys, glngModule, "1")
        
BUGEX "ReadParameter 4", True
    '初始化科室级参数==============================================================================
    IsUseAfterCapture = GetDeptPara(glngDepartId, "启用后台采集", 0)
    AfterCaptureModality = GetDeptPara(glngDepartId, "后台影像类别", "OT")
    
    '读取并检测存储设备号
    CurStorageDeviceNo = GetDeptPara(glngDepartId, "存储设备号")
    BakStorageDeviceNo = GetDeptPara(glngDepartId, "备份设备号")
    
    '启用采集锁定
    IsUseCaptureLock = GetDeptPara(glngDepartId, "启用采集锁定", True)
    '改变采集区域大小
    IsAllowChangeSize = GetDeptPara(glngDepartId, "允许改变采集区域大小", True)
    
BUGEX "ReadParameter 5", True
End Sub
