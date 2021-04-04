VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public InstanceUID As String        '图像的UID（主关键字）
Public ImageNo As String            '图像号 (0020,0013) : Image Number
Public ImageName As String          '图像文件名，对于直接从本地目录打开的图像，需要填写真实名称，对于从工作站打开的，填写图像UID。
Public blnSelected As Boolean       '是否被选择，仅在ZLShowSeriesInfos中使用。
Public blnDisplayed As Boolean      '是否已显示，仅在ZLShowSeriesInfos中使用。
Public AcquisitionTime As String    '采集时间 (0008,0022) : Acquisition Date+ (0008,0032) : Acquisition Time
Public ImageTime As String          '图像时间 (0008,0023) : Image Date  + (0008,0033) : Image Time
Public SliceThickness As String     '层厚 (0018,0050): Slice Thickness
Public ImagePositionPatient As String       '图像位置病人 (0020,0032) : Image Position (Patient)
Public ImageOrientationPatient As String    '图像方向病人 (0020,0037) : Image Orientation (Patient)
Public FrameOfReferenceUID As String        '参考帧UID (0020,0052) : Frame of Reference UID
Public SliceLocation As String      '切片位置 (0020,1041) : Slice Location
Public Rows As String               '行数 (0028,0010) : Rows
Public Columns As String            '列数 (0028,0011) : Columns
Public PixelSpacing As String       '像素距离 (0028,0030): Pixel Spacing
Public int3DLabelIndex  As Integer     '三维鼠标中定位线的索引
Public blnPrinted As Boolean        '是否已经打印胶片
Public ImgTag As Long               '图像类型， 0：图像，1：多帧图，2：录像，3：录音（不使用）

