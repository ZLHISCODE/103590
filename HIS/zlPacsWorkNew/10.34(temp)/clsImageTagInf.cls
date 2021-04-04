VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageTagInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public EncoderName As String '编码名称
Public CaptureTime As String '采集时间
Public RecordTimeLen As Long '录制长度
Public VideoFile As String   '视频文件名称
Public Tag As Variant        '标记 VIDEOTAG: 表示视频，IMGTAG:表示图像

