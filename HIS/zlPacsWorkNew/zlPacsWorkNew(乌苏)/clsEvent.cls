VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'工作模块回调事件
Public Enum TWorkEventType
    wetDelAllImg = 0        '删除所有图像
    wetGetImg = 1           '获取图像
    wetTechDo = 2           '技师执行
    wetChangeImgType = 3    '更改图像类型
    
    wetLockStudy = 4        '锁定检查
    wetUnLockStudy = 5      '解锁检查
    wetCaptureFirstImg = 6  '采集第一幅图像
    wetUpdateImg = 7        '更新图像
    wetChangeUser = 8       '交换用户
    
    wetPatholRequest = 9    '病理申请
    wetPatholQuality = 10   '病理质量
    wetPatholBatSlices = 11 '制片批量处理
    wetPatholBatSpeExm = 12 '特检批量处理
    
    wetSpecimenAccept = 13  '标本核收
    wetSpecimenReject = 14  '标本拒收
    wetSpecimenSave = 15    '标本保存
    
    wetMaterialSure = 16    '取材确认
    wetMaterialSave = 17    '材块保存
    
    wetSlicesSure = 18      '制片确认
    
    wetSpeExamSure = 19     '特检确认
    
    wetViewEprReport = 20   '预览电子病历报告
    wetViewPacsImage = 21   '预览Pacs图像
    wetRejectReport = 22    '驳回报告
    
    wetPrintFilm = 23       '胶片打印
End Enum



Public Event OnWork(objEvent As Object, ByVal lngWorkType As TWorkEventType, ByVal lngAdviceID As Long, ByVal other As Variant)



Public Sub DoWork(objEvent As Object, ByVal lngWorkType As TWorkEventType, ByVal lngAdviceID As Long, Optional other As Variant = "")
On Error GoTo ErrHandle
    RaiseEvent OnWork(objEvent, lngWorkType, lngAdviceID, other)
    
ErrHandle:
End Sub
