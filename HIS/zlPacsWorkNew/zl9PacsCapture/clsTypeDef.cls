VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTypeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'工作模块回调事件类型
Public Enum TVideoEventType
    vetDelAllImg = 0        '删除所有图像
    vetGetImg = 1           '获取图像

    vetLockStudy = 2        '锁定检查
    vetUnLockStudy = 3      '解锁检查

    vetCaptureFirstImg = 4  '采集第一幅图像
    vetUpdateImg = 5        '更新图像
End Enum
