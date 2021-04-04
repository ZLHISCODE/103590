VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRDemoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'病历文件示范
Public ID As Long                           '示范ID
Public 文件ID As Long                       '原始文件定义ID
Public 编号 As String                       '示范编号
Public 名称 As String                       '示范名称
Public 性质 As Integer                      '示范性质:0-范文;1-片段
Public 通用级 As Long                       '0-全院通用;1-部门通用;2-个人使用
Public 科室ID As Long                       '记录科室ID
Public 人员ID As Long                       '记录人员ID

