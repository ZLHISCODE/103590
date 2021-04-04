VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScSerachCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Name As String
Public InputType As TInputType  '录入类型0-弹出录入，1-列表上方下拉快捷录入，2-两者都有
Public ControlType As Long  '弹出录入的控件类型
Public ExtProperty As String  '扩展属性
'Public SourceType As Long
Public Default As String    '默认值
Public DataFrom As String   '数据来源配置
'Public IsCustom As Boolean
