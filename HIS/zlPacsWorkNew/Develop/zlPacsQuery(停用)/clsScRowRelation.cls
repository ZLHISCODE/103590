VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScRowRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

  

Public TiggerData As String         '触发数据
Public Icon As String               '图标
Public IconPerformCol As String     '图标呈现列
Public IsStateIcon As Boolean       '是否状态图标
Public RowFontColor As Long         '行字体颜色
Public RowBackColor As Long         '行背景颜色
Public CellFontColor As Long        '单元格字体颜色
Public CellBackColor As Long        '单元格背景颜色
Public ColorPerformCol As String    '颜色呈现列
Public FlickerTimeOut As Long       '颜色超时闪烁时长
Public TimeOutReferCol As String    '超时时间参考列

