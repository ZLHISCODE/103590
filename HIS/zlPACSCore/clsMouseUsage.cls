VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''
''''保存鼠标的用法设置
''''黄捷
'''''''''''''''''''''''''''''''''''''''''''
Public bModified As Boolean                                ''是否被更改了
Public lngFuncNo As Long                                   ''功能序号
Public lngMouseKey As Long                                 ''鼠标键位
Public lngShift As Long                                    ''Shift键位
Public strShowName As String                               ''按键名称
Public bSelected As Boolean                                ''是否被选中
Public strProgramName As String                            ''按钮程序名称
Public ButtomID As Long                                    ''按钮ID
