VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueryPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public objPatiIdentify As Object
Public objFilterCmdBar As Object
Attribute objFilterCmdBar.VB_VarHelpID = -1
Public ImageList16 As Object
Public ImageList24 As Object
Public picList As Object

Public rtpAppend As Object
Public vsfList As Object
Public PicContainer As Object
Public cmdDo As Object


Public picFollow As Object
Public img1 As Object

Public img2 As Object
Public img3 As Object
Public img4 As Object
Public cmdFind As Object
Public TimerFunc As Object
Attribute TimerFunc.VB_VarHelpID = -1

Public TabCtl As Object
Public rtfHisFollow As Object
Public PicHisFollow As Object
Public TimerHisFunc As Object

Public picTemp As Object
Public labPatiInfo As Object

Public Type TColConfig
    strName As String
    lngWidth As Long
    lngColOrder As Long
    lngCfgOrder As Long
    blnIsUserHide As Boolean '是否用户自己手动隐藏
    blnIsSysHide As Boolean '是否配置中设置隐藏列
    blnHideTitle As Boolean    '是否配置中设置隐藏标题
    blnHiddenData As Boolean '是否配置中设置内容显示
End Type






