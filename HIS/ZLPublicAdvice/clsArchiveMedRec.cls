VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchiveMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'类的模块变量
Private mclsArchive As zlMedRecPage.clsArchive

'首页是否是打开状态，只读属性
Public Property Get IsOpen() As Boolean
    IsOpen = mclsArchive.IsOpen
End Property
'首页是否合并过,只读属性
Public Property Get IsMerged() As Boolean
    IsMerged = mclsArchive.IsMerged
End Property
'出院首要诊断是否发生改变,只读属性
Public Property Get IsDiagChange() As Boolean
    IsDiagChange = mclsArchive.IsDiagChange
End Property
'IsDiagInput:指定诊断的首页诊断是否输入
Public Property Get IsDiagInput() As Boolean
    IsDiagInput = mclsArchive.IsDiagInput
End Property

Public Sub InitArchiveMedRec(cnMain As ADODB.Connection, ByVal lngSys As Long, Optional ByVal lngModule As Long, Optional ByRef objMip As Object)
'功能：初始化
'参数：objMip=消息对象
    Call mclsArchive.InitArchiveMedRec(cnMain, lngSys, lngModule, objMip)
End Sub

Private Sub Class_Initialize()
    Set mclsArchive = New zlMedRecPage.clsArchive
End Sub

Private Sub Class_Terminate()
    If Not mclsArchive Is Nothing Then
        Set mclsArchive = Nothing
    End If
End Sub

Public Function zlGetForm(ByVal intType As Integer) As Object
'功能：获取子业务窗口,由主程序初始化时只调用一次获取
'参数：intType 0-门诊病案查阅 1-住院病案查阅
    Set zlGetForm = mclsArchive.zlGetForm(intType)
End Function

Public Function zlRefresh(ByVal intType As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal blnMoved As Boolean) As Boolean
'功能：刷新或清除医嘱清单
'参数：intType 0-门诊病案查阅 1-住院病案查阅
'      lng主页ID :住院首页-主页ID， 门诊首页-挂号ID
'      blnMoved 是否转储
    zlRefresh = mclsArchive.zlRefresh(intType, lng病人ID, lng主页ID, blnMoved)
End Function
