VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExternalCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long, ByVal int场合 As Integer)
'功能：接口初始化
'参数：cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    Set gcnOracle = cnOracle
End Sub


Public Function ExternalCheck(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int场合 As Integer, ByVal lngID As Long, ByVal strName As String, ByRef strXml As String, ByRef strVerificationCode As String, ByRef blnGranted As Boolean) As Boolean
'功能：三方配置授权检查
'入参： lngSys,lngModual=当前调用接口的主程序系统号及模块号
'       int场合：调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'       lngID：三方配置功能的菜单ID(三方调用目录.ID)
'       strName：三方配置功能的菜单名称(三方调用目录.菜单名称)
'       strXml:扩展参数

'出参： strVerificationCode：返回医院名称，调用时检查是否与当前医院名称一致，可用于渠道一个部件统一分不同用户授权
'       blnGranted：授权是正确

End Function
