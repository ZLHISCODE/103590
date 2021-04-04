VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsWin"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InitBrower(ByRef objRelogin As Object, ByRef cnOracle As ADODB.Connection, ByRef rsMenu As ADODB.Recordset)
    On Error Resume Next
    Set gobjRelogin = objRelogin
    Set gcnOracle = cnOracle
    Set grsMenus = rsMenu.Clone
    Set FrmMainface = FrmWin
    '在启动程序中已初始化公共部件的连接
    FrmMainface.Show
    
End Sub

Public Sub SetEnvironment(strSysNameIn As String, strVersionIn As String, strAviPathIn As String, _
                          strUserFlagIn As String, strDbUserIn As String, lngUserIdIn As Long, _
                          strUserCodeIn As String, strUserNameIn As String, strUserAbbrIn As String, _
                          lngDeptIdIn As Long, strDeptCodeIn As String, strDeptNameIn As String, _
                          strStationIn As String, strMenusysIn As String, Optional strCommand As String)
    '设置环境变量
    gstrSysName = strSysNameIn
    gstrVersion = strVersionIn
    gstrAviPath = strAviPathIn
    gstrUserFlag = strUserFlagIn
    gstrDbUser = strDbUserIn
    glngUserId = lngUserIdIn
    gstrUserCode = strUserCodeIn
    gstrUserName = strUserNameIn
    gstrUserAbbr = strUserAbbrIn
    glngDeptId = lngDeptIdIn
    gstrDeptCode = strDeptCodeIn
    gstrDeptName = strDeptNameIn
    gstrStation = strStationIn
    gstrMenuSys = strMenusysIn
    gstrCommand = strCommand
End Sub

Private Sub Class_Terminate()
    Set gobjPlugIn = Nothing
End Sub

