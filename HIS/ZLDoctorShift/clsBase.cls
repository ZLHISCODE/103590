VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    Set gcnOracle = cnMain
    glngSys = lngSys
    gstrDbaUser = strDbUser
    glngModul = lngModul
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    On Error Resume Next
    Set gobjEmr = frmMain.mobjEMR  '新病历需要部件
    If Err.Number <> 0 Then
        Set gobjEmr = Nothing
        Err.Clear
    End If
    
    Select Case lngModul
        Case 1242
            frmDoctorShift.Show , frmMain
    End Select
End Sub
