VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoctorShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

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
    
    Call InitObjPublicAdvice
    If Err.Number <> 0 Then
        Set gobjEmr = Nothing
        Err.Clear
    End If
    
    Select Case lngModul
        Case 1242
            frmDoctorShift.Show , frmMain
    End Select
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub


Private Sub Class_Terminate()
    If Not gobjPublicAdvice Is Nothing Then Set gobjPublicAdvice = Nothing
End Sub


Public Sub InitDoctorShift(ByVal lngSys As Long, cnMain As ADODB.Connection)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    Set gcnOracle = cnMain
    glngSys = lngSys
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
End Sub

Public Sub ShowViewShift(frmMain As Object, str病人类型 As String)
    '预览交接班内容格式
    ' frmMain：父窗体对象
    ' str病人类型:默认查看病人类型的字符串，格式：新入,一级护理,XXX,XXX
    Dim frm As frmShiftEdit
    
    Set frm = New frmShiftEdit
    frm.gstr预览类型 = str病人类型
    frm.BorderStyle = FormBorderStyleConstants.vbSizable    '设置为无边框
    frm.Caption = "病人交接班内容预览"      '重点是这一句
    
    frm.Show 1, frmMain
End Sub
