VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDisReportCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mfrmReport As Object
Private mobjReportEx As Object


Public Sub Init(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long)
'功能：部件初始化

    On Error GoTo errHand
    Set gcnOracle = cnMain
    'Call InitCommon(gcnOracle)
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    Call GetUserInfo
    glngSys = lngSys
    
    '获取zlDisReportCardEx部件
    Call CreateCardExOK(1277)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub OpenDoc(ByVal frmParent As Object, ByVal bytType As Byte, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal bytFrom As Byte, ByVal bytBabyNo As Byte, ByVal lngDeptID As Long, ByVal lngFileId As Long, Optional ByVal blnHand As Boolean = False)
'功能：用于新增文档时调用，实现数据初始至编辑窗口打开
'参数说明:
'frmParent:编辑窗口用于显示的父窗口
'bytType:编辑方式0-新增　1-修改，用于区别提取数据
'lngPatiID:病人ID
'lngPageID:主页ID
'bytFrom:病人来源
'lngDeptID:当前科室
'lngFileId:bytType=1时，传入修改的文件ID,来源于电子病历记录.ID
'blnHand:是否为手动打开传染病报告卡
    On Error GoTo errHand
    Call frmMain.ShowMe(frmParent, bytType, lngPatiID, lngPageID, bytFrom, bytBabyNo, lngDeptID, lngFileId, blnHand)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub


Public Sub PrintDoc(ByVal frmParent As Object, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal strPrintDeviceName As String)
'功能:打印
'参数说明:
'frmParent:编辑窗口用于显示的父窗口
'lngPatiID:病人ID
'lngPageID:主页ID
'lngFileId:文件ID
'strPrintDeviceName:用于批量打印时不指定打印机，不指定时弹出窗口提供打印机选择。
    Call mfrmReport.mclsReport.PrintReport(frmParent, lngPatiID, lngPageID, lngFileId, strPrintDeviceName)
End Sub

Public Sub zlRefresh(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal blnMoved As Boolean)
'功能:刷新
'参数说明:
'lngPatiID:病人ID
'lngPageID:主页ID
'lngFileId:文件ID
'blnMoved:数据是否被转移
    On Error GoTo errHand
    Call mfrmReport.mclsReport.zlRefresh(lngPatiID, lngPageID, lngFileId, blnMoved)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Function zlGetForm() As Object
    If mfrmReport Is Nothing Then Set mfrmReport = GetReport
    If mfrmReport.hWnd = 0 Then Set mfrmReport = GetReport '外部UnLoad
    Set zlGetForm = mfrmReport
End Function

Public Sub SetCaption身份证()
    Call mfrmReport.mclsReport.SetCaption身份证
End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    '获取zlDisReportCardEx部件
    Call CreateCardExOK(1277)
    
    Set mfrmReport = GetReport
    mfrmReport.mclsReport.blnHaveStatus = False
    Err.Clear
End Sub

Private Sub Class_Terminate()
'    Set gobjCardEx = Nothing
    If Not mfrmReport Is Nothing Then
        Unload mfrmReport
        Set mfrmReport = Nothing
    End If
End Sub

