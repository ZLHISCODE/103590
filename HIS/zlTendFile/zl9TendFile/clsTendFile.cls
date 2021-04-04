VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'继承窗体事件
Public Event Activate()
Public Event RefreshPrompt(ByVal strInfo As String, ByVal blnImportant As Boolean)

Private WithEvents frmInTend As frmDockInTendMain
Attribute frmInTend.VB_VarHelpID = -1
Private WithEvents frmInTendFile As frmDockInTend_File
Attribute frmInTendFile.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set frmInTend = New frmDockInTendMain
    Set frmInTendFile = New frmDockInTend_File
End Sub

Private Sub Class_Terminate()
    If Not frmInTend Is Nothing Then
        Unload frmInTend
        Set frmInTend = Nothing
    End If
    If Not frmInTendFile Is Nothing Then
        Unload frmInTendFile
        Set frmInTendFile = Nothing
    End If
    
    If Not gobjBodyEditor Is Nothing Then Set gobjBodyEditor = Nothing
    If Not gobjPartogram Is Nothing Then Set gobjPartogram = Nothing
    If Not mclsUnzip Is Nothing Then Set mclsUnzip = Nothing
    If Not mclsZip Is Nothing Then Set mclsZip = Nothing
End Sub

Private Sub frmInTend_Activate()
    RaiseEvent Activate
End Sub

Private Sub frmInTend_RefreshPrompt(ByVal strInfo As String, ByVal blnImportant As Boolean)
    RaiseEvent RefreshPrompt(strInfo, blnImportant)
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = frmInTend
End Function

Public Function zlGetfrmInTendFile() As Object
    Set zlGetfrmInTendFile = frmInTendFile
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object, Optional ByVal blnChildToolBar As Boolean = False)
    Call frmInTend.zlDefCommandBars(CommandBars, blnChildToolBar)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInTend.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInTend.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal bnEdit As Boolean, _
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal lngUnitId As Long, Optional ByVal intState As Integer, Optional ByVal intCurveReSize As Integer = 0) As Long
    '功能:调用刷新指定病人的病历内容，并根据情况提供编辑功能
    '参数:  lngPatiId-病人id;
    '       lngPageId-主页id;
    '       lngDeptId-当前操作部门，注意不是病人当前所在病区或出院病区；
    '       lngUnitID，lngDeptId-当intState=“5-最近转科病人”时为病人原病区或原科室
    '       blnEdit-是否允许编辑，通常当前操作部门不是病人所在病区或出院病区，就应该不允许编辑，医生站直接设置为不允许编辑
    '       intState-护士站或医生站传入的病人状态TYPE_PATI_State
    '       intCurveReSize 体温单查阅是是否为缩放模式 0缩放模式 1原始大小
    '56965:刘鹏飞,2012-12-24
    glng病区ID = IIf(blnDoctorStation = True, lngUnitId, lngDeptID)
    zlRefresh = frmInTend.zlRefresh(lngPatiID, lngPageId, glng病区ID, bnEdit, , blnDoctorStation, , intCurveReSize)
End Function

Public Function zlRefreshTendFile(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, ByVal lngDeptID As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal intSEL As Integer, Optional ByVal lngKey As Long, Optional ByVal intCurveReSize As Integer = 0) As Long
    'bytSel:0-体温单;1-记录单;2 产程图
    Call frmInTendFile.zlRefresh(lngPatiID, lngPageId, intBaby, lngDeptID, blnEdit, blnDoctorStation, intSEL, lngKey, intCurveReSize)
End Function

Public Function zlPrintTendFile(ByVal bytKind As Byte, Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDevice As String = "") As Long
    '完成记录单、体温单、产程图的预览打印
    '参数:bytkind  1-体温单 2-记录单  3-产程图
    '     bytMode 1 预览 2打印 3输出到Excel
    '     strPrintDevice 打印机名称，不为空时打印病人所有内容，为空时弹出询问窗体
    Call frmInTendFile.zlPrintTendFile(bytKind, bytMode, strPrintDevice)
End Function

Public Sub zlPrintDocument(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, ByVal lngDeptID As Long, ByVal lngFileID As Long, ByVal bytKind As Byte, ByVal strPrintDeviceName As String, Optional ByVal blnPrint As Boolean = True)
    '功能:完成记录单、体温单、产程图输出(可直接使用)
    '参数:lngPatiID-病人ID
    '     lngPageId-主页ID
    '     intBaby-婴儿序号
    '     lngDeptID-科室ID
    '     lngFileID-文件ID
    '     bytKind:1-体温单:2-记录单 3-产程图
    '     strPrintDeviceName:打印机名称
    '     blnPrint:TRUE 打印，False预览
    Call frmInTendFile.zlPrintDocument(lngPatiID, lngPageId, intBaby, lngDeptID, lngFileID, bytKind, strPrintDeviceName, blnPrint)
End Sub

Public Sub zlLocateData(ByVal intType As Integer)
'功能:定位到护理记录或才护理病历数据界面
'参数:intType=0-护理记录界面,1-护理病历界面
    Call frmInTend.zlLocateData(intType)
End Sub

Public Sub TendFileMutilEditor(ByVal objParent As Object, ByVal lngDeptID As Long, ByVal strPrivs As String, Optional ByVal bytSize As Byte = 0)
    Dim frmTendFileMutil As New frmTendFileMutilEditor
    
    Call frmTendFileMutil.ShowMe(objParent, lngDeptID, strPrivs, bytSize)
End Sub

Public Sub InitTendFile(cnMain As ADODB.Connection, ByVal lngSys As Long)
    Set gcnOracle = cnMain
    Call InitCommon(gcnOracle)
        
    glngSys = lngSys
    gstrDbOwner = GetDbOwner(glngSys)
    gstrProductName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("产品名称"), Default:="中联")
    gstrAviPath = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    
    Call GetUserInfo
End Sub

'Public Function zlTendPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
'    '完成记录单参数设置
'    '参数:frmMain  调用窗体
'    '     strPrivs 权限
'    zlTendPara = frmTendPara.ShowPara(frmMain, strPrivs)
'End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '区别在于是否提供文件信息列表
    '编制:刘鹏飞
    '日期:2012-06-19 15:16
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If frmInTend Is Nothing Then Exit Sub
    Call frmInTend.SetFontSize(bytSize)
End Sub

Public Sub UnloadClass()
'新版电子病历使用，因新版电子病历无法在卸载对象时调用Class_Terminate
    Call Class_Terminate
End Sub
