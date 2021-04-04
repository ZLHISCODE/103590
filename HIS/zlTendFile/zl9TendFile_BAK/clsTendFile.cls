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

Private Sub Class_Initialize()
    Set frmInTend = New frmDockInTendMain
End Sub

Private Sub Class_Terminate()
    Unload frmInTend
    Set frmInTend = Nothing
    
    Set mclsUnzip = Nothing
    Set mclsZip = Nothing
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
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal lngUnitID As Long, Optional ByVal intState As Integer) As Long
    '功能:调用刷新指定病人的病历内容，并根据情况提供编辑功能
    '参数:  lngPatiId-病人id;
    '       lngPageId-主页id;
    '       lngDeptId-当前操作部门，注意不是病人当前所在病区或出院病区；
    '       lngUnitID，lngDeptId-当intState=“5-最近转科病人”时为病人原病区或原科室
    '       blnEdit-是否允许编辑，通常当前操作部门不是病人所在病区或出院病区，就应该不允许编辑，医生站直接设置为不允许编辑
    '       intState-护士站或医生站传入的病人状态TYPE_PATI_State
    zlRefresh = frmInTend.zlRefresh(lngPatiID, lngPageId, lngDeptID, bnEdit, , blnDoctorStation)
End Function

Public Sub zlLocateData(ByVal intType As Integer)
'功能:定位到护理记录或才护理病历数据界面
'参数:intType=0-护理记录界面,1-护理病历界面
    Call frmInTend.zlLocateData(intType)
End Sub

Public Sub TendFileMutilEditor(ByVal objParent As Object, ByVal lngDeptID As Long, ByVal strPrivs As String)
    Dim frmTendFileMutil As New frmTendFileMutilEditor
    
    Call frmTendFileMutil.ShowMe(objParent, lngDeptID, strPrivs)
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

