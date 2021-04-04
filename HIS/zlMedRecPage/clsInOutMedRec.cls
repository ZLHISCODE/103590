VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInOutMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' strTag=附加信息，现在存储门诊病人照片文件的路径，以后扩展时，以|分割
Public Event Closed(ByVal blnEditCancel As Boolean, ByVal str疾病ID As String, ByVal str诊断ID As String, ByVal strTag As String) '首页关闭事件

'Private mfrmInMedRecEdit As Object
Private mclsProperty As clsProperty                   '属性类，用来缓存部分变量，隔离数据
'属性类
Public Property Get Pros() As clsProperty
    Set Pros = mclsProperty
End Property
'首页是否是打开状态，只读属性
Public Property Get IsOpen() As Boolean
    IsOpen = mclsProperty.IsOpen
End Property
'首页是否合并过,只读属性
Public Property Get IsMerged() As Boolean
    IsMerged = mclsProperty.IsMerged
End Property
'出院首页诊断是否发生改变,只读属性
Public Property Get IsDiagChange() As Boolean
    IsDiagChange = mclsProperty.IsDiagChange
End Property
'IsDiagInput:指定诊断的首页诊断是否输入
Public Property Get IsDiagInput() As Boolean
    IsDiagInput = mclsProperty.IsDiagInput
End Property

Public Sub InitMedRec(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal lngModule As Long, Optional ByRef objMip As zl9ComLib.clsMipModule, Optional objCommunity As Object, Optional objInsure As clsInsure)
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = GetUnitName
    If lngModule = P病人入出管理 Then Exit Sub
    If Not (objMip Is Nothing) Then Set gclsMipModule = objMip
    Set mclsProperty = New clsProperty
    Call mclsProperty.InitProperty(f医生首页, lngSys, lngModule, objMip, objCommunity, objInsure)
End Sub

Public Function ShowInMedRecEdit(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng科室ID As Long, ByVal lngPathState As Long, _
                            Optional ByVal str类型 As String, Optional ByVal strPrivs As String, Optional ByVal intOpenMode As Integer, Optional ByVal blnModal As Boolean) As Boolean
'参数：
'      frmParent=调用窗体
'      lngModule=调用模块
'      lng病人ID=当前病人ID
'      lng主页ID=当前病人的主页ID
'      lng科室ID=当前病人的科室ID
'      lngPathState=当前病人的路径状态 -1=未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
'      str类型=要示录入的诊断类型，如"3,13"格式
'      strPrivs=当前用户权限
'      intOpenMode=首页打开模式
'      blnModal=窗体打开方式，True-模式打开窗体（窗体只有关闭才能操作其他窗体），False-无模打开窗体（窗体未关闭也可以操作其他窗体）
'返回：ShowInMedRecEdit=是否点击确定按钮关闭窗体
    If Not gclsMain Is Me Or gclsMain Is Nothing Then
        Set gclsMain = Me
    End If
    If Not gclsPros Is mclsProperty Or gclsPros Is Nothing Then
        Set gclsPros = mclsProperty
    End If
    mclsProperty.PatiType = PF_住院
    mclsProperty.MedPageSandard = Val(zlDatabase.GetPara("病案首页标准", mclsProperty.SysNo, p住院医生站, "0")) + 1
    '防止对象有时卸载不掉
    If Not mclsProperty.CurrentForm Is Nothing Then
        On Error Resume Next
        Unload mclsProperty.CurrentForm
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If
    mclsProperty.Privs = "" '自动读取
    mclsProperty.RefreshPara
    Set mclsProperty.MainForm = frmParent
    mclsProperty.OpenMode = intOpenMode
    mclsProperty.病人ID = lng病人ID
    mclsProperty.主页ID = lng主页ID
    mclsProperty.出院科室ID = lng科室ID
    mclsProperty.PathState = lngPathState
    mclsProperty.MustDiagType = str类型
    mclsProperty.InfosChange = False
    mclsProperty.LoadFinish = False
    
    ShowInMedRecEdit = frmMain.ShowMe(blnModal)
    '返回
    If blnModal Then
        ShowInMedRecEdit = mclsProperty.IsOK
    End If
End Function

Public Function ShowOutMedRecEdit(frmParent As Object, ByVal str挂号单 As String, ByVal strPrivs As String, Optional ByVal intOpenMode As Integer, Optional ByRef strPictureFile As String) As Boolean
'参数：blnDiagnose=是否调用用于填写诊断
'返回：blnDiagnose=是否填写了病人的诊断

    ShowOutMedRecEdit = False
End Function


Public Function PrintOrPriviewInMedRec(ByVal intType As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
        ByRef objReport As Object, ByVal lng科室ID As Long, ByRef objForm As Object, Optional intPage As Integer) As Boolean
'功能：首页打印，预览
'参数：intType=2（打印），=1（预览）0=设置
'     mobjReport-打印部件，lng科室ID-病人科室，mobjForm-主窗口
'     intPage=1-4打印的页数（格式）=5打印正面+附页1，=6打印反面+附页2
    Call PrintInMedRec(intType, lng病人ID, lng主页ID, lng科室ID, intPage, objReport, objForm)
    PrintOrPriviewInMedRec = True
End Function
'
Public Sub Closed(ByVal blnEditCancel As Boolean, ByVal str疾病ID As String, ByVal str诊断ID As String, ByVal strTag As String)
'功能：调用事件
' strTag=附加信息，现在存储门诊病人照片文件的路径，以后扩展时，以|分割
    RaiseEvent Closed(blnEditCancel, str疾病ID, str诊断ID, strTag)
End Sub

Public Sub EditDelivery(ByRef objFrmMain As Object, ByVal lngPatiID As Long, ByVal lngMainID As Long)
'功能:允许新生儿登记时登记分娩信息
    Call CmdDeliceryInfoClick(1, objFrmMain, lngPatiID, lngMainID)
End Sub

Public Sub Hide()
    On Error Resume Next
    Call frmMain.Hide
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Public Function FormUnLoad() As Boolean
    On Error Resume Next
    Unload frmMain
    If Not gblnUnload Then
        FormUnLoad = True
    Else
        FormUnLoad = False
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Public Sub PlugWndProc(msg As Long, wParam As Long, lParam As Long, Result As Long)
    '外挂附页自定义的消息处理函数
    Dim wzDelta As Integer
    Select Case msg
        Case WM_MOUSEWHEEL   '滚动
            wzDelta = HIWORD(wParam)
            With gclsPros.CurrentForm
                If wzDelta > 0 Then        '向上滚动
                    Call ChangePage(False, , , False)
                Else                        '向下滚动
                    Call ChangePage(True, , , False)
                End If
            End With
    End Select
End Sub

Private Sub Class_Terminate()
    Set mclsProperty = Nothing
End Sub
