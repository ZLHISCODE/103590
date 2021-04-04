VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'类的模块变量
Private mfrmInMedRec As Object
Private mfrmOutMedRec As frmArchiveOutMedRec
Private mstZY As MedPage_Standard                '住院首页标准,防止首页标准变更带来不能展示数据的问题
Private mclsProperty As clsProperty                   '属性类，用来缓存部分变量，隔离数据
'首页是否是打开状态，只读属性
Public Property Get IsOpen() As Boolean
    IsOpen = mclsProperty.IsOpen
End Property
'首页是否合并过,只读属性
Public Property Get IsMerged() As Boolean
    IsMerged = mclsProperty.IsMerged
End Property
'出院首要诊断是否发生改变,只读属性
Public Property Get IsDiagChange() As Boolean
    IsDiagChange = mclsProperty.IsDiagChange
End Property
'IsDiagInput:指定诊断的首页诊断是否输入
Public Property Get IsDiagInput() As Boolean
    IsDiagInput = mclsProperty.IsDiagInput
End Property

Public Sub InitArchiveMedRec(cnMain As ADODB.Connection, ByVal lngSys As Long, Optional ByVal lngModule As Long, Optional ByRef objMip As zl9ComLib.clsMipModule)
'功能：初始化
'参数：objMip=消息对象
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = GetUnitName
    If Not (objMip Is Nothing) Then Set gclsMipModule = objMip
    Set mclsProperty = New clsProperty
    Call mclsProperty.InitProperty(f电子病案, lngSys, lngModule, objMip)
End Sub

Private Sub Class_Terminate()
    If Not mfrmInMedRec Is Nothing Then
        Unload mfrmInMedRec
        Set mfrmInMedRec = Nothing
    End If
    If Not mfrmOutMedRec Is Nothing Then
        Unload mfrmOutMedRec
        Set mfrmOutMedRec = Nothing
    End If
    Set mclsProperty = Nothing
End Sub

Public Function zlGetForm(ByVal intType As Integer) As Object
'功能：获取子业务窗口,由主程序初始化时只调用一次获取
'参数：intType 0-门诊病案查阅 1-住院病案查阅
    If Not gclsMain Is Me Or gclsMain Is Nothing Then
        Set gclsMain = Me
    End If
    If Not gclsPros Is mclsProperty Or gclsPros Is Nothing Then
        Set gclsPros = mclsProperty
    End If
    If intType + 1 = PF_门诊 Then
        mclsProperty.PatiType = PF_门诊
        mclsProperty.MedPageSandard = ST_门诊首页
        If mfrmOutMedRec Is Nothing Then
            Set mfrmOutMedRec = New frmArchiveOutMedRec
        End If
        Set zlGetForm = mfrmOutMedRec
        Set mclsProperty.CurrentForm = mfrmOutMedRec
    Else
        mclsProperty.PatiType = PF_住院
        mclsProperty.MedPageSandard = Val(zlDatabase.GetPara("病案首页标准", mclsProperty.SysNo, p住院医生站, "0")) + 1
        If mfrmInMedRec Is Nothing Then
            mstZY = mclsProperty.MedPageSandard
            '病案首页标准
            Select Case mstZY
                Case ST_卫生部标准  '卫生部标准
                    Set mfrmInMedRec = New frmArchiveInMedRec
                Case ST_四川省标准   '四川省标准
                    Set mfrmInMedRec = New frmArchiveInMedRec_SC
                Case ST_云南省标准    '云南省标准
                    Set mfrmInMedRec = New frmArchiveInMedRec_YN
                Case ST_湖南省标准    '湖南省标准
                    Set mfrmInMedRec = New frmArchiveInMedRec_HN
            End Select
        End If
        Set zlGetForm = mfrmInMedRec
        Set mclsProperty.CurrentForm = mfrmInMedRec
    End If
End Function

Public Function zlRefresh(ByVal intType As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal blnMoved As Boolean) As Boolean
'功能：刷新或清除医嘱清单
'参数：intType 0-门诊病案查阅 1-住院病案查阅
'      lng主页ID :住院首页-主页ID， 门诊首页-挂号ID
'      blnMoved 是否转储
    If intType + 1 = PF_门诊 Then
        mclsProperty.PatiType = PF_门诊
        mclsProperty.MedPageSandard = ST_门诊首页
        Set mclsProperty.CurrentForm = mfrmOutMedRec
    Else
        mclsProperty.PatiType = PF_住院
        mclsProperty.MedPageSandard = mstZY
        Set mclsProperty.CurrentForm = mfrmInMedRec
    End If
    mclsProperty.病人ID = lng病人ID
    mclsProperty.主页ID = lng主页ID
    Set gclsPros = mclsProperty
    Call ArchivezlRefresh
End Function

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function









