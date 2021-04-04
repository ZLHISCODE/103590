VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPageMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'调用病案存放位置,点击确定时调用
Public Event MedRecSaveLocation(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
'保存首页，刷新病人列表
Public Event SavePage(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
Private mfrmPageMedRecEdit As Object
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

Public Sub InitPageMedRec(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal lngModule As Long, Optional ByRef objMip As zl9ComLib.clsMipModule)
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = GetUnitName
    If Not (objMip Is Nothing) Then Set gclsMipModule = objMip
    Set mclsProperty = New clsProperty
    Call mclsProperty.InitProperty(f病案首页, lngSys, lngModule, objMip)
End Sub

Public Function ShowPageMedRecEdit(frmParent As Object, ByVal str住院号 As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal intOpenMode As Integer, Optional bln编目 As Boolean, Optional blnChange编码员 As Boolean) As Boolean
'参数：lngModel=调用模块
'      frmParent=调用窗体
'      lng病人ID=当前病人ID
'      lng主页ID=当前病人的主页ID
'      str住院号=当前病人的住院号
'      bln编目=对已有的主页编码
'      intOpenMode=首页打开模式
'      blnChange编码员=是否可以更换编码员
'返回：ShowInMedRecEdit=是否点击确定按钮关闭窗体
    If Not gclsMain Is Me Or gclsMain Is Nothing Then
        Set gclsMain = Me
    End If
    If Not gclsPros Is mclsProperty Or gclsPros Is Nothing Then
        Set gclsPros = mclsProperty
    End If
    mclsProperty.PatiType = PF_住院
    If Not mclsProperty.ShareMedRec Then
        mclsProperty.MedPageSandard = Val(zlDatabase.GetPara("病案首页标准", mclsProperty.SysNo, p病案管理, "0")) + 1
    Else
        mclsProperty.MedPageSandard = Val(zlDatabase.GetPara("病案首页标准", 100, p住院医生站, "0")) + 1
    End If
    If gblnUnload Then
        Unload frmMain
    Else
        If Not mclsProperty.CurrentForm Is Nothing Then
            '82639:问题产生原因未知，具体跟踪发现，某次关闭窗体，该函数执行结束，调用Form_Load数据，导致窗体未卸载掉，再次show时
            '直接跳过Form_load，显示出上一个病人的信息，此处应对策略是，强制卸载窗体
            On Error Resume Next
            Unload mclsProperty.CurrentForm
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    End If
    
    mclsProperty.Privs = "" '自动读取
    mclsProperty.RefreshPara
    Set mclsProperty.MainForm = frmParent
    mclsProperty.OpenMode = intOpenMode
    If intOpenMode = EM_编辑 Then
        mclsProperty.IsExistPati = True
    End If
    mclsProperty.病人ID = lng病人ID
    mclsProperty.主页ID = lng主页ID
    mclsProperty.InNo = str住院号
    mclsProperty.Is编目 = bln编目
    mclsProperty.Change编码员 = blnChange编码员
    ShowPageMedRecEdit = frmMain.ShowMe(False)
End Function

Private Sub Class_Terminate()
    Set mclsProperty = Nothing
    If Not mfrmPageMedRecEdit Is Nothing Then
        Unload mfrmPageMedRecEdit
        Set mfrmPageMedRecEdit = Nothing
    End If
End Sub

Public Sub MedRecSaveLocation(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    RaiseEvent MedRecSaveLocation(lng病人ID, lng主页ID)
End Sub

Public Sub SavePage(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    RaiseEvent SavePage(lng病人ID, lng主页ID)
End Sub
Public Function unLoadPage() As Boolean
    Unload frmMain
    If Not gblnUnload Then
        unLoadPage = True
    Else
        unLoadPage = False
    End If
End Function
