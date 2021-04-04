VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTechCore"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Public Sub InitTechCore(cnMain As ADODB.Connection)
'------------------------------------------------
'功能： 部件的初始化，将主程序的必要内容传递下来
'参数：
'   cnMain:主程序的数据库连接
'返回：
'------------------------------------------------
    Set gcnOracle = cnMain
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    glngSys = 100
    Call GetUserInfo
    
    gstrUnitName = GetUnitName
    Call gclsInsure.InitOracle(gcnOracle)
    Call InitCISCore(gcnOracle)
    
    If Not gblnPass Then
        gblnPass = GetSysParVal(30) = "1"
        If gblnPass Then
            If Not PassInitialize() Then gblnPass = False
        End If
    End If
End Sub

Private Sub Class_Terminate()
    'PASS接口终止
    If gblnPass Then Call PassQuit
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub


'外部接口：医嘱编辑
'用于手麻等外围系统
'By:赵彤宇
Public Function Edit医嘱(frmMain As Object, _
                            ByVal PatientID As Long, _
                            ByVal PageID As Variant, _
                            ByVal lng医嘱ID As Long, _
                            cnMain As ADODB.Connection, _
                            ByVal strDBUser As String, _
                            ByVal lngSys As Long, _
                            Optional ByVal strPrivs As String = "") As Boolean
    
    '---------------------------------------------------------------------------------------------------------
    '功能:
    '---------------------------------------------------------------------------------------------------------
    If TypeName(PageID) = "String" Then
        If Not CheckAdviceWindow("门诊医嘱编辑") Then Exit Function
    Else
        If Not CheckAdviceWindow("住院医嘱编辑") Then Exit Function
    End If
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")

    gstrDBUser = strDBUser

    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If

    Set gfrmMain = frmMain
    glngSys = lngSys

    gstrUnitName = GetUnitName
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call InitCISCore(gcnOracle)
    
    If TypeName(PageID) = "String" Then
        Edit医嘱 = frmOutAdviceEdit.ShowMe(frmMain, strPrivs, PatientID, PageID, lng医嘱ID)
    Else
        Edit医嘱 = frmInAdviceEdit.ShowMe(frmMain, strPrivs, PatientID, PageID, lng医嘱ID)
    End If
End Function

'外部接口：在指定容器中显示医嘱列表
'用于手麻等外围系统
'By:赵彤宇
Public Function ListAdviceInObject(frmMain As Object, objContainer As Object, _
                            ByVal PatientID As Long, _
                            ByVal PageID As Variant, _
                            ByVal lng病区ID As Long, ByVal lng科室ID As Long, _
                            ByVal lng医嘱ID As Long, _
                            cnMain As ADODB.Connection, _
                            ByVal strDBUser As String, _
                            ByVal lngSys As Long, _
                            Optional ByVal strPrivs As String = "") As Object

    '---------------------------------------------------------------------------------------------------------
    '功能:
    '---------------------------------------------------------------------------------------------------------
    Dim lngStyle As Long

    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")

    gstrDBUser = strDBUser

    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If

    Set gfrmMain = frmMain
    glngSys = lngSys

    gstrUnitName = GetUnitName
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call InitCISCore(gcnOracle)
    
    If TypeName(PageID) = "String" Then
        Set ListAdviceInObject = New frmOutDoctorAdvice
    Else
        Set ListAdviceInObject = New frmInDoctorAdvice
    End If
    ListAdviceInObject.mstrPrivs = strPrivs
    Set ListAdviceInObject.mfrmParent = frmMain
    
    Load ListAdviceInObject
    
    lngStyle = GetWindowLong(ListAdviceInObject.Hwnd, GWL_STYLE)
    SetWindowLong ListAdviceInObject.Hwnd, GWL_STYLE, lngStyle Or WS_CHILD
    
    SetParent ListAdviceInObject.Hwnd, objContainer.Hwnd
    MoveWindow ListAdviceInObject.Hwnd, 0, 0, objContainer.ScaleWidth / Screen.TwipsPerPixelX, objContainer.ScaleHeight / Screen.TwipsPerPixelY, 1
            
    '调用功能窗口数据刷新
    If TypeName(PageID) = "String" Then
        Call ListAdviceInObject.zlRefresh(PatientID, CStr(PageID), 1, 0)
    Else
        Call ListAdviceInObject.zlRefresh(PatientID, CLng(PageID), lng病区ID, lng科室ID, False)
    End If

    ListAdviceInObject.Show
End Function

Public Function ListChargeInObject(ByVal frmMain As Object, _
                                    ByVal objContainer As Object, _
                                    ByVal lng医嘱ID As Long, _
                                    ByVal lng发送id As Long, _
                                    ByVal cnMain As ADODB.Connection, _
                                    ByVal strDBUser As String, _
                                    ByVal lngSys As Long, _
                                    Optional ByVal strPrivs As String = "", _
                                    Optional ByVal strClass As String = "检验", _
                                    Optional ByVal strSys As String = "") As Object
    Dim lngStyle As Long

    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")

    gstrDBUser = strDBUser

    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If

    Set gfrmMain = frmMain
    glngSys = lngSys

    gstrUnitName = GetUnitName
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call InitCISCore(gcnOracle)
        
    Set ListChargeInObject = New frmChargeRequest
    
    Call InitSysPar '初始化系统参数
    
    Load ListChargeInObject
    
    lngStyle = GetWindowLong(ListChargeInObject.Hwnd, GWL_STYLE)
    
    Call SetWindowLong(ListChargeInObject.Hwnd, GWL_STYLE, lngStyle Or WS_CHILD)
    Call SetParent(ListChargeInObject.Hwnd, objContainer.Hwnd)
    Call MoveWindow(ListChargeInObject.Hwnd, 0, 0, objContainer.ScaleWidth / Screen.TwipsPerPixelX, objContainer.ScaleHeight / Screen.TwipsPerPixelY, 1)
            
    '调用功能窗口数据刷新
    
    Call ListChargeInObject.zlRefresh(frmMain, lng医嘱ID, lng发送id, strPrivs, strClass, strSys)
    
    ListChargeInObject.Show
End Function

Public Function InDoctorAdvice() As Object
    Set InDoctorAdvice = New frmInDoctorAdvice
End Function

Public Function OutDoctorAdvice() As Object
    Set OutDoctorAdvice = New frmOutDoctorAdvice
End Function
'编辑费用
Public Function EditExpense(objParent As Object, ByVal mbytInState As Integer, ByVal int记录性质 As Integer, ByVal mstrPrivs As String, ByVal strNO As String, _
    ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal int病人来源 As Integer, ByVal lng开单科室ID As Long, ByVal lng病人科室ID As Long, Optional ByVal mbln费用登记 As Boolean = False) As Boolean
    
    Dim bln零耗 As Boolean
    
    Call InitSysPar '初始化系统参数
    
    If int记录性质 = 2 And mbytInState = 0 And Len(strNO) > 0 Then
        bln零耗 = BillisZeroLog(strNO)
    Else
        bln零耗 = mbln费用登记
    End If

    frmTechnicExpense.mstrPrivs = mstrPrivs
    frmTechnicExpense.mbytInState = mbytInState
    frmTechnicExpense.mbln费用登记 = bln零耗
    frmTechnicExpense.mstrInNO = strNO
    frmTechnicExpense.mlng医嘱ID = lng医嘱ID
    frmTechnicExpense.mlng发送号 = lng发送号
    frmTechnicExpense.mlng病人ID = lng病人ID
    frmTechnicExpense.mlng主页ID = lng主页ID
    frmTechnicExpense.mint病人来源 = int病人来源
    frmTechnicExpense.mint记录性质 = int记录性质
    frmTechnicExpense.mlng开单科室ID = lng开单科室ID
    frmTechnicExpense.mlng病人科室id = lng病人科室ID
    On Error Resume Next
    frmTechnicExpense.Show 1, objParent
    On Error GoTo 0
    EditExpense = gblnOK
End Function
