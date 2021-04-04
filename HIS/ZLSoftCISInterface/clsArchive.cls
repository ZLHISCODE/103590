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
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'接口功能说明:
'
'1、RIS查看电子病历
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'接口调用顺序说明
'1、支持两种方式查看电子病历：嵌入式窗口和独立窗口
'2、独立窗口调用方法：（1）初始化zlInitCommon （2）电子病案查阅zlSoftCISInterface
'3、嵌入式窗口调用方法：（1）初始化zlInitCommon（2）创建窗口，提取窗口句柄zlGetFormHwnd（3）刷新窗口zlRefresh

Private mfrmDockArchive As Object ' 是一个无边框窗体，可绑定到其它窗体上‘此功能暂时只支持35及之后的版本
Private mfrmArchive As Object '单人病人调用电子病案查阅窗体对象

Private mobjArchive As Object '类 zlPublicAdvice.clsPublicAdvice
Private mfrmArchiveView As frmArchiveView   '电子病案查阅窗体

'******************************************************************************************************************
'功能：打开电子病案查阅窗口
'参数：
'       lngPatientID - 病人ID
'       lngClinicID - 就诊ID，门诊为挂号ID 病人挂号记录.ID，住院为主页ID
'返回：True - 成功；False - 失败
'说明：可以直接手工关闭窗口，不需要调用zlCloseArchiveForm。手工关闭的窗口会被隐藏
'******************************************************************************************************************
Public Function zlOpenArchiveForm(ByVal lngPatientID As Long, ByVal lngClinicID As Long, Optional ByVal blnLoad As Boolean = False)
    Dim lngLeft As Long
    Dim lngTop As Long
    
    
    
    '如果窗口已经打开，则重新显示，不再创建
    If mfrmArchiveView Is Nothing Then
        '第一次调用电子病案处查询，赋值，并创建窗体
        Set mfrmArchiveView = New frmArchiveView
        Call mfrmArchiveView.ShowMe(lngPatientID, lngClinicID, Not blnLoad)
        mfrmArchiveView.WindowState = vbNormal
    End If
    
    '电子病案查询窗体已经存在，将它显示出来，并处理消息
    Call mfrmArchiveView.zlRefresh(lngPatientID, lngClinicID)
    
    If blnLoad Then
        zlOpenArchiveForm = True
        Exit Function
    End If
    Call mfrmArchiveView.Show
    mfrmArchiveView.WindowState = 0
    
    '把窗口展示到最前端，有部分机器窗口无法显示到最前端，因此用API强制显示
    If (Screen.Width > mfrmArchiveView.Width) And (Screen.Height > mfrmArchiveView.Height) Then
        lngLeft = (Screen.Width - mfrmArchiveView.Width) / 2 / Screen.TwipsPerPixelX
        lngTop = (Screen.Height - mfrmArchiveView.Height) / 2 / Screen.TwipsPerPixelY
    End If
    SetWindowPos mfrmArchiveView.hWnd, HWND_TOPMOST, lngLeft, lngTop, 0, 0, &H10 Or &H1
    SetWindowPos mfrmArchiveView.hWnd, HWND_NOTOPMOST, lngLeft, lngTop, 0, 0, &H10 Or &H1
    BringWindowToTop mfrmArchiveView.hWnd
    SetForegroundWindow mfrmArchiveView.hWnd
    
    
    zlOpenArchiveForm = True
    
End Function

'******************************************************************************************************************
'功能：关闭电子病案查阅窗口
'参数：无
'返回：True - 成功；False - 失败
'说明：
'******************************************************************************************************************
Public Function zlCloseArchiveForm()
    
    '如果窗口不存在，则直接返回成功
    If mfrmArchiveView Is Nothing Then
        zlCloseArchiveForm = True
        Exit Function
    End If
    
    '关闭窗口
    mfrmArchiveView.zlCloseMe
    
    Set mobjArchive = Nothing
    Set mfrmArchiveView = Nothing
    
    zlCloseArchiveForm = True
    
End Function

'******************************************************************************************************************
'功能： 初始化部件
'参数： 无
'返回： True，成功；False，失败
'说明： 窗体初始化或对象初始化时，调用:zlInitCommon,一般只调用一次。
'******************************************************************************************************************
Public Function zlInitCommon() As Boolean
    Dim objEmr As Object
    
    If CreateArchive = False Then Exit Function
    
    Set objEmr = UpdateEmrInterface
    
    If mobjArchive.InitCommon(gcnOracle, glngSys, , , , , , objEmr) = False Then Exit Function
    zlInitCommon = True
    
 
End Function

'******************************************************************************************************************
'功能： 获取嵌入式病案查阅窗口句柄
'参数： 无
'返回： 窗口句柄，0失败
'说明：
'******************************************************************************************************************
Public Function zlGetFormHwnd() As Long
    On Error GoTo err

    If CreateArchive = False Then Exit Function
    
    zlGetFormHwnd = zlGetForm.hWnd
    
    Exit Function
err:
    MsgBox err.Description
   
End Function

'******************************************************************************************************************
'功能： 刷新嵌入式病案查阅窗口
'参数： lngPatientID - 病人ID
'       lngClinicID - 就诊ID，门诊为挂号ID 病人挂号记录.ID，住院为主页ID
'返回： 0 成功；非0，失败
'说明：
'******************************************************************************************************************
Public Function zlRefresh(ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Long
    
    zlRefresh = 1
    
    
    If CreateArchive = False Then Exit Function
    Call mfrmDockArchive.zlRefresh(lngPatientID, lngClinicID)
    zlRefresh = 0
    
End Function


'******************************************************************************************************************
'功能： 独立窗口的电子病案查阅
'参数： lngPatientID - 病人ID
'       lngClinicID - 就诊ID，门诊为挂号ID 病人挂号记录.ID，住院为主页ID
'返回： 0 成功；非0，失败
'说明：
'******************************************************************************************************************
Public Function zlShowArchive(ByVal lngPatientID As Long, ByVal lngClinicID As Long) As Long
'功能：电子病案查阅  frmParent 可以传 Nothing
    
    Dim frmParent As Object
    Set frmParent = Nothing
    
    zlShowArchive = 1
    

    If CreateArchive = False Then Exit Function
    Call mobjArchive.ShowArchive(frmParent, lngPatientID, lngClinicID)
    '获取当前显示的窗体对象，用于卸载
    If mfrmArchive Is Nothing Then
        Set mfrmArchive = mobjArchive.GetArchiveOneFrom
    End If
    zlShowArchive = 0
   
End Function

Private Function CreateArchive() As Boolean
'功能：创建对象
    err = 0: On Error Resume Next
    If Not mobjArchive Is Nothing Then CreateArchive = True: Exit Function
    Set mobjArchive = CreateObject("zlPublicAdvice.clsPublicAdvice")
    If err <> 0 Then
        MsgBox "未找到zlPublicAdvice部件,请检查该站点是部署了此部件!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    CreateArchive = False
End Function

Private Function zlGetForm() As Object
'功能：电子病案查阅无框窗体
    If CreateArchive = False Then Exit Function
    If mfrmDockArchive Is Nothing Then
        Set mfrmDockArchive = mobjArchive.GetArchiveFrom
    End If
    Set zlGetForm = mfrmDockArchive
End Function

Private Sub Class_Initialize()
    Call CreateArchive
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Unload mfrmDockArchive
    Unload mfrmArchive
    Set mfrmDockArchive = Nothing
    Set mfrmArchive = Nothing
    Set mobjArchive = Nothing
End Sub



'****************临床查看报告
Public Sub zlViewEPRReport(frmParent As Object, ByVal lng报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告  frmParent 可以传 Nothing
    If CreateArchive = False Then Exit Sub
    Call mobjArchive.ViewEPRReport(frmParent, lng报告ID, CanPrint)
End Sub

'****************临床查看PACS图像
Public Sub zlViewPACSImage(frmParent As Object, ByVal lng医嘱id As Long, ByVal blnMoved As Boolean)
'功能：PACS观片处理 frmParent 可以传 Nothing
    If CreateArchive = False Then Exit Sub
    Call mobjArchive.ViewPACSImage(frmParent, lng医嘱id, blnMoved)
End Sub
