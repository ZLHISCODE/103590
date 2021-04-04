VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockDiagReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'说明：用于挂接辅诊子窗口,如果要取消,删除以下内容：
'1.去掉工程引用：zl9CISCore,zl9PrintMode,Microsoft Scripting Runtime
'2.移除工程文件：frmDockDiagReport,frmPubIcons,clsDockDiagReport,clsFtp
'3.删除几个工作站中mclsOldDiag,"旧辅诊"相关代码

Public Event Activate() '子窗体激活
Public Event RequestRefresh(ByVal RefreshNotify As Boolean) '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字
Private WithEvents mfrmDiagReport As frmDockDiagReport
Attribute mfrmDiagReport.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mfrmDiagReport = New frmDockDiagReport
End Sub

Private Sub Class_Terminate()
    Set mfrmDiagReport = Nothing
End Sub

Private Sub mfrmDiagReport_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmDiagReport_RequestRefresh(ByVal RefreshNotify As Boolean)
    RaiseEvent RequestRefresh(RefreshNotify)
End Sub

Private Sub mfrmDiagReport_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Public Function zlGetForm() As Object
'功能：获取子业务窗口,由主程序初始化时只调用一次获取

    '初始读取权限,以加快后续速度(不是必要的)
    Call GetInsidePrivs(p辅诊记录管理, True)
    
    Set zlGetForm = mfrmDiagReport
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int场合 As Integer, Optional ByVal bln护士站 As Boolean)
'功能：定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)
'参数：int场合=1-门诊,2-住院
    Call mfrmDiagReport.zlDefCommandBars(frmParent, CommandBars, int场合, bln护士站)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmDiagReport.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mfrmDiagReport.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmDiagReport.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal var就诊ID As Variant, ByVal int类型 As Integer, Optional ByVal blnMoved As Boolean)
'功能：刷新住院医嘱数据
'参数：var就诊ID=门诊为挂号NO,住院为主页ID
'      int类型：
'         门诊病人：0-不可编辑,1-可以编辑
'         住院病人：0-在院病人,1-预出院病人,2-出院病人,3-待会诊病人(在院),4-已会诊病人
'      blnMoved=该病人的数据是否已转出
    Call mfrmDiagReport.zlRefresh(lng病人ID, var就诊ID, int类型, blnMoved)
End Sub
