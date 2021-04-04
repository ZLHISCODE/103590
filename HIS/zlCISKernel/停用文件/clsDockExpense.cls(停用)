VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Activate() '子窗体激活
Public Event RequestRefresh() '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字
Private WithEvents mfrmExpense As frmDockExpense
Attribute mfrmExpense.VB_VarHelpID = -1
Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘兴洪
    '日期:2012-06-18 16:50:35
    '问题:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmExpense Is Nothing Then Exit Sub
    mfrmExpense.SetFontSize (bytSize)
End Sub
Private Sub Class_Initialize()
    Set mfrmExpense = New frmDockExpense
End Sub

Private Sub Class_Terminate()
    Unload mfrmExpense
    Set mfrmExpense = Nothing
    Set grs医疗付款方式 = Nothing
End Sub

Private Sub mfrmExpense_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmExpense_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub mfrmExpense_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Public Function zlGetForm() As Object
'功能：获取子业务窗口,由主程序初始化时只调用一次获取

    '初始化医保部件:可分散到具体的GetForm中去初始化
    Call gclsInsure.InitOracle(gcnOracle)

    '初始读取权限,以加快后续速度(不是必要的)
    Call GetInsidePrivs(p医嘱附费管理, True)
    
    Set zlGetForm = mfrmExpense
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, Optional ByRef objSquareCard As Object)
'功能：定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)
'参数：CommandBars=仅用于查看时可以不传(传入Nothing)
'      objSquareCard=通过医技工作站传入的一卡通结算部件
    Call mfrmExpense.zlDefCommandBars(frmParent, CommandBars, objSquareCard)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmExpense.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mfrmExpense.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmExpense.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng科室id As Long, ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, Optional ByVal blnMoved As Boolean, Optional ByVal bln单独执行 As Boolean)
'功能：刷新医嘱附费数据
'参数：lng科室ID=当前医技科室ID
'      lng医嘱ID=相当于医技站的前提ID
'      blnMoved=该病人的数据是否已转出
'      bln单独执行=用于检验项目，一并采集的一组项目，是否针对其中的某一个单独执行
    Call mfrmExpense.zlRefresh(lng科室id, lng医嘱ID, lng发送号, blnMoved, bln单独执行)
End Sub
