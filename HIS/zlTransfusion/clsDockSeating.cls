VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockSeating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Seatings"
Option Explicit

Public Event Activate() '子窗体激活
Public Event RequestRefresh() '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字
Private WithEvents mfrmSeating As frmDockSeat
Attribute mfrmSeating.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mfrmSeating = New frmDockSeat
End Sub

Private Sub Class_Terminate()
    Unload mfrmSeating
    Set mfrmSeating = Nothing
End Sub

Private Sub mfrmSeating_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmSeating_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub mfrmSeating_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Public Function zlGetForm() As Object
'功能：获取子业务窗口,由主程序初始化时只调用一次获取

    '初始读取权限,以加快后续速度(不是必要的)
    Call GetInsidePrivs(1264, True)
    
    '一些接口初始化,只初始化一次
    '因涉及具体模块的权限,因此这里初始化
    
    '电子签名初始化:因涉及具体模块的权限,因此这里初始化
    '因为几个模块可能同时使用,且共用变量gobjESign，因此还需要在具体模块中判断权限

    Set zlGetForm = mfrmSeating
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int场合 As Integer)
'功能：定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)
'参数：int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    Call mfrmSeating.zlDefCommandBars(frmParent, CommandBars, int场合)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmSeating.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mfrmSeating.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmSeating.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal curSeatings As Seatings, ByVal lng病人ID As Long, ByVal ObjPati As cPatient)
'功能：刷新住院医嘱数据
'参数：lng前提ID=当由医技站调用时传入
'      blnMoved=该病人的数据是否已转出
    mfrmSeating.lng病人ID = lng病人ID
    Set mfrmSeating.ObjPati = ObjPati
    Call mfrmSeating.zlRefresh(curSeatings)
End Sub

