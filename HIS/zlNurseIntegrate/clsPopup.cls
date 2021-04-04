VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'弹窗属性
Private Type TYPE_Popup_Param
    Params As String '传回的参数串原格式
    Type As String ': "CloseDialog" || "ShowDialog", // CloseDialog 关闭弹窗  ShowDialog 打开弹窗
    ModuleUrl As String ': "/shiftReport", //功能Url
    Title As String ': "交班报告",
    Width As Long ': "100" || null,
    Height As Long ': "100" || null,
    Minimal As Boolean ': true,  //最大化
    Max As Boolean ': false,     //最小化
    IsRefresh As Boolean ': true  //是否刷新父窗体
    Data As String ': "xxxxxxxxxxxx"  //打开弹窗是需要带上的参数
    UnitID As String '整体护理中的病区ID
    UserID As String '整体护理中的用户ID
    ParentUrl As String '触发事件的主窗体URL
    ParentParam As String '触发时间的主窗体的参数
    PatientID As String '整体护理中的病人ID
    PrentForm As Object  '刷新时间调用主窗体
End Type

Private mPopupParam As TYPE_Popup_Param

'传回的参数串原格式
Public Property Let PopupParams(strValue As String)
    mPopupParam.Params = strValue
End Property

Public Property Get PopupParams() As String
    PopupParams = mPopupParam.Params
End Property

'刷新时间调用主窗体
Public Property Set PopupParentForm(strValue As Object)
    Set mPopupParam.PrentForm = strValue
End Property

Public Property Get PopupParentForm() As Object
    Set PopupParentForm = mPopupParam.PrentForm
End Property

'触发事件的窗体URL
Public Property Let PopupParentUrl(strValue As String)
    mPopupParam.ParentUrl = strValue
End Property

Public Property Get PopupParentUrl() As String
    PopupParentUrl = mPopupParam.ParentUrl
End Property

'触发事件主窗体的参数
Public Property Let PopupParentParam(strValue As String)
    mPopupParam.ParentParam = strValue
End Property

Public Property Get PopupParentParam() As String
    PopupParentParam = mPopupParam.ParentParam
End Property

'整体护理中的病区ID
Public Property Let PopupUnitID(strValue As String)
    mPopupParam.UnitID = strValue
End Property

Public Property Get PopupUnitID() As String
    PopupUnitID = mPopupParam.UnitID
End Property

'整体护理中的用户ID
Public Property Let PopupUserID(strValue As String)
    mPopupParam.UserID = strValue
End Property

Public Property Get PopupUserID() As String
    PopupUserID = mPopupParam.UserID
End Property

'整体护理中的病人ID
Public Property Let PopupPatientID(strValue As String)
    mPopupParam.PatientID = strValue
End Property

Public Property Get PopupPatientID() As String
    PopupPatientID = mPopupParam.PatientID
End Property


': "/shiftReport", //功能Url
Public Property Let PopupModuleUrl(strValue As String)
    mPopupParam.ModuleUrl = strValue
End Property

Public Property Get PopupModuleUrl() As String
    PopupModuleUrl = mPopupParam.ModuleUrl
End Property

': "CloseDialog" || "ShowDialog", // CloseDialog 关闭弹窗  ShowDialog 打开弹窗
Public Property Let PopupType(strValue As String)
    mPopupParam.Type = strValue
End Property

Public Property Get PopupType() As String
    PopupType = mPopupParam.Type
End Property

': "交班报告",
Public Property Let PopupTitle(strValue As String)
    mPopupParam.Title = strValue
End Property

Public Property Get PopupTitle() As String
    PopupTitle = mPopupParam.Title
End Property


 ': "100" || null,
Public Property Let PopupWidth(strValue As String)
    mPopupParam.Width = strValue
End Property

Public Property Get PopupWidth() As String
    PopupWidth = mPopupParam.Width
End Property
': "100" || null,
Public Property Let PopupHeight(strValue As String)
    mPopupParam.Height = strValue
End Property

Public Property Get PopupHeight() As String
    PopupHeight = mPopupParam.Height
End Property
 ': true,  //最大化
Public Property Let PopupMinimal(blnValue As Boolean)
    mPopupParam.Minimal = blnValue
End Property

Public Property Get PopupMinimal() As Boolean
    PopupMinimal = mPopupParam.Minimal
End Property
': false,     //最小化
Public Property Let PopupMax(blnValue As Boolean)
    mPopupParam.Max = blnValue
End Property

Public Property Get PopupMax() As Boolean
    PopupMax = mPopupParam.Max
End Property
': true  //是否刷新父窗体
Public Property Let PopupIsRefresh(blnValue As Boolean)
    mPopupParam.IsRefresh = blnValue
End Property

Public Property Get PopupIsRefresh() As Boolean
    PopupIsRefresh = mPopupParam.IsRefresh
End Property
': "xxxxxxxxxxxx"  //打开弹窗是需要带上的参数
Public Property Let PopupData(strValue As String)
    mPopupParam.Data = strValue
End Property

Public Property Get PopupData() As String
    PopupData = mPopupParam.Data
End Property


