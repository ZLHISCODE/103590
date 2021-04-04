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

'��������
Private Type TYPE_Popup_Param
    Params As String '���صĲ�����ԭ��ʽ
    Type As String ': "CloseDialog" || "ShowDialog", // CloseDialog �رյ���  ShowDialog �򿪵���
    ModuleUrl As String ': "/shiftReport", //����Url
    Title As String ': "���౨��",
    Width As Long ': "100" || null,
    Height As Long ': "100" || null,
    Minimal As Boolean ': true,  //���
    Max As Boolean ': false,     //��С��
    IsRefresh As Boolean ': true  //�Ƿ�ˢ�¸�����
    Data As String ': "xxxxxxxxxxxx"  //�򿪵�������Ҫ���ϵĲ���
    UnitID As String '���廤���еĲ���ID
    UserID As String '���廤���е��û�ID
    ParentUrl As String '�����¼���������URL
    ParentParam As String '����ʱ���������Ĳ���
    PatientID As String '���廤���еĲ���ID
    PrentForm As Object  'ˢ��ʱ�����������
End Type

Private mPopupParam As TYPE_Popup_Param

'���صĲ�����ԭ��ʽ
Public Property Let PopupParams(strValue As String)
    mPopupParam.Params = strValue
End Property

Public Property Get PopupParams() As String
    PopupParams = mPopupParam.Params
End Property

'ˢ��ʱ�����������
Public Property Set PopupParentForm(strValue As Object)
    Set mPopupParam.PrentForm = strValue
End Property

Public Property Get PopupParentForm() As Object
    Set PopupParentForm = mPopupParam.PrentForm
End Property

'�����¼��Ĵ���URL
Public Property Let PopupParentUrl(strValue As String)
    mPopupParam.ParentUrl = strValue
End Property

Public Property Get PopupParentUrl() As String
    PopupParentUrl = mPopupParam.ParentUrl
End Property

'�����¼�������Ĳ���
Public Property Let PopupParentParam(strValue As String)
    mPopupParam.ParentParam = strValue
End Property

Public Property Get PopupParentParam() As String
    PopupParentParam = mPopupParam.ParentParam
End Property

'���廤���еĲ���ID
Public Property Let PopupUnitID(strValue As String)
    mPopupParam.UnitID = strValue
End Property

Public Property Get PopupUnitID() As String
    PopupUnitID = mPopupParam.UnitID
End Property

'���廤���е��û�ID
Public Property Let PopupUserID(strValue As String)
    mPopupParam.UserID = strValue
End Property

Public Property Get PopupUserID() As String
    PopupUserID = mPopupParam.UserID
End Property

'���廤���еĲ���ID
Public Property Let PopupPatientID(strValue As String)
    mPopupParam.PatientID = strValue
End Property

Public Property Get PopupPatientID() As String
    PopupPatientID = mPopupParam.PatientID
End Property


': "/shiftReport", //����Url
Public Property Let PopupModuleUrl(strValue As String)
    mPopupParam.ModuleUrl = strValue
End Property

Public Property Get PopupModuleUrl() As String
    PopupModuleUrl = mPopupParam.ModuleUrl
End Property

': "CloseDialog" || "ShowDialog", // CloseDialog �رյ���  ShowDialog �򿪵���
Public Property Let PopupType(strValue As String)
    mPopupParam.Type = strValue
End Property

Public Property Get PopupType() As String
    PopupType = mPopupParam.Type
End Property

': "���౨��",
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
 ': true,  //���
Public Property Let PopupMinimal(blnValue As Boolean)
    mPopupParam.Minimal = blnValue
End Property

Public Property Get PopupMinimal() As Boolean
    PopupMinimal = mPopupParam.Minimal
End Property
': false,     //��С��
Public Property Let PopupMax(blnValue As Boolean)
    mPopupParam.Max = blnValue
End Property

Public Property Get PopupMax() As Boolean
    PopupMax = mPopupParam.Max
End Property
': true  //�Ƿ�ˢ�¸�����
Public Property Let PopupIsRefresh(blnValue As Boolean)
    mPopupParam.IsRefresh = blnValue
End Property

Public Property Get PopupIsRefresh() As Boolean
    PopupIsRefresh = mPopupParam.IsRefresh
End Property
': "xxxxxxxxxxxx"  //�򿪵�������Ҫ���ϵĲ���
Public Property Let PopupData(strValue As String)
    mPopupParam.Data = strValue
End Property

Public Property Get PopupData() As String
    PopupData = mPopupParam.Data
End Property

