VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWardMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enabled As Boolean
Private mobjWardMonitor As Object

Public Sub HideWindow()
    Call mobjWardMonitor.ShowWindow(False)
End Sub

Public Function GetListPati() As String
'功能：获取连接到监护中心的病人（住院号）列表，以，逗号分开
    Dim strList As String
    
    Call mobjWardMonitor.GetListBedNo(strList)
    GetListPati = strList
End Function

Public Function zlGetForm() As Object
    Set zlGetForm = frmWardMonitor
End Function

Public Sub ShowInfor(ByVal strInPatiNO As String)
    Call mobjWardMonitor.ShowInfor(strInPatiNO)
End Sub

Private Sub Class_Initialize()
    
    On Error Resume Next
    Set mobjWardMonitor = CreateObject("zlWardMonitor.clsWardMonitor")
    If Err.Number = 0 Then
        Load zlGetForm  '加载窗体，以便监护仪能够通过窗体事件获取病人信息，否则监护仪上下载病人信息时，会导致本程序异常退出
        Enabled = mobjWardMonitor.Init(gcnOracle, frmWardMonitor.hwnd)
        If Enabled = False Then
            Unload zlGetForm
        End If
    Else
        Enabled = False
    End If
End Sub

Private Sub Class_Terminate()
    Set mobjWardMonitor = Nothing
End Sub
