VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStudyImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mobjImages As frmImages
Attribute mobjImages.VB_VarHelpID = -1

Public Event OnSelChange(ByVal lngSelectIndex As Long)
Public Event OnClick(ByVal lngSelectIndex As Long)
Public Event OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
Public Event OnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)


Property Get hwnd() As Long
'窗体句柄
    hwnd = mobjImages.hwnd
End Property


Property Get Window() As Object
'窗体对象
    Set Window = mobjImages
End Property

Property Get Images() As Object
    Set Images = mobjImages.ImagePreviewObj
End Property

Public Sub Init(ByVal lngPageImgCount As Long)
    mobjImages.ImagePreviewObj.PageImgCount = lngPageImgCount
End Sub

Public Sub RefreshImages(ByVal lngAdviceId As Long, ByVal strStudyUid As String, _
    ByVal blnMoved As Boolean, Optional ByVal blnForceRefresh As Boolean)
    
BUGEX "RefreshImages 1", gblnUseDebugLog
    Call mobjImages.RefreshImage(lngAdviceId, strStudyUid, blnMoved, blnForceRefresh)
    
BUGEX "RefreshImages End", gblnUseDebugLog
End Sub

Private Sub Class_Initialize()
    Set mobjImages = New frmImages
End Sub

Private Sub Class_Terminate()
    Unload mobjImages
    Set mobjImages = Nothing
End Sub

Private Sub mobjImages_OnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
On Error Resume Next
    RaiseEvent OnCheckChange(lngSelectedIndex, blnSelected)
err.Clear
End Sub

Private Sub mobjImages_OnClick(ByVal lngSelectedIndex As Long)
On Error Resume Next
    RaiseEvent OnClick(lngSelectedIndex)
err.Clear
End Sub

Private Sub mobjImages_OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
On Error Resume Next
    RaiseEvent OnDbClick(lngSelectedIndex, blnContinue)
err.Clear
End Sub

Private Sub mobjImages_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
err.Clear
End Sub

Private Sub mobjImages_OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
err.Clear
End Sub

Private Sub mobjImages_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
err.Clear
End Sub

Private Sub mobjImages_OnMouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
On Error Resume Next
    RaiseEvent OnMouseWheel(Shift, Delta, X, Y)
err.Clear
End Sub

Private Sub mobjImages_OnSelChange(ByVal lngSelectedIndex As Long)
On Error Resume Next
    RaiseEvent OnSelChange(lngSelectedIndex)
err.Clear
End Sub
