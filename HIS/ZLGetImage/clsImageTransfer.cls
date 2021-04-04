VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjEvent As Object
Private WithEvents mobjFrmMain As frmMain
Attribute mobjFrmMain.VB_VarHelpID = -1

Public Sub zlInitModule(ByVal blnIsUpload As Boolean, ByVal lngThreadID As Long)
    frmMain.zlInitModule blnIsUpload, lngThreadID
End Sub

Public Sub zlLoadImage()
    frmMain.zlLoadImage
End Sub

Public Sub MsgInQueue(objMsg As Object)
mobjFrmMain.BUGEX "MsgInQueue 0"
    Call frmMain.mobjDataQueue.MsgInQueue(objMsg)
mobjFrmMain.BUGEX "MsgInQueue 1"
End Sub

Public Sub RegEventObj(ByVal objForm As Object)
    Set mobjEvent = objForm
End Sub

Private Sub Class_Initialize()
    Set mobjFrmMain = frmMain
End Sub

Public Sub zlUnloadObj()
'ÊÍ·Å×ÊÔ´
mobjFrmMain.BUGEX "zlGetImage                                               Class_Terminate"
On Error GoTo errHandle
    Set mobjFrmMain = Nothing
    
    Set mobjEvent = Nothing

    Unload frmMain

    Exit Sub
errHandle:
frmMain.BUGEX "Class_Terminate Err " & err.Description
End Sub

Private Sub mobjFrmMain_OnComPlete(ByVal curMsg As Object)
mobjFrmMain.BUGEX "mobjFrmMain_OnComPlete Start"
    On Error GoTo err
    
    If Not mobjEvent Is Nothing Then
mobjFrmMain.BUGEX "mobjFrmMain_OnComPlete 0"
        Call mobjEvent.OnComPlete(curMsg)
mobjFrmMain.BUGEX "mobjFrmMain_OnComPlete 1"
    End If
mobjFrmMain.BUGEX "mobjFrmMain_OnComPlete End"
    Exit Sub
err:
mobjFrmMain.BUGEX "mobjFrmMain_OnComPlete Err " & err.Description
End Sub

Private Sub mobjFrmMain_OnState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean, ByVal lngThreadID As Long)
    On Error GoTo err
mobjFrmMain.BUGEX IIf(mobjEvent Is Nothing, "mobjEvent is nothing", "mobjEvent not is nothing")
    If Not mobjEvent Is Nothing Then
        Call mobjEvent.OnState(blnLoadFinish, blnUpLoad, lngThreadID)
    End If

    Exit Sub
err:
mobjFrmMain.BUGEX "mobjFrmMain_OnState Err " & err.Description
End Sub

Public Property Get ImgInfo() As Object
    Set ImgInfo = New clsImgInfo
End Property
