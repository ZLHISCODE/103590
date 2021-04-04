VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZLOKVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private vfwControl As New clsVfwCapture
Private mlngParentHwnd As Long

'*********************** 1.初始化 ***********************
Public Sub zlInit(cnOracle As ADODB.Connection, _
                    ByVal lngUserId As Long, _
                    ByVal lngCurDeptId As Long, _
                    ByVal lngParentHwnd As Long)
'如果需要初始化资源的，则在这里进行初始化处理
    mlngParentHwnd = lngParentHwnd
'    Call vfwControl.StartPreview(lngParentHwnd)
End Sub

'35.110版本支持
Public Sub GetSizeInfo(ByRef lngWidth As Long, ByRef lngHeight As Long)
    lngWidth = 800
    lngHeight = 600
End Sub

'35.110版本支持
Public Sub UpdateWindow(ByVal lngWidth As Long, ByVal lngHeight As Long)
    Call vfwControl.UpdateCaptureWindowPos(lngWidth, lngHeight)
End Sub

'35.110版本支持
Public Sub StartPreview()
    'vfwControl.StartPreview
    Call vfwControl.StartPreview(mlngParentHwnd)
End Sub

'35.110版本支持
Public Sub StopPreview()
    Call vfwControl.StopPreview
End Sub



'*********************** 2.图像采集 ***********************
Public Function zlCaptureImage(ByVal lngAdviceId As Long, _
                                ByRef objCapPic As StdPicture, _
                                ByRef strCapImgFiles As String, _
                                ByRef blnUseCustom As Boolean) As Boolean
'实现该方法进行图像采集
    Set objCapPic = vfwControl.CaptureImage ' LoadPicture(App.Path & "\capture1.bmp")
    
    blnUseCustom = False
    
    zlCaptureImage = True
End Function

'*********************** 3.视频录像 ***********************
Public Function zlCaptureVideo(ByVal lngAdviceId As Long, _
                        ByRef strVideoFiles As String, _
                        ByRef blnUseCustom As Boolean) As Boolean
'实现该方法进行视频录像

End Function

'*********************** 4.参数配置 ***********************
Public Sub zlConfigDevice()
'如需要进行参数配置的，则在该过程中处理
    vfwControl.ShowCaptureVideoSourceDialog
End Sub

'*********************** 5.资源释放 ***********************
Public Sub zlFree()
'如需要对相关资源进行释放的，则在该过程中处理

End Sub
