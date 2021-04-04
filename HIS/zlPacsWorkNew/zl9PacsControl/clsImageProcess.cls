VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mobjImageProcess As frmImageProcess
Attribute mobjImageProcess.VB_VarHelpID = -1

Public Event OnSaveImage(ByVal dcmImage As DicomImage, ByVal lngImageType As Long)
Public Event OnUnload()
Public Event AfterSaveStady()

Public Sub ShowImageProcess(ByVal strQueryValue As String, dcmImage As DicomImage, _
    lngImgIndex As Long, objParent As Object, blnMoved As Boolean, Optional lngLeval As Long, Optional lngType As Long, Optional lngPreviewTime As Long, Optional blnDoShiled As Boolean)
    Dim blnRefresh As Boolean
    
    If gobjImageProcess Is Nothing Then
        Set mobjImageProcess = New frmImageProcess
        Set gobjImageProcess = mobjImageProcess
        blnRefresh = False
    Else
         
        Set mobjImageProcess = gobjImageProcess
        '²»ÊÇÔ¤ÀÀÍ¼ÏñÊ±£¬ÒÆ¶¯Êó±êÇÐ»»Í¼Ïñ²»Ë¢ÐÂ
        If lngType = 1 And lngPreviewTime > 0 And Not mobjImageProcess.IsRefresh Then
            Exit Sub
        End If
        
        blnRefresh = True
    End If
    
    mobjImageProcess.zlShowMe strQueryValue, dcmImage, lngImgIndex, objParent, blnRefresh, blnMoved, lngLeval, lngType, lngPreviewTime, blnDoShiled
End Sub
'
'Public Sub Free()
'On Error GoTo errHandle
'
'    If Not mobjImageProcess Is Nothing Then
'
'        Unload mobjImageProcess
'        Set mobjImageProcess = Nothing
'    End If
'
'
'errHandle:
'End Sub

Private Sub mobjImageProcess_AfterSaveStady()
    RaiseEvent AfterSaveStady
End Sub


Private Sub mobjImageProcess_OnSaveImage(ByVal dcmImage As DicomObjects.DicomImage, ByVal emImageType As TImageType)
    RaiseEvent OnSaveImage(dcmImage, emImageType)
End Sub

Private Sub mobjImageProcess_OnUnload()
    RaiseEvent OnUnload

    Set mobjImageProcess = Nothing
    Set gobjImageProcess = Nothing
End Sub

Public Sub AfterSaveStudy(dcmImage As DicomImage)
    If Not mobjImageProcess Is Nothing Then
        mobjImageProcess.AfterSaveStudy dcmImage
    End If
End Sub

