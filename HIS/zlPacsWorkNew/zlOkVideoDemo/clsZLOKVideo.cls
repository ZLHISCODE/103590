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

'*********************** 1.��ʼ�� ***********************
Public Sub zlInit(cnOracle As ADODB.Connection, _
                    ByVal lngUserId As Long, _
                    ByVal lngCurDeptId As Long, _
                    ByVal lngParentHwnd As Long)
'�����Ҫ��ʼ����Դ�ģ�����������г�ʼ������
    mlngParentHwnd = lngParentHwnd
'    Call vfwControl.StartPreview(lngParentHwnd)
End Sub

'35.110�汾֧��
Public Sub GetSizeInfo(ByRef lngWidth As Long, ByRef lngHeight As Long)
    lngWidth = 800
    lngHeight = 600
End Sub

'35.110�汾֧��
Public Sub UpdateWindow(ByVal lngWidth As Long, ByVal lngHeight As Long)
    Call vfwControl.UpdateCaptureWindowPos(lngWidth, lngHeight)
End Sub

'35.110�汾֧��
Public Sub StartPreview()
    'vfwControl.StartPreview
    Call vfwControl.StartPreview(mlngParentHwnd)
End Sub

'35.110�汾֧��
Public Sub StopPreview()
    Call vfwControl.StopPreview
End Sub



'*********************** 2.ͼ��ɼ� ***********************
Public Function zlCaptureImage(ByVal lngAdviceId As Long, _
                                ByRef objCapPic As StdPicture, _
                                ByRef strCapImgFiles As String, _
                                ByRef blnUseCustom As Boolean) As Boolean
'ʵ�ָ÷�������ͼ��ɼ�
    Set objCapPic = vfwControl.CaptureImage ' LoadPicture(App.Path & "\capture1.bmp")
    
    blnUseCustom = False
    
    zlCaptureImage = True
End Function

'*********************** 3.��Ƶ¼�� ***********************
Public Function zlCaptureVideo(ByVal lngAdviceId As Long, _
                        ByRef strVideoFiles As String, _
                        ByRef blnUseCustom As Boolean) As Boolean
'ʵ�ָ÷���������Ƶ¼��

End Function

'*********************** 4.�������� ***********************
Public Sub zlConfigDevice()
'����Ҫ���в������õģ����ڸù����д���
    vfwControl.ShowCaptureVideoSourceDialog
End Sub

'*********************** 5.��Դ�ͷ� ***********************
Public Sub zlFree()
'����Ҫ�������Դ�����ͷŵģ����ڸù����д���

End Sub