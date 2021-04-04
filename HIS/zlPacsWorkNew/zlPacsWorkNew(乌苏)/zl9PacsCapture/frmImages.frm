VERSION 5.00
Begin VB.Form frmImages 
   BorderStyle     =   0  'None
   Caption         =   "检查图像"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   Icon            =   "frmImages.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmerRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   915
      Top             =   150
   End
   Begin zl9PacsCapture.ucImagePreview ucStudyImages 
      Height          =   1980
      Left            =   390
      TabIndex        =   0
      Top             =   720
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   3493
      BackColor       =   4210752
   End
   Begin VB.Timer tmerRefreshImages 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   150
      Top             =   150
   End
End
Attribute VB_Name = "frmImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private mlngAdviceId As Long
Private mstrStudyUid As String
Private mblnMoved As Boolean
Private mblnForceRefresh As Boolean

Private WithEvents mobjMsg As clsMsg
Attribute mobjMsg.VB_VarHelpID = -1

Public Event OnSelChange(ByVal lngSelectIndex As Long)
Public Event OnClick(ByVal lngSelectIndex As Long)
Public Event OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
Public Event OnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)



Property Get ImagePreviewObj() As ucImagePreview
'获取检查图像显示控件
    Set ImagePreviewObj = ucStudyImages
End Property


Public Sub RefreshImage(ByVal lngAdviceId As Long, ByVal strStudyUid As String, ByVal blnMoved As Boolean, _
    Optional ByVal blnForceRefresh As Boolean)
    
    mlngAdviceId = lngAdviceId
    mstrStudyUid = strStudyUid
    mblnMoved = blnMoved
    mblnForceRefresh = blnForceRefresh
    
BUGEX "RefreshImage 1 studyUid:" & strStudyUid, gblnUseDebugLog

    tmerRefreshImages.Enabled = True
    
BUGEX "RefreshImage End", gblnUseDebugLog
End Sub


Private Sub Form_Load()
    
    Set mobjMsg = New clsMsg
    Call mobjMsg.SetMsgHook(Me.hWnd)
    
    Call RegBroadcastRec(Me.hWnd)
    
    ucStudyImages.PageImgCount = Val(GetSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "影像缩略图数量", 5))
End Sub

Private Sub Form_Paint()
On Error GoTo errHandle

BUGEX "Form_Paint(frmImages)"
    tmerRefresh.Enabled = True
Exit Sub
errHandle:
End Sub

Private Sub Form_Resize()
On Error Resume Next
    ucStudyImages.Left = 0
    ucStudyImages.Top = 0
    ucStudyImages.Width = Me.ScaleWidth
    ucStudyImages.Height = Me.ScaleHeight
err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)

    mobjMsg.SetMsgUnHook
    Set mobjMsg = Nothing
    
    Call UnBroadcastRec(Me.hWnd)
    
    Call SaveSetting("ZLSOFT", G_STR_REG_PATH_PUBLIC, "影像缩略图数量", ucStudyImages.PageImgCount)
End Sub


Private Sub mobjMsg_OnWindowMessage(result As Long, ByVal lngHwnd As Long, ByVal lngMessage As Long, ByVal lngWParam As Long, ByVal lngLParam As Long)
On Error GoTo errHandle

    Select Case lngMessage
        Case WM_REFRESH_IMAGE
            If lngWParam = mlngAdviceId Then
                
BUGEX "mobjMsg_OnWindowMessage 1"
                tmerRefreshImages.Enabled = True
                
                result = 1
                Exit Sub
            End If
    End Select
    
    '其他我们不关心的消息自己不处理，必须由 VB 的默认处理函数处理
    result = mobjMsg.CallDefaultWindowProc(lngHwnd, lngMessage, lngWParam, lngLParam)
Exit Sub
errHandle:
End Sub

Private Sub tmerRefresh_Timer()
On Error GoTo errHandle
    tmerRefresh.Enabled = False
    
    Call ucStudyImages.RedrawSelf
Exit Sub
errHandle:
    BUGEX "tmerRefresh_Timer Err:" & err.Description
End Sub

Private Sub tmerRefreshImages_Timer()
On Error GoTo errHandle
    tmerRefreshImages.Enabled = False
'    Debug.Print "tmerRefreshStart" & Now
    Call ucStudyImages.RefreshImage(slStudy, mstrStudyUid, mblnMoved, mblnForceRefresh, False)
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub









Private Sub ucStudyImages_OnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
On Error Resume Next
    RaiseEvent OnCheckChange(lngSelectedIndex, blnSelected)
err.Clear
End Sub

Private Sub ucStudyImages_OnClick(ByVal lngSelectedIndex As Long)
On Error Resume Next
    RaiseEvent OnClick(lngSelectedIndex)
err.Clear
End Sub

Private Sub ucStudyImages_OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
On Error Resume Next
    RaiseEvent OnDbClick(lngSelectedIndex, blnContinue)
err.Clear
End Sub

Private Sub ucStudyImages_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
err.Clear
End Sub

Private Sub ucStudyImages_OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
err.Clear
End Sub

Private Sub ucStudyImages_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
err.Clear
End Sub

Private Sub ucStudyImages_OnMouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
On Error Resume Next
    RaiseEvent OnMouseWheel(Shift, Delta, X, Y)
err.Clear
End Sub

Private Sub ucStudyImages_OnSelChange(ByVal lngSelectedIndex As Long)
On Error Resume Next
    RaiseEvent OnSelChange(lngSelectedIndex)
err.Clear
End Sub
