VERSION 5.00
Begin VB.UserControl PictureEditor 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "PictureEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mfMain As fMain
Attribute mfMain.VB_VarHelpID = -1
Public lngKeyOfPic As Long                      '图片的Key值

Public Event pOK(ByRef FinalPicture As StdPicture, ByVal lngWidth As Long, ByVal lngHeight As Long)    '保存，返回修改后的临时图片路径（JPEG图片）
Public Event pCancel()                          '取消并退出

'################################################################################################################
'## 功能：  显示编辑主窗体
'##
'## 参数：  lngSys      :系统号
'##         cnMain      :数据库连接
'##         srcPic      :源图片 StcPicture
'##         lngKey      :图片Key值
'##         frmParent   :父窗体
'##         bln保留     :该图片对象是否保留，如果是，则不允许编辑时打开其他图片
'################################################################################################################
Public Sub ShowPicEditor(ByVal lngSys As Long, _
    ByRef cnMain As ADODB.Connection, _
    ByRef srcPic As StdPicture, _
    Optional lngKey As Long = 0, _
    Optional bln保留 As Boolean, _
    Optional ByRef frmParent As Object = Nothing)
    
    Call InitCommon(cnMain)
    glngSys = lngSys
    lngKeyOfPic = lngKey
    gbln保留 = bln保留
    
    If mfMain Is Nothing Then
        Set mfMain = New fMain
        Set gfrmMain = mfMain
    End If
    If gfDialogEx Is Nothing Then Set gfDialogEx = New fDialogEx
    If gfFilter Is Nothing Then Set gfFilter = New fFilter
    If gfOrientation Is Nothing Then Set gfOrientation = New fOrientation
    If gfPanView Is Nothing Then Set gfPanView = New fPanView
    If gfPrint Is Nothing Then Set gfPrint = New fPrint
    If gfProperties Is Nothing Then Set gfProperties = New fProperties
    If gfResize Is Nothing Then Set gfResize = New fResize
    If gfTexturize Is Nothing Then Set gfTexturize = New fTexturize
    
    Call gfrmMain.ShowMe(srcPic, frmParent)
End Sub


'################################################################################################################
'## 功能：  释放资源
'################################################################################################################

Private Sub mfMain_pCancel()
    RaiseEvent pCancel
End Sub

Private Sub mfMain_pOK(ByRef FinalPicture As StdPicture, ByVal lngWidth As Long, ByVal lngHeight As Long)
    RaiseEvent pOK(FinalPicture, lngWidth, lngHeight)
End Sub

Private Sub UserControl_Initialize()
    If UserControl.Ambient.UserMode Then
        If mfMain Is Nothing Then Set mfMain = New fMain
        Set gfrmMain = mfMain
    End If
End Sub

Private Sub UserControl_Terminate()
    If UserControl.Ambient.UserMode Then
        On Error Resume Next
        Unload gfrmMain
        Unload gfDialogEx
        Unload gfFilter
        Unload gfOrientation
        Unload gfPanView
        Unload gfPrint
        Unload gfProperties
        Unload gfResize
        Unload gfTexturize
        Unload mfMain
        
        Set gfrmMain = Nothing
        Set gfDialogEx = Nothing
        Set gfFilter = Nothing
        Set gfOrientation = Nothing
        Set gfPanView = Nothing
        Set gfPrint = Nothing
        Set gfProperties = Nothing
        Set gfResize = Nothing
        Set gfTexturize = Nothing
        Set mfMain = Nothing
    End If
End Sub
