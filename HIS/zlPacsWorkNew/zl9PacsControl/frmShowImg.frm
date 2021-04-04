VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmShowImg 
   BackColor       =   &H00000000&
   Caption         =   "放大图像"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   555
   ClientWidth     =   4680
   Icon            =   "frmShowImg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin DicomObjects.DicomViewer Viewer 
      Height          =   2655
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   4335
      _Version        =   262147
      _ExtentX        =   7646
      _ExtentY        =   4683
      _StockProps     =   35
   End
End
Attribute VB_Name = "frmShowImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintStyle As Integer    '窗口模式1-固定窗口；2-独立窗口
Private mintOldStyle As Integer
Private mDcmImg As DicomImage
Private mlngleft As Long
Private mlngtop As Long
Private mdblBigImgZoom As Double        '报告大图放大倍数
Private mintLoadState As Integer
Private mblnBigImageCtl As Boolean

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Sub ShowMe(img As DicomImage, ObjFrm As Object, intStyle As Integer, Left As Long, Top As Long, blnBigImageCtl As Boolean, Optional dblBigImgZoom As Double = 1)

    mdblBigImgZoom = dblBigImgZoom
    
    If mintStyle <> intStyle Then mintOldStyle = mintStyle
    
    mintStyle = intStyle
    mlngleft = Left
    mlngtop = Top
    mblnBigImageCtl = blnBigImageCtl
    
    Set mDcmImg = img
    
    If mintStyle = 2 Then Me.BorderStyle = 2
    
    If mintLoadState <> 1 Then
        Me.Show
    Else
        Me.Viewer.Images.Clear
        Me.Viewer.Images.Add mDcmImg
        
        Call InitFaceSize(mDcmImg.SizeX * mdblBigImgZoom, mDcmImg.SizeY * mdblBigImgZoom)
    End If

End Sub

Public Sub HideMe()
    Unload Me
End Sub

Private Sub InitFaceSize(ByVal dblOldWidth As Double, ByVal dblOldHeight As Double)
    Dim blnZoomControl As Boolean
    Dim strImgMaxSize As String
    Dim dblNewWidth As Double, dblNewHeight As Double
    
    dblNewWidth = dblOldWidth
    dblNewHeight = dblOldHeight
    
    If mblnBigImageCtl Then
        blnZoomControl = Val(zlDatabase.GetPara("大图显示范围限制", glngSys, glngMoudle, "0")) <> 0
        
        If blnZoomControl Then
            strImgMaxSize = zlDatabase.GetPara("大图显示最大分辨率", glngSys, glngMoudle, "800*600")
            If Trim(strImgMaxSize) = "" Then strImgMaxSize = "800*600"
            
            If UBound(Split(strImgMaxSize, "*")) > 0 Then
                If dblOldWidth > Split(strImgMaxSize, "*")(0) Or dblOldHeight > Split(strImgMaxSize, "*")(1) Then
                    If dblOldWidth > dblOldHeight Then
                        If dblOldWidth > Split(strImgMaxSize, "*")(0) Then
                            dblNewWidth = Split(strImgMaxSize, "*")(0)
                            dblNewHeight = dblNewWidth * dblOldHeight / dblOldWidth
                        End If
                    Else
                        If dblOldHeight > Split(strImgMaxSize, "*")(1) Then
                            dblNewHeight = Split(strImgMaxSize, "*")(1)
                            dblNewWidth = dblNewHeight * dblOldWidth / dblOldHeight
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Me.Viewer.Left = 0
    Me.Viewer.Top = 0
    Me.Viewer.Width = dblNewWidth * Screen.TwipsPerPixelX
    Me.Viewer.Height = dblNewHeight * Screen.TwipsPerPixelY
    Me.Width = Me.Viewer.Width
    Me.Height = Me.Viewer.Height
End Sub

Public Sub Form_Load()
    Dim blnZoomControl As Boolean
    Dim strImgMaxSize As String
    Dim dblNewWidth As Double, dblNewHeight As Double
    
On Error GoTo ErrorHand

    Me.Viewer.Images.Clear
    Me.Viewer.Images.Add mDcmImg
    
    If mintStyle = 1 Then        '移动时显示大图，始终显示在界面的左上角
        If mintStyle <> mintOldStyle Then zlControl.FormSetCaption Me, False, False
        
        Me.Left = mlngleft
        Me.Top = mlngtop
        
        Call InitFaceSize(mDcmImg.SizeX * mdblBigImgZoom, mDcmImg.SizeY * mdblBigImgZoom)
    Else                        '鼠标单击显示大图窗口，读取窗口最后的位置
        If mintStyle <> mintOldStyle Then
            Call RestoreWinState(Me, App.ProductName)
        End If
    End If
        
    Me.Visible = True
    
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '&H1 Or &H2 Or &H10 Or &H40 '将窗口置顶
    mintLoadState = 1
    
    Exit Sub
ErrorHand:
    BUGEX "显示大图错误 err=" & err.Description
End Sub

Private Sub Form_Resize()
    If mintStyle = 2 Then
        Viewer.Left = 0
        Viewer.Top = 0
        Viewer.Width = Me.ScaleWidth
        Viewer.Height = Me.ScaleHeight
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '保存独立窗口的位置
    If mintStyle = 2 Then
        Call SaveWinState(Me, App.ProductName)
    End If
    
    mintLoadState = 2
End Sub

Private Sub Viewer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        ReleaseCapture      '解锁鼠标
        Call HideMe
    End If
End Sub
