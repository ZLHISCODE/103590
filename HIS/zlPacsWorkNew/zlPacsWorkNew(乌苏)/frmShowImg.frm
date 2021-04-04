VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmShowImg 
   BorderStyle     =   0  'None
   Caption         =   "放大图像"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   4680
   Icon            =   "frmShowImg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin DicomObjects.DicomViewer Viewer 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
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

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub ShowMe(img As DicomImage, ObjFrm As Object, intStyle As Integer, Left As Long, Top As Long, Optional dblBigImgZoom As Double = 1)

    mdblBigImgZoom = dblBigImgZoom
    mintOldStyle = mintStyle
    mintStyle = intStyle
    mlngleft = Left
    mlngtop = Top
    Set mDcmImg = img
    
    If mintLoadState <> 1 Then
        Load Me
    Else
        Me.Viewer.Images.Clear
        Me.Viewer.Images.Add mDcmImg
    End If

End Sub

Public Sub HideMe()
    Unload Me
End Sub

Public Sub Form_Load()

    Me.Viewer.Images.Clear
    Me.Viewer.Images.Add mDcmImg
    
    If mintStyle = 1 Then        '移动时显示大图，始终显示在界面的左上角
        If mintStyle <> mintOldStyle Then FormSetCaption Me, False, False
        Me.Left = mlngleft
        Me.Top = mlngtop
        
'        If Viewer.Images(1).SizeX * Screen.TwipsPerPixelX * mdblBigImgZoom >= Screen.Width Or _
'            Viewer.Images(1).SizeY * Screen.TwipsPerPixelY * mdblBigImgZoom >= Screen.Height Then
'            Viewer.Width = Viewer.Images(1).SizeX * Screen.TwipsPerPixelX
'            Viewer.Height = Viewer.Images(1).SizeY * Screen.TwipsPerPixelY
'        Else
            Viewer.Width = Viewer.Images(1).SizeX * Screen.TwipsPerPixelX * mdblBigImgZoom
            Viewer.Height = Viewer.Images(1).SizeY * Screen.TwipsPerPixelY * mdblBigImgZoom
'        End If
        Me.Width = Viewer.Width
        Me.Height = Viewer.Height
        Viewer.Left = 0
        Viewer.Top = 0
        Viewer.Width = Me.ScaleWidth
        Viewer.Height = Me.ScaleHeight
    Else                        '鼠标单击显示大图窗口，读取窗口最后的位置
        If mintStyle <> mintOldStyle Then
            Call RestoreWinState(Me, App.ProductName)
        End If
    End If
    
    
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, &H1 Or &H2 Or &H10 Or &H40 '将窗口置顶
    mintStyle = 0
    mintLoadState = 1
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
