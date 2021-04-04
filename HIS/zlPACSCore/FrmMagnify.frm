VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form FrmMagnify 
   Caption         =   "放大镜"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   Icon            =   "FrmMagnify.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   262
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkInvert 
      Caption         =   "反白"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   3360
      Width           =   660
   End
   Begin VB.CheckBox chkOrganLens 
      Caption         =   "透镜"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3360
      Width           =   660
   End
   Begin VB.CommandButton CmdHid 
      Caption         =   "隐藏"
      Height          =   345
      Left            =   3360
      TabIndex        =   2
      Top             =   3360
      Width           =   585
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      _Version        =   393216
      Min             =   1
      Max             =   80
      SelStart        =   20
      TickStyle       =   3
      Value           =   20
   End
   Begin DicomObjects.DicomViewer Viewer1 
      Height          =   3165
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3855
      _Version        =   262147
      _ExtentX        =   6800
      _ExtentY        =   5583
      _StockProps     =   35
      BackColor       =   -2147483641
      AsyncReceive    =   -1  'True
      UseScrollBars   =   0   'False
   End
   Begin VB.Label lblZoomState 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   3135
   End
End
Attribute VB_Name = "FrmMagnify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public f As Form

Dim MP As POINTAPI
'用于保存父窗体和需要放大的DICOM控件
Dim FrmClick As Boolean
Dim NowImg As Integer               '当前图像
Dim IntBC As Integer                '设置步长
Dim BeginWidth, BeginHeight As Integer  '当前窗体位置
Dim blnOrganLens As Boolean     '是否进入了组织透镜状态
Dim blnInvert As Boolean        '是否进入反白状态
Dim intMaxTop As Long           '窗体最大的TOP
Dim intMaxLeft As Long          '窗体最大的Left
Dim lngBaseXX As Long
Dim lngBaseYY As Long
Dim lngMagnifyWidth As Long
Dim lngMagnifyLevel As Long


Private Sub chkInvert_Click()
    blnInvert = IIf(chkInvert.Value = 1, True, False)
    If Me.Viewer1.Images.Count = 0 Then Exit Sub
    Call subFlipRotate(Viewer1.Images(1), "Invert")
End Sub

Private Sub chkOrganLens_Click()
    blnOrganLens = IIf(chkOrganLens.Value = 1, True, False)
    If blnOrganLens Then
        subOrganLens
    Else
        If Me.Viewer1.Images.Count > 0 Then
            Me.Viewer1.Images(1).SetDefaultWindows
        End If
    End If
End Sub

'退出
Private Sub CmdHid_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    '隐藏标题栏
    Call zlcontrol.FormSetCaption(Me, False)
        
    Me.left = f.left + f.width / 2
    Me.top = f.top + f.height / 2
    intMaxTop = GetToolBarBottomOrRight(2)
    intMaxLeft = GetToolBarBottomOrRight(1)
End Sub
'自适应窗体
Private Sub Form_Resize()
    On Error Resume Next
    'viewer1
    Me.Viewer1.top = 1
    Me.Viewer1.left = 1
    Me.Viewer1.width = Me.ScaleWidth - 2
    Me.Viewer1.height = Me.ScaleHeight - Me.Slider1.height - Me.lblZoomState.height - 5
    'lblZoomState
    Me.lblZoomState.top = Me.Viewer1.height + 2
    Me.lblZoomState.left = 1
    Me.lblZoomState.width = Me.ScaleWidth - Me.CmdHid.width - 1
    If Me.Viewer1.Images.Count > 0 Then
        Me.lblZoomState.Caption = "   放大倍数：" & Format(Me.Viewer1.Images(1).ActualZoom, "###0.0000")
    Else
        Me.lblZoomState.Caption = "   放大倍数："
    End If
    'slider1
    Me.Slider1.top = Me.lblZoomState.top + Me.lblZoomState.height + 2
    Me.Slider1.left = 1
    Me.Slider1.width = Abs(Me.ScaleWidth - Me.CmdHid.width - Me.chkOrganLens.width - Me.chkInvert.width - 1)
    'chkOrganLens
    Me.chkOrganLens.top = Me.Slider1.top
    Me.chkOrganLens.left = Me.Slider1.left + Me.Slider1.width
    'chkInvert
    Me.chkInvert.top = Me.Slider1.top
    Me.chkInvert.left = Me.chkOrganLens.left + Me.chkOrganLens.width
    'cmd
    Me.CmdHid.top = Me.Slider1.top - 2
    Me.CmdHid.left = Me.chkInvert.left + Me.chkInvert.width
    '刷新
    ImgMagnify
End Sub

'改变缩放比例
Private Sub Slider1_Change()
    On Error Resume Next
    If Me.Viewer1.Images.Count = 0 Then Exit Sub
    Me.CmdHid.SetFocus
    ImgMagnify
    Me.lblZoomState.Caption = "   放大倍数：" & Format(Me.Viewer1.Images(1).ActualZoom, "###0.0000")
End Sub

Private Sub Viewer1_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = cMouseUsage("102").lngMouseKey Then
        FrmClick = True
        lngBaseXX = x
        lngBaseYY = y
    Else
        '可以开始拖移
        FrmClick = True
        BeginWidth = x
        BeginHeight = y
    End If
End Sub

Private Sub Viewer1_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    '用于移动窗体
    Dim TmpMp As POINTAPI
    Dim Mx, My As Integer
    On Error Resume Next
    If FrmClick = True Then
        If Button = cMouseUsage("102").lngMouseKey Then
            If Abs(y - lngBaseYY) >= lngWidthLevelStep / 5 Or Abs(x - lngBaseXX) >= lngWidthLevelStep / 5 Then  ''''调窗步长控制
                Me.Viewer1.Images(1).width = Me.Viewer1.Images(1).width + (x - lngBaseXX) * lngWidthLevelStep / 5
                Me.Viewer1.Images(1).Level = Me.Viewer1.Images(1).Level + (y - lngBaseYY) * lngWidthLevelStep / 5
                lngMagnifyWidth = Me.Viewer1.Images(1).width
                lngMagnifyLevel = Me.Viewer1.Images(1).Level
                Me.Viewer1.Refresh
                lngBaseXX = x
                lngBaseYY = y
            End If
        Else
            GetCursorPos TmpMp
            Mx = (TmpMp.x * Screen.TwipsPerPixelX) - (BeginWidth * Screen.TwipsPerPixelX)
            My = (TmpMp.y * Screen.TwipsPerPixelY) - (BeginHeight * Screen.TwipsPerPixelY)
            Me.Move Mx, My
            ImgMagnify
            Me.lblZoomState.Caption = "   放大倍数：" & Format(Me.Viewer1.Images(1).ActualZoom, "###0.0000")
            '刷新放大窗体
            f.Refresh
        End If
    End If
End Sub
Private Sub Viewer1_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    FrmClick = False
End Sub
'放大
Public Sub ImgMagnify()
    '****************************************************************************
    '参数：MainFrm 取得父窗体对象
    '      DicView 取得DICOM对象
    '作用：对DICOM图像放大
    '****************************************************************************
    Dim Ox1, Ox2, Oy1, Oy2 As Integer   '用于得到父窗体内控件在屏幕上坐标
    Dim Ix1, Ix2, Iy1, Iy2 As Integer   '用于得到放大窗体在屏幕上坐标
    Dim Dx, Dy As Integer               '放大窗体的DICOM控件的中心点坐标
    Dim HowImg As Integer               '得到当前是第几幅图
    Dim Sx1, Sx2, Sy1, Sy2 As Integer   '小图像的坐标
    Dim intRow, intCol As Integer       '窗体的行和列个数
    Dim MWidth, MHeight As Integer      '缩小后的图像宽和高
    Dim i As Integer                    '临时变量
    Dim ViewIndex As Integer
    Dim a As Double
'    On Error Resume Next
    'Dim A As New DicomImage
    ViewIndex = 0
   '放大窗体位置
    With Me
        Ix1 = (.left / Screen.TwipsPerPixelX) + .Viewer1.left - intMaxLeft
        Ix2 = (.left / Screen.TwipsPerPixelX) + .Viewer1.left + .Viewer1.width - intMaxLeft
        Iy1 = (.top / Screen.TwipsPerPixelY) + .Viewer1.top - GetMenuHeight - intMaxTop - GetSystemMetrics(11) '+ 6
        Iy2 = (.top / Screen.TwipsPerPixelY) + .Viewer1.top + .Viewer1.height - intMaxTop - GetMenuHeight - GetSystemMetrics(11) '+ 6
    End With
    '放大窗体中心点
    Dx = (Ix2 - Ix1) / 2 + Ix1
    Dy = (Iy2 - Iy1) / 2 + Iy1
    
    ViewIndex = FunIsViewer(Dx * Screen.TwipsPerPixelX, Dy * Screen.TwipsPerPixelY)
    
    
    '得到父窗体内DIDOM控件位置
    With f
        Ox1 = (.left / Screen.TwipsPerPixelX) + (.Viewer(ViewIndex).left / Screen.TwipsPerPixelX)
        Ox2 = (.left / Screen.TwipsPerPixelX) + (.Viewer(ViewIndex).left / Screen.TwipsPerPixelX) + (.Viewer(ViewIndex).width / Screen.TwipsPerPixelX)
        Oy1 = (.top / Screen.TwipsPerPixelY) + (.Viewer(ViewIndex).top / Screen.TwipsPerPixelY)
        Oy2 = (.top / Screen.TwipsPerPixelY) + (.Viewer(ViewIndex).top / Screen.TwipsPerPixelY) + (.Viewer(ViewIndex).height / Screen.TwipsPerPixelY)
    End With
    '生成图像
    With Me.Viewer1
        HowImg = f.Viewer(ViewIndex).ImageIndex(Dx - Ox1, Dy - Oy1)
        '加上图像的起终位置(在屏屏上看不见的部份)
        Ox1 = Ox1 - f.Viewer(ViewIndex).Images(HowImg).ScrollX
        Oy1 = Oy1 - f.Viewer(ViewIndex).Images(HowImg).ScrollY
        If HowImg < 1 Then
            .Images.Clear
            Exit Sub
        End If
        '是当前图像时不刷新图像
            .Images.Clear
            .Images.Add f.Viewer(ViewIndex).Images(HowImg)
            .Images(1).MagnificationMode = doFilterBSpline
            .Images(1).Zoom = f.Viewer(ViewIndex).Images(HowImg).ActualZoom * (Me.Slider1.Value / 10)
            .Images(1).StretchToFit = False
            .Images(1).Labels.Clear
            If .Images(1).VOILUT = 1 Then
                .Images(1).width = f.Viewer(ViewIndex).Images(HowImg).width
                .Images(1).Level = f.Viewer(ViewIndex).Images(HowImg).Level
                .Images(1).VOILUT = 0
            End If
            If lngMagnifyWidth <> 0 And lngMagnifyLevel <> 0 Then
                .Images(1).width = lngMagnifyWidth
                .Images(1).Level = lngMagnifyLevel
            End If
    End With
    With f.Viewer(ViewIndex)
        '当图像超过当前可示图像总量时处理
        If (.MultiColumns * .MultiRows) < .Images.Count Then
            i = HowImg - f.Viewer(ViewIndex).CurrentIndex + 1
        Else
            i = HowImg
        End If
        '得到当前图像位置
        If (i Mod .MultiColumns) = 0 Then
            intRow = i / .MultiColumns
            intCol = .MultiColumns
        Else
            intRow = Int(i / .MultiColumns) + 1
            intCol = HowImg Mod .MultiColumns
            If intCol = 0 Then
                intCol = 1
            End If
        End If
        MWidth = (.width / .MultiColumns) / Screen.TwipsPerPixelX
        MHeight = (.height / .MultiRows) / Screen.TwipsPerPixelY
    End With
    '放大小图像位坐标
    If intCol = 1 Then
        Sx1 = 0
        Sx2 = MWidth
    Else
        Sx1 = MWidth * (intCol - 1)
        Sx2 = MWidth * intCol
    End If
    If intRow = 1 Then
        Sy1 = 0
        Sy2 = MHeight
    Else
        Sy1 = MHeight * (intRow - 1)
        Sy2 = MHeight * intRow
    End If

    '计算放大后的位置
    With Viewer1
        If Dx > Ox1 And Dx < Ox2 And Dy > Oy1 And Dy < Oy2 Then
'            .Images(1).ScrollX = ((Dx - Ox1 - Sx1) * (Slider1.Value / 10)) - Abs(f.viewer(ViewIndex).Images(HowImg).ActualScrollX * (Slider1.Value / 10)) - (.width / 2)
'            .Images(1).ScrollY = ((Dy - Oy1 - Sy1) * (Slider1.Value / 10)) - Abs(f.viewer(ViewIndex).Images(HowImg).ActualScrollY * (Slider1.Value / 10)) - (.height / 2)
            .Images(1).ScrollX = (((Dx - Ox1 + Abs(f.Viewer(ViewIndex).Images(HowImg).ScrollX) - Sx1) - Abs(f.Viewer(ViewIndex).Images(HowImg).ActualScrollX)) / f.Viewer(ViewIndex).Images(HowImg).ActualZoom * .Images(1).ActualZoom) - (.width / 2)
            .Images(1).ScrollY = (((Dy - Oy1 + Abs(f.Viewer(ViewIndex).Images(HowImg).ScrollY) - Sy1) - Abs(f.Viewer(ViewIndex).Images(HowImg).ActualScrollY)) / f.Viewer(ViewIndex).Images(HowImg).ActualZoom * .Images(1).ActualZoom) - (.height / 2)
        Else
            Me.Viewer1.Images.Clear
        End If
    End With
    
    '滤镜
    If blnOrganLens Then subOrganLens
    '反白
    If blnInvert Then Call subFlipRotate(Viewer1.Images(1), "Invert")
    
    '增加一个刷新，是图像反映速度加快。
    Me.Viewer1.Refresh
End Sub

'得到当前Viewer数组位置
Function FunIsViewer(x As Long, y As Long) As Integer
    Dim v As DicomViewer

    With f
        FunIsViewer = 0
        For Each v In .Viewer
            If v.Visible And .left + v.left <= x And _
            .left + v.left + v.width >= x And _
            .top + v.top <= y And _
            .top + v.top + v.height >= y Then
                FunIsViewer = v.Index
                Exit Function
            End If
        Next
    End With
End Function

Private Sub subOrganLens()
    If Me.Viewer1.Images.Count = 0 Then Exit Sub
    Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
    Dim ww As Long, wl As Long
    
    '计算组织透镜的覆盖区域
    lngLeft = Me.Viewer1.Images(1).ActualScrollX / Me.Viewer1.Images(1).ActualZoom
    lngTop = Me.Viewer1.Images(1).ActualScrollY / Me.Viewer1.Images(1).ActualZoom
    lngWidth = Me.Viewer1.width / Me.Viewer1.Images(1).ActualZoom
    lngHeight = Me.Viewer1.height / Me.Viewer1.Images(1).ActualZoom
    
    '计算组织透镜的新窗宽窗位，利用自适应调窗算法
    If funAutoWinWL(Me.Viewer1.Images(1), lngLeft, lngTop, lngWidth, lngHeight, ww, wl) Then
        Me.Viewer1.Images(1).width = ww
        Me.Viewer1.Images(1).Level = wl
    End If
End Sub

Private Function GetToolBarBottomOrRight(LeftORTop As Integer) As Long
    '------------------------------------------------
    '功能：                                  得到工具条的Left或Right的高
    '参数：                                  LeftORRight 1=left  2= Right
    '返回：                                  工具条的高
    '上级函数或过程：                        ImgMagnify
    '下级函数或过程：                        无
    '引用的外部参数：                        f主界面窗体
    '编制人：                                曾超 2005-8-9
    '------------------------------------------------
    Dim intMaxTop  As Long                  '最大的高
    Dim intMaxLeft As Long                  '最大的边
    Dim intToolBarLeft  As Long             '工具条Left
    Dim intToolBarTop   As Long             '工具条Top
    Dim intToolBarRight As Long             '工具条Right
    Dim intToolBarBottom As Long            '工具条Bottom
    Dim a As CommandBar
    Dim i As Integer
    
    With f.ComToolBar
        If LeftORTop = 1 Then
            For i = 2 To 8
                If .Item(i).Position = xtpBarLeft Then
                    .Item(i).GetWindowRect intToolBarLeft, intToolBarTop, intToolBarRight, intToolBarBottom
                    If intMaxLeft < intToolBarLeft Or intMaxLeft = 0 Then
                        intMaxLeft = intToolBarLeft
                        GetToolBarBottomOrRight = GetToolBarBottomOrRight + (intToolBarRight - intToolBarLeft)
                    End If
                End If
            Next
            GetToolBarBottomOrRight = GetToolBarBottomOrRight / Screen.TwipsPerPixelX
        Else
            For i = 2 To 8
                If .Item(i).Position = xtpBarTop Then
                    .Item(i).GetWindowRect intToolBarLeft, intToolBarTop, intToolBarRight, intToolBarBottom
                     If intMaxTop < intToolBarTop Or intMaxTop = 0 Then
                        intMaxTop = intToolBarTop
                        GetToolBarBottomOrRight = GetToolBarBottomOrRight + (intToolBarBottom - intToolBarTop)
                    End If
                End If
            Next
            GetToolBarBottomOrRight = GetToolBarBottomOrRight / Screen.TwipsPerPixelY
        End If
    End With
End Function

Private Function GetMenuHeight() As Long
    '------------------------------------------------
    '功能：                                  得到菜单和缩略图的高度
    '参数：
    '返回：                                  菜单和缩略图的高度
    '------------------------------------------------
    Dim lngToolBarLeft  As Long             '工具条Left
    Dim lngToolBarTop   As Long             '工具条Top
    Dim lngToolBarRight As Long             '工具条Right
    Dim lngToolBarBottom As Long            '工具条Bottom
    
    f.ComToolBar.Item(ToolBar_Menu).GetWindowRect lngToolBarLeft, lngToolBarTop, lngToolBarRight, lngToolBarBottom
    GetMenuHeight = (lngToolBarBottom - lngToolBarTop) / Screen.TwipsPerPixelY
    
    '如果缩略图设置为停靠，而且显示了缩略图，则计算缩略图的高度
    If blnDockMiniImage = True And frmMiniSeries.Visible = True Then
        GetMenuHeight = GetMenuHeight + frmMiniSeries.height / Screen.TwipsPerPixelY
    End If
End Function








