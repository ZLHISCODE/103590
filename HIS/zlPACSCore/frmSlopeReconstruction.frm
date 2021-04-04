VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmSlopeReconstruction 
   Caption         =   "MPR斜面重建"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   Icon            =   "frmSlopeReconstruction.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9915
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic 
      Height          =   2775
      Index           =   2
      Left            =   5160
      ScaleHeight     =   2715
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   240
      Width           =   2535
      Begin DicomObjects.DicomViewer Viewer 
         Height          =   1455
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   1215
         _Version        =   262147
         _ExtentX        =   2143
         _ExtentY        =   2566
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin VB.PictureBox pic 
      Height          =   2775
      Index           =   1
      Left            =   960
      ScaleHeight     =   2715
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   480
      Width           =   2535
      Begin DicomObjects.DicomViewer Viewer 
         Height          =   1455
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   1215
         _Version        =   262147
         _ExtentX        =   2143
         _ExtentY        =   2566
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin VB.PictureBox pic 
      Height          =   2775
      Index           =   3
      Left            =   5280
      ScaleHeight     =   2715
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   3240
      Width           =   2535
      Begin DicomObjects.DicomViewer Viewer 
         Height          =   1455
         Index           =   3
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   1215
         _Version        =   262147
         _ExtentX        =   2143
         _ExtentY        =   2566
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin MSComctlLib.ImageList ImageListMouse 
      Left            =   0
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlopeReconstruction.frx":038A
            Key             =   "Stack"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlopeReconstruction.frx":04EC
            Key             =   "Rotate2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlopeReconstruction.frx":0806
            Key             =   "Rotate"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlopeReconstruction.frx":0B20
            Key             =   "WindowWL"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlopeReconstruction.frx":0E3A
            Key             =   "Zoom"
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   960
      Top             =   1800
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSlopeReconstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mParForm As frmViewer
Dim mImageViewer As DicomViewer      '观片窗口中，图像所在的Viewer
Dim mAxialViewer As DicomViewer      '斜面重建窗口中，轴位图像所在的Viewer
Dim mCoronalViewer As DicomViewer     '斜面重建窗口中，冠状位图像所在的Viewer
Dim mSagittalViewer As DicomViewer   '鞋面重建窗口中，矢状位图像所在的Viewer

Dim SelectedImage As DicomImage     '当前选中的图像
Dim SelectedLabel As DicomLabel     '当前选择的标注
Dim blnMoveLabel  As Boolean        '开始移动标注的标识
Dim blnMoveImage As Boolean         '开始拖动鼠标的标识，拖动鼠标，冠状位和矢状位图会改变位置
Dim lngBaseX As Long                '鼠标移动基准位置
Dim lngBaseY As Long                '鼠标移动基准位置
Dim lngBaseCenterX As Long          '控制线的中心点基准位置X
Dim lngBaseCenterY As Long          '控制线的中心点基准位置Y

Dim blnRebuild As Boolean           '是否成功重建了图像

Private Enum RebuildType
    rt翻页 = 0
    rt斜面重建 = 1
    rt矢冠状位平移 = 2
End Enum

Dim mMouseAction As MouseAction
Private Enum MouseAction
    ma默认 = 0
    ma旋转标注 = 1
    ma翻图 = 2
    ma平移标注 = 3
End Enum

Public Function zlShowMe(parForm As frmViewer) As Boolean
    Dim iIndex As Integer
    
    On Error GoTo err
    
    blnRebuild = False
    
    Set mParForm = parForm
    
    If mParForm.intSelectedSerial = 0 Then
        MsgBox "请先选择一个图像序列后，再开始斜面重建。"
        zlShowMe = True
        Exit Function
    End If
    
    '设置Viewer变量们
    Set mImageViewer = mParForm.Viewer(mParForm.intSelectedSerial)
    Set mAxialViewer = Viewer(1)
    Set mCoronalViewer = Viewer(2)
    Set mSagittalViewer = Viewer(3)
    
    '先把序列中的所有图像都加载到Viewer中
    Call funAddAllImages(mImageViewer)
    
    '判断是否满足矢冠状位重建的条件
    If LeagelToACRebuild(mImageViewer.Images) = 1 Then
        zlShowMe = False   '退出重建
        Exit Function
    End If
    
    '初始化窗体
    Call InitForm
    
    '显示中间的图像
    iIndex = mParForm.Viewer(mParForm.intSelectedSerial).Images.Count / 2
    If iIndex <= 0 Then
        iIndex = 1
    End If
    Call ShowAxialImage(mParForm.Viewer(mParForm.intSelectedSerial).Images(iIndex), iIndex)

    '加载重建的图像
    If ShowImage = False Then
        zlShowMe = False
        Exit Function
    Else
        blnRebuild = True
    End If
    
    '先显示窗体，再加载重建的图像
    Me.Show 1, mParForm
    
    zlShowMe = True
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitForm()
'------------------------------------------------
'功能：初始化窗口
'参数：无
'返回：无
'------------------------------------------------
    Dim Pane1 As Pane
    Dim Pane2 As Pane
    Dim Pane3 As Pane
    Dim dGlabal As New DicomGlobal
    
    On Error GoTo err
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .TabPaintManager.BoldSelected = True
        .Options.DefaultPaneOptions = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        Set Pane1 = .CreatePane(1, 200, 200, DockTopOf)
        Set Pane2 = .CreatePane(2, 200, 200, DockBottomOf)
        Set Pane3 = .CreatePane(3, 600, 400, DockLeftOf, Pane1 And Pane2)
    End With
    
    '先创建本次MPR斜面重建的序列UID
    ZLMPRSlopeSeriesUID = dGlabal.NewUID
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Id = 1 Then
        Item.Handle = pic(2).hwnd
    ElseIf Item.Id = 2 Then
        Item.Handle = pic(3).hwnd
    ElseIf Item.Id = 3 Then
        Item.Handle = pic(1).hwnd
    End If
End Sub

Private Function ShowAxialImage(img As DicomImage, imageIndex As Integer) As Boolean
'------------------------------------------------
'功能：显示轴位图和MPR重建结果图
'参数： img -- 添加到轴位图位置的图像
'       imageIndex -- 图像在主窗体Viewer中的ImageIndex
'返回：True--成功；False--失败
'------------------------------------------------
    Dim oldImage As New DicomImage
    Dim blnCopyLabels As Boolean
    
    On Error GoTo err
    
    '如果原来已经有轴位图像，先保存这个图像，后面需要画控制线
    If mAxialViewer.Images.Count = 1 Then
        Set oldImage = mAxialViewer.Images(1)
        blnCopyLabels = True
    Else
        Set oldImage = Nothing
        blnCopyLabels = False
    End If
    
    mAxialViewer.Images.Clear
    mAxialViewer.Images.Add img
    '这里Add图像之后，是新建了一个图像吗？tag改变，会影响原图吗？
    mAxialViewer.Images(1).Tag = imageIndex
    
    '初始化轴位图像的MPR重建控制线
    Call funInitMPRControlLines(mAxialViewer.Images(1), False)
    
    '复制标注
    If blnCopyLabels = True Then
        Call funCopyMPRControlLines(mAxialViewer.Images(1), oldImage)
    End If
    
    ShowAxialImage = True
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 
Private Function ShowImage() As Boolean
'------------------------------------------------
'功能：显示轴位图和MPR重建结果图
'参数：无
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    On Error GoTo err
    
    '冠状位重建，右上角Viewer(2)
    ShowImage = funMPRslope(mImageViewer, mAxialViewer, mCoronalViewer, mSagittalViewer, mParForm)
    
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim lngResult As Long
    Dim resImage As DicomImage
    
    If blnRebuild = True Then
        lngResult = MsgBox("是否保存重建结果？", vbYesNoCancel, "提示信息", Me)
        If lngResult = vbCancel Then
            Cancel = -1
            Exit Sub
        ElseIf lngResult = vbYes Then
            '保存重建结果图
            Set resImage = mAxialViewer.Images(1)
            resImage.SeriesUID = ZLMPRSlopeSeriesUID
            '保存结果图
            Call subSaveImage(resImage, mImageViewer.Images(1).SeriesUID)
            '把图像追加到观片站中
            Call subOpenCurrentImage(mParForm, resImage)
        End If
        
        Set mImageViewer = Nothing
        Set mAxialViewer = Nothing
        Set mCoronalViewer = Nothing
        Set mSagittalViewer = Nothing
    End If
End Sub

Private Sub pic_Resize(Index As Integer)
'------------------------------------------------
'功能：picture背景改变大小，同时改变所有Viewer的大小
'参数：无
'返回：无
'------------------------------------------------
    If Viewer.Count = 3 Then
        Viewer(Index).left = 0
        Viewer(Index).top = 0
        Viewer(Index).width = pic(Index).width
        Viewer(Index).height = pic(Index).height
    End If
End Sub

Private Sub Viewer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
'------------------------------------------------
'功能：MouseDown，选择当前图像，当前标注
'参数：无
'返回：无
'------------------------------------------------
    Dim ls As DicomLabels
    Dim j As Integer
    Dim m As Integer
    
    On Error GoTo err
    
    ''记录鼠标的基准位置
    lngBaseX = Viewer(Index).ImageXPosition(x, y)
    lngBaseY = Viewer(Index).ImageYPosition(x, y)
    
    mMouseAction = ma默认
    MousePointer = vbDefault
    
    '鼠标单击时，先选择图像
    If Viewer(Index).Images.Count <= 0 Then
        Set SelectedImage = Nothing
        Exit Sub
    Else
        Set SelectedImage = Viewer(Index).Images(1)
    End If
    
    If Button = 1 Then  '鼠标左键
        Set ls = Viewer(Index).LabelHits(x, y, False, False, True)
        If ls.Count > 0 Then
            
            If Index = 1 And SelectedImage.Labels.IndexOf(ls(1)) >= G_INT_SYS_LABEL_MPRV And SelectedImage.Labels.IndexOf(ls(1)) <= G_INT_SYS_LABEL_MPR_POINT_O Then
                '是轴位图上的“MPR重建”相关的标注
                For j = 1 To ls.Count
                    If SelectedImage.Labels.IndexOf(ls(j)) > m Then m = SelectedImage.Labels.IndexOf(ls(j))
                Next
                Set SelectedLabel = SelectedImage.Labels(m)     'm为序号最大的标注。
            ElseIf (Index = 2 Or Index = 3) And ((SelectedImage.Labels.IndexOf(ls(1)) = G_INT_SYS_LABEL_MPR_RESULT_H) _
                Or (SelectedImage.Labels.IndexOf(ls(1)) = G_INT_SYS_LABEL_MPR_RESULT_V)) Then
                '是“矢冠状重建”结果图中的横线和竖线
                Set SelectedLabel = ls(1)
                
                lngBaseCenterX = SelectedLabel.left + SelectedLabel.width / 2
                lngBaseCenterY = SelectedLabel.top + SelectedLabel.height / 2
        
                '只有横线可以旋转，改编鼠标形状
                If SelectedImage.Labels.IndexOf(ls(1)) = G_INT_SYS_LABEL_MPR_RESULT_H Then
                    If (Abs(SelectedLabel.width) > Abs(SelectedLabel.height) And _
                        (lngBaseX < SelectedLabel.left + SelectedLabel.width / 8 Or lngBaseX > SelectedLabel.left + SelectedLabel.width * 7 / 8)) _
                        Or (Abs(SelectedLabel.height) > Abs(SelectedLabel.width) And _
                        (lngBaseY < SelectedLabel.top + SelectedLabel.height / 8 Or lngBaseY > SelectedLabel.top + SelectedLabel.height * 7 / 8)) Then
                        MousePointer = vbCustom ' vbNoDrop 'vbCustom
                        MouseIcon = ImageListMouse.ListImages("Rotate").Picture
                        mMouseAction = ma旋转标注
                    Else
                        MousePointer = vbSizeAll
                        mMouseAction = ma平移标注
                    End If
                Else
                    MousePointer = vbSizeAll
                    mMouseAction = ma平移标注
                End If
            End If
            
            blnMoveLabel = True
        Else
            '没有选中标注，则是移动图像
            blnMoveImage = True
            MousePointer = vbCustom
            MouseIcon = ImageListMouse.ListImages("Stack").Picture
            mMouseAction = ma翻图
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Viewer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
'------------------------------------------------
'功能：移动标注，重建
'参数：无
'返回：无
'------------------------------------------------
    If SelectedImage Is Nothing Then Exit Sub
    
    On Error GoTo err

    '移动标注
    If blnMoveLabel And Not SelectedLabel Is Nothing Then
        subaCorrectCursor Viewer(Index), SelectedImage, x, y     ''''鼠标移动如果超出图像范围则修正其鼠标位置
        
        '移动标注的所有操作，包括移动MPR线并显示重建结果图
        subMoveSlopeLabel SelectedLabel, Viewer(Index).ImageXPosition(x, y), _
            Viewer(Index).ImageYPosition(x, y), lngBaseX, lngBaseY, Index
        
        lngBaseX = Viewer(Index).ImageXPosition(x, y)
        lngBaseY = Viewer(Index).ImageYPosition(x, y)
        
        Viewer(Index).Refresh
    ElseIf blnMoveImage = True Then
        subaCorrectCursor Viewer(Index), SelectedImage, x, y     ''''鼠标移动如果超出图像范围则修正其鼠标位置
        '拖拽鼠标的时候，切换轴位，冠状位和矢状位图像的位置
        subMoveImage Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y), lngBaseX, lngBaseY, Index
        
        lngBaseX = Viewer(Index).ImageXPosition(x, y)
        lngBaseY = Viewer(Index).ImageYPosition(x, y)
        
        Viewer(Index).Refresh
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub subMoveSlopeLabel(la As DicomLabel, newX As Long, newY As Long, _
    baseX As Long, baseY As Long, ViewerIndex As Integer)
'------------------------------------------------
'功能：移动一个标注，包括轴位图，矢状位和冠状位的控制线
'参数： la -- 被移动的标注
'       newX -- 新位置的图像像素X坐标
'       newY -- 新位置的图像像素Y坐标
'       basex -- 旧位置的图像像素X坐标
'       baseY -- 旧位置的图像像素Y坐标
'       ViewerIndex -- 图像所在Viewer的Index；1-轴位图；2-冠状位图；3-矢状位图
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    '在轴位图像上，移动控制线
    If SelectedImage.Labels.IndexOf(la) >= G_INT_SYS_LABEL_MPRV And SelectedImage.Labels.IndexOf(la) <= G_INT_SYS_LABEL_MPR_POINT_O Then ''[矢冠状线的移动]
        '移动矢冠状重建控制点、线，且生成新的重建图像。
        Call subMoveAxialMPRLabel(la, newX, newY, baseX, baseY)
    ElseIf (SelectedImage.Labels.IndexOf(la) = G_INT_SYS_LABEL_MPR_RESULT_H) _
        Or (SelectedImage.Labels.IndexOf(la) = G_INT_SYS_LABEL_MPR_RESULT_V) Then
        '在冠状位或矢状位图像上，移动控制线
        Call subMoveCAndSLabel(la, SelectedImage, newX, newY, baseX, baseY, IIf(ViewerIndex = 2, True, False))
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub subMoveAxialMPRLabel(la As DicomLabel, xx As Long, Yy As Long, _
    baseX As Long, baseY As Long)
'------------------------------------------------
'功能：移动轴位图矢冠状重建控制点、线，且生成新的重建图像。
'参数：
'       la -- 被移动的矢冠状重建控制点或控制线；
'       xx -- 标注新位置在图像上的X坐标；
'       yy -- 标注新位置在图像上的Y坐标；
'       basex -- 标注旧位置的图像上的x坐标；
'       baseY -- 标注旧位置的图像上的y坐标。
'返回：无，直接移动矢冠状重建的控制点和线，并生成重建结果图像。
'------------------------------------------------
    Dim intIndex As Integer
    Dim axialImage As DicomImage
    
    On Error GoTo err
    
    '不能直接使用SelectedImage，调用这个方法的时候，SelectedImage可能是冠状位或矢状位图
    Set axialImage = mAxialViewer.Images(1)
    
    ''''''''''''''''''''''[是四角点的移动]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    intIndex = axialImage.Labels.IndexOf(la)
    
    '矢冠状控制点中四个边点的移动处理
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_V1 Then
        Call subPeriodMove(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), xx, Yy, _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPRV), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), axialImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_H1 Then
        Call subPeriodMove(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), xx, Yy, _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPRH), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), axialImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_V2 Then
        Call subPeriodMove(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), xx, Yy, _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPRV), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), axialImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_H2 Then
        Call subPeriodMove(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), xx, Yy, _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPRH), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), axialImage)
    '''''''''''''''''''''''''''中心点的移动'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_O Then
        axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top + Yy - baseY
        axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left + xx - baseX
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top < -G_INT_MPR_RADIUS / 2 + 1 Then
            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = -G_INT_MPR_RADIUS / 2 + 1
        End If
        If axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left < -G_INT_MPR_RADIUS / 2 + 1 Then
            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = -G_INT_MPR_RADIUS / 2 + 1
        End If
        If axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top > axialImage.sizeY - G_INT_MPR_RADIUS / 2 - 1 Then
            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = axialImage.sizeY - G_INT_MPR_RADIUS - 1
        End If
        If axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left > axialImage.sizeX - G_INT_MPR_RADIUS / 2 - 1 Then
            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = axialImage.sizeX - G_INT_MPR_RADIUS - 1
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '矢冠状中心点的移动
        If xx <> baseX Then
            Call subPeriodMovee5X(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), _
                                axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), _
                                axialImage.Labels(G_INT_SYS_LABEL_MPRV), xx, Yy, _
                                axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), baseX, baseY, axialImage)
        End If
        
        If Yy <> baseY Then
            Call subPeriodMovee5X(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), _
                                axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), _
                                axialImage.Labels(G_INT_SYS_LABEL_MPRH), xx, Yy, _
                                axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), baseX, baseY, axialImage)
        End If
    End If
    
    '旋转标注后，刷新图像显示
    Call axialImage.Refresh(False)
    
    ''''''''''进行重建''''''''''''''''''''''''''''''''''''''''''''''
    '标注是MPR控制线竖线的两个端点，或者是MPR控制线竖线的中心点，此时需要移动的是MPR竖线
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_V1 Or intIndex = G_INT_SYS_LABEL_MPR_POINT_V2 _
        Or (intIndex = G_INT_SYS_LABEL_MPR_POINT_O And xx <> baseX) Then
        
        '根据传入的控制线，对MPRViewer内的图像进行重建，并将返回重建的结果图显示到ShowViewerIndex指定的Viewer中
        If funGetCandSImageAndShow(axialImage.Labels(G_INT_SYS_LABEL_MPRV), mImageViewer, _
                                        mAxialViewer, mSagittalViewer, ToltalHeight, 1, False, True) = False Then
            '重建出错，退出MPR重建
            Exit Sub
        End If
    End If
    
    '标注是MPR控制线横线的两个端点，或者是MPR控制线的横线中心点，此时需要移动的是MPR横线
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_H1 Or intIndex = G_INT_SYS_LABEL_MPR_POINT_H2 _
        Or (intIndex = G_INT_SYS_LABEL_MPR_POINT_O And Yy <> baseY) Then
        
        '根据传入的控制线，对MPRViewer内的图像进行重建，并将返回重建的结果图显示到ShowViewerIndex指定的Viewer中
        If funGetCandSImageAndShow(axialImage.Labels(G_INT_SYS_LABEL_MPRH), mImageViewer, _
                                        mAxialViewer, mCoronalViewer, ToltalHeight, 2, False, True) = False Then
            '重建出错，退出MPR重建
            Exit Sub
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub subMoveCAndSLabel(la As DicomLabel, img As DicomImage, newX As Long, newY As Long, _
    baseX As Long, baseY As Long, blnIsCoronal As Boolean)
'------------------------------------------------
'功能：移动矢冠状重建冠状位和矢状位图的控制线，横线控制轴位图像的自动翻页，竖线控制结果图
'参数：
'       la--被移动的矢冠状重建控制点或控制线；
'       img -- 标注所在的图像
'       newX--新位置的图像像素X坐标；
'       newY--新位置的图像像素Y坐标；
'       basex--旧位置的图像像素x坐标；
'       baseY--旧位置的图像像素y坐标
'       blnIsCoronal -- 是否冠状位，True-冠状位；False-矢状位
'返回：无，直接移动矢冠状重建的结果线
'------------------------------------------------
    Dim iImageIndex As Integer
    Dim intIndex As Integer
    Dim rtType As RebuildType
    Dim k As Double
    Dim laAxial As DicomLabel
    
    On Error GoTo err
    
    '在冠状位和矢状位上移动控制线，SelectedImage是冠状位或矢状位图
    '将标注线的移动和图像重建操作分开，不同的标注移动方法，创建方式不同
    
    intIndex = img.Labels.IndexOf(la)
    
    If intIndex = G_INT_SYS_LABEL_MPR_RESULT_H And (mMouseAction = ma平移标注 Or mMouseAction = ma翻图) Then    '横线平移，轴位图更换图像或者斜面重建
        '先判断是更换轴位图像，还是斜面重建
        If la.height = 0 Then
            If (blnIsCoronal = True And mSagittalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height = 0) _
                Or (blnIsCoronal = False And mCoronalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height = 0) Then
                rtType = rt翻页
            Else
                rtType = rt斜面重建
            End If
        Else
            rtType = rt斜面重建
        End If

        '移动标注线
        If la.width <> 0 Then
            '计算斜率
            k = la.height / la.width
            If k = 0 Then
                '斜率为0 ，直接移动Y坐标
                If la.top + (newY - baseY) > 0 And la.top + (newY - baseY) < img.sizeY Then
                    la.top = la.top + (newY - baseY)
                End If
            Else
                Call funGetLine(la, img, k, la.left + (newX - baseX), la.top + (newY - baseY))
            End If
        End If
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_RESULT_H And mMouseAction = ma旋转标注 Then  '横线旋转，斜面重建
        '旋转横向控制线
        Call subRotateLabel(la, newX, newY, img)
        
        rtType = rt斜面重建
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_RESULT_V Then '竖线平移，是重建图像
        '先移动矢冠状重建结果线
        If la.left + (newX - baseX) > 0 And la.left + (newX - baseX) < img.sizeX Then
            la.left = la.left + (newX - baseX)
        End If
        
        rtType = rt矢冠状位平移
    End If
    
    If rtType = rt翻页 Then
        '根据结果线的位置，显示新的轴位图像
        iImageIndex = la.top / img.sizeY * mImageViewer.Images.Count
        If iImageIndex > 0 And iImageIndex <= mImageViewer.Images.Count Then
            '将对应的图像，显示到轴位图中
            
            Call ShowAxialImage(mParForm.Viewer(mParForm.intSelectedSerial).Images(iImageIndex), iImageIndex)
            
            '调整另一个重建图的标注位置
            If blnIsCoronal = True Then
                Call subMPRSlopeDrawResultControlLabels(la, mSagittalViewer.Images(1), mImageViewer, mAxialViewer)
            Else
                Call subMPRSlopeDrawResultControlLabels(la, mCoronalViewer.Images(1), mImageViewer, mAxialViewer)
            End If
        End If
    ElseIf rtType = rt斜面重建 Then
        '如果旋转或平移了冠状位（矢状位）的横向控制线，就需要调整矢状位（冠状位）图控制线的位置
        '保持矢状位图和冠状位图控制线的中心点重合，两根线在同一个平面上
        Call funTranslateLabel(blnIsCoronal)
        '斜面重建
        Call funGetSlopeImageAndShow
    ElseIf rtType = rt矢冠状位平移 Then
        '如果平移了冠状位（矢状位）的竖向控制线，就需要调整矢状位（冠状位）图控制线的位置
        '保持矢状位图和冠状位图控制线的中心点重合，两根线在同一个平面上
        Call funTranslateLabel(blnIsCoronal)
        '重新做冠状位或矢状位的重建
        If blnIsCoronal = True Then
            Set laAxial = mImageViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRV)
            laAxial.left = la.left
            laAxial.top = 0
            laAxial.width = 0
            laAxial.height = mImageViewer.Images(1).sizeY
            Call funGetCandSImageAndShow(laAxial, mImageViewer, mAxialViewer, mSagittalViewer, ToltalHeight, 2, False, True)
        Else
            Set laAxial = mImageViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRH)
            laAxial.left = 0
            laAxial.top = la.left
            laAxial.width = mImageViewer.Images(1).sizeX
            laAxial.height = 0
            Call funGetCandSImageAndShow(laAxial, mImageViewer, mAxialViewer, mCoronalViewer, ToltalHeight, 1, False, True)
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Viewer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    blnMoveLabel = False
    blnMoveImage = False
    MousePointer = vbDefault
    mMouseAction = ma默认
End Sub

Private Sub subRotateLabel(la As DicomLabel, xNew As Long, yNew As Long, im As DicomImage)
'------------------------------------------------
'功能：以线的中心点为轴心，旋转冠状位和矢状位图上的横线
'参数：
'       la--被旋转的矢冠状位图中的控制线；
'       xNew--新位置在图像中的X坐标；
'       yNew--新位置在图像中的Y坐标；
'       im--控制线所在的图像；
'返回：无，直接旋转标注
'------------------------------------------------
    Dim x0 As Double, y0 As Double  '直线中心点坐标
    Dim k As Double             '直线的斜率
    
    On Error GoTo err
    
    '中心点位置
    x0 = lngBaseCenterX
    y0 = lngBaseCenterY
    
    '计算斜率
    k = (yNew - y0) / (xNew - x0)
    
    '根据斜率和一个点，画出标注线
    Call funGetLine(la, im, k, x0, y0)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function funSlopeRebuild() As DicomImage
'------------------------------------------------
'功能：以冠状位和矢状位图上的横线为基准，在轴位图中进行斜面重建
'参数：
'返回：无，返回重建后的图像
'------------------------------------------------
    Dim resImage As DicomImage
    Dim X1 As Double, Y1 As Double, Z1 As Double
    Dim X2 As Double, Y2 As Double, Z2 As Double
    Dim X3 As Double, Y3 As Double, Z3 As Double
    Dim X4 As Double, Y4 As Double, Z4 As Double
    Dim A As Double, B As Double, C  As Double, D As Double
    Dim zIndex As Long                  '斜面图像在z方向的坐标
    Dim ZZ As Double
    Dim sizeX As Long, sizeY As Long    '新图像的宽度和高度
    Dim i As Integer, j As Integer
    Dim v As Variant
    Dim lines() As Integer              '保存图像灰度值的二维数组
    Dim zIndexOld As Long               '保存上次读取图像的位置
    
    '算法：
    '将已知三个点的坐标分别用P1(x1,y1,z1)，P2(x2,y2,z2)，P3(x3,y3,z3)表示。（P1，P2，P3不在同一条直线上。）
    '设通过P1，P2，P3三点的平面方程为A(x - x1) + B(y - y1) + C(z - z1) = 0 。
    '化简为一般式：Ax + By + Cz + D = 0。
    '将P1(x1,y1,z1)点数值代入方程Ax + By + Cz + D = 0。
    '即可得到：Ax1 + By 1+ Cz1 + D = 0。
    '化简得D = -(A * x1 + B * y1 + C * z1)。
    '则可以根据P1(x1,y1,z1)，P2(x2,y2,z2)，P3(x3,y3,z3)三点坐标分别求得A、B、C的值，如下：
    'A = (y3 - y1)*(z3 - z1) - (z2 -z1)*(y3 - y1);
    'B = (x3 - x1)*(z2 - z1) - (x2 - x1)*(z3 - z1);
    'C = (x2 - x1)*(y3 - y1) - (x3 - x1)*(y2 - y1);
    '又D = -(A * x1 + B * y1 + C * z1)，所以可以求得D的值。
    '将求得的A、B、C、D值代入一般式方程就可得过P1，P2，P3的平面方程:
    'Ax + By + Cz + D = 0 (一般式)
    
    '冠状位图中，横向控制线在三维体数据中的两个交点为P1，P2，计算P1(X1,Y1,Z1),P2(X2,Y2,Z2)
    '矢状位图中，横向控制线在三维体数据中的两个交点为P3，P4，计算P3(X3,Y3,Z3),P4(X4,Y4,Z4)
    
    On Error GoTo err
    
    Set funSlopeRebuild = Nothing
    zIndexOld = -1
    
    '直接从冠状位和矢状位图和横线，获取到AB,CD的坐标
    If funGetTwoPointsFromImg(mCoronalViewer.Images(1), X1, Y1, Z1, X2, Y2, Z2) = False Then
        Exit Function
    End If
    If funGetTwoPointsFromImg(mSagittalViewer.Images(1), X3, Y3, Z3, X4, Y4, Z4) = False Then
        Exit Function
    End If

    '使用三点确定一个平面，所以P1P2和P3P4之间的交点不是我们预计的中心点，关系也不大，通过P1,P2,P3也可以得到这个位置附近的斜面图
    '求平面方程的A,B,C,D，使用已知四个点中的三个，P1(X1,Y1,Z1),P2(X2,Y2,Z2),P3(X3,Y3,Z3)
    
'    '这个算法有错误？
'    ''A = (y3 - y1)*(z3 - z1) - (z2 -z1)*(y3 - y1);
'    A = (Y3 - Y1) * (Z3 - Z1) - (Z2 - Z1) * (Y3 - Y1)
'    ''B = (x3 - x1)*(z2 - z1) - (x2 - x1)*(z3 - z1);
'    B = (X3 - X1) * (Z2 - Z1) - (X2 - X1) * (Z3 - Z1)
'    ''C = (x2 - x1)*(y3 - y1) - (x3 - x1)*(y2 - y1);
'    C = (X2 - X1) * (Y3 - Y1) - (X3 - X1) * (Y2 - Y1)
'    ''D = -(A * x1 + B * y1 + C * z1)
'    D = -(A * X1 + B * Y1 + C * Z1)
    
    '另一种计算ABCD的方法，a=y1z2-y1z3-y2z1+y2z3+y3z1-y3z2,b=-x1z2+x1z3+x2z1-x2z3-x3z1+x3z2,
    'c=x1y2-x1y3-x2y1+x2y3+x3y1-x3y2,d=-x1y2z3+x1y3z2+x2y1z3-x2y3z1-x3y1z2+x3y2z1
    A = Y1 * Z2 - Y1 * Z3 - Y2 * Z1 + Y2 * Z3 + Y3 * Z1 - Y3 * Z2
    B = -X1 * Z2 + X1 * Z3 + X2 * Z1 - X2 * Z3 - X3 * Z1 + X3 * Z2
    C = X1 * Y2 - X1 * Y3 - X2 * Y1 + X2 * Y3 + X3 * Y1 - X3 * Y2
    D = -X1 * Y2 * Z3 + X1 * Y3 * Z2 + X2 * Y1 * Z3 - X2 * Y3 * Z1 - X3 * Y1 * Z2 + X3 * Y2 * Z1
    
    If C = 0 Then
        Exit Function
    End If
    
    '提取重建结果图的体数据，在重建斜面中，从原点开始，逐个点提取像素值
    sizeX = mImageViewer.Images(1).sizeX
    sizeY = mImageViewer.Images(1).sizeY
    
    '重新定义原图图像灰度值二维数组
    ReDim lines(sizeX, sizeY) As Integer
    
    '利用平面的一般式方程，获取Z坐标 Ax + By + Cz + D = 0
    'z = (-D-Ax-By)/C
    If SafeArrayGetDim(aPixels) = 0 Then
        'MPR的缓存三维体数据维度=0，说明超出内存许可，将直接使用图像数据做重建，图像越多，重建越慢
        For i = 1 To sizeX
            For j = 1 To sizeY
                ZZ = (-D - A * i - B * j) / C   '总高度上的z坐标
                
                '将zIndex换算成图像号
                zIndex = mImageViewer.Images.Count / ToltalHeight * ZZ
                
                If zIndex < 1 Or zIndex > mImageViewer.Images.Count Then
                    lines(i, j) = 0
                Else
                    If zIndexOld <> zIndex Then
                        v = mImageViewer.Images(zIndex).Pixels
                        zIndexOld = zIndex
                    End If
                    lines(i, j) = v(i, j, 1)
                End If
            Next j
        Next i
    Else
        '使用三维数组保存三维体数据，做MPR重建，每次重建速度在1秒钟内
        For i = 1 To sizeX
            For j = 1 To sizeY
                ZZ = (-D - A * i - B * j) / C   '总高度上的z坐标
                
                '将zIndex换算成图像号
                zIndex = mImageViewer.Images.Count / ToltalHeight * ZZ
                
                If zIndex < 1 Or zIndex > mImageViewer.Images.Count Then
                    lines(i, j) = 0
                Else
                    lines(i, j) = aPixels(i, j, zIndex)
                End If
            Next j
        Next i
    End If
    
    '平滑处理，大于512的图像，重建速度慢，精度高，不做平滑
    If sizeX <= 512 Then
        Call funImageSmoothing(lines(), 1)
    End If
    
    '生成新图像
    Set resImage = mImageViewer.Images(1).SubImage(0, 0, sizeX, sizeY, 1, 1)
    
    '删掉一些无用的位置属性
    resImage.Attributes.Remove &H18, &H50
    resImage.Attributes.Remove &H18, &H1110
    resImage.Attributes.Remove &H18, &H1111
    resImage.Attributes.Remove &H18, &H1120     'Tilt
    resImage.Attributes.Remove &H18, &H1140     'Rotation Direction
    resImage.Attributes.Remove &H18, &H5100     'Patient Position
    resImage.Attributes.Remove &H20, &H32       'Image Position(Patient)
    resImage.Attributes.Remove &H20, &H37       'Image Orientation (Patient)
    resImage.Attributes.Remove &H20, &H1041     'Slice Location
    
    '设置结果图的属性，这些不需要修改
'    resImage.Attributes.Add &H28, &H10, intToltalHeight
'    resImage.Attributes.Add &H28, &H11, iPointsCount
'    If intType = 1 Then
'        resImage.Attributes.Add &H20, &H11, LineLong(1).y
'    Else
'        resImage.Attributes.Add &H20, &H11, LineLong(1).x
'    End If
'    resImage.Attributes.Add &H20, &H13, intType
    resImage.Pixels = lines
    resImage.width = mImageViewer.Images(1).width
    resImage.Level = mImageViewer.Images(1).Level
    
    '返回结果图
    Set funSlopeRebuild = resImage
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funGetTwoPointsFromImg(im As DicomImage, ByRef X1 As Double, _
    ByRef Y1 As Double, ByRef Z1 As Double, ByRef X2 As Double, ByRef Y2 As Double, _
    ByRef Z2 As Double) As Boolean
'------------------------------------------------
'功能：获取冠状位图和矢状位图上面，横线在三维体数据中的坐标
'参数： im -- 冠状位或矢状位图
'       x1,y1,z1 -- 第一个点的坐标
'       x2,y2,z2 -- 第二个点的坐标
'返回：true -- 成功； false -- 失败
'------------------------------------------------
    Dim isCoronal As Boolean
    Dim la As DicomLabel
    Dim AX As Double, AY  As Double, BX  As Double, BY  As Double '图像上标注线AB两个端点的坐标
    Dim blnDoublePos As Boolean
    Dim Poss() As String
    
    On Error GoTo err
        
    '先判断是冠状位还是矢状位
    If im.Attributes(&H20, &H13).Value = 1 Then
        isCoronal = True
    Else
        isCoronal = False
    End If
    
    '提取冠状位和矢状位图中的横向控制线
    Set la = im.Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
    
    '计算la在三维体数据中的坐标
    '三维体数据，从头向脚看，左上角为原点（0,0,0），即轴位图像左上角点为原点（0,0）
    '计算la在图像上两个端点的坐标，不考虑两个点的位置顺序
    blnDoublePos = False
    If la.Tag <> "" Then
        Poss = Split(la.Tag, ":")
        If UBound(Poss) = 3 Then
            AX = CDbl(Poss(0))
            AY = CDbl(Poss(1))
            BX = CDbl(Poss(0)) + CDbl(Poss(2))
            BY = CDbl(Poss(1)) + CDbl(Poss(3))
            blnDoublePos = True
        End If
    End If
    If blnDoublePos = False Then
        AX = la.left
        AY = la.top
        BX = la.left + la.width
        BY = la.top + la.height
    End If
    
    '二维坐标，转换成三维坐标
    If isCoronal = True Then
        X1 = AX
        Y1 = im.Attributes(&H20, &H11).Value
        Z1 = AY
        X2 = BX
        Y2 = Y1
        Z2 = BY
    Else
        X1 = im.Attributes(&H20, &H11).Value
        Y1 = AX
        Z1 = AY
        X2 = X1
        Y2 = BX
        Z2 = BY
    End If
    
    funGetTwoPointsFromImg = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funGetSlopeImageAndShow() As Boolean
'------------------------------------------------
'功能： 画斜面重建图像，对imageViewer中的图像进行重建，并将结果图显示到mAxialViewer中
'参数：
'返回:无
'------------------------------------------------
    Dim resImage As DicomImage
    
    On Error GoTo err
    
    '获取重建结果图
    Set resImage = funSlopeRebuild()
    
    '显示结果图
    If resImage Is Nothing Then
        funGetSlopeImageAndShow = False
        Exit Function
    Else
        '将结果图添加到轴位图像位置
        mAxialViewer.Images.Clear
        mAxialViewer.Images.Add resImage
        If mAxialViewer.Images(1).Labels.Count = 0 Then
            Call subInitAImage(mAxialViewer.Images(1), 0, mAxialViewer)
        End If
    End If
    
    funGetSlopeImageAndShow = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funTranslateLabel(isCoronal As Boolean) As Boolean
'------------------------------------------------
'功能： 平移横向控制线，确保冠状位图和矢状位图的横向控制线中心点重合
'       这样他们两条控制线才可以确定一个重建的斜面
'参数：
'返回:无
'------------------------------------------------
    Dim lblRotate As DicomLabel         '旋转的标注
    Dim lblTranslate As DicomLabel      '平移的标注
    Dim imgTranslate As DicomImage      '平移标注的图像
    Dim x0 As Double, y0 As Double      '旋转后标注的中心点
    Dim xT0 As Double, yT0 As Double    '平移前标注的中心点
    Dim k As Double                     '斜率
    
    On Error GoTo err
    
    If isCoronal = True Then    '旋转了冠状位的横向控制线，就平移矢状位的控制线
        Set lblRotate = mCoronalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
        Set lblTranslate = mSagittalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
        Set imgTranslate = mSagittalViewer.Images(1)
        x0 = mCoronalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left
    Else    '旋转了矢状位的横向控制线，就平移冠状位的控制线
        Set lblRotate = mSagittalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
        Set lblTranslate = mCoronalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
        Set imgTranslate = mCoronalViewer.Images(1)
        x0 = mSagittalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left
    End If
    
    '计算平移标注所在平面的交点位置
    
    xT0 = lblTranslate.left + lblTranslate.width / 2
    yT0 = lblTranslate.top + lblTranslate.height / 2
    
    If lblRotate.width = 0 Then
        x0 = xT0
        y0 = yT0
    Else
        y0 = (x0 - lblRotate.left) * lblRotate.height / lblRotate.width + lblRotate.top
    End If
        
    If xT0 = x0 And yT0 = y0 Then
        '不用处理
    Else
        If lblTranslate.width = 0 Then  '是直线
            lblTranslate.left = x0
            lblTranslate.top = 0
            lblTranslate.width = 0
            lblTranslate.height = imgTranslate.sizeY
            lblTranslate.Tag = x0 & ":0:0:" & imgTranslate.sizeY
        ElseIf lblTranslate.height = 0 Then '是横线
            lblTranslate.left = 0
            lblTranslate.top = y0
            lblTranslate.width = imgTranslate.sizeX
            lblTranslate.height = 0
            lblTranslate.Tag = "0:" & y0 & ":" & imgTranslate.sizeY & ":0"
        Else
            '先计算斜率
            k = lblTranslate.height / lblTranslate.width
            Call funGetLine(lblTranslate, imgTranslate, k, x0, y0)
        End If
        imgTranslate.Refresh (False)
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funGetLine(la As DicomLabel, im As DicomImage, k As Double _
    , x0 As Double, y0 As Double) As Boolean
'------------------------------------------------
'功能： 根据斜率和一个点，画出标注线
'参数：  la -- 标注线
'       im -- 标注线所在的图像
'       k -- 斜率
'       x0，y0 -- 直线上一个点的坐标
'返回: 直接画标注线
'------------------------------------------------
    Dim xA As Double, yA As Double  '直线AB的A点坐标
    Dim xB As Double, yB As Double  '直线AB的B点坐标
    Dim xC As Double, yC As Double
    Dim xD As Double, yD As Double
    Dim xAA As Double, yAA As Double
    Dim xBB As Double, yBB As Double
    Dim strTag As String
    Dim intPoints As Integer
    
    On Error GoTo err
    
    If k = 0 Then Exit Function
    
    '算法：
    '设这两点坐标分别为（x1，y1）（x2，y2）,斜截式
    '求斜率: k = (y2 - y1) / (x2 - x1)
    '直线方程 y - y1 = k(x - x1)
    '再把k代入y-y1=k(x-x1)即可得到直线方程。
    '实际上是计算一条直线和矩形的交点
    '设矩形为EFGH,分别计算跟四条边线的交点
    'E-------F
    '|       |
    'G-------H
    
    intPoints = 0
    
    '和EG的交点
    xA = 0
    yA = k * (xA - x0) + y0
    If yA < 0 Then
        yA = 0
        xA = (yA - y0) / k + x0
    ElseIf yA > im.sizeY Then
        yA = im.sizeY
        xA = (yA - y0) / k + x0
    End If
    If xA >= 0 And xA <= im.sizeX Then
        intPoints = 1
        xAA = xA
        yAA = yA
    End If
    
    '和FH的交点
    xB = im.sizeX
    yB = k * (xB - x0) + y0
    If yB < 0 Then
        yB = 0
        xB = (yB - y0) / k + x0
    ElseIf yB > im.sizeY Then
        yB = im.sizeY
        xB = (yB - y0) / k + x0
    End If
    If xB >= 0 And xB <= im.sizeX Then
        If intPoints = 1 Then
            xBB = xB
            yBB = yB
            intPoints = 2
        Else
            xAA = xB
            yAA = yB
            intPoints = 1
        End If
    End If
    
    '和EF的交点
    If intPoints < 2 Then
        yC = 0
        xC = k * (xC - x0) + y0
        If xC < 0 Then
            xC = 0
            yC = k * (xC - x0) + y0
        ElseIf xC > im.sizeX Then
            xC = im.sizeX
            yC = k * (xC - x0) + y0
        End If
        If yC >= 0 And yC <= im.sizeY Then
            If intPoints = 1 Then
                xBB = xC
                yBB = yC
                intPoints = 2
            Else
                xAA = xC
                yAA = yC
                intPoints = 1
            End If
        End If
    End If
    
    '和GH的交点
    If intPoints < 2 Then
        yD = im.sizeY
        xD = k * (xD - x0) + y0
        If xD < 0 Then
            xD = 0
            yD = k * (xD - x0) + y0
        ElseIf xD > im.sizeX Then
            xD = im.sizeX
            yD = k * (xD - x0) + y0
        End If
        If yD >= 0 And yD <= im.sizeY Then
            If intPoints = 1 Then
                xBB = xD
                yBB = yD
                intPoints = 2
            End If
        End If
    End If
    
    '如果A点在B点的右边，交换AB点的位置
    If xBB < xAA Then
        xA = xAA
        yA = yAA
        xAA = xBB
        yAA = yBB
        xBB = xA
        yBB = yA
    End If
    
    '如果找到两个交点，才调整标注的位置
    If intPoints = 2 And Not (xAA = xBB And yAA = yBB) Then
        la.top = yAA
        la.left = xAA
        
        strTag = xAA & ":" & yAA
        
        If la.top = im.sizeY Then
            la.width = xBB - xAA
            la.height = yBB - yAA
            strTag = strTag & ":" & (xBB - xAA) & ":" & (yBB - yAA)
        ElseIf la.left = 0 Then
            la.width = xBB
            la.height = yBB - yAA
            strTag = strTag & ":" & xBB & ":" & (yBB - yAA)
        Else
            la.width = xBB - xAA
            la.height = yBB
            strTag = strTag & ":" & (xBB - xAA) & ":" & yBB
        End If
        la.Tag = strTag
    End If
    funGetLine = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub subMoveImage(newX As Long, newY As Long, baseX As Long, baseY As Long, ViewerIndex As Integer)
'------------------------------------------------
'功能：在轴位、冠状位和矢状位中拖拽鼠标，上下移动，切换图像
'参数：
'       newX -- 新位置的图像像素X坐标
'       newY -- 新位置的图像像素Y坐标
'       basex -- 旧位置的图像像素X坐标
'       baseY -- 旧位置的图像像素Y坐标
'       ViewerIndex -- 图像所在Viewer的Index；1-轴位图；2-冠状位图；3-矢状位图
'返回：无
'------------------------------------------------
    Dim la As DicomLabel
    Dim newImgX As Long, newImgY As Long, baseImgX As Long, baseImgY As Long
    Dim Size As Long
    Dim img As DicomImage
    
    '提取图像
    If ViewerIndex = 1 Then
        '轴位中鼠标上下移动翻图，相当于冠状位中横线上下移动
        Set img = mCoronalViewer.Images(1)
    ElseIf ViewerIndex = 2 Then
        '冠状位中鼠标上下移动翻图，相当于矢状位竖线左右移动
        Set img = mSagittalViewer.Images(1)
    ElseIf ViewerIndex = 3 Then
        '矢状位中鼠标上下翻图，相当于冠状位竖线左右移动
        Set img = mCoronalViewer.Images(1)
    End If
    
    If ViewerIndex = 1 Then
        Set la = img.Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
        Size = img.sizeY
        
        baseImgX = la.left
        baseImgY = la.top
        newImgX = baseImgX
        newImgY = (newY - baseY) * Size / Viewer(ViewerIndex).Images(1).sizeY + baseImgY
        
    ElseIf ViewerIndex = 2 Or ViewerIndex = 3 Then
        Set la = img.Labels(G_INT_SYS_LABEL_MPR_RESULT_V)
        Size = img.sizeX
            
        baseImgX = la.left
        baseImgY = la.height / 2
        newImgY = baseImgY
        newImgX = Size / Viewer(ViewerIndex).Images(1).sizeY * (newY - baseY) + baseImgX
    End If
    
    Call subMoveCAndSLabel(la, img, newImgX, newImgY, baseImgX, baseImgY, IIf(ViewerIndex = 2, False, True))
    Call img.Refresh(False)
    
    On Error GoTo err
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
