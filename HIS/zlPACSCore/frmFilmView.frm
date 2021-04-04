VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmFilmView 
   Caption         =   "胶片打印--查看图像"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7995
   Icon            =   "frmFilmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7995
   Begin DicomObjects.DicomViewer dcmViewer 
      Height          =   6000
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _Version        =   262147
      _ExtentX        =   12938
      _ExtentY        =   10583
      _StockProps     =   35
      BackColor       =   0
      UseScrollBars   =   0   'False
   End
End
Attribute VB_Name = "frmFilmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'-----------------------------------------------------
'窗体事件
'-----------------------------------------------------
Public Event AfterClose(dcmImage As DicomImage, intViewerIndex As Integer, intImageIndex As Integer)

'窗体内部公共变量
Public SelectedImage As DicomImage      '当前被选中的图像
Public blnDefaultWW2 As Boolean         '记录双窗宽窗位的状态

'窗体内部私有变量
Private mfrmParent As frmFilm
Private mintMouseState As Integer
Private mdcmSelectLabel As DicomLabel   '当前被选中的标注
Private mblnDcmViewDown As Boolean      '用于判断dcmView中鼠标是否被按下
Private mblnLabelMoving As Boolean      '正在移动裁剪框
Private intBaseX As Long                '记录鼠标原来的X位置
Private intBaseY As Long                '记录鼠标原来的Y位置
Private mintSourceViewerIndex As Integer    '记录当前处理的图像所在的Viewer索引
Private mintSourceImageIndex As Integer     '记录当前处理的图像所在的图像索引
Private mdblViewerRatio As Double       '记录Viewer高度/宽度的比例
Private mblnAfterShow As Boolean        '显示完成
Private mintDriverType As Integer       '记录当前设备类型，用来提取对应的调窗快捷键

''''''''''''''''裁剪''''''''''''''''''''''''''''''''''''''''''''''
Private mintCutOutLabel                 '裁剪框所在的标注序号
''''''''''''''''裁剪''''''''''''''''''''''''''''''''''''''''''''''

Private Sub dcmViewer_DblClick()
    '在图像上双击时，进行图像裁剪
    If mintCutOutLabel = 0 Then Exit Sub
    If dcmViewer.Images.Count = 0 Then Exit Sub
    If mintCutOutLabel <> dcmViewer.Images(1).Labels.Count Then Exit Sub
    
    Dim Image As DicomImage
    Dim i As Integer
    Dim lblTemp As DicomLabel
    Dim sourceImage As DicomImage
    
    Set sourceImage = dcmViewer.Images(1)
    Set Image = CutOutAImage(sourceImage)
    
    Image.Name = "ZLPIC"
    '删除框选用的临时标注
    sourceImage.Labels.Remove mintCutOutLabel
    Set mdcmSelectLabel = Nothing
    
    Call subWriteDicomPara(sourceImage, Image)
    
    '把原来图像的标注，添加到现在的图像中
    Image.Labels.Clear
    For i = 1 To sourceImage.Labels.Count
        Image.Labels.Add sourceImage.Labels(i)
    Next i
    
    '把新生成的图像，添加到Viewer中
    dcmViewer.Images.Clear
    dcmViewer.Images.Add Image
    
    '图像放入Viewer中后，重新显示标尺，这个时候标尺和单位才是准确的
    Call UpdateRuler(Image, True)
    
    mintCutOutLabel = 0
    Me.MousePointer = vbArrow
End Sub

Private Sub dcmViewer_KeyDown(KeyCode As Integer, Shift As Integer)
    '处理窗宽窗位快捷按钮，F2,F3-F12
    If KeyCode >= VK_F2 And KeyCode <= VK_F12 Then
        Call subWWWLShortCut(KeyCode)
    End If
End Sub

Private Sub dcmViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    '处理图像操作
    Dim ls As DicomLabels

    If dcmViewer.Images.Count = 0 Then Exit Sub
    intBaseX = x
    intBaseY = y

    If Button = 1 Then
        mintMouseState = mfrmParent.intMouseState

        'mintMouseState 鼠标的状态：0－无；1－调窗；2－漫游；3－缩放;4-选中图像;5-框选缩放;6-裁剪:7-文字标注
        If mintMouseState = 6 Then  '裁剪
            '裁剪状态下的鼠标down，有三种操作：1、画裁剪框（记录标记）；2、移动裁剪框(有焦点) ；3、双击进行裁剪
            If mintCutOutLabel = 0 Then  '画裁剪框
                '增加框选标注
                dcmViewer.Images(1).Labels.Add GetNewLabel(doLabelRectangle, dcmViewer.ImageXPosition(x, y), dcmViewer.ImageYPosition(x, y), 0, 0)
                Set mdcmSelectLabel = dcmViewer.Images(1).Labels(dcmViewer.Images(1).Labels.Count)
                mdcmSelectLabel.Tag = CUT_LABEL
                mblnDcmViewDown = True
                mintCutOutLabel = dcmViewer.Images(1).Labels.Count
            Else    '开始移动裁剪框
                Set ls = dcmViewer.LabelHits(x, y, False, False, True)
                If ls.Count <> 0 And Me.MousePointer <> vbArrow Then
                    '开始移动裁剪框
                    If ls(1).Tag = CUT_LABEL Then
                        mblnLabelMoving = True
                    End If
                End If
            End If
        End If

        If mintMouseState = 5 Then  '框选缩放
            '增加框选标注
            dcmViewer.Images(1).Labels.Add GetNewLabel(doLabelRectangle, dcmViewer.ImageXPosition(x, y), dcmViewer.ImageYPosition(x, y), 0, 0)
            Set mdcmSelectLabel = dcmViewer.Images(1).Labels(dcmViewer.Images(1).Labels.Count)
            mblnDcmViewDown = True
        End If

        If mintMouseState = 7 Then  '文字标注
            Dim dcmLabel As DicomLabel
            Set dcmLabel = GetNewLabel(doLabelText, dcmViewer.ImageXPosition(x, y), dcmViewer.ImageYPosition(x, y), 0, 0)
            dcmViewer.Images(1).Labels.Add dcmLabel
            dcmLabel.AutoSize = True
            dcmLabel.Margin = 0
            dcmLabel.Text = mfrmParent.pstrSideMarker
            dcmLabel.Shadow = doShadowAll
            dcmLabel.ShowTextBox = True
            dcmLabel.Font.Bold = True
            dcmLabel.Tag = POSTURE_LABEL
            mintMouseState = 0
            '设置父窗体的两个参数
            mfrmParent.pstrSideMarker = ""
            mfrmParent.intMouseState = 0
        End If

        dcmViewer.Refresh
    End If
End Sub

Private Sub dcmViewer_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    '处理鼠标移动事件
    Dim dblZoom As Double
    Dim ls As DicomLabels
    Dim lngXOffset As Long
    Dim lngYOffset As Long
    Dim lblCUT As DicomLabel
    
    On Error GoTo err
    
    If dcmViewer.Images.Count = 0 Then Exit Sub
    
    If (Button = 1 And mintMouseState = 1) Or (Button = 4 And intMouseWheelDrag = 2) _
        Or (Button = 2 And cMouseUsage("102").lngMouseKey = 2 And Button_miWidthLevel) Then  '调窗
        If SelectedImage.VOILUT = 1 Then SelectedImage.VOILUT = 0
        SelectedImage.width = SelectedImage.width + (x - intBaseX) * lngWidthLevelStep / 5
        SelectedImage.Level = SelectedImage.Level + (y - intBaseY) * lngWidthLevelStep / 5
        SelectedImage.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & SelectedImage.width & "-L:" & SelectedImage.Level
        intBaseX = x
        intBaseY = y
        dcmViewer.Refresh
    ElseIf (Button = 1 And mintMouseState = 2) Or (Button = 4 And intMouseWheelDrag = 0) _
        Or (Button = 2 And cMouseUsage("103").lngMouseKey = 2 And Button_miCruise) Then '漫游
        subCenterZoom SelectedImage, dcmViewer, SelectedImage.ActualZoom
        SelectedImage.ScrollX = SelectedImage.ScrollX - (x - intBaseX) * lngCruiseStep / 5
        SelectedImage.ScrollY = SelectedImage.ScrollY - (y - intBaseY) * lngCruiseStep / 5
        intBaseX = x
        intBaseY = y
    ElseIf (Button = 1 And mintMouseState = 3) Or (Button = 4 And intMouseWheelDrag = 1) _
        Or (Button = 2 And cMouseUsage("104").lngMouseKey = 2 And Button_miZoom) Then '缩放
        '缩放单位是0.01倍
        dblZoom = SelectedImage.ActualZoom * (1 + (intBaseY - y) * lngZoomStep / 5 * 0.001)
        If dblZoom < 0.01 Then dblZoom = 0.01
        If dblZoom > 64 Then dblZoom = 64
        Call subCenterZoom(SelectedImage, dcmViewer, dblZoom)
        
        If SelectedImage.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
            If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then '更新标尺单位
                Call UpdateRuler(SelectedImage, True)
            End If
        End If
        
        intBaseX = x
        intBaseY = y
    ElseIf Button = 1 And (mintMouseState = 5 Or mintMouseState = 6) Then  '框选缩放和裁剪
        If mblnDcmViewDown = True Then
            mdcmSelectLabel.width = dcmViewer.ImageXPosition(x, y) - mdcmSelectLabel.left
            mdcmSelectLabel.height = dcmViewer.ImageYPosition(x, y) - mdcmSelectLabel.top
            dcmViewer.Refresh
        End If
    End If
    
    '继续处理裁剪
    If mintMouseState = 6 And mintCutOutLabel <> 0 Then
        Set ls = dcmViewer.LabelHits(x, y, False, False, True)
        If Button = 1 Then  '鼠标按下
            If mblnLabelMoving = True Then
                Call subaCorrectCursor(dcmViewer, SelectedImage, x, y)
                Set lblCUT = SelectedImage.Labels(SelectedImage.Labels.Count)
                
                If (Me.MousePointer = vbSizeWE And (SelectedImage.RotateState = doRotateNormal Or SelectedImage.RotateState = doRotate180)) _
                    Or (Me.MousePointer = vbSizeNS And (SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight)) Then    '左右移动
                    
                    lngXOffset = (dcmViewer.ImageXPosition(x, y) - dcmViewer.ImageXPosition(intBaseX, intBaseY))
                    If Abs(lblCUT.left - dcmViewer.ImageXPosition(x, y)) > Abs(lblCUT.left + lblCUT.width - dcmViewer.ImageXPosition(x, y)) Then '右边的移动
                        lblCUT.width = lblCUT.width + lngXOffset
                    Else    '左边的移动
                        lblCUT.left = lblCUT.left + lngXOffset
                        lblCUT.width = lblCUT.width - lngXOffset
                    End If
                ElseIf (Me.MousePointer = vbSizeNS And (SelectedImage.RotateState = doRotateNormal Or SelectedImage.RotateState = doRotate180)) _
                    Or (Me.MousePointer = vbSizeWE And (SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight)) Then    '上下移动
                    
                    lngYOffset = (dcmViewer.ImageYPosition(x, y) - dcmViewer.ImageYPosition(intBaseX, intBaseY))
                    If Abs(lblCUT.top - dcmViewer.ImageYPosition(x, y)) > Abs(lblCUT.top + lblCUT.height - dcmViewer.ImageYPosition(x, y)) Then '下面线的移动
                        lblCUT.height = lblCUT.height + lngYOffset
                    Else    '上面线的移动
                        lblCUT.top = lblCUT.top + lngYOffset
                        lblCUT.height = lblCUT.height - lngYOffset
                    End If
                ElseIf Me.MousePointer = vbSizePointer Then     '整体移动
                    lngXOffset = (dcmViewer.ImageXPosition(x, y) - dcmViewer.ImageXPosition(intBaseX, intBaseY))
                    lngYOffset = (dcmViewer.ImageYPosition(x, y) - dcmViewer.ImageYPosition(intBaseX, intBaseY))
                    lblCUT.top = lblCUT.top + lngYOffset
                    lblCUT.left = lblCUT.left + lngXOffset
                End If
                intBaseX = x
                intBaseY = y
                dcmViewer.Refresh
            End If
        ElseIf Button = 0 Then    '鼠标没有被按下，只改变鼠标指针
            If ls.Count <> 0 Then
                If Abs(ls(1).left - dcmViewer.ImageXPosition(x, y)) < 4 Or Abs(ls(1).left + ls(1).width - dcmViewer.ImageXPosition(x, y)) < 4 Then
                    If SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight Then
                        Me.MousePointer = vbSizeNS
                    Else
                        Me.MousePointer = vbSizeWE
                    End If
                ElseIf Abs(ls(1).top - dcmViewer.ImageYPosition(x, y)) < 4 Or Abs(ls(1).top + ls(1).height - dcmViewer.ImageYPosition(x, y)) < 4 Then
                    If SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight Then
                        Me.MousePointer = vbSizeWE
                    Else
                        Me.MousePointer = vbSizeNS
                    End If
                Else
                    Me.MousePointer = vbSizePointer
                End If
            Else
                Me.MousePointer = vbArrow
            End If
        End If
    End If
    Exit Sub
err:
    
End Sub

Private Sub dcmViewer_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    '处理鼠标弹起事件
    Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
    
    On Error GoTo err
    
    If Button = 1 Then
        If mintMouseState <> 0 Then
            If mintMouseState = 5 And mblnDcmViewDown Then    '框选缩放
                lngLeft = SelectedImage.Labels(SelectedImage.Labels.Count).left * SelectedImage.ActualZoom
                lngTop = SelectedImage.Labels(SelectedImage.Labels.Count).top * SelectedImage.ActualZoom
                lngWidth = SelectedImage.Labels(SelectedImage.Labels.Count).width * SelectedImage.ActualZoom
                lngHeight = SelectedImage.Labels(SelectedImage.Labels.Count).height * SelectedImage.ActualZoom
                
                '调整宽高
                If lngWidth < 0 Then
                    lngLeft = lngLeft + lngWidth
                    lngWidth = -lngWidth
                End If
                
                If lngHeight < 0 Then
                    lngTop = lngTop + lngHeight
                    lngHeight = -lngHeight
                End If
                
                RectangleZoom dcmViewer, SelectedImage, lngLeft, lngTop, lngWidth, lngHeight
                
                '删除框选用的临时标注
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
                Set mdcmSelectLabel = Nothing
                dcmViewer.Refresh
            ElseIf mintMouseState = 6 Then
                If mblnDcmViewDown Then       '裁剪
                    '不做任何操作
                    '如果裁剪框为0 ，则取删除裁剪框，清除裁剪的标记
                    If mdcmSelectLabel.width = 0 Or mdcmSelectLabel.height = 0 Then
                        '删除框选用的临时标注
                        SelectedImage.Labels.Remove SelectedImage.Labels.Count
                        Set mdcmSelectLabel = Nothing
                        dcmViewer.Refresh
                        
                        mintCutOutLabel = 0
                    End If
                End If
            End If
        End If
    End If
    mblnDcmViewDown = False
    mblnLabelMoving = False
    Exit Sub
err:
End Sub

Private Sub Form_Load()
    '读取窗体位置
    Call RestoreWinState(Me, App.ProductName)
    
    '每次打开窗口，都设置默认值
    mintDriverType = 0
    blnDefaultWW2 = False
End Sub

Private Sub Form_Resize()
    Dim lngOldWidth As Long
    Dim lngOldHeight As Long
    
    lngOldWidth = dcmViewer.width
    lngOldHeight = dcmViewer.height
    
    dcmViewer.left = 0
    dcmViewer.top = 0
    dcmViewer.width = Me.ScaleWidth
    dcmViewer.height = Me.ScaleHeight
    
    '调整图像的位置和缩放比例
    If mblnAfterShow And Not SelectedImage Is Nothing Then
        If SelectedImage.StretchToFit = False Then
            Call subScaleImage(SelectedImage, dcmViewer, lngOldWidth, lngOldHeight)
        End If
    End If
End Sub


Public Sub zlShowMe(img As DicomImage, frmParent As frmFilm, intViewerIndex As Integer, intImageIndex As Integer)
'------------------------------------------------
'功能：打开图像处理窗口
'参数： img - 需要处理和显示的图像
'       frmParent - 胶片打印预览窗体
'       intViewerIndex - 当前打开图像所在的Viewer索引
'       intImageIndex -- 当前打开图像所在的图像索引
'返回：无
'------------------------------------------------
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    On Error GoTo err
    
    If img Is Nothing Then Exit Sub
    
    mblnAfterShow = False
    
    Set mfrmParent = frmParent
    mintSourceViewerIndex = intViewerIndex
    mintSourceImageIndex = intImageIndex
    
    lngWidth = frmParent.FilmViewer(mintSourceViewerIndex).width / frmParent.FilmViewer(mintSourceViewerIndex).MultiColumns
    lngHeight = frmParent.FilmViewer(mintSourceViewerIndex).height / frmParent.FilmViewer(mintSourceViewerIndex).MultiRows
    
    mdblViewerRatio = lngHeight / lngWidth
    
    dcmViewer.Images.Clear
    dcmViewer.Images.Add img
    Set SelectedImage = dcmViewer.Images(1)
    
    '调整图像标注的显示
    If SelectedImage.Labels.Count > 0 Then
        Call subChangeLabelForPrint(SelectedImage, 1)
    End If
    
    Me.height = mdblViewerRatio * Abs(Me.width - 115) + 510 '加上标题高度 510,边缘宽度115
    If Me.height < mdblViewerRatio * Abs(Me.width - 115) + 510 Then
        '高度超出了，使用高度计算看宽度
        Me.width = Abs(Me.height - 510) / mdblViewerRatio + 115
    End If
    
    '调整图像的位置和缩放比例
    If img.StretchToFit = False Then
        Call subScaleImage(SelectedImage, dcmViewer, lngWidth, lngHeight)
    End If
    
    Me.Show , mfrmParent
    mblnAfterShow = True
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '    卸载hook
    Call FilmViewUnhook(Me.hwnd, plngFilmViewPreWndProc)
    
    '返回当前处理的图像
    If dcmViewer.Images.Count = 1 Then
        '恢复图像标注的显示
        If dcmViewer.Images(1).Labels.Count > 0 Then
            Call subChangeLabelForPrint(dcmViewer.Images(1), 0)
        End If
        RaiseEvent AfterClose(dcmViewer.Images(1), mintSourceViewerIndex, mintSourceImageIndex)
    End If
    
    '保存窗体位置
    Call SaveWinState(Me, App.ProductName)
End Sub

Public Sub ZLToolButtonClick(control As CommandBarControl)
'------------------------------------------------
'功能：处理胶片预览窗体的工具栏按钮事件
'参数： lngControlID -- 工具栏ID
'返回：直接修改图像
'------------------------------------------------

    On Error GoTo err
    If SelectedImage Is Nothing Then Exit Sub
    
    '''''''''''''''''''''''''''''[功能键设置窗宽窗位处理]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id >= 349 And control.Id <= 360 Then
        subFunctionWL control, Me
        Exit Sub
    End If
    
    Select Case control.Id
        Case ID_frmFilm_FilterLengthUp      '平滑增加
            Call SubImageFiltering("miFilterLengthUp", SelectedImage)
        Case ID_frmFilm_FilterLengthDown     ''平滑减少
            Call SubImageFiltering("miFilterLengthDown", SelectedImage)
        Case ID_frmFilm_Invert               ''反白
            Call subFlipRotate(SelectedImage, "Invert")
        Case ID_frmFilm_RotateLeft           ''向左旋转90度
            Call subFlipRotate(SelectedImage, "RotateAnticlockwise")
        Case ID_frmFilm_RotateRight          ''向右旋转90度
            Call subFlipRotate(SelectedImage, "RotateClockwise")
        Case ID_frmFilm_FlipHorizontal       ''左右镜象
            Call subFlipRotate(SelectedImage, "FlipHorizontal")
        Case ID_frmFilm_FlipVertical         ''上下镜象
            Call subFlipRotate(SelectedImage, "FlipVertical")
        Case ID_frmFilm_Resume               ''恢复
            SelectedImage.SetDefaultWindows
            SelectedImage.FlipState = doFlipNormal
            SelectedImage.RotateState = doRotateNormal
            SelectedImage.StretchToFit = True
            SelectedImage.UnsharpEnhancement = 0
            SelectedImage.UnsharpLength = 0
            SelectedImage.FilterLength = 0
            
            If SelectedImage.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
                If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '更新标尺单位
                    UpdateRuler SelectedImage, True
                End If
            End If
    End Select
    
    dcmViewer.Refresh
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subWWWLShortCut(KeyCode As Integer)
'------------------------------------------------
'功能：处理窗宽窗位快捷键
'参数： KeyCode -- 当时按下的快捷键
'返回：直接修改图像
'------------------------------------------------
    Dim intWidth As Integer
    Dim intLevel As Integer
    Dim strDriverType As String
    Dim i As Integer
    
    On Error GoTo err
    
    If KeyCode < VK_F2 Or KeyCode > VK_F12 Then Exit Sub
    If SelectedImage Is Nothing Then Exit Sub
    
    If KeyCode = VK_F2 Then '默认窗口
        SelectedImage.VOILUT = 1
        '判断是否有两个默认窗口
        If blnDefaultWW2 = False Then
            '显示第二个窗口
            If SelectedImage.Attributes(&H28, &H1050).VM = 2 And SelectedImage.Attributes(&H28, &H1051).VM = 2 Then
                intWidth = SelectedImage.Attributes(&H28, &H1051).ValueByIndex(2)
                intLevel = SelectedImage.Attributes(&H28, &H1050).ValueByIndex(2)
                SelectedImage.width = intWidth
                SelectedImage.Level = intLevel
                blnDefaultWW2 = True
            Else
                SelectedImage.SetDefaultWindows
            End If
        Else
            SelectedImage.SetDefaultWindows
            blnDefaultWW2 = False
        End If
        
        If SelectedImage.Attributes(&H6000, &H15).Value = 1 Then
            If SelectedImage.Level = 0 Then SelectedImage.Level = 1
        End If
    Else    '预设窗口
        '先判断是否需要提取图像类别
        If mintDriverType = 0 Then
            If IsNull(SelectedImage.Attributes(&H8, &H60).Value) Then Exit Sub         '获取Modality
            strDriverType = SelectedImage.Attributes(&H8, &H60).Value
            
            For i = 1 To UBound(aPresetWinWL, 2)        '[找到图像对应设备]
                If UCase(aPresetWinWL(3, i).strModality) = UCase(strDriverType) Then
                    mintDriverType = i
                    Exit For
                End If
            Next i
        End If
        
        If mintDriverType > 0 Then
            If aPresetWinWL(KeyCode - VK_F2 + 2, mintDriverType).bInUse Then
                SelectedImage.width = aPresetWinWL(KeyCode - VK_F2 + 2, mintDriverType).lngWinWidth
                SelectedImage.Level = aPresetWinWL(KeyCode - VK_F2 + 2, mintDriverType).lngWinLevel
            End If
        End If
    End If
    
    SelectedImage.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & SelectedImage.width & "-L:" & SelectedImage.Level
    
    dcmViewer.Refresh
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub MouseWheel(intDirection As Integer)
'------------------------------------------------
'功能：处理鼠标滚轮的消息
'参数：intDirection--滚轮滚动方向 1-鼠标上滚；2-鼠标下滚
'返回：无
'------------------------------------------------
    Dim dblScale As Double
    
    '发生错误，不做任何提示
    On Error Resume Next
    
    If SelectedImage Is Nothing Then Exit Sub
    If dcmViewer.Images.Count <= 0 Then Exit Sub
    
    If intDirection = 1 Then
        dblScale = 1 + (lngZoomStep * 0.01)
    Else
        dblScale = 1 - (lngZoomStep * 0.01)
    End If
    
    Call subCenterZoom(SelectedImage, dcmViewer, SelectedImage.ActualZoom * dblScale)
        
    If SelectedImage.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
        If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then '更新标尺单位
            Call UpdateRuler(SelectedImage, True)
        End If
    End If
    
End Sub
