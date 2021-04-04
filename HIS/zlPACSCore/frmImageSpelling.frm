VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmImageSpelling 
   BackColor       =   &H8000000B&
   Caption         =   "图像拼接"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   10005
   Icon            =   "frmImageSpelling.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicViewer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5292
      Left            =   600
      ScaleHeight     =   5265
      ScaleWidth      =   8955
      TabIndex        =   1
      Top             =   960
      Width           =   8988
      Begin DicomObjects.DicomViewer viewer 
         Height          =   2505
         Index           =   0
         Left            =   2640
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   2790
         _Version        =   262147
         _ExtentX        =   4932
         _ExtentY        =   4424
         _StockProps     =   35
      End
   End
   Begin VB.ListBox lstSort 
      Height          =   240
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7350
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Text            =   "中联软件"
            TextSave        =   "中联软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10186
            Text            =   "双击观片器上图像直接把图像加入拼接器"
            TextSave        =   "双击观片器上图像直接把图像加入拼接器"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "缩放倍数："
            TextSave        =   "缩放倍数："
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "大写"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   926
            MinWidth        =   706
            Text            =   "数字"
            TextSave        =   "NUM"
            Object.ToolTipText     =   "数字"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImgIcons 
      Left            =   3000
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmImageSpelling.frx":0CCA
   End
   Begin XtremeCommandBars.CommandBars CommBar_ImageSelling 
      Left            =   120
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmImageSpelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intImageCount As Integer
Private intBaseX As Long                     '''记录鼠标原来的X位置
Private intBaseY As Long                     '''记录鼠标原来的Y位置
Dim intSelectedViewer As Integer
Dim iMaxTag As Long
Dim iMaxViewer As Integer
Private mintMouseState As Integer        '''记录鼠标的状态：0－无（漫游）；2－缩放;3-裁剪
Private mdcmSelectLabel As DicomLabel   '当前被选中的标注
Private mblnMouseDown As Boolean        '鼠标是否被按下
Private SelectedImg As DicomImage               '''记录当前被选中的图像
Private mRViewerWidth As Long                   '''记录拼接完成后图像的宽度
Private mRViewerHeight As Long                  '''记录拼接完成后图像的高度
Private mstrSeriesUID As String                 '''记录拼接完成的图像所使用的原图的序列UID，用于图像保存

Public f As frmViewer

''''''''''''''''裁剪''''''''''''''''''''''''''''''''''''''''''''''
Private mintCutOutViewer                '裁剪框所在的viewer序号
Private mintCutOutImage                 '裁剪框所在的图像序号
Private mintCutOutLabel                 '裁剪框所在的标注序号
Private mblnLabelMoving As Boolean      '正在移动裁剪框
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function funCompleteSplling() As DicomImage
'------------------------------------------------
'功能： 对当前摆好的图像进行拼接，生成结果图像
'参数： 无
'返回： 拼接的结果图像
'说明： 当图像分别是 Monochrome1和Monochrome2时，Window=False,以哪个为主，另一个就会被反色。
'       当图像分别是 Monochrome1和Monochrome2时，Window=true，两个图都会被反色，而且如果图像是Signed的，还不能作为目的图。
'       当图像中包含非Monochrome1和Monochrome2，比如包含RGB时，Window=False,会提示错误，需要设置Window=True。
'       当图像中包含非Monochrome1和Monochrome2，比如包含RGB时，Window=True，以黑白图为主，得到黑白图，以彩色图为主得到彩色图。
'------------------------------------------------
    If intImageCount < 1 Then
        Exit Function
    End If
    Dim i As Integer
    Dim dblLeft As Double, dblTop As Double
    Dim dblWidth As Double, dblHeight As Double
    Dim dblRight As Double, dblBottom As Double
    Dim NewImg As New DicomImage
    Dim sizex As Long   '第一个图像的x像素数量
    Dim sizey As Long   '第一个图像的y像素数量
    Dim iViewerIndex As Integer     '表示当前使用的是哪个viewer
    Dim ClipSizex As Integer    '被剪切的图像的x方向像素数量
    Dim ClipSizey As Integer    '被剪切的图像的y方向像素数量
    Dim view As DicomViewer
    Dim dblMaxZoom As Double
    Dim dcmGlobal As New DicomGlobal
    Dim blnWindow As Boolean    '拼接的时候，Window参数的值
    Dim intMainImage As Integer     '拼接的时候，使用哪个图像作为主图象，优先使用彩色图像为主图像
    Dim MainImage As New DicomImage     '拼接时用的主图像

    On Error GoTo err
    
    blnWindow = False       '默认为False
    intMainImage = 0        '默认为0
    
    '初始化新图像的位置和大小
    dblLeft = Viewer(intSelectedViewer).left
    dblTop = Viewer(intSelectedViewer).top
    dblRight = Viewer(intSelectedViewer).left + Viewer(intSelectedViewer).width
    dblBottom = Viewer(intSelectedViewer).top + Viewer(intSelectedViewer).height
    dblMaxZoom = Viewer(intSelectedViewer).Images(1).ActualZoom

    '清空现有的lstSort
    Me.lstSort.Clear
    
    '对原图像进行排序，按照tag的大小来排序，从小到大.
    '获取新图像的位置:左，顶，右，底,取当前最大的图像
    '循环所有的图像
    For Each view In Viewer
        If view.Index <> 0 Then
            If view.left < dblLeft Then
                dblLeft = view.left
            End If
            If view.top < dblTop Then
                dblTop = view.top
            End If
            If view.left + view.width > dblRight Then
                dblRight = view.left + view.width
            End If
            If view.top + view.height > dblBottom Then
                dblBottom = view.top + view.height
            End If
            If dblMaxZoom < view.Images(1).ActualZoom Then
                dblMaxZoom = view.Images(1).ActualZoom
            End If
            
            '将图像的TAG和viewer 的index放到lstSort中，后面进行排序
            Me.lstSort.AddItem Format(view.Images(1).Tag, "0000")
            Me.lstSort.ItemData(Me.lstSort.NewIndex) = view.Index
            
            '记录图像的(0028,0004) Photometric Interpretation
            If intMainImage = 0 And view.Images(1).Attributes(&H28, &H4).Exists And Not IsNull(view.Images(1).Attributes(&H28, &H4).Value) Then
                If UCase(view.Images(1).Attributes(&H28, &H4).Value) = "MONOCHROME1" Or UCase(view.Images(1).Attributes(&H28, &H4).Value) = "MONOCHROME2" Then
                    '不做处理
                Else
                    blnWindow = True
                    intMainImage = view.Index
                End If
            End If
        End If
    Next
    
    If intMainImage = 0 Then intMainImage = intSelectedViewer
    
    '保存新Viewer的宽度和高度
    dblWidth = dblRight - dblLeft
    dblHeight = dblBottom - dblTop
    mRViewerWidth = dblWidth
    mRViewerHeight = dblHeight
    
    '将新图像的位置从缇转换为像素
    dblLeft = dblLeft / Screen.TwipsPerPixelX
    dblWidth = dblWidth / Screen.TwipsPerPixelX
    dblTop = dblTop / Screen.TwipsPerPixelY
    dblHeight = dblHeight / Screen.TwipsPerPixelY
    
    '如果存在彩色图像，先将主图像保存成BMP格式的,避免JPG图像拼接后变成绿图。
    If blnWindow = True Then
        Viewer(intMainImage).Images(1).FileExport "tmpBMPFile", "BMP"
        MainImage.FileImport "tmpBMPFile", "BMP"
        MainImage.StudyUID = Viewer(intMainImage).Images(1).StudyUID
        MainImage.SeriesUID = Viewer(intMainImage).Images(1).SeriesUID
        MainImage.PatientID = Viewer(intMainImage).Images(1).PatientID
        MainImage.Name = Viewer(intMainImage).Images(1).Name
    Else
        Set MainImage = Viewer(intMainImage).Images(1)
    End If

    '创建新图像,使用sizex和sizey是为了在创新新图像时，不出现原来的图像。
    sizex = MainImage.sizex
    sizey = MainImage.sizey
    
    Set NewImg = MainImage.SubImage(sizex, sizey, dblWidth / dblMaxZoom, dblHeight / dblMaxZoom, 1, MainImage.Frame)

    '删除图像中原有的遮盖Shutter信息
    NewImg.Attributes.Remove &H18, &H1600
    NewImg.Attributes.Remove &H18, &H1602
    NewImg.Attributes.Remove &H18, &H1604
    NewImg.Attributes.Remove &H18, &H1606
    NewImg.Attributes.Remove &H18, &H1608
    NewImg.Attributes.Remove &H18, &H1610
    NewImg.Attributes.Remove &H18, &H1612
    NewImg.Attributes.Remove &H18, &H1620
    NewImg.Attributes.Remove &H18, &H1622
    
    '将原有图像一个个复制到新图像中。
    For i = 0 To Me.lstSort.ListCount - 1
        iViewerIndex = Me.lstSort.ItemData(i)
        ClipSizex = (Viewer(iViewerIndex).width / Screen.TwipsPerPixelX) / dblMaxZoom
        ClipSizey = (Viewer(iViewerIndex).height / Screen.TwipsPerPixelY) / dblMaxZoom
                        
        NewImg.Blt Viewer(iViewerIndex).Images(1), Viewer(iViewerIndex).Images(1).ActualScrollX, _
                Viewer(iViewerIndex).Images(1).ActualScrollY, (Viewer(iViewerIndex).left / Screen.TwipsPerPixelX - dblLeft) / dblMaxZoom, _
                (Viewer(iViewerIndex).top / Screen.TwipsPerPixelY - dblTop) / dblMaxZoom, ClipSizex, ClipSizey, Viewer(iViewerIndex).Images(1).Frame, 1, Viewer(iViewerIndex).Images(1).ActualZoom / dblMaxZoom, blnWindow
    Next
    
    NewImg.width = MainImage.width
    NewImg.Level = MainImage.Level
    
    '保存和修改序列UID
    mstrSeriesUID = NewImg.SeriesUID
    NewImg.SeriesUID = dcmGlobal.NewUID

    Set funCompleteSplling = NewImg
    
    '删除参与拼接的图像
    If intImageCount >= 1 Then
        For Each view In Viewer
            If view.Index <> 0 Then Unload view
        Next
    End If
    intImageCount = 0
    intSelectedViewer = 0
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CommBar_ImageSelling_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.Id
        '完成拼接
        Case ID_frmImageSpelling_CompleteSpelling
            subLoadImage funCompleteSplling
            Viewer(iMaxViewer).width = mRViewerWidth
            Viewer(iMaxViewer).height = mRViewerHeight
            Call subDrawFrame(Viewer(iMaxViewer), False, True)
        '保存图像，退出
        Case ID_frmImageSpelling_SavePhoto
            If intSelectedViewer = 0 Then Exit Sub
            
            '把图像保存到服务器
            If subSaveImage(Viewer(intSelectedViewer).Images(1), mstrSeriesUID) = True Then
                '打开并显示这个图像
                Call subOpenCurrentImage(f, Viewer(intSelectedViewer).Images(1))
            End If
            '退出
            Unload Me
        '删除图像
        Case ID_frmImageSpelling_DelPhoto
            If intImageCount < 1 Then Exit Sub
            Unload Viewer(intSelectedViewer)
            intImageCount = intImageCount - 1
            intSelectedViewer = Viewer.UBound
        '移动
        Case ID_frmImageSpelling_Move
            subSetToolBarChecked control.Id
            mintMouseState = 0
        '缩放
        Case ID_frmImageSpelling_ZoomOut
            subSetToolBarChecked control.Id
            mintMouseState = 2
        '裁剪
        Case ID_frmImageSpelling_CutOut
            subSetToolBarChecked control.Id
            
            '设置鼠标状态
            mintMouseState = 3
            
        '退出
        Case ID_frmImageSpelling_Quit
            Unload Me
    End Select
End Sub

Private Sub CommBar_ImageSelling_GetClientBordersWidth(left As Long, top As Long, Right As Long, Bottom As Long)
    If sbStatusBar.Visible Then
        Bottom = sbStatusBar.height
    End If
End Sub

Private Sub CommBar_ImageSelling_Resize()
    On Error Resume Next
    
    Dim left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    Me.CommBar_ImageSelling.GetClientRect left, top, Right, Bottom
    If Right >= left And Bottom >= top Then
        picViewer.Move left, top, Right - left, Bottom - top
    Else
        picViewer.Move 0, 0, 0, 0
    End If
    
End Sub

Private Sub CommBar_ImageSelling_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.Id
        Case ID_frmImageSpelling_CompleteSpelling, ID_frmImageSpelling_SavePhoto, ID_frmImageSpelling_DelPhoto, _
             ID_frmImageSpelling_Move, ID_frmImageSpelling_ZoomOut, ID_frmImageSpelling_CutOut
            If Viewer.Count <= 1 Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
    End Select
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    '创建工具条
    Call CreateBar
    '设置状态栏图标
    'Set sbStatusBar.Panels(1).Picture = f.ImgList32.ListImages(4).Picture
    intImageCount = 0
    iMaxViewer = 0
    iMaxTag = 0
End Sub

Public Sub subLoadImage(im As DicomImage)
'------------------------------------------------
'功能： 装载图像，把图像加载进入拼接窗口
'参数： im---需要加载的图像
'返回： 无
'------------------------------------------------
    
    If im Is Nothing Then Exit Sub
    
    On Error GoTo err
    
    If im.Attributes(&H28, &H4) = "PALETTE COLOR" Then
        MsgBox "彩色图像不能进行图像拼接。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    intImageCount = intImageCount + 1
    iMaxViewer = iMaxViewer + 1
    load Viewer(iMaxViewer)
    Viewer(iMaxViewer).Visible = True
    Viewer(iMaxViewer).MultiColumns = 1
    Viewer(iMaxViewer).MultiRows = 1
    Viewer(iMaxViewer).UseScrollBars = False
    Viewer(iMaxViewer).Images.Add im
    Viewer(iMaxViewer).BackColour = vbBlack
    Viewer(iMaxViewer).width = im.sizex * im.ActualZoom * Screen.TwipsPerPixelX
    Viewer(iMaxViewer).height = im.sizey * im.ActualZoom * Screen.TwipsPerPixelY
    subDrawFrame Viewer(iMaxViewer), True, True
    
    '删除标注
    Viewer(iMaxViewer).Images(1).Labels.Clear
    
    Viewer(iMaxViewer).Images(1).StretchToFit = True
    
    iMaxTag = iMaxTag + 1
    Viewer(iMaxViewer).Images(1).Tag = iMaxTag
    If intImageCount = 1 Then
        intSelectedViewer = iMaxViewer
        Viewer(iMaxViewer).Labels(1).ForeColour = vbRed
    End If
    Viewer(iMaxViewer).Refresh
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub subDrawFrame(v As DicomViewer, isNew As Boolean, isSelected As Boolean)
'------------------------------------------------
'功能： 画图像的选择边框
'参数： v ---图像所在的Viwer
'       isNew ---是否新增加的图像,新增加的图像需要新增加边框label
'       isSelected --- 图像是否被选择，被选择时边框颜色不同
'返回： 无
'------------------------------------------------
    Dim l As New DicomLabel
    
    On Error GoTo err
    
    If Not isNew Then Set l = v.Labels(1)
    
    l.LabelType = doLabelRectangle
    l.left = 0
    l.top = 0
    l.width = v.width / Screen.TwipsPerPixelX
    l.height = v.height / Screen.TwipsPerPixelY
    l.ForeColour = IIf(isSelected, vbRed, vbWhite)
    If isNew Then
        l.ImageTied = True
        v.Labels.Add l
    End If
    v.Refresh
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    f.blnfis = False
End Sub

Private Sub picViewer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = 0
End Sub

Private Sub viewer_DblClick(Index As Integer)
    Call sub裁剪
End Sub

Private Sub Viewer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    Dim v As DicomViewer
    Dim intImgIndex As Integer
    Dim ls As DicomLabels
    
    intBaseX = x
    intBaseY = y
    
    intImgIndex = Viewer(Index).ImageIndex(x, y)
    If Viewer(Index).Images.Count > 0 And intImgIndex <> 0 Then
        Set SelectedImg = Viewer(Index).Images(intImgIndex)
        If Button = 1 Then
            If mintMouseState = 0 Then      '移动图像
                mblnMouseDown = True
            ElseIf mintMouseState = 1 Then       '调窗
                mblnMouseDown = True
            ElseIf mintMouseState = 2 Then       '缩放
                mblnMouseDown = True
            ElseIf mintMouseState = 3 Then       '裁剪
                '裁剪状态下的鼠标down，有三种操作：1、画裁剪框（记录标记）；2、移动裁剪框(有焦点) ；3、双击进行裁剪
                If mintCutOutViewer = 0 Or mintCutOutImage = 0 Or mintCutOutLabel = 0 Then  '画裁剪框
                    '增加框选标注
                    Viewer(Index).Images(intImgIndex).Labels.Add GetNewLabel(doLabelRectangle, Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y), 0, 0)
                    Set mdcmSelectLabel = Viewer(Index).Images(intImgIndex).Labels(Viewer(Index).Images(intImgIndex).Labels.Count)
                    mdcmSelectLabel.Tag = CUT_LABEL
                    mblnMouseDown = True
                    mintCutOutViewer = Index
                    mintCutOutImage = intImgIndex
                    mintCutOutLabel = Viewer(Index).Images(intImgIndex).Labels.Count
                    Viewer(Index).Refresh
                Else            '开始移动裁剪框
                    Set ls = Viewer(Index).LabelHits(x, y, False, False, True)
                    If ls.Count <> 0 And Screen.MousePointer <> vbArrow Then
                        '开始移动裁剪框
                        If ls(1).Tag = CUT_LABEL And SelectedImg.Labels(SelectedImg.Labels.Count).Tag = CUT_LABEL Then
                            mblnLabelMoving = True
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Viewer(Index).ZOrder
    iMaxTag = iMaxTag + 1
    Viewer(Index).Images(1).Tag = iMaxTag
    intSelectedViewer = Index
    For Each v In Viewer
        If v.Index <> 0 Then
            subDrawFrame v, False, IIf((v.Index = Index), True, False)
        End If
    Next
End Sub

Private Sub Viewer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim ls As DicomLabels
    Dim lngXOffset As Long
    Dim lngYOffset As Long
    Dim lblCUT As DicomLabel
    Dim i As Integer
    Dim v As DicomViewer
    Dim dblZoom As Double
    Dim dblZoomRatio As Double
    
    If SelectedImg Is Nothing Then Exit Sub
    
    If Button = 1 Then
        Select Case mintMouseState
            Case 0                      '移动图像
                If mblnMouseDown = True Then
                    If Viewer(Index).left + (x - intBaseX) * Screen.TwipsPerPixelX > 24 And x <> intBaseX Then
                        Viewer(Index).left = Viewer(Index).left + (x - intBaseX) * Screen.TwipsPerPixelX
                    End If
                    If Viewer(Index).top + (y - intBaseY) * Screen.TwipsPerPixelX > 24 And y <> intBaseY Then
                        Viewer(Index).top = Viewer(Index).top + (y - intBaseY) * Screen.TwipsPerPixelX
                    End If
                End If
            Case 2                  '缩放
                If mblnMouseDown = True Then
                    '缩放单位是0.01倍
                    dblZoom = SelectedImg.ActualZoom * (1 + (intBaseY - y) * 0.001)
                    If dblZoom < 0.01 Then dblZoom = 0.01
                    If dblZoom > 64 Then dblZoom = 64
                    dblZoomRatio = dblZoom / SelectedImg.ActualZoom
                    Viewer(Index).width = Viewer(Index).width * dblZoomRatio
                    Viewer(Index).height = Viewer(Index).height * dblZoomRatio

                    Call subDrawFrame(Viewer(Index), False, True)

                    intBaseX = x
                    intBaseY = y
                End If
            Case 3                  '裁剪
                If mblnMouseDown = True Then
                    mdcmSelectLabel.width = Viewer(Index).ImageXPosition(x, y) - mdcmSelectLabel.left
                    mdcmSelectLabel.height = Viewer(Index).ImageYPosition(x, y) - mdcmSelectLabel.top
                    Viewer(Index).Refresh
                End If
        End Select
    End If
    
    '单独处理裁剪
    If mintMouseState = 3 And mintCutOutViewer <> 0 And mintCutOutImage <> 0 And mintCutOutLabel <> 0 Then
        Set ls = Viewer(Index).LabelHits(x, y, False, False, True)
        If Button = 1 Then          '鼠标被按下
            If mblnLabelMoving = True Then
                Call subaCorrectCursor(Viewer(Index), SelectedImg, x, y)
                Set lblCUT = SelectedImg.Labels(SelectedImg.Labels.Count)

                If (Screen.MousePointer = vbSizeWE And (SelectedImg.RotateState = doRotateNormal Or SelectedImg.RotateState = doRotate180)) _
                    Or (Screen.MousePointer = vbSizeNS And (SelectedImg.RotateState = doRotateLeft Or SelectedImg.RotateState = doRotateRight)) Then       '左右移动

                    lngXOffset = (Viewer(Index).ImageXPosition(x, y) - Viewer(Index).ImageXPosition(intBaseX, intBaseY))
                    If Abs(lblCUT.left - Viewer(Index).ImageXPosition(x, y)) > Abs(lblCUT.left + lblCUT.width - Viewer(Index).ImageXPosition(x, y)) Then '右边的移动
                            lblCUT.width = lblCUT.width + lngXOffset
                    Else    '左边线移动
                            lblCUT.left = lblCUT.left + lngXOffset
                            lblCUT.width = lblCUT.width - lngXOffset
                    End If
                ElseIf (Screen.MousePointer = vbSizeNS And (SelectedImg.RotateState = doRotateNormal Or SelectedImg.RotateState = doRotate180)) _
                    Or (Screen.MousePointer = vbSizeWE And (SelectedImg.RotateState = doRotateLeft Or SelectedImg.RotateState = doRotateRight)) Then   '上下移动

                    lngYOffset = (Viewer(Index).ImageYPosition(x, y) - Viewer(Index).ImageYPosition(intBaseX, intBaseY))
                    If Abs(lblCUT.top - Viewer(Index).ImageYPosition(x, y)) > Abs(lblCUT.top + lblCUT.height - Viewer(Index).ImageYPosition(x, y)) Then    '下面线的移动
                        lblCUT.height = lblCUT.height + lngYOffset

                    Else    '上面线移动
                        lblCUT.top = lblCUT.top + lngYOffset
                        lblCUT.height = lblCUT.height - lngYOffset
                    End If
                ElseIf Screen.MousePointer = vbSizePointer Then     '整体移动

                    lngXOffset = (Viewer(Index).ImageXPosition(x, y) - Viewer(Index).ImageXPosition(intBaseX, intBaseY))
                    lngYOffset = (Viewer(Index).ImageYPosition(x, y) - Viewer(Index).ImageYPosition(intBaseX, intBaseY))
                    lblCUT.top = lblCUT.top + lngYOffset
                    lblCUT.left = lblCUT.left + lngXOffset
                End If
                intBaseX = x
                intBaseY = y
                Viewer(Index).Refresh
            End If
        ElseIf Button = 0 Then
            If ls.Count <> 0 Then
                If Abs(ls(1).left - Viewer(Index).ImageXPosition(x, y)) < 4 Or Abs(ls(1).left + ls(1).width - Viewer(Index).ImageXPosition(x, y)) < 4 Then
                    If SelectedImg.RotateState = doRotateLeft Or SelectedImg.RotateState = doRotateRight Then
                        Screen.MousePointer = vbSizeNS
                    Else
                        Screen.MousePointer = vbSizeWE
                    End If
                ElseIf Abs(ls(1).top - Viewer(Index).ImageYPosition(x, y)) < 4 Or Abs(ls(1).top + ls(1).height - Viewer(Index).ImageYPosition(x, y)) < 4 Then
                    If SelectedImg.RotateState = doRotateLeft Or SelectedImg.RotateState = doRotateRight Then
                        Screen.MousePointer = vbSizeWE
                    Else
                        Screen.MousePointer = vbSizeNS
                    End If
                Else
                    Screen.MousePointer = vbSizePointer
                End If
            Else
                Screen.MousePointer = vbArrow
            End If
        End If
    End If
End Sub
Private Sub Viewer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 1 Then
        If mintMouseState = 2 Then          '缩放
            '取消缩放操作
            Call subSetToolBarChecked(ID_frmImageSpelling_Move)
            mintMouseState = 0
        ElseIf mintMouseState = 3 Then
            If mblnMouseDown Then           '裁剪
                '不做任何操作
                '如果裁剪框为0 ，则取删除裁剪框，清除裁剪的标记
                If mdcmSelectLabel.width = 0 Or mdcmSelectLabel.height = 0 Then
                    '删除框选用的临时标注
                    SelectedImg.Labels.Remove SelectedImg.Labels.Count
                    Set mdcmSelectLabel = Nothing
                    Viewer(Index).Refresh
                    
                    mintCutOutViewer = 0
                    mintCutOutImage = 0
                    mintCutOutLabel = 0
                End If
            End If
        End If
    End If
    mblnMouseDown = False
    mblnLabelMoving = False
End Sub

Private Sub CreateBar()
    '------------------------------------------------
    '功能：                                  创建菜单
    '参数：
    '返回：                                  无
    '------------------------------------------------
    Dim ToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.CommBar_ImageSelling.VisualTheme = xtpThemeOffice2003
    Me.CommBar_ImageSelling.Icons = ImgIcons.Icons
    Me.CommBar_ImageSelling.Item(1).Visible = False                                 '隐藏菜单栏
    
    With Me.CommBar_ImageSelling.Options
        .ShowExpandButtonAlways = False     '去掉扩展按钮
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 32, 32
    End With
    
    '建立主工具栏
    Set ToolBar = Me.CommBar_ImageSelling.Add("主工具栏", xtpBarBottom)
    ToolBar.Position = xtpBarTop
    ToolBar.ShowTextBelowIcons = True
    With ToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_CompleteSpelling, "拼接")
            cbrControl.IconId = 1010: cbrControl.ToolTipText = "进行图像拼接"
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_SavePhoto, "保存退出")
            cbrControl.IconId = 1009: cbrControl.ToolTipText = "保存拼接完成的图像，并退出系统"
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_Move, "移动图像"): cbrControl.BeginGroup = True
            cbrControl.IconId = 1007: cbrControl.ToolTipText = "移动图像"
            cbrControl.Checked = True                                       '默认为移动图像
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_ZoomOut, "缩放")
            cbrControl.IconId = 1005: cbrControl.ToolTipText = "图像缩放"
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_CutOut, "裁剪")
            cbrControl.IconId = 1006: cbrControl.ToolTipText = "图像裁剪"
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_DelPhoto, "删除图像")
            cbrControl.IconId = 1002: cbrControl.ToolTipText = "删除图像": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_Quit, "退出")
            cbrControl.IconId = 1003: cbrControl.ToolTipText = "直接退出系统"
    End With
End Sub

Private Sub sub裁剪()
'------------------------------------------------
'功能： 裁剪图像，裁剪当前选中的图像
'参数： 无
'返回： 无，直接显示裁剪的结果图
'------------------------------------------------
    If mintCutOutViewer = 0 Or mintCutOutImage = 0 Or mintCutOutLabel = 0 Then Exit Sub
    If mintCutOutImage > Viewer(mintCutOutViewer).Images.Count Then Exit Sub
    If mintCutOutLabel <> Viewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Count Then Exit Sub
    
    Dim Image As DicomImage
    Dim i As Integer
    Dim lblCUT As DicomLabel
    Dim sourceImage As DicomImage
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    Dim lngNewLeft As Long
    Dim lngNewTop As Long
    
    On Error GoTo err
    
    Set sourceImage = Viewer(mintCutOutViewer).Images(mintCutOutImage)
    Set lblCUT = sourceImage.Labels(sourceImage.Labels.Count)
    
    If lblCUT.width < 0 Then
        lngNewLeft = (lblCUT.left + lblCUT.width) * sourceImage.ActualZoom * Screen.TwipsPerPixelX
    Else
        lngNewLeft = lblCUT.left * sourceImage.ActualZoom * Screen.TwipsPerPixelX
    End If
    If lblCUT.height < 0 Then
        lngNewTop = (lblCUT.top + lblCUT.height) * sourceImage.ActualZoom * Screen.TwipsPerPixelY
    Else
        lngNewTop = lblCUT.top * sourceImage.ActualZoom * Screen.TwipsPerPixelY
    End If
    lngNewWidth = Abs(sourceImage.Labels(sourceImage.Labels.Count).width) * sourceImage.ActualZoom * Screen.TwipsPerPixelX
    lngNewHeight = Abs(sourceImage.Labels(sourceImage.Labels.Count).height) * sourceImage.ActualZoom * Screen.TwipsPerPixelY
    
    Set Image = CutOutAImage(sourceImage)
    
    '删除框选用的临时标注
    sourceImage.Labels.Remove mintCutOutLabel
    Set mdcmSelectLabel = Nothing
    
    '把新生成的图像，添加到Viewer中
    If mintCutOutImage = 1 And Viewer(mintCutOutViewer).Images.Count = 1 Then
        Viewer(mintCutOutViewer).Images.Clear
        Viewer(mintCutOutViewer).Images.Add Image
    Else
        Viewer(mintCutOutViewer).Images.Remove mintCutOutImage
        Viewer(mintCutOutViewer).Images.Add Image
        Viewer(mintCutOutViewer).Images.Move Viewer(mintCutOutViewer).Images.Count, mintCutOutImage
    End If
    
    '设置viewer的长度和宽度,并把Viewer移动到原来画裁剪框的位置
    Viewer(mintCutOutViewer).left = Viewer(mintCutOutViewer).left + lngNewLeft
    Viewer(mintCutOutViewer).top = Viewer(mintCutOutViewer).top + lngNewTop
    Viewer(mintCutOutViewer).width = lngNewWidth
    Viewer(mintCutOutViewer).height = lngNewHeight
    

    mintCutOutViewer = 0
    mintCutOutImage = 0
    mintCutOutLabel = 0
    Screen.MousePointer = vbArrow
    
    '完成裁剪后，取消鼠标裁剪的功能，恢复菜单项
    
    subSetToolBarChecked ID_frmImageSpelling_Move
    mintMouseState = 0
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subSetToolBarChecked(lngControlID As Long)
    On Error GoTo err
    
    '先把所有按钮设置成False
    CommBar_ImageSelling.Item(2).FindControl(, ID_frmImageSpelling_Move, , True).Checked = False
    CommBar_ImageSelling.Item(2).FindControl(, ID_frmImageSpelling_ZoomOut, , True).Checked = False
    CommBar_ImageSelling.Item(2).FindControl(, ID_frmImageSpelling_CutOut, , True).Checked = False
    '设置对应的按钮为True
    CommBar_ImageSelling.Item(2).FindControl(, lngControlID, , True).Checked = True
    
    '处理裁剪的框
    '如果原来已经有裁剪框，则先删除这个裁剪框
    If mintCutOutViewer <> 0 And mintCutOutImage <> 0 And mintCutOutLabel <> 0 Then
        If mintCutOutViewer < Viewer.Count Then
            If mintCutOutImage <= Viewer(mintCutOutViewer).Images.Count Then
                If mintCutOutLabel = Viewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Count Then
                    '删除框选用的临时标注
                    Viewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Remove mintCutOutLabel
                    Set mdcmSelectLabel = Nothing
                    Viewer(mintCutOutViewer).Refresh
                End If
            End If
        End If
    End If
    
    '初始化裁剪参数
    mintCutOutViewer = 0
    mintCutOutImage = 0
    mintCutOutLabel = 0
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

