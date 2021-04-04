VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{5C493D4E-FD57-4FF4-9BA4-C6C670BFF9A7}#58.0#0"; "zl9PacsControl.ocx"
Begin VB.Form frmReportImage 
   BorderStyle     =   0  'None
   Caption         =   "报告图像"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMiniViewer 
      Height          =   1845
      Left            =   4440
      ScaleHeight     =   1785
      ScaleWidth      =   3615
      TabIndex        =   16
      Top             =   4185
      Width           =   3675
      Begin zl9PacsControl.ucImagePreview ucMiniImageViewer 
         Height          =   1695
         Left            =   45
         TabIndex        =   17
         Top             =   45
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   2990
         BackColor       =   4210752
      End
   End
   Begin VB.PictureBox picMenu 
      Height          =   540
      Left            =   2100
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   585
      Begin XtremeCommandBars.CommandBars cbrMain 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin MSComctlLib.ImageList listCur 
      Bindings        =   "frmReportImage.frx":0000
      Left            =   1200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportImage.frx":0014
            Key             =   "Pen"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMiniImageC 
      Height          =   1935
      Left            =   285
      ScaleHeight     =   1875
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   4170
      Width           =   3735
      Begin VB.VScrollBar vscrollMini 
         Height          =   1815
         Left            =   3360
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin DicomObjects.DicomViewer dcmMiniImageC 
         Height          =   1695
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3135
         _Version        =   262147
         _ExtentX        =   5530
         _ExtentY        =   2990
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin VB.PictureBox picReportImage 
      Height          =   2055
      Left            =   3480
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
      Begin DicomObjects.DicomViewer dcmReportImage 
         Height          =   1695
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _Version        =   262147
         _ExtentX        =   3413
         _ExtentY        =   2990
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin VB.PictureBox picMark 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
      Begin VB.PictureBox picNumMark 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1030
         Left            =   300
         ScaleHeight     =   1035
         ScaleWidth      =   2040
         TabIndex        =   6
         Top             =   1300
         Width           =   2040
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":032E
            Height          =   510
            Index           =   1
            Left            =   490
            Picture         =   "frmReportImage.frx":0F70
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":1BB2
            Height          =   510
            Index           =   4
            Left            =   510
            Picture         =   "frmReportImage.frx":27F4
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":3436
            Height          =   510
            Index           =   2
            Left            =   1000
            Picture         =   "frmReportImage.frx":4078
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":4CBA
            Height          =   510
            Index           =   5
            Left            =   1010
            Picture         =   "frmReportImage.frx":58FC
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":653E
            Height          =   510
            Index           =   3
            Left            =   1560
            Picture         =   "frmReportImage.frx":7180
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":7DC2
            Height          =   510
            Index           =   6
            Left            =   1510
            Picture         =   "frmReportImage.frx":8A04
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":9646
            Height          =   1020
            Index           =   0
            Left            =   0
            Picture         =   "frmReportImage.frx":A288
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "自动编号"
            Top             =   0
            Value           =   1  'Checked
            Width           =   510
         End
      End
      Begin DicomObjects.DicomViewer dcmMark 
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _Version        =   262147
         _ExtentX        =   2990
         _ExtentY        =   1720
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   120
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSingleWindow As Boolean     '是否使用独立窗口显示报告编辑器，True-独立窗口显示；False-嵌入式显示
Private mlngAdviceID As Long    '医嘱ID
Private mintEditType As Integer '病历状态 0 创建，1书写，2 修订
Private mlngReportID As Long    '报告内容ID
Private mlngFileID As Long      '报告单格式ID
Private mlngShowBigImg As Long          '是否显示大图,0-不显示；1-鼠标移动时显示；2-鼠标单击显示独立窗口
Private mdblBigImgZoom As Double        '报告大图放大倍数
Private mintImageDblClick As Integer    '缩略图双击后的操作 0--直接写入报告；1--打开图片编辑窗口
Private mblnEditable As Boolean         '是否可以编辑内容
Private mintMoustType As Integer        '鼠标工作类型
Private mblnUserInvoke As Boolean       '是否用户操作触发
Private mblnMoved As Boolean            '是否已经转储
Private mintCurImgIndex As Integer      '当前选中的图像
Private mintShowPhotoNumber As Integer  '当前界面能够显示出图像的最佳数量
Private mlngModule As Long

Public mSelMiniImg As DicomImage
Private mSelReportImg As DicomImage
Private mSelViewerIndex As Integer  '当前被选中的报告图象框ID，从1开始计数
Private mselReportImgIndex As Integer   '当前被选中的报告图像ID，从1开始计数
Private mdblMarkZoom As Double          '当前标记图中实际像素和标记之间的缩放比例
Private lngColor(10) As Long             '标记图中圆形编号使用的9个颜色
Private mlngCY1 As Long                 '标记图的高度
Private mlngMarkW As Long               '标记图的宽度
Private mlngCY2 As Long                 '报告图的高度
Private mlngRptImgW As Long             '报告图的宽度
Private mlngCY3 As Long                 '缩略图图的高度

Public pMarkModified As Boolean        '标记图的标记有改动
Public pImageModified As Boolean       '记录报告图像是否修改，如果没有修改，则保存报告的时候不再保存图像
Public pobjMarks As cPicMarks          '当前标记图的标注对象
Public pMarkImageID As Long            '当前标记图在数据库表“电子病历内容”表中的ID
Public pTableID As String              '当前图像所在表格的ID串，用“;”分隔。


Private mintShowMarkImage As Integer   '是否显示标记图   0-隐藏标记图  1-显示标记图
Private mblnIsInitFace As Boolean        '是否已经加载窗体

Private blnLoadImages As Boolean        '记录本次刷新是否加载了图像


Private mdcmGlobal As New DicomGlobal    '定义UIDRoot=1

Private mblnUseActiveVideo As Boolean

Private Enum MarkType
    自动编号 = 0: 编号1: 编号2: 编号3: 编号4: 编号5: 编号6
End Enum

Property Get ImageCount() As Long
    If mblnUseActiveVideo Then
'        ImageCount = mobjStudyImage.Images.CurImageCount
        ImageCount = ucMiniImageViewer.CurImageCount
    Else
        ImageCount = dcmMiniImageC.Images.Count
    End If
End Property

Property Get dcmImages() As Object
    If mblnUseActiveVideo Then
'        Set dcmImages = mobjStudyImage.Images.ImgViewer.Images
        Set dcmImages = ucMiniImageViewer.ImgViewer.Images
    Else
        Set dcmImages = dcmMiniImageC.Images
    End If
End Property

Public Sub MovePage(ByVal lngPageType As TMoveType)
'移动缩略图页面
    If mblnUseActiveVideo Then
        ucMiniImageViewer.MovePage (lngPageType)
    End If
End Sub


Public Sub zlRefresh(ByVal lngAdviceID As Long, FileID As Long, _
        ReportID As Long, blnSingleWindow As Boolean, lngShowBigImg As Long, _
        dblBigImgZoom As Double, intImageDblClick As Integer, blnEditable As Boolean, _
        ByVal blnMoved As Boolean, ByVal intMinImageCount As Integer, blnFormIsSelected As Boolean, ByVal lngModule As Long)
    Dim i As Integer
    Dim intShowMarkImage As Integer
    
    mlngAdviceID = lngAdviceID
    mlngFileID = FileID
    mlngReportID = ReportID
    mlngShowBigImg = lngShowBigImg
    mdblBigImgZoom = dblBigImgZoom
    mintImageDblClick = intImageDblClick
    mblnEditable = blnEditable
    mblnMoved = blnMoved
    mintShowPhotoNumber = intMinImageCount
    mlngModule = lngModule
    mblnSingleWindow = blnSingleWindow
    
    intShowMarkImage = DecideMarkImagesVisible    '判断标记图是否可见
    
    
    ucMiniImageViewer.BigImageWay = lngShowBigImg
    ucMiniImageViewer.MouseMoveZoom = dblBigImgZoom
    ucMiniImageViewer.ShowPopup = False
    
    
    '判断如果是 独立窗口 或者 没有加载过窗体 或者 标记图状态已经改变，则重新加载初始化界面
    If mblnSingleWindow Or (Not mblnIsInitFace Or (mintShowMarkImage <> intShowMarkImage)) Then
        mintShowMarkImage = intShowMarkImage
        Call InitLoaclParas     '读取本机参数
        Call InitFaceScheme     '初始化窗体界面
    End If
    
    
    
    '重新初始化内部参数
    pTableID = ""
    pMarkImageID = 0
    pImageModified = False
    pMarkModified = False
    dcmMark.Images.Clear
    
    If Not (pobjMarks Is Nothing) Then
        For i = 1 To pobjMarks.Count
            pobjMarks.Remove 1
        Next i
    End If
    
    
    '标记本次刷新还没有加载图像
    blnLoadImages = False
    '如果窗体是正在被显示的，则加载图像
    If blnFormIsSelected = True And Me.Visible Then
        '根据需要加载图像
        Call LoadImages
    Else
        Call ClearReportImages
    End If
    
    '设置界面控件是否可以编辑
    picMark.Enabled = mblnEditable
    picReportImage.Enabled = mblnEditable
    picMiniImageC.Enabled = mblnEditable
    picMiniViewer.Enabled = mblnEditable
End Sub

Private Sub ClearReportImages()
    Dim i As Integer
    
    '初始化各个对象
    For i = 1 To dcmReportImage.Count - 1
        Unload dcmReportImage(i)
    Next i
    dcmMark.Images.Clear
End Sub

Public Sub RefPacsPic()
    '读取和显示当前可选报告图像
    Call LoadMiniImages
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    
    Select Case control.ID
        Case comMenu_Cap_Process '图像处理
            Call OpenImageProcessWind
        Case conMenu_Cap_DevSet
            Call ucMiniImageViewer.ShowPageConfig
        Case conMenu_PacsReport_DelImage    '删除图像
            If dcmReportImage(mSelViewerIndex).Images.Count >= mselReportImgIndex And mselReportImgIndex <> 0 Then
                dcmReportImage(mSelViewerIndex).Images.Remove mselReportImgIndex
                Call picReportImage_Resize
                pImageModified = True
            End If
        Case conMenu_PacsReport_MoveUp      '前移图像
            If mselReportImgIndex > 1 And dcmReportImage(mSelViewerIndex).Images.Count >= mselReportImgIndex Then
                dcmReportImage(mSelViewerIndex).Images.Move mselReportImgIndex, mselReportImgIndex - 1
                pImageModified = True
            End If
        Case conMenu_PacsReport_MoveDown    '后移图像
            If mselReportImgIndex > 0 And dcmReportImage(mSelViewerIndex).Images.Count > mselReportImgIndex Then
                dcmReportImage(mSelViewerIndex).Images.Move mselReportImgIndex, mselReportImgIndex + 1
                pImageModified = True
            End If
        Case conMenu_PacsReport_DelMarks    '清除标注
            If dcmMark.Images.Count > 0 Then
                dcmMark.Images(1).Labels.Clear
                dcmMark.Refresh
                For i = 1 To pobjMarks.Count
                    pobjMarks.Remove 1
                Next i
                pMarkModified = True
            End If
        Case conMenu_View_Refresh           '刷新
            '读取和显示当前可选报告图像
            Call LoadMiniImages
        Case conMenu_PacsReport_DelMiniImage    '删除报告图
            
        Case conMenu_PacsReport_SelMiniImage    '提取报告图
            Dim resImages As DicomImages
            
            Set resImages = frmSelectRepImage.ShowMe(Me, mlngAdviceID, mlngShowBigImg, mdblBigImgZoom)
            '把当前图形添加到图象框中
            If resImages.Count > 0 Then
                For i = 1 To resImages.Count
                    dcmReportImage(mSelViewerIndex).Images.Add resImages(i)
                    dcmReportImage(mSelViewerIndex).Images(dcmReportImage(mSelViewerIndex).Images.Count).BorderColour = vbWhite
                Next i
                dcmReportImage(mSelViewerIndex).CurrentIndex = 1
                Call picReportImage_Resize
                pImageModified = True
            End If
    End Select
End Sub

Private Sub chkMark_Click(Index As Integer)
    Dim i As Integer
    If mblnUserInvoke = False Then
        mblnUserInvoke = True
    Select Case Index
        Case 0
            mintMoustType = MarkType.自动编号
        Case 1
            mintMoustType = MarkType.编号1
        Case 2
            mintMoustType = MarkType.编号2
        Case 3
            mintMoustType = MarkType.编号3
        Case 4
            mintMoustType = MarkType.编号4
        Case 5
            mintMoustType = MarkType.编号5
        Case 6
            mintMoustType = MarkType.编号6
    End Select
    For i = 0 To 6
        chkMark(i).value = 0
    Next i
    chkMark(Index).value = 1
    mblnUserInvoke = False
    End If
End Sub

Private Sub dcmMark_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim lTemp As DicomLabel
    Dim strNum As Integer
    
    If Button = 1 And dcmMark.Images.Count > 0 And picMark.MousePointer = 99 Then
        '画标注
        '两种类型的标注，一种是直接自动编号，另一种是手工编号
        pobjMarks.Add pobjMarks.Count + 1
        pobjMarks(pobjMarks.Count).Selected = False
        pobjMarks(pobjMarks.Count).类型 = 6     '圆形编号
        If mintMoustType = MarkType.自动编号 Then
            pobjMarks(pobjMarks.Count).内容 = pobjMarks.Count
        Else
            Select Case mintMoustType
                Case MarkType.编号1
                    pobjMarks(pobjMarks.Count).内容 = 1
                Case MarkType.编号2
                    pobjMarks(pobjMarks.Count).内容 = 2
                Case MarkType.编号3
                    pobjMarks(pobjMarks.Count).内容 = 3
                Case MarkType.编号4
                    pobjMarks(pobjMarks.Count).内容 = 4
                Case MarkType.编号5
                    pobjMarks(pobjMarks.Count).内容 = 5
                Case MarkType.编号6
                    pobjMarks(pobjMarks.Count).内容 = 6
            End Select
        End If
        '点集没有留空
        Set lTemp = New DicomLabel
        lTemp.Left = X
        lTemp.Top = Y
        lTemp.Width = 20
        lTemp.Height = 20
        lTemp.ImageTied = True
        lTemp.Rescale dcmMark.Images(1)
        pobjMarks(pobjMarks.Count).X1 = lTemp.Left / mdblMarkZoom
        pobjMarks(pobjMarks.Count).Y1 = lTemp.Top / mdblMarkZoom
        pobjMarks(pobjMarks.Count).X2 = pobjMarks(pobjMarks.Count).X1
        pobjMarks(pobjMarks.Count).Y2 = pobjMarks(pobjMarks.Count).Y1
        pobjMarks(pobjMarks.Count).填充色 = lngColor(pobjMarks.Count Mod 9 + 1)
        pobjMarks(pobjMarks.Count).填充方式 = -2
        '线条色留空，字体色留空
        pobjMarks(pobjMarks.Count).线型 = 1
        pobjMarks(pobjMarks.Count).线宽 = 1
        Set pobjMarks(pobjMarks.Count).字体 = New StdFont '  "宋体"
        drawPicMarks dcmMark.Images(1), pobjMarks
        dcmMark.Refresh
        
        pMarkModified = True
    End If
End Sub

Private Sub dcmMark_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If dcmMark.Images.Count = 1 Then
        '设置鼠标
        If dcmMark.ImageXPosition(X, Y) > 0 And dcmMark.ImageXPosition(X, Y) < dcmMark.Images(1).SizeX _
           And dcmMark.ImageYPosition(X, Y) > 0 And dcmMark.ImageYPosition(X, Y) < dcmMark.Images(1).SizeY Then
            picMark.MousePointer = 99
            picMark.MouseIcon = listCur.ListImages("Pen").Picture
        Else
            picMark.MousePointer = 0
            picMark.MouseIcon = Nothing
        End If
    End If
End Sub

Private Sub dcmMark_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 2 Then ShowPopupMark
End Sub

Private Sub OpenImageProcessWind()
    Call frmReportImageEdit.zlShowMe(mSelMiniImg, Me, mintCurImgIndex, mSelViewerIndex, mlngModule)
End Sub

Private Sub dcmMiniImageC_DblClick()
    
    If dcmMiniImageC.Images.Count > 0 And mSelViewerIndex <> 0 And dcmReportImage.Count > 1 Then
        '判断当前双击的操作动作
        If mintImageDblClick = 0 Then   '直接写入报告
            Dim dcmImage As DicomImage
            Set dcmImage = mSelMiniImg
            
            '调用将当前图形添加到图象框过程
            Call DcmAddImage(dcmImage, mSelViewerIndex)
            
        Else                            '先打开图片编辑窗口
            '先关闭大图窗口
            ReleaseCapture      '解锁鼠标
            frmShowImg.HideMe
            
            Call OpenImageProcessWind
        End If

    End If
End Sub

Public Sub DcmAddImage(dcmImage As DicomImage, SelViewerIndex As Integer)
'把当前图形添加到图象框中
    If Not dcmImage Is Nothing Then
        dcmReportImage(SelViewerIndex).Images.Add dcmImage
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).BorderColour = vbWhite
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).Tag = mdcmGlobal.NewUID & ".jpg"
        dcmReportImage(SelViewerIndex).CurrentIndex = 1
        Call picReportImage_Resize
        pImageModified = True
    End If
End Sub

Private Sub dcmMiniImageC_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    
    If dcmMiniImageC.Images.Count > 0 Then
        For i = 1 To dcmMiniImageC.Images.Count
            dcmMiniImageC.Images(i).BorderColour = vbWhite
            
        Next i
        
        i = dcmMiniImageC.ImageIndex(X, Y)
        If i = 0 Then
            Set mSelMiniImg = dcmMiniImageC.Images(1)
        Else
            Set mSelMiniImg = dcmMiniImageC.Images(i)
        End If
        
        mSelMiniImg.BorderColour = vbRed
        
        mintCurImgIndex = i
        
        '判断是否需要显示大图
        If mlngShowBigImg = 2 Then
            frmShowImg.ShowMe mSelMiniImg, Me, 2, 0, 0
        End If
    End If
End Sub

Private Sub dcmMiniImageC_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim blnShowImg As Boolean
    Dim intCurrImg As Integer

    If dcmMiniImageC.Images.Count <= 0 Or mlngShowBigImg <> 1 Then Exit Sub
    
    '判断是否需要显示图像
    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= dcmMiniImageC.Width) And _
       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= dcmMiniImageC.Height) Then
        blnShowImg = True
    End If
    If blnShowImg Then      '显示图像
        SetCapture dcmMiniImageC.hWnd    '锁定鼠标
        
        intCurrImg = dcmMiniImageC.ImageIndex(X, Y)
        If intCurrImg <> 0 Then
            '加载图像并显示
            frmShowImg.ShowMe dcmMiniImageC.Images(intCurrImg), Me, 1, 0, 0, mdblBigImgZoom
        Else
            frmShowImg.HideMe
        End If
    Else        '关闭图像显示
        ReleaseCapture      '解锁鼠标
        frmShowImg.HideMe
    End If
End Sub

Private Sub dcmMiniImageC_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mlngShowBigImg = 1 Then  '关闭大图显示
        ReleaseCapture      '解锁鼠标
        frmShowImg.HideMe
    End If
    
    If Button = 2 Then Call ShowPopupImage(True)
End Sub

Private Sub dcmReportImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    
'    If dcmReportImage(Index).Images.Count = 0 Then Exit Sub
    
    mSelViewerIndex = Index
    mselReportImgIndex = dcmReportImage(Index).ImageIndex(X, Y)
    
    For i = 1 To dcmReportImage.Count - 1
        dcmReportImage(i).Labels(1).ForeColour = vbWhite
        dcmReportImage(i).Refresh
    Next i
    dcmReportImage(Index).Labels(1).ForeColour = vbRed
    dcmReportImage(Index).Refresh
    
    If mselReportImgIndex <> 0 Then
        For i = 1 To dcmReportImage(Index).Images.Count
            dcmReportImage(Index).Images(i).BorderColour = vbWhite
        Next i
        dcmReportImage(Index).Images(mselReportImgIndex).BorderColour = vbBlue
    End If
    
    
End Sub

Private Sub dcmReportImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    If Button = 2 Then Call ShowPopupImage(False)
End Sub

Private Sub ShowPopupImage(ByVal blnIsDcmMiniImage As Boolean)
'------------------------------------------------
'功能：创建鼠标右键弹出菜单
'------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    If mblnUseActiveVideo Then
'        If mobjStudyImage.Images.CurImageCount < 1 Then Exit Sub
        If ucMiniImageViewer.CurImageCount < 1 Then Exit Sub
    Else
        '如果缩略图没有图像，则禁止右键弹出
        If Me.dcmMiniImageC.Images.Count < 1 Then Exit Sub
    End If

    
    '鼠标右键弹出菜单
    Set cbrToolBar = cbrMain.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        If Not blnIsDcmMiniImage Then
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelImage, "删除")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_MoveUp, "前移")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_MoveDown, "后移")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_SelMiniImage, "提取报告图")
         Else
            Set cbrControl = .Add(xtpControlButton, comMenu_Cap_Process, "图像处理")
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "分页设置")
            Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
         End If
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub ShowPopupMark()
    '------------------------------------------------
'功能：创建鼠标右键弹出菜单
'------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '鼠标右键弹出菜单
    Set cbrToolBar = cbrMain.Add("鼠标右键", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelMarks, "清除标注")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub


Private Sub Form_Activate()
    '根据需要加载图像
    
    '注：在Form的Activate和Paint时间中必须调用LoadImages方法
    '因为如果只在Activate方法中调用LoadImages方法，可能造成报告图不会在第一时间显示，必须用鼠标点击一下报告图才会显示
    '如果只在Paint方法中调用LoadImages方法，由于该方法中使用了UnLoad卸载控件数组，可能造成“不能从该上下文中卸载”的错误
    
    Call LoadImages
End Sub

Private Sub Form_Load()
    
    '标记本次刷新已经加载图像
    blnLoadImages = True
    
    '标记窗体已经首次加载
    mblnIsInitFace = False
        
    mintMoustType = MarkType.自动编号

    
    '设置默认颜色
    lngColor(1) = RGB(186, 186, 186)
    lngColor(2) = RGB(255, 215, 0)
    lngColor(3) = RGB(255, 0, 255)
    lngColor(4) = RGB(255, 0, 130)
    lngColor(5) = RGB(0, 255, 0)
    lngColor(6) = RGB(130, 255, 255)
    lngColor(7) = RGB(255, 255, 0)
    lngColor(8) = RGB(0, 0, 255)
    lngColor(9) = RGB(0, 160, 0)
    
    '定义UIDRoot=1
    mdcmGlobal.RegString("UIDRoot") = "1"
    
End Sub

Public Sub MouseWheel(intDirection As Integer)
'处理鼠标滚轮的事件
'参数：intDirection --- 鼠标滚轮的方向；1--向上；0--向下
    
    On Error Resume Next
    
    If vscrollMini.Visible = False Then Exit Sub
    
    If intDirection = 1 Then '上翻一页
        If vscrollMini.value - 1 < 1 Then
            vscrollMini.value = 1
        Else
            vscrollMini.value = vscrollMini.value - 1
        End If
    Else        '下翻一页
        If vscrollMini.value + 1 > vscrollMini.Max Then
            vscrollMini.value = vscrollMini.Max
        Else
            vscrollMini.value = vscrollMini.value + 1
        End If
    End If
End Sub

Public Sub subDispScroll()
'------------------------------------------------
'功能：自动判断是否需要显示或隐藏滚动条
'返回：无，直接显示或隐藏滚动条。
'------------------------------------------------
    Dim ii As Integer
    
    If dcmMiniImageC.Images.Count > dcmMiniImageC.MultiColumns * dcmMiniImageC.MultiRows Then       '图像总数大于显示数，显示滚动条
        '摆放滚动条位置，并显示滚动条
        vscrollMini.Move dcmMiniImageC.Width - vscrollMini.Width, dcmMiniImageC.Top, vscrollMini.Width, dcmMiniImageC.Height
        vscrollMini.Visible = True
        vscrollMini.ZOrder
        vscrollMini.Refresh
        
        ''''''''''''''''''[关于滚动条需要单独仔细分析]'''''''''''''''''''''''''
        vscrollMini.Min = 1
        vscrollMini.Max = dcmMiniImageC.Images.Count - dcmMiniImageC.MultiColumns * dcmMiniImageC.MultiRows + 1
        If vscrollMini.Max < 1 Then vscrollMini.Max = 1
        vscrollMini.LargeChange = dcmMiniImageC.MultiColumns * dcmMiniImageC.MultiRows
        If dcmMiniImageC.CurrentIndex > vscrollMini.Max Then
            vscrollMini.value = vscrollMini.Max
            dcmMiniImageC.CurrentIndex = vscrollMini.Max
        Else
            vscrollMini.value = dcmMiniImageC.CurrentIndex
        End If
    Else                '图像数少于可显示数，隐藏滚动条
'        ii = dcmMiniature.Images.Count - dcmMiniature.MultiColumns * dcmMiniature.MultiRows + 1
'        If dcmMiniature.Images.Count - dcmMiniature.CurrentIndex + 1 < dcmMiniature.MultiColumns * dcmMiniature.MultiRows Then
'            dcmMiniature.CurrentIndex = IIf(ii < 1, 1, ii)
'        End If
'        vscrollMini.Value = dcmMiniature.CurrentIndex
        vscrollMini.Visible = False
    End If
    
    If vscrollMini.Visible = True Then
        dcmMiniImageC.Width = dcmMiniImageC.Width - vscrollMini.Width - 20
        
        vscrollMini.Height = dcmMiniImageC.Height - 40
        vscrollMini.Left = dcmMiniImageC.Width - 20
    Else
        dcmMiniImageC.Width = dcmMiniImageC.Width
    End If
End Sub


Private Sub InitLoaclParas()
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage"
    End If
    
    ucMiniImageViewer.PageImgCount = Val(GetSetting("ZLSOFT", strRegPath, "报告缩略图数量", 5))
    
    '读取标记图区域，报告图区域 和缩略图区域的高度
    mlngCY1 = GetSetting("ZLSOFT", strRegPath, "CY1", 180)
    mlngMarkW = GetSetting("ZLSOFT", strRegPath, "MarkW", 300)
    mlngCY2 = GetSetting("ZLSOFT", strRegPath, "CY2", 400)
    mlngRptImgW = GetSetting("ZLSOFT", strRegPath, "RptImgW", 100)
    mlngCY3 = GetSetting("ZLSOFT", strRegPath, "CY3", 200)
End Sub

Private Sub Form_Paint()
    '根据需要加载图像
    Call LoadImages
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReportImage"
    End If
    
    Call SaveSetting("ZLSOFT", strRegPath, "报告缩略图数量", ucMiniImageViewer.PageImgCount)
    
    '保存标记图区域，报告图区域和缩略图区域的高度
    '285是Pane的标题高度，使用了标题，就需要加回这个高度
    SaveSetting "ZLSOFT", strRegPath, "CY1", picMark.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "MarkW", picMark.Width
    SaveSetting "ZLSOFT", strRegPath, "CY2", picReportImage.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "RptImgW", picReportImage.Width
    SaveSetting "ZLSOFT", strRegPath, "CY3", picMiniImageC.Height + 285
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
    End If
    SaveSetting "ZLSOFT", strRegPath, "CX3", Me.Width
    SaveSetting "ZLSOFT", strRegPath, "CY3", Me.Height
End Sub

Private Sub picMark_Resize()
    If picMark.Height = 0 Or picMark.Width = 0 Then Exit Sub
    
    On Error Resume Next
    
    '判断宽高比
    If picMark.Width / picMark.Height > 2 Then  '数字标记放在右边
        dcmMark.Left = 0
        dcmMark.Top = 0
        dcmMark.Width = Abs(picMark.ScaleWidth - picNumMark.ScaleWidth - 50)
        dcmMark.Height = picMark.ScaleHeight
        
        picNumMark.Left = dcmMark.Width
        If picMark.Height > picNumMark.Height Then
            picNumMark.Top = (picMark.ScaleHeight - picNumMark.ScaleHeight) / 2
        Else
            picNumMark.Top = 0
        End If
    Else    '数字标记放在下面
        dcmMark.Left = 0
        dcmMark.Top = 0
        dcmMark.Width = picMark.ScaleWidth
        dcmMark.Height = Abs(picMark.ScaleHeight - picNumMark.ScaleHeight - 50)
        
        If picMark.Width > picNumMark.Width Then
            picNumMark.Left = (picMark.ScaleWidth - picNumMark.ScaleWidth) / 2
        Else
            picNumMark.Left = 0
        End If
        picNumMark.Top = dcmMark.Height
    End If
End Sub

Private Sub picMiniImageC_Resize()
'    If picMiniImage.Width < 50 Or picMiniImage.Height < 50 Then Exit Sub
'    dcmMiniImage.Left = 0
'    dcmMiniImage.Top = 0
'    dcmMiniImage.Width = picMiniImage.Width - 50
'    dcmMiniImage.Height = picMiniImage.Height - 50

    Dim iRows As Integer
    Dim iCols As Integer
    
    On Error Resume Next
    
    dcmMiniImageC.Left = 0
    dcmMiniImageC.Top = 0
    dcmMiniImageC.Width = picMiniImageC.Width
    dcmMiniImageC.Height = picMiniImageC.Height
    
    '自动对图像做布局
    '计算缩略图的图像布局
    If mintShowPhotoNumber < dcmMiniImageC.Images.Count Then
        ResizeRegion mintShowPhotoNumber, dcmMiniImageC.Width, picMiniImageC.Height, iRows, iCols
    Else
        ResizeRegion dcmMiniImageC.Images.Count, dcmMiniImageC.Width, dcmMiniImageC.Height, iRows, iCols
    End If
    
    dcmMiniImageC.MultiColumns = iCols
    dcmMiniImageC.MultiRows = iRows
    '处理滚动条
    'If vscrollMini.Visible = True Then
    dcmMiniImageC.Width = picMiniImageC.Width - vscrollMini.Width - 20
    
    vscrollMini.Height = dcmMiniImageC.Height - 40
    vscrollMini.Left = dcmMiniImageC.Width - 20
    'End If
End Sub


Private Sub InitFaceScheme()
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    
    With dkpMain
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    If mintShowMarkImage = 1 Then
        picMark.Visible = True
        dcmMark.Visible = True
        picNumMark.Visible = True
        
        Set Pane1 = dkpMain.CreatePane(1, mlngMarkW, mlngCY1, DockTopOf, Nothing)
        Pane1.Title = "标记图"
        Pane1.Handle = picMark.hWnd
        Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        '根据宽高比，摆放报告图的位置
        If ((mlngCY1 = mlngCY2) And (mlngMarkW + mlngRptImgW > mlngCY1)) _
            Or (((mlngCY1 <> mlngCY2)) And (mlngMarkW + mlngRptImgW > mlngCY1 + mlngCY2)) Then
            Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockLeftOf, Pane1)
        Else
            Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockBottomOf, Pane1)
        End If
    Else
        picMark.Visible = False
        dcmMark.Visible = False
        picNumMark.Visible = False
        
        Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockTopOf, Nothing)
    End If
    
'    If mobjStudyImage Is Nothing Then
'        Set mobjStudyImage = New clsStudyImages
'    End If
    
    mblnUseActiveVideo = False
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Or G_LNG_PATHSTATION_MODULE Then
        mblnUseActiveVideo = GetSetting("ZLSOFT", "公共模块", "UseActiveVideo", "true")
        Call SaveSetting("ZLSOFT", "公共模块", "UseActiveVideo", mblnUseActiveVideo)
    End If

    Pane2.Title = "报告图"
    Pane2.Handle = picReportImage.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    picMiniImageC.Visible = Not mblnUseActiveVideo
    picMiniViewer.Visible = mblnUseActiveVideo

    Set Pane3 = dkpMain.CreatePane(3, 0, mlngCY3, DockBottomOf, Nothing)
    Pane3.Title = "缩略图"
    Pane3.Handle = IIf(mblnUseActiveVideo, picMiniViewer.hWnd, picMiniImageC.hWnd)
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    mblnIsInitFace = True
End Sub


Private Function LoadMiniImages() As Boolean
    Dim lngMsgHwnd As Long
    
    
    If mblnUseActiveVideo Then
'        lngMsgHwnd = mobjStudyImage.hWnd
'
'        Call mobjStudyImage.RefreshImages(mlngAdviceID, mlngAdviceID, mblnMoved, True)

        Call ucMiniImageViewer.RefreshImage(0, mlngAdviceID, mblnMoved, True)
    Else
        Call GetRptImages(dcmMiniImageC, mlngAdviceID, mblnMoved)
    
        Call AdjustDicomViewerLayout
    End If

End Function


Private Sub AdjustDicomViewerLayout()
'------------------------------------------------
'功能：将图像添加到缩略图dcmMiniature中
'参数：img－－输入的DICOM图像
'返回：无，直接将图像添加到缩略图dcmMiniature中
'------------------------------------------------
    Dim iRows As Integer
    Dim iCols As Integer
    
    '计算缩略图的图像布局
    If mintShowPhotoNumber < dcmMiniImageC.Images.Count + 1 Then
        ResizeRegion mintShowPhotoNumber, dcmMiniImageC.Width, dcmMiniImageC.Height, iRows, iCols
    Else
        ResizeRegion dcmMiniImageC.Images.Count + 1, dcmMiniImageC.Width, dcmMiniImageC.Height, iRows, iCols
    End If
            
    dcmMiniImageC.MultiColumns = iCols
    dcmMiniImageC.MultiRows = iRows

    
'    '根据缩略图的检查UID和序列UID，修改img的值
'    subUniteUID img
'    dcmMiniature.Images.Add img
    
'    '处理缩略图中被选中的状态
'    If mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniature.Images.Count Then
'        dcmMiniature.Images(mintCurImgIndex).BorderColour = vbWhite
'    End If
    
    If dcmMiniImageC.Images.Count > 0 Then
         With dcmMiniImageC.Images(1)
            .BorderWidth = 1
            .BorderStyle = 6
            .BorderColour = vbRed
        End With
    End If
    
    mintCurImgIndex = 1
    
'    mintCurImgIndex = dcmMiniature.Images.Count
    '显示滚动条
    Call subDispScroll
End Sub


Private Sub LoadReportImages()
    
    On Error GoTo errH
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim iRImageCount As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim blnGetTable As Boolean
    Dim i As Integer
    
    '初始化各个对象
    Call ClearReportImages
        
    '如果存在报告内容，则从报告内容中读取报告图和标记图，否则从报告单格式中读取标记图
    If mlngReportID <> 0 Then
        strSql = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
            " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By 对象序号"
        If mblnMoved = True Then
            strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngReportID)
    Else
        strSql = "Select Id As 表格Id From 病历文件结构" & vbNewLine & _
            " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngFileID)
    End If
    
    iRImageCount = 0
    Do While Not rsTemp.EOF
    
        Set cTable = New cEPRTable
        If mlngReportID = 0 Then
            blnGetTable = cTable.GetTableFromDB(cprET_病历文件定义, mlngFileID, Val("" & rsTemp!表格ID))
        Else
            blnGetTable = cTable.GetTableFromDB(cprET_单病历审核, mlngReportID, Val("" & rsTemp!表格ID))
        End If
        If blnGetTable Then
            
            '记录图像所在表格ID
            If pTableID = "" Then
                pTableID = cTable.ID
            Else
                pTableID = pTableID & ";" & cTable.ID
            End If
            
            '创建viewer
            iRImageCount = iRImageCount + 1
            Load dcmReportImage(iRImageCount)
            dcmReportImage(iRImageCount).BorderStyle = 1
            dcmReportImage(iRImageCount).Labels.AddNew
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LabelType = doLabelRectangle
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Left = 1
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Top = 1
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LineWidth = 2
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).ForeColour = vbWhite
            dcmReportImage(iRImageCount).Visible = True
            
            mSelViewerIndex = iRImageCount

            '记录图像框的宽度和高度，该宽高比例用于后续对图像行列布局
            If cTable.ExtendTag <> "" Then
                If Val(Split(cTable.ExtendTag, "|")(0)) = 0 Then
                    dcmReportImage(iRImageCount).Tag = cTable.Width & "|" & cTable.Height
                Else
                    dcmReportImage(iRImageCount).Tag = (cTable.Width - Val(Split(cTable.ExtendTag, "|")(1)) - 30) & "|" & cTable.Height
                End If
            Else
                dcmReportImage(iRImageCount).Tag = cTable.Width & "|" & cTable.Height
            End If
            
            
            For i = 1 To cTable.Pictures.Count
                strPicFile = App.Path & "\PACSPic" & i & ".JPG"
                If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True

                Set oPicture = cTable.Pictures(i).OrigPic
                SavePicture oPicture, strPicFile
                If objFile.FileExists(strPicFile) Then
                    '显示标记图和报告图
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture And dcmMark.Images.Count = 0 Then

                        '只处理第一个标记图
                        dcmMark.Images.AddNew
                        
                        dcmMark.Images(1).FileImport strPicFile, "BMP"
                        dcmMark.Images(1).Tag = cTable.Pictures(i).ID
                        '保存标记图基础数据
                        Set pobjMarks = cTable.Pictures(i).PicMarks
                        pMarkImageID = cTable.Pictures(i).ID

                        mdblMarkZoom = dcmMark.Images(1).SizeX / cTable.Pictures(i).Width * Screen.TwipsPerPixelX
                        '显示标注
                        If cTable.Pictures(i).PicMarks.Count > 0 Then
                            drawPicMarks dcmMark.Images(1), cTable.Pictures(i).PicMarks
                        End If
                    Else

                        dcmReportImage(iRImageCount).Images.AddNew
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).FileImport strPicFile, "BMP"
                        If cTable.Pictures(i).PicName = "" Then
                            dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).Tag = mdcmGlobal.NewUID & ".jpg"
                        Else
                            dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).Tag = cTable.Pictures(i).PicName
                        End If
                        
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).BorderWidth = 3
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).BorderColour = vbWhite
                        dcmReportImage(iRImageCount).CurrentIndex = 1
                        mselReportImgIndex = 1
                    End If
                    '删除临时图像
                    Kill strPicFile
                End If
            Next
        End If
        
        rsTemp.MoveNext
    Loop
    If dcmReportImage.Count > 1 Then dcmReportImage(1).Labels(1).ForeColour = vbRed
    Call picReportImage_Resize
    

Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function DecideMarkImagesVisible() As Integer
'------------------------------------------------
'功能：判断当前选中检查标记图是否可见
'参数：无
'返回：int类型，1-显示标记图  2-隐藏标记图
'-----------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim objFile As New Scripting.FileSystemObject
    Dim blnGetTable As Boolean
    Dim i As Integer
    
        
    '如果存在报告内容，则从报告内容中读取数据，否则从报告单格式中读取数据
    If mlngReportID <> 0 Then
        strSql = "Select Id As 表格Id From 电子病历内容" & vbNewLine & _
            " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By 对象序号"
        If mblnMoved = True Then
            strSql = Replace(strSql, "电子病历内容", "H电子病历内容")
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "报告内容中读取", mlngReportID)
    Else
        strSql = "Select Id As 表格Id From 病历文件结构" & vbNewLine & _
            " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "从报告单格式中读取", mlngFileID)
    End If
    
    Do While Not rsTemp.EOF
        Set cTable = New cEPRTable
        If mlngReportID = 0 Then
            blnGetTable = cTable.GetTableFromDB(cprET_病历文件定义, mlngFileID, Val("" & rsTemp!表格ID))
        Else
            blnGetTable = cTable.GetTableFromDB(cprET_单病历审核, mlngReportID, Val("" & rsTemp!表格ID))
        End If
        
        If blnGetTable Then
            If cTable.Pictures.Count > 0 Then
                For i = 1 To cTable.Pictures.Count
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                        DecideMarkImagesVisible = 1
                        Exit Do
                    Else
                        DecideMarkImagesVisible = 0
                    End If
                Next
            Else
                DecideMarkImagesVisible = 0
            End If
        End If
        rsTemp.MoveNext
    Loop

End Function


Private Sub drawPicMarks(img As DicomImage, thisMarks As cPicMarks)
'显示标注，只支持数字编号标注
    Dim i As Integer
    Dim iLabelCount As Integer
    
    img.Labels.Clear
    For i = 1 To thisMarks.Count
        If thisMarks(i).类型 = 6 Then   '圆形编号
            With thisMarks(i)
                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelEllipse
                img.Labels(iLabelCount).BackColour = IIf(.填充色 = 0, vbYellow, .填充色)
                img.Labels(iLabelCount).Transparent = False
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Width = 20
                img.Labels(iLabelCount).Height = 20
                img.Labels(iLabelCount).ImageTied = True
                
                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelEllipse
                img.Labels(iLabelCount).ForeColour = vbBlack
                img.Labels(iLabelCount).Transparent = True
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Width = 20
                img.Labels(iLabelCount).Height = 20
                img.Labels(iLabelCount).ImageTied = True

                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelText
                img.Labels(iLabelCount).Transparent = True
                img.Labels(iLabelCount).ForeColour = vbBlack
                img.Labels(iLabelCount).FontSize = 11
                img.Labels(iLabelCount).FontName = "Arial Bold"
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).AutoSize = True
                img.Labels(iLabelCount).Text = .内容
                img.Labels(iLabelCount).ImageTied = True
            End With
        End If
    Next i
End Sub
 
Private Sub picMiniViewer_Resize()
On Error Resume Next
    ucMiniImageViewer.Left = 0
    ucMiniImageViewer.Top = 0
    ucMiniImageViewer.Width = picMiniViewer.ScaleWidth
    ucMiniImageViewer.Height = picMiniViewer.ScaleHeight
End Sub

Private Sub picReportImage_Resize()
    Dim i As Integer
    Dim rectH As Long, rectW As Long    '图象框可以使用的区域宽高
    Dim picH As Long, picW As Long      '图像实际宽高，作为比例使用
    Dim iCols As Integer, iRows As Integer
    Dim dImg As DicomImage
    
    If dcmReportImage.Count = 1 Then Exit Sub
    
    On Error Resume Next
    
    '首先计算每个图象框可占用的最大宽高
    
    rectH = picReportImage.Height / (dcmReportImage.Count - 1)
    rectW = picReportImage.Width
    If rectH < 100 Or rectW < 100 Then Exit Sub
    
    For i = 1 To dcmReportImage.Count - 1
        '按照图像比例，计算图象框的真实宽度和高度
        picW = Val(Split(dcmReportImage(i).Tag, "|")(0))
        picH = Val(Split(dcmReportImage(i).Tag, "|")(1))
        
        dcmReportImage(i).Height = rectH - 100
        dcmReportImage(i).Width = rectW - 100
        
        dcmReportImage(i).Left = 0
        dcmReportImage(i).Top = rectH * (i - 1)
        
        dcmReportImage(i).Labels(1).Width = Abs(dcmReportImage(i).Width / Screen.TwipsPerPixelX - 2)
        dcmReportImage(i).Labels(1).Height = Abs(dcmReportImage(i).Height / Screen.TwipsPerPixelY - 1)

        
        '调整图像显示布局
        ResizeRegion dcmReportImage(i).Images.Count, picW, picH, iRows, iCols
        dcmReportImage(i).MultiColumns = iCols
        dcmReportImage(i).MultiRows = iRows
    Next i
End Sub

Public Sub zlChangeFormat(FormatID As Long)

    On Error GoTo errH
    
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim iRImageCount As Integer
    Dim blnHasMarkImage As Boolean
    Dim blnGetTable As Boolean
    Dim objFile As New Scripting.FileSystemObject
    Dim i As Integer
    
    '提示格式变换中图象框、标记图等的数量变化
    If FormatID = 0 Then     '标准格式，查 病历文件结构
        strSql = "Select Id As 表格Id From 病历文件结构" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngFileID)
    Else        '范文格式，查 病历范文内容
        strSql = "Select Id As 表格Id From 病历范文内容" & vbNewLine & _
        "       Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By 对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, FormatID)
    End If
    If rsTemp.RecordCount > 0 Then
        If rsTemp.RecordCount < dcmReportImage.Count - 1 Then
            If MsgBoxD(Me, "新格式中图象框数量少于当前格式，当前的部分图象框会被删除，是否更换格式？", vbOKCancel) = vbCancel Then
                Exit Sub
            Else
                '先删除多余的图象框
                For i = dcmReportImage.Count - 1 - rsTemp.RecordCount To 1 Step -1
                    Unload dcmReportImage(dcmReportImage.Count - 1)
                Next i
            End If
        End If
        
        '读取图象框中的标记图和报告图
        iRImageCount = 0
        pTableID = ""
        Do While Not rsTemp.EOF
            Set cTable = New cEPRTable
            If FormatID = 0 Then
                blnGetTable = cTable.GetTableFromDB(cprET_病历文件定义, mlngFileID, Val("" & rsTemp!表格ID))
            Else
                blnGetTable = cTable.GetTableFromDB(cprET_全文示范编辑, FormatID, Val("" & rsTemp!表格ID))
            End If
            If blnGetTable Then
                iRImageCount = iRImageCount + 1
                If iRImageCount > dcmReportImage.Count - 1 Then
                    '创建Viewer
                    Load dcmReportImage(iRImageCount)
                    dcmReportImage(iRImageCount).BorderStyle = 1
                    dcmReportImage(iRImageCount).Labels.AddNew
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LabelType = doLabelRectangle
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Left = 1
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Top = 1
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LineWidth = 2
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).ForeColour = vbWhite
                    dcmReportImage(iRImageCount).Visible = True
                End If
                mSelViewerIndex = iRImageCount
                
                '记录图像框的宽度和高度，该宽高比例用于后续对图像行列布局
                If cTable.ExtendTag <> "" Then
                    If Val(Split(cTable.ExtendTag, "|")(0)) = 0 Then
                        dcmReportImage(iRImageCount).Tag = cTable.Width & "|" & cTable.Height
                    Else
                        dcmReportImage(iRImageCount).Tag = (cTable.Width - Val(Split(cTable.ExtendTag, "|")(1)) - 30) & "|" & cTable.Height
                    End If
                Else
                    dcmReportImage(iRImageCount).Tag = cTable.Width & "|" & cTable.Height
                End If
                
                '更新标记图
                For i = 1 To cTable.Pictures.Count
                    strPicFile = App.Path & "\PACSPic" & i & ".JPG"
                    If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True

                    Set oPicture = cTable.Pictures(i).OrigPic
                    SavePicture oPicture, strPicFile
                    If objFile.FileExists(strPicFile) Then
                        '显示标记图
                        If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                            blnHasMarkImage = True
                            '先清除当前标记图，再更新
                            dcmMark.Images.Clear
                            dcmMark.Images.AddNew
                            dcmMark.Images(1).FileImport strPicFile, "BMP"
                            dcmMark.Images(1).Tag = cTable.Pictures(i).ID
                            '如果当前没有标记，则读取新格式中标记图的标记
                            If pobjMarks Is Nothing Then
                                Set pobjMarks = cTable.Pictures(i).PicMarks
                            End If
                            pMarkImageID = cTable.Pictures(i).ID
    
                            mdblMarkZoom = dcmMark.Images(1).SizeX / cTable.Pictures(i).Width * Screen.TwipsPerPixelX
                            '显示标注
                            If pobjMarks.Count > 0 Then
                                drawPicMarks dcmMark.Images(1), pobjMarks
                            End If
                        End If
                        '删除临时图像
                        Kill strPicFile
                    End If
                Next
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    If blnHasMarkImage = False Then
        '当前格式没有标记图，删除当前显示的标记图
        pMarkImageID = 0
        
        dcmMark.Images.Clear
        If Not (pobjMarks Is Nothing) Then
            For i = 1 To pobjMarks.Count
                pobjMarks.Remove 1
            Next i
        End If
    End If
    Call picReportImage_Resize

Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ucMiniImageViewer_OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
    If ucMiniImageViewer.CurImageCount > 0 And mSelViewerIndex <> 0 And dcmReportImage.Count > 1 Then
        '判断当前双击的操作动作
        If mintImageDblClick = 0 Then   '直接写入报告
            Dim dcmImage As DicomImage
            Set dcmImage = mSelMiniImg
            
            '调用将当前图形添加到图象框过程
            Call DcmAddImage(dcmImage, mSelViewerIndex)
            
        Else                            '先打开图片编辑窗口
            '先关闭大图窗口
            ReleaseCapture      '解锁鼠标
            frmShowImg.HideMe
            
            Call OpenImageProcessWind
        End If

    End If
    
    blnContinue = False
End Sub

Private Sub ucMiniImageViewer_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
If Button = 2 Then Call ShowPopupImage(True)
End Sub

Private Sub ucMiniImageViewer_OnSelChange(ByVal lngSelectedIndex As Long)
    Set mSelMiniImg = ucMiniImageViewer.SelectImage
End Sub

Private Sub vscrollMini_Change()
    Dim iImgIndex As Integer
    
    If dcmMiniImageC.Images.Count > 0 And (mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniImageC.Images.Count) Then
        iImgIndex = vscrollMini.value + mintCurImgIndex - dcmMiniImageC.CurrentIndex
        If iImgIndex <= 0 Then
            iImgIndex = 1
        ElseIf iImgIndex > dcmMiniImageC.Images.Count Then
            iImgIndex = dcmMiniImageC.Images.Count
        End If
        dcmMiniImageC.CurrentIndex = vscrollMini.value
        
        dcmMiniImageC.Images(mintCurImgIndex).BorderColour = vbWhite
        mintCurImgIndex = iImgIndex
        dcmMiniImageC.Images(mintCurImgIndex).BorderColour = vbRed
    End If

End Sub

Private Sub LoadImages()
'------------------------------------------------
'功能：加载报告图和缩略图
'参数：
'返回：无，直接加载图像，并修噶 blnLoadImages状态
'-----------------------------------------------
    '如果本次刷新没有加载图像，则加载图像
    If blnLoadImages = False Then
        '读取和显示当前可选报告图像
        Call LoadMiniImages
        '根据报告单格式，或者报告内容格式，读取标记图和报告图
        Call LoadReportImages
        '标记本次刷新已经加载图像
        blnLoadImages = True
    End If
End Sub
