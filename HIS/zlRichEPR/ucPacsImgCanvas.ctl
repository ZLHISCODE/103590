VERSION 5.00
Begin VB.UserControl ucPacsImgCanvas 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   ScaleHeight     =   3285
   ScaleWidth      =   4785
   Begin VB.PictureBox picResize 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   3330
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   3
      Top             =   1665
      Width           =   120
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2745
      ScaleHeight     =   435
      ScaleWidth      =   660
      TabIndex        =   2
      Top             =   675
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   1350
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1815
      ScaleWidth      =   30
      TabIndex        =   1
      Top             =   45
      Width           =   30
   End
   Begin VB.PictureBox picBuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2790
      ScaleHeight     =   435
      ScaleWidth      =   660
      TabIndex        =   0
      Top             =   1305
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Height          =   915
      Left            =   1575
      Top             =   1890
      Width           =   1410
   End
   Begin VB.Line linH 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   0
      X2              =   4875
      Y1              =   3015
      Y2              =   3030
   End
   Begin VB.Line LinV 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   3950
      X2              =   3950
      Y1              =   0
      Y2              =   3240
   End
   Begin VB.Menu mnuContextMenu 
      Caption         =   "右键菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "删除(&D)"
      End
      Begin VB.Menu mnuAdjust 
         Caption         =   "布局调整(&A)"
         Begin VB.Menu mnuMarkedPicOnLeft 
            Caption         =   "标记图居左(&L)"
         End
         Begin VB.Menu mnuMarkedPicOnRight 
            Caption         =   "标记图居右(&R)"
         End
         Begin VB.Menu mnuNoMarkdedPic 
            Caption         =   "无标记图(&N)"
         End
      End
   End
End
Attribute VB_Name = "ucPacsImgCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'局部变量
Private mblnMouseDown As Boolean, OldX As Long, OldY As Long    '鼠标移动需要的参数
Private lngMarkedPicPosition As Long        '0-无标记图；1-左边；2-右边；

Private mobjTable As cEPRTable              '传入的表格对象（只读取其中的Pictures）
Private mPictures As cEPRPictures           '存储所有加入的Pacs图片(>=0)
Private mPicturesPosition() As RECT         '存储每个图片的位置信息
Public mMarkedPicture As cEPRPicture        '标记图
Private mMarkedPicturePosition As RECT      '标记图外框位置
Private mMarkedPictureEditPosition As RECT  '标记图编辑位置

Private Space As Integer                    '边距
Private SelectedIndex As Long               '选中的图片

'全局事件
Public Event Resize(lngWidth As Long, lngHeight As Long)
Public Event SelectedMarkedPic(lLeft As Long, lTOp As Long, lWidth As Long, lHeight As Long) '选中标记图
Public Event SelectedPacsPic()


Private WithEvents cbsThis As CommandBars
Attribute cbsThis.VB_VarHelpID = -1
Private blnInited As Boolean
Private mBar报告图 As CommandBar
Private mBar布局 As CommandBarPopup
Private mfrmParent As frmMain
Private SumWidth As Long            '当前总宽度
Private dblZoomFactor As Double     '缩放比例
Private OldWidth As Long            '
Private lngRegionLeft As Long, lngRegionWidth As Long

Private mvarZoomFactor As Double         '缩放比例

Public Property Get zoomFactor() As Double
    zoomFactor = mvarZoomFactor
End Property

Public Property Let zoomFactor(ByVal vData As Double)
    mvarZoomFactor = vData
    Space = 60 * vData
    PropertyChanged "ZoomFactor"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Let MarkedPicPosition(vData As Long)
    lngMarkedPicPosition = vData
'    Call LayoutPictures
End Property

Public Property Get MarkedPicPosition() As Long
    MarkedPicPosition = lngMarkedPicPosition
End Property

Private Sub FixSplitPosition()
    With picSplit
        If lngMarkedPicPosition = 0 Or mMarkedPicture Is Nothing Then
            .Tag = ""
            .Visible = False
        ElseIf lngMarkedPicPosition = 1 Then
'            If .Tag = "R" Then
'                .Left = UserControl.ScaleWidth - .Left
'            Else
'                .Left = UserControl.ScaleWidth / 4
'            End If
            .Tag = "L"
            .Visible = True
        ElseIf lngMarkedPicPosition = 2 Then
'            If .Tag = "L" Then
'                .Left = UserControl.ScaleWidth - .Left
'            Else
'                .Left = UserControl.ScaleWidth * 3 / 4
'            End If
            .Tag = "R"
            .Visible = True
        End If
        .ZOrder 0
    End With
End Sub

Public Sub AddMarkedPicture(pic As StdPicture, lngPosition As Long)
    '添加标记图
    Dim lW As Long, lH As Long
    Set mMarkedPicture = New cEPRPicture
    Set mMarkedPicture.OrigPic = pic
    mMarkedPicture.PictureType = EPRMarkedPicture
    mMarkedPicture.OrigWidth = UserControl.ScaleX(pic.Width, vbHimetric, vbTwips)
    mMarkedPicture.OrigHeight = UserControl.ScaleX(pic.Height, vbHimetric, vbTwips)
    lW = IIf(mMarkedPicture.OrigWidth > 3000, 3000, mMarkedPicture.OrigWidth)
    lH = lW * pic.Height / pic.Width    '保持比例
    mMarkedPicture.Width = lW
    mMarkedPicture.Height = lH
    
    lngMarkedPicPosition = lngPosition
    Call SavePictures '保存到oTable中
    RaiseEvent SelectedPacsPic
    FixSplitPosition
    Call LayoutPictures
End Sub

Public Sub AddPacsPicture(pic As StdPicture, ByVal strUid As String, ByVal lngAdviceID As Long)
    '添加Pacs报告图片
    Dim lW As Long, lH As Long
    Dim lngKey As Long
'    Dim strPath As String, strF As String
'    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    
    '图片压缩保存
    picTmp.Cls
    picTmp.AutoRedraw = True
    
'    picTmp.Width = 4000         '此时BMP图片大小固定为 207K！！！
'    picTmp.Height = picTmp.Width * pic.Height / pic.Width
    picTmp.Width = UserControl.ScaleX(pic.Width, vbHimetric, vbTwips)
    picTmp.Height = UserControl.ScaleX(pic.Height, vbHimetric, vbTwips)
    
    picTmp.PaintPicture pic, 0, 0, picTmp.Width, picTmp.Height
    Set picTmp.Picture = picTmp.Image
    picTmp.Refresh
'    strF = strPath & "\TMP" & App.hInstance & "_" & CStr(Timer) & ".JPG" '保存临时文件
'    SavePicture picTMP.Image, strF
    
    lngKey = mPictures.Add
    Set mPictures("K" & lngKey).OrigPic = picTmp.Picture
    mPictures("K" & lngKey).PicName = strUid
    mPictures("K" & lngKey).AdviceID = lngAdviceID
    mPictures("K" & lngKey).PictureType = EPRInnerPicture
    mPictures("K" & lngKey).OrigWidth = UserControl.ScaleX(pic.Width, vbHimetric, vbTwips)
    mPictures("K" & lngKey).OrigHeight = UserControl.ScaleX(pic.Height, vbHimetric, vbTwips)
    lW = IIf(mPictures("K" & lngKey).OrigWidth > 4000, 4000, mPictures("K" & lngKey).OrigWidth)
    lH = lW * pic.Height / pic.Width    '保持比例
    mPictures("K" & lngKey).Width = lW
    mPictures("K" & lngKey).Height = lH
    
    Call SavePictures '保存到oTable中
    RaiseEvent SelectedPacsPic
    FixSplitPosition
    Call LayoutPictures
End Sub

'################################################################################################################
'   用途：  系统入口。
'################################################################################################################
Public Sub ShowMe(ByVal frmParent As frmMain, ByVal hWndParent As Long, ByVal cbsMain As CommandBars, ByVal objTable As cEPRTable, _
    ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long)
    
    shpBorder.Visible = False
    Set mfrmParent = frmParent
    
    Set cbsThis = cbsMain
'    If blnInited = False Then Call InitCommandBars
    blnInited = True
'    DockingRightOf mfrmParent.Bar表格, mBar报告图
'    mBar报告图.Visible = True
'    mBar布局.Visible = True
    UserControl.KeyPreview = True
'    SumWidth = IIf(objTable.Pictures.Count = 0, 8800, objTable.Width)
    
    ReadPicturesFromTable objTable, False
    '在父容器中显示控件
    SetParent UserControl.hwnd, hWndParent
    UserControl.Extender.Left = lngLeft
    UserControl.Extender.Top = lngTop
    UserControl.Extender.Width = lngWidth
    UserControl.Extender.Height = lngHeight
    UserControl.BackColor = vbWhite         ' &H8000000F
    UserControl.BorderStyle = 0
    UserControl.Cls
    UserControl.AutoRedraw = True
    UserControl.Extender.Visible = True
End Sub

Private Sub InitCommandBars()
    Dim cbpPopup As CommandBarPopup     '临时对象
    Dim cbpPopupSub As CommandBarPopup  '临时对象
    Dim objControl As CommandBarControl                 '工具栏控件
    Dim objCustControl As CommandBarControlCustom       '自定义控件
    Dim Combo As CommandBarComboBox     '工具栏下拉框控件
    
    Set mBar报告图 = cbsThis.Add("报告图", xtpBarTop)
    mBar报告图.EnableDocking xtpFlagHideWrap
    mBar报告图.ModifyStyle XTP_CBRS_GRIPPER, 0
    With mBar报告图.Controls
        Set objControl = .Add(xtpControlButton, ID_PACS_DeletePacsImg, "删除报告图片(&C)")
        Set mBar布局 = .Add(xtpControlButtonPopup, ID_PACS_Layout, "布局调整")
        mBar布局.BeginGroup = True
        mBar布局.Style = xtpButtonIconAndCaption
        mBar布局.CommandBar.Controls.Add xtpControlButton, ID_PACS_Left, "标记图在左边"
        mBar布局.CommandBar.Controls.Add xtpControlButton, ID_PACS_Right, "标记图在右边"
        mBar布局.CommandBar.Controls.Add xtpControlButton, ID_PACS_None, "无标记图"
    End With
    DockingRightOf mBar报告图, mfrmParent.CommBar(ID_BAR_FORMAT)
End Sub

'################################################################################################################
'## 功能：  将工具条A放置到工具条B的同一行
'##
'## 参数：  BarToDock   ：加入的工具栏
'##         BarOnLeft   ：位于左边的工具条
'################################################################################################################
Private Sub DockingRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position
End Sub

Public Sub SavePictures()
    '保存所有图片到表格中
    If Not mMarkedPicture Is Nothing Then
        Set mobjTable.Pictures = New cEPRPictures
        If Not mMarkedPicture Is Nothing Then mobjTable.Pictures.AddExistNode mMarkedPicture
        Dim i As Long
        For i = 1 To mPictures.Count
            mobjTable.Pictures.AddExistNode mPictures(i)
        Next
    Else
        Set mobjTable.Pictures = mPictures      '保存图片到表格中
    End If
    mobjTable.Width = UserControl.Width
    mobjTable.Height = UserControl.Height
    mobjTable.ExtendTag = lngMarkedPicPosition & "|" & picSplit.Left & "|" & CStr(IIf(picSplit.Visible, 1, 0))
End Sub

Public Sub CloseMe()
    Call SavePictures
    UserControl.Extender.Visible = False    '隐藏控件
    UserControl.Extender.Tag = ""
'    mBar报告图.Visible = False
'    mBar布局.Visible = False
End Sub

Public Sub ReadPicturesFromTable(objTable As cEPRTable, Optional RaiseResizeEvent As Boolean = True)
    '从一个Table中读取所有图片
'    SumWidth = IIf(objTable.Pictures.Count = 0 Or objTable.Width = 0, 8800, objTable.Width)
    UserControl.Width = IIf(objTable.Width = 0, 6000, objTable.Width)
    UserControl.Height = IIf(objTable.Height = 0, 4000, objTable.Height)
    If objTable.Width = 0 Then objTable.Width = 6000
    If objTable.Height = 0 Then objTable.Height = 4000
'
    On Error Resume Next
    Dim i As Long, lKey As Long, T As Variant
    T = Split(objTable.ExtendTag, "|")
    lngMarkedPicPosition = Val(T(0))
    picSplit.Visible = (T(2) = 1)
    picSplit.Left = Val(T(1))
    
    Set mobjTable = objTable                '保存表格的引用
    
    Set mPictures = New cEPRPictures        '提取表格中的图片
    If objTable.Pictures.Count = 0 Then
        Set mMarkedPicture = Nothing
    Else
        If objTable.Pictures(1).PictureType <> EPRInnerPicture Then
            '第一个是标记图
            If lngMarkedPicPosition = 0 Then lngMarkedPicPosition = 1
            Set mMarkedPicture = objTable.Pictures(1)
            For i = 2 To objTable.Pictures.Count
                mPictures.AddExistNode objTable.Pictures(i).Clone
            Next
        Else
            Set mMarkedPicture = Nothing
            Set mPictures = objTable.Pictures.Clone
        End If
    End If
    If SelectedIndex > mPictures.Count Then SelectedIndex = 0
    FixSplitPosition
    Call LayoutPictures(RaiseResizeEvent)    '然后给每个图片定位
End Sub

'-----------------------------------------
'以下为公用方法
'-----------------------------------------
Private Sub ResizeRegion(ByVal PicCount As Integer, _
    ByVal RegionWidth As Long, ByVal RegionHeight As Long, _
    Rows As Integer, Cols As Integer)
    '-----------------------------------------------------------
    '功能： 根据需要显示的图像数量和显示区域，计算可显示图像的行列数。
    '参数： PicCount-图像数量
    '       RegionWidth,RegionHeight-区域宽度高度
    '       Rows,Cols-返回自动排列的行列数
    '-----------------------------------------------------------
    Dim intRows As Integer, intCols As Integer
    If RegionHeight = 0 Or RegionWidth = 0 Then
        Rows = 1
        Cols = 1
        Exit Sub
    Else
        intRows = CInt(Sqr(PicCount * RegionHeight / RegionWidth))
        intCols = CInt(Sqr(PicCount * RegionWidth / RegionHeight))
    End If
        
    '下面4句是避免在只有1幅标记图和1幅报告图时报错
    intRows = IIf(intRows > PicCount, PicCount, intRows)
    intCols = IIf(intCols > PicCount, PicCount, intCols)
    intRows = IIf(intRows <= 0, 1, intRows)
    intCols = IIf(intCols <= 0, 1, intCols)
    
    Do While intRows * intCols < PicCount
        If RegionWidth / RegionHeight > 1 Then
            intCols = intCols + 1
        Else
            intRows = intRows + 1
        End If
    Loop
    Rows = intRows: Cols = intCols
End Sub

Private Function DrawPicture(ByVal Picture As StdPicture, _
    ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, _
    Optional BorderLine As Boolean = True, Optional BoderColor As Long = 12845056, Optional lngNumber As Integer = 0) As RECT
    '-----------------------------------------------------------
    '功能： 在指定区域按图像本身横纵向比例绘制图像
    '参数： Picture-要绘制的图像
    '       Left,Top,Width,Height-区域起点和大小
    '       BorderLine-是否需要边框
    '       BoderColor-边框的颜色
    '-----------------------------------------------------------
    Dim PicWidth As Long, PicHeight As Long     '图片尺寸，先取得原始尺寸，计算处理为打印尺寸
    Dim lLeft As Long, lTOp As Long
    Dim ActualRect As RECT  '实际位置
    Dim clsDIB As New clsDIB
    Dim clsDIBTemp As New clsDIB
    Dim sngScale As Single
    Dim W As Long, H As Long
    
    lLeft = Left
    lTOp = Top
    
    PicWidth = Picture.Width
    PicHeight = Picture.Height
    
    If BorderLine Then
        UserControl.Line (Left, Top)-(Left + Width - Space, Top), BoderColor
        UserControl.Line (Left, Top)-(Left, Top + Height - Space), BoderColor
        UserControl.Line (Left, Top + Height - Space)-(Left + Width - Space, Top + Height - Space), BoderColor
        UserControl.Line (Left + Width - Space, Top)-(Left + Width - Space, Top + Height - Space), BoderColor
    End If
   
    Width = Width - Space * 3
    Height = Height - Space * 3
    
    If Width / PicWidth < Height / PicHeight Then
        PicHeight = Int(PicHeight * (Width / PicWidth))
        PicWidth = Width
    Else
        PicWidth = Int(PicWidth * (Height / PicHeight))
        PicHeight = Height
    End If
    Left = Left + Int((Width - PicWidth) / 2)
    Top = Top + Int((Height - PicHeight) / 2)
    
    If lngNumber = 0 Then
        '标记图
        '缩放标记
        dblZoomFactor = PicWidth / mMarkedPicture.Width
        Set mMarkedPicture.PicMarks = ScalePicMarks(mMarkedPicture.PicMarks, dblZoomFactor)
        UserControl.PaintPicture mMarkedPicture.DrawFinalPic, Left + Space, Top + Space, PicWidth, PicHeight
'        '恢复标记
'        Set mMarkedPicture.PicMarks = ScalePicMarks(mMarkedPicture.PicMarks, 1 / dblZoomFactor)
        mMarkedPicture.Width = PicWidth
        mMarkedPicture.Height = PicHeight
    Else
    
        Call SavePicture(Picture, App.Path & "\dibtmp.tmp")
        If clsDIB.DIBLoadMap(App.Path & "\dibtmp.tmp", True, 24) Then
            If clsDIB.DataPtr <> 0 Then
                                
                If Width / (clsDIB.Width * Screen.TwipsPerPixelX) > Height / (clsDIB.Height * Screen.TwipsPerPixelX) Then
                    sngScale = Height / (clsDIB.Height * Screen.TwipsPerPixelX)
                Else
                    sngScale = Width / (clsDIB.Width * Screen.TwipsPerPixelX)
                End If
                
                W = clsDIB.Width * sngScale
                H = clsDIB.Height * sngScale
                If W < 1 Then W = 1
                If H < 1 Then H = 1
                
                If clsDIBTemp.DIBScale(clsDIB, W, H) Then
                    clsDIBTemp.PutTo UserControl.hDC, (Left + Space) / Screen.TwipsPerPixelX, (Top + Space) / Screen.TwipsPerPixelY
                End If
            End If
        End If
        Kill App.Path & "\dibtmp.tmp"
        
'        UserControl.PaintPicture Picture, Left + Space, Top + Space, PicWidth, PicHeight
    End If
    
    ActualRect.Left = Left + Space
    ActualRect.Top = Top + Space
    ActualRect.Right = PicWidth
    ActualRect.Bottom = PicHeight
    
    '报告图片编号
    If lngNumber > 0 Then
        UserControl.FontName = "Arial"
        UserControl.FontSize = 9 * zoomFactor
        Dim LL As Long, lT As Long
        LL = lLeft + (Space + 30) * zoomFactor
        lT = lTOp + (Space + 30) * zoomFactor
        DrawText LL, lT, CStr(lngNumber), vbWhite
        LL = lLeft + (Space + 15) * zoomFactor
        lT = lTOp + (Space + 15) * zoomFactor
        DrawText LL, lT, CStr(lngNumber), vbBlack
    End If
    
    Set clsDIB = Nothing
    Set clsDIBTemp = Nothing
    DrawPicture = ActualRect
End Function

Public Sub DrawText(ByVal x As Single, ByVal y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0)
    '在(X,Y)处输出Text文本
    Dim lngSaveForeColor As Long
    
    With UserControl
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        .CurrentX = x
        .CurrentY = y
        .FontTransparent = True
        UserControl.Print Text
        .ForeColor = lngSaveForeColor
    End With
End Sub

Public Sub LayoutPictures(Optional RaiseResizeEvent As Boolean = True, Optional ByVal BorderLine As Boolean = True)
    '-----------------------------------------------------------
    '功能： 重新绘制所有图像
    '-----------------------------------------------------------
    Dim introw As Integer, intCol As Integer, intIndex As Integer
    Dim intRows As Integer, intCols As Integer
    Dim lngPerWidth As Long, lngPerHeight As Long
    Dim lngPicLeft As Long, lngPicTop As Long
        
    UserControl.Cls
    Set UserControl.Picture = Nothing
    
    '绘制标记图
    If lngMarkedPicPosition <> 0 And Not mMarkedPicture Is Nothing Then
        If lngMarkedPicPosition = 1 Then
            lngPicLeft = 0
            lngRegionWidth = picSplit.Left
        Else
            lngPicLeft = picSplit.Left + picSplit.Width
            lngRegionWidth = UserControl.ScaleWidth - (picSplit.Left + picSplit.Width)
        End If
        lngPicTop = 0
        mMarkedPicturePosition.Left = lngPicLeft
        mMarkedPicturePosition.Top = lngPicTop
        mMarkedPicturePosition.Right = lngRegionWidth
        mMarkedPicturePosition.Bottom = UserControl.ScaleHeight
        mMarkedPictureEditPosition = DrawPicture(mMarkedPicture.OrigPic, lngPicLeft, lngPicTop, lngRegionWidth, UserControl.ScaleHeight, BorderLine)
    End If
    
    '计算PACS图片的区域
    If mMarkedPicture Is Nothing Then
        lngRegionLeft = 0
        lngRegionWidth = UserControl.ScaleWidth
    Else
        Select Case lngMarkedPicPosition
        Case 1
            lngRegionLeft = picSplit.Left + picSplit.Width
            lngRegionWidth = UserControl.ScaleWidth - (picSplit.Left + picSplit.Width)
        Case 2
            lngRegionLeft = 0
            lngRegionWidth = picSplit.Left
        Case Else
            lngRegionLeft = 0
            lngRegionWidth = UserControl.ScaleWidth
        End Select
    End If
    
    '计算自动排列的行列数，并循环重新绘制图像
    Call ResizeRegion(mPictures.Count, lngRegionWidth, UserControl.ScaleHeight, intRows, intCols)
    lngPerWidth = Fix(lngRegionWidth / intCols)
    lngPerHeight = Fix(UserControl.ScaleHeight / intRows)
    ReDim mPicturesPosition(1 To mPictures.Count + 1) As RECT
    
    For introw = 0 To intRows - 1
        For intCol = 0 To intCols - 1
            intIndex = introw * intCols + intCol + 1
            If intIndex > mPictures.Count Then Exit For
            lngPicLeft = lngRegionLeft + intCol * lngPerWidth
            lngPicTop = introw * lngPerHeight
            
            mPicturesPosition(intIndex).Left = lngPicLeft
            mPicturesPosition(intIndex).Top = lngPicTop
            mPicturesPosition(intIndex).Right = lngPerWidth
            mPicturesPosition(intIndex).Bottom = lngPerHeight

            Call DrawPicture(mPictures(intIndex).OrigPic, lngPicLeft, lngPicTop, lngPerWidth, lngPerHeight, BorderLine, , intIndex)
        Next
    Next
    Set UserControl.Picture = UserControl.Image
    DrawPicBoder
End Sub

Private Sub DrawPic(objDest As PictureBox, pic As cEPRPicture, lWidth As Long, lHeight As Long, Optional lngNumber As Long = 0)
    '在指定图片上绘制规定大小的图片（缩放）
'    On Error Resume Next
    
    objDest.BorderStyle = 0
    objDest.AutoRedraw = True
    objDest.Width = lWidth
    objDest.Height = lHeight
    objDest.Cls
    If objDest.Name = "picMarkedPic" Then
        objDest.PaintPicture pic.DrawFinalPic, 0, 0, lWidth, lHeight
    Else
        objDest.PaintPicture pic.DrawFinalPic, Space + 15, Space + 15, lWidth - Space + 15, lHeight - Space + 15
    End If
    
    '报告图片编号
    If lngNumber > 0 Then
'        objDest.FontName = "Arial"
        DrawText 145, 125, CStr(lngNumber), vbBlack
        DrawText 130, 110, CStr(lngNumber), vbWhite
    End If
    
    Set objDest.Picture = objDest.Image
    objDest.Refresh
End Sub

Public Function FinalPic(Optional ByVal BorderLine As Boolean = True) As StdPicture
    '返回最终图片供显示、打印
'    On Error Resume Next
'    LayoutPictures
    picBuff.Width = UserControl.Width
    picBuff.Height = UserControl.Height
    picBuff.BorderStyle = 0
    picBuff.AutoRedraw = True
    picBuff.BackColor = vbWhite ' &H8000000F
    picBuff.Cls
    
    If mPictures.Count > 0 Or Not mMarkedPicture Is Nothing Then
        LayoutPictures , BorderLine
    Else
        '在没有图片时的默认结果
'        UserControl.Width = 6000
'        UserControl.Height = 4000
        picBuff.Width = UserControl.Width
        picBuff.Height = UserControl.Height
        DrawText 100, 60, IIf(BorderLine, "请加入报告图片...", "")
        Set UserControl.Picture = UserControl.Image
    End If
    Set picBuff.Picture = UserControl.Picture
    
    '绘制边框
    Dim hPen As Long
    Dim hPenOld As Long
    Dim m_hDC As Long
    If mPictures.Count = 0 And mMarkedPicture Is Nothing Then
        m_hDC = picBuff.hDC
    
        hPen = CreatePen(PS_SOLID, 1, IIf(BorderLine, vbBlack, vbWhite))    '设置边框颜色画笔
        hPenOld = SelectObject(m_hDC, hPen)         '选入画笔，保存旧画笔
        Rectangle m_hDC, 0, 0, picBuff.Width / 15, picBuff.Height / 15
        SelectObject m_hDC, hPenOld
        DeleteObject hPen
        hPen = 0
    
        Set picBuff.Picture = picBuff.Image
    End If
    picBuff.Refresh
    Set FinalPic = picBuff.Image
End Function

Private Sub ShowRightMenu()
    '显示右键菜单
    Dim Popup As CommandBar
    Dim cbpPopup As CommandBarPopup
    Dim Control As CommandBarControl
    
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_PACS_DeletePacsImg, "删除报告图片(&C)")
        Set Control = .Add(xtpControlButton, ID_PACS_Left, "标记图在左边(&L)")
        Set Control = .Add(xtpControlButton, ID_PACS_Right, "标记图在右边(&R)")
        Set Control = .Add(xtpControlButton, ID_PACS_None, "无标记图(&N)")
        Popup.ShowPopup
    End With
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_PACS_DeletePacsImg
        If (SelectedIndex > 0) Then
            mPictures.Remove "K" & mPictures(SelectedIndex).Key
            Dim i As Long
            SelectedIndex = IIf(mPictures.Count = 0, 0, IIf(SelectedIndex > mPictures.Count, mPictures.Count, SelectedIndex))
            UserControl.Cls
        End If
        RaiseEvent SelectedPacsPic
        LayoutPictures True
    Case ID_PACS_Left
        lngMarkedPicPosition = 1
        FixSplitPosition
        RaiseEvent SelectedPacsPic
        LayoutPictures True
    Case ID_PACS_Right
        lngMarkedPicPosition = 2
        FixSplitPosition
        RaiseEvent SelectedPacsPic
        LayoutPictures True
    Case ID_PACS_None
        Set mMarkedPicture = Nothing
        lngMarkedPicPosition = 0
        FixSplitPosition
        UserControl.Cls
        RaiseEvent SelectedPacsPic
        LayoutPictures True
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_PACS_DeletePacsImg
        If SelectedIndex = 0 Then
            Control.Enabled = False
        Else
            Control.Enabled = (SelectedIndex > 0)
        End If
    Case ID_PACS_Left
        Control.Checked = (lngMarkedPicPosition = 1)
    Case ID_PACS_Right
        Control.Checked = (lngMarkedPicPosition = 2)
    Case ID_PACS_None
        Control.Checked = (lngMarkedPicPosition = 0)
    End Select
End Sub

Private Sub DrawPicBoder()
    '选中当前图片，绘制边框效果
'    shpBorder.BorderWidth = 2
    If SelectedIndex > 0 Then
        shpBorder.Move mPicturesPosition(SelectedIndex).Left + 30, mPicturesPosition(SelectedIndex).Top + 30, _
            mPicturesPosition(SelectedIndex).Right - 105, mPicturesPosition(SelectedIndex).Bottom - 105
        shpBorder.Visible = True
        shpBorder.ZOrder 0
    ElseIf SelectedIndex = -1 Then
        shpBorder.Move mMarkedPicturePosition.Left + 30, mMarkedPicturePosition.Top + 30, _
            mMarkedPicturePosition.Right - 105, mMarkedPicturePosition.Bottom - 105
        shpBorder.Visible = True
        shpBorder.ZOrder 0
    Else
        shpBorder.Visible = False
    End If
End Sub

Private Sub UserControl_Initialize()
    Space = 60
    zoomFactor = 1#
    Set mPictures = New cEPRPictures
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (x > mMarkedPicturePosition.Left And x < (mMarkedPicturePosition.Left + mMarkedPicturePosition.Right) And _
        y > mMarkedPicturePosition.Top And y < (mMarkedPicturePosition.Top + mMarkedPicturePosition.Bottom) And Not mMarkedPicture Is Nothing) Then
        SelectedIndex = -1
    Else
        SelectedIndex = 0
        Dim i As Long
        For i = 1 To UBound(mPicturesPosition)
            If (x > mPicturesPosition(i).Left And x < (mPicturesPosition(i).Left + mPicturesPosition(i).Right) And _
                y > mPicturesPosition(i).Top And y < (mPicturesPosition(i).Top + mPicturesPosition(i).Bottom)) Then
                SelectedIndex = i
                RaiseEvent SelectedPacsPic
                Exit For
            End If
        Next
    End If
    DrawPicBoder
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        ShowRightMenu
    Else
        If SelectedIndex = -1 Then
            '编辑标记图
            RaiseEvent SelectedMarkedPic(mMarkedPictureEditPosition.Left, mMarkedPictureEditPosition.Top, mMarkedPictureEditPosition.Right, mMarkedPictureEditPosition.Bottom)
        End If
    End If
End Sub

Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent SelectedPacsPic
    With UserControl.linH
        .X1 = 0: .Y1 = UserControl.ScaleHeight - Screen.TwipsPerPixelY
        .X2 = UserControl.ScaleWidth - Screen.TwipsPerPixelX: .Y2 = .Y1
        .Visible = True: .Tag = .Y1: .ZOrder 0
    End With
    With UserControl.LinV
        .X1 = UserControl.ScaleWidth - Screen.TwipsPerPixelX: .Y1 = 0
        .X2 = .X1: .Y2 = UserControl.ScaleHeight - Screen.TwipsPerPixelY
        .Visible = True: .Tag = .X1: .ZOrder 0
    End With
End Sub

Private Sub picResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If UserControl.linH.Visible = False Or UserControl.LinV.Visible = False Then Exit Sub
    
    If Val(UserControl.linH.Tag) + y < 2000 Then y = 2000 - Val(UserControl.linH.Tag)
    If Val(UserControl.linH.Tag) + y > 10500 Then y = 10500 - Val(UserControl.linH.Tag)
    If Val(UserControl.LinV.Tag) + x < 3000 Then x = 3000 - Val(UserControl.LinV.Tag)
    If Val(UserControl.LinV.Tag) + x > 10500 Then x = 10500 - Val(UserControl.LinV.Tag)
    With UserControl.linH
        .X1 = 0: .Y1 = Val(.Tag) + y
        .X2 = Val(UserControl.LinV.Tag) + x: .Y2 = .Y1
    End With
    With UserControl.LinV
        .X1 = Val(.Tag) + x: .Y1 = 0
        .X2 = .X1: .Y2 = Val(UserControl.linH.Tag) + y
    End With
End Sub

Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngWidth As Long, lngHeight As Long
    
    UserControl.linH.Visible = False: UserControl.LinV.Visible = False
    With UserControl
        lngWidth = .Width + UserControl.LinV.X1 - Val(UserControl.LinV.Tag)
        lngHeight = .Height + UserControl.linH.Y1 - Val(UserControl.linH.Tag)
        picSplit.Left = picSplit.Left * lngWidth / .Width
        .Width = lngWidth
        .Height = lngHeight
    End With
    
    LayoutPictures
    RaiseEvent Resize(UserControl.Width, UserControl.Height)
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent SelectedPacsPic
    If Not mMarkedPicture Is Nothing Then OldWidth = mMarkedPicture.Width
    With LinV
        .X1 = picSplit.Left: .Y1 = 0
        .X2 = .X1: .Y2 = UserControl.ScaleHeight
        .Visible = True: .Tag = .X1: .ZOrder 0
    End With
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LinV.Visible = False Then Exit Sub
    
    If Val(LinV.Tag) + x < 1000 Then x = 1000 - Val(LinV.Tag)
    If Val(LinV.Tag) + x > UserControl.ScaleWidth - 1000 Then x = UserControl.ScaleWidth - 1000 - Val(LinV.Tag)
    With LinV
        .X1 = Val(.Tag) + x
        .X2 = .X1
    End With
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    LinV.Visible = False
    picSplit.Left = LinV.X1
    
    Call UserControl_Resize
    Call LayoutPictures
    RaiseEvent Resize(UserControl.Width, UserControl.Height)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim lKey As Long, i As Long
'    Select Case KeyCode
'    Case vbKeyLeft, vbKeyUp
'        If picSelected Is Nothing Then
'            If mPictures.Count > 0 Then Set picSelected = picImgs(0)
'        Else
'            For i = 1 To picImgs.Count - 1
'                If picImgs(i) Is picSelected Then
'                    picSelected.Cls
'                    Set picSelected = picImgs(i - 1)
'                    DrawPicBoder picSelected
'                    Exit For
'                End If
'            Next
'        End If
'    Case vbKeyRight, vbKeyDown
'        If picSelected Is Nothing Then
'            If mPictures.Count > 0 Then Set picSelected = picImgs(picImgs.UBound)
'        Else
'            For i = 0 To picImgs.UBound - 1
'                If picImgs(i) Is picSelected Then
'                    picSelected.Cls
'                    Set picSelected = picImgs(i + 1)
'                    DrawPicBoder picSelected
'                    Exit For
'                End If
'            Next
'        End If
'    End Select
End Sub

Private Sub UserControl_Resize()
    picSplit.Top = 0: picSplit.Height = UserControl.Height
    picResize.Left = UserControl.ScaleWidth - picResize.Width
    picResize.Top = UserControl.ScaleHeight - picResize.Height
'    Call LayoutPictures
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    Set mobjTable = Nothing
    Set mPictures = Nothing
    Set mMarkedPicture = Nothing
    Set cbsThis = Nothing
    Set mBar报告图 = Nothing
    Set mBar布局 = Nothing
    Set mfrmParent = Nothing
End Sub
