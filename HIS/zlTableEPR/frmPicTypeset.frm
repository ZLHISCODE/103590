VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPicTypeset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报告图排版"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "frmPicTypeset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picBak 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9600
      Left            =   3270
      ScaleHeight     =   9600
      ScaleWidth      =   5280
      TabIndex        =   0
      Top             =   -15
      Width           =   5280
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   -345
         ScaleHeight     =   525
         ScaleWidth      =   615
         TabIndex        =   12
         Top             =   8715
         Width           =   615
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3885
         TabIndex        =   1
         Top             =   9165
         Width           =   1200
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "应用(&A)"
         Height          =   350
         Left            =   2700
         TabIndex        =   2
         Top             =   9165
         Width           =   1200
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "向右(&R)"
         Height          =   350
         Index           =   3
         Left            =   3885
         TabIndex        =   5
         Top             =   8730
         Width           =   1200
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "向左(&L)"
         Height          =   350
         Index           =   2
         Left            =   2700
         TabIndex        =   6
         Top             =   8730
         Width           =   1200
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全清(&B)"
         Height          =   350
         Left            =   1515
         TabIndex        =   11
         Top             =   9165
         Width           =   1200
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "清除(&K)"
         Height          =   350
         Left            =   330
         TabIndex        =   10
         Top             =   9165
         Width           =   1200
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "向下(&D)"
         Height          =   350
         Index           =   1
         Left            =   1515
         TabIndex        =   7
         Top             =   8730
         Width           =   1200
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "向上(&U)"
         Height          =   350
         Index           =   0
         Left            =   330
         TabIndex        =   8
         Top             =   8730
         Width           =   1200
      End
      Begin VB.Frame fraTypeset 
         Caption         =   "图象排版"
         Height          =   8565
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   5220
         Begin VSFlex8Ctl.VSFlexGrid vsPic 
            Height          =   8235
            Left            =   45
            TabIndex        =   4
            Top             =   240
            Width           =   5130
            _cx             =   9049
            _cy             =   14526
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   -2147483640
            BackColorBkg    =   16761024
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   -1  'True
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
            Begin VB.Shape shpBorder 
               BorderColor     =   &H00FF0000&
               Height          =   255
               Left            =   1245
               Top             =   1605
               Width           =   270
            End
         End
      End
   End
   Begin VB.PictureBox picResult 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   315
      ScaleHeight     =   2310
      ScaleWidth      =   2295
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   2325
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   30
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPicTypeset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mPacsimg As frmPACSImg
Attribute mPacsimg.VB_VarHelpID = -1
Private mlngAdviceID As Long, mlngWidth As Long, mlngHeight As Long, mselKey As String, mParent As Object
Private mlngModule As Long

Public Sub ShowTypeset(ByVal fParent As Object, ByVal selKey As String, ByVal lngAdviceID As Long, ByVal lngWidth As Long, _
    ByVal lngHeight As Long, ByVal SImg As StdPicture, ByVal AddImg As StdPicture, ByVal lngModule As Long)
'功能：完成传入图片的排版
'参数：lngWidth单元格原始宽度,lngHeight单元格原始高度
'返回：排版后的图片
    If Me.Visible Then Exit Sub
    picTmp(0).Visible = False
    mselKey = selKey
    mlngAdviceID = lngAdviceID
    mlngModule = lngModule
    mlngWidth = lngWidth
    mlngHeight = lngHeight
    Set mParent = fParent
    Me.Show 0, fParent
    mPacsimg.zlRefresh mlngAdviceID, mlngModule
    zlAddPic SImg
    zlAddPic AddImg
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
Dim l As Long
    For l = 1 To picTmp.UBound
        Unload picTmp(l)
    Next
    
    vsPic.Rows = 1: vsPic.Cols = 1: vsPic.RowHeight(0) = vsPic.Height: vsPic.ColWidth(0) = vsPic.Width
    Set vsPic.Cell(flexcpPicture, 0, 0) = Nothing
    Call vsPic_SelChange
End Sub

Private Sub cmdDel_Click()
'picTmp第0维没有使用
    If Not shpBorder.Visible Then Exit Sub
    Dim l As Long, i As Long
    '从左数起第N个,=(当前行-1)*列数+当前列
    If picTmp.UBound = 0 Then Exit Sub
    l = vsPic.Row * vsPic.Cols + vsPic.Col
    If vsPic.Cell(flexcpPicture, vsPic.Row, vsPic.Col) Is Nothing Then Exit Sub
    
    For i = l + 1 To picTmp.UBound - 1
        If i > picTmp.UBound Then Exit For '计算错误
        Set picTmp(i).Picture = picTmp(i + 1).Picture
    Next
    Unload picTmp(picTmp.UBound)
    Call FillPic
End Sub
Private Sub cmdMove_Click(Index As Integer)
'用picTmp第0维作交换
Dim lS As Long, lD As Long
    With vsPic
        lS = vsPic.Row * vsPic.Cols + vsPic.Col '源图片处于Pictmp第几维
        lS = lS + 1
        Select Case Index
            Case 0 '上
                If .Row = 0 Then Call MsgBox("不能向上调整", vbInformation, gstrSysName): Exit Sub
                lD = (vsPic.Row - 1) * vsPic.Cols + vsPic.Col '目标图片处于第N维
            Case 1 '下
                If .Row = .Rows - 1 Then Call MsgBox("不能向下调整", vbInformation, gstrSysName): Exit Sub
                lD = (vsPic.Row + 1) * vsPic.Cols + vsPic.Col '目标图片处于第N维
            Case 2 '左
                If .Col = 0 Then Call MsgBox("不能向左调整", vbInformation, gstrSysName): Exit Sub
                lD = vsPic.Row * vsPic.Cols + (vsPic.Col - 1)
            Case 3 '右
                If .Col = .Cols - 1 Then Call MsgBox("不能向右调整", vbInformation, gstrSysName): Exit Sub
                lD = vsPic.Row * vsPic.Cols + (vsPic.Col + 1)
        End Select
        lD = lD + 1
        If lD > picTmp.UBound Then Call MsgBox("不能调整到指定位置", vbInformation, gstrSysName): Exit Sub
        Set picTmp(0).Picture = Nothing
        Set picTmp(0).Picture = picTmp(lD).Picture
        Set picTmp(lD).Picture = picTmp(lS).Picture
        Set picTmp(lS).Picture = picTmp(0).Picture
        Set picTmp(0).Picture = Nothing
        Call FillPic
    End With
End Sub

Private Sub cmdOK_Click()
    Call MakeResultPic
    If picResult.Picture.Handle = 0 Then Exit Sub
    Set mParent.Document.Pictures("K" & mParent.Document.Cells(mselKey).PictureKey).OrigPic = picResult.Image
    mParent.PaintPictureOnTable mselKey
    Unload Me
End Sub
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = mPacsimg.hWnd
        Case 2
            Item.Handle = picBak.hWnd
    End Select
End Sub

Private Sub Form_Load()
Dim paneList As Pane, paneApply As Pane
    With Me.dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
    Set paneList = dkpMain.CreatePane(1, 250, 0, DockLeftOf)
    paneList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set paneApply = dkpMain.CreatePane(2, 400, 0, DockRightOf)
    paneApply.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set mPacsimg = New frmPACSImg
    vsPic.Cols = 1: vsPic.Rows = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mPacsimg
    Set mPacsimg = Nothing
    Set mParent = Nothing
End Sub
Private Sub zlAddPic(ByVal AddPic As StdPicture)
    shpBorder.Visible = False
    Load picTmp(picTmp.UBound + 1)
    Set picTmp(picTmp.UBound).Picture = AddPic
    Call FillPic
End Sub
Private Sub FillPic()
Dim lRows As Integer, lCols As Integer, R As Long, C As Long, l As Long
Dim lWidth As Long, lHeight As Long, loRow As Long, loCol As Long

    loRow = vsPic.Row: loCol = vsPic.Col
    'picTmp第0维不使用
    Call ResizeRegion(picTmp.Count - 1, mlngWidth, mlngHeight, lRows, lCols)
    '用表格宽度计算每列宽度，用表格高度计算每行高度
    With vsPic
        .Clear
        .Rows = lRows
        .Cols = lCols
        lWidth = (vsPic.Width - 50) / .Cols
        lHeight = (vsPic.Height - 50) / .Rows
        For C = 0 To .Cols - 1
            .ColWidth(C) = lWidth
        Next
        For R = 0 To .Rows - 1
            .RowHeight(R) = lHeight
        Next
        
        l = 1
        For R = 0 To .Rows - 1
            For C = 0 To .Cols - 1
                Set .Cell(flexcpPicture, R, C) = Nothing
                If picTmp.UBound = 0 Then Exit Sub
                picTmp(l).Width = lWidth
                picTmp(l).Height = lHeight
                picTmp(l).Cls
                picTmp(l).PaintPicture picTmp(l).Picture, 0, 0, lWidth, lHeight
                
                Set .Cell(flexcpPicture, R, C) = picTmp(l).Image
                .Cell(flexcpPictureAlignment, R, C) = flexAlignCenterCenter
                l = l + 1
                If l > picTmp.UBound Then
                    If loRow > vsPic.Rows - 1 Then loRow = vsPic.Rows - 1
                    If loCol > vsPic.Cols - 1 Then loCol = vsPic.Cols - 1
                    vsPic.Row = loRow: vsPic.Col = loCol
                    Call vsPic_SelChange
                    Exit Sub
                End If
            Next
        Next
    End With
End Sub
Private Sub MakeResultPic()
Dim lWidth As Long, lHeight As Long '预定绘图区域宽高
Dim i As Integer, x As Long, y As Long
Dim Row As Long, Col As Long        '当前图片所在行列
Dim lMoveWidth As Long, lMoveHeight As Long, lDWidth As Long, lDHeight As Long

    picResult.Width = mlngWidth: picResult.Height = mlngHeight
    lWidth = mlngWidth / vsPic.Cols '预定绘图区域宽高
    lHeight = mlngHeight / vsPic.Rows
    
    For i = 1 To picTmp.UBound
        '计算图片应该在第几行第几列
        Row = i \ vsPic.Cols: If i Mod vsPic.Cols <> 0 Then Row = Row + 1
        Col = i Mod vsPic.Cols: If Col = 0 Then Col = vsPic.Cols
        Row = Row - 1: Col = Col - 1 'vs行列以0起计
        
        '计算图片宽高偏移,从而保持宽高比
        lDWidth = (picTmp(i).Picture.Width / picTmp(i).Picture.Height) * lHeight '用图片宽高比*目标区域高度，得出目标区域宽度
        If lDWidth <= lWidth Then '如果目标宽度小于预定区域,表明保持宽度能容下
            lMoveWidth = (lWidth - lDWidth) / 2
            lMoveHeight = 0
        Else
            lDHeight = (picTmp(i).Picture.Height / picTmp(i).Picture.Width) * lWidth '用图片高宽比*目标区域宽度，得出目标区域高度
            If lDHeight <= lHeight Then
                lMoveWidth = 0
                lMoveHeight = (lHeight - lDHeight) / 2
            Else
                lMoveWidth = 0
                lMoveHeight = 0
            End If
        End If
                
        '按设定绘图
        picResult.PaintPicture picTmp(i).Picture, Col * lWidth + lMoveWidth, Row * lHeight + lMoveHeight, lWidth - (lMoveWidth * 2), lHeight - (lMoveHeight * 2)
    Next
    Set picResult.Picture = picResult.Image
End Sub
Private Sub mPacsimg_InsertPicture(pic As stdole.StdPicture)
    zlAddPic pic
End Sub
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

Private Sub vsPic_SelChange()
    shpBorder.Move vsPic.ColWidth(vsPic.Col) * vsPic.Col - 1, vsPic.RowHeight(vsPic.Row) * vsPic.Row - 1, vsPic.ColWidth(vsPic.Col) + 2, vsPic.RowHeight(vsPic.Row) + 2
    shpBorder.Visible = True
End Sub
