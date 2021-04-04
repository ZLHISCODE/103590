VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPACSImg 
   Caption         =   "打开图像"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   Icon            =   "frmPacsImg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   9420
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2295
      ScaleWidth      =   5775
      TabIndex        =   1
      Top             =   4200
      Width           =   5775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消"
         Height          =   350
         Left            =   3480
         TabIndex        =   5
         Top             =   1920
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定"
         Height          =   350
         Left            =   600
         TabIndex        =   4
         Top             =   1920
         Width           =   1100
      End
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   1455
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   2415
         _Version        =   262147
         _ExtentX        =   4260
         _ExtentY        =   2566
         _StockProps     =   35
         BackColor       =   0
      End
   End
   Begin MSComctlLib.ListView lvwSeq 
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwImage 
      Height          =   1695
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2990
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "iLsTree32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin XtremeCommandBars.ImageManager ImgIcons 
      Left            =   960
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPacsImg.frx":0CCA
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "frmPacsImg.frx":3C0A
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPACSImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintSelectAllSeq As Integer                 '0--无状态；1--选择全部序列；2--不选择全部序列
Private mintSelectAllImg As Integer                 '0--无状态；1--选择全部图像；2--不选择全部图像

Private mstrImageIDs As String                      '记录需要打开的图像在正本图像中的ID串，用“-”分隔
Private mfrmViewer As frmViewer
Private strRegPath As String

Public Function zlOpenImages(frmParent As frmFilm, frmViewer As frmViewer) As String
    Set mfrmViewer = frmViewer
    Call ShowSeqImg
    Me.Show 1, frmParent
    
    zlOpenImages = mstrImageIDs
End Function

'执行菜单命令
Public Sub zlMenuClick(mnuClick As String)
    
    Select Case mnuClick
        Case "全选序列"
            If mintSelectAllSeq = 0 Or mintSelectAllSeq = 2 Then
                mintSelectAllSeq = 1
            ElseIf mintSelectAllSeq = 1 Then
                mintSelectAllSeq = 0
            End If
            Call subSetMenuState
            SelectAllSeq True
        Case "全清序列"
            If mintSelectAllSeq = 0 Or mintSelectAllSeq = 1 Then
                mintSelectAllSeq = 2
            ElseIf mintSelectAllSeq = 2 Then
                mintSelectAllSeq = 0
            End If
            Call subSetMenuState
            SelectAllSeq False
        Case "全选图像"
            If mintSelectAllImg = 0 Or mintSelectAllImg = 2 Then
                mintSelectAllImg = 1
            ElseIf mintSelectAllImg = 1 Then
                mintSelectAllImg = 0
            End If
            Call subSetMenuState
            SelectAllImg True
        Case "全清图像"
            If mintSelectAllImg = 0 Or mintSelectAllImg = 1 Then
                mintSelectAllImg = 2
            ElseIf mintSelectAllImg = 2 Then
                mintSelectAllImg = 0
            End If
            Call subSetMenuState
            SelectAllImg False
        Case "反选图像"
            Dim i As Integer
            With lvwImage
                For i = 1 To .ListItems.Count
                    .ListItems(i).Checked = Not .ListItems(i).Checked
                Next
            End With
            Call WriteSelectdImages(lvwImage.Tag)
    End Select
End Sub

Private Sub subSetMenuState()
    If mintSelectAllSeq = 0 Then            '0--无状态
        Me.cbrMain.FindControl(, ID_PacsImg_SelectAllSeries).Checked = False
        Me.cbrMain.FindControl(, ID_PacsImg_UnSelectAllSeries).Checked = False
    ElseIf mintSelectAllSeq = 1 Then        '1--选择全部序列
        Me.cbrMain.FindControl(, ID_PacsImg_SelectAllSeries).Checked = True
        Me.cbrMain.FindControl(, ID_PacsImg_UnSelectAllSeries).Checked = False
    ElseIf mintSelectAllSeq = 2 Then        '2--不选择全部序列
        Me.cbrMain.FindControl(, ID_PacsImg_SelectAllSeries).Checked = False
        Me.cbrMain.FindControl(, ID_PacsImg_UnSelectAllSeries).Checked = True
    End If
    
    If mintSelectAllImg = 0 Then            '0--无状态
        Me.cbrMain.FindControl(, ID_PacsImg_SelectAllImages).Checked = False
        Me.cbrMain.FindControl(, ID_PacsImg_UnSelectAllImages).Checked = False
    ElseIf mintSelectAllImg = 1 Then        '1--选择全部图像
        Me.cbrMain.FindControl(, ID_PacsImg_SelectAllImages).Checked = True
        Me.cbrMain.FindControl(, ID_PacsImg_UnSelectAllImages).Checked = False
    ElseIf mintSelectAllImg = 2 Then        '2--不选择全部图像
        Me.cbrMain.FindControl(, ID_PacsImg_SelectAllImages).Checked = False
        Me.cbrMain.FindControl(, ID_PacsImg_UnSelectAllImages).Checked = True
    End If
End Sub

Private Sub SelectAllSeq(ByVal blnSelect As Boolean)
    Dim i As Integer
    With lvwSeq
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next
        If Not lvwSeq.SelectedItem Is Nothing Then
            ShowImageList lvwSeq.SelectedItem
        Else
            ShowImageList Nothing
        End If
    End With
End Sub

Private Sub SelectAllImg(ByVal blnSelect As Boolean)
    Dim i As Integer
    With lvwImage
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next
    End With
    Call WriteSelectdImages(lvwImage.Tag)
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.Id
        Case ID_PacsImg_SelectAllSeries    '全选序列
            Call zlMenuClick("全选序列")
        Case ID_PacsImg_UnSelectAllSeries      '全清序列
            Call zlMenuClick("全清序列")
        Case ID_PacsImg_SelectAllImages     '全选图像
            Call zlMenuClick("全选图像")
        Case ID_PacsImg_UnSelectAllImages   '全清图像
            Call zlMenuClick("全清图像")
        Case ID_PacsImg_ReverseSelectImages '反选图像
            Call zlMenuClick("反选图像")
    End Select
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.Id
        Case ID_PacsImg_SelectAllSeries, ID_PacsImg_UnSelectAllSeries, ID_PacsImg_SelectAllImages, _
             ID_PacsImg_UnSelectAllImages, ID_PacsImg_ReverseSelectImages
            control.Enabled = lvwSeq.ListItems.Count > 0
    End Select
End Sub

Private Sub CmdCancel_Click()
    mstrImageIDs = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim Item As MSComctlLib.ListItem
    Dim intSeriesNo As Integer
    
    On Error GoTo err
    '组织返回的图像串，规则是“序列号1|1-3;5-27;33-100+序列号2|全部”,全部表示打开全部图象
    For Each Item In lvwSeq.ListItems
        intSeriesNo = Val(Item.SubItems(3))
        If Item.Checked Then    '只有当前序列被勾选了，而且选择可部分图象或者全部图象，才打开该序列
            If Item.SubItems(1) <> "" Then      '为空表示没有选择任何图象
                If mstrImageIDs = "" Then
                    mstrImageIDs = intSeriesNo & "|" & Item.SubItems(1)
                Else
                    mstrImageIDs = mstrImageIDs & "+" & intSeriesNo & "|" & Item.SubItems(1)
                End If
            End If
        End If
    Next Item
    
    Unload Me
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Id = 1 Then
        Item.Handle = lvwSeq.hwnd
    ElseIf Item.Id = 2 Then
        Item.Handle = lvwImage.hwnd
    ElseIf Item.Id = 3 Then
        Item.Handle = picView.hwnd
    End If
End Sub

Private Sub Form_Load()
    mstrImageIDs = ""
    
    '读取本地参数
    strRegPath = "公共模块\zl9PacsCore\frmPacsImg"
    mintSelectAllSeq = Val(GetSetting("ZLSOFT", strRegPath, "SelectAllSeq", 0))
    mintSelectAllImg = Val(GetSetting("ZLSOFT", strRegPath, "SelectAllImg", 0))
    
    '-----------------------------------------------------
    '功能创建工具条
    Call InitCommandBars
    Call subSetMenuState
    Call InitFaceScheme
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub InitFaceScheme()
    Dim pane1 As Pane
    
    With Me.DkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    With DkpMain
        Set pane1 = .CreatePane(1, 0, 300, DockTopOf, Nothing)
            pane1.Handle = lvwSeq.hwnd
            pane1.Options = PaneNoCaption Or PaneNoCloseable
            
        Set pane1 = .CreatePane(2, 0, 300, DockBottomOf, pane1)
            pane1.Handle = lvwImage.hwnd
            pane1.Options = PaneNoCaption Or PaneNoCloseable
            
        Set pane1 = .CreatePane(3, 0, 400, DockBottomOf, Nothing)
            pane1.Handle = picView.hwnd
            pane1.Options = PaneNoCaption Or PaneNoCloseable
    End With
    DkpMain.LoadStateFromString GetSetting("ZLSOFT", strRegPath, DkpMain.Name, "")
End Sub

Private Sub InitCommandBars()
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOfficeXP
    Me.cbrMain.Icons = ImgIcons.Icons
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
    End With

    Me.cbrMain.Item(1).Visible = False                                 '隐藏菜单栏

    '建立主工具栏
    Set cbrToolBar = Me.cbrMain.Add("主工具栏", xtpBarBottom)
    cbrToolBar.Position = xtpBarTop

    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ID_PacsImg_SelectAllSeries, "全选序列")
            cbrControl.IconId = 1001: cbrControl.style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "选中当前所有序列"
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ID_PacsImg_UnSelectAllSeries, "全清序列")
            cbrControl.IconId = 1002: cbrControl.style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "清除选中当前所有序列"
        Set cbrControl = .Add(xtpControlButton, ID_PacsImg_SelectAllImages, "全选图像")
            cbrControl.IconId = 1003: cbrControl.style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "选中当前所有图像"
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ID_PacsImg_UnSelectAllImages, "全清图像")
        cbrControl.IconId = 1004: cbrControl.style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "清除选中当前所有图像"
        Set cbrControl = .Add(xtpControlButton, ID_PacsImg_ReverseSelectImages, "反选图像")
        cbrControl.IconId = 1005: cbrControl.style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "反向选择所有图像"
    End With
End Sub

Private Sub ShowSeqList()
'-----------------------------------------------------------------------------------------
'功能：查询检查序列
'参数：无
'返回：无
'-----------------------------------------------------------------------------------------
    Dim tmpItem As MSComctlLib.ListItem
    Dim strCurKey As String
    Dim i As Integer
    
    On Error GoTo DBError
    If Not lvwSeq.SelectedItem Is Nothing Then strCurKey = lvwSeq.SelectedItem.Key
    
    With lvwSeq
        With .ColumnHeaders
            .Clear
            .Add , , "影像类别", 2000
            .Add , , "打开图像", 2000
            .Add , , "病人ID", 800, 1
            .Add , , "序列号", 800, 1
            .Add , , "图像数", 800, 1
            .Add , , "说明", 2500
        End With
        .ListItems.Add , , "Temp"
        .ListItems.Clear
    End With
    
    '从正本zlSeriesInfo中读取序列信息
    For i = 1 To ZLSeriesInfos.Count
        Set tmpItem = lvwSeq.ListItems.Add(, "_" & ZLSeriesInfos(i).SeriesUID, ZLSeriesInfos(i).strModality)
        With tmpItem
            If mintSelectAllImg = 0 Or mintSelectAllImg = 1 Then    '0--无状态；1--选择全部图像；2--不选择全部图像
                .SubItems(1) = "全部"
            Else
                .SubItems(1) = ""
            End If
            
            .SubItems(2) = ZLSeriesInfos(i).strModality
            .SubItems(3) = i
            .SubItems(4) = ZLSeriesInfos(i).ImageInfos.Count
            .SubItems(5) = ZLSeriesInfos(i).SeriesNo
            
            If .Key = strCurKey Then .Selected = True
        End With
    Next i
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowImageList(ByVal Item As MSComctlLib.ListItem)
'-----------------------------------------------------------------------------------------
'功能：查询检查序列
'参数：无
'返回：无
'-----------------------------------------------------------------------------------------
    Dim strSeriesUID As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim strCurKey As String
    Dim strOpenImages As String
    Dim ImagesArray() As String
    Dim iSegment As Integer
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iSegCount As Integer
    Dim i As Integer
    Dim iSeriesNo As Integer
    Dim imgs As DicomImages
    
    If Not lvwImage.SelectedItem Is Nothing Then strCurKey = lvwImage.SelectedItem.Key
    With lvwImage
        With .ColumnHeaders
            .Clear
            .Add , , "图像号", 2000
            .Add , , "图像描述", 6000
        End With
        .ListItems.Add , , "Temp"
        .ListItems.Clear
    End With
    
    If Item Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo err
    strOpenImages = Item.SubItems(1)
    If strOpenImages <> "全部" And strOpenImages <> "" Then
        ImagesArray = Split(strOpenImages, ";")
        iSegment = 0
        iSegCount = UBound(ImagesArray)
        iStart = Split(ImagesArray(iSegment), "-")(0)
        iEnd = Split(ImagesArray(iSegment), "-")(1)
    End If
    strSeriesUID = Mid(Item.Key, 2)
    
    ' 查找序列所在的Index
    iSeriesNo = 0
    For i = 1 To ZLSeriesInfos.Count
        If ZLSeriesInfos(i).SeriesUID = strSeriesUID Then
            iSeriesNo = i
            Exit For
        End If
    Next i
    
    If iSeriesNo <> 0 Then
        lvwImage.Tag = strSeriesUID
        For i = 1 To ZLSeriesInfos(iSeriesNo).ImageInfos.Count
            Set tmpItem = lvwImage.ListItems.Add(, ZLSeriesInfos(iSeriesNo).ImageInfos(i).InstanceUID, i)
            With tmpItem
                .SubItems(1) = ZLSeriesInfos(iSeriesNo).ImageInfos(i).ImageName
                If strOpenImages = "全部" Then
                    tmpItem.Checked = True
                ElseIf strOpenImages = "" Then
                    tmpItem.Checked = False
                Else
                    If i >= iStart And i <= iEnd Then
                        '满足条件，是需要选中的
                        tmpItem.Checked = True
                    ElseIf i > iEnd Then
                        '大于本段终止号码，则段号加1 ，重新调整起始号码和终止号码
                        iSegment = iSegment + 1
                        If iSegment > iSegCount Then
                            tmpItem.Checked = False
                        Else
                            iStart = Split(ImagesArray(iSegment), "-")(0)
                            iEnd = Split(ImagesArray(iSegment), "-")(1)
                            If i >= iStart And i <= iEnd Then
                                tmpItem.Checked = True
                            Else
                                tmpItem.Checked = False
                            End If
                        End If
                    Else
                        '小于本段起始号码，则不选中
                        tmpItem.Checked = False
                    End If
                End If
                If .Key = strCurKey Then .Selected = True
            End With
        Next i
    End If
    
    DViewer.Images.Clear
    
    If lvwImage.ListItems.Count >= 1 Then
        Call ShowLvwImage(lvwImage.ListItems(1))
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    SaveSetting "ZLSOFT", strRegPath, "SelectAllSeq", mintSelectAllSeq
    SaveSetting "ZLSOFT", strRegPath, "SelectAllImg", mintSelectAllImg
    
    Call SaveSetting("ZLSOFT", strRegPath, DkpMain.Name, DkpMain.SaveStateToString)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwImage_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call WriteSelectdImages(lvwImage.Tag)
End Sub

Private Sub lvwImage_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If Item.Checked <> Item.Selected Then
        Item.Checked = Item.Selected
        Call WriteSelectdImages(lvwImage.Tag)
    End If
    Call ShowLvwImage(Item)
End Sub

Private Sub ShowLvwImage(ByVal Item As MSComctlLib.ListItem)
'显示ListView中的图像
    
    Dim intSeriesNo As Integer
    Dim intImageNo As Integer
    Dim tmpImg As DicomImage
    
    On Error GoTo DBError
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
    intSeriesNo = Val(lvwSeq.SelectedItem.SubItems(3))
    intImageNo = Val(Item.Text)
    
    '读取图像到DViewer中
    DViewer.Images.Clear
    Set tmpImg = funLoadAImage(intSeriesNo, intImageNo, 0)
    If Not tmpImg Is Nothing Then
        DViewer.Images.Add tmpImg
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwSeq_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    lvwSeq.SelectedItem = Item
    Call ShowImageList(Item)
End Sub

Private Sub lvwSeq_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked <> Item.Selected Then
        Item.Checked = Item.Selected
    End If
    Call ShowImageList(Item)
End Sub

Private Sub picView_Resize()
    Dim iCols As Integer, iRows As Integer
    
    On Error Resume Next
    With DViewer
        .left = 0: .top = 0
        .width = picView.ScaleWidth
        .height = picView.ScaleHeight - cmdOK.height - 400
        
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .width, .height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
    
    cmdOK.left = picView.ScaleWidth / 4
    cmdOK.top = picView.ScaleHeight - cmdOK.height - 200
    cmdCancel.left = picView.ScaleWidth / 2
    cmdCancel.top = picView.ScaleHeight - cmdCancel.height - 200
    
End Sub

Private Sub ShowSeqImg()
    Call ShowSeqList     '显示序列
    If lvwSeq.SelectedItem Is Nothing Then
        DViewer.Images.Clear
        Call ShowImageList(Nothing)
    ElseIf mintSelectAllSeq = 0 Then
        lvwSeq_ItemClick lvwSeq.SelectedItem
    ElseIf mintSelectAllSeq = 1 Then
        SelectAllSeq True
    ElseIf mintSelectAllSeq = 2 Then
        SelectAllSeq False
    End If
    
    If lvwImage.SelectedItem Is Nothing Then
        DViewer.Images.Clear
    Else
        ShowLvwImage lvwImage.SelectedItem
    End If
End Sub

Private Sub WriteSelectdImages(strSeriesUID As String)
    Dim i As Integer
    Dim j As Integer
    Dim strOpenImages As String
    Dim blnSelectAll As Boolean
    Dim blnSelectNone As Boolean
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iSegment As Integer
    
    blnSelectNone = True
    blnSelectAll = True
    For j = 1 To lvwImage.ListItems.Count
        If lvwImage.ListItems(j).Checked = True Then
            blnSelectNone = False
            '开始记录本段
            If iStart <> 0 Then
                iEnd = lvwImage.ListItems(j).Text
            Else
                iStart = lvwImage.ListItems(j).Text
                iEnd = lvwImage.ListItems(j).Text
            End If
        Else
            blnSelectAll = False
            '结束记录本段
            If iStart <> 0 Then
                iSegment = iSegment + 1
                If strOpenImages = "" Then
                    strOpenImages = iStart & "-" & iEnd
                Else
                    strOpenImages = strOpenImages & ";" & iStart & "-" & iEnd
                End If
                iStart = 0
                iEnd = 0
            End If
        End If
    Next j
    If iStart <> 0 Then
        iSegment = iSegment + 1
        If strOpenImages = "" Then
            strOpenImages = iStart & "-" & iEnd
        Else
            strOpenImages = strOpenImages & ";" & iStart & "-" & iEnd
        End If
    End If
    If blnSelectAll = True Then
        strOpenImages = "全部"
    End If
    If blnSelectNone = True Then
        strOpenImages = ""
    End If
    
    For i = 1 To lvwSeq.ListItems.Count
        If lvwSeq.ListItems(i).Key = "_" & strSeriesUID Then
            lvwSeq.ListItems(i).ListSubItems(1) = strOpenImages
        End If
    Next i
End Sub
