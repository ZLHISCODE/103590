VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmMiniSeries 
   Caption         =   "序列缩略图"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3780
   ControlBox      =   0   'False
   Icon            =   "frmMiniSeries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.TabStrip tabMini 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      TabFixedHeight  =   527
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DicomObjects.DicomViewer MiniVeiwer 
      DragIcon        =   "frmMiniSeries.frx":0CCA
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2415
      _Version        =   262147
      _ExtentX        =   4260
      _ExtentY        =   1085
      _StockProps     =   35
      BackColor       =   0
   End
   Begin XtremeCommandBars.CommandBars cbrPopup 
      Left            =   3240
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMiniSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fViewer As frmViewer
Dim intX As Integer
Dim intY As Integer
Dim mstrStudyUIDArray() As String
Dim mImages As New DicomImages

Private WithEvents mfrmMain As frmViewer
Attribute mfrmMain.VB_VarHelpID = -1


Private Sub cbrPopup_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim iImgGroupNo As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    
    If Me.MiniVeiwer.Images.Count > 0 Then
        iImgGroupNo = Me.MiniVeiwer.ImageIndex(intX, intY) '提取图像在ZLSeriesInfos结构中的索引
         
        If control.Id < fViewer.Viewer.Count Then   '如果Viewer已经存在，则替换Viewer中的内容
            '用新序列的图像代替viewer(index)中的图像
            Call fViewer.funcSwapSeries(control.Id, iImgGroupNo)
        Else    '如果Viewer不存在，则创建新Viewer
            If (control.Id Mod fViewer.intCountX) = 0 Then
                intCol = fViewer.intCountX
                intRow = control.Id / fViewer.intCountX
            Else
                intCol = control.Id Mod fViewer.intCountX
                intRow = Int(control.Id / fViewer.intCountX) + 1
            End If
            Call fViewer.subCreateAndPlaceAViewer(iImgGroupNo, intRow, intCol)
        End If
    End If
End Sub

Private Sub mfrmMain_AfterSeriesChanged(strStudyUID As String, strSeriesUID As String)
    '观片战中选择的序列发生改变，则修改缩略图中当前选中序列的标记
    Dim i As Integer
    Dim iTabIndex As Integer
    
    On Error GoTo err
    
    If SafeArrayGetDim(mstrStudyUIDArray) = 0 Then Exit Sub
    
    iTabIndex = -1
    '通过检查UID，查找对应的TAB index
    For i = 0 To UBound(mstrStudyUIDArray)
        If mstrStudyUIDArray(i) = strStudyUID Then
            iTabIndex = i
            Exit For
        End If
    Next i
    
    '显示序列选择标记
    If iTabIndex <> -1 Then Call ShowTabImage(iTabIndex, strSeriesUID)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub MiniVeiwer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    
    If Button = 1 And Me.MiniVeiwer.Images.Count > 0 And MiniVeiwer.ImageIndex(x, y) <> 0 Then
        MiniVeiwer.Tag = Me.MiniVeiwer.Images(MiniVeiwer.ImageIndex(x, y)).Tag
        MiniVeiwer.Drag
    End If
End Sub

Private Sub MiniVeiwer_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 2 And MiniVeiwer.ImageIndex(x, y) <> 0 Then
        intX = x
        intY = y
        ShowPopup
    End If
End Sub

Private Sub Form_Load()
    Set mfrmMain = frmMain
    Call RestoreWinState(Me, App.ProductName)
    SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶

End Sub

Private Sub Form_Resize()
    Me.tabMini.left = 1
    Me.tabMini.top = 1
    Me.tabMini.width = Abs(Me.ScaleWidth - 1)
    Me.MiniVeiwer.left = 1
    Me.MiniVeiwer.top = Me.tabMini.top + Me.tabMini.height
    Me.MiniVeiwer.height = Abs(Me.ScaleHeight - Me.tabMini.height - 1)
    Me.MiniVeiwer.width = tabMini.width
End Sub

Public Sub CloseMe(Optional dkpPane As DockingPane = Nothing)
    If Not dkpPane Is Nothing Then
        If dkpPane.PanesCount = 2 Then
            dkpPane.Panes(2).Closed = True
        End If
    End If
    Unload Me
End Sub

Public Sub ShowMe(imgs As DicomImages, f As frmViewer, Optional dkpPane As DockingPane = Nothing)
    Dim i As Integer
    Dim iStudyCount As Integer
    Dim blnFound As Boolean
    
    Set fViewer = f
    Set mImages = imgs
    
    Me.MiniVeiwer.Images.Clear
    ReDim mstrStudyUIDArray(0) As String
    
    tabMini.Tabs.Clear
    tabMini.Visible = False
    
    If imgs.Count > 0 Then
        For i = 1 To imgs.Count
            '判断图像是否已经被增加，如果没有，则增加到Tab中，否则不增加
            blnFound = False
            For iStudyCount = 0 To UBound(mstrStudyUIDArray) - 1
                If mstrStudyUIDArray(iStudyCount) = imgs(i).StudyUID Then
                    blnFound = True
                    Exit For
                End If
            Next iStudyCount
            
            If blnFound = False Then
                
                ReDim Preserve mstrStudyUIDArray(UBound(mstrStudyUIDArray) + 1) As String
                mstrStudyUIDArray(UBound(mstrStudyUIDArray) - 1) = imgs(i).StudyUID
                tabMini.Tabs.Add , , imgs(i).Name & "(" & imgs(i).Attributes(&H8, &H60).Value & " " & imgs(i).Attributes(&H8, &H20).Value & ")"
                tabMini.Visible = True
            End If
        Next i
        
        '显示当前Tab对应的图像
        Call ShowTabImage(0, "")
    End If
    
    '显示
    If Not dkpPane Is Nothing Then
        '去掉PACS报告窗体的控制框
        Call zlcontrol.FormSetCaption(Me, False, False)
        
        If dkpPane.PanesCount = 1 Then
            Dim pane1 As Pane
            Dim dblScale As Double
    
            dblScale = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.EXEName & "\frmMiniSeries", "缩略图高度", 0.1)
            dblScale = dblScale * 200
            
            Set pane1 = dkpPane.CreatePane(2, 200, dblScale, DockTopOf, dkpPane.Panes(1))
            pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable
        Else
            dkpPane.Panes(2).Closed = False
        End If
        
        dkpPane.Panes(2).Handle = Me.hwnd
    Else
        If Me.Visible = False Then
            '去掉PACS报告窗体的控制框
            Call zlcontrol.FormSetCaption(Me, True, True)
            Me.Show , fViewer
        End If
    End If
End Sub

Private Sub ShowTabImage(iTabIndex As Integer, strSeriesUID As String)
'------------------------------------------------
'功能：显示指定页面的缩略图
'参数：     iTabIndex --- 缩略图所在的页面
'           strSeriesUID --- 缩略图中被选中的序列
'返回：无，直接显示
'------------------------------------------------
    Dim img As DicomImage
    Dim i As Integer
    Dim strStudyUID As String
    Dim iRows As Integer, iCols As Integer
    Dim strLabel As String
    
    On Error GoTo err
    
    If iTabIndex < 0 Or iTabIndex >= UBound(mstrStudyUIDArray) Then Exit Sub
    
    iRows = 1
    iCols = 1
    Me.MiniVeiwer.Images.Clear
    strStudyUID = mstrStudyUIDArray(iTabIndex)
    For i = 1 To mImages.Count
        If mImages(i).StudyUID = strStudyUID Then
            Set img = mImages(i)
            img.Labels.Clear
            
            '显示图像信息
            If blnShowMiniImageInfo Then
                img.Labels.AddNew
                '显示序列号+ PatientID
                strLabel = img.PatientID & vbCrLf & img.Attributes(&H20, &H11).Value
                'Study Description
                If Not IsNull(img.Attributes(&H8, &H1030).Value) Then
                    strLabel = strLabel & vbCrLf & img.Attributes(&H8, &H1030).Value
                End If
                'Series Description
                If Not IsNull(img.Attributes(&H8, &H103E).Value) Then
                    strLabel = strLabel & vbCrLf & img.Attributes(&H8, &H103E).Value
                End If
                img.Labels(1).Text = strLabel
                img.Labels(1).LabelType = doLabelText
                img.Labels(1).left = 0
                img.Labels(1).top = 0
                img.Labels(1).FontSize = 12
            End If
            img.Tag = i
            '设置图像的一些特殊属性
            '取消自动剪影,因为DicomObjects控件本身对处理剪影有BUG，存在（0028，6100）时，会自动对图像进行剪影，
            '导致晋煤的DSA图像不能正常显示
            '虽然设置图像的mask=0 ,可以取消剪影，但是每次图像被添加到新的Dicomimages之后，自动又将mask设置成1了，
            '这样在程序中无法很好的控制，因此直接去掉（0028，6100）这个属性。
            If Not IsNull(img.Attributes(&H28, &H6100).Value) Then
                img.Attributes.Remove &H28, &H6100
            End If
            
            '处理一个Overlay的显示,Overlay的文字一般是白色的，因此最好把图像底色设置成1
            If Not IsNull(img.Attributes(&H6000, &H15).Value) Then
                If img.Attributes(&H6000, &H15).Value = 1 Then
                    If img.Level = 0 Then img.Level = 1
                    img.OverlayVisible(0) = True
                    img.OverlayColour(0) = lngLabelColor
                End If
            End If
            
            '修改图像的VOILUT
            img.VOILUT = 0
            
            img.BorderColour = vbWhite
            If strSeriesUID = "" And i = 1 Then
                img.BorderWidth = 2
            ElseIf strSeriesUID = img.SeriesUID Then
                img.BorderWidth = 2
            Else
                img.BorderWidth = 0
            End If
            
            Me.MiniVeiwer.Images.Add img
        End If
    Next i
    
    '图像布局
    ResizeRegion Me.MiniVeiwer.Images.Count, Me.MiniVeiwer.width, Me.MiniVeiwer.height, iRows, iCols
    Me.MiniVeiwer.MultiColumns = iCols
    Me.MiniVeiwer.MultiRows = iRows
    Me.MiniVeiwer.CurrentIndex = 1
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    If Not fViewer Is Nothing Then
        If fViewer.DkpMain.PanesCount = 2 Then
            Dim dblScale As Double
            If fViewer.picViewer.ScaleHeight <> 0 Then
                dblScale = Me.ScaleHeight / fViewer.picViewer.ScaleHeight
                SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.EXEName & "\frmMiniSeries", "缩略图高度", dblScale
            End If
        End If
    End If
End Sub

Private Sub ShowPopup()
'功能创建弹出菜单
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim iSeriesCount As Integer
Dim i As Integer

'计算需要显示几个序列选择
iSeriesCount = fViewer.intCountX * fViewer.intCountY

    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrPopup.VisualTheme = xtpThemeOffice2003
    
    
    With Me.cbrPopup.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
    End With
    Me.cbrPopup.EnableCustomization False
    Me.cbrPopup.ActiveMenuBar.Visible = False
    
    '采集工具栏定义
    Set cbrToolBar = Me.cbrPopup.Add("序列选择", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        For i = 1 To iSeriesCount
            Set cbrControl = .Add(xtpControlButton, i, i): cbrControl.ToolTipText = "在第" & i & "个序列打开"
        Next i
    End With
    
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub tabMini_Click()
    Call ShowTabImage(tabMini.SelectedItem.Index - 1, "")
End Sub
