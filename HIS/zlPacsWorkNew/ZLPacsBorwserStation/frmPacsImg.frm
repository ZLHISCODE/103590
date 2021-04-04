VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPACSImg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picView 
      Height          =   2775
      Left            =   360
      ScaleHeight     =   2715
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   4560
      Width           =   4815
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   2055
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2415
         _Version        =   262147
         _ExtentX        =   4260
         _ExtentY        =   3625
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin MSComctlLib.ListView lvwSeq 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   480
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
      Left            =   120
      TabIndex        =   3
      Top             =   2280
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "frmPacsImg.frx":0000
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
Private mlngAdviceID As Long, mlngSendNo As Long
Private mblnShowPic As Boolean
Private mstrPrivs As String
Private mblnMoved As Boolean
Private mblnAddImage As Boolean                 '是否追加图像
Private mShowPhotoNumber As Integer
Private mblnLocalizerBackward As Boolean        '定位片后置
Private iCurImageIndex As Integer
Public pobjPacsCore As zl9PacsCore.clsViewer
Private mintSelectAllSeq As Integer                 '0--无状态；1--选择全部序列；2--不选择全部序列
Private mintSelectAllImg As Integer                 '0--无状态；1--选择全部图像；2--不选择全部图像

Public Function zlRefresh(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal strPrivs As String, _
    Optional ByVal blnMoved As Boolean = False, Optional ByVal blnRefresh As Boolean = False) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo DBError
    If mlngAdviceID = lngAdviceID And mlngSendNo = lngSendNO And Not blnRefresh Then Exit Function
    mblnMoved = blnMoved
    mblnShowPic = False
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mstrPrivs = strPrivs

    '转出的影像不能保存报告
    If mblnMoved Then
        mstrPrivs = Replace(mstrPrivs, "图像操作处理", "")
        mstrPrivs = Replace(mstrPrivs, "图像标注测量", "")
        mstrPrivs = Replace(mstrPrivs, "清除图像", "")
    End If
    
    mShowPhotoNumber = 15
    strSQL = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID =  " & _
             "(Select 执行部门ID From 病人医嘱发送 Where 医嘱ID =[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    While Not rsTemp.EOF
        Select Case rsTemp!参数名
        Case "最大显示缩略图数"
            mShowPhotoNumber = Abs(Nvl(rsTemp!参数值, 15))
        Case "定位片后置"
            mblnLocalizerBackward = Nvl(rsTemp!参数值)
        End Select
        rsTemp.MoveNext
    Wend
    
    Call ShowSeqImg
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'执行菜单命令
Public Sub zlMenuClick(mnuClick As String)
    
    mblnAddImage = False
    Select Case mnuClick
        Case "影像处理"
            DViewer_DblClick
        Case "影像对比"
            mblnAddImage = True
            DViewer_DblClick
        Case "影像显示"
            If Not lvwImage.SelectedItem Is Nothing Then ShowLvwImage lvwImage.SelectedItem
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
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllCollapse).Checked = False
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllExpend).Checked = False
    ElseIf mintSelectAllSeq = 1 Then        '1--选择全部序列
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllCollapse).Checked = True
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllExpend).Checked = False
    ElseIf mintSelectAllSeq = 2 Then        '2--不选择全部序列
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllCollapse).Checked = False
        Me.cbrMain.FindControl(, conMenu_View_Expend_AllExpend).Checked = True
    End If
    
    If mintSelectAllImg = 0 Then            '0--无状态
        Me.cbrMain.FindControl(, conMenu_Manage_SelectAllImages).Checked = False
        Me.cbrMain.FindControl(, conMenu_Manage_UnSelectAllImages).Checked = False
    ElseIf mintSelectAllImg = 1 Then        '1--选择全部图像
        Me.cbrMain.FindControl(, conMenu_Manage_SelectAllImages).Checked = True
        Me.cbrMain.FindControl(, conMenu_Manage_UnSelectAllImages).Checked = False
    ElseIf mintSelectAllImg = 2 Then        '2--不选择全部图像
        Me.cbrMain.FindControl(, conMenu_Manage_SelectAllImages).Checked = False
        Me.cbrMain.FindControl(, conMenu_Manage_UnSelectAllImages).Checked = True
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
    Select Case control.ID
        Case conMenu_View_Show          '显示图像
            mblnShowPic = Not mblnShowPic
            control.Checked = mblnShowPic
            Call zlMenuClick("影像显示")
        Case conMenu_View_Expend_AllCollapse    '全选序列
            Call zlMenuClick("全选序列")
        Case conMenu_View_Expend_AllExpend      '全清序列
            Call zlMenuClick("全清序列")
        Case conMenu_Manage_SelectAllImages     '全选图像
            Call zlMenuClick("全选图像")
        Case conMenu_Manage_UnSelectAllImages   '全清图像
            Call zlMenuClick("全清图像")
        Case conMenu_Manage_ReverseSelectImages '反选图像
            Call zlMenuClick("反选图像")
        Case conMenu_View_Refresh
            Call zlRefresh(mlngAdviceID, mlngSendNo, mstrPrivs, mblnMoved, True)
    End Select
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case conMenu_View_Expend_AllCollapse, conMenu_View_Expend_AllExpend, conMenu_Manage_SelectAllImages, _
             conMenu_Manage_UnSelectAllImages, conMenu_Manage_ReverseSelectImages
            control.Enabled = lvwSeq.ListItems.Count > 0
        Case conMenu_View_Show
            control.Enabled = lvwSeq.ListItems.Count > 0
            control.Checked = mblnShowPic
    End Select
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = lvwSeq.Hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = lvwImage.Hwnd
    ElseIf Item.ID = 3 Then
        Item.Handle = picView.Hwnd
    End If
End Sub

Private Sub DViewer_DblClick()
'显示观片站
    Dim strSerials As String, strSeqUID As String
    Dim Item As MSComctlLib.ListItem
    Dim intImageInverval As Integer
    Dim strImages As String
    
    On Error GoTo CallError
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
    
    '规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象
    strImages = ""
    strSerials = ""
    For Each Item In lvwSeq.ListItems
        strSeqUID = Mid(Item.Key, 2)
        If Item.Checked Then
            '只有当前序列被勾选了，而且选择可部分图象或者全部图象，才打开该序列
            If Item.SubItems(1) <> "" Then          '为空表示没有选择任何图象
                strSerials = strSerials & ",'" & strSeqUID & "'"
                If strImages = "" Then
                    strImages = strSeqUID & "|" & Item.SubItems(1)
                Else
                    strImages = strImages & "+" & strSeqUID & "|" & Item.SubItems(1)
                End If
            End If
        End If
    Next
    If Len(strSerials) = 0 Then         '没有选择任何序列,则默认打开该序列的全部图象
        strSerials = ",'" & Mid(lvwSeq.SelectedItem.Key, 2) & "'"
        strImages = Mid(lvwSeq.SelectedItem.Key, 2) & "|全部"
    End If
    
    strSerials = Mid(strSerials, 2)
    
    intImageInverval = Val(Me.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)

    OpenViewer pobjPacsCore, mlngAdviceID, mblnAddImage, Me, strSerials, mblnMoved, mblnLocalizerBackward, intImageInverval, strImages
    Exit Sub
CallError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    
    With DViewer
        i = .ImageIndex(x, y)
        If i > 0 And i <= .Images.Count And i <> iCurImageIndex Then
            .Images(iCurImageIndex).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            iCurImageIndex = i
        End If
    End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.Tag = "Loading" Then
        Me.Tag = ""
        
    End If
End Sub

Private Sub Form_Load()
    Dim objFileSystem As New Scripting.FileSystemObject
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim Pane1 As Pane
    Dim strRegPath As String
    
    '读取本地参数
    strRegPath = "公共模块\" & App.ProductName & "\frmPacsImg"
    mintSelectAllSeq = Val(GetSetting("ZLSOFT", strRegPath, "SelectAllSeq", 0))
    mintSelectAllImg = Val(GetSetting("ZLSOFT", strRegPath, "SelectAllImg", 0))
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOfficeXP
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        '.SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "影像显示")
            cbrControl.IconId = 825: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "显示当前序列影像缩略图"
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "全选序列")
            cbrControl.IconId = 3010: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "选中当前所有序列"
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "全清序列")
            cbrControl.IconId = 3004: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "清除选中当前所有序列"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_SelectAllImages, "全选图像")
            cbrControl.IconId = 227: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "选中当前所有图像"
            cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_UnSelectAllImages, "全清图像")
        cbrControl.IconId = 229: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "清除选中当前所有图像"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReverseSelectImages, "反选图像")
        cbrControl.IconId = 3012: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "反向选择所有图像"
        Set cbrControl = .Add(xtpControlComboBox, conMenu_Manage_ImageInterval, "图像间隔")
            cbrControl.ToolTipText = "设置打开图像时，图像之间的间隔数量"
            cbrControl.AddItem "0"
            cbrControl.AddItem "2"
            cbrControl.AddItem "3"
            cbrControl.AddItem "4"
            cbrControl.AddItem "5"
            cbrControl.AddItem "7"
            cbrControl.AddItem "10"
            cbrControl.ListIndex = 0
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
            cbrControl.IconId = 791: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.ToolTipText = "刷新当前病人图像序列": cbrControl.Flags = xtpFlagRightAlign
    End With
        
    Call subSetMenuState
       
    With dkpMain
        .SetCommandBars Me.cbrMain
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = False
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
        Set Pane1 = .CreatePane(1, 0, 300, DockTopOf, Nothing)
            Pane1.Handle = lvwSeq.Hwnd
            Pane1.Options = PaneNoCaption Or PaneNoCloseable
            
        Set Pane1 = .CreatePane(2, 0, 300, DockBottomOf, Pane1)
            Pane1.Handle = lvwImage.Hwnd
            Pane1.Options = PaneNoCaption Or PaneNoCloseable
            
        Set Pane1 = .CreatePane(3, 0, 400, DockBottomOf, Nothing)
            Pane1.Handle = picView.Hwnd
            Pane1.Options = PaneNoCaption Or PaneNoCloseable
    End With
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub ShowSeqList()
'-----------------------------------------------------------------------------------------
'功能：查询检查序列
'参数：无
'返回：无
'-----------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim tmpItem As MSComctlLib.ListItem
    Dim strCurKey As String
    
    On Error GoTo DBError
    If Not lvwSeq.SelectedItem Is Nothing Then strCurKey = lvwSeq.SelectedItem.Key
    
    With lvwSeq
        With .ColumnHeaders
            .Clear
            .Add , , "影像类别", 2000
            .Add , , "打开图像", 2000
            .Add , , "检查号", 800, 1
            .Add , , "序列号", 800, 1
            .Add , , "图像数", 800, 1
            .Add , , "说明", 2500
            .Add , , "采集时间", 1800
        End With
        .ListItems.Add , , "Temp"
        .ListItems.Clear
    End With
    
    strSQL = "Select A.序列UID,A.序列号,A.序列描述,A.采集时间,B.影像类别,B.检查号," & _
        " B.检查UID,Sum(1) As 图像数 " & _
        "From 影像检查序列 A,影像检查记录 B,影像检查图象 D " & _
        "Where B.医嘱ID= [1]  And B.发送号= [2] And A.检查UID=B.检查UID  And A.序列UID=D.序列UID " & _
        "Group By A.序列UID,A.序列号,A.序列描述,A.采集时间,B.影像类别,B.检查号,B.检查UID " & _
        "Order By B.影像类别,B.检查号,A.序列号"
    If mblnMoved Then
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
    End If
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID, mlngSendNo)
   
    lvwSeq.Tag = ""
    If Not rsTmp.EOF Then
        lvwSeq.Tag = Nvl(rsTmp("检查UID"))
        Do While Not rsTmp.EOF
            
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTmp("序列UID"), rsTmp("影像类别"))
            With tmpItem
                If mintSelectAllImg = 0 Or mintSelectAllImg = 1 Then
                    .SubItems(1) = "全部"
                Else
                    .SubItems(1) = ""
                End If
                
                .SubItems(2) = Nvl(rsTmp("检查号"))
                .SubItems(3) = Nvl(rsTmp("序列号"))
                .SubItems(4) = Nvl(rsTmp("图像数"), 0)
                .SubItems(5) = Nvl(rsTmp("序列描述"))
                .SubItems(6) = Nvl(rsTmp("采集时间"), Date)
                
                If .Key = strCurKey Then .Selected = True
            End With
            rsTmp.MoveNext
        Loop
    End If

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
    strSQL = "Select 图像号,图像描述,图像UID From 影像检查图象 Where 序列UID = [1] Order By 图像号"
    If mblnMoved Then
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取图像信息", strSeriesUID)
    
    lvwImage.Tag = ""
    If Not rsTmp.EOF Then
        lvwImage.Tag = strSeriesUID
        Do While Not rsTmp.EOF
            Set tmpItem = lvwImage.ListItems.Add(, rsTmp("图像UID"), rsTmp("图像号"))
            With tmpItem
                .SubItems(1) = Nvl(rsTmp("图像描述"))
                If strOpenImages = "全部" Then
                    tmpItem.Checked = True
                ElseIf strOpenImages = "" Then
                    tmpItem.Checked = False
                Else
                    If rsTmp("图像号") >= iStart And rsTmp("图像号") <= iEnd Then
                        '满足条件，是需要选中的
                        tmpItem.Checked = True
                    ElseIf rsTmp("图像号") > iEnd Then
                        '大于本段终止号码，则段号加1 ，重新调整起始号码和终止号码
                        iSegment = iSegment + 1
                        If iSegment > iSegCount Then
                            tmpItem.Checked = False
                        Else
                            iStart = Split(ImagesArray(iSegment), "-")(0)
                            iEnd = Split(ImagesArray(iSegment), "-")(1)
                            If rsTmp("图像号") >= iStart And rsTmp("图像号") <= iEnd Then
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
            rsTmp.MoveNext
        Loop
    End If
    
    DViewer.Images.Clear: iCurImageIndex = 0
    
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
    Dim strRegPath As String
    
    strRegPath = "公共模块\" & App.ProductName & "\frmPacsImg"
    SaveSetting "ZLSOFT", strRegPath, "SelectAllSeq", mintSelectAllSeq
    SaveSetting "ZLSOFT", strRegPath, "SelectAllImg", mintSelectAllImg
    
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
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
    Dim strImageUID As String
    
    If mblnShowPic = False Then
        DViewer.Images.Clear
        Exit Sub
    End If
    
    On Error GoTo DBError
    strImageUID = Item.Key
    '读取图像到DViewer中
    GetAllImages DViewer, mblnMoved, 3, 0, lvwImage.Tag, 1, 1, False, "", strImageUID

    If DViewer.Images.Count > 0 Then
        iCurImageIndex = 1
    Else
        iCurImageIndex = 0
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwSeq_DblClick()
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
    DViewer_DblClick
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
        .Left = 0: .Top = 0
        .Width = picView.ScaleWidth: .Height = picView.ScaleHeight
        
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
End Sub

Public Function ZLfun3DImgProcess() As String
'------------------------------------------------
'功能：三维重建预处理，移动当前被选中序列的图像
'参数：无
'返回：图像被移动的目的目录，如果移动失败则返回空
'------------------------------------------------

    Dim strSeriesUID As String
    Dim Item As MSComctlLib.ListItem
    Dim iSeriesCount As Integer
    
    On Error GoTo CallError
    If lvwSeq.SelectedItem Is Nothing Then
        MsgBox "请选择一个序列进行三维重建。"
        ZLfun3DImgProcess = ""
        Exit Function
    End If
    
    iSeriesCount = 0
    For Each Item In lvwSeq.ListItems
        If Item.Checked Then
            iSeriesCount = iSeriesCount + 1
            strSeriesUID = Mid(Item.Key, 2)
        End If
    Next
    
    '判断是否只有多个序列被选择，三维重建一次只能处理一个序列
    If iSeriesCount <> 1 Then
        MsgBox "请选择一个序列进行三维重建，每次重建只能选择一个系列。"
        ZLfun3DImgProcess = ""
        Exit Function
    End If
    
    '移动指定序列UID的图像
    ZLfun3DImgProcess = funMove3DImage(strSeriesUID, mblnMoved)
    Exit Function
CallError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ZLfun3DImgProcess = ""
End Function

Private Function funMove3DImage(strSeriesUID As String, blnMoved As Boolean) As String
'------------------------------------------------
'功能：将一个序列的图像移动到3D临时目录中，等待三维重建软件的调用
'参数：
'       lngAdviceID --  医嘱ID
'       strSeriesUID -- 图像的序列UID
'       blnMoved -- 图像是否被转储
'返回：图像被移动的目的目录，如果移动失败则返回空
'------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim str3DCachePath As String
    Dim strTmpFile As String
    Dim strImageFullPath As String
    
    strSQL = "Select A.图像号,D.FTP用户名 As User1,D.FTP密码 As Pwd1," & _
        "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As 图像目录,A.图像UID,d.设备号 as 设备号1, " & _
        "E.FTP用户名 As User2,E.FTP密码 As Pwd2," & _
        "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2," & _
        "e.设备号 as 设备号2,C.检查UID,B.序列UID " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) "
    If blnMoved Then
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If

    On Error GoTo DBError
    strSQL = strSQL & "And A.序列UID= [1] Order By A.图像号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取图像", strSeriesUID)
    
    If rsTmp.RecordCount > 0 Then
        
        '创建本地目录,3D图像目录由前缀"App.Path & "\TmpImage\3D"+接收日期+检查UID+序列UID
        str3DCachePath = App.Path & "\TmpImage\3D\" & Replace(Nvl(rsTmp("图像目录")), "/", "\") & "\" & strSeriesUID & "\"
        strImageFullPath = App.Path & "\TmpImage\" & Replace(Nvl(rsTmp("图像目录")), "/", "\") & "\"
        MkLocalDir str3DCachePath

        On Error GoTo DBError
        
        Do While Not rsTmp.EOF
            '如果3D目录下没有图像，再检查本地缓存目录，最后再从FTP下载图像
            strTmpFile = str3DCachePath & Nvl(rsTmp("图像UID"))
            If Dir(strTmpFile) = vbNullString Then  '有图像则不需要做任何操作
                If Dir(strImageFullPath & Nvl(rsTmp("图像UID"))) = vbNullString Then
                    '本地缓存图像不存在，则读取FTP图像
                    '建立FTP连接
                    If rsTmp("设备号1") <> vbNullString And Inet1.hConnection = 0 Then
                        If Inet1.FuncFtpConnect(Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))) = 0 Then
                            If rsTmp("设备号2") <> vbNullString Then
                                If Inet2.FuncFtpConnect(Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))) = 0 Then
                                    MsgBox "FTP不能正常连接，请检查网络设置。"
                                    funMove3DImage = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    If Inet1.FuncDownloadFile(Nvl(rsTmp("Root1")) & rsTmp("图像目录"), strTmpFile, rsTmp("图像UID")) <> 0 Then
                        '从设备号1提取图像失败，则从设备号2提取图像
                        If rsTmp("设备号2") <> vbNullString Then
                            If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
                            Call Inet2.FuncDownloadFile(Nvl(rsTmp("Root2")) & rsTmp("图像目录"), strTmpFile, rsTmp("图像UID"))
                        End If
                    End If
                Else
                '本地观片缓存中图像存在，直接复制到3D目录
                    FileCopy strImageFullPath & Nvl(rsTmp("图像UID")), strTmpFile
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    funMove3DImage = str3DCachePath
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    funMove3DImage = ""
End Function

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
