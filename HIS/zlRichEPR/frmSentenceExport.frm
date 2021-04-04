VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmSentenceExport 
   AutoRedraw      =   -1  'True
   Caption         =   "词句导出"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmSentenceExport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   9615
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picNowList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   480
      ScaleHeight     =   3735
      ScaleWidth      =   2775
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton cmdYes 
         Caption         =   " 确 定"
         Height          =   380
         Left            =   840
         TabIndex        =   6
         Top             =   3300
         Width           =   1100
      End
      Begin MSComctlLib.TreeView tvwNowList 
         Height          =   3255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   5741
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "imgClass"
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   3600
      ScaleHeight     =   1335
      ScaleWidth      =   2160
      TabIndex        =   1
      Top             =   2400
      Width           =   2160
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   975
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   1575
         _cx             =   2778
         _cy             =   1720
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
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
      End
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1931
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "imgClass"
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6915
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14076
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgClass 
      Left            =   2160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":6852
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":6DEC
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":7386
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":7920
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":7EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":8454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmSentenceExport.frx":8D2E
   End
   Begin XtremeCommandBars.ImageManager imgTools 
      Left            =   600
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmSentenceExport.frx":8DCB
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSentenceExport.frx":1087B
      Left            =   1080
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSentenceExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    Range = 0: Choose: Num: pName: Depart: personnel: clasID: ID: Class
End Enum
Public Event zlRefParentTree()      '刷新词句列表
Private mlngFileID As Long, mlngWordId As Long
Private mtion As Collection         '存储不需要导入或者导出的词句id
Private mblnDifference As Boolean   '标示导入导出，为真表示导出，为假表示导入
Private msrtXmlPathName As String   '记录xml文件的路径
Private mstrHigher As String        '保存上级分类名称
Private objBar As CommandBar        '工具栏
Private mColClass As Collection     '记录分类

Private oDoc As DOMDocument         'xml文档
Private oRoot  As IXMLDOMElement    '根节点
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strTag As String
    Select Case Control.ID
        Case 10:
            Call ImportOrExport
        Case 11:
            Call ImportOrExport
        Case 15:
            On Error GoTo ErrHandle
            With dlgThis
                On Error Resume Next
                .DialogTitle = "打开文件"
                .Filter = "*.ZIP|*.zip"
                .flags = &H80000 + &H1000 + &H200000 + &H800
                .CancelError = True
                .InitDir = "C:\APPSOFT"
                .ShowOpen
            If Err.Number = 32755 Then Err.Clear: Exit Sub
            '进行解压缩处理
            Dim strFilePath As String '临时文件，解压过后的xml文件；关闭窗体时删除（因为点击导入的时候也需要，为节约时间）
            strFilePath = zlFilesUnZip(.Filename)
            Set oDoc = New DOMDocument
            oDoc.Load strFilePath
            If gobjFSO.FileExists(strFilePath) Then gobjFSO.DeleteFile strFilePath, True
            '如果不包含任何元素，则退出
            If oDoc.documentElement Is Nothing Then
                MsgBox "该文件数据格式不正确或已被损坏！", vbInformation, gstrSysName: Exit Sub
            End If
            If oDoc Is Nothing Then
                Set oDoc = New DOMDocument
            End If
            msrtXmlPathName = strFilePath
            oDoc.Load strFilePath
            If gobjFSO.FileExists(strFilePath) Then gobjFSO.DeleteFile msrtXmlPathName, True
            Set oRoot = oDoc.selectSingleNode("Document")       'oRoot置为根节点
            '如果不包含任何元素，则退出
            If Not oDoc.documentElement Is Nothing Then
                Call zlXmlTree
            End If
            End With
        Case 16:
            Unload Me
        Case 17:
            CheckAllOrClearAll (True)
        Case 18:
            CheckAllOrClearAll (False)
        Case 22:
            strTag = Me.tvwClass.SelectedItem.Tag
            Me.tvwClass.SelectedItem.Tag = Split(strTag, vbCrLf)(0) & vbCrLf & Split(strTag, vbCrLf)(1) & vbCrLf & vbCrLf & Split(strTag, vbCrLf)(3)
            Me.tvwClass.SelectedItem.ForeColor = RGB(0, 0, 0)
        Case 23, 24:
            Me.tvwClass.SelectedItem.Checked = Not Me.tvwClass.SelectedItem.Checked
        Case 26:
            strTag = Me.tvwClass.SelectedItem.Tag
            Me.tvwClass.SelectedItem.Tag = Split(strTag, vbCrLf)(0) & vbCrLf & Split(strTag, vbCrLf)(1) & vbCrLf & Split(strTag, vbCrLf)(2) & vbCrLf
            Me.tvwClass.SelectedItem.ForeColor = RGB(0, 0, 0)
    End Select
    Exit Sub
ErrHandle:
    If Err.Number = 32755 Then
        MsgBox "取消操作", vbInformation, "提示"
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub


Private Sub cbsThis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    '对树形结构进行定位
    If CommandBar.Title = "上级分类" Then
        If Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(3) <> "" Then
            Me.tvwNowList.Nodes("_" & Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(3)).Selected = True
        Else
            If Me.tvwNowList.Nodes.Count > 0 Then Me.tvwNowList.Nodes(1).Selected = True
        End If
        Me.tvwNowList.Tag = 1
        
    ElseIf CommandBar.Title = "添加到指定分类" Then
        If Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(2) <> "" Then
            Me.tvwNowList.Nodes("_" & Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(2)).Selected = True
        Else
           If Me.tvwNowList.Nodes.Count > 0 Then Me.tvwNowList.Nodes(1).Selected = True
        End If
        Me.tvwNowList.Tag = 0
        
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    Select Case Control.ID
        Case 2:
            Control.Enabled = Me.tvwClass.Nodes.Count > 0
        Case 10
            Control.Visible = mblnDifference
        Case 11
            Control.Visible = Not mblnDifference
        Case 15
            Control.Visible = Not mblnDifference
        Case 21:
            If Me.tvwNowList.Nodes.Count > 0 Then
                Control.Enabled = Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(2) = "" And Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(3) = ""
            Else
                Control.Enabled = False
            End If
        Case 22:
            If Me.tvwClass.SelectedItem.Parent Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(2) <> ""
            End If
        Case 23:
            Control.Enabled = Not Me.tvwClass.SelectedItem.Checked
        Case 24:
            Control.Enabled = Me.tvwClass.SelectedItem.Checked
        Case 25:
            If Me.tvwNowList.Nodes.Count > 0 Then
                Control.Enabled = Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(3) = "" And Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(2) = ""
            Else
                Control.Enabled = False
            End If
        Case 26:
            Control.Enabled = Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(3) <> ""
    End Select
End Sub
Private Sub ImportOrExport()
    On Error GoTo ErrHandle
    Dim strFileXML As String
    If mblnDifference Then
        With dlgThis
                On Error Resume Next
                .DialogTitle = "保存文件"
                .Filter = "*.ZIP|*.zip"
                .flags = &H200000 + &H2000 + &H2 + &H800
                .CancelError = True
                .InitDir = "C:\APPSOFT"
                .Filename = "Sentence.zip"
                .ShowSave
                If Err.Number = 32755 Then Err.Clear: Exit Sub
                zlCommfun.ShowFlash "请稍候，正在导出..."
                Screen.MousePointer = vbHourglass
                strFileXML = "Sentence.xml"
                Call ToXml(strFileXML)
                '进行压缩处理
                Call zlFilesZip(strFileXML, .Filename)
                zlCommfun.StopFlash
                Screen.MousePointer = vbDefault
                MsgBox "导出成功,文件地址：" & .Filename, vbOKOnly, "提示"
        End With
    Else
        Dim i As Integer
        For i = 1 To Me.tvwClass.Nodes.Count
            If Me.tvwClass.Nodes(i).Checked = True Then Exit For
        Next
        If i = Me.tvwClass.Nodes.Count + 1 Then
            MsgBox "没有选择要导入的分类", vbOKOnly, "提示": Exit Sub
        Else
            zlCommfun.ShowFlash "请稍候，正在导入..."
            Screen.MousePointer = vbHourglass
            
             Call ImportXMLFile
            
            zlCommfun.StopFlash
            Screen.MousePointer = vbDefault
            MsgBox "导入成功!", vbOKOnly, "提示"
            RaiseEvent zlRefParentTree
        End If
    End If
    Unload Me
    Exit Sub
ErrHandle:
    If Err.Number = 32755 Then
        MsgBox "取消操作", vbInformation, "提示"
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub cmdYes_Click()
    Dim strTag As String
    strTag = Me.tvwClass.SelectedItem.Tag
    If Me.tvwNowList.Tag = "" Then
    ElseIf Me.tvwNowList.Tag = 1 Then
        Me.tvwClass.SelectedItem.Tag = Split(strTag, vbCrLf)(0) & vbCrLf & Split(strTag, vbCrLf)(1) & vbCrLf & Split(strTag, vbCrLf)(3) & vbCrLf & Mid(Me.tvwNowList.SelectedItem.Key, 2)
        Me.tvwClass.SelectedItem.ForeColor = RGB(128, 0, 128)
    ElseIf Me.tvwNowList.Tag = 0 Then
        Me.tvwClass.SelectedItem.Tag = Split(strTag, vbCrLf)(0) & vbCrLf & Split(strTag, vbCrLf)(1) & vbCrLf & Mid(Me.tvwNowList.SelectedItem.Key, 2) & vbCrLf & Split(strTag, vbCrLf)(3)
        Me.tvwClass.SelectedItem.Text = Me.tvwNowList.SelectedItem.Text
        Me.tvwClass.SelectedItem.ForeColor = RGB(0, 0, 255)
    End If
    Me.tvwNowList.Tag = ""
    Me.cbsThis.ClosePopups
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = Me.tvwClass.hwnd
        Case 2
            Item.Handle = Me.picList.hwnd
        Case 3
            Item.Handle = Me.rtbText.hwnd
    End Select
End Sub

Private Sub Form_Load()
    Dim cbpPopup As CommandBarPopup
    Dim cbpNew As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbpCustom As CommandBarControlCustom
    Set objBar = Me.cbsThis.Add("Tools", xtpBarTop)
    objBar.ContextMenuPresent = False           '工具栏上点击鼠标右键时不弹出设置菜单
    objBar.ShowTextBelowIcons = False           '工具栏中的按钮文字显示在图标右侧
    objBar.EnableDocking xtpFlagStretched
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = Me.imgTools.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True                 '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbsThis.EnableCustomization False
    Me.cbsThis.ActiveMenuBar.Visible = False
    With objBar.Controls
        Set cbrControl = .Add(xtpControlButton, 15, "打开"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, 10, "导出"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, 11, "导入"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, 17, "全选"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, 18, "全清"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbpNew = .Add(xtpControlPopup, 21, "指定分类"): cbpNew.CommandBar.Title = "添加到指定分类"
        Set cbpCustom = cbpNew.CommandBar.Controls.Add(xtpControlCustom, 211, "添加到指定分类列表"): cbpCustom.Handle = Me.picNowList.hwnd
        cbpNew.ID = 21: cbpNew.Visible = False: cbpNew.BeginGroup = True
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, 22, "取消指定"): cbrControl.Style = xtpButtonIconAndCaption: cbrControl.Visible = False
        Set cbpNew = .Add(xtpControlPopup, 25, "指定上级"): cbpNew.CommandBar.Title = "上级分类"
        Set cbpCustom = cbpNew.CommandBar.Controls.Add(xtpControlCustom, 251, "上级分类列表"): cbpCustom.Handle = Me.picNowList.hwnd
        cbpNew.ID = 25: cbpNew.Visible = False: cbpNew.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, 26, "取消上级"): cbrControl.BeginGroup = True: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 16, "退出"): cbrControl.BeginGroup = True: cbrControl.Style = xtpButtonIconAndCaption
    End With
    
       '设置窗体布局
    dkpMan.SetCommandBars Me.cbsThis
    dkpMan.Options.ThemedFloatingFrames = True
    Dim panThis As Pane, panSub As Pane, panOper As Pane
    
    Set panThis = dkpMan.CreatePane(1, 400, 800, DockLeftOf, Nothing)
    panThis.Title = "词句分类"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    
    Set panThis = dkpMan.CreatePane(2, 1100, 300, DockRightOf, panThis)
    panThis.Title = "词句列表"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    

    Set panSub = dkpMan.CreatePane(3, 1100, 500, DockBottomOf, panThis)
    panSub.Title = "词句内容"
    panSub.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    With vsfList
        .Rows = 1
        .Cols = 9
        .TextMatrix(0, mCol.Choose) = "选择"
        .TextMatrix(0, mCol.Class) = "分类"
        .TextMatrix(0, mCol.Depart) = "部门"
        .TextMatrix(0, mCol.Num) = "编号"
        .TextMatrix(0, mCol.pName) = "名称"
        .TextMatrix(0, mCol.personnel) = "人员"
        .FixedCols = 0
        .ColDataType(mCol.Choose) = flexDTBoolean
        .SelectionMode = flexSelectionByRow
    End With
    
End Sub
Public Function ShowMe(blnDifference As Boolean, frmParent As Object) As Boolean
'---------------------------------------------------------
    '显示病历导入导出主窗体
    'blnDifference 如果为真表示导出，为假则为导入
    'frmParent 上级窗体，用于模态化显示
    '返回值 真为成功 假为失败
'---------------------------------------------------------
    mblnDifference = blnDifference
    On Error GoTo ErrHandle
    If blnDifference Then '导出
         Me.Caption = "病历词句导出"
         If zlRefTree = -1 Then
            MsgBox "数据库不存在词句分类信息", vbOKOnly, "提示"
            Exit Function
         End If
         Me.cbsThis.ActiveMenuBar.Visible = False
         objBar.Visible = True
         
    Else '导入
        Me.Caption = "病历词句导入"
        On Error Resume Next
        With dlgThis
            .DialogTitle = "打开文件"
            .Filter = "*.ZIP|*.zip"
            .flags = &H80000 + &H1000 + &H200000 + &H800
            .CancelError = True
            .InitDir = "C:\APPSOFT"
            .ShowOpen
            If Err.Number = 32755 Then Err.Clear: Exit Function
            '进行解压缩处理
            Dim strFilePath As String '临时文件，解压过后的xml文件；关闭窗体时删除（因为点击导入的时候也需要，为节约时间）
            strFilePath = zlFilesUnZip(.Filename)
            Set oDoc = New DOMDocument
            oDoc.Load strFilePath
            If gobjFSO.FileExists(strFilePath) Then gobjFSO.DeleteFile strFilePath, True
            '如果不包含任何元素，则退出
            If oDoc.documentElement Is Nothing Then
                MsgBox "该文件数据格式不正确或已被损坏！", vbInformation, gstrSysName: Exit Function
            End If
            Set oRoot = oDoc.selectSingleNode("Document")       'oRoot置为根节点
            Call zlXmlTree
            Call picNowList_Resize
        End With
    End If
    ShowMe = True
    Set mtion = New Collection
    Me.Show 0, frmParent
    Exit Function
ErrHandle:
    If Err.Number = 32755 Then
        MsgBox "取消操作", vbInformation, "提示"
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
    ShowMe = False
End Function
Private Function zlRefTree(Optional lngID As Long) As Long
    '功能：刷新装入指定种类的病历文件清单，并定位到指定的文件上
    Dim rsTemp As New ADODB.Recordset
    Dim objNode As MSComctlLib.Node
    
    gstrSQL = "Select ID, 上级id, 编码, 名称, 说明, 范围" & vbNewLine & _
            "From 病历词句分类" & vbNewLine & _
            "Start With 上级id Is Null" & vbNewLine & _
            "Connect By Prior ID = 上级id" & vbNewLine & _
            "Order By Level, 编码"
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTemp.RecordCount < 1 Then
        zlRefTree = -1
        Exit Function
    End If
    With rsTemp
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
                If IsNull(!上级ID) Then
                    Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, !编码 & "-" & !名称, "close")
                Else
                    Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, !编码 & "-" & !名称, "close")
                End If
            objNode.Tag = !说明 & vbCrLf & !范围: objNode.Sorted = True: objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    If Me.tvwClass.Nodes.Count > 0 Then
            If lngID <> 0 Then
                Me.tvwClass.Nodes("_" & lngID).Selected = True
            Else
                Me.tvwClass.Nodes(1).Selected = True
            End If
            If Me.tvwClass.SelectedItem.Children > 0 Then Me.tvwClass.SelectedItem.Expanded = True
            Call tvwClass_NodeClick(Me.tvwClass.SelectedItem)
    Else
        Call tvwClass_NodeClick(Nothing)
    End If
    zlRefTree = Me.tvwClass.Nodes.Count
    Exit Function

Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefTree = Me.tvwClass.Nodes.Count
End Function
Private Function zlXmlTree() As Integer
    Dim oNode As IXMLDOMNode            '父节点
    Dim oSubNode1 As IXMLDOMNode        '子节点
    Dim i As Long
    '判断是否是词句xml
    If oRoot.getAttribute("EditType") <> "MedicalWords" Then
        MsgBox "此xml文件不是病历词句导出的xml,无法从此xml导入病历词句", vbOKOnly, "提示"
        Exit Function
    End If
    '获取基础信息
    On Error Resume Next
    mstrHigher = ""
    Me.tvwClass.Nodes.Clear
    Call zlTree(oRoot)
    zlXmlTree = tvwClass.Nodes.Count
    
    '初始化目标库的病历词句
    Dim rsTemp As New ADODB.Recordset
    Dim objNode As MSComctlLib.Node
    gstrSQL = "Select ID, 上级id, 编码, 名称, 说明, 范围" & vbNewLine & _
            "From 病历词句分类" & vbNewLine & _
            "Start With 上级id Is Null" & vbNewLine & _
            "Connect By Prior ID = 上级id" & vbNewLine & _
            "Order By Level, 编码"
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.tvwNowList.Nodes.Clear
        Do While Not .EOF
                If IsNull(!上级ID) Then
                    Set objNode = Me.tvwNowList.Nodes.Add(, , "_" & !ID, !编码 & "-" & !名称, "close")
                Else
                    Set objNode = Me.tvwNowList.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, !编码 & "-" & !名称, "close")
                End If
                objNode.Expanded = True
                objNode.Sorted = True: objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub zlTree(oNode As IXMLDOMNode, Optional strHigher As String)
'----------------------------------------------------------------------------------------------------------
'把xml文件中的节点绑定到树形列表中
'----------------------------------------------------------------------------------------------------------
    Dim objNode As MSComctlLib.Node
    Dim oSubNode1 As IXMLDOMNode        '子节点
    Dim i As Long, rsCount As ADODB.Recordset
    Dim j As Long, introw As Long
    
    For i = 0 To oNode.selectNodes("Class").Length - 1
        Set oSubNode1 = oNode.selectNodes("Class")(i)
        If Not oSubNode1 Is Nothing Then
            If GetNodeValue(oSubNode1, "上级id", "") = "" Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & GetNodeValue(oSubNode1, "ID", ""), GetNodeValue(oSubNode1, "编码", "") & "-" & GetNodeValue(oSubNode1, "名称", ""), "close")
                 objNode.Expanded = True
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & GetNodeValue(oSubNode1, "上级id", ""), tvwChild, "_" & GetNodeValue(oSubNode1, "ID", ""), GetNodeValue(oSubNode1, "编码", "") & "-" & GetNodeValue(oSubNode1, "名称", ""), "close")
            End If
            objNode.Tag = GetNodeValue(oSubNode1, "说明", "") & vbCrLf & GetNodeValue(oSubNode1, "范围", "")
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
    
            If GetNodeValue(oSubNode1, "上级id", "") = "" Then
                gstrSQL = "select a.id from 病历词句分类 a where  a.名称=[1]"
            Else
                gstrSQL = "select a.id from 病历词句分类 a,病历词句分类 b where  a.名称=[1] and  a.上级id=b.id and b.名称=[2]"
            End If
            Set rsCount = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetNodeValue(oSubNode1, "名称", ""), strHigher)
            '判断分类在目标库中是否存在
            If rsCount.RecordCount < 1 Then
                objNode.ForeColor = RGB(255, 20, 147)
                objNode.Tag = objNode.Tag & vbCrLf & vbCrLf
            Else
                objNode.Tag = objNode.Tag & vbCrLf & rsCount("id") & vbCrLf
            End If
            If oSubNode1.selectNodes("Class").Length > 0 Then
                Call zlTree(oSubNode1, GetNodeValue(oSubNode1, "名称", ""))
            End If
        End If
    Next
End Sub

Private Function zlSubRefList(lngFileID As Long, Optional ByVal blnCheck As Boolean) As Long
    '******************************************************************************************************************
    '功能：刷新装入清单，并定位到指定的记录上
    '参数：
    '返回：
    '******************************************************************************************************************
Dim rsTemp As New ADODB.Recordset
Dim i As Integer
    '如果选中的是同一个分类，则不刷新
    If lngFileID = mlngFileID Then Exit Function
    mlngFileID = lngFileID
    '------------------------------------------------------------------------------------------------------------------
       gstrSQL = "Select L.ID, L.分类id, C.编码 || '-' || C.名称 As 分类, L.编号, L.名称, L.通用级, D.名称 As 部门, P.姓名 As 人员" & vbNewLine & _
                "From 病历词句分类 C, 病历词句示范 L, 部门表 D, 人员表 P" & vbNewLine & _
                "Where C.ID = L.分类id And L.科室id = D.ID And L.人员id = P.ID And L.分类id = [1] "

    '------------------------------------------------------------------------------------------------------------------
    If InStr(1, gstrPrivsEpr, "全院病历范文") <> 0 Then
    
    ElseIf InStr(1, gstrPrivsEpr, "科室病历词句") <> 0 Then
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 In (1, 2) And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"

    Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 = 1 And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      L.通用级 = 2 And L.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
     End If
    
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    '--------------------------------------------------------------------------------------------------------------
    With Me.vsfList
        .Rows = 1
        .ScrollBars = flexScrollBarNone '为防止滚动条抖动，先不显示滚动条
        On Error Resume Next
        Do While Not rsTemp.EOF
            .AddItem ""
            i = i + 1
            Select Case rsTemp!通用级
                    Case 0:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-全院"
                    Case 1:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(2).Picture '"2-科室"
                    Case 2:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(3).Picture '"3-个人"
                    Case Else:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-全院"
                End Select
                .Cell(flexcpAlignment, i, mCol.Range) = flexAlignCenterCenter
            .TextMatrix(i, mCol.ID) = rsTemp!ID
            .TextMatrix(i, mCol.clasID) = rsTemp!分类id
            .TextMatrix(i, mCol.Class) = rsTemp!分类
            .TextMatrix(i, mCol.Num) = rsTemp!编号
            .TextMatrix(i, mCol.pName) = rsTemp!名称
            .TextMatrix(i, mCol.Depart) = rsTemp!部门
            .TextMatrix(i, mCol.personnel) = rsTemp!人员
            If blnCheck Then
                '判断是否是不选择不导入词句
                If mtion("_" & rsTemp!ID) <> rsTemp!ID & "" Then
                    .TextMatrix(i, mCol.Choose) = 1
                Else
                    .TextMatrix(i, mCol.Choose) = 0
                End If
            End If
            rsTemp.MoveNext
        Loop
        .ScrollBars = flexScrollBarVertical
    End With
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlSubRefList = Me.vsfList.Rows
End Function
Private Sub zlSubRefText(lngWordId As Long)
    Dim rsTemp As ADODB.Recordset, lngStart As Long, strText As String
    If lngWordId = mlngWordId Then Exit Sub
    mlngWordId = lngWordId
    rtbText.Text = ""
    Err = 0: On Error GoTo Errhand
    gstrSQL = "Select 内容性质, 内容文本, 要素名称, 要素单位 From 病历词句组成 Where 词句id = [1] Order By 排列次序"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngWordId)
    With rsTemp
        Do While Not .EOF
            lngStart = Len(Me.rtbText.Text)
            Me.rtbText.SelStart = lngStart
            Me.rtbText.SelLength = 0
            Select Case !内容性质
            Case 0 '自由文字
                strText = IIf(IsNull(!内容文本), " ", !内容文本)
                With Me.rtbText
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = False
                End With
            Case 1, 2 '1-临时诊治要素,2-固定诊治要素
                strText = IIf(IsNull(!内容文本), "{" & !要素名称 & "}" & !要素单位, "{" & !内容文本 & "}")
                With Me.rtbText
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = True
                End With
            End Select
            .MoveNext
        Loop
        Me.rtbText.SelStart = 0
    End With
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Me.vsfList.Move 0, 0, Me.picList.Width, Me.vsfList.Height
    Me.rtbText.Move 0, Me.vsfList.Height, Me.ScaleWidth, Me.rtbText.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlCommfun.StopFlash
    Set mtion = Nothing
    Set oDoc = Nothing
    Set oRoot = Nothing
    If gobjFSO.FileExists(msrtXmlPathName) Then gobjFSO.DeleteFile msrtXmlPathName, True
    msrtXmlPathName = ""
End Sub
Private Sub piclist_Resize()
    With vsfList
        .Move 0, 0, picList.Width, picList.Height
        .ColWidth(mCol.Choose) = 450
        .ColWidth(mCol.Range) = 300
        .ColWidth(mCol.Class) = 0
        .ColWidth(mCol.Num) = (picList.Width - 750) / 4 - 300
        .ColWidth(mCol.pName) = (picList.Width - 750) / 4 + 300
        .ColWidth(mCol.Depart) = (picList.Width - 750) / 4
        .ColWidth(mCol.personnel) = (picList.Width - 750) / 4
        .ColWidth(mCol.clasID) = 0
        .ColWidth(mCol.ID) = 0
    End With
End Sub
Private Sub ToXml(strFilePath As String)

    Dim oNode As IXMLDOMNode
    Dim Node As MSComctlLib.Node
    Dim j As Integer
    'XML文档
    Set oDoc = New DOMDocument
    '注释
    oDoc.appendChild oDoc.createComment(gstrSysName & "  " & _
        "操作员:" & gstrUserName & "，部门:" & gstrDeptName & "，时间:" & _
        Format(Now(), "YYYY年MM月DD日"))
    '根节点
    Set oRoot = oDoc.createElement("Document")
    Set oDoc.documentElement = oRoot    '设置为根节点
    Call oRoot.setAttribute("EditType", "MedicalWords")
    Set mColClass = New Collection

    On Error Resume Next
    For j = 1 To tvwClass.Nodes.Count
        Set Node = tvwClass.Nodes(j)
        If Node.Checked Then
            If mColClass(Node.Parent.Key) Is Nothing Then
                Set oNode = oRoot
            Else
               Set oNode = mColClass(Node.Parent.Key)
            End If
            Call CreateChild(oNode, 1, Node)
        End If
    Next
    
    Dim pi As IXMLDOMProcessingInstruction
    Set pi = oDoc.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call oDoc.insertBefore(pi, oDoc.childNodes(0))
    '直接保存成文件即可
    oDoc.Save strFilePath
    
    Set mtion = New Collection
    Set mColClass = Nothing
    Set oDoc = Nothing
End Sub
Private Sub CreateChild(Parent As IXMLDOMNode, intNodeId As Integer, Node As MSComctlLib.Node)
'----------------------------------------------------------------------------------------------
'参数说明
'Parent 父亲节点
'intNodeId 等级
'Node 选中的节点
'----------------------------------------------------------------------------------------------
     
    Dim intSenId As Double
    Dim rsSentence As ADODB.Recordset, rsContent As ADODB.Recordset, rsCondition As ADODB.Recordset
    Dim oNode As IXMLDOMNode            '父节点
    Dim oSubNode1 As IXMLDOMNode        '子节点
    Dim oSubNode2 As IXMLDOMNode         '节点
    Dim oSubNode3 As IXMLDOMNode         '节点
    Dim intNodeId1 As Integer, intNodeId2 As Integer, intNodeId3 As Integer
        'If Node.Checked Then
            '词句分类信息
            intNodeId1 = intNodeId + 1
            intNodeId2 = intNodeId + 2
            intNodeId3 = intNodeId + 3
            Set oNode = CreateNode(intNodeId1, Parent, "Class", NODE_ELEMENT, "")
            '把导入过的节点保存
            On Error Resume Next
            If mColClass(Node.Key) Is Nothing Then
                mColClass.Add oNode, Node.Key
            End If
            
            
            CreateNode 1, oNode, "ID", , Mid(Node.Key, 2)
            If Node.Parent Is Nothing Or Parent.nodeName = "Document" Then
                CreateNode intNodeId1, oNode, "上级id", , ""
            Else
                CreateNode intNodeId1, oNode, "上级id", , Mid(Node.Parent.Key, 2)
            End If
            CreateNode intNodeId1, oNode, "编码", , Split(Node.Text, "-")(0)
            CreateNode intNodeId1, oNode, "名称", , Split(Node.Text, "-")(1)
            CreateNode intNodeId1, oNode, "说明", , Split(Node.Tag, vbCrLf)(0)
            CreateNode intNodeId1, oNode, "范围", , Split(Node.Tag, vbCrLf)(1)
                 '词句示范
            Set rsSentence = GetSentence(Mid(Node.Key, 2))
            
            
            Do While Not rsSentence.EOF
                intSenId = NVL(rsSentence!ID)
                On Error Resume Next
                If mtion("_" & intSenId) <> intSenId & "" Then
                    Set oSubNode1 = CreateNode(intNodeId1, oNode, "Sentence", NODE_ELEMENT, "")
                    CreateNode intNodeId2, oSubNode1, "ID", , rsSentence!ID
                    CreateNode intNodeId2, oSubNode1, "分类id", , rsSentence!分类id
                    CreateNode intNodeId2, oSubNode1, "编号", , rsSentence!编号
                    CreateNode intNodeId2, oSubNode1, "名称", , NVL(rsSentence!名称)
                    CreateNode intNodeId2, oSubNode1, "通用级", , NVL(rsSentence!通用级)
                    CreateNode intNodeId2, oSubNode1, "科室id", , NVL(rsSentence!科室ID)
                    CreateNode intNodeId2, oSubNode1, "人员id", , NVL(rsSentence!人员ID)
                    
                    
                    '获取到病历词句条件
                    gstrSQL = "select t.词句id,t.条件项,t.条件值 from 病历词句条件 t where t.词句id=[1]"
                    Set rsCondition = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, intSenId)
                    
                    Do While Not rsCondition.EOF
                        Set oSubNode3 = CreateNode(intNodeId2, oSubNode1, "Condition", NODE_ELEMENT, "")
                        CreateNode intNodeId3, oSubNode3, "词句id", , rsCondition!词句id
                        CreateNode intNodeId3, oSubNode3, "条件项", , rsCondition!条件项
                        CreateNode intNodeId3, oSubNode3, "条件值", , rsCondition!条件值
                        rsCondition.MoveNext
                    Loop
                    '获取到对应词句的内容
                    gstrSQL = "select t.词句id,t.排列次序,t.内容性质, t.内容文本,t.诊治要素id,t.替换域,t.要素名称," & _
                                "t.要素类型,t.要素长度,t.要素小数,t.要素单位,t.要素表示,t.要素值域,t.输入形态,t.对象属性 " & _
                                " From 病历词句组成 t Where 词句id = [1] Order By t.排列次序"

                    Set rsContent = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, intSenId)
                    Do While Not rsContent.EOF
                        Set oSubNode2 = CreateNode(intNodeId2, oSubNode1, "Content", NODE_ELEMENT, "")
                        CreateNode intNodeId3, oSubNode2, "词句id", , rsContent!词句id
                        CreateNode intNodeId3, oSubNode2, "排列次序", , rsContent!排列次序
                        CreateNode intNodeId3, oSubNode2, "内容性质", , rsContent!内容性质
                        CreateNode intNodeId3, oSubNode2, "内容文本", , NVL(rsContent!内容文本)
                        CreateNode intNodeId3, oSubNode2, "诊治要素id", , NVL(rsContent!诊治要素ID)
                        CreateNode intNodeId3, oSubNode2, "替换域", , NVL(rsContent!替换域)
                        CreateNode intNodeId3, oSubNode2, "要素名称", , NVL(rsContent!要素名称)
                        CreateNode intNodeId3, oSubNode2, "要素长度", , NVL(rsContent!要素长度)
                        CreateNode intNodeId3, oSubNode2, "要素小数", , NVL(rsContent!要素小数)
                        CreateNode intNodeId3, oSubNode2, "要素单位", , NVL(rsContent!要素单位)
                        CreateNode intNodeId3, oSubNode2, "要素表示", , NVL(rsContent!要素表示)
                        CreateNode intNodeId3, oSubNode2, "要素值域", , NVL(rsContent!要素值域)
                        CreateNode intNodeId3, oSubNode2, "输入形态", , NVL(rsContent!输入形态)
                        CreateNode intNodeId3, oSubNode2, "对象属性", , NVL(rsContent!对象属性)
                        rsContent.MoveNext
                    Loop
                End If
            rsSentence.MoveNext
        Loop
    'End If
    '子分类
'    If Node.Children > 0 Then
'        For j = 1 To tvwClass.Nodes.Count
'            If Not tvwClass.Nodes(j).Parent Is Nothing Then
'                If tvwClass.Nodes(j).Parent.Key = Node.Key Then
'                    If oNode Is Nothing Then
'                        Call CreateChild(Parent, intNodeId, tvwClass.Nodes(j))
'                    Else
'                        Call CreateChild(oNode, intNodeId, tvwClass.Nodes(j))
'                    End If
'                End If
'            End If
'        Next
'    End If
End Sub
Private Function GetSentence(lngFileID As Long) As ADODB.Recordset
'------------------------------------------------------------------------------------------------
'作用：用于获取病历词句分类数据
'参数：分类id
'------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------
       gstrSQL = "select L.ID, L.分类id, L.编号, L.名称, L.通用级,l.科室id,l.人员id " & vbNewLine & _
                "From 病历词句分类 C, 病历词句示范 L, 部门表 D, 人员表 P" & vbNewLine & _
                "Where C.ID = L.分类id And L.科室id = D.ID And L.人员id = P.ID And L.分类id = [1] "

    '------------------------------------------------------------------------------------------------------------------
    If InStr(1, gstrPrivsEpr, "全院病历范文") <> 0 Then
    
    ElseIf InStr(1, gstrPrivsEpr, "科室病历词句") <> 0 Then
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 In (1, 2) And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"

    Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 = 1 And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      L.通用级 = 2 And L.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
     End If
    
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    Set GetSentence = rsTemp
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub ImportXMLFile()
'-------------------------------------------------------------------------------------------------
'作用：用于导入词句的时候进行容错处理，主要是对分类的处理
'-------------------------------------------------------------------------------------------------
    Dim objNode As MSComctlLib.Node
    Dim j As Long, i As Long, strHigherId As String
    Dim lngItemID As Long, strMaxNum As String
    Dim rsMaxNum As ADODB.Recordset
    Dim oNode As IXMLDOMNode
    On Error GoTo Errhand
    For j = 1 To Me.tvwClass.Nodes.Count
        Set objNode = tvwClass.Nodes(j)
            '判断是否导入该分类
            If objNode.Checked = True Then
            Set oNode = oRoot.getElementsByTagName("Class[ID=" & Mid(objNode.Key, 2) & "]")(0)
                Debug.Print oNode.selectSingleNode("名称").Text
                '判断分类是否加入指定分类
                If Split(objNode.Tag, vbCrLf)(2) <> "" Then
                    '是 导入该分类
                    lngItemID = Split(objNode.Tag, vbCrLf)(2)
                Else '否 创建分类，导入
                    '判读是否指定上级分类
                    If Split(objNode.Tag, vbCrLf)(3) = "" Then
                        '判读是否是根节点
                        If GetNodeValue(oNode, "上级id", "") = "" Then
                            lngItemID = zldatabase.GetNextId("病历词句分类")
                            gstrSQL = "select max(to_number(编码)) as 编码 from 病历词句分类 t where t.上级id is null"
                            strHigherId = "null"
                        Else
                            '判读当前导入分类的上级分类是否指定
                            If Split(objNode.Parent.Tag, vbCrLf)(2) <> "" Then
                                strHigherId = Split(objNode.Parent.Tag, vbCrLf)(2)
                                lngItemID = zldatabase.GetNextId("病历词句分类")
                                gstrSQL = "select 编码 from 病历词句分类 where 编码=(select max(to_number(编码)) as 编码 from 病历词句分类 t where t.上级id=[1])"
                            Else
                                '在数据库中查询上级分类的id
                                Dim strSQL As String, rsLevel As ADODB.Recordset
                                strSQL = "Select ID, 上级id, 编码, 名称, 说明, 范围 " & vbNewLine & _
                                            "From 病历词句分类 where 名称=[1] and level=[2] " & vbNewLine & _
                                            "Start With 上级id Is Null" & vbNewLine & _
                                            "Connect By Prior ID = 上级id" & vbNewLine & _
                                            "Order By ID desc"
                                Set rsLevel = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(objNode.Parent.Text, InStr(objNode.Parent.Text, "-") + 1), UBound(Split(objNode.FullPath, "\")))
                                Debug.Print Mid(objNode.Parent.Text, InStr(objNode.Parent.Text, "-") + 1) & "|" & UBound(Split(objNode.FullPath, "\"))
                                If rsLevel.RecordCount > 0 Then
                                    strHigherId = rsLevel("ID")
                                    strMaxNum = rsLevel("编码")
                                    gstrSQL = "select 编码 from 病历词句分类 where 编码=(select max(to_number(编码)) as 编码 from 病历词句分类 t where t.上级id=[1])"
                                Else
                                    strHigherId = "null"
                                    gstrSQL = "select max(to_number(编码)) as 编码 from 病历词句分类 t where t.上级id is null"
                                End If
                                lngItemID = zldatabase.GetNextId("病历词句分类")
                            End If
                        End If
                    Else
                        '获取词句id，上级分类id
                        lngItemID = zldatabase.GetNextId("病历词句分类")
                        strHigherId = Split(objNode.Tag, vbCrLf)(3)
                        gstrSQL = "select max(编码) as 编码 from 病历词句分类 t where t.id=[1]"
                        Set rsMaxNum = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strHigherId)
                        If rsMaxNum.RecordCount > 0 Then strMaxNum = NVL(rsMaxNum!编码)
                        gstrSQL = "select max(to_number(编码)) as 编码 from 病历词句分类 t where t.上级id=[1]"
                    End If
                    
                    Set rsMaxNum = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strHigherId)
                    
                    If rsMaxNum.RecordCount > 0 Then
                        strMaxNum = NVL(rsMaxNum!编码)
                        If Mid(strMaxNum, 1, 1) = 0 Then
                            strMaxNum = "0" & strMaxNum + 1
                        Else
                            strMaxNum = Val(strMaxNum) + 1
                        End If
                    Else
                        strMaxNum = strMaxNum & "01"
                    End If
                    If Val(strMaxNum) < 10 Then strMaxNum = "0" & Val(strMaxNum)
                    gstrSQL = "Zl_病历词句分类_Edit(1," & lngItemID & "," & IIf(strHigherId = "", "Null", strHigherId) & ",'" & strMaxNum & "'," & _
                        " '" & GetNodeValue(oNode, "名称", "") & "','" & GetNodeValue(oNode, "说明", "") & "','" & GetNodeValue(oNode, "范围", "") & "')"
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
                Call ImportXmlSentence(oNode, lngItemID)
            End If
    Next
    Exit Sub
Errhand:
    MsgBox Err.Description
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ImportXmlSentence(oNode As IXMLDOMNode, lngItemID As Long)
'-------------------------------------------------------------------------------------------------
'作用:从xml文件中导入词句示范到目标库
'参数说明：
'onode xml分类节点
'lngItemID 分类id
'-------------------------------------------------------------------------------------------------
    Dim oSubNode1 As IXMLDOMNode, oSubNode2 As IXMLDOMNode, oSubNode3 As IXMLDOMNode      '子节点
    Dim i As Long, j As Long, k As Long, ArraySQL() As String, lngCount As Long
    Dim lngWordId As Long '病历词句id
    Dim blnTran As Boolean
    Dim rsMaxNumber As ADODB.Recordset, strNumber As String
    
    On Error GoTo Errhand
    '获取到对应树形结构中的分类
    For i = 0 To oNode.selectNodes("Sentence").Length - 1
        Set oSubNode1 = oNode.selectNodes("Sentence")(i)
        On Error Resume Next
        If mtion("_" & GetNodeValue(oSubNode1, "ID", "")) <> GetNodeValue(oSubNode1, "ID", "") Then
            gstrSQL = "select max(编号) as 编号 from 病历词句示范  where 分类id=[1]"
            Set rsMaxNumber = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngItemID)
            If rsMaxNumber.RecordCount > 0 Then
                Dim intLen As Integer
                strNumber = NVL(rsMaxNumber("编号"))
                If strNumber = "" Then
                    strNumber = "00001"
                Else
                    intLen = Len(strNumber)
                    strNumber = strNumber + 1
                    For j = 1 To intLen - Len(strNumber)
                        strNumber = "0" & strNumber
                    Next
                End If
            Else
                strNumber = "00001"
            End If
            '添加病历词句示范
            lngWordId = zldatabase.GetNextId("病历词句示范")
            gstrSQL = lngWordId & "," & lngItemID & ",'" & strNumber & "','" & GetNodeValue(oSubNode1, "名称", "") & "'"
            Select Case GetNodeValue(oSubNode1, "通用级", "")
                Case 0:
                    gstrSQL = gstrSQL & ",0"
                Case 1:
                    gstrSQL = gstrSQL & ",1"
                Case 2:
                    gstrSQL = gstrSQL & ",2"
                Case Else:
                    gstrSQL = gstrSQL & ",null"
            End Select
            gstrSQL = gstrSQL & "," & glngDeptId & "," & glngUserId
            gstrSQL = "Zl_病历词句示范_Edit(1," & gstrSQL & ")"
            
            '获取SQL语句数组
            ReDim ArraySQL(1 To 2) As String
            ArraySQL(1) = gstrSQL
            '前期处理
            ArraySQL(2) = "Zl_病历词句组成_Beforesave(" & lngWordId & ")"
            
            For j = 0 To oSubNode1.selectNodes("Content").Length - 1
                Set oSubNode2 = oSubNode1.selectNodes("Content")(j)
                '内容等于零是自由文字，1、2是要素
                If GetNodeValue(oSubNode2, "内容性质", "") = 0 Then
                    Dim strIn As String, lngLen As Long, inti As Integer, strSub As String
                    strIn = GetNodeValue(oSubNode2, "内容文本", "")
                    strIn = Replace(strIn, "'", "' || chr(39) || '")
                    strIn = Replace(strIn, vbCrLf, "' || chr(13) || chr(10) || '")  '本来strIn是不允许有vbCrlf的。
                    lngLen = Len(strIn)
                    
                    '按照4000为界分段存储。
                    inti = 0
                    Do While (inti * 2000 + 1 <= lngLen)
                        lngCount = UBound(ArraySQL) + 1
                        ReDim Preserve ArraySQL(1 To lngCount) As String
                    
                        strSub = Mid(strIn, inti * 2000 + 1, 2000)
                    
                        gstrSQL = "Zl_病历词句组成_Insert(" & lngWordId & "," & GetNodeValue(oSubNode2, "排列次序", "") & ",0,'" & strSub & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)"
                        
                        ArraySQL(lngCount) = gstrSQL
                       
                        inti = inti + 1
                    Loop
                Else
                    lngCount = UBound(ArraySQL) + 1
                    ReDim Preserve ArraySQL(1 To lngCount) As String
                    Dim rsCount As ADODB.Recordset, Treatmentid As String
                    If GetNodeValue(oSubNode2, "内容性质", "") = 2 And GetNodeValue(oSubNode2, "诊治要素ID", "") <> "" Then
                        gstrSQL = "select id from 诊治所见项目 where id=[1]"
                        Set rsCount = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetNodeValue(oSubNode2, "诊治要素ID", ""))
                        If rsCount.RecordCount > 0 Then
                            Treatmentid = GetNodeValue(oSubNode2, "诊治要素ID", "")
                        Else
                            Treatmentid = "null"
                        End If
                    Else
                        Treatmentid = "null"
                    End If
                    gstrSQL = "Zl_病历词句组成_Insert(" & lngWordId & _
                                "," & GetNodeValue(oSubNode2, "排列次序", "") & _
                                ",1,'" & GetNodeValue(oSubNode2, "内容文本", "") & _
                                "','" & GetNodeValue(oSubNode2, "要素名称", "") & "'," & _
                                Treatmentid & "," & GetNodeValue(oSubNode2, "替换域", "") & _
                                "," & IIf(GetNodeValue(oSubNode2, "要素类型", "") = "", "Null", GetNodeValue(oSubNode2, "要素类型", "")) & _
                                "," & GetNodeValue(oSubNode2, "要素长度", "") & _
                                "," & GetNodeValue(oSubNode2, "要素小数", "") & _
                                ",'" & GetNodeValue(oSubNode2, "要素单位", "") & _
                                "'," & GetNodeValue(oSubNode2, "要素表示", "") & _
                                ",'" & GetNodeValue(oSubNode2, "要素值域", "") & _
                                "'," & GetNodeValue(oSubNode2, "输入形态", "") & _
                                ",'" & GetNodeValue(oSubNode2, "对象属性", "") & "')"
                    ArraySQL(lngCount) = gstrSQL
                End If
            Next
                
            '后期处理
            lngCount = UBound(ArraySQL) + 1
            ReDim Preserve ArraySQL(1 To lngCount) As String
            gstrSQL = "Zl_病历词句组成_Aftersave(" & lngWordId & ")"
            ArraySQL(lngCount) = gstrSQL
            
                    '设置病历词句条件
            For j = 0 To oSubNode1.selectNodes("Condition").Length - 1
                Set oSubNode3 = oSubNode1.selectNodes("Condition")(j)
                lngCount = UBound(ArraySQL) + 1
                ReDim Preserve ArraySQL(1 To lngCount) As String
                
                gstrSQL = "Zl_病历词句条件_Edit(" & lngWordId & ",'" & GetNodeValue(oSubNode3, "条件项", "") & "','" & GetNodeValue(oSubNode3, "条件值", "") & "')"
                
                ArraySQL(lngCount) = gstrSQL
            Next
            
            '执行保存操作
            Err = 0: On Error GoTo Errhand
            gcnOracle.BeginTrans
            blnTran = True
            For k = 1 To UBound(ArraySQL)
                gstrSQL = ArraySQL(k)
                Call zldatabase.ExecuteProcedure(gstrSQL, "cEPRDocument")
            Next
            gcnOracle.CommitTrans
            blnTran = False
        End If
    Next
    Exit Sub
Errhand:
    If InStr(1, Err.Description, "病历词句分类的(名称、上级ID)出现重复") > 0 Then MsgBox "病历词句分类的(名称、上级ID)出现重复！", vbInformation, gstrSysName
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picNowList_Resize()
    With picNowList
        Me.tvwNowList.Move 0, 0, .Width, .Height - Me.cmdYes.Height - 100
    End With
End Sub
Private Sub tvwClass_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '创建右键菜单
    If Button = 2 And mblnDifference = False And Me.tvwClass.Nodes.Count > 0 Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        Dim cbpPopup As CommandBarPopup
        Dim cbpCustom As CommandBarControlCustom
    
        Set Popup = Me.cbsThis.Add("Popup", xtpBarPopup)
        
        With Popup.Controls
            Set cbpPopup = .Add(xtpControlPopup, 21, "添加到指定分类(&R)")
                cbpPopup.CommandBar.Title = "添加到指定分类"
            Set cbpCustom = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 211, "添加到指定分类列表")
                cbpCustom.Handle = Me.picNowList.hwnd
                cbpPopup.ID = 21
            Set Control = .Add(xtpControlButton, 22, "取消指定分类(&A)")
            Set Control = .Add(xtpControlButton, 23, "导入此分类(&D)")
            Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, 24, "不导入此分类类(&L)")
            Set cbpPopup = .Add(xtpControlPopup, 25, "指定上级分类(&S)")
                cbpPopup.CommandBar.Title = "上级分类"
            cbpPopup.BeginGroup = True
            Set cbpCustom = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 251, "上级分类列表")
                cbpCustom.Handle = Me.picNowList.hwnd
            cbpPopup.ID = 25
            Set Control = .Add(xtpControlButton, 26, "取消上级分类(&S)")
        End With
        Popup.ShowPopup
    End If
   
End Sub

Private Sub tvwClass_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    Call CheckTvw(Node)
    Call CheckParentsNodes(Node)
    If Mid(Node.Key, 2) <> Me.vsfList.Tag Then Exit Sub
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, mCol.Choose) = Node.Checked
            '对mtion进行操作
            On Error Resume Next
            If Node.Checked = False Then
                mtion.Add .TextMatrix(i, mCol.ID), "_" & .TextMatrix(i, mCol.ID)
            Else
                mtion.Remove "_" & .TextMatrix(i, mCol.ID)
            End If
        Next
    End With
End Sub
Private Sub CheckTvw(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    If Node.Children < 1 Then Exit Sub
    With tvwClass
        For i = 1 To .Nodes.Count
            If Not .Nodes(i).Parent Is Nothing Then
                If .Nodes(i).Parent.Key = Node.Key Then
                    .Nodes(i).Checked = Node.Checked
                    Call CheckTvw(.Nodes(i))
                End If
            End If
        Next
    End With
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    If mblnDifference Then
        Call zlSubRefList(Mid(Node.Key, 2), Node.Checked)
    Else
        Call xmlSubRefList(Mid(Node.Key, 2), Node.Checked, Node.Text)
    End If
    Me.vsfList.Tag = Mid(Node.Key, 2) '标记此时列表数据属于那个树节点
    If Me.vsfList.Rows > 1 Then
        Me.stbThis.Panels(2).Text = "该分类下有" & Me.vsfList.Rows - 1 & "条词句"
    Else
        Me.stbThis.Panels(2).Text = ""
    End If
End Sub
Private Sub CheckParentsNodes(ByVal oNode As Node)
 Do While (Not oNode.Parent Is Nothing)
    oNode.Parent.Checked = True
    Set oNode = oNode.Parent
 Loop
End Sub
Private Sub xmlSubRefList(lngFileID As Long, blnCheck As Boolean, strClassName As String)
    Dim strPath As String

    If oRoot Is Nothing Then Exit Sub
    strPath = "Class[ID=" & lngFileID & "]"
    Call xmlSenceList(oRoot.getElementsByTagName(strPath)(0), lngFileID, blnCheck, strClassName)
End Sub
Private Sub xmlSenceList(Node As IXMLDOMNode, lngFileID As Long, blnCheck As Boolean, strClassName As String)
    Dim oSubNode2 As IXMLDOMNode        '子节点
    Dim i As Long, j As Long
    Dim NodeList As IXMLDOMNodeList
    Dim k As Long

        With vsfList
            .ScrollBars = flexScrollBarNone '为防止加载数据的时候滚动条抖动，先不显示滚动条
            .Rows = 1
            For j = 0 To Node.selectNodes("Sentence").Length - 1
                Set oSubNode2 = Node.selectNodes("Sentence")(j)
                .AddItem ""
                k = k + 1
                Select Case GetNodeValue(oSubNode2, "通用级", "")
                    Case 0:
                        .TextMatrix(k, 5) = ""
                        .Cell(flexcpPicture, k, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-全院"
                    Case 1:
                        .TextMatrix(k, 5) = ""
                        .Cell(flexcpPicture, k, mCol.Range) = Me.imgList.ListImages(2).Picture '"2-科室"
                    Case 2:
                        .TextMatrix(k, 5) = ""
                        .Cell(flexcpPicture, k, mCol.Range) = Me.imgList.ListImages(3).Picture '"3-个人"
                    Case Else:
                        .TextMatrix(k, 5) = ""
                        .Cell(flexcpPicture, k, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-全院"
                End Select
                .Cell(flexcpAlignment, k, mCol.Range) = flexAlignCenterCenter
                .TextMatrix(k, mCol.ID) = GetNodeValue(oSubNode2, "ID", "")
                .TextMatrix(k, mCol.clasID) = GetNodeValue(oSubNode2, "分类id", "")
                .TextMatrix(k, mCol.Class) = GetNodeValue(oSubNode2, "分类", "")
                .TextMatrix(k, mCol.Num) = GetNodeValue(oSubNode2, "编号", "")
                .TextMatrix(k, mCol.pName) = GetNodeValue(oSubNode2, "名称", "")
                If blnCheck Then
                    On Error Resume Next '在此处如果发生异常表示在mtion中没有记录
                    If mtion("_" & GetNodeValue(oSubNode2, "ID", "")) <> GetNodeValue(oSubNode2, "ID", "") Then
                        .TextMatrix(k, mCol.Choose) = 1
                    Else
                        .TextMatrix(k, mCol.Choose) = 0
                    End If
                End If
                '获取病历内容
                Dim oSubNode3 As IXMLDOMNode        '子节点
                Dim M As Long, lngStart As Long, strText As String
                For M = 0 To oSubNode2.selectNodes("Content").Length - 1
                    Set oSubNode3 = oSubNode2.selectNodes("Content")(M)
                    lngStart = Len(Me.rtbText.Text)
                    Me.rtbText.SelStart = lngStart
                    Me.rtbText.SelLength = 0
                    Select Case GetNodeValue(oSubNode3, "内容性质", "")
                    Case 0 '自由文字
                        strText = GetNodeValue(oSubNode3, "内容文本", "")
                        With Me.rtbText
                            .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                            .SelUnderline = False
                        End With
                    Case 1, 2 '1-临时诊治要素,2-固定诊治要素
                        If GetNodeValue(oSubNode3, "内容文本", "") = "" Then
                            strText = "{" & GetNodeValue(oSubNode3, "要素名称", "") & "}" & GetNodeValue(oSubNode3, "要素单位", "")
                        Else
                            strText = "{" & GetNodeValue(oSubNode3, "内容文本", "") & "}"
                        End If
                        With Me.rtbText
                            .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                            .SelUnderline = True
                        End With
                    End Select
                Next
                Me.rtbText.SelStart = 0
                .Cell(flexcpData, k, mCol.Range) = Me.rtbText.Text
                Me.rtbText.Text = ""
                
                Dim rsCount As ADODB.Recordset
                gstrSQL = "select a.id from 病历词句示范 a,病历词句分类 b where a.分类id=b.id and a.名称=[1] and b.名称=[2]"
                Set rsCount = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetNodeValue(oSubNode2, "名称", ""), (Split(strClassName, "-")(1)))
                '如果在对应分类下不存着该词句，则红色显示
                If rsCount.RecordCount < 1 Then
                    .Cell(flexcpForeColor, k, mCol.Range, k, mCol.Class) = RGB(255, 20, 147)
                End If
            Next
            .ScrollBars = flexScrollBarVertical
        End With
        Exit Sub
End Sub
Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Integer
        With vsfList
            If .Rows <= 1 Then
                Exit Sub
            End If
            If (.ColWidth(mCol.Range) < x And x < .ColWidth(mCol.Range) + .ColWidth(mCol.Choose)) And Button = 1 Then
                '当选者词句的时候，同时把未选中的词句记录到mtion中
                If .TextMatrix(.Row, mCol.Choose) = "1" Then
                    .TextMatrix(.Row, mCol.Choose) = "0"
                    mtion.Add .TextMatrix(.Row, mCol.ID), "_" & .TextMatrix(.Row, mCol.ID)
                Else
                    .TextMatrix(.Row, mCol.Choose) = "1"
                    For i = 1 To tvwClass.Nodes.Count
                        If Mid(tvwClass.Nodes(i).Key, 2) = .TextMatrix(.Row, mCol.clasID) Then
                            If tvwClass.Nodes(i).Checked = False Then tvwClass.Nodes(i).Checked = True
                            Exit For
                        End If
                    Next
                    On Error GoTo Errhand
                    mtion.Remove "_" & .TextMatrix(.Row, mCol.ID)
                End If
            End If
            If Button = 1 Then
                If mblnDifference Then
                    If .TextMatrix(.Row, mCol.ID) <> "" Then Call zlSubRefText(.TextMatrix(.Row, mCol.ID))
                Else
                    Me.rtbText.Text = Me.vsfList.Cell(flexcpData, vsfList.Row, mCol.Range)
                End If
            End If
        End With
        Exit Sub
Errhand:
        For i = 1 To vsfList.Rows - 1
            If vsfList.TextMatrix(i, mCol.Choose) = "" Then
                mtion.Add vsfList.TextMatrix(i, mCol.ID), "_" & vsfList.TextMatrix(i, mCol.ID)
            Else
               ' mtion.Remove "_" & vsfList.TextMatrix(i, mCol.ID)
            End If
        Next
End Sub
'################################################################################################################
'## 功能：  创建一个XML节点并赋值
'##
'## 参数：  TabNumber   :   缩进层次数（表示有多少个Tab制表符，便于阅读）
'##         Parent      :   父节点
'##         Node_Type   :   节点类型（目前支持 NODE_ELEMENT 、NODE_CDATA_SECTION 、NODE_COMMENT 、NODE_ATTRIBUTE等）
'##         Node_Name   :   节点名称
'##         Node_Value  :   节点文本
'################################################################################################################
Private Function CreateNode(ByVal TabNumber As Integer, _
    ByVal Parent As IXMLDOMNode, _
    Optional ByVal node_name As String, _
    Optional ByVal Node_Type As tagDOMNodeType = NODE_ELEMENT, _
    Optional ByVal Node_Value As String = "")
    Dim New_Node As IXMLDOMNode
    
    '字符缩进值设置（不影响数据），只影响阅读美观度
    Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf & String(TabNumber, vbKeyTab))   '创建文本节点
    '创建新节点
    Set New_Node = Parent.ownerDocument.CreateNode(Node_Type, node_name, "")
    '设置文本值
    New_Node.Text = Node_Value
    '添加到父节点
    Parent.appendChild New_Node
    '添加末尾回车（不影响数据），只影响阅读美观度
    'Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf)   '创建文本节点
    Set CreateNode = New_Node
End Function

'################################################################################################################
'## 功能：  获取一个节点的值
'##
'## 参数：  CurNode         :   当前节点对象
'##         SubNodeName     :   子节点名称
'##         DefaultValue    :   默认值
'################################################################################################################
Private Function GetNodeValue(ByVal CurNode As IXMLDOMNode, _
    ByVal SubNodeName As String, _
    Optional ByVal DefaultValue As String = "") As String
    
    On Error Resume Next
    Dim NodeTMP As IXMLDOMNode
    Set NodeTMP = CurNode.selectSingleNode(".//" & SubNodeName)
    If NodeTMP Is Nothing Then
        GetNodeValue = DefaultValue
    Else
        GetNodeValue = NodeTMP.Text
    End If
    
    If InStr(GetNodeValue, vbCr) > 0 And InStr(GetNodeValue, vbLf) = 0 Then '只有回车符无换行符
        GetNodeValue = Replace(GetNodeValue, vbCr, vbCrLf)
    ElseIf InStr(GetNodeValue, vbLf) > 0 And InStr(GetNodeValue, vbCr) = 0 Then '只有换行符无回车符
        GetNodeValue = Replace(GetNodeValue, vbLf, vbCrLf)
    End If
End Function

'################################################################################################################
'## 功能：  全选/全清
'################################################################################################################
Private Function CheckAllOrClearAll(ByVal blnOn As Boolean)
    Dim oNode As Node, i As Long
    For Each oNode In Me.tvwClass.Nodes
        oNode.Checked = blnOn
        Call CheckTvw(oNode)
    Next
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, mCol.Choose) = blnOn
            '对mtion进行操作
            On Error Resume Next
            If Not blnOn Then
                mtion.Add .TextMatrix(i, mCol.ID), "_" & .TextMatrix(i, mCol.ID)
            Else
                mtion.Remove "_" & .TextMatrix(i, mCol.ID)
            End If
        Next
    End With
End Function

