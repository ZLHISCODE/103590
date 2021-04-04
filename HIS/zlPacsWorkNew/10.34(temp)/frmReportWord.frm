VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReportWord 
   Caption         =   "词句示范"
   ClientHeight    =   8100
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9270
   Icon            =   "frmReportWord.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   9270
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picCommandButton 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   1800
      ScaleHeight     =   1935
      ScaleWidth      =   4935
      TabIndex        =   10
      Top             =   6000
      Width           =   4935
      Begin VB.CommandButton cmdSure 
         Caption         =   "确 定(&S)"
         Height          =   350
         Left            =   2040
         TabIndex        =   13
         Top             =   1560
         Width           =   1100
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "取 消(&C)"
         Height          =   350
         Left            =   3600
         TabIndex        =   11
         Top             =   1560
         Width           =   1100
      End
      Begin RichTextLib.RichTextBox rtbEditWord 
         Height          =   975
         Left            =   360
         TabIndex        =   12
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1720
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmReportWord.frx":0CCA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picPrivateWord 
      Height          =   1575
      Left            =   4920
      ScaleHeight     =   1515
      ScaleWidth      =   3675
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   3735
      Begin RichTextLib.RichTextBox rtxtPrivateWord 
         Height          =   975
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "双击进入编辑状态，再双击保存修改"
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1720
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmReportWord.frx":0D67
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picWordTree 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   3795
      TabIndex        =   5
      Top             =   0
      Width           =   3855
      Begin VB.CheckBox ChkAutoExpand 
         Caption         =   "自动展开"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chk条件过滤 
         Caption         =   "按限制条件过滤"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin MSComctlLib.TreeView trvWordTree 
         Height          =   1935
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3413
         _Version        =   393217
         Indentation     =   176
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CheckBox chk直接编辑 
         Caption         =   "直接编辑"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picWordShow 
      AutoSize        =   -1  'True
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3075
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   2760
      Width           =   4455
      Begin VB.PictureBox picWordContainer 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   3495
         TabIndex        =   2
         Top             =   240
         Width           =   3495
         Begin VB.CommandButton cmdSelect 
            Height          =   375
            Index           =   0
            Left            =   0
            Picture         =   "frmReportWord.frx":0DFF
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "写入报告"
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin RichTextLib.RichTextBox rtxtWord 
            Height          =   975
            Index           =   0
            Left            =   480
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1720
            _Version        =   393217
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmReportWord.frx":1389
         End
      End
      Begin VB.VScrollBar vscroWordH 
         Height          =   1215
         Left            =   3720
         Max             =   500
         TabIndex        =   1
         Top             =   1440
         Value           =   200
         Width           =   250
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   1200
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
            Picture         =   "frmReportWord.frx":1426
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportWord.frx":1B20
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   5040
      Top             =   480
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Menu menuPopup 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu menuAddWord 
         Caption         =   "新增词句"
      End
      Begin VB.Menu menuModifyWord 
         Caption         =   "修改词句"
      End
      Begin VB.Menu menuSplit 
         Caption         =   "-"
      End
      Begin VB.Menu menuSaveAllWord 
         Caption         =   "全套存入"
      End
   End
End
Attribute VB_Name = "frmReportWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mpreWinProc As Long
Private mFileID As Long                 '报告ID
Private mstrReportViewType As String    '词句示范内容类型， “检查所见”“诊断意见”，“建议”等
Private mlngAdviceID As Long            '医嘱ID
Private mdbOwner As String              '数据库所有者
Private mlngDeptID As Long              '科室ID
Private mblnSingleWindow As Boolean     '是否使用独立窗口显示报告编辑器，True-独立窗口显示；False-嵌入式显示
Public mblnShowWord As Boolean         '显示词句示范，True--显示词句示范；False--双击标题才显示词句示范
Private mlngModul As Long               '模块号
Private mintWordDblClick As Integer     '词句双击后的操作：0--直接写入报告；1--打开词句编辑窗口
Private mintWordPower As Integer        '词句管理权范围
Private mstrReportViewTypeAlias As String

Private mblnEditable As Boolean         '是否可以编辑报告

Private mlngWordTreeH As Long               '词库模板树的高度
Private mlngWordShowH As Long               '词库模板内容的高度
Private mlngPrivateWordH As Long            '私人常用词句的高度
Private mlngButtonH As Long                 '确定和取消按钮区域的高度

Private mPrivatePane As Pane                '私人常用词句区域的页面
Private mblnInitFaseScheme As Boolean       '初始化界面，只执行一次

Private miWordScale As Integer

'本窗体的事件
Public Event WordSelected(strWord As String, strReportViewType As String, blnIsPopupWindInsert As Boolean)   '词句被选择
Public Event AddSampleWord(ByVal blnIsAllWord As Boolean)     '新增词句示范
Public Event ModifySampleWord() '修改词句示范


Private Sub ChkAutoExpand_Click()
    Call LoadWordTree(mFileID, mstrReportViewType, True)
End Sub

Private Sub chk条件过滤_Click()
    Call LoadWordTree(mFileID, mstrReportViewType, True)
End Sub

Private Sub chk直接编辑_Click()
    If CBool(chk直接编辑.value) Then
        mlngButtonH = Round(Me.Height / 4)
    End If
    
    Call InitFaceScheme
    
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    If Not mblnShowWord And CBool(chk直接编辑.value) Then
        rtbEditWord.SelLength = 0
        rtbEditWord.SelText = rtxtWord(Index).Text & vbCrLf
    Else
        '写入报告填空框
        RaiseEvent WordSelected(rtxtWord(Index).Text, rtxtWord(Index).Tag, False)
    End If
End Sub

Private Sub cmdSure_Click()
    '写入报告填空框
    RaiseEvent WordSelected(rtbEditWord.Text, mstrReportViewType, True)
    
    Unload Me
End Sub

Private Sub Form_Load()

    mdbOwner = GetDbOwner(glngSys)
    
    trvWordTree.ImageList = ImageList1
    chk条件过滤.value = 1
    mstrReportViewType = ""
    mblnInitFaseScheme = False
    
    ''''''''''''''''''''''''''处理鼠标滚轮'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If App.LogMode <> 0 Then
        Dim ret As Long
        Set mReport.fReport = Me
    '    '记录原来的window程序地址
        preWinProc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
    '    '用自定义程序代替原来的window程序
        ret = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Wndproc)
    End If
End Sub


Private Sub InitFaceScheme()
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane
    With Me.dkpMain
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 0, mlngWordTreeH, DockTopOf, Nothing)
    Pane1.Title = "词句示范"
    Pane1.Handle = picWordTree.hWnd
    Pane1.Options = PaneNoCaption Or PaneNoCloseable
    
    Set Pane2 = dkpMain.CreatePane(2, 0, mlngWordShowH, DockBottomOf, Pane1)
    Pane2.Title = "词句内容"
    Pane2.Handle = picWordShow.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable
    
    Set Pane3 = dkpMain.CreatePane(3, 0, mlngPrivateWordH, DockBottomOf, Pane2)
    Pane3.Title = "常用词句"
    Pane3.Handle = picPrivateWord.hWnd
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set mPrivatePane = Pane3
    
    chk直接编辑.Visible = Not mblnShowWord
    picCommandButton.Visible = Not mblnShowWord

    
    If Not mblnShowWord Then    '通过双击打开，则显示确定和取消按钮
        
        If Not CBool(chk直接编辑.value) Then
            Set pane4 = dkpMain.CreatePane(4, 0, cmdClose.Height + 50, DockBottomOf, Pane3)
            pane4.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        Else
            Set pane4 = dkpMain.CreatePane(4, 0, mlngButtonH, DockBottomOf, Pane3)
            pane4.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        End If
        
        
        pane4.Title = mstrReportViewTypeAlias
        pane4.Handle = picCommandButton.hWnd
        
        
        cmdSure.Visible = CBool(chk直接编辑.value)
        rtbEditWord.Visible = CBool(chk直接编辑.value)
    End If
End Sub


Private Sub LoadWordTree(FileID As Long, strReportViewType As String, Optional blnForceRefresh As Boolean = False)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strTextName As String
    Dim lng提纲ID As Long, lng病人ID As Long, lng主页ID As Long
    Dim objNode As Node
    Dim objPnode As Node
    Dim strKey As String
    
    If strReportViewType = mstrReportViewType And FileID = mFileID And blnForceRefresh = False Then Exit Sub
    
    mstrReportViewType = strReportViewType
    mFileID = FileID
    
    '清空模板内容
    Call ClearWordShow
    '调用引用API，并且采用逆序循环删除TreeView的方法，这个方法速度更快
    Call TrvwClear
    
    strTextName = mstrReportViewType
    
    '打开对应的词句示范，检查所见对应的词句示范
    strSql = "Select nvl(C.父id,0) as 提纲ID ,a.病人ID,nvl(a.主页id ,0) as 主页id  From 病历文件结构 C ,病人医嘱记录 a " & _
             " Where C.文件ID=[1] and C.内容文本=[2] And C.对象类型=3 And Rownum =1 And a.Id=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, FileID, strTextName, mlngAdviceID)
    If rsTemp.EOF = True Then Exit Sub
    lng提纲ID = rsTemp!提纲ID
    lng病人ID = rsTemp!病人ID
    lng主页ID = rsTemp!主页ID
        
    If chk条件过滤.value = 0 Then   '不按照部位过滤
        strSql = "Select Distinct Id,上级id,编码,名称 From 病历词句分类 Start With Id In (" & _
             "Select L.分类id " & vbNewLine & _
            " From 病历词句分类 C, 病历词句示范 L, 病历提纲词句 A, 部门表 D, 人员表 P " & vbNewLine & _
            " Where C.ID = L.分类id And L.分类id = A.词句分类id And L.科室id = D.ID And L.人员id = P.ID And A.提纲id = [1] And" & vbNewLine & _
            "      (L.通用级=0 Or (L.通用级=1 And L.科室ID=[2]) Or (L.通用级=2 And L.人员ID= [3]))" & vbNewLine & _
            " ) Connect By Prior 上级id=Id  Order By 编码"
    Else                            '按照部位过滤
        strSql = "Select Distinct Id,上级id,编码,名称 From 病历词句分类 Start With Id In (" & _
             "Select /*+RULE*/ Distinct L.分类id " & vbNewLine & _
            "From 病历词句分类 C, 病历词句示范 L, 病历提纲词句 A, 部门表 D, 人员表 P," & vbNewLine & _
            "     Table(Cast(f_Sentence_Usable([1], [4], [5], [6]) As " & mdbOwner & ".t_Dic_Rowset)) U" & vbNewLine & _
            "Where C.ID = L.分类id And L.分类id = A.词句分类id And L.科室id = D.ID And L.人员id = P.ID And A.提纲id = [1] And" & vbNewLine & _
            "      L.ID = To_Number(U.编码) And (L.通用级=0 Or (L.通用级=1 And L.科室ID=[2]) Or (L.通用级=2 And L.人员ID= [3]))" & vbNewLine & _
            " ) Connect By Prior 上级id=Id  Order By 编码"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng提纲ID, mlngDeptID, UserInfo.ID, lng病人ID, lng主页ID, mlngAdviceID)
    
    Do While Not rsTemp.EOF
        
        Set objNode = Nothing
        
        On Error Resume Next
        Set objNode = trvWordTree.Nodes("T-" & rsTemp("ID").value)
        If zlCommFun.Nvl(rsTemp("上级id").value, 0) <> 0 Then
            Set objPnode = trvWordTree.Nodes("T-" & rsTemp("上级id").value)
        Else
            Set objPnode = Nothing
        End If
        On Error GoTo errHandle
        
        If objNode Is Nothing Then
            If objPnode Is Nothing Then
                Set objNode = trvWordTree.Nodes.Add(, , "T-" & rsTemp("ID").value, rsTemp("名称").value, 2)
                '根据CheckBox判断是否自动加载
                If ChkAutoExpand.value = 0 Then
                    objNode.Expanded = True
                    
                If Not objNode.Parent Is Nothing Then
                    If InStr(strKey, objNode.Parent.Key) <= 0 Then
                        strKey = strKey & "," & objNode.Parent.Key
                        '装载叶子节点
                        Call LoadClassWork(objNode.Parent)
                    End If
                End If
                End If
            Else
                Set objNode = trvWordTree.Nodes.Add("T-" & zlCommFun.Nvl(rsTemp("上级id").value, 0), tvwChild, "T-" & rsTemp("ID").value, rsTemp("名称").value, 2)
            End If
            objNode.Tag = lng提纲ID & "-" & lng病人ID & "-" & lng主页ID & "-" & mlngAdviceID
            '根据CheckBox判断是否自动加载
            If ChkAutoExpand.value = 1 Then
                objNode.Expanded = True
                
                If Not objNode.Parent Is Nothing Then
                    If InStr(strKey, objNode.Parent.Key) <= 0 Then
                        strKey = strKey & "," & objNode.Parent.Key
                        '装载叶子节点
                        Call LoadClassWork(objNode.Parent)
                    End If
                End If
            End If
            
        End If
        rsTemp.MoveNext
    Loop
    
    Exit Sub
errHandle:
    If err.Number <> 35602 Then
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReportWord\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReportWord"
    End If
    
    '保存词句示范区域的高度
    '285是Pane的标题高度，使用了标题，就需要加回这个高度
    SaveSetting "ZLSOFT", strRegPath, "WordTreeH", picWordTree.Height
    SaveSetting "ZLSOFT", strRegPath, "WordShowH", picWordShow.Height
    SaveSetting "ZLSOFT", strRegPath, "PrivateWordH", picPrivateWord.Height ' + 285
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmReportWord", "直接编辑", CLng(chk直接编辑.value)
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmReportWord", "自动展开", CLng(ChkAutoExpand.value)
    
    If mblnShowWord = False Then    '通过双击打开，则显示确定和取消按钮,记录这个高度
        SaveSetting "ZLSOFT", strRegPath, "ButtonH", picCommandButton.Height
    End If
    
    '保存词句示范区域的宽度
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
    End If
    SaveSetting "ZLSOFT", strRegPath, "CX1", picWordTree.Width
    
    '窗口模式,此模式下记录窗口位置
    If mblnShowWord = False Then
        Call SaveWinState(Me, App.ProductName)
    End If
End Sub

Private Sub menuAddWord_Click()
    RaiseEvent AddSampleWord(False)
End Sub

Private Sub menuModifyWord_Click()
    RaiseEvent ModifySampleWord
End Sub

Private Sub menuSaveAllWord_Click()
    RaiseEvent AddSampleWord(True)
End Sub

Private Sub picCommandButton_Resize()
    On Error Resume Next
    
    If mblnShowWord = False Then
        rtbEditWord.Left = 0
        rtbEditWord.Top = 0
        
        If CBool(chk直接编辑.value) Then
            rtbEditWord.Width = picCommandButton.ScaleWidth
            rtbEditWord.Height = picCommandButton.ScaleHeight - cmdSure.Height - 100
        Else
            rtbEditWord.Width = 0
            rtbEditWord.Height = 0
        End If
        
        cmdClose.Left = picCommandButton.ScaleWidth - cmdClose.Width - 200
        cmdSure.Left = cmdClose.Left - cmdSure.Width - 200
        
        cmdClose.Top = picCommandButton.ScaleHeight - cmdClose.Height - 50
        cmdSure.Top = picCommandButton.ScaleHeight - cmdSure.Height - 50
    End If
End Sub

Private Sub picPrivateWord_Resize()
On Error Resume Next

    rtxtPrivateWord.Left = 0
    rtxtPrivateWord.Top = 0
    rtxtPrivateWord.Width = picPrivateWord.ScaleWidth
    rtxtPrivateWord.Height = picPrivateWord.ScaleHeight
End Sub

Private Sub picWordShow_Resize()
    Dim i As Integer
    
    On Error Resume Next
    
    '调整每一个RichTextBox的宽度
    For i = 1 To rtxtWord.Count - 1
        rtxtWord(i).Width = Abs(picWordContainer.Width - rtxtWord(i).Left - 60)
    Next i
    
    Call ResizeWordContainer
End Sub

Private Sub picWordTree_Resize()
On Error Resume Next
    
    chk条件过滤.Left = 10
    chk条件过滤.Top = 10
    
    chk直接编辑.Left = chk条件过滤.Left + chk条件过滤.Width + 20
    chk直接编辑.Top = chk条件过滤.Top
    
    ChkAutoExpand.Left = IIf(chk直接编辑.Visible, chk直接编辑.Left + chk直接编辑.Width + 80, chk条件过滤.Left + chk条件过滤.Width + 20)
    ChkAutoExpand.Top = chk条件过滤.Top
    
    trvWordTree.Left = 0
    trvWordTree.Top = chk条件过滤.Top + chk条件过滤.Height
    trvWordTree.Width = picWordTree.Width
    trvWordTree.Height = Abs(picWordTree.Height - 10 - chk条件过滤.Top - chk条件过滤.Height)
End Sub

Private Sub rtxtPrivateWord_DblClick()
    rtxtPrivateWord.Locked = Not rtxtPrivateWord.Locked
    If rtxtPrivateWord.Locked = True Then
        '先判断词句内容的长度是否超过1000个字符，如果超过，则提醒
        If Len(rtxtPrivateWord.Text) > 1000 Then
            MsgBoxD Me, "私人词句的长度不能超过 1000个字符，请重新修改后再保存。"
            mPrivatePane.Title = "词语编辑模式，双击保存"
            rtxtPrivateWord.Locked = False
            Exit Sub
        End If
        rtxtPrivateWord.BackColor = vbWhite
        mPrivatePane.Title = "常用词语"
        Call zlDatabase.SetPara("报告常用词句", rtxtPrivateWord.Text, glngSys, mlngModul)
    Else
        mPrivatePane.Title = "词语编辑模式，双击保存"
        rtxtPrivateWord.BackColor = &H80000013
    End If
End Sub

Private Sub rtxtWord_DblClick(Index As Integer)
    Call richTextBoxShowElements(rtxtWord(Index))
End Sub

Private Sub trvWordTree_DblClick()
    Dim i As Integer
    
    If Not trvWordTree.SelectedItem Is Nothing Then
        If Left(trvWordTree.SelectedItem.Key, 1) = "L" Then
        
        If mblnEditable Then
            If mintWordDblClick = 1 And (mstrReportViewType = ReportViewType_检查所见 _
                Or mstrReportViewType = ReportViewType_诊断意见 Or mstrReportViewType = ReportViewType_建议) Then              '词句双击后直接写入报告
                '词句双击后，打开词句编辑窗口
                WriteWordEdit Right(trvWordTree.SelectedItem.Key, Len(trvWordTree.SelectedItem.Key) - 2)
            Else
                For i = 1 To cmdSelect.Count - 1
                    cmdSelect_Click i
                Next i
            End If
        End If
            
        
        ElseIf Left(trvWordTree.SelectedItem.Key, 1) = "T" And trvWordTree.SelectedItem.Checked = False Then
            '装载叶子节点
            Call LoadClassWork(trvWordTree.SelectedItem)
        End If
    End If
End Sub

Private Sub LoadClassWork(ByVal objNode As Object)
    '载入分类下的词句
    Dim strSql As String
    Dim strPara() As String
    Dim rsLeaf As ADODB.Recordset
    Dim objCurNode As Node
    Dim objSubNode As Node
    Dim lngClassID As Long
    
    If objNode Is Nothing Then Exit Sub
    
    Set objCurNode = objNode
    
    If objCurNode.Tag = "" Then Exit Sub
    
    lngClassID = Right(objNode.Key, Len(objNode.Key) - 2)
    
    '装载叶子节点
    objNode.Checked = True
    
    strPara = Split(objNode.Tag, "-")
    If chk条件过滤.value = 0 Then       '不按照条件过滤
        strSql = "Select  L.Id as 示范ID,L.名称 as 示范名称 From 病历词句示范 L " & _
            " Where L.分类id=[7] and (L.通用级=0 Or (L.通用级=1 And L.科室ID=[1]) Or (L.通用级=2 And L.人员ID= [2]))" & _
            " Order By 编号"
    Else                                '按照条件过滤
        strSql = "Select /*+RULE*/ L.Id as 示范ID,L.名称 as 示范名称 From 病历词句示范 L, Table(Cast(f_Sentence_Usable([3], [4], [5], [6]) As " & mdbOwner & ".t_Dic_Rowset)) U " & _
            " Where L.分类id=[7] and L.ID = To_Number(U.编码) And (L.通用级=0 Or (L.通用级=1 And L.科室ID=[1]) Or (L.通用级=2 And L.人员ID= [2]))" & _
            " Order By 编码"
    End If
    Set rsLeaf = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID, UserInfo.ID, strPara(0), strPara(1), strPara(2), strPara(3), lngClassID)
    
    Do While Not rsLeaf.EOF
        Set objSubNode = trvWordTree.Nodes.Add(objNode, tvwChild, "L-" & rsLeaf("示范ID").value, rsLeaf("示范名称").value, 1)
        rsLeaf.MoveNext
    Loop
    
    objNode.Expanded = True
End Sub

Private Sub trvWordTree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '处理右键弹出菜单，判断是否右键
    If Button = vbRightButton And mintWordPower <> -1 Then
        If trvWordTree.SelectedItem Is trvWordTree.HitTest(X, Y) And Not trvWordTree.SelectedItem Is Nothing Then
            If Left(trvWordTree.SelectedItem.Key, 1) = "L" Then '叶子节点，可以修改词句示范
                menuModifyWord.Visible = True
            Else    '分类结点，不能修改词句示范
                menuModifyWord.Visible = False
            End If
            PopupMenu menuPopup
        End If
    End If
End Sub

Private Sub trvWordTree_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngWordID As Long
    Dim blnNextLine As Boolean
    Dim iFieldCount As Integer
    
    Dim blnStartSegment As Boolean      '开始一个段落
    Dim str内容文本 As String
    Dim str插入文本 As String
    
    '清空原有控件
    Call ClearWordShow
    blnNextLine = True
    miWordScale = 0
    
    If Left(Node.Key, 1) = "L" Then
        lngWordID = Right(Node.Key, Len(Node.Key) - 2)
        strSql = "Select 词句id,排列次序,内容性质,内容文本,诊治要素ID,替换域,要素名称,要素类型,要素长度,要素小数," & _
                 " 要素单位,要素表示,要素值域,输入形态 From 病历词句组成 Where 词句ID=[1] order by 排列次序 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngWordID)
        blnStartSegment = False
        
        On Error GoTo errHandle
        '从数据库中读取词句后，逐行分析并显示
        While rsTemp.EOF = False
            '先把记录中的词句内容读取到str内容文本中
            str内容文本 = Nvl(rsTemp!内容文本)
            
            If blnNextLine = True And Not (rsTemp!内容性质 = 0 And (Trim(str内容文本) = "" Or Trim(str内容文本) = vbCrLf)) Then      '用新的一行来显示，则创建一套cmdSelect和rtxtWord显示内容文本
                blnNextLine = False
                iFieldCount = rtxtWord.Count
                '创建按钮和文本框
                Load rtxtWord(iFieldCount)
                rtxtWord(iFieldCount).Visible = True
                Load cmdSelect(iFieldCount)
                cmdSelect(iFieldCount).Visible = True
                
                If Not mblnEditable Then
                    rtxtWord(iFieldCount).Enabled = False
                    cmdSelect(iFieldCount).Enabled = False
                End If
                

                '先读取内容文本，判断当前内容的类型，如果带了标记，则按照标记来记录类型，否则就使用当前的默认类型
                If Left(str内容文本, 6) = "<<所见>>" Then
                    rtxtWord(iFieldCount).Tag = ReportViewType_检查所见
                ElseIf Left(str内容文本, 6) = "<<诊断>>" Then
                    rtxtWord(iFieldCount).Tag = ReportViewType_诊断意见
                ElseIf Left(str内容文本, 6) = "<<建议>>" Then
                    rtxtWord(iFieldCount).Tag = ReportViewType_建议
                Else
                    rtxtWord(iFieldCount).Tag = mstrReportViewType
                End If
                
                '摆放位置
                If iFieldCount = 1 Then
                    cmdSelect(iFieldCount).Top = 5
                Else
                    cmdSelect(iFieldCount).Top = rtxtWord(iFieldCount - 1).Top + rtxtWord(iFieldCount - 1).Height + 5
                End If
                cmdSelect(iFieldCount).Left = 5
                rtxtWord(iFieldCount).Left = cmdSelect(iFieldCount).Left + cmdSelect(iFieldCount).Width + 10
                rtxtWord(iFieldCount).Top = cmdSelect(iFieldCount).Top
                rtxtWord(iFieldCount).Width = picWordContainer.Width - rtxtWord(iFieldCount).Left - 60
                rtxtWord(iFieldCount).Height = 400
            End If
            
            If rsTemp!内容性质 = 0 Then     '是自由文本，直接加入内容
                If Trim(str内容文本) <> "" And Trim(str内容文本) <> vbCrLf Then '内容文本不为空或者空回车，则解析并显示内容文本
                    '准备插入文字，设置光标位置
                    rtxtWord(iFieldCount).SelStart = Len(rtxtWord(iFieldCount).Text)
                    rtxtWord(iFieldCount).SelLength = 0
                    rtxtWord(iFieldCount).SelColor = vbBlack
                    '如果文字串前面有报告填写位置标识，删除该标识
                    If Left(str内容文本, 6) = "<<所见>>" Or Left(str内容文本, 6) = "<<诊断>>" _
                        Or Left(str内容文本, 6) = "<<建议>>" Then
                        str插入文本 = Right(str内容文本, Len(str内容文本) - 6)
                    ElseIf UCase(Left(str内容文本, 3)) = "<P>" Then
                        '判断是否被<P>和</P>包围了一个独立的段
                        If UCase(Right(str内容文本, 4)) = "</P>" Then
                            str插入文本 = Mid(str内容文本, 4, Len(str内容文本) - 7)
                        ElseIf UCase(Right(str内容文本, 6)) = "</P>" & vbCrLf Then
                            str插入文本 = Mid(str内容文本, 4, Len(str内容文本) - 9)
                        Else
                            str插入文本 = Right(str内容文本, Len(str内容文本) - 3)
                        End If
                        blnStartSegment = True
                    ElseIf UCase(Right(str内容文本, 4)) = "</P>" Then
                        str插入文本 = Left(str内容文本, Len(str内容文本) - 4)
                    ElseIf UCase(Right(str内容文本, 6)) = "</P>" & vbCrLf Then
                        str插入文本 = Left(str内容文本, Len(str内容文本) - 6)
                    Else
                        str插入文本 = str内容文本
                    End If
                    
                    '把内容文本添加到文本框
                    '删除文本末尾的回车换行，如果是<P></P>封装的段落组合，则不删除回车换行
                    If Right(str插入文本, 2) = vbCrLf And blnStartSegment = False Then
                        str插入文本 = Left(str插入文本, Len(str插入文本) - 2)
                    End If
                    rtxtWord(iFieldCount).SelText = str插入文本
                    '判断是否需要换行
                    If blnStartSegment = True Then      '已经启用段落标记，则查找结束段落的标记</P>
                        If UCase(Right(str内容文本, 4)) = "</P>" Or UCase(Right(str内容文本, 6)) = "</P>" & vbCrLf Then
                            blnNextLine = True
                            blnStartSegment = False
                        End If
                    Else    '查找回车作为段落结束标记
                        If Right(str内容文本, 2) = vbCrLf Then
                            blnNextLine = True
                        End If
                    End If
                End If
            Else        'rsTemp!内容性质<>0 ,是要素，需要解析
                If rsTemp!要素表示 = 0 Then     '文本要素解析成空“ ”
                    rtxtWord(iFieldCount).SelStart = Len(rtxtWord(iFieldCount).Text)
                    rtxtWord(iFieldCount).SelLength = 0
                    rtxtWord(iFieldCount).SelText = "  " & Nvl(rsTemp!要素单位)
                    
                    rtxtWord(iFieldCount).SelStart = Len(rtxtWord(iFieldCount).Text) - Len(Nvl(rsTemp!要素单位))
                    rtxtWord(iFieldCount).SelLength = Len("  " & Nvl(rsTemp!要素单位))
                    rtxtWord(iFieldCount).SelColor = vbBlue
                ElseIf rsTemp!要素表示 = 1 Then     '上下
                    '目前没有使用这个方式
                ElseIf rsTemp!要素表示 = 2 Then     '单选
                    rtxtWord(iFieldCount).SelStart = Len(rtxtWord(iFieldCount).Text)
                    rtxtWord(iFieldCount).SelLength = 0
                    rtxtWord(iFieldCount).SelText = "{{" & Nvl(rsTemp!要素值域) & "}}" & Nvl(rsTemp!要素单位)
                    
                    rtxtWord(iFieldCount).SelStart = Len(rtxtWord(iFieldCount).Text) - Len("{{" & Nvl(rsTemp!要素值域) & "}}" & Nvl(rsTemp!要素单位))
                    rtxtWord(iFieldCount).SelLength = Len("{{" & Nvl(rsTemp!要素值域) & "}}" & Nvl(rsTemp!要素单位))
                    rtxtWord(iFieldCount).SelColor = vbBlue
                ElseIf rsTemp!要素表示 = 3 Then     '复选
                    rtxtWord(iFieldCount).SelStart = Len(rtxtWord(iFieldCount).Text)
                    rtxtWord(iFieldCount).SelLength = 0
                    rtxtWord(iFieldCount).SelText = "{<" & Nvl(rsTemp!要素值域) & ">}" & Nvl(rsTemp!要素单位)
                    
                    rtxtWord(iFieldCount).SelStart = Len(rtxtWord(iFieldCount).Text) - Len("{<" & Nvl(rsTemp!要素值域) & ">}" & Nvl(rsTemp!要素单位))
                    rtxtWord(iFieldCount).SelLength = Len("{<" & Nvl(rsTemp!要素值域) & ">}" & Nvl(rsTemp!要素单位))
                    rtxtWord(iFieldCount).SelColor = vbBlue
                End If
            End If
            ResizeRichTextBox rtxtWord(iFieldCount)
            If iFieldCount = 1 Then
                miWordScale = rtxtWord(iFieldCount).Height / IIf(Len(rtxtWord(iFieldCount).Text) = 0, 1, Len(rtxtWord(iFieldCount).Text))
            End If
            rsTemp.MoveNext
        Wend
        Call ResizeWordContainer
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ClearWordShow()
    Dim i As Integer
    
    For i = 1 To rtxtWord.Count - 1
        Unload rtxtWord(i)
    Next i
    For i = 1 To cmdSelect.Count - 1
        Unload cmdSelect(i)
    Next i
End Sub

Private Sub TrvwClear()
     Dim X As Integer
     With trvWordTree
        SendMessage .hWnd, WM_SETREDRAW, 0, 0
        For X = .Nodes.Count To 1 Step -1
            .Nodes.Remove X
        Next X
        SendMessage .hWnd, WM_SETREDRAW, 1, 0
     End With
End Sub

Private Sub ResizeWordContainer()
    Dim lngH As Long
    
    '调整滚动条的位置和高度
        vscroWordH.Left = picWordShow.Width - vscroWordH.Width
        vscroWordH.Top = 0
        vscroWordH.Height = picWordShow.Height

        '调整词句容器的位置和宽度
        picWordContainer.Left = 0
        picWordContainer.Top = 0
        If picWordShow.Width > vscroWordH.Width Then
            picWordContainer.Width = picWordShow.Width - vscroWordH.Width
        Else
            picWordContainer.Width = 10
        End If

        '调整词句容器的高度
        lngH = 0
        If rtxtWord.Count > 1 Then
            lngH = rtxtWord(rtxtWord.Count - 1).Top + rtxtWord(rtxtWord.Count - 1).Height + 200
        End If

        If lngH < picWordShow.Height Then
            picWordContainer.Height = picWordShow.Height
            vscroWordH.Visible = False
        Else
            picWordContainer.Height = lngH
            vscroWordH.Visible = True
        End If

        '设置滚动条的幅度
        vscroWordH.Max = picWordContainer.Height / 1000
        vscroWordH.value = 0

End Sub



Private Sub vscroWordH_Change()
    picWordContainer.Top = -vscroWordH.value * 1000
End Sub

Public Sub zlRefresh(FileID As Long, strReportViewType As String, strReportViewTypeAlias As String, strContext As String, lngAdviceID As Long, lngDeptID As Long, _
    blnSingleWindow As Boolean, lngModul As Long, blnShowWord As Boolean, intWordDblClick As Integer, _
    intWordPower As Integer, Optional blnEditable As Boolean)
'参数：
'    intWordPower=-1，不具备词句管理权;
'    intWordPower=0，全院，这时显示所有的示范，也可以更改;
'    intWordPower=1，科室，这时显示全院通用示范(科室id is null)和所在科室公有或部门内人员私有的示范，但不能更改全院通用示范;
'    intWordPower=2，个人，这时显示全院通用示范(科室id is null)和所在科室通用示范(人员id is null)和个人示范，仅个人示范可更改

    mlngAdviceID = lngAdviceID
    mlngDeptID = lngDeptID
    mblnSingleWindow = blnSingleWindow
    mlngModul = lngModul
    mintWordDblClick = intWordDblClick
    mintWordPower = intWordPower
    mstrReportViewTypeAlias = strReportViewTypeAlias
    
    '将传入的 是否可以编辑的标记 给模块变量
'    If mblnEditable = False Then
        mblnEditable = blnEditable
'    End If
    
    If mblnSingleWindow <> blnSingleWindow Or mblnShowWord <> blnShowWord Or blnShowWord = False Then
        mblnSingleWindow = blnSingleWindow
        mblnShowWord = blnShowWord
        Call InitLoaclParas     '读取本机参数
        Call InitFaceScheme     '初始界面布局
        mblnInitFaseScheme = True
    ElseIf mblnInitFaseScheme = False Then
        Call InitLoaclParas     '读取本机参数
        Call InitFaceScheme     '初始界面布局
        mblnInitFaseScheme = True
    End If
    
    '独立窗口模式下，双击词句模板，直接写入报告，不再支持打开词句编辑窗口
    If mblnShowWord = False Then mintWordDblClick = 0
    
    Call LoadWordTree(FileID, strReportViewType, False)
    
    rtxtPrivateWord.Text = zlDatabase.GetPara("报告常用词句", glngSys, mlngModul)
    rtxtPrivateWord.Locked = True
    rtxtPrivateWord.BackColor = vbWhite
    mPrivatePane.Title = "常用词语"
    
    rtbEditWord.Text = strContext
End Sub

Private Function GetDbOwner(ByVal lngSys As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSql  As String

    GetDbOwner = ""
    err = 0: On Error GoTo errHand
    strSql = "Select 所有者 From Zlsystems Where 编号 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "GetDbOwner", lngSys)
    If rsTemp.RecordCount <> 0 Then GetDbOwner = "" & rsTemp!所有者
    rsTemp.Close
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitLoaclParas()
    Dim strRegPath As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    '读取词句示范和私人词句的宽度和高度
    If mblnSingleWindow = True Then
        strRegPath = "公共模块\" & App.ProductName & "\frmReportWord\SingleWindow"
    Else
        strRegPath = "公共模块\" & App.ProductName & "\frmReportWord"
    End If
    
    mlngWordTreeH = GetSetting("ZLSOFT", strRegPath, "WordTreeH", 200)
    mlngWordShowH = GetSetting("ZLSOFT", strRegPath, "WordShowH", 300) - 15
    mlngPrivateWordH = GetSetting("ZLSOFT", strRegPath, "PrivateWordH", 200) + 355
    mlngButtonH = GetSetting("ZLSOFT", strRegPath, "ButtonH", 500) + 325
    chk直接编辑.value = IIf(CBool(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmReportWord", "直接编辑", False)), 1, 0)
    ChkAutoExpand.value = IIf(CBool(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmReportWord", "自动展开", False)), 1, 0)
    
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub WriteWordEdit(lngWordID As Long)
    Dim strCheckView As String
    Dim strResult As String
    Dim strAdvice As String
    Dim intReportViewType As Integer
    
    Select Case mstrReportViewType
        Case ReportViewType_检查所见
            intReportViewType = 0
        Case ReportViewType_诊断意见
            intReportViewType = 1
        Case ReportViewType_建议
            intReportViewType = 2
    End Select
    
    frmReportWordEdit.zlShowMe lngWordID, Me, intReportViewType, strCheckView, strResult, strAdvice
    
    If strCheckView <> "" Then
        RaiseEvent WordSelected(strCheckView, ReportViewType_检查所见, False)
    End If
    
    If strResult <> "" Then
        RaiseEvent WordSelected(strResult, ReportViewType_诊断意见, False)
    End If
    
    If strAdvice <> "" Then
        RaiseEvent WordSelected(strAdvice, ReportViewType_建议, False)
    End If
    
    dkpMain.RecalcLayout
End Sub

Public Function ResizeRichTextBox(rtxtBox As RichTextBox) As Boolean           '判断垂直滚动条的可见性
    Dim wndStyle As Long
    Dim i As Integer
    
    i = 0
    rtxtBox.Refresh
    wndStyle = GetWindowLong(rtxtBox.hWnd, GWL_STYLE)
    
    While (wndStyle And WS_VSCROLL) <> 0 And i < 20
        rtxtBox.Height = rtxtBox.Height + 200
        rtxtBox.Refresh
        If miWordScale <> 0 Then
            '判断当前高度和文字数量之间的比例是否大于第一个文本框该比例的2倍
            If rtxtBox.Height / Len(rtxtBox.Text) > miWordScale * 2 Then
                i = 20
            End If
        End If
        wndStyle = GetWindowLong(rtxtBox.hWnd, GWL_STYLE)
        i = i + 1
    Wend
End Function

Public Sub zlShowMe(frmParent As Form, FileID As Long, strReportViewType As String, strReportViewTypeAlias As String, strContext As String, _
    lngAdviceID As Long, lngDeptID As Long, blnSingleWindow As Boolean, lngModul As Long, intWordPower As Integer, blnEditable As Boolean)
    
'    If blnEditable Then
        '将传入的 是否可以编辑的标记 给模块变量
        mblnEditable = blnEditable
        
        Call zlRefresh(FileID, strReportViewType, strReportViewTypeAlias, strContext, lngAdviceID, lngDeptID, blnSingleWindow, lngModul, False, 0, intWordPower)
        Call RestoreWinState(Me, App.ProductName)
        
        Me.Show 0, frmParent
'    End If
    
End Sub
