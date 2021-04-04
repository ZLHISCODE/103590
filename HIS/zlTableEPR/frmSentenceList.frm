VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSentenceList 
   BorderStyle     =   0  'None
   Caption         =   "词句示范列表"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   405
      ScaleHeight     =   1800
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   2520
      Width           =   2355
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   1275
         Left            =   75
         TabIndex        =   3
         Top             =   45
         Width           =   2070
         _cx             =   3651
         _cy             =   2249
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
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
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
         ExplorerBar     =   7
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
   Begin MSComctlLib.TreeView TreeList 
      Height          =   1125
      Left            =   750
      TabIndex        =   1
      Top             =   1125
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   1984
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgList"
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   240
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":0000
            Key             =   "全院"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":059A
            Key             =   "科室"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":0B34
            Key             =   "个人"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":19A8
            Key             =   "Close"
            Object.Tag             =   "折叠"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":1F42
            Key             =   "Expend"
            Object.Tag             =   "打开"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   2160
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4530
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   3810
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmSentenceList.frx":24DC
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1065
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmSentenceList.frx":2579
      Left            =   75
      Top             =   30
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSentenceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String         '当前使用者权限串
Private Enum mCol
    ID = 0: 编号: 名称: 通用级: 分类id
End Enum
Private Const conPane_Tree = 400
Private Const conPane_List = 401
Private Const conPane_Text = 404

Private mintPower As Integer       '词句使用范围
Private mbytFileType As Byte, mlngPatID As Long, mlngPageID As Long, mlngAdviceID As Long
Private mblnInit As Boolean
Private mfrmParent As Object
Public Event RowDblClick(ByVal lngWordId As Long)   '双击一行或在行上按回车

Private Function zlGetPower() As Integer
'******************************************************************************************************************
'功能：获得当前用户的词句管理的权限
'返回：词句管理权限数值
'    mintPower=-1，不具备词句管理权;
'    mintPower=0，全院，这时显示所有的示范，也可以更改;
'    mintPower=1，科室，这时显示全院通用示范(科室id is null)和所在科室公有或部门内人员私有的示范，但不能更改全院通用示范;
'    mintPower=2，个人，这时显示全院通用示范(科室id is null)和所在科室通用示范(人员id is null)和个人示范，仅个人示范可更改
'******************************************************************************************************************
Dim intPower As Integer
    If InStr(1, mstrPrivs, "全院病历词句") <> 0 Then
        intPower = 0
    ElseIf InStr(1, mstrPrivs, "科室病历词句") <> 0 Then
        intPower = 1
    ElseIf InStr(1, mstrPrivs, "个人病历词句") <> 0 Then
        intPower = 2
    Else
        intPower = -1
    End If
    zlGetPower = intPower
End Function

Public Sub zlSubRefClass(ByVal bytFileType As Byte, ByVal lngPatID As Long, ByVal lngPageID As Long, ByVal lngAdviceID As Long, ByVal frmParent As Object)
'******************************************************************************************************************
'功能：刷新分类
'参数：bytFileType 文件类型, lngPatID 病人ID,lngPageID 主页ID, lngAdviceID 医嘱ID
'8位标志11111111,分别对应8类病历:1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-疾病证明报告;6-知情文件;7-诊疗报告;8-诊疗申请
'返回：
'******************************************************************************************************************
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    mbytFileType = bytFileType: mlngPatID = lngPatID: mlngPageID = lngPageID: mlngAdviceID = lngAdviceID: Set mfrmParent = frmParent
    
    gstrSQL = "Select /*+ rule*/ Id,上级id,编码,名称 From 病历词句分类 Start With Id In ("

    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select Distinct L.分类id " & vbNewLine & _
            "From 病历词句示范 L,Table(Cast(f_Sentence_Usable([1],[2],[3],[4],[5]) As " & gstrDbOwner & ".t_Dic_Rowset)) U" & vbNewLine & _
            "Where L.ID = To_Number(U.编码)"
            
    Select Case mintPower
    Case 0
    Case 1
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 In (1, 2) And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"
    Case Else
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 = 1 And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      L.通用级 = 2 And L.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
    End Select

    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = gstrSQL & strSQL
    gstrSQL = gstrSQL & ") Connect By Prior 上级id=Id  Order By 编码"

    Dim objNode As Node

    TreeList.Nodes.Clear
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(0), mlngPatID, mlngPageID, mlngAdviceID, mbytFileType)
    If rsTemp.BOF = False Then

        Set objNode = TreeList.Nodes.Add(, , "K0", "所有词句", "Close", "Expend")
        objNode.Expanded = True
        Do While Not rsTemp.EOF

            Set objNode = Nothing

            On Error Resume Next
            Set objNode = TreeList.Nodes("K" & rsTemp("ID").Value)
            On Error GoTo errHand

            If objNode Is Nothing Then
                Set objNode = TreeList.Nodes.Add("K" & Nvl(rsTemp!上级id, 0), tvwChild, "K" & rsTemp("ID").Value, rsTemp("名称").Value, "Close", "Expend")
                objNode.Expanded = False
            End If
            rsTemp.MoveNext
        Loop
    End If
    If TreeList.Nodes.Count > 0 Then
        TreeList.Nodes(1).Selected = True
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlSubRefList(Optional ByVal lng分类id As Long)
'******************************************************************************************************************
'功能：刷新装入清单，并定位到指定的记录上
'参数：
'返回：
'******************************************************************************************************************
Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
    
    gstrSQL = "Select /*+ rule*/ L.ID,L.编号,L.名称,L.通用级,L.分类id" & vbNewLine & _
            "From 病历词句示范 L,Table(Cast(f_Sentence_Usable([1],[2],[3],[4],[5]) As " & gstrDbOwner & ".t_Dic_Rowset)) U" & vbNewLine & _
            "Where L.ID = To_Number(U.编码)"
    If lng分类id > 0 Then gstrSQL = gstrSQL & "  And L.分类id=[6] "
    '------------------------------------------------------------------------------------------------------------------
    Select Case mintPower
    Case 0
    Case 1
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 In (1, 2) And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"

    Case Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 = 1 And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      L.通用级 = 2 And L.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
    End Select
    gstrSQL = gstrSQL & " Order By 编号"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 0, mlngPatID, mlngPageID, mlngAdviceID, mbytFileType, lng分类id)
    '------------------------------------------------------------------------------------------------------------------
    Call InitList
    With vsList
        .Rows = rsTemp.RecordCount + 1
        Do Until rsTemp.EOF
            .TextMatrix(rsTemp.AbsolutePosition, mCol.ID) = rsTemp!ID
            .TextMatrix(rsTemp.AbsolutePosition, mCol.编号) = Nvl(rsTemp!编号, "N000" & rsTemp.AbsolutePosition)
            .TextMatrix(rsTemp.AbsolutePosition, mCol.名称) = rsTemp!名称
            .TextMatrix(rsTemp.AbsolutePosition, mCol.通用级) = Nvl(rsTemp!通用级, 0)
            Set .Cell(flexcpPicture, rsTemp.AbsolutePosition, mCol.名称) = imgList.ListImages(Decode(Nvl(rsTemp!通用级, 0), 0, "全院", 1, "科室", "个人")).Picture
            .TextMatrix(rsTemp.AbsolutePosition, mCol.分类id) = Nvl(rsTemp!分类id, 0)
            rsTemp.MoveNext
        Loop
        If .Rows > 1 Then Call .Select(1, 0)
    End With

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Control.Visible = False
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Control.Visible = False
    End Select
    Err.Clear
End Sub
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    Select Case Item.ID
    Case conPane_Tree
        Item.Handle = TreeList.hWnd
    Case conPane_List
        Item.Handle = picList.hWnd
    Case conPane_Text
        Item.Handle = Me.rtbText.hWnd
    End Select
    Err.Clear
End Sub

Private Sub Form_Load()
Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar, cbrControl As CommandBarControl
    mintPower = zlGetPower
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单": Me.cbsThis.ActiveMenuBar.Visible = False
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    End With
    '设置词句显示停靠窗格
    Dim panThis As Pane
    
    Set panThis = dkpMain.CreatePane(conPane_Tree, 600, 400, DockTopOf, Nothing)
    panThis.Title = "树型结构"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMain.CreatePane(conPane_List, 600, 600, DockBottomOf, panThis)
    panThis.Title = "条件列表"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMain.CreatePane(conPane_Text, 600, 400, DockBottomOf, panThis)
    panThis.Title = "示范内容"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    dkpMain.SetCommandBars Me.cbsThis
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    
    Call InitList
End Sub

Private Sub InitList()
    mblnInit = True
    With vsList
        .Clear
        .FixedRows = 1
        .Rows = 1
        .Cols = 5
        'Id = 0: 编号: 名称
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.编号) = 800: .ColWidth(mCol.名称) = 1800: .ColWidth(mCol.通用级) = 0: .ColWidth(mCol.分类id) = 0
        
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.编号) = "编号": .TextMatrix(0, mCol.名称) = "名称": .TextMatrix(0, mCol.通用级) = "通用级": .TextMatrix(0, mCol.分类id) = "分类ID"
    
        Dim i As Integer
        For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignLeftCenter
        Next
    End With
    mblnInit = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmParent = Nothing
End Sub

Private Sub TreeList_NodeClick(ByVal Node As MSComctlLib.Node)
    Call zlSubRefList(Val(Mid(Node.Key, 2)))
End Sub

Private Sub vsList_DblClick()
Dim lngWordId As Long
    If vsList.Rows = 1 Then Exit Sub
    If mblnInit Then Exit Sub
    lngWordId = vsList.TextMatrix(vsList.Row, mCol.ID)
    If lngWordId = 0 Then Exit Sub
    
    RaiseEvent RowDblClick(lngWordId)
End Sub

Private Sub vsList_RowColChange()
    Dim rsTemp As New ADODB.Recordset, lngWordId As Long
    Dim lngStart As Long, strText As String
    
    On Error GoTo errHand
    rtbText.Text = ""
    If vsList.Rows = 1 Then Exit Sub
    If mblnInit Then Exit Sub
    lngWordId = vsList.TextMatrix(vsList.Row, mCol.ID)
    If lngWordId = 0 Then Exit Sub
    
    
    gstrSQL = "Select 内容性质, 内容文本, 要素名称, 要素单位 From 病历词句组成 Where 词句id = [1] Order By 排列次序"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngWordId)
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
                End With
            Case 1, 2 '1-临时诊治要素,2-固定诊治要素
                strText = IIf(IsNull(!内容文本), "{" & !要素名称 & "}" & !要素单位, "{" & !内容文本 & "}")
                With Me.rtbText
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                End With
            End Select
            .MoveNext
        Loop
        Me.rtbText.SelStart = 0
    End With
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    vsList.Top = 0: vsList.Left = 0: vsList.Width = picList.Width: vsList.Height = picList.Height
    Err.Clear
End Sub
